VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dInduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InsertaRHInduccion(ByVal pnCodInduccion As Long, _
                    ByVal pcCodProceso As String, _
                    ByVal pnEstado As Long)

Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHInduccion
    sSQL = "INSERT INTO RHInduccion(nCodInduccion, cCodProceso,nEstado)"
    sSQL = sSQL & " VALUES(" & pnCodInduccion & ",'" & pcCodProceso & "'," & pnEstado & ")"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHInduccion:
    Err.Raise Err.Number, "InsertaRHInduccion", Err.Description
End Sub

Public Sub ModificaRHInduccion(ByVal pnCodInduccion As Long, _
                    ByVal pcCodProceso As String, _
                    ByVal pnEstado As Long)

Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHInduccion
    sSQL = "UPDATE RHInduccion" '
    sSQL = sSQL & " SET nCodInduccion=" & pnCodInduccion & ",cCodProceso='" & pcCodProceso & "'" & ",nEstado=" & pnEstado & " WHERE nCodInduccion=" & pnCodInduccion
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHInduccion:
    Err.Raise Err.Number, "ModificaRHInduccion", Err.Description
End Sub

Public Function EliminaRHInduccion(ByVal pnCodInduccion As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHInduccion
    sSQL = "DELETE FROM RHInduccion" '
    sSQL = sSQL & " WHERE nCodInduccion=" & pnCodInduccion
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHInduccion:
    Err.Raise Err.Number, "EliminaRHInduccion", Err.Description
End Function

Public Sub InsertaRHInduccionFase(ByVal pnCodInduccion As Long, _
                    ByVal pnCodTema As Long, _
                    ByVal pcExpositor As String, _
                    ByVal pnTipoInduccion As Integer, _
                    ByVal pdInicio As Date, _
                    ByVal pdFin As Date, _
                    ByVal pdHoraIni As Date, _
                    ByVal pdHoraFin As Date, _
                    ByVal pcComentario As String, _
                    ByVal pnAceptado As Long, _
                    ByVal pcUltimaActualizacion As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHInduccionFase
    
    If pnTipoInduccion = 1 Then
        sSQL = "INSERT INTO RHInduccionFase(nCodInduccion ,nCodTema,cExpositor,dInicio,dFin ,dHoraIni,dHoraFin,cComentario ,nAceptado,cUltimaActualizacion,nTipoInduccion)"
        sSQL = sSQL & " VALUES(" & pnCodInduccion & "," & pnCodTema & ",'" & RTrim(pcExpositor) & "','" & Format(pdInicio, gsFormatoFecha) & "','" & Format(pdFin, gsFormatoFecha) & "','" & Format(pdHoraIni, gsFormatoFechaHora) & "','" & Format(pdHoraFin, gsFormatoFechaHora) & "','" & pcComentario & "'," & pnAceptado & ",'" & pcUltimaActualizacion & "',1)"
    Else    'Induccion en Campo
        sSQL = "INSERT INTO RHInduccionFase(nCodInduccion ,nCodTema,dInicio,dFin ,cComentario ,nAceptado,cUltimaActualizacion,nTipoInduccion)"
        sSQL = sSQL & " VALUES(" & pnCodInduccion & "," & pnCodTema & ",'" & Format(pdInicio, gsFormatoFecha) & "','" & Format(pdFin, gsFormatoFecha) & "','" & pcComentario & "'," & pnAceptado & ",'" & pcUltimaActualizacion & "', 2)"
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHInduccionFase:
    Err.Raise Err.Number, "InsertaRHInduccionFase", Err.Description
End Sub

Public Sub ModificaRHInduccionFase(ByVal pnCodInduccion As Long, _
                    ByVal pnCodTema As Long, _
                    ByVal pcExpositor As String, _
                    ByVal pdInicio As Date, _
                    ByVal pdFin As Date, _
                    ByVal pdHoraIni As Date, _
                    ByVal pdHoraFin As Date, _
                    ByVal pcComentario As String, _
                    ByVal pnAceptado As Long, _
                    ByVal pcUltimaActualizacion As String)

Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHInduccionFase
    sSQL = "UPDATE RHInduccionFase"
    sSQL = sSQL & " SET nCodInduccion=" & pnCodInduccion & _
                        ",nCodTema=" & pnCodTema & _
                        ",cExpositor='" & pcExpositor & "'" & _
                        ",dInicio='" & Format(pdInicio, gsFormatoFecha) & "'" & _
                        ",dFin='" & Format(pdFin, gsFormatoFecha) & "'" & _
                        ",dHoraIni='" & Format(pdHoraIni, gsFormatoFechaHora) & "'" & _
                        ",dHoraFin='" & Format(pdHoraFin, gsFormatoFechaHora) & "'" & _
                        ",cComentario='" & pcComentario & "'" & _
                        ",nAceptado=" & pnAceptado & _
                        ",cUltimaActualizacion='" & pcUltimaActualizacion & "'" & _
                        " WHERE nCodInduccion=" & pnCodInduccion & " AND nCodTema=" & pnCodTema
                                                
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    
    If pnAceptado = 2 Then
        sSQL = "UPDATE RHInduccion SET nEstado=2 WHERE nCodInduccion=" & pnCodInduccion
    End If
    If pnAceptado = 3 Then
        sSQL = "UPDATE RHInduccion SET nEstado=1 WHERE nCodInduccion=" & pnCodInduccion
    End If
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHInduccionFase:
    Err.Raise Err.Number, "ModificaRHInduccion", Err.Description
End Sub

Public Function EliminaRHInduccionFase(ByVal pnCodInduccion As Long, _
                                        ByVal pnCodTema As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHInduccionFase
    sSQL = "DELETE FROM RHInduccionFase"
    sSQL = sSQL & " WHERE nCodInduccion=" & pnCodInduccion & " AND nCodTema=" & pnCodTema
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHInduccionFase:
    Err.Raise Err.Number, "EliminaRHInduccionFase", Err.Description
End Function

Public Sub InsertaRHInduccionNota(ByVal pnCodInduccion As Long, _
                    ByVal pcPersCod As String, _
                    ByVal pcExpositor As String, _
                    ByVal pnTipoInduccion As Integer, _
                    ByVal pnNota As Double, _
                    ByVal pbAprobado As Integer, _
                    ByVal pcComentario As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHInduccionNota
    
    sSQL = "INSERT INTO RHInduccionNota(nCodInduccion ,cPersCod,cPersCodEvaluador,nTipoInduccion,nNota,bAprobado,cComentario)"
    sSQL = sSQL & " VALUES(" & pnCodInduccion & ",'" & pcPersCod & "','" & pcExpositor & "'," & pnTipoInduccion & "," & pnNota & "," & pbAprobado & ",'" & pcComentario & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHInduccionNota:
    Err.Raise Err.Number, "InsertaRHInduccionNota", Err.Description
End Sub

Public Sub ModificaRHInduccionNota(ByVal pnCodInduccion As Long, _
                    ByVal pcPersCod As String, _
                    ByVal pnTipoInduccion As Integer, _
                    ByVal pnNota As Double, _
                    ByVal pbAprobado As Integer, _
                    ByVal pcComentario As String)

Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHInduccionNota
    sSQL = "UPDATE RHInduccionNota"
    sSQL = sSQL & " SET nNota=" & pnNota & _
                        ",bAprobado=" & pbAprobado & _
                        ",cComentario='" & pcComentario & _
                        "' WHERE nCodInduccion=" & pnCodInduccion & " AND nTipoInduccion=" & pnTipoInduccion & _
                                " AND cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHInduccionNota:
    Err.Raise Err.Number, "ModificaRHInduccionNota", Err.Description
End Sub

Public Function EliminaRHInduccionNota(ByVal pnCodInduccion As Long, _
                                        ByVal pnCodTema As Long, _
                                        ByVal pcPersCod As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHInduccionNota
    sSQL = "DELETE FROM RHInduccionNota" '
    sSQL = sSQL & " WHERE nCodInduccion=" & pnCodInduccion & " AND nCodTema=" & pnCodTema & _
                        " AND cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHInduccionNota:
    Err.Raise Err.Number, "EliminaRHInduccionNota", Err.Description
End Function

Public Sub InsertaRHInduccionSesion(ByVal pnCodInduccion As Long, _
                    ByVal pnCodTema As Long, _
                    ByVal pnCodSesion As Long, _
                    ByVal pdFechaInicio As Date, _
                    ByVal pdHoraInicio As Date, _
                    ByVal pdHoraFin As Date)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHInduccionSesion
    
    sSQL = "INSERT INTO RHInduccionSesion(nCodInduccion ,nCodTema,nCodSeccion,dFecha,dHoraInicio,dHoraFin)"
    sSQL = sSQL & " VALUES(" & pnCodInduccion & "," & pnCodTema & "," & pnCodSesion & ",'" & Format(pdFechaInicio, gsFormatoFecha) & "',' " & Format(pdHoraInicio, gsFormatoFechaHora) & "','" & Format(pdHoraFin, gsFormatoFechaHora) & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHInduccionSesion:
    Err.Raise Err.Number, "InsertaRHInduccionSesion", Err.Description
End Sub

Public Sub ModificaRHInduccionSesion(ByVal pnCodInduccion As Long, _
                    ByVal pnCodTema As Long, _
                    ByVal pnCodSesion As Long, _
                    ByVal pdHoraInicio As Date, _
                    ByVal pdHoraFin As Date)

Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHInduccionSesion
    sSQL = "UPDATE RHInduccionSesion"
    sSQL = sSQL & " SET nCodInduccion=" & pnCodInduccion & _
                        ",nCodTema=" & pnCodTema & _
                        ",nCodSeccion=" & pnCodSesion & _
                        ",dHoraInicio='" & Format(pdHoraInicio, gsFormatoFechaHora) & "'" & _
                        ",dHoraFin='" & Format(pdHoraFin, gsFormatoFechaHora) & "'" & _
                        " WHERE nCodInduccion=" & pnCodInduccion & " AND nCodTema=" & pnCodTema & _
                                " AND nCodSeccion=" & pnCodSesion
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHInduccionSesion:
    Err.Raise Err.Number, "ModificaRHInduccionSesion", Err.Description
End Sub

Public Function EliminaRHInduccionSesion(ByVal pnCodInduccion As Long, _
                                        ByVal pnCodTema As Long, _
                                        ByVal pnCodSesion As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHInduccionSesion
    sSQL = "DELETE FROM RHInduccionSesion" '
    sSQL = sSQL & " WHERE nCodInduccion=" & pnCodInduccion & " AND nCodTema=" & pnCodTema & _
                        " AND nCodSeccion=" & pnCodSesion
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHInduccionSesion:
    Err.Raise Err.Number, "EliminaRHInduccionSesion", Err.Description
End Function

Public Sub InsertaRHInduccionSesionAsistencia(ByVal pnCodInduccion As Long, _
                    ByVal pnCodTema As Long, _
                    ByVal pnCodSesion As Long, _
                    ByVal pcPersCod As String, _
                    ByVal pnTipoAsistencia As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHInduccionSesionAsistencia
    
    sSQL = "INSERT INTO RHInduccionSesionAsistencia(nCodInduccion ,nCodTema,nCodSeccion,cPersCod,nTipoAsistencia)"
    sSQL = sSQL & " VALUES(" & pnCodInduccion & "," & pnCodTema & "," & pnCodSesion & ",'" & pcPersCod & "'," & pnTipoAsistencia & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHInduccionSesionAsistencia:
    Err.Raise Err.Number, "InsertaRHInduccionSesionAsistencia", Err.Description
End Sub

Public Sub ModificaRHInduccionSesionAsistencia(ByVal pnCodInduccion As Long, _
                    ByVal pnCodTema As Long, _
                    ByVal pnCodSesion As Long, _
                    ByVal pcPersCod As String, _
                    ByVal pnTipoAsistencia As Long)

Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHInduccionSesionAsistencia
    sSQL = "UPDATE RHInduccionSesionAsistencia"
    sSQL = sSQL & " SET nCodInduccion=" & pnCodInduccion & _
                        ",nCodTema=" & pnCodTema & _
                        ",nCodSeccion=" & pnCodSesion & _
                        ",cPersCod='" & pcPersCod & "'" & _
                        ",nTipoAsistencia=" & pnTipoAsistencia & _
                        " WHERE nCodInduccion=" & pnCodInduccion & " AND nCodTema=" & pnCodTema & _
                                " AND nCodSeccion=" & pnCodSesion & _
                                " AND cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHInduccionSesionAsistencia:
    Err.Raise Err.Number, "ModificaRHInduccionSesionAsistencia", Err.Description
End Sub

Public Function EliminaRHInduccionSesionAsistencia(ByVal pnCodInduccion As Long, _
                                        ByVal pnCodTema As Long, _
                                        ByVal pnCodSesion As Long, _
                                        ByVal pcPersCod As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHInduccionSesionAsistencia
    sSQL = "DELETE FROM RHInduccionSesionAsistencia" '
    sSQL = sSQL & " WHERE nCodInduccion=" & pnCodInduccion & " AND nCodTema=" & pnCodTema & _
                        " AND nCodSeccion=" & pnCodSesion & _
                        " AND cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHInduccionSesionAsistencia:
    Err.Raise Err.Number, "EliminaRHInduccionSesionAsistencia", Err.Description
End Function

Public Sub InsertaRHInduccionAsignacionFuncion(ByVal pnCodInduccion As Long, _
                    ByVal pcPersCod As String, _
                    ByVal pbMateriales As Integer, _
                    ByVal pbFunciones As Integer, _
                    ByVal pbClaves As Integer, _
                    ByVal pcAgencia As String, _
                    ByVal pcCargo As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHInduccionAsignacionFuncion
    
    sSQL = "INSERT INTO RHInduccionAsignacionFuncion(nCodInduccion ,cPersCod,bMateriales,bFunciones,bClaves,cAgencia,cCargo)"
    sSQL = sSQL & " VALUES(" & pnCodInduccion & ",'" & pcPersCod & "'," & pbMateriales & "," & pbFunciones & "," & pbClaves & ",'" & pcAgencia & "','" & pcCargo & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHInduccionAsignacionFuncion:
    Err.Raise Err.Number, "InsertaRHInduccionAsignacionFuncion", Err.Description
End Sub

Public Sub ModificaRHInduccionAsignacionFuncion(ByVal pnCodInduccion As Long, _
                    ByVal pcPersCod As String, _
                    ByVal pbMateriales As Integer, _
                    ByVal pbFunciones As Integer, _
                    ByVal pbClaves As Integer, _
                    ByVal pcAgencia As String, _
                    ByVal pcCargo As String)

Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHInduccionAsignacionFuncion
    sSQL = "UPDATE RHInduccionAsignacionFuncion"
    sSQL = sSQL & " SET nCodInduccion=" & pnCodInduccion & _
                        ",cPersCod='" & pcPersCod & "'" & _
                        ",bMateriales=" & pbMateriales & _
                        ",bFunciones=" & pbFunciones & _
                        ",bClaves=" & pbClaves & _
                        ",cAgencia='" & pcAgencia & "'" & _
                        ",cCargo='" & pcCargo & "'" & _
                        " WHERE nCodInduccion=" & pnCodInduccion & _
                                " AND cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHInduccionAsignacionFuncion:
    Err.Raise Err.Number, "ModificaRHInduccionAsignacionFuncion", Err.Description
End Sub

Public Function EliminaRHInduccionAsignacionFuncion(ByVal pnCodInduccion As Long, _
                                                ByVal pcPersCod As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHInduccionAsignacionFuncion
    sSQL = "DELETE FROM RHInduccionAsignacionFuncion" '
    sSQL = sSQL & " WHERE nCodInduccion=" & pnCodInduccion & _
                        " AND cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHInduccionAsignacionFuncion:
    Err.Raise Err.Number, "EliminaRHInduccionAsignacionFuncion", Err.Description
End Function

Public Sub InsertaRHInduccionTema(ByVal pnCodTema As Long, _
                    ByVal pcDescTema As String, _
                    ByVal pnTipoTema As Long, _
                    ByVal pcCargaHoraria As String, _
                    ByVal pnTipoCarga As Long, _
                    ByVal pnTipoInduc As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHInduccionTema
    
    sSQL = "INSERT INTO RHInduccionTema(nCodTema,cDescTema,nTipoTema,cCargaHoraria,nTipoCarga,nTipoInduc)"
    sSQL = sSQL & " VALUES(" & pnCodTema & ",'" & pcDescTema & "'," & pnTipoTema & ",'" & pcCargaHoraria & "'," & pnTipoCarga & "," & pnTipoInduc & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHInduccionTema:
    Err.Raise Err.Number, "InsertaRHInduccionTema", Err.Description
End Sub

Public Sub ModificaRHInduccionTema(ByVal pnCodTema As Long, _
                    ByVal pcDescTema As String, _
                    ByVal pnTipoTema As Long, _
                    ByVal pcCargaHoraria As String, _
                    ByVal pnTipoCarga As Long, _
                    ByVal pnTipoInduc As Long)

Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHInduccionTema
    sSQL = "UPDATE RHInduccionTema"
    sSQL = sSQL & " SET nCodTema=" & pnCodTema & _
                        ",cDescTema ='" & pcDescTema & _
                        "',nTipoTema =" & pnTipoTema & _
                        ",cCargaHoraria='" & pcCargaHoraria & _
                        "',nTipoCarga=" & pnTipoCarga & _
                        ",nTipoInduc=" & pnTipoInduc & _
                        " WHERE nCodTema=" & pnCodTema
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHInduccionTema:
    Err.Raise Err.Number, "ModificaRHInduccionTema", Err.Description
End Sub

Public Function EliminaRHInduccionTema(ByVal pnCodTema As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHInduccionTema
    sSQL = "DELETE FROM RHInduccionTema " '
    sSQL = sSQL & " WHERE  nCodTema=" & pnCodTema
                        
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHInduccionTema:
    Err.Raise Err.Number, "EliminaRHInduccionTema", Err.Description
End Function

Public Sub InsertaRHInduccionPlantillaTemas(ByVal pcRHCargoCod As String, _
                    ByVal pnCodTema As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHInduccionPlantillaTemas
    sSQL = "INSERT INTO RHInduccionPlantillaTemas(cRHCargoCod,nCodTema)"
    sSQL = sSQL & " VALUES('" & pcRHCargoCod & "'," & pnCodTema & ")"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHInduccionPlantillaTemas:
    Err.Raise Err.Number, "InsertaRHInduccionPlantillaTemas", Err.Description
End Sub

Public Function EliminaRHInduccionPlantillaTemas(ByVal pcRHCargoCod As String, _
                                                Optional ByVal pnCodTema As Long = 0)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHInduccionPlantillaTemas
    sSQL = "DELETE FROM RHInduccionPlantillaTemas" '
    sSQL = sSQL & " WHERE cRHCargoCod='" & pcRHCargoCod & IIf(pnCodTema = 0, "'", "' AND nCodTema=" & pnCodTema)
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHInduccionPlantillaTemas:
    Err.Raise Err.Number, "EliminaRHInduccionPlantillaTemas", Err.Description
End Function

Public Function Obtiene_Temas_x_Cargo(ByVal pcRHCargoCod As String) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtiene_Temas_x_Cargo
    sSQL = " SELECT IT.nCodTema,cDescTema,TipoTema=C.cConsDescripcion FROM RHInduccionPlantillaTemas IPT " & _
           " INNER JOIN RHInduccionTema IT ON IPT.nCodTema=IT.nCodTema INNER JOIN Constante C ON " & _
           " IT.nTipoTema=C.nConsValor AND C.nConsCod='9031'"
    sSQL = sSQL & " WHERE IPT.cRHCargoCod='" & pcRHCargoCod & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtiene_Temas_x_Cargo = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtiene_Temas_x_Cargo:
    Err.Raise Err.Number, "Obtiene_Temas_x_Cargo", Err.Description
End Function

Public Function Obtiene_Temas_Induccion() As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtiene_Temas_Induccion
    sSQL = "SELECT nCodTema,cDescTema,1 FROM RHInduccionTema "
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtiene_Temas_Induccion = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtiene_Temas_Induccion:
    Err.Raise Err.Number, "Obtiene_Temas_Induccion", Err.Description
End Function

Public Function Obtiene_Datos_Tema(pnCodTema As Integer) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtiene_Datos_Tema
   sSQL = "SELECT nCodTema,cDescTema,C.cConsDescripcion,cCargaHoraria,nTipoCarga,TipoCarga=C1.cConsDescripcion,nTipoInduc,TipoInduc=C2.cConsDescripcion FROM RHInduccionTema IT INNER JOIN Constante C ON IT.nTipoTema=C.nConsValor "
    sSQL = sSQL & "AND C.nConsCod='9031' INNER JOIN Constante C1 ON IT.nTipoCarga=C1.nConsValor AND C1.nConsCod='9034' INNER JOIN Constante C2 ON IT.nTipoInduc=C2.nConsValor AND C2.nConsCod='9035' WHERE nCodTema=" & pnCodTema
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtiene_Datos_Tema = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtiene_Datos_Tema:
    Err.Raise Err.Number, "Obtiene_Datos_Tema", Err.Description
End Function

Public Function Obtiene_Plantilla_x_Proceso(ByVal pcCodProceso As String, Optional ByVal pnTipoTema As Integer = 0) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtiene_Plantilla_x_Proceso
    
    Select Case pnTipoTema
        Case 0
            sSQL = " SELECT IT.nCodTema,IT.cDescTema,1 FROM RHSeleccionProceso SP INNER JOIN RHSeleccionRequerimientoproceso SRP " & _
                " ON SP.cCodProceso = SRP.cCodProceso AND SRP.cCodRequerimiento=(SELECT TOP 1 cCodRequerimiento FROM RHSeleccionRequerimientoProceso WHERE cCodProceso='" & pcCodProceso & "')" & _
                " INNER JOIN RHSeleccionRequerimiento SR ON SRP.cCodRequerimiento=SR.cCodRequerimiento INNER JOIN RHInduccionPlantillaTemas " & _
                " IPT ON IPT.cRHCargoCod=SR.cCodCargoSol INNER JOIN RHInduccionTema IT ON IPT.nCodTema=IT.nCodTema " & _
                " WHERE SP.cCodProceso='" & pcCodProceso & "'"
        Case 1
            sSQL = " SELECT IT.nCodTema,IT.cDescTema,1 FROM RHSeleccionProceso SP INNER JOIN RHSeleccionRequerimientoproceso SRP " & _
                " ON SP.cCodProceso = SRP.cCodProceso AND SRP.cCodRequerimiento=(SELECT TOP 1 cCodRequerimiento FROM RHSeleccionRequerimientoProceso WHERE cCodProceso='" & pcCodProceso & "')" & _
                " INNER JOIN RHSeleccionRequerimiento SR ON SRP.cCodRequerimiento=SR.cCodRequerimiento INNER JOIN RHInduccionPlantillaTemas " & _
                " IPT ON IPT.cRHCargoCod=SR.cCodCargoSol INNER JOIN RHInduccionTema IT ON IPT.nCodTema=IT.nCodTema " & _
                " WHERE SP.cCodProceso='" & pcCodProceso & "' AND IT.nTipoTema=1"
        Case 2
            sSQL = " SELECT IT.nCodTema,IT.cDescTema,1 FROM RHSeleccionProceso SP INNER JOIN RHSeleccionRequerimientoproceso SRP " & _
                " ON SP.cCodProceso = SRP.cCodProceso AND SRP.cCodRequerimiento=(SELECT TOP 1 cCodRequerimiento FROM RHSeleccionRequerimientoProceso WHERE cCodProceso='" & pcCodProceso & "')" & _
                " INNER JOIN RHSeleccionRequerimiento SR ON SRP.cCodRequerimiento=SR.cCodRequerimiento INNER JOIN RHInduccionPlantillaTemas " & _
                " IPT ON IPT.cRHCargoCod=SR.cCodCargoSol INNER JOIN RHInduccionTema IT ON IPT.nCodTema=IT.nCodTema " & _
                " WHERE SP.cCodProceso='" & pcCodProceso & "' AND IT.nTipoTema=2"
    End Select
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtiene_Plantilla_x_Proceso = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtiene_Plantilla_x_Proceso:
    Err.Raise Err.Number, "Obtiene_Plantilla_x_Proceso", Err.Description
End Function

Public Function Obtener_Personal_x_Agencia(ByVal pcCodProceso As String) As ADODB.Recordset
                    
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorObtener_Personal_x_Agencia
    
    sSQL = "SELECT P.cPersCod,Participante=P.cPersNombre,'Cod Agencia'=A.cAgeCod,Descripcion=A.cAgeDescripcion " & _
           " FROM RHSeleccionPostulante SELP INNER JOIN Persona P ON SELP.cPersCod=P.cPersCod INNER JOIN " & _
           " RHSeleccionRequerimiento SR ON SR.cCodRequerimiento = SELP.cCodRequerimiento INNER JOIN Agencias A " & _
           " ON A.cAgeCod = SR.cCodAgeSol WHERE SELP.cCodProceso='" & pcCodProceso & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Personal_x_Agencia = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Personal_x_Agencia:
    Err.Raise Err.Number, "Obtener_Personal_x_Agencia", Err.Description
End Function

Public Function Obtener_Cargos_Responsables(ByVal pcPersCod As String, _
                                            ByVal pnCodInduccion As Integer) As ADODB.Recordset
                    
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorObtener_Cargos_Responsables
    
    sSQL = "SELECT  CodTutor=Tutor.cPersCod,Tutor=Tutor.cPersNombre,CodResponsable=Responsable.cPersCod,Responsable=Responsable.cPersNombre " & _
           " FROM RHInduccionParticipante IP INNER JOIN Persona Tutor ON IP.cPersCodTutor=Tutor.cPersCod INNER JOIN Persona Responsable ON " & _
           " IP.cPersCodResponsable=Responsable.cPersCod " & _
           " WHERE IP.nCodInduccion=" & pnCodInduccion & " AND IP.cPersCod='" & pcPersCod & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Cargos_Responsables = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Cargos_Responsables:
    Err.Raise Err.Number, "Obtener_Cargos_Responsables", Err.Description
End Function

Public Sub InsertaRHInduccionParticipante(ByVal pnCodInduccion As Integer, _
                                    ByVal pcPersCod As String, _
                                    ByVal pcPersCodTutor As String, _
                                    ByVal pcPersCodResponsable As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHInduccionParticipante
    
    sSQL = "INSERT INTO RHInduccionParticipante(nCodInduccion,cPersCod,cPersCodTutor,cPersCodResponsable)"
    sSQL = sSQL & " VALUES(" & pnCodInduccion & ",'" & pcPersCod & "','" & pcPersCodTutor & "','" & pcPersCodResponsable & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHInduccionParticipante:
    Err.Raise Err.Number, "InsertaRHInduccionParticipante", Err.Description
End Sub

Public Sub ModificaRHInduccionParticipante(ByVal pnCodInduccion As Integer, _
                                    ByVal pcPersCod As String, _
                                    ByVal pcPersCodTutor As String, _
                                    ByVal pcPersCodResponsable As String)

Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHInduccionParticipante
    sSQL = "UPDATE RHInduccionParticipante"
    sSQL = sSQL & " SET cPersCodTutor ='" & pcPersCodTutor & _
                        "',cPersCodResponsable='" & pcPersCodResponsable & _
                        "' WHERE nCodInduccion=" & pnCodInduccion & " AND cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHInduccionParticipante:
    Err.Raise Err.Number, "ModificaRHInduccionParticipante", Err.Description
End Sub

'Public Function EliminaRHInduccionParticipante(ByVal pcPersCodTutor As String, _
'                                            ByVal pcPersCodResponsable As String)
'Dim sSQL As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorEliminaRHInduccionParticipante
'    sSQL = "DELETE FROM RHInduccionParticipante "
'    sSQL = sSQL & " WHERE nCodInduccion=" & pnCodInduccion & " AND cPersCod='" & pcPersCod & "'"
'
'
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSQL
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorEliminaRHInduccionParticipante:
'    Err.Raise Err.Number, "EliminaRHInduccionParticipante", Err.Description
'End Function

Public Sub InsertaRHInduccionFaseNota(ByVal pnCodInduccion As Integer, _
                                ByVal pnCodTema As Integer, _
                                ByVal pcPersCod As String, _
                                ByVal pnNota As Double, _
                                ByVal pbAprobado As Integer, _
                                ByVal pdFecha As Date)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHInduccionFaseNota
    
    sSQL = "INSERT INTO RHInduccionFaseNota(nCodInduccion ,nCodTema,cPersCod,nNota,bAprobado,dFecha)"
    sSQL = sSQL & " VALUES(" & pnCodInduccion & "," & pnCodTema & ",'" & pcPersCod & "'," & pnNota & "," & pbAprobado & ",'" & Format(pdFecha, gsFormatoFecha) & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHInduccionFaseNota:
    Err.Raise Err.Number, "InsertaRHInduccionFaseNota", Err.Description
End Sub

Public Sub ModificaRHInduccionFaseNota(ByVal pnCodInduccion As Long, _
                    ByVal pnCodTema As Long, _
                    ByVal pcPersCod As String, _
                    ByVal pnNota As Double, _
                    ByVal pbAprobado As Integer, _
                    ByVal pdFecha As Date)

Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHInduccionFaseNota
    sSQL = "UPDATE RHInduccionFaseNota"
    sSQL = sSQL & " SET nNota=" & pnNota & _
                        ",bAprobado=" & pbAprobado & _
                        ",dFecha='" & pdFecha & _
                        "' WHERE nCodInduccion=" & pnCodInduccion & " AND nCodTema=" & pnCodTema & _
                        " AND cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHInduccionFaseNota:
    Err.Raise Err.Number, "ModificaRHInduccionFaseNota", Err.Description
End Sub

Public Function EliminaRHInduccionFaseNota(ByVal pnCodInduccion As Integer, _
                                        ByVal pnCodTema As Integer, _
                                        ByVal pcPersCod As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHInduccionFaseNota
    sSQL = "DELETE FROM RHInduccionFaseNota" '
    sSQL = sSQL & " WHERE nCodInduccion=" & pnCodInduccion & " AND nCodTema=" & pnCodTema & _
                        " AND cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHInduccionFaseNota:
    Err.Raise Err.Number, "EliminaRHInduccionFaseNota", Err.Description
End Function

Public Function Obtener_Personal_x_Tutor(ByVal pnCodInduccion As Integer, ByVal pcPersCodTutor As String) As ADODB.Recordset
                    
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorObtener_Personal_x_Tutor
    
    sSQL = "SELECT  P.cPersCod, P.cPersNombre,A.cAgeDescripcion,dRHContratoInicio=CONVERT(varchar(10),CD.dRHContratoInicio,103) FROM RHInduccionParticipante IP INNER JOIN Persona P " & _
           " ON IP.cPersCod= P.cPersCod INNER JOIN RRHH RH ON RH.cPersCod=P.cPersCod INNER JOIN RHContrato CON ON P.cPersCod=CON.cPersCod " & _
           " AND CON.cRHContratoNro=(SELECT MAX(CONVERT(int,cRHContratoNro))FROM RHContrato WHERE cPersCod=P.cPersCod) INNER JOIN RHContratoDet " & _
           " CD ON CON.cPersCod=CD.cPersCod AND CON.cRHContratoNro=CD.cRHContratoNro AND CD.dRHContratoFecha=(SELECT MAX(dRHContratoFecha) FROM RHContratoDet " & _
           " WHERE cPersCod=CON.cPersCod AND cRHContratoNro=CON.cRHContratoNro) INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha= " & _
           " (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod " & _
           " INNER JOIN Agencias A ON A.cAgeCod = CAR.cRHAgenciaCod " & _
           " WHERE IP.nCodInduccion=" & pnCodInduccion & " AND IP.cPersCodTutor='" & pcPersCodTutor & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Personal_x_Tutor = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Personal_x_Tutor:
    Err.Raise Err.Number, "Obtener_Personal_x_Tutor", Err.Description
End Function

Public Function Obtener_Notas_x_Sesion_Campo(ByVal pnCodInduccion As Integer, ByVal pcPersCod As String, ByVal pcPersCodTutor As String) As ADODB.Recordset
                    
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorObtener_Notas_x_Sesion_Campo
    
    sSQL = " SELECT P.cPerscod,T.nCodTema,T.cDescTema,IFN.dFecha,bAprobado=CONVERT(int,IFN.bAprobado)FROM RHInduccionParticipante " & _
           " IP INNER JOIN Persona P ON IP.cPersCod=P.cPersCod INNER JOIN RHInduccionFase IFA ON IFA.nCodInduccion= IP.nCodInduccion AND " & _
           " IFA.nTipoInduccion=2 INNER JOIN RHInduccionTema T ON IFA.nCodTema=T.nCodTema INNER JOIN RHInduccionFaseNota IFN ON " & _
           " IFA.nCodInduccion=IFN.nCodInduccion AND IFA.nCodTema = IFN.nCodTema AND IFN.cPersCod=IP.cPersCod" & _
           " WHERE IP.nCodInduccion=" & pnCodInduccion & " AND IP.cPersCod='" & pcPersCod & "' AND IP.cPersCodTutor='" & pcPersCodTutor & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Notas_x_Sesion_Campo = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Notas_x_Sesion_Campo:
    Err.Raise Err.Number, "Obtener_Notas_x_Sesion_Campo", Err.Description
End Function

Public Function Obtener_Temas_a_Evaluar(ByVal pnCodInduccion As Integer, ByVal pcPersCod As String, ByVal pcPersCodTutor As String) As ADODB.Recordset
                    
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorObtener_Temas_a_Evaluar
    
    sSQL = "SELECT P.cPersCod,T.nCodTema,T.cDescTema FROM RHInduccionParticipante IP INNER JOIN Persona P ON P.cPersCod = IP.cPersCod INNER JOIN " & _
           " RHInduccionFase IFA ON IFA.nCodInduccion= IP.nCodInduccion AND IFA.nTipoInduccion=2 INNER JOIN RHInduccionTema T ON IFA.nCodTema=T.nCodTema " & _
           " WHERE IP.nCodInduccion=" & pnCodInduccion & " AND P.cPersCod='" & pcPersCod & "' AND IP.cPersCodTutor='" & pcPersCodTutor & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Temas_a_Evaluar = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Temas_a_Evaluar:
    Err.Raise Err.Number, "Obtener_Temas_a_Evaluar", Err.Description
End Function

Public Function Obtener_Notas_x_Campo_Detalle(ByVal pnCodInduccion As Integer, ByVal pcPersCod As String) As ADODB.Recordset
                    
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorObtener_Notas_x_Campo_Detalle
    
    sSQL = "SELECT P.cPersNombre,T.nCodTema,T.cDescTema,dFecha=CONVERT(varchar(10),IFN.dFecha,103),bAprobado=CONVERT(int,IFN.bAprobado)FROM RHInduccionParticipante " & _
           " IP INNER JOIN Persona P ON IP.cPersCodTutor=P.cPersCod INNER JOIN RHInduccionFase IFA ON IFA.nCodInduccion= IP.nCodInduccion AND " & _
           " IFA.nTipoInduccion=2 INNER JOIN RHInduccionTema T ON IFA.nCodTema=T.nCodTema  INNER JOIN RHInduccionFaseNota IFN ON IFA.nCodInduccion=IFN.nCodInduccion " & _
           " AND IFA.nCodTema = IFN.nCodTema AND IFN.cPersCod=IP.cPersCod" & _
           " WHERE IP.nCodInduccion=" & pnCodInduccion & " AND IP.cPersCod='" & pcPersCod & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Notas_x_Campo_Detalle = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Notas_x_Campo_Detalle:
    Err.Raise Err.Number, "Obtener_Notas_x_Campo_Detalle", Err.Description
End Function
