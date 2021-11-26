VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DPeriodoNoLaborado"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9C0D090262"
'Actualiza en Base Autorización Fisica
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB903050003
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB90305014E
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB903050112
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'Agrega a una nueva autorizacion fisica
'##ModelId=3A9C0FBA03AD
'Public Function AgregaPerNoLab(psPersCod As String, psAutorizacionTpo As String, pdSolicitadoInicio As Date, pdSolicitadoFin As Date, pdProgramadoInicio As Date, pdProgramadoFin As Date, pdEjecutadoInicio As Date, pdEjecutadoFin As Date, psComentario As String, psUltimaActualizacion As String) As Boolean
'    On Error GoTo AgregaAutorizacionFisicaErr
'    Dim oCon As DConecta
'    Dim sqlC As String
'    Set oCon = New DConecta
'
'    sqlC = " Insert RHPeriodosNoLaborados " _
'         & " (cPersCod,nRHPeriodoTpo,dSolicitadoInicio,dSolicitadoFin,dEjecutadoInicio,dEjecutadoFin,cComentario,nRHPeriodoEstado,cUltimaActualizacion)" _
'         & " Values('" & psPersCod & "'," & psRHPeriodoTpo & ",'" & psSolicitadoInicio & "','" & psSolicitadoFin & "','" & psEjecutadoInicio & "','" & psEjecutadoFin & "','" & psComentario & "'," & psRHPeriodoEstado & ",'" & psUltimaActualizacion & "')"
'
'    If oCon.AbreConexion Then
'        oCon.Ejecutar sqlC
'        oCon.CierraConexion
'    End If
'
'    Set oCon = Nothing
'    Exit Function
'AgregaAutorizacionFisicaErr:
'    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:AgregaAutorizacionFisica Method")
'End Function

'Elimina una autorizacion fisica
'##ModelId=3A9C0FBB0015
Public Function EliminaAutorizacionFisica(psCodPers As String, psAutorizacionTpo As String, pdSolicitadoInicio As Date) As Boolean
    On Error GoTo EliminaAutorizacionFisicaErr

    'your code goes here...

    Exit Function
EliminaAutorizacionFisicaErr:
    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:EliminaAutorizacionFisica Method")
End Function

'Actualiza una autorizacion fisica
'##ModelId=3A9C0FBB005B
'Public Function ModificaPeriodo(psPersCod As String, psGrupo As String, prPeriodo As ADODB.Recordset, psUltimaActualizacion As String, psFormatoFechaHora As String) As Boolean
'    On Error GoTo ModificaAutorizacionFisicaErr
'    Dim oCon As DConecta
'    Dim sqlC As String
'    Dim sqlD As String
'    Dim lsFechas As String
'    Dim lsTipos As String
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'
'    Set oCon = New DConecta
'
'    If prPeriodo Is Nothing Then
'
'    Else
'        lsFechas = ""
'        While Not prPeriodo.EOF
'            If IsDate(prPeriodo.Fields(10)) Then
'                If lsFechas = "" Then
'                    lsFechas = lsFechas & Format(CDate(prPeriodo.Fields(10)), psFormatoFechaHora)
'                Else
'                    lsFechas = lsFechas & "','" & Format(CDate(prPeriodo.Fields(10)), psFormatoFechaHora)
'                End If
'            End If
'            prPeriodo.MoveNext
'        Wend
'
'        prPeriodo.MoveFirst
'    End If
'    sqlD = " Delete RHPeriodosNoLaborados" _
'         & " Where  cPersCod = '" & psPersCod & "' And nRHPeriodoTpo like '" & psGrupo & "%' And dSolicitadoInicio Not In ('" & lsFechas & "')"
'
'    If oCon.AbreConexion Then
'        oCon.Ejecutar sqlD
'        If prPeriodo Is Nothing Then
'
'        Else
'            While Not prPeriodo.EOF
'                sqlD = " Select cPersCod From RHPeriodosNoLaborados" _
'                     & " Where cPersCod = '" & psPersCod & "' And nRHPeriodoTpo = '" & prPeriodo.Fields(0) & "' And dSolicitadoInicio = '" & Format(CDate(prPeriodo.Fields(2)), psFormatoFechaHora) & "'"
'                Set rs = oCon.CargaRecordSet(sqlD)
'                If Not (rs.EOF And rs.BOF) Then
'                    sqlC = " Update RHPeriodosNoLaborados " _
'                         & " Set dSolicitadoFin = '" & Format(CDate(prPeriodo.Fields(3)), psFormatoFechaHora) & "', dEjecutadoInicio = '" & Format(CDate(prPeriodo.Fields(4)), psFormatoFechaHora) & "', dEjecutadoFin = '" & Format(CDate(prPeriodo.Fields(5)), psFormatoFechaHora) & "', cComentario = '" & prPeriodo.Fields(6) & "', nRHPeriodoEstado = " & prPeriodo.Fields(7) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "', cObservaciones = '" & prPeriodo.Fields(9) & "', dSolicitadoInicio = '" & Format(CDate(prPeriodo.Fields(2)), psFormatoFechaHora) & "'  " _
'                         & " " _
'                         & " Where cPersCod = '" & psPersCod & "' And  nRHPeriodoTpo = " & prPeriodo.Fields(11) & " And dSolicitadoInicio = '" & Format(CDate(prPeriodo.Fields(10)), psFormatoFechaHora) & "'"
'                Else
'                    sqlC = " Insert RHPeriodosNoLaborados " _
'                         & " (cPersCod,nRHPeriodoTpo,dSolicitadoInicio,dSolicitadoFin,dEjecutadoInicio,dEjecutadoFin,cComentario,nRHPeriodoEstado,cUltimaActualizacion,cDocJustificacion,nTipoAsistenciaMedica)" _
'                         & " Values('" & psPersCod & "'," & prPeriodo.Fields(0) & ",'" & Format(CDate(prPeriodo.Fields(2)), psFormatoFechaHora) & "','" & Format(CDate(prPeriodo.Fields(3)), psFormatoFechaHora) & "','" & Format(CDate(prPeriodo.Fields(4)), psFormatoFechaHora) & "','" & Format(CDate(prPeriodo.Fields(5)), psFormatoFechaHora) & "','" & prPeriodo.Fields(6) & "'," & prPeriodo.Fields(7) & ",'" & psUltimaActualizacion & "','" & "0" & "',0)"
'                End If
'
'                oCon.Ejecutar sqlC
'                prPeriodo.MoveNext
'            Wend
'        End If
'        oCon.CierraConexion
'    End If
'
'    Set oCon = Nothing
'    Exit Function
'ModificaAutorizacionFisicaErr:
'    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:ModificaAutorizacionFisica Method")
'End Function

Public Function ModificaPeriodo(psPersCod As String, psGrupo As String, prPeriodo As ADODB.Recordset, psUltimaActualizacion As String, psFormatoFechaHora As String) As Boolean
    On Error GoTo ModificaAutorizacionFisicaErr
    Dim oCon As DConecta
    Dim sqlC As String
    Dim sqlD As String
    Dim lsFechas As String
    Dim lsTipos As String
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    
    Set oCon = New DConecta
      
    If prPeriodo Is Nothing Then
    
    Else
        lsFechas = ""
        While Not prPeriodo.EOF
            If IsDate(prPeriodo.Fields(10)) Then
                If lsFechas = "" Then
                    lsFechas = lsFechas & Format(CDate(prPeriodo.Fields(10)), psFormatoFechaHora)
                Else
                    lsFechas = lsFechas & "','" & Format(CDate(prPeriodo.Fields(10)), psFormatoFechaHora)
                End If
            End If
            prPeriodo.MoveNext
        Wend
        
        prPeriodo.MoveFirst
    End If
    sqlD = " Delete RHPeriodosNoLaborados" _
         & " Where  cPersCod = '" & psPersCod & "' And nRHPeriodoTpo like '" & psGrupo & "%' And dSolicitadoInicio Not In ('" & lsFechas & "')"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlD
        If prPeriodo Is Nothing Then
        
        Else
            While Not prPeriodo.EOF
                sqlD = " Select cPersCod From RHPeriodosNoLaborados" _
                     & " Where cPersCod = '" & psPersCod & "' And nRHPeriodoTpo = '" & prPeriodo.Fields(0) & "' And dSolicitadoInicio = '" & Format(CDate(prPeriodo.Fields(2)), psFormatoFechaHora) & "'"
                Set Rs = oCon.CargaRecordSet(sqlD)
                If Not (Rs.EOF And Rs.BOF) Then
                    sqlC = " Update RHPeriodosNoLaborados " _
                         & " Set dSolicitadoFin = '" & Format(CDate(prPeriodo.Fields(3)), psFormatoFechaHora) & "', dEjecutadoInicio = '" & Format(CDate(prPeriodo.Fields(4)), psFormatoFechaHora) & "', dEjecutadoFin = '" & Format(CDate(prPeriodo.Fields(5)), psFormatoFechaHora) & "', cComentario = '" & prPeriodo.Fields(6) & "', nRHPeriodoEstado = " & prPeriodo.Fields(7) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "', cObservaciones = '" & prPeriodo.Fields(9) & "', dSolicitadoInicio = '" & Format(CDate(prPeriodo.Fields(2)), psFormatoFechaHora) & "', nRHPeriodoTpo = " & prPeriodo.Fields(0) & "" _
                         & " Where cPersCod = '" & psPersCod & "' And  nRHPeriodoTpo = " & prPeriodo.Fields(11) & " And dSolicitadoInicio = '" & Format(CDate(prPeriodo.Fields(10)), psFormatoFechaHora) & "'"
                Else
                    sqlC = " Insert RHPeriodosNoLaborados " _
                         & " (cPersCod,nRHPeriodoTpo,dSolicitadoInicio,dSolicitadoFin,dEjecutadoInicio,dEjecutadoFin,cComentario,nRHPeriodoEstado,cUltimaActualizacion)" _
                         & " Values('" & psPersCod & "'," & prPeriodo.Fields(0) & ",'" & Format(CDate(prPeriodo.Fields(2)), psFormatoFechaHora) & "','" & Format(CDate(prPeriodo.Fields(3)), psFormatoFechaHora) & "','" & Format(CDate(prPeriodo.Fields(4)), psFormatoFechaHora) & "','" & Format(CDate(prPeriodo.Fields(5)), psFormatoFechaHora) & "','" & prPeriodo.Fields(6) & "'," & prPeriodo.Fields(7) & ",'" & psUltimaActualizacion & "')"
                End If
                
                oCon.Ejecutar sqlC
                prPeriodo.MoveNext
            Wend
        End If
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaAutorizacionFisicaErr:
    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:ModificaAutorizacionFisica Method")
End Function


'Obtiene la autorizacion de descanso fisico, en un rango de fechas
'##ModelId=3A9C102D0204
Public Function GetAutorizacionFisica(psCodPers As String, pdInicio As Date, pdFin As Date) As Integer
    On Error GoTo GetAutorizacionFisicaErr

    'your code goes here...

    Exit Function
GetAutorizacionFisicaErr:
    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:GetAutorizacionFisica Method")
End Function

'Obtiene la autorizacion de descanso fisico, en un rango de fechas
'##ModelId=3A9C102D0204
'Public Function GetRHPeriodoNoLabPersona(psCodPers As String, pnRHEstadosTpo As Integer, Optional pbSolPendientes As Boolean = False) As ADODB.Recordset
'    On Error GoTo GetAutorizacionFisicaErr
'    Dim sqlP As String
'    Dim rsP As ADODB.Recordset
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'    Set rsP = New ADODB.Recordset
'
'    If pbSolPendientes Then
'       sqlP = " Select  RHPNL.nRHPeriodoTpo CodTpo , rtrim(ltrim(str(nRHPeriodoUnidad))) +  '-' + CO.cConsDescripcion Tpo," _
'            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoInicio,103) Else Convert(varchar(10),dSolicitadoInicio,103) + ' ' + Convert(varchar(8),dSolicitadoInicio,108) end SolIni," _
'            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoFin,103) Else Convert(varchar(10),dSolicitadoFin,103) + ' ' + Convert(varchar(10),dSolicitadoFin,108) end SolFin," _
'            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dEjecutadoInicio,103) Else Convert(varchar(10),dEjecutadoInicio,103) + ' ' + Convert(varchar(10),dEjecutadoInicio,108) end EjecIni," _
'            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dEjecutadoFin,103) Else Convert(varchar(10),dEjecutadoFin,103) + ' ' + Convert(varchar(10),dEjecutadoFin,108) end EjecFin," _
'            & " cComentario Comentario,nRHPeriodoEstado Est, CO1.cConsDescripcion De,cObservaciones Obs," _
'            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoInicio,103) Else Convert(varchar(10),dSolicitadoInicio,103) + ' ' + Convert(varchar(8),dSolicitadoInicio,108) end  SolIni1," _
'            & " RHPNL.nRHPeriodoTpo CodTpo1 " _
'            & " From RHPeriodosNoLaborados RHPNL" _
'            & " Inner Join Constante CO On CO.nConsValor = RHPNL.nRHPeriodoTpo And CO.nConsCod = '" & gRHEstado & "'" _
'            & " Inner Join Constante CO1 On CO1.nConsValor = RHPNL.nRHPeriodoEstado  And CO1.nConsCod = '" & gRHPermisosCab & "'" _
'            & " Inner Join RHPeriodosNoLaboradosTpo RHPNLT On RHPNLT.nRHPeriodoTpo = RHPNL.nRHPeriodoTpo" _
'            & " Where cPersCod = '" & psCodPers & "' And RHPNL.nRHPeriodoTpo like '" & pnRHEstadosTpo & "%' And RHPNL.nRHPeriodoEstado = 0" _
'            & " Order by dSolicitadoInicio Desc"
'    Else
'       sqlP = " Select  RHPNL.nRHPeriodoTpo CodTpo , rtrim(ltrim(str(nRHPeriodoUnidad))) +  '-' + CO.cConsDescripcion Tpo," _
'            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoInicio,103) Else Convert(varchar(10),dSolicitadoInicio,103) + ' ' + Convert(varchar(8),dSolicitadoInicio,108) end SolIni," _
'            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoFin,103) Else Convert(varchar(10),dSolicitadoFin,103) + ' ' + Convert(varchar(10),dSolicitadoFin,108) end SolFin," _
'            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dEjecutadoInicio,103) Else Convert(varchar(10),dEjecutadoInicio,103) + ' ' + Convert(varchar(10),dEjecutadoInicio,108) end EjecIni," _
'            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dEjecutadoFin,103) Else Convert(varchar(10),dEjecutadoFin,103) + ' ' + Convert(varchar(10),dEjecutadoFin,108) end EjecFin," _
'            & " cComentario Comentario,nRHPeriodoEstado Est, CO1.cConsDescripcion De,cObservaciones Obs," _
'            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoInicio,103) Else Convert(varchar(10),dSolicitadoInicio,103) + ' ' + Convert(varchar(8),dSolicitadoInicio,108) end  SolIni1," _
'            & " RHPNL.nRHPeriodoTpo CodTpo1 " _
'            & " , RHPNL.cDocJustificacion, (TAM.cConsDescripcion+'" & Space(100) & "'+ Convert(varchar(4),TAM.nConsValor,4)) nTipoAsistenciaMedica " _
'            & " From RHPeriodosNoLaborados RHPNL" _
'            & " Inner Join Constante CO On CO.nConsValor = RHPNL.nRHPeriodoTpo And CO.nConsCod = '" & gRHEstado & "'" _
'            & " Inner Join Constante CO1 On CO1.nConsValor = RHPNL.nRHPeriodoEstado  And CO1.nConsCod = '" & gRHPermisosCab & "'" _
'            & " Inner Join RHPeriodosNoLaboradosTpo RHPNLT On RHPNLT.nRHPeriodoTpo = RHPNL.nRHPeriodoTpo" _
'            & " left join Constante TAM On  TAM.nConsValor = RHPNL.nTipoAsistenciaMedica And TAM.nConsCod = '9090'" _
'            & " Where cPersCod = '" & psCodPers & "' And RHPNL.nRHPeriodoTpo like '" & pnRHEstadosTpo & "%'" _
'            & " Order by dSolicitadoInicio Desc"
'    End If
'    If oCon.AbreConexion Then
'        Set GetRHPeriodoNoLabPersona = oCon.CargaRecordSet(sqlP)
'        oCon.CierraConexion
'    End If
'
'    Exit Function
'GetAutorizacionFisicaErr:
'    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:GetAutorizacionFisica Method")
'End Function

Public Function GetRHPeriodoNoLabPersona(psCodPers As String, pnRHEstadosTpo As Integer, Optional pbSolPendientes As Boolean = False) As ADODB.Recordset
    On Error GoTo GetAutorizacionFisicaErr
    Dim sqlP As String
    Dim rsP As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Set rsP = New ADODB.Recordset
     
    If pbSolPendientes Then
       sqlP = " Select  RHPNL.nRHPeriodoTpo CodTpo , rtrim(ltrim(str(nRHPeriodoUnidad))) +  '-' + CO.cConsDescripcion Tpo," _
            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoInicio,103) Else Convert(varchar(10),dSolicitadoInicio,103) + ' ' + Convert(varchar(8),dSolicitadoInicio,108) end SolIni," _
            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoFin,103) Else Convert(varchar(10),dSolicitadoFin,103) + ' ' + Convert(varchar(10),dSolicitadoFin,108) end SolFin," _
            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dEjecutadoInicio,103) Else Convert(varchar(10),dEjecutadoInicio,103) + ' ' + Convert(varchar(10),dEjecutadoInicio,108) end EjecIni," _
            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dEjecutadoFin,103) Else Convert(varchar(10),dEjecutadoFin,103) + ' ' + Convert(varchar(10),dEjecutadoFin,108) end EjecFin," _
            & " cComentario Comentario,nRHPeriodoEstado Est, CO1.cConsDescripcion De,cObservaciones Obs," _
            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoInicio,103) Else Convert(varchar(10),dSolicitadoInicio,103) + ' ' + Convert(varchar(8),dSolicitadoInicio,108) end  SolIni1," _
            & " RHPNL.nRHPeriodoTpo CodTpo1 " _
            & " From RHPeriodosNoLaborados RHPNL" _
            & " Inner Join Constante CO On CO.nConsValor = RHPNL.nRHPeriodoTpo And CO.nConsCod = '" & gRHEstado & "'" _
            & " Inner Join Constante CO1 On CO1.nConsValor = RHPNL.nRHPeriodoEstado  And CO1.nConsCod = '" & gRHPermisosCab & "'" _
            & " Inner Join RHPeriodosNoLaboradosTpo RHPNLT On RHPNLT.nRHPeriodoTpo = RHPNL.nRHPeriodoTpo" _
            & " Where cPersCod = '" & psCodPers & "' And RHPNL.nRHPeriodoTpo like '" & pnRHEstadosTpo & "%' And RHPNL.nRHPeriodoEstado = 0" _
            & " Order by dSolicitadoInicio Desc"
    Else
       sqlP = " Select  RHPNL.nRHPeriodoTpo CodTpo , rtrim(ltrim(str(nRHPeriodoUnidad))) +  '-' + CO.cConsDescripcion Tpo," _
            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoInicio,103) Else Convert(varchar(10),dSolicitadoInicio,103) + ' ' + Convert(varchar(8),dSolicitadoInicio,108) end SolIni," _
            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoFin,103) Else Convert(varchar(10),dSolicitadoFin,103) + ' ' + Convert(varchar(10),dSolicitadoFin,108) end SolFin," _
            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dEjecutadoInicio,103) Else Convert(varchar(10),dEjecutadoInicio,103) + ' ' + Convert(varchar(10),dEjecutadoInicio,108) end EjecIni," _
            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dEjecutadoFin,103) Else Convert(varchar(10),dEjecutadoFin,103) + ' ' + Convert(varchar(10),dEjecutadoFin,108) end EjecFin," _
            & " cComentario Comentario,nRHPeriodoEstado Est, CO1.cConsDescripcion De,cObservaciones Obs," _
            & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoInicio,103) Else Convert(varchar(10),dSolicitadoInicio,103) + ' ' + Convert(varchar(8),dSolicitadoInicio,108) end  SolIni1," _
            & " RHPNL.nRHPeriodoTpo CodTpo1 " _
            & " From RHPeriodosNoLaborados RHPNL" _
            & " Inner Join Constante CO On CO.nConsValor = RHPNL.nRHPeriodoTpo And CO.nConsCod = '" & gRHEstado & "'" _
            & " Inner Join Constante CO1 On CO1.nConsValor = RHPNL.nRHPeriodoEstado  And CO1.nConsCod = '" & gRHPermisosCab & "'" _
            & " Inner Join RHPeriodosNoLaboradosTpo RHPNLT On RHPNLT.nRHPeriodoTpo = RHPNL.nRHPeriodoTpo" _
            & " Where cPersCod = '" & psCodPers & "' And RHPNL.nRHPeriodoTpo like '" & pnRHEstadosTpo & "%'" _
            & " Order by dSolicitadoInicio Desc"
    End If
    If oCon.AbreConexion Then
        Set GetRHPeriodoNoLabPersona = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Exit Function
GetAutorizacionFisicaErr:
    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:GetAutorizacionFisica Method")
End Function



'Obtiene la autorizacion de descanso fisico, en un rango de fechas
'##ModelId=3A9C102D0204
Public Function GetRHPeriodoNoLabPersonaReporte(psCodPers As String, pnRHEstadosTpo As Integer) As ADODB.Recordset
    On Error GoTo GetAutorizacionFisicaErr
    Dim sqlP As String
    Dim rsP As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Set rsP = New ADODB.Recordset
     
    sqlP = " Select  RHPNL.nRHPeriodoTpo CodTpo , rtrim(ltrim(str(nRHPeriodoUnidad))) +  '-' + CO.cConsDescripcion Tpo," _
         & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoInicio,103) Else Convert(varchar(10),dSolicitadoInicio,103) + ' ' + Convert(varchar(8),dSolicitadoInicio,108) end SolIni," _
         & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoFin,103) Else Convert(varchar(10),dSolicitadoFin,103) + ' ' + Convert(varchar(10),dSolicitadoFin,108) end SolFin," _
         & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dEjecutadoInicio,103) Else Convert(varchar(10),dEjecutadoInicio,103) + ' ' + Convert(varchar(10),dEjecutadoInicio,108) end EjecIni," _
         & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dEjecutadoFin,103) Else Convert(varchar(10),dEjecutadoFin,103) + ' ' + Convert(varchar(10),dEjecutadoFin,108) end EjecFin," _
         & " cComentario Comentario,nRHPeriodoEstado Est, CO1.cConsDescripcion De,cObservaciones Obs," _
         & " Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoInicio,103) Else Convert(varchar(10),dSolicitadoInicio,103) + ' ' + Convert(varchar(8),dSolicitadoInicio,108) end  SolIni1," _
         & " RHPNL.nRHPeriodoTpo CodTpo1, CO2.cConsDescripcion Cabecera " _
         & " From RHPeriodosNoLaborados RHPNL" _
         & " Inner Join Constante CO On CO.nConsValor = RHPNL.nRHPeriodoTpo And CO.nConsCod = '" & gRHEstado & "'" _
         & " Inner Join Constante CO1 On CO1.nConsValor = RHPNL.nRHPeriodoEstado  And CO1.nConsCod = '" & gRHPermisosCab & "' Inner Join Constante CO2 On CO2.nConsValor = left(RHPNL.nRHPeriodoTpo,1) And CO2.nConsCod = '" & gRHConceptoFuncionesConstantes & "'" _
         & " Inner Join RHPeriodosNoLaboradosTpo RHPNLT On RHPNLT.nRHPeriodoTpo = RHPNL.nRHPeriodoTpo " _
         & " Where cPersCod = '" & psCodPers & "' And RHPNL.nRHPeriodoTpo like '" & pnRHEstadosTpo & "%'" _
         & " Order by dSolicitadoInicio Desc"
    
    If oCon.AbreConexion Then
        Set GetRHPeriodoNoLabPersonaReporte = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Exit Function
GetAutorizacionFisicaErr:
    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:GetAutorizacionFisica Method")
End Function

'Obtiene la autorizacion de descanso fisico, en un rango de fechas
'##ModelId=3A9C102D0204
Public Function GetRHPeriodoNoLabPersonaReporteTotal(pnRHEstadosTpo As Integer, pdIni As Date, pdFin As Date) As ADODB.Recordset
    On Error GoTo GetAutorizacionFisicaErr
    Dim sqlP As String
    Dim rsP As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Set rsP = New ADODB.Recordset
     
    sqlP = " Select  cConsDescripcion + space(75) + ltrim(rtrim(nConsValor)) Tipo, PE.cPersCod, cPersNombre, Convert(Varchar(10),dSolicitadoInicio,103) dSolicitadoInicio, DateDiff(Day,dSolicitadoInicio,dSolicitadoFin)  Dias , Convert(Varchar(10),dSolicitadoFin - 1,103) dSolicitadoFin," _
         & " Case nAplicacionTpo When 1 Then '1' Else '' End nAplicacionTpo, cComentario, Convert(Varchar(10),dVacSuspend,103) dVacSuspend, Convert(Varchar(10),dDesVacIni,103) dDesVacIni, Convert(Varchar(10),dDesVacFin,103) dDesVacFin, cObservaciones, " _
         & " ltrim(rtrim(nConsValor)) TipoAnt, PE.cPersCod cPersCodAnt, Convert(Varchar(10),dSolicitadoInicio,103) dSolicitadoInicioAnt, Convert(Varchar(10),dSolicitadoFin - 1,103) dSolicitadoFinAnt," _
         & " Case nAplicacionTpo When 1 Then '1' Else '' End nAplicacionTpoAnt, cComentario cComentarioAnt, Convert(Varchar(10),dVacSuspend,103) dVacSuspendAnt, Convert(Varchar(10),dDesVacIni,103) dDesVacIniAnt, Convert(Varchar(10),dDesVacFin,103) dDesVacFinAnt, cObservaciones cObservacionesAnt, nId, '' Eliminado ,  PNL.cUltimaActualizacion" _
         & " from RHPeriodosNoLaborados PNL" _
         & " Inner Join Constante CO On PNL.nRHPeriodoTpo = CO.nConsValor And Co.nConsCod = 6006 And CO.nConsValor Between 300 And 700" _
         & " Inner Join RHPeriodosNoLaboradosTpo PNLTPO On PNL.nRHPeriodoTpo = PNLTPO.nRHPeriodoTpo And Co.nConsCod = 6006 And CO.nConsValor Between 300 And 700" _
         & " Inner Join Persona PE On PNL.cPersCod = PE.cPersCod" _
         & " Where nRHPeriodoUnidad = 0 And dSolicitadoInicio >= '" & Format(pdIni, gsFormatoFecha) & "' And dSolicitadoInicio <= '" & Format(pdFin, gsFormatoFecha) & "' Order By PE.cPersNombre"
    
    If oCon.AbreConexion Then
        Set GetRHPeriodoNoLabPersonaReporteTotal = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Exit Function
GetAutorizacionFisicaErr:
    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:GetAutorizacionFisica Method")
End Function


Public Function ActualizaEstado(psPersCod As String, pdFecha As Date, psUltimaActualizacion As String) As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    Dim lnEstadoNuevo As Long
    Dim lnEstado  As Long
    
    oCon.AbreConexion
    oCon.BeginTrans
        
        Sql = " Select Top 1 nRHPeriodoTpo  from rhperiodosnolaborados where cPersCod = '" & psPersCod & "'" _
            & " And '" & Format(pdFecha, gsFormatoFecha) & "' Between dEjecutadoInicio And Dateadd(Day,-1,dEjecutadoFin) Order by dEjecutadoInicio Desc"
   
        Set Rs = oCon.CargaRecordSet(Sql)
        If Rs.EOF And Rs.BOF Then
            If MsgBox("No se encontro algun periodo, Si desea modificar el estado a activo elija < Si > ?", vbQuestion + vbYesNo) = vbYes Then
                If MsgBox("SI desea adicionar un estado de activo elija < Si >, caso contrario se actualiazara con estado retirado ?", vbQuestion + vbYesNo) = vbYes Then
                    lnEstadoNuevo = 201
                Else
                    'MAVM 20130814 ***
                    Sql = " Select Top 1 nRHEstado From RHEstado Where cPersCod = '" & psPersCod & "'" _
                    & " Order By cUltimaActualizacion Desc"
                    Set Rs = oCon.CargaRecordSet(Sql)
                    If Not (Rs.EOF And Rs.BOF) Then
                        lnEstadoNuevo = Rs!nRHEstado
                    Else
                        lnEstadoNuevo = 701
                    End If
                    '***
                End If
            Else
                ActualizaEstado = False
                Exit Function
            End If
        Else
            lnEstadoNuevo = Rs!nRHPeriodoTpo
        End If
        
        Rs.Close
        
        Sql = " Select nRHEstado from RRHH Where cPersCod = '" & psPersCod & "'"
        Set Rs = oCon.CargaRecordSet(Sql)
        lnEstado = Rs!nRHEstado
        Rs.Close
        oCon.GetHoraServer
        Sql = " Insert RHEstado (cPersCod,dRHEstadoFecha,nRHEstado,cRHEstadoComentario,cUltimaActualizacion)" _
            & " Values('" & psPersCod & "','" & Format(pdFecha & " " & oCon.GetHoraServer, gsFormatoFechaHora) & "'," & lnEstado & ",'Actualizaciuon de Estado','" & psUltimaActualizacion & "')"
        oCon.Ejecutar Sql
        
        Sql = " Update RRHH Set nRHEstado = " & lnEstadoNuevo & " Where cPersCod = '" & psPersCod & "'"
        oCon.Ejecutar Sql
    oCon.CommitTrans
        
    ActualizaEstado = True
End Function


Public Function ActualizaEstadoTodos(pdFecha As Date, psUltimaActualizacion As String) As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim Sql As String
    
    oCon.AbreConexion
    oCon.BeginTrans
        Sql = ""
    
    oCon.CommitTrans
        
    ActualizaEstadoTodos = True
End Function


Public Function GetPNLADesactivar(pdFecha As Date) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim Sql As String
    
    oCon.AbreConexion
    
    Sql = " Select * from rhPeriodosNoLaborados PNL" _
        & " Inner Join RRHH RH On PNL.cPersCod = RH.cPersCod And RH.nRHEstado < 700 And RH.nRHEstado <> PNL.nRHPeriodoTpo" _
        & " Inner Join Persona PE On PE.cPersCod = RH.cPersCod" _
        & " Where '" & Format(pdFecha, gsFormatoFecha) & "' Between dSolicitadoInicio And DateAdd(Day,-1,dSolicitadoFin) "
        
    Set GetPNLADesactivar = oCon.CargaRecordSet(Sql)
End Function

Public Function GetPNLAActivar(pdFecha As Date) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim Sql As String
    
    oCon.BeginTrans
    Sql = " Select RH.cRHCod, PE.cPersNombre, CO.cConsDescripcion Anterior, 'ACTIVO' Nuevo From rhPeriodosNoLaborados PNL" _
        & " Inner Join Persona PE On PNL.cPersCod = PE.cPersCod" _
        & " Inner Join RRHH RH On PNL.cPersCod = RH.cPersCod" _
        & " Inner Join Constante CO On CO.nConsCod = 6006 And CO.nConsValor = PNL.nRHPeriodoTpo" _
        & " Where '" & Format(pdFecha, gsFormatoFecha) & "' Between dSolicitadoInicio And DateAdd(Day,-1,dSolicitadoFin) " _
        & " and cPersCod in (Select cPersCod from RRHH Where nRHEstado = 201)"
       
    Set GetPNLAActivar = oCon.CargaRecordSet(Sql)
 End Function


#If DebugMode Then
    '##ModelId=3AB90305005E
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If


'Actualiza una autorizacion fisica
'##ModelId=3A9C0FBB005B
Public Function ActualizaPeriodoNLTodos(prPeriodo As ADODB.Recordset, pdIni As Date, pdFin As Date, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaAutorizacionFisicaErr
    Dim oCon As DConecta
    Dim sqlC As String
    Dim sqlD As String
    Dim lsFechas As String
    Dim lsTipos As String
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    Dim lsVacSusp As String
    Dim lsVacSuspAnt As String
    Dim lsIniDesVac As String
    Dim lsIniDesVacAnt As String
    Dim lsFinDesVac As String
    Dim lsFinDesVacAnt As String
    
    Set oCon = New DConecta
        
    If oCon.AbreConexion Then
        
        If prPeriodo Is Nothing Then
        
        Else
            prPeriodo.MoveFirst
            While Not prPeriodo.EOF
                            
                If Not IsDate(prPeriodo.Fields(8)) Then
                    lsVacSusp = "NULL"
                Else
                    lsVacSusp = "'" & Format(CDate(prPeriodo.Fields(8)), gsFormatoFecha) & "'"
                End If
                
                If Not IsDate(prPeriodo.Fields(9)) Then
                    lsIniDesVac = "NULL"
                Else
                    lsIniDesVac = "'" & Format(CDate(prPeriodo.Fields(9)), gsFormatoFecha) & "'"
                End If
                
                If Not IsDate(prPeriodo.Fields(10)) Then
                    lsFinDesVac = "NULL"
                Else
                    lsFinDesVac = "'" & Format(CDate(prPeriodo.Fields(10)), gsFormatoFecha) & "'"
                End If
                '------------------
                If Not IsDate(prPeriodo.Fields(18)) Then
                    lsVacSuspAnt = "NULL"
                Else
                    lsVacSuspAnt = "'" & Format(CDate(prPeriodo.Fields(18)), gsFormatoFecha) & "'"
                End If
                
                If Not IsDate(prPeriodo.Fields(19)) Then
                    lsIniDesVacAnt = "NULL"
                Else
                    lsIniDesVacAnt = "'" & Format(CDate(prPeriodo.Fields(19)), gsFormatoFecha) & "'"
                End If
                
                If Not IsDate(prPeriodo.Fields(20)) Then
                    lsFinDesVacAnt = "NULL"
                Else
                    lsFinDesVacAnt = "'" & Format(CDate(prPeriodo.Fields(20)), gsFormatoFecha) & "'"
                End If
                
                If prPeriodo.Fields(24) <> "" Then 'YA EXISTIA
                    If prPeriodo.Fields(23) = "X" Then 'ELIMINA
                        sqlC = " Delete RHPeriodosNoLaborados " _
                             & " Where cPersCod = '" & prPeriodo.Fields(13) & "' And nRHPeriodoTpo = " & prPeriodo.Fields(12) & " And dSolicitadoInicio = '" & Format(CDate(prPeriodo.Fields(14)), gsFormatoFechaHora) & "'"
                    Else                                'MODIFICA
                                                                                                                                         ' cod pers
                        If lsVacSusp = lsVacSuspAnt And lsIniDesVac = lsIniDesVacAnt And lsFinDesVac = lsFinDesVacAnt And prPeriodo.Fields(1) = prPeriodo.Fields(13) And prPeriodo.Fields(12) = Right(prPeriodo.Fields(12), 5) And prPeriodo.Fields(3) = prPeriodo.Fields(14) And prPeriodo.Fields(15) = prPeriodo.Fields(5) And prPeriodo.Fields(16) = prPeriodo.Fields(6) And prPeriodo.Fields(17) = prPeriodo.Fields(7) And prPeriodo.Fields(21) = prPeriodo.Fields(11) Then
                        
                        Else
                            sqlC = " Update RHPeriodosNoLaborados " _
                                 & " Set cPersCod = '" & prPeriodo.Fields(1) & "', nRHPeriodoTpo = " & Right(prPeriodo.Fields(0), 3) & ", dSolicitadoInicio = '" & Format(CDate(prPeriodo.Fields(3)), gsFormatoFechaHora) & "', dSolicitadoFin = '" & Format(DateAdd("d", 1, CDate(prPeriodo.Fields(5))), gsFormatoFechaHora) & "', dEjecutadoInicio = '" & Format(CDate(prPeriodo.Fields(3)), gsFormatoFechaHora) & "'," _
                             & "         dEjecutadoFin = '" & Format(DateAdd("d", 1, CDate(prPeriodo.Fields(5))), gsFormatoFechaHora) & "', cComentario = '" & prPeriodo.Fields(7) & "',cObservaciones = '" & prPeriodo.Fields(11) & "', cUltimaActualizacion = '" & psUltimaActualizacion & "'," _
                             & "         nTipoAsistenciaMedica = 0, cDocJustificacion = '',nAplicacionTpo = " & prPeriodo.Fields(6) & ",dVacSuspend = " & lsVacSusp & ",dDesVacIni = " & lsIniDesVac & "," _
                             & "         dDesVacFin = " & lsFinDesVac & "" _
                             & " Where cPersCod = '" & prPeriodo.Fields(13) & "' And nRHPeriodoTpo = " & prPeriodo.Fields(12) & " And dSolicitadoInicio = '" & Format(CDate(prPeriodo.Fields(14)), gsFormatoFechaHora) & "'"
                        End If
                    End If
                Else
                    sqlC = " Insert rhperiodosnolaborados ( cPersCod,nRHPeriodoTpo,dSolicitadoInicio,dSolicitadoFin,dEjecutadoInicio," _
                         & "                                dEjecutadoFin,cComentario,cObservaciones,nRHPeriodoEstado,cUltimaActualizacion," _
                         & "                                nTipoAsistenciaMedica,cDocJustificacion,nAplicacionTpo,dVacSuspend,dDesVacIni," _
                         & "                                dDesVacFin)" _
                         & " Values('" & prPeriodo.Fields(1) & "'," & Right(prPeriodo.Fields(0), 3) & ",'" & Format(CDate(prPeriodo.Fields(3)), gsFormatoFechaHora) & "','" & Format(DateAdd("d", 1, CDate(prPeriodo.Fields(5))), gsFormatoFechaHora) & "','" & Format(CDate(prPeriodo.Fields(3)), gsFormatoFechaHora) & "'," _
                         & " '" & Format(DateAdd("d", 1, CDate(prPeriodo.Fields(5))), gsFormatoFechaHora) & "','" & prPeriodo.Fields(7) & "','" & prPeriodo.Fields(11) & "',1,'" & psUltimaActualizacion & "'," _
                         & " 0,''," & prPeriodo.Fields(6) & "," & lsVacSusp & "," & lsIniDesVac & "," & lsFinDesVac & ")"
                End If
                
                If sqlC <> "" Then oCon.Ejecutar sqlC
                prPeriodo.MoveNext
            Wend
        End If
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaAutorizacionFisicaErr:
    Call RaiseError(MyUnhandledError, "DActualizacionDatosContrlAutorizacionFisica:ModificaAutorizacionFisica Method")
End Function

'MAVM 20110923 ***
Public Function CargarSubsidioConsol(ByVal sFI As String, ByVal sFF As String) As ADODB.Recordset
Dim Sql As String
Dim Co As DConecta
Dim Rs As ADODB.Recordset
Set Co = New DConecta

Sql = "Select"
Sql = Sql & " cPersNombre, A.cAgeDescripcion, cComentario Motivo, Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoInicio,103) Else Convert(varchar(10),dSolicitadoInicio,103) + ' ' + Convert(varchar(8),dSolicitadoInicio,108) end SolIni"
Sql = Sql & " , Case nRHPeriodoUnidad when  0 Then Convert(varchar(10),dSolicitadoFin,103) Else Convert(varchar(10),dSolicitadoFin,103) + ' ' + Convert(varchar(10),dSolicitadoFin,108) end SolFin , nRHPeriodoEstado Est"
Sql = Sql & " From RHPeriodosNoLaborados RHPNL Inner Join Constante CO On CO.nConsValor = RHPNL.nRHPeriodoTpo And CO.nConsCod = '6006'"
Sql = Sql & " Inner Join Constante CO1 On CO1.nConsValor = RHPNL.nRHPeriodoEstado  And CO1.nConsCod = '6025' Inner Join RHPeriodosNoLaboradosTpo RHPNLT On RHPNLT.nRHPeriodoTpo = RHPNL.nRHPeriodoTpo"
Sql = Sql & " Inner Join Persona P on RHPNL.cPersCod = P.cPersCod Inner Join RRHH R on P.cPersCod = R.cPersCod Inner Join Agencias A on R.cAgenciaActual = A.cAgeCod"
Sql = Sql & " Where RHPNL.nRHPeriodoTpo In ('501','502')"
Sql = Sql & " And Convert(Varchar (10), SubString(RHPNL.cUltimaActualizacion, 1, 8)) Between" & "'" & sFI & "' And '" & sFF & "'"
Sql = Sql & " Order by dSolicitadoInicio Desc"
    
If Co.AbreConexion Then
    Set CargarSubsidioConsol = Co.CargaRecordSet(Sql)
    Co.CierraConexion
End If
Set Rs = Nothing
Set Co = Nothing
End Function
'MAVM 20110920 ***
