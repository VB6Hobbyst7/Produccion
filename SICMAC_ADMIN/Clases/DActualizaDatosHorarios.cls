VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaDatosHorarios"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9D38F30272"
'
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB9030602D6
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB9030700D9
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB903070093
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

'Agrega a un nuevo horario tabla
'##ModelId=3A9D39150112
Public Function AgregaHorariosTabla(psRHHorarioCod As String, psRHHorarioDescripcion As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaHorariosTablaErr

    'your code goes here...

    Exit Function
AgregaHorariosTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:AgregaHorariosTabla Method")
End Function

'Elimina un horario tabla
'##ModelId=3A9D39150144
Public Function EliminaHorariosTabla(psRHHorarioCod As String) As Boolean
    On Error GoTo EliminaHorariosTablaErr

    'your code goes here...

    Exit Function
EliminaHorariosTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:EliminaHorariosTabla Method")
End Function

'Actualiza un Horario tabla
'##ModelId=3A9D39150180
Public Function ModificaHorariosTabla(psRHHorarioCod As String, psRHHorarioDescripcion As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaHorariosTablaErr

    'your code goes here...

    Exit Function
ModificaHorariosTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:ModificaHorariosTabla Method")
End Function


'Elimina una relacion horario persona
'##ModelId=3A9D391501EE
Public Function EliminaHorarios(psRHPersCod As String, pdRHHorarioFecha As Date) As Boolean
    On Error GoTo EliminaHorariosErr

    'your code goes here...

    Exit Function
EliminaHorariosErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:EliminaHorarios Method")
End Function

'Actualiza una relacion horario persona
'##ModelId=3A9D39150234
Public Function ModificaHorarios(psRHPersCod As String, pdRHHorarioFecha As Date, psUltimaActualizacion As String, psRHHorarioCod As String, psRHHorarioComentario As String) As Boolean
    On Error GoTo ModificaHorariosErr

    'your code goes here...

    Exit Function
ModificaHorariosErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:ModificaHorarios Method")
End Function

'Agrega a un nuevo horario detalle
'##ModelId=3A9D39150266
Public Function AgregaAsistencia(psFechaRef As String, prAsis As ADODB.Recordset, psUltimaActualizacion As String, pgsFormatoFecha As String) As Boolean
    On Error GoTo AgregaHorariosDetalleErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lsAux As String
    Dim lsAuxIng As String


    If oCon.AbreConexion Then

        While Not prAsis.EOF

            If prAsis.Fields(1) <> prAsis.Fields(6) Or prAsis.Fields(2) <> prAsis.Fields(7) Then
                sqlC = " Delete RHAsistenciaDet Where DateDiff(Day, dRHAsistenciaFechaRef,'" & psFechaRef & "') = 0 And cPersCod = '" & prAsis.Fields(3) & "' And cRHTurno = '" & Right(prAsis.Fields(0), 1) & "'"
                oCon.Ejecutar sqlC


                If IsDate(prAsis.Fields(1)) Then
                    lsAuxIng = "'" & Format(CDate(prAsis.Fields(1)), pgsFormatoFecha) & "'"
                Else
                    lsAuxIng = "NULL"
                End If

                If IsDate(prAsis.Fields(2)) Then
                    lsAux = "'" & Format(CDate(prAsis.Fields(2)), pgsFormatoFecha) & "'"
                Else
                    lsAux = "NULL"
                End If

                If prAsis.Fields(0) <> "" Then
                    sqlC = " Insert RHAsistenciaDet(cPersCod, dRHAsistenciaFechaRef, cRHTurno, dRHAsistenciaIngreso, dRHAsistenciaSalida, cUltimaActualizacion)" _
                         & " Values('" & prAsis.Fields(3) & "','" & psFechaRef & "','" & Right(prAsis.Fields(0), 1) & "'," & lsAuxIng & "," & lsAux & ",'" & psUltimaActualizacion & "')"
                    oCon.Ejecutar sqlC
                End If
            End If
            prAsis.MoveNext
        Wend

        oCon.CierraConexion
    End If

    Set oCon = Nothing
    Exit Function
AgregaHorariosDetalleErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:AgregaHorariosDetalle Method")
End Function

'Public Function AgregaAsistencia(psFechaRef As String, prAsis As ADODB.Recordset, psUltimaActualizacion As String, pgsFormatoFecha As String) As Boolean
'    On Error GoTo AgregaHorariosDetalleErr
'    Dim sqlC As String
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'    Dim lsAux As String
'
'
'    If oCon.AbreConexion Then
'
'        While Not prAsis.EOF
'            sqlC = " Delete RHAsistenciaDet Where DateDiff(Day, dRHAsistenciaFechaRef,'" & psFechaRef & "') = 0 And cPersCod = '" & prAsis.Fields(3) & "' And cRHTurno = '" & Right(prAsis.Fields(0), 1) & "'"
'            oCon.Ejecutar sqlC
'
'            If IsDate(prAsis.Fields(2)) Then
'                lsAux = "'" & Format(CDate(prAsis.Fields(2)), pgsFormatoFecha) & "'"
'            Else
'                lsAux = "NULL"
'            End If
'
'            If prAsis.Fields(0) <> "" And prAsis.Fields(2) <> "" And prAsis.Fields(1) <> "" Then
'                sqlC = " Insert RHAsistenciaDet(cPersCod, dRHAsistenciaFechaRef, cRHTurno, dRHAsistenciaIngreso, dRHAsistenciaSalida, cUltimaActualizacion)" _
'                     & " Values('" & prAsis.Fields(3) & "','" & psFechaRef & "','" & Right(prAsis.Fields(0), 1) & "','" & Format(CDate(prAsis.Fields(1)), pgsFormatoFecha) & "'," & lsAux & ",'" & psUltimaActualizacion & "')"
'                oCon.Ejecutar sqlC
'            End If
'            prAsis.MoveNext
'        Wend
'
'        oCon.CierraConexion
'    End If
'
'    Set oCon = Nothing
'    Exit Function
'AgregaHorariosDetalleErr:
'    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:AgregaHorariosDetalle Method")
'End Function


'Agrega a un nuevo horario detalle
'##ModelId=3A9D39150266
Public Function AgregaHorarios(psPersCod As String, prHor As ADODB.Recordset, prDia As ADODB.Recordset, psUltimaActualizacion As String, pgsFormatoFecha As String) As Boolean
    On Error GoTo AgregaHorariosDetalleErr
    Dim sqlC As String
    Dim lsCadena As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If oCon.AbreConexion Then
        oCon.BeginTrans
        lsCadena = ""
        
        If Not prHor Is Nothing Then
            prHor.MoveFirst
            While Not prHor.EOF
                If prHor.Fields(3) <> "" Then
                    If lsCadena = "" Then
                        lsCadena = Format(CDate(prHor.Fields(0)), pgsFormatoFecha)
                    Else
                        lsCadena = lsCadena & "','" & Format(CDate(prHor.Fields(0)), pgsFormatoFecha)
                    End If
                End If
                prHor.MoveNext
            Wend
        End If
        
        sqlC = " Delete RHHorarioDet Where dRHHorarioFecha Not in ('" & lsCadena & "') And cPersCod = '" & psPersCod & "'"
        oCon.Ejecutar sqlC
        sqlC = " Delete RHHorario Where dRHHorarioFecha Not in ('" & lsCadena & "') And cPersCod = '" & psPersCod & "'"
        oCon.Ejecutar sqlC
        
        If Not prHor Is Nothing Then
            prHor.MoveFirst
            While Not prHor.EOF
                If prHor.Fields(3) = "" Then
                    sqlC = " Insert RHHorario (cPersCod,dRHHorarioFecha,nRHRangoControl,cRHHorarioComentario,cUltimaActualizacion)" _
                         & " Values ('" & psPersCod & "','" & Format(CDate(prHor.Fields(0)), pgsFormatoFecha) & "'," & prHor.Fields(1) & ",'" & prHor.Fields(2) & "','" & psUltimaActualizacion & "')"
                    oCon.Ejecutar sqlC
                End If
                prHor.MoveNext
            Wend
            
            If Not prDia Is Nothing Then
                While Not prDia.EOF
                    If prDia.Fields(5) <> "" Then
                        sqlC = " Delete RHHorarioDet Where cPersCod = '" & psPersCod & "' And dRHHorarioFecha = '" & Format(CDate(prDia.Fields(5)), pgsFormatoFecha) & "' And cRHHorarioTurno = '" & Right(prDia.Fields(2), 1) & "' And cRHHorarioDias = '" & prDia.Fields(0) & "'"
                        oCon.Ejecutar sqlC
                        sqlC = " Insert RHHorarioDet (cPersCod,dRHHorarioFecha,cRHHorarioTurno,cRHHorarioDias,dRHHorarioInicio,dRHHorarioFin,cUltimaActualizacion)" _
                             & " Values('" & psPersCod & "','" & Format(CDate(prDia.Fields(5)), pgsFormatoFecha) & "','" & Right(prDia.Fields(2), 1) & "','" & prDia.Fields(0) & "','" & Format(CDate(prDia.Fields(5) & " " & prDia.Fields(3)), pgsFormatoFecha) & "','" & Format(CDate(prDia.Fields(5) & " " & prDia.Fields(4)), pgsFormatoFecha) & "','" & psUltimaActualizacion & "')"
                        oCon.Ejecutar sqlC
                    End If
                    prDia.MoveNext
                Wend
            End If
        End If
        
        oCon.CommitTrans
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
AgregaHorariosDetalleErr:
    oCon.RollBackTrans
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:AgregaHorariosDetalle Method")
End Function

'Elimina un horario detalle
'##ModelId=3A9D391502A3
Public Function EliminaHorariosDetalle(psRHHorarioCod As String, psRHHorarioTurno As String) As Boolean
    On Error GoTo EliminaHorariosDetalleErr

    'your code goes here...

    Exit Function
EliminaHorariosDetalleErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:EliminaHorariosDetalle Method")
End Function

'Actualiza un Horario detalle
'##ModelId=3A9D391502E9
Public Function ModificaHorariosDetalle(psRHHorarioCod As String, psRHHorarioTurno As String, psRHHorarioDias As String, pdRHHorarioInicio As Date, pdRHHorarioFin As Date, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaHorariosDetalleErr

    'your code goes here...

    Exit Function
ModificaHorariosDetalleErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:ModificaHorariosDetalle Method")
End Function

'Actualiza un Horario detalle
'##ModelId=3A9D391502E9
Public Function GetHorarios(psPersCod As String) As ADODB.Recordset
    On Error GoTo ModificaHorariosDetalleErr
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlC As String
    
    sqlC = " Select convert(varchar(10),dRHHorarioFecha,103) Fecha, nRHRangoControl,  cRHHorarioComentario, 1 Bandera From RHHorario where cPersCod  = '" & psPersCod & "' order by dRHHorarioFecha desc"
    
    If oCon.AbreConexion Then
        Set GetHorarios = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
      
    Set oCon = Nothing
    Exit Function
ModificaHorariosDetalleErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:ModificaHorariosDetalle Method")
End Function


'Actualiza un Horario detalle
'##ModelId=3A9D391502E9
Public Function GetHorarioTabla(pdFecha As Date, pnTipoGrupo As Integer) As ADODB.Recordset
    On Error GoTo ModificaHorariosDetalleErr
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlC As String
    
    sqlC = " Select cRHHorarioDias, CO.cConsDescripcion Dia, COTu.cConsDescripcion + space(50) + cRHHorarioTurno Turno, Convert(varchar(8),dRHHorarioInicio,108) H_Ini, Convert(varchar(8),dRHHorarioFin,108) H_Fin,'" & Format(pdFecha, gsFormatoFechaView) & "'" _
         & " From RHHorarioTabla RHD" _
         & " Inner Join Constante CO On CO.nConsValor = RHD.cRHHorarioDias And CO.nConsCod = '" & gDiasSemana & "'" _
         & " Inner Join Constante COTu On RHD.cRHHorarioTurno = COTu.nConsValor And COTu.nConsCod = '" & gRHConceptoVariablesGlobales & "'" _
         & " Where nGrupoHorario = " & pnTipoGrupo
    
    If oCon.AbreConexion Then
        Set GetHorarioTabla = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
      
    Set oCon = Nothing
    Exit Function
ModificaHorariosDetalleErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:ModificaHorariosDetalle Method")
End Function



'Actualiza un Horario detalle
'##ModelId=3A9D391502E9
Public Function GetHorariosDetalle(psPersCod As String, psRHHorarioFecha As String) As ADODB.Recordset
    On Error GoTo ModificaHorariosDetalleErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Select cRHHorarioDias, CO.cConsDescripcion Dia, COTu.cConsDescripcion + space(50) + cRHHorarioTurno Turno, Convert(varchar(8),dRHHorarioInicio,108) H_Ini, Convert(varchar(8),dRHHorarioFin,108) H_Fin,'" & Format(psRHHorarioFecha, gsFormatoFechaView) & "' from rhhorariodet RHD" _
         & " Inner Join Constante CO On CO.nConsValor = RHD.cRHHorarioDias And CO.nConsCod = '" & gDiasSemana & "'" _
         & " Inner Join Constante COTu On RHD.cRHHorarioTurno = COTu.nConsValor And COTu.nConsCod = '" & gRHConceptoVariablesGlobales & "'" _
         & " where cPersCod = '" & psPersCod & "' And  datediff(day,dRHHorarioFecha ,'" & psRHHorarioFecha & "') = 0"
    
    If oCon.AbreConexion Then
        Set GetHorariosDetalle = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
      
    Set oCon = Nothing
    Exit Function
ModificaHorariosDetalleErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:ModificaHorariosDetalle Method")
End Function

'Actualiza un Horario detalle
'##ModelId=3A9D391502E9
Public Function GetHorarioDiaDet(psFecha As String, Optional psPersCod As String = "") As ADODB.Recordset

    On Error GoTo ModificaHorariosDetalleErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    Dim sqlAge As String
    
    sqlAge = ""
    If psPersCod <> "" Then sqlAge = " And RH.cPersCod = '" & psPersCod & "'"
        
    sqlC = " Select  RH.cPersCod , ltrim(rtrim(CO.cConsDescripcion)) + space(50) + cRHTurno Turno," _
         & "         convert(varchar(10),dRHAsistenciaIngreso,103) + ' ' + convert(varchar(8),dRHAsistenciaIngreso,108) H1," _
         & "         convert(varchar(10),dRHAsistenciaSalida,103) + ' ' + convert(varchar(8),dRHAsistenciaSalida,108) H2" _
         & "         " _
         & " From RHAsistenciaDet RHAD" _
         & " Inner Join RRHH RH On RH.cPersCod = RHAD.cPersCod" _
         & "    Inner Join Constante CO On RHAD.cRHTurno = CO.nConsValor" _
         & " And CO.nConsCod = '" & gRHConceptoVariablesGlobales & "' Where Datediff(Day, dRHAsistenciaFechaRef, '" & psFecha & "' ) = 0 " & sqlAge & " Order by RH.cRHCod, cRHTurno"
    
    If oCon.AbreConexion Then
        Set GetHorarioDiaDet = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
      
    Set oCon = Nothing
    Exit Function
ModificaHorariosDetalleErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:ModificaHorariosDetalle Method")
End Function

'Actualiza un Horario detalle
'##ModelId=3A9D391502E9
'Public Function GetHorarioDia(pdFecha As Date, pgsFormatoFecha As String, Optional psAgencia As String = "", Optional psTipoContrato As String, Optional pbConMotivo As Boolean = True) As ADODB.Recordset
'    On Error GoTo ModificaHorariosDetalleErr
'    Dim oCon As DConecta
'    Dim sqlC As String
'    Dim sqlAge As String
'    Dim sqlTC As String
'    Set oCon = New DConecta
'
'    sqlAge = ""
'    If psAgencia <> "" Then sqlAge = " And RH.cAgenciaActual = '" & psAgencia & "'"
'    sqlTC = ""
'    If psTipoContrato <> "" Then sqlTC = " And Left(RH.cRHCod,1) = '" & psTipoContrato & "'"
'
'    sqlC = " Select distinct RH.cRHCod, PE.cPersNombre, T1.H1, T1.H2, T2.H1, T2.H2, RH.cPersCod, '' idx, case when " & IIf(pbConMotivo, 1, 0) & " = 1 then IsNull((Select Top 1 Motivo From (Select 'COMISION' + str(Dias) + ' Dias' Motivo, cperscod from dbo.RHAsistenciaTempTabla VVV Where VVV.cPersCod = RH.cPersCod And '" & Format(pdFecha, pgsFormatoFecha) & "' >  PartidaFecha And '" & Format(pdFecha, pgsFormatoFecha) & "' <= DateAdd(Day,Dias,PartidaFecha) Union Select cConsDescripcion Motivo, cperscod from rhperiodosnolaborados PNL Inner Join Constante CO on CO.nConsCod = 6006 And CO.nConsValor = PNL.nRHPeriodoTpo Where PNL.cPersCod = RH.cPersCod And  '" & Format(pdFecha, pgsFormatoFecha) & "' >= dsolicitadoInicio And '" & Format(pdFecha, pgsFormatoFecha) & "' < dSolicitadoFin) As AAA),'') else '' End Motivo From RRHH RH" _
'         & "        Inner Join Persona PE On RH.cPersCod = Pe.cPersCod" _
'         & "        Inner Join" _
'         & "        ( Select RHH.cPersCod, Convert(varchar(8),RHHD.dRHHorarioInicio,108) H1, Convert(varchar(8),RHHD.dRHHorarioFin,108) H2 from RHHorario RHH" _
'         & "          Inner Join RHHorarioDet RHHD" _
'         & "            On RHH.cPersCod = RHHD.cPersCod And RHH.dRHHorarioFecha = RHHD.dRHHorarioFecha" _
'         & "          where RHHD.cRHHorarioDias = '" & Weekday(pdFecha, vbMonday) & "'" _
'         & "            And RHHD.cRHHorarioTurno = '" & RHEmpleadoTurno.RHEmpleadoTurnoUno & "'" _
'         & "            And RHH.dRHHorarioFecha in" _
'         & "                    (Select Max(dRHHorarioFecha) From RHHorario RHH1" _
'         & "                where RHH1.cPersCod = RHH.cPersCod And dRHHorarioFecha <= '" & Format(pdFecha, pgsFormatoFecha) & "')" _
'         & "        ) T1 On T1.cPersCod = RH.cPersCod" _
'         & "        Left Join" _
'         & "        ( Select RHH.cPersCod, Convert(varchar(8),RHHD.dRHHorarioInicio,108) H1, Convert(varchar(8),RHHD.dRHHorarioFin,108) H2 from RHHorario RHH" _
'         & "          Inner Join RHHorarioDet RHHD" _
'         & "            On RHH.cPersCod = RHHD.cPersCod And RHH.dRHHorarioFecha = RHHD.dRHHorarioFecha" _
'         & "          where RHHD.cRHHorarioDias = '" & Weekday(pdFecha, vbMonday) & "'" _
'         & "            And RHHD.cRHHorarioTurno = '" & RHEmpleadoTurno.RHEmpleadoTurnoDos & "'" _
'         & "            And RHH.dRHHorarioFecha in" _
'         & "                    (Select Max(dRHHorarioFecha) From RHHorario RHH1" _
'         & "                where RHH1.cPersCod = RHH.cPersCod And dRHHorarioFecha <= '" & Format(pdFecha, pgsFormatoFecha) & "')" _
'         & "        ) T2 On T2.cPersCod = RH.cPersCod" _
'         & "        Left Join RHPeriodosNoLaborados RHPNL On RHPNL.cPersCod = RH.cPersCod And  '" & Format(pdFecha, pgsFormatoFecha) & "' Not Between RHPNL.dEjecutadoInicio And RHPNL.dEjecutadoFin  And RHPNL.nRHPeriodoEstado = '1'" _
'         & "        Where nRHEstado < 700 And RH.dIngreso <= " & "'" & Format(pdFecha, pgsFormatoFecha) & "' " & sqlAge & sqlTC _
'         & " Order by PE.cPersNombre"
'
'    If oCon.AbreConexion Then
'        Set GetHorarioDia = oCon.CargaRecordSet(sqlC)
'        oCon.CierraConexion
'    End If
'
'    Set oCon = Nothing
'    Exit Function
'ModificaHorariosDetalleErr:
'    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:ModificaHorariosDetalle Method")
'End Function

'Public Function GetHorarioDia(pdFecha As Date, pgsFormatoFecha As String, Optional psAgencia As String = "", Optional psTipoContrato As String, Optional pbConMotivo As Boolean = True) As ADODB.Recordset
'    On Error GoTo ModificaHorariosDetalleErr
'    Dim oCon As DConecta
'    Dim sqlC As String
'    Dim sqlAge As String
'    Dim sqlTC As String
'    Set oCon = New DConecta
'
'    sqlAge = ""
'    If psAgencia <> "" Then sqlAge = " And RH.cAgenciaActual = '" & psAgencia & "'"
'    sqlTC = ""
'    If psTipoContrato <> "" Then sqlTC = " And Left(RH.cRHCod,1) = '" & psTipoContrato & "'"
'
'    sqlC = " Select distinct RH.cRHCod, PE.cPersNombre, T1.H1, T1.H2, T2.H1, T2.H2, RH.cPersCod, '' idx, case when " & IIf(pbConMotivo, 1, 0) & " = 1 then IsNull((Select Top 1 Motivo From (Select 'COMISION' + str(Dias) + ' Dias' Motivo, cperscod from dbo.RHAsistenciaTempTabla VVV Where VVV.cPersCod = RH.cPersCod And '" & Format(pdFecha, pgsFormatoFecha) & "' >  PartidaFecha And '" & Format(pdFecha, pgsFormatoFecha) & "' <= DateAdd(Day,Dias,PartidaFecha) Union Select cConsDescripcion Motivo, cperscod from rhperiodosnolaborados PNL Inner Join Constante CO on CO.nConsCod = 6006 And CO.nConsValor = PNL.nRHPeriodoTpo Where PNL.cPersCod = RH.cPersCod And  '" & Format(pdFecha, pgsFormatoFecha) & "' >= dsolicitadoInicio And '" & Format(pdFecha, pgsFormatoFecha) & "' < dSolicitadoFin) As AAA),'') else '' End Motivo From RRHH RH" _
'         & "        Inner Join Persona PE On RH.cPersCod = Pe.cPersCod" _
'         & "        Inner Join" _
'         & "        ( Select RHH.cPersCod, Convert(varchar(8),RHHD.dRHHorarioInicio,108) H1, Convert(varchar(8),RHHD.dRHHorarioFin,108) H2 from RHHorario RHH" _
'         & "          Inner Join RHHorarioDet RHHD" _
'         & "            On RHH.cPersCod = RHHD.cPersCod And RHH.dRHHorarioFecha = RHHD.dRHHorarioFecha" _
'         & "          where RHHD.cRHHorarioDias = '" & Weekday(pdFecha, vbMonday) & "'" _
'         & "            And RHHD.cRHHorarioTurno = '" & RHEmpleadoTurno.RHEmpleadoTurnoUno & "'" _
'         & "            And RHH.dRHHorarioFecha in" _
'         & "                    (Select Max(dRHHorarioFecha) From RHHorario RHH1" _
'         & "                where RHH1.cPersCod = RHH.cPersCod And dRHHorarioFecha <= '" & Format(pdFecha, pgsFormatoFecha) & "')" _
'         & "        ) T1 On T1.cPersCod = RH.cPersCod" _
'         & "        Left Join" _
'         & "        ( Select RHH.cPersCod, Convert(varchar(8),RHHD.dRHHorarioInicio,108) H1, Convert(varchar(8),RHHD.dRHHorarioFin,108) H2 from RHHorario RHH" _
'         & "          Inner Join RHHorarioDet RHHD" _
'         & "            On RHH.cPersCod = RHHD.cPersCod And RHH.dRHHorarioFecha = RHHD.dRHHorarioFecha" _
'         & "          where RHHD.cRHHorarioDias = '" & Weekday(pdFecha, vbMonday) & "'" _
'         & "            And RHHD.cRHHorarioTurno = '" & RHEmpleadoTurno.RHEmpleadoTurnoDos & "'" _
'         & "            And RHH.dRHHorarioFecha in" _
'         & "                    (Select Max(dRHHorarioFecha) From RHHorario RHH1" _
'         & "                where RHH1.cPersCod = RHH.cPersCod And dRHHorarioFecha <= '" & Format(pdFecha, pgsFormatoFecha) & "')" _
'         & "        ) T2 On T2.cPersCod = RH.cPersCod" _
'         & "        Left Join RHPeriodosNoLaborados RHPNL On RHPNL.cPersCod = RH.cPersCod And  '" & Format(pdFecha, pgsFormatoFecha) & "' Not Between RHPNL.dEjecutadoInicio And RHPNL.dEjecutadoFin  And RHPNL.nRHPeriodoEstado = '1'" _
'         & "        Where nRHEstado < 700 And RH.dIngreso <= " & "'" & Format(pdFecha, pgsFormatoFecha) & "' " & sqlAge & sqlTC _
'         & " Order by PE.cPersNombre"
'
'    If oCon.AbreConexion Then
'        Set GetHorarioDia = oCon.CargaRecordSet(sqlC)
'        oCon.CierraConexion
'    End If
'
'    Set oCon = Nothing
'    Exit Function
'ModificaHorariosDetalleErr:
'    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:ModificaHorariosDetalle Method")
'End Function

Public Function GetHorarioDia(pdFecha As Date, pgsFormatoFecha As String, Optional psAgencia As String = "", Optional psTipoContrato As String) As ADODB.Recordset
    On Error GoTo ModificaHorariosDetalleErr
    Dim oCon As DConecta
    Dim sqlC As String
    Dim sqlAge As String
    Dim sqlTC As String
    Set oCon = New DConecta
    
    sqlAge = ""
    If psAgencia <> "" Then sqlAge = " And RH.cAgenciaActual = '" & psAgencia & "'"
    sqlTC = ""
    If psTipoContrato <> "" Then sqlTC = " And Left(RH.cRHCod,1) = '" & psTipoContrato & "'"
    
    sqlC = " Select distinct RH.cRHCod, PE.cPersNombre, T1.H1, T1.H2, T2.H1, T2.H2, RH.cPersCod  From RRHH RH" _
         & "        Inner Join Persona PE On RH.cPersCod = Pe.cPersCod" _
         & "        Inner Join" _
         & "        ( Select RHH.cPersCod, Convert(varchar(8),RHHD.dRHHorarioInicio,108) H1, Convert(varchar(8),RHHD.dRHHorarioFin,108) H2 from RHHorario RHH" _
         & "          Inner Join RHHorarioDet RHHD" _
         & "            On RHH.cPersCod = RHHD.cPersCod And RHH.dRHHorarioFecha = RHHD.dRHHorarioFecha" _
         & "          where RHHD.cRHHorarioDias = '" & Weekday(pdFecha, vbMonday) & "'" _
         & "            And RHHD.cRHHorarioTurno = '" & RHEmpleadoTurno.RHEmpleadoTurnoUno & "'" _
         & "            And RHH.dRHHorarioFecha in" _
         & "                    (Select Max(dRHHorarioFecha) From RHHorario RHH1" _
         & "                where RHH1.cPersCod = RHH.cPersCod And dRHHorarioFecha < '" & Format(pdFecha, pgsFormatoFecha) & "')" _
         & "        ) T1 On T1.cPersCod = RH.cPersCod" _
         & "        Left Join" _
         & "        ( Select RHH.cPersCod, Convert(varchar(8),RHHD.dRHHorarioInicio,108) H1, Convert(varchar(8),RHHD.dRHHorarioFin,108) H2 from RHHorario RHH" _
         & "          Inner Join RHHorarioDet RHHD" _
         & "            On RHH.cPersCod = RHHD.cPersCod And RHH.dRHHorarioFecha = RHHD.dRHHorarioFecha" _
         & "          where RHHD.cRHHorarioDias = '" & Weekday(pdFecha, vbMonday) & "'" _
         & "            And RHHD.cRHHorarioTurno = '" & RHEmpleadoTurno.RHEmpleadoTurnoDos & "'" _
         & "            And RHH.dRHHorarioFecha in" _
         & "                    (Select Max(dRHHorarioFecha) From RHHorario RHH1" _
         & "                where RHH1.cPersCod = RHH.cPersCod And dRHHorarioFecha < '" & Format(pdFecha, pgsFormatoFecha) & "')" _
         & "        ) T2 On T2.cPersCod = RH.cPersCod" _
         & "        Left Join RHPeriodosNoLaborados RHPNL On RHPNL.cPersCod = RH.cPersCod And  '" & Format(pdFecha, pgsFormatoFecha) & "' Not Between RHPNL.dEjecutadoInicio And RHPNL.dEjecutadoFin  And RHPNL.nRHPeriodoEstado = '1'" _
         & "        Where nRHEstado in ('201','301') " & sqlAge & sqlTC _
         & " Order by PE.cPersNombre"

    If oCon.AbreConexion Then
        Set GetHorarioDia = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
      
    Set oCon = Nothing
    Exit Function
ModificaHorariosDetalleErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:ModificaHorariosDetalle Method")
End Function

Public Function GetRepoAsistenciaDiaAge(pdFecha As Date, psAgeCod As String, pgsNomaAge As String, gsempresa As String, pgdFecSis As Date, psTipoContrato As String) As String
    Dim sql As String
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnPagina As Long
    Dim lnItem As Long
    
    Dim lsCod As String * 8
    Dim lsNombre As String * 30
    
    Dim lsT1I As String * 12
    Dim lsT1S As String * 12
    Dim lsT2I As String * 12
    Dim lsT2S As String * 12
    
    Dim lsA1I As String * 12
    Dim lsA1S As String * 12
    Dim lsA2I As String * 12
    Dim lsA2S As String * 12
    
    Dim lsCadena As String
    
    Dim sqlAdd As String
    Dim lnCorr As Long
    Dim lsCorr As String * 6
    
    If psAgeCod = "" Then
        sqlAdd = ""
    Else
        sqlAdd = " And cAgenciaAsig = '" & psAgeCod & "'"
    End If
    
    
    sql = " Set DateFirst  1 ;" _
        & " Select RH.cRHCod, PE.cPersNombre, HT1.T1Ini, HT1.T1Fin, HT2.T2Ini, HT2.T2Fin, AT1.A1Ini, AT1.A1Fin, AT2.A2Ini, AT2.A2Fin" _
        & " From RRHH RH" _
        & " Inner Join" _
        & "    ( Select cPersCod, dRHHorarioFecha" _
        & "         From RHHorario RH" _
        & "         Where dRHHorarioFecha = (Select Top 1 dRHHorarioFecha" _
        & "            From RHHorario RHD Where dRHHorarioFecha < '" & Format(pdFecha, gcFormatoFecha) & "'" _
        & "                 And RHD.cPersCod = RH.cPersCod" _
        & "                 Order By dRHHorarioFecha Desc)) As RHH" _
        & "     On RH.cPersCod = RHH.cPersCod" _
        & "         And RH.cPersCod Not In (Select cPersCod" _
        & "                 From RHPeriodosNoLaborados" _
        & "                     Where nRHPeriodoEstado = 1 And '" & Format(pdFecha, gcFormatoFecha) & "'" _
        & "                         Between dEjecutadoInicio And DateAdd(Day,-1,dEjecutadoFin))"
    
      sql = sql & " Inner Join" _
                & " (Select dRHHorarioInicio T1Ini, dRHHorarioFin T1Fin, cPersCod, dRHHorarioFecha" _
                & " From RHHorarioDet where cRHHorarioDias = DatePart(dw,'" & Format(pdFecha, gcFormatoFecha) & "')" _
                & " And cRHHorarioTurno = 1) As HT1 On HT1.cPersCod = RH.cPersCod And HT1.dRHHorarioFecha = RHH.dRHHorarioFecha" _
                & " Left Join" _
                & "     (Select dRHHorarioInicio T2Ini, dRHHorarioFin T2Fin, cPersCod, dRHHorarioFecha" _
                & "          From RHHorarioDet where cRHHorarioDias = DatePart(dw,'" & Format(pdFecha, gcFormatoFecha) & "')" _
                & "             And cRHHorarioTurno = 2) As HT2 On HT2.cPersCod = RH.cPersCod And HT2.dRHHorarioFecha = RHH.dRHHorarioFecha" _
                & "     Left Join" _
                & " (Select dRHAsistenciaIngreso A1Ini, dRHAsistenciaSalida A1Fin, cPersCod, dRHAsistenciaFechaRef" _
                & "     From RHAsistenciaDet where Datediff(Day,dRHAsistenciaFechaRef,'" & Format(pdFecha, gcFormatoFecha) & "') = 0" _
                & "             And cRHTurno = 1) As AT1 On AT1.cPersCod = RH.cPersCod" _
                & "     Left Join" _
                & "         (Select dRHAsistenciaIngreso A2Ini, dRHAsistenciaSalida A2Fin, cPersCod, dRHAsistenciaFechaRef" _
                & "             From RHAsistenciaDet where Datediff(Day,dRHAsistenciaFechaRef,'" & Format(pdFecha, gcFormatoFecha) & "') = 0" _
                & "                 And cRHTurno = 2) As AT2 On AT2.cPersCod = RH.cPersCod" _
                & "                     Inner Join Persona PE On RH.cPersCod = PE.cPersCod" _
                & "         Where RH.nRHEstado Like '[2]%' And RH.cRHCod Like '" & Trim(psTipoContrato) & "%' " & sqlAdd & " Order By PE.cPersNombre "
 
    oCon.AbreConexion
    
    Set rsP = oCon.CargaRecordSet(sql)
    
    lsCod = ""
    lsNombre = ""
    
    lsT1I = ""
    lsT1S = ""
    lsT2I = ""
    lsT2S = ""
    
    lsA1I = ""
    lsA1S = ""
    lsA2I = ""
    lsA2S = ""
    
    lsCadena = ""
    
    lsCadena = lsCadena & CabeceraPagina("ASISTENCIA " & Format(pgdFecSis, gsFormatoFechaView), lnPagina, lnItem, pgsNomaAge, gsempresa, pgdFecSis, "")
    lsCadena = lsCadena & Encabezado("Corr;4; ;2;Codigo;8;Nombre;10; ;20;H.INI;7; ;6;H.FIN;7; ;6;H.INI;7; ;6;H.FIN;7; ;6;A.INI;7; ;6;A.FIN;7; ;6;A.INI;7; ;6;A.FIN;7; ;6;", lnItem)
    
    lnCorr = 0
    
    While Not rsP.EOF
        lsCod = rsP!cRHCod
        lsNombre = PstaNombre(rsP!cPersNombre)
        lnCorr = lnCorr + 1
        LSet lsCorr = Format(lnCorr, "0000")
        lsT1I = Format(rsP!T1Ini, "hh:mm:ss AMPM")
        lsT1S = Format(rsP!T1Fin, "hh:mm:ss AMPM")
        lsT2I = Format(rsP!T2Ini, "hh:mm:ss AMPM")
        lsT2S = Format(rsP!T2Fin, "hh:mm:ss AMPM")
        
        lsA1I = Format(rsP!A1Ini, "hh:mm:ss AMPM")
        lsA1S = Format(rsP!A1Fin, "hh:mm:ss AMPM")
        lsA2I = Format(rsP!A2Ini, "hh:mm:ss AMPM")
        lsA2S = Format(rsP!A2Fin, "hh:mm:ss AMPM")
    
        lsCadena = lsCadena & lsCorr & lsCod & lsNombre & " " & lsT1I & " " & lsT1S & " " & lsT2I & " " & lsT2S & " " & lsA1I & " " & lsA1S & " " & lsA2I & " " & lsA2S & oImpresora.gPrnSaltoLinea
        
        lnItem = lnItem + 1
        
        If lnItem > 54 Then
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            lsCadena = lsCadena & CabeceraPagina("ASISTENCIA " & Format(pgdFecSis, gsFormatoFechaView), lnPagina, lnItem, pgsNomaAge, gsempresa, pgdFecSis, "")
            lsCadena = lsCadena & Encabezado("Corr;4; ;2;Codigo;8;Nombre;10; ;20;H.INI;7; ;6;H.FIN;7; ;6;H.INI;7; ;6;H.FIN;7; ;6;A.INI;7; ;6;A.FIN;7; ;6;A.INI;7; ;6;A.FIN;7; ;6;", lnItem)
            lnItem = 0
        End If
        
        rsP.MoveNext
    Wend
    
    GetRepoAsistenciaDiaAge = lsCadena
End Function


#If DebugMode Then
    '##ModelId=3AB903060344
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If


Public Function GetListadoFaltas(psFechaIni As String, psFechaFin As String, psCodagencia As String, psTurno As String, psSalida As String, pnMin As Integer) As ADODB.Recordset
    On Error GoTo ModificaHorariosDetalleErr
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlC As String
    
    sqlC = " SPRH_GetFaltas '" & psCodagencia & "','" & psFechaIni & "','" & psFechaFin & "','" & psTurno & "','" & psSalida & "'," & pnMin & ""
    
    If oCon.AbreConexion Then
        Set GetListadoFaltas = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
      
    Set oCon = Nothing
    Exit Function
ModificaHorariosDetalleErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosHorarios:GetListadoTar Method")
End Function

