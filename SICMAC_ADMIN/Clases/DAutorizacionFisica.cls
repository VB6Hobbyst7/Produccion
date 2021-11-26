VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAutorizacionFisica"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB9030C01C6
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB9030C02CB
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

Public Function GetAutorizacionFisica(psPersCod As String, Optional pnRHAutorizacionGrupo As Integer = 0) As ADODB.Recordset
    On Error GoTo ERROR
    Dim sqlA As String
    Dim sqlAux As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If pnRHAutorizacionGrupo = 0 Then
        sqlAux = ""
    Else
        sqlAux = " And nRHAutorizacionTpoGrupo = " & pnRHAutorizacionGrupo & ""
    End If

    sqlA = " Select  PE.cPersNombre + Space(50) + RHA.cPersCod Nombre, RHAT.cRHAutorizacionDesc + Space(50) + RHA.cRHAutorizacionTpo Tipo," _
         & "     Case nRHAutorizacionUnidad" _
         & "         When 1 Then Convert(varchar(10),RHA.dSolicitadoInicio,103) + ' ' + Convert(varchar(8),RHA.dSolicitadoInicio,108)" _
         & "         When 0 Then Convert(varchar(10),RHA.dSolicitadoInicio,103) end SolicitudIni," _
         & "     Case nRHAutorizacionUnidad" _
         & "         When 1 Then Convert(varchar(10),RHA.dSolicitadoFin,103) + ' ' + Convert(varchar(8),RHA.dSolicitadoFin,108)" _
         & "         When 0 Then Convert(varchar(10),RHA.dSolicitadoFin,103) end SolicitudFin," _
         & "     Case nRHAutorizacionUnidad" _
         & "         when 1 Then Convert(varchar(10),RHA.dProgramadoInicio,103) + ' ' + Convert(varchar(8),RHA.dProgramadoInicio,108)" _
         & "         when 0 Then Convert(varchar(10),RHA.dProgramadoInicio,103) end ProgramadoInicio," _
         & "     Case nRHAutorizacionUnidad" _
         & "         when 1 Then Convert(varchar(10),RHA.dProgramadoFin,103) + ' ' + Convert(varchar(8),RHA.dProgramadoFin,108)" _
         & "         when 0 Then Convert(varchar(10),RHA.dProgramadoFin,103) end ProgramadoFin," _
         & "     Case nRHAutorizacionUnidad" _
         & "         when 1 Then Convert(varchar(10),RHA.dEjecutadoInicio,103) + ' ' + Convert(varchar(8),RHA.dEjecutadoInicio,108)" _
         & "         when 0 Then Convert(varchar(10),RHA.dEjecutadoInicio,103) end EjecutadoInicio," _
         & "     Case nRHAutorizacionUnidad" _
         & "         when 1 Then Convert(varchar(10),RHA.dEjecutadoFin,103) + ' ' + Convert(varchar(8),RHA.dEjecutadoFin,108)" _
         & "         when 0 Then Convert(varchar(10),RHA.dEjecutadoFin,103) end EjecutadoFin," _
         & "     RHA.cComentario Comentario, CO.cConsDescripcion + Space(50) + RHA.cRHAutorizacionEstado" _
         & "   From RHAutorizacionFisica RHA" _
         & "   Inner Join Persona PE On PE.cPersCod = RHA.cPersCod" _
         & "   Inner Join RHAutorizacionTpo RHAT On RHA.cRHAutorizacionTpo = RHAT.cRHAutorizacionTpo And nRHAutorizacionTpoGrupo = " & pnRHAutorizacionGrupo & "" _
         & "   Inner Join Constante CO On RHA.cRHAutorizacionEstado = CO.nConsValor And CO.nConsCod = '" & gRHPermisosCab & "'" _
         & " Where RHA.cPersCod = '" & psPersCod & "'" & sqlAux
    
    If oCon.AbreConexion Then
        Set GetAutorizacionFisica = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ERROR:
    MsgBox Str(Err.Number) & Err.Description, vbInformation, "Aviso"
End Function

Public Function GetAutorizacionFisicaTipo(pnRHAutorizacionGrupo As Integer) As ADODB.Recordset
    On Error GoTo ERROR
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlA = " Select cRHAutorizacionDesc Descripcion, convert(varchar(1),nRHAutorizacionUnidad) + cRHAutorizacionTpo Codigo from rhAutorizacionTpo Where nRHAutorizacionTpoGrupo = " & pnRHAutorizacionGrupo & ""
    If oCon.AbreConexion Then
        Set GetAutorizacionFisicaTipo = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ERROR:
    MsgBox Str(Err.Number) & Err.Description, vbInformation, "Aviso"
End Function

Public Function AgredaDatos(psPersCod As String, psRHAutorizacionTpo As String, psSolicitadoInicio As String, psSolicitadoFin As String, psProgramadoInicio As String, psProgramadoFin As String, psEjecutadoInicio As String, psEjecutadoFin As String, psComentario As String, psRHAutorizacionEstado As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ERROR
    Dim sqlA As String
    Dim sqlAux As String
    Dim rsAux As ADODB.Recordset
    Set rsAux = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlA = " Insert RHAutorizacionFisica(cPersCod, cRHAutorizacionTpo, dSolicitadoInicio, dSolicitadoFin, dProgramadoInicio, dProgramadoFin, dEjecutadoInicio, dEjecutadoFin, cComentario, cRHAutorizacionEstado, cUltimaActualizacion)" _
         & " Values ('" & psPersCod & "','" & psRHAutorizacionTpo & "','" & psSolicitadoInicio & "','" & psSolicitadoFin & "','" & psProgramadoInicio & "','" & psProgramadoFin & "','" & psEjecutadoInicio & "','" & psEjecutadoFin & "','" & psComentario & "','" & psRHAutorizacionEstado & "','" & psUltimaActualizacion & "')"
    
    sqlAux = " Select cPersCod From RHAutorizacionFisica Where cPersCod = '" & psPersCod & "' And cRHAutorizacionTpo = '" & psRHAutorizacionTpo & "' And dSolicitadoInicio = '" & psSolicitadoInicio & "'"
    
    If oCon.AbreConexion Then
        Set rsAux = oCon.CargaRecordSet(sqlAux)
        If rsAux.EOF And rsAux.BOF Then
            oCon.Ejecutar sqlA
            AgredaDatos = True
        Else
            AgredaDatos = False
        End If
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ERROR:
    AgredaDatos = False
    MsgBox Str(Err.Number) & Err.Description, vbInformation, "Aviso"
End Function

Public Function ModificaDatos(psPersCod As String, psRHAutorizacionTpo As String, psSolicitadoInicio As String, psSolicitadoFin As String, psProgramadoInicio As String, psProgramadoFin As String, psEjecutadoInicio As String, psEjecutadoFin As String, psComentario As String, psRHAutorizacionEstado As String, psUltimaActualizacion As String, psRHAutorizacionTpoT As String, psSolicitadoInicioT As String) As Boolean
    On Error GoTo ERROR
    Dim sqlA As String
    Dim sqlAux As String
    Dim oCon As DConecta
    Dim rsAux As ADODB.Recordset
    Set rsAux = New ADODB.Recordset
    Set oCon = New DConecta
    
    sqlA = " Update RHAutorizacionFisica" _
         & " Set cRHAutorizacionTpo = '" & psRHAutorizacionTpo & "', dSolicitadoInicio = '" & psSolicitadoInicio & "', dSolicitadoFin = '" & psSolicitadoFin & "'," _
         & " dProgramadoInicio = '" & psProgramadoInicio & "', dProgramadoFin = '" & psProgramadoFin & "'," _
         & " dEjecutadoInicio = '" & psEjecutadoInicio & "',  dEjecutadoFin = '" & psEjecutadoFin & "'," _
         & " cComentario = '" & psComentario & "', cRHAutorizacionEstado = '" & psRHAutorizacionEstado & "', cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
         & " Where cPersCod = '" & psPersCod & "' And cRHAutorizacionTpo = '" & psRHAutorizacionTpoT & "' And dSolicitadoInicio = '" & psSolicitadoInicioT & "'"
    
    If oCon.AbreConexion Then
        If Not (psSolicitadoInicioT = psSolicitadoInicio And psRHAutorizacionTpoT = psRHAutorizacionTpo) Then
            sqlAux = " Select cPersCod From RHAutorizacionFisica Where cPersCod = '" & psPersCod & "' And cRHAutorizacionTpo = '" & psRHAutorizacionTpo & "' And dSolicitadoInicio = '" & psSolicitadoInicio & "'"
            Set rsAux = oCon.CargaRecordSet(sqlAux)
            If rsAux.EOF And rsAux.BOF Then
                oCon.Ejecutar sqlA
                ModificaDatos = True
            Else
                ModificaDatos = False
            End If
        Else
            oCon.Ejecutar sqlA
            ModificaDatos = True
        End If
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
ERROR:
    MsgBox Str(Err.Number) & Err.Description, vbInformation, "Aviso"
End Function

Public Function EliminaDatos(psPersCod As String, psRHAutorizacionTpo As String, psSolicitadoInicio As String) As Boolean
    On Error GoTo ERROR
    Dim sqlA As String
    Dim sqlAux As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlA = " Delete RHAutorizacionFisica" _
         & " Where cPersCod = '" & psPersCod & "' And cRHAutorizacionTpo = '" & psRHAutorizacionTpo & "' And dSolicitadoInicio = '" & psSolicitadoInicio & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlA
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
ERROR:
    MsgBox Str(Err.Number) & Err.Description, vbInformation, "Aviso"
End Function

Public Function GetNroDias(psPersCod As String, psFecIni As String, psFecFin As String, psTipoPer As String, Optional pnMinutos As Double = 0, Optional pnDias As Double = 0, Optional pbMinutos As Boolean = True) As Double
    On Error GoTo ERROR
    Dim sqlN As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rsN As ADODB.Recordset
    Set rsN = New ADODB.Recordset
    
'    sqlN = " Select S1.Minutos + S2.Minutos As Minutos, S1.Dias + S2.Dias As Dias From " _
         & " (Select datediff(day,dEjecutadoInicio,dEjecutadoFin) * " & gnRHMinutosDia & " Minutos, datediff(day,dEjecutadoInicio,dEjecutadoFin) Dias From RHAutorizacionfisica RHAF" _
         & " Inner Join RHAutorizacionTpo RHAFT On  RHAF.cRHAutorizacionTpo = RHAFT.cRHAutorizacionTpo And nRHAutorizacionUnidad = 0" _
         & " Where cPersCod = '" & psPersCod & "' And cRHAutorizacionEstado = '1' And RHAF.cRHAutorizacionTpo in ('" & psTipoPer & "')) As S1," _
         & " (Select datediff(minute,dEjecutadoInicio,dEjecutadoFin)  Minutos, datediff(day,dEjecutadoInicio,dEjecutadoFin) / " & gnRHMinutosDia & " Dias From RHAutorizacionfisica RHAF" _
         & " Inner Join RHAutorizacionTpo RHAFT On  RHAF.cRHAutorizacionTpo = RHAFT.cRHAutorizacionTpo And nRHAutorizacionUnidad = 1" _
         & " Where cPersCod = '" & psPersCod & "' And cRHAutorizacionEstado = '1' And RHAF.cRHAutorizacionTpo in ('" & psTipoPer & "') ) As S2"
    
    sqlN = " Select S1.Minutos As Minutos, S1.Dias  As Dias From " _
         & " (Select datediff(day,dEjecutadoInicio,dEjecutadoFin) * " & gnRHMinutosDia & " Minutos, datediff(day,dEjecutadoInicio,dEjecutadoFin) Dias From RHAutorizacionfisica RHAF" _
         & " Inner Join RHAutorizacionTpo RHAFT On  RHAF.cRHAutorizacionTpo = RHAFT.cRHAutorizacionTpo And nRHAutorizacionUnidad = 0" _
         & " Where cPersCod = '" & psPersCod & "' And cRHAutorizacionEstado = '1' And RHAF.cRHAutorizacionTpo in ('" & psTipoPer & "')) As S1    "
    
    If oCon.AbreConexion Then
        Set rsN = oCon.CargaRecordSet(sqlN)
            If rsN.EOF And rsN.BOF Then
                GetNroDias = 0
                pnDias = 0
                pnMinutos = 0
            Else
                If pbMinutos Then
                    GetNroDias = rsN!Minutos
                Else
                    GetNroDias = rsN!Dias
                End If
                pnDias = rsN!Dias
                pnMinutos = rsN!Minutos
            End If
        oCon.CierraConexion
    End If
    
    rsN.Close
    Set rsN = Nothing
    Set oCon = Nothing
    Exit Function
ERROR:
    MsgBox Str(Err.Number) & Err.Description, vbInformation, "Aviso"
End Function

'##ModelId=3AB9030C028E
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

#If DebugMode Then
    '##ModelId=3AB9030C0220
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If


