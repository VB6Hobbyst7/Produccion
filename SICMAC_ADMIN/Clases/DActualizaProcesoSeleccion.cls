VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaProcesoSeleccion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9426E60262"
'Realiza la Actualizacion del Proceso de seleccion
Option Base 0
Option Explicit

Private lsServerComunes As String
Private lsServerPersona As String
Private lsServerAdministracion As String
Private lsServerNegocio As String
Private lsServerImagenes As String

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902EB007E
    Private mlClassDebugID As Long
#End If

Public Sub Inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub

'Agrega un nuevo Proceso de Seleccion
'##ModelId=3A9429520029
'Public Function AgregaProSelec(psRHProcSeleccNro As String, psRHProcSeleccTpo As String, psRHProcSeleccAreaCod As String, psRHProcSeleccCargoCod As String, pdRHProcSeleccInicio As String, pdRHProcSeleccFin As String, psRHProcSeleccEstado As String, psRHProcSeleccComentario As String, ptRHProcSeleccExamenPsicologico As String, ptRHProcSeleccExamenEscrito As String, psRHProcSeleccExaCur As String, psRHProcSeleccExaEsc As String, psRHProcSeleccExaPsi As String, psRHProcSeleccExaEnt As String, psRHProcSeleccNotaMax As String, psRHProcSeleccConTpo As String, psUltimaActualizacion As String) As Boolean
'    Dim oCon As DConecta
'    Dim sqlE As String
'    Set oCon = New DConecta
'    On Error GoTo AgregaProSelecErr
'
'    sqlE = " Insert " & lsServerAdministracion & "RHSeleccion (cRHSeleccionNro,cRHSeleccionTpo,cRHSeleccionAreaCod,cRHSeleccionCargoCod,dRHSeleccionInicio,dRHSeleccionFin,cRHSeleccionEstado,cRHSeleccionComentario,tRHSeleccionExamenPsicologico,tRHSeleccionExamenEscrito,cUltimaActualizacion,  bRHSeleccionExaCur,bRHSeleccionExaEsc,bRHSeleccionExaPsico,bRHSeleccionExaEnt,nRHSeleccionNotaMaxima,cRHSeleccionTipoContrato)" _
'         & " Values('" & psRHProcSeleccNro & "','" & psRHProcSeleccTpo & "','" & psRHProcSeleccAreaCod & "','" & psRHProcSeleccCargoCod & "','" & pdRHProcSeleccInicio & "','" & pdRHProcSeleccFin & "','" & psRHProcSeleccEstado & "','" & psRHProcSeleccComentario & "','','','" & psUltimaActualizacion & "'," & psRHProcSeleccExaCur & "," & psRHProcSeleccExaEsc & "," & psRHProcSeleccExaPsi & "," & psRHProcSeleccExaEnt & "," & psRHProcSeleccNotaMax & ",'" & psRHProcSeleccConTpo & "')"
'    If oCon.AbreConexion Then
'
'        oCon.Ejecutar sqlE
'        sqlE = " Update RHSeleccion " _
'             & " Set tRHSeleccionExamenPsicologico = '" & ptRHProcSeleccExamenPsicologico & "', tRHSeleccionExamenEscrito = '" & ptRHProcSeleccExamenEscrito & "'" _
'             & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"
'        oCon.Ejecutar sqlE
'
'        oCon.CierraConexion
'        AgregaProSelec = True
'    Else
'        AgregaProSelec = False
'    End If
'
'    Set oCon = Nothing
'    Exit Function
'    AgregaProSelec = False
'AgregaProSelecErr:
'    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:AgregaProSelec Method")
'End Function

'Agrega un nuevo Proceso de Evalcion
'##ModelId=3A9429520029
Public Function AgregaProEval(psRHProcEvalcNro As String, pdRHProcEvalcInicio As String, pdRHProcEvalcFin As String, psRHProcEvalcEstado As String, psRHProcEvalcComentario As String, ptRHProcEvalcExamenPsicologico As String, ptRHProcEvalcExamenEscrito As String, psUltimaActualizacion As String, prRHEvaArea As ADODB.Recordset, prRHEvaAreaEmp As ADODB.Recordset, psRHEvaluacionNotaMaxima As String, psRHEvaluacionExaCur As String, psRHEvaluacionExaEsc As String, psRHEvaluacionExaPsico As String, psRHEvaluacionExaEnt As String) As Boolean
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    On Error GoTo AgregaProEvalErr
    
    sqlE = " Insert " & lsServerAdministracion & "RHEvaluacion (cRHEvaluacionNro,dRHEvaluacionInicio,dRHEvaluacionFin,cRHEvaluacionEstado,cRHEvaluacionComentario,tRHEvaluacionExamenPsicologico,tRHEvaluacionExamenEscrito,cUltimaActualizacion,nRHEvaluacionNotaMaxima,bRHEvaluacionExaCur,bRHEvaluacionExaEsc,bRHEvaluacionExaPsico,bRHEvaluacionExaEnt)" _
         & " Values('" & psRHProcEvalcNro & "','" & pdRHProcEvalcInicio & "','" & pdRHProcEvalcFin & "','" & psRHProcEvalcEstado & "','" & psRHProcEvalcComentario & "','" & ptRHProcEvalcExamenPsicologico & "','" & ptRHProcEvalcExamenEscrito & "','" & psUltimaActualizacion & "'," & psRHEvaluacionNotaMaxima & "," & psRHEvaluacionExaCur & "," & psRHEvaluacionExaEsc & "," & psRHEvaluacionExaPsico & "," & psRHEvaluacionExaEnt & ")"
    
    If oCon.AbreConexion Then
        oCon.BeginTrans
            oCon.Ejecutar sqlE
            If Not prRHEvaArea Is Nothing Then
                If Not (prRHEvaArea.EOF And prRHEvaArea.BOF) Then
                    While Not prRHEvaArea.EOF
                        sqlE = " Insert RHEvaluaciondet(cRHEvaluacionNro, cAreaCod, cUltimaActualizacion)" _
                             & " Values('" & psRHProcEvalcNro & "','" & prRHEvaArea.Fields(0) & "','" & psUltimaActualizacion & "')"
                        oCon.Ejecutar sqlE
                        prRHEvaArea.MoveNext
                    Wend
                End If
            End If
            
            If Not prRHEvaAreaEmp Is Nothing Then
                If Not (prRHEvaAreaEmp.EOF And prRHEvaAreaEmp.BOF) Then
                    While Not prRHEvaAreaEmp.EOF
                        sqlE = " Insert RHEvaluacionRRHH (cRHEvaluacionNro,cAreaCod,cPersCod,bExonerado,cComentario,cUltimaActualizacion)" _
                             & " Values('" & psRHProcEvalcNro & "','" & prRHEvaAreaEmp.Fields(5) & "','" & prRHEvaAreaEmp.Fields(0) & "'," & prRHEvaAreaEmp.Fields(2) & ",'" & prRHEvaAreaEmp.Fields(3) & "','" & psUltimaActualizacion & "')"
                        oCon.Ejecutar sqlE
                        prRHEvaAreaEmp.MoveNext
                    Wend
                End If
            End If
        oCon.CommitTrans
        oCon.CierraConexion
        AgregaProEval = True
    Else
        AgregaProEval = False
    End If
    
    Set oCon = Nothing
    Exit Function
    AgregaProEval = False
AgregaProEvalErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvalcion:AgregaProEval Method")
End Function

Public Function GetCodigoEvaluacion() As String
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    On Error GoTo AgregaProSelecErr
    
    sqlE = " Select Top 1 cRHEvaluacionNro Cod From " & lsServerAdministracion & "RHEvaluacion Order by cRHEvaluacionNro Desc"
    If oCon.AbreConexion Then
        Set rsE = oCon.CargaRecordSet(sqlE)
        If rsE.EOF And rsE.BOF Then
            GetCodigoEvaluacion = "000001"
        Else
            GetCodigoEvaluacion = Format(CCur(rsE!Cod) + 1, "000000")
        End If
        oCon.CierraConexion
    End If
    
    rsE.Close
    Set rsE = Nothing
    Set oCon = Nothing
    Exit Function
AgregaProSelecErr:
    
End Function

Public Function ValidaNotasNulas(psRHProcSeleccNro As String) As Boolean
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    On Error GoTo AgregaProSelecErr
    
    sqlE = " Select RH.cRHSeleccionNro  From RHSeleccionDet RHD Inner Join RHSeleccion RH On RH.cRHSeleccionNro = RHD.cRHSeleccionNro Where ((nRHSeleccionNotaPsicologico Is Null And bRHSeleccionExaPsico = 1)     Or (nRHSeleccionNotaEscrito Is Null And bRHSeleccionExaEsc = 1)     Or (nRHSeleccionNotaCurricular Is Null And bRHSeleccionExaCur = 1)     Or (nRHSeleccionNotaEntrevista Is Null And bRHSeleccionExaEnt = 1)) And RH.cRHSeleccionNro = '" & psRHProcSeleccNro & "'"

    If oCon.AbreConexion Then
        Set rsE = oCon.CargaRecordSet(sqlE)
        If rsE.EOF And rsE.BOF Then
            ValidaNotasNulas = False
        Else
            ValidaNotasNulas = True
        End If
        oCon.CierraConexion
    End If
    
    rsE.Close
    Set rsE = Nothing
    Set oCon = Nothing
    Exit Function
AgregaProSelecErr:
    
End Function

Public Function ValidaNotasNulasEva(psRHProcSeleccNro As String) As Boolean
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    On Error GoTo AgregaProSelecErr
    
    sqlE = " Select RH.cRHEvaluacionNro From RHEvaluacionRRHH RHD Inner Join RHEvaluacion RH On RH.cRHEvaluacionNro = RHD.cRHEvaluacionNro" _
         & " Where ((nRHEvaluacionNotaPsicologica Is Null And bRHEvaluacionExaPsico = 1) " _
         & " Or (nRHEvaluacionNotaEscrita Is Null And bRHEvaluacionExaEsc = 1) " _
         & " Or (nRHEvaluacionNotaCurricular Is Null And bRHEvaluacionExaCur = 1) " _
         & " Or (nRHEvaluacionNotaEntrevista Is Null And bRHEvaluacionExaEnt = 1)) " _
         & " And (bExonerado is null or bExonerado = 0) And RH.cRHEvaluacionNro = '" & psRHProcSeleccNro & "'"

    If oCon.AbreConexion Then
        Set rsE = oCon.CargaRecordSet(sqlE)
        If rsE.EOF And rsE.BOF Then
            ValidaNotasNulasEva = False
        Else
            ValidaNotasNulasEva = True
        End If
        oCon.CierraConexion
    End If
    
    rsE.Close
    Set rsE = Nothing
    Set oCon = Nothing
    Exit Function
AgregaProSelecErr:
    
End Function


Public Function GetCodigoEva() As String
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    On Error GoTo AgregaProSelecErr
    
    sqlE = " Select Top 1 cRHSeleccionNro Cod From " & lsServerAdministracion & "RHSeleccion Order by cRHSeleccionNro Desc"
    If oCon.AbreConexion Then
        Set rsE = oCon.CargaRecordSet(sqlE)
        If rsE.EOF And rsE.BOF Then
            GetCodigoEva = "000001"
        Else
            GetCodigoEva = Format(CCur(rsE!Cod) + 1, "000000")
        End If
        oCon.CierraConexion
    End If
    
    rsE.Close
    Set rsE = Nothing
    Set oCon = Nothing
    Exit Function
AgregaProSelecErr:
    
End Function

'Elimina un proceso de Seleccion, la eliminacion es logica
'##ModelId=3A942952005B
Public Function EliminaProSelec(psRHProcSeleccNro As String) As Boolean
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    On Error GoTo EliminaProSelecErr
    
    If oCon.AbreConexion Then
        sqlE = " Delete " & lsServerAdministracion & "RHSeleccionComite Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"
        oCon.Ejecutar sqlE
        sqlE = " Delete " & lsServerAdministracion & "RHSeleccionDet Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"
        oCon.Ejecutar sqlE
        sqlE = " Delete " & lsServerAdministracion & "RHSeleccion Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"
        oCon.Ejecutar sqlE
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing

    Exit Function
EliminaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:EliminaProSelec Method")
End Function

'Elimina un proceso de Evalcion, la eliminacion es logica
'##ModelId=3A942952005B
Public Function EliminaProEval(psRHProcEvalcNro As String) As Boolean
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    On Error GoTo EliminaProEvalErr
    
    If oCon.AbreConexion Then
        oCon.BeginTrans
        sqlE = " Delete " & lsServerAdministracion & "RHEvaluacionComite Where cRHEvaluacionNro = '" & psRHProcEvalcNro & "'"
        oCon.Ejecutar sqlE
        sqlE = " Delete " & lsServerAdministracion & "RHEvaluacionRRHH Where cRHEvaluacionNro = '" & psRHProcEvalcNro & "'"
        oCon.Ejecutar sqlE
        sqlE = " Delete " & lsServerAdministracion & "RHEvaluacionDet Where cRHEvaluacionNro = '" & psRHProcEvalcNro & "'"
        oCon.Ejecutar sqlE
        sqlE = " Delete " & lsServerAdministracion & "RHEvaluacion Where cRHEvaluacionNro = '" & psRHProcEvalcNro & "'"
        oCon.Ejecutar sqlE
        oCon.CommitTrans
        oCon.CierraConexion
    End If
    Set oCon = Nothing

    Exit Function
EliminaProEvalErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvalcion:EliminaProEval Method")
End Function

'Actualiza un proceso de Seleccion
'##ModelId=3A94295200A1
'Public Function ModificaProSelec(psRHProcSeleccNro As String, psRHProcSeleccTpo As String, psRHProcSeleccAreaCod As String, psRHProcSeleccCargoCod As String, pdRHProcSeleccInicio As String, pdRHProcSeleccFin As String, psRHProcSeleccEstado As String, psRHProcSeleccComentario As String, ptRHProcSeleccExamenPsicologico As String, ptRHProcSeleccExamenEscrito As String, psRHProcSeleccExaCur As String, psRHProcSeleccExaEsc As String, psRHProcSeleccExaPsi As String, psRHProcSeleccExaEnt As String, psRHProcSeleccNotaMax As String, psRHProcSeleccConTpo As String, psUltimaActualizacion As String) As Boolean
'    Dim oCon As DConecta
'    Dim sqlE As String
'    Set oCon = New DConecta
'    On Error GoTo ModificaProSelecErr
'
'    sqlE = " Update " & lsServerAdministracion & "RHSeleccion" _
'         & " Set cRHSeleccionTpo = '" & psRHProcSeleccTpo & "', cRHSeleccionAreaCod = '" & psRHProcSeleccAreaCod & "',cRHSeleccionCargoCod = '" & psRHProcSeleccCargoCod & "'," _
'         & " dRHSeleccionInicio = '" & pdRHProcSeleccInicio & "',dRHSeleccionFin = '" & pdRHProcSeleccFin & "',cRHSeleccionEstado = '" & psRHProcSeleccEstado & "',cRHSeleccionComentario= '" & psRHProcSeleccComentario & "' ,tRHSeleccionExamenPsicologico = '" & ptRHProcSeleccExamenPsicologico & "',tRHSeleccionExamenEscrito= '" & ptRHProcSeleccExamenEscrito & "',cUltimaActualizacion = '" & psUltimaActualizacion & "'," _
'         & " bRHSeleccionExaCur = " & psRHProcSeleccExaCur & ", bRHSeleccionExaEsc = " & psRHProcSeleccExaEsc & ", bRHSeleccionExaPsico = " & psRHProcSeleccExaPsi & ",bRHSeleccionExaEnt = " & psRHProcSeleccExaEnt & ",nRHSeleccionNotaMaxima = " & psRHProcSeleccNotaMax & ", cRHSeleccionTipoContrato = '" & psRHProcSeleccConTpo & "'" _
'         & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"
'
'    If oCon.AbreConexion Then
'        oCon.Ejecutar sqlE
'        oCon.CierraConexion
'    End If
'
'    Set oCon = Nothing
'    Exit Function
'ModificaProSelecErr:
'    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:ModificaProSelec Method")
'End Function

'Actualiza un proceso de Seleccion
'##ModelId=3A94295200A1
Public Function ModificaEstadoProSelec(psRHProcSeleccNro As String, psRHProcSeleccEstado As String, psUltimaActualizacion As String) As Boolean
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    On Error GoTo ModificaProSelecErr

    sqlE = " Update " & lsServerAdministracion & "RHSeleccion " _
         & " Set cRHSeleccionEstado = '" & psRHProcSeleccEstado & "', cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
         & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlE
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:ModificaProSelec Method")
End Function

'Actualiza un proceso de Seleccion
'##ModelId=3A94295200A1
Public Function ModificaEstadoEva(psRHEvaluacionNro As String, psRHEvaluacionEstado As String, psUltimaActualizacion As String) As Boolean
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    On Error GoTo ModificaProSelecErr

    sqlE = " Update " & lsServerAdministracion & "RHEvaluacion " _
         & " Set cRHEvaluacionEstado = '" & psRHEvaluacionEstado & "', cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
         & " Where cRHEvaluacionNro = '" & psRHEvaluacionNro & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlE
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:ModificaProSelec Method")
End Function

'Actualiza un proceso de Seleccion
'##ModelId=3A94295200A1
Public Function ModificaProSelecComentarioFinal(psRHProcSeleccNro As String, psRHProcSeleccComentario As String, psUltimaActualizacion As String) As Boolean
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    On Error GoTo ModificaProSelecErr

    sqlE = " Update " & lsServerAdministracion & "RHSeleccion" _
         & " Set cRHSeleccionComentarioFinal = '" & psRHProcSeleccComentario & "',cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
         & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlE
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:ModificaProSelec Method")
End Function

'Actualiza un proceso de Seleccion
'##ModelId=3A94295200A1
Public Function ModificaEvaluacionComentarioFinal(psRHEvaluacionNro As String, psRHEvaluacionComentario As String, psUltimaActualizacion As String) As Boolean
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    On Error GoTo ModificaProSelecErr

    sqlE = " Update " & lsServerAdministracion & "RHEvaluacion" _
         & " Set cRHEvaluacionComentarioFinal = '" & psRHEvaluacionComentario & "',cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
         & " Where cRHEvaluacionNro = '" & psRHEvaluacionNro & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlE
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:ModificaProSelec Method")
End Function


'Actualiza un proceso de Evaluacion
'##ModelId=3A94295200A1
Public Function ModificaProEval(psRHProcEvalcNro As String, pdRHProcEvalcInicio As String, pdRHProcEvalcFin As String, psRHProcEvalcEstado As String, psRHProcEvalcComentario As String, ptRHProcEvalcExamenPsicologico As String, ptRHProcEvalcExamenEscrito As String, psUltimaActualizacion As String, prRHEvaArea As ADODB.Recordset, prRHEvaAreaEmp As ADODB.Recordset, psRHEvaluacionNotaMaxima As String, psRHEvaluacionExaCur As String, psRHEvaluacionExaEsc As String, psRHEvaluacionExaPsico As String, psRHEvaluacionExaEnt As String) As Boolean
    Dim oCon As DConecta
    Dim sqlE As String
    Dim lsArea As String
    Dim lbBan As Boolean
    Dim rsEva As ADODB.Recordset
    Set rsEva = New ADODB.Recordset
    Set oCon = New DConecta
    On Error GoTo ModificaProEvalErr

    sqlE = " Update " & lsServerAdministracion & "RHEvaluacion" _
         & " Set dRHEvaluacionInicio = '" & pdRHProcEvalcInicio & "',dRHEvaluacionFin = '" & pdRHProcEvalcFin & "',cRHEvaluacionEstado = '" & psRHProcEvalcEstado & "',cRHEvaluacionComentario= '" & psRHProcEvalcComentario & "' ,tRHEvaluacionExamenPsicologico = '" & ptRHProcEvalcExamenPsicologico & "',tRHEvaluacionExamenEscrito= '" & ptRHProcEvalcExamenEscrito & "',cUltimaActualizacion = '" & psUltimaActualizacion & "',nRHEvaluacionNotaMaxima= " & psRHEvaluacionNotaMaxima & ", bRHEvaluacionExaCur= " & psRHEvaluacionExaCur & ",bRHEvaluacionExaEsc = " & psRHEvaluacionExaEsc & ", bRHEvaluacionExaPsico = " & psRHEvaluacionExaPsico & " ,bRHEvaluacionExaEnt = " & psRHEvaluacionExaEnt & " " _
         & " Where cRHEvaluacionNro = '" & psRHProcEvalcNro & "'"
    
    lsArea = ""
    lbBan = True
    
    If Not prRHEvaArea Is Nothing Then
        If Not (prRHEvaArea.EOF And prRHEvaArea.BOF) Then
            While Not prRHEvaArea.EOF
                If lbBan Then
                    lsArea = lsArea & prRHEvaArea.Fields(0)
                    lbBan = False
                End If
                lsArea = lsArea & "','" & prRHEvaArea.Fields(0)
                prRHEvaArea.MoveNext
            Wend
            prRHEvaArea.MoveFirst
        End If
    End If
    
    If oCon.AbreConexion Then
        oCon.BeginTrans
            oCon.Ejecutar sqlE
            If Not prRHEvaArea Is Nothing Then
                If Not (prRHEvaArea.EOF And prRHEvaArea.BOF) Then
                
                    sqlE = " Delete RHEvaluacionRRHH" _
                         & "    Where cRHEvaluacionNro = '" & psRHProcEvalcNro & "' And" _
                         & " cAreaCod not in ('" & lsArea & "'); " _
                         & " Delete RHEvaluacionDet " _
                         & "    Where cRHEvaluacionNro = '" & psRHProcEvalcNro & "' And" _
                         & " cAreaCod not in ('" & lsArea & "') " _
                         & ""
                    oCon.Ejecutar sqlE
                    While Not prRHEvaArea.EOF
                        sqlE = "Select cRHEvaluacionNro from  RHEvaluacionDet Where cRHEvaluacionNro = '" & psRHProcEvalcNro & "' And cAreaCod = '" & prRHEvaArea.Fields(0) & "'"
                        Set rsEva = oCon.CargaRecordSet(sqlE)
                        If rsEva.EOF And rsEva.BOF Then
                            sqlE = " Insert RHEvaluaciondet(cRHEvaluacionNro, cAreaCod, cUltimaActualizacion)" _
                                 & " Values('" & psRHProcEvalcNro & "','" & prRHEvaArea.Fields(0) & "','" & psUltimaActualizacion & "')"
                            oCon.Ejecutar sqlE
                        End If
                        rsEva.Close
                        prRHEvaArea.MoveNext
                    Wend
                    
                    While Not prRHEvaAreaEmp.EOF
                        sqlE = "Select cRHEvaluacionNro from  RHEvaluacionRRHH Where cRHEvaluacionNro = '" & psRHProcEvalcNro & "' And cAreaCod = '" & prRHEvaAreaEmp.Fields(5) & "' And cPersCod = '" & prRHEvaAreaEmp.Fields(0) & "'"
                        Set rsEva = oCon.CargaRecordSet(sqlE)
                        If rsEva.EOF And rsEva.BOF Then
                            sqlE = " Insert RHEvaluacionRRHH(cRHEvaluacionNro,cAreaCod,cPersCod,bExonerado,cComentario,cUltimaActualizacion)" _
                                 & " Values('" & psRHProcEvalcNro & "','" & prRHEvaAreaEmp.Fields(5) & "','" & prRHEvaAreaEmp.Fields(0) & "','" & prRHEvaAreaEmp.Fields(2) & "','" & prRHEvaAreaEmp.Fields(3) & "','" & psUltimaActualizacion & "')"
                            oCon.Ejecutar sqlE
                        End If
                        rsEva.Close
                        prRHEvaAreaEmp.MoveNext
                    Wend
                End If
            End If
        oCon.CommitTrans
        oCon.CierraConexion
    End If
    
    Set rsEva = Nothing
    Set oCon = Nothing
    Exit Function
ModificaProEvalErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvaluacion:ModificaProEval Method")
End Function

'Clausura el proceso de seleccion, no se permite actualizar un proyecto cerrado
'##ModelId=3A94295200D3
Public Function CierraProSelec(psRHProcSeleccNro As String, psCodCierre As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo CierraProSelecErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    sqlE = " Update " & lsServerAdministracion & "RHSeleccion" _
         & " Set cRHSeleccionEstado = '" & psCodCierre & "',cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
         & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlE
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
CierraProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:CierraProSelec Method")
End Function


'Clausura el proceso de Evaluacion, no se permite actualizar un proyecto cerrado
'##ModelId=3A94295200D3
Public Function CierraProEval(psRHProcEvalcNro As String, psCodCierre As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo CierraProEvalErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    sqlE = " Update " & lsServerAdministracion & "RHEvaluacion" _
         & " Set cRHEvaluacionEstado = '" & psCodCierre & "',cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
         & " Where cRHEvaluacionNro = '" & psRHProcEvalcNro & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlE
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
CierraProEvalErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvaluacion:CierraProEval Method")
End Function

        'Select RHS.cRHSeleccionPersCod Codigo, PE.cPersNombre Nombre, RHS.nRHSeleccionNotaCurricular Exa_Curri, RHS.nRHSeleccionNotaEscrito Exa_Escrito, RHS.nRHSeleccionNotaPsicologico Exa_Psico, RHS.nRHSeleccionNotaEntrevista Exa_Entrev, RHS.cRHSeleccionResultado Seleccion
        'From RHSeleccionDet RHS
        'Inner Join Persona PE On RHS.cRHSeleccionPersCod = PE.cPersCod
        'where RHS.cRHSeleccionNro = ''

'##ModelId=3A942AA801B1
Public Function GetProcesoSeleccion(psRHSeleccionNro As String) As Recordset
    On Error GoTo GetProcesoSeleccionErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    sqlE = " Select cRHSeleccionTpo Tipo,cRHSeleccionAreaCod Area,cRHSeleccionCargoCod Cargo,dRHSeleccionInicio FI,dRHSeleccionFin FF,cRHSeleccionEstado Estado,cRHSeleccionComentario Comentario,tRHSeleccionExamenPsicologico Psico,tRHSeleccionExamenEscrito Escrito, cUltimaActualizacion, bRHSeleccionExaCur bCur , bRHSeleccionExaEsc bEsc, bRHSeleccionExaPsico bPsi, bRHSeleccionExaEnt bEnt, cRHSeleccionTipoContrato TipoContrato, nRHSeleccionNotaMaxima NotaMax, cRHSeleccionComentarioFinal ComentarioFinal  from RHSeleccion Where cRHSeleccionNro = '" & psRHSeleccionNro & "'"
    If oCon.AbreConexion Then
        Set GetProcesoSeleccion = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetProcesoSeleccionErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:GetProcesoSeleccion Method")
End Function

'##ModelId=3A942AA801B1
Public Function GetProcesoEvaluacion(psRHEvaluacionNro As String) As Recordset
    On Error GoTo GetProcesoEvaluacionErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    sqlE = " Select dRHEvaluacionInicio FI,dRHEvaluacionFin FF,cRHEvaluacionEstado Estado,cRHEvaluacionComentario Comentario,tRHEvaluacionExamenPsicologico Psico, tRHEvaluacionExamenEscrito Escrito, cUltimaActualizacion, nRHEvaluacionNotaMaxima NMax, bRHEvaluacionExaCur bCur, bRHEvaluacionExaEsc bEsc, bRHEvaluacionExaPsico bPsico, bRHEvaluacionExaEnt bEnt, cRHEvaluacionComentarioFinal Com from RHEvaluacion Where cRHEvaluacionNro = '" & psRHEvaluacionNro & "'"
    If oCon.AbreConexion Then
        Set GetProcesoEvaluacion = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetProcesoEvaluacionErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvaluacion:GetProcesoEvaluacion Method")
End Function

Public Function GetAreasEvaluacion(psRHEvaluacionNro As String) As Recordset
    On Error GoTo GetProcesoEvaluacionErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    sqlE = " Select RHE.cAreaCod Codigo ,cAreaDescripcion Area from RHEvaluacionDet RHE Inner Join Areas ARE On RHE.cAreaCod = ARE.cAreaCod Where cRHEvaluacionNro = '" & psRHEvaluacionNro & "' Order by RHE.cAreaCod"
    If oCon.AbreConexion Then
        Set GetAreasEvaluacion = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetProcesoEvaluacionErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvaluacion:GetProcesoEvaluacion Method")
End Function

Public Function GetAreasEvaluacionRRHH(psRHEvaluacionNro As String) As Recordset
    On Error GoTo GetProcesoEvaluacionErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    sqlE = " Select RHE.cPersCod Codigo , PE.cPersNombre, bExonerado Exo, cComentario Comentario, 1 Bit,cAreaCod From RHEvaluacionRRHH RHE" _
         & " Inner Join Persona PE On RHE.cPersCod = PE.cPersCod" _
         & " where cRHEvaluacionNro = '" & psRHEvaluacionNro & "'"
    If oCon.AbreConexion Then
        Set GetAreasEvaluacionRRHH = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetProcesoEvaluacionErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvaluacion:GetProcesoEvaluacion Method")
End Function

'-1 Todos
'0 Abiertos
'1 Cerrados
Public Function GetProcesosSeleccion(Optional pnTipos As Integer = 0) As Recordset
    On Error GoTo GetProcesoSeleccionErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    
    If pnTipos = -1 Then
        sqlE = " Select cRHSeleccionNro Codigo,cRHSeleccionComentario Comentario From RHSeleccion"
    ElseIf pnTipos = 0 Then
        sqlE = " Select cRHSeleccionNro Codigo,cRHSeleccionComentario Comentario From RHSeleccion Where cRHSeleccionEstado = '" & RHProcesoSeleccionEstado.gRHProcSelEstIniCiado & "'"
    ElseIf pnTipos = 1 Then
        sqlE = " Select cRHSeleccionNro Codigo,cRHSeleccionComentario Comentario, len(cRHSeleccionNro) Nivel From RHSeleccion Where cRHSeleccionEstado = '" & RHProcesoSeleccionEstado.gRHProcSelEstFinalizado & "'"
    End If
    
    If oCon.AbreConexion Then
        Set GetProcesosSeleccion = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetProcesoSeleccionErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:GetProcesoSeleccion Method")
End Function

Public Function GetProcesosEvaluacion(Optional pbSoloAbiertos As Boolean = False) As Recordset
    On Error GoTo GetProcesoEvaluacionErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    
    If Not pbSoloAbiertos Then
        sqlE = " Select cRHEvaluacionNro Codigo,cRHEvaluacionComentario Comentario From RHEvaluacion Order by cRHEvaluacionNro Desc"
    Else
        sqlE = " Select cRHEvaluacionNro Codigo,cRHEvaluacionComentario Comentario From RHEvaluacion Where cRHEvaluacionEstado = '" & RHProcesoSeleccionEstado.gRHProcSelEstIniCiado & "' Order by cRHEvaluacionNro Desc"
    End If
    If oCon.AbreConexion Then
        Set GetProcesosEvaluacion = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetProcesoEvaluacionErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvaluacion:GetProcesoEvaluacion Method")
End Function

Public Function GetProcesosSeleccionDet(psSeleccion As String, Optional pbGanador As Boolean = False) As Recordset
    On Error GoTo GetProcesoSeleccionErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    
    If Not pbGanador Then
        sqlE = " Select RH.cRHSeleccionPersCod cCodPers, cPersNombre Nombre from RHSeleccionDet RH Inner Join Persona PE On RH.cRHSeleccionPersCod = PE.cPersCod where cRHSeleccionNro = '" & psSeleccion & "'"
    Else
        sqlE = " Select RH.cRHSeleccionPersCod cCodPers, cPersNombre Nombre from RHSeleccionDet RH Inner Join Persona PE On RH.cRHSeleccionPersCod = PE.cPersCod where cRHSeleccionNro = '" & psSeleccion & "' And cRHSeleccionResultado = '1'"
    End If

    If oCon.AbreConexion Then
        Set GetProcesosSeleccionDet = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetProcesoSeleccionErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:GetProcesoSeleccion Method")
End Function

Public Function GetProcesosEvaluacionDet(psRHEvaluacionNro As String) As Recordset
    On Error GoTo GetProcesoEvaluacionErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
        
    sqlE = " Select RH.cAreaCod + RH.cPersCod cCodPers, cPersNombre Nombre" _
         & " From RHEvaluacionRRHH RH " _
         & " Inner Join Persona PE On RH.cPersCod = PE.cPersCod" _
         & " where cRHEvaluacionNro = '" & psRHEvaluacionNro & "'"

    If oCon.AbreConexion Then
        Set GetProcesosEvaluacionDet = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetProcesoEvaluacionErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvaluacion:GetProcesoEvaluacion Method")
End Function

Public Function GetProcesosSeleccionDetExamen(pcRHSeleccionNro As String, pnTipoExamen As Integer, Optional pbConConfirmacion As Boolean = False) As Recordset
    On Error GoTo GetProcesoSeleccionErr
    Dim oCon As DConecta
    Dim sqlE As String
    Dim sqlA As String
    Dim rsRes As ADODB.Recordset
    Set rsRes = New ADODB.Recordset
    Set oCon = New DConecta
    
    
    If oCon.AbreConexion Then
        If pnTipoExamen = RHTipoOpeEvaEscrito Then
            sqlE = " Select RH.cRHSeleccionPersCod cCodPers, cPersNombre Nombre, Convert(Varchar(8),nRHSeleccionNotaEscrito) Eva_Escrita, Activa = Case When nRHSeleccionNotaEscrito is null Then '1' else '0' end  from RHSeleccionDet RH Inner Join Persona PE On RH.cRHSeleccionPersCod = PE.cPersCod where cRHSeleccionNro = '" & pcRHSeleccionNro & "' Order by cPersNombre"
        ElseIf pnTipoExamen = RHTipoOpeEvaPsicologico Then
            sqlE = " Select RH.cRHSeleccionPersCod cCodPers, cPersNombre Nombre, Convert(Varchar(8),nRHSeleccionNotaPsicologico) Eva_Psicologica, Activa = Case When nRHSeleccionNotaPsicologico is null Then '1' else '0' end from RHSeleccionDet RH Inner Join Persona PE On RH.cRHSeleccionPersCod = PE.cPersCod where cRHSeleccionNro = '" & pcRHSeleccionNro & "' Order by cPersNombre"
        ElseIf pnTipoExamen = RHTipoOpeEvaEntrevista Then
            sqlE = " Select RH.cRHSeleccionPersCod cCodPers, cPersNombre Nombre, Convert(Varchar(8),nRHSeleccionNotaEntrevista) Eva_Entrevista, Activa = Case When nRHSeleccionNotaEntrevista is null Then '1' else '0' end from RHSeleccionDet RH Inner Join Persona PE On RH.cRHSeleccionPersCod = PE.cPersCod where cRHSeleccionNro = '" & pcRHSeleccionNro & "' Order by cPersNombre"
        ElseIf pnTipoExamen = RHTipoOpeEvaCurricular Then
            sqlE = " Select RH.cRHSeleccionPersCod cCodPers, cPersNombre Nombre, Convert(Varchar(8),nRHSeleccionNotaCurricular) Eva_Curricular, Activa = Case When nRHSeleccionNotaCurricular is null Then '1' else '0' end from RHSeleccionDet RH Inner Join Persona PE On RH.cRHSeleccionPersCod = PE.cPersCod where cRHSeleccionNro = '" & pcRHSeleccionNro & "' Order By cPersNombre"
        ElseIf pnTipoExamen = RHTipoOpeEvaConsolidado Then
            sqlA = "select convert(int,bRHSeleccionExaCur) + convert(int,bRHSeleccionExaEsc) + convert(int,bRHSeleccionExaPsico) + convert(int,bRHSeleccionExaEnt) div from rhseleccion where cRHSeleccionNro = '" & pcRHSeleccionNro & "'"
            Set rsRes = oCon.CargaRecordSet(sqlA)
            If Not (rsRes.EOF And rsRes.BOF) Then
                If Not pbConConfirmacion Then
                    sqlE = " Select RH.cRHSeleccionPersCod cCodPers, cPersNombre Nombre, Convert(Varchar(8),nRHSeleccionNotaEscrito) Eva_Escrita, Convert(Varchar(8),nRHSeleccionNotaPsicologico) Eva_Psicologica, Convert(Varchar(8),nRHSeleccionNotaEntrevista) Eva_Entrevista, Convert(Varchar(8),nRHSeleccionNotaCurricular) Eva_Curricular, convert(numeric(8,2),(case when nRHSeleccionNotaEscrito is null then 0 else nRHSeleccionNotaEscrito end + case when nRHSeleccionNotaPsicologico is null then 0 else nRHSeleccionNotaPsicologico end + case when nRHSeleccionNotaEntrevista is null then 0 else nRHSeleccionNotaEntrevista end + case when nRHSeleccionNotaCurricular is null then 0 else nRHSeleccionNotaCurricular end) / " & rsRes!Div & ") Promedio, cRHSeleccionResultado OK from RHSeleccionDet RH Inner Join Persona PE On RH.cRHSeleccionPersCod = PE.cPersCod where cRHSeleccionNro = '" & pcRHSeleccionNro & "' order by cPersNombre"
                Else
                    sqlE = " Select RH.cRHSeleccionPersCod cCodPers, cPersNombre Nombre, Convert(Varchar(8),nRHSeleccionNotaEscrito) Eva_Escrita, Convert(Varchar(8),nRHSeleccionNotaPsicologico) Eva_Psicologica, Convert(Varchar(8),nRHSeleccionNotaEntrevista) Eva_Entrevista, Convert(Varchar(8),nRHSeleccionNotaCurricular) Eva_Curricular, convert(numeric(8,2),(case when nRHSeleccionNotaEscrito is null then 0 else nRHSeleccionNotaEscrito end + case when nRHSeleccionNotaPsicologico is null then 0 else nRHSeleccionNotaPsicologico end + case when nRHSeleccionNotaEntrevista is null then 0 else nRHSeleccionNotaEntrevista end + case when nRHSeleccionNotaCurricular is null then 0 else nRHSeleccionNotaCurricular end) / " & rsRes!Div & ") Promedio, Case cRHSeleccionResultado When '1' Then 'SI' Else 'NO' end OK, Case When cRHSeleccionConfirmacion is not Null Then '1' Else '0' End "
                    sqlE = sqlE & " Conf from RHSeleccionDet RH Inner Join Persona PE On RH.cRHSeleccionPersCod = PE.cPersCod where cRHSeleccionNro = '" & pcRHSeleccionNro & "' order by cPersNombre"
                End If
            Else
                sqlE = " Select RH.cRHSeleccionPersCod from RHSeleccionDet RH Inner Join Persona PE On RH.cRHSeleccionPersCod = PE.cPersCod where cRHSeleccionNro = '" & pcRHSeleccionNro & "' order by cPersNombre"
            End If
            rsRes.Close
        End If
        Set GetProcesosSeleccionDetExamen = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    
    Set rsRes = Nothing
    Set oCon = Nothing
    Exit Function
GetProcesoSeleccionErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:GetProcesoSeleccion Method")
End Function

Public Function GetProcesosEvaluacionDetExamen(pcRHEvaluacion As String, pnTipoExamen As Integer) As Recordset
    On Error GoTo GetProcesoEvaluacionErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    If pnTipoExamen = 0 Then
        sqlE = " Select RH.cPersCod cCodPers, cPersNombre Nombre, nRHEvaluacionNotaEscrita Eva_Escrita, Activa = Case When nRHEvaluacionNotaEscrita is null And (bExonerado = 0 Or bExonerado is null) Then '1' When nRHEvaluacionNotaEscrita is not null And (bExonerado = 0 Or bExonerado is null) Then '2' else '0' end  from RHEvaluacionRRHH RH Inner Join Persona PE On RH.cPersCod = PE.cPersCod where cRHEvaluacionNro = '" & pcRHEvaluacion & "' Order by cPersNombre"
    ElseIf pnTipoExamen = 1 Then
        sqlE = " Select RH.cPersCod cCodPers, cPersNombre Nombre, nRHEvaluacionNotaPsicologica Eva_Psicologica, Activa = Case When nRHEvaluacionNotaPsicologica is null And (bExonerado = 0 Or bExonerado is null) Then '1' When nRHEvaluacionNotaPsicologica is not null And (bExonerado = 0 Or bExonerado is null) Then '2' else '0' end from RHEvaluacionRRHH RH Inner Join Persona PE On RH.cPersCod = PE.cPersCod where cRHEvaluacionNro = '" & pcRHEvaluacion & "' Order by cPersNombre"
    ElseIf pnTipoExamen = 2 Then
        sqlE = " Select RH.cPersCod cCodPers, cPersNombre Nombre, nRHEvaluacionNotaEntrevista Eva_Entrevista, Activa = Case When nRHEvaluacionNotaEntrevista is null And (bExonerado = 0 Or bExonerado is null) Then '1' When nRHEvaluacionNotaEntrevista Is Not null And (bExonerado = 0 Or bExonerado is null) Then '2' else '0' end from RHEvaluacionRRHH RH Inner Join Persona PE On RH.cPersCod = PE.cPersCod where cRHEvaluacionNro = '" & pcRHEvaluacion & "' Order by cPersNombre"
    ElseIf pnTipoExamen = 3 Then
        sqlE = " Select RH.cPersCod cCodPers, cPersNombre Nombre, nRHEvaluacionNotaCurricular Eva_Curricular, Activa = Case When nRHEvaluacionNotaCurricular is null And (bExonerado = 0 Or bExonerado is null) Then '1' When nRHEvaluacionNotaCurricular is not null And (bExonerado = 0 Or bExonerado is null) Then '2' else '0' end from RHEvaluacionRRHH RH Inner Join Persona PE On RH.cPersCod = PE.cPersCod where cRHEvaluacionNro = '" & pcRHEvaluacion & "' Order By cPersNombre"
    ElseIf pnTipoExamen = 4 Then
        sqlE = " Select RH.cPersCod cCodPers, cPersNombre Nombre, Case when nRHEvaluacionNotaEscrita Is null Then 0 Else nRHEvaluacionNotaEscrita End Eva_Escrita,  Case when nRHEvaluacionNotaPsicologica Is null Then 0 Else nRHEvaluacionNotaPsicologica End Eva_Psicologica, Case when nRHEvaluacionNotaEntrevista Is null Then 0 Else nRHEvaluacionNotaEntrevista End Eva_Entrevista, Case when nRHEvaluacionNotaCurricular Is null Then 0 Else nRHEvaluacionNotaCurricular End Eva_Curricular, Convert(numeric(8,2),(Case when nRHEvaluacionNotaEscrita Is null Then 0 Else nRHEvaluacionNotaEscrita End + Case when nRHEvaluacionNotaPsicologica Is null Then 0 Else nRHEvaluacionNotaPsicologica End        + Case when nRHEvaluacionNotaEntrevista Is null Then 0 Else nRHEvaluacionNotaEntrevista End + Case when nRHEvaluacionNotaCurricular Is null Then 0 Else nRHEvaluacionNotaCurricular End) / 4) Promedio, cRHEvaluacionResultado OK from RHEvaluacionRRHH RH Inner Join Persona PE On RH.cPersCod = PE.cPersCod "
        sqlE = sqlE & " where cRHEvaluacionNro = '" & pcRHEvaluacion & "' order by cPersNombre"
    End If
    If oCon.AbreConexion Then
        Set GetProcesosEvaluacionDetExamen = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetProcesoEvaluacionErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvaluacion:GetProcesoEvaluacion Method")
End Function


Public Function GetNomPersonasComite(psRHProcSeleccNro As String) As Recordset
    On Error GoTo GetNomPersonasComiteErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    sqlE = " Select distinct PE.cPersCod Codigo, cPersNombre Nombre, (cConsDescripcion + space(50) + convert(varchar(1),CO.nConsValor)) Cargo , cRHCargoDescripcion CargaCaja" _
         & " from rhSeleccioncomite PS" _
         & " Inner Join Persona PE On PE.cPersCod = PS.cPersCod" _
         & " LEFT  Join RHCargos RHC On RHC.cPersCod = PS.cPersCod  And RHC.dRHCargoFecha in (Select Max(dRHCargoFecha) From RHCargos RHC1 Where RHC1.cPersCod = PS.cPersCod)" _
         & " LEFT  Join RHCargosTabla RHTC On RHTC.cRHCargoCod = RHC.cRHCargoCod" _
         & " Inner Join Constante CO On CO.nConsValor = PS.nComiteCargo and CO.nConsCod = '" & gRHEvaluacionComite & "'" _
         & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"

    If oCon.AbreConexion Then
        Set GetNomPersonasComite = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetNomPersonasComiteErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:GetNomPersonasComite Method")
End Function

Public Function GetNomPersonasComiteEval(psRHProcSeleccNro As String) As Recordset
    On Error GoTo GetNomPersonasComiteErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    sqlE = " Select distinct PE.cPersCod Codigo, cPersNombre Nombre, (cConsDescripcion + space(50) + convert(varchar(1),CO.nConsValor)) Cargo , cRHCargoDescripcion CargaCaja" _
         & " From RHEvaluacionComite PS" _
         & " Inner Join Persona PE On PE.cPersCod = PS.cPersCod" _
         & " LEFT  Join RHCargos RHC On RHC.cPersCod = PS.cPersCod And RHC.dRHCargoFecha in (Select Max(dRHCargoFecha) From RHCargos RHC1 Where RHC1.cPersCod = PS.cPersCod)" _
         & " LEFT  Join RHCargosTabla RHTC On RHTC.cRHCargoCod = RHC.cRHCargoCod" _
         & " Inner Join Constante CO On CO.nConsValor = PS.nComiteCargo and CO.nConsCod = '" & gRHEvaluacionComite & "'" _
         & " Where cRHEvaluacionNro = '" & psRHProcSeleccNro & "'"

    If oCon.AbreConexion Then
        Set GetNomPersonasComiteEval = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetNomPersonasComiteErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:GetNomPersonasComite Method")
End Function


Public Function GetNomCargo(psRHProcSeleccNro As String) As String
    On Error GoTo GetNomPersonasComiteErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    sqlE = " Select cRHCargoDescripcion from rhSeleccion PS " _
         & " Inner Join RHCargosTabla RHTC On PS.cRHSeleccionCargoCod = RHTC.cRHCargoCod" _
         & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"
    If oCon.AbreConexion Then
        GetNomCargo = oCon.CargaRecordSet(sqlE).Fields(0)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetNomPersonasComiteErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:GetNomPersonasComite Method")
End Function

'Agrega una persona a un Proceso de Seleccion
'##ModelId=3A942D610376
Public Function AgregaPersonaProSelec(psRHProcSeleccNro As String, psRHProcSeleccPersCod As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaPersonaProSelecErr
    Dim oCon As DConecta
    Dim sqlE As String
    Dim sqlAdd As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Set oCon = New DConecta



    sqlE = " Insert RHSeleccionDet (cRHSeleccionNro, cRHSeleccionPersCod,cUltimaActualizacion)" _
         & " Values ('" & psRHProcSeleccNro & "','" & psRHProcSeleccPersCod & "','" & psUltimaActualizacion & "')"
    
    sqlAdd = " Select cRHSeleccionNro From RHSeleccionDet " _
           & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "' And cRHSeleccionPersCod = '" & psRHProcSeleccPersCod & "'"
    
    If oCon.AbreConexion Then
        Set rsE = oCon.CargaRecordSet(sqlAdd)
        If rsE.EOF And rsE.BOF Then
            If psRHProcSeleccNro <> "" Then oCon.Ejecutar sqlE
        End If
        oCon.CierraConexion
    End If
    
    rsE.Close
    Set rsE = Nothing
    Set oCon = Nothing
        
    Exit Function
AgregaPersonaProSelecErr:
    oCon.RollBackTrans
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:AgregaPersonaProSelec Method")
End Function

'Agrega una persona a un Proceso de Evaluacion
'##ModelId=3A942D610376
Public Function AgregaPersonasEvaluacion(psRHProcSeleccNro As String, prRHPersonas As ADODB.Recordset, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaPersonaProSelecErr
    Dim oCon As DConecta
    Dim sqlE As String
    Dim sqlAdd As String
    Dim rsE As ADODB.Recordset
    Set oCon = New DConecta
    
    If Not (prRHPersonas.EOF And prRHPersonas.BOF) Then
        sqlAdd = ""
        While Not prRHPersonas.EOF
            If sqlAdd = "" Then
                sqlAdd = Mid(prRHPersonas.Fields(0), 4)
            Else
                sqlAdd = sqlAdd & "','" & Mid(prRHPersonas.Fields(0), 4)
            End If
            prRHPersonas.MoveNext
        Wend
        
        prRHPersonas.MoveFirst
        
        If oCon.AbreConexion Then
            oCon.BeginTrans
            sqlE = " Delete RHEvaluacionRRHH Where cRHEvaluacionNro = '" & psRHProcSeleccNro & "' And cPersCod not in ('" & sqlAdd & "')"
            oCon.Ejecutar sqlE
            While Not prRHPersonas.EOF
                sqlAdd = "Select cRHEvaluacionNro From RHEvaluacionRRHH  Where cPersCod = '" & Right(prRHPersonas.Fields(0), 13) & "' And cRHEvaluacionNro = '" & psRHProcSeleccNro & "'"
                Set rsE = oCon.CargaRecordSet(sqlAdd)
                If rsE.EOF And rsE.BOF Then
                    sqlE = " Insert RHEvaluacionRRHH (cRHEvaluacionNro,cAreaCod,cPersCod,cUltimaActualizacion)" _
                    & " Values ('" & psRHProcSeleccNro & "','" & Left(prRHPersonas.Fields(0), 3) & "','" & Right(prRHPersonas.Fields(0), 13) & "','" & psUltimaActualizacion & "')"
                    oCon.Ejecutar sqlE
                End If
                prRHPersonas.MoveNext
            Wend
            oCon.CommitTrans
            oCon.CierraConexion
        End If
    End If
    Set oCon = Nothing
    
    Exit Function
AgregaPersonaProSelecErr:
    oCon.RollBackTrans
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvaluacion:AgregaPersonaProSelec Method")
End Function


'Agrega una persona a un Proceso de Seleccion
'##ModelId=3A942D610376
Public Function AgregaPersonaComite(psRHProcSeleccNro As String, prRHProcSeleccComite As ADODB.Recordset, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaPersonaProSelecErr
    Dim oCon As DConecta
    Dim sqlD As String
    Dim sqlE As String
    Set oCon = New DConecta
    
    If prRHProcSeleccComite Is Nothing Then
        If oCon.AbreConexion Then
            sqlD = " Delete RHSeleccionComite" _
                 & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"
            oCon.Ejecutar sqlD
            oCon.CierraConexion
        End If
        Set oCon = Nothing
        Exit Function
    End If
    
    If Not (prRHProcSeleccComite.EOF And prRHProcSeleccComite.BOF) Then
        If oCon.AbreConexion Then
            oCon.BeginTrans
            sqlD = " Delete RHSeleccionComite" _
                 & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"
            oCon.Ejecutar sqlD
            While Not prRHProcSeleccComite.EOF
                sqlE = " Insert RHSeleccionComite (cRHSeleccionNro,cPersCod,nComiteCargo,cUltimaActualizacion)" _
                     & " Values ('" & psRHProcSeleccNro & "','" & prRHProcSeleccComite.Fields(0) & "','" & Right(prRHProcSeleccComite.Fields(2), 1) & "','" & psUltimaActualizacion & "')"
                oCon.Ejecutar sqlE
                prRHProcSeleccComite.MoveNext
            Wend
            oCon.CommitTrans
        End If
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
AgregaPersonaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:AgregaPersonaProSelec Method")
End Function

'Agrega una persona a un Proceso de Evalion
'##ModelId=3A942D610376
Public Function AgregaPersonaComiteEval(psRHProcEvalNro As String, prRHProcEvalComite As ADODB.Recordset, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaPersonaProSelecErr
    Dim oCon As DConecta
    Dim sqlD As String
    Dim sqlE As String
    Set oCon = New DConecta
    
    If Not prRHProcEvalComite Is Nothing Then
        If Not (prRHProcEvalComite.EOF And prRHProcEvalComite.BOF) Then
            If oCon.AbreConexion Then
                oCon.BeginTrans
                While Not prRHProcEvalComite.EOF
                    sqlD = " Delete RHEvaluacionComite" _
                         & " Where cRHEvaluacionNro = '" & psRHProcEvalNro & "' And cPersCod =  '" & prRHProcEvalComite.Fields(0) & "'"
                    sqlE = " Insert RHEvaluacionComite (cRHEvaluacionNro,cPersCod,nComiteCargo,cUltimaActualizacion)" _
                         & " Values ('" & psRHProcEvalNro & "','" & prRHProcEvalComite.Fields(0) & "','" & Right(prRHProcEvalComite.Fields(2), 1) & "','" & psUltimaActualizacion & "')"
                    oCon.Ejecutar sqlD
                    oCon.Ejecutar sqlE
                    prRHProcEvalComite.MoveNext
                Wend
                oCon.CommitTrans
            End If
            oCon.CierraConexion
        End If
    End If
    Set oCon = Nothing
    Exit Function
AgregaPersonaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvalion:AgregaPersonaProSelec Method")
End Function


'Agrega una persona a un Proceso de Evalion
'##ModelId=3A942D610376
Public Function AgregaAreasEvaluacion(psRHProcEvalNro As String, prRHProcEvalAreas As ADODB.Recordset, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaPersonaProSelecErr
    Dim oCon As DConecta
    Dim sqlD As String
    Dim sqlE As String
    Set oCon = New DConecta
    
    If Not (prRHProcEvalAreas.EOF And prRHProcEvalAreas.BOF) Then
        If oCon.AbreConexion Then
            oCon.BeginTrans
            While Not prRHProcEvalAreas.EOF
                sqlD = " Delete RHEvaluacionComite" _
                     & " Where cRHEvaluacionNro = '" & psRHProcEvalNro & "' And cPersCod =  '" & prRHProcEvalAreas.Fields(0) & "'"
                sqlE = " Insert RHEvaluacionComite (cRHEvaluacionNro,cPersCod,nComiteCargo,cUltimaActualizacion)" _
                     & " Values ('" & psRHProcEvalNro & "','" & prRHProcEvalAreas.Fields(0) & "','" & Right(prRHProcEvalAreas.Fields(2), 1) & "','" & psUltimaActualizacion & "')"
                oCon.Ejecutar sqlD
                oCon.Ejecutar sqlE
                prRHProcEvalAreas.MoveNext
            Wend
            oCon.CommitTrans
        End If
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
AgregaPersonaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvalion:AgregaPersonaProSelec Method")
End Function


'Elimina una persona en un Proceso de Seleccion
'##ModelId=3A942D6103B2
Public Function EliminaPersonaProSelec(psRHProcSeleccNro As String, psRHProcSeleccPersCod As String) As Boolean
    On Error GoTo EliminaPersonaProSelecErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    sqlE = " Delete RHSeleccionDet Where cRHSeleccionNro = '" & psRHProcSeleccNro & "' And cRHSeleccionPersCod = '" & psRHProcSeleccPersCod & "'"
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlE
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
EliminaPersonaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:EliminaPersonaProSelec Method")
End Function

'Elimina una persona en un Proceso de Evaluacion
'##ModelId=3A942D6103B2
Public Function EliminaPersonaProEval(psRHProcEvalcNro As String, psRHProcEvalcPersCod As String) As Boolean
    On Error GoTo EliminaPersonaProEvalErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    sqlE = " Delete RHEvaluacionDet Where cRHEvaluacionNro = '" & psRHProcEvalcNro & "' And cRHEvaluacionPersCod = '" & psRHProcEvalcPersCod & "'"
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlE
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
EliminaPersonaProEvalErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvaluacion:EliminaPersonaProEval Method")
End Function


'Elimina una persona en un Proceso de Seleccion
'##ModelId=3A942D6103B2
Public Function EliminaPersonasProSelec(psRHProcSeleccNro As String, rsRH As ADODB.Recordset) As Boolean
    On Error GoTo EliminaPersonaProSelecErr
    Dim oCon As DConecta
    Dim sqlE As String
    Dim sqlAdd As String
    Set oCon = New DConecta
    
    sqlAdd = ""
    
    If rsRH Is Nothing Then
        sqlAdd = ""
    Else
        If Not (rsRH.EOF And rsRH.BOF) Then
            While Not rsRH.EOF
                If sqlAdd = "" Then
                    sqlAdd = rsRH.Fields(0)
                End If
                sqlAdd = sqlAdd & "','" & rsRH.Fields(0)
                rsRH.MoveNext
            Wend
        End If
    End If
    
    sqlE = " Delete RHSeleccionDet Where cRHSeleccionNro = '" & psRHProcSeleccNro & "' and cRHSeleccionPersCod not in ('" & sqlAdd & "')"
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlE
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
EliminaPersonaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:EliminaPersonaProSelec Method")
End Function

'Elimina una persona en un Proceso de Seleccion
'##ModelId=3A942D6103B2
Public Function EliminaComiteProSelec(psRHProcSeleccNro As String, psRHProcSeleccPersCod As String) As Boolean
    On Error GoTo EliminaComiteProSelecErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    sqlE = " Delete RHSeleccionComite Where cRHSeleccionNro = '" & psRHProcSeleccNro & "' And cPersCod = '" & psRHProcSeleccPersCod & "'"
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlE
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
EliminaComiteProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:EliminaPersonaProSelec Method")
End Function

'Actualiza una persona en un Proceso de Seleccion
'##ModelId=3A942D6103E4
Public Function ModificaPersonaProSelec(psRHProcSeleccNro As String, psRHrs As ADODB.Recordset, pnTpoEva As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaPersonaProSelecErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    Dim sqlPsi As String
    Dim sqlEst As String
    Dim sqlCur As String
    Dim sqlEnt As String
    
    If Not (psRHrs.EOF And psRHrs.BOF) Then
        If oCon.AbreConexion Then
            psRHrs.MoveFirst
            oCon.BeginTrans
            While Not psRHrs.EOF
                If pnTpoEva = 4 Then
                    If psRHrs.Fields(3) <> 0 Or psRHrs.Fields(5) <> 0 Or psRHrs.Fields(4) <> 0 Or psRHrs.Fields(7) <> 0 Or psRHrs.Fields(2) <> 0 Then
                        
                        If psRHrs.Fields(3) <> "" Then
                            sqlPsi = "nRHSeleccionNotaPsicologico = " & psRHrs.Fields(3) & ""
                        Else
                            sqlPsi = "nRHSeleccionNotaPsicologico = 0 "
                        End If
                        
                        If psRHrs.Fields(2) <> "" Then
                            sqlEst = "nRHSeleccionNotaEscrito = " & psRHrs.Fields(2) & ""
                        Else
                            sqlEst = "nRHSeleccionNotaEscrito = 0"
                        End If
                        
                        If psRHrs.Fields(5) <> "" Then
                            sqlCur = "nRHSeleccionNotaCurricular = " & psRHrs.Fields(5) & ""
                        Else
                            sqlCur = "nRHSeleccionNotaCurricular = 0"
                        End If
                        
                        If psRHrs.Fields(4) <> "" Then
                            sqlEnt = "nRHSeleccionNotaEntrevista = " & psRHrs.Fields(4) & ""
                        Else
                            sqlEnt = "nRHSeleccionNotaEntrevista = 0"
                        End If
                        
                        If sqlPsi <> "" Or sqlEst <> "" Or sqlCur <> "" Or sqlEnt <> "" Then
                            sqlE = " Update RHSeleccionDet" _
                                 & " Set " & sqlPsi & " , " & sqlEst & " , " & sqlCur & " , " & sqlEnt & ",  cUltimaActualizacion = '" & psUltimaActualizacion & "', cRHSeleccionResultado = '" & psRHrs.Fields(7) & "' " _
                                 & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "' And cRHSeleccionPersCod = '" & psRHrs.Fields(0) & "'"
                        End If
                    End If
                ElseIf pnTpoEva = 3 And psRHrs.Fields(2) <> 0 And psRHrs.Fields(2) <> "" Then
                    sqlE = " Update RHSeleccionDet" _
                         & " Set nRHSeleccionNotaCurricular = " & psRHrs.Fields(2) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                         & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "' And cRHSeleccionPersCod = '" & psRHrs.Fields(0) & "'"
                ElseIf pnTpoEva = 2 And psRHrs.Fields(2) <> 0 And psRHrs.Fields(2) <> "" Then
                    sqlE = " Update RHSeleccionDet" _
                         & " Set nRHSeleccionNotaEntrevista = " & psRHrs.Fields(2) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                         & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "' And cRHSeleccionPersCod = '" & psRHrs.Fields(0) & "'"
                ElseIf pnTpoEva = 1 And psRHrs.Fields(2) <> 0 And psRHrs.Fields(2) <> "" Then
                    sqlE = " Update RHSeleccionDet" _
                         & " Set nRHSeleccionNotaPsicologico = " & psRHrs.Fields(2) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                         & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "' And cRHSeleccionPersCod = '" & psRHrs.Fields(0) & "'"
                ElseIf pnTpoEva = 0 And psRHrs.Fields(2) <> 0 And psRHrs.Fields(2) <> "" Then
                    sqlE = " Update RHSeleccionDet" _
                         & " Set nRHSeleccionNotaEscrito = " & psRHrs.Fields(2) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                         & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "' And cRHSeleccionPersCod = '" & psRHrs.Fields(0) & "'"
                End If
                If sqlE <> "" Then oCon.Ejecutar sqlE
                sqlE = ""
                psRHrs.MoveNext
            Wend
            oCon.CommitTrans
            oCon.CierraConexion
        End If
    End If
    Set oCon = Nothing
    
    Exit Function
ModificaPersonaProSelecErr:
    oCon.RollBackTrans
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:ModificaPersonaProSelec Method")
End Function

'Actualiza una persona en un Proceso de Evaluacion
'##ModelId=3A942D6103E4
Public Function ModificaPersonaProEval(psRHProcEvalNro As String, psRHrs As ADODB.Recordset, pnTpoEva As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaPersonaProSelecErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    
    If Not (psRHrs.EOF And psRHrs.BOF) Then
        If oCon.AbreConexion Then
            psRHrs.MoveFirst
            oCon.BeginTrans
            While Not psRHrs.EOF
                If pnTpoEva = 4 Then
                    If psRHrs.Fields(3) <> 0 Or psRHrs.Fields(5) <> 0 Or psRHrs.Fields(4) <> 0 Or psRHrs.Fields(7) <> 0 Or psRHrs.Fields(2) <> 0 Then
                        sqlE = " Update RHEvaluacionRRHH" _
                             & " Set nRHEvaluacionNotaPsicologica = " & psRHrs.Fields(3) & ", nRHEvaluacionNotaEscrita = " & psRHrs.Fields(2) & ", nRHEvaluacionNotaCurricular = " & psRHrs.Fields(5) & ", nRHEvaluacionNotaEntrevista = " & psRHrs.Fields(4) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "', cRHEvaluacionResultado = '" & psRHrs.Fields(7) & "' " _
                             & " Where cRHEvaluacionNro = '" & psRHProcEvalNro & "' And cPersCod = '" & psRHrs.Fields(0) & "'"
                    End If
                ElseIf pnTpoEva = 3 And psRHrs.Fields(2) <> 0 And psRHrs.Fields(2) <> "" Then
                    sqlE = " Update RHEvaluacionRRHH" _
                         & " Set nRHEvaluacionNotaCurricular = " & psRHrs.Fields(2) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                         & " Where cRHEvaluacionNro = '" & psRHProcEvalNro & "' And cPersCod = '" & psRHrs.Fields(0) & "'"
                ElseIf pnTpoEva = 2 And psRHrs.Fields(2) <> 0 And psRHrs.Fields(2) <> "" Then
                    sqlE = " Update RHEvaluacionRRHH" _
                         & " Set nRHEvaluacionNotaEntrevista = " & psRHrs.Fields(2) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                         & " Where cRHEvaluacionNro = '" & psRHProcEvalNro & "' And cPersCod = '" & psRHrs.Fields(0) & "'"
                ElseIf pnTpoEva = 1 And psRHrs.Fields(2) <> 0 And psRHrs.Fields(2) <> "" Then
                    sqlE = " Update RHEvaluacionRRHH" _
                         & " Set nRHEvaluacionNotaPsicologica = " & psRHrs.Fields(2) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                         & " Where cRHEvaluacionNro = '" & psRHProcEvalNro & "' And cPersCod = '" & psRHrs.Fields(0) & "'"
                ElseIf pnTpoEva = 0 And psRHrs.Fields(2) <> 0 And psRHrs.Fields(2) <> "" Then
                    sqlE = " Update RHEvaluacionRRHH" _
                         & " Set nRHEvaluacionNotaEscrita = " & psRHrs.Fields(2) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                         & " Where cRHEvaluacionNro = '" & psRHProcEvalNro & "' And cPersCod = '" & psRHrs.Fields(0) & "'"
                End If
                If sqlE <> "" Then oCon.Ejecutar sqlE
                sqlE = ""
                psRHrs.MoveNext
            Wend
            oCon.CommitTrans
            oCon.CierraConexion
        End If
    End If
    Set oCon = Nothing
    
    Exit Function
ModificaPersonaProSelecErr:
    oCon.RollBackTrans
    Call RaiseError(MyUnhandledError, "DActualizaProcesoEvaluacion:ModificaPersonaProSelec Method")
End Function

'Actualiza una persona en un Proceso de Seleccion
'##ModelId=3A942D6103E4
Public Function ModificaComiteProSelec(psRHProcSeleccNro As String, prRHProcComite As ADODB.Recordset, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaComiteProSelecErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    If Not (prRHProcComite.BOF And prRHProcComite.EOF) Then
        If oCon.AbreConexion Then
            While Not prRHProcComite.EOF
                sqlE = " Update rhevaluacioncomite" _
                     & " Set nComiteCargo = '" & Trim(Right(prRHProcComite.Fields(1), 3)) & "', cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                     & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "' And cPersCod = '" & prRHProcComite.Fields(1) & "'"
                oCon.Ejecutar sqlE
                prRHProcComite.MoveNext
            Wend
            oCon.CierraConexion
        End If
    End If
    Set oCon = Nothing
    Exit Function
ModificaComiteProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:ModificaPersonaProSelec Method")
End Function

'##ModelId=3AB902EB01D3
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

'Actualiza una persona en un Proceso de Seleccion
'##ModelId=3A942C0E00A7
Public Function GetReporte() As Recordset
    On Error GoTo GetReporteErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta

    sqlE = " Select EV.cRHSeleccionNro Cod, CO1.cConsDescripcion Tipo, AR.cAreaDescripcion Area, CA.cRHCargoDescripcion Cargo, dRHSeleccionInicio FechaIni, dRHSeleccionFin FechaFin, CO3.cConsDescripcion Estado, cRHSeleccionComentario from RHSeleccion EV" _
         & " Inner Join Constante CO1 On EV.cRHSeleccionTpo = CO1.nConsValor and CO1.nConsCod = '" & gRHProcesoSeleccionTipo & "'" _
         & " " _
         & " Inner Join Constante CO3 On EV.cRHSeleccionEstado = CO3.nConsValor and CO3.nConsCod = '" & gRHProcesoSeleccionEstado & "'" _
         & " Inner Join Areas AR On EV.cRHSeleccionAreaCod = AR.cAreaCod" _
         & " Inner Join RHCargosTabla CA On EV.cRHSeleccionCargoCod = CA.cRHCargoCod"

    If oCon.AbreConexion Then
        Set GetReporte = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetReporteErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:ModificaPersonaProSelec Method")
End Function

'Actualiza una persona en un Proceso de Seleccion
'##ModelId=3A942C0E00A7
Public Function PersonaComite(psSeleccionCod As String, psPersCod As String) As Boolean
    On Error GoTo GetReporteErr
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    sqlE = " Select cRHSeleccionNro from rhseleccioncomite Where cRHSeleccionNro = '" & psSeleccionCod & "' And cPersCod = '" & psPersCod & "'"

    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlE)
        
        If rs.EOF And rs.BOF Then
            PersonaComite = False
        Else
            PersonaComite = True
        End If
        
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetReporteErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:ModificaPersonaProSelec Method")
End Function

'##ModelId=3AB902EB020F
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

#If DebugMode Then
    '##ModelId=3AB902EB00E2
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If


Public Function ModificaProSelec(psRHProcSeleccNro As String, psRHProcSeleccTpo As String, psRHProcSeleccAreaCod As String, psRHProcSeleccCargoCod As String, pdRHProcSeleccInicio As String, pdRHProcSeleccFin As String, psRHProcSeleccEstado As String, psRHProcSeleccComentario As String, ptRHProcSeleccExamenPsicologico As String, ptRHProcSeleccExamenEscrito As String, psRHProcSeleccExaCur As String, psRHProcSeleccExaEsc As String, psRHProcSeleccExaPsi As String, psRHProcSeleccExaEnt As String, psRHProcSeleccNotaMax As String, psRHProcSeleccConTpo As String, psUltimaActualizacion As String, pnRHPesoExaCur As Integer, pnRHPesoExaEsc As Integer, pnRHPesoExaPsico As Integer, pnRHPesoExaEnt As Integer) As Boolean
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    On Error GoTo ModificaProSelecErr

    sqlE = " Update " & lsServerAdministracion & "RHSeleccion" _
         & " Set cRHSeleccionTpo = '" & psRHProcSeleccTpo & "', cRHSeleccionAreaCod = '" & psRHProcSeleccAreaCod & "',cRHSeleccionCargoCod = '" & psRHProcSeleccCargoCod & "'," _
         & " dRHSeleccionInicio = '" & pdRHProcSeleccInicio & "',dRHSeleccionFin = '" & pdRHProcSeleccFin & "',cRHSeleccionEstado = '" & psRHProcSeleccEstado & "',cRHSeleccionComentario= '" & psRHProcSeleccComentario & "' ,tRHSeleccionExamenPsicologico = '" & ptRHProcSeleccExamenPsicologico & "',tRHSeleccionExamenEscrito= '" & ptRHProcSeleccExamenEscrito & "',cUltimaActualizacion = '" & psUltimaActualizacion & "'," _
         & " bRHSeleccionExaCur = " & psRHProcSeleccExaCur & ", bRHSeleccionExaEsc = " & psRHProcSeleccExaEsc & ", bRHSeleccionExaPsico = " & psRHProcSeleccExaPsi & ",bRHSeleccionExaEnt = " & psRHProcSeleccExaEnt & ",nRHSeleccionNotaMaxima = " & psRHProcSeleccNotaMax & ", cRHSeleccionTipoContrato = '" & psRHProcSeleccConTpo & "' ," _
         & " nRHPesoExaCur = " & pnRHPesoExaCur & ", " _
         & " nRHPesoExaEsc = " & pnRHPesoExaEsc & ",  " _
         & " nRHPesoExaPsico =" & pnRHPesoExaPsico & ", " _
         & " nRHPesoExaEnt =" & pnRHPesoExaEnt & " " _
         & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlE
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:ModificaProSelec Method")
End Function


Public Function AgregaProSelec(psRHProcSeleccNro As String, psRHProcSeleccTpo As String, psRHProcSeleccAreaCod As String, psRHProcSeleccCargoCod As String, pdRHProcSeleccInicio As String, pdRHProcSeleccFin As String, psRHProcSeleccEstado As String, psRHProcSeleccComentario As String, ptRHProcSeleccExamenPsicologico As String, ptRHProcSeleccExamenEscrito As String, psRHProcSeleccExaCur As String, psRHProcSeleccExaEsc As String, psRHProcSeleccExaPsi As String, psRHProcSeleccExaEnt As String, psRHProcSeleccNotaMax As String, psRHProcSeleccConTpo As String, psUltimaActualizacion As String, pnRHPesoExaCur As Integer, pnRHPesoExaEsc As Integer, pnRHPesoExaPsico As Integer, pnRHPesoExaEnt As Integer) As Boolean
    Dim oCon As DConecta
    Dim sqlE As String
    Set oCon = New DConecta
    On Error GoTo AgregaProSelecErr
    
    sqlE = " Insert " & lsServerAdministracion & "RHSeleccion (cRHSeleccionNro,cRHSeleccionTpo,cRHSeleccionAreaCod,cRHSeleccionCargoCod,dRHSeleccionInicio,dRHSeleccionFin,cRHSeleccionEstado,cRHSeleccionComentario,tRHSeleccionExamenPsicologico,tRHSeleccionExamenEscrito,cUltimaActualizacion,  bRHSeleccionExaCur,bRHSeleccionExaEsc,bRHSeleccionExaPsico,bRHSeleccionExaEnt,nRHSeleccionNotaMaxima,cRHSeleccionTipoContrato,nRHPesoExaCur,nRHPesoExaEsc,nRHPesoExaPsico,nRHPesoExaEnt)" _
         & " Values('" & psRHProcSeleccNro & "','" & psRHProcSeleccTpo & "','" & psRHProcSeleccAreaCod & "','" & psRHProcSeleccCargoCod & "','" & pdRHProcSeleccInicio & "','" & pdRHProcSeleccFin & "','" & psRHProcSeleccEstado & "','" & psRHProcSeleccComentario & "','','','" & psUltimaActualizacion & "'," & psRHProcSeleccExaCur & "," & psRHProcSeleccExaEsc & "," & psRHProcSeleccExaPsi & "," & psRHProcSeleccExaEnt & "," & psRHProcSeleccNotaMax & ",'" & psRHProcSeleccConTpo & "'," & pnRHPesoExaCur & "," & pnRHPesoExaEsc & "," & pnRHPesoExaPsico & "," & pnRHPesoExaEnt & ")"
    If oCon.AbreConexion Then
        
        oCon.Ejecutar sqlE
        sqlE = " Update RHSeleccion " _
             & " Set tRHSeleccionExamenPsicologico = '" & ptRHProcSeleccExamenPsicologico & "', tRHSeleccionExamenEscrito = '" & ptRHProcSeleccExamenEscrito & "'" _
             & " Where cRHSeleccionNro = '" & psRHProcSeleccNro & "'"
        oCon.Ejecutar sqlE
        
        oCon.CierraConexion
        AgregaProSelec = True
    Else
        AgregaProSelec = False
    End If
    
    Set oCon = Nothing
    Exit Function
    AgregaProSelec = False
AgregaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:AgregaProSelec Method")
End Function
