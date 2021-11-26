VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaProcesoSeleccion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A93FAB00159"
'Actualiza el Proceso de Seleccion es decir, lo crea, le asigna participantes,
'lo actualiza , lo elimina  y lo cierra.
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
    '##ModelId=3AB902CE027B
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902CE036C
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

Public Sub Inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub

'##ModelId=3AB902CE0358
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

'Agrega un nuevo Proceso de Seleccion
'##ModelId=3A9425BE028E
'Public Function AgregaProSelec(psRHProcSeleccNro As String, psRHProcSeleccTpo As String, psRHProcSeleccAreaCod As String, psRHProcSeleccCargoCod As String, pdRHProcSeleccInicio As String, pdRHProcSeleccFin As String, psRHProcSeleccEstado As String, psRHProcSeleccComentario As String, ptRHProcSeleccExamenPsicologico As String, ptRHProcSeleccExamenEscrito As String, psRHProcSeleccExaCur As String, psRHProcSeleccExaEsc As String, psRHProcSeleccExaPsi As String, psRHProcSeleccExaEnt As String, psRHProcSeleccNotaMax As String, psRHProcSeleccConTpo As String, psUltimaActualizacion As String) As Boolean
'    On Error GoTo AgregaProSelecErr
'    Dim oEva As DActualizaProcesoSeleccion
'    Dim lsCodigo As String
'    Set oEva = New DActualizaProcesoSeleccion
'
'    lsCodigo = oEva.GetCodigoEva()
'    psRHProcSeleccNro = lsCodigo
'    oEva.AgregaProSelec lsCodigo, Trim(psRHProcSeleccTpo), psRHProcSeleccAreaCod, psRHProcSeleccCargoCod, pdRHProcSeleccInicio, pdRHProcSeleccFin, psRHProcSeleccEstado, psRHProcSeleccComentario, ptRHProcSeleccExamenPsicologico, ptRHProcSeleccExamenEscrito, psRHProcSeleccExaCur, psRHProcSeleccExaEsc, psRHProcSeleccExaPsi, psRHProcSeleccExaEnt, psRHProcSeleccNotaMax, psRHProcSeleccConTpo, psUltimaActualizacion
'
'    Set oEva = Nothing
'    Exit Function
'AgregaProSelecErr:
'    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:AgregaProSelec Method")
'End Function

'Agrega un nuevo Proceso de Seleccion
'##ModelId=3A9425BE028E
Public Function AgregaProEval(psRHProcSeleccNro As String, pdRHProcSeleccInicio As String, pdRHProcSeleccFin As String, psRHProcSeleccEstado As String, psRHProcSeleccComentario As String, ptRHProcSeleccExamenPsicologico As String, ptRHProcSeleccExamenEscrito As String, psUltimaActualizacion As String, prRHEvaArea As ADODB.Recordset, prRHEvaAreaEmp As ADODB.Recordset, psRHEvaluacionNotaMaxima As String, psRHEvaluacionExaCur As String, psRHEvaluacionExaEsc As String, psRHEvaluacionExaPsico As String, psRHEvaluacionExaEnt As String) As Boolean
    On Error GoTo AgregaProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Dim lsCodigo As String
    Set oEva = New DActualizaProcesoSeleccion
    
    lsCodigo = oEva.GetCodigoEvaluacion()
    psRHProcSeleccNro = lsCodigo
    oEva.AgregaProEval lsCodigo, pdRHProcSeleccInicio, pdRHProcSeleccFin, psRHProcSeleccEstado, psRHProcSeleccComentario, ptRHProcSeleccExamenPsicologico, ptRHProcSeleccExamenEscrito, psUltimaActualizacion, prRHEvaArea, prRHEvaAreaEmp, psRHEvaluacionNotaMaxima, psRHEvaluacionExaCur, psRHEvaluacionExaEsc, psRHEvaluacionExaPsico, psRHEvaluacionExaEnt
       
    Set oEva = Nothing
    Exit Function
AgregaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:AgregaProSelec Method")
End Function

'Elimina un proceso de Seleccion, la eliminacion es logica
'##ModelId=3A9425BE02C0
Public Function EliminaProSelec(psRHProcSeleccNro As String) As Boolean
    On Error GoTo EliminaProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.EliminaProSelec psRHProcSeleccNro
       
    Set oEva = Nothing
    Exit Function
EliminaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:EliminaProSelec Method")
End Function


'Elimina un proceso de Seleccion, la eliminacion es logica
'##ModelId=3A9425BE02C0
Public Function EliminaProEval(psRHProcSeleccNro As String) As Boolean
    On Error GoTo EliminaProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.EliminaProEval psRHProcSeleccNro
       
    Set oEva = Nothing
    Exit Function
EliminaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:EliminaProSelec Method")
End Function

'Actualiza un proceso de Seleccion
'##ModelId=3A9425BE02E8
'Public Function ModificaProSelec(psRHProcSeleccNro As String, psRHProcSeleccTpo As String, psRHProcSeleccAreaCod As String, psRHProcSeleccCargoCod As String, pdRHProcSeleccInicio As String, pdRHProcSeleccFin As String, psRHProcSeleccEstado As String, psRHProcSeleccComentario As String, ptRHProcSeleccExamenPsicologico As String, ptRHProcSeleccExamenEscrito As String, psRHProcSeleccExaCur As String, psRHProcSeleccExaEsc As String, psRHProcSeleccExaPsi As String, psRHProcSeleccExaEnt As String, psRHProcSeleccNotaMax As String, psRHProcSeleccConTpo As String, psUltimaActualizacion As String) As Boolean
'    On Error GoTo ModificaProSelecErr
'    Dim oEva As DActualizaProcesoSeleccion
'    Set oEva = New DActualizaProcesoSeleccion
'
'    oEva.ModificaProSelec psRHProcSeleccNro, psRHProcSeleccTpo, psRHProcSeleccAreaCod, psRHProcSeleccCargoCod, pdRHProcSeleccInicio, pdRHProcSeleccFin, psRHProcSeleccEstado, psRHProcSeleccComentario, ptRHProcSeleccExamenPsicologico, ptRHProcSeleccExamenEscrito, psRHProcSeleccExaCur, psRHProcSeleccExaEsc, psRHProcSeleccExaPsi, psRHProcSeleccExaEnt, psRHProcSeleccNotaMax, psRHProcSeleccConTpo, psUltimaActualizacion
'
'    Set oEva = Nothing
'    Exit Function
'ModificaProSelecErr:
'    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:ModificaProSelec Method")
'End Function

'Actualiza un proceso de Seleccion
'##ModelId=3A9425BE02E8
Public Function ModificaEstadoProSelec(psRHProcSeleccNro As String, psRHProcSeleccEstado As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.ModificaEstadoProSelec psRHProcSeleccNro, psRHProcSeleccEstado, psUltimaActualizacion
    
    Set oEva = Nothing
    Exit Function
ModificaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:ModificaProSelec Method")
End Function

'Actualiza un proceso de Seleccion
'##ModelId=3A9425BE02E8
Public Function ModificaEstadoEvaluacion(psRHEvaluacionNro As String, psRHEvaluacionEstado As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.ModificaEstadoEva psRHEvaluacionNro, psRHEvaluacionEstado, psUltimaActualizacion
    
    Set oEva = Nothing
    Exit Function
ModificaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:ModificaProSelec Method")
End Function


'Actualiza un proceso de Evalcion
'##ModelId=3A9425BE02E8
Public Function ModificaProEval(psRHProcEvalcNro As String, pdRHProcEvalcInicio As String, pdRHProcEvalcFin As String, psRHProcEvalcEstado As String, psRHProcEvalcComentario As String, ptRHProcEvalcExamenPsicologico As String, ptRHProcEvalcExamenEscrito As String, psUltimaActualizacion As String, prRHEvaArea As ADODB.Recordset, prRHEvaAreaEmp As ADODB.Recordset, psRHEvaluacionNotaMaxima As String, psRHEvaluacionExaCur As String, psRHEvaluacionExaEsc As String, psRHEvaluacionExaPsico As String, psRHEvaluacionExaEnt As String) As Boolean
    On Error GoTo ModificaProEvalErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.ModificaProEval psRHProcEvalcNro, pdRHProcEvalcInicio, pdRHProcEvalcFin, psRHProcEvalcEstado, psRHProcEvalcComentario, ptRHProcEvalcExamenPsicologico, ptRHProcEvalcExamenEscrito, psUltimaActualizacion, prRHEvaArea, prRHEvaAreaEmp, psRHEvaluacionNotaMaxima, psRHEvaluacionExaCur, psRHEvaluacionExaEsc, psRHEvaluacionExaPsico, psRHEvaluacionExaEnt
       
    Set oEva = Nothing
    Exit Function
ModificaProEvalErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoEvalcion:ModificaProEval Method")
End Function

'Clausura el proceso de seleccion, no se permite actualizar un proyecto cerrado
'##ModelId=3A94260F0082
Public Function CierraProSelec(psRHProcSeleccNro As String, psCodCierre As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo CierraProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.CierraProSelec psRHProcSeleccNro, psCodCierre, psUltimaActualizacion
       
    Set oEva = Nothing
    Exit Function
CierraProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:CierraProSelec Method")
End Function

'Agrega una persona a un Proceso de Seleccion
'##ModelId=3A942C0E004D
Public Function AgregaPersonaProSelec(psRHProcSeleccNro As String, psRHProcSeleccPersCod As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaPersonaProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.AgregaPersonaProSelec psRHProcSeleccNro, psRHProcSeleccPersCod, psUltimaActualizacion
       
    Set oEva = Nothing
    Exit Function
AgregaPersonaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:AgregaPersonaProSelec Method")
End Function

'Agrega una persona a un Proceso de Seleccion
'##ModelId=3A942C0E004D
Public Function AgregaComiteProSelec(psRHProcSeleccNro As String, prRHProcSeleccComite As ADODB.Recordset, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaPersonaProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.AgregaPersonaComite psRHProcSeleccNro, prRHProcSeleccComite, psUltimaActualizacion
    
    Set oEva = Nothing
    Exit Function
AgregaPersonaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:AgregaPersonaProSelec Method")
End Function

'Agrega una persona a un Proceso de Evalcion
'##ModelId=3A942C0E004D
Public Function AgregaComiteProEval(psRHProcEvalcNro As String, prRHProcEvalcComite As ADODB.Recordset, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaPersonaProEvalErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.AgregaPersonaComiteEval psRHProcEvalcNro, prRHProcEvalcComite, psUltimaActualizacion
    
    Set oEva = Nothing
    Exit Function
AgregaPersonaProEvalErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoEvalcion:AgregaPersonaProEval Method")
End Function

'Elimina una persona en un Proceso de Seleccion
'##ModelId=3A942C0E0075
Public Function EliminaPersonaProSelec(psRHProcSeleccNro As String, psRHProcSeleccPersCod As String) As Boolean
    On Error GoTo EliminaPersonaProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.EliminaPersonaProSelec psRHProcSeleccNro, psRHProcSeleccPersCod
       
    Set oEva = Nothing
    Exit Function
EliminaPersonaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:EliminaPersonaProSelec Method")
End Function

'Elimina una persona en un Proceso de Seleccion
'##ModelId=3A942C0E0075
Public Function EliminaComiteProSelec(psRHProcSeleccNro As String, psRHProcSeleccPersCod As String) As Boolean
    On Error GoTo EliminaPersonaProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.EliminaComiteProSelec psRHProcSeleccNro, psRHProcSeleccPersCod
       
    Set oEva = Nothing
    Exit Function
EliminaPersonaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:EliminaPersonaProSelec Method")
End Function

'Actualiza una persona en un Proceso de Evalcion
'##ModelId=3A942C0E00A7
Public Function ModificaPersonaProEval(psRHProcEvalNro As String, psRHrs As ADODB.Recordset, pnTpoEva As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaPersonaProEvalErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion

    oEva.ModificaPersonaProEval psRHProcEvalNro, psRHrs, pnTpoEva, psUltimaActualizacion

    Set oEva = Nothing
    Exit Function
ModificaPersonaProEvalErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoEvalcion:ModificaPersonaProEval Method")
End Function


'Actualiza una persona en un Proceso de Seleccion
'##ModelId=3A942C0E00A7
Public Function ModificaPersonaProSelec(psRHProcSeleccNro As String, psRHrs As ADODB.Recordset, pnTpoEva As Integer, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaPersonaProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.ModificaPersonaProSelec psRHProcSeleccNro, psRHrs, pnTpoEva, psUltimaActualizacion
       
    Set oEva = Nothing
    Exit Function
ModificaPersonaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:ModificaPersonaProSelec Method")
End Function

'Actualiza una persona en un Proceso de Evalion
'##ModelId=3A942C0E00A7
Public Function ModificaComiteProEval(psRHProcEvalNro As String, prRHProcEvalComite As ADODB.Recordset, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaComiteProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.ModificaComiteProSelec psRHProcEvalNro, prRHProcEvalComite, psUltimaActualizacion
       
    Set oEva = Nothing
    Exit Function
ModificaComiteProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoEvalion:ModificaPersonaProSelec Method")
End Function

'Actualiza una persona en un Proceso de Seleccion
'##ModelId=3A942C0E00A7
Public Function GetReporte(pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo GetReporteErr
    Dim Cod As String * 8
    Dim Tipo As String * 12
    Dim Area As String * 30
    Dim Cargo As String * 30
    Dim FechaIni As String * 10
    Dim FechaFin As String * 10
    Dim Estado As String * 15
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long

    Dim rs As Recordset
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    Set rs = oEva.GetReporte
    
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("Evaluaciones", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Codigo;8;Tipo;9;Area;10; ;20;Cargo;18; ;20;F.Ini;10;F.Fin;10;Estado;10; ;3;", lnItem)
        While Not rs.EOF
            Cod = rs!Cod
            Tipo = rs!Tipo
            'Modo = rs!Modo
            Area = rs!Area
            Cargo = rs!Cargo
            FechaIni = rs!FechaIni
            FechaFin = rs!FechaFin
            Estado = rs!Estado
            
            lsCadena = lsCadena & " " & Cod & " " & Tipo & " " & Area & " " & Cargo & "  " & FechaIni & "  " & FechaFin & " " & Estado & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            
            If lnItem > 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Evaluaciones", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Codigo;8;Tipo;9;Area;10; ;10;Cargo;18; ;10;F.Ini;10;F.Fin;10;Estado;10; ;6;", lnItem)
            End If
            
            rs.MoveNext
        Wend
    End If
    
    GetReporte = lsCadena
    Exit Function
GetReporteErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:ModificaPersonaProSelec Method")
End Function

'Actualiza una persona en un Proceso de Seleccion
'##ModelId=3A942C0E00A7
Public Function GetActa(psSeleccionNro As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo GetReporteErr
    Dim lsCod As String
    Dim richT As RichTextBox
    
    Dim lsHora As String
    Dim lsFecha As String
    Dim lsPostulantes As String
    Dim lsPostulantesNotas As String
    Dim lsPostulantesTotalIng As String
    Dim lsComite As String
    
    Dim lsCadena As String
    Dim rsC As ADODB.Recordset
    Dim rsP As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Set rsP = New ADODB.Recordset
    
    Dim lsNombre As String * 30
    Dim lsCargo As String * 20
    Dim lsCargoCaja As String * 20
    Dim lsNota1 As String * 7
    Dim lsNota2 As String * 7
    Dim lsNota3 As String * 7
    Dim lsNota4 As String * 7
    Dim lsNomCargo As String
     
    Dim lsNombreArch As String
    Dim lnLong As Long
    lsNombreArch = App.Path & "\Txt\ActaSeleccion.txt"
    
    Open lsNombreArch For Input As #1
    lnLong = LOF(1)
    lsNombreArch = Input$(lnLong, #1)
    Close #1
    lsHora = Format(pgdFecSis, "hh:mm:ss AMPM")
    lsFecha = Format(pgdFecSis, "dddd, d mmm yyyy")
    
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    Set rsC = oEva.GetNomPersonasComite(psSeleccionNro)
    Set rsP = oEva.GetProcesosSeleccionDetExamen(psSeleccionNro, 4)
    lsNomCargo = oEva.GetNomCargo(psSeleccionNro)
    
    lsComite = oImpresora.gPrnSaltoLinea
    lsNombre = "NOMBRE"
    lsCargo = "CARGO"
    lsComite = lsComite & lsNombre & lsCargo & oImpresora.gPrnSaltoLinea
    If Not (rsC.EOF And rsC.BOF) Then
        While Not rsC.EOF
            lsNombre = PstaNombre(rsC.Fields(1))
            lsCargo = rsC.Fields(2)
            lsCargoCaja = IIf(IsNull(rsC.Fields(3)), "", rsC.Fields(3))
            lsComite = lsComite & lsNombre & lsCargo & lsCargoCaja & oImpresora.gPrnSaltoLinea
            rsC.MoveNext
        Wend
    End If
    
    lsPostulantesNotas = oImpresora.gPrnSaltoLinea
    lsPostulantes = oImpresora.gPrnSaltoLinea
    lsPostulantesTotalIng = oImpresora.gPrnSaltoLinea
    lsNombre = "NOMBRE"
    RSet lsCargo = "Promedio"
    lsPostulantesTotalIng = lsPostulantesTotalIng & lsNombre & lsCargo & oImpresora.gPrnSaltoLinea
    lsNota1 = "Escrit."
    lsNota2 = "Psicol."
    lsNota3 = "Curric."
    lsNota4 = "Entrevi."
    lsPostulantesNotas = lsPostulantesNotas & lsNombre & " " & " " & lsNota1 & " " & lsNota2 & " " & lsNota3 & " " & lsNota4 & " " & lsCargo & oImpresora.gPrnSaltoLinea
    lsPostulantes = lsPostulantes & lsNombre & oImpresora.gPrnSaltoLinea
    If Not (rsP.EOF And rsP.BOF) Then
        While Not rsP.EOF
            lsNombre = PstaNombre(rsP.Fields(1))
            RSet lsCargo = Format(rsP.Fields(6), "000.00")
            RSet lsNota4 = Format(rsP.Fields(5), "000.00")
            RSet lsNota3 = Format(rsP.Fields(4), "000.00")
            RSet lsNota2 = Format(rsP.Fields(3), "000.00")
            RSet lsNota1 = Format(rsP.Fields(2), "000.00")
            lsPostulantesNotas = lsPostulantesNotas & lsNombre & " " & " " & lsNota1 & " " & lsNota2 & " " & lsNota3 & " " & lsNota4 & " " & lsCargo & oImpresora.gPrnSaltoLinea
            lsPostulantes = lsPostulantes & lsNombre & oImpresora.gPrnSaltoLinea
            If rsP.Fields(7) = 1 Then lsPostulantesTotalIng = lsPostulantesTotalIng & lsNombre & " " & lsCargo & oImpresora.gPrnSaltoLinea
            rsP.MoveNext
        Wend
    End If
    
    'rsP.Sort = 1
    
    lsCadena = lsNombreArch
    lsCadena = Replace(lsCadena, "<<HORA>>", lsHora)
    lsCadena = Replace(lsCadena, "<<FECHA>>", lsFecha)
    lsCadena = Replace(lsCadena, "<<COMITE>>", lsComite)
    lsCadena = Replace(lsCadena, "<<CARGO>>", lsNomCargo)
    lsCadena = Replace(lsCadena, "<<POSTULANTES_NOTAS>>", lsPostulantesNotas)
    lsCadena = Replace(lsCadena, "<<POSTULANTES>>", lsPostulantes)
    lsCadena = Replace(lsCadena, "<<POSTULANTES_INGRESANTES>>", lsPostulantesTotalIng)
    
    rsC.Close
    rsP.Close
    Set rsC = Nothing
    Set rsP = Nothing
    Set oEva = Nothing
    GetActa = lsCadena
    Exit Function
GetReporteErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:ModificaPersonaProSelec Method")
End Function


'Actualiza una persona en un Proceso de Seleccion
'##ModelId=3A942C0E00A7
Public Function GetReporteEvaPersonas(psRHEvaCod As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo GetReporteErr
    Dim lsCodigo As String * 25
    Dim lsNombre As String * 80
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long

    Dim rs As Recordset
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    Set rs = oEva.GetProcesosSeleccionDet(psRHEvaCod)
    
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("Postulantes de la Evalucion - " & psRHEvaCod, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Codigo;8; ;18;Nombre;18; ;75;", lnItem)
        While Not rs.EOF
            lsCodigo = rs.Fields(0)
            lsNombre = PstaNombre(rs.Fields(1))
            
            lsCadena = lsCadena & "  " & lsCodigo & "  " & lsNombre & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            
            If lnItem > 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Postulantes de la Evalucion - " & psRHEvaCod, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Codigo;8; ;18;Nombre;18; ;75;", lnItem)
            End If
            
            rs.MoveNext
        Wend
    End If
    
    GetReporteEvaPersonas = lsCadena
    Exit Function
GetReporteErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:ModificaPersonaProSelec Method")
End Function

'Actualiza una persona en un Proceso de Seleccion
'##ModelId=3A942C0E00A7
Public Function GetReporteEvaPersonasNotas(psRHEvaCod As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, pnTpoEva As Integer) As String
    On Error GoTo GetReporteErr
    Dim lsCodigo As String * 13
    Dim lsNombre As String * 48
    Dim lsNota As String * 15
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsIng As String * 1
    
    Dim lsColumna As String

    Dim rs As Recordset
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    Set rs = oEva.GetProcesosEvaluacionDetExamen(psRHEvaCod, pnTpoEva)
        
    If pnTpoEva = 0 Then
        lsColumna = "EXAMEN ESCRITO"
    ElseIf pnTpoEva = 1 Then
        lsColumna = "EXAMEN PSICOLOGICO"
    ElseIf pnTpoEva = 2 Then
        lsColumna = "EXAMEN ENTREVISTA"
    ElseIf pnTpoEva = 3 Then
        lsColumna = "EXAMEN CURRICULAR"
    ElseIf pnTpoEva = 4 Then
        lsColumna = "EXAMEN CONSOLIDADO"
    End If
    
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("Postulantes de la Evalucion - " & psRHEvaCod, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        If pnTpoEva = 4 Then
            lsCadena = lsCadena & Encabezado("Codigo;8; ;8;Nombre;12; ;38;ESCRITO;10; ;8;PSICOLOG;10; ;8;ENTREVIS;8; ;8;CURRICU;7; ;2;PROMEDIO;10;ING;4;", lnItem)
        Else
            lsCadena = lsCadena & Encabezado("Codigo;8; ;8;Nombre;12; ;38;" & lsColumna & ";30; ;23;", lnItem)
        End If
        While Not rs.EOF
            lsNombre = PstaNombre(rs.Fields(1))
            If pnTpoEva = 4 Then
                lsCodigo = rs.Fields(0)
                RSet lsNota = Format(rs.Fields(2), "#,##0.00")
                lsCadena = lsCadena & lsCodigo & "  " & lsNombre & lsNota
                
                RSet lsNota = Format(rs.Fields(3), "#,##0.00")
                lsCadena = lsCadena & lsNota
                
                RSet lsNota = Format(rs.Fields(4), "#,##0.00")
                lsCadena = lsCadena & lsNota
                
                RSet lsNota = Format(rs.Fields(5), "#,##0.00")
                lsCadena = lsCadena & lsNota
                
                RSet lsNota = Format(rs.Fields(6), "#,##0.00")
                lsCadena = lsCadena & lsNota
                
                RSet lsIng = IIf(rs.Fields(7) = "0", "", "X")
                lsCadena = lsCadena & Space(2) & lsIng & oImpresora.gPrnSaltoLinea
                
            Else
                lsCodigo = rs.Fields(0)
                RSet lsNota = Format(rs.Fields(2), "#,##0.00")
                lsCadena = lsCadena & lsCodigo & "   " & lsNombre & Space(18) & lsNota & oImpresora.gPrnSaltoLinea
            End If
            
            lnItem = lnItem + 1
            
            If lnItem > 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                If pnTpoEva = 4 Then
                    lsCadena = lsCadena & Encabezado("Codigo;8; ;8;Nombre;12; ;38;ESCRITO;10; ;8;PSICOLOG;10; ;8;ENTREVIS;8; ;8;CURRICU;7; ;2;PROMEDIO;11;", lnItem)
                Else
                    lsCadena = lsCadena & Encabezado("Codigo;8; ;8;Nombre;12; ;38;" & lsColumna & ";30; ;23;", lnItem)
                End If
            End If
            
            rs.MoveNext
        Wend
    End If
    
    GetReporteEvaPersonasNotas = lsCadena
    Exit Function
GetReporteErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:ModificaPersonaProSelec Method")
End Function


'Actualiza una persona en un Proceso de Seleccion
'##ModelId=3A942C0E00A7
Public Function GetReporteSelectPersonasNotas(psRHEvaCod As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, pnTpoEva As Integer) As String
    On Error GoTo GetReporteErr
    Dim lsCodigo As String * 13
    Dim lsNombre As String * 48
    Dim lsNota As String * 15
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsIng As String * 1
    
    Dim lsColumna As String

    Dim rs As Recordset
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    Set rs = oEva.GetProcesosSeleccionDetExamen(psRHEvaCod, pnTpoEva)
        
    If pnTpoEva = 0 Then
        lsColumna = "EXAMEN ESCRITO"
    ElseIf pnTpoEva = 1 Then
        lsColumna = "EXAMEN PSICOLOGICO"
    ElseIf pnTpoEva = 2 Then
        lsColumna = "EXAMEN ENTREVISTA"
    ElseIf pnTpoEva = 3 Then
        lsColumna = "EXAMEN CURRICULAR"
    ElseIf pnTpoEva = 4 Then
        lsColumna = "EXAMEN CONSOLIDADO"
    End If
    
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("Postulantes de la Evalucion - " & psRHEvaCod, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        If pnTpoEva = 4 Then
            lsCadena = lsCadena & Encabezado("Codigo;8; ;8;Nombre;12; ;38;ESCRITO;10; ;8;PSICOLOG;10; ;8;ENTREVIS;8; ;8;CURRICU;7; ;2;PROMEDIO;10;ING;4;", lnItem)
        Else
            lsCadena = lsCadena & Encabezado("Codigo;8; ;8;Nombre;12; ;38;" & lsColumna & ";30; ;23;", lnItem)
        End If
        While Not rs.EOF
            lsNombre = PstaNombre(rs.Fields(1))
            If pnTpoEva = 4 Then
                lsCodigo = rs.Fields(0)
                RSet lsNota = Format(rs.Fields(2), "#,##0.00")
                lsCadena = lsCadena & lsCodigo & "  " & lsNombre & lsNota
                
                RSet lsNota = Format(rs.Fields(3), "#,##0.00")
                lsCadena = lsCadena & lsNota
                
                RSet lsNota = Format(rs.Fields(4), "#,##0.00")
                lsCadena = lsCadena & lsNota
                
                RSet lsNota = Format(rs.Fields(5), "#,##0.00")
                lsCadena = lsCadena & lsNota
                
                RSet lsNota = Format(rs.Fields(6), "#,##0.00")
                lsCadena = lsCadena & lsNota
                
                RSet lsIng = IIf(rs.Fields(7) = "0", "", "X")
                lsCadena = lsCadena & Space(2) & lsIng & oImpresora.gPrnSaltoLinea
                
            Else
                lsCodigo = rs.Fields(0)
                RSet lsNota = Format(rs.Fields(2), "#,##0.00")
                lsCadena = lsCadena & lsCodigo & "   " & lsNombre & Space(18) & lsNota & oImpresora.gPrnSaltoLinea
            End If
            
            lnItem = lnItem + 1
            
            If lnItem > 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                If pnTpoEva = 4 Then
                    lsCadena = lsCadena & Encabezado("Codigo;8; ;8;Nombre;12; ;38;ESCRITO;10; ;8;PSICOLOG;10; ;8;ENTREVIS;8; ;8;CURRICU;7; ;2;PROMEDIO;11;", lnItem)
                Else
                    lsCadena = lsCadena & Encabezado("Codigo;8; ;8;Nombre;12; ;38;" & lsColumna & ";30; ;23;", lnItem)
                End If
            End If
            
            rs.MoveNext
        Wend
    End If
    
    GetReporteSelectPersonasNotas = lsCadena
    Exit Function
GetReporteErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:ModificaPersonaProSelec Method")
End Function


Public Function ModificaProSelecComentarioFinal(psRHProcSeleccNro As String, psRHProcSeleccComentario As String, psUltimaActualizacion As String) As Boolean
    Dim oPro As DActualizaProcesoSeleccion
    Dim sqlE As String
    Set oPro = New DActualizaProcesoSeleccion
    On Error GoTo ModificaProSelecErr
    
    oPro.ModificaProSelecComentarioFinal psRHProcSeleccNro, psRHProcSeleccComentario, psUltimaActualizacion
    
    Set oPro = Nothing
    Exit Function
ModificaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:ModificaProSelec Method")
End Function

Public Function ModificaProSelecComentarioFinalEval(psRHEvalucionNro As String, psRHEvaluacionComentario As String, psUltimaActualizacion As String) As Boolean
    Dim oPro As DActualizaProcesoSeleccion
    Dim sqlE As String
    Set oPro = New DActualizaProcesoSeleccion
    On Error GoTo ModificaProSelecErr
    
    oPro.ModificaEvaluacionComentarioFinal psRHEvalucionNro, psRHEvaluacionComentario, psUltimaActualizacion
    
    Set oPro = Nothing
    Exit Function
ModificaProSelecErr:
    Call RaiseError(MyUnhandledError, "DActualizaProcesoSeleccion:ModificaProSelec Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902CE0299
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Public Function ModificaProSelec(psRHProcSeleccNro As String, psRHProcSeleccTpo As String, psRHProcSeleccAreaCod As String, psRHProcSeleccCargoCod As String, pdRHProcSeleccInicio As String, pdRHProcSeleccFin As String, psRHProcSeleccEstado As String, psRHProcSeleccComentario As String, ptRHProcSeleccExamenPsicologico As String, ptRHProcSeleccExamenEscrito As String, psRHProcSeleccExaCur As String, psRHProcSeleccExaEsc As String, psRHProcSeleccExaPsi As String, psRHProcSeleccExaEnt As String, psRHProcSeleccNotaMax As String, psRHProcSeleccConTpo As String, psUltimaActualizacion As String, pnRHPesoExaCur As Integer, pnRHPesoExaEsc As Integer, pnRHPesoExaPsico As Integer, pnRHPesoExaEnt As Integer) As Boolean
    On Error GoTo ModificaProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    
    oEva.ModificaProSelec psRHProcSeleccNro, psRHProcSeleccTpo, psRHProcSeleccAreaCod, psRHProcSeleccCargoCod, pdRHProcSeleccInicio, pdRHProcSeleccFin, psRHProcSeleccEstado, psRHProcSeleccComentario, ptRHProcSeleccExamenPsicologico, ptRHProcSeleccExamenEscrito, psRHProcSeleccExaCur, psRHProcSeleccExaEsc, psRHProcSeleccExaPsi, psRHProcSeleccExaEnt, psRHProcSeleccNotaMax, psRHProcSeleccConTpo, psUltimaActualizacion, pnRHPesoExaCur, pnRHPesoExaEsc, pnRHPesoExaPsico, pnRHPesoExaEnt
       
    Set oEva = Nothing
    Exit Function
ModificaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:ModificaProSelec Method")
End Function

Public Function AgregaProSelec(psRHProcSeleccNro As String, psRHProcSeleccTpo As String, psRHProcSeleccAreaCod As String, psRHProcSeleccCargoCod As String, pdRHProcSeleccInicio As String, pdRHProcSeleccFin As String, psRHProcSeleccEstado As String, psRHProcSeleccComentario As String, ptRHProcSeleccExamenPsicologico As String, ptRHProcSeleccExamenEscrito As String, psRHProcSeleccExaCur As String, psRHProcSeleccExaEsc As String, psRHProcSeleccExaPsi As String, psRHProcSeleccExaEnt As String, psRHProcSeleccNotaMax As String, psRHProcSeleccConTpo As String, psUltimaActualizacion As String, pnRHPesoExaCur As Integer, pnRHPesoExaEsc As Integer, pnRHPesoExaPsico As Integer, pnRHPesoExaEnt As Integer) As Boolean
    On Error GoTo AgregaProSelecErr
    Dim oEva As DActualizaProcesoSeleccion
    Dim lsCodigo As String
    Set oEva = New DActualizaProcesoSeleccion
    
    lsCodigo = oEva.GetCodigoEva()
    psRHProcSeleccNro = lsCodigo
    oEva.AgregaProSelec lsCodigo, Trim(psRHProcSeleccTpo), psRHProcSeleccAreaCod, psRHProcSeleccCargoCod, pdRHProcSeleccInicio, pdRHProcSeleccFin, psRHProcSeleccEstado, psRHProcSeleccComentario, ptRHProcSeleccExamenPsicologico, ptRHProcSeleccExamenEscrito, psRHProcSeleccExaCur, psRHProcSeleccExaEsc, psRHProcSeleccExaPsi, psRHProcSeleccExaEnt, psRHProcSeleccNotaMax, psRHProcSeleccConTpo, psUltimaActualizacion, pnRHPesoExaCur, pnRHPesoExaEsc, pnRHPesoExaPsico, pnRHPesoExaEnt
       
    Set oEva = Nothing
    Exit Function
AgregaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccion:AgregaProSelec Method")
End Function
