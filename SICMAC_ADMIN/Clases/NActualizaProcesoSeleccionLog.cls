VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaProcesoSelecLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

#If DebugMode Then
    '##ModelId=3AB902CE0299
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Public Function AgregaProSeleccionLog(psPeriodo As Integer, psTipoBien As Integer, psTipoSeleccion As Integer, psMoneda As Integer, psRangoIni As Long, psRangoFin As Long, psFechaConvo As String, psActualiza As String, psDescripcion As String, psNumeroCotizacion As String, pnClaseProceso As Integer) As Long
    On Error GoTo AgregaProSelecErr
    Dim clsDGAdqui As DLogAdquisi
    Dim lnCodigo As Long
    Set clsDGAdqui = New DLogAdquisi
    lnCodigo = clsDGAdqui.LogSeleccionCod(psPeriodo)
    clsDGAdqui.InsertaSeleccionLog lnCodigo, psPeriodo, psTipoBien, psTipoSeleccion, psMoneda, psRangoIni, psRangoFin, psFechaConvo, psActualiza, psDescripcion, psNumeroCotizacion, pnClaseProceso
    Set clsDGAdqui = Nothing
    AgregaProSeleccionLog = lnCodigo
    Exit Function
AgregaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaProSeleccionLog Method")

End Function
Public Function AgregaSeleccionConfigPuntaje(pnLogSelNro As Long, pnCodTipoEvaluacion As Integer, pnPuntajeMinimo As Integer, pnPuntajeMaximo As Integer, pnPesoPonderado As Double, psActualiza As String) As Long
    On Error GoTo AgregaProSelecErr
    Dim clsDGAdqui As DLogAdquisi
    Dim lnCodigo As Long
    Set clsDGAdqui = New DLogAdquisi
    AgregaSeleccionConfigPuntaje = 0
    clsDGAdqui.InsertaSeleccionConfPuntajes pnLogSelNro, pnCodTipoEvaluacion, pnPuntajeMinimo, pnPuntajeMaximo, pnPesoPonderado, psActualiza
    Set clsDGAdqui = Nothing
    AgregaSeleccionConfigPuntaje = 1
    Exit Function
AgregaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaSeleccionConfigPuntaje Method")

End Function
Public Function EliminaSeleccionConfigPuntaje(pnLogSelNro As Long) As Long
    On Error GoTo EliminaSeleccionConfigPuntajeErr
    Dim clsDGAdqui As DLogAdquisi
    Dim lnCodigo As Long
    Set clsDGAdqui = New DLogAdquisi
    EliminaSeleccionConfigPuntaje = 0
    clsDGAdqui.EliminaSeleccionConfPuntajes pnLogSelNro
    Set clsDGAdqui = Nothing
    EliminaSeleccionConfigPuntaje = 1
    Exit Function
EliminaSeleccionConfigPuntajeErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaSeleccionConfigPuntaje Method")

End Function
Public Function AgregaSeleccionComite(pnLogSelNro As Long, prLogProcSelecComite As ADODB.Recordset, psActualiza As String) As Long
    On Error GoTo AgregaProSelecErr
    Dim clsDGAdqui As DLogAdquisi
    Set clsDGAdqui = New DLogAdquisi
    AgregaSeleccionComite = 0
    clsDGAdqui.EliminaSeleccionComite pnLogSelNro
    Do While Not prLogProcSelecComite.EOF = True
        clsDGAdqui.InsertaSeleccionComite pnLogSelNro, prLogProcSelecComite.Fields(0), Right(prLogProcSelecComite.Fields(2), 1), psActualiza
        prLogProcSelecComite.MoveNext
    Loop
    Set clsDGAdqui = Nothing
    AgregaSeleccionComite = 1
    Exit Function
AgregaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaSeleccionComite Method")
End Function

Public Function AgregaSeleccionConfigBienes(pnLogSelNro As Long, prLogSelecionBienes As ADODB.Recordset, psActualiza As String) As Long
    On Error GoTo AgregaProSelecErr
    Dim clsDGAdqui As DLogAdquisi
    Set clsDGAdqui = New DLogAdquisi
    AgregaSeleccionConfigBienes = 0
    clsDGAdqui.EliminaSeleccionConfigBienes pnLogSelNro
    'clsDGAdqui.InsertaLogSelReferencia
    'pnLogSelNro As Long, pscBSCod As String, psDescripcion As String, pnCantidad As Integer, pnUnidad As Integer, pnPrecio As Double, psActualiza As String
    Do While Not prLogSelecionBienes.EOF = True
            clsDGAdqui.InsertaSeleccionConfigBienes pnLogSelNro, prLogSelecionBienes.Fields(0), prLogSelecionBienes.Fields(4), prLogSelecionBienes.Fields(5), prLogSelecionBienes.Fields(3), prLogSelecionBienes.Fields(6), psActualiza
            prLogSelecionBienes.MoveNext
    Loop
    Set clsDGAdqui = Nothing
    AgregaSeleccionConfigBienes = 1
    Exit Function
AgregaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaSeleccionConfigBienes Method")
End Function
Public Function AgregaSeleccionCotizacionProveedores(pnLogSelNro As Long, psCodProveedor As String, prLogSelecionCotProveedores As ADODB.Recordset, psActualiza As String) As Long
    On Error GoTo AgregaProSelecErr
    Dim clsDGAdqui As DLogAdquisi
    Set clsDGAdqui = New DLogAdquisi
    AgregaSeleccionCotizacionProveedores = 0
    clsDGAdqui.EliminaLogSeleccionCotDetalle pnLogSelNro, psCodProveedor
    Do While Not prLogSelecionCotProveedores.EOF = True
            'pnLogSelNro As Long, pscBSCod As String, psCodPer As String,pnCantidad As Integer,
            'pnUnidad As Integer, pnPrecio As Double, pnIndicadorParticipa As Integer, psActualiza
            
            clsDGAdqui.InsertaLogSeleccionCotDetalle pnLogSelNro, prLogSelecionCotProveedores.Fields(0), psCodProveedor, _
            prLogSelecionCotProveedores.Fields(5), prLogSelecionCotProveedores.Fields(3), IIf(prLogSelecionCotProveedores.Fields(6) = "0", 0, prLogSelecionCotProveedores.Fields(6)), _
            prLogSelecionCotProveedores.Fields(8), psActualiza
            prLogSelecionCotProveedores.MoveNext
    Loop
    Set clsDGAdqui = Nothing
    AgregaSeleccionCotizacionProveedores = 1
    Exit Function
AgregaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaSeleccionCotizacionProveedores Method")
End Function
Public Function AgregaSeleccionCriteriosTecnicos(pnLogSelNro As Long, prLogSelecionCriterios As ADODB.Recordset, psActualiza As String) As Long
    On Error GoTo AgregaSeleccionCriteriosTecnicosErr
    Dim clsDGAdqui As DLogAdquisi
    Set clsDGAdqui = New DLogAdquisi
    AgregaSeleccionCriteriosTecnicos = 0
    clsDGAdqui.EliminaSeleccionCriteriosProceso pnLogSelNro
    Do While Not prLogSelecionCriterios.EOF = True
            clsDGAdqui.InsertaSeleccionCriteriosProceso pnLogSelNro, prLogSelecionCriterios.Fields(0), prLogSelecionCriterios.Fields(2), psActualiza
            prLogSelecionCriterios.MoveNext
    Loop
    Set clsDGAdqui = Nothing
    AgregaSeleccionCriteriosTecnicos = 1
    Exit Function
AgregaSeleccionCriteriosTecnicosErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaSeleccionCriteriosTecnicos Method")
End Function

Public Function AgregaSeleccionEvaluacionTecnica(pnLogSelNro As Long, prLogSelecionCriterios As ADODB.Recordset, psProveedor As String, psActualiza As String) As Long
    On Error GoTo AgregaSeleccionEvaluacionTecnicaErr
    Dim clsDGAdqui As DLogAdquisi
    Set clsDGAdqui = New DLogAdquisi
    AgregaSeleccionEvaluacionTecnica = 0
    clsDGAdqui.EliminaSeleccionCalificacionTecnica pnLogSelNro, psProveedor
    'pnLogSelNro ,pnCodCriterioTec ,psCodProveedor ,pnPuntaje,psDescripcion,psactualiza
    Do While Not prLogSelecionCriterios.EOF = True
            clsDGAdqui.InsertaSeleccionCalificacionTecnica pnLogSelNro, prLogSelecionCriterios.Fields(0), psProveedor, prLogSelecionCriterios.Fields(3), prLogSelecionCriterios.Fields(4), psActualiza
            prLogSelecionCriterios.MoveNext
    Loop
    Set clsDGAdqui = Nothing
    AgregaSeleccionEvaluacionTecnica = 1
    Exit Function
AgregaSeleccionEvaluacionTecnicaErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaSeleccionEvaluacionTecnica Method")
End Function
Public Function AgregaSeleccionProveedores(pnLogSelNro As Long, prLogSelecionProveedor As ADODB.Recordset, psActualiza As String) As Long
    On Error GoTo AgregaSeleccionProveedoresErr
    Dim clsDGAdqui As DLogAdquisi
    Set clsDGAdqui = New DLogAdquisi
    AgregaSeleccionProveedores = 0
    clsDGAdqui.EliminaSeleccionProveedores pnLogSelNro
    Do While Not prLogSelecionProveedor.EOF = True
            clsDGAdqui.InsertaSeleccionProveedores pnLogSelNro, prLogSelecionProveedor.Fields(0), psActualiza
            prLogSelecionProveedor.MoveNext
    Loop
    Set clsDGAdqui = Nothing
    AgregaSeleccionProveedores = 1
    Exit Function
AgregaSeleccionProveedoresErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaSeleccionProveedores Method")
End Function
Public Function ActualizaSelPuntajeEvaluacionEco(pnLogSelNro As Long, psCodPers As String, psCodBien As String, pnPuntaje As Double, cfechaEvaluacionEco As String) As Integer
    On Error GoTo ActualizaSelPuntajeEvaluacionEcoErr
        Dim clsDGAdqui As DLogAdquisi
        Set clsDGAdqui = New DLogAdquisi
        ActualizaSelPuntajeEvaluacionEco = 0
        clsDGAdqui.ActualizaSelEvaluacionEco pnLogSelNro, psCodPers, psCodBien, pnPuntaje, cfechaEvaluacionEco
        Set clsDGAdqui = Nothing
        ActualizaSelPuntajeEvaluacionEco = 1
        Exit Function
ActualizaSelPuntajeEvaluacionEcoErr:
        Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog : ActualizaSelPuntajeEvaluacionEco Method")
End Function



Public Function AgregaSeleccionReferencia(pnLogSelNro As Long, pnLogSelPeriodo As Integer, pnLogSelTpoReq As Integer, _
                pnConsolidado As Integer, psDescripcionConsol As String, _
                pnmesIni As String, pnmesFin As String, pscategoriabien, psActualiza As String) As Long
    On Error GoTo AgregaProSelecErr
    Dim clsDGAdqui As DLogAdquisi
    Set clsDGAdqui = New DLogAdquisi
    AgregaSeleccionReferencia = 0
    clsDGAdqui.EliminaLogSelReferencia pnLogSelNro
    clsDGAdqui.InsertaLogSelReferencia pnLogSelNro, pnLogSelPeriodo, pnLogSelTpoReq, pnConsolidado, psDescripcionConsol, pnmesIni, pnmesFin, pscategoriabien, psActualiza
    Set clsDGAdqui = Nothing
    AgregaSeleccionReferencia = 1
    Exit Function
AgregaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaSeleccionConfigBienes Method")
End Function

Public Function AgregaSeleccionMantCriterios(pnCodCriterioTec As Integer, psDescripcion As String, pnPuntajeDefault As Double, psUltimaActualizacion As String) As Long
    On Error GoTo AgregaSeleccionMantCriteriosErr
    Dim clsDGAdqui As DLogAdquisi
    Set clsDGAdqui = New DLogAdquisi
    AgregaSeleccionMantCriterios = 0
    clsDGAdqui.InsertaLogSelMantCriterios pnCodCriterioTec, psDescripcion, pnPuntajeDefault, psUltimaActualizacion
    Set clsDGAdqui = Nothing
    AgregaSeleccionMantCriterios = 1
    Exit Function
AgregaSeleccionMantCriteriosErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaSeleccionMantCriterios Method")
End Function
Public Function EliminaSeleccionMantCriterios(pnCodCriterioTec As Integer) As Long
    On Error GoTo EliminaSeleccionMantCriteriosErr
    Dim clsDGAdqui As DLogAdquisi
    Set clsDGAdqui = New DLogAdquisi
    EliminaSeleccionMantCriterios = 0
    clsDGAdqui.EliminaLogSelMantCriterios pnCodCriterioTec
    Set clsDGAdqui = Nothing
    EliminaSeleccionMantCriterios = 1
    Exit Function
EliminaSeleccionMantCriteriosErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:EliminaSeleccionMantCriterios Method")
End Function
Public Function ActualizaSeleccionMantCriterios(pnCodCriterioTec As Integer, psDescripcion As String, pnPuntajeDefault As Double, psUltimaActualizacion As String) As Long
    On Error GoTo ActualizaSeleccionMantCriteriosErr
    Dim clsDGAdqui As DLogAdquisi
    Set clsDGAdqui = New DLogAdquisi
    
    ActualizaSeleccionMantCriterios = 0
    clsDGAdqui.ActualizaLogSelMantCriterios pnCodCriterioTec, psDescripcion, pnPuntajeDefault, psUltimaActualizacion
    Set clsDGAdqui = Nothing
    ActualizaSeleccionMantCriterios = 1
    Exit Function
ActualizaSeleccionMantCriteriosErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:ActualizaSeleccionMantCriterios Method")
End Function

Public Function ActualizaProSeleccionLog(pnLogSelNro As Long, psPeriodo As Integer, psTipoBien As Integer, psTipoSeleccion As Integer, psMoneda As Integer, psRangoIni As Long, psRangoFin As Long, psFechaConvo As String, psActualiza As String, psDescripcion As String, psNumeroCotizacion As String, pnClaseProceso As Integer) As Integer
    On Error GoTo ActualizaProSeleccionLogErr
    Dim clsDGAdqui As DLogAdquisi
    Set clsDGAdqui = New DLogAdquisi
    ActualizaProSeleccionLog = 1
    clsDGAdqui.EliminaLogSelReferencia pnLogSelNro
    clsDGAdqui.ActualizaSeleccionLog pnLogSelNro, psPeriodo, psTipoBien, psTipoSeleccion, psMoneda, psRangoIni, psRangoFin, psFechaConvo, psActualiza, psDescripcion, psNumeroCotizacion, pnClaseProceso
    Set clsDGAdqui = Nothing
    ActualizaProSeleccionLog = 0
    Exit Function
ActualizaProSeleccionLogErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaProSeleccionLog Method")
End Function

Public Function ActualizaSeleccionConfigPuntaje(pnLogSelNro As Long, pnCodTipoEvaluacion As Integer, pnPuntajeMinimo As Integer, pnPuntajeMaximo As Integer, pnPesoPonderado As Double, psActualiza As String) As Long
    On Error GoTo AgregaProSelecErr
    Dim clsDGAdqui As DLogAdquisi
    Dim lnCodigo As Long
    Set clsDGAdqui = New DLogAdquisi
    ActualizaSeleccionConfigPuntaje = 0
    clsDGAdqui.InsertaSeleccionConfPuntajes pnLogSelNro, pnCodTipoEvaluacion, pnPuntajeMinimo, pnPuntajeMaximo, pnPesoPonderado, psActualiza
    Set clsDGAdqui = Nothing
    ActualizaSeleccionConfigPuntaje = 1
    Exit Function
AgregaProSelecErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaSeleccionConfigPuntaje Method")

End Function

Public Function GetReporteLogSelCriteriosTecnicos(pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo GetReporteLogSelCriteriosTecnicosErr
    Dim lsCod As String * 10
    Dim lsDescripcion As String * 50
    Dim lsTipo As String * 20
    Dim lsUltiMov As String * 27
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim clsDGAdqui As DLogAdquisi
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set clsDGAdqui = New DLogAdquisi
    
    Set rs = clsDGAdqui.CargaSelCriteriosTecnicos(1)
    lsCadena = ""
    lnPagina = 0
    lnItem = 0
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("Criterios de Evaluacion", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Codigo;8;Descripcion;20; Ult. Movimiento;56; ;10;", lnItem)
        While Not rs.EOF
            lsCod = rs.Fields(0)
            lsDescripcion = rs.Fields(1)
            lsUltiMov = rs.Fields(2)
            lsCadena = lsCadena & "   " & lsCod & "   " & lsDescripcion & "   " & lsUltiMov & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Criterios de Evaluacion", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Codigo;30;Descripcion;30; ;30;Ult. Movimiento;25; ;10;", lnItem)
            End If
            rs.MoveNext
        Wend
    End If
    GetReporteLogSelCriteriosTecnicos = lsCadena
    Set clsDGAdqui = Nothing
    Exit Function
GetReporteLogSelCriteriosTecnicosErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosMerDem:ModificaMerDem Method")
End Function


Public Function AgregaSeleccionTipoComite(pnTipoComite As Long, prLogSelecioncomite As ADODB.Recordset, psActualiza As String) As Long
    On Error GoTo AgregaSeleccionProveedoresErr
    Dim clsDGAdqui As DLogAdquisi
    Set clsDGAdqui = New DLogAdquisi
    AgregaSeleccionTipoComite = 0
    clsDGAdqui.EliminaSeleccionTipoComite pnTipoComite
    Do While Not prLogSelecioncomite.EOF = True
            clsDGAdqui.InsertaSeleccionTipoComite pnTipoComite, prLogSelecioncomite.Fields(0), Right(prLogSelecioncomite.Fields(2), 1), psActualiza
            prLogSelecioncomite.MoveNext
    Loop
    Set clsDGAdqui = Nothing
    AgregaSeleccionTipoComite = 1
    Exit Function
AgregaSeleccionProveedoresErr:
    Call RaiseError(MyUnhandledError, "NActualizaProcesoSeleccionLog:AgregaSeleccionTipoComite Method")
End Function



