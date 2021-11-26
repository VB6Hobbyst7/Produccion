VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPeriodoNoLaborado"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Public Function AgredaDatos(psPersCod As String, psRHAutorizacionTpo As String, psSolicitadoInicio As String, psSolicitadoFin As String, psProgramadoInicio As String, psProgramadoFin As String, psEjecutadoInicio As String, psEjecutadoFin As String, psComentario As String, psRHAutorizacionEstado As String, psUltimaActualizacion As String) As Boolean
    Dim oAut As DAutorizacionFisica
    Set oAut = New DAutorizacionFisica
    
    AgredaDatos = oAut.AgredaDatos(psPersCod, psRHAutorizacionTpo, psSolicitadoInicio, psSolicitadoFin, psProgramadoInicio, psProgramadoFin, psEjecutadoInicio, psEjecutadoFin, psComentario, psRHAutorizacionEstado, psUltimaActualizacion)
    
    Set oAut = Nothing
    Exit Function
ERROR:
    MsgBox Str(Err.Number) & Err.Description, vbInformation, "Aviso"
End Function

Public Function AgredaPerNoLab(psPersCod As String, psPeriodoTpo As String, psSolicitadoInicio As String, psSolicitadoFin As String, psEjecutadoInicio As String, psEjecutadoFin As String, psComentario As String, psRHPeriodoEstado As String, psUltimaActualizacion As String) As Boolean
    Dim oAut As DPeriodoNoLaborado
    Set oAut = New DPeriodoNoLaborado
    
    'AgredaPerNoLab = oAut.ModificaPeriodo(psPersCod, psPeriodoTpo, psUltimaActualizacion, gsFormatoFechaHora)
    
    Set oAut = Nothing
    Exit Function
ERROR:
    MsgBox Str(Err.Number) & Err.Description, vbInformation, "Aviso"
End Function

Public Function ModificaDatos(psPersCod As String, psRHAutorizacionTpo As String, psSolicitadoInicio As String, psSolicitadoFin As String, psProgramadoInicio As String, psProgramadoFin As String, psEjecutadoInicio As String, psEjecutadoFin As String, psComentario As String, psRHAutorizacionEstado As String, psUltimaActualizacion As String, psRHAutorizacionTpoT As String, psSolicitadoInicioT As String) As Boolean
    Dim oAut As DAutorizacionFisica
    Set oAut = New DAutorizacionFisica
    
    ModificaDatos = oAut.ModificaDatos(psPersCod, psRHAutorizacionTpo, psSolicitadoInicio, psSolicitadoFin, psProgramadoInicio, psProgramadoFin, psEjecutadoInicio, psEjecutadoFin, psComentario, psRHAutorizacionEstado, psUltimaActualizacion, psRHAutorizacionTpoT, psSolicitadoInicioT)
    
    Set oAut = Nothing
    Exit Function
ERROR:
    MsgBox Str(Err.Number) & Err.Description, vbInformation, "Aviso"
End Function

Public Function ModificaPerNoLab(psPersCod As String, psGrupo As String, prPeriodo As ADODB.Recordset, psUltimaActualizacion As String, psFormatoFechaHora As String) As Boolean
    Dim oAut As DPeriodoNoLaborado
    Set oAut = New DPeriodoNoLaborado
    
    ModificaPerNoLab = oAut.ModificaPeriodo(psPersCod, psGrupo, prPeriodo, psUltimaActualizacion, psFormatoFechaHora)
    
    Set oAut = Nothing
    Exit Function
ERROR:
    MsgBox Str(Err.Number) & Err.Description, vbInformation, "Aviso"
End Function

Public Function GetReporte(psPersCod As String, psNombre As String, pnRHAutorizacionGrupo As Integer, psMensaje As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    Dim rsA As ADODB.Recordset
    Dim lsCadena As String
    Set rsA = New ADODB.Recordset
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsTipo As String * 73
    Dim lsFecSolIni As String * 21
    Dim lsFecSolFin As String * 21
    Dim lsFecProgIni As String * 21
    Dim lsFecProgFin As String * 21
    Dim lsFecEjecIni As String * 21
    Dim lsFecEjecFin As String * 21
    Dim lsComentario As String * 73
    Dim lsEstado As String * 73
    Dim lsObserva As String * 73
    
    Dim oPNL As DPeriodoNoLaborado
    Set oPNL = New DPeriodoNoLaborado
    Set rsA = oPNL.GetRHPeriodoNoLabPersonaReporte(psPersCod, CInt(pnRHAutorizacionGrupo))
    
    lsCadena = ""
    
    If Not (rsA.BOF And rsA.EOF) Then
        lsCadena = lsCadena & CabeceraPagina(rsA.Fields(12) & " - " & PstaNombre(Trim(Left(psNombre, 50))), lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Tipo/Comentario/Estado/Observacion;40; ;17;Sol_Ini/Ejec_Ini;33;Sol_Fin/Ejec_Fin;27; ;3;", lnItem)
        While Not rsA.EOF
            lsTipo = Left(rsA.Fields(1), 30)
            lsFecSolIni = rsA.Fields(2)
            lsFecSolFin = rsA.Fields(3)
            lsFecProgIni = rsA.Fields(4)
            lsFecProgFin = rsA.Fields(5)
            lsFecEjecIni = ""
            lsFecEjecFin = ""
            lsComentario = rsA!Comentario
            lsEstado = rsA!de
            
            If IsNull(rsA!Obs) Then
                lsObserva = ""
            Else
                lsObserva = rsA!Obs
            End If
            
            lsCadena = lsCadena & lsTipo & lsFecSolIni & Space(6) & lsFecSolFin & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & lsComentario & lsFecProgIni & Space(6) & lsFecProgFin & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & lsEstado & lsFecEjecIni & Space(6) & lsFecEjecFin & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & lsObserva & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & String(120, "-") & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 4
        
            If lnItem >= 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina(rsA.Fields(12) & " - " & PstaNombre(Trim(Left(psNombre, 50))), lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Tipo/Comentario/Estado/Observacion;40; ;23;Sol_Ini/Ejec_Ini;30;Sol_Fin/Ejec_Fin;27;", lnItem)
            End If
            rsA.MoveNext
        Wend
        
    End If
    
    GetReporte = lsCadena
    rsA.Close
    Set rsA = Nothing
    Set oPNL = Nothing
    Exit Function
ERROR:
    MsgBox Str(Err.Number) & Err.Description, vbInformation, "Aviso"
End Function

Public Function Elimina(psPersCod As String, psRHAutorizacionTpo As String, psSolicitadoInicio As String) As Boolean
    Dim oAut As DAutorizacionFisica
    Set oAut = New DAutorizacionFisica
    
    oAut.EliminaDatos psPersCod, psRHAutorizacionTpo, psSolicitadoInicio
    
    Set oAut = Nothing
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
