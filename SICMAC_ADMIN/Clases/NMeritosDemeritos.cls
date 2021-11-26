VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NMeritosDemeritos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Realiza el mantenimiento de los MerDems de los empleados y la asignación
'al examen a al que postula
Option Base 0
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902CE00F5
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902CE018B
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

'##ModelId=3AB902CE01A9
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'Agrega un nuevo Item de MerDem Tabla
'##ModelId=3B02D75603C3
Public Function AgregaMerDemTabla(psRHMerDemTablaCod As String, psRHpcRHMerDemTablaCodDescripcion As String, psMeritoTpo As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaMerDemTablaErr
    Dim oMer As DMeritosDemeritos
    Set oMer = New DMeritosDemeritos
    
    oMer.AgregaMerDemTabla psRHMerDemTablaCod, psRHpcRHMerDemTablaCodDescripcion, psMeritoTpo, psUltimaActualizacion
    
    Set oMer = Nothing
    
    Exit Function
AgregaMerDemTablaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosMerDem:AgregaMerDemTabla Method")
End Function

'Elimina  un Item de MerDem Tabla
'##ModelId=3B02D7570035
Public Function EliminaMerDemTabla(psRHMerDemTablaCod As String) As Boolean
    On Error GoTo EliminaMerDemTablaErr
    Dim oMer As DMeritosDemeritos
    Set oMer = New DMeritosDemeritos
    
    oMer.EliminaMerDemTabla psRHMerDemTablaCod

    Exit Function
EliminaMerDemTablaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosMerDem:EliminaMerDemTabla Method")
End Function

'Elimina  un Item de MerDem Tabla
'##ModelId=3B02D7570035
Public Function TipoUsado(psRHMerDemTablaCod As String) As Boolean
    On Error GoTo EliminaMerDemTablaErr
    Dim oMer As DMeritosDemeritos
    Set oMer = New DMeritosDemeritos
    
    TipoUsado = oMer.TipoUsado(psRHMerDemTablaCod)

    Exit Function
EliminaMerDemTablaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosMerDem:EliminaMerDemTabla Method")
End Function

'Actualiza un Item de MerDem Tabla
'##ModelId=3B02D757008F
Public Function ModificaMerDemTabla(psRHMerDemTablaCod As String, psRHpcRHMerDemTablaCodDescripcion As String, psMeritoTpo As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaMerDemTablaErr
    Dim oMer As DMeritosDemeritos
    Set oMer = New DMeritosDemeritos
    
    oMer.ModificaMerDemTabla psRHMerDemTablaCod, psRHpcRHMerDemTablaCodDescripcion, psMeritoTpo, psUltimaActualizacion

    Exit Function
ModificaMerDemTablaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosMerDem:ModificaMerDemTabla Method")
End Function

'Agrega un nuevo Item de MerDem
'##ModelId=3B02D75700F4
Public Function AgregaMerDem(psRHPersCod As String, psRHMerDemCod As String, pnRHItem As Integer, psRHCurLugar As String, psRHCurInicio As String, psRHCurFin As String, pnRHCurUndTiempo As Integer, pnRHCurTiempo As Integer, psComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaMerDemErr
    Dim oMer As DMeritosDemeritos
    Set oMer = New DMeritosDemeritos
    
    oMer.AgregaMerDem psRHPersCod, psRHMerDemCod, pnRHItem, psRHCurLugar, psRHCurInicio, psRHCurFin, pnRHCurUndTiempo, pnRHCurTiempo, psComentario, psUltimaActualizacion
    
    Set oMer = Nothing
    Exit Function
AgregaMerDemErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosMerDem:AgregaMerDem Method")
End Function

'Elimina  un Item de MerDem
'##ModelId=3B02D757014E
Public Function EliminaMerDem(psRHMerDemCod As String, psRHPersCod As String, pnRHItem As Integer) As Boolean
    On Error GoTo EliminaMerDemErr
    Dim oMer As DMeritosDemeritos
    Set oMer = New DMeritosDemeritos
    
    oMer.EliminaMerDem psRHPersCod, psRHMerDemCod, pnRHItem
    
    Set oMer = Nothing
    Exit Function
EliminaMerDemErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosMerDem:EliminaMerDem Method")
End Function

'Actualiza un Item de MerDem
'##ModelId=3B02D75701A8
Public Function ModificaMerDem(psPersCod As String, prRH As ADODB.Recordset, psUltimaActualizacion As String, pgsFormatoFecha As String) As Boolean
    On Error GoTo ModificaMerDemErr
    Dim oMer As DMeritosDemeritos
    Set oMer = New DMeritosDemeritos
    
    oMer.ModificaMerDem psPersCod, prRH, psUltimaActualizacion, pgsFormatoFecha
    
    Set oMer = Nothing
    Exit Function
ModificaMerDemErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosMerDem:ModificaMerDem Method")
End Function

'psPersCod As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date
'Actualiza un Item de MerDem
'##ModelId=3B02D75701A8
Public Function GetReporteMerDemTabla(pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo GetReporteMerDemTablaErr
    Dim lsCod As String * 10
    Dim lsDescripcion As String * 50
    Dim lsTipo As String * 20
    Dim lsUltiMov As String * 27
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim oMer As DMeritosDemeritos
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oMer = New DMeritosDemeritos
    
    Set rs = oMer.GetMerDemTabla
    
    lsCadena = ""
    lnPagina = 0
    lnItem = 0
    
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("MerDem Tabla", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Codigo;8;Descripcion;20; ;30;Tipo;10; ;10;Ult. Movimiento;20; ;10;", lnItem)
        While Not rs.EOF
            lsCod = rs.Fields(1)
            lsDescripcion = rs.Fields(2)
            lsTipo = rs.Fields(3)
            lsUltiMov = rs.Fields(4)
            
            lsCadena = lsCadena & "   " & lsCod & "   " & lsDescripcion & lsTipo & lsUltiMov & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("MerDem Tabla", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Codigo;8;Descripcion;20; ;30;Tipo;10; ;10;Ult. Movimiento;20; ;10;", lnItem)
            End If
            
            rs.MoveNext
        Wend
    End If
    
    GetReporteMerDemTabla = lsCadena
    Set oMer = Nothing
    Exit Function
GetReporteMerDemTablaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosMerDem:ModificaMerDem Method")
End Function

'Actualiza un Item de MerDem
'##ModelId=3B02D75701A8
Public Function GetReporteMerDem(psPersCod As String, psNombre As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo GetReporteMerDemErr
    Dim oPersona As UPersona
    Set oPersona = New UPersona
    Dim lsFecha As String * 12
    Dim lsTipo As String * 4
    Dim lsTipoDef As String * 48
    Dim lsDescripcion As String * 50
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsCadena As String
    Dim oMer As DMeritosDemeritos
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oMer = New DMeritosDemeritos
    
    Set rs = oMer.GetMerDems(psPersCod)
    lsCadena = ""
    lnPagina = 0
    lnItem = 0
    
    lsCadena = lsCadena & CabeceraPagina("MERITOS DEMERITOS - " & PstaNombre(psNombre), lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
    lsCadena = lsCadena & Encabezado("Tipo;10; ;29;Fecha;20;Observacion;20; ;39;", lnItem)
    
    If Not (rs.EOF And rs.BOF) Then
        While Not rs.EOF
            lsTipo = rs.Fields(0)
            lsTipoDef = rs.Fields(1)
            lsFecha = rs.Fields(2)
            lsDescripcion = rs.Fields(3)
            lsCadena = lsCadena & lsTipo & lsTipoDef & "  " & lsFecha & "  " & lsDescripcion & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            
            If lnItem > 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("MERITOS DEMERITOS - " & PstaNombre(psNombre), lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Tipo;10; ;29;Fecha;20;Observacion;20; ;39;", lnItem)
            End If
            
            rs.MoveNext
        Wend
    End If
    
    GetReporteMerDem = lsCadena
    Set oMer = Nothing
    Exit Function
GetReporteMerDemErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosMerDem:ModificaMerDem Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902CE0113
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

