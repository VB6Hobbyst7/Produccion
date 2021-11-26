VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaDatosCurriculum"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A93FAA0025A"

'Realiza el mantenimiento de los curriculums de los empleados y la asignación
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

'Agrega un nuevo Item de Curriculum Tabla
'##ModelId=3B02D75603C3
Public Function AgregaCurriculumTabla(psRHCurriculumTablaCod As String, psRHpcRHCurriculumTablaCodDescripcion As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaCurriculumTablaErr
    Dim oCur As DActualizaDatosCurriculum
    Set oCur = New DActualizaDatosCurriculum
    
    oCur.AgregaCurriculumTabla psRHCurriculumTablaCod, psRHpcRHCurriculumTablaCodDescripcion, psUltimaActualizacion
    
    Set oCur = Nothing
    
    Exit Function
AgregaCurriculumTablaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCurriculum:AgregaCurriculumTabla Method")
End Function

'Elimina  un Item de Curriculum Tabla
'##ModelId=3B02D7570035
Public Function EliminaCurriculumTabla(psRHCurriculumTablaCod As String) As Boolean
    On Error GoTo EliminaCurriculumTablaErr
    Dim oCur As DActualizaDatosCurriculum
    Set oCur = New DActualizaDatosCurriculum
    
    oCur.EliminaCurriculumTabla psRHCurriculumTablaCod

    Exit Function
EliminaCurriculumTablaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCurriculum:EliminaCurriculumTabla Method")
End Function

'Es usado
'##ModelId=3B02D7570035
Public Function TipoUsado(psRHCurriculumTablaCod As String) As Boolean
    On Error GoTo EliminaCurriculumTablaErr
    Dim oCur As DActualizaDatosCurriculum
    Set oCur = New DActualizaDatosCurriculum
    
    TipoUsado = oCur.TipoUsado(psRHCurriculumTablaCod)
    
    Exit Function
EliminaCurriculumTablaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCurriculum:EliminaCurriculumTabla Method")
End Function

'Actualiza un Item de Curriculum Tabla
'##ModelId=3B02D757008F
Public Function ModificaCurriculumTabla(psRHCurriculumTablaCod As String, psRHpcRHCurriculumTablaCodDescripcion As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaCurriculumTablaErr
    Dim oCur As DActualizaDatosCurriculum
    Set oCur = New DActualizaDatosCurriculum
    
    oCur.ModificaCurriculumTabla psRHCurriculumTablaCod, psRHpcRHCurriculumTablaCodDescripcion, psUltimaActualizacion

    Exit Function
ModificaCurriculumTablaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCurriculum:ModificaCurriculumTabla Method")
End Function

'Agrega un nuevo Item de Curriculum
'##ModelId=3B02D75700F4
Public Function AgregaCurriculum(psRHPersCod As String, psRHCurriculumCod As String, pnRHItem As Integer, psRHCurLugar As String, psRHCurInicio As String, psRHCurFin As String, pnRHCurUndTiempo As Integer, pnRHCurTiempo As Integer, psComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaCurriculumErr
    Dim oCur As DActualizaDatosCurriculum
    Set oCur = New DActualizaDatosCurriculum
    
    oCur.AgregaCurriculum psRHPersCod, psRHCurriculumCod, pnRHItem, psRHCurLugar, psRHCurInicio, psRHCurFin, pnRHCurUndTiempo, pnRHCurTiempo, psComentario, psUltimaActualizacion
    
    Set oCur = Nothing
    Exit Function
AgregaCurriculumErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCurriculum:AgregaCurriculum Method")
End Function

'Elimina  un Item de Curriculum
'##ModelId=3B02D757014E
Public Function EliminaCurriculum(psRHCurriculumCod As String, psRHPersCod As String, pnRHItem As Integer) As Boolean
    On Error GoTo EliminaCurriculumErr
    Dim oCur As DActualizaDatosCurriculum
    Set oCur = New DActualizaDatosCurriculum
    
    oCur.EliminaCurriculum psRHPersCod, psRHCurriculumCod, pnRHItem
    
    Set oCur = Nothing
    Exit Function
EliminaCurriculumErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCurriculum:EliminaCurriculum Method")
End Function

'Actualiza un Item de Curriculum
'##ModelId=3B02D75701A8
Public Function ModificaCurriculum(psRHPersCod As String, prRS As ADODB.Recordset, prRSExtra As ADODB.Recordset, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaCurriculumErr
    Dim oCur As DActualizaDatosCurriculum
    Set oCur = New DActualizaDatosCurriculum
    
    oCur.ModificaCurriculum psRHPersCod, prRS, prRSExtra, psUltimaActualizacion
    
    Set oCur = Nothing
    Exit Function
ModificaCurriculumErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCurriculum:ModificaCurriculum Method")
End Function

'psPersCod As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date
'Actualiza un Item de Curriculum
'##ModelId=3B02D75701A8
Public Function GetReporteCurriculumTabla(pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo GetReporteCurriculumTablaErr
    Dim lsCod As String * 10
    Dim lsDescripcion As String * 80
    Dim lsUltiMov As String * 27
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim oCur As DActualizaDatosCurriculum
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCur = New DActualizaDatosCurriculum
    
    Set rs = oCur.GetCurriculumTabla
    
    lsCadena = ""
    lnPagina = 0
    lnItem = 0
    
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("Curriculum Tabla", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Codigo;8;Descripcion;20; ;60;Ult. Movimiento;25; ;5;", lnItem)
        While Not rs.EOF
            lsCod = rs.Fields(1)
            lsDescripcion = rs.Fields(2)
            lsUltiMov = rs.Fields(3)
            
            lsCadena = lsCadena & "   " & lsCod & "   " & lsDescripcion & lsUltiMov & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Curriculum Tabla", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Codigo;8;Descripcion;20; ;60;Ult. Movimiento;25; ;5;", lnItem)
            End If
            
            rs.MoveNext
        Wend
    End If
    
    GetReporteCurriculumTabla = lsCadena
    Set oCur = Nothing
    Exit Function
GetReporteCurriculumTablaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCurriculum:ModificaCurriculum Method")
End Function

'Actualiza un Item de Curriculum
'##ModelId=3B02D75701A8
Public Function GetReporteCurriculum(psPersCod As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo GetReporteCurriculumErr
    Dim oPersona As UPersona
    Set oPersona = New UPersona
    Dim lsItem As String * 5
    Dim lsTipo As String * 25
    Dim lsLugar As String * 25
    Dim lsDescripcion As String * 90
    Dim lsPeridoTipo As String * 15
    Dim lsPerido As String * 15
    Dim lsPeridoTipoAdd As String * 15
    Dim lsPeridoAdd As String * 15
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsCadena As String
    Dim oCur As DActualizaDatosCurriculum
    Dim rs As ADODB.Recordset
    Dim lbBit As Boolean
    Set rs = New ADODB.Recordset
    Set oCur = New DActualizaDatosCurriculum
    
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    oPersona.ObtieneClientexCodigo psPersCod
    
    Set rs = oCur.GetCurriculums(psPersCod)
    lsCadena = ""
    lnPagina = 0
    lnItem = 0
    
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & CentrarCadena("CURRICULUM VITAE", 80) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & CentrarCadena("================", 80) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "NOMBRE    : " & PstaNombre(oPersona.sPersNombre) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "DIRECCION : " & PstaNombre(oPersona.sPersDireccDomicilio) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "TELEFONO  : " & PstaNombre(oPersona.sPersTelefono) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "DIRECCION : " & PstaNombre(oPersona.sPersDireccDomicilio) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "DNI       : " & PstaNombre(oPersona.sPersIdnroDNI) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "RUC       : " & PstaNombre(oPersona.sPersIdnroRUC) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "FECHA NAC : " & Format(oRH.GetFecNac(psPersCod), gsFormatoFechaView) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & String(80, "-") & oImpresora.gPrnSaltoLinea
    
    lnItem = 10
    If Not (rs.EOF And rs.BOF) Then
        While Not rs.EOF
            lsItem = rs.Fields(13)
            lsLugar = rs.Fields(2)
            lsTipo = rs.Fields(1)
            lsDescripcion = rs.Fields(16)
            lbBit = False
            If Not IsNull(rs.Fields(8)) And Not IsNull(rs.Fields(5)) Then
                lbBit = True
                lsPeridoTipo = rs.Fields(8)
                lsPerido = rs.Fields(9)
                lsPeridoTipoAdd = rs.Fields(5)
                lsPeridoAdd = rs.Fields(6)
            ElseIf Not IsNull(rs.Fields(8)) Then
                lsPeridoTipo = rs.Fields(8)
                lsPerido = rs.Fields(9)
            Else
                lsPeridoTipo = rs.Fields(5)
                lsPerido = rs.Fields(6)
            End If
            
            lsCadena = lsCadena & "TIPO         : " & lsTipo & IIf(rs.Fields(15), " - OTORGADO POR LA CAJA", "") & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "LUGAR        : " & lsLugar & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "DESCRIPCION  : " & lsDescripcion & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "LUGAR        : " & lsLugar & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "PROFESION    : " & rs.Fields(11) & " - " & rs.Fields(14) & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "COSTO        : " & Format(rs.Fields(12), "#,##0.00") & " NUEVOS SOLES" & oImpresora.gPrnSaltoLinea
            If lbBit Then
                lsCadena = lsCadena & "PERIODO      : " & Trim(lsPerido) & " (" & Trim(lsPeridoTipo) & ")" & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & "DESDE        : " & lsPeridoTipoAdd & "  HASTA  : " & lsPeridoAdd & oImpresora.gPrnSaltoLinea
            Else
                If Not IsDate(lsPeridoTipo) Then
                    lsCadena = lsCadena & "PERIODO      : " & Trim(lsPerido) & " (" & Trim(lsPeridoTipo) & ")" & oImpresora.gPrnSaltoLinea
                Else
                    lsCadena = lsCadena & "DESDE        : " & lsPeridoTipo & "  HASTA  : " & lsPerido & oImpresora.gPrnSaltoLinea
                End If
            End If
            
            
            lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 8
            
            If lnItem > 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lnItem = 1
            End If
            
            rs.MoveNext
        Wend
    End If
    
    'Extra
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & CentrarCadena("ACTIVIDADES EXTRACURRICULARES", 80) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & CentrarCadena("=============================", 80) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    Set rs = oCur.GetCurriculumsExtra(psPersCod)
    lnItem = lnItem + 5
    If Not (rs.EOF And rs.BOF) Then
        While Not rs.EOF
            lsItem = rs.Fields(12)
            lsLugar = rs.Fields(3)
            lsTipo = rs.Fields(1)
            lsDescripcion = rs.Fields(9)
            
            lsCadena = lsCadena & "TIPO         : " & lsTipo & IIf(rs.Fields(8), " - OTORGADO POR LA CAJA", "") & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "ACTIVIDAD    : " & lsLugar & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "DESCRIPCION  : " & lsDescripcion & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "NIVEL        : " & rs.Fields(7) & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "AÑOS PRACTIC : " & rs.Fields(4) & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "COSTO        : " & Format(rs.Fields(5), "#,##0.00") & " NUEVOS SOLES" & oImpresora.gPrnSaltoLinea
            
            lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 7
            
            If lnItem > 52 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lnItem = 1
            End If
            
            rs.MoveNext
        Wend
    End If
    
    
    
    
    GetReporteCurriculum = lsCadena
    Set oCur = Nothing
    Exit Function
GetReporteCurriculumErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosCurriculum:ModificaCurriculum Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902CE0113
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

