VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DConstantes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3AC2BBBF0051"
'
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
    '##ModelId=3AC2BC280142
    Private mlClassDebugID As Long
#End If

Public Sub Inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub

'##ModelId=3AC2BC2803E1
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AC2BC28035F
Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'##ModelId=3AC2BBDD023F
'Public Function GetConstante(psGrupo As ConstanteCabecera, Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True) As Recordset
 Public Function GetConstante(ByVal psGrupo As ConstanteCabecera, Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True, Optional pbEstructuraArbol As Boolean = False, Optional psSubArbol As String = "") As Recordset
    Dim sqlC As String
    Dim oCon As DConecta
    Dim lsOrdena As String
    Set oCon = New DConecta
    
    On Error GoTo GetConstanteErr
    
    lsOrdena = " Order by " & IIf(pbOrdenCol1, " Columna1 ", " Columna2 ") & IIf(pbAscendente, " ASC ", " DESC ")
    If Not pbEstructuraArbol Then
        sqlC = "Select cConsDescripcion as Columna1, nConsValor as Columna2 From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' " & lsOrdena
    Else
        If psSubArbol = "" Then
            sqlC = "Select nConsValor as Columna2, cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' " & lsOrdena
        Else
            sqlC = "Select nConsValor as Columna2, cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' And nConsValor like '" & psSubArbol & "%'" & lsOrdena
        End If
    End If
    
    If oCon.AbreConexion Then
        Set GetConstante = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    
    Exit Function
GetConstanteErr:
    Call RaiseError(MyUnhandledError, "DConstantes:GetConstante Method")
End Function

Public Function GetPeriodosNoLabTpo(Optional psSubArbol As String = "", Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True) As Recordset
    Dim sqlC As String
    Dim oCon As DConecta
    Dim lsOrdena As String
    Set oCon = New DConecta
    
    On Error GoTo GetConstanteErr
    
    lsOrdena = " Order by " & IIf(pbOrdenCol1, " Columna1 ", " Columna2 ") & IIf(pbAscendente, " ASC ", " DESC ")
    sqlC = " Select nConsValor as Columna2, ltrim(rtrim(str(nRHPeriodoUnidad))) + '-' + cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante CO" _
         & " Inner Join RHPeriodosNolaboradosTpo RHT ON CO.nConsValor = RHT.nRHPeriodoTpo where nConsCod like '6006' And nConsValor like '" & psSubArbol & "%'" & lsOrdena
    
    If oCon.AbreConexion Then
        Set GetPeriodosNoLabTpo = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    
    Exit Function
GetConstanteErr:
    Call RaiseError(MyUnhandledError, "DConstantes:GetConstante Method")
End Function


'##ModelId=3AC2BBDD023F
Public Function GetAgencias(Optional pbOrdenCol1 As Boolean = False, Optional pbAscendente As Boolean = True, Optional pbTipoArbol As Boolean = False, Optional psArea As String = "", Optional psPerscod As String = "") As Recordset
    Dim sqlC As String
    Dim oCon As DConecta
    Dim lsOrdena As String
    Set oCon = New DConecta
    
    On Error GoTo GetAgenciasErr
    
    lsOrdena = " Order by " & IIf(pbOrdenCol1, " Columna1 ", " Columna2 ") & IIf(pbAscendente, " ASC ", " DESC ")
    If Not pbTipoArbol Then
        sqlC = "Select cAgeDescripcion as Columna1, cAgeCod as Columna2 From " & lsServerComunes & "Agencias" & lsOrdena
    Else
        If psArea <> "" Then
            sqlC = " Select AA.cUbicaCod as Columna2, cAgeDescripcion as Columna1, Len(AG.cAgeCod) Nivel From " & lsServerComunes & "Agencias AG" _
                 & " Inner Join AreaAgencia AA On AG.cAgeCod = AA.cUbicaCod And AA.cAreaCod  = '" & psArea & "' " & lsOrdena
        ElseIf psPerscod <> "" Then
            sqlC = " Select Cargo.cAreaCod + AA.cUbicaCod as Columna2, cAgeDescripcion as Columna1, Len(AG.cAgeCod) Nivel From " & lsServerComunes & "Agencias AG" _
                 & " Inner Join AreaAgencia AA On AG.cAgeCod = AA.cUbicaCod" _
                 & " Inner Join (Select cAreaCod from RRHH where cPersCod = '" & psPerscod & "') Cargo On Cargo.cAreaCod = AA.cAreaCod  " _
                 & " " & lsOrdena
        Else
            sqlC = "Select cAgeCod as Columna2, cAgeDescripcion as Columna1, Len(cAgeCod) Nivel  From " & lsServerComunes & "Agencias" & lsOrdena
        End If
    End If
    If oCon.AbreConexion Then
        Set GetAgencias = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetAgenciasErr:
    Call RaiseError(MyUnhandledError, "DConstantes:GetConstante Method")
End Function

'##ModelId=3AC2BBDD023F
Public Function GetAreas(Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True) As Recordset
    Dim sqlC As String
    Dim oCon As DConecta
    Dim lsOrdena As String
    Set oCon = New DConecta
    
    On Error GoTo GetAreasErr
    
    lsOrdena = " Order by " & IIf(pbOrdenCol1, " Columna1 ", " Columna2 ") & IIf(pbAscendente, " ASC ", " DESC ")
    sqlC = "Select cAreaDescripcion as Columna1, cAreaCod as Columna2 From " & lsServerComunes & "Areas " & lsOrdena
    
    If oCon.AbreConexion Then
        Set GetAreas = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetAreasErr:
    Call RaiseError(MyUnhandledError, "DConstantes:GetConstante Method")
End Function

'##ModelId=3AC2BBDD023F
Public Function GetRHTipoContrato() As Recordset
    Dim sqlC As String
    Dim oCon As DConecta
    Dim lsOrdena As String
    Set oCon = New DConecta
    
    On Error GoTo GetAreasErr
    
    sqlC = " Select  Case nConsValor When 2 then 'L' " _
         & "    When 3 then 'F'" _
         & "    When 4 then 'P'" _
         & "    When 5 then 'S'" _
         & "    end Valor, cConsDescripcion Descrip , 2 Nivel from constante" _
         & " where nconscod like '6012%' and nConsvalor > 1" _
         & " Union" _
         & " Select 'E' Valor, 'LABORAL' Descrip , 2 Nivel Order By Valor"
    
    If oCon.AbreConexion Then
        Set GetRHTipoContrato = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetAreasErr:
    Call RaiseError(MyUnhandledError, "DConstantes:GetConstante Method")
End Function

#If DebugMode Then
    '##ModelId=3AC2BC2801F6
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
