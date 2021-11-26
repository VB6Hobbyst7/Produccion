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

Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Public sConexion As String
Public bError As Boolean
Public sMsgError As String
Private sCentralCom As String

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

'##ModelId=3AC2BBDD023F
'Public Function GetConstante(psGrupo As ConstanteCabecera, Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True) As Recordset
 Public Function GetConstante(ByVal psGrupo As ConstanteCabecera, Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True, Optional pbEstructuraArbol As Boolean = False, Optional psSubArbol As String = "", Optional psValorExcluido As String = "", Optional pbUnido As Boolean = False) As Recordset
    Dim sqlC As String
    Dim oCon As DConecta
    Dim lsOrdena As String
    Set oCon = New DConecta
    
    On Error GoTo GetConstanteErr
    
    lsOrdena = " Order by " & IIf(pbOrdenCol1, " Columna1 ", " Columna2 ") & IIf(pbAscendente, " ASC ", " DESC ")
    If Not pbEstructuraArbol Then
        If psValorExcluido = "" Then
            If pbUnido Then
                sqlC = "Select cConsDescripcion + space(50) + Str(nConsValor) As Columna1  From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' " & lsOrdena
            Else
                sqlC = "Select cConsDescripcion as Columna1, nConsValor as Columna2 From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "'  And nConsValor <> nConsCod " & lsOrdena
            End If
        Else
            If pbUnido Then
                sqlC = "Select cConsDescripcion + space(50) + Str(nConsValor) As Columna1 From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' And nConsValor not in ('" & psValorExcluido & "')  And nConsValor <> nConsCod " & lsOrdena
            Else
                sqlC = "Select cConsDescripcion as Columna1, nConsValor as Columna2 From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' And nConsValor Not in ('" & psValorExcluido & "')  And nConsValor <> nConsCod " & lsOrdena
            End If
        End If
    Else
        If psSubArbol = "" Then
            If psValorExcluido = "" Then
                sqlC = "Select nConsValor as Columna2, cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' And nConsValor <> nConsCod " & lsOrdena
            Else
                sqlC = "Select nConsValor as Columna2, cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' And nConsValor not in ('" & psValorExcluido & "')  And nConsValor <> nConsCod " & lsOrdena
            End If
        Else
            If psValorExcluido = "" Then
                sqlC = "Select nConsValor as Columna2, cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' And nConsValor like '" & psSubArbol & "%'  And nConsValor <> nConsCod " & lsOrdena
            Else
                sqlC = "Select nConsValor as Columna2, cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' And nConsValor like '" & psSubArbol & "%' And nConsValor not in ('" & psValorExcluido & "') And nConsValor <> nConsCod " & lsOrdena
            End If
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

Public Function GetPeriodosNoLabTpo(Optional psSubArbol As String = "", Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True, Optional pbSoloAprovacionRechazo As Boolean = False) As Recordset
    Dim sqlC As String
    Dim oCon As DConecta
    Dim lsOrdena As String
    Set oCon = New DConecta
    
    On Error GoTo GetConstanteErr
    
    lsOrdena = " Order by " & IIf(pbOrdenCol1, " Columna1 ", " Columna2 ") & IIf(pbAscendente, " ASC ", " DESC ")
    If pbSoloAprovacionRechazo Then
        sqlC = " Select nConsValor as Columna2, ltrim(rtrim(str(nRHPeriodoUnidad))) + '-' + cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante CO" _
             & " Inner Join RHPeriodosNolaboradosTpo RHT ON CO.nConsValor = RHT.nRHPeriodoTpo where nConsCod like '" & gRHEstado & "' And nConsValor like '" & psSubArbol & "%' And nConsValor > 0 " & lsOrdena
    Else
        sqlC = " Select nConsValor as Columna2, ltrim(rtrim(str(nRHPeriodoUnidad))) + '-' + cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante CO" _
             & " Inner Join RHPeriodosNolaboradosTpo RHT ON CO.nConsValor = RHT.nRHPeriodoTpo where nConsCod like '" & gRHEstado & "' And nConsValor like '" & psSubArbol & "%'" & lsOrdena
    End If
    
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
Public Function GetAgencias(Optional pbOrdenCol1 As Boolean = False, Optional pbAscendente As Boolean = True, Optional pbTipoArbol As Boolean = False, Optional psArea As String = "", Optional psPersCod As String = "") As Recordset
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
        ElseIf psPersCod <> "" Then
            sqlC = " Select Cargo.cAreaCod + AA.cUbicaCod as Columna2, cAgeDescripcion as Columna1, Len(AG.cAgeCod) Nivel From " & lsServerComunes & "Agencias AG" _
                 & " Inner Join AreaAgencia AA On AG.cAgeCod = AA.cUbicaCod" _
                 & " Inner Join (Select cAreaCod from RRHH where cPersCod = '" & psPersCod & "') Cargo On Cargo.cAreaCod = AA.cAreaCod  " _
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
         & "    When 6 then 'D'" _
         & "    end Valor, cConsDescripcion Descrip , 2 Nivel from constante" _
         & " where nconscod like '" & gRHTipoContrato & "' and nConsvalor > 1" _
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

Public Function RecuperaConstantes(ByVal psConstanteCab As ConstanteCabecera) As ADODB.Recordset
Dim Conn As DConecta
    On Error GoTo ErrorRecuperaConstantes
    Set Conn = New DConecta
    Conn.AbreConexion
    Set RecuperaConstantes = Conn.CargaRecordSet("Select cConsDescripcion,nConsValor from " & sCentralCom & "Constante where nConsValor <> '" & psConstanteCab & "' And  nConsCod = '" & psConstanteCab & "' order by nConsValor")
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

ErrorRecuperaConstantes:
    Err.Raise Err.Number, "Error En Proceso RecuperaConstantes", Err.Description
    
End Function
Public Function DameDescripcionConstante(ByVal pnCodigo As ConstanteCabecera, ByVal pnValor As Long) As String
Dim sSQL As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorDameDescripcionConstante
    sSQL = "Select cConsDescripcion from Constante where nConsCod = " & pnCodigo & " AND nConsValor = " & pnValor
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.BOF And Not R.EOF Then
        DameDescripcionConstante = Trim(R!cConsDescripcion)
    Else
        DameDescripcionConstante = ""
    End If
    R.Close
    Set R = Nothing
    Exit Function

ErrorDameDescripcionConstante:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function CargaConstante(Optional psConsCod As ConstanteCabecera = 0, Optional psConsValor As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
Dim sCond As String
   On Error GoTo CargaConstanteErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion Then
   If psConsCod <> 0 Then
      sCond = " WHERE nCONSCOD = " & psConsCod & " and nConsValor <> " & psConsCod
   End If
   If psConsValor <> "" Then
      sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " nConsValor = " & psConsValor
   End If
      psSql = "SELECT nConsCod, nConsValor, cConsDescripcion FROM Constante " & sCond
      Set CargaConstante = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaConstanteErr:
   Call RaiseError(MyUnhandledError, "DConstantes:CargaConstante Method")
End Function
