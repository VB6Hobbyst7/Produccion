VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCampanas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function CargarCampanas() As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    
    sSql = "Select * From Campanas"
    Set oConec = New DConecta
    oConec.AbreConexion
    Set CargarCampanas = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
End Function

Public Function ObtenerMaxIdCampanas() As Integer
    Dim sSql As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Select isnull(Max(IdCampana),0) as IdCampanas"
    sSql = sSql & " From Campanas"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerMaxIdCampanas = rs!IdCampanas
    End If
    Set rs = Nothing
End Function


Public Function InsertarCampanas(ByVal pnIdCampana As Integer, ByVal psDescripcionCampana As String, _
ByVal pnEstado As Integer) As Boolean
    Dim oConec As DConecta
    Dim sSql As String
    
    On Error GoTo ErrHandler
    
    InsertarCampanas = True
    
    sSql = "Insert Into Campanas Values(" & pnIdCampana & ",'" & psDescripcionCampana & "'," & pnEstado & ")"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    oConec.ConexionActiva.Execute sSql
    oConec.CierraConexion
    Set oConec = Nothing
    
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    InsertarCampanas = False
End Function


Public Function CargarAgencias() As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select * From Agencias"
    sSql = sSql & " Order By Cast(cAgeCod as int)"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set CargarAgencias = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function AsignacionAgenCamp() As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select A.cAgeDescripcion as cAgencia,C.cDescripcion as cCampanas,"
    sSql = sSql & " Cast(A.cAgeCod as int) as cAgeCod ,C.IdCampana"
    sSql = sSql & " From CampAgencia CA"
    sSql = sSql & " Inner Join Agencias A on A.cAgeCod=CA.cAgeCod"
    sSql = sSql & " Inner Join Campanas C on C.IdCampana=CA.IdCampanas"
    sSql = sSql & " Where C.bEstado = 1"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set AsignacionAgenCamp = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function VerificarRegistro(ByVal psCodAgencias As String, ByVal pnCodCampana As Integer) As Boolean
    ' Verifica que el agencia y la campaña exista o no en la tabla
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select Count(*) as nCantidad From CampAgencia"
    sSql = sSql & " Where IdCampanas=" & pnCodCampana & " And cAgeCod='" & psCodAgencias & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    
    If Not rs.EOF And Not rs.BOF Then
        If rs!nCantidad > 0 Then
            VerificarRegistro = True
        Else
            VerificarRegistro = False
        End If
    Else
        VerificarRegistro = False
    End If
    Set rs = Nothing
    ' si es true significa que la agencia y la campaña se encuentran registrador
End Function

Public Function InsertarAsignacion(ByVal psCodAgencia As String, ByVal pnCodCampana As Integer, ByVal sFecha As String) As Boolean
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    InsertarAsignacion = True
    sSql = "Insert Into CampAgencia Values(" & pnCodCampana & ",'" & psCodAgencia & "','" & sFecha & "')"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    oConec.ConexionActiva.Execute sSql
    oConec.CierraConexion
    Set oConec = Nothing
    
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    InsertarAsignacion = False
End Function

Public Function VerificacionEliminacion(ByVal pnIdCampana As Integer) As Boolean
 'Funcion que determina si se puede eliminar si o no?
 'si es False existe registros depedientes
 'si es True se puede eliminar
 
    Dim sSql As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Select Count(*) as nCantidad"
    sSql = sSql & " From CampAgencia"
    sSql = sSql & " Where IdCampanas=" & pnIdCampana
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        If Val(rs!nCantidad) > 0 Then
            VerificacionEliminacion = False
        Else
            VerificacionEliminacion = True
        End If
    End If
    Set rs = Nothing
End Function

Public Function EliminacionCampanas(ByVal pnIdCampana As Integer) As Boolean
    Dim sSql As String
    Dim oConec As DConecta
    
    On Error GoTo ErrHandler
    
    sSql = "Delete From Campanas Where IdCampana=" & pnIdCampana
    Set oConec = New DConecta
    oConec.AbreConexion
    oConec.ConexionActiva.Execute sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminacionCampanas = True
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    EliminacionCampanas = False
End Function

Public Function EliminacionCampanasAgencias(ByVal pnIdCampana As Integer, ByVal pscAgeCod As String) As Boolean
    Dim sSql As String
    Dim oConec As DConecta
    On Error GoTo ErrHandler
        sSql = "Delete From CampAgencia Where IdCampanas=" & pnIdCampana & " And cAgeCod='" & pscAgeCod & "'"
        Set oConec = New DConecta
        oConec.AbreConexion
        oConec.ConexionActiva.Execute sSql
        oConec.CierraConexion
        Set oConec = Nothing
        EliminacionCampanasAgencias = True
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    EliminacionCampanasAgencias = False
End Function


Public Function ListaCampanasXAgencia(ByVal psAgenciaCod As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select C.IdCampana,C.cDescripcion"
    sSql = sSql & " From CampAgencia CA"
    sSql = sSql & " Inner Join Campanas C on CA.IdCampanas=C.IdCampana"
    sSql = sSql & " Where CA.cAgeCod='" & psAgenciaCod & "' and C.bEstado=1"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCampanasXAgencia = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerIdCampanaxCuenta(ByVal psctacod As String) As Integer
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select C.IdCampana"
    sSql = sSql & " from ColocacCred CC"
    sSql = sSql & " Inner Join Campanas C on C.IdCampana=CC.IdCampana"
    sSql = sSql & " Where CC.cCtaCod='" & psctacod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs Is Nothing Then
        If Not rs.EOF And Not rs.BOF Then
            ObtenerIdCampanaxCuenta = rs!IdCampana
        End If
    Else
        ObtenerIdCampanaxCuenta = -1
    End If
    Set rs = Nothing
End Function

Public Function ListaCampanasAgenciacXCampanas(ByVal pnIdCampana As Integer) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select A.cAgeDescripcion as cAgencia,C.cDescripcion as cCampanas,"
    sSql = sSql & " Cast(A.cAgeCod as int) as cAgeCod ,C.IdCampana"
    sSql = sSql & " From CampAgencia CA"
    sSql = sSql & " Inner Join Agencias A on A.cAgeCod=CA.cAgeCod"
    sSql = sSql & " Inner Join Campanas C on C.IdCampana=CA.IdCampanas"
    sSql = sSql & " Where C.bEstado = 1 and CA.IdCampanas=" & pnIdCampana
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCampanasAgenciacXCampanas = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCampanasAgenciacXAgencia(ByVal pscAgenciaCod As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select A.cAgeDescripcion as cAgencia,C.cDescripcion as cCampanas,"
    sSql = sSql & " Cast(A.cAgeCod as int) as cAgeCod ,C.IdCampana"
    sSql = sSql & " From CampAgencia CA"
    sSql = sSql & " Inner Join Agencias A on A.cAgeCod=CA.cAgeCod"
    sSql = sSql & " Inner Join Campanas C on C.IdCampana=CA.IdCampanas"
    sSql = sSql & " Where C.bEstado = 1 and CA.cAgeCod='" & pscAgenciaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCampanasAgenciacXAgencia = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ActualizarCampanas(ByVal pnIdCampanas As Integer, ByVal psDescripcion As String, _
ByVal pbEstado As Boolean) As Boolean
    Dim oConec As DConecta
    Dim sSql As String
    
    On Error GoTo ErrHandler
    sSql = "Update Campanas"
    sSql = sSql & " set cDescripcion='" & psDescripcion & "',"
    sSql = sSql & " bEstado=" & IIf(pbEstado = True, 1, 0)
    sSql = sSql & " Where IdCampana=" & pnIdCampanas
    
    Set oConec = New DConecta
    oConec.AbreConexion
    oConec.ConexionActiva.Execute sSql
    oConec.CierraConexion
    Set oConec = Nothing
    ActualizarCampanas = True
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    ActualizarCampanas = False
End Function


Public Function ListaCampanaXAgecCtaCod(ByVal psAgecod As String, ByVal psctacod As String) As String
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select C.cDescripcion"
    sSql = sSql & " From CampAgencia CA"
    sSql = sSql & " Inner Join Campanas C on C.IdCampana=CA.IdCampanas"
    sSql = sSql & " Inner Join ColocacCred CC on CC.IdCampana=C.IdCampana"
    sSql = sSql & " Where CA.cAgeCod='" & psAgecod & "' and CC.cCtaCod='" & psctacod & "'"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CadenaConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        ListaCampanaXAgecCtaCod = IIf(IsNull(rs!CDescripcion), "", rs!CDescripcion)
    End If
    Set rs = Nothing
End Function


Public Function DesCampanaXIdCampana(ByVal pnIdCampana As Integer) As String
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select cDescripcion"
    sSql = sSql & " From Campanas"
    sSql = sSql & " Where IdCampana=" & pnIdCampana
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        DesCampanaXIdCampana = rs!CDescripcion
    End If
    Set rs = Nothing
End Function

