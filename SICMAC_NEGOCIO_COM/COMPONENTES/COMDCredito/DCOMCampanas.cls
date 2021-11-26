VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCampanas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Function CargarCampanas() As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
    '**Modificado por DAOR 20071102
    'sSql = "Select * From Campanas Order By cDescripcion"
    sSql = "Select  IdCampana, cDescripcion,bEstado,isnull(dFechaIni,'') as dFechaIni ,isnull(dFechaFin,'') as dFechaFin "
    sSql = sSql & "From Campanas Where cProd <> 'A' or cProd is null Order By cDescripcion" 'ADD PARAMN BY ARLO 20171013
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set CargarCampanas = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
End Function

Public Function ObtenerMaxIdCampanas() As Integer
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Select isnull(Max(IdCampana),0) as IdCampanas"
    sSql = sSql & " From Campanas"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerMaxIdCampanas = rs!IdCampanas
    End If
    Set rs = Nothing
End Function
Public Function InsertarCampanas(ByVal pnIdCampana As Integer, ByVal psDescripcionCampana As String, _
ByVal pnEstado As Integer, ByVal pdFechaIni As String, ByVal pdFechaFin As String) As Boolean 'ADD PARAMN BY ARLO 20171013
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    On Error GoTo ErrHandler
    
    InsertarCampanas = True
    
    '**Modificado por DAOR 20071102
    'sSql = "Insert Into Campanas Values(" & pnIdCampana & ",'" & psDescripcionCampana & "'," & pnEstado & ")"
    sSql = "Insert Into Campanas(cDescripcion,bEstado,dfechaIni,dfechaFin)"
    sSql = sSql & "Values('" & psDescripcionCampana & "'," & pnEstado & ",'" & pdFechaIni & "','" & pdFechaFin & "')" 'ADD PARAMN BY ARLO 20171013
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.ConexionActiva.Execute sSql
    oConec.CierraConexion
    Set oConec = Nothing
    
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    InsertarCampanas = False
End Function


Public Function CargarAgencias() As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select * From Agencias"
    sSql = sSql & " Order By Cast(cAgeCod as int)"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set CargarAgencias = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function AsignacionAgenCamp() As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select A.cAgeDescripcion as cAgencia,C.cDescripcion as cCampanas,"
    sSql = sSql & " Cast(A.cAgeCod as int) as cAgeCod ,C.IdCampana"
    sSql = sSql & " From CampAgencia CA"
    sSql = sSql & " Inner Join Agencias A on A.cAgeCod=CA.cAgeCod"
    sSql = sSql & " Inner Join Campanas C on C.IdCampana=CA.IdCampanas"
    sSql = sSql & " Where C.bEstado = 1"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set AsignacionAgenCamp = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function VerificarRegistro(ByVal psCodAgencias As String, ByVal pnCodCampana As Integer) As Boolean
    ' Verifica que el agencia y la campaña exista o no en la tabla
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select Count(*) as nCantidad From CampAgencia"
    sSql = sSql & " Where IdCampanas=" & pnCodCampana & " And cAgeCod='" & psCodAgencias & "'"
    
    Set oConec = New COMConecta.DCOMConecta
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
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    InsertarAsignacion = True
    sSql = "Insert Into CampAgencia Values(" & pnCodCampana & ",'" & psCodAgencia & "','" & sFecha & "')"
    
    Set oConec = New COMConecta.DCOMConecta
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
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Select Count(*) as nCantidad"
    sSql = sSql & " From CampAgencia"
    sSql = sSql & " Where IdCampanas=" & pnIdCampana
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        If val(rs!nCantidad) > 0 Then
            VerificacionEliminacion = False
        Else
            VerificacionEliminacion = True
        End If
    End If
    Set rs = Nothing
End Function

Public Function EliminacionCampanas(ByVal pnIdCampana As Integer) As Boolean
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
    On Error GoTo ErrHandler
    
    sSql = "Delete From Campanas Where IdCampana=" & pnIdCampana
    Set oConec = New COMConecta.DCOMConecta
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
    Dim oConec As COMConecta.DCOMConecta
    On Error GoTo ErrHandler
        sSql = "Delete From CampAgencia Where IdCampanas=" & pnIdCampana & " And cAgeCod='" & pscAgeCod & "'"
        Set oConec = New COMConecta.DCOMConecta
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


Public Function ListaCampanasXAgencia(ByVal psAgenciaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select C.IdCampana,C.cDescripcion"
    sSql = sSql & " From CampAgencia CA"
    sSql = sSql & " Inner Join Campanas C on CA.IdCampanas=C.IdCampana"
    sSql = sSql & " Where CA.cAgeCod='" & psAgenciaCod & "' and C.bEstado=1"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCampanasXAgencia = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerIdCampanaxCuenta(ByVal psCtaCod As String) As Integer
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select C.IdCampana"
    sSql = sSql & " from ColocacCred CC"
    sSql = sSql & " Inner Join Campanas C on C.IdCampana=CC.IdCampana"
    sSql = sSql & " Where CC.cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
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

Public Function ListaCampanasAgenciacXCampanas(ByVal pnIdCampana As Integer) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select A.cAgeDescripcion as cAgencia,C.cDescripcion as cCampanas,"
    sSql = sSql & " Cast(A.cAgeCod as int) as cAgeCod ,C.IdCampana"
    sSql = sSql & " From CampAgencia CA"
    sSql = sSql & " Inner Join Agencias A on A.cAgeCod=CA.cAgeCod"
    sSql = sSql & " Inner Join Campanas C on C.IdCampana=CA.IdCampanas"
    sSql = sSql & " Where C.bEstado = 1 and CA.IdCampanas=" & pnIdCampana
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCampanasAgenciacXCampanas = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCampanasAgenciacXAgencia(ByVal pscAgenciaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select A.cAgeDescripcion as cAgencia,C.cDescripcion as cCampanas,"
    sSql = sSql & " Cast(A.cAgeCod as int) as cAgeCod ,C.IdCampana"
    sSql = sSql & " From CampAgencia CA"
    sSql = sSql & " Inner Join Agencias A on A.cAgeCod=CA.cAgeCod"
    sSql = sSql & " Inner Join Campanas C on C.IdCampana=CA.IdCampanas"
    sSql = sSql & " Where C.bEstado = 1 and CA.cAgeCod='" & pscAgenciaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCampanasAgenciacXAgencia = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function ActualizarCampanas(ByVal pnIdCampanas As Integer, ByVal psDescripcion As String, _
ByVal pbEstado As Boolean, ByVal pdFechaIni As String, ByVal pdFechaFin As String) As Boolean 'ADD PARAMN BY ARLO 20171013
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    On Error GoTo ErrHandler
    sSql = "Update Campanas"
    sSql = sSql & " set cDescripcion='" & psDescripcion & "',"
    sSql = sSql & " bEstado=" & IIf(pbEstado = True, 1, 0) & ","
    sSql = sSql & " dFechaIni='" & pdFechaIni & "',"    'BY ARLO 20171013
    sSql = sSql & " dFechaFin='" & pdFechaFin & "' "    'BY ARLO 20171013"
    sSql = sSql & " Where IdCampana=" & pnIdCampanas
    
    Set oConec = New COMConecta.DCOMConecta
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


Public Function ListaCampanaXAgecCtaCod(ByVal psAgeCod As String, ByVal psCtaCod As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select C.cDescripcion"
    sSql = sSql & " From CampAgencia CA"
    sSql = sSql & " Inner Join Campanas C on C.IdCampana=CA.IdCampanas"
    sSql = sSql & " Inner Join ColocacCred CC on CC.IdCampana=C.IdCampana"
    sSql = sSql & " Where CA.cAgeCod='" & psAgeCod & "' and CC.cCtaCod='" & psCtaCod & "'"

    Set oConec = New COMConecta.DCOMConecta
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
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select cDescripcion"
    sSql = sSql & " From Campanas"
    sSql = sSql & " Where IdCampana=" & pnIdCampana
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        DesCampanaXIdCampana = rs!CDescripcion
    End If
    Set rs = Nothing
End Function

Public Function Cargar_Objetos_Controles(ByRef pnIdCampana As Integer, ByVal psCtaCod As String, ByRef prs As ADODB.Recordset, ByVal psAgenciaCod As String) As String
    
    If psCtaCod <> "" Then
        pnIdCampana = ObtenerIdCampanaxCuenta(psCtaCod)
    End If
    
    Set prs = ListaCampanasXAgencia(psAgenciaCod)
    
End Function

Public Function GeneraIdCampanas() As Integer
Dim nIdCampana As Integer
    
    nIdCampana = ObtenerMaxIdCampanas
    nIdCampana = nIdCampana + 1
    
    GeneraIdCampanas = nIdCampana
End Function

Public Function GrabarAsignacionCampanas(ByVal pdFecSis As Date, _
                                        ByVal pMatListaInsercion As Variant, _
                                        ByRef psMensaje As String) As Boolean

Dim nAsignacion As Integer
Dim bResultado As Boolean
Dim i As Integer
Dim dFecha As String
    
On Error GoTo ErrorGrabarAsignacionCampanas

    bResultado = True
    
    dFecha = Format(pdFecSis, "MM/dd/yyyy")
    nAsignacion = VerificarPreviewInsercion(pMatListaInsercion)
    If nAsignacion = -1 Then
        psMensaje = "Usted no ha asignado ninguna campaña a una agencia"
    Else
        For i = nAsignacion To UBound(pMatListaInsercion) - 1
            bResultado = InsertarAsignacion(pMatListaInsercion(i, 0), pMatListaInsercion(i, 1), dFecha)
            If bResultado = False Then
                psMensaje = "Existe error en la insercción"
            End If
        Next i
    End If
    
    GrabarAsignacionCampanas = bResultado
    
    Exit Function
    
ErrorGrabarAsignacionCampanas:
    Err.Raise Err.Number, "Grabar Asignacion Campanas", Err.Description
End Function

Function VerificarPreviewInsercion(ByVal pMatListaInsercion As Variant) As Integer
 
 Dim i As Integer
 Dim nValor As Integer
 
 nValor = -1
 
 For i = 0 To UBound(pMatListaInsercion) - 1
    If VerificarRegistro(pMatListaInsercion(i, 0), pMatListaInsercion(i, 1)) = False Then
        'significa que el registro de aca empieza
        nValor = i
        Exit For
    End If
 Next i
 
 VerificarPreviewInsercion = nValor

End Function

