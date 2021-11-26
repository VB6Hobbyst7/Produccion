VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************
'    DE ACA EN ADELANTE SE AGREGO PARA MANEJAR LOS ROLES DE LAS PERSONAS
'**************************************************************************
Public Function CargaTipoRoles() As ADODB.Recordset
Dim lsSql As String
Dim oConecta As COMConecta.DCOMConecta

    lsSql = " SELECT cConsDescripcion,nConsValor " _
          & " FROM Constante WHERE nConsCod = " & gPersTipo _
          & " ORDER BY cConsDescripcion "

On Error GoTo ERRORCargaTipoRoles
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaTipoRoles = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ERRORCargaTipoRoles:
    Err.Raise Err.Number, "Carga Tipo Roles", Err.Description
End Function

' Carga Lista Personas segun el rol seleccionado en el combo
Public Function CargaPersonas(ByVal pnTipoRol As Integer) As ADODB.Recordset

Dim lsSql As String

Dim loConecta As COMConecta.DCOMConecta

On Error GoTo ErrorCargaPersonas
 lsSql = "SELECT DISTINCT PTp.cPersCod, P.cPersNombre, PersEstado  = CASE PTp.nPersEstado WHEN 1 THEN 'ACTIVO' WHEN 2 THEN 'INACTIVO' " _
        & "ELSE 'NO DEFINIDO' END " _
        & "FROM PersTpo PTp INNER JOIN Persona P ON PTp.cPersCod = P.cPersCod  " _
        & "WHERE PTp.nPersTipo = " & pnTipoRol & " "

    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    Set CargaPersonas = loConecta.CargaRecordSet(lsSql)
    loConecta.CierraConexion
    Set loConecta = Nothing
    Exit Function
    
ErrorCargaPersonas:
    Err.Raise Err.Number, "Error Carga Personas", Err.Description
End Function


'Modificado por RIRO el 20130314
'Cambio permite saber si el cliente solicitado se encuentra activo o de baja dentro de la tabla PersTpo
Public Function ExistePersonaRol(ByVal psPersCod As String, ByVal pnTipoRol As PersTipo, _
                                Optional ByVal bEstado As Boolean = False) As Boolean
Dim lsSql As String
Dim lrReg As ADODB.Recordset
Dim loConecta As COMConecta.DCOMConecta

If Not bEstado Then
    lsSql = "SELECT * FROM PersTpo " & _
        "WHERE cPersCod = '" & psPersCod & " ' AND nPersTipo = '" & pnTipoRol & "' "
Else
    lsSql = "SELECT * FROM PersTpo " & _
        "WHERE cPersCod = '" & psPersCod & " ' AND nPersTipo = '" & pnTipoRol & "' " & _
        "and  nPersEstado = 1"
End If

    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    Set lrReg = loConecta.CargaRecordSet(lsSql)
    loConecta.CierraConexion
    Set loConecta = Nothing
    If lrReg.BOF And lrReg.EOF Then
       ExistePersonaRol = False
    Else
       ExistePersonaRol = True
   End If
End Function

Public Sub GrabarOperacionRolPersona(ByVal psOperacion As String, _
                                    ByVal psPersCod As String, _
                                    ByVal pnTipoRol As Integer, _
                                    ByVal pnPersTipo As Integer, _
                                    ByVal pnPersEstado As Integer, _
                                    ByRef psMensaje As String)
Dim lsSql As String
Dim loConex As COMConecta.DCOMConecta
On Error GoTo ErrorGrabarOperacion
If psOperacion = "A" Then ' Agregar
    If ExistePersonaRol(psPersCod, pnTipoRol) = False Then
        lsSql = "INSERT INTO PersTpo (cPersCod,nPersTipo,nPersEstado) " & _
                " VALUES ('" & psPersCod & "'," & pnPersTipo & "," & pnPersEstado & ") "
        Set loConex = New COMConecta.DCOMConecta
        loConex.AbreConexion
        loConex.Ejecutar lsSql
        Set loConex = Nothing
    Else
       psMensaje = "Persona ya se encuentra registrada en el rol"
       Exit Sub
    End If
    
Else  ' Editar
    lsSql = " UPDATE PersTpo SET " _
    & " nPersEstado = " & pnPersEstado & " " _
    & " WHERE cPersCod = '" & psPersCod & "' AND nPersTipo  = " & pnTipoRol
    
    Set loConex = New COMConecta.DCOMConecta
    loConex.AbreConexion
    loConex.Ejecutar lsSql
    Set loConex = Nothing
End If
    
Exit Sub

ErrorGrabarOperacion:
     Err.Raise Err.Number, "Error Grabar Operacion Rol", Err.Description
End Sub

'ALPA 20150205****************************************
Public Function GetCargarDatosParametroPermisoAprobacion() As ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSql = "exec stp_sel_ParametroPermisoAprobacion"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetCargarDatosParametroPermisoAprobacion = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Sub GetActualizarDatosParametroPermisoAprobacion(ByVal psCargo As String, ByVal psMovNroSol As String, ByVal pnEstado As Integer)
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    psCargo = Right("000000" & psCargo, 6)
    lsSql = "exec stp_ins_ParametroPermisoAprobacion '" & psCargo & "','" & psMovNroSol & "'," & pnEstado
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
'*****************************************************
