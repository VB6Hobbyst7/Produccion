VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DInstFinanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum TInstFinExistenciaRegistro
    gInstFinExistenciaRegistro = 1
    gInstFinExistenciaReferencia = 2
End Enum
Private gConsPersona As String
Private gConsComunes As String
Private gConsImagenes As String
Public Function CargaCmaCs() As Recordset
Dim sSql As String
Dim oConecta As DConecta

On Error GoTo ErrorCargaInstituciones
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "Select I.cPersCod,P.cPersNombre, C.cConsDescripcion + space(50) + I.cIFTpo as sTipo, C.cConsDescripcion, I.cSubCtaContCod "
    sSql = sSql & " From InstitucionFinanc I inner join Persona P on I.cPersCod = P.cPersCod "
    sSql = sSql & " Inner join Constante C ON convert(int,I.cIFTpo) = C.nConsValor"
    sSql = sSql & " WHERE C.nConsCod = '" & gCGTipoIF & "' And nConsValor <> '" & gCGTipoIF & "'"
    sSql = sSql & " AND cIFTpo=3 Order By P.cPersNombre "
    
    Set CargaCmaCs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargaInstituciones:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Instituciones Financieras", Err.Description
End Function

Public Function CargaInstituciones() As Recordset
Dim sSql As String
Dim oConecta As DConecta

On Error GoTo ErrorCargaInstituciones
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "Select I.cPersCod,P.cPersNombre, C.cConsDescripcion + space(50) + I.cIFTpo as sTipo, C.cConsDescripcion, I.cSubCtaContCod "
    sSql = sSql & " From InstitucionFinanc I inner join Persona P on I.cPersCod = P.cPersCod "
    sSql = sSql & " Inner join Constante C ON convert(int,I.cIFTpo) = C.nConsValor"
    sSql = sSql & " WHERE C.nConsCod = '" & gCGTipoIF & "' And nConsValor <> '" & gCGTipoIF & "'"
    sSql = sSql & " Order By P.cPersNombre "
    
    Set CargaInstituciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargaInstituciones:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Instituciones Financieras", Err.Description
End Function

Public Sub InsertaNuevaInstitucion(ByVal psPersCod As String, ByVal psTipoInst As String, ByVal psSubCta As String)
Dim sSql As String
Dim oConecta As DConecta

On Error GoTo ErrorInsertaNuevaInstitucion
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "INSERT INTO InstitucionFinanc(cPersCod,cIFTpo,cSubCtaContCod)"
    sSql = sSql & "VALUES('" & psPersCod & "','" & Right("00" & psTipoInst, 2) & "','" & psSubCta & "')"
    oConecta.Ejecutar sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertaNuevaInstitucion:
    Err.Raise Err.Number, "Nueva Institucion Financiera", Err.Description
End Sub

Public Sub ActualizaInstitucion(ByVal psPersCod As String, ByVal psTipoInst As String, ByVal psSubCta As String, ByVal psTipoInstAnt As String)
Dim sSql As String
Dim oConecta As DConecta

On Error GoTo ErrorActualizaInstitucion
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "UPDATE InstitucionFinanc SET cIFTpo = '" & Right("00" & Trim(psTipoInst), 2) & "', cSubCtaContCod = '" & psSubCta & "'"
    sSql = sSql & " Where cPersCod = '" & psPersCod & "' And cIFTpo = '" & Right("00" & Trim(psTipoInstAnt), 2) & "'"
    oConecta.Ejecutar sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizaInstitucion:
    Err.Raise Err.Number, "Actualizar Institucion Financiera", Err.Description
End Sub

Public Sub EliminarInstitucion(ByVal psPersCod As String, ByVal psTpoInst As String)
Dim sSql As String
Dim oConecta As DConecta

On Error GoTo ErrorActualizaInstitucion
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "DELETE InstitucionFinanc "
    sSql = sSql & " Where cPersCod = '" & psPersCod & "' And cIFTpo = '" & Right("00" & Trim(psTpoInst), 2) & "'"
    oConecta.Ejecutar sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizaInstitucion:
    Err.Raise Err.Number, "Actualizar Institucion Financiera", Err.Description
End Sub

Public Function CargaChequesBanco(ByVal psPersCod As String, ByVal pChqEstado As ChequeEstado) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = "Select * From DocRec Where cPerscod = '" & psPersCod & "' And nEstado = " & pChqEstado
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set CargaChequesBanco = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function CargaChequesBancoCred(ByVal psPersCod As String, ByVal pChqEstado As ChequeEstado) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = "Select * From DocRec Where cPerscod = '" & psPersCod & "' And nEstado = " & pChqEstado & " and nProducto not in (232,233,234)"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set CargaChequesBancoCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaBancos(Optional ByVal pbTodos As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = "Select I.cPersCod, P.cpersNombre  "
    sSql = sSql & " From InstitucionFinanc I Inner Join Persona P ON I.cPersCod = P.cPersCod "
    If Not pbTodos Then
        sSql = sSql & " where I.cIFTpo='01' "
    End If
    sSql = sSql & " Order by P.cPersNombre "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaBancos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaIFinancieraPersCod(Optional ByVal pCombo As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaIFinancieraPersCod
    Set oConecta = New DConecta
    If pCombo Then
        sSql = "Select P.cPersNombre + space(50) +  P.cpersCod  from InstitucionFinanc IFin inner join " & gConsPersona & "Persona P on IFin.cPersCod = P.cpersCod where cIFTpo='05' order by P.cPersNombre"
    Else
        sSql = "select P.cpersCod,P.cPersNombre from InstitucionFinanc IFin inner join " & gConsPersona & "Persona P on IFin.cPersCod = P.cpersCod where cIFTpo='05' order by P.cPersNombre"
    End If
    oConecta.AbreConexion
    Set RecuperaIFinancieraPersCod = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaIFinancieraPersCod:
    Err.Raise Err.Number, "Recupera Institucion Financiera y Codigo de Persona", Err.Description
    
End Function

Public Function ExisteRegistro(ByVal psPersCod As String, ByVal psIFTpo As String, ByVal gOpcExisReg As TInstFinExistenciaRegistro) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConec As DConecta

    On Error GoTo ErrorExisteInstitucion
    Set oConec = New DConecta
    oConec.AbreConexion
    If gOpcExisReg = gInstFinExistenciaRegistro Then
        sSql = "Select cPersCod,cIFTpo from " & gConsPersona & "InstitucionFinanc "
        sSql = sSql & " Where cPerscod = '" & psPersCod & "' And cIFTpo = '" & Right("00" & Trim(psIFTpo), 2) & "'"
    End If
    If gOpcExisReg = gInstFinExistenciaReferencia Then
        sSql = "Select cPersCod,cIFTpo from CtaIF "
        sSql = sSql & " Where cPerscod = '" & psPersCod & "' And cIFTpo = '" & Right("00" & Trim(psIFTpo), 2) & "'"
    End If
    
    Set R = oConec.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        ExisteRegistro = True
    Else
        ExisteRegistro = False
    End If
    R.Close
    Set R = Nothing
    oConec.CierraConexion
    Exit Function

ErrorExisteInstitucion:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function


Public Function GetIFSubCuenta(ByVal psPersCod As String, ByVal psIFTpo As String) As String
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConec As DConecta

    On Error GoTo ErrorGetIFSubCuenta
    Set oConec = New DConecta
    oConec.AbreConexion
    sSql = "SELECT cPersCod, cIFTpo,cSubCtaContCod " _
         & "FROM InstitucionFinanc " _
         & "WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' "
    Set R = oConec.CargaRecordSet(sSql)
    If Not R.RecordCount = 0 Then
        GetIFSubCuenta = R!cSubCtaContCod
    End If
    RSClose R
    oConec.CierraConexion
    Exit Function

ErrorGetIFSubCuenta:
        Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim CIni As ClsIni.ClasIni
    Set CIni = New ClsIni.ClasIni
    gConsPersona = CIni.BasePersonas
    gConsComunes = CIni.BaseComunes
    gConsImagenes = CIni.BaseImagenes
    Set CIni = Nothing
    
End Sub


