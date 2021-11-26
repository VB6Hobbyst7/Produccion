VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Dim vsConexion As String
'Dim vsCentralPer As String
'Dim vsCentralCom As String
'Dim vsCentralImg As String
Dim sSQL As String


Private Sub Class_Initialize()
    'Dim oImp As DImpresoras
    'Set oImp = New DImpresoras
    
    'oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    'Set oImp = Nothing

    'Dim oIni As ClasIni
    
    'Set oIni = New ClasIni
        'vsConexion = oIni.CadenaConexion
        'vsCentralPer = oIni.BasePersonas
        'vsCentralCom = oIni.BaseComunes
        'vsCentralImg = oIni.BaseImagenes
    'Set oIni = Nothing
End Sub

'Devuelva un RecordSet con las agencias
Public Function CargaAgencias(ByVal gsCodAgeN As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSQL = "select cAgeCod, cAgeDescripcion from agencias where cAgecod = '" & gsCodAgeN & "' And nEstado = 1"
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    
    Set CargaAgencias = rs
End Function

Public Function DevuelveProveedorAG(ByVal sCodProv As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSQL = "select cAgeCod from proveedorAG where cPersCod='" & sCodProv & "'"
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    
    Set DevuelveProveedorAG = rs
End Function


