VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMConstSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim vsFormatoFecha As String
Public Sub inicio(ByVal psFormatoFecha As String)
    vsFormatoFecha = psFormatoFecha
End Sub

Public Function ActualizaConstSistemas(ByVal pnConsSistCod As ConstSistemas, ByVal psMovUltAct As String, _
                                Optional ByVal psConsSistValor As String = "", _
                                Optional ByVal psConsSistDesc As String = "", Optional pbEjecBatch As Boolean = True) As Integer
                                       
    Dim oConSist As DCOMConstSistema
    Set oConSist = New DCOMConstSistema
    ActualizaConstSistemas = oConSist.ActualizaConsSistemas(pnConsSistCod, psMovUltAct, psConsSistValor, psConsSistDesc, pbEjecBatch)
                                           
    Set oConSist = Nothing
End Function
    
    Public Function LeeConstSistema(ByVal psConstSistCod As ConstSistemas) As String
    Dim rsVar As Recordset
    Dim sSQL As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    sSQL = "Select nConsSisDesc, nConsSisValor From ConstSistema where nConsSisCod =" & psConstSistCod & ""
    Set rsVar = New Recordset
    Set rsVar = oCon.CargaRecordSet(sSQL)
    LeeConstSistema = ""
    If Not rsVar.EOF And Not rsVar.BOF Then
        LeeConstSistema = rsVar("nConsSisValor")
    End If
    rsVar.Close
    Set rsVar = Nothing
    Set oCon = Nothing
End Function

Public Function GetRutaAcceso(psAgenciaCod As String) As String
    Dim sql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    sql = "Select cRuta From RutaActualizacion Where cAgeCod = '" & psAgenciaCod & "'"
    
    oCon.AbreConexion
    GetRutaAcceso = oCon.CargaRecordSet(sql).Fields(0)
    oCon.CierraConexion
    
    Set oCon = Nothing
End Function

Public Function ObtenerValoresCargaLogin(ByVal psAgenciaCod As String, _
                                        ByRef psEmpresa As String, _
                                        ByRef psRutaUltActualiz As String, _
                                        ByRef psFlagActualizaEXE As String, _
                                        ByRef psProyectoActual As String)

    psEmpresa = LeeConstSistema(gConstSistCMACNombreCompleto)
    psRutaUltActualiz = GetRutaAcceso(psAgenciaCod)
    psFlagActualizaEXE = LeeConstSistema(49)
    psProyectoActual = LeeConstSistema(300)
End Function

