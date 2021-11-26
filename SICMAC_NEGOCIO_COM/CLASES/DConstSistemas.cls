VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DConstSistemas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsConexion As String
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsFormatoFecha As String
Dim oCon As DConecta
Public Sub inicio(ByVal psFormatoFecha As String)
vsFormatoFecha = psFormatoFecha
End Sub

Private Sub Class_Initialize()
Dim oIni As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni
Set oCon = New DConecta

If oCon.AbreConexion = False Then
    Call RaiseError(MyUnhandledError, "DMov:Initialize Method. Error en Conexion a Base de datos")
End If
vsConexion = oIni.CadenaConexion
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub
Public Function ActualizaConsSistemas(ByVal pnConsSistCod As ConstSistemas, ByVal psMovUltAct As String, _
                                Optional ByVal psConsSistValor As String = "", _
                                Optional ByVal psConsSistDesc As String = "", Optional pbEjecBatch As Boolean = True) As Integer

Dim sql As String
Dim lsFiltro As String
lsFiltro = ""
If psConsSistValor <> "" Then
    lsFiltro = " ,nConsSisValor=" & psConsSistValor & ""
End If
If psConsSistDesc <> "" Then
    lsFiltro = " ,nConsSisDesc=" & psConsSistDesc & ""
End If
ActualizaConsSistemas = 1
If lsFiltro <> "" Then
    sql = "UPDATE ConstSistema SET cUltimaActuzalizacion='" & psMovUltAct & " " & lsFiltro _
        & " WHERE nConsSisCod =" & pnConsSistCod & ""
    
End If

If pbEjecBatch Then
    oCon.AdicionaCmdBatch sql
Else
    oCon.Ejecutar sql
End If
ActualizaConsSistemas = 0

End Function
Private Sub Class_Terminate()
oCon.CierraConexion
Set oCon = Nothing
End Sub
