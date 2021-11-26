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
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsFormatoFecha As String
Dim oCon As DConecta
Public Sub Inicio(ByVal psFormatoFecha As String)
vsFormatoFecha = psFormatoFecha
End Sub

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
Set oCon = New DConecta

If oCon.AbreConexion = False Then
    Call RaiseError(MyUnhandledError, "DMov:Initialize Method. Error en Conexion a Base de datos")
End If
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub
Public Function ActualizaConsSistemas(ByVal pnConsSistCod As ConstSistemas, ByVal psMovUltAct As String, _
                                Optional ByVal psConsSistValor As String = "", _
                                Optional ByVal psConsSistDesc As String = "") As Integer
Dim sql As String
Dim lsFiltro As String
lsFiltro = ""
If psConsSistValor <> "" Then
    lsFiltro = " ,nConsSisValor= '" & psConsSistValor & "' "
End If
If psConsSistDesc <> "" Then
    lsFiltro = " ,nConsSisDesc=" & psConsSistDesc & ""
End If
ActualizaConsSistemas = 1
If lsFiltro <> "" Then
    sql = "UPDATE ConstSistema SET cUltimaActuzalizacion='" & psMovUltAct & "' " & lsFiltro _
        & " WHERE nConsSisCod =" & pnConsSistCod & ""
    
End If
oCon.Ejecutar sql
ActualizaConsSistemas = 0

End Function
Private Sub Class_Terminate()
oCon.CierraConexion
Set oCon = Nothing
End Sub

Public Function RecuperaUltimaFecha(ByVal pnConsSistCod As ConstSistemas) As Date
Dim sql As String
Dim rs As ADODB.Recordset

On Error GoTo RecuperaFechaErr
    sql = "Select nConsSisValor FROM ConstSistema " & _
          " WHERE nConsSisCod =" & pnConsSistCod & ""
          
    Set rs = New ADODB.Recordset
    Set rs = oCon.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        RecuperaUltimaFecha = rs(0)
    End If
    Set rs = Nothing
    Exit Function
RecuperaFechaErr:
    Call RaiseError(MyUnhandledError, "DMov:BuscarMov Method")
   

oCon.Ejecutar sql

End Function
'ALPA 20110430
Public Function LeeConstSistema(ByVal psConstSistCod As ConstSistemas) As String
    Dim rsVar As Recordset
    Dim sSql As String
    Dim oCon  As DConecta
    Set oCon = New DConecta
    
    If oCon.AbreConexion = False Then Exit Function
    sSql = "Select nConsSisDesc, nConsSisValor From ConstSistema where nConsSisCod =" & psConstSistCod & ""
    Set rsVar = New Recordset
    Set rsVar = oCon.CargaRecordSet(sSql)
    LeeConstSistema = ""
    If Not rsVar.EOF And Not rsVar.BOF Then
        LeeConstSistema = rsVar("nConsSisValor")
    End If
    rsVar.Close
    Set rsVar = Nothing
    Set oCon = Nothing
End Function
