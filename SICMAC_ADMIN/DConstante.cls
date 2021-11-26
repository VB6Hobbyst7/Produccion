VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DConstante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public bError As Boolean
Public sMsgError As String
Private sCentralCom As String

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
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorDameDescripcionConstante
    sSql = "Select cConsDescripcion from Constante where nConsCod = " & pnCodigo & " AND nConsValor = " & pnValor
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
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
Private Sub Class_Initialize()
Dim ClsIni As ClsIni.ClasIni
    Set ClsIni = New ClsIni.ClasIni
    sCentralCom = ClsIni.BaseComunes
    Set ClsIni = Nothing
End Sub
