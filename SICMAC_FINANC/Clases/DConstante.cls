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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset
Public sConexion As String
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

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
    Dim oConec As New ClasIni
    sCentralCom = oConec.BaseComunes
    Set oConec = Nothing
End Sub

