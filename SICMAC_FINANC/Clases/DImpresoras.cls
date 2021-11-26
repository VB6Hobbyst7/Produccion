VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DImpresoras"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Public Function GetMaquina() As String
    Dim buffMaq As String
    Dim lSizeMaq As Long
    buffMaq = Space(255)
    lSizeMaq = Len(buffMaq)
    GetComputerName buffMaq, lSizeMaq
    GetMaquina = Trim(Left$(buffMaq, lSizeMaq))
End Function

Public Function GetImpreSetup(psMaquina As String) As Integer
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = "Select nTpoCar from ImpreSetup Where cPCCod = '" & psMaquina & "'"
    
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        'GetImpreSetup = -1
        GetImpreSetup = 0
    Else
        GetImpreSetup = rs!nTpoCar
    End If
    
    Set oCon = Nothing
End Function

Public Function SetImpreSetup(psMaquina As String, psTpoImpresora As Impresoras) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    sql = " Delete ImpreSetup Where cPCCod = '" & psMaquina & "'"
    oCon.CargaRecordSet (sql)
    
    sql = " Insert ImpreSetup (cPCCod, nTpoCar)" _
        & " Values ('" & psMaquina & "'," & psTpoImpresora & ")"
    oCon.CargaRecordSet (sql)
    
    Set oCon = Nothing
End Function
