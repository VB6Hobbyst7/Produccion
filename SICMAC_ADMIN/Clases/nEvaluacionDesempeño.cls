VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nEvaluacionDesempeño"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Function Abrir_Periodo_EvaluacionDesempeño(ByVal pnPeriodo As Integer, _
                                                    ByVal pnAnio As Integer)
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorAbrir_Periodo_EvaluacionDesempeño
    sSQL = "sp_RHAbrir_Periodo_EvaluacionDesempeño " & pnPeriodo & "," & pnAnio
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
       oConecta.ConexionActiva.Execute sSQL
    '    pnRpta = oConecta.Ejecutar(sSQL)
        
    End If
    'oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function
ErrorAbrir_Periodo_EvaluacionDesempeño:
    Err.Raise Err.Number, "Abrir_Periodo_EvaluacionDesempeño", Err.Description
End Function

