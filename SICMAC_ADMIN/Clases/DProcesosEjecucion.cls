VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DProcesosEjecucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim oConect As DConecta

Private Sub Class_Initialize()
    Set oConect = New DConecta
    oConect.AbreConexion
End Sub

Public Function ProcesoEjecucionValida(pdFecha As Date, pnSistema As Integer, psProceso As String) As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select A.nID, nTipo, nDia, dFecha, dHoraIni, dHoraFin from ProcesosProgEjecucion A" _
        & " Inner Join ProcesosProgEjecucionDet B On A.nId = B.nId" _
        & " Where cProceso = '" & psProceso & "' And bEstado =  1 And nSistema = " & pnSistema & " And  convert(varchar(10),dHoraIni,108) <= '" & Format(pdFecha, "hh:mm:ss") & "' And convert(varchar(10),dHoraFin,108) >= '" & Format(pdFecha, "hh:mm:ss") & "' Order By A.nId Desc "
    Set rs = oConect.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        ProcesoEjecucionValida = False
    Else
        If rs!nTipo = 0 Then
            ProcesoEjecucionValida = True
        ElseIf rs!nTipo = 1 Then
            If rs!nDia = -1 Then
                If IsNull(rs!dFecha) Then
                    ProcesoEjecucionValida = False
                Else
                    If DateDiff("d", pdFecha, rs!dFecha) = 0 Then
                        If Format(pdFecha, "hhmmss") >= Format(rs!dHoraIni, "hhmmss") And Format(pdFecha, "hhmmss") <= Format(rs!dHoraFin, "hhmmss") Then
                            ProcesoEjecucionValida = True
                        Else
                            ProcesoEjecucionValida = False
                        End If
                    Else
                        ProcesoEjecucionValida = False
                    End If
                End If
            ElseIf rs!nDia = 0 Then
                If Format(pdFecha, "hhmmss") >= Format(rs!dHoraIni, "hhmmss") And Format(pdFecha, "hhmmss") <= Format(rs!dHoraFin, "hhmmss") Then
                    ProcesoEjecucionValida = True
                Else
                    ProcesoEjecucionValida = False
                End If
            Else
                If DatePart("w", pdFecha, vbMonday) = rs!nDia Then
                    If Format(pdFecha, "hhmmss") >= Format(rs!dHoraIni, "hhmmss") And Format(pdFecha, "hhmmss") <= Format(rs!dHoraFin, "hhmmss") Then
                        ProcesoEjecucionValida = True
                    Else
                        ProcesoEjecucionValida = False
                    End If
                End If
            End If
        ElseIf rs!nTipo = 2 Then
            sql = "Update ProcesosProgEjecucion Set bEstado =  0 Where nId = " & rs!nId
            oConect.Ejecutar sql
        Else
            ProcesoEjecucionValida = True
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub Class_Terminate()
    oConect.CierraConexion
    Set oConect = Nothing
End Sub

