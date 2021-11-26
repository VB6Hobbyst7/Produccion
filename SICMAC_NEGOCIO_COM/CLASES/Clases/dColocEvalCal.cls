VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dColocEvalCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'* Modulo de Colocaciones Pignoraticio
'* Clase con todas las Actualizaciones a BD
Option Explicit
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String

Dim coConex As DConecta

Private Sub Class_Initialize()
    Dim loIni As ClasIni
    
    Set loIni = New ClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing
    
Set coConex = New DConecta
If coConex.AbreConexion(csConexion) = False Then
    Call RaiseError(MyUnhandledError, "DColPCalculos:Initialize. Error en Conexion a Base de datos")
End If

End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

Public Sub dInsertColocEvalCalif(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psEvalCalif As String, ByVal psUltimaAct As String, ByVal psEvalObs As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocEvalCalif (cPersCod, nEvalTipo, cEvalCalif, cUltimaAct, cEvalObs ) " _
        & "VALUES ('" & psPersCod & "'," & pnEvalTipo & ",'" & psEvalCalif & "','" _
        & psUltimaAct & "','" & psEvalObs & "'  )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub


Public Sub dInsertColocEvalCalifDetalle(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psCtaCod As String, ByVal pdEval As Date, _
        ByVal psEvalEstado As String, ByVal psEvalCalifDet As String, _
        ByVal pnSaldoCap As Double, ByVal pnDiasAtraso As Integer, _
        ByVal psUltimaAct As String, ByVal psEvalObsDet As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocEvalCalifDetalle (cPersCod, nEvalTipo, cCtaCod, dEval, cEvalEstado, cEvalCalifDet, " _
        & "nSaldoCap, nDiasAtraso, cUltimaAct, cEvalObsDet ) " _
        & "VALUES ('" & psPersCod & "'," & pnEvalTipo & ",'" & psCtaCod & "','" & Format(pdEval, "mm/dd/yyyy") & "','" _
        & psEvalEstado & "','" & psEvalCalifDet & "'," & pnSaldoCap & "," & pnDiasAtraso & ",'" & psUltimaAct & "','" & psEvalObsDet & "'  )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dUpdateColocEvalCalif(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        Optional ByVal psEvalCalif As String = "@", Optional ByVal psUltimaAct As String = "@", _
        Optional ByVal psEvalObs As String = "@", _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE ColocEvalCalif SET "

    If psEvalCalif <> "@" Then
        lsSQL = lsSQL & " cEvalCalif = '" & psEvalCalif & "',"
    End If
    If psUltimaAct <> "@" Then
        lsSQL = lsSQL & " cUltimaAct = '" & psUltimaAct & "',"
    End If
    If psEvalObs <> "@" Then
        lsSQL = lsSQL & " cEvalObs = '" & psEvalObs & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cPersCod ='" & psPersCod & "' AND nEvalTipo = " & pnEvalTipo & " "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dDeleteColocEvalCalif(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocEvalCalif " _
          & " WHERE cPersCod ='" & psPersCod & "' AND nEvalTipo = " & pnEvalTipo & " "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dUpdateColocEvalCalifDetalle(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psCtaCod As String, ByVal pdEval As Date, Optional ByVal psEvalEstado As String, _
        Optional ByVal psEvalCalifDet As String = "@", Optional ByVal pnSaldoCap As Double = -999, _
        Optional ByVal pnDiasAtraso As Integer = -999, Optional ByVal psUltimaAct As String = "@", _
        Optional ByVal psEvalObsDet As String = "@", _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE ColocEvalCalifDetalle SET "

    If psEvalEstado <> "@" Then
        lsSQL = lsSQL & " cEvalEstado = '" & psEvalEstado & "',"
    End If
    If psEvalCalifDet <> "@" Then
        lsSQL = lsSQL & " cEvalCalifDet = '" & psEvalCalifDet & "',"
    End If
    If pnSaldoCap <> -999 Then
         lsSQL = lsSQL & " nSaldoCap = " & pnSaldoCap & ","
    End If
    If pnDiasAtraso <> -999 Then
         lsSQL = lsSQL & " nDiasAtraso = " & pnDiasAtraso & ","
    End If
    If psUltimaAct <> "@" Then
        lsSQL = lsSQL & " cUltimaAct = '" & psUltimaAct & "',"
    End If
    If psEvalObsDet <> "@" Then
        lsSQL = lsSQL & " cEvalObsDet = '" & psEvalObsDet & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cPersCod ='" & psPersCod & "' AND nEvalTipo = " & pnEvalTipo _
                  & " AND cCtaCod ='" & psCtaCod & "' AND dEval = '" & Format(pdEval, "yyyy/mm/dd") & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dDeleteColocEvalCalifDetalle(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psCtaCod As String, ByVal pdEval As Date, ByVal psEvalEstado As String, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocEvalCalifDetalle  " _
          & " WHERE cPersCod ='" & psPersCod & "' AND nEvalTipo = " & pnEvalTipo _
                  & " AND cCtaCod ='" & psCtaCod & "' AND dEval = '" & Format(pdEval, "mm/dd/yyyy") & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


