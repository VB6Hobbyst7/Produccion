VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String
Dim oConect As DConecta
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
Set oConect = New DConecta

If oConect.AbreConexion = False Then
    Call RaiseError(MyUnhandledError, "DTipoCambio:Initialize Method. Error en Conexion de Base de datos")
    Exit Sub
End If
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub

'Public Function InsertaTipoCambio(ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
'                                    ByVal pnValVentEsp As Currency, ByVal pnValCompEsp As Currency, _
'                                    ByVal pnValFijo As Currency, ByVal pnValFijoDia As Currency, ByVal pnValPonderado As Currency, _
'                                    ByVal psUltimaActualizacion As String, _
'                                    Optional ByVal pbEjectBatch As Boolean) As Integer
'    Dim sql As String
'    InsertaTipoCambio = 1
'    sql = " INSERT INTO TIPOCAMBIO(dFecCamb,nValVent,nValComp,nValVentESP,nValCompEsp,nValFijo,nValFijoDia, cUltimaActualizacion,nValPond) " _
'        & " Values('" & Format(pdFecCamb & " " & oConect.GetHoraServer, gsFormatoFechaHora) & "'," _
'        & pnValVent & "," & pnValComp & "," & pnValVentEsp & "," & pnValCompEsp & "," & pnValFijo & "," & pnValFijoDia & ",'" & psUltimaActualizacion & "'," & pnValPonderado & ")"
'
'    If pbEjectBatch Then
'        oConect.AdicionaCmdBatch sql
'    Else
'        oConect.Ejecutar sql
'    End If
'    InsertaTipoCambio = 0
'End Function

Public Function InsertaTipoCambio(ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
                                    ByVal pnValVentEsp As Currency, ByVal pnValCompEsp As Currency, _
                                    ByVal pnValFijo As Currency, ByVal pnValFijoDia As Currency, ByVal pnValPonderado As Currency, _
                                    ByVal psUltimaActualizacion As String, ByVal pnValPonderadoVenta As Currency, _
                                    Optional ByVal pbEjectBatch As Boolean) As Integer
    Dim sql As String
    InsertaTipoCambio = 1
    sql = " INSERT INTO TIPOCAMBIO(dFecCamb,nValVent,nValComp,nValVentESP,nValCompEsp,nValFijo,nValFijoDia, cUltimaActualizacion,nValPond,nValPondVenta) " _
        & " Values('" & Format(pdFecCamb & " " & oConect.GetHoraServer, gsFormatoFechaHora) & "'," _
        & pnValVent & "," & pnValComp & "," & pnValVentEsp & "," & pnValCompEsp & "," & pnValFijo & "," & pnValFijoDia & ",'" & psUltimaActualizacion & "'," & pnValPonderado & "," & pnValPonderadoVenta & ")"

    If pbEjectBatch Then
        oConect.AdicionaCmdBatch sql
    Else
        oConect.Ejecutar sql
    End If
    InsertaTipoCambio = 0
End Function

Public Function ActualizaTipoCambioDiario(ByVal pdFecCamb As Date, ByVal psUltimaActualizacion As String, _
                                            ByVal pnValFijoDia As Currency, _
                                            Optional ByVal pbEjectBatch As Boolean) As Integer
    Dim sql As String
    ActualizaTipoCambioDiario = 1
    
    sql = " UPDATE TIPOCAMBIO" _
        & " SET cUltimaActualizacion = '" & psUltimaActualizacion & "', nValFijoDia=" & pnValFijoDia & "  " _
        & " WHERE DATEDIFF(day,dFecCamb,'" & Format(pdFecCamb, gsFormatoFecha) & "')=0"
    If pbEjectBatch Then
        oConect.AdicionaCmdBatch sql
    Else
        oConect.Ejecutar sql
    End If
    ActualizaTipoCambioDiario = 0
End Function

'Public Function ActualizaTipoCambio(ByVal pdFecCamb As Date, ByVal psUltimaActualizacion As String, _
'                                    Optional ByVal pnValVent As Currency = -99999, Optional ByVal pnValComp As Currency = -99999, _
'                                    Optional ByVal pnValVentEsp As Currency = -99999, Optional ByVal pnValCompEsp As Currency = -99999, _
'                                    Optional ByVal pnValFijo As Currency = -99999, _
'                                    Optional ByVal pnValFijoDia As Currency = -99999, Optional ByVal pnValPonderado As Currency = -99999, _
'                                    Optional ByVal pbEjectBatch As Boolean) As Integer
'    Dim sql As String
'    Dim lsFiltro As String
'    ActualizaTipoCambio = 1
'    lsFiltro = ""
'    If pnValVent <> -99999 Then
'        lsFiltro = lsFiltro + ",nValVent=" & pnValVent
'    End If
'    If pnValComp <> -99999 Then
'        lsFiltro = lsFiltro + ",nValComp=" & pnValComp
'    End If
'    If pnValVentEsp <> -99999 Then
'        lsFiltro = lsFiltro + ",nValVentEsp=" & pnValVentEsp
'    End If
'    If pnValCompEsp <> -99999 Then
'        lsFiltro = lsFiltro + ",nValCompEsp=" & pnValCompEsp
'    End If
'    If pnValFijo <> -99999 Then
'        lsFiltro = lsFiltro + ",nValFijo=" & pnValFijo
'    End If
'    If pnValFijoDia <> -99999 Then
'        lsFiltro = lsFiltro + ",nValFijoDia=" & pnValFijoDia
'    End If
'    If pnValPonderado <> -99999 Then
'        lsFiltro = lsFiltro + ",nValPond=" & pnValPonderado
'    End If
'
'
'
'    If lsFiltro <> "" Then
'        sql = " UPDATE TIPOCAMBIO SET  cUltimaActualizacion='" & psUltimaActualizacion & "' " & lsFiltro _
'            & " WHERE dFecCamb = '" & Format(pdFecCamb, gsFormatoFecha & " hh:mm:ss AMPM") & "'"
'    End If
'    If pbEjectBatch Then
'        oConect.AdicionaCmdBatch sql
'    Else
'        oConect.Ejecutar sql
'    End If
'    ActualizaTipoCambio = 0
'End Function

Public Function ActualizaTipoCambio(ByVal pdFecCamb As Date, ByVal psUltimaActualizacion As String, _
                                    Optional ByVal pnValVent As Currency = -99999, Optional ByVal pnValComp As Currency = -99999, _
                                    Optional ByVal pnValVentEsp As Currency = -99999, Optional ByVal pnValCompEsp As Currency = -99999, _
                                    Optional ByVal pnValFijo As Currency = -99999, _
                                    Optional ByVal pnValFijoDia As Currency = -99999, Optional ByVal pnValPonderado As Currency = -99999, _
                                    Optional ByVal pbEjectBatch As Boolean, Optional ByVal pnValPonderadoVenta As Currency = -99999) As Integer
    Dim sql As String
    Dim lsFiltro As String
    ActualizaTipoCambio = 1
    lsFiltro = ""
    If pnValVent <> -99999 Then
        lsFiltro = lsFiltro + ",nValVent=" & pnValVent
    End If
    If pnValComp <> -99999 Then
        lsFiltro = lsFiltro + ",nValComp=" & pnValComp
    End If
    If pnValVentEsp <> -99999 Then
        lsFiltro = lsFiltro + ",nValVentEsp=" & pnValVentEsp
    End If
    If pnValCompEsp <> -99999 Then
        lsFiltro = lsFiltro + ",nValCompEsp=" & pnValCompEsp
    End If
    If pnValFijo <> -99999 Then
        lsFiltro = lsFiltro + ",nValFijo=" & pnValFijo
    End If
    If pnValFijoDia <> -99999 Then
        lsFiltro = lsFiltro + ",nValFijoDia=" & pnValFijoDia
    End If
    If pnValPonderado <> -99999 Then
        lsFiltro = lsFiltro + ",nValPond=" & pnValPonderado
    End If
    If pnValPonderadoVenta <> -99999 Then
        lsFiltro = lsFiltro + ",nValPondVenta=" & pnValPonderadoVenta
    End If
    
    
    If lsFiltro <> "" Then
        sql = " UPDATE TIPOCAMBIO SET  cUltimaActualizacion='" & psUltimaActualizacion & "' " & lsFiltro _
            & " WHERE dFecCamb = '" & Format(pdFecCamb, gsFormatoFecha & " hh:mm:ss AMPM") & "'"
    End If
    If pbEjectBatch Then
        oConect.AdicionaCmdBatch sql
    Else
        oConect.Ejecutar sql
    End If
    ActualizaTipoCambio = 0
End Function

Public Function EjecutaBatch() As Integer
    EjecutaBatch = oConect.EjecutarBatch
End Function

'ALPA 20140331****************************************************************************************
Public Sub ActualizarAprobacionMovCompraVenta(ByVal psMovNro As String)
    Dim sql As String
    sql = "exec stp_upd_AprobacionMovCompraVenta '" & psMovNro & "'"
    oConect.Ejecutar sql
End Sub

Public Function ObtenerNivelesAprobacionCompraVentaxMonto(ByVal pnMonto As Currency, ByVal nTipoEspecial As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim lrDatos As ADODB.Recordset
    sSql = "exec stp_sel_ObtenerNivelesAprobacionCompraVentaxMonto '" & pnMonto & "'," & nTipoEspecial & ""
    Set ObtenerNivelesAprobacionCompraVentaxMonto = oConect.CargaRecordSet(sSql)
End Function

Public Sub AprobacionMovCompraVenta(ByVal cMovNro As String, ByVal cOpeCod As String, ByVal cNivelCod As String, ByVal nMonto As Currency, ByVal nTipoCambioNormal As Currency, ByVal nTipoCambioSolicitado As Currency, ByVal nTipoEspecial As Integer, ByVal psPersCod As String)
    Dim lsSql As String
    lsSql = "exec stp_sel_AprobacionMovCompraVenta '" & cMovNro & "','" & cOpeCod & "','" & cNivelCod & "','" & nMonto & "'," & nTipoCambioNormal & "," & nTipoCambioSolicitado & "," & nTipoEspecial & ",'" & psPersCod & "'"
    oConect.Ejecutar (lsSql)
End Sub

Public Function ObtenerCantidadAprobacionMovCompraVenta(ByVal psMovNro As String) As ADODB.Recordset
    Dim lsSql As String
    Dim lrDatos As ADODB.Recordset
    lsSql = "exec stp_sel_ObtenerCantidadAprobacionMovCompraVenta '" & psMovNro & "'"
    Set ObtenerCantidadAprobacionMovCompraVenta = oConect.CargaRecordSet(lsSql)
End Function
'*****************************************************************************************************

