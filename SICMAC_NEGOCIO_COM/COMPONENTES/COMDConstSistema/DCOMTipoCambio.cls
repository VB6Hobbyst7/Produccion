VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsFormatoFecha As String
Dim oCon As COMConecta.DCOMConecta
Dim oerror As New COMConecta.COMErrorHandling

Private Sub Class_Initialize()

    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then
        oerror.RaiseError oerror.MyUnhandledError, "DTipoCambio:Initialize Method. Error en Conexion de Base de datos"
        Exit Sub
    End If
End Sub

Public Sub inicio(ByVal psFormatoFecha As String)
    vsFormatoFecha = psFormatoFecha
End Sub
'ALPA 20081003*******************************************************************************
'Se agrego el parametro pnPonREU para el REU
'ALPA 20140225*******************************************************************************
'Se agrego los parametros pnSBSDia,pnCompraTraider,pnVentaTraider
'********************************************************************************************
Public Function InsertaTipoCambio(ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
                                    ByVal pnValVentEsp As Currency, ByVal pnValCompEsp As Currency, _
                                    ByVal pnValFijo As Currency, ByVal pnValFijoDia As Currency, ByVal pnValPonderado As Currency, _
                                    ByVal psUltimaActualizacion As String, _
                                    Optional ByVal pbEjectBatch As Boolean, Optional ByVal pnPondVenta As Currency, Optional ByVal pnPonREU As Currency = 0, _
                                    Optional ByVal pnSBSDia As Currency = 0, Optional ByVal pnCompraTraider As Currency = 0, Optional ByVal pnVentaTraider As Currency = 0) As Integer
    Dim sql As String
    InsertaTipoCambio = 1
    sql = " INSERT INTO TIPOCAMBIO(dFecCamb,nValVent,nValComp,nValVentESP,nValCompEsp,nValFijo,nValFijoDia, cUltimaActualizacion,nValPond,nValPondVenta,nValPondREU,nValSBSDia,nValCompTr,nValVentTr) " _
        & " Values('" & Format(pdFecCamb & " " & oCon.GetHoraServer, vsFormatoFecha & " hh:mm:ss AMPM") & "'," _
        & pnValVent & "," & pnValComp & "," & pnValVentEsp & "," & pnValCompEsp & "," & pnValFijo & "," & pnValFijoDia & ",'" & psUltimaActualizacion & "'," & pnValPonderado & "," & pnPondVenta & "," & pnPonREU & "," & pnSBSDia & ", " & pnCompraTraider & ", " & pnVentaTraider & ")"
    
    If pbEjectBatch Then
        oCon.AdicionaCmdBatch sql
    Else
        oCon.Ejecutar sql
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
        & " WHERE DATEDIFF(day,dFecCamb,'" & Format(pdFecCamb, vsFormatoFecha) & "')=0"
    If pbEjectBatch Then
        oCon.AdicionaCmdBatch sql
    Else
        oCon.Ejecutar sql
    End If
    ActualizaTipoCambioDiario = 0
End Function
'ALPA 20081003*******************************************************************************
'Se agrego el parametro pnPonREU para el REU
'********************************************************************************************
Public Function ActualizaTipoCambio(ByVal pdFecCamb As Date, ByVal psUltimaActualizacion As String, _
                                    Optional ByVal pnValVent As Currency = -99999, Optional ByVal pnValComp As Currency = -99999, _
                                    Optional ByVal pnValVentEsp As Currency = -99999, Optional ByVal pnValCompEsp As Currency = -99999, _
                                    Optional ByVal pnValFijo As Currency = -99999, _
                                    Optional ByVal pnValFijoDia As Currency = -99999, Optional ByVal pnValPonderado As Currency = -99999, _
                                    Optional ByVal pbEjectBatch As Boolean, Optional ByVal pnValPondVenta As Currency, Optional ByVal pnPonREU As Currency = 0, _
                                    Optional ByVal pnSBSDia As Currency = 0, Optional ByVal pnCompraTraider As Currency = 0, Optional ByVal pnVentaTraider As Currency = 0) As Integer
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
    
    If pnValPondVenta <> -99999 Then
        lsFiltro = lsFiltro + ",nValPondVenta=" & pnValPondVenta
    End If
    If pnPonREU <> -99999 Then
        lsFiltro = lsFiltro + ",nValPondREU=" & pnPonREU
    End If
    'ALPA20140225*********************************************
    If pnSBSDia <> -99999 Then
        lsFiltro = lsFiltro + ",nValSBSDia= " & pnSBSDia
    End If
    If pnCompraTraider <> -99999 Then
        lsFiltro = lsFiltro + ",nValCompTr= " & pnCompraTraider
    End If
    If pnVentaTraider <> -99999 Then
        lsFiltro = lsFiltro + ",nValVentTr= " & pnVentaTraider
    End If
    '*********************************************************
    
    If lsFiltro <> "" Then
        sql = " UPDATE TIPOCAMBIO SET  cUltimaActualizacion='" & psUltimaActualizacion & "' " & lsFiltro _
            & " WHERE dFecCamb = '" & Format(pdFecCamb, vsFormatoFecha & " hh:mm:ss AMPM") & "'"
    End If
    If pbEjectBatch Then
        oCon.AdicionaCmdBatch sql
    Else
        oCon.Ejecutar sql
    End If
    ActualizaTipoCambio = 0
End Function

Public Function EjecutaBatch() As Integer
    EjecutaBatch = oCon.EjecutarBatch
End Function

'ALPA20140225*******************************************
Public Function ObtenerTipoCambioDiario(ByVal pdFecCamb As Date) As ADODB.Recordset
    Dim sql As String
    
    sql = "exec stp_sel_obtener_TipoCambio '" & Format(pdFecCamb, "YYYY/MM/DD") & "'"
    Set ObtenerTipoCambioDiario = oCon.CargaRecordSet(sql)
End Function
'*******************************************************

