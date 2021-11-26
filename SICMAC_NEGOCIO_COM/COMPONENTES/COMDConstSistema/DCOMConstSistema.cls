VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMConstSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsFormatoFecha As String
Dim oCon As New COMConecta.DCOMConecta
Dim oerror As New COMConecta.COMErrorHandling
Dim ocons As New COMDConstantes.DCOMValores

Public Sub inicio(ByVal psFormatoFecha As String)
    vsFormatoFecha = psFormatoFecha
End Sub

Public Function ActualizaConsSistemas(ByVal pnConsSistCod As ConstSistemas, ByVal psMovUltAct As String, _
                                Optional ByVal psConsSistValor As String = "", _
                                Optional ByVal psConsSistDesc As String = "", Optional pbEjecBatch As Boolean = True) As Integer
    
    Dim Sql As String
    Dim lsFiltro As String
   
    
    lsFiltro = ""
    If psConsSistValor <> "" Then
        lsFiltro = " ,nConsSisValor=" & psConsSistValor & ""
    End If
    If psConsSistDesc <> "" Then
        lsFiltro = " ,nConsSisDesc=" & psConsSistDesc & ""
    End If
    ActualizaConsSistemas = 1

    If oCon.AbreConexion = False Then
       oerror.RaiseError oerror.MyUnhandledError, "DMov:Initialize Method. Error en Conexion a Base de datos"
    End If
    
    If lsFiltro <> "" Then
        Sql = "UPDATE ConstSistema SET cUltimaActuzalizacion='" & psMovUltAct & " " & lsFiltro _
            & " WHERE nConsSisCod =" & pnConsSistCod & ""
        
    End If
    
    If pbEjecBatch Then
        oCon.AdicionaCmdBatch Sql
    Else
        oCon.Ejecutar Sql
    End If
    ActualizaConsSistemas = 0
    
    oCon.CierraConexion
End Function

Public Function ObtenerVarSistema() As ADODB.Recordset
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dError
    lsSQL = " Select nConsSisValor From ConstSistema WHERE nConsSisCod = 107 "
     oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(lsSQL)
        Set ObtenerVarSistema = rs
    Set rs = Nothing
    Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function

Public Function ObtenerVarSistema80() As ADODB.Recordset
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dError
    lsSQL = "select * from ConstSistema where nConsSisCod=80"
     oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(lsSQL)
        Set ObtenerVarSistema80 = rs
    Set rs = Nothing
    Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function
'ALPA 20160623****************************************************************
Public Function ObtenerVarSistemaCargoMultiple(pnCodigo As Integer, psCargo As String) As Boolean
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dErrorObtenerVarSistemaCargoMultiple
    lsSQL = " stp_sel_ObtenerConstSistemaxCargo " & pnCodigo & ",'" & psCargo & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(lsSQL)
    If Not (rs.BOF Or rs.EOF) Then
        ObtenerVarSistemaCargoMultiple = True
    Else
        ObtenerVarSistemaCargoMultiple = False
    End If
    Set rs = Nothing
    Exit Function
dErrorObtenerVarSistemaCargoMultiple:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function
Public Function RecuperaSeguroMultiriesgo(ByVal psCtaCod As String) As Recordset  'GEMO 13/02/2020
  Dim lsSQL As String
  Dim nConsSisCod As Integer
  nConsSisCod = IIf(Mid(psCtaCod, 9, 1) = "1", 722, 723)
  Dim rs As New ADODB.Recordset
On Error GoTo dErrorRecuperaSeguroMultiriesgo
    lsSQL = "select nConsSisValor from Constsistema where nConsSisCod =" & nConsSisCod & ""
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(lsSQL)
        Set RecuperaSeguroMultiriesgo = rs
    Set rs = Nothing
    Exit Function
dErrorRecuperaSeguroMultiriesgo:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function
'*****************************************************************************
'*************************APRI 20170630 TI-ERS025*****************************
Public Function HabilitarBotonEditarEliminarRelacion(ByVal cRHCargoCod) As Boolean
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dErrorObtenerVarSistemaCargoMultiple
    lsSQL = "Exec SP_SEL_HABILITAR_BTN_EDITAR_ELIMINAR '" & cRHCargoCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(lsSQL)
    If rs!bValue = 1 Then
        HabilitarBotonEditarEliminarRelacion = True
    Else
        HabilitarBotonEditarEliminarRelacion = False
    End If
    Set rs = Nothing
    Exit Function
dErrorObtenerVarSistemaCargoMultiple:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function
'********************************END APRI*************************************
