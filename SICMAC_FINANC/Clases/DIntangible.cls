VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DIntangible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**-------------------------------------------------------------------------------------**'
'** Clase : DIntangible
'** Finalidad  : Esta Clase contiene los procedimientos relacionados
'**              a las Intangibles
'** Programador: Paolo Hector Sinti Cabrera - PASI
'** Fecha/Hora : 20140505 11:50 AM
'**-------------------------------------------------------------------------------------

Option Base 0
Option Explicit
Dim dbconec As DConecta
Dim sSql As String
'*** PASI 20140331 TI-ERS001-2014
Public Function RegistraIntangibleActiv(ByVal pnMovNroProv As Long, ByVal psMovNroIntang, ByVal psMovDesc As String, ByVal psCodOpe As String, ByVal psTipo As String, ByVal pnPerAmort As Integer, ByVal pdFechaAct As Date, ByRef psMensaje As String) As Boolean
    On Error GoTo ErrorRegistraIntangibleActiv
        
    Dim oDMOv As DMov
    Dim lnMovNro As Long
    Dim lbtrans As Boolean
    Dim lsIntCod As String
    Set oDMOv = New DMov
    
    oDMOv.BeginTrans
        lbtrans = True
        'oDMOv.InsertaMov psMovNroIntang, psCodOpe, "Activación de Intangible", gMovEstContabMovContable
        oDMOv.InsertaMov psMovNroIntang, psCodOpe, psMovDesc, gMovEstContabMovContable
        lnMovNro = oDMOv.GetnMovNro(psMovNroIntang)
        lsIntCod = ObtenerCorrelativoIntnagibleActiv(psTipo)
        
        Set dbconec = New DConecta
        dbconec.AbreConexion
        sSql = "stp_ins_ERS0012014_InsertaIntangibleActiv " & lnMovNro & ",'" & lsIntCod & "'," & pnPerAmort & ",'" & Format(pdFechaAct, "yyyyMMdd") & "'"
        dbconec.ConexionActiva.Execute sSql
        dbconec.CierraConexion
        Set dbconec = Nothing
        
        oDMOv.InsertaMovRef lnMovNro, pnMovNroProv
        RegistraIntangibleActiv = lbtrans
    oDMOv.CommitTrans
    Exit Function
ErrorRegistraIntangibleActiv:
    psMensaje = Err.Description
    If lbtrans Then
        oDMOv.RollbackTrans
        RegistraIntangibleActiv = False
    End If
    'Err.Raise Err.Number, "Registro de Activación de Intangible", Err.Description
End Function
Public Function ObtenerCorrelativoIntnagibleActiv(ByVal psTipo As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lsCodIntang As String
    
    Set dbconec = New DConecta
    dbconec.AbreConexion
    sSql = "stp_sel_ERS0012014_CorrelativoIntangibleActiv " & psTipo
    Set rs = dbconec.CargaRecordSet(sSql)
    
    If Not rs.EOF Then
        lsCodIntang = rs!cCorrelativo
    Else
        lsCodIntang = ""
    End If
    ObtenerCorrelativoIntnagibleActiv = lsCodIntang
    Set rs = Nothing
    dbconec.CierraConexion
    Set dbconec = Nothing
End Function
Public Function ListaIntangiblesActivadas() As ADODB.Recordset
    On Error GoTo ErrorListaIntangiblesActivadas
        Set dbconec = New DConecta
        If dbconec.AbreConexion Then
            sSql = "exec stp_sel_ERS0012014_ListaIntangibleActiv"
            Set ListaIntangiblesActivadas = dbconec.CargaRecordSet(sSql)
            dbconec.CierraConexion
        End If
        Set dbconec = Nothing
        Exit Function
ErrorListaIntangiblesActivadas:
    Call RaiseError(MyUnhandledError, "DIntangible:ListaIntangiblesActivadas Method")
End Function
Public Function ActualizaIntangibleActiv(ByVal pnMovNroIntang, ByVal psMovDesc As String, ByVal psIntCod As String, ByVal pnPerAmort As Integer, ByVal pdFechaAct As Date, ByRef psMensaje As String) As Boolean
    On Error GoTo ErrorActualizaIntangibleActiv
        Set dbconec = New DConecta
        dbconec.AbreConexion
        sSql = "exec stp_upd_ERS0012014_ActualizaIntangibleActiv " & pnMovNroIntang & ",'" & psMovDesc & "','" & psIntCod & "'," & pnPerAmort & ",'" & Format(pdFechaAct, "yyyyMMdd") & "'"
        dbconec.ConexionActiva.Execute sSql
        dbconec.CierraConexion
        Set dbconec = Nothing
        ActualizaIntangibleActiv = True
        Exit Function
ErrorActualizaIntangibleActiv:
    psMensaje = Err.Description
    ActualizaIntangibleActiv = False
    'Err.Raise Err.Number, "Actualizar de Activación de Intangible", Err.Description
End Function
Public Function ObtenerIntangibleActiv(ByVal pnMovNro As Long) As ADODB.Recordset
    On Error GoTo ErrorObtenerIntangibleActiv
        Set dbconec = New DConecta
        If dbconec.AbreConexion Then
            sSql = "stp_sel_ERS0012014_ObtenerIntangibleActiv " & pnMovNro
            Set ObtenerIntangibleActiv = dbconec.CargaRecordSet(sSql)
            dbconec.CierraConexion
        End If
        Set dbconec = Nothing
        Exit Function
ErrorObtenerIntangibleActiv:
    Call RaiseError(MyUnhandledError, "DIntangible:ObtenerIntangibleActiv Method")
End Function
Public Function EliminarIntangible(ByVal psMovNro As String, ByRef psMensaje As String) As Boolean
    On Error GoTo ErrorEliminarIntangible
    Dim oDMOv As DMov
    Dim lbtrans As Boolean
    Set oDMOv = New DMov
    oDMOv.BeginTrans
        lbtrans = True
        oDMOv.EliminaMov psMovNro, ""
        EliminarIntangible = True
    oDMOv.CommitTrans
    Exit Function
ErrorEliminarIntangible:
    psMensaje = Err.Description
    If lbtrans Then
        oDMOv.RollbackTrans
        EliminarIntangible = False
    End If
End Function
Public Function ListaTipoIntangible() As ADODB.Recordset
    On Error GoTo ErrorListaTipoIntangible
    Set dbconec = New DConecta
    If dbconec.AbreConexion Then
        sSql = "exec stp_sel_ERS0012014_ListaTipoIntangible"
        Set ListaTipoIntangible = dbconec.CargaRecordSet(sSql)
        dbconec.CierraConexion
    End If
    Set dbconec = Nothing
    Exit Function
ErrorListaTipoIntangible:
    Call RaiseError(MyUnhandledError, "DIntangible:ListaTipoIntangible Method")
End Function
Public Function ListaIntangibleAmortizar(ByVal psTipo As String, ByVal pdFecIni As Date, ByVal pdFecFin As Date) As ADODB.Recordset
    On Error GoTo ErrorListaIntangibleAmortizar
    Set dbconec = New DConecta
    If dbconec.AbreConexion Then
        sSql = "stp_sel_ERS0012014_ListaIntangibleAmortizar '" & psTipo & "','" & Format(pdFecIni, "yyyyMMdd") & "','" & Format(pdFecFin, "yyyyMMdd") & "'"
        Set ListaIntangibleAmortizar = dbconec.CargaRecordSet(sSql)
        dbconec.CierraConexion
    End If
    Set dbconec = Nothing
    Exit Function
ErrorListaIntangibleAmortizar:
    Call RaiseError(MyUnhandledError, "DIntangible:ListaIntangibleAmortizar Method")
End Function
Public Function ObtenerSiAmortPendiente(ByVal pnMovNro As Long, ByVal psCodIntang As String) As ADODB.Recordset
    On Error GoTo ErrorObtenerSiAmortPendiente
    Set dbconec = New DConecta
    If dbconec.AbreConexion Then
        sSql = "stp_sel_ERS0012014_ObtenerSiAmortPendiente  " & pnMovNro & ",'" & psCodIntang & "'"
        Set ObtenerSiAmortPendiente = dbconec.CargaRecordSet(sSql)
        dbconec.CierraConexion
    End If
    Set dbconec = Nothing
    Exit Function
ErrorObtenerSiAmortPendiente:
    Call RaiseError(MyUnhandledError, "DIntangible:ObtenerSiAmortPendiente Method")
End Function
Public Function IntangibleTieneAmortizacion(ByVal pnMovNro As Long, ByVal cIntgcod As String) As ADODB.Recordset
    On Error GoTo ErrorIntangibleTieneAmortizacion
    Set dbconec = New DConecta
    If dbconec.AbreConexion Then
        sSql = "stp_sel_ERS0012014_BuscaIntgAmort " & pnMovNro & ",'" & cIntgcod & "'"
        Set IntangibleTieneAmortizacion = dbconec.CargaRecordSet(sSql)
        dbconec.CierraConexion
    End If
    Set dbconec = Nothing
    Exit Function
ErrorIntangibleTieneAmortizacion:
    Call RaiseError(MyUnhandledError, "DIntangible:ObtenerSiAmortPendiente Method")
End Function
Public Function ListaAmortizacionesparaExtorno(ByVal psTipo As String, ByVal pdFecIni As Date, ByVal pdFecFin As Date) As ADODB.Recordset
    On Error GoTo ErrorListaAmortizacionesparaExtorno
    Set dbconec = New DConecta
    If dbconec.AbreConexion Then
        sSql = "stp_sel_ERS0012014_ListaIntangAmortparaExtorno '" & psTipo & "','" & Format(pdFecIni, "yyyyMMdd") & "','" & Format(pdFecFin, "yyyyMMdd") & "'"
        Set ListaAmortizacionesparaExtorno = dbconec.CargaRecordSet(sSql)
        dbconec.CierraConexion
    End If
    Set dbconec = Nothing
    Exit Function
ErrorListaAmortizacionesparaExtorno:
    Call RaiseError(MyUnhandledError, "DIntangible:ErrorListaAmortizacionesparaExtorno Method")
End Function
Public Function ObtenerAmortizacionesPosteriores(ByVal pnMovNro As Long, ByVal psCodInt As String, ByVal pdFecIni As Date, ByVal pdFecFin As Date) As ADODB.Recordset
    On Error GoTo ErrorObtenerAmortizacionesPosteriores
    Set dbconec = New DConecta
    If dbconec.AbreConexion Then
        sSql = "stp_sel_ERS0012014_ObtenerAmortPost " & pnMovNro & ",'" & psCodInt & "','" & Format(pdFecIni, "yyyyMMdd") & "','" & Format(pdFecFin, "yyyyMMdd") & "'"
        Set ObtenerAmortizacionesPosteriores = dbconec.CargaRecordSet(sSql)
        dbconec.CierraConexion
    End If
    Set dbconec = Nothing
    Exit Function
ErrorObtenerAmortizacionesPosteriores:
    Call RaiseError(MyUnhandledError, "DIntangible:ObtenerAmortizacionesPosteriores Method")
End Function
Public Function ExtornaAmortizacion(ByVal psMovNroExt As String, ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal pnMovNroAnt As Long) As Integer
    On Error GoTo ErrorExtornaAmortizacio
    Dim oDMOv As DMov
    Set oDMOv = New DMov
    Dim lbtrans As Boolean
    'Dim lnMovNro As Long
    oDMOv.BeginTrans
        lbtrans = True
        oDMOv.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpeCod, psMovDesc
    oDMOv.CommitTrans
    ExtornaAmortizacion = 0
    Exit Function
ErrorExtornaAmortizacio:
    If lbtrans Then
        lbtrans = False
        oDMOv.RollbackTrans
    End If
    Call RaiseError(MyUnhandledError, "DIntangible:ExtornaAmortizacion Method")
End Function
Public Function ListaAmortizacionesCompletas(ByVal psTipo As String) As ADODB.Recordset
    On Error GoTo ErrorListaAmortizacionesCompletas
    Set dbconec = New DConecta
    If dbconec.AbreConexion Then
        sSql = "stp_sel_ERS0012014_ListaAmortizacionesCompletas '" & psTipo & "'"
        Set ListaAmortizacionesCompletas = dbconec.CargaRecordSet(sSql)
        dbconec.CierraConexion
    End If
    Set dbconec = Nothing
    Exit Function
ErrorListaAmortizacionesCompletas:
    Call RaiseError(MyUnhandledError, "DIntangible:ErrorListaAmortizacionesCompletas Method")
End Function
Public Function ActualizaBajaIntangible(ByVal pnMovNro As Long, ByVal psCodintg As String, ByVal pdBaja As Date) As Boolean
    On Error GoTo ErrorActualizaBajaIntangible
    Set dbconec = New DConecta
    
    If dbconec.AbreConexion Then
        sSql = "stp_upd_ActualizaBajaIntangible " & pnMovNro & ",'" & psCodintg & "','" & Format(pdBaja, "yyyyMMdd") & "'"
        dbconec.ConexionActiva.Execute sSql
        ActualizaBajaIntangible = True
        dbconec.CierraConexion
    End If
    Set dbconec = Nothing
    Exit Function
    
ErrorActualizaBajaIntangible:
    Call RaiseError(MyUnhandledError, "DIntangible:ActualizaBajaIntangible Method")
End Function

