VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMBPPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private gConsPersona As String
Private gConsComunes As String
Private gConsImagenes As String
Private Conn As COMConecta.DCOMConecta
'JACA 20110318******************************************************************************************
Public Sub RegistrarParametrosBPP(ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer, ByVal iClienteNuevo As Integer, ByVal iNroOpe As Integer, ByVal dMetaCrecimiento As Double, ByVal dCMDeAceptable As Double, ByVal dCMHastaAceptacion As Double, ByVal dCMDeSinDescuento As Double, ByVal dCMHastaSinDescuento As Double, _
                                        ByVal dCMDeConDescuento1 As Double, ByVal dCMHastaConDescuento1 As Double, ByVal dCMDescConDescuento1 As Double, _
                                        ByVal dCMDeConDescuento2 As Double, ByVal dCMHastaConDescuento2 As Double, ByVal dCMDescConDescuento2 As Double, _
                                        ByVal dCMDeConDescuento3 As Double, ByVal dCMHastaConDescuento3 As Double, ByVal dCMDescConDescuento3 As Double, _
                                        ByVal dCMDeConDescuento4 As Double, ByVal dCMHastaConDescuento4 As Double, ByVal dCMDescConDescuento4 As Double, _
                                        ByVal dCMDeConDescuento5 As Double, ByVal dCMHastaConDescuento5 As Double, ByVal dCMDescConDescuento5 As Double, _
                                        ByVal dCMDeConDescuento6 As Double, ByVal dCMHastaConDescuento6 As Double, ByVal dCMDescConDescuento6 As Double, _
                                        ByVal dCMDeConDescuento7 As Double, ByVal dCMHastaConDescuento7 As Double, ByVal dCMDescConDescuento7 As Double, _
                                        ByVal dCMDeConDescuento8 As Double, ByVal dCMHastaConDescuento8 As Double, ByVal dCMDescConDescuento8 As Double, _
                                        ByVal dCMDeConDescuento9 As Double, ByVal dCMHastaConDescuento9 As Double, ByVal dCMDescConDescuento9 As Double, _
                                        ByVal dCMDeConDescuento10 As Double, ByVal dCMHastaConDescuento10 As Double, ByVal dCMDescConDescuento10 As Double, _
                                        ByVal dCMDeConDescuento11 As Double, ByVal dCMHastaConDescuento11 As Double, ByVal dCMDescConDescuento11 As Double, _
                                        ByVal dCMDeConDescuento12 As Double, ByVal dCMHastaConDescuento12 As Double, ByVal dCMDescConDescuento12 As Double, _
                                        ByVal sMovNro As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_ParametrosBPP  " & iTipoCarteraId & ", " & iTipoClasificacionId & ", " & iClienteNuevo & ", " & iNroOpe & ", " & dMetaCrecimiento & ", " & dCMDeAceptable & ", " & dCMHastaAceptacion & ", " & dCMDeSinDescuento & ", " & dCMHastaSinDescuento & ", " & dCMDeConDescuento1 & ", " & dCMHastaConDescuento1 & ", " & dCMDescConDescuento1 & ""
    lsSQL = lsSQL & " , " & dCMDeConDescuento2 & ", " & dCMHastaConDescuento2 & ", " & dCMDescConDescuento2 & ""
    lsSQL = lsSQL & " , " & dCMDeConDescuento3 & ", " & dCMHastaConDescuento3 & ", " & dCMDescConDescuento3 & ""
    lsSQL = lsSQL & " , " & dCMDeConDescuento4 & ", " & dCMHastaConDescuento4 & ", " & dCMDescConDescuento4 & ""
    lsSQL = lsSQL & " , " & dCMDeConDescuento5 & ", " & dCMHastaConDescuento5 & ", " & dCMDescConDescuento5 & ""
    lsSQL = lsSQL & " , " & dCMDeConDescuento6 & ", " & dCMHastaConDescuento6 & ", " & dCMDescConDescuento6 & ""
    lsSQL = lsSQL & " , " & dCMDeConDescuento7 & ", " & dCMHastaConDescuento7 & ", " & dCMDescConDescuento7 & ""
    lsSQL = lsSQL & " , " & dCMDeConDescuento8 & ", " & dCMHastaConDescuento8 & ", " & dCMDescConDescuento8 & ""
    lsSQL = lsSQL & " , " & dCMDeConDescuento9 & ", " & dCMHastaConDescuento9 & ", " & dCMDescConDescuento9 & ""
    lsSQL = lsSQL & " , " & dCMDeConDescuento10 & ", " & dCMHastaConDescuento10 & ", " & dCMDescConDescuento10 & ""
    lsSQL = lsSQL & " , " & dCMDeConDescuento11 & ", " & dCMHastaConDescuento11 & ", " & dCMDescConDescuento11 & ""
    lsSQL = lsSQL & " , " & dCMDeConDescuento12 & ", " & dCMHastaConDescuento12 & ", " & dCMDescConDescuento12 & ""
    lsSQL = lsSQL & " , '" & sMovNro & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub

Public Function CargarAgencias() As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_Agencias_BPPR"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarAgencias = lrDatos
    Set lrDatos = Nothing
End Function
'END JACA*********************************************************************

Public Function CargarComite() As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_Comite"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarComite = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarDatosMeta(ByVal sAgeCod As String, ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer) As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_CargarDatosMeta '" & sAgeCod & "', " & iTipoCarteraId & ", " & iTipoClasificacionId & ""
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarDatosMeta = lrDatos
    Set lrDatos = Nothing
End Function

Public Function RecuperaResumenSaldosCarteraPorAnalistaConsolXCierre(psServerCons As String, ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Integer, ByVal P1I As Integer, ByVal P1F As Integer, _
    ByVal P2I As Integer, ByVal P2F As Integer, ByVal P3I As Integer, ByVal P3F As Integer, ByVal P4F As Integer, _
    ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, Optional ByVal pnTipCam As Double, Optional ByVal pbCondicBN As Boolean = False, _
    Optional dFCierre As Date) As ADODB.Recordset
    
Dim sSql As String
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadCondicionBN As String
Dim sCadProd As String
Dim oConecta As COMConecta.DCOMConecta

    sCadProd = ""
    If UBound(pMatProd) > 0 Then
        For i = 0 To UBound(pMatProd) - 1
            sCadProd = sCadProd & pMatProd(i) & ","
        Next i
        sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)
    End If

    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
    
    sCadCondicion = ""
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1)

    If pbCondicBN Then
        sCadCondicionBN = "1" ''INCLUYE LA TABLA ColocConvBcoNac EN EL STP
    Else
        sCadCondicionBN = "2" ''NO INCLUYE LA TABLA ColocConvBcoNac EN EL STP
    End If
    
    sSql = "exec stp_sel_RecuperaResumenSaldosCarteraPorAnalistaConsolXCierre " & pnTipCam & "," & P1I & "," & P1F & "," & P2I & "," & P2F & "," & P3I & "," & P3F & "," & P4F & ",'" & sCadCondicion & "','" & sCadProd & "','" & sCadAge & "','" & sCadCondicionBN & "','" & Format(dFCierre, "yyyymmdd") & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaResumenSaldosCarteraPorAnalistaConsolXCierre = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function DarCargoAnalistas() As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_sel_DarCargoAnalista"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set DarCargoAnalistas = lrDatos
    Set lrDatos = Nothing
End Function
Public Function CargarTipoCarteraReporte() As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
        lsSQL = "exec stp_sel_TipoCartera_Reporte"
        
        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set CargarTipoCarteraReporte = lrDatos
        Set lrDatos = Nothing
End Function
'JACA 20110318*****************************************************************
'Public Function CargarAnalistasXAgencia(ByVal sAgeCod As String, ByVal nConsSisValor As String, ByVal nTpoCambio As Double, ByVal nTipoCartera As Integer) As ADODB.Recordset
Public Function CargarAnalistasXAgencia(ByVal sAgeCod As String, ByVal nTpoCambio As Double, ByVal nTipoCartera As Integer, ByVal nTipoCarteraClasificacion As Integer, ByVal dFecha As Date) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    If nTipoCartera = 2 Then 'JACA 20110606
        lsSQL = "exec stp_sel_AnalistasXAgenciaConvenio'" & sAgeCod & "', " & nTpoCambio & ", " & nTipoCartera & ", " & nTipoCarteraClasificacion & ",'" & Format(dFecha, "yyyymmdd") & "'"
    Else
        lsSQL = "exec stp_sel_AnalistasXAgenciaMinorista_NoMinorista'" & sAgeCod & "', " & nTpoCambio & ", " & nTipoCartera & ", " & nTipoCarteraClasificacion & ",'" & Format(dFecha, "yyyymmdd") & "'"
    End If
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarAnalistasXAgencia = lrDatos
    Set lrDatos = Nothing
End Function
Public Function CargarCordinadorXAgencia(ByVal sAgeCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_sel_CordinadorXAgencia'" & sAgeCod & "'"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarCordinadorXAgencia = lrDatos
    Set lrDatos = Nothing
End Function
Public Sub RegistrarAnalistaMetaBPP(ByVal sAgeCodAct As String, ByVal sCodAna As String, ByVal sUser As String, ByVal iTpoCar As Integer, ByVal iTpoCarClas As Integer, ByVal iCred As Integer, ByVal iClie As Integer, ByVal iOpe As Integer, ByVal iSaldo As Currency, ByVal iMora As Double, ByVal iTasa As Double, ByVal sCodCord As String, ByVal iMetaClie As Integer, ByVal iMetaOpe As Integer, ByVal iMetaSaldo As Long, ByVal sFechaMeta As String, ByVal sMovNro As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_BPPAnlistaMeta '" & sAgeCodAct & "','" & sCodAna & "','" & sUser & "'," & iTpoCar & "," & iTpoCarClas & ", " & iCred & ", " & iClie & "," & iOpe & "," & iSaldo & "," & iMora & "," & iTasa & ",'" & sCodCord & "'," & iMetaClie & "," & iMetaOpe & "," & iMetaSaldo & ", '" & sFechaMeta & "', '" & sMovNro & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub

Public Sub modificarParametrosBPP(ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_upd_ParametrosBPP " & iTipoCarteraId & "," & iTipoClasificacionId & ""
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
Public Function CargarTipoCartera() As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
        lsSQL = "exec stp_sel_TipoCartera"
        
        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set CargarTipoCartera = lrDatos
        Set lrDatos = Nothing
End Function

Public Function CargarTipoClasificacion(ByVal iTipoCarteraId As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
        lsSQL = "exec stp_sel_TipoClasificacion " & iTipoCarteraId & ""
        
        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set CargarTipoClasificacion = lrDatos
        Set lrDatos = Nothing
End Function
Public Function CargarParametrosTipoClasificacion(ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
        lsSQL = "exec stp_sel_TipoParametrosClasificacion " & iTipoCarteraId & "," & iTipoClasificacionId & ""
        
        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set CargarParametrosTipoClasificacion = lrDatos
        Set lrDatos = Nothing
End Function
Public Function ObtenerFechaUlitmoMeta(ByVal sAgeCod As String, ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
        lsSQL = "exec stp_sel_ObtenerFechaUltimoMeta '" & sAgeCod & "'," & iTipoCarteraId & "," & iTipoClasificacionId & ""
        
        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set ObtenerFechaUlitmoMeta = lrDatos
        Set lrDatos = Nothing
End Function
Public Function ObtenerListaAnalistaMetas(ByVal sAgeCod As String, ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer, ByVal sFechaMeta As String) As ADODB.Recordset
    
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
        lsSQL = "exec stp_sel_ListaAnalistaMetas '" & sAgeCod & "'," & iTipoCarteraId & "," & iTipoClasificacionId & ",'" & sFechaMeta & "'"
        
        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set ObtenerListaAnalistaMetas = lrDatos
        Set lrDatos = Nothing
End Function
Public Function ObtenerListaAnalistaMetasAnteriores(ByVal sAgeCod As String, ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer, ByVal sFechaMeta As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
        lsSQL = "exec stp_sel_ListaAnalistaMetasAnt '" & sAgeCod & "'," & iTipoCarteraId & "," & iTipoClasificacionId & ",'" & sFechaMeta & "'"
        
        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set ObtenerListaAnalistaMetasAnteriores = lrDatos
        Set lrDatos = Nothing
End Function
Public Sub modificarEstadoAnalistaMeta(ByVal sAgeCod As String, ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer, ByVal sFechaMeta As String, ByVal sMovNro As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_upd_ListaAnalistaMetas '" & sAgeCod & "'," & iTipoCarteraId & "," & iTipoClasificacionId & ",'" & sFechaMeta & "','" & sMovNro & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
Public Function getBonoAnalistas(ByVal sAgeCod As String, ByVal nTpoCambio As Double, ByVal nTipoCarteraId As Integer, ByVal nTipoClasificacionId As Integer, ByVal dFecha As Date) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

        If nTipoCarteraId <> 2 Then
            lsSQL = "exec stp_sel_BonificacionBPP_Mino_NOMino_Consol '" & sAgeCod & "'," & nTpoCambio & "," & nTipoCarteraId & "," & nTipoClasificacionId & ",'" & Format(dFecha, "yyyymmdd") & "'"
        Else
            lsSQL = "exec stp_sel_BonificacionBPP_Convenio_Consol  '" & sAgeCod & "'," & nTpoCambio & "," & nTipoCarteraId & "," & nTipoClasificacionId & ",'" & Format(dFecha, "yyyymmdd") & "'"
        End If

        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set getBonoAnalistas = lrDatos
        Set lrDatos = Nothing
End Function

Public Sub insertarBonoAnalistas(ByVal psAgeCodAct As String, ByVal psCodAna As String, ByVal pnTpoCar As Integer, ByVal pnTpoCarClas As Integer, _
                                ByVal pnCredAnt As Integer, ByVal pnClieAnt As Integer, ByVal pnSaldoAnt As Currency, _
                                ByVal pnCredTransf As Integer, ByVal pnClieTransf As Integer, ByVal pnSaldoTransf As Currency, _
                                ByVal pnMetaClie As Integer, ByVal pnMetaOpe As Integer, ByVal pnMetaSald As Currency, ByVal psCoord As String, _
                                ByVal pnCred As Integer, ByVal pnClie As Integer, ByVal pnSaldo As Currency, ByVal pnClieNue As Integer, ByVal pnClieRecu As Integer, _
                                ByVal pnOpe As Integer, ByVal pnMora As Double, ByVal pnTasaPond As Double, _
                                ByVal pnBonoClieNue As Currency, ByVal pnBonoClieRecu As Currency, ByVal pnBonoSaldo As Currency, ByVal pnBonoOpe As Currency, _
                                ByVal pnBonoMora As Currency, ByVal pnBonoTasa As Currency, ByVal pnBonoTotal As Currency, _
                                ByVal pnFecha As Date, ByVal sMovNro As String)

    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_BonoAnalistas '" & psAgeCodAct & "','" & psCodAna & "'," & pnTpoCar & "," & pnTpoCarClas & ","
    lsSQL = lsSQL & pnCredAnt & "," & pnClieAnt & "," & pnSaldoAnt & ","
    lsSQL = lsSQL & pnCredTransf & "," & pnClieTransf & "," & pnSaldoTransf & ","
    lsSQL = lsSQL & pnMetaClie & "," & pnMetaOpe & "," & pnMetaSald & ",'" & psCoord & "',"
    lsSQL = lsSQL & pnCred & "," & pnClie & "," & pnSaldo & "," & pnClieNue & "," & pnClieRecu & ","
    lsSQL = lsSQL & pnOpe & "," & pnMora & "," & pnTasaPond & ","
    lsSQL = lsSQL & pnBonoClieNue & "," & pnBonoClieRecu & "," & pnBonoSaldo & "," & pnBonoOpe & ","
    lsSQL = lsSQL & pnBonoMora & "," & pnBonoTasa & "," & pnBonoTotal & ","
    lsSQL = lsSQL & "'" & Format(pnFecha, "yyyymmdd") & "','" & sMovNro & "'"

    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub
Public Function getBonoAnalistasConsol(ByVal sAgeCod As String, ByVal nTipoCarteraId As Integer, ByVal nTipoClasificacionId As Integer, ByVal dFecha As Date) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_sel_BonificacionBPP_Consol '" & sAgeCod & "',0," & nTipoCarteraId & "," & nTipoClasificacionId & ",'" & Format(dFecha, "yyyymmdd") & "'"

    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set getBonoAnalistasConsol = lrDatos
    Set lrDatos = Nothing
End Function
'END JACA*****************************************************************

'JUEZ 20121010 *************************************************************************
Public Function RecuperaCredTiposCartera(Optional psFiltroId As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCredTiposCartera
    
    sSql = "exec stp_sel_RecuperaCredBPPTiposCartera '" & psFiltroId & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCredTiposCartera = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCredTiposCartera:
    Err.Raise Err.Number, "Recuperar Datos Tipos de Cartera", Err.Description
End Function

Public Sub dInsertaCredTipoCartera(ByVal psCartera As String, ByVal pnCredCorpo As Integer, ByVal pnGrandeEmp As Integer, _
                                    ByVal pnMedianaEmp As Integer, ByVal pnPequenaEmp As Integer, ByVal pnMicro As Integer, _
                                    ByVal pnConsumoRev As Integer, ByVal pnConsumoNoRev As Integer, ByVal pnHipotecario As Integer, _
                                    Optional ByVal pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_GrabarBPPTipoCartera '" & psCartera & "'," & pnCredCorpo & "," & pnGrandeEmp & "," & pnMedianaEmp & "," & _
            pnPequenaEmp & "," & pnMicro & "," & pnConsumoRev & "," & pnConsumoNoRev & "," & pnHipotecario
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dActualizaCredTipoCartera(ByVal psIdTipo As String, ByVal psCartera As String, ByVal pnCredCorpo As Integer, _
                                    ByVal pnGrandeEmp As Integer, ByVal pnMedianaEmp As Integer, ByVal pnPequenaEmp As Integer, _
                                    ByVal pnMicro As Integer, ByVal pnConsumoRev As Integer, ByVal pnConsumoNoRev As Integer, _
                                    ByVal pnHipotecario As Integer, Optional ByVal pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_upd_GrabarBPPTipoCartera '" & psIdTipo & "','" & psCartera & "'," & pnCredCorpo & "," & pnGrandeEmp & "," & _
            pnMedianaEmp & "," & pnPequenaEmp & "," & pnMicro & "," & pnConsumoRev & "," & pnConsumoNoRev & "," & pnHipotecario

    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dEliminaCredTipoCartera(ByVal psIdTipo As String, Optional ByVal pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_del_EliminaBPPTipoCartera '" & psIdTipo & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function ListarTiposCartera() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorListarTiposCartera
    
    sSql = "exec stp_sel_ListaCredBPPTiposCartera"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ListarTiposCartera = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorListarTiposCartera:
    Err.Raise Err.Number, "Listar Tipos de Cartera", Err.Description
End Function

Public Function ListarCategoriaAnalistas() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorListarCategoriaAnalistas
    
    sSql = "exec stp_sel_ListaCredBPPCategoriasAnalistas"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ListarCategoriaAnalistas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorListarCategoriaAnalistas:
    Err.Raise Err.Number, "Listar Categorias de Analista", Err.Description
End Function

Public Sub dInsertaParametroDatos(ByVal psTipoCartera As String, ByVal psCatAnalista As String, ByVal pnMinCartera As Double, ByVal pnMaxCartera As Double)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_GrabarParametroDatos '" & psTipoCartera & "','" & psCatAnalista & "'," & pnMinCartera & "," & pnMaxCartera
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function RecuperaParametro(ByVal psTipoCartera As String, ByVal psCatAnalista As String, ByVal pnMinCartera As Double, ByVal pnMaxCartera As Double) As String
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaParametro
    
    sSql = "exec stp_sel_RecuperaParametro '" & psTipoCartera & "','" & psCatAnalista & "'," & pnMinCartera & "," & pnMaxCartera
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.EOF Then
        RecuperaParametro = rs!cIdParametro
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaParametro:
    Err.Raise Err.Number, "Recuperado ID Parametro", Err.Description
End Function

Public Sub dInsertaParametrosLista(ByVal psIdParam As String, ByVal psValor As String, ByVal pnTipo As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_GrabarParametrosListas '" & psIdParam & "','" & psValor & "'," & pnTipo
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dInsertaParametrosIndicadoresCartera(ByVal psIdParam As String, ByVal pnDiaDesde As Long, ByVal pnDiaHasta As Long, ByVal pnJudicial As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_GrabarParametrosIndicadoresCartera '" & psIdParam & "'," & pnDiaDesde & "," & pnDiaHasta & "," & pnJudicial
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dActualizaParametroDatos(ByVal psIdParam As String, ByVal psTipoCartera As String, ByVal psCatAnalista As String, _
                                    ByVal pnMinCartera As Double, ByVal pnMaxCartera As Double)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_ActualizaParametroDatos '" & psIdParam & "','" & psTipoCartera & "','" & psCatAnalista & "'," & _
                                                    pnMinCartera & "," & pnMaxCartera
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dEliminaParametrosListas(ByVal psIdParam As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_del_EliminaParametrosListas '" & psIdParam & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function RecuperaParametroDatos(Optional ByVal psIdParametro As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaParametroDatos
    
    sSql = "exec stp_sel_RecuperaParametroDatos '" & psIdParametro & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaParametroDatos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaParametroDatos:
    Err.Raise Err.Number, "Recupera Datos Parametros", Err.Description
End Function

Public Function RecuperaParametrosLista(ByVal psIdParametro As String, ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaParametrosLista
    
    sSql = "exec stp_sel_RecuperaParametrosLista '" & psIdParametro & "'," & pnTipo
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaParametrosLista = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaParametrosLista:
    Err.Raise Err.Number, "Recupera Listas Parametros", Err.Description
End Function

Public Sub dEliminaParametroDatos(ByVal psIdParam As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_del_EliminaParametroDatos '" & psIdParam & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function VerificaSiExisteParametro(ByVal psTipoCartera As String, ByVal psCateAnalista As String, ByVal psIdParam As String) As Boolean
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    VerificaSiExisteParametro = False
    lsSQL = "exec stp_sel_VerificaSiExisteParametro '" & psTipoCartera & "','" & psCateAnalista & "','" & psIdParam & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        VerificaSiExisteParametro = True
    End If
    Set lrDatos = Nothing
End Function

Public Function VerificaTamanoCartera(ByVal psIdParam As String, ByVal psTipoCartera As String, ByVal pnMinCart As Double, _
                                        ByVal pnMaxCart As Double, ByRef cCatPertenece As String) As Boolean
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    VerificaTamanoCartera = False
    lsSQL = "exec stp_sel_VerificaTamanoCartera '" & psIdParam & "','" & psTipoCartera & "'," & pnMinCart & "," & pnMaxCart
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        VerificaTamanoCartera = True
        cCatPertenece = lrDatos!cCatAnalistaDesc
    End If
    Set lrDatos = Nothing
End Function

'**END JUEZ *************************************************************************

'WIOR 20130524 **********************************************************************
Public Sub OpeSubProductosCredNiveles(ByVal pnOperacion As Integer, Optional ByVal psSubProducto As String = "", Optional ByVal pnNivel As Integer = 0, Optional pnTipo As Integer = 1)
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

Select Case pnOperacion
    Case 1: lsSQL = "INSERT INTO CredBPPSubProductosCredNiveles(cTpoSubProd,nNivel,nTipo) VALUES ('" & psSubProducto & "'," & pnNivel & "," & pnTipo & ")"
    Case 2: lsSQL = "UPDATE CredBPPSubProductosCredNiveles SET nNivel=" & pnNivel & " WHERE cTpoSubProd='" & psSubProducto & "' and nTipo=" & pnTipo
    Case 3: lsSQL = "DELETE FROM CredBPPSubProductosCredNiveles WHERE nTipo=" & pnTipo & IIf(pnNivel = 0, "", " and nNivel=" & pnNivel)
End Select

Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub
Public Function ObtenerNivelesXSubProdCred(ByVal psSubProducto As String, Optional pnTipo As Integer = 1) As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "stp_sel_BPPObtenerNivelesXSubProd '" & psSubProducto & "'," & pnTipo

Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
Set ObtenerNivelesXSubProdCred = oConec.CargaRecordSet(lsSQL)
oConec.CierraConexion
End Function
Public Function ObtenerSubProTpoCredXNivel(ByVal pnNivel As Integer, Optional pnTipo As Integer = 1) As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "exec stp_sel_BPPObtenerSubProdTpoCredXNivel " & pnNivel & "," & pnTipo

Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
Set ObtenerSubProTpoCredXNivel = oConec.CargaRecordSet(lsSQL)
oConec.CierraConexion
End Function

Public Function ObtenerCatOParamCabXNivel(ByVal pnOperacion As Integer, ByVal pnNivel As String) As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

Select Case pnOperacion
    Case 1: lsSQL = "SELECT cCategoria FROM CredBPPNivelesCat WHERE nNivel =" & pnNivel & " ORDER BY cCategoria DESC "
    Case 2: lsSQL = "SELECT bSaldoCartera,bNumClientes FROM CredBPPNivelesParamCab WHERE nNivel =" & pnNivel
End Select

Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
Set ObtenerCatOParamCabXNivel = oConec.CargaRecordSet(lsSQL)
oConec.CierraConexion
End Function
Public Function OpeCatAnalistaXNivel(ByVal pnOperacion As Integer, ByVal pnNivel As Integer, _
                                    Optional ByVal psCategoria As String, Optional ByVal pnSalMax As Double, _
                                    Optional ByVal pnSalMin As Double, Optional ByVal pnNumMax As Double, _
                                    Optional ByVal pnNumMin As Double) As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta


Select Case pnOperacion
    Case 1: lsSQL = "stp_ins_BPPRegistrarCatAnalista " & pnNivel & ",'" & psCategoria & "'," & pnSalMax & "," & pnSalMin & "," & pnNumMax & "," & pnNumMin
    Case 2: lsSQL = "DELETE FROM CredBPPCatAnalista WHERE nNivel=" & pnNivel
End Select

Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Function
Public Function ObtenerCatAnalistaXNivel(ByVal pnNivel As String, ByVal psCategoria As String) As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "SELECT nNivel,cCategoria,nSalCarteraMin,nSalCarteraMax,nNumClientesMin,nNumClientesMax  " & _
        "FROM CredBPPCatAnalista WHERE nNivel =" & pnNivel & " and cCategoria='" & psCategoria & "'"

Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
Set ObtenerCatAnalistaXNivel = oConec.CargaRecordSet(lsSQL)
oConec.CierraConexion
End Function
Public Function ObtenerAnalistaXAgencia(Optional ByVal psAgencia As String = "%", Optional ByVal pUser As String = "%") As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "exec stp_sel_BPPObtenerMorasXAnalista '" & psAgencia & "','" & pUser & "'"

Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
Set ObtenerAnalistaXAgencia = oConec.CargaRecordSet(lsSQL)
oConec.CierraConexion
End Function
Public Function ObtenerDetSaldoXAnalista(ByVal psPersCod As String, ByVal psPeriodo As String) As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "exec stp_sel_BPPMoraBaseDetalleAnalista '" & psPersCod & "','" & psPeriodo & "'"

Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
Set ObtenerDetSaldoXAnalista = oConec.CargaRecordSet(lsSQL)
oConec.CierraConexion
End Function
Public Sub OpeMoraBase(ByVal psPersCod As String, ByVal dFecMoraBase As Date, ByVal pnPorcMoraBase As Double, ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal pnCambioMora As Integer)
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "Exec stp_upd_CredBPPMoraBase '" & psPersCod & "','" & Format(dFecMoraBase, "yyyymmdd") & "'," & pnPorcMoraBase & "," & pnAnio & "," & pnMes & "," & pnCambioMora

Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub
Public Sub EliminaParametrosCumplimiento(ByVal pnNivel As Integer, _
                                        ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, _
                                        ByVal pnTipoParam As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_del_ParametroCumplimiento " & pnNivel & "," & pnMes & "," & pnAnio & ",'" & psCodAge & "'," & pnTipoParam

            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub
Public Sub OpeCarteraTpoCred(ByVal pnOperacion As Integer, Optional ByVal pnNivel As Integer = 0, Optional pnMin As Double = 0, Optional pnMax As Double = 0, Optional ByVal pnTipo As Integer = 1)
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta


Select Case pnOperacion
    Case 1: lsSQL = "INSERT INTO CredBPPTpoCredCartera(nNivel,nMin,nMax,nTipo) VALUES (" & pnNivel & "," & pnMin & "," & pnMax & "," & pnTipo & ")"
    Case 2: lsSQL = "DELETE FROM CredBPPTpoCredCartera WHERE nNivel=" & pnNivel & " and nTipo=" & pnTipo
End Select

Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub

Public Function ObtenerCarteraTpoCred(Optional ByVal pnNivel As Integer = 0, Optional ByVal pnTipo As Integer = 1) As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "select nNivel,nMin,nMax  FROM CredBPPTpoCredCartera WHERE nNivel=" & pnNivel & " and nTipo=" & pnTipo

Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
Set ObtenerCarteraTpoCred = oConec.CargaRecordSet(lsSQL)
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function DevolverFactoresCumplimiento(ByVal pnNivel As Integer, ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorDevolverFactoresCumplimiento
    
    sSql = "exec spt_sel_BPPDevuelveFactoresXNivel " & pnNivel & ", " & pnTipo
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverFactoresCumplimiento = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorDevolverFactoresCumplimiento:
    Err.Raise Err.Number, "Devolver Factores de Cumplimiento", Err.Description
End Function

Public Sub InsertaParametrosCumplimiento(ByVal pnNivel As Integer, ByVal psCodFactor As String, ByVal pnCategoriaA As Double, _
                                        ByVal pnCategoriaB As Double, ByVal pnCategoriaC As Double, ByVal pnCategoriaD As Double, _
                                        ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, ByVal psUserReg As String, _
                                        ByVal pdFecReg As String, ByVal pnTipoParam As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "Exec stp_ins_InsertaParametroCumplimiento " & pnNivel & ",'" & psCodFactor & "'," & pnCategoriaA & "," & pnCategoriaB & "," & _
            pnCategoriaC & "," & pnCategoriaD & "," & pnMes & "," & pnAnio & ",'" & psCodAge & "','" & psUserReg & "','" & pdFecReg & "'," & pnTipoParam

            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub
Public Function DevolverParametrosCumplimiento(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, _
                                                ByVal pnTipoParam As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorDevolverParametrosCumplimiento
    
    sSql = "exec stp_sel_DevuelveParametrosCumplimiento " & pnNivel & ", " & pnMes & ", " & pnAnio & ", '" & psCodAge & "'," & pnTipoParam
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverParametrosCumplimiento = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorDevolverParametrosCumplimiento:
    Err.Raise Err.Number, "Devolver Factores de Cumplimiento", Err.Description
End Function

Public Sub InsertaRangosMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, ByVal pnMora8a30 As Double, _
                             ByVal pnMoraMaya30 As Double, ByVal psUserReg As String, ByVal pdFecReg As String, _
                             ByVal pnNivel As Integer, ByVal psCategoria As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_InsertaRangosMora " & pnMes & "," & pnAnio & ",'" & psCodAge & "'," & pnMora8a30 & "," & pnMoraMaya30 & ",'" & psUserReg & "','" & pdFecReg & "'," & pnNivel & ",'" & psCategoria & "'"

            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub EliminaRangosMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, ByVal pnNivel As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta

    lsSQL = "exec stp_del_RangosMora " & pnMes & "," & pnAnio & ",'" & psCodAge & "'," & pnNivel

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar lsSQL
    oConec.CierraConexion
    Set oConec = Nothing
End Sub

Public Function DevolverRangosMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, ByVal pnNivel As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorDevolverRangosMora
    
    sSql = "exec stp_sel_CargaRangosMora " & pnMes & ", " & pnAnio & ",'" & psCodAge & "'," & pnNivel
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverRangosMora = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorDevolverRangosMora:
    Err.Raise Err.Number, "Devolver Rangos de Mora", Err.Description
End Function
Public Function DevolverFactorRend(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorDevolverFactorRend
    
    sSql = "exec stp_sel_CredCredFacRendBPP " & pnMes & ", " & pnAnio & ",'" & psCodAge & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverFactorRend = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorDevolverFactorRend:
    Err.Raise Err.Number, "DevolverFactorRend", Err.Description
End Function

Public Function DevolverFechasAMigrarXParam(ByVal pnTipo As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorDevolverFechasAMigrarXParam
    
    sSql = "exec stp_sel_CredBPPFecMigrarXParam " & pnTipo & "," & pnMes & "," & pnAnio & ",'" & psCodAge & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverFechasAMigrarXParam = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorDevolverFechasAMigrarXParam:
    Err.Raise Err.Number, "DevolverFechasAMigrarXParam", Err.Description
End Function

Public Sub EliminaFactorRend(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta

    lsSQL = "exec stp_del_CredCredFacRendBPP " & pnMes & "," & pnAnio & ",'" & psCodAge & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar lsSQL
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
Public Sub MigrarDatosXParam(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, ByVal pnTipo As Integer, ByVal pnMesAct As Integer, ByVal pnAnioAct As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta

    lsSQL = "exec stp_ins_CredBPPMigrarXParam " & pnMes & "," & pnAnio & ",'" & psCodAge & "'," & pnTipo & "," & pnMesAct & "," & pnAnioAct

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar lsSQL
    oConec.CierraConexion
    Set oConec = Nothing
End Sub

Public Sub InsertaFactorRend(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCategoria As String, ByVal pnNivel1 As Double, _
                             ByVal pnNivel2 As Double, ByVal pnNivel3 As Double, ByVal psUserReg As String, ByVal pdFecReg As String, _
                             ByVal psCodAge As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_ins_InsertaCredFacRendBPP " & pnMes & "," & pnAnio & ",'" & psCategoria & "'," & pnNivel1 & "," & pnNivel2 & "," & pnNivel3 & ",'" & psUserReg & "','" & pdFecReg & "','" & psCodAge & "'"
      
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub
Public Function DevolverFactorOpe(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorDevolverFactorOpe
    
    sSql = "exec stp_sel_CredFactOpeBPP " & pnMes & ", " & pnAnio & ",'" & psCodAge & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverFactorOpe = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorDevolverFactorOpe:
    Err.Raise Err.Number, "ErrorDevolverFactorOpe", Err.Description
End Function

Public Sub EliminaFactorOpe(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta

    lsSQL = "exec stp_del_CredFactOpeBPP " & pnMes & "," & pnAnio & ",'" & psCodAge & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar lsSQL
    oConec.CierraConexion
    Set oConec = Nothing
End Sub

Public Sub InsertaFactorOpe(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal pnQuincena1 As Double, _
                             ByVal pnQuincena2 As Double, ByVal psUserReg As String, ByVal pdFecReg As String, _
                             ByVal psCodAge As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_InsertaCredFactOpeBPP " & pnMes & "," & pnAnio & "," & pnQuincena1 & "," & pnQuincena2 & ",'" & psUserReg & "','" & pdFecReg & "','" & psCodAge & "'"
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function DevolverCategorias() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorDevolverCategorias
    
    sSql = "exec stp_sel_DevolverCategoriasBPP"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverCategorias = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorDevolverCategorias:
    Err.Raise Err.Number, "Devolver Categorias", Err.Description
End Function

Public Sub InsertaPesoyTopes(ByVal pnNivel As Integer, ByVal psCodFactor As String, ByVal pnPesoCateA As Double, ByVal pnTopeCateA As Double, _
                             ByVal pnPesoCateB As Double, ByVal pnTopeCateB As Double, ByVal pnPesoCateC As Double, ByVal pnTopeCateC As Double, _
                             ByVal pnPesoCateD As Double, ByVal pnTopeCateD As Double, ByVal pnMes As Integer, ByVal pnAnio As Integer, _
                             ByVal psCodAge As String, ByVal psUserReg As String, ByVal pdFecReg As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_InsertaPesosTopesBPP " & pnNivel & ",'" & psCodFactor & "'," & pnPesoCateA & "," & pnTopeCateA & "," & _
            pnPesoCateB & "," & pnTopeCateB & "," & pnPesoCateC & "," & pnTopeCateC & "," & pnPesoCateD & "," & pnTopeCateD & ", " & _
            pnMes & "," & pnAnio & ",'" & psCodAge & "','" & psUserReg & "','" & pdFecReg & "'"

    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub
Public Sub EliminaPesoyTopes(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_del_PesosTopesBPP " & pnNivel & "," & pnMes & "," & pnAnio & ",'" & psCodAge & "'"

    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function DevolverPesoyTopes(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorDevolverPesoyTopes
    
    sSql = "exec stp_sel_DevuelvePesoyTopes " & pnNivel & ", " & pnMes & ", " & pnAnio & ", '" & psCodAge & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverPesoyTopes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorDevolverPesoyTopes:
    Err.Raise Err.Number, "Devolver Pesos y Topes de BPP", Err.Description
End Function

Public Sub InsertaPenalidadXIncreMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, _
                                      ByVal pnDesde As Currency, ByVal pnHasta As Currency, ByVal pnDscto As Currency, _
                                      ByVal psUserReg As String, ByVal pdFecReg As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_InsertaPenalidadXIncreMora " & pnMes & "," & pnAnio & ",'" & psCodAge & "'," & pnDesde & "," & pnHasta & "," & pnDscto & _
                                                    ",'" & psUserReg & "','" & pdFecReg & "'"

            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub EliminaPenalidadXIncreMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta

    lsSQL = "exec stp_del_PenalidadXIncreMora " & pnMes & "," & pnAnio & ",'" & psCodAge & "'"

            
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar lsSQL
    Set oConec = Nothing
End Sub
Public Function DevolverPenalidadXIncreMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorDevolverPenalidadXIncreMora
    
    sSql = "exec stp_sel_PenalidadXIncreMora  " & pnMes & ", " & pnAnio & ", '" & psCodAge & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverPenalidadXIncreMora = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorDevolverPenalidadXIncreMora:
    Err.Raise Err.Number, "Devolver PenalidadXIncreMora", Err.Description
End Function
Public Function ObtenerMetasAgencia(ByVal pdFecha As Date, ByVal psCodAge As String, ByVal pnNivel As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerMetasAgencia
    
    sSql = "exec stp_sel_MetasXAgencias  '" & Format(pdFecha, "yyyymmdd") & "','" & psCodAge & "', " & pnNivel
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerMetasAgencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtenerMetasAgencia:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Function
Public Sub OpeMetasAgencia(ByVal psPersCod As String, ByVal pnSaldoAg As Double, ByVal pnNCliAg As Double, _
                        ByVal pnNOpeAg As Double, ByVal pnMoraAg As Double, ByVal pnMes As Long, ByVal pnAnio As Long, _
                        ByVal pcUserReg As String, ByVal pdFecha As Date, _
                        ByVal pnNivel As Integer)
                        
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "Exec stp_ins_upd_CredMetasAgenciasBPP '" & psPersCod & "'," & pnSaldoAg & "," & pnNCliAg & "," & _
            pnNOpeAg & "," & pnMoraAg & "," & pnMes & "," & pnAnio & ",'" & pcUserReg & "','" & Format(pdFecha, "yyyymmdd hh:mm:s") & "'," & _
            pnNivel


Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub

Public Function ObtenerComiteAnaCord(ByVal psCodAge As String, ByVal pnMes As Integer, ByVal pnAnio As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerComiteAnaCord
    
    sSql = "exec stp_sel_CredBPPAnalistas  '" & psCodAge & "'," & pnMes & ", " & pnAnio
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerComiteAnaCord = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtenerComiteAnaCord:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Function
Public Sub OpeComiteAnaCord(ByVal psPersCod As String, ByVal pnComite As Integer, ByVal psCodAge As String, _
                        ByVal pnMes As Long, ByVal pnAnio As Long, ByVal pcUserReg As String, ByVal pdFecha As Date, _
                        ByVal pnEstado As Integer, ByVal pnCargo As Integer)
                        
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "Exec stp_ins_upd_CredBPPComite '" & psPersCod & "'," & pnComite & ",'" & psCodAge & "'," & _
            pnMes & "," & pnAnio & ",'" & pcUserReg & "','" & Format(pdFecha, "yyyymmdd hh:mm:s") & "'," & pnEstado & "," & pnCargo

Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub

Public Function GeneracionBPP(ByVal psCodAge As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorGeneracionBPP
    
    sSql = "exec stp_sel_GeneracionBPP  '" & psCodAge & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GeneracionBPP = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrorGeneracionBPP:
    Err.Raise Err.Number, "En Proceso", Err.Description
    
End Function

Public Sub OpeParametrosCoorgJefAg(ByVal pnOperacion As Integer, Optional ByVal pnCodigo As Integer, _
                                    Optional ByVal pnDesdeMinA As Long = 0, Optional ByVal pnHastaMinA As Long = 0, _
                                    Optional ByVal pnDesdeMaxA As Long = 0, Optional ByVal pnHastaMaxA As Long = 0, _
                                    Optional ByVal pnMinBoni As Double = 0, Optional ByVal pnMaxBoni As Double = 0)
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta


Select Case pnOperacion
    Case 1: lsSQL = "EXEC stp_ins_CredBPPConfigCoordJefes " & pnCodigo & "," & pnDesdeMinA & "," & pnHastaMinA & "," & pnDesdeMaxA & "," & pnHastaMaxA & "," & pnMinBoni & "," & pnMaxBoni
    Case 2: lsSQL = "EXEC stp_del_CredBPPConfigCoordJefes"
End Select

Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub

Public Function ObtenerParametrosCoorgJefAg(ByVal pnCodigo As Integer) As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "stp_sel_CredBPPConfigCoordJefes " & pnCodigo

Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
Set ObtenerParametrosCoorgJefAg = oConec.CargaRecordSet(lsSQL)
oConec.CierraConexion
End Function
Public Function ObtenerAgenciaConfig() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerAgenciaConfig
    
    sSql = "EXEC stp_sel_CredBPPAgenciaZonaCom"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerAgenciaConfig = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtenerAgenciaConfig:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Function

Public Sub OpeAgenciaConfig(ByVal psAgeCod As String, ByVal pnZona As Integer, ByVal pnComite As Long)
                        
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "Exec stp_ins_upd_CredBPPAgenciaZonaCom '" & psAgeCod & "'," & pnZona & "," & pnComite

Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub


Public Function ObtenerMetaTopeCoord(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psAgeCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerMetaTopeCoord
    
    sSql = "EXEC stp_sel_CredBPPMetaTopeCoord " & pnMes & "," & pnAnio & ",'" & psAgeCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerMetaTopeCoord = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtenerMetaTopeCoord:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Function

Public Sub OpeMetaTopeCoord(ByVal psAgeCod As String, ByVal pnComite As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal pnSaldo As Double, ByVal pnSaldoVenJud As Double, ByVal pnTope As Double, ByVal psUserReg As String, ByVal pdFecReg As Date)
                        
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "Exec stp_ins_upd_CredBPPMetaTopeCoord '" & psAgeCod & "'," & pnComite & "," & pnMes & "," & pnAnio & "," & pnSaldo & "," & pnSaldoVenJud & "," & pnTope & ",'" & psUserReg & "','" & Format(pdFecReg, "yyyymmdd hh:mm:s") & "'"

Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub

Public Function ObtenerMetaTopeJefA(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psAgeCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerMetaTopeJefA
    
    sSql = "EXEC stp_sel_CredBPPMetaTopeJefAg " & pnMes & "," & pnAnio & ",'" & psAgeCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerMetaTopeJefA = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtenerMetaTopeJefA:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Function

Public Sub OpeMetaTopeJefA(ByVal psAgeCod As String, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal pnSaldo As Double, ByVal pnSaldoVenJud As Double, ByVal pnTope As Double, ByVal psUserReg As String, ByVal pdFecReg As Date)
                        
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "Exec stp_ins_upd_CredBPPMetaTopeJefAg '" & psAgeCod & "'," & pnMes & "," & pnAnio & "," & pnSaldo & "," & pnSaldoVenJud & "," & pnTope & ",'" & psUserReg & "','" & Format(pdFecReg, "yyyymmdd hh:mm:s") & "'"

Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub


Public Function ObtenerMetaTopeJefTerritorial(ByVal pnMes As Integer, ByVal pnAnio As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerMetaTopeJefTerritorial
    
    sSql = "EXEC stp_sel_CredBPPMetaTopeJefTerritorial " & pnMes & "," & pnAnio
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerMetaTopeJefTerritorial = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtenerMetaTopeJefTerritorial:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Function


Public Sub OpeMetaTopeJefTerritorial(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal pnSaldo As Double, ByVal pnSaldoVenJud As Double, ByVal pnTope As Double, ByVal psUserReg As String, ByVal pdFecReg As Date, ByVal pnZona As Integer)
                        
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "Exec stp_ins_upd_CredBPPMetaTopeJefTerritorial " & pnMes & "," & pnAnio & "," & pnSaldo & "," & pnSaldoVenJud & "," & pnTope & ",'" & psUserReg & "','" & Format(pdFecReg, "yyyymmdd hh:mm:s") & "'," & pnZona

Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub

Public Function ObtenerJefAgenciaAct(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psAgeCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerJefAgenciaAct
    
    sSql = "EXEC stp_sel_CredBPPJefesAgencia " & pnMes & "," & pnAnio & ",'" & psAgeCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerJefAgenciaAct = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtenerJefAgenciaAct:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Function

Public Function ObtenerJefTerritorialAct(ByVal pnMes As Integer, ByVal pnAnio As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerJefTerritorialAct
    
    sSql = "EXEC stp_sel_CredBPPJefesTerritoriales " & pnMes & "," & pnAnio
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerJefTerritorialAct = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtenerJefTerritorialAct:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Function

Public Function ObtenerUsuariosJefTerritorialAct(ByVal pnMes As Integer, ByVal pnAnio As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerUsuariosJefTerritorialAct
    
    sSql = "EXEC stp_sel_MostrarJefesTerritoriales " & pnMes & "," & pnAnio
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerUsuariosJefTerritorialAct = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtenerUsuariosJefTerritorialAct:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Function

Public Sub OpeJefAgenciaAsig(ByVal psAgeCod As String, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psPersCod As String, ByVal pnEstado As Integer, ByVal psUserReg As String, ByVal pdFecReg As Date)

Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "Exec stp_ins_upd_CredBPPJefesAgencia '" & psAgeCod & "'," & pnMes & "," & pnAnio & ",'" & psPersCod & "'," & pnEstado & ",'" & psUserReg & "','" & Format(pdFecReg, "yyyymmdd hh:mm:s") & "'"

Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub

Public Sub OpeJefTerritorialAsig(ByVal pnZona As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psPersCod As String, ByVal psUserReg As String, ByVal pdFecReg As Date)
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "Exec stp_ins_upd_CredBPPJefesTerritoriales " & pnZona & "," & pnMes & "," & pnAnio & ",'" & psPersCod & "','" & psUserReg & "','" & Format(pdFecReg, "yyyymmdd hh:mm:s") & "'"

Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub

Public Function GenerarBPPCoordJA(ByVal psCodAge As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorGeneracionBPP
    
    sSql = "exec stp_sel_GeneraBPPCoordJefAG  '" & psCodAge & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GenerarBPPCoordJA = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrorGeneracionBPP:
    Err.Raise Err.Number, "En Proceso", Err.Description
    
End Function

Public Function GenerarBPPJT() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorGeneracionBPP
    
    sSql = "exec stp_sel_GeneraBPPJefeTerritorial"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GenerarBPPJT = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrorGeneracionBPP:
    Err.Raise Err.Number, "En Proceso", Err.Description
    
End Function

Public Function BonoGenerados(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorBonoGenerados
    
    sSql = "exec stp_sel_BPPGenerados " & pnMes & "," & pnAnio & "," & pnTipo
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set BonoGenerados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrorBonoGenerados:
    Err.Raise Err.Number, "En Proceso", Err.Description
    
End Function

Public Sub RegistraCierreBonosGenerados(ByVal pdFecha As Date, ByVal psMovNro As String, Optional ByVal pnEstado As Integer = 0)
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

lsSQL = "Exec stp_ins_CredBPPBonosGenerados '" & Format(pdFecha, "yyyymmdd") & "','" & psMovNro & "'," & pnEstado

Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub

Public Function ObtenerCierreBonoGenerados(ByVal pnAnio As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerCierreBonoGenerados
    
    sSql = "exec stp_sel_CredBPPBonosGenerados " & pnAnio
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerCierreBonoGenerados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrorObtenerCierreBonoGenerados:
    Err.Raise Err.Number, "En Proceso", Err.Description
    
End Function
'WIOR FIN ***************************************************************************
'WIOR 20140512 **********************************************************************
Public Sub OpeConfigPromotoresGen(ByVal pnTipo As Integer, Optional ByVal pnMontoMinCol As Double = 0, Optional ByVal pnMontoMaxBon As Double = 0, Optional ByVal pnTEM As Double = 0)
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta

If pnTipo = 1 Then
    lsSQL = "Exec stp_ins_CredBPPPromotoresConfGen " & pnMontoMinCol & "," & pnMontoMaxBon & "," & pnTEM
Else
    lsSQL = "Exec stp_del_CredBPPPromotoresConfGen"
End If

Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub

Public Sub OpeConfigPromotoresGenDet(ByVal pnOpe As Integer, ByVal pnTipo As Integer, Optional ByVal pnDesde As Double = 0, Optional ByVal pnHasta As Double = 0, Optional ByVal pnPorcentaje As Double = 0)
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta
    
If pnOpe = 1 Then
    lsSQL = "Exec stp_ins_CredBPPPromotoresConfGenDet " & pnTipo & "," & pnDesde & "," & pnHasta & "," & pnPorcentaje
Else
    lsSQL = "Exec stp_del_CredBPPPromotoresConfGenDet " & pnTipo
End If

Set oConec = New COMConecta.DCOMConecta

oConec.AbreConexion
oConec.Ejecutar lsSQL
oConec.CierraConexion
End Sub

Public Function ObtenerConfigPromotoresGen() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerConfigPromotoresGen
    
    sSql = "EXEC stp_sel_CredBPPPromotoresConfGen "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerConfigPromotoresGen = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrorObtenerConfigPromotoresGen:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Function
Public Function ObtenerConfigPromotoresGenDet(ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerConfigPromotoresGenDet
    
    sSql = "EXEC stp_sel_CredBPPPromotoresConfGenDet " & pnTipo
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerConfigPromotoresGenDet = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrorObtenerConfigPromotoresGenDet:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Function

Public Function ObtenerPromotoresMes(ByVal psMesAnio As String, Optional ByVal psAge As String = "%", Optional psPromotor As String = "%") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerPromotoresMes
    
    sSql = "EXEC stp_sel_CredPromotores '" & psMesAnio & "','" & psAge & "','" & psPromotor & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerPromotoresMes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrorObtenerPromotoresMes:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Function
Public Function ObtenerPromotoresMesDet(ByVal psMesAnio As String, Optional ByVal psAge As String = "%", Optional psPromotor As String = "%") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerPromotoresMes
    
    sSql = "EXEC stp_sel_CredPromotoresDet '" & psMesAnio & "','" & psAge & "','" & psPromotor & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerPromotoresMesDet = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrorObtenerPromotoresMes:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Function
'WIOR FIN ***************************************************************************
