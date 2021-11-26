VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMBPPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'JACA 20110318***************************************************************
Option Explicit
Public dbCmact As ADODB.Connection

Public Sub InsertarParametrosBPP(ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer, ByVal iClienteNuevo As Integer, ByVal iNroOpe As Integer, ByVal dMetaCrecimiento As Double, ByVal dCMDeAceptable As Double, ByVal dCMHastaAceptacion As Double, ByVal dCMDeSinDescuento As Double, ByVal dCMHastaSinDescuento As Double, _
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
                                        
    Dim objCOMDCredito As COMDCredito.DCOMBPPR
    Set objCOMDCredito = New COMDCredito.DCOMBPPR
    objCOMDCredito.RegistrarParametrosBPP iTipoCarteraId, iTipoClasificacionId, iClienteNuevo, iNroOpe, dMetaCrecimiento, dCMDeAceptable, dCMHastaAceptacion, dCMDeSinDescuento, dCMHastaSinDescuento, _
    dCMDeConDescuento1, dCMHastaConDescuento1, dCMDescConDescuento1, _
    dCMDeConDescuento2, dCMHastaConDescuento2, dCMDescConDescuento2, _
    dCMDeConDescuento3, dCMHastaConDescuento3, dCMDescConDescuento3, _
    dCMDeConDescuento4, dCMHastaConDescuento4, dCMDescConDescuento4, _
    dCMDeConDescuento5, dCMHastaConDescuento5, dCMDescConDescuento5, _
    dCMDeConDescuento6, dCMHastaConDescuento6, dCMDescConDescuento6, _
    dCMDeConDescuento7, dCMHastaConDescuento7, dCMDescConDescuento7, _
    dCMDeConDescuento8, dCMHastaConDescuento8, dCMDescConDescuento8, _
    dCMDeConDescuento9, dCMHastaConDescuento9, dCMDescConDescuento9, _
    dCMDeConDescuento10, dCMHastaConDescuento10, dCMDescConDescuento10, _
    dCMDeConDescuento11, dCMHastaConDescuento11, dCMDescConDescuento11, _
    dCMDeConDescuento12, dCMHastaConDescuento12, dCMDescConDescuento12, _
    sMovNro
End Sub
Public Sub actualizarParametrosBPP(ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer)
    Dim objCOMDCredito As COMDCredito.DCOMBPPR
    Set objCOMDCredito = New COMDCredito.DCOMBPPR
    objCOMDCredito.modificarParametrosBPP iTipoCarteraId, iTipoClasificacionId
End Sub
Public Sub InsertarAnalistaMetaBPP(ByVal sAgeCodAct As String, ByVal sCodAna As String, ByVal sUser As String, ByVal iTpoCar As Integer, ByVal iTpoCarClas As Integer, ByVal iCred As Integer, ByVal iClie As Integer, ByVal iOpe As Integer, ByVal iSaldo As Currency, ByVal iMora As Double, ByVal iTasa As Double, ByVal sCodCord As String, ByVal iMetaClie As Integer, ByVal iMetaOpe As Integer, ByVal iMetaSaldo As Long, ByVal sFechaMeta As String, ByVal sMovNro As String)
    Dim objCOMDCredito As COMDCredito.DCOMBPPR
    Set objCOMDCredito = New COMDCredito.DCOMBPPR
    objCOMDCredito.RegistrarAnalistaMetaBPP sAgeCodAct, sCodAna, sUser, iTpoCar, iTpoCarClas, iCred, iClie, iOpe, iSaldo, iMora, iTasa, sCodCord, iMetaClie, iMetaOpe, iMetaSaldo, sFechaMeta, sMovNro
End Sub
Public Function getCargarTipoClasificacion(ByVal iTipoCarteraId As Integer) As ADODB.Recordset
        Dim objCOMDCredito As COMDCredito.DCOMBPPR
        Set objCOMDCredito = New COMDCredito.DCOMBPPR
        Set getCargarTipoClasificacion = objCOMDCredito.CargarTipoClasificacion(iTipoCarteraId)
End Function
Public Function getCargarTipoCartera() As ADODB.Recordset
        Dim objCOMDCredito As COMDCredito.DCOMBPPR
        Set objCOMDCredito = New COMDCredito.DCOMBPPR
        Set getCargarTipoCartera = objCOMDCredito.CargarTipoCartera

End Function
Public Function getCargarParametrosTipoClasificacion(ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer) As ADODB.Recordset
        Dim objCOMDCredito As COMDCredito.DCOMBPPR
        Set objCOMDCredito = New COMDCredito.DCOMBPPR
        Set getCargarParametrosTipoClasificacion = objCOMDCredito.CargarParametrosTipoClasificacion(iTipoCarteraId, iTipoClasificacionId)

End Function
Public Function getCargarAgencias() As ADODB.Recordset
        Dim objCOMDCredito As COMDCredito.DCOMBPPR
        Set objCOMDCredito = New COMDCredito.DCOMBPPR
        Set getCargarAgencias = objCOMDCredito.CargarAgencias
End Function
'Public Function getCargarAnalistasXAgencia(ByVal sAgeCod As String, ByVal nConsSisValor As String, ByVal nTpoCambio As Double, ByVal nTipoCartera As Integer) As ADODB.Recordset
Public Function getCargarAnalistasXAgencia(ByVal sAgeCod As String, ByVal nTpoCambio As Double, ByVal nTipoCartera As Integer, ByVal nTipoCarteraClasificacion As Integer, ByVal dFecha As Date) As ADODB.Recordset
        Dim objCOMDCredito As COMDCredito.DCOMBPPR
        Set objCOMDCredito = New COMDCredito.DCOMBPPR
        Set getCargarAnalistasXAgencia = objCOMDCredito.CargarAnalistasXAgencia(sAgeCod, nTpoCambio, nTipoCartera, nTipoCarteraClasificacion, dFecha)
End Function
Public Function getCargarCordinadorXAgencia(ByVal sAgeCod As String) As ADODB.Recordset
        Dim objCOMDCredito As COMDCredito.DCOMBPPR
        Set objCOMDCredito = New COMDCredito.DCOMBPPR
        Set getCargarCordinadorXAgencia = objCOMDCredito.CargarCordinadorXAgencia(sAgeCod)
End Function

Public Function getFechaUlitmoMeta(ByVal sAgeCod As String, ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer) As ADODB.Recordset
        Dim objCOMDCredito As COMDCredito.DCOMBPPR
        Set objCOMDCredito = New COMDCredito.DCOMBPPR
        Set getFechaUlitmoMeta = objCOMDCredito.ObtenerFechaUlitmoMeta(sAgeCod, iTipoCarteraId, iTipoClasificacionId)
End Function
Public Function getListaAnalistaMetas(ByVal sAgeCod As String, ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer, ByVal sFechaMeta As String) As ADODB.Recordset
        Dim objCOMDCredito As COMDCredito.DCOMBPPR
        Set objCOMDCredito = New COMDCredito.DCOMBPPR
        Set getListaAnalistaMetas = objCOMDCredito.ObtenerListaAnalistaMetas(sAgeCod, iTipoCarteraId, iTipoClasificacionId, sFechaMeta)
End Function
Public Function getListaAnalistaMetasAnteriores(ByVal sAgeCod As String, ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer, ByVal sFechaMeta As String) As ADODB.Recordset
        Dim objCOMDCredito As COMDCredito.DCOMBPPR
        Set objCOMDCredito = New COMDCredito.DCOMBPPR
        Set getListaAnalistaMetasAnteriores = objCOMDCredito.ObtenerListaAnalistaMetasAnteriores(sAgeCod, iTipoCarteraId, iTipoClasificacionId, sFechaMeta)
End Function
Public Sub actualizarEstadoAnalistaMeta(ByVal sAgeCod As String, ByVal iTipoCarteraId As Integer, ByVal iTipoClasificacionId As Integer, ByVal sFechaMeta As String, ByVal sMovNro As String)
    Dim objCOMDCredito As COMDCredito.DCOMBPPR
    Set objCOMDCredito = New COMDCredito.DCOMBPPR
    objCOMDCredito.modificarEstadoAnalistaMeta sAgeCod, iTipoCarteraId, iTipoClasificacionId, sFechaMeta, sMovNro
End Sub
Public Function obtenerBonoAnalistas(ByVal sAgeCod As String, ByVal nTpoCambio As Double, ByVal nTipoCarteraId As Integer, ByVal nTipoClasificacionId As Integer, ByVal dFecha As Date) As ADODB.Recordset
        Dim objCOMDCredito As COMDCredito.DCOMBPPR
        Set objCOMDCredito = New COMDCredito.DCOMBPPR
        Set obtenerBonoAnalistas = objCOMDCredito.getBONOAnalistas(sAgeCod, nTpoCambio, nTipoCarteraId, nTipoClasificacionId, dFecha)
End Function
Public Sub guardarBonoAnalistas(ByVal psAgeCodAct As String, ByVal psCodana As String, ByVal pnTpoCar As Integer, ByVal pnTpoCarClas As Integer, _
                                ByVal pnCredAnt As Integer, ByVal pnClieAnt As Integer, ByVal pnSaldoAnt As Currency, _
                                ByVal pnCredTransf As Integer, ByVal pnClieTransf As Integer, ByVal pnSaldoTransf As Currency, _
                                ByVal pnMetaClie As Integer, ByVal pnMetaOpe As Integer, ByVal pnMetaSald As Currency, ByVal psCoord As String, _
                                ByVal pnCred As Integer, ByVal pnClie As Integer, ByVal pnSaldo As Currency, ByVal pnClieNue As Integer, ByVal pnClieRecu As Integer, _
                                ByVal pnOpe As Integer, ByVal pnMora As Double, ByVal pnTasaPond As Double, _
                                ByVal pnBonoClieNue As Currency, ByVal pnBonoClieRecu As Currency, ByVal pnBonoSaldo As Currency, ByVal pnBonoOpe As Currency, _
                                ByVal pnBonoMora As Currency, ByVal pnBonoTasa As Currency, ByVal pnBonoTotal As Currency, _
                                ByVal pnFecha As Date, ByVal sMovNro As String)
    Dim objCOMDCredito As COMDCredito.DCOMBPPR
    Set objCOMDCredito = New COMDCredito.DCOMBPPR
    objCOMDCredito.insertarBonoAnalistas psAgeCodAct, psCodana, pnTpoCar, pnTpoCarClas, _
                                          pnCredAnt, pnClieAnt, pnSaldoAnt, _
                                         pnCredTransf, pnClieTransf, pnSaldoTransf, _
                                         pnMetaClie, pnMetaOpe, pnMetaSald, psCoord, _
                                         pnCred, pnClie, pnSaldo, pnClieNue, pnClieRecu, _
                                         pnOpe, pnMora, pnTasaPond, _
                                         pnBonoClieNue, pnBonoClieRecu, pnBonoSaldo, pnBonoOpe, _
                                         pnBonoMora, pnBonoTasa, pnBonoTotal, _
                                         pnFecha, sMovNro
End Sub
Public Function obtenerBonoAnalistasConsol(ByVal sAgeCod As String, ByVal nTipoCarteraId As Integer, ByVal nTipoClasificacionId As Integer, ByVal dFecha As Date) As ADODB.Recordset
        Dim objCOMDCredito As COMDCredito.DCOMBPPR
        Set objCOMDCredito = New COMDCredito.DCOMBPPR
        Set obtenerBonoAnalistasConsol = objCOMDCredito.getBonoAnalistasConsol(sAgeCod, nTipoCarteraId, nTipoClasificacionId, dFecha)
End Function
'END JACA **********************************************************************

'JUEZ 20121010 *************************************************************************
Public Function RecuperaCredTiposCartera(Optional psFiltroId As String = "") As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    If psFiltroId = "" Then
        Set RecuperaCredTiposCartera = objCCredBPP.RecuperaCredTiposCartera()
    Else
        Set RecuperaCredTiposCartera = objCCredBPP.RecuperaCredTiposCartera(psFiltroId)
    End If
End Function

Public Sub dInsertaCredTipoCartera(ByVal psCartera As String, ByVal pnCredCorpo As Integer, ByVal pnGrandeEmp As Integer, _
                                    ByVal pnMedianaEmp As Integer, ByVal pnPequenaEmp As Integer, ByVal pnMicro As Integer, _
                                    ByVal pnConsumoRev As Integer, ByVal pnConsumoNoRev As Integer, ByVal pnHipotecario As Integer)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.dInsertaCredTipoCartera(psCartera, pnCredCorpo, pnGrandeEmp, pnMedianaEmp, pnPequenaEmp, pnMicro, _
                                            pnConsumoRev, pnConsumoNoRev, pnHipotecario)
End Sub

Public Sub dActualizaCredTipoCartera(ByVal psIdTipo As String, ByVal psCartera As String, ByVal pnCredCorpo As Integer, _
                                    ByVal pnGrandeEmp As Integer, ByVal pnMedianaEmp As Integer, ByVal pnPequenaEmp As Integer, _
                                    ByVal pnMicro As Integer, ByVal pnConsumoRev As Integer, ByVal pnConsumoNoRev As Integer, _
                                    ByVal pnHipotecario As Integer, Optional ByVal pbEjecBatch As Boolean = False)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.dActualizaCredTipoCartera(psIdTipo, psCartera, pnCredCorpo, pnGrandeEmp, pnMedianaEmp, pnPequenaEmp, pnMicro, _
                                                pnConsumoRev, pnConsumoNoRev, pnHipotecario)
End Sub

Public Sub dEliminaCredTipoCartera(ByVal psIdTipo As String, Optional ByVal pbEjecBatch As Boolean = False)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.dEliminaCredTipoCartera(psIdTipo)
End Sub

Public Function ListarTiposCartera() As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    Set ListarTiposCartera = objCCredBPP.ListarTiposCartera()
End Function

Public Function ListarCategoriaAnalistas() As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    Set ListarCategoriaAnalistas = objCCredBPP.ListarCategoriaAnalistas
End Function

Public Sub dInsertaParametro(ByVal psTipoCartera As String, ByVal pnMatTpoProd As Variant, ByVal pnMatAgencias As Variant, _
                            ByVal pnMatIndicad As Variant, ByVal psCatAnalista As String, ByVal pnMinCartera As Double, _
                            ByVal pnMaxCartera As Double)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Dim psIdParametro As String
    
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.dInsertaParametroDatos(psTipoCartera, psCatAnalista, pnMinCartera, pnMaxCartera)
    psIdParametro = objCCredBPP.RecuperaParametro(psTipoCartera, psCatAnalista, pnMinCartera, pnMaxCartera)
    
    Call dInsertaParametrosListas(psIdParametro, pnMatTpoProd, 1)
    Call dInsertaParametrosListas(psIdParametro, pnMatAgencias, 2)
    Call dInsertaParametrosListas(psIdParametro, pnMatIndicad, 3)
    
End Sub

Public Sub dActualizaParametro(ByVal psTipoCartera As String, ByVal pnMatTpoProd As Variant, ByVal pnMatAgencias As Variant, _
                                            ByVal pnMatIndicad As Variant, ByVal psCatAnalista As String, ByVal pnMinCartera As Double, _
                                            ByVal pnMaxCartera As Double, ByVal psIdParametro As String)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.dActualizaParametroDatos(psIdParametro, psTipoCartera, psCatAnalista, pnMinCartera, pnMaxCartera)
    Call objCCredBPP.dEliminaParametrosListas(psIdParametro)
    Call dInsertaParametrosListas(psIdParametro, pnMatTpoProd, 1)
    Call dInsertaParametrosListas(psIdParametro, pnMatAgencias, 2)
    Call dInsertaParametrosListas(psIdParametro, pnMatIndicad, 3)
End Sub

Public Sub dInsertaParametrosListas(ByVal psIdParam As String, ByVal pnMatLista As Variant, ByVal pnTipoMat As Integer)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Dim i As Integer
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    If pnTipoMat = 1 Or pnTipoMat = 2 Then
        For i = 1 To UBound(pnMatLista)
            Call objCCredBPP.dInsertaParametrosLista(psIdParam, pnMatLista(i, 0), pnTipoMat)
        Next i
    ElseIf pnTipoMat = 3 Then
        For i = 1 To UBound(pnMatLista)
            Call objCCredBPP.dInsertaParametrosIndicadoresCartera(psIdParam, pnMatLista(i - 1, 0), pnMatLista(i - 1, 1), pnMatLista(i - 1, 2))
        Next i
    End If
End Sub

Public Sub dEliminaParametro(ByVal psIdParam As String)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.dEliminaParametroDatos(psIdParam)
    'Call objCCredBPP.dEliminaParametrosListas(psIdParam)
End Sub

Public Function VerificaSiExisteParametro(ByVal psTipoCartera As String, ByVal psCateAnalista As String, ByVal psIdParam As String) As Boolean
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    VerificaSiExisteParametro = objCCredBPP.VerificaSiExisteParametro(psTipoCartera, psCateAnalista, psIdParam)
End Function

Public Function VerificaTamanoCartera(ByVal psIdParam As String, ByVal psTipoCartera As String, ByVal pnMinCart As Double, _
                                        ByVal pnMaxCart As Double, ByRef cCatPertenece As String) As Boolean
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    VerificaTamanoCartera = objCCredBPP.VerificaTamanoCartera(psIdParam, psTipoCartera, pnMinCart, pnMaxCart, cCatPertenece)
End Function

'END JUEZ ******************************************************************************
'WIOR 20130524 **********************************************************************
Public Function ObtenerParamCabXNivel(ByVal pnNivel As String, ByRef pbSalCartera As Boolean, ByRef pbNumClientes As Boolean) As ADODB.Recordset
Dim oDBPP As COMDCredito.DCOMBPPR
Dim rsDBPP As ADODB.Recordset

Set oDBPP = New COMDCredito.DCOMBPPR

Set rsDBPP = oDBPP.ObtenerCatOParamCabXNivel(2, pnNivel)

If Not (rsDBPP.BOF And rsDBPP.EOF) Then
    pbSalCartera = CBool(rsDBPP!bSaldoCartera)
    pbNumClientes = CBool(rsDBPP!bNumClientes)
Else
    pbSalCartera = False
    pbNumClientes = False
End If

Set rsDBPP = Nothing
End Function
Public Sub EliminaParametrosCumplimiento(ByVal pnNivel As Integer, _
                                        ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, _
                                        ByVal pnTipoParam As Integer)
    
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.EliminaParametrosCumplimiento(pnNivel, pnMes, pnAnio, psCodAge, pnTipoParam)
End Sub
Public Function DevolverFactoresCumplimiento(ByVal pnNivel As Integer, ByVal pnTipo As Integer) As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    Set DevolverFactoresCumplimiento = objCCredBPP.DevolverFactoresCumplimiento(pnNivel, pnTipo)
End Function

Public Sub InsertaParametroCumplimiento(ByVal pnNivel As Integer, ByVal psCodFactor As String, ByVal pnCategoriaA As Double, _
                                        ByVal pnCategoriaB As Double, ByVal pnCategoriaC As Double, ByVal pnCategoriaD As Double, _
                                        ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, ByVal psUserReg As String, _
                                        ByVal pdFecReg As String, ByVal pnTipoParam As Integer)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.InsertaParametrosCumplimiento(pnNivel, psCodFactor, pnCategoriaA, pnCategoriaB, pnCategoriaC, pnCategoriaD, pnMes, _
                                                   pnAnio, psCodAge, psUserReg, pdFecReg, pnTipoParam)
End Sub
Public Function DevolverParametrosCumplimiento(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, _
                                        ByVal pnTipoParam As Integer) As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Set DevolverParametrosCumplimiento = objCCredBPP.DevolverParametrosCumplimiento(pnNivel, pnMes, pnAnio, psCodAge, pnTipoParam)
End Function

Public Sub InsertaRangosdeMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, ByVal pnMora8a30 As Double, _
                               ByVal pnMoraMaya30 As Double, ByVal psUserReg As String, ByVal pdFecReg As String, _
                               ByVal pnNivel As Integer, ByVal psCategoria As String)
                               
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.InsertaRangosMora(pnMes, pnAnio, psCodAge, pnMora8a30, pnMoraMaya30, psUserReg, pdFecReg, pnNivel, psCategoria)
End Sub

Public Sub EliminaRangosMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, ByVal pnNivel As Integer)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.EliminaRangosMora(pnMes, pnAnio, psCodAge, pnNivel)
End Sub


Public Function DevolverRangosMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, ByVal pnNivel As Integer) As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Set DevolverRangosMora = objCCredBPP.DevolverRangosMora(pnMes, pnAnio, psCodAge, pnNivel)
End Function
Public Function DevolverFechasAMigrarXParam(ByVal pnTipo As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String) As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Set DevolverFechasAMigrarXParam = objCCredBPP.DevolverFechasAMigrarXParam(pnTipo, pnMes, pnAnio, psCodAge)
End Function

Public Sub MigrarDatosXParam(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, ByVal pnTipo As Integer, ByVal pnMesAct As Integer, ByVal pnAnioAct As Integer)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.MigrarDatosXParam(pnMes, pnAnio, psCodAge, pnTipo, pnMesAct, pnAnioAct)
End Sub


Public Function DevolverFactorRend(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String) As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set DevolverFactorRend = objCCredBPP.DevolverFactorRend(pnMes, pnAnio, psCodAge)
End Function

Public Sub EliminaFactorRend(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
    
Call objCCredBPP.EliminaFactorRend(pnMes, pnAnio, psCodAge)
End Sub
Public Sub InsertaFactorRend(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCategoria As String, ByVal pnNivel1 As Double, _
                             ByVal pnNivel2 As Double, ByVal pnNivel3 As Double, ByVal psUserReg As String, ByVal pdFecReg As String, _
                             ByVal psCodAge As String)
                             
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.InsertaFactorRend(pnMes, pnAnio, psCategoria, pnNivel1, pnNivel2, pnNivel3, psUserReg, pdFecReg, psCodAge)
End Sub
Public Function DevolverFactorOpe(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String) As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
    
Set DevolverFactorOpe = objCCredBPP.DevolverFactorOpe(pnMes, pnAnio, psCodAge)
End Function

Public Sub EliminaFactorOpe(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
    
Call objCCredBPP.EliminaFactorOpe(pnMes, pnAnio, psCodAge)
End Sub

Public Sub InsertaFactorOpe(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal pnQuincena1 As Double, _
                             ByVal pnQuincena2 As Double, ByVal psUserReg As String, ByVal pdFecReg As String, _
                             ByVal psCodAge As String)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.InsertaFactorOpe(pnMes, pnAnio, pnQuincena1, pnQuincena2, psUserReg, pdFecReg, psCodAge)
End Sub

Public Function DevolverCategorias() As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Set DevolverCategorias = objCCredBPP.DevolverCategorias()
End Function

Public Sub InsertaPesoyTopes(ByVal pnNivel As Integer, ByVal psCodFactor As String, ByVal pnPesoCateA As Double, ByVal pnTopeCateA As Double, _
                             ByVal pnPesoCateB As Double, ByVal pnTopeCateB As Double, ByVal pnPesoCateC As Double, ByVal pnTopeCateC As Double, _
                             ByVal pnPesoCateD As Double, ByVal pnTopeCateD As Double, ByVal pnMes As Integer, ByVal pnAnio As Integer, _
                             ByVal psCodAge As String, ByVal psUserReg As String, ByVal pdFecReg As String)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.InsertaPesoyTopes(pnNivel, psCodFactor, pnPesoCateA, pnTopeCateA, pnPesoCateB, pnTopeCateB, pnPesoCateC, pnTopeCateC, _
                                       pnPesoCateD, pnTopeCateD, pnMes, pnAnio, psCodAge, psUserReg, pdFecReg)
End Sub

Public Sub EliminaPesoyTopes(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
 Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.EliminaPesoyTopes(pnNivel, pnMes, pnAnio, psCodAge)
End Sub

Public Function DevolverPesoyTopes(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String) As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Set DevolverPesoyTopes = objCCredBPP.DevolverPesoyTopes(pnNivel, pnMes, pnAnio, psCodAge)
End Function

Public Sub InsertaPenalidadXIncreMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, _
                                      ByVal pnDesde As Currency, ByVal pnHasta As Currency, ByVal pnDscto As Currency, _
                                      ByVal psUserReg As String, ByVal pdFecReg As String)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.InsertaPenalidadXIncreMora(pnMes, pnAnio, psCodAge, pnDesde, pnHasta, pnDscto, psUserReg, pdFecReg)
End Sub
Public Sub EliminaPenalidadXIncreMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Call objCCredBPP.EliminaPenalidadXIncreMora(pnMes, pnAnio, psCodAge)
End Sub
Public Function DevolverPenalidadXIncreMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String) As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Set DevolverPenalidadXIncreMora = objCCredBPP.DevolverPenalidadXIncreMora(pnMes, pnAnio, psCodAge)
End Function
Public Function ObtenerMetasAgencia(ByVal pdFecha As Date, ByVal psCodAge As String, ByVal pnNivel As Integer) As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Set ObtenerMetasAgencia = objCCredBPP.ObtenerMetasAgencia(pdFecha, psCodAge, pnNivel)
End Function
Public Sub OpeMetasAgencia(ByVal psPersCod As String, ByVal pnSaldoAg As Double, ByVal pnNCliAg As Double, _
                        ByVal pnNOpeAg As Double, ByVal pnMoraAg As Double, ByVal pnMes As Long, ByVal pnAnio As Long, _
                        ByVal pcUserReg As String, ByVal pdFecha As Date, _
                        ByVal pnNivel As Integer)
                        
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
    
Call objCCredBPP.OpeMetasAgencia(psPersCod, pnSaldoAg, pnNCliAg, pnNOpeAg, pnMoraAg, pnMes, pnAnio, pcUserReg, pdFecha, pnNivel)
End Sub

Public Function ObtenerComiteAnaCord(ByVal psCodAge As String, ByVal pnMes As Integer, ByVal pnAnio As Integer) As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Set ObtenerComiteAnaCord = objCCredBPP.ObtenerComiteAnaCord(psCodAge, pnMes, pnAnio)
End Function
Public Sub OpeComiteAnaCord(ByVal psPersCod As String, ByVal pnComite As Integer, ByVal psCodAge As String, _
                        ByVal pnMes As Long, ByVal pnAnio As Long, ByVal pcUserReg As String, ByVal pdFecha As Date, _
                        ByVal pnEstado As Integer, ByVal pnCargo As Integer)
                                      
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
    
Call objCCredBPP.OpeComiteAnaCord(psPersCod, pnComite, psCodAge, pnMes, pnAnio, pcUserReg, pdFecha, pnEstado, pnCargo)

End Sub
Public Function GeneracionBPP(ByVal psCodAge As String) As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Set GeneracionBPP = objCCredBPP.GeneracionBPP(psCodAge)
End Function

Public Function ObtenerAgenciaConfig() As ADODB.Recordset
    Dim objCCredBPP As COMDCredito.DCOMBPPR
    Set objCCredBPP = New COMDCredito.DCOMBPPR
    
    Set ObtenerAgenciaConfig = objCCredBPP.ObtenerAgenciaConfig()
End Function
Public Sub OpeAgenciaConfig(ByVal psAgeCod As String, ByVal pnZona As Integer, ByVal pnComite As Long)
                                      
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
    
Call objCCredBPP.OpeAgenciaConfig(psAgeCod, pnZona, pnComite)

End Sub

Public Function ObtenerMetaTopeCoord(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psAgeCod As String) As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set ObtenerMetaTopeCoord = objCCredBPP.ObtenerMetaTopeCoord(pnMes, pnAnio, psAgeCod)
End Function

Public Sub OpeMetaTopeCoord(ByVal psAgeCod As String, ByVal pnComite As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal pnSaldo As Double, ByVal pnSaldoVenJud As Double, ByVal pnTope As Double, ByVal psUserReg As String, ByVal pdFecReg As Date)
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
    
Call objCCredBPP.OpeMetaTopeCoord(psAgeCod, pnComite, pnMes, pnAnio, pnSaldo, pnSaldoVenJud, pnTope, psUserReg, pdFecReg)
End Sub

Public Function ObtenerMetaTopeJefA(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psAgeCod As String) As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set ObtenerMetaTopeJefA = objCCredBPP.ObtenerMetaTopeJefA(pnMes, pnAnio, psAgeCod)
End Function

Public Sub OpeMetaTopeJefA(ByVal psAgeCod As String, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal pnSaldo As Double, ByVal pnSaldoVenJud As Double, ByVal pnTope As Double, ByVal psUserReg As String, ByVal pdFecReg As Date)
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
    
Call objCCredBPP.OpeMetaTopeJefA(psAgeCod, pnMes, pnAnio, pnSaldo, pnSaldoVenJud, pnTope, psUserReg, pdFecReg)
End Sub


Public Function ObtenerMetaTopeJefTerritorial(ByVal pnMes As Integer, ByVal pnAnio As Integer) As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set ObtenerMetaTopeJefTerritorial = objCCredBPP.ObtenerMetaTopeJefTerritorial(pnMes, pnAnio)
End Function

Public Sub OpeMetaTopeJefTerritorial(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal pnSaldo As Double, ByVal pnSaldoVenJud As Double, ByVal pnTope As Double, ByVal psUserReg As String, ByVal pdFecReg As Date, ByVal pnZona As Integer)
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
    
Call objCCredBPP.OpeMetaTopeJefTerritorial(pnMes, pnAnio, pnSaldo, pnSaldoVenJud, pnTope, psUserReg, pdFecReg, pnZona)
End Sub

Public Function ObtenerJefAgenciaAct(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psAgeCod As String) As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set ObtenerJefAgenciaAct = objCCredBPP.ObtenerJefAgenciaAct(pnMes, pnAnio, psAgeCod)
End Function

Public Function ObtenerJefTerritorialAct(ByVal pnMes As Integer, ByVal pnAnio As Integer) As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set ObtenerJefTerritorialAct = objCCredBPP.ObtenerJefTerritorialAct(pnMes, pnAnio)
End Function

Public Function ObtenerUsuariosJefTerritorialAct(ByVal pnMes As Integer, ByVal pnAnio As Integer) As ADODB.Recordset
    Dim objCCredBPP As New COMDCredito.DCOMBPPR
    Set ObtenerUsuariosJefTerritorialAct = objCCredBPP.ObtenerUsuariosJefTerritorialAct(pnMes, pnAnio)
    Set objCCredBPP = Nothing
End Function

Public Sub OpeJefAgenciaAsig(ByVal psAgeCod As String, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psPersCod As String, ByVal pnEstado As Integer, ByVal psUserReg As String, ByVal pdFecReg As Date)
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
    
Call objCCredBPP.OpeJefAgenciaAsig(psAgeCod, pnMes, pnAnio, psPersCod, pnEstado, psUserReg, pdFecReg)
End Sub

Public Sub OpeJefTerritorialAsig(ByVal pnZona As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psPersCod As String, ByVal psUserReg As String, ByVal pdFecReg As Date)
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
    
Call objCCredBPP.OpeJefTerritorialAsig(pnZona, pnMes, pnAnio, psPersCod, psUserReg, pdFecReg)
End Sub
Public Function GenerarBPPCoordJA(ByVal psCodAge As String) As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set GenerarBPPCoordJA = objCCredBPP.GenerarBPPCoordJA(psCodAge)
End Function

Public Function GenerarBPPJT() As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set GenerarBPPJT = objCCredBPP.GenerarBPPJT
End Function

Public Function BonoGenerados(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal pnTipo As Integer) As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set BonoGenerados = objCCredBPP.BonoGenerados(pnMes, pnAnio, pnTipo)
End Function

Public Sub RegistraCierreBonosGenerados(ByVal pdFecha As Date, ByVal psMovNro As String, Optional ByVal pnEstado As Integer = 0)
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
    
Call objCCredBPP.RegistraCierreBonosGenerados(pdFecha, psMovNro, pnEstado)
End Sub

Public Function ObtenerCierreBonoGenerados(ByVal pnAnio As Integer) As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set ObtenerCierreBonoGenerados = objCCredBPP.ObtenerCierreBonoGenerados(pnAnio)
   
End Function
'WIOR FIN ***************************************************************************
'WIOR 20140512 **********************************************************************
Public Sub OpeConfigPromotoresGen(ByVal pnTipo As Integer, Optional ByVal pnMontoMinCol As Double = 0, Optional ByVal pnMontoMaxBon As Double = 0, Optional ByVal pnTEM As Double = 0)
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
Call objCCredBPP.OpeConfigPromotoresGen(pnTipo, pnMontoMinCol, pnMontoMaxBon, pnTEM)
Set objCCredBPP = Nothing
End Sub

Public Sub OpeConfigPromotoresGenDet(ByVal pnOpe As Integer, ByVal pnTipo As Integer, Optional ByVal pnDesde As Double = 0, Optional ByVal pnHasta As Double = 0, Optional ByVal pnPorcentaje As Double = 0)
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR
Call objCCredBPP.OpeConfigPromotoresGenDet(pnOpe, pnTipo, pnDesde, pnHasta, pnPorcentaje)
Set objCCredBPP = Nothing
End Sub

Public Function ObtenerConfigPromotoresGen() As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set ObtenerConfigPromotoresGen = objCCredBPP.ObtenerConfigPromotoresGen()
Set objCCredBPP = Nothing
End Function
Public Function ObtenerConfigPromotoresGenDet(ByVal pnTipo As Integer) As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set ObtenerConfigPromotoresGenDet = objCCredBPP.ObtenerConfigPromotoresGenDet(pnTipo)
Set objCCredBPP = Nothing
End Function

Public Function ObtenerPromotoresMes(ByVal psMesAnio As String, Optional ByVal psAge As String = "%", Optional psPromotor As String = "%") As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set ObtenerPromotoresMes = objCCredBPP.ObtenerPromotoresMes(psMesAnio, psAge, psPromotor)
Set objCCredBPP = Nothing
End Function
Public Function ObtenerPromotoresMesDet(ByVal psMesAnio As String, Optional ByVal psAge As String = "%", Optional psPromotor As String = "%") As ADODB.Recordset
Dim objCCredBPP As COMDCredito.DCOMBPPR
Set objCCredBPP = New COMDCredito.DCOMBPPR

Set ObtenerPromotoresMesDet = objCCredBPP.ObtenerPromotoresMesDet(psMesAnio, psAge, psPromotor)
Set objCCredBPP = Nothing
End Function
'WIOR FIN ***************************************************************************
