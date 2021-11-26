VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMColRecCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim coConex As COMConecta.DCOMConecta
Dim csCentralCom As String
Dim oError As New COMConecta.COMErrorHandling

Private Sub Class_Initialize()
    Dim loIni As COMConecta.DCOMClasIni
    Dim csConexion As String
    Dim csNegocio As String
    Dim csCentralPer As String
    Dim csCentralImg As String
    Dim csAdminist As String
    
    Set loIni = New COMConecta.DCOMClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing

Set coConex = New COMConecta.DCOMConecta
If coConex.AbreConexion(csConexion) = False Then
    Call oError.RaiseError(oError.MyUnhandledError, "DColPContrato:Initialize. Error en Conexion a Base de datos")
End If
End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub


Public Function dObtieneCreditosDePersona(ByVal psPersCod As String, _
        ByVal psEstados As String, Optional ByVal psAgencia As String = "") As ADODB.Recordset
'Obtiene Listado Creditos en Recuperaciones  de una Persona
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Dim lsAgencia As String

If Trim(psAgencia) = "" Then
    lsAgencia = "__"
Else
    lsAgencia = Trim(psAgencia)
End If

On Error GoTo dError
'MAVM 201000610
'lssql = "SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
'        & "T1.cConsDescripcion cRelacion, UPPER(T2.cConsDescripcion) cProducto, " _
'        & "UPPER(T3.cConsDescripcion) cMoneda " _
'        & "FROM ProductoPersona PP INNER JOIN Producto P  " _
'        & "ON P.cCtaCod = PP.cCtaCod INNER JOIN " _
'        & csCentralCom & "Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN " & csCentralCom & "" _
'        & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN " & csCentralCom & "Constante T2 " _
'        & "ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.nConsValor) INNER JOIN " & csCentralCom & "" _
'        & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) " _
'        & "WHERE PP.cPersCod = '" & psPersCod & "' " _
'        & "AND T1.nConsCod = " & gColocRelacPers & " AND T.nConsCod = " & gColocEstado & " AND " _
'        & "T2.nConsCod = " & gProducto & " AND T3.nConsCod = " & gMoneda _
'        & "AND P.cCtaCod like '___" & lsAgencia & "%' " _
'        & "AND P.nPrdEstado in ( " & psEstados & " ,2205 , 2206  ) "
'lssql = lssql & " ORDER BY PP.cCtaCod"

'*** PEAC 20120711
'lssql = "SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
'        & "T1.cConsDescripcion cRelacion, isnull(CN4.nConsValor,0) nTipoProdCod, ISNULL(CN4.cConsDescripcion,'') as cTipoProdDescrip, " _
'        & "UPPER(T3.cConsDescripcion) cMoneda " _
'        & "FROM ProductoPersona PP INNER JOIN Producto P  " _
'        & "INNER JOIN Colocaciones CO ON P.cCtaCod = CO.cCtaCod " _
'        & "ON P.cCtaCod = PP.cCtaCod INNER JOIN " _
'        & csCentralCom & "Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN " & csCentralCom & "" _
'        & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN Constante CN4 " _
'        & "ON convert(int,substring(cTpoProdCod,1,2)+'0') = CN4.nConsValor AND CN4.nConsCod = " & gTpoProducto & " INNER JOIN " & csCentralCom & "" _
'        & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) " _
'        & "WHERE PP.cPersCod = '" & psPersCod & "' " _
'        & "AND T1.nConsCod = " & gColocRelacPers & " AND T.nConsCod = " & gColocEstado & " AND " _
'        & "T3.nConsCod = " & gMoneda _
'        & " AND P.cCtaCod like '___" & lsAgencia & "%' " _
'        & "AND P.nPrdEstado in ( " & psEstados & " , 2205, 2206  ) "
'lssql = lssql & " ORDER BY PP.cCtaCod"

lsSQL = "exec stp_sel_ObtieneCreditosDePersona '" & psPersCod & "','" & psEstados & "','" & lsAgencia & "'"

'*** FIN PEAC

Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneCreditosDePersona = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function dObtieneListaGastosxCredito(ByVal psCtaCod As String, ByRef psmensaje As String, Optional ByVal pbCredTransferido As Boolean = False) As ADODB.Recordset
'FRHU 20150415 ERS022-2015 Se agrego pbCredTransferido
'************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset
    'FRHU 20150415 ERS022-2015
    'lsSQL = " SELECT G.dAsigna, G.nMonto, G.nMontoPagado, G.cMotivoGasto,  " _
    '      & " G.nPrdConceptoCod, G.nNroGastoCta, G.nColocRecGastoEstado  " _
    '      & " FROM ColocRecup CRec INNER JOIN ColocRecupGastos G " _
    '      & " ON CRec.cCtaCod = G.cCtaCod " _
    '      & " " _
    '      & " WHERE CRec.cCtaCod ='" & psctacod & "' " _
    '      & " AND G.nColocRecGastoEstado not in ( " & gColRecGastoEstEliminado & " ) " _
    '      & " ORDER BY G.nNroGastoCta "

    If pbCredTransferido Then
        lsSQL = "stp_sel_ObtieneListaGastosxCreditoTransferido '" & psCtaCod & "'"
    Else
        lsSQL = "stp_sel_ObtieneListaGastosxCredito '" & psCtaCod & "'"
    End If
    'FIN FRHU 20150415
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
        Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    
    Set dObtieneListaGastosxCredito = lrValida
    Set lrValida = Nothing
End Function

Public Function dObtieneDatosAsignaGastoCredRecup(ByVal psCtaCod As String, Optional ByVal pbCredTransferido As Boolean = False) As ADODB.Recordset
'Se agrego pbCredTransferido 'FRHU 20150415 ERS022-2015
'Obtiene Datos de Credito en Recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New ADODB.Recordset

On Error GoTo dError
'FRHU 20150415 ERS022-2015
'lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.nMontoCol, P.nSaldo,  " _
'    & " CRec.dIngRecup, CRec.nSaldoIntComp, CRec.nSaldoIntMor, CRec.nSaldoGasto, " _
'    & " Pers.cPersCod, Pers.cPersNombre, " _
'    & " nUltGasto = (SELECT ISNULL(MAX(nNroGastoCta),0) FROM ColocRecupGastos " _
'    & "              WHERE cCtaCod ='" & psctacod & "' ) " _
'    & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'    & " Inner Join ColocRecup CRec ON C.cctacod = CRec.cctacod " _
'    & " Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod " _
'    & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod " _
'    & " WHERE P.cCtaCod ='" & psctacod & "' " _
'    & " AND PP.nPrdPersRelac = " & gColRelPersTitular & " "
If pbCredTransferido Then
    lsSQL = "EXEC stp_sel_ObtieneDatosAsignaGastoCredTransferidos '" & psCtaCod & "'"
Else
    lsSQL = "EXEC stp_sel_ObtieneDatosAsignaGastoCredRecup '" & psCtaCod & "'"
End If
'FIN FRHU 20150415
Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneDatosAsignaGastoCredRecup = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function dObtieneGastosRecup() As ADODB.Recordset
'Obtiene Datos de Credito en Recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New ADODB.Recordset

On Error GoTo dError

lsSQL = " SELECT nPrdConceptoCod, cDescripcion " _
    & " FROM ProductoConcepto " _
    & " WHERE nPrdConceptoCod like '32%' AND cAplicaProceso = 'MA' " _
    & " ORDER BY nPrdConceptoCod "
    
Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneGastosRecup = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function dObtieneDatosPagoCredRecup(ByVal psCtaCod As String, ByRef psmensaje As String) As ADODB.Recordset
'Obtiene Datos para Pago de Credito en Recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Dim cRecDesc As DCOMColRecDesc
Set lrs = New ADODB.Recordset
On Error GoTo dError

'Comentado por DAOR
'lssql = " SELECT P.cCtaCod, PP.cPersCod, P.nPrdEstado, C.nMontoCol, P.nSaldo, C.cUltimaActualizacion,  " _
'    & " CRec.nSaldoIntComp, CRec.nSaldoIntMor, CRec.nSaldoGasto, CRec.nIntCompGen, CRec.cNroNeg, " _
'    & " CRec.nComisionCod, CRec.cMetLiquid, CRec.nTipCj, CRec.nDemanda, CRec.nNroCalen, IsNull(Com.nValor,0) nValorCom,  " _
'    & " Pers.cPersCod, Pers.cPersNombre, Pers2.cPersCod cPersCodAbog, Pers2.cPersNombre cPersNombreAbog,  " _
'    & " nUltGas = (Select Isnull(max(nNroGastoCta),0) From ColocRecupGastos CRG Where CRG.cCtaCod = p.cCtaCod ) , " _
'    & " nTasaInt = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT  WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntCompNormal & " ) , " _
'    & " nTasaIntMor = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT  WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ), " _
'    & " cEstudioAbogCod = (SELECT cPersCod FROM ProductoPersona  " _
'    & "           WHERE cCtaCod ='" & psctacod & "' AND nPrdPersRelac = " & gColRelPersEstudioJuridico & " ) " _
'    & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'    & " Inner Join ColocRecup CRec ON C.cctacod = CRec.cctacod " _
'    & " Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular _
'    & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod " _
'    & " Inner Join ProductoPersona PP2 ON C.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersEstudioJuridico _
'    & " Inner Join Persona Pers2 ON Pers2.cPersCod = PP2.cPersCod " _
'    & " Left Join ColocRecupComision Com ON Com.nComisionCod = CRec.nComisionCod " _
'    & " WHERE P.cCtaCod ='" & psctacod & "'"

'DAOR 20070421
'COMENTO ARLO20190417 BEGIN
'DAOR 20070802, se aumentó en el select el campo nIntMoraGen
'lsSQL = " SELECT P.cCtaCod, PP.cPersCod, P.nPrdEstado, C.nMontoCol, P.nSaldo, C.cUltimaActualizacion,  " _
'    & " CRec.nSaldoIntComp, CRec.nSaldoIntMor, CRec.nSaldoGasto, CRec.nIntCompGen,CRec.nIntMoraGen, CRec.cNroNeg, " _
'    & " CRec.nComisionCod, CRec.cMetLiquid, CRec.nTipCj, CRec.nDemanda, CRec.nNroCalen, IsNull(Com.nValor,0) nValorCom,  " _
'    & " Pers.cPersCod, Pers.cPersNombre, Pers2.cPersCod cPersCodAbog, Pers2.cPersNombre cPersNombreAbog,  " _
'    & " nUltGas = (Select Isnull(max(nNroGastoCta),0) From ColocRecupGastos CRG Where CRG.cCtaCod = p.cCtaCod ) , " _
'    & " nTasaInt = (SELECT ISNULL(nTasaInteres, 0) From ProductoTasaInteres LCT  WHERE LCT.cCtaCod = P.cCtaCod and LCT.nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & " ) , " _
'    & " nTasaIntMor = (SELECT ISNULL(nTasaInteres, 0) From ProductoTasaInteres LCT  WHERE LCT.cCtaCod = P.cCtaCod and LCT.nPrdTasaInteres = " & gColocLineaCredTasasIntMoratNormal & " ), " _
'    & " cEstudioAbogCod = (SELECT cPersCod FROM ProductoPersona  " _
'    & "           WHERE cCtaCod ='" & psCtaCod & "' AND nPrdPersRelac = " & gColRelPersEstudioJuridico & " ) " _
'    & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'    & " Inner Join ColocRecup CRec ON C.cctacod = CRec.cctacod " _
'    & " Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular _
'    & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod " _
'    & " Inner Join ProductoPersona PP2 ON C.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersEstudioJuridico _
'    & " Inner Join Persona Pers2 ON Pers2.cPersCod = PP2.cPersCod " _
'    & " Left Join ColocRecupComision Com ON Com.nComisionCod = CRec.nComisionCod " _
'    & " WHERE P.cCtaCod ='" & psCtaCod & "'"
'ARLO END

lsSQL = "EXEC stp_sel_dObtieneDatosPagoCredRecup '" & psCtaCod & "'" 'ARLO20190417

Set lrs = coConex.CargaRecordSet(lsSQL)
    
    If lrs.BOF And lrs.EOF Then
        psmensaje = "No se ha realizado registro de expediente "
        Exit Function
    End If
    If lrs!nprdestado = gColocEstRecCanJud Or lrs!nprdestado = gColocEstRecCanCast Then
        Set cRecDesc = New DCOMColRecDesc
        psmensaje = "Credito se encuentra en estado " & cRecDesc.mfgEstadoColocRecupDesc(lrs!nprdestado)
        Set cRecDesc = Nothing
        Exit Function
    End If
    
    
Set dObtieneDatosPagoCredRecup = lrs
Set lrs = Nothing
Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosPagoCredRecup>>", Err.Description
    
End Function


Public Function dObtieneDatosCancelaCredRecup(ByVal psCtaCod As String, ByRef psmensaje As String) As ADODB.Recordset
'Obtiene Datos para Cancelacion de Credito en Recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Dim cRecDesc As New DCOMColRecDesc

'Set lrs = New ADODB.Recordset

On Error GoTo dError

lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, P.nTasaInteres, C.nMontoCol, P.nSaldo, C.cUltimaActualizacion,  " _
    & " CRec.nSaldoIntComp, CRec.nSaldoIntMor, CRec.nSaldoGasto, CRec.nIntCompGen, CRec.dIngRecup,  " _
    & " CRec.nComisionCod, CRec.cMetLiquid, CRec.nTipCj, CRec.nDemanda, CRec.nNroCalen, IsNull(Com.nValor,0) nValorCom,  " _
    & " Pers.cPersCod, Pers.cPersNombre, Pers2.cPersCod cPersCodAbog, Pers2.cPersNombre cPersNombreAbog,  " _
    & " cEstudioAbogCod = (SELECT cPersCod FROM ProductoPersona  " _
    & "           WHERE cCtaCod ='" & psCtaCod & "' AND nPrdPersRelac = " & gColRelPersEstudioJuridico & " ) , " _
    & " nTasaInt = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
    & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntCompNormal & " )  " _
    & "  " _
    & "  " _
    & " ,nVisMetodo,nVisNegociacion,nVisCancelacion" _
    & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
    & " Inner Join ColocRecup CRec ON C.cctacod = CRec.cctacod " _
    & " Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular _
    & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod " _
    & " Inner Join ProductoPersona PP2 ON C.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersEstudioJuridico _
    & " Inner Join Persona Pers2 ON Pers2.cPersCod = PP2.cPersCod " _
    & " Left Join ColocRecupComision Com ON Com.nComisionCod = CRec.nComisionCod " _
    & " WHERE P.cCtaCod ='" & psCtaCod & "'"

Set lrs = coConex.CargaRecordSet(lsSQL)

    If Not (lrs.EOF And lrs.BOF) Then
        If lrs!nprdestado = gColocEstRecCanJud Or lrs!nprdestado = gColocEstRecCanCast Then
            psmensaje = "Credito ya se encuentra en estado " & cRecDesc.mfgEstadoColocRecupDesc(lrs!nprdestado)
            Exit Function
        End If
        
        Set dObtieneDatosCancelaCredRecup = lrs
        Set lrs = Nothing
        
    Else
        
        '20060505
        'si el recordset esta vacio debe devolver nothing
        Set dObtieneDatosCancelaCredRecup = Nothing
        
    End If
    

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosPagoCredRecup>>", Err.Description
    
End Function

Public Function dObtieneDatosCierreMes(ByRef psmensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim lrs As ADODB.Recordset
Set lrs = New ADODB.Recordset

On Error GoTo ErrorRecupera
    
lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, P.nTasaInteres, C.nMontoCol, P.nSaldo, C.cUltimaActualizacion,  " _
    & " CRec.nSaldoIntComp, CRec.nSaldoIntMor, CRec.nSaldoGasto, CRec.nIntCompGen, CRec.dIngRecup,  " _
    & " CRec.nTipCj, CRec.nDemanda, CRec.nNroCalen, " _
    & " nTasaIntComp = (SELECT ISNULL(nTasaInteres, 0) From ProductoTasaInteres  PT    WHERE PT.cCtaCod = C.cCtacod  and  nPrdTasaInteres=1 ),  " _
    & " nTasaIntMorat = (SELECT ISNULL(nTasaInteres, 0) From ProductoTasaInteres  PT    WHERE PT.cCtaCod = C.cCtacod and  nPrdTasaInteres=3)  " _
    & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
    & " Inner Join ColocRecup CRec ON C.cctacod = CRec.cctacod " _
    & " WHERE P.nPrdEstado in (" & gColocEstRecVigJud & "," & gColocEstRecVigCast & " ) "
    
Set lrs = coConex.CargaRecordSet(lsSQL)

    If lrs.EOF And lrs.BOF Then
            psmensaje = "No se encuentran Creditos en Recuperaciones "
            Exit Function
    End If
    
Set dObtieneDatosCierreMes = lrs
Set lrs = Nothing
    
Exit Function

ErrorRecupera:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
    
End Function

Public Function dObtieneDatosNegociaCuotas(ByVal psCodCta As String, ByVal psNroNeg As String, ByRef psmensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim lrs As ADODB.Recordset
Set lrs = New ADODB.Recordset

On Error GoTo ErrorRecupera
    
lsSQL = "SELECT * FROM ColocRecupNegocPlanPagos WHERE cCtaCod ='" & psCodCta & "' " & _
      " And cNroNeg = '" & psNroNeg & "' " & _
      " And cEstado ='P' ORDER BY nNroCuota "
    
Set lrs = coConex.CargaRecordSet(lsSQL)
    If lrs.EOF And lrs.BOF Then
            psmensaje = "No se encuentran Calendario de Negociacion "
            Exit Function
    End If
Set dObtieneDatosNegociaCuotas = lrs
Set lrs = Nothing
    
Exit Function

ErrorRecupera:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
End Function


Public Function bVerificarVisto(ByVal psCtaCod As String, pnTipoVisto As Integer) As Boolean
Dim lsSQL As String
Dim lrs As ADODB.Recordset
Set lrs = New ADODB.Recordset

On Error GoTo ErrorRecupera

'1=Metodo/2=Negociacion/3=Cancelacion
    
lsSQL = "SELECT * FROM ColocRecup where cCtaCod='" & psCtaCod & "' and "

If pnTipoVisto = 1 Then
    lsSQL = lsSQL & " nVisMetodo=1 "
ElseIf pnTipoVisto = 2 Then
    lsSQL = lsSQL & " nVisNegociacion=1"
ElseIf pnTipoVisto = 3 Then
    lsSQL = lsSQL & " nVisCancelacion=1"
End If
    
Set lrs = coConex.CargaRecordSet(lsSQL)
    If lrs.EOF And lrs.BOF Then
        bVerificarVisto = False
    Else
        bVerificarVisto = True
    End If
Set lrs = Nothing
    
Exit Function

ErrorRecupera:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
End Function


'**DAOR 20070414, Función que Obtiene la distribución de los montos CIMG para cobranza
'**de una determinada cuenta
Public Function dObtieneDistribucionCIMGCobranza(psCtaCod As String, pdFecha As Date) As ADODB.Recordset
Dim lsSQL As String

On Error GoTo ErrordObtieneDistribucionCIMGCobranza

    lsSQL = " Select * From ColocRecupDistribucionCobr  Where cCtaCod ='" & psCtaCod & "' " _
          & " AND dFecha='" & Format(pdFecha, "yyyymmdd") & "'"
           
    Set dObtieneDistribucionCIMGCobranza = coConex.CargaRecordSet(lsSQL)
    
    Exit Function

ErrordObtieneDistribucionCIMGCobranza:
    Err.Raise Err.Number, "Obtiene Distribución de Montos CIMG para Cobranza <<dObtieneDistribucionCIMGCobranza>>", Err.Description
    
End Function

'peac 20070912 reporte para pase de judiciales a castigados

Public Function dObtienegRepCredJudToCastigar(psCodAge As String, ByVal pdFecSis As Date, _
ByVal pnAtrIni As Integer, ByVal pnAtrFin As Integer, ByVal pMatAgencias As String, Optional ByVal pListaCreditos As String = "") As ADODB.Recordset

Dim sSql As String
Dim lrs As ADODB.Recordset
Dim sCriterio As String
Dim qAgencias As String, nAgencias As String, nCreditos As String

Dim I As Integer
Dim sCadAge As String

Set lrs = New ADODB.Recordset

On Error GoTo ErrorJudToCas

'    sCadAge = ""
'    For i = 0 To UBound(pMatAgencias) - 1
'        sCadAge = sCadAge & pMatAgencias(i) & ","
'    Next i
'    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

'    If pMatAgencias = "" Then
'        qAgencias = " where substring(cr.cctacod,4,2)='" & psCodAge & "' "
'    Else
'        qAgencias = " where substring(cr.cctacod,4,2) in (" & pMatAgencias & " ) "
'    End If


'*** PEAC 20080311
    If pMatAgencias = "" Then
        nAgencias = Replace(Replace(psCodAge, "'", ""), " ", "")
    Else
        nAgencias = Replace(Replace(pMatAgencias, "'", ""), " ", "")
    End If
    
    'MADM 20110513 pMatCreditos
    If pListaCreditos <> "" Then
        If Len(pListaCreditos) > 20 Then
            nCreditos = Replace(Replace(pListaCreditos, "'", ""), " ", "")
        Else
            nCreditos = Replace(Replace(pListaCreditos, "'", ""), " ", "")
        End If
    End If

'*** PEAC 20080311
'sSql = " select"
'sSql = sSql & "     CODAge=substring(cr.cctacod,4,2),"
'sSql = sSql & "     NomAge=(select cAgeDescripcion from agencias where cAgeCod=substring(cr.cctacod,4,2)),"
'sSql = sSql & "     Credito=cr.cctacod,"
'sSql = sSql & "     cMoneda=case substring(cr.cctacod,9,1) when '1' then 'NUEVOS SOLES' ELSE 'DOLARES US' END,"
'sSql = sSql & "     cMoney=substring(cr.cctacod,9,1),"
'sSql = sSql & "     Estado=(SELECT cConsDescripcion from CONSTANTE where nConsCod=3001 and nConsValor=p.nPrdEstado),"
'sSql = sSql & "     Analista=rh1.cUser,"
'sSql = sSql & "     Linea=cl.cDescripcion, "
'sSql = sSql & "     CodCliente=pp2.cPersCod,"
'sSql = sSql & "     Cliente=per.cPersNombre,"
'sSql = sSql & "     DNI=isnull(Pid.cpersIDnro,''),"
'sSql = sSql & "     DIRECCION=Per.cPersDireccDomicilio,"
'sSql = sSql & "     CodSbs=isnull(per.cPersCodSbs,''),"
'sSql = sSql & "     FechaVige=c.dVigencia,"
'sSql = sSql & "     Atraso=case when cc.nDiasAtraso < 0 then 0 else cc.nDiasAtraso end,"
'sSql = sSql & "     MontoDesem=c.nMontoCol,"
'sSql = sSql & "     ProviConst=IsNull(Cal.nProvision, 0),"
'sSql = sSql & "     SaldoCap=p.nsaldo,"
'sSql = sSql & "     nDiasTransUltPgo = datediff(d,substring(c.cUltimaActualizacion,1,8),'" & Format(pdFecSis, "mm/dd/yyyy") & "'),"
'sSql = sSql & "     TasaIntComp =(select isnull(nTasaInteres, 0)"
'sSql = sSql & "             From ProductoTasaInteres"
'sSql = sSql & "             Where cCtaCod = cr.cCtaCod"
'sSql = sSql & "             and nPrdTasaInteres = 1 ) ,"
'sSql = sSql & "     TasaIntMora =(select isnull(nTasaInteres, 0)"
'sSql = sSql & "             From ProductoTasaInteres"
'sSql = sSql & "             Where cCtaCod = cr.cCtaCod"
'sSql = sSql & "             and nPrdTasaInteres = 3 ),"
'sSql = sSql & "     IntCom=CR.nSaldoIntComp,"
'sSql = sSql & "     IntMor=CR.nSaldoIntMor,"
'sSql = sSql & "     Gastos=isnull((select sum(nMonto-nMontoPagado)SumGastos"
'sSql = sSql & "             from ColocRecupGastos gas"
'sSql = sSql & "             where cCtaCod=CR.cCtaCod and nColocRecGastoEstado=0),0),"
'sSql = sSql & "     Calificacion = IsNull(Cal.cCalCmac, 0)"
'sSql = sSql & "     from colocRecup cr"
'sSql = sSql & "     inner join producto p on cr.cctacod = p.cctacod and p.nPrdEstado in (2201,2205)"
'sSql = sSql & "     inner join productopersona pp1 on cr.cctacod=pp1.cctacod and pp1.nPrdPersRelac=28"
'sSql = sSql & "     inner join productopersona pp2 on cr.cctacod=pp2.cctacod and pp2.nPrdPersRelac=20"
'sSql = sSql & "     inner join rrhh rh1 on pp1.cperscod=rh1.cperscod"
'sSql = sSql & "     inner join persona per on pp2.cperscod=per.cperscod"
'sSql = sSql & "     inner join persId PId on per.cperscod = pid.cperscod and pid.cpersidtpo='1'"
'sSql = sSql & "     inner join colocaciones c on cr.cctacod=c.cctacod"
'sSql = sSql & "     inner join colocLineaCredito cl on c.cLineaCred=cl.cLineaCred"
'sSql = sSql & "     inner join ColocacCred cc on cr.cctacod=cc.cctacod and cc.nDiasAtraso between " & pnAtrIni & " and " & pnAtrFin
'sSql = sSql & "     inner join dbConsolidada..ColocCalifProvTotal Cal on Cr.cCtaCod=Cal.cCtaCod"
'sSql = sSql & "     and datediff(day,Cal.dFecha,(select max(dFecha)from dbConsolidada..ColocCalifProvTotal))=0"
'sSql = sSql & " " & qAgencias
'sSql = sSql & "order by codage,cmoney"

'*** PEAC 20080311
If nCreditos <> "" Then
    sSql = "exec stp_sel_RepCredJudToCastigarExcel '" & nCreditos & "'" 'MADM 20110513
Else
    sSql = "exec stp_sel_RepCredJudToCastigar  " & pnAtrIni & "," & pnAtrFin & ",'" & nAgencias & "'"
End If
Set lrs = coConex.CargaRecordSet(sSql)
   
Set dObtienegRepCredJudToCastigar = lrs
Set lrs = Nothing
    
Exit Function

ErrorJudToCas:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
End Function

'peac 20070917 reporte para cobranza de gestores

Public Function dObtienegRepCobranzaGestores( _
psCodAge As String, _
ByVal pMatAgencias As String, _
ByVal pnAtrIni As Integer, _
ByVal pnAtrFin As Integer, _
ByVal pdFecIni As Date, _
ByVal pdFecFin As Date) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim I As Integer
Dim lrs As ADODB.Recordset
Dim qAgencias As String, nAgencias As String

Set lrs = New ADODB.Recordset

'    If pMatAgencias = "" Then
'        qAgencias = " and substring(cc.cctacod,4,2)='" & psCodAge & "' "
'    Else
'        qAgencias = " and substring(cc.cctacod,4,2) in (" & pMatAgencias & " ) "
'    End If

If pMatAgencias = "" Then
        nAgencias = Replace(Replace(psCodAge, "'", ""), " ", "")
    Else
        nAgencias = Replace(Replace(pMatAgencias, "'", ""), " ", "")
End If

'ssql = " select cc.*"
'ssql = ssql & " from (select substring(a.cctacod,4,2)Agencia,substring(a.cctacod,9,1)Moneda,case substring(a.cctacod,9,1) when '1' then 'NUEVOS SOLES' ELSE 'DOLARES US' END Moneey,rh.cuser,a.cctacod,a.nCuota,p.cpersnombre,a.nColocCalendEstado, a.dpago,a.nnrocalen,ncoloccalendapl,"
'ssql = ssql & "     capital=isnull((select sum(nMontoPagado)"
'ssql = ssql & "         From coloccalenddet"
'ssql = ssql & "         Where nnrocalen = a.nnrocalen And ncoloccalendapl = a.ncoloccalendapl"
'ssql = ssql & "         and nCuota=a.nCuota and cctacod=a.cctacod and nPrdConceptoCod=1000),0),"
'ssql = ssql & "     Interes=isnull((select sum(nMontoPagado)"
'ssql = ssql & "         From coloccalenddet"
'ssql = ssql & "         Where nnrocalen = a.nnrocalen And ncoloccalendapl = a.ncoloccalendapl"
'ssql = ssql & "         and nCuota=a.nCuota and cctacod=a.cctacod and nPrdConceptoCod in (1100,1105)),0),"
'ssql = ssql & "     Mora=isnull((select sum(nMontoPagado)"
'ssql = ssql & "         From coloccalenddet"
'ssql = ssql & "         Where nnrocalen = a.nnrocalen And ncoloccalendapl = a.ncoloccalendapl"
'ssql = ssql & "         and nCuota=a.nCuota and cctacod=a.cctacod and nPrdConceptoCod=1101),0),"
'ssql = ssql & "     Gastos=isnull((select sum(nMontoPagado)"
'ssql = ssql & "         From coloccalenddet"
'ssql = ssql & "         Where nnrocalen = a.nnrocalen And ncoloccalendapl = a.ncoloccalendapl"
'ssql = ssql & "         and nCuota=a.nCuota and cctacod=a.cctacod and nPrdConceptoCod like '12%'),0),"
'ssql = ssql & "     Atraso = DateDiff(d, dVenc, dpago)"
'ssql = ssql & "     from coloccalendario a"
'ssql = ssql & " inner join productopersona pp1 on a.cctacod=pp1.cctacod and pp1.nPrdPersRelac=20"
'ssql = ssql & " inner join productopersona pp2 on a.cctacod=pp2.cctacod and pp2.nPrdPersRelac=28"
'ssql = ssql & " inner join persona p on pp1.cperscod=p.cperscod"
'ssql = ssql & " inner join rrhh rh on pp2.cperscod=rh.cperscod) cc"
'ssql = ssql & " inner join colocaccred b on cc.cctacod=b.cctacod and cc.nnrocalen=b.nnrocalen"
''ALPA 20090312*******************************************
''ssql = ssql & " where cc.dpago between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "'"
'ssql = ssql & " where (datediff(d,cc.dpago,'" & Format(pdFecIni, "yyyymmdd") & "')<=0 and datediff(d,cc.dpago,'" & Format(pdFecFin, "yyyymmdd") & "')>=0) "
''********************************************************
'ssql = ssql & " and cc.nColocCalendEstado in (1,0) and cc.ncoloccalendapl=1"
'ssql = ssql & " and cc.atraso between " & pnAtrIni & " and " & pnAtrFin
'ssql = ssql & " " & qAgencias
'ssql = ssql & " order by cc.agencia,cc.moneda,cc.cuser,cc.cctacod,cc.ncuota"
    
sSql = "exec stp_sel_ReporteGestores " & pnAtrIni & "," & pnAtrFin & ",'" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & nAgencias & "'"
Set lrs = coConex.CargaRecordSet(sSql)
   
Set dObtienegRepCobranzaGestores = lrs
Set lrs = Nothing
    
Exit Function
    
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set dObtienegRepCobranzaGestores = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
       
End Function

'peac 20071010 reporte de gastos de cred jud y cas

Public Function dObtienegRepGastosJudCas( _
psCodAge As String, _
ByVal pEstado As String, _
ByVal pMoneda As String, _
ByVal pMatAgencias As String, _
ByVal pdFecIni As Date, _
ByVal pdFecFin As Date, _
ByVal pMatProd As Variant) As ADODB.Recordset

'psCodAge, pEstado, pMoneda, psListaAgencias, pdDel, pdAl, pMatProd

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim I As Integer
Dim lrs As ADODB.Recordset
Dim qAgencias As String, qEstado As String, sCadProd As String, qmoneda As String

Set lrs = New ADODB.Recordset

    sCadProd = " And Co.cTpoCredCod in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ") "
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If

    If pEstado = "1" Then
        qEstado = " a.nPrdEstado = 2201 "
    Else
        qEstado = " a.nPrdEstado = 2202 "
    End If
    
    If pMoneda = "1" Then
        qmoneda = " and substring(a.cctacod,9,1) = '1' "
    Else
        qmoneda = " and substring(a.cctacod,9,1) = '2' "
    End If
    
    If pMatAgencias = "" Then
        qAgencias = " and Co.cAgeCodAct='" & psCodAge & "' "
    Else
        qAgencias = " and Co.cAgeCodAct in (" & pMatAgencias & " ) "
    End If

sSql = " select  Co.cAgeCodAct Agencia,i.cAgeDescripcion,h.cConsDescripcion Estado,"
sSql = sSql & "     case substring(a.cctacod,9,1) when '1' then 'NUEVOS SOLES' ELSE 'DOLARES US' END Moneda,"
sSql = sSql & "     Co.cTpoCredCod Producto, j.cConsDescripcion NomProd,isnull(cCtaCodAnt,'')cCtaCodAnt,a.cctacod,d.cpersnombre,"
sSql = sSql & "     b.nPrdConceptoCod,b.cmotivogasto,b.nmonto,b.dasigna,b.nmontopagado,"
sSql = sSql & "     isnull(substring(g.cMovNro,1,8),'')Pago,b.nmonto-b.nmontopagado Saldo"
sSql = sSql & " from producto a"
sSql = sSql & " inner join Colocaciones Co on Co.cCtaCod = a.cCtaCod "
sSql = sSql & " inner join constante h on a.nprdestado=h.nConsValor and h.nconscod=3001"
sSql = sSql & " inner join constante j on Co.cTpoCredCod=j.nConsValor and j.nconscod=3034"
sSql = sSql & " inner join agencias i on Co.cAgeCodAct =i.cAgeCod"
sSql = sSql & " inner join colocRecupGastos b on a.cctacod=b.cctacod and b.dasigna between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "' "
sSql = sSql & " inner join productopersona c on a.cctacod=c.cctacod and nprdpersrelac=20"
sSql = sSql & " inner join persona d on c.cperscod=d.cperscod"
sSql = sSql & " left join relcuentas e on a.cctacod=e.cctacod"
sSql = sSql & " left join MOVCOLDET f on b.cctacod=f.cctacod and b.nnrogastocta=f.nnrocuota and f.nprdconceptocod=b.nPrdConceptoCod"
sSql = sSql & " left join mov g on f.nMovNro=g.nMovNro"
sSql = sSql & " where " 'a.nprdestado like '220[12]'
sSql = sSql & " " & qEstado & qmoneda & qAgencias & sCadProd
sSql = sSql & " order by Agencia,producto,a.cctacod,b.dasigna"

Set lrs = coConex.CargaRecordSet(sSql)
   
Set dObtienegRepGastosJudCas = lrs
Set lrs = Nothing
    
Exit Function
           
End Function

'peac 20071011 reporte de creditos con saldos condonados

Public Function dObtienegRepCredSaldosCondo( _
psCodAge As String, _
ByVal pMatAgencias As String, _
ByVal pdFecIni As Date, _
ByVal pdFecFin As Date) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim I As Integer
Dim lrs As ADODB.Recordset
Dim qAgencias As String, qEstado As String, sCadProd As String, qmoneda As String

Set lrs = New ADODB.Recordset
   
    If pMatAgencias = "" Then
        qAgencias = " and substring(p.cctacod,4,2)='" & psCodAge & "' "
    Else
        qAgencias = " and C.cAgeCodAct in (" & pMatAgencias & " ) " ' BRGO BASILEA II
    End If

sSql = " select  CodAge=C.cAgeCodAct,"
sSql = sSql & "     Agencia=upper(ag.cAgeDescripcion),"
sSql = sSql & "     Moneda=case substring(p.cctacod,9,1) when '1' then 'NUEVOS SOLES' ELSE 'DOLARES US' END,"
sSql = sSql & "     CtaAnt=isnull(rc.cctacodant,space(18)),"
sSql = sSql & "     CtaAct=p.cctacod,"
sSql = sSql & "     per.cpersnombre,"
sSql = sSql & "     Analista=isnull(rh.cuser,space(4)),"
sSql = sSql & "     Capital=P.nSaldo,"
sSql = sSql & "     Interes=cr.nSaldoIntComp+cr.nIntCompGen,"
sSql = sSql & "     Mora=cr.nSaldoIntMor,"
sSql = sSql & "     Gastos=cr.nSaldoGasto,"
sSql = sSql & "     Total=P.nSaldo+cr.nSaldoIntComp+cr.nIntCompGen+cr.nSaldoIntMor+cr.nSaldoGasto,"
sSql = sSql & "     Fecha=substring(M.cMovNro,7,2)+'/'+substring(M.cMovNro,5,2)+'/'+substring(M.cMovNro,1,4),"
sSql = sSql & "     estado = upper(cons.cConsDescripcion)"
sSql = sSql & " from Mov M"
sSql = sSql & "     inner join MovCol MC on M.nMovNro=MC.nMovNro"
sSql = sSql & "     inner join Producto P on MC.cCtaCod=P.cCtaCod"
sSql = sSql & "     inner join Colocaciones C on C.cCtaCod = P.cCtaCod " ' BRGO BASILEA II
sSql = sSql & "     inner join agencias ag on C.cAgeCodAct = ag.cAgeCod" ' BRGO BASILEA II
sSql = sSql & "     inner join constante cons on p.nPrdEstado=cons.nConsValor and cons.nConsCod=3001"
sSql = sSql & "     inner join relcuentas rc on p.cctacod=rc.cctacod"
sSql = sSql & "     inner join productopersona pp1 on p.cctacod=pp1.cctacod and pp1.nprdpersrelac=20"
sSql = sSql & "     inner join productopersona pp2 on p.cctacod=pp2.cctacod and pp2.nprdpersrelac=28"
sSql = sSql & "     inner join persona per on pp1.cperscod=per.cperscod"
sSql = sSql & "     left join rrhh rh on pp2.cperscod=rh.cperscod"
sSql = sSql & "     inner join ColocRecup CR on P.cCtaCod=CR.cCtaCod"
sSql = sSql & " where M.cOpeCod='130600' and M.nMovFlag=0 and substring(m.cMovNro,1,8) between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "' "
sSql = sSql & " " & qAgencias
sSql = sSql & " ORDER BY codage,MONEDA,substring(M.cMovNro,1,8)"

Set lrs = coConex.CargaRecordSet(sSql)
   
Set dObtienegRepCredSaldosCondo = lrs
Set lrs = Nothing
    
Exit Function
           
End Function

'peac 20071023 reporte de cartera de creditos judi y casti

Public Function dObtienegRepCarteraJudCas( _
psCodAge As String, _
ByVal pEstado As String, _
ByVal pMoneda As String, _
ByVal pMatAgencias As String, _
ByVal pdFecIni As Date, _
ByVal pdFecFin As Date, _
ByVal pMatProd As Variant) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim I As Integer
Dim lrs As ADODB.Recordset
Dim qAgencias As String, qEstado As String, sCadProd As String, qmoneda As String

Set lrs = New ADODB.Recordset
   
    If pEstado = "1" Then
        qEstado = " p.nprdestado in (2201,2205) "
    Else
        qEstado = " p.nprdestado in (2202,2206) "
    End If
    
    If pMoneda = "1" Then
        qmoneda = " and substring(p.cctacod,9,1) = '1' "
    Else
        qmoneda = " and substring(p.cctacod,9,1) = '2' "
    End If
    
    If pMatAgencias = "" Then
        qAgencias = " and substring(p.cctacod,4,2)='" & psCodAge & "' "
    Else
        qAgencias = " and substring(p.cctacod,4,2) in (" & pMatAgencias & " ) "
    End If

    sCadProd = " And SUBSTRING(p.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ") "
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If

sSql = " select  case substring(p.cCtaCod,9,1) when '1' then 'NUEVOS SOLES' ELSE 'DOLARES US' END Moneda, p.cCtaCod,"
sSql = sSql & "     isnull(upper(age.cAgeDescripcion),space(27)) Agencia,"
sSql = sSql & "     isnull(cons.cConsDescripcion,space(37)) Producto,"
sSql = sSql & "     isnull(upper(est.cConsDescripcion),space(33)) Estado,"
sSql = sSql & "     per.cPersNombre,"
sSql = sSql & "     per.cPersDireccDomicilio,"
sSql = sSql & "     cc.nDiasAtraso,"
sSql = sSql & "     p.nsaldo Capital,"
sSql = sSql & "     cr.nSaldoIntComp+cr.nIntCompGen Interes,"
sSql = sSql & "     cr.nSaldoIntMor Mora,"
sSql = sSql & "     cr.nSaldoGasto Gastos,"
sSql = sSql & "     p.nsaldo+cr.nSaldoIntComp+cr.nIntCompGen+cr.nSaldoIntMor+cr.nSaldoGasto Total"
sSql = sSql & " from producto p"
sSql = sSql & " left join constante cons on substring(p.cCtaCod,6,3)=cons.nConsValor and cons.nConsCod=1001"
sSql = sSql & " left join agencias age on substring(p.cCtaCod,4,2)=age.cAgeCod"
sSql = sSql & " left join constante est on p.nPrdEstado=est.nConsValor and est.nConsCod=3001"
sSql = sSql & " inner join colocaccred cc on p.cctacod=cc.cctacod"
sSql = sSql & " inner join colocRecup cr on p.cctacod=cr.cctacod"
sSql = sSql & " inner join productopersona pp on p.cctacod=pp.cctacod and pp.nPrdPersRelac='20'"
sSql = sSql & " inner join persona per on pp.cperscod=per.cperscod"
sSql = sSql & " where"
sSql = sSql & " " & qEstado & qmoneda & qAgencias & sCadProd
sSql = sSql & " order by substring(p.cCtaCod,4,2),substring(p.cCtaCod,6,3),p.nPrdEstado,per.cPersNombre --agencia,producto,estado,cliente"

Set lrs = coConex.CargaRecordSet(sSql)
   
Set dObtienegRepCarteraJudCas = lrs
Set lrs = Nothing
    
Exit Function
           
End Function


'JACA 20110719*****************************************************************
Public Function getCuentaEmbargo(ByVal psCtaCod As String) As Recordset
    Dim sql As String

    On Error GoTo ErrorgetCuentaEmbargo

    sql = "exec stp_sel_CuentaEmbargo '" & psCtaCod & "' "
           
    Set getCuentaEmbargo = coConex.CargaRecordSet(sql)
    
    Exit Function

ErrorgetCuentaEmbargo:
    Err.Raise Err.Number, "Obtener Datos de Cuenta a Embargo", Err.Description
    
End Function
Public Function getEstadoBienEmbargo() As Recordset
    Dim sql As String

    On Error GoTo ErrorgetEstadoBienEmbargo

    sql = "exec stp_sel_EstadoBienEmbargo"
           
    Set getEstadoBienEmbargo = coConex.CargaRecordSet(sql)
    
    Exit Function

ErrorgetEstadoBienEmbargo:
    Err.Raise Err.Number, "Obtener Estado del Bien a Embargar", Err.Description
    
End Function
Public Function getConsValorEmbargo(ByVal nConsCod As Integer, ByVal nConsValorLeft As String, ByVal nConsValorRight As String) As Recordset
    Dim sql As String

    On Error GoTo ErrorgetConsValorEmbargo

    sql = "exec stp_sel_ConsValorEmbargo " & nConsCod & ",'" & nConsValorLeft & "','" & nConsValorRight & "'"
           
    Set getConsValorEmbargo = coConex.CargaRecordSet(sql)
    
    Exit Function

ErrorgetConsValorEmbargo:
    Err.Raise Err.Number, "Obtener Valores Constante de Embargo", Err.Description
    
End Function
Public Function getAlmacenEmbargo(ByVal psCodAge As String, ByVal psCodAlmacen As String) As Recordset
    Dim sql As String

    On Error GoTo ErrorgetAlmacenEmbargo

    sql = "exec stp_sel_AlmacenEmbargo '" & psCodAge & "','" & psCodAlmacen & "'"
           
    Set getAlmacenEmbargo = coConex.CargaRecordSet(sql)
    
    Exit Function

ErrorgetAlmacenEmbargo:
    Err.Raise Err.Number, "Obtener Almacenes de Embargo", Err.Description
    
End Function
Public Function getAlmacenCorrelativo(ByVal psCodAge As String) As String
    Dim sql As String
    Dim rs As Recordset
    On Error GoTo ErrorgetAlmacenCorrelativo

    sql = "exec stp_sel_AlmEmbgo_Correlativo '" & psCodAge & "'"
    
    Set rs = coConex.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) And Not IsNull(rs!Correlativo) Then
        getAlmacenCorrelativo = rs!Correlativo
    Else
        getAlmacenCorrelativo = "001"
    End If
    
    Exit Function

ErrorgetAlmacenCorrelativo:
    Err.Raise Err.Number, "Obtener Correlativo de Almacenes", Err.Description
    
End Function
Public Sub insertarAlmacenEmbargo(ByVal psCodAlmacen As String, ByVal psNomAlmacen As String, ByVal psDescripcion As String, ByVal psNroMov As String)
    Dim sql As String

    On Error GoTo ErrorinsertarAlmacenEmbargo

    sql = "exec stp_ins_AlmacenEmbargo '" & psCodAlmacen & "','" & psNomAlmacen & "','" & psDescripcion & "','" & psNroMov & "'"
    coConex.CargaRecordSet (sql)
    
    Exit Sub

ErrorinsertarAlmacenEmbargo:
    Err.Raise Err.Number, "Guardar Almacenes", Err.Description
End Sub
Public Function getBienesEmbargo(ByVal psSubTpoBien As String) As Recordset
    Dim sql As String

    On Error GoTo ErrorgetBienesEmbargo

    sql = "exec stp_sel_BienesEmbargo '" & psSubTpoBien & "'"
    Set getBienesEmbargo = coConex.CargaRecordSet(sql)
    
    Exit Function

ErrorgetBienesEmbargo:
    Err.Raise Err.Number, "Obtener Bienes", Err.Description
End Function
Public Sub insertarBienEmbargo(ByVal pnConsValor As Integer, ByVal pnCodBien As Integer, ByVal psNomBien As String, ByVal psNroMov As String)
    Dim sql As String

    On Error GoTo ErrorinsertarBienEmbargo

    sql = "exec stp_ins_BienEmbargo " & pnConsValor & "," & pnCodBien & ",'" & psNomBien & "','" & psNroMov & "'"
    coConex.CargaRecordSet (sql)
    
    Exit Sub

ErrorinsertarBienEmbargo:
    Err.Raise Err.Number, "Guardar Bien", Err.Description
End Sub
Public Function getBienCorrelativo(ByVal pnConsValor As Integer) As Integer
    Dim sql As String
    Dim rs As Recordset
    On Error GoTo ErrorgetBienCorrelativo

    sql = "exec stp_sel_BienEmbargo_Correlativo " & pnConsValor & ""
    
    Set rs = coConex.CargaRecordSet(sql)
    getBienCorrelativo = rs!Correlativo
    
    Exit Function

ErrorgetBienCorrelativo:
    Err.Raise Err.Number, "Obtener Correlativo de Bienes", Err.Description
    
End Function
Public Sub actualizarBienEmbargo(ByVal pnConsValor As Integer, ByVal pnCodBien As Integer)
    Dim sql As String

    On Error GoTo ErroractualizarBienEmbargo

    sql = "exec stp_upd_BienEmbargo " & pnConsValor & "," & pnCodBien & ""
    coConex.CargaRecordSet (sql)
    
    Exit Sub

ErroractualizarBienEmbargo:
    Err.Raise Err.Number, "Actualizar Bien", Err.Description
End Sub
Public Sub actualizarAlmacenEmbargo(ByVal psCodAlmacen As String)
    Dim sql As String

    On Error GoTo ErroractualizarAlmacenEmbargo

    sql = "exec stp_upd_AlmacenEmbargo '" & psCodAlmacen & "'"
    coConex.CargaRecordSet (sql)
    
    Exit Sub

ErroractualizarAlmacenEmbargo:
    Err.Raise Err.Number, "Actualizar Almacenes", Err.Description
End Sub
Public Sub insertarEmbargo(ByVal pnMovNro As Long, ByVal psCtaCod As String, ByVal pnCapital As Currency, ByVal pnInteres As Currency, ByVal pnGastos As Currency, _
                          ByVal psNroExpediente As String, ByVal psNroResolucion As String, ByVal pdFecEmbargo As Date, ByVal psUltimaActualizacion As String)
    Dim sql As String

    On Error GoTo ErrorinsertarEmbargo

    sql = "exec stp_ins_Embargo " & pnMovNro & ",'" & psCtaCod & "'," & pnCapital & "," & pnInteres & "," & pnGastos & ",'" & psNroExpediente & "','" & psNroResolucion & "','" & Format(pdFecEmbargo, "yyyymmdd") & "','" & psUltimaActualizacion & "'"
    coConex.CargaRecordSet (sql)
    
    Exit Sub

ErrorinsertarEmbargo:
    Err.Raise Err.Number, "Guardar Embargo", Err.Description
End Sub
Public Sub insertarEmbargoDetalle(ByVal psCtaCod As String, ByVal pnNroBienEmbargo As Integer, ByVal psDocSalida As String, ByVal pdFecEmbargo As Date, _
                                ByVal pnCodSubTpoBien As Integer, ByVal pnCodBien As Integer, ByVal pnEstadoBien As Integer, ByVal pnEstadoEmbargo As Integer, ByVal psCodAlmacen As String, _
                                ByVal psBienDescripcion As String, ByVal pnCantidad As String, ByVal psColor As String, ByVal psMarca As String, ByVal psModelo As String, _
                                ByVal psSerie As String, ByVal psMotor As String, ByVal pnTasacion As Currency, ByVal psPlaca As String, ByVal psPartida As String, ByVal psUltimaActualizacion As String, _
                                Optional pdFecSalida As String = "", Optional pdFecTasacion As String = "", Optional psMonTasacion As String = "", _
                                Optional psDepositario As String = "", Optional psTitularBien As String = "")
    Dim sql As String

    On Error GoTo ErrorinsertarEmbargoDetalle

    sql = "exec stp_ins_EmbargoDetalle  '" & psCtaCod & "'," & pnNroBienEmbargo & ",'" & psDocSalida & "' , '" & Format(pdFecEmbargo, "yyyymmdd") & "','" & IIf(pdFecSalida <> "", Format(pdFecSalida, "yyyymmdd"), pdFecSalida) & "' ,"
    sql = sql + "" & pnCodSubTpoBien & " ," & pnCodBien & "," & pnEstadoBien & "," & pnEstadoEmbargo & ",'" & psCodAlmacen & "' ,"
    sql = sql + "'" & psBienDescripcion & "' ,'" & pnCantidad & "','" & psColor & "','" & psMarca & "','" & psModelo & "',"
    sql = sql + "'" & psSerie & "','" & psMotor & "' ," & pnTasacion & ",'" & IIf(pdFecTasacion <> "", Format(pdFecTasacion, "yyyymmdd"), pdFecTasacion) & "','" & psMonTasacion & "','" & psPlaca & "','" & psPartida & "' ,'" & psUltimaActualizacion & "', "
    sql = sql + "'" & psDepositario & "','" & psTitularBien & "' "
    coConex.CargaRecordSet (sql)
    
    Exit Sub

ErrorinsertarEmbargoDetalle:
    Err.Raise Err.Number, "Guardar Bienes de Embargo", Err.Description
End Sub
'*** PEAC 20130618
'Public Function getEmbargosListar(ByVal pdFecha As Date) As Recordset
Public Function getEmbargosListar(ByVal psFecDel As String, ByVal psFecAl As String, ByVal psNumCta As String, ByVal psCodCli As String, ByVal psNroExp As String, ByVal psNroRes As String) As Recordset
    
    Dim sql As String
    Dim rs As Recordset
    On Error GoTo ErrorgetEmbargosListar

    'sql = "exec stp_sel_Embargos_Listar '" & Format(pdFecha, "yyyymmdd") & "'"
    sql = "exec stp_sel_Embargos_Listar '" & psFecDel & "','" & psFecAl & "','" & psNumCta & "','" & psCodCli & "','" & psNroExp & "','" & psNroRes & "'"
    
    Set rs = coConex.CargaRecordSet(sql)
    Set getEmbargosListar = rs
    
    Exit Function

ErrorgetEmbargosListar:
    Err.Raise Err.Number, "Obtener Listado de embargos", Err.Description
    
End Function
Public Function getEmbargo(ByVal psCtaCod As String) As Recordset
    Dim sql As String
    Dim rs As Recordset
    On Error GoTo ErrorgetEmbargo

    sql = "exec stp_sel_Embargo '" & psCtaCod & "'"
    
    Set rs = coConex.CargaRecordSet(sql)
    Set getEmbargo = rs
    
    Exit Function

ErrorgetEmbargo:
    Err.Raise Err.Number, "Obtener Embargo", Err.Description
    
End Function
Public Function getEmbargoDetalle(ByVal psCtaCod As String) As Recordset
    Dim sql As String
    Dim rs As Recordset
    On Error GoTo ErrorgetEmbargoDetalle

    sql = "exec stp_sel_EmbargoDetalle '" & psCtaCod & "'"
    
    Set rs = coConex.CargaRecordSet(sql)
    Set getEmbargoDetalle = rs
    
    Exit Function

ErrorgetEmbargoDetalle:
    Err.Raise Err.Number, "Obtener Embargo Detalle", Err.Description
    
End Function
Public Sub actualizarEmbargo(ByVal psCtaCod As String, Optional pnNroBienEmbargo As Integer = 0)
    Dim sql As String

    On Error GoTo ErroractualizarEmbargo

    sql = "exec stp_upd_Embargo '" & psCtaCod & "'," & pnNroBienEmbargo & ""
    coConex.CargaRecordSet (sql)
    
    Exit Sub

ErroractualizarEmbargo:
    Err.Raise Err.Number, "Actualizar Embargo", Err.Description
End Sub
Public Function getVerificarEmbargo(ByVal psCtaCod As String) As Recordset
    Dim sql As String
    Dim rs As Recordset
    On Error GoTo ErrorgetVerificarEmbargo

    sql = "exec stp_sel_VerificarEmbargo '" & psCtaCod & "'"
    
    Set rs = coConex.CargaRecordSet(sql)
    Set getVerificarEmbargo = rs
    
    Exit Function

ErrorgetVerificarEmbargo:
    Err.Raise Err.Number, "Verificar Embargo", Err.Description
    
End Function
Public Function getDistribucionSaldos(ByVal psCtaCod As String) As Recordset
    Dim sql As String
    Dim rs As Recordset
    On Error GoTo ErrorgetDistribucionSaldos

    sql = "exec stp_sel_DistribucionSaldos'" & psCtaCod & "'"
    
    Set rs = coConex.CargaRecordSet(sql)
    Set getDistribucionSaldos = rs
    
    Exit Function

ErrorgetDistribucionSaldos:
    Err.Raise Err.Number, "Obtener Distribucion Saldos", Err.Description
    
End Function
Public Sub insertarPagoCredAdjudicacion(ByVal psOpeCod As String, ByVal psCtaCod As String, ByVal pnMonto As Currency, ByVal pdFecha As Date, ByVal pnMoneda As Integer, ByVal psMovNro As String)
     Dim sql As String

    On Error GoTo ErrorinsertarCredAdjudicacion

    sql = "exec stp_ins_PagoCredAdjudicacion '" & psOpeCod & "','" & psCtaCod & "'," & pnMonto & ",'" & Format(pdFecha, "yyyy/mm/dd") & "'," & pnMoneda & ",'" & psMovNro & "'"
    coConex.CargaRecordSet (sql)
    
    Exit Sub

ErrorinsertarCredAdjudicacion:
    Err.Raise Err.Number, "Guardar Pago de Credito x Adjudicacion", Err.Description
End Sub
Public Function getPagoCredAdjudicacion(ByVal psCtaCod As String, ByVal psOpeCod As String) As Recordset
    Dim sql As String

    On Error GoTo ErrorgetCuentaEmbargo

    sql = "exec stp_sel_PagoCredAdjudicacion '" & psCtaCod & "','" & psOpeCod & "' "
           
    Set getPagoCredAdjudicacion = coConex.CargaRecordSet(sql)
    
    Exit Function

ErrorgetCuentaEmbargo:
    Err.Raise Err.Number, "Obtener Datos de Pago de Credito x Adjudicacion", Err.Description
    
End Function
Public Function getNroBienesEmbargados(ByVal psCtaCod As String) As Integer
    Dim sql As String
    Dim rs As Recordset
    On Error GoTo ErrorRep

    Set rs = New Recordset
    sql = "exec stp_sel_NroBienesEmbargados '" & psCtaCod & "'"
           
    Set rs = coConex.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) Then
        getNroBienesEmbargados = rs!NroBienes
    Else
         getNroBienesEmbargados = 0
    End If
    
    
    Exit Function

ErrorRep:
    Err.Raise Err.Number, "Obtener Nro de Bienes", Err.Description
    
End Function
Public Function getAgenciaCiudad(ByVal psAgeCod As String) As String
    Dim sql As String
    Dim rs As Recordset
    Dim nPosicion As Integer
    On Error GoTo ErrorRep

    Set rs = New Recordset
    sql = "exec stp_sel_AgenciaCiudad '" & psAgeCod & "'"
           
    Set rs = coConex.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) Then
        If rs!Ciudad <> "" Then
            nPosicion = InStr(rs!Ciudad, "(")
        
            If nPosicion > 0 Then
                getAgenciaCiudad = Left(rs!Ciudad, nPosicion - 1)
            Else
                getAgenciaCiudad = ""
            End If
            
            
        Else
             getAgenciaCiudad = ""
        End If
    Else
         getAgenciaCiudad = ""
    End If
    
    
    Exit Function

ErrorRep:
    Err.Raise Err.Number, "Obtener Nro de Bienes", Err.Description
    
End Function

Public Function InsertaMovPendientesRend(ByVal pnMovNro As Long, ByVal psCtaContCod As String, ByVal pnSaldo As Currency) As Integer
    On Error GoTo InsertaMovPendientesRendErr
    Dim rs As Recordset
    Dim sql As String
    Set rs = New Recordset
    InsertaMovPendientesRend = 1
    sql = "SELECT nMovNro FROM MovPendientesRend Where nMovNro = " & pnMovNro & " and cCtaContCod = '" & psCtaContCod & "' "
    Set rs = coConex.CargaRecordSet(sql)
    If rs.EOF Then
        sql = "INSERT INTO MovPendientesRend (nMovNro, cCtaContCod, nSaldo ) " _
             & "VALUES (" & pnMovNro & ", '" & psCtaContCod & "'," & pnSaldo & ") "
    Else
        sql = "UPDATE MovPendientesRend SET nSaldo = " & pnSaldo & " WHERE nMovNro = " & pnMovNro & " "
    End If
    coConex.Ejecutar sql
    InsertaMovPendientesRend = 0
    Exit Function
InsertaMovPendientesRendErr:
    Err.Raise Err.Number, "Insertar Movimiento Pendiente", Err.Description
End Function
'JACA END*********************************************************************

'*** PEAC 20120710
Public Function ObtieneVisitaDeGestores(ByVal psCodCta As String) As ADODB.Recordset

    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
    lsSQL = "exec stp_sel_VisitasDeGestoresRealizadas '" & psCodCta & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set ObtieneVisitaDeGestores = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
    Err.Raise Err.Number, "Error <<ObtieneVisitaDeGestores>>", Err.Description
End Function


'*** PEAC 20120711
Public Sub AnulaRegVisitaGestores(ByVal pnMovNro As Long)
    
    Dim sql As String

    On Error GoTo ErrorAnulaRegVisitaGestores

    sql = "exec stp_upd_AnulaRegVisitaGestores " & pnMovNro
    coConex.CargaRecordSet (sql)
    
    Exit Sub

ErrorAnulaRegVisitaGestores:
    Err.Raise Err.Number, "Anula Reg Visita Gestores.", Err.Description
End Sub

' *** RIRO 20131023 SEGUN TI-ERS146-2013 ***
Public Sub actualizarRelacionVoucherCaptacion(ByVal pnMovNroOpe As Long, _
                                              ByVal pnMovNroRVD As Long)
    
    On Error GoTo Error
    
    Dim lsSQL As String
    lsSQL = "exec stp_upd_ActualizarCapVoucherDeposito_2  " & pnMovNroOpe & ", " & pnMovNroRVD & ""
    coConex.CargaRecordSet lsSQL
    
    Exit Sub
    
Error:
    Err.Raise Err.Number, "Anula Reg Visita Gestores.", Err.Description
    
End Sub
' *** FIN RIRO
'RIRO20140610 ERS017 ************
Public Sub ActualizaMovPendientesRend(ByVal nMovNroPend As Long, ByVal pnMonto As Currency)
    Dim lsSQL As String
        
    On Error GoTo Error
        
    lsSQL = " UPDATE MovPendientesRend " _
         & " Set nSaldo = nSaldo - " & pnMonto & "" _
         & " Where nMovNro = " & nMovNroPend & ""
    
    coConex.CargaRecordSet lsSQL
    Exit Sub
    
Error:
    Err.Raise Err.Number, "Anula Reg Visita Gestores.", Err.Description
End Sub
Public Sub ActualizaMovPendientesOpe(ByVal nMovNroOpe As Long, ByVal pnMonto As Currency)
    Dim sSql As String
    sSql = "STP_UPD_ACTUALIZAPENDOPE " & nMovNroOpe & "," & pnMonto
    coConex.CargaRecordSet sSql
    Exit Sub
End Sub
'END RIRO ***********************
'FRHU 20150415 ERS022-2015
Public Function ObtieneDatosPersCredTranferidos(ByVal psCodCta As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    lsSQL = "exec stp_sel_DatosPersCredTranferidos '" & psCodCta & "'"
    Set lrs = coConex.CargaRecordSet(lsSQL)
    Set ObtieneDatosPersCredTranferidos = lrs
    Set lrs = Nothing
    
    Exit Function
dError:
    Err.Raise Err.Number, "Error Obtiene Datos Creditos Transferidos", Err.Description
End Function
Public Function CredTransfObtieneDatosCancelacion(ByVal psCtaCod As String, ByRef psmensaje As String) As ADODB.Recordset
    'Obtiene Datos para Cancelacion de Credito Transfridos
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    Dim cRecDesc As New DCOMColRecDesc
On Error GoTo dError

    lsSQL = " EXEC stp_sel_CredTransfObtieneDatosCancelaCredRecup '" & psCtaCod & "'"

    Set lrs = coConex.CargaRecordSet(lsSQL)
    If Not (lrs.EOF And lrs.BOF) Then
        If lrs!nprdestado = gColocEstTransfCance Then
            psmensaje = "Credito ya se encuentra en estado " & cRecDesc.mfgEstadoColocRecupDesc(lrs!nprdestado)
            Exit Function
        End If
        Set CredTransfObtieneDatosCancelacion = lrs
        Set lrs = Nothing
    Else
        Set CredTransfObtieneDatosCancelacion = Nothing
    End If
Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Datos Creditos Transferidos Cancelación", Err.Description
End Function
Public Function CredTransfObtieneDistribucionCIMGCobranza(psCtaCod As String, pdFecha As Date) As ADODB.Recordset
    Dim lsSQL As String

On Error GoTo ErrorCredTransfObtieneDistribucionCIMGCobranza

    lsSQL = "EXEC stp_sel_CredTransfObtieneDistribucionCIMGCobranza '" & psCtaCod & "','" & Format(pdFecha, "yyyyMMdd") & "'"
    Set CredTransfObtieneDistribucionCIMGCobranza = coConex.CargaRecordSet(lsSQL)
    
Exit Function
ErrorCredTransfObtieneDistribucionCIMGCobranza:
    Err.Raise Err.Number, "Obtiene Distribución de Montos CIMG para Cobranza Creditos Transferidos", Err.Description
End Function
Public Function getCtaAhorroPorCredito(ByVal psCtaCod As String) As Recordset
    Dim sql As String

    On Error GoTo ErrorgetCtaAhorroPorCredito
    sql = "exec stp_sel_obtenerCtaAhorroPorCredito '" & psCtaCod & "'"
    Set getCtaAhorroPorCredito = coConex.CargaRecordSet(sql)
    
    Exit Function
ErrorgetCtaAhorroPorCredito:
    Err.Raise Err.Number, "Obtener Cuenta de Ahorro Por Credito", Err.Description
End Function
'FIN FRHU 20150415
'WIOR 20150430 ***
Public Function RecuperaColocaciones(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaColocaciones
    
    '*** peac 20071201
    'sSql = "Select * from Colocaciones where cCtaCod = '" & psCtaCod & "'"
    
    '*** peac 20071204
    sSql = " Select a.*, isnull(w.nmonto,0) nMontoColUltCal, isnull(w.dvenc,0) dMontoColUltCal"
    sSql = sSql & " from Colocaciones a"
    sSql = sSql & " Left Join"
    sSql = sSql & " (SELECT cc.cctacod,cc.nmonto, ccc.dvenc"
    sSql = sSql & " from coloccalenddet cc"
    sSql = sSql & " inner join colocaccred c on cc.cctacod=c.cctacod and c.nnrocalen=cc.nnrocalen"
    sSql = sSql & " inner join coloccalendario ccc on ccc.cctacod=cc.cctacod and ccc.nnrocalen=cc.nnrocalen and ccc.ncoloccalendapl=0"
    sSql = sSql & " where cc.ncoloccalendapl=0 and nprdconceptocod=1000) w on  a.cctacod=w.cctacod"
    sSql = sSql & " where a.cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocaciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaColocaciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function ObtenerInteresDiferidoCredito(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_ObtenerInteresDiferido '" & psCtaCod & "'"
    oCon.AbreConexion
    Set ObtenerInteresDiferidoCredito = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerInteresDiferidoPendCredito(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_ObtenerIntDiferidoPend '" & psCtaCod & "'"
    oCon.AbreConexion
    Set ObtenerInteresDiferidoPendCredito = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'WIOR FIN ********
