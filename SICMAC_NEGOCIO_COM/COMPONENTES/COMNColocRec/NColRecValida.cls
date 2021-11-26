VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NColRecValida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************
'* Clase que contiene las validaciones de las operaciones de Creditos en Recuperaciones
'* LAYG - 01/08/2001
'****************************
Option Explicit

Public Function nValidaCambioMetodoLiquid(ByVal psCtaCod As String, ByRef psmensaje As String) As Recordset
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As New ADODB.Recordset

    'Comentado por DAOR 20070421
    'lsSQL = "SELECT P.cCtaCod, P.nSaldo, P.nPrdEstado, CR.nSaldoIntComp, nSaldoIntMor, nSaldoGasto, CR.cMetLiquid " _
    '    & " FROM Producto P Inner Join ColocRecup CR ON P.cCtaCod = CR.cCtaCod " _
    '    & " WHERE P.cCtaCod = '" & psCtaCod & "'"

    '**DAOR 20070421**********************************************************************
'    lsSQL = "SELECT P.cCtaCod, P.nSaldo, P.nPrdEstado,C.cUltimaActualizacion, CR.nSaldoIntComp, CR.nSaldoIntMor, " _
'        & " CR.nSaldoGasto,CR.nIntCompGen, CR.cMetLiquid, IsNull(Com.nValor,0) nValorCom, " _
'        & " PT1.nTasaInteres as nTasaInt, PT2.nTasaInteres as nTasaIntMor " _
'        & "FROM Producto P inner join Colocaciones C on P.cCtaCod=C.cCtaCod " _
'        & " inner Join ColocRecup CR ON P.cCtaCod = CR.cCtaCod " _
'        & " inner join ProductoTasaInteres PT1 on P.cCtaCod=PT1.cCtaCod and PT1.nPrdTasaInteres=" & gColocLineaCredTasasIntCompNormal _
'        & " inner join ProductoTasaInteres PT2 on P.cCtaCod=PT2.cCtaCod and PT2.nPrdTasaInteres=" & gColocLineaCredTasasIntMoratNormal _
'        & " left Join ColocRecupComision Com ON Com.nComisionCod = CR.nComisionCod  " _
'        & "WHERE P.cCtaCod = '" & psCtaCod & "'" COMENTADO POR MACM RIRO 20210810
    '*************************************************************************************
    lsSQL = "stp_sel_nValidaCambioMetodoLiquid '" & psCtaCod & "'" 'MACM RIRO 20210811 ADD


    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Or (lrValida.BOF And lrValida.EOF) Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
  
    'If lrValida!nPrdEstado <> gColPEstRegis Then
    '    MsgBox "NO SE PUEDE ANULAR, Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
    '    Exit Function
    'End If
    
    Set nValidaCambioMetodoLiquid = lrValida
    Set lrValida = Nothing
End Function

Public Function nValidaAsignacionGasto(ByVal psCtaCod As String, ByRef psmensaje As String) As Recordset

'************************************

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As New ADODB.Recordset

    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.nMontoCol, ,  " _
          & " P.nSaldo, CP.nTasacion, P.nTasaInteres " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " Inner Join ColocRecup CRec on C.cctacod = CRec.cctacod " _
          & " WHERE P.cCtaCod ='" & psCtaCod & "' "
    
    
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
        Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    
    If (lrValida!nPrdEstado <> gColocEstRecVigJud) And (lrValida!nPrdEstado <> gColocEstRecVigCast) Then
        psmensaje = " Contrato se encuentra " '& ContratoEstado(RegCredPrend!cEstado)
        Exit Function
    End If
    
    Set nValidaAsignacionGasto = lrValida
    Set lrValida = Nothing
End Function
'MADM 20110701 - CODAGE
Public Function nBuscaCreditoParaCastigar(Optional ByVal psServerConsol As String = "", _
        Optional ByVal psCtaCod As String = "@", Optional ByVal pbOrdenAsc As Boolean = False, Optional psCodAge As String = "01", Optional ByRef psmensaje As String) As Recordset

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As New ADODB.Recordset

'DAOR 20070731, Indicaciones por Yomar
'    lsSQL = " SELECT * FROM (SELECT P.cCtaCod, P.nSaldo, P.nPrdEstado, C.cUltimaActualizacion, CRec.nSaldoIntComp, CRec.nSaldoIntMor," _
'          & " CRec.nSaldoGasto, CRec.cMetLiquid, CRec.nDemanda, CRec.nIntCompGen, C.nMontoCol,  " _
'          & " cNomClie = ISNULL ( ( SELECT cPersNombre FROM Persona Per INNER JOIN ProductoPersona PP " _
'          & "    ON Per.cPersCod = PP.cPersCod WHERE PP.cCtaCod = P.cCtaCod AND nPrdPersRelac = " & gColRelPersTitular & " ), '' ),  " _
'          & " nTasaIntComp = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntCompNormal & " ) , " _
'          & " nProvision = (SELECT ISNULL(nProvision, 0) From " & psServerConsol & "ColocCalifProv Cal Where Cal.cCtaCod = C.cCtaCod ) " _
'          & " FROM Producto P INNER JOIN Colocaciones C ON C.cCtaCod = P.cCtaCod " _
'          & " INNER JOIN ColocRecup CRec ON P.cCtaCod = CRec.cCtaCod and SUBSTRING(P.cCtaCod,4,2)='" & psCodAge & "' " _
'          & " WHERE P.nPrdEstado in (" & gColocEstRecVigJud & ",2205) "

    
    '**Comentado por DAOR 20070728
'    lsSQL = " SELECT * FROM (SELECT P.cCtaCod, P.nSaldo, P.nPrdEstado, C.cUltimaActualizacion, CRec.nSaldoIntComp, CRec.nSaldoIntMor," _
'          & " CRec.nSaldoGasto, CRec.cMetLiquid, CRec.nDemanda, CRec.nIntCompGen, C.nMontoCol,  " _
'          & " cNomClie = ISNULL ( ( SELECT cPersNombre FROM Persona Per INNER JOIN ProductoPersona PP " _
'          & "    ON Per.cPersCod = PP.cPersCod WHERE PP.cCtaCod = P.cCtaCod AND nPrdPersRelac = " & gColRelPersTitular & " ), '' ),  " _
'          & " nTasaIntComp = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntCompNormal & " ) , " _
'          & " nProvision = (SELECT ISNULL(nProvision, 0) From " & psServerConsol & "ColocCalifProv Cal Where Cal.cCtaCod = C.cCtaCod ) " _
'          & " FROM Producto P INNER JOIN Colocaciones C ON C.cCtaCod = P.cCtaCod " _
'          & " INNER JOIN ColocRecup CRec ON P.cCtaCod = CRec.cCtaCod " _
'          & " WHERE P.nPrdEstado in (" & gColocEstRecVigJud & ",2205) "
    
    '**DAOR 20070828, se corrigió TasaIntComp y se agregó TasaIntMora para que tome de ProductoTasaInteres
    lsSQL = " select *," _
          & "       nSaldoGasto=isnull((select sum(nMonto-nMontoPagado) from ColocRecupGastos where cCtaCod=CRED.cCtaCod and nColocRecGastoEstado=0),0) " _
          & " from " _
          & "   (select P.cCtaCod, P.nSaldo, P.nPrdEstado, C.cUltimaActualizacion, CRec.nSaldoIntComp, CRec.nSaldoIntMor," _
          & "       CRec.cMetLiquid, CRec.nDemanda, CRec.nIntCompGen, C.nMontoCol,  " _
          & "       cNomClie = isnull((select cPersNombre from Persona Per inner join ProductoPersona PP " _
          & "               on Per.cPersCod = PP.cPersCod where PP.cCtaCod = P.cCtaCod and PP.nPrdPersRelac = " & gColRelPersTitular & " ), '' ),  " _
          & "       nTasaIntComp =(select isnull(nTasaInteres, 0) From ProductoTasaInteres PTI  where PTI.cCtaCod = P.cCtaCod and PTI.nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & " ) , " _
          & "       nTasaIntMora =(select isnull(nTasaInteres, 0) From ProductoTasaInteres PTI  where PTI.cCtaCod = P.cCtaCod and PTI.nPrdTasaInteres = " & gColocLineaCredTasasIntMoratNormal & " ), " _
          & "       isnull(Cal.nProvision ,0) as nProvision,C.cAgeCodAct cAgeCod  " _
          & "   from Producto P inner join Colocaciones C ON C.cCtaCod = P.cCtaCod " _
          & "       inner join ColocRecup CRec ON P.cCtaCod = CRec.cCtaCod " _
          & "       inner join " & psServerConsol & "ColocCalifProvTotal Cal on C.cCtaCod=Cal.cCtaCod " _
          & "               and datediff(day,Cal.dFecha,(select max(dFecha) from " & psServerConsol & "ColocCalifProvTotal))=0 " _
          & "   where P.nPrdEstado in (" & gColocEstRecVigJud & ",2205) " _
          & "   and Cal.nProvision>=P.nSaldo "
          
    If psCtaCod <> "@" Then
        lsSQL = lsSQL & " AND P.cCtaCod ='" & psCtaCod & "'  "
    End If
    lsSQL = lsSQL & " ) AS CRED "
    lsSQL = lsSQL & " ORDER BY  cNomClie ASC  "
    
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Castigar Creditos "
        Exit Function
    End If
    
    Set nBuscaCreditoParaCastigar = lrValida
    Set lrValida = Nothing
End Function

Public Function nValidaExpediente(ByVal psCtaCod As String, ByRef psmensaje As String) As Recordset

'************************************

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim orecdes As New COMDColocRec.DCOMColRecDesc
Dim lrValida As New ADODB.Recordset

    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.nMontoCol, P.nSaldo, CRec.dIngRecup, CRec.nDemanda, CRec.nTipCJ, " _
          & " ISNULL(CRec.nComisionCod,'') nComisionCod, " _
          & " nTipComis = ISNULL ( ( SELECT nTipComis FROM ColocRecupComision WHERE nComisionCod = CRec.nComisionCod ), '') , " _
          & " nComisionValor = ISNULL ( ( SELECT nValor FROM ColocRecupComision WHERE nComisionCod = CRec.nComisionCod ), 0 ) , " _
          & " ISNULL(CRecEx.cNumExp,'') cNumExp, ISNULL(CRecEx.nMonPetit, 0 ) nMonPetit, ISNULL(CRecEx.nViaProce, 0 ) nViaProce, " _
          & " ISNULL(InfPetit.mInforme,'') mPetit, ISNULL(InfHechos.mInforme,'') mHechos, ISNULL(InfFundJur.mInforme,'') mFundJur, " _
          & " ISNULL(InfMedProb.mInforme,'') mMedProb, ISNULL(InfDatComp.mInforme,'') mDatComp , ISNULL(CRecEx.nViaProce, 9 ) nControl,  " _
          & " cCodClie = ISNULL ( ( SELECT cPersCod FROM ProductoPersona WHERE cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersTitular & " ), '' ) , " _
          & " cNomClie = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ProductoPersona PP " _
          & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersTitular & " ), '' ) , " _
          & " cCodAbog = ISNULL ( ( SELECT cPersCod FROM ProductoPersona WHERE cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersEstudioJuridico & " ), '' ) , " _
          & " cNomAbog = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ProductoPersona PP " _
          & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersEstudioJuridico & " ), '' ) , " _
          & " cCodJuzg = ISNULL ( ( SELECT cPersCod FROM ProductoPersona WHERE cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersJuzgado & " ), '' ) , " _
          & " cNomJuzg = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ProductoPersona PP " _
          & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersJuzgado & " ), '' ) , "
    lsSQL = lsSQL & " " _
          & " cCodJuez = ISNULL ( ( SELECT cPersCod FROM ProductoPersona WHERE cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersJuez & " ), '' ) , " _
          & " cNomJuez = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ProductoPersona PP " _
          & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersJuez & " ), '' ) , " _
          & " cCodSecre = ISNULL ( ( SELECT cPersCod FROM ProductoPersona WHERE cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersSecretario & " ), '' ) , " _
          & " cNomSecre = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ProductoPersona PP " _
          & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersSecretario & " ), '' ),  " _
          & " ISNULL(CRecEx.nTipoProceso,'') TipoProceso, " _
          & " ISNULL(CRecEx.nTipMedCau,0) nTipMedCau " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " Inner Join ColocRecup CRec on C.cCtaCod = CRec.cCtaCod " _
          & " Left Join ColocRecupExpediente CRecEx ON CRec.cCtaCod = CRecEx.cCtaCod " _
          & " Left Join ColocRecupExpedienteInf InfPetit ON CRec.cCtaCod = InfPetit.cCtaCod AND InfPetit.nTipoInf = " & gColRecExpedTipoInfPetit _
          & " Left Join ColocRecupExpedienteInf InfHechos ON CRec.cCtaCod = InfHechos.cCtaCod AND InfHechos.nTipoInf = " & gColRecExpedTipoInfHechos _
          & " Left Join ColocRecupExpedienteInf InfFundJur ON CRec.cCtaCod = InfFundJur.cCtaCod AND InfFundJur.nTipoInf =  " & gColRecExpedTipoInfFundJur _
          & " Left Join ColocRecupExpedienteInf InfMedProb ON CRec.cCtaCod = InfMedProb.cCtaCod AND InfMedProb.nTipoInf = " & gColRecExpedTipoInfMedProb _
          & " Left Join ColocRecupExpedienteInf InfDatComp ON CRec.cCtaCod = InfDatComp.cCtaCod AND InfDatComp.nTipoInf = " & gColRecExpedTipoInfDatComp _
          & " WHERE P.cCtaCod ='" & psCtaCod & "' "
    
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
        Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
        If lrValida Is Nothing Then
            psmensaje = "ERROR: al Buscar datos para Validación "
            Exit Function
        End If
        If lrValida.BOF And lrValida.EOF Then
            psmensaje = " Credito No se encuentra en Recuperaciones "
            Exit Function
        Else
            
            If (lrValida!nPrdEstado <> gColocEstRecVigJud) And (lrValida!nPrdEstado <> gColocEstRecVigCast And _
                lrValida!nPrdEstado <> 2205 And lrValida!nPrdEstado <> 2206) Then
                psmensaje = " Credito se encuentra " & orecdes.mfgEstadoColocRecupDesc(lrValida!nPrdEstado)
                Exit Function
            End If
        End If
    Set nValidaExpediente = lrValida
    Set lrValida = Nothing
End Function


Public Function nValidaExpedienteRelacionados(ByVal psCtaCod As String) As ADODB.Recordset

'************************************

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim orecdes As New COMDColocRec.DCOMColRecDesc
Dim lrValida As New ADODB.Recordset

    lsSQL = " SELECT cPersNombre,PE.cPersCod,nSaldo,dIngRecup,nMontoCol,nTipCj,nDemanda" & _
            " FROM  Producto P" & _
            " INNER JOIN ProductoPersona  PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=20 " & _
            " INNER JOIN Persona PE ON PE.cPersCod=PP.cPersCod" & _
            " INNER JOIN Colocaciones C ON C.cCtaCod=P.cCtaCod" & _
            " INNER JOIN ColocRecup CR ON P.cCtaCod=CR.cCtaCod" & _
            " WHERE P.cCtaCod='" & psCtaCod & "'"
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
        Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
        Set nValidaExpedienteRelacionados = lrValida
    Set lrValida = Nothing
End Function


Public Function nBuscaOperacionesParaExtorno(ByVal psFecTrans As String, ByVal psOpeExtor As String, _
        Optional ByVal psCtaCod As String = "@", Optional ByVal pbOrdenAsc As Boolean = False, Optional ByRef psmensaje As String) As Recordset

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As New ADODB.Recordset
    
    'FRHU 20150520 ERS022-2015
    'lsSQL = " SELECT M.cMovNro, M.nMovNro, M.cOpeCod, M.nMovFlag,cPersNombre," _
    '      & " MC.cCtaCod, MC.nMonto, Ope.cOpeDesc  " _
    '      & " FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
    '      & " JOIN ProductoPersona  PP ON PP.cCtaCod=MC.cCtaCod and nPrdPersRelac=20" _
    '      & " JOIN Persona P ON P.cPersCod=PP.cPersCod " _
    '      & " LEFT JOIN OpeTpo Ope ON MC.cOpeCod = Ope.cOpeCod " _
    '      & " WHERE M.cMovNro LIKE '" & psFecTrans & "%' " _
    '      & " AND M.nMovFlag = " & gMovFlagVigente & " " _
    '      & " AND MC.cOpeCod in " & psOpeExtor & "  " 'JUEZ 20130921 Para buscar el codigo en MovCol
    '      '& " AND M.cOpeCod in " & psOpeExtor & "  "
    'If psCtaCod <> "@" Then
        'lsSQL = lsSQL & " AND MC.cCtaCod ='" & psCtaCod & "' "
    'End If
    'lsSQL = lsSQL & " ORDER BY M.nMovNro DESC  "
    lsSQL = "EXEC stp_sel_buscaOperacionesParaExtorno '" & psFecTrans & "','" & psOpeExtor & "','" & psCtaCod & "'"
    'FIN FRHU 20150520
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Extorno "
        Exit Function
    End If
    
    Set nBuscaOperacionesParaExtorno = lrValida
    Set lrValida = Nothing
End Function

'** Modificado por DAOR 20070124
'** En la consulta se incluye el tipo de actuación procesal
Public Function nValidaActuacionesProc(ByVal psCtaCod As String, ByRef psmensaje As String) As Recordset
'************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As New ADODB.Recordset

    lsSQL = " SELECT A.cCtaCod, A.cMovNro, A.cComenta, A.dFechaAviso,A.dFechaVencimiento, C.cConsDescripcion TipoAct  " _
          & " FROM ColocRecupActProcesales A inner join Constante C on A.nTipoAct=C.nConsValor and C.nConsCod=3310 " _
          & " WHERE A.cCtaCod ='" & psCtaCod & "' "
    
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
        Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
        If lrValida Is Nothing Then
            psmensaje = "ERROR: al Buscar datos para Validación "
            Exit Function
        End If
        If lrValida.BOF And lrValida.EOF Then
            psmensaje = "Credito No posee Actuaciones Procesales "
            Exit Function
        Else
            'If (lrValida!nPrdEstado <> gColocEstRecVigJud) And (lrValida!nPrdEstado <> gColocEstRecVigCast) Then
            '    MsgBox " Credito se encuentra " & mfgEstadoColocRecupDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
            '    Exit Function
            'End If
        End If
    Set nValidaActuacionesProc = lrValida
    Set lrValida = Nothing
End Function

Public Function ObtenerCorrelativoExped(ByVal psCtaCod As String) As String
    Dim lsSQL As String
    Dim lrCorre As New ADODB.Recordset
    Dim loRegCorre As COMDColocPig.DCOMColPFunciones
    lsSQL = " SELECT Max(nNroCorreRelacion) nNro  FROM ColocRecupExpedienteRelacionados " & _
            " WHERE cCtaCod='" & psCtaCod & "' "
    Set loRegCorre = New COMDColocPig.DCOMColPFunciones
        Set lrCorre = loRegCorre.dObtieneRecordSet(lsSQL)
            
        If lrCorre.BOF And lrCorre.EOF Then
            ObtenerCorrelativoExped = 1
        Else
            ObtenerCorrelativoExped = IIf(IsNull(lrCorre!nNro), 1, (lrCorre!nNro + 1))
        End If
    
    Set loRegCorre = Nothing
End Function

Public Function nValidaExpedientePersona(ByVal psCtaCod As String, ByVal pnCorrelativo As Integer, ByRef psmensaje As String) As Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim orecdes As New COMDColocRec.DCOMColRecDesc
Dim lrValida As New ADODB.Recordset

    lsSQL = " SELECT " _
          & " ISNULL(CRecEx.cNumExp,'') cNumExp, ISNULL(CRecEx.nMonPetit, 0 ) nMonPetit, ISNULL(CRecEx.nViaProce, 0 ) nViaProce, " _
          & " ISNULL(InfPetit.mInforme,'') mPetit, ISNULL(InfHechos.mInforme,'') mHechos, ISNULL(InfFundJur.mInforme,'') mFundJur, " _
          & " ISNULL(InfMedProb.mInforme,'') mMedProb, ISNULL(InfDatComp.mInforme,'') mDatComp , ISNULL(CRecEx.nViaProce, 9 ) nControl,  " _
          & " cCodClie = ISNULL ( ( SELECT cPersCod FROM ColocRecupExpedientePersona WHERE cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersTitular & " ), '' ) , " _
          & " cNomClie = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ColocRecupExpedientePersona PP " _
          & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersTitular & " ), '' ) , " _
          & " cCodAbog = ISNULL ( ( SELECT cPersCod FROM ColocRecupExpedientePersona WHERE cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersEstudioJuridico & " ), '' ) , " _
          & " cNomAbog = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ColocRecupExpedientePersona PP " _
          & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersEstudioJuridico & " ), '' ) , " _
          & " cCodJuzg = ISNULL ( ( SELECT cPersCod FROM ColocRecupExpedientePersona WHERE cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersJuzgado & " ), '' ) , " _
          & " cNomJuzg = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ColocRecupExpedientePersona PP " _
          & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersJuzgado & " ), '' ) , "
    lsSQL = lsSQL & " " _
          & " cCodJuez = ISNULL ( ( SELECT cPersCod FROM ColocRecupExpedientePersona WHERE cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersJuez & " ), '' ) , " _
          & " cNomJuez = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ColocRecupExpedientePersona PP " _
          & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersJuez & " ), '' ) , " _
          & " cCodSecre = ISNULL ( ( SELECT cPersCod FROM ColocRecupExpedientePersona WHERE cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersSecretario & " ), '' ) , " _
          & " cNomSecre = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ColocRecupExpedientePersona PP " _
          & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersSecretario & " ), '' ),"
    lsSQL = lsSQL & " " _
          & " cCodEncargado = ISNULL ( ( SELECT cPersCod FROM ColocRecupExpedientePersona WHERE cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersEncargado & " ), '' ) , " _
          & " cNomEncargado = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ColocRecupExpedientePersona PP " _
          & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psCtaCod & "' " _
          & "              AND nPrdPersRelac = " & gColRelPersEncargado & " ), '' )," _
          & " ISNULL(nTipoProceso,'') TipoProceso,ISNULL(nTipoProcesal,'') TipoProcesal " _
          & " FROM Persona P " _
          & " INNER JOIN ColocRecupExpedientePersona RP ON P.cPersCod= RP.cPersCod " _
          & " INNER JOIN ColocRecupExpedienteRelacionados CRecEx ON RP.cCtaCod=CRecEx.cCtaCod " _
          & " Left Join ColocRecupExpedienteRelacionadosInf InfPetit ON CRecEx.cCtaCod = InfPetit.cCtaCod AND InfPetit.nTipoInf = " & gColRecExpedTipoInfPetit _
          & " Left Join ColocRecupExpedienteRelacionadosInf InfHechos ON CRecEx.cCtaCod = InfHechos.cCtaCod AND InfHechos.nTipoInf = " & gColRecExpedTipoInfHechos _
          & " Left Join ColocRecupExpedienteRelacionadosInf InfFundJur ON CRecEx.cCtaCod = InfFundJur.cCtaCod AND InfFundJur.nTipoInf =  " & gColRecExpedTipoInfFundJur _
          & " Left Join ColocRecupExpedienteRelacionadosInf InfMedProb ON CRecEx.cCtaCod = InfMedProb.cCtaCod AND InfMedProb.nTipoInf = " & gColRecExpedTipoInfMedProb _
          & " Left Join ColocRecupExpedienteRelacionadosInf InfDatComp ON CRecEx.cCtaCod = InfDatComp.cCtaCod AND InfDatComp.nTipoInf = " & gColRecExpedTipoInfDatComp _
          & " WHERE RP.cCtaCod ='" & psCtaCod & "' and CRecEx.nNroCorreRelacion =" & pnCorrelativo & ""
    
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
        Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
        If lrValida Is Nothing Then
            psmensaje = "ERROR: al Buscar datos para Validación "
            Exit Function
        End If
    Set nValidaExpedientePersona = lrValida
    Set lrValida = Nothing
End Function

'*** PEAC 20120705
Public Function ObtieneDatosVisitaCliente(ByVal psCtaCod As String, ByRef psmensaje As String) As Recordset

'************************************

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim orecdes As New COMDColocRec.DCOMColRecDesc
Dim lrValida As New ADODB.Recordset

    lsSQL = "exec stp_sel_ObtieneDatosClienteVisita '" & psCtaCod & "'"
    
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
        If lrValida Is Nothing Then
            psmensaje = "ERROR: al Buscar datos del crédito"
            Exit Function
        End If
        If lrValida.BOF And lrValida.EOF Then
            psmensaje = " Credito No tiene movimientos"
            Exit Function
        End If
    Set ObtieneDatosVisitaCliente = lrValida
    Set lrValida = Nothing
End Function
'FRHU 20150428 ERS022-2015
Public Function CredTransfValidaCambioMetodoLiquid(ByVal psCtaCod As String, ByRef psmensaje As String) As Recordset
    Dim lsSQL As String
    Dim loRegValida As COMDColocPig.DCOMColPFunciones
    Dim lrValida As New ADODB.Recordset

    lsSQL = "EXEC stp_sel_CredTransfValidaCambioMetodoLiquid '" & psCtaCod & "'"

    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Or (lrValida.BOF And lrValida.EOF) Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    
    Set CredTransfValidaCambioMetodoLiquid = lrValida
    Set lrValida = Nothing
End Function
'FIN FRHU 20150428
