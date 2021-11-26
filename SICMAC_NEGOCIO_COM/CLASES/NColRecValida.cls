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

Public Function nValidaCambioMetodoLiquid(ByVal psCtaCod As String) As Recordset
Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = "SELECT P.cCtaCod, P.nSaldo, P.nPrdEstado, CR.nSaldoIntComp, nSaldoIntMor, nSaldoGasto, CR.cMetLiquid " _
        & " FROM Producto P Inner Join ColocRecup CR ON P.cCtaCod = CR.cCtaCod " _
        & " WHERE P.cCtaCod = '" & psCtaCod & "'"

    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Or (lrValida.BOF And lrValida.EOF) Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
  
    'If lrValida!nPrdEstado <> gColPEstRegis Then
    '    MsgBox "NO SE PUEDE ANULAR, Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
    '    Exit Function
    'End If
    
    Set nValidaCambioMetodoLiquid = lrValida
    Set lrValida = Nothing
End Function

Public Function nValidaAsignacionGasto(ByVal psCtaCod As String) As Recordset

'************************************

Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.nMontoCol, ,  " _
          & " P.nSaldo, CP.nTasacion, P.nTasaInteres " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " Inner Join ColocRecup CRec on C.cctacod = CRec.cctacod " _
          & " WHERE P.cCtaCod ='" & psCtaCod & "' "
    
    
    Set loRegValida = New dColPFunciones
        Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    
    If (lrValida!nPrdEstado <> gColocEstRecVigJud) And (lrValida!nPrdEstado <> gColocEstRecVigCast) Then
        MsgBox " Contrato se encuentra " '& ContratoEstado(RegCredPrend!cEstado), vbInformation, " Aviso "
        Exit Function
    End If
    
    Set nValidaAsignacionGasto = lrValida
    Set lrValida = Nothing
End Function

Public Function nBuscaCreditoParaCastigar(Optional ByVal psServerConsol As String = "", _
        Optional ByVal psCtaCod As String = "@", Optional ByVal pbOrdenAsc As Boolean = False, Optional psCodAge As String = "01") As Recordset

Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT * FROM (SELECT P.cCtaCod, P.nSaldo, P.nPrdEstado, C.cUltimaActualizacion, CRec.nSaldoIntComp, CRec.nSaldoIntMor," _
          & " CRec.nSaldoGasto, CRec.cMetLiquid, CRec.nDemanda, CRec.nIntCompGen, C.nMontoCol,  " _
          & " cNomClie = ISNULL ( ( SELECT cPersNombre FROM Persona Per INNER JOIN ProductoPersona PP " _
          & "    ON Per.cPersCod = PP.cPersCod WHERE PP.cCtaCod = P.cCtaCod AND nPrdPersRelac = " & gColRelPersTitular & " ), '' ),  " _
          & " nTasaIntComp = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntCompNormal & " ) , " _
          & " nProvision = (SELECT ISNULL(nProvision, 0) From " & psServerConsol & "ColocCalifProv Cal Where Cal.cCtaCod = C.cCtaCod ) " _
          & " FROM Producto P INNER JOIN Colocaciones C ON C.cCtaCod = P.cCtaCod " _
          & " INNER JOIN ColocRecup CRec ON P.cCtaCod = CRec.cCtaCod and SUBSTRING(P.cCtaCod,4,2)='" & psCodAge & "' " _
          & " WHERE P.nPrdEstado in (" & gColocEstRecVigJud & ",2205) "
    If psCtaCod <> "@" Then
        lsSQL = lsSQL & " AND P.cCtaCod ='" & psCtaCod & "'  "
    End If
    lsSQL = lsSQL & " ) AS CRED "
    lsSQL = lsSQL & " ORDER BY  cNomClie ASC  "
    
    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Castigar Creditos ", vbInformation, "Aviso"
        Exit Function
    End If
    
    Set nBuscaCreditoParaCastigar = lrValida
    Set lrValida = Nothing
End Function

Public Function nValidaExpediente(ByVal psCtaCod As String) As Recordset

'************************************

Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

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
          & "              AND nPrdPersRelac = " & gColRelPersSecretario & " ), '' )  " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " Inner Join ColocRecup CRec on C.cCtaCod = CRec.cCtaCod " _
          & " Left Join ColocRecupExpediente CRecEx ON CRec.cCtaCod = CRecEx.cCtaCod " _
          & " Left Join ColocRecupExpedienteInf InfPetit ON CRec.cCtaCod = InfPetit.cCtaCod AND InfPetit.nTipoInf = " & gColRecExpedTipoInfPetit _
          & " Left Join ColocRecupExpedienteInf InfHechos ON CRec.cCtaCod = InfHechos.cCtaCod AND InfHechos.nTipoInf = " & gColRecExpedTipoInfHechos _
          & " Left Join ColocRecupExpedienteInf InfFundJur ON CRec.cCtaCod = InfFundJur.cCtaCod AND InfFundJur.nTipoInf =  " & gColRecExpedTipoInfFundJur _
          & " Left Join ColocRecupExpedienteInf InfMedProb ON CRec.cCtaCod = InfMedProb.cCtaCod AND InfMedProb.nTipoInf = " & gColRecExpedTipoInfMedProb _
          & " Left Join ColocRecupExpedienteInf InfDatComp ON CRec.cCtaCod = InfDatComp.cCtaCod AND InfDatComp.nTipoInf = " & gColRecExpedTipoInfDatComp _
          & " WHERE P.cCtaCod ='" & psCtaCod & "' "
    
    Set loRegValida = New dColPFunciones
        Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
        If lrValida Is Nothing Then
            MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
            Exit Function
        End If
        If lrValida.BOF And lrValida.EOF Then
            MsgBox " Credito No se encuentra en Recuperaciones ", vbInformation, "Aviso"
            Exit Function
        Else
            
            If (lrValida!nPrdEstado <> gColocEstRecVigJud) And (lrValida!nPrdEstado <> gColocEstRecVigCast And _
                lrValida!nPrdEstado <> 2205 And lrValida!nPrdEstado <> 2206) Then
                MsgBox " Credito se encuentra " & mfgEstadoColocRecupDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
                Exit Function
            End If
        End If
    Set nValidaExpediente = lrValida
    Set lrValida = Nothing
End Function

Public Function nBuscaOperacionesParaExtorno(ByVal psFecTrans As String, ByVal psOpeExtor As String, _
        Optional ByVal psCtaCod As String = "@", Optional ByVal pbOrdenAsc As Boolean = False) As Recordset

Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT M.cMovNro, M.nMovNro, M.cOpeCod, M.nMovFlag, " _
          & " MC.cCtaCod, MC.nMonto, Ope.cOpeDesc  " _
          & " FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
          & " LEFT JOIN OpeTpo Ope ON MC.cOpeCod = Ope.cOpeCod " _
          & " WHERE M.cMovNro LIKE '" & psFecTrans & "%' " _
          & " AND M.nMovFlag = " & gMovFlagVigente & " " _
          & " AND M.cOpeCod in " & psOpeExtor & "  "
    If psCtaCod <> "@" Then
        lsSQL = lsSQL & " AND MC.cCtaCod ='" & psCtaCod & "' "
    End If
    lsSQL = lsSQL & " ORDER BY M.nMovNro DESC  "
    
    
    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Extorno ", vbInformation, "Aviso"
        Exit Function
    End If
    
    Set nBuscaOperacionesParaExtorno = lrValida
    Set lrValida = Nothing
End Function

Public Function nValidaActuacionesProc(ByVal psCtaCod As String) As Recordset

'************************************

Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT A.cCtaCod, A.cMovNro, A.cComenta  " _
          & " FROM ColocRecupActProcesales A " _
          & " WHERE A.cCtaCod ='" & psCtaCod & "' "
    
    Set loRegValida = New dColPFunciones
        Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
        If lrValida Is Nothing Then
            'MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
            Exit Function
        End If
        If lrValida.BOF And lrValida.EOF Then
            'MsgBox " Credito No se encuentra en Recuperaciones ", vbInformation, "Aviso"
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


