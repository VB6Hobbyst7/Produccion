VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nPigValida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim sql As String
Dim oConn As DConecta
Private Sub Class_Initialize()
Dim oIni As ClasIni
Dim sConn As String

    Set oIni = New ClasIni
    sConn = oIni.CadenaConexion
    Set oIni = Nothing
    
Set oConn = New DConecta
If oConn.AbreConexion(sConn) = False Then
    Call RaiseError(MyUnhandledError, "DPigActualizaBD:Initialize. Error en Conexion a Base de datos")
End If

End Sub
Private Sub Class_Terminate()
    oConn.CierraConexion
    Set oConn = Nothing
End Sub

Public Function nValidaDesembolsoCredPignoraticio(ByVal psCtaCod As String, ByVal pdFecSis As Date, _
     ByVal pnMaxDiasDesembolso As Integer) As Recordset
     
Dim lrValida As ADODB.Recordset

    sql = " SELECT P.cCtaCod, P.nPrdEstado, dVigencia, " _
         & " cBloqueo = (SELECT CASE WHEN COUNT(nBlqMotivo) > 0 THEN 'S' ELSE 'N' END FROM ProductoBloqueos PB LEFT JOIN Constante Cons " _
         & "             ON PB.nBlqMotivo = Cons.nConsCod WHERE PB.cCtaCod = '" & psCtaCod & "' AND cMovNroDbl Is Null )  " _
         & " FROM Colocaciones c Inner join Producto P on P.cCtaCod = c.cCtaCod " _
         & " Inner join ColocCalendario CCal on CCal.cCtaCod = C.cCtaCod " _
         & " Inner join ColoccalendDetPig CCalD on CCal.cCtaCod = CCalD.cCtaCod and CCal.nNroCalen = CCalD.nNroCalen " _
         & " And CCal.nColocCAlendApl = CCalD.nColocCalendApl " _
         & " Where P.cCtaCod  = '" & psCtaCod & "' AND CCalD.nColocCalendApl =  " & gColocCalendAplDesembolso & " " _
         & " GROUP BY P.cCtaCod, P.nPrdEstado, dVigencia "
   
    Set lrValida = oConn.CargaRecordSet(sql)
    If lrValida.BOF And lrValida.EOF Then
        MsgBox " No se ha encontrado Contrato ", vbInformation, " Aviso "
        Set lrValida = Nothing
        Exit Function
    End If
    
    If lrValida!nPrdEstado <> gPigEstRegis Then 'estado 2100 - registrado
        MsgBox " Contrato se encuentra " & mfgEstCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
        Exit Function
    ElseIf DateDiff("d", lrValida!dVigencia, pdFecSis) > pnMaxDiasDesembolso Then 'Controla los dias de vigencias
        MsgBox " Contrato está fuera de la Fecha para su Desembolso ", vbInformation, " Aviso "
        Set lrValida = Nothing
        Exit Function
    End If
    
    If lrValida!cBloqueo = "S" Then 'en caso de que este bloqueado
        MsgBox " Contrato se encuentra Bloqueado ", vbInformation, " Aviso "
        Set lrValida = Nothing
        Exit Function
    End If
    
    Set nValidaDesembolsoCredPignoraticio = lrValida
    Set lrValida = Nothing

End Function
Public Function nValidaAnulacionCredPignoraticio(ByVal psCtaCod As String, ByVal pdFecSis As Date) As Recordset
Dim Rs As ADODB.Recordset

    sql = " SELECT P.cCtaCod, P.nPrdestado,  " _
          & " dRegistrado = (SELECT CONVERT(datetime, Substring(cMovNro,1,8),  103 )  " _
          & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
          & "              WHERE MC.cCtaCod = C.cCtaCod AND MC.cOpeCod in ( " & gPigOpeRegContrato & ")  )" _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " Inner Join ColocPigno CP on C.cctacod = CP.cctacod " _
          & " WHERE P.cCtaCod ='" & psCtaCod & "' "
    
    Set Rs = oConn.CargaRecordSet(sql)
    If Rs Is Nothing Or (Rs.BOF And Rs.EOF) Then
        MsgBox "No se encontraron datos del Contrato", vbInformation, "Aviso"
        Exit Function
    End If
  
    If Rs!nPrdEstado <> gPigEstRegis Then
        MsgBox "NO SE PUEDE ANULAR, Contrato se encuentra " & mfgEstCredPigDesc(Rs!nPrdEstado), vbInformation, " Aviso "
        Set Rs = Nothing
        Exit Function
    End If
    
    If DateDiff("d", Rs!dRegistrado, pdFecSis) <> 0 Then
        MsgBox " Contrato se encuentra fuera de su Fecha Limite ", vbInformation, " Aviso "
        Set Rs = Nothing
        Exit Function
    End If
    
    Set nValidaAnulacionCredPignoraticio = Rs
    Set Rs = Nothing
    
End Function
Public Function nValidaBloqueoCredPignoraticio(ByVal psCtaCod As String) As Recordset
Dim Rs As ADODB.Recordset

sql = "SELECT P.nPrdEstado, PBloq.nBlqTpo, PBloq.nBlqMotivo, ISNULL(PBloq.cComentario,'') cComentario, PBloq.cMovNro as cMovNroBloqueo , " _
    & " cBloqueo = (SELECT CASE WHEN COUNT(nBlqMotivo) > 0 THEN 'S' ELSE 'N' END FROM ProductoBloqueos PB " _
    & "WHERE PB.cCtaCod = '" & psCtaCod & "' AND cMovNroDbl Is Null ) " _
    & " FROM Producto P INNER JOIN ColocPigno CP ON P.cCtaCod = CP.cCtaCod  " _
    & " LEFT JOIN ProductoBloqueos PBloq ON CP.cCtaCod = PBloq.cCtaCod " _
    & " WHERE P.cCtaCod ='" & psCtaCod & "' ORDER BY cMovNro DESC "
     
    Set Rs = oConn.CargaRecordSet(sql)
    If Rs Is Nothing Or (Rs.BOF And Rs.EOF) Then
        MsgBox "Contrato no existe", vbInformation, "Aviso"
        Exit Function
    End If

    If (Rs!nPrdEstado <> gPigEstRegis) And (Rs!nPrdEstado <> gPigEstDesemb) And _
       (Rs!nPrdEstado <> gPigEstAmortiz) And (Rs!nPrdEstado <> gPigEstReusoLin) And _
       (Rs!nPrdEstado <> gPigEstCancelPendRes) And (Rs!nPrdEstado <> gPigEstRemat) And _
       (Rs!nPrdEstado <> gPigEstRematPRes) And (Rs!nPrdEstado <> gPigEstRematPFact) And (Rs!nPrdEstado <> gPigEstPResRematPFact) Then
       
        MsgBox " Contrato se encuentra " & mfgEstCredPigDesc(Rs!nPrdEstado), vbInformation, " Aviso "
        Set Rs = Nothing
        Exit Function
    End If
    
    Set nValidaBloqueoCredPignoraticio = Rs
    Set Rs = Nothing
    
End Function

Public Function nValidaDuplicadoContratoCredPignoraticio(ByVal psCtaCod As String) As Recordset
Dim Rs As ADODB.Recordset

    sql = " SELECT P.cCtaCod, P.nPrdEstado FROM Producto P Where P.cCtaCod  = '" & psCtaCod & "'"
    
    Set Rs = oConn.CargaRecordSet(sql)
    
    If Rs Is Nothing Then
        MsgBox "Cuenta no existe", vbInformation, "Aviso"
        Exit Function
    End If
   
    If IsNull(Rs!nPrdEstado) Or (Rs.EOF And Rs.BOF) Then
        MsgBox "Cuenta no valida", vbInformation, "Aviso"
        Exit Function
    Else
        If Rs!nPrdEstado <> gPigEstRegis And Rs!nPrdEstado <> gPigEstDesemb _
            And Rs!nPrdEstado <> gPigEstAmortiz And Rs!nPrdEstado <> gPigEstReusoLin _
            And Rs!nPrdEstado <> gPigEstCancelPendRes And Rs!nPrdEstado <> gPigEstRemat _
            And Rs!nPrdEstado <> gPigEstRematPRes Then
            MsgBox " Contrato se encuentra " & mfgEstCredPigDesc(Rs!nPrdEstado), vbInformation, " Aviso "
            Exit Function
            Set Rs = Nothing
        End If
    End If
    Set nValidaDuplicadoContratoCredPignoraticio = Rs
    Set Rs = Nothing

End Function

Public Function nValidaComision(ByVal psCtaCod As String, ByVal pnEstado As String) As Recordset
Dim lrValida As ADODB.Recordset

    sql = "SELECT P.cCtaCod, nTasaInteres, nSaldo, nPrdEstado, dPrdEstado, nTransacc, " _
        & "SUM(nTasacion + nTasacionAdicional) Tasacion, nUbicaLote " _
        & "FROM Producto P INNER JOIN COLOCPIGJOYATASACION T ON P.CCTACOD = T.CCTACOD " _
        & "INNER JOIN COLOCPIGNO CP ON CP.cCtaCod = P.cCtaCod " _
        & "WHERE nPrdEstado IN (" & pnEstado & ") AND T.nTipoTasacion = 1 " _
        & "AND P.cCtaCod = '" & psCtaCod & "' " _
        & "GROUP BY P.cCtaCod, nTasaInteres, nSaldo, nPrdEstado, dPrdEstado, nTransacc, nUbicaLote "
        
    Set lrValida = oConn.CargaRecordSet(sql)
    
    If lrValida.BOF And lrValida.EOF Then
        MsgBox " No se ha encontrado Contrato ", vbInformation, " Aviso "
    End If
    
    Set nValidaComision = lrValida
    Set lrValida = Nothing

End Function

Public Function nValidaAmortizacionCredPignoraticio(ByVal psCtaCod As String, ByVal pdFecSis As Date, _
                                                                          ByVal pnMaxDiasDesembolso As Integer) As Recordset
Dim lrValida As ADODB.Recordset

    sql = " SELECT P.cCtaCod, P.nPrdestado, nTransacc,  C.dVenc, C.nPlazo, CP.nNroAmort, CP.nUsoLineaNro,  CP.nDiasAtraso, " _
          & " P.nSaldo, P.nTasaInteres, CP.nNumCalend, C.nMontoCol, P.dPrdEstado, CP.nNumCalendDesem, sum(ntasacion + ntasacionadicional) totTasacion, " _
          & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " )  " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & "                 INNER JOIN ColocPigno CP on C.cctacod = CP.cctacod " _
          & "                 INNER JOIN ColocPigJoyaTasacion CPT on CP.cctacod = CPT.cctacod " _
          & "                                        AND CP.nTipoTasacion = CPT.nTipoTasacion " _
          & " WHERE P.cCtaCod ='" & psCtaCod & "' " _
          & " GROUP BY P.cCtaCod, P.nPrdestado, nTransacc, C.dVenc, C.nPlazo, CP.nNroAmort, CP.nUsoLineaNro,  CP.nDiasAtraso, " _
          & "          P.nSaldo , P.nTasaInteres, CP.nNumCalend, C.cLineaCred, C.nMontoCol, P.dPrdEstado, CP.nNumCalendDesem "
    
    Set lrValida = oConn.CargaRecordSet(sql)
    If lrValida.BOF And lrValida.EOF Then
        MsgBox "No se encontro el Credito ", vbInformation, "Aviso"
        Set lrValida = Nothing
        Exit Function
    End If
    
    If (lrValida!nPrdEstado <> gPigEstDesemb) And (lrValida!nPrdEstado <> gPigEstReusoLin) And _
       (lrValida!nPrdEstado <> gPigEstRemat) And (lrValida!nPrdEstado <> gPigEstRematPRes) And _
       (lrValida!nPrdEstado <> gPigEstAmortiz) And (lrValida!nPrdEstado <> gPigEstRematPRes) Then
        MsgBox " Contrato se encuentra " & mfgEstCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
        Set lrValida = Nothing
        Exit Function
    End If
           
    If IsNull(lrValida!nTasaIntVenc) Then
        MsgBox " No se ha definido la Tasa de Interes Vencida para linea de Credito ", vbInformation, " Aviso "
        Set lrValida = Nothing
        Exit Function
    End If
    
    Set nValidaAmortizacionCredPignoraticio = lrValida
    Set lrValida = Nothing
    
End Function

Public Function nValidaCancelacionCredPignoraticio(ByVal psCtaCod As String, ByVal pdFecSis As Date, _
                                                   ByVal pnMaxDiasDesembolso As Integer) As Recordset
Dim lrValida As ADODB.Recordset
    
    sql = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, CP.nNroAmort, CP.nUsoLineaNro, CP.nDiasAtraso, nTransacc, nNroDuplic, " _
          & " P.nSaldo, P.nTasaInteres, CP.nNumCalend, C.nMontoCol, P.dPrdEstado, C.dVigencia, sum(ntasacion + nTasacionAdicional) totTasacion, " _
          & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " )  " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & "                 INNER JOIN ColocPigno CP on C.cctacod = CP.cctacod " _
          & "                 INNER JOIN ColocPigJoyaTasacion CPT ON CP.cctacod = CPT.cctacod " _
          & "                                        AND CP.nTipoTasacion = CPT.nTipoTasacion " _
          & " WHERE P.cCtaCod ='" & psCtaCod & "' " _
          & " GROUP BY P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, CP.nNroAmort, CP.nUsoLineaNro,  CP.nDiasAtraso, " _
          & "          P.nSaldo , P.nTasaInteres, CP.nNumCalend, C.cLineaCred, C.nMontoCol, P.dPrdEstado, C.dVigencia, P.nTransacc, nNroDuplic"
    
    Set lrValida = oConn.CargaRecordSet(sql)
    
    If lrValida Is Nothing Or (lrValida.EOF And lrValida.BOF) Then
        MsgBox "No se encontro datos del Contrato", vbInformation, "Aviso"
        Set lrValida = Nothing
        Exit Function
    End If
    
    If (lrValida!nPrdEstado <> gPigEstDesemb) And (lrValida!nPrdEstado <> gPigEstReusoLin) And _
       (lrValida!nPrdEstado <> gPigEstRemat) And (lrValida!nPrdEstado <> gPigEstRematPRes) And _
       (lrValida!nPrdEstado <> gPigEstAmortiz) And (lrValida!nPrdEstado <> gPigEstRematPRes) Then
        MsgBox " Contrato se encuentra " & mfgEstCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
        Set lrValida = Nothing
        Exit Function
    End If
    
    Set nValidaCancelacionCredPignoraticio = lrValida
    Set lrValida = Nothing
    
End Function

Public Function nBuscaOperacionesCredPigParaExtorno(ByVal psFecTrans As String, ByVal pnOpeCod As Long, Optional ByVal psCtaCod As String = "@", _
                                 Optional ByVal psCodAge As String = "@", Optional ByVal pbOrdenAsc As Boolean = False) As Recordset

Dim lrValida As ADODB.Recordset
Dim lsCod As String

    If pnOpeCod = "1509" Then
        lsCod = "3,4"
    ElseIf pnOpeCod = "1502" Then
        lsCod = "2"
    Else
        lsCod = "99"
    End If
    
    If pnOpeCod <> "1521" Then
    
        sql = "  SELECT M.cMovNro, M.nMovNro, M.cOpeCod, M.nMovFlag, cPersNombre, " _
                 & "    MC.cCtaCod, Ope.cOpeDesc, sum(MC.nMonto) nMonto, nSaldoCap, nTransacc, " _
                 & " NroCalen = (CASE WHEN M.cOpeCod = '" & gPigOpeReusoLinea & "' " _
                 & "                                then (select MCL.nNroCalen from MovCol MCL where MCL.nMovNro = M.nMovNro " _
                 & "                                           and MCL.cOpeCod in ('" & gPigOpeReusoLinea & "')) " _
                 & "                                else (select MCL.nNroCalen from MovCol MCL where MCL.nMovNro = M.nMovNro) end) " _
                 & "    FROM Mov M LEFT JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
                 & "                           LEFT JOIN OpeTpo Ope ON M.cOpeCod = Ope.cOpeCod " _
                 & "INNER JOIN PRODUCTOPERSONA PP ON MC.CCTACOD = PP.CCTACOD INNER JOIN PERSONA P ON P.cPersCod = PP.CPERSCOD " _
                 & "INNER JOIN PRODUCTO C ON C.CCTACOD = PP.CCTACOD " _
                 & " WHERE M.cMovNro LIKE '" & psFecTrans & "%' AND " _
                 & " M.nMovFlag = " & gMovFlagVigente & " AND M.cOpeCod like '" & pnOpeCod & "%' AND " _
                 & " Substring(M.cOpeCod ,6,1) NOT IN (" & lsCod & ") AND Substring(M.cOpeCod,3,1) <> '" & Mid(gPigOpeExtDesembolso, 3, 1) & "' " _
        
        If psCtaCod <> "@" Then
            sql = sql & " AND MC.cCtaCod = '" & psCtaCod & "' "
        End If
        
        If psCodAge <> "@" Then
            sql = sql & " AND Substring(M.cMovNro, 18,2)  = '" & psCodAge & "' "
        End If
        
        sql = sql & " GROUP BY M.cMovNro, M.nMovNro, M.cOpeCod, M.nMovFlag, MC.cCtaCod, Ope.cOpeDesc, nSaldoCap, cPersNombre, nTransacc "
        sql = sql & " ORDER BY M.nMovNro DESC  "
    
    Else

        sql = "SELECT M.nMovNro, cMovNro, M.cOpeCod, SUM(nMonto) nMonto, cOpeDesc, cPersNombre, nTransacc = 0, " _
            & "M.nMovFlag, nSaldoCap = 0,  NroCalen = (SELECT TOP 1 MCL.nNroCalen from MovCol MCL " _
            & "WHERE MCL.nMovNro = M.nMovNro and MCL.cOpeCod in ('" & gPigOpePagoSobrantes & "')), " _
            & "cCtaCod = (SELECT TOP 1 MCL.cCtaCod from MovCol MCL where MCL.nMovNro = M.nMovNro " _
            & "AND MCL.cOpeCod in ('" & gPigOpePagoSobrantes & "'))" _
            & "FROM Mov M LEFT JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
            & "LEFT JOIN OpeTpo Ope ON M.cOpeCod = Ope.cOpeCod " _
            & "INNER JOIN PRODUCTOPERSONA PP ON MC.CCTACOD = PP.CCTACOD " _
            & "INNER JOIN PERSONA P ON P.CPERSCOD = PP.CPERSCOD INNER JOIN PRODUCTO C ON C.CCTACOD = PP.CCTACOD " _
            & "WHERE M.cMovNro LIKE '" & psFecTrans & "%' AND  M.nMovFlag = " & gMovFlagVigente & " AND M.cOpeCod like '" & pnOpeCod _
            & "%' AND SUBSTRING(M.cOpeCod ,6,1) NOT IN (99) AND SUBSTRING(M.cOpeCod,3,1) <> '9' " _

        If psCtaCod <> "@" Then
            sql = sql & " AND MC.cCtaCod = '" & psCtaCod & "' "
        End If

        If psCodAge <> "@" Then
            sql = sql & " AND Substring(M.cMovNro, 18,2)    = '" & psCodAge & "' "
        End If

        sql = sql & "GROUP BY M.nMovNro, cMovNro, M.cOpeCod, cOpeDesc, cPersNombre, nMovFlag"

    End If
    
    Set lrValida = oConn.CargaRecordSet(sql)
    If lrValida Is Nothing Then
        MsgBox "No se encontraron datos para el Extorno", vbInformation, "Aviso"
        Set lrValida = Nothing
        Exit Function
    End If
    
    Set nBuscaOperacionesCredPigParaExtorno = lrValida
    Set lrValida = Nothing
    
End Function

'**********************************************************************************
' CAFF - 03/10/2002 - Valida Si no se genero el Fin de Retasacion y 1era Subasta
'**********************************************************************************
Public Function ValidaPrimSubasta(ByVal pdFecha As Date) As Boolean
Dim rsValida As Recordset

sql = "SELECT dInicio, nTipoProceso FROM COLOCPIGREMATE " _
      & "WHERE datediff(dd, '" & FechaHora(pdFecha) & "', dInicio ) <= 0 AND nTipoProceso = 0 " _
      & "ORDER BY nRemate DESC"

Set rsValida = oConn.CargaRecordSet(sql)

If rsValida.EOF And rsValida.BOF Then
    ValidaPrimSubasta = False
Else
    ValidaPrimSubasta = True
End If
    
    Set rsValida = Nothing

End Function

'*****************************************************
' CAFF - 03/10/2002 - Valida Si 2da o 3era Subasta
'*****************************************************
Public Function ValidaSegTerSubasta(ByVal pdFecha As Date, ByVal pnTipoProceso As Integer) As Boolean
Dim rsValida As Recordset

sql = "SELECT dInicio, nTipoProceso FROM COLOCPIGREMATE " _
      & "WHERE '" & FechaHora(pdFecha) & "' BETWEEN dInicio AND dFin AND nTipoProceso = " & pnTipoProceso _
      & " ORDER BY nRemate DESC"

Set rsValida = oConn.CargaRecordSet(sql)

If rsValida Is Nothing Then
    ValidaSegTerSubasta = False
Else
    ValidaSegTerSubasta = True
End If

    Set rsValida = Nothing

End Function

Public Function ValidaAdjudicacion(ByVal pdFecha As Date) As Boolean

Dim rsValida As Recordset

sql = "SELECT nAdjudicado FROM COLOCPIGREMATE " _
      & "WHERE DATEDIFF(dd, dInicio, '" & FechaHora(pdFecha) & "') >= 0 AND DATEDIFF(dd, dFin, '" & _
      FechaHora(pdFecha) & "') <= 0 AND nAdjudicado = 1 " _
      & " ORDER BY nRemate DESC"

Set rsValida = oConn.CargaRecordSet(sql)

If rsValida Is Nothing Then
    ValidaAdjudicacion = False
Else
    ValidaAdjudicacion = True
End If

    Set rsValida = Nothing

End Function
Public Function ValidaSiRemesaAdjudica(ByVal pnRemate As Long) As Boolean
Dim Rs As Recordset

sql = "SELECT * FROM COLOCPIGJOYATASACION WHERE cCtaCod IN " _
    & "(SELECT cCtaCod FROM COLOCPIGPROCESO WHERE nRemate = " & pnRemate _
    & " AND nSituacionPieza = 1) AND nUbicaPieza <> " & gPigUbicaBovValor

Set Rs = oConn.CargaRecordSet(sql)

If (Rs.EOF And Rs.BOF) Then
    ValidaSiRemesaAdjudica = True
Else
    ValidaSiRemesaAdjudica = False
End If

End Function
Public Function ValidaCancelacionSobrante(ByVal psCodCta As String, Optional ByVal psCodPers As String = "") As Recordset
Dim Rs As Recordset

If psCodPers = "" Then

    sql = "SELECT cPersCod FROM PRODUCTOPERSONA WHERE cCtaCod = '" & psCodCta & "'"
    Set Rs = oConn.CargaRecordSet(sql)
    
    If Not (Rs.EOF And Rs.BOF) Then
           psCodPers = Rs!cPersCod
    End If
    
End If

    sql = "SELECT S.cCtaCod, PP.cPersCod, nSobrante, cCtaAbono, nRemate, C.cPersNombre, " _
        & "NroDNI = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = C.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ) " _
        & "FROM COLOCPIGSOBRANTE S INNER JOIN PRODUCTOPERSONA PP ON " _
        & "S.cCtaCod = PP.cCtaCod INNER JOIN PRODUCTO P ON P.cCtaCod = PP.cCtaCod " _
        & "INNER JOIN PERSONA C ON C.cPersCod = PP.cPersCod " _
        & "WHERE nPrdEstado NOT IN (1300, 1301, 1400, 1401) AND PP.cPersCod = '" & psCodPers & "' AND nEstadoSobrante = 0"

    Set ValidaCancelacionSobrante = oConn.CargaRecordSet(sql)

End Function

Public Function ValidaSinoOpePrevias(ByVal pnMovNro As Long, ByVal psCodCta As String) As Boolean
Dim Rs As Recordset

    sql = "SELECT MC.nMovNro FROM MOVCOL MC INNER JOIN MOV M ON MC.nMovNro = M.nMovNro " _
        & "WHERE CCTACOD = '" & psCodCta & "' AND MC.nMovNro > " & pnMovNro _
        & " AND nMovFLag <> 2 AND SUBSTRING(M.cOpeCod,1,3) <> '159'"
        
    Set Rs = oConn.CargaRecordSet(sql)
    
    If Not (Rs.EOF And Rs.BOF) Then
        ValidaSinoOpePrevias = True
    Else
        ValidaSinoOpePrevias = False
    End If
    Set Rs = Nothing

End Function
'***********************************************************************
' VALIDA EL DETALLE DE RESCATES DE SOBRANTE DE PIEZAS DE UN CONTRATO
' JAPP - 30/06/03
'***********************************************************************
Public Function nValidaSobrantePieza(ByVal psCtaCod As String) As Recordset
Dim lrValida As ADODB.Recordset
Dim oRema As DPigFunciones
Dim lsRemate As Integer

    Set oRema = New DPigFunciones
    lsRemate = oRema.GetRemate(psCtaCod)
    Set oRema = Nothing
    
    sql = "SELECT t1.cCtaCod, " _
        & "SUM(nTasacion + nTasacionAdicional) Tasacion " _
        & "From ColocPigProceso  t1 Inner Join ColocPigJoyaTasacion t2 " _
        & "On t1.cCtaCod = t2.cCtaCod and t1.nItemPieza = t2.nItemPieza " _
        & "Where t1.nTipoProceso = 4 And t2.nTipoTasacion = 3 " _
        & "And t1.nSituacionPieza = 1 And t1.nRemate = " & lsRemate & " " _
        & "And t1.cCtaCod = '" & psCtaCod & "'" _
        & "Group by t1.cCtaCod"
        
    Set lrValida = oConn.CargaRecordSet(sql)
    
    If lrValida.BOF And lrValida.EOF Then
        MsgBox " No se ha encontrado Contrato ", vbInformation, " Aviso "
    End If
    
    Set nValidaSobrantePieza = lrValida
    Set lrValida = Nothing

End Function



