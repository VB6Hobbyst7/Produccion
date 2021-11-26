VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCaptaFideicomiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim oError As COMConecta.COMErrorHandling
Dim oConn As COMConecta.DCOMConecta
Dim lsSQL As String

'************************************************
' CAFF - 12/11/2002 - Inicializacion de la Clase
'************************************************
Private Sub Class_Initialize()
Dim oIni As COMConecta.DCOMClasIni

Dim sConn As String

    Set oIni = New COMConecta.DCOMClasIni
    sConn = oIni.CadenaConexion
    Set oIni = Nothing
    
Set oError = New COMConecta.COMErrorHandling
Set oConn = New COMConecta.DCOMConecta
If oConn.AbreConexion(sConn) = False Then
    Call oError.RaiseError(oError.MyUnhandledError, "DCapFideicomiso:Initialize. Error en Conexion a Base de datos")
End If

End Sub

Private Sub Class_Terminate()
    oConn.CierraConexion
    Set oConn = Nothing
    Set oError = Nothing
End Sub

Public Sub dBeginTrans()
    oConn.BeginTrans
End Sub

Public Sub dRollbackTrans()
    oConn.RollbackTrans
End Sub

Public Sub dCommitTrans()
    oConn.CommitTrans
End Sub

Public Function dGetCuentasFideicomiso(ByVal psNombre As String) As ADODB.Recordset
Dim rs As ADODB.Recordset

    On Error GoTo dErrorGetDatos
    
    lsSQL = "SELECT cCtaCod FROM FIDEICOMISO WHERE cNombre = '" & psNombre & "' AND nNroCuotasPend <> 0"
       
    Set rs = oConn.CargaRecordSet(lsSQL)
    
    Set dGetCuentasFideicomiso = rs
    Set rs = Nothing
    
    Exit Function
    
dErrorGetDatos:
    Err.Raise Err.Number, "Error en << dGetCuentasFideicomiso >>  " + Err.Description, "Aviso"
    
End Function
' FONCODES
' CRSF - 23/06
Public Function dGetCuentasFoncodes(ByVal psNombre As String) As ADODB.Recordset
Dim rs As ADODB.Recordset

    On Error GoTo dErrorGetDatos
    
    lsSQL = "SELECT cCtaCod FROM FONCODES WHERE cNombre = '" & psNombre & "' AND nNroCuotasPend <> 0"
       
    Set rs = oConn.CargaRecordSet(lsSQL)
    
    Set dGetCuentasFoncodes = rs
    Set rs = Nothing
    
    Exit Function
    
dErrorGetDatos:
    Err.Raise Err.Number, "Error en << dGetCuentasFoncodes >>  " + Err.Description, "Aviso"
    
End Function
' PLAN BICI
' CRSF - 24/06
Public Function dGetCuentasPlanBici(ByVal psNombre As String) As ADODB.Recordset
Dim rs As ADODB.Recordset

    On Error GoTo dErrorGetDatos
    
    lsSQL = "SELECT cCtaCod FROM PLANBICI WHERE cNombre = '" & psNombre & "' AND nNroCuotasPend <> 0"
       
    Set rs = oConn.CargaRecordSet(lsSQL)
    
    Set dGetCuentasPlanBici = rs
    Set rs = Nothing
    
    Exit Function
    
dErrorGetDatos:
    Err.Raise Err.Number, "Error en << dGetCuentasPlanBici >>  " + Err.Description, "Aviso"
    
End Function

Public Function dGetDatosFideicomiso(ByVal psCodCta As String) As ADODB.Recordset
Dim rs As ADODB.Recordset

    On Error GoTo dErrorGetDatos
        
    lsSQL = "SELECT cCtaCod, cNombre,  nIndCalendario, nNroCuotasPend, (nMontoTotalDeuda  - ISNULL(nMontoTotalPagado,0)) nMontoTotalDeuda " _
          & "FROM FIDEICOMISO WHERE cCtaCod = '" & psCodCta & "'"
        
    Set rs = oConn.CargaRecordSet(lsSQL)
    
    Set dGetDatosFideicomiso = rs
    Set rs = Nothing
    
    Exit Function
    
dErrorGetDatos:
    Err.Raise Err.Number, "Error en << dGetDatosFideicomiso >>  ", Err.Description
End Function
' FONCODES
' CRSF - 24/06
Public Function dGetDatosFoncodes(ByVal psCodCta As String) As ADODB.Recordset
Dim rs As ADODB.Recordset

    On Error GoTo dErrorGetDatos
        
    lsSQL = "SELECT cCtaCod, cNombre,  nIndCalendario, nNroCuotasPend, (nMontoTotalDeuda  - ISNULL(nMontoTotalPagado,0)) nMontoTotalDeuda " _
          & "FROM Foncodes WHERE cCtaCod = '" & psCodCta & "'"
        
    Set rs = oConn.CargaRecordSet(lsSQL)
    
    Set dGetDatosFoncodes = rs
    Set rs = Nothing
    
    Exit Function
    
dErrorGetDatos:
    Err.Raise Err.Number, "Error en << dGetDatosFoncodes >>  ", Err.Description
End Function
' PLAN BICI
' CRSF - 24/06
Public Function dGetDatosPlanBici(ByVal psCodCta As String) As ADODB.Recordset
Dim rs As ADODB.Recordset

    On Error GoTo dErrorGetDatos
        
    lsSQL = "SELECT cCtaCod, cNombre,  nIndCalendario, nNroCuotasPend, (nMontoTotalDeuda  - ISNULL(nMontoTotalPagado,0)) nMontoTotalDeuda " _
          & "FROM PLANBICI WHERE cCtaCod = '" & psCodCta & "'"
        
    Set rs = oConn.CargaRecordSet(lsSQL)
    
    Set dGetDatosPlanBici = rs
    Set rs = Nothing
    
    Exit Function
    
dErrorGetDatos:
    Err.Raise Err.Number, "Error en << dGetDatosPlanBici >>  ", Err.Description
End Function

Public Function dGetDatosFideicomisoDet(ByVal psCodCta As String) As ADODB.Recordset
Dim rs As ADODB.Recordset

    On Error GoTo dErrorGetDatos
    
    lsSQL = " SELECT S.nNroCuota, Sum(S.nCapital - nCapitalPag) nCapital, Sum(S.nInteres - nInteresPag) nInteres, " _
        & "Sum(S.nPortes - nPortesPag) nPortes, Sum(S.nInteresComp - nInteresCompPag) nInteresComp, " _
        & "Sum(S.nMora-nMoraPag) nMora, Sum(S.nProtesto-nProtestoPag) nProtesto, Sum(S.nComVcdo - nComVcdoPag) nComVcdo, " _
        & "Sum(S.nGastos - nGastosPag) nGastos, FD.dFecVenc " _
        & "FROM (" _
            & "SELECT nNroCuota, nCapital, nInteres, nPortes, nInteresComp, nMora, nProtesto, nComVcdo, nGastos, " _
            & "nCapitalPag = 0, nInteresPag = 0, nPortesPag = 0, nInteresCompPag = 0, nMoraPag = 0, nProtestoPag = 0, " _
            & "nComVcdoPag = 0, nGastosPag = 0 " _
            & "        FROM FIDEICOMISODET WHERE cCtaCod = '" & psCodCta & "' " _
            & "UNION " _
            & "SELECT nNroCuota, nCapital = 0, nInteres = 0, nPortes = 0, nInteresComp = 0, nMora = 0, nProtesto = 0, " _
            & "nComVcdo = 0, nGastos = 0, ISNULL(SUM(nCapitalPag),0) nCapitalPag, ISNULL(SUM(nInteresPag),0) nInteresPag, " _
            & "ISNULL(SUM(nPortesPag),0) nPortesPag, ISNULL(SUM(nInteresCompPag),0) nInteresCompPag, " _
            & "ISNULL(SUM(nMoraPag),0) nMoraPag, ISNULL(SUM(nProtestoPag),0) nProtestoPag , " _
            & "ISNULL(SUM(nComVcdoPag),0) nComVcdoPag, ISNULL(SUM(nGastosPag),0) nGastosPag " _
            & "         FROM FIDEICOMISOMOV WHERE nMovFlag <> 2 AND cCtaCod = '" & psCodCta & "' " _
            & "GROUP BY cCtaCod, nNroCuota) S " _
        & "INNER JOIN FIDEICOMISODET FD ON S.nNroCuota = FD.nNroCuota " _
        & "      WHERE FD.cCtaCod = '" & psCodCta & "' Group By S.nNroCuota, dFecVenc " _
        & "      HAVING SUM(S.nCapital - nCapitalPag + S.nInteres - nInteresPag + S.nPortes - nPortesPag + " _
        & "           S.nInteresComp - nInteresCompPag + S.nMora-nMoraPag + S.nProtesto-nProtestoPag + " _
        & "           S.nComVcdo - nComVcdoPag + S.nGastos - nGastosPag) > 0 " _
        & "      ORDER BY S.nNroCuota"
    
    Set rs = oConn.CargaRecordSet(lsSQL)
    
    Set dGetDatosFideicomisoDet = rs
    Set rs = Nothing
    
    Exit Function
    
dErrorGetDatos:
    Err.Raise Err.Number, "Error en << dGetDatosFideicomisoDet >>  ", Err.Description
    
End Function
' FONCODES
' CRSF - 24/06
Public Function dGetDatosFoncodesDet(ByVal psCodCta As String) As ADODB.Recordset
Dim rs As ADODB.Recordset

    On Error GoTo dErrorGetDatos
    
    lsSQL = " SELECT S.nNroCuota, Sum(S.nCapital - nCapitalPag) nCapital, Sum(S.nInteres - nInteresPag) nInteres, " _
        & "Sum(S.nPortes - nPortesPag) nPortes, Sum(S.nInteresComp - nInteresCompPag) nInteresComp, " _
        & "Sum(S.nMora-nMoraPag) nMora, Sum(S.nProtesto-nProtestoPag) nProtesto, Sum(S.nComVcdo - nComVcdoPag) nComVcdo, " _
        & "Sum(S.nGastos - nGastosPag) nGastos, FD.dFecVenc " _
        & "FROM (" _
            & "SELECT nNroCuota, nCapital, nInteres, nPortes, nInteresComp, nMora, nProtesto, nComVcdo, nGastos, " _
            & "nCapitalPag = 0, nInteresPag = 0, nPortesPag = 0, nInteresCompPag = 0, nMoraPag = 0, nProtestoPag = 0, " _
            & "nComVcdoPag = 0, nGastosPag = 0 " _
            & "        FROM FoncodesDET WHERE cCtaCod = '" & psCodCta & "' " _
            & "UNION " _
            & "SELECT nNroCuota, nCapital = 0, nInteres = 0, nPortes = 0, nInteresComp = 0, nMora = 0, nProtesto = 0, " _
            & "nComVcdo = 0, nGastos = 0, ISNULL(SUM(nCapitalPag),0) nCapitalPag, ISNULL(SUM(nInteresPag),0) nInteresPag, " _
            & "ISNULL(SUM(nPortesPag),0) nPortesPag, ISNULL(SUM(nInteresCompPag),0) nInteresCompPag, " _
            & "ISNULL(SUM(nMoraPag),0) nMoraPag, ISNULL(SUM(nProtestoPag),0) nProtestoPag , " _
            & "ISNULL(SUM(nComVcdoPag),0) nComVcdoPag, ISNULL(SUM(nGastosPag),0) nGastosPag " _
            & "         FROM FoncodesMOV WHERE nMovFlag <> 2 AND cCtaCod = '" & psCodCta & "' " _
            & "GROUP BY cCtaCod, nNroCuota) S " _
        & "INNER JOIN FoncodesDET FD ON S.nNroCuota = FD.nNroCuota " _
        & "      WHERE FD.cCtaCod = '" & psCodCta & "' Group By S.nNroCuota, dFecVenc " _
        & "      HAVING SUM(S.nCapital - nCapitalPag + S.nInteres - nInteresPag + S.nPortes - nPortesPag + " _
        & "           S.nInteresComp - nInteresCompPag + S.nMora-nMoraPag + S.nProtesto-nProtestoPag + " _
        & "           S.nComVcdo - nComVcdoPag + S.nGastos - nGastosPag) > 0 " _
        & "      ORDER BY S.nNroCuota"
    
    Set rs = oConn.CargaRecordSet(lsSQL)
    
    Set dGetDatosFoncodesDet = rs
    Set rs = Nothing
    
    Exit Function
    
dErrorGetDatos:
    Err.Raise Err.Number, "Error en << dGetDatosFoncodesDet >>  ", Err.Description
    
End Function
' PLAN BICI
' CRSF - 24/06
Public Function dGetDatosPlanBiciDet(ByVal psCodCta As String) As ADODB.Recordset
Dim rs As ADODB.Recordset

    On Error GoTo dErrorGetDatos
    
    lsSQL = " SELECT S.nNroCuota, Sum(S.nCapital - nCapitalPag) nCapital, Sum(S.nInteres - nInteresPag) nInteres, " _
        & "Sum(S.nPortes - nPortesPag) nPortes, Sum(S.nInteresComp - nInteresCompPag) nInteresComp, " _
        & "Sum(S.nMora-nMoraPag) nMora, Sum(S.nProtesto-nProtestoPag) nProtesto, Sum(S.nComVcdo - nComVcdoPag) nComVcdo, " _
        & "Sum(S.nGastos - nGastosPag) nGastos, FD.dFecVenc " _
        & "FROM (" _
            & "SELECT nNroCuota, nCapital, nInteres, nPortes, nInteresComp, nMora, nProtesto, nComVcdo, nGastos, " _
            & "nCapitalPag = 0, nInteresPag = 0, nPortesPag = 0, nInteresCompPag = 0, nMoraPag = 0, nProtestoPag = 0, " _
            & "nComVcdoPag = 0, nGastosPag = 0 " _
            & "        FROM PLANBICIDET WHERE cCtaCod = '" & psCodCta & "' " _
            & "UNION " _
            & "SELECT nNroCuota, nCapital = 0, nInteres = 0, nPortes = 0, nInteresComp = 0, nMora = 0, nProtesto = 0, " _
            & "nComVcdo = 0, nGastos = 0, ISNULL(SUM(nCapitalPag),0) nCapitalPag, ISNULL(SUM(nInteresPag),0) nInteresPag, " _
            & "ISNULL(SUM(nPortesPag),0) nPortesPag, ISNULL(SUM(nInteresCompPag),0) nInteresCompPag, " _
            & "ISNULL(SUM(nMoraPag),0) nMoraPag, ISNULL(SUM(nProtestoPag),0) nProtestoPag , " _
            & "ISNULL(SUM(nComVcdoPag),0) nComVcdoPag, ISNULL(SUM(nGastosPag),0) nGastosPag " _
            & "         FROM PLANBICIMOV WHERE nMovFlag <> 2 AND cCtaCod = '" & psCodCta & "' " _
            & "GROUP BY cCtaCod, nNroCuota) S " _
        & "INNER JOIN PLANBICIDET FD ON S.nNroCuota = FD.nNroCuota " _
        & "      WHERE FD.cCtaCod = '" & psCodCta & "' Group By S.nNroCuota, dFecVenc " _
        & "      HAVING SUM(S.nCapital - nCapitalPag + S.nInteres - nInteresPag + S.nPortes - nPortesPag + " _
        & "           S.nInteresComp - nInteresCompPag + S.nMora-nMoraPag + S.nProtesto-nProtestoPag + " _
        & "           S.nComVcdo - nComVcdoPag + S.nGastos - nGastosPag) > 0 " _
        & "      ORDER BY S.nNroCuota"
    
    Set rs = oConn.CargaRecordSet(lsSQL)
    
    Set dGetDatosPlanBiciDet = rs
    Set rs = Nothing
    
    Exit Function
    
dErrorGetDatos:
    Err.Raise Err.Number, "Error en << dGetDatosPlanBiciDet >>  ", Err.Description
    
End Function

Public Sub dInsertMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)

lsSQL = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"

oConn.Ejecutar lsSQL

End Sub

Public Function GetnMovNro(ByVal sMovNro As String) As Long
Dim rsMov As ADODB.Recordset

lsSQL = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"

Set rsMov = oConn.CargaRecordSet(lsSQL)

If rsMov.EOF And rsMov.BOF Then
    GetnMovNro = 0
Else
    GetnMovNro = rsMov("nMovNro")
End If

Set rsMov = Nothing
End Function

Public Sub dInsertMovServicios(ByVal pnMovNro As Long, ByVal psNumRecibo As String, ByVal psCodigo As String, _
                               ByVal pnMonto As Currency, ByVal pnMoneda As String, Optional ByVal pnFlag As MovFlag = gMovFlagVigente)

    
    On Error GoTo dErrorInsertMovServicios
    'todocompleta
    lsSQL = "INSERT MOVSERVICIOS (nMovNro, cNumRecibo, cOpeCod, nMonto, nMoneda, nFlag) VALUES (" _
            & pnMovNro & ",'" & psNumRecibo & "','" & psCodigo & "'," & pnMonto & "," & pnMoneda & "," & pnFlag & ")"
'    lsSQL = "INSERT MOVSERVICIOS (nMovNro, cNumRecibo, cCodigo, nMonto, nMoneda, nFlag) VALUES (" _
'            & pnMovNro & ",'" & psNumRecibo & "','" & psCodigo & "'," & pnMonto & "," & pnMoneda & "," & pnFlag & ")"
    oConn.Ejecutar lsSQL
    
    Exit Sub
    
dErrorInsertMovServicios:
    Err.Raise Err.Number, "Error en << dInsertMovServicios >> ", Err.Description
    
End Sub

Public Sub dInsertMovServiciosDet(ByVal psCodCta As String, ByVal pnMovNro As Long, ByVal pnPrdConcepto As Integer, ByVal pnNroCuota As Integer, _
                               ByVal pnMonto As Currency, ByVal pscodope As String)
 
    On Error GoTo dErrorInsertMovServiciosDet
    'TODOCOMPLETA
    lsSQL = "INSERT MOVSERVICIOSDET (nMovNro, cNumRecibo, nPrdConcepto, nNroCuota,copecod, nMonto) VALUES (" _
            & pnMovNro & ",'" & psCodCta & "'," & pnPrdConcepto & "," & pnNroCuota & ",'" & pscodope & "'," & pnMonto & ")"
            
'    lsSQL = "INSERT MOVSERVICIOSDET (nMovNro, cNumRecibo, nPrdConcepto, nNroCuota, nMonto) VALUES (" _
'            & pnMovNro & ",'" & psCodCta & "'," & pnPrdConcepto & "," & pnNroCuota & "," & pnMonto & ")"
    oConn.Ejecutar lsSQL
    
    Exit Sub
    
dErrorInsertMovServiciosDet:
    Err.Raise Err.Number, "Error en << dInsertMovServiciosDet >> ", Err.Description
    
End Sub

Public Sub dInsertFideicomisoMov(ByVal psCodCta As String, ByVal pnTipoPag As Integer, ByVal pnNroCuota As Integer, _
                ByVal pnCapital As Currency, ByVal pnInteres As Currency, ByVal pnPortes As Currency, ByVal pnIntCom As Currency, _
                ByVal pnMora As Currency, ByVal pnProtesto As Currency, ByVal pnComVcdo As Currency, ByVal pnGasto As Currency, _
                ByVal pnMovNro As Long, Optional ByVal pnMovFlag As MovFlag = gMovFlagVigente)

    On Error GoTo dErrorInsertFideicomisoMov
    
    lsSQL = "INSERT FIDEICOMISOMOV (cCtaCod, cTipoPago, nNroCuota, nCapitalPag, nInteresPag, nPortesPag, nInteresCompPag, " _
                                & " nMoraPag, nProtestoPag, nComVcdoPag, nGastosPag, nNroMov, nMovFlag) " _
        & " VALUES ('" & psCodCta & "'," & pnTipoPag & "," & pnNroCuota & "," & pnCapital & "," & pnInteres & "," & pnPortes _
                       & "," & pnIntCom & "," & pnMora & "," & pnProtesto & "," & pnComVcdo & "," & pnGasto & "," & pnMovNro _
                       & "," & pnMovFlag & ")"
    
    oConn.Ejecutar lsSQL
    
    Exit Sub
   
   
dErrorInsertFideicomisoMov:
    Err.Raise Err.Number, "Error en << dInsertFideicomisoMov >> ", Err.Description
    
End Sub
'FONCODES
' 01/07 - CMCPL
Public Sub dInsertFONCODESMov(ByVal psCodCta As String, ByVal pnTipoPag As Integer, ByVal pnNroCuota As Integer, _
                ByVal pnCapital As Currency, ByVal pnInteres As Currency, ByVal pnPortes As Currency, ByVal pnIntCom As Currency, _
                ByVal pnMora As Currency, ByVal pnProtesto As Currency, ByVal pnComVcdo As Currency, ByVal pnGasto As Currency, _
                ByVal pnMovNro As Long, Optional ByVal pnMovFlag As MovFlag = gMovFlagVigente)

    On Error GoTo dErrorInsertFONCODESMov
    
    lsSQL = "INSERT FONCODESMOV (cCtaCod, cTipoPago, nNroCuota, nCapitalPag, nInteresPag, nPortesPag, nInteresCompPag, " _
                                & " nMoraPag, nProtestoPag, nComVcdoPag, nGastosPag, nNroMov, nMovFlag) " _
        & " VALUES ('" & psCodCta & "'," & pnTipoPag & "," & pnNroCuota & "," & pnCapital & "," & pnInteres & "," & pnPortes _
                       & "," & pnIntCom & "," & pnMora & "," & pnProtesto & "," & pnComVcdo & "," & pnGasto & "," & pnMovNro _
                       & "," & pnMovFlag & ")"
    
    oConn.Ejecutar lsSQL
    
    Exit Sub
   
   
dErrorInsertFONCODESMov:
    Err.Raise Err.Number, "Error en << dInsertFONCODESMov >> ", Err.Description
    
End Sub
' PLAN BICI
' CRSF - 01/07
Public Sub dInsertPLANBICIMov(ByVal psCodCta As String, ByVal pnTipoPag As Integer, ByVal pnNroCuota As Integer, _
                ByVal pnCapital As Currency, ByVal pnInteres As Currency, ByVal pnPortes As Currency, ByVal pnIntCom As Currency, _
                ByVal pnMora As Currency, ByVal pnProtesto As Currency, ByVal pnComVcdo As Currency, ByVal pnGasto As Currency, _
                ByVal pnMovNro As Long, Optional ByVal pnMovFlag As MovFlag = gMovFlagVigente)

    On Error GoTo dErrorInsertPLANBICIMov
    
    lsSQL = "INSERT PLANBICIMOV (cCtaCod, cTipoPago, nNroCuota, nCapitalPag, nInteresPag, nPortesPag, nInteresCompPag, " _
                                & " nMoraPag, nProtestoPag, nComVcdoPag, nGastosPag, nNroMov, nMovFlag) " _
        & " VALUES ('" & psCodCta & "'," & pnTipoPag & "," & pnNroCuota & "," & pnCapital & "," & pnInteres & "," & pnPortes _
                       & "," & pnIntCom & "," & pnMora & "," & pnProtesto & "," & pnComVcdo & "," & pnGasto & "," & pnMovNro _
                       & "," & pnMovFlag & ")"
    
    oConn.Ejecutar lsSQL
    
    Exit Sub
   
   
dErrorInsertPLANBICIMov:
    Err.Raise Err.Number, "Error en << dInsertPLANBICIMov >> ", Err.Description
    
End Sub
Public Sub dUpdateFideicomiso(ByVal psCodCta As String, ByVal pnCuotasPag As Integer, ByVal pnMontoPag As Currency)

     On Error GoTo dErrorUpdateFideicomiso
     
     lsSQL = "UPDATE FIDEICOMISO SET nNroCuotasPend = nNroCuotasPend - " & pnCuotasPag & ", nMontoTotalPagado = nMontoTotalPagado + " _
            & pnMontoPag & " WHERE cCtaCod = '" & psCodCta & "'"
            
    oConn.Ejecutar lsSQL
     
    Exit Sub
    
dErrorUpdateFideicomiso:
    Err.Raise Err.Number, "Error en << dUpdateFideicomiso >> ", Err.Description
     
End Sub
' FONCODES
' CMCPL - CRSF - 01/07
Public Sub dUpdateFONCODES(ByVal psCodCta As String, ByVal pnCuotasPag As Integer, ByVal pnMontoPag As Currency)

     On Error GoTo dErrorUpdateFoncodes
     
     lsSQL = "UPDATE FONCODES SET nNroCuotasPend = nNroCuotasPend - " & pnCuotasPag & ", nMontoTotalPagado = nMontoTotalPagado + " _
            & pnMontoPag & " WHERE cCtaCod = '" & psCodCta & "'"
            
    oConn.Ejecutar lsSQL
     
    Exit Sub
    
dErrorUpdateFoncodes:
    Err.Raise Err.Number, "Error en << dUpdateFONCODES>> ", Err.Description
     
End Sub
' PLAN BICI
' CMCPL - CRSF - 01/07
Public Sub dUpdatePLANBICI(ByVal psCodCta As String, ByVal pnCuotasPag As Integer, ByVal pnMontoPag As Currency)

     On Error GoTo dErrorUpdatePLANBICI
     
     lsSQL = "UPDATE PLANBICI SET nNroCuotasPend = nNroCuotasPend - " & pnCuotasPag & ", nMontoTotalPagado = nMontoTotalPagado + " _
            & pnMontoPag & " WHERE cCtaCod = '" & psCodCta & "'"
            
    oConn.Ejecutar lsSQL
     
    Exit Sub
    
dErrorUpdatePLANBICI:
    Err.Raise Err.Number, "Error en << dUpdatePLANBICI>> ", Err.Description
     
End Sub

Public Function dGetDatosPersona(Optional ByVal psNombre As String = "@", Optional ByVal psDocId As String = "@") As ADODB.Recordset

     On Error GoTo dErrorGetDatosPersona
     
     lsSQL = "SELECT DISTINCT cNombre, cNumDocId FROM FIDEICOMISO "
     
     If psNombre <> "@" Then
        lsSQL = lsSQL & " WHERE cNombre like '" & psNombre & "%'"
     End If
     
     If psDocId <> "@" Then
        lsSQL = lsSQL & " WHERE cNumDocId like '" & psDocId & "%'"
     End If
     
    Set dGetDatosPersona = oConn.Ejecutar(lsSQL)
     
    Exit Function
    
dErrorGetDatosPersona:
    Err.Raise Err.Number, "Error en << dGetDatosPersona >> ", Err.Description

End Function
' FONCODES
' CRSF - 24/06
Public Function dGetDatosPersonaFoncodes(Optional ByVal psNombre As String = "@", Optional ByVal psDocId As String = "@") As ADODB.Recordset

     On Error GoTo dErrorGetDatosPersona
     lsSQL = "SELECT DISTINCT cNombre, cNumDocId FROM FONCODES "
     
     If psNombre <> "@" Then
        lsSQL = lsSQL & " WHERE cNombre like '" & psNombre & "%'"
     End If
     If psDocId <> "@" Then
        lsSQL = lsSQL & " WHERE cNumDocId like '" & psDocId & "%'"
     End If
     Set dGetDatosPersonaFoncodes = oConn.Ejecutar(lsSQL)
     Exit Function
    
dErrorGetDatosPersona:
    Err.Raise Err.Number, "Error en << dGetDatosPersona >> ", Err.Description

End Function
' PLAN BICI
' CRSF - 24/06
Public Function dGetDatosPersonaPlan(Optional ByVal psNombre As String = "@", Optional ByVal psDocId As String = "@") As ADODB.Recordset

     On Error GoTo dErrorGetDatosPersonaPlan
     
     lsSQL = "SELECT DISTINCT cNombre, cNumDocId FROM PLANBICI "
     
     If psNombre <> "@" Then
        lsSQL = lsSQL & " WHERE cNombre like '" & psNombre & "%'"
     End If
     
     If psDocId <> "@" Then
        lsSQL = lsSQL & " WHERE cNumDocId like '" & psDocId & "%'"
     End If
     
    Set dGetDatosPersonaPlan = oConn.Ejecutar(lsSQL)
     
    Exit Function
    
dErrorGetDatosPersonaPlan:
    Err.Raise Err.Number, "Error en << dGetDatosPersonaPlan >> ", Err.Description

End Function

Public Sub dUpdateFideicomisoMov(ByVal pnMovNro As Long, ByVal pnMovFlag As Integer)

    On Error GoTo dErrorUpdateFideicomisoDet
    
    lsSQL = "UPDATE FIDEICOMISOMOV SET nMovFlag = " & pnMovFlag & " WHERE nNroMov = " & pnMovNro

    oConn.Ejecutar lsSQL
    
    Exit Sub
    
dErrorUpdateFideicomisoDet:
    Err.Raise Err.Number, "Error en << dUpdateFideicomisoMov >> ", Err.Description
    
End Sub
' FONCODES
' CRSF - 24/06
Public Sub dUpdateFoncodesMov(ByVal pnMovNro As Long, ByVal pnMovFlag As Integer)

    On Error GoTo dErrorUpdateFoncodesDet
    
    lsSQL = "UPDATE FoncodesMOV SET nMovFlag = " & pnMovFlag & " WHERE nNroMov = " & pnMovNro

    oConn.Ejecutar lsSQL
    
    Exit Sub
    
dErrorUpdateFoncodesDet:
    Err.Raise Err.Number, "Error en << dUpdateFoncodesMov >> ", Err.Description
    
End Sub

' PLAN BICI
' CRSF - 24/06
Public Sub dUpdatePlanBiciMov(ByVal pnMovNro As Long, ByVal pnMovFlag As Integer)

    On Error GoTo dErrorUpdatePlanBiciDet
    
    lsSQL = "UPDATE PlanBiciMOV SET nMovFlag = " & pnMovFlag & " WHERE nNroMov = " & pnMovNro

    oConn.Ejecutar lsSQL
    
    Exit Sub
    
dErrorUpdatePlanBiciDet:
    Err.Raise Err.Number, "Error en << dUpdatePlanBiciMov >> ", Err.Description
    
End Sub

Public Function GetcMovNro(ByVal pnMovNro As Long) As String
Dim rs As ADODB.Recordset

    lsSQL = "SELECT cMovNro FROM Mov WHERE nMovNro ='" & pnMovNro & "'"

    Set rs = oConn.CargaRecordSet(lsSQL)
    
    If Not rs.EOF And Not rs.BOF Then
        GetcMovNro = rs!cMovNro
    End If

    Set rs = Nothing

End Function

Public Sub dInsertMovRef(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long)
    
    On Error GoTo dErrorInsertMovRef
    
    lsSQL = "INSERT INTO MovRef(nMovNro, nMovNroRef) " _
         & "VALUES (" & pnMovNro & "," & pnMovNroRef & ")"
     
    oConn.Ejecutar lsSQL
    
    Exit Sub
    
dErrorInsertMovRef:
    Err.Raise Err.Number, "Error en <<dInsertMovRef >> ", Err.Description

End Sub

Public Function dGetCuotasPag(ByVal pnMovNro As Long) As Integer
Dim rs As ADODB.Recordset

    On Error GoTo dErrorGetCuotasPag
    
    lsSQL = "SELECT COUNT(DISTINCT(nNroCuota)) nCuotasPag FROM FIDEICOMISOMOV WHERE nNroMov = " & pnMovNro
    
    Set rs = oConn.CargaRecordSet(lsSQL)
    
    If Not rs.EOF And Not rs.BOF Then
        dGetCuotasPag = rs!nCuotasPag
    End If

    Exit Function
    
dErrorGetCuotasPag:
    Err.Raise Err.Number, "Error en << dGetCuotasPag >> ", Err.Description
    
End Function
'FONCODES
'CRSF - 24/06
Public Function dGetCuotasPagFon(ByVal pnMovNro As Long) As Integer
Dim rs As ADODB.Recordset

    On Error GoTo dErrorGetCuotasPagFon
    
    lsSQL = "SELECT COUNT(DISTINCT(nNroCuota)) nCuotasPag FROM FONCODESMOV WHERE nNroMov = " & pnMovNro
    
    Set rs = oConn.CargaRecordSet(lsSQL)
    
    If Not rs.EOF And Not rs.BOF Then
        dGetCuotasPagFon = rs!nCuotasPag
    End If

    Exit Function
    
dErrorGetCuotasPagFon:
    Err.Raise Err.Number, "Error en << dGetCuotasPagFon >> ", Err.Description
    
End Function
'PLAN BICI
' CRSF - 24/06
Public Function dGetCuotasPagPlan(ByVal pnMovNro As Long) As Integer
Dim rs As ADODB.Recordset

    On Error GoTo dErrorGetCuotasPagPlan
    
    lsSQL = "SELECT COUNT(DISTINCT(nNroCuota)) nCuotasPag FROM PLANBICIMOV WHERE nNroMov = " & pnMovNro
    
    Set rs = oConn.CargaRecordSet(lsSQL)
    
    If Not rs.EOF And Not rs.BOF Then
        dGetCuotasPagPlan = rs!nCuotasPag
    End If

    Exit Function
    
dErrorGetCuotasPagPlan:
    Err.Raise Err.Number, "Error en << dGetCuotasPagPlan >> ", Err.Description
    
End Function

Public Sub dUpdateMov(ByVal pnMovNro As Long)

    On Error GoTo dErrorUpdateMov

    lsSQL = "UPDATE MOV SET nMovFlag = '" & gMovFlagExtornado & "' where nMovNro = " & pnMovNro

    oConn.Ejecutar lsSQL
    
    Exit Sub
    
dErrorUpdateMov:
    Err.Raise Err.Number, "Error en << dUpdateMov >> ", Err.Description

End Sub
