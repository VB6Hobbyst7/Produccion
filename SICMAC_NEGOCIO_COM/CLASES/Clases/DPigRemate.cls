VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DPigRemate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim lsSQL As String
Dim oConn As DConecta

'*************************************************
' INICIALIZACION DE LA CLASE
'*************************************************
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

Public Sub dBeginTrans()
    oConn.BeginTrans
End Sub

Public Sub dRollbackTrans()
    oConn.RollbackTrans
End Sub

Public Sub dCommitTrans()
    oConn.CommitTrans
End Sub

'***************************************************
'** CMCPL - Obtiene el nMovNro a partir del cMovNro
'***************************************************
Public Function dGetnMovNro(ByVal psMovNro As String) As Long
Dim Rs As Recordset

lsSQL = "SELECT nMovNro FROM Mov WHERE cMovNro ='" & psMovNro & "'"

Set Rs = oConn.CargaRecordSet(lsSQL)

If Not Rs.EOF And Not Rs.BOF Then
    dGetnMovNro = Rs!nMovNro
End If

Rs.Close
Set Rs = Nothing

End Function

'********* CAFF - 18/10/2002 *******************
Public Sub dInsertMov(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal pnMovEstado As MovEstado, _
    Optional pnMovFlag As MovFlag = gMovFlagVigente, Optional pbEjecBatch As Boolean = False)
   
On Error GoTo InsertMovErr
    
    lsSQL = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado, nMovFlag) " _
          & "VALUES ('" & psMovNro & "','" & psOpeCod & "','" & Replace(psMovDesc, "'", "''") & "'," _
          & pnMovEstado & "," & pnMovFlag & ")"
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    Exit Sub

InsertMovErr:
    Call RaiseError(MyUnhandledError, "DPigActualizaBD Insert Mov")
End Sub

'***************************************************
'CAFF - 22/08/2002 -  Insert MovCol
'***************************************************

Public Sub dInsertMovCol(ByVal pnMovNro As Long, ByVal psOperacion As String, ByVal psCuenta As String, ByVal pnNroCalend As Integer, _
        ByVal pnMonto As Currency, ByVal pnDiasMora As Integer, ByVal psMetLiq As String, ByVal pnPlazo As Integer, _
        ByVal pnCredEstado As ColocEstado, ByVal pnSaldoCap As Currency, Optional pbEjecBatch As Boolean = False)

On Error GoTo ErrInsertMovCol

    lsSQL = "INSERT MovCol (nMovNro,cOpeCod,cCtaCod,nNroCalen,nMonto,nDiasMora,cMetLiq,nPlazo, nCredEstado, nSaldoCap) " _
        & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psCuenta & "'," & pnNroCalend & "," _
        & pnMonto & "," & pnDiasMora & ",'" & psMetLiq & "'," & pnPlazo & "," & pnCredEstado & "," & pnSaldoCap & ") "

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

Exit Sub

ErrInsertMovCol:
    Call RaiseError(MyUnhandledError, "DPigActualizaBD Insert Mov Col")

End Sub

'************* CAFF - 18/10/2002 **********************
Public Sub dInsertMovRef(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long, _
    Optional pbEjecBatch As Boolean = False)
    
On Error GoTo InsertMovRefErr
    
    lsSQL = "INSERT MovRef (nMovNro,nMovNroRef) " _
          & "VALUES (" & pnMovNro & "," & pnMovNroRef & ")"
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    Exit Sub

InsertMovRefErr:
    Call RaiseError(MyUnhandledError, "DPigActualizaBD Insert Mov Ref")
End Sub
'*************************
'CAFF - 17/10/2002
'*************************
Public Function ValidaSiCtaAhorro(ByVal psCodPers As String) As String
Dim Rs As Recordset

    lsSQL = "SELECT cCtaAbono FROM COLOCPIGSOBRANTE S INNER JOIN PRODUCTO P ON S.cCtaAbono = P.cCtaCod " _
        & "     INNER JOIN PRODUCTOPERSONA PP ON PP.cCtaCod = P.cCtaCod " _
        & " WHERE nPrdEstado NOT IN  (1300, 1301, 1400, 1401) AND cPersCod = '" & psCodPers & "'"
        
Set Rs = oConn.CargaRecordSet(lsSQL)
        
If Not (Rs.EOF And Rs.BOF) Then
    ValidaSiCtaAhorro = IIf(IsNull(Rs!cCtaAbono), "", Rs!cCtaAbono)
Else
    ValidaSiCtaAhorro = ""
End If

Set Rs = Nothing

End Function
'*******************
'CAFF - 17/10/2002
'*******************
Public Sub dInsertaSobrante(ByVal psCodCta As String, ByVal pnRemate As Long, ByVal psCtaAhorro As String, _
                    ByVal pnSobrante As Currency, ByVal psMovNro As String, ByVal pnEstadoSob As Integer)


On Error GoTo ErrordInsertaSobrante
    
    lsSQL = "Insert COLOCPIGSOBRANTE(cCtaCod, nRemate, cCtaAbono, nEstadoSobrante, nSobrante, cMovNro) " _
        & "VALUES ('" & psCodCta & "'," & pnRemate & ",'" & psCtaAhorro & "'," & pnEstadoSob & "," & pnSobrante & ",'" & psMovNro & "')"
        
    oConn.Ejecutar lsSQL
    
Exit Sub

ErrordInsertaSobrante:
    Err.Raise Err.Number, "Error En Proceso ErrordInsertaSobrante", Err.Description
        
End Sub
Public Sub dUpdateSobrante(ByVal psCodCta As String, ByVal pnEstado As Integer)

    On Error GoTo dError
        
    lsSQL = "UPDATE COLOCPIGSOBRANTE SET nEstadoSobrante = " & pnEstado & " WHERE cCtaCod = '" & psCodCta & "'"
           
    oConn.Ejecutar lsSQL
    
    Exit Sub

dError:
    Err.Raise Err.Number, "Error en la Actualizacion del Estado del Sobrante <<dUpdateSobrante>>", Err.Description

End Sub

'******************************************************
' CAFF - 18/10/2002 - CAPTACIONES - BLOQUEO DE CTA
'******************************************************
Public Function NuevoBloqueoTotal(ByVal sCuenta As String, ByVal nMotivo As CaptacMotBloqueoTot, _
        cComentario As String, ByVal cMovNro As String)

    lsSQL = "INSERT ProductoBloqueos (cCtaCod,nBlqTpo,nBlqMotivo,cComentario,cMovNro,cMovNroDbl) " _
            & "VALUES ('" & sCuenta & "'," & gCapTpoBlqTotal & "," & nMotivo & ",'" & cComentario & "','" _
            & cMovNro & "',NULL)"
            
    oConn.Ejecutar lsSQL

End Function

'******************************************************
' CAFF - 18/10/2002 - CAPTACIONES - BLOQUEO DE CTA
'******************************************************
Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, nEstado As CaptacEstado, Optional ByVal pdFecha As String)

    lsSQL = "UPDATE Producto Set nPrdEstado = " & nEstado & ", dPrdEstado = '" & Format(pdFecha, "mm/dd/yyyy") _
            & "' WHERE cCtaCod = '" & sCuenta & "'"
            
    oConn.Ejecutar lsSQL
    
End Sub

'******************************************************
' CAFF - 18/10/2002 - CAPTACIONES - BLOQUEO DE CTA
'******************************************************
Public Sub AgregaCaptacEstado(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, ByVal sMovNro As String)

    lsSQL = "INSERT CaptacEstado (cCtaCod,nPrdEstado,cMovNro) " _
                & "VALUES ('" & sCuenta & "'," & nEstado & ",'" & sMovNro & "')"
                
    oConn.Ejecutar lsSQL
    
End Sub

'******************************************************
' CAFF - 18/10/2002 - CAPTACIONES - BLOQUEO DE CTA
'******************************************************
Public Sub ActualizaBloqueoTot(ByVal sCuenta As String, ByVal sComentario As String, ByVal sMovNroDbl As String, _
                               ByVal nMotivo As CaptacMotBloqueoTot, ByVal sMovNro As String)

    lsSQL = "UPDATE ProductoBloqueos SET cComentario = '" & sComentario & "', cMovNroDbl = '" & sMovNroDbl & "' " _
            & "WHERE cCtaCod = '" & sCuenta & "' AND nBlqTpo = " & gCapTpoBlqTotal & " AND nBlqMotivo = " & nMotivo _
            & " AND cMovNro = '" & sMovNro & "'"
        
    oConn.Ejecutar lsSQL
    
End Sub

'******************************************************
' CAFF - 15/10/2002 - CAPTACIONES
'******************************************************
Public Sub CapAbonoCuentaAho(ByVal sCuenta As String, ByVal nMonto As Double, ByVal nOperacion As CaptacOperacion, _
         ByVal psMovNro As String, ByVal pnMovNro As Long, ByVal sGlosa As String)

Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean
Dim nDiasTranscurridos As Long, nExtracto As Long
Dim sCodUser As String
Dim nMovNro As Long
Dim dFecSis As Date

dFecSis = CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4))
sCodUser = Right(psMovNro, 4)

'Obtiene los datos para el calculo
Set rsCta = GetDatosCuentaAho(sCuenta)
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

On Error GoTo ErrGraba

If sGlosa = "" Then sGlosa = "Abono Sobrante Remate Cuenta = " & sCuenta

'Calcula intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1

nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)

dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt + nMonto
    
    ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
    nSaldoDisp = nSaldoDisp + nMonto

    AgregaMovCap pnMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
    AgregaMovCapDet pnMovNro, nOperacion, sCuenta, gConcCapital, nMonto

    ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, psMovNro

Exit Sub

ErrGraba:
    Err.Raise Err.Number, "Error ", Err.Description

End Sub

'******************************************************
' CAFF - 16/10/2002 - CAPTACIONES - ABONO A CTA
'******************************************************
Public Function GetCapTasaInteres(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal nTipoTasa As CaptacTipoTasa, Optional nValor As Double = 0, _
        Optional sCodAge As String = "") As Double

Dim rsTasa As Recordset
Dim cOrdPag As String

If sCodAge = "" Then sCodAge = gsCodAge

    lsSQL = "SELECT nTasaValor FROM CaptacTasas WHERE nTasaProd = " & nProducto & " AND " _
        & "nTasaMon = " & nMoneda & " AND nTasaTpo = " & nTipoTasa & " And (" & nValor & " >= nValorIni " _
        & "And " & nValor & " < nValorFin) And cOrdPag = '0' And " _
        & "cCodAge = '" & sCodAge & "'"

    Set rsTasa = oConn.CargaRecordSet(lsSQL)

    If Not (rsTasa.EOF Or rsTasa.BOF) Then
        GetCapTasaInteres = rsTasa("nTasaValor")
    Else
        GetCapTasaInteres = 0
    End If

Set rsTasa = Nothing

End Function

'******************************************************
' CAFF - 16/10/2002 - CAPTACIONES - APERTURA DE CUENTA
'******************************************************
Public Function CapAperturaCuenta(ByVal nProducto As Producto, ByVal nMoneda As Moneda, ByVal cPersCod As String, _
        ByVal sAgencia As String, ByVal nTasa As Double, ByVal nSaldo As Double, ByVal dFecha As Date, ByVal nFirmas As Integer, _
        ByVal nPersoneria As PersPersoneria, ByVal sGlosa As String, ByVal nTipoCuenta As ProductoCuentaTipo, ByVal sMovNro As String, _
        ByVal nMovNro As Long, ByVal nTipoTasa As CaptacTipoTasa, ByVal nOperacion As Long) As String

Dim sCuenta As String

On Error GoTo ErrGraba

    sCuenta = AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecha, nFirmas, nPersoneria, _
                                        nTipoCuenta, sMovNro, nTipoTasa)
             
    AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, nSaldo, nSaldo
    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo

    AgregaNuevoProdPers sCuenta, cPersCod, gCapRelPersTitular
    
CapAperturaCuenta = sCuenta

Exit Function

ErrGraba:
    CapAperturaCuenta = ""
    Err.Raise Err.Number, "CapAperturaCuenta", Err.Description
End Function

'*********************************************************
' CAFF - 16/10/2002 - CAPTACIONES - CANCELACION DE CUENTA
'*********************************************************
Public Function CapCancelaCuentaAho(ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, _
        Optional nEstado As CaptacEstado = gCapEstCancelada) As Long
        

Dim oPigGraba As DPigRemate
Dim rsCta As Recordset
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nMonto As Double, dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, I As Long
Dim nExtracto As Long, nMovNro As Long, sCodUser As String
Dim dFecSis As Date
Dim nOperacion As String

nOperacion = gAhoCancSobRemate

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)

'Obtiene los datos para el calculo
Set rsCta = GetDatosCuentaAho(sCuenta)

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, True, False)
nExtracto = rsCta("nTransacc") + 1

Set rsCta = Nothing

On Error GoTo ErrGraba

nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
nMonto = nSaldoCnt + nIntAcum + nIntGanado
dUltMov = DateAdd("d", -1, dFecSis)

If sGlosa = "" Then sGlosa = "Cancelación Cuenta = " & sCuenta

ActualizaCargoCaptacion sCuenta, 0, 0, nIntAcum + nIntGanado, dUltMov, sMovNro, True
ActualizaEstadoCuenta sCuenta, gCapEstCancelada
UltimaActualizacionCuenta sCuenta, sMovNro
AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = GetnMovNro(sMovNro)
AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldoDisp + nIntAcum + nIntGanado, 0, 0
AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldoDisp
AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcInteres, nIntAcum + nIntGanado

CapCancelaCuentaAho = nMovNro

Exit Function

ErrGraba:
    Err.Raise Err.Number, "", Err.Description
End Function

'********************************************************************
' CAFF -  15/10/2002 - DATOS DE LA CUENTA PARA EL ABONO DE SOBRANTES
'********************************************************************
Public Function GetDatosCuentaAho(ByVal sCuenta As String) As Recordset
Dim rsCta As Recordset

    lsSQL = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, " _
        & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, " _
        & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.bOrdPag, A.dUltContacto, T.cConsDescripcion cEstado, " _
        & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.bInactiva, A.nSobregiro FROM Producto P " _
        & "INNER JOIN Captaciones C INNER JOIN CaptacAhorros A ON C.cCtaCod = A.cCtaCod ON " _
        & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
        & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
        & "ON C.nPrdTasaInteres = T2.nConsValor WHERE P.cCtaCod = '" & sCuenta & "' AND " _
        & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
        & "AND T2.nConsCod = " & gCaptacTipoTasa

Set rsCta = oConn.CargaRecordSet(lsSQL)
Set GetDatosCuentaAho = rsCta
Set rsCta = Nothing

End Function
'******************************************************
' CAFF - 16/10/2002 - CAPTACIONES - ABONO A CTA
'******************************************************
Public Function GetInteres(ByVal nCapital As Double, ByVal nTasa As Double, _
            ByVal nPlazo As Long, Optional nTipoInteres As TipoCalculoInteres = TpoCalcIntSimple) As Double
If nTipoInteres = TpoCalcIntSimple Then
    GetInteres = Round((nTasa / 36000) * nPlazo * nCapital, 2)
ElseIf nTipoInteres = TpoCalcIntCompuesto Then
    GetInteres = Round((((nTasa / 36000) + 1) ^ nPlazo - 1) * nCapital, 2)
End If
End Function

'******************************************************
' CAFF - 15/10/2002 - CAPTACIONES
'******************************************************
Public Sub ActualizaSaldoAnteriorAho(ByVal sCuenta As String, ByVal nSaldoAnt As Double)
    
    lsSQL = "UPDATE CAPTACAHORROS SET nSaldoAnterior = " & nSaldoAnt & "  WHERE cCtaCod = '" & sCuenta & "'"
    
    oConn.Ejecutar lsSQL
    
End Sub

Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)

lsSQL = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"

oConn.Ejecutar lsSQL

End Sub

'******************************************************
' CAFF - 15/10/2002 - CAPTACIONES
'******************************************************
Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, ByVal sCuenta As String, _
        ByVal nMonto As Double, ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)

    lsSQL = "INSERT MOVCAP (nMovNro, cOpeCod, cCtaCod, nMonto, nSaldoDisponible, nSaldoContable) " _
        & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & ")"
    
    oConn.Ejecutar lsSQL

End Sub
'******************************************************
' CAFF - 15/10/2002 - CAPTACIONES
'******************************************************
Public Sub AgregaMovCapDet(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, ByVal sCuenta As String, _
        ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double)
        
    lsSQL = "INSERT MOVCAPDET (nMovNro,cOpeCod,cCtaCod,nConceptoCod,nMonto) " _
        & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nConcepto & "," & nMonto & ")"
        
    oConn.Ejecutar lsSQL
    
End Sub

Public Function GetnMovNro(ByVal sMovNro As String) As Long
Dim rsMov As Recordset

lsSQL = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"

Set rsMov = oConn.CargaRecordSet(lsSQL)

If rsMov.EOF And rsMov.BOF Then
    GetnMovNro = 0
Else
    GetnMovNro = rsMov("nMovNro")
End If

Set rsMov = Nothing
End Function

'******************************************************
' CAFF - 15/10/2002 - CAPTACIONES
'******************************************************
Public Sub ActualizaAbonoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String)
            
    lsSQL = "UPDATE PRODUCTO SET nSaldo = nSaldo + " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"

    oConn.Ejecutar lsSQL

    lsSQL = "UPDATE CAPTACIONES SET nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
    
    oConn.Ejecutar lsSQL

End Sub

'*****************************************************************
' CAFF - 16/10/2002 - CAPTACIONES APERTURA DE CUENTA DE AHORROS
'*****************************************************************
Public Function AgregaNuevaCaptacion(ByVal nProducto As Producto, ByVal nMoneda As Moneda, ByVal sAgencia As String, _
        ByVal nTasa As Double, ByVal nSaldo As Double, ByVal dFecha As Date, ByVal nFirmas As Integer, ByVal nPersoneria As PersPersoneria, _
        ByVal nTipoCuenta As ProductoCuentaTipo, ByVal sMovNro As String, ByVal nTipoTasa As CaptacTipoTasa) As String

Dim sCuenta As String, sFecha As String
Dim sFecUltCierre As String
Dim clsGen As DGeneral

Set clsGen = New DGeneral
sCuenta = clsGen.GeneraNuevaCuenta(sAgencia, nProducto, nMoneda)
Set clsGen = Nothing

sFecha = Format$(dFecha, "mm/dd/yyyy") & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)

lsSQL = "INSERT PRODUCTO (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado, nTransacc) " _
        & "VALUES ('" & sCuenta & "'," & nTasa & "," & nSaldo & "," & gCapEstActiva & ",'" & sFecha & "',0)"

oConn.Ejecutar lsSQL

sFecUltCierre = Format$(DateAdd("d", -1, dFecha), "mm/dd/yyyy")

    lsSQL = "INSERT CAPTACIONES (cCtaCod, nSaldoDisp, nPersoneria, nFirmas, nIntAcum, dUltCierre, dApertura, nPrdCtaTpo, nPrdTasaInteres, " _
                                & "cUltimaActualizacion) " _
            & "VALUES ('" & sCuenta & "'," & nSaldo & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," _
                        & nTipoCuenta & "," & nTipoTasa & ",'" & sMovNro & "')"

    oConn.Ejecutar lsSQL
        
    lsSQL = "INSERT CAPTACAHORROS (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto) " _
        & "VALUES ('" & sCuenta & "',0," & "0" & ",0,'" & sFecha & "')"

    oConn.Ejecutar lsSQL

AgregaNuevaCaptacion = sCuenta

End Function

'******************************************************
' CAFF - 16/10/2002 - CAPTACIONES - APERTURA DE CTA
'******************************************************
Public Sub AgregaNuevoProdPers(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)

lsSQL = "INSERT PRODUCTOPERSONA (cCtaCod,cPersCod,nPrdPersRelac) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ")"
    
oConn.Ejecutar lsSQL

End Sub

'******************************************************
' CAFF - 18/10/2002 - CAPTACIONES - CANCELACION DE CTA
'******************************************************
Public Sub ActualizaCargoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

If bActExtracto Then
    lsSQL = "UPDATE PRODUCTO SET nSaldo = nSaldo - " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    lsSQL = "UPDATE PRODUCTO SET nSaldo = nSaldo - " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If

oConn.Ejecutar lsSQL

    lsSQL = "UPDATE CAPTACIONES SET nSaldoDisp = nSaldoDisp - " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
        & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
        & "WHERE cCtaCod = '" & sCuenta & "'"

oConn.Ejecutar lsSQL

End Sub

'******************************************************
' CAFF - 18/10/2002 - CAPTACIONES - CANCELACION DE CTA
'******************************************************
Public Sub UltimaActualizacionCuenta(ByVal sCuenta As String, ByVal sMovNro As String)

'Actualiza la fecha de ultima actualizacion de la cuenta de captaciones
    lsSQL = "Update Captaciones Set cUltimaActualizacion = '" & sMovNro & "' " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
        
    oConn.Ejecutar lsSQL
    
End Sub

Public Function GetSaldoCancelacion(ByVal sCuenta As String, ByVal dFecSis As Date, ByVal sCodAge As String, _
            ByRef pnIntGanado As Currency, ByRef pnSaldo As Currency) As Double
Dim nProd As Producto
Dim rsProd As Recordset
Dim nSaldoDisp As Double, nSaldoCnt As Double
Dim nIntAcum As Double, nIntGanado As Double
Dim nTasa As Double
Dim dUltMov As Date
Dim nDiasTranscurridos As Long

    Set rsProd = GetDatosCuentaAho(sCuenta)
    
    nSaldoDisp = rsProd("nSaldoDisp")
    pnSaldo = nSaldoDisp
    nSaldoCnt = rsProd("nSaldo")
    dUltMov = rsProd("dUltCierre")
    nTasa = rsProd("nTasaInteres")
    nIntAcum = rsProd("nIntAcum")
    nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
    nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
    pnIntGanado = nIntAcum + nIntGanado
    GetSaldoCancelacion = nSaldoDisp + nIntAcum + nIntGanado
    Set rsProd = Nothing

End Function

Public Function GetMovBloqueo(psCtaCod As String) As String
Dim Rs As Recordset

lsSQL = "SELECT cMovNro FROM PRODUCTOBLOQUEOS " _
           & " WHERE cMovNroDbl Is Null And nBlqMotivo = 12 " _
                    & " AND cCtaCod = '" & psCtaCod & "'"
                                     
Set Rs = oConn.CargaRecordSet(lsSQL)

If Rs.EOF And Rs.BOF Then
    GetMovBloqueo = ""
Else
    GetMovBloqueo = Rs!cMovNro
End If

Set Rs = Nothing

End Function

Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    On Error GoTo GeneraMovNroErr
    Dim Rs As ADODB.Recordset
    Dim oConect As DConecta
    Dim sql As String
    Set oConect = New DConecta
    Set Rs = New ADODB.Recordset
    If oConect.AbreConexion = False Then Exit Function
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & oConect.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set Rs = oConect.Ejecutar(sql)
    If Not Rs.EOF Then
        GeneraMovNro = Rs.Fields(0)
    End If
    Rs.Close
    Set Rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
GeneraMovNroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function

Public Function GetNumRemate(Optional ByVal pnNumRemate As Integer) As Recordset

    lsSQL = "SELECT TOP 1 nRemate AS NumRemate, dInicio, dFin, dReferencia, nTipoProceso, dAviso, cUbicacion, " _
        & "ISNULL(nAdjudicado, 0) nAdjudicado, ISNULL(nSobrante,0) nSobrante, CR.cPerscod, cPersNombre,  " _
        & "Ubicacion = (SELECT cConsDescripcion FROM CONSTANTE C WHERE C.nConsValor = CR.cUbicacion AND C.nConsCod = 8004) " _
        & "FROM COLOCPIGREMATE CR LEFT JOIN PERSONA P ON CR.cPersCod = P.cPersCod  "
          
        If pnNumRemate > 0 Then
            lsSQL = lsSQL & "WHERE nRemate = " & pnNumRemate
        End If
          
    lsSQL = lsSQL & "ORDER BY nRemate Desc "

    Set GetNumRemate = oConn.CargaRecordSet(lsSQL)
    
End Function

Public Function GetClientesAvisoRemate(ByVal pnDiasVenc As Integer, ByVal pdFecIniRem As Date) As Recordset


    lsSQL = "SELECT Distinct (PP.cPerscod), cPersNombre, cPersDireccDomicilio, cUbiGeoDescripcion, " _
        & "CodPostal = (SELECT cDesCodPostal + ' (' + cUbiGeoDescripcion + ')' " _
        & "FROM UBICACIONGEOGRAFICA UG1 INNER JOIN CODIGOPOSTAL CP ON CP.cCodZon = UG1.CUBIGEOCOD " _
        & "WHERE SUBSTRING(UG1.CUBIGEOCOD ,2,6) = SUBSTRING(H.CPERSDIRECCUBIGEO,2,6) AND SUBSTRING(UG1.CUBIGEOCOD,1,1) = '3') " _
        & "FROM PRODUCTOPERSONA PP INNER JOIN COLOCPIGNO CP ON PP.CCTACOD = CP.CCTACOD " _
        & "INNER JOIN COLOCACIONES C ON CP.CCTACOD = C.CCTACOD INNER JOIN PRODUCTO P ON P.cCtaCod = CP.cCtaCod " _
        & "INNER JOIN PERSONA H ON PP.CPERSCOD = H.CPERSCOD LEFT JOIN UBICACIONGEOGRAFICA UG " _
        & "ON H.CPERSDIRECCUBIGEO = UG.CUBIGEOCOD " _
        & "WHERE DATEDIFF(dd, dVenc, '" & Format(pdFecIniRem, "mm/dd/yyyy") & "') >= " & pnDiasVenc _
        & " AND nPrdEstado IN (" & gPigEstDesemb & "," & gPigEstAmortiz & "," & gPigEstReusoLin & ")"
        
        
    Set GetClientesAvisoRemate = oConn.CargaRecordSet(lsSQL)
    
End Function

Public Function GetContratosAvisoRemate(ByVal psCodPers As String, ByVal psFecFin As String) As Recordset

'    lsSQL = "SELECT C.cCtaCod, dPrdEstado, nSaldo, TI1.nTasaInteres IntComp, TI2.nTasaInteres IntMor, dVenc " _
'        & "FROM PRODUCTOPERSONA PP INNER JOIN PRODUCTO P ON PP.cCtaCod = P.cCtaCod " _
'        & "INNER JOIN COLOCACIONES C ON C.cCtaCod = PP.cCtaCod " _
'        & "INNER JOIN PRODUCTOTASAINTERES TI1 ON P.cCtaCod = TI1.cCtaCod " _
'        & "INNER JOIN PRODUCTOTASAINTERES TI2 ON P.cCtaCod = TI2.cCtaCod " _
'        & "Where TI1.nPrdTasaInteres = 200 And TI2.nPrdTasaInteres = 210 " _
'        & "AND cPersCod = '" & psCodPers & "' ORDER BY C.cCtaCod"

    lsSQL = "SELECT C.cCtaCod, dPrdEstado, nSaldo, TI1.nTasaInteres IntComp, TI2.nTasaInteres IntMor, dVenc," _
        & "dbo.ColocPignoInteres(TI1.nTasaInteres, datediff(dd, dPrdEstado, '" _
        & Format(psFecFin, "mm/dd/yyyy") & "'), nSaldo) Interes, " _
        & "dbo.ColocPignoInteres(TI2.nTasaInteres, datediff(dd, dVenc, '" _
        & Format(psFecFin, "mm/dd/yyyy") & "'), nSaldo) Mora " _
        & "FROM PRODUCTOPERSONA PP INNER JOIN PRODUCTO P ON PP.cCtaCod = P.cCtaCod " _
        & "INNER JOIN COLOCACIONES C ON C.cCtaCod = PP.cCtaCod " _
        & "INNER JOIN PRODUCTOTASAINTERES TI1 ON P.cCtaCod = TI1.cCtaCod " _
        & "INNER JOIN PRODUCTOTASAINTERES TI2 ON P.cCtaCod = TI2.cCtaCod " _
        & "WHERE TI1.nPrdTasaInteres = 200 And TI2.nPrdTasaInteres = 210 " _
        & "AND cPersCod = '" & psCodPers & "' ORDER BY C.cCtaCod"
   
    Set GetContratosAvisoRemate = oConn.CargaRecordSet(lsSQL)

End Function

' MODIFICADO CRSF - 17/06
Public Function GetContratosRemate(ByVal pnNumRemate As Integer) As Recordset

    On Error GoTo dError

    lsSQL = "SELECT Distinct CR.cCtaCod, nPiezas, CP.nUbicaLote, nPrdEstado, cConsDescripcion Estado, nDiasAtraso, " _
        & "UbicaLote = (SELECT cConsDescripcion FROM CONSTANTE C WHERE C.nConsValor = CP.nUbicaLote AND nConsCod = 8004 ) " _
        & "FROM COLOCPIGPROCESO CR INNER JOIN COLOCPIGNO CP ON CR.cCtaCod = CP.cCtaCod " _
        & "INNER JOIN PRODUCTO P ON P.cCtaCod = CR.cCtaCod " _
        & "INNER JOIN CONSTANTE C ON C.nConsValor = P.nPrdEstado AND nConsCod = 3001 " _
        & "WHERE nRemate = " & pnNumRemate & " AND nSituacionPieza = " & gPigSituacionDisponible
        
    Set GetContratosRemate = oConn.CargaRecordSet(lsSQL)
    
    Exit Function
    
dError:
    Err.Raise Err.Number, "Obtiene Contratos para Remate <<GetContratosRemate>>", Err.Description

End Function


Public Function dObtieneColocPigProcesoDet(ByVal psCodCta As String) As Recordset

    On Error GoTo dError
    
    lsSQL = "SELECT nCodConcepto, nMonto FROM COLOCPIGPROCESODET CRD " _
        & "INNER JOIN COLOCPIGPROCESO CR ON CRD.nRemate = CR.nRemate AND CR.nTipoProceso = CRD.nTipoProceso " _
        & "AND CRD.cCtaCod = CR.cCtaCod " _
        & "WHERE nSituacionPieza = " & gPigSituacionPendFacturar & " AND CR.cCtaCod = '" & psCodCta & "'"
    
    Set dObtieneColocPigProcesoDet = oConn.CargaRecordSet(lsSQL)
    
    Exit Function
        
dError:
    Err.Raise Err.Number, "Obtiene detalle de Remate <<dObtieneColocPigProcesoDet>>", Err.Description
    
End Function

Public Sub dUpdateColocPigRemate(ByVal pnNumRemate)

On Error GoTo dError

    lsSQL = "UPDATE COLOCPIGREMATE SET nSobrante = 1 WHERE nRemate = " & pnNumRemate
    oConn.Ejecutar lsSQL
    
    Exit Sub
    
dError:
    Err.Raise Err.Number, "Error en Actualiza ColocPigRemate", Err.Description
    
End Sub

Public Function CapExtornoCancelacion(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal nMonto As Double) As Long

Dim nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim dFecSis As Date
Dim bTrans As Boolean
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nExtracto As Long
Dim sGlosa As String

    dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
    sCodUser = Right(sMovNro, 4)
    bTrans = False

    'Obtiene los datos para el calculo
    Set rsCta = GetDatosCuentaAho(sCuenta)
    nSaldoDisp = rsCta("nSaldoDisp")
    nSaldoCnt = rsCta("nSaldo")
    nEstado = rsCta("nPrdEstado")
    nIntAcum = rsCta("nIntAcum")
    nExtracto = rsCta("nTransacc")
    rsCta.Close
    Set rsCta = Nothing

    On Error GoTo ErrGraba

    sGlosa = "Extorno Cancelacion Cuenta Por Sobrante = " & sCuenta

    '=========== DESBLOQUEO DE LA CUENTA ==================
    ActualizaBloqueoTot sCuenta, "Bloqueo Total de Cuenta", sMovNro, gCapMotBlqTotSobRemate, ""
    ActualizaEstadoCuenta sCuenta, gCapEstBloqTotal, dFecSis

    dUpdateMov nMovNroBus, , , , gMovFlagExtornado
    AgregaMov sMovNro, nOperacion, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno

    nMovNro = GetnMovNro(sMovNro)
    AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
    UltimaActualizacionCuenta sCuenta, sMovNro

    CapExtornoCancelacion = nMovNro
    
    Exit Function

ErrGraba:
    Err.Raise Err.Number, "CapExtornoCancelacionAho", Err.Description
    CapExtornoCancelacion = 0
End Function

Public Sub dUpdateMov(ByVal pnMovNro As Long, Optional ByVal psOpeCod As String = "@", _
    Optional ByVal psMovDesc As String = "@", Optional ByVal pnMovEstado As MovEstado = -1, _
    Optional ByVal pnMovFlag As MovFlag = -1, Optional pbEjecBatch As Boolean = False)

    lsSQL = "UPDATE Mov SET "

    If psOpeCod <> "@" Then
        lsSQL = lsSQL & " cOpeCod = '" & psOpeCod & "',"
    End If
    If psMovDesc <> "@" Then
        lsSQL = lsSQL & " cMovDesc = '" & psMovDesc & "',"
    End If
    If pnMovEstado <> -1 Then
        lsSQL = lsSQL & " nMovEstado = " & pnMovEstado & ","
    End If
    If pnMovFlag <> -1 Then
        lsSQL = lsSQL & " nMovFlag = " & pnMovFlag & ","
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE nMovNro =" & pnMovNro & " "

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

End Sub

Public Function GetMovRef(ByVal pnMovNroRef As Long) As Long
Dim Rs As Recordset

    lsSQL = "SELECT nMovNro FROM MOVREF WHERE nMovNroRef = " & pnMovNroRef
    Set Rs = oConn.Ejecutar(lsSQL)

    If Not (Rs.EOF And Rs.BOF) Then
         
        GetMovRef = Rs!nMovNro
    Else
        GetMovRef = 0
    End If

    Set Rs = Nothing
    
End Function

Public Function GetMovCol(ByVal pnMovNro As Long) As Recordset

    lsSQL = "SELECT cCtaCod FROM MOVCOL WHERE nMovNro = " & pnMovNro
    Set GetMovCol = oConn.CargaRecordSet(lsSQL)

End Function

Public Function GetMovCap(ByVal pnMovNro As Long) As Recordset

    lsSQL = "SELECT cCtaCod, nMonto FROM MOVCAP WHERE NMOVNRO = " & pnMovNro
    Set GetMovCap = oConn.CargaRecordSet(lsSQL)

End Function

'NUEVO CRSF - 17/06 Ubicacion Actual del Remate
Public Function GetUbicaLoteRemate(ByVal pnRemate) As Integer
Dim Rs As Recordset

    lsSQL = "select cUbicacion  from colocpigremate where nRemate = " & pnRemate
    Set Rs = oConn.Ejecutar(lsSQL)

    If Not (Rs.EOF And Rs.BOF) Then
        GetUbicaLoteRemate = Rs!cUbicacion
    End If

    Set Rs = Nothing
    
End Function
