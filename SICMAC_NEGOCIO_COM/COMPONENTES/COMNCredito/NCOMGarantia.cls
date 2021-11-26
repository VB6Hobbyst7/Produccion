VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private lsNegritaOn As String
Private lsNegritaOff  As String
Dim oImpre As COMFunciones.FCOMImpresion

Public Sub LiberaGarantia(ByVal psNumGarant As String, ByVal pnTipoGarant As Integer, _
    ByVal pnMontoLiberado As Double, ByVal pnMoneda As Integer, ByVal psAge As String, ByVal psUser As String, _
    ByVal pdFecha As Date, ByVal psCodCta As String, ByVal pnEstado As Integer)

Dim oBase As COMDCredito.DCOMCredActBD
Dim nMovNro As Long
Dim sMovNro As String
Dim sOpeCod As String
Dim R As ADODB.Recordset

    If pnMoneda = 1 Then
        Select Case pnTipoGarant
            Case 1
                sOpeCod = "107301"
            Case 2
                sOpeCod = "107302"
            Case 3
                sOpeCod = "107303"
            Case 4
                sOpeCod = "107304"
            Case 5
                sOpeCod = "107305"
            Case 6
                sOpeCod = "107306"
            Case 7
                sOpeCod = "107307"
            Case 8
                sOpeCod = "107308"
            Case 9
                sOpeCod = "107309"
            Case 10
                sOpeCod = "107310"
            Case 11
                sOpeCod = "107311"
            Case 12
                sOpeCod = "107312"
        End Select
    Else
        Select Case pnTipoGarant
            Case 1
                sOpeCod = "107401"
            Case 2
                sOpeCod = "107402"
            Case 3
                sOpeCod = "107403"
            Case 4
                sOpeCod = "107404"
            Case 5
                sOpeCod = "107405"
            Case 6
                sOpeCod = "107406"
            Case 7
                sOpeCod = "107407"
            Case 8
                sOpeCod = "107408"
            Case 9
                sOpeCod = "107409"
            Case 10
                sOpeCod = "107410"
            Case 11
                sOpeCod = "107411"
            Case 12
                sOpeCod = "107412"
        End Select
    End If
    Set oBase = New COMDCredito.DCOMCredActBD
    oBase.dBeginTrans
    '20090928 - MADM el proceso es automatico al finalizar pago credito, gravament 0 y estado es 3
    'Call oBase.dUpdateGarantiaLibera(psNumGarant, pnEstado, pnMontoLiberado)
    Call oBase.dUpdateColocGarantia(psNumGarant, psCodCta, 0)
    
    If pnMontoLiberado > 0 Then
        sMovNro = oBase.GeneraMovNro(Format(pdFecha, "dd/mm/yyyy"), psAge, psUser)
        Call oBase.dInsertMov(sMovNro, sOpeCod, "Liberacion de Garantia Manual", gMovEstContabMovContable, gMovFlagVigente, False)
        nMovNro = oBase.GetnMovNro(sMovNro)
        Call oBase.dInsertMovCol(nMovNro, sOpeCod, psCodCta, 0, pnMontoLiberado, 0, "", 0, 0, 0)
        Call oBase.dInsertMovColDet(nMovNro, sOpeCod, psCodCta, 0, gColocConceptoCodCapital, 0, pnMontoLiberado)
    End If
    oBase.dCommitTrans
    
    Dim odGar As COMDCredito.DCOMGarantia
    Dim sSql As String
    Set odGar = New COMDCredito.DCOMGarantia
    Set R = odGar.RecuperaGarantia(psNumGarant)
    Set odGar = Nothing
    
    If R!nGravament = 0 Then 'Liberara
        sSql = " UPDATE Garantias Set nEstado = 4 Where cNumGarant = '" & psNumGarant & "'"
    Else 'Contabilizada
        sSql = " UPDATE Garantias Set nEstado = 3 Where cNumGarant = '" & psNumGarant & "'"
    End If
    Call oBase.coConex.Ejecutar(sSql)
    
    Set oBase = Nothing
End Sub


Public Sub BloqueaGarantia(ByVal psNumGarant As String)
Dim oBase As COMDCredito.DCOMCredActBD
    
    Set oBase = New COMDCredito.DCOMCredActBD
    Call oBase.dUpdateGarantiaBloquea(psNumGarant, gPersGarantEstadoBloqueada)
    Set oBase = Nothing
End Sub

Public Sub DesbloqueaGarantia(ByVal psNumGarant As String)
Dim oBase As COMDCredito.DCOMCredActBD
Dim odGar As COMDCredito.DCOMGarantia
Dim bGarVig As Boolean

    Set odGar = New COMDCredito.DCOMGarantia
    bGarVig = odGar.PerteneceACreditoVigente(psNumGarant)
    Set oBase = New COMDCredito.DCOMCredActBD
    If bGarVig Then
        Call oBase.dUpdateGarantiaBloquea(psNumGarant, gPersGarantEstadoContabilizado)
    Else
        bGarVig = odGar.PerteneceACredito(psNumGarant)
        If Not bGarVig Then
            Call oBase.dUpdateGarantiaBloquea(psNumGarant, gPersGarantEstadoAsignado)
        Else
            Call oBase.dUpdateGarantiaBloquea(psNumGarant, gPersGarantEstadoLiberado)
        End If
    End If
    
    Set odGar = Nothing

End Sub

Public Function PorcentajeGarantia(ByVal psConsGarantia As String) As Double
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset

    On Error GoTo ErrorPorcentajeGarantia
    sSql = "Select nParamValor from ColocParametro Where nParamVar = " & psConsGarantia
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        PorcentajeGarantia = R!nParamValor / 100
    Else
        PorcentajeGarantia = 0
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    R.Close
    Set R = Nothing
    Exit Function

ErrorPorcentajeGarantia:
    Set oConecta = Nothing
    Set R = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ValidaDatos(ByVal RelPers As Variant, ByVal pnMontoTasac As Double, _
    ByVal pnMontoReali As Double, ByVal pnMontoDisp As Double, _
    Optional ByVal pbSoloGaran As Boolean = False, _
    Optional ByVal pnTipoGarant As Long = -1) As String
Dim i As Integer
Dim nNumTit As Integer
Dim nPorc As Double

     ValidaDatos = ""
    'Valida Numero de Titulares
    If Not pbSoloGaran Then
        nNumTit = 0
        For i = 0 To UBound(RelPers) - 1
            If CInt(Trim(Right(RelPers(i), 10))) = gPersRelGarantiaTitular Then
                nNumTit = nNumTit + 1
            End If
        Next i
        If nNumTit > 1 Then
            ValidaDatos = "No puede haber mas de un Titular"
        End If
    End If
    'Valida Montos Tasacion, Realizacion, Disponible
    If pnMontoReali > pnMontoTasac Then
        ValidaDatos = "El Monto de Realizacion No Puede Ser Mayor que el Monto de Tasacion"
        Exit Function
    End If
    
    If pnMontoDisp > pnMontoReali Then
        ValidaDatos = "El Monto Disponible No Puede Ser Mayor que el Monto de Realizacion"
        Exit Function
    End If
            
End Function
Public Function RecuperaTitularGarantReal(ByVal Nrog As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecupera
     
sSql = "Select PG.cPersCod as CodPers, P.cPersNombre as Nombre From GarantReal GR"
sSql = sSql & " Inner Join PersGarantia PG on PG.cNumGarant=GR.cNumGarant"
sSql = sSql & " Inner Join Persona P on P.cPersCod=Pg.cPersCod"
sSql = sSql & " Where GR.cNumGarant =" & Nrog
sSql = sSql & " And PG.nRelacion = 1"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaTitularGarantReal = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecupera:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
Public Function ImprimeCabeceraGarantiaRealExtorno(ByVal psNomAge As String, _
        ByVal psFecha As String, ByVal psHora As String, _
        ByVal psPersNombre As String) As String

Dim lscadena As String
Dim oVar As COMFunciones.FCOMCadenas
'lsCadena = Chr$(10)
Set oVar = New COMFunciones.FCOMCadenas
lscadena = Chr$(10) & Chr(10) & lsNegritaOn
lscadena = lscadena & oImpre.ImpreFormat("CMACT - CREDITOS ", 43, , False) & Space(6) & oImpre.ImpreFormat("CMACT - CREDITOS ", 43, , False) & Chr(10)
lscadena = lscadena & oImpre.ImpreFormat(psNomAge, 43, , False) & Space(6) & oImpre.ImpreFormat(psNomAge, 43, , False) & Chr(10) & lsNegritaOff
lscadena = lscadena & oImpre.ImpreFormat("Fecha:" & psFecha & Space(5) & "Hora:" & psHora, 43, , False) & Space(6) & oImpre.ImpreFormat("Fecha:" & psFecha & Space(5) & "Hora:" & psHora, 43, , False) & Chr(10)
lscadena = lscadena & lsNegritaOn & oImpre.ImpreFormat(oVar.PstaNombre(psPersNombre), 43, , False) & Space(6) & oImpre.ImpreFormat(oVar.PstaNombre(psPersNombre), 43, , False) & Chr(10)
lscadena = lscadena & oImpre.ImpreFormat("----------EXTORNO DE LEVANTAMIENTO----------", 43, , False) & Space(6) & oImpre.ImpreFormat("----------EXTORNO DE LEVANTAMIENTO----------", 43, , False) & Chr(10) & lsNegritaOff
Set oVar = Nothing
ImprimeCabeceraGarantiaRealExtorno = lscadena
End Function

Public Function ImprimeBoletaGarantiaExt(ByVal psNomAge As String, _
        ByVal psFecha As String, ByVal psHora As String, _
        ByVal Nrog As String, ByVal psCodUser As String, ByVal pnMovG As String) As String
Dim Cabecera As String
Dim cuerpo As String
Dim loConecta As COMConecta.DCOMConecta
Dim Dg As COMDCredito.DCOMGarantia
Dim rs As ADODB.Recordset
Dim sSql As String
Dim Total As Double
Dim Lineas As Integer
Dim llena As String
Dim i As Integer
Dim MovG As Long
Lineas = 7
Set Dg = New COMDCredito.DCOMGarantia
Set rs = New ADODB.Recordset
Set rs = RecuperaTitularGarantReal(Nrog)

Cabecera = ImprimeCabeceraGarantiaRealExtorno(psNomAge, psFecha, psHora, rs!Nombre)
Total = 0

Set loConecta = New COMConecta.DCOMConecta
sSql = "select nMovNro, cOpeCod from Mov Where cMovNro='" & pnMovG & "'"
        loConecta.AbreConexion
Set rs = loConecta.CargaRecordSet(sSql)
MovG = IIf(IsNull(rs!nMovNro), 0, rs!nMovNro)

'Detalle
sSql = " Select( Select cDescripcion from ProductoConcepto where nPrdConceptoCod=MGD.nPrdConceptoCod) Concepto,"
sSql = sSql & " nMonto"
sSql = sSql & " from MovGarantDet MGD"
sSql = sSql & " Where MGD.nMovNro =" & MovG

Set rs = loConecta.CargaRecordSet(sSql)
loConecta.CierraConexion
  
cuerpo = Cabecera
If rs.BOF And rs.EOF Then

Else

    cuerpo = cuerpo & Chr(10)
    While Not rs.EOF
    cuerpo = cuerpo & oImpre.ImpreFormat(rs!Concepto, 23, 0) & Space(8) & oImpre.ImpreFormat(rs!nMonto, 10, 2, True) & Space(6) & oImpre.ImpreFormat(rs!Concepto, 23, 0, False) & Space(10) & oImpre.ImpreFormat(rs!nMonto, 10, 2, True) & Chr(10)
    Lineas = Lineas + 1
    Total = Total + rs!nMonto
    rs.MoveNext
    Wend
    
End If
cuerpo = cuerpo & lsNegritaOn & String(45, "-") & Space(6) & String(45, "-") & Chr(10)
cuerpo = cuerpo & oImpre.ImpreFormat("Total Gastos", 23, 0) & Space(8) & oImpre.ImpreFormat(Total, 10, 2, True) & Space(8) & oImpre.ImpreFormat("Total Gastos", 23, 0, False) & Space(8) & oImpre.ImpreFormat(Total, 10, 2, True) & Chr(10)
cuerpo = cuerpo & psCodUser & Space(48) & psCodUser & lsNegritaOff
Lineas = Lineas + 3

If Lineas <= 18 Then
    For i = Lineas To 18
        cuerpo = cuerpo + Chr(10)
        
    Next i
End If
Set rs = Nothing
Set loConecta = Nothing
Set Dg = Nothing
ImprimeBoletaGarantiaExt = cuerpo
End Function

Private Sub Class_Initialize()
 lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)
Set oImpre = New COMFunciones.FCOMImpresion
End Sub

Public Function ValidaDatosGarantia(ByVal psCtaCod As String, _
                                    ByRef pbValidaCuenta As Boolean, _
                                    ByVal psConsGarantia As String, _
                                    ByRef pnValorCuota As Double, _
                                    ByRef pnValorPorcen As Double)
Dim sSql As String
Dim oGarant As COMDCredito.DCOMGarantia

    On Error GoTo ErrorValidaDatosGarantia
     
    Set oGarant = New COMDCredito.DCOMGarantia
    
    'If oGarant.ValidaPFCTS(psCtaCod) = False Then
   '     Set oGarant = Nothing
        pbValidaCuenta = oGarant.ValidaPFCTS(psCtaCod)  'False
    '    Exit Function
    'End If
    Set oGarant = Nothing
    
    pnValorPorcen = PorcentajeGarantia(gPersGarantia & psConsGarantia)
    pnValorCuota = PorcentajeGarantia("3052")
    
    Exit Function

ErrorValidaDatosGarantia:
    Err.Raise Err.Number, "Valida Datos Garantia", Err.Description
End Function


'Grabar Datos
Public Function GrabarDatos(ByVal pnTipoOperacion As Integer, _
                            ByRef psValidar As String, _
                            ByRef pbExisteGar As Boolean, _
                           ByVal pnMonto As Double, _
                            ByVal pnMontoDisponible As Double, _
                           ByVal psCtaCod As String, _
                           ByVal psNumGarant As String, _
                           ByVal psMoneda As String, _
                           ByVal pnGravado As Double, _
                           ByVal pnGravadoAnt As Double)
    
    
    Dim oNegCredito As COMNCredito.NCOMCredito
    Dim oCredito As COMDCredito.DCOMCredito
    
    On Error GoTo ErrorGrabarDatos
    
    
    Set oNegCredito = New COMNCredito.NCOMCredito
    
    psValidar = oNegCredito.ValidaDatosGarantiaCredito(pnMonto, pnMontoDisponible)
    If psValidar <> "" Then
        Set oNegCredito = Nothing
        Exit Function
    End If
    
    Set oNegCredito = Nothing
    Set oCredito = New COMDCredito.DCOMCredito
    If pnTipoOperacion = 1 Then
        pbExisteGar = oCredito.ExisteGarantia(psCtaCod, psNumGarant)
        If Not pbExisteGar Then
            Call oCredito.NuevaGarantia(psCtaCod, psNumGarant, psMoneda, pnGravado)
        Else
            'psValidar = "Garantia Ya Existe"
            Set oCredito = Nothing
            Exit Function
        End If
    Else
        Call oCredito.ActualizaGarantias(psCtaCod, psNumGarant, psMoneda, pnGravado, pnGravadoAnt)
    End If
    Set oCredito = Nothing
    
    Exit Function

ErrorGrabarDatos:
    Err.Raise Err.Number, "Grabar Datos", Err.Description
End Function

'****************************************************
'Funciones para el Reporte de la Declaracion Jurada

Public Function CargarPersonaDJ(ByVal pcNumGarant As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
    sSql = "Select P.cPersCod,P.cPersNombre,PID.cPersIDNro,P.cPersDireccDomicilio"
    sSql = sSql & " From PersGarantia PG"
    sSql = sSql & " Inner Join Persona P on PG.cPersCod=P.cPersCod"
    sSql = sSql & " Inner Join PersId PID on PID.cPersCod=P.cPersCod"
    sSql = sSql & " Inner Join Constante CN on CN.nConsValor=PG.nRelacion and CN.nConsValor<>1021"
    sSql = sSql & " Where PG.cNumGarant='" & pcNumGarant & "' and CN.nConsValor=1 and CN.nConsCod=1021"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set CargarPersonaDJ = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
End Function

Public Function CargarDJ(ByVal psNumGarant As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim nPos As Integer
    Dim nSMonto As Double
    
    sSql = "Select cGarDjDescripcion,nGarDJCantidad*nGarDJPrecioUnit as nMonto"
    sSql = sSql & " From GarantDeclaracionJur GJ"
    sSql = sSql & " Inner Join Garantias G on G.cNumGarant=GJ.cNUmGarant"
    sSql = sSql & " Where GJ.cNumGarant='" & psNumGarant & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set CargarDJ = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function TitularFiador(ByVal psNumGarant As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
    sSql = "Select nTpoGarantia From Garantias Where cNumGarant='" & psNumGarant & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set TitularFiador = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
End Function

'****************************************************
Private Sub Class_Terminate()
    Set oImpre = Nothing
End Sub
Public Function InsertarGarantiaSaneamiento(ByVal psNumGarant As String, ByVal psCtaCod As String, ByVal pnTipoSan As Integer, ByVal pnPeriSan As Integer, ByVal pnMontSan As Double, ByVal pdFecSan As Date, ByVal psUsuariAdju As String, ByVal pnContador As Integer, ByVal pnEstadoAdj As Integer, ByVal dfechaAdj As Date, ByVal pnEstado As Integer, ByVal pdHoy As Date, ByVal psUsuarioGaran As String, ByVal psCodAge As String, ByVal pnMoneda As Integer, ByVal nTESan As Integer)
Dim sSql As String
Dim oGarant As COMDCredito.DCOMGarantia
'Dim oBase As COMDCredito.DCOMCredActBD
Dim pbTran As Boolean

    On Error GoTo ErrorInsertarGarantiaSaneamiento
    
    'Genera cMovNro
    Dim psMovAct As String
    Dim loContFunct As COMNContabilidad.NCOMContFunciones
    
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
    psMovAct = loContFunct.GeneraMovNro(pdHoy, psCodAge, psUsuariAdju) ' gsCodUser)
    Set loContFunct = Nothing
    'Fin de Generacion de cMovNro
     
    pbTran = True
    'oBase.dBeginTrans
    Set oGarant = New COMDCredito.DCOMGarantia
        Call oGarant.InsertarGarantiaSaneamiento(psNumGarant, psCtaCod, pnTipoSan, pnPeriSan, pnMontSan, pdFecSan, psMovAct, pnContador, pnMoneda, nTESan)
    Set oGarant = Nothing
    If pnContador = 0 Then
        Set oGarant = New COMDCredito.DCOMGarantia
        Call oGarant.dUpdateGarantiasAdjudicados(psNumGarant, pdHoy, pnEstado, psMovAct) ', psUsuariAdju)
        'Call oBase.dCommitTrans
        Set oGarant = Nothing
    End If
    
    
    Exit Function

ErrorInsertarGarantiaSaneamiento:
    'If pbTran Then
     '   Call oBase.dRollbackTrans
    'End If
    Err.Raise Err.Number, "Valida Datos Garantia", Err.Description
End Function
Public Function InsertarGarantiaRemate(ByVal psNumGarant As String, ByVal psCtaCod As String, ByVal nNumRemate As Integer, ByVal cLugarRemate As String, ByVal dFechaRemate As Date, ByVal dFechaSis As Date, ByVal nEstadoAdju As Integer, ByVal cPersCod As String, ByVal nMonto As Double, ByVal nContador As Integer, ByVal psUsuario As String, ByVal pbVendido As Boolean, ByVal psCodAge As String, ByVal pnMoneda As Integer, ByVal psPerCod2 As String, ByVal nInteres As Double)
Dim sSql As String
Dim oGarant As COMDCredito.DCOMGarantia
'Dim oBase As COMDCredito.DCOMCredActBD
Dim nEstadoA As Integer
Dim pbTran As Boolean

    On Error GoTo ErrorInsertarGarantiaRemate
    
    '
    'Genera cMovNro
    Dim psMovAct As String
    Dim loContFunct As COMNContabilidad.NCOMContFunciones
    
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
    psMovAct = loContFunct.GeneraMovNro(dFechaSis, psCodAge, psUsuario) ' gsCodUser)
    Set loContFunct = Nothing
    'Fin de Generacion de cMovNro
     
    pbTran = True
    'oBase.dBeginTrans***
    Set oGarant = New COMDCredito.DCOMGarantia
        Call oGarant.InsertarGarantiaRemate(psNumGarant, psCtaCod, nNumRemate, cLugarRemate, dFechaRemate, dFechaSis, nEstadoAdju, cPersCod, nMonto, nContador, psMovAct, pnMoneda, psPerCod2, nInteres)
    Set oGarant = Nothing
    If (nContador = 2 And Not pbVendido) Or (pbVendido And nContador = 0) Then
        'Set oBase = New COMDCredito.DCOMCredActBD
        Set oGarant = New COMDCredito.DCOMGarantia
        
        If Not pbVendido Then
            If nContador = 2 Then
                nEstadoA = gPersGarantEstadoAdjudicado
            End If
        
        Else
            If nContador = 0 Then
                nEstadoA = gPersGarantEstadoRematado
            End If
        End If
        Call oGarant.dUpdateGarantiasAdjudicados(psNumGarant, dFechaSis, nEstadoA, psMovAct) ' psUsuario)
        'Call oBase.dCommitTrans
        Set oGarant = Nothing
    End If
    
    
    Exit Function

ErrorInsertarGarantiaRemate:
    'If pbTran Then
     '   Call oBase.dRollbackTrans
    'End If
    Err.Raise Err.Number, "Valida Datos Garantia", Err.Description
End Function
'ALPA***********************20080805***********************************
Public Sub TransferirGarantiasARecuperacionesLote(ByVal prs As ADODB.Recordset, _
                                        ByVal psCodUser As String, ByVal psCodAge As String, Optional ByRef psMensaje As String)

On Error GoTo ErrorTransferirGarantiasARecuperacionesLote

    prs.MoveFirst
    While Not prs.EOF
        Call TransferirGarantiasARecuperaciones(prs("cGarantia"), prs("dEstadoAdju"), _
                  prs("nEstadoAdju"), psCodUser, psCodAge, psMensaje)

        prs.MoveNext
    Wend

    Exit Sub

ErrorTransferirGarantiasARecuperacionesLote:
    Err.Raise Err.Number, "Transferir a Recuperaciones en Lote", Err.Description
End Sub
'**********************************************************************
'ALPA***********************20080805***********************************
Public Sub TransferirGarantiasARecuperaciones(ByVal psGaranCod As String, ByVal pdHoy As Date, _
    ByVal pnEstado As Integer, ByVal psUser As String, ByVal psAgencia As String, Optional ByVal psMensaje As String = "")
Dim oBase As COMDCredito.DCOMGarantia
Dim pbTran As Boolean
Dim oGen As COMFunciones.FCOMFechas
Dim oConecta As COMConecta.DCOMConecta

Dim sSql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    On Error GoTo ErrorTransferirGarantiasARecuperaciones

    pbTran = False

    Set oBase = New COMDCredito.DCOMGarantia

    'Genera cMovNro
    Dim psMovAct As String
    Dim loContFunct As COMNContabilidad.NCOMContFunciones

    Set loContFunct = New COMNContabilidad.NCOMContFunciones
    psMovAct = loContFunct.GeneraMovNro(pdHoy, psAgencia, psUser) ' gsCodUser)
    Set loContFunct = Nothing
    'Fin de Generacion de cMovNro

    pbTran = True
    oBase.dBeginTrans
    Call oBase.dUpdateGarantiasAdjudicados(psGaranCod, pdHoy, pnEstado, psMovAct) ' psUser)

    Call oBase.dCommitTrans
    Set oBase = Nothing
    Exit Sub

ErrorTransferirGarantiasARecuperaciones:
    If pbTran Then
        Call oBase.dRollbackTrans
    End If
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Sub
'***************************************************/*******************
Public Function CargarDatosCambioGarantia(ByVal psCtaCod As String, _
                                            ByVal pdFecSis As Date, _
                                            ByRef prsDatosGarantiaReemplazada As ADODB.Recordset, _
                                            ByRef prsDatosGarantiaReemplazant As ADODB.Recordset)
    
Dim oGarantia As COMDCredito.DCOMGarantia
Set oGarantia = New COMDCredito.DCOMGarantia
On Error GoTo ErrorCargarDatosCambioGarantia

   
   Set prsDatosGarantiaReemplazada = oGarantia.RecuperaGarantiasReemplazadas(psCtaCod)
   Set prsDatosGarantiaReemplazant = oGarantia.RecuperaGarantiasReemplazante(psCtaCod)
   Set oGarantia = Nothing
   
Exit Function

ErrorCargarDatosCambioGarantia:
    Err.Raise Err.Number, "Cargar Datos Garantias Reemplazante", Err.Description
End Function

Public Function RecuperaCantidadCreditosxGarantia(ByVal psNumGarant As String) As Integer
                                    

Dim oGarantia As COMDCredito.DCOMGarantia
Dim prsDatos As ADODB.Recordset

On Error GoTo ErrorCantidadCreditosxGarantia

   Set oGarantia = New COMDCredito.DCOMGarantia
   Set prsDatos = oGarantia.RecuperaCantidadCreditosxGarantia(psNumGarant)
   Set oGarantia = Nothing
   If Not (prsDatos.BOF Or prsDatos.EOF) Then
    RecuperaCantidadCreditosxGarantia = prsDatos!nCantidad
   End If
Exit Function

ErrorCantidadCreditosxGarantia:
    Err.Raise Err.Number, "Cargar Datos CantidadCreditosxGarantia", Err.Description
End Function
'EJVG20140422 ***
Public Sub CargarDatosDocsPendiente(ByVal pbFiltraAge As Boolean, ByVal psUserCod As String, ByVal psAgeCod As String, ByVal pdFecha As Date, _
                                    pRSGarXVenc As ADODB.Recordset, ByRef pRSGarVenc As ADODB.Recordset, _
                                    ByRef pRSPolXVenc As ADODB.Recordset, ByRef pRSPolVenc As ADODB.Recordset, _
                                    ByRef pRSTasacXVenc As ADODB.Recordset, ByRef pRSTasacVenc As ADODB.Recordset)
    On Error GoTo ErrCargarDatosDocsPendiente
    Dim obj As COMDCredito.DCOMGarantia
    Dim lsTblNombre As String
    
    lsTblNombre = psUserCod & Format(pdFecha, "yyyymmdd") & Format(Time, "hhmmss")
    Set obj = New COMDCredito.DCOMGarantia
    obj.GenerarDocsPendiente pbFiltraAge, psAgeCod, pdFecha, lsTblNombre
    Set pRSGarXVenc = obj.ConsultarDocsPendiente(1, lsTblNombre)
    Set pRSGarVenc = obj.ConsultarDocsPendiente(2, lsTblNombre)
    Set pRSPolXVenc = obj.ConsultarDocsPendiente(3, lsTblNombre)
    Set pRSPolVenc = obj.ConsultarDocsPendiente(4, lsTblNombre)
    Set pRSTasacXVenc = obj.ConsultarDocsPendiente(5, lsTblNombre)
    Set pRSTasacVenc = obj.ConsultarDocsPendiente(6, lsTblNombre)
    obj.EliminarDocsPendiente lsTblNombre
    Set obj = Nothing
    Exit Sub
ErrCargarDatosDocsPendiente:
    Err.Raise Err.Number, "Error al cargar método CargarDatosDocsPendiente", Err.Description
End Sub
Public Function MostrarDocsPendiente(ByVal pdFecha As Date) As Boolean
    Dim obj As New COMDCredito.DCOMGarantia
    MostrarDocsPendiente = obj.MostrarDocsPendiente(pdFecha)
    Set obj = Nothing
End Function
Public Function NivelVerDocsPendiente(ByVal psGrupoUsu As String, ByVal psUserCod As String, ByVal psRHCargoCod As String) As Integer
    Dim obj As New COMDCredito.DCOMGarantia
    NivelVerDocsPendiente = obj.NivelVerDocsPendiente(psGrupoUsu, psUserCod, psRHCargoCod)
    Set obj = Nothing
End Function
'END EJVG *******
'RECO20150326 ERS149-2014*********************************
Public Function RecuperaListaGarantPolizaPersona(ByVal psPersCod As String) As Recordset
    Dim obj As New COMDCredito.DCOMGarantia
    Set RecuperaListaGarantPolizaPersona = obj.RecuperaListaGarantPolizaPersona(psPersCod)
    Set obj = Nothing
End Function
Public Function RegistraSolicitudCoberturaSeguro(ByVal psPersCod As String, ByVal psNumGaran As String, ByVal psFecSiniestro As String, ByVal psGlosa As String _
                                                 , ByVal psNumCarta As String, ByVal psFecEmision As String, ByVal pnTipoSeguro As TipoSeguro _
                                                 , Optional ByVal psPersGestiona As String = "", Optional ByVal pnPersRelac As Integer = 0, Optional ByVal psAgeCodSiniestro As String = "" _
                                                 ) As Integer
    Dim obj As New COMDCredito.DCOMGarantia
    
    On Error GoTo Error
        
    RegistraSolicitudCoberturaSeguro = obj.RegistraSolicitudCoberturaSeguro(psPersCod, psNumGaran, Format(psFecSiniestro, "yyyy/MM/dd"), psGlosa, psNumCarta, Format(psFecEmision, "yyyy/MM/dd"), pnTipoSeguro, psPersGestiona, pnPersRelac, psAgeCodSiniestro)
    Set obj = Nothing
    Exit Function
    
Error:
    RegistraSolicitudCoberturaSeguro = 0
End Function
Public Function ObtieneDatosSolicitudCobert(ByVal psNumGarant As String, ByVal pnTpoSeguro As TipoSeguro) As Recordset
    Dim obj As New COMDCredito.DCOMGarantia
    Set ObtieneDatosSolicitudCobert = obj.ObtieneDatosSolicitudCobert(psNumGarant, pnTpoSeguro)
    Set obj = Nothing
End Function
Public Function ObtieneDocAdjuntosSolicitud(ByVal pnIdSolicitud As Integer) As Recordset
    Dim obj As New COMDCredito.DCOMGarantia
    Set ObtieneDocAdjuntosSolicitud = obj.ObtieneDocAdjuntosSolicitud(pnIdSolicitud)
    Set obj = Nothing
End Function
Public Function ObtieneNumCartaSolicCobert(ByVal psPersCod As String) As String
    Dim obj As New COMDCredito.DCOMGarantia
    Dim rs As New Recordset
    
    Set rs = obj.ObtieneNumCartaSolicCobert(psPersCod)
    If Not (rs.EOF And rs.BOF) Then
        ObtieneNumCartaSolicCobert = rs!cNumCarta
    End If
    Set obj = Nothing
End Function
Public Sub RegistraResolucionSolicitud(ByVal pnIdSolicitud As Integer, ByVal pnTpoDoc As Integer, ByVal pnMoneda As Integer, ByVal pnMonto As Double _
                                             , ByVal psFecEmision As String, ByVal psFecDeposito As String, ByVal psNumero As String, ByVal psCtaCodBanco As String _
                                             , ByVal pnEstado As Integer, ByVal psGlosa As String, ByVal pnTpoSeguro As TipoSeguro)
    Dim obj As New COMDCredito.DCOMGarantia
    Call obj.RegistraResolucionSolicitud(pnIdSolicitud, pnTpoDoc, pnMoneda, pnMonto, psFecEmision, psFecDeposito, psNumero, psCtaCodBanco, pnEstado, psGlosa, pnTpoSeguro)
    Set obj = Nothing
End Sub
Public Sub RegsitraHistorialSegSolic(ByVal pnIdSolicitud As Integer, ByVal psDescripcion As String, ByVal psNumCarta As String, ByVal psGlosa As String _
                                    , ByVal psFecRespuesta As String, ByVal psFecEstado As String, ByVal pnEstado As SolicitudCoberEstado, ByVal pnTpoSeguro As TipoSeguro)
    Dim obj As New COMDCredito.DCOMGarantia
    Call obj.RegsitraHistorialSegSolic(pnIdSolicitud, psDescripcion, psNumCarta, psGlosa, psFecRespuesta, psFecEstado, pnEstado, pnTpoSeguro)
    Set obj = Nothing
End Sub
Public Function BuscaSolicitudCobert(ByVal psNumCarta As String, ByVal psEstados As String, ByVal pnTpoSeguro As TipoSeguro) As String
    Dim obj As New COMDCredito.DCOMGarantia
    Dim rs As New Recordset
    Set rs = obj.BuscaSolicitudCobert(psNumCarta, psEstados, pnTpoSeguro)
    BuscaSolicitudCobert = ""
    If Not (rs.BOF And rs.EOF) Then
        BuscaSolicitudCobert = rs!nCodigo
    End If
    Set obj = Nothing
End Function
Public Sub RegistraRechazoSolicitud(ByVal pnIdSolicitud As Integer, ByVal psFecRespuesta As String, ByVal psGlosa As String, ByVal psNumCarta As String)
    Dim obj As New COMDCredito.DCOMGarantia
    Call obj.RegistraRechazoSolicitud(pnIdSolicitud, psFecRespuesta, psGlosa, psNumCarta)
    Set obj = Nothing
End Sub
Public Sub ActualizaEstadoSolicCobert(ByVal pnEstado As SolicitudCoberEstado, ByVal pnIdSolicitud As Integer, ByVal psFecReg As String, ByVal pnTpoSeguro As TipoSeguro)
    Dim obj As New COMDCredito.DCOMGarantia
    Call obj.ActualizaEstadoSolicCobert(pnEstado, pnIdSolicitud, psFecReg, pnTpoSeguro)
    Set obj = Nothing
End Sub
Public Function ExisteSolicCobert(ByVal psNumGarant As String) As Boolean
    Dim obj As New COMDCredito.DCOMGarantia
    Dim rs As New Recordset
    Set rs = obj.ObtieneSolicitCoberXGarant(psNumGarant)
    ExisteSolicCobert = False
    If Not (rs.BOF And rs.EOF) Then
        ExisteSolicCobert = True
    End If
    Set obj = Nothing
End Function
Public Function ObtieneDatosExtornarSeg(ByVal psPersCod As String, ByVal psFecDesde As String, ByVal psFecHasta As String, ByVal pnTpoSeguro As TipoSeguro) As Recordset
    Dim obj As New COMDCredito.DCOMGarantia
    Set ObtieneDatosExtornarSeg = obj.ObtieneDatosExtornarSeg(psPersCod, psFecDesde, psFecHasta, pnTpoSeguro)
    Set obj = Nothing
End Function
Public Sub ExtornoSolicitud(ByVal pnIdSolicitud As Integer, ByVal pnEstado As Integer, ByVal psFecReg As String, ByVal pnTpoSeguro As TipoSeguro)
    Dim obj As New COMDCredito.DCOMGarantia
    Call obj.ExtornoSolicitud(pnIdSolicitud, pnEstado, psFecReg, pnTpoSeguro)
    Set obj = Nothing
End Sub
Public Sub EliminaRegistroEditar(ByVal pnIdSolid As Integer, ByVal pnTpoSeguro As TipoSeguro)
    Dim obj As New COMDCredito.DCOMGarantia
    Call obj.EliminaRegistroEditar(pnIdSolid, pnTpoSeguro)
    Set obj = Nothing
End Sub
Public Sub ActulizaSolicitudSeg(ByVal pnIdSolic As Integer, ByVal psFecSiniestro As String, ByVal psGlosa As String, ByVal pnTpoSeguro As TipoSeguro)
    Dim obj As New COMDCredito.DCOMGarantia
    Call obj.ActulizaSolicitudSeg(pnIdSolic, psFecSiniestro, psGlosa, pnTpoSeguro)
    Set obj = Nothing
End Sub
Public Function SegSolicitudHistorial(ByVal pnIdSolicitud As Long, ByVal pnTpoSeguro As TipoSeguro) As ADODB.Recordset
    Dim obj As New COMDCredito.DCOMGarantia
    Set SegSolicitudHistorial = obj.SegSolicitudHistorial(pnIdSolicitud, pnTpoSeguro)
    Set obj = Nothing
End Function
'RECO FIN*****************************
'EJVG20150326 ***
Public Function ObtenerDocsGarantias(Optional ByVal pnCodConfig As Long = 0, Optional ByVal pnClasificacion As Long = 0) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    'Set ObtenerDocsGarantias = oGarant.ObtenerDocsGarantias(pnCodConfig)
    Set ObtenerDocsGarantias = oGarant.ObtenerDocsGarantias(pnCodConfig, pnClasificacion)
    Set oGarant = Nothing
End Function
Public Function ObtenerDocsGarantiasXID(Optional ByVal pnCodConfig As Long = 0) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenerDocsGarantiasXID = oGarant.ObtenerDocsGarantiasXID(pnCodConfig)
    Set oGarant = Nothing
End Function
Public Function ObtenerConfigGarantParam(Optional ByVal pnClasificacion As Long = 0) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenerConfigGarantParam = oGarant.ObtenerConfigGarantParam(pnClasificacion)
    Set oGarant = Nothing
End Function
Public Function ObtenerConfigGarantParamTpoValor(Optional ByVal pnClasificacion As Long = 0) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenerConfigGarantParamTpoValor = oGarant.ObtenerConfigGarantParamTpoValor(pnClasificacion)
    Set oGarant = Nothing
End Function
Public Function ObtenerConfigGarantTpoValor(Optional ByVal pnCodConfig As Long = 0) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenerConfigGarantTpoValor = oGarant.ObtenerConfigGarantTpoValor(pnCodConfig)
    Set oGarant = Nothing
End Function
Public Function ObtenerConfigGarant() As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenerConfigGarant = oGarant.ObtenerConfigGarant()
    Set oGarant = Nothing
End Function
Public Function ObtenerConfigGarantXID(ByVal pnGarantConfigID As Long) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenerConfigGarantXID = oGarant.ObtenerConfigGarantXID(pnGarantConfigID)
    Set oGarant = Nothing
End Function

Public Function GarantValidaCod(ByVal pnCod As Long, ByVal pnTipo As Long) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set GarantValidaCod = oGarant.GarantValidaCod(pnCod, pnTipo)
    Set oGarant = Nothing
End Function

Public Function RegistrarConfigGarant(ByVal pnCod As Long, ByVal pnClasificacion As Integer, ByVal pnObjGarantCod As Integer, ByVal pcObjGarantDesc As String, _
                                    ByVal pnConstitucionCod As Integer, ByVal pcConstitucionDesc As String, ByVal pnLeasingCod As Integer, _
                                    ByVal pcLeasingDesc As String, ByVal pnSinTramite As Integer, ByVal pbActivado As Boolean, ByVal pMatDocs As Variant, ByVal pMatTpoValor As Variant) As String
On Error GoTo ErrRegistrarConfigGarant
Dim oGarant As COMDCredito.DCOMGarantia
Dim bTransac As Boolean
Dim nGarantConfigID As Long
Dim i As Long

RegistrarConfigGarant = ""

Set oGarant = New COMDCredito.DCOMGarantia
Call oGarant.dBeginTrans
bTransac = True

nGarantConfigID = oGarant.RegistarActConfigGarant(pnCod, pnClasificacion, pnObjGarantCod, pcObjGarantDesc, pnConstitucionCod, pcConstitucionDesc, pnLeasingCod, pcLeasingDesc, pnSinTramite, pbActivado)

Call oGarant.EliminaConfigGarantDocs(nGarantConfigID)
For i = 0 To UBound(pMatDocs)
    Call oGarant.RegistrarConfigGarantDocs(nGarantConfigID, CLng(pMatDocs(i)))
Next i

Call oGarant.EliminaConfigGarantTpoValor(nGarantConfigID)
For i = 0 To UBound(pMatTpoValor)
    Call oGarant.RegistrarConfigGarantTpoValor(nGarantConfigID, CLng(pMatTpoValor(i)))
Next i
Call oGarant.dCommitTrans
bTransac = False

Exit Function
ErrRegistrarConfigGarant:
    RegistrarConfigGarant = "Ocurrió un error al procesar el registro de la configuración de la garantía"
    If bTransac Then
        Call oGarant.dRollbackTrans
        Set oGarant = Nothing
    End If
    Err.Raise Err.Number, "Ocurrió un error al procesar el registro de la configuración de la garantía", Err.Description
End Function

Public Function ObtenerConfigGarantClas(ByVal pnClasificacion As Long) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenerConfigGarantClas = oGarant.ObtenerConfigGarantClas(pnClasificacion)
    Set oGarant = Nothing
End Function
'END EJVG ***
'RECO20150416 ERS001-2015***************************
Public Function ObtieneCreditosPorGarantia(ByVal psNumGarant As String)
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtieneCreditosPorGarantia = oGarant.ObtieneCreditosPorGarantia(psNumGarant)
    Set oGarant = Nothing
End Function
'RECO FIN*******************************************
'RECO20150509 ERS023-2015***************************
Public Function ObtieneGarantMultiRiesgoMYPE(ByVal psPersCod As String, Optional ByVal psCtaCod As String = "")
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtieneGarantMultiRiesgoMYPE = oGarant.ObtieneGarantMultiRiesgoMYPE(psPersCod, psCtaCod)
    Set oGarant = Nothing
End Function
Public Sub RegistroGastoMicroSegMYPE(ByVal psCtaCod As String, ByVal psNumGarant As String, ByVal pnValorAseg As Double, ByVal psFecReg As String)
    Dim oGarant As New COMDCredito.DCOMGarantia
    Call oGarant.RegistroGastoMicroSegMYPE(psCtaCod, psNumGarant, pnValorAseg, psFecReg)
    Set oGarant = Nothing
End Sub
Public Function RecuperaDatosClienteDJMYPE(ByVal psCtaCod As String)
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set RecuperaDatosClienteDJMYPE = oGarant.RecuperaDatosClienteDJMYPE(psCtaCod)
    Set oGarant = Nothing
End Function
Public Function RecuperaDatosGarantDJMYPE(ByVal psNumGarant As String)
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set RecuperaDatosGarantDJMYPE = oGarant.RecuperaDatosGarantDJMYPE(psNumGarant)
    Set oGarant = Nothing
End Function
Public Sub RegistrarNumPoliza(ByVal psCtaCod As String, ByVal psNumPoliza As String)
    Dim oGarant As New COMDCredito.DCOMGarantia
    Call oGarant.RegistrarNumPoliza(psCtaCod, psNumPoliza)
    Set oGarant = Nothing
End Sub
'RECO FIN*******************************************
'RECO20150715 ERS034-2015****************************************************
Public Function RecuperaListaGarantMoviliariaPersona(ByVal psPersCod As String, Optional ByVal psCtaCod As String = "") As Recordset
    Dim obj As New COMDCredito.DCOMGarantia
    Set RecuperaListaGarantMoviliariaPersona = obj.RecuperaListaGarantMoviliariaPersona(psPersCod, psCtaCod)
    Set obj = Nothing
End Function
'RECO FIN********************************************************************
'RECO20160125 ERS001-2016****************************************************
Public Function ObtieneDatosGarantRealVenc(ByVal psFecha As String, ByVal psAgeCad As String, ByVal psCadTpoCred As String, ByVal psMoneda As String, ByVal psEstado As String) As ADODB.Recordset
    Dim obj As New COMDCredito.DCOMGarantia
    Set ObtieneDatosGarantRealVenc = obj.ObtieneDatosGarantRealVenc(psFecha, psAgeCad, psCadTpoCred, psMoneda, psEstado)
    Set obj = Nothing
End Function
Public Function GarantRealRegistraCartaVenc(ByVal psCtaCod As String, ByVal psUser As String) As String
    Dim obj As New COMDCredito.DCOMGarantia
    GarantRealRegistraCartaVenc = obj.GarantRealRegistraCartaVenc(psCtaCod, psUser)
    Set obj = Nothing
End Function

'RECO FIN********************************************************************
'EJVG20150206 ***
'Public Function GrabarDatosGarantia(ByRef psGarantiaID As String, ByVal pRsGarantia As ADODB.Recordset, ByRef pvPropietario() As tPropietarioRelacion, ByRef pvValorizacion() As tValorizacion, ByRef pvTramiteLegal() As tTramiteLegal) As Boolean
Public Function GrabarDatosGarantia(ByRef psGarantiaID As String, ByVal pRsGarantia As ADODB.Recordset, ByRef pvPropietario As Variant, ByRef pvValorizacion As Variant, ByRef pvTramiteLegal As Variant, ByRef pbAutoriza As Boolean) As Boolean 'JGPA 20180324 added pbAutoriza TI-ERS026-2018
    Dim obj As New COMDCredito.DCOMGarantia
    Dim i As Integer, iDet As Integer
    Dim lsNumGarant As String
    Dim bTrans As Boolean
    Dim bAutorizaElimina As Boolean 'JGPA TI-ERS026-2018
    Dim lvPropietario() As tPropietarioRelacion
    Dim lvValorizacion() As tValorizacion
    Dim lvTramiteLegal() As tTramiteLegal
    Dim nTpoDoc As Integer
    Dim nContratoId As Integer
    'Set lvPropietario() = New tPropietarioRelacion
    'Set lvValorizacion() = New tValorizacion
    'Set lvTramiteLegal() = New tTramiteLegal
    lvPropietario = pvPropietario
    lvValorizacion = pvValorizacion
    lvTramiteLegal = pvTramiteLegal
    bAutorizaElimina = pbAutoriza 'JGPA TI-ERS026-2018
    
    nContratoId = 0
    nTpoDoc = pRsGarantia("nDocTpo")
    If (pRsGarantia("nClasificacion") = 1 And pRsGarantia("nBienGarantia") = 6) Then
        If (pRsGarantia("nTpoBienContrato") = 2) And nTpoDoc = 1 Then
            nTpoDoc = 104
        ElseIf (pRsGarantia("nTpoBienContrato") = 2) And nTpoDoc <> 1 Then
            nContratoId = nTpoDoc
            nTpoDoc = 161
        End If
    End If
    
    On Error GoTo ErrRegistrar
    
    obj.dBeginTrans
    bTrans = True
    
    If psGarantiaID = "" Then
        lsNumGarant = obj.RegistrarGarantia(pRsGarantia("nMoneda"), pRsGarantia("nClasificacion"), pRsGarantia("nBienGarantia"), pRsGarantia("cEmisorCod"), nTpoDoc, pRsGarantia("cDocNro"), pRsGarantia("dConstatacion"), pRsGarantia("sUltimacionActualizacion"), pRsGarantia("nTpoBienContrato"), nContratoId)
    Else
        lsNumGarant = psGarantiaID
        obj.ActualizarGarantia lsNumGarant, pRsGarantia("nMoneda"), pRsGarantia("nClasificacion"), pRsGarantia("nBienGarantia"), pRsGarantia("cEmisorCod"), nTpoDoc, pRsGarantia("cDocNro"), pRsGarantia("dConstatacion"), pRsGarantia("sUltimacionActualizacion"), pRsGarantia("nTpoBienContrato"), nContratoId
        If pRsGarantia("bCambiadoBienFuturo") = True Then
            obj.GrabarPolizaBienFuturo lsNumGarant, pRsGarantia("dFechaSistema"), pRsGarantia("sMovNroCreditoPoliza")
            
            Dim pnMonto As Double
            Dim pdFechaVigencia As Date
            Dim pnTasaInteres As Double
            Dim psCtaCod As String
            Dim pnTipoPeriodo As Integer
            Dim nTasaCostoEfectivoAnual As Double
            Dim oCredito As COMNCredito.NCOMCredito
            Set oCredito = New COMNCredito.NCOMCredito
        
            Dim oRsDatos As ADODB.Recordset
            Set oRsDatos = New ADODB.Recordset
            
            Set oRsDatos = obj.ObtenerDatosCredito(lsNumGarant)

            If Not oRsDatos.BOF And Not oRsDatos.EOF Then
                pnMonto = oRsDatos!nMontoCol
                pdFechaVigencia = oRsDatos!dVigencia
                pnTasaInteres = oRsDatos!nTasaInteres
                psCtaCod = oRsDatos!cCtaCod
            End If
            
            Set oRsDatos = New ADODB.Recordset
            Set oRsDatos = obj.IdentificarTipoPeriodoMIVIENDA(psCtaCod)
            If Not oRsDatos.BOF And Not oRsDatos.EOF Then
                pnTipoPeriodo = oRsDatos!nTpPeriodo
            End If
            Dim MatCalend As Variant
            MatCalend = ObtenerCalendarioCredito(psCtaCod, obj)

            nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(pdFechaVigencia, pnMonto, MatCalend, pnTasaInteres, psCtaCod, pnTipoPeriodo) 'JUEZ 20140814
            Call obj.ActualizaTCEA(psCtaCod, nTasaCostoEfectivoAnual)
            Set oRsDatos = Nothing
        End If
    End If
    
    If pRsGarantia("bEliminarRegCobCredProceso") Then
        obj.EliminarRegCoberturaCredProceso lsNumGarant
    End If
    
    'Propietarios
    obj.EliminarGarantiaPersonas lsNumGarant
    For i = 1 To UBound(lvPropietario)
        If lvPropietario(i).nGarantiaCambiosFila <> GarantFilaEliminada And lvPropietario(i).nGarantiaCambiosFila <> GarantFilaOculta Then
            obj.RegistrarGarantiaPersona lsNumGarant, lvPropietario(i).sPersCod, lvPropietario(i).nRelacionTpo
        End If
    Next
    
    'Primero eliminanos las Valorizaciones para evitar duplicidad
    For i = 1 To UBound(lvValorizacion)
        If lvValorizacion(i).nGarantiaCambiosFila = GarantFilaEliminada Then
            Select Case lvValorizacion(i).nValorizacionTpo
                Case eGarantiaTipoValorizacion.ValorDirectoTotal
                    obj.EliminarValorizacionDirectaTotal lsNumGarant, lvValorizacion(i).dFecha
                Case eGarantiaTipoValorizacion.ValorDirectoDetallado
                    obj.EliminarValorizacionDirectaDet lsNumGarant, lvValorizacion(i).dFecha
                Case eGarantiaTipoValorizacion.GarantiaAutoliquidable
                    obj.EliminarValorizacionAutoLiquidable lsNumGarant, lvValorizacion(i).dFecha
                Case eGarantiaTipoValorizacion.TasacionInmobiliaria
                    obj.EliminarValorizacionTasacInmobiliaria lsNumGarant, lvValorizacion(i).dFecha
                Case eGarantiaTipoValorizacion.TasacionVehicular
                    obj.EliminarValorizacionTasacVehicular lsNumGarant, lvValorizacion(i).dFecha
                Case eGarantiaTipoValorizacion.OtrasTasacionesMobiliarias
                    obj.EliminarValorizacionOtrasTasacMobiliarias lsNumGarant, lvValorizacion(i).dFecha
                Case eGarantiaTipoValorizacion.GravamenFavorOtraIFi 'EJVG20160226
                    obj.EliminarValorizacionGravamenOtraIFi lsNumGarant, lvValorizacion(i).dFecha
            End Select
            obj.EliminarValorizacion lsNumGarant, lvValorizacion(i).dFecha
        End If
    Next
    'Insertamos y Actualizamos las Valorizaciones
    For i = 1 To UBound(lvValorizacion)
        If lvValorizacion(i).nGarantiaCambiosFila = GarantFilaNueva Then
            obj.RegistrarValorizacion lsNumGarant, lvValorizacion(i).dFecha, lvValorizacion(i).sGlosa, lvValorizacion(i).sUltimaActualizacion
        ElseIf lvValorizacion(i).nGarantiaCambiosFila = GarantFilaModificada Then
            obj.ActualizarValorizacion lsNumGarant, lvValorizacion(i).dFecha, lvValorizacion(i).sGlosa, lvValorizacion(i).sUltimaActualizacion
        End If
        Select Case lvValorizacion(i).nValorizacionTpo
            Case eGarantiaTipoValorizacion.ValorDirectoTotal
                If lvValorizacion(i).nGarantiaCambiosFila = GarantFilaNueva Then
                    obj.RegistrarValorizacionDirectaTotal lsNumGarant, lvValorizacion(i).dFecha, lvValorizacion(i).vValorDirectoTotal
                ElseIf lvValorizacion(i).nGarantiaCambiosFila = GarantFilaModificada Then
                    obj.ActualizarValorizacionDirectaTotal lsNumGarant, lvValorizacion(i).dFecha, lvValorizacion(i).vValorDirectoTotal
                End If
            Case eGarantiaTipoValorizacion.ValorDirectoDetallado
                If lvValorizacion(i).nGarantiaCambiosFila = GarantFilaNueva Or lvValorizacion(i).nGarantiaCambiosFila = GarantFilaModificada Then
                    obj.EliminarValorizacionDirectaDet lsNumGarant, lvValorizacion(i).dFecha
                    For iDet = 1 To UBound(lvValorizacion(i).vValorDirDet)
                        obj.RegistrarValorizacionDirectaDet lsNumGarant, lvValorizacion(i).dFecha, iDet, lvValorizacion(i).vValorDirDet(iDet)
                    Next
                End If
            Case eGarantiaTipoValorizacion.GarantiaAutoliquidable
                If lvValorizacion(i).nGarantiaCambiosFila = GarantFilaNueva Then
                    obj.RegistrarValorizacionAutoLiquidable lsNumGarant, lvValorizacion(i).dFecha, lvValorizacion(i).vValorAutoLiquidable
                ElseIf lvValorizacion(i).nGarantiaCambiosFila = GarantFilaModificada Then
                    obj.ActualizarValorizacionAutoLiquidable lsNumGarant, lvValorizacion(i).dFecha, lvValorizacion(i).vValorAutoLiquidable
                End If
            Case eGarantiaTipoValorizacion.TasacionInmobiliaria
                If lvValorizacion(i).nGarantiaCambiosFila = GarantFilaNueva Then
                    obj.RegistrarValorizacionTasacInmobiliaria lsNumGarant, lvValorizacion(i).dFecha, lvValorizacion(i).vValorTasacionInmobiliaria
                ElseIf lvValorizacion(i).nGarantiaCambiosFila = GarantFilaModificada Then
                    obj.ActualizarValorizacionTasacInmobiliaria lsNumGarant, lvValorizacion(i).dFecha, lvValorizacion(i).vValorTasacionInmobiliaria
                End If
            Case eGarantiaTipoValorizacion.TasacionVehicular
                If lvValorizacion(i).nGarantiaCambiosFila = GarantFilaNueva Then
                    obj.RegistrarValorizacionTasacVehicular lsNumGarant, lvValorizacion(i).dFecha, lvValorizacion(i).vValorTasacionVehicular
                ElseIf lvValorizacion(i).nGarantiaCambiosFila = GarantFilaModificada Then
                    obj.ActualizarValorizacionTasacVehicular lsNumGarant, lvValorizacion(i).dFecha, lvValorizacion(i).vValorTasacionVehicular
                End If
            Case eGarantiaTipoValorizacion.OtrasTasacionesMobiliarias
                If lvValorizacion(i).nGarantiaCambiosFila = GarantFilaNueva Then
                    obj.RegistrarValorizacionOtrasTasacMobiliarias lsNumGarant, lvValorizacion(i).dFecha, lvValorizacion(i).vValorTasacionMobiliariaOtras
                ElseIf lvValorizacion(i).nGarantiaCambiosFila = GarantFilaModificada Then
                    obj.ActualizarValorizacionOtrasTasacMobiliarias lsNumGarant, lvValorizacion(i).dFecha, lvValorizacion(i).vValorTasacionMobiliariaOtras
                End If
            Case eGarantiaTipoValorizacion.GravamenFavorOtraIFi 'EJVG20160226
                If lvValorizacion(i).nGarantiaCambiosFila = GarantFilaNueva Or lvValorizacion(i).nGarantiaCambiosFila = GarantFilaModificada Then
                    If lvValorizacion(1).nVTpRegPub = 2 Then 'JOEP20180524 ERS009-2018
                        obj.EliminarValorizacionGravamenPerNatJudDet lsNumGarant, lvValorizacion(i).dFecha  'JOEP20180524 ERS009-2018
                    Else
                        obj.EliminarValorizacionGravamenOtraIFi lsNumGarant, lvValorizacion(i).dFecha
                    End If
                    obj.RegistrarValorizacionGravamenOtraIFi lsNumGarant, lvValorizacion(i).dFecha, lvValorizacion(i).vValorGravamenFavorOtraIFi
                End If
        End Select
        'Consolidamos datos de Garantías
        If lvValorizacion(i).nGarantiaCambiosFila = GarantFilaNueva Or lvValorizacion(i).nGarantiaCambiosFila = GarantFilaModificada Then
            obj.ConsolidarGarantiaValoriza lsNumGarant, lvValorizacion(i).dFecha
        End If
    Next
    
    'Primero eliminanos los Tramites Legales para evitar duplicidad
    For i = 1 To UBound(lvTramiteLegal)
        If lvTramiteLegal(i).nGarantiaCambiosFila = GarantFilaEliminada Then
            obj.EliminarTramiteLegal lsNumGarant, lvTramiteLegal(i).dFecha
        End If
    Next
    For i = 1 To UBound(lvTramiteLegal)
        If lvTramiteLegal(i).nGarantiaCambiosFila = GarantFilaNueva Then
            obj.RegistrarTramiteLegal lsNumGarant, lvTramiteLegal(i)
        ElseIf lvTramiteLegal(i).nGarantiaCambiosFila = GarantFilaModificada Then
            obj.ActualizarTramiteLegal lsNumGarant, lvTramiteLegal(i)
        End If
    Next
    
    'Damos de baja la garantía
    '***JGPA TI-ERS026-2018
    If bAutorizaElimina = True Then
        obj.DarDeBajaGarantia lsNumGarant
    End If
    '*** End JGPA
    
    obj.ConsolidarGarantia lsNumGarant
    psGarantiaID = lsNumGarant
    
    obj.dCommitTrans
    bTrans = False
    GrabarDatosGarantia = True
    
    Set obj = Nothing
    Exit Function
ErrRegistrar:
    GrabarDatosGarantia = False
    If bTrans Then
        obj.dRollbackTrans
        Set obj = Nothing
    End If
    Err.Raise Err.Number, "Ocurrió un error al registrar la Garantía", Err.Description
End Function
Public Function GrabarCoberturasConfig(ByVal psTpoProdCod As String, ByVal pbCliPreferencial As Boolean, ByRef pvBienGarantia As Variant) As Boolean
    Dim obj As New COMDCredito.DCOMGarantia
    Dim bTrans As Boolean
    Dim Index As Integer
    
    On Error GoTo ErrGrabarCoberturasConfig
    
    obj.dBeginTrans
    bTrans = True
    
    obj.EliminarCoberturaConfig psTpoProdCod, pbCliPreferencial
    For Index = 1 To UBound(pvBienGarantia, 2)
        obj.RegistrarCoberturaConfig psTpoProdCod, pbCliPreferencial, pvBienGarantia(1, Index), pvBienGarantia(2, Index), pvBienGarantia(3, Index)
    Next
    
    bTrans = False
    obj.dCommitTrans
    GrabarCoberturasConfig = True
    Exit Function
ErrGrabarCoberturasConfig:
    GrabarCoberturasConfig = False
    If bTrans Then
        obj.dRollbackTrans
        Set obj = Nothing
    End If
    Err.Raise Err.Number, "Ocurrió un error al grabar la configuración de coberturas", Err.Description
End Function
Public Function GrabarCoberturasConfigEX(ByVal pnRetiroIntAdelantado As Currency, ByVal pnRetiroIntMensual As Currency, ByVal pnRetiroIntFinal As Currency, _
    ByVal pnRetiroIntAdelantadoCAMP As Currency, ByVal pnRetiroIntMensualCAMP As Currency, ByVal pnRetiroIntFinalCAMP As Currency) As Boolean
    Dim obj As New COMDCredito.DCOMGarantia
    Dim bTrans As Boolean
    
    On Error GoTo ErrGrabarCoberturasConfigEX
    
    obj.dBeginTrans
    bTrans = True
    
    obj.ActualizarCoberturaConfigRapiFlash 0, eGarantiaTipoRetiroRapiFlash.nRetiroIntAdelantado, pnRetiroIntAdelantado
    obj.ActualizarCoberturaConfigRapiFlash 0, eGarantiaTipoRetiroRapiFlash.nRetiroIntMensual, pnRetiroIntMensual
    obj.ActualizarCoberturaConfigRapiFlash 0, eGarantiaTipoRetiroRapiFlash.nRetiroIntFinal, pnRetiroIntFinal
    
    obj.ActualizarCoberturaConfigRapiFlash 88, eGarantiaTipoRetiroRapiFlash.nRetiroIntAdelantado, pnRetiroIntAdelantadoCAMP
    obj.ActualizarCoberturaConfigRapiFlash 88, eGarantiaTipoRetiroRapiFlash.nRetiroIntMensual, pnRetiroIntMensualCAMP
    obj.ActualizarCoberturaConfigRapiFlash 88, eGarantiaTipoRetiroRapiFlash.nRetiroIntFinal, pnRetiroIntFinalCAMP
    
    bTrans = False
    obj.dCommitTrans
    GrabarCoberturasConfigEX = True
    Exit Function
ErrGrabarCoberturasConfigEX:
    GrabarCoberturasConfigEX = False
    If bTrans Then
        obj.dRollbackTrans
        Set obj = Nothing
    End If
    Err.Raise Err.Number, "Ocurrió un error al grabar las excepciones de la configuración de coberturas", Err.Description
End Function
Public Sub CargarDatosxGravamen(ByVal psCtaCod As String, ByVal pbCartaFianza As Boolean, _
                                        ByVal pdFecSis As Date, ByRef psPersCodTit As String, ByRef pbCliPreferencial As Boolean, _
                                        ByRef prsDatos As ADODB.Recordset, _
                                        ByRef prsPersonas As ADODB.Recordset, _
                                        ByRef prsGarantiasPersonas As ADODB.Recordset, _
                                        ByRef prsGarantiasCredito As ADODB.Recordset, _
                                        ByVal psTpoProdCod As String, _
                                        ByRef prsColocGarantiaAnt As ADODB.Recordset)
    Dim obj As COMDCredito.DCOMGarantia
    Dim oNCred As COMNCredito.NCOMCredito
    Dim lrsPersonasTmp As New ADODB.Recordset, lrsGarantiasPersonas As New ADODB.Recordset
    Dim lsPersCodLista As String, lsNumGarantLista As String
    
    On Error GoTo ErrorCargarDatosGarantCredito

    Set obj = New COMDCredito.DCOMGarantia
    Set oNCred = New COMNCredito.NCOMCredito
    
    'Recupera Datos del Crédito
    If pbCartaFianza Then
        Set prsDatos = obj.RecuperaDatosCFxGravamen(psCtaCod)
    Else
        Set prsDatos = obj.RecuperaDatosCreditoxGravamen(psCtaCod)
    End If
    
    If prsDatos.EOF Then Exit Sub
    
    psPersCodTit = prsDatos!cPersCodTit
    pbCliPreferencial = oNCred.ValidarClientePreferencial(psPersCodTit)
    
    'Recupera los propietarios/intervinientes en la Garantía
    Set prsPersonas = obj.RecuperaPersonaCreditoxGravamen(psCtaCod)
    Set lrsPersonasTmp = prsPersonas.Clone
    If Not lrsPersonasTmp.EOF Then
        Do While Not lrsPersonasTmp.EOF
            lsPersCodLista = lsPersCodLista & lrsPersonasTmp!cPersCod & ","
            lrsPersonasTmp.MoveNext
        Loop
        lsPersCodLista = Mid(lsPersCodLista, 1, Len(lsPersCodLista) - 1)
    End If
    lrsPersonasTmp.Close
    Set lrsPersonasTmp = Nothing
    
    'Recupera las garantías en la que los propietarios/intervinientes son titulares
    Set prsGarantiasPersonas = obj.RecuperaGarantiasPersonaxGravamen(psCtaCod, lsPersCodLista, pbCliPreferencial, psTpoProdCod)
    Set lrsGarantiasPersonas = prsGarantiasPersonas.Clone
    If Not lrsGarantiasPersonas.EOF Then
        Do While Not lrsGarantiasPersonas.EOF
            lsNumGarantLista = lsNumGarantLista & lrsGarantiasPersonas!cNumGarant & ","
            lrsGarantiasPersonas.MoveNext
        Loop
        lsNumGarantLista = Mid(lsNumGarantLista, 1, Len(lsNumGarantLista) - 1)
    End If
    lrsGarantiasPersonas.Close
    Set lrsGarantiasPersonas = Nothing
       
    'Recupera las garantías ya vinculadas con el crédito
    Set prsGarantiasCredito = obj.RecuperaGarantiasCreditoxGravamen(psCtaCod, lsNumGarantLista)
    
    'Recupera créditos anteriores (Ampliaciciones y Refinanciaciones)
    If prsGarantiasCredito.RecordCount = 0 Then
        Set prsColocGarantiaAnt = obj.RecuperaGarantiasCreditoAnterior(psCtaCod)
    End If

    'Recupera si tiene Coberturas que faltan migrar
    'psCadenaGarantiasAMigrar = obj.CadenaGarantiasPendienteMigracion(psCtaCod)'EJVG20160320->En el registro de cobertura se podrá actualizar las garantías por migrar

    Set oNCred = Nothing
    Set obj = Nothing
    Exit Sub
ErrorCargarDatosGarantCredito:
    Err.Raise Err.Number, "Error al Cargar Datos Garantia Credito", Err.Description
End Sub
'Public Function GrabarCoberturaGarantia(ByVal psCtaCod As String, ByVal psTpoProdCod As String, ByVal pnMontoColocado As Currency, ByRef pvDatos() As tGarantiaGravamen) As Boolean
'    Dim obj As New COMDCredito.DCOMGarantia
'    Dim bTrans As Boolean
'    Dim i As Integer
'
'    On Error GoTo ErrGrabarCoberturaGarantia
'
'    obj.dBeginTrans
'    bTrans = True
'
'    If UBound(pvDatos) > 0 Then
'        obj.EliminarCoberturaGarantia psCtaCod
'        obj.RegistrarColocGravamen psCtaCod, psTpoProdCod, pnMontoColocado
'        For i = 1 To UBound(pvDatos)
'            obj.RegistrarCoberturaGarantia pvDatos(i).sNumgarant, pvDatos(i).nMoneda, psCtaCod, pvDatos(i).dFechaValorizacion, pvDatos(i).nSaldoCobertura, pvDatos(i).nRatio, pvDatos(i).nMaxCobertura, pvDatos(i).nCobertura, pvDatos(i).dFechaTramiteLegal
'        Next
'    End If
'
'    obj.dCommitTrans
'    bTrans = False
'    Set obj = Nothing
'
'    GrabarCoberturaGarantia = True
'    Exit Function
'ErrGrabarCoberturaGarantia:
'    GrabarCoberturaGarantia = False
'    If bTrans Then
'        obj.dRollbackTrans
'        Set obj = Nothing
'    End If
'    Err.Raise Err.Number, "Ocurrió un error al grabar la cobertura de garantías con el préstamo " & psCtaCod, Err.Description
'End Function
Public Function ActualizarGarantiaInmobiliaria(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByVal pnClase As Integer, _
                                                ByVal pnNroLocales As Integer, ByVal pnNroPisos As Integer, ByVal pnCategoria As Integer, _
                                                ByVal pnNroSotanos As Integer, ByVal pnAnioConstruccion As Integer) As Boolean
    Dim obj As New COMDCredito.DCOMGarantia
    Dim bTrans As Boolean
    
    On Error GoTo ErrActualizarGarantiaInmobiliaria
    
    obj.dBeginTrans
    bTrans = True
    
    obj.ActualizarGarantiaInmobiliaria psNumGarant, pdRegistro, pnClase, pnNroLocales, pnNroPisos, pnCategoria, pnNroSotanos, pnAnioConstruccion
        
    bTrans = False
    obj.dCommitTrans
    ActualizarGarantiaInmobiliaria = True
    Exit Function
ErrActualizarGarantiaInmobiliaria:
    ActualizarGarantiaInmobiliaria = False
    If bTrans Then
        obj.dRollbackTrans
        Set obj = Nothing
    End If
    Err.Raise Err.Number, "Ocurrió un error al actualizar la Garantía Inmobiliaria", Err.Description
End Function
'Public Sub CargarDatosPolizaInterna(ByVal psNumGarant As String, ByVal pdValorizacion As Date, ByVal psCtaCod As String, ByRef pRsPoliza As ADODB.Recordset)
'    Dim oGarantia As New COMDCredito.DCOMGarantia
'
'    On Error GoTo errRecuperaDatosPolizaInterna
'    Set pRsPoliza = oGarantia.RecuperaUltPolizaxGarantia(psNumGarant, psCtaCod)
'    'Set pRsGarantia = oGarantia.RecuperaDatosGarantiaxCalculoPoliza(psNumGarant, pdValorizacion)
'
'    Set oGarantia = Nothing
'    Exit Sub
'errRecuperaDatosPolizaInterna:
'    Err.Raise Err.Number, "Ocurrió un error al recuperar datos para calculo de la Poliza Interna", Err.Description
'End Sub
'END EJVG ********

'RECO20160419 ERS013-2016********************************************
Public Function ObtieneGarantLiq(ByVal psNumGarant As String) As Boolean
    Dim oGarant As New COMDCredito.DCOMGarantia
    Dim oRs As New ADODB.Recordset
    
    Set oRs = oGarant.ObtieneGarantLiq(psNumGarant)
    
    If Not (oRs.EOF And oRs.BOF) Then
        If oRs!nGarantLiq = 1 Then
            ObtieneGarantLiq = True
        Else
            ObtieneGarantLiq = False
        End If
    End If
End Function
'RECO FIN************************************************************
'CTI5 ERS0012020*****************************************************
Public Function ObtenertTpoDocumentoBienFuturo() As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenertTpoDocumentoBienFuturo = oGarant.ObtenerTpoBienFuturo
    Set oGarant = Nothing
End Function
Public Function ObtenerDocumentoIdentidadTitularGarantia(ByVal psPersCod As String) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenerDocumentoIdentidadTitularGarantia = oGarant.ObtenerDocumentoIdentidadTitularGarantia(psPersCod)
    Set oGarant = Nothing
End Function
Public Function ObtenertTpoDocumentoBienFuturoxGarantia(ByVal cNumGarant As String, ByVal cNumDoc As String, ByVal cTipo As String) As String
    Dim oGarant As New COMDCredito.DCOMGarantia
    Dim oRs As New ADODB.Recordset
    Set oRs = oGarant.ObtenertTpoDocumentoBienFuturoxGarantia(cNumGarant, cNumDoc)
    
    If Not (oRs.EOF And oRs.BOF) Then
     If cTipo = "Correlativo" Then
        ObtenertTpoDocumentoBienFuturoxGarantia = oRs!cCorrelativo
     ElseIf cTipo = "Etapa" Then
         ObtenertTpoDocumentoBienFuturoxGarantia = oRs!cEtapa
     End If
    End If
    
    Set oGarant = Nothing
End Function
Public Function ObtenerGarantiaCreditoDesemboldado(ByVal cNumGarant As String) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenerGarantiaCreditoDesemboldado = oGarant.ObtenerGarantiaCreditoDesemboldado(cNumGarant)
    Set oGarant = Nothing
End Function
Public Function ObtenerGarantiaBienFuturoCreditoHipo(ByVal cNumGarant As String) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenerGarantiaBienFuturoCreditoHipo = oGarant.ObtenerGarantiaBienFuturoCreditoHipo(cNumGarant)
    Set oGarant = Nothing
End Function
Public Function ObtenerGarantiaBienFuturo(ByVal nContratoId As Integer) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenerGarantiaBienFuturo = oGarant.ObtenerGarantiaBienFuturo(nContratoId)
    Set oGarant = Nothing
End Function
Public Function ObtenerGarantiaColocaciones(ByVal cNumGarant As String) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenerGarantiaColocaciones = oGarant.ObtenerGarantiaColocaciones(cNumGarant)
    Set oGarant = Nothing
End Function
Public Function ObtenerGarantiaColocacionesCobertura(ByVal cNumGarant As String) As ADODB.Recordset
    Dim oGarant As New COMDCredito.DCOMGarantia
    Set ObtenerGarantiaColocacionesCobertura = oGarant.ObtenerGarantiaColocacionesCobertura(cNumGarant)
    Set oGarant = Nothing
End Function
Public Function ObtenerCalendarioCredito(ByVal pcCtaCod As String, obj As COMDCredito.DCOMGarantia) As Variant
    Dim nIndex As Integer
    Dim oRsbj As ADODB.Recordset
    Set oRsbj = New ADODB.Recordset
        
    Set oRsbj = obj.ObtenerCalendarioMIVIENDA(pcCtaCod)
    nIndex = 0
    If Not oRsbj.BOF And Not oRsbj.EOF Then
        ReDim sCalendPagos(oRsbj.RecordCount, 3)
        Do While Not oRsbj.EOF
            sCalendPagos(nIndex, 0) = oRsbj!dVenc
            sCalendPagos(nIndex, 1) = oRsbj!nCuota
            sCalendPagos(nIndex, 2) = oRsbj!nMontoCuota
            nIndex = nIndex + 1
            oRsbj.MoveNext
        Loop
    End If
   ObtenerCalendarioCredito = sCalendPagos
End Function
'********************************************************************

