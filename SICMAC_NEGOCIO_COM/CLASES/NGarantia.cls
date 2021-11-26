VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private lsNegritaOn As String
Private lsNegritaOff  As String

Public Sub LiberaGarantia(ByVal psNumgarant As String, ByVal pnTipoGarant As Integer, _
    ByVal pnMontoLiberado As Double, ByVal pnMoneda As Integer, ByVal psAge As String, ByVal psUser As String, _
    ByVal pdFecha As Date, ByVal psCodCta As String, ByVal pnEstado As Integer)
Dim oBase As DCredActualizaBD
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
    Set oBase = New DCredActualizaBD
    oBase.dBeginTrans
    Call oBase.dUpdateGarantiaLibera(psNumgarant, pnEstado, pnMontoLiberado)
    Call oBase.dUpdateColocGarantia(psNumgarant, psCodCta, 0)
    
    If pnMontoLiberado > 0 Then
        sMovNro = oBase.GeneraMovNro(Format(pdFecha, "dd/mm/yyyy"), psAge, psUser)
        Call oBase.dInsertMov(sMovNro, sOpeCod, "Liberacion de Garantia Manual", gMovEstContabMovContable, gMovFlagVigente, False)
        nMovNro = oBase.GetnMovNro(sMovNro)
        Call oBase.dInsertMovCol(nMovNro, sOpeCod, psCodCta, 0, pnMontoLiberado, 0, "", 0, 0, 0)
        Call oBase.dInsertMovColDet(nMovNro, sOpeCod, psCodCta, 0, gColocConceptoCodCapital, 0, pnMontoLiberado)
    End If
    oBase.dCommitTrans
    
    Dim odGar As DGarantia
    Dim sSql As String
    Set odGar = New DGarantia
    Set R = odGar.RecuperaGarantia(psNumgarant)
    Set odGar = Nothing
    
    If R!nGravament = 0 Then 'Liberara
        sSql = " UPDATE Garantias Set nEstado = 4 Where cNumGarant = '" & psNumgarant & "'"
    Else 'Contabilizada
        sSql = " UPDATE Garantias Set nEstado = 3 Where cNumGarant = '" & psNumgarant & "'"
    End If
    Call oBase.coConex.Ejecutar(sSql)
    
    Set oBase = Nothing
End Sub


Public Sub BloqueaGarantia(ByVal psNumgarant As String)
Dim oBase As DCredActualizaBD
    
    Set oBase = New DCredActualizaBD
    Call oBase.dUpdateGarantiaBloquea(psNumgarant, gPersGarantEstadoBloqueada)
    Set oBase = Nothing
End Sub

Public Sub DesbloqueaGarantia(ByVal psNumgarant As String)
Dim oBase As DCredActualizaBD
Dim odGar As DGarantia
Dim bGarVig As Boolean

    Set odGar = New DGarantia
    bGarVig = odGar.PerteneceACreditoVigente(psNumgarant)
    Set oBase = New DCredActualizaBD
    If bGarVig Then
        Call oBase.dUpdateGarantiaBloquea(psNumgarant, gPersGarantEstadoContabilizado)
    Else
        bGarVig = odGar.PerteneceACredito(psNumgarant)
        If Not bGarVig Then
            Call oBase.dUpdateGarantiaBloquea(psNumgarant, gPersGarantEstadoAsignado)
        Else
            Call oBase.dUpdateGarantiaBloquea(psNumgarant, gPersGarantEstadoLiberado)
        End If
    End If
    
    Set odGar = Nothing

End Sub

Public Function PorcentajeGarantia(ByVal psConsGarantia As String) As Double
Dim oConecta As DConecta
Dim sSql As String
Dim R As ADODB.Recordset

    On Error GoTo ErrorPorcentajeGarantia
    sSql = "Select nParamValor from ColocParametro Where nParamVar = " & psConsGarantia
    Set oConecta = New DConecta
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
Dim I As Integer
Dim nNumTit As Integer
Dim nPorc As Double

     ValidaDatos = ""
    'Valida Numero de Titulares
    If Not pbSoloGaran Then
        nNumTit = 0
        For I = 0 To UBound(RelPers) - 1
            If CInt(Trim(Right(RelPers(I), 10))) = gPersRelGarantiaTitular Then
                nNumTit = nNumTit + 1
            End If
        Next I
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
Dim oConecta As DConecta

    On Error GoTo ErrorRecupera
     
sSql = "Select PG.cPersCod as CodPers, P.cPersNombre as Nombre From GarantReal GR"
sSql = sSql & " Inner Join PersGarantia PG on PG.cNumGarant=GR.cNumGarant"
sSql = sSql & " Inner Join Persona P on P.cPersCod=Pg.cPersCod"
sSql = sSql & " Where GR.cNumGarant =" & Nrog
sSql = sSql & " And PG.nRelacion = 1"

    Set oConecta = New DConecta
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

Dim lsCadena As String
'lsCadena = Chr$(10)
lsCadena = Chr$(10) & Chr(10) & lsNegritaOn
lsCadena = lsCadena & ImpreFormat("CMACT - CREDITOS ", 43, , False) & Space(6) & ImpreFormat("CMACT - CREDITOS ", 43, , False) & Chr(10)
lsCadena = lsCadena & ImpreFormat(psNomAge, 43, , False) & Space(6) & ImpreFormat(psNomAge, 43, , False) & Chr(10) & lsNegritaOff
lsCadena = lsCadena & ImpreFormat("Fecha:" & psFecha & Space(5) & "Hora:" & psHora, 43, , False) & Space(6) & ImpreFormat("Fecha:" & psFecha & Space(5) & "Hora:" & psHora, 43, , False) & Chr(10)
lsCadena = lsCadena & lsNegritaOn & ImpreFormat(PstaNombre(psPersNombre), 43, , False) & Space(6) & ImpreFormat(PstaNombre(psPersNombre), 43, , False) & Chr(10)
lsCadena = lsCadena & ImpreFormat("----------EXTORNO DE LEVANTAMIENTO----------", 43, , False) & Space(6) & ImpreFormat("----------EXTORNO DE LEVANTAMIENTO----------", 43, , False) & Chr(10) & lsNegritaOff
ImprimeCabeceraGarantiaRealExtorno = lsCadena
End Function

Public Function ImprimeBoletaGarantiaExt(ByVal psNomAge As String, _
        ByVal psFecha As String, ByVal psHora As String, _
        ByVal Nrog As String, ByVal psCodUser As String, ByVal pnMovG As String) As String
Dim Cabecera As String
Dim cuerpo As String
Dim loConecta As DConecta
Dim Dg As DGarantia
Dim Rs As ADODB.Recordset
Dim sSql As String
Dim Total As Double
Dim Lineas As Integer
Dim llena As String
Dim I As Integer
Dim MovG As Long
Lineas = 7
Set Dg = New DGarantia
Set Rs = New ADODB.Recordset
Set Rs = RecuperaTitularGarantReal(Nrog)

Cabecera = ImprimeCabeceraGarantiaRealExtorno(psNomAge, psFecha, psHora, Rs!Nombre)
Total = 0

Set loConecta = New DConecta
sSql = "select nMovNro, cOpeCod from Mov Where cMovNro='" & pnMovG & "'"
        loConecta.AbreConexion
Set Rs = loConecta.CargaRecordSet(sSql)
MovG = IIf(IsNull(Rs!nMovNro), 0, Rs!nMovNro)

'Detalle
sSql = " Select( Select cDescripcion from ProductoConcepto where nPrdConceptoCod=MGD.nPrdConceptoCod) Concepto,"
sSql = sSql & " nMonto"
sSql = sSql & " from MovGarantDet MGD"
sSql = sSql & " Where MGD.nMovNro =" & MovG

Set Rs = loConecta.CargaRecordSet(sSql)
loConecta.CierraConexion
  
cuerpo = Cabecera
If Rs.BOF And Rs.EOF Then

Else

    cuerpo = cuerpo & Chr(10)
    While Not Rs.EOF
    cuerpo = cuerpo & ImpreFormat(Rs!concepto, 23, 0) & Space(8) & ImpreFormat(Rs!nMonto, 10, 2, True) & Space(6) & ImpreFormat(Rs!concepto, 23, 0, False) & Space(10) & ImpreFormat(Rs!nMonto, 10, 2, True) & Chr(10)
    Lineas = Lineas + 1
    Total = Total + Rs!nMonto
    Rs.MoveNext
    Wend
    
End If
cuerpo = cuerpo & lsNegritaOn & String(45, "-") & Space(6) & String(45, "-") & Chr(10)
cuerpo = cuerpo & ImpreFormat("Total Gastos", 23, 0) & Space(8) & ImpreFormat(Total, 10, 2, True) & Space(8) & ImpreFormat("Total Gastos", 23, 0, False) & Space(8) & ImpreFormat(Total, 10, 2, True) & Chr(10)
cuerpo = cuerpo & psCodUser & Space(48) & psCodUser & lsNegritaOff
Lineas = Lineas + 3

If Lineas <= 18 Then
    For I = Lineas To 18
        cuerpo = cuerpo + Chr(10)
        
    Next I
End If
Set Rs = Nothing
Set loConecta = Nothing
Set Dg = Nothing
ImprimeBoletaGarantiaExt = cuerpo
End Function


Private Sub Class_Initialize()
 lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)
End Sub
