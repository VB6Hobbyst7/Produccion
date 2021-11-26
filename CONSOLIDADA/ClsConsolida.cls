VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsConsolida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim gConn As ADODB.Connection

Private Type TPlanPagos
    NumCuota As Integer
    FecVenc As Date
    Capital As Double
    CapPag As Double
    Interes As Double
    IntPag As Double
    IntGra As Double
    IntGraPag As Double
    MontoCuota As Double
    Gasto As Double
    Mora As Double
    MoraPag As Double
    Total As Double
    estado As String * 1
    Modificado As Boolean
    DiasAtr As Integer
End Type

Private Type TSCCuadre
    Descrip As String
    NumCred As Integer
    SaldoCap As Double
    IntDeveng As Double
    MontoGar1 As Double
    MontoGar2 As Double
    estado As String
    CapVenc As Double
    Abrev As String
End Type
Private Type TGastos
    Codgasto As String
    NumCuota As Integer
    MontoGasto As Double
    Gastopag As Double
    estado As String * 1
    Modificado As Boolean
End Type

Dim MatSaldos() As Double
Dim MatGastos() As TGastos
Dim MatPagos() As TPlanPagos
Dim nSaldos As Integer

Dim ContCuotas As Integer
Dim ContGastos As Integer
Dim TasaInt As Double
Dim Periodo As Integer
Dim TotalInteresaFecha As Double
Dim FecPagPend As Date
Dim FecUltPago As Date
Dim gdFecSis As Date

Public Event IniciaBarra(ByVal pnMaxValor As Long)

Public Event Progreso(ByVal i As Long, lsTitulo As String, lsSubtitulo As String)

Public Event finBarra()


'****************************************
'ReadParametros: Funcion que devuelve el valor de un parametro de
'la tabla parametros, se ingresa el codigo del parametro
'****************************************
'Private Function ReadParametros(txtCodPar As String) As Double
'
'    Dim RecVar As New ADODB.Recordset
'    Dim qryVar As String
'
'    On Error GoTo error
'
'    qryVar = "SELECT nValor1 FROM [128.107.2.2].DBCmact07.dbo.Parametro WHERE cCodPar = '" & txtCodPar & "' "
'
'    RecVar.Open qryVar, gConn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    If RecVar.EOF And RecVar.BOF Then
'        ReadParametros = 0
'        Set RecVar = Nothing
'    Else
'        ReadParametros = RecVar!nValor1
'        RecVar.Close
'        Set RecVar = Nothing
'    End If
'
'    Exit Function
'error:
'    ReadParametros = -1
'
'End Function

Public Sub Abreconexion(Optional pbMensajeOK As Boolean)
Dim sCadCon As String
'    sCadCon = "DSN=DSNConsolidada;UID=sa;PWD=confianza"

sCadCon = "DSN=DSNConsolidada;UID=sa;PWD=migrasa" 'CUSCO

    Set gConn = New ADODB.Connection
    gConn.Open sCadCon
    gConn.Execute "SET DATEFORMAT mdy"
    If pbMensajeOK Then
        If gConn.State <> adStateOpen Then
            MsgBox "No se pudo abrir conexion a Base Datos Consolidada", vbInformation, "Aviso"
        End If
    End If
End Sub

Public Sub CierraConexion()
    If gConn.State = 1 Then
        gConn.Close
    End If
End Sub

Function VerSiFeriado(ByVal FecVer As String) As Boolean
' Determina si una fecha es feriado
Dim RegVerFec As New ADODB.Recordset
Dim VSQL As String
    FecVer = Format(FecVer, "mm/dd/yyyy")
    VSQL = "select dFeriado from Feriado where dFeriado = '" & FecVer & "' "
    RegVerFec.Open VSQL, gConn, adOpenStatic, adLockOptimistic, adCmdText
    
    If RegVerFec.EOF And RegVerFec.BOF Then
        VerSiFeriado = False
    Else
        VerSiFeriado = True
    End If
    RegVerFec.Close
    Set RegVerFec = Nothing
End Function

Private Function fgCalculaInteresReprogramados(ByVal pCodCta As String, ByVal pConexion As ADODB.Connection) As Double
'** Halla los Intereses Reprogramados
Dim lsSQL As String
Dim r As New ADODB.Recordset
Dim lnCuotaFinal As Integer

   lsSQL = "SELECT nCuotasApr FROM CreditoConsol Where cCtaCod = '" & pCodCta & "'"
   r.Open lsSQL, pConexion, adOpenStatic, adLockReadOnly, adCmdText
   If Not IsNull(r!nCuotasApr) Then
         lnCuotaFinal = r!nCuotasApr
         r.Close
         fgCalculaInteresReprogramados = 0
         lsSQL = "Select * from PlandesPagConsol Where cCtaCod = '" & pCodCta & "' And nTipo=1 And cNroCuo > " & Str(Trim(lnCuotaFinal))
         r.Open lsSQL, pConexion, adOpenStatic, adLockReadOnly, adCmdText
            Do While Not r.EOF
               fgCalculaInteresReprogramados = fgCalculaInteresReprogramados + CDbl(Format(IIf(IsNull(r!nInteres), 0, r!nInteres) - IIf(IsNull(r!nintcompag), 0, r!nintcompag), "#0.00"))
               r.MoveNext
            Loop
        r.Close
        Set r = Nothing
   Else
       fgCalculaInteresReprogramados = 0
       r.Close
       Set r = Nothing
   End If
End Function


Private Function CobrarMora(ByVal Fecha As Date, ByVal param As Integer, ByVal FechaHoy As Date) As Boolean
Dim DiasAtraso, K As Integer
Dim fec As Date
    DiasAtraso = FechaHoy - Fecha
    If DiasAtraso <= param Then
        fec = FechaHoy
        For K = 1 To DiasAtraso
            fec = fec - 1
            If Not VerSiFeriado(Format(fec, "dd/mm/yyyy")) Then
                CobrarMora = True
                Exit Function
            End If
        Next K
        CobrarMora = False
    Else
        CobrarMora = True
    End If
End Function

Private Function CalculaInteresAdelantado(SaldoCapital As Double, TasaInteres As Double, Plazo As Integer) As Double
    CalculaInteresAdelantado = SaldoCapital * (1 - (1 / ((1 + TasaInteres) ^ (Plazo / 360))))
End Function

'********************************************
'** FUNCIONES DE CALCULOS DE CREDITOS  ******
'********************************************

Private Function fgCalculaInteresaFecha(ByVal pCodCta As String, ByVal pConexion As ADODB.Connection) As Double

Dim lsSQL As String
Dim r As New ADODB.Recordset
Dim ldFecUltDes As Date
Dim lnNroProxCuota As Integer
Dim lnDiasTrans As Integer
'Dim nInteres As Double
Dim lnGraciaApr As Integer
Dim ldFecVenGracia As Date
Dim lnMontoApr As Currency
Dim lnTasa As Double
Dim lnInteresFecha As Double
Dim lnPeriodo As Integer
Dim lnSaldoCap As Double
Dim ldFecVenCuota As Date
Dim lnIntPagFecha As Currency
Dim lbCreditoAdelantado As Boolean
Dim lnIntTemp As Currency
Dim lnCuotaCorresp As Integer
Dim ldFecCuotaCorresp As Date


'    On Error GoTo error
    'Obteniendo fecha de Desmbolso y Nro de Proxima Cuota
    lsSQL = "SELECT C.dfecVig, C.nTasaInt, C.nMontoApr, C.nGraciaApr, CS.nNroProxCuota, " _
        & " dFecUltDesemb = (select MAX(dFecPag) from PLandespagconsol where nEstado = 1 and nTipo=0 " _
        & " AND cctacod = CS.cCtaCod) "
    lsSQL = lsSQL & " FROM CreditoConsol C Inner Join CreditoSaldoConsol CS ON C.cCtaCod = CS.cCtaCod And " _
        & " DATEDIFF(dd, CS.dFecha, '" & Format(gdFecSis, "yyyy/mm/dd") & "') = 0 WHERE C.cCtaCod = '" & pCodCta & "'"
    r.Open lsSQL, pConexion, adOpenStatic, adLockReadOnly, adCmdText
        If Not r.BOF And Not r.EOF Then
            lnMontoApr = r!nMontoApr
            lnNroProxCuota = IIf(IsNull(r!nNroProxCuota), 0, r!nNroProxCuota)
            ldFecUltDes = Format(IIf(IsNull(r!dFecUltDesemb), r!dFecVig, r!dFecUltDesemb), "dd/mm/yyyy")
            lnGraciaApr = IIf(IsNull(r!nGraciaApr) Or r!nGraciaApr < 0, 0, r!nGraciaApr)
            If lnGraciaApr > 0 Then
                ldFecVenGracia = ldFecUltDes + lnGraciaApr
            End If
            lnTasa = r!nTasaInt
        End If
    r.Close
    
    lbCreditoAdelantado = False
    lnIntPagFecha = 0
    lnInteresFecha = 0
    If IsDate(ldFecUltDes) Then
        lnDiasTrans = (gdFecSis + 1) - ldFecUltDes
    Else
        lnDiasTrans = 0
    End If
    'Si es la primera cuota
    If lnNroProxCuota <= 1 And lnGraciaApr > 0 And lnDiasTrans < lnGraciaApr Then
        'Calculo de Dias Transcurridos y Si todavia no vence la gracia
         lnDiasTrans = CDate(Format(gdFecSis + 1, "dd/mm/yyyy")) - CDate(Format(ldFecUltDes, "dd/mm/yyyy"))
         lnInteresFecha = IntPerDias(lnTasa, lnDiasTrans, 30) * lnMontoApr
    Else
        'Halla intereses de Gracia Ganados
        lsSQL = "SELECT Sum(nIntGra - nIntGraPag) AS IntTot FROM PlandespagConsol WHERE cCtaCod = '" & pCodCta & "' And nTipo=1"
        r.Open lsSQL, pConexion, adOpenStatic, adLockReadOnly, adCmdText
            lnInteresFecha = lnInteresFecha + CDbl(Format(IIf(IsNull(r!IntTot), 0, r!IntTot), "#0.00"))
        r.Close
        
        lnSaldoCap = lnMontoApr
        If lnGraciaApr > 0 Then
            ldFecVenCuota = ldFecVenGracia
        Else
            ldFecVenCuota = ldFecUltDes
        End If
        
        lnCuotaCorresp = -1
        lsSQL = "SELECT * from PlandespagConsol WHERE cCtaCod = '" & pCodCta & "' And nTipo=1 order by cNroCuo"
        r.Open lsSQL, pConexion, adOpenStatic, adLockReadOnly, adCmdText
            Do While Not r.EOF
                If gdFecSis >= r!dFecVenc Then
                    ldFecVenCuota = Format(r!dFecVenc, "dd/mm/yyyy")
                    lnSaldoCap = lnSaldoCap - IIf(IsNull(r!nCapital), 0, r!nCapital)
                    'Capital mayor a cero para no tocar reprogramados ni intereses de gracia  a la primmera cuota o a la final
                    If r!nEstado = 0 And r!nCapital > 0 Then
                        lnInteresFecha = lnInteresFecha + (IIf(IsNull(r!nInteres), 0, r!nInteres) - IIf(IsNull(r!nintcompag), 0, r!nintcompag))
                    End If
                Else
                    lnCuotaCorresp = r!cNroCuo
                    If r!nEstado = 1 Then
                        lbCreditoAdelantado = True
                    Else
                        lbCreditoAdelantado = False
                    End If
                    ldFecCuotaCorresp = CDate(Format(r!dFecVenc, "dd/mm/yyyy"))
                    
                    'ARCV 26 - 7 - 2006
                    'lnIntPagFecha = IIf(IsNull(r!nintcompag), 0, IsNull(r!nintcompag))
                    lnIntPagFecha = IIf(IsNull(r!nintcompag), 0, r!nintcompag)
                    Exit Do
                End If
                r.MoveNext
            Loop
        r.Close
        
        If lnCuotaCorresp <> -1 And lbCreditoAdelantado = False Then
            '' Modificado x SGA le suma 2 dias de mas
            ''lnDiasTrans = (CDate(Format(gdFecSis, "dd/mm/yyyy")) + 1) - CDate(Format(ldFecVenCuota, "dd/mm/yyyy")) + 1
            lnDiasTrans = (CDate(Format(gdFecSis, "dd/mm/yyyy")) + 1) - (CDate(Format(ldFecVenCuota, "dd/mm/yyyy")) + 1)
            lnPeriodo = ldFecCuotaCorresp - ldFecVenCuota
            lnIntTemp = (IntPerDias(lnTasa, lnDiasTrans, 30) * lnSaldoCap) - lnIntPagFecha
            If lnIntTemp > 0 Then
               lnInteresFecha = lnInteresFecha + lnIntTemp
            End If
        End If
    End If
Set r = Nothing

fgCalculaInteresaFecha = CDbl(Format(lnInteresFecha, "#0.00"))
Exit Function

error:
    MsgBox Err.Description
    

End Function

Private Function TotalDeuda() As Double
Dim Monto, IntAPagar, MontoIAPagar As Double

Dim i, NumDias As Integer
    TotalInteresaFecha = 0
    Monto = 0
    For i = 0 To ContCuotas - 1
        'si es una cuota adelantada al pago pendiente
        
        If MatPagos(i).FecVenc > FecPagPend Then
            'Verificar si es un pago adelantado
            If gdFecSis < MatPagos(i).FecVenc Then
                If gdFecSis > MatPagos(i - 1).FecVenc Then
                   If MatPagos(i).Capital = 0 And MatPagos(i).NumCuota > 1 Then
                         Monto = Monto + MatPagos(i).Total
                         TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
                    Else
                        NumDias = gdFecSis - MatPagos(i - 1).FecVenc
                        'IntAPagar = IntPerDias(TasaInt, NumDias, Periodo) ' ARTURO
                        IntAPagar = IntPerDias(TasaInt, NumDias, IIf(Periodo = 0, 30, Periodo))
                        MontoIAPagar = MatSaldos(MatPagos(i).NumCuota - 1) * IntAPagar
                        
                        'Si monto de interes a pagar es menor que lo que ya pago entonces
                        'ese dinero del cliente que sobra se da por perdido
                        If MontoIAPagar < MatPagos(i).IntPag Then
                            MontoIAPagar = 0#
                            TotalInteresaFecha = TotalInteresaFecha + MontoIAPagar
                        Else
                            MontoIAPagar = MontoIAPagar - MatPagos(i).IntPag
                            
                        End If
                        Monto = Monto + MatPagos(i).Total - MatPagos(i).Interes + MontoIAPagar
                        MatPagos(i).Interes = MontoIAPagar
                        TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
                    End If
                Else
                    If MatPagos(i).Capital = 0 And MatPagos(i).NumCuota > 1 Then
                       Monto = Monto + MatPagos(i).Total
                       TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
                    Else
                        
                       Monto = Monto + MatPagos(i).Total - MatPagos(i).Interes
                       MatPagos(i).Interes = 0#
                       TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
                    End If
                End If
            Else
                Monto = Monto + MatPagos(i).Total
                TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
            End If
        Else
            'Preguntar si el Usuario esta al dia
            If gdFecSis < MatPagos(i).FecVenc Then
                NumDias = gdFecSis - CDate(Format(FecUltPago, "dd/mm/yyyy"))
                If NumDias < 0 Then
                   If MatPagos(i).Capital = 0 And MatPagos(i).NumCuota > 1 Then
                      Monto = Monto + MatPagos(i).Total
                      TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
                   Else
                      MontoIAPagar = 0
                    
                     'Si monto de interes a pagar es menor que lo que ya pago entonces
                     'ese dinero del cliente que sobra se da por perdido
                     If MontoIAPagar < MatPagos(i).IntPag Then
                         MontoIAPagar = 0#
                     Else
                         MontoIAPagar = MontoIAPagar - MatPagos(i).IntPag
                     End If
                     Monto = Monto + MatPagos(i).Total - MatPagos(i).Interes + MontoIAPagar
                     MatPagos(i).Interes = MontoIAPagar
                     TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
                  End If
                Else
                    If MatPagos(i).Capital = 0 And MatPagos(i).NumCuota > 1 Then
                       Monto = Monto + MatPagos(i).Total
                       TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
                    Else
                        IntAPagar = IntPerDias(TasaInt, NumDias, IIf(Periodo = 0, 30, Periodo))
                        MontoIAPagar = MatSaldos(MatPagos(i).NumCuota - 1) * IntAPagar
                        'Si monto de interes a pagar es menor que lo que ya pago entonces
                        'ese dinero del cliente que sobra se da por perdido
                        If MontoIAPagar < MatPagos(i).IntPag Then
                            MontoIAPagar = 0#
                        Else
                            MontoIAPagar = MontoIAPagar - MatPagos(i).IntPag
                        End If
                        Monto = Monto + MatPagos(i).Total - MatPagos(i).Interes + MontoIAPagar
                        MatPagos(i).Interes = MontoIAPagar
                        TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
                     End If
                End If
            Else
                Monto = Monto + MatPagos(i).Total
                TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
            End If
        End If
    Next i
    TotalDeuda = Monto
End Function

Private Function TotalDeudaParcial(Conexion As ADODB.Connection, pCuenta As String, Optional pTipoDesemb As String = "P") As Double
Dim sql1 As String
Dim RC As New ADODB.Recordset
Dim Interes As Double

    On Error GoTo error
    If pTipoDesemb = "P" And gdFecSis <= MatPagos(0).FecVenc Then
        sql1 = "SELECT * FROM PlandespagConsol WHERE cCtaCod='" & pCuenta & "' and nTipo = 0 and nEstado=1 ORDER BY cNroCuo ASC"
        RC.Open sql1, Conexion, adOpenStatic, adLockReadOnly, adCmdText
        Interes = 0
        Do While Not RC.EOF
            Interes = Interes + IntPerDias(TasaInt, (gdFecSis - RC!dFecVenc) + 1, Periodo) * RC!nCapital
            RC.MoveNext
        Loop
        RC.Close
        If Interes <= MatPagos(0).Interes Then
            Interes = Format(Interes, "#0.00")
        Else
            Interes = 0
        End If
        TotalInteresaFecha = Interes
        TotalDeudaParcial = MatPagos(0).Total - MatPagos(0).Interes + Interes
        MatPagos(0).Interes = Interes
    Else
        TotalInteresaFecha = 0
        TotalDeudaParcial = TotalDeuda
    End If
Set RC = Nothing
Exit Function

error:
    MsgBox Err.Description
    
End Function

Private Sub SaldoCapitalesCuota(ByVal CodCta As String, pConex As ADODB.Connection)
Dim RCred As New ADODB.Recordset
Dim sql1 As String
Dim i As Integer
Dim Monto As Double
    nSaldos = 1
    ReDim MatSaldos(nSaldos)
    sql1 = "SELECT CreditoConsol.nCuotasApr,CreditoConsol.nMontoDesemb,PlandespagConsol.nCapital FROM CreditoConsol CreditoConsol Inner Join PlandespagConsol PlandespagConsol" & _
        " On CreditoConsol.cCtaCod = PlandespagConsol.cCtaCod WHERE PlandespagConsol.nTipo=1 and CreditoConsol.cCtaCod='" & CodCta & "' ORDER BY PlandespagConsol.cNroCuo ASC"
    RCred.Open sql1, pConex, adOpenStatic, adLockReadOnly, adCmdText
    If Not RCred.BOF And Not RCred.EOF Then
        Monto = Format(IIf(IsNull(RCred!nMontoDesemb), 0, RCred!nMontoDesemb), "#0.00")
        MatSaldos(0) = IIf(IsNull(RCred!nMontoDesemb), 0, RCred!nMontoDesemb)
    End If
    Do While Not RCred.EOF
        Monto = Monto - Format(RCred!nCapital, "#0.00")
        nSaldos = nSaldos + 1
        ReDim Preserve MatSaldos(nSaldos)
        MatSaldos(nSaldos - 1) = Monto
        RCred.MoveNext
    Loop
    
    RCred.Close
    Set RCred = Nothing
End Sub

Private Function GastoCuota(NumC As Integer) As Double
Dim i As Integer
Dim Acum As Double
    Acum = 0
    For i = 0 To ContGastos - 1
        If NumC = MatGastos(i).NumCuota Then
            Acum = Acum + MatGastos(i).MontoGasto
        End If
    Next i
    GastoCuota = Acum
End Function

Private Sub CargaCuotas(ByVal CodCta As String, pConex As ADODB.Connection)
Dim RegCuotas  As New ADODB.Recordset
Dim sql1 As String
     ContCuotas = 0
    'Realiza Carga de Datos
    On Error GoTo error
    sql1 = "SELECT * FROM PlandespagConsol WHERE cCtaCod='" & CodCta & "' and nTipo = 1 and nEstado = 0" & _
           " ORDER BY cNroCuo"
    RegCuotas.Open sql1, pConex, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not RegCuotas.EOF
        If RegCuotas!nEstado = 0 Then
            ContCuotas = ContCuotas + 1
            ReDim Preserve MatPagos(ContCuotas)
            MatPagos(ContCuotas - 1).NumCuota = RegCuotas!cNroCuo
            MatPagos(ContCuotas - 1).FecVenc = RegCuotas!dFecVenc
            MatPagos(ContCuotas - 1).Capital = Format(RegCuotas!nCapital - IIf(IsNull(RegCuotas!nCapPag), 0, RegCuotas!nCapPag), "#0.00")
            MatPagos(ContCuotas - 1).Interes = Format(RegCuotas!nInteres - IIf(IsNull(RegCuotas!nintcompag), 0, RegCuotas!nintcompag), "#0.00")
            MatPagos(ContCuotas - 1).IntGra = Format(IIf(IsNull(RegCuotas!nIntGra), 0, RegCuotas!nIntGra) - IIf(IsNull(RegCuotas!nIntGraPag), 0, RegCuotas!nIntGraPag), "#0.00")
            MatPagos(ContCuotas - 1).MontoCuota = Format(MatPagos(ContCuotas - 1).Capital + MatPagos(ContCuotas - 1).Interes + MatPagos(ContCuotas - 1).IntGra, "#0.00")

            MatPagos(ContCuotas - 1).Gasto = Format(GastoCuota(RegCuotas!cNroCuo), "#0.00")
            MatPagos(ContCuotas - 1).Mora = Format(IIf(IsNull(RegCuotas!nMora), 0, CDbl(Format(IIf(IsNull(RegCuotas!nMora), 0, RegCuotas!nMora), "#0.00")) - CDbl(Format(IIf(IsNull(RegCuotas!nIntMorPag), 0, RegCuotas!nIntMorPag), "#0.00"))), "#0.00")
            
            If ContCuotas = 1 And MatPagos(ContCuotas - 1).Mora > 0 Then
                If Not CobrarMora(MatPagos(ContCuotas - 1).FecVenc, 3, gdFecSis) Then
                    MatPagos(ContCuotas - 1).Mora = 0
                End If
            End If
            MatPagos(ContCuotas - 1).Total = Format(MatPagos(ContCuotas - 1).MontoCuota + MatPagos(ContCuotas - 1).Gasto + MatPagos(ContCuotas - 1).Mora, "#0.00")
            MatPagos(ContCuotas - 1).CapPag = 0
            MatPagos(ContCuotas - 1).IntPag = Format(IsNull(RegCuotas!nintcompag), "#0.00")
            MatPagos(ContCuotas - 1).IntGraPag = 0
            MatPagos(ContCuotas - 1).MoraPag = 0
            MatPagos(ContCuotas - 1).estado = RegCuotas!nEstado
            MatPagos(ContCuotas - 1).Modificado = False
        End If
        RegCuotas.MoveNext
    Loop
    RegCuotas.Close
    FecPagPend = MatPagos(0).FecVenc
    'Cargar la Ultima Cuota
    sql1 = "SELECT * FROM PlandespagConsol WHERE cCtaCod='" & CodCta & "' and nTipo = 1 and nEstado = 1 " & _
           " ORDER BY cNroCuo DESC"
    RegCuotas.Open sql1, pConex, adOpenStatic, adLockReadOnly, adCmdText
    If Not RegCuotas.BOF And Not RegCuotas.EOF Then
        FecUltPago = RegCuotas!dFecVenc
    End If
    RegCuotas.Close
    'Carga los saldos de capitales de las cuotas
    Call SaldoCapitalesCuota(CodCta, pConex)
    
    Exit Sub
    
error:
    MsgBox Err.Description
    
End Sub


Private Function fgCalculaInteresaFechaCuotaLibre(ByVal pCodCta As String, ByVal pConexion As ADODB.Connection)

Dim r As New ADODB.Recordset
Dim lsSQL As String
Dim lnTasaInt As Double
Dim lnSaldoCapital As Double
Dim ldFecUltPago As Date
Dim ldFecUltCuoCal As Date
Dim lnDias As Integer
Dim lnInteresFecha As Double

    On Error GoTo error
    lsSQL = " SELECT C.nTasaInt, CS.nPrdEstado, " _
            & "dFecUltDesemb = (select MAX(dFecPag) from PLandespagconsol where nEstado = 1 and nTipo=0 and cctacod = CS.cCtaCod), " _
            & "C.dFecUltPago, C.dAsignacion, CS.nSaldoCap, C.dFecVig  " & _
            " FROM CreditoConsol C Inner Join CreditoSaldoConsol CS ON C.cCtaCod = CS.cCtaCod And Datediff(day,'" & Format(gdFecSis, "yyyy/mm/dd") & "',CS.dfecha)=0    Where CS.cCtaCod = '" & pCodCta _
            & "' And CS.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
    r.Open lsSQL, pConexion, adOpenStatic, adLockReadOnly, adCmdText
    
        lnTasaInt = r!nTasaInt
        lnSaldoCapital = r!nSaldoCap
    If r!nprdestado = 2020 Or r!nprdestado = 2021 Or r!nprdestado = 2022 Or r!nprdestado = 2030 Or r!nprdestado = 2031 Or r!nprdestado = 2032 Then
       If IsNull(r!dFecUltPago) Then
         ldFecUltPago = r!dFecUltDesemb
       Else
         ldFecUltPago = IIf(r!dFecUltDesemb > r!dFecUltPago, r!dFecUltDesemb, r!dFecUltPago)
       End If
    Else
        ldFecUltPago = IIf(IsNull(r!dAsignacion), r!dFecVig, r!dAsignacion)
    End If
    r.Close
    
    lsSQL = "SELECT dFecVenc FROM PlandespagConsol WHERE cCtaCod = '" & pCodCta & "' AND nTipo =1 ORDER BY dFecVenc DESC "
    r.Open lsSQL, pConexion, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not r.BOF And Not r.EOF Then
        ldFecUltCuoCal = r!dFecVenc
    Else
        ldFecUltCuoCal = gdFecSis
    End If
    r.Close
    Set r = Nothing
    
    If gdFecSis > ldFecUltCuoCal Then
        If ldFecUltPago < ldFecUltCuoCal Then
            lnDias = ldFecUltCuoCal - CDate(Format(ldFecUltPago, "dd/mm/yyyy"))
            lnInteresFecha = InteresReal(lnTasaInt / 100, lnDias) * lnSaldoCapital
        Else
            lnInteresFecha = 0
        End If
    Else
        lnDias = gdFecSis - CDate(Format(ldFecUltPago, "dd/mm/yyyy"))
        lnInteresFecha = InteresReal(lnTasaInt / 100, lnDias) * lnSaldoCapital
    End If
    lnInteresFecha = CDbl(Format(lnInteresFecha, "#0.00"))
    
    fgCalculaInteresaFechaCuotaLibre = lnInteresFecha
    Exit Function
    
error:
    MsgBox Err.Description
    
End Function

'********************************************************************
'*FUNCION QUE HALLA EL INTERES REAL DE ACUERDO A LOS DIAS DE PLAZO
'* Ejemplo1 : TI = 4.0 a 30 dias = 4.0
'* Ejemplo2 : TI = 4.0 a 40 dias = 5.16
'********************************************************************
Private Function InteresReal(ByVal Ti As Double, ByVal Periodo As Integer) As Double
    InteresReal = ((1 + Ti) ^ (Periodo / 30)) - 1
End Function

'***********************************************
'*  FUNCION QUE HALLA EL INTERES DE UN PERIODO DE DIAS TRANACURRIDOS
'***********************************************
Private Function IntPerDias(ByVal inter As Double, ByVal DiasTrans As Integer, ByVal Periodo As Double) As Double
    IntPerDias = ((1 + inter / 100) ^ (DiasTrans / Periodo)) - 1
End Function


Private Function ObtieneCapVencido(psCodCta As String, Psfecha As String, pConex As ADODB.Connection) As Double
Dim sSql As String
Dim Reg As New ADODB.Recordset
  
  sSql = "Select Sum(PlanDesPagConsol.nCapital - ISNULL(PlanDesPagConsol.nCapPag,0))  as CapVenc from PlanDesPagConsol " & _
          "where cCtaCod = '" & Trim(psCodCta) & "' and nTipo = 1 and nEstado = 0 " & _
          "and dFecVenc <= '" & Format(Psfecha, "mm/dd/yyyy") & " 23:59' "
          
  Reg.Open sSql, pConex, adOpenStatic, adLockReadOnly, adCmdText

  If Reg.EOF And Reg.BOF Then
      ObtieneCapVencido = 0
  Else
      ObtieneCapVencido = IIf(IsNull(Reg!CapVenc), 0, Reg!CapVenc)
  End If
  Reg.Close
  Set Reg = Nothing
End Function

Public Sub ActualizarCapitalVencido()
'**** Actualiza Cap Vencido
Dim rsCred As New ADODB.Recordset
Dim sSql As String
Dim lnCapVenc As Double
Dim dFecha As Date
Dim lsCmac As String

        Abreconexion
                
        sSql = "Select TOP 1 SUBSTRING(cCodAge,1,3) CMAC FROM SERVIDOR"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        lsCmac = rsCred!CMAC
        rsCred.Close
        
        sSql = "Select dfecConsol from varconsolida"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        dFecha = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close
        
        sSql = "SELECT cCtaCod, nDiasAtraso " & _
               "FROM Creditoconsol WHERE nPrdestado in (2020,2021,2022,2030,2031,2032)"
        
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        
        Do While Not rsCred.EOF
              lnCapVenc = ObtieneCapVencido(rsCred!cctacod, Format(dFecha, "dd/mm/yyyy"), gConn)
              sSql = "UPDATE Creditoconsol set nCapVencido = " & Round(lnCapVenc, 2) & _
                     " WHERE cCtaCod = '" & rsCred!cctacod & "'"
              gConn.Execute sSql
              
              sSql = "UPDATE CreditoSaldoConsol set nCapVencido = " & Round(lnCapVenc, 4) & _
                 " WHERE cCtaCod = '" & rsCred!cctacod & "' and datediff(dd, dFecha, '" & Format(dFecha, "yyyy/mm/dd") & "') = 0"
              gConn.Execute sSql
              rsCred.MoveNext
        Loop
    
        rsCred.Close
        Set rsCred = Nothing
        
        'If lsCmac = "102" Then  'Pignoraticio Lima - CAFF
        '    sSql = "SELECT cCtaCod, nDiasAtraso " & _
        '        "FROM CreditoConsol WHERE nPrdEstado IN (2101,2107,2104,2106)"
                
        '    rsCred.CursorLocation = adUseClient
        '    rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        '    Set rsCred.ActiveConnection = Nothing
            
        '    Do While Not rsCred.EOF
        '        If rsCred!nDiasAtraso >= 30 Then    'Aca Modificar esto para q se lea de alguna tabla
        '
        '            lnCapVenc = ObtieneCapVencido(rsCred!cctacod, Format(dFecha, "dd/mm/yyyy"), gConn)
        '            sSql = "UPDATE Creditoconsol set nCapVencido = " & Round(lnCapVenc, 2) & _
        '                   " WHERE cCtaCod = '" & rsCred!cctacod & "'"
        '            gConn.Execute sSql
        '
        '            sSql = "UPDATE CreditoSaldoConsol set nCapVencido = " & Round(lnCapVenc, 4) & _
        '               " WHERE cCtaCod = '" & rsCred!cctacod & "' and datediff(dd, dFecha, '" & Format(dFecha, "yyyy/mm/dd") & "') = 0"
        '
        '            gConn.Execute sSql
        '
        '        End If
        '        rsCred.MoveNext
        '    Loop
        'End If
        
        CierraConexion

End Sub

Public Sub ActualizaGarantiasCreditoDetallado()
Dim rsCred As New ADODB.Recordset
Dim rsGar  As ADODB.Recordset
Dim K As Integer, J As Integer
Dim lnTotal As Double

Dim sSql As String
Dim lsCodCred As String
Dim lsFFinanc As String
Dim lnIntDevengado As Double

Dim lnHipoSol As Double
Dim lnHipoDol As Double
Dim lnPrenSol As Double
Dim lnPrenDol As Double

    Abreconexion
    sSql = "Select dFecConsol from varconsolida"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        gdFecSis = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close

    sSql = "DELETE CredGarantiasConsol WHERE datediff(d,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
    gConn.Execute sSql

        'ARCV 04-07-2006
        'sSql = "SELECT cCtaCod, nTipoDesemb, dFecUltDesemb = (select MAX(dFecPag) from PLandespagconsol where nEstado = 1 and nTipo=0 and cctacod = C.cCtaCod), nPlazoApr, nTasaInt, SUBSTRING(cLineaCred,1,2) cFF " & _
               "FROM CreditoConsol C WHERE C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2202,2205,2206) "
        'ARCV 05-07-2006
        sSql = "SELECT cCtaCod, nTipoDesemb, dFecUltDesemb = (select MAX(dFecPag) from PLandespagconsol where nEstado = 1 and nTipo=0 and cctacod = C.cCtaCod), nPlazoApr, nTasaInt, " & _
               " cFF= (SELECT ISNULL(cCtaCont,'') FROM ColocLineaCreditoEquiv WHERE cLineaCred=SUBSTRING(C.cLineaCred,1,4)) " & _
               " FROM CreditoConsol C WHERE C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2202,2205,2206) "
                       
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        RaiseEvent IniciaBarra(rsCred.RecordCount)
        Set rsCred.ActiveConnection = Nothing
           Do While Not rsCred.EOF
               lsCodCred = rsCred!cctacod
               lsFFinanc = rsCred!cFF
               sSql = "SELECT G.nTipoGarant, G.nMoneda, SUM(GC.nMontoGrava) AS SumGarant " & _
                      " FROM GarantCredConsol GC INNER JOIN GarantiasConsol G ON GC.cNumGarant = G.cNumGarant " & _
                      " WHERE GC.cCtaCod= '" & lsCodCred & "' GROUP BY G.nTipoGarant, G.nMoneda "
               Set rsGar = New ADODB.Recordset
               rsGar.CursorLocation = adUseClient
               rsGar.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
               Set rsGar.ActiveConnection = Nothing
               
               Do While Not rsGar.EOF
                   
                       sSql = " INSERT INTO CredGarantiasConsol(dFecha, cCtaCod, cFF, nMoneda, cTipoGarant, nMonto) "
                       sSql = sSql & " VALUES('" & Format(gdFecSis, "mm/dd/yyyy") & "','" & lsCodCred & "','" & rsCred!cFF & "',"
                       sSql = sSql & rsGar!nMoneda & ",'" & rsGar!nTipoGarant & "'," & Format(rsGar!SumGarant, "#0.00") & ")"
                       
                       gConn.Execute sSql
                       
                  rsGar.MoveNext
               Loop
               rsGar.Close: Set rsGar = Nothing
               
               RaiseEvent Progreso(rsCred.Bookmark, "ACtualiza Garantias Credito Detallado", "")
              rsCred.MoveNext
           Loop
        
        RaiseEvent finBarra
        rsCred.Close
        Set rsCred = Nothing
        
        CierraConexion

End Sub

Public Sub ActualizaGarantiasJudicialDetallado()
Dim rsCred As New ADODB.Recordset
Dim rsGar  As ADODB.Recordset
Dim K As Integer, J As Integer
Dim lnTotal As Double

Dim sSql As String
Dim lsCodCred As String
Dim lsFFinanc As String
Dim lnIntDevengado As Double

Dim lnHipoSol As Double
Dim lnHipoDol As Double
Dim lnPrenSol As Double
Dim lnPrenDol As Double

    Abreconexion
    sSql = "Select dFecConsol from varconsolida"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        gdFecSis = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close

        sSql = "SELECT cCtaCod, nTipoDesemb, dFecUltDesemb = (select MAX(dFecPag) from PLandespagconsol where nEstado = 1 and nTipo=0 and cctacod = C.cCtaCod), nPlazoApr, nTasaInt, SUBSTRING(cLineaCred,1,2) cFF " & _
               "FROM CreditoConsol C WHERE C.nPrdEstado in (2201,2202)"
                       
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        RaiseEvent IniciaBarra(rsCred.RecordCount)
        Set rsCred.ActiveConnection = Nothing
           Do While Not rsCred.EOF
               lsCodCred = rsCred!cctacod
               lsFFinanc = rsCred!cFF
               sSql = "SELECT G.nTipoGarant, G.nMoneda, SUM(GC.nMontoGrava) AS SumGarant " & _
                      " FROM GarantCredConsol GC INNER JOIN GarantiasConsol G ON GC.cNumGarant = G.cNumGarant " & _
                      " WHERE GC.cCtaCod= '" & lsCodCred & "' GROUP BY G.nTipoGarant, G.nMoneda "
               Set rsGar = New ADODB.Recordset
               rsGar.CursorLocation = adUseClient
               rsGar.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
               Set rsGar.ActiveConnection = Nothing
               
               Do While Not rsGar.EOF
                   
                       sSql = " INSERT INTO CredGarantiasConsol(dFecha, cCtaCod, cFF, nMoneda, cTipoGarant, nMonto) "
                       sSql = sSql & " VALUES('" & Format(gdFecSis, "mm/dd/yyyy") & "','" & lsCodCred & "','" & rsCred!cFF & "',"
                       sSql = sSql & rsGar!nMoneda & ",'" & rsGar!nTipoGarant & "'," & Format(rsGar!SumGarant, "#0.00") & ")"
                       
                       gConn.Execute sSql
                       
                  rsGar.MoveNext
               Loop
               rsGar.Close: Set rsGar = Nothing
            
                RaiseEvent Progreso(rsCred.Bookmark, "Actualiza Garantias Jud.Detallada", "")
              rsCred.MoveNext
           Loop
        RaiseEvent finBarra
        
        rsCred.Close
        Set rsCred = Nothing
        
        CierraConexion

End Sub

Public Sub ActualizaInteresDevengado()
Dim rsCred As New ADODB.Recordset
Dim sSql As String
Dim lsCodCred As String
Dim lnDeuda As Currency
Dim lnIntDevengado As Currency
        'On Error GoTo error
        Abreconexion
        sSql = "select dfecConsol from varconsolida "
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        gdFecSis = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close

        sSql = "SELECT CS.cCtaCod, CS.nDiasAtraso, C.ntipCuota, C.nTipoDesemb, "
        sSql = sSql & " dFecUltDesemb = (select MAX(dFecPag) from PLandespagconsol where nEstado = 1 and nTipo=0 and cctacod = CS.cCtaCod), C.nPlazoApr, C.nTasaInt "
        'ARCV 12-06-2006
        sSql = sSql & ", nIntPend=ISNULL(C.nIntPend,0) "
        sSql = sSql & " FROM CreditoSaldoconsol CS Inner Join CreditoConsol C ON CS.cctacod = C.cctacod "
        sSql = sSql & " WHERE CS.nPrdEstado in (2020,2021,2022,2030,2031,2032) and DATEDIFF(DD, CS.dFecha,'" & Format(gdFecSis, "yyyy/mm/dd") & "') = 0 "
        sSql = sSql & " AND SUBSTRING(CS.cCtaCod,7,2) <> '21'"
        
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        
        RaiseEvent IniciaBarra(rsCred.RecordCount)
        
    Set rsCred.ActiveConnection = Nothing
        
        
       Do While Not rsCred.EOF
           lsCodCred = rsCred!cctacod
           
           If Not IsNull(rsCred!dFecUltDesemb) Then FecUltPago = rsCred!dFecUltDesemb
            Periodo = IIf(IIf(IsNull(rsCred!nPlazoApr), 0, rsCred!nPlazoApr) = 0, 30, rsCred!nPlazoApr)
            TasaInt = Format(InteresReal(Format(rsCred!nTasaInt / 100, "#0.000"), IIf(IIf(IsNull(rsCred!nPlazoApr), 0, rsCred!nPlazoApr) = 0, 30, rsCred!nPlazoApr)) * 100, "#0.00")
           
            'Calculo de la Deuda
            If rsCred!nTipCuota = 70 Then
               'ARCV 12-06-2006 --Se Agrego el nIntPend
               lnIntDevengado = fgCalculaInteresaFechaCuotaLibre(rsCred!cctacod, gConn) + rsCred!nIntPend
               
            Else
               If rsCred!nTipodesemb = 1 Then
                  Call CargaCuotas(rsCred!cctacod, gConn)
                  lnDeuda = TotalDeudaParcial(gConn, rsCred!cctacod)
                  lnIntDevengado = TotalInteresaFecha
                  
               Else
                  lnIntDevengado = fgCalculaInteresaFecha(rsCred!cctacod, gConn) + fgCalculaInteresReprogramados(rsCred!cctacod, gConn)
                  
               End If
            End If
          
          sSql = "UPDATE CreditoConsol set nIntDev = " & Round(lnIntDevengado, 4) & _
                 " WHERE cCtaCod = '" & rsCred!cctacod & "'"
          gConn.Execute sSql
          
          sSql = "UPDATE CreditoConsolTotal set nIntDev = " & Round(lnIntDevengado, 4) & " WHERE cCtaCod = '" & rsCred!cctacod & "'"
          gConn.Execute sSql
          
          sSql = "UPDATE CreditoSaldoConsol set nIntDev = " & Round(lnIntDevengado, 4) & _
                 " WHERE cCtaCod = '" & rsCred!cctacod & "' and datediff(dd, dFecha, '" & Format(gdFecSis, "yyyy/mm/dd") & "') = 0"
          gConn.Execute sSql
          
          RaiseEvent Progreso(rsCred.Bookmark, "Actualizando Interes devengado", "credito N° " & rsCred.RecordCount & " de " & rsCred.Bookmark)
           
          rsCred.MoveNext
          
       Loop
    
    rsCred.Close
    Set rsCred = Nothing
    RaiseEvent finBarra
    
    CierraConexion
    Exit Sub
    
error:
    MsgBox Err.Description
    
End Sub

Public Sub ActualizaInteresDevengadoPrendario()
Dim rsCred As New ADODB.Recordset
Dim sSql As String
Dim lsCodCred As String
Dim lnIntDevengado As Double
Dim pTasaInteresVencido As Double
Dim lnDiasAtraso As Integer
Dim vFecVen As String
Dim vDiaPro As Integer, vDiaAct As Integer
Dim vIntAde As Double, vIntAct As Double

    Abreconexion
    
    sSql = "Select dfecConsol from varconsolida"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        gdFecSis = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close

                sSql = "SELECT cCtaCod, dFecVenc, nSaldoCap, nTasaInt, nPlazoApr " & _
                       "FROM CreditoConsol WHERE nPrdEstado in (2101,2107,2104,2106) "
                
                rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
                    RaiseEvent IniciaBarra(rsCred.RecordCount)
                    
                   Do While Not rsCred.EOF
                       lsCodCred = rsCred!cctacod
                       pTasaInteresVencido = (rsCred!nTasaInt / 100)
                       lnDiasAtraso = 0
                       lnIntDevengado = 0
                       vDiaPro = 0: vDiaAct = 0
                       vIntAde = 0:    vIntAct = 0
                       vFecVen = Format(rsCred!dFecVenc, "dd/mm/yyyy")
                        If DateDiff("d", rsCred!dFecVenc, (gdFecSis + 1)) > 0 Then
                           lnIntDevengado = ((1 + pTasaInteresVencido) ^ (DateDiff("d", rsCred!dFecVenc, (gdFecSis + 1)) / 30) - 1) * rsCred!nSaldoCap
                           lnIntDevengado = Round(lnIntDevengado, 2)
                        End If
                        
                        lnDiasAtraso = IIf(DateDiff("d", vFecVen, (gdFecSis + 1)) > 0, DateDiff("d", vFecVen, (gdFecSis + 1)), 0)
                        If lnDiasAtraso = 0 Then
                             vDiaPro = IIf(DateDiff("d", gdFecSis, vFecVen) > 0, DateDiff("d", gdFecSis, vFecVen), 0)
                             vDiaAct = rsCred!nPlazoApr - vDiaPro
                             vIntAde = CalculaInteresAdelantado(rsCred!nSaldoCap, (rsCred!nTasaInt / 100), rsCred!nPlazoApr)
                             vIntAct = rsCred!nSaldoCap * ((((vIntAde / rsCred!nSaldoCap + 1) ^ (1 / rsCred!nPlazoApr)) ^ vDiaAct) - 1)
                             vIntAct = Round(vIntAct, 2)
                        End If
                        lnIntDevengado = lnIntDevengado + vIntAct
                          
                        sSql = "UPDATE CreditoConsol set nIntDev = " & Round(lnIntDevengado, 3) & _
                               " WHERE cCtaCod = '" & rsCred!cctacod & "'"
                        
                        gConn.Execute sSql
                        RaiseEvent Progreso(rsCred.Bookmark, "Interes Devengado Prendario", "Credito N° " & rsCred.Bookmark & " de " & rsCred.RecordCount)
                        rsCred.MoveNext
                   Loop
                
                
                rsCred.Close
                Set rsCred = Nothing
                RaiseEvent finBarra
                CierraConexion
End Sub

Public Sub ActualizaInteresdevengadoJudicial()
Dim rsCred As New ADODB.Recordset
Dim sSql As String
Dim lsCodCred As String
Dim lnIntDevengado As Double
Dim pTasaInteresVencido As Double

    Abreconexion
    
    sSql = "Select dfecConsol from varconsolida"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        gdFecSis = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close
                
    sSql = "SELECT cCodCta, nSaldIntCom, nMoraCalc " & _
           "FROM CreditoConsol WHERE cEstado = 'V' "
    rsCred.CursorLocation = adUseClient
    rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCred.ActiveConnection = Nothing
    Do While Not rsCred.EOF
       lsCodCred = rsCred!ccodCta
       
       lnIntDevengado = rsCred!nSaldIntCom + rsCred!nMoraCalc
       
       sSql = "UPDATE CreditoConsol SET nIntDev = " & Round(lnIntDevengado, 3) & _
              " WHERE cCodCta = '" & rsCred!ccodCta & "'"
       
       gConn.Execute sSql
       rsCred.MoveNext
    Loop
    
    rsCred.Close
    Set rsCred = Nothing
    CierraConexion

End Sub

Public Sub ActualizaGarantiasCredito()
Dim rsCred As New ADODB.Recordset
Dim rsGar  As ADODB.Recordset
Dim K As Integer, J As Integer
Dim lnTotal As Double

Dim sSql As String
Dim lsCodCred As String
Dim lsFFinanc As String
Dim lnIntDevengado As Double

Dim lnHipoSol As Double
Dim lnHipoDol As Double
Dim lnPrenSol As Double
Dim lnPrenDol As Double

    Abreconexion
    sSql = "Select dFecConsol from varconsolida"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        gdFecSis = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close

    sSql = "DELETE CredGarantias WHERE datediff(d,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
    gConn.Execute sSql
        
        'ARCV 04-07-2006
        'sSql = "SELECT cCtaCod, nTipoDesemb, dFecUltDesemb = (select MAX(dFecPag) from PLandespagconsol where nEstado = 1 and nTipo=0 and cctacod = C.cCtaCod), nPlazoApr, nTasaInt, SUBSTRING(cLineaCred,1,2) cFF " & _
               "FROM CreditoConsol C WHERE C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2202,2205,2206) "
        'ARCV 05-07-2006
        sSql = "SELECT cCtaCod, nTipoDesemb, dFecUltDesemb = (select MAX(dFecPag) from PLandespagconsol where nEstado = 1 and nTipo=0 and cctacod = C.cCtaCod), nPlazoApr, nTasaInt, " & _
               " cFF= (SELECT ISNULL(cCtaCont,'') FROM ColocLineaCreditoEquiv WHERE cLineaCred=SUBSTRING(C.cLineaCred,1,4)) " & _
               "FROM CreditoConsol C WHERE C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2202,2205,2206) "
                       
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        'rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        
        RaiseEvent IniciaBarra(rsCred.RecordCount)
        Set rsCred.ActiveConnection = Nothing
           Do While Not rsCred.EOF
               lsCodCred = rsCred!cctacod
               lsFFinanc = rsCred!cFF
               'sSql = "SELECT G.nTipoGarant, G.nMoneda, SUM(G.nMontoRealiz) AS SumGarant " & _
                      " FROM GarantCredConsol GC INNER JOIN GarantiasConsol G ON GC.cNumGarant = G.cNumGarant " & _
                      " WHERE GC.cCtaCod= '" & lsCodCred & "' GROUP BY G.nTipoGarant, G.nMoneda "
                '29-05-2006 ARCV(Se Agrego la Tabla GrupoGarantConsol)
                sSql = "SELECT G.nTipoGarant, G.nMoneda, SUM(G.nMontoRealiz) AS SumGarant,nGrupoGarant=ISNULL(GG.nGrupoGarant,10)  " & _
                      " FROM GarantCredConsol GC INNER JOIN GarantiasConsol G ON GC.cNumGarant = G.cNumGarant " & _
                      " LEFT JOIN GrupoGarantConsol GG ON G.nTipoGarant = GG.nTipoGarant " & _
                      " WHERE GC.cCtaCod= '" & lsCodCred & "' GROUP BY G.nTipoGarant, G.nMoneda ,GG.nGrupoGarant "
                '***********************
               Set rsGar = New ADODB.Recordset
               rsGar.CursorLocation = adUseClient
               rsGar.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
               Set rsGar.ActiveConnection = Nothing
               
               Dim aGarantias(2, 99) As Currency
               For J = 1 To 2
                For K = 1 To 99
                    aGarantias(J, K) = 0
                    'modificado x cjnm
                    aGarantias(0, K) = 0
                Next K
               Next J
               
               Do While Not rsGar.EOF
                'modificado x cjnm
                '29-05-2006 ARCV
                  If rsGar!nMoneda <> 3 Then
                    'aGarantias(rsGar!nMoneda, Val(rsGar!nTipoGarant)) = aGarantias(rsGar!nMoneda, Val(rsGar!nTipoGarant)) + rsGar!SumGarant
                    aGarantias(rsGar!nMoneda, Val(rsGar!nGrupoGarant)) = aGarantias(rsGar!nMoneda, Val(rsGar!nGrupoGarant)) + rsGar!SumGarant
                  Else
                    'aGarantias(0, Val(rsGar!nTipoGarant)) = aGarantias(0, Val(rsGar!nTipoGarant)) + rsGar!SumGarant
                    aGarantias(0, Val(rsGar!nGrupoGarant)) = aGarantias(0, Val(rsGar!nGrupoGarant)) + rsGar!SumGarant
                  End If
                  rsGar.MoveNext
               Loop
               rsGar.Close: Set rsGar = Nothing
               For J = 1 To 2
                  lnTotal = 0
                  For K = 1 To 12
                    lnTotal = lnTotal + aGarantias(J, K)
                    'modificado x cjnm
                    lnTotal = lnTotal + aGarantias(0, K)
                  Next
                  If lnTotal <> 0 Then
                    sSql = "SELECT cCtaCod FROM CredGarantias WHERE cCtaCod = '" & lsCodCred & "' and nMoneda = '" & J & "' and datediff(d,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
                    Set rsGar = New ADODB.Recordset
                    rsGar.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
                    If rsGar.EOF Then
                      sSql = "INSERT CredGarantias (dFecha, cCtaCod,cFF, nMoneda,nHipoteca,nVehicular, nIndustrial,nAgricola, nCartaFianza, nDeposito, nJoyas, nGlobalFlo, nFianzaSolida, nOtrosG, nPersonal, nAutorizacion ) VALUES ('" & Format(gdFecSis, "mm/dd/yyyy") & "', '" & lsCodCred & "', '" & lsFFinanc & "'," _
                           & "'" & J & "'," & aGarantias(J, 1) & "," & aGarantias(J, 2) & "," & aGarantias(J, 3) & "," & aGarantias(J, 4) & "," & aGarantias(J, 5) & "," & aGarantias(J, 6) & "," & aGarantias(J, 7) & "," & aGarantias(J, 8) & "," & aGarantias(J, 9) & "," & aGarantias(J, 10) & "," & aGarantias(J, 11) & "," & aGarantias(J, 12) & ")"
                    Else
                      sSql = "UPDATE CredGarantias SET cFF = '" & lsFFinanc & "', nHipoteca = " & aGarantias(J, 1) & ", nVehicular = " & aGarantias(J, 2) & ", nIndustrial = " & aGarantias(J, 3) & ",nAgricola = " & aGarantias(J, 4) & ", nCartaFianza = " & aGarantias(J, 5) & ", nDeposito = " & aGarantias(J, 6) _
                           & ", nJoyas = " & aGarantias(J, 7) & ", nGlobalFlo = " & aGarantias(J, 8) & ", nFianzaSolida = " & aGarantias(J, 9) & ", nOtrosG = " & aGarantias(J, 10) & ", nPersonal = " & aGarantias(J, 11) & ", nAutorizacion = " & aGarantias(J, 12) _
                           & " WHERE cCodCta='" & lsCodCred & "' and cMoneda = '" & J & "' and datediff(d,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
                    End If
                    rsGar.Close: Set rsGar = Nothing
                    gConn.Execute sSql
                  End If
              Next
              RaiseEvent Progreso(rsCred.Bookmark, "Actualizando Garantia de Credito", "Registro N°" & rsCred.Bookmark & " de " & rsCred.RecordCount)
              rsCred.MoveNext
           Loop
        
        RaiseEvent finBarra
        rsCred.Close
        Set rsCred = Nothing
        
        CierraConexion
End Sub

Public Sub ActualizaCredGarantiasPrendario()
Dim rsCred As New ADODB.Recordset
Dim rsGar As New ADODB.Recordset
Dim sSql As String
Dim lsCodCred As String
Dim lnIntDevengado As Double

Dim lnHipoSol As Double
Dim lnHipoDol As Double
Dim lnPrenSol As Double
Dim lnPrenDol As Double

    Abreconexion
    
    sSql = "Select dFecConsol from varconsolida"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        gdFecSis = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close
    
    sSql = "SELECT cCtaCod, nPrdEstado, nValTasac " & _
           "FROM CreditoConsol WHERE nPrdEstado in (2101,2107,2104,2106) "
    rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
    RaiseEvent IniciaBarra(rsCred.RecordCount)
    Do While Not rsCred.EOF

    lsCodCred = rsCred!cctacod
    
     sSql = "SELECT cCtaCod FROM CredGarantias WHERE cCtaCod = '" & lsCodCred & "' and nMoneda = 1 and datediff(d,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
     rsGar.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
     If rsGar.EOF Then
       sSql = "INSERT CredGarantias (dFecha,cCtaCod,cFF,nMoneda,nJoyas) VALUES ('" & Format(gdFecSis, "mm/dd/yyyy") & "','" & lsCodCred & "',1," _
            & "1," & Round(IIf(IsNull(rsCred!nValTasac), 0, rsCred!nValTasac), 4) & ")"
     Else
       sSql = "UPDATE CredGarantias SET nJoyas = " & Round(IIf(IsNull(rsCred!nValTasac), 0, rsCred!nValTasac), 4) & " WHERE cCtaCod='" & lsCodCred & "' and nMoneda = 1 and datediff(d,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0"
     End If
     rsGar.Close: Set rsGar = Nothing
     gConn.Execute sSql
        RaiseEvent Progreso(rsCred.Bookmark, "Actualiza Garantias de creditos prendario", "credito N° " & rsCred.RecordCount & " de " & rsCred.Bookmark)
    rsCred.MoveNext
    Loop
    
    RaiseEvent finBarra
    rsCred.Close
    Set rsCred = Nothing
    
    CierraConexion


End Sub

Public Sub ActualizaCredGarantiasPrendarioDetalle()
Dim rsCred As New ADODB.Recordset
Dim rsGar As New ADODB.Recordset
Dim sSql As String
Dim lsCodCred As String
Dim lnIntDevengado As Double

Dim lnHipoSol As Double
Dim lnHipoDol As Double
Dim lnPrenSol As Double
Dim lnPrenDol As Double

    Abreconexion
    
    sSql = "Select dFecConsol from varconsolida"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        gdFecSis = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close
    
    sSql = "SELECT cCtaCod, nPrdEstado, nValTasac " & _
           "FROM CreditoConsol WHERE nPrdEstado in (2101,2107,2104,2106) "
    rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
    RaiseEvent IniciaBarra(rsCred.RecordCount)
    Do While Not rsCred.EOF

    lsCodCred = rsCred!cctacod
    
        sSql = " INSERT INTO CredGarantiasConsol(dFecha, cCtaCod, cFF, nMoneda, cTipoGarant, nMonto) "
                       sSql = sSql & " VALUES('" & Format(gdFecSis, "mm/dd/yyyy") & "','" & lsCodCred & "','1',"
                       sSql = sSql & Mid(lsCodCred, 9, 1) & ",'7'," & Format(Round(IIf(IsNull(rsCred!nValTasac), 0, rsCred!nValTasac), 4), "#0.00") & ")"
                       
        gConn.Execute sSql
        RaiseEvent Progreso(rsCred.Bookmark, "ACtualiza Garantias Prendario Detalle", "")
    rsCred.MoveNext
    Loop
    RaiseEvent finBarra
    rsCred.Close
    Set rsCred = Nothing
    
    CierraConexion
End Sub

Public Sub ActualizaGarantiasJudicial()

Dim rsCred As New ADODB.Recordset
Dim rsGar  As ADODB.Recordset
Dim J As Integer, K As Integer
Dim lnTotal As Double

Dim sSql As String
Dim lsCodCred As String
Dim lsFFinanc As String
Dim lnIntDevengado As Double

Dim lnHipoSol As Double
Dim lnHipoDol As Double
Dim lnPrenSol As Double
Dim lnPrenDol As Double

    Abreconexion
    
    sSql = "Select dFecConsol from varconsolida"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        gdFecSis = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close
                
    sSql = "SELECT CJ.cCtaCod, CJ.nPrdEstado, Substring(cJ.cLineaCred,1,2) cFF " & _
           "FROM CreditoConsol CJ WHERE CJ.nPrdEstado in (2201,2202) "
    rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
       
       RaiseEvent IniciaBarra(rsCred.RecordCount)
       
       Do While Not rsCred.EOF

           lsCodCred = rsCred!cctacod
           lsFFinanc = rsCred!cFF
           sSql = "SELECT G.NTipoGarant, G.nMoneda, SUM(G.nMontoRealiz) AS SumGarant " & _
                  " FROM GARANTCREDCONSOL GC INNER JOIN GARANTIASCONSOL G ON GC.cNumGarant = G.cNumGarant " & _
                  " WHERE GC.cCtaCod = '" & lsCodCred & "' GROUP BY G.NTipoGarant, G.nMoneda "
           Set rsGar = New ADODB.Recordset
           rsGar.CursorLocation = adUseClient
           rsGar.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
           Set rsGar.ActiveConnection = Nothing

           Dim aGarantias(2, 50) As Currency
           For J = 1 To 2
            For K = 1 To 50
                aGarantias(J, K) = 0
            Next
           Next
           
           Do While Not rsGar.EOF
              aGarantias(Val(rsGar!nMoneda), Val(rsGar!nTipoGarant)) = aGarantias(Val(rsGar!nMoneda), Val(rsGar!nTipoGarant)) + rsGar!SumGarant
              rsGar.MoveNext
           Loop
           rsGar.Close: Set rsGar = Nothing
           For J = 1 To 2
              lnTotal = 0
              For K = 1 To 12
                lnTotal = lnTotal + aGarantias(J, K)
              Next
              If lnTotal <> 0 Then
                sSql = "SELECT cCtaCod FROM CredGarantias WHERE cCtaCod = '" & lsCodCred & "' and nMoneda = '" & J & "' and datediff(d,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
                Set rsGar = New ADODB.Recordset
                rsGar.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
                If rsGar.EOF Then
                  sSql = "INSERT CredGarantias (dFecha,cCtaCod, cFF, nMoneda,nHipoteca,nVehicular, nIndustrial,nAgricola, nCartaFianza, nDeposito, nJoyas, nGlobalFlo, nFianzaSolida, nOtrosG, nPersonal, nAutorizacion ) VALUES ('" & Format(gdFecSis, "mm/dd/yyyy") & "','" & lsCodCred & "','" & lsFFinanc & "', " _
                       & J & "," & aGarantias(J, 1) & "," & aGarantias(J, 2) & "," & aGarantias(J, 3) & "," & aGarantias(J, 4) & "," & aGarantias(J, 5) & "," & aGarantias(J, 6) & "," & aGarantias(J, 7) & "," & aGarantias(J, 8) & "," & aGarantias(J, 9) & "," & aGarantias(J, 10) & "," & aGarantias(J, 11) & "," & aGarantias(J, 12) & ")"
                Else
                  sSql = "UPDATE CredGarantias SET cFF = '" & lsFFinanc & "', nHipoteca = " & aGarantias(J, 1) & ", nVehicular = " & aGarantias(J, 2) & ", nIndustrial = " & aGarantias(J, 3) & ",nAgricola = " & aGarantias(J, 4) & ", nCartaFianza = " & aGarantias(J, 5) & ", nDeposito = " & aGarantias(J, 6) _
                       & ", nJoyas = " & aGarantias(J, 7) & ", nGlobalFlo = " & aGarantias(J, 8) & ", nFianzaSolida = " & aGarantias(J, 9) & ", nOtrosG = " & aGarantias(J, 10) & ", nPersonal = " & aGarantias(J, 11) & ", nAutorizacion = " & aGarantias(J, 12) _
                       & " WHERE cCtaCod='" & lsCodCred & "' and nMoneda = " & J & " and datediff(d,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
                End If
                rsGar.Close: Set rsGar = Nothing
                gConn.Execute sSql
              End If
           Next
          RaiseEvent Progreso(rsCred.Bookmark, "Actualiza garantias Judiciales", " registro N°" & rsCred.Bookmark & " de " & rsCred.RecordCount)
          rsCred.MoveNext
       Loop
    
    rsCred.Close
    Set rsCred = Nothing
     RaiseEvent finBarra
    CierraConexion
End Sub

Public Sub ActualizaEstadisticaMensualCredito()
Dim rsLC As New ADODB.Recordset
Dim rsRan As New ADODB.Recordset
Dim rsAg As New ADODB.Recordset
Dim rsTot As New ADODB.Recordset

Dim sqlY As String
Dim lcNorRef As String
Dim lnNum As Single
Dim lnCap As Currency
Dim lnCapVen As Currency
Dim i As Integer
Dim lnDatos As Single
Dim sSql As String
Dim rsCred As New ADODB.Recordset
    Abreconexion
    
    sSql = "Select dFecConsol from varconsolida"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        gdFecSis = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close


sqlY = " SELECT  cDescripcion as  cNomTab,IdAgencia as cValor FROM Agencias "
rsAg.Open sqlY, gConn, adOpenStatic, adLockReadOnly, adCmdText

sqlY = "SELECT nRanIniTab, nRanFinTab FROM Rangos " & _
   " WHERE substring(cCodTab,1,2) = 'A8' " & _
   " AND len(cCodTab)= 4 "
rsRan.Open sqlY, gConn, adOpenStatic, adLockReadOnly, adCmdText
   
    
sqlY = "SELECT cLineaCred FROM ColoclineaCredito "
rsLC.Open sqlY, gConn, adOpenStatic, adLockReadOnly, adCmdText
' Verifica si estan las tablas
    
' ****************************************
'Para Cada Linea de credito
rsLC.MoveFirst
Do While Not rsLC.EOF   ' Lineas Credito
   
   'Para Cada Agencia
   rsAg.MoveFirst
   Do While Not rsAg.EOF   'Agencia
       
        '===== Si hay datos en Agencia =====
        sqlY = "SELECT Count(cCtaCod) AS Num " & _
        " FROM CreditoConsol WHERE nPrdEstado in (2020,2021,2022,2030,2031,2032) AND cLineaCred = '" & rsLC!cLineaCred & "'" & _
        " AND substring(cCtaCod,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'"

        rsTot.Open sqlY, gConn, adOpenStatic, adLockReadOnly, adCmdText
        lnDatos = rsTot!Num
        rsTot.Close
        'Set rsTot = Nothing
    
        If lnDatos > 0 Then ' Si hay datos de la Agencia
           'Para Cada Rango
           rsRan.MoveFirst
           Do While Not rsRan.EOF  ' Rangos
           
              '===== Si hay datos en Rango =====
              sqlY = "SELECT Count(cCtaCod) AS Num " & _
              " FROM CreditoConsol WHERE nPrdEstado in (2020,2021,2022,2030,2031,2032) AND cLineaCred = '" & rsLC!cLineaCred & "'" & _
              " AND substring(cCtaCod,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
              " AND nDiasAtraso >= " & rsRan!nRanIniTab & _
              " AND nDiasAtraso <= " & rsRan!nRanFinTab
  
              rsTot.Open sqlY, gConn, adOpenStatic, adLockReadOnly, adCmdText
              lnDatos = rsTot!Num
              rsTot.Close
              'Set rsTot = Nothing
            
              If lnDatos > 0 Then ' Si hay datos del Rango
           
                 For i = 0 To 1  ' Norm / Refinan
                    lcNorRef = IIf(i = 0, "N", "R")
                    sqlY = "SELECT Count(cCtaCod) AS Num, Sum(nSaldoCap) AS Cap, " & _
                         " Sum(nCapVencido) AS CapVenc  FROM CreditoConsol  " & _
                         " WHERE nPrdEstado in (2020,2021,2022,2030,2031,2032)  AND cLineaCred = '" & rsLC!cLineaCred & "'" & _
                         " AND substring(cCtaCod,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                         " AND nDiasAtraso >= " & rsRan!nRanIniTab & _
                         " AND nDiasAtraso <= " & rsRan!nRanFinTab & _
                         " AND cRefinan = '" & lcNorRef & "'"
                     rsTot.Open sqlY, gConn, adOpenStatic, adLockReadOnly, adCmdText
                     lnNum = rsTot!Num
                     lnCap = IIf(IsNull(rsTot!Cap), 0, rsTot!Cap)
                     lnCapVen = IIf(IsNull(rsTot!CapVenc), 0, rsTot!CapVenc)
                     rsTot.Close
                     Set rsTot = Nothing
                    
                     ' Actualizo en la tabla
                     
                     If lnNum > 0 Then  ' **********
           
                        sqlY = "INSERT INTO EstadMensCartera (dFecha,cLineaCred, " & _
                               " nRangIni,nRangFin,cRefinan,nNumero,nCapital,nCapVenc,cAgencia, cClase) " & _
                               " VALUES ('" & Format(gdFecSis, "mm/dd/yyyy") & "','" & _
                               rsLC!cLineaCred & "'," & rsRan!nRanIniTab & "," & _
                               rsRan!nRanFinTab & ",'" & lcNorRef & "'," & lnNum & "," & _
                               lnCap & "," & lnCapVen & ",'" & Trim(rsAg!cValor) & "','CR')"
                               
                        gConn.Execute sqlY
                     
                     End If  ' **********************
                 
                 Next i  ' Norm / Refinan
        
              End If ' Si hay Datos Rango
                 
              rsRan.MoveNext
           Loop  ' Rangos
              
        End If ' Si hay Datos Agencia
          
        rsAg.MoveNext
                 
   Loop   ' Agencias
   rsLC.MoveNext
       
Loop  ' Lineas Credito
    
rsLC.Close
Set rsLC = Nothing
rsRan.Close
Set rsRan = Nothing
rsAg.Close
Set rsAg = Nothing

End Sub

Public Sub ActualizaEstadisticasMensualPrendario()
Dim rsLC As New ADODB.Recordset
Dim rsRan As New ADODB.Recordset
Dim rsAg As New ADODB.Recordset
Dim rsTot As New ADODB.Recordset

Dim sqlY As String
Dim lcNorRef As String
Dim lnNum As Single
Dim lnCap As Currency
Dim lnCapVen As Currency
Dim i As Integer
Dim lnDatos As Single
Dim sSql As String
Dim rsCred As New ADODB.Recordset
    Abreconexion
    
    sSql = "Select dFecConsol from varconsolida"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        gdFecSis = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close

sqlY = " SELECT  cDescripcion as  cNomTab,IdAgencia as cValor FROM Agencias "

rsAg.Open sqlY, gConn, adOpenStatic, adLockReadOnly, adCmdText

sqlY = "SELECT nRanIniTab, nRanFinTab FROM Rangos " & _
   " WHERE substring(cCodTab,1,2) = 'A8' " & _
   " AND len(cCodTab)= 4 "
rsRan.Open sqlY, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
   'Para Cada Agencia
   rsAg.MoveFirst
   Do While Not rsAg.EOF   'Agencia
           'Para Cada Rango
           rsRan.MoveFirst
           Do While Not rsRan.EOF  ' Rangos
                            
                    sqlY = "SELECT Count(cCtaCod) AS Num, Sum(nSaldoCap) AS Cap " & _
                         " FROM CreditoConsol   " & _
                         " WHERE nPrdEstado in (2101,2107,2104,2106) " & _
                         " AND substring(cCtaCod,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                         " And DATEDIFF(dd ,dFecVenc,'" & Format(gdFecSis, "mm/dd/yyyy") & "') >= " & rsRan!nRanIniTab & _
                         " And DATEDIFF(dd ,dFecVenc,'" & Format(gdFecSis, "mm/dd/yyyy") & "') <= " & rsRan!nRanFinTab
                     rsTot.Open sqlY, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
                     lnNum = rsTot!Num
                     lnCap = IIf(IsNull(rsTot!Cap), 0, rsTot!Cap)
                     'lnCapVen = IIf(IsNull(rsTot!CapVenc), 0, rsTot!CapVenc)
                     rsTot.Close
                     Set rsTot = Nothing
                    
                     ' Actualizo en la tabla
                     
                     If lnNum > 0 Then  ' **********
           
                        sqlY = "INSERT INTO EstadMensCartera (dFecha,cLineaCred, " & _
                               " nRangIni,nRangFin,cRefinan,nNumero,nCapital,nCapVenc,cAgencia, cClase) " & _
                               " VALUES ('" & Format(gdFecSis, "mm/dd/yyyy") & "','01011130501'," & _
                               rsRan!nRanIniTab & "," & _
                               rsRan!nRanFinTab & ",'N'," & lnNum & "," & _
                               lnCap & ",0,'" & Trim(rsAg!cValor) & "','CR')"
                               
                        gConn.Execute sqlY
                     
                     End If  ' **********************
                 
                 'Next i  ' Norm / Refinan
              rsRan.MoveNext
           Loop  ' Rangos
          
        rsAg.MoveNext
   Loop   ' Agencias
    
rsRan.Close
Set rsRan = Nothing
rsAg.Close
Set rsAg = Nothing

End Sub

Public Sub ActualizaEstadisticaMensualJudicial()

Dim rsLC As New ADODB.Recordset
Dim rsRan As New ADODB.Recordset
Dim rsAg As New ADODB.Recordset
Dim rsTot As New ADODB.Recordset

Dim sqlY As String
Dim lcNorRef As String
Dim lnNum As Single
Dim lnCap As Currency
Dim lnCapVen As Currency
Dim i As Integer
Dim lnDatos As Single
Dim rsCred As New ADODB.Recordset
Dim sSql As String

    Abreconexion
    
    sSql = "Select dFecConsol from varconsolida"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        gdFecSis = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close

sqlY = " SELECT  cDescripcion as  cNomTab,IdAgencia as cValor FROM Agencias "
rsAg.Open sqlY, gConn, adOpenStatic, adLockReadOnly, adCmdText
   
sqlY = "SELECT cLineaCred FROM ColoclineaCredito "
rsLC.Open sqlY, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
' ****************************************
'Para Cada Linea de credito
rsLC.MoveFirst
Do While Not rsLC.EOF   ' Lineas Credito
   
   'Para Cada Agencia
   rsAg.MoveFirst
   Do While Not rsAg.EOF   'Agencia
       
        '===== Si hay datos en Agencia =====
        sqlY = "SELECT Count(cCtaCod) AS Num " & _
        " FROM CreditoConsol WHERE nPrdEstado in (2201,2202)  AND cLineaCred = '" & rsLC!cLineaCred & "'" & _
        " AND substring(cCtaCod,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'"

        rsTot.Open sqlY, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
        lnDatos = rsTot!Num
        rsTot.Close
        Set rsTot = Nothing
    
        If lnDatos > 0 Then ' Si hay datos de la Agencia
                    
                    sqlY = "SELECT Count(cCtaCod) AS Num, Sum(nSaldoCap) AS Cap, " & _
                         " Sum(nSaldoCap) AS CapVenc  FROM Creditoconsol  " & _
                         " WHERE nPrdEstado in (2201,2202) AND cLineaCred = '" & rsLC!cLineaCred & "'" & _
                         " AND nSaldoCap > 0 " & _
                         " AND substring(cCtaCod,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'"
                     rsTot.Open sqlY, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
                     lnNum = rsTot!Num
                     lnCap = IIf(IsNull(rsTot!Cap), 0, rsTot!Cap)
                     lnCapVen = IIf(IsNull(rsTot!CapVenc), 0, rsTot!CapVenc)
                     rsTot.Close
                     Set rsTot = Nothing
                    
                     ' Actualizo en la tabla
                     
                     If lnNum > 0 Then  ' **********
           
                        sqlY = "INSERT INTO EstadMensCartera (dFecha,cLineaCred, " & _
                               " nRangIni,nRangFin,cRefinan,nNumero,nCapital,nCapVenc,cAgencia, cClase) " & _
                               " VALUES ('" & Format(gdFecSis, "mm/dd/yyyy") & "','" & _
                               rsLC!cLineaCred & "',999,999,'N'," & lnNum & "," & _
                               lnCap & "," & lnCapVen & ",'" & Trim(rsAg!cValor) & "','JU')"
                               
                        gConn.Execute sqlY
                     
                     End If  ' **********************
              
        End If ' Si hay Datos Agencia
          
        rsAg.MoveNext
   Loop   ' Agencias
   rsLC.MoveNext
       
Loop  ' Lineas Credito
    
rsLC.Close
Set rsLC = Nothing
rsAg.Close
Set rsAg = Nothing

End Sub

'=========================================================
'============  OBTIENE GARANTIAS      ====================
Private Function GetGarantias(lsCodCta As String, lsTipGar As String, lsTipoMon As String, pConex As ADODB.Connection) As Double
Dim Reg As New ADODB.Recordset
Dim sql As String

'sql = "SELECT SUM(GARANTIAS.nMontoRealiz) AS SumGarant " & _
'   " FROM CREDITOCONSOL CREDITO INNER JOIN GARANTCREDCONSOL GARANTCRED ON CREDITO.cCtaCod = " & _
'   " GARANTCRED.cCtaCod INNER JOIN GARANTIASCONSOL GARANTIAS " & _
'   " ON GARANTCRED.cNumGarant = GARANTIAS.cNumGarant " & _
'   " WHERE CREDITO.cCtaCod = '" & lsCodCta & _
'   "' AND GARANTIAS.cTipoGarant " & lsTipGar & _
'   " AND Garantias.nMoneda = '" & lsTipoMon & "'"
   
   
sql = "SELECT SUM(GARANTCRED.nMontoGrava) AS SumGarant " & _
        " FROM CREDITOCONSOL CREDITO INNER JOIN GARANTCREDCONSOL GARANTCRED ON CREDITO.cCtaCod = " & _
        " GARANTCRED.cCtaCod INNER JOIN GARANTIASCONSOL GARANTIAS " & _
        " ON GARANTCRED.cNumGarant = GARANTIAS.cNumGarant " & _
        " WHERE CREDITO.cCtaCod = '" & lsCodCta & _
        "' AND GARANTIAS.nTipoGarant " & lsTipGar & _
        " AND Garantias.nMoneda = '" & lsTipoMon & "'"
   
   sql = sql & " AND GARANTIAS.nTipoGarant<> 99 " 'ARCV 05-07-2006
   
Reg.Open sql, pConex, adOpenStatic, adLockReadOnly, adCmdText

If Reg.EOF And Reg.BOF Then
   GetGarantias = 0
Else
   GetGarantias = IIf(IsNull(Reg!SumGarant), 0, Reg!SumGarant)
End If
Reg.Close
Set Reg = Nothing
End Function

Private Function GetGarantiasJudicial(pCodCta As String, pTipoGar As String, pTipoMon As String, pConex As ADODB.Connection) As Double
Dim tmpReg As New ADODB.Recordset
Dim tmpSql As String

tmpSql = " SELECT SUM(CjGarantias.nMontoRealiz) as SumGarant " & _
    " FROM CjGarantCred INNER JOIN CjGarantias ON CjGarantCred.cNumGarant = " & _
    "CjGarantias.cNumGarant " & _
    " WHERE CjGarantCred.cCodCta = '" & pCodCta & _
    "' AND CjGarantias.cTipoGarant " & pTipoGar & _
    " AND CjGarantias.cMoneda = '" & pTipoMon & "'"
tmpReg.Open tmpSql, pConex, adOpenStatic, adLockReadOnly, adCmdText
If (tmpReg.BOF Or tmpReg.EOF) Then
    GetGarantiasJudicial = 0
Else
    GetGarantiasJudicial = IIf(IsNull(tmpReg!SumGarant), 0, tmpReg!SumGarant)
End If
tmpReg.Close
Set tmpReg = Nothing
End Function


Private Function RSVacio(rs1 As ADODB.Recordset) As Boolean
 RSVacio = (rs1.BOF And rs1.EOF)
End Function

Private Function VerificaGarantia(lsNumGarant As String) As Boolean
Dim sql As String
Dim rs As New ADODB.Recordset

sql = "Select * from Garantias where cNumGarant='" & lsNumGarant & "'"
rs.Open sql, gConn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RSVacio(rs) Then
    VerificaGarantia = True
Else
    VerificaGarantia = False
End If
rs.Close
Set rs = Nothing
End Function


Private Sub GarantiasPersonas(lsCodPers As String, dbConexion As ADODB.Connection)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim sql1 As String
sql = "Select * from Garantias where cCodPers='" & lsCodPers & "'"
rs.Open sql, dbConexion, adOpenStatic, adLockReadOnly, adCmdText
If Not RSVacio(rs) Then
    Do While Not rs.EOF
        If VerificaGarantia(rs!cNumGarant) = False Then
                sql1 = " INSERT INTO GARANTIAS (cNumGarant,cCodPers,cTipoGarant,cDocGarant,cCodZon,cMoneda,nMontoTasac,nMontoRealiz,nMontoxGrav,cEstado)" & _
                    " VALUES('" & rs!cNumGarant & "','" _
                    & rs!cCodPers & "','" & rs!cTipoGarant & "','" & rs!cDocGarant & "'," _
                    & "'" & rs!cCodZon & "','" & rs!cMoneda & "'," _
                    & rs!nMontoTasac & "," & rs!nMontoRealiz & "," & rs!nMontoxGrav & ",'" _
                    & rs!cEstado & "')"
                gConn.Execute sql1
        End If
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing
End Sub

Private Function ExistePersCredito1(lsCodCta As String, lsCodPers As String, lsRelacion As String, Optional lbCredito As Boolean = True) As Boolean
Dim sql As String
Dim rs As New ADODB.Recordset
If lbCredito = True Then
    sql = "Select * From PersCredito where ccodcta='" & lsCodCta & "' and ccodpers='" & lsCodPers & "' and crelacta='" & lsRelacion & "'"
Else
    sql = "Select * From PersCuenta where ccodcta='" & lsCodCta & "' and ccodpers='" & lsCodPers & "' and crelacta='" & lsRelacion & "'"
End If
rs.Open sql, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
If Not RSVacio(rs) Then
     ExistePersCredito1 = True
Else
     ExistePersCredito1 = False
End If
rs.Close
Set rs = Nothing

End Function


Private Sub PersCredito(lsCodCred As String, dbConexion As ADODB.Connection)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim sql1 As String
sql = "Select * from PersCredito where cCodCta='" & lsCodCred & "'"
rs.Open sql, dbConexion, adOpenStatic, adLockReadOnly, adCmdText
If Not RSVacio(rs) Then
    Do While Not rs.EOF
        If ExistePersCredito1(rs!ccodCta, rs!cCodPers, rs!cRelaCta) = False Then
            sql1 = "INSERT INTO PERSCREDITO VALUES('" & rs!cCodPers & "','" _
                & rs!ccodCta & "','" & rs!cRelaCta & "')"
        
            gConn.Execute sql1
            GarantiasPersonas rs!cCodPers, dbConexion
        End If
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing
End Sub

Public Sub ConsolidaCreditosConexion()
Dim sql As String
Dim rs As New ADODB.Recordset
Dim i As Long
Dim M As Long
Dim lsCadena As String
Dim lnCapitalVenc As Currency
Dim lnDeuda As Currency
Dim lnIntDevengado As Currency
Dim lnHipotecaSoles As Currency
Dim lnHipotecaDolares As Currency
Dim lnGarantiasSoles As Currency
Dim lnGarantiasDolares As Currency

lsCadena = ""
lsCadena = "'1',"
lsCadena = lsCadena & "'2',"
lsCadena = lsCadena & "'3','4',"
lsCadena = Mid(lsCadena, 1, Len(Trim(lsCadena)) - 1)
lsCadena = "WHERE SUBSTRING(cCtaCod,6,1) IN (" & lsCadena & ")"
            
            
            Abreconexion
            sql = "Select dFecConsol from varconsolida"
                rs.CursorLocation = adUseClient
                rs.Open sql, gConn, adOpenStatic, adLockReadOnly, adCmdText
                Set rs.ActiveConnection = Nothing
                gdFecSis = CDate(Format(rs!dfecConsol, "dd/mm/yyyy"))
                rs.Close
            
            sql = " SELECT *  " _
                & " FROM CreditoConsol " & lsCadena _
                & "  AND cCtaCod not in (SELECT cCtaCod FROM CreditoConsol WHERE nPrdEstado=2050 and dcancelado < DATEADD(MONTH , -3, '" & Format(gdFecSis, "mm/dd/yyyy") & "') ) "
            M = 0
            
            rs.CursorLocation = adUseClient
            rs.Open sql, gConn, adOpenStatic, adLockReadOnly, adCmdText
            rs.ActiveConnection = Nothing
            
            RaiseEvent IniciaBarra(rs.RecordCount)
            Do While Not rs.EOF
                M = M + 1
                lnCapitalVenc = rs!nCapVencido
                lnIntDevengado = rs!nIntDev
                'lnCapitalVenc = 0
                'lnIntDevengado = 0
                lnHipotecaSoles = 0
                lnHipotecaDolares = 0
                lnGarantiasSoles = 0
                lnGarantiasDolares = 0
                                        
                    If rs!nprdestado = 2020 Or rs!nprdestado = 2021 Or rs!nprdestado = 2022 Or rs!nprdestado = 2030 Or rs!nprdestado = 2031 Or rs!nprdestado = 2032 Or rs!nprdestado = 2201 _
                        Or rs!nprdestado = 2202 Or rs!nprdestado = 2205 Or rs!nprdestado = 2206 Then '    Then    (ARCV 06-07-2006)
'                        '=================== [ GARANTIA HIPOTECARIA SOLES ] ========================
'                        lnHipotecaSoles = GetGarantias(rs!cctacod, " IN (1, 19, 8) ", "1", gConn)
'                        '================= GARANTIA HIPOTECARIA DOLARES ==========================
'                        lnHipotecaDolares = GetGarantias(rs!cctacod, " IN (1, 19, 8) ", "2", gConn)
'                        '===================== [OTRAS GARANTIAS SOLES] =========================
'                        lnGarantiasSoles = GetGarantias(rs!cctacod, " NOT IN (1, 19, 8) ", "1", gConn)
'                        '===================== OTRAS GARANTIAS DOLARES ========================
'                        lnGarantiasDolares = GetGarantias(rs!cctacod, " NOT IN (1, 19, 8) ", "2", gConn)
'CUSCO
                        '=================== [ GARANTIA HIPOTECARIA SOLES ] ========================
                        lnHipotecaSoles = GetGarantias(rs!cctacod, " IN (1,2, 19, 30) ", "1", gConn)
                        '================= GARANTIA HIPOTECARIA DOLARES ==========================
                        lnHipotecaDolares = GetGarantias(rs!cctacod, " IN (1,2, 19, 30) ", "2", gConn)
                        '===================== [OTRAS GARANTIAS SOLES] =========================
                        lnGarantiasSoles = GetGarantias(rs!cctacod, " NOT IN (1,2, 19, 30) ", "1", gConn)
                        '===================== OTRAS GARANTIAS DOLARES ========================
                        lnGarantiasDolares = GetGarantias(rs!cctacod, " NOT IN (1,2, 19, 30) ", "2", gConn)
'FIN CUSCO
                   End If
                    
                    sql = "UPDATE CreditoConsolTotal set  " _
                        & "nHipoSoles=" & lnHipotecaSoles & ",nHipoDol=" & lnHipotecaDolares & ",nGarantSoles=" & lnGarantiasSoles & ",nGarantDol=" & lnGarantiasDolares & _
                        " WHERE cCtaCod='" & rs!cctacod & "'"
                    gConn.Execute sql
                    
                    sql = "UPDATE CreditoConsol set  " _
                        & "nHipoSoles=" & lnHipotecaSoles & ",nHipoDol=" & lnHipotecaDolares & ",nGarantSoles=" & lnGarantiasSoles & ",nGarantDol=" & lnGarantiasDolares & _
                        " WHERE cCtaCod='" & rs!cctacod & "'"
                    gConn.Execute sql
                    
                    'sql = "UPDATE CreditoConsolTotal SET nCapVencido = " & lnCapitalVenc & ", nIntDev = " & lnIntDevengado & _
                    '    " WHERE cCtaCod='" & rs!cctacod & "'"
                    'gConn.Execute sql
                    
                    RaiseEvent Progreso(rs.Bookmark, "Consolida Creditos", "")
                    
                    rs.MoveNext
            Loop
            RaiseEvent finBarra
            rs.Close
            Set rs = Nothing
        
            CierraConexion
    
End Sub

Public Function ObtieneFechaConsolida() As String
Dim sql As String
Dim rs As New ADODB.Recordset
            
    Abreconexion
    sql = "Select dFecConsol from varconsolida"
    rs.Open sql, gConn, adOpenStatic, adLockReadOnly, adCmdText
    gdFecSis = CDate(Format(rs!dfecConsol, "dd/mm/yyyy"))
    rs.Close
    CierraConexion
    ObtieneFechaConsolida = gdFecSis
End Function


Public Sub ActualizarEstadoRFA()
'**** Actualiza Estado RFA
Dim sSql As String
Dim dFecha As Date
Dim rsCred As New ADODB.Recordset


        Abreconexion

        sSql = "Select dfecConsol from varconsolida"
        rsCred.CursorLocation = adUseClient
        rsCred.Open sSql, gConn, adOpenStatic, adLockReadOnly, adCmdText
        Set rsCred.ActiveConnection = Nothing
        dFecha = CDate(Format(rsCred!dfecConsol, "dd/mm/yyyy"))
        rsCred.Close
        
        sSql = "Update creditoconsol set " & _
                " nprdestado = case when nprdestado = 2030 then 2060 " & _
                "                   when nprdestado = 2031 then 2061 " & _
                "                   when nprdestado = 2032 then 2062 " & _
                "                   when nprdestado = 2201 then 2262 " & _
                "                   when nprdestado = 2202 then 2262 " & _
                "                   when nprdestado = 2205 then 2265 " & _
                "                   when nprdestado = 2206 then 2262 end " & _
                " where nNroRepro = 50 "
              
        gConn.Execute sSql
                
        sSql = "Update creditoconsolTotal set " & _
                " nprdestado = case when nprdestado = 2030 then 2060 " & _
                "                   when nprdestado = 2031 then 2061 " & _
                "                   when nprdestado = 2032 then 2062 " & _
                "                   when nprdestado = 2201 then 2262 " & _
                "                   when nprdestado = 2202 then 2262 " & _
                "                   when nprdestado = 2205 then 2265 " & _
                "                   when nprdestado = 2206 then 2262 end " & _
                " where nNroRepro = 50 and nprdestado in (2030,2031,2032,2201,2202,2205,2206) "
              
        gConn.Execute sSql
                
        sSql = "Update creditoSaldoconsol set " & _
                " nprdestado = case when nprdestado = 2030 then 2060 " & _
                "                   when nprdestado = 2031 then 2061 " & _
                "                   when nprdestado = 2032 then 2062 " & _
                "                   when nprdestado = 2201 then 2262 " & _
                "                   when nprdestado = 2202 then 2262 " & _
                "                   when nprdestado = 2205 then 2265 " & _
                "                   when nprdestado = 2206 then 2262 end " & _
                " where cctaCod in (select cctacod from Creditoconsol where nNroRepro = 50 ) " & _
                " and datediff(dd, dFecha, '" & Format(dFecha, "yyyy/mm/dd") & "') = 0 "
              
        gConn.Execute sSql
                
        CierraConexion

End Sub
Public Sub CorrigeDescoberturaGarantia(ByVal lnTCFijoMes As Currency)
Dim sql As String
Abreconexion
sql = "EXEC CorrigeDescoberturaSoles " & lnTCFijoMes
gConn.Execute sql


sql = "EXEC CorrigeDescoberturaDolares " & lnTCFijoMes
gConn.Execute sql
CierraConexion
End Sub


