VERSION 5.00
Begin VB.Form frmAnx15AEstadDia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte Estadistico Diario"
   ClientHeight    =   975
   ClientLeft      =   4260
   ClientTop       =   4470
   ClientWidth     =   4815
   Icon            =   "frmAnx15AEstadDia.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmAnx15AEstadDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsSaldos() As String
Dim lsOtros() As String
Dim lnPorcEncSoles As Currency


Dim lnMontoPFCMACT As Currency
Dim lnObligacionesCaja As Currency
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim lnChequeCartera As Currency
Dim lnNumChequesCartera As Long
Dim lnMontoChqCust As Currency
Dim lnNumChequesCust As Long

Dim lnMontoCmacs As Currency
Dim lnMontoChequesVal As Currency
Dim lnTotalJudiciales As Integer
Dim lnSaldosJudiciales As Currency

Dim lnMontoBCR As Currency
Dim lnMontoCaja As Currency
Dim lnMontoCajaDiaAnt As Currency
Dim lnMontoCajaDia    As Currency

Dim lnMontoCartaFianza As Currency
Dim lnTotalCartaFianza As Currency


Dim lnToseBase As Double
Dim lsArchivo As String
Dim lsNomArch As String

Dim lnEncajeBase() As Currency
Dim lnTotalCaptaciones() As Currency
Dim lnFondoExigible() As Currency
Dim lnMontoChequePF As Currency
Dim lnSaldoTotalCTS As Currency

Dim lsRutaReferencia As String

'Saldos de C.Rurales
Dim nSaldoCRACPF As Currency
Dim nSaldoCRACAC As Currency
Dim lbExcel As Boolean

Dim oCon As DConecta
Dim oBarra As clsProgressBar

Public Sub ImprimeEstadisticaDiaria(psOpeCod As String, psMoneda As String, pdFecha As Date)
Dim lbHojaActiva As Boolean
On Error GoTo ImprimeEstadisticaDiariaErr
   GeneraEstadisticaDiaria psOpeCod, psMoneda, pdFecha
   lsNomArch = "Anx15A_EstCaja_" & Format(pdFecha, "mmyyyy") & IIf(psMoneda = "1", "MN", "ME") & ".XLS"
   lsArchivo = App.path & "\SPOOLER\" & lsNomArch
   lbHojaActiva = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)

   If lbHojaActiva Then
      ExcelAddHoja Format(pdFecha, "dd-mm-yyyy"), xlLibro, xlHoja1
      GeneraReporte psOpeCod, psMoneda, pdFecha
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
      If lsArchivo <> "" Then
          MsgBox "Anexo Generado satisfactoriamente", vbInformation, "Aviso!!!"
          CargaArchivo lsArchivo, App.path & "\SPOOLER\"
      End If
   End If
Exit Sub
ImprimeEstadisticaDiariaErr:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
    If lbHojaActiva Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub


Private Sub GeneraEstadisticaDiaria(psOpeCod As String, psMoneda As String, pdFecha As Date)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim I As Integer
Dim J As Integer
Dim lsNomAge As String * 20
Dim lsCad As String
Dim lsMoneda As String
Dim Total, m As Long

Dim N1 As New nCajaGenImprimir


ReDim lsSaldos(12, 1)

ReDim lnEncajeBase(1)
ReDim lnTotalCaptaciones(1)
ReDim lnFondoExigible(1)

I = 0: J = 0
'lnMontoCmacs = 0
lnMontoPFCMACT = 0
lnMontoChequesVal = 0

lnChequeCartera = 0
lnNumChequesCartera = 0
lnMontoChqCust = 0
lnNumChequesCust = 0

lnMontoBCR = 0
lnMontoCaja = 0
'lnPPD = 0

lnMontoCartaFianza = 0
lnTotalCartaFianza = 0
'lnCredHipotecarios = 0

lnSaldoTotalCTS = 0
lnMontoChequePF = 0
 
Dim oAge As New DActualizaDatosArea
Set rs = oAge.GetAgencias(, False)
Set oAge = Nothing
lsMoneda = psMoneda
If Not RSVacio(rs) Then

   lnMontoCaja = SaldosPromedioCaja("1", lsMoneda, pdFecha)
   lnMontoCajaDia = SaldosCuentas(psOpeCod, lsMoneda, "1", pdFecha)
   lnMontoCajaDiaAnt = SaldosCuentas(psOpeCod, lsMoneda, "1", pdFecha - 1)
   lnMontoBCR = SaldosCuentas(psOpeCod, lsMoneda, "2", pdFecha)
   
   lnObligacionesCaja = N1.SaldoObligInmediatas(1, "76" & lsMoneda & "201", pdFecha, pdFecha, "99")
   
   SaldosChequeCartera pdFecha, lsMoneda
   CartasFianzas psOpeCod, pdFecha, lsMoneda
   
   Do While Not rs.EOF
      m = m + 1
      lsNomAge = rs!Descripcion
      I = I + 1
'      If CInt(rs!Codigo) = 7 Then
'         CreditosSanta lsMoneda
'         CreditosJudiciales lsMoneda
'      End If
      ReDim Preserve lnEncajeBase(I)
      ReDim Preserve lnTotalCaptaciones(I)
      ReDim Preserve lsSaldos(12, I)
     
      lnEncajeBase(I) = EncajeBaseAgencia(gsCodCMAC & rs!Codigo, lsMoneda, pdFecha)
      lsSaldos(1, I) = Trim(lsNomAge)   'Nombre de Agencia
      If rs!Codigo = "01" Then
        ' Saldos de ahorro cte
         lsSaldos(2, I) = Format(SaldosAhorros(gCapAhorros, lsMoneda, pdFecha, rs!Codigo), gsFormatoNumeroView)
      Else
        ' Saldos de ahorro cte
         lsSaldos(2, I) = Format(SaldosAhorros(gCapAhorros, lsMoneda, pdFecha, rs!Codigo), gsFormatoNumeroView)
      End If
       ' saldos de P.F. + CTS
      lsSaldos(3, I) = Format(SaldosAhorros(gCapPlazoFijo, lsMoneda, pdFecha, rs!Codigo) + SaldosAhorros(gCapCTS, lsMoneda, pdFecha, rs!Codigo, False), gsFormatoNumeroView)
      
      'Agregado Pp
      
      'Interes Antiguo
      'nMontoInteresesPF = nMontoInteresesPF + Format(GetInteresPF(psOpeCod, lsMoneda, rs!Codigo, pdFecha, False), gsFormatoNumeroView)
      
      'GetInteresPF psOpeCod, lsMoneda, rs!Codigo, pdFecha, True
      
      lsSaldos(4, I) = 0
      'lsSaldos(4, I) = EncajeGetInteresEntidades(pdFecha, lsMoneda, Right(rs!Codigo, 2))
      lsSaldos(4, I) = N1.SaldoObligInmediatas(1, "76" & lsMoneda & "201", pdFecha, pdFecha, rs!Codigo)
         
      'lsSaldosProm(4, I) = Format(GetInteresPFProm(psOpeCod, lsMoneda, rs!Codigo, pdFecha), gsFormatoNumeroView)
      
      'Fin Agregado Pp
      
      'lsSaldos(4, I) = Format(GetInteresPF(psOpeCod, lsMoneda, rs!Codigo, pdFecha), gsFormatoNumeroView)
      
      lnSaldoTotalCTS = lnSaldoTotalCTS + SaldosAhorros(gCapCTS, lsMoneda, pdFecha, rs!Codigo)
      
      If I = 1 Then
        lsSaldos(5, 0) = lnObligacionesCaja
      End If
      lsSaldos(5, I) = Format(CCur(lsSaldos(2, I)) + CCur(lsSaldos(3, I)) + CCur(lsSaldos(4, I)), "#,#0.00")     'Total de Captaciones
      lnTotalCaptaciones(I) = Format(CCur(lsSaldos(2, I)) + CCur(lsSaldos(3, I)) + CCur(lsSaldos(4, I)), "#,#0.00")  'Total de Captaciones
      
      lsSaldos(6, I) = Format(CajasPromedioAgencias(lsMoneda, pdFecha, rs!Codigo), "#,#0.00")              'Caja de Agencias
      lsSaldos(7, I) = Format(CajasAgencias(psOpeCod, lsMoneda, pdFecha - 1, rs!Codigo), "#,#0.00")        'Caja de Agencias
      lsSaldos(8, I) = Format(CajasAgencias(psOpeCod, lsMoneda, pdFecha, rs!Codigo), "#,#0.00")                   'Caja de Agencias
      
      If lsMoneda = "1" Then
          lsSaldos(9, I) = lsSaldos(6, I)                                     'fondos de encaje Caja +bcr
      Else
          lsSaldos(9, I) = lsSaldos(8, I)                                     'Saldo de Caja Dia +bcr
      End If
        
      If lsMoneda = "1" Then
          lsSaldos(10, I) = Format(CCur(lsSaldos(9, I)) * 0.07, "#,#0.00")                    'fondos exigible Soles 0.07 del fondo de encaje
      Else
          '***************************************
          'calculo de encaje para dolares
          lsSaldos(10, I) = Format(CCur(lsSaldos(9, I)) * 0.07, "#,#0.00")
          
      End If
      lsSaldos(11, I) = Format(CCur(lsSaldos(9, I)) - CCur(lsSaldos(10, I)), "#,#0.00") 'SUPERAVIT DEFICIT DIARIO A-B
      'Variación de Caja
      lsSaldos(12, I) = Format(CCur(lsSaldos(8, I)) - CCur(lsSaldos(7, I)), "#,#0.00")

     '************* Datos de prendario **********************
      'ReDim Preserve lsPrendario(4, I)
      'ReDim Preserve lsAdjudicados(4, I)
     
      'lnOroVig = 0: lnOroAdj = 0: lnNumVig = 0: lnNumAdj = 0: lnCapVig = 0: lnCapAdj = 0
     
'      lnMontoCmacs = lnMontoCmacs + SaldosCmact(pdFecha, lsMoneda, rs!Codigo, True)
     
      lnMontoPFCMACT = lnMontoPFCMACT + SaldosCmact(pdFecha, lsMoneda, rs!Codigo, False)
   
      lnMontoChequesVal = lnMontoChequesVal + ChequesValorizacion(pdFecha, lsMoneda, rs!Codigo)
     
      lnMontoChequePF = lnMontoChequePF + ChequesValorizacion(pdFecha, lsMoneda, rs!Codigo, False)
     
''''      If lsMoneda = "1" Then
''''         SaldosPrendario pdFecha, lsMoneda, rs!Codigo
''''         lsPrendario(1, I) = lsNomAge
''''         lsPrendario(2, I) = Format(lnCapVig, "#,#0.00")
''''         lsPrendario(3, I) = Format(lnNumVig, "#,#0")
''''         lsPrendario(4, I) = Format(lnOroVig, "#,#0.00")
''''
''''         lsAdjudicados(1, I) = "Adjudicados Ag." & rs!Codigo
''''         lsAdjudicados(2, I) = Format(lnCapAdj, "#,#0.00")
''''         lsAdjudicados(3, I) = Format(lnNumAdj, "#,#0")
''''         lsAdjudicados(4, I) = Format(lnOroAdj, "#,#0.00")
''''      End If
''''      '*************** Créditos Personales **********************
''''      SaldosCreditos lsMoneda, Trim(lsNomAge), rs!Codigo, pdFecha, I
''''      '**********************************************************
      rs.MoveNext
   Loop
   CalculoFondoExigible lsMoneda, pdFecha
End If

RSClose rs

Set N1 = Nothing

End Sub

'Private Sub CreditosSanta(lsMoneda As String)
'Dim sql As String
'Dim rs As New ADODB.Recordset
'Dim oConR As New DConecta
'
''Modificado por Pepe Para Trujillo y lima
'
'If gbBitCentral = True Then
'
'    If oConR.AbreConexion() Then
'       lnCreditosSanta(1) = "CPE.Santa"
'       sql = "SELECT  SUM(p.nSaldo) AS Saldo, COUNT(*) AS TOTAL " _
'        & "FROM  ColocRecup cr JOIN Producto p ON p.cCtaCod = cr.cCtaCod " _
'        & "WHERE SubString(cr.cCtaCod,4,2)='99' and substring(cr.cCtaCod,9,1)='" & lsMoneda & "' " _
'        & "      and nPrdEstado = " & gColocEstRecVigJud
'
'       Set rs = oConR.CargaRecordSet(sql)
'
'       If Not RSVacio(rs) Then
'          lnCreditosSanta(2) = Format(IIf(IsNull(rs!Saldo), 0, rs!Saldo), gsFormatoNumeroView)
'          lnCreditosSanta(3) = Format(IIf(IsNull(rs!Total), 0, rs!Total), gsFormatoNumeroView)
'       End If
'    End If
'Else
'    If oConR.AbreConexionRemota("07", False, False, "01") Then
'       lnCreditosSanta(1) = "CPE.Santa"
'        sql = "SELECT  SUM(nSaldCap) AS SALDO, COUNT(*) AS TOTAL " _
'            & "From CredcJudi " _
'            & "WHERE   SubString(ccodcta,1,2)='99' and substring(cCodCta,6,1)='" & lsMoneda & "'" _
'            & "and cEstado='V' and cCondicion='J'"
'
'       Set rs = oConR.CargaRecordSet(sql)
'       If Not RSVacio(rs) Then
'          lnCreditosSanta(2) = Format(IIf(IsNull(rs!Saldo), 0, rs!Saldo), gsFormatoNumeroView)
'          lnCreditosSanta(3) = Format(IIf(IsNull(rs!Total), 0, rs!Total), gsFormatoNumeroView)
'       End If
'    End If
'End If
'oConR.CierraConexion
'Set oConR = Nothing
'End Sub

'Private Sub CreditosJudiciales(lsMoneda As String)
'Dim sql As String
'Dim rs As New ADODB.Recordset
'Dim oConR As New DConecta
'
'lnTotalJudiciales = 0
'lnSaldosJudiciales = 0
'
'     If gbBitCentral = True Then
'        If oConR.AbreConexion Then
'            sql = " SELECT  SUM(p.nSaldo) AS Saldo, COUNT(*) AS TotalCredJud " _
'                & " FROM    ColocRecup cr JOIN Producto p ON p.cCtaCod = cr.cCtaCod " _
'                & " WHERE   nSaldo > 0 and " _
'                & "         NOT SubString(cr.cCtaCod,4,2) = '99' and substring(cr.cCtaCod,9,1)='" & lsMoneda & "' and " _
'                & "         nPrdEstado = " & gColocEstRecVigJud
'        Else
'            Exit Sub
'        End If
'
'    Else
'        If oConR.AbreConexionRemota("07", False, False, "01") Then
'
'            sql = " SELECT  ISNULL(SUM(nSaldCap),0) as Saldo , Count(*) TotalCredJud " _
'            & " From CREDCJUDI " _
'            & " WHERE   cEstado ='V' AND cCondicion ='J'  and nSaldCap> 0 " _
'            & "         AND Substring(cCodCta,6,1)='" & lsMoneda & "' and SubString(ccodcta,1,2) not in ('99') "
'
'        Else
'            Exit Sub
'        End If
'    End If
'
'    Set rs = oConR.CargaRecordSet(sql)
'    If Not RSVacio(rs) Then
'       lnTotalJudiciales = rs!TotalCredJud
'       lnSaldosJudiciales = IIf(IsNull(rs!Saldo), 0, rs!Saldo)
'    End If
'    RSClose rs
'
'oConR.CierraConexion
'Set oConR = Nothing
'End Sub
'
'Private Sub SaldosPrendario(ldFecha As Date, psMoneda As String, lsCodAge As String)
'Dim sql As String
'Dim rs As New ADODB.Recordset
'Dim o As New DConecta
'Dim oEst As NEstadisticas
'   Set oEst = New NEstadisticas
'   Set rs = oEst.GetEstadisticaPrendario(gbBitCentral, ldFecha, psMoneda, lsCodAge)
'   Do While Not rs.EOF
'      lnOroVig = lnOroVig + rs!nOroVig
'      lnOroAdj = lnOroAdj + rs!nOroAdj
'      lnNumVig = lnNumVig + rs!nNumCredVig
'      lnNumAdj = lnNumAdj + rs!nNumCredAdj
'      lnCapVig = lnCapVig + rs!nCapVig
'      lnCapAdj = lnCapAdj + rs!nCapAdj
'      rs.MoveNext
'   Loop
'   RSClose rs
'End Sub

Function CajasAgencias(psOpeCod As String, lsMoneda As String, ldFecha As Date, lsCodAge As String) As Currency
Dim oEfe As New Defectivo
CajasAgencias = oEfe.BilletajeCajaAgencias("", lsMoneda, ldFecha, lsCodAge)
Set oEfe = Nothing
End Function

Function SaldosAhorros(pProd As String, Moneda As String, ldFecha As Date, lsCodAge As String, Optional pbIncluyeChqValoriza As Boolean = True) As Currency
Dim rs As New ADODB.Recordset
Dim lNumChq As Long
Dim lSalChq As Currency
Dim SaldoAge As Currency
Dim SQLBil As String
Dim lnPFCmact As Currency

Dim oEst As New NEstadisticas
Select Case pProd
    Case gCapAhorros
         Set rs = oEst.GetEstadisticaAhorro(gbBitCentral, ldFecha, ldFecha, Moneda, lsCodAge)
    Case gCapPlazoFijo
         Set rs = oEst.GetEstadisticaPlazoFijo(gbBitCentral, ldFecha, ldFecha, Moneda, lsCodAge)
    Case gCapCTS
         Set rs = oEst.GetEstadisticaCTS(gbBitCentral, ldFecha, ldFecha, Moneda, lsCodAge)
End Select
SaldoAge = 0
If Not RSVacio(rs) Then
   Do While Not rs.EOF
      Select Case pProd
         Case gCapAhorros
            If gbBitCentral = True Then
                SaldoAge = SaldoAge + rs!nSaldoSinCMAC '+ rs!nSaldCoop
                'SaldoAge = SaldoAge + rs!nSaldoSinCMAC + rs!nSaldoCMAC + rs!nSaldCRAC  '- rs!nChqCMAC - rs!nChqCrAC
            Else
                SaldoAge = SaldoAge + rs!nSaldoAC - rs!nSaldoCMAC - rs!nChqCMAC - rs!nSaldCRAC
            End If
            nSaldoCRACAC = nSaldoCRACAC + rs!nSaldCRAC
            '-----------------------------
         Case gCapPlazoFijo
            If gbBitCentral = True Then
                SaldoAge = SaldoAge + rs!nSaldoSinCMAC
                'SaldoAge = SaldoAge + rs!nSaldoSinCMAC + rs!nSaldoCMAC + rs!nSaldCRAC
            Else
                SaldoAge = SaldoAge + rs!nSaldoPF - rs!nSaldoCMAC - rs!nSaldCRAC
            End If
            nSaldoCRACPF = nSaldoCRACPF + rs!nSaldCRAC
         Case gCapCTS
            If gbBitCentral = True Then
                If pbIncluyeChqValoriza Then
                    SaldoAge = SaldoAge + rs!nSaldoSinCMAC + rs!MontoChq
                    'SaldoAge = SaldoAge + rs!nSaldoSinCMAC + rs!nSaldoCMAC + rs!nSaldCRAC + rs!MontoChq
                Else
                    SaldoAge = SaldoAge + rs!nSaldoSinCMAC
                End If
            Else
                If pbIncluyeChqValoriza Then
                    SaldoAge = SaldoAge + rs!nSaldo + rs!MontoChq - rs!nSaldoCMAC - rs!nSaldCRAC
                Else
                    SaldoAge = SaldoAge + rs!nSaldo - rs!nSaldoCMAC - rs!nSaldCRAC
                End If
            End If
      End Select
      rs.MoveNext
   Loop
End If
RSClose rs
SaldosAhorros = SaldoAge
End Function

Private Function SaldosCmact(ldFecha As Date, lsMoneda As String, lsCodAge As String, Optional lbAhorros As Boolean = True) As Currency
Dim sql As String
Dim rs As New ADODB.Recordset
Dim oEst As New NEstadisticas
If lbAhorros Then
   Set rs = oEst.GetEstadisticaAhorro(gbBitCentral, ldFecha, ldFecha, lsMoneda, lsCodAge)
Else
   Set rs = oEst.GetEstadisticaPlazoFijo(gbBitCentral, ldFecha, ldFecha, lsMoneda, lsCodAge)
End If
SaldosCmact = 0
If Not RSVacio(rs) Then
    Do While Not rs.EOF
        SaldosCmact = SaldosCmact + IIf(IsNull(rs!nSaldoCMAC + rs!nChqCMAC), 0, rs!nSaldoCMAC + rs!nChqCMAC)
        rs.MoveNext
    Loop
End If
RSClose rs
End Function

Private Function ChequesValorizacion(ldFecha As Date, lsMoneda As String, lsCodAge As String, Optional lbChqVal As Boolean = True) As Currency
Dim sql As String
Dim rs As New ADODB.Recordset
Dim oEst As New NEstadisticas
If lbChqVal = True Then
   Set rs = oEst.GetEstadisticaAhorro(gbBitCentral, ldFecha, ldFecha, lsMoneda, lsCodAge)
Else
   Set rs = oEst.GetEstadisticaPlazoFijoCts(gbBitCentral, ldFecha, ldFecha, lsMoneda, lsCodAge)
End If
ChequesValorizacion = 0
Do While Not rs.EOF
   ChequesValorizacion = ChequesValorizacion + IIf(IsNull(rs!nMonChqVal), 0, rs!nMonChqVal)
   rs.MoveNext
Loop
RSClose rs
End Function

Private Function SaldosCuentas(psOpeCod As String, lsMoneda As String, lsOrden As String, ldFecha As Date) As Currency
Dim sql As String
Dim rs As New ADODB.Recordset
Dim oSdo As New NCtasaldo
Dim oOpe As New DOperacion
Dim lsCtaCod As String
Dim lnSaldo  As Currency
BarraShow 1
BarraProgress 0, "REPORTE ESTADISTICO DIARIO", "", "Procesando datos...", vbBlue
Set rs = oOpe.CargaOpeCtaUltimoNivel(psOpeCod, "D", lsOrden)
SaldosCuentas = 0
BarraClose
BarraShow rs.RecordCount
Do While Not rs.EOF
   lsCtaCod = rs!cCtaContCod
   If lsMoneda <> Mid(rs!cCtaContCod, 3, 1) Then
      lsCtaCod = Left(rs!cCtaContCod, 2) & lsMoneda & Mid(rs!cCtaContCod, 4, 22)
   End If
   BarraProgress rs.Bookmark, "REPORTE ESTADISTICO DIARIO", "", "Procesando datos...", vbBlue
   If lsMoneda = "1" Then
       lnSaldo = oSdo.GetCtaSaldo(lsCtaCod, Format(ldFecha, gsFormatoFecha))
   Else
       lnSaldo = oSdo.GetCtaSaldo(lsCtaCod, Format(ldFecha, gsFormatoFecha), False)
   End If
   SaldosCuentas = SaldosCuentas + lnSaldo
   rs.MoveNext

Loop
BarraClose
RSClose rs
Set oSdo = Nothing
Set oOpe = Nothing
End Function

Private Sub SaldosChequeCartera(ldFecha As Date, lsMoneda As String)
Dim oChq As NDocRec
Set oChq = New NDocRec
lnChequeCartera = 0
lnNumChequesCartera = 0
lnChequeCartera = oChq.GetImporteChequesCartera(lsMoneda, ldFecha, lnNumChequesCartera)

Set oChq = Nothing
End Sub

Private Sub SaldosChequeCustodia(ldFecha As Date, lsMoneda As String)
Dim sql As String
Dim rs As New ADODB.Recordset

sql = "SELECT SUM(nMontoChq) AS MontoChq, COUNT(*) AS TOTALCHQ " _
    & "From ChequeCaja " _
    & "WHERE cDepBco IN ('0') AND SUBSTRING(cCodCta,6,1)='" & lsMoneda & "'"
Set rs = oCon.CargaRecordSet(sql)

lnMontoChqCust = 0
lnNumChequesCust = 0
If Not RSVacio(rs) Then
   lnMontoChqCust = IIf(IsNull(rs!MontoChq), 0, rs!MontoChq)
   lnNumChequesCust = IIf(IsNull(rs!TotalChq), 0, rs!TotalChq)
End If
RSClose rs
End Sub

Private Function EncajeBaseAgencia(pnTipoBala As String, lsMoneda As String, ldFecha As Date) As Currency
Dim oBal As New NBalanceCont
EncajeBaseAgencia = oBal.GetUtilidadAcumulada(pnTipoBala, CInt(lsMoneda), Month(ldFecha), Year(ldFecha), True, False)
Set oBal = Nothing
End Function

''''Private Sub SaldosCreditos(lsMoneda As String, lsNomAge As String, lsCodAge As String, ldFecha As Date, Cont As Integer)
''''Dim rs   As ADODB.Recordset
''''Dim oEst As New NEstadisticas
'''''Creditos pequeña empresa
''''If gbBitCentral = True Then
''''    Set rs = oEst.GetEstadisticaCreditos(gbBitCentral, ldFecha, lsMoneda, lsCodAge, , " and substring(cLineaCred,7,1) in ('1','2') AND substring(cLineaCred,7,3) NOT IN ('" & gColComercAgro & "','" & gColPYMEAgro & "')")
''''Else
''''    Set rs = oEst.GetEstadisticaCreditos(gbBitCentral, ldFecha, lsMoneda, lsCodAge, , " and Left(cCodLinCred,1) in ('1','2') AND LEFT(cCodLinCred,3) NOT IN ('" & gColComercAgro & "','" & gColPYMEAgro & "')")
''''End If
''''ReDim Preserve lsPymes(3, Cont)
''''If Not RSVacio(rs) Then
''''    lsPymes(1, Cont) = lsNomAge
''''    lsPymes(2, Cont) = Format(IIf(IsNull(rs!Saldo), 0, rs!Saldo), "#,#0.00")
''''    lsPymes(3, Cont) = Format(IIf(IsNull(rs!TotalCreditos), 0, rs!TotalCreditos), "#,#0")
''''End If
''''RSClose rs
''''
'''''Creditos Personales Descuento por Planilla
''''Set rs = oEst.GetEstadisticaCreditos(gbBitCentral, ldFecha, lsMoneda, lsCodAge, gColConsuDctoPlan)
''''ReDim Preserve lsPersonales(3, Cont)
''''If Not RSVacio(rs) Then
''''    lsPersonales(1, Cont) = lsNomAge
''''    lsPersonales(2, Cont) = Format(IIf(IsNull(rs!Saldo), 0, rs!Saldo), "#,#0.00")
''''    lsPersonales(3, Cont) = Format(IIf(IsNull(rs!TotalCreditos), 0, rs!TotalCreditos), "#,#0.00")
''''End If
''''RSClose rs
'''''Creditos Otros (cts-Pf-Diversos)
''''
''''
''''If gbBitCentral = True Then
''''    Set rs = oEst.GetEstadisticaCreditos(gbBitCentral, ldFecha, lsMoneda, lsCodAge, , " and substring(cLineaCred,7,1) ='3' AND NOT substring(cLineaCred,7,3) = '" & gColConsuDctoPlan & "'")
''''Else
''''    Set rs = oEst.GetEstadisticaCreditos(gbBitCentral, ldFecha, lsMoneda, lsCodAge, , " and Left(cCodLinCred,1) = '3' AND NOT LEFT(cCodLinCred,3) = '" & gColConsuDctoPlan & "' ")
''''End If
''''
''''ReDim Preserve lsPersOtros(3, Cont)
''''If Not RSVacio(rs) Then
''''   lsPersOtros(1, Cont) = lsNomAge
''''   lsPersOtros(2, Cont) = Format(IIf(IsNull(rs!Saldo), 0, rs!Saldo), "#,#0.00")
''''   lsPersOtros(3, Cont) = Format(IIf(IsNull(rs!TotalCreditos), 0, rs!TotalCreditos), "#,#0.00")
''''End If
''''RSClose rs
''''
'''''CREDITOS AGRICOLAS
''''If gbBitCentral = True Then
''''    Set rs = oEst.GetEstadisticaCreditos(gbBitCentral, ldFecha, lsMoneda, lsCodAge, , " and substring(cLineaCred,7,3) IN ('" & gColComercAgro & "','" & gColPYMEAgro & "')")
''''Else
''''    Set rs = oEst.GetEstadisticaCreditos(gbBitCentral, ldFecha, lsMoneda, lsCodAge, , " and LEFT(cCodLinCred,3) IN ('" & gColComercAgro & "','" & gColPYMEAgro & "')")
''''End If
''''
''''ReDim Preserve lsAgricolas(3, Cont)
''''If Not RSVacio(rs) Then
''''    lsAgricolas(1, Cont) = lsNomAge
''''    lsAgricolas(2, Cont) = Format(rs!Saldo, "#,#0.00")
''''    lsAgricolas(3, Cont) = Format(rs!TotalCreditos, "#,#0")
''''End If
''''RSClose rs
''''Set oEst = Nothing
''''End Sub

Private Sub GeneraReporte(psOpeCod As String, psMoneda As String, pdFecha As Date)
    Dim fs As New Scripting.FileSystemObject
    Dim lsMoneda As String
    Dim lnFila As Integer, I As Integer, lnCol As Integer, J As Integer
'    Dim lsTotalSaldos() As String
    
'    Dim lsTotalPrenda() As String
'    Dim lsTotalCredPE() As String
'    Dim lsTotalCredPPDP() As String
'    Dim lsTotalCredOtros() As String
'    Dim lsTotalCredAgric() As String
'    Dim lnFilaPers As Integer
'    Dim lnTotalColocaciones(3) As String
    Dim Y1 As Currency, Y2 As Currency
    
    Dim lsFormulasSaldos As String
    Dim lsFormulasTotales As String
    Dim lbExisteHoja As Boolean
'    Dim lnFilaAgric As Integer
    Dim lsTotalPlazoFijo As String
    Dim lsTotalObligaciones As String
    
Dim oEst As New NEstadisticas
   lsRutaReferencia = "+^" & App.path & "\SPOOLER\[" & lsNomArch & "]" & Format(pdFecha, "dd-mm-yyyy") & "^!"

    lsMoneda = psMoneda
    lsTotalPlazoFijo = ""
    
    xlHoja1.PageSetup.Zoom = 75
    xlAplicacion.Range("A1:R100").Font.Size = 9
    
    xlHoja1.Range("A1").ColumnWidth = 18
    xlHoja1.Range("B1:P1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 13
    xlHoja1.Range("D1").ColumnWidth = 18
    xlHoja1.Range("E1").ColumnWidth = 18
    xlHoja1.Range("F1:H1").ColumnWidth = 12
    
    xlHoja1.Cells(1, 1) = gsNomCmac
    xlAplicacion.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 1)).Font.Bold = True
    xlHoja1.Cells(1, 7) = "Fecha :" & Format(pdFecha, "dd mmmm yyyy")
    xlAplicacion.Range(xlHoja1.Cells(1, 7), xlHoja1.Cells(1, 7)).Font.Bold = True
    xlHoja1.Cells(2, 1) = "Area de Caja General"
    xlAplicacion.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 1)).Font.Bold = True
    xlHoja1.Cells(3, 3) = "INFORME ESTADISTICO DIARIO EN " & IIf(lsMoneda = "1", "MONEDA NACIONAL", "MONEDA EXTRANJERA")
    xlAplicacion.Range(xlHoja1.Cells(3, 3), xlHoja1.Cells(3, 3)).Font.Bold = True
    
    xlAplicacion.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(9, UBound(lsSaldos, 1))).HorizontalAlignment = xlHAlignCenter
    xlAplicacion.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(9, UBound(lsSaldos, 1))).Font.Bold = True
    
    ExcelCuadro xlHoja1, 1, 7, UBound(lsSaldos, 1), 9
'    xlHoja1.Cells(7, 1) = "AGENCIA":
'    xlHoja1.Cells(7, 2) = "SALDO":
'    xlHoja1.Cells(7, 3) = "SALDO":
'    xlHoja1.Cells(7, 4) = "INTERES":
'    xlHoja1.Cells(7, 5) = "TOTAL":
'    xlHoja1.Cells(7, 6) = "CAJA":
'    xlHoja1.Cells(7, 7) = "CAJA":
'    xlHoja1.Cells(7, 8) = "CAJA":
'    xlHoja1.Cells(7, 9) = "FONDOS DE":
'    xlHoja1.Cells(7, 10) = "ENCAJE":
'    xlHoja1.Cells(7, 11) = "SUPERAVIT/"
'    xlHoja1.Cells(7, 12) = "VARIACION"
'
'
'    xlHoja1.Cells(8, 2) = "AHORRO CTE":
'    xlHoja1.Cells(8, 3) = "P. FIJO":
'    xlHoja1.Cells(8, 4) = "P. FIJO":
'    xlHoja1.Cells(8, 5) = "TOSE":
'    xlHoja1.Cells(8, 6) = "PROMEDIO":
'    xlHoja1.Cells(8, 7) = "DEL DIA":
'    xlHoja1.Cells(8, 8) = "DEL DIA":
'    xlHoja1.Cells(8, 9) = "ENCAJE":
'    xlHoja1.Cells(8, 10) = "EXIGIBLE":
'    xlHoja1.Cells(8, 11) = "DEFICIT"
'    xlHoja1.Cells(8, 12) = "DE"
'
'    xlHoja1.Cells(9, 7) = "ANTERIOR":
'    xlHoja1.Cells(9, 8) = "REPORTADO":
'    xlHoja1.Cells(9, 9) = "Caja + BCR":
'    xlHoja1.Cells(9, 11) = "DIARIO"
'    xlHoja1.Cells(9, 12) = "CAJA"

    xlHoja1.Cells(7, 1) = "AGENCIA":
    xlHoja1.Cells(7, 2) = "SALDO":
    xlHoja1.Cells(7, 3) = "SALDO":
    xlHoja1.Cells(7, 4) = "OBLIGACIONES":
    xlHoja1.Cells(7, 5) = "TOTAL":
    xlHoja1.Cells(7, 6) = "CAJA":
    xlHoja1.Cells(7, 7) = "CAJA":
    xlHoja1.Cells(7, 8) = "CAJA":
    xlHoja1.Cells(7, 9) = "FONDOS DE":
    xlHoja1.Cells(7, 10) = "ENCAJE":
    xlHoja1.Cells(7, 11) = "SUPERAVIT/"
    xlHoja1.Cells(7, 12) = "VARIACION"
    
    
    xlHoja1.Cells(8, 2) = "AHORRO CTE":
    xlHoja1.Cells(8, 3) = "P. FIJO":
    xlHoja1.Cells(8, 4) = "INMEDIATAS":
    xlHoja1.Cells(8, 5) = "TOSE":
    xlHoja1.Cells(8, 6) = "PROMEDIO":
    xlHoja1.Cells(8, 7) = "DEL DIA":
    xlHoja1.Cells(8, 8) = "DEL DIA":
    xlHoja1.Cells(8, 9) = "ENCAJE":
    xlHoja1.Cells(8, 10) = "EXIGIBLE":
    xlHoja1.Cells(8, 11) = "DEFICIT"
    xlHoja1.Cells(8, 12) = "DE"
    
    xlHoja1.Cells(9, 7) = "ANTERIOR":
    xlHoja1.Cells(9, 8) = "REPORTADO":
    xlHoja1.Cells(9, 9) = "Caja + BCR":
    xlHoja1.Cells(9, 11) = "DIARIO"
    xlHoja1.Cells(9, 12) = "CAJA"
    
    '***************** datos de Saldos de Ahorros y Plazo Fijo *****************************
    ReDim lsTotalSaldos(UBound(lsSaldos, 1))
    Y1 = 10
    xlHoja1.Cells(10, 1) = "Caja General":
    
    xlHoja1.Cells(10, 4) = Format(lnObligacionesCaja, "#,#0.00")
    lsTotalSaldos(4) = xlHoja1.Range(xlHoja1.Cells(10, 4), xlHoja1.Cells(10, 4)).Address(False, False) & "+"
    
    xlHoja1.Cells(10, 5) = Format(lnObligacionesCaja, "#,#0.00")
    lsTotalSaldos(5) = xlHoja1.Range(xlHoja1.Cells(10, 5), xlHoja1.Cells(10, 5)).Address(False, False) & "+"
    
    xlHoja1.Cells(10, 6) = Format(lnMontoCaja, "#,#0.00")
    lsTotalSaldos(6) = xlHoja1.Range(xlHoja1.Cells(10, 6), xlHoja1.Cells(10, 6)).Address(False, False) & "+"


    xlHoja1.Cells(10, 7) = Format(lnMontoCajaDiaAnt, "#,#0.00"):
    lsTotalSaldos(7) = xlHoja1.Range(xlHoja1.Cells(10, 7), xlHoja1.Cells(10, 7)).Address(False, False) & "+"

    xlHoja1.Cells(10, 8) = Format(lnMontoCajaDia, "#,#0.00"):
    lsTotalSaldos(8) = xlHoja1.Range(xlHoja1.Cells(10, 8), xlHoja1.Cells(10, 8)).Address(False, False) & "+"

    If lsMoneda = "1" Then
        xlHoja1.Cells(10, 9) = Format(lnMontoCaja + lnMontoBCR, "#,#0.00"):
        lsTotalSaldos(9) = xlHoja1.Range(xlHoja1.Cells(10, 9), xlHoja1.Cells(10, 9)).Address(False, False) & "+"
        
        xlHoja1.Cells(10, 10) = Format(lnFondoExigible(0), "#,#0.00")
        lsTotalSaldos(10) = xlHoja1.Range(xlHoja1.Cells(10, 10), xlHoja1.Cells(10, 10)).Address(False, False) & "+"
        
        xlHoja1.Cells(10, 11) = Format(lnMontoCaja + lnMontoBCR - lnFondoExigible(0), "#,#0.00")
        lsTotalSaldos(11) = xlHoja1.Range(xlHoja1.Cells(10, 11), xlHoja1.Cells(10, 11)).Address(False, False) & "+"
    Else
        xlHoja1.Cells(10, 9) = Format(lnMontoCajaDia + lnMontoBCR, "#,#0.00"):
        lsTotalSaldos(9) = xlHoja1.Range(xlHoja1.Cells(10, 9), xlHoja1.Cells(10, 9)).Address(False, False) & "+"
        
        xlHoja1.Cells(10, 10) = Format(lnFondoExigible(0), "#,#0.00")
        lsTotalSaldos(10) = xlHoja1.Range(xlHoja1.Cells(10, 10), xlHoja1.Cells(10, 10)).Address(False, False) & "+"
        
        xlHoja1.Cells(10, 11) = Format(lnMontoCajaDia + lnMontoBCR, "#,#0.00")
        lsTotalSaldos(11) = xlHoja1.Range(xlHoja1.Cells(10, 11), xlHoja1.Cells(10, 11)).Address(False, False) & "+"
    End If
    lnFila = 10
    
    'Variacion de Caja
    xlHoja1.Cells(10, 12) = Format(lnMontoCajaDia - lnMontoCajaDiaAnt, "#,#0.00")
    lsTotalSaldos(12) = xlHoja1.Range(xlHoja1.Cells(10, 12), xlHoja1.Cells(10, 12)).Address(False, False) & "+"

    For I = 1 To UBound(lsSaldos, 2)
        lnCol = 0
        lnFila = lnFila + 1
        For J = 1 To UBound(lsSaldos, 1)
            lnCol = lnCol + 1
            Select Case J
                Case 10
                    'If i = 1 Then
                    '    lsSaldos(J, 0) = lnFondoExigible(0)
                    'End If
                    lsSaldos(J, I) = Format(lnFondoExigible(I), "#,#0.00")
                Case 11
                    'If i = 1 Then
                    '    lsSaldos(J, 0) = lnFondoExigible(0)
                    'End If
                    lsSaldos(J, I) = Format(CCur(lsSaldos(9, I)) - lnFondoExigible(I), "#,#0.00")
            End Select
            If J = 10 And lnFila = 11 Then
                xlHoja1.Cells(lnFila - 1, lnCol) = lsSaldos(J, 0)
            End If
            xlHoja1.Cells(lnFila, lnCol) = lsSaldos(J, I)
            If J = 1 Then
                lsTotalSaldos(J) = "TOTAL "
            Else
               'If J = 10 And lnFila = 11 Then
               '     lsTotalSaldos(J) = lsTotalSaldos(J) + xlHoja1.Range(xlHoja1.Cells(lnFila - 1, lnCol), xlHoja1.Cells(lnFila - 1, lnCol)).Address(False, False) & "+" + xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
               'Else
                    lsTotalSaldos(J) = lsTotalSaldos(J) + xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
               'End If
            End If
        Next J
    Next I
    
    xlHoja1.Cells(10, 10) = Format(lnFondoExigible(0), "#,#0.00")
    
    lnCol = 0
    Y2 = lnFila
    ExcelCuadro xlHoja1, 1, Y1, 12, Y2
    
    Y1 = lnFila + 1
    lnFila = lnFila + 2
    oEst.EliminaEstadAnexos pdFecha, "AHORROS", psMoneda
    oEst.EliminaEstadAnexos pdFecha, "OBLIGPLAZO", psMoneda
    oEst.EliminaEstadAnexos pdFecha, "CAJA", psMoneda
    oEst.EliminaEstadAnexos pdFecha, "FONDOCAJA", psMoneda
    oEst.EliminaEstadAnexos pdFecha, "EXIGIBLE", psMoneda
    oEst.EliminaEstadAnexos pdFecha, "FONDOSBCR", psMoneda
    oEst.EliminaEstadAnexos pdFecha, "CMACSAHO", psMoneda
    oEst.EliminaEstadAnexos pdFecha, "CMACSPF", psMoneda
    
    Dim lnImporteObligPlazo As Currency
    For J = 1 To UBound(lsTotalSaldos)
        lnCol = lnCol + 1
        If J = 1 Then
            xlHoja1.Cells(lnFila, lnCol) = lsTotalSaldos(J)
        Else
            xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=SUM(" & Mid(lsTotalSaldos(J), 1, Len(lsTotalSaldos(J)) - 1) & ")"
        End If
        xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
        Select Case J
            Case 2
                oEst.InsertaEstadAnexos pdFecha, "AHORROS", psMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address
            Case 3
                lsTotalPlazoFijo = lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address
                lsTotalObligaciones = lsTotalPlazoFijo
            Case 4
                'lsTotalObligaciones = lsTotalObligaciones & "+" & lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address
                oEst.InsertaEstadAnexos pdFecha, "OBLIGPLAZO", psMoneda, lsTotalObligaciones
            Case 6
                If lsMoneda = "1" Then
                    oEst.InsertaEstadAnexos pdFecha, "FONDOCAJA", psMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address
                End If
            Case 8
                oEst.InsertaEstadAnexos pdFecha, "CAJA", lsMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address
                If lsMoneda = "2" Then
                    oEst.InsertaEstadAnexos pdFecha, "FONDOCAJA", lsMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address
                End If
            Case 10
                oEst.InsertaEstadAnexos pdFecha, "EXIGIBLE", psMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address
        End Select
    Next J
    Y2 = lnFila + 1
    ExcelCuadro xlHoja1, 1, Y1, 12, Y2
    
    xlHoja1.Cells(lnFila + 1, 8) = "SALDO BCR :"
    xlHoja1.Cells(lnFila + 1, 9) = Format(lnMontoBCR, "#,#0.00"):
    oEst.InsertaEstadAnexos pdFecha, "FONDOSBCR", lsMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 9), xlHoja1.Cells(lnFila + 1, 9)).Address
    
    lnFila = lnFila + 2
        
    '********************** OTROS DATOS ADICIONALES **************************************************************************
    Y1 = lnFila
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlHAlignCenter
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
    xlHoja1.Cells(lnFila, 1) = "OTROS": xlHoja1.Cells(lnFila, 2) = "SALDO" & IIf(lsMoneda = "1", gcPEN_SIMBOLO, "$.") 'marg ers044-2016
    xlHoja1.Cells(lnFila, 3) = "Nº":
    Y2 = lnFila
    ExcelCuadro xlHoja1, 1, Y1, 6, Y2

    lnFila = lnFila + 1
    Y1 = lnFila
    xlHoja1.Cells(lnFila, 1) = "CHQ. CARTERA":
    xlHoja1.Cells(lnFila, 2) = Format(lnChequeCartera, "#,#0.00"):
    xlHoja1.Cells(lnFila, 3) = Format(lnNumChequesCartera, "#,#0"):

    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "CHQ. CUSTODIA":
    xlHoja1.Cells(lnFila, 2) = Format(lnMontoChqCust, "#,#0.00"):
    xlHoja1.Cells(lnFila, 3) = Format(lnNumChequesCust, "#,#0.00")

    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "CARTAS FIANZAS":
    xlHoja1.Cells(lnFila, 3) = Format(lnTotalCartaFianza, "#,#0"):
    xlHoja1.Cells(lnFila, 2) = Format(lnMontoCartaFianza, "#,#0.00"):
'    xlHoja1.Cells(lnFila, 10) = Format(lnMontoCmacs, "#,#0.00")

    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "TOTAL CTS "
    xlHoja1.Cells(lnFila, 2) = Format(lnSaldoTotalCTS, "#,#0.00")

    Y2 = lnFila + 1
    ExcelCuadro xlHoja1, 1, Y1, 3, Y2

    MuestraDatosComplementarios lsMoneda, Y1 - 1, pdFecha
    If lsMoneda = "2" And lnFila < Y1 + 5 Then
       lnFila = Y1 + 5
    End If
    
    
    '***************** datos Prendario *****************************
'''''''    If lsMoneda = "1" Then
'''''''        lnFila = lnFila + 2
'''''''        Y1 = lnFila
'''''''        xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 4)).HorizontalAlignment = xlHAlignCenter
'''''''        xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 4)).Font.Bold = True
'''''''        xlHoja1.Cells(lnFila, 1) = "CREDITO PRENDARIO": xlHoja1.Cells(lnFila, 2) = "SALDO " & IIf(lsMoneda = "1", "S/.", "$."): xlHoja1.Cells(lnFila, 3) = "Nº": xlHoja1.Cells(lnFila, 4) = "GRAMOS"
'''''''        Y2 = lnFila
'''''''        ExcelCuadro xlHoja1, 1, Y1, 4, Y2
'''''''
'''''''        Y1 = lnFila + 1
'''''''        ReDim Preserve lsTotalPrenda(UBound(lsPrendario, 1))
'''''''        For I = 1 To UBound(lsPrendario, 2)
'''''''            lnCol = 0
'''''''            If CCur(lsPrendario(2, I)) <> 0 Then
'''''''                lnFila = lnFila + 1
'''''''                For J = 1 To UBound(lsPrendario, 1)
'''''''                    lnCol = lnCol + 1
'''''''                    xlHoja1.Cells(lnFila, lnCol) = lsPrendario(J, I)
'''''''                    If J = 1 Then
'''''''                        lsTotalPrenda(J) = "TOTAL"
'''''''                    Else
'''''''                        lsTotalPrenda(J) = lsTotalPrenda(J) + xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'''''''                    End If
'''''''                Next J
'''''''            End If
'''''''        Next I
'''''''
'''''''        For I = 1 To UBound(lsAdjudicados, 2)
'''''''            lnCol = 0
'''''''            If CCur(lsAdjudicados(2, I)) <> 0 Then
'''''''                lnFila = lnFila + 1
'''''''                For J = 1 To UBound(lsAdjudicados, 1)
'''''''                    lnCol = lnCol + 1
'''''''                    xlHoja1.Cells(lnFila, lnCol) = lsAdjudicados(J, I)
'''''''                    If J = 1 Then
'''''''                        lsTotalPrenda(J) = "SUBTOTALES"
'''''''                    Else
'''''''                        lsTotalPrenda(J) = lsTotalPrenda(J) + xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'''''''                    End If
'''''''                Next J
'''''''            End If
'''''''        Next I
'''''''
'''''''        Y2 = lnFila
'''''''        ExcelCuadro xlHoja1, 1, Y1, 4, Y2
'''''''
'''''''        lnCol = 0
'''''''        lnFila = lnFila + 1
'''''''        Y1 = lnFila
'''''''        For J = 1 To UBound(lsTotalPrenda)
'''''''            lnCol = lnCol + 1
'''''''            If J = 1 Then
'''''''                xlHoja1.Cells(lnFila, lnCol) = lsTotalPrenda(J)
'''''''            Else
'''''''                If lsTotalPrenda(J) <> "" Then
'''''''                    xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=SUM(" & Mid(lsTotalPrenda(J), 1, Len(lsTotalPrenda(J)) - 1) & ")"
'''''''                End If
'''''''            End If
'''''''
'''''''            xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'''''''            Select Case J
'''''''                Case 2
'''''''                    lsFormulasSaldos = xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) + "+"
'''''''                Case 3
'''''''                    lsFormulasTotales = xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) + "+"
'''''''            End Select
'''''''        Next J
'''''''        Y2 = lnFila
'''''''        ExcelCuadro xlHoja1, 1, Y1, 4, Y2
'''''''
'''''''    End If
'''''''    '***************** datos Creditos Personales Descuento por Planilla *****************************
'''''''    lnFila = lnFila + 2
'''''''    Y1 = lnFila
'''''''    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
'''''''    xlHoja1.Cells(lnFila, 1) = "CREDITOS PERSONALES"
'''''''
'''''''    xlAplicacion.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
'''''''    xlHoja1.Cells(lnFila, 4) = "Creditos Hipotecarios":
'''''''    xlHoja1.Cells(lnFila, 5) = Format(lnCredHipotecarios, "#,#0.00")
'''''''
'''''''    lsFormulasSaldos = lsFormulasSaldos + xlAplicacion.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Address(False, False) & "+"
'''''''    lsFormulasTotales = lsFormulasTotales + xlAplicacion.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False) & "+"
'''''''
'''''''    lnFila = lnFila + 1
'''''''    xlAplicacion.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
'''''''    xlHoja1.Cells(lnFila, 4) = "Creditos Judiciales":
'''''''    xlHoja1.Cells(lnFila, 5) = Format(lnSaldosJudiciales, "#,#0.00")
'''''''    xlHoja1.Cells(lnFila, 6) = Format(lnTotalJudiciales, "#,#0")
'''''''    lsFormulasSaldos = lsFormulasSaldos + xlAplicacion.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Address(False, False) & "+"
'''''''    lsFormulasTotales = lsFormulasTotales + xlAplicacion.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False) & "+"
'''''''    Y2 = lnFila
'''''''    ExcelCuadro xlHoja1, 4, Y1, 6, Y2
'''''''
'''''''    lnFila = lnFila + 1
'''''''    lnFilaPers = lnFila
'''''''    Y1 = lnFila
'''''''    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlHAlignCenter
'''''''    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
'''''''    xlHoja1.Cells(lnFila, 1) = "DSCTO.PLANILLA ": xlHoja1.Cells(lnFila, 2) = "SALDO " & IIf(lsMoneda = "1", "S/.", "$."): xlHoja1.Cells(lnFila, 3) = "Nº"
'''''''    Y2 = lnFila
'''''''    ExcelCuadro xlHoja1, 1, Y1, 6, Y2
'''''''    Y1 = lnFila + 1
'''''''    ReDim Preserve lsTotalCredPPDP(UBound(lsPersonales, 1))
'''''''    For I = 1 To UBound(lsPersonales, 2)
'''''''        lnCol = 0
'''''''        If CCur(lsPersonales(2, I)) <> 0 Then
'''''''            lnFila = lnFila + 1
'''''''            For J = 1 To UBound(lsPersonales, 1)
'''''''                lnCol = lnCol + 1
'''''''                xlHoja1.Cells(lnFila, lnCol) = lsPersonales(J, I)
'''''''                If J = 1 Then
'''''''                    lsTotalCredPPDP(J) = "SUBTOTALES"
'''''''                Else
'''''''                    lsTotalCredPPDP(J) = lsTotalCredPPDP(J) + xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'''''''                End If
'''''''            Next J
'''''''        End If
'''''''    Next I
'''''''
'''''''    '***************** datos Creditos Personales OTROS *****************************
'''''''    xlHoja1.Cells(lnFilaPers, 4) = "CTS-PF-DIVERSOS": xlHoja1.Cells(lnFilaPers, 5) = "SALDO " & IIf(lsMoneda = "1", "S/.", "$."): xlHoja1.Cells(lnFilaPers, 6) = "Nº"
'''''''    ReDim Preserve lsTotalCredOtros(UBound(lsPersOtros, 1))
'''''''    For I = 1 To UBound(lsPersOtros, 2)
'''''''        lnCol = 3
'''''''        If CCur(lsPersOtros(2, I)) <> 0 Then
'''''''            lnFilaPers = lnFilaPers + 1
'''''''            For J = 1 To UBound(lsPersOtros, 1)
'''''''                lnCol = lnCol + 1
'''''''                xlHoja1.Cells(lnFilaPers, lnCol) = lsPersOtros(J, I)
'''''''                If J = 1 Then
'''''''                    lsTotalCredOtros(J) = "SUBTOTALES"
'''''''                Else
'''''''                    'lsTotalCredOtros(j) = Format(CCur(IIf(lsTotalCredOtros(j) = "", "0", lsTotalCredOtros(j))) + CCur(lsPersOtros(j, i)), "#,#0.00")
'''''''                    lsTotalCredOtros(J) = lsTotalCredOtros(J) + xlHoja1.Range(xlHoja1.Cells(lnFilaPers, lnCol), xlHoja1.Cells(lnFilaPers, lnCol)).Address(False, False) & "+"
'''''''                End If
'''''''            Next J
'''''''        End If
'''''''    Next I
'''''''    If lnFilaPers > lnFila Then
'''''''        lnFila = lnFilaPers
'''''''    End If
'''''''    Y2 = lnFila
'''''''    ExcelCuadro xlHoja1, 1, Y1, 6, Y2
'''''''    '****** TOTALES DE CREDITOS PERSONALES *****************************
'''''''    lnCol = 0
'''''''    lnFila = lnFila + 1
'''''''    Y1 = lnFila
'''''''    For J = 1 To UBound(lsTotalCredPPDP)
'''''''        lnCol = lnCol + 1
'''''''        If J = 1 Then
'''''''            xlHoja1.Cells(lnFila, lnCol) = lsTotalCredPPDP(J)
'''''''        Else
'''''''            If lsTotalCredPPDP(J) <> "" Then
'''''''               xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=Sum(" & Mid(lsTotalCredPPDP(J), 1, Len(lsTotalCredPPDP(J)) - 1) & ")"
'''''''            End If
'''''''        End If
'''''''
'''''''        xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'''''''        Select Case J
'''''''            Case 2
'''''''                lsFormulasSaldos = lsFormulasSaldos + xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'''''''            Case 3
'''''''                lsFormulasTotales = lsFormulasTotales + xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'''''''        End Select
'''''''    Next J
'''''''    For J = 1 To UBound(lsTotalCredOtros)
'''''''        lnCol = lnCol + 1
'''''''        If J = 1 Then
'''''''            xlHoja1.Cells(lnFila, lnCol) = lsTotalCredOtros(J)
'''''''        Else
'''''''            If lsTotalCredOtros(J) <> "" Then
'''''''               xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=Sum(" & Mid(lsTotalCredOtros(J), 1, Len(lsTotalCredOtros(J)) - 1) & ")"
'''''''            End If
'''''''        End If
'''''''
'''''''        xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'''''''        Select Case J
'''''''            Case 2
'''''''                lsFormulasSaldos = lsFormulasSaldos + xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'''''''            Case 3
'''''''                lsFormulasTotales = lsFormulasTotales + xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'''''''        End Select
'''''''    Next J
'''''''    Y2 = lnFila
'''''''    ExcelCuadro xlHoja1, 1, Y1, 6, Y2
'''''''    '***************** datos Creditos pequeña empresa*****************************
'''''''    lnFila = lnFila + 2
'''''''    Y1 = lnFila
'''''''    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlHAlignCenter
'''''''    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
'''''''    xlHoja1.Cells(lnFila, 1) = "CREDITO P.E.": xlHoja1.Cells(lnFila, 2) = "SALDO " & IIf(lsMoneda = "1", "S/.", "$."): xlHoja1.Cells(lnFila, 3) = "Nº": xlHoja1.Cells(lnFila, 4) = "AGRICOLAS": xlHoja1.Cells(lnFila, 5) = "SALDO " & IIf(lsMoneda = "1", "S/.", "$."): xlHoja1.Cells(lnFila, 6) = "Nº"
'''''''    Y2 = lnFila
'''''''    ExcelCuadro xlHoja1, 1, Y1, 6, Y2
'''''''
'''''''    '********************** Creditos del Santa ***********************************
'''''''    ReDim lsTotalCredPE(UBound(lsPymes, 1))
'''''''
'''''''    ReDim lsTotalCredAgric(UBound(lsPymes, 1))
'''''''
'''''''    lnFila = lnFila + 1
'''''''    Y1 = lnFila
'''''''    If Val(lnCreditosSanta(2)) > 0 Then
'''''''        xlHoja1.Cells(lnFila, 1) = lnCreditosSanta(1): xlHoja1.Cells(lnFila, 2) = lnCreditosSanta(2): xlHoja1.Cells(lnFila, 3) = lnCreditosSanta(3)
'''''''        lsTotalCredPE(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).Address(False, False) & "+"
'''''''        lsTotalCredPE(3) = xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Address(False, False) & "+"
'''''''    End If
'''''''    lnFilaAgric = 0
'''''''    For I = 1 To UBound(lsPymes, 2)
'''''''        lnCol = 0
'''''''        If CCur(lsPymes(2, I)) <> 0 Then
'''''''            lnFila = lnFila + 1
'''''''            If lnFilaAgric = 0 Then
'''''''               lnFilaAgric = lnFila
'''''''            End If
'''''''            For J = 1 To UBound(lsPymes, 1)
'''''''                lnCol = lnCol + 1
'''''''                xlHoja1.Cells(lnFila, lnCol) = lsPymes(J, I)
'''''''                If J = 1 Then
'''''''                    lsTotalCredPE(J) = "SUBTOTALES"
'''''''                Else
'''''''                    lsTotalCredPE(J) = lsTotalCredPE(J) + xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'''''''                End If
'''''''            Next J
'''''''        End If
'''''''    Next I
'''''''
'''''''
'''''''    '******************  TOTALES DE CREDITOS AGRICOLAS ************************
'''''''    For I = 1 To UBound(lsAgricolas, 2)
'''''''        lnCol = 3
'''''''        If CCur(lsAgricolas(2, I)) <> 0 Then  '**************** CAMBIAR A MATRIZ DE AGRICOLAS
'''''''            For J = 1 To UBound(lsAgricolas, 1)    '**************** CAMBIAR A MATRIZ DE AGRICOLAS
'''''''                lnCol = lnCol + 1
'''''''                xlHoja1.Cells(lnFilaAgric, lnCol) = lsAgricolas(J, I)
'''''''                If J = 1 Then
'''''''                    '******************  TOTALES DE CREDITOS AGRICOLAS ************************
'''''''                    lsTotalCredAgric(J) = "SUBTOTALES"
'''''''                Else
'''''''                    '******************  TOTALES DE CREDITOS AGRICOLAS ************************
'''''''                    lsTotalCredAgric(J) = lsTotalCredAgric(J) + xlHoja1.Range(xlHoja1.Cells(lnFilaAgric, lnCol), xlHoja1.Cells(lnFilaAgric, lnCol)).Address(False, False) & "+"
'''''''                End If
'''''''            Next J
'''''''            lnFilaAgric = lnFilaAgric + 1
'''''''        End If
'''''''    Next I
'''''''    '*********************************************************************************
'''''''    lnCol = 0
'''''''    lnFila = lnFila + 1
'''''''    If lnFila < lnFilaAgric Then
'''''''        lnFila = lnFilaAgric
'''''''    End If
'''''''    Y2 = lnFila
'''''''    ExcelCuadro xlHoja1, 1, Y1, 6, Y2
'''''''    Y1 = lnFila
'''''''
'''''''    For J = 1 To UBound(lsTotalCredPE)
'''''''        lnCol = lnCol + 1
'''''''        If J = 1 Then
'''''''            xlHoja1.Cells(lnFila, lnCol) = lsTotalCredPE(J)
'''''''        Else
'''''''            If lsTotalCredPE(J) <> "" Then
'''''''               xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=Sum(" & Mid(lsTotalCredPE(J), 1, Len(lsTotalCredPE(J)) - 1) & ")"
'''''''            End If
'''''''        End If
'''''''
'''''''        xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'''''''        Select Case J
'''''''            Case 2
'''''''                lsFormulasSaldos = lsFormulasSaldos + xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'''''''            Case 3
'''''''                lsFormulasTotales = lsFormulasTotales + xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'''''''        End Select
'''''''    Next J
'''''''    '******************** TOTALES DE CREDITOS AGRICOLAS ********************************
'''''''    lnCol = 3
'''''''    For J = 1 To UBound(lsTotalCredAgric)
'''''''        lnCol = lnCol + 1
'''''''        If J = 1 Then
'''''''            xlHoja1.Cells(lnFila, lnCol) = lsTotalCredAgric(J)
'''''''        Else
'''''''            If lsTotalCredAgric(J) <> "" Then
'''''''               xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=Sum(" & Mid(lsTotalCredAgric(J), 1, Len(lsTotalCredAgric(J)) - 1) & ")"
'''''''            End If
'''''''        End If
'''''''        xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'''''''        Select Case J
'''''''            Case 2
'''''''                lsFormulasSaldos = lsFormulasSaldos + xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'''''''            Case 3
'''''''                lsFormulasTotales = lsFormulasTotales + xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'''''''        End Select
'''''''    Next J
    
  
'    Y2 = lnFila
'    ExcelCuadro xlHoja1, 1, Y1, 6, Y2
'    lnTotalColocaciones(1) = "TOTAL COLOCAC."
'    Y1 = lnFila + 1
'    lnFila = lnFila + 2
'    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
'    xlHoja1.Cells(lnFila, 1) = lnTotalColocaciones(1)
'    xlAplicacion.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).Formula = "=Sum(" & Mid(lsFormulasSaldos, 1, Len(lsFormulasSaldos) - 1) & ")"
'    xlAplicacion.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Formula = "=Sum(" & Mid(lsFormulasTotales, 1, Len(lsFormulasTotales) - 1) & ")"
'    Y2 = lnFila
'    ExcelCuadro xlHoja1, 1, Y1, 3, Y2
    
End Sub

Private Sub MuestraDatosComplementarios(lsMoneda As String, ByVal pnFila As Integer, pdFecha As Date)
   Dim sSql As String
   Dim rs   As ADODB.Recordset
    
    xlHoja1.Cells(pnFila, 5) = "OTRAS FINANC.": xlHoja1.Cells(pnFila, 6) = "AHORROS": xlHoja1.Cells(pnFila, 7) = "PLAZO FIJO"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 5), xlHoja1.Cells(pnFila, 5)).Font.Bold = True
    xlHoja1.Cells(pnFila + 1, 5) = "CMACS"
    xlHoja1.Cells(pnFila + 2, 5) = "OTRAS.IFIS"
    
    xlHoja1.Cells(pnFila + 3, 5) = "CHEQUES"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 3, 5), xlHoja1.Cells(pnFila + 3, 5)).Font.Bold = True
    xlHoja1.Cells(pnFila + 4, 5) = "CMAC-ICA"
    xlHoja1.Cells(pnFila + 5, 5) = "CMACS"
    xlHoja1.Cells(pnFila + 6, 5) = "OTRAS.IFIS"
        
   Dim oEst As New NEstadisticas
   Set rs = oEst.GetEstadisticaAhorro(gbBitCentral, pdFecha, pdFecha, lsMoneda)
   If Not rs.EOF Then
      xlHoja1.Cells(pnFila + 1, 6) = rs!nSaldoCMAC
      xlHoja1.Cells(pnFila + 2, 6) = Format(rs!nSaldCRAC + rs!nSaldEdiPyme, "#,#0.00")
      xlHoja1.Cells(pnFila + 4, 6) = rs!nMonChqVal
      xlHoja1.Cells(pnFila + 5, 6) = rs!nChqCMAC
      xlHoja1.Cells(pnFila + 6, 6) = rs!nChqCRAC
      
   End If

   Set rs = oEst.GetEstadisticaPlazoFijoCts(gbBitCentral, pdFecha, pdFecha, lsMoneda)
   If Not rs.EOF Then
      xlHoja1.Cells(pnFila + 1, 7) = rs!nSaldoCMAC
      xlHoja1.Cells(pnFila + 2, 7) = Format(rs!nSaldCRAC + rs!nSaldEdiPyme, "#,#0.00")
      xlHoja1.Cells(pnFila + 4, 7) = rs!nMonChqVal
      xlHoja1.Cells(pnFila + 5, 7) = rs!nChqCMAC
   End If
   xlHoja1.Range(xlHoja1.Cells(pnFila + 1, 6), xlHoja1.Cells(pnFila + 1, 6)).NumberFormat = "#,##0.00;-#,##0.00"
   xlHoja1.Range(xlHoja1.Cells(pnFila + 1, 7), xlHoja1.Cells(pnFila + 1, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
   oEst.InsertaEstadAnexos pdFecha, "CMACSAHO", lsMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(pnFila + 1, 6), xlHoja1.Cells(pnFila + 1, 6)).Address
   oEst.InsertaEstadAnexos pdFecha, "CMACSPF", lsMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(pnFila + 1, 7), xlHoja1.Cells(pnFila + 1, 7)).Address
   
   ExcelCuadro xlHoja1, 5, CCur(pnFila), 7, CCur(pnFila)
   ExcelCuadro xlHoja1, 5, CCur(pnFila), 7, CCur(pnFila) + 6
   
End Sub

Private Sub CartasFianzas(psOpeCod As String, ldFecha As Date, lsMoneda As String)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim lsCodCtaCont As String

Dim oOpe As New DOperacion
lsCodCtaCont = oOpe.EmiteOpeCta(psOpeCod, "D", "3")
lsCodCtaCont = Left(lsCodCtaCont, 2) & lsMoneda & Mid(lsCodCtaCont, 4, 22)
Set oOpe = Nothing
Dim oCaja As nCajaGeneral
Set oCaja = New nCajaGeneral
BarraShow 1
Set rs = oCaja.GetDatosCartaFianza(lsCodCtaCont, "0")
lnTotalCartaFianza = rs.RecordCount
BarraClose
BarraShow rs.RecordCount
Do While Not rs.EOF
   BarraProgress rs.Bookmark, "REPORTE ESTADISTICO DIARIO", "CARTAS FIANZA", "Procesando...", vbBlue
   lnMontoCartaFianza = lnMontoCartaFianza + IIf(IsNull(rs!nMovImporte), 0, rs!nMovImporte)
   rs.MoveNext
Loop
BarraClose
RSClose rs
Set oCaja = Nothing
End Sub

Private Sub CalculoFondoExigible(lsMoneda As String, pdFecha As Date)
Dim lnBaseAgencia() As Currency
Dim lnDiferencia() As Currency
Dim lnTotalDiferencia As Currency
Dim I As Integer
ReDim lnFondoExigible(1)
If lsMoneda = "2" Then
    ReDim lnBaseAgencia(1)
    gnEncajeExig = EncajeBaseAgencia("E", lsMoneda, pdFecha)
    gnTotalOblig = EncajeBaseAgencia("T", lsMoneda, pdFecha)
    If gnTotalOblig <> 0 Then
        lnToseBase = Format(gnEncajeExig / gnTotalOblig, "#0.000000")
        If CDate(pdFecha) >= "01/09/2000" And CDate(pdFecha) <= "30/11/2000" Then
            lnToseBase = lnToseBase - 0.03
        End If
    Else
        lnToseBase = 0
    End If

    For I = 1 To UBound(lnEncajeBase)
        ReDim Preserve lnBaseAgencia(I)
        lnBaseAgencia(I) = lnEncajeBase(I) * lnToseBase
    Next I
    lnTotalDiferencia = 0
    ReDim lnDiferencia(1)
    For I = 1 To UBound(lnEncajeBase)
        ReDim Preserve lnDiferencia(I)
        lnDiferencia(I) = lnTotalCaptaciones(I) - lnEncajeBase(I)
        lnTotalDiferencia = lnTotalDiferencia + lnDiferencia(I)
    Next I

    If lnTotalDiferencia > 0 Then
        For I = 1 To UBound(lnBaseAgencia)
            ReDim Preserve lnFondoExigible(I)
            lnFondoExigible(I) = lnDiferencia(I) * 0.2 + lnBaseAgencia(I)
        Next I
    Else
        For I = 1 To UBound(lnBaseAgencia)
            ReDim Preserve lnFondoExigible(I)
            lnFondoExigible(I) = lnTotalCaptaciones(I) * lnToseBase
        Next I
    End If
Else
    lnPorcEncSoles = EncajeBaseAgencia("P", lsMoneda, pdFecha) / 100
    For I = 0 To UBound(lsSaldos, 2)
      ReDim Preserve lnFondoExigible(I)
      lnFondoExigible(I) = CCur(lsSaldos(5, I)) * lnPorcEncSoles
    Next I
End If
End Sub

Private Function GetInteresPF(psOpeCod As String, ByVal psMoneda As String, psAgeCod As String, ByVal pdFecha As Date) As Currency
    Dim oSdo As New NCtasaldo
    GetInteresPF = oSdo.GetOpeCtaSaldo(psOpeCod, Format(pdFecha, gsFormatoFecha), IIf(psMoneda = "1", True, False), "6", " AND RIGHT(CS.cCtaContCod,2) = '" & Right(psAgeCod, 2) & "' ")
    Set oSdo = Nothing
End Function

Private Function SaldosPromedioCaja(lsTipo As String, lsMoneda As String, ldFecha As Date) As Currency
Dim sql As String
Dim rs As New ADODB.Recordset
Dim ldFecIni As Date
Dim lsCodCtaCaja As String
Dim lnSaldo As Currency
Dim lnDiaMes As Integer
Set oCon = New DConecta
oCon.AbreConexion
ldFecIni = CDate("01/" & Format(Month(ldFecha), "00") & "/" & Format(Year(ldFecha), "0000"))

'Probado por Pepe en Trujillo

Dim oSdo As New NCtasaldo
Dim oOpe As New DOperacion
lsCodCtaCaja = oOpe.EmiteOpeCta(gsOpeCod, "D", "1")
lsCodCtaCaja = Left(lsCodCtaCaja, 2) & lsMoneda & Mid(lsCodCtaCaja, 4, 22)
Set oOpe = Nothing
BarraShow 1
BarraProgress 0, "REPORTE ESTADISTICO DIARIO", "BILLETAJE PROMEDIO CAJA GENERAL", "Procesando...", vbBlue
Set rs = oSdo.GetCtaSaldoRango(lsCodCtaCaja, DateAdd("m", -1, ldFecIni), ldFecIni - 1)
lnSaldo = 0
lnDiaMes = Day(ldFecIni - 1)
BarraShow rs.RecordCount
Do While Not rs.EOF
    BarraProgress rs.Bookmark, "REPORTE ESTADISTICO DIARIO", "BILLETAJE PROMEDIO CAJA GENERAL", "Procesando...", vbBlue
   'lnDiaMes = lnDiaMes + 1
   lnSaldo = lnSaldo + CCur(IIf(lsMoneda = "1", rs!nCtaSaldoImporte, rs!nCtaSaldoImporteme))
   rs.MoveNext
Loop
BarraClose

RSClose rs
oCon.CierraConexion
Set oCon = Nothing

If lnDiaMes = 0 Then
    SaldosPromedioCaja = 0
Else
    SaldosPromedioCaja = Round(lnSaldo / lnDiaMes, 2)
End If

End Function

Function CajasPromedioAgencias(lsMoneda As String, Fecha As Date, lsCodAge As String) As Currency
Dim lsQryEnc As String
Dim rsQryEnc As ADODB.Recordset
Dim ldFecIni As Date
Dim lnSaldo  As Currency
Dim lnDiaMes As Currency
Set oCon = New DConecta

oCon.AbreConexion
ldFecIni = CDate("01/" & Format(Month(Fecha), "00") & "/" & Format(Year(Fecha), "0000"))
CajasPromedioAgencias = 0

If DateAdd("m", -1, ldFecIni) = "01/09/2006" Then
    lsQryEnc = ""
    lsQryEnc = "SELECT Fec.dFecha, SubString(m.cMovNro,18,2) cAgeCod, SUM(nMonto) nMonto " _
          & "FROM   MovUserEfectivo MUE JOIN Mov M ON MUE.nMovNro = M.nMovNro, " _
          & "       FechaTmp('" & Format(DateAdd("m", -1, ldFecIni), gsFormatoFecha) & "')  Fec " _
          & "WHERE  M.nMovEstado = " & gMovEstContabNoContable & " and M.nMovFlag = " & gMovFlagVigente & " and Fec.dFecha BETWEEN '" & Format(DateAdd("m", -1, ldFecIni), gsFormatoFecha) & "' and '" & Format(ldFecIni - 1, gsFormatoFecha) & "' and " _
          & "       LEFT(m.cMovNro,8) = (SELECT Max( LEFT(cMovNro,8)) FROM Mov m1 JOIN MovUserEfectivo ME1 ON me1.nMovNro = m1.nMovNro WHERE M1.nMovEstado = " & gMovEstContabNoContable & " and M1.nMovFlag = " & gMovFlagVigente & " and LEFT(me1.cEfectivoCod,1) = " & lsMoneda & " and m1.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and SubString(m1.cMovNro,18,2) =  '" & lsCodAge & "' and LEFT(m1.cMovNro,8) <= Convert(varchar(10), Fec.dFecha,112) ) and " _
          & "       LEFT(mue.cEfectivoCod,1) = " & lsMoneda & " and m.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and " _
          & "       SubString(m.cMovNro,18,2) = '" & lsCodAge & "' " _
          & "GROUP BY Fec.dFecha, SubString(m.cMovNro,18,2) " _
          & "ORDER BY Fec.dFecha"
Else
    lsQryEnc = "SELECT Fec.dFecha, SubString(m.cMovNro,18,2) cAgeCod, SUM(nMonto) nMonto " _
          & "FROM   MovUserEfectivo MUE JOIN Mov M ON MUE.nMovNro = M.nMovNro, " _
          & "       FechaTmp('" & Format(DateAdd("m", -1, ldFecIni), gsFormatoFecha) & "')  Fec " _
          & "WHERE  M.nMovEstado = " & gMovEstContabNoContable & " and M.nMovFlag = " & gMovFlagVigente & " and Fec.dFecha BETWEEN '" & Format(DateAdd("m", -1, ldFecIni), gsFormatoFecha) & "' and '" & Format(ldFecIni - 1, gsFormatoFecha) & "' and " _
          & "       LEFT(m.cMovNro,8) = (SELECT Max( LEFT(cMovNro,8)) FROM Mov m1 JOIN MovUserEfectivo ME1 ON me1.nMovNro = m1.nMovNro WHERE M1.nMovEstado = " & gMovEstContabNoContable & " and M1.nMovFlag = " & gMovFlagVigente & " and LEFT(me1.cEfectivoCod,1) = " & lsMoneda & " and m1.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and SubString(m1.cMovNro,18,2) =  '" & lsCodAge & "' and LEFT(m1.cMovNro,8) <= Convert(varchar(10), Fec.dFecha,112) ) and " _
          & "       LEFT(mue.cEfectivoCod,1) = " & lsMoneda & " and m.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and " _
          & "       SubString(m.cMovNro,18,2) = '" & lsCodAge & "' " _
          & "GROUP BY Fec.dFecha, SubString(m.cMovNro,18,2) " _
          & "ORDER BY Fec.dFecha"
End If
         
   Set rsQryEnc = oCon.CargaRecordSet(lsQryEnc)
    lnSaldo = 0
    lnDiaMes = Day(ldFecIni - 1)
    If Not rsQryEnc.EOF Then
        Do While Not rsQryEnc.EOF
           'lnDiaMes = lnDiaMes + 1
           lnSaldo = lnSaldo + rsQryEnc!nMonto
           rsQryEnc.MoveNext
        Loop
        
        CajasPromedioAgencias = Round(lnSaldo / lnDiaMes, 2)
    End If
RSClose rsQryEnc
oCon.CierraConexion
Set oCon = Nothing
End Function

Private Sub Form_Initialize()
Set oCon = New DConecta
oCon.AbreConexion
End Sub

Private Sub Form_Load()
CentraForm Me
Set oCon = New DConecta
oCon.AbreConexion
End Sub

Private Sub Form_Terminate()
oCon.CierraConexion
Set oCon = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
oCon.CierraConexion
Set oCon = Nothing
End Sub

Private Sub BarraClose()
oBarra.CloseForm Me
Set oBarra = Nothing
End Sub

Private Sub BarraProgress(value As Variant, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
oBarra.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Sub

Private Sub BarraShow(pnMax As Variant)
Set oBarra = New clsProgressBar
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.ShowForm Me
oBarra.Max = pnMax
End Sub





