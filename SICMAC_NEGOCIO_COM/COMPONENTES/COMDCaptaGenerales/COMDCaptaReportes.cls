VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COMDCaptaReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Actualiza en Base la AFP
Option Base 0
Option Explicit

Private lsServerComunes As String
Private lsServerPersona As String
Private lsServerAdministracion As String
Private lsServerNegocio As String
Private lsServerImagenes As String
Dim oError As COMConecta.COMErrorHandling

'##ModelId=3AB902F20151
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902F20114
Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

Public Sub inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub

'Agrega una nueva APF
'##ModelId=3A969FF6015C
Public Function AgregaAFP(psRHAFPPersCod As String, psRHAFPAbreviatura As String, pnRHAFPPrima As Double, pnRHAFPVariable As Double, psUltimaActualizacion As String) As Boolean
    
    On Error GoTo AgregaAFPErr

    'your code goes here...
    
    Exit Function
AgregaAFPErr:
    Set oError = New COMConecta.COMErrorHandling
        Call oError.RaiseError(oError.MyUnhandledError, "DActualizaDatosAFP:AgregaAFP Method")
    Set oError = Nothing
End Function

'Devuelve los Valores de las tasas de descuento de la AFP
'##ModelId=3A96A04300D0
Public Function GetCaptaReportes(pbRHAFPPrima As Boolean, psRHAFPPersCod As String) As ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim sqlA As String
    Dim rsA As ADODB.Recordset
    Set rsA = New ADODB.Recordset
     Set oError = New COMConecta.COMErrorHandling
    On Error GoTo GetValorAFPErr

    If oCon.AbreConexion() Then
        If pbRHAFPPrima Then
            sqlA = " Select nRHAFPPrima as Monto From " & lsServerAdministracion & "RHAFP where cRHAFPPersCod  = '" & psRHAFPPersCod & "'"
        Else
            sqlA = " Select nRHAFPVariable as Monto From " & lsServerAdministracion & "RHAFP where cRHAFPPersCod  = '" & psRHAFPPersCod & "'"
        End If
        
        Set rsA = oCon.CargaRecordSet(sqlA)
        Set GetCaptaReportes = rsA
    End If
    
    rsA.Close
    Set rsA = Nothing
    Set oCon = Nothing
    Exit Function
GetValorAFPErr:
    Set oError = New COMConecta.COMErrorHandling
        Call oError.RaiseError(oError.MyUnhandledError, "DActualizaDatosAFP:GetValorAFP Method")
    Set oError = Nothing
End Function
'----------TODO COMPLETA-----------
'Devuelve los Valores de las tasas de descuento de la AFP
'##ModelId=3A96A04300D0
'Public Function GetCreditosVinculados(pbTitulares As Boolean, pbFinMes As Boolean, psCadenaPersona As String, pnTipoCambioFijo As Currency, psAgenciaCod As String) As ADODB.Recordset
'    Dim oCon As COMConecta.DCOMConecta
'    Set oCon = New COMConecta.DCOMConecta
'    Dim sql As String
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'
'    On Error GoTo GetValorAFPErr
'
'        If pbFinMes Then
'            oCon.AbreConexionRemota "07", , , "03"
'
'            If pbTitulares Then
'
'                sql = " Select cPersCod, Sum(Desem) Desem, Sum(Saldo) Saldo" _
'                    & "    From  (" _
'                    & "        Select cPersCod, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoDesemb * " & pnTipoCambioFijo & " Else nMontoDesemb End) Desem," _
'                    & "                 Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldoCap * " & pnTipoCambioFijo & " Else nSaldoCap End) Saldo From CreditoConsol CR" _
'                    & "        Inner Join ProductoPersonaConsol PC On CR.cCtaCod = PC.cCtaCod And Substring(PC.cCtaCod,6,3) <> '" & Producto.gColConsuPrendario & "' And PC.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
'                    & "             Where CR.nPrdEstado = " & gColocEstVigNorm & " And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & gColRelPersTitular & "" _
'                    & "         Group By cPersCod" _
'                    & "         Union" _
'                    & "         Select cPersCod, Sum(CR.nMontoDesemb) Desem," _
'                    & "                 Sum(nSaldoCap) Saldo       From CreditoConsol CR" _
'                    & "         Inner Join ProductoPersonaConsol PC On CR.cCtaCod = PC.cCtaCod And Substring(PC.cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "' And PC.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
'                    & "             Where CR.nPrdEstado In (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & gColRelPersTitular & "" _
'                    & "         Group By cPersCod )" _
'                    & "         As Creditos Group By cPersCod"
'
'                'sql = " Select cCodPers, Sum(Desem) Desem, Sum(Saldo) Saldo From " _
'                    & " ( Select cCodPers, Sum(Case Substring(CR.cCodCta,6,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoDesemb * " & pnTipoCambioFijo & " Else nMontoDesemb End) Desem, Sum(Case Substring(CR.cCodCta,6,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldoCap * " & pnTipoCambioFijo & " Else nSaldoCap End) Saldo From CreditoConsol CR" _
'                    & "     Inner Join PersCreditoConsol PC On CR.cCodCta = PC.cCodCta And Substring(PC.cCodCta,3,3) <> '305' And PC.cCodCta Like '" & Right(psAgenciaCod, 2) & "%'" _
'                    & "   Where CR.cEstado = 'F' And PC.cCodPers In (" & psCadenaPersona & ") And PC.cRelaCta = 'TI' " _
'                    & "   Group By cCodPers" _
'                    & "         Union" _
'                    & "   Select cCodPers, Sum(CR.nMontoDesemb) Desem, Sum(nSaldoCap) Saldo From CreditoConsol CR" _
'                    & "     Inner Join PersCreditoConsol PC On CR.cCodCta = PC.cCodCta And Substring(PC.cCodCta,3,3) = '305' And PC.cCodCta Like '" & Right(psAgenciaCod, 2) & "%'" _
'                    & "   Where CR.cEstado In ('1','4','6','7') And PC.cCodPers In (" & psCadenaPersona & ")" _
'                    & "   Group By cCodPers" _
'                    & " ) As Creditos Group By cCodPers"
'            Else
'                sql = "  Select cPersCod, Sum(Desem) Desem, Sum(Saldo) Saldo " _
'                    & "     From  ( " _
'                    & "         Select cPersCod, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoDesemb * " & pnTipoCambioFijo & " Else nMontoDesemb End) Desem, " _
'                    & "                  Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldoCap * " & pnTipoCambioFijo & " Else nSaldoCap End) Saldo " _
'                    & "             From CreditoConsol CR " _
'                    & "             Inner Join ProductoPersonaConsol PC On CR.cCtaCod = PC.cCtaCod And Substring(PC.cCtaCod,6,3) <> '" & Producto.gColConsuPrendario & "' And PC.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%' " _
'                    & "             Where CR.nPrdEstado = " & gColocEstVigNorm & " And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & gColRelPersTitular & "" _
'                    & "             Group By cPersCod " _
'                    & "             Union " _
'                    & "             Select cPersCod, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoDesemb * " & pnTipoCambioFijo & " Else nMontoDesemb End) Desem, " _
'                    & "                      Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldoCap * " & pnTipoCambioFijo & " Else nSaldoCap End) Saldo " _
'                    & "             From CreditoConsol CR " _
'                    & "             Inner Join ProductoPersonaConsol PC On CR.cCtaCod = PC.cCtaCod And Substring(PC.cCtaCod,6,3) <> '" & Producto.gColConsuPrendario & "' And PC.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%' " _
'                    & "             Where CR.nPrdEstado = 2020 And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & gColRelPersGarante & " And CR.cCtaCod Not In " _
'                    & "                 ( Select CR.cCtaCod  From CreditoConsol CR " _
'                    & "                     Inner Join ProductoPersonaConsol PC On CR.cCtaCod = PC.cCtaCod And PC.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%' " _
'                    & "                     Where CR.nPrdEstado = " & gColocEstVigNorm & " And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & gColRelPersTitular & ") " _
'                    & "             Group By cPersCod " _
'                    & "             Union " _
'                    & "             Select cPersCod, Sum(nMontoDesemb) Desem, Sum(nSaldoCap) Saldo " _
'                    & "             From CreditoConsol CR " _
'                    & "             Inner Join ProductoPersonaConsol PC On CR.cCtaCod = PC.cCtaCod  And Substring(PC.cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "'  And PC.nPrdPersRelac = " & gColRelPersTitular & " And PC.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%' " _
'                    & "             Where CR.nPrdEstado In (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") And PC.cPersCod In (" & psCadenaPersona & ") " _
'                    & "             Group By cPersCod " _
'                    & "     ) As Creditos Group By cPersCod"
'
'                'sql = " Select cCodPers, Sum(Desem) Desem, Sum(Saldo) Saldo From " _
'                    & " ( Select cCodPers, Sum(Case Substring(CR.cCodCta,6,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoDesemb * " & pnTipoCambioFijo & " Else nMontoDesemb End) Desem, Sum(Case Substring(CR.cCodCta,6,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldoCap * " & pnTipoCambioFijo & " Else nSaldoCap End) Saldo From CreditoConsol CR" _
'                    & "     Inner Join PersCreditoConsol PC On CR.cCodCta = PC.cCodCta And Substring(PC.cCodCta,3,3) <> '305' And PC.cCodCta Like '" & Right(psAgenciaCod, 2) & "%'" _
'                    & "   Where CR.cEstado = 'F' And PC.cCodPers In (" & psCadenaPersona & ") And PC.cRelaCta = 'TI' " _
'                    & "   Group By cCodPers" _
'                    & "         Union" _
'                    & "   Select cCodPers, Sum(Case Substring(CR.cCodCta,6,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoDesemb * " & pnTipoCambioFijo & " Else nMontoDesemb End) Desem, Sum(Case Substring(CR.cCodCta,6,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldoCap * " & pnTipoCambioFijo & " Else nSaldoCap End) Saldo From CreditoConsol CR" _
'                    & "        Inner Join PersCreditoConsol PC On CR.cCodCta = PC.cCodCta And Substring(PC.cCodCta,3,3) <> '305' And PC.cCodCta Like '" & Right(psAgenciaCod, 2) & "%'" _
'                    & "   Where CR.cEstado = 'F' And PC.cCodPers In (" & psCadenaPersona & ") And PC.cRelaCta = 'GA' And CR.cCodCta Not In" _
'                    & "         ( Select CR.cCodCta  From CreditoConsol CR" _
'                    & "                  Inner Join PersCreditoConsol PC On CR.cCodCta = PC.cCodCta And PC.cCodCta Like '" & Right(psAgenciaCod, 2) & "%'" _
'                    & "           Where CR.cEstado = 'F' And PC.cCodPers In (" & psCadenaPersona & ") And PC.cRelaCta = 'TI')" _
'                    & "   Group By cCodPers" _
'                    & "         Union" _
'                    & "   Select cCodPers, Sum(nMontoDesemb) Desem, Sum(nSaldoCap) Saldo From CreditoConsol CR" _
'                    & "     Inner Join PersCreditoConsol PC On CR.cCodCta = PC.cCodCta  And Substring(PC.cCodCta,3,3) = '305' And PC.cCodCta Like '" & Right(psAgenciaCod, 2) & "%'" _
'                    & "   Where CR.cEstado In ('1','4','6','7') And PC.cCodPers In (" & psCadenaPersona & ")" _
'                    & "   Group By cCodPers" _
'                    & " ) As Creditos Group By cCodPers"
'            End If
'        Else
'            oCon.AbreConexion
'            If pbTitulares Then
'                sql = " Select cPersCod, Sum(Desem) Desem, Sum(Saldo) Saldo From  " _
'                    & "     (" _
'                    & "        Select cPersCod , Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoCol * " & pnTipoCambioFijo & " Else nMontoCol End) Desem, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then PRD.nSaldo * " & pnTipoCambioFijo & " Else PRD.nSaldo End) Saldo" _
'                    & "         From Colocaciones CR" _
'                    & "         Inner Join Producto PRD ON CR.cCtaCod = PRD.cCtaCod And PRD.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
'                    & "         Inner Join ProductoPersona PC On CR.cCtaCod = PC.cCtaCod" _
'                    & "         Where Left(PRD.nPrdEstado,3) = '" & Left(ColocEstado.gColocEstVigNorm, 3) & "' And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & ColocRelacPers.gColRelPersTitular & "  " _
'                    & "     Group By cPersCod" _
'                    & "     Union" _
'                    & "         Select cPersCod , Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoCol * " & pnTipoCambioFijo & " Else nMontoCol End) Desem, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then PRD.nSaldo * " & pnTipoCambioFijo & " Else PRD.nSaldo End) Saldo" _
'                    & "           From Colocaciones CR" _
'                    & "           Inner Join Producto PRD ON CR.cCtaCod = PRD.cCtaCod And PRD.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
'                    & "           Inner Join ProductoPersona PC On CR.cCtaCod = PC.cCtaCod" _
'                    & "           Where PRD.nPrdEstado In ('" & ColocEstado.gColPEstDesem & "','" & ColocEstado.gColPEstVenci & "','" & ColocEstado.gColPEstPRema & "','" & ColocEstado.gColPEstRenov & "') And PC.cPersCod In (" & psCadenaPersona & ")" _
'                    & "           Group By cPersCod ) As Creditos Group By cPersCod"
'            Else
'                sql = " Select cPersCod, Sum(Desem) Desem, Sum(Saldo) Saldo From  " _
'                    & "     (" _
'                    & "       Select cPersCod, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoCol * " & pnTipoCambioFijo & " Else nMontoCol End) Desem, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldo * " & pnTipoCambioFijo & " Else nSaldo End) Saldo" _
'                    & "         From Colocaciones CR" _
'                    & "         Inner Join ProductoPersona PC On CR.cCtaCod = PC.cCtaCod" _
'                    & "         Inner Join Producto PRD On CR.cCtaCod = PRD.cCtaCod And PRD.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
'                    & "         Where Left(PRD.nPrdEstado,3) = '" & Left(ColocEstado.gColocEstVigNorm, 3) & "' And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & ColocRelacPers.gColRelPersTitular & "" _
'                    & "         Group By cPersCod" _
'                    & "         Union" _
'                    & "       Select cPersCod, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoCol * " & pnTipoCambioFijo & " Else nMontoCol End) Desem, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldo * " & pnTipoCambioFijo & " Else nSaldo End) Saldo" _
'                    & "         From Colocaciones CR" _
'                    & "         Inner Join ProductoPersona PC On CR.cCtaCod = PC.cCtaCod" _
'                    & "         Inner Join Producto PRD On CR.cCtaCod = PRD.cCtaCod And PRD.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
'                    & "         Where Left(PRD.nPrdEstado,3) = '" & Left(ColocEstado.gColocEstVigNorm, 3) & "' And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & ColocRelacPers.gColRelPersGarante & " And CR.cCtaCod Not In" _
'                    & "             ( Select CR.cCtaCod  From Colocaciones CR" _
'                    & "                 Inner Join ProductoPersona PC On CR.cCtaCod = PC.cCtaCod" _
'                    & "                 Inner Join Producto PRD On CR.cCtaCod = PRD.cCtaCod And PRD.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
'                    & "         Where Left(PRD.nPrdEstado,3) = '" & Left(ColocEstado.gColocEstVigNorm, 3) & "' And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & ColocRelacPers.gColRelPersTitular & ")" _
'                    & "         Group By cPersCod" _
'                    & "         Union" _
'                    & "         Select cPersCod , Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoCol * " & pnTipoCambioFijo & " Else nMontoCol End) Desem, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then PRD.nSaldo * " & pnTipoCambioFijo & " Else PRD.nSaldo End) Saldo" _
'                    & "           From Colocaciones CR" _
'                    & "           Inner Join Producto PRD ON CR.cCtaCod = PRD.cCtaCod And PRD.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
'                    & "           Inner Join ProductoPersona PC On CR.cCtaCod = PC.cCtaCod" _
'                    & "           Where PRD.nPrdEstado In ('" & ColocEstado.gColPEstDesem & "','" & ColocEstado.gColPEstVenci & "','" & ColocEstado.gColPEstPRema & "','" & ColocEstado.gColPEstRenov & "') And PC.cPersCod In (" & psCadenaPersona & ") Group By cPersCod ) As Creditos Group By cPersCod"
'            End If
'        End If
'
'        Set rs = oCon.CargaRecordSet(sql)
'        Set GetCreditosVinculados = rs
'
'    oCon.CierraConexion
'    Set oCon = Nothing
'    Exit Function
'GetValorAFPErr:
'    Set oError = New COMConecta.COMErrorHandling
'        Call oError.RaiseError(oError.MyUnhandledError, "DActualizaDatosAFP:GetValorAFP Method")
'    Set oError = Nothing
'End Function

'Devuelve los Valores de las tasas de descuento de la AFP
'##ModelId=3A96A04300D0
Public Function GetFamiliares() As ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    On Error GoTo GetValorAFPErr

    If oCon.AbreConexion() Then
        sql = " Select cRHCod, cPersCod, NEmp, NFam, cNomTab From EmpleadoFamilia  Order By NEmp, NFam"

        Set rs = oCon.CargaRecordSet(sql)
        Set GetFamiliares = rs
    End If

    Set oCon = Nothing
    Exit Function
GetValorAFPErr:
    Set oError = New COMConecta.COMErrorHandling
        Call oError.RaiseError(oError.MyUnhandledError, "DActualizaDatosAFP:GetValorAFP Method")
    Set oError = Nothing
End Function

'Public Function Cap_ReportesCaptaciones() As ADODB.Recordset 'comentado por pti1 ACTA N° 034-2019
Public Function Cap_ReportesCaptaciones(Optional ByVal cUser As String, Optional ByVal cGrupo As String) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim sql As String
    Dim rs As New ADODB.Recordset
    oCon.AbreConexion
    'sql = "exec  Cap_ReportesCaptaciones 'PTI1 20-03-2019 COMENTADO ACTA N° 034-2019
    sql = "exec  Cap_ReportesCaptaciones_Usuario '" & cUser & "','" & cGrupo & "'" 'ADD PTI1 20-03-2019 ACTA N° 034-2019
    
    Set rs = oCon.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.BOF) Then
        Set Cap_ReportesCaptaciones = rs
    Else
        Set Cap_ReportesCaptaciones = Nothing
    End If
End Function

Public Function GetAperturaPromotorAho(ByVal pcProd As String, ByVal pcProm As String, ByVal pnMoneda As Integer, ByVal pdFecI As Date, ByVal pdFecF As Date) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    
    Select Case pcProd
    Case "232"
        sSql = "Select Cuenta = PP.cCtaCod, Cliente = PPC.cPersNombre, " & _
               "nSaldo = P.nSaldo, nApertura = M.nMonto, dApertura = C.dApertura, " & _
               "nPlazo = CA.nPlazo " & _
               "From ProductoPersona PP " & _
               "Inner Join Persona PE On PP.cPersCod = PE.cPersCod " & _
               "Inner Join Captaciones C On PP.cCtaCod = C.cCtaCod " & _
               "Inner Join Producto P On PP.cCtaCod = P.cCtaCod " & _
               "Inner Join MovCapDet M On PP.cCtaCod = M.cCtaCod " & _
               "Inner Join ProductoPersona PC On PP.cCtaCod = PC.cCtaCod " & _
               "Inner Join Persona PPC On PC.cPersCod = PPC.cPersCod " & _
               "Inner Join CaptacAhorros CA On P.cCtaCod = CA.cCtaCod " & _
               "Where " & _
               "PP.nPrdPersRelac = " & gCapRelPersPromotor & " " & _
               "And M.cOpeCod In ('" & gAhoApeEfec & "','" & gAhoApeChq & "','" & gAhoApeTransf & "','" & gAhoApeLoteEfec & "','" & gAhoApeLoteChq & "') " & _
               "And PC.nPrdPersRelac = " & gCapRelPersTitular & " " & _
               "And PP.cPersCod = '" & pcProm & "' " & _
               "And substring(PP.cCtaCod,9,1) = " & pnMoneda & " " & _
               "And P.nPrdEstado Not In (1300,1400) " & _
               "And C.dApertura Between '" & Format(pdFecI, "yyyymmdd") & "' And '" & Format(pdFecF, "yyyymmdd") & "' "
    Case "233"
        sSql = "Select Cuenta = PP.cCtaCod, Cliente = PPC.cPersNombre, " & _
               "nSaldo = P.nSaldo, nApertura = CA.nApertura, dApertura = C.dApertura, " & _
               "nPlazo = CA.nPlazo " & _
               "From ProductoPersona PP " & _
               "Inner Join Persona PE On PP.cPersCod = PE.cPersCod " & _
               "Inner Join Captaciones C On PP.cCtaCod = C.cCtaCod " & _
               "Inner Join Producto P On PP.cCtaCod = P.cCtaCod " & _
               "Inner Join ProductoPersona PC On PP.cCtaCod = PC.cCtaCod " & _
               "Inner Join Persona PPC On PC.cPersCod = PPC.cPersCod " & _
               "Inner Join CaptacPlazoFijo CA On P.cCtaCod = CA.cCtaCod " & _
               "Where " & _
               "PP.nPrdPersRelac = " & gCapRelPersPromotor & " " & _
               "And PC.nPrdPersRelac = " & gCapRelPersTitular & " " & _
               "And PP.cPersCod = '" & pcProm & "' " & _
               "And substring(PP.cCtaCod,9,1) = " & pnMoneda & " " & _
               "And P.nPrdEstado Not In (1300,1400) " & _
               "And C.dApertura Between '" & Format(pdFecI, "yyyymmdd") & "' And '" & Format(pdFecF, "yyyymmdd") & "' "
    Case "234" 'By Capi Agosto 2007 Acta 014-2007
        sSql = "Select Cuenta = PP.cCtaCod, Cliente = PEC.cPersNombre, " & _
               "nSaldo = P.nSaldo,nTangible=Cts.nSaldRetiro,nIntangible=P.nSaldo-Cts.nSaldRetiro, " & _
               "nApertura = M.nMonto, dApertura = C.dApertura, " & _
               "nPlazo = 0 " & _
               "From ProductoPersona PP " & _
               "Inner Join Captaciones C On PP.cCtaCod = C.cCtaCod " & _
               "Inner join Producto P On PP.cCtaCod = P.cCtaCod " & _
               "Inner Join CaptacCTS Cts On PP.cCtaCod = Cts.cCtaCod " & _
               "Inner Join MovCapDet M On PP.cCtaCod = M.cCtaCod " & _
               "Inner Join ProductoPersona PC On PP.cCtaCod = PC.cCtaCod " & _
               "Inner Join Persona PEC On PC.cPersCod = PEC.cPersCod " & _
               "Where " & _
               "PP.NPRDPERSRELAC = " & gCapRelPersPromotor & " " & _
               "And M.cOpecod In ('" & gCTSApeEfec & " ','" & gCTSApeChq & "') " & _
               "And PP.cPersCod = '" & pcProm & "' " & _
               "And PC.nPrdPersRelac = " & gCapRelPersTitular & " " & _
               "And substring(PP.cCtaCod,9,1) = " & pnMoneda & " " & _
               "And P.nPrdEstado Not In (1300,1400) " & _
               "And C.dApertura Between '" & Format(pdFecI, "yyyymmdd") & "' And '" & Format(pdFecF, "yyyymmdd") & "' "
    End Select
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    Set GetAperturaPromotorAho = rs
End Function

'MADM 20101210
Public Function RecuperaRfxAgencia(ByVal psCodAge As String, Optional x As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim x1 As Integer
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaRfxAgencia
    x1 = IIf(x, 1, 0)
    sSql = " exec stp_sel_devuelveRfxAgencia '" & psCodAge & "'," & x1 & ""
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaRfxAgencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaRfxAgencia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'***Agregado por ELRO el 20120905, según OYP-RFC087-2012
Public Function obtenerIndConCli() As ADODB.Recordset
    Dim oDCOMConecta As New COMConecta.DCOMConecta
    Dim sSql As String
    
    oDCOMConecta.AbreConexion
    sSql = "exec  stp_sel_RFC0872012_DevolverIndConCli"
    
    Set obtenerIndConCli = oDCOMConecta.CargaRecordSet(sSql)

    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
End Function

Public Function actualizarIndConCli(ByVal pnId_IndConCli As Long, _
                                    ByVal pnSBS As Currency, _
                                    ByVal pnInterno As Currency) As Long
    Dim oDCOMConecta As New COMConecta.DCOMConecta
    Dim sSql As String
    Dim rsIndConCli As New ADODB.Recordset
    
    oDCOMConecta.AbreConexion
    sSql = "exec stp_upd_RFC0872012_ActualizarIndConCli " & pnId_IndConCli & ", " & pnSBS & ", " & pnInterno & ""
    
    Set rsIndConCli = oDCOMConecta.CargaRecordSet(sSql)
    
    If Not (rsIndConCli.EOF And rsIndConCli.BOF) Then
        actualizarIndConCli = rsIndConCli!nId_IndConCli
    Else
        actualizarIndConCli = 0
    End If
    
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
End Function

Public Function obtenerIndComDep() As ADODB.Recordset
    Dim oDCOMConecta As New COMConecta.DCOMConecta
    Dim sSql As String
    
    oDCOMConecta.AbreConexion
    sSql = "exec  stp_sel_RFC0872012_DevolverIndComDep"
    
    Set obtenerIndComDep = oDCOMConecta.CargaRecordSet(sSql)
    
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
End Function

Public Function actualizarIndComDep(ByVal pnId_IndConCli As Long, ByVal pnIndice As Currency) As Long
    Dim oDCOMConecta As New COMConecta.DCOMConecta
    Dim sSql As String
    Dim rsIndComDep As New ADODB.Recordset
    
    oDCOMConecta.AbreConexion
    sSql = "exec stp_upd_RFC0872012_ActualizarIndComDep " & pnId_IndConCli & ", " & pnIndice & ""
    
    Set rsIndComDep = oDCOMConecta.CargaRecordSet(sSql)
    
    If Not (rsIndComDep.EOF And rsIndComDep.BOF) Then
        actualizarIndComDep = rsIndComDep!nId_IndComDep
    Else
        actualizarIndComDep = 0
    End If
    
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
End Function

Public Function obtenerAperturaVoBo(ByVal psAgeCod As String) As ADODB.Recordset
    Dim oDCOMConecta As New COMConecta.DCOMConecta
    Dim sSql As String
    
    oDCOMConecta.AbreConexion
    sSql = "exec  stp_sel_REF0872012_DevolverAperturaVoBo '" & psAgeCod & "' "
    Set obtenerAperturaVoBo = oDCOMConecta.CargaRecordSet(sSql)
    
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
End Function

Public Function actualizarVisBueCapIndConCli(ByVal pnId_VoBoConCli As Long, ByVal pnEstado As Integer, ByVal psMov As String) As Long
    Dim oDCOMConecta As New COMConecta.DCOMConecta
    Dim sSql As String
    Dim rsVoBo As New ADODB.Recordset
    
    oDCOMConecta.AbreConexion
    sSql = "exec stp_sel_REF0872012_ActualizarVisBueCapIndConCli " & pnId_VoBoConCli & ", " & pnEstado & ", '" & psMov & "'"
    
    Set rsVoBo = oDCOMConecta.CargaRecordSet(sSql)
    
    If Not (rsVoBo.EOF And rsVoBo.BOF) Then
        actualizarVisBueCapIndConCli = rsVoBo!Id_VoBoConCli
    Else
        actualizarVisBueCapIndConCli = 0
    End If
    
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
End Function
'***Fin Agregado por ELRO el 20120905*******************
