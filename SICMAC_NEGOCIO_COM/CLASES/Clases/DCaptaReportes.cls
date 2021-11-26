VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCaptaReportes"
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
    Call RaiseError(MyUnhandledError, "DActualizaDatosAFP:AgregaAFP Method")
End Function

'Devuelve los Valores de las tasas de descuento de la AFP
'##ModelId=3A96A04300D0
Public Function GetCaptaReportes(pbRHAFPPrima As Boolean, psRHAFPPersCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlA As String
    Dim rsA As ADODB.Recordset
    Set rsA = New ADODB.Recordset
    
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
    Call RaiseError(MyUnhandledError, "DActualizaDatosAFP:GetValorAFP Method")
End Function

'Devuelve los Valores de las tasas de descuento de la AFP
'##ModelId=3A96A04300D0
Public Function GetCreditosVinculados(pbTitulares As Boolean, pbFinMes As Boolean, psCadenaPersona As String, pnTipoCambioFijo As Currency, psAgenciaCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    On Error GoTo GetValorAFPErr

        If pbFinMes Then
            oCon.AbreConexionRemota "07", , , "03"
            
            If pbTitulares Then
            
                sql = " Select cPersCod, Sum(Desem) Desem, Sum(Saldo) Saldo" _
                    & "    From  (" _
                    & "        Select cPersCod, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoDesemb * " & pnTipoCambioFijo & " Else nMontoDesemb End) Desem," _
                    & "                 Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldoCap * " & pnTipoCambioFijo & " Else nSaldoCap End) Saldo From CreditoConsol CR" _
                    & "        Inner Join ProductoPersonaConsol PC On CR.cCtaCod = PC.cCtaCod And Substring(PC.cCtaCod,6,3) <> '" & Producto.gColConsuPrendario & "' And PC.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
                    & "             Where CR.nPrdEstado = " & gColocEstVigNorm & " And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & gColRelPersTitular & "" _
                    & "         Group By cPersCod" _
                    & "         Union" _
                    & "         Select cPersCod, Sum(CR.nMontoDesemb) Desem," _
                    & "                 Sum(nSaldoCap) Saldo       From CreditoConsol CR" _
                    & "         Inner Join ProductoPersonaConsol PC On CR.cCtaCod = PC.cCtaCod And Substring(PC.cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "' And PC.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
                    & "             Where CR.nPrdEstado In (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & gColRelPersTitular & "" _
                    & "         Group By cPersCod )" _
                    & "         As Creditos Group By cPersCod"

                'sql = " Select cCodPers, Sum(Desem) Desem, Sum(Saldo) Saldo From " _
                    & " ( Select cCodPers, Sum(Case Substring(CR.cCodCta,6,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoDesemb * " & pnTipoCambioFijo & " Else nMontoDesemb End) Desem, Sum(Case Substring(CR.cCodCta,6,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldoCap * " & pnTipoCambioFijo & " Else nSaldoCap End) Saldo From CreditoConsol CR" _
                    & "     Inner Join PersCreditoConsol PC On CR.cCodCta = PC.cCodCta And Substring(PC.cCodCta,3,3) <> '305' And PC.cCodCta Like '" & Right(psAgenciaCod, 2) & "%'" _
                    & "   Where CR.cEstado = 'F' And PC.cCodPers In (" & psCadenaPersona & ") And PC.cRelaCta = 'TI' " _
                    & "   Group By cCodPers" _
                    & "         Union" _
                    & "   Select cCodPers, Sum(CR.nMontoDesemb) Desem, Sum(nSaldoCap) Saldo From CreditoConsol CR" _
                    & "     Inner Join PersCreditoConsol PC On CR.cCodCta = PC.cCodCta And Substring(PC.cCodCta,3,3) = '305' And PC.cCodCta Like '" & Right(psAgenciaCod, 2) & "%'" _
                    & "   Where CR.cEstado In ('1','4','6','7') And PC.cCodPers In (" & psCadenaPersona & ")" _
                    & "   Group By cCodPers" _
                    & " ) As Creditos Group By cCodPers"
            Else
                sql = "  Select cPersCod, Sum(Desem) Desem, Sum(Saldo) Saldo " _
                    & "     From  ( " _
                    & "         Select cPersCod, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoDesemb * " & pnTipoCambioFijo & " Else nMontoDesemb End) Desem, " _
                    & "                  Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldoCap * " & pnTipoCambioFijo & " Else nSaldoCap End) Saldo " _
                    & "             From CreditoConsol CR " _
                    & "             Inner Join ProductoPersonaConsol PC On CR.cCtaCod = PC.cCtaCod And Substring(PC.cCtaCod,6,3) <> '" & Producto.gColConsuPrendario & "' And PC.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%' " _
                    & "             Where CR.nPrdEstado = " & gColocEstVigNorm & " And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & gColRelPersTitular & "" _
                    & "             Group By cPersCod " _
                    & "             Union " _
                    & "             Select cPersCod, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoDesemb * " & pnTipoCambioFijo & " Else nMontoDesemb End) Desem, " _
                    & "                      Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldoCap * " & pnTipoCambioFijo & " Else nSaldoCap End) Saldo " _
                    & "             From CreditoConsol CR " _
                    & "             Inner Join ProductoPersonaConsol PC On CR.cCtaCod = PC.cCtaCod And Substring(PC.cCtaCod,6,3) <> '" & Producto.gColConsuPrendario & "' And PC.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%' " _
                    & "             Where CR.nPrdEstado = 2020 And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & gColRelPersGarante & " And CR.cCtaCod Not In " _
                    & "                 ( Select CR.cCtaCod  From CreditoConsol CR " _
                    & "                     Inner Join ProductoPersonaConsol PC On CR.cCtaCod = PC.cCtaCod And PC.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%' " _
                    & "                     Where CR.nPrdEstado = " & gColocEstVigNorm & " And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & gColRelPersTitular & ") " _
                    & "             Group By cPersCod " _
                    & "             Union " _
                    & "             Select cPersCod, Sum(nMontoDesemb) Desem, Sum(nSaldoCap) Saldo " _
                    & "             From CreditoConsol CR " _
                    & "             Inner Join ProductoPersonaConsol PC On CR.cCtaCod = PC.cCtaCod  And Substring(PC.cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "'  And PC.nPrdPersRelac = " & gColRelPersTitular & " And PC.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%' " _
                    & "             Where CR.nPrdEstado In (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") And PC.cPersCod In (" & psCadenaPersona & ") " _
                    & "             Group By cPersCod " _
                    & "     ) As Creditos Group By cPersCod"

                'sql = " Select cCodPers, Sum(Desem) Desem, Sum(Saldo) Saldo From " _
                    & " ( Select cCodPers, Sum(Case Substring(CR.cCodCta,6,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoDesemb * " & pnTipoCambioFijo & " Else nMontoDesemb End) Desem, Sum(Case Substring(CR.cCodCta,6,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldoCap * " & pnTipoCambioFijo & " Else nSaldoCap End) Saldo From CreditoConsol CR" _
                    & "     Inner Join PersCreditoConsol PC On CR.cCodCta = PC.cCodCta And Substring(PC.cCodCta,3,3) <> '305' And PC.cCodCta Like '" & Right(psAgenciaCod, 2) & "%'" _
                    & "   Where CR.cEstado = 'F' And PC.cCodPers In (" & psCadenaPersona & ") And PC.cRelaCta = 'TI' " _
                    & "   Group By cCodPers" _
                    & "         Union" _
                    & "   Select cCodPers, Sum(Case Substring(CR.cCodCta,6,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoDesemb * " & pnTipoCambioFijo & " Else nMontoDesemb End) Desem, Sum(Case Substring(CR.cCodCta,6,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldoCap * " & pnTipoCambioFijo & " Else nSaldoCap End) Saldo From CreditoConsol CR" _
                    & "        Inner Join PersCreditoConsol PC On CR.cCodCta = PC.cCodCta And Substring(PC.cCodCta,3,3) <> '305' And PC.cCodCta Like '" & Right(psAgenciaCod, 2) & "%'" _
                    & "   Where CR.cEstado = 'F' And PC.cCodPers In (" & psCadenaPersona & ") And PC.cRelaCta = 'GA' And CR.cCodCta Not In" _
                    & "         ( Select CR.cCodCta  From CreditoConsol CR" _
                    & "                  Inner Join PersCreditoConsol PC On CR.cCodCta = PC.cCodCta And PC.cCodCta Like '" & Right(psAgenciaCod, 2) & "%'" _
                    & "           Where CR.cEstado = 'F' And PC.cCodPers In (" & psCadenaPersona & ") And PC.cRelaCta = 'TI')" _
                    & "   Group By cCodPers" _
                    & "         Union" _
                    & "   Select cCodPers, Sum(nMontoDesemb) Desem, Sum(nSaldoCap) Saldo From CreditoConsol CR" _
                    & "     Inner Join PersCreditoConsol PC On CR.cCodCta = PC.cCodCta  And Substring(PC.cCodCta,3,3) = '305' And PC.cCodCta Like '" & Right(psAgenciaCod, 2) & "%'" _
                    & "   Where CR.cEstado In ('1','4','6','7') And PC.cCodPers In (" & psCadenaPersona & ")" _
                    & "   Group By cCodPers" _
                    & " ) As Creditos Group By cCodPers"
            End If
        Else
            oCon.AbreConexion
            If pbTitulares Then
                sql = " Select cPersCod, Sum(Desem) Desem, Sum(Saldo) Saldo From  " _
                    & "     (" _
                    & "        Select cPersCod , Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoCol * " & pnTipoCambioFijo & " Else nMontoCol End) Desem, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then PRD.nSaldo * " & pnTipoCambioFijo & " Else PRD.nSaldo End) Saldo" _
                    & "         From Colocaciones CR" _
                    & "         Inner Join Producto PRD ON CR.cCtaCod = PRD.cCtaCod And PRD.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
                    & "         Inner Join ProductoPersona PC On CR.cCtaCod = PC.cCtaCod" _
                    & "         Where Left(PRD.nPrdEstado,3) = '" & Left(ColocEstado.gColocEstVigNorm, 3) & "' And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & ColocRelacPers.gColRelPersTitular & "  " _
                    & "     Group By cPersCod" _
                    & "     Union" _
                    & "         Select cPersCod , Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoCol * " & pnTipoCambioFijo & " Else nMontoCol End) Desem, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then PRD.nSaldo * " & pnTipoCambioFijo & " Else PRD.nSaldo End) Saldo" _
                    & "           From Colocaciones CR" _
                    & "           Inner Join Producto PRD ON CR.cCtaCod = PRD.cCtaCod And PRD.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
                    & "           Inner Join ProductoPersona PC On CR.cCtaCod = PC.cCtaCod" _
                    & "           Where PRD.nPrdEstado In ('" & ColocEstado.gColPEstDesem & "','" & ColocEstado.gColPEstVenci & "','" & ColocEstado.gColPEstPRema & "','" & ColocEstado.gColPEstRenov & "') And PC.cPersCod In (" & psCadenaPersona & ")" _
                    & "           Group By cPersCod ) As Creditos Group By cPersCod"
            Else
                sql = " Select cPersCod, Sum(Desem) Desem, Sum(Saldo) Saldo From  " _
                    & "     (" _
                    & "       Select cPersCod, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoCol * " & pnTipoCambioFijo & " Else nMontoCol End) Desem, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldo * " & pnTipoCambioFijo & " Else nSaldo End) Saldo" _
                    & "         From Colocaciones CR" _
                    & "         Inner Join ProductoPersona PC On CR.cCtaCod = PC.cCtaCod" _
                    & "         Inner Join Producto PRD On CR.cCtaCod = PRD.cCtaCod And PRD.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
                    & "         Where Left(PRD.nPrdEstado,3) = '" & Left(ColocEstado.gColocEstVigNorm, 3) & "' And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & ColocRelacPers.gColRelPersTitular & "" _
                    & "         Group By cPersCod" _
                    & "         Union" _
                    & "       Select cPersCod, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoCol * " & pnTipoCambioFijo & " Else nMontoCol End) Desem, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nSaldo * " & pnTipoCambioFijo & " Else nSaldo End) Saldo" _
                    & "         From Colocaciones CR" _
                    & "         Inner Join ProductoPersona PC On CR.cCtaCod = PC.cCtaCod" _
                    & "         Inner Join Producto PRD On CR.cCtaCod = PRD.cCtaCod And PRD.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
                    & "         Where Left(PRD.nPrdEstado,3) = '" & Left(ColocEstado.gColocEstVigNorm, 3) & "' And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & ColocRelacPers.gColRelPersGarante & " And CR.cCtaCod Not In" _
                    & "             ( Select CR.cCtaCod  From Colocaciones CR" _
                    & "                 Inner Join ProductoPersona PC On CR.cCtaCod = PC.cCtaCod" _
                    & "                 Inner Join Producto PRD On CR.cCtaCod = PRD.cCtaCod And PRD.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
                    & "         Where Left(PRD.nPrdEstado,3) = '" & Left(ColocEstado.gColocEstVigNorm, 3) & "' And PC.cPersCod In (" & psCadenaPersona & ") And PC.nPrdPersRelac = " & ColocRelacPers.gColRelPersTitular & ")" _
                    & "         Group By cPersCod" _
                    & "         Union" _
                    & "         Select cPersCod , Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then nMontoCol * " & pnTipoCambioFijo & " Else nMontoCol End) Desem, Sum(Case Substring(CR.cCtaCod,9,1) When '" & Moneda.gMonedaExtranjera & "' Then PRD.nSaldo * " & pnTipoCambioFijo & " Else PRD.nSaldo End) Saldo" _
                    & "           From Colocaciones CR" _
                    & "           Inner Join Producto PRD ON CR.cCtaCod = PRD.cCtaCod And PRD.cCtaCod Like '___" & Right(psAgenciaCod, 2) & "%'" _
                    & "           Inner Join ProductoPersona PC On CR.cCtaCod = PC.cCtaCod" _
                    & "           Where PRD.nPrdEstado In ('" & ColocEstado.gColPEstDesem & "','" & ColocEstado.gColPEstVenci & "','" & ColocEstado.gColPEstPRema & "','" & ColocEstado.gColPEstRenov & "') And PC.cPersCod In (" & psCadenaPersona & ") Group By cPersCod ) As Creditos Group By cPersCod"
            End If
        End If
                
        Set rs = oCon.CargaRecordSet(sql)
        Set GetCreditosVinculados = rs
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
GetValorAFPErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosAFP:GetValorAFP Method")
End Function

'Devuelve los Valores de las tasas de descuento de la AFP
'##ModelId=3A96A04300D0
Public Function GetFamiliares() As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
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
    Call RaiseError(MyUnhandledError, "DActualizaDatosAFP:GetValorAFP Method")
End Function


#If DebugMode Then
    '##ModelId=3AB902F20056
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

