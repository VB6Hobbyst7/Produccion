VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCajeroImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim oImp As COMFunciones.FCOMVarImpresion

Public Function GeneraRepoCompraVenta(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                    ByVal psCodUser As String, ByVal psCodAge As String, _
                                    ByVal psNomCmact As String, ByVal psNomAge As String, ByVal psNomSeccion As String, _
                                    ByVal pdFecSis As Date, ByVal pnColPage As Integer, ByVal pnLinPage As Integer) As String
                              
Dim lsTexto As String
Dim Rs As ADODB.Recordset
Dim oCajero As NCOMCajero
Dim lsCabecera As String
Dim lnNumPage As Integer
Dim lnReg As Long
Dim lnTotal As Long
Dim lnTotalsoles As Currency
Dim lnTotalDolares As Currency
Dim lsAgencia As String
Dim oFunI As New COMFunciones.FCOMImpresion
Set oCajero = New NCOMCajero
Set Rs = New ADODB.Recordset
Select Case psOpeCod
    Case gOpeCajeroMECompra
        lsCabecera = "LISTADO DIARIO MONEDA EXTRANJERA - COMPRA"
    Case gOpeCajeroMEVenta
        lsCabecera = "LISTADO DIARIO MONEDA EXTRANJERA - VENTA" '
End Select
'If pdHasta < FechaTopeHistorica Then
'    Set rs = oCajero.GetCompraVentaRepo(psOpeCod, pdDesde, pdHasta, psCodUser, psCodAge, True)
'Else
    Set Rs = oCajero.GetCompraVentaRepo(psOpeCod, pdDesde, pdHasta, psCodUser, psCodAge)
'End If
lsAgencia = ""
lnReg = 0: lnTotalsoles = 0: lnTotalDolares = 0: lnNumPage = 0: lsTexto = ""
lnTotalsolesAgencia = 0: lnTotalDolaresAgencia = 0
If Not Rs.EOF And Not Rs.BOF Then
    lnTotal = Rs.RecordCount
   
    lsAgencia = Rs!Agencia
    oFunI.Linea lsTexto, CabeceraRepoCompraVenta(pdDesde, pdHasta, lsCabecera, psNomCmact, psNomAge, psNomSeccion, pdFecSis, pnColPage, pnLinPage, lnReg, lnNumPage, lsAgencia)
    Do While Not Rs.EOF
        lnReg = lnReg + 1
        oFunI.Linea lsTexto, Con + oFunI.ImpreFormat(Rs!nMovNro, 6, 0, True) + oFunI.ImpreFormat(Rs!Fecha, 12) + oFunI.ImpreFormat(Rs!cPersNombre, 30) + oFunI.ImpreFormat(Rs!Dolares, 12, 2, True) + _
                oFunI.ImpreFormat(Rs!Cambio, 10, 2, True) + oFunI.ImpreFormat(Rs!SOLES, 12, 2, True) & oFunI.ImpreFormat(Rs!cUser, 8, 5) & oFunI.ImpreFormat(Format(Rs!Hora, "hh:mm:ss AMPM"), 12) & oFunI.ImpreFormat(Rs!Agencia, 15) + COFF

        lnTotalsoles = lnTotalsoles + Rs!SOLES
        lnTotalDolares = lnTotalDolares + Rs!Dolares
        lnTotalsolesAgencia = lnTotalsolesAgencia + Rs!SOLES
        lnTotalDolaresAgencia = lnTotalDolaresAgencia + Rs!Dolares
        
        
        If lnReg > pnLinPage Then
            lnReg = 0
            lstext = lsTexto + Chr$(12)
            lnNumPage = lnNumPage + 1
            oFunI.Linea lsTexto, CabeceraRepoCompraVenta(pdDesde, pdHasta, psCabecera, psNomCmact, psNomAge, psNomSeccion, pdFecSis, pnColPage, pnLinPage, lnReg, lnNumPage, lsAgencia)
        End If
        Rs.MoveNext
        If Rs.EOF Then
            oFunI.Linea lsTexto, BON + String(pnColPage + 5, "-")
            oFunI.Linea lsTexto, Con + BON + oFunI.ImpreFormat("TOTAL AGENCIA", 20) + oFunI.ImpreFormat("", 28) + oFunI.ImpreFormat(lnTotalDolaresAgencia, 12, 2, True) + _
                oFunI.ImpreFormat("", 12, 2, True) + oFunI.ImpreFormat(lnTotalsolesAgencia, 11, 2, True) + COFF + BOFF, 2
        Else
            If lsAgencia <> Rs!Agencia Or Rs.EOF Then
                lsAgencia = Rs!Agencia
                oFunI.Linea lsTexto, BON + String(pnColPage + 5, "-")
                oFunI.Linea lsTexto, Con + BON + oFunI.ImpreFormat("TOTAL AGENCIA", 20) + oFunI.ImpreFormat("", 28) + oFunI.ImpreFormat(lnTotalDolaresAgencia, 12, 2, True) + _
                    oFunI.ImpreFormat("", 12, 2, True) + oFunI.ImpreFormat(lnTotalsolesAgencia, 11, 2, True) + COFF + BOFF, 2
                lnTotalDolaresAgencia = 0
                lnTotalsolesAgencia = 0
                oFunI.Linea lsTexto, BON + lsAgencia + BOFF, 2
                If lnReg > pnLinPage - 2 Then
                    lnReg = 0
                    lstext = lsTexto + Chr$(12)
                    lnNumPage = lnNumPage + 1
                    oFunI.Linea lsTexto, CabeceraRepoCompraVenta(pdDesde, pdHasta, psCabecera, psNomCmact, psNomAge, psNomSeccion, pdFecSis, pnColPage, pnLinPage, lnReg, lnNumPage, lsAgencia)
                End If
            End If
        End If
    Loop
   
    oFunI.Linea lsTexto, BON + String(pnColPage + 5, "=")
    oFunI.Linea lsTexto, BON + oFunI.ImpreFormat("TOTAL GENERAL :", 24) + oFunI.ImpreFormat(lnTotalDolares, 12, 2, True) + _
                oFunI.ImpreFormat(lnTotalsoles, 14, 2, True) + BOFF

End If
Rs.Close
Set Rs = Nothing
GeneraRepoCompraVenta = lsTexto
Set oCajero = Nothing
End Function
Private Function CabeceraRepoCompraVenta(ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                 ByVal psCabecera As String, ByVal psNomCmact As String, _
                                 ByVal psNomAge As String, ByVal psNomSeccion As String, _
                                 ByVal pdFecSis As Date, ByVal pnColPage As Integer, ByVal pnLinPage As Integer, _
                                 ByRef pnReg As Long, ByRef lnNumPage As Integer, ByVal psAgencia As String)
Dim lsTexto As String
Dim oFunI As New COMFunciones.FCOMImpresion
lsTexto = ""
lsTexto = lsTexto + oFunI.CabeRepo(psNomCmact, psNomAge, psNomSeccion, "", Format(pdFecSis, "dd/mm/yyyy"), psCabecera, _
                        " DEL  " & pdDesde & " HASTA " & pdHasta, "", "", lnNumPage, pnColPage, 0, oImp)
pnReg = pnReg + 5
oFunI.Linea lsTexto, ""
oFunI.Linea lsTexto, BON + psAgencia + BOFF
oFunI.Linea lsTexto, BON + String(pnColPage + 5, "=")
oFunI.Linea lsTexto, Con + oFunI.ImpreFormat("NUMERO", 6) + oFunI.ImpreFormat("FECHA", 12) & oFunI.ImpreFormat("C L I E N T E", 35) + oFunI.ImpreFormat("DOLARES", 12) + _
             oFunI.ImpreFormat("CAMBIO", 12) + oFunI.ImpreFormat("SOLES", 8) & oFunI.ImpreFormat("USUARIO", 10) & oFunI.ImpreFormat("HORA", 12) & oFunI.ImpreFormat("AGENCIA", 12) + COFF
oFunI.Linea lsTexto, String(pnColPage + 5, "=") + BOFF
pnReg = pnReg + 5
CabeceraRepoCompraVenta = lsTexto
End Function

Public Function ReporteHabilitacionDevolucion(ByVal sCajero As String, ByVal sAreaCod As String, _
            ByVal sAgecod As String, ByVal dFecha As Date, ByVal sNombreCajero As String, _
            ByVal sNomAge As String) As String

Dim rsHab As New ADODB.Recordset
Dim oRep As NCOMCajero
Dim sCad As String
Dim i As Long
Dim nCarLin As Integer
Dim sTitRp1 As String, sTitRp2 As String
Dim nSaldoHab As Double, nSaldoDev As Double
Dim sItem As String * 4, sNumPag As String * 4
Dim nLinPag As Integer, nCntPag As Integer
Dim sOperacion As String * 35
Dim sHabilitacion As String * 12, sDevolucion As String * 12
Dim sUsuario As String, sHora As String
Dim nItem As Integer
Dim oCap As COMNCaptaGenerales.NCOMCaptaImpresion

sCad = ""
For i = 1 To 2
    Set oRep = New NCOMCajero
    Set rsHab = oRep.GetDetalleHabDevCajero(sAreaCod, sAgecod, dFecha, sCajero, i)
    Set oRep = Nothing
    
    If rsHab.EOF And rsHab.BOF Then
       sCad = sCad & ""
    Else
        
        If i = 1 Then
            sNumPag = ""
            nCarLin = 90
            sTitRp1 = "DETALLE DE HABILITACION / DEVOLUCION"
            sTitRp2 = sCajero & " - " & sNombreCajero
            
            nCntPag = 0
            nCntPag = nCntPag + 1
            RSet sNumPag = Trim(nCntPag)
            Set oCap = New COMNCaptaGenerales.NCOMCaptaImpresion
                sCad = sCad & oCap.CabeRepoCaptac("", "", nCarLin, "OPERACIONES", sTitRp1, sTitRp2, "", sNumPag, Trim(sNomAge), dFecha, Chr$(10))
            Set oCap = Nothing
            sCad = sCad & Chr$(10)
            
            sCad = sCad & String(nCarLin, "-") & Chr$(10)
            sCad = sCad & "ITEM  HORA      OPERACION                            USUA  HABILITACION    DEVOLUCION" & Chr$(10)
            sCad = sCad & String(nCarLin, "-") & Chr$(10)
            sCad = sCad & Chr$(10)
            sCad = sCad & "MONEDA NACIONAL" & Chr$(10)
            nItem = 0
            nLinPag = 7
        Else
            sCad = sCad & Chr$(10)
            sCad = sCad & "MONEDA EXTRANJERA" & Chr$(10)
        End If
        Do While Not rsHab.EOF
            sOperacion = rsHab("cOpeDesc")
            If rsHab("cUsuDest") = sCajero Then
                sUsuario = rsHab("cUsuOrig")
                RSet sHabilitacion = Format$(rsHab("nMovImporte"), "#,##0.00")
                sDevolucion = ""
            Else
                sHabilitacion = ""
                RSet sDevolucion = Format$(rsHab("nMovImporte"), "#,##0.00")
                sUsuario = rsHab("cUsuDest")
            End If
            nItem = nItem + 1
            sHora = Trim(rsHab("Fecha"))
            RSet sItem = Trim(nItem)
            sCad = sCad & sItem & "  " & sHora & "  " & sOperacion & "  " & sUsuario & "  " & sHabilitacion & "  " & sDevolucion & Chr$(10)
            rsHab.MoveNext
        Loop
        
    End If
Next i
Set oRep = Nothing
ReporteHabilitacionDevolucion = sCad
End Function

Public Function ImprimeVaucherCtaPendiente(ByVal psNroVaucher As String, ByVal psPersNom As String, ByVal pnMonto As Double, ByVal pdFechaDoc As Date, _
                                           ByVal psNomCmac As String, ByVal psNomAge As String, ByVal pdFecSis As Date) As String

Dim lsCadImp  As String
Dim oFun As New COMFunciones.FCOMImpresion
lsCadImp = ""
lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
lsCadImp = lsCadImp & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif

lsCadImp = Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)
lsCadImp = lsCadImp & psNomCmac & Space(5) & Format(pdFecSis, "dd/mm/yyyy") & Space(6) & psNomCmac & Space(5) & Format(pdFecSis, "dd/mm/yyyy") & Chr$(10)
lsCadImp = lsCadImp & psNomAge & Space(15) & Format(pdFecSis, "hh:mm:ss") & Space(8) & psNomAge & Space(15) & Format(pdFecSis, "hh:mm:ss") & Chr$(10) & Chr$(10) & Chr$(10)
lsCadImp = lsCadImp & oFun.CentrarCadena("RECIBO VOUCHER CAJA CHICA ", 30) & Space(16) & oFun.CentrarCadena("RECIBO VOUCHER CAJA CHICA ", 30) & Chr$(10)
lsCadImp = lsCadImp & String(38, "-") & Space(4) & String(38, "-") & Chr$(10)
lsCadImp = lsCadImp & "Nro Voucher   :" & psNroVaucher & Space(19) & "Nro Voucher :" & psNroVaucher & Chr$(10) & Chr$(10)
lsCadImp = lsCadImp & "Nombre :" & oFun.Justifica(psPersNom, 30) & Space(4) & "Nombre :" & oFun.Justifica(psPersNom, 30) & Chr$(10)
lsCadImp = lsCadImp & "Monto :" & oFun.ImpreFormat(pnMonto, 11, 2) & Space(21) & "Monto :" & oFun.ImpreFormat(pnMonto, 11, 2) & Chr$(10) & Chr$(10)
lsCadImp = lsCadImp & "Fecha Emisión :" & pdFechaDoc & Space(17) & "Fecha Emisión :" & pdFechaDoc & Chr$(10)
          
ImprimeVaucherCtaPendiente = lsCadImp
Set oFun = Nothing
End Function

'ARCV 18-07-2006

Public Sub IniciaImpresora(Optional ByVal nImpresora As COMFunciones.Impresoras = gEPSON)
    Set oImp = New COMFunciones.FCOMVarImpresion
    oImp.Inicia nImpresora
End Sub

Private Sub Class_Initialize()
IniciaImpresora
End Sub

Private Sub Class_Terminate()
Set oImp = Nothing
End Sub

