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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Dim vsBaseComunes As String
Dim vsBasePesonas As String
'Para activar este evento, use RaiseEvent con la siguiente sintaxis:
'RaiseEvent ShowProgress[(arg1, arg2, ... , argn)]
Public Event ShowProgress()
'Para activar este evento, use RaiseEvent con la siguiente sintaxis:
'RaiseEvent CloseProgress[(arg1, arg2, ... , argn)]
Public Event CloseProgress()
'Para activar este evento, use RaiseEvent con la siguiente sintaxis:
'RaiseEvent Progress[(arg1, arg2, ... , argn)]
Public Event Progress(ByVal pnValor As Long, ByVal pnTotal As Long)

Private Sub Class_Initialize()
Dim oIni As ClasIni
Set oIni = New ClasIni
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
sLpt = "LPT1"
BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")
Set oIni = Nothing
End Sub
Public Function GeneraRepoCompraVenta(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                    ByVal psCodUser As String, ByVal psCodAge As String, _
                                    ByVal psNomCmact As String, ByVal psNomAge As String, ByVal psNomSeccion As String, _
                                    ByVal pdFecSis As Date, ByVal pnColPage As Integer, ByVal pnLinPage As Integer) As String
                              
Dim lsTexto As String
Dim Rs As ADODB.Recordset
Dim oCajero As nCajero
Dim lsCabecera As String
Dim lnNumPage As Integer
Dim lnReg As Long
Dim lnTotal As Long
Dim lnTotalsoles As Currency
Dim lnTotalDolares As Currency
Dim lsAgencia As String
Set oCajero = New nCajero
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
    RaiseEvent ShowProgress
    lsAgencia = Rs!Agencia
    Linea lsTexto, CabeceraRepoCompraVenta(pdDesde, pdHasta, lsCabecera, psNomCmact, psNomAge, psNomSeccion, pdFecSis, pnColPage, pnLinPage, lnReg, lnNumPage, lsAgencia)
    Do While Not Rs.EOF
        lnReg = lnReg + 1
        Linea lsTexto, CON + ImpreFormat(Rs!nMovNro, 6, 0, True) + ImpreFormat(Rs!Fecha, 12) + ImpreFormat(Rs!cpersNombre, 30) + ImpreFormat(Rs!Dolares, 12, 2, True) + _
                ImpreFormat(Rs!Cambio, 10, 2, True) + ImpreFormat(Rs!SOLES, 12, 2, True) & ImpreFormat(Rs!cUser, 8, 5) & ImpreFormat(Format(Rs!Hora, "hh:mm:ss AMPM"), 12) & ImpreFormat(Rs!Agencia, 15) + COFF

        lnTotalsoles = lnTotalsoles + Rs!SOLES
        lnTotalDolares = lnTotalDolares + Rs!Dolares
        lnTotalsolesAgencia = lnTotalsolesAgencia + Rs!SOLES
        lnTotalDolaresAgencia = lnTotalDolaresAgencia + Rs!Dolares
        
        RaiseEvent Progress(lnReg, lnTotal)
        If lnReg > pnLinPage Then
            lnReg = 0
            lstext = lsTexto + Chr$(12)
            lnNumPage = lnNumPage + 1
            Linea lsTexto, CabeceraRepoCompraVenta(pdDesde, pdHasta, psCabecera, psNomCmact, psNomAge, psNomSeccion, pdFecSis, pnColPage, pnLinPage, lnReg, lnNumPage, lsAgencia)
        End If
        Rs.MoveNext
        If Rs.EOF Then
            Linea lsTexto, BON + String(pnColPage + 5, "-")
            Linea lsTexto, CON + BON + ImpreFormat("TOTAL AGENCIA", 20) + ImpreFormat("", 28) + ImpreFormat(lnTotalDolaresAgencia, 12, 2, True) + _
                ImpreFormat("", 12, 2, True) + ImpreFormat(lnTotalsolesAgencia, 11, 2, True) + COFF + BOFF, 2
        Else
            If lsAgencia <> Rs!Agencia Or Rs.EOF Then
                lsAgencia = Rs!Agencia
                Linea lsTexto, BON + String(pnColPage + 5, "-")
                Linea lsTexto, CON + BON + ImpreFormat("TOTAL AGENCIA", 20) + ImpreFormat("", 28) + ImpreFormat(lnTotalDolaresAgencia, 12, 2, True) + _
                    ImpreFormat("", 12, 2, True) + ImpreFormat(lnTotalsolesAgencia, 11, 2, True) + COFF + BOFF, 2
                lnTotalDolaresAgencia = 0
                lnTotalsolesAgencia = 0
                Linea lsTexto, BON + lsAgencia + BOFF, 2
                If lnReg > pnLinPage - 2 Then
                    lnReg = 0
                    lstext = lsTexto + Chr$(12)
                    lnNumPage = lnNumPage + 1
                    Linea lsTexto, CabeceraRepoCompraVenta(pdDesde, pdHasta, psCabecera, psNomCmact, psNomAge, psNomSeccion, pdFecSis, pnColPage, pnLinPage, lnReg, lnNumPage, lsAgencia)
                End If
            End If
        End If
    Loop
    RaiseEvent CloseProgress
    Linea lsTexto, BON + String(pnColPage + 5, "=")
    Linea lsTexto, BON + ImpreFormat("TOTAL GENERAL :", 24) + ImpreFormat(lnTotalDolares, 12, 2, True) + _
                ImpreFormat(lnTotalsoles, 14, 2, True) + BOFF

End If
Rs.Close
Set Rs = Nothing
GeneraRepoCompraVenta = lsTexto
End Function
Private Function CabeceraRepoCompraVenta(ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                 ByVal psCabecera As String, ByVal psNomCmact As String, _
                                 ByVal psNomAge As String, ByVal psNomSeccion As String, _
                                 ByVal pdFecSis As Date, ByVal pnColPage As Integer, ByVal pnLinPage As Integer, _
                                 ByRef pnReg As Long, ByRef lnNumPage As Integer, ByVal psAgencia As String)
Dim lsTexto As String

lsTexto = ""
lsTexto = lsTexto + CabeRepo(psNomCmact, psNomAge, psNomSeccion, "", Format(pdFecSis, "dd/mm/yyyy"), psCabecera, _
                        " DEL  " & pdDesde & " HASTA " & pdHasta, "", "", lnNumPage, pnColPage)
pnReg = pnReg + 5
Linea lsTexto, ""
Linea lsTexto, BON + psAgencia + BOFF
Linea lsTexto, BON + String(pnColPage + 5, "=")
Linea lsTexto, CON + ImpreFormat("NUMERO", 6) + ImpreFormat("FECHA", 12) & ImpreFormat("C L I E N T E", 35) + ImpreFormat("DOLARES", 12) + _
             ImpreFormat("CAMBIO", 12) + ImpreFormat("SOLES", 8) & ImpreFormat("USUARIO", 10) & ImpreFormat("HORA", 12) & ImpreFormat("AGENCIA", 12) + COFF
Linea lsTexto, String(pnColPage + 5, "=") + BOFF
pnReg = pnReg + 5
CabeceraRepoCompraVenta = lsTexto
End Function


Public Function ReporteHabilitacionDevolucion(ByVal sCajero As String, ByVal sAreaCod As String, _
            ByVal sAgeCod As String, ByVal dFecha As Date, ByVal sNombreCajero As String, _
            ByVal sNomAge As String) As String

Dim rsHab As Recordset
Dim oRep As nCajero
Dim sCad As String
Dim I As Long
Dim nCarLin As Integer
Dim sTitRp1 As String, sTitRp2 As String
Dim nSaldoHab As Double, nSaldoDev As Double
Dim sItem As String * 4, sNumPag As String * 4
Dim nLinPag As Integer, nCntPag As Integer
Dim sOperacion As String * 35
Dim sHabilitacion As String * 12, sDevolucion As String * 12
Dim sUsuario As String, sHora As String
Dim nItem As Integer
Dim oImp As clsConstImp

sCad = ""
For I = 1 To 2
    Set oRep = New nCajero
    Set rsHab = oRep.GetDetalleHabDevCajero(sAreaCod, sAgeCod, dFecha, sCajero, I)
    Set oRep = Nothing
    
    If rsHab.EOF And rsHab.BOF Then
       sCad = sCad & ""
    Else
        Set oImp = New clsConstImp
        If I = 1 Then
            sNumPag = ""
            nCarLin = 90
            sTitRp1 = "DETALLE DE HABILITACION / DEVOLUCION"
            sTitRp2 = sCajero & " - " & sNombreCajero
            
            nCntPag = 0
            nCntPag = nCntPag + 1
            RSet sNumPag = Trim(nCntPag)
            
            sCad = sCad & CabeRepoCaptac("", "", nCarLin, "OPERACIONES", sTitRp1, sTitRp2, "", sNumPag, Trim(sNomAge), dFecha) & Chr$(10)
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
Next I
ReporteHabilitacionDevolucion = sCad
End Function
