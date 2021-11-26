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
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
sLPT = "LPT1"
BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")
Set oIni = Nothing
End Sub
Public Function GeneraRepoCompraVenta(ByVal psOpeCod As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                    ByVal psCodUser As String, ByVal psCodAge As String, _
                                    ByVal psNomCmact As String, ByVal psNomAge As String, ByVal psNomSeccion As String, _
                                    ByVal pdFecsis As Date, ByVal pnColPage As Integer, ByVal pnLinPage As Integer) As String
                              
Dim lsTexto As String
Dim rs As ADODB.Recordset
Dim oCajero As nCajero
Dim lsCabecera As String
Dim lnNumPage As Integer
Dim lnReg As Long
Dim lnTotal As Long
Dim lnTotalsoles As Currency
Dim lnTotalDolares As Currency
Dim lsAgencia As String
Set oCajero = New nCajero
Set rs = New ADODB.Recordset
Select Case psOpeCod
    Case gOpeCajeroMECompra
        lsCabecera = "LISTADO DIARIO MONEDA EXTRANJERA - COMPRA"
    Case gOpeCajeroMEVenta
        lsCabecera = "LISTADO DIARIO MONEDA EXTRANJERA - VENTA" '
End Select
'If pdHasta < FechaTopeHistorica Then
'    Set rs = oCajero.GetCompraVentaRepo(psOpeCod, pdDesde, pdHasta, psCodUser, psCodAge, True)
'Else
    Set rs = oCajero.GetCompraVentaRepo(psOpeCod, pdDesde, pdHasta, psCodUser, psCodAge)
'End If
lsAgencia = ""
lnReg = 0: lnTotalsoles = 0: lnTotalDolares = 0: lnNumPage = 0: lsTexto = ""
lnTotalsolesAgencia = 0: lnTotalDolaresAgencia = 0
If Not rs.EOF And Not rs.BOF Then
    lnTotal = rs.RecordCount
    RaiseEvent ShowProgress
    lsAgencia = rs!Agencia
    Linea lsTexto, CabeceraRepoCompraVenta(pdDesde, pdHasta, lsCabecera, psNomCmact, psNomAge, psNomSeccion, pdFecsis, pnColPage, pnLinPage, lnReg, lnNumPage, lsAgencia)
    Do While Not rs.EOF
        lnReg = lnReg + 1
        Linea lsTexto, CON + ImpreFormat(rs!nMovNro, 6, 0, True) + ImpreFormat(rs!Fecha, 12) + ImpreFormat(rs!cPersNombre, 30) + ImpreFormat(rs!Dolares, 12, 2, True) + _
                ImpreFormat(rs!Cambio, 10, 2, True) + ImpreFormat(rs!SOLES, 12, 2, True) & ImpreFormat(rs!cUser, 8, 5) & ImpreFormat(Format(rs!Hora, "hh:mm:ss AMPM"), 12) & ImpreFormat(rs!Agencia, 15) + COFF

        lnTotalsoles = lnTotalsoles + rs!SOLES
        lnTotalDolares = lnTotalDolares + rs!Dolares
        lnTotalsolesAgencia = lnTotalsolesAgencia + rs!SOLES
        lnTotalDolaresAgencia = lnTotalDolaresAgencia + rs!Dolares
        
        RaiseEvent Progress(lnReg, lnTotal)
        If lnReg > pnLinPage Then
            lnReg = 0
            lstext = lsTexto + oImpresora.gPrnSaltoPagina
            lnNumPage = lnNumPage + 1
            Linea lsTexto, CabeceraRepoCompraVenta(pdDesde, pdHasta, psCabecera, psNomCmact, psNomAge, psNomSeccion, pdFecsis, pnColPage, pnLinPage, lnReg, lnNumPage, lsAgencia)
        End If
        rs.MoveNext
        If rs.EOF Then
            Linea lsTexto, BON + String(pnColPage + 5, "-")
            Linea lsTexto, CON + BON + ImpreFormat("TOTAL AGENCIA", 20) + ImpreFormat("", 28) + ImpreFormat(lnTotalDolaresAgencia, 12, 2, True) + _
                ImpreFormat("", 12, 2, True) + ImpreFormat(lnTotalsolesAgencia, 11, 2, True) + COFF + BOFF, 2
        Else
            If lsAgencia <> rs!Agencia Or rs.EOF Then
                lsAgencia = rs!Agencia
                Linea lsTexto, BON + String(pnColPage + 5, "-")
                Linea lsTexto, CON + BON + ImpreFormat("TOTAL AGENCIA", 20) + ImpreFormat("", 28) + ImpreFormat(lnTotalDolaresAgencia, 12, 2, True) + _
                    ImpreFormat("", 12, 2, True) + ImpreFormat(lnTotalsolesAgencia, 11, 2, True) + COFF + BOFF, 2
                lnTotalDolaresAgencia = 0
                lnTotalsolesAgencia = 0
                Linea lsTexto, BON + lsAgencia + BOFF, 2
                If lnReg > pnLinPage - 2 Then
                    lnReg = 0
                    lstext = lsTexto + oImpresora.gPrnSaltoPagina
                    lnNumPage = lnNumPage + 1
                    Linea lsTexto, CabeceraRepoCompraVenta(pdDesde, pdHasta, psCabecera, psNomCmact, psNomAge, psNomSeccion, pdFecsis, pnColPage, pnLinPage, lnReg, lnNumPage, lsAgencia)
                End If
            End If
        End If
    Loop
    RaiseEvent CloseProgress
    Linea lsTexto, BON + String(pnColPage + 5, "=")
    Linea lsTexto, BON + ImpreFormat("TOTAL GENERAL :", 24) + ImpreFormat(lnTotalDolares, 12, 2, True) + _
                ImpreFormat(lnTotalsoles, 14, 2, True) + BOFF

End If
rs.Close
Set rs = Nothing
GeneraRepoCompraVenta = lsTexto
End Function
Private Function CabeceraRepoCompraVenta(ByVal pdDesde As Date, ByVal pdHasta As Date, _
                                 ByVal psCabecera As String, ByVal psNomCmact As String, _
                                 ByVal psNomAge As String, ByVal psNomSeccion As String, _
                                 ByVal pdFecsis As Date, ByVal pnColPage As Integer, ByVal pnLinPage As Integer, _
                                 ByRef pnReg As Long, ByRef lnNumPage As Integer, ByVal psAgencia As String)
Dim lsTexto As String

lsTexto = ""
lsTexto = lsTexto + CabeRepo(psNomCmact, psNomAge, psNomSeccion, "", Format(pdFecsis, gsFormatoFechaView), psCabecera, _
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


