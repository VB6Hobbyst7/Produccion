VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSdoEncaje 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmSdoEncaje"
   ClientHeight    =   1215
   ClientLeft      =   3630
   ClientTop       =   3585
   ClientWidth     =   3540
   Icon            =   "frmSdoEncaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin RichTextLib.RichTextBox RTFEncaje 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   503
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmSdoEncaje.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSdoEncaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim laSdoEnc() As String
Dim lsMoneda As String
Dim lntotal1 As Currency
Dim lntotal2 As Currency
Dim lntotal3 As Currency
Dim lntotal4 As Currency
Dim lntotal5 As Currency
Dim lntotal6 As Currency
Dim lnTotalG As Currency
Dim ldFchIni As Date
Dim lsRTFEnc As String
Dim pdFecIni As Date, pdFecFin As Date
Dim sNomAge  As String
Dim sSql As String
Dim rs As New ADODB.Recordset

Public Sub CalculaSdoEnc(psMoneda As String, psFecIni As String, psFecFin As String)
    Dim lsQryEnc As String, lsMensaj As String
    Dim rsEnc As ADODB.Recordset
    Dim lnDiaMes As Integer
    Dim lnDiaDif As Integer
    Dim I As Integer, j As Integer
    Dim nTotal1 As Currency, nTotal2 As Currency, nTotal3 As Currency
    Dim nTotal4 As Currency, nTotal5 As Currency, nTotal6 As Currency
    Dim nTotal7 As Currency, nTotal8 As Currency, nTotal9 As Currency
    Dim nTotal10 As Currency, nTotal11 As Currency, nTotal12 As Currency, nTotal13 As Currency
    Dim lnNumPag As Integer
    Dim lnLinPag As Integer
    Dim lnCarLin As Integer
    Dim lsTitRp1 As String
    Dim lsTitRp2 As String
    Dim lnCntPag As Integer
    Dim lsCodMon As String
    
    lsRTFEnc = ""
    lsMoneda = psMoneda
    Dim oCon As New DConecta
    Dim oAge As New DActualizaDatosArea
    Dim oDesc As New ClassDescObjeto
    Set rs = oAge.GetAgencias()
    If RSVacio(rs) Then
      Exit Sub
    End If
    oDesc.Show rs, ""
    If Not oDesc.lbOk Then
       Exit Sub
    End If
    frmMdiMain.staMain.Panels(2) = "Conectandose a Agencia " & oDesc.gsSelecDesc & ". Espere Por favor..."
    If Not oCon.AbreConexion Then 'Remota(Right(oDesc.gsSelecCod, 2), , False, "01")
        frmMdiMain.staMain.Panels(2) = ""
        Exit Sub
    End If
    
    frmMdiMain.staMain.Panels(2) = "Generando Información de Agencia " & oDesc.gsSelecDesc & " Espere Por favor..."
    sNomAge = gsNomAge
    gsNomAge = oDesc.gsSelecDesc
    
    nTotal1 = 0: nTotal2 = 0: nTotal3 = 0: nTotal4 = 0
    nTotal5 = 0: nTotal6 = 0: nTotal12 = 0: nTotal13 = 0
    nTotal7 = 0: nTotal8 = 0: nTotal9 = 0
    nTotal10 = 0: nTotal11 = 0
    
    lsMoneda = IIf(psMoneda = "1", "SOLES", "DOLARES")
    
    lsQryEnc = "SELECT A.dEstadAC dEstad, (A.nSaldoAC+A.nChqCMAC+A.nMonChqVal) SaldCntAC, A.nSaldCMAC SaldCMACAC, " _
        & "       (A.nChqCMAC+A.nMonChqVal) ChqValAC, (A.nSaldoAC-A.nSaldCMAC) SaldDispAC, " _
        & "       (PF.nSaldoPF+PF.nMonChqVal) SaldCntPF, PF.nSaldCMAC SaldCMACPF, PF.nMonChqVal ChqValPF, " _
        & "       (CTS.nSaldoCTS+CTS.nMonChqVal) SaldCntCTS, CTS.nMonChqVal ChqValCTS, " _
        & "       (A.nSaldoAC+A.nChqCMAC+A.nMonChqVal+PF.nSaldoPF+PF.nMonChqVal+CTS.nSaldoCTS+CTS.nMonChqVal) Acum, " _
        & "       nSaldCracAC = ISNULL((Select nSaldCRAC From EstadMensAho EA WHERE DateDiff(dd,EA.dEstadMens,A.dEstadAC) = 0 And " _
        & "       EA.cMoneda = A.cMoneda),0), " _
        & "       nSaldCracPF = ISNULL((Select nSaldCRAC From EstadMensPF EP WHERE DateDiff(dd,EP.dEstadMens,A.dEstadAC) = 0 And " _
        & "       EP.cMoneda = A.cMoneda),0) " _
        & "FROM EstadDiaAC A INNER JOIN " _
        & "EstadDiaPF PF INNER JOIN EstadDiaCTS CTS ON Convert(Varchar(10),PF.dEstadPF,103) = " _
        & "Convert(Varchar(10),CTS.dEstadCTS,103) And PF.cMoneda = CTS.cMoneda ON " _
        & "Convert(Varchar(10),A.dEstadAC,103) = Convert(Varchar(10),PF.dEstadPF,103) And " _
        & "A.cMoneda = PF.cMoneda Where A.dEstadAC between '" & Format(CDate(psFecIni), "mm/dd/yyyy") & "' " _
        & "and '" & Format(DateAdd("d", 1, CDate(psFecFin)), "mm/dd/yyyy") & "' " _
        & "And A.cMoneda = '" & psMoneda & "' Order by A.dEstadAC"
    
    Set rsEnc = oCon.CargaRecordSet(lsQryEnc)
    If Not (rsEnc.EOF And rsEnc.BOF) Then
        lnNumPag = 0
        lnLinPag = 65
        lnCarLin = 226
        lsTitRp1 = "A H O R R O   P A R A   E F E C T O S   D E   E N C A J E"
        lnCntPag = 1
        lsTitRp2 = ""

        lsRTFEnc = lsRTFEnc & CabeRepo(gsNomCmac, gsNomAge, "SECCION AHORROS", psMoneda, Format(gdFecSis, gsFormatoFechaView), lsTitRp1, lsTitRp2, "", "", lnNumPag, gnColPage) & oImpresora.gPrnSaltoLinea
        lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        lsRTFEnc = lsRTFEnc & "                                A  H  O  R  R  O  S                                                      P  L  A  Z  O   F  I  J  O                                         C    T    S" & Chr(10)
        lsRTFEnc = lsRTFEnc & "            S  A  L  D  O    S  A  L  D  O    S  A  L  D  O   C H E Q U E S     S A L  D  O    S  A  L  D  O    S  A  L  D  O    S  A  L  D  O    C H E Q U E S    S  A  L  D  O    C H E Q U E S     S A L  D  O      ACUMULADO" & Chr(10)
        lsRTFEnc = lsRTFEnc & "  FECHAS   C O N T A B L E   C  M  A  C  S    OTR.INST.FIN    VALORIZACION      DISPONIBLE    C O N T A B L E   C  M  A  C  S    OTR.INST.FIN     VALORIZACION   C O N T A B L E    VALORIZACION      DISPONIBLE      A  LA FECHA" & Chr(10)
        lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        lnLinPag = 9
        Do While Not rsEnc.EOF
            lsRTFEnc = lsRTFEnc & Format(rsEnc!dEstad, "dd/mm/yyyy") & " "
            lsRTFEnc = lsRTFEnc & PrnVal(rsEnc!SaldCntAC, 15, 2) & " "
            lsRTFEnc = lsRTFEnc & PrnVal(rsEnc!SaldCMACAC, 15, 2) & " "
            lsRTFEnc = lsRTFEnc & PrnVal(rsEnc!nSaldCracAC, 15, 2) & " "
            lsRTFEnc = lsRTFEnc & PrnVal(rsEnc!ChqValAC, 15, 2) & " "
            lsRTFEnc = lsRTFEnc & PrnVal(rsEnc!SaldDispAC - rsEnc!nSaldCracAC, 15, 2) & "   "
            lsRTFEnc = lsRTFEnc & PrnVal(rsEnc!SaldCntPF, 15, 2) & " "
            lsRTFEnc = lsRTFEnc & PrnVal(rsEnc!SaldCMACPF, 15, 2) & " "
            lsRTFEnc = lsRTFEnc & PrnVal(rsEnc!nSaldCracPF, 15, 2) & "   "
            lsRTFEnc = lsRTFEnc & PrnVal(rsEnc!ChqValPF, 15, 2) & " "
            lsRTFEnc = lsRTFEnc & PrnVal(rsEnc!SaldCntCTS, 15, 2) & " "
            lsRTFEnc = lsRTFEnc & PrnVal(rsEnc!ChqValCTS, 15, 2) & "   "
            lsRTFEnc = lsRTFEnc & PrnVal(rsEnc!SaldCntPF + rsEnc!SaldCntCTS - rsEnc!SaldCMACPF - rsEnc!nSaldCracPF - rsEnc!ChqValPF - rsEnc!ChqValCTS, 15, 2) & "   "
            lsRTFEnc = lsRTFEnc & PrnVal(rsEnc!Acum + rsEnc!nSaldCracAC + rsEnc!nSaldCracPF, 15, 2) & oImpresora.gPrnSaltoLinea
            nTotal1 = nTotal1 + rsEnc!SaldCntAC
            nTotal2 = nTotal2 + rsEnc!SaldCMACAC
            nTotal11 = nTotal11 + rsEnc!nSaldCracAC
            nTotal3 = nTotal3 + rsEnc!ChqValAC
            nTotal4 = nTotal4 + (rsEnc!SaldDispAC - rsEnc!nSaldCracAC)
            nTotal5 = nTotal5 + rsEnc!SaldCntPF
            nTotal6 = nTotal6 + rsEnc!SaldCMACPF
            nTotal12 = nTotal12 + rsEnc!nSaldCracPF
            nTotal7 = nTotal7 + rsEnc!ChqValPF
            nTotal8 = nTotal8 + rsEnc!SaldCntCTS
            nTotal9 = nTotal9 + rsEnc!ChqValCTS
            nTotal10 = nTotal10 + rsEnc!Acum + rsEnc!nSaldCracAC + rsEnc!nSaldCracPF
            nTotal13 = nTotal13 + (rsEnc!SaldCntPF + rsEnc!SaldCntCTS - rsEnc!SaldCMACPF - rsEnc!nSaldCracPF - rsEnc!ChqValPF - rsEnc!ChqValCTS)
            rsEnc.MoveNext
        Loop
        lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        lsRTFEnc = lsRTFEnc & String(11, " ")
        lsRTFEnc = lsRTFEnc & PrnVal(nTotal1, 15, 2) & " "
        lsRTFEnc = lsRTFEnc & PrnVal(nTotal2, 15, 2) & " "
        lsRTFEnc = lsRTFEnc & PrnVal(nTotal11, 15, 2) & " "
        lsRTFEnc = lsRTFEnc & PrnVal(nTotal3, 15, 2) & " "
        lsRTFEnc = lsRTFEnc & PrnVal(nTotal4, 15, 2) & "   "
        lsRTFEnc = lsRTFEnc & PrnVal(nTotal5, 15, 2) & " "
        lsRTFEnc = lsRTFEnc & PrnVal(nTotal6, 15, 2) & " "
        lsRTFEnc = lsRTFEnc & PrnVal(nTotal12, 15, 2) & "   "
        lsRTFEnc = lsRTFEnc & PrnVal(nTotal7, 15, 2) & " "
        lsRTFEnc = lsRTFEnc & PrnVal(nTotal8, 15, 2) & " "
        lsRTFEnc = lsRTFEnc & PrnVal(nTotal9, 15, 2) & "   "
        lsRTFEnc = lsRTFEnc & PrnVal(nTotal13, 15, 2) & "   "
        lsRTFEnc = lsRTFEnc & PrnVal(nTotal10, 15, 2) & oImpresora.gPrnSaltoLinea
        lsRTFEnc = lsRTFEnc & String(lnCarLin, "=")
    Else
        lsRTFEnc = ""
    End If
    lsRTFEnc = oImpresora.gPrnCondensadaON & lsRTFEnc & oImpresora.gPrnCondensadaOFF
    EnviaPrevio lsRTFEnc, "AHORROS PARA EFECTO DE ENCAJE", gnLinPage, False
    gsNomAge = sNomAge
    frmMdiMain.staMain.Panels(2) = ""
End Sub

'Public Function CalculaSdoEnc(psMoneda As String, psFecIni As String, psFecFin As String) As RichTextBox
'
'    Dim lsQryEnc As String, lsMensaj As String
'    Dim rsQryEnc As New ADODB.Recordset
'    Dim lnDiaMes As Integer
'    Dim lnDiaDif As Integer
'    Dim I As Integer, J As Integer
'    Dim lntotal1 As Double, lntotal2 As Double, lntotal3 As Double
'    Dim lntotal4 As Double, lntotal5 As Double, lntotal6 As Double
'    Dim lntotal7 As Double, lntotal8 As Double, lntotal9 As Double
'    Dim lntotal10 As Double, lntotal11 As Double
'    lsRTFEnc = ""
'    lsMoneda = psMoneda
'    RTFEncaje.Text = ""
'    Set CalculaSdoEnc = RTFEncaje
'    sSql = gcCentralCom & "spGetTreeObj '112', 3,''"
'    Set rs = dbCmact.Execute(sSql)
'    If RSVacio(rs) Then
'      Exit Function
'    End If
'    frmDescObjeto.Inicio rs, "", 3
'    If Not frmDescObjeto.lOk Then
'      Exit Function
'    End If
'    frmMdiMain.staMain.Panels(2) = "Conectandose a Agencia " & gaObj(0, 0, 0) & " Espere Por favor..."
'    If Not AbreConeccion(Right(gaObj(0, 0, 0), 2)) Then
'        frmMdiMain.staMain.Panels(2) = ""
'        Exit Function
'    End If
'    frmMdiMain.staMain.Panels(2) = "Generando Información de Agencia " & gaObj(0, 0, 0) & " Espere Por favor..."
'    sNomAge = gsNomAge
'    gsNomAge = gaObj(0, 1, 0)
'
''    ldFchIni = CDate("01/" & FillNum(Trim(Str(Month(gdFecSis))), 2, "0") & "/" & Trim(Str(Year(gdFecSis))))
'    ldFchIni = CDate(psFecIni)
'    lnDiaDif = DateDiff("d", ldFchIni, CDate(psFecFin))
'
'    lntotal1 = 0: lntotal2 = 0: lntotal3 = 0: lntotal4 = 0
'    lntotal5 = 0: lntotal6 = 0: lnTotalG = 0
'    lntotal7 = 0: lntotal8 = 0: lntotal9 = 0
'    lntotal10 = 0: lntotal11 = 0
'
'    ReDim laSdoEnc(1 To 12, 1 To 31)
'
'    For I = 1 To UBound(laSdoEnc, 2)
'        For J = 1 To UBound(laSdoEnc, 1)
'            laSdoEnc(J, I) = ""
'        Next J
'    Next I
'
'    lsMoneda = IIf(psMoneda = "1", "SOLES", "DOLARES")
'
'    Do While Not ldFchIni > CDate(psFecFin)    'gdFecSis
'       lnDiaMes = Day(ldFchIni)
'       frmMdiMain.staMain.Panels(2) = "Generando Información de Ag. " & gaObj(0, 0, 0) & " - dia :" & ldFchIni & " Espere..."
'       lsQryEnc = ""
'       lsQryEnc = "SELECT dEstadAC, nSaldoAC, nMonChqAC, nSaldCMAC, nChqCMAC, nMonChqVal  FROM  EstadDiaAC  " _
'                & "WHERE dEstadAC BETWEEN '" & Format(ldFchIni, "mm/dd/yyyy") & "'  " _
'                & "AND '" & Format(ldFchIni + 1, "mm/dd/yyyy") & "' AND cMoneda = '" & psMoneda & "'"
'       If rsQryEnc.State = adStateOpen Then rsQryEnc.Close
'       rsQryEnc.Open lsQryEnc, dbCmactN, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'       If RSVacio(rsQryEnc) Then
'          laSdoEnc(1, lnDiaMes) = ldFchIni
'          laSdoEnc(2, lnDiaMes) = 0
'          laSdoEnc(3, lnDiaMes) = 0
'          laSdoEnc(6, lnDiaMes) = 0
'          laSdoEnc(7, lnDiaMes) = 0
'          laSdoEnc(8, lnDiaMes) = 0
'          laSdoEnc(9, lnDiaMes) = 0
'       Else
'          laSdoEnc(1, lnDiaMes) = ldFchIni
'          laSdoEnc(2, lnDiaMes) = CCur(rsQryEnc!nSaldoAC) + CCur(rsQryEnc!nMonChqVal) + CCur(rsQryEnc!nChqCMAC)
'          laSdoEnc(3, lnDiaMes) = CCur(rsQryEnc!nSaldCMAC) + CCur(rsQryEnc!nChqCMAC)      'LeeChequesValor(gsCodProAC, psMoneda)
'          laSdoEnc(4, lnDiaMes) = CCur(rsQryEnc!nMonChqVal) + CCur(rsQryEnc!nChqCMAC)
'          laSdoEnc(5, lnDiaMes) = CCur(rsQryEnc!nSaldoAC) - CCur(rsQryEnc!nSaldCMAC)
'       End If
'
'       lsQryEnc = ""
'       lsQryEnc = "SELECT dEstadPF, nSaldoPF, nMonChqVal, nNumCMAC, ISNULL(nSaldCMAC,0) AS nSaldCMAC  FROM  EstadDiaPF  " _
'                & "WHERE dEstadPF BETWEEN '" & Format(ldFchIni, "mm/dd/yyyy") & "'  " _
'                & "AND '" & Format(ldFchIni + 1, "mm/dd/yyyy") & "' AND cMoneda = '" & psMoneda & "'"
'       If rsQryEnc.State = adStateOpen Then rsQryEnc.Close
'       rsQryEnc.Open lsQryEnc, dbCmactN, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'       If RSVacio(rsQryEnc) Then
'          laSdoEnc(6, lnDiaMes) = 0
'          laSdoEnc(7, lnDiaMes) = 0
'          laSdoEnc(8, lnDiaMes) = 0
'          laSdoEnc(9, lnDiaMes) = 0
'       Else
'          laSdoEnc(6, lnDiaMes) = CCur(rsQryEnc!nSaldoPF) + CCur(rsQryEnc!nMonChqVal)   'SALDO CONTABLE
'          laSdoEnc(7, lnDiaMes) = CCur(rsQryEnc!nSaldCMAC)     'SALDO CMACT
'          laSdoEnc(8, lnDiaMes) = CCur(rsQryEnc!nMonChqVal)    'SALDO CHEQUES 'LeeChequesValor(gsCodProPF, psMoneda)
'          laSdoEnc(9, lnDiaMes) = CCur(rsQryEnc!nSaldoPF) - CCur(rsQryEnc!nSaldCMAC)   'SALDO DISPONIBLE 'LeeChequesValor(gsCodProPF, psMoneda)
'       End If
'
'       lsQryEnc = ""
'       lsQryEnc = "SELECT dEstadCTS, (nSaldoCTS) AS nSaldo, nMonChqVal FROM  EstadDiaCTS  " _
'                & "WHERE dEstadCTS BETWEEN '" & Format(ldFchIni, "mm/dd/yyyy") & "'  " _
'                & "AND '" & Format(ldFchIni + 1, "mm/dd/yyyy") & "' AND cMoneda = '" & psMoneda & "'"
'       If rsQryEnc.State = adStateOpen Then rsQryEnc.Close
'       rsQryEnc.Open lsQryEnc, dbCmactN, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'       If RSVacio(rsQryEnc) Then
'          laSdoEnc(10, lnDiaMes) = 0
'          laSdoEnc(11, lnDiaMes) = 0
'       Else
'          laSdoEnc(10, lnDiaMes) = CCur(rsQryEnc!nSaldo) + CCur(rsQryEnc!nMonChqVal)
'          laSdoEnc(11, lnDiaMes) = CCur(rsQryEnc!nMonChqVal)  'LeeChequesValor(gsCodProPF, psMoneda)
'          laSdoEnc(12, lnDiaMes) = CCur(laSdoEnc(2, lnDiaMes)) + CCur(laSdoEnc(6, lnDiaMes)) + CCur(laSdoEnc(10, lnDiaMes))
'       End If
'
'       lntotal1 = lntotal1 + CCur(IIf(laSdoEnc(2, lnDiaMes) = "", 0, laSdoEnc(2, lnDiaMes)))
'       lntotal2 = lntotal2 + CCur(IIf(laSdoEnc(3, lnDiaMes) = "", 0, laSdoEnc(3, lnDiaMes)))
'       lntotal3 = lntotal3 + CCur(IIf(laSdoEnc(4, lnDiaMes) = "", 0, laSdoEnc(4, lnDiaMes)))
'       lntotal4 = lntotal4 + CCur(IIf(laSdoEnc(5, lnDiaMes) = "", 0, laSdoEnc(5, lnDiaMes)))
'       lntotal5 = lntotal5 + CCur(IIf(laSdoEnc(6, lnDiaMes) = "", 0, laSdoEnc(6, lnDiaMes)))
'       lntotal6 = lntotal6 + CCur(IIf(laSdoEnc(7, lnDiaMes) = "", 0, laSdoEnc(7, lnDiaMes)))
'       lntotal7 = lntotal7 + CCur(IIf(laSdoEnc(8, lnDiaMes) = "", 0, laSdoEnc(8, lnDiaMes)))
'       lntotal8 = lntotal8 + CCur(IIf(laSdoEnc(9, lnDiaMes) = "", 0, laSdoEnc(9, lnDiaMes)))
'       lntotal10 = lntotal10 + CCur(IIf(laSdoEnc(10, lnDiaMes) = "", 0, laSdoEnc(10, lnDiaMes)))
'       lntotal11 = lntotal11 + CCur(IIf(laSdoEnc(11, lnDiaMes) = "", 0, laSdoEnc(11, lnDiaMes)))
'
'       ldFchIni = ldFchIni + 1
'    Loop
'
'    CierraConeccion
'
'    lnTotalG = lntotal1 + lntotal5 + lntotal10
'
'    If lnTotalG > 0 Then
'       PrtSdoEnc
'    Else
'       lsRTFEnc = ""
'    End If
'    RTFEncaje.Text = CON & lsRTFEnc & COFF
'    Set CalculaSdoEnc = RTFEncaje
'    gsNomAge = sNomAge
'    frmMdiMain.staMain.Panels(2) = ""
'End Function

'Private Function PrtSdoEnc()
'    Dim lnCarLin As Integer, I As Integer
'    Dim lsTitRp1 As String
'    Dim lsTitRp2 As String
'    Dim lsNumPag As String
'    Dim lnLinPag As Integer
'    Dim lnCntPag As Integer
'    Dim lsCodMon As String
'    Dim Total1 As Currency, Total2 As Currency, total3 As Currency
'    Dim total4 As Currency, total5 As Currency, total6 As Currency
'    Dim total7 As Currency, total8 As Currency, total9 As Currency
'    Dim total10 As Currency, total11 As Currency
'
'    lsNumPag = ""
'    lnLinPag = 65
'    lnCarLin = 200
'    lsTitRp1 = "A H O R R O   P A R A   E F E C T O S   D E   E N C A J E"
'    lnCntPag = 1
'    lsTitRp2 = gsNomAge
'
'    lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
'
'    lsRTFEnc = lsRTFEnc & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag) & oImpresora.gPrnSaltoLinea
'
'    lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
'    lsRTFEnc = lsRTFEnc & "                                A  H  O  R  R  O  S                                           P L A Z O   F I J O                                                C    T    S                             " & Chr(10)
'    lsRTFEnc = lsRTFEnc & "            S  A  L  D  O    S  A  L  D  O   C H E Q U E S   S  A  L  D  O    S  A  L  D  O    S  A  L  D  O   C H E Q U E S   S  A  L  D  O    S  A  L  D  O    C H E Q U E S      ACUMULADO   " & Chr(10)
'    lsRTFEnc = lsRTFEnc & "  FECHAS   C O N T A B L E   C  M  A  C  S    VALORIZACION    DISPONIBLE     C O N T A B L E   C  M  A  C  S    VALORIZACION    DISPONIBLE     C O N T A B L E    VALORIZACION     A  LA FECHA  " & Chr(10)
'    lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
'    lnLinPag = 9
'
'    For I = 1 To 31
'        If laSdoEnc(1, I) <> "" Then
'           lsRTFEnc = lsRTFEnc & Format(laSdoEnc(1, I), "dd/mm/yyyy") & " "
'           lsRTFEnc = lsRTFEnc & JDNum(laSdoEnc(2, I), 15, True, 12, 2) & " "
'           lsRTFEnc = lsRTFEnc & JDNum(laSdoEnc(3, I), 15, True, 12, 2) & " "
'           lsRTFEnc = lsRTFEnc & JDNum(laSdoEnc(4, I), 15, True, 12, 2) & " "
'           lsRTFEnc = lsRTFEnc & JDNum(laSdoEnc(5, I), 15, True, 12, 2) & " "
'           lsRTFEnc = lsRTFEnc & JDNum(laSdoEnc(6, I), 15, True, 12, 2) & " "
'           lsRTFEnc = lsRTFEnc & JDNum(laSdoEnc(7, I), 15, True, 12, 2) & " "
'           lsRTFEnc = lsRTFEnc & JDNum(laSdoEnc(8, I), 15, True, 12, 2) & " "
'           lsRTFEnc = lsRTFEnc & JDNum(laSdoEnc(9, I), 15, True, 12, 2) & "   "
'           lsRTFEnc = lsRTFEnc & JDNum(laSdoEnc(10, I), 15, True, 12, 2) & "   "
'           lsRTFEnc = lsRTFEnc & JDNum(laSdoEnc(11, I), 15, True, 12, 2) & "   "
'           lsRTFEnc = lsRTFEnc & JDNum(laSdoEnc(12, I), 15, True, 12, 2) & oImpresora.gPrnSaltoLinea
'
'           Total1 = Total1 + CCur(IIf(laSdoEnc(2, I) = "", 0, laSdoEnc(2, I)))
'           Total2 = Total2 + CCur(IIf(laSdoEnc(3, I) = "", 0, laSdoEnc(3, I)))
'           total3 = total3 + CCur(IIf(laSdoEnc(4, I) = "", 0, laSdoEnc(4, I)))
'           total4 = total4 + CCur(IIf(laSdoEnc(5, I) = "", 0, laSdoEnc(5, I)))
'           total5 = total5 + CCur(IIf(laSdoEnc(6, I) = "", 0, laSdoEnc(6, I)))
'           total6 = total6 + CCur(IIf(laSdoEnc(7, I) = "", 0, laSdoEnc(7, I)))
'           total7 = total7 + CCur(IIf(laSdoEnc(8, I) = "", 0, laSdoEnc(8, I)))
'           total8 = total8 + CCur(IIf(laSdoEnc(9, I) = "", 0, laSdoEnc(9, I)))
'           total9 = total9 + CCur(IIf(laSdoEnc(10, I) = "", 0, laSdoEnc(10, I)))
'           total10 = total10 + CCur(IIf(laSdoEnc(11, I) = "", 0, laSdoEnc(11, I)))
'           total11 = total11 + CCur(IIf(laSdoEnc(12, I) = "", 0, laSdoEnc(12, I)))
'        End If
'    Next I
'    lsRTFEnc = lsRTFEnc & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
'    lsRTFEnc = lsRTFEnc & String(11, " ")
'    lsRTFEnc = lsRTFEnc & PrnVal(Total1, 15, 2) & " "
'    lsRTFEnc = lsRTFEnc & PrnVal(Total2, 15, 2) & " "
'    lsRTFEnc = lsRTFEnc & PrnVal(total3, 15, 2) & " "
'    lsRTFEnc = lsRTFEnc & PrnVal(total4, 15, 2) & " "
'    lsRTFEnc = lsRTFEnc & PrnVal(total5, 15, 2) & " "
'    lsRTFEnc = lsRTFEnc & PrnVal(total6, 15, 2) & " "
'    lsRTFEnc = lsRTFEnc & PrnVal(total7, 15, 2) & " "
'    lsRTFEnc = lsRTFEnc & PrnVal(total8, 15, 2) & "   "
'    lsRTFEnc = lsRTFEnc & PrnVal(total9, 15, 2) & "   "
'    lsRTFEnc = lsRTFEnc & PrnVal(total10, 15, 2) & "   "
'    lsRTFEnc = lsRTFEnc & PrnVal(total11, 15, 2) & oImpresora.gPrnSaltoLinea
'    lsRTFEnc = lsRTFEnc & String(lnCarLin, "=")
'End Function

Private Sub Form_Load()
AbreConexion
CentraForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub
