VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCaptaImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public oImp As COMFunciones.FCOMVarImpresion
Public gcPEN As New COMDConstantes.DCOMConstantes
Public Sub IniciaImpresora(Optional ByVal nImpresora As COMFunciones.Impresoras = gEPSON)
Set oImp = New COMFunciones.FCOMVarImpresion
oImp.Inicia nImpresora
End Sub

Public Function EmiteBoleta(ByVal sMsgProd As String, ByVal sMsgOpe As String, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double, _
            ByVal nIntMes As Double, ByVal nExtracto As Long, Optional bDocumento As Boolean = False, Optional nDocumento As TpoDoc = TpoDocCheque, _
            Optional sNroDoc As String = "", Optional dFechaValor As Date, Optional bImpSaldos As Boolean = True, _
            Optional pdFecSis As Date, Optional psNomAge As String = "", Optional psCodUser As String = "") As String

'Dim bReImp As Boolean
'Dim sTipDep As String, sCodOpe As String
'Dim sModDep As String, sTipApe As String
'Dim sNomTit As String
'Dim clsMant As NCOMCaptaGenerales
'Set clsMant = New NCOMCaptaGenerales
'Dim oImpre As New COMFunciones.FCOMImpresion
'
'sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
'sCodOpe = Trim(nOperacion)
'sModDep = sMsgOpe
'sTipApe = sMsgProd
'sNomTit = oImpre.ImpreCarEsp(clsMant.GetNombreTitulares(sCuenta))
'Set clsMant = Nothing
'bReImp = False
'If bDocumento Then
'    Select Case nDocumento
'        Case TpoDocCheque
'            EmiteBoleta = oImp.ImprimeBoleta sTipApe, oimpre.ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, oFunF.gsFormatoFechaView), nSaldoDisp, nIntMes, "Fecha Valor", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecsis, psNomAge, psCodUser
'        Case TpoDocNotaAbono, TpoDocNotaCargo
'            EmiteBoleta = oImp.ImprimeBoleta sTipApe, oimpre.ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecsis, psNomAge, psCodUser
'        Case TpoDocOrdenPago
'            EmiteBoleta = oImp.ImprimeBoleta sTipApe, oimpre.ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecsis, psNomAge, psCodUser
'    End Select
'End If
'    If MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
'        bReImp = True
'    Else
'        bReImp = False
'   End If
'Loop Until Not bReImp
End Function

Public Function CabeRepoCaptac(ByVal sCabe01 As String, ByVal sCabe02 As String, _
        ByVal nCarLin As Long, ByVal sSeccion As String, ByVal sTitRp1 As String, _
        ByVal sTitRp2 As String, ByVal sMoneda As String, ByVal sNumPag As String, _
        ByVal sNomAge As String, ByVal dFecSis As Date, ByVal sSaltoLinea As String) As String

Dim sTit1 As String, sTit2 As String
Dim sMon As String
Dim sCad As String
sTit1 = "": sTit2 = ""

CabeRepoCaptac = ""


' Definición de Cabecera 1
sMoneda = IIf(sMoneda = "", String(10, " "), " - " & sMoneda)
sCad = UCase(Trim(sNomAge)) & sMoneda
sCabe01 = sCad & String(50 - Len(sCad), " ")
sCabe01 = sCabe01 & Space((IIf(nCarLin <= 36, 80, nCarLin) - 36) - (Len(Mid(Trim(sCabe01), 1, 45)) - 2))
sCabe01 = sCabe01 & "PAGINA: " & sNumPag
sCabe01 = sCabe01 & Space(5) & "FECHA: " & Format$(dFecSis, "dd/mm/yyyy")

' Definición de Cabecera 2
sCabe02 = sSeccion & String(19 - Len(sSeccion), " ")
sCabe02 = sCabe02 & Space((IIf(nCarLin <= 19, 100, nCarLin) - 19) - (Len(sCabe02) - 2))
sCabe02 = sCabe02 & "HORA :   " & Format$(Now(), "hh:mm:ss")

' Definición del Titulo del Reporte
sTit1 = String(Int((IIf(nCarLin <= Len(sTitRp1), Len(sTitRp1) + 1, nCarLin) - Len(sTitRp1)) / 2), " ") & sTitRp1
sTit2 = String(Int((IIf(nCarLin <= Len(sTitRp2), Len(sTitRp2) + 1, nCarLin) - Len(sTitRp2)) / 2), " ") & sTitRp2
    
CabeRepoCaptac = CabeRepoCaptac & sCabe01 & sSaltoLinea
CabeRepoCaptac = CabeRepoCaptac & sCabe02 & sSaltoLinea
CabeRepoCaptac = CabeRepoCaptac & sTit1 & sSaltoLinea
CabeRepoCaptac = CabeRepoCaptac & sTit2

End Function
Public Function CabeRepoCaptacSM(ByVal sCabe01 As String, ByVal sCabe02 As String, _
        ByVal nCarLin As Integer, ByVal sSeccion As String, ByVal sTitRp1 As String, _
        ByVal sTitRp2 As String, ByVal sMoneda As String, ByVal sNumPag As String, _
        ByVal sNomAge As String, ByVal dFecSis As Date) As String

Dim sTit1 As String, sTit2 As String
Dim sMon As String
Dim sCad As String
sTit1 = "": sTit2 = ""

' Definición de Cabecera 1
sMoneda = IIf(sMoneda = "", String(10, " "), " - " & sMoneda)
sCad = UCase(Trim(sNomAge)) & sMoneda
sCabe01 = sCad & String(50 - Len(sCad), " ")
sCabe01 = sCabe01 & Space((nCarLin - 36) - (Len(Mid(Trim(sCabe01), 1, 45)) - 2))
sCabe01 = sCabe01 & "PAGINA: " & sNumPag
sCabe01 = sCabe01 & Space(5) & "FECHA: " & Format$(dFecSis, "dd/mm/yyyy")

' Definición de Cabecera 2
sCabe02 = sSeccion & String(19 - Len(sSeccion), " ")
sCabe02 = sCabe02 & Space((nCarLin - 19) - (Len(sCabe02) - 2))
sCabe02 = sCabe02 & "HORA :   " & Format$(Now(), "hh:mm:ss")

' Definición del Titulo del Reporte
sTit1 = String(Int((nCarLin - Len(sTitRp1)) / 2), " ") & sTitRp1
sTit2 = String(Int((nCarLin - Len(sTitRp2)) / 2), " ") & sTitRp2
    
CabeRepoCaptacSM = CabeRepoCaptacSM & Space(30) & sCabe01 & oImp.gPrnSaltoLinea
CabeRepoCaptacSM = CabeRepoCaptacSM & Space(30) & sCabe02 & oImp.gPrnSaltoLinea
CabeRepoCaptacSM = CabeRepoCaptacSM & sTit1 & oImp.gPrnSaltoLinea
CabeRepoCaptacSM = CabeRepoCaptacSM & sTit2
End Function

Public Function ConvierteTNAaTEA(ByVal nTasa As Double) As Double
ConvierteTNAaTEA = ((1 + nTasa / 36000) ^ 360 - 1) * 100
End Function

Public Function ConvierteTEAaTNA(ByVal nTasa As Double) As Double
ConvierteTEAaTNA = ((1 + nTasa / 100) ^ (1 / 360) - 1) * 36000
End Function

Public Function CargaDataReport(ByVal drObjeto As Object, ByVal pRs As ADODB.Recordset)

With drObjeto
   Set .DataSource = pRs
      .DataMember = ""
      .WindowState = vbMaximized
      .Title = "CONVENIO ORDEN DE PAGO"
      .inicio
      .Refresh
      .Show 1
 End With

End Function

'Public Function ImprimeBoletaLavadoDinero(ByVal sNomCmac As String, ByVal sNomAge As String, _
'        ByVal dFecha As Date, ByVal sCuenta As String, ByVal sTitNombre As String, _
'        ByVal sTitDocId As String, ByVal sTitDireccion As String, ByVal sReaNombre As String, _
'        ByVal sReaDocId As String, ByVal sReaDireccion As String, ByVal sBenNombre As String, _
'        ByVal sBenDocId As String, ByVal sBenDireccion As String, ByVal sOperacion As String, _
'        ByVal nImporte As Double, ByVal sLpt As String, _
'        ByVal sOrdNombre As String, ByVal sOrdDocId As String, ByVal sOrdDireccion As String, _
'        Optional sAgeRem As String = "", _
'        Optional bCadena As Boolean = False, Optional sTipo As String = "", Optional nMoneda As Integer = 1, Optional pbImpTMU As Boolean = False, Optional gsCodAge As String, Optional psOrigenEfectivo As String = "") As String
        'By Capi 20012008 se inserto parametros de persona que ordena transaccion

'JACA 20110224 SE HA INCREMENTADO A 4 BENEFICIARIOS
Public Function ImprimeBoletaLavadoDinero(ByVal sNomCmac As String, ByVal sNomAge As String, _
        ByVal dFecha As Date, ByVal sCuenta As String, _
        ByVal sTitNombre As String, ByVal sTitDocId As String, ByVal sTitDireccion As String, ByVal sTitOcupacion As String, _
        ByVal sReaNombre As String, ByVal sReaDocId As String, ByVal sReaDireccion As String, ByVal sReaOcupacion As String, _
        ByVal sBenNombre As String, ByVal sBenDocId As String, ByVal sBenDireccion As String, ByVal sBenOcupacion As String, _
        ByVal sOperacion As String, ByVal nImporte As Double, ByVal sLpt As String, _
        ByVal sOrdNombre As String, ByVal sOrdDocId As String, ByVal sOrdDireccion As String, ByVal sOrdOcupacion As String, _
        Optional sAgeRem As String = "", _
        Optional bCadena As Boolean = False, Optional sTipo As String = "", Optional nMoneda As Integer = 1, Optional pbImpTMU As Boolean = False, Optional gsCodAge As String, Optional psOrigenEfectivo As String = "", _
        Optional sBenNombre2 As String = "", Optional sBenDocId2 As String = "", Optional sBenDireccion2 As String = "", Optional sBenOcupacion2 As String = "", _
        Optional sBenNombre3 As String = "", Optional sBenDocId3 As String = "", Optional sBenDireccion3 As String = "", Optional sBenOcupacion3 As String = "", _
        Optional sBenNombre4 As String = "", Optional sBenDocId4 As String = "", Optional sBenDireccion4 As String = "", Optional sBenOcupacion4 As String = "", _
        Optional nTipoREU As Integer = 1, Optional pnNroREU As String = "0", Optional poListaRealizan As Variant = Nothing, Optional poListaOrdenan As Variant = Nothing, _
        Optional poListaBeneficia As Variant = Nothing) As String 'EJVG20120329 Se agrego los que realizan y ordenan de posicion 2 al 4
        'Optional sBenNombre4 As String = "", Optional sBenDocId4 As String = "", Optional sBenDireccion4 As String = "", Optional sBenOcupacion4 As String = "", Optional nTipoREU As Integer = 1, Optional pnNroREU As String = "0") As String
        'By Capi 20012008 se inserto parametros de persona que ordena transaccion
        'WIOR 20130301 AGREGO poListaBeneficia
        
    Dim sImpre As String, sCad As String
    Dim sOpe As String
    Dim nFila As Integer, nSep As Integer, nMax As Integer
    Dim sTexto As String
    'By Capi 28022008
    Dim lrs As ADODB.Recordset
    Dim lsCorrLavDinero As String
    Dim loLavDin As COMDPersona.DCOMPersonas
    Set loLavDin = New COMDPersona.DCOMPersonas
    Dim iRea As Integer, iOrd As Integer 'EJVG20120329
    Dim iBen As Integer 'WIOR 20130301
    Dim oImpre As New COMFunciones.FCOMImpresion
    Dim objPersona As PersonaLavado
    sCad = sCad & oImp.gPrnInicializa
    sCad = sCad & oImp.gPrnEspaLineaN   'espaciamiento lineas 1/6 pulg.
    sCad = sCad & oImp.gPrnTamPagina22  'Longitud de página a 22 líneas'
    sCad = sCad & oImp.gPrnTamLetra10CPI   'Tamaño 10 cpi
    sCad = sCad & oImp.gPrnTpoLetraRoman     'Tipo de Letra Sans Serif
    sCad = sCad & oImp.gPrnCondensadaOFF ' cancela condensada
    sCad = sCad & oImp.gPrnNegritaOFF ' desactiva negrita
    
    nSep = 14
    
    'nMax = 39
    'By Capi 14022008
    nMax = 41
    
    'JACA 20110304
        'By Capi 28022008
    'Comentado By JACA 20110930****************************************************
'        If pnNroREU = "0" Then
'            Set lrs = loLavDin.ObtenerAgenCorrelativoLavDinero(gsCodAge)
'            Call loLavDin.ActualizarAgenCorrelativoLavDinero(gsCodAge)
'            Set loLavDin = Nothing
'            'lsCorrLavDinero = gsCodAge & "-" & (Format(lrs!correlativo, "000000"))
            'lsCorrLavDinero = gsCodAge & "-" & (Format(lrs!nNroREU, "000000"))
'        Else
            lsCorrLavDinero = gsCodAge & "-" & (Format(pnNroREU, "000000"))
'        End If
    'END JACA**********************************************************************
    
    'WIOR 20150708 ***
    Dim oCtaTit As ADODB.Recordset
    Dim oCtaTpo As ADODB.Recordset
    Dim cTpoPers As String
    Dim indx As Integer
    Dim nCtaTpo As Integer
    Dim oCap As New COMDCaptaGenerales.DCOMCaptaGenerales
    
    Set oCtaTit = New ADODB.Recordset
    Set oCtaTpo = New ADODB.Recordset
    If Mid(sCuenta, 6, 3) = "232" Or Mid(sCuenta, 6, 3) = "233" Or Mid(sCuenta, 6, 3) = "234" Then
        
        Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
        
        Set oCtaTit = oCap.GetProductoPersona(sCuenta)
        
        If Mid(sCuenta, 6, 3) = "232" Then
            Set oCtaTpo = oCap.GetDatosCuentaAho(sCuenta)
        ElseIf Mid(sCuenta, 6, 3) = "233" Then
            Set oCtaTpo = oCap.GetDatosCuentaPF(sCuenta)
        ElseIf Mid(sCuenta, 6, 3) = "234" Then
            Set oCtaTpo = oCap.GetDatosCuentaCTS(sCuenta)
        End If
        
        nCtaTpo = 1
    ElseIf Mid(sCuenta, 6, 3) = "239" Then
       Dim oCapSer As New COMDCaptaServicios.DCOMCaptaServicios
       Set oCapSer = New COMDCaptaServicios.DCOMCaptaServicios
       Set oCtaTit = oCapSer.RecuperaDatosRelacGiro(sCuenta)
       nCtaTpo = 2
    ElseIf sCuenta = "" Then
       nCtaTpo = 2
    Else
        Set oCtaTit = RecuperaDatosRelacCred(sCuenta)
        Set oCtaTpo = RecuperaDatosCreditoVigente(sCuenta)
        nCtaTpo = 3
    End If
    'WIOR FIN ********
    
    If pbImpTMU = False Then

        'By capi 28122008
        'sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica(sNomCmac, nMax - 14) & " " & "Nro." & Left(lsCorrLavDinero, 12) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica(sNomCmac, nMax - 15) & " " & "Nro." & Left(lsCorrLavDinero, 12) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        '''''sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica(sNomCmac, nMax - 14) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica(sNomCmac, nMax - 15) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea ' Comentado by JACA 20110930
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica(sNomCmac, nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica(sNomCmac, nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea 'Modificado by JACA 20110930
        'sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica(sNomCmac, nMax - 11) & " " & Left(dFecha, 10) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica(sNomCmac, nMax - 11) & " " & Left(dFecha, 10) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        'sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("PROC.:" & Trim(sNomAge) & "/" & sAgeRem, nMax - 12) & " " & Format$(Time, "hh:mm:ss AMPM") & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("PROC.:" & Trim(sNomAge) & "/" & sAgeRem, nMax - 12) & " " & Format$(Time, "hh:mm:ss AMPM") & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        'By capi 28122008
        'sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("PROC.:" & Trim(sNomAge) & "/" & sAgeRem, nMax - 1) & " " & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("PROC.:" & Trim(sNomAge) & "/" & sAgeRem, nMax - 12) & " " & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        '''''sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("PROC.:" & Trim(sNomAge) & "/" & sAgeRem, nMax - 1) & " " & " " & "Nro." & Left(lsCorrLavDinero, 12) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("PROC.:" & Trim(sNomAge) & "/" & sAgeRem, nMax - 12) & " " & "Nro." & Left(lsCorrLavDinero, 12) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea'Comentado by JACA 20110930
        '''''sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("PROC.:" & Trim(sNomAge) & "/" & sAgeRem & " Nro." & Left(lsCorrLavDinero, 12), nMax - 1) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("PROC.:" & Trim(sNomAge) & "/" & sAgeRem & " Nro." & Left(lsCorrLavDinero, 12), nMax - 1) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea 'Modificado by JACA 20110930
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("PROC.:" & Trim(sNomAge) & "/" & sAgeRem, nMax - 1) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica(" PROC.:" & Trim(sNomAge) & "/" & sAgeRem, nMax - 1) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea   'Modificado by JACA 20110930
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("Nro.RO: " & lsCorrLavDinero, nMax - 1) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica(" Nro.RO: " & lsCorrLavDinero, nMax - 1) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea 'Modificado by JACA 20110930
        
        'sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("REU: " & lsCorrLavDinero & oImp.gPrnNegritaON & "(Cta: " & sCuenta & ") " & oImp.gPrnNegritaOFF, nMax + 2) & oImp.gPrnNegritaOFF & Space(nSep + 2) & oImp.gPrnNegritaON & oImpre.Justifica("REU: " & lsCorrLavDinero & oImp.gPrnNegritaON & "(Cta: " & sCuenta & ") " & oImp.gPrnNegritaOFF, nMax + 2) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("==========REGISTRO DE OPERACIONES========== ", nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("==========REGISTRO DE OPERACIONES========== ", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        '**************************
        '**************************
        '**************************
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA REALIZA TRANSACCIÓN  ----", nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA REALIZA TRANSACCIÓN  ----", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        
        'sCad = sCad & oImpre.Justifica("NOMBRE : " & sReaNombre, nMax) & Space(nSep) & oImpre.Justifica("NOMBRE : " & sReaNombre, nMax) & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("DOC ID : " & sReaDocId, nMax - 9) & "---------" & Space(nSep) & oImpre.Justifica("DOC ID : " & sReaDocId, nMax - 9) & "---------" & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("DIREC. : " & sReaDireccion, nMax - 9) & "         " & Space(nSep) & oImpre.Justifica("DIREC. : " & sReaDireccion, nMax - 9) & "        " & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("OCUPA. : " & sReaOcupacion, nMax - 9) & "         " & Space(nSep) & oImpre.Justifica("OCUPA. : " & sReaOcupacion, nMax - 9) & "        " & oImp.gPrnSaltoLinea 'JACA 20110225
        'sCad = sCad & oImp.gPrnSaltoLinea 'JACA 20110225
        
        
        'EJVG20120329
        For iRea = 0 To UBound(poListaRealizan)
            objPersona = poListaRealizan(iRea)
            sCad = sCad & oImpre.Justifica("NOMBRE : " & objPersona.Nombre, nMax) & Space(nSep) & oImpre.Justifica("NOMBRE : " & objPersona.Nombre, nMax) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("DOC ID : " & objPersona.DocumentoId, nMax - 9) & "---------" & Space(nSep) & oImpre.Justifica("DOC ID : " & objPersona.DocumentoId, nMax - 9) & "---------" & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("DIREC. : " & objPersona.Direccion, nMax - 9) & "         " & Space(nSep) & oImpre.Justifica("DIREC. : " & objPersona.Direccion, nMax - 9) & "        " & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("OCUPA. : " & objPersona.Ocupacion, nMax - 9) & "         " & Space(nSep) & oImpre.Justifica("OCUPA. : " & objPersona.Ocupacion, nMax - 9) & "        " & oImp.gPrnSaltoLinea
            'WIOR 20130301 ****************************************************
            sCad = sCad & oImpre.Justifica("NAC.   : " & objPersona.Nacionalidad, nMax - 9) & "         " & Space(nSep) & oImpre.Justifica("NAC.   : " & objPersona.Nacionalidad, nMax - 9) & "        " & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("RESIDE : " & objPersona.Residente, nMax - 9) & "         " & Space(nSep) & oImpre.Justifica("RESIDE : " & objPersona.Residente, nMax - 9) & "        " & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("PEPS   : " & objPersona.Peps, nMax - 9) & "         " & Space(nSep) & oImpre.Justifica("PEPS   : " & objPersona.Peps, nMax - 9) & "        " & oImp.gPrnSaltoLinea
            'WIOR FIN *********************************************************
            sCad = sCad & oImp.gPrnSaltoLinea
        Next
        
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA ORDENA TRANSACCIÓN ------", nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA ORDENA LA TRANSACCIÓN ------", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        'WIOR 20130301 COMENTO LINEAS ABAJO ************
        'sCad = sCad & oImpre.Justifica("NOMBRE : " & sOrdNombre, nMax) & Space(nSep) & oImpre.Justifica("NOMBRE : " & sOrdNombre, nMax) & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("DOC ID : " & sOrdDocId, nMax) & Space(nSep) & oImpre.Justifica("DOC ID : " & sOrdDocId, nMax) & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("DIREC. : " & sOrdDireccion, nMax) & Space(nSep) & oImpre.Justifica("DIREC. : " & sOrdDireccion, nMax) & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("OCUPA. : " & sOrdOcupacion, nMax) & Space(nSep) & oImpre.Justifica("OCUPA. : " & sOrdOcupacion, nMax) & oImp.gPrnSaltoLinea 'JACA 20110225
        'sCad = sCad & oImp.gPrnSaltoLinea 'JACA 20110225
        'WIOR FIN  **************************************
        
        'EJVG20120329
        For iOrd = 0 To UBound(poListaOrdenan)
            objPersona = poListaOrdenan(iOrd)
            sCad = sCad & oImpre.Justifica("NOMBRE : " & objPersona.Nombre, nMax) & Space(nSep) & oImpre.Justifica("NOMBRE : " & objPersona.Nombre, nMax) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("DOC ID : " & objPersona.DocumentoId, nMax - 9) & "---------" & Space(nSep) & oImpre.Justifica("DOC ID : " & objPersona.DocumentoId, nMax - 9) & "---------" & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("DIREC. : " & objPersona.Direccion, nMax - 9) & "         " & Space(nSep) & oImpre.Justifica("DIREC. : " & objPersona.Direccion, nMax - 9) & "        " & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("OCUPA. : " & objPersona.Ocupacion, nMax - 9) & "         " & Space(nSep) & oImpre.Justifica("OCUPA. : " & objPersona.Ocupacion, nMax - 9) & "        " & oImp.gPrnSaltoLinea
            'WIOR 20130301 ****************************************************
            sCad = sCad & oImpre.Justifica("NAC.   : " & objPersona.Nacionalidad, nMax - 9) & "         " & Space(nSep) & oImpre.Justifica("NAC.   : " & objPersona.Nacionalidad, nMax - 9) & "        " & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("RESIDE : " & objPersona.Residente, nMax - 9) & "         " & Space(nSep) & oImpre.Justifica("RESIDE : " & objPersona.Residente, nMax - 9) & "        " & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("PEPS   : " & objPersona.Peps, nMax - 9) & "         " & Space(nSep) & oImpre.Justifica("PEPS   : " & objPersona.Peps, nMax - 9) & "        " & oImp.gPrnSaltoLinea
            'WIOR FIN *********************************************************
            sCad = sCad & oImp.gPrnSaltoLinea
        Next
        
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA BENEFICIA TRANSACCIÓN ------", nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA BENEFICIA TRANSACCIÓN ------", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        'WIOR 20130301 COMENTO LINEAS ABAJO ************
        'sCad = sCad & oImpre.Justifica("NOMBRE : " & sBenNombre, nMax) & Space(nSep) & oImpre.Justifica("NOMBRE : " & sBenNombre, nMax) & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("DOC ID : " & sBenDocId, nMax) & Space(nSep) & oImpre.Justifica("DOC ID : " & sBenDocId, nMax) & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("DIREC. : " & sBenDireccion, nMax) & Space(nSep) & oImpre.Justifica("DIREC. : " & sBenDireccion, nMax) & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("OCUPA. : " & sBenOcupacion, nMax) & Space(nSep) & oImpre.Justifica("OCUPA. : " & sBenOcupacion, nMax) & oImp.gPrnSaltoLinea 'JACA 20110225
        'sCad = sCad & oImp.gPrnSaltoLinea 'JACA 20110225
        ''JACA 20110224
        '    If sBenNombre2 <> "" Then
        '        sCad = sCad & oImpre.Justifica("NOMBRE : " & sBenNombre2, nMax) & Space(nSep) & oImpre.Justifica("NOMBRE : " & sBenNombre2, nMax) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("DOC ID : " & sBenDocId2, nMax) & Space(nSep) & oImpre.Justifica("DOC ID : " & sBenDocId2, nMax) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("DIREC. : " & sBenDireccion2, nMax) & Space(nSep) & oImpre.Justifica("DIREC. : " & sBenDireccion2, nMax) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("OCUPA. : " & sBenOcupacion2, nMax) & Space(nSep) & oImpre.Justifica("OCUPA. : " & sBenOcupacion2, nMax) & oImp.gPrnSaltoLinea 'JACA 20110225
        '        sCad = sCad & oImp.gPrnSaltoLinea 'JACA 20110225
        '    End If
        '    If sBenNombre3 <> "" Then
        '        sCad = sCad & oImpre.Justifica("NOMBRE : " & sBenNombre3, nMax) & Space(nSep) & oImpre.Justifica("NOMBRE : " & sBenNombre3, nMax) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("DOC ID : " & sBenDocId3, nMax) & Space(nSep) & oImpre.Justifica("DOC ID : " & sBenDocId3, nMax) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("DIREC. : " & sBenDireccion3, nMax) & Space(nSep) & oImpre.Justifica("DIREC. : " & sBenDireccion3, nMax) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("OCUPA. : " & sBenOcupacion3, nMax) & Space(nSep) & oImpre.Justifica("OCUPA. : " & sBenOcupacion3, nMax) & oImp.gPrnSaltoLinea 'JACA 20110225
        '        sCad = sCad & oImp.gPrnSaltoLinea 'JACA 20110225
        '    End If
        '    If sBenNombre4 <> "" Then
        '        sCad = sCad & oImpre.Justifica("NOMBRE : " & sBenNombre4, nMax) & Space(nSep) & oImpre.Justifica("NOMBRE : " & sBenNombre4, nMax) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("DOC ID : " & sBenDocId4, nMax) & Space(nSep) & oImpre.Justifica("DOC ID : " & sBenDocId4, nMax) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("DIREC. : " & sBenDireccion4, nMax) & Space(nSep) & oImpre.Justifica("DIREC. : " & sBenDireccion4, nMax) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("OCUPA. : " & sBenOcupacion4, nMax) & Space(nSep) & oImpre.Justifica("OCUPA. : " & sBenOcupacion4, nMax) & oImp.gPrnSaltoLinea 'JACA 20110225
        '        sCad = sCad & oImp.gPrnSaltoLinea 'JACA 20110225
        '    End If
        ''JACA END
        'WIOR FIN  **************************************
        'WIOR 20130301 ****************************************************
         For iBen = 0 To UBound(poListaBeneficia)
            objPersona = poListaBeneficia(iBen)
            sCad = sCad & oImpre.Justifica("NOMBRE : " & objPersona.Nombre, nMax) & Space(nSep) & oImpre.Justifica("NOMBRE : " & objPersona.Nombre, nMax) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("DOC ID : " & objPersona.DocumentoId, nMax) & Space(nSep) & oImpre.Justifica("DOC ID : " & objPersona.DocumentoId, nMax) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("DIREC. : " & objPersona.Direccion, nMax) & Space(nSep) & oImpre.Justifica("DIREC. : " & objPersona.Direccion, nMax) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("OCUPA. : " & objPersona.Ocupacion, nMax) & Space(nSep) & oImpre.Justifica("OCUPA. : " & objPersona.Ocupacion, nMax) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("NAC.   : " & objPersona.Nacionalidad, nMax) & Space(nSep) & oImpre.Justifica("NAC.   : " & objPersona.Nacionalidad, nMax) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("RESIDE : " & objPersona.Residente, nMax) & Space(nSep) & oImpre.Justifica("RESIDE : " & objPersona.Residente, nMax) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("PEPS   : " & objPersona.Peps, nMax) & Space(nSep) & oImpre.Justifica("PEPS   : " & objPersona.Peps, nMax) & oImp.gPrnSaltoLinea 'JACA 20110225
            sCad = sCad & oImp.gPrnSaltoLinea
        Next
        'WIOR FIN *********************************************************
        
        '**************************
        '**************************
        '***************************
''        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("-------------  TITULAR  ---------------", nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("-------------  TITULAR  ---------------", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
''        sCad = sCad & oImpre.Justifica("NOMBRE : " & sTitNombre, nMax) & Space(nSep) & oImpre.Justifica("NOMBRE : " & sTitNombre, nMax) & oImp.gPrnSaltoLinea
''        sCad = sCad & oImpre.Justifica("DOC ID : " & sTitDocId, nMax) & Space(nSep) & oImpre.Justifica("DOC ID : " & sTitDocId, nMax) & oImp.gPrnSaltoLinea
''        sCad = sCad & oImpre.Justifica("DIREC. : " & sTitDireccion, nMax) & Space(nSep) & oImpre.Justifica("DIREC. : " & sTitDireccion, nMax) & oImp.gPrnSaltoLinea
''
''        'By Capi 20012008
''        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA ORDENA TRANSACCIÓN ------", nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA ORDENA LA TRANSACCIÓN ------", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
''        sCad = sCad & oImpre.Justifica("NOMBRE : " & sOrdNombre, nMax) & Space(nSep) & oImpre.Justifica("NOMBRE : " & sOrdNombre, nMax) & oImp.gPrnSaltoLinea
''        sCad = sCad & oImpre.Justifica("DOC ID : " & sOrdDocId, nMax) & Space(nSep) & oImpre.Justifica("DOC ID : " & sOrdDocId, nMax) & oImp.gPrnSaltoLinea
''        sCad = sCad & oImpre.Justifica("DIREC. : " & sOrdDireccion, nMax) & Space(nSep) & oImpre.Justifica("DIREC. : " & sOrdDireccion, nMax) & oImp.gPrnSaltoLinea
''
''
''
''        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA REALIZA TRANSACCIÓN  ----", nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA REALIZA TRANSACCIÓN  ----", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
''        sCad = sCad & oImpre.Justifica("NOMBRE : " & sReaNombre, nMax) & Space(nSep) & oImpre.Justifica("NOMBRE : " & sReaNombre, nMax) & oImp.gPrnSaltoLinea
''        sCad = sCad & oImpre.Justifica("DOC ID : " & sReaDocId, nMax - 9) & "---------" & Space(nSep) & oImpre.Justifica("DOC ID : " & sReaDocId, nMax - 9) & "---------" & oImp.gPrnSaltoLinea
''        sCad = sCad & oImpre.Justifica("DIREC. : " & sReaDireccion, nMax - 9) & "  (FIRMA)" & Space(nSep) & oImpre.Justifica("DIREC. : " & sReaDireccion, nMax - 9) & "  (FIRMA)" & oImp.gPrnSaltoLinea
''
''        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA BENEFICIA TRANSACCIÓN ------", nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA BENEFICIA TRANSACCIÓN ------", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
''        sCad = sCad & oImpre.Justifica("NOMBRE : " & sBenNombre, nMax) & Space(nSep) & oImpre.Justifica("NOMBRE : " & sBenNombre, nMax) & oImp.gPrnSaltoLinea
''        sCad = sCad & oImpre.Justifica("DOC ID : " & sBenDocId, nMax) & Space(nSep) & oImpre.Justifica("DOC ID : " & sBenDocId, nMax) & oImp.gPrnSaltoLinea
''        sCad = sCad & oImpre.Justifica("DIREC. : " & sBenDireccion, nMax) & Space(nSep) & oImpre.Justifica("DIREC. : " & sBenDireccion, nMax) & oImp.gPrnSaltoLinea
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("----------DATOS DE LA OPERACION----------", nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("----------DATOS DE LA OPERACION----------", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        If Len(sCuenta) = 18 Then
            '''sOpe = "TIPO   : " & sOperacion & " (" & IIf(CLng(Mid(sCuenta, 9, 1)) = gMonedaNacional, "SOLES", IIf(CLng(Mid(sCuenta, 9, 1)) = 2, "DOLARES", "EUROS")) & ")" 'marg ers044-2016
            sOpe = "TIPO   : " & sOperacion & " (" & IIf(CLng(Mid(sCuenta, 9, 1)) = gMonedaNacional, StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), IIf(CLng(Mid(sCuenta, 9, 1)) = 2, "DOLARES", "EUROS")) & ")" 'marg ers044-2016
        Else
            If nMoneda = 1 Then
                '''sOpe = sOperacion & " (SOLES)" 'marg ers044-2016
                sOpe = sOperacion & " (" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & ")" 'marg ers044-2016
            ElseIf nMoneda = 2 Then
                sOpe = sOperacion & " (DOLARES)"
            '*** PEAC 20090727
            ElseIf nMoneda = 3 Then
                sOpe = sOperacion & " (EUROS)"
                
            End If
        End If
        
        If Len(sOpe) <= nMax Then
           sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        Else
           '''sOpe = Left(sOperacion, nMax - 10) & " (" & IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", IIf(Mid(sCuenta, 9, 1) = "2", "DOLARES", "EUROS")) & ")" 'marg ers044-2016
           sOpe = Left(sOperacion, nMax - 10) & " (" & IIf(Mid(sCuenta, 9, 1) = "1", StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), IIf(Mid(sCuenta, 9, 1) = "2", "DOLARES", "EUROS")) & ")" 'marg ers044-2016
           sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        End If
        'JACA 20110225
        If nTipoREU = 1 Then
             sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("TPO.RO : OPE. UNICO", nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("TPO.RO :  OPE. UNICO", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        ElseIf nTipoREU = 2 Then
            sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("TPO.RO : OPE. MULTIPLE", nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("TPO.RO :  OPE. MULTIPLE", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        End If
        
        'RECO20140311 ERS160-2013************************************************************************************
            'WIOR 20150708 COMENTO ESTA PARTE***
            'Dim oCtaTit As ADODB.Recordset
            'Dim oCtaTpo As ADODB.Recordset
            'Dim cTpoPers As String
            'Dim indx As Integer
            'Dim nCtaTpo As Integer
            
            'Set oCtaTit = New ADODB.Recordset
            'Set oCtaTpo = New ADODB.Recordset
            'If Mid(sCuenta, 6, 3) = "232" Or Mid(sCuenta, 6, 3) = "233" Or Mid(sCuenta, 6, 3) = "234" Then
            '    Dim oCap As New COMDCaptaGenerales.DCOMCaptaGenerales
            '    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
            '
            '    Set oCtaTit = oCap.GetProductoPersona(sCuenta)
            '    Set oCtaTpo = oCap.GetDatosCuentaAho(sCuenta)
            '
            '    nCtaTpo = 1
            'ElseIf Mid(sCuenta, 6, 3) = "239" Then
            '   Dim oCapSer As New COMDCaptaServicios.DCOMCaptaServicios
            '   Set oCapSer = New COMDCaptaServicios.DCOMCaptaServicios
            '   Set oCtaTit = oCapSer.RecuperaDatosRelacGiro(sCuenta)
            '   nCtaTpo = 2
            'ElseIf sCuenta = "" Then
            '   nCtaTpo = 2
            'Else
            '    'Dim oCred As New COMDCredito.DCOMCredito
            '    'Set oCred = New COMDCredito.DCOMCredito
            '
            '    Set oCtaTit = RecuperaDatosRelacCred(sCuenta)
            '    Set oCtaTpo = RecuperaDatosCreditoVigente(sCuenta)
            '    nCtaTpo = 3
            'End If
            'WIOR FIN ***
            If nCtaTpo <> 2 Then
                If Not (oCtaTit.BOF And oCtaTit.EOF) Then
                    For indx = 1 To oCtaTit.RecordCount
                        cTpoPers = RTrim(Mid(oCtaTit!cRelacion, 1, 8))
                            'If (Val(Mid(oCtaTit!cRelacion, Len(oCtaTit!cRelacion) - 3, Len(oCtaTit!cRelacion))) = 10) Then
                            'sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("TIT.CTA : " & oCtaTit!Nombre, nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("TIT.CTA : " & oCtaTit!Nombre, nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                            sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica(cTpoPers & Space(IIf(Len(cTpoPers) < 8, 8 - Len(cTpoPers), 0)) & ": " & oCtaTit!Nombre, nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica(cTpoPers & ": " & oCtaTit!Nombre, nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                        'End If
                        oCtaTit.MoveNext
                    Next
                End If
            Else
                sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("TIT.CTA : ", nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("TIT.CTA : ", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            End If
            If nCtaTpo <> 2 Then
                sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("CUENTA  :", nMax - 19) & " " & Right(Space(14) & Format(sCuenta), 18) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("CUENTA  :", nMax - 19) & " " & Right(Space(14) & Format(sCuenta), 18) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            Else
                sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("CUENTA  :", nMax - 19) & " " & Right(Space(32) & Format(sCuenta), 18) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("CUENTA  :", nMax - 19) & " " & Right(Space(14) & Format(sCuenta), 18) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            End If
            If nCtaTpo <> 2 Then
                If Not (oCtaTpo.BOF And oCtaTpo.EOF) Then
                    If nCtaTpo = 1 Then 'CAPTACIONES
                        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("TPO.CTA :", nMax - 19) & " " & Right(Space(14) & oCtaTpo!cTipoCuenta, 18) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("TPO.CTA :", nMax - 19) & " " & Right(Space(14) & oCtaTpo!cTipoCuenta, 18) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                    Else 'COLOCACIONES
                        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("TPO.CTA :", nMax - 19) & " " & Right(Space(14) & oCtaTpo!cTpoCredDes, 18) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("TPO.CTA :", nMax - 19) & " " & Right(Space(14) & oCtaTpo!cTpoCredDes, 18) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                    End If
                End If
            Else
                sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("TPO.CTA :", nMax - 19) & " " & Right(Space(32) & "", 18) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("TPO.CTA :", nMax - 19) & " " & Right(Space(14) & "", 18) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            End If
        'RECO FIN****************************************************************************************************

        'JACA END
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("MONTO   :", nMax - 15) & " " & Right(Space(14) & Format(nImporte, "##,##0.00"), 14) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("MONTO   :", nMax - 15) & " " & Right(Space(14) & Format(nImporte, "##,##0.00"), 14) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        'sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("ORIGEN EFECTIVO", nMax - 15) & " " & Right(Space(14) & Format(psOrigenEfectivo), 14) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("ORIGEN EFECTIVO", nMax - 15) & " " & Right(Space(14) & Format(psOrigenEfectivo), 14) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        'WIOR 20120709
        Dim linea As String
        Dim tamano As Integer
        Dim blanco As String
        Dim nValidar As Double
        Dim nTamano As Integer
        Dim NumLineas As Integer
        Dim i As Integer
        Dim hecho As Integer
        Dim NumEspacios As Integer
        Dim final As String
        Dim recorte As String
        Dim palabra As String
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("ORIGEN EFECTIVO:", nMax - 15) & " " & Space(14) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("ORIGEN EFECTIVO:", nMax - 15) & " " & Space(14) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        tamano = Len(psOrigenEfectivo)
        nTamano = Len(psOrigenEfectivo)
        palabra = Trim(psOrigenEfectivo)
        linea = Mid(palabra, 1, 41)
        blanco = " "
        If Len(palabra) > 41 Then
            blanco = Mid(palabra, 42, 1)
        End If
        
        
        nValidar = nTamano / 41
        NumLineas = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
        
        For i = 0 To NumLineas - 1
            If Trim(linea) <> "" Then
                recorte = linea
                NumEspacios = 0
                If blanco <> " " Then
                    hecho = InStr(1, linea, " ")
                    Do While hecho <> 0
                        NumEspacios = NumEspacios + hecho
                        recorte = Mid(recorte, hecho + 1, 41)
                        hecho = InStr(1, recorte, " ")
                    Loop
                    
                    final = Mid(linea, 1, NumEspacios - 1)
                    palabra = Mid(palabra, Len(final) + 1, Len(palabra) - Len(final))
                    If Len(palabra) > 41 Then
                        blanco = Mid(palabra, 42, 1)
                        linea = Mid(palabra, 1, 41)
                    Else
                        blanco = " "
                        linea = palabra
                    End If
                Else
                    final = linea
                    palabra = Mid(palabra, Len(final) + 1, Len(palabra) - Len(final))
                    If Len(palabra) > 41 Then
                        blanco = Mid(palabra, 42, 1)
                        linea = Mid(palabra, 1, 41)
                    Else
                        blanco = " "
                        linea = palabra
                    End If
                End If
                
                sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica(Trim(UCase(final)), nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica(Trim(UCase(final)), nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            End If
        Next i
        
        'sCad = sCad & oImp.gPrnNegritaON & oImpre.JustificaTextoSinCarEsp(psOrigenEfectivo, nMax) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImp.gPrnSaltoLinea '& oImpre.JustificaTextoSinCarEsp(psOrigenEfectivo, (nMax - 15) + 15) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        'WIOR -FIN
        'By Capi 27122008 se agrego Fecha
        'sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("FECHA PROCESO : ", nMax - 15) & Format(dFecha, "dd/mm/yyyy") & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("FECHA PROCESO : ", nMax - 15) & Format(dFecha, "dd/mm/yyyy") & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea' Comentado by JACA 20110930
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("FECHA PROCESO : ", nMax - 15) & " " & Right(Space(14) & Format(dFecha, "dd/mm/yyyy"), 14) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("FECHA PROCESO : ", nMax - 15) & " " & Right(Space(14) & Format(dFecha, "dd/mm/yyyy"), 14) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        
        sCad = sCad & String(nMax, "-") & Space(nSep) & String(nMax, "-") & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("FIRMA.:", nMax) & " " & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("FIRMA:", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica(" ", nMax - 12) & " " & Format$(Time, "hh:mm:ss AMPM") & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica(" ", nMax - 12) & " " & Format$(Time, "hh:mm:ss AMPM") & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        
        'sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("TIPO CUENTA", nMax - 15) & " " & Right(Space(14) & sTipo, 14) & oImp.gPrnNegritaOFF & Space(nSep) & oImp.gPrnNegritaON & oImpre.Justifica("TIPO CUENTA", nMax - 15) & " " & Right(Space(14) & sTipo, 14) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        sCad = sCad & String(nMax, "=") & Space(nSep) & String(nMax, "=") & oImp.gPrnSaltoLinea
        sCad = oImpre.ImpreCarEsp(sCad)
        
        'sCad = sCad & oImp.gPrnSaltoLinea
        'By Capi 14022008
        sCad = sCad & String(nMax, "-") & Space(nSep) & String(nMax, "-") & oImp.gPrnSaltoLinea
        
        
        sCad = sCad & oImp.gPrnSaltoPagina
    Else
        'WIOR 20120709 SE AJUSTO EN GENERAL LA IMPRESION EN LA TIKETERA SEGUN RFC-OYP059-2012
        sCad = Replace(sCad, "k", "")
        sCad = sCad & oImpre.Justifica(sNomCmac, nMax - 13) & " " & Left(dFecha, 10) & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.Justifica("PROC.:" & Trim(sNomAge) & "/" & sAgeRem, nMax - 14) & " " & Format$(Time, "hh:mm:ss AMPM") & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("REU: " & lsCorrLavDinero & " (Cta: " & sCuenta & ") ", nMax - 1) & oImp.gPrnSaltoLinea
        'WIOR 20120710
        sCad = sCad & oImpre.Justifica("REU  : " & lsCorrLavDinero, nMax - 1) & oImp.gPrnSaltoLinea

'        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("-------------  TITULAR  ---------------", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
'        sCad = sCad & oImpre.Justifica("NOMBRE : " & sTitNombre, nMax - 1) & oImp.gPrnSaltoLinea
'        sCad = sCad & oImpre.Justifica("DOC ID : " & sTitDocId, nMax - 1) & oImp.gPrnSaltoLinea
'        sCad = sCad & oImpre.Justifica("DIREC. : " & sTitDireccion, nMax - 1) & oImp.gPrnSaltoLinea
'        sCad = sCad & oImpre.Justifica("OCUPA. : " & sTitOcupacion, nMax - 1) & oImp.gPrnSaltoLinea 'JACA 20110225
'        sCad = sCad & oImp.gPrnSaltoLinea 'JACA 20110225
        
        'By Capi 20012008
        
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA ORDENA TRANSACCIÓN  ----", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        'WIOR 20130301 COMENTO LINEAS ABAJO ************
        'sCad = sCad & oImpre.Justifica("NOMBRE : " & sOrdNombre, nMax - 1) & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("DOC ID : " & sOrdDocId, nMax - 9) & oImp.gPrnSaltoLinea
        ''sCad = sCad & oImpre.Justifica("DIREC. : " & sOrdDireccion, nMax - 9) & "  (FIRMA)" & oImp.gPrnSaltoLinea COMENTADO X JACA 20110225
        'sCad = sCad & oImpre.Justifica("DIREC. : " & sOrdDireccion, nMax - 10) & "---------" & oImp.gPrnSaltoLinea ''JACA 20110225
        'sCad = sCad & oImpre.Justifica("OCUPA. : " & sOrdOcupacion, nMax - 10) & "  (FIRMA)" & oImp.gPrnSaltoLinea 'JACA 20110225
        'sCad = sCad & oImp.gPrnSaltoLinea 'JACA 20110225
        'WIOR FIN  **************************************
        
        'EJVG20120329
        For iOrd = 0 To UBound(poListaOrdenan)
            objPersona = poListaOrdenan(iOrd)
            sCad = sCad & oImpre.Justifica("NOMBRE : " & objPersona.Nombre, nMax - 1) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("DOC ID : " & objPersona.DocumentoId, nMax - 9) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("DIREC. : " & objPersona.Direccion, nMax - 10) & "---------" & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("OCUPA. : " & objPersona.Ocupacion, nMax - 10) & "  (FIRMA)" & oImp.gPrnSaltoLinea
            'WIOR 20130301 ****************************************************
            sCad = sCad & oImpre.Justifica("NAC.   : " & objPersona.Nacionalidad, nMax - 10) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("RESIDE : " & objPersona.Residente, nMax - 10) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("PEPS   : " & objPersona.Peps, nMax - 10) & oImp.gPrnSaltoLinea
            'WIOR FIN *********************************************************
            sCad = sCad & oImp.gPrnSaltoLinea
        Next
        
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA REALIZA TRANSACCIÓN  ----", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        'WIOR 20130301 COMENTO LINEAS ABAJO ************
        'sCad = sCad & oImpre.Justifica("NOMBRE : " & sReaNombre, nMax - 1) & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("DOC ID : " & sReaDocId, nMax - 9) & oImp.gPrnSaltoLinea
        ''sCad = sCad & oImpre.Justifica("DIREC. : " & sReaDireccion, nMax - 9) & "  (FIRMA)" & oImp.gPrnSaltoLinea COMENTADO X JACA 20110225
        'sCad = sCad & oImpre.Justifica("DIREC. : " & sReaDireccion, nMax - 10) & "---------" & oImp.gPrnSaltoLinea  'JACA 20110225
        'sCad = sCad & oImpre.Justifica("OCUPA. : " & sReaOcupacion, nMax - 10) & "  (FIRMA)" & oImp.gPrnSaltoLinea 'JACA 20110225
        'sCad = sCad & oImp.gPrnSaltoLinea 'JACA 20110225
        'WIOR FIN  **************************************
        
        'EJVG20120329
        For iRea = 0 To UBound(poListaRealizan)
            objPersona = poListaRealizan(iRea)
            sCad = sCad & oImpre.Justifica("NOMBRE : " & objPersona.Nombre, nMax - 1) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("DOC ID : " & objPersona.DocumentoId, nMax - 9) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("DIREC. : " & objPersona.Direccion, nMax - 10) & "---------" & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("OCUPA. : " & objPersona.Ocupacion, nMax - 10) & "  (FIRMA)" & oImp.gPrnSaltoLinea
            'WIOR 20130301 ****************************************************
            sCad = sCad & oImpre.Justifica("NAC.   : " & objPersona.Nacionalidad, nMax - 10) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("RESIDE : " & objPersona.Residente, nMax - 10) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("PEPS   : " & objPersona.Peps, nMax - 10) & oImp.gPrnSaltoLinea
            'WIOR FIN *********************************************************
            sCad = sCad & oImp.gPrnSaltoLinea
        Next
        
        sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("---- PERSONA BENEFICIA TRANSACCIÓN ------", nMax) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        'WIOR 20130301 COMENTO LINEAS ABAJO ******
        'sCad = sCad & oImpre.Justifica("NOMBRE : " & sBenNombre, nMax - 1) & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("DOC ID : " & sBenDocId, nMax - 1) & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("DIREC. : " & sBenDireccion, nMax - 1) & oImp.gPrnSaltoLinea
        'sCad = sCad & oImpre.Justifica("OCUPA. : " & sBenOcupacion, nMax - 1) & oImp.gPrnSaltoLinea 'JACA 20110225
        'sCad = sCad & oImp.gPrnSaltoLinea 'JACA 20110225
        
        ''JACA 20110224
        '    If sBenNombre2 <> "" Then
        '        sCad = sCad & oImpre.Justifica("NOMBRE : " & sBenNombre2, nMax - 1) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("DOC ID : " & sBenDocId2, nMax - 1) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("DIREC. : " & sBenDireccion2, nMax - 1) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("OCUPA. : " & sBenOcupacion2, nMax - 1) & oImp.gPrnSaltoLinea 'JACA 20110225
        '        sCad = sCad & oImp.gPrnSaltoLinea 'JACA 20110225
        '    End If
        '    If sBenNombre3 <> "" Then
        '        sCad = sCad & oImpre.Justifica("NOMBRE : " & sBenNombre3, nMax - 1) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("DOC ID : " & sBenDocId3, nMax - 1) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("DIREC. : " & sBenDireccion3, nMax - 1) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("OCUPA. : " & sBenOcupacion3, nMax - 1) & oImp.gPrnSaltoLinea 'JACA 20110225
        '        sCad = sCad & oImp.gPrnSaltoLinea 'JACA 20110225
        '    End If
        '    If sBenNombre4 <> "" Then
        '        sCad = sCad & oImpre.Justifica("NOMBRE : " & sBenNombre4, nMax - 1) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("DOC ID : " & sBenDocId4, nMax - 1) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("DIREC. : " & sBenDireccion4, nMax - 1) & oImp.gPrnSaltoLinea
        '        sCad = sCad & oImpre.Justifica("OCUPA. : " & sBenOcupacion4, nMax - 1) & oImp.gPrnSaltoLinea 'JACA 20110225
        '        sCad = sCad & oImp.gPrnSaltoLinea 'JACA 20110225
        '    End If
        ''JACA END
        'WIOR FIN  **************************************
        'WIOR 20130301 ***************************************************
        For iBen = 0 To UBound(poListaBeneficia)
            objPersona = poListaBeneficia(iBen)
            sCad = sCad & oImpre.Justifica("NOMBRE : " & objPersona.Nombre, nMax - 1) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("DOC ID : " & objPersona.DocumentoId, nMax - 1) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("DIREC. : " & objPersona.Direccion, nMax - 1) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("OCUPA. : " & objPersona.Ocupacion, nMax - 1) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("NAC.   : " & objPersona.Nacionalidad, nMax - 1) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("RESIDE : " & objPersona.Residente, nMax - 1) & oImp.gPrnSaltoLinea
            sCad = sCad & oImpre.Justifica("PEPS   : " & objPersona.Peps, nMax - 1) & oImp.gPrnSaltoLinea
            sCad = sCad & oImp.gPrnSaltoLinea
        Next
        'WIOR FIN *********************************************************
        
        If Len(sCuenta) = 18 Then
            '''sOpe = sOperacion & "(" & IIf(CLng(Mid(sCuenta, 9, 1)) = gMonedaNacional, "SOLES", IIf(CLng(Mid(sCuenta, 9, 1)) = 2, "DOLARES", "EUROS")) & ")" 'marg ers044-2016
            sOpe = sOperacion & "(" & IIf(CLng(Mid(sCuenta, 9, 1)) = gMonedaNacional, StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), IIf(CLng(Mid(sCuenta, 9, 1)) = 2, "DOLARES", "EUROS")) & ")" 'marg ers044-2016
        Else
            If nMoneda = 1 Then
                '''sOpe = sOperacion & "(SOLES)" 'marg ers044-2016
                sOpe = sOperacion & "(" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & ")" 'marg ers044-2016
            ElseIf nMoneda = 2 Then
                sOpe = sOperacion & "(DOLARES)"
            '*** PEAC 20090727
            ElseIf nMoneda = 3 Then
                sOpe = sOperacion & "(EUROS)"
            End If
        End If
        
        'WIOR 20120710
        'sCad = sCad & String(nMax - 1, "-") & oImp.gPrnSaltoLinea'WIOR 20150708 COMENTO
        'WIOR FIN
                
        'WIOR 20120710
        'sCad = sCad & oImpre.Justifica("CTA.    : " & sCuenta, nMax - 1) & oImp.gPrnSaltoLinea'WIOR 20150708 COMENTO
        'sCad = sCad & oImpre.Justifica("TIT. CTA: " & sTitNombre, nMax - 1) & oImp.gPrnSaltoLinea'WIOR 20150708 COMENTO
        'WIOR FIN
        
        'WIOR 20150708 ***
        sCad = sCad & oImpre.Justifica("----------DATOS DE LA OPERACION----------", nMax - 1) & oImp.gPrnSaltoLinea
        
        If Len(sOpe) <= nMax Then
           sCad = sCad & oImpre.Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & oImp.gPrnSaltoLinea
        Else
           '''sOpe = Left(sOperacion, nMax - 10) & " (" & IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", IIf(Mid(sCuenta, 9, 1) = "2", "DOLARES", "EUROS")) & ")" 'marg ers044-2016
           sOpe = Left(sOperacion, nMax - 10) & " (" & IIf(Mid(sCuenta, 9, 1) = "1", StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), IIf(Mid(sCuenta, 9, 1) = "2", "DOLARES", "EUROS")) & ")" 'marg ers044-2016
           sCad = sCad & oImpre.Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & oImp.gPrnSaltoLinea
        End If
        
        'MAGR 20160718
         'JACA 20110225
        If nTipoREU = 1 Then
            sCad = sCad & oImpre.Justifica("TPO.RO : OPE. UNICO", nMax - 1) & oImp.gPrnSaltoLinea
        ElseIf nTipoREU = 2 Then
            sCad = sCad & oImpre.Justifica("TPO.RO : OPE. MULTIPLE", nMax - 1) & oImp.gPrnSaltoLinea
        End If
        'JACA END
        
        If nCtaTpo <> 2 Then
            If Not (oCtaTit.BOF And oCtaTit.EOF) Then
                For indx = 1 To oCtaTit.RecordCount
                    cTpoPers = RTrim(Mid(oCtaTit!cRelacion, 1, 8))
                    sCad = sCad & oImpre.Justifica(cTpoPers & Space(IIf(Len(cTpoPers) < 8, 8 - Len(cTpoPers), 0)) & ": " & oCtaTit!Nombre, nMax - 1) & oImp.gPrnSaltoLinea
                    oCtaTit.MoveNext
                Next
            End If
        Else
            sCad = sCad & oImpre.Justifica("TIT. CTA: " & sTitNombre, nMax - 1) & oImp.gPrnSaltoLinea
        End If
        
        sCad = sCad & oImpre.Justifica("CTA.    : " & sCuenta, nMax - 1) & oImp.gPrnSaltoLinea

        
        If nCtaTpo <> 2 Then
            If Not (oCtaTpo.BOF And oCtaTpo.EOF) Then
                If nCtaTpo = 1 Then 'CAPTACIONES
                    sCad = sCad & oImpre.Justifica("TPO. CTA: " & oCtaTpo!cTipoCuenta, nMax - 1) & oImp.gPrnSaltoLinea
                Else 'COLOCACIONES
                    sCad = sCad & oImpre.Justifica("TPO. CTA: " & oCtaTpo!cTpoCredDes, nMax - 1) & oImp.gPrnSaltoLinea
                End If
            End If
        Else
            sCad = sCad & oImpre.Justifica("TPO. CTA: " & "", nMax - 1) & oImp.gPrnSaltoLinea
        End If
        'WIOR FIN ********
        
        sCad = sCad & oImpre.Justifica("MONTO   :", nMax - 16) & " " & Right(Space(14) & Format(nImporte, "##,##0.00"), 14) & oImp.gPrnSaltoLinea
                        
        'WIOR 20120710 AGREGO LA DESCRIPCION DEL ORIGEN DEL EFECTIVO
        sCad = sCad & oImpre.Justifica("ORIGEN EFECTIVO:", nMax - 1) & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.JustificaTextoSinCarEsp(Trim(UCase(psOrigenEfectivo)), nMax - 1) & oImp.gPrnSaltoLinea
        'WIOR FIN
        'sCad = sCad & oImp.gPrnNegritaON & oImpre.Justifica("TIPO CUENTA", nMax - 15) & " " & Right(Space(14) & sTipo, 14) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        
        sCad = sCad & String(nMax - 1, "=") & oImp.gPrnSaltoLinea
        sCad = oImpre.ImpreCarEsp(sCad)
        
        sCad = sCad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
       
    End If
If bCadena Then
    ImprimeBoletaLavadoDinero = sCad
End If

End Function

Public Function ImprimeBoleta(ByVal psTit As String, ByVal pstext As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
        ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
        ByVal pnInteresA As String, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
        Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
        Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, _
        Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
        Optional pdFecSis As Date, Optional sNomAge As String = "", Optional sCodUser As String = "", _
        Optional sLpt As String = "", Optional bCadena As Boolean = False, Optional bPreguntaSaldo As Boolean = True, _
        Optional pnITFValor As Currency = 0, Optional pbImprimeSaldos As Boolean = True, Optional pnMovNro As Long, _
        Optional psDNI = "", Optional psMot As Integer, Optional ByVal pbImpTMU As Boolean = False) As String
'SE AGREGO ByVal psMot As Integer PARA PONER GLOSA ---ANPS

    Dim nFicSal As Integer
    Dim sFecha As String
    Dim sHora As String
    Dim sSep As Integer
    Dim sIni As Integer
    Dim sMonto As String
    Dim sSDisp As String
    Dim sIntAcum As String
    Dim sMax As Integer
    Dim saux As Integer
    Dim lsNroExt As String
    Dim lnTope As Integer
    Dim lsSaldoC As String
    Dim lsCadArg As String
    Dim lsInteres As String
    Dim sCad As String
    Dim lnCliAux As Long
    Dim lsCliAux1 As String
    Dim lsCliAux2 As String

    Dim lnChq As Long
    Dim lsChqAux1 As String
    Dim lsChqAux2 As String
    Dim lsNomAge As String

    Dim lnNumLinCmac As Integer

    Dim lsMensaje As String * 39

    Dim oImpre As New COMFunciones.FCOMImpresion

ETIQ:

    On Error GoTo Error

    lnTope = 0 '6 'Tope de lineas en Boleta

    lsNroExt = Str(pnNumExt)


    sCad = sCad & oImp.gPrnInicializa
    sCad = sCad & oImp.gPrnEspaLineaN   'espaciamiento lineas 1/6 pulg.
    sCad = sCad & oImp.gPrnTamPagina22  'Longitud de página a 22 líneas'
    sCad = sCad & oImp.gPrnTamLetra10CPI   'Tamaño 10 cpi
    sCad = sCad & oImp.gPrnTpoLetraRoman     'Tipo de Letra Sans Serif
    sCad = sCad & oImp.gPrnCondensadaOFF ' cancela condensada
    sCad = sCad & oImp.gPrnNegritaOFF ' desactiva negrita

    sSep = 15
    sIni = 1
    sMax = 33
    saux = 5

    sFecha = Format$(pdFecSis, "dd/mm/yyyy")
    sHora = Format$(Time, "hh:mm:ss")
    sMonto = Format$(pnMonto, "#,##0.00")
    sSDisp = Format$(pnSaldo, "#,##0.00")
    lsSaldoC = Format$(pnSaldoC, "#,##0.00")

    lsNomAge = sNomAge

    sCad = sCad & oImp.gPrnNegritaON 'Activa Negrita


    sCad = sCad & oImp.gPrnSaltoLinea & oImp.gPrnNegritaON 'Activa Negrita
    If pbImpTMU Then
        sCad = sCad & Space(sIni) & "CMAC-MAYNAS - AHORRO" & Space(16 + sSep + saux) & oImp.gPrnSaltoLinea
    Else
        sCad = sCad & Space(sIni) & "CMAC-MAYNAS - AHORRO" & Space(16 + sSep + saux) & "CMAC-MAYNAS - AHORRO" & oImp.gPrnSaltoLinea
    End If
    '*** PEAC 20090924
    If Mid(psCodCta, 3, 3) = "000" Then
        If Mid(psCodCta, 6, 1) = 1 Then
            '''sCad = sCad & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
            If pbImpTMU Then
                sCad = sCad & Space(sIni) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & Space(sSep) & oImp.gPrnSaltoLinea 'marg ers044-2016
            Else
                sCad = sCad & Space(sIni) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
            End If
        Else
            If pbImpTMU Then
                sCad = sCad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & oImp.gPrnSaltoLinea  'Trim(sNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & oImp.gPrnSaltoLinea
            Else
                sCad = sCad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & oImp.gPrnSaltoLinea 'Trim(sNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & oImp.gPrnSaltoLinea
            End If
        End If
    Else
        If Mid(psCodCta, 9, 1) = 1 Then
            If pbImpTMU Then
                sCad = sCad & Space(sIni) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & oImp.gPrnSaltoLinea  'marg ers044-2016
            Else
                sCad = sCad & Space(sIni) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
            End If
        Else
            If pbImpTMU Then
                sCad = sCad & Space(sIni) & Trim(sNomAge) & oImp.gPrnSaltoLinea 'Trim(sNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & oImp.gPrnSaltoLinea
            Else
                sCad = sCad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & oImp.gPrnSaltoLinea 'Trim(sNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & oImp.gPrnSaltoLinea
            End If
        End If
    End If
    '*** FIN PEAC
    
    If psNomAgeRem = "" Then
        'sCad = sCad & oImp.gPrnSaltoLinea
    Else
        If pbImpTMU Then
            sCad = sCad & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & oImp.gPrnSaltoLinea
        Else
            sCad = sCad & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & "Ag.Rem: " & Trim(psNomAgeRem) & oImp.gPrnSaltoLinea
        End If
    End If

    If psComCmac = "XXX" Then
        If psLin3 = "XXX" Then
            sCad = sCad & oImp.gPrnNegritaOFF 'Desactiva Negrita
        Else
            If pbImpTMU Then
                sCad = sCad & Space(sIni) & psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea    'Desactiva Negrita
            Else
                sCad = sCad & Space(sIni) & psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & oImp.gPrnNegritaOFF & psLin3 & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea   'Desactiva Negrita
            End If
            lnNumLinCmac = 1
        End If
        lnNumLinCmac = 0
    Else
        If pbImpTMU Then
            sCad = sCad & Space(sIni) & "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea   '"NroDocCmac:" & psComCmac & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea   'Desactiva Negrita
        Else
            sCad = sCad & Space(sIni) & "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea   '"NroDocCmac:" & psComCmac & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea   'Desactiva Negrita
        End If
        lnNumLinCmac = 1
    End If

    If pbImpTMU Then
        sCad = sCad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea
    Else
        sCad = sCad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea
    End If
    If psCodOpe = "200904" Or psCodOpe = "200905" Then
        sCad = sCad & Space(sIni) & Space(21) & "Nro. Mov.:" & CStr(pnMovNro) & Space(saux + sSep - 6) & Space(21) & "Nro. Mov.:" & CStr(pnMovNro) & oImp.gPrnSaltoLinea
    End If

    lnCliAux = InStr(1, psCliente, "*", vbTextCompare)

    If lnCliAux = 0 Then
        If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
        If pbImpTMU Then
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & oImp.gPrnSaltoLinea
        Else
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & oImpre.ImpreCarEsp(psCliente) & oImp.gPrnSaltoLinea
        End If
    Else
        lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
        lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
    
        If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
        If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
        If pbImpTMU Then
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & oImp.gPrnSaltoLinea
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & oImp.gPrnSaltoLinea
        Else
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & Space(saux + sMax + sSep - Len(lsCliAux1)) & oImpre.ImpreCarEsp(lsCliAux1) & oImp.gPrnSaltoLinea
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & Space(saux + sMax + sSep - Len(lsCliAux2)) & oImpre.ImpreCarEsp(lsCliAux2) & oImp.gPrnSaltoLinea
        End If
        lnCliAux = 1
    End If

    'INICIO ORCR20140714
    'If psCodOpe = 300528 Then
    If psCodOpe = "300528" Then 'JUEZ 20140910
        If pbImpTMU Then
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(psDNI) & Space(saux + sMax + sSep - Len(psDNI)) & oImpre.ImpreCarEsp(psDNI) & oImp.gPrnSaltoLinea
        Else
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(psDNI) & oImp.gPrnSaltoLinea
        End If
    End If
    'FIN ORCR20140714

    If pbSaldoInt Or pbCuenta Then
        If pbImpTMU Then
            sCad = sCad & Space(sIni) & "Cuenta:" & psCodCta & oImp.gPrnSaltoLinea
        Else
            sCad = sCad & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + saux) & "Cuenta:" & psCodCta & oImp.gPrnSaltoLinea
        End If
    Else
        'sCad = sCad & oImp.gPrnSaltoLinea
    End If

    psTit = Trim(psTit)
    psTit = oImpre.CentrarCadena(psTit, 28)
    sCad = sCad & oImp.gPrnNegritaON 'Activa Negrita
    If pbImpTMU Then
        sCad = sCad & Space(sIni + 1) & "-----" & psTit & "-----" & oImp.gPrnSaltoLinea
    Else
        sCad = sCad & Space(sIni + 1) & "-----" & psTit & "-----" & Space(-1 + sSep) & "-----" & psTit & "-----" & oImp.gPrnSaltoLinea
    End If

    lnChq = InStr(1, pstext, "*", vbTextCompare)

    If psTexto = "XXX" Then
        If lnChq = 0 Then
            'sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid(psText, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(Mid(psText, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)) & sMonto & oImp.gPrnSaltoLinea
            If pbImpTMU Then
                sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - IIf(Len(sMonto) > 11, 11, Len(sMonto))) & sMonto & oImp.gPrnSaltoLinea 'EJVG20140123
            Else
                sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - IIf(Len(sMonto) > 11, 11, Len(sMonto))) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - IIf(Len(sMonto) > 11, 11, Len(sMonto))) & sMonto & oImp.gPrnSaltoLinea 'EJVG20140123
            End If
            sCad = sCad & oImp.gPrnSaltoLinea
        Else
            lsChqAux1 = (Mid(pstext, 1, lnChq - 1))
            lsChqAux2 = (Mid(pstext, lnChq + 1))
            If pbImpTMU Then
                sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & oImp.gPrnSaltoLinea
                sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & oImp.gPrnSaltoLinea
            Else
                sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & oImp.gPrnSaltoLinea
                sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & oImp.gPrnSaltoLinea
            End If
        End If
    Else
        If pbImpTMU Then
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & oImp.gPrnSaltoLinea
        Else
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & Space(saux + sSep + sMax - Len(Trim(psTexto))) & oImpre.ImpreCarEsp(Trim(psTexto)) & oImp.gPrnSaltoLinea
        End If
        sCad = sCad & oImp.gPrnSaltoLinea
    End If

    sCad = sCad & oImp.gPrnNegritaOFF 'Desactiva Negrita

    If pbSaldoInt Then
        If bPreguntaSaldo Then
            'If MsgBox("Desea Imprimir el Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            If pbImprimeSaldos Then
                If pbImpTMU Then
                    sCad = sCad & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & oImp.gPrnSaltoLinea
                Else
                    sCad = sCad & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & oImp.gPrnSaltoLinea
                End If
                If pbOpSaldoC Then
                    If pbImpTMU Then
                        sCad = sCad & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & oImp.gPrnSaltoLinea
                    Else
                        sCad = sCad & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & oImp.gPrnSaltoLinea
                    End If
                Else
                    'sCad = sCad & oImp.gPrnSaltoLinea
                End If
            Else
                'sCad = sCad & oImp.gPrnSaltoLinea
                'sCad = sCad & oImp.gPrnSaltoLinea
                pbSaldoInt = False
            End If
        Else
            If pbImpTMU Then
                sCad = sCad & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & oImp.gPrnSaltoLinea
            Else
                sCad = sCad & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & oImp.gPrnSaltoLinea
            End If
            If pbOpSaldoC Then
                If pbImpTMU Then
                    sCad = sCad & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & oImp.gPrnSaltoLinea
                Else
                    sCad = sCad & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & oImp.gPrnSaltoLinea
                End If
            Else
                'sCad = sCad & oImp.gPrnSaltoLinea
            End If
        End If
    Else
        'sCad = sCad & oImp.gPrnSaltoLinea
    End If

    lsInteres = pnInteresA

    If pbSaldoInt Then
        If lsInteres <> "No Valido" Then
            lsInteres = Format(lsInteres, "#,##0.00")
            If pbImpTMU Then
                sCad = sCad & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & oImp.gPrnSaltoLinea
            Else
                sCad = sCad & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Space(-1 + sSep) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & oImp.gPrnSaltoLinea
            End If
        End If
    Else
        'sCad = sCad & oImp.gPrnSaltoLinea
    End If

    If Not psNumDoc = "" Then
        If pbImpTMU Then
            sCad = sCad & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & NomDoc & oImp.gPrnSaltoLinea
        Else
            sCad = sCad & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & oImp.gPrnSaltoLinea
        End If
    Else
        'sCad = sCad & oImp.gPrnSaltoLinea
    End If

    If Not psNumDias = "----" Then
        If pbImpTMU Then
            sCad = sCad & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & oImp.gPrnSaltoLinea
        Else
            sCad = sCad & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Space(-1 + sSep) & "Nro Dias Interes" & Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")) & psNumDias & oImp.gPrnSaltoLinea
        End If
        lnTope = 5 - lnCliAux
    Else
        lnTope = 4 - lnCliAux
    End If

    If pnITFValor <> 0 Then
        Dim nTotalP As Double
        Dim sTotalP As String
        'sTotalP = Format$((pnITFValor + pnMonto), "#,##0.00")
        sTotalP = Format$((pnMonto), "#,##0.00")
        lnNumLinCmac = lnNumLinCmac + 2
        If pbImpTMU Then
            sCad = sCad & Space(sIni) & "ITF : " & Space(29) & Format(Abs(pnITFValor), "#,##0.00") & oImp.gPrnSaltoLinea
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid("TOTAL A PAGAR", 1, 28))) & Space(sMax + 6 - Len(Trim(Mid("TOTAL A PAGAR", 1, 28))) - Len(sTotalP)) & sTotalP & oImp.gPrnSaltoLinea
        Else
            sCad = sCad & Space(sIni) & "ITF : " & Space(29) & Format(Abs(pnITFValor), "#,##0.00") & Space(14) & "ITF : " & Space(29) & Format(Abs(pnITFValor), "#,##0.00") & oImp.gPrnSaltoLinea
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid("TOTAL A PAGAR", 1, 28))) & Space(sMax + 6 - Len(Trim(Mid("TOTAL A PAGAR", 1, 28))) - Len(sTotalP)) & sTotalP & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(Mid("TOTAL A PAGAR", 1, 28))) & Space(sMax + 6 - Len(Trim(Mid("TOTAL A PAGAR", 1, 28))) - Len(sTotalP)) & sTotalP & oImp.gPrnSaltoLinea
        End If
    Else
        lnNumLinCmac = lnNumLinCmac + 1
    End If
    If pbImpTMU Then
        sCad = sCad & Space(sIni) & "---------------------------------------" & oImp.gPrnSaltoLinea
    Else
        sCad = sCad & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & "---------------------------------------" & oImp.gPrnSaltoLinea
    End If
         'VALIDACION PARA EL MOTIVO -------- ANPS2021407
    
    If psMot = 1 Then
            If pbImpTMU Then
                sCad = sCad & Space(sIni) & Right("Estimado Cliente, la comisión cobrada", 37) & oImp.gPrnSaltoLinea
                sCad = sCad & Space(sIni) & Right("estará sujeta a evaluación y en caso resulte", 37) & oImp.gPrnSaltoLinea
                sCad = sCad & Space(sIni) & Right("procedente, se realizará la devolución ", 37) & oImp.gPrnSaltoLinea
                sCad = sCad & Space(sIni) & Right("correspondiente.", 37) & oImp.gPrnSaltoLinea
            Else
                sCad = sCad & Space(sIni) & Right("Estimado Cliente, la comisión cobrada", 37) & Space(2 + sSep) & Right("Estimado Cliente, la comisión cobrada", 37) & oImp.gPrnSaltoLinea
                sCad = sCad & Space(sIni) & Right("estará sujeta a evaluación y en caso resulte", 37) & Space(2 + sSep) & Right("estará sujeta a evaluación y en caso resulte", 37) & oImp.gPrnSaltoLinea
                sCad = sCad & Space(sIni) & Right("procedente, se realizará la devolución ", 37) & Space(2 + sSep) & Right("procedente, se realizará la devolución", 37) & oImp.gPrnSaltoLinea
                sCad = sCad & Space(sIni) & Right("correspondiente.", 37) & Space(23 + sSep) & Right("correspondiente.", 37) & oImp.gPrnSaltoLinea
            End If
    End If
        
    'FIN -------- ANPS2021407
    If psCodUsuRem = "" Then
        If pbImpTMU Then
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & oImp.gPrnSaltoLinea
        Else
            sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Space(29 + sSep + saux) & oImpre.ImpreCarEsp(sCodUser) & oImp.gPrnSaltoLinea
        End If
    Else
        If pbImpTMU Then
            sCad = sCad & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & oImp.gPrnSaltoLinea
        Else
            sCad = sCad & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & Space(sSep) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & oImp.gPrnSaltoLinea
        End If
    End If
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Set clsGen = New COMDConstSistema.DCOMGeneral
    lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
    Set clsGen = Nothing
    'By Capi 10042008 para que corte despues de la impresion
    'sCad = sCad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(lsMensaje) & Space(-1 + sSep) & oImp.gPrnNegritaOFF & oImpre.ImpreCarEsp(lsMensaje) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
    If pbImpTMU Then
        sCad = sCad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(lsMensaje) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
    Else
        sCad = sCad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(lsMensaje) & Space(-1 + sSep) & oImp.gPrnNegritaOFF & oImpre.ImpreCarEsp(lsMensaje) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
    End If
    '
    
    'scad = scad & Chr$(12)
    Dim nmenos As Integer
    nmenos = 0

    If psCodOpe = CStr(gAhoTransCargo) Or psCodOpe = CStr(gAhoTransAbono) Then
        nmenos = 2
    End If

    'For saux = 1 To (lnTope - lnNumLinCmac) - nmenos
    '    sCad = sCad & oImp.gPrnSaltoLinea
    'Next saux

    If bCadena Then
        ImprimeBoleta = sCad
    Else
        ImprimeBoleta = sCad

    End If
    Exit Function
Error:
   
    Err.Raise Err.Number, "Error al realizar la Impresión...", Err.Description
End Function

Private Function LinRec(ByVal sTexto As String, pnSepara As Integer) As String
LinRec = sTexto & Space(pnSepara) & oImp.gPrnSaltoLinea ' sTexto & oImp.gPrnSaltoLinea
End Function


Public Function ImprimeBoletaDepCtaBanco(ByVal sNomCmac As String, ByVal sNomAge As String, _
        ByVal dFecha As Date, ByVal sBanco As String, ByVal sTipoCuenta As String, _
        ByVal sDocumento As String, ByVal nImporte As Double, ByVal sLpt As String, _
        ByVal nMoneda As COMDConstantes.Moneda, ByVal sUsuario As String, ByVal sOperacion As String, _
        Optional bCadena As Boolean = False) As String
    
Dim sImpre As String, sCad As String
Dim nFicSal As Integer
Dim gPrnNegritaON As String
Dim gPrnNegritaOFF As String
Dim sOpe As String
Dim nFila As Integer, nSep As Integer, nMax As Integer
Dim sTexto As String

Dim oImpre As New COMFunciones.FCOMImpresion

ETIQ:
    On Error GoTo ImpreErr
    

sCad = sCad & Chr$(27) & Chr$(64)

sCad = sCad & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
sCad = sCad & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
sCad = sCad & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
sCad = sCad & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
sCad = sCad & Chr$(27) + Chr$(18) ' cancela condensada
sCad = sCad & Chr$(27) + Chr$(72) ' desactiva negrita
    
nSep = 14
nMax = 39

sCad = sCad & oImp.gPrnSaltoLinea
sCad = sCad & oImp.gPrnSaltoLinea
sCad = sCad & oImp.gPrnSaltoLinea
sCad = sCad & LinRec(oImp.gPrnNegritaON & oImpre.Justifica(sNomCmac, nMax - 11) & " " & Left(dFecha, 10) & oImp.gPrnNegritaOFF, nSep)
sCad = sCad & LinRec(oImp.gPrnNegritaON & oImpre.Justifica(Trim(sNomAge), nMax - 12) & " " & Format$(Time, "hh:mm:ss AMPM") & oImp.gPrnNegritaOFF, nSep)
sCad = sCad & oImp.gPrnSaltoLinea
sCad = sCad & LinRec(String(nMax, "-"), nSep)
sCad = sCad & LinRec(oImp.gPrnNegritaON & oImpre.Justifica(oImp.gPrnNegritaON & "Entidad: " & sBanco & oImp.gPrnNegritaOFF, nMax + 2) & oImp.gPrnNegritaOFF, nSep)
sCad = sCad & LinRec(oImp.gPrnNegritaON & oImpre.Justifica(oImp.gPrnNegritaON & "Cuenta : " & sTipoCuenta & oImp.gPrnNegritaOFF, nMax + 2) & oImp.gPrnNegritaOFF, nSep)
sCad = sCad & LinRec(oImp.gPrnNegritaON & oImpre.Justifica(oImp.gPrnNegritaON & "Doc    : " & sDocumento & oImp.gPrnNegritaOFF, nMax + 2) & oImp.gPrnNegritaOFF, nSep)
sCad = sCad & LinRec(String(nMax, "-"), nSep)
sCad = sCad & oImp.gPrnSaltoLinea
'''sOpe = sOperacion & " (" & IIf(nMoneda = COMDConstantes.gMonedaNacional, "SOLES", "DOLARES") & ")" 'marg ers044-2016
sOpe = sOperacion & " (" & IIf(nMoneda = COMDConstantes.gMonedaNacional, StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), "DOLARES") & ")" 'marg ers044-2016

If Len(sOpe) <= nMax Then
   sCad = sCad & LinRec(oImp.gPrnNegritaON & oImpre.Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & oImp.gPrnNegritaOFF, nSep)
Else
   '''sOpe = Left(sOperacion, nMax - 10) & " (" & IIf(nMoneda = gMonedaNacional, "SOLES", "DOLARES") & ")" 'marg ers044-2016
   sOpe = Left(sOperacion, nMax - 10) & " (" & IIf(nMoneda = gMonedaNacional, StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), "DOLARES") & ")" 'marg ers044-2016
   sCad = sCad & LinRec(oImp.gPrnNegritaON & oImpre.Justifica(String(Int((nMax - Len(sOpe)) / 2), "-") & sOpe & String(Int((nMax - Len(sOpe)) / 2), "-"), nMax) & oImp.gPrnNegritaOFF, nSep)
End If
sCad = sCad & oImp.gPrnSaltoLinea
sCad = sCad & LinRec(oImp.gPrnNegritaON & oImpre.Justifica("MONTO", nMax - 15) & " " & Right(Space(14) & Format(nImporte, "#,##0.00"), 14) & oImp.gPrnNegritaOFF, nSep)
sCad = sCad & LinRec(String(nMax, "="), nSep)
sCad = sCad & LinRec(oImpre.Justifica(sUsuario, nMax), nSep)

sCad = oImpre.ImpreCarEsp(sCad)
sCad = sCad & oImp.gPrnSaltoLinea
sCad = sCad & oImp.gPrnSaltoLinea
sCad = sCad & oImp.gPrnSaltoLinea
sCad = sCad & oImp.gPrnSaltoLinea
sCad = sCad & oImp.gPrnSaltoLinea
    ImprimeBoletaDepCtaBanco = sCad
Exit Function
ImpreErr:
  
    Err.Raise Err.Number, "Comprueba la conexion de su impresora...", Err.Description
End Function

Public Function ImprimirParametros(ByVal rs As ADODB.Recordset, ByVal gsNomAge As String, ByVal gdFecSis As Date) As String
    Dim sCad As String
    Dim nCarLin As Integer, i As Integer
    Dim sTit1 As String, sTit2 As String
    Dim sNumPag As String
    Dim nLinPag As Integer, nCntPag As Integer
    Dim sDesc As String * 55
    Dim sVal As String * 12
    sNumPag = "01"
    nLinPag = 65
    nCarLin = 85
    sTit1 = "P A R A M E T R O S   D E   C A P T A C I O N E S"
    sTit2 = ""
    nCntPag = 0
    sCad = sCad & CabeRepoCaptac("", "", nCarLin, "CAPTACIONES", sTit1, sTit2, "", sNumPag, gsNomAge, gdFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
    sCad = sCad & String(nCarLin, "=") & oImp.gPrnSaltoLinea
    sCad = sCad & "DESCRIPCION" & Space(55) & "VALOR" & oImp.gPrnSaltoLinea
    sCad = sCad & String(nCarLin, "=") & oImp.gPrnSaltoLinea
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            sDesc = rs!cDescripcion
            RSet sVal = rs!cValor
            sCad = sCad & sDesc & Space(2) & sVal & oImp.gPrnSaltoLinea
            rs.MoveNext
        Loop
    End If
    ImprimirParametros = sCad
    rs.Close
End Function

Public Function ImprimirTasaInt(ByVal rs As ADODB.Recordset, ByVal sTitProd As String, ByVal gMonedaNacional As String, _
                                ByVal gsNomAge As String, ByVal gdFecSis As Date, ByVal gsNomCmac As String, ByVal pnMoneda As Integer) As String
    Dim sMontoIni As String * 15, sMontoFin As String * 15
    Dim sPlazoIni As String * 8, sPlazoFin As String * 8
    Dim sTasa As String * 8
    Dim sOrdPag As String * 2
    Dim nCntPag As Integer, nCarLin As Integer
    Dim sCad As String, sTit1 As String, sTit2 As String
    Dim i As Long
    Dim sBoldON As String, sBoldOFF As String
    Dim sMoneda As String, sNumPag As String * 4
    Dim ofun As COMFunciones.FCOMImpresion
    sCad = ""
    nCarLin = 100
    sTit1 = UCase(sTitProd)
    RSet sNumPag = Trim(nCntPag)
    Set ofun = New COMFunciones.FCOMImpresion
    'By Capi 20012008
    Dim bVigente As Boolean
    bVigente = True
    If MsgBox("Desea Solo Tarifario Vigente ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        bVigente = False
    End If
    If bVigente Then
        sTit1 = "TARIFARIO VIGENTE - " & sTitProd & " - " & IIf(pnMoneda = gMonedaNacional, "MONEDA NACIONAL", "MONEDA EXTRANJERA")
    Else
        sTit1 = "TARIFARIO HISTORICO - " & sTitProd & " - " & IIf(pnMoneda = gMonedaNacional, "MONEDA NACIONAL", "MONEDA EXTRANJERA")
    End If
    sTit2 = "(EN TASAS NOMINALES ANUALES)"
    
    '''sMoneda = IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") 'marg ers044-2016
    sMoneda = IIf(pnMoneda = gMonedaNacional, StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), "DOLARES") 'marg ers044-2016
    sCad = sCad & CabeRepoCaptac(gsNomCmac, "", nCarLin, "CAPTACIONES", sTit1, sTit2, sMoneda, sNumPag, gsNomAge, gdFecSis, oImp.gPrnSaltoLinea)
    sCad = sCad & oImp.gPrnSaltoLinea
    sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & ofun.ImpreFormat("MONTO INICIAL", 12, 2) & ofun.ImpreFormat("MONTO FINAL", 12, 2) & ofun.ImpreFormat("PLAZO INICIAL", 12, 2) & ofun.ImpreFormat("PLAZO FINAL", 12, 2) & ofun.ImpreFormat("ORD. PAGO", 12, 2) & ofun.ImpreFormat("TASA INT", 12, 2) & ofun.ImpreFormat("ACTIVA", 12, 2) & ofun.ImpreFormat("CAMBIO", 12, 2) & oImp.gPrnSaltoLinea
    sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            'By Capi 20012008
            If bVigente And rs!sActiva = "SI" Then
                sCad = sCad & ofun.ImpreFormat(CDbl(rs!nMontoI), 6, 2)
                sCad = sCad & ofun.ImpreFormat(CDbl(rs!nMontoF), 14, 2)
                sCad = sCad & ofun.ImpreFormat(CDbl(rs!nPlazoI), 6, 4)
                sCad = sCad & ofun.ImpreFormat(CDbl(rs!nPlazoF), 10, 4)
                sCad = sCad & ofun.ImpreFormat(rs!sOrdPago, 8, 8)
                sCad = sCad & ofun.ImpreFormat(rs!sTasaI, 6, 4)
                sCad = sCad & ofun.ImpreFormat(rs!sActiva, 8, 8)
                sCad = sCad & ofun.ImpreFormat(rs!LogCambio, 28, 8)
                sCad = sCad & oImp.gPrnSaltoLinea
            End If
            If Not bVigente Then
                sCad = sCad & ofun.ImpreFormat(CDbl(rs!nMontoI), 6, 2)
                sCad = sCad & ofun.ImpreFormat(CDbl(rs!nMontoF), 14, 2)
                sCad = sCad & ofun.ImpreFormat(CDbl(rs!nPlazoI), 6, 4)
                sCad = sCad & ofun.ImpreFormat(CDbl(rs!nPlazoF), 10, 4)
                sCad = sCad & ofun.ImpreFormat(rs!sOrdPago, 8, 8)
                sCad = sCad & ofun.ImpreFormat(rs!sTasaI, 6, 4)
                sCad = sCad & ofun.ImpreFormat(rs!sActiva, 8, 8)
                sCad = sCad & ofun.ImpreFormat(rs!LogCambio, 28, 8)
                sCad = sCad & oImp.gPrnSaltoLinea
            End If
            rs.MoveNext
        Loop
    End If
    sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
    rs.Close
    ImprimirTasaInt = sCad
End Function

Public Function ImprimeCertificadoPlazoFijo(ByVal dApertura As Date, ByVal sNomCli As String, _
        ByVal sDirCli As String, ByVal sCuenta As String, nExtracto As String, _
        ByVal nPlazo As Long, ByVal nSaldo As Double, ByVal nTasa As Double, _
        ByVal sRetInt As String, ByVal sLetras As String, Optional nDuplicado As Integer = 0, Optional ByVal bCopia As Boolean = False, _
        Optional ByVal sTipoCta As String, Optional ByVal RsR As ADODB.Recordset, Optional ByVal psNomAge As String, Optional psFormaRetiro As String, Optional nCostoMan As Currency = 0) As String
        
Dim intLinBla As Integer, nFicSal As Integer
Dim sFchLet As String, sFchVct As String
Dim sFchCan As String, sForMon As String
Dim sVctLet As String, dVencimiento As Date, dCancelacion As Date
Dim sTaInNo As String, sTaInEf As String
Dim sTaInAn As String, sPlazo As String
Dim i As Integer
Dim sCad As String

Dim ofunf  As New COMFunciones.FCOMFechas
Dim oFunI As New COMFunciones.FCOMImpresion
Dim oFunN As New COMFunciones.FCOMCadenas
Dim scadTitu As String
Dim NroTitu As Integer
Dim sDocumento As String
Dim sql As String

'Fechas
sFchLet = ofunf.ArmaFecha(dApertura)
dVencimiento = DateAdd("d", nPlazo, dApertura)
dCancelacion = DateAdd("d", nPlazo + 1, dApertura)

sVctLet = ofunf.ArmaFecha(dVencimiento)
sFchCan = ofunf.ArmaFecha(dCancelacion)

'*** PEAC 20090727
'***Modificado por ELRO el 20110915, según Acta 260-2011/TI-D
'sForMon = IIf(Right(sLetras, 5) = "SOLES", "S/. ", IIf(Right(sLetras, 5) = "DOLAR", "US$ ", "Eu. "))
'''sForMon = IIf(Right(sLetras, 5) = "SOLES", "S/. ", IIf(Right(sLetras, 7) = "DOLARES", "US$ ", "Eu. ")) 'marg ers044-2016
sForMon = IIf(Right(sLetras, 5) = StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), gcPEN.gcPEN_SIMBOLO & " ", IIf(Right(sLetras, 7) = "DOLARES", "US$ ", "Eu. ")) 'marg ers044-2016
'***Fin Modificado por ELRO**********************************
sTaInNo = Format$((nTasa / 12), "#0.00")
'sTaInEf = Format$(((((((nTasa / 12) / 3000) + 1) ^ nPlazo) - 1) * 100), "#0.00")
'sTaInAn = Format(((((((nTasa / 12) / 3000) + 1) ^ 360) - 1) * 100), "#0.00")
sTaInEf = Format$(((1 + nTasa / 36000) ^ 360 - 1) * 100, "#0.00")
'ALPA 20091118************************************************
Dim objCaptac As COMNCaptaGenerales.NCOMCaptaGenerales
Set objCaptac = New COMNCaptaGenerales.NCOMCaptaGenerales
Dim pnPlazoT As Integer
Dim nTREA As Currency
pnPlazoT = CInt(nPlazo)
nTREA = objCaptac.ObtenerTREA(Mid$(sCuenta, 6, 3), nSaldo, sTaInEf, pnPlazoT, nCostoMan)
    '*************************************************************

'Documento

If Not (RsR.EOF And RsR.BOF) Then
    RsR.MoveFirst
    Dim oPers As COMDPersona.DCOMPersonas
    Dim rs As New ADODB.Recordset
    Dim sNomPers As String
    Dim sDireccion As String
    Dim OBJF As COMFunciones.FCOMFechas
    Set OBJF = New COMFunciones.FCOMFechas
    
    sPlazo = Trim(nPlazo) & " dia(s)"
    sCad = ""
    'sCad = sCad & Chr$(15)'Retorna al tipo de letra normal
    sCad = sCad & oImp.gPrnBoldON
    sCad = sCad & oImp.gPrnTamLetra15CPI
    sCad = sCad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    sCad = sCad & oImp.gPrnBoldON
    'sCad = sCad & Chr$(27) + Chr$(107) + Chr$(1)
    sCad = sCad & psNomAge & Space(11) & dApertura & " " & Time & oImp.gPrnSaltoLinea
    sCad = sCad & oImp.gPrnBoldOFF
    'sCad = sCad & Chr$(27) + Chr$(107) + Chr$(1)     'Tipo de Letra Roman
    sCad = sCad & oImp.gPrnTamPagina22
   ' sCad = sCad & Chr$(27) + Chr$(50)                'Espaciado entre lineas 1/16
    
    sCad = sCad & "" & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    
    If Not (RsR.EOF And RsR.BOF) Then
       Do Until RsR.EOF
        If Right(RsR("Relacion"), 2) = 10 Or Right(RsR("Relacion"), 2) = 12 Then
                Set oPers = New COMDPersona.DCOMPersonas
                    Set rs = oPers.BuscaCliente(RsR(0), 2)
                Set oPers = Nothing
                If Not (rs.EOF And rs.BOF) Then
                    NroTitu = NroTitu + 1
                    
                    sDocumento = IIf(Trim(IsNull(rs!cPersIDnroDNI)), "R.U.C " & Trim(rs!cPersIDnroRUC), "D.N.I " & Trim(rs!cPersIDnroDNI))
                    sNomPers = oFunN.PstaNombre(rs!cPersNombre)
                    scadTitu = scadTitu & oFunI.ImpreFormat(oFunI.ImpreCarEsp(Trim(sNomPers)), 60) & Space(5) & sDocumento & oImp.gPrnSaltoLinea
                    sDireccion = IIf(IsNull(rs(3)) = True, "", rs(3))
                    sCad = sCad & oImp.gPrnBoldON
                    sCad = sCad & oFunI.ImpreFormat(sNomPers, 30) & Space(5) & oFunI.ImpreFormat(sDocumento, 15) & oImp.gPrnSaltoLinea
                    
                Else
                    sDocumento = ""
                End If
            End If
            RsR.MoveNext
        Loop
    End If
End If

sCad = sCad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
sCad = sCad & " DEPOSITO            : " & sForMon & " " & nSaldo & oImp.gPrnSaltoLinea
sCad = sCad & oImp.gPrnBoldOFF
sCad = sCad & "                       ( " & sLetras & ")" & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea

sCad = sCad & " PLAZO               : " & nPlazo & " DIAS " & oImp.gPrnSaltoLinea
sCad = sCad & " NRO PLAZO FIJO      : " & sCuenta & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea

sCad = sCad & oImp.gPrnBoldON
sCad = sCad & " TIPO DE CUENTA      : " & sTipoCta & oImp.gPrnSaltoLinea
sCad = sCad & oImp.gPrnBoldOFF
sCad = sCad & " FECHA VENCIMIENTO   : " & sVctLet & oImp.gPrnSaltoLinea
'sCad = sCad & " FECHA CANCELACION   : " & sFchCan & oImp.gPrnSaltoLinea
'ALPA 200091118
'sCad = sCad & " T.N.A               : " & nTasa & oImp.gPrnSaltoLinea
sCad = sCad & " T.R.E.A             : " & Format$(nTREA, "#0.00") & oImp.gPrnSaltoLinea
sCad = sCad & " T.E.A               : " & Format$(sTaInEf, "#0.00") & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'************************
sCad = sCad & " RET. DE INTERES     : " & psFormaRetiro & oImp.gPrnSaltoLinea & oImp.gPrnSaltoPagina

'sFchLet = ofunf.ArmaFecha(dApertura)
'dVencimiento = DateAdd("d", nPlazo, dApertura)
'dCancelacion = DateAdd("d", nPlazo + 1, dApertura)

'sVctLet = ofunf.ArmaFecha(dVencimiento)
'sFchCan = ofunf.ArmaFecha(dCancelacion)
'sForMon = IIf(Right(sLetras, 5) = "SOLES", "S/. ", "US$ ")

'sTaInNo = Format$((nTasa / 12), "#0.00")
'sTaInEf = Format$(((((((nTasa / 12) / 3000) + 1) ^ nPlazo) - 1) * 100), "#0.00")
'sTaInAn = Format(((((((nTasa / 12) / 3000) + 1) ^ 360) - 1) * 100), "#0.00")


'*****agregado--

'If bCopia = True Then
'    sCad = sCad & Chr$(27) & Chr$(67) & Chr$(17)  'longitud de 17 lineas
'Else
'    sCad = sCad & Chr$(27) & Chr$(67) & Chr$(16)  'Longitud de página a 24 líneas
'End If
'sCad = sCad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'sCad = sCad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'sCad = sCad & ofuni.ImpreCarEsp(sNomPers) & Space(20) & sDocumento & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea

'If NroTitu = 3 Then
'    sCad = sCad & scadTitu & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'ElseIf NroTitu = 2 Then
'    sCad = sCad & scadTitu & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'Else
'    sCad = sCad & scadTitu & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'End If

'sCad = sCad & "*** " & sLetras & " ***" & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'sCad = sCad & sCuenta & Space(40) & sForMon & oFunN.JDNum(Trim(nSaldo), 12, True, 9, 2) & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'sCad = sCad & sDireccion & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'sCad = sCad & Space(10) & dVencimiento & Space(40) & sTaInAn & "%" & Space(10) & sTaInEf & "%" & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'sCad = sCad & Space(10) & "SI" & Space(58) & "DIARIA" & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'sCad = sCad & Space(10) & Trim(nPlazo) & " dias" & Space(52) & Trim(Left(psFormaRetiro, 20)) & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'
'sCad = sCad & "LOS INTERESES SE ABONARAN      :   " & Left(sRetInt, 20) & oImp.gPrnSaltoLinea

'sCad = sCad & oImp.gPrnSaltoLinea  'Avance de Página

ImprimeCertificadoPlazoFijo = sCad

End Function

Public Function fgImprimirConsultaSaldos(ByVal rs As ADODB.Recordset, ByVal pnCarLin As Integer, ByVal psTitRp1 As String, ByVal psTitRp2 As String, _
                                         ByVal psMoneda As String, ByVal psNumPag As String, ByVal gsNomAge As String, ByVal gdFecSis As String, _
                                         ByVal psCuenta As String, ByVal psapertura As String, ByVal psSaldoDisponible As String, ByVal psSaldoContable As String, _
                                         ByVal PsSaldoRetiro As String, ByVal psMon As String, ByVal psInteresMes As String, ByVal psEstado As String, _
                                         ByVal gsCodUser As String, ByVal pnProducto As Integer, ByVal psCapCTS As String, Optional ByVal pbImpTMU As Boolean = False) As String
    Dim sCad As String
    Dim sCliente As String
    Dim sRelacionCta As String
    
    If pbImpTMU = False Then
        sCad = sCad & oImp.gPrnSaltoLinea
        sCad = sCad & CabeRepoCaptac("", "", pnCarLin, "SECCION AHORROS", psTitRp1, psTitRp2, psMoneda, psNumPag, gsNomAge, gdFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
        sCad = sCad & Space(30) & "CUENTA         : " & psCuenta & oImp.gPrnSaltoLinea
        sCad = sCad & Space(30) & "FECHA APERTURA : " & psapertura & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            Do Until rs.EOF
             sCliente = rs(0)
             sRelacionCta = rs(1)
             sCad = sCad & Space(30) & Space(3) & sRelacionCta & "  " & sCliente & oImp.gPrnSaltoLinea
             rs.MoveNext
            Loop
         End If
        
        
        sCad = sCad & oImp.gPrnSaltoLinea
        sCad = sCad & Space(30) & String(pnCarLin, "-") & oImp.gPrnSaltoLinea           ' hasta 160 caracteres x línea
        
        sCad = sCad & Space(30) & "SALDO DISPONIBLE       " & Space(30) & psMon & " " & psSaldoDisponible & oImp.gPrnSaltoLinea
        sCad = sCad & Space(30) & "SALDO CONTABLE         " & Space(30) & psMon & " " & psSaldoContable & oImp.gPrnSaltoLinea
        If pnProducto = psCapCTS Then
            sCad = sCad & Space(30) & "SALDO RETIRO           " & Space(30) & psMon & " " & PsSaldoRetiro & oImp.gPrnSaltoLinea
        End If
        sCad = sCad & Space(30) & "INTERÉSES AL " & Format$(gdFecSis, "dd/mm/yyyy") & Space(30) & psMon & " " & psInteresMes & oImp.gPrnSaltoLinea
        
        sCad = sCad & Space(30) & String(pnCarLin, "=") & oImp.gPrnSaltoLinea              ' hasta 160 caracteres x línea
        sCad = sCad & Space(30) & psEstado & oImp.gPrnSaltoLinea   '   & Chr$(12)
        sCad = sCad & Space(30) & String(pnCarLin, "-") & oImp.gPrnSaltoLinea
        sCad = sCad & Space(30) & gsCodUser
    Else
        'sCad = sCad & oImp.gPrnSaltoLinea
        sCad = sCad & "SECCION AHORROS" & oImp.gPrnSaltoLinea
        sCad = sCad & psTitRp1 & Space(10) & psTitRp2 & oImp.gPrnSaltoLinea
        sCad = sCad & "Moneda:" & psMoneda & oImp.gPrnSaltoLinea
        sCad = sCad & gsNomAge & Space(10) & gdFecSis & oImp.gPrnSaltoLinea
        sCad = sCad & "CUENTA         : " & psCuenta & oImp.gPrnSaltoLinea
        sCad = sCad & "FECHA APERTURA : " & psapertura & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            Do Until rs.EOF
             sCliente = rs(0)
             sRelacionCta = rs(1)
             sCad = sCad & Space(2) & sRelacionCta & "  " & sCliente & oImp.gPrnSaltoLinea
             rs.MoveNext
            Loop
         End If
        
        sCad = sCad & oImp.gPrnSaltoLinea
        sCad = sCad & String(38, "-") & oImp.gPrnSaltoLinea
        
        sCad = sCad & "SALDO DISPONIBLE  " & Space(2) & Trim(psMon) & " " & psSaldoDisponible & oImp.gPrnSaltoLinea
        sCad = sCad & "SALDO CONTABLE    " & Space(2) & Trim(psMon) & " " & psSaldoContable & oImp.gPrnSaltoLinea
        If pnProducto = psCapCTS Then
            sCad = sCad & Space(2) & "SALDO RETIRO      " & Space(2) & psMon & " " & PsSaldoRetiro & oImp.gPrnSaltoLinea
        End If
        sCad = sCad & "INTERES AL " & Format$(gdFecSis, "dd/mm/yyyy") & Space(2) & Trim(psMon) & psInteresMes & oImp.gPrnSaltoLinea
        sCad = sCad & String(38, "=") & oImp.gPrnSaltoLinea
        sCad = sCad & Space(2) & psEstado & oImp.gPrnSaltoLinea
        sCad = sCad & String(38, "-") & oImp.gPrnSaltoLinea
        sCad = sCad & Space(2) & gsCodUser
        sCad = sCad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
    End If
    fgImprimirConsultaSaldos = sCad
End Function

Public Function ImprimeConsultaMovimiento(ByVal rs As ADODB.Recordset, ByVal rs1 As ADODB.Recordset, ByVal gsCodUser As String, _
        ByVal gsNomCmac As String, ByVal gsNomAge As String, ByVal sMoneda As String, ByVal gdFecSis As String, _
        ByVal sTitRp1 As String, ByVal sTitRp2 As String, ByVal psCuenta As String, ByVal psFecSaldIni As String, _
        ByVal pdSaldoInicial As String, ByVal psSaldoDisponible As String, ByVal psSaldoContable As String, ByVal psSaldoRet As String, _
        ByVal psInteresMes As String, ByVal sIntPag As String, ByVal psEstado As String, ByVal psTotAbono As String, ByVal psTotCargo As String, _
        ByRef pnNroPag As Integer, Optional ByVal psTEA As String) As String
     
      
Dim sCad As String
Dim sCadR As String '--=========================EAAS20180511 SEGUN ERS TI-006-2018
Dim sCadAux As String
Dim nCarLin As Integer
Dim nLinPag As Integer, nEspaciosIzq As Integer
Dim sNumPag As String
Dim nCntPag As Integer
Dim sOperacion As String * 28
Dim sDocumento As String * 15
Dim sDeposito As String * 13
Dim sRetiro As String * 13
Dim sSaldoCnt As String * 13
Dim sAgencia  As String '* 10
Dim sFecha As String * 19
Dim sUser As String * 4
Dim ofunc As New COMFunciones.FCOMCadenas
Dim oFunI As New COMFunciones.FCOMImpresion


nCarLin = 140 'EAAS20180511 SEGUN ERS TI-006-2018 152
nLinPag = 65
sNumPag = ""
nCntPag = 0
nEspaciosIzq = 2


'By capi 21112008 ordenando el recordset
'Rs.Sort = "dFecha,sOperacion" 'Comentado por JUEZ 20131128
'rs.Sort = "dFecha,sOpeCod" 'JUEZ 20140128 'EAAS20180511 SEGUN ERS TI-006-2018

If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    
    'Do Until rs.EOF
        
        'If nLinPag > 58 Then
            
                
            nCntPag = nCntPag + 1
            sNumPag = ofunc.FillNum(Trim(nCntPag), 4, " ")
            sCad = sCad & oFunI.CabeRepo(gsNomCmac, gsNomAge, "SECCION AHORROS", sMoneda, Format$(gdFecSis, "dd/mm/yyyy"), sTitRp1, sTitRp2, "", "", nCntPag, nCarLin, nEspaciosIzq, oImp) & oImp.gPrnSaltoLinea
            sCad = sCad & Space(nEspaciosIzq) & "Cuenta No. :   " & oImp.gPrnNegritaON & psCuenta & oImp.gPrnNegritaOFF & Space(5) & "TEA :   " & oImp.gPrnNegritaON & psTEA & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            sCad = sCad & Space(nEspaciosIzq) & "CLIENTE" & Space(50) & "RELACION" & Space(20) & "DIRECCION" & oImp.gPrnSaltoLinea
            nLinPag = 7
            If Not (rs1.EOF And rs1.BOF) Then
                rs1.MoveFirst
                Do Until rs1.EOF
                Dim sCCI As String '--========= EAAS20180511 SEGUN ERS TI-006-2018
                
                    nLinPag = nLinPag + 1
                    sCad = sCad & Space(nEspaciosIzq) & Left(Trim(rs1(0)), 59) & Space(60 - Len(Left(Trim(rs1(0)), 50))) & Trim(rs1(1)) & Space(10) & Trim(rs1(2)) & oImp.gPrnSaltoLinea
                    sCCI = rs1!sCCI '--========= EAAS20180511 SEGUN ERS TI-006-2018
                    rs1.MoveNext
                Loop
            End If
            
            sCad = sCad & oImp.gPrnSaltoLinea
            sCad = sCad & Space(nEspaciosIzq) & "CCI :" & sCCI & oImp.gPrnSaltoLinea   '--=========INICIO EAAS20180511 SEGUN ERS TI-006-2018
            sCad = sCad & Space(nEspaciosIzq) & "SALDO AL " & psFecSaldIni & " : " & ofunc.JDNum(pdSaldoInicial, 12, True, 7, 2) & oImp.gPrnSaltoLinea
            sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea
            sCad = sCad & Space(nEspaciosIzq) & Space(103 - nEspaciosIzq) & "SALDO" & oImp.gPrnSaltoLinea
            sCad = sCad & Space(nEspaciosIzq) & "      FECHAS       "
            sCad = sCad & " " & "     TIPO DE OPERACION      "
            sCad = sCad & " " & "   DOCUMENTO   "
            sCad = sCad & " " & "    DEPOSITOS"
            sCad = sCad & " " & "      RETIROS"
            sCad = sCad & " " & "     CONTABLE   "
            sCad = sCad & " " & "AGENCIA/ESTABLECIMIENTO  "
            sCad = sCad & " " & "USUA" & oImp.gPrnSaltoLinea
            sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea            ' hasta 160 caracteres x línea
            nLinPag = nLinPag + 6
            If (nCntPag Mod 3) = 0 Then
                sCadAux = sCadAux & sCad
                sCad = ""
            End If
        'End If
        Do Until rs.EOF
        If nLinPag > 58 Then
            nCntPag = nCntPag + 1
            sNumPag = ofunc.FillNum(Trim(nCntPag), 4, " ")
            If sNumPag <> "" Then
                sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea
                sCad = sCad & Space(nEspaciosIzq) & gsCodUser
                sCad = sCad & oImp.gPrnSaltoPagina
                nLinPag = 0
            End If
        End If
        
        sCad = sCad & Space(nEspaciosIzq)
        sFecha = Format$(rs(0), "dd/mm/yyyy hh:mm:ss")
        sOperacion = oFunI.ImpreCarEsp(rs(1))
        sDocumento = Trim(rs(2))
        RSet sDeposito = rs(3)
        RSet sRetiro = rs(4)
        RSet sSaldoCnt = rs(5)
        sAgencia = Space(3) & Trim(rs(6)) '==============EAAS20180511 SEGUN ERS TI-006-2018
        sUser = Trim(rs(7))
        sCad = sCad & sFecha
        sCad = sCad & " " & sOperacion
        sCad = sCad & " " & sDocumento
        sCad = sCad & " " & sDeposito
        sCad = sCad & " " & sRetiro
        If Trim(sSaldoCnt) = "999999999" Then
            sCad = sCad & " " & Space(13)
        Else
            sCad = sCad & " " & sSaldoCnt
        End If
        '--=========================================
        '--INICIO EAAS20180511 SEGUN ERS TI-006-2018
        '--=========================================
        Dim nTam As Integer
        nTam = 27 - Len(sAgencia)
        sCad = sCad & " " & sAgencia & " " & Space(nTam)
        '--=========================================
        '--FIN EAAS20180511 SEGUN ERS TI-006-2018
        '--=========================================
        sCad = sCad & " " & sUser & oImp.gPrnSaltoLinea
        nLinPag = nLinPag + 1
        rs.MoveNext
        
    Loop
End If

sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea              ' hasta 160 caracteres x línea
sCad = sCad & Space(nEspaciosIzq) & "  TOTAL    " & Space(54) & psTotAbono & " " & psTotCargo & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea

''by capi 11112008 ****************************************
'
''By Capi 11112008 para operaciones pendientes cajeros automaticos
'sCad = sCad & oImp.gPrnSaltoLinea
'sCad = sCad & "  MOVIMIENTOS PENDIENTES" & oImp.gPrnSaltoLinea
'sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea
'nLinPag = nLinPag + 20
'
'If Not (Rs.EOF And Rs.BOF) Then
'
'    Do Until Rs.EOF
'        If nLinPag > 58 Then
'            If sNumPag <> "" Then
'                sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea
'                sCad = sCad & Space(nEspaciosIzq) & gsCodUser
'                sCad = sCad & oImp.gPrnSaltoPagina
'                nLinPag = 0
'            End If
'
'            nCntPag = nCntPag + 1
'            sNumPag = ofunc.FillNum(Trim(nCntPag), 4, " ")
'            sCad = sCad & ofuni.CabeRepo(gsNomCmac, gsNomAge, "SECCION AHORROS", sMoneda, Format$(gdFecSis, "dd/mm/yyyy"), sTitRp1, sTitRp2, "", "", nCntPag, nCarLin, nEspaciosIzq, oImp) & oImp.gPrnSaltoLinea
'            sCad = sCad & Space(nEspaciosIzq) & "Cuenta No. :   " & oImp.gPrnNegritaON & psCuenta & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
'            sCad = sCad & Space(nEspaciosIzq) & "CLIENTE" & Space(50) & "RELACION" & Space(20) & "DIRECCION" & oImp.gPrnSaltoLinea
'            nLinPag = 7
'            If Not (rs1.EOF And rs1.BOF) Then
'                rs1.MoveFirst
'                Do Until rs1.EOF
'                    nLinPag = nLinPag + 1
'                    sCad = sCad & Space(nEspaciosIzq) & Left(Trim(rs1(0)), 59) & Space(60 - Len(Left(Trim(rs1(0)), 50))) & Trim(rs1(1)) & Space(10) & Trim(rs1(2)) & oImp.gPrnSaltoLinea
'                    rs1.MoveNext
'                Loop
'            End If
'
'            sCad = sCad & oImp.gPrnSaltoLinea
'            sCad = sCad & Space(nEspaciosIzq) & "SALDO AL " & psFecSaldIni & " : " & ofunc.JDNum(pdSaldoInicial, 12, True, 7, 2) & oImp.gPrnSaltoLinea
'            sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea
'            sCad = sCad & Space(nEspaciosIzq) & Space(103 - nEspaciosIzq) & "SALDO" & oImp.gPrnSaltoLinea
'            sCad = sCad & Space(nEspaciosIzq) & "      FECHAS       "
'            sCad = sCad & " " & "     TIPO DE OPERACION      "
'            sCad = sCad & " " & "   DOCUMENTO   "
'            sCad = sCad & " " & "    DEPOSITOS"
'            sCad = sCad & " " & "      RETIROS"
'            sCad = sCad & " " & "     CONTABLE"
'            sCad = sCad & " " & " AGENCIA  " '& oImp.gPrnSaltoLinea
'            sCad = sCad & " " & "USUA" & oImp.gPrnSaltoLinea
'            sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea            ' hasta 160 caracteres x línea
'            nLinPag = nLinPag + 6
'            If (nCntPag Mod 3) = 0 Then
'                sCadAux = sCadAux & sCad
'                sCad = ""
'            End If
'        End If
'
'        sCad = sCad & Space(nEspaciosIzq)
'        sFecha = Format$(Rs(0), "dd/mm/yyyy hh:mm:ss")
'        sOperacion = ofuni.ImpreCarEsp(Rs(1))
'        sDocumento = Trim(Rs(2))
'        RSet sDeposito = Rs(3)
'        RSet sRetiro = Rs(4)
'        RSet sSaldoCnt = Rs(5)
'        sAgencia = Trim(Rs(6))
'        sUser = Trim(Rs(7))
'        sCad = sCad & sFecha
'        sCad = sCad & " " & sOperacion
'        sCad = sCad & " " & sDocumento
'        sCad = sCad & " " & sDeposito
'        sCad = sCad & " " & sRetiro
'        sCad = sCad & " " & Space(13) 'sSaldoCnt
'        sCad = sCad & " " & sAgencia
'        sCad = sCad & " " & sUser & oImp.gPrnSaltoLinea
'
'        nLinPag = nLinPag + 1
'        Rs.MoveNext
'    Loop
'End If
'
'sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea              ' hasta 160 caracteres x línea
'sCad = sCad & Space(nEspaciosIzq) & "  TOTAL    " & Space(54) & psTotAbono & " " & psTotCargo & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'
'
'
''**********************************************************


sCad = sCad & Space(nEspaciosIzq) & "SALDO DISPONIBLE  " & Space(75) & psSaldoDisponible & oImp.gPrnSaltoLinea
sCad = sCad & Space(nEspaciosIzq) & "SALDO CONTABLE    " & Space(75) & psSaldoContable & oImp.gPrnSaltoLinea

If Mid(psCuenta, 6, 3) = gCapCTS Then
    sCad = sCad & Space(nEspaciosIzq) & "SALDO RETIRO      " & Space(75) & psSaldoRet & oImp.gPrnSaltoLinea
ElseIf Mid(psCuenta, 6, 3) = gCapAhorros Then
    sCad = sCad & Space(nEspaciosIzq) & "BLOQUEO PARCIAL   " & Space(75) & psSaldoRet & oImp.gPrnSaltoLinea
End If
sCad = sCad & Space(nEspaciosIzq) & "INTERÉS DEL MES   " & Space(75) & psInteresMes & oImp.gPrnSaltoLinea

If Mid(psCuenta, 6, 3) = gCapPlazoFijo Then
    sCad = sCad & Space(nEspaciosIzq) & "INTERÉS PAGADO    " & Space(75) & sIntPag & oImp.gPrnSaltoLinea
End If
sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "=") & oImp.gPrnSaltoLinea             ' hasta 160 caracteres x línea

'JUEZ 20130727 *****************************
If Mid(psCuenta, 6, 3) = gCapCTS Then
    sCad = sCad & "  * Saldo total de retiro de la(s) cuenta(s) CTS del cliente con el mismo empleador." & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
End If
'END JUEZ **********************************

'INICIO EAAS20180511 COMENTO
'Dim lsMsg As String
'lsMsg = "Advertencia: Si dentro de 30 días no se formula observación al presente estado, daremos por conforme la cuenta y aprobado el saldo. En caso contrario sírvase dirigirse a nuestras oficinas para atender sus observaciones." & oImp.gPrnSaltoLinea & "En caso de reclamos y servicios, el cliente podrá recurrir indistintamente a las siguientes instancias:" & oImp.gPrnSaltoLinea & "1. A  nuestra  Red  de  Oficinas.  2.  A  la  Plataforma de Atención al Usuario de la Superintendencia de Banca y Seguros 3. INDECOPI."
'sCad = sCad & JustificaTextoSinCarEsp(lsMsg, nCarLin, nEspaciosIzq) & oImp.gPrnSaltoLinea
'sCad = sCad & Space(nEspaciosIzq) & psEstado & oImp.gPrnSaltoLinea
'sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea
'sCad = sCad & Space(nEspaciosIzq) & gsCodUser EAAS
'FIN EAAS20180511

'--=========================================
'--INICIO EAAS20180511 SEGUN ERS TI-006-2018
'--=========================================
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsCta As ADODB.Recordset
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsCta = New ADODB.Recordset
Set rsCta = clsMant.GetDatosCuenta(psCuenta)
Dim nTpoPrograma As Integer
nTpoPrograma = rsCta("nTpoPrograma")

Dim lsMsgReten As String
If Mid(psCuenta, 6, 3) = "232" Then
    If (nTpoPrograma <> 0 And nTpoPrograma <> 5 And nTpoPrograma <> 6) Then
    lsMsgReten = "Nuestro compromiso es servirle cada día mejor, por lo que si tuviese alguna consulta o desee efectuar un reclamo puede apersonarse a nuestra red de agencias o llamar al Fono Maynas 065-581800 (Loreto) o 0801-10700 (Lima y provincias). Adicionalmente, usted podrá recurrir a Indecopi o  a  la  Plataforma de Atención al Usuario de la SBS si considera que su reclamo no es atendido adecuadamente."
    ElseIf (nTpoPrograma = 0 Or nTpoPrograma = 5) Then
    lsMsgReten = "Estimado Cliente, se le recuerda que si existiera alguna diferencia entre su Saldo Contable y Saldo Disponible se puede atribuir a las siguientes razones:" & oImp.gPrnSaltoLinea & "" & oImp.gPrnSaltoLinea & " -Retiros en Cajeros Automáticos (Global Net y VISA)" & oImp.gPrnSaltoLinea & "-Compras POS (Red VISA)" & oImp.gPrnSaltoLinea & "-Reembolsos VISA NET - Establecimientos afiliados al servicio de banco pagador." & oImp.gPrnSaltoLinea & " " & oImp.gPrnSaltoLinea & "Nuestro compromiso es servirle cada día mejor, por lo que si tuviese alguna consulta o desee efectuar un reclamo puede apersonarse a nuestra red de agencias o llamar al Fono Maynas 065-581800 (Loreto) o 0801-10700 (Lima y provincias). Adicionalmente, usted podrá recurrir a Indecopi o  a  la  Plataforma de Atención al Usuario de la SBS si considera que su reclamo no es atendido adecuadamente."
    Else
    lsMsgReten = "Estimado Cliente, se le recuerda que si existiera alguna diferencia entre su Saldo Contable y Saldo Disponible se puede atribuir a las siguientes razones:" & oImp.gPrnSaltoLinea & "" & oImp.gPrnSaltoLinea & " -Retiros en Cajeros Automáticos (Global Net y VISA)" & oImp.gPrnSaltoLinea & "-Compras POS (Red VISA)" & oImp.gPrnSaltoLinea & "" & oImp.gPrnSaltoLinea & " " & oImp.gPrnSaltoLinea & "Nuestro compromiso es servirle cada día mejor, por lo que si tuviese alguna consulta o desee efectuar un reclamo puede apersonarse a nuestra red de agencias o llamar al Fono Maynas 065-581800 (Loreto) o 0801-10700 (Lima y provincias). Adicionalmente, usted podrá recurrir a Indecopi o  a  la  Plataforma de Atención al Usuario de la SBS si considera que su reclamo no es atendido adecuadamente."
    End If
    'lsMsgReten = lsMsgReten & oImp.gPrnSaltoLinea & "Le recordamos verificar la información contenida en el estado de cuenta. Si tuviese  alguna observación al respecto,  le solicitamos contactarse con nosotros a través de nuestra red de agencias dentro de los 30 días calendario siguientes a la recepción de este documento; en caso contrario, daremos por conforme la cuenta y aprobado el saldo. "
ElseIf Mid(psCuenta, 6, 3) = "233" Then
    lsMsgReten = "Nuestro compromiso es servirle cada día mejor, por lo que si tuviese alguna consulta o desee efectuar un reclamo puede apersonarse a nuestra red de agencias o llamar al Fono Maynas 065-581800 (Loreto) o 0801-10700 (Lima y provincias). Adicionalmente, usted podrá recurrir a Indecopi o  a  la  Plataforma de Atención al Usuario de la SBS si considera que su reclamo no es atendido adecuadamente."
    'lsMsgReten = lsMsgReten & oImp.gPrnSaltoLinea & "Le recordamos verificar la información contenida en el estado de cuenta. Si tuviese  alguna observación al respecto,  le solicitamos contactarse con nosotros a través de nuestra red de agencias dentro de los 30 días calendario siguientes a la recepción de este documento; en caso contrario, daremos por conforme la cuenta y aprobado el saldo. "
ElseIf Mid(psCuenta, 6, 3) = "234" Then
    lsMsgReten = "Estimado Cliente, se le recuerda que si existiera alguna diferencia entre su Saldo Contable y Saldo Disponible, se puede atribuir a las siguientes razones:" & oImp.gPrnSaltoLinea & " -Retiros en Cajeros Automáticos (Global Net y VISA)" & oImp.gPrnSaltoLinea & "-Compras POS (Red VISA)" & oImp.gPrnSaltoLinea & "Nuestro compromiso es servirle cada día mejor, por lo que si tuviese alguna consulta o desee efectuar un reclamo puede apersonarse a nuestra red de agencias o llamar al Fono Maynas 065-581800 (Loreto) o 0801-10700 (Lima y provincias) donde le estaremos atendiendo gustosamente. Adicionalmente, en caso no estuviese conforme también podrá recurrir al Indecopi o  a  la  Plataforma de Atención al Usuario de la Superintendencia de Banca, Seguros y AFP."
    'lsMsgReten = lsMsgReten & oImp.gPrnSaltoLinea & "Le recordamos verificar la información contenida en el estado de cuenta. Si tuviese  alguna observación al respecto,  le solicitamos contactarse con nosotros a través de nuestra red de agencias dentro de los 30 días calendario siguientes a la recepción de este documento; en caso contrario, daremos por conforme la cuenta y aprobado el saldo. "
End If

sCadR = sCadR & JustificaTextoSinCarEsp(lsMsgReten, nCarLin, nEspaciosIzq) & oImp.gPrnSaltoLinea
sCadR = sCadR & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea
sCadR = sCadR & Space(nEspaciosIzq) & gsCodUser
'--======================================
'--FIN EAAS20180511 SEGUN ERS TI-006-2018
'--======================================

sCadAux = sCadAux & sCad & sCadR '--===========================EAAS20180511 sCadR
pnNroPag = nCntPag
Set oImp = Nothing
ImprimeConsultaMovimiento = sCadAux
        
End Function

Public Function ImprimeConsultaMovimientoCobranza(ByVal rs As ADODB.Recordset, ByVal rs1 As ADODB.Recordset, ByVal gsCodUser As String, _
                ByVal gsNomCmac As String, ByVal gsNomAge As String, ByVal sMoneda As String, ByVal gdFecSis As String, _
                ByVal sTitRp1 As String, ByVal sTitRp2 As String, ByVal sCuenta As String, ByVal psFecSaldIni As String, _
                ByVal pdSaldoInicial As String, ByVal psSaldoDisponible As String, ByVal psSaldoContable As String, ByVal psSaldoRet As String, _
                ByVal psInteresMes As String, ByVal psCapPlazoFijo As String, ByVal pnIntPag As Integer, ByVal psEstado As String, ByVal psTotAbono As String, ByVal pnNroPag As Integer) As String
      
  Dim sCad As String
  Dim sCadAux As String
  Dim nCarLin As Integer
  Dim nLinPag As Integer
  Dim sNumPag As String
  Dim nCntPag As Integer
  Dim sOperacion As String * 28
  Dim sDocumento As String * 15
  Dim sDeposito As String * 13
  Dim sPersona As String * 14
  Dim sSaldoCnt As String * 14
  Dim ofunc As New COMFunciones.FCOMCadenas
  Dim oFunI As New COMFunciones.FCOMImpresion
  
  nCarLin = 118
  nLinPag = 65
  sNumPag = ""
  nCntPag = 0
  
  If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    Do Until rs.EOF
        If nLinPag > 58 Then
            If sNumPag <> "" Then
                sCad = sCad & Space(30) & String(nCarLin, "-") & oImp.gPrnSaltoLinea
                sCad = sCad & Space(30) & gsCodUser
                'sCad = sCad & Chr$(12)
            End If
                
            sNumPag = ofunc.FillNum(Trim(nCntPag), 4, " ")
            sCad = sCad & oImp.gPrnSaltoLinea                                                                                                                    'ncarlin
            sCad = sCad & oFunI.CabeRepo(gsNomCmac, gsNomAge, "SECCION AHORROS", sMoneda, Format$(gdFecSis, "dd/mm/yyyy"), sTitRp1, sTitRp2, "", "", nCntPag, 80, 0, oImp) & oImp.gPrnSaltoLinea
            'nCntPag = nCntPag + 1
            sCad = sCad & Space(30) & "Cuenta No. :   " & sCuenta & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
            sCad = sCad & Space(30) & "CLIENTE" & Space(50) & "RELACION" & Space(20) & "DIRECCION" & oImp.gPrnSaltoLinea
            If Not (rs1.EOF And rs1.BOF) Then
                rs1.MoveFirst
                Do Until rs1.EOF
                    sCad = sCad & Space(30) & Left(Trim(rs1(0)), 59) & Space(60 - Len(Left(Trim(rs1(0)), 50))) & Trim(rs1(1)) & Space(10) & Trim(rs1(2)) & oImp.gPrnSaltoLinea
                    rs1.MoveNext
                Loop
            End If
                'sCad = sCad & Space(30) & "------------------------" & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
                sCad = sCad & Space(30) & "SALDO AL " & psFecSaldIni & " : " & ofunc.JDNum(pdSaldoInicial, 12, True, 7, 2) & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
                sCad = sCad & Space(30) & String(nCarLin, "-") & oImp.gPrnSaltoLinea
                sCad = sCad & Space(30) & "      FECHAS        "
                sCad = sCad & "     TIPO DE OPERACION      "
                sCad = sCad & "      DEPOSITOS   "
                sCad = sCad & "CONTABLE "
                sCad = sCad & Space(30) & String(nCarLin, "-") & oImp.gPrnSaltoLinea            ' hasta 160 caracteres x línea
                nLinPag = 31
                If (nCntPag Mod 3) = 0 Then
                    sCadAux = sCadAux & sCad
                    sCad = ""
                End If
            End If
            sCad = sCad & Space(30) & Trim(rs(0)) & " "
            sOperacion = oFunI.ImpreCarEsp(rs(1)) & " "
            RSet sDeposito = rs(3) & " "
            sPersona = rs(4) & " "
             
            sCad = sCad & sOperacion
            sCad = sCad & sDeposito
            sCad = sCad & sSaldoCnt
            
            nLinPag = nLinPag + 1
           rs.MoveNext
     Loop
   End If
        sCad = sCad & Space(30) & String(nCarLin, "-") & oImp.gPrnSaltoLinea              ' hasta 160 caracteres x línea
        sCad = sCad & Space(30) & "  TOTAL    " & Space(52)
        sCad = sCad & psTotAbono & " "
        'scad = scad & sTotCargo & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        sCad = sCad & Space(30) & "SALDO DISPONIBLE  " & Space(71) & psSaldoDisponible & oImp.gPrnSaltoLinea
        sCad = sCad & Space(30) & "SALDO CONTABLE    " & Space(71) & psSaldoContable & oImp.gPrnSaltoLinea
        If Mid(sCuenta, 6, 3) = "234" Then
            sCad = sCad & Space(30) & "SALDO RETIRO ACTUAL  " & Space(74) & psSaldoRet & oImp.gPrnSaltoLinea
        End If
        sCad = sCad & Space(30) & "INTERÉS DEL MES   " & Space(71) & psInteresMes & oImp.gPrnSaltoLinea
        If Mid(sCuenta, 6, 3) = psCapPlazoFijo Then
            sCad = sCad & Space(30) & "INTERÉS PAGADO    " & Space(80) & Format(pnIntPag, "#0.00") & oImp.gPrnSaltoLinea
        End If
        sCad = sCad & Space(30) & String(nCarLin, "=") & oImp.gPrnSaltoLinea             ' hasta 160 caracteres x línea
        sCad = sCad & Space(30) & psEstado & oImp.gPrnSaltoLinea    '   & Chr$(12)
        sCad = sCad & Space(30) & String(nCarLin, "-") & oImp.gPrnSaltoLinea
        sCad = sCad & Space(30) & gsCodUser
        sCadAux = sCadAux & sCad
        
        pnNroPag = nCntPag
        
        ImprimeConsultaMovimientoCobranza = sCadAux
        
End Function

Public Function ImprimeBeneficiario(ByVal gsNomAge As String, ByVal gdFecSis As String, ByVal psNomCli As String, ByVal psDirCli As String, _
                                    ByVal psCodCli As String, ByVal psDocId As String, ByVal pnPorcAcum As Double, ByVal rs As ADODB.Recordset) As String
Dim lnCarLin As Integer
Dim lsRTFExt As String, lsTitRp1 As String
Dim lsTitRp2 As String, lsMoneda As String
Dim lsNumPag As String, lsFchtra As String
Dim lnLinPag As Integer, lnCntPag As Integer
Dim lsNombre As String * 35
Dim lsPar As String * 10
Dim lnItem As Integer
Dim lnLenNom As Integer
Dim i As Integer
Dim lnAcum As Double
Dim oFunI As New COMFunciones.FCOMImpresion
Dim ofunc As New COMFunciones.FCOMCadenas


lsNumPag = ""
lnLinPag = 0
lnCarLin = 90
lsTitRp1 = "D E C L A R A C I O N   D E   B E N E F I C I A R I O S"
lsTitRp2 = "S E G U R O  D E  S E P E L I O"
lnCntPag = 1
lnAcum = 0
lnItem = 0
lsMoneda = ""
    lnLenNom = Len(oFunI.ImpreCarEsp(psNomCli))
    lsNumPag = ofunc.FillNum(Trim(Str(lnCntPag)), 4, " ")
    lsRTFExt = lsRTFExt & CabeRepoCaptac("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag, gsNomAge, gdFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & String(lnCarLin, "=") & oImp.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & "CODIGO    : " & Trim(psCodCli) & oImp.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & "ASEGURADO : " & oFunI.ImpreCarEsp(psNomCli) & oImp.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & "DIRECCION : " & Trim(psDirCli) & oImp.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & "DOC. ID.  : " & Trim(psDocId) & oImp.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & oImp.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & String(lnCarLin, "=") & oImp.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & "ITEM  CODIGO      NOMBRE                               EDAD   PARENTESCO   PORCENTAJE" & oImp.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & String(lnCarLin, "=") & oImp.gPrnSaltoLinea
    lnLinPag = 13
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            lnItem = lnItem + 1
            lsNombre = oFunI.ImpreCarEsp(rs(1))
            lsPar = rs(2)
            lsRTFExt = lsRTFExt & ofunc.JDNum(Trim(Str(lnItem)), 3, False, 3, 0) & "   " & Trim(rs(0)) & "  " & lsNombre & "  " & ofunc.JDNum(Trim(CDbl(rs(3))), 5, True, 5, 0) & "  " & lsPar & "  " & ofunc.JDNum(Trim(Val(rs(4))), 12, True, 9, 2) & oImp.gPrnSaltoLinea
            lnLinPag = lnLinPag + 1
            rs.MoveNext
        Loop
    End If
    lsRTFExt = lsRTFExt & String(lnCarLin, "-") & oImp.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & String(56, " ") & "PORCENTAJE ACUMULADO    " & ofunc.JDNum(Trim(Str(pnPorcAcum)), 5, True, 3, 2) & oImp.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & String(lnCarLin, "-") & oImp.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    
    lsRTFExt = lsRTFExt & Space((lnCarLin - lnLenNom) / 2) & String(lnLenNom + 4, "_") & oImp.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & Space(((lnCarLin - lnLenNom) / 2) + 2) & oFunI.ImpreCarEsp(psNomCli) & oImp.gPrnSaltoLinea
    
    ImprimeBeneficiario = lsRTFExt
    
 Set oFunI = Nothing
 Set ofunc = Nothing
End Function

Private Sub Class_Initialize()
IniciaImpresora
End Sub

Private Sub Class_Terminate()
Set oImp = Nothing
End Sub

Public Function ImprimeBoletaAutorizacionTCE(ByVal psTit As String, ByVal pstext As String, ByVal psCodOpe As String, ByVal pnIDAprob As Integer, _
                                             ByVal pnTCNew As Double, ByVal pnValor As Double, ByVal psNomAge As String, ByVal psNomCmac As String, ByVal pdFecSis As Date, ByVal psCodUser As String) As String

   
Dim sCad As String
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax As Integer
Dim saux As Integer
Dim lsNroExt As String
Dim lnTope As Integer
Dim lsSaldoC As String
Dim lsCadArg As String
Dim lsInteres As String
Dim sNomAge As String
Dim oImpre As New COMFunciones.FCOMImpresion

sCad = sCad & oImp.gPrnInicializa
sCad = sCad & oImp.gPrnEspaLineaN   'espaciamiento lineas 1/6 pulg.
sCad = sCad & oImp.gPrnTamPagina22  'Longitud de página a 22 líneas'
sCad = sCad & oImp.gPrnTamLetra10CPI   'Tamaño 10 cpi
sCad = sCad & oImp.gPrnTpoLetraRoman     'Tipo de Letra Sans Serif
sCad = sCad & oImp.gPrnCondensadaOFF ' cancela condensada
sCad = sCad & oImp.gPrnNegritaOFF ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
saux = 5

sFecha = Format$(pdFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sNomAge = psNomAge

sCad = sCad & oImp.gPrnNegritaON 'Activa Negrita

sCad = sCad & oImp.gPrnSaltoLinea & oImp.gPrnNegritaON 'Activa Negrita
sCad = sCad & Space(sIni) & "CMAC-MAYNAS -AHORRO" & Space(14 + sSep + saux) & "CMAC-MAYNAS - AHORRO" & oImp.gPrnSaltoLinea

lsNroExt = ""

sCad = sCad & Space(sIni) & Trim(sNomAge) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt)) + lsNroExt & Space(sSep) & Trim(sNomAge) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt)) + lsNroExt & oImp.gPrnSaltoLinea

sCad = sCad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea

psTit = Trim(psTit)
psTit = oImpre.CentrarCadena(psTit, 28)
sCad = sCad & oImp.gPrnNegritaON 'Activa Negrita
sCad = sCad & Space(sIni + 1) & "-----" & psTit & "-----" & Space(-1 + sSep) & "-----" & psTit & "-----" & oImp.gPrnSaltoLinea

sCad = sCad & oImp.gPrnNegritaOFF 'Desactiva Negrita

sCad = sCad & Space(sIni) & "Nuevo Tipo de Cambio :" & Space(sMax + 6 - Len("Nuevo Tipo de Cambio") - Len(pnTCNew)) & pnTCNew & Space(-1 + sSep + 2) & "Nuevo Tipo de Cambio :" & Space(sMax + 6 - Len("Nuevo Tipo de Cambio") - Len(pnTCNew)) & pnTCNew & oImp.gPrnSaltoLinea
sCad = sCad & Space(sIni) & "Monto a Cambiar : " & Space(sMax + 6 - Len("Monto a Cambiar") - Len(pnValor)) & pnValor & Space(-1 + sSep + 2) & "Monto a Cambiar :" & Space(sMax + 6 - Len("Monto a Cambiar") - Len(pnValor)) & pnValor & oImp.gPrnSaltoLinea
            
sCad = sCad & Space(sIni) & "ID Autorización : " & Space(17) & pnIDAprob & Space(16) & "ID Autorización :" & Space(17) & pnIDAprob & oImp.gPrnSaltoLinea

sCad = sCad & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & "---------------------------------------" & oImp.gPrnSaltoLinea
'By Capi 10042008 para que corte despues de la impresion
'sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(psCodUser) & Space(29 + sSep + saux) & oImpre.ImpreCarEsp(psCodUser) & oImp.gPrnSaltoLinea
sCad = sCad & Space(sIni) & oImpre.ImpreCarEsp(psCodUser) & Space(29 + sSep + saux) & oImpre.ImpreCarEsp(psCodUser) & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
'

ImprimeBoletaAutorizacionTCE = sCad
Exit Function
Error:

    Err.Raise Err.Number, "Error al realizar la Impresión...", Err.Description
End Function

Public Function JustificaTextoSinCarEsp(sTemp As String, lnColPage As Integer, Optional lsEspIzq As Integer = 0) As String
Dim vTextFin As String
Dim Letra As String * 1, i As Integer, K As Integer, N As Integer
Dim nVeces As Long, m As Integer, Fin As Integer, Ini As Integer
Dim nAncho1 As Integer, nSpa As Integer
i = 0
K = 0
N = Len(sTemp)
nAncho1 = lnColPage
Do While i <= N
   K = K + 1
   i = i + 1
   If i > N Then
      Exit Do
   End If
   Letra = Mid(sTemp, i, 1)
   If Letra = Chr$(27) Then
      vTextFin = vTextFin & Letra & Mid(sTemp, i + 1, 1)
      i = i + 1
      K = K + 1
      nAncho1 = nAncho1 + 2
   Else
      If Asc(Letra) <> 13 And Asc(Letra) <> 10 Then
         If K > nAncho1 Then
            m = 0
            If Mid(sTemp, i, 1) = Chr(32) Then
               vTextFin = Trim(vTextFin)
            Else
               m = InStrRev(vTextFin, " ", , vbTextCompare)
               If m = 0 Then m = 1
               If InStr(Mid(vTextFin, m, Len(vTextFin)), Chr$(27)) Then
                  nAncho1 = nAncho1 - 2
               End If
               i = i - (nAncho1 + 1 - m)
               vTextFin = Mid(vTextFin, 1, m - 1)
            End If
            nSpa = nAncho1 - Len(Trim(vTextFin))
            vTextFin = Trim(vTextFin)
            If nSpa <> 0 Then
               Fin = 1
               nVeces = 0
               m = 1
               Do While m <= nSpa
                  Ini = InStr(Fin, vTextFin, " ", vbTextCompare)
                  If Ini = 0 Then
                     Fin = 1
                     nVeces = nVeces + 1
                     m = m + 1
                  Else
                      vTextFin = Mid(vTextFin, 1, Ini) & " " & RTrim(Mid(vTextFin, Ini + 1, nAncho1))
                      Fin = Ini + 2 + nVeces
                      m = m + 1
                  End If
               Loop
            End If
            vTextFin = vTextFin & Chr(10)
            JustificaTextoSinCarEsp = JustificaTextoSinCarEsp & Space(lsEspIzq) & Trim(vTextFin)
            nAncho1 = lnColPage
            vTextFin = ""
            Letra = ""
            K = 0
         Else
            vTextFin = vTextFin & Letra
         End If
      Else
        If i < N Then
            If Asc(Mid(sTemp, i + 1, 1)) = 13 Or Asc(Mid(sTemp, i + 1, 1)) = 10 Then
                i = i + 1
            End If
        End If
         JustificaTextoSinCarEsp = JustificaTextoSinCarEsp & Space(lsEspIzq) & Trim(vTextFin) & Chr(10)
         nAncho1 = lnColPage
         vTextFin = ""
         Letra = ""
         K = 0
      End If
   End If
Loop
JustificaTextoSinCarEsp = JustificaTextoSinCarEsp & Space(lsEspIzq) & Trim(vTextFin)
End Function

'**************************************************************
Public Function ImprimeDetalleCuenta(ByVal sCuenta As String, ByVal sProducto As String, ByVal sPrograma As String, ByVal dFecApe As Date, _
                                     ByVal sTipCuenta As String, ByVal nPlazo As String, ByVal sTipTasa As String, ByVal sEstado As String, _
                                     ByVal sFormaRet As String, ByVal nTasa As Currency, ByVal sMoneda As String, ByVal dFecVenc As String, _
                                     ByVal bOrdPag As Boolean, ByVal bAboIntCtaAho As Boolean, ByVal rsProdPers As ADODB.Recordset, _
                                     ByVal sAgencia As String, ByVal dFecSis As Date, ByVal sUser As String, ByVal nTpoPrograma As Integer, _
                                     ByVal sCuentaAbo As String) As String
         
    Dim oConect As New COMConecta.DCOMConecta
    Dim i As Integer
    Dim sCad As String
    Dim sNumPag As String

    sCad = ""
    
    Dim nTamMayor As Long, nTamMenor As Long, nTamAdd As Long
    Dim nItem As Integer
    Dim nPagina As Integer

    nTamMayor = 70
    nTamMenor = 38
    nTamAdd = 22
    nItem = 12
    nPagina = 1
     
    Set oConect = New COMConecta.DCOMConecta
    
    'sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "", sTitRp1, sTitRp2, "", Trim(CStr(nPagina)), sAgencia, dFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
    sCad = sCad & sAgencia & Space(80 - Len(sAgencia)) & "PAGINA: " & Trim(CStr(nPagina)) & oImp.gPrnSaltoLinea
    sCad = sCad & Space(80) & "FECHA: " & Format$(dFecSis, "dd/mm/yyyy") & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    
    sCad = sCad & "Cuenta: " & sCuenta & Space(10) & "Producto: " & sProducto & Space(65 - Len(sProducto) - 46) & "SubProducto: " & sPrograma & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    
    sCad = sCad & "Apertura  : " & dFecApe & Space(13) & "Tipo Cuenta : " & sTipCuenta & Space(65 - 44 - Len(sTipCuenta)) & "Plazo (días)  :" & nPlazo & oImp.gPrnSaltoLinea
    sCad = sCad & "Tipo Tasa : " & sTipTasa & Space(35 - 12 - Len(sTipTasa)) & "Estado      : " & sEstado & Space(65 - 44 - Len(sEstado)) & "Forma Retiro  :" & sFormaRet & oImp.gPrnSaltoLinea
    sCad = sCad & "Tasa (%)  : " & Format(nTasa, "#0.0000") & Space(35 - 12 - Len(Format(nTasa, "#0.0000"))) & "Moneda      : " & sMoneda & Space(65 - 44 - Len(sMoneda)) & "Vencimiento   :" & dFecVenc & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    sCad = sCad & "CLIENTES:" & oImp.gPrnSaltoLinea
    sCad = sCad & String(9, "=") & oImp.gPrnSaltoLinea
    
    sCad = sCad & "CODIGO" & Space(12) & "NOMBRE" & Space(50) & "RELACION" & Space(15) & "OBLIGATORIO" & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 16, "=") & oImp.gPrnSaltoLinea

    rsProdPers.MoveFirst
    Do While Not rsProdPers.EOF
        sCad = sCad & rsProdPers!Codigo & Space(5) & Trim(rsProdPers!Nombre) & Space(74 - Len(Trim(rsProdPers!Nombre)) - 18) & Trim(Left(rsProdPers!Relacion, Len(rsProdPers!Relacion) - 2)) & Space(89 - Len(Trim(Left(rsProdPers!Relacion, Len(rsProdPers!Relacion) - 2))) - 62) & rsProdPers!Obligatorio & oImp.gPrnSaltoLinea
        rsProdPers.MoveNext
    Loop
    
    sCad = sCad & String(nTamMayor + nTamAdd + 16, "=") & oImp.gPrnSaltoLinea
    sCad = sCad & "Orden de Pago         :" & Space(5) & IIf(bOrdPag = True, "SI", "NO") & oImp.gPrnSaltoLinea
    'sCad = sCad & "Abonar Int. Cta. Aho. :" & Space(5) & IIf(bOrdPag = True, "SI", "NO") & oImp.gPrnSaltoLinea
    sCad = sCad & "Abonar Int. Cta. Aho. :" & Space(5) & IIf(bAboIntCtaAho = True, "SI", "NO") & oImp.gPrnSaltoLinea 'FRHU 20140212 RQ14049
    'sCad = sCad & "Cuenta Abono Int.     :" & oImp.gPrnSaltoLinea
    sCad = sCad & "Cuenta Abono Int.     :" & sCuentaAbo & oImp.gPrnSaltoLinea 'FRHU 20140212 RQ14049

    ImprimeDetalleCuenta = sCad
End Function

'***Agregado por ELRO el 20130710, según RFC1306270002****
Public Function imprimirBoletaServicioPago(ByVal psNombreCMAC As String, ByVal psNomAge As String, ByVal psMoneda As String, _
                                           ByVal pnMovNro As Long, ByVal psFecha As String, ByVal psHora As String, _
                                           ByVal psEmpresa As String, ByVal psDOI As String, ByVal psBeneficiario As String, _
                                           ByVal pnDebito As Currency, ByVal psUsurio As String, ByRef psLpt As String, _
                                           ByRef pbImpTMU As Boolean, Optional ByVal pbExtorno As Boolean = False, _
                                           Optional ByVal psMovNroExt As String = "", _
                                           Optional ByVal psNombreConvenio As String = "") As String
                                           

    Dim sSep As Integer
    Dim sIni As Integer
    Dim sMax As Integer
    Dim saux As Integer
    Dim sCad As String
    Dim lnTope As Integer
    Dim oFCOMImpresion As New COMFunciones.FCOMImpresion
    Dim lscad As String
    Dim lnNomConSerPag As String * 20


ETIQ:

    On Error GoTo Error

    lnTope = 0 '6 'Tope de lineas en Boleta

    lscad = lscad & oImp.gPrnInicializa
    lscad = lscad & oImp.gPrnEspaLineaN   'espaciamiento lineas 1/6 pulg.
    lscad = lscad & oImp.gPrnTamPagina22   'Longitud de página a 22 líneas'
    lscad = lscad & oImp.gPrnTamLetra10CPI   'Tamaño 10 cpi
    lscad = lscad & oImp.gPrnTpoLetraRoman     'Tipo de Letra Sans Serif
    lscad = lscad & oImp.gPrnCondensadaOFF ' cancela condensada
    lscad = lscad & oImp.gPrnNegritaOFF ' desactiva negrita

    sSep = 19
    sIni = 0
    sMax = 33
    saux = 17
    
    lnNomConSerPag = psNombreConvenio
    
    If pbImpTMU = False Then
        lscad = lscad & Space(sIni) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psNombreCMAC & Space(sSep + sMax - Len(Trim(psNombreCMAC)) + 2) & psNombreCMAC & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & UCase(psNomAge) & "-" & UCase(psMoneda) & Space(2) & "MOV: " & CStr(pnMovNro) & Space(saux + sMax - Len(Trim(psNomAge)) - Len(Trim(psMoneda)) - Len(Trim(pnMovNro)) - 4) & UCase(psNomAge) & "-" & UCase(psMoneda) & Space(2) & "MOV: " & CStr(pnMovNro) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "FECHA: " & UCase(psFecha) & Space(7) & "HORA: " & CStr(psHora) & Space(saux + sMax - Len(Trim(psFecha)) - 24) & "FECHA: " & UCase(psFecha) & Space(7) & "HORA: " & CStr(psHora) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psEmpresa & Space(saux + sMax - Len(Trim(psEmpresa)) + 4) & psEmpresa & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "DNI: " & psDOI & Space(saux + sMax - Len(Trim(psDOI)) - 1) & "DNI: " & psDOI & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psBeneficiario & Space(saux + sMax - Len(Trim(psBeneficiario)) + 4) & psBeneficiario & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        '''''''''''''''''''''''''''''''--------------------------------------
        '''''''''''''''''''''''''''''''01234567890123456789012345678901234567
        If pbExtorno Then
        lscad = lscad & Space(sIni) & "-------EXTORNO SERVICIO DE PAGO-------" & Space(16) & "-----------SERVICIO DE PAGO-----------" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Ext. Debi. Ser. Pago" & Space(18 - Len(Format$(pnDebito, "##,##0.00"))) & Format$(pnDebito, "##,##0.00") & Space(16) & "Ext. Debi. Ser. Pago" & Space(18 - Len(Format$(pnDebito, "##,##0.00"))) & Format$(pnDebito, "#,##0.00") & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "--------------------------------------" & Space(16) & "--------------------------------------" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & Right(psMovNroExt, 4) & Space(saux + sMax - Len(Right(psMovNroExt, 4)) + 4) & Right(psMovNroExt, 4) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psUsurio & Space(saux + sMax - Len(psUsurio) + 4) & psUsurio & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        Else
        lscad = lscad & Space(sIni) & "-----------SERVICIO DE PAGO-----------" & Space(16) & "-----------SERVICIO DE PAGO-----------" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & Left(lnNomConSerPag, 20) & Space(18 - Len(Format$(pnDebito, "##,##0.00"))) & Format$(pnDebito, "##,##0.00") & Space(16) & Left(lnNomConSerPag, 20) & Space(18 - Len(Format$(pnDebito, "##,##0.00"))) & Format$(pnDebito, "#,##0.00") & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "--------------------------------------" & Space(16) & "--------------------------------------" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psUsurio & Space(saux + sMax - Len(psUsurio) + 4) & psUsurio & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "***ESTIMADO CLIENTE,ANTES DE RETIRARSE" & Space(16) & "***ESTIMADO CLIENTE,ANTES DE RETIRARSE" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "DE LA VENTANILLA VERIFIQUE QUE LOS DA-" & Space(16) & "DE LA VENTANILLA VERIFIQUE QUE LOS DA-" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "TOS CONSIGNADOS CORRESPONDAN A LA OPE-" & Space(16) & "TOS CONSIGNADOS CORRESPONDAN A LA OPE-" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "RACIÓN SOLICITADA POR USTED***        " & Space(16) & "RACIÓN SOLICITADA POR USTED***        " & oImp.gPrnSaltoLinea
        End If
        lscad = lscad & oFCOMImpresion.ImpreCarEsp(sCad)
        imprimirBoletaServicioPago = lscad & oImp.gPrnSaltoLinea
    Else 'TMU
        lscad = lscad & Space(sIni) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psNombreCMAC & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & UCase(psNomAge) & "-" & UCase(psMoneda) & Space(2) & "MOV: " & CStr(pnMovNro) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "FECHA: " & UCase(psFecha) & Space(7) & "HORA: " & CStr(psHora) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psEmpresa & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "DNI: " & psDOI & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psBeneficiario & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        '''''''''''''''''''''''''''''''--------------------------------------
        '''''''''''''''''''''''''''''''01234567890123456789012345678901234567
        If pbExtorno Then
        lscad = lscad & Space(sIni) & "-------EXTORNO SERVICIO DE PAGO-------" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Ext. Debi. Ser. Pago" & Space(18 - Len(Format$(pnDebito, "##,##0.00"))) & Format$(pnDebito, "##,##0.00") & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "--------------------------------------" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & Right(psMovNroExt, 4) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psUsurio & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        Else
        lscad = lscad & Space(sIni) & "-----------SERVICIO DE PAGO-----------" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & Left(lnNomConSerPag, 20) & Space(18 - Len(Format$(pnDebito, "##,##0.00"))) & Format$(pnDebito, "##,##0.00") & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "--------------------------------------" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psUsurio & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "***ESTIMADO CLIENTE,ANTES DE RETIRARSE" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "DE LA VENTANILLA VERIFIQUE QUE LOS DA-" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "TOS CONSIGNADOS CORRESPONDAN A LA OPE-" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "RACIÓN SOLICITADA POR USTED***        " & oImp.gPrnSaltoLinea
        End If
        lscad = lscad & oFCOMImpresion.ImpreCarEsp(sCad)
        'Linea que corta el ticket en TMU
        imprimirBoletaServicioPago = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
    End If
    
Exit Function
Error:
    Err.Raise Err.Number, "Comprueba la conexion de su impresora... ", Err.Description
'    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
'        GoTo ETIQ
'    End If
End Function

'***Fin Agregado por ELRO el 20130710, según RFC1306270002

'RECO20140311 ERS160-2013************************************************************************************
Public Function RecuperaDatosRelacCred(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_RecuperaDatosRelacCred '" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set RecuperaDatosRelacCred = rs
    'oConec.CierraConexion
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function RecuperaDatosCreditoVigente(ByVal psCtaCod As String, Optional ByVal pdHoy As Date = CDate("01/01/1900"), _
    Optional ByVal pbIncluirbAprobados As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosCreditoVigente
    'ALPA 20100607 B2**************************************************************
    'sSql = "Select PP.cPersCod, CC.cMetLiquidacion, dbo.CuotaAprobada(P.cCtaCod) as CuotaAprobada, CC.bCuotaCom, CC.bMiVivienda, CC.nCalendDinamTipo, CC.bPrepago, CC.nCalPago, Pers.cPersNombre, C.cLineaCred,C.nMontoCol, P.nSaldo, "
    sSql = "Select C.cAgeCodAct,PP.cPersCod, CC.cMetLiquidacion, dbo.CuotaAprobada(P.cCtaCod) as CuotaAprobada, CC.bCuotaCom, CC.bMiVivienda, CC.nCalendDinamTipo, CC.bPrepago, CC.nCalPago, Pers.cPersNombre, C.cLineaCred,C.nMontoCol, P.nSaldo, Pers.nPersPersoneria, " 'JUEZ 20130925 Se agregó Pers.nPersPersoneria
    '******************************************************************************
    sSql = sSql & " cMoneda = (Select cConsDescripcion from Constante where nConsCod = " & COMDConstantes.gMoneda & " AND nConsValor = " & CInt(Mid(psCtaCod, 9, 1)) & "),"
    sSql = sSql & " CE.nCuotas nCuotasApr, CE.nColocCalendCod " & ", "
    sSql = sSql & " CC.nDiasAtraso, CC.cMetLiquidacion, P.nTransacc, CC.nCalendDinamico, CC.nIntPend, CC.nNroCalen, CC.nNroProxCuota, CC.nNroProxDesemb "
    If pdHoy <> CDate("01/01/1900") Then
        sSql = sSql & ", DATEDIFF(year,C.dVigencia,'" & Format(pdHoy, "mm/dd/yyyy") & "') as nPlazoTranscurrido "
    End If
    'ALPA 20100607 B2*******************
    sSql = sSql & ", isnull(cTpoProdCod,0) cTpoProdCod,isnull((select cConsDescripcion from constante where nConsValor=cTpoProdCod and nConsCod=3033),'') cTpoProdDes "
    sSql = sSql & ", isnull(cTpoCredCod,0) cTpoCredCod,isnull((select cConsDescripcion from constante where nConsValor=cTpoCredCod and nConsCod=3034),'') cTpoCredDes "
    '***********************************
    'BRGO 20111227 *********************
    sSql = sSql & ", CE.dPrdEstado dFecAprobacion "
    '***********************************
    sSql = sSql & " From Producto P Inner join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
    sSql = sSql & "                 Inner Join Persona Pers ON PP.cPersCod = Pers.cPersCod "
    sSql = sSql & "                 Inner Join Colocaciones C ON C.cCtaCod =  P.cCtaCod "
    sSql = sSql & "                 Inner Join ColocacCred CC ON CC.cCtaCod =  P.cCtaCod "
    sSql = sSql & "                 Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod And CE.nPrdEstado = " & COMDConstantes.gColocEstAprob
    sSql = sSql & " WHERE P.cCtaCod = '" & psCtaCod & "'"

    If pbIncluirbAprobados Then
        sSql = sSql & " AND P.nPrdEstado in (" & COMDConstantes.gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & COMDConstantes.gColocEstRefVenc & "," & COMDConstantes.gColocEstVigMor & "," & COMDConstantes.gColocEstVigNorm & "," & COMDConstantes.gColocEstVigVenc & "," & COMDConstantes.gColocEstAprob & ")"
    Else
        sSql = sSql & " AND P.nPrdEstado in (" & COMDConstantes.gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & COMDConstantes.gColocEstRefVenc & "," & COMDConstantes.gColocEstVigMor & "," & COMDConstantes.gColocEstVigNorm & "," & COMDConstantes.gColocEstVigVenc & ")"
    End If

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosCreditoVigente = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaDatosCreditoVigente:
    Err.Raise Err.Number, "Error En Proceso RecuperaDatosCreditoVigente", Err.Description

End Function
'RECO FIN*****************************************************************************************************
'FRHU 20140926 ERS099-2014
Public Function CabeRepoCaptacSolicitud(ByVal sCabe01 As String, ByVal sCabe02 As String, _
        ByVal nCarLin As Long, ByVal sSeccion As String, ByVal sTitRp1 As String, _
        ByVal sTitRp2 As String, ByVal sMoneda As String, ByVal sNumPag As String, _
        ByVal sNomAge As String, ByVal dFecSis As Date, ByVal sSaltoLinea As String, _
        Optional ByVal sTitRp3 As String = "", Optional ByVal psUser As String = "") As String

Dim sTit1 As String, sTit2 As String
Dim sTit3 As String
Dim nCantCaracTit1 As Integer, nEspacioTit1 As Integer
Dim sMon As String
Dim sCad As String
sTit1 = "": sTit2 = ""

CabeRepoCaptacSolicitud = ""


' Definición de Cabecera 1
sMoneda = IIf(sMoneda = "", String(10, " "), " - " & sMoneda)
sCad = UCase(Trim(sNomAge)) & sMoneda
sCabe01 = sCad & String(50 - Len(sCad), " ")
sCabe01 = sCabe01 & Space((IIf(nCarLin <= 36, 80, nCarLin) - 36) - (Len(Mid(Trim(sCabe01), 1, 45)) - 2))
sCabe01 = sCabe01 & "PAGINA: " & sNumPag
sCabe01 = sCabe01 & Space(5) & "FECHA  : " & Format$(dFecSis, "dd/mm/yyyy")

' Definición de Cabecera 2
sCabe02 = sSeccion & String(19 - Len(sSeccion), " ")
sCabe02 = sCabe02 & Space((IIf(nCarLin <= 19, 100, nCarLin) - 19) - (Len(sCabe02) - 2))
sCabe02 = sCabe02 & Space(15) & "HORA   : " & Format$(Now(), "hh:mm:ss")

' Definición del Titulo del Reporte
If Len(Trim(sTitRp1)) = 21 Then
    nEspacioTit1 = 43
Else
    nEspacioTit1 = 29
End If
sTit1 = String(Int((IIf(nCarLin <= Len(sTitRp1), Len(sTitRp1) + 1, nCarLin) - Len(sTitRp1)) / 2), " ") & sTitRp1 & Space(nEspacioTit1) & "USUARIO: " & psUser
sTit2 = String(Int((IIf(nCarLin <= Len(sTitRp2), Len(sTitRp2) + 1, nCarLin) - Len(sTitRp2)) / 2), " ") & sTitRp2
sTit3 = String(Int((IIf(nCarLin <= Len(sTitRp3), Len(sTitRp3) + 1, nCarLin) - Len(sTitRp3)) / 2), " ") & sTitRp3
    
CabeRepoCaptacSolicitud = CabeRepoCaptacSolicitud & sCabe01 & sSaltoLinea
CabeRepoCaptacSolicitud = CabeRepoCaptacSolicitud & sCabe02 & sSaltoLinea
CabeRepoCaptacSolicitud = CabeRepoCaptacSolicitud & sTit1 & sSaltoLinea

CabeRepoCaptacSolicitud = CabeRepoCaptacSolicitud & sTit2 & sSaltoLinea
CabeRepoCaptacSolicitud = CabeRepoCaptacSolicitud & sTit3

End Function
'FIN FRHU 20140926
'JUEZ 20150224 **********************************************************
Public Function ImprimeConsultaMovimientoDebitoAuto(ByVal rs As ADODB.Recordset, ByVal rs1 As ADODB.Recordset, ByVal gsCodUser As String, _
        ByVal gsNomCmac As String, ByVal gsNomAge As String, ByVal sMoneda As String, ByVal gdFecSis As String, _
        ByVal sTitRp1 As String, ByVal sTitRp2 As String, ByVal psCuenta As String, ByVal psFecSaldIni As String, _
        ByVal pdSaldoInicial As String, ByVal psSaldoDisponible As String, ByVal psSaldoContable As String, ByVal psSaldoRet As String, _
        ByVal psInteresMes As String, ByVal sIntPag As String, ByVal psEstado As String, ByVal psTotCargo As String, _
        ByRef pnNroPag As Integer, Optional ByVal psTEA As String) As String
     
      
Dim sCad As String
Dim sCadAux As String
Dim nCarLin As Integer
Dim nLinPag As Integer, nEspaciosIzq As Integer
Dim sNumPag As String
Dim nCntPag As Integer
Dim sOperacion As String * 40
Dim sDocumento As String * 18
Dim sDeposito As String * 13
Dim sRetiro As String * 13
Dim sSaldoCnt As String * 13
Dim sAgencia  As String * 10
Dim sFecha As String * 19
Dim sUser As String * 4
Dim ofunc As New COMFunciones.FCOMCadenas
Dim oFunI As New COMFunciones.FCOMImpresion


nCarLin = 122
nLinPag = 65
sNumPag = ""
nCntPag = 0
nEspaciosIzq = 2

rs.Sort = "dFecha,sOpeCod"

If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    
    Do Until rs.EOF
        
        If nLinPag > 58 Then
            If sNumPag <> "" Then
                sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea
                sCad = sCad & Space(nEspaciosIzq) & gsCodUser
                sCad = sCad & oImp.gPrnSaltoPagina
                nLinPag = 0
            End If
                
            nCntPag = nCntPag + 1
            sNumPag = ofunc.FillNum(Trim(nCntPag), 4, " ")
            sCad = sCad & oFunI.CabeRepo(gsNomCmac, gsNomAge, "SECCION AHORROS", sMoneda, Format$(gdFecSis, "dd/mm/yyyy"), sTitRp1, sTitRp2, "", "", nCntPag, nCarLin, nEspaciosIzq, oImp) & oImp.gPrnSaltoLinea
            sCad = sCad & Space(nEspaciosIzq) & "Cuenta No. :   " & oImp.gPrnNegritaON & psCuenta & oImp.gPrnNegritaOFF & Space(5) & "TEA :   " & oImp.gPrnNegritaON & psTEA & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            sCad = sCad & Space(nEspaciosIzq) & "CLIENTE" & Space(50) & "RELACION" & Space(20) & "DIRECCION" & oImp.gPrnSaltoLinea
            nLinPag = 7
            If Not (rs1.EOF And rs1.BOF) Then
                rs1.MoveFirst
                Do Until rs1.EOF
                    nLinPag = nLinPag + 1
                    sCad = sCad & Space(nEspaciosIzq) & Left(Trim(rs1(0)), 59) & Space(60 - Len(Left(Trim(rs1(0)), 50))) & Trim(rs1(1)) & Space(10) & Trim(rs1(2)) & oImp.gPrnSaltoLinea
                    rs1.MoveNext
                Loop
            End If
            
            sCad = sCad & oImp.gPrnSaltoLinea
            sCad = sCad & Space(nEspaciosIzq) & "SALDO AL " & psFecSaldIni & " : " & ofunc.JDNum(pdSaldoInicial, 12, True, 7, 2) & oImp.gPrnSaltoLinea
            sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea
            sCad = sCad & Space(nEspaciosIzq) & "      FECHAS       "
            sCad = sCad & " " & "           TIPO DE OPERACION            "
            sCad = sCad & " " & "    DOCUMENTO     "
            sCad = sCad & " " & "      RETIROS"
            sCad = sCad & " " & "     CONTABLE"
            sCad = sCad & " " & " AGENCIA  "
            sCad = sCad & " " & "USUA" & oImp.gPrnSaltoLinea
            sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea            ' hasta 160 caracteres x línea
            nLinPag = nLinPag + 6
            If (nCntPag Mod 3) = 0 Then
                sCadAux = sCadAux & sCad
                sCad = ""
            End If
        End If
        
        sCad = sCad & Space(nEspaciosIzq)
        sFecha = Format$(rs(0), "dd/mm/yyyy hh:mm:ss")
        sOperacion = oFunI.ImpreCarEsp(rs(1))
        sDocumento = Trim(rs(2))
        RSet sRetiro = rs(3)
        RSet sSaldoCnt = rs(4)
        sAgencia = Trim(rs(5))
        sUser = Trim(rs(6))
        sCad = sCad & sFecha
        sCad = sCad & " " & sOperacion
        sCad = sCad & " " & sDocumento
        sCad = sCad & " " & sRetiro
        If Trim(sSaldoCnt) = "999999999" Then
            sCad = sCad & " " & Space(13)
        Else
            sCad = sCad & " " & sSaldoCnt
        End If
        sCad = sCad & " " & sAgencia
        sCad = sCad & " " & sUser & oImp.gPrnSaltoLinea
               
        nLinPag = nLinPag + 1
        rs.MoveNext
        
    Loop
End If

sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea              ' hasta 160 caracteres x línea
sCad = sCad & Space(nEspaciosIzq) & "  TOTAL    " & Space(75) & psTotCargo & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea

sCad = sCad & Space(nEspaciosIzq) & "SALDO DISPONIBLE  " & Space(75) & psSaldoDisponible & oImp.gPrnSaltoLinea
sCad = sCad & Space(nEspaciosIzq) & "SALDO CONTABLE    " & Space(75) & psSaldoContable & oImp.gPrnSaltoLinea

If Mid(psCuenta, 6, 3) = gCapCTS Then
    sCad = sCad & Space(nEspaciosIzq) & "SALDO RETIRO      " & Space(75) & psSaldoRet & oImp.gPrnSaltoLinea
ElseIf Mid(psCuenta, 6, 3) = gCapAhorros Then
    sCad = sCad & Space(nEspaciosIzq) & "BLOQUEO PARCIAL   " & Space(75) & psSaldoRet & oImp.gPrnSaltoLinea
End If
sCad = sCad & Space(nEspaciosIzq) & "INTERÉS DEL MES   " & Space(75) & psInteresMes & oImp.gPrnSaltoLinea

If Mid(psCuenta, 6, 3) = gCapPlazoFijo Then
    sCad = sCad & Space(nEspaciosIzq) & "INTERÉS PAGADO    " & Space(75) & sIntPag & oImp.gPrnSaltoLinea
End If
sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "=") & oImp.gPrnSaltoLinea             ' hasta 160 caracteres x línea

If Mid(psCuenta, 6, 3) = gCapCTS Then
    sCad = sCad & "  * Saldo total de retiro de todas las cuentas CTS del cliente com el mismo empleador" & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
End If

Dim lsMsg As String
lsMsg = "Advertencia: Si dentro de 30 días no se formula observación al presente estado, daremos por conforme la cuenta y aprobado el saldo. En caso contrario sírvase dirigirse a nuestras oficinas para atender sus observaciones." & oImp.gPrnSaltoLinea & "En caso de reclamos y servicios, el cliente podrá recurrir indistintamente a las siguientes instancias:" & oImp.gPrnSaltoLinea & "1. A  nuestra  Red  de  Oficinas.  2.  A  la  Plataforma de Atención al Usuario de la Superintendencia de Banca y Seguros 3. INDECOPI."
sCad = sCad & JustificaTextoSinCarEsp(lsMsg, nCarLin, nEspaciosIzq) & oImp.gPrnSaltoLinea
sCad = sCad & Space(nEspaciosIzq) & psEstado & oImp.gPrnSaltoLinea
sCad = sCad & Space(nEspaciosIzq) & String(nCarLin, "-") & oImp.gPrnSaltoLinea
sCad = sCad & Space(nEspaciosIzq) & gsCodUser

sCadAux = sCadAux & sCad
pnNroPag = nCntPag
Set oImp = Nothing
ImprimeConsultaMovimientoDebitoAuto = sCadAux
End Function
'END JUEZ ***************************************************************

'***AMDO20150326 - TI-ERS018-2015 *************************************************************
Public Function ImprimirAfiliaTarjetaCuentas(ByVal psNombreCMAC As String, ByVal psNomAge As String, _
                                           ByVal psFecha As String, ByVal psHora As String, _
                                           ByVal psTitular As String, ByVal psDOI As String, _
                                           ByVal psUsurio As String, ByRef psLpt As String, _
                                           ByRef pbImpTMU As Boolean, ByVal psTarjeta As String, ByRef prsCuentas As Recordset) As String
                                           
                                           
    Dim sSep As Integer
    Dim sIni As Integer
    Dim sMax As Integer
    Dim saux As Integer
    Dim sCad As String
    Dim lnTope As Integer
    Dim oFCOMImpresion As New COMFunciones.FCOMImpresion
    Dim lscad As String
    Dim lnNomConSerPag As String * 20


ETIQ:

    On Error GoTo Error

    lnTope = 0 '6 'Tope de lineas en Boleta

    lscad = lscad & oImp.gPrnInicializa
    lscad = lscad & oImp.gPrnEspaLineaN   'espaciamiento lineas 1/6 pulg.
    lscad = lscad & oImp.gPrnTamPagina22   'Longitud de página a 22 líneas'
    lscad = lscad & oImp.gPrnTamLetra10CPI   'Tamaño 10 cpi
    lscad = lscad & oImp.gPrnTpoLetraRoman     'Tipo de Letra Sans Serif
    lscad = lscad & oImp.gPrnCondensadaOFF ' cancela condensada
    lscad = lscad & oImp.gPrnNegritaOFF ' desactiva negrita

    sSep = 19
    sIni = 0
    sMax = 33
    saux = 17
    
'    lnNomConSerPag = psNombreConvenio
    
    If pbImpTMU = False Then
        lscad = lscad & Space(sIni) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psNombreCMAC & Space(sSep + sMax - Len(Trim(psNombreCMAC)) + 2) & psNombreCMAC & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & UCase(psNomAge) & Space(37) & UCase(psNomAge) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "FECHA: " & UCase(psFecha) & Space(7) & "HORA: " & CStr(psHora) & Space(saux + sMax - Len(Trim(psFecha)) - 24) & "FECHA: " & UCase(psFecha) & Space(7) & "HORA: " & CStr(psHora) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psTitular & Space(saux + sMax - Len(Trim(psTitular)) + 4) & psTitular & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Tarjeta: " & psTarjeta & Space(saux + sMax - 25 + 4) & "Tarjeta: " & psTarjeta & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        '''''''''''''''''''''''''''''''--------------------------------------

            lscad = lscad & Space(sIni) & "-ACTUALIZACION DE CUENTAS VINCULADAS--" & Space(16) & "-ACTUALIZACION DE CUENTAS VINCULADAS--" & oImp.gPrnSaltoLinea
            lscad = lscad & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "CUENTAS:" & Space(46) & "CUENTAS:" & oImp.gPrnSaltoLinea
                        
                Do While Not prsCuentas.EOF

                If prsCuentas(7) = 1 Then
                    '''lscad = lscad & Space(sIni) & prsCuentas(0) & " " & prsCuentas(1) & " " & IIf(Mid(prsCuentas(0), 9, 1) = "1", "SOLES", "DOLARES") & Space(saux + sMax - Len(Trim(prsCuentas(0) & prsCuentas(1) & IIf(Mid(prsCuentas(0), 9, 1) = "1", "SOLES", "DOLARES"))) + 2) & prsCuentas(0) & " " & prsCuentas(1) & " " & IIf(Mid(prsCuentas(0), 9, 1) = "1", "SOLES", "DOLARES") & oImp.gPrnSaltoLinea 'marg ers044-2016
                    lscad = lscad & Space(sIni) & prsCuentas(0) & " " & prsCuentas(1) & " " & IIf(Mid(prsCuentas(0), 9, 1) = "1", StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), "DOLARES") & Space(saux + sMax - Len(Trim(prsCuentas(0) & prsCuentas(1) & IIf(Mid(prsCuentas(0), 9, 1) = "1", StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), "DOLARES"))) + 2) & prsCuentas(0) & " " & prsCuentas(1) & " " & IIf(Mid(prsCuentas(0), 9, 1) = "1", StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), "DOLARES") & oImp.gPrnSaltoLinea 'marg ers044-2016
                    lscad = lscad & Space(sIni) & "PRIORIDAD: " & CStr(CInt(prsCuentas(6))) & " " & Replace(prsCuentas(5), "Ñ", Chr(165), , , vbTextCompare) & Space(saux + sMax - (Len(Trim(CStr(CInt(prsCuentas(6))) & prsCuentas(5))) + 9) + 1) & "PRIORIDAD: " & CStr(CInt(prsCuentas(6))) & " " & Replace(prsCuentas(5), "Ñ", Chr(165), , , vbTextCompare) & oImp.gPrnSaltoLinea
                    lscad = lscad & oImp.gPrnSaltoLinea
                End If

                    prsCuentas.MoveNext
                Loop
                lscad = lscad & Space(sIni) & "--------------------------------------" & Space(16) & "--------------------------------------" & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "Queda sin efecto las vinculaciones de" & Space(17) & "Queda sin efecto las vinculaciones de" & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "cuentas a mi tarjeta, realizados antes" & Space(16) & "cuentas a mi tarjeta, realizados antes" & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "de esta fecha." & Space(41) & "de esta fecha." & oImp.gPrnSaltoLinea
                lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
                
                lscad = lscad & Space(sIni) & "FIRMA : ____________________" & Space(26) & "FIRMA : ____________________" & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "NOMBRE: " & psTitular & Space(saux + sMax - (Len(Trim(psTitular)) + 8) + 4) & "NOMBRE: " & psTitular & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "DOI   : " & psDOI & Space(saux + sMax - (Len(Trim(psDOI)) + 8) + 4) & "DOI   : " & psDOI & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "--------------------------------------" & Space(16) & "--------------------------------------" & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & psUsurio & Space(50) & psUsurio & oImp.gPrnSaltoLinea

            lscad = lscad & oImp.gPrnSaltoLinea
            lscad = lscad & oImp.gPrnSaltoLinea

        lscad = lscad & oFCOMImpresion.ImpreCarEsp(sCad)
        ImprimirAfiliaTarjetaCuentas = lscad & oImp.gPrnSaltoLinea
    Else 'TMU
    
        lscad = lscad & Space(sIni) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psNombreCMAC & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & UCase(psNomAge) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "FECHA: " & UCase(psFecha) & Space(7) & "HORA: " & CStr(psHora) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psTitular & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Tarjeta: " & psTarjeta & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        '''''''''''''''''''''''''''''''--------------------------------------
        '''''''''''''''''''''''''''''''01234567890123456789012345678901234567

            lscad = lscad & Space(sIni) & "-ACTUALIZACION DE CUENTAS VINCULADAS--" & oImp.gPrnSaltoLinea
            lscad = lscad & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "CUENTAS:" & oImp.gPrnSaltoLinea

                Do While Not prsCuentas.EOF
                If prsCuentas(7) = 1 Then
                    '''lscad = lscad & Space(sIni) & prsCuentas(0) & " " & prsCuentas(1) & " " & IIf(Mid(prsCuentas(0), 9, 1) = "1", "SOLES", "DOLARES") & oImp.gPrnSaltoLinea 'marg ers044-2016
                    lscad = lscad & Space(sIni) & prsCuentas(0) & " " & prsCuentas(1) & " " & IIf(Mid(prsCuentas(0), 9, 1) = "1", StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), "DOLARES") & oImp.gPrnSaltoLinea 'marg ers044-2016
                    lscad = lscad & Space(sIni) & "PRIORIDAD: " & CStr(CInt(prsCuentas(6))) & " " & Replace(prsCuentas(5), "Ñ", Chr(165), , , vbTextCompare) & oImp.gPrnSaltoLinea
                    lscad = lscad & oImp.gPrnSaltoLinea
                End If
                
                    prsCuentas.MoveNext
                Loop
                lscad = lscad & Space(sIni) & "--------------------------------------" & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "Queda sin efecto las vinculaciones de" & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "cuentas a mi tarjeta, realizados antes" & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "de esta fecha." & oImp.gPrnSaltoLinea
                lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
                
                lscad = lscad & Space(sIni) & "FIRMA : ____________________" & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "NOMBRE: " & psTitular & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "DOI   : " & psDOI & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "--------------------------------------" & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & psUsurio & oImp.gPrnSaltoLinea
                
            lscad = lscad & oImp.gPrnSaltoLinea
            lscad = lscad & oImp.gPrnSaltoLinea

        lscad = lscad & oFCOMImpresion.ImpreCarEsp(sCad)
        ImprimirAfiliaTarjetaCuentas = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)

    End If
    
Exit Function
Error:
    Err.Raise Err.Number, "Comprueba la conexion de su impresora... ", Err.Description
'    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
'        GoTo ETIQ
'    End If
End Function

'END AMDO**********************************************************************************************

Public Function ImprimeBoletaNew(ByVal psTit As String, ByVal pstext As String, _
                                 ByVal psCodOpe As String, ByVal pnMonto As String, _
                                 ByVal psCliente As String, ByVal psCodCta As String, _
                                 ByVal pnNumExt As Long, Optional pbSaldoInt As Boolean = True, _
                                 Optional pbCuenta As Boolean = False, Optional psTexto As String = "XXX", _
                                 Optional pdFecSis As Date, Optional sNomAge As String = "", _
                                 Optional sCodUser As String = "", Optional pnMonedaCMAC As Integer = 1, _
                                 Optional psDisponible As String = "", Optional psIntangible As String = "", _
                                 Optional psDispCta As String = "", Optional psIntaCta As String = "", _
                                 Optional pbImpTMU As Boolean = False, Optional pcPersNomInst As String = "", _
                                 Optional ByVal psTipoCambio As String = "XXX", Optional ByVal psDOI As String = "") As String

    Dim sFecha As String
    Dim sHora As String
    Dim sSep As Integer
    Dim sIni As Integer
    Dim sMonto As String
    Dim sMax As Integer
    Dim saux As Integer
    Dim lsNroExt As String
    Dim sCad As String
    Dim lnCliAux As Long
    Dim lsCliAux1 As String
    Dim lsCliAux2 As String
    Dim lnChq As Long
    Dim lsChqAux1 As String
    Dim lsChqAux2 As String
    Dim oVarImpre As New COMFunciones.FCOMVarImpresion
    '** Para la Impresion del Codigo Antiguo en las Boletas
    Dim lsCodAntiguo As String
    Dim lscad As String
    Dim pcImrimePer As String

    Dim loConstSistema As COMDConstSistema.NCOMConstSistema
    Set loConstSistema = New COMDConstSistema.NCOMConstSistema
    If loConstSistema.LeeConstSistema(122) = "1" Then
    Dim loGeneral As COMDConstSistema.DCOMGeneral
    Set loGeneral = New COMDConstSistema.DCOMGeneral
    lsCodAntiguo = loGeneral.GetCuentaAntigua(psCodCta)
    Else
        lsCodAntiguo = ""
    End If
    Set loConstSistema = Nothing
    
    'RIRO20160318 *********
    Dim nI As Integer
    Dim sComent As String
    'END RIRO *************
    
ETIQ:

    On Error GoTo Error

    lsNroExt = Str(pnNumExt)

    lscad = lscad & oImp.gPrnInicializa
    lscad = lscad & oImp.gPrnEspaLineaN   'espaciamiento lineas 1/6 pulg.
    lscad = lscad & oImp.gPrnTamPagina22  'Longitud de página a 22 líneas'
    lscad = lscad & oImp.gPrnTamLetra10CPI   'Tamaño 10 cpi
    lscad = lscad & oImp.gPrnTpoLetraRoman     'Tipo de Letra Sans Serif
    lscad = lscad & oImp.gPrnCondensadaOFF ' cancela condensada
    lscad = lscad & oImp.gPrnNegritaOFF ' desactiva negrita

    sSep = 15
    sIni = 0
    sMax = 33
    saux = 5

    sFecha = Format$(pdFecSis, "dd/mm/yyyy")
    sHora = Format$(Time, "hh:mm:ss")
    sMonto = Format$(pnMonto, "#,##0.00")

    '' Declaracion para los mensajes
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Dim oImpre As New COMFunciones.FCOMImpresion

    '' liquidacion Trujillo
    If pbImpTMU = False Then
        lscad = lscad & oImp.gPrnNegritaON  'Activa Negrita
        lscad = lscad & oImp.gPrnSaltoLinea

        lscad = lscad & Space(sIni) & "CMAC-MAYNAS - AHORRO" & Space(16 + sSep + saux) & "CMAC-MAYNAS - AHORRO" & oImp.gPrnSaltoLinea
        If Len(Trim(psCodCta)) < 18 Then
            If pnMonedaCMAC = 1 Then
                '''lscad = lscad & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
            Else
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & oImp.gPrnSaltoLinea
            End If
        Else
            If Mid(psCodCta, 9, 1) = 1 Then
                '''lscad = lscad & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
            Else
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & oImp.gPrnSaltoLinea
            End If
        End If
        lscad = lscad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea
        lnCliAux = InStr(1, pcPersNomInst, "*", vbTextCompare)
        If lnCliAux = 0 Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & oImpre.ImpreCarEsp(psCliente) & oImp.gPrnSaltoLinea
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & Space(saux + sMax + sSep - Len(lsCliAux1)) & oImpre.ImpreCarEsp(lsCliAux1) & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & Space(saux + sMax + sSep - Len(lsCliAux2)) & oImpre.ImpreCarEsp(lsCliAux2) & oImp.gPrnSaltoLinea
            lnCliAux = 1
        End If
        If pbSaldoInt Or pbCuenta Then
            If lsCodAntiguo <> "" Then
                lscad = lscad & Space(sIni) & "Cuenta:" & lsCodAntiguo & Space(8 + sSep + saux) & "Cuenta:" & lsCodAntiguo & oImp.gPrnSaltoLinea
            Else
                lscad = lscad & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + saux) & "Cuenta:" & psCodCta & oImp.gPrnSaltoLinea
            End If
        Else
            lscad = lscad & ""
        End If
        
        If Len(Trim(psDOI)) > 0 Then
            lscad = lscad & Space(sIni) & "DOI: " & psDOI & Space(16 + sSep + saux) & "DOI: " & psDOI & oImp.gPrnSaltoLinea
        End If
        
        psTit = Trim(psTit)
        psTit = oImpre.CentrarCadena(psTit, 28)
        lscad = lscad & oImp.gPrnNegritaON 'Activa Negrita
        lscad = lscad & Space(sIni + 1) & "-----" & psTit & "-----" & Space(-1 + sSep) & "-----" & psTit & "-----" & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnNegritaOFF
        lnChq = InStr(1, pstext, "*", vbTextCompare)
                
        If psTexto = "XXX" Then
            If lnChq = 0 Then
                
                'RIRO20160318 **************

                nI = 1
                If gAhoBloqDesbCuenta = psCodOpe Then ' Si es Bloqueo/Desbloqueo Cuentas
                    Do While Len(Trim(Mid(pstext, (nI - 1) * (sMax + 5) + 1, sMax + 5))) > 0
                        sComent = Trim(Mid(pstext, (nI - 1) * (sMax + 5) + 1, sMax + 5))
                        lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(Mid(sComent, 1, sMax + 5))) & Space((sMax + 5) + 1 - Len(Trim(Mid(sComent, 1, (sMax + 5)))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(Mid(sComent, 1, (sMax + 5)))) & Space((sMax + 5) + 1 - Len(Trim(Mid(pstext, 1, (sMax + 5)))) - Len(sMonto)) & sMonto & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                        nI = nI + 1
                    Loop
                Else
                    lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                End If
                lscad = lscad & ""
                'RIRO20160318 *************
            Else
                lsChqAux1 = (Mid(pstext, 1, lnChq - 1))
                lsChqAux2 = (Mid(pstext, lnChq + 1))
                lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            End If
        Else
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & Space(saux + sSep + sMax - Len(Trim(psTexto))) & oImpre.ImpreCarEsp(Trim(psTexto)) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        End If
        If psDisponible <> "" And psIntangible <> "" Then
           lscad = lscad & oVarImpre.gPrnCondensadaON & Space(sIni) & oImpre.ImpreCarEsp(Trim(psDispCta)) & Space(saux + sSep + sMax - Len(Trim(psDispCta))) & oImpre.ImpreCarEsp(Trim(psDispCta)) & oImp.gPrnSaltoLinea
           lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psIntaCta)) & Space(saux + sSep + sMax - Len(Trim(psIntaCta))) & oImpre.ImpreCarEsp(Trim(psIntaCta)) & oImp.gPrnSaltoLinea
           lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psDisponible)) & Space(saux + sSep + sMax - Len(Trim(psDisponible))) & oImpre.ImpreCarEsp(Trim(psDisponible)) & oImp.gPrnSaltoLinea
           lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psIntangible)) & Space(saux + sSep + sMax - Len(Trim(psIntangible))) & oImpre.ImpreCarEsp(Trim(psIntangible)) & oVarImpre.gPrnCondensadaOFF & oImp.gPrnSaltoLinea
        End If
        lscad = lscad & oImp.gPrnNegritaON & oImp.gPrnSaltoLinea & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & "---------------------------------------" & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        If psTipoCambio = "XXX" Then
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Space(29 + sSep + saux) & oImpre.ImpreCarEsp(sCodUser) & oImp.gPrnSaltoLinea
        Else
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Space(5) & psTipoCambio & Space(sSep + saux - 2) & oImpre.ImpreCarEsp(sCodUser) & Space(5) & psTipoCambio & oImp.gPrnSaltoLinea
        End If
        ImprimeBoletaNew = lscad & oImp.gPrnSaltoLinea
    Else 'TMU
        
        lscad = lscad & oImp.gPrnNegritaON  'Activa Negrita
        If psCodOpe <> gOtrOpeDesParGas And psCodOpe <> gOtrOpeDesParVia And psCodOpe <> gOtrOpeDesParCaj Then
            lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        End If
        lscad = lscad & Space(sIni) & "CMAC-MAYNAS - AHORRO" & Space(16 + sSep + saux) & oImp.gPrnSaltoLinea
        If Len(Trim(psCodCta)) < 18 Then
            If pnMonedaCMAC = 1 Then
                '''lscad = lscad & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
            Else
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & oImp.gPrnSaltoLinea
            End If
        Else
            If Mid(psCodCta, 9, 1) = 1 Then
                '''lscad = lscad & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - 3 - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - 3 - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
            Else
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - 3 - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & oImp.gPrnSaltoLinea
            End If
        End If
        lscad = lscad & Space(sIni) & "Fecha:" & sFecha & Space(6) & "Hora:" & sHora & Space(saux + sSep - 6) & oImp.gPrnSaltoLinea
        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)

        If lnCliAux = 0 Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Space(saux + sMax - 3 + sSep - Len(psCliente)) & oImp.gPrnSaltoLinea
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
    
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
            
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & oImpre.ImpreCarEsp(lsCliAux1) & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & oImpre.ImpreCarEsp(lsCliAux2) & oImp.gPrnSaltoLinea

            lnCliAux = 1
        End If
        
        If pbSaldoInt Or pbCuenta Then
            If lsCodAntiguo <> "" Then  ' Para Codigo Antiguo - LAYG
                lscad = lscad & Space(sIni) & "Cuenta:" & lsCodAntiguo & oImp.gPrnSaltoLinea
            Else
                lscad = lscad & Space(sIni) & "Cuenta:" & psCodCta & oImp.gPrnSaltoLinea
            End If
        Else
            lscad = lscad & ""
        End If
        
        If Len(Trim(psDOI)) > 0 Then
            lscad = lscad & Space(sIni) & "DOI: " & psDOI & oImp.gPrnSaltoLinea
        End If
        
        psTit = Trim(psTit)
        psTit = oImpre.CentrarCadena(psTit, 25)
        lscad = lscad & oImp.gPrnNegritaON 'Activa Negrita
        lscad = lscad & Space(sIni + 1) & "-----" & psTit & "-----" & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnNegritaOFF
        lnChq = InStr(1, pstext, "*", vbTextCompare)
        
        If psTexto = "XXX" Then
            If lnChq = 0 Then
            
                'RIRO20160318 **************
                
                nI = 1
                If gAhoBloqDesbCuenta = psCodOpe Then ' Si es Bloqueo/Desbloqueo Cuentas
                    Do While Len(Trim(Mid(pstext, (nI - 1) * (sMax + 2) + 1, sMax + 2))) > 0
                        sComent = Trim(Mid(pstext, (nI - 1) * (sMax + 2) + 1, sMax + 2))
                        
                        'lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(Mid(sComent, 1, sMax + 5))) & Space((sMax + 5) + 1 - Len(Trim(Mid(sComent, 1, (sMax + 5)))) - Len(sMonto)) & sMonto & oImp.gPrnSaltoLinea
                        lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(sComent, 1, sMax + 2))) & Space((sMax + 2) + 1 - Len(Trim(Mid(sComent, 1, (sMax + 2)))) + Len(Format(sMonto, "0.00"))) & Format(sMonto, "0.00") & oImp.gPrnSaltoLinea
                        
                        nI = nI + 1
                    Loop
                Else
                    'lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(Mid(psText, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(Mid(psText, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)) & sMonto & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                    lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(sComent, 1, sMax + 2))) & Space((sMax + 2) + 1 - Len(Trim(Mid(sComent, 1, (sMax + 2)))) + Len(Format(sMonto, "0.00"))) & Format(sMonto, "0.00") & oImp.gPrnSaltoLinea
                End If
                lscad = lscad & ""
                'RIRO20160318 *************
            
            
            
                
                lscad = lscad & ""
            Else
                lsChqAux1 = (Mid(pstext, 1, lnChq - 1))
                lsChqAux2 = (Mid(pstext, lnChq + 1))
                lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Space(-1 + sSep) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            End If
        Else
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            If psDisponible = "" And psIntangible = "" Then
                lscad = lscad & ""
            End If
        End If
        
        If psDisponible <> "" And psIntangible <> "" Then

           lscad = lscad & oVarImpre.gPrnCondensadaON & Space(sIni) & oImpre.ImpreCarEsp(Trim(psDispCta)) & oImp.gPrnSaltoLinea
           lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psIntaCta)) & oImp.gPrnSaltoLinea
           lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psDisponible)) & oImp.gPrnSaltoLinea
           lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psIntangible)) & Space(saux + sSep + sMax - Len(Trim(psIntangible))) & oImpre.ImpreCarEsp(Trim(psIntangible)) & oVarImpre.gPrnCondensadaOFF & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea

        End If

        lscad = lscad & oImp.gPrnNegritaON & Space(sIni) & oImp.gPrnSaltoLinea & "-----------------------------------" & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
        ImprimeBoletaNew = lscad
    End If
Exit Function
Error:
    Err.Raise Err.Number, "Comprueba la conexion de su impresora... ", Err.Description
End Function
'END RIRO ***************************************

'GIPO ERS070-2016  13-03-2017
Public Function imprimirBoletaDevolucionSobrantesPignoraticio(ByVal psTitCab As String, ByVal psTitDet As String, _
                ByVal psNomPers As String, ByVal psDirPers As String, _
                ByVal psDocumento As String, ByVal pnTipoCambio As Double, _
                ByVal psCodOpe As String, ByVal pnTotal As Currency, ByVal ITF As Currency, _
                ByVal prsDebitos As ADODB.Recordset, _
                ByVal psNomAge As String, ByVal psMovNro As String, ByVal psLpt As String, _
                Optional ByVal psCodCMAC As String, Optional ByVal psNomCmac As String, _
                Optional ByVal pbImpTMU As Boolean = False) As String
    Dim oImpre As COMFunciones.FCOMImpresion
    Set oImpre = New COMFunciones.FCOMImpresion
    Dim oConstPEN As New COMDConstantes.DCOMConstantes
    Dim sCad As String
    Dim nFicSal As Integer
    Dim lsFecha As String
    Dim lsHora As String
    Dim lsUserOrig As String
    Dim sBoldON As String
    Dim sBoldOFF As String
    Dim lsMensaje As String * 39
    Dim i As Integer
    Dim lnNumLineas As Integer
    Dim lnLineasBol As Integer
    Dim lnEspaceIni As Integer
    Dim lnAnchoBol As Long
    Dim psSubTit As String
    Dim psSubTit2 As String
    Dim sAgencia As String * 25, sNumMov As String * 25
    Dim lnMovNro As Long
    Dim oMov As COMDMov.DCOMMov
    Set oMov = New COMDMov.DCOMMov
    lnMovNro = oMov.GetnMovNro(psMovNro)
    Set oMov = Nothing
    psNomPers = Replace(Replace(psNomPers, "/", " "), "\", " ")
ETIQ:
On Error GoTo Error
    sBoldON = Chr$(27) + Chr$(71)
    sBoldOFF = Chr$(27) + Chr$(72)
    
    sCad = sCad & Chr$(27) & Chr$(64)
    sCad = sCad & Chr$(27) + Chr$(50)   'espaciamiento lineas 1/6 pulg.
    sCad = sCad & Chr$(27) + Chr$(67) + Chr$(22)  'Longitud de página a 22 líneas'
    sCad = sCad & Chr$(27) + Chr$(77)   'Tamaño 10 cpi
    sCad = sCad & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
    sCad = sCad & Chr$(27) + Chr$(18) ' cancela condensada
    sCad = sCad & Chr$(27) + Chr$(72) ' desactiva negrita
    
    psTitDet = "Total Devuelto " & oConstPEN.gcPEN_SIMBOLO()
    psSubTit = "Crédito 109017051000403083"
    psSubTit2 = "Crédito 109017051000365793"

    lsFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
    lsHora = Format$(Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2), "hh:mm:ss")
    lsUserOrig = Right(psMovNro, 4)
    lnEspaceIni = 2
    lnLineasBol = 22
    lnAnchoBol = 45


'    Global Const gcPEN_SINGULAR = "SOL"
'    Global Const gcPEN_PLURAL = "SOLES"
'    Global Const gcPEN_SIMBOLO = "S/"

If psCodCMAC = "102" Then
    sCad = sCad & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)
    sCad = sCad & oImpre.ImpreCarEsp(sBoldON & oImpre.ImpreFormat(psNomCmac & "- OPERACIONES", lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(oImpre.ImpreFormat(psNomAge, 30) & oImpre.ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10) '& oimpre.ImpreFormat(oimpre.ImpreFormat(psNomAge, 30) & oimpre.ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10)
    sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(oImpre.ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni)) & Chr$(10) ' & oimpre.ImpreFormat(oimpre.ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni)) & Chr$(10) ' & oimpre.ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat("Dir:" & psDirPers, lnAnchoBol, lnEspaceIni)) & Chr$(10) ' & oimpre.ImpreFormat("Dir:" & psDirPers, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    If psDocumento <> "" Then
        sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat("Nro Doc:" & psDocumento, lnAnchoBol, lnEspaceIni)) & Chr$(10) ' & oimpre.ImpreFormat("Nro Doc:" + psDocumento, lnAnchoBol, lnEspaceIni)) & Chr$(10)
    Else
        sCad = sCad & Chr$(10)
    End If
    sCad = sCad & oImpre.ImpreCarEsp(sBoldON & Space(lnEspaceIni) + oImpre.CentrarCadena(psTitCab, lnAnchoBol - 3)) & Chr$(10) ' & oimpre.CentrarCadena(psTitCab, lnAnchoBol - 3, lnEspaceIni + 4)) & Chr$(10)
    sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni)) & Chr$(10) ' & oimpre.ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(oImpre.ImpreFormat(psTitDet, lnAnchoBol / 2) & oImpre.ImpreFormat(CCur(pnTotal), 15, 4, True), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10) ' & oimpre.ImpreFormat(oimpre.ImpreFormat(psTitDet, lnAnchoBol / 2) & oimpre.ImpreFormat(CCur(pnTotal), 15, 2, True), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10)
    sCad = sCad & Chr$(10)

    sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(oImpre.ImpreFormat(psSubTit, lnAnchoBol / 2) & oImpre.ImpreFormat(CCur(ITF), 15, 4, True), lnAnchoBol, lnEspaceIni)) & Chr$(10) ' & oimpre.ImpreFormat(oimpre.ImpreFormat(psSubTit, lnAnchoBol / 2) & oimpre.ImpreFormat(CCur(ITF), 15, 2, True), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(oImpre.ImpreFormat(psSubTit2, lnAnchoBol / 2) & oImpre.ImpreFormat(CCur(pnTipoCambio), 15, 4, True), lnAnchoBol, lnEspaceIni)) & Chr$(10) '& oimpre.ImpreFormat(oimpre.ImpreFormat(psSubTit2, lnAnchoBol / 2) & oimpre.ImpreFormat(CCur(pnTipoCambio), 15, 4, True), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni)) & Chr$(10) '& oimpre.ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni)) & Chr$(10) '' & oimpre.ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni)) & Chr$(10)
Else
    If pbImpTMU = False Then
        sCad = sCad & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)
        sCad = sCad & oImpre.ImpreCarEsp(sBoldON & oImpre.ImpreFormat(psNomCmac & " - OPERACIONES", lnAnchoBol, lnEspaceIni) & oImpre.ImpreFormat(psNomCmac & "  - OPERACIONES", lnAnchoBol, lnEspaceIni)) & Chr$(10)
        sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(oImpre.ImpreFormat(psNomAge, 30) & oImpre.ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & oImpre.ImpreFormat(oImpre.ImpreFormat(psNomAge, 30) & oImpre.ImpreFormat(lnMovNro, 10, 0, False), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10)
        sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(oImpre.ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni) & oImpre.ImpreFormat(oImpre.ImpreFormat("Fecha:" & lsFecha, 25) & "Hora:" & lsHora, lnAnchoBol, lnEspaceIni)) & Chr$(10)
        sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni) & oImpre.ImpreFormat(psNomPers, lnAnchoBol, lnEspaceIni)) & Chr$(10)
        sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat("Dir:" & psDirPers, lnAnchoBol, lnEspaceIni) & oImpre.ImpreFormat("Dir:" & psDirPers, lnAnchoBol, lnEspaceIni)) & Chr$(10)
        If psDocumento <> "" Then
            sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat("Nro Doc:" & psDocumento, lnAnchoBol, lnEspaceIni) & oImpre.ImpreFormat("Nro Doc:" + psDocumento, lnAnchoBol, lnEspaceIni)) & Chr$(10)
        Else
            sCad = sCad & Chr$(10)
        End If
        sCad = sCad & oImpre.ImpreCarEsp(sBoldON & Space(lnEspaceIni) + oImpre.CentrarCadena(psTitCab, lnAnchoBol - 3) & oImpre.CentrarCadena(psTitCab, lnAnchoBol - 3, lnEspaceIni + 4)) & Chr$(10)
        sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni) & oImpre.ImpreFormat(String(lnAnchoBol - 3, "="), lnAnchoBol, lnEspaceIni)) & Chr$(10)
        
        prsDebitos.MoveFirst
        Do While Not prsDebitos.EOF
            sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(oImpre.ImpreFormat("Crédito " & prsDebitos!NroContrato, lnAnchoBol / 2) & oImpre.ImpreFormat(CCur(prsDebitos!Importe), 15, 2, True), lnAnchoBol, lnEspaceIni) & oImpre.ImpreFormat(oImpre.ImpreFormat("Crédito " & prsDebitos!NroContrato, lnAnchoBol / 2) & oImpre.ImpreFormat(CCur(prsDebitos!Importe), 15, 2, True), lnAnchoBol, lnEspaceIni)) & Chr$(10)
            prsDebitos.MoveNext
        Loop
        sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(oImpre.ImpreFormat("I.T.F (-)", lnAnchoBol / 2) & oImpre.ImpreFormat(CCur(ITF), 15, 2, True), lnAnchoBol, lnEspaceIni) & oImpre.ImpreFormat(oImpre.ImpreFormat("I.T.F. (-)", lnAnchoBol / 2) & oImpre.ImpreFormat(CCur(ITF), 15, 2, True), lnAnchoBol, lnEspaceIni)) & Chr$(10)   'para el ITF
        sCad = sCad & Chr$(10)
        sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(oImpre.ImpreFormat(psTitDet, lnAnchoBol / 2) & oImpre.ImpreFormat(CCur(pnTotal), 15, 2, True), lnAnchoBol, lnEspaceIni) & oImpre.ImpreFormat(oImpre.ImpreFormat(psTitDet, lnAnchoBol / 2) & oImpre.ImpreFormat(CCur(pnTotal), 15, 2, True), lnAnchoBol, lnEspaceIni) & sBoldOFF) & Chr$(10)
        sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni) & oImpre.ImpreFormat(String(lnAnchoBol - 3, "-"), lnAnchoBol, lnEspaceIni)) & Chr$(10)
        sCad = sCad & oImpre.ImpreCarEsp(oImpre.ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni) & oImpre.ImpreFormat(Right(psMovNro, 4), lnAnchoBol, lnEspaceIni)) & Chr$(10)
    Else
        sCad = sCad & Chr$(10)
        sCad = sCad & sBoldON & "C" & psNomCmac & " - OPERACIONES" & Chr$(10)
        sCad = sCad & psNomAge & Space(8) & lnMovNro & sBoldOFF & Chr$(10)
        sCad = sCad & "Fecha: " & lsFecha & Space(8) & "Hora: " & lsHora & Chr$(10)
        sCad = sCad & psNomPers & Chr$(10)
        sCad = sCad & "Dir: " & psDirPers & Chr$(10)
        If psDocumento <> "" Then
            sCad = sCad & "Nro Doc: " & psDocumento & Chr$(10)
        Else
            sCad = sCad & Chr$(10)
        End If
        psTitCab = "DEVOLUCIÓN DE SOBRANTES - CRED. PIGN"
        sCad = sCad & sBoldON & Space(5) & psTitCab & sBoldOFF & Chr$(10)
        sCad = sCad & String(38, "=") & Chr$(10)
'        scad = scad & psTitDet & Space(10) & Format(pnTotal, "###,###.0000") & Chr$(10) 'Sub título
        Dim sTotal As String
        Dim sImporte As String
        Dim sDetalle As String
        Dim sITF As String
        Dim c As Integer
        Dim x As Integer
       
        prsDebitos.MoveFirst
        Do While Not prsDebitos.EOF
            sDetalle = "Cred. " & prsDebitos!NroContrato
            sImporte = Format(prsDebitos!Importe, "###,###.00")
            c = 38 - Len(sDetalle)
            x = c - Len(sImporte)
            sCad = sCad & sDetalle & Space(x) & sImporte & Chr$(10)
            prsDebitos.MoveNext
        Loop
        sITF = "I.T.F.(-)"
        c = 38 - Len(sITF)
        x = c - Len(Format(ITF, "###,##0.00"))
        sCad = sCad & sITF & Space(x) & Format(ITF, "###,##0.00") & Chr$(10)
        sCad = sCad & Chr$(10)
        sTotal = Format(pnTotal, "###,###.00")
        c = 38 - Len(psTitDet)
        x = c - Len(sTotal)
        sCad = sCad & psTitDet & Space(x) & sTotal & Chr$(10) 'Sub título
        
        sCad = sCad & String(38, "-") & Chr$(10)
        sCad = sCad & Right(psMovNro, 4) & Chr$(10)
    End If

End If
Dim oGen As COMDConstSistema.DCOMGeneral
Set oGen = New COMDConstSistema.DCOMGeneral
lsMensaje = oGen.GetMensajeBoletas("")
Set oGen = Nothing
If pbImpTMU = False Then
    sCad = sCad & sBoldON & oImpre.ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & oImpre.ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & sBoldOFF & Chr$(10)
    sCad = sCad & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)
Else
    sCad = sCad & sBoldON & oImpre.ImpreFormat(lsMensaje, lnAnchoBol, lnEspaceIni) & sBoldOFF & Chr$(10)
    sCad = sCad & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10) & Chr(27) & "m" & Chr(10)
End If

imprimirBoletaDevolucionSobrantesPignoraticio = sCad
Exit Function
Error:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'CTI4 ERS0112020**********************
Public Function ImprimeBoletaAhorro(ByVal psTit As String, ByVal pstext As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
        ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
        ByVal pnInteresA As String, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
        Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
        Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, _
        Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
        Optional pdFecSis As Date, Optional sNomAge As String = "", Optional sCodUser As String = "", _
        Optional sLpt As String = "", Optional ByVal psCodCMAC As String, Optional ByVal pnMontoIFT As Double, Optional ByVal psCtaCodCred As String = "", Optional ByVal psNombreTitularEco As String = "", _
        Optional ByVal pbImpTMU As Boolean = False, Optional bEsCTS As Boolean = False) As String

Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax As Integer
Dim saux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsNroExt As String
Dim lnTope As Integer
Dim lsSaldoC As String
Dim lsCadArg As String
Dim lsInteres As String
Dim sCad As String
Dim lnCliAux As Long
Dim lsCliAux1 As String
Dim lsCliAux2 As String

Dim lnChq As Long
Dim lsChqAux1 As String
Dim lsChqAux2 As String
Dim lsNomAge As String

Dim lnNumLinCmac As Integer

Dim lsMensaje As String * 39
Dim sText As String

Dim oImpre As COMFunciones.FCOMImpresion ' CTI4 ERS0112020***


''  quitar este comentario

        Dim clsGen As COMDConstSistema.DCOMGeneral
        Set clsGen = New COMDConstSistema.DCOMGeneral
        
        Set oImpre = New COMFunciones.FCOMImpresion 'CTI4 ERS0112020***
 

ETIQ:

'On Error GoTo Error

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)

lsNroExt = Str(pnNumExt)


 ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(64)

ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) + Chr$(18)  ' cancela condensada
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) + Chr$(72)  ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
saux = 5

sFecha = Format$(pdFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(pnMonto, "#,##0.00")
sSDisp = Format$(pnSaldo, "#,##0.00")
lsSaldoC = Format$(pnSaldoC, "#,##0.00")

lsNomAge = sNomAge

ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(10)
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(10)

'Print #nFicSal, Chr$(10);
If psCodCMAC = "102" Then  ' CAJA LIMA
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "CM - AHORRO" & Chr$(10)
        
        If Mid(psCodCta, 9, 1) = 1 Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Chr$(10)
        End If
        
        If psNomAgeRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & Chr$(10)
        End If
        
        If psComCmac = "XXX" Then
            If psLin3 = "XXX" Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff 'Desactiva Negrita
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & lsNegritaOff   'Desactiva Negrita
                lnNumLinCmac = 1
            End If
            lnNumLinCmac = 0
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & lsNegritaOff & Chr$(10)   'Desactiva Negrita
            lnNumLinCmac = 1
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & Chr$(10)
        
        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
        
        If lnCliAux = 0 Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & Chr$(10)
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
            
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & Space(saux + sMax + sSep - Len(lsCliAux1)) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & Space(saux + sMax + sSep - Len(lsCliAux2)) & Chr$(10)
            
            lnCliAux = 1
        End If
        
        If pbSaldoInt Or pbCuenta Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + saux) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        psTit = Trim(psTit)
        psTit = oImpre.CentrarCadena(psTit, 28)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita & chr$(10)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni + 1) & "-----" & psTit & "-----" & Chr$(10)
        
        lnChq = InStr(1, pstext, "*", vbTextCompare)
        
        If psTexto = "XXX" Then
            If lnChq = 0 Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
            Else
                lsChqAux1 = (Mid(pstext, 1, lnChq - 1))
                lsChqAux2 = (Mid(pstext, lnChq + 1))

                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & Space(saux + sSep + sMax - Len(Trim(psTexto))) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff  'Desactiva Negrita
   
        If pbSaldoInt Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & Chr$(10)
                If pbOpSaldoC Then
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & Chr$(10)
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
                End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        lsInteres = pnInteresA
        
        If pbSaldoInt Then
            If lsInteres <> "No Valido" Then
                lsInteres = Format(lsInteres, "#,##0.00")
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Space(-1 + sSep) & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDoc = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDias = "----" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Space(-1 + sSep) & Chr$(10)
            lnTope = 4 - lnCliAux
        Else
            lnTope = 3 - lnCliAux
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & Chr$(10)
        If psCodUsuRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Chr$(10)  ''& Space(29 + sSep + sAux)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & Space(sSep) & Chr$(10)
        End If
        'Dim clsGen As COMDConstSistema.DCOMGeneral
        'Set clsGen = New COMDConstSistema.DCOMGeneral
        lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
        Set clsGen = Nothing
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & lsNegritaOn & oImpre.ImpreCarEsp(lsMensaje) & Space(-1 + sSep) & oImpre.ImpreCarEsp(lsMensaje) & lsNegritaOff
        
        lnNumLinCmac = lnNumLinCmac + 1
        
        For saux = 1 To (lnTope - lnNumLinCmac)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Next saux

Else   '' Caja Trujillo
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita
        If Not pbImpTMU Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "CMAC-MAYNAS - AHORROS" & Space(19 + sSep + saux) & "CMAC-MAYNAS - AHORRO" & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "CMAC-MAYNAS - AHORROS" & Chr$(10)
        End If
        If Mid(psCodCta, 9, 1) = 1 Then
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Chr$(10)
            End If
        Else
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Chr$(10)
            End If
        End If
        
        If psNomAgeRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Else
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & "Ag.Rem: " & Trim(psNomAgeRem) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Chr$(10)
            End If
        End If
        
        If psComCmac = "XXX" Then
            If psLin3 = "XXX" Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff 'Desactiva Negrita
            Else
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff  'Desactiva Negrita
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & psLin3 & lsNegritaOff   'Desactiva Negrita
            End If
                lnNumLinCmac = 1
            End If
            lnNumLinCmac = 0
        Else
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff & Chr$(10)   'Desactiva Negrita
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "NroDocCmac:" & psComCmac & lsNegritaOff & Chr$(10)   'Desactiva Negrita
            End If
            lnNumLinCmac = 1
        End If
        
        If Not pbImpTMU Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Fecha:" & Format(sFecha, "dd-mm-yyyy") & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & "Fecha:" & Format(sFecha, "dd-mm-yyyy") & Space(10) & "Hora:" & sHora & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Fecha:" & Format(sFecha, "dd-mm-yyyy") & Space(10) & "Hora:" & sHora & Chr$(10)
        End If
        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
        
        If lnCliAux = 0 Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & oImpre.ImpreCarEsp(psCliente) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Chr$(10)
            End If
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & Space(saux + sMax + sSep - Len(lsCliAux1)) & oImpre.ImpreCarEsp(lsCliAux1) & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & Space(saux + sMax + sSep - Len(lsCliAux2)) & oImpre.ImpreCarEsp(lsCliAux2) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & Chr$(10)
            End If
            lnCliAux = 1
        End If
        
        If pbSaldoInt Or pbCuenta Then
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + saux) & "Cuenta:" & psCodCta & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Cuenta:" & psCodCta & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        psTit = Trim(psTit)
        psTit = oImpre.CentrarCadena(psTit, 28)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita & chr$(10)
        If Not pbImpTMU Then
           ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni + 1) & "-----" & psTit & "-----" & Space(-1 + sSep) & "-----" & psTit & "-----" & Chr$(10)
        Else
          ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni + 1) & "-----" & psTit & "-----" & Chr$(10)
        End If
        lnChq = InStr(1, pstext, "*", vbTextCompare)
        
        If psTexto = "XXX" Then
            If lnChq = 0 Then
                If Not pbImpTMU Then
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Chr$(10)
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Chr$(10)
                End If
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
            Else
                lsChqAux1 = (Mid(pstext, 1, lnChq - 1))
                lsChqAux2 = (Mid(pstext, lnChq + 1))
                If Not pbImpTMU Then
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Chr$(10)
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Chr$(10)
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Chr$(10)
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Chr$(10)
                End If
            End If
        Else
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & Space(saux + sSep + sMax - Len(Trim(psTexto))) & oImpre.ImpreCarEsp(Trim(psTexto)) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & Chr$(10)
            End If
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff  'Desactiva Negrita
        
        If pnMontoIFT > 0 Then
            sText = "ITF Efectivo"
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & (Trim(Mid(sText, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(sText, 1, 28))) - Len(pnMontoIFT)) & Format(pnMontoIFT, "0.00") & Space(-1 + 15) & (Trim(Mid(sText, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(sText, 1, 28))) - Len(pnMontoIFT)) & Format(pnMontoIFT, "0.00") & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & (Trim(Mid(sText, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(sText, 1, 28))) - Len(pnMontoIFT)) & Format(pnMontoIFT, "0.00") & Chr$(10)
            End If
        End If
        If pbSaldoInt Then
                If Not bEsCTS Then
                    If Not pbImpTMU Then
                        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Chr$(10)
                    Else
                        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Chr$(10)
                    End If
                End If
                If pbOpSaldoC And Not bEsCTS Then
                    If Not pbImpTMU Then
                        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Chr$(10)
                    Else
                        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Chr$(10)
                    End If
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
                End If
 
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        lsInteres = pnInteresA
        
        If pbSaldoInt Then
            If lsInteres <> "No Valido" And Not bEsCTS Then
                lsInteres = Format(lsInteres, "#,##0.00")
                If Not pbImpTMU Then
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Space(-1 + sSep) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Chr$(10)
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Chr$(10)
                End If
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDoc = "" Then
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDias = "----" Then
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Space(-1 + sSep) & "Nro Dias Interes" & Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")) & psNumDias & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Chr$(10)
            End If
            lnTope = 4 - lnCliAux
        Else
            lnTope = 3 - lnCliAux
        End If
        If Not pbImpTMU Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & "---------------------------------------" & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "---------------------------------------" & Chr$(10)
        End If
        If psCodUsuRem = "" Then
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Space(29 + sSep + saux) & oImpre.ImpreCarEsp(sCodUser) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Chr$(10)
            End If
        Else
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & Space(sSep) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & Chr$(10)
            End If
        End If

        If Mid(Trim(psCtaCodCred), 6, 3) = "517" Then
            If Not pbImpTMU Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp("Cliente : " & Left(Trim(psNombreTitularEco) & Space(40), 29)) & Space(14) & oImpre.ImpreCarEsp("Cliente : " & Trim(psNombreTitularEco)) & Chr$(10)
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & oImpre.ImpreCarEsp("Cliente : " & Left(Trim(psNombreTitularEco) & Space(40), 29)) & Chr$(10)
            End If
        End If
        '*************************************************
        lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
        Set clsGen = Nothing
        If Not pbImpTMU Then
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & lsNegritaOn & oImpre.ImpreCarEsp(lsMensaje) & Space(-1 + sSep) & oImpre.ImpreCarEsp(lsMensaje) & lsNegritaOff
        Else
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & lsNegritaOn & oImpre.ImpreCarEsp(lsMensaje) & lsNegritaOff
        End If
        lnNumLinCmac = lnNumLinCmac + 1
        
        For saux = 1 To (lnTope - lnNumLinCmac)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Next saux
        
End If
    Exit Function
End Function
Public Function nPrintReciboExtorCargoCta(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psCtaAhoExt As String, ByVal pnITFAho As Currency, ByVal psNombCliente As String, ByVal psOperacion As String, _
        ByVal pnMonto As Currency, ByVal pnSaldoCapAnt As Currency, ByVal pnMovExtornado As Long, _
        ByVal psCodUser As String, ByVal psVarPersCMAC As String, _
        Optional ByVal psmensaje As String = "", Optional ByVal psUsuOri As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal pbImpTMU As Boolean = False) As String
        
    Dim lsCadImp As String
    Dim ofun As New COMFunciones.FCOMImpresion
    '** Para la Impresion del Codigo Antiguo en las Boletas
    Dim lsCodAntiguo As String
    Dim loConstSistema As New COMDConstSistema.NCOMConstSistema
    If loConstSistema.LeeConstSistema(122) = "1" Then
        Dim loGeneral As COMDConstSistema.DCOMGeneral
        Set loGeneral = New COMDConstSistema.DCOMGeneral
        lsCodAntiguo = loGeneral.GetCuentaAntigua(psCodCta)
    Else
        lsCodAntiguo = ""
    End If
    Set loConstSistema = Nothing
        
    Dim vEspacio As Integer
    vEspacio = 28
        
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    oFunI.Inicia psImpresora

    If pbImpTMU = False Then
        lsCadImp = "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        
        lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 30, 0) & Space(12) & Space(vEspacio + 4) _
                            & " " & ofun.ImpreFormat(psNomAgencia, 30, 0) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & " " & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & Space(28 + vEspacio) & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & oFunI.gPrnSaltoLinea
        
        If psVarPersCMAC <> "" Then
             lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & Space(5 + vEspacio) _
                                 & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        End If
        
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        
        lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & Space(vEspacio + 4) _
                            & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & " Cuenta    : " & Mid(psCtaAhoExt, 1, 5) & "-" & Mid(psCtaAhoExt, 6) & Space(10) & Space(5 + vEspacio) _
                            & " Cuenta    : " & Mid(psCtaAhoExt, 1, 5) & "-" & Mid(psCtaAhoExt, 6) & Space(10) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
                
        lsCadImp = lsCadImp & " ---- " & ofun.ImpreFormat("EXTORNO CARGO CTA", 25) & " ---- " & Space(9 + vEspacio) _
                            & " ---- " & ofun.ImpreFormat("EXTORNO CARGO CTA", 25) & " ---- " & oFunI.gPrnSaltoLinea

        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " MONTO     :" & ofun.ImpreFormat(pnMonto, 21, , True) & Space(11 + vEspacio) _
                            & " MONTO     :" & ofun.ImpreFormat(pnMonto, 21, , True) & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & " I.T.F.    :" & ofun.ImpreFormat(pnITFAho, 21, , True) & Space(11 + vEspacio) _
                            & " I.T.F.    :" & ofun.ImpreFormat(pnITFAho, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & " -------------------------------------" & Space(10 + vEspacio) _
                            & " -------------------------------------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & ofun.ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & Space(10) & Space(10 + vEspacio) _
                            & " " & ofun.ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & Space(10) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & Space(24 + vEspacio) & "  " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & Space(24 + vEspacio) & "  " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    Else
        lsCadImp = "" & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        
        lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psNomAgencia, 30, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & oFunI.gPrnSaltoLinea
        
        If psVarPersCMAC <> "" Then
             lsCadImp = lsCadImp & " " & ofun.ImpreFormat(psVarPersCMAC, 41, 0) & oFunI.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        End If
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " CLIENTE   : " & ofun.ImpreFormat(psNombCliente, 28) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " Cuenta    : " & Mid(psCtaAhoExt, 1, 5) & "-" & Mid(psCtaAhoExt, 6) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " ---- " & ofun.ImpreFormat("EXTORNO CARGO CTA", 25) & " ---- " & oFunI.gPrnSaltoLinea

        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " MONTO     :" & ofun.ImpreFormat(pnMonto, 21, , True) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " I.T.F.    :" & ofun.ImpreFormat(pnITFAho, 21, , True) & oFunI.gPrnSaltoLinea
                            
        lsCadImp = lsCadImp & " -------------------------------------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & ofun.ImpreCarEsp("Trans. Extornada : ") & pnMovExtornado & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO EXTORNA : " & Format(psCodUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & " " & "USUARIO OPERADOR: " & Format(psUsuOri, "@@@@") & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Space(15) & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)

    End If
    nPrintReciboExtorCargoCta = lsCadImp
    Set oFunI = Nothing
End Function
'END CTI4
