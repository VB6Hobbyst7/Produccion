VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCaptaServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Function DigChequeo(ByVal sRecibo As String, ByVal sCodigoCliente As String, _
        ByVal nMonto As Double) As Integer

Dim sFacSerNro As String, sFacNro As String, sDigCheq As String
Dim sValor As String
Dim nLongValor As Integer, nValTotal As Integer, N As Integer
sFacSerNro = Mid(sRecibo, 1, 3)
sFacNro = Mid(sRecibo, 4, 8)
sDigCheq = Right(sRecibo, 2)
sValor = sCodigoCliente & sFacSerNro & sFacNro & Right(String((10 - Len(Str(nMonto * 100))), " ") & Str(nMonto * 100), 8)
nLongValor = Len(sValor)
N = 0
nValTotal = 0
Do While N < nLongValor
    N = N + 1
    nValTotal = nValTotal + N * Val(Mid(sValor, N, 1))
Loop
DigChequeo = 97 - (nValTotal - Int(nValTotal / 97) * 97)
End Function

Public Function ValSedalib(ByVal psRecibo As String, ByVal psCodigoCliente As String, ByVal pcMonto As Currency) As Boolean
If DigChequeo(psRecibo, psCodigoCliente, pcMonto) = Val(Right(psRecibo, 2)) Then
    ValSedalib = True
Else
    ValSedalib = False
End If
End Function

Public Sub ImpBolUNT(ByVal psCodAlu As String, ByVal psNumDoc As String, ByVal psNombre As String, _
                      ByVal psMonTot As String, ByVal psNomEscuela As String, ByVal psHoy As String, _
                      ByVal psFecPag As String, ByVal psMonEsSalud As String, ByVal psOrdOtros As String, _
                      ByVal psMonCuoFac As String, ByVal psMonCarnet As String, ByVal sNomAge As String, _
                      ByVal sCodUser As String, ByVal sLpt As String)
                      
    Dim lsNegritaOn As String
    Dim lsNegritaOff As String
    
    lsNegritaOn = Chr$(27) + Chr$(71)
    lsNegritaOff = Chr$(27) + Chr$(72)

    Dim lsCadena As String
    Dim lsAux1 As String
    Dim lsAux2 As String
    Dim lsAux3 As String * 44
    Dim lsAux4 As String * 44
    Dim lsAux5 As String
    Dim lsAux6 As String
    Dim lsAux7 As String
    Dim lsAux8 As String
    Dim lsAux9 As String
    Dim lsAux10 As String
    Dim lsAux11 As String
    Dim lsAux12 As String
    Dim lsAux13 As String
    Dim lsAux14 As String
    Dim lsTot As String
    Dim lsMat As String
    
    Dim lsConcep(14) As String
    Dim lnConcep As Integer
    
    Dim lnLMax As Long
    Dim lnEMax As Long
    Dim lnTMax As Long
    
    lnTMax = 3
    lnEMax = 15
    lnLMax = 45
    If Trim(psMonCuoFac) = "" Then psMonCuoFac = "0.00"
    If Trim(psMonCarnet) = "" Then psMonCarnet = "0.00"
    lsCadena = ""
    lsAux1 = "CMACT - UNIVERSIDAD NACIONAL DE TRUJILLO"
    lsAux2 = sNomAge
    lsAux3 = "ALUMNO  : " & psNombre
    lsAux4 = "ESCUELA : " & psNomEscuela
    lsAux5 = "FEC.PAG : " & Left(psHoy, 10) & "  HORA : " & Right(psHoy, 8)
    lsAux6 = "ORD.PAG : " & Mid(psNumDoc, 1, 2) & "-" & Right(psNumDoc, 6) & Space(3) & "COD.ALU. : " & psCodAlu
    lsAux7 = "-------PAGO DE SERVICIOS EDUCATIVOS-------"
    lsAux8 = "EsSalud        : " & Mid(psOrdOtros, 1, 2) & "-" & Right(psOrdOtros, 6) & Space(14 - Len(psMonEsSalud)) & psMonEsSalud
    lsAux9 = "Cuota Facultad : " & Space(23 - Len(psMonCuoFac)) & psMonCuoFac
    lsMat = Format$(CDbl(psMonTot) - CDbl(psMonCuoFac) - CDbl(psMonCarnet), "#,##0.00")
    lsAux10 = "Matricula      : " & Space(23 - Len(lsMat)) & lsMat
    lsAux14 = "Carnet Univ.   : " & Space(23 - Len(psMonCarnet)) & psMonCarnet
    lsTot = Format$(CDbl(psMonTot) + CDbl(psMonEsSalud), "#,##0.00")
    lsAux11 = "TOTAL A PAGAR  : " & Space(23 - Len(lsTot)) & lsTot
    lsAux12 = String(42, "-")
    lsAux13 = sCodUser
    
    lsCadena = lsNegritaOn & lsCadena & "." & Space(lnTMax - 1) & lsAux1 & Space(lnLMax - Len(lsAux1)) & Space(lnEMax) & lsAux1 & Chr(10)
    lsCadena = lsCadena & Space(lnTMax) & lsAux2 & Space(lnLMax - Len(lsAux2)) & Space(lnEMax) & lsAux2 & Chr(10) & lsNegritaOff
    lsCadena = lsCadena & Space(lnTMax) & lsNegritaOn & lsAux3 & lsNegritaOff & Space(lnLMax - Len(lsAux3)) & Space(lnEMax) & lsNegritaOn & lsAux3 & Chr(10) & lsNegritaOff
    lsCadena = lsCadena & Space(lnTMax) & lsNegritaOn & lsAux4 & lsNegritaOff & Space(lnLMax - Len(lsAux4)) & Space(lnEMax) & lsNegritaOn & lsAux4 & Chr(10) & lsNegritaOff
    lsCadena = lsCadena & Space(lnTMax) & lsNegritaOn & lsAux5 & lsNegritaOff & Space(lnLMax - Len(lsAux5)) & Space(lnEMax) & lsNegritaOn & lsAux5 & Chr(10) & lsNegritaOff
    lsCadena = lsCadena & Space(lnTMax) & lsNegritaOn & lsAux6 & lsNegritaOff & Space(lnLMax - Len(lsAux6)) & Space(lnEMax) & lsNegritaOn & lsAux6 & Chr(10) & lsNegritaOff
    lsCadena = lsCadena & Space(lnTMax) & lsNegritaOn & lsAux7 & Space(lnLMax - Len(lsAux7)) & Space(lnEMax) & lsAux7 & Chr(10) & lsNegritaOff
    
    lsCadena = lsCadena & Space(lnTMax) & lsAux8 & Space(lnLMax - Len(lsAux8)) & Space(lnEMax) & lsAux8 & Chr(10)
    lsCadena = lsCadena & Space(lnTMax) & lsAux9 & Space(lnLMax - Len(lsAux9)) & Space(lnEMax) & lsAux9 & Chr(10)
    lsCadena = lsCadena & Space(lnTMax) & lsAux10 & Space(lnLMax - Len(lsAux10)) & Space(lnEMax) & lsAux10 & Chr(10)
    lsCadena = lsCadena & Space(lnTMax) & lsAux14 & Space(lnLMax - Len(lsAux14)) & Space(lnEMax) & lsAux14 & Chr(10)
    
    lsCadena = lsCadena & Space(lnTMax) & lsNegritaOn & lsAux11 & Space(lnLMax - Len(lsAux11)) & Space(lnEMax) & lsAux11 & Chr(10)
    
    lsCadena = lsCadena & Space(lnTMax) & lsAux12 & Space(lnLMax - Len(lsAux12)) & Space(lnEMax) & lsAux12 & Chr(10)
    lsCadena = lsCadena & Space(lnTMax) & lsAux13 & Space(lnLMax - Len(lsAux13)) & Space(lnEMax) & lsAux13 & Chr(10) & lsNegritaOff
    
    Dim ArcSal As Integer
    ArcSal = FreeFile
    Open sLpt For Output As ArcSal
    Print #ArcSal, Chr$(27) & Chr$(64);            'Inicializa Impresora
    Print #ArcSal, Chr$(18);
    Print #ArcSal, Chr$(27) & Chr$(50);            'Espaciamiento lineas 1/6 pulg.1
    Print #ArcSal, Chr$(27) & Chr$(67) & Chr$(22) '   Chr$(nLineas); 'Longitud de página a 66 líneas
    Print #ArcSal, Chr$(27) & Chr$(108) & Chr$(0); 'Tipo letra : 0,1,2 - Roman,SansS,Courier
    Print #ArcSal, Chr$(27) & Chr$(103);            'Tamaño  : 80, 77, 103
    Print #ArcSal, Chr$(27) & Chr$(120) & Chr$(0)  'Draf : 1 pasada
    Print #ArcSal, lsCadena
    Print #ArcSal, Chr$(12);   'Nueva página
    Print #ArcSal, Chr$(18);   'Retorna al tipo de letra normal
    Close ArcSal
End Sub

Private Function ImprimeBoletaCancelacionGiros(ByVal sNumGiro As String, ByVal sRemitente As String, _
        ByRef sDestinatario() As String, ByVal nMontoGiro As Double, ByVal dFecha As Date, _
        ByVal sHora As String, ByVal sNomAge As String, ByVal sCodUser As String, _
        ByVal sLpt As String, Optional ByVal bCadena As Boolean = False) As String
                      
                      
    Dim sNegritaOn As String
    Dim sNegritaOff As String
    
    sNegritaOn = Chr$(27) + Chr$(71)
    sNegritaOff = Chr$(27) + Chr$(72)


    Dim sCadena As String
    Dim sAux1 As String * 44, sAux2 As String * 44
    Dim sAux3 As String * 44, sAux4 As String * 44
    Dim sAux5 As String * 44, sAux6 As String * 44
    Dim sAux7 As String * 44, sAux8 As String * 44
    Dim sAux9 As String * 44, sAux10 As String * 44
    Dim sAux11 As String * 44
    Dim sAux12 As String, sAux13 As String
    Dim sAux14 As String, sAux15 As String, sAux16 As String
    Dim sMonto As String * 23, sComision As String * 23, sTot As String * 23
    
    Dim nLMax As Long
    Dim nEMax As Long
    Dim nTMax As Long
    
     Dim sSep As Integer
    Dim sIni As Integer
    Dim sMax As Integer
    Dim saux As Integer
    
    sSep = 15
    sIni = 0
    sMax = 38
    saux = 15
    
    nTMax = 3
    nEMax = 15
    nLMax = 45
    RSet sMonto = Format$(nMontoGiro, "#,##0.00")
    sCadena = ""
    sAux1 = "CMACICA - AHORROS"
    sAux2 = sNomAge & IIf(Mid(sNumGiro, 9, 1) = "1", "- SOLES", "- DOLARES")
    sAux3 = ""
    sAux4 = "Fecha: " & Format$(dFecha, "dd/mm/yyyy") & "Hora: " & sHora
    sAux5 = "REMITENTE: " & Trim(sRemitente)
    sAux6 = "DESTINAT.: " & Trim(sDestinatario(0))
    sAux7 = "           " & Trim(sDestinatario(1))
    sAux8 = "           " & Trim(sDestinatario(2))
    sAux9 = "           " & Trim(sDestinatario(3))
    
    sAux10 = "GIRO NRO.: " & sNumGiro
    sAux11 = "---- CANCELACION DE GIRO EFECTIVO ----"
    sAux12 = "Monto de Giro: " & IIf(Mid(sNumGiro, 9, 1) = "1", "S/.", "US$") & Trim(sMonto)
    sAux13 = ""
    sAux14 = "TOTAL A PAGAR: " & IIf(Mid(sNumGiro, 9, 1) = "1", "S/.", "US$") & Trim(sMonto)
    sAux15 = String(38, "-")
    sAux16 = sCodUser
    
    
    sCadena = sNegritaOn & sCadena & Space(sIni) & Trim(sAux1) & Space(saux + sMax - Len(Trim(sAux1))) & Trim(sAux1) & Chr$(10)  ' & Chr$(10)
    sCadena = sCadena & Space(sIni) & Trim(sAux2) & Space(saux + sMax - Len(Trim(sAux2))) & Trim(sAux2) & Chr$(10) & sNegritaOff  ' sAux2 & Chr$(10) & sNegritaOff
    sCadena = sCadena & Space(sIni) & Trim(sAux3) & Space(saux + sMax - Len(Trim(sAux3))) & Trim(sAux3) & Chr$(10) ' sAux3 & Chr$(10)
    sCadena = sCadena & Space(sIni) & Trim(sAux4) & Space(saux + sMax - Len(Trim(sAux4))) & Trim(sAux4) & Chr$(10) ' sAux4 & Chr$(10)
    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux5) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux5))) & Trim(sAux5) & Chr$(10)   'Space(nEMax) & sNegritaOn & sAux5 & Chr$(10) & sNegritaOff
    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux6) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux6))) & Trim(sAux6) & Chr$(10) 'Space(nEMax) & sNegritaOn & sAux6 & Chr$(10) & sNegritaOff
    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux7) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux7))) & Trim(sAux7) & Chr$(10)
    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux8) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux8))) & Trim(sAux8) & Chr$(10)
    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux9) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux9))) & Trim(sAux9) & Chr$(10)
    
    
    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux10) & Space(saux + sMax - Len(Trim(sAux10))) & Trim(sAux10) & Chr$(10)  'sAux7 & Chr$(10) & sNegritaOff
    sCadena = sCadena & Space(sIni) & Trim(sAux11) & Space(saux + sMax - Len(Trim(sAux11))) & Trim(sAux11) & Chr$(10) '& Space(nEMax) & sAux8 & Chr$(10) '
    sCadena = sCadena & Space(sIni) & Trim(sAux12) & Space(saux + sMax - Len(Trim(sAux12))) & Trim(sAux12) & Chr$(10) ' Space(nEMax) & sAux9 & Chr$(10)
    sCadena = sCadena & Space(sIni) & Trim(sAux13) & Space(saux + sMax - Len(Trim(sAux13))) & Trim(sAux13) & Chr$(10) 'Space(nEMax) & sAux10 & Chr$(10)
    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux14) & Space(saux + sMax - Len(Trim(sAux14))) & Trim(sAux14) & Chr$(10) 'Space(nEMax) & sAux11 & Chr$(10)
    
    sCadena = sCadena & Space(sIni) & Trim(sAux15) & Space(saux + sMax - Len(Trim(sAux15))) & Trim(sAux15) & Chr$(10)  'Space(nEMax) & sAux12 & Chr$(10)
    sCadena = sCadena & Space(sIni) & Trim(sAux16) & Space(saux + sMax - Len(Trim(sAux16))) & Trim(sAux16) & Chr$(10) 'Space(nEMax) & sAux13 & Chr$(10) & sNegritaOff
    
    If bCadena Then
        ImprimeBoletaCancelacionGiros = sCadena
    Else
        Dim ArcSal As Integer
        ArcSal = FreeFile
        Open sLpt For Output As ArcSal
        Print #ArcSal, Chr$(27) & Chr$(64);            'Inicializa Impresora
        Print #ArcSal, Chr$(18);
        Print #ArcSal, Chr$(27) & Chr$(50);            'Espaciamiento lineas 1/6 pulg.1
        Print #ArcSal, Chr$(27) & Chr$(67) & Chr$(22); '   Chr$(nLineas); 'Longitud de página a 66 líneas
        Print #ArcSal, Chr$(27) + Chr$(107) + Chr$(0); 'Tipo letra : 0,1,2 - Roman,SansS,Courier
        Print #ArcSal, Chr$(27) & Chr$(77);           'Tamaño  : 80, 77, 103
        Print #ArcSal, Chr$(27) & Chr$(120) & Chr$(0)  'Draf : 1 pasada
        Print #ArcSal, sCadena
        Print #ArcSal, Chr$(12);   'Nueva página
        Print #ArcSal, Chr$(18);   'Retorna al tipo de letra normal
        Close ArcSal
    End If
End Function



'Private Function PRINTBoletaCancelacionGirosANT(ByVal sNumGiro As String, ByVal sRemitente As String, _
'        ByVal sDestinatario As String, ByVal nMontoGiro As Double, ByVal dFecha As Date, _
'        ByVal sHora As String, ByVal sNomAge As String, ByVal sCodUser As String, _
'        ByVal sLpt As String, Optional ByVal bCadena As Boolean = False) As String
'
'
'    Dim sNegritaOn As String
'    Dim sNegritaOff As String
'
'    sNegritaOn = Chr$(27) + Chr$(71)
'    sNegritaOff = Chr$(27) + Chr$(72)
'
'
'    Dim sCadena As String
'    Dim sAux1 As String * 44, sAux2 As String * 44
'    Dim sAux3 As String * 44, sAux4 As String * 44
'    Dim sAux5 As String * 44, sAux6 As String * 44
'    Dim sAux7 As String * 44, sAux8 As String * 44
'    Dim sAux9 As String, sAux10 As String
'    Dim sAux11 As String, sAux12 As String, sAux13 As String
'    Dim sMonto As String * 23, sComision As String * 23, sTot As String * 23
'
'    Dim nLMax As Long
'    Dim nEMax As Long
'    Dim nTMax As Long
'
'     Dim sSep As Integer
'    Dim sIni As Integer
'    Dim sMax As Integer
'    Dim saux As Integer
'
'    sSep = 15
'    sIni = 1
'    sMax = 38
'    saux = 15
'
'    nTMax = 3
'    nEMax = 15
'    nLMax = 45
'    RSet sMonto = Format$(nMontoGiro, "#,##0.00")
'    sCadena = ""
'    sAux1 = "CMACICA - AHORROS"
'    sAux2 = sNomAge & IIf(Mid(sNumGiro, 9, 1) = "1", "- SOLES", "- DOLARES")
'    sAux3 = ""
'    sAux4 = "Fecha     : " & Format$(dFecha, "dd/mm/yyyy") & "        Hora : " & sHora
'    sAux5 = "REMITENTE : " & sRemitente
'    sAux6 = "DESTINAT. : " & sDestinatario
'    sAux7 = "GIRO NRO. : " & sNumGiro
'    sAux8 = "------  CANCELACION DE GIRO EFECTIVO  ------"
'    sAux9 = "Monto de Giro  : " & IIf(Mid(sNumGiro, 9, 1) = "1", "S/.", "US$") & sMonto
'    sAux10 = ""
'    sAux11 = "TOTAL A PAGAR  : " & IIf(Mid(sNumGiro, 9, 1) = "1", "S/.", "US$") & sMonto
'    sAux12 = String(38, "-")
'    sAux13 = sCodUser
'
'
'    sCadena = sNegritaOn & sCadena & "." & Space(nTMax - 1) & sAux1 & Space(nLMax - Len(sAux1)) & Space(nEMax) & Chr$(10) 'sAux1 & Chr$(10)
'    sCadena = sCadena & Space(nTMax) & sAux2 & Space(nLMax - Len(sAux2)) & Space(nEMax) & Chr$(10) & sNegritaOff ' sAux2 & Chr$(10) & sNegritaOff
'    sCadena = sCadena & Space(nTMax) & sAux3 & Space(nLMax - Len(sAux3)) & Space(nEMax) & Chr$(10) ' sAux3 & Chr$(10)
'    sCadena = sCadena & Space(nTMax) & sAux4 & Space(nLMax - Len(sAux4)) & Space(nEMax) & Chr$(10) ' sAux4 & Chr$(10)
'    sCadena = sCadena & Space(nTMax) & sNegritaOn & sAux5 & sNegritaOff & Space(nLMax - Len(sAux5)) & Chr$(10)  'Space(nEMax) & sNegritaOn & sAux5 & Chr$(10) & sNegritaOff
'    sCadena = sCadena & Space(nTMax) & sNegritaOn & sAux6 & sNegritaOff & Space(nLMax - Len(sAux6)) & Chr$(10) 'Space(nEMax) & sNegritaOn & sAux6 & Chr$(10) & sNegritaOff
'    sCadena = sCadena & Space(nTMax) & sNegritaOn & sAux7 & Space(nLMax - Len(sAux7)) & Space(nEMax) & Chr$(10)  'sAux7 & Chr$(10) & sNegritaOff
'
'    sCadena = sCadena & Space(nTMax) & sAux8 & Space(nLMax - Len(sAux8)) & Chr$(10) '& Space(nEMax) & sAux8 & Chr$(10) '
'    sCadena = sCadena & Space(nTMax) & sAux9 & Space(nLMax - Len(sAux9)) & Chr$(10) ' Space(nEMax) & sAux9 & Chr$(10)
'    sCadena = sCadena & Space(nTMax) & sAux10 & Space(nLMax - Len(sAux10)) & Chr$(10) 'Space(nEMax) & sAux10 & Chr$(10)
'    sCadena = sCadena & Space(nTMax) & sNegritaOn & sAux11 & Space(nLMax - Len(sAux11)) & Chr$(10) 'Space(nEMax) & sAux11 & Chr$(10)
'
'    sCadena = sCadena & Space(nTMax) & sAux12 & Space(nLMax - Len(sAux12)) & Chr$(10) 'Space(nEMax) & sAux12 & Chr$(10)
'    sCadena = sCadena & Space(nTMax) & sAux13 & Space(nLMax - Len(sAux13)) & Chr$(10) 'Space(nEMax) & sAux13 & Chr$(10) & sNegritaOff
'
'    If bCadena Then
'        ImprimeBoletaCancelacionGiros = sCadena
'    Else
'        Dim ArcSal As Integer
'        ArcSal = FreeFile
'        Open sLpt For Output As ArcSal
'        Print #ArcSal, Chr$(27) & Chr$(64);            'Inicializa Impresora
'        Print #ArcSal, Chr$(18);
'        Print #ArcSal, Chr$(27) & Chr$(50);            'Espaciamiento lineas 1/6 pulg.1
'        Print #ArcSal, Chr$(27) & Chr$(67) & Chr$(22) '   Chr$(nLineas); 'Longitud de página a 66 líneas
'        Print #ArcSal, Chr$(27) & Chr$(108) & Chr$(0); 'Tipo letra : 0,1,2 - Roman,SansS,Courier
'        Print #ArcSal, Chr$(27) & Chr$(103);            'Tamaño  : 80, 77, 103
'        Print #ArcSal, Chr$(27) & Chr$(120) & Chr$(0)  'Draf : 1 pasada
'        Print #ArcSal, sCadena
'        Print #ArcSal, Chr$(12);   'Nueva página
'        Print #ArcSal, Chr$(18);   'Retorna al tipo de letra normal
'        Close ArcSal
'    End If
'End Function

Public Function ImprimeBoletaGiros(ByVal sNumGiro As String, ByVal sRemitente As String, _
        ByRef sDestinatario() As String, ByVal nMontoGiro As Double, ByVal nComision As Double, _
        ByVal dFecha As Date, ByVal sHora As String, ByVal sNomAge As String, ByVal sCodUser As String, _
        ByVal sLpt As String, Optional ByVal bCadena As Boolean = False, Optional pnValorItf As Double = 0, _
        Optional sAgenciaDestino As String = "") As String
                      
                      
    Dim sNegritaOn As String
    Dim sNegritaOff As String
    Dim sSep As Integer
    Dim sIni As Integer
    Dim sMax As Integer
    Dim saux As Integer
    
   ' ReDim sDestinatario(4)
    
    
    sNegritaOn = Chr$(27) + Chr$(71)
    sNegritaOff = Chr$(27) + Chr$(72)

    Dim sCadena As String
    Dim sAux1 As String * 44, sAux2 As String * 44
    Dim sAux3 As String * 44, sAux4 As String * 44
    Dim sAux5 As String * 30, sAux6 As String * 30
    Dim sAux7 As String * 30, sAux8 As String * 30
    Dim sAux9 As String * 44
    Dim sAux10 As String * 44, sAux11 As String * 44
    Dim sAux12 As String, sAux13 As String
    Dim sAux14 As String, sAux15 As String, sAux16 As String, saux17 As String * 44
    
    Dim sMonto As String * 23, sComision As String * 23, sTot As String * 23
    
    Dim nLMax As Long
    Dim nEMax As Long
    Dim nTMax As Long
    
    'nTMax = 3
    'nEMax = 15
    'nLMax = 45
 
    
    sSep = 15
    sIni = 0
    'sMax = 38
    sMax = 35
    saux = 15
    
    RSet sMonto = Format$(nMontoGiro, "#,##0.00")
    RSet sComision = Format$(nComision, "#,##0.00")
    RSet sTot = Trim(Format$(nMontoGiro + nComision - pnValorItf, "#,##0.00"))
    sCadena = ""
    sAux1 = "CMACICA - AHORROS"
    sAux2 = sNomAge & IIf(Mid(sNumGiro, 9, 1) = "1", "- SOLES", "- DOLARES")
    sAux3 = "Fecha: " & Format$(dFecha, "dd/mm/yyyy") & "Hora : " & sHora
    sAux4 = "REMIT.: " & Trim(sRemitente)
    sAux5 = "DEST.: " & sDestinatario(0)
    sAux6 = "       " & sDestinatario(1)
    sAux7 = "       " & sDestinatario(2)
    sAux8 = "       " & sDestinatario(3)
    sAux9 = "GIRO NRO.: " & sNumGiro
    sAux10 = "----- REGISTRO DE GIROS EFECTIVO -----"
    sAux11 = "Monto Giro: " & IIf(Mid(sNumGiro, 9, 1) = "1", "S/.", "US$") & sMonto
    sAux12 = "Comisión  : " & IIf(Mid(sNumGiro, 9, 1) = "1", "S/.", "US$") & sComision
    sAux13 = ""
    sAux14 = "TOTAL A PAGAR:" & IIf(Mid(sNumGiro, 9, 1) = "1", "S/.", "US$") & sTot
    sAux15 = String(38, "-")
    sAux16 = sCodUser
    saux17 = "DESTINO:" & sAgenciaDestino
    
'    sCadena = sNegritaOn & sCadena & Space(sIni) & Trim(sAux1) & Space(saux + sMax - Len(Trim(sAux1))) & Trim(sAux1) & Chr$(10)  'sAux1 & Chr$(10)
'    sCadena = sCadena & Space(sIni) & Trim(sAux2) & Space(saux + sMax - Len(Trim(sAux2))) & Trim(sAux2) & Chr$(10) & sNegritaOff ' sAux2 & Chr$(10) & sNegritaOff
'    sCadena = sCadena & Space(sIni) & Trim(sAux3) & Space(saux + sMax - Len(Trim(sAux3))) & Trim(sAux3) & Chr$(10) ' sAux3 & Chr$(10)
'    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux4) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux4))) & Trim(sAux4) & sNegritaOn & Chr$(10) & sNegritaOff  'sAux4 & Chr$(10) & sNegritaOff
'
'    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux5) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux5))) & Trim(sAux5) & sNegritaOn & Chr$(10) & sNegritaOff 'sAux5 & Chr$(10) & sNegritaOff
'    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux6) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux6))) & Trim(sAux6) & sNegritaOn & Chr$(10) & sNegritaOff 'sAux5 & Chr$(10) & sNegritaOff
'    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux7) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux7))) & Trim(sAux7) & sNegritaOn & Chr$(10) & sNegritaOff 'sAux5 & Chr$(10) & sNegritaOff
'    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux8) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux8))) & Trim(sAux8) & sNegritaOn & Chr$(10) & sNegritaOff 'sAux5 & Chr$(10) & sNegritaOff
'
'
'
'    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux9) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux9))) & Trim(sAux9) & sNegritaOn & Chr$(10) & sNegritaOff 'sAux6 & Chr$(10) & sNegritaOff
'    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux10) & Space(saux + sMax - Len(Trim(sAux10))) & Trim(sAux10) & Chr$(10) & sNegritaOff   ' sAux7 & Chr$(10) & sNegritaOff
'
'    sCadena = sCadena & Space(sIni) & Trim(sAux11) & Space(saux + sMax - Len(Trim(sAux11))) & Trim(sAux11) & Chr$(10) ' sAux8 & Chr$(10)
'    sCadena = sCadena & Space(sIni) & Trim(sAux12) & Space(saux + sMax - Len(Trim(sAux12))) & Trim(sAux12) & Chr$(10) ' sAux9 & Chr$(10)
'    sCadena = sCadena & Space(sIni) & Trim(sAux13) & Space(saux + sMax - Len(Trim(sAux13))) & Trim(sAux13) & Chr$(10) ' sAux10 & Chr$(10)
'    sCadena = sCadena & Space(sIni) & sNegritaOn & Trim(sAux14) & Space(saux + sMax - Len(Trim(sAux14))) & Trim(sAux14) & Chr$(10) ' sAux11 & Chr$(10)
'
'    sCadena = sCadena & Space(sIni) & Trim(sAux15) & Space(saux + sMax - Len(Trim(sAux15))) & Trim(sAux15) & Chr$(10)  ' sAux12 &  Chr$(10)
'    sCadena = sCadena & Space(sIni) & Trim(sAux16) & Space(saux + sMax - Len(Trim(sAux16))) & Trim(sAux16) & Chr$(10) ' sAux13 & Chr$(10) & sNegritaOff
'    sCadena = sCadena & Space(sIni) & Trim(saux17) & Space(saux + sMax - Len(Trim(saux17))) & Trim(saux17) & Chr$(10)

    sCadena = Chr$(27) & Chr$(77)
    sCadena = sNegritaOn & sCadena & Space(sIni) & Trim(sAux1) & Space(saux + sMax - Len(Trim(sAux1))) & Trim(sAux1) & Chr$(10)  'sAux1 & Chr$(10)
    sCadena = sCadena & Trim(sAux2) & Space(saux + sMax - Len(Trim(sAux2))) & Trim(sAux2) & Chr$(10) & sNegritaOff  ' sAux2 & Chr$(10) & sNegritaOff
    sCadena = sCadena & Trim(sAux3) & Space(saux + sMax - Len(Trim(sAux3))) & Trim(sAux3) & Chr$(10)  ' sAux3 & Chr$(10)
    sCadena = sCadena & sNegritaOn & Trim(sAux4) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux4))) & Trim(sAux4) & sNegritaOn & Chr$(10) & sNegritaOff   'sAux4 & Chr$(10) & sNegritaOff

    sCadena = sCadena & sNegritaOn & Trim(sAux5) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux5))) & Trim(sAux5) & sNegritaOn & Chr$(10) & sNegritaOff  'sAux5 & Chr$(10) & sNegritaOff
    sCadena = sCadena & sNegritaOn & Trim(sAux6) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux6))) & Trim(sAux6) & sNegritaOn & Chr$(10) & sNegritaOff  'sAux5 & Chr$(10) & sNegritaOff
    sCadena = sCadena & sNegritaOn & Trim(sAux7) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux7))) & Trim(sAux7) & sNegritaOn & Chr$(10) & sNegritaOff  'sAux5 & Chr$(10) & sNegritaOff
    sCadena = sCadena & sNegritaOn & Trim(sAux8) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux8))) & Trim(sAux8) & sNegritaOn & Chr$(10) & sNegritaOff  'sAux5 & Chr$(10) & sNegritaOff



    sCadena = sCadena & sNegritaOn & Trim(sAux9) & sNegritaOff & Space(saux + sMax - Len(Trim(sAux9))) & Trim(sAux9) & sNegritaOn & Chr$(10) & sNegritaOff  'sAux6 & Chr$(10) & sNegritaOff
    sCadena = sCadena & sNegritaOn & Trim(sAux10) & Space(saux + sMax - Len(Trim(sAux10))) & Trim(sAux10) & Chr$(10) & sNegritaOff    ' sAux7 & Chr$(10) & sNegritaOff

    sCadena = sCadena & Trim(sAux11) & Space(saux + sMax - Len(Trim(sAux11))) & Trim(sAux11) & Chr$(10)  ' sAux8 & Chr$(10)
    sCadena = sCadena & Trim(sAux12) & Space(saux + sMax - Len(Trim(sAux12))) & Trim(sAux12) & Chr$(10)  ' sAux9 & Chr$(10)
    sCadena = sCadena & Trim(sAux13) & Space(saux + sMax - Len(Trim(sAux13))) & Trim(sAux13) & Chr$(10)  ' sAux10 & Chr$(10)
    sCadena = sCadena & sNegritaOn & Trim(sAux14) & Space(saux + sMax - Len(Trim(sAux14))) & Trim(sAux14) & Chr$(10)  ' sAux11 & Chr$(10)

    sCadena = sCadena & Trim(sAux15) & Space(saux + sMax - Len(Trim(sAux15))) & Trim(sAux15) & Chr$(10)   ' sAux12 &  Chr$(10)
    sCadena = sCadena & Trim(sAux16) & Space(saux + sMax - Len(Trim(sAux16))) & Trim(sAux16) & Chr$(10)  ' sAux13 & Chr$(10) & sNegritaOff
    sCadena = sCadena & Trim(saux17) & Space(saux + sMax - Len(Trim(saux17))) & Trim(saux17) & Chr$(10)


    
    If bCadena Then
        ImprimeBoletaGiros = sCadena
    Else
        Dim ArcSal As Integer
        ArcSal = FreeFile
        Open sLpt For Output As ArcSal
        Print #ArcSal, Chr$(27) & Chr$(64);            'Inicializa Impresora
        'Print #ArcSal, Chr$(18);
        Print #ArcSal, Chr$(27) & Chr$(50);            'Espaciamiento lineas 1/6 pulg.1
        Print #ArcSal, Chr$(27) & Chr$(67) & Chr$(22) '   Chr$(nLineas); 'Longitud de página a 22 líneas
        Print #ArcSal, Chr$(27) + Chr$(107) + Chr$(0);  'Tipo letra : 0,1,2 - Roman,SansS,Courier
        Print #ArcSal, Chr$(27) & Chr$(77);            'Tamaño  : 80, 77, 103
        Print #ArcSal, Chr$(27) & Chr$(120) & Chr$(0)  'Draf : 1 pasada
        Print #ArcSal, sCadena
        Print #ArcSal, Chr$(12);   'Nueva página
        'Print #ArcSal, Chr$(18);   'Retorna al tipo de letra normal
        Close ArcSal
    End If
End Function

'Public Function PRINTBoletaGirosAnt(ByVal sNumGiro As String, ByVal sRemitente As String, _
'        ByVal sDestinatario As String, ByVal nMontoGiro As Double, ByVal nComision As Double, _
'        ByVal dFecha As Date, ByVal sHora As String, ByVal sNomAge As String, ByVal sCodUser As String, _
'        ByVal sLpt As String, Optional ByVal bCadena As Boolean = False) As String
'
'
'    Dim sNegritaOn As String
'    Dim sNegritaOff As String
'    Dim sSep As Integer
'    Dim sIni As Integer
'    Dim sMax As Integer
'    Dim saux As Integer
'
'    sSep = 15
'    sIni = 1
'    sMax = 33
'    saux = 5
'
'
'    sNegritaOn = Chr$(27) + Chr$(71)
'    sNegritaOff = Chr$(27) + Chr$(72)
'
'    Dim sCadena As String
'    Dim sAux1 As String * 44, sAux2 As String * 44
'    Dim sAux3 As String * 44, sAux4 As String * 44
'    Dim sAux5 As String * 44, sAux6 As String * 44
'    Dim sAux7 As String * 44, sAux8 As String * 44
'    Dim sAux9 As String, sAux10 As String
'    Dim sAux11 As String, sAux12 As String, sAux13 As String
'    Dim sMonto As String * 23, sComision As String * 23, sTot As String * 23
'
'    Dim nLMax As Long
'    Dim nEMax As Long
'    Dim nTMax As Long
'
'    nTMax = 3
'    nEMax = 15
'    nLMax = 45
'    RSet sMonto = Format$(nMontoGiro, "#,##0.00")
'    RSet sComision = Format$(nComision, "#,##0.00")
'    RSet sTot = Format$(nMontoGiro + nComision, "#,##0.00")
'    sCadena = ""
'    sAux1 = "CMACICA - AHORROS"
'    sAux2 = sNomAge & IIf(Mid(sNumGiro, 9, 1) = "1", "- SOLES", "- DOLARES")
'    sAux3 = "Fecha     : " & Format$(dFecha, "dd/mm/yyyy") & "      Hora : " & sHora
'    sAux4 = "REMITENTE : " & sRemitente
'    sAux5 = "DESTINAT. : " & sDestinatario
'    sAux6 = "GIRO NRO. : " & sNumGiro
'    sAux7 = "-------  REGISTRO DE GIROS EFECTIVO  --------"
'    sAux8 = "Monto Giro     : " & IIf(Mid(sNumGiro, 9, 1) = "1", "S/.", "US$") & sMonto
'    sAux9 = "Comisión       : " & IIf(Mid(sNumGiro, 9, 1) = "1", "S/.", "US$") & sComision
'    sAux10 = ""
'    sAux11 = "TOTAL A PAGAR  : " & IIf(Mid(sNumGiro, 9, 1) = "1", "S/.", "US$") & sTot
'    sAux12 = String(44, "-")
'    sAux13 = sCodUser
'
'
'    sCadena = sNegritaOn & sCadena & "." & Space(nTMax - 1) & sAux1 & Space(nLMax - Len(sAux1)) & Space(nEMax) & Chr$(10) 'sAux1 & Chr$(10)
'    sCadena = sCadena & Space(nTMax) & sAux2 & Space(nLMax - Len(sAux2)) & Space(nEMax) & Chr$(10) & sNegritaOff ' sAux2 & Chr$(10) & sNegritaOff
'    sCadena = sCadena & Space(nTMax) & sAux3 & Space(nLMax - Len(sAux3)) & Space(nEMax) & Chr$(10) ' sAux3 & Chr$(10)
'    sCadena = sCadena & Space(nTMax) & sNegritaOn & sAux4 & sNegritaOff & Space(nLMax - Len(sAux4)) & Space(nEMax) & sNegritaOn & Chr$(10) & sNegritaOff 'sAux4 & Chr$(10) & sNegritaOff
'    sCadena = sCadena & Space(nTMax) & sNegritaOn & sAux5 & sNegritaOff & Space(nLMax - Len(sAux5)) & Space(nEMax) & sNegritaOn & Chr$(10) & sNegritaOff 'sAux5 & Chr$(10) & sNegritaOff
'    sCadena = sCadena & Space(nTMax) & sNegritaOn & sAux6 & sNegritaOff & Space(nLMax - Len(sAux6)) & Space(nEMax) & sNegritaOn & Chr$(10) & sNegritaOff 'sAux6 & Chr$(10) & sNegritaOff
'    sCadena = sCadena & Space(nTMax) & sNegritaOn & sAux7 & Space(nLMax - Len(sAux7)) & Space(nEMax) & Chr$(10) & sNegritaOff  ' sAux7 & Chr$(10) & sNegritaOff
'
'    sCadena = sCadena & Space(nTMax) & sAux8 & Space(nLMax - Len(sAux8)) & Space(nEMax) & Chr$(10) ' sAux8 & Chr$(10)
'    sCadena = sCadena & Space(nTMax) & sAux9 & Space(nLMax - Len(sAux9)) & Space(nEMax) & Chr$(10) ' sAux9 & Chr$(10)
'    sCadena = sCadena & Space(nTMax) & sAux10 & Space(nLMax - Len(sAux10)) & Space(nEMax) & Chr$(10) ' sAux10 & Chr$(10)
'    sCadena = sCadena & Space(nTMax) & sNegritaOn & sAux11 & Space(nLMax - Len(sAux11)) & Space(nEMax) & Chr$(10) ' sAux11 & Chr$(10)
'
'    sCadena = sCadena & Space(nTMax) & sAux12 & Space(nLMax - Len(sAux12)) & Space(nEMax) & Chr$(10) ' sAux12 &  Chr$(10)
'    sCadena = sCadena & Space(nTMax) & sAux13 & Space(nLMax - Len(sAux13)) & Space(nEMax) & Chr$(10) ' sAux13 & Chr$(10) & sNegritaOff
'
'    If bCadena Then
'        ImprimeBoletaGiros = sCadena
'    Else
'        Dim ArcSal As Integer
'        ArcSal = FreeFile
'        Open sLpt For Output As ArcSal
'        Print #ArcSal, Chr$(27) & Chr$(64);            'Inicializa Impresora
'        Print #ArcSal, Chr$(18);
'        Print #ArcSal, Chr$(27) & Chr$(50);            'Espaciamiento lineas 1/6 pulg.1
'        Print #ArcSal, Chr$(27) & Chr$(67) & Chr$(22) '   Chr$(nLineas); 'Longitud de página a 66 líneas
'        Print #ArcSal, Chr$(27) & Chr$(108) & Chr$(0); 'Tipo letra : 0,1,2 - Roman,SansS,Courier
'        Print #ArcSal, Chr$(27) & Chr$(103);            'Tamaño  : 80, 77, 103
'        Print #ArcSal, Chr$(27) & Chr$(120) & Chr$(0)  'Draf : 1 pasada
'        Print #ArcSal, sCadena
'        Print #ArcSal, Chr$(12);   'Nueva página
'        Print #ArcSal, Chr$(18);   'Retorna al tipo de letra normal
'        Close ArcSal
'    End If
'End Function



Public Sub ImprimeBoletaConvenio(ByVal sTit As String, ByVal sText As String, ByVal sTextOPe As String, _
            ByVal nMonto As Double, ByVal STextTit As String, ByVal STextCta As String, _
            ByVal sTextRecibo As String, ByVal dFecSis As Date, ByVal sNomAge As String, _
            ByVal sCodUser As String, Optional ByVal psLpt As String = "", Optional psRef As String, Optional psDIRef As String, _
            Optional ByVal pbITFAplica As Boolean = True, Optional ByVal pnITFValor As Double = 0, Optional ByVal pbITFAsumido As Boolean = False, Optional ByVal psItfOperacion As String = "")

                      
Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String, sComision As String, sTotal As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax  As Integer
Dim saux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsCadArg As String
Dim lsInteres As String
Dim lsMensaje As String * 39
Dim oCon As New COMConecta.DCOMConecta
Dim oITF As New COMDConstSistema.FCOMITF

ETIQ:
On Error GoTo Error

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)
nFicSal = FreeFile
Open psLpt For Output As nFicSal
Print #nFicSal, Chr$(27) & Chr$(64);
Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
saux = 5

sMonto = Format$(nMonto, "#,##0.00")
'sComision = Format$(nComision, "#,##0.00")
'sTotal = Format$(nMonto + nComision, "#,##0.00")
sFecha = Format$(dFecSis, "dd/mm/yyyy")
sHora = oCon.GetHoraServer()

Dim clsGen As COMDConstSistema.DCOMGeneral
    Set clsGen = New COMDConstSistema.DCOMGeneral
Dim oImpre As COMFunciones.FCOMImpresion
    Set oImpre = New COMFunciones.FCOMImpresion
'sNumComp = Trim(sNumComp)
    Print #nFicSal, lsNegritaOn; 'Activa Negrita
    Print #nFicSal, Tab(sIni); Trim(Left("CMACICA - AHORRO", 33)); Space(saux + sMax - Len(Trim(Left("CMACICA - AHORRO", 33)))); Space(sSep); Trim(Left("CMACICA - AHORRO", 33)); Space(saux + sMax - Len(Trim(Left("CMACICA - AHORRO", 33))) + 1);
    Print #nFicSal, Tab(sIni); Trim(sNomAge); Space(saux + sMax - Len(Trim(sNomAge)) - Len(Trim(sTextRecibo))); Trim(sTextRecibo); Space(sSep); Trim(sNomAge); Space(saux + sMax - Len(Trim(sNomAge)) - Len(Trim(sTextRecibo)) + 1); Trim(sTextRecibo)
    Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    
    Print #nFicSal, ""
    Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(saux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora
    
    Print #nFicSal, Tab(sIni); "-----------------TITULAR---------------"; Space(-1 + sSep); "-----------------TITULAR---------------"
    Print #nFicSal, Tab(sIni); Trim(Left(STextTit, 39)); Space(39 - Len(Trim(Left(STextTit, 38)))); Space(14); Trim(STextTit)
    Print #nFicSal, Tab(sIni); Trim(STextCta); Space(39 - Len(Trim(STextCta))); Space(14); Trim(STextCta)
    Print #nFicSal, Tab(sIni); "--- PERSONA REALIZA  LA TRANSACCION ---"; Space(-1 + sSep); "--- PERSONA REALIZA  LA TRANSACCION ---"
    Print #nFicSal, Tab(sIni); Trim(Left(psRef, 39)); Space(39 - Len(Trim(Left(psRef, 38)))); Space(14); Trim(Left(psRef, 39))
    Print #nFicSal, Tab(sIni); Trim(Left(psDIRef, 39)); Space(39 - Len(Trim(Left(psDIRef, 38)))); Space(14); Trim(Left(psDIRef, 39))
    
    
    
'    Print #nFicSal, Tab(sIni); sTextRecibo; Space(sMax + 6 - Len(sTextRecibo) - Len(sNumRecibo)); sNumRecibo; Space(-1 + sSep); sTextRecibo; Space(sMax + 6 - Len(sTextRecibo) - Len(sNumRecibo)); sNumRecibo
    Print #nFicSal, ""
    
    sTit = Trim(sTit)
    sTit = oImpre.CentrarCadena(sTit, 29)
    Print #nFicSal, lsNegritaOn; 'Activa Negrita
    Print #nFicSal, Tab(sIni + 1); "-----" & sTit & "-----"; Space(-1 + sSep); "-----" & sTit & "-----"
    Print #nFicSal, Tab(sIni); sText; Space(sMax + 6 - Len(sText) - Len(sMonto)); sMonto; Space(-1 + sSep); sText; Space(sMax + 6 - Len(sText) - Len(sMonto)); sMonto
    'Print #nFicSal, Tab(sIni); sTextCom; Space(sMax + 6 - Len(sTextCom) - Len(sComision)); sComision; Space(-1 + sSep); sTextCom; Space(sMax + 6 - Len(sTextCom) - Len(sComision)); sComision
    Print #nFicSal, ""
    'Print #nFicSal, Tab(sIni); sTextTot; Space(sMax + 6 - Len(sTextTot) - Len(sTotal)); sTotal; Space(-1 + sSep); sTextTot; Space(sMax + 6 - Len(sTextTot) - Len(sTotal)); sTotal
    Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
    Print #nFicSal, Tab(sIni); oImpre.ImpreCarEsp(sCodUser); Space(29 + sSep + saux); oImpre.ImpreCarEsp(sCodUser)
'    Dim clsGen As COMDConstSistema.DCOMGeneral
'    Set clsGen = New COMDConstSistema.DCOMGeneral
    lsMensaje = clsGen.GetMensajeBoletas("")
    Print #nFicSal, Tab(sIni); lsNegritaOn & oImpre.ImpreCarEsp(lsMensaje); Space(-1 + sSep); oImpre.ImpreCarEsp(lsMensaje); lsNegritaOff



For saux = 1 To 5
    Print #nFicSal, Tab(sIni); ""
Next saux
Close nFicSal


 If pbITFAplica And pnITFValor > 0 Then
        If pbITFAsumido Then
            If psItfOperacion = gITFCobroEfectivo Then
                oITF.fgITFImprimeBoleta Trim(STextTit), pnITFValor, sText, dFecSis, , , 1, Mid(STextCta, 9, 1), STextCta, "ASUMIDO POR LA CMAC-ICA", False, psLpt, sNomAge
            Else
                oITF.fgITFImprimeBoleta Trim(STextTit), pnITFValor, sText, dFecSis, , , 2, Mid(STextCta, 9, 1), STextCta, "ASUMIDO POR LA CMAC-ICA", False, psLpt, sNomAge
            End If
        Else
            If psItfOperacion = gITFCobroEfectivo Then
                oITF.fgITFImprimeBoleta Trim(STextTit), pnITFValor, sText, dFecSis, , , 1, Mid(STextCta, 9, 1), STextCta, "", False, psLpt, sNomAge
                
            Else
                oITF.fgITFImprimeBoleta Trim(STextTit), pnITFValor, sText, dFecSis, , , 2, Mid(STextCta, 9, 1), STextCta, "", False, psLpt, sNomAge
            End If
        End If
    End If


Exit Sub
Error:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Error") = vbYes Then
        GoTo ETIQ
    End If
            

End Sub

Public Sub ImprimeBolTarjeta(ByVal sTextOPe As String, _
            ByVal STextTit As String, ByVal STextCta As String, _
            ByVal sTextRecibo As String, ByVal dFecSis As Date, ByVal sNomAge As String, _
            ByVal sCodUser As String, Optional ByVal psLpt As String = "", Optional sRelCuentas As String = "")

          
            
Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String, sComision As String, sTotal As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax  As Integer
Dim saux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsCadArg As String
Dim lsInteres As String
Dim lsMensaje As String * 39

Dim oCone As New COMConecta.DCOMConecta
Dim oImpre As New COMFunciones.FCOMImpresion

ETIQ:
On Error GoTo Error

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)
nFicSal = FreeFile
Open psLpt For Output As nFicSal
Print #nFicSal, Chr$(27) & Chr$(64);
Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
saux = 5

'sMonto = Format$(nMonto, "#,##0.00")
'sComision = Format$(nComision, "#,##0.00")
'sTotal = Format$(nMonto + nComision, "#,##0.00")
sFecha = Format$(dFecSis, "dd/mm/yyyy")
sHora = oCone.GetHoraServer()

Dim clsGen As COMDConstSistema.DCOMGeneral
    Set clsGen = New COMDConstSistema.DCOMGeneral
'sNumComp = Trim(sNumComp)
    Print #nFicSal, lsNegritaOn; 'Activa Negrita
    Print #nFicSal, Tab(sIni); Trim(Left("CMACICA - CAPTACIONES", 33)); Space(saux + sMax - Len(Trim(Left("CMACICA - CAPTACIONES", 33)))); Space(sSep); Trim(Left("CMACICA - CAPTACIONES", 33)); Space(saux + sMax - Len(Trim(Left("CMACICA - CAPTACIONES", 33))) + 1);
    Print #nFicSal, Tab(sIni); Trim(sNomAge); Space(saux + sMax - Len(Trim(sNomAge)) - Len(Trim(sTextRecibo))); Trim(sTextRecibo); Space(sSep); Trim(sNomAge); Space(saux + sMax - Len(Trim(sNomAge)) - Len(Trim(sTextRecibo)) + 1); Trim(sTextRecibo)
    Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    
    Print #nFicSal, ""
    Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(saux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora
    
    Print #nFicSal, Tab(sIni); "-----------------TITULAR---------------"; Space(-1 + sSep); "-----------------TITULAR---------------"
    Print #nFicSal, Tab(sIni); Trim(Left(STextTit, 39)); Space(39 - Len(Trim(Left(STextTit, 38)))); Space(14); Trim(STextTit)
    Print #nFicSal, Tab(sIni); Trim(STextCta); Space(39 - Len(Trim(STextCta))); Space(14); Trim(STextCta)
    Print #nFicSal, ""
    sTextOPe = oImpre.CentrarCadena(Trim(sTextOPe), 33)
    Print #nFicSal, lsNegritaOn; 'Activa Negrita
    Print #nFicSal, Tab(sIni + 1); "---" & sTextOPe & "---"; Space(-1 + sSep); "---" & sTextOPe & "---"
    Print #nFicSal, Tab(sIni); Left(Trim(sRelCuentas), 33); Space(sMax + 6 - Len(Left(Trim(sRelCuentas), 33))); Space(sSep - 1); Left(Trim(sRelCuentas), 33)
    'Print #nFicSal, Tab(sIni); sText; Space(sMax + 6 - Len(sText) - Len(sMonto)); sMonto; Space(-1 + sSep); sText; Space(sMax + 6 - Len(sText) - Len(sMonto)); sMonto
    'Print #nFicSal, Tab(sIni); sTextCom; Space(sMax + 6 - Len(sTextCom) - Len(sComision)); sComision; Space(-1 + sSep); sTextCom; Space(sMax + 6 - Len(sTextCom) - Len(sComision)); sComision
    'If sRelCuentas <> "" Then
    '    Print #nFicSal, ""
    'Else
    '        Print #nFicSal, ""
            Print #nFicSal, ""
    'End If
    'Print #nFicSal, Tab(sIni); sTextTot; Space(sMax + 6 - Len(sTextTot) - Len(sTotal)); sTotal; Space(-1 + sSep); sTextTot; Space(sMax + 6 - Len(sTextTot) - Len(sTotal)); sTotal
    Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
    Print #nFicSal, Tab(sIni); oImpre.ImpreCarEsp(sCodUser); Space(29 + sSep + saux); oImpre.ImpreCarEsp(sCodUser)
    lsMensaje = clsGen.GetMensajeBoletas("")
    Print #nFicSal, Tab(sIni); lsNegritaOn & oImpre.ImpreCarEsp(lsMensaje); Space(-1 + sSep); oImpre.ImpreCarEsp(lsMensaje); lsNegritaOff




For saux = 1 To 5
    Print #nFicSal, Tab(sIni); ""
Next saux
Close nFicSal
 
Exit Sub
Error:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Error") = vbYes Then
        GoTo ETIQ
    End If
            

End Sub


Public Sub ImprimeBoleta(ByVal sTit As String, ByVal sText As String, ByVal sTextCom As String, _
            ByVal sTextTot As String, ByVal sCodOpe As String, ByVal nMonto As Double, nComision As Double, _
            ByVal sTextRecibo As String, ByVal sNumRecibo As String, _
            ByVal dFecSis As Date, ByVal sNomCmac As String, ByVal sNomAge As String, _
            ByVal sCodUser As String, Optional ByVal sLpt As String = "", Optional ByVal psCodCmac As String)

Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String, sComision As String, sTotal As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax  As Integer
Dim saux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsCadArg As String
Dim lsInteres As String
Dim lsMensaje As String * 39

Dim oImpre As New COMFunciones.FCOMImpresion

ETIQ:
On Error GoTo Error

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)
nFicSal = FreeFile
Open sLpt For Output As nFicSal
Print #nFicSal, Chr$(27) & Chr$(64);
Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
saux = 5

sMonto = Format$(nMonto, "#,##0.00")
sComision = Format$(nComision, "#,##0.00")
sTotal = Format$(nMonto + nComision, "#,##0.00")
sFecha = Format$(dFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")

Dim clsGen As COMDConstSistema.DCOMGeneral
    Set clsGen = New COMDConstSistema.DCOMGeneral
'sNumComp = Trim(sNumComp)
If psCodCmac = "102" Then  '' Lima
    Print #nFicSal, lsNegritaOn; 'Activa Negrita
    Print #nFicSal, Tab(sIni); Trim(Left(sNomCmac, 33)); Space(saux + sMax - Len(Trim(Left(sNomCmac, 33))));
    Print #nFicSal, Tab(sIni); Trim(sNomAge); Space(saux + sMax - Len(Trim(sNomAge)));
    Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    
    Print #nFicSal, ""
    Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora;
    Print #nFicSal, ""
    Print #nFicSal, Tab(sIni); sTextRecibo; Space(sMax + 6 - Len(sTextRecibo) - Len(sNumRecibo)); sNumRecibo;
    
    sTit = Trim(sTit)
    sTit = oImpre.CentrarCadena(sTit, 23)
    Print #nFicSal, lsNegritaOn; 'Activa Negrita
    Print #nFicSal, Tab(sIni + 1); "-------" & sTit & "--------";
    Print #nFicSal, Tab(sIni); sText; Space(sMax + 6 - Len(sText) - Len(sMonto)); sMonto;
    Print #nFicSal, Tab(sIni); sTextCom; Space(sMax + 6 - Len(sTextCom) - Len(sComision)); sComision; Space(-1 + sSep);
    Print #nFicSal, ""
    Print #nFicSal, Tab(sIni); sTextTot; Space(sMax + 6 - Len(sTextTot) - Len(sTotal)); sTotal;
    Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    Print #nFicSal, Tab(sIni); "---------------------------------------";
    Print #nFicSal, Tab(sIni); oImpre.ImpreCarEsp(sCodUser); Space(29 + sSep + saux);
'    Dim clsGen As COMDConstSistema.DCOMGeneral
'    Set clsGen = New COMDConstSistema.DCOMGeneral
    lsMensaje = clsGen.GetMensajeBoletas("")
    Print #nFicSal, Tab(sIni); lsNegritaOn & oImpre.ImpreCarEsp(lsMensaje); lsNegritaOff
Else
    Print #nFicSal, lsNegritaOn; 'Activa Negrita
    Print #nFicSal, Tab(sIni); Trim(Left(sNomCmac, 33)); Space(saux + sMax - Len(Trim(Left(sNomCmac, 33)))); Space(sSep); Trim(Left(sNomCmac, 33)); Space(saux + sMax - Len(Trim(Left(sNomCmac, 33))) + 1);
    Print #nFicSal, Tab(sIni); Trim(sNomAge); Space(saux + sMax - Len(Trim(sNomAge))); Space(sSep); Trim(sNomAge); Space(saux + sMax - Len(Trim(sNomAge)) + 1);
    Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    
    Print #nFicSal, ""
    Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(saux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora
    Print #nFicSal, ""
    Print #nFicSal, Tab(sIni); sTextRecibo; Space(sMax + 6 - Len(sTextRecibo) - Len(sNumRecibo)); sNumRecibo; Space(-1 + sSep); sTextRecibo; Space(sMax + 6 - Len(sTextRecibo) - Len(sNumRecibo)); sNumRecibo
    
    sTit = Trim(sTit)
    sTit = oImpre.CentrarCadena(sTit, 23)
    Print #nFicSal, lsNegritaOn; 'Activa Negrita
    Print #nFicSal, Tab(sIni + 1); "-------" & sTit & "--------"; Space(-1 + sSep); "----------" & sTit & "-----------"
    Print #nFicSal, Tab(sIni); sText; Space(sMax + 6 - Len(sText) - Len(sMonto)); sMonto; Space(-1 + sSep); sText; Space(sMax + 6 - Len(sText) - Len(sMonto)); sMonto
    Print #nFicSal, Tab(sIni); sTextCom; Space(sMax + 6 - Len(sTextCom) - Len(sComision)); sComision; Space(-1 + sSep); sTextCom; Space(sMax + 6 - Len(sTextCom) - Len(sComision)); sComision
    Print #nFicSal, ""
    Print #nFicSal, Tab(sIni); sTextTot; Space(sMax + 6 - Len(sTextTot) - Len(sTotal)); sTotal; Space(-1 + sSep); sTextTot; Space(sMax + 6 - Len(sTextTot) - Len(sTotal)); sTotal
    Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
    Print #nFicSal, Tab(sIni); oImpre.ImpreCarEsp(sCodUser); Space(29 + sSep + saux); oImpre.ImpreCarEsp(sCodUser)
'    Dim clsGen As COMDConstSistema.DCOMGeneral
'    Set clsGen = New COMDConstSistema.DCOMGeneral
    lsMensaje = clsGen.GetMensajeBoletas("")
    Print #nFicSal, Tab(sIni); lsNegritaOn & oImpre.ImpreCarEsp(lsMensaje); Space(-1 + sSep); oImpre.ImpreCarEsp(lsMensaje); lsNegritaOff

End If
For saux = 1 To 5
    Print #nFicSal, Tab(sIni); ""
Next saux
Close nFicSal
Exit Sub
Error:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Error") = vbYes Then
        GoTo ETIQ
    End If
End Sub

Public Function GetServicioParametros(ByVal nServicio As CaptacInstServicios) As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetServicioParametros = clsServ.GetServicioParametros(nServicio)
Set clsServ = Nothing
End Function

Public Sub ActualizaServicioParametro(ByVal nServicio As CaptacInstServicios, _
        ByVal sCuenta As String, ByVal nComision As Double, ByVal nTipoComision As CapServTipoComision)
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Dim bTrans As Boolean
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
On Error GoTo ErrGraba
clsServ.dbCmact.BeginTrans
bTrans = True
clsServ.ActualizaServicioParametro nServicio, sCuenta, nComision, nTipoComision
clsServ.dbCmact.CommitTrans
bTrans = False
Exit Sub
ErrGraba:
    If bTrans Then clsServ.dbCmact.RollbackTrans
    Set clsServ = Nothing
    Err.Raise Err.Number, "", Err.Description
End Sub

Public Sub AgregaServicioParametro(ByVal nServicio As CaptacInstServicios, _
        ByVal sCuenta As String, ByVal nComision As Double, ByVal nTipoComision As CapServTipoComision)

Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Dim bTrans As Boolean
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
On Error GoTo ErrGraba
clsServ.dbCmact.BeginTrans
bTrans = True
clsServ.AgregaServicioParametro nServicio, sCuenta, nComision, nTipoComision
clsServ.dbCmact.CommitTrans
bTrans = False
Exit Sub
ErrGraba:
    If bTrans Then clsServ.dbCmact.RollbackTrans
    Set clsServ = Nothing
    Err.Raise Err.Number, "", Err.Description
End Sub


Public Function EsReciboCobrado(ByVal nOperacion As CaptacOperacion, ByVal sNumRecibo As String) As Boolean

Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
EsReciboCobrado = clsServ.EsReciboCobrado(nOperacion, sNumRecibo)
Set clsServ = Nothing
End Function

Public Function GetMovServicios(ByVal nServicio As CaptacOperacion, _
        ByVal dInicio As Date, ByVal dFin As Date, _
        Optional bReciboYaAbono As Boolean = False) As ADODB.Recordset

Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetMovServicios = clsServ.GetMovServicios(nServicio, dInicio, dFin, bReciboYaAbono)
Set clsServ = Nothing
End Function

' ACTUALIZACION - CMCPL
Public Sub CapCobranzaServicios(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
            ByVal sNumRecibo As String, ByVal sCodigo As String, ByVal nMonto As Double, _
            ByVal sNomCmac As String, ByVal sNomAge As String, ByVal sLpt As String, _
            Optional nComision As Double = 0, Optional nmoneda As Moneda = gMonedaNacional, _
            Optional nGastos As Double = 0, Optional nCostas As Double = 0, Optional nDerEmis As Double = 0, _
            Optional ByVal psCodCmac As String)
Dim dFecSis As Date
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Dim bTrans As Boolean
Dim nMovNro As Long
Dim sServicio As String, sGlosa As String, sTexto As String
Dim sReciboImp As String
Select Case nOperacion
    Case gServCobSedalib
        sServicio = "SEDALIB"
        sTexto = "Recibo : "
        sReciboImp = Format$(sNumRecibo, "###-########-##")
        sGlosa = "Cobranza " & sServicio & " Recibo N° = " & sNumRecibo
    Case gServCobHidrandina
        sServicio = "HIDRANDINA"
        sTexto = "Recibo : "
        sReciboImp = Format$(sNumRecibo, "###-##-##-####")
        sGlosa = "Cobranza " & sServicio & " Recibo N° = " & sNumRecibo
    Case gServCobEdelnor
        sServicio = "EDELNOR"
        sTexto = "Suministro : "
        sReciboImp = sCodigo
        sGlosa = "Cobranza " & sServicio & " Suministro N° = " & sCodigo
    Case gServCobSATTInfraccion 'SAT - INFRANCCIONES
         sServicio = "SAT"
         sTexto = "Recibo: "
         sReciboImp = sCodigo
         sGlosa = "Cobranza " & sServicio & " Recibo N° = " & sCodigo
    Case gServCobSATTReciboDerecho 'SAT - RECIBO/DERECHO
         sServicio = "Tributo/Derecho"
         sTexto = "Recibo: "
         sReciboImp = sCodigo
         sGlosa = "Cobranza " & sServicio & " Recibo N° = " & sCodigo
End Select

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
On Error GoTo ErrGraba
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
clsServ.dbCmact.BeginTrans
bTrans = True
clsServ.AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = clsServ.GetnMovNro(sMovNro)
clsServ.AgregaMovServicios nMovNro, sNumRecibo, sCodigo, nMonto + nComision + nDerEmis + nCostas + nGastos
If nMonto > 0 Then 'CAPITAL
    clsServ.AgregaMovServiciosDet nMovNro, sNumRecibo, gServConcImporte, nMonto
End If
If nComision > 0 Then 'COMISION
    clsServ.AgregaMovServiciosDet nMovNro, sNumRecibo, gServConcComision, nComision
End If
If nGastos > 0 Then 'GASTOS ADMINISTRATIVOS
    clsServ.AgregaMovServiciosDet nMovNro, sNumRecibo, gServConcGastos, nGastos
End If
If nCostas > 0 Then 'COSTAS
    clsServ.AgregaMovServiciosDet nMovNro, sNumRecibo, gServConcCostas, nCostas
End If
If nDerEmis > 0 Then 'DERECHO DE EMISION
    clsServ.AgregaMovServiciosDet nMovNro, sNumRecibo, gServConcDerEmis, nDerEmis
End If
clsServ.dbCmact.CommitTrans
bTrans = False
Do
    ImprimeBoleta "COBRANZA " & sServicio, "Monto Recibo", "Comision", "Total a Pagar", Trim(nOperacion), _
                            nMonto + nDerEmis, nComision, sTexto, sReciboImp, dFecSis, sNomCmac, sNomAge, _
                            Right(sMovNro, 4), sLpt, psCodCmac
Loop Until MsgBox("¿Desea reimprimir la boleta?", vbQuestion + vbYesNo, "Aviso") = vbNo
Exit Sub
ErrGraba:
    If bTrans Then clsServ.dbCmact.RollbackTrans
    Set clsServ = Nothing
    Err.Raise Err.Number, "CapCobranzaServicios", Err.Description
End Sub

Private Function GetFormatoFile(ByVal dFechaCobro As Date) As String
Dim vMes As String
Dim vNumMes As String
Dim vYear As String
Dim vNumYear As String
Dim vDia As String
Dim vNom_cadena1 As String
Dim cFecha As String
cFecha = Format$(dFechaCobro, "dd/mm/yyyy")
vMes = Mid(cFecha, 4, 2)
vYear = Mid(cFecha, 7, 4)
vDia = Mid(cFecha, 1, 2)

Select Case vMes
    Case "01"
        vNumMes = "a"
    Case "02"
        vNumMes = "b"
    Case "03"
        vNumMes = "c"
    Case "04"
        vNumMes = "d"
    Case "05"
        vNumMes = "e"
    Case "06"
        vNumMes = "f"
    Case "07"
        vNumMes = "g"
    Case "08"
        vNumMes = "h"
    Case "09"
        vNumMes = "i"
    Case "10"
        vNumMes = "j"
    Case "11"
        vNumMes = "k"
    Case "12"
        vNumMes = "l"
End Select

Select Case vYear
    Case "1999"
        vNumYear = "c"
    Case "2000"
        vNumYear = "d"
    Case "2001"
        vNumYear = "e"
    Case "2002"
        vNumYear = "f"
    Case "2003"
        vNumYear = "g"
    Case "2004"
        vNumYear = "h"
    Case "2005"
        vNumYear = "i"
    Case "2006"
        vNumYear = "j"
    Case "2007"
        vNumYear = "k"
    Case "2008"
        vNumYear = "l"
    Case "2009"
        vNumYear = "m"
    Case "2010"
        vNumYear = "n"
End Select
GetFormatoFile = "P050" & vDia & vNumMes & vNumYear
End Function

Public Sub GeneraDBFSedalib(ByVal dFechaCobro As Date)

Dim rsServ As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Dim sNumRecibo As String, sCodigo As String, sError As String
Dim nMonto As Double
Dim nMovNro As Long
Dim sNombreArchivo As String, sAgencia As String, sUsuario As String
Dim sHora As String

On Error GoTo ControlError
Screen.MousePointer = 11
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios

'Set rsServ = clsServ.GetMovServicios(gServCobSedalib, dFechaCobro)  ROCHABUS
If Not (rsServ.EOF And rsServ.BOF) Then
    clsServ.dbCmact.BeginTrans
    Do While Not rsServ.EOF
        sNumRecibo = rsServ("cNumRecibo")
        sCodigo = rsServ("cCodigo")
        nMonto = rsServ("nMonto")
        sAgencia = Mid(rsServ("cMovNro"), 15, 5)
        sUsuario = Right(rsServ("cMovNro"), 4)
        nMovNro = rsServ("nMovNro")
        If Len(Trim(sNumRecibo)) < 11 Then
            sCodigo = Left(sCodigo, Len(RTrim(sCodigo))) & Space(11 - Len(RTrim(sCodigo)))
        End If
        If Not ValSedalib(sNumRecibo, sCodigo, nMonto) Then
            MsgBox "Códigos Errados, Verifique El Recibo   " & sNumRecibo & "      Código      " & sCodigo & "       Monto    " & nMonto & "  Age   " & sAgencia & "  User  " & sUsuario, vbInformation, "Aviso"       'Actualiza el campo Cflag con la Letra Y si esta errado
            clsServ.ActualizaFlagServicios nMovNro, gCapServFlagNoValido
        Else
            clsServ.ActualizaFlagServicios nMovNro, gCapServFlagProcesado
        End If
        rsServ.MoveNext
    Loop
    clsServ.dbCmact.CommitTrans
    rsServ.MoveFirst
    sNombreArchivo = GetFormatoFile(dFechaCobro)
    clsServ.CreaTablaDBFSedalib sNombreArchivo
    Do While Not rsServ.EOF
        sNumRecibo = rsServ("cNumRecibo")
        sCodigo = rsServ("cCodigo")
        nMonto = rsServ("nMonto")
        sHora = Mid(rsServ("cMovNro"), 9, 2) & ":" & Mid(rsServ("cMovNro"), 11, 2) & ":" & Mid(rsServ("cMovNro"), 13, 2)
        clsServ.AgregaServDBF sNombreArchivo, sNumRecibo, sCodigo, nMonto, dFechaCobro, sHora
        rsServ.MoveNext
    Loop
    Screen.MousePointer = 0
    MsgBox "Archivo " & sNombreArchivo & " generado correctamente", vbInformation, "Aviso"
Else
    Screen.MousePointer = 0
    MsgBox "No existen datos para la fecha indicada", vbInformation, "Aviso"
End If
rsServ.Close
Set rsServ = Nothing
Set clsServ = Nothing
Exit Sub
ControlError:
    Err.Raise Err.Number, "", Err.Description
    Screen.MousePointer = 0
End Sub

Public Function GeneraReporteServicios(ByVal rsServ As ADODB.Recordset, ByVal nServicio As CaptacInstServicios, _
                ByVal nNumRecibo As Long, ByVal nMontoTotal As Double, _
                ByVal nValorComision As Double, ByVal nMontoComision As Double, _
                ByVal nTipoComision As CapServTipoComision, ByVal dInicio As Date, ByVal dFin As Date, _
                ByVal dFecSis As Date, ByVal sNomCmac As String) As String

Dim nItem As Long
Dim nLinPag As Integer
Dim nCntPag As Long
Dim nImporteDep As Double
Dim nMonto As Double
Dim sCad As String, sItem As String * 4
Dim sRecibo As String * 13, sCodigo As String * 11
Dim sMonto As String * 12
Dim sTitRp1 As String, sTitRp2 As String, sNumPag As String
Dim nCarLin As Integer
Dim sFechaHora As String * 19
Dim sAgencia As String * 15, sUsuario As String * 4
Dim sMoneda As String

Dim oFun As New COMFunciones.FCOMCadenas
Dim oImpreCap As New COMNCaptaGenerales.NCOMCaptaImpresion

sCad = ""
sNumPag = ""
nLinPag = 0
nCarLin = 90
sMoneda = "SOLES"

If nServicio = gCapServEdelnor Then
    sTitRp1 = "COBRANZA SERVICIO EDELNOR"
ElseIf nServicio = gCapServSedalib Then
    sTitRp1 = "COBRANZA SERVICIO SEDALIB"
ElseIf nServicio = gCapServHidrandina Then
    sTitRp1 = "COBRANZA SERVICIO HIDRANDINA"
End If
sTitRp2 = "DE " & Format$(dInicio, "dd/mm/yyyy") & " AL " & Format$(dFin, "dd/mm/yyyy")

rsServ.MoveFirst
If Not (rsServ.EOF And rsServ.BOF) Then
    nCntPag = 1
    sNumPag = oFun.FillNum(Trim(Str(nCntPag)), 4, " ")
    sCad = sCad & sNomCmac & Chr$(10)
    sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "SECCION OPERACIONES", sTitRp1, sTitRp2, sMoneda, sNumPag, "", dFecSis) & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & String(nCarLin, "-") & Chr$(10)
    sCad = sCad & "ITEM  FECHA                RECIBO         CODIGO              MONTO  AGENCIA          USUA" & Chr$(10)
    sCad = sCad & String(nCarLin, "-") & Chr$(10)
    sCad = sCad & Chr$(10)
    nLinPag = 7
    nItem = 0
    Do While Not rsServ.EOF
        sFechaHora = rsServ("Fecha")
        sUsuario = rsServ("Usuario")
        sAgencia = rsServ("Agencia")
        sCodigo = rsServ("Codigo")
        sRecibo = rsServ("Recibo")
        RSet sMonto = Format$(rsServ("Monto"), "#,##0.00")
        nItem = nItem + 1
        RSet sItem = nItem
        
        sCad = sCad & sItem & Space(2) & sFechaHora & Space(2) & sRecibo & Space(2) & sCodigo & Space(2) & sMonto & Space(2) & sAgencia & Space(2) & sUsuario & Chr$(10)
                
        If nLinPag > 62 Then
            sCad = sCad & Chr$(12)
            nCntPag = nCntPag + 1
            sNumPag = oFun.FillNum(Trim(Str(nCntPag)), 4, " ")
            sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "SECCION OPERACIONES", sTitRp1, sTitRp2, sMoneda, sNumPag, "", dFecSis) & Chr$(10)
            sCad = sCad & Chr$(10)
            sCad = sCad & String(nCarLin, "-") & Chr$(10)
            sCad = sCad & "ITEM  FECHA                RECIBO         CODIGO              MONTO  AGENCIA          USUA" & Chr$(10)
            sCad = sCad & String(nCarLin, "-") & Chr$(10)
            sCad = sCad & Chr$(10)
            nLinPag = 7
        End If
        rsServ.MoveNext
    Loop
    sCad = sCad & String(nCarLin, "-") & Chr$(10)
    
    
    If nLinPag + 10 > 62 Then
        sCad = sCad & Chr$(12)
        nCntPag = nCntPag + 1
        sNumPag = oFun.FillNum(Trim(Str(nCntPag)), 4, " ")
        sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "SECCION OPERACIONES", sTitRp1, sTitRp2, sMoneda, sNumPag, "", dFecSis) & Chr$(10)
        sCad = sCad & Chr$(10)
        sCad = sCad & String(nCarLin, "-") & Chr$(10)
        sCad = sCad & "ITEM  FECHA                RECIBO         CODIGO              MONTO  AGENCIA          USUA" & Chr$(10)
        sCad = sCad & String(nCarLin, "-") & Chr$(10)
        sCad = sCad & Chr$(10)
        nLinPag = 7
    End If
    sCad = sCad & Chr$(10)
    RSet sMonto = Format$(nNumRecibo, "#,##0.00")
    sCad = sCad & Space(5) & "Recibos           = " & sMonto & Chr$(10)
    RSet sMonto = Format$(nValorComision, "#,##0.0000")
    
    If nTipoComision = gCapServTpoComMontoRecibo Then
        sCad = sCad & Space(5) & "Comisión x Recibo = " & sMonto & Chr$(10)
    ElseIf nTipoComision = gCapServTpoComPorcentaje Then
        sCad = sCad & Space(5) & "Comisión (%)      = " & sMonto & Chr$(10)
    End If
    RSet sMonto = Format$(nMontoComision, "#,##0.00")
    sCad = sCad & Space(5) & "Total Comisión    = " & sMonto & Chr$(10)
    RSet sMonto = Format$(nMontoTotal, "#,##0.00")
    sCad = sCad & Space(5) & "Importe a Abonar  = " & sMonto & Chr$(10)
    Screen.MousePointer = 0
End If
GeneraReporteServicios = sCad
End Function

Public Function GeneraReporteHidrandina(ByVal dFechaCobro As Date, ByVal dFecSis As Date) As String
Dim nLineas As Long
Dim nPage As Long, nPageLin As Long
Dim nNumRecibo As Long
Dim nValorComision As Double, nMontoTotal As Double, nImporteDep As Double
Dim nMonto As Double
Dim sCad As String
Dim rsServ As ADODB.Recordset, rsParam As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Dim sCodigo As String * 8, sRecibo As String * 12

Dim oImpre As New COMFunciones.FCOMImpresion
Dim oFun As New COMFunciones.FCOMCadenas

sCad = ""
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
'Set rsServ = clsServ.GetMovServicios(gServCobHidrandina, dFechaCobro)  ROCHABUS
If Not (rsServ.EOF And rsServ.BOF) Then
    nPage = 1
    sCad = sCad & Space(8) & oImpre.PrnSet("B+") & "CAJA MUNICIPAL DE TRUJILLO" & Space(62) & "Página: " & nPage & oImpre.PrnSet("B-") & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & Space(40) & oImpre.PrnSet("B+") & oImpre.PrnSet("RM") & "PLANILLA DE COBRANZA HIDRANDINA DEL " & Format(dFechaCobro, "dd/mm/yyyy") & oImpre.PrnSet("B-") & Chr$(10)
    sCad = sCad & Space(40) & oImpre.PrnSet("B+") & oImpre.PrnSet("RM") & "________________________________________________" & oImpre.PrnSet("B-") & Chr$(10)
    sCad = sCad & Space(94) & " Fecha : " & Format(dFecSis, "dd/mm/yyyy") & Chr$(10)
    sCad = sCad & Space(94) & " Hora  : " & Format(Time, "hh:mm:ss") & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & Space(9) & String(50, "-") & Space(4) & String(48, "-") & Chr$(10)
    sCad = sCad & Space(9) & oImpre.PrnSet("B+") & "|   Item | Numero de   |        Recibo           |" & Space(4) & "| Item | Numero de   |        Recibo            |" & oImpre.PrnSet("B-") & Chr$(10)
    sCad = sCad & Space(9) & oImpre.PrnSet("B+") & "|        |  Recibo     |   Suministro |Importe   |" & Space(4) & "|      |  Recibo     |  Suministro  |Importe    |" & oImpre.PrnSet("B-") & Chr$(10)
    sCad = sCad & Space(9) & String(50, "-") & Space(4) & String(48, "-") & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & Space(10) & oImpre.PrnSet("B+") & oImpre.PrnSet("U+") & "  Cobranza del día :" & Format(dFechaCobro, "dd/mm/yyyy") & oImpre.PrnSet("B-") & oImpre.PrnSet("U-") & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & Chr$(10)
    nPageLin = 16
    nLineas = 7
    nNumRecibo = 0
    nMontoTotal = 0
    
    Do While Not rsServ.EOF
        sCodigo = rsServ("cCodigo")
        sRecibo = Format$(Left(rsServ("cNumRecibo"), 9), "###-##-##-##")
        nMonto = rsServ("nMonto")
        nNumRecibo = nNumRecibo + 1
        nMontoTotal = nMontoTotal + nMonto
        If (nNumRecibo Mod 2) = 0 Then
            sCad = sCad & oFun.JDNum(Trim(nNumRecibo), 10, False, 10, 0) & Space(4) & sRecibo & Space(4) & sCodigo & Space(3) & oFun.JDNum(Trim(nMonto), 8, True, 5, 2) & Chr$(10)
            nPageLin = nPageLin + 1
        Else
            sCad = sCad & oFun.JDNum(Trim(nNumRecibo), 21, False, 21, 0) & Space(4) & sRecibo & Space(4) & sCodigo & Space(3) & oFun.JDNum(Trim(nMonto), 8, True, 5, 2)
        End If
        If nPageLin > 62 Then
            sCad = sCad & Chr$(12)
            nPage = nPage + 1
            sCad = sCad & Space(8) & oImpre.PrnSet("B+") & "CAJA MUNICIPAL DE TRUJILLO" & Space(62) & "Página: " & nPage & oImpre.PrnSet("B-") & Chr$(10)
            sCad = sCad & Chr$(10)
            sCad = sCad & Space(40) & oImpre.PrnSet("B+") & oImpre.PrnSet("RM") & "PLANILLA DE COBRANZA HIDRANDINA DEL " & Format(dFechaCobro, "dd/mm/yyyy") & oImpre.PrnSet("B-") & Chr$(10)
            sCad = sCad & Space(40) & oImpre.PrnSet("B+") & oImpre.PrnSet("RM") & "________________________________________________" & oImpre.PrnSet("B-") & Chr$(10)
            sCad = sCad & Space(94) & " Fecha : " & Format(dFecSis, "dd/mm/yyyy") & Chr$(10)
            sCad = sCad & Space(94) & " Hora  : " & Format(Time, "hh:mm:ss") & Chr$(10)
            sCad = sCad & Chr$(10)
            sCad = sCad & Space(9) & String(50, "-") & Space(4) & String(48, "-") & Chr$(10)
            sCad = sCad & Space(9) & oImpre.PrnSet("B+") & "|   Item | Numero de   |        Recibo           |" & Space(4) & "| Item | Numero de   |        Recibo            |" & oImpre.PrnSet("B-") & Chr$(10)
            sCad = sCad & Space(9) & oImpre.PrnSet("B+") & "|        |  Recibo     |   Suministro |Importe   |" & Space(4) & "|      |  Recibo     |  Suministro  |Importe    |" & oImpre.PrnSet("B-") & Chr$(10)
            sCad = sCad & Space(9) & String(50, "-") & Space(4) & String(48, "-") & Chr$(10)
            sCad = sCad & Chr$(10)
            sCad = sCad & Space(10) & oImpre.PrnSet("B+") & oImpre.PrnSet("U+") & "  Cobranza del día :" & Format(dFechaCobro, "dd/mm/yyyy") & oImpre.PrnSet("B-") & oImpre.PrnSet("U-") & Chr$(10)
            sCad = sCad & Chr$(10)
            sCad = sCad & Chr$(10)
            nPageLin = 16
            nLineas = 7
        End If
        rsServ.MoveNext
    Loop
        
    Set rsParam = clsServ.GetServicioParametros(gCapServHidrandina)
    nValorComision = rsParam("nComision")
    Set clsServ = Nothing
    rsParam.Close
    Set rsParam = Nothing
    nImporteDep = nValorComision * nNumRecibo
    sCad = sCad & Chr$(10) & Chr$(10)
    sCad = sCad & Space(6) & oImpre.PrnSet("B+") & "SubT. Agencia [50]             ====>    " & oFun.JDNum(Trim(nNumRecibo), 5, False, 5, 0) & Space(4) & oFun.JDNum(Trim(nMontoTotal), 7, True, 4, 2) & oImpre.PrnSet("B-") & Chr$(10)
    sCad = sCad & Chr$(10)
    nPageLin = nPageLin + 10
    If nPageLin > 62 Then
        sCad = sCad & Chr$(12)
        nPage = nPage + 1
        sCad = sCad & Space(8) & oImpre.PrnSet("B+") & "CAJA MUNICIPAL DE TRUJILLO" & Space(62) & "Página: " & nPage & oImpre.PrnSet("B-") & Chr$(10)
        sCad = sCad & Chr$(10)
        sCad = sCad & Space(40) & oImpre.PrnSet("B+") & oImpre.PrnSet("RM") & "PLANILLA DE COBRANZA HIDRANDINA DEL " & Format(dFechaCobro, "dd/mm/yyyy") & oImpre.PrnSet("B-") & Chr$(10)
        sCad = sCad & Space(40) & oImpre.PrnSet("B+") & oImpre.PrnSet("RM") & "________________________________________________" & oImpre.PrnSet("B-") & Chr$(10)
        sCad = sCad & Space(94) & " Fecha : " & Format(dFecSis, "dd/mm/yyyy") & Chr$(10)
        sCad = sCad & Space(94) & " Hora  : " & Format(Time, "hh:mm:ss") & Chr$(10)
        sCad = sCad & Chr$(10)
        sCad = sCad & Space(9) & String(50, "-") & Space(4) & String(48, "-") & Chr$(10)
        sCad = sCad & Space(9) & oImpre.PrnSet("B+") & "|   Item | Numero de   |        Recibo           |" & Space(4) & "| Item | Numero de   |        Recibo            |" & oImpre.PrnSet("B-") & Chr$(10)
        sCad = sCad & Space(9) & oImpre.PrnSet("B+") & "|        |  Recibo     |   Suministro |Importe   |" & Space(4) & "|      |  Recibo     |  Suministro  |Importe    |" & oImpre.PrnSet("B-") & Chr$(10)
        sCad = sCad & Space(9) & String(50, "-") & Space(4) & String(48, "-") & Chr$(10)
        sCad = sCad & Chr$(10)
        sCad = sCad & Space(10) & oImpre.PrnSet("B+") & oImpre.PrnSet("U+") & "  Cobranza del día :" & Format(dFechaCobro, "dd/mm/yyyy") & oImpre.PrnSet("B-") & oImpre.PrnSet("U-") & Chr$(10)
        sCad = sCad & Chr$(10)
        sCad = sCad & Chr$(10)
    End If
    sCad = sCad & Space(8) & "...................................................." & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & Space(8) & oImpre.PrnSet("B+") & "                   Comisión dia " & Format(dFechaCobro, "dd/mm/yyyy") & oImpre.PrnSet("B-") & Chr$(10)
    sCad = sCad & "                                Recibos  =" & oFun.JDNum(Trim(nNumRecibo), 9, False, 9, 0) & Chr$(10)
    sCad = sCad & "                     Comisión x Recibo   =" & oFun.JDNum(Trim(nValorComision), 9, False, 6, 2) & Chr$(10)
    sCad = sCad & "                     Total Comisión día  =" & oFun.JDNum(Trim(nImporteDep), 9, False, 6, 2) & Chr$(10)
    sCad = sCad & "                     Importe a Depositar =" & oFun.JDNum(Trim(nMontoTotal), 9, False, 6, 2) & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & Space(8) & "________________________________________________________________" & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & Space(10) & "SubT  Fecha [" & Format(dFechaCobro, "dd/mm/yyyy") & "] ===>   " & oFun.JDNum(Trim(nNumRecibo), 5, False, 5, 0) & Space(4) & oFun.JDNum(Trim(nMontoTotal), 7, True, 4, 2) & Chr$(10)
    sCad = sCad & Chr$(10)
    sCad = sCad & Space(8) & "======================================================================================================" & Chr$(10)
    sCad = sCad & Space(3) & "       TOTAL REGISTRADO                    " & oFun.JDNum(Trim(nNumRecibo), 5, False, 5, 0) & Space(4) & oFun.JDNum(Trim(nMontoTotal), 7, True, 4, 2) & Chr$(10)
    Screen.MousePointer = 0
End If
GeneraReporteHidrandina = sCad
End Function

Public Function GetServConvenios(Optional ByVal nTipoCOnv As Integer) As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetServConvenios = clsServ.GetServConvenios(nTipoCOnv)

Set clsServ = Nothing
End Function




Public Function GetServConvenios2() As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetServConvenios2 = clsServ.GetServConvenios2()

Set clsServ = Nothing
End Function

Public Function GetServConveniosArbol() As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetServConveniosArbol = clsServ.GetServConveniosArbol()
Set clsServ = Nothing
End Function

Public Function GetServConveniosCodArbol() As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetServConveniosCodArbol = clsServ.GetServConveniosCodArbol()
Set clsServ = Nothing
End Function

Public Sub ActualizaServConvenios(ByVal rsServ As ADODB.Recordset)
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Dim nConvenio As CaptacConvenios
Dim bTrans As Boolean
Dim sFlag As String
On Error GoTo ErrGraba
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
clsServ.dbCmact.BeginTrans
bTrans = True
Do While Not rsServ.EOF
    sFlag = rsServ("Flag")
    nConvenio = 101 'CLng(Trim(Right(rsServ("Tipo"), 4)))
    If sFlag = "N" Then
        
        clsServ.AgregaServConvenio rsServ("Codigo"), nConvenio
    Else
        clsServ.ActualizaServConvenio rsServ("Codigo"), nConvenio
        
    End If
    rsServ.MoveNext
Loop
clsServ.dbCmact.CommitTrans
bTrans = False
Exit Sub
ErrGraba:
    If bTrans Then clsServ.dbCmact.RollbackTrans
    Set clsServ = Nothing
    MsgBox Err.Description, vbCritical, "Error"
End Sub
Public Sub EliminaServConvenios(ByVal sPersona As String)
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Dim bTrans As Boolean
On Error GoTo ErrGraba
clsServ.dbCmact.BeginTrans

  clsServ.EliminaServConvCuentas sPersona
  
  clsServ.EliminaServConvenio sPersona

clsServ.dbCmact.CommitTrans
bTrans = False
Exit Sub
ErrGraba:
    If bTrans Then clsServ.dbCmact.RollbackTrans
    Set clsServ = Nothing
    MsgBox Err.Description, vbCritical, "Error"

End Sub


Public Sub ActualizaServConvCuentas(ByVal sPersona As String, ByVal rsServ As ADODB.Recordset)
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Dim nTipoCuenta As CaptacConvTipoCuenta
Dim bTrans As Boolean
Dim sCuenta As String, nMontoMora As Currency
'On Error GoTo ErrGraba
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
clsServ.dbCmact.BeginTrans
bTrans = True
clsServ.EliminaServConvCuentas sPersona

Do While Not rsServ.EOF
    sCuenta = rsServ("cuenta")
    nMontoMora = rsServ("monto")
    nTipoCuenta = CLng(Trim(Left(rsServ("Tipo Cuenta"), 1)))
    clsServ.AgregaServConvenioCuenta sPersona, sCuenta, nTipoCuenta, nMontoMora
    rsServ.MoveNext
Loop
clsServ.dbCmact.CommitTrans
bTrans = False
'Exit Sub
'ErrGraba:
'    If btrans Then clsServ.dbCmact.RollbackTrans
'    Set clsServ = Nothing
'    Err.Raise Err.Number, "ActualizaServConvCuentas", Err.Description
End Sub

Public Function GetServPlanPagos(Optional sPersona As String = "", Optional nConvenio As CaptacConvenios) As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetServPlanPagos = clsServ.GetServPlanPagos(sPersona, nConvenio)
Set clsServ = Nothing
End Function

Public Function GetServConvCuentas(Optional sPersona As String = "", Optional nConvenio As CaptacConvenios, Optional cSoloActivas As String = "N") As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetServConvCuentas = clsServ.GetServConvCuentas(sPersona, nConvenio, cSoloActivas)
Set clsServ = Nothing
End Function

Public Sub ActualizaServConvPlanPagos(ByVal sPersona As String, _
        ByVal rsServ As ADODB.Recordset)

Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Dim bTrans As Boolean
Dim bMora As Boolean, bFeriado As Boolean
Dim i As Integer
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
On Error GoTo ErrGraba
clsServ.dbCmact.BeginTrans
bTrans = True
clsServ.EliminaPlanPagos sPersona
i = 1
Do While Not rsServ.EOF
    bMora = IIf(rsServ("Afecto Mora") = "SI", True, False)
    bFeriado = IIf(rsServ("Afecto Feriado") = "SI", True, False)
    clsServ.AgregaPlanPagos sPersona, i, CDate(rsServ("Vencimiento")), CDbl(rsServ("Monto")), CDbl(rsServ("Gasto")), bMora, bFeriado
    rsServ.MoveNext
    i = i + 1
Loop
clsServ.dbCmact.CommitTrans
bTrans = False
Exit Sub
ErrGraba:
    If bTrans Then clsServ.dbCmact.RollbackTrans
    Set clsServ = Nothing
    Err.Raise Err.Number, "", Err.Description
End Sub

Public Function GetUNTConceptosAlumno(ByVal sCurriculo As String, ByVal sTipoMatricula As String, _
        ByVal sIngresante As String, ByVal s2daProf As String) As ADODB.Recordset

Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetUNTConceptosAlumno = clsServ.GetUNTConceptosAlumno(sCurriculo, sTipoMatricula, sIngresante, s2daProf)

Set clsServ = Nothing
End Function

Public Function GetUNTAlumno(ByVal sMatricula As String) As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetUNTAlumno = clsServ.GetUNTAlumno(sMatricula)
Set clsServ = Nothing
End Function

Public Function GetUNTCurriculoConcepto(ByVal sCurriculo As String, ByVal sTipoMatricula As String, _
        ByVal sIngresante As String, ByVal s2daProf As String) As ADODB.Recordset

Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetUNTCurriculoConcepto = clsServ.GetUNTCurriculoConcepto(sCurriculo, sTipoMatricula, sIngresante, s2daProf)
Set clsServ = Nothing
End Function

Public Function GetUNTConceptosOtros() As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetUNTConceptosOtros = clsServ.GetUNTConceptosOtros()
Set clsServ = Nothing
End Function

Public Function GetConvenioCuenta(ByVal nConvenio As CaptacConvenios) As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetConvenioCuenta = clsServ.GetConvenioCuenta(nConvenio)
Set clsServ = Nothing
End Function

Public Function GetConvenioCuenta2(ByVal sPersCod As String, ByVal nConvenio As CaptacConvenios) As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetConvenioCuenta2 = clsServ.GetConvenioCuenta2(sPersCod, nConvenio)
Set clsServ = Nothing
End Function


Public Function GetUNTAlumnoNombre(ByVal sAlumno As String) As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetUNTAlumnoNombre = clsServ.GetUNTAlumnoNombre(sAlumno)
Set clsServ = Nothing
End Function

Public Sub ActualizaOrdenPagoUNT(ByVal sOrden As String, ByVal rsOrden As ADODB.Recordset, _
        ByVal nTotal As Double, ByVal sAlumno As String, ByVal sMovNro As String, _
        ByVal sCuentaPension As String, Optional sCuentaDonacion As String = "", _
        Optional sTipo As String = "M", Optional nCuoFac As Double = 0, _
        Optional sNomAge As String = "", Optional sLpt As String = "")

Dim dFecSis As Date
Dim sCodUser As String
Dim sAgencia As String
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim nConcepto As CapServUNTConcepto
Dim nMonto As Double
Dim nMovNro As Long
'Genera las variables globales del MovNro
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
sAgencia = Mid(sMovNro, 18, 2)

'Abona el monto de la operacion menos el monto de la cuota de facultad
Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
clsCap.CapAbonoCuentaAho sCuentaPension, nTotal - nCuoFac, gAhoDepEfec, sMovNro, "UNT-" & sTipo & "-" & sOrden, , , , , , , , , sNomAge, sLpt, , , False
If nCuoFac > 0 Then
    clsCap.CapAbonoCuentaAho sCuentaDonacion, nCuoFac, gAhoDepEfec, sMovNro, "UNT-" & sTipo & "-" & sOrden, , , , , , , , , sNomAge, sLpt, , , False
End If

Set clsCap = Nothing
'Graba la orden con su detalle
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
nMovNro = clsServ.GetnMovNro(sMovNro)
clsServ.dbCmact.BeginTrans
clsServ.AgregaUNTOrdenPago nMovNro, sOrden, sTipo, dFecSis, sAlumno, nTotal, sAgencia
rsOrden.MoveFirst
Do While Not rsOrden.EOF
    nMonto = rsOrden("Monto")
    nConcepto = rsOrden("nConceptoCod")
    If nMonto > 0 Then
        clsServ.AgregaUNTOrdenPagoDet nMovNro, nConcepto, nMonto
    End If
    rsOrden.MoveNext
Loop
clsServ.dbCmact.CommitTrans
Set clsServ = Nothing
End Sub

Public Function GetConvenioAlumno(ByVal nConvenio As CaptacConvenios, ByVal sCodigo As String) As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetConvenioAlumno = clsServ.GetConvenioAlumno(nConvenio, sCodigo)
Set clsServ = Nothing
End Function

Public Function GetGiroTarifario() As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetGiroTarifario = clsServ.GetGiroTarifario()
Set clsServ = Nothing
End Function

Public Function ServGiroAperturaEfectivo(ByVal nMonto As Double, ByVal nComision As Double, _
    ByVal nTipoGiro As ProductoCuentaTipo, ByVal sAgenciaDest As String, ByVal sMovNro As String, _
    ByVal rsDest As ADODB.Recordset, ByVal sCodRemitente As String, ByVal sRemitente As String, _
    ByVal nmoneda As Moneda, ByVal sNomAge As String, ByVal sLpt As String, ByVal sCMAC As String, _
    Optional sPersLavDinero As String = "", _
    Optional pbITFAplica As Boolean = True, Optional pnITFValor As Currency = 0, Optional pbITFAsumido As Boolean = False, Optional ByVal psITFOpeCod As String = "", Optional psAgeDestino As String = "") As String

Dim clsMov As COMDCaptaServicios.DCOMCaptaServicios
Dim citf_ As New COMNCaptaGenerales.NCOMCaptaMovimiento
Dim clsDmov As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim sMsgOpe As String, sCodUser As String, sCuenta As String, sAgencia As String
Dim sDestinatario As String, sReferencia As String, sGlosa As String, sDocId As String
Dim sDestImp(4) As String, sDestImp2(4) As String, sDestImp3(4) As String
Dim nMovNro As Long, pITFImp As Double
Dim dFecSis As Date
Dim bTrans As Boolean

'On Error GoTo ErrGraba
sDocId = ""
sDestImp(0) = "": sDestImp(2) = "": sDestImp(3) = "": sDestImp(4) = ""
sDestImp2(0) = "": sDestImp2(1) = "": sDestImp2(2) = "": sDestImp2(3) = ""
sDestImp3(0) = "": sDestImp3(1) = "": sDestImp3(2) = "": sDestImp3(3) = ""
'sDestImp2 = "": sDestImp3 = ""

bTrans = False
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sAgencia = Mid(sMovNro, 18, 2)
sCodUser = Right(sMovNro, 4)
Set clsGen = New COMDConstSistema.DCOMGeneral
sCuenta = sCMAC & clsGen.GeneraNuevaCuenta(sAgencia, gGiro, nmoneda)

Set clsGen = Nothing
Set clsMov = New COMDCaptaServicios.DCOMCaptaServicios
clsMov.dbCmact.BeginTrans
bTrans = True



'*******05-05-2004
If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido = False Then
        If psITFOpeCod = gITFCobroCargo Then
            clsMov.AgregaGiro sCuenta, dFecSis, nMonto - pnITFValor, nTipoGiro, sAgenciaDest
           
        Else
            clsMov.AgregaGiro sCuenta, dFecSis, nMonto, nTipoGiro, sAgenciaDest
        End If
    Else
        clsMov.AgregaGiro sCuenta, dFecSis, nMonto, nTipoGiro, sAgenciaDest
    End If
Else
        clsMov.AgregaGiro sCuenta, dFecSis, nMonto, nTipoGiro, sAgenciaDest
End If




clsMov.AgregaProdPersGiro sCuenta, sCodRemitente, gGiroRelPersRemitente
Dim nCont As Integer, nItem As Integer
nCont = 1
nItem = 0

Do While Not rsDest.EOF
    

    If Trim(rsDest("Doc ID")) = "" Then
        sDestinatario = rsDest("Nombre")
        If nCont <= 4 Then
                sDestImp(nItem) = sDestinatario
                nItem = nItem + 1
        ElseIf nCont > 4 And nCont <= 8 Then
                sDestImp2(nItem) = sDestinatario
                nItem = nItem + 1
        ElseIf nCont > 8 And nCont <= 12 Then
                sDestImp3(nItem) = sDestinatario
                nItem = nItem + 1
           
        End If
        sReferencia = rsDest("Referencia")
        clsMov.AgregaGiroDestinatario sCuenta, sDestinatario, sReferencia
    Else
        sDestinatario = rsDest("cPersCod")
        If nCont <= 4 Then
            sDestImp(nItem) = rsDest("Nombre")
            nItem = nItem + 1
            
        ElseIf nCont > 4 And nCont <= 8 Then
            sDestImp2(nItem) = rsDest("Nombre")
            nItem = nItem + 1
            
        ElseIf nCont > 8 And nCont <= 12 Then
            sDestImp3(nItem) = rsDest("Nombre")
            nItem = nItem + 1
            
        End If
        sDocId = rsDest("Doc ID")
        clsMov.AgregaProdPersGiro sCuenta, sDestinatario, gGiroRelPersDestinatario
    End If
    
    nCont = nCont + 1
    If nCont = 5 Or nCont = 9 Then nItem = 0
    
    rsDest.MoveNext
Loop
sGlosa = "GIRO - Apertura Efectivo Cuenta : " & sCuenta

clsMov.AgregaMov sMovNro, gServGiroApertEfec, sGlosa, gMovEstContabMovContable, gMovFlagVigente

nMovNro = clsMov.GetnMovNro(sMovNro)

pITFImp = 0
'*******05-05-2004
If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido = False Then
        If psITFOpeCod = gITFCobroCargo Then
            clsMov.AgregaMovServicios nMovNro, sCuenta, "", nMonto + nComision, , nmoneda, gServGiroApertEfec
            pITFImp = pnITFValor
        Else
            clsMov.AgregaMovServicios nMovNro, sCuenta, "", nMonto + nComision, , nmoneda, gServGiroApertEfec
        End If
    Else
        clsMov.AgregaMovServicios nMovNro, sCuenta, "", nMonto + nComision, , nmoneda, gServGiroApertEfec
    End If
Else
    clsMov.AgregaMovServicios nMovNro, sCuenta, "", nMonto + nComision, , nmoneda, gServGiroApertEfec
End If

clsMov.AgregaMovServiciosDet nMovNro, sCuenta, gConcCapital, nMonto, , gServGiroApertEfec
clsMov.AgregaMovServiciosDet nMovNro, sCuenta, gConcInteres, nComision, , gServGiroApertEfec

clsMov.dbCmact.CommitTrans
bTrans = False

clsMov.dbCmact.BeginTrans
bTrans = True

If pbITFAplica And pnITFValor > 0 Then
    'If pnITFTipoExoneracion = 0 Then
       If pbITFAsumido Then
          clsMov.AgregaMovServicios nMovNro, sCuenta, "", pnITFValor, , nmoneda, psITFOpeCod
          clsMov.AgregaMovServiciosDet nMovNro, sCuenta, gConcITFAsumido, pnITFValor, , psITFOpeCod
       Else
          'If psITFOpeCod = gITFCobroEfectivo Then
             clsMov.AgregaMovServicios nMovNro, sCuenta, "", pnITFValor, , nmoneda, psITFOpeCod
             clsMov.AgregaMovServiciosDet nMovNro, sCuenta, gConcITFCliente, pnITFValor, , psITFOpeCod
          'Else
          '   clsMov.AgregaMovCap nMovNro, psITFOpeCod, sCuenta, pnITFValor, nSaldo - pnITFValor, nSaldo - pnITFValor
          '   clsMov.AgregaMovCapDet nMovNro, psITFOpeCod, sCuenta, gConcITFCliente, pnITFValor
          'End If
       End If
    'End If
End If


If sPersLavDinero <> "" Then
    clsDmov.AgregaMovLavDinero nMovNro, sPersLavDinero, IIf(sDocId = "", sPersLavDinero, sDestinatario)
End If

clsMov.dbCmact.CommitTrans
bTrans = False
Do
    ImprimeBoletaGiros sCuenta, sRemitente, sDestImp(), nMonto, nComision, dFecSis, Format$(Time, "hh:mm:ss"), sNomAge, sCodUser, sLpt, , pITFImp, psAgeDestino
    
    If sDestImp2(0) <> "" Then
       ImprimeBoletaGiros sCuenta, sRemitente, sDestImp2(), nMonto, nComision, dFecSis, Format$(Time, "hh:mm:ss"), sNomAge, sCodUser, sLpt, , pITFImp, psAgeDestino
    End If
    
    If sDestImp3(0) <> "" Then
       ImprimeBoletaGiros sCuenta, sRemitente, sDestImp3(), nMonto, nComision, dFecSis, Format$(Time, "hh:mm:ss"), sNomAge, sCodUser, sLpt, , pITFImp, psAgeDestino
    End If
    
    If pbITFAplica And pnITFValor > 0 Then
        If pbITFAsumido Then
            citf_.fgITFImprimeBoleta sRemitente, pnITFValor, sGlosa, dFecSis, sLpt, 0, , 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, , , , , psAgeDestino
        Else
            citf_.fgITFImprimeBoleta sRemitente, pnITFValor, sGlosa, dFecSis, sLpt, 0, , 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser, , , , , psAgeDestino
        End If
    End If
      
Loop Until MsgBox("Desea Reimprimir Boletas de Giro???", vbQuestion + vbYesNo, "Aviso") = vbNo
ServGiroAperturaEfectivo = sCuenta
Exit Function
'ErrGraba:
 '   If bTrans Then
  '      clsMov.dbCmact.RollbackTrans
  '      Set clsMov = Nothing
 '   End If
  '  MsgBox Err.Description, vbExclamation, "Error"
  '  ServGiroAperturaEfectivo = ""
End Function

Public Sub ServGiroCancelacionEfectivo(ByVal nMonto As Double, ByVal sMovNro As String, _
    ByVal rsDest As ADODB.Recordset, ByVal sCuenta As String, ByVal sNomAge As String, _
    ByVal sRemitente As String, ByVal sLpt As String, Optional nmoneda As Moneda = gMonedaNacional, _
    Optional sPersLavDinero As String = "")

Dim clsMov As COMDCaptaServicios.DCOMCaptaServicios
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim sMsgOpe As String, sCodUser As String, sAgencia As String
Dim sDestinatario(4) As String, sReferencia As String, sGlosa As String, i As Integer, nCont As Integer
Dim sDestinatario2(4) As String, sDestinatario3(4) As String
Dim nMovNro As Long
Dim dFecSis As Date
Dim bTrans As Boolean

sDestinatario(0) = "": sDestinatario(1) = "": sDestinatario(2) = "": sDestinatario(3) = ""
sDestinatario2(0) = "": sDestinatario2(1) = "": sDestinatario2(2) = "": sDestinatario2(3) = ""
sDestinatario3(0) = "": sDestinatario3(1) = "": sDestinatario3(2) = "": sDestinatario3(3) = ""

On Error GoTo ErrGraba
bTrans = False
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sAgencia = Mid(sMovNro, 18, 2)
sCodUser = Right(sMovNro, 4)

Set clsMov = New COMDCaptaServicios.DCOMCaptaServicios
clsMov.dbCmact.BeginTrans
bTrans = True
clsMov.ActualizaEstadoProducto sCuenta, gCapEstCancelada, dFecSis
nCont = 1
i = 0
Do While Not rsDest.EOF
    If rsDest("bRegistrado") = "0" Then
        clsMov.AgregaProdPersGiro sCuenta, rsDest("cPersCod"), gGiroRelPersDestinatario
        
    End If
    
      If nCont <= 4 Then
          sDestinatario(i) = rsDest("Nombre")
          i = i + 1
      ElseIf nCont > 4 And nCont <= 8 Then
          sDestinatario2(i) = rsDest("Nombre")
          i = i + 1
         
      ElseIf nCont >= 9 And nCont < 13 Then
          sDestinatario3(i) = rsDest("Nombre")
          i = i + 1
      End If
    
    nCont = nCont + 1
    If nCont = 5 Or nCont = 9 Then i = 0
    rsDest.MoveNext
Loop
sGlosa = "GIRO - Cancelacion Efectivo Cuenta : " & sCuenta
clsMov.AgregaMov sMovNro, gServGiroCancEfec, sGlosa, gMovEstContabMovContable, gMovFlagVigente
nMovNro = clsMov.GetnMovNro(sMovNro)
clsMov.AgregaMovServicios nMovNro, sCuenta, "", nMonto, , nmoneda
clsMov.AgregaMovServiciosDet nMovNro, sCuenta, gConcCapital, nMonto
If sPersLavDinero <> "" Then
    clsMov.AgregaMovLavDinero nMovNro, sPersLavDinero
End If
clsMov.dbCmact.CommitTrans
bTrans = False
Do
    
    ImprimeBoletaCancelacionGiros sCuenta, sRemitente, sDestinatario(), nMonto, dFecSis, Format$(Time, "hh:mm:ss"), sNomAge, sCodUser, sLpt
    
    If sDestinatario2(0) <> "" Then
    
        ImprimeBoletaCancelacionGiros sCuenta, sRemitente, sDestinatario2(), nMonto, dFecSis, Format$(Time, "hh:mm:ss"), sNomAge, sCodUser, sLpt
    End If
    
    If sDestinatario3(0) <> "" Then
    
        ImprimeBoletaCancelacionGiros sCuenta, sRemitente, sDestinatario3(), nMonto, dFecSis, Format$(Time, "hh:mm:ss"), sNomAge, sCodUser, sLpt
    End If
    
    
Loop Until MsgBox("Desea Reimprimir Boletas de Giro???", vbQuestion + vbYesNo, "Aviso") = vbNo

Exit Sub
ErrGraba:
    If bTrans Then
        clsMov.dbCmact.RollbackTrans
        Set clsMov = Nothing
    End If
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub OtrasOpeDepCtaBanco(ByVal sMovNro As String, ByVal sPersCod As String, _
        ByVal sDocumento As String, ByVal nMonto As Double, ByVal sGlosa As String, _
        ByVal sNomAge As String, ByVal nmoneda As Moneda)

Dim clsMov As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim nMovNro As Long
Dim bTrans As Boolean
Dim dFecSis As Date
bTrans = False
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
Set clsMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
clsMov.dbCmact.BeginTrans
bTrans = True
On Error GoTo ErrGraba
clsMov.AgregaMov sMovNro, gOtrOpeDepCtaBcoEfec, sGlosa, gMovEstContabMovContable, gMovFlagVigente
nMovNro = clsMov.GetnMovNro(sMovNro)
clsMov.AgregaMovDoc nMovNro, TpoDocBolDep, sDocumento, dFecSis
clsMov.AgregaMovOpeVarias nMovNro, nMonto, sDocumento, sPersCod, nmoneda
clsMov.dbCmact.CommitTrans
bTrans = False
Set clsMov = Nothing
Exit Sub
ErrGraba:
    If bTrans Then
        clsMov.dbCmact.RollbackTrans
        Set clsMov = Nothing
    End If
    MsgBox Err.Description, vbExclamation, "Error"

End Sub

Public Function GetGiroDatos(ByVal sCuenta As String) As ADODB.Recordset
Dim clsGiro As COMDCaptaServicios.DCOMCaptaServicios
Set clsGiro = New COMDCaptaServicios.DCOMCaptaServicios
Set GetGiroDatos = clsGiro.GetGiroDatos(sCuenta)
Set clsGiro = Nothing
End Function


Public Function GetGiroPendiente(ByVal sCodage As String) As ADODB.Recordset
Dim clsGiro As COMDCaptaServicios.DCOMCaptaServicios
Set clsGiro = New COMDCaptaServicios.DCOMCaptaServicios
Set GetGiroPendiente = clsGiro.GetGiroPendiente(sCodage)
Set clsGiro = Nothing
End Function

Public Function GetPersonasExoLavDinero() As ADODB.Recordset
Dim clsLav As COMDCaptaServicios.DCOMCaptaServicios
Set clsLav = New COMDCaptaServicios.DCOMCaptaServicios
Set GetPersonasExoLavDinero = clsLav.GetPersonasExoLavDinero()
Set clsLav = Nothing
End Function

Public Function GetPersonaHistExoLavDinero(ByVal sPersCod As String) As ADODB.Recordset
Dim clsLav As COMDCaptaServicios.DCOMCaptaServicios
Set clsLav = New COMDCaptaServicios.DCOMCaptaServicios
Set GetPersonaHistExoLavDinero = clsLav.GetPersonaHistExoLavDinero(sPersCod)
Set clsLav = Nothing
End Function

Public Function GetInfoRP(ByVal sAgencia As String) As ADODB.Recordset
Dim clsRP As COMDCaptaServicios.DCOMCaptaServicios
Set clsRP = New COMDCaptaServicios.DCOMCaptaServicios
Set GetInfoRP = clsRP.GetInfoRP(sAgencia)
Set clsRP = Nothing

End Function

Public Function GetInfoActa(ByVal sUser As String) As ADODB.Recordset
Dim clsRP As COMDCaptaServicios.DCOMCaptaServicios
Set clsRP = New COMDCaptaServicios.DCOMCaptaServicios
Set GetInfoActa = clsRP.GetInfoActa(sUser)
Set clsRP = Nothing

End Function
Public Function GetInfoDetActa(ByVal sNroActa As String) As ADODB.Recordset
Dim clsRP As COMDCaptaServicios.DCOMCaptaServicios
Set clsRP = New COMDCaptaServicios.DCOMCaptaServicios
Set GetInfoDetActa = clsRP.GetInfoDetActa(sNroActa)
Set clsRP = Nothing

End Function




Public Function EsCuentaExoneradaLavadoDinero(ByVal sCuenta As String) As Boolean
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Dim bExito As Boolean
Dim rsLav As ADODB.Recordset
Dim sPersona As String
'bExito = True
bExito = False
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set rsLav = clsServ.GetTitularExoLavDinero(sCuenta)
Set clsServ = Nothing
Do While Not rsLav.EOF
    If IsNull(rsLav("PersonaExo")) Then
        bExito = False
        Exit Do
    Else
        sPersona = Trim(rsLav("PersonaExo"))
        MsgBox "Persona: " & sPersona & " EXONERADA de Control de Lavado de Dinero", vbInformation, "Aviso"
        bExito = True
        
    End If
    rsLav.MoveNext
Loop
rsLav.Close
Set rsLav = Nothing
EsCuentaExoneradaLavadoDinero = bExito
End Function


Public Function EsPersonaExoneradaLavadoDinero(ByVal sPersona As String) As Boolean
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Dim bExito As Boolean
Dim rsLav As ADODB.Recordset
bExito = True
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set rsLav = clsServ.GetPersonaExoneradaLavDinero(sPersona)
Set clsServ = Nothing
If rsLav.EOF And rsLav.BOF Then
    bExito = False
Else
    If rsLav("nEstado") = gPersLavDinConControl Then
        bExito = False
    End If
End If
rsLav.Close
Set rsLav = Nothing
EsPersonaExoneradaLavadoDinero = bExito
End Function

Public Sub ActualizaPersExoLavDinero(ByVal rsPers As ADODB.Recordset, ByVal rsHist As ADODB.Recordset, ByVal sMovNro As String)
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Dim bTrans As Boolean
Dim sComentario As String, sPersCod As String, sMov As String
Dim nEstado As PersEstLavDinero

'Agregamos primero a las nuevas personas
bTrans = False
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
clsServ.dbCmact.BeginTrans
bTrans = True
On Error GoTo ErrGraba
If Not rsPers Is Nothing Then
    Do While Not rsPers.EOF
        If rsPers("Flag") = "N" Then
            sPersCod = rsPers("Codigo")
            sComentario = Trim(rsPers("Comentario"))
            nEstado = CLng(Right(Trim(rsPers("Estado")), 2))
            clsServ.AgregaPersExoLavDinero sPersCod, sMovNro, nEstado, sComentario
        End If
        rsPers.MoveNext
    Loop
End If
'Agregamos los nuevos estados para las personas ya existentes
If Not rsHist Is Nothing Then
    Do While Not rsHist.EOF
        If rsHist("Flag") = "N" Then
            sPersCod = rsHist("Codigo")
            sMov = rsHist("sMovNro")
            sComentario = Trim(rsHist("Comentario"))
            nEstado = CLng(Right(Trim(rsHist("Estado")), 2))
            clsServ.AgregaPersExoLavDinero sPersCod, sMovNro, nEstado, sComentario
        ElseIf rsHist("Flag") = "M" Then
            sPersCod = rsHist("Codigo")
            sMov = rsHist("sMovNro")
            sComentario = Trim(rsHist("Comentario"))
            clsServ.ActualizaPersExoLavDinero sPersCod, sMovNro, sComentario, sMov
        End If
        rsHist.MoveNext
    Loop
End If
clsServ.dbCmact.CommitTrans
bTrans = False
Exit Sub
ErrGraba:
    If bTrans Then
        clsServ.dbCmact.RollbackTrans
    End If
    Set clsServ = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub ActualizaChequeFechaValorizacion(ByVal sMovNro As String, ByVal sGlosa As String, _
            ByVal rsChq As ADODB.Recordset)
Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
Dim sPersCod As String
Dim sNroDoc As String
Dim sIFTpo As String
Dim dFechaValor As Date

Set oserv = New COMDCaptaServicios.DCOMCaptaServicios
oserv.dbCmact.BeginTrans
Do While Not rsChq.EOF
    If rsChq("Flag") = "M" Then
        sPersCod = Right(Trim(rsChq("Banco")), 13)
        sNroDoc = Trim(rsChq("Cheque"))
        sIFTpo = Left(Right(Trim(rsChq("Banco")), 15), 2)
        dFechaValor = CDate(rsChq("Fecha Valor"))
        oserv.AgregaMov sMovNro, gChqOpeModFechaValor, sGlosa, gMovEstContabNoContable, gMovFlagVigente
        oserv.ActualizaChequeFechaValorizacion sNroDoc, sPersCod, sIFTpo, dFechaValor
        rsChq.MoveNext
    End If
Loop
oserv.dbCmact.CommitTrans
Set oserv = Nothing
End Sub

Public Function GetInteres(ByVal nCapital As Double, ByVal nTasa As Double, _
            ByVal nPlazo As Long, Optional nTipoInteres As TipoCalculoInteres = TpoCalcIntSimple) As Double
If nTipoInteres = TpoCalcIntSimple Then
    GetInteres = Round((nTasa / 36000) * nPlazo * nCapital, 2)
ElseIf nTipoInteres = TpoCalcIntCompuesto Then
    GetInteres = Round((((nTasa / 36000) + 1) ^ nPlazo - 1) * nCapital, 2)
End If
End Function

Public Sub ValorizaChequeInmediato(ByVal sMovNro As String, ByVal sGlosa As String, _
            ByVal sNroDoc As String, ByVal sPersCod As String, ByVal sIFTpo As String, Optional ByVal sIfCta As String)

Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
Dim oCta As COMDCaptaGenerales.DCOMCaptaGenerales
Dim rsCuenta As ADODB.Recordset
Dim sCuenta As String
Dim nMonto As Double
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nSaldoRetiro As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bTrans As Boolean
Dim nDiasTranscurridos As Long
Dim nMovNro As Long
Dim dFecSis As Date
Dim nProducto As Producto
Dim rsCta As ADODB.Recordset
Dim nPorcDisp As Double

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))

Set oserv = New COMDCaptaServicios.DCOMCaptaServicios
Set rsCuenta = oserv.GetChequeCuenta(sNroDoc, sPersCod, sIFTpo)
On Error GoTo ErrGraba
oserv.dbCmact.BeginTrans
bTrans = True

    Dim rstemp1 As ADODB.Recordset
    Dim oMant As COMDCaptaGenerales.DCOMCaptaMovimiento
    Set oMant = New COMDCaptaGenerales.DCOMCaptaMovimiento


Do While Not rsCuenta.EOF
    sCuenta = rsCuenta("cCtaCod")
    nMonto = rsCuenta("nMonto")
    nProducto = CLng(Mid(sCuenta, 6, 3))
    
    If nProducto = gCapAhorros Then
        'Obtiene los datos para el calculo
        Set oCta = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsCta = oCta.GetDatosCuentaAho(sCuenta)
        Set oCta = Nothing
        nSaldoDisp = rsCta("nSaldoDisp")
        
        nSaldoCnt = rsCta("nSaldo")
        dUltMov = rsCta("dUltCierre")
        nTasa = rsCta("nTasaInteres")
        nIntAcum = rsCta("nIntAcum")
        rsCta.Close
        Set rsCta = Nothing

        'Calcula intereses
        nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
        nIntGanado = GetInteres(nSaldoDisp + nMonto, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
        dUltMov = DateAdd("d", -1, dFecSis)
        oserv.AgregaMov sMovNro, gChqOpeValorInmediata, sGlosa, gMovEstContabNoContable, gMovFlagVigente
        oserv.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
        
        oserv.ActualizaAbonoCaptacion sCuenta, 0, nMonto, nIntGanado, dUltMov, sMovNro, False
        nMovNro = oserv.GetnMovNro(sMovNro)
        oserv.AgregaMovCap nMovNro, gChqOpeValorInmediata, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
    
    ElseIf nProducto = gCapPlazoFijo Then
        Set oCta = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsCta = oCta.GetDatosCuentaPF(sCuenta)
        Set oCta = Nothing
        nSaldoDisp = rsCta("nSaldoDisp")
        nSaldoCnt = rsCta("nSaldo")
        rsCta.Close
        Set rsCta = Nothing
        
        oserv.ActualizaFechaRenovacionPF sCuenta, dFecSis
        oserv.ActualizaSaldoDisponiblePF sCuenta, dFecSis, nMonto
        oserv.AgregaMov sMovNro, gChqOpeValorInmediata, sGlosa
        nMovNro = oserv.GetnMovNro(sMovNro)
        oserv.AgregaMovCap nMovNro, gChqOpeValorInmediata, sCuenta, nMonto, nSaldoDisp + nMonto, nSaldoCnt
        
    ElseIf nProducto = gCapCTS Then
        Set oCta = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsCta = oCta.GetDatosCuentaCTS(sCuenta)
        Set oCta = Nothing
        nSaldoDisp = rsCta("nSaldoDisp")
        nSaldoCnt = rsCta("nSaldo")
        dUltMov = rsCta("dUltCierre")
        nTasa = rsCta("nTasaInteres")
        nIntAcum = rsCta("nIntAcum")
        nSaldoRetiro = rsCta("nSaldRetiro")
        rsCta.Close
        Set rsCta = Nothing
        
        nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
        nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntCompuesto)
        nIntSaldo = GetInteres(nSaldoRetiro, nTasa, nDiasTranscurridos, TpoCalcIntCompuesto)
        nSaldoCnt = nSaldoCnt + nMonto
        dUltMov = DateAdd("d", -1, dFecSis)

        oserv.AgregaMov sMovNro, gChqOpeValorInmediata, sGlosa, gMovEstContabNoContable, gMovFlagVigente
        nMovNro = oserv.GetnMovNro(sMovNro)

        nPorcDisp = oserv.GetPorcentajeRetiroCTS(sCuenta, sNroDoc, sPersCod, sIFTpo)
        nSaldoDisp = nSaldoDisp + nMonto
        nSaldoRetiro = nSaldoRetiro + (nMonto * nPorcDisp / 100)

        oserv.AgregaMovCap nMovNro, gChqOpeValorInmediata, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
        oserv.AgregaMovCapDet nMovNro, gChqOpeValorInmediata, sCuenta, gConcProvision, nIntGanado
        oserv.ActualizaSaldoRetiroCTS sCuenta, nSaldoRetiro, nIntSaldo
        oserv.ActualizaAbonoCaptacion sCuenta, 0, nMonto, nIntGanado, dUltMov, sMovNro, False
    
    End If
    
    oserv.ActualizaEstadoCheque sNroDoc, sPersCod, sIFTpo, gChqEstValorizado, dFecSis
    
'    Set rstemp1 = oMant.GetDatosCheque(sNroDoc, sPersCod, sIFTpo)
    
    oserv.AgregaDocRecEstado sNroDoc, sPersCod, sIFTpo, gChqEstValorizado, sMovNro, sIfCta
    
'    Set rstemp1 = Nothing
    
    rsCuenta.MoveNext
Loop
oserv.dbCmact.CommitTrans
bTrans = False
Set oserv = Nothing


Exit Sub
ErrGraba:
    If bTrans Then oserv.dbCmact.RollbackTrans
    Set oserv = Nothing
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Function ValidaSaldos(ByVal rsCuenta As ADODB.Recordset) As Boolean
Dim sCuenta As String
Dim nMonto As Double
Dim nSaldoDisp As Double, nSaldoCnt As Double
Dim nProducto As Producto
Dim oCta As COMDCaptaGenerales.DCOMCaptaGenerales
Dim rsCta As ADODB.Recordset

Do While Not rsCuenta.EOF
    sCuenta = rsCuenta("cCtaCod")
    nMonto = rsCuenta("nMonto")
    nProducto = CLng(Mid(sCuenta, 6, 3))
    If nProducto = gCapAhorros Then
        Dim nMontoMinimo As Double
        Dim nmoneda As Moneda
        Dim clsPar As COMNCaptaGenerales.NCOMCaptaDefinicion
        Dim nPersoneria As PersPersoneria
        Dim bOrdPag As Boolean
        
        nmoneda = CLng(Mid(sCuenta, 9, 1))
        'Obtiene los datos para el calculo
        Set oCta = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsCta = oCta.GetDatosCuentaAho(sCuenta)
        Set oCta = Nothing
        nSaldoDisp = rsCta("nSaldoDisp")
        nSaldoCnt = rsCta("nSaldo")
        nPersoneria = rsCta("nPersoneria")
        bOrdPag = rsCta("bOrdPag")
        rsCta.Close
        Set rsCta = Nothing
        
        Set clsPar = New COMNCaptaGenerales.NCOMCaptaDefinicion
        nMontoMinimo = clsPar.GetSaldoMinimoPersoneria(nProducto, nmoneda, nPersoneria, bOrdPag)
        Set clsPar = Nothing
        
        If nSaldoDisp - nMontoMinimo < nMonto Then
            ValidaSaldos = False
            Exit Function
        End If
    ElseIf nProducto = gCapCTS Then
        'Obtiene los datos para el calculo
        Set oCta = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsCta = oCta.GetDatosCuentaCTS(sCuenta)
        Set oCta = Nothing
        nSaldoDisp = rsCta("nSaldoDisp")
        nSaldoCnt = rsCta("nSaldo")
        rsCta.Close
        Set rsCta = Nothing
        If nSaldoDisp < nMonto Then
            ValidaSaldos = False
            Exit Function
        End If
    
    End If
    rsCuenta.MoveNext
Loop
ValidaSaldos = True
End Function

Public Function ExtornaValorizaChequeInmediato(ByVal sMovNro As String, ByVal sGlosa As String, _
            ByVal sNroDoc As String, ByVal sPersCod As String, ByVal sIFTpo As String, _
            ByVal nMovNroExt As Long) As Integer

Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
Dim oCta As COMDCaptaGenerales.DCOMCaptaGenerales
Dim rsCuenta As ADODB.Recordset
Dim sCuenta As String
Dim nMonto As Double
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nSaldoRetiro As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bTrans As Boolean
Dim nDiasTranscurridos As Long
Dim nMovNro As Long
Dim dFecSis As Date
Dim nProducto As Producto
Dim rsCta As ADODB.Recordset
Dim nPorcDisp As Double

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))

Set oserv = New COMDCaptaServicios.DCOMCaptaServicios
Set rsCuenta = oserv.GetChequeCuenta(sNroDoc, sPersCod, sIFTpo)

If Not ValidaSaldos(rsCuenta) Then
    MsgBox "Cuenta NO Posee Saldo Suficiente para realizar el extorno.", vbExclamation, "Aviso"
    Set rsCuenta = Nothing
    Set oserv = Nothing
    ExtornaValorizaChequeInmediato = -1
    Exit Function
End If
On Error GoTo ErrGraba
oserv.dbCmact.BeginTrans
bTrans = True
rsCuenta.MoveFirst
Do While Not rsCuenta.EOF
    sCuenta = rsCuenta("cCtaCod")
    nMonto = rsCuenta("nMonto")
    nProducto = CLng(Mid(sCuenta, 6, 3))
    
    If nProducto = gCapAhorros Then
        'Obtiene los datos para el calculo
        Set oCta = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsCta = oCta.GetDatosCuentaAho(sCuenta)
        Set oCta = Nothing
        nSaldoDisp = rsCta("nSaldoDisp")
        nSaldoCnt = rsCta("nSaldo")
        dUltMov = rsCta("dUltCierre")
        rsCta.Close
        Set rsCta = Nothing
        
        oserv.AgregaMov sMovNro, gChqOpeExtValorInmediata, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
        oserv.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
        nSaldoDisp = nSaldoDisp - nMonto
        oserv.ActualizaCargoCaptacion sCuenta, 0, nMonto, 0, dUltMov, sMovNro, False
        nMovNro = oserv.GetnMovNro(sMovNro)
        oserv.AgregaMovCap nMovNro, gChqOpeExtValorInmediata, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
        
    ElseIf nProducto = gCapPlazoFijo Then
        Set oCta = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsCta = oCta.GetDatosCuentaPF(sCuenta)
        Set oCta = Nothing
        nSaldoDisp = rsCta("nSaldoDisp")
        nSaldoCnt = rsCta("nSaldo")
        rsCta.Close
        Set rsCta = Nothing
        
        oserv.ActualizaSaldoDisponiblePF sCuenta, dFecSis, nSaldoDisp - nMonto
        oserv.AgregaMov sMovNro, gChqOpeExtValorInmediata, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
        nMovNro = oserv.GetnMovNro(sMovNro)
        oserv.AgregaMovCap nMovNro, gChqOpeExtValorInmediata, sCuenta, nMonto, nSaldoDisp - nMonto, nSaldoCnt
        
    ElseIf nProducto = gCapCTS Then
        Dim nMontoProvision As Double
        Set oCta = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsCta = oCta.GetDatosCuentaCTS(sCuenta)
        Set oCta = Nothing
        nSaldoDisp = rsCta("nSaldoDisp")
        nSaldoCnt = rsCta("nSaldo")
        dUltMov = rsCta("dUltCierre")
        nSaldoRetiro = rsCta("nSaldRetiro")
        rsCta.Close
        Set rsCta = Nothing
        nPorcDisp = oserv.GetPorcentajeRetiroCTS(sCuenta, sNroDoc, sPersCod, sIFTpo)
        nMontoProvision = oserv.GetProvisionCTS(nMovNroExt)
        oserv.AgregaMov sMovNro, gChqOpeExtValorInmediata, sGlosa
        nMovNro = oserv.GetnMovNro(sMovNro)
        
        nSaldoDisp = nSaldoDisp - nMonto
        nSaldoRetiro = nSaldoRetiro - (nMonto * nPorcDisp / 100)

        oserv.AgregaMovCap nMovNro, gChqOpeExtValorInmediata, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
        oserv.AgregaMovCapDet nMovNro, gChqOpeExtValorInmediata, sCuenta, gConcProvision, nMontoProvision
        oserv.ActualizaSaldoRetiroCTS sCuenta, nSaldoRetiro, 0
        oserv.ActualizaCargoCaptacion sCuenta, 0, nMonto, 0, dUltMov, sMovNro, False
    
    End If
    oserv.ActualizaEstadoMov nMovNroExt, gMovEstContabNoContable, gMovFlagExtornado
    oserv.AgregaMovRef nMovNro, nMovNroExt
    oserv.ActualizaEstadoCheque sNroDoc, sPersCod, sIFTpo, gChqEstEnValorizacion
    oserv.AgregaDocRecEstado sNroDoc, sPersCod, sIFTpo, gChqEstEnValorizacion, sMovNro
        
    rsCuenta.MoveNext
Loop
oserv.dbCmact.CommitTrans
bTrans = False
Set oserv = Nothing
ExtornaValorizaChequeInmediato = 0
Exit Function
ErrGraba:
    If bTrans Then oserv.dbCmact.RollbackTrans
    Set oserv = Nothing
    MsgBox Err.Description, vbCritical, "Error"
End Function

Public Sub ExtornaChequeRegistro(ByVal sMovNro As String, ByVal sGlosa As String, _
            ByVal sNroDoc As String, ByVal sPersCod As String, ByVal sIFTpo As String, _
            ByVal nMovNroExt As Long)
Dim bTrans As Boolean
Dim oCap As COMDCaptaServicios.DCOMCaptaServicios
Dim nMovNro As Long

On Error GoTo ErrGraba
Set oCap = New COMDCaptaServicios.DCOMCaptaServicios
oCap.dbCmact.BeginTrans
bTrans = True
oCap.ActualizaEstadoMov nMovNroExt, gMovEstContabMovContable, gMovFlagExtornado
oCap.AgregaMov sMovNro, gChqOpeExtRegistro, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno

oCap.ActualizaEstadoCheque sNroDoc, sPersCod, sIFTpo, gsChqEstExtornado
oCap.AgregaDocRecEstado sNroDoc, sPersCod, sIFTpo, gsChqEstExtornado, sMovNro
nMovNro = oCap.GetnMovNro(sMovNro)
oCap.AgregaMovRef nMovNro, nMovNroExt
oCap.dbCmact.CommitTrans
bTrans = False
Set oCap = Nothing
Exit Sub
ErrGraba:
    If bTrans Then oCap.dbCmact.RollbackTrans
    Set oCap = Nothing
    MsgBox Err.Description, vbCritical, "Error"
    
End Sub




Public Sub AgregaMovConvCobranza(ByVal nMovNro As Long, ByVal cPersCod As String, ByVal cCodigo As String, ByVal nTpoCuenta As Integer, ByVal nMonto As Double, ByVal cReferencia As String)
   Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
   Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
   
            Call clsServ.AgregaMovConvCobranza(nMovNro, cPersCod, cCodigo, nTpoCuenta, nMonto, cReferencia)

End Sub

Public Sub GrabaAbonoServicios(ByVal rsServ As ADODB.Recordset, ByVal nMontoAbono As Double, _
        ByVal nMontoCargo As Double, ByVal sCuenta As String, ByVal nOpeAbono As CaptacOperacion, _
        ByVal nOpeCargo As CaptacOperacion, ByVal sGlosa As String, ByVal dFecha As Date, _
        ByVal sUsuario As String, ByVal sAgencia As String, ByVal sNomCmac As String, ByVal sNomAge As String, _
        ByVal sLpt As String)

Dim oCap As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
Dim oMov As COMNContabilidad.NCOMContFunciones
Dim sMovNroAbono As String, sMovNroCargo As String
Dim nMovNroAbono As Long, nMovNroCargo As Long, nMovNroRef
'Abona el monto total cobrado
Set oMov = New COMNContabilidad.NCOMContFunciones
sMovNroAbono = oMov.GeneraMovNro(dFecha, sAgencia, sUsuario)
Set oMov = Nothing
Set oCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
oCap.CapAbonoCuentaAho sCuenta, nMontoAbono, nOpeAbono, sMovNroAbono, sGlosa, , , , , , , , sNomCmac, sNomAge, sLpt
Set oCap = Nothing
Set oMov = New COMNContabilidad.NCOMContFunciones
sMovNroCargo = oMov.GeneraMovNro(dFecha, sAgencia, sUsuario)
Set oMov = Nothing
Set oCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
oCap.CapCargoCuentaAho sCuenta, nMontoCargo, nOpeCargo, sMovNroCargo, sGlosa, , , , , , , , sNomCmac, sNomAge, sLpt
Set oCap = Nothing

Set oserv = New COMDCaptaServicios.DCOMCaptaServicios
nMovNroAbono = oserv.GetnMovNro(sMovNroAbono)
nMovNroCargo = oserv.GetnMovNro(sMovNroCargo)
rsServ.MoveFirst
oserv.dbCmact.BeginTrans
Do While Not rsServ.EOF
    nMovNroRef = rsServ("nMovNro")
    oserv.AgregaMovRef nMovNroAbono, nMovNroRef
    rsServ.MoveNext
Loop
oserv.AgregaMovRef nMovNroCargo, nMovNroAbono
oserv.dbCmact.CommitTrans
Set oserv = Nothing
End Sub

Public Sub GrabaMovCobranzas(ByVal nMovNro As Long, ByVal cPersCod As String, ByVal cCodigo As String, ByVal nTpoCuenta As Integer, ByVal nMonto As Double, ByVal cReferencia As String)
Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
Set oserv = New COMDCaptaServicios.DCOMCaptaServicios
    
    Call oserv.GrabaMovCobranzas(nMovNro, cPersCod, cCodigo, nTpoCuenta, nMonto, cReferencia)
    
    Set oserv = Nothing
    
End Sub

Public Function GrabaActa(ByVal CodPortador As String, ByVal FechaIncautacion As Date, ByVal CodUsuDetector As String, ByVal COdUsuVerificador As String, ByVal cMovNro As String, ByVal cMoneda As String, ByVal RSTMP As ADODB.Recordset, ByRef NroActa As String) As Boolean
Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
Set oserv = New COMDCaptaServicios.DCOMCaptaServicios

'GrabaActa(ByVal CodPortador As String, ByVal FechaIncautacion As Date, ByVal CodUsuDetector As String, ByVal COdUsuVerificador As String, ByVal cMovnro As String, ByVal cMoneda As String, ByVal RSTMP As ADODB.Recordset)
    GrabaActa = oserv.GrabaActa(CodPortador, FechaIncautacion, CodUsuDetector, COdUsuVerificador, cMovNro, cMoneda, RSTMP, NroActa)
    
    Set oserv = Nothing
    
End Function

'GrabaActa(ByVal CodPortador As String, ByVal FechaIncautacion As Date, ByVal CodUsuDetector As String, ByVal COdUsuVerificador As String, ByVal cMovnro As String, ByVal cMoneda As String, ByVal RSTMP As ADODB.Recordset)



' NUEVO - CMCPL
Public Function AgregaServSATTributo(ByVal rsTrib As ADODB.Recordset) As Integer
Dim sCodTributo As String, sTributo As String, sAno As String
Dim sRecibo As String, sDigCheq As String, sPeriodo As String
Dim nValDeuda As Double, nValDerEmis As Double, nValAjuste As Double
Dim nValIntMor As Double, dFecVenc As Date, sNombre As String
Dim sCuenta As String, sContrib As String, nValGastos As Double
Dim nValCostas As Double, sCodSis As String, nTotal As Double
Dim bTrans As Boolean
Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
Set oserv = New COMDCaptaServicios.DCOMCaptaServicios

On Error GoTo ErrGraba
rsTrib.MoveFirst

oserv.dbCmact.BeginTrans
oserv.EliminaServSATTributo
bTrans = True
Do While Not rsTrib.EOF
    sCodTributo = Trim(rsTrib("cCodTributo")) 'Nombre de cabecera del Flex
    sTributo = Trim(rsTrib("cTributo"))
    sAno = Trim(rsTrib("cAno"))
    sRecibo = rsTrib("cRecibo")
    sDigCheq = rsTrib("cDigCheq")
    sPeriodo = rsTrib("cPeriodo")
    nValDeuda = rsTrib("nValDeuda")
    nValDerEmis = rsTrib("nValDerEmis")
    nValAjuste = rsTrib("nValAjuste")
    nValIntMor = rsTrib("nValIntMor")
    dFecVenc = rsTrib("dFecVenc")
    sNombre = Trim(rsTrib("cNombre"))
    sCuenta = Trim(rsTrib("cCuenta"))
    sContrib = Trim(rsTrib("cContrib"))
    nValGastos = rsTrib("nValGastos")
    nValCostas = rsTrib("nValCostas")
    sCodSis = Trim(rsTrib("cCodSis"))
    nTotal = rsTrib("nTotal")
    oserv.AgregaServSATTributo sCodTributo, sTributo, sAno, sRecibo, sDigCheq, sPeriodo, nValDeuda, _
            nValDerEmis, nValAjuste, nValIntMor, dFecVenc, sNombre, sCuenta, sContrib, nValGastos, _
            nValCostas, sCodSis, nTotal
    rsTrib.MoveNext
Loop
oserv.dbCmact.CommitTrans
bTrans = False
Set oserv = Nothing
AgregaServSATTributo = 0
Exit Function
ErrGraba:
    If bTrans Then oserv.dbCmact.RollbackTrans
    Set oserv = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
    AgregaServSATTributo = -1
End Function

' NUEVO - CMCPL
Public Function AgregaServSATpapeletas(ByVal rsPap As ADODB.Recordset) As Integer
Dim sCodTributo As String, sTributo As String, sAno As String
Dim sPlaca As String, sDigCheq As String, sPeriodo As String
Dim nImporte As Double, nValReincide As Double, nValGastAdm As Double
Dim nValCostas As Double
Dim dFecVenc As Date, dFecInf As Date
Dim dFecLect As String, dFecNoti As String
Dim sPapeleta As String, sTipo As String, sFalta As String, sCodAux As String, sLicCond As String
Dim bTrans As Boolean
Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
Set oserv = New COMDCaptaServicios.DCOMCaptaServicios

'On Error GoTo ErrGraba
rsPap.MoveFirst

oserv.dbCmact.BeginTrans
oserv.EliminaServSATPapeletas
bTrans = True
Do While Not rsPap.EOF
    sCodTributo = Trim(rsPap("cCodTributo")) 'Nombre de cabecera del Flex
    sTributo = Trim(rsPap("cTributo"))
    sAno = Trim(rsPap("cAno"))
    sPlaca = Trim(rsPap("cPlaca"))
    sDigCheq = Trim(rsPap("cDigCheq"))
    sPeriodo = Trim(rsPap("cPeriodo"))
    
    nImporte = rsPap("nImporte")
    nValReincide = rsPap("nValReincide")
    nValGastAdm = rsPap("nValGastAdm")
    nValCostas = rsPap("nValCostas")
    
    dFecVenc = rsPap("dFecVenc")
    dFecInf = Trim(rsPap("dFecInf"))
    
    dFecLect = Trim(rsPap("dFecLect"))
    dFecNoti = Trim(rsPap("dFecNoti"))
    
    
    sPapeleta = Trim(rsPap("spapeletas"))
    sTipo = Trim(rsPap("sTipo"))
    sFalta = Trim(rsPap("sFalta"))
    sCodAux = Trim(rsPap("sCodAux"))
    sLicCond = Trim(rsPap("sLicCond"))
    
    oserv.AgregaServSATpapeletas sCodTributo, sTributo, sAno, sPlaca, sDigCheq, sPeriodo, nImporte, _
            nValReincide, nValGastAdm, nValCostas, dFecVenc, dFecInf, dFecLect, dFecNoti, sPapeleta, _
            sTipo, sFalta, sCodAux, sLicCond
    rsPap.MoveNext
Loop
MsgBox "Exportación realizada con éxito", vbInformation, "Aviso"
oserv.dbCmact.CommitTrans
bTrans = False
Set oserv = Nothing
AgregaServSATpapeletas = 0
Exit Function
'ErrGraba:
'    If bTrans Then oServ.dbCmact.RollbackTrans
'    Set oServ = Nothing
'    MsgBox Err.Description, vbExclamation, "Error"
'    AgregaServSATpapeletas = -1
End Function

' NUEVO - CMCPL
Public Function GetServSATPapeletas(ByVal sDato As String, ByVal nTipo As Integer) As ADODB.Recordset
Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
Set oserv = New COMDCaptaServicios.DCOMCaptaServicios
If nTipo = 0 Then
    Set GetServSATPapeletas = oserv.GetServSATPapeletaLincencia(sDato)
ElseIf nTipo = 1 Then
    Set GetServSATPapeletas = oserv.GetServSATPapeletaPlaca(sDato)
ElseIf nTipo = 2 Then
    Set GetServSATPapeletas = oserv.GetServSATPapeletaNPapeleta(sDato)
End If
Set oserv = Nothing
End Function

' NUEVO - CMCPL
Public Function GetServSATTributos(ByVal sDato As String, ByVal nTipo As Integer) As ADODB.Recordset
Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
Set oserv = New COMDCaptaServicios.DCOMCaptaServicios
If nTipo = 0 Then 'Derechos
    Set GetServSATTributos = oserv.GetServSATDerecho(sDato)
ElseIf nTipo = 1 Then 'Recibos
    Set GetServSATTributos = oserv.GetServSATRecibo(sDato)
ElseIf nTipo = 2 Then 'Otros Pagos
    'Set GetServSATTributos = oServ.GetServSATOtros(sDato)
End If
Set oserv = Nothing
End Function

' NUEVO - CMCPL
Public Function GetSatMontoFecha(ByVal dInicio As String, ByVal dFin As String) As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetSatMontoFecha = clsServ.GetSatMontoFecha(dInicio, dFin)
Set clsServ = Nothing
End Function

' NUEVO - CMCPL
Public Function GetSatMontoFechaDis(ByVal dInicio As String, ByVal dFin As String) As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetSatMontoFechaDis = clsServ.GetSatMontoFechaDis(dInicio, dFin)
Set clsServ = Nothing
End Function
' NUEVO - CMCPL - CRSF 11/06 Monto de Parametros para  Distribucion de Fondos
Public Function GetSatDistribucion(ByVal sParam As String) As ADODB.Recordset
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
Set GetSatDistribucion = clsServ.GetSatDistribucion(sParam)
Set clsServ = Nothing
End Function
' NUEVO - CMCPL - CRSF 11/06 Suma de Totales de Cobro SAT
Public Function GetSatSuma(ByVal sParam As String) As Double
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios
GetSatSuma = clsServ.GetSatSuma(sParam)
Set clsServ = Nothing
End Function


'NUEVO CMCPL CRSF 11/06
Public Sub GrabaAbonoServiciosSAT(ByVal rsServ As ADODB.Recordset, ByVal nMontoAbono As Double, _
        ByVal nMontoCargo As Double, ByVal sCuenta As String, ByVal nOpeAbono As CaptacOperacion, _
        ByVal nOpeCargo As CaptacOperacion, ByVal sGlosa As String, ByVal dFecha As Date, _
        ByVal sUsuario As String, ByVal sAgencia As String, ByVal sNomCmac As String, ByVal sNomAge As String, _
        ByVal sLpt As String)

Dim oCap As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
Dim oMov As COMNContabilidad.NCOMContFunciones
Dim sMovNroAbono As String, sMovNroCargo As String
Dim nMovNroAbono As Long, nMovNroCargo As Long, nMovNroRef
'Abona el monto total cobrado
Set oMov = New COMNContabilidad.NCOMContFunciones
sMovNroAbono = oMov.GeneraMovNro(dFecha, sAgencia, sUsuario)
Set oMov = Nothing
Set oCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
oCap.CapAbonoCuentaAho sCuenta, nMontoAbono, nOpeAbono, sMovNroAbono, sGlosa, , , , , , , , sNomCmac, sNomAge, sLpt
Set oCap = Nothing
Set oMov = New COMNContabilidad.NCOMContFunciones
sMovNroCargo = oMov.GeneraMovNro(dFecha, sAgencia, sUsuario)
Set oMov = Nothing
Set oCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
oCap.CapCargoCuentaAho sCuenta, nMontoCargo, nOpeCargo, sMovNroCargo, sGlosa, , , , , , , , sNomCmac, sNomAge, sLpt
Set oCap = Nothing

Set oserv = New COMDCaptaServicios.DCOMCaptaServicios
nMovNroAbono = oserv.GetnMovNro(sMovNroAbono)
nMovNroCargo = oserv.GetnMovNro(sMovNroCargo)
rsServ.MoveFirst
oserv.dbCmact.BeginTrans
Do While Not rsServ.EOF
    nMovNroRef = rsServ("nMovNro")
    oserv.AgregaMovRef nMovNroAbono, nMovNroRef
    rsServ.MoveNext
Loop
oserv.AgregaMovRef nMovNroCargo, nMovNroAbono
oserv.dbCmact.CommitTrans
Set oserv = Nothing
End Sub

'Public Sub oITF.fgITFImprimeBoleta(ByVal sNomCli As String, ByVal nMonITF As Double, ByVal psDesOpeOrigen As String, ByVal pdFecSis As Date, _
'            Optional ByVal nNumTran As Long = 1, Optional ByVal sTitBol As String = "IMP. TRANSAC. FINANCIERAS", _
'            Optional ByVal pnTipoPago As Integer = 1, Optional nMonedaITF As Integer = 1, Optional sCodCta As String = "", _
'            Optional sGlosa As String, Optional pCancelaImp As Boolean = True, Optional ByVal sImp As String = "", Optional ByVal sNomAg As String = "", _
'            Optional ByVal psNomAge As String = "")
'
'If pCancelaImp Then
'    Exit Sub 'cancela impresion de itf
'End If
'
'Dim oVarImpre As New COMFunciones.FCOMVarImpresion
'Dim oImpre As New COMFunciones.FCOMImpresion
'
'Dim sLpt As String
'
'If sImp <> "" Then
'   sLpt = sImp
'End If
'
'Dim psTit As String
'Dim sMonto As String
'Dim psDesOpe As String
'Dim nFicSal As Integer
'Dim sFecha As String
'Dim sHora As String
'Dim sSep As Integer
'Dim sIni As Integer
'Dim sMax  As Integer
'Dim saux As Integer
'Dim lsNegritaOn As String
'Dim lsNegritaOff As String
'Dim lnTope As Integer
'Dim lsNroExt As String
'Dim lnCliAux As Integer
'Dim lsNomAge As String
'
'Dim lnNumLinCmac As Integer
'
'Dim lsMensaje As String * 39
'Dim lsCliAux1 As String
'Dim lsCliAux2 As String
'Dim psTexto As String
'
'Dim psTextoTC1 As String
'Dim psTextoTC2 As String
'Dim psMontoSoles As String
'Dim sPonderadoSoles As String
'
''Moneda y Cuenta porque en algunos casos no se manda la moneda... JHVP
'If Len(sCodCta) = 18 Then
'    nMonedaITF = Mid(sCodCta, 9, 1)
'End If
'
'If Len(psNomAge) > 15 Then
'    psNomAge = Mid(psNomAge, 1, 15)
'End If
'
'If nMonedaITF = 1 Then
'    If pnTipoPago = 1 Then
'        psTexto = "ITF Efectivo"
'    ElseIf pnTipoPago = 2 Then
'        psTexto = "ITF Cargo a Cuenta"
'    End If
'ElseIf nMonedaITF = 2 Then
'    If pnTipoPago = 1 Then
'        psTexto = "ITF Efectivo Dolares"
'    ElseIf pnTipoPago = 2 Then
'        psTexto = "ITF Cargo a Cuenta Dolares"
'    End If
'End If
'
'lsNroExt = Str(nNumTran)
'
'ETIQ:
'
'On Error GoTo Error
'
'lnTope = 0 '6 'Tope de lineas en Boleta
'
'lsNegritaOn = oVarImpre.gPrnNegritaON
'lsNegritaOff = oVarImpre.gPrnNegritaOFF
'
'nFicSal = FreeFile
''If bsImp = False Then
'Open sLpt For Output As nFicSal
'
''End If
'
'Print #nFicSal, oVarImpre.gPrnInicializa;
'
''Print #nFicSal, gPrnEspaLinea6;   'espaciamiento lineas 1/6 pulg.
''Print #nFicSal, gPrnTamPagina22;  'Longitud de página a 22 líneas'
''Print #nFicSal, gPrnTamLetra10CPI;    'Tamaño 10 cpi
''Print #nFicSal, gPrnTpoLetraRoman;       'Tipo de Letra Sans Serif
''Print #nFicSal, gPrnCondensadaOFF ' cancela condensada
''Print #nFicSal, gPrnNegritaOFF ' desactiva negrita
'
'sSep = 15
'sIni = 1
'sMax = 33
'saux = 5
'
'
'sFecha = Format$(pdFecSis, "dd/mm/yyyy")
'sHora = Format$(Time, "hh:mm:ss")
'sMonto = Format$(nMonITF, "#,##0.00")
'
'
'lsNomAge = psNomAge
'
'If sNomAg <> "" Then
'lsNomAge = sNomAg
'End If
'
'If Len(lsNomAge) > 15 Then
'    lsNomAge = Mid(lsNomAge, 1, 15)
'End If
'
''Print #nFicSal, gPrnSaltoLinea;
''Print #nFicSal, ""
'Print #nFicSal, ""
'Print #nFicSal, lsNegritaOn; 'Activa Negrita
'Print #nFicSal, Tab(sIni); lsNegritaOn & "CMACICA"; Space(26 + sSep + saux); "CMACICA" & lsNegritaOff
'
'If nMonedaITF = 1 Then
'    Print #nFicSal, Tab(sIni); Trim(lsNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(lsNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(lsNomAge) & "-SOLES"; Space(saux + sMax - Len(Trim(lsNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
'Else
'    Print #nFicSal, Tab(sIni); Trim(lsNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(lsNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(lsNomAge) & "-DOLARES"; Space(saux + sMax - Len(Trim(lsNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
'End If
'
'Print #nFicSal, ""
'
'Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
'lnNumLinCmac = 0
'
'Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(saux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora
'
'lnCliAux = InStr(1, sNomCli, "*", vbTextCompare)
'
'If lnCliAux = 0 Then
'    If saux + sMax - Len(sNomCli) < 0 Then sNomCli = Mid(sNomCli, 1, sMax + saux)
'    Print #nFicSal, Tab(sIni); oImpre.ImpreCarEsp(sNomCli); Space(saux + sMax + sSep - Len(sNomCli)); oImpre.ImpreCarEsp(sNomCli)
'Else
'    lsCliAux1 = (Mid(sNomCli, 1, lnCliAux - 1))
'    lsCliAux2 = (Mid(sNomCli, lnCliAux + 1))
'
'    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
'    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
'
'    Print #nFicSal, Tab(sIni); oImpre.ImpreCarEsp(lsCliAux1); Space(saux + sMax + sSep - Len(lsCliAux1)); oImpre.ImpreCarEsp(lsCliAux1)
'    Print #nFicSal, Tab(sIni); oImpre.ImpreCarEsp(lsCliAux2); Space(saux + sMax + sSep - Len(lsCliAux2)); oImpre.ImpreCarEsp(lsCliAux2)
'
'    lnCliAux = 1
'End If
'
'If Len(Trim(sCodCta)) > 0 Then
'    sCodCta = Right("                       " & sCodCta, 18)
'    Print #nFicSal, Tab(sIni); "Cuenta:" & sCodCta; Space(8 + sSep + saux); "Cuenta:" & sCodCta
'Else
'    Print #nFicSal, ""
'End If
'
'psTit = Trim(sTitBol)
'psTit = oImpre.CentrarCadena(psTit, 28)
'Print #nFicSal, lsNegritaOn; 'Activa Negrita
'Print #nFicSal, Tab(sIni + 1); "-----" & psTit & "-----"; Space(-1 + sSep); "-----" & psTit & "-----"
'
'Print #nFicSal, Tab(sIni); oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(sMonto)); sMonto; Space(-1 + sSep); oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(sMonto)); sMonto
'
'If nMonedaITF = 1 Then
'    Print #nFicSal, ""  'JVP1
'ElseIf nMonedaITF = 2 Then
'    psTextoTC1 = "** Tipo de Cambio SBS"
'    If pnTipoPago = 1 Then
'        psTextoTC2 = "** ITF Efectivo Soles"
'    ElseIf pnTipoPago = 2 Then
'        psTextoTC2 = "** ITF Cargo a Cuenta Soles"
'    End If
'
'    If gTCPonderadoSBS = 0 Then
'        psMontoSoles = "0.00"
'        sPonderadoSoles = "0.00"
'    Else
'        psMontoSoles = Format(nMonITF * gTCPonderadoSBS, "0.00")
'        sPonderadoSoles = Format(gTCPonderadoSBS, "0.0000")
'    End If
'
'    Print #nFicSal, Tab(sIni); oImpre.ImpreCarEsp(Trim(Mid(psTextoTC1, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC1, 1, 28))) - Len(sPonderadoSoles)); sPonderadoSoles; Space(-1 + sSep); oImpre.ImpreCarEsp(Trim(Mid(psTextoTC1, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC1, 1, 28))) - Len(sPonderadoSoles)); sPonderadoSoles
'    Print #nFicSal, Tab(sIni); oImpre.ImpreCarEsp(Trim(Mid(psTextoTC2, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC2, 1, 28))) - Len(psMontoSoles)); psMontoSoles; Space(-1 + sSep); oImpre.ImpreCarEsp(Trim(Mid(psTextoTC2, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psTextoTC2, 1, 28))) - Len(psMontoSoles)); psMontoSoles
'
'    lnCliAux = lnCliAux + 1
'End If
'
'Print #nFicSal, lsNegritaOff; 'Desactiva Negrita 'JVP2 JVP3
'
'lsMensaje = "Operacion Origen:"
'Print #nFicSal, Tab(sIni); lsNegritaOn & oImpre.ImpreCarEsp(lsMensaje); Space(-1 + sSep); oImpre.ImpreCarEsp(lsMensaje); lsNegritaOff
'lsMensaje = psDesOpeOrigen
'Print #nFicSal, Tab(sIni); oImpre.ImpreCarEsp(lsMensaje); Space(-1 + sSep); oImpre.ImpreCarEsp(lsMensaje) 'JVP4
'
'Print #nFicSal, "" 'JVP5 JVP6
'
'Print #nFicSal, sGlosa & Space(sMax + 5 - Len(sGlosa) + sSep) & sGlosa ''JVP6 JVP7
'
'lnTope = 7 - lnCliAux 'JVP 3-0=3 3-1=2
'
'Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
'Print #nFicSal, Tab(sIni); oImpre.ImpreCarEsp(gsCodUser); Space(29 + sSep + saux); oImpre.ImpreCarEsp(gsCodUser)
'
'Dim clsGen As COMDConstSistema.DCOMGeneral
'Set clsGen = New COMDConstSistema.DCOMGeneral
'lsMensaje = clsGen.GetMensajeBoletas(sCodCta)
'Print #nFicSal, Tab(sIni); lsNegritaOn & oImpre.ImpreCarEsp(lsMensaje); Space(-1 + sSep); oImpre.ImpreCarEsp(lsMensaje); lsNegritaOff
'
'lnNumLinCmac = lnNumLinCmac + 1
'
'For saux = 1 To (lnTope - lnNumLinCmac) 'JVP DE 1 A 3-1=2 DE 1 A 2-1=1
'    Print #nFicSal, ""
'Next saux
'Close nFicSal
'Exit Sub
'Error:
'    Close nFicSal
'    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
'        GoTo ETIQ
'    End If
'
'End Sub

