Attribute VB_Name = "gCapFunciones"
Option Explicit

Public Function CabeRepoCaptac(ByVal sCabe01 As String, ByVal sCabe02 As String, _
        ByVal nCarLin As Long, ByVal sSeccion As String, ByVal sTitRp1 As String, _
        ByVal sTitRp2 As String, ByVal sMoneda As String, ByVal sNumPag As String, _
        ByVal sNomAge As String, ByVal dFecSis As Date) As String

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
    
CabeRepoCaptac = CabeRepoCaptac & sCabe01 & Chr$(10)
CabeRepoCaptac = CabeRepoCaptac & sCabe02 & Chr$(10)
CabeRepoCaptac = CabeRepoCaptac & sTit1 & Chr$(10)
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
    
CabeRepoCaptacSM = CabeRepoCaptacSM & Space(30) & sCabe01 & Chr$(10)
CabeRepoCaptacSM = CabeRepoCaptacSM & Space(30) & sCabe02 & Chr$(10)
CabeRepoCaptacSM = CabeRepoCaptacSM & sTit1 & Chr$(10)
CabeRepoCaptacSM = CabeRepoCaptacSM & sTit2
End Function



Public Sub ImprimeCertificadoPlazoFijo(ByVal dApertura As Date, ByVal sNomCli As String, _
        ByVal sDirCli As String, ByVal sCuenta As String, nExtracto As String, _
        ByVal nPlazo As Long, ByVal nSaldo As Double, ByVal nTasa As Double, _
        ByVal sRetInt As String, ByVal sLetras As String, Optional nDuplicado As Integer = 0, Optional ByVal bCopia As Boolean = False, _
        Optional ByVal sTipoCta As String)
        
Dim intLinBla As Integer, nFicSal As Integer
Dim sFchLet As String, sFchVct As String
Dim sFchCan As String, sForMon As String
Dim sVctLet As String, dVencimiento As Date, dCancelacion As Date
Dim sTaInNo As String, sTaInEf As String
Dim sTaInAn As String, sPlazo As String
Dim I As Integer
sFchLet = ArmaFecha(dApertura)
dVencimiento = DateAdd("d", nPlazo, dApertura)
dCancelacion = DateAdd("d", nPlazo + 1, dApertura)

sVctLet = ArmaFecha(dVencimiento)
sFchCan = ArmaFecha(dCancelacion)
sForMon = IIf(Right(sLetras, 5) = "SOLES", "S/. ", "US$ ")

sTaInNo = Format$((nTasa / 12), "#0.00")
sTaInEf = Format$(((((((nTasa / 12) / 3000) + 1) ^ 30) - 1) * 100), "#0.00")
sTaInAn = Format(((((((nTasa / 12) / 3000) + 1) ^ 360) - 1) * 100), "#0.00")

sPlazo = Trim(nPlazo) & " dia(s)"

nFicSal = FreeFile
Open sLpt For Output As nFicSal

Print #nFicSal, Chr$(15);                           'Retorna al tipo de letra normal
Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(1);     'Tipo de Letra Roman
Print #nFicSal, Chr$(27) + Chr$(103);                'Tamaño 10.5 - 15 CPI
Print #nFicSal, Chr$(27) + Chr$(50);                'Espaciado entre lineas 1/16
'*****agregado--
'*
If bCopia = True Then
    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(17)  'longitud de 17 lineas
Else
    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(16)
'    Print #nFicSal, Chr$(27) + Chr$(67) + Chr$(18);     'Longitud de página a 24 líneas
End If

'****agregado

Print #nFicSal, Chr$(27) + Chr$(108) + Chr$(6);     'Margen Izquierdo - 6ta. Columna

For I = 1 To 6
    Print #nFicSal, ""     '   "Fila No. " & Trim(Str(i))
Next I

'       Impresión
'Print #nFicSal, Tab(43); UCase(sFchLet)
Print #nFicSal, Chr$(27) + Chr$(69);                           'Establece tipo de letra negrita
Print #nFicSal, Tab(83); UCase(sFchLet)
Print #nFicSal, ""
Print #nFicSal, ImpreCarEsp(sNomCli)      '  psNomCli
'Print #nFicSal, sDirCli
'Print #nFicSal, ""
Print #nFicSal, "PLAZO FIJO " & Space(24) & sCuenta; Tab(57); FillNum(Trim(nExtracto), 5, " ")
Print #nFicSal, "TIPO DE CUENTA                 :   " & sTipoCta
Print #nFicSal, "DEPOSITO A PLAZO FIJO " & Space(9) & ":" & Space(3) & Trim(nPlazo) & " dias"; Tab(54); Chr$(27) + Chr$(69); sForMon & JDNum(Trim(nSaldo), 12, True, 9, 2) '; Chr$(27) + Chr$(70)
Print #nFicSal, "FECHA DE VENCIMIENTO               " & UCase(sVctLet)
Print #nFicSal, "Y/O CANCELACION                :   " & UCase(sFchCan)
Print #nFicSal, "" '"TASA DE INTERES NOMINAL        :   " & sTaInNo & "%  MENSUAL"
Print #nFicSal, "TASA DE INTERES EFECTIVA       :   " & sTaInEf & "%  MENSUAL "
Print #nFicSal, "FRECUENCIA DE CAPITALIZACION   :   DIARIA"
Print #nFicSal, "LOS INTERESES SE ABONARAN      :   " & Left(sRetInt, 20)
'Print #nFicSal, ""
'Print #nFicSal, Chr$(27) + Chr$(69);                           'Establece tipo de letra negrita
Print #nFicSal, "SON :" & sLetras  ' & Space(80 - Len(sLetras) - 5)
Print #nFicSal, "TASA DE INTERES EFECTIVA ANUAL :   " & sTaInAn & "%"
'Print #nFicSal, sLetras
'Print #nFicSal, Tab(54); 'sForMon & JDNum(Trim(nSaldo), 12, True, 9, 2)
'Print #nFicSal, ""
'Print #nFicSal, "TASA DE INTERES EFECTIVA ANUAL :   " & sTaInAn & "%"
Print #nFicSal, Chr$(27) + Chr$(70);
Print #nFicSal, ""
Print #nFicSal, Chr$(27) + Chr$(33) + Chr$(40);
Print #nFicSal, IIf(nDuplicado = 0, "", "DUPLICADO " & Trim(nDuplicado));
Print #nFicSal, Chr$(27) + Chr$(33) + Chr$(0);
Print #nFicSal, Chr$(27) + Chr$(70);                           'Desactiva tipo de letra negrita
Print #nFicSal, Chr$(12);                           'Avance de Página
Close nFicSal

End Sub


Public Function ConvierteTNAaTEA(ByVal nTasa As Double) As Double
ConvierteTNAaTEA = ((1 + nTasa / 36000) ^ 360 - 1) * 100
End Function

Public Function ConvierteTEAaTNA(ByVal nTasa As Double) As Double
ConvierteTEAaTNA = ((1 + nTasa / 100) ^ (1 / 360) - 1) * 36000
End Function

