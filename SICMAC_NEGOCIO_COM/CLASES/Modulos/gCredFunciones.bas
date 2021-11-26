Attribute VB_Name = "gCredFunciones"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function MontoTotalGastosGenerado(ByVal MatGastos As Variant, ByVal pnNumGastosCancel As Integer, _
    Optional ByVal psTipoGastoProc As Variant = "") As Double
Dim i As Integer
    MontoTotalGastosGenerado = 0
    For i = 0 To pnNumGastosCancel - 1
        If MatGastos(i, 4) = psTipoGastoProc(0) Or MatGastos(i, 4) = psTipoGastoProc(1) Or MatGastos(i, 4) = psTipoGastoProc(2) Then
            MontoTotalGastosGenerado = MontoTotalGastosGenerado + CDbl(MatGastos(i, 3))
        End If
    Next i

End Function

Public Function DiferencialMatricesMiVivienda(ByVal pMat1 As Variant, ByVal pMat2 As Variant) As Variant
Dim i As Integer
Dim MatResul As Variant
    ReDim MatResul(UBound(pMat1), 7)
    For i = 0 To UBound(pMat1) - 1
        MatResul(i, 0) = pMat1(i, 0) 'fecha
        MatResul(i, 1) = pMat1(i, 1) 'Cuota
        MatResul(i, 2) = pMat1(i, 2) 'Monto Cuota
        MatResul(i, 3) = Format(CDbl(pMat2(i, 3)) - CDbl(pMat1(i, 3)), "#0.00") 'Capital
        MatResul(i, 4) = Format(CDbl(pMat2(i, 4)) - CDbl(pMat1(i, 4)), "#0.00") 'Interes
        'MatResul(i, 5) = Format(CDbl(pMat2(i, 5)) - CDbl(pMat1(i, 5)), "#0.00") 'Gracia
        'MatResul(i, 6) = Format(CDbl(pMat2(i, 6)) - CDbl(pMat1(i, 6)), "#0.00") 'Gasto
    Next i
    DiferencialMatricesMiVivienda = MatResul
End Function

Public Function UnirMatricesMiViviendaAmortizacion(ByVal pMat1 As Variant, ByVal pMat2 As Variant) As Variant
Dim i, J, K As Integer
Dim MatResul As Variant
Dim nMonto As Double

    ReDim MatResul(UBound(pMat1), 13)
    For i = 0 To UBound(pMat1) - 1
        MatResul(i, 0) = pMat1(i, 0) 'fecha
        MatResul(i, 1) = pMat1(i, 1) 'Cuota
        MatResul(i, 2) = pMat1(i, 2) 'Monto Cuota
        For J = 3 To 12 'unimos concepto por concepto
            nMonto = 0
            For K = 0 To UBound(pMat2) - 1 'buscamos su cuota equivalente en calendatio paralelo
                If pMat1(i, 1) = pMat2(K, 1) Then 'si se encuentra la fila de la cuota
                    nMonto = CDbl(pMat2(K, J))
                    Exit For
                End If
            Next K
            MatResul(i, J) = Format(CDbl(pMat1(i, J)) + nMonto, "#0.00")
        Next J
        'MatResul(i, 3) = Format(CDbl(pMat1(i, 3)) + CDbl(pMat2(i, 3)), "#0.00") 'Capital
        'MatResul(i, 4) = Format(CDbl(pMat1(i, 4)) + CDbl(pMat2(i, 4)), "#0.00") 'Interes
        'MatResul(i, 5) = Format(CDbl(pMat1(i, 5)) + CDbl(pMat2(i, 5)), "#0.00") 'Interes Gracia
        'MatResul(i, 6) = Format(CDbl(pMat1(i, 6)) + CDbl(pMat2(i, 6)), "#0.00") 'Interes Mora
        'MatResul(i, 7) = Format(CDbl(pMat1(i, 7)) + CDbl(pMat2(i, 7)), "#0.00") 'Interes Reprog
        'MatResul(i, 8) = Format(CDbl(pMat1(i, 8)) + CDbl(pMat2(i, 8)), "#0.00") 'Interes Suspenso
        'MatResul(i, 9) = Format(CDbl(pMat1(i, 9)) + CDbl(pMat2(i, 9)), "#0.00") 'Interes Gasto
        'MatResul(i, 10) = Format(CDbl(pMat1(i, 10)) + CDbl(pMat2(i, 10)), "#0.00") 'Saldo
    Next i
    
    UnirMatricesMiViviendaAmortizacion = MatResul
    
End Function

Public Function UnirMatricesCalendarioIguales(ByVal pMat1 As Variant, ByVal pMat2 As Variant, ByVal pnMonto As Double)
Dim i As Integer
Dim nMontoSaldo As Double
Dim pMat2Temp As Variant
        pMat2Temp = pMat2
        For i = 0 To UBound(pMat1) - 1
            pMat2Temp(i, 3) = Format(CDbl(pMat1(i, 3)) + CDbl(pMat2(i, 3)), "#0.00") 'Capital
            pMat2Temp(i, 4) = Format(CDbl(pMat1(i, 4)) + CDbl(pMat2(i, 4)), "#0.00") 'Interes
            pMat2Temp(i, 5) = Format(CDbl(pMat1(i, 5)) + CDbl(pMat2(i, 5)), "#0.00") 'Gracia
            pMat2Temp(i, 6) = Format(CDbl(pMat1(i, 6)) + CDbl(pMat2(i, 6)), "#0.00") 'Gasto
            
            pMat2Temp(i, 2) = Format(CDbl(pMat2Temp(i, 3)) + CDbl(pMat2Temp(i, 4)) + CDbl(pMat2Temp(i, 5)) + CDbl(pMat2Temp(i, 6)), "#0.00") 'Monto Cuota
        Next i
        nMontoSaldo = pnMonto
        For i = 0 To UBound(pMat2) - 2
            nMontoSaldo = Format(nMontoSaldo - CDbl(pMat2Temp(i, 3)), "#0.00")
            pMat2Temp(i, 7) = Format(nMontoSaldo, "#0.00")
        Next i
        
        pMat2Temp(UBound(pMat2) - 1, 7) = "0.00"
        
        UnirMatricesCalendarioIguales = pMat2Temp
        
End Function


Public Function UnirMatricesMiVivienda(ByVal pMat1 As Variant, ByVal pMat2 As Variant, _
    ByVal pnMonto As Double) As Variant
Dim i As Integer
Dim MatResul() As String
Dim nNumProrat As Integer
Dim nMontoTemp As Double
Dim nCuotaTemp As Double
Dim nMontoTotal As Double
Dim nIndS As Integer
        nMontoTemp = pnMonto
        ReDim MatResul(UBound(pMat1), 8)
        nNumProrat = 0
        nIndS = -1
        For i = 0 To UBound(pMat1) - 1
            If i < 6 Then
                MatResul(i, 0) = pMat1(i, 0) 'fecha
                MatResul(i, 1) = pMat1(i, 1) 'Cuota
                MatResul(i, 2) = pMat1(i, 2) 'Monto Cuota
                MatResul(i, 3) = pMat1(i, 3) 'Capital
                MatResul(i, 4) = pMat1(i, 4) 'Interes
                MatResul(i, 5) = pMat1(i, 5) 'Gracia
                MatResul(i, 6) = pMat1(i, 6) 'Gasto
                'MatResul(i, 7) = pMat1(i, 7) 'Saldo
            Else
                MatResul(i, 0) = pMat1(i, 0) 'fecha
                MatResul(i, 1) = pMat1(i, 1) 'Cuota
                MatResul(i, 3) = Format(CDbl(pMat1(i, 3)) + CDbl(pMat2(nIndS, 3)) / nNumProrat, "#0.00") 'Capital
                MatResul(i, 4) = Format(CDbl(pMat1(i, 4)) + CDbl(pMat2(nIndS, 4)) / nNumProrat, "#0.00") 'Interes
                MatResul(i, 5) = Format(CDbl(pMat1(i, 5)) + CDbl(pMat2(nIndS, 5)) / nNumProrat, "#0.00") 'Gracia
                MatResul(i, 6) = Format(CDbl(pMat1(i, 6)) + CDbl(pMat2(nIndS, 6)) / nNumProrat, "#0.00") 'Gasto
                MatResul(i, 2) = Format(CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)), "#0.00")
                'MatResul(i, 7) = pMat1(i, 7) 'Saldo
                
                nCuotaTemp = Format(CDbl(pMat1(i, 2)) + CDbl(pMat2(nIndS, 2)) / nNumProrat, "#0.00") 'Monto Cuota
                If nCuotaTemp <> (CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6))) Then
                    MatResul(i, 4) = Format(CDbl(MatResul(i, 4)) + (nCuotaTemp - (CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)))), "#0.00")
                    MatResul(i, 2) = Format(CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)), "#0.00")
                End If
                
            End If
                        
            
            If (i + 1) Mod 6 = 0 Then
                If i <> 0 Then
                    nIndS = nIndS + 1
                End If
                If (UBound(pMat1) - i) >= 6 Then
                    nNumProrat = 6
                Else
                    nNumProrat = UBound(pMat1) - i
                End If
            End If
        Next i
                
        'Agregando la ultima Cuota Semestral
        MatResul(UBound(MatResul) - 1, 3) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(pMat2(UBound(pMat2) - 1, 3)), "#0.00")  'Capital
        MatResul(UBound(MatResul) - 1, 4) = Format(CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(pMat2(UBound(pMat2) - 1, 4)), "#0.00")  'Interes
        MatResul(UBound(MatResul) - 1, 5) = Format(CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(pMat2(UBound(pMat2) - 1, 5)), "#0.00")  'Gracia
        MatResul(UBound(MatResul) - 1, 6) = Format(CDbl(MatResul(UBound(MatResul) - 1, 6)) + CDbl(pMat2(UBound(pMat2) - 1, 6)), "#0.00")  'Gasto
        MatResul(UBound(MatResul) - 1, 2) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(MatResul(UBound(MatResul) - 1, 6)), "#0.00")
                
        'Comprobando Capital Total sea Igual a Prestamo
        nMontoTotal = 0
        For i = 0 To UBound(MatResul) - 1
            nMontoTotal = nMontoTotal + CDbl(MatResul(i, 3))
        Next i
        
        If nMontoTotal <> pnMonto Then
            MatResul(UBound(MatResul) - 1, 3) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) - (nMontoTotal - pnMonto), "#0.00")
            MatResul(UBound(MatResul) - 1, 2) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(MatResul(UBound(MatResul) - 1, 6)), "#0.00")
        End If
        
        nMontoTemp = pnMonto
        For i = 0 To UBound(MatResul) - 1
            nMontoTemp = nMontoTemp - CDbl(MatResul(i, 3))
            nMontoTemp = Format(nMontoTemp, "#0.00")
            MatResul(i, 7) = Format(nMontoTemp, "#0.00")
        Next i
        
        
        UnirMatricesMiVivienda = MatResul
End Function

Public Function UnirMatricesMiViviendaReprogramado(ByVal pMat1 As Variant, ByVal pMat2 As Variant, _
    ByVal pnMonto As Double, ByVal nCalifBuenPagador As Integer, ByVal pnFinalCuota As Integer) As Variant
Dim i As Integer
Dim MatResul() As String
Dim nNumProrat As Integer
Dim nMontoTemp As Double
Dim nCuotaTemp As Double
Dim nMontoTotal As Double
Dim nIndS As Integer
Dim K As Integer

        nMontoTemp = pnMonto
        ReDim MatResul(UBound(pMat1), 8)
        nNumProrat = 0
        nIndS = 0
        K = -10
        For i = 0 To UBound(pMat1) - 1
            If i < pnFinalCuota And nCalifBuenPagador = 1 Then
                MatResul(i, 0) = pMat1(i, 0) 'fecha
                MatResul(i, 1) = pMat1(i, 1) 'Cuota
                MatResul(i, 2) = pMat1(i, 2) 'Monto Cuota
                MatResul(i, 3) = pMat1(i, 3) 'Capital
                MatResul(i, 4) = pMat1(i, 4) 'Interes
                MatResul(i, 5) = pMat1(i, 5) 'Gracia
                MatResul(i, 6) = pMat1(i, 6) 'Gasto
                'MatResul(i, 7) = pMat1(i, 7) 'Saldo
            Else
                If nNumProrat = 0 Then
                    nNumProrat = 6
                    K = 0
                End If
                MatResul(i, 0) = pMat1(i, 0) 'fecha
                MatResul(i, 1) = pMat1(i, 1) 'Cuota
                MatResul(i, 3) = Format(CDbl(pMat1(i, 3)) + CDbl(pMat2(nIndS, 3)) / nNumProrat, "#0.00") 'Capital
                MatResul(i, 4) = Format(CDbl(pMat1(i, 4)) + CDbl(pMat2(nIndS, 4)) / nNumProrat, "#0.00") 'Interes
                MatResul(i, 5) = Format(CDbl(pMat1(i, 5)) + CDbl(pMat2(nIndS, 5)) / nNumProrat, "#0.00") 'Gracia
                MatResul(i, 6) = Format(CDbl(pMat1(i, 6)) + CDbl(pMat2(nIndS, 6)) / nNumProrat, "#0.00") 'Gasto
                MatResul(i, 2) = Format(CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)), "#0.00")
                'MatResul(i, 7) = pMat1(i, 7) 'Saldo
                
                nCuotaTemp = Format(CDbl(pMat1(i, 2)) + CDbl(pMat2(nIndS, 2)) / nNumProrat, "#0.00") 'Monto Cuota
                If nCuotaTemp <> (CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6))) Then
                    MatResul(i, 4) = Format(CDbl(MatResul(i, 4)) + (nCuotaTemp - (CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)))), "#0.00")
                    MatResul(i, 2) = Format(CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)), "#0.00")
                End If
                
            End If
            
            If (K + 1) Mod 6 = 0 Then
                If K <> 0 Then
                    nIndS = nIndS + 1
                End If
                If (UBound(pMat1) - K) >= 6 Then
                    nNumProrat = 6
                Else
                    nNumProrat = UBound(pMat1) - K
                End If
            End If
            If K <> -10 Then
                K = K + 1
            End If
        Next i
                
        'Agregando la ultima Cuota Semestral
        'MatResul(UBound(MatResul) - 1, 3) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(pMat2(UBound(pMat2) - 1, 3)), "#0.00")  'Capital
        'MatResul(UBound(MatResul) - 1, 4) = Format(CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(pMat2(UBound(pMat2) - 1, 4)), "#0.00")  'Interes
        'MatResul(UBound(MatResul) - 1, 5) = Format(CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(pMat2(UBound(pMat2) - 1, 5)), "#0.00")  'Gracia
        'MatResul(UBound(MatResul) - 1, 6) = Format(CDbl(MatResul(UBound(MatResul) - 1, 6)) + CDbl(pMat2(UBound(pMat2) - 1, 6)), "#0.00")  'Gasto
        'MatResul(UBound(MatResul) - 1, 2) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(MatResul(UBound(MatResul) - 1, 6)), "#0.00")
                
        'Comprobando Capital Total sea Igual a Prestamo
        nMontoTotal = 0
        For i = 0 To UBound(MatResul) - 1
            nMontoTotal = nMontoTotal + CDbl(MatResul(i, 3))
        Next i
        
        If nMontoTotal <> pnMonto Then
            MatResul(UBound(MatResul) - 1, 3) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) - (nMontoTotal - pnMonto), "#0.00")
            MatResul(UBound(MatResul) - 1, 2) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(MatResul(UBound(MatResul) - 1, 6)), "#0.00")
        End If
        
        nMontoTemp = pnMonto
        For i = 0 To UBound(MatResul) - 1
            nMontoTemp = nMontoTemp - CDbl(MatResul(i, 3))
            nMontoTemp = Format(nMontoTemp, "#0.00")
            MatResul(i, 7) = Format(nMontoTemp, "#0.00")
        Next i
        
        
        UnirMatricesMiViviendaReprogramado = MatResul
End Function


