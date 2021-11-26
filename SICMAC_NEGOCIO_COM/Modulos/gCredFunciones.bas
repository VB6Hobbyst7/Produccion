Attribute VB_Name = "gCredFunciones"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

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
Dim i, J, k As Integer
Dim MatResul As Variant
Dim nMonto As Double

    ReDim MatResul(UBound(pMat1), 13)
    For i = 0 To UBound(pMat1) - 1
        MatResul(i, 0) = pMat1(i, 0) 'fecha
        MatResul(i, 1) = pMat1(i, 1) 'Cuota
        MatResul(i, 2) = pMat1(i, 2) 'Monto Cuota
        For J = 3 To 12 'unimos concepto por concepto
            nMonto = 0
            For k = 0 To UBound(pMat2) - 1 'buscamos su cuota equivalente en calendatio paralelo
                If pMat1(i, 1) = pMat2(k, 1) Then 'si se encuentra la fila de la cuota
                    nMonto = CDbl(pMat2(k, J))
                    Exit For
                End If
            Next k
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
    ByVal pnMonto As Double, Optional ByVal psTipoSegDes As String = "", Optional ByVal pnTramoNoConsPorcen As Currency = 0) As Variant 'MAVM 20120828 psTipoSegDes
Dim i As Integer
Dim MatResul() As String
Dim nNumProrat As Integer
Dim nMontoTemp As Double
Dim nCuotaTemp As Double
Dim nMontoTotal As Double
Dim nIndS As Integer

'MAVM 20120828
Dim nValorSegDes As Double
Dim nTramoCons As Double 'MAVM 20121113
Dim pnMontoPivot As Double
Dim rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim sql As String
Set oConect = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If oConect.AbreConexion = False Then Exit Function
    sql = "Select nValor, nValorDosTit From ProductoConcepto Where nPrdConceptoCod = '1217'"
Set rs = oConect.Ejecutar(sql)
If Not rs.EOF Then
    If psTipoSegDes = "I" Then
        nValorSegDes = rs.Fields(0)
    Else
        nValorSegDes = rs.Fields(1)
    End If
End If
rs.Close
Set rs = Nothing
oConect.CierraConexion
Set oConect = Nothing
'***

        nMontoTemp = pnMonto
        pnMontoPivot = pnMonto 'MAVM 20120828
        'ReDim MatResul(UBound(pMat1), 8)
        ReDim MatResul(UBound(pMat1), 10) 'MAVM 20120630 ***
        If pnTramoNoConsPorcen = 0 Then
            nTramoCons = 12500 'MAVM 20121113
        Else
            nTramoCons = pnTramoNoConsPorcen 'ALPA 20121113
        End If
        nNumProrat = 0
        nIndS = -1
        For i = 0 To UBound(pMat1) - 1
            If i < 6 Then
                MatResul(i, 0) = pMat1(i, 0) 'fecha
                MatResul(i, 1) = pMat1(i, 1) 'Cuota
                'MatResul(i, 2) = pMat1(i, 2) 'Monto Cuota
                MatResul(i, 3) = pMat1(i, 3) 'Capital
                MatResul(i, 4) = pMat1(i, 4) 'Interes
                MatResul(i, 5) = pMat1(i, 5) 'Gracia
                'MatResul(i, 6) = pMat1(i, 6) 'Gasto
                'MatResul(i, 7) = pMat1(i, 7) 'Saldo
                
                'MAVM 20120828 ***
                MatResul(i, 8) = pMat1(i, 8) 'Gasto
                If i = 0 Then
                    'ALPA 20140618****************************************
                    'MatResul(i, 6) = Format(Format(nValorSegDes / 100, "#0.000000") * (CDbl(pnMontoPivot) + CDbl(pMat1(i, 4))), "#0.00") 'Seg Desg
                    MatResul(i, 6) = Format(Format(nValorSegDes / 100, "#0.000000") * (CDbl(pnMontoPivot)), "#0.00") 'Seg Desg
                Else
                    'ALPA 20140618****************************************
                    'MatResul(i, 6) = Format(Format(nValorSegDes / 100, "#0.000000") * (CDbl(pnMontoPivot) + CDbl(pMat1(i, 4))), "#0.00") 'Seg Desg
                    MatResul(i, 6) = Format(Format(nValorSegDes / 100, "#0.000000") * (CDbl(pnMontoPivot)), "#0.00")  'Seg Desg
                End If
                pnMontoPivot = pnMontoPivot - pMat1(i, 3) 'MAVM 20120828
                MatResul(i, 9) = pMat1(i, 9) 'Seg Inmueb
                MatResul(i, 2) = Format(CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)) + CDbl(MatResul(i, 8)) + CDbl(MatResul(i, 9)), "#0.00")
                '***
            Else
                MatResul(i, 0) = pMat1(i, 0) 'fecha
                MatResul(i, 1) = pMat1(i, 1) 'Cuota
                'MatResul(i, 3) = Format(CDbl(pMat1(i, 3)) + CDbl(pMat2(nIndS, 3)) / nNumProrat, "#0.00") 'Capital
                'MatResul(i, 4) = Format(CDbl(pMat1(i, 4)) + CDbl(pMat2(nIndS, 4)) / nNumProrat, "#0.00") 'Interes
                'MatResul(i, 5) = Format(CDbl(pMat1(i, 5)) + CDbl(pMat2(nIndS, 5)) / nNumProrat, "#0.00") 'Gracia
                'MatResul(i, 6) = Format(CDbl(pMat1(i, 6)) + CDbl(pMat2(nIndS, 6)) / nNumProrat, "#0.00") 'Gasto
                'MatResul(i, 2) = Format(CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)), "#0.00")
                'MatResul(i, 7) = pMat1(i, 7) 'Saldo
                
                'MAVM 20120709 ***
                MatResul(i, 3) = Format(CDbl(pMat1(i, 3)) + CDbl(pMat2(i, 3)), "#0.00") 'Capital
                MatResul(i, 4) = Format(CDbl(pMat1(i, 4)) + CDbl(pMat2(i, 4)), "#0.00") 'Interes
                MatResul(i, 5) = Format(CDbl(pMat1(i, 5)) + CDbl(pMat2(i, 5)), "#0.00") 'Gracia
                'ALPA20140619*******************
                'MatResul(i, 8) = Format(CDbl(pMat1(i, 8)) + CDbl(pMat2(i, 8)) + ((nTramoCons * 0.02081) / 100), "#0.00") 'Gasto
                MatResul(i, 8) = CDbl(pMat2(i, 8))
                '***********************
                nTramoCons = nTramoCons - Format(CDbl(pMat2(i, 3)), "#0.00")
                MatResul(i, 6) = Format(Format(nValorSegDes / 100, "#0.000000") * (CDbl(pnMontoPivot) + CDbl(MatResul(i, 4))), "#0.00")
                MatResul(i, 9) = Format(CDbl(pMat1(i, 9)) + CDbl(pMat2(nIndS, 9)), "#0.00") 'Seg Inmueb
                MatResul(i, 2) = Format(CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)) + CDbl(MatResul(i, 8)) + CDbl(MatResul(i, 9)), "#0.00")
                pnMontoPivot = pnMontoPivot - MatResul(i, 3)
                '***
                
                'nCuotaTemp = Format(CDbl(pMat1(i, 2)) + CDbl(pMat2(nIndS, 2)) / nNumProrat, "#0.00") 'Monto Cuota
                'If nCuotaTemp <> (CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6))) Then
                    'MatResul(i, 4) = Format(CDbl(MatResul(i, 4)) + (nCuotaTemp - (CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)))), "#0.00")
                    'MatResul(i, 2) = Format(CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)), "#0.00")
                'End If
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
        'MatResul(UBound(MatResul) - 1, 3) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(pMat2(UBound(pMat2) - 1, 3)), "#0.00")  'Capital
        'MatResul(UBound(MatResul) - 1, 4) = Format(CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(pMat2(UBound(pMat2) - 1, 4)), "#0.00")  'Interes
        'MatResul(UBound(MatResul) - 1, 5) = Format(CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(pMat2(UBound(pMat2) - 1, 5)), "#0.00")  'Gracia
        
        'MatResul(UBound(MatResul) - 1, 6) = Format(CDbl(MatResul(UBound(MatResul) - 1, 6)) + CDbl(pMat2(UBound(pMat2) - 1, 6)), "#0.00")  'Gasto
        'MatResul(UBound(MatResul) - 1, 2) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(MatResul(UBound(MatResul) - 1, 6)), "#0.00")
                
        'MatResul(UBound(MatResul) - 1, 6) = Format(CDbl(MatResul(UBound(MatResul) - 1, 8)) + CDbl(pMat2(UBound(pMat2) - 1, 8)), "#0.00")  'Seg Desg
        'MatResul(UBound(MatResul) - 1, 8) = Format(CDbl(MatResul(UBound(MatResul) - 1, 6)) + CDbl(pMat2(UBound(pMat2) - 1, 6)), "#0.00")  'Gasto
        'MatResul(UBound(MatResul) - 1, 9) = Format(CDbl(MatResul(UBound(MatResul) - 1, 9)) + CDbl(pMat2(UBound(pMat2) - 1, 9)), "#0.00")  'Seg Inm
        'MatResul(UBound(MatResul) - 1, 2) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(MatResul(UBound(MatResul) - 1, 6)) + CDbl(MatResul(UBound(MatResul) - 1, 8)) + CDbl(MatResul(UBound(MatResul) - 1, 9)), "#0.00")
                
        'Comprobando Capital Total sea Igual a Prestamo
        nMontoTotal = 0
        For i = 0 To UBound(MatResul) - 1
            nMontoTotal = nMontoTotal + CDbl(MatResul(i, 3))
        Next i
        
        If Format(nMontoTotal, "#0.00") <> Format(pnMonto, "#0.00") Then
            MatResul(UBound(MatResul) - 1, 3) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) - (nMontoTotal - pnMonto), "#0.00")
            'MatResul(UBound(MatResul) - 1, 2) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(MatResul(UBound(MatResul) - 1, 6)), "#0.00")
            MatResul(UBound(MatResul) - 1, 2) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(MatResul(UBound(MatResul) - 1, 6)) + CDbl(MatResul(UBound(MatResul) - 1, 8)) + CDbl(MatResul(UBound(MatResul) - 1, 9)), "#0.00")
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
Dim k As Integer

        nMontoTemp = pnMonto
        ReDim MatResul(UBound(pMat1), 8)
        nNumProrat = 0
        nIndS = 0
        k = -10
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
                    k = 0
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
            
            If (k + 1) Mod 6 = 0 Then
                If k <> 0 Then
                    nIndS = nIndS + 1
                End If
                If (UBound(pMat1) - k) >= 6 Then
                    nNumProrat = 6
                Else
                    nNumProrat = UBound(pMat1) - k
                End If
            End If
            If k <> -10 Then
                k = k + 1
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

'*** PEAC 20100723
Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "01", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    'On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    Dim sql As String
    Set oConect = New COMConecta.DCOMConecta
    Set rs = New ADODB.Recordset
    If oConect.AbreConexion = False Then Exit Function
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & oConect.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set rs = oConect.Ejecutar(sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
'GeneraMovNroErr:
'    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function
'Comentar en clases cada vez que se compila
'EJVG20160509 ***
Public Function EmiteInformeRiesgo(ByVal pnProceso As eProcesoEmiteInformeRiesgo, _
                                        ByVal psCtaCod As String, _
                                        Optional ByVal psTpoProdCod As String = "", _
                                        Optional ByVal psTpoCredCod As String = "", _
                                        Optional ByVal psPersCod As String = "", _
                                        Optional ByVal psPersNombre As String = "", _
                                        Optional ByVal pnMontoCol As Double = 0#, _
                                        Optional ByVal pbEsAmpliado As Boolean = False, _
                                        Optional ByVal pnNroCuotas As Integer = 0) As Boolean
    Dim rsMontoRiesgo As ADODB.Recordset 'JOEP-ERS064 20161028
    Dim oPar As COMDCredito.DCOMParametro
    Dim oNCredito As COMNCredito.NCOMCredito
    Dim oDCredito As COMDCredito.DCOMCredito
    Dim oDPersona As COMDpersona.DCOMPersona
    Dim oTipoCam As COMDConstSistema.NCOMTipoCambio
    Dim rsCredito As ADODB.Recordset
    Dim rsPersona As ADODB.Recordset
    Dim nMontoColTC As Double
    Dim bVerificaDPF As Boolean
    Dim nRiesgo11 As Double, nRiesgo12 As Double, nRiesgo21 As Double, nRiesgo22 As Double
    Dim X As Integer, Y As Integer
    Dim IR_nRiesgo_NEW As Integer
    Dim IR_nExposicion_NEW As Double
    Dim IR_nID As Long
    Dim IR_nRiesgo As Integer, IR_nNivel As Integer, IR_nEstado As Integer
    Dim IR_nExposicion As Double
    Dim IR_nNroCuotas As Integer
    Dim CR_nEstado As Integer
    Dim nTC As Double
    Dim nMontoAmpliado As Double
    Dim IAmp As Integer
    Dim lsMovNro As String
    Dim lnTipo As Integer
    'Dim lsMsg As String

    EmiteInformeRiesgo = True

    Set oDCredito = New COMDCredito.DCOMCredito
    If pnProceso = NivelAprobacion Then
        Set rsCredito = oDCredito.RecuperaColocacionesXInformeRiesgos(psCtaCod)
        If Not rsCredito.EOF Then
            psTpoProdCod = rsCredito!cTpoProdCod
            psTpoCredCod = rsCredito!cTpoCredCod
            psPersCod = rsCredito!cPersCod
            psPersNombre = rsCredito!cPersNombre
            pnMontoCol = rsCredito!nMontoCol
            pbEsAmpliado = rsCredito!bEsAmpliado
            pnNroCuotas = rsCredito!nNroCuotas
        End If
    End If

    IR_nEstado = -1
    CR_nEstado = -1

    lsMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

    Set oPar = New COMDCredito.DCOMParametro
    nRiesgo11 = oPar.RecuperaValorParametro(3207)
    nRiesgo12 = oPar.RecuperaValorParametro(3208)
    nRiesgo21 = oPar.RecuperaValorParametro(3205)
    nRiesgo22 = oPar.RecuperaValorParametro(3206)
    Set oPar = Nothing

    Set rsCredito = oDCredito.ObtenerComunicarRiesgo(psCtaCod)
    If rsCredito.RecordCount > 0 Then
        CR_nEstado = CInt(rsCredito!nEstado)
    End If

    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
    nTC = oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoMes)
    Set oTipoCam = Nothing

    Set oNCredito = New COMNCredito.NCOMCredito
    bVerificaDPF = oNCredito.VerificaGarantiaDPF(psCtaCod)

    'Comunicar Riesgo x RapiFlash o Garantías (todas) Autoliquidables en Crédito
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000080", psTpoProdCod) Or bVerificaDPF Then
    'If psTpoProdCod = "703" Or bVerificaDPF Then
    '**ARLO20180712 ERS042 - 2018
        nMontoColTC = CDbl(pnMontoCol) * CDbl(IIf(Mid(psCtaCod, 9, 1) = "1", 1, nTC))

        If pnProceso = NivelAprobacion Then
            Exit Function
        End If

        If nMontoColTC >= nRiesgo22 Then
            If CR_nEstado = -1 Or CR_nEstado = 0 Then
                If CR_nEstado = -1 Then
                    Call oDCredito.OpeComunicarRiesgo(psCtaCod, 3, , , 0)
                    Call oDCredito.OpeComunicarRiesgo(psCtaCod, 1, lsMovNro, , 0)
                End If
                If pnProceso = Aprobacion Then
                    MsgBox "Para la aprobación de este crédito se requiere la recepción del correo electrónico informativo de la Gerencia de Riesgos, verifíquelo e inténtelo de nuevo.", vbInformation
                Else
                    MsgBox "Debido al monto del crédito Rapiflash se requerirá que usted remita un correo electrónico a la Gerencia de Riesgos." & vbNewLine & _
                       "No se podrá aprobrar el crédito mientras no se recepcione el mencionado email.", vbInformation, "Aviso"
                End If
                If pnProceso = Aprobacion Then
                    EmiteInformeRiesgo = False
                End If
                Exit Function
            End If
            If CR_nEstado = 1 Then
                If pnProceso = Sugerencia Then
                    Call oDCredito.OpeComunicarRiesgo(psCtaCod, 1, lsMovNro, , 0)
                    MsgBox "Debido al monto del crédito Rapiflash se requerirá que usted remita un correo electrónico a la Gerencia de Riesgos." & vbNewLine & _
                       "No se podrá aprobrar el crédito mientras no se recepcione el mencionado email.", vbInformation, "Aviso"
                    Exit Function
                End If
            End If
        Else
            If CR_nEstado = 0 Then
                Call oDCredito.OpeComunicarRiesgo(psCtaCod, 3, , , 0)
                Exit Function
            End If
        End If
    'Informe de Riesgos
    Else
        Set oDPersona = New COMDpersona.DCOMPersona
        If psTpoCredCod = "152" Or psTpoCredCod = "252" Or psTpoCredCod = "352" Or psTpoCredCod = "452" Or psTpoCredCod = "552" Then
            lnTipo = 1 'AGROPECUARIO
        Else
            lnTipo = 2 'NO AGROPECUARIO
        End If

        IR_nRiesgo_NEW = oDPersona.ObtenerNivelRiesgo(psCtaCod, lnTipo, pnMontoCol)

        Set oDCredito = New COMDCredito.DCOMCredito
        IR_nExposicion_NEW = oDCredito.ObtenerMontoExposicionRiesgoUnico(psCtaCod, pnMontoCol)

        Set rsCredito = oDCredito.ObtenerInformeRiesgoNEW(psCtaCod)
        If rsCredito.RecordCount > 0 Then
            IR_nID = rsCredito!nInformeID
            IR_nRiesgo = rsCredito!nRiesgo
            IR_nNivel = rsCredito!nNivel
            IR_nEstado = rsCredito!nEstado
            IR_nExposicion = rsCredito!nExposicion
            IR_nNroCuotas = rsCredito!nNroCuotas
        End If

        'JOEP-ERS064 20161028
    Dim rsObtSaldObs As ADODB.Recordset
    Set rsObtSaldObs = oDCredito.ObtSaldObs(psCtaCod)

    If Not (rsObtSaldObs.EOF And rsObtSaldObs.BOF) Then
        If rsObtSaldObs!nObs = 0 Then '1:Reingreso,0:Salida
            IR_nEstado = 4
        End If
    End If

    Dim oCredSegRieg As COMDCredito.DCOMCredito
    Set oCredSegRieg = New COMDCredito.DCOMCredito

        Dim sValCatAge As String
        Dim sValorCriterio As String

        Dim rsMoraAge As ADODB.Recordset
        Dim rsCatAge As ADODB.Recordset
        Dim rsMoraCosecha As ADODB.Recordset
        Dim rsNivelAge As ADODB.Recordset
        Dim rsReslMoraCos As ADODB.Recordset
        'JOEP-ERS064 20161028

 'JOEP-ERS064 20161028
    Set rsMoraAge = oCredSegRieg.MostrarMorarAgencia(Format(gdFecSis, "yyyyMMdd"), gsCodAge)
    Set rsCatAge = oCredSegRieg.MostrarCategoriaAgencia(gsCodAge)
    Set rsMoraCosecha = oCredSegRieg.MostrarMoraCosechaRiesgo(psCtaCod)

        'Identificar Mora Agencia
        If Not (rsMoraAge.EOF And rsMoraAge.BOF) Then
            If Not (rsCatAge.EOF And rsCatAge.BOF) Then
                Set rsNivelAge = oCredSegRieg.IdentificarMoraAge(rsMoraAge!nValor, rsCatAge!nBaja, rsCatAge!nModerado, rsCatAge!nAlto, rsCatAge!nExtremo)
                    If rsNivelAge!nValor <> -1 Then
                        sValCatAge = rsNivelAge!nValor
                    Else
                        MsgBox "Favor de Verificar los Niveles de Riesgo de Agencia", vbInformation, "Aviso"
                        EmiteInformeRiesgo = False
                        Exit Function
                    End If
            Else
                MsgBox "Hubo un error, Verifique el Nivel de Riesgo de Agencia", vbInformation, "Aviso"
                EmiteInformeRiesgo = False
                Exit Function
            End If
        Else
            MsgBox "Hubo un error, Verifique su Mora Agencia", vbInformation, "Aviso"
            EmiteInformeRiesgo = False
            Exit Function
        End If

        If Not (rsMoraCosecha.EOF And rsMoraCosecha.BOF) Then
            Set rsReslMoraCos = oCredSegRieg.IdentificaMoraCos(rsMoraCosecha!nIndicador)
                sValorCriterio = rsReslMoraCos!nValor
        Else
            MsgBox "Hubo un error, Verifique la Mora Cosecha", vbInformation, "Aviso"
            EmiteInformeRiesgo = False
            Exit Function
        End If
    'JOEP-ERS064 20161028

        Set rsMontoRiesgo = oDCredito.MostrarMontoRiesgo(sValCatAge, sValorCriterio) 'JOEP-ERS064 20161028

        If IR_nExposicion_NEW >= IIf(IR_nRiesgo_NEW = 11, rsMontoRiesgo!nRiesgo1, IIf(IR_nRiesgo_NEW = 12, rsMontoRiesgo!nRiesgo2, IIf(IR_nRiesgo_NEW = 21, rsMontoRiesgo!nRiesgo1, IIf(IR_nRiesgo_NEW = 22, rsMontoRiesgo!nRiesgo2, IR_nExposicion_NEW + 1#)))) Then 'JOEP-ERS064 20161104

        'If IR_nExposicion_NEW >= IIf(IR_nRiesgo_NEW = 11, nRiesgo11, IIf(IR_nRiesgo_NEW = 12, nRiesgo12, IIf(IR_nRiesgo_NEW = 21, nRiesgo21, IIf(IR_nRiesgo_NEW = 22, nRiesgo22, IR_nExposicion_NEW + 1#)))) Then 'Comento JOEP20161104
            'No hay informe anterior
            If IR_nEstado = -1 Then
                Call oDCredito.InsertarInformeRiesgo(psCtaCod, lsMovNro, IR_nRiesgo_NEW, IR_nExposicion_NEW, pnProceso)
                MsgBox "Crédito necesitará Informe de Riesgo para su aprobación." & Chr(13) & Chr(13) & "- Monto Exposición: " & gcPEN_SIMBOLO & " " & Format(IR_nExposicion_NEW, gsFormatoNumeroView) & Chr(13) & "- Nivel de Riesgo: R" & Right(CStr(IR_nRiesgo_NEW), 1), vbInformation, "Aviso"

                If pnProceso = Aprobacion Then
                    EmiteInformeRiesgo = False
                    Exit Function
                End If
            'Informe Registrado
            ElseIf IR_nEstado = 0 Then
                Call oDCredito.ActualizarInformeRiesgo(psCtaCod, IR_nID, , IR_nRiesgo_NEW, , IR_nExposicion_NEW)
                MsgBox "Crédito necesitará Informe de Riesgo para su aprobación." & Chr(13) & Chr(13) & "- Monto Exposición: " & gcPEN_SIMBOLO & " " & Format(IR_nExposicion_NEW, gsFormatoNumeroView) & Chr(13) & "- Nivel de Riesgo: R" & Right(CStr(IR_nRiesgo_NEW), 1), vbInformation, "Aviso"

                If pnProceso = Aprobacion Then
                    EmiteInformeRiesgo = False
                    Exit Function
                End If
            'Informe Recepcionado
            ElseIf IR_nEstado = 1 Then
                Call oDCredito.ActualizarInformeRiesgo(psCtaCod, IR_nID, , IR_nRiesgo_NEW, , IR_nExposicion_NEW)
                MsgBox "Expediente del Crédito se encuentra en la Gerencia de Riesgos.", vbInformation, "Aviso"

                If pnProceso = Sugerencia Or pnProceso = Aprobacion Then
                    EmiteInformeRiesgo = False
                    Exit Function
                End If
            'Informe Aprobado
            ElseIf IR_nEstado = 2 Then
                If pnProceso = Sugerencia Then
                    MsgBox "Expediente del Crédito, ya se encuentra con Informe de Riesgo.", vbInformation, "Aviso"
                    EmiteInformeRiesgo = False
                    Exit Function
                ElseIf pnProceso = NivelAprobacion Then

                ElseIf pnProceso = Aprobacion Then
                    'Validamos Nivel de Riesgo
                    If IR_nNivel >= 3 Then
                        MsgBox "El crédito tiene un Nivel de Riesgo NO ACEPTABLE, comunicarse con la Gerencia de Riesgos para su atención.", vbInformation, "Aviso"
                        EmiteInformeRiesgo = False
                        Exit Function
                    End If
                    'Incremento de Nro. de Cuotas
                    If pnNroCuotas > IR_nNroCuotas Then
                        If MsgBox("El Plazo del crédito sería de " & pnNroCuotas & " cuota" & IIf(pnNroCuotas > 1, "s", "") & ", pero el Informe de Riesgos se emitió con un plazo de " & IR_nNroCuotas & " cuota" & IIf(IR_nNroCuotas > 1, "s", "") & "." & Chr(13) & Chr(13) & "El crédito necesitaría nuevo Informe de Riesgo para su aprobación." & Chr(13) & Chr(13) & "- Monto Exposición: " & gcPEN_SIMBOLO & " " & Format(IR_nExposicion_NEW, gsFormatoNumeroView) & Chr(13) & "- Nivel de Riesgo: R" & Right(CStr(IR_nRiesgo_NEW), 1) & Chr(13) & Chr(13) & "¿Está seguro de continuar?", vbInformation + vbYesNo, "Confirmación") = vbYes Then
                            Call oDCredito.ActualizarInformeRiesgo(psCtaCod, IR_nID, 3, , , IR_nExposicion_NEW, lsMovNro, , , , "El Nro. de cuotas ha variado de " & IR_nNroCuotas & " a " & pnNroCuotas)
                            Call oDCredito.InsertarInformeRiesgo(psCtaCod, lsMovNro, IR_nRiesgo_NEW, IR_nExposicion_NEW, pnProceso)
                        End If
                        EmiteInformeRiesgo = False
                        Exit Function
                    End If
                    'Incremento de la Exposición
                    If IR_nExposicion_NEW > IR_nExposicion Then
                        If MsgBox("El monto de Exposición ha variado de " & Format(IR_nExposicion, gsFormatoNumeroView) & " a " & Format(IR_nExposicion_NEW, gsFormatoNumeroView) & " del último Informe de Riesgos que se emitió." & Chr(13) & Chr(13) & "El crédito necesitaría nuevo Informe de Riesgo para su aprobación." & Chr(13) & Chr(13) & "- Monto Exposición: " & gcPEN_SIMBOLO & " " & Format(IR_nExposicion_NEW, gsFormatoNumeroView) & Chr(13) & "- Nivel de Riesgo: R" & Right(CStr(IR_nRiesgo_NEW), 1) & Chr(13) & Chr(13) & "¿Está seguro de continuar?", vbInformation + vbYesNo, "Confirmación") = vbYes Then
                            Call oDCredito.ActualizarInformeRiesgo(psCtaCod, IR_nID, 3, , , IR_nExposicion_NEW, lsMovNro, , , , "El monto de Exposición ha variado de " & Format(IR_nExposicion, gsFormatoNumeroView) & " a " & Format(IR_nExposicion_NEW, gsFormatoNumeroView))
                            Call oDCredito.InsertarInformeRiesgo(psCtaCod, lsMovNro, IR_nRiesgo_NEW, IR_nExposicion_NEW, pnProceso)
                        End If
                        EmiteInformeRiesgo = False
                        Exit Function
                    End If
                    'Nivel de Riesgo no haya variado de R2 a R1
                    If (IR_nRiesgo = 12 Or IR_nRiesgo = 22) And (IR_nRiesgo_NEW = 11 Or IR_nRiesgo_NEW = 21) Then
                        If MsgBox("El Nivel de Riesgo del Crédito ha variado de R2 a R1 del último Informe de Riesgos que se emitió." & Chr(13) & Chr(13) & "El crédito necesitaría nuevo Informe de Riesgo para su aprobación." & Chr(13) & Chr(13) & "- Monto Exposición: " & gcPEN_SIMBOLO & " " & Format(IR_nExposicion_NEW, gsFormatoNumeroView) & Chr(13) & "- Nivel de Riesgo: R" & Right(CStr(IR_nRiesgo_NEW), 1) & Chr(13) & Chr(13) & "¿Está seguro de continuar?", vbInformation + vbYesNo, "Confirmación") = vbYes Then
                            Call oDCredito.ActualizarInformeRiesgo(psCtaCod, IR_nID, 3, , , IR_nExposicion_NEW, lsMovNro, , , , "El Nivel de Riesgo del Crédito ha variado de R2 a R1.")
                            Call oDCredito.InsertarInformeRiesgo(psCtaCod, lsMovNro, IR_nRiesgo_NEW, IR_nExposicion_NEW, pnProceso)
                        End If
                        EmiteInformeRiesgo = False
                        Exit Function
                    End If
                End If
            'Informe Desestimado
            ElseIf IR_nEstado = 3 Then
                Call oDCredito.InsertarInformeRiesgo(psCtaCod, lsMovNro, IR_nRiesgo_NEW, IR_nExposicion_NEW, pnProceso)
                MsgBox "Crédito necesitará Informe de Riesgo para su aprobación." & Chr(13) & Chr(13) & "- Monto Exposición: " & gcPEN_SIMBOLO & " " & Format(IR_nExposicion_NEW, gsFormatoNumeroView) & Chr(13) & "- Nivel de Riesgo: R" & Right(CStr(IR_nRiesgo_NEW), 1), vbInformation, "Aviso"

                If pnProceso = Aprobacion Then
                    EmiteInformeRiesgo = False
                    Exit Function
                End If
        'Inicio JOEP-ERS64 20170619
            ElseIf IR_nEstado = 4 Then
                MsgBox "Crédito tiene Observaciones de Riesgo", vbInformation, "Aviso"
        'FIN JOEP-ERS64 20170619
            End If
        Else
            'No hay informe anterior
            If IR_nEstado = -1 Then
            'Informe Registrado
            ElseIf IR_nEstado = 0 Then
                Call oDCredito.EliminarInformeRiesgo(psCtaCod, IR_nID)
            'Informe Recepcionado
            ElseIf IR_nEstado = 1 Then
                Call oDCredito.ActualizarInformeRiesgo(psCtaCod, IR_nID, 3, , , , , , , "Desestimado porque ya no necesitará Informe de Riesgos.")
            'Informe Aprobado
            ElseIf IR_nEstado = 2 Then
            'Informe Desestimado
            ElseIf IR_nEstado = 3 Then

            End If
        End If
    End If

    RSClose rsPersona
    RSClose rsCredito
    Set oNCredito = Nothing
End Function

Public Function IntervinientesSonVinculados(ByVal psCtaCod As String, Optional ByVal psPersCod As String = "", Optional ByVal psPersNombre As String = "") As Boolean
    Dim oDCredito As New COMDCredito.DCOMCredito
    Dim rsPersona As ADODB.Recordset
    Dim lsPropietariosPendientes As String

    If psPersCod = "" Then
        Set rsPersona = oDCredito.RecuperaColocacionesXInformeRiesgos(psCtaCod)
        If Not rsPersona.EOF Then
            psPersCod = rsPersona!cPersCod
            psPersNombre = rsPersona!cPersNombre
        End If
    End If

    'Validamos que los propietarios estén como Vinculados en los GRUPOS ECONÓMICOS
    lsPropietariosPendientes = oDCredito.CadenaPropietarioPendientexVincularRiesgo(psCtaCod)
    Set oDCredito = Nothing
    If Len(lsPropietariosPendientes) > 0 Then
        MsgBox "Primero debe de registrar como vinculados del titular " & Chr(13) & UCase(psPersNombre) & " en el módulo de VINCULADOS Y GRUPOS ECONÓMICOS a las siguientes personas:" & Chr(13) & Chr(13) & lsPropietariosPendientes, vbInformation, "No se podrá continuar"
        Exit Function
    End If

    IntervinientesSonVinculados = True
End Function
Public Function GenerarDataExposicionRiesgoUnico(ByVal psCtaCod As String, Optional ByVal psPersCod As String = "", Optional ByVal psPersNombre As String = "") As Boolean
    Dim oDCredito As New COMDCredito.DCOMCredito
    Dim oDPersona As COMDpersona.DCOMPersona
    Dim rsPersona As ADODB.Recordset
    Dim rsCodPersonas As ADODB.Recordset
    Dim rsPersonasVin As ADODB.Recordset
    Dim sCodPersonas As String
    Dim bVinculados As Boolean
    Dim CantVinculados As Long, Recorrido As Long
    Dim a As Integer, X As Integer, Y As Integer

    If psPersCod = "" Then
        Set rsPersona = oDCredito.RecuperaColocacionesXInformeRiesgos(psCtaCod)
        If Not rsPersona.EOF Then
            psPersCod = rsPersona!cPersCod
            psPersNombre = rsPersona!cPersNombre
        End If
    End If

    If Not IntervinientesSonVinculados(psCtaCod, psPersCod, psPersNombre) Then
        Exit Function
    End If

    Set oDPersona = New COMDpersona.DCOMPersona
    'Eliminar y Regenerar la lista de vinculados
    Call oDPersona.EliminarVinculadosPersona(psCtaCod)

    'Generar Lista de Vinculados de Nivel 0
    Set rsPersona = oDPersona.ObtenerVinculadosPersona(psPersCod, psPersCod)
    'Registro de Titular (Nivel 0)
    Call oDPersona.RegistrarVinculadosPersona(psCtaCod, psPersCod, psPersCod, , 0, gdFecSis, Recorrido)

    'Registro de Vinculados (Nivel 1)
    If rsPersona.RecordCount > 0 Then
        For a = 0 To rsPersona.RecordCount - 1
            Call oDPersona.RegistrarVinculadosPersona(psCtaCod, psPersCod, Trim(rsPersona!cPersCod), , 1, gdFecSis, Recorrido)
            rsPersona.MoveNext
        Next a
        bVinculados = True
    End If

    Set rsPersona = oDPersona.ListaVinculadosPersona(psCtaCod)
    'Registro de Vinculados de Vinculados (Nivel 2)
    Do While bVinculados
        Recorrido = Recorrido + 1
        If rsPersona.RecordCount > 0 Then
            For X = 0 To rsPersona.RecordCount - 1
                Set rsCodPersonas = oDPersona.DelvolverCodViculados(psCtaCod)
                If Not rsCodPersonas.EOF Then
                    sCodPersonas = Trim(rsCodPersonas!CodPersonas) & "," & psPersCod 'Aumentamos el codigo del titular del credito
                End If

                Set rsPersonasVin = oDPersona.ObtenerVinculadosPersona(Trim(rsPersona!cPersCodVin), sCodPersonas)
                For Y = 0 To rsPersonasVin.RecordCount - 1
                    Call oDPersona.RegistrarVinculadosPersona(psCtaCod, Trim(rsPersona!cPersCodVin), Trim(rsPersonasVin!cPersCod), , 2, gdFecSis, Recorrido)
                    rsPersonasVin.MoveNext
                    CantVinculados = CantVinculados + 1
                Next Y
                rsPersona.MoveNext
            Next X

            If CantVinculados > 0 Then
                bVinculados = True
            Else
                bVinculados = False
            End If

            Set rsPersona = oDPersona.ListaVinculadosPersona(psCtaCod, 2, Recorrido)
            CantVinculados = 0
        Else
            Recorrido = 0
            bVinculados = False
        End If
    Loop

    GenerarDataExposicionRiesgoUnico = True
End Function
Public Function GenerarDataExposicionEsteCredito(ByVal psCtaCod As String, ByVal pnMontoCol As Double, ByRef pnExpEsteCred As Double) As Boolean
    Dim oDCred As New COMDCredito.DCOMCredito

    pnExpEsteCred = oDCred.ObtenerMontoExposicionEsteCredito(psCtaCod, pnMontoCol)

    GenerarDataExposicionEsteCredito = True
    Set oDCred = Nothing
End Function
Public Function NecesitaFormatoEvaluacion(ByVal psCtaCod As String, _
                                            ByVal pnPrdEstado As Integer, _
                                            ByVal pnProducto As Integer, _
                                            ByVal pnSubProducto As Integer, _
                                            ByVal pnExpoEsteCred_NEW As Double, _
                                            ByRef pbEliminarEvaluacion As Boolean, _
                                            Optional ByRef pnFormatoEliminado As Integer = -1, _
                                            Optional ByRef pnFormato_NEW As Integer = -1, _
                                            Optional ByVal psTpoCredCodPadre As String, _
                                            Optional ByVal pnTpIngr As Long) As Boolean
    'LUCV20181220 Agregó:pnFormatoEliminado, pnFormato_NEW, Anexo01 de Acta 199-2018

    Dim oEval As New COMDCredito.DCOMFormatosEval
    Dim rsCred As New ADODB.Recordset
    Dim nExpEsteCred As Double
    Dim nFormato As Integer, nFormato_NEW As Integer
    'CTI3 ERS0032020****************
    Dim lnTpIngr  As Long
    Dim lsTpoCredCodPadre As String
    Dim lnTpoCliente As Integer
    Dim rsFE As ADODB.Recordset
    'END****************************
    'pnPrdEstado->2000: Solicitado, 2001: Sugerido
    NecesitaFormatoEvaluacion = True

    pbEliminarEvaluacion = False

    On Error GoTo ErrNecesitaFormato

    Set rsCred = oEval.RecuperaFormatoEvaluacion(psCtaCod)
    If rsCred.RecordCount <= 0 Then
        'En la Solicitud no exigimos que tenga, o si es RapiFlash
        '**ARLO20180712 ERS042 - 2018
        Set objProducto = New COMDCredito.DCOMCredito
        If pnPrdEstado = 2000 Or objProducto.GetResultadoCondicionCatalogo("N0000085", pnSubProducto) Then
        'If pnPrdEstado = 2000 Or pnSubProducto = 703 Then
        '**ARLO20180712 ERS042 - 2018
            NecesitaFormatoEvaluacion = False
        'En la Sugerencia exigimos que tenga
        Else
            NecesitaFormatoEvaluacion = True
            MsgBox "Aún no se ha ingresado la Evaluación del Crédito, verifique..", vbInformation, "Aviso"
        End If
        Exit Function
    End If

    nFormato = rsCred!nCodForm
    nExpEsteCred = rsCred!nMontoExpCredito
    
    'CTI3 ERS0032020********************************************************************************
    Dim oDCreditos As COMDCredito.DCOMCreditos
    Set oDCreditos = New COMDCredito.DCOMCreditos
    Call oDCreditos.ActualizaCategoriaCredito(psCtaCod, psTpoCredCodPadre)
    Set rsFE = oEval.RecuperaCreditoxEvaluacion(psCtaCod)
    If Not (rsFE.BOF And rsFE.EOF) Then
        lnTpIngr = rsFE!nTpIngr
        lsTpoCredCodPadre = IIf(IsNull(rsFE!cTpoCredCodPadre), "000", rsFE!cTpoCredCodPadre)
    Else
         lsTpoCredCodPadre = psTpoCredCodPadre
         lnTpIngr = pnTpIngr
    End If

    If Left(lsTpoCredCodPadre, 1) = "4" Or Left(lsTpoCredCodPadre, 1) = "5" Or Left(lsTpoCredCodPadre, 1) = "7" Or Left(lsTpoCredCodPadre, 1) = "8" Or Left(lsTpoCredCodPadre, 1) = "0" Then
        lnTpoCliente = 1
    ElseIf Left(lsTpoCredCodPadre, 1) = "1" Or Left(lsTpoCredCodPadre, 1) = "2" Or Left(lsTpoCredCodPadre, 1) = "3" Then
        lnTpoCliente = 2
    ElseIf Left(psTpoCredCodPadre, 1) = "0" Or Trim(psTpoCredCodPadre) = "" Then
        lnTpoCliente = 3
    End If
    nFormato_NEW = oEval.AsignarFormato(pnProducto, pnSubProducto, pnExpoEsteCred_NEW, lnTpoCliente, lnTpIngr)
    'nFormato_NEW = oEval.AsignarFormato(pnProducto, pnSubProducto, pnExpoEsteCred_NEW)
    '*********************************************************************************************
'    If nFormato <> 7 Then 'LUCV20160831, No considera a formato consumo sin convenio
        If nFormato_NEW <> nFormato Then
            If MsgBox("Monto de Exposición con Este Crédito: " & gcPEN_SIMBOLO & " " & Format(pnExpoEsteCred_NEW, gsFormatoNumeroView) & "." _
                    & Chr(13) & Chr(13) & "Con los datos ingresados la Evaluación de Crédito está variando de [Formato " & nFormato & "] a [Formato " & nFormato_NEW & "]" _
                    & Chr(13) & Chr(13) & "Si continúa con la operación el [Formato " & nFormato & "] se va a eliminar." _
                    & Chr(13) & Chr(13) & "                      ¿Está seguro de continuar?", vbInformation + vbYesNo, "Confirmación") = vbYes Then
                pbEliminarEvaluacion = True
                pnFormatoEliminado = nFormato 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                pnFormato_NEW = nFormato_NEW 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018

                If pnPrdEstado > 2000 Then
                    NecesitaFormatoEvaluacion = True

                    oEval.EliminaFormatoEvaluacion (psCtaCod)
                    MsgBox "Se ha eliminado la actual Evaluación del Crédito", vbInformation, "Aviso"

                    NecesitaFormatoEvaluacion = EvaluarCredito(psCtaCod, False, pnPrdEstado, pnProducto, pnSubProducto, pnExpoEsteCred_NEW, False)

                    If oEval.RecuperaFormatoEvaluacion(psCtaCod).RecordCount = 0 Then
                        NecesitaFormatoEvaluacion = True
                    End If
                Else
                    NecesitaFormatoEvaluacion = False
                End If
                Exit Function
            Else
                NecesitaFormatoEvaluacion = True
                Exit Function
            End If
        End If
'    End If
    NecesitaFormatoEvaluacion = False
    Exit Function
ErrNecesitaFormato:
    NecesitaFormatoEvaluacion = True
    MsgBox Err.Description, vbInformation, "Aviso"
End Function
Public Function EvaluarCredito(ByVal psCtaCod As String, _
                                Optional ByVal pbCargaDataNuevamente As Boolean = True, _
                                Optional ByVal pnEstado As Integer = 0, _
                                Optional ByVal pnProducto As Integer = 0, _
                                Optional ByVal pnSubProducto As Integer = 0, _
                                Optional ByVal pnMontoExpEsteCred As Double = 0#, _
                                Optional ByVal pbConsultar As Boolean = False, _
                                Optional ByVal pbImprimir As Boolean = False, _
                                Optional ByVal pbFormEmpr As Boolean = False, _
                                Optional ByVal pbTipoEliminaFormato As Boolean = False, _
                                Optional ByRef pbEliminaFormato As Boolean = False, _
                                Optional ByVal pbImprimirVB As Boolean = False) As Boolean
    Dim oEval As New COMDCredito.DCOMFormatosEval
    Dim oTipoCam  As New COMDConstSistema.NCOMTipoCambio
    Dim rs As ADODB.Recordset
    Dim nEstado As Integer
    Dim nFormato As Integer
    Dim retorno As Boolean
    Dim fnInicio As Integer
    'CTI3 ERS0032020****************
    Dim pnTpIngr  As Long
    Dim psTpoCredCodPadre As String
    Dim pnTpoCliente As Integer
    Dim nFormatoGrabado As Integer
    Dim nFormatoNuevo As Integer
    'END****************************
    'JOEP20190123 CP
    Dim rsMulPro As ADODB.Recordset
    Dim cProMult As String
    
    'JOEP20190123 CP
    retorno = False
    EvaluarCredito = False

 'JOEP20190123 CP
    Set rsMulPro = oEval.RecuperaProdMultFormato(pnSubProducto)
    If Not (rsMulPro.BOF And rsMulPro.EOF) Then
        cProMult = rsMulPro!SubProducto
    End If
 'JOEP20190123 CP

    If pbCargaDataNuevamente Then
        Set rs = oEval.RecuperaCreditoxEvaluacion(psCtaCod)
        pnEstado = rs!nPrdEstado
        pnSubProducto = CInt(rs!cTpoProdCod)
        pnProducto = CInt(Mid(rs!cTpoProdCod, 1, 1) & "00")
        pnMontoExpEsteCred = rs!nMontoExpCredito
        pnTpIngr = rs!nTpIngr
        psTpoCredCodPadre = IIf(IsNull(rs!cTpoCredCodPadre), "000", rs!cTpoCredCodPadre)
    Else
        Set rs = oEval.RecuperaCreditoxEvaluacion(psCtaCod)
        pnTpIngr = rs!nTpIngr
        psTpoCredCodPadre = IIf(IsNull(rs!cTpoCredCodPadre), "000", rs!cTpoCredCodPadre)
    End If
    
    
    Set rs = oEval.RecuperaFormatoEvaluacion(psCtaCod)
        
    

    If pnEstado = 2000 Or pnEstado = 2001 Then
        If rs.RecordCount = 0 Then
            fnInicio = 1
            'CTI3 ERS0032020******************************************************************************
            If Left(psTpoCredCodPadre, 1) = "4" Or Left(psTpoCredCodPadre, 1) = "5" Or Left(psTpoCredCodPadre, 1) = "7" Or Left(psTpoCredCodPadre, 1) = "8" Or Left(psTpoCredCodPadre, 1) = "0" Then
                pnTpoCliente = 1
            ElseIf Left(psTpoCredCodPadre, 1) = "1" Or Left(psTpoCredCodPadre, 1) = "2" Or Left(psTpoCredCodPadre, 1) = "3" Then
                pnTpoCliente = 2
            ElseIf Left(psTpoCredCodPadre, 1) = "0" Or Trim(psTpoCredCodPadre) = "" Then
                pnTpoCliente = 3
            End If
            'CTI3 ERS0032020*****************************************
            If pbTipoEliminaFormato = False Then
                Call oEval.ActualizarRSECuotaEstimada(psCtaCod)
            End If
            '********************************************************
            nFormato = oEval.AsignarFormato(pnProducto, pnSubProducto, pnMontoExpEsteCred, pnTpoCliente, pnTpIngr)
            nFormatoGrabado = 0
            'nFormato = oEval.AsignarFormato(pnProducto, pnSubProducto, pnMontoExpEsteCred)
            ' nFormato = 3 'Pruebas
            ' If cProMult <> "" And pbFormEmpr = False Then 'JOEP20190123 CP
            ''If pnProducto = "800" And pbFormEmpr = False Then'Comento JOEP20190123 CP
            ' If pnSubProducto = cProMult And pbFormEmpr = False Then 'JOEP20190123 CP
            'nFormato = 7
            ' End If
            ' End If
            'End******************************************************************************************
        Else
            fnInicio = 2
            nFormato = rs!nCodForm
            nFormatoGrabado = rs!nCodForm
            If pbImprimirVB Then
                Call oEval.ActualizarRSECuotaEstimada(psCtaCod)
            End If
        End If
    Else
        fnInicio = 3
        If rs.RecordCount > 0 Then
            nFormato = rs!nCodForm
        Else
            nFormato = 0
        End If
        nFormatoGrabado = nFormato
    End If
    
    'CTI3 ERS0032020******************************************************************************
    If pbTipoEliminaFormato = True Then
        If nFormatoGrabado = 0 Then
            pbEliminaFormato = False
        Else
            If Left(psTpoCredCodPadre, 1) = "4" Or Left(psTpoCredCodPadre, 1) = "5" Or Left(psTpoCredCodPadre, 1) = "7" Or Left(psTpoCredCodPadre, 1) = "8" Or Left(psTpoCredCodPadre, 1) = "0" Then
                pnTpoCliente = 1
            ElseIf Left(psTpoCredCodPadre, 1) = "1" Or Left(psTpoCredCodPadre, 1) = "2" Or Left(psTpoCredCodPadre, 1) = "3" Then
                pnTpoCliente = 2
            ElseIf Left(psTpoCredCodPadre, 1) = "0" Or Trim(psTpoCredCodPadre) = "" Then
                pnTpoCliente = 3
            End If
           
            nFormatoNuevo = oEval.AsignarFormato(pnProducto, pnSubProducto, pnMontoExpEsteCred, pnTpoCliente, pnTpIngr)
            
            If nFormatoGrabado = nFormatoNuevo Then
                pbEliminaFormato = False
            Else
                pbEliminaFormato = True
            End If
        End If
        EvaluarCredito = False
    End If
    'END**********************************************************************************************
    
    If pbTipoEliminaFormato = False Then
    If pbConsultar Then
        fnInicio = 3
    'CTI320200110 ERS003-2020. Agregó
    Else
        Call oEval.GrabarRSEAdmisionCuota(psCtaCod)
        Call oEval.GrabarRSERepAdmision(psCtaCod)
    'Fin CTI320200110 ERS003-2020
    End If

    Select Case nFormato
        Case 0: MsgBox "No existe Formato para este Credito.", vbInformation, "Aviso"
        Case 1: retorno = frmCredFormEvalFormato1.inicio(fnInicio, psCtaCod, nFormato, pnProducto, pnSubProducto, pnMontoExpEsteCred, pbImprimir, pnEstado, pbImprimirVB)
        Case 2: retorno = frmCredFormEvalFormato2.inicio(fnInicio, psCtaCod, nFormato, pnProducto, pnSubProducto, pnMontoExpEsteCred, pbImprimir, pnEstado, pbImprimirVB)
        Case 3: retorno = frmCredFormEvalFormato3.inicio(fnInicio, psCtaCod, nFormato, pnProducto, pnSubProducto, pnMontoExpEsteCred, pbImprimir, pnEstado, pbImprimirVB)
        Case 4: retorno = frmCredFormEvalFormato4.inicio(fnInicio, psCtaCod, nFormato, pnProducto, pnSubProducto, pnMontoExpEsteCred, pbImprimir, pnEstado, pbImprimirVB)
        Case 5: retorno = frmCredFormEvalFormato5.inicio(fnInicio, psCtaCod, nFormato, pnProducto, pnSubProducto, pnMontoExpEsteCred, pbImprimir, pnEstado, pbImprimirVB)
        Case 6: retorno = frmCredFormEvalFormato6.inicio(fnInicio, psCtaCod, nFormato, pnProducto, pnSubProducto, pnMontoExpEsteCred, pbImprimir, pbImprimirVB)
        Case 7: retorno = frmCredFormEvalFormatoSinConvenio.inicio(fnInicio, psCtaCod, nFormato, pnProducto, pnSubProducto, pnMontoExpEsteCred, pbImprimir, pnEstado, pbImprimirVB)
        Case 8: retorno = frmCredFormEvalFormatoConsumoConvenio.inicio(fnInicio, psCtaCod, nFormato, pnProducto, pnSubProducto, pnMontoExpEsteCred, pbImprimir, pnEstado, pbImprimirVB)
        Case 9: retorno = frmCredFormEvalFormatoParalelo.inicio(fnInicio, psCtaCod, nFormato, pnProducto, pnSubProducto, pnMontoExpEsteCred, pbImprimir, pnEstado, pbImprimirVB)
        Case Else: MsgBox "No existe Formato para este Credito.", vbInformation, "Aviso"
    End Select

    'La idea era que [retorno] devuelva si el formato seleccionado se ha grabado
    EvaluarCredito = retorno
    End If
End Function
'END EJVG *******
'LUCV20160720 **********************************************************
Public Function CumpleCriteriosRatios(ByVal psCtaCod As String) As Boolean
    Dim oConect As New COMConecta.DCOMConecta
    Dim rs As New ADODB.Recordset
    Dim sSql As String

    CumpleCriteriosRatios = False
    If oConect.AbreConexion = False Then Exit Function

    sSql = "stp_sel_ValidaCumpleCriterioRatiosFinancieros '" & psCtaCod & "'"

    Set rs = oConect.CargaRecordSet(sSql)
    If Not rs.EOF Then
        CumpleCriteriosRatios = IIf(rs!bCumple = 0, False, True)
    End If
    rs.Close
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
End Function
'CTI3 ERS0032020 **********************************************************
Public Function CumpleCriteriosRatioLimite(ByVal psCtaCod As String) As Boolean
    Dim oConect As New COMConecta.DCOMConecta
    Dim rs As New ADODB.Recordset
    Dim sSql As String

    CumpleCriteriosRatioLimite = False
    If oConect.AbreConexion = False Then Exit Function

    sSql = "stp_sel_validaRatioLimite '" & psCtaCod & "'"

    Set rs = oConect.CargaRecordSet(sSql)
    If Not rs.EOF Then
        CumpleCriteriosRatioLimite = IIf(rs!bCumple = 0, False, True)
    End If
    rs.Close
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
End Function
'MARG20170606 ERS007-2017***
Public Function CumpleCriteriosAprobacionCredito(ByVal psCtaCod As String, ByVal psCodUser As String) As Boolean
    Dim oConect As New COMConecta.DCOMConecta
    Dim rs As New ADODB.Recordset
    Dim sSql As String

    CumpleCriteriosAprobacionCredito = False
    If oConect.AbreConexion = False Then Exit Function

    sSql = "stp_sel_ValidaCumpleCriteriosAprobacionCredito '" & psCtaCod & "','" & psCodUser & "'"

    Set rs = oConect.CargaRecordSet(sSql)
    If Not rs.EOF Then
        CumpleCriteriosAprobacionCredito = IIf(rs!bCumple = 0, False, True)
    End If
    rs.Close
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
End Function
'END MARG*********************
Public Function CargaMatrizDatosMantenimientoCtaCobrar(ByRef pvDetalleRef() As tFormEvalDetalleActivosCtasCobrarFormato5, ByVal psCtaCod As String, ByVal pnConsCod As Integer, _
                            ByVal pnConsValor As Integer, ByVal pnTpoPat As Integer) As Integer

    Dim oCredEval As New COMNCredito.NCOMFormatosEval
    Dim oRS As ADODB.Recordset
    Dim i As Integer

    Set oRS = oCredEval.RecuperaDatosCtaCobrar(psCtaCod, pnConsCod, pnConsValor, pnTpoPat, 2)
    If Not (oRS.EOF And oRS.BOF) Then
        CargaMatrizDatosMantenimientoCtaCobrar = oRS.RecordCount
        ReDim Preserve pvDetalleRef(CargaMatrizDatosMantenimientoCtaCobrar)
        For i = 1 To CargaMatrizDatosMantenimientoCtaCobrar
            pvDetalleRef(i).dFecha = Format(oRS!dCtaFecha, "DD/MM/YYYY")
            pvDetalleRef(i).cCtaporCobrar = oRS!CDescripcion
            pvDetalleRef(i).nTotal = Format(oRS!nTotal, "#,##0.00")
            oRS.MoveNext
        Next
    Else
        ReDim Preserve pvDetalleRef(0)
    End If
End Function
Public Function CargarMatrizDatosMantenimientoInvetario(ByRef pvDetalleActivoFlex() As tFormEvalDetalleActivosInventarioFormato5, ByVal psCtaCod As String, _
                                                        ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnTipoPat As Integer) As Integer
    Dim oCred As New COMNCredito.NCOMFormatosEval
    Dim rsDatos As ADODB.Recordset
    Dim i As Integer

    Set rsDatos = oCred.ObtieneDetalleInventario(psCtaCod, pnConsCod, pnConsValor, pnTipoPat)

    If Not (rsDatos.EOF And rsDatos.BOF) Then
        CargarMatrizDatosMantenimientoInvetario = rsDatos.RecordCount

        ReDim Preserve pvDetalleActivoFlex(CargarMatrizDatosMantenimientoInvetario)
        For i = 1 To CargarMatrizDatosMantenimientoInvetario
            pvDetalleActivoFlex(i).cMercaderia = rsDatos!Mercaderia
            pvDetalleActivoFlex(i).nCantidad = rsDatos!cantidad
            pvDetalleActivoFlex(i).cUnidMed = Trim(Right(rsDatos!UnidadMedida, 3))
            pvDetalleActivoFlex(i).nCostoUnit = rsDatos!CostoUnitario
            pvDetalleActivoFlex(i).nTotal = rsDatos!Total
        Next i
    Else
        ReDim Preserve pvDetalleActivoFlex(0)
    End If
End Function
Public Function CargarMatrizDatosMantenimientoActivoFijo(ByRef pvDetalleActivoFlex() As tFormEvalDetalleActivosActivoFijoFormato5, ByVal psCtaCod As String, _
                                                    ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnTipoPat As Integer) As Integer
    Dim oCred As New COMNCredito.NCOMFormatosEval
    Dim rsDatos As ADODB.Recordset
    Dim nIndice As Integer

    Set rsDatos = oCred.ObtieneDetalleActiFijo(psCtaCod, pnConsCod, pnConsValor, pnTipoPat)

    If Not (rsDatos.EOF And rsDatos.BOF) Then
        CargarMatrizDatosMantenimientoActivoFijo = rsDatos.RecordCount

        ReDim Preserve pvDetalleActivoFlex(CargarMatrizDatosMantenimientoActivoFijo)
        For nIndice = 1 To CargarMatrizDatosMantenimientoActivoFijo
            pvDetalleActivoFlex(nIndice).CDescripcion = rsDatos!Descripcion
            pvDetalleActivoFlex(nIndice).nCantidad = rsDatos!cantidad
            pvDetalleActivoFlex(nIndice).nPrecio = rsDatos!CostoUnitario
            pvDetalleActivoFlex(nIndice).nTotal = rsDatos!Total
        Next nIndice
    Else
        ReDim Preserve pvDetalleActivoFlex(0)
    End If
End Function

'LUCV FIN **************************************************************
'FRHU 20160802 ERS0022016
Public Function ValidarExisteNivelAprobacionParaAutorizacion(ByVal psCtaCod As String) As Boolean
    Dim objCredito As New COMDCredito.DCOMNivelAprobacion
    Dim rs As New ADODB.Recordset
    Dim lsValor As String
    Dim lsExoneraDesc As String, lsmensaje As String
    Dim lnMontoExposicion As Double

    ValidarExisteNivelAprobacionParaAutorizacion = True
    Set rs = objCredito.ValidarExisteNivelDeAprobacion(psCtaCod)
    If Not (rs.BOF And rs.EOF) Then
        lsValor = rs!valor
        lsExoneraDesc = rs!cExoneraDesc
        lsmensaje = rs!cMensaje
        lnMontoExposicion = rs!nMontoExposicion
        If lsValor = "1" Then
            ValidarExisteNivelAprobacionParaAutorizacion = False
            MsgBox lsmensaje & " " & Format(lnMontoExposicion, "#,###,##0.00") & vbNewLine & _
                   "Autorización: " & vbNewLine & _
                   "     " & lsExoneraDesc, vbInformation, "Aviso"
            Exit Function
        End If
    End If
End Function
'FIN FRHU ERS002-2016

'*** PEAC 20160811
Public Function DevolverNombreMes(ByVal pnMes As Integer) As String
    DevolverNombreMes = Choose(pnMes, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
End Function

'*** FRHU 20160823
Public Function CreditoTieneFormatoEvaluacion(ByVal psCtaCod As String) As Boolean
    Dim objCredito As New COMDCredito.DCOMFormatosEval
    Dim rs As New ADODB.Recordset

    Set rs = objCredito.CreditoTieneFormatoEvaluacion(psCtaCod)
    If Not (rs.BOF And rs.EOF) Then
        CreditoTieneFormatoEvaluacion = True
    Else
        CreditoTieneFormatoEvaluacion = False
    End If
    Set rs = Nothing
    Set objCredito = Nothing
End Function
'*** FIN FRHU 20160823

'LUCV20161213->Según ERS068-2016
Public Function ValidaIfiExisteCompraDeuda(ByVal pcCtaCod As String, ByVal MatIfiGastoFami As Variant, Optional ByVal MatIfiGastoNego As Variant, Optional ByRef psMensajeIfi As String, Optional ByVal MatIfiGastoFamiNoSupervisada As Variant, Optional ByVal MatIfiGastoNegoNoSupervisada As Variant) As Boolean
    Dim i As Integer
    Dim J As Integer
    Dim k As Integer
    Dim rsListaCompraDeuda As New ADODB.Recordset
    Dim oDCOMInstFinac As New COMDpersona.DCOMInstFinac
    Set rsListaCompraDeuda = oDCOMInstFinac.ObtieneCtaIFIxCompraDeuda(pcCtaCod)
    Dim lsCodIfi As String
    Dim lsCodIfiMsj As String
    
    Dim nCantidadFamiliar As Integer
    Dim nCantidadFamiliarNoSupervisada As Integer
    Dim nCantidadNegocio As Integer
    Dim nCantidadNegocioNoSujpervisada As Integer
    Dim nTotalIfis As Integer
    If IsArray(MatIfiGastoFami) Then
        nCantidadFamiliar = UBound(MatIfiGastoFami)
    End If
    If IsArray(MatIfiGastoNego) Then
        nCantidadNegocio = UBound(MatIfiGastoNego)
    End If
    If IsArray(MatIfiGastoFamiNoSupervisada) Then
        nCantidadFamiliarNoSupervisada = UBound(MatIfiGastoFamiNoSupervisada)
    End If
    If IsArray(MatIfiGastoNegoNoSupervisada) Then
        nCantidadNegocioNoSujpervisada = UBound(MatIfiGastoNegoNoSupervisada)
    End If
    nTotalIfis = nCantidadFamiliar + nCantidadNegocio + nCantidadFamiliarNoSupervisada + nCantidadNegocioNoSujpervisada
    Dim MatIfiGeneral As Variant
    Dim nContador As Integer
    ReDim MatIfiGeneral(nTotalIfis, 2)
    nContador = 0
    If IsArray(MatIfiGastoFami) Then
        For J = 0 To nCantidadFamiliar - 1
            MatIfiGeneral(nContador, 1) = MatIfiGastoFami(J, 1)
            MatIfiGeneral(nContador, 2) = MatIfiGastoFami(J, 2)
            nContador = nContador + 1
          Next J
    End If
    
    If IsArray(MatIfiGastoNego) Then
        For J = 0 To nCantidadNegocio - 1
            MatIfiGeneral(nContador, 1) = MatIfiGastoNego(J, 1)
            MatIfiGeneral(nContador, 2) = MatIfiGastoNego(J, 2)
            nContador = nContador + 1
          Next J
    End If
    
    If IsArray(MatIfiGastoFamiNoSupervisada) Then
        For J = 0 To nCantidadFamiliarNoSupervisada - 1
            MatIfiGeneral(nContador, 1) = MatIfiGastoFamiNoSupervisada(J, 1)
            MatIfiGeneral(nContador, 2) = MatIfiGastoFamiNoSupervisada(J, 2)
            nContador = nContador + 1
          Next J
    End If
     If IsArray(MatIfiGastoNegoNoSupervisada) Then
        For J = 0 To nCantidadNegocioNoSujpervisada - 1
            MatIfiGeneral(nContador, 1) = MatIfiGastoNegoNoSupervisada(J, 1)
            MatIfiGeneral(nContador, 2) = MatIfiGastoNegoNoSupervisada(J, 2)
            nContador = nContador + 1
          Next J
    End If
    
    ValidaIfiExisteCompraDeuda = True
    lsCodIfiMsj = ""

    Do While Not rsListaCompraDeuda.EOF
    ValidaIfiExisteCompraDeuda = False
        lsCodIfi = ""
            'Valida- Gastos Familiares(Compra Deuda)
            If IsArray(MatIfiGeneral) Then
                For J = 0 To UBound(MatIfiGeneral)
                    If rsListaCompraDeuda!cPersCodIfi = MatIfiGeneral(J, 1) And CDbl(rsListaCompraDeuda!nMontoCuoPaga) = CDbl(MatIfiGeneral(J, 2)) Then
                        ValidaIfiExisteCompraDeuda = True
                        lsCodIfi = ""
                        Exit For
                    End If
                        ValidaIfiExisteCompraDeuda = False
                        lsCodIfi = Trim(rsListaCompraDeuda!cPersNombre) & " = " & Format(rsListaCompraDeuda!nMontoCuoPaga, "#,###,##0.00") & " Soles"
                Next J
            Else
                        lsCodIfi = Trim(rsListaCompraDeuda!cPersNombre) & " = " & Format(rsListaCompraDeuda!nMontoCuoPaga, "#,###,##0.00") & " Soles"
            End If

           
            If Len(Trim(lsCodIfi)) > 0 Then
                lsCodIfiMsj = " - " & lsCodIfi & Chr(10) & lsCodIfiMsj
            End If
            rsListaCompraDeuda.MoveNext
    Loop
'    Do While Not rsListaCompraDeuda.EOF
'    ValidaIfiExisteCompraDeuda = False
'        lsCodIfi = ""
'            'Valida- Gastos Familiares(Compra Deuda)
'            If IsArray(MatIfiGastoFami) Then
'                For j = 0 To UBound(MatIfiGastoFami)
'                    If rsListaCompraDeuda!cPersCodIfi = MatIfiGastoFami(j, 1) And CDbl(rsListaCompraDeuda!nMontoCuoPaga) = CDbl(MatIfiGastoFami(j, 2)) Then
'                        ValidaIfiExisteCompraDeuda = True
'                        lsCodIfi = ""
'                        Exit For
'                    End If
'                        ValidaIfiExisteCompraDeuda = False
'                        lsCodIfi = Trim(rsListaCompraDeuda!cPersNombre) & " = " & Format(rsListaCompraDeuda!nMontoCuoPaga, "#,###,##0.00") & " Soles"
'                Next j
'            Else
'                        lsCodIfi = Trim(rsListaCompraDeuda!cPersNombre) & " = " & Format(rsListaCompraDeuda!nMontoCuoPaga, "#,###,##0.00") & " Soles"
'            End If
'
'            'Valida - Gastos Negocio (Compra Deuda)
'            If Not ValidaIfiExisteCompraDeuda Then
'                If IsArray(MatIfiGastoNego) Then
'                    For K = 0 To UBound(MatIfiGastoNego)
'                        If rsListaCompraDeuda!cPersCodIfi = MatIfiGastoNego(K, 1) And CDbl(rsListaCompraDeuda!nMontoCuoPaga) = CDbl(MatIfiGastoNego(K, 2)) Then
'                            ValidaIfiExisteCompraDeuda = True
'                            lsCodIfi = ""
'                            Exit For
'                        End If
'                            ValidaIfiExisteCompraDeuda = False
'                            lsCodIfi = Trim(rsListaCompraDeuda!cPersNombre) & " = " & Format(rsListaCompraDeuda!nMontoCuoPaga, "#,###,##0.00") & " Soles"
'                    Next K
'                Else
'                            lsCodIfi = Trim(rsListaCompraDeuda!cPersNombre) & " = " & Format(rsListaCompraDeuda!nMontoCuoPaga, "#,###,##0.00") & " Soles"
'                End If
'
'            End If
'            If Len(Trim(lsCodIfi)) > 0 Then
'                lsCodIfiMsj = " - " & lsCodIfi & Chr(10) & lsCodIfiMsj
'            End If
'            rsListaCompraDeuda.MoveNext
'    Loop

    'If Len(lsCodIfiMsj) > 0 Then 'LUCV20170410, Comentó
        'lsCodIfiMsj = Mid(Trim(lsCodIfiMsj), 1, Len(Trim(lsCodIfiMsj)) - 2)
    'End If

    psMensajeIfi = lsCodIfiMsj
    rsListaCompraDeuda.Close
    Set rsListaCompraDeuda = Nothing
End Function
'FRHU 20170914 ERS049-2017
Public Function ValidarTasaMaxima(ByVal psCtaCod As String, ByVal pnTasaInt As Double, _
                                  Optional ByVal pnIdCampana As Integer = -1, Optional ByVal psTpoProdCod As String = "", Optional ByVal pnMonto As Currency = 0) As Boolean
    Dim oCredito As New COMDCredito.DCOMCredito
    Dim oCredActBD As New COMDCredito.DCOMCredActBD
    Dim oLineas As New COMDCredito.DCOMLineaCredito
    Dim RLinea As New ADODB.Recordset, RCredito As New ADODB.Recordset
    Dim lnIdCampana As Integer, lnTasaInicial As Currency, lnTasaFinal As Currency
    Dim lsTpoProdCod As String

    If pnIdCampana = -1 Then
        Set RCredito = oCredito.RecuperaColocacCredCampos(psCtaCod, "IdCampana")
        lnIdCampana = RCredito!idCampana
    Else
        lnIdCampana = pnIdCampana
    End If

    If psTpoProdCod = "" Then
        Set RCredito = oCredActBD.RecuperaColocaciones(psCtaCod)
        lsTpoProdCod = RCredito!cTpoProdCod
    Else
        lsTpoProdCod = psTpoProdCod
    End If

    Set RLinea = oLineas.RecuperaLineadeCreditoProductoCrediticio(lsTpoProdCod, lnIdCampana, "", "", "", Mid(psCtaCod, 9, 1), pnMonto, 0)
    If RLinea.RecordCount > 0 Then
        lnTasaInicial = RLinea!nTasaIni
        lnTasaFinal = RLinea!nTasaFin
    Else
        lnTasaInicial = 0
        lnTasaFinal = 0
    End If

    If pnTasaInt <> lnTasaFinal Then
        If MsgBox("El nuevo valor de la tasa """ & Format(pnTasaInt, "#,##0.00") & """ difiere de la tasa máxima """ & Format(lnTasaFinal, "#,##0.00") & """" & _
                  Chr(13) & "Tasa Mínima  : " & Format(lnTasaInicial, "#,##0.00") & _
                  Chr(13) & "Tasa Máxima : " & Format(lnTasaFinal, "#,##0.00") & _
                  Chr(13) & _
                  Chr(13) & "Desea Continuar ? ", vbInformation + vbYesNo, "Aviso") = vbYes Then
            ValidarTasaMaxima = True
        Else
            ValidarTasaMaxima = False
        End If
    Else
        ValidarTasaMaxima = True
    End If

    Set oCredito = Nothing
    Set oCredActBD = Nothing
    Set oLineas = Nothing
    Set RCredito = Nothing
    Set RLinea = Nothing
End Function
'FIN FRHU 20170914
'FRHU 20171014: Se agrego de gFunCartaAfectacion
Public Sub ImprimeCartaAfectacion(ByVal cCtaCod As String, ByVal nTipoCredito As Integer, ByVal nMontoCol As Currency, Optional ByVal nPoliza As Long)
    'ByVal cPersCod As String, ByVal cPersGarantia As String, ByVal cDoi As String, ByVal nDoi As String, Optional ByVal cEstadoCivil As String, Optional ByVal cDomicilio As String,

    Dim cPersCod As String
    Dim cPersGarantia As String
    Dim cDOI As String
    Dim nDoi As String
    Dim cEstadoCivil As String
    Dim cDomicilio As String
    Dim PersonaGarantia As ADODB.Recordset
    Dim PersonaTitular As ADODB.Recordset 'EAAS 20170831
    Dim obj As COMNCredito.NCOMCredito
    Set obj = New COMNCredito.NCOMCredito
    Set PersonaGarantia = obj.ObtenerGarantiaPersona(cCtaCod)
    Set PersonaTitular = obj.ObtenerTitularCredito(cCtaCod) 'EAAS 20170831
    Set obj = Nothing

    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000079", nTipoCredito) Or PersonaGarantia.RecordCount = 1 Then
    'If nTipoCredito = 504 Or PersonaGarantia.RecordCount = 1 Then
    '**ARLO20180712 ERS042 - 2018
    Do While Not PersonaGarantia.EOF
        cPersCod = PersonaGarantia!cPersCod
        cPersGarantia = PersonaGarantia!cPersNombre
        cDOI = PersonaGarantia!cDOI
        nDoi = PersonaGarantia!nDoi
        '**ARLO20180712 ERS042 - 2018
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("N0000080", nTipoCredito) Then
        'If nTipoCredito = 703 Then
        '**ARLO20180712 ERS042 - 2018
            cEstadoCivil = PersonaGarantia!cEstadoCivil
            cDomicilio = PersonaGarantia!cDireccion
        End If

    PersonaGarantia.MoveNext
    Loop
    Else 'EAAS 20170831
    cPersCod = PersonaGarantia!cPersCod
    End If
'cf

'    Set obj = New COMNCredito.NCOMCredito
'    Set PersonaGarantia = obj.ObtenerGarantiaPersona(vCodCta)
'    Set obj = Nothing
'
'    Do While Not PersonaGarantia.EOF
'    cPersCod = PersonaGarantia!cPersCod
'    cPersGarantia = PersonaGarantia!cPersNombre
'    cDoi = PersonaGarantia!cDoi
'    nDoi = PersonaGarantia!nDoi
'    PersonaGarantia.MoveNext
'
'    Loop
    Dim oDCOMCartaFianza As COMNCredito.NCOMCredito
    Dim rsGarantias As ADODB.Recordset

    On Error GoTo ErrorImprimirPDF
    Dim sParrafo1 As String
    Dim sParrafo1a As String
    Dim sParrafo1b As String
    Dim sParrafo1c As String
    Dim sParrafo2 As String
    Dim sParrafo3 As String
    Dim sParrafo4 As String
    Dim sParrafo5a As String
    Dim sParrafo5b As String
    Dim sParrafo5c As String
    Dim oDoc  As cPDF
    Set oDoc = New cPDF

    Dim sMontoColocado As String
    Dim sMontoCol As String
    sMontoCol = Format(nMontoCol, "#,###0.00")
    sMontoColocado = IIf(Mid(cCtaCod, 9, 1) = "1", "S/ ", "$. ") & sMontoCol & " " & "(" & UCase(NumLet(sMontoCol)) & IIf(Mid(cCtaCod, 9, 1) = "2", "", " Y " & IIf(InStr(1, sMontoCol, ".") = 0, "00", Mid(sMontoCol, InStr(1, sMontoCol, ".") + 1, 2)) & "/100 ") & IIf(Mid(cCtaCod, 9, 1) = "1", " SOLES)", " DOLARES)")
    sMontoColocado = IIf(Mid(cCtaCod, 9, 1) = "1", "S/ ", "$. ") & Format(sMontoCol, "#,###0.00") & " " & "(" & UCase(NumLet(sMontoCol)) & IIf(Mid(cCtaCod, 9, 1) = "2", "", " Y " & IIf(InStr(1, Str(Format(sMontoCol, "#,###0.00")), ".") = 0, "00", Mid(Str(Format(sMontoCol, "#,###0.00")), InStr(1, Str(Format(sMontoCol, "#,###0.00")), ".") + 1, 2)) & "/100 ") & IIf(Mid(cCtaCod, 9, 1) = "1", "SOLES)", " US DOLARES)")
    sMontoColocado = IIf(Mid(cCtaCod, 9, 1) = "1", "S/ ", "$. ") & Format(sMontoCol, "#,###0.00") & " " & "(" & UCase(NumLet(sMontoCol)) & IIf(Mid(cCtaCod, 9, 1) = "2", "", " Y " & IIf(InStr(1, sMontoCol, ".") = 0, "00", Mid(sMontoCol, InStr(1, sMontoCol, ".") + 1, 2)) & "/100 ") & IIf(Mid(cCtaCod, 9, 1) = "1", "SOLES)", " US DOLARES)")

    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000081", nTipoCredito) Then
    'If nTipoCredito = 514 Then
    '**ARLO20180712 ERS042 - 2018
        Dim rsCartaFianza As ADODB.Recordset
        Set rsCartaFianza = New ADODB.Recordset
        Set oDCOMCartaFianza = New COMNCredito.NCOMCredito
        Set rsCartaFianza = oDCOMCartaFianza.ObtenerCFRelacion(cCtaCod)
        Set oDCOMCartaFianza = Nothing
        Dim nTpoRelac As Integer
        Dim cAcreedor As String
        Dim cAval As String
        Dim cDoiAval As String
        Dim nDoiAval As String

        Do While Not rsCartaFianza.EOF
        If rsCartaFianza!nAval = 1 Then
            If rsCartaFianza!NPRDPERSRELAC = 38 Then
                cAval = rsCartaFianza!cPersNombre
                cDoiAval = rsCartaFianza!cDOI
                nDoiAval = rsCartaFianza!nDoi
            Else
                cAcreedor = rsCartaFianza!cPersNombre
            End If
        Else 'CONSORCIOS
            If rsCartaFianza!NPRDPERSRELAC = 20 Then
                cAval = rsCartaFianza!cPersNombre
                cDoiAval = rsCartaFianza!cDOI
                nDoiAval = rsCartaFianza!nDoi
            Else
                cAcreedor = rsCartaFianza!cPersNombre
            End If
        End If
        rsCartaFianza.MoveNext
        Loop

          sParrafo2 = "Incluyendo sus intereses compensatorios, frutos y demás bienes que produzca, con la finalidad de honrar y cumplir " & _
                      "las obligaciones contraídas, por " & cAval & " con " & cDoiAval & "° " & nDoiAval & ", con su representada por la " & _
                      "CARTA FIANZA N° " & Format(nPoliza, "0000000") & " (crédito N° " & cCtaCod & ") por la suma de " & sMontoColocado & " " & _
                      ", a favor  de " & cAcreedor & " en caso de incumplimiento de la deuda; de conformidad a lo dispuesto en el Art. 132 " & _
                      "inc. 11 de la Ley General del Sistema Financiero y del Sistema de Seguros y Orgánica de la Superintendencia de " & _
                      "Banca y Seguros - Ley N° 26702, que permite a las entidades del Sector Financiero la compensación de sus acreencias " & _
                      "por cobrar con los activos de los deudores."

                If Not oDoc.PDFCreate(App.Path & "\Spooler\CartaAfectacion_" & nPoliza & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
                    Exit Sub
                End If

    ElseIf PersonaGarantia.RecordCount = 1 And cPersCod = PersonaTitular!cPersCod Then 'EAAS 20170831
          sParrafo2 = "Incluyendo sus respectivos intereses y demás bienes que produzca, AFECTEN y RETIREN el importe de " & _
                    "las obligaciones pendientes de pago que mantenga el crédito N° " & cCtaCod & " aprobado y/o otorgado " & _
                    "a mi persona;     afectación que se hace, de conformidad a lo dispuesto en el Art. 132 inc. 11 de la " & _
                    "Ley General del Sistema Financiero y del Sistema de Seguros y Orgánica de la Superintendencia de " & _
                    "Banca y Seguros - Ley N° 26702, que permite a las entidades del Sector Financiero, realizar la " & _
                    "compensación de sus acreencias por cobrar con los activos de los deudores, hasta por el monto de " & _
                    "aquellas. Así mismo, manifiesto que la presente autorización es irrevocable y se mantendrá hasta " & _
                    "que haya cumplido con cancelar todas las obligaciones pendientes de pago con Caja Maynas."

                    If Not oDoc.PDFCreate(App.Path & "\Spooler\CartaAfectacion_" & cCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
                    Exit Sub
                    End If
    Else 'EAAS 20170831
    sParrafo2 = "Incluyendo sus respectivos intereses y demás bienes que produzca, AFECTEN y RETIREN el importe de " & _
                    "las obligaciones pendientes de pago que mantenga el crédito N° " & cCtaCod & " aprobado y/o otorgado " & _
                    "a favor del cliente " & PersonaTitular!cPersNombre & " con DNI." & PersonaTitular!cPersIDNro & " en el cual intervengo/intervenimos en calidad de garante(s) y fiador(es) solidario(s) ;     afectación que se hace, de conformidad a lo dispuesto en el Art. 132 inc. 11 de la " & _
                    "Ley General del Sistema Financiero y del Sistema de Seguros y Orgánica de la Superintendencia de " & _
                    "Banca y Seguros - Ley N° 26702, que permite a las entidades del Sector Financiero, realizar la " & _
                    "compensación de sus acreencias por cobrar con los activos de los deudores, hasta por el monto de " & _
                    "aquellas. Así mismo, manifiesto que la presente autorización es irrevocable y se mantendrá hasta " & _
                    "que haya cumplido con cancelar todas las obligaciones pendientes de pago con Caja Maynas."
         'END EAAS 20170831

'            sParrafo2a = "Incluyendo sus respectivos intereses y demás bienes que produzca, AFECTEN y RETIREN el importe de "
'            sParrafo2b = "las obligaciones pendientes de pago que mantenga el crédito N° 109011011000110124 aprobado y/o otorgado a "
'            sParrafo2c = "mi persona; afectación que se hace, de conformidad a lo dispuesto en el Art. 132 inc. 11 de la Ley"
'            sParrafo2d = "General del Sistema Financiero y del Sistema de Seguros y Orgánica de la Superintendencia de Banca "
'            sParrafo2e = "y Seguros - Ley N° 26702, que permite a las entidades del Sector Financiero, realizar la compensación "
'            sParrafo2f = "de sus acreencias por cobrar con los activos de los deudores, hasta por el monto de aquellas."
'            sParrafo2g = "Así mismo, manifiesto que la presente autorización es irrevocable y se mantendrá hasta que haya"
'            sParrafo2h = "cumplido con cancelar todas las obligaciones pendientes de pago con Caja Maynas."
'            sParrafo2i = ""

          If Not oDoc.PDFCreate(App.Path & "\Spooler\CartaAfectacion_" & cCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
                    Exit Sub
            End If
    End If

    Dim nTamano As Integer
    Dim nValidar As Double
    Dim nTop As Integer
    Dim sFechaActual As String

    Dim cNroDoc As String
    Dim nSaldo As String
    Dim sMontoGravado As String
    Dim nNro As Integer
    nNro = 0

    Set oDCOMCartaFianza = New COMNCredito.NCOMCredito
    Set rsGarantias = oDCOMCartaFianza.ObtenerGarantias(cCtaCod)
    Set oDCOMCartaFianza = Nothing

    'Creacion de Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Carta Afectacion Nº " & cCtaCod
    oDoc.Title = "Carta Afectacion Nº " & cCtaCod


    oDoc.Fonts.Add "F1", "Times New Roman", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Times New Roman", TrueType, Bold, WinAnsiEncoding

    oDoc.NewPage A4_Vertical

    sFechaActual = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")

    oDoc.WTextBox 70, 70, 10, 450, "CARTA DE AUTORIZACIÓN DE AFECTACIÓN", "F1", 13, vMiddle
    oDoc.WTextBox 95, 50, 10, 450, "Iquitos, " & sFechaActual, "F2", 11, hLeft
    oDoc.WTextBox 120, 50, 10, 450, "Señores:", "F1", 11, hLeft
    oDoc.WTextBox 140, 50, 10, 450, "CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS", "F2", 11, hLeft '
    oDoc.WTextBox 160, 50, 10, 450, "Presente.-", "F1", 11, hLeft
    oDoc.WTextBox 200, 50, 10, 450, "De mi consideración:", "F1", 11, hLeft

    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000082", nTipoCredito) Then
    'If nTipoCredito = 514 Then
    '**ARLO20180712 ERS042 - 2018
            sParrafo1 = "Por medio de la presente me dirijo a ustedes para saludarlos y principalmente para AUTORIZAR a la " & _
            "CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS S.A., a que pueda constituir como garantía mobiliaria, " & _
            "de ser el caso, así como AFECTAR Y RETIRAR DE MI(s) CUENTA(s) DE DEPOSITO A PLAZO FIJO:"
    Else
            If PersonaGarantia.RecordCount = 1 Then
                sParrafo1 = "Por el presente documento Yo,   " & cPersGarantia & "  , identificado con   " & cDOI & "   N°   " & nDoi & "  , estado civil   " & cEstadoCivil & "   y   " & _
                    "domiciliado en   " & cDomicilio & "  ; me dirijo a ustedes para saludarlos y manifestar expresamente mi voluntad de AUTORIZARLOS para " & _
                    "que de mi(s) CUENTA(s) DE DEPÓSITO A PLAZO FIJO:"
             Else
                    sParrafo1 = "Por el presente documento "

                    Dim nTotal As Integer
                     Set obj = New COMNCredito.NCOMCredito
                    Set PersonaGarantia = obj.ObtenerGarantiaPersona(cCtaCod)
                    Set obj = Nothing
                    nTotal = PersonaGarantia.RecordCount

                     Do While Not PersonaGarantia.EOF
                        cPersCod = PersonaGarantia!cPersCod
                        cPersGarantia = PersonaGarantia!cPersNombre
                        cDOI = PersonaGarantia!cDOI
                        nDoi = PersonaGarantia!nDoi
                        cEstadoCivil = PersonaGarantia!cEstadoCivil
                        cDomicilio = PersonaGarantia!cDireccion

                        If nTotal = PersonaGarantia.RecordCount Then
                              sParrafo1a = "Yo,   " & cPersGarantia & "  , identificado con   " & cDOI & "   N°   " & nDoi & "  , estado civil   " & cEstadoCivil & "   y   " & _
                            "domiciliado en   " & cDomicilio & "  "
                        ElseIf nTotal = 1 Then
                            sParrafo1a = "y Yo,   " & cPersGarantia & "  , identificado con   " & cDOI & "   N°   " & nDoi & "  , estado civil   " & cEstadoCivil & "   y   " & _
                            "domiciliado en   " & cDomicilio & "  "
                        Else
                            sParrafo1a = ", Yo,   " & cPersGarantia & "  , identificado con   " & cDOI & "   N°   " & nDoi & "  , estado civil   " & cEstadoCivil & "   y   " & _
                            "domiciliado en   " & cDomicilio & "  "
                        End If

                        sParrafo1 = sParrafo1 + sParrafo1a
                        nTotal = nTotal - 1
                    PersonaGarantia.MoveNext
                    Loop

                    sParrafo1b = "; nos dirigimos a ustedes para saludarlos y manifestar expresamente nuestra voluntad de AUTORIZARLOS para "

                    If PersonaGarantia.RecordCount = 1 Then
                    sParrafo1c = "que de mi(s) CUENTA(s) DE DEPÓSITO A PLAZO FIJO:"
                    Else
                    sParrafo1c = "que de nuestra(s) CUENTA(S) DE DEPÓSITO A PLAZO FIJO:"
                    End If

                    sParrafo1 = sParrafo1 + sParrafo1b + sParrafo1c
             End If

    End If

    'String(20, "-") & " " &
    'oDoc.WTextBox 220, 50, 50, 580, sParrafo1, "F1", 11, hjustify, , , , , , 50

    nTamano = Len(sParrafo1)
    nValidar = nTamano / 75
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    nTop = 180
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo1, "F1", 12, hjustify
    oDoc.WTextBox nTop, 0, nTamano * 8, 580, sParrafo1, "F1", 10, hjustify, , , , , , 50 'esto
    'oDoc.WText nTop, 0, sParrafo1, "F1", 11
    'oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True

    nTop = nTop + (nTamano * 8) + 10

     Dim counter As Integer
    counter = nTop
    Dim sSaldo As String

            Do While Not rsGarantias.EOF
                cNroDoc = rsGarantias!cNroDoc
                nSaldo = rsGarantias!nSaldo
                nNro = nNro + 1

                sSaldo = Format(nSaldo, "#,###0.00")
                sMontoGravado = IIf(Mid(cNroDoc, 9, 1) = "1", "S/ ", "$. ") & sSaldo & " " & "(" & UCase(NumLet(sSaldo)) & IIf(Mid(cNroDoc, 9, 1) = "2", "", " Y " & IIf(InStr(1, sSaldo, ".") = 0, "00", Mid(sSaldo, InStr(1, sSaldo, ".") + 1, 2)) & "/100 ") & IIf(Mid(cNroDoc, 9, 1) = "1", "SOLES)", " DOLARES)")
                sMontoGravado = IIf(Mid(cNroDoc, 9, 1) = "1", "S/ ", "$. ") & Format(sSaldo, "#,###0.00") & " " & "(" & UCase(NumLet(sSaldo)) & IIf(Mid(cNroDoc, 9, 1) = "2", "", " Y " & IIf(InStr(1, Str(Format(sSaldo, "#,###0.00")), ".") = 0, "00", Mid(Str(Format(sSaldo, "#,###0.00")), InStr(1, Str(Format(sSaldo, "#,###0.00")), ".") + 1, 2)) & "/100 ") & IIf(Mid(cNroDoc, 9, 1) = "1", "SOLES)", " US DOLARES)")
                sMontoGravado = IIf(Mid(cNroDoc, 9, 1) = "1", "S/ ", "$. ") & Format(sSaldo, "#,###0.00") & " " & "(" & UCase(NumLet(sSaldo)) & IIf(Mid(cNroDoc, 9, 1) = "2", "", " Y " & IIf(InStr(1, sSaldo, ".") = 0, "00", Mid(sSaldo, InStr(1, sSaldo, ".") + 1, 2)) & "/100 ") & IIf(Mid(cNroDoc, 9, 1) = "1", "SOLES)", " US DOLARES)")


                oDoc.WTextBox counter, 0, 10, 580, nNro & ") " & cNroDoc & " por un monto de " & sMontoGravado & ".", "F1", 10, hjustify, , , , , , 50
                '485
                counter = counter + 20
                rsGarantias.MoveNext

            Loop

    'PARRAFO 2

    nTop = counter + 10
    nTamano = Len(sParrafo2)
    nValidar = nTamano / 80
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))

    oDoc.WTextBox nTop, 0, nTamano * 10, 580, sParrafo2, "F1", 10, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 10, 10, nTamano * 10, 580, sParrafo2b, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 20, 10, nTamano * 10, 580, sParrafo2c, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 30, 10, nTamano * 10, 580, sParrafo2d, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 40, 10, nTamano * 10, 580, sParrafo2e, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 50, 10, nTamano * 10, 580, sParrafo2f, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 60, 10, nTamano * 10, 580, sParrafo2g, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 70, 10, nTamano * 10, 580, sParrafo2h, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 80, 10, nTamano * 10, 580, sParrafo2i, "F1", 11, hjustify, , , , , , 50

    nTop = nTop + (nTamano * 10)

    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000083", nTipoCredito) Then
    'If nTipoCredito = 514 Then
    '**ARLO20180712 ERS042 - 2018
                    sParrafo3 = "Así mismo DECLARO BAJO JURAMENTO que el dinero depositado en dicha cuenta(s), la misma que " & _
                    "ofrezco en garantía y autorizo su afectación por medio del apresente carta, no es un bien " & _
                    "inerbagable, según lo establecido en el Art. 648° del Código Procesal Civil, y no pertenece a " & _
                    "sociedad conyugal, puesto que es un bien propio y por lo tanto tengo plena facultad para disponer " & _
                    "del crédito depositado."
     Else

                    sParrafo3 = "Sin otro particular, me despido agradeciendo la atención a la presente."
        End If

    nTamano = Len(sParrafo3)
    nValidar = nTamano / 90
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo3, "F1", 12, hjustify, vMiddle, , , , False
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, sParrafo3, "F1", 10, hjustify, , , , , , 50
    'oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 12)
    'FRHU20131126
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000084", nTipoCredito) Then
    'If nTipoCredito = 514 Then
    '**ARLO20180712 ERS042 - 2018
     sParrafo4 = "" & _
                        "Sin otro particular."
    Else
    sParrafo4 = ""
    End If
    nTamano = Len(sParrafo4)
    nValidar = nTamano / 90
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo4, "F1", 12, hjustify, vMiddle, , , , False
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, sParrafo4, "F1", 10, hjustify, , , , , , 50
    'oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    'nTop = nTop + 50

    oDoc.WTextBox nTop + 80, 50, 10, 580, "Atentamente,", "F1", 10, hLeft, , , , , 0
'
'    oDoc.WTextBox nTop + 100, 60, 10, 210, sParrafo5, "F1", 10, hLeft, , , 1, , False '1
'    oDoc.WTextBox nTop + 100, 300, 10, 210, sParrafo5, "F1", 10, hLeft, , , 1, , False '2
'    oDoc.WTextBox nTop + 170, 60, 10, 210, sParrafo5, "F1", 10, hLeft, , , 1, , False '3
'    oDoc.WTextBox nTop + 170, 300, 10, 210, sParrafo5, "F1", 10, hLeft, , , 1, , False '4
'    oDoc.WTextBox nTop + 240, 60, 10, 210, sParrafo5, "F1", 10, hLeft, , , 1, , False '5
'    oDoc.WTextBox nTop + 240, 300, 10, 210, sParrafo5, "F1", 10, hLeft, , , 1, , False '6
     If PersonaGarantia.RecordCount = 1 Then
'        sParrafo5 = "............................................................................" & _
'                "   " & cPersGarantia & "  " & _
'                "  " & cDoi & " N° " & nDoi & " "
            sParrafo5a = "............................................................................"
            sParrafo5b = "" & cPersGarantia & ""
            sParrafo5c = "" & cDOI & " N° " & nDoi & ""
        oDoc.WTextBox nTop + 150, 50, 10, 220, sParrafo5a, "F1", 10, hLeft, , , 0, , False '1
        oDoc.WTextBox nTop + 160, 50, 10, 220, sParrafo5b, "F1", 10, hLeft, , , 0, , False '1
        oDoc.WTextBox nTop + 170, 50, 10, 180, sParrafo5c, "F1", 10, hLeft, , , 0, , False '1
     Else
         Dim nCounter As Integer
         Dim nCounterb As Integer 'FRHU 20171014 : agregado por error
         nCounter = 1
         nCounterb = 0
         Set obj = New COMNCredito.NCOMCredito
         Set PersonaGarantia = obj.ObtenerGarantiaPersona(cCtaCod)
         Set obj = Nothing
         Dim a As Integer

         a = 140 '90

         Do While Not PersonaGarantia.EOF
            cPersCod = PersonaGarantia!cPersCod
            cPersGarantia = PersonaGarantia!cPersNombre
            cDOI = PersonaGarantia!cDOI
            nDoi = PersonaGarantia!nDoi

            sParrafo5a = "............................................................................"
            sParrafo5b = "" & cPersGarantia & ""
            sParrafo5c = "" & cDOI & " N° " & nDoi & ""

            If nCounterb = 2 Then
                a = a + 70
                nCounterb = 0
            End If

            If nCounter Mod 2 <> 0 Then
                oDoc.WTextBox nTop + a, 50, 10, 220, sParrafo5a, "F1", 10, hLeft, , , 0, , False '1
                oDoc.WTextBox nTop + a + 10, 50, 10, 220, sParrafo5b, "F1", 10, hLeft, , , 0, , False '1
                oDoc.WTextBox nTop + a + 20, 50, 10, 180, sParrafo5c, "F1", 10, hLeft, , , 0, , False '1
            Else
                oDoc.WTextBox nTop + a, 330, 10, 220, sParrafo5a, "F1", 10, hLeft, , , 0, , False '2
                oDoc.WTextBox nTop + a + 10, 330, 10, 220, sParrafo5b, "F1", 10, hLeft, , , 0, , False '2
                oDoc.WTextBox nTop + a + 20, 330, 10, 180, sParrafo5c, "F1", 10, hLeft, , , 0, , False '2
            End If

            nCounter = nCounter + 1
            nCounterb = nCounterb + 1
        PersonaGarantia.MoveNext
        Loop
     End If

    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirPDF:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
'FRHU 20171015: SE QUITO DE gFunCorreo
Public Sub EnviarMail(ByVal psHost As String, ByVal psEmailEnvia As String, ByVal psEmailDestino As String, ByVal psAsunto As String, ByVal psContenido As String, _
                      Optional ByVal psEmailCC As String = "", Optional ByVal psEmailCCO As String = "", Optional ByVal psRutaDocAdjunto As String = "")

Dim oSendMail As clsSendMail
Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean

    Set oSendMail = New clsSendMail
    bHtml = True

    psContenido = FormateaContenido(psContenido)

    With oSendMail
        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        .SMTPHost = psHost                          ' Required the fist time, optional thereafter
        .from = psEmailEnvia                        ' Required the fist time, optional thereafter
        '.FromDisplayName = Nombre Envia            ' Optional, saved after first use
        .Recipient = psEmailDestino                 ' Required, separate multiple entries with delimiter character
        '.RecipientDisplayName = Nombre Destino     ' Optional, separate multiple entries with delimiter character
        .CcRecipient = psEmailCC                    ' Optional, separate multiple entries with delimiter character
        '.CcDisplayName = Nombre CC                 ' Optional, separate multiple entries with delimiter character
        .BccRecipient = psEmailCCO                  ' Optional, separate multiple entries with delimiter character
        '.ReplyToAddress = txtFrom.Text             ' Optional, used when different than 'From' address
        .Subject = psAsunto                         ' Optional
        .Message = psContenido                      ' Optional
        .Attachment = Trim(psRutaDocAdjunto)        ' Optional, separate multiple entries with delimiter character

        .AsHTML = bHtml                              ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MyEncodeType                     ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised

        .Send                                       ' Required
    End With

    Set oSendMail = Nothing
End Sub

Private Function FormateaContenido(ByVal psCadena As String) As String

    psCadena = "<font style='font-family:Calibri,Arial; font-size:14.5px'>" & psCadena & "<p><p>" & _
               "<b>TECNOLOGIA DE INFORMACIÓN</b></font>"

    FormateaContenido = psCadena
End Function
'***********APRI2018 ERS036-2017
Public Sub GeneraEstadoCuentaCredito(ByVal FEDatos As Variant)

    Dim oDoc As New cPDF
    Dim rsAge As New ADODB.Recordset
    Dim R As New ADODB.Recordset
    Dim nIndex As Integer
    Dim Contador As Integer
    Dim nCentrar As Integer
    Dim nTamTit As Integer
    Dim nTamLet As Integer
    Dim nTamSubTit As Integer
    Dim nTamPie As Integer
    Dim sCiudadEmite As String
    Dim nCantDoc As Integer
    Dim sParrafo1 As String
    Dim sParrafo2 As String
    Dim sParrafo3 As String


    nCantDoc = 1
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Operaciones"
    oDoc.Producer = gsNomCmac
    oDoc.Subject = "EMISIÓN DE ESTADO DE CUENTA"
    oDoc.Title = "EMISIÓN DE ESTADO DE CUENTA"
    If Not oDoc.PDFCreate(App.Path & "\Spooler\EstadoCuentaCredito_" & Trim(FEDatos.TextMatrix(nIndex, 3)) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then Exit Sub

    oDoc.Fonts.Add "F1", "Arial Narrow", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial Narrow", TrueType, Bold, WinAnsiEncoding

    oDoc.LoadImageFromFile App.Path & "\Logo_2015.jpg", "Logo"
    nTamTit = 16: nTamSubTit = 15: nTamLet = 11: nTamPie = 8: Contador = 0: nCentrar = 80

    For nIndex = 1 To FEDatos.Rows - 1
        If FEDatos.TextMatrix(nIndex, 1) = "." Then

        oDoc.NewPage A4_Vertical


        oDoc.WImage 110 + Contador, 450, 70, 120, "Logo"
        oDoc.WTextBox 45 + Contador, 25, 20, 300, "ESTADO DE SITUACIÓN DEL PRESTAMO", "F2", nTamTit, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 85 + Contador, 25, 20, 100, "AL: " & gdFecSis, "F2", nTamSubTit, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 135 + Contador, 25, 20, 200, "Sr(a): ", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 135 + Contador, 55, 20, 200, FEDatos.TextMatrix(nIndex, 2), "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        'oDoc.WTextBox 195 + contador, 25, 20, 300, "DIRECCIÓN / DISTRITO / PROVINCIA: ", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 155 + Contador, 25, 20, 500, FEDatos.TextMatrix(nIndex, 4) & " / " & FEDatos.TextMatrix(nIndex, 19) & " / " & FEDatos.TextMatrix(nIndex, 18), "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack


        oDoc.WTextBox 195 + Contador, 30, 15, 100, "Crédito N°", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 195 + Contador, 135, 15, 100, ":   " & FEDatos.TextMatrix(nIndex, 3), "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 210 + Contador, 30, 15, 100, "Fecha de Desembolso", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 210 + Contador, 135, 15, 100, ":   " & Format(FEDatos.TextMatrix(nIndex, 29), "dd/MM/yyyy"), "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 225 + Contador, 30, 15, 50, "Plazo", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 225 + Contador, 135, 15, 50, ":   " & FEDatos.TextMatrix(nIndex, 23) & " Cuotas", "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 240 + Contador, 30, 15, 100, "Cuotas Pagadas", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 240 + Contador, 135, 15, 100, ":   " & FEDatos.TextMatrix(nIndex, 5), "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 255 + Contador, 30, 15, 50, "TEA        : ", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 255 + Contador, 135, 15, 50, ":   " & FEDatos.TextMatrix(nIndex, 21) & "%", "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack

        oDoc.WTextBox 210 + Contador, 350, 15, 100, "Monto", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 210 + Contador, 430, 15, 100, ":   " & Format(FEDatos.TextMatrix(nIndex, 20), gsFormatoNumeroView), "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 225 + Contador, 350, 15, 50, "Moneda", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 225 + Contador, 430, 15, 50, ":   " & IIf(Mid(FEDatos.TextMatrix(nIndex, 3), 9, 1) = "1", "SOLES", "DÓLARES"), "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 240 + Contador, 350, 15, 100, "Cuotas por Pagar", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 240 + Contador, 430, 15, 100, ":   " & FEDatos.TextMatrix(nIndex, 6), "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 255 + Contador, 350, 15, 50, "TCEA        : ", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 255 + Contador, 430, 15, 50, ":   " & FEDatos.TextMatrix(nIndex, 22) & "%", "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack


         oDoc.WTextBox 190 + Contador, 25, 85, 545, "", "F1", nTamLet, hCenter, , , 1, vbBlack

        oDoc.WTextBox 285 + Contador, 25, 20, 230, "DETALLE DEL ÚLTIMO PAGO REALIZADO", "F2", nTamLet, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oDoc.WTextBox 285 + Contador, 25, 165, 150, "", "F1", nTamLet, hCenter, , , 1, vbBlack
        oDoc.WTextBox 285 + Contador, 175, 165, 80, "", "F1", nTamLet, hCenter, , , 1, vbBlack

        oDoc.WTextBox 310 + Contador, 30, 15, 150, "Cuota N°", "F1", nTamLet, hLeft, vMiddle, vbBlack
        oDoc.WTextBox 325 + Contador, 30, 15, 150, "Capital", "F1", nTamLet, hLeft, vMiddle, vbBlack
        oDoc.WTextBox 340 + Contador, 30, 15, 150, "Inter. Compesatorios", "F1", nTamLet, hLeft, vMiddle, vbBlack
        oDoc.WTextBox 355 + Contador, 30, 15, 150, "Inter. Moratorios", "F1", nTamLet, hLeft, vMiddle, vbBlack
        oDoc.WTextBox 370 + Contador, 30, 15, 150, "Gastos", "F1", nTamLet, hLeft, vMiddle, vbBlack
        oDoc.WTextBox 385 + Contador, 30, 15, 150, "Inter. Gracia", "F1", nTamLet, hLeft, vMiddle, vbBlack
        oDoc.WTextBox 400 + Contador, 30, 15, 150, "Comisión por envío de EECC", "F1", nTamLet, hLeft, vMiddle, vbBlack
        oDoc.WTextBox 415 + Contador, 30, 15, 150, "Monto Total Pagado (1)", "F1", nTamLet, hLeft, vMiddle, vbBlack
        oDoc.WTextBox 430 + Contador, 30, 15, 150, "Fecha Pago", "F1", nTamLet, hLeft, vMiddle, vbBlack


            oDoc.WTextBox 310 + Contador, 190, 15, 50, FEDatos.TextMatrix(nIndex, 7), "F1", nTamLet, hCenter, vMiddle, vbBlack
            oDoc.WTextBox 325 + Contador, 190, 15, 50, Format(FEDatos.TextMatrix(nIndex, 8), gsFormatoNumeroView), "F1", nTamLet, hCenter, vMiddle, vbBlack
            oDoc.WTextBox 340 + Contador, 190, 15, 50, Format(FEDatos.TextMatrix(nIndex, 9), gsFormatoNumeroView), "F1", nTamLet, hCenter, vMiddle, vbBlack
            oDoc.WTextBox 355 + Contador, 190, 15, 50, Format(FEDatos.TextMatrix(nIndex, 10), gsFormatoNumeroView), "F1", nTamLet, hCenter, vMiddle, vbBlack
            oDoc.WTextBox 370 + Contador, 190, 15, 50, Format(FEDatos.TextMatrix(nIndex, 11), gsFormatoNumeroView), "F1", nTamLet, hCenter, vMiddle, vbBlack
            oDoc.WTextBox 385 + Contador, 190, 15, 50, Format(FEDatos.TextMatrix(nIndex, 12), gsFormatoNumeroView), "F1", nTamLet, hCenter, vMiddle, vbBlack
            oDoc.WTextBox 400 + Contador, 190, 15, 50, Format(FEDatos.TextMatrix(nIndex, 24), gsFormatoNumeroView), "F1", nTamLet, hCenter, vMiddle, vbBlack
            oDoc.WTextBox 415 + Contador, 190, 15, 50, FEDatos.TextMatrix(nIndex, 16), "F1", nTamLet, hCenter, vMiddle, vbBlack
            oDoc.WTextBox 430 + Contador, 190, 15, 50, FEDatos.TextMatrix(nIndex, 14), "F1", nTamLet, hCenter, vMiddle, vbBlack


            oDoc.WTextBox 450 + Contador, 25, 20, 150, "", "F1", nTamLet, hCenter, , , 1, vbBlack, 1, vbBlack, True
            oDoc.WTextBox 450 + Contador, 30, 20, 150, "SALDO DE CAPITAL", "F2", nTamLet, hLeft, vMiddle, vbWhite ', 1, vbBlack, True
            oDoc.WTextBox 450 + Contador, 175, 20, 80, "", "F1", nTamLet, hCenter, , , 1, vbBlack
            oDoc.WTextBox 450 + Contador, 190, 20, 50, Format(FEDatos.TextMatrix(nIndex, 25), gsFormatoNumeroView), "F1", nTamLet, hCenter, vMiddle, vbBlack ', 1, vbBlack

            oDoc.WTextBox 325 + Contador, 315, 30, 250, "PRÓXIMA CUOTA A PAGAR N° " & FEDatos.TextMatrix(nIndex, 15) & " (2)", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack

            oDoc.WTextBox 355 + Contador, 285, 15, 70, "FECHA", "F2", nTamLet, hCenter, vMiddle, vbWhite, 1, vbBlack, True
            oDoc.WTextBox 355 + Contador, 285, 30, 70, "", "F2", nTamLet, hCenter, , , 1, vbBlack
            oDoc.WTextBox 355 + Contador, 355, 15, 90, "DESCRIPCIÓN", "F2", nTamLet, hCenter, vMiddle, vbWhite, 1, vbBlack, True
            oDoc.WTextBox 355 + Contador, 355, 30, 90, "", "F2", nTamLet, hCenter, , , 1, vbBlack
            oDoc.WTextBox 355 + Contador, 445, 15, 70, " IMPORTE", "F2", nTamLet, hCenter, vMiddle, vbWhite, 1, vbBlack, True
            oDoc.WTextBox 355 + Contador, 445, 30, 70, "", "F2", nTamLet, hCenter, , , 1, vbBlack

            oDoc.WTextBox 370 + Contador, 295, 15, 70, Format(FEDatos.TextMatrix(nIndex, 26), "DD/MM/YYYY"), "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
            oDoc.WTextBox 370 + Contador, 365, 15, 70, "Cuota Pendiente", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
            oDoc.WTextBox 370 + Contador, 460, 15, 50, Format(FEDatos.TextMatrix(nIndex, 27), gsFormatoNumeroView), "F2", nTamLet, hCenter, vMiddle, vbBlack ', 1, vbBlack


        sParrafo1 = "Si tiene alguna consulta o requiere mayor información sobre las operaciones detalladas en su estado de cuenta " & _
                    "o sobre los servicios que le ofrecemos, puede contactarnos a través de nuestro Call Center desde Lima y Provincias " & _
                    "al 0-801-10700 y desde Iquitos Llamando al (065) 581800 o acudir a nuestra de red de agencias."
         sParrafo2 = "En caso tuviera algún reclamo que presentar, podrá realizarlo a través de nuestra pagina Web, Call Center " & _
                    "u Oficinas con un plazo de 30 días calendario, posteriores a los cuales no se aceptará reclamo alguno. Si no " & _
                    "se encuentra conforme con la respuesta emitida, puede proceder a la reiteración de su reclamo, en caso de " & _
                    "de continuar disconforme con la respuesta useted podrá recurrir a Indecopi ó a la Plataforma de Atención al Usuario " & _
                    "de la Superintendencia de Banca y Seguros según corresponda."


        sParrafo3 = "EL PRESENTE DOCUMENTO TIENE DE CARÁCTER INFORMATIVO Y UNICAMENTE TIENE VALOR PARA LOS FINES DE " & _
                    "CUMPLIMIENTO DE LAS OBLIGACIONES DEL CLIENTE FRENTE A LA CAJA QUE SON MATERIA DEL PRESTAMO " & _
                    "RESPECTIVO."


          oDoc.WTextBox 480 + Contador, 30, 15, 200, "INFORMACIÓN AL CLIENTE ", "F2", nTamPie, hLeft, vMiddle, vbBlack ', 1, vbBlack
          oDoc.WTextBox 490 + Contador, 30, 30, 540, sParrafo1, "F1", nTamPie, hjustify, vMiddle, vbBlack ', 1, vbBlack
          oDoc.WTextBox 505 + Contador, 30, 50, 540, sParrafo2, "F1", nTamPie, hjustify, vMiddle, vbBlack ', 1, vbBlack
          oDoc.WTextBox 480 + Contador, 25, 75, 545, "", "F1", nTamLet, hCenter, , , 1, vbBlack

          oDoc.WTextBox 560 + Contador, 25, 15, 100, "Notas:", "F2", nTamPie, hLeft, vMiddle, vbBlack ', 1, vbBlack
          oDoc.WTextBox 570 + Contador, 25, 15, 150, "(1) Importe no incluye ITF.", "F1", nTamPie, hLeft, vMiddle, vbBlack ', 1, vbBlack
          oDoc.WTextBox 580 + Contador, 25, 15, 500, "(2) Importe de cuota aplicable únicamente en el caso de pagos puntuales. Importe no incluye ITF.", "F1", nTamPie, hLeft, vMiddle, vbBlack ', 1, vbBlack
          oDoc.WTextBox 590 + Contador, 25, 15, 300, "Tasa Actual del ITF es 0.005%", "F1", nTamPie, hLeft, vMiddle, vbBlack

          oDoc.WTextBox 600 + Contador, 25, 30, 545, sParrafo3, "F1", nTamPie, hLeft, vMiddle, vbBlack




            End If
        Next

    oDoc.PDFClose
    oDoc.Show
End Sub
'***********END APRI

'JOEP20180725 ERS034-2018
Public Function ConsultaRiesgoCamCred(ByVal pcCtaCod As String) As Boolean
Dim objCredRiegCamb As COMDCredito.DCOMCredito
Dim rsRiegCamb As ADODB.Recordset
ConsultaRiesgoCamCred = True
    Set objCredRiegCamb = New COMDCredito.DCOMCredito
    Set rsRiegCamb = objCredRiegCamb.ValidadRigCambCred(pcCtaCod)
        If Not (rsRiegCamb.BOF And rsRiegCamb.EOF) Then
            If rsRiegCamb!nApli = 1 Then
                ConsultaRiesgoCamCred = False
            End If
        End If
RSClose rsRiegCamb
Set objCredRiegCamb = Nothing
End Function
Public Sub EmiteFormRiesgoCamCred(ByVal pcCtaCod As String)
Dim objCredRiegCamb As COMDCredito.DCOMCredito
Dim rsRiegCambF As ADODB.Recordset
    Set objCredRiegCamb = New COMDCredito.DCOMCredito
    Set rsRiegCambF = objCredRiegCamb.ValidadRigCambCred(pcCtaCod)
        If Not (rsRiegCambF.BOF And rsRiegCambF.EOF) Then
            If rsRiegCambF!nApli = 1 Then
                Call frmCredFormEvalCredCel.inicio(pcCtaCod, 10)
            End If
        End If
RSClose rsRiegCambF
Set objCredRiegCamb = Nothing
End Sub
'JOEP20180725 ERS034-2018
