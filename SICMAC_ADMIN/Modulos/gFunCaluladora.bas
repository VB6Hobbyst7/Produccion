Attribute VB_Name = "gFunCaluladora"
Option Explicit

Public Function CalEvaluaBil(psCadena As String, Optional pbSoloPositivos As Boolean = True) As Currency
    Dim lsCadena As String
    Dim lsOpe As String
    Dim lsNum1 As String
    Dim lsNum2 As String
    Dim lsResp As String
    
    Dim lnNum1 As String
    Dim lnNum2 As String
    
    Dim lnPosMas As Integer
    Dim lnPosMenos As Integer
    Dim lnPosMasN As Integer
    Dim lnPosMenosN As Integer
    
    psCadena = CalPreCadenaBil(psCadena)
    
    If psCadena = "" Then
        CalEvaluaBil = Format(0, "#0.00")
        MsgBox "Error de validación.", vbInformation, "Aviso"
        Exit Function
    End If
    
    lsCadena = psCadena
    
    lnPosMas = InStr(1, psCadena, "+")
    lnPosMenos = InStr(1, psCadena, "-")

    If lnPosMas = 0 And lnPosMenos = 0 Then
        If IsNumeric(psCadena) Then
            lnNum1 = CCur(psCadena)
            lsCadena = ""
        Else
            GoTo NOVALIDO
        End If
    ElseIf (lnPosMas < lnPosMenos) And lnPosMas <> 0 Or lnPosMenos = 0 Then
        If IsNumeric(Mid(psCadena, 1, lnPosMas - 1)) Then
            lnNum1 = CCur(Mid(psCadena, 1, lnPosMas - 1))
            lsCadena = Mid(psCadena, lnPosMas)
        Else
            GoTo NOVALIDO
        End If
    ElseIf (lnPosMas > lnPosMenos) And lnPosMenos <> 0 Or lnPosMas = 0 Then
        If IsNumeric(Mid(psCadena, 1, lnPosMenos - 1)) Then
            lnNum1 = CCur(Mid(psCadena, 1, lnPosMenos - 1))
            lsCadena = Mid(psCadena, lnPosMenos)
        Else
            GoTo NOVALIDO
        End If
    End If
    
    While lsCadena <> ""
        CalObtNumBil lsCadena, lsNum2, lsOpe
        If IsNumeric(lsNum2) Then
            lnNum1 = CalGetResOpeBil(Str(lnNum1), lsNum2, lsOpe)
        Else
            GoTo NOVALIDO
        End If
    Wend
    
    If pbSoloPositivos Then
        If lnNum1 < 0 Then
            CalEvaluaBil = Format(0, "#0.00")
            MsgBox "El resultado no puede ser Negativo.", vbInformation, "Aviso"
            Exit Function
        End If
    End If
    
    CalEvaluaBil = Format(lnNum1, "#,#00.00")
    
    Exit Function
    
NOVALIDO:
    CalEvaluaBil = 0
    MsgBox "Error Ud. ha ingresado un valor no Valido", vbInformation, "Aviso"
End Function

Public Function CalObtNumBil(psCadena As String, psNum2 As String, psOpe As String) As String
    Dim lnPosMas As Integer
    Dim lnPosMenos As Integer
    Dim lnPosMasN As Integer
    Dim lnPosMenosN As Integer
    
    psOpe = Mid(psCadena, 1, 1)
    psCadena = Mid(psCadena, 2)
    
    lnPosMas = InStr(1, psCadena, "+")
    lnPosMenos = InStr(1, psCadena, "-")

    If lnPosMas = 0 And lnPosMenos = 0 Then
        psNum2 = psCadena
        psCadena = ""
    ElseIf (lnPosMas < lnPosMenos) And lnPosMas <> 0 Or lnPosMenos = 0 Then
        psNum2 = Mid(psCadena, 1, lnPosMas - 1)
        psCadena = Mid(psCadena, lnPosMas)
    ElseIf (lnPosMas > lnPosMenos) And lnPosMenos <> 0 Or lnPosMas = 0 Then
        psNum2 = Mid(psCadena, 1, lnPosMenos - 1)
        psCadena = Mid(psCadena, lnPosMenos)
    End If
    
End Function

Public Function CalPreCadenaBil(psCadena As String) As String
    Dim i As Integer
    Dim lsCadena As String
 
    psCadena = Trim(psCadena)
    lsCadena = ""
    For i = 1 To Len(psCadena)
        If Mid(psCadena, i, 1) <> " " Then
            lsCadena = lsCadena & Mid(psCadena, i, 1)
        End If
    Next i
    
    If InStr(1, psCadena, "++") = 0 _
    And InStr(1, psCadena, "+-") = 0 _
    And InStr(1, psCadena, "-+") = 0 _
    And InStr(1, psCadena, "--") = 0 Then
        CalPreCadenaBil = lsCadena
    Else
        CalPreCadenaBil = ""
    End If
End Function

Public Function CalGetResOpeBil(psNum1 As String, psNum2 As String, psOpe As String) As Currency
    If psOpe = "+" Then
        CalGetResOpeBil = CCur(psNum1) + CCur(psNum2)
    Else
        CalGetResOpeBil = CCur(psNum1) - CCur(psNum2)
    End If
End Function

Public Function CalValidaIngBil(pnCodigo As Integer) As Integer
     If InStr(1, "0123456789+-.", Chr$(pnCodigo)) = 0 And pnCodigo <> 8 Then
        CalValidaIngBil = 0
     Else
        CalValidaIngBil = pnCodigo
     End If
End Function


