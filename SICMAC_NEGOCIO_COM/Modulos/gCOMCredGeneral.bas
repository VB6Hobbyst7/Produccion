Attribute VB_Name = "gCOMCredGeneral"
Global gnTipCambio As Double
Global gnTipCambioV As Double
Global gnTipCambioC As Double
Public Function FillText(psCadena As String, pnLenTex As Integer, ChrFil As String) As String
    If pnLenTex > Len(Trim(psCadena)) Then
       FillText = Trim(psCadena) & String((pnLenTex - Len(Trim(psCadena))), ChrFil)
    Else
       FillText = Trim(Left(psCadena, pnLenTex)) & String((pnLenTex - Len(Trim(psCadena))), ChrFil)
    End If
End Function

Public Function FechaHora(ByVal psFecha As Date) As String
    FechaHora = Format(psFecha & Space(1) & GetHoraServer, "mm/dd/yyyy hh:mm:ss")
End Function
Public Function GetFechaMov(cMovnro, lDia As Boolean) As String
Dim lFec As Date
lFec = Mid(cMovnro, 7, 2) & "/" & Mid(cMovnro, 5, 2) & "/" & Mid(cMovnro, 1, 4)
If lDia Then
   GetFechaMov = Format(lFec, gsFormatoFechaView)
Else
   GetFechaMov = Format(lFec, gsFormatoFecha)
End If
End Function
'Devuelve un string formateado de acuerdo a los parametros ingresados
' se utiliza con numeros y caracteres
Public Function ImpreFormat(ByVal pNumero As Variant, ByVal pLongitudEntera As Integer, _
        Optional ByVal pLongitudDecimal As Integer = 2, _
        Optional ByVal pMoneda As Boolean = False) As String
Dim vPosPto As Integer
Dim vParEnt As String
Dim vParDec As String
Dim vLonEnt As Integer
Dim vLonDec As Integer
Dim X As Integer

On Error GoTo ErrHandler
vParDec = ""
If IsNull(pNumero) Then
    If pLongitudDecimal > 0 Then vParDec = "." & String(pLongitudDecimal, "0")
    If pLongitudEntera <= 0 Then pLongitudEntera = 1
    ImpreFormat = String(pLongitudEntera - 1, " ") & "0" & vParDec
ElseIf VarType(pNumero) = 8 Then
    pNumero = Trim(pNumero)
    vLonEnt = Len(pNumero)
    If vLonEnt > pLongitudEntera Then
        pNumero = Left(pNumero, pLongitudEntera)
        vLonEnt = pLongitudEntera
    End If
    ImpreFormat = String(pLongitudDecimal, " ") & pNumero & String(pLongitudEntera - vLonEnt, " ")
Else
    vPosPto = InStr(Trim(CStr(pNumero)), ".")
    If vPosPto > 0 Then
        vParEnt = Trim(CStr(Left(pNumero, vPosPto - 1)))
        vParDec = Trim(CStr(Mid(pNumero, vPosPto + 1)))
        vLonEnt = Len(vParEnt)
        vLonDec = Len(vParDec)
    Else
        vParEnt = Trim(Str(pNumero))
        vParDec = ""
        vLonEnt = Len(vParEnt)
        vLonDec = 0
    End If
    If pMoneda And vLonEnt > 3 Then
        vParEnt = Format(vParEnt, "#,###,###")
        For X = 1 To Len(vParEnt)
            If Mid(vParEnt, X, 1) = "," Then pLongitudEntera = pLongitudEntera - 1
        Next X
    End If
    If vLonEnt > pLongitudEntera Then pLongitudEntera = vLonEnt + 1
    If vLonDec > pLongitudDecimal Then
        vLonDec = pLongitudDecimal
        vParDec = Left(vParDec, vLonDec)
    End If
    ImpreFormat = String(pLongitudEntera - vLonEnt, " ") & vParEnt
    If pLongitudDecimal > 0 Then
        ImpreFormat = ImpreFormat & "." & vParDec & String(pLongitudDecimal - vLonDec, "0")
    End If
End If
Exit Function

ErrHandler:     ' Errores obtenidos
    MsgBox " Operación no válida " & vbCr & _
        " Error " & Err.Number & " : " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " ! Aviso ! "
End Function
Public Function Encripta(pnTexto As String, Valor As Boolean) As String
'true = encripta
'false = desencripta
Dim MiClase As cEncrypt
Set MiClase = New cEncrypt
Encripta = MiClase.ConvertirClave(pnTexto, , Valor)
End Function

Public Function GetTipCambio(pdFecha As Date) As Boolean

Dim oDGeneral As DCOMGeneral
Set oDGeneral = New DCOMGeneral
GetTipCambio = True
gnTipCambio = 0
gnTipCambioV = 0
gnTipCambioC = 0

 gnTipCambio = oDGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
 gnTipCambioV = oDGeneral.EmiteTipoCambio(pdFecha, TCVenta)
 gnTipCambioC = oDGeneral.EmiteTipoCambio(pdFecha, TCCompra)

If gnTipCambio = 0 Then
    GetTipCambio = False
End If
End Function
