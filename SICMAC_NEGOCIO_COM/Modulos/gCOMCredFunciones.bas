Attribute VB_Name = "gCOMCredFunciones"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function CentrarCadena(psCadena As String, pnNroLineas As Long, Optional lnEspaciosIzq As Integer = 0, Optional lsCarImp As String = " ") As String
    Dim psNinf As Long
    Dim lnPosIni As Long
    
    psCadena = Trim(psCadena)
    If Len(psCadena) > pnNroLineas Then
        'psCadena = Left(psCadena, pnNroLineas)
        'MsgBox "EL valor de la Cadena enviada es mayor al espacio destinado", vbInformation, "Aviso"
        psCadena = Left(psCadena, pnNroLineas)
    End If
    'Else
    psNinf = Len(psCadena) / 2
    lnPosIni = Int(pnNroLineas / 2) - Int(Len(psCadena) / 2)
    
    'psCadena = String((pnNroLineas / 2) - psNinf, " ") & psCadena & String(pnNroLineas - Len(psCadena), " ")
    psCadena = String(lnEspaciosIzq, " ") & String(lnPosIni, lsCarImp) & psCadena & String(lnPosIni, lsCarImp)
    CentrarCadena = psCadena
   'End If
End Function
Public Sub CargaComboPersonasTipo(ByVal psConstante As PersTipo, ByRef pM_cmbInstitucion As Variant)
Dim oPersonas As DCOMPersonas
Dim Datos() As String
Dim R As ADODB.Recordset
    On Error GoTo ERRORCargaComboPersonasTipo
    
    Set oPersonas = New DCOMPersonas
    Set R = oPersonas.RecuperaPersonasTipo(Trim(Str(psConstante)))
    Set oPersonas = Nothing
    ReDim Datos(R.RecordCount)
    Do While Not R.EOF
        Datos(R.Bookmark - 1) = PstaNombre(R!cPersNombre) & Space(250) & R!cPersCod
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    pM_cmbInstitucion = Datos
    Exit Sub

ERRORCargaComboPersonasTipo:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Public Sub CargaComboConstante(ByVal pnCodCons As Integer, ByRef Combo As Variant, _
    Optional ByVal pnFiltro As Integer = -1)
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConstante As DCOMConstante
Dim Datos() As String

    Set oConstante = New DCOMConstante
    Set R = oConstante.RecuperaConstantes(Trim(Str(pnCodCons)), pnFiltro)
    ReDim Datos(R.RecordCount)
    Do While Not R.EOF
        Datos(R.Bookmark - 1) = Trim(R!cConsDescripcion) & Space(100) & Trim(Str(R!nConsValor))
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oConstante = Nothing
    
    Combo = Datos
End Sub
Public Function FillNum(intNumero As String, intLenNum As Integer, ChrFil As String) As String
'On Error Resume Next
  FillNum = Left(String(intLenNum, ChrFil), (Len(String(intLenNum, ChrFil)) - Len(Trim(intNumero)))) + Trim(intNumero)
End Function

'Prepara una cadena especial (cadena con caracteres con tilde y/o otros)
' para que se imprima en el modo FREEFILE.
Public Function ImpreCarEsp(ByVal vCadena As String) As String
    vCadena = Replace(vCadena, "á", Chr(160), , , vbTextCompare)
    vCadena = Replace(vCadena, "é", Chr(130), , , vbTextCompare)
    vCadena = Replace(vCadena, "í", Chr(161), , , vbTextCompare)
    vCadena = Replace(vCadena, "ó", Chr(162), , , vbTextCompare)
    vCadena = Replace(vCadena, "ú", Chr(163), , , vbTextCompare)
    vCadena = Replace(vCadena, "ñ", Chr(164), , , vbTextCompare)
    vCadena = Replace(vCadena, "Ñ", Chr(165), , , vbTextCompare)
    vCadena = Replace(vCadena, "°", Chr(248), , , vbTextCompare)
    vCadena = Replace(vCadena, "¦", Chr(179), , , vbTextCompare)
    ImpreCarEsp = vCadena
End Function
Public Function fgFechaHoraGrab(ByVal psMovNro As String) As String
    fgFechaHoraGrab = Mid(psMovNro, 1, 4) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 7, 2) & " " & Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2)
End Function
Public Function PstaNombre(psNombre As String, Optional pbNombApell As Boolean = False) As String
Dim Total As Long
Dim pos As Long
Dim CadAux As String
Dim lsApellido As String
Dim lsNombre As String
Dim lsMaterno As String
Dim lsConyugue As String
Dim CadAux2 As String
Dim posAux As Integer
Dim lbVda As Boolean
lbVda = False
Total = Len(Trim(psNombre))
pos = InStr(psNombre, "/")
If pos <> 0 Then
    lsApellido = Left(psNombre, pos - 1)
    CadAux = Mid(psNombre, pos + 1, Total)
    pos = InStr(CadAux, "\")
    If pos <> 0 Then
        lsMaterno = Left(CadAux, pos - 1)
        CadAux = Mid(CadAux, pos + 1, Total)
        pos = InStr(CadAux, ",")
        If pos > 0 Then
            CadAux2 = Left(CadAux, pos - 1)
            posAux = InStr(CadAux, "VDA")
            If posAux = 0 Then
                lsConyugue = CadAux2
            Else
                lbVda = True
                lsConyugue = CadAux2
            End If
        Else
            lsMaterno = CadAux
        End If
    Else
        CadAux = Mid(CadAux, pos + 1, Total)
        pos = InStr(CadAux, ",")
        If pos <> 0 Then
            lsMaterno = Left(CadAux, pos - 1)
            lsConyugue = ""
        Else
            lsMaterno = CadAux
        End If
    End If
    lsNombre = Mid(CadAux, pos + 1, Total)
    If pbNombApell = True Then
        If Len(Trim(lsConyugue)) > 0 Then
            PstaNombre = Trim(lsNombre) & " " & Trim(lsApellido) & " " & Trim(lsMaterno) & IIf(lbVda = False, " DE ", " ") & Trim(lsConyugue)
        Else
            PstaNombre = Trim(lsNombre) & " " & Trim(lsApellido) & " " & Trim(lsMaterno)
        End If
    Else
        If Len(Trim(lsConyugue)) > 0 Then
            PstaNombre = Trim(lsApellido) & " " & Trim(lsMaterno) & IIf(lbVda = False, " DE ", " ") & Trim(lsConyugue) & " " & Trim(lsNombre)
        Else
            PstaNombre = Trim(lsApellido) & " " & Trim(lsMaterno) & " " & Trim(lsNombre)
        End If
    End If
Else
    PstaNombre = Trim(psNombre)
End If
End Function

Public Function GetHoraServer() As String
Dim sql As String
Dim rsH As New ADODB.Recordset
Dim coConex As DCOMConecta

Set coConex = New DCOMConecta
If coConex.AbreConexion = False Then Exit Function
sql = "Select Convert(varchar(10),getdate(),108) as sHora"
Set rsH = coConex.CargaRecordSet(sql)
If Not rsH.EOF Then
   GetHoraServer = rsH!sHora
Else
   GetHoraServer = Format(Time, "hh:mm:ss")
End If
rsH.Close
Set rsH = Nothing
coConex.CierraConexion
Set coConex = Nothing
End Function

Public Function TasaIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer) As Double
    TasaIntPerDias = ((1 + pnTasaInter / 100) ^ (pnDiasTrans / 30)) - 1
End Function

Public Function MontoIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer, ByVal pnMonto As Double) As Double
    MontoIntPerDias = (((1 + pnTasaInter / 100) ^ (pnDiasTrans / 30)) - 1) * pnMonto
    MontoIntPerDias = CDbl(Format(MontoIntPerDias, "#0.00"))
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
