VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private MatGastos() As String

Public Function GeneraCalendarioGastos(ByRef MatCalend As Variant, _
    ByRef MatDesemb As Variant, ByRef nNumGastos As Integer, _
    ByVal pdHoy As Date, ByVal psCtaCod As String, ByVal pnColocCred As Integer, _
    ByVal psAplicaproceso As String, Optional ByVal psGastoFijo As String = "F", Optional ByVal pnCuotaApr As Double = -1, _
    Optional ByVal pnPrestamo As Double = -1, _
    Optional ByVal pnMontoCapital As Double = 0, Optional ByVal pnCuotaPend As Integer = -1, _
    Optional ByVal pHabRecorsed As Boolean = False, _
    Optional ByVal PR As ADODB.Recordset = Nothing, Optional ByVal pRFiltro As ADODB.Recordset = Nothing, _
    Optional ByVal pRFiltroGar As ADODB.Recordset = Nothing, Optional pnDiasAtraso As Integer = -1) As Variant
        
Dim oGastos As DGastos
Dim R As ADODB.Recordset
Dim i As Integer
Dim sProd As String
Dim RFiltro As ADODB.Recordset
Dim bAplicaGasto As Boolean
Dim nRes As Double
Dim oCred As DCredito
Dim oNCred As NCredito
Dim nPrestamo As Double
Dim nCuotaApr As Double
Dim nMontoGarantia As Double
Dim oGar As DGarantia
Dim nDiasAtraso As Integer
Dim sGastoFijoVariable As String
Dim oCredRel As UCredRelacion
Dim RFiltroGar As ADODB.Recordset
Dim RGar As ADODB.Recordset
Dim sTipoRelacGarantia As String
Dim bCumpleRestricc As Boolean
Dim nMontoTemp As Double
Dim RCredVig As ADODB.Recordset
Dim MatCalTemp As Variant
Dim MontoPrePagado As Double
Dim MatGastosExon() As String
Dim RGastosExon As ADODB.Recordset

    On Error GoTo ErrorGeneraCalendarioGastos
    'On Error GoTo 0
    'Recuperando el Monto de prestamo y la cuota aprobada
    Set oCred = New DCredito
    If psAplicaproceso <> "CD" Then
        Set R = oCred.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
        If pnCuotaApr = -1 Then
            sGastoFijoVariable = IIf(IsNull(R!cTipoGasto), "F", "V")
            Set oNCred = New NCredito
            MatCalTemp = oNCred.RecuperaMatrizCalendarioInicial(psCtaCod)
            Set oNCred = Nothing
            nCuotaApr = Format(CDbl(MatCalTemp(0, 3)) + CDbl(MatCalTemp(0, 4)) + CDbl(MatCalTemp(0, 5)) + CDbl(MatCalTemp(0, 6)) + CDbl(MatCalTemp(0, 7)), "#0.00")
        Else
            sGastoFijoVariable = psGastoFijo
            nCuotaApr = pnCuotaApr
        End If
        If pnPrestamo = -1 Then
            nPrestamo = Format(R!nMonto, "#0.000000")
        Else
            nPrestamo = pnPrestamo
        End If
        If psAplicaproceso = "PA" Or psAplicaproceso = "CA" Or psAplicaproceso = "PP" Then
            ReDim MatDesemb(1, 1)
            MatDesemb(0, 0) = Format(R!nMonto, "#0.000000")
        End If
        
        R.Close
        Set R = Nothing
        
        Set R = oCred.RecuperaColocacCred(psCtaCod)
        nDiasAtraso = IIf(IsNull(R!nDiasAtraso), 0, R!nDiasAtraso)
        R.Close
        Set oCred = Nothing
        
        Set oCredRel = New UCredRelacion
        Call oCredRel.CargaRelacPersCred(psCtaCod)
        If oCredRel.ExisteMasDeUnTitular Then
            sTipoRelacGarantia = "M"
        Else
            sTipoRelacGarantia = "I"
        End If
        
        Set oGar = New DGarantia
'        nMontoGarantia = oGar.RecuperaMontoGarantiaCredito(psCtaCod)
        Set RGar = oGar.RecuperaGarantiaCredito(psCtaCod)
        Set oGar = Nothing
        
        
    Else
        nPrestamo = pnPrestamo
    End If
        
    
    If psAplicaproceso = "DE" Then
        
        pdHoy = CDate(MatCalend(UBound(MatCalend) - 1, 0))
        
        nNumGastos = 0
        ReDim MatGastos(10000, 4)
        Set oGastos = New DGastos
        Set R = oGastos.RecuperaGastosAplicablesDesembolso(CInt(Mid(psCtaCod, 9, 1)), , psAplicaproceso, sGastoFijoVariable)
        Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P")
        Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G")
        Set oGastos = Nothing
        
        'Aplicar Gastos al Desembolso
        If IsArray(MatDesemb) Then
            ReDim Preserve MatDesemb(UBound(MatDesemb), 3)
            Do While Not R.EOF
                For i = 0 To UBound(MatDesemb) - 1
                    bAplicaGasto = False
                    RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'P'"
                    If RFiltro.RecordCount = 0 Then
                        bAplicaGasto = True
                    End If
                    If Not bAplicaGasto Then
                        RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & Mid(psCtaCod, 6, 3) & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "' AND cTpoFiltro = 'P'"
                        If RFiltro.RecordCount > 0 Then
                            bAplicaGasto = True
                        Else
                            bAplicaGasto = False
                        End If
                    End If
                    
                    '************************************************************
                    'Si tiene Filtro Por Tipos de Garantias
                    '************************************************************
                    If bAplicaGasto Then
                        bAplicaGasto = False
                        RFiltroGar.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'G'"
                        If RFiltroGar.RecordCount = 0 Then
                            bAplicaGasto = True
                        End If
                        If Not bAplicaGasto Then
                            RGar.MoveFirst
                            Do While Not RGar.EOF
                                If Not bAplicaGasto Then
                                    RFiltroGar.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & RGar!nTpoGar & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "' AND cTpoFiltro = 'G'"
                                    If RFiltroGar.RecordCount > 0 Then
                                        bAplicaGasto = True
                                        Exit Do
                                    Else
                                        bAplicaGasto = False
                                    End If
                                End If
                                RGar.MoveNext
                            Loop
                        End If
                    End If
                    
                    If bAplicaGasto Then
                        If CDbl(MatDesemb(i, 1)) >= R!nInicial And CDbl(MatDesemb(i, 1)) <= R!nFinal Then
                            'Si el gasto se aplica por relaciones de credito
                            If R!cFiltro = "I" Or R!cFiltro = "M" Then
                                If sTipoRelacGarantia = "I" Then 'Si el Credito es de Tipo Individual
                                    If R!cFiltro = "I" Then
                                        'Aplico a todos el mismo porcentaje
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                            
                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then
                                                        nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    End If
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "D" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "H" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        End If
                                                        If nRes < R!nMontoMin Then
                                                            nRes = R!nMontoMin
                                                        End If
                                                        If nRes > R!nMontoMax Then
                                                           nRes = R!nMontoMax
                                                        End If
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                            oCredRel.siguiente
                                        Loop
                                        MatDesemb(i, 2) = Format(nMontoTemp, "#0.000000")
                                    End If
                                Else 'Si es Mancomunada cobrar el porcentaje mancomunado al Titular y al Conyuge
                                     'y al resto el gasto de Individual
                                     'Para Titular y Conyuge
                                     
                                     If R!cFiltro = "M" Then
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            If oCredRel.ObtenerValorRelac = gColRelPersConyugue Or _
                                                oCredRel.ObtenerValorRelac = gColRelPersTitular Then
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                            
                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then
                                                        nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    End If
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "D" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "H" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        End If
                                                        If nRes < R!nMontoMin Then
                                                            nRes = R!nMontoMin
                                                        End If
                                                        If nRes > R!nMontoMax Then
                                                           nRes = R!nMontoMax
                                                        End If
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                            End If
                                            oCredRel.siguiente
                                        Loop
                                        MatDesemb(i, 2) = Format(nMontoTemp, "#0.000000")
                                        
                                     End If
                                     
                                     If R!cFiltro = "I" Then
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            If oCredRel.ObtenerValorRelac <> gColRelPersConyugue And _
                                                oCredRel.ObtenerValorRelac <> gColRelPersTitular Then
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                            
                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then
                                                        nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    End If
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "D" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "H" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        End If
                                                        If nRes < R!nMontoMin Then
                                                            nRes = R!nMontoMin
                                                        End If
                                                        If nRes > R!nMontoMax Then
                                                           nRes = R!nMontoMax
                                                        End If
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                            End If
                                            oCredRel.siguiente
                                        Loop
                                        MatDesemb(i, 2) = Format(nMontoTemp, "#0.000000")
                                     End If
                                End If
                                
                            Else 'Si no se considera las Relaciones entonces
                                'Filtramos por Edad
                                bCumpleRestricc = False
                                If Not IsNull(R!nEdadOper) Then
                                    Select Case R!nEdadOper
                                        Case 2 'Mayor
                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 3 'Mayor Igual
                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 4 'Menor
                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 5 'Menor Igual
                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 6 'Igual a
                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                    End Select
                                Else
                                    bCumpleRestricc = True
                                End If
                            
                                'Si es Fijo
                                If bCumpleRestricc Then
                                    If R!nTpoValor = 1 Then
                                        MatDesemb(i, 2) = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.00")
                                    End If
                                    'Si es Porcentaje
                                    If R!nTpoValor = 2 Then
                                        If R!cAplicaMonto = "C" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.00")
                                        End If
                                        If R!cAplicaMonto = "D" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.00")
                                        End If
                                        If R!cAplicaMonto = "H" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.00")
                                        End If
                                        If nRes < R!nMontoMin Then
                                            nRes = R!nMontoMin
                                        End If
                                        If nRes > R!nMontoMax Then
                                           nRes = R!nMontoMax
                                        End If
                                        MatDesemb(i, 2) = Format(nRes, "#0.000000")
                                    End If
                                End If
                            End If
                            
                            nNumGastos = nNumGastos + 1
                            MatGastos(nNumGastos - 1, 0) = "Desembolso" & Space(50) & gColocCalendAplDesembolso
                            MatGastos(nNumGastos - 1, 1) = i + 1
                            MatGastos(nNumGastos - 1, 2) = R!cDescripcion & Space(150) & R!nPrdConceptoCod
                            MatGastos(nNumGastos - 1, 3) = MatDesemb(i, 2)
                            If R!nAplicado = gColocConceptoAplDesembolso Then
                                Exit For
                            End If
                        End If
                    End If
                Next i
                R.MoveNext
            Loop
            R.Close
            Set R = Nothing
        End If
        
        '****************************************************
        'Aplicar Gastos a la Cuota
        '****************************************************
        Set oGastos = New DGastos
        Set R = oGastos.RecuperaGastosAplicablesCuotas(CInt(Mid(psCtaCod, 9, 1)), psAplicaproceso, sGastoFijoVariable)
        Set oGastos = Nothing
        'ReDim Preserve MatCalend(UBound(MatCalend), 7)
        Do While Not R.EOF
            
            For i = 0 To UBound(MatCalend) - 1
                bAplicaGasto = False
                    RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'P'"
                    If RFiltro.RecordCount = 0 Then
                        bAplicaGasto = True
                    End If
                    If Not bAplicaGasto Then
                        RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & Mid(psCtaCod, 6, 3) & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "' AND cTpoFiltro = 'P'"
                        
                        If RFiltro.RecordCount > 0 Then
                            bAplicaGasto = True
                        Else
                            bAplicaGasto = False
                        End If
                    End If
                                    
                    '************************************************************
                    'Si tiene Filtro Por Tipos de Garantias
                    '************************************************************
                    If bAplicaGasto Then
                        bAplicaGasto = False
                        RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'G'"
                        If RFiltro.RecordCount = 0 Then
                            bAplicaGasto = True
                        End If
                        
                        If Not bAplicaGasto Then
                            RGar.MoveFirst
                            Do While Not RGar.EOF
                                If Not bAplicaGasto Then
                                    RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & RGar!nTpoGar & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "' AND cTpoFiltro = 'G'"
                                    If RFiltro.RecordCount > 0 Then
                                        bAplicaGasto = True
                                        Exit Do
                                    Else
                                        bAplicaGasto = False
                                    End If
                                End If
                                RGar.MoveNext
                            Loop
                        End If
                    End If
                    
                    If bAplicaGasto Then
                        If CDbl(MatCalend(i, 2)) >= R!nInicial And CDbl(MatCalend(i, 2)) <= R!nFinal Then
                            'Si el gasto se aplica por relaciones de credito
                            If R!cFiltro = "I" Or R!cFiltro = "M" Then
                                If sTipoRelacGarantia = "I" Then 'Si el Credito es de Tipo Individual
                                    If R!cFiltro = "I" Then
                                        'Aplico a todos el mismo porcentaje
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                                
                                                If bCumpleRestricc Then
                                                'Si es Fijo
                                                    If R!nTpoValor = 1 Then
                                                        nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    End If
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 2)), "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "D" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "H" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "S" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                        End If
                                                        If nRes < R!nMontoMin Then
                                                            nRes = R!nMontoMin
                                                        End If
                                                        If nRes > R!nMontoMax Then
                                                           nRes = R!nMontoMax
                                                        End If
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                            oCredRel.siguiente
                                        Loop
                                        MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                    Else
                                        Exit For
                                    End If
                                Else 'Si es Mancomunada cobrar el porcentaje mancomunado todos menos al Titular
                                     'y al resto el gasto de Mancomunado
                                     'Para desiguales al Titular
                                     If R!cFiltro = "M" Then
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            If oCredRel.ObtenerValorRelac <> gColRelPersTitular Then
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                            
                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then
                                                        nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    End If
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "D" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "H" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "S" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                        End If
                                                        If nRes < R!nMontoMin Then
                                                            nRes = R!nMontoMin
                                                        End If
                                                        If nRes > R!nMontoMax Then
                                                           nRes = R!nMontoMax
                                                        End If
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                            End If
                                            oCredRel.siguiente
                                        Loop
                                        MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                        
                                     End If
                                     
                                     If R!cFiltro = "I" Then
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            If oCredRel.ObtenerValorRelac = gColRelPersTitular Then
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                            
                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then
                                                        nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    End If
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "D" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "H" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "S" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.0000000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                        End If
                                                        If nRes < R!nMontoMin Then
                                                            nRes = R!nMontoMin
                                                        End If
                                                        If nRes > R!nMontoMax Then
                                                           nRes = R!nMontoMax
                                                        End If
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                            End If
                                            oCredRel.siguiente
                                        Loop
                                        MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                     End If
                                
                                End If
                            Else    'Si no se considera las Relaciones entonces
                                'Filtramos por Edad
                                bCumpleRestricc = False
                                oCredRel.nPuntMat = oCredRel.PosicionTitular
                                If Not IsNull(R!nEdadOper) Then
                                    Select Case R!nEdadOper
                                        Case 2 'Mayor
                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 3 'Mayor Igual
                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 4 'Menor
                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 5 'Menor Igual
                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 6 'Igual a
                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                    End Select
                                Else
                                    bCumpleRestricc = True
                                End If
                            
                                'Si es Fijo
                                If bCumpleRestricc Then
                                    If R!nTpoValor = 1 Then
                                        MatCalend(i, 6) = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.00")
                                    End If
                                    'Si es Porcentaje
                                    If R!nTpoValor = 2 Then
                                        If R!cAplicaMonto = "C" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                        End If
                                        If R!cAplicaMonto = "D" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                        End If
                                        If R!cAplicaMonto = "H" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                        End If
                                        If R!cAplicaMonto = "S" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                        End If
                                        If nRes < R!nMontoMin Then
                                            nRes = R!nMontoMin
                                        End If
                                        If nRes > R!nMontoMax Then
                                           nRes = R!nMontoMax
                                        End If
                                        MatCalend(i, 6) = Format(nRes, "#0.00")
                                    End If
                                End If
                            End If
                                                        
                            If CDbl(MatCalend(i, 6)) > 0 And bCumpleRestricc Then
                                nNumGastos = nNumGastos + 1
                                MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                MatGastos(nNumGastos - 1, 1) = i + 1
                                MatGastos(nNumGastos - 1, 2) = R!cDescripcion & Space(150) & R!nPrdConceptoCod
                                MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                                If R!nAplicado = gColocConceptoAplCuota Then
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        End If
                    End If
            Next i
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
    End If
        
        
        
    'Para los Gastos que se cargan en los Cancelaciones y Pagos
    If psAplicaproceso = "PA" Or psAplicaproceso = "CA" Or psAplicaproceso = "PP" Then
        
        Set oCred = New DCredito
        Set RGastosExon = oCred.RecuperaGastosExonerados(psCtaCod)
        Set oCred = Nothing
        ReDim MatGastosExon(0)
        Do While Not RGastosExon.EOF
            ReDim Preserve MatGastosExon(UBound(MatGastosExon) + 1)
            MatGastosExon(UBound(MatGastosExon) - 1) = Str(Trim(RGastosExon!nPrdConceptoCod))
            RGastosExon.MoveNext
        Loop
        
        nNumGastos = 0
        ReDim MatGastos(1000, 5)
        Set oGastos = New DGastos
        Set R = oGastos.RecuperaGastosAplicablesCuotas(CInt(Mid(psCtaCod, 9, 1)), psAplicaproceso, sGastoFijoVariable, , MatGastosExon)
        
        Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P")
        Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G")
        Set oGastos = Nothing
        
        Do While Not R.EOF
            bAplicaGasto = False
            RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod
            If RFiltro.RecordCount = 0 Then
                bAplicaGasto = True
            End If
            If Not bAplicaGasto Then
                RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & Mid(psCtaCod, 6, 3) & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "' AND cTpoFiltro = 'P'"
                If RFiltro.RecordCount > 0 Then
                    bAplicaGasto = True
                Else
                    bAplicaGasto = False
                End If
            End If
                            
            '************************************************************
            'Si tiene Filtro Por Tipos de Garantias
            '************************************************************
            If bAplicaGasto Then
                bAplicaGasto = False
                RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'G'"
                If RFiltro.RecordCount = 0 Then
                    bAplicaGasto = True
                End If
                
                If Not bAplicaGasto Then
                    RGar.MoveFirst
                    Do While Not RGar.EOF
                        If Not bAplicaGasto Then
                            RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & RGar!nTpoGar & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "' AND cTpoFiltro = 'G'"
                            If RFiltro.RecordCount > 0 Then
                                bAplicaGasto = True
                                Exit Do
                            Else
                                bAplicaGasto = False
                            End If
                        End If
                        RGar.MoveNext
                    Loop
                End If
            End If
            If bAplicaGasto Then
                Set oCred = New DCredito
                Set RCredVig = oCred.RecuperaDatosCreditoVigente(psCtaCod)
                Set oCred = Nothing
                
                'Si se encuentra en el Rango Permitido
                If nPrestamo >= R!nInicial And nPrestamo <= R!nFinal Then
                    bCumpleRestricc = False
                    'Si esta sujeto al vencimiento de dias de atraso
                    If Not IsNull(R!nDiasApl) Then
                       If pnDiasAtraso = R!nDiasApl Then
                            bCumpleRestricc = True
                       Else
                            bCumpleRestricc = False
                       End If
                    Else
                        bCumpleRestricc = True
                    End If
                    
                    
                    If bCumpleRestricc Then
                        'Si esta sujeto a restriccion de Operacion
                        bCumpleRestricc = False
                        If Not IsNull(R!nOperador) Then
                            Select Case R!cOperMonto
                                Case "P" 'Prestamo
                                    Select Case R!nEdadOper
                                        Case 2 'Mayor
                                            If nPrestamo > CDbl(Format((R!nOperPorc / 100) * RCredVig!nMontoCol)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 3 'Mayor Igual
                                            If nPrestamo >= CDbl(Format((R!nOperPorc / 100) * RCredVig!nMontoCol)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 4 'Menor
                                            If nPrestamo < CDbl(Format((R!nOperPorc / 100) * RCredVig!nMontoCol)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 5 'Menor Igual
                                            If nPrestamo <= CDbl(Format((R!nOperPorc / 100) * RCredVig!nMontoCol)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 6 'Igual a
                                            If nPrestamo = CDbl(Format((R!nOperPorc / 100) * RCredVig!nMontoCol)) Then
                                                bCumpleRestricc = True
                                            End If
                                    End Select
                                Case "C" 'Cuota
                                    Select Case R!nEdadOper
                                        Case 2 'Mayor
                                            If nPrestamo > CDbl(Format((R!nOperPorc / 100) * nCuotaApr)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 3 'Mayor Igual
                                            If nPrestamo >= CDbl(Format((R!nOperPorc / 100) * nCuotaApr)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 4 'Menor
                                            If nPrestamo < CDbl(Format((R!nOperPorc / 100) * nCuotaApr)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 5 'Menor Igual
                                            If nPrestamo <= CDbl(Format((R!nOperPorc / 100) * nCuotaApr)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 6 'Igual a
                                            If nPrestamo = CDbl(Format((R!nOperPorc / 100) * nCuotaApr)) Then
                                                bCumpleRestricc = True
                                            End If
                                    End Select
                                Case "G" 'Garantia
                                    Select Case R!nEdadOper
                                        Case 2 'Mayor
                                            If nPrestamo > CDbl(Format((R!nOperPorc / 100) * nMontoGarantia)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 3 'Mayor Igual
                                            If nPrestamo >= CDbl(Format((R!nOperPorc / 100) * nMontoGarantia)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 4 'Menor
                                            If nPrestamo < CDbl(Format((R!nOperPorc / 100) * nMontoGarantia)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 5 'Menor Igual
                                            If nPrestamo <= CDbl(Format((R!nOperPorc / 100) * nMontoGarantia)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 6 'Igual a
                                            If nPrestamo = CDbl(Format((R!nOperPorc / 100) * nMontoGarantia)) Then
                                                bCumpleRestricc = True
                                            End If
                                    End Select
                                Case "S" 'Saldo
                                    Select Case R!nEdadOper
                                        Case 2 'Mayor
                                            If nPrestamo > CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 3 'Mayor Igual
                                            If nPrestamo >= CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 4 'Menor
                                            If nPrestamo < CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 5 'Menor Igual
                                            If nPrestamo <= CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 6 'Igual a
                                            If nPrestamo = CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then
                                                bCumpleRestricc = True
                                            End If
                                    End Select
                                Case "H" 'Hipoteca
                                    Set oCred = New DCredito
                                    MontoPrePagado = oCred.RecuperaDatosPrepago(psCtaCod, pdHoy)
                                    Set oCred = Nothing
                                    Select Case R!nOperador
                                        Case 2 'Mayor
                                            If pnMontoCapital + MontoPrePagado > CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 3 'Mayor Igual
                                            If pnMontoCapital + MontoPrePagado >= CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 4 'Menor
                                            If pnMontoCapital + MontoPrePagado < CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 5 'Menor Igual
                                            If pnMontoCapital + MontoPrePagado <= CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 6 'Igual a
                                            If pnMontoCapital + MontoPrePagado = CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then
                                                bCumpleRestricc = True
                                            End If
                                    End Select
                            End Select
                        Else
                            bCumpleRestricc = True
                        End If
                    End If
                    
                    'Si cumple la primera Restriccion
                    'Buecar Por Restriccion por edad del Titular
                    If bCumpleRestricc Then
                        bCumpleRestricc = False
                        oCredRel.nPuntMat = oCredRel.PosicionTitular
                        If Not IsNull(R!nEdadOper) Then
                            Select Case R!nEdadOper
                                Case 2 'Mayor
                                    If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                        bCumpleRestricc = True
                                    End If
                                Case 3 'Mayor Igual
                                    If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                        bCumpleRestricc = True
                                    End If
                                Case 4 'Menor
                                    If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                        bCumpleRestricc = True
                                    End If
                                Case 5 'Menor Igual
                                    If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                        bCumpleRestricc = True
                                    End If
                                Case 6 'Igual a
                                    If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                        bCumpleRestricc = True
                                    End If
                            End Select
                        Else
                            bCumpleRestricc = True
                        End If
                        
                        If bCumpleRestricc Then
                            If Not IsNull(R!nOperDiasVenc) Then
                                bCumpleRestricc = False
                                nDiasAtraso = pnDiasAtraso
                                Select Case R!nOperDiasVenc
                                    Case 2 'Mayor
                                        If nDiasAtraso > R!nDiasVenc Then
                                            bCumpleRestricc = True
                                        End If
                                    Case 3 'Mayor Igual
                                        If nDiasAtraso >= R!nDiasVenc Then
                                            bCumpleRestricc = True
                                        End If
                                    Case 4 'Menor
                                        If nDiasAtraso < R!nDiasVenc Then
                                            bCumpleRestricc = True
                                        End If
                                    Case 5 'Menor Igual
                                        If nDiasAtraso <= R!nDiasVenc Then
                                            bCumpleRestricc = True
                                        End If
                                    Case 6 'Igual a
                                        If nDiasAtraso = R!nDiasVenc Then
                                            bCumpleRestricc = True
                                        End If
                                End Select
                                
                            End If
                        End If
                        
                        nMontoTemp = 0
                        'Si cumple la restriccion
                        If bCumpleRestricc Then
                            If R!nTpoValor = 1 Then
                                nMontoTemp = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                If nMontoTemp > 0 Then
                                    nNumGastos = nNumGastos + 1
                                    MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                    MatGastos(nNumGastos - 1, 1) = pnCuotaPend
                                    MatGastos(nNumGastos - 1, 2) = "GASTO POR CANCELACION, PAGO O PREPAGO " & Space(150) & R!nPrdConceptoCod
                                    MatGastos(nNumGastos - 1, 3) = Format(nMontoTemp, "#0.00")
                                    MatGastos(nNumGastos - 1, 4) = Trim(R!cAplicaProceso)
                                End If
                            End If
                            'Si es Porcentaje
                            If R!nTpoValor = 2 Then
                                If Not IsNull(R!cOperMonto) Then
                                    If R!cOperMonto = "H" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * ((pnMontoCapital + MontoPrePagado) - CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo))), "#0.000000")
                                    Else
                                        If R!cAplicaMonto = "C" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                        End If
                                        If R!cAplicaMonto = "D" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(0, 0)), "#0.000000")
                                        End If
                                        If R!cAplicaMonto = "H" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                        End If
                                        If R!cAplicaMonto = "S" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * RCredVig!nSaldo, "#0.000000")
                                        End If
                                        If nRes < R!nMontoMin Then
                                            nRes = R!nMontoMin
                                        End If
                                        If nRes > R!nMontoMax Then
                                           nRes = R!nMontoMax
                                        End If
                                    End If
                                Else
                                    If R!cAplicaMonto = "C" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "D" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(0, 0)), "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "H" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "S" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * RCredVig!nSaldo, "#0.000000")
                                    End If
                                    If nRes < R!nMontoMin Then
                                        nRes = R!nMontoMin
                                    End If
                                    If nRes > R!nMontoMax Then
                                       nRes = R!nMontoMax
                                    End If
                                End If
                                nMontoTemp = Format(nRes, "#0.00")
                                
                                nNumGastos = nNumGastos + 1
                                MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                MatGastos(nNumGastos - 1, 1) = pnCuotaPend
                                MatGastos(nNumGastos - 1, 2) = "GASTO POR CANCELACION, PAGO O PREPAGO " & Space(150) & R!nPrdConceptoCod
                                MatGastos(nNumGastos - 1, 3) = Format(nMontoTemp, "#0.00")
                                MatGastos(nNumGastos - 1, 4) = Trim(R!cAplicaProceso)
                            End If
                        End If
                    End If
                End If
            End If
            
            R.MoveNext
        Loop
    End If
    
    'Para Cierre de Dia
    If psAplicaproceso = "CD" Then
        nNumGastos = 0
        ReDim MatGastos(1000, 4)
        If Not pHabRecorsed Then
            Set oGastos = New DGastos
            Set R = oGastos.RecuperaGastosAplicablesCuotas(CInt(Mid(psCtaCod, 9, 1)), psAplicaproceso, , True)
            Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P")
            Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G")
            Set oGastos = Nothing
        Else
            Set R = PR
            Set RFiltro = pRFiltro
            Set RFiltroGar = pRFiltroGar
            If R.RecordCount > 0 Then
                R.MoveFirst
            End If
            If RFiltro.RecordCount > 0 Then
                RFiltro.MoveFirst
            End If
            If RFiltroGar.RecordCount > 0 Then
                RFiltroGar.MoveFirst
            End If
        End If
        
        Do While Not R.EOF
            For i = 0 To UBound(MatCalend) - 1
                bAplicaGasto = False
                RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod
                If RFiltro.RecordCount = 0 Then
                    bAplicaGasto = True
                End If
                If Not bAplicaGasto Then
                    RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & Mid(psCtaCod, 6, 3) & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "' AND cTpoFiltro = 'P'"
                    If RFiltro.RecordCount > 0 Then
                        bAplicaGasto = True
                    Else
                        bAplicaGasto = False
                    End If
                End If
                                
                '************************************************************
                'Si tiene Filtro Por Tipos de Garantias
                '************************************************************
                If bAplicaGasto Then
                    bAplicaGasto = False
                    RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'G'"
                    If RFiltro.RecordCount = 0 Then
                        bAplicaGasto = True
                    End If
                    
                    If Not bAplicaGasto Then
                        RGar.MoveFirst
                        Do While Not RGar.EOF
                            If Not bAplicaGasto Then
                                RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & RGar!nTpoGar & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "' AND cTpoFiltro = 'G'"
                                If RFiltro.RecordCount > 0 Then
                                    bAplicaGasto = True
                                    Exit Do
                                Else
                                    bAplicaGasto = False
                                End If
                            End If
                            RGar.MoveNext
                        Loop
                    End If
                End If
                
                If bAplicaGasto Then
                
                    If nPrestamo >= R!nInicial And nPrestamo <= R!nFinal Then
                        If Not IsNull(R!nDiasApl) Then
                            nDiasAtraso = pdHoy - CDate(MatCalend(i, 0))
                            If R!nDiasApl = nDiasAtraso Then
                                If R!nTpoValor = 1 Then
                                    nMontoTemp = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                    MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                End If
                                'Si es Porcentaje
                                If R!nTpoValor = 2 Then
                                    If R!cAplicaMonto = "C" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 2)), "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "D" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "H" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "S" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 3)), "#0.000000")
                                    End If
                                    If nRes < R!nMontoMin Then
                                        nRes = R!nMontoMin
                                    End If
                                    If nRes > R!nMontoMax Then
                                       nRes = R!nMontoMax
                                    End If
                                    MatCalend(i, 6) = Format(nRes, "#0.00")
                                End If
                                If CDbl(MatCalend(i, 6)) > 0 Then
                                        nNumGastos = nNumGastos + 1
                                        MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                        MatGastos(nNumGastos - 1, 1) = i + 1
                                        MatGastos(nNumGastos - 1, 2) = R!cDescripcion & Space(150) & R!nPrdConceptoCod
                                        MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                                End If
                            End If
                            
                        End If
                        
                        'Para el otro gasto de Cierre de Dia
                        If Not IsNull(R!nOperDiasVenc) Then
                            bCumpleRestricc = False
                            nDiasAtraso = pdHoy - CDate(MatCalend(i, 0))
                            Select Case R!nOperDiasVenc
                                Case 2 'Mayor
                                    If nDiasAtraso > R!nDiasVenc Then
                                        bCumpleRestricc = True
                                    End If
                                Case 3 'Mayor Igual
                                    If nDiasAtraso >= R!nDiasVenc Then
                                        bCumpleRestricc = True
                                    End If
                                Case 4 'Menor
                                    If nDiasAtraso < R!nDiasVenc Then
                                        bCumpleRestricc = True
                                    End If
                                Case 5 'Menor Igual
                                    If nDiasAtraso <= R!nDiasVenc Then
                                        bCumpleRestricc = True
                                    End If
                                Case 6 'Igual a
                                    If nDiasAtraso = R!nDiasVenc Then
                                        bCumpleRestricc = True
                                    End If
                            End Select
                            
                            If bCumpleRestricc Then
                                If R!nTpoValor = 1 Then
                                    nMontoTemp = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                    nRes = nMontoTemp
                                    MatCalend(i, 6) = Format(nRes, "#0.00")
                                End If
                                'Si es Porcentaje
                                If R!nTpoValor = 2 Then
                                    If R!cAplicaMonto = "C" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 2)), "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "D" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "H" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "S" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 3)), "#0.000000")
                                    End If
                                    If nRes < R!nMontoMin Then
                                        nRes = R!nMontoMin
                                    End If
                                    If nRes > R!nMontoMax Then
                                       nRes = R!nMontoMax
                                    End If
                                    MatCalend(i, 6) = Format(nRes, "#0.00")
                                End If
                                If CDbl(MatCalend(i, 6)) > 0 Then
                                        nNumGastos = nNumGastos + 1
                                        MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                        MatGastos(nNumGastos - 1, 1) = i + 1
                                        MatGastos(nNumGastos - 1, 2) = R!cDescripcion & Space(150) & R!nPrdConceptoCod
                                        MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
            R.MoveNext
        Loop
        
    End If
    
    'Para los Gastos que se cargan en las Refinanciaciones
    'Igual que los desemboloso pero para la parte de las cuotas
    If psAplicaproceso = "RE" Then
    
        Set oGastos = New DGastos
        Set R = oGastos.RecuperaGastosAplicablesCuotas(CInt(Mid(psCtaCod, 9, 1)))
        Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P")
        Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G")
        Set oGastos = Nothing
        'ReDim Preserve MatCalend(UBound(MatCalend), 7)
        Do While Not R.EOF
            
            For i = 0 To UBound(MatCalend) - 1
                bAplicaGasto = False
                    RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod
                    If RFiltro.RecordCount = 0 Then
                        bAplicaGasto = True
                    End If
                    If Not bAplicaGasto Then
                        RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And cProdCod = '" & Mid(psCtaCod, 6, 3) & "' And nMoneda = " & Mid(psCtaCod, 9, 1) & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "'"
                        If RFiltro.RecordCount > 0 Then
                            bAplicaGasto = True
                        Else
                            bAplicaGasto = False
                        End If
                    End If
                                    
                    '************************************************************
                    'Si tiene Filtro Por Tipos de Garantias
                    '************************************************************
                    If bAplicaGasto Then
                        RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'G'"
                        If RFiltro.RecordCount = 0 Then
                            bAplicaGasto = True
                        End If
                        If Not bAplicaGasto Then
                            RGar.MoveFirst
                            Do While Not RGar.EOF
                                If Not bAplicaGasto Then
                                    RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & RGar!nTpoGar & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "' AND cTpoFiltro = 'G'"
                                    If RFiltro.RecordCount > 0 Then
                                        bAplicaGasto = True
                                        Exit Do
                                    Else
                                        bAplicaGasto = False
                                    End If
                                End If
                                RGar.MoveNext
                            Loop
                        End If
                    End If
                    If bAplicaGasto Then
                        If CDbl(MatCalend(i, 2)) >= R!nInicial And CDbl(MatCalend(i, 2)) <= R!nFinal Then
                            'Si el gasto se aplica por relaciones de credito
                            If R!cFiltro = "I" Or R!cFiltro = "M" Then
                                If sTipoRelacGarantia = "I" Then 'Si el Credito es de Tipo Individual
                                    If R!cFiltro = "I" Then
                                        'Aplico a todos el mismo porcentaje
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                                
                                                If bCumpleRestricc Then
                                                'Si es Fijo
                                                    If R!nTpoValor = 1 Then
                                                        nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    End If
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 2)), "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "D" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "H" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "S" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                        End If
                                                        If nRes < R!nMontoMin Then
                                                            nRes = R!nMontoMin
                                                        End If
                                                        If nRes > R!nMontoMax Then
                                                           nRes = R!nMontoMax
                                                        End If
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                            oCredRel.siguiente
                                        Loop
                                        MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                    End If
                                Else 'Si es Mancomunada cobrar el porcentaje mancomunado al Titular y al Conyuge
                                     'y al resto el gasto de Individual
                                     'Para Titular y Conyuge
                                     If R!cFiltro = "M" Then
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            If oCredRel.ObtenerValorRelac = gColRelPersConyugue Or _
                                                oCredRel.ObtenerValorRelac = gColRelPersTitular Then
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                            
                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then
                                                        nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    End If
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "D" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "H" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        End If
                                                        If nRes < R!nMontoMin Then
                                                            nRes = R!nMontoMin
                                                        End If
                                                        If nRes > R!nMontoMax Then
                                                           nRes = R!nMontoMax
                                                        End If
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                            End If
                                            oCredRel.siguiente
                                        Loop
                                        MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                        
                                     End If
                                     
                                     If R!cFiltro = "I" Then
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            If oCredRel.ObtenerValorRelac <> gColRelPersConyugue And _
                                                oCredRel.ObtenerValorRelac <> gColRelPersTitular Then
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                                bCumpleRestricc = True
                                                            End If
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                            
                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then
                                                        nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    End If
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "D" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.000000")
                                                        End If
                                                        If R!cAplicaMonto = "H" Then
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        End If
                                                        If nRes < R!nMontoMin Then
                                                            nRes = R!nMontoMin
                                                        End If
                                                        If nRes > R!nMontoMax Then
                                                           nRes = R!nMontoMax
                                                        End If
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                            End If
                                            oCredRel.siguiente
                                        Loop
                                        MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                     End If
                                
                                End If
                            Else    'Si no se considera las Relaciones entonces
                                'Filtramos por Edad
                                bCumpleRestricc = False
                                oCredRel.nPuntMat = oCredRel.PosicionTitular
                                If Not IsNull(R!nEdadOper) Then
                                    Select Case R!nEdadOper
                                        Case 2 'Mayor
                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 3 'Mayor Igual
                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 4 'Menor
                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 5 'Menor Igual
                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                        Case 6 'Igual a
                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                bCumpleRestricc = True
                                            End If
                                    End Select
                                Else
                                    bCumpleRestricc = True
                                End If
                            
                                'Si es Fijo
                                If bCumpleRestricc Then
                                    If R!nTpoValor = 1 Then
                                        MatCalend(i, 6) = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                    End If
                                    'Si es Porcentaje
                                    If R!nTpoValor = 2 Then
                                        If R!cAplicaMonto = "C" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                        End If
                                        If R!cAplicaMonto = "D" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.000000")
                                        End If
                                        If R!cAplicaMonto = "H" Then
                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                        End If
                                        If nRes < R!nMontoMin Then
                                            nRes = R!nMontoMin
                                        End If
                                        If nRes > R!nMontoMax Then
                                           nRes = R!nMontoMax
                                        End If
                                        MatCalend(i, 6) = Format(nRes, "#0.00")
                                    End If
                                End If
                            End If
                            
                            nNumGastos = nNumGastos + 1
                            MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                            MatGastos(nNumGastos - 1, 1) = i + 1
                            MatGastos(nNumGastos - 1, 2) = R!cDescripcion & Space(150) & R!nPrdConceptoCod
                            MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                            If R!nAplicado = gColocConceptoAplCuota Then
                                Exit For
                            End If
                        End If
                    End If
            Next i
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
    End If
                    
    Set oCredRel = Nothing
    
    GeneraCalendarioGastos = MatGastos
    Exit Function

ErrorGeneraCalendarioGastos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


