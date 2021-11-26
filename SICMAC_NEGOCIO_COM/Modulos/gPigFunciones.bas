Attribute VB_Name = "gPigFunciones"
Public Function fgIniciaAxCuentaPignoraticio1() As String
    fgIniciaAxCuentaPignoraticio1 = gsCodCMAC & "07" & "305" 'gsCodAge ' CMCPL
End Function

'Permite Muestra el Credito Pignoraticio en el AXDesCon
Public Function fgMuestraPig_AXPigDesCon(ByVal psNroContrato As String, pAXPigDesCon As ActXPigDesCon, Optional pbHabilitaDescrLote As Boolean = False) As Boolean
'On Error GoTo ControlError

Dim lrCredPig As ADODB.Recordset
Dim lrCredPigJoyas As ADODB.Recordset
Dim lrCredPigPersonas As ADODB.Recordset
Dim lrComision As ADODB.Recordset ' CMCPL cambio para calendario
Dim loMuestraContrato As dPigContrato
Dim lrCalendarioCon As ADODB.Recordset ' CMCPL cambio para calendario

fgMuestraPig_AXPigDesCon = True
Set loMuestraContrato = New dPigContrato
Set lrCredPig = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psNroContrato)
Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psNroContrato)
Set lrCalendarioCon = loMuestraContrato.dObtieneCalendario(psNroContrato)


Set loMuestraContrato = Nothing
        
    If lrCredPig.BOF And lrCredPig.EOF Then
        lrCredPig.Close
        Set lrCredPig = Nothing
        Set lrCredPigJoyas = Nothing
        Set lrCredPigPersonas = Nothing
        Set lrCalendarioCon = Nothing
        MsgBox " No se encuentra el Credito Pignoraticio " & psNroContrato, vbInformation, " Aviso "
        fgMuestraPig_AXPigDesCon = False
        Exit Function
    Else
        'pAXPigDesCon.Limpiar
        pAXPigDesCon.prestamo1 = lrCalendarioCon!PigCapital ' CAMBIO CMCPL
        pAXPigDesCon.comision1 = lrCalendarioCon!Comisiones 'CAMBIO CMCPL
        pAXPigDesCon.neto1 = pAXPigDesCon.prestamo1 - pAXPigDesCon.comision1
        '  pAXPigDesCon.Piezas = lrCredPig!npiezas
        pAXPigDesCon.Piezas = lrCredPig!nPlazo 'cambiar ahora esta el plazo antes estaba la pieza
       
        pAXPigDesCon.FechaPrestamo = Format(lrCredPig!dVigencia, "dd/mm/yyyy")
        pAXPigDesCon.FechaVencimiento = Format(lrCredPig!dVenc, "dd/mm/yyyy")
     
        lrCredPig.Close
        Set lrCredPig = Nothing

      ' Mostrar Clientes
        If fgMostrarClientes(pAXPigDesCon.listaClientes, lrCredPigPersonas) = False Then
            MsgBox " No se encuentra Datos de Clientes de Contrato " & psNroContrato, vbInformation, " Aviso "
            fgMuestraPig_AXPigDesCon = False
            Exit Function
        End If
        
        lrCredPigPersonas.Close
        Set lrCredPigPersonas = Nothing
        
        If pbHabilitaDescrLote = True Then
            pAXPigDesCon.EnabledDescLot = True
            pAXPigDesCon.SetFocusDesLot
        Else
            pAXPigDesCon.EnabledDescLot = False
        End If
   End If
Exit Function

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Function

'**************************************************************
' CRSF -5 / 7 / 2002 FUNCION DE COBRO DE COMISION DE CUSTODIA
'**************************************************************
Public Function fgMuestraPig_AXPigComision(ByVal psNroContrato As String, pAXPigComision As ActXPigComision, Optional pbHabilitaDescrLote As Boolean = False) As Boolean

Dim lrCredPig As ADODB.Recordset
Dim lrCredPigJoyas As ADODB.Recordset
Dim lrCredPigPersonas As ADODB.Recordset
Dim lrComision As ADODB.Recordset ' CMCPL cambio para calendario
Dim loMuestraContrato  As dPigContrato
Dim livalorfecha As Integer

fgMuestraPig_AXPigComision = True
Set loMuestraContrato = New dPigContrato


Set lrCredPig = loMuestraContrato.dObtieneCreditoPigno(psNroContrato)
Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psNroContrato)
Set lrComision = loMuestraContrato.dComision(psNroContrato, gPigEstCancelPendRes)

Set loMuestraContrato = Nothing

        pAXPigComision.Fechapago = lrComision!dPrdEstado ' CAMBIO CMCPL
        livalorfecha = DateDiff("d", pAXPigComision.Fechapago, Date)
        pAXPigComision.DiasAtraso = CInt(livalorfecha)
        pAXPigComision.CuotaCosto = CInt(calculo(livalorfecha, gColPigConceptoCodCustodiaDif, gPigParamDiasMinComision)) 'llamado a la funcion CALCULO
        lrCredPig.Close
        Set lrCredPig = Nothing
      ' Mostrar Clientes
        If fgMostrarClientes(pAXPigComision.listaClientes, lrCredPigPersonas) = False Then
            MsgBox " No se encuentra Datos de Clientes de Contrato " & psNroContrato, vbInformation, " Aviso "
            fgMuestraPig_AXPigComision = False
            Exit Function
        End If
        lrCredPigPersonas.Close
        Set lrCredPigPersonas = Nothing
 Exit Function

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Function
Public Function calculo(ByVal pnDiasAtraso As Integer, ByVal pnComision As Integer, _
                                        ByVal pnDiasCobro As Integer) As Currency
                                        
    Dim lnMesAtrasado As Integer
    Dim lnDiasAtrasado  As Integer
    Dim lnTotalPagar   As Currency
    Dim lnDiasMinimoCobro  As Integer
    Dim lnComisionPagar As Currency
    Dim rsCom As Recordset
    Dim oFunc As DPigFunciones
        
    Set oFunc = New DPigFunciones
    lnDiasMinimoCobro = oFunc.GetParamValor(pnDiasCobro)
    Set rsCom = oFunc.GetConceptoValor(pnComision)
    
    lnComisionPagar = rsCom("nValor")
    Set rsCom = Nothing
   
    If pnDiasAtraso > lnDiasMinimoCobro Then
        lnMesAtrasado = pnDiasAtraso \ lnDiasMinimoCobro
        lnDiasAtrasado = pnDiasAtraso Mod lnDiasMinimoCobro
        
        If lnMesAtrasado > 0 Then 'mes de atraso
            If lnDiasAtrasado > 0 Then
                 lnTotalPagar = lnComisionPagar * (lnMesAtrasado + 1)  'Aca verificar esto
            Else
                lnTotalPagar = lnComisionPagar * (lnMesAtrasado)
            End If
        Else
            lnTotalPagar = 0
        End If
    End If
    
    calculo = lnTotalPagar
    
End Function

Public Function fgEstadoVigenteCredito(ByVal pnEstado As Integer) As Boolean
Dim lbVigenteOK As Boolean
lbVigenteOK = False
    Select Case pnEstado
        Case gColocEstVigNorm, gColocEstVigMor, gColocEstVigVenc
            lbVigenteOK = True
        Case gColocEstRefNorm, gColocEstRefMor, gColocEstRefVenc
            lbVigenteOK = True
        Case gColocEstRecVigJud, gColocEstRecVigCast
            lbVigenteOK = True
        Case gColPEstRegis, gColPEstVenci, gColPEstRenov, gColPEstPRema
            lbVigenteOK = True
    End Select
    fgEstadoVigenteCredito = lbVigenteOK
End Function




Public Function fgProductoCreditoTipo(ByVal psCuenta As String) As String
Dim lsDesc As String
    Select Case Mid(psCuenta, 6, 3)
        Case "101"
            lsDesc = "COMERCIAL"
        Case "201"
            lsDesc = "PYME"
        Case "301"
            lsDesc = "CONSUMO DSCTO PLANILLA"
        Case "302"
            lsDesc = "CONSUMO AVAL PF"
        Case "303"
            lsDesc = "CONSUMO AVAL CTS"
        Case "304"
            lsDesc = "CONSUMO USOS DIVERSOS"
        Case "305"
            lsDesc = "PIGNORATICIO"
        Case "320"
            lsDesc = "CONSUMO ADMINISTRATIVO"
        Case "401"
            lsDesc = "HIPOTECARIO "
        Case "423"
            lsDesc = "MI VIVIENDA"
        Case "121"
            lsDesc = "CARTA FIANZA"
        Case "221"
            lsDesc = "CARTA FIANZA"
          
    End Select
    fgProductoCreditoTipo = lsDesc
End Function
