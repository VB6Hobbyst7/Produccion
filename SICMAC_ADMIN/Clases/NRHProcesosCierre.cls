VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NRHProcesosCierre"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Actualiza Datos de la Relacion concepto Empleado
Option Base 0
Option Explicit

Public Event ShowProgress()
Public Event CloseProgress()
Public Event Progress(pnValor As Long, pnTotal As Long)

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB9030101E8
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB90301035B
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB90301031F
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    'fgITFParametros

    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'Agrega una nueva relación entre concepto y Empleado
'##ModelId=3A9BEAB1011F
Public Function CierreDia(pdFecha As Date, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo AgregaConceptoEmpleadoErr
    Dim oPla As DRHProcesosCierre
    Set oPla = New DRHProcesosCierre
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Dim lsCadena As String
    Dim lnItem As Long
    Dim lnPagina As Long
    Dim lsCodEmp As String * 8
    Dim lsNombre As String * 40
    Dim lsEstadoAnt As String * 30
    Dim lsEstadoNue As String * 30
    
    Set rsP = oPla.CierreDia(pdFecha)
     
    
    lsCadena = ""
    RaiseEvent ShowProgress
    If Not (rsP.EOF And rsP.BOF) Then
        rsP.MoveFirst
        lsCadena = lsCadena & CabeceraPagina("CAMBIO DE ESTADOS DEL RECURSO HUMANO", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Codigo;6; ;2;Nombre;20; ;20;Est. Anterior;25; ;10;Est. Nuevo;25; ;10;", lnItem)
        While Not rsP.EOF
            lsCodEmp = rsP!EmpCod
            lsNombre = PstaNombre(rsP!Nombre)
            lsEstadoAnt = rsP!DesAnt
            lsEstadoNue = rsP!DesNue
            
            lnItem = lnItem + 1
            lsCadena = lsCadena & lsCodEmp & "  " & lsNombre & "  " & lsEstadoAnt & "  " & lsEstadoNue & oImpresora.gPrnSaltoLinea
            
            If lnItem > 56 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("CAMBIO DE ESTADOS DEL RECURSO HUMANO", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Codigo;6; ;2;Nombre;20; ;20;Est. Anterior;25; ;10;Est. Nuevo;25; ;10;", lnItem)
            End If
            
            RaiseEvent Progress(rsP.Bookmark, rsP.RecordCount)
            rsP.MoveNext
        Wend
    End If
    RaiseEvent CloseProgress
    Set oPla = Nothing
    rsP.Close
    Set rsP = Nothing
    
    CierreDia = lsCadena
    
    Exit Function
AgregaConceptoEmpleadoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoEmpleado:AgregaConceptoEmpleado Method")
End Function

'Agrega una nueva relación entre concepto y Empleado
'##ModelId=3A9BEAB1011F
Public Function AbonoPagos(prRS As ADODB.Recordset, psPlanillaPeriodo As String, psPlanillaCod As String, psPlanillaDes As String, pnTipoCambioVenta As Currency, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, pgsCodAge As String, pgsCodUser As String, pSNomPlanilla As String) As String
    'pSNomPlanilla 'APRI20170328
    On Error GoTo AgregaConceptoEmpleadoErr
        
    Dim lsNotaAbono As String
    Dim lsCuentaRef As String
    Dim lsNotaAbonoRef As String
    Dim lsMonto As String
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lsCodCta As String
    Dim oPla As DRHProcesosCierre
    Set oPla = New DRHProcesosCierre
    Dim oExtra As DActualizaMovExtraPlanilla
    Set oExtra = New DActualizaMovExtraPlanilla
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    Dim rsOpe As ADODB.Recordset
    Set rsOpe = New ADODB.Recordset
    Dim oMov As DMov
    Set oMov = New DMov
    'Dim oPla As DRHProcesosCierre
    'Set oPla = New DRHProcesosCierre
    Dim oCapta As NCapAbonoCargo
    Set oCapta = New NCapAbonoCargo
    Dim lnSaldo As Double
    Dim lnTipCambioVenta As Currency
    Dim lsCadBol As String
    Dim lsCadBolInd As String 'APRI20270328
    Dim Planilla As String 'APRI20270328
    Dim lsNum As String
    Dim lsOpeCod As String
    
    Dim lnITFValor As Currency
    
    gsNomAge = Trim(pgsNomAge)
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim oConRef As DConecta
    Set oConRef = New DConecta
    
    Dim lsOpeAbonoCod As String
    
    Dim lbAbonado As Boolean
    
    lsOpeAbonoCod = GetOpeAbonoPlanilla(psPlanillaCod)
    
    If lsOpeAbonoCod = "" Then
        MsgBox "No se ha realizado ningun abono, no existe una operacion Valida de Abono a las cuentas.", vbInformation, "Aviso"
        Exit Function
    End If
    
    If Not (prRS.EOF And prRS.BOF) Then
        lsMovNro = oMov.GeneraMovNro(pgdFecSis, pgsCodAge, pgsCodUser)
        oMov.InsertaMov lsMovNro, lsOpeAbonoCod, psPlanillaDes, gMovEstContabPendiente, gMovFlagVigente
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        
        While Not prRS.EOF
            If prRS.Fields(0) = 1 Then
                If (prRS.Fields(gsRHConceptoITOTING) - prRS.Fields(gsRHConceptoDTOTDES)) > 0 Then
                    lbAbonado = oRH.GetPlanillaPagada(prRS.Fields(2), psPlanillaCod, psPlanillaPeriodo)
                    
                    lsCodCta = oRH.GetCuentaRRHH(prRS.Fields(2), psPlanillaCod)
                
                    If lsCodCta <> "" And Not lbAbonado Then
                        
                        If Mid(lsCodCta, 6, 1) = Moneda.gMonedaExtranjera Then
                            lsMonto = Round(CCur(prRS.Fields(gsRHConceptoITOTING) - prRS.Fields(gsRHConceptoDTOTDES)) / pnTipoCambioVenta, 2)
                        Else
                            lsMonto = prRS.Fields(gsRHConceptoITOTING) - prRS.Fields(gsRHConceptoDTOTDES)
                        End If
                        
                        lsNotaAbono = oPla.GetNroNotaCargoAbono(TpoDocNotaAbono)
                        oMov.InsertaNotaAbonoCargo TpoDocNotaAbono, lsNotaAbono, gNCNARegistrado, gsRHMotivoAbono, Abs(lsMonto)
                        oMov.InsertaNotaAbonoCargoEst TpoDocNotaAbono, lsNotaAbono, gNCNARegistrado, lsMovNro
                        oMov.InsertaRegDocCuenta TpoDocNotaAbono, lsNotaAbono, lsCodCta
                        oMov.Inicio gsFormatoFecha
                        oMov.InsertaMovDoc lnMovNro, TpoDocNotaAbono, lsNotaAbono, Format(pgdFecSis, gsFormatoFecha)
                        
                        'lnSaldo = oCapta.CapAbonoCuentaAho(lsCodCta, lsMonto, 200203, lsMovNro, psPlanillaDes, TpoDocNotaAbono, lsNotaAbono)
                        'Abono de planilla
                        If oCon.AbreConexion Then 'Remota(Left(lsCodCta, 2), , True, , pgdFecSis) Then
                            oCon.BeginTrans
                            lnSaldo = Abono(lsCodCta, CCur(lsMonto), lsOpeAbonoCod, lsOpeAbonoCod, "112" & Left(lsCodCta, 2), pgsCodUser, lsNotaAbono, "PAGO : " & psPlanillaDes & " - " & psPlanillaPeriodo, oCon, pgdFecSis)
                            oRH.SetPlanillaPagada prRS.Fields(2), psPlanillaCod, psPlanillaPeriodo
                            If gbOpeOk Then lsCadBolInd = lsCadBolInd & ImprimeBoletaCad(pgdFecSis, "ABONO RRHH", "Depósito RRHH*Nro." & lsNotaAbono, "", lsMonto, prRS!Nombre, lsCodCta, "", CCur(lnSaldo), 0, "Nota Abono", 0, 0, False, False, , , , , , , , False) & oImpresora.gPrnSaltoPagina 'APRI20170328 SE CAMBIO lsCadBol POR lsCadBolInd
        
                            'Marca la ope de pago
                            'Ver Otros Movimientos
                            Set rsOpe = oExtra.GetExtPlanilla(Left(psPlanillaPeriodo, 6), psPlanillaCod, False, prRS.Fields(2), lsCodCta)
                            
                            If Not (rsOpe.EOF And rsOpe.BOF) Then
                                While Not rsOpe.EOF
                                    If Right(rsOpe!Tipo, 1) = 0 Then
                                        lsNotaAbono = oPla.GetNroNotaCargoAbono(TpoDocNotaCargo)
                                        oMov.InsertaNotaAbonoCargo TpoDocNotaCargo, lsNotaAbono, gNCNARegistrado, gsRHMotivoAbono, Abs(rsOpe!Monto)
                                        oMov.InsertaNotaAbonoCargoEst TpoDocNotaCargo, lsNotaAbono, gNCNARegistrado, lsMovNro
                                        oMov.InsertaRegDocCuenta TpoDocNotaCargo, lsNotaAbono, lsCodCta
                                        oMov.Inicio gsFormatoFecha
                                        oMov.InsertaMovDoc lnMovNro, TpoDocNotaCargo, lsNotaAbono, Format(pgdFecSis, gsFormatoFecha)
                                        
                                        'lnSaldo = oCapta.CapCargoCuentaAho(lsCodCta, rsOpe!Monto, 200303, lsMovNro, psPlanillaDes, TpoDocNotaAbono, lsNotaAbono)
                                        lnSaldo = Cargo(lsCodCta, rsOpe!Monto, gsACRetNCRH, "112" & Left(lsCodCta, 2), pgsCodUser, lsNotaAbono, rsOpe!Comentario, oCon, pgdFecSis)
                                        If gbOpeOk Then lsCadBolInd = lsCadBolInd & ImprimeBoletaCad(pgdFecSis, "RETIRO RRHH", "Retiro RRHH*Nro." & lsNotaAbono, "", rsOpe!Monto, prRS!Nombre, lsCodCta, "", CCur(lnSaldo), 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False) & oImpresora.gPrnSaltoPagina 'APRI20170330 SE CAMBIO lsCadBol POR lsCadBolInd
                                        
                                        'CUENTA ADICIONAL
                                        If Right(rsOpe.Fields(7), 1) <> 0 Then
                                            lsCuentaRef = rsOpe.Fields(10)
                                            lsNotaAbonoRef = oPla.GetNroNotaCargoAbono(TpoDocNotaAbono)
                                            oMov.InsertaNotaAbonoCargo TpoDocNotaAbono, lsNotaAbonoRef, gNCNARegistrado, gsRHMotivoAbono, Abs(rsOpe!Monto)
                                            oMov.InsertaNotaAbonoCargoEst TpoDocNotaAbono, lsNotaAbonoRef, gNCNARegistrado, lsMovNro
                                            oMov.InsertaRegDocCuenta TpoDocNotaAbono, lsNotaAbonoRef, lsCuentaRef
                                            oMov.Inicio gsFormatoFecha
                                            oMov.InsertaMovDoc lnMovNro, TpoDocNotaAbono, lsNotaAbonoRef, Format(pgdFecSis, gsFormatoFecha)
                                            
                                            lnITFValor = 0
                                            If gbITFAplica Then
                                                If VerificaExoneracion(lsCuentaRef) Then
                                                    lnITFValor = 0
                                                Else
                                                    lnITFValor = fgITFCalculaImpuesto(rsOpe!Monto)
                                                End If
                                            End If
                                            
                                            If Left(lsCuentaRef, 2) = Left(lsCodCta, 2) Then
                                                lnSaldo = Abono(lsCuentaRef, rsOpe!Monto, gsACDepNARH, gsCTSDepNARH, "112" & Left(lsCuentaRef, 2), pgsCodUser, lsNotaAbonoRef, rsOpe!Comentario, oCon, pgdFecSis, lnITFValor)
                                                If gbOpeOk Then lsCadBolInd = lsCadBolInd & ImprimeBoletaCad(pgdFecSis, "ABONO RRHH", "Depósito RRHH*Nro." & lsNotaAbonoRef, "", lsMonto, prRS!Nombre, lsCuentaRef, "", CCur(lnSaldo), 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False) & oImpresora.gPrnSaltoPagina 'APRI20170330 SE CAMBIO lsCadBol POR lsCadBolInd
                                            Else
                                                If oConRef.AbreConexion Then 'Remota(Left(lsCuentaRef, 2), , True, , pgdFecSis) Then
                                                    oConRef.BeginTrans
                                                    lnSaldo = Abono(lsCuentaRef, rsOpe!Monto, gsACDepNARH, gsCTSDepNARH, "112" & Left(lsCuentaRef, 2), pgsCodUser, lsNotaAbonoRef, rsOpe!Comentario, oConRef, pgdFecSis, lnITFValor)
                                                    oConRef.CommitTrans
                                                    If gbOpeOk Then lsCadBolInd = lsCadBolInd & ImprimeBoletaCad(pgdFecSis, "ABONO RRHH", "Depósito RRHH*Nro." & lsNotaAbonoRef, "", lsMonto, prRS!Nombre, lsCuentaRef, "", CCur(lnSaldo), 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False) & oImpresora.gPrnSaltoPagina 'APRI20170330 SE CAMBIO lsCadBol POR lsCadBolInd
                                                    oConRef.CierraConexion
                                                Else
                                                    lsNotaAbonoRef = ""
                                                End If
                                            End If
                                        End If
                                    
                                    Else
                                        lsNotaAbono = oPla.GetNroNotaCargoAbono(TpoDocNotaAbono)
                                        oMov.InsertaNotaAbonoCargo TpoDocNotaAbono, lsNotaAbono, gNCNARegistrado, gsRHMotivoAbono, Abs(rsOpe!Monto)
                                        oMov.InsertaNotaAbonoCargoEst TpoDocNotaAbono, lsNotaAbono, gNCNARegistrado, lsMovNro
                                        oMov.InsertaRegDocCuenta TpoDocNotaAbono, lsNotaAbono, lsCodCta
                                        oMov.Inicio gsFormatoFecha
                                        oMov.InsertaMovDoc lnMovNro, TpoDocNotaAbono, lsNotaAbono, Format(pgdFecSis, gsFormatoFecha)
                                        
                                        'lnSaldo = oCapta.CapAbonoCuentaAho(lsCodCta, rsOpe!Monto, 200203, lsMovNro, psPlanillaDes, TpoDocNotaAbono, lsNotaAbono)
                                        lnSaldo = Abono(lsCodCta, rsOpe!Monto, gsACDepNARH, gsCTSDepNARH, "112" & Left(lsCodCta, 2), pgsCodUser, lsNotaAbono, rsOpe!Comentario, oCon, pgdFecSis)
                                        If gbOpeOk Then lsCadBolInd = lsCadBolInd & ImprimeBoletaCad(pgdFecSis, "ABONO RRHH", "Depósito RRHH*Nro." & lsNotaAbono, "", lsMonto, prRS!Nombre, lsCodCta, "", CCur(lnSaldo), 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False) & oImpresora.gPrnSaltoPagina 'APRI20170330 SE CAMBIO lsCadBol POR lsCadBolInd
                                    End If
                                    'If gbOpeOk Then
                                    oExtra.MarcaExtPlanilla Left(psPlanillaPeriodo, 6), psPlanillaCod, Right(rsOpe!Tipo, 1), prRS.Fields(2), lsCodCta, lsNotaAbono, lsNotaAbonoRef, lsCuentaRef
                                    rsOpe.MoveNext
                                Wend
                            End If
                            rsOpe.Close
                    
                        oCon.CommitTrans
                        oCon.CierraConexion
                        End If
                    Else
                        If lbAbonado Then
                            MsgBox "Para el recurso Humamo : " & prRS.Fields(1) & " - " & prRS!Nombre & " . Su sueldo ya fue abonado.", vbInformation, "Aviso"
                        Else
                            MsgBox "El recurso Humamo : " & prRS.Fields(1) & " - " & prRS!Nombre & " . No tiene cuenta asignada o la cuenta asignada tiene un estado no valido.", vbInformation, "Aviso"
                        End If
                    End If
                Else
                    MsgBox "El recurso Humamo : " & prRS.Fields(1) & " - " & prRS!Nombre & " . Tiene un neto a Pagar Negativo.", vbInformation, "Aviso"
                End If
                
                'APRI20170328
                If (psPlanillaCod = "21" Or psPlanillaCod = "22" Or psPlanillaCod = "27" Or psPlanillaCod = "28" Or psPlanillaCod = "41" Or psPlanillaCod = "42") And lsCadBolInd <> "" Then
                    Planilla = psPlanillaCod & pSNomPlanilla
                    GeneraLogAbonoRH lsCadBolInd, Planilla
                End If
                    lsCadBol = lsCadBol & lsCadBolInd
                   lsCadBolInd = ""
                 'END APRI
            End If
            prRS.MoveNext
        Wend
    End If
        
    AbonoPagos = lsCadBol
    
    Exit Function
AgregaConceptoEmpleadoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoEmpleado:AgregaConceptoEmpleado Method")
End Function

''Agrega una nueva relación entre concepto y Empleado
''##ModelId=3A9BEAB1011F
Public Function AbonoPagos22(prRS As ADODB.Recordset, psPlanillaPeriodo As String, psPlanillaCod As String, psPlanillaDes As String, pnTipoCambioVenta As Currency, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, pgsCodAge As String, pgsCodUser As String, pSNomPlanilla As String) As String
    'pSNomPlanilla 'APRI20170328
    On Error GoTo AgregaConceptoEmpleadoErr

    Dim lsNotaAbono As String
    Dim lsNotaAbonoRef As String
    
    
    Dim lsCuentaRef As String
    
    Dim oConsistema As NConstSistemas
    Set oConsistema = New NConstSistemas
    Dim lnDispPorcent As Double
    
    lnDispPorcent = oConsistema.LeeConstSistema(222)
    
    Dim lsMonto As String
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lsCodCta As String
    Dim oPla As DRHProcesosCierre
    Set oPla = New DRHProcesosCierre
    Dim oExtra As DActualizaMovExtraPlanilla
    Set oExtra = New DActualizaMovExtraPlanilla
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    Dim rsOpe As ADODB.Recordset
    Set rsOpe = New ADODB.Recordset
    Dim oMov As DMov
    Set oMov = New DMov
    'Dim oPla As DRHProcesosCierre
    'Set oPla = New DRHProcesosCierre
    Dim oCapta As NCapMovimientos
    Set oCapta = New NCapMovimientos
    Dim lnSaldo As Double
    Dim lsCadBol As String
    Dim lsCadBolInd As String 'APRI20270328
    Dim Planilla As String 'APRI20270328
    Dim lsOpeAbonoCod As String
    
    Dim lbAbonado As Boolean
    
    'MAVM 20120504 ***
    Dim rsDatosFractal As ADODB.Recordset
    Set rsDatosFractal = New ADODB.Recordset
    Dim lnIntangibleCTS As Currency
    '***
    
    lsOpeAbonoCod = GetOpeAbonoPlanilla(psPlanillaCod)

    If lsOpeAbonoCod = "" Then
        MsgBox "No se ha realizado ningun abono, no existe una operacion Valida de Abono a las cuentas.", vbInformation, "Aviso"
        Exit Function
    End If
    
    gsNomAge = Trim(pgsNomAge)
    
    If Not (prRS.EOF And prRS.BOF) Then
        
        While Not prRS.EOF
            If prRS.Fields(0) = 1 Then
                
                lbAbonado = oRH.GetPlanillaPagada(prRS.Fields(2), psPlanillaCod, psPlanillaPeriodo)
                
                'MAVM 20120223 Cambios Fractal
                'lsCodCta = oRH.GetCuentaRRHH(prRS.Fields(2), psPlanillaCod)
                lsCodCta = oRH.GetCuentaRRHH(prRS.Fields(2), psPlanillaCod, psPlanillaPeriodo)
                '***
    
                If lsCodCta <> "" And prRS.Fields(2) <> "" And Not lbAbonado Then
                    lsMovNro = oMov.GeneraMovNro(pgdFecSis, pgsCodAge, pgsCodUser)
                    oMov.InsertaMov lsMovNro, lsOpeAbonoCod, psPlanillaDes, gMovEstContabPendiente, gMovFlagVigente
                    lnMovNro = oMov.GetnMovNro(lsMovNro)
    
                    If Mid(lsCodCta, 9, 1) = "2" Then
                        'MAVM 20120505 *** FRACTAL
                        'lsMonto = CCur(prRS.Fields(gsRHConceptoITOTING) - prRS.Fields(gsRHConceptoDTOTDES)) / pnTipoCambioVenta
                        lsMonto = CCur(prRS.Fields(gsRHConceptoINETOPAGAR))
                        '***
                    Else
                        'MAVM 20120214 *** FRACTAL
                        'lsMonto = prRS.Fields(gsRHConceptoITOTING) - prRS.Fields(gsRHConceptoDTOTDES)
                        lsMonto = prRS.Fields(gsRHConceptoINETOPAGAR)
                        '***
                    End If
    
                    lsNotaAbono = oPla.GetNroNotaCargoAbono(TpoDocNotaAbono)
                    oMov.InsertaNotaAbonoCargo TpoDocNotaAbono, lsNotaAbono, gNCNARegistrado, gsRHMotivoAbono, Abs(lsMonto)
                    oMov.InsertaNotaAbonoCargoEst TpoDocNotaAbono, lsNotaAbono, gNCNARegistrado, lsMovNro
                    oMov.InsertaRegDocCuenta TpoDocNotaAbono, lsNotaAbono, lsCodCta
                    oMov.Inicio gsFormatoFecha
                    'oMov.InsertaMovDoc lnMovNro, TpoDocNotaAbono, lsNotaAbono, Format(pgdFecSis, gsFormatoFecha)
                    
                    'MAVM 20120504 ***
                    'If psPlanillaCod = gsRHPlanillaCTS Then
                    If psPlanillaCod = gsRHPlanillaCTSFractal Then
                        'Set rsDatosFractal = oRH.GetRHPlanillaDetConFractal(prRS.Fields(2), psPlanillaCod, psPlanillaPeriodo)
                        Set rsDatosFractal = oRH.ObtenerCapSaldosCuentasCTS(lsCodCta, pnTipoCambioVenta) 'APRI20191025 MEJORA PROCESO
                        If rsDatosFractal.RecordCount <> "0" Then
                            'lnIntangibleCTS = rsDatosFractal!nMontoIntangibleCTS
                            lnIntangibleCTS = rsDatosFractal!TotalSueldos 'APRI20191025 MEJORA PROCESO
                        End If
                        rsDatosFractal.Close
                        'lnSaldo = oCapta.CapAbonoCuentaCTS(lsCodCta, lsMonto, lsOpeAbonoCod, lsMovNro, psPlanillaDes, lnDispPorcent, TpoDocNotaAbono, lsNotaAbono, , , , , , , , , , , pnTipoCambioVenta)
                        lnSaldo = oCapta.CapAbonoCuentaCTS(lsCodCta, lsMonto, lsOpeAbonoCod, lsMovNro, psPlanillaDes, lnDispPorcent, TpoDocNotaAbono, lsNotaAbono, , , , , , , , , , , pnTipoCambioVenta, lnIntangibleCTS)
                        '***
                    Else
                        lnSaldo = oCapta.CapAbonoCuentaAho(lsCodCta, lsMonto, lsOpeAbonoCod, lsMovNro, psPlanillaDes, TpoDocNotaAbono, lsNotaAbono, , , , , , , , , , , False)
                    End If
                    oRH.SetPlanillaPagada prRS.Fields(2), psPlanillaCod, psPlanillaPeriodo
                                    
                    lsCadBolInd = lsCadBolInd & ImprimeBoletaCad(pgdFecSis, "ABONO RRHH", "Depósito RRHH*Nro." & lsNotaAbono, "", lsMonto, prRS!Nombre, lsCodCta, "", CCur(lnSaldo), 0, "Nota Abono", 0, 0, False, False, , , , , , , , False) & oImpresora.gPrnSaltoPagina 'APRI20170328 SE CAMBIO lsCadBol POR lsCadBolInd
                    
                    'Marca la ope de pago
                    'Ver Otros Movimientos
                    Set rsOpe = oExtra.GetExtPlanilla(psPlanillaPeriodo, psPlanillaCod, False, prRS.Fields(2), lsCodCta)
    
                    If Not (rsOpe.EOF And rsOpe.BOF) Then
                        While Not rsOpe.EOF
                            If rsOpe!nTipoOpe = 0 Then
                                lsNotaAbono = oPla.GetNroNotaCargoAbono(TpoDocNotaCargo)
                                oMov.InsertaNotaAbonoCargo TpoDocNotaCargo, lsNotaAbono, gNCNARegistrado, gsRHMotivoAbono, Abs(rsOpe!nRHExtraMonto)
                                oMov.InsertaNotaAbonoCargoEst TpoDocNotaCargo, lsNotaAbono, gNCNARegistrado, lsMovNro
                                oMov.InsertaRegDocCuenta TpoDocNotaCargo, lsNotaAbono, lsCodCta
                                oMov.Inicio gsFormatoFecha
                                'oMov.InsertaMovDoc lnMovNro, TpoDocNotaCargo, lsNotaAbono, Format(pgdFecSis, gsFormatoFecha)
                                
                                If psPlanillaCod = gsRHPlanillaCTS Then
                                    lnSaldo = oCapta.CapCargoCuentaCTS(lsCodCta, rsOpe!nRHExtraMonto, 220303, lsMovNro, psPlanillaDes, TpoDocNotaAbono, lsNotaAbono)
                                Else
                                    lnSaldo = oCapta.CapCargoCuentaAho(lsCodCta, rsOpe!nRHExtraMonto, 200303, lsMovNro, psPlanillaDes, TpoDocNotaAbono, lsNotaAbono)
                                End If
                                lsCadBolInd = lsCadBolInd & ImprimeBoletaCad(pgdFecSis, "RETIRO RRHH", "Retiro RRHH*Nro." & lsNotaAbono, "", rsOpe!nRHExtraMonto, prRS!Nombre, lsCodCta, "", CCur(lnSaldo), 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False) & oImpresora.gPrnSaltoPagina 'APRI20170328 SE CAMBIO lsCadBol POR lsCadBolInd
                                
                                
                                
                                
                            Else
                                lsNotaAbono = oPla.GetNroNotaCargoAbono(TpoDocNotaAbono)
                                oMov.InsertaNotaAbonoCargo TpoDocNotaAbono, lsNotaAbono, gNCNARegistrado, gsRHMotivoAbono, Abs(rsOpe!nRHExtraMonto)
                                oMov.InsertaNotaAbonoCargoEst TpoDocNotaAbono, lsNotaAbono, gNCNARegistrado, lsMovNro
                                oMov.InsertaRegDocCuenta TpoDocNotaAbono, lsNotaAbono, lsCodCta
                                oMov.Inicio gsFormatoFecha
                                'oMov.InsertaMovDoc lnMovNro, TpoDocNotaAbono, lsNotaAbono, Format(pgdFecSis, gsFormatoFecha)
                                
                                If psPlanillaCod = gsRHPlanillaCTS Then
                                    lnSaldo = oCapta.CapAbonoCuentaAho(lsCodCta, rsOpe!nRHExtraMonto, 220203, lsMovNro, psPlanillaDes, TpoDocNotaAbono, lsNotaAbono)
                                Else
                                    lnSaldo = oCapta.CapAbonoCuentaAho(lsCodCta, rsOpe!nRHExtraMonto, 200203, lsMovNro, psPlanillaDes, TpoDocNotaAbono, lsNotaAbono)
                                End If
                                lsCadBolInd = lsCadBolInd & ImprimeBoletaCad(pgdFecSis, "ABONO RRHH", "Depósito RRHH*Nro." & lsNotaAbono, "", lsMonto, prRS!Nombre, lsCodCta, "", CCur(lnSaldo), 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False) & oImpresora.gPrnSaltoPagina 'APRI20170328 SE CAMBIO lsCadBol POR lsCadBolInd
                            End If
    
                            oExtra.MarcaExtPlanilla psPlanillaPeriodo, psPlanillaCod, rsOpe!nTipoOpe, prRS.Fields(2), lsCodCta, lsNotaAbono, lsNotaAbonoRef, lsCuentaRef
    
                            rsOpe.MoveNext
                        Wend
                    End If
                    rsOpe.Close
                Else
                    If lbAbonado Then
                        MsgBox "Para el recurso Humamo : " & prRS.Fields(2) & " - " & prRS.Fields(3) & ", ya se proceso el abono de esta planilla.", vbInformation, "Aviso"
                    Else
                        If prRS.Fields(2) <> "" Then MsgBox "El recurso Humamo : " & prRS.Fields(2) & " - " & prRS.Fields(3) & ". No tiene cuenta asignada o la cuenta asignada tiene un estado no valido.", vbInformation, "Aviso"
                    End If
                End If
                  'APRI20170328
                If (psPlanillaCod = "21" Or psPlanillaCod = "22" Or psPlanillaCod = "27" Or psPlanillaCod = "28" Or psPlanillaCod = "41" Or psPlanillaCod = "42") And lsCadBolInd <> "" Then
                    Planilla = psPlanillaCod & pSNomPlanilla
                    GeneraLogAbonoRH lsCadBolInd, Planilla
                End If
                    lsCadBol = lsCadBol & lsCadBolInd
                   lsCadBolInd = ""
                 'END APRI
            End If
            prRS.MoveNext
        Wend
    End If
    AbonoPagos22 = lsCadBol
    Exit Function
AgregaConceptoEmpleadoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoEmpleado:AgregaConceptoEmpleado Method")
End Function

Public Function CierreMesRRHH(ByVal pdFecha As Date, pgsCodAge As String, pgsCodUser As String) As String
    Dim rsEmp As ADODB.Recordset
    Dim nGratEst As Double, nTerEst As Double
    Dim nCTSEst As Double, nVacEst As Double
    Dim nBonVacEst As Double
    Dim nGratCon As Double, nTerCon As Double
    Dim nCTSCon As Double, nVacCon As Double
    Dim nBonVacCon As Double
    Dim oProc As DRHProcesosCierre
    Set oProc = New DRHProcesosCierre
    Dim nDiasV As Double
    Dim bTrans As Boolean
    Set rsEmp = New ADODB.Recordset
    Dim lsPeriodo As String
    Dim lsMov1 As String
    Dim lsMov2 As String
    Dim lsMov3 As String
    Dim lsMov4 As String
    Dim oMov As NContImprimir
    Set oMov = New NContImprimir
    
    lsPeriodo = Format(pdFecha, "yyyymm")
    
    If Not oProc.VerificaPlanillaProvision(lsPeriodo) Then
        'Primero se debe incrementar los dias de vacaciones por empleado
        'por el mes transcurrido y si no lo haces, sonamos!!!!
        nDiasV = gnRHNumDiasVac / 12
        oProc.IncrementoVac nDiasV
        
        'Lista todos los empleados activos, no vacaciones, no retirado, será cierto eso...!!!!
        Set rsEmp = oProc.CierreMesListaRRHH
        
        Set rsEmp.ActiveConnection = Nothing
        nGratEst = 0: nTerEst = 0
        nCTSEst = 0: nVacEst = 0
        nBonVacEst = 0
        nGratCon = 0: nTerCon = 0
        nCTSCon = 0: nVacCon = 0
        nBonVacCon = 0
        
        
        With rsEmp
            RaiseEvent ShowProgress
            Do While Not .EOF
                'Generar provision por la Gratificacion y realizar el asiento contable correspondiente.
                If !TipoContrato = RHContratoTipo.RHContratoTipoIndeterminado Then
                    oProc.SetPlanillaProvision lsPeriodo, !cPersCod, gsRHPlanillaProvGratificacion, oProc.GetProvGrat(!cPersCod, pdFecha)
                    'oProc.SetPlanillaProvision lsPeriodo, !cPersCod, gsRHPlanillaProvVacaciones, oProc.GetProvVac(!cPersCod, pdFecha)
                    oProc.SetPlanillaProvision lsPeriodo, !cPersCod, gsRHPlanillaProvCTS, oProc.GetProvCTS(!cPersCod, pdFecha, "")
                    'oProc.SetPlanillaProvision lsPeriodo, !cPersCod, gsRHPlanillaProvTercio, oProc.GetProvTercio(!cPersCod, pdFecha)
                    'If Not IsNull(!cBonVac) Then nBonVacEst = nBonVacEst + oProc.GetProvBonifVac(!cPersCod, pdFecha)
                Else
                    oProc.SetPlanillaProvision lsPeriodo, !cPersCod, gsRHPlanillaProvGratificacion, oProc.GetProvGrat(!cPersCod, pdFecha)
                    'oProc.SetPlanillaProvision lsPeriodo, !cPersCod, gsRHPlanillaProvVacaciones, oProc.GetProvVac(!cPersCod, pdFecha)
                    oProc.SetPlanillaProvision lsPeriodo, !cPersCod, gsRHPlanillaProvCTS, oProc.GetProvCTS(!cPersCod, pdFecha, "")
                    'oProc.SetPlanillaProvision lsPeriodo, !cPersCod, gsRHPlanillaProvTercio, oProc.GetProvTercio(!cPersCod, pdFecha)
                    'If Not IsNull(!cBonVac) Then oProc.SetPlanillaProvision lsPeriodo, !cPersCod, gsRHPlanillaProvBonVacaciones, oProc.GetProvBonifVac(!cPersCod, pdFecha)
                End If
                
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
    Else
        If MsgBox("La planilla de Provision de este mes ya ha sido generada, desea volver a generar el asiento.", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
            Set rsEmp = Nothing
            Set oProc = Nothing
            Exit Function
        End If
    End If
    
    'Genera asiento de la provision
    lsMov1 = oProc.GrabaAsientoProvision(lsPeriodo, pdFecha, pgsCodAge, pgsCodUser, "0", gsRHPlanillaGratificacionProvEst)
    lsMov2 = oProc.GrabaAsientoProvision(lsPeriodo, pdFecha, pgsCodAge, pgsCodUser, "1", gsRHPlanillaGratificacionProvCon)
    lsMov3 = oProc.GrabaAsientoProvision(lsPeriodo, pdFecha, pgsCodAge, pgsCodUser, "0", gsRHPlanillaCTSProvEst)
    lsMov4 = oProc.GrabaAsientoProvision(lsPeriodo, pdFecha, pgsCodAge, pgsCodUser, "1", gsRHPlanillaCTSProvCon)
    
    CierreMesRRHH = oMov.ImprimeAsientoContable(lsMov1, 66, 80, , , , , False) & oImpresora.gPrnSaltoPagina & oMov.ImprimeAsientoContable(lsMov2, 66, 80, , , , , False) & oImpresora.gPrnSaltoPagina & oMov.ImprimeAsientoContable(lsMov3, 66, 80, , , , , False) & oImpresora.gPrnSaltoPagina & oMov.ImprimeAsientoContable(lsMov4, 66, 80, , , , , False)
      
    Set oMov = Nothing
    Set rsEmp = Nothing
    Set oProc = Nothing
End Function

#If DebugMode Then
    '##ModelId=3AB90301024D
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If


Private Function Abono(psCodCta As String, pnMonto As Currency, pgsACDepNA As String, pgsCTSDepNA As String, pgsCodAge As String, pgsCodUser As String, NroDoc As String, psDes As String, poCon As DConecta, pgdFecSis As Date, Optional pnITFValor As Currency = 0)
    Dim sqlA As String
    Dim sqlT As String
    Dim sqlTra As String
    Dim sqlVerifica As String
    Dim sqlNota As String
    Dim lnInteres As Currency
    Dim RegCuenta As ADODB.Recordset
    Set RegCuenta = New ADODB.Recordset
    
    Dim lnSaldoDisponible As Currency
    Dim lnSaldCnt As Currency
    Dim lnSaldDisp As Currency
    Dim lbSAux As Boolean
    Dim lcEst As String
    Dim lNTInt As Double
    Dim lsHoy As String
    Dim lsAyer As String
    Dim I As Long
    Dim lsAgeTmp As String
    
    lsHoy = Format(pgdFecSis, "mm/dd/yyyy") + " " + Format(Time, "hh:mm:ss AMPM")
    lsAyer = Format(pgdFecSis - 1, "mm/dd/yyyy") + " " + Format(Time, "hh:mm:ss AMPM")

    
    If Mid(psCodCta, 3, 3) = Producto.gCapAhorros Then  'Ahorro
        sqlA = "Select * from AhorroC where cCodCta = '" & psCodCta & "'"
        Set RegCuenta = poCon.CargaRecordSet(sqlA)
        
        lnInteres = RegCuenta!nInteres + GetInteres(RegCuenta!nTasaIntAC, RegCuenta!nsalddispac, DateDiff("d", RegCuenta!dUltMovAC, pgdFecSis) - 1, True)
        
        lnSaldoDisponible = RegCuenta!nsalddispac
        lnSaldCnt = RegCuenta!nSaldCntAC + CCur(pnMonto)
        lnSaldDisp = RegCuenta!nsalddispac + CCur(pnMonto)

        lbSAux = False
        If RegCuenta!bInactiva Then
            lcEst = "0"
            lbSAux = True
            'Realiza la Transaccion de activación

            sqlT = "Insert TranDiaria (dFecTran, cCodUsu, cCodOpe, cCodCta, cNumDoc, nMonTran, nSaldCnt, cCodUsuRem, cCodAge, cFlag, nTipCambio) values('" & lsHoy & "' , '" & pgsCodUser & "' , '" _
                    & gsACInaAct & "' , '" & psCodCta & "' , null , " & Str(lnSaldoDisponible) & " , " & lnSaldCnt & ", null , '" & pgsCodAge & "' , NULL,0)"
        Else
            lcEst = "0"
        End If

        If RegCuenta!cTipTasAC = "T" Then
            lNTInt = GetNuevaTasaAC(psCodCta, lnSaldDisp, RegCuenta!nTasaIntAC, poCon)
        Else
            lNTInt = RegCuenta!nTasaIntAC
        End If

        sqlA = " Update AhorroC " _
             & " Set nSaldAntAC=nSaldDispAC, nSaldDispAC=nSaldDispAC+" & pnMonto & ", nSaldCntAC=nSaldCntAC+" & pnMonto & ", dUltMovAC='" & lsAyer & "', nInteres = " & lnInteres & " , dUltActAC='" & lsHoy & "', dUltCntAC = '" + lsHoy + "' ,cCodUsu= '" & pgsCodUser & "', nNumExt = nNumExt + 1 , ntasaintAC = " + Str(lNTInt) + ", bInactiva = " & lcEst & " " _
             & " where cCodCta='" & psCodCta & "'"
        
        sqlTra = " Insert TranDiaria (dFecTran, cCodUsu, cCodOpe, cCodCta, cNumDoc, nMonTran, nSaldCnt, cCodUsuRem, cCodAge, cFlag, nTipCambio) values('" & lsHoy & "' , '" & pgsCodUser & "' , '" _
               & pgsACDepNA & "' , '" & psCodCta & "' , '" & NroDoc & "', " & pnMonto & " , " & Str(lnSaldCnt) & " , null , '" & pgsCodAge & "' , NULL,0)"
        
        sqlVerifica = "Select Max(nNumTran) nNumTran from TranDiaria Where dFecTran = '" & lsHoy & "' And cCodCta = '" & psCodCta & "' AND  cCodAge = '" & pgsCodAge & "' AND cCodUsu = '" & pgsCodUser & "' AND cCodOpe = '" & pgsACDepNA & "' "
        
        RegCuenta.Close
        Set RegCuenta = Nothing
            
        Randomize
        For I = 0 To Rnd(2000) * 1000
        Next I
        
        poCon.Ejecutar sqlA
        If lbSAux Then poCon.Ejecutar sqlT
        poCon.Ejecutar sqlTra
        
        If gbITFAplica And pnITFValor > 0 Then
            lsAgeTmp = gsCodAge
            gsCodAge = pgsCodAge
            gnITFNumTranOrigen = fgITFGetNumtranOrigen(poCon.ConexionActiva)
            gnITFNumTran = fgITFACRetiroImpuesto(psCodCta, pnITFValor, "", gsITFACCargoCuenta, FechaHora(pgdFecSis), pgsCodUser, poCon.ConexionActiva)
            fgITFDetalleInserta gnITFNumTranOrigen, Right(pgsCodAge, 2), CCur(pnMonto), pnITFValor, gnITFNumTran, poCon.ConexionActiva
            gsCodAge = lsAgeTmp
        End If
        
        Set RegCuenta = poCon.CargaRecordSet(sqlVerifica)
        
        If RegCuenta.EOF And RegCuenta.BOF Then
            Set RegCuenta = Nothing
            GoTo ERROR
        End If
                  
        sqlT = Str(RegCuenta!nNumTran)
        
        RegCuenta.Close
        Set RegCuenta = Nothing
                
        sqlNota = " INSERT OtrasOpe(dFecTran,nNumTran,cCodOpe,cGlosa,cNumDoc,cCodPers,cFlag) " _
                & " Values('" & lsHoy & "'," & sqlT & ",'" & pgsACDepNA & "','" & psDes & "','" & NroDoc & "',null,null)"
        poCon.Ejecutar sqlNota
                  
        gbOpeOk = True
        Abono = lnSaldDisp
    
    'Fin Ahorros
    ElseIf Mid(psCodCta, 3, 3) = Producto.gCapCTS Then 'Inicio CTS
        Dim lsPorRet As Currency
        Dim lnInteres1 As Currency
        Dim lnSaldAcu As Currency
        Dim lnIntAnt As Currency
        
        'Calcula el nuevos saldos
        sqlA = "SELECT * FROM CTS WHERE cCodCta = '" & psCodCta & "'"
        Set RegCuenta = poCon.CargaRecordSet(sqlA)
    
        lsPorRet = (ReadParametros("23110", poCon) / 100)
    
        If lsPorRet = -1 Then
            MsgBox "Error no se encontro el porcentaje de Retiro de CTS, llame a la Oficina de Sistemas (VarSistema)", vbInformation, "Aviso"
            GoTo ERROR
        End If
          
        lnIntAnt = GetInteres(RegCuenta!nTasaIntCTS, RegCuenta!nCapitalCTS, DateDiff("d", Format(RegCuenta!dUltMovCTS, "dd/mm/yyyy"), gdFecSis) - 1, True)
        
        lnInteres = RegCuenta!nInteres + lnIntAnt
        
        lnInteres1 = RegCuenta!nIntSald + GetInteres(RegCuenta!nTasaIntCTS, RegCuenta!nSaldDispCTS, DateDiff("d", Format(RegCuenta!dUltMovCTS, "dd/mm/yyyy"), gdFecSis) - 1, True)
        
        lnSaldCnt = RegCuenta!nSaldCntCTS
        lnSaldDisp = RegCuenta!nSaldDispCTS
        lnSaldAcu = RegCuenta!nCapitalCTS
        
        lnSaldCnt = lnSaldCnt + CCur(pnMonto)
        lnSaldAcu = lnSaldAcu + CCur(pnMonto)
        lnSaldDisp = lnSaldDisp + CCur(pnMonto) * lsPorRet
        
        'Actualiza la tabla de CTS
        sqlA = " Update CTS " _
             & " Set nSaldCntCTS = nSaldCntCTS + " & Str(Round(CCur(pnMonto), 2)) & ", nUltDepCTS = " & Str(Round(CCur(pnMonto), 2)) & ", nCapitalCTS = nCapitalCTS +" & Str(Round(CCur(pnMonto), 2)) & ", " _
             & " nSaldDispCTS = nSaldDispCTS + " & Str(Round(CCur(pnMonto) * lsPorRet, 2)) & ", dUltDepCTS = '" & lsHoy & "', dUltCntCTS = '" & lsHoy & "', " _
             & " dUltActCTS = '" & lsHoy & "', cCodUsu = '" & pgsCodUser & "', nNumExt = nNumExt + 1, dUltMovCTS= '" & lsAyer & "' , nInteres = " + Str(Round(lnInteres, 2)) + ", nIntSald = " + Str(Round(lnInteres1, 2)) + " " _
             & " where cCodCta= '" & psCodCta & "'"
        
        'Inserta nueva operación en Transacción Diaria
        sqlT = "Insert TranDiaria (dFecTran, cCodUsu, cCodOpe, cCodCta, cNumDoc, nMonTran, nSaldCnt, cCodUsuRem, cCodAge, cFlag, nTipCambio) values('" & lsHoy & "','" & pgsCodUser & "','" _
                & pgsCTSDepNA & "','" & psCodCta & "',Null," _
                & pnMonto & "," & Str(lnSaldCnt) & ",null,'" & pgsCodAge & "', null,0)"
        
        sqlVerifica = " Select nNumTran from TranDiaria Where dFecTran = '" & lsHoy & "' And cCodCta = '" & psCodCta & "' And  cCodAge = '" & pgsCodAge & "' AND cCodUsu = '" & pgsCodUser & "' and cCodOpe = '" & pgsCTSDepNA & "'"
       
        RegCuenta.Close
       
        Randomize
        For I = 0 To Rnd(2000) * 1000
        Next I
                
        poCon.Ejecutar sqlA
        poCon.Ejecutar sqlT
'*************************************
            If lnIntAnt > 0 Then
                sqlT = " Insert TranDiaria (dFecTran, cCodUsu, cCodOpe, cCodCta, cNumDoc, nMonTran, nSaldCnt, cCodUsuRem, cCodAge, cFlag, nTipCambio) values('" & lsHoy & "','" _
                     & pgsCodUser & "','" & gsCTSProInt & "','" & psCodCta & "',Null," _
                          & Str(Round(lnIntAnt, 2)) & "," & Str(Round(lnSaldCnt, 2)) & ",null,'" & pgsCodAge & "',null,0)"
                poCon.Ejecutar sqlT
            End If
'*************************************
        Set RegCuenta = poCon.CargaRecordSet(sqlVerifica)
        
        sqlA = Str(RegCuenta!nNumTran)
                
        RegCuenta.Close
        Set RegCuenta = Nothing
                        
        sqlNota = " INSERT OtrasOpe(dFecTran,nNumTran,cCodOpe,cGlosa,cNumDoc,cCodPers,cFlag) " _
                & " values('" & lsHoy & "'," & sqlA & ",'" & pgsCTSDepNA & "','" & psDes & "','" & NroDoc & "',null,null)"
        poCon.Ejecutar sqlNota
        
        gbOpeOk = True
        Abono = lnSaldDisp
    'fin CTS
    End If
    
    Exit Function
ERROR:
    MsgBox Err.Description, vbInformation, "Aviso"
End Function

'***********************************************
' GetInteres: Funcion que devuelve el interes de la cuenta este interes puede ser
' simple o compuesto de acuerdo al ultimo argumento de la funcion
' si este es True es I simple de lo contrario sera compuesto
'***********************************************
Public Function GetInteres(pnTasa As Currency, pnSaldo As Currency, pnDias As Integer, Simple As Boolean) As Double
    Dim lnTasaEfec As Double
    
    If Simple Then
        If pnDias > 0 Then
            lnTasaEfec = (pnTasa / 100) / 360
            GetInteres = Round((lnTasaEfec) * pnDias * pnSaldo, 2)
        Else
            GetInteres = 0
        End If
    Else
        If pnDias > 0 Then
            lnTasaEfec = (pnTasa / 100) / 360
            GetInteres = Round((((lnTasaEfec + 1) ^ pnDias) - 1) * pnSaldo, 2)
        Else
            GetInteres = 0
        End If
    End If

End Function

'*************************************************
'GetNuevaTasaAC:Funcion que devuelve la nueva tasa de Interes de la cuenta
'segun el nuevo saldo de la cuenta, solo se usa para
'ahorro.
'recibe como parametro el numero de cuenta
'El nuevo saldo, el interes actual y la conexion
'*************************************************
Public Function GetNuevaTasaAC(pCuenta As String, pSaldo As Currency, pInt As Double, poCon As DConecta) As Double
    Dim lcCodInt As String
    On Error GoTo ERROR
    lcCodInt = Mid(pCuenta, 3, 4)
    Dim RegCuenta As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT nTasaInt FROM TASAINT WHERE cTipTas = 'T' AND nMonInf<= " + Str(pSaldo) + " AND nMonSup>=" + Str(pSaldo) + " AND cCodTas = '" + lcCodInt + "'"
    
    Set RegCuenta = poCon.CargaRecordSet(sql)
    
    If RegCuenta.BOF And RegCuenta.EOF Then
        MsgBox "No Existe un Saldo asignable al monto existente en la cuenta, la cuenta permanecera con el interes anterior "
        GetNuevaTasaAC = pInt
        Set RegCuenta = Nothing
    Else
        GetNuevaTasaAC = RegCuenta!nTasaInt
        RegCuenta.Close
        Set RegCuenta = Nothing
    End If
    
    'Set oCon = Nothing
    
    Exit Function
ERROR:
'        MsgBox Err.Description + "Funciones Principales", vbCritical, frmRealizaOpe.caption
        GetNuevaTasaAC = RegCuenta!nTasaInt
End Function


' Procedimeinto que imprime la boleta de Operación
' luego de una operacion de cajero por impresora
Public Function ImprimeBoletaCad(pgdFecSis As Date, ByVal psTit As String, psText As String, _
                            psCodOpe As String, pnMonto As String, _
                            psCliente As String, psCodCta As String, _
                            psNumDoc As String, pnSaldo As Currency, pnInteresA As String, NomDoc As String, _
                            pnNumExt As Long, pnSaldoC As Double, Optional pbOpSaldoC As Boolean = True, _
                            Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
                            Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", Optional pbConfirmacion As Boolean = True) As String
    Dim sFecha As String
    Dim sHora As String
    Dim sSep As Integer
    Dim sIni As Integer
    Dim sMonto As String
    Dim sSDisp As String
    Dim sIntAcum As String
    Dim sMax  As Integer
    Dim sAux As Integer
    Dim lsNegritaOn As String
    Dim lsNegritaOff As String
    Dim lsNroExt As String
    Dim lnTope As Integer
    Dim lsSaldoC As String
    Dim lsCadArg As String
    Dim lsInteres As String
    
    Dim lnCliAux As Long
    Dim lsCliAux1 As String
    Dim lsCliAux2 As String
    
    Dim lnChq As Long
    Dim lsChqAux1 As String
    Dim lsChqAux2 As String
    Dim lsNomAge As String
    
    Dim lnNumLinCmac As Integer
    
    Dim lsCadena As String
    
ETIQ:
    
    On Error GoTo ERROR
    
    lnTope = 0 '6 'Tope de lineas en Boleta
    
    lsNegritaOn = oImpresora.gPrnBoldON
    lsNegritaOff = oImpresora.gPrnBoldOFF
    
    lsNroExt = Str(pnNumExt)
    
    lsCadena = lsCadena & oImpresora.gPrnInicializa
    
    lsCadena = lsCadena & oImpresora.gPrnEspaLineaN       'espaciamiento lineas 1/6 pulg.
    lsCadena = lsCadena & oImpresora.gPrnTamPagina22     'Longitud de página a 22 líneas'
    lsCadena = lsCadena & oImpresora.gPrnTamLetra10CPI      'Tamaño 10 cpi
    lsCadena = lsCadena & oImpresora.gPrnTpoLetraSansSerif        'Tipo de Letra Sans Serif
    lsCadena = lsCadena & oImpresora.gPrnCondensadaOFF   ' cancela condensada
    lsCadena = lsCadena & oImpresora.gPrnBoldOFF   ' desactiva negrita
    
    sSep = 15
    sIni = 1
    sMax = 33
    sAux = 5
    
    
    sFecha = Format$(pgdFecSis, "dd/mm/yyyy")
    sHora = Format$(Time, "hh:mm:ss")
    sMonto = Format$(pnMonto, "#,##0.00")
    sSDisp = Format$(pnSaldo, "#,##0.00")
    lsSaldoC = Format$(pnSaldoC, "#,##0.00")
    
    lsNomAge = gsNomAge
    
    'lsCadena = lscadena &  oImpresora.gPrnSaltoLinea  &
    lsCadena = lsCadena & lsNegritaOn  'Activa Negrita
    lsCadena = lsCadena & Space(sIni) & "CMACM - AHORRO" & Space(19 + sSep + sAux) & "CMACM - AHORRO" & oImpresora.gPrnSaltoLinea
    
    If Len(psCodCta) = 12 Then
        If Mid(psCodCta, 6, 1) = 1 Then
            lsCadena = lsCadena & Space(sIni) & Trim(gsNomAge) & "-SOLES" & Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(gsNomAge) & "-SOLES" & Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-SOLES")) & lsNroExt & oImpresora.gPrnSaltoLinea
        Else
            lsCadena = lsCadena & Space(sIni) & Trim(gsNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(gsNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & oImpresora.gPrnSaltoLinea
        End If
    Else
        If Mid(psCodCta, 9, 1) = 1 Then
            lsCadena = lsCadena & Space(sIni) & Trim(gsNomAge) & "-SOLES" & Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(gsNomAge) & "-SOLES" & Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-SOLES")) & lsNroExt & oImpresora.gPrnSaltoLinea
        Else
            lsCadena = lsCadena & Space(sIni) & Trim(gsNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(gsNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(gsNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & oImpresora.gPrnSaltoLinea
        End If
    End If
    If psNomAgeRem = "" Then
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    Else
        lsCadena = lsCadena & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(sAux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & "Ag.Rem: " & Trim(psNomAgeRem) & oImpresora.gPrnSaltoLinea
    End If
    
    If psComCmac = "XXX" Then
        If psLin3 = "XXX" Then
            lsCadena = lsCadena & lsNegritaOff & oImpresora.gPrnSaltoLinea     'Desactiva Negrita
        Else
            lsCadena = lsCadena & Space(sIni) & psLin3 & Space(sAux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff & oImpresora.gPrnSaltoLinea      'Desactiva Negrita
            lnNumLinCmac = 1
        End If
        lnNumLinCmac = 0
    Else
        lsCadena = lsCadena & Space(sIni) & "NroDocCmac:" & psComCmac & Space(sAux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff & oImpresora.gPrnSaltoLinea      'Desactiva Negrita
        lnNumLinCmac = 1
    End If
    
    lsCadena = lsCadena & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(sAux + sSep - 6) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImpresora.gPrnSaltoLinea
    
    'psCliente = PstaNombre(psCliente)
    
    lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
    
    If lnCliAux = 0 Then
        If sAux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + sAux)
        lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(psCliente) & Space(sAux + sMax + sSep - Len(psCliente)) & ImpreCarEsp(psCliente) & oImpresora.gPrnSaltoLinea
    Else
        lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
        lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
        
        If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + sAux)
        If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + sAux)
        
        lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(lsCliAux1) & Space(sAux + sMax + sSep - Len(lsCliAux1)) & ImpreCarEsp(lsCliAux1) & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(lsCliAux2) & Space(sAux + sMax + sSep - Len(lsCliAux2)) & ImpreCarEsp(lsCliAux2) & oImpresora.gPrnSaltoLinea
        
        lnCliAux = 1
    End If
    
    If pbSaldoInt Or pbCuenta Then
        lsCadena = lsCadena & Space(sIni) & "Cuenta:" & psCodCta & Space(14 + sSep + sAux) & "Cuenta:" & psCodCta & oImpresora.gPrnSaltoLinea
    Else
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    End If
    
    psTit = Trim(psTit)
    psTit = CentrarCadena(psTit, 28)
    lsCadena = lsCadena & lsNegritaOn    'Activa Negrita
    lsCadena = lsCadena & Space(sIni + 1) & "-----" & psTit & "-----" & Space(-1 + sSep) & "-----" & psTit & "-----" & oImpresora.gPrnSaltoLinea
    
    
    lnChq = InStr(1, psText, "*", vbTextCompare)
    
    If psTexto = "XXX" Then
        If lnChq = 0 Then
            lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(Trim(psText)) & Space(sMax + 6 - Len(Trim(psText)) - Len(sMonto)) & sMonto & Space(-1 + sSep) & ImpreCarEsp(Trim(psText)) & Space(sMax + 6 - Len(Trim(psText)) - Len(sMonto)) & sMonto & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
        Else
            lsChqAux1 = (Mid(psText, 1, lnChq - 1))
            lsChqAux2 = (Mid(psText, lnChq + 1))
            lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Space(-1 + sSep) & ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Space(-1 + sSep) & ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & oImpresora.gPrnSaltoLinea
        End If
    Else
        lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(Trim(psTexto)) & Space(sAux + sSep + sMax - Len(Trim(psTexto))) & ImpreCarEsp(Trim(psTexto)) & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    End If
    
    lsCadena = lsCadena & lsNegritaOff    'Desactiva Negrita
    
    If pbSaldoInt Then
        If pbConfirmacion Then
            If MsgBox("Desea Imprimir el Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                pbConfirmacion = True
            Else
                pbConfirmacion = False
            End If
        Else
            pbConfirmacion = True
        End If
        If pbConfirmacion Then
            lsCadena = lsCadena & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & oImpresora.gPrnSaltoLinea
            If pbOpSaldoC Then
                lsCadena = lsCadena & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & oImpresora.gPrnSaltoLinea
            Else
                lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
            End If
        Else
            lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
            pbSaldoInt = False
        End If
    Else
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    End If
    
    lsInteres = pnInteresA
    
    If pbSaldoInt Then
        If lsInteres <> "No Valido" Then
            lsInteres = Format(lsInteres, "#,##0.00")
            lsCadena = lsCadena & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Space(-1 + sSep) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & oImpresora.gPrnSaltoLinea
        End If
    Else
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    End If
    
    If Not psNumDoc = "" Then
        lsCadena = lsCadena & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & oImpresora.gPrnSaltoLinea
    Else
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    End If
    
    If Not psNumDias = "----" Then
        lsCadena = lsCadena & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Space(-1 + sSep) & "Nro Dias Interes" & Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")) & psNumDias & oImpresora.gPrnSaltoLinea
        lnTope = 4 - lnCliAux
    Else
        lnTope = 3 - lnCliAux
    End If
    
    lsCadena = lsCadena & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & "---------------------------------------" & oImpresora.gPrnSaltoLinea
    If psCodUsuRem = "" Then
        lsCadena = lsCadena & Space(sIni) & ImpreCarEsp(gsCodUser) & Space(29 + sSep + sAux) & ImpreCarEsp(gsCodUser) & oImpresora.gPrnSaltoLinea
    Else
        lsCadena = lsCadena & Space(sIni) & "Loc/Rem" & Space(sMax + sAux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(gsCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & Space(sSep) & "Loc/Rem" & Space(sMax + sAux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(gsCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & oImpresora.gPrnSaltoLinea
    End If
    
    'lsCadena = lscadena &  oImpresora.gPrnSaltoPagina
    For sAux = 1 To (lnTope - lnNumLinCmac)
        lsCadena = lsCadena & "" & oImpresora.gPrnSaltoLinea
    Next sAux
    
    ImprimeBoletaCad = lsCadena
    
    Exit Function
ERROR:
        Close nFicSal
        If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
            GoTo ETIQ
        End If
End Function


Function Cargo(psCodCta As String, pnMonto As Currency, pgsACRetNC As String, pgsCodAge As String, pgsCodUser As String, NroDoc As String, psDes As String, poCon As DConecta, pgdFecSis As Date, Optional pnITFValor As Currency = 0) As Currency
    Dim sqlA As String
    Dim sqlT As String
    Dim sqlTra As String
    Dim sqlVerifica As String
    Dim sqlNota As String
    Dim lnInteres As Currency
    Dim RegCuenta As ADODB.Recordset
    Set RegCuenta = New ADODB.Recordset
    
    Dim lnSaldoDisponible As Currency
    Dim lnSaldCnt As Currency
    Dim lnSaldDisp As Currency
    Dim lbSAux As Boolean
    Dim lcEst As String
    Dim lNTInt As Double
    Dim lsHoy As String
    Dim lsAyer As String
    Dim I As Long
    Dim inicioTrans As Boolean
    Dim lsAgeTmp As String
    
    lsHoy = Format(pgdFecSis, "mm/dd/yyyy") + " " + Format(Time, "hh:mm:ss AMPM")
    lsAyer = Format(pgdFecSis - 1, "mm/dd/yyyy") + " " + Format(Time, "hh:mm:ss AMPM")
    
    lbSAux = False
     
    If Mid(psCodCta, 3, 3) = Producto.gCapAhorros Then 'Ahorro
        sqlA = "Select * from AhorroC where cCodCta = '" & psCodCta & "'"
        Set RegCuenta = poCon.CargaRecordSet(sqlA)
        
        lnInteres = RegCuenta!nInteres + GetInteres(RegCuenta!nTasaIntAC, RegCuenta!nsalddispac, DateDiff("d", RegCuenta!dUltMovAC, pgdFecSis) - 1, True)

        lnSaldCnt = RegCuenta!nSaldCntAC - CCur(pnMonto)
        lnSaldDisp = RegCuenta!nsalddispac - CCur(pnMonto)
        
        If RegCuenta!cTipTasAC = "T" Then
            lNTInt = GetNuevaTasaAC(psCodCta, lnSaldDisp, RegCuenta!nTasaIntAC, poCon)
        Else
            lNTInt = RegCuenta!nTasaIntAC
        End If
    
        lbSAux = False
        If RegCuenta!bInactiva Then
            lcEst = "0"
            lbSAux = True
            'Realiza la Transaccion de activación
            sqlT = " Insert TranDiaria (dFecTran, cCodUsu, cCodOpe, cCodCta, cNumDoc, nMonTran, nSaldCnt, cCodUsuRem, cCodAge, cFlag, nTipCambio) values('" & lsHoy & "' , '" & pgsCodUser & "' , '" _
                 & gsACInaAct & "' , '" & psCodCta & "' , null , " & Str(Round(lnSaldDisp, 2)) & " ," & Str(lnSaldCnt) & ", null , '" & pgsCodAge & "' , NULL,0)"
        Else
            lcEst = "0"
        End If

        'Actualiza los datos de la tabla de Ahorro
        sqlA = " Update AhorroC " _
             & " set nSaldAntAC=nSaldDispAC, nSaldDispAC=nSaldDispAC-" & Str(Round(CCur(pnMonto), 2)) & ", dUltCntAC = '" + lsHoy + "', bInactiva = " & lcEst & "  ," _
             & " nSaldCntAC=nSaldCntAC-" & Str(Round(CCur(pnMonto), 2)) & ", dUltMovAC='" & lsAyer & "', ntasaintAC = " + Str(lNTInt) + " " _
             & " , dUltActAC='" & lsHoy & "', cCodUsu='" & gsCodUser & "', nNumExt = nNumExt + 1 ,nInteres = " + Str(Round(lnInteres, 2)) _
             & " where cCodCta='" & psCodCta & "'"

        sqlTra = "Insert TranDiaria (dFecTran, cCodUsu, cCodOpe, cCodCta, cNumDoc, nMonTran, nSaldCnt, cCodUsuRem, cCodAge, cFlag, nTipCambio) values('" & lsHoy & "' , '" _
               & pgsCodUser & "', '" & pgsACRetNC & "','" & psCodCta & "','" & NroDoc & "', " _
               & "-" + Str(Round(CCur(pnMonto), 2)) & ", " & Str(Round(lnSaldCnt, 2)) & ",null,'" & pgsCodAge & "', Null,0)"
    
        sqlVerifica = " Select Max(nNumTran) nNumTran from TranDiaria Where dFecTran = '" & lsHoy & "' And cCodCta = '" & psCodCta & "' AND  cCodAge = '" & pgsCodAge & "' AND cCodUsu = '" & pgsCodUser & "' and cCodOpe = '" & pgsACRetNC & "'"
        
        'Inserta nueva operación en Transacción Diaria
        RegCuenta.Close
        Set RegCuenta = Nothing

        Randomize
        For I = 0 To Rnd(2000) * 1000
        Next I
        
        If Not TieneSaldo(psCodCta, Str(pnMonto), poCon) Then
            MsgBox "Se ha realizado otra operación de Retiro, la cuenta ya no tiene saldo suficiente", vbInformation, "Aviso"
            gbOpeOk = False
            Cargo = 0
            Exit Function
        End If
        
        poCon.Ejecutar sqlA
        poCon.Ejecutar sqlTra
        If gbITFAplica And pnITFValor > 0 Then
            lsAgeTmp = gsCodAge
            gsCodAge = pgsCodAge
            gnITFNumTranOrigen = fgITFGetNumtranOrigen(poCon.ConexionActiva)
            gnITFNumTran = fgITFACRetiroImpuesto(psCodCta, pnITFValor, "", gsITFACCargoCuenta, FechaHora(pgdFecSis), pgsCodUser, poCon.ConexionActiva)
            fgITFDetalleInserta gnITFNumTranOrigen, Right(pgsCodAge, 2), CCur(pnMonto), pnITFValor, gnITFNumTran, poCon.ConexionActiva
            gsCodAge = lsAgeTmp
        End If
        
        If lbSAux Then poCon.Ejecutar sqlT
        
        Set RegCuenta = poCon.CargaRecordSet(sqlVerifica)
        
        If RegCuenta.EOF And RegCuenta.BOF Then
            Set RegCuenta = Nothing
            GoTo ERROR
        End If
          
        sqlA = Str(RegCuenta!nNumTran)
        
        RegCuenta.Close
        Set RegCuenta = Nothing
                
        sqlNota = " INSERT OtrasOpe(dFecTran,nNumTran,cCodOpe,cGlosa,cNumDoc,cCodPers,cFlag) " _
                & " values('" & lsHoy & "'," & sqlA & ",'" & pgsACRetNC & "','" & psDes & "','" & NroDoc & "',null,null)"

        poCon.Ejecutar sqlNota
 
        gbOpeOk = True
        Cargo = lnSaldDisp
    
    End If 'fin Ahorro
    Exit Function
ERROR:
    
End Function


'*************************************************
'TieneSaldo: Funcion que devuelve el true si la cuenta tiene
'el saldo suficiente para realizar el retiro que se esta realizando
'se le ingresa como paramentro el Nro de cuenta el monto a retirar y
'conexion a la agencia que es propietaria de esta cuenta
'con que se esta haciendo un retiro con orden de pago
'caso contrario es False
'*************************************************
Public Function TieneSaldo(psCodCta As String, pnMonto As String, poCon As DConecta) As Boolean
    Dim RegCuenta As New ADODB.Recordset
    Dim lnMontoMinimo As Double
    Dim lnCuentaDescarAuto As Double
    Dim lnResta As Double
    Dim sql As String
    On Error GoTo ERROR
        
    Select Case Mid(psCodCta, 3, 3)
        Case Producto.gCapAhorros
                sql = "Select nSaldDispAC as nSaldDisp from AhorroC where cCodCta = '" & psCodCta & "' and cEstCtaAC Not in ('C','U') "
        Case Producto.gCapPlazoFijo
                sql = "Select nIntDev as nSaldDisp from PlazoFijo where cCodCta = '" & psCodCta & "' and cEstCtaPF Not in ('C','U')"
        Case Producto.gCapCTS
                sql = "Select nSaldDispCTS as nSaldDisp from CTS where cCodCta = '" & psCodCta & "' and cEstCtaCTS Not in ('C','U')"
    End Select
    
    Set RegCuenta = poCon.CargaRecordSet(sql)
    
    If RSVacio(RegCuenta) Then
        Set RegCuenta = Nothing
        TieneSaldo = False
        Exit Function
    End If
    
    Select Case Mid(psCodCta, 3, 3)
        Case Producto.gCapAhorros
            If Mid(psCodCta, 6, 1) = "1" Then
                lnMontoMinimo = ReadParametros("23102", poCon)
                lnResta = lnMontoMinimo
                If RegCuenta!nSaldDisp - lnResta >= CCur(pnMonto) Then
                    RegCuenta.Close
                    Set RegCuenta = Nothing
                    TieneSaldo = True
                Else
                    RegCuenta.Close
                    Set RegCuenta = Nothing
                    TieneSaldo = False
                End If
            Else
                lnMontoMinimo = ReadParametros("23103", poCon)
               
                lnResta = lnMontoMinimo
                If RegCuenta!nSaldDisp - lnResta >= CCur(pnMonto) Then
                    RegCuenta.Close
                    Set RegCuenta = Nothing
                    TieneSaldo = True
                Else
                    RegCuenta.Close
                    Set RegCuenta = Nothing
                    TieneSaldo = False
                End If
            End If
        
        Case Producto.gCapPlazoFijo
                If RegCuenta!nSaldDisp >= CCur(pnMonto) Then
                RegCuenta.Close
                Set RegCuenta = Nothing
                TieneSaldo = True
            Else
                RegCuenta.Close
                Set RegCuenta = Nothing
                TieneSaldo = False
            End If
        Case Producto.gCapCTS
            
            If (RegCuenta!nSaldDisp) >= CCur(pnMonto) Then
                RegCuenta.Close
                Set RegCuenta = Nothing
                TieneSaldo = True
            Else
                RegCuenta.Close
                Set RegCuenta = Nothing
                TieneSaldo = False
            End If
    End Select

    Exit Function
ERROR:
    MsgBox Err.Description & ". Funciones Principales", vbCritical, "Aviso"
End Function

'****************************************
'ReadParametros: Funcion que devuelve el valor de un parametro de
'la tabla parametros, se ingresa el codigo del parametro
'****************************************
Public Function ReadParametros(txtCodPar As String, poCon As DConecta) As Double

    Dim RecVar As New ADODB.Recordset
    Dim qryVar As String
    
    On Error GoTo ERROR
    
    qryVar = "SELECT nValor1 FROM Parametro WHERE cCodPar = '" & txtCodPar & "'"
    Set RecVar = poCon.CargaRecordSet(qryVar)
    
    If RecVar.EOF And RecVar.BOF Then
        ReadParametros = 0
        Set RecVar = Nothing
    Else
        ReadParametros = RecVar!nValor1
        RecVar.Close
        Set RecVar = Nothing
    End If
    
    Exit Function
ERROR:
    MsgBox "Error en Conexión + " + Err.Description, vbCritical, "Aviso"
    ReadParametros = -1
    
End Function

Public Function GetOpeAbonoPlanilla(psPlaCod As String) As String
    Dim oCon As DConecta
    Dim sql As String
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    
    sql = "Select cOpeAbono from rhplanillatabla where cRHPlanillaCod = '" & psPlaCod & "'"
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        GetOpeAbonoPlanilla = ""
    ElseIf IsNull(rs.Fields(0)) Then
        GetOpeAbonoPlanilla = ""
    Else
        GetOpeAbonoPlanilla = rs.Fields(0)
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    rs.Close
    Set rs = Nothing
End Function

'*** PEAC 20131112
Public Function ValidaCtasParaAbono(prRS As ADODB.Recordset, psPlanillaPeriodo As String, psPlanillaCod As String, psPlanillaDes As String, pnTipoCambioVenta As Currency, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, pgsCodAge As String, pgsCodUser As String) As String
    On Error GoTo AgregaConceptoEmpleadoErr

    Dim lsNotaAbono As String
    Dim lsNotaAbonoRef As String
    
    Dim lsCuentaRef As String
    
    Dim oConsistema As NConstSistemas
    Set oConsistema = New NConstSistemas
    Dim lnDispPorcent As Double
    
    lnDispPorcent = oConsistema.LeeConstSistema(222)
    
    Dim lsMonto As String
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lsCodCta As String
    Dim oPla As DRHProcesosCierre
    Set oPla = New DRHProcesosCierre
    Dim oExtra As DActualizaMovExtraPlanilla
    Set oExtra = New DActualizaMovExtraPlanilla
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    Dim rsOpe As ADODB.Recordset
    Set rsOpe = New ADODB.Recordset
    Dim oMov As DMov
    Set oMov = New DMov
    Dim oCapta As NCapMovimientos
    Set oCapta = New NCapMovimientos
    Dim lnSaldo As Double
    Dim lsCadBol As String
    Dim lsOpeAbonoCod As String
    Dim lsCabezal As String
    Dim lbAbonado As Boolean
    
    Dim rsDatosFractal As ADODB.Recordset
    Set rsDatosFractal = New ADODB.Recordset
    Dim lnIntangibleCTS As Currency

    Dim lsEstadoCuenta As String
    Dim lsTitularCta As String
    
    lsOpeAbonoCod = GetOpeAbonoPlanilla(psPlanillaCod)

    If lsOpeAbonoCod = "" Then
        MsgBox "No se ha realizado ningun abono, no existe una operacion Valida de Abono a las cuentas.", vbInformation, "Aviso"
        Exit Function
    End If
    
    gsNomAge = Trim(pgsNomAge)
    lsCadBol = ""
    lsCabezal = ""
    
    If Not (prRS.EOF And prRS.BOF) Then
        
        While Not prRS.EOF
            If prRS.Fields(0) = 1 Then
                
                lbAbonado = oRH.GetPlanillaPagada(prRS.Fields(2), psPlanillaCod, psPlanillaPeriodo)
                lsCodCta = oRH.GetCuentaRRHH(prRS.Fields(2), psPlanillaCod, psPlanillaPeriodo)
                lsEstadoCuenta = oRH.GetEstadoCuentaRRHH(lsCodCta)

                If lsEstadoCuenta = "1400" Or lsEstadoCuenta = "1300" Or lsEstadoCuenta = "" Then
                    lsCadBol = lsCadBol & lsCodCta & " " & Left(prRS.Fields("Nombre") + Space(50), 50) & " " & IIf(lsEstadoCuenta = "1400", "CUENTA CANCELADA", IIf(lsEstadoCuenta = "1300", "CUENTA ANULADA", "CUENTA NO EXISTE")) & oImpresora.gPrnSaltoLinea
                End If
                
            End If
            prRS.MoveNext
        Wend
    End If
    
    If lsCadBol <> "" Then
    lsCabezal = CStr(Date) + " " + CStr(Time()) & Space(70) & "CAJA MAYNAS S.A." & oImpresora.gPrnSaltoLinea
    lsCabezal = lsCabezal & "INCONSISTENCIAS ENCONTRADAS AL MOMENTO DE REALIZAR EL ABONO" & oImpresora.gPrnSaltoLinea
    lsCabezal = lsCabezal & "===========================================================" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCabezal = lsCabezal & "NUM. CUENTA         TITULAR                                           OBSERVACION" & oImpresora.gPrnSaltoLinea
    lsCabezal = lsCabezal & "=================================================================================" & oImpresora.gPrnSaltoLinea
    End If
    
    ValidaCtasParaAbono = IIf(lsCadBol <> "", lsCabezal & lsCadBol, "")
    Exit Function
AgregaConceptoEmpleadoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoEmpleado:AgregaConceptoEmpleado Method")
End Function
'APRI 20170328
Private Sub GeneraLogAbonoRH(ByVal psTexto As String, Optional ByVal pPlanilla As String)
    Dim FechaAbono As String
    FechaAbono = Replace(pPlanilla, " ", "") & Format(gdFecSis, "yyyymmdd")
    'C:\sicmacm\DLLComunes\Spooler
    Open App.Path & "\Spooler\" & FechaAbono & ".TXT" For Append As #1
        Print #1, psTexto
    Close #1
End Sub
'END APRI
