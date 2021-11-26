VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPigContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***************************************************
' CMPCL - Registro de un Contrato
' CAFF  - 20/08/2002
'***************************************************
Public Function nRegistraContratoPigno(ByVal psAgencia As String, ByVal pnMoneda As Moneda, ByVal pnTasaComp As Double, _
        ByVal pnTasaMora As Double, ByVal pnSaldo As Currency, ByVal psFechaHora As String, ByVal psCodPersona As String, _
        ByVal pnPlazo As Integer, ByVal psFecVenc As String, ByVal psMovNro As String, ByVal pnCostoTasac As Double, _
        ByVal prJoyas As Recordset, ByVal pnPiezas As Integer, ByVal pnUbica As ColocPigUbica, ByVal psCodPers As String, _
        ByVal psCodLin As String, ByVal pnIntComp As Currency, ByVal pnMontoTasac As Currency, ByVal pnPenalidad As Currency, _
        ByVal pnMontoPrestamo As Currency, ByVal pnTotalBalanza As Currency, ByVal pnCalCte As Integer, ByVal pnEvalCte As Integer)
        
Dim oGen As DGeneral
Dim lsCuenta As String
Dim lnMovNro As Long
Dim oRegPig As DPigActualizaBD
Dim mbTrans As Boolean
Dim oPigFunc As DPigFunciones
Dim lsAgeBovedaVal As String

On Error GoTo ErrorRegPig

Set oPigFunc = New DPigFunciones
lsAgeBovedaVal = CStr(oPigFunc.GetParamValor(gPigAgeBovedaVal))
Set oPigFunc = Nothing

'Genera Nro Cuenta
Set oGen = New DGeneral
    lsCuenta = oGen.GeneraNuevaCuenta(psAgencia, gColConsuPrendario, pnMoneda)
Set oGen = Nothing

Set oRegPig = New DPigActualizaBD

    oRegPig.dBeginTrans
    mbTrans = True
    
    Call oRegPig.dInsertProducto(lsCuenta, pnTasaComp, pnMontoPrestamo, gPigEstRegis, psFechaHora, 0, False)
  
    Call oRegPig.dInsertProductoPersona(lsCuenta, psCodPersona, gColRelPersTitular, False)
    
    If pnCalCte = 4 Then 'Cliente Nuevo
        Call oRegPig.dInsertColocPigEvalCliente(psCodPersona, pnCalCte, pnEvalCte, False)
    End If
    
    'Aca se ingresan dos registros uno para la Tasa Compensatoria y la otra para la Tasa Moratoria
    Call oRegPig.dInsertProductoTasaInteres(lsCuenta, gColocTasaCompNormal, pnTasaComp, False)
    Call oRegPig.dInsertProductoTasaInteres(lsCuenta, gColocTasaMoraNormal, pnTasaMora, False)
     
    Call oRegPig.dInsertColocaciones(lsCuenta, pnPlazo, psFecVenc, pnMontoPrestamo, gColocCalendCodPFCF, psMovNro, psCodLin, psFechaHora, False)
    
    Call oRegPig.dInsertColocacEstado(lsCuenta, psFechaHora, gPigEstRegis, 1, pnMontoPrestamo, "Registro Contrato", gColocCalendCodPFCF, 0, 0, _
            pnPlazo, "0", 0, 0, 0, 0, False)
          
    '======= Calendario para el Desembolso
    Call oRegPig.dInsertColocCalendario(lsCuenta, 1, gColocCalendAplDesembolso, 1, psFecVenc, 0, "Registro Contrato", "", False)
    
   '********** Detalle para el Desembolso
    'Inserta ColocCalendDetPig - Capital
    Call oRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplDesembolso, 1, gColPigConceptoCodCapital, pnMontoPrestamo, 0, "", False)
    
    'Inserta ColocCalendDetPig - Comision de Tasacion
    If pnCostoTasac >= 0 Then
        Call oRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplDesembolso, 1, gColPigConceptoCodTasacion, pnCostoTasac, 0, "", False)
    End If
          
    'Calendario para la Cuota
    Call oRegPig.dInsertColocCalendario(lsCuenta, 1, gColocCalendAplCuota, 1, psFecVenc, 0, "Registro Contrato", "", False)
          
   '********** Detalle para la cuota
    'Inserta ColocCalendDetPig - Capital
    Call oRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplCuota, 1, gColPigConceptoCodCapital, pnMontoPrestamo, 0, "", False)
    
    'Inserta ColocCalendDetPig - Interes Compensatorio
    Call oRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplCuota, 1, gColPigConceptoCodInteresComp, pnIntComp, 0, "", False)
    
    'Inserta ColocCalendDetPig - Interes Moratorio
    Call oRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplCuota, 1, gcolPigConceptoCodInteresMora, 0, 0, "", False)
    
    '========== Inserta en el Mov
    Call oRegPig.dInsertMov(psMovNro, gPigOpeRegContrato, "Registro Contrato", gMovEstContabMovContable, gMovFlagVigente, False)
    
    'Obtiene nMovNro
    lnMovNro = oRegPig.dGetnMovNro(psMovNro)
        
    'Inserta MovCol
    Call oRegPig.dInsertMovCol(lnMovNro, gPigOpeRegContrato, lsCuenta, 1, pnMontoPrestamo, 0, "GMIC", pnPlazo, gPigEstRegis, pnMontoPrestamo, False)
 
    'Valor Tasacion  ==  PARA EL TEMA DE LA CUSTODIA
    '(Cuando es la Agencia de la Boveda de Valores lo envia directamente a la Boveda de Valores
    If Right(lsAgeBovedaVal, 2) = Right(psAgencia, 2) Then
        Call oRegPig.dInsertMovColDet(lnMovNro, gPigOpeRegContrato, lsCuenta, 1, gColPigConceptoCodGarantCustBoveda, 1, pnMontoTasac, False)
        pnUbica = gPigUbicaBovValor
    Else
        Call oRegPig.dInsertMovColDet(lnMovNro, gPigOpeRegContrato, lsCuenta, 1, gColPigConceptoCodGarantCustAge, 1, pnMontoTasac, False)
    End If
        
    'Inserta en ColocPigno los datos del Contrato

    Call oRegPig.dInsertColocPigno(lsCuenta, 0, pnPiezas, gPigEstRegis, pnUbica, 1, gPigTipoTasacNor, 0, 0, pnTotalBalanza, 1, False)
   
    'Inserta el detalle de las Joyas
    If Not prJoyas.EOF And Not prJoyas.BOF Then
        prJoyas.MoveFirst
        Do While Not prJoyas.EOF
            Call oRegPig.dInsertColocPigJoyaTasacion(lsCuenta, gPigTipoTasacNor, CInt(prJoyas!Item), CInt(Right(prJoyas!Material, 3)), _
                        CInt(Right(prJoyas!Tipo, 3)), CInt(IIf(Right(prJoyas!SubTipo, 3) = "", 0, Right(prJoyas!SubTipo, 3))), CInt(Right(prJoyas!Estado, 3)), _
                        CDbl(IIf(prJoyas!PBruto = "", 0, prJoyas!PBruto)), CDbl(IIf(prJoyas!PNeto = "", 0, prJoyas!PNeto)), _
                        CCur(IIf(prJoyas!Tasacion = "", 0, prJoyas!Tasacion)), CCur(IIf(prJoyas!TasAdicion = "", 0, prJoyas!TasAdicion)), _
                        CStr(prJoyas!Observacion), CStr(prJoyas!ObsAdicion), pnUbica, psCodPers, psMovNro, False)
            prJoyas.MoveNext
        Loop
    End If

    nRegistraContratoPigno = lsCuenta
  
    oRegPig.dCommitTrans
    Set oRegPig = Nothing
    mbTrans = False
    
Exit Function
    
ErrorRegPig:
    If mbTrans Then
        oRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Function

'****************************************************************
'**  CRSF -  28/08/02- Modificacion de Grabacion de Desembolso
'****************************************************************
Public Function nDesembolsoCredPignoraticio(ByVal psCtaCod As String, ByVal pnSaldoCap As Currency, _
            ByVal psFechaHora As String, ByVal psMovNro As String, ByVal pnMontoEntregar As Currency, _
            ByVal pnComision As Currency, ByVal pnImpuesto As Currency, ByVal pnTasacion As Currency, _
            ByVal pnPlazo As Integer, Optional sPersLavDinero As String = "", Optional sPersCod As String = "") As Long

Dim lsSQL As String
Dim loRegPigno As DPigActualizaBD 'Nuevo CMCPL
Dim lnMovNro As Long
Dim mbTrans As Boolean

Set loRegPigno = New DPigActualizaBD
     loRegPigno.dBeginTrans
     mbTrans = True
        '** Actualiza Producto - CMCPL
    Call loRegPigno.dUpdateProducto(psCtaCod, -1, -1, gPigEstDesemb, psFechaHora, -2, False) ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones - CMCPL
     Call loRegPigno.dUpdateColocaciones(psCtaCod, -1, "@", -1, -1, psMovNro, "@", psFechaHora, False)
    
    '** Insert ColocEstado - CMCPL
     Call loRegPigno.dInsertColocacEstado(psCtaCod, psFechaHora, gPigEstDesemb, 1, pnSaldoCap, "Desembolso Contrato", _
        gColocCalendCodFFCF, 0, 0, pnPlazo, "0", 0, 0, 0, 0, False)
  
    '** Actualiza Calendario ColocCalendario  - CMCPL
    Call loRegPigno.dUpdateColocCalendario(psCtaCod, 1, gColocCalendAplDesembolso, gColocCalendEstadoPagado, , , , False)
    
    '** Actualiza Calendario Detalle ColocCalendDetPig - CMCPL  (MONTO DESEMBOLSADO)
    Call loRegPigno.dUpdateColocCalendDetPig(psCtaCod, 1, gColocCalendAplDesembolso, gColPigConceptoCodCapital, -1, pnSaldoCap, -1, False)
   
    '** Actualiza Calendario Detalle ColocCalendDetPig - CMCPL  (COMISION DE SERVICIO)
    Call loRegPigno.dUpdateColocCalendDetPig(psCtaCod, 1, gColocCalendAplDesembolso, gColPigConceptoCodTasacion, -1, pnComision, -1, False)
   
   '** Actualiza ColocPigno - CMCPL
   Call loRegPigno.dUpdateColocPigno(psCtaCod, , gPigEstDesemb, , , , , , False)
    
   '** Actualiza MOV  - CMCPL
   Call loRegPigno.dInsertMov(psMovNro, gPigOpeDesembolsoEFE, "Desembolso Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
     
   '** Obtiene nMovNro
   lnMovNro = loRegPigno.dGetnMovNro(psMovNro)

 ' **************************************************
    ' INSERTA DE LAVADO DE DINERO
    Dim clsMant As DCapMovimientos
    Set clsMant = New DCapMovimientos
    If sPersLavDinero <> "" Then
          clsMant.AgregaMovLavDinero lnMovNro, sPersCod
      End If
    Set clsMant = Nothing
  ' **************************************************

   '** Actualiza MOVCOL - CMCPL
   Call loRegPigno.dInsertMovCol(lnMovNro, gPigOpeDesembolsoEFE, psCtaCod, 1, pnSaldoCap - pnComision, 0, "", 0, gPigEstDesemb, pnSaldoCap, False)

    '** Inserta MovColDet -  Monto Capital
    Call loRegPigno.dInsertMovColDet(lnMovNro, gPigOpeDesembolsoEFE, psCtaCod, 1, gColPigConceptoCodCapital, 1, pnSaldoCap, False)

    '** Inserta MovColDet - Garantia (Valor de Tasacion del Contrato)
    Call loRegPigno.dInsertMovColDet(lnMovNro, gPigOpeDesembolsoEFE, psCtaCod, 1, gColPigConceptoCodGarantia, 1, pnTasacion, False)
    
    '** Inserta MovColDet -  Comision
    Call loRegPigno.dInsertMovColDet(lnMovNro, gPigOpeDesembolsoEFE, psCtaCod, 1, gColPigConceptoCodTasacion, 1, 0 - pnComision, False)
    
    loRegPigno.dCommitTrans
    mbTrans = False
    Set loRegPigno = Nothing
    nDesembolsoCredPignoraticio = lnMovNro
    
Exit Function

ErrorModPig:
    If mbTrans Then
        loRegPigno.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "
End Function

'***************************************************
' CMPCL - Anulacion de Contratos
' CAFF  - 20/08/2002
'***************************************************
Public Sub nAnulaCredPignoraticio(ByVal psCtaCod As String, ByVal pnEstado As ColocEstado, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnPrestamo As Currency, _
        Optional pbEjecBatch As Boolean = False)

Dim oRegPig As DPigActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim mbTrans As Boolean

lsOpeCod = gPigOpeAnulacion

'On Error GoTo ErrorModPig
Set oRegPig = New DPigActualizaBD
    oRegPig.dBeginTrans
    mbTrans = True
    
    Call oRegPig.dUpdateProducto(psCtaCod, , , pnEstado, psFechaHora, -2, False)
    
    Call oRegPig.dUpdateColocaciones(psCtaCod, , , , , psMovNro, , , False)

    Call oRegPig.dInsertColocacEstado(psCtaCod, psFechaHora, gPigEstAnula, 1, 0, "Anula Contrato", _
        gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)

    Call oRegPig.dInsertMov(psMovNro, lsOpeCod, "Anula Contrato Pigno", gMovEstContabNoContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = oRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol - Con el estado antes del movimiento
    Call oRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, 0, 0, "", 0, gPigEstRegis, pnPrestamo, False)
    
    '** Update Mov Anterior
    Call oRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)
    
    Call oRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    oRegPig.dCommitTrans
    mbTrans = False
Set oRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        oRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nAnulaCredPignoraticio", "Error en Funcion de Anulación de Contrato "

End Sub

'***************************************************
' CMPCL - Bloqueo y Desbloqueo de Contratos
' CAFF  - 20/08/2002
'***************************************************
Public Sub nBloqueoDesBloqueoCredPignoraticio(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
      ByVal pbBloqIni As Boolean, ByVal pnMotBloqueo As Integer, ByVal pnEstado As Integer, ByVal pnSaldo As Currency, _
      Optional ByVal psDescripcion As String = "@", Optional ByVal psMovNroBloqueo As String = "@", Optional pbEjecBatch As Boolean = False)

Dim oRegPig As DPigActualizaBD
Dim lnMovNro As Long
Dim mbTrans As Boolean

Set oRegPig = New DPigActualizaBD
    oRegPig.dBeginTrans
    mbTrans = True
    
    Call oRegPig.dUpdateProducto(psCtaCod, , , , psFechaHora, -2, False)  ' (-2) aumenta el ste transac
    
    Call oRegPig.dUpdateColocaciones(psCtaCod, , , , , psMovNro, , , False)

    If pbBloqIni = False Then ' Se va ha Bloquear
        
        Call oRegPig.dInsertProductoBloqueos(psCtaCod, 11, pnMotBloqueo, psMovNro, psDescripcion, , False)
       
        Call oRegPig.dInsertMov(psMovNro, gPigOpeBloqueoJoyas, "Bloqueo de Joyas", gMovEstContabNoContable, gMovFlagVigente, False)
        
         'Obtengo el Numero de Movimiento
        lnMovNro = oRegPig.dGetnMovNro(psMovNro)
        
        Call oRegPig.dInsertMovCol(lnMovNro, gPigOpeBloqueoJoyas, psCtaCod, 0, 0, 0, "", 0, pnEstado, pnSaldo, False)
    Else ' Se va ha Desbloquear
        
        Call oRegPig.dUpdateProductoBloqueos(psCtaCod, psMovNroBloqueo, psMovNro, False)
        
        Call oRegPig.dInsertMov(psMovNro, gPigOpeDesbloqueoJoyas, "DesBloqueo de Joyas", gMovEstContabNoContable, gMovFlagVigente, False)
        
        'Obtengo el Numero de Movimiento
        lnMovNro = oRegPig.dGetnMovNro(psMovNro)
        
        Call oRegPig.dInsertMovCol(lnMovNro, gPigOpeDesbloqueoJoyas, psCtaCod, 0, 0, 0, "", 0, pnEstado, pnSaldo, False)
    End If
    
    oRegPig.dCommitTrans
Set oRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        oRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nBloqueoDesbloqueoCredPignoraticio", "Error en Funcion de Bloqueo / Desbloqueo de Contrato "
End Sub

'***************************************************
' CMPCL - Duplicado de Contratos
' CAFF  - 20/08/2002
'***************************************************

Public Sub nDuplicadoContratoCredPignoraticio(ByVal psCtaCod As String, ByVal pnNroDuplic As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnMontoTransac As Currency, ByVal pnEstado As Integer, Optional pbEjecBatch As Boolean = False)

Dim oRegPig As DPigActualizaBD
Dim lnMovNro As Long
Dim mbTrans As Boolean

'On Error GoTo ErrorModPig
Set oRegPig = New DPigActualizaBD
    oRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call oRegPig.dUpdateProducto(psCtaCod, , , , psFechaHora, -2, False)           ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call oRegPig.dUpdateColocaciones(psCtaCod, , , , , psMovNro, , , False)

    '** Actualiza ColocPignoraticio
    Call oRegPig.dUpdateColocPigno(psCtaCod, , , , , , pnNroDuplic + 1, , False)

    '** Inserta Mov
    Call oRegPig.dInsertMov(psMovNro, gPigOpeImpDuplicado, "Duplicado Credito Pigno", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = oRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call oRegPig.dInsertMovCol(lnMovNro, gPigOpeImpDuplicado, psCtaCod, 1, pnMontoTransac, 0, "", 0, pnEstado, 0, False)

    '** Inserta MovColDet -  Costo Duplicado
    If pnMontoTransac > 0 Then
        Call oRegPig.dInsertMovColDet(lnMovNro, gPigOpeImpDuplicado, psCtaCod, 1, gColPigConceptoCodDuplicado, 1, pnMontoTransac, False)
    End If

    oRegPig.dCommitTrans
    mbTrans = False
    
Set oRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        oRegPig.dRollbackTrans
        mbTrans = False
    End If
    
    Err.Raise vbObjectError + 100, "Error nDuplicadoContratoPignoraticio", "Error en Funcion Duplicado de Contrato "

End Sub

'***************************************************************************
'CAFF CMCPL MODIFICACION DE CONTRATO
'***************************************************************************
Public Sub nModificaCredPignoraticio(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal pnEstado As Integer, _
        ByVal prJoyas As Recordset, ByVal psMovNro As String, Optional ByVal pnTipoTasac As Integer, _
        Optional pbEjecBatch As Boolean = False)

Dim oRegPig As DPigActualizaBD
Dim lnMovNro As Long
Dim lnTipo As Integer
Dim lnSTipo As Integer
Dim lsDes As String
Dim lsDesAdic As String
Dim lnItem As Integer
Dim mbTrans As Boolean

'On Error GoTo ErrorModPig
Set oRegPig = New DPigActualizaBD

    oRegPig.dBeginTrans
    mbTrans = True
    '** Actualiza Producto
    Call oRegPig.dUpdateProducto(psCtaCod, , , , , -2, False) ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call oRegPig.dUpdateColocaciones(psCtaCod, , , , , psMovNro, , , False)
    
    '**** Actualiza ColocPigJoyaTasacion
    Do While Not prJoyas.EOF
        lnItem = prJoyas!Item
        lnTipo = Val(Right(prJoyas!Tipo, 3))
        lnSTipo = Val(Right(prJoyas!SubTipo, 3))
        lsDes = Trim(prJoyas!Observacion)
        lsDesAdic = Trim(prJoyas!ObsAdicion)
        If pnTipoTasac = 0 Then
            pnTipoTasac = 1
        End If
        Call oRegPig.dUpdateColocPigJoyaTasacion(psCtaCod, pnTipoTasac, lnItem, , , lnTipo, lnSTipo, , , , , , lsDes, lsDesAdic, , , False)
        
        prJoyas.MoveNext
    Loop
    
    '** Inserta Mov
    Call oRegPig.dInsertMov(psMovNro, gPigOpeModifDescJoyas, "Modificacion Contrato", gMovEstContabNoContable, gMovFlagVigente, False)
    
    '** Obtiene nMovNro
    lnMovNro = oRegPig.dGetnMovNro(psMovNro)

    '** Inserta MovCol
    Call oRegPig.dInsertMovCol(lnMovNro, gPigOpeModifDescJoyas, psCtaCod, 0, 0, 0, "", 0, pnEstado, 0, False)


    oRegPig.dCommitTrans
    mbTrans = False
    
Set oRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        oRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Function nAmortizacionCredPignoraticio(ByVal psCtaCod As String, ByVal pnNewSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal psFecVenci As String, ByVal pnNewPlazo As Integer, _
        ByVal pnMontoTransac As Currency, ByVal pnCapPag As Currency, ByVal pnIntCompensatorio As Currency, _
        ByVal pnIntMoratorio As Currency, ByVal pnComServicio As Currency, ByVal pnComPenalidad As Currency, _
        ByVal pnComVencida As Currency, ByVal pnDerRemate As Currency, ByVal pnDiasAtraso As Integer, _
        ByVal pnNroAmort As Integer, ByVal pnDiasCambCart As Integer, ByVal pnValorTasacion As Currency, _
        ByVal pnOpeCod As Long, pnNewIntComp As Currency, ByVal psOpeDesc As String, ByVal psPersCodCMAC As String, _
        ByVal pnNroCalend As Integer, ByVal pnPlazo As Integer, ByVal pnEstado As Integer, ByVal pnSaldoCap As Currency, _
        Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal sPersLavDinero As String = "", _
        Optional ByVal sPersCod As String = "") As Long

'** Actualiza Producto (Estado, nSaldoCap,dPrdEstado, NroTrans)

'** Actualiza Colocaciones (Plazo, Vencimiento, Ultima Modificacion, dVigencia)
'** Actualiza ColocPigno (NroRenov)
'** Actualiza ColocCalendario (nColocCalendEstado)
'** Actualiza ColocCalendDetPig (nMontoPagado)

'** dInsertColocCalendario, ColocCalendDetPig
'** dInsertColocacEstado
'** dInsertMov, MovCol, MovColDet
'************************************

Dim lsSQL As String
Dim loRegPig As DPigActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim mbTrans As Boolean
  
Select Case pnOpeCod
    Case gPigOpeAmortizEFE
        If pnDiasAtraso > pnDiasCambCart Then
            lsOpeCod = gPigOpeAmortMorEFE
        Else
            lsOpeCod = gPigOpeAmortNorEFE
        End If
End Select

'On Error GoTo ErrorModPig
Set loRegPig = New DPigActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtaCod, , pnNewSaldoCap, gPigEstAmortiz, psFechaHora, -2, False)          ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtaCod, pnNewPlazo, psFecVenci, , , psMovNro, , , False)

    '** Actualiza ColocPigno
    Call loRegPig.dUpdateColocPigno(psCtaCod, , gPigEstAmortiz, , pnNroCalend + 1, , , pnNroAmort, False)

    Call loRegPig.dUpdateColocCalendario(psCtaCod, pnNroCalend, gColocCalendAplCuota, 1, , , psFechaHora)
    
    '** Insert ColocCalendario         (Se genera un nuevo Calendario)
    Call loRegPig.dInsertColocCalendario(psCtaCod, pnNroCalend + 1, gColocCalendAplCuota, 1, psFecVenci, _
                                                         gColocCalendEstadoPendiente, "Nuevo Calendario de Pago", "", gColocCalendConceptoProcOtro, False)
    
    '** Insert ColocCalendDetPig     (Se genera un nuevo Calendario)
    Call loRegPig.dInsertColocCalendDetPig(psCtaCod, pnNroCalend + 1, gColocCalendAplCuota, 1, gColPigConceptoCodCapital, _
                                                             pnNewSaldoCap, 0, "", False)
    Call loRegPig.dInsertColocCalendDetPig(psCtaCod, pnNroCalend + 1, gColocCalendAplCuota, 1, gColPigConceptoCodInteresComp, _
                                                             pnNewIntComp, 0, "", False)
    Call loRegPig.dInsertColocCalendDetPig(psCtaCod, pnNroCalend + 1, gColocCalendAplCuota, 1, gcolPigConceptoCodInteresMora, _
                                                             0, 0, "", False)
    
    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtaCod, psFechaHora, gPigEstAmortiz, 1, pnMontoTransac, "Amortización de Cred. Pignoraticio", _
                                                       gColocCalendCodFFCF, 0, 0, pnNewPlazo, "0", -1, -1, -1, -1, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Amortización de Cred. Pignoraticio", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
   ' **************************************************
    ' INSERTA DE LAVADO DE DINERO
    Dim clsMant As DCapMovimientos
    Set clsMant = New DCapMovimientos
    If sPersLavDinero <> "" Then
          clsMant.AgregaMovLavDinero lnMovNro, sPersCod
      End If
    Set clsMant = Nothing
  ' **************************************************
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, pnMontoTransac, pnDiasAtraso, "", pnPlazo, pnEstado, pnSaldoCap, False)

    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Capital
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gColPigConceptoCodCapital, , pnCapPag, , False)
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodCapital, 1, pnCapPag, False)
    
    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Interes Compensatorio
    If pnIntCompensatorio > 0 Then
       Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gColPigConceptoCodInteresComp, pnIntCompensatorio, pnIntCompensatorio, , False)
       Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodInteresComp, 1, pnIntCompensatorio, False)
    End If
    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Interes Moratorio
    If pnIntMoratorio > 0 Then
       Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gcolPigConceptoCodInteresMora, , pnIntMoratorio, , False)
       Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gcolPigConceptoCodInteresMora, 1, pnIntMoratorio, False)
    End If
    '** Inserta MovColDet - ColocCalendDetPig (Se actualiza el Actual) - Comision Servicio
    If pnComServicio > 0 Then
       Call loRegPig.dInsertColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, 1, gColPigConceptoCodComiServ, _
                                                                pnComServicio, pnComServicio, "", False)
       Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodComiServ, 1, pnComServicio, False)
    End If

    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Comision Vencida
    If pnComVencida > 0 Then
       Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gColPigConceptoCodComiVencida, , pnComVencida, , False)
       Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodComiVencida, 1, pnComVencida, False)
    End If
    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Derecho Remate
    If pnDerRemate > 0 Then
       Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gColPigConceptoCodPreparaRemate, , pnDerRemate, , False)
       Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodPreparaRemate, 1, pnDerRemate, False)
    End If
    
    '**** Cambio de Cartera - Regresa a Normal
'    If pnDiasAtraso > pnDiasCambCart Then
'        'Cambio de Cartera - Regresa a Normal
'        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodCambioCartNormalVenc, 1, pnNewSaldoCap, False)
'    End If
    
    nAmortizacionCredPignoraticio = lnMovNro
    
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Function
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Function
Public Function nCancelacionCredPignoraticio(ByVal psCtaCod As String, ByVal pnNewSaldoCap As Currency, _
               ByVal psFechaHora As String, ByVal psMovNro As String, ByVal pnMontoTransac As Currency, _
               ByVal pnCapPag As Currency, ByVal pnIntCompensatorio As Currency, ByVal pnIntMoratorio As Currency, _
               ByVal pnComServicio As Currency, ByVal pnComPenalidad As Currency, ByVal pnComVencida As Currency, _
               ByVal pnDerRemate As Currency, ByVal pnDiasAtraso As Integer, ByVal pnValorTasacion As Currency, _
               ByVal pnOpeCod As Long, ByVal psOpeDesc As String, ByVal psPersCodCMAC As String, _
               ByVal pnNroCalend As Integer, ByVal pnPlazo As Integer, ByVal pnEstado As Integer, ByVal pnSaldoCap As Currency, _
               ByVal pnDiasCambCart As Integer, Optional ByVal sPersLavDinero As String = "", _
               Optional ByVal sPersCod As String = "") As Long

Dim lsSQL As String
Dim loRegPig As DPigActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim mbTrans As Boolean

'Codigo de Operacion
Select Case pnOpeCod
    Case gPigOpeCancelacEFE
        If pnDiasAtraso > pnDiasCambCart Then
            lsOpeCod = gPigOpeCancelMorEFE
        Else
            lsOpeCod = gPigOpeCancelNorEFE
        End If
End Select

'On Error GoTo ErrorModPig
Set loRegPig = New DPigActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtaCod, , pnNewSaldoCap, gPigEstCancelPendRes, psFechaHora, -2, False)           ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtaCod, , , , , psMovNro, , , False)

    Call loRegPig.dUpdateColocPigno(psCtaCod, , gPigEstCancelPendRes, , , , , , False)
    
    Call loRegPig.dUpdateColocCalendario(psCtaCod, pnNroCalend, gColocCalendAplCuota, 1, , , psFechaHora)
    
    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtaCod, psFechaHora, gPigEstCancelPendRes, 1, pnMontoTransac, "Cancelacion Contrato", _
        gColocCalendCodFFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Cancelacion Cred.Pignoraticio", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    ' **************************************************
    ' INSERTA DE LAVADO DE DINERO
    Dim clsMant As DCapMovimientos
    Set clsMant = New DCapMovimientos
    If sPersLavDinero <> "" Then
          clsMant.AgregaMovLavDinero lnMovNro, sPersCod
      End If
    Set clsMant = Nothing
  ' **************************************************
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, pnMontoTransac, pnDiasAtraso, "", pnPlazo, pnEstado, pnSaldoCap, False)
    
    '** Inserta MovColDet -  Descargo de la Garantia en Agencia al momento de la cancelación
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodGarantia, 1, pnValorTasacion, False)

    '** Inserta MovColDet -  Movimiento de Valores de Lotes en Custodia a Custodia Pendiente de Rescate
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodGarantCustBoveda, 1, pnValorTasacion, False)
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodGarantPendRescBoveda, 1, pnValorTasacion, False)
    
    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Capital
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodCapital, 1, pnCapPag, False)
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gColPigConceptoCodCapital, , pnCapPag, , False)

    '** Actualiza ColocCalendDetPig (Se actualiza el Actual) / Inserta MovColDet - Interes Compensatorio
    If pnIntCompensatorio > 0 Then
       Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gColPigConceptoCodInteresComp, pnIntCompensatorio, pnIntCompensatorio, , False)
       Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodInteresComp, 1, pnIntCompensatorio, False)
    End If
    
    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Interes Moratorio
    If pnIntMoratorio > 0 Then
       Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gcolPigConceptoCodInteresMora, 1, pnIntMoratorio, False)
       Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gcolPigConceptoCodInteresMora, , pnIntMoratorio, , False)
    End If
    
    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Comision Servicio
    If pnComServicio > 0 Then
       Call loRegPig.dInsertColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, 1, gColPigConceptoCodComiServ, _
                                                                pnComServicio, pnComServicio, "", False)
       Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodComiServ, 1, pnComServicio, False)
    End If
    
    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Comision Penalidad
    If pnComPenalidad > 0 Then
       Call loRegPig.dInsertColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, 1, gColPigConceptoCodPenalidad, _
                                                                pnComPenalidad, pnComPenalidad, "", False)
       Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodPenalidad, 1, pnComPenalidad, False)
    End If
    
    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Comision Vencida
    If pnComVencida > 0 Then
       Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodComiVencida, 1, pnComVencida, False)
       Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gColPigConceptoCodComiVencida, , pnComVencida, , False)
    End If
    
    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Derecho Remate
    If pnDerRemate > 0 Then
       Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodPreparaRemate, 1, pnDerRemate, False)
       Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gColPigConceptoCodPreparaRemate, , pnDerRemate, , False)
    End If
    
    nCancelacionCredPignoraticio = lnMovNro
    
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Function
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCredPignoraticio", "Error en Funcion de Registro de Contrato "

End Function

'***************************************************************************
'CMCPL USO DE LINEA 'EAFA 'MODIFICACION CAFF
'***************************************************************************

'***************************************************************************
'CMCPL USO DE LINEA 'EAFA
'MODIFICACION CAFF
'EAFA 19/06/2003
'***************************************************************************

Public Function nUsoLineaCredPignoraticio(ByVal psCtaCod As String, ByVal pnNewSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal psFecVenci As String, ByVal pnNewPlazo As Integer, _
        ByVal pnMontoTransac As Currency, ByVal pnCapPag As Currency, ByVal pnIntCompensatorio As Currency, _
        ByVal pnIntMoratorio As Currency, ByVal pnComServicio As Currency, ByVal pnComPenalidad As Currency, _
        ByVal pnComVencida As Currency, ByVal pnDerRemate As Currency, ByVal pnDiasAtraso As Integer, _
        ByVal pnNroAmort As Integer, ByVal pnDiasCambCart As Integer, ByVal pnValorTasacion As Currency, _
        ByVal pnOpeCod As Long, ByVal pnNewIntComp As Currency, ByVal psOpeDesc As String, _
        ByVal psPersCodCMAC As String, ByVal pnNroCalend As Integer, ByVal pnPlazo As Integer, _
        ByVal pnEstado As Integer, ByVal pnSaldoCap As Currency, ByVal pnNewPrestamo As Currency, _
        ByVal pnTipoTasacion As Integer, ByVal pnUsoLineaActual As Integer, ByVal pnNroCalendDesem As Integer, _
        ByVal pnMontoColAnt As Currency, pdFecha As Date, Optional ByVal sPersLavDinero As String = "", _
               Optional ByVal sPersCod As String = "") As Long

Dim lsSQL As String
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim fnVarTasacionAdic As Currency
Dim fnVarTasacion As Currency
Dim fnVarPorcConserv As Currency
Dim fnVarPrecioMaterial As Currency
Dim fnVarComTasacion As Currency
Dim fnCalendDesem As Integer
Dim oCont As NContFunciones
Dim prPigJoya As ADODB.Recordset
Dim loFunctVal As ADODB.Recordset
Dim loObtieneJoyas As DPigContrato
Dim loRegPig As DPigActualizaBD
Dim loFunct As NPigCalculos
Dim loValores As DPigFunciones
Dim mbTrans As Boolean
Dim lsMovNroAmort As String
Dim lnMovNroAmort As Long
Dim lsMovNroDesemb As String
Dim lnMovNroDesemb As Long
Dim lnMontoTransac As Currency


If pnDiasAtraso > pnDiasCambCart Then
   lsOpeCod = gPigOpeAmortMorUL
Else
   lsOpeCod = gPigOpeAmortNorUL
End If

'On Error GoTo ErrorModPig
Set loRegPig = New DPigActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    lnMontoTransac = pnNewPrestamo - pnMontoTransac - fnVarComTasacion
    '***************************************** POR REUSO DE LINEA *************************************
    Call loRegPig.dInsertMov(psMovNro, gPigOpeReusoLinea, "Reuso de Linea", gMovEstContabMovContable, gMovFlagVigente, False)
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    ' **************************************************
    ' INSERTA DE LAVADO DE DINERO
    Dim clsMant As DCapMovimientos
    Set clsMant = New DCapMovimientos
    If sPersLavDinero <> "" Then
          clsMant.AgregaMovLavDinero lnMovNro, sPersCod
      End If
    Set clsMant = Nothing
  ' **************************************************
    
    '** Inserta Mov por la Amortizacion
    lsMovNroAmort = loRegPig.GeneraMovNro(pdFecha, gsCodAge, gsCodUser, psMovNro)
    Call loRegPig.dInsertMov(lsMovNroAmort, lsOpeCod, "Amortización por Reuso de Linea", gMovEstContabMovContable, gMovFlagVigente, False)
    'Obtiene nMovNro
    lnMovNroAmort = loRegPig.dGetnMovNro(lsMovNroAmort)
    
    '** Inserta Mov por el Desembolso
    lsMovNroDesemb = loRegPig.GeneraMovNro(pdFecha, gsCodAge, gsCodUser, lsMovNroAmort)
    Call loRegPig.dInsertMov(lsMovNroDesemb, gPigOpeDesembolsoUL, "Desembolso por Reuso de Linea", gMovEstContabMovContable, gMovFlagVigente, False)
    lnMovNroDesemb = loRegPig.dGetnMovNro(lsMovNroDesemb)
    
    Call loRegPig.dInsertMovCol(lnMovNro, gPigOpeReusoLinea, psCtaCod, pnNroCalend, lnMontoTransac, pnDiasAtraso, "", pnPlazo, pnEstado, pnSaldoCap, False)
    
    '***************************************** POR AMORTIZACION ***************************************
    '** Actualiza ColocCalendario
    Call loRegPig.dUpdateColocCalendario(psCtaCod, pnNroCalend, gColocCalendAplCuota, gColocCalendEstadoPagado, , , pdFecha, False)
            
    Call loRegPig.dInsertMovRef(lnMovNroAmort, lnMovNro)

    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNroAmort, lsOpeCod, psCtaCod, pnNroCalend, pnMontoTransac, pnDiasAtraso, "", pnPlazo, pnEstado, pnSaldoCap, False)

    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Capital
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gColPigConceptoCodCapital, , pnCapPag, , False)
    Call loRegPig.dInsertMovColDet(lnMovNroAmort, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodCapital, 1, pnCapPag, False)
    
    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Interes Compensatorio
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gColPigConceptoCodInteresComp, pnIntCompensatorio, pnIntCompensatorio, , False)
    If pnIntCompensatorio > 0 Then
       Call loRegPig.dInsertMovColDet(lnMovNroAmort, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodInteresComp, 1, pnIntCompensatorio, False)
    End If
    
    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Interes Moratorio
    If pnIntMoratorio > 0 Then
       Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gcolPigConceptoCodInteresMora, , pnIntMoratorio, , False)
       Call loRegPig.dInsertMovColDet(lnMovNroAmort, lsOpeCod, psCtaCod, pnNroCalend, gcolPigConceptoCodInteresMora, 1, pnIntMoratorio, False)
    End If
    
    '** Inserta MovColDet - ColocCalendDetPig (Se actualiza el Actual) - Comision Servicio
    If pnComServicio > 0 Then
       Call loRegPig.dInsertColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, 1, gColPigConceptoCodComiServ, _
                                                                pnComServicio, pnComServicio, "", False)
       Call loRegPig.dInsertMovColDet(lnMovNroAmort, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodComiServ, 1, pnComServicio, False)
    End If
    
    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Comision Vencida
    If pnComVencida > 0 Then
       Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gColPigConceptoCodComiVencida, , pnComVencida, , False)
       Call loRegPig.dInsertMovColDet(lnMovNroAmort, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodComiVencida, 1, pnComVencida, False)
    End If
    '** Inserta MovColDet / Actualiza ColocCalendDetPig (Se actualiza el Actual) - Derecho Remate
    If pnDerRemate > 0 Then
       Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalend, gColocCalendAplCuota, gColPigConceptoCodPreparaRemate, , pnDerRemate, , False)
       Call loRegPig.dInsertMovColDet(lnMovNroAmort, lsOpeCod, psCtaCod, pnNroCalend, gColPigConceptoCodPreparaRemate, 1, pnDerRemate, False)
    End If
    
    '***************************************** POR EL DESEMBOLSO ***************************************
    'Inicializa
    fnVarTasacionAdic = 0:     fnVarTasacion = 0
    fnVarPorcConserv = 0:        fnVarPrecioMaterial = 0
    fnVarComTasacion = 0
    
    Set loFunctVal = New ADODB.Recordset
    Set loValores = New DPigFunciones
         Set loFunctVal = loValores.GetConceptoValor(gColPigConceptoCodTasacion)     'Comision Tasación
    
    Set loFunct = New NPigCalculos
    fnVarComTasacion = loFunct.nCalculaConcepto(loFunctVal!nTpoValor, loFunctVal!nValor, IIf(IsNull(loFunctVal!nMontoMin), 0, loFunctVal!nMontoMin), _
                                                     IIf(IsNull(loFunctVal!nMontoMax), 0, loFunctVal!nMontoMax), pnNewSaldoCap)
    fnCalendDesem = pnNroCalendDesem + 1
    '** Insert ColocCalendario         (Se genera un nuevo Calendario de Desembolso y se cancela al mismo tiempo)
    Call loRegPig.dInsertColocCalendario(psCtaCod, pnNroCalend + 1, gColocCalendAplDesembolso, 1, psFecVenci, _
                                                         gColocCalendEstadoPagado, "Desembolso por.UsoLinea de Cred. Pignoraticio", "", 0, False)
    
    '** Insert ColocCalendDetPig     (Se genera un nuevo Calendario de Desembolso y se cancela al mismo tiempo)
    Call loRegPig.dInsertColocCalendDetPig(psCtaCod, pnNroCalend + 1, gColocCalendAplDesembolso, 1, gColPigConceptoCodCapital, _
                                                             pnNewPrestamo, pnNewPrestamo, "", False)
    Call loRegPig.dInsertColocCalendDetPig(psCtaCod, pnNroCalend + 1, gColocCalendAplDesembolso, 1, gColPigConceptoCodTasacion, _
                                                             fnVarComTasacion, fnVarComTasacion, "", False)
            
    Call loRegPig.dInsertMovRef(lnMovNroDesemb, lnMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNroDesemb, gPigOpeDesembolsoUL, psCtaCod, pnNroCalend + 1, pnNewPrestamo - fnVarComTasacion, pnDiasAtraso, "", pnPlazo, 0, pnMontoColAnt, False)

    '** Inserta MovColDet - Capital
    Call loRegPig.dInsertMovColDet(lnMovNroDesemb, gPigOpeDesembolsoUL, psCtaCod, pnNroCalend + 1, gColPigConceptoCodCapital, 1, pnNewPrestamo, False)
    
    '** Inserta MovColDet - Comision Tasacion
    Call loRegPig.dInsertMovColDet(lnMovNroDesemb, gPigOpeDesembolsoUL, psCtaCod, pnNroCalend + 1, gColPigConceptoCodTasacion, 1, 0 - fnVarComTasacion, False)
    
    ' Calcula Nuevo Tasación
    Set prPigJoya = New ADODB.Recordset
    Set loObtieneJoyas = New DPigContrato
         Set prPigJoya = loObtieneJoyas.dObtieneDatosPignoraticioJoyas(psCtaCod) 'Obtiene Caracteristicas de las joyas
    
    Do While Not prPigJoya.EOF
         fnVarTasacionAdic = fnVarTasacionAdic + prPigJoya!nTasacionAdicional
         If prPigJoya!nConservacion = 1 Then
            fnVarPorcConserv = loValores.GetParamValor(8010)          ' Conservacion de Joya Buena
         ElseIf prPigJoya!nConservacion = 2 Then
            fnVarPorcConserv = loValores.GetParamValor(8012)          ' Conservacion de Joya Regular
         ElseIf prPigJoya!nConservacion = 3 Then
            fnVarPorcConserv = loValores.GetParamValor(8011)          ' Conservacion de Joya Mala
         End If
         
         fnVarPrecioMaterial = loValores.GetPrecioMaterial(1, prPigJoya!nMaterial, 1)
         fnVarTasacion = fnVarTasacion + Round(loFunct.CalcValorTasacion(fnVarPorcConserv, prPigJoya!nPesoNeto, fnVarPrecioMaterial), 2)
         
         If pnTipoTasacion = gPigTipoTasacNor Then
            Call loRegPig.dInsertColocPigJoyaTasacion(psCtaCod, gPigTipoTasacUsoLin, prPigJoya!nItemPieza, prPigJoya!nMaterial, _
                                                  prPigJoya!nTipoJoya, prPigJoya!nSubTipoJoya, prPigJoya!nConservacion, _
                                                   prPigJoya!nPesoBruto, prPigJoya!nPesoNeto, fnVarTasacion, prPigJoya!nTasacionAdicional, _
                                                  IIf(IsNull(prPigJoya!cDescripcion), "", prPigJoya!cDescripcion), _
                                                  IIf(IsNull(prPigJoya!cDescripcionAdic), "", prPigJoya!cDescripcionAdic), prPigJoya!nUbicaPieza, _
                                                  prPigJoya!cPersCod, psMovNro, False)
         Else
             Call loRegPig.dUpdateColocPigJoyaTasacion(psCtaCod, gPigTipoTasacUsoLin, prPigJoya!nItemPieza, psMovNro, , , , , , , fnVarTasacion, , , , , , False)
         End If
         prPigJoya.MoveNext
     Loop
    
    '** Insert ColocCalendario         (Se genera un nuevo Calendario de Pagos)
    Call loRegPig.dInsertColocCalendario(psCtaCod, pnNroCalend + 1, gColocCalendAplCuota, 1, psFecVenci, _
                                                         gColocCalendEstadoPendiente, "Nuevo Calendario de Pagos", "", 0, False)
    
    '** Insert ColocCalendDetPig     (Se genera un nuevo Calendario de Pagos)
    Call loRegPig.dInsertColocCalendDetPig(psCtaCod, pnNroCalend + 1, gColocCalendAplCuota, 1, gColPigConceptoCodCapital, _
                                                             pnNewPrestamo, 0, "", False)
    Call loRegPig.dInsertColocCalendDetPig(psCtaCod, pnNroCalend + 1, gColocCalendAplCuota, 1, gColPigConceptoCodInteresComp, _
                                                             pnNewIntComp, 0, "", False)
    Call loRegPig.dInsertColocCalendDetPig(psCtaCod, pnNroCalend + 1, gColocCalendAplCuota, 1, gcolPigConceptoCodInteresMora, _
                                                             0, 0, "", False)
    
   '************************************ POR LA UTILIZACION DE LINEA ***********************************
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtaCod, , pnNewPrestamo, gPigEstReusoLin, psFechaHora, -2, False)          ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtaCod, pnNewPlazo, psFecVenci, pnNewPrestamo, , psMovNro, , , False)

    '** Actualiza ColocPigno
    Call loRegPig.dUpdateColocPigno(psCtaCod, pnUsoLineaActual + 1, gPigEstReusoLin, , pnNroCalend + 1, gPigTipoTasacUsoLin, 0, pnNroAmort, False)
    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtaCod, psFechaHora, gPigEstReusoLin, 1, pnNewPrestamo, "Uso de Linea de Cred. Pignoraticio", _
                                                       gColocCalendCodFFCF, 0, 0, pnNewPlazo, "0", -1, -1, -1, -1, False)
    
    loRegPig.dCommitTrans
    mbTrans = False
    Set loRegPig = Nothing
    nUsoLineaCredPignoraticio = lnMovNro

Exit Function
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nUsoLineaCredPignoraticio", "Error en Proceso de grabación de Uso de Linea"

End Function

'***************************************************************************
'CMCPL EXTORNO DESEMBOLSO CREDITO   'EAFA
'***************************************************************************
Public Function nExtornoDesembolsoCredPig(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        ByVal pnNroCalen As Integer, Optional pbEjecBatch As Boolean = False) As Long

Dim lsSQL As String
Dim loRegPig As DPigActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim mbTrans As Boolean

lsOpeCod = gPigOpeExtDesembolso

'On Error GoTo ErrorModPig
Set loRegPig = New DPigActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtaCod, , , gPigEstRegis, psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtaCod, , , , , psMovNro, , , False)

    '** Actualiza ColocCalendario
    Call loRegPig.dUpdateColocCalendario(psCtaCod, pnNroCalen, gColocCalendAplDesembolso, gColocCalendEstadoPendiente, , , , False)
    
    '** Actualiza ColocCalendDetPig
        '**** Capital
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplDesembolso, gColPigConceptoCodCapital, , 0, , False)
        '**** Comision de Tasacion
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplDesembolso, gColPigConceptoCodTasacion, , 0, , False)
    
    '** Actualiza ColocPigno
    Call loRegPig.dUpdateColocPigno(psCtaCod, , gPigEstRegis, , , , , , False)
    'Call loRegPig.dUpdateColocPigno(psCtaCod, , , , , , , , , gPigEstRegis, False)
    
    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtaCod, psFechaHora, gPigEstRegis, 1, pnMonto, "Extorno Desembolso Pignoraticio", _
         gColocCalendCodFFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Extorno Desembolso Pignoraticio", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 1, pnMonto, 0, "", 0, gPigEstRegis, 0, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    nExtornoDesembolsoCredPig = lnMovNro
    
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Function
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nExtornoDesembolsoCredPig", "Error en Proceso de Extorno de Desembolso"

End Function

Public Function nExtornoAmortizacionCredPig(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        ByVal pnNroCalen As Integer, Optional pbEjecBatch As Boolean = False) As Long

Dim lsSQL As String
Dim loRegPig As DPigActualizaBD
Dim loCalc As NPigCalculos
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lnInteres As Currency

Dim lnAntSaldoCap As Currency
Dim lnAntEstado As Integer
Dim lnAntPlazo As Integer
Dim lsAntFecVenc As String
Dim mbTrans As Boolean

Dim loDataExt As DPigFunciones
Dim lrDataExt As ADODB.Recordset
Dim Rs As Recordset
Dim ldPrdEstadoAnt As Date

lsOpeCod = gPigOpeExtAmortizEFE

' *** Obtiene Datos para el Extorno
lsSQL = " SELECT CE.nPlazo, CE.cCtaCod, " _
         & "              dVencAnt = (SELECT dVenc FROM ColocCalendario WHERE cCtaCod = '" & psCtaCod & "' and nNroCalen = " & pnNroCalen & "  " _
         & "                                                                                            and nColocCalendApl = " & gColocCalendAplCuota & ") , " _
         & "              nSaldoCapAnt = (SELECT nSaldoCap FROM MovCol where nMovNro = " & pnMovNroAnt & "  ) , " _
         & "              nPlazoAnt = (SELECT nPlazo FROM MovCol where nMovNro = " & pnMovNroAnt & "  ) , " _
         & "              nDiasAtraso = (SELECT nDiasMora FROM MovCol where nMovNro = " & pnMovNroAnt & "  ) , " _
         & "              nPrdEstadoAnt = (SELECT nCredEstado FROM MovCol where nMovNro = " & pnMovNroAnt & "  ), " _
         & "              nAmort = (SELECT nNroAmort FROM ColocPigno where cCtaCod = '" & psCtaCod & "'  ),  " _
         & "              nTasa = (SELECT nTasaInteres FROM Producto where cCtaCod = '" & psCtaCod & "'  )  " _
         & "   FROM ColocacEstado CE WHERE CE.cCtaCod = '" & psCtaCod & "' AND CE.nPrdEstado = " & gPigEstAmortiz & "  " _
         & "    AND CE.dPrdEstado = (SELECT MAX(dPrdEstado) FROM ColocacEstado " _
         & "                                      WHERE cCtaCod = CE.cCtaCod AND nPrdEstado = CE.nPrdEstado ) "

Set loDataExt = New DPigFunciones
    Set lrDataExt = loDataExt.dObtieneRecordSet(lsSQL)

If lrDataExt Is Nothing Then
    MsgBox "ERROR: al Buscar datos para Extorno ", vbInformation, "Aviso"
    Exit Function
End If
If lrDataExt.BOF And lrDataExt.EOF Then
    MsgBox "ERROR: No encontro datos de la Operacion a extornar ", vbInformation, "Aviso"
    Exit Function
End If

lnAntSaldoCap = lrDataExt!nSaldoCapAnt
lnAntEstado = lrDataExt!nPrdEstadoAnt
lnAntPlazo = lrDataExt!nPlazoAnt
lsAntFecVenc = Format(lrDataExt!dVencAnt, "mm/dd/yyyy")

'   lsSQL = "SELECT Max(dPrdEstado) dPrdEstado FROM ColocacEstado WHERE cCtaCod = '" & psCtaCod & "' AND nPrdEstado = " & lnAntEstado
    lsSQL = "Select Top 2 * from ColocacEstado Where cCtaCod = '" & psCtaCod & "' Order By dPrdEstado Desc"
    
    Set Rs = loDataExt.dObtieneRecordSet(lsSQL)
    
    If Not (Rs.EOF And Rs.BOF) Then
        Rs.MoveLast
        ldPrdEstadoAnt = Rs!dPrdEstado
    End If
    
    Set loDataExt = Nothing

Set loRegPig = New DPigActualizaBD
    
    loRegPig.dBeginTrans
    mbTrans = True
    
    Set loCalc = New NPigCalculos
        lnInteres = loCalc.nCalculaIntCompensatorio(lnAntSaldoCap, lrDataExt!nTasa, lnAntPlazo)
    
    Call loRegPig.dDeleteColocEstadoExtorno(psCtaCod, gPigEstAmortiz)
        
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtaCod, , lnAntSaldoCap, lnAntEstado, Format(ldPrdEstadoAnt, "mm/dd/yyyy hh:mm:ss"), -2, False)       ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtaCod, lnAntPlazo, lsAntFecVenc, , , psMovNro, , , False)

    '** Actualiza ColocCalendario
    Call loRegPig.dUpdateColocCalendario(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColocCalendEstadoPendiente, , , , False)
    
    '** Actualiza ColocPigno
    'Call loRegPig.dUpdateColocPigno(psCtaCod, , lrDataExt!nAmort - 1, , pnNroCalen, , , lrDataExt!nDiasAtraso, , lnAntEstado, False)
    Call loRegPig.dUpdateColocPigno(psCtaCod, , lnAntEstado, , pnNroCalen, , 0, lrDataExt!nAmort - 1, False)
    
    '** Actualiza ColocCalendDetPig
        '**** Capital
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColPigConceptoCodCapital, , 0, , False)
        '**** Interes Compensatorio
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColPigConceptoCodInteresComp, lnInteres, 0, , False)
        '**** Interes Moratorio
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gcolPigConceptoCodInteresMora, , 0, , False)
        '**** Comision de Vencimiento
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColPigConceptoCodComiVencida, , 0, , False)
        '**** Derecho de Remate
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColPigConceptoCodPreparaRemate, , 0, , False)
    
   '** Elimina el nuevo Calendario que se genero
       '**** Detalle
     Call loRegPig.dDeleteColocCalendDetPig(psCtaCod, pnNroCalen + 1, gColocCalendAplCuota, 1, , False)
       '**** Cabecera
     Call loRegPig.dDeleteColocCalendario(psCtaCod, pnNroCalen + 1, gColocCalendAplCuota, 1, False)
    
    '** Elimina del Anterior Calendario la Comision de Servicio y Penalidad
       '**** Comision de Servicio
     Call loRegPig.dDeleteColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, 1, gColPigConceptoCodComiServ, False)
       '**** Penalidad
     Call loRegPig.dDeleteColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, 1, gColPigConceptoCodPenalidad, False)
    
'    '** Insert ColocEstado
'    Call loRegPig.dInsertColocacEstado(psCtaCod, psFechaHora, lnAntEstado, 1, pnMonto, "Extorno Amortización Pignoraticio", _
'         gColocCalendCodFFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Extorno Amortización Pignoraticio", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, pnNroCalen, pnMonto, 0, "", 0, lnAntEstado, 0, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert MovRef
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    nExtornoAmortizacionCredPig = lnMovNro
    
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing
Set loCalc = Nothing

Exit Function
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nExtornoAmortizacionCredPig", "Error en Proceso de Extorno de Amortización"

End Function

Public Function nExtornoCancelacionCredPig(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        ByVal pnNroCalen As Integer, Optional pbEjecBatch As Boolean = False) As Long

Dim lsSQL As String
Dim loRegPig As DPigActualizaBD
Dim loCalc As NPigCalculos
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lnInteres As Currency
Dim ldPrdEstadoAnt As Date
Dim loDataExt As DPigFunciones
Dim lrDataExt As ADODB.Recordset
Dim Rs As Recordset

lsOpeCod = gPigOpeExtCancelacEFE

' *** Obtiene Datos para el Extorno
lsSQL = "SELECT CE.nPlazo, CE.cCtaCod, CE.dPrdEstado, " _
         & "              nSaldoCapAnt = (SELECT nSaldoCap FROM MovCol where nMovNro = " & pnMovNroAnt & "  ) , " _
         & "              nDiasAtraso = (SELECT nDiasMora FROM MovCol where nMovNro = " & pnMovNroAnt & "  ) , " _
         & "              nPlazoAnt = (SELECT nPlazo FROM MovCol where nMovNro = " & pnMovNroAnt & "  ) , " _
         & "              nPrdEstadoAnt = (SELECT nCredEstado FROM MovCol where nMovNro = " & pnMovNroAnt & "  ), " _
         & "              nAmort = (SELECT nNroAmort FROM ColocPigno where cCtaCod = '" & psCtaCod & "'  ) , " _
         & "              nTasa = (SELECT nTasaInteres FROM Producto where cCtaCod = '" & psCtaCod & "'  )  " _
         & "   FROM ColocacEstado CE WHERE CE.cCtaCod = '" & psCtaCod & "' AND CE.nPrdEstado = " & gPigEstCancelPendRes & "  " _
         & "     AND CE.dPrdEstado = (SELECT MAX(dPrdEstado) FROM ColocacEstado " _
         & "                                      WHERE cCtaCod = CE.cCtaCod AND nPrdEstado = CE.nPrdEstado ) "

Set loDataExt = New DPigFunciones
    Set lrDataExt = loDataExt.dObtieneRecordSet(lsSQL)

If lrDataExt Is Nothing Then
    MsgBox "ERROR: al Buscar datos para Extorno ", vbInformation, "Aviso"
    Exit Function
End If
If lrDataExt.BOF And lrDataExt.EOF Then
    MsgBox "ERROR: No encontro datos de la Operacion a extornar ", vbInformation, "Aviso"
    Exit Function
End If

' Asigna los valores
Dim lnAntSaldoCap As Currency
Dim lnAntEstado As Integer
Dim lnAntPlazo As Integer
Dim lsAntFecVenc As String
Dim mbTrans As Boolean

lnAntSaldoCap = lrDataExt!nSaldoCapAnt
lnAntEstado = lrDataExt!nPrdEstadoAnt
lnAntPlazo = lrDataExt!nPlazoAnt

    lsSQL = "SELECT Max(dPrdEstado) dPrdEstado FROM ColocacEstado WHERE cCtaCod = '" & psCtaCod & "' AND nPrdEstado = " & lnAntEstado
    Set Rs = loDataExt.dObtieneRecordSet(lsSQL)
    
    If Not (Rs.EOF And Rs.BOF) Then
        ldPrdEstadoAnt = Rs!dPrdEstado
    End If
    
    Set loDataExt = Nothing
'*** Realiza el Extorno
Set loRegPig = New DPigActualizaBD
    
    loRegPig.dBeginTrans
    mbTrans = True
    
    Set loCalc = New NPigCalculos
        lnInteres = loCalc.nCalculaIntCompensatorio(lnAntSaldoCap, lrDataExt!nTasa, lnAntPlazo)
    
    Call loRegPig.dDeleteColocEstadoExtorno(psCtaCod, gPigEstCancelPendRes)
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtaCod, , lnAntSaldoCap, lnAntEstado, Format(ldPrdEstadoAnt, "mm/dd/yyyy hh:mm:ss"), -2, False)        ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtaCod, , , , , psMovNro, , , False)

    '** Actualiza ColocPigno
    'Call loRegPig.dUpdateColocPigno(psCtaCod, , , , , , , lrDataExt!nDiasAtraso, , lnAntEstado, False)
    Call loRegPig.dUpdateColocPigno(psCtaCod, , lnAntEstado, , , , 0, , False)
    
    '** Actualiza ColocCalendario
    Call loRegPig.dUpdateColocCalendario(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColocCalendEstadoPendiente, , , , False)
    
    '** Actualiza ColocCalendDetPig
        '**** Capital
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColPigConceptoCodCapital, , 0, , False)
        '**** Interes Compensatorio
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColPigConceptoCodInteresComp, lnInteres, 0, , False)
        '**** Interes Moratorio
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gcolPigConceptoCodInteresMora, , 0, , False)
        '**** Comision de Vencimiento
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColPigConceptoCodComiVencida, , 0, , False)
        '**** Derecho de Remate
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColPigConceptoCodPreparaRemate, , 0, , False)
    
    '** Elimina del Ultimo Calendario la Comision de Servicio y Penalidad
       '**** Comision de Servicio
    Call loRegPig.dDeleteColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, 1, gColPigConceptoCodComiServ, False)
       '**** Penalidad
    Call loRegPig.dDeleteColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, 1, gColPigConceptoCodPenalidad, False)
    
'    '** Insert ColocEstado
'    Call loRegPig.dInsertColocacEstado(psCtaCod, psFechaHora, lnAntEstado, 1, pnMonto, "Extorno Cancelacion Cred Pig", _
'         gColocCalendCodFFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Cancelacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    '** Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 1, pnMonto, 0, "", 0, lnAntEstado, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    nExtornoCancelacionCredPig = lnMovNro
    
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing
Set loCalc = Nothing

Exit Function
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nExtornoCancelacionCredPig", "Error en Proceso de Extorno de Cancelación"

End Function
'***************************************************************************
'CMCPL EXTORNO DE USO DE LINEA
'EAFA
'***************************************************************************
Public Function nExtornoUsoLineaCredPig(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnMovNroAnt As Long, ByVal pnMovAmort As Long, ByVal pnMovDesem As Long, _
        ByVal pnMonto As Currency, ByVal pnNroCalen As Integer) As Long

Dim lsSQL As String
Dim loRegPig As DPigActualizaBD
Dim loCalc As NPigCalculos
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lnInteres As Currency
Dim lnTipoTasacion As Integer

Dim loDataExt As DPigFunciones
Dim lrDataExt As ADODB.Recordset
Dim lnAntColoc As Currency
Dim lnAntSaldoCap As Currency
Dim lnAntEstado As Integer
Dim lnAntPlazo As Integer
Dim lsAntFecVenc As String
Dim mbTrans  As Boolean
Dim ldPrdEstadoAnt As Date
Dim Rs As Recordset

lsOpeCod = gPigOpeExtReusoLinea

On Error GoTo ErrorModPig

lsSQL = " SELECT CE.nPlazo, CE.cCtaCod, " _
         & "              dVencAnt = (SELECT dVenc FROM ColocCalendario WHERE cCtaCod = '" & psCtaCod & "' and nNroCalen = " & pnNroCalen & "  " _
         & "                                                                                            and nColocCalendApl = " & gColocCalendAplCuota & ") , " _
         & "              nSaldoCapAnt = (SELECT nSaldoCap FROM MovCol where nMovNro = " & pnMovAmort & " and cOpeCod in ('" & gPigOpeAmortNorUL & "','" & gPigOpeAmortMorUL & "')), " _
         & "              nPlazoAnt = (SELECT nPlazo FROM MovCol where nMovNro = " & pnMovAmort & " and cOpeCod in ('" & gPigOpeAmortNorUL & "','" & gPigOpeAmortMorUL & "')), " _
         & "              nDiasAtraso = (SELECT nDiasMora FROM MovCol where nMovNro = " & pnMovAmort & " and cOpeCod in ('" & gPigOpeAmortNorUL & "','" & gPigOpeAmortMorUL & "')), " _
         & "              nPrdEstadoAnt = (SELECT nCredEstado FROM MovCol where nMovNro = " & pnMovAmort & " and cOpeCod in ('" & gPigOpeAmortNorUL & "','" & gPigOpeAmortMorUL & "')), " _
         & "              nColAnt = (SELECT nSaldoCap FROM MovCol where nMovNro = " & pnMovDesem & " and cOpeCod = '" & gPigOpeDesembolsoUL & "'), " _
         & "              nAmort = (SELECT nNroAmort FROM ColocPigno where cCtaCod = '" & psCtaCod & "'), " _
         & "              nUsoLinea = (SELECT nUsoLineaNro FROM ColocPigno where cCtaCod = '" & psCtaCod & "'), " _
         & "              nNroDesem = (SELECT nNumCalendDesem FROM ColocPigno where cCtaCod = '" & psCtaCod & "'), " _
         & "              nTasa = (SELECT nTasaInteres FROM Producto where cCtaCod = '" & psCtaCod & "') " _
         & "   FROM ColocacEstado CE WHERE CE.cCtaCod = '" & psCtaCod & "' AND CE.nPrdEstado = " & gPigEstReusoLin & "  " _
         & "    AND CE.dPrdEstado = (SELECT MAX(dPrdEstado) FROM ColocacEstado " _
         & "                                      WHERE cCtaCod = CE.cCtaCod AND nPrdEstado = CE.nPrdEstado ) "

    Set loDataExt = New DPigFunciones
        Set lrDataExt = loDataExt.dObtieneRecordSet(lsSQL)
    
    If lrDataExt.BOF And lrDataExt.EOF Then
        MsgBox "No encontro datos de la Operacion a extornar ", vbInformation, "Aviso"
        Exit Function
    End If

    lnAntSaldoCap = lrDataExt!nSaldoCapAnt
    lnAntEstado = lrDataExt!nPrdEstadoAnt
    lnAntPlazo = lrDataExt!nPlazoAnt
    lsAntFecVenc = Format$(lrDataExt!dVencAnt, "mm/dd/yyyy")
    lnAntColoc = lrDataExt!nColAnt

'    lsSQL = "SELECT Max(dPrdEstado) dPrdEstado FROM ColocacEstado WHERE cCtaCod = '" & psCtaCod & "' AND nPrdEstado = " & lnAntEstado
    lsSQL = "Select Top 2 * from ColocacEstado Where cCtaCod = '" & psCtaCod & "' Order By dPrdEstado Desc"
    
    Set Rs = loDataExt.dObtieneRecordSet(lsSQL)
    
    If Not (Rs.EOF And Rs.BOF) Then
        Rs.MoveLast
        ldPrdEstadoAnt = Rs!dPrdEstado
    End If
    
    Set loDataExt = Nothing
    
    Set loRegPig = New DPigActualizaBD
    
    loRegPig.dBeginTrans
    mbTrans = True
    
    Set loCalc = New NPigCalculos
    lnInteres = loCalc.nCalculaIntCompensatorio(lnAntSaldoCap, lrDataExt!nTasa, lnAntPlazo)
    
    Call loRegPig.dDeleteColocEstadoExtorno(psCtaCod, gPigEstReusoLin)
    
    '************************************* POR EL DESEMBOLSO **************************************
    '**** Elimina el nuevo Calendario de Pagos que se genero con la nueva Deuda
    Call loRegPig.dDeleteColocCalendDetPig(psCtaCod, pnNroCalen + 1, gColocCalendAplCuota, 1, , False)
    Call loRegPig.dDeleteColocCalendario(psCtaCod, pnNroCalen + 1, gColocCalendAplCuota, 1, False)
    
    '** Elimina el nuevo Calendario de Desembolso que se genero con el nuevo Desembolso
    Call loRegPig.dDeleteColocCalendDetPig(psCtaCod, pnNroCalen + 1, gColocCalendAplDesembolso, 1, , False)
    Call loRegPig.dDeleteColocCalendario(psCtaCod, pnNroCalen + 1, gColocCalendAplDesembolso, 1, False)
    
    '************************************* POR LA AMORTIZACION **************************************
    '** Actualiza ColocCalendario
    Call loRegPig.dUpdateColocCalendario(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColocCalendEstadoPendiente, , , , False)
    
    '** Actualiza ColocCalendDetPig
        '**** Capital
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColPigConceptoCodCapital, , 0, , False)
        '**** Interes Compensatorio
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColPigConceptoCodInteresComp, lnInteres, 0, , False)
        '**** Interes Moratorio
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gcolPigConceptoCodInteresMora, , 0, , False)
        '**** Comision de Vencimiento
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColPigConceptoCodComiVencida, , 0, , False)
        '**** Derecho de Remate
    Call loRegPig.dUpdateColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, gColPigConceptoCodPreparaRemate, , 0, , False)
    
    '** Elimina del Anterior Calendario la Comision de Servicio y Penalidad
    '**** Comision de Servicio
    Call loRegPig.dDeleteColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, 1, gColPigConceptoCodComiServ, False)
        '**** Penalidad
    Call loRegPig.dDeleteColocCalendDetPig(psCtaCod, pnNroCalen, gColocCalendAplCuota, 1, gColPigConceptoCodPenalidad, False)
    
    '************************************* POR LA UTILIZACION DE LINEA **************************************
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtaCod, , lnAntSaldoCap, lnAntEstado, Format(ldPrdEstadoAnt, "mm/dd/yyyy hh:mm:ss"), -2, False)      ' (-2) Aumenta la Sgte transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtaCod, lnAntPlazo, lsAntFecVenc, lnAntColoc, , psMovNro, , , False)

    If lrDataExt!nUsoLinea = 1 Then
    '** Elimina el registro en ColocPigJoyaTasacion
       Call loRegPig.dDeleteColocPigJoyaTasacion(psCtaCod, gPigTipoTasacUsoLin, , False)
        lnTipoTasacion = gPigTipoTasacNor
    Else
        lnTipoTasacion = gPigTipoTasacUsoLin
    End If
    
    '** Actualiza ColocPigno
    Call loRegPig.dUpdateColocPigno(psCtaCod, lrDataExt!nUsoLinea - 1, lnAntEstado, , pnNroCalen, lnTipoTasacion, 0, lrDataExt!nAmort - 1, False)
    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtaCod, psFechaHora, lnAntEstado, 1, pnMonto, "Extorno Utilizacion de Linea de Pignoraticio", _
         gColocCalendCodFFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Extorno Utilizacion de Linea de Pignoraticio", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, pnNroCalen, pnMonto, 0, "", 0, lnAntEstado, 0, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)
    Call loRegPig.dUpdateMov(pnMovAmort, , , , gMovFlagExtornado, False)
    Call loRegPig.dUpdateMov(pnMovDesem, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovAmort, False)
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovDesem, False)
    
    nExtornoUsoLineaCredPig = lnMovNro
    
    loRegPig.dCommitTrans
    mbTrans = False
    
    Set loRegPig = Nothing
    Set loCalc = Nothing

Exit Function

ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nExtornoUsoLineaCredPig", "Error en Proceso de Extorno de Uso de Linea"

End Function

'******************************************************
'**  CMCPL - Grabacion de Movimiento de Facturacion
'**  CRSF -  21/10/02
'******************************************************
  Public Sub nMovimientoPagoMov(ByVal psMovNro As String, Optional pbEjecBatch As Boolean = False)
'************************************
'** dInsertMov
'************************************
Dim lsSQL As String
Dim loRegPigno As DPigActualizaBD 'Nuevo CMCPL
Dim lnMovNro As Long
Dim mbTrans As Boolean

Set loRegPigno = New DPigActualizaBD
     loRegPigno.dBeginTrans
     mbTrans = True
            
    '** Actualiza MOV  - CMCPL
       Call loRegPigno.dInsertMov(psMovNro, gPigOpeVentaRemate, "Venta en el Remate", gMovEstContabMovContable, gMovFlagVigente, False)
     
    loRegPigno.dCommitTrans
    mbTrans = False
    Set loRegPigno = Nothing

Exit Sub

ErrorModPig:
    If mbTrans Then
        loRegPigno.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "
End Sub

'************************************************************************************
''**  CRSF -  21/10/02 - Grabacion de Movimiento de Facturacion -MOVCOL por cuenta
'************************************************************************************
Public Sub nMovimientoPagoMovCol(ByVal psCtaCod As String, ByVal pnSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal pnMontoEntregar As Currency, _
        ByVal pnComision As Currency, ByVal pnImpuesto As Currency, _
        Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
Dim loRegPigno As DPigActualizaBD 'Nuevo CMCPL
Dim lnMovNro As Long
Dim mbTrans As Boolean

Set loRegPigno = New DPigActualizaBD
     loRegPigno.dBeginTrans
     mbTrans = True
     
     '** Obtiene nMovNro
        lnMovNro = loRegPigno.dGetnMovNro(psMovNro)

     '** Actualiza MOVCOL - CMCPL
       Call loRegPigno.dInsertMovCol(lnMovNro, gPigOpeVentaRemate, psCtaCod, 1, pnSaldoCap, 0, "", 0, gPigEstRematFact, pnSaldoCap, False)

'    '** Inserta MovColDet -  Monto Capital
'    Call loRegPigno.dInsertMovColDet(lnMovNro, gPigOpeDesembolsoEFE, psCtaCod, 1, gColPigConceptoCodCapital, 1, pnSaldoCap, False)
'
'    '** Inserta MovColDet -  Interes Comision
'    Call loRegPigno.dInsertMovColDet(lnMovNro, gPigOpeDesembolsoEFE, psCtaCod, 1, gColPigConceptoCodTasacion, 1, pnComision, False)

   loRegPigno.dCommitTrans
   mbTrans = False
    Set loRegPigno = Nothing

Exit Sub

ErrorModPig:
    If mbTrans Then
        loRegPigno.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "
End Sub


'************************************************************************************
''**  CRSF -  29/10/02 - Grabacion de Movimiento de Facturacion -MOVCOLDET por concepto
'************************************************************************************
Public Sub nMovimientoPagoMovColDet(ByVal psCtaCod As String, ByVal pnMonto As Currency, _
         ByVal psMovNro As String, ByVal pnOpeCod As ColocPOperaciones, _
         ByVal pnConcepto As Integer, Optional pbEjecBatch As Boolean = False)
         
Dim lsSQL As String
Dim loRegPigno As DPigActualizaBD 'Nuevo CMCPL
Dim lnMovNro As Long
Dim mbTrans As Boolean


Set loRegPigno = New DPigActualizaBD
     loRegPigno.dBeginTrans
     mbTrans = True
     
     '** Obtiene nMovNro
        lnMovNro = loRegPigno.dGetnMovNro(psMovNro)

    '** Inserta MovColDet -  Monto Capital
  Call loRegPigno.dInsertMovColDet(lnMovNro, gPigOpeVentaRemate, psCtaCod, 1, pnConcepto, 1, pnMonto, False)
  'Call loRegPigno.dInsertMovColDet(lnMovNro, gPigOpeDesembolsoEFE, psCtaCod, 1, gColPigConceptoCodCapital, 1, pnSaldoCap, False)
   
   loRegPigno.dCommitTrans
   mbTrans = False
    Set loRegPigno = Nothing

Exit Sub

ErrorModPig:
    If mbTrans Then
        loRegPigno.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "
End Sub

'***************************************************************************
'CMCPL - REGISTRO DE VENTA DE JOYAS EN TIENDA
'EAFA - 09/10/2002
'***************************************************************************
Public Sub nRegistraVentaJoyas(ByVal psMovNro As String, ByVal pnTipDoc As Integer, ByVal psDocumento As String, _
                                                           ByVal psCliente As String, ByVal pnSubVenta As Currency, ByVal pnIGV As Currency, _
                                                           ByVal pnTotVenta As Currency, ByVal pnMoneda As Integer, ByVal pnTipoVenta As Integer, _
                                                           ByVal pmJoyas As Variant, ByVal pnFilas As Integer)
Dim loVtaJoya As DPigActualizaBD
Dim loRegPig As DPigActualizaBD
Dim lnMovNro As Long
Dim mbTrans As Boolean
Dim lsMotivo As String
Dim lsCont As String
Dim lnMonto As Currency
Dim lnIgv As Currency
Dim m As Integer
Dim f As Integer

'On Error GoTo ErrorModPig
Set loRegPig = New DPigActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True

    Set loVtaJoya = New DPigActualizaBD
  
    '**** Insert Mov
    Call loVtaJoya.dInsertMov(psMovNro, gPigOpeVentaJoya, "VENTA DE JOYAS", gMovEstContabMovContable, gMovFlagVigente, False)
    
    '**** Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '**** Insert MovDocPig
    Call loVtaJoya.dInsertMovDocPig(pnTipDoc, psDocumento, lnMovNro, pnTotVenta, psCliente, pnTipoVenta, pnMoneda, psMovNro, False)
    
    For f = 0 To pnFilas
          If pmJoyas(f, 6) = "" Then
              lnMonto = 0: lnIgv = 0
              For m = 0 To pnFilas
                    If pmJoyas(f, 1) = pmJoyas(m, 1) Then
                        lsCont = pmJoyas(m, 1)
                        lnMonto = lnMonto + Val(pmJoyas(m, 3))
                        lnIgv = lnIgv + Val(pmJoyas(m, 4))
                        pmJoyas(m, 6) = "X"
                    End If
              Next
              '**** Insert MovCol - Monto de Venta + IGV
              Call loVtaJoya.dInsertMovCol(lnMovNro, gPigOpeVentaJoya, lsCont, 0, lnMonto + lnIgv, 0, "", 0, 0, 0, False)
              '**** Insert MovColDet - Monto de Venta S/IGV
              Call loVtaJoya.dInsertMovColDet(lnMovNro, gPigOpeVentaJoya, lsCont, 0, gColPigConceptoCodPrecioVenta, 0, lnMonto, False)
              '**** Insert MovColDet - Monto de IGV
              Call loVtaJoya.dInsertMovColDet(lnMovNro, gPigOpeVentaJoya, lsCont, 0, gColPigConceptoCodIgv, 0, lnIgv, False)
          End If
          '**** Insert MovDocPigDet - Monto de Venta S/IGV
          Call loVtaJoya.dInsertMovDocPigDet(lnMovNro, pnTipDoc, psDocumento, pmJoyas(f, 0), pmJoyas(f, 1), pmJoyas(f, 2), gColPigConceptoCodPrecioVenta, pmJoyas(f, 3), False)
   
          '**** Insert MovDocPigDet - Monto de IGV
          Call loVtaJoya.dInsertMovDocPigDet(lnMovNro, pnTipDoc, psDocumento, pmJoyas(f, 0), pmJoyas(f, 1), pmJoyas(f, 2), gColPigConceptoCodIgv, pmJoyas(f, 4), False)
  
          '**** Update ColocPigProceso
          Call loVtaJoya.dUpdateColocPigProceso(pmJoyas(f, 5), gPigTipoVentas, pmJoyas(f, 1), pmJoyas(f, 2), psMovNro, , , gPigSituacionFacturado, , Val(pmJoyas(f, 3)) + Val(pmJoyas(f, 4)), , False)
    
          '**** Update ColocPigEstadoJoya
          Call loVtaJoya.dUpdateColocPigEstadoJoya(pmJoyas(f, 1), pmJoyas(f, 2), psMovNro, , , "1", False)
          
          '**** Update ColocPigJoyaTasacion
          Call loVtaJoya.dUpdateColocPigJoyaTasacion(pmJoyas(f, 1), gPigTipoTasacNor, pmJoyas(f, 2), , , , , , , , , , , , gPigUbicaSinCaja, , False)
          Call loVtaJoya.dUpdateColocPigJoyaTasacion(pmJoyas(f, 1), gPigTipoTasacUsoLin, pmJoyas(f, 2), , , , , , , , , , , , gPigUbicaSinCaja, , False)
          Call loVtaJoya.dUpdateColocPigJoyaTasacion(pmJoyas(f, 1), gPigTipoTasacRetasac, pmJoyas(f, 2), , , , , , , , , , , , gPigUbicaSinCaja, , False)
          
    Next
    
    loRegPig.dCommitTrans
    mbTrans = False

    Exit Sub

ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraVentaJoyas", "Error en Proceso de grabación de Venta de Joyas"
    
End Sub
'***************************************************************************
'CMCPL - REGISTRO DE JOYAS SELECCIONADAS PARA VENTA/FUNDICION
'EAFA - 30/10/2002
'***************************************************************************
Public Sub nRegistraSelVentaFundicion(ByVal pmJoyas As Variant, ByVal pnFilas As Long, ByVal pnTipoProceso As Integer, ByVal psMovNro As String)
Dim loVtaJoya As DPigActualizaBD
Dim loRegPig As DPigActualizaBD
Dim mbTrans As Boolean
Dim f As Integer

'On Error GoTo ErrorModPig
Set loRegPig = New DPigActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True

    Set loVtaJoya = New DPigActualizaBD
  
    For f = 0 To pnFilas
          If pnTipoProceso <> pmJoyas(f, 8) Then
              '**** Insert ColocPigProceso
              'Call loVtaJoya.dInsertColocPigProceso(pmJoyas(f, 3) , pnTipoProceso, pmJoyas(f, 0), pmJoyas(f, 1), pmJoyas(f, 6), pmJoyas(f, 4), gPigSituacionDisponible, pmJoyas(f, 5), 0, pmJoyas(f, 7), psMovNro)
          Else
              '**** Update ColocPigProceso
              Call loVtaJoya.dUpdateColocPigProceso(Val(pmJoyas(f, 3)), pnTipoProceso, pmJoyas(f, 0), Val(pmJoyas(f, 1)), psMovNro, , , , , , Val(pmJoyas(f, 7)), False)
          End If
          '**** Update ColocPigEstadoJoya
          Call loVtaJoya.dUpdateColocPigEstadoJoya(pmJoyas(f, 0), pmJoyas(f, 1), psMovNro, , pnTipoProceso, "0", False)
    Next
    
    loRegPig.dCommitTrans
    mbTrans = False

    Exit Sub

ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraSelVentaFundicion", "Error en la Grabación de Selección Venta/Fundición"
    
End Sub

'***************************************************************************
'CMCPL - ANULACION DE FACTURA/BOLETA DE VENTA DE JOYAS
'EAFA - 16/10/2002
'***************************************************************************
Public Sub nAnularVentaJoyas(ByVal pnMovNroAnt As Long, ByVal pmJoyas As Variant, ByVal psMovNro As String, ByVal pnFilas As Integer)
Dim loVtaJoya As DPigActualizaBD
Dim loRegPig As DPigActualizaBD
Dim lnMovNro As Long
Dim mbTrans As Boolean
Dim lsMotivo As String
Dim lsCont As String
Dim lnMonto As Currency
Dim lnIgv As Currency
Dim f As Integer
Dim m As Integer

'On Error GoTo ErrorModPig
Set loRegPig = New DPigActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True

    Set loVtaJoya = New DPigActualizaBD
  
    '**** Insert Mov
    Call loVtaJoya.dInsertMov(psMovNro, gPigOpeAnulaVentaJoya, "ANULACION DE VENTA DE JOYAS", gMovEstContabMovContable, gMovFlagVigente, False)
    
    '**** Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        
    For f = 0 To pnFilas
          If pmJoyas(f, 5) = "" Then
              lnMonto = 0: lnIgv = 0
              For m = 0 To pnFilas
                    If pmJoyas(f, 0) = pmJoyas(m, 0) Then
                        lsCont = pmJoyas(m, 0)
                        lnMonto = lnMonto + Val(pmJoyas(m, 3))
                        lnIgv = lnIgv + Val(pmJoyas(m, 4))
                        pmJoyas(m, 5) = "X"
                    End If
              Next
              Call loVtaJoya.dInsertMovCol(lnMovNro, gPigOpeAnulaVentaJoya, lsCont, 0, lnMonto + lnIgv, 0, "", 0, 0, 0, False)
          End If
          '**** Update ColocPigProceso
          Call loVtaJoya.dUpdateColocPigProceso(Val(pmJoyas(f, 2)), gPigTipoVentas, Trim(pmJoyas(f, 0)), Val(pmJoyas(f, 1)), psMovNro, , , gPigSituacionDisponible, , 0, , False)
    
          '**** Update ColocPigEstadoJoya
          Call loVtaJoya.dUpdateColocPigEstadoJoya(pmJoyas(f, 0), pmJoyas(f, 1), psMovNro, , , "0", False)
          
          '**** Update ColocPigJoyaTasacion
          Call loVtaJoya.dUpdateColocPigJoyaTasacion(pmJoyas(f, 1), gPigTipoTasacNor, pmJoyas(f, 2), , , , , , , , , , , , gPigUbicaTdaLarco, , False)
          Call loVtaJoya.dUpdateColocPigJoyaTasacion(pmJoyas(f, 1), gPigTipoTasacUsoLin, pmJoyas(f, 2), , , , , , , , , , , , gPigUbicaTdaLarco, , False)
          Call loVtaJoya.dUpdateColocPigJoyaTasacion(pmJoyas(f, 1), gPigTipoTasacRetasac, pmJoyas(f, 2), , , , , , , , , , , , gPigUbicaTdaLarco, , False)
    Next
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)


    loRegPig.dCommitTrans
    mbTrans = False

Exit Sub

ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nAnularVentaJoyas", "Error en Proceso de Anulación de Venta de Joyas"
    
End Sub

Public Sub nDespachoRemesa(ByVal psMovNro As String, ByVal psNumGuia As String, ByVal psSerie As String, _
                ByVal psSunat As String, psPersCod As String, ByVal pnTipoGuia As Integer, _
                ByVal pnUbicaTrans As Integer, ByVal Rs As Recordset, ByVal pnMotivo As Integer)

Dim oPigGrabar As DPigActualizaBD
Dim oDatos As DPigContrato
Dim lnMovNro As Long
Dim bTrans As Boolean

    On Error GoTo dError
    
    Set oPigGrabar = New DPigActualizaBD
    oPigGrabar.dBeginTrans
    bTrans = True
                 
    Call oPigGrabar.dUpdateColocPigGuia(psNumGuia, 2, , , , psPersCod, , , , , , , psSerie, psSunat, psMovNro, False)

    Call oPigGrabar.dInsertColocPigGuiaEtapa(psNumGuia, 2, psMovNro, False)
    
    If pnTipoGuia = 1 Then
        Call oPigGrabar.dUpdateColocPignoLote(psNumGuia, pnUbicaTrans)
        Call oPigGrabar.dUpdateColocPigJoyaTasacionPieza(psNumGuia, pnUbicaTrans, pnTipoGuia)
        
    ElseIf pnTipoGuia = 2 Then
        Call oPigGrabar.dUpdateColocPigJoyaTasacionPieza(psNumGuia, pnUbicaTrans, pnTipoGuia)
    End If

    '========== Inserta en el Mov
    Call oPigGrabar.dInsertMov(psMovNro, gPigOpeDespRemesa, "Despacho Remesa de Joyas", gMovEstContabMovContable, gMovFlagVigente, False)
            
    'Obtiene nMovNro
    lnMovNro = oPigGrabar.dGetnMovNro(psMovNro)
    
    'CUSTODIA DE VALORES
    Select Case pnMotivo
    
    Case 1      'CONTRATOS NUEVOS
        
        If Not (Rs.EOF And Rs.BOF) Then
            Do While Not Rs.EOF
                'Inserta MovCol
                Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeDespRemesa, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                
                'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespRemesa, Rs!cCtaCod, 1, gColPigConceptoCodGarantCustAge, 1, Rs!Tasacion)
                Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespRemesa, Rs!cCtaCod, 1, gColPigConceptoCodGarantTrans, 1, Rs!Tasacion)
            
                Rs.MoveNext
            Loop
        End If
        
    Case 2      'PENDIENTES DE RESCATE MAS DE 30 DIAS
        If Not (Rs.EOF And Rs.BOF) Then
            Do While Not Rs.EOF
                Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeDespPendRescMay, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespPendRescMay, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescAge, 1, Rs!Tasacion)
                Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespPendRescMay, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescTrans, 1, Rs!Tasacion)
                Rs.MoveNext
            Loop
        End If

    Case 3      'PENDIENTES DE RESCATE
        If Not (Rs.EOF And Rs.BOF) Then
            Do While Not Rs.EOF
                Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeDespPendResc, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespPendResc, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescBoveda, 1, Rs!Tasacion)
                Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespPendResc, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescTrans, 1, Rs!Tasacion)
                Rs.MoveNext
            Loop
        End If
    Case 4      'LOTES PARA REMATE (VENCIDOS Y PENDIENTES DE RESCATE)
        If Not (Rs.EOF And Rs.BOF) Then
            Do While Not Rs.EOF
                If Rs!nPrdEstado = gPigEstRemat Then              ' Los Contratos Vencidos
                    Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeDespRemateVenc, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                    'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantCustBoveda, 1, Rs!Tasacion)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantTrans, 1, Rs!Tasacion)
                ElseIf Rs!nPrdEstado = gPigEstRematPRes Then      ' Los Pendientes de Rescate
                    Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeDespRemateVenc, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                    'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescBoveda, 1, Rs!Tasacion)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescTrans, 1, Rs!Tasacion)
                End If
                Rs.MoveNext
            Loop
        End If
    Case 5      'DEVOLUCION DEL REMATE (VENCIDOS Y PENDIENTES DE RESCATE)
        If Not (Rs.EOF And Rs.BOF) Then
            Do While Not Rs.EOF
                If Rs!nPrdEstado = gPigEstRemat Then
                    Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeDespDevRemateVenc, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                    'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespDevRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantCustAge, 1, Rs!Tasacion)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespDevRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantTrans, 1, Rs!Tasacion)
                ElseIf Rs!nPrdEstado = gPigEstRematPRes Then      ' Los Pendientes de Rescate
                    Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeDespRemateVenc, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                    'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespDevRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescAge, 1, Rs!Tasacion)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespDevRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescTrans, 1, Rs!Tasacion)
                End If
                Rs.MoveNext
            Loop
        End If
    Case 6      'JOYAS PARA VENTA OP
        Do While Not Rs.EOF
            Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeDespJoyasVenta, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
            'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespJoyasVenta, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjBoveda, 1, Rs!Tasacion)
            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespJoyasVenta, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjTrans, 1, Rs!Tasacion)
            
            Rs.MoveNext
        Loop
    Case 7      'DEVOLUCION DE JOYAS A BOVEDA DE VENTAS EN OP
        Do While Not Rs.EOF
            Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeDespDevJoyasVenta, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
            'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespDevJoyasVenta, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjCustAge, 1, Rs!Tasacion)
            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespDevJoyasVenta, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjTrans, 1, Rs!Tasacion)
            
            Rs.MoveNext
        Loop
    
    Case 8      'JOYAS PARA VENTAS Y REMATES
        Do While Not Rs.EOF
            Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeDespJoyasVentaRemate, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
            'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespJoyasVentaRemate, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjBoveda, 1, Rs!Tasacion)
            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespJoyasVentaRemate, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjTrans, 1, Rs!Tasacion)
            
            Rs.MoveNext
        Loop
    Case 9      'DEVOLUCION DE JOYAS A BOVEDA DE VENTAS Y REMATES
        Do While Not Rs.EOF
            Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeDespDevJoyasVentaRemate, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
            'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespDevJoyasVentaRemate, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjCustAge, 1, Rs!Tasacion)
            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeDespDevJoyasVentaRemate, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjTrans, 1, Rs!Tasacion)
            
            Rs.MoveNext
        Loop
            
    End Select

    oPigGrabar.dCommitTrans
    bTrans = False
    
    Set oPigGrabar = Nothing
    Set Rs = Nothing
    Set oDatos = Nothing

    Exit Sub
    
dError:
    If bTrans Then
        oPigGrabar.dRollbackTrans
        bTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nDespachoRemesa", "Error en Despacho de Remesa"

End Sub

Public Sub nRecepcionRemesa(ByVal psNumGuia As String, psMovNro As String, pnTipoGuia As Integer, _
        ByVal psCodAge As String, ByVal Rs As Recordset, ByVal pnMotivo As Integer)

Dim oPigGrabar As DPigActualizaBD
Dim oDatos As DPigContrato
Dim lnMovNro As Long
Dim bTrans As Boolean

    On Error GoTo dError
    
    Set oPigGrabar = New DPigActualizaBD
    oPigGrabar.dBeginTrans
    bTrans = True

    Call oPigGrabar.dUpdateColocPignoLote(psNumGuia, psCodAge)
    Call oPigGrabar.dUpdateColocPigJoyaTasacionPieza(psNumGuia, psCodAge, pnTipoGuia)
    
    Call oPigGrabar.dInsertColocPigGuiaEtapa(psNumGuia, 3, psMovNro, False) 'Un registro por la Recepcion
        
    If MsgBox("Desea realizar la Confirmación del Contenido?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    
        Call oPigGrabar.dUpdateColocPigGuiaDet(psNumGuia, , , , , 1, False)
        Call oPigGrabar.dUpdateColocPigGuia(psNumGuia, 3, , , , , , , , , , , , , psMovNro, False)
    
    End If
    
    '========== Inserta en el Mov
    Call oPigGrabar.dInsertMov(psMovNro, gPigOpeRecepRemesa, "Recepcion Remesa de Joyas", gMovEstContabMovContable, gMovFlagVigente, False)
            
    'Obtiene nMovNro
    lnMovNro = oPigGrabar.dGetnMovNro(psMovNro)
            
    If Not (Rs.EOF And Rs.BOF) Then
        
        Select Case pnMotivo
            Case 1      'CONTRATOS NUEVOS
                If Not (Rs.EOF And Rs.BOF) Then
                    Do While Not Rs.EOF
                        'Inserta MovCol
                        Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeRecepRemesa, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                        
                        'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                        Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepRemesa, Rs!cCtaCod, 1, gColPigConceptoCodGarantTrans, 1, Rs!Tasacion)
                        Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepRemesa, Rs!cCtaCod, 1, gColPigConceptoCodGarantCustAge, 1, Rs!Tasacion)
                    
                        Rs.MoveNext
                    Loop
                End If
                
            Case 2      'PENDIENTES DE RESCATE MAS DE 30 DIAS
                If Not (Rs.EOF And Rs.BOF) Then
                    Do While Not Rs.EOF
                        Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeRecepPendRescMay, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                        'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                        Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepPendRescMay, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescTrans, 1, Rs!Tasacion)
                        Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepPendRescMay, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescBoveda, 1, Rs!Tasacion)
                        Rs.MoveNext
                    Loop
                End If
        
            Case 3      'PENDIENTES DE RESCATE
                If Not (Rs.EOF And Rs.BOF) Then
                    Do While Not Rs.EOF
                        Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeRecepPendResc, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                        'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                        Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepPendResc, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescTrans, 1, Rs!Tasacion)
                        Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepPendResc, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescAge, 1, Rs!Tasacion)
                        Rs.MoveNext
                    Loop
                End If
            Case 4      'LOTES PARA REMATE (VENCIDOS Y PENDIENTES DE RESCATE
                If Not (Rs.EOF And Rs.BOF) Then
                    Do While Not Rs.EOF
                        If Rs!nPrdEstado = gPigEstRemat Then     ' Los Contratos Vencidos
                            Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeRecepRemateVenc, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                            'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantTrans, 1, Rs!Tasacion)
                            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantCustAge, 1, Rs!Tasacion)
                        ElseIf Rs!nPrdEstado = gPigEstRematPRes Then      ' Los Pendientes de Rescate
                            Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeRecepRemateVenc, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                            'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescTrans, 1, Rs!Tasacion)
                            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescAge, 1, Rs!Tasacion)
                        End If
                        Rs.MoveNext
                    Loop
                End If
            Case 5      'DEVOLUVION DE REMATE - LOTES NO REMATADOS
                If Not (Rs.EOF And Rs.BOF) Then
                    Do While Not Rs.EOF
                        If Rs!nPrdEstado = gPigEstRemat Then
                            Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeRecepDevRemateVenc, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                            'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepDevRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantTrans, 1, Rs!Tasacion)
                            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepDevRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantCustBoveda, 1, Rs!Tasacion)
                        ElseIf Rs!nPrdEstado = gPigEstRematPRes Then      ' Los Pendientes de Rescate
                            Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeRecepRemateVenc, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                            'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepDevRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescTrans, 1, Rs!Tasacion)
                            Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepDevRemateVenc, Rs!cCtaCod, 1, gColPigConceptoCodGarantPendRescBoveda, 1, Rs!Tasacion)
                        End If
                        Rs.MoveNext
                    Loop
                End If
            Case 6      'JOYAS PARA VENTA
                Do While Not Rs.EOF
                    Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeRecepJoyasVenta, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                    'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepJoyasVenta, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjTrans, 1, Rs!Tasacion)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepJoyasVenta, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjCustAge, 1, Rs!Tasacion)
                    Rs.MoveNext
                Loop
            Case 7      'DEVOLUCION DE JOYAS A BOVEDA DE LA TIENDA
                Do While Not Rs.EOF
                
                    Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeRecepDevJoyasVenta, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                    'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepDevJoyasVenta, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjTrans, 1, Rs!Tasacion)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepDevJoyasVenta, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjBoveda, 1, Rs!Tasacion)
                    
                    Rs.MoveNext
                Loop
            Case 8      'JOYAS PARA VENTAS Y REMATES
                Do While Not Rs.EOF
                    Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeRecepJoyasVentaRemate, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                    'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepJoyasVentaRemate, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjTrans, 1, Rs!Tasacion)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepJoyasVentaRemate, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjCustAge, 1, Rs!Tasacion)
                    Rs.MoveNext
                Loop
            Case 9      'DEVOLUCION DE JOYAS A BOVEDA DE VENTAS Y REMATES
                Do While Not Rs.EOF
                    Call oPigGrabar.dInsertMovCol(lnMovNro, gPigOpeRecepDevJoyasVentaremate, Rs!cCtaCod, 1, Rs!Tasacion, 0, "", 0, 0, 0)
                    'Valor Tasacion  -- PARA LA CUSTODIA (En Transito)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepDevJoyasVentaremate, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjTrans, 1, Rs!Tasacion)
                    Call oPigGrabar.dInsertMovColDet(lnMovNro, gPigOpeRecepDevJoyasVentaremate, Rs!cCtaCod, 1, gColPigConceptoCodGarantBienAdjBoveda, 1, Rs!Tasacion)
                    
                    Rs.MoveNext
                Loop
            
            End Select
        
    End If

    oPigGrabar.dCommitTrans
    bTrans = False
    
    Set oPigGrabar = Nothing
    Set Rs = Nothing
    Set oDatos = Nothing

    Exit Sub
    
dError:
    If bTrans Then
        oPigGrabar.dRollbackTrans
        bTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRecepcionRemesa", "Error en Recepcion de Remesa"

End Sub

Public Sub nRescateLotes(ByVal psCtaCod As String, ByVal psMovNro As String, ByVal pnComision As Currency, _
                ByVal pnValTasacion As Currency, ByVal psFechaHora As String, Optional ByVal psCodAge As String)
                
Dim lnMovNro As Long
Dim oPigGraba As DPigActualizaBD
Dim mbTrans As Boolean
Dim oPigFunc As DPigFunciones
Dim lsAgeBovedaVal As String

Set oPigFunc = New DPigFunciones
lsAgeBovedaVal = CStr(oPigFunc.GetParamValor(gPigAgeBovedaVal))
Set oPigFunc = Nothing
    
    Set oPigGraba = New DPigActualizaBD
    oPigGraba.dBeginTrans
    mbTrans = True
    
    On Error GoTo ErrorPig
    
    '** Actualiza Producto - CMCPL
    Call oPigGraba.dUpdateProducto(psCtaCod, -1, -1, gPigEstRescate, psFechaHora, -2, False)
    
    '** Insert ColocEstado - CMCPL
     Call oPigGraba.dInsertColocacEstado(psCtaCod, psFechaHora, gPigEstRescate, 1, 0, "Rescate de Lote ", _
        gColocCalendCodFFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)

    '** Actualiza ColocPigno
    Call oPigGraba.dUpdateColocPigno(psCtaCod, , gPigEstRescate)
        
     '** Inserta Mov
    Call oPigGraba.dInsertMov(psMovNro, gPigOpeDevJoyas, "Rescate de Joyas", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = oPigGraba.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call oPigGraba.dInsertMovCol(lnMovNro, gPigOpeDevJoyas, psCtaCod, 0, pnComision, 0, "", 0, gPigEstCancelPendRes, 0, False)

    '** Inserta MovColDet - Descargo de CustodiaRescatePendientes
    '(Cuando es la Agencia de la Boveda de Valores lo envia directamente a la Boveda de Valores
    If Right(lsAgeBovedaVal, 2) = Right(psCodAge, 2) Then
        Call oPigGraba.dInsertMovColDet(lnMovNro, gPigOpeDevJoyas, psCtaCod, 0, gColPigConceptoCodGarantCustBoveda, 0, pnValTasacion, False)
    Else
        Call oPigGraba.dInsertMovColDet(lnMovNro, gPigOpeDevJoyas, psCtaCod, 0, gColPigConceptoCodGarantPendRescAge, 0, pnValTasacion, False)
    End If
    
    If pnComision > 0 Then
        'Comision
        Call oPigGraba.dInsertMovColDet(lnMovNro, gPigOpeDevJoyas, psCtaCod, 0, gColPigConceptoCodCustodia, 0, pnComision, False)
        'IGV
   '     Call oPigGraba.dInsertMovColDet(lnMovNro, gPigOpeDevJoyas, psCtaCod, 0, gColPigConceptoCodIgv, 0, pnIGV, False)
    End If
    
    oPigGraba.dCommitTrans
    mbTrans = False
    Set oPigGraba = Nothing
    Exit Sub
    
ErrorPig:
    If mbTrans Then
        oPigGraba.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRescateLotes", "Error en Funcion de Rescate de Lotes"

End Sub
'**********************************
'CAFF CMCPL RETASACION MANUAL
'**********************************
Public Sub nRetasacionManual(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal pnEstado As Integer, _
                            ByVal prJoyas As Recordset, ByVal psMovNro As String, ByVal psTasador As String)

Dim oRegPig As DPigActualizaBD
Dim lnTipo As Integer, lnSTipo As Integer
Dim lnMaterial As Integer, lnEstado As Integer
Dim lsDes As String, lsDesAdic As String
Dim lnPesoB As Double, lnPesoN As Double
Dim lnTasac As Currency, lnTasacAdic As Currency
Dim lnItem As Integer
Dim mbTrans As Boolean

On Error GoTo ErrorModPig
Set oRegPig = New DPigActualizaBD

    oRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Colocaciones
    Call oRegPig.dUpdateColocaciones(psCtaCod, , , , , psMovNro, , , False)
    Call oRegPig.dUpdateColocPigno(psCtaCod, , , , , gPigTipoTasacRetasac)
    
    '**** Actualiza ColocPigJoyaTasacion
    Do While Not prJoyas.EOF
        lnItem = prJoyas!Item
        lnTipo = Val(Right(prJoyas!Tipo, 3))
        lnSTipo = Val(Right(prJoyas!SubTipo, 3))
        lnMaterial = Val(Right(prJoyas!Material, 3))
        lnEstado = Val(Right(prJoyas!Estado, 3))
        lnPesoB = prJoyas!PBruto
        lnPesoN = prJoyas!PNeto
        lnTasac = prJoyas!Tasac
        lnTasacAdic = prJoyas!TasAdic
        lsDes = Trim(prJoyas!Observacion)
        lsDesAdic = Trim(prJoyas!ObsAdicion)
   
        Call oRegPig.dUpdateColocPigJoyaTasacion(psCtaCod, gPigTipoTasacRetasac, lnItem, , lnMaterial, lnTipo, lnSTipo, lnEstado, _
                                        lnPesoB, lnPesoN, lnTasac, lnTasacAdic, lsDes, lsDesAdic, , psTasador, False)
        
        prJoyas.MoveNext
    Loop

    oRegPig.dCommitTrans
    mbTrans = False
    
Set oRegPig = Nothing

Exit Sub

ErrorModPig:
    If mbTrans Then
        oRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRetasacionManual", "Error en Funcion Retasación Manual"

End Sub

'***************************************************************************
'CMCPL EXTORNO RESCATE JOYAS
'DFAO - 16/06/2003
'***************************************************************************
Public Sub nExtornoRescateJoyas(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        ByVal pnNroCalen As Integer, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
Dim loRegPig As DPigActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim mbTrans As Boolean

lsOpeCod = gPigOpeExtRescateJoya

'On Error GoTo ErrorModPig
Set loRegPig = New DPigActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtaCod, , , gPigEstCancelPendRes, psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    
    '** Actualiza ColocPigno
    Call loRegPig.dUpdateColocPigno(psCtaCod, , gPigEstCancelPendRes, , , , , , False)
    
    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtaCod, psFechaHora, gPigEstCancelPendRes, 1, 0, "Extorno Rescate de Joyas", _
         gColocCalendCodFFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
     
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Extorno Rescate de Joyas", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMonto, 0, "", 0, gPigEstCancelPendRes, 0, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nExtornoRescateJoyas", "Error en Proceso de Extorno de Rescate de Joyas"

End Sub

'***************************************************************************
'CMCPL - ANULACION DE POLIZA DE VENTA DE JOYAS EN REMATE
'DFAO - 16/06/2003
'***************************************************************************
Public Sub nAnularVentaJoyasRemate(ByVal pnMovNroAnt As Long, ByVal psFechaHora As String, ByVal pmJoyas As Variant, ByVal psMovNro As String, ByVal pnFilas As Integer)
Dim loVtaJoya As DPigActualizaBD
Dim loRegPig As DPigActualizaBD
Dim lnMovNro As Long
Dim mbTrans As Boolean
Dim lsMotivo As String
Dim lsCont As String
Dim lsCodCta As String
Dim lnMonto As Currency
Dim lnIgv As Currency
Dim f As Integer
Dim m As Integer

Dim lsContx As String

'On Error GoTo ErrorModPig
lsCodCta = ""
Set loRegPig = New DPigActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True

    Set loVtaJoya = New DPigActualizaBD
  
    '**** Insert Mov
    Call loVtaJoya.dInsertMov(psMovNro, gPigOpeAnulaVentaJoyaRemate, "ANULACION DE VENTA DE JOYAS", gMovEstContabMovContable, gMovFlagVigente, False)
    
    '**** Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        
    For f = 0 To pnFilas
          lsContx = pmJoyas(f, 0)
          If pmJoyas(f, 6) = "" Then
              lnMonto = 0: lnIgv = 0
              For m = 0 To pnFilas
                    If pmJoyas(f, 0) = pmJoyas(m, 0) Then
                        lsCont = pmJoyas(m, 0)
                        lnMonto = lnMonto + Val(pmJoyas(m, 3))
                        lnIgv = lnIgv + Val(pmJoyas(m, 4))
                        pmJoyas(m, 6) = "X"
                    End If
              Next
              Call loVtaJoya.dInsertMovCol(lnMovNro, gPigOpeAnulaVentaJoyaRemate, lsCont, 0, lnMonto + lnIgv, 0, "", 0, 0, 0, False)
          End If
          If lsCodCta <> pmJoyas(f, 0) Then
            lsCodCta = pmJoyas(f, 0)
            loVtaJoya.dInsertColocacEstado pmJoyas(f, 0), psFechaHora, gPigEstRematPFact, 0, 0, "Rematado Pendiente de Facturar", gColocCalendCodFFCF, 0, 0, 0, 0, 0, 0, 0, 0
          End If
          '**** Update Producto
          loVtaJoya.dUpdateProducto lsContx, , , gPigEstRematPFact, psFechaHora

          '**** Update ColocPigProceso
          Call loVtaJoya.dUpdateColocPigProceso(Val(pmJoyas(f, 2)), Val(pmJoyas(f, 5)), Trim(pmJoyas(f, 0)), Val(pmJoyas(f, 1)), psMovNro, , , gPigSituacionPendFacturar, , , , False)
    
          '**** Update ColocPigno
           loVtaJoya.dUpdateColocPigno Trim(pmJoyas(f, 0)), , gPigEstRematPFact
    Next
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)


    loRegPig.dCommitTrans
    mbTrans = False

Exit Sub

ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nAnularVentaJoyasRemate", "Error en Proceso de Anulación de Venta de Joyas"
    
End Sub

'***************************************************************************
'CMCPL - RESCATE DE SOBRANTE DE PIEZAS
'JAPP - 04/07/2003
'***************************************************************************
Public Sub nRescatePiezas(ByVal psCtaCod As String, ByVal psMovNro As String, ByVal pnComision As Currency, _
                ByVal pnValTasacion As Currency, ByVal psFechaHora As String, ByVal pnNroRemate As Integer, _
                ByVal pmJoyas As Variant, ByVal pnFilas As Integer)
                
Dim lnMovNro As Long
Dim oPigGraba As DPigActualizaBD
Dim mbTrans As Boolean
Dim f As Integer
    
    Set oPigGraba = New DPigActualizaBD
    oPigGraba.dBeginTrans
    mbTrans = True
    
    On Error GoTo ErrorPig
    
    '** Insert ColocEstado - CMCPL
    Call oPigGraba.dInsertColocacEstado(psCtaCod, psFechaHora, _
        gPigEstRescateSobranteP, 1, 0, "Rescate de Sobrante de Piezas", _
        gColocCalendCodFFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)

     '** Inserta Mov
    Call oPigGraba.dInsertMov(psMovNro, gPigOpeSobraPieza, _
         "Rescate de Sobrante de Piezas", gMovEstContabMovContable, _
         gMovFlagVigente, False)

    ' Obtiene nMovNro
    lnMovNro = oPigGraba.dGetnMovNro(psMovNro)

    '** Inserta MovCol
    Call oPigGraba.dInsertMovCol(lnMovNro, gPigOpeSobraPieza, psCtaCod, 0, _
         pnComision, 0, "", 0, gPigEstCancelPendRes, 0, False)

    '** Inserta MovColDet - Descargo de CustodiaRescatePendientes
    Call oPigGraba.dInsertMovColDet(lnMovNro, gPigOpeSobraPieza, psCtaCod, 0, _
         gColPigConceptoCodGarantPendRescAge, 0, pnValTasacion, False)
    If pnComision > 0 Then
        'Comision
        Call oPigGraba.dInsertMovColDet(lnMovNro, gPigOpeSobraPieza, psCtaCod, 0, _
             gColPigConceptoCodCustodia, 0, pnComision, False)
        'IGV
   '     Call oPigGraba.dInsertMovColDet(lnMovNro, gPigOpeDevJoyas, psCtaCod, _
   '     0, gColPigConceptoCodIgv, 0, pnIGV, False)
    End If
    
    For f = 0 To pnFilas
        '** Actualiza ColocPignoProceso
        Call oPigGraba.dUpdateColocPigProceso(pnNroRemate, 4, psCtaCod, _
             pmJoyas(f, 1), psFechaHora, -1, -1, 5, -1, -1, -1, False)
        '** Actualiza ColocPigJoyaTasacion
        Call oPigGraba.dUpdateColocPigJoyaTasacion(psCtaCod, 3, _
             pmJoyas(f, 1), "@", -1, -1, -1, -1, -1, -1, -1, -1, "@", "@", _
             gPigUbicaSinCaja, "@", False)
    Next
    
    oPigGraba.dCommitTrans
    mbTrans = False
    Set oPigGraba = Nothing
    Exit Sub
    
ErrorPig:
    If mbTrans Then
        oPigGraba.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRescateLotes", "Error en Funcion de Rescate de Lotes"

End Sub
'***************************************************************************
'CMCPL EXTORNO RESCATE SOBRANTE PIEZAS 'JAPP - 04/07/2003
'***************************************************************************
Public Sub nExtornoRescateSobrantePiezas(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        ByVal pnCodAge As String, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
Dim loRegPig As DPigActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim mbTrans As Boolean
Dim oRema As DPigFunciones
Dim oRema2 As DPigFunciones
Dim lnRemate As Integer
Dim lnUbicacion As Integer
Dim lnItemPieza As Integer
Dim N As Integer


    lsOpeCod = gPigOpeExtPagoSobranteP
    
    Set oRema = New DPigFunciones
    lnRemate = oRema.GetRemate(psCtaCod)
    
    Set oRema = New DPigFunciones
    lnUbicacion = oRema.GetUbicacion(pnCodAge)
    
    Set oRema = New DPigFunciones
    lnItemPieza = oRema.GetNroPieza(psCtaCod)
    
    Set oRema = Nothing
    'On Error GoTo ErrorModPig
    Set loRegPig = New DPigActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
     
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Extorno Rescate de Sobrante de Piezas", gMovEstContabMovContable, gMovFlagVigente, False)
    
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMonto, 0, "", 0, gPigEstRescateSobranteP, 0, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
    
    For N = 1 To lnItemPieza
        Call loRegPig.dUpdateColocPigProceso(lnRemate, 4, psCtaCod, _
             N, psFechaHora, -1, -1, 1, -1, -1, -1, False)
    
        Call loRegPig.dUpdateColocPigJoyaTasacion(psCtaCod, 3, _
             N, "@", -1, -1, -1, -1, -1, -1, -1, -1, "@", "@", _
             lnUbicacion, "@", False)
    Next
    
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nExtornoRescateJoyas", "Error en Proceso de Extorno de Rescate de Sobrante de Piezas"

End Sub



