VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMColPContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mnEjecutaBatch As Integer
Dim mbTrans As Boolean

' Registra Contrato Pignoraticio
Public Function nRegistraContratoPignoraticio(ByVal psAgencia As String, ByVal pnMoneda As Moneda, _
                ByVal prPers As ADODB.Recordset, ByVal pnTasa As Double, ByVal pnSaldo As Currency, _
                ByVal psFechaHora As String, ByVal pnPlazo As Integer, ByVal psFecVenc As String, _
                ByVal pnOroBruto As Currency, ByVal pnOroNeto As Currency, ByVal pnValTasacion As Currency, _
                ByVal pnPiezas As Integer, ByVal psTipoContrato As String, ByVal psLote As String, _
                ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, ByVal pnKi21 As Double, _
                ByVal psMovNro As String, ByVal pnIntAdelantado As Currency, ByVal pnCostoTasac As Currency, _
                ByVal pnCostoCustodia As Currency, ByVal pnImpuesto As Currency) As String
'** Genera Nuevo Nro Contrato
'** dInsertCreditoPignoraticio
'** dInsertColocEstado
'** dInsertCreditoPigRelaPersona
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************
Dim loGen As New COMDConstSistema.DCOMGeneral
Dim lsCuenta As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

On Error GoTo ErrorRegPig

lsOpeCod = gColPOpeRegContrato

'Genera Nro Cuenta

lsCuenta = loGen.GeneraNuevaCuenta(psAgencia, gColConsuPrendario, pnMoneda)
Set loGen = Nothing

    loRegPig.dBeginTrans
    mbTrans = True

    Call loRegPig.dInsertCredPignoraticio(lsCuenta, pnTasa, pnSaldo, psFechaHora, pnPlazo, _
        psFecVenc, pnOroBruto, pnOroNeto, pnValTasacion, pnPiezas, psTipoContrato, psLote, _
        pnki14, pnki16, pnKi18, pnKi21, psMovNro, False)
        
    Call loRegPig.dInsertColocacEstado(lsCuenta, psFechaHora, gColPEstRegis, 1, pnSaldo, _
        "Registro Contrato", gColocCalendCodPFCF, 0, 0, pnPlazo, "0", 0, 0, 0, 0, False)

    'inserta personas para cada registro de prPersonas
    Do While Not prPers.EOF
        Call loRegPig.dInsertProductoPersona(lsCuenta, prPers("cPersCod"), gColRelPersTitular, False)
        prPers.MoveNext
    Loop
    'inserta ColocCalendario
    Call loRegPig.dInsertColocCalendario(lsCuenta, 1, gColocCalendAplDesembolso, 1, psFecVenc, 0, "Reg Contrato", "", False)
   '**********
    'inserta ColocCalendDetPig - Capital
    Call loRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplDesembolso, 1, gColPConceptoCodCapital, pnSaldo, 0, "", False)
    
    'inserta ColocCalendDetPig - Interes Adelantado
    Call loRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplDesembolso, 1, gColPConceptoCodInteresCompensatorio, pnIntAdelantado, 0, "", False)
    
    'inserta ColocCalendDetPig - Costo Tasacion
    If pnCostoTasac > 0 Then
        Call loRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplDesembolso, 1, gColPConceptoCodTasacion, pnCostoTasac, 0, "", False)
    End If
    'inserta ColocCalendDetPig - Costo Custodia
    If pnCostoCustodia > 0 Then
        Call loRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplDesembolso, 1, gColPConceptoCodCustodia, pnCostoCustodia, 0, "", False)
    End If
    'inserta ColocCalendDetPig - Costo Impuesto
    If pnImpuesto > 0 Then
        Call loRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplDesembolso, 1, gColPConceptoCodImpuesto, pnImpuesto, 0, "", False)
    End If
   '**********
    ' inserta Mov '
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Registro Contrato", gMovEstContabMovContable, gMovFlagVigente, False)
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    ' inserta MovCol '
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, lsCuenta, 1, pnSaldo, 0, "GMIC", pnPlazo, 0, 0, False)
    ' Monto Prestamo(Capital)
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lsCuenta, 1, gColPConceptoCodCapital, 1, pnSaldo, False)
    ' Valor Tasacion
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lsCuenta, 1, gColPConceptoCodTasacion, 1, pnValTasacion, False)

    ' Gramos de Oro
    If pnki14 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lsCuenta, 1, gColPConceptoCodOro14, 1, pnki14, False)
    End If
    If pnki16 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lsCuenta, 1, gColPConceptoCodOro16, 1, pnki16, False)
    End If
    If pnKi18 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lsCuenta, 1, gColPConceptoCodOro18, 1, pnKi18, False)
    End If
    If pnKi21 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lsCuenta, 1, gColPConceptoCodOro21, 1, pnKi21, False)
    End If

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    nRegistraContratoPignoraticio = lsCuenta
    
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Function

ErrorRegPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "
End Function



Public Sub nModificaLoteCredPignoraticio(ByVal psctacod As String, ByVal psFechaHora As String, _
            ByVal psLote As String, ByVal psMovNro As String, Optional pbEjecBatch As Boolean = False)


'** Actualiza Producto (Nro Transacc)
'** Actualiza Colocaciones (Ultima Modificacion)
'** Actualiza ColocPig (Lote)
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long

'On Error GoTo ErrorModPig


    loRegPig.dBeginTrans
    mbTrans = True
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , , , -2, False) ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Actualiza ColocPig
    Call loRegPig.dUpdateColocPignoraticio(psctacod, , , , , , , , , , psLote, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, gColPOpeModifDescJoyas, "Modifica Lote", gMovEstContabNoContable, gMovFlagVigente, False)
    
    '** Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)

    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, gColPOpeModifDescJoyas, psctacod, 0, 0, 0, "", 0, 0, 0, False)

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nAnulaCredPignoraticio(ByVal psctacod As String, ByVal pnEstado As ColocEstado, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal pnMovNroAnt As Long, _
        Optional pbEjecBatch As Boolean = False, Optional psMotAnula As String = "")

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'** Actualiza Mov ( MovNro del Registro-Anterior )
'** dInsertMovRef
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = gColPOpeAnula

'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    '****RECO 20140113-INC1401130010****************************************************************************
    'Call loRegPig.dUpdateProducto(psctacod, , , pnEstado, psFechaHora, -2, False)  ' (-2) aumenta el ste transac
    Call loRegPig.dUpdateProducto(psctacod, , 0, pnEstado, psFechaHora, -2, False)
    '****END RECO***********************************************************************************************
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstAnula, 1, 0, "Anula Contrato", _
        gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    'Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Anula Cred Pign " & psMotAnula, gMovEstContabNoContable, gMovFlagVigente, False)
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, psMotAnula, gMovEstContabNoContable, gMovFlagVigente, False)

    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 0, 0, 0, "", 0, 0, 0, False)
     
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
    
    'WIOR 20140203 ******************************
    Call loRegPig.ActualizaCredVinculados(psctacod, "", 1, 3)
    'WIOR FIN ***********************************
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nBloqueo_DesBloqueoCredPignoraticio(ByVal psctacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pbBloqIni As Boolean, Optional ByVal psDescripcion As String = "@", _
        Optional ByVal psMovNroBloqueo As String = "@", Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertProductoBloqueos
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
' ******************************************************Falta cambiar
'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , , psFechaHora, -2, False)  ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    If pbBloqIni = False Then ' Se va ha Bloquear
        '** Insert ProductoBloqueos
        Call loRegPig.dInsertProductoBloqueos(psctacod, 11, 3, psMovNro, psDescripcion, , False)
        '** Inserta Mov
        Call loRegPig.dInsertMov(psMovNro, geColPBloqueoJoyas, "Bloqueo de Joyas", gMovEstContabNoContable, gMovFlagVigente, False)
        ' Obtiene nMovNro
        lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        '** Inserta MovCol
        Call loRegPig.dInsertMovCol(lnMovNro, geColPBloqueoJoyas, psctacod, 0, 0, 0, "", 0, 0, 0, False)
    Else ' Se va ha Desbloquear
        '** Update ProductoBloqueos
        Call loRegPig.dUpdateProductoBloqueos(psctacod, psMovNroBloqueo, psMovNro, False)
        '** Inserta Mov
        Call loRegPig.dInsertMov(psMovNro, geColPDesbloqueoJoyas, "DesBloqueo de Joyas", gMovEstContabNoContable, gMovFlagVigente, False)
        ' Obtiene nMovNro
        lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        '** Inserta MovCol
        Call loRegPig.dInsertMovCol(lnMovNro, geColPDesbloqueoJoyas, psctacod, 0, 0, 0, "", 0, 0, 0, False)
    End If
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "
End Sub

Public Sub nDesembolsoCredPignoraticio(ByVal psctacod As String, ByVal pnSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal pnMontoEntregar As Currency, _
        ByVal pnInteresCompensat As Currency, ByVal pnImpuesto As Currency, _
        ByVal pnCostoTasacion As Currency, ByVal pnCostoCustodia As Currency, _
        ByVal pbITFAplica As Boolean, ByVal pbITFAsumidoCreditos As Boolean, _
        ByVal pnMontoITf1 As Currency, ByVal pnMontoITf2 As Currency, ByVal pnMontoDesembolsar As Currency, _
        Optional pbEjecBatch As Boolean = False)
    
'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

'Codigo de Operacion
lsOpeCod = gColPOpeDesembolsoEFE

'On Error GoTo ErrorModPig

    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , gColPEstDesem, psFechaHora, -2, False)    ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstDesem, 1, pnSaldoCap, "Desembolso Contrato", _
        gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Desembolso Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    
    '******** modificado 01/12/2004
    
    '** Inserta MovCol
    'Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 1, pnMontoEntregar, 0, "", 0, 0, pnSaldoCap, False)
    
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoDesembolsar, 0, "", 0, 0, pnSaldoCap, False)

    '** Inserta MovColDet -  Monto Entregar
    '****ORIGINAL ANTES DEL 07/02/2005
    
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodMontoEntregar, 1, pnMontoDesembolsar, False)
    
    
    '*****antes del 15/01/2005
    'Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodMontoEntregar, 1, pnMontoEntregar, False)
    
    'Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodCapital, 1, pnMontoEntregar, False)
    
    '********
    '** Inserta MovColDet -  Interes Compensatorio
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresCompensatorio, 1, pnInteresCompensat, False)
    
    '** Inserta MovColDet -  Impuesto
    If pnImpuesto > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodImpuesto, 1, pnImpuesto, False)
    End If
    '** Inserta MovColDet -  Costo Custodia
    If pnCostoCustodia > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCustodia, 1, pnCostoCustodia, False)
    End If
    '** Inserta MovColDet -  Costo Tasacion
    If pnCostoTasacion > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodTasacion, 1, pnCostoTasacion, False)
    End If
    
    If pbITFAplica Then
        If Not pbITFAsumidoCreditos Then
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, pnMontoITf1 + pnMontoITf2, 0, "", 0, 0, pnSaldoCap, False)
        
            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, gConcITFCliente, 1, pnMontoITf1, False)
            
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, 22, 1, pnMontoITf2, False)
        Else
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, pnMontoITf1 + pnMontoITf2, 0, "", 0, 0, pnSaldoCap, False)
        
            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, gConcITFAsumido, 1, pnMontoITf1, False)
            
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, 22, 1, pnMontoITf2, False)
        End If
    End If
    
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nRescataJoyaCredPignoraticio(ByVal psctacod As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, ByVal pnKi21 As Double, _
        ByVal pnDiasCustodia As Integer, ByVal pnValTasac As Double, _
        ByVal psOperacion As String, Optional pbEjecBatch As Boolean = False, _
        Optional ByRef pnMovNro As Long = 0, Optional loRegPig2 As COMDColocPig.DCOMColPActualizaBD = Nothing, Optional pbTrans As Boolean = False)
        'RECO20140206 ERS002 SE AGREGO LOS PARAMETROS:pnMovNro,loRegPig2


'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

'Codigo de Operacion
'lsOpeCod = geColPDevJoyas
lsOpeCod = psOperacion

On Error GoTo ErrorModPig
    'RECO20140206 ERS002*********************************
    If pbTrans = False Then
        loRegPig.dBeginTrans
    Else
        Set loRegPig = loRegPig2
    End If
    'loRegPig.dBeginTrans
    mbTrans = True
    'RECO FIN********************************************


    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , gColPEstCance, psFechaHora, -2, False)    ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstCance, 1, 0, "Entrega Joya", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    'RECO 20140207 ERS002**********************************
    If pbTrans Then
        lnMovNro = pnMovNro
    Else
        Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Entrega Joya Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
        
        ' Obtiene nMovNro
        lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        pnMovNro = lnMovNro 'WIOR 20130301
    End If
    'RECO FIN************************************************

    'Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Entrega Joya Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, 0, 0, "", 0, 0, 0, False)
    '** Insert MovColDet Valor Tasacion
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodTasacion, 1, pnValTasac, False)

    '** Inserta MovColDet ' Oro 14/16/18/21
    If pnki14 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro14, 1, pnki14, False)
    End If
    If pnki16 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro16, 1, pnki16, False)
    End If
    If pnKi18 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro18, 1, pnKi18, False)
    End If
    If pnKi21 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro21, 1, pnKi21, False)
    End If

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    'RECO 20140207 ERS002**********************************
    If pbTrans = False Then
        loRegPig.dCommitTrans
     End If
     'loRegPig.dCommitTrans
     mbTrans = False
    'RECO FIN************************************************

    
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
    'RECO20140208 ERS002*************************
        If pbTrans = False Then
            loRegPig.dRollbackTrans
        End If

'        loRegPig.dRollbackTrans
    'RECO FIN*************************************
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nRetasacionCredPignoraticio(ByVal psctacod As String, ByVal psMovNro As String, ByVal pnValTasacion As Currency, ByVal pnPiezas As Integer, _
                                        ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, ByVal pnKi21 As Double, _
                                         ByVal prjoyas As ADODB.Recordset, Optional pnOroBruto As Double, Optional pnOroNeto As Double)
                                        
Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim loPigDetJoya As New COMDColocPig.DCOMColPActualizaBD
Dim loCalcula As New COMDColocPig.DCOMColPCalculos
Dim lrJoyasAnt As New ADODB.Recordset
Dim lnMovNro As Long, nNumRet As Integer
Dim lsOpeCod As String

    
    lsOpeCod = "121300"
    nNumRet = loCalcula.nNumRetasacion(psctacod)
    'todocompleta verificar funciones
    'Set lrJoyasAnt = loPigDetJoya.dObtieneDatosCreditoPignoraticioJoyasDet(psCtaCod)
    'Call loRegPig.dInsertHistColocPigJoyaDetRet(lrJoyasAnt, nNumRet, psMovNro, psCtaCod)
    Call loRegPig.dUpdateColocPignoraticio(psctacod, pnOroBruto, pnOroNeto, pnPiezas, pnValTasacion)
    Call loRegPig.dUpdateColocPigJoya(psctacod, pnki14, pnki16, pnKi18, pnKi21)
    Call loRegPig.dUpdateColocPigJoyaDet(prjoyas, psctacod)
      '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Retasacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnValTasacion, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodTasacion, 1, pnValTasacion, False)
    Set loCalcula = Nothing
    'Set lrJoyasAnt = Nothing
    Set loRegPig = Nothing
    Set loPigDetJoya = Nothing
    
End Sub

Public Sub nRenovacionCredPignoraticio(ByVal psctacod As String, _
        ByVal pnNewSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal psFecVenci As String, ByVal pnNewPlazo As Integer, _
        ByVal pnMontoTransac As Currency, ByVal pnCapPag As Currency, _
        ByVal pnIntVenc As Currency, ByVal pnCustodiaVenc As Currency, _
        ByVal pnPreparaRemate As Currency, ByVal pnIntCompensat As Currency, _
        ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, _
        ByVal pnDiasAtraso As Integer, ByVal pnNroRenovacion As Integer, _
        ByVal pnDiasCambCart As Integer, ByVal pnValorTasacion As Currency, _
        ByVal pnOpeCod As Long, ByVal psOpeDesc As String, _
        ByVal psPersCodCMAC As String, ByVal pnGastoCorrespondencia As Currency, _
        ByVal pbITFAplica As Boolean, ByVal pbITFAsumidoCreditos As Boolean, _
        ByVal pnMontoITF As Currency, _
        Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnIntMora As Currency = 0, _
        Optional ByVal pnMontoCostoNotificacion As Currency = 0, _
        Optional ByRef pnMovNro As Long = 0, Optional ByVal pnPagosInt As Double = 0, Optional ByVal psFechaVencAnt As String = "@", _
        Optional ByVal pnTipoPago As Integer = 1, Optional pnMovNroRVD As Long = 0, Optional pnMovNroPend As Long = 0, _
        Optional ByVal psCtaCodCargo As String = "", Optional ByRef pMatAho As Variant = "")
        
        'WIOR 20130301 AGREGO pnMovNro
        'PEAC 20161213 se agregó psFechaVencAnt
        'CTI4 ERS0112020 Add: pnTipoPago, pnMovNroRVD, pnMovNroPend, psCtaCodCargo, pMatAho

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsOpeDesc As String
lsOpeDesc = ""
Dim lsOpeCodCargoCta As String   'CTI4 ERS0112020
'Codigo de Operacion
Select Case pnOpeCod
    Case gColPOpeRenovacEFE  ' En la CMAC
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeRenovNorEFE
        Else
            lsOpeCod = gColPOpeRenovMorEFE
        End If
    Case gColPOpeRenovacCHQ
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeRenovNorCHQ
        Else
            lsOpeCod = gColPOpeRenovMorCHQ
        End If
    Case gColPOpeRenovacEnOtCjEFE  ' En otra CMAC
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeRenovNorEnOtCjEFE
        Else
            lsOpeCod = gColPOpeRenovMorEnOtCjEFE
        End If
    Case gColPOpeRenovacEnOtCjCHQ
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeRenovNorEnOtCjCHQ
        Else
            lsOpeCod = gColPOpeRenovMorEnOtCjCHQ
        End If
     Case gColPOpeRenovacVoucher 'CTI4 ERS0112020
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeRenovNorVoucher
        Else
            lsOpeCod = gColPOpeRenovMorVoucher
        End If
        lsOpeDesc = " con voucher"
    Case gColPOpeRenovacCargoCta 'CTI4 ERS0112020
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeRenovNorCargoCta
            lsOpeCodCargoCta = gAhoCargoCtaRenovaPignoNor
        Else
            lsOpeCod = gColPOpeRenovMorCargoCta
            lsOpeCodCargoCta = gAhoCargoCtaRenovaPignoMor
        End If
End Select

'On Error GoTo ErrorModPig
'Para verificar la fecha de Vencimiento -- AVMM --11-06-2007************
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
Dim Rs As ADODB.Recordset
Dim ldVarFecVencimiento As Date
Dim ldFechaAct As Date
Set oCon = New COMConecta.DCOMConecta
    If oCon.AbreConexion Then
        sSql = "select dVenc from Colocaciones where cCtaCod='" & psctacod & "' "
        Set Rs = New ADODB.Recordset
        Set Rs = oCon.CargaRecordSet(sSql)
            If Not (Rs.EOF And Rs.BOF) Then
                ldVarFecVencimiento = Rs("dVenc")
                ldFechaAct = Format(psFechaHora, "dd/mm/yyyy")
            End If
        Set Rs = Nothing
    End If
Set oCon = Nothing
'***********************************************************************

'ARCV 04-07-2007
'If Format(ldFechaAct, "dd/mm/yyyy") <> Format(ldVarFecVencimiento, "dd/mm/yyyy") Then
'    psFecVenci = Format$(ldFechaAct + pnNewPlazo, "mm/dd/yyyy")
'End If

'peac 20070820
'If Format(ldFechaAct, "mm/dd/yyyy") > Format(ldVarFecVencimiento, "mm/dd/yyyy") Then
'    psFecVenci = Format$(ldFechaAct + pnNewPlazo, "mm/dd/yyyy") 'Vencidos
'Else
'    psFecVenci = Format$(ldVarFecVencimiento + pnNewPlazo, "mm/dd/yyyy") 'No Vencidos
'End If
psFecVenci = Format$(ldFechaAct + pnNewPlazo, "mm/dd/yyyy")
'end peac
'-----------

    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , pnNewSaldoCap, gColPEstRenov, psFechaHora, -2, False)          ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, pnNewPlazo, psFecVenci, , , psMovNro, , False, psFechaVencAnt)

    '** Actualiza ColocPignoraticio
    Call loRegPig.dUpdateColocPignoraticio(psctacod, , , , , pnNroRenovacion, 0, , , , , False)
    
    '*** PEAC 20080515
    '** Actualiza ColocPigNotifi
    Call loRegPig.dUpdateColocPigNotifi(psctacod, , 1, psMovNro, False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstRenov, 1, pnMontoTransac, "Renovacion Contrato", _
        gColocCalendCodPFCF, 0, 0, pnNewPlazo, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Renovacion Cred Pign" & lsOpeDesc, gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    pnMovNro = lnMovNro 'WIOR 20130301
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, pnDiasAtraso, "GMIC", pnNewPlazo, gColPEstRenov, pnNewSaldoCap, False)

    '** Inserta MovColDet -  Capital
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCapital, 1, pnCapPag, False)

    '** Inserta MovColDet -  Interes Vencido
    If pnIntVenc > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresVencido, 1, pnIntVenc, False)
    End If
    
    '** Inserta MovColDet -  Interes Moratorio
    If pnIntMora > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresMoratorio, 1, pnIntMora, False)
    End If
   
    '** Inserta MovColDet -  Custod Venc
    If pnCustodiaVenc > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCustodiaVencida, 1, pnCustodiaVenc, False)
    End If
    '** Inserta MovColDet -  Preparac Remate
    If pnPreparaRemate > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodPreparaRemate, 1, pnPreparaRemate, False)
    End If
    '** Inserta MovColDet -  Interes Compensatorio (Adelantado)
    If pnIntCompensat > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresCompensatorio, 1, pnIntCompensat, False)
    End If
    '** Inserta MovColDet -  Impuesto
    If pnImpuesto > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodImpuesto, 1, pnImpuesto, False)
    End If
    '** Inserta MovColDet -  Costo Custodia
    If pnCostoCustodia > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCustodia, 1, pnCostoCustodia, False)
    End If
    
    '** Inserta MovColDet -  Gasto Correspondencia
    If pnGastoCorrespondencia > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, 2215, 1, pnGastoCorrespondencia, False)
        
        'Actualiza ColocCalendDetPig
        lsSQL = "UPDATE ColocCalendDetPig Set nMontoPagado = nMontoPagado + " & pnGastoCorrespondencia & _
                " WHERE cCtaCod ='" & psctacod & "' And nNroCalen = 1 And nColocCalendApl = 0 " & _
                " And nCuota = 1 And nPrdConceptoCod = 2215 "

        loRegPig.dEjecutar lsSQL
        
    End If

    'ARCV 14-03-2007
    If pnMontoCostoNotificacion > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, "2103", 1, pnMontoCostoNotificacion, False)
    End If
    
    If pbITFAplica Then
        If Not pbITFAsumidoCreditos Then
            If Left(gColPOpeRenovacEnOtCjEFE, 4) = Left(lsOpeCod, 4) Then
                '** Inserta MovCol
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroCMACCred, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                '** Inserta MovColDet -  Monto Entregar
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroCMACCred, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
            Else
                If pnTipoPago = gColocTipoPagoVoucher Then 'CTI4 ERS0112020
                    Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroOpePignoVoucher, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                    Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroOpePignoVoucher, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
                ElseIf pnTipoPago = gColocTipoPagoCargoCta Then 'CTI4 ERS0112020
                    Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroOpePignoCargoACta, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                    Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroOpePignoCargoACta, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
                Else
                    '** Inserta MovCol
                    Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                    '** Inserta MovColDet -  Monto Entregar
                    Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
                End If
            End If
        Else
            If pnTipoPago = gColocTipoPagoVoucher Then 'CTI4 ERS0112020
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroOpePignoVoucher, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroOpePignoVoucher, psctacod, 1, gConcITFAsumido, 1, pnMontoITF, False)
            ElseIf pnTipoPago = gColocTipoPagoCargoCta Then 'CTI4 ERS0112020
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroOpePignoCargoACta, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroOpePignoCargoACta, psctacod, 1, gConcITFAsumido, 1, pnMontoITF, False)
            Else
                '** Inserta MovCol
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
            
                '** Inserta MovColDet -  Monto Entregar
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, gConcITFAsumido, 1, pnMontoITF, False)
            End If
        End If
    End If

    '**** Cambio de Cartera - Regresa a Normal
    If pnDiasAtraso > pnDiasCambCart Then
        'Cambio de Cartera - Regresa a Normal
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCambioCarteraMorosoNormal, 1, pnNewSaldoCap, False)
        'Cambio a Valor de Tasación
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCambioTasacionMorosoNormal, 1, pnValorTasacion, False)
    End If
    
    '** Inserta MovCMAC - Para Recepcion de Llamada de cmac
    If pnOpeCod = gColPOpeRenovacEnOtCjEFE Or pnOpeCod = gColPOpeRenovacEnOtCjCHQ Then
        Call loRegPig.dInsertMovCMAC(lnMovNro, psPersCodCMAC, Format$(gTpoIFCmac, "00"), CInt(Mid(psctacod, 9, 1)), "", _
                "", lsOpeCod, pnMontoTransac, False)
    End If
    
    '*** PEAC 20161120
    If pnPagosInt > 0 Then
        Call loRegPig.dUpdateColocPigPagosParciales(psctacod, lnMovNro)
    End If
  
     
    'CTI4 ERS0112020 **************
    If pnTipoPago = gColocTipoPagoVoucher Then
        loRegPig.actualizarRelacionVoucherCaptacion lnMovNro, pnMovNroRVD
        If pnMovNroPend <> -1 Then loRegPig.ActualizaMovPendientesRend pnMovNroPend, pnMontoTransac + pnMontoITF
    End If
    If pnTipoPago = gColocTipoPagoCargoCta Then
        loRegPig.CapCargoCuentaAho pMatAho, psCtaCodCargo, pnMontoTransac + pnMontoITF, lsOpeCodCargoCta, psMovNro, "Cargo a Cuenta Renov. Pigno.: " & psctacod, psFechaHora, , , , , , , , , , , , , , True, CDbl(pnMontoITF), True, gITFCobroCargo
    End If
    'END CTI4
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

'*** PEAC 20080412
Public Sub ModificarEstadoAdj(ByVal psCta As String, ByVal psNumProceso As String)

    '** Actualiza ColocPignoraticio
    Dim loExcluyePigAdj As New COMDColocPig.DCOMColPActualizaBD
    Call loExcluyePigAdj.PignoExcluyeAdjudica(psCta, psNumProceso)
    Set loExcluyePigAdj = Nothing

Exit Sub

End Sub

'*** PEAC 20080412
Public Sub nMarcaCredPignoAdjudica(ByVal pdFecCorte As Date, ByVal psAge As String, _
                                ByVal pdFechaLey As Date, Optional pdFecNotifi As Date, _
                                Optional pnNotifica As Integer, Optional pnEstado As Integer, _
                                Optional psMovNroNotifi As String, Optional psNumProceso As String)

'    '** Actualiza ColocPignoraticio
'    Dim loRegPigAdj As New COMDColocPig.DCOMColPActualizaBD
'    Call loRegPigAdj.dActualizaPignoAdjudica(pdFecCorte, psAge, pdFechaLey)
'    Set loRegPigAdj = Nothing
    
    '** Actualiza ColocPignoraticio
    Dim loRegPigAdj As New COMDColocPig.DCOMColPActualizaBD
    Call loRegPigAdj.dActualizaPignoAdjudica(pdFecCorte, psAge, pdFechaLey, pdFecNotifi, pnNotifica, _
                                            pnEstado, psMovNroNotifi, psNumProceso)
    Set loRegPigAdj = Nothing

Exit Sub

End Sub


Public Sub nCancelacionCredPignoraticio(ByVal psctacod As String, ByVal pnNewSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnMontoTransac As Currency, ByVal pnCapPag As Currency, ByVal pnIntVenc As Currency, _
        ByVal pnCustodiaVenc As Currency, ByVal pnPreparaRemate As Currency, _
        ByVal pnImpuesto As Currency, ByVal pnDiasAtraso As Integer, _
        ByVal pnDiasCambCart As Integer, ByVal pnValorTasacion As Currency, _
        ByVal pnOpeCod As Long, ByVal psOpeDesc As String, ByVal psPersCodCMAC As String, _
        ByVal pbITFAplica As Boolean, ByVal pbITFAsumidoCreditos As Boolean, _
        ByVal pnMontoITF As Currency, _
        Optional pbEjecBatch As Boolean = False, Optional pnGastoCorrespondencia As Currency = 0, _
        Optional ByVal pnIntMora As Currency = 0, Optional ByVal pnMontoCostoNotificacion As Currency = 0, _
        Optional ByVal pnInteresAdel As Currency, _
        Optional ByRef pnMovNro As Long = 0, _
        Optional loRegPig2 As COMDColocPig.DCOMColPActualizaBD = Nothing, Optional pbTrans As Boolean = False, _
        Optional ByVal pnTipoPago As Integer = 1, Optional pnMovNroRVD As Long = 0, Optional pnMovNroPend As Long = 0, _
        Optional ByVal psCtaCodCargo As String = "", Optional ByRef pMatAho As Variant = "")
        'RECO20140208 ERS022 SE AGREGO:loRegPig2,pbTrans
        'WIOR 20130301 AGREGO pnMovNro
        'PEAC 20070926
        '***PEAC 20080118 se agrego pnInteresAdel
        'CTI4 ERS0112020 Add: pnTipoPago, pnMovNroRVD, pnMovNroPend, psCtaCodCargo, pMatAho

'ByVal pnInteresAdel As Currency

'** Actualiza Producto (Estado, nSaldoCap,dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Plazo, Vencimiento, Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsOpeCodCargoCta As String 'CTI4 ERS0112020
Dim lsOpeDesc As String
lsOpeDesc = "Cancelacion Cred Pign"
'Codigo de Operacion
Select Case pnOpeCod
    Case gColPOpeCancelacEFE   ' En la CMAC
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeCancelNorEFE
        Else
            lsOpeCod = gColPOpeCancelMorEFE
        End If
      Case gColPOpeCancelacCHQ
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeCancelNorCHQ
        Else
            lsOpeCod = gColPOpeCancelMorCHQ
        End If
    Case gColPOpeCancelacEnOtCjEFE    ' En otra CMAC
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeCanceNorEnOtCjEFE
        Else
            lsOpeCod = gColPOpeCanceMorEnOtCjEFE
        End If
    Case gColPOpeCancelacEnOtCjCHQ
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeCanceNorEnOtCjCHQ
        Else
            lsOpeCod = gColPOpeCanceMorEnOtCjCHQ
        End If
    Case gColPOpeCancelVoucher 'CTI4 ERS0112020
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeCancelNorVoucher
        Else
            lsOpeCod = gColPOpeCancelMorVoucher
        End If
        lsOpeDesc = "Cancelacion Cred Pign con voucher"
    Case gColPOpeCancelCargoCta 'CTI4 ERS0112020
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeCancelNorCargoCta
            lsOpeCodCargoCta = gAhoCargoCtaCancelaPignoNor
        Else
            lsOpeCod = gColPOpeCancelMorCargoCta
            lsOpeCodCargoCta = gAhoCargoCtaCancelaPignoMor
        End If
End Select

On Error GoTo ErrorModPig
    'RECO20140206 ERS002*********************************
    If pbTrans = False Then
        loRegPig.dBeginTrans
    Else
        Set loRegPig = loRegPig2
    End If
    'loRegPig.dBeginTrans
    mbTrans = True
    'RECO FIN********************************************

    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , pnNewSaldoCap, gColPEstDifer, psFechaHora, -2, False)           ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

'    '*** PEAC 20080515
'    '** Actualiza ColocPignoraticio
'    Call loRegPig.dUpdateColocPignoraticio(psctacod, , , , , , , , , , , False)

    '*** PEAC 20080515
    '** Actualiza ColocPigNotifi
    Call loRegPig.dUpdateColocPigNotifi(psctacod, , 1, psMovNro, False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstDifer, 1, pnMontoTransac, "Cancelacion Contrato", _
        gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    'RECO 20140206 ERS002****************************************************
    If pbTrans = False Then
        '** Inserta Mov
        Call loRegPig.dInsertMov(psMovNro, lsOpeCod, lsOpeDesc, gMovEstContabMovContable, gMovFlagVigente, False)
        ' Obtiene nMovNro
        lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        pnMovNro = lnMovNro 'WIOR 20130301
    Else
        lnMovNro = pnMovNro
        'lsOpeCod = gColPOpeCancelAmplEFE 'RECO20140911 CANCELACION POR AMPLIACION
    End If
    'Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Cancelacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    'RECO FIN ******************************************************************

    
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    pnMovNro = lnMovNro 'WIOR 20130301
    
    '** Inserta MovCol
    'Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, 0, "GMIC", 0, 0, 0, False)
    '*** PEAC 20080701 - pnDiasAtraso
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, pnDiasAtraso, "GMIC", 0, 0, 0, False)


    '** Inserta MovColDet -  Capital
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCapital, 1, pnCapPag, False)

    '** Inserta MovColDet -  Valor Tasacion
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodTasacion, 1, pnValorTasacion, False)

    '** Inserta MovColDet -  Interes Compensatorio (al vencimiento)
    If pnInteresAdel > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresCompensatorio, 1, pnInteresAdel, False)
    End If


    '** Inserta MovColDet -  Interes Vencido
    If pnIntVenc > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresVencido, 1, pnIntVenc, False)
    End If
    
     '** Inserta MovColDet -  Interes Moratorio
    If pnIntMora > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresMoratorio, 1, pnIntMora, False)
    End If
  
    'PEAC 20070926
    '*** inserta en MovColDet - costo de notificacion
    If pnMontoCostoNotificacion > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, "2103", 1, pnMontoCostoNotificacion, False)
    End If
    
    '** Inserta MovColDet -  Custod Venc
    If pnCustodiaVenc > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCustodiaVencida, 1, pnCustodiaVenc, False)
    End If
    '** Inserta MovColDet -  Impuesto
    If pnImpuesto > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodImpuesto, 1, pnImpuesto, False)
    End If
    '** Inserta MovColDet -  Preparac Remate
    If pnPreparaRemate > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodPreparaRemate, 1, pnPreparaRemate, False)
    End If
    
        '** Inserta MovColDet -  Gasto Correspondencia
    If pnGastoCorrespondencia > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, 2215, 1, pnGastoCorrespondencia, False)
               
        'Actualiza ColocCalendDetPig
        lsSQL = "UPDATE ColocCalendDetPig Set nMontoPagado = nMontoPagado + " & pnGastoCorrespondencia & _
                " WHERE cCtaCod ='" & psctacod & "' And nNroCalen = 1 And nColocCalendApl = 0 " & _
                " And nCuota = 1 And nPrdConceptoCod = 2215 "

        loRegPig.dEjecutar lsSQL
        
        
    End If
    
    '** Inserta MovCMAC - Para Recepcion de Llamada de cmac
    If pnOpeCod = gColPOpeCancelacEnOtCjEFE Or pnOpeCod = gColPOpeCancelacEnOtCjCHQ Then
        Call loRegPig.dInsertMovCMAC(lnMovNro, psPersCodCMAC, Format$(gTpoIFCmac, "00"), CInt(Mid(psctacod, 9, 1)), "", _
                "", lsOpeCod, pnMontoTransac, False)
    End If
    
    '+++++++++++++   itf  +++++++++++++++++
    If pbITFAplica Then
        If Not pbITFAsumidoCreditos Then
            If Left(gColPOpeCancelacEnOtCjEFE, 4) = Left(lsOpeCod, 4) Then
                '** Inserta MovCol
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroCMACCred, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                '** Inserta MovColDet -  Monto Entregar
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroCMACCred, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
            Else
                If pnTipoPago = gColocTipoPagoVoucher Then 'CTI4 ERS0112020
                    Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroOpePignoVoucher, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                    Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroOpePignoVoucher, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
                ElseIf pnTipoPago = gColocTipoPagoCargoCta Then 'CTI4 ERS0112020
                    Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroOpePignoCargoACta, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                    Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroOpePignoCargoACta, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
                Else
                    ''** Inserta MovCol
                    Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                    ''** Inserta MovColDet -  Monto Entregar
                    Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
                End If
            End If
        Else
            If pnTipoPago = gColocTipoPagoVoucher Then 'CTI4 ERS0112020
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroOpePignoVoucher, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroOpePignoVoucher, psctacod, 1, gConcITFAsumido, 1, pnMontoITF, False)
            ElseIf pnTipoPago = gColocTipoPagoCargoCta Then 'CTI4 ERS0112020
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroOpePignoCargoACta, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroOpePignoCargoACta, psctacod, 1, gConcITFAsumido, 1, pnMontoITF, False)
            Else
                '** Inserta MovCol
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
            
                '** Inserta MovColDet -  Monto Entregar
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, gConcITFAsumido, 1, pnMontoITF, False)
            End If
        End If
    End If
    '++++++++++++++++++++++++++++++++++++++
    
    'CTI4 ERS0112020 **************
    If pnTipoPago = gColocTipoPagoVoucher Then
        loRegPig.actualizarRelacionVoucherCaptacion lnMovNro, pnMovNroRVD
        If pnMovNroPend <> -1 Then loRegPig.ActualizaMovPendientesRend pnMovNroPend, pnMontoTransac + pnMontoITF
    End If
    If pnTipoPago = gColocTipoPagoCargoCta Then
        loRegPig.CapCargoCuentaAho pMatAho, psCtaCodCargo, pnMontoTransac + pnMontoITF, lsOpeCodCargoCta, psMovNro, "Cargo a Cuenta Cancel. Pigno.: " & psctacod, psFechaHora, , , , , , , , , , , , , , True, CDbl(pnMontoITF), True, gITFCobroCargo
    End If
    'END CTI4
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    'RECO20140206 ERS002*********************************
    If pbTrans = False Then
        loRegPig.dCommitTrans
    End If
    'loRegPig.dCommitTrans
    'RECO FIN********************************************

    'loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
    'RECO20140206 ERS002*********************************
        If pbTrans = False Then
            loRegPig.dRollbackTrans
        End If
        'loRegPig.dRollbackTrans
    'RECO FIN********************************************

    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCredPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nCustodiaDiferidaCredPignoraticio(ByVal psctacod As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnMontoTransac As Currency, ByVal pnCustodNeto As Currency, ByVal pnCustodImpuesto As Currency, _
        ByVal psTpoDoc As String, ByVal psNroDoc As String, _
        Optional pbEjecBatch As Boolean = False, Optional ByVal pnITF As Double = 0#)

'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'** dInsertMovDoc
'************************************

Dim lsSQL As String
Dim loRegPig  As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = gColPOpeCobCusDiferida

'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Custodia Diferida Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet -  Custodia Neto
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCustodiaDiferida, 1, pnCustodNeto, False)

    '** Inserta MovColDet -  Custodia Impuesto
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCustodiaDiferidaImpuesto, 1, pnCustodImpuesto, False)

    '** Inserta MovDoc -  Registra la boleta
    Call loRegPig.dInsertMovDoc(lnMovNro, 3, psNroDoc, psFechaHora, False)
     
     '** ITF
    Call loRegPig.dInsertMovCol(lnMovNro, "990104", psctacod, 1, pnITF, 0, "", 0, 0, 0, False)
    Call loRegPig.dInsertMovColDet(lnMovNro, "990104", psctacod, 1, gConcITFCliente, 1, pnITF, False)
     
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCredPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub



Public Sub nDuplicadoContratoCredPignoraticio(ByVal psctacod As String, ByVal pnNroDuplic As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnMontoTransac As Currency, Optional pbEjecBatch As Boolean = False, Optional ByVal pnITF As Double = 0#)

'** Actualiza Producto (NroTrans )
'** Actualiza Colocaciones (Ultima Modificacion)
'** Actualiza ColocPignoraticio (NroDuplicado)
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = gColPOpeImpDuplicado

'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , , psFechaHora, -2, False)           ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Actualiza ColocPignoraticio
    Call loRegPig.dUpdateColocPignoraticio(psctacod, , , , , , , pnNroDuplic + 1, , , , False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Duplicado Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet -  Costo Duplicado
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCostoDuplicado, 1, pnMontoTransac, False)
    
    '**ITF
    Call loRegPig.dInsertMovCol(lnMovNro, "990104", psctacod, 1, pnITF, 0, "", 0, 0, 0, False)
    Call loRegPig.dInsertMovColDet(lnMovNro, "990104", psctacod, 1, gConcITFCliente, 1, pnITF, False)
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCredPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

'--**** EXTORNOS
Public Sub nExtornoDesembolsoCredPig(ByVal psctacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False, Optional ByRef pbProcedeExtAho As Boolean = False, _
        Optional ByRef pnITF As Currency = 0, Optional ByRef psCtaAhoExt As String = "", _
        Optional ByVal psCodUser As String = "SIST", _
        Optional ByVal psMotExtorno As Variant)
'*** PEAC 20090703 - se aumento los parametros pnITF, psCtaAhoExt

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol

'*********************************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsOpeCodExtAcIn As String 'ALPA20130827
lsOpeCod = geColPExtDesemb

'**********************************************************

Dim ofun As COMNContabilidad.NCOMContFunciones
Dim lbProcedeExtAho As Boolean
Dim sMovNroCap As String
Dim RCapMov, RCap As ADODB.Recordset
Set RCapMov = RecuperaMovimientosAhorros(pnMovNroAnt)
Set ofun = New COMNContabilidad.NCOMContFunciones
lbProcedeExtAho = False

If Not (RCapMov.EOF And RCapMov.BOF) Then
    'ALPA20130827**********************************************************************
    'sMovNroCap = oFun.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
    'sMovNroCap = ofun.GeneraMovNro(CDate(Format(psFechaHora, "dd/mm/yyyy")), Mid(RCapMov("cctacod"), 4, 2), "PEAC")
    sMovNroCap = ofun.GeneraMovNro(CDate(Format(psFechaHora, "dd/mm/yyyy")), Mid(RCapMov("cctacod"), 4, 2), psCodUser)
    '**********************************************************************************
    'psFechaHora
    lbProcedeExtAho = True
End If

pbProcedeExtAho = lbProcedeExtAho

'**********************************************************

'On Error GoTo ErrorModPig
    
    loRegPig.dBeginTrans
    mbTrans = True




'**************************

psCtaAhoExt = ""

If lbProcedeExtAho Then

    Set RCap = RecuperaMovimientoCapataciones(pnMovNroAnt)
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    psCtaAhoExt = RCap!cCtaCod
    Do While Not RCap.EOF
        If RCap!nMonto > 0 Then
            If Mid(RCap!copecod, 1, 2) <> "99" Then
                'ALPA20130827*******************
                If Left(RCap!copecod, 4) = "2008" Then
                    lsOpeCodExtAcIn = gCTSExtDepActiInact
                Else
                    lsOpeCodExtAcIn = gAhoExtDepEfec
                End If
                '********************************
                Call loRegPig.CapExtornoAbonoAho(sMovNroCap, lnMovNro, pnMovNroAnt, lsOpeCodExtAcIn, RCap!cCtaCod, psMovNro, "Extorno de Desembolso Abono Cuenta", RCap!nMonto)
                'Call loRegPig.CapExtornoAbonoAho(sMovNroCap, lnMovNro, pnMovNroAnt, gAhoExtDepEfec, RCap!cCtaCod, psMovNro, "Extorno de Desembolso Abono Cuenta", RCap!nMonto)
            Else
                Call loRegPig.CapExtornoCargoAho(lnMovNro, 0, gITFCobroCargoExt, RCap!cCtaCod, sMovNroCap, "Extorno de Desembolso Abono Cuenta", RCap!nMonto)
                pnITF = RCap!nMonto
            End If
        End If
        RCap.MoveNext
    Loop
End If

'**************************
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , gColPEstRegis, psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstRegis, 1, pnMonto, "Extorno Desembolso Cont", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Extorno Desembolso Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        '**CTI3 (ferimoro) 02102018
    Call loRegPig.dInsertDetExtMov(lnMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    'WIOR 20140203********************************
    Call loRegPig.ActualizaCredVinculados(psctacod, "", 5, 2)
    'WIOR FIN ************************************
    
    'JOEP20210915 campana prendario
    Call loRegPig.CampPrendExtorno(psctacod, psFechaHora, 0, sMovNroCap)
    'JOEP20210915 campana prendario
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nExtornoDuplicadoContratoCredPig(ByVal psctacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExtorno As Variant)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = geColPExtDupli

'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , , psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Duplicado Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
   
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
        '**CTI3 (ferimoro) 02102018
    Call loRegPig.dInsertDetExtMov(lnMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nExtornoRenovacionCredPig(ByVal psctacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, ByRef psmensaje As String, _
        Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExtorno As Variant, Optional psOpeTpo As String = "", _
        Optional pnITFOpeVoucher As Currency = 0, Optional psMovNroCap As String = "", Optional ByRef psCtaAhoExt As String = "", _
        Optional ByRef psClienteExt As String = "", Optional ByRef pnMontoExt As Currency = 0)
        'CTI4 ERS0112020 Add: psOpeTpo, pnITFOpeVoucher,psCtaAhoExt,psClienteExt,psClienteExt,psOperacionDesc,pnMontoExt


'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long, lnNroRenov As Integer
Dim lsOpeCod As String

Dim loDataExt As COMDColocPig.DCOMColPFunciones
Dim lrDataExt As New ADODB.Recordset
Dim RCapMov  As ADODB.Recordset 'CTI4 ERS0112020
'On Error GoTo ErrorModPig

lsOpeCod = geColPExtRenov

' *** Obtiene Datos para el Extorno

'*** PEAC 20080708
'lsSQL = "SELECT CE.nPlazo, CE.cCtaCod, " _
      & " nSaldoCap = (SELECT nSaldo FROM Producto WHERE cCtaCod = '" & psctacod & "'  ) , " _
      & " nCapitalPag = (SELECT ISNULL(SUM(nMonto),0) FROM MovColDet WHERE nMovNro = " & pnMovNroAnt _
      & "               AND nPrdConceptoCod IN (" & gColPConceptoCodCapital & "," & gColPConceptoCodCapitalVencido & ") ) ," _
      & " nGastoCorrespond = (SELECT ISNULL(SUM(nMonto),0) FROM MovColDet WHERE nMovNro = " & pnMovNroAnt _
      & "               AND nPrdConceptoCod IN (2215)  ) ," _
      & " dVenc = (SELECT dVenc FROM Colocaciones WHERE cCtaCod = '" & psctacod & "'  ) , " _
      & " nDiasAtraso = (SELECT nDiasMora FROM MovCol where nMovNro = " & pnMovNroAnt & " and not cOpeCod like '99%'  ) , " _
      & " nPrdEstadoAnt = (SELECT nEstAnt = CASE WHEN nDiasMora <= 0 AND nNroRenov = 0 THEN " & gColPEstDesem & "  " _
      & "                                        WHEN nDiasMora > 0 THEN " & gColPEstVenci & " " _
      & "                                        WHEN nDiasMora > 30 THEN " & gColPEstPRema & " " _
      & "                                        WHEN nDiasMora <= 0 AND nNroRenov > 0 then  " & gColPEstRenov & " " _
      & "                                   END " _
      & "                  FROM MovCol MovCol INNER JOIN ColocPignoraticio ColPig " _
      & "                  ON MovCol.cCtaCod = ColPig.cCtaCod WHERE nMovNro = " & pnMovNroAnt & " and not cOpeCod like '99%' )  " _
      & " ,CP.NNRORENOV " _
      & " FROM ColocacEstado CE " _
      & " JOIN COLOCPIGNORATICIO CP  ON CP.CCTACOD=CE.CCTACOD " _
      & " WHERE CE.cCtaCod = '" & psctacod & "' AND CE.nPrdEstado = " & gColPEstRenov & "  " _
      & " AND CE.dPrdEstado = (SELECT MAX(dPrdEstado) FROM ColocacEstado " _
      & "                     WHERE cCtaCod = CE.cCtaCod AND nPrdEstado = CE.nPrdEstado ) "


'*** PEAC 20080708
lsSQL = " exec stp_sel_RecuperaDatosExtornoRenoCredPig '" & psctacod & "'," & pnMovNroAnt

Set loDataExt = New COMDColocPig.DCOMColPFunciones
    Set lrDataExt = loDataExt.dObtieneRecordSet(lsSQL)
Set loDataExt = Nothing
If lrDataExt Is Nothing Then
    psmensaje = "ERROR: al Buscar datos para Extorno "
    Exit Sub
ElseIf lrDataExt.BOF And lrDataExt.EOF Then
    psmensaje = "ERROR: No encontro datos de la Operacion a extornar "
    Exit Sub
End If


' Asigna los valores
Dim lnAntSaldoCap As Currency
Dim lnAntEstado As Integer
Dim lnAntPlazo As Integer
Dim lsAntFecVenc As String

lnAntSaldoCap = lrDataExt!nSaldoCap + lrDataExt!nCapitalPag
lnAntEstado = lrDataExt!nPrdEstadoAnt
lnAntPlazo = lrDataExt!nPlazo  ' OJO (Falta el plazo anterior )
'lsAntFecVenc = Format$(DateAdd("d", -1 * (lrDataExt!nDiasAtraso + lrDataExt!nPlazo), lrDataExt!dVenc), "mm/dd/yyyy")
'ARCV 13-07-2007

'*** PEAC 20190723
'If lrDataExt!nDiasAtraso > 0 Then
'    lsAntFecVenc = Format$(DateAdd("d", -1 * (lrDataExt!nDiasAtraso + lrDataExt!nPlazo), lrDataExt!dVenc), "mm/dd/yyyy") 'Vencidos
'Else
'    '**Comentado por DAOR 20070807
'    'lsAntFecVenc = Format$(Format(lrDataExt!dVenc, "mm/dd/yyyy") - lnAntPlazo, "mm/dd/yyyy") 'No Vencidos
'
'    '**DAOR 20070807***********************
'    'lsAntFecVenc = Format$(DateAdd("d", -1 * lrDataExt!nPlazo, lrDataExt!dVenc), "mm/dd/yyyy") 'No Vencidos
'    '*** PEAC 20080123
'    lsAntFecVenc = Format$(DateAdd("d", -1 * lrDataExt!nPlazo - (lrDataExt!nDiasAtraso), lrDataExt!dVenc), "mm/dd/yyyy")  'No Vencidos
'    '**************************************
'End If

lsAntFecVenc = Format$(lrDataExt!dFecVencAnt, "mm/dd/yyyy")
'*** FIN PEAC


'----------
lnNroRenov = lrDataExt!nNroRenov
' *** Realiza el Extorno
    
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , lnAntSaldoCap, lnAntEstado, psFechaHora, -2, False)       ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, lnAntPlazo, lsAntFecVenc, , , psMovNro, , False, IIf(lrDataExt!dFecVencAnt = "", "@", Format(lrDataExt!dFecVencAnt, "yyyymmdd")), 1)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Renovacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '**CTI3 (ferimoro) 17102018
    Call loRegPig.dInsertDetExtMov(lnMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
      
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMonto, 0, "", 0, 0, lnAntSaldoCap, False)
     
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
    
    '** Actualiza NroRenovaciones
    Call loRegPig.dUpdateColocPignoraticio(psctacod, , , , , lnNroRenov - 1)

    '*** PEAC 20080715
    Call loRegPig.dUpdateColocPigNotifi(psctacod, , 0, psMovNro, False)

    '*** Gasto de Correspondencia
    If lrDataExt!nGastoCorrespond > 0 Then
        'Actualiza ColocCalendDetPig
        lsSQL = "UPDATE ColocCalendDetPig Set nMontoPagado = nMontoPagado - " & lrDataExt!nGastoCorrespond & _
                " WHERE cCtaCod ='" & psctacod & "' And nNroCalen = 1 And nColocCalendApl = 0 " & _
                " And nCuota = 1 And nPrdConceptoCod = 2215 "

        loRegPig.dEjecutar lsSQL
    End If
    
    '*** PEAC 20161122
    Call loRegPig.dUpdateColocPigPagosParcialesExtorno(psctacod, pnMovNroAnt)
    
    'CTI4 ERS0112020****************
    If psOpeTpo = gColPOpeRenovNorVoucher Or psOpeTpo = gColPOpeRenovMorVoucher Then
        loRegPig.ActualizaMovPendientesOpe pnMovNroAnt, pnMonto + pnITFOpeVoucher
        loRegPig.actualizarRelacionExtornoVoucherCaptacion pnMovNroAnt
    ElseIf psOpeTpo = gColPOpeRenovNorCargoCta Or psOpeTpo = gColPOpeRenovMorCargoCta Then
        Set RCapMov = loRegPig.RecuperaMovimientoCapataciones(pnMovNroAnt)
        Call loRegPig.CapExtornoCargoAho(pnMovNroAnt, 0, gAhoExtCargoCtaRenovPigno, RCapMov!cCtaCod, psMovNroCap, "Extorno Cargo a Renov. Pigno.", RCapMov!nMonto)
        psCtaAhoExt = RCapMov!cCtaCod
        psClienteExt = RCapMov!cCliente
        pnMontoExt = RCapMov!nMonto
    End If
    'end CTI4
    
    'joep20210922 campana prendario
    Call loRegPig.CampPrendExtorno(psctacod, psFechaHora, 1, psMovNro)
    'joep20210922 campana prendario
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub


Public Sub nExtornoCancelacionCredPig(ByVal psctacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, ByRef psmensaje As String, _
        Optional pbEjecBatch As Boolean = False, Optional ByVal psMotExtorno As Variant, Optional psOpeTpo As String = "", _
        Optional pnITFOpeVoucher As Currency = 0, Optional psMovNroCap As String = "", Optional ByRef psCtaAhoExt As String = "", _
        Optional ByRef psClienteExt As String = "", Optional ByRef pnMontoExt As Currency = 0)
        'CTI4 ERS0112020 Add: psOpeTpo, pnITFOpeVoucher, psMovNroCap

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

Dim loDataExt As New COMDColocPig.DCOMColPFunciones
Dim lrDataExt As New ADODB.Recordset
Dim RCapMov As ADODB.Recordset 'CTI4 ERS0112020
'On Error GoTo ErrorModPig

lsOpeCod = geColPExtRenov

' *** Obtiene Datos para el Extorno

'lsSQL = "SELECT CE.nPlazo, CE.cCtaCod, " _
      & " nSaldoCap = (SELECT nSaldo FROM Producto WHERE cCtaCod = '" & psctacod & "'  ) , " _
      & " nCapitalPag = (SELECT ISNULL(SUM(nMonto),0) FROM MovColDet WHERE nMovNro = " & pnMovNroAnt _
      & "               AND nPrdConceptoCod IN (" & gColPConceptoCodCapital & "," & gColPConceptoCodCapitalVencido & ") ) ," _
      & " dVenc = (SELECT dVenc FROM Colocaciones WHERE cCtaCod = '" & psctacod & "'  ) , " _
      & " nGastoCorrespond = (SELECT ISNULL(SUM(nMonto),0) FROM MovColDet WHERE nMovNro = " & pnMovNroAnt _
      & "               AND nPrdConceptoCod IN (2215) ) ," _
      & " nDiasAtraso = (SELECT nDiasMora FROM MovCol where nMovNro = " & pnMovNroAnt & " and NOT cOpeCod like '99%'   ) , " _
      & " nPrdEstadoAnt = (SELECT nEstAnt = CASE WHEN nDiasMora <= 0 AND nNroRenov = 0 THEN " & gColPEstDesem & "  " _
      & "                                        WHEN nDiasMora > 0 THEN " & gColPEstVenci & " " _
      & "                                        WHEN nDiasMora > 30 THEN " & gColPEstPRema & " " _
      & "                                        WHEN nDiasMora <= 0 AND nNroRenov > 0 then  " & gColPEstRenov & " " _
      & "                                   END " _
      & "                  FROM MovCol MovCol INNER JOIN ColocPignoraticio ColPig " _
      & "                  ON MovCol.cCtaCod = ColPig.cCtaCod WHERE nMovNro = " & pnMovNroAnt & " and NOT cOpeCod like '99%' ) " _
      & "FROM ColocacEstado CE WHERE CE.cCtaCod = '" & psctacod & "' AND CE.nPrdEstado = " & gColPEstDifer & "  " _
      & "AND CE.dPrdEstado = (SELECT MAX(dPrdEstado) FROM ColocacEstado " _
      & "                     WHERE cCtaCod = CE.cCtaCod AND nPrdEstado = CE.nPrdEstado ) "

'*** PEAC 20080701
lsSQL = "exec stp_sel_RecuperaDatosExtornoCancCredPig '" & psctacod & "'," & pnMovNroAnt

Set loDataExt = New COMDColocPig.DCOMColPFunciones
    Set lrDataExt = loDataExt.dObtieneRecordSet(lsSQL)
Set loDataExt = Nothing
If lrDataExt Is Nothing Then
    psmensaje = "ERROR: al Buscar datos para Extorno "
    Exit Sub
End If
If lrDataExt.BOF And lrDataExt.EOF Then
    psmensaje = "ERROR: No encontro datos de la Operacion a extornar "
    Exit Sub
End If
' Asigna los valores
Dim lnAntSaldoCap As Currency
Dim lnAntEstado As Integer
Dim lnAntPlazo As Integer
Dim lsAntFecVenc As String

lnAntSaldoCap = lrDataExt!nSaldoCap + lrDataExt!nCapitalPag
lnAntEstado = lrDataExt!nPrdEstadoAnt
'lnAntPlazo = lrDataExt!nPlazo  ' OJO (Falta el plazo anterior )
'lsAntFecVenc = Format$(DateAdd("d", -1 * (lrDataExt!nDiasAtraso + lrDataExt!nPlazo), lrDataExt!dVenc), "mm/dd/yyyy")
' *** Realiza el Extorno

    
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , lnAntSaldoCap, lnAntEstado, psFechaHora, -2, False)       ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)


    '*** PEAC 20080715
    Call loRegPig.dUpdateColocPigNotifi(psctacod, , 0, psMovNro, False)


    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, lnAntEstado, 1, pnMonto, "Extorno Cancelacion Cred Pig", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Cancelacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        
    '**CTI3 (ferimoro) 17102018
    Call loRegPig.dInsertDetExtMov(lnMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
      
   
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMonto, 0, "", 0, 0, lnAntSaldoCap, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    '*** Gasto de Correspondencia
    If lrDataExt!nGastoCorrespond > 0 Then
        'Actualiza ColocCalendDetPig
        lsSQL = "UPDATE ColocCalendDetPig Set nMontoPagado = nMontoPagado - " & lrDataExt!nGastoCorrespond & _
                " WHERE cCtaCod ='" & psctacod & "' And nNroCalen = 1 And nColocCalendApl = 0 " & _
                " And nCuota = 1 And nPrdConceptoCod = 2215 "

        loRegPig.dEjecutar lsSQL
    End If

    'CTI4 ERS0112020****************
    If psOpeTpo = gColPOpeCancelNorVoucher Or psOpeTpo = gColPOpeCancelMorVoucher Then
        loRegPig.ActualizaMovPendientesOpe pnMovNroAnt, pnMonto + pnITFOpeVoucher
        loRegPig.actualizarRelacionExtornoVoucherCaptacion pnMovNroAnt
    ElseIf psOpeTpo = gColPOpeCancelNorCargoCta Or psOpeTpo = gColPOpeCancelMorCargoCta Then
        Set RCapMov = loRegPig.RecuperaMovimientoCapataciones(pnMovNroAnt)
        Call loRegPig.CapExtornoCargoAho(pnMovNroAnt, 0, gAhoExtCargoCtaCancelPigno, RCapMov!cCtaCod, psMovNroCap, "Extorno Cargo a Cuenta por Cancel. Pigno.", RCapMov!nMonto)
        psCtaAhoExt = RCapMov!cCtaCod
        psClienteExt = RCapMov!cCliente
        pnMontoExt = RCapMov!nMonto
    End If
    'end CTI4

     'JOEP20210923 campana Prendario
    Call loRegPig.CampPrendExtorno(psctacod, psFechaHora, 3, psMovNro)
    'JOEP20210923 campana Prendario

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nExtornoAmortizacionCredPig(ByVal psctacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        ByVal psOpeExt, ByRef psmensaje As String, Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExtorno As Variant)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

Dim loDataExt As COMDColocPig.DCOMColPFunciones
Dim lrDataExt As New ADODB.Recordset

'On Error GoTo ErrorModPig

lsOpeCod = gColPOpeExtRenov

' *** Obtiene Datos para el Extorno

'lsSQL = "SELECT CE.nPlazo, CE.cCtaCod, " _
      & " nSaldoCap = (SELECT nSaldo FROM Producto WHERE cCtaCod = '" & psCtaCod & "'  ) , " _
      & " nCapitalPag = (SELECT ISNULL(SUM(nMonto),0) FROM MovColDet WHERE nMovNro = " & pnMovNroAnt _
      & "               AND nPrdConceptoCod IN (" & gColPConceptoCodCapital & "," & gColPConceptoCodCapitalVencido & ") ) ," _
      & " dVenc = (SELECT dVenc FROM Colocaciones WHERE cCtaCod = '" & psCtaCod & "'  ) , " _
      & " nDiasAtraso = (SELECT nDiasMora FROM MovCol where nMovNro = " & pnMovNroAnt & " and copecod='" & psOpeExt & "'  ) , " _
      & " nPrdEstadoAnt = (SELECT nEstAnt = CASE WHEN nDiasMora <= 0 AND nNroRenov = 0 THEN " & gColPEstDesem & "  " _
      & "                                        WHEN nDiasMora > 0 THEN " & gColPEstVenci & " " _
      & "                                        WHEN nDiasMora > 30 THEN " & gColPEstPRema & " " _
      & "                                        WHEN nDiasMora <= 0 AND nNroRenov > 0 then  " & gColPEstRenov & " " _
      & "                                   END " _
      & "                  FROM MovCol MovCol INNER JOIN ColocPignoraticio ColPig " _
      & "                  ON MovCol.cCtaCod = ColPig.cCtaCod WHERE nMovNro = " & pnMovNroAnt & " and copecod='" & psOpeExt & "'   ) " _
      & "FROM ColocacEstado CE WHERE CE.cCtaCod = '" & psCtaCod & "' AND CE.nPrdEstado = " & gColPEstRenov & "  " _
      & "AND CE.dPrdEstado = (SELECT MAX(dPrdEstado) FROM ColocacEstado " _
      & "                     WHERE cCtaCod = CE.cCtaCod AND nPrdEstado = CE.nPrdEstado ) "
      
'*** PEAC 20161122
lsSQL = "exec stp_sel_ExtornoAmortizacionCredPig '" & psctacod & "'," & pnMovNroAnt & ",'" & psOpeExt & "' "

Set loDataExt = New COMDColocPig.DCOMColPFunciones
    Set lrDataExt = loDataExt.dObtieneRecordSet(lsSQL)
Set loDataExt = Nothing
If lrDataExt Is Nothing Then
    psmensaje = "ERROR: al Buscar datos para Extorno "
    Exit Sub
End If
If lrDataExt.BOF And lrDataExt.EOF Then
    psmensaje = "ERROR: No encontro datos de la Operacion a extornar "
    Exit Sub
End If
' Asigna los valores
Dim lnAntSaldoCap As Currency
Dim lnAntEstado As Integer
Dim lnAntPlazo As Integer
Dim lsAntFecVenc As String

lnAntSaldoCap = lrDataExt!nSaldoCap + lrDataExt!nCapitalPag
lnAntEstado = lrDataExt!nPrdEstadoAnt
'lnAntPlazo = lrDataExt!nPlazo  ' OJO (Falta el plazo anterior )
lsAntFecVenc = Format$(DateAdd("d", -1 * (lrDataExt!nDiasAtraso + lrDataExt!nPlazo), lrDataExt!dVenc), "mm/dd/yyyy")
' *** Realiza el Extorno

    
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , lnAntSaldoCap, lnAntEstado, psFechaHora, -2, False)       ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    '***PEAC 20080125
    'Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)
    '***PEAC 20080125 PARA QUE AL EXTORNAR VUELVA A SU FECHA CORRECTA
    Call loRegPig.dUpdateColocaciones(psctacod, , lsAntFecVenc, , , psMovNro, , False, IIf(lrDataExt!dFecVencAnt = "", "@", Format(lrDataExt!dFecVencAnt, "yyyymmdd")), 1)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Amortizacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    '**CTI3 (ferimoro) 02102018
    Call loRegPig.dInsertDetExtMov(psMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
  
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMonto, 0, "", 0, 0, lnAntSaldoCap, False)
     
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    '*** PEAC 20161122
    Call loRegPig.dUpdateColocPigPagosParcialesExtorno(psctacod, pnMovNroAnt)

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

'*** PEAC 20161212
Public Sub nExtornoPagoParcialCredPig(ByVal psctacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        ByVal psOpeExt, ByRef psmensaje As String, Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExtorno As Variant, Optional psOpeTpo As String = "", _
        Optional pnITFOpeVoucher As Currency = 0, Optional psMovNroCap As String = "", Optional ByRef psCtaAhoExt As String = "", _
        Optional ByRef psClienteExt As String = "", Optional ByRef pnMontoExt As Currency = 0)
        'CTI4 ERS0112020 Add: psOpeTpo, pnITFOpeVoucher, psMovNroCap,psCtaAhoExt,psClienteExt,pnMontoExt


'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

Dim loDataExt As COMDColocPig.DCOMColPFunciones
Dim lrDataExt As New ADODB.Recordset
Dim RCapMov As ADODB.Recordset 'CTI4 ERS0112020
'On Error GoTo ErrorModPig

lsOpeCod = gColPOpeExtPagoParc 'gColPOpeExtRenov

' *** Obtiene Datos para el Extorno

lsSQL = "exec stp_sel_ExtornoPagoParcialCredPig '" & psctacod & "'," & pnMovNroAnt & ",'" & psOpeExt & "' "

Set loDataExt = New COMDColocPig.DCOMColPFunciones
    Set lrDataExt = loDataExt.dObtieneRecordSet(lsSQL)
Set loDataExt = Nothing
If lrDataExt Is Nothing Then
    psmensaje = "ERROR: al Buscar datos para Extorno "
    Exit Sub
ElseIf lrDataExt.BOF And lrDataExt.EOF Then
    psmensaje = "ERROR: No encontro datos de la Operacion a extornar "
    Exit Sub
ElseIf lrDataExt!nMovRela > 0 Then
    psmensaje = "Antes debe extornar la operación: " & lrDataExt!cOpeRela
    Exit Sub
End If
' Asigna los valores
Dim lnAntSaldoCap As Currency
Dim lnAntEstado As Integer
Dim lnAntPlazo As Integer
Dim lsAntFecVenc As String

lnAntSaldoCap = lrDataExt!nSaldoCap + lrDataExt!nCapitalPag
lnAntEstado = lrDataExt!nPrdEstadoAnt
'lnAntPlazo = lrDataExt!nPlazo  ' OJO (Falta el plazo anterior )
lsAntFecVenc = Format$(DateAdd("d", -1 * (lrDataExt!nDiasAtraso + lrDataExt!nPlazo), lrDataExt!dVenc), "mm/dd/yyyy")
' *** Realiza el Extorno

    
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , lnAntSaldoCap, lnAntEstado, psFechaHora, -2, False)       ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    '***PEAC 20080125
    'Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)
    '***PEAC 20080125 PARA QUE AL EXTORNAR VUELVA A SU FECHA CORRECTA
    '*** PEAC 20161213 - NO MODIFICA FECHA VENCIMIENTO
    'Call loRegPig.dUpdateColocaciones(psCtaCod, , lsAntFecVenc, , , psMovNro, , False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Pago Parcial Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '**CTI3 (ferimoro) 02102018
    Call loRegPig.dInsertDetExtMov(lnMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
          
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMonto, 0, "", 0, 0, lnAntSaldoCap, False)
     
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    '*** PEAC 20161122
    'Call loRegPig.dUpdateColocPigPagosParcialesExtorno(psCtaCod, pnMovNroAnt)
    
    '*** GEMO 20210810
    Call loRegPig.dDeleteColocPigPagosParcialesExtorno(psctacod, pnMovNroAnt, lnMovNro)
    '*** GEMO 20210810
    
    'CTI4 ERS0112020****************
    If psOpeTpo = gColPOpePagoParcialNorVoucher Then
        loRegPig.ActualizaMovPendientesOpe pnMovNroAnt, pnMonto + pnITFOpeVoucher
        loRegPig.actualizarRelacionExtornoVoucherCaptacion pnMovNroAnt
    ElseIf psOpeTpo = gColPOpePagoParcialNorCargoCta Then
        Set RCapMov = loRegPig.RecuperaMovimientoCapataciones(pnMovNroAnt)
        Call loRegPig.CapExtornoCargoAho(pnMovNroAnt, 0, gAhoExtCargoCtaPagoParcialPigno, RCapMov!cCtaCod, psMovNroCap, "Extorno Cargo a Cuenta por Amort. Pigno.", RCapMov!nMonto)
        psCtaAhoExt = RCapMov!cCtaCod
        psClienteExt = RCapMov!cCliente
        pnMontoExt = RCapMov!nMonto
    End If
    'end CTI4

    'JOEP20210922 campna prendario
    Call loRegPig.CampPrendExtorno(psctacod, psFechaHora, 2, psMovNro)
    'JOEP20210922 campna prendario
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nExtornoRescateJoyaCredPig(ByVal psctacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExtorno As Variant)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = geColPExtDevJoyas

'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , gColPEstDifer, psFechaHora, -2, False)       ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstDifer, 1, 0, "Extorno Entrega Joya", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Rescate Joyas Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '**CTI3 (ferimoro) 02102018
    Call loRegPig.dInsertDetExtMov(lnMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
          
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nExtornoCustodiaDiferidaCredPig(ByVal psctacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExtorno As Variant)

'** Actualiza Producto ( NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = geColPExtCustodDifer

'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , , psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Custodia Diferida Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '**CTI3 (ferimoro) 02102018
    Call loRegPig.dInsertDetExtMov(lnMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
   
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nExtornoRemateVentaCredPig(ByVal psctacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExtorno As Variant)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** Actualiza ColocPigRGDet (Estado, Monto Venta, nMovVta)
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim loNroPro As New COMDColocPig.DCOMColPFunciones
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsNroProceso As String

Dim Rs As New ADODB.Recordset

lsOpeCod = "129701"
'Obtiene el nro de proceso
'Dim loDat As NCOMColPRecGar
'Set loDat = New NCOMColPRecGar
'    lsNroProceso = loDat.nObtieneRemateIniciado()
'Set loDat = Nothing

lsSQL = "SELECT RG.cTpoProceso, RG.cNroProceso, RG.nRGEstado " _
        & " FROM ColocPigRecGar RG  WHERE cTpoProceso ='R' and Right(cCodAge,2)= '" & Mid(psctacod, 4, 2) & "'" _
        & " and  nRGEstado ='" & gColPRecGarEstIniciado & "'" _
        & " ORDER BY cNroProceso DESC "
Set Rs = loNroPro.dObtieneRecordSet(lsSQL)
If Not (Rs.EOF And Rs.BOF) Then
    lsNroProceso = Rs!cNroProceso
End If


'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , gColPEstPRema, psFechaHora, -2, False)        ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("R", lsNroProceso, psctacod, gColPRecGarVentaNoVendido, 0, , , , 0, False)
    
    '** Elimina ColocPigRecupVta
    Call loRegPig.dDeleteColocPigRecupVta("R", lsNroProceso, psctacod, False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Venta Remate ", gMovEstContabMovContable, gMovFlagVigente, False)
    
    '**CTI3 (ferimoro) 02102018
    Call loRegPig.dInsertDetExtMov(psMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nExtornoSubastaVentaCredPig(ByVal psctacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExtorno As Variant)
Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsNroProceso As String

lsOpeCod = "129703"
'Obtiene el nro de proceso
Dim loDat As NCOMColPRecGar
Set loDat = New NCOMColPRecGar
    lsNroProceso = loDat.nObtieneRemateIniciado()
Set loDat = Nothing

'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , gColPEstAdjud, psFechaHora, -2, False)        ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("A", lsNroProceso, psctacod, gColPRecGarAdjudicado, 0, , , , 0, False)
    
    '** Elimina ColocPigRecupVta
    'Call loRegPig.dDeleteColocPigRecupVta("R", lsNroProceso, psCtacod, False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Venta Subasta ", gMovEstContabMovContable, gMovFlagVigente, False)
    
    '**CTI3 (ferimoro) 02102018
    Call loRegPig.dInsertDetExtMov(psMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
      
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "
        
End Sub


Public Sub nExtornoRecuperacionCredPig(ByVal psctacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExtorno As Variant, Optional psOpeTpo As String = "", _
        Optional pnITFOpeVoucher As Currency = 0, Optional psMovNroCap As String = "", Optional ByRef psCtaAhoExt As String = "", _
        Optional ByRef psClienteExt As String = "", Optional ByRef pnMontoExt As Currency = 0)
        'CTI4 ERS0112020 Add: psOpeTpo, psMovNroCap
Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsNroProceso As String
Dim RCapMov As ADODB.Recordset 'CTI4 ERS0112020

lsOpeCod = "129704"
'Obtiene el nro de proceso
Dim loDat As NCOMColPRecGar
Set loDat = New NCOMColPRecGar
    lsNroProceso = loDat.nObtieneRemateIniciado()
Set loDat = Nothing

'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , gColPEstAdjud, psFechaHora, -2, False)        ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Update ColocPigRGDet
    'Call loRegPig.dUpdateColocPigRGDet("R", lsNroProceso, psCtacod, gColPRecGarVentaNoVendido, 0, , , , 0, False)
    
    '** Elimina ColocPigRecupVta
   ' Call loRegPig.dDeleteColocPigRecupVta("R", lsNroProceso, psCtacod, False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Recup. Subasta ", gMovEstContabMovContable, gMovFlagVigente, False)
        
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '**CTI3 (ferimoro) 02102018
    Call loRegPig.dInsertDetExtMov(lnMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
    
       '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
    
    'CTI4 ERS0112020****************
    If psOpeTpo = gColPOpeRecuperaContratoAdjVoucher Then
        loRegPig.ActualizaMovPendientesOpe pnMovNroAnt, pnMonto + pnITFOpeVoucher
        loRegPig.actualizarRelacionExtornoVoucherCaptacion pnMovNroAnt
    ElseIf psOpeTpo = gColPOpeRecuperaContratoAdjCargoCta Then
        Set RCapMov = loRegPig.RecuperaMovimientoCapataciones(pnMovNroAnt)
        Call loRegPig.CapExtornoCargoAho(pnMovNroAnt, 0, gAhoExtCargoCtaRecuperaContratoAdjPigno, RCapMov!cCtaCod, psMovNroCap, "Extorno Cargo a Cuenta por Recup. Contrato. Adj", RCapMov!nMonto)
        psCtaAhoExt = RCapMov!cCtaCod
        psClienteExt = RCapMov!cCliente
        pnMontoExt = RCapMov!nMonto
    End If
    'end CTI4

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "
        
End Sub


Public Sub nRemateVentaCredPignoraticio(ByVal psctacod As String, ByVal psNroProceso As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnMontoTransac As Currency, ByVal pnCapital As Currency, ByVal pnIntVenc As Currency, _
        ByVal pnCustodiaVenc As Currency, ByVal pnPreparaRemate As Currency, ByVal pnComisionRemate As Currency, _
        ByVal pnImpuesto As Currency, ByVal pnSobranteRemate As Currency, ByVal pnValTasacion As Currency, _
        ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, ByVal pnKi21 As Double, _
        ByVal psPersCodComp As String, ByVal psNroDoc As String, ByVal psNroBoleta As String, _
        ByVal pnIGVBoleta As Currency, ByVal pnSubtotalBoleta As Currency, Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertColocacEstado
'** Actualiza ColocPigRGDet (Estado, Monto Venta, nMovVta)
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = gColPOpeVtaRemate

'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , gColPEstRemat, psFechaHora, -2, False)           ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstRenov, 1, pnMontoTransac, "Venta Joyas en Remate", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Venta Remate", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("R", psNroProceso, psctacod, gColPRecGarVentaVendido, pnMontoTransac, , , , lnMovNro, False)
    
    '** Inserta ColocPigRecupVta
    Call loRegPig.dInsertColocPigRecupVta("R", psNroProceso, psctacod, 23, psNroDoc, psPersCodComp, pnMontoTransac, pnComisionRemate, pnSobranteRemate, False)

    '** Inserta MovCol Monto Venta
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, (pnMontoTransac - pnIGVBoleta), 0, "", 0, 0, 0, False)
    
    '--------------------- AVMM 21-09-2006 -----------------------------------
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivo, psctacod, 1, pnIGVBoleta, 0, "", 0, 0, 0, False)
    '** Inserta MovColDet -  Monto Entregar
    Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivo, psctacod, 1, gConcITFCliente, 1, pnIGVBoleta, False)
    '--------------------------------------------------------------------------

    '** Inserta MovColDet -  Capital
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCapital, 1, pnCapital, False)

    '** Inserta MovColDet -  Interes Vencido
    If pnIntVenc > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresMoratorio, 1, pnIntVenc, False)
    End If
    '** Inserta MovColDet -  Custod Venc
    If pnCustodiaVenc > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCustodiaVencida, 1, pnCustodiaVenc, False)
    End If
    '** Inserta MovColDet -  Preparac Remate
    If pnPreparaRemate > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodPreparaRemate, 1, pnPreparaRemate, False)
    End If
    '** Inserta MovColDet -  Costo Remate - Comision
    If pnComisionRemate > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodComisionRemate, 1, pnComisionRemate, False)
    End If
    '** Inserta MovColDet -  Impuesto
    If pnImpuesto > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodImpuesto, 1, pnImpuesto, False)
    End If
    '** Inserta MovColDet -  Sobrante Remate
    If pnSobranteRemate > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodSobranteRemate, 1, pnSobranteRemate, False)
    End If
    
    '** Inserta MovColDet ' Valor Tasacion
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodTasacion, 1, pnValTasacion, False)
    

    '** Inserta MovColDet -  IGV de Boleta
'    If pnIGVBoleta > 0 Then
'        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, 2216, 1, pnIGVBoleta, False)
'    End If

    '** Inserta MovColDet -  Subtotal
    If pnSubtotalBoleta > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, 2217, 1, pnSubtotalBoleta, False)
    End If


    '** Inserta MovColDet ' Oro 14/16/18/21
    If pnki14 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro14, 1, pnki14, False)
    End If
    If pnki16 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro16, 1, pnki16, False)
    End If
    If pnKi18 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro18, 1, pnKi18, False)
    End If
    If pnKi21 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro21, 1, pnKi21, False)
    End If
    'Inserta la Poliza
    Call loRegPig.dInsertMovDoc(lnMovNro, 23, psNroDoc, psFechaHora, False)
    
    'Inserta la Boleta de Venta
    Call loRegPig.dInsertMovDoc(lnMovNro, 3, psNroBoleta, psFechaHora, False)
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nSubastaVentaCredPignoraticio(ByVal psctacod As String, ByVal psNroProceso As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnMontoTransac As Currency, ByVal pnVtaNetaSubasta As Currency, _
        ByVal pnImpuestoVtaSubasta As Currency, ByVal pnValorAdjudicacion As Currency, _
        ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, ByVal pnKi21 As Double, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertColocacEstado
'** dInsertMov
'** Actualiza ColocPigRGDet (Estado, Monto Venta, nMovVta)
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = geColPVtaSubasta

'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , gColPEstSubas, psFechaHora, -2, False)            ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstSubas, 1, pnMontoTransac, "Venta Joyas en Subasta", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Venta Joyas en Subasta", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("S", psNroProceso, psctacod, gColPRecGarVentaVendido, pnMontoTransac, , , , lnMovNro, False)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet -  Venta Neta
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, geColPConceptoCodVtaNetaSubasta, 1, pnVtaNetaSubasta, False)

    '** Inserta MovColDet -  Impuesto Vta Subasta
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, geColPConceptoCodImpuestoVtaSubasta, 1, pnImpuestoVtaSubasta, False)
    
    '** Inserta MovColDet -  Valor Adjudicacion
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, geColPConceptoCodValorAdjudica, 1, pnValorAdjudicacion, False)

    '** Inserta MovColDet ' Oro 14/16/18/21
    If pnki14 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro14, 1, pnki14, False)
    End If
    If pnki16 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro16, 1, pnki16, False)
    End If
    If pnKi18 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro18, 1, pnKi18, False)
    End If
    If pnKi21 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro21, 1, pnKi21, False)
    End If
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

'FOTOCOMENTADO POR FALTA DE DLL ******************************************
'*************************************************************************

Public Sub nOpeCMACLlamadaCredPignoraticio(ByVal pnOpeCod As Long, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal psGlosa As String, ByVal psPersCodCMAC As String, _
        ByVal psIFTipo As String, ByVal pnMoneda As Moneda, ByVal psCtaIf As String, _
        ByVal psDocumento As String, ByVal pnMontoTransac As Currency, _
        Optional pbEjecBatch As Boolean = False, Optional nComision As Double = 0, _
        Optional nMonedaComision As Moneda = gMonedaNacional, Optional sCuentaAho As String = "", _
        Optional sNomAge As String = "", Optional sLpt As String = "", Optional sCliente As String = "", Optional psPersLavDinero As String = "", Optional ByRef psmensaje As String, _
        Optional ByRef psBoleta As String, Optional ByRef psBoletaITF As String, Optional ByVal pbImpTMU As Boolean = False, _
        Optional psTitPersLavDinero As String = "", Optional psOrdPersLavDinero As String = "", Optional psReaPersLavDinero As String = "", _
        Optional psBenPersLavDinero As String = "", Optional psVisPersLavDinero As String = "", _
        Optional pnMovNro As Long = 0)

        'By Capi 18022008 se agrego 5 parametros opcionales al final
    '** dInsertMov
    '** dInsertMovCMAC
    '************************************
    
    Dim lsSQL As String
    Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
    Dim lnMovNro As Long
    Dim lnMovNroRef As Long
    Dim lsOpeCod As String, sCodUser As String, sCodage As String
    Dim oCont As New COMNContabilidad.NCOMContFunciones
    Dim oCon As New COMConecta.DCOMConecta
    Dim dFecSis As Date
    
    oCon.AbreConexion
    
    dFecSis = CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4))
    sCodUser = Right(psMovNro, 4)
    sCodage = Mid(psMovNro, 18, 2)
    
    'Codigo de Operacion
    lsOpeCod = pnOpeCod
    
    'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, psGlosa, gMovEstContabMovContable, gMovFlagVigente, False)

    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    'ALPA 20081010****************************
    pnMovNro = lnMovNro
    '** Inserta MovCMAC
    Call loRegPig.dInsertMovCMAC(lnMovNro, psPersCodCMAC, psIFTipo, pnMoneda, psCtaIf, _
            psDocumento, pnOpeCod, pnMontoTransac, False, sCliente)

    ' xxx loRegPig.dCommitTrans
    ' xxx mbTrans = False
   ' Set loRegPig = Nothing


    If nComision > 0 Then
        Dim sMovNroCom As String
        Dim nMovNroCom As Long
        Set oCont = New COMNContabilidad.NCOMContFunciones

        sMovNroCom = oCont.IncrementaMovNro(psMovNro)

        'xxx sMovNroCom = oCont.GeneraMovNro(dFecSis, sCodAge, sCodUser)

        Set oCont = Nothing
        'loRegPig.dBeginTrans
        Call loRegPig.dInsertMov(sMovNroCom, gOtrOpeComisionCMACLlam, psGlosa, gMovEstContabMovContable, gMovFlagVigente, False)
        nMovNroCom = loRegPig.dGetnMovNro(sMovNroCom)
        Call loRegPig.dInsertMovOpeVarias(nMovNroCom, nComision, psCtaIf, psPersCodCMAC, nMonedaComision, False)
        Call loRegPig.dInsertMovRef(lnMovNro, nMovNroCom)
        'loRegPig.dCommitTrans
    End If

    If sCuentaAho <> "" Then
        Dim oCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
        Dim sMovNroReg As String
        Set oCont = New COMNContabilidad.NCOMContFunciones
        'sMovNroReg = oCont.GeneraMovNro(dFecSis, sCodAge, sCodUser)

        'xxx agregado
        If Len(Trim(sMovNroCom)) > 0 Then
            sMovNroReg = oCont.IncrementaMovNro(sMovNroCom)
        Else
            sMovNroReg = oCont.IncrementaMovNro(psMovNro)
        End If

        Set oCont = Nothing

        Set oCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento

        oCapMov.CapAbonoCuentaAho sCuentaAho, pnMontoTransac, gAhoDepRegCMACLlam, sMovNroReg, psGlosa, , , , , , , , , sNomAge, sLpt, psPersLavDinero, True, , , , , oCon.ConexionActiva, , , , , , , , psmensaje, psBoleta, psBoletaITF, pbImpTMU

        Set oCapMov = Nothing

        'Pepe
        lnMovNroRef = loRegPig.dGetnMovNro(sMovNroReg)
        Call loRegPig.dInsertMovRef(lnMovNro, lnMovNroRef)
        'Fin Pepe

    End If

    'Lavado de Dinero

    'Dim oBase As DCredActualizaBD
    'Set oBase = New DCredActualizaBD

     'Lavado de Dinero
''    If psPersLavDinero <> "" Then
''        'ALPA 20081009********************************************************
''        'loRegPig.dInsertaMovLavDinero lnMovNro, psPersLavDinero, psTitPersLavDinero, psOrdPersLavDinero, psReaPersLavDinero, psBenPersLavDinero, psVisPersLavDinero
''        loRegPig.dInsertaMovLavDinero lnMovNro, psPersLavDinero, psTitPersLavDinero, psOrdPersLavDinero, psReaPersLavDinero, psBenPersLavDinero, psVisPersLavDinero, nTipoREU, nMontoAcu, sOrigen
''        '*********************************************************************
''    End If

    loRegPig.dCommitTrans

    Set loRegPig = Nothing

    'Set oBase = Nothing
    oCon.CierraConexion
Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCredPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nExtornoOpeCMACLlamadaCredPig(ByVal pnOpeCod As Long, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsOpeDesc As String
Select Case pnOpeCod
    Case geColPRenDEOtCj
        lsOpeCod = geColPExtRenovDeOtCj
        lsOpeDesc = "Renovacion "
    Case geColPCanDEOtCj
        lsOpeCod = geColPExtCancelDeOtCj
        lsOpeDesc = "Cancelacion "
End Select
'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Extorno de " & lsOpeDesc & " DE Otra CMAC ", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub


' Registra Contrato Pignoraticio con Detalle
Public Function nRegistraContratoPignoraticioDetalle(ByVal psAgencia As String, ByVal pnMoneda As Moneda, _
                ByVal prPers As ADODB.Recordset, ByVal pnTasa As Double, ByVal pnSaldo As Currency, _
                ByVal psFechaHora As String, ByVal pnPlazo As Integer, ByVal psFecVenc As String, _
                ByVal pnOroBruto As Currency, ByVal pnOroNeto As Currency, ByVal pnValTasacion As Currency, _
                ByVal pnPiezas As Integer, ByVal psTipoContrato As String, ByVal psLote As String, _
                ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, ByVal pnKi21 As Double, _
                ByVal psMovNro As String, ByVal pnIntAdelantado As Currency, ByVal pnCostoTasac As Currency, _
                ByVal pnCostoCustodia As Currency, ByVal pnImpuesto As Currency, ByVal prjoyas As ADODB.Recordset, _
                Optional ByRef pnMovNro As Long, Optional pbMalCalificacion As Boolean = False, _
                Optional ByVal pnTasaIntVencidoCrePigCliMalCal As Double = 0#, Optional pbTrabajadoVinc As Boolean = False, _
                Optional ByVal psCtaCodAmp = "", Optional ByVal pnSaldoCredAmp As Double = 0#, _
                Optional ByVal psUser As String = "", _
                Optional ByVal psCtaCampAdj As String = "", Optional ByVal pnMontoinicial As Double, Optional psLineaCredito As String = "", Optional ByVal pnTasaGracia As Double = 0#, Optional ByVal pnTasaMora As Double = 0#, Optional ByVal cUser As String = "SIST", Optional ByVal Holograma As Long = 0#) As String 'JUCS Agrego cUser and Holograma Optional
                'RECO20140130 ERS002 Se agrego el parametro psCtaCodAmp,pnSaldoCredAmp ) As String
                'WIOR 20140203 AGREGO pbTrabajadoVinc
                '*** PEAC 20080813 - se agregó el parametro pnMovNro
'** Genera Nuevo Nro Contrato
'** dInsertCreditoPignoraticio
'** dInsertColocEstado
'** dInsertCreditoPigJoya
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************
               '***Agregado por ELRO los parametros "pbMalCalificacion" y "pnTasaIntVencidoCrePigCliMalCal" el 20120120, según Acta N° 002-2012/TI-D
Dim loGen As COMDConstSistema.DCOMGeneral
Dim lsCuenta As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lnColPigTpoProducto As Integer

On Error GoTo ErrorRegPig 'ARLO20190302 ERS042-2018

lsOpeCod = gColPOpeRegContrato
lnColPigTpoProducto = 709 'ARLO20190302 ERS042-2018

'Genera Nro Cuenta
Set loGen = New COMDConstSistema.DCOMGeneral
    'lsCuenta = loGen.GeneraNuevaCuenta(psAgencia, gColConsuPrendario, pnMoneda)
    'MAVM 20100608 BAS II
    
    lsCuenta = loGen.GeneraNuevaCuenta(psAgencia, lnColPigTpoProducto, pnMoneda) 'ARLO20190302 ERS042-2018
Set loGen = Nothing
    
    'CROB20180524 begin
    Dim oDNiv As COMDColocPig.DCOMColPContrato
    Set oDNiv = New COMDColocPig.DCOMColPContrato
    
    Call oDNiv.RegistrarSolicitudNivAprobPigno(lsCuenta, 1, pnTasa, prPers("cPersCod"), pnValTasacion, pnSaldo, pnPiezas, pnOroBruto, pnOroNeto, psMovNro, psFechaHora)
    'CROB20180524 end
    
    loRegPig.dBeginTrans
    mbTrans = True

    Call loRegPig.dInsertCredPignoraticioDetalle(lsCuenta, pnTasa, pnSaldo, psFechaHora, pnPlazo, _
        psFecVenc, pnOroBruto, pnOroNeto, pnValTasacion, pnPiezas, psTipoContrato, psLote, _
        pnki14, pnki16, pnKi18, pnKi21, psMovNro, prjoyas, False, Right(psAgencia, 2), _
        lnColPigTpoProducto, pbMalCalificacion, pnTasaIntVencidoCrePigCliMalCal, psLineaCredito, pnTasaGracia, pnTasaMora, Holograma)
        '***Agregado por ELRO los parametros "pbMalCalificacion" y "pnTasaIntVencidoCrePigCliMalCal"  el 20120120, según Acta N° 002-2012/TI-D
         '***Modificación TI ERS 063-2017 JUCS AGREGÓ Hologramas
    Call loRegPig.dInsertColocacEstado(lsCuenta, psFechaHora, gColPEstRegis, 1, pnSaldo, _
        "Registro Contrato", gColocCalendCodPFCF, 0, 0, pnPlazo, "0", 0, 0, 0, 0, False)

    'inserta personas para cada registro de prPersonas
    Do While Not prPers.EOF
        Call loRegPig.dInsertProductoPersona(lsCuenta, prPers("cPersCod"), gColRelPersTitular, False)
        prPers.MoveNext
    Loop
    prPers.MoveFirst
    
     'Inserta Tasador Modificación JUCS TI-ERS063-2017
    If cUser <> "SIST" Then
        Dim cCodUserTasador As ADODB.Recordset
        Dim gColRelPersTasador As String
        Dim sSql As String
        Dim oConecta As COMConecta.DCOMConecta
        'Obtiene Codigo del Tasador con su cUser
         sSql = "exec stp_sel_ObtieneCodPersona '" & cUser & "'"
         Set oConecta = New COMConecta.DCOMConecta
         oConecta.AbreConexion
         Set cCodUserTasador = oConecta.CargaRecordSet(sSql)
         oConecta.CierraConexion
         Set oConecta = Nothing
         'Constante: Cod de Tasador Nro 39
         gColRelPersTasador = "39"
        Call loRegPig.dInsertProductoPersona(lsCuenta, cCodUserTasador("cPersCod"), gColRelPersTasador, False)
   End If
    '************************************************************
    
    'inserta ColocCalendario
    Call loRegPig.dInsertColocCalendario(lsCuenta, 1, gColocCalendAplDesembolso, 1, psFecVenc, 0, "Reg Contrato", "", False)
   '**********
    'inserta ColocCalendDetPig - Capital
    Call loRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplDesembolso, 1, gColPConceptoCodCapital, pnSaldo, 0, "", False)
    
    'inserta ColocCalendDetPig - Interes Adelantado
    Call loRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplDesembolso, 1, gColPConceptoCodInteresCompensatorio, pnIntAdelantado, 0, "", False)
    
    'inserta ColocCalendDetPig - Costo Tasacion
    If pnCostoTasac > 0 Then
        Call loRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplDesembolso, 1, gColPConceptoCodTasacion, pnCostoTasac, 0, "", False)
    End If
    'inserta ColocCalendDetPig - Costo Custodia
    If pnCostoCustodia > 0 Then
        Call loRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplDesembolso, 1, gColPConceptoCodCustodia, pnCostoCustodia, 0, "", False)
    End If
    'inserta ColocCalendDetPig - Costo Impuesto
    If pnImpuesto > 0 Then
        Call loRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplDesembolso, 1, gColPConceptoCodImpuesto, pnImpuesto, 0, "", False)
    End If
   '**********
    'inserta ColocCalendDetPig - Costo Cobro Correspondencia
    Call loRegPig.dInsertColocCalendDetPig(lsCuenta, 1, gColocCalendAplDesembolso, 1, 2215, 0, 0, "", False)
   
    ' inserta Mov '
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Registro Contrato", gMovEstContabMovContable, gMovFlagVigente, False)
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    ' inserta MovCol '
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, lsCuenta, 1, pnSaldo, 0, "GMIC", pnPlazo, 0, 0, False)
    ' Monto Prestamo(Capital)
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lsCuenta, 1, gColPConceptoCodCapital, 1, pnSaldo, False)
    ' Valor Tasacion
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lsCuenta, 1, gColPConceptoCodTasacion, 1, pnValTasacion, False)

    ' Gramos de Oro
    If pnki14 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lsCuenta, 1, gColPConceptoCodOro14, 1, pnki14, False)
    End If
    If pnki16 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lsCuenta, 1, gColPConceptoCodOro16, 1, pnki16, False)
    End If
    If pnKi18 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lsCuenta, 1, gColPConceptoCodOro18, 1, pnKi18, False)
    End If
    If pnKi21 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lsCuenta, 1, gColPConceptoCodOro21, 1, pnKi21, False)
    End If
    
    'RECO20140130 ERS002 RECO-C*******************************
    If psCtaCodAmp <> "" Then
        Call nInsertaRegistroCreditoPignoAmpliado(lsCuenta, psCtaCodAmp, gColPEstRegis, psFechaHora, pnSaldo, pnSaldoCredAmp, psUser)
    End If
    'RECO FIN******************************************


    pnMovNro = lnMovNro
    'WIOR 20140203 **************************
    Call loRegPig.dDeleteCredVinculados(lsCuenta, 1)
    If pbTrabajadoVinc Then
        Call loRegPig.dInsertCredVinculados(lsCuenta, pnSaldo, psMovNro, 2, 1)
    End If
    'WIOR FIN *******************************
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    'RECO20141204 ERS163-2014****************************
    If psCtaCampAdj <> "" Then
        Call nInsertaRegistroCreditoCampAdj(lnMovNro, psCtaCampAdj, lsCuenta, pnMontoinicial)
    End If
    'RECO FIN********************************************
    nRegistraContratoPignoraticioDetalle = lsCuenta
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Function

ErrorRegPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "
End Function

Public Sub nAmortizacionCredPignoraticio(ByVal psctacod As String, ByVal pnNewSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal psFecVenci As String, ByVal pnNewPlazo As Integer, _
        ByVal pnMontoTransac As Currency, ByVal pnCapPag As Currency, ByVal pnIntVenc As Currency, _
        ByVal pnCustodiaVenc As Currency, ByVal pnPreparaRemate As Currency, ByVal pnIntCompensat As Currency, _
        ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, ByVal pnDiasAtraso As Integer, _
        ByVal pnDiasCambCart As Integer, ByVal pnValorTasacion As Currency, _
        ByVal pnOpeCod As Long, ByVal psOpeDesc As String, ByVal psPersCodCMAC As String, _
        ByVal pbITFAplica As Boolean, ByVal pbITFAsumidoCreditos As Boolean, _
        ByVal pnMontoITF As Currency, Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnMontoCostoNotificacion As Currency = 0, _
        Optional ByRef pnMovNro As Long = 0, Optional ByVal psFechaVencAnt As String = "@")
        'WIOR 20130301 AGREGO pnMovNro
'** Actualiza Producto (Estado, nSaldoCap,dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Plazo, Vencimiento, Ultima Modificacion, dVigencia)
'** Actualiza ColocPignoraticio (NroRenov)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
Select Case pnOpeCod
    Case gColPOpeAmortizEFE  ' En la CMAC
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeAmortNorEFE
        'Else
        '    lsOpeCod = gColPOpeAmortMorEFE
        End If
    Case gColPOpeAmortizCHQ
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeAmortNorCHQ
        'Else
        '    lsOpeCod = gColPOpeamortMorCHQ
        End If
    Case gColPOpeAmortizEnOtCjEFE  ' En otra CMAC
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeAmortNorEnOtCjEFE
        'Else
        '    lsOpeCod = gColPOpeAmortMorEnOtCjEFE
        End If
    Case gColPOpeAmortizEnOtCjCHQ
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeAmortNorEnOtCjCHQ
        'Else
        '    lsOpeCod = gColPOpeAmortMorEnOtCjCHQ
        End If
End Select

'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , pnNewSaldoCap, gColPEstRenov, psFechaHora, -2, False)           ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
        
    '*** PEAC 20071206 - al momento de amortizar se grava la nueva fecha de venc a 30 dias *****
    'Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)
     
     '*** PEAC 20161213 - se agrego el campo "psFechaVencAnt"
     Call loRegPig.dUpdateColocaciones(psctacod, , psFecVenci, , , psMovNro, , False, psFechaVencAnt)
    '********************************************************************************************


    '** Actualiza ColocPignoraticio ' activado por peac 20071106
    Call loRegPig.dUpdateColocPignoraticio(psctacod, , , , , , , , , , , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstRenov, 1, pnMontoTransac, "Amortizacion Contrato", _
        gColocCalendCodPFCF, 0, 0, pnNewPlazo, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Amortizacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    pnMovNro = lnMovNro 'WIOR 20130301
    '** Inserta MovCol
    '***PEAC 20080125
    'Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, 0, "GMIC", pnNewPlazo, 0, pnNewSaldoCap, False)
    '***PEAC 20080125 SE AGREGO LOS DIAS DE ATRASO PARA PODER ESTORNAR A LA FECHA CORRECTA
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, pnDiasAtraso, "GMIC", pnNewPlazo, 0, pnNewSaldoCap, False)

    '** Inserta MovColDet -  Capital
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCapital, 1, pnCapPag, False)
    
    '** Inserta MovColDet -  Interes Compensatorio (al vencimiento)
    If pnIntCompensat > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresCompensatorio, 1, pnIntCompensat, False)
    End If
    
    If pbITFAplica Then
        If Not pbITFAsumidoCreditos Then
            If Left(gColPOpeAmortizEnOtCjEFE, 4) = Left(lsOpeCod, 4) Then
                '** Inserta MovCol
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroCMACCred, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                '** Inserta MovColDet -  Monto Entregar
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroCMACCred, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
            Else
                '** Inserta MovCol
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                '** Inserta MovColDet -  Monto Entregar
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
            End If
        Else
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
        
            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, gConcITFAsumido, 1, pnMontoITF, False)
        End If
    End If

'    '** Inserta MovColDet -  Interes Vencido
'    If pnIntVenc > 0 Then
'        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodInteresMoratorio, 1, pnIntVenc, False)
'    End If
'    '** Inserta MovColDet -  Custod Venc
'    If pnCustodiaVenc > 0 Then
'        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodCustodiaVencida, 1, pnCustodiaVenc, False)
'    End If
'    '** Inserta MovColDet -  Preparac Remate
'    If pnPreparaRemate > 0 Then
'        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodPreparaRemate, 1, pnPreparaRemate, False)
'    End If
'    '** Inserta MovColDet -  Interes Compensatorio (Adelantado)
'    If pnIntCompensat > 0 Then
'        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodInteresCompensatorio, 1, pnIntCompensat, False)
'    End If
'    '** Inserta MovColDet -  Impuesto
'    If pnImpuesto > 0 Then
'        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodImpuesto, 1, pnImpuesto, False)
'    End If
'    '** Inserta MovColDet -  Costo Custodia
'    If pnCostoCustodia > 0 Then
'        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodCustodia, 1, pnCostoCustodia, False)
'    End If
    
    '**** Cambio de Cartera - Regresa a Normal
'    If pnDiasAtraso > pnDiasCambCart Then
'        'Cambio de Cartera - Regresa a Normal
'        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodCambioCarteraMorosoNormal, 1, pnNewSaldoCap, False)
'        'Cambio a Valor de Tasación
'        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodCambioTasacionMorosoNormal, 1, pnValorTasacion, False)
'    End If
       
    If pnMontoCostoNotificacion > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, "2103", 1, pnMontoCostoNotificacion, False)
    End If
    
    '** Inserta MovCMAC - Para Recepcion de Llamada de cmac
    If pnOpeCod = gColPOpeRenovacEnOtCjEFE Or pnOpeCod = gColPOpeRenovacEnOtCjCHQ Then
        Call loRegPig.dInsertMovCMAC(lnMovNro, psPersCodCMAC, Format$(gTpoIFCmac, "00"), CInt(Mid(psctacod, 9, 1)), "", _
                "", lsOpeCod, pnMontoTransac, False)
    End If
    
    '*** PEAC 20161116
    Call loRegPig.dUpdateColocPigPagosParciales(psctacod, lnMovNro, False)
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

'*** PEAC 20161028
Public Sub nPagoParcialCredPignoraticio(ByVal psctacod As String, ByVal pnNewSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal psFecVenci As String, ByVal pnNewPlazo As Integer, _
        ByVal pnMontoTransac As Currency, ByVal pnCapPag As Currency, ByVal pnIntVenc As Currency, _
        ByVal pnCustodiaVenc As Currency, ByVal pnPreparaRemate As Currency, ByVal pnIntCompensat As Currency, _
        ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, ByVal pnDiasAtraso As Integer, _
        ByVal pnDiasCambCart As Integer, ByVal pnValorTasacion As Currency, _
        ByVal pnOpeCod As Long, ByVal psOpeDesc As String, ByVal psPersCodCMAC As String, _
        ByVal pbITFAplica As Boolean, ByVal pbITFAsumidoCreditos As Boolean, _
        ByVal pnMontoITF As Currency, Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnMontoCostoNotificacion As Currency = 0, _
        Optional ByRef pnMovNro As Long = 0, Optional ByRef pnIntPend As Double = 0, Optional ByVal pnIntMora As Double = 0, _
        Optional ByVal pnTipoPago As Integer = 1, Optional pnMovNroRVD As Long = 0, Optional pnMovNroPend As Long = 0, _
        Optional ByVal psCtaCodCargo As String = "", Optional ByRef pMatAho As Variant = "", Optional ByRef psCtaAhoExt As String = "", _
        Optional ByRef psClienteExt As String = "", Optional ByRef pnMontoExt As Currency = 0)
        'CTI4 ERS0112020 Add: pnTipoPago, pnMovNroRVD, pnMovNroPend,psCtaCodCargo,pMatAho

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsOpeDesc As String
lsOpeDesc = "Amortización Cred Pign"
'Codigo de Operacion
Select Case pnOpeCod
    Case gColPOpeAmortizEFE  ' En la CMAC
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeAmortNorEFE
        End If
    Case gColPOpeAmortizCHQ
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeAmortNorCHQ
        End If
    Case gColPOpeAmortizEnOtCjEFE  ' En otra CMAC
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeAmortNorEnOtCjEFE
        End If
    Case gColPOpeAmortizEnOtCjCHQ
        If pnIntVenc = 0 Then
            lsOpeCod = gColPOpeAmortNorEnOtCjCHQ
        End If
    Case gColPOpePagoParcialEfectivo '*** PEAC 20161111
            lsOpeCod = gColPOpePagoParcialNorEfectivo
    Case gColPOpePagoParcialNorVoucher 'CTI4 ERS0112020
        lsOpeCod = gColPOpePagoParcialNorVoucher
        lsOpeDesc = "Amortización Cred Pign con voucher"
    Case gColPOpePagoParcialNorCargoCta 'CTI4 ERS0112020
        lsOpeCod = gColPOpePagoParcialNorCargoCta
End Select


'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    
    '*** PEAC 20170524 - si amortiza despues de vencido su estado se mantiene como vencido
    'Call loRegPig.dUpdateProducto(psctacod, , pnNewSaldoCap, gColPEstRenov, psFechaHora, -2, False)          ' (-2) aumenta el ste transac
    Call loRegPig.dUpdateProducto(psctacod, , pnNewSaldoCap, IIf(pnDiasAtraso > 30, gColPEstVenci, gColPEstRenov), psFechaHora, -2, False)      ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    '*** PEAC 20071206 - al momento de amortizar se grava la nueva fecha de venc a 30 dias *****
    'Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)
     Call loRegPig.dUpdateColocaciones(psctacod, , psFecVenci, , , psMovNro, , False)
    '********************************************************************************************

    '** Actualiza ColocPignoraticio ' activado por peac 20071106
    Call loRegPig.dUpdateColocPignoraticio(psctacod, , , , , , , , , , , False)

    '** Insert ColocEstado
'    Call loRegPig.dInsertColocacEstado(psCtaCod, psFechaHora, gColPEstRenov, 1, pnMontoTransac, "Pago Parcial Contrato", _
'        gColocCalendCodPFCF, 0, 0, pnNewPlazo, "0", 0, 0, 0, 0, False)
    
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstRenov, 1, pnMontoTransac, "Amortización Contrato", _
        gColocCalendCodPFCF, 0, 0, pnNewPlazo, "0", 0, 0, 0, 0, False)
    
    
    '** Inserta Mov
    'Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Pago Parcial Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, lsOpeDesc, gMovEstContabMovContable, gMovFlagVigente, False)

    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    pnMovNro = lnMovNro
    '** Inserta MovCol
    '***PEAC 20080125
    'Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, 0, "GMIC", pnNewPlazo, 0, pnNewSaldoCap, False)
    '***PEAC 20080125 SE AGREGO LOS DIAS DE ATRASO PARA PODER ESTORNAR A LA FECHA CORRECTA
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, pnDiasAtraso, "GMIC", pnNewPlazo, 0, pnNewSaldoCap, False)

    '** Inserta MovColDet -  Capital
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCapital, 1, pnCapPag, False)
    '** Inserta MovColDet -  interes comp
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresCompensatorio, 1, pnIntCompensat, False)
    '** Inserta MovColDet -  interes mora
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresMoratorio, 1, pnIntMora, False)
    '** Inserta MovColDet -  interes vencido
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresVencido, 1, pnIntVenc, False)

'    '** Inserta MovColDet -  Interes Compensatorio (al vencimiento)
'    If pnIntPend > 0 Then
'        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodInteresCompensatorio, 1, pnMontoTransac, False)
'    ElseIf pnIntMora > 0 Then '*** PEAC 20170329
'        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodInteresMoratorio, 1, pnMontoTransac, False)
'    ElseIf pnIntPend = 0 And pnIntCompensat > 0 Then
'        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodInteresCompensatorio, 1, pnIntCompensat, False)
'    End If

    If pbITFAplica Then
        If Not pbITFAsumidoCreditos Then
            If Left(gColPOpeAmortizEnOtCjEFE, 4) = Left(lsOpeCod, 4) Then
                '** Inserta MovCol
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroCMACCred, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                '** Inserta MovColDet -  Monto Entregar
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroCMACCred, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
            Else
                If pnTipoPago = gColocTipoPagoVoucher Then 'CTI4 ERS0112020
                    Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroOpePignoVoucher, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                    Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroOpePignoVoucher, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
                ElseIf pnTipoPago = gColocTipoPagoCargoCta Then 'CTI4 ERS0112020
                    Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroOpePignoCargoACta, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                    Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroOpePignoCargoACta, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
                Else
                    ''** Inserta MovCol
                    Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                    ''** Inserta MovColDet -  Monto Entregar
                    Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False)
                End If
            End If
        Else
            If pnTipoPago = gColocTipoPagoVoucher Then
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroOpePignoVoucher, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroOpePignoVoucher, psctacod, 1, gConcITFAsumido, 1, pnMontoITF, False)
            ElseIf pnTipoPago = gColocTipoPagoCargoCta Then 'CTI4 ERS0112020
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroOpePignoCargoACta, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroOpePignoCargoACta, psctacod, 1, gConcITFAsumido, 1, pnMontoITF, False)
            Else
                '** Inserta MovCol
                Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False)
        
                '** Inserta MovColDet -  Monto Entregar
                Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, gConcITFAsumido, 1, pnMontoITF, False)
            End If
        End If
    End If
    
    'CTI4 ERS0112020 **************
    If pnTipoPago = gColocTipoPagoVoucher Then
        loRegPig.actualizarRelacionVoucherCaptacion lnMovNro, pnMovNroRVD
        If pnMovNroPend <> -1 Then loRegPig.ActualizaMovPendientesRend pnMovNroPend, pnMontoTransac + pnMontoITF
    End If
    If pnTipoPago = gColocTipoPagoCargoCta Then
        'loRegPig.CapCargoCuentaAho pMatAho, psCtaCodCargo, pnMontoTransac + pnMontoITF, gAhoCargoCtaAmortizaPignoNor, psMovNro, "Cargo a Cuenta Amort. Pigno.: " & psCtaCod, psFechaHora NDX
        loRegPig.CapCargoCuentaAho pMatAho, psCtaCodCargo, pnMontoTransac + pnMontoITF, gAhoCargoCtaAmortizaPignoNor, psMovNro, "Cargo a Cuenta Amort. Pigno.: " & psctacod, psFechaHora, , , , , , , , , , , , , , True, CDbl(pnMontoITF), True, gITFCobroCargo 'NDX
    End If
    'END CTI4

    If pnMontoCostoNotificacion > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, "2103", 1, pnMontoCostoNotificacion, False)
    End If
    
    '** Inserta MovCMAC - Para Recepcion de Llamada de cmac
    If pnOpeCod = gColPOpeRenovacEnOtCjEFE Or pnOpeCod = gColPOpeRenovacEnOtCjCHQ Then
        Call loRegPig.dInsertMovCMAC(lnMovNro, psPersCodCMAC, Format$(gTpoIFCmac, "00"), CInt(Mid(psctacod, 9, 1)), "", _
                "", lsOpeCod, pnMontoTransac, False)
    End If
    
    'inserta pago parcial de interes
    Call loRegPig.dInsertColocPigPagosParciales(psctacod, psFecVenci, gColPConceptoCodInteresCompensatorio, pnIntCompensat + pnIntPend, pnIntCompensat, lnMovNro, False)
    'Call loRegPig.dInsertColocPigPagosParciales(psCtaCod, psFecVenci, gColPConceptoCodInteresCompensatorio, pnIntCompensat, pnIntPend, lnMovNro, False)
    
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nPagoParcialCredPignoraticio", "Error en Funcion de Pago Parcial de Contrato "

End Sub


Public Sub nSubastaVentaCredPignoraticioSINSubasta(ByVal psctacod As String, ByVal psNroProceso As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnMontoTransac As Currency, ByVal pnVtaNetaSubasta As Currency, _
        ByVal pnImpuestoVtaSubasta As Currency, ByVal pnValorAdjudicacion As Currency, _
        ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, ByVal pnKi21 As Double, _
        Optional pbEjecBatch As Boolean = False, Optional psNroDoc As String = "", Optional psPersCod As String = "", _
        Optional ByVal pbITFAplica As Boolean = False, Optional ByVal pbITFAsumidoCreditos As Boolean = False, _
        Optional ByVal pnMontoITF As Currency = 0, Optional ByVal pnRecupInteres As Currency = 0, _
        Optional pbRecuperacion As Boolean = False, _
        Optional pnMovNro As Long = 0, Optional ByRef pnMovNroResult As Long = 0, _
        Optional pnValorVtaLotAdj As Double = 0, Optional pnDevolVtaLotAdj As Double = 0, _
        Optional pnMontoDsctoInt As Double = 0, _
        Optional ByVal pnTipoPago As Integer = 1, Optional pnMovNroRVD As Long = 0, Optional pnMovNroPend As Long = 0, _
        Optional ByVal psCtaCodCargo As String = "", Optional ByRef pMatAho As Variant = "")
        'CTI4 ERS0112020 Add: pnTipoPago, pnMovNroRVD, pnMovNroPend, psCtaCodCargo, pMatAho

'*** PEAC 20200904 - SE AGREGÓ pnMontoDsctoInt
'*** PEAC 20180712 - Se agregó los parámetros pnValorVtaLotAdj y pnDevolVtaLotAdj para control de precio de
'*** venta en el modulo de venta de Lote Adjudicado

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertColocacEstado
'** dInsertMov  psPersCod
'** Actualiza ColocPigRGDet (Estado, Monto Venta, nMovVta)
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsOpeDescVoucher As String
lsOpeDescVoucher = ""
'*** PEAC 20190807
Dim lnTasacion As Double
Dim lrDataAdj As ADODB.Recordset
Dim loDataRecup As New COMDColocPig.DCOMColPFunciones
'*** FIN PEAC
Dim lsOpeTpoCobroITF As String 'CTI4 ERS0112020
'Codigo de Operacion
If pbRecuperacion Then
   If pnTipoPago = gColocTipoPagoVoucher Then 'CTI4 ERS0112020
        lsOpeCod = gColPOpeRecuperaContratoAdjVoucher
        lsOpeTpoCobroITF = gITFCobroOpePignoVoucher
        lsOpeDescVoucher = " con voucher"
    ElseIf pnTipoPago = gColocTipoPagoCargoCta Then 'CTI4 ERS0112020
        lsOpeCod = gColPOpeRecuperaContratoAdjCargoCta
        lsOpeTpoCobroITF = gITFCobroOpePignoCargoACta
    Else
        lsOpeCod = "122900"
        lsOpeTpoCobroITF = gITFCobroEfectivoPrend
    End If
Else
   lsOpeCod = geColPVtaSubasta
   lsOpeTpoCobroITF = gITFCobroEfectivoPrend
End If

'*** PEAC 20190807
lnTasacion = 0

lsSQL = "exec stp_sel_DatosParaRecupContraAdjudiCustDif '" & psctacod & "'"

Set loDataRecup = New COMDColocPig.DCOMColPFunciones
    Set lrDataAdj = loDataRecup.dObtieneRecordSet(lsSQL)
Set loDataRecup = Nothing

If Not (lrDataAdj.BOF And lrDataAdj.EOF) Then
    lnTasacion = lrDataAdj!nTasacion
End If

'*** FIN PEAC

'On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
        
    If pbRecuperacion = False Then
        '** Actualiza Producto
        Call loRegPig.dUpdateProducto(psctacod, , , gColPEstSubas, psFechaHora, -2, False) ' (-2) aumenta el ste transac
        '** Actualiza Colocaciones
        Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)
        '** Insert ColocEstado
        Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstSubas, 1, pnMontoTransac, "Venta Joyas en Subasta", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
        '** Inserta Mov
    
        Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Venta Joyas Adjudicadas" & lsOpeDescVoucher, gMovEstContabMovContable, gMovFlagVigente, False)
        ' Obtiene nMovNro
        'lnMovNro = loRegPig.dGetnMovNro(psMovNro) RECO COMENTADO
        'RECO20141210 ERS162-2014*****************
        If pnMovNro = 0 Then
            lnMovNro = loRegPig.dGetnMovNro(psMovNro)
            pnMovNroResult = lnMovNro
        Else
            lnMovNro = pnMovNro
            pnMovNroResult = lnMovNro
        End If
        'FIN RECO*********************************
        '** Inserta MovCol
        Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)
        
        '** Inserta MovColDet -  Venta Neta
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, geColPConceptoCodVtaNetaSubasta, 1, pnVtaNetaSubasta, False)
    
        '** Inserta MovColDet -  Impuesto Vta Subasta
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, geColPConceptoCodImpuestoVtaSubasta, 1, pnImpuestoVtaSubasta, False)
        
        '** Inserta MovColDet -  Valor Adjudicacion
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, geColPConceptoCodValorAdjudica, 1, pnValorAdjudicacion, False)
    
        '** Inserta MovColDet ' Oro 14/16/18/21
        If pnki14 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro14, 1, pnki14, False)
        End If
        If pnki16 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro16, 1, pnki16, False)
        End If
        If pnKi18 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro18, 1, pnKi18, False)
        End If
        If pnKi21 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro21, 1, pnKi21, False)
        End If

        '*** PEAC 20190807
        If lnTasacion > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, "2218", 1, lnTasacion, False)
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, "2219", 1, lnTasacion, False)
        End If
        '*** FIN PEAC
        
        'PEAC 20200904
        If pnMontoDsctoInt > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, "2220", 1, pnMontoDsctoInt, False)
        End If
        'FIN PEAC
        
        ''** Inserta Boleta
        If psNroDoc <> "" Then
             Call loRegPig.dInsertMovDoc(lnMovNro, TpoDoc.TpoDocBoletaVenta, psNroDoc, psFechaHora)
        End If
    
        '** Update ColocPigRGDet
        'Call loRegPig.dUpdateColocPigRGDet("S", psNroProceso, psCtaCod, gColPRecGarVentaVendido, pnMontoTransac, , , , lnMovNro, False)
    
        '** Inserta ColocPigRecupVta
        Call loRegPig.dInsertColocPigRecupVta("S", psNroProceso, psctacod, 3, psNroDoc, psPersCod, pnMontoTransac, 0, 0, False)
    
    
    Else
    
        '** Actualiza Producto
        Call loRegPig.dUpdateProducto(psctacod, , , "2113", psFechaHora, -2, False)              ' (-2) aumenta el ste transac
        '** Actualiza Colocaciones
        Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)
        '** Insert ColocEstado
        
        '*** PEAC 20080828
        'Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, "2113", 1, pnMontoTransac, "Recuperacion Joyas Adjudicadas", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
        Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, "2113", 1, pnVtaNetaSubasta, "Recuperacion Joyas Adjudicadas", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
        '** Inserta Mov
        'RECO20141210 ERS162-2014*****************
        If pnMovNro = 0 Then
            Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Recuperacion Joyas Adjudicadas" & lsOpeDescVoucher, gMovEstContabMovContable, gMovFlagVigente, False)
            lnMovNro = loRegPig.dGetnMovNro(psMovNro)
            pnMovNroResult = lnMovNro
        Else
            lnMovNro = pnMovNro
            pnMovNroResult = lnMovNro
        End If
        'FIN RECO*********************************
        
        ' Obtiene nMovNro
        'lnMovNro = loRegPig.dGetnMovNro(psMovNro) RECO COMENTADO
        
        
        '** Inserta MovCol
        
        '*** PEAC 20080828
        Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False) 'APRI20181110 DESCOMENTÓ POR INCIDENTE
        'Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 1, pnVtaNetaSubasta, 0, "", 0, 0, 0, False) 'APRI20181110 COMENTÓ POR INCIDENTE

        '** Inserta MovColDet -  Deuda Vencida
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCapitalVencido, 1, pnVtaNetaSubasta, False)
    
        '** Inserta MovColDet -  Interes
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresMoratorio, 1, pnRecupInteres, False)
    
        ''** Inserta MovColDet -  Impuesto Vta Subasta
        'Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, geColPConceptoCodImpuestoVtaSubasta, 1, pnImpuestoVtaSubasta, False)
        
        '** Inserta MovColDet -  Valor Adjudicacion
        'Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, geColPConceptoCodValorAdjudica, 1, pnValorAdjudicacion, False)
        
        '** Inserta MovColDet - PEAC 20180712 - Valor venta lote adjudicado
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, 2214, 1, pnValorVtaLotAdj, False)
    
        '** Inserta MovColDet - PEAC 20180712 - Devolucion venta lote adjudicado
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, 2215, 1, pnDevolVtaLotAdj, False)
        
        
        '*** PEAC 20190807
        If lnTasacion > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, "2218", 1, lnTasacion, False)
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, "2219", 1, lnTasacion, False)
        End If
        '*** FIN PEAC
        
        'PEAC 20200904
        If pnMontoDsctoInt > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, "2220", 1, pnMontoDsctoInt, False)
        End If
        'FIN PEAC
        
        '** Inserta MovColDet ' Oro 14/16/18/21
        If pnki14 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro14, 1, pnki14, False)
        End If
        If pnki16 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro16, 1, pnki16, False)
        End If
        If pnKi18 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro18, 1, pnKi18, False)
        End If
        If pnKi21 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro21, 1, pnKi21, False)
        End If
        
        If pbITFAplica Then
                    If Not pbITFAsumidoCreditos Then
                '** Inserta MovCol
                'Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psCtaCod, 1, pnMontoITF, 0, "", 0, 0, 0, False) 'CTI4 ERS0112020 Comento
                Call loRegPig.dInsertMovCol(lnMovNro, lsOpeTpoCobroITF, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False) 'CTI4 ERS0112020
            
                '** Inserta MovColDet -  Monto Entregar
                'Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtaCod, 1, gConcITFCliente, 1, pnMontoITF, False) 'CTI4 ERS0112020 Comento
                Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeTpoCobroITF, psctacod, 1, gConcITFCliente, 1, pnMontoITF, False) 'CTI4 ERS0112020
            Else
                '** Inserta MovCol
                'Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psCtaCod, 1, pnMontoITF, 0, "", 0, 0, 0, False) 'CTI4 ERS0112020 Comento
                Call loRegPig.dInsertMovCol(lnMovNro, lsOpeTpoCobroITF, psctacod, 1, pnMontoITF, 0, "", 0, 0, 0, False) 'CTI4 ERS0112020
            
                '** Inserta MovColDet -  Monto Entregar
                'Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtaCod, 1, gConcITFAsumido, 1, pnMontoITF, False) 'CTI4 ERS0112020 Comento
                Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeTpoCobroITF, psctacod, 1, gConcITFAsumido, 1, pnMontoITF, False) 'CTI4 ERS0112020
            End If
        End If
        
        'CTI4 ERS0112020 **************
        If pnTipoPago = gColocTipoPagoVoucher Then
            loRegPig.actualizarRelacionVoucherCaptacion lnMovNro, pnMovNroRVD
            If pnMovNroPend <> -1 Then loRegPig.ActualizaMovPendientesRend pnMovNroPend, pnMontoTransac + pnMontoITF
        End If
        If pnTipoPago = gColocTipoPagoCargoCta Then
            loRegPig.CapCargoCuentaAho pMatAho, psCtaCodCargo, pnMontoTransac + pnMontoITF, gAhoCargoCtaRecupCoAdj, psMovNro, "Cargo a Cuenta Recup. Co. Adj.: " & psctacod, psFechaHora, , , , , , , , , , , , , , True, CDbl(pnMontoITF), True, gITFCobroCargo
        End If
        'END CTI4
 End If
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nPagoSobranteRemate(ByVal psctacod As String, ByVal psNroProceso As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnMontoTransac As Currency, ByVal psCtaAhorroRemate As String, _
         ByVal psLpt As String, Optional pbEjecBatch As Boolean = False, Optional psNroDoc As String = "", Optional ByRef pnNumMovPgoSobRema As Long = 0)

'*** PEAC 20100111 - se agrego el parametro "pnNumMovPgoSobRema" para la implementacion del visto electronico

'** Actualiza Producto ( NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertMov
'** Actualiza ColocPigRGDet (Estado,  nMovVta)
'** Actualiza ColocPigRecupVta (FechaPago)
'** dInsertMovCol
'** dInsertMovColDet (Monto Sobrante)
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = geColPPagSobrante

'On Error GoTo ErrorModPig

    'Retira de la la Cta de Ahorros de sobrante
    '****
    Dim oCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMovNroReg As String
    Dim nMovNroReg As Long
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Set oCont = New COMNContabilidad.NCOMContFunciones
        sMovNroReg = oCont.GeneraMovNro(Format(psFechaHora, "yyyy/mm/dd"), Mid(psMovNro, 18, 2), Mid(psMovNro, 22, 4))
    Set oCont = Nothing
    Set oCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
        'oCapMov.CapAbonoCuentaAho psCtaAhorroRemate, pnMontoTransac, gAhoDepSobRemate, sMovNroReg, "Dep. Sobrante Remate" & psNroRemate & " - " & psProcesoCadaAgencia, , , , , , , , , , psLpt, , True
        oCapMov.CapCargoCuentaAho psCtaAhorroRemate, pnMontoTransac, gAhoRetOtrosConceptos, sMovNroReg, "Ret. Pago Sobrante Remate-" & psctacod, , , , , , , , , , psLpt, , True
    Set oCapMov = Nothing

    '****

    loRegPig.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , , psFechaHora, -2, False)              ' (-2) aumenta el ste transac

    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Pago Sobrante Remate", gMovEstContabMovContable, gMovFlagVigente, False)

    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)

    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("R", psNroProceso, psctacod, gColPRecGarVentaSobranteCobrado, , , , , lnMovNro, False)

    '** Update ColocPigRecupVta
    Call loRegPig.dUpdateColocPigRecupVta("R", psNroProceso, psctacod, , , , , , , psFechaHora, False)

    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet -  Monto Sobrante
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodSobranteRemate, 1, pnMontoTransac, False)

    nMovNroReg = loRegPig.dGetnMovNro(sMovNroReg)
    '*** Relaciona El retiro con el Pago Sobrante
    Call loRegPig.dInsertMovRef(lnMovNro, nMovNroReg)

    pnNumMovPgoSobRema = lnMovNro

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nPagoSobranteRemate", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nExtornoPagoSobranteRemate(ByVal psctacod As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal pnMovNroAnt As Long, _
        ByVal pnMontoTransac As Currency, _
        ByVal psLpt As String, ByVal psNroProceso As String, ByVal psCtaAhorroRemate As String, _
        Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExtorno As Variant)

'** Actualiza Producto ( NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** Actualiza ColocPigRGDet (Estado,  nMovVta)
'** Actualiza ColocPigRecupVta (FechaPago)
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = "129702"
'On Error GoTo ErrorModPig

    'Deposita a la Cta de Ahorros de sobrante
    '****
    Dim oCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMovNroReg As String
    Dim nMovNroReg As Long
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Set oCont = New COMNContabilidad.NCOMContFunciones
        sMovNroReg = oCont.GeneraMovNro(Format(psFechaHora, "yyyy/mm/dd"), Mid(psMovNro, 16, 2), Mid(psMovNro, 22, 4))
    Set oCont = Nothing
    Set oCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
        oCapMov.CapAbonoCuentaAho psCtaAhorroRemate, pnMontoTransac, gAhoDepEfec, sMovNroReg, "Extorno Pago Sobrante Remate-" & psctacod, , , , , , , , , , psLpt, , True
        'oCapMov.CapCargoCuentaAho psCtaAhorroRemate, pnMontoTransac, gAhoRetEfec, psMovNro, "Ret. Pago Sobrante Remate" & psCtacod, , , , , , , , , , psLpt, , True
    Set oCapMov = Nothing

    '****

    loRegPig.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , , psFechaHora, -2, False)              ' (-2) aumenta el ste transac

    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("R", psNroProceso, psctacod, gColPRecGarVentaVendido, , , , , 0, False)

    '** Update ColocPigRecupVta
    'Call loRegPig.dUpdateColocPigRecupVta("R", psNroProceso, psCtacod, , , , , , , , False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Extorno Pago Sobrante Remate", gMovEstContabMovContable, gMovFlagVigente, False)
    
    '**CTI3 (ferimoro) 02102018
    Call loRegPig.dInsertDetExtMov(psMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
  
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)

    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)

    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    nMovNroReg = loRegPig.dGetnMovNro(sMovNroReg)
    '*** Relaciona con la cuenta de ahorros
    Call loRegPig.dInsertMovRef(lnMovNro, nMovNroReg)

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

''*** PEAC 20090313
Public Sub nPagoSobranteAdjudicado(ByVal psctacod As String, ByVal psNroProceso As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnMontoTransac As Currency, ByVal psCtaAhorroAdjudicado As String, _
         ByVal psLpt As String, Optional pbEjecBatch As Boolean = False, Optional psNroDoc As String = "")

'** Actualiza Producto ( NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertMov
'** Actualiza ColocPigRGDet (Estado,  nMovVta)
'** Actualiza ColocPigRecupVta (FechaPago)
'** dInsertMovCol
'** dInsertMovColDet (Monto Sobrante)
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = geColPPagSobraAdjudicado

'On Error GoTo ErrorModPig

    'Retira de la la Cta de Ahorros de sobrante
    '****
    Dim oCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMovNroReg As String
    Dim nMovNroReg As Long
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Set oCont = New COMNContabilidad.NCOMContFunciones
        sMovNroReg = oCont.GeneraMovNro(Format(psFechaHora, "yyyy/mm/dd"), Mid(psMovNro, 18, 2), Mid(psMovNro, 22, 4))
    Set oCont = Nothing
    Set oCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
        oCapMov.CapCargoCuentaAho psCtaAhorroAdjudicado, pnMontoTransac, gAhoRetOtrosConceptos, sMovNroReg, "Ret. Pago Sobrante Adjudicado-" & psctacod, , , , , , , , , , psLpt, , True
    Set oCapMov = Nothing

    '****

    loRegPig.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , , psFechaHora, -2, False)              ' (-2) aumenta el ste transac

    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Pago Sobrante Adjudicado", gMovEstContabMovContable, gMovFlagVigente, False)

    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)

    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("A", psNroProceso, psctacod, gColPRecGarVentaSobranteCobrado, , , , , lnMovNro, False)

'    '** Update ColocPigRecupVta
'    Call loRegPig.dUpdateColocPigRecupVta("A", psNroProceso, psctacod, , , , , , , psFechaHora, False)

    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet -  Monto Sobrante
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodSobranteAdjudicado, 1, pnMontoTransac, False)

    nMovNroReg = loRegPig.dGetnMovNro(sMovNroReg)
    '*** Relaciona El retiro con el Pago Sobrante
    Call loRegPig.dInsertMovRef(lnMovNro, nMovNroReg)

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nPagoSobranteAdjudicado", "Error en Funcion de Registro de Contrato "

End Sub

'*** PEAC 20090317
Public Sub nExtornoPagoSobranteAdjudica(ByVal psctacod As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal pnMovNroAnt As Long, _
        ByVal pnMontoTransac As Currency, _
        ByVal psLpt As String, ByVal psNroProceso As String, ByVal psCtaAhorroAdjudica As String, _
        Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExtorno As Variant)

'** Actualiza Producto ( NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** Actualiza ColocPigRGDet (Estado,  nMovVta)
'** Actualiza ColocPigRecupVta (FechaPago)
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = "129705"  ''EXTORNO PAGO SOBRANTE ADJUDICADO
'On Error GoTo ErrorModPig

    'Deposita a la Cta de Ahorros de sobrante
    '****
    Dim oCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMovNroReg As String
    Dim nMovNroReg As Long
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Set oCont = New COMNContabilidad.NCOMContFunciones
        sMovNroReg = oCont.GeneraMovNro(Format(psFechaHora, "yyyy/mm/dd"), Mid(psMovNro, 16, 2), Mid(psMovNro, 22, 4))
    Set oCont = Nothing
    Set oCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
        oCapMov.CapAbonoCuentaAho psCtaAhorroAdjudica, pnMontoTransac, gAhoDepEfec, sMovNroReg, "Extorno Pago Sobrante Adjudicado-" & psctacod, , , , , , , , , , psLpt, , True
    Set oCapMov = Nothing

    '****

    loRegPig.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , , psFechaHora, -2, False)

    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("A", psNroProceso, psctacod, gColPRecGarVentaVendido, , , , , 0, False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Extorno Pago Sobrante Adjudicado", gMovEstContabMovContable, gMovFlagVigente, False)
    
    '**CTI3 (ferimoro) 02102018
    Call loRegPig.dInsertDetExtMov(psMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
  
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)

    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)

    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    nMovNroReg = loRegPig.dGetnMovNro(sMovNroReg)
    '*** Relaciona con la cuenta de ahorros
    Call loRegPig.dInsertMovRef(lnMovNro, nMovNroReg)

    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

'''*** PEAC 20090505
'''*** Desembolso de credito pignoraticio en efcetivo y con abono a cuenta de ahorro

Public Sub nDesembolsoCredPignoEfectivoAbono(ByVal psctacod As String, ByVal pbAbonoCtaAho As Boolean, ByVal pbCtaAhoNva As Boolean, ByVal pnSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal pnMontoEntregar As Currency, _
        ByVal pnInteresCompensat As Currency, ByVal pnImpuesto As Currency, ByVal pnCostoTasacion As Currency, ByVal pnCostoCustodia As Currency, _
        ByVal pbITFAplica As Boolean, ByVal pbITFAsumidoCreditos As Boolean, ByVal pnMontoITf1 As Currency, ByVal pnMontoITf2 As Currency, ByVal pnMontoDesembolsar As Currency, _
        ByRef pnMovNro As Long, Optional pbEjecBatch As Boolean = False, _
        Optional loRegPig2 As COMDColocPig.DCOMColPActualizaBD, Optional pbTrans As Boolean = False)
        'RECO20140208 ERS002 SE AGREGO:loRegPig2,pbTrans
'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

''Codigo de Operacion
If pbAbonoCtaAho Then
    If pbCtaAhoNva Then
        lsOpeCod = gColPOpeDesembolsoAboCtaNva '' desembolso con abono cta Nueva
    Else
        lsOpeCod = gColPOpeDesembolsoAboCta '' desembolso con abono en cta ahorro existente
    End If
Else
    lsOpeCod = gColPOpeDesembolsoEFE  '' desembolso en efectivo
End If

On Error GoTo ErrorModPig

    'RECO20140206 ERS002*********************************
    If pbTrans = False Then
        loRegPig.dBeginTrans
    Else
        Set loRegPig = loRegPig2
    End If
    'loRegPig.dBeginTrans
    mbTrans = True
    'RECO FIN********************************************

'''**************************** descomentar desde aqui
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , gColPEstDesem, psFechaHora, -2, False)    ' (-2) aumenta el ste transac

    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstDesem, 1, pnSaldoCap, "Desembolso Contrato", _
        gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Desembolso Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)

    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)

    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMontoDesembolsar, 0, "", 0, 0, pnSaldoCap, False)

    '** Inserta MovColDet -  Monto Entregar

    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodMontoEntregar, 1, pnMontoDesembolsar, False)

    '** Inserta MovColDet -  Interes Compensatorio
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodInteresCompensatorio, 1, pnInteresCompensat, False)

    '** Inserta MovColDet -  Impuesto
    If pnImpuesto > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodImpuesto, 1, pnImpuesto, False)
    End If
    '** Inserta MovColDet -  Costo Custodia
    If pnCostoCustodia > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodCustodia, 1, pnCostoCustodia, False)
    End If
    '** Inserta MovColDet -  Costo Tasacion
    If pnCostoTasacion > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodTasacion, 1, pnCostoTasacion, False)
    End If

    If pbITFAplica Then
        If Not pbITFAsumidoCreditos Then
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, pnMontoITf1 + pnMontoITf2, 0, "", 0, 0, pnSaldoCap, False)

            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, gConcITFCliente, 1, pnMontoITf1, False)

            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, 22, 1, pnMontoITf2, False)
        Else
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, pnMontoITf1 + pnMontoITf2, 0, "", 0, 0, pnSaldoCap, False)

            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, gConcITFAsumido, 1, pnMontoITf1, False)

            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psctacod, 1, 22, 1, pnMontoITf2, False)
        End If
    End If
    
    'WIOR 20140203 ************************
    Call loRegPig.ActualizaCredVinculados(psctacod, "", 2, 5)
    'WIOR FIN *****************************
    
    pnMovNro = lnMovNro
    
    'RECO20140206 ERS002*********************************
    If pbTrans = False Then
        loRegPig.dCommitTrans
        mbTrans = False
        Set loRegPig = Nothing
    End If
    'loRegPig.dCommitTrans
    'RECO FIN********************************************

    'loRegPig.dCommitTrans
    mbTrans = False
'Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
    'RECO20140206 ERS002*********************************
        If pbTrans = False Then
            loRegPig.dRollbackTrans
        End If
        mbTrans = False
    'RECO FIN********************************************

    End If
    Err.Raise vbObjectError + 100, "Error nDesembolsoCredPignoEfectivoAbono", "Error en Funcion de desembolso de credito pignoraticio "
End Sub

''*** PEAC 20090507
Public Function RecuperaAgencia(ByVal psAge As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaAgencia
    sSql = "Select * from Agencias Where cAgeCod = '" & psAge & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaAgencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaAgencia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'''*** PEAC 20090508
Public Function DesembolsoPignoConAbonoCta(ByRef pnMovNro As Long, ByRef pMatDatosAho As Variant, ByVal psctacod As String, ByVal pnprestamo As Double, ByVal pnMontoDesembolso As Double, _
        ByVal pMatGastos As Variant, ByVal MatCargoAutom As Variant, ByVal pdHoy As Date, ByVal psCodAge As String, psCodUser As String, _
        ByVal pbDesembCC As Boolean, Optional ByVal pbCtaAhoNueva As Boolean = False, _
        Optional ByVal MatCredCanc As Variant = Nothing, Optional ByVal pnContMatCredCanc As Integer = -1, Optional ByVal psCtaAbo As String = "", _
        Optional ByVal pRSRela As ADODB.Recordset = Nothing, Optional ByVal pnTasa As Double = 0#, Optional ByVal pnPersoneria As Integer = 0, _
        Optional ByVal pnTipoCuenta As Integer = 0, Optional ByVal pnTipoTasa As Integer = 0, Optional ByVal pbDocumento As Boolean = False, _
        Optional ByVal psNroDoc As String = "", Optional ByVal psCodIF As String = "", _
        Optional ByVal pnMontoGastos As Double = 0#, Optional ByVal pnMontoCancelaciones As Double = 0#, Optional psPersLavDinero As String = "", _
        Optional ByVal pnITF As Double = 0#, _
        Optional pbITFAplica As Boolean = True, Optional pnITFAbono As Double = 0, Optional pnITFGasto As Double = 0, Optional pnITFCancelacion As Double = 0, Optional pbITFAsumido As Boolean = False, _
        Optional psItfOperacion As String = "", _
        Optional ByVal psProyectoActual As String = "", Optional ByVal pvbDesembCheque As Boolean = False, Optional ByVal psBanco As String = "", Optional ByVal psNumCheque As String = "", Optional ByVal psCtaBco As String = "", Optional ByRef psCtaAboR As String, _
        Optional ByVal pMatDatosAhoNew As Variant = "", Optional ByVal pnMontoPoliza As Double = 0, Optional psVisLavDinero As String = "", Optional ByVal psMovNro As String = "", _
        Optional loRegPig2 As COMDColocPig.DCOMColPActualizaBD = Nothing, Optional pbTrans As Boolean = False) As String 'DAOR 20070511,psVisLavDinero

Dim R As ADODB.Recordset
'Dim oCredito As COMDCredito.DCOMCredito
'Dim oBase As COMDCredito.DCOMCredActBD
'Dim oLinea As COMDCredito.DCOMLineaCredito
'Dim oCalend As NCOMCalendario
'Dim oDCalend As COMDCredito.DCOMCalendario
'Dim oGastos As NCOMGasto

Dim oBase As New COMDColocPig.DCOMColPActualizaBD

Dim sLineaCod As String
Dim nMontoReservado As Double
Dim nMontoColocado As Double
Dim MatCalendPagos As Variant
Dim MatCalendPagos_2 As Variant
Dim MatGracia As Variant
Dim nPeriodoGracia As Integer
Dim nTasaInteres As Double
Dim nNroCalen As Integer
Dim nNroProxDesemb As Integer
Dim nCuotaCom As Integer
Dim nMiViv As Integer
Dim i As Integer
Dim MatGastos As Variant
Dim MatDesemb() As String
Dim ContDesemb As Integer
Dim nNumGastos As Integer
Dim nNumTransac As Long
Dim sMovNro As String
Dim nmovnro As Long
Dim nMovNroAnt As Long
Dim nConsCred As String
Dim sMetLiquid As String
Dim bTransac As Boolean
Dim MatDesParcial As Variant

'Para MiVivienda
'Dim oParam As COMDCredito.DCOMParametro
Dim nTramoNoConsMonto As Double
Dim nTramoConsMonto As Double
Dim nTramoNoConsPorcen As Double
Dim nTramoConsTasa As Double
Dim nMontoGarantia As Double
'Dim oGar As COMDCredito.DCOMGarantia
Dim nCuotaTemp As String
Dim dFecTran As Date

'Para Cancelacion de Creditos
Dim MatCalendTemp As Variant
Dim MatCalendDistrTemp As Variant
Dim sMetLiqTmp As String
Dim RTmp As ADODB.Recordset
Dim MatRSCancel() As ADODB.Recordset
Dim MatRSCancelCalPend() As Variant
Dim RCalendActual As ADODB.Recordset
Dim RCalendDesembActual As ADODB.Recordset
Dim RColocCalendDesembActual As ADODB.Recordset

'Para Apertura de Cuenta de Ahorros
Dim sCtaApe As String
Dim nMivivMontoCondonadoCapital As Double
Dim nMivivMontoCondonadoInteres As Double
Dim sTipoGasto As String
Dim sConsAhoRetGas As String
Dim sConsAhoRetCan As String

Dim bAmpliado As Boolean
'Dim oAmpliado As COMDCredito.DCOMAmpliacion
Dim oITF As COMDConstSistema.FCOMITF

Dim sOpeCodItfAbonoCuenta As String
Dim bAbonoCuenta As Boolean

Dim sEstadoCreditoParcial As String
'Dim oDCred As COMDCredito.DCOMCredito
Dim nDias As Integer
Dim nPeriodo As Integer
Dim sSql As String

'**DAOR 20070413*********************************
Dim rPlzFjoGar As ADODB.Recordset
Dim sMovNroBloq As String
Dim nMovNroBloq As Long
Dim sOperacionBloq As String
'************************************************

        sMovNro = psMovNro

        On Error GoTo ErrorDesembolsoPignoConAbonoCta
        bAbonoCuenta = False
        
        Set oITF = New COMDConstSistema.FCOMITF
        
        oITF.fgITFParametros

        
        'inicializa Datos de Ahorros
        pMatDatosAho(0) = "" 'Cuenta de Ahorros
        pMatDatosAho(1) = "0.00" 'Monto de Apertura
        pMatDatosAho(2) = "0.00" 'Interes Ganado de Abono
        pMatDatosAho(3) = "0.00" 'Interes Ganado de Retiro Gastos
        pMatDatosAho(4) = "0.00" 'Interes Ganado de Retiro Cancelaciones
        pMatDatosAho(5) = "0.00" 'Monto de Abono
        pMatDatosAho(6) = "0.00" 'Monto de Retiro de Gastos
        pMatDatosAho(7) = "0.00" 'Monto de Retiro de Cancelaciones
        pMatDatosAho(8) = "0.00" 'Saldo Disponible Abono
        pMatDatosAho(9) = "0.00" 'Saldo Contable Abono
        pMatDatosAho(10) = "0.00" 'Saldo Disponible Retiro de Gastos
        pMatDatosAho(11) = "0.00" 'Saldo Contable Retiro de Gastos
        pMatDatosAho(12) = "0.00" 'Saldo Disponible Retiro de Cancelaciones
        pMatDatosAho(13) = "0.00" 'Saldo Contable Retiro de Cancelaciones

        bTransac = False


        If pbDesembCC Then
            If pbCtaAhoNueva Then
                    nConsCred = gColPOpeDesembolsoAboCtaNva
            Else
                If Mid(psCtaAbo, 4, 2) <> Mid(psctacod, 4, 2) Then 'Cuenta de Otra Age
                    nConsCred = gColPOpeDesembolsoAboCtaOtraAge
                Else
                    nConsCred = gColPOpeDesembolsoAboCta
                End If
            End If
            bAbonoCuenta = True
        End If
        
        '********************************************************************
        'INICIO DE TRANSACCION
        '********************************************************************
        'RECO20140206 ERS002*********************************
        If pbTrans = False Then
            oBase.dBeginTrans
        Else
            Set oBase = loRegPig2
        End If
        'loRegPig.dBeginTrans
        mbTrans = True
        'RECO FIN********************************************
        'Call oBase.dBeginTrans

        bTransac = True


        '**********************************************************
        'Apertura de Cuenta de Ahorro
        '**********************************************************
        If pbDesembCC And pbCtaAhoNueva Then
                '***** Cobrar ITF
                sCtaApe = oBase.CapAperturaCuenta(gCapAhorros, CInt(Mid(psctacod, 9, 1)), _
                         pRSRela, sMovNro, psCodAge, gColPOpeDesembolsoAboCta, pnTasa, 0, pdHoy, 1, _
                          pnPersoneria, "", pnTipoCuenta, pnTipoTasa, pbDocumento, psNroDoc, psCodIF, _
                         , , , , , , , , CInt(pMatDatosAhoNew(0)), CDbl(pMatDatosAhoNew(1)), CInt(pMatDatosAhoNew(2)), CStr(pMatDatosAhoNew(3)))
                '**********************
                psCtaAbo = sCtaApe
                psCtaAboR = sCtaApe
                pMatDatosAho(1) = Format(pnprestamo, "#0.00")
                
        End If
        pMatDatosAho(0) = psCtaAbo
        '**********************************************************

 
        '*****************************************************************
        'Actualiza Cuentas de Ahorro en caso de Abono Cta
        '*****************************************************************
        If pbDesembCC And Not pbCtaAhoNueva Then
            oBase.CapAbonoCuentaAho pnMovNro, pMatDatosAho, psCtaAbo, pnprestamo, nConsCred, sMovNro, "", , , , , , , pdHoy, "", pbITFAplica, pnITFAbono, pbITFAsumido, psItfOperacion
            pMatDatosAho(5) = pnprestamo
            
        End If
        
        'Para que exonere de inactivas
        If pbDesembCC Then
            oBase.ExoneraInactivas psCtaAbo, sMovNro
        End If

        '***************************************************************************************
     
        'Si es con Apertura Abonar el Monto del Prestamo
        If pbDesembCC And pbCtaAhoNueva Then
            oBase.CapAbonoCuentaAho pnMovNro, pMatDatosAho, sCtaApe, pnprestamo, nConsCred, sMovNro, "", , , , , , , pdHoy, "", pbITFAplica, pnITFAbono, pbITFAsumido, psItfOperacion
            pMatDatosAho(5) = pnprestamo
        End If
            
        '***************************************************************************************
            
        'RECO20140207 ERS002****************************************
        If pbTrans = False Then
            oBase.dCommitTrans
        End If
        'Call oBase.dCommitTrans 'ARCV 30-04-2007
        'RECO FIN***************************************************

            
        Set oBase = Nothing

        Exit Function

ErrorDesembolsoPignoConAbonoCta:
    If bTransac Then
        'RECO20140208 ERS002****************
        If pbTrans = False Then
            Call oBase.dRollbackTrans
        End If
        'RECO FIN **************************
        Set oBase = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


'*** PEAC 20090609
Public Function RecuperaMovimientosAhorros(ByVal pnMov As Long, Optional ByVal pbSinMovRef As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConec As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion

    sSql = " exec stp_sel_RecuperaMovimientosAhorros " & pnMov

    Set RecuperaMovimientosAhorros = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

'*** PEAC 20090609
Public Function RecuperaMovimientoCapataciones(ByVal pnMovNro As Long) As ADODB.Recordset

    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim R As ADODB.Recordset

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    
    sSql = " exec stp_sel_RecuperaMovimientoCapataciones " & pnMovNro
    
    Set RecuperaMovimientoCapataciones = oConec.CargaRecordSet(sSql)
    
    oConec.CierraConexion
    Set oConec = Nothing
    
End Function

'*** MAVM 20100717 JG
Public Function RegistrarJoyaGarantia(ByVal psPersCod As String, ByVal pnValTasacion As Currency, ByVal pnPiezas As Integer, ByVal pnPesoBruto As Currency, ByVal pnPesoNeto As Currency, ByVal psMovNro As String, ByVal prjoyas As ADODB.Recordset) As String
Dim loGen As COMDConstSistema.DCOMGeneral
Dim lsCodJoyas As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
On Error GoTo ErrorRegPig
Set loGen = New COMDConstSistema.DCOMGeneral
    lsCodJoyas = loGen.GeneraNvoNumeroJoyGar()
Set loGen = Nothing
    loRegPig.dBeginTrans
    mbTrans = True

    Call loRegPig.dInsertarJoyaGarantia(lsCodJoyas, psPersCod, pnValTasacion, pnPiezas, pnPesoBruto, pnPesoNeto, psMovNro, prjoyas)
    RegistrarJoyaGarantia = lsCodJoyas
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Function

ErrorRegPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraJoyaGarantia", "Error en Funcion de Registro de Joya Garantia "
End Function

Public Sub nRescataGarantJoya(ByVal psctacod As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, ByVal pnKi21 As Double, _
        ByVal pnDiasCustodia As Integer, ByVal pnValTasac As Double, _
        ByVal psOperacion As String, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = psOperacion

    loRegPig.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , gColPEstCance, psFechaHora, -2, False)    ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstCance, 1, 0, "Entrega Joya", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Entrega Joya Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, 0, 0, "", 0, 0, 0, False)
    '** Insert MovColDet Valor Tasacion
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodTasacion, 1, pnValTasac, False)

    '** Inserta MovColDet ' Oro 14/16/18/21
    If pnki14 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro14, 1, pnki14, False)
    End If
    If pnki16 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro16, 1, pnki16, False)
    End If
    If pnKi18 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro18, 1, pnKi18, False)
    End If
    If pnKi21 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psctacod, 1, gColPConceptoCodOro21, 1, pnKi21, False)
    End If

    loRegPig.dCommitTrans
    mbTrans = False
    
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "
End Sub

Public Sub nRescataJoyaGarantia(ByVal psJoyGarCod As String, Optional pnEstado As Integer, Optional psMovNroRes As String, _
        Optional pbEjecBatch As Boolean = False)
        
Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

    loRegPig.dBeginTrans
    mbTrans = True

    Call loRegPig.dUpdateJoyaGarantia(psJoyGarCod, pnEstado, psMovNroRes, -2, False)
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
    
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub
'RECO20140130 ERS002 RECO-N***************************************************************************
Public Sub nInsertaRegistroCreditoPignoAmpliado(ByVal psctacod As String, ByVal psCtaCodAmp As String, ByVal pnEstado As Integer, ByVal psEstadoFec As String, ByVal nMonto As Double, ByVal nMontoAmp As Double, ByVal psUser As String)
        
Dim lsSQL As String
Dim loDCONColPContrato As New COMDColocPig.DCOMColPContrato
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

    loRegPig.dBeginTrans
    mbTrans = True

    Call loDCONColPContrato.InsertaRegistroCreditoPignoAmpliado(psctacod, psCtaCodAmp, pnEstado, psEstadoFec, nMonto, nMontoAmp, psUser)
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loRegPig.dCommitTrans
    mbTrans = False
    
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub
Public Function ObtieneEstadoCredPignoAmp(ByVal psctacod As String) As Integer
    Dim oDCOMColPContrato As COMDColocPig.DCOMColPContrato
    Dim lrCredPig As ADODB.Recordset
    Set oDCOMColPContrato = New COMDColocPig.DCOMColPContrato
    Set lrCredPig = oDCOMColPContrato.ObtieneCuentaCancPignoAmpliado(psctacod)
    If Not (lrCredPig.EOF And lrCredPig.BOF) Then
        ObtieneEstadoCredPignoAmp = lrCredPig!nEstado
    Else
        ObtieneEstadoCredPignoAmp = 0
    End If
End Function

Public Function ListaCredPignoEstado(ByVal psEstadoCad As String, ByVal psAgeCod As String, ByVal pnTpoBusca As Integer, ByVal psPersCod As String) As ADODB.Recordset
    Dim oDCOMColPContrato As COMDColocPig.DCOMColPContrato
    Set oDCOMColPContrato = New COMDColocPig.DCOMColPContrato
    Set ListaCredPignoEstado = oDCOMColPContrato.ListaCredPignoEstado(psEstadoCad, psAgeCod, pnTpoBusca, psPersCod)
End Function

Public Function ObtieneCuentaCancPignoAmpliado(ByVal psctacod As String) As ADODB.Recordset
    Dim oDCOMColPContrato As COMDColocPig.DCOMColPContrato
    Set oDCOMColPContrato = New COMDColocPig.DCOMColPContrato
    Set ObtieneCuentaCancPignoAmpliado = oDCOMColPContrato.ObtieneCuentaCancPignoAmpliado(psctacod)
End Function

Public Sub InsertaMovRedondeoITF(ByVal psMovNro As String, ByVal pnMovItem As Integer, ByVal pnITFCalculado As Currency, ByVal pnITFRedondeado As Currency, Optional ByVal pnMovNro As Long = 0, Optional loRegPig2 As COMDColocPig.DCOMColPActualizaBD, Optional pbTrans As Boolean = False)
    Dim lsSQL As String
    Dim lnMovNro As Long
    Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
    
    On Error GoTo InsertaMovRedondeoITFErr
        If pnMovNro > 0 Then
            lnMovNro = pnMovNro
        Else
            lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        End If
        
        If pbTrans = False Then
            loRegPig.dBeginTrans
        Else
            Set loRegPig = loRegPig2
        End If
        
        mbTrans = True
        
        lsSQL = " exec stp_ins_MovRedondeoITF " & lnMovNro & "," & pnMovItem & "," & pnITFCalculado & "," & pnITFRedondeado
        
        loRegPig.dEjecutar lsSQL
        
        If pbTrans = False Then
            loRegPig.dCommitTrans
        End If
        mbTrans = False
        
        Exit Sub
InsertaMovRedondeoITFErr:
        If mbTrans Then
            If pbTrans = False Then
                loRegPig.dRollbackTrans
            End If
            mbTrans = False
        End If
        Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "
End Sub

Public Sub nExtornoDesembolsoCredPigAmpliado(ByVal psctacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False, Optional ByRef pbProcedeExtAho As Boolean = False, _
        Optional ByRef pnITF As Currency = 0, Optional ByRef psCtaAhoExt As String = "", Optional ByVal psCodUser As String = "SIST", _
        Optional ByRef psmensaje As String = "", _
        Optional ByVal psMotExtorno As Variant)
'*** PEAC 20090703 - se aumento los parametros pnITF, psCtaAhoExt

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol

'*********************************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsOpeCodExtAcIn As String 'ALPA20130827
lsOpeCod = geColPExtDesemb

'**********************************************************
'RECO20140326**********************************************
Dim loCredPig As New COMDColocPig.DCOMColPContrato
Set loCredPig = New COMDColocPig.DCOMColPContrato
'RECO FIN**************************************************
Dim ofun As COMNContabilidad.NCOMContFunciones
Dim lbProcedeExtAho As Boolean
Dim sMovNroCap As String
Dim RCapMov, RCap As ADODB.Recordset
Set RCapMov = RecuperaMovimientosAhorros(pnMovNroAnt)
Set ofun = New COMNContabilidad.NCOMContFunciones
lbProcedeExtAho = False

If Not (RCapMov.EOF And RCapMov.BOF) Then
    'ALPA20130827**********************************************************************
    'sMovNroCap = oFun.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
    'sMovNroCap = ofun.GeneraMovNro(CDate(Format(psFechaHora, "dd/mm/yyyy")), Mid(RCapMov("cctacod"), 4, 2), "PEAC")
    sMovNroCap = ofun.GeneraMovNro(CDate(Format(psFechaHora, "dd/mm/yyyy")), Mid(RCapMov("cctacod"), 4, 2), psCodUser)
    '**********************************************************************************
    lbProcedeExtAho = True
End If

pbProcedeExtAho = lbProcedeExtAho

On Error GoTo ErrorModPig
    loRegPig.dBeginTrans
    mbTrans = True
'**************************
psCtaAhoExt = ""
'DESEMBOLSO**********************************************************************************

If lbProcedeExtAho Then

    Set RCap = RecuperaMovimientoCapataciones(pnMovNroAnt)
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    psCtaAhoExt = RCap!cCtaCod
    Do While Not RCap.EOF
        If RCap!nMonto > 0 Then
            If Mid(RCap!copecod, 1, 2) <> "99" Then
                'ALPA20130827*******************
                If Left(RCap!copecod, 4) = "2008" Then
                    lsOpeCodExtAcIn = gCTSExtDepActiInact
                Else
                    lsOpeCodExtAcIn = gAhoExtDepEfec
                End If
                '********************************
                Call loRegPig.CapExtornoAbonoAho(sMovNroCap, lnMovNro, pnMovNroAnt, lsOpeCodExtAcIn, RCap!cCtaCod, psMovNro, "Extorno de Desembolso Abono Cuenta", RCap!nMonto)
                'Call loRegPig.CapExtornoAbonoAho(sMovNroCap, lnMovNro, pnMovNroAnt, gAhoExtDepEfec, RCap!cCtaCod, psMovNro, "Extorno de Desembolso Abono Cuenta", RCap!nMonto)
            Else
                Call loRegPig.CapExtornoCargoAho(lnMovNro, 0, gITFCobroCargoExt, RCap!cCtaCod, sMovNroCap, "Extorno de Desembolso Abono Cuenta", RCap!nMonto)
                pnITF = RCap!nMonto
            End If
        End If
        RCap.MoveNext
    Loop
End If

'**************************
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , , gColPEstRegis, psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, gColPEstRegis, 1, pnMonto, "Extorno Desembolso Cont", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Extorno Desembolso Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '***CTI3
    Call loRegPig.dInsertDetExtMov(lnMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
'RECO20140326**************************************
    Dim loDRPigEstadoAnt As ADODB.Recordset
    Dim lsCtaCodAmp As String
    Dim lnCtaCodAmpEst As Integer
    Set loDRPigEstadoAnt = New ADODB.Recordset
    
    Set loDRPigEstadoAnt = loCredPig.RecuperaEstadoCredPignoAnterior(psctacod)
    If Not (loDRPigEstadoAnt.BOF And loDRPigEstadoAnt.EOF) Then
        lsCtaCodAmp = loDRPigEstadoAnt!cCtaCod
        lnCtaCodAmpEst = loDRPigEstadoAnt!nPrdEstado
        'Call loRegPig.dUpdateProducto(lsCtaCodAmp, , , loDRPigEstadoAnt!nPrdEstado, psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    End If
    'RECO FIN******************************************
 'FIN DESEMBOLSO********************************************************************************************************************
 'RESCATE***************************************************************************************************************************
    'Dim oPigCredAnt As ADODB.Recordset
    'Set oPigCredAnt = loRegPig.ObtieneCtaRescateXAmpliacion(pnMovNroAnt)
    'psctacod = oPigCredAnt!cCtaCd
    'Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
 'FIN RESCATE***********************************************************************************************************************
 'CANCELACION***********************************************************************************************************************
    Dim lrDataExt As ADODB.Recordset
    lsOpeCod = geColPExtRenov
    psctacod = lsCtaCodAmp
    Set lrDataExt = loRegPig.dRecuperaDatosExtornoCancCredPig(psctacod, pnMovNroAnt)
    If lrDataExt Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Extorno "
        Exit Sub
    End If
    If lrDataExt.BOF And lrDataExt.EOF Then
        psmensaje = "ERROR: No encontro datos de la Operacion a extornar "
        Exit Sub
    End If
    ' Asigna los valores
    Dim lnAntSaldoCap As Currency
    Dim lnAntEstado As Integer
    Dim lnAntPlazo As Integer
    Dim lsAntFecVenc As String
    lnAntSaldoCap = lrDataExt!nSaldoCap + lrDataExt!nCapitalPag
    lnAntEstado = lrDataExt!nPrdEstadoAnt
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psctacod, , lnAntSaldoCap, lnCtaCodAmpEst, psFechaHora, -2, False)       ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psctacod, , , , , psMovNro, , False)
    '*** PEAC 20080715
    Call loRegPig.dUpdateColocPigNotifi(psctacod, , 0, psMovNro, False)
    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psctacod, psFechaHora, lnAntEstado, 1, pnMonto, "Extorno Cancelacion Cred Pig", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Cancelacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    ' Obtiene nMovNro
    'lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psctacod, 1, pnMonto, 0, "", 0, 0, lnAntSaldoCap, False)
    '** Update Mov Anterior
    'Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)
    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
    '*** Gasto de Correspondencia
    If lrDataExt!nGastoCorrespond > 0 Then
        'Actualiza ColocCalendDetPig
        lsSQL = "UPDATE ColocCalendDetPig Set nMontoPagado = nMontoPagado - " & lrDataExt!nGastoCorrespond & _
                " WHERE cCtaCod ='" & psctacod & "' And nNroCalen = 1 And nColocCalendApl = 0 " & _
                " And nCuota = 1 And nPrdConceptoCod = 2215 "
        loRegPig.dEjecutar lsSQL
    End If
 'FIN CANCELACION*******************************************************************************************************************
    loRegPig.dCommitTrans
    mbTrans = False
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

'RECO FIN**********************************************************************************************************
'RECO20140311 ERS160-2013******************************************************************************************
Public Function ObtieneMontoDesembCreDPigMes(ByVal psPersCod As String, ByVal psPeriodo As String) As Double
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Dim ODR As ADODB.Recordset
    
    Set oCredP = New COMDColocPig.DCOMColPContrato
    Set ODR = New ADODB.Recordset
    
    Set ODR = oCredP.ObtieneMontoDesembCreDPigMes(psPersCod, psPeriodo)
    If Not (ODR.BOF And ODR.EOF) Then
        ObtieneMontoDesembCreDPigMes = ODR!nMonto
    Else
        ObtieneMontoDesembCreDPigMes = 0
    End If
    Set ODR = Nothing
End Function
'RECO FIN**********************************************************************************************************

'RECO20140707 ERS074-2014********************************************************************************
'Modificación: TORE20180417 ERS054-2017
Public Function ListaCredPrepRetasacion(ByVal psAgeCad As String, ByVal psFecDesde As String _
                                        , ByVal psFecHasta As String, ByVal pnEstadoJoy As Integer _
                                        , ByVal pnTrimestre As Integer, ByVal psAnio As String) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set ListaCredPrepRetasacion = oCredP.ListaCredPrepRetasacion(psAgeCad, psFecDesde, psFecHasta, pnEstadoJoy, pnTrimestre, psAnio)
End Function
'Modificación: TORE20180417 ERS054-2017
Public Function DevuelveCtaCodCredPRetasacion(ByVal psAgeCod As String, ByVal psTop As String, ByVal psEstado As String _
                                                , ByVal psFecDesde As String, ByVal psFecHasta As String _
                                                , ByVal pnTrimestre As Integer, ByVal psAnio As String) As String
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Dim ODR As New ADODB.Recordset
    Set ODR = oCredP.DevuelveCtaCodCredPRetasacion(psAgeCod, psTop, psEstado, psFecDesde, psFecHasta, pnTrimestre, psAnio)
    DevuelveCtaCodCredPRetasacion = ""
    If Not (ODR.EOF And ODR.BOF) Then
        DevuelveCtaCodCredPRetasacion = ODR!cCtaCad
    End If
End Function
'Modificación: TORE20180417 ERS054-2017
Public Function DevuelveDatosListaCredPrepRetasacion(ByVal psCtaCad As String, ByVal pnEstado As Integer, ByVal pnCodigoID As Integer) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set DevuelveDatosListaCredPrepRetasacion = oCredP.DevuelveDatosListaCredPrepRetasacion(psCtaCad, pnEstado, pnCodigoID)
End Function
'TORE ERS054-2017
Public Function RegistraPreparacionRetasacion(ByVal psAgeCod As String, ByVal pnTpoPrepara As Integer, _
                                                ByVal pnNumCred As Integer, ByVal pdfecha As String, _
                                                ByVal psUser As String, ByVal psCodRetasacion As String, _
                                                ByVal pnTotLotes As Integer, ByVal pnCriterio As Integer, _
                                                ByVal pdFechaIni As String, ByVal pdFechaFin As String, _
                                                ByVal pnTrimestre As Integer, ByVal psAnio As String) As Integer
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Dim ODR As New ADODB.Recordset
    Set ODR = oCredP.RegistraPreparacionRetasacion(psAgeCod, pnTpoPrepara, pnNumCred, pdfecha, _
                                                   psUser, psCodRetasacion, pnTotLotes, pnCriterio, _
                                                   pdFechaIni, pdFechaFin, pnTrimestre, psAnio)  '[TORE RFC1811260001: ADD pnTotLotes]
    If Not (ODR.EOF And ODR.BOF) Then
        RegistraPreparacionRetasacion = ODR!nCodigoID
    End If
End Function
'TORE ERS054-2017
Public Function RegistraPreparacionRetasacionDet(ByVal pnCodigoID As Integer, ByVal psctacod As String, ByVal pnNroRetasacion As String) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    oCredP.RegistraPreparacionRetasacionDet pnCodigoID, psctacod, pnNroRetasacion
End Function
'TORE ERS054-2017
Public Function CargaListaRetasaciones(ByVal pnTpoBusc As Integer, ByVal psNumBusc As String, _
                                        ByVal psDel As String, ByVal psHasta As String, _
                                        ByVal psCtaCad As String, ByVal pnTrimestre As Integer, _
                                        ByVal psAnio As String, ByVal pnEstado As Integer) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set CargaListaRetasaciones = oCredP.CargaListaRetasaciones(pnTpoBusc, psNumBusc, psDel, psHasta, psCtaCad, pnTrimestre, psAnio, pnEstado)
End Function
'TORE - ERS054-2017
Public Function ObtieneCredPignoRetasacion(ByVal psAgeCod As String) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set ObtieneCredPignoRetasacion = oCredP.ObtieneCredPignoRetasacion(psAgeCod)
End Function
'TORE - ERS054-2017
Public Function ObtenerDetRetasacion(ByVal psctacod As String, ByVal pnCodigoID As Integer, ByVal pnItem As Integer) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set ObtenerDetRetasacion = oCredP.ObtenerDetRetasacion(psctacod, pnCodigoID, pnItem)
End Function
'TORE ERS054-2017
Public Function ObtieneCodPrepaRetasacion(ByVal psNombreArhivo As String) As Integer
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Dim ODR As New Recordset
    Set ODR = oCredP.ObtieneCodPrepaRetasacion(psNombreArhivo)
    If Not (ODR.EOF And ODR.BOF) Then
        ObtieneCodPrepaRetasacion = ODR!nRetasado
    End If
End Function
'TORE ERS054-2017
'Public Sub RegistraRetasacion(ByVal pnCodigoID As Integer, ByVal psCtaCod As String, ByVal pnItem As Integer, ByVal psKilataje As String, _
'                              ByVal pnPiezas As Integer, ByVal pnPesoBruto As Double, ByVal pnPesoNeto As Double, ByVal pnValTasac As Double, _
'                              ByVal psNroHolograma As String, ByVal psObservaciones As String)
'    Dim oCredP As New COMDColocPig.DCOMColPContrato
'    oCredP.RegistraRetasacion pnCodigoID, psCtaCod, pnItem, psKilataje, pnPiezas, pnPesoBruto, pnPesoNeto, pnValTasac, psNroHolograma, psObservaciones
'End Sub
'TORE - Correcion
Public Function RegistraRetasacion(ByVal pnCodigoID As Long, ByVal psctacod As String, ByVal pnItem As Integer, ByVal psKilataje As String, _
                              ByVal pnPiezas As Integer, ByVal pnPesoBruto As Double, ByVal pnPesoNeto As Double, ByVal pnValTasac As Double, _
                              ByVal psNroHolograma As String, ByVal psObservaciones As String, ByVal pdFechaRetas As Date) As Integer
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    RegistraRetasacion = oCredP.RegistraRetasacion(pnCodigoID, psctacod, pnItem, psKilataje, pnPiezas, pnPesoBruto, pnPesoNeto, pnValTasac, psNroHolograma, psObservaciones, pdFechaRetas)
End Function

'TORE ERS054-2017
Public Function ObtieneMiembrosRetasacion(ByVal pnCodigoID As Integer) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set ObtieneMiembrosRetasacion = oCredP.ObtieneMiembrosRetasacion(pnCodigoID)
End Function
'TORE ERS054-2017
Public Sub RegistraMiembrosRetasacion(ByVal pnCodigoID As Integer, ByVal pnOrden As Integer, ByVal psNombreMiembro As String, ByVal psRolMiembro As String)
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    oCredP.RegistraMiembrosRetasacion pnCodigoID, pnOrden, psNombreMiembro, psRolMiembro
End Sub
'TORE ERS054-2017
Public Sub RegistraFechaRetasacion(ByVal pnCodigoID As Integer, ByVal psUser As String)
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    oCredP.RegistraFechaRetasacion pnCodigoID, psUser
End Sub
'[TORE ADD 01062019]
Public Function ObtenerCodRetasacion(ByVal pnCodigoID As Integer, ByVal psCredito As String) As String
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Dim ODR As New ADODB.Recordset
    Set ODR = oCredP.ObtenerCodRetasacion(pnCodigoID, psCredito)
    ObtenerCodRetasacion = ""
    If Not (ODR.EOF And ODR.BOF) Then
        ObtenerCodRetasacion = ODR!nNroRetasacion
    End If
End Function
Public Function ObtieneHistorialCredRetasacion(ByVal psctacod As String) As Boolean
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Dim ODR As New ADODB.Recordset
    Set ODR = oCredP.ObtieneHistorialCredRetasacion(psctacod)
    ObtieneHistorialCredRetasacion = False
    If Not (ODR.EOF And ODR.BOF) Then
        ObtieneHistorialCredRetasacion = True
    End If
End Function
'TORE ERS054-2017
Public Function ObtieneCtaCodRetasacionPrep(ByVal pnCodigoID As Integer) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set ObtieneCtaCodRetasacionPrep = oCredP.ObtieneCtaCodRetasacionPrep(pnCodigoID)
End Function
Public Sub FinalizaRetasacionPreparada(ByVal pnCodigoID As Integer, ByVal psFecRetas As String, ByVal psUserRetas As String)
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Call oCredP.FinalizaRetasacionPreparada(pnCodigoID, psFecRetas, psUserRetas)
End Sub
Public Function DevuelveHistorialRetasacion(ByVal psctacod As String) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set DevuelveHistorialRetasacion = oCredP.DevuelveHistorialRetasacion(psctacod)
End Function
'RECO FIN ***********************************************************************************************
'RECO20141204 ERS163-2014********************************************************************************
Public Sub nInsertaRegistroCreditoCampAdj(ByVal pnMovNro As Long, ByVal psCtaCodAdj As String, ByVal psctacod As String, ByVal pnMontoinicial As Double)
        
    Dim lsSQL As String
    Dim loDCONColPContrato As New COMDColocPig.DCOMColPContrato
    Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
    Dim lnMovNro As Long
    Dim lsOpeCod As String
    
    loRegPig.dBeginTrans
    mbTrans = True
    Call loDCONColPContrato.InsertaRegistroCreditoCampAdj(pnMovNro, psCtaCodAdj, psctacod, pnMontoinicial)
    loRegPig.dCommitTrans
    mbTrans = False
    Set loRegPig = Nothing
    
    Exit Sub
ErrorModPig:
        If mbTrans Then
            loRegPig.dRollbackTrans
            mbTrans = False
        End If
        Err.Raise vbObjectError + 100, "Error nInsertaRegistroCreditoCampAdj", "Error en Funcion de Registro de Contrato "
End Sub
Public Function DevuelveValorDesemCampAdjudicado(ByVal psctacod As String) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set DevuelveValorDesemCampAdjudicado = oCredP.DevuelveValorDesemCampAdjudicado(psctacod)
End Function
Public Function ObtieneValorPesoNetoJoya(ByVal psctacod As String) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set ObtieneValorPesoNetoJoya = oCredP.ObtieneValorPesoNetoJoya(psctacod)
End Function
'RECO FIN***********************************************************************************************
'RECO20140827 ERS032-2014********************************************************************************
Public Sub RegistroConfigTpoPrenda(ByVal psTipoDesc As String, ByVal pnPiezas As Integer)
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    oCredP.RegistroConfigTpoPrenda psTipoDesc, pnPiezas
End Sub
Public Sub ActualizaConfigTpoPrenda(ByVal pnTipoCod As Integer, ByVal psTipoDesc As String, ByVal pnPiezas As Integer)
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    oCredP.ActualizaConfigTpoPrenda pnTipoCod, psTipoDesc, pnPiezas
End Sub
Public Function QuitaConfigTipoPrenda(ByVal pnTipoCod As Integer) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    oCredP.QuitaConfigTipoPrenda pnTipoCod
End Function
Public Function ObtieneTipoPrenda() As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set ObtieneTipoPrenda = oCredP.ObtieneTipoPrenda
End Function
Public Sub RegistroConfigSubTpoPrenda(ByVal pnTipoCod As Integer, ByVal psSubTipoDesc As String)
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    oCredP.RegistroConfigSubTpoPrenda pnTipoCod, psSubTipoDesc
End Sub
Public Sub ActualizaConfigSubTpoPrenda(ByVal pnSubTipoCod As Integer, ByVal psSubTipoDesc As String)
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    oCredP.ActualizaConfigSubTpoPrenda pnSubTipoCod, psSubTipoDesc
End Sub
Public Function ObtieneSubTipoPrenda(ByVal pnTipoCod As Integer) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set ObtieneSubTipoPrenda = oCredP.ObtieneSubTipoPrendaXTipo(pnTipoCod)
End Function
Public Function QuitaConfigSubTipoPrenda(ByVal pnSubTipoCod As Integer) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    oCredP.QuitaConfigSubTipoPrenda pnSubTipoCod
End Function
Public Function TasaXCalificSBS() As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set TasaXCalificSBS = oCredP.TasaXCalificSBS
End Function
Public Sub ActualizaConfiguracionTasaXCalificSBS(ByVal pnCodigo As Integer, ByVal pnTasa As Integer)
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    oCredP.ActualizaConfiguracionTasaXCalificSBS pnCodigo, pnTasa
End Sub
Public Function ListaCriteriosEvaluacionClientePignoraticio(ByVal pnCalificacion As Integer) As ADODB.Recordset
    Dim oCredP As New DCOMColPContrato
    Set ListaCriteriosEvaluacionClientePignoraticio = oCredP.ListaCriteriosEvaluacionClientePignoraticio(pnCalificacion)
End Function
Public Sub ActualizaValoresCalificClientePigno(ByVal pnCalificacion As Integer, ByVal pnCriterio As Integer, ByVal pnValor As Integer, ByVal pnValorPrestTas As Integer)
    Dim oCredP As New DCOMColPContrato
    Call oCredP.ActualizaValoresCalificClientePigno(pnCalificacion, pnCriterio, pnValor, pnValorPrestTas)
End Sub
Public Function ListaTipoPrendaPigno() As ADODB.Recordset
    Dim oCredP As New DCOMColPContrato
    Set ListaTipoPrendaPigno = oCredP.ListaTipoPrendaPigno()
End Function
Public Function ListaSubTipoPrendaPigno(ByVal pnTipoCod As Integer) As ADODB.Recordset
    Dim oCredP As New DCOMColPContrato
    Set ListaSubTipoPrendaPigno = oCredP.ListaSubTipoPrendaPigno(pnTipoCod)
End Function
Public Function ObtieneValorCriterioCalificSBSPigno(ByVal psPersCod As String) As ADODB.Recordset
    Dim oCredP As New DCOMColPContrato
    Set ObtieneValorCriterioCalificSBSPigno = oCredP.ObtieneValorCriterioCalificSBSPigno(psPersCod)
End Function
Public Function ValorCriterioInactividadPigno(ByVal psPersCod As String, ByVal pnInactividad As Integer, ByVal psFecSis As String) As Integer
    Dim oCredP As New DCOMColPContrato
    Dim Rs As New ADODB.Recordset
    Set Rs = oCredP.ValorCriterioInactividadPigno(psPersCod, pnInactividad, psFecSis)
    If Not (Rs.EOF And Rs.BOF) Then
        ValorCriterioInactividadPigno = Rs!nCanInactiva
    End If
End Function
Public Function ObtieneDatosCalificClientePignoraticio() As ADODB.Recordset
    Dim oCredP As New DCOMColPContrato
    Set ObtieneDatosCalificClientePignoraticio = oCredP.ObtieneDatosCalificClientePignoraticio
End Function
Public Function ObtieneTasaPignoCalificacionSBS(ByVal pnCodigo As Integer) As Integer
    Dim oCredP As New DCOMColPContrato
    Dim Rs As New ADODB.Recordset
    Set Rs = oCredP.ObtieneTasaPignoCalificacionSBS(pnCodigo)
    If Not (Rs.BOF And Rs.EOF) Then
        ObtieneTasaPignoCalificacionSBS = Rs!nTasa
    End If
End Function
'RECO FIN************************************************************************************************
'RECO20160225 ERS040-2015******************************************************************
Public Function PignoDevuelveNroComprobanteAdj(ByVal psAgeCod As String) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set PignoDevuelveNroComprobanteAdj = oCredP.PignoDevuelveNroComprobanteAdj(psAgeCod)
End Function
Public Sub PignoRegistraNroCompAdj(ByVal psComprobanteNro As String, ByVal psNumero As String, ByVal psctacod As String, ByVal pnMovNro As Long)
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Call oCredP.PignoRegistraNroCompAdj(psComprobanteNro, psNumero, psctacod, pnMovNro)
End Sub
Public Function PignoReimpresionCompAdj(ByVal psctacod As String) As ADODB.Recordset
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Set PignoReimpresionCompAdj = oCredP.PignoReimpresionCompAdj(psctacod)
End Function
'RECO FIN**********************************************************************************

'NAGL ERS012-2017 20170710**************************************************************************
Public Function PignoDescripcionJoyaAdj(ByVal psctacod As String) As ADODB.Recordset
   Dim oCredP As New COMDColocPig.DCOMColPContrato
   Set PignoDescripcionJoyaAdj = oCredP.PignoDescripcionJoyaAdj(psctacod)
End Function

Public Sub RegistroVentaContratoPigno(ByVal psOpeTpo As String, ByVal pnDocTpo As Long, ByVal psDocNro As String, ByVal pdDocFecha As Date, ByVal psPersCod As String, _
                                      ByVal psctacod As String, ByVal psDescrip As String, ByVal pnVVenta As Currency, ByVal pnIGV As Currency, ByVal pnPVentas As Currency, ByVal pnMovNro As Long, ByVal pnMoneda As Integer, _
                                      Optional psDocNroRefe As String = "", Optional pnDocTpoRefe As Integer = 0, _
                                      Optional pnTipoCambio As Currency = 0, Optional psAreaAgecod As String = "")
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Call oCredP.RegistroVentaContratoPigno(psOpeTpo, pnDocTpo, psDocNro, pdDocFecha, psPersCod, psctacod, psDescrip, pnVVenta, pnIGV, pnPVentas, pnMovNro, pnMoneda, psDocNroRefe, pnDocTpoRefe, pnTipoCambio, psAreaAgecod)
End Sub
'NAGL ERS 012-2017 FIN**********************************************************************

'Inicio JOEP - ARLO ERS082-2017 - 20171213 ---AGREGADO DESDE LA 60
'JOEP
Public Function RegistroDatosSegmentacion(ByVal nTpFlex As Integer, ByVal pMatDatos As Variant, Optional ByVal pMatDatosExt As Variant = Nothing, Optional ByVal pcMovnro As String = "")
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    Dim i As Integer
    Dim j As Integer
    Dim x As Integer
    If nTpFlex = 1 Then
        If IsArray(pMatDatos) Then 'Pregunto Matriz no Sea Vacia
            If UBound(pMatDatos) Then
                For i = 1 To UBound(pMatDatos)
                    'Call oCredP.RegistraDatosSegmentacion(nTpFlex, pMatDatos(i, 1), pMatDatos(i, 2), pMatDatos(i, 3), pMatDatos(i, 4), pMatDatos(i, 5), pMatDatos(i, 6), pMatDatos(i, 7), pMatDatos(i, 8), pMatDatos(i, 9), pMatDatos(i, 10), pMatDatos(i, 11)) 'Agrego pMatDatos(i, 11) JOEP20180410 Mejora de Segmenatcion Pig
                    Call oCredP.RegistraDatosSegmentacion(nTpFlex, pMatDatos(i, 1), pMatDatos(i, 2), pMatDatos(i, 3), pMatDatos(i, 4), pMatDatos(i, 5), pMatDatos(i, 6), pMatDatos(i, 7), pMatDatos(i, 8), pMatDatos(i, 9), pMatDatos(i, 10), pMatDatos(i, 11), , pcMovnro) 'JOEP20210929 Prendario Externo
                Next i
            End If
            RegistroDatosSegmentacion = True
        Else
            RegistroDatosSegmentacion = False
        End If
    ElseIf nTpFlex = 2 Then
        If IsArray(pMatDatos) Then 'Pregunto Matriz no Sea Vacia
            If UBound(pMatDatos) Then
                For i = 1 To UBound(pMatDatos)
                    'Call oCredP.RegistraDatosSegmentacion(nTpFlex, pMatDatos(i, 1), pMatDatos(i, 2), pMatDatos(i, 3), pMatDatos(i, 4), pMatDatos(i, 5), 0, 0, 0, 0, 0, pMatDatos(i, 6))
                    Call oCredP.RegistraDatosSegmentacion(nTpFlex, pMatDatos(i, 1), pMatDatos(i, 2), pMatDatos(i, 3), pMatDatos(i, 4), pMatDatos(i, 5), 0, 0, 0, 0, 0, pMatDatos(i, 6), , pcMovnro) 'JOEP20210929 Prendario Externo
                Next i
            End If
            RegistroDatosSegmentacion = True
        Else
            RegistroDatosSegmentacion = False
        End If
    ElseIf nTpFlex = 3 Then
        'If IsArray(pMatDatos) Then 'Pregunto Matriz no Sea Vacia
        If IsArray(pMatDatos) And IsArray(pMatDatosExt) Then 'Pregunto Matriz no Sea Vacia
            'If UBound(pMatDatos) Then
            If UBound(pMatDatos) >= 1 And UBound(pMatDatosExt) >= 1 Then
                'Segementacion Interna
                For i = 1 To UBound(pMatDatos)
                    'Call oCredP.RegistraDatosSegmentacion(nTpFlex, pMatDatos(i, 1), pMatDatos(i, 2), "", pMatDatos(i, 3))
                    Call oCredP.RegistraDatosSegmentacion(nTpFlex, pMatDatos(i, 1), pMatDatos(i, 2), "", pMatDatos(i, 3), , , , , , , , 0, pcMovnro)
                Next i
                'Segementacion Externa
                x = 1
                For i = 1 To 9 'UBound(pMatDatosExt)
                    For j = 1 To 16 'UBound(pMatDatosExt)
                        Call oCredP.RegistraDatosSegmentacion(nTpFlex, i, "", pMatDatosExt(x + 2, 3), pMatDatosExt(x + 3, 4), , , , , , , , pMatDatosExt(x + 1, 2), pcMovnro)
                        x = x + 4
                    Next j
                Next i
            End If
            RegistroDatosSegmentacion = True
        Else
            RegistroDatosSegmentacion = False
        End If
    End If
End Function
'JOEP
'ARLO
'Comento JOEP20210514 Seg. Prendario Externo
'Public Sub RegistroFormatoEvalPig(ByVal pscCtaCod As String, ByVal pnIngreso As Double, ByVal pnEgreso As Double, ByVal psMovNro As String, ByVal pnTipCli As Integer, ByVal pnCantAdjuSeg As Integer, ByVal pnDiasSeg As Integer, ByVal pcSegmento As String, ByVal pnDiasSinAdj As Integer, ByVal pnSegmentoAnt As String) 'Agrego JOEP20180220 pnTipCli,pnCantAdjuSeg,pnDiasSeg,pcSegmento: joep pig pase pnDiasSinAdj
'    Dim oCredP As New COMDColocPig.DCOMColPContrato
'    oCredP.RegistroFormatoEvalPig pscCtaCod, pnIngreso, pnEgreso, psMovNro, pnTipCli, pnCantAdjuSeg, pnDiasSeg, pcSegmento, pnDiasSinAdj, pnSegmentoAnt 'Agrego JOEP20180220 pnTipCli,pnCantAdjuSeg,pnDiasSeg,pcSegmento: joep pig pase pnDiasSinAdj,pnSegmentoAnt
'End Sub
'Comento JOEP20210514 Seg. Prendario Externo

'Add JOEP20210514 Seg. Prendario Externo
Public Sub RegistroFormatoEvalPig(ByVal pscCtaCod As String, ByVal pnIngreso As Double, ByVal pnEgreso As Double, ByVal psMovNro As String, ByVal pnTipCli As Integer, ByVal pnCantAdjuSeg As Integer, ByVal pnDiasSeg As Integer, ByVal pcSegmento As String, ByVal pnDiasSinAdj As Integer, ByVal pnSegmentoAnt As String, Optional ByRef pArrayDatosSegPredExt As Variant = Nothing) 'Add JOEP20210513 Seg. Prendario Externo
    Dim oCredP As New COMDColocPig.DCOMColPContrato
    oCredP.RegistroFormatoEvalPig pscCtaCod, pnIngreso, pnEgreso, psMovNro, pnTipCli, pnCantAdjuSeg, pnDiasSeg, pcSegmento, pnDiasSinAdj, pnSegmentoAnt, pArrayDatosSegPredExt 'Add JOEP20210513 Seg. Prendario Externo
End Sub
'Add JOEP20210514 Seg. Prendario Externo

'ARLO
'Fin JOEP - ARLO ERS082-2017 - 20171213
'APRI20181109 MEJORA
Public Function ObtenerFechaEstado(ByVal psctacod As String, ByVal pnPrdEstado) As String
    Dim sql As String
    Dim Rs As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion

    sql = "EXEC stp_sel_ObtenerFechaEstadoCred '" & psctacod & "'," & pnPrdEstado
    
    Set Rs = oConec.CargaRecordSet(sql)
    
    If Not Rs.BOF And Not Rs.EOF Then
        ObtenerFechaEstado = Rs!dPrdEstado
    Else
        ObtenerFechaEstado = ""
    End If
    Set Rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'END APRI

'End jhcu 09-07-2020 - 09-07-2020
Public Sub ReversionReprogramacion(ByVal psCmovnro, ByVal psctacod)
    Dim lsSQL As String
    On Error GoTo dError
    Dim oConec As COMConecta.DCOMConecta
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion

    lsSQL = "EXEC stp_ins_ColocPigReversion '" + psCmovnro + "','" + psctacod + "'"
    oConec.CargaRecordSet (lsSQL)


    oConec.CierraConexion
    Set oConec = Nothing
    
    Exit Sub
dError:
    Err.Raise Err.Number, "Hubo un Error Comunicarse con TI-Desarrollo.", Err.Description
End Sub
'Add jhcu 09-07-2020 - 09-07-2020

'PEAC 20200904
Public Sub RegistroPignoSegmentacion(ByVal pscPersCod As String, ByVal psCmovnro As String)

    Dim lsSQL As String
    On Error GoTo dError
    Dim oConec As COMConecta.DCOMConecta
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion

    lsSQL = "EXEC stp_ins_RegistroPignoSegmentacion  '" + pscPersCod + "','" + psCmovnro + "' "
    oConec.CargaRecordSet (lsSQL)

    oConec.CierraConexion
    Set oConec = Nothing
    
    Exit Sub
dError:
    Err.Raise Err.Number, "Hubo un Error RegistroPignoSegmentacion.", Err.Description

End Sub
'CTI4 ERS0112020***
Public Function ObtieneNombreTitularCargoCta(ByVal psctacod As String) As String
    Dim sql As String

    Dim Rs As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion

    sql = "EXEC stp_sel_ERS0112020_ObtieneNombreTitularCargoCta '" & psctacod & "'"
    
    Set Rs = oConec.CargaRecordSet(sql)
    
    If Not Rs.BOF And Not Rs.EOF Then
        ObtieneNombreTitularCargoCta = Rs!cPersNombre
    Else
        ObtieneNombreTitularCargoCta = ""
    End If
    Set Rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'END CTI4
'JAOR 20200904
Public Sub AfectarPignoSegmentacion(ByVal pscCtaCod As String)

    Dim lsSQL As String
    On Error GoTo dError
    Dim oConec As COMConecta.DCOMConecta
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion

    lsSQL = "EXEC stp_upd_ColocPigAfectaCalificacion  '" + pscCtaCod + "'"
    oConec.CargaRecordSet (lsSQL)

    oConec.CierraConexion
    Set oConec = Nothing
    
    Exit Sub
dError:
    Err.Raise Err.Number, "Hubo un Error RegistroPignoSegmentacion.", Err.Description

End Sub

'JOEP20210927 Campana Prnedario
Public Function CampPrenRegTasas(ByVal prsDatos As Variant, ByVal pcMovnro As String) As Boolean
    Dim i As Integer
    Dim objCP As COMDColocPig.DCOMColPContrato
On Error GoTo dErrorCampPrenRegTasas
    Set objCP = New COMDColocPig.DCOMColPContrato
    CampPrenRegTasas = False
    If IsArray(prsDatos) Then
        If UBound(prsDatos) > 0 Then
            For i = 1 To UBound(prsDatos)
                Call objCP.CampPrendarioRegConf(prsDatos(i, 1), prsDatos(i, 2), prsDatos(i, 3), prsDatos(i, 4), prsDatos(i, 5), pcMovnro, i)
            Next i
            CampPrenRegTasas = True
        End If
    End If
    Set objCP = Nothing
   Exit Function
dErrorCampPrenRegTasas:
    CampPrenRegTasas = False
    Err.Raise Err.Number, "Hubo un Error Registrar la Configuracion de tasas Retencion.", Err.Description
End Function
'JOEP20210927 Campana Prnedario

