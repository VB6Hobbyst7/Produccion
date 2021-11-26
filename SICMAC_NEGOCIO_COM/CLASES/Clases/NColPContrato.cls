VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NColPContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mnEjecutaBatch As Integer
Dim mbTrans As Boolean

' Registra Contrato Pignoraticio
Public Function nRegistraContratoPignoraticio(ByVal psAgencia As String, ByVal pnMoneda As Moneda, _
                ByVal prPers As Recordset, ByVal pnTasa As Double, ByVal pnSaldo As Currency, _
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
Dim loGen As DGeneral
Dim lsCuenta As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

On Error GoTo ErrorRegPig

lsOpeCod = gColPOpeRegContrato

'Genera Nro Cuenta
Set loGen = New DGeneral
    lsCuenta = loGen.GeneraNuevaCuenta(psAgencia, gColConsuPrendario, pnMoneda)
Set loGen = Nothing


Set loRegPig = New DColPActualizaBD

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



Public Sub nModificaLoteCredPignoraticio(ByVal psCtacod As String, ByVal psFechaHora As String, _
            ByVal psLote As String, ByVal psMovNro As String, Optional pbEjecBatch As Boolean = False)


'** Actualiza Producto (Nro Transacc)
'** Actualiza Colocaciones (Ultima Modificacion)
'** Actualiza ColocPig (Lote)
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD

    loRegPig.dBeginTrans
    mbTrans = True
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , , , -2, False) ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Actualiza ColocPig
    Call loRegPig.dUpdateColocPignoraticio(psCtacod, , , , , , , , , , psLote, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, gColPOpeModifDescJoyas, "Modifica Lote", gMovEstContabNoContable, gMovFlagVigente, False)
    
    '** Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)

    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, gColPOpeModifDescJoyas, psCtacod, 0, 0, 0, "", 0, 0, 0, False)

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

Public Sub nAnulaCredPignoraticio(ByVal psCtacod As String, ByVal pnEstado As ColocEstado, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal pnMovNroAnt As Long, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'** Actualiza Mov ( MovNro del Registro-Anterior )
'** dInsertMovRef
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = gColPOpeAnula

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , pnEstado, psFechaHora, -2, False)  ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtacod, psFechaHora, gColPEstAnula, 1, 0, "Anula Contrato", _
        gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Anula Cred Pign", gMovEstContabNoContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 0, 0, 0, "", 0, 0, 0, False)
     
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

Public Sub nBloqueo_DesBloqueoCredPignoraticio(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pbBloqIni As Boolean, Optional ByVal psDescripcion As String = "@", _
        Optional ByVal psMovNroBloqueo As String = "@", Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertProductoBloqueos
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
' ******************************************************Falta cambiar
'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , , psFechaHora, -2, False)  ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    If pbBloqIni = False Then ' Se va ha Bloquear
        '** Insert ProductoBloqueos
        Call loRegPig.dInsertProductoBloqueos(psCtacod, 11, 3, psMovNro, psDescripcion, , False)
        '** Inserta Mov
        Call loRegPig.dInsertMov(psMovNro, geColPBloqueoJoyas, "Bloqueo de Joyas", gMovEstContabNoContable, gMovFlagVigente, False)
        ' Obtiene nMovNro
        lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        '** Inserta MovCol
        Call loRegPig.dInsertMovCol(lnMovNro, geColPBloqueoJoyas, psCtacod, 0, 0, 0, "", 0, 0, 0, False)
    Else ' Se va ha Desbloquear
        '** Update ProductoBloqueos
        Call loRegPig.dUpdateProductoBloqueos(psCtacod, psMovNroBloqueo, psMovNro, False)
        '** Inserta Mov
        Call loRegPig.dInsertMov(psMovNro, geColPDesbloqueoJoyas, "DesBloqueo de Joyas", gMovEstContabNoContable, gMovFlagVigente, False)
        ' Obtiene nMovNro
        lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        '** Inserta MovCol
        Call loRegPig.dInsertMovCol(lnMovNro, geColPDesbloqueoJoyas, psCtacod, 0, 0, 0, "", 0, 0, 0, False)
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


Public Sub nDesembolsoCredPignoraticio(ByVal psCtacod As String, ByVal pnSaldoCap As Currency, _
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
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

'Codigo de Operacion
lsOpeCod = gColPOpeDesembolsoEFE

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , gColPEstDesem, psFechaHora, -2, False)    ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtacod, psFechaHora, gColPEstDesem, 1, pnSaldoCap, "Desembolso Contrato", _
        gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Desembolso Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    
    '******** modificado 01/12/2004
    
    '** Inserta MovCol
    'Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 1, pnMontoEntregar, 0, "", 0, 0, pnSaldoCap, False)
    
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMontoDesembolsar, 0, "", 0, 0, pnSaldoCap, False)

    '** Inserta MovColDet -  Monto Entregar
    '****ORIGINAL ANTES DEL 07/02/2005
    
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodMontoEntregar, 1, pnMontoDesembolsar, False)
    
    
    '*****antes del 15/01/2005
    'Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodMontoEntregar, 1, pnMontoEntregar, False)
    
    'Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColPConceptoCodCapital, 1, pnMontoEntregar, False)
    
    '********
    

    '** Inserta MovColDet -  Interes Compensatorio
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodInteresCompensatorio, 1, pnInteresCompensat, False)
    
    '** Inserta MovColDet -  Impuesto
    If pnImpuesto > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodImpuesto, 1, pnImpuesto, False)
    End If
    '** Inserta MovColDet -  Costo Custodia
    If pnCostoCustodia > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCustodia, 1, pnCostoCustodia, False)
    End If
    '** Inserta MovColDet -  Costo Tasacion
    If pnCostoTasacion > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodTasacion, 1, pnCostoTasacion, False)
    End If
    
    If pbITFAplica Then
        If Not pbITFAsumidoCreditos Then
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, pnMontoITf1 + pnMontoITf2, 0, "", 0, 0, pnSaldoCap, False)
        
            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, gConcITFCliente, 1, pnMontoITf1, False)
            
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, 22, 1, pnMontoITf2, False)
        Else
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, pnMontoITf1 + pnMontoITf2, 0, "", 0, 0, pnSaldoCap, False)
        
            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, gConcITFAsumido, 1, pnMontoITf1, False)
            
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, 22, 1, pnMontoITf2, False)
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

Public Sub nRescataJoyaCredPignoraticio(ByVal psCtacod As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, ByVal pnKi21 As Double, _
        ByVal pnDiasCustodia As Integer, ByVal pnValTasac As Double, _
        ByVal psOperacion As String, Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

'Codigo de Operacion
'lsOpeCod = geColPDevJoyas
lsOpeCod = psOperacion

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , gColPEstCance, psFechaHora, -2, False)    ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtacod, psFechaHora, gColPEstCance, 1, 0, "Entrega Joya", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Entrega Joya Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, 0, 0, "", 0, 0, 0, False)
    '** Insert MovColDet Valor Tasacion
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodTasacion, 1, pnValTasac, False)

    '** Inserta MovColDet ' Oro 14/16/18/21
    If pnki14 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro14, 1, pnki14, False)
    End If
    If pnki16 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro16, 1, pnki16, False)
    End If
    If pnKi18 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro18, 1, pnKi18, False)
    End If
    If pnKi21 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro21, 1, pnKi21, False)
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

Public Sub nRetasacionCredPignoraticio(ByVal psCtacod As String, ByVal psMovNro As String, ByVal pnValTasacion As Currency, ByVal pnPiezas As Integer, _
                                        ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, ByVal pnKi21 As Double, _
                                         ByVal prjoyas As Recordset, Optional pnOroBruto As Double, Optional pnOroNeto As Double)
                                        
Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim loPigDetJoya As DColPContrato
Dim loCalcula As DColPCalculos
Dim lrJoyasAnt As ADODB.Recordset
Dim lnMovNro As Long, nNumRet As Integer
Dim lsOpeCod As String
Set loRegPig = New DColPActualizaBD
Set loPigDetJoya = New DColPContrato
Set loCalcula = New DColPCalculos
    
    lsOpeCod = "121300"
    nNumRet = loCalcula.nNumRetasacion(psCtacod)
    Set lrJoyasAnt = loPigDetJoya.dObtieneDatosCreditoPignoraticioJoyasDet(psCtacod)
    Call loRegPig.dInsertHistColocPigJoyaDetRet(lrJoyasAnt, nNumRet, psMovNro, psCtacod)
    Call loRegPig.dUpdateColocPignoraticio(psCtacod, pnOroBruto, pnOroNeto, pnPiezas, pnValTasacion)
    Call loRegPig.dUpdateColocPigJoya(psCtacod, pnki14, pnki16, pnKi18, pnKi21)
    Call loRegPig.dUpdateColocPigJoyaDet(prjoyas, psCtacod)
      '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Retasacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnValTasacion, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodTasacion, 1, pnValTasacion, False)
    Set loCalcula = Nothing
    Set lrJoyasAnt = Nothing
    Set loRegPig = Nothing
    Set loPigDetJoya = Nothing
    
End Sub

Public Sub nRenovacionCredPignoraticio(ByVal psCtacod As String, ByVal pnNewSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal psFecVenci As String, ByVal pnNewPlazo As Integer, _
        ByVal pnMontoTransac As Currency, ByVal pnCapPag As Currency, ByVal pnIntVenc As Currency, _
        ByVal pnCustodiaVenc As Currency, ByVal pnPreparaRemate As Currency, ByVal pnIntCompensat As Currency, _
        ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, ByVal pnDiasAtraso As Integer, _
        ByVal pnNroRenovacion As Integer, ByVal pnDiasCambCart As Integer, ByVal pnValortasacion As Currency, _
        ByVal pnOpeCod As Long, ByVal psOpeDesc As String, ByVal psPersCodCMAC As String, _
        ByVal pnGastoCorrespondencia As Currency, _
        ByVal pbITFAplica As Boolean, ByVal pbITFAsumidoCreditos As Boolean, _
        ByVal pnMontoITf As Currency, _
        Optional ByVal pbEjecBatch As Boolean = False)

'** Actualiza Producto (Estado, nSaldoCap,dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Plazo, Vencimiento, Ultima Modificacion, dVigencia)
'** Actualiza ColocPignoraticio (NroRenov)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
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
End Select

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , pnNewSaldoCap, gColPEstRenov, psFechaHora, -2, False)          ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, pnNewPlazo, psFecVenci, , , psMovNro, , False)

    '** Actualiza ColocPignoraticio
    Call loRegPig.dUpdateColocPignoraticio(psCtacod, , , , , pnNroRenovacion, 0, , , , , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtacod, psFechaHora, gColPEstRenov, 1, pnMontoTransac, "Renovacion Contrato", _
        gColocCalendCodPFCF, 0, 0, pnNewPlazo, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Renovacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMontoTransac, pnDiasAtraso, "GMIC", pnNewPlazo, gColPEstRenov, pnNewSaldoCap, False)

    '** Inserta MovColDet -  Capital
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCapital, 1, pnCapPag, False)

    '** Inserta MovColDet -  Interes Vencido
    If pnIntVenc > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodInteresMoratorio, 1, pnIntVenc, False)
    End If
    '** Inserta MovColDet -  Custod Venc
    If pnCustodiaVenc > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCustodiaVencida, 1, pnCustodiaVenc, False)
    End If
    '** Inserta MovColDet -  Preparac Remate
    If pnPreparaRemate > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodPreparaRemate, 1, pnPreparaRemate, False)
    End If
    '** Inserta MovColDet -  Interes Compensatorio (Adelantado)
    If pnIntCompensat > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodInteresCompensatorio, 1, pnIntCompensat, False)
    End If
    '** Inserta MovColDet -  Impuesto
    If pnImpuesto > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodImpuesto, 1, pnImpuesto, False)
    End If
    '** Inserta MovColDet -  Costo Custodia
    If pnCostoCustodia > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCustodia, 1, pnCostoCustodia, False)
    End If
    
    '** Inserta MovColDet -  Gasto Correspondencia
    If pnGastoCorrespondencia > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, 2215, 1, pnGastoCorrespondencia, False)
        
        'Actualiza ColocCalendDetPig
        lsSQL = "UPDATE ColocCalendDetPig Set nMontoPagado = nMontoPagado + " & pnGastoCorrespondencia & _
                " WHERE cCtaCod ='" & psCtacod & "' And nNroCalen = 1 And nColocCalendApl = 0 " & _
                " And nCuota = 1 And nPrdConceptoCod = 2215 "

        loRegPig.dEjecutar lsSQL
        
    End If
    
    If pbITFAplica Then
        If Not pbITFAsumidoCreditos Then
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, pnMontoITf, 0, "", 0, 0, 0, False)
        
            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, gConcITFCliente, 1, pnMontoITf, False)
        Else
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, pnMontoITf, 0, "", 0, 0, 0, False)
        
            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, gConcITFAsumido, 1, pnMontoITf, False)
        End If
    End If
    
    
    '**** Cambio de Cartera - Regresa a Normal
    If pnDiasAtraso > pnDiasCambCart Then
        'Cambio de Cartera - Regresa a Normal
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCambioCarteraMorosoNormal, 1, pnNewSaldoCap, False)
        'Cambio a Valor de Tasación
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCambioTasacionMorosoNormal, 1, pnValortasacion, False)
    End If
    
    '** Inserta MovCMAC - Para Recepcion de Llamada de cmac
    If pnOpeCod = gColPOpeRenovacEnOtCjEFE Or pnOpeCod = gColPOpeRenovacEnOtCjCHQ Then
        Call loRegPig.dInsertMovCMAC(lnMovNro, psPersCodCMAC, Format$(gTpoIFCmac, "00"), CInt(Mid(psCtacod, 9, 1)), "", _
                "", lsOpeCod, pnMontoTransac, False)
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

Public Sub nCancelacionCredPignoraticio(ByVal psCtacod As String, ByVal pnNewSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnMontoTransac As Currency, ByVal pnCapPag As Currency, ByVal pnIntVenc As Currency, _
        ByVal pnCustodiaVenc As Currency, ByVal pnPreparaRemate As Currency, _
        ByVal pnImpuesto As Currency, ByVal pnDiasAtraso As Integer, _
        ByVal pnDiasCambCart As Integer, ByVal pnValortasacion As Currency, _
        ByVal pnOpeCod As Long, ByVal psOpeDesc As String, ByVal psPersCodCMAC As String, _
        ByVal pbITFAplica As Boolean, ByVal pbITFAsumidoCreditos As Boolean, _
        ByVal pnMontoITf As Currency, _
        Optional pbEjecBatch As Boolean = False, Optional pnGastoCorrespondencia As Currency = 0)

'** Actualiza Producto (Estado, nSaldoCap,dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Plazo, Vencimiento, Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
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
End Select

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , pnNewSaldoCap, gColPEstDifer, psFechaHora, -2, False)           ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtacod, psFechaHora, gColPEstDifer, 1, pnMontoTransac, "Cancelacion Contrato", _
        gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Cancelacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMontoTransac, 0, "GMIC", 0, 0, 0, False)

    '** Inserta MovColDet -  Capital
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCapital, 1, pnCapPag, False)

    '** Inserta MovColDet -  Valor Tasacion
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodTasacion, 1, pnValortasacion, False)

    '** Inserta MovColDet -  Interes Vencido
    If pnIntVenc > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodInteresMoratorio, 1, pnIntVenc, False)
    End If
    '** Inserta MovColDet -  Custod Venc
    If pnCustodiaVenc > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCustodiaVencida, 1, pnCustodiaVenc, False)
    End If
    '** Inserta MovColDet -  Impuesto
    If pnImpuesto > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodImpuesto, 1, pnImpuesto, False)
    End If
    '** Inserta MovColDet -  Preparac Remate
    If pnPreparaRemate > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodPreparaRemate, 1, pnPreparaRemate, False)
    End If
    
        '** Inserta MovColDet -  Gasto Correspondencia
    If pnGastoCorrespondencia > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, 2215, 1, pnGastoCorrespondencia, False)
               
        'Actualiza ColocCalendDetPig
        lsSQL = "UPDATE ColocCalendDetPig Set nMontoPagado = nMontoPagado + " & pnGastoCorrespondencia & _
                " WHERE cCtaCod ='" & psCtacod & "' And nNroCalen = 1 And nColocCalendApl = 0 " & _
                " And nCuota = 1 And nPrdConceptoCod = 2215 "

        loRegPig.dEjecutar lsSQL
        
        
    End If
    
    '** Inserta MovCMAC - Para Recepcion de Llamada de cmac
    If pnOpeCod = gColPOpeCancelacEnOtCjEFE Or pnOpeCod = gColPOpeCancelacEnOtCjCHQ Then
        Call loRegPig.dInsertMovCMAC(lnMovNro, psPersCodCMAC, Format$(gTpoIFCmac, "00"), CInt(Mid(psCtacod, 9, 1)), "", _
                "", lsOpeCod, pnMontoTransac, False)
    End If
    
    '+++++++++++++   itf  +++++++++++++++++
    If pbITFAplica Then
        If Not pbITFAsumidoCreditos Then
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, pnMontoITf, 0, "", 0, 0, 0, False)
            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, gConcITFCliente, 1, pnMontoITf, False)
        Else
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, pnMontoITf, 0, "", 0, 0, 0, False)
        
            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, gConcITFAsumido, 1, pnMontoITf, False)
        End If
    End If
    '++++++++++++++++++++++++++++++++++++++
    
    
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

Public Sub nCustodiaDiferidaCredPignoraticio(ByVal psCtacod As String, _
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
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = gColPOpeCobCusDiferida

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Custodia Diferida Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet -  Custodia Neto
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCustodiaDiferida, 1, pnCustodNeto, False)

    '** Inserta MovColDet -  Custodia Impuesto
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCustodiaDiferidaImpuesto, 1, pnCustodImpuesto, False)

    '** Inserta MovDoc -  Registra la boleta
    Call loRegPig.dInsertMovDoc(lnMovNro, 3, psNroDoc, psFechaHora, False)
     
     '** ITF
    Call loRegPig.dInsertMovCol(lnMovNro, "990104", psCtacod, 1, pnITF, 0, "", 0, 0, 0, False)
    Call loRegPig.dInsertMovColDet(lnMovNro, "990104", psCtacod, 1, gConcITFCliente, 1, pnITF, False)
     
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



Public Sub nDuplicadoContratoCredPignoraticio(ByVal psCtacod As String, ByVal pnNroDuplic As Currency, _
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
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = gColPOpeImpDuplicado

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , , psFechaHora, -2, False)           ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Actualiza ColocPignoraticio
    Call loRegPig.dUpdateColocPignoraticio(psCtacod, , , , , , , pnNroDuplic + 1, , , , False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Duplicado Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet -  Costo Duplicado
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCostoDuplicado, 1, pnMontoTransac, False)
    
    '**ITF
    Call loRegPig.dInsertMovCol(lnMovNro, "990104", psCtacod, 1, pnITF, 0, "", 0, 0, 0, False)
    Call loRegPig.dInsertMovColDet(lnMovNro, "990104", psCtacod, 1, gConcITFCliente, 1, pnITF, False)
    
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
Public Sub nExtornoDesembolsoCredPig(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = geColPExtDesemb

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , gColPEstRegis, psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtacod, psFechaHora, gColPEstRegis, 1, pnMonto, "Extorno Desembolso Cont", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Extorno Desembolso Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
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

Public Sub nExtornoDuplicadoContratoCredPig(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = geColPExtDupli

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , , psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Duplicado Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
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


Public Sub nExtornoRenovacionCredPig(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long, lnNroRenov As Integer
Dim lsOpeCod As String

Dim loDataExt As dColPFunciones
Dim lrDataExt As ADODB.Recordset

'On Error GoTo ErrorModPig

lsOpeCod = geColPExtRenov

' *** Obtiene Datos para el Extorno

lsSQL = "SELECT CE.nPlazo, CE.cCtaCod, " _
      & " nSaldoCap = (SELECT nSaldo FROM Producto WHERE cCtaCod = '" & psCtacod & "'  ) , " _
      & " nCapitalPag = (SELECT ISNULL(SUM(nMonto),0) FROM MovColDet WHERE nMovNro = " & pnMovNroAnt _
      & "               AND nPrdConceptoCod IN (" & gColPConceptoCodCapital & "," & gColPConceptoCodCapitalVencido & ") ) ," _
      & " nGastoCorrespond = (SELECT ISNULL(SUM(nMonto),0) FROM MovColDet WHERE nMovNro = " & pnMovNroAnt _
      & "               AND nPrdConceptoCod IN (2215)  ) ," _
      & " dVenc = (SELECT dVenc FROM Colocaciones WHERE cCtaCod = '" & psCtacod & "'  ) , " _
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
      & " WHERE CE.cCtaCod = '" & psCtacod & "' AND CE.nPrdEstado = " & gColPEstRenov & "  " _
      & " AND CE.dPrdEstado = (SELECT MAX(dPrdEstado) FROM ColocacEstado " _
      & "                     WHERE cCtaCod = CE.cCtaCod AND nPrdEstado = CE.nPrdEstado ) "

Set loDataExt = New dColPFunciones
    Set lrDataExt = loDataExt.dObtieneRecordSet(lsSQL)
Set loDataExt = Nothing
If lrDataExt Is Nothing Then
    MsgBox "ERROR: al Buscar datos para Extorno ", vbInformation, "Aviso"
    Exit Sub
End If
If lrDataExt.BOF And lrDataExt.EOF Then
    MsgBox "ERROR: No encontro datos de la Operacion a extornar ", vbInformation, "Aviso"
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
lsAntFecVenc = Format$(DateAdd("d", -1 * (lrDataExt!nDiasAtraso + lrDataExt!nPlazo), lrDataExt!dVenc), "mm/dd/yyyy")
lnNroRenov = lrDataExt!nNroRenov
' *** Realiza el Extorno

Set loRegPig = New DColPActualizaBD
    
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , lnAntSaldoCap, lnAntEstado, psFechaHora, -2, False)       ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, lnAntPlazo, lsAntFecVenc, , , psMovNro, , False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Renovacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMonto, 0, "", 0, 0, lnAntSaldoCap, False)
     
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
    
    '** Actualiza NroRenovaciones
    Call loRegPig.dUpdateColocPignoraticio(psCtacod, , , , , lnNroRenov - 1)

    '*** Gasto de Correspondencia
    If lrDataExt!nGastoCorrespond > 0 Then
        'Actualiza ColocCalendDetPig
        lsSQL = "UPDATE ColocCalendDetPig Set nMontoPagado = nMontoPagado - " & lrDataExt!nGastoCorrespond & _
                " WHERE cCtaCod ='" & psCtacod & "' And nNroCalen = 1 And nColocCalendApl = 0 " & _
                " And nCuota = 1 And nPrdConceptoCod = 2215 "

        loRegPig.dEjecutar lsSQL
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


Public Sub nExtornoCancelacionCredPig(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

Dim loDataExt As dColPFunciones
Dim lrDataExt As ADODB.Recordset

'On Error GoTo ErrorModPig

lsOpeCod = geColPExtRenov

' *** Obtiene Datos para el Extorno

lsSQL = "SELECT CE.nPlazo, CE.cCtaCod, " _
      & " nSaldoCap = (SELECT nSaldo FROM Producto WHERE cCtaCod = '" & psCtacod & "'  ) , " _
      & " nCapitalPag = (SELECT ISNULL(SUM(nMonto),0) FROM MovColDet WHERE nMovNro = " & pnMovNroAnt _
      & "               AND nPrdConceptoCod IN (" & gColPConceptoCodCapital & "," & gColPConceptoCodCapitalVencido & ") ) ," _
      & " dVenc = (SELECT dVenc FROM Colocaciones WHERE cCtaCod = '" & psCtacod & "'  ) , " _
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
      & "FROM ColocacEstado CE WHERE CE.cCtaCod = '" & psCtacod & "' AND CE.nPrdEstado = " & gColPEstDifer & "  " _
      & "AND CE.dPrdEstado = (SELECT MAX(dPrdEstado) FROM ColocacEstado " _
      & "                     WHERE cCtaCod = CE.cCtaCod AND nPrdEstado = CE.nPrdEstado ) "

Set loDataExt = New dColPFunciones
    Set lrDataExt = loDataExt.dObtieneRecordSet(lsSQL)
Set loDataExt = Nothing
If lrDataExt Is Nothing Then
    MsgBox "ERROR: al Buscar datos para Extorno ", vbInformation, "Aviso"
    Exit Sub
End If
If lrDataExt.BOF And lrDataExt.EOF Then
    MsgBox "ERROR: No encontro datos de la Operacion a extornar ", vbInformation, "Aviso"
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

Set loRegPig = New DColPActualizaBD
    
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , lnAntSaldoCap, lnAntEstado, psFechaHora, -2, False)       ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtacod, psFechaHora, lnAntEstado, 1, pnMonto, "Extorno Cancelacion Cred Pig", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Cancelacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMonto, 0, "", 0, 0, lnAntSaldoCap, False)
    
    '** Update Mov Anterior
    Call loRegPig.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loRegPig.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

    '*** Gasto de Correspondencia
    If lrDataExt!nGastoCorrespond > 0 Then
        'Actualiza ColocCalendDetPig
        lsSQL = "UPDATE ColocCalendDetPig Set nMontoPagado = nMontoPagado - " & lrDataExt!nGastoCorrespond & _
                " WHERE cCtaCod ='" & psCtacod & "' And nNroCalen = 1 And nColocCalendApl = 0 " & _
                " And nCuota = 1 And nPrdConceptoCod = 2215 "

        loRegPig.dEjecutar lsSQL
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

Public Sub nExtornoAmortizacionCredPig(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        ByVal psOpeExt, Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

Dim loDataExt As dColPFunciones
Dim lrDataExt As ADODB.Recordset

'On Error GoTo ErrorModPig

lsOpeCod = gColPOpeExtRenov

' *** Obtiene Datos para el Extorno

lsSQL = "SELECT CE.nPlazo, CE.cCtaCod, " _
      & " nSaldoCap = (SELECT nSaldo FROM Producto WHERE cCtaCod = '" & psCtacod & "'  ) , " _
      & " nCapitalPag = (SELECT ISNULL(SUM(nMonto),0) FROM MovColDet WHERE nMovNro = " & pnMovNroAnt _
      & "               AND nPrdConceptoCod IN (" & gColPConceptoCodCapital & "," & gColPConceptoCodCapitalVencido & ") ) ," _
      & " dVenc = (SELECT dVenc FROM Colocaciones WHERE cCtaCod = '" & psCtacod & "'  ) , " _
      & " nDiasAtraso = (SELECT nDiasMora FROM MovCol where nMovNro = " & pnMovNroAnt & " and copecod='" & psOpeExt & "'  ) , " _
      & " nPrdEstadoAnt = (SELECT nEstAnt = CASE WHEN nDiasMora <= 0 AND nNroRenov = 0 THEN " & gColPEstDesem & "  " _
      & "                                        WHEN nDiasMora > 0 THEN " & gColPEstVenci & " " _
      & "                                        WHEN nDiasMora > 30 THEN " & gColPEstPRema & " " _
      & "                                        WHEN nDiasMora <= 0 AND nNroRenov > 0 then  " & gColPEstRenov & " " _
      & "                                   END " _
      & "                  FROM MovCol MovCol INNER JOIN ColocPignoraticio ColPig " _
      & "                  ON MovCol.cCtaCod = ColPig.cCtaCod WHERE nMovNro = " & pnMovNroAnt & " and copecod='" & psOpeExt & "'   ) " _
      & "FROM ColocacEstado CE WHERE CE.cCtaCod = '" & psCtacod & "' AND CE.nPrdEstado = " & gColPEstRenov & "  " _
      & "AND CE.dPrdEstado = (SELECT MAX(dPrdEstado) FROM ColocacEstado " _
      & "                     WHERE cCtaCod = CE.cCtaCod AND nPrdEstado = CE.nPrdEstado ) "

Set loDataExt = New dColPFunciones
    Set lrDataExt = loDataExt.dObtieneRecordSet(lsSQL)
Set loDataExt = Nothing
If lrDataExt Is Nothing Then
    MsgBox "ERROR: al Buscar datos para Extorno ", vbInformation, "Aviso"
    Exit Sub
End If
If lrDataExt.BOF And lrDataExt.EOF Then
    MsgBox "ERROR: No encontro datos de la Operacion a extornar ", vbInformation, "Aviso"
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

Set loRegPig = New DColPActualizaBD
    
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , lnAntSaldoCap, lnAntEstado, psFechaHora, -2, False)       ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Amortizacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMonto, 0, "", 0, 0, lnAntSaldoCap, False)
     
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

Public Sub nExtornoRescateJoyaCredPig(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion, dVigencia)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = geColPExtDevJoyas

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , gColPEstDifer, psFechaHora, -2, False)       ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtacod, psFechaHora, gColPEstDifer, 1, 0, "Extorno Entrega Joya", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Rescate Joyas Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
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

Public Sub nExtornoCustodiaDiferidaCredPig(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto ( NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = geColPExtCustodDifer

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , , psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Custodia Diferida Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
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

Public Sub nExtornoRemateVentaCredPig(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** Actualiza ColocPigRGDet (Estado, Monto Venta, nMovVta)
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsNroProceso As String

lsOpeCod = "129701"
'Obtiene el nro de proceso
Dim loDat As NColPRecGar
Set loDat = New NColPRecGar
    lsNroProceso = loDat.nObtieneRemateIniciado()
Set loDat = Nothing

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , gColPEstPRema, psFechaHora, -2, False)        ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("R", lsNroProceso, psCtacod, gColPRecGarVentaNoVendido, 0, , , , 0, False)
    
    '** Elimina ColocPigRecupVta
    Call loRegPig.dDeleteColocPigRecupVta("R", lsNroProceso, psCtacod, False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Venta Remate ", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
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

Public Sub nExtornoSubastaVentaCredPig(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsNroProceso As String

lsOpeCod = "129703"
'Obtiene el nro de proceso
Dim loDat As NColPRecGar
Set loDat = New NColPRecGar
    lsNroProceso = loDat.nObtieneRemateIniciado()
Set loDat = Nothing

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , gColPEstAdjud, psFechaHora, -2, False)        ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("A", lsNroProceso, psCtacod, gColPRecGarAdjudicado, 0, , , , 0, False)
    
    '** Elimina ColocPigRecupVta
    'Call loRegPig.dDeleteColocPigRecupVta("R", lsNroProceso, psCtacod, False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Venta Subasta ", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
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


Public Sub nExtornoRecuperacionCredPig(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsNroProceso As String

lsOpeCod = "129704"
'Obtiene el nro de proceso
Dim loDat As NColPRecGar
Set loDat = New NColPRecGar
    lsNroProceso = loDat.nObtieneRemateIniciado()
Set loDat = Nothing

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , gColPEstAdjud, psFechaHora, -2, False)        ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Update ColocPigRGDet
    'Call loRegPig.dUpdateColocPigRGDet("R", lsNroProceso, psCtacod, gColPRecGarVentaNoVendido, 0, , , , 0, False)
    
    '** Elimina ColocPigRecupVta
   ' Call loRegPig.dDeleteColocPigRecupVta("R", lsNroProceso, psCtacod, False)

    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Ext Recup. Subasta ", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMonto, 0, "", 0, 0, 0, False)
    
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


Public Sub nRemateVentaCredPignoraticio(ByVal psCtacod As String, ByVal psNroProceso As String, _
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
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = gColPOpeVtaRemate

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , gColPEstRemat, psFechaHora, -2, False)           ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtacod, psFechaHora, gColPEstRenov, 1, pnMontoTransac, "Venta Joyas en Remate", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Venta Remate", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("R", psNroProceso, psCtacod, gColPRecGarVentaVendido, pnMontoTransac, , , , lnMovNro, False)
    
    '** Inserta ColocPigRecupVta
    Call loRegPig.dInsertColocPigRecupVta("R", psNroProceso, psCtacod, 23, psNroDoc, psPersCodComp, pnMontoTransac, pnComisionRemate, pnSobranteRemate, False)

    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet -  Capital
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCapital, 1, pnCapital, False)

    '** Inserta MovColDet -  Interes Vencido
    If pnIntVenc > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodInteresMoratorio, 1, pnIntVenc, False)
    End If
    '** Inserta MovColDet -  Custod Venc
    If pnCustodiaVenc > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCustodiaVencida, 1, pnCustodiaVenc, False)
    End If
    '** Inserta MovColDet -  Preparac Remate
    If pnPreparaRemate > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodPreparaRemate, 1, pnPreparaRemate, False)
    End If
    '** Inserta MovColDet -  Costo Remate - Comision
    If pnComisionRemate > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodComisionRemate, 1, pnComisionRemate, False)
    End If
    '** Inserta MovColDet -  Impuesto
    If pnImpuesto > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodImpuesto, 1, pnImpuesto, False)
    End If
    '** Inserta MovColDet -  Sobrante Remate
    If pnSobranteRemate > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodSobranteRemate, 1, pnSobranteRemate, False)
    End If
    
    '** Inserta MovColDet ' Valor Tasacion
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodTasacion, 1, pnValTasacion, False)
    

    '** Inserta MovColDet -  IGV de Boleta
    If pnIGVBoleta > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, 2216, 1, pnIGVBoleta, False)
    End If

    '** Inserta MovColDet -  Subtotal
    If pnSubtotalBoleta > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, 2217, 1, pnSubtotalBoleta, False)
    End If


    '** Inserta MovColDet ' Oro 14/16/18/21
    If pnki14 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro14, 1, pnki14, False)
    End If
    If pnki16 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro16, 1, pnki16, False)
    End If
    If pnKi18 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro18, 1, pnKi18, False)
    End If
    If pnKi21 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro21, 1, pnKi21, False)
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

Public Sub nSubastaVentaCredPignoraticio(ByVal psCtacod As String, ByVal psNroProceso As String, _
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
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = geColPVtaSubasta

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , gColPEstSubas, psFechaHora, -2, False)            ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtacod, psFechaHora, gColPEstSubas, 1, pnMontoTransac, "Venta Joyas en Subasta", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Venta Joyas en Subasta", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("S", psNroProceso, psCtacod, gColPRecGarVentaVendido, pnMontoTransac, , , , lnMovNro, False)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet -  Venta Neta
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, geColPConceptoCodVtaNetaSubasta, 1, pnVtaNetaSubasta, False)

    '** Inserta MovColDet -  Impuesto Vta Subasta
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, geColPConceptoCodImpuestoVtaSubasta, 1, pnImpuestoVtaSubasta, False)
    
    '** Inserta MovColDet -  Valor Adjudicacion
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, geColPConceptoCodValorAdjudica, 1, pnValorAdjudicacion, False)

    '** Inserta MovColDet ' Oro 14/16/18/21
    If pnki14 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro14, 1, pnki14, False)
    End If
    If pnki16 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro16, 1, pnki16, False)
    End If
    If pnKi18 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro18, 1, pnKi18, False)
    End If
    If pnKi21 > 0 Then
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro21, 1, pnKi21, False)
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

Public Sub nOpeCMACLlamadaCredPignoraticio(ByVal pnOpeCod As Long, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal psGlosa As String, ByVal psPersCodCMAC As String, _
        ByVal psIFTipo As String, ByVal pnMoneda As Moneda, ByVal psCtaIf As String, _
        ByVal psDocumento As String, ByVal pnMontoTransac As Currency, _
        Optional pbEjecBatch As Boolean = False, Optional nComision As Double = 0, _
        Optional nMonedaComision As Moneda = gMonedaNacional, Optional sCuentaAho As String = "", _
        Optional sNomAge As String = "", Optional sLpt As String = "", Optional sCliente As String = "", Optional psPersLavDinero As String = "")

'** dInsertMov
'** dInsertMovCMAC
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lnMovNroRef As Long
Dim lsOpeCod As String, sCodUser As String, sCodage As String
Dim oCont As NContFunciones
Dim dFecSis As Date

dFecSis = CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4))
sCodUser = Right(psMovNro, 4)
sCodage = Mid(psMovNro, 18, 2)

'Codigo de Operacion
lsOpeCod = pnOpeCod

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, psGlosa, gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCMAC
    Call loRegPig.dInsertMovCMAC(lnMovNro, psPersCodCMAC, psIFTipo, pnMoneda, psCtaIf, _
            psDocumento, pnOpeCod, pnMontoTransac, False, sCliente)
    
    ' xxx loRegPig.dCommitTrans
    ' xxx mbTrans = False
   ' Set loRegPig = Nothing
    

    If nComision > 0 Then
        Dim sMovNroCom As String
        Dim nMovNroCom As Long
        Set oCont = New NContFunciones
        
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
        Dim oCapMov As NCapMovimientos
        Dim sMovNroReg As String
        Set oCont = New NContFunciones
        'sMovNroReg = oCont.GeneraMovNro(dFecSis, sCodAge, sCodUser)
        
        'xxx agregado
        If Len(Trim(sMovNroCom)) > 0 Then
            sMovNroReg = oCont.IncrementaMovNro(sMovNroCom)
        Else
            sMovNroReg = oCont.IncrementaMovNro(psMovNro)
        End If
        
        Set oCont = Nothing
        
        Set oCapMov = New NCapMovimientos
        
        oCapMov.CapAbonoCuentaAho sCuentaAho, pnMontoTransac, gAhoDepRegCMACLlam, sMovNroReg, psGlosa, , , , , , , , , sNomAge, sLpt, psPersLavDinero, True, , , , , loRegPig.coConex.ConexionActiva
        
        Set oCapMov = Nothing
        
        'Pepe
        lnMovNroRef = loRegPig.dGetnMovNro(sMovNroReg)
        Call loRegPig.dInsertMovRef(lnMovNro, lnMovNroRef)
        'Fin Pepe
         
    End If

    'Lavado de Dinero

    'Dim oBase As DCredActualizaBD
    'Set oBase = New DCredActualizaBD
    
    ' Lavado de Dinero
    'If psPersLavDinero <> "" Then
    '    loRegPig.InsertaMovLavDinero lnMovNro, psPersLavDinero
    'End If

    loRegPig.coConex.CommitTrans
    
    Set loRegPig = Nothing

    'Set oBase = Nothing
    
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
Dim loRegPig As DColPActualizaBD
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
Set loRegPig = New DColPActualizaBD
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
                ByVal prPers As Recordset, ByVal pnTasa As Double, ByVal pnSaldo As Currency, _
                ByVal psFechaHora As String, ByVal pnPlazo As Integer, ByVal psFecVenc As String, _
                ByVal pnOroBruto As Currency, ByVal pnOroNeto As Currency, ByVal pnValTasacion As Currency, _
                ByVal pnPiezas As Integer, ByVal psTipoContrato As String, ByVal psLote As String, _
                ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, ByVal pnKi21 As Double, _
                ByVal psMovNro As String, ByVal pnIntAdelantado As Currency, ByVal pnCostoTasac As Currency, _
                ByVal pnCostoCustodia As Currency, ByVal pnImpuesto As Currency, ByVal prjoyas As Recordset) As String
'** Genera Nuevo Nro Contrato
'** dInsertCreditoPignoraticio
'** dInsertColocEstado
'** dInsertCreditoPigJoya
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************
Dim loGen As DGeneral
Dim lsCuenta As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

On Error GoTo ErrorRegPig

lsOpeCod = gColPOpeRegContrato

'Genera Nro Cuenta
Set loGen = New DGeneral
    lsCuenta = loGen.GeneraNuevaCuenta(psAgencia, gColConsuPrendario, pnMoneda)
Set loGen = Nothing


Set loRegPig = New DColPActualizaBD

    loRegPig.dBeginTrans
    mbTrans = True

    Call loRegPig.dInsertCredPignoraticioDetalle(lsCuenta, pnTasa, pnSaldo, psFechaHora, pnPlazo, _
        psFecVenc, pnOroBruto, pnOroNeto, pnValTasacion, pnPiezas, psTipoContrato, psLote, _
        pnki14, pnki16, pnKi18, pnKi21, psMovNro, prjoyas, False)
        
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

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
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

Public Sub nAmortizacionCredPignoraticio(ByVal psCtacod As String, ByVal pnNewSaldoCap As Currency, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal psFecVenci As String, ByVal pnNewPlazo As Integer, _
        ByVal pnMontoTransac As Currency, ByVal pnCapPag As Currency, ByVal pnIntVenc As Currency, _
        ByVal pnCustodiaVenc As Currency, ByVal pnPreparaRemate As Currency, ByVal pnIntCompensat As Currency, _
        ByVal pnImpuesto As Currency, ByVal pnCostoCustodia As Currency, ByVal pnDiasAtraso As Integer, _
        ByVal pnDiasCambCart As Integer, ByVal pnValortasacion As Currency, _
        ByVal pnOpeCod As Long, ByVal psOpeDesc As String, ByVal psPersCodCMAC As String, _
        ByVal pbITFAplica As Boolean, ByVal pbITFAsumidoCreditos As Boolean, _
        ByVal pnMontoITf As Currency, _
        Optional ByVal pbEjecBatch As Boolean = False)

'** Actualiza Producto (Estado, nSaldoCap,dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Plazo, Vencimiento, Ultima Modificacion, dVigencia)
'** Actualiza ColocPignoraticio (NroRenov)
'** dInsertColocacEstado
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
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
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , pnNewSaldoCap, gColPEstRenov, psFechaHora, -2, False)          ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Actualiza ColocPignoraticio
    'Call loRegPig.dUpdateColocPignoraticio(psCtaCod, , , , , pnNroRenovacion, , , , , , False)

    '** Insert ColocEstado
    Call loRegPig.dInsertColocacEstado(psCtacod, psFechaHora, gColPEstRenov, 1, pnMontoTransac, "Amortizacion Contrato", _
        gColocCalendCodPFCF, 0, 0, pnNewPlazo, "0", 0, 0, 0, 0, False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Amortizacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMontoTransac, 0, "GMIC", pnNewPlazo, 0, pnNewSaldoCap, False)

    '** Inserta MovColDet -  Capital
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCapital, 1, pnCapPag, False)
    
    If pbITFAplica Then
        If Not pbITFAsumidoCreditos Then
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, pnMontoITf, 0, "", 0, 0, 0, False)
        
            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, gConcITFCliente, 1, pnMontoITf, False)
        Else
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, pnMontoITf, 0, "", 0, 0, 0, False)
        
            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, gConcITFAsumido, 1, pnMontoITf, False)
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
    
'    '** Inserta MovCMAC - Para Recepcion de Llamada de cmac
'    If pnOpeCod = gColPOpeRenovacEnOtCjEFE Or pnOpeCod = gColPOpeRenovacEnOtCjCHQ Then
'        Call loRegPig.dInsertMovCMAC(lnMovNro, psPersCodCMAC, Format$(gTpoIFCmac, "00"), CInt(Mid(psCtaCod, 9, 1)), "", _
'                "", lsOpeCod, pnMontoTransac, False)
'    End If
    
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


Public Sub nSubastaVentaCredPignoraticioSINSubasta(ByVal psCtacod As String, ByVal psNroProceso As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnMontoTransac As Currency, ByVal pnVtaNetaSubasta As Currency, _
        ByVal pnImpuestoVtaSubasta As Currency, ByVal pnValorAdjudicacion As Currency, _
        ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, ByVal pnKi21 As Double, _
        Optional pbEjecBatch As Boolean = False, Optional psNroDoc As String = "", Optional psPersCod As String = "", _
        Optional ByVal pbITFAplica As Boolean = False, Optional ByVal pbITFAsumidoCreditos As Boolean = False, _
        Optional ByVal pnMontoITf As Currency = 0, Optional ByVal pnRecupInteres As Currency = 0, _
        Optional pbRecuperacion As Boolean = False)

'** Actualiza Producto (Estado, dPrdEstado, NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertColocacEstado
'** dInsertMov  psPersCod
'** Actualiza ColocPigRGDet (Estado, Monto Venta, nMovVta)
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
If pbRecuperacion Then
   lsOpeCod = "122900"
Else
   lsOpeCod = geColPVtaSubasta
End If

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
        
    If pbRecuperacion = False Then
        '** Actualiza Producto
        Call loRegPig.dUpdateProducto(psCtacod, , , gColPEstSubas, psFechaHora, -2, False)             ' (-2) aumenta el ste transac
        '** Actualiza Colocaciones
        Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)
        '** Insert ColocEstado
        Call loRegPig.dInsertColocacEstado(psCtacod, psFechaHora, gColPEstSubas, 1, pnMontoTransac, "Venta Joyas en Subasta", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
        '** Inserta Mov
    
        Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Venta Joyas Adjudicadas", gMovEstContabMovContable, gMovFlagVigente, False)
        ' Obtiene nMovNro
        lnMovNro = loRegPig.dGetnMovNro(psMovNro)
                
        '** Inserta MovCol
        Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)
    
        '** Inserta MovColDet -  Venta Neta
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, geColPConceptoCodVtaNetaSubasta, 1, pnVtaNetaSubasta, False)
    
        '** Inserta MovColDet -  Impuesto Vta Subasta
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, geColPConceptoCodImpuestoVtaSubasta, 1, pnImpuestoVtaSubasta, False)
        
        '** Inserta MovColDet -  Valor Adjudicacion
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, geColPConceptoCodValorAdjudica, 1, pnValorAdjudicacion, False)
    
        '** Inserta MovColDet ' Oro 14/16/18/21
        If pnki14 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro14, 1, pnki14, False)
        End If
        If pnki16 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro16, 1, pnki16, False)
        End If
        If pnKi18 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro18, 1, pnKi18, False)
        End If
        If pnKi21 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro21, 1, pnKi21, False)
        End If
        
        ''** Inserta Boleta
        If psNroDoc <> "" Then
             Call loRegPig.dInsertMovDoc(lnMovNro, TpoDoc.TpoDocBoletaVenta, psNroDoc, psFechaHora)
        End If
    
        '** Update ColocPigRGDet
        'Call loRegPig.dUpdateColocPigRGDet("S", psNroProceso, psCtaCod, gColPRecGarVentaVendido, pnMontoTransac, , , , lnMovNro, False)
    
        '** Inserta ColocPigRecupVta
        Call loRegPig.dInsertColocPigRecupVta("S", psNroProceso, psCtacod, 3, psNroDoc, psPersCod, pnMontoTransac, 0, 0, False)
    
    
    Else
    
        '** Actualiza Producto
        Call loRegPig.dUpdateProducto(psCtacod, , , "2113", psFechaHora, -2, False)             ' (-2) aumenta el ste transac
        '** Actualiza Colocaciones
        Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)
        '** Insert ColocEstado
        Call loRegPig.dInsertColocacEstado(psCtacod, psFechaHora, "2113", 1, pnMontoTransac, "Recuperacion Joyas Adjudicadas", _
         gColocCalendCodPFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
    
        '** Inserta Mov
        Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Recuperacion Joyas Adjudicadas", gMovEstContabMovContable, gMovFlagVigente, False)
        
        ' Obtiene nMovNro
        lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        
        '** Inserta MovCol
        Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)
    
        '** Inserta MovColDet -  Deuda Vencida
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodCapitalVencido, 1, pnVtaNetaSubasta, False)
    
        '** Inserta MovColDet -  Interes
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodInteresMoratorio, 1, pnRecupInteres, False)
    
        ''** Inserta MovColDet -  Impuesto Vta Subasta
        'Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, geColPConceptoCodImpuestoVtaSubasta, 1, pnImpuestoVtaSubasta, False)
        
        '** Inserta MovColDet -  Valor Adjudicacion
        'Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, geColPConceptoCodValorAdjudica, 1, pnValorAdjudicacion, False)
    
        '** Inserta MovColDet ' Oro 14/16/18/21
        If pnki14 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro14, 1, pnki14, False)
        End If
        If pnki16 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro16, 1, pnki16, False)
        End If
        If pnKi18 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro18, 1, pnKi18, False)
        End If
        If pnKi21 > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodOro21, 1, pnKi21, False)
        End If
        
    If pbITFAplica Then
        If Not pbITFAsumidoCreditos Then
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, pnMontoITf, 0, "", 0, 0, 0, False)
        
            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, gConcITFCliente, 1, pnMontoITf, False)
        Else
            '** Inserta MovCol
            Call loRegPig.dInsertMovCol(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, pnMontoITf, 0, "", 0, 0, 0, False)
        
            '** Inserta MovColDet -  Monto Entregar
            Call loRegPig.dInsertMovColDet(lnMovNro, gITFCobroEfectivoPrend, psCtacod, 1, gConcITFAsumido, 1, pnMontoITf, False)
        End If
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

Public Sub nPagoSobranteRemate(ByVal psCtacod As String, ByVal psNroProceso As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnMontoTransac As Currency, ByVal psCtaAhorroRemate As String, _
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
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = geColPPagSobrante

'On Error GoTo ErrorModPig

    'Retira de la la Cta de Ahorros de sobrante
    '****
    Dim oCapMov As NCapMovimientos
    Dim sMovNroReg As String
    Dim nMovNroReg As Long
    Dim oCont As NContFunciones
    Set oCont = New NContFunciones
        sMovNroReg = oCont.GeneraMovNro(Format(psFechaHora, "yyyy/mm/dd"), Mid(psMovNro, 18, 2), Mid(psMovNro, 22, 4))
    Set oCont = Nothing
    Set oCapMov = New NCapMovimientos
        'oCapMov.CapAbonoCuentaAho psCtaAhorroRemate, pnMontoTransac, gAhoDepSobRemate, sMovNroReg, "Dep. Sobrante Remate" & psNroRemate & " - " & psProcesoCadaAgencia, , , , , , , , , , psLpt, , True
        oCapMov.CapCargoCuentaAho psCtaAhorroRemate, pnMontoTransac, gAhoRetEfec, sMovNroReg, "Ret. Pago Sobrante Remate-" & psCtacod, , , , , , , , , , psLpt, , True
    Set oCapMov = Nothing

    '****
    
Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , , psFechaHora, -2, False)              ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)
   
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Pago Sobrante Remate", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("R", psNroProceso, psCtacod, gColPRecGarVentaSobranteCobrado, , , , , lnMovNro, False)
    
    '** Update ColocPigRecupVta
    Call loRegPig.dUpdateColocPigRecupVta("R", psNroProceso, psCtacod, , , , , , , psFechaHora, False)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet -  Monto Sobrante
    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColPConceptoCodSobranteRemate, 1, pnMontoTransac, False)
    
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
    Err.Raise vbObjectError + 100, "Error nPagoSobranteRemate", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nExtornoPagoSobranteRemate(ByVal psCtacod As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal pnMovNroAnt As Long, _
        ByVal pnMontoTransac As Currency, _
        ByVal psLpt As String, ByVal psNroProceso As String, ByVal psCtaAhorroRemate As String, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto ( NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** Actualiza ColocPigRGDet (Estado,  nMovVta)
'** Actualiza ColocPigRecupVta (FechaPago)
'** dInsertMov
'** dInsertMovCol
'************************************
 
Dim lsSQL As String
Dim loRegPig As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = "129702"
'On Error GoTo ErrorModPig

    'Deposita a la Cta de Ahorros de sobrante
    '****
    Dim oCapMov As NCapMovimientos
    Dim sMovNroReg As String
    Dim nMovNroReg As Long
    Dim oCont As NContFunciones
    Set oCont = New NContFunciones
        sMovNroReg = oCont.GeneraMovNro(Format(psFechaHora, "yyyy/mm/dd"), Mid(psMovNro, 16, 2), Mid(psMovNro, 22, 4))
    Set oCont = Nothing
    Set oCapMov = New NCapMovimientos
        oCapMov.CapAbonoCuentaAho psCtaAhorroRemate, pnMontoTransac, gAhoDepEfec, sMovNroReg, "Extorno Pago Sobrante Remate-" & psCtacod, , , , , , , , , , psLpt, , True
        'oCapMov.CapCargoCuentaAho psCtaAhorroRemate, pnMontoTransac, gAhoRetEfec, psMovNro, "Ret. Pago Sobrante Remate" & psCtacod, , , , , , , , , , psLpt, , True
    Set oCapMov = Nothing

    '****

Set loRegPig = New DColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegPig.dUpdateProducto(psCtacod, , , , psFechaHora, -2, False)              ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loRegPig.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)
   
    '** Update ColocPigRGDet
    Call loRegPig.dUpdateColocPigRGDet("R", psNroProceso, psCtacod, gColPRecGarVentaVendido, , , , , 0, False)
    
    '** Update ColocPigRecupVta
    'Call loRegPig.dUpdateColocPigRecupVta("R", psNroProceso, psCtacod, , , , , , , , False)
    
    '** Inserta Mov
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Extorno Pago Sobrante Remate", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnMontoTransac, 0, "", 0, 0, 0, False)
    
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

