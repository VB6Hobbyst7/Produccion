VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NColRecCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mnEjecutaBatch As Integer
Dim mbTrans As Boolean


Public Sub nCambiaMetodoLiquidCredRecup(ByVal psCtacod As String, ByVal psMetLiquid As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (NroTrans )
'** Actualiza ColocCredRecup (Metodo Liquid)
'************************************

Dim lsSQL As String
Dim loRegRec As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
'lsOpeCod = geColPImpDuplicado

'On Error GoTo ErrorModRec
Set loRegRec = New DColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegRec.dUpdateProducto(psCtacod, , , , psFechaHora, -2, False)           ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    'Call loRegRec.dUpdateColocaciones(psCtaCod, , , , ,  psMovNro, , False)

    '** Actualiza ColocRecup
    Call loRegRec.dUpdateColocRecup(psCtacod, , , , , , , psMetLiquid, , , , False)

    '** Inserta Mov
    'Call loRegRec.dInsertMov(psMovNro, lsOpeCod, "Cambio Metodo Liquid", gMovEstContabNoContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    'lnMovNro = loRegrec.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    'Call loRegRec.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 1, pnMontoTransac, 0, "", 0, False)
    
    'mnEjecutaBatch = loRegRec.dEjecutaBatch
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCambiaMetodoLiquidCredRecup", "Error en Funcion de Cambio Metodo Liquid Recuperaciones "

End Sub

Public Sub nGastoRecupAsignaNuevo(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnNroGastoCta As Integer, ByVal psFecAsigna As String, ByVal pnColocConceptoCod As Integer, _
        ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, ByVal pnColocRecGastoEstado As ColocRecGastoEstado, _
        ByVal psMotivoGasto As String, ByVal pnNewSaldoGasto As Currency, Optional pbEjecBatch As Boolean = False)

'** Inserta ColocRecGastos
'** Actualiza ColocRecup
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loReg As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = geColPExtDupli

On Error GoTo ControlError
Set loReg = New DColPActualizaBD
    loReg.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    'Call loReg.dUpdateProducto(psCtaCod, , , , psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    
    '** Actualiza ColocRecup
    Call loReg.dUpdateColocRecup(psCtacod, , , , , pnNewSaldoGasto, , , , , , False)

    '** Inserta ColocRecupGastos
    Call loReg.dInsertColocRecupGastos(psCtacod, pnNroGastoCta, pnColocConceptoCod, psFecAsigna, pnMonto, pnMontoPagado, _
            pnColocRecGastoEstado, psMotivoGasto, False)

    '** Inserta Mov
    'Call loReg.dInsertMov(psMovNro, lsOpeCod, "Ext Duplicado Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    'lnMovNro = loReg.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    'Call loReg.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 1, pnMonto, 0, "", 0, False)
    

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ControlError:
    If mbTrans Then
        loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nGastoRecupModifica(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnNroGastoCta As Integer, ByVal psFecAsigna As String, ByVal pnColocConceptoCod As Integer, _
        ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, ByVal pnColocRecGastoEstado As ColocRecGastoEstado, _
        ByVal psMotivoGasto As String, ByVal pnNewSaldoGasto As Currency, Optional pbEjecBatch As Boolean = False)

'** Update ColocRecGastos
'** Actualiza ColocRecup
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loReg As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = geColPExtDupli

On Error GoTo ControlError
Set loReg = New DColPActualizaBD
    loReg.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    'Call loReg.dUpdateProducto(psCtaCod, , , , psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    
    '** Actualiza ColocRecup
    Call loReg.dUpdateColocRecup(psCtacod, , , , , pnNewSaldoGasto, , , , , , False)

    '** Update ColocRecupGastos
    Call loReg.dUpdateColocRecupGastos(psCtacod, pnNroGastoCta, , pnMonto, , _
            , psMotivoGasto, False)

    '** Inserta Mov
    'Call loReg.dInsertMov(psMovNro, lsOpeCod, " ", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    'lnMovNro = loReg.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    'Call loReg.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 1, pnMonto, 0, "", 0, False)
    

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ControlError:
    If mbTrans Then
        loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nGastoRecupModifica", "Error en Funcion de Modifica Gasto Recup "

End Sub

Public Sub nGastoRecupEliminaGasto(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnNroGastoCta As Integer, _
        ByVal pnColocRecGastoEstado As ColocRecGastoEstado, _
        ByVal pnNewSaldoGasto As Currency, Optional pbEjecBatch As Boolean = False)

'** Actualiza ColocRecup
'** Update ColocRecGastos
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loReg As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = geColPExtDupli

On Error GoTo ControlError
Set loReg = New DColPActualizaBD
    loReg.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    'Call loReg.dUpdateProducto(psCtaCod, , , , psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    
    '** Actualiza ColocRecup
    Call loReg.dUpdateColocRecup(psCtacod, , , , , pnNewSaldoGasto, , , , , , False)

    '** Update ColocRecupGastos
    Call loReg.dUpdateColocRecupGastos(psCtacod, pnNroGastoCta, , , , pnColocRecGastoEstado, , False)
    
    '** Inserta Mov
    'Call loReg.dInsertMov(psMovNro, lsOpeCod, "Ext Duplicado Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    'lnMovNro = loReg.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    'Call loReg.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 1, pnMonto, 0, "", 0, False)
    

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ControlError:
    If mbTrans Then
        loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nGastoRecupEliminaGasto", "Error en Funcion de Elimina Gasto Recuperaciones"

End Sub

Public Sub nCastigaCredRecup(ByVal psCtacod As String, ByVal psMetLiquid As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal pnSaldoCap As Double, _
        ByVal pnSaldoIntComp As Double, ByVal pnSaldoIntMora As Double, ByVal pnSaldoGasto As Double, _
        ByVal pnIntComGen As Double, ByVal pnDemanda As ColocRecDemandado, ByVal pnProvision As Double, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (NroTrans )
'** Actualiza ColocCredRecup (Metodo Liquid)
'** Actualiza Mov (Metodo Liquid)
'************************************

Dim lsSQL As String
Dim loRegRec As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lnGasto4 As Double, lnIngreso6 As Double
'Codigo de Operacion
If pnDemanda = gColRecDemandaNo Then
    lsOpeCod = gColRecOpeCastigSinDem     ' Cambiar  la constante
ElseIf pnDemanda = gColRecDemandaSi Then
    lsOpeCod = gColRecOpeCastigConDem
End If

' Determina Gastos o Ingresos
If pnSaldoCap > pnProvision Then
    lnGasto4 = pnSaldoCap - pnProvision
    lnIngreso6 = 0
ElseIf pnProvision > pnSaldoCap Then
    lnIngreso6 = pnProvision - pnSaldoCap
    lnGasto4 = 0
End If
        
'On Error GoTo ErrorModRec
Set loRegRec = New DColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loRegRec.dUpdateProducto(psCtacod, , , gColocEstRecVigCast, psFechaHora, -2, False)              ' (-2) aumenta el ste transac

    '** Actualiza Colocaciones
    Call loRegRec.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)

    '** Actualiza ColocRecup
    Call loRegRec.dUpdateColocRecup(psCtacod, , , pnSaldoIntComp, , , pnIntComGen, psMetLiquid, , , , False)
 
    '** Inserta Mov
    Call loRegRec.dInsertMov(psMovNro, lsOpeCod, "Catigar Credito", gMovEstContabMovContable, gMovFlagVigente, False)

    ' Obtiene nMovNro
    lnMovNro = loRegRec.dGetnMovNro(psMovNro)

    '** Inserta MovCol
    Call loRegRec.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 0, pnSaldoCap, 0, "", 0, 0, pnSaldoCap, False)
    
    '** Inserta MovColDet - Capital
    Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 0, gColRecConceptoCodCapital, 0, pnSaldoCap, False)
    
    If lnGasto4 > 0 Then  ' Ingresos
        Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 0, gColRecConceptoCodCastigGasto4, 0, lnGasto4, False)
    End If
    
    If lnIngreso6 > 0 Then  ' Gastos
        Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 0, gColRecConceptoCodCastigIngreso6, 0, lnIngreso6, False)
    End If
    
    ' Capital
    Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 0, gColRecConceptoCodCastigCtaOrdenCap, 0, pnSaldoCap, False)
    
    ' Int Comp
    If pnSaldoIntComp > 0 Then
        Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 0, gColRecConceptoCodCastigCtaOrdenIntCom, 0, pnSaldoIntComp, False)
    End If
    
    ' Int Morat
    If pnSaldoIntMora > 0 Then
        Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 0, gColRecConceptoCodCastigCtaOrdenIntMor, 0, pnSaldoIntMora, False)
    End If
    
    If pnSaldoGasto > 0 Then
        Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 0, gColRecConceptoCodCastigCtaOrdenGasto, 0, pnSaldoGasto, False)
    End If

    'mnEjecutaBatch = loRegRec.dEjecutaBatch
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCambiaMetodoLiquidCredRecup", "Error en Funcion de Cambio Metodo Liquid Recuperaciones "

End Sub

Public Sub nRegistraCredEnRecup(ByVal psCtacod As String, ByVal pnTasaInteres As Double, ByVal pnSaldo As Currency, _
        ByVal pnPlazo As Integer, ByVal psLineaCred, ByVal prClientes As Recordset, ByVal psFecIngRecup As String, _
        ByVal pnSaldoIntComp As Currency, ByVal pnSaldoIntMor As Currency, ByVal pnSaldoGasto As Currency, _
        ByVal pnDemanda As Integer, ByVal psFechaHora As String, ByVal psMovNro As String, _
        Optional pbEjecBatch As Boolean = False)

'** Insert Producto
'** Insert Colocaciones
'** Insert ColocacEstado
'** Insert ColocRecup
'** Insert ColocRecupPersona
'** Insert ColocRecupEstado
'** Insert Mov
'** Insert MovCol
'************************************

Dim lsSQL As String
Dim loRegRec As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = gColPOpeImpDuplicado

'On Error GoTo ErrorModRec
Set loRegRec = New DColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    '** Inserta Producto
    Call loRegRec.dInsertProducto(psCtacod, pnTasaInteres, pnSaldo, gColocEstRecVigJud, psFechaHora, 1, False)
    
    '** Inserta Colocaciones
    Call loRegRec.dInsertColocaciones(psCtacod, 30, psFecIngRecup, pnSaldo, gColocCalendCodPFCF, psMovNro, psLineaCred, psFechaHora, False)
    
    '** Inserta ColocEstado
    Call loRegRec.dInsertColocacEstado(psCtacod, psFechaHora, gColocEstRecVigJud, 1, pnSaldo, "Ingreso a Recuperaciones", gColocCalendCodPFCF, 0, 0, 0, 0, 0, 0, 0, 0, False)
    
    '** Inserta ColocRecup
    Call loRegRec.dInsertColocRecup(psCtacod, psFecIngRecup, -1, pnSaldoIntComp, pnSaldoIntMor, pnSaldoGasto, 0, "GCIM", -1, pnDemanda, 1, False)
    
    '** Inserta ColocRecupEstado
    Call loRegRec.dInsertColocRecupEstado(psCtacod, psFechaHora, gColocEstRecVigJud, "", False)
    
    '** Inserta ColocRecupExpediente
    'Call loRegRec.dInsertColocRecupExpediente(psCtaCod, psFechaHora, gColocEstRecVigJud, "", False)
    
    '** Inserta ColocRecupComisionHist
    'Call loRegRec.dInsertColocRecupExpediente(psCtaCod, psFechaHora, gColocEstRecVigJud, "", False)
    
    '** Inserta ColocRecupPersonas
    'inserta Clientes para cada registro de prClientes
    Do While Not prClientes.EOF
        Call loRegRec.dInsertProductoPersona(psCtacod, prClientes("cPersCod"), prClientes("cPersRela"), False)
        prClientes.MoveNext
    Loop
    
    '** Inserta Mov
    Call loRegRec.dInsertMov(psMovNro, lsOpeCod, "Registro En Recuperaciones", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegRec.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegRec.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 1, pnSaldo, 0, "", 0, gColocEstRecVigJud, pnSaldo, False)
    
    '** Inserta MovColDet
    Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, 1, gColRecConceptoCodCapital, 0, pnSaldo, False)
    
    'mnEjecutaBatch = loRegRec.dEjecutaBatch
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCambiaMetodoLiquidCredRecup", "Error en Funcion de Cambio Metodo Liquid Recuperaciones "

End Sub

Public Sub nExtornoTransfRecup(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto ( Estado )
'** Elimina ColocRecupExpedInf
'** Elimina ColocRecupExped
'** Elimina ColocRecupGasto
'** Elimina ColocRecup
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loBase As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim loDataExt As DConecta
Dim lrDataExt As Recordset
Dim loFunciones As NContFunciones

lsOpeCod = gColRecOpeExtTransfRecup

' *** Obtiene Datos para el Extorno

lsSQL = "SELECT CE.nPrdEstado " _
      & "FROM ColocacEstado CE WHERE CE.cCtaCod = '" & psCtacod & "' " _
      & "AND CE.nPrdEstado <> " & gColocEstRecVigJud & "   " _
      & "ORDER BY CE.dPrdEstado DESC "

Set loDataExt = New DConecta
    loDataExt.AbreConexion
    Set lrDataExt = loDataExt.CargaRecordSet(lsSQL)
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
Dim lnAntEstado As Integer

lnAntEstado = lrDataExt!nPrdEstado

' *** Realiza el Extorno
'On Error GoTo Error
Set loBase = New DColPActualizaBD
    loBase.dBeginTrans
    mbTrans = True
    
    '** Update Producto
    Call loBase.dUpdateProducto(psCtacod, , , lnAntEstado, psFechaHora, -2, False)
    '** Delete ColocRecupExpedienteInf
    Call loBase.dDeleteColocRecupExpedienteInf(psCtacod, , False)
    '** Delete ColocRecupExpediente
    Call loBase.dDeleteColocRecupExpediente(psCtacod, False)
    '** Delete ColocRecupGastos
    Call loBase.dDeleteColocRecupGastos(psCtacod, , False)
    '** Delete ColocRecup
    Call loBase.dDeleteColocRecup(psCtacod, False)
    
    '** Insert ColocEstado
    Call loBase.dInsertColocacEstado(psCtacod, psFechaHora, lnAntEstado, 1, pnMonto, "Extorno Transferencia Recup", _
          0, 0, 0, 0, 0, 0, 0, 0, 0, False)

    '** Inserta Mov
    Call loBase.dInsertMov(psMovNro, lsOpeCod, "Ext Transferencia Recup", gMovEstContabNoContable, gMovFlagDeExtorno, False)
    
    ' Obtiene nMovNro
    lnMovNro = loBase.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 0, pnMonto, 0, "", 0, 0, 0, False)

    '** Update Mov Anterior
    Call loBase.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loBase.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
    
    '**
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loBase.dCommitTrans
    mbTrans = False
Set loBase = Nothing

Exit Sub
ERROR:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nExtornoTransFRecup", "Error en Funcion de Extorno de Transferencia"

End Sub
Public Sub nRegistraExpedienteRecup(ByVal psCtacod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnComisionCod As Integer, ByVal pnTipCj As Integer, ByVal psNumExp As String, ByVal pnMonPetit As Currency, _
        ByVal pnMoneda As Integer, ByVal prClientes As Recordset, ByVal psPetit As String, ByVal psHechos As String, _
        ByVal psFundJur As String, ByVal psMedProb As String, ByVal psDatComp As String, _
        ByVal pnViaProce As Integer, ByVal pnEstadoProceso As Integer, ByVal pbCambioComision As Boolean, _
        ByVal pbNuevoExpediente As Boolean, Optional pbEjecBatch As Boolean = False)

'** Actualiza ColocRecup
'** Actualiza ColocRecupExpediente
'** Actualiza ColocRecupComisionHist
'** Inserta Mov
'************************************

Dim lsSQL As String
Dim loRegRec As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = 135001

'On Error GoTo ErrorModRec
Set loRegRec = New DColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    '** Actualiza ColocRecup
    Call loRegRec.dUpdateColocRecup(psCtacod, , pnComisionCod, , , , , , pnTipCj, , , False)
    
    '** Actualiza ColocRecupExpediente
    If pbNuevoExpediente = False Then ' Existe Expediente
        Call loRegRec.dInsertColocRecExpediente(psCtacod, psNumExp, pnMonPetit, pnMoneda, _
                pnViaProce, pnEstadoProceso, False)
    Else
        Call loRegRec.dUpdateColocRecupExpediente(psCtacod, psNumExp, pnMonPetit, pnMoneda, _
                pnViaProce, pnEstadoProceso, False)
    End If
    '** Registra el cambio de comision
    If pbCambioComision = True Then
        '** Actualiza ColocRecupComisionHist
        Call loRegRec.dInsertColocRecComisionHist(psCtacod, pnComisionCod, lnMovNro, False)
    End If

    '** Borra las personas relacionadas
    Call loRegRec.dDeleteProductoPersona(psCtacod, "", gColRelPersEstudioJuridico, False)
    Call loRegRec.dDeleteProductoPersona(psCtacod, "", gColRelPersJuzgado, False)
    Call loRegRec.dDeleteProductoPersona(psCtacod, "", gColRelPersJuez, False)
    Call loRegRec.dDeleteProductoPersona(psCtacod, "", gColRelPersSecretario, False)
    '** Registra las personas relacionadas
    prClientes.MoveFirst
    Do While Not prClientes.EOF
        Call loRegRec.dInsertProductoPersona(psCtacod, prClientes("cPersCod"), prClientes("cPersRelac"), False)
        prClientes.MoveNext
    Loop
    
    If Len(psPetit) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInf(psCtacod, gColRecExpedTipoInfPetit, psPetit, False)
    End If
    If Len(psHechos) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInf(psCtacod, gColRecExpedTipoInfHechos, psHechos, False)
    End If
    If Len(psFundJur) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInf(psCtacod, gColRecExpedTipoInfFundJur, psFundJur, False)
    End If
    If Len(psMedProb) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInf(psCtacod, gColRecExpedTipoInfMedProb, psMedProb, False)
    End If
    If Len(psDatComp) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInf(psCtacod, gColRecExpedTipoInfDatComp, psDatComp, False)
    End If
    
    '** Inserta Mov
    Call loRegRec.dInsertMov(psMovNro, lsOpeCod, "Modifica Exped Judic", gMovEstContabNoContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegRec.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    'Call loRegRec.dInsertMovCol( lnMovNro, lsOpeCod, psCtaCod, 1, pnMontoTransac, 0, "", 0, False)
    
    'mnEjecutaBatch = loRegRec.dEjecutaBatch
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraExpedienteRecup", "Error en Funcion de Registro Expediente Recuperaciones "

End Sub


Public Sub nPagoCreditoRecup(ByVal psCtacod As String, ByVal psFechaHora As String, ByVal psOpeCod As String, _
        ByVal psMovNro As String, ByVal pnMontoPagar As Currency, ByVal psMetLiquid As String, _
        ByVal pnNewSaldoCap As Currency, ByVal pnNewSaldoIntCom As Currency, _
        ByVal pnNewSaldoIntMor As Currency, ByVal pnNewSaldoGasto As Currency, _
        ByVal pnNewSaldoIntComGener As Currency, ByVal pnComisionAbog As Currency, ByVal pnNroGastoCta As Integer, _
        ByVal pnGastoAdmin As Currency, ByVal pnCapPag As Currency, _
        ByVal pnIntCompPag As Currency, ByVal pnIntMoratPag As Currency, _
        ByVal pnGastoPag As Currency, ByVal pmGastos As Variant, ByVal pnNroCalend As Integer, ByVal pnEstadoNew As ColocEstado, _
        Optional pbEjecBatch As Boolean = False, Optional psPersLavDinero As String = "", Optional ByVal pnITF As Double = 0#, Optional pnDocTpo As Integer = 0, Optional psNroDoc As String = "")

'** Update Producto
'** Update Colocaciones
'** Update ColocRecup
'** Insert ColocRecupGasto (Comision)
'** Insert ColocRecupGasto (Gasto Admin)
'** Update ColocRecGastos
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loReg As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lnNroGastoSgte As Integer, I As Integer
Dim lnNroGastoComis As Integer, lnNroGastoAdmin As Integer

lsOpeCod = psOpeCod

lnNroGastoSgte = pnNroGastoCta

'On Error GoTo ControlError
Set loReg = New DColPActualizaBD
    loReg.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loReg.dUpdateProducto(psCtacod, , pnNewSaldoCap, pnEstadoNew, psFechaHora, -2, False)          ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loReg.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)
    
    '** Actualiza ColocRecup
    Call loReg.dUpdateColocRecup(psCtacod, , , pnNewSaldoIntCom, pnNewSaldoIntMor, pnNewSaldoGasto, pnNewSaldoIntComGener, , , , , False)

    'Si tiene Comision
    If pnComisionAbog > 0 Then
        '** Insert ColocRecupGastos
        lnNroGastoSgte = lnNroGastoSgte + 1
        lnNroGastoComis = lnNroGastoSgte
        Call loReg.dInsertColocRecupGastos(psCtacod, lnNroGastoComis, gColRecConceptoCodGastoComisionAbog, psFechaHora, pnComisionAbog, pnComisionAbog, gColRecGastoEstPagado, "Comision Abogado", False)
    End If
    
    'Para cada Gasto que se tenga
    For I = 0 To UBound(pmGastos) - 1
        If pmGastos(I, 5) = "S" Then  ' Si ha sido modificado
            '** Update ColocRecupGastos
            Call loReg.dUpdateColocRecupGastos(psCtacod, CInt(pmGastos(I, 1)), , , Format(CDbl(pmGastos(I, 3)), "#0.00"), CInt(pmGastos(I, 4)), , False)
        End If
    Next I

    'Si tiene Gasto Administrativo
    If pnGastoAdmin > 0 Then
        '** Insert ColocRecupGastos
        lnNroGastoSgte = lnNroGastoSgte + 1
        lnNroGastoAdmin = lnNroGastoSgte
        Call loReg.dInsertColocRecupGastos(psCtacod, lnNroGastoAdmin, gColRecConceptoCodGastoAdmin, psFechaHora, pnComisionAbog, pnGastoAdmin, gColRecGastoEstPagado, "Gasto Administrativ", False)
    End If

    '** Inserta Mov
    Call loReg.dInsertMov(psMovNro, lsOpeCod, "Pago Credito Recup", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loReg.dGetnMovNro(psMovNro)
    
    If psNroDoc <> "" Then
        loReg.dInsertMovDoc lnMovNro, pnDocTpo, psNroDoc, psFechaHora
    End If
    'ITF
    If pnITF > 0 Then
        Call loReg.dInsertMovCol(lnMovNro, "990105", psCtacod, pnNroCalend, pnITF, 0, psMetLiquid, 0, 0, pnNewSaldoCap, False)
        Call loReg.dInsertMovColDet(lnMovNro, "990105", psCtacod, pnNroCalend, gConcITFCliente, 0, pnITF, False)
    End If
    
    '** Inserta MovCol
    Call loReg.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, pnNroCalend, pnMontoPagar, 0, psMetLiquid, 0, 0, pnNewSaldoCap, False)
    
    '** Inserta MovColDet - Capital
    If pnCapPag > 0 Then
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, pnNroCalend, gColRecConceptoCodCapital, 0, pnCapPag, False)
    End If
    
    
    
    '** Inserta MovColDet - Interes Comp
    If pnIntCompPag > 0 Then
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, pnNroCalend, gColRecConceptoCodInteresCompensatorio, 0, pnIntCompPag, False)
    End If
    
    '** Inserta MovColDet - Interes Morat
    If pnIntMoratPag > 0 Then
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, pnNroCalend, gColRecConceptoCodInteresMoratorio, 0, pnIntMoratPag, False)
    End If
    
    '** Inserta MovColDet - Gastos (Sin Comision, ni Gasto Administrativo )
    'If pnGastoPag > 0 Then
    '    Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColRecConceptoCodGastoCobranza, 0, pnGastoPag, False)
    'End If
    
    'Para cada Gasto que se tenga (Sin Comision, ni Gasto Administrativo )
    For I = 0 To UBound(pmGastos) - 1
        If pmGastos(I, 5) = "S" Then  ' Si ha sido modificado
            '** Inserta MovColDet - Gastos (Sin Comision, ni Gasto Administrativo )
            Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, pnNroCalend, pmGastos(I, 7), pmGastos(I, 1), Format(CDbl(pmGastos(I, 3)), "#0.00"), False)
        End If
    Next I
    
    '** Inserta MovColDet - Comision Abogado
    If pnComisionAbog > 0 Then
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, pnNroCalend, gColRecConceptoCodGastoComisionAbog, lnNroGastoComis, pnComisionAbog, False)
    End If
    
    '** Inserta MovColDet - Gasto Administ
    If pnGastoAdmin > 0 Then
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtacod, pnNroCalend, gColRecConceptoCodGastoAdmin, lnNroGastoAdmin, pnGastoAdmin, False)
    End If
    
    ''''''''''''''''''''
    ' Lavado de Dinero
    If psPersLavDinero <> "" Then
        Call loReg.dInsertaMovLavDinero(lnMovNro, psPersLavDinero)
    End If
    ''''''''''''''''''''
    
    loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ControlError:
    If mbTrans Then
        loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nExtornoPagoCreditoRecup(ByVal psCtacod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto ( Estado )
'** Actualiza ColocRecup
'** Update ColocRecupGasto(Comision) <Flag ='X'>
'** Update ColocRecupGasto Estado , (-) Monto Pagado
'** Update ColocRecupGasto (Administrativo )
'** dInsertMov
'** dInsertMovRef
'************************************

Dim lsSQL As String, lsSQL2 As String
Dim loBase As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim loDataExt As DConecta
Dim lrDataExt As Recordset
Dim loFunciones As NContFunciones

Dim lnSaldCap As Currency, lnSaldIntComp As Currency, lnSaldIntMor As Currency, lnSaldGasto As Currency
Dim lnCapPag As Currency, lnIntComPag As Currency, lnIntMorPag As Currency, lnGastoPag As Currency
Dim lsUltMov As String

lsOpeCod = gColRecOpeExtPagRecup

' *** Obtiene Datos de Montos Cubiertos

lsSQL = "SELECT p.nSaldo, cr.nSaldoIntComp, cr.nSaldoIntMor, cr.nSaldoGasto, cr.nIntCompGen, " _
    & " CapPag = (Select ISNULL(SUM(nMonto),0) From MovColDet mcd  Where mcd.nMovNro = " & pnMovNroAnt _
    & "           And mc.cCtaCod = '" & psCtacod & "' and mcd.nPrdConceptoCod in (" & gColRecConceptoCodCapital & " )  ) , " _
    & " IntComPag = (Select ISNULL(SUM(nMonto),0) From MovColDet mcd  Where mcd.nMovNro = " & pnMovNroAnt _
    & "           And mc.cCtaCod =  '" & psCtacod & "' and mcd.nPrdConceptoCod in (" & gColRecConceptoCodInteresCompensatorio & " )  ) , " _
    & " IntMorPag = (Select ISNULL(SUM(nMonto),0) From MovColDet mcd  Where mcd.nMovNro = " & pnMovNroAnt _
    & "           And mc.cCtaCod =  '" & psCtacod & "' and mcd.nPrdConceptoCod in (" & gColRecConceptoCodInteresMoratorio & " )  ) , " _
    & " GastoPag = (Select ISNULL(SUM(nMonto),0) From MovColDet mcd  Where mcd.nMovNro = " & pnMovNroAnt _
    & "           And mc.cCtaCod =  '" & psCtacod & "' and mcd.nPrdConceptoCod in (" & gColRecConceptoCodGastoCobranza & " )  ),  " _
    & " cUltMov = (SELECT MAX(M.cMOVNRO) FROM MOV M JOIN MOVCOL MC1 ON MC1.NMOVNRO = M.NMOVNRO " _
    & "             WHERE  M.NMOVNRO <" & pnMovNroAnt & " AND M.NMOVFLAG =0 AND MC1.CCTACOD ='" & psCtacod & "') " _
    & " From MovCol mc " _
    & " Inner join Producto p on p.cCtaCod = mc.cCtaCod " _
    & " Inner join ColocRecup cr on p.cCtaCod = cr.cCtaCod " _
    & " Where mc.cCtaCod = '" & psCtacod & "' And mc.nmovnro =  " & pnMovNroAnt & " "

lsSQL2 = "SELECT mcd.nPrdConceptoCod, mcd.nNroCuota nNroGastoCta, mcd.nMonto nMontoPag, " _
    & " nMontoGasto = (Select CRG.nMonto From ColocRecupGastos CRG Where CRG.cCtaCod = '" & psCtacod & "' And CRG.nNroGastoCta = mcd.nNroCuota ) " _
    & " FROM Mov m INNER JOIN MovCol mc on m.nMovNro = mc.nMovNro " _
    & " INNER JOIN MovColDet mcd on mc.nMovNro = mcd.nMovNro " _
    & " WHERE mc.cCtaCod = '" & psCtacod & "'  and m.nmovnro = " & pnMovNroAnt & " " _
    & " AND mcd.nPrdConceptoCod in (" & gColRecConceptoCodGastoCobranza & "," & gColRecConceptoCodGastoComisionAbog & "," & gColRecConceptoCodGastoAdmin & " ) "

Set loDataExt = New DConecta
    loDataExt.AbreConexion
    '** Carga Conceptos Pagados
    Set lrDataExt = loDataExt.CargaRecordSet(lsSQL)
    If lrDataExt Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Extorno ", vbInformation, "Aviso"
        Exit Sub
    End If
    If lrDataExt.BOF And lrDataExt.EOF Then
        MsgBox "ERROR: No encontro datos de la Operacion a extornar ", vbInformation, "Aviso"
        Exit Sub
    End If
    ' Asigna los valores
    lnSaldCap = lrDataExt!nSaldo:           lnSaldIntComp = lrDataExt!nSaldoIntComp - lrDataExt!nIntCompGen
    lnSaldIntMor = lrDataExt!nSaldoIntMor:  lnSaldGasto = lrDataExt!nSaldoGasto
    lnCapPag = lrDataExt!CapPag:      lnIntComPag = lrDataExt!IntComPag
    lnIntMorPag = lrDataExt!IntMorPag: lnGastoPag = lrDataExt!GastoPag
    lsUltMov = IIf(IsNull(lrDataExt!cUltMov), "@", lrDataExt!cUltMov)
    
    Set lrDataExt = Nothing
    '** Carga los Gastos Cobrados ' Concepto / NroGasto / MontoPagado
    Set lrDataExt = loDataExt.CargaRecordSet(lsSQL2)
    
Set loDataExt = Nothing

' *** Realiza el Extorno
'On Error GoTo Error
Set loBase = New DColPActualizaBD
    loBase.dBeginTrans
    mbTrans = True
    
    '** Update Producto
    Call loBase.dUpdateProducto(psCtacod, , lnSaldCap + lnCapPag, , psFechaHora, -2, False)
    
    '** Update COLOCACIONES
    Call loBase.dUpdateColocaciones(psCtacod, , , , , lsUltMov, , False)
    
    '** Update ColocRecup
    Call loBase.dUpdateColocRecup(psCtacod, , , lnSaldIntComp + lnIntComPag, lnSaldIntMor + lnIntMorPag, lnSaldGasto + lnGastoPag, 0, , , , , False)
    
    'Regresa los Gastos a su estado anterior
    Do While Not lrDataExt.EOF
        If lrDataExt!nPrdConceptoCod = gColRecConceptoCodGastoComisionAbog Then
            Call loBase.dUpdateColocRecupGastos(psCtacod, lrDataExt!nNroGastoCta, , , , gColRecGastoEstEliminado, , False)
        ElseIf lrDataExt!nPrdConceptoCod = gColRecConceptoCodGastoComisionAbog Then
            Call loBase.dUpdateColocRecupGastos(psCtacod, lrDataExt!nNroGastoCta, , , , gColRecGastoEstEliminado, , False)
        Else
            Call loBase.dUpdateColocRecupGastos(psCtacod, lrDataExt!nNroGastoCta, , , lrDataExt!nMontoGasto - lrDataExt!nMontoPag, gColRecGastoEstPendiente, , False)
        End If
        lrDataExt.MoveNext
    Loop
    
    '** Inserta Mov
    Call loBase.dInsertMov(psMovNro, lsOpeCod, "Ext.Pago.Credito.Recup", gMovEstContabNoContable, gMovFlagDeExtorno, False)
    
    ' Obtiene nMovNro
    lnMovNro = loBase.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, 0, pnMonto, 0, "", 0, 0, 0, False)

    '** Update Mov Anterior
    Call loBase.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loBase.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
    
    '**
    'mnEjecutaBatch = lobase.dEjecutaBatch
    loBase.dCommitTrans
    mbTrans = False
Set loBase = Nothing

Exit Sub
ERROR:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nExtornoTransFRecup", "Error en Funcion de Extorno de Transferencia"

End Sub

Public Sub nCancelacionCreditoRecup(ByVal psCtacod As String, ByVal psFechaHora As String, ByVal psOpeCod As String, _
        ByVal psMovNro As String, ByVal pnPrdEstado As Integer, _
        ByVal pnNewSaldoCap As Currency, ByVal pnNewSaldoIntCom As Currency, _
        ByVal pnNewIntComGener As Currency, ByVal pnNroCalend As Integer, _
        Optional pbEjecBatch As Boolean = False)

'** Update Producto
'** Update Colocaciones
'** Update ColocRecup
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loReg As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String


lsOpeCod = psOpeCod

'On Error GoTo ControlError
Set loReg = New DColPActualizaBD
    loReg.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loReg.dUpdateProducto(psCtacod, , , pnPrdEstado, psFechaHora, -2, False)               ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loReg.dUpdateColocaciones(psCtacod, , , , , psMovNro, , False)
    
    '** Actualiza ColocRecup
    Call loReg.dUpdateColocRecup(psCtacod, , , pnNewSaldoIntCom, , , pnNewIntComGener, , , , , False)

    '** Inserta Mov
    Call loReg.dInsertMov(psMovNro, lsOpeCod, "Cancelacion Cred Recup", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loReg.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loReg.dInsertMovCol(lnMovNro, lsOpeCod, psCtacod, pnNroCalend, pnNewSaldoCap, 0, "", 0, 0, pnNewSaldoCap, False)
    
    '** Inserta MovColDet - Capital
    'If pnCapPag > 0 Then
    '    Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColRecConceptoCodCapital, 0, pnCapPag, False)
    'End If
    
    'mnEjecutaBatch = loReg.dEjecutaBatch
    loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ControlError:
    If mbTrans Then
        loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCreditoRecup", "Error en Funcion Cancelacion Credito Recup "

End Sub


Public Sub nCierreMesRecuperaciones(ByVal psFechaHora As String, ByVal psOpeCod As String, _
        ByVal psMovNro As String, ByVal pnTipoCalcIntComp As Integer, ByVal pnTipoCalcIntMora As Integer, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (NroTrans )
'** Actualiza ColocCredRecup (Metodo Liquid)
'************************************

Dim lsSQL As String
Dim loRegRec As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim loRec As DColRecCredito
Dim lrRec As ADODB.Recordset
Dim loCalcula As NColRecCalculos
Dim lsFecUltPago As String
Dim lnDiasUltTrans As Integer

Dim lnNewSaldoIntCom As Currency, lnNewSaldoIntMor As Currency, lnNewSaldoIntComGener As Currency
Dim lnIntCompGenerado As Currency, lnIntMoraGenerado As Currency
'Codigo de Operacion
lsOpeCod = "131000"

'On Error GoTo ErrorModRec


    Set loRec = New DColRecCredito
        Set lrRec = loRec.dObtieneDatosCierreMes
    Set loRec = Nothing
    If lrRec Is Nothing Then Exit Sub
    If lrRec.BOF And lrRec.EOF Then Exit Sub
    Set loRegRec = New DColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    Do While Not lrRec.EOF
        lnIntCompGenerado = 0
        lnIntMoraGenerado = 0
        lsFecUltPago = CDate(fgFechaHoraGrab(lrRec!cUltimaActualizacion))
                            
        'lnDiasUltTrans = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(fsFecUltPago, "dd/mm/yyyy"))
        lnDiasUltTrans = CDate(fgFechaHoraGrab(psMovNro)) - CDate(Format(lsFecUltPago, "dd/mm/yyyy"))
        
        'Calcula los intereses Generados
        Set loCalcula = New NColRecCalculos
            If pnTipoCalcIntComp = 0 Then ' NoCalcula
                lnIntCompGenerado = 0
            ElseIf pnTipoCalcIntComp = 1 Then ' En base al capital
                lnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, lrRec!nTasaIntComp, lrRec!nSaldo)
            ElseIf pnTipoCalcIntComp = 2 Then ' En base a capit + int Comp
                lnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, lrRec!nTasaIntComp, lrRec!nSaldo + lrRec!nSaldoIntComp)
            ElseIf pnTipoCalcIntComp = 3 Then ' En base a capit + int Comp + int Morat
                lnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, lrRec!nTasaIntComp, lrRec!nSaldo + lrRec!nSaldoIntComp + lrRec!nSaldoIntMor)
            End If
            If pnTipoCalcIntMora = 0 Then  ' NoCalcula
                lnIntMoraGenerado = 0
            ElseIf pnTipoCalcIntMora = 1 Then ' En base al capital
                lnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, lrRec!nTasaIntMorat, lrRec!nSaldo)
            ElseIf pnTipoCalcIntMora = 2 Then ' En base a capit + int Comp
                lnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, lrRec!nTasaIntMorat, lrRec!nSaldo + lrRec!nSaldoIntComp)
            ElseIf pnTipoCalcIntMora = 3 Then ' En base a capit + int Comp + int Morat
                lnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, lrRec!nTasaIntMorat, lrRec!nSaldo + lrRec!nSaldoIntComp + lrRec!nSaldoIntMor)
            End If
        Set loCalcula = Nothing
        
        lnNewSaldoIntComGener = lrRec!nIntCompGen + Format(lnIntCompGenerado, "#.00") + Format(lnIntMoraGenerado, "#.00")
        lnNewSaldoIntCom = lrRec!nSaldoIntComp + Format(lnIntCompGenerado, "#.00")
        lnNewSaldoIntMor = lrRec!nSaldoIntMor + Format(lnIntMoraGenerado, "#.00")
        '******
        
        '** Actualiza Producto
        Call loRegRec.dUpdateProducto(lrRec!cCtacod, , , , psFechaHora, -2, False)           ' (-2) aumenta el ste transac
        
        '** Actualiza Colocaciones
        Call loRegRec.dUpdateColocaciones(lrRec!cCtacod, , , , , psMovNro, , False)
    
        '** Actualiza ColocRecup
        Call loRegRec.dUpdateColocRecup(lrRec!cCtacod, , , lnNewSaldoIntCom, lnNewSaldoIntMor, , lnNewSaldoIntComGener, , , , , False)
    
        '** Inserta Mov
        Call loRegRec.dInsertMov(psMovNro, lsOpeCod, "Cierre Mensual", gMovEstContabNoContable, gMovFlagVigente, False)
        
        ' Obtiene nMovNro
        lnMovNro = loRegRec.dGetnMovNro(psMovNro)
        
        '** Inserta MovCol
        Call loRegRec.dInsertMovCol(lnMovNro, lsOpeCod, lrRec!cCtacod, 0, lnNewSaldoIntComGener, 0, "", 0, 0, 0, False)
        
        '** Inserta MovColDet - Int Comp
        If lnIntCompGenerado > 0 Then
            Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, lrRec!cCtacod, 0, gColRecConceptoCodInteresCompensatorio, 0, lnIntCompGenerado, False)
        End If
        
        '** Inserta MovColDet - Int Comp
        If lnIntMoraGenerado > 0 Then
            Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, lrRec!cCtacod, 0, gColRecConceptoCodInteresMoratorio, 0, lnIntMoraGenerado, False)
        End If
        
        'mnEjecutaBatch = loRegRec.dEjecutaBatch
        
        lrRec.MoveNext
    Loop
    
    
    loRegRec.dCommitTrans
    mbTrans = False

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCambiaMetodoLiquidCredRecup", "Error en Funcion de Cambio Metodo Liquid Recuperaciones "

End Sub

Public Sub nRegistraActProcesales(ByVal psCtacod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal psComenta As String, Optional pbEjecBatch As Boolean = False)

'** Inserta Mov
'** Inserta ColocRecupActProc
'************************************

Dim lsSQL As String
Dim loRegRec As DColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = 135003

'On Error GoTo ErrorModRec
Set loRegRec = New DColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    '** Inserta Mov
    Call loRegRec.dInsertMov(psMovNro, lsOpeCod, "Modifica Exped Judic", gMovEstContabNoContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegRec.dGetnMovNro(psMovNro)
    
    '** Inserta ColocRecupActProcesales
    Call loRegRec.dInsertColocRecupActProcesales(psCtacod, psMovNro, psComenta, False)
    
    'mnEjecutaBatch = loRegRec.dEjecutaBatch
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraExpedienteRecup", "Error en Funcion de Registro Expediente Recuperaciones "

End Sub

