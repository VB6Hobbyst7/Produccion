VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NComColRecComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mnEjecutaBatch As Integer
Dim mbTrans As Boolean

Public Sub nGrabarNuevaComisionAbogado(ByVal psCodAbogado As String, ByVal pnTipComision As Integer, _
        ByVal pnValor As Double, ByVal pnRangIni As Currency, ByVal pnRangFin As Currency, _
        ByVal pnCategoria As Integer, Optional pbEjecBatch As Boolean = False)


'** Inserta ColocRecComision
'************************************

Dim lsSQL As String
Dim loReg As COMDColocPig.DCOMColPActualizaBD
Dim loNueComi As COMDColocRec.DCOMColRecComision
Dim lnNueComis As Integer
'On Error GoTo ControlError

Set loNueComi = New COMDColocRec.DCOMColRecComision
    lnNueComis = loNueComi.nGeneraNuevaComision
Set loNueComi = Nothing

Set loReg = New COMDColocPig.DCOMColPActualizaBD
    loReg.dBeginTrans
    mbTrans = True
    
    '** Inserta Comision
    Call loReg.dInsertColocRecComision(lnNueComis, psCodAbogado, pnTipComision, pnValor, pnRangIni, pnRangFin, pnCategoria, False)

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

Public Sub nGrabarModificaComisionAbogado(ByVal pnNroComis As Long, ByVal pnTipComision As Integer, _
        ByVal pnValor As Double, ByVal pnRangIni As Currency, ByVal pnRangFin As Currency, _
        ByVal pnCategoria As Integer, Optional pbEjecBatch As Boolean = False)


'** Update ColocRecupComision
'************************************

Dim lsSQL As String
Dim loReg As COMDColocPig.DCOMColPActualizaBD

'On Error GoTo ControlError

Set loReg = New COMDColocPig.DCOMColPActualizaBD
    loReg.dBeginTrans
    mbTrans = True
    
    '** Update Comision
    Call loReg.dUpdateColocRecComision(pnNroComis, pnTipComision, pnValor, pnRangIni, pnRangFin, pnCategoria, False)

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

Public Function nObtieneListaComisionAbogado(ByVal psCodAbog As String, ByRef psmensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loRegLista As COMDColocPig.DCOMColPFunciones
Dim lrLista As ADODB.Recordset

    lsSQL = " SELECT nRangIni, nRangFin, nTipComis, nValor, nComisionCod, cPersCod, Modif = ' ', nCategoria " _
          & " FROM ColocRecupComision " _
          & " WHERE cPersCod = '" & psCodAbog & "' " _
          & " ORDER BY  nRangIni, nRangFin "
    
    Set loRegLista = New COMDColocPig.DCOMColPFunciones
        Set lrLista = loRegLista.dObtieneRecordSet(lsSQL)
        If lrLista Is Nothing Then
            psmensaje = "ERROR: al Buscar datos para Validación "
            Exit Function
        End If
    Set nObtieneListaComisionAbogado = lrLista
    Set lrLista = Nothing
End Function



