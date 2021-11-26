VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NColRecComision"
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
Dim loReg As DColPActualizaBD
Dim loNueComi As DColRecComision
Dim lnNueComis As Integer
'On Error GoTo ControlError

Set loNueComi = New DColRecComision
    lnNueComis = loNueComi.nGeneraNuevaComision
Set loNueComi = Nothing

Set loReg = New DColPActualizaBD
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
Dim loReg As DColPActualizaBD

'On Error GoTo ControlError

Set loReg = New DColPActualizaBD
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


Public Function nObtieneListaComisionAbogado(ByVal psCodAbog As String) As Recordset
Dim lsSQL As String
Dim loRegLista As dColPFunciones
Dim lrLista As ADODB.Recordset

    lsSQL = " SELECT nRangIni, nRangFin, nTipComis, nValor, nComisionCod, cPersCod, Modif = ' ', nCategoria " _
          & " FROM ColocRecupComision " _
          & " WHERE cPersCod = '" & psCodAbog & "' " _
          & " ORDER BY  nRangIni, nRangFin "
    
    Set loRegLista = New dColPFunciones
        Set lrLista = loRegLista.dObtieneRecordSet(lsSQL)
        If lrLista Is Nothing Then
            MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
            Exit Function
        End If
    Set nObtieneListaComisionAbogado = lrLista
    Set lrLista = Nothing
End Function

