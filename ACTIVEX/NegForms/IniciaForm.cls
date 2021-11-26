VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniciaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************** TORE ***************************'
'Razon: Creado con la finalidad de migrar todos los formularios
'       SICMACM NEGOCIO a este DLL.
'Fecha: 30/04/2021
'*************************************************************'
Option Explicit



Public Sub AbirForm(ByVal psCodForm As Integer)
    Select Case psCodForm
        Case 10000
            frmColPPreparacionRetasacionVigDif.Show 1
        Case 10001
            frmColPRetasacionConsulta.Show 1
        Case 10002
            frmColPRetasacionVigenteDiferida.Show 1
    End Select
End Sub



