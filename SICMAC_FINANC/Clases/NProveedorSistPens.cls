VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NProveedorSistPens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim rs As ADODB.Recordset

Public Function ObtieneDatosSistemaPension(ByVal psPersCod As String) As ADODB.Recordset
    On Error GoTo ErrObtieneDatosSistemaPension
    Dim obj As New DProveedorSistPens
    Set ObtieneDatosSistemaPension = obj.ObtieneDatosSistemaPension(psPersCod)
    Exit Function
ErrObtieneDatosSistemaPension:
   Call RaiseError(MyUnhandledError, "NProveedorSistPens:ObtieneDatosSistemaPension Method")
End Function
Public Function ExisteDatosSistemaPension(ByVal psPersCod As String) As Boolean
    On Error GoTo ErrExisteDatosSistemaPension
    Dim obj As New DProveedorSistPens
    Set rs = obj.ObtieneDatosSistemaPension(psPersCod)
    If Not rs.EOF Then
        ExisteDatosSistemaPension = True
    Else
        ExisteDatosSistemaPension = False
    End If
    Set obj = Nothing
    Set rs = Nothing
    Exit Function
ErrExisteDatosSistemaPension:
   Call RaiseError(MyUnhandledError, "NProveedorSistPens:ExisteDatosSistemaPension Method")
End Function
Public Function AplicaRetencionSistemaPension(ByVal psPersCod As String, ByVal pdFecha As Date, ByVal pnMontoBase As Currency) As Boolean
    On Error GoTo ErrAplicaRetencionSistemaPension
    Dim obj As New DProveedorSistPens
    AplicaRetencionSistemaPension = obj.AplicaRetencionSistemaPension(psPersCod, pdFecha, pnMontoBase)
    Set obj = Nothing
    Exit Function
ErrAplicaRetencionSistemaPension:
   Call RaiseError(MyUnhandledError, "NProveedorSistPens:AplicaRetencionSistemaPension Method")
End Function
Public Function RegistrarDatosSistemaPension(ByVal psMovNro As String, ByVal psPersCod As String, ByVal pnTpoSistPens As TipoSistemaPensionProveeedor, _
                                             ByVal psCUSP As String, ByVal psPersCodAFP As String, ByVal pnTpoComision As TipoComisionAFPProveeedor) As Boolean
    On Error GoTo ErrRegistrarDatosSistemaPension
    Dim obj As New DProveedorSistPens
    Dim bTrans As Boolean
    
    obj.BeginTrans
    bTrans = True
    
    obj.InsertarSistemaPension psPersCod, pnTpoSistPens, psMovNro
    If pnTpoSistPens = AFP Then
        obj.InsertarSistemaPensionAFP psPersCod, psCUSP, psPersCodAFP, pnTpoComision
    End If
    
    obj.CommitTrans
    bTrans = False
    RegistrarDatosSistemaPension = True
    
    Set obj = Nothing
    Exit Function
ErrRegistrarDatosSistemaPension:
    RegistrarDatosSistemaPension = False
    If bTrans Then
        obj.RollbackTrans
        Set obj = Nothing
    End If
   Call RaiseError(MyUnhandledError, "NProveedorSistPens:RegistrarDatosSistemaPension Method")
End Function
Public Sub SetDatosRetencionSistPens(ByVal psPersCod As String, ByVal pdFecha As Date, ByVal pnMontoBase As Currency, ByVal pnMoneda As Integer, _
                                        ByRef pnMontoAporte As Currency, ByRef pnMontoSeguroAFP As Currency, ByRef pnMontoComisionAFP As Currency, _
                                        Optional ByRef pnTpoSistPens As TipoSistemaPensionProveeedor = 0)
    On Error GoTo ErrSetDatosRetencionSistPens
    Dim obj As New DProveedorSistPens
    Dim rs As New ADODB.Recordset
    
    Set rs = obj.ObtenerMontosRetencSistPens(psPersCod, pdFecha, pnMontoBase, pnMoneda)
    pnMontoAporte = 0#: pnMontoSeguroAFP = 0#: pnMontoComisionAFP = 0#
    If Not rs.EOF Then
        pnMontoAporte = rs!nMontoAporte
        pnMontoSeguroAFP = rs!nMontoSeguroAFP
        pnMontoComisionAFP = rs!nMontoComisionAFP
        pnTpoSistPens = rs!nTpoSistPens
    End If
    RSClose rs
    Set obj = Nothing
    Exit Sub
ErrSetDatosRetencionSistPens:
   Call RaiseError(MyUnhandledError, "NProveedorSistPens:SetDatosRetencionSistPens Method")
End Sub
Public Function ListaAFP() As ADODB.Recordset
    On Error GoTo ErrListaAFP
    Dim obj As New DProveedorSistPens
    Set ListaAFP = obj.ListaAFP()
    Set obj = Nothing
    Exit Function
ErrListaAFP:
   Call RaiseError(MyUnhandledError, "NProveedorSistPens:ListaAFP Method")
End Function
Public Function ListaProveedorSistemaPension() As ADODB.Recordset
    On Error GoTo ErrListaAFP
    Dim obj As New DProveedorSistPens
    Set ListaProveedorSistemaPension = obj.ListaProveedorSistemaPension()
    Set obj = Nothing
    Exit Function
ErrListaAFP:
   Call RaiseError(MyUnhandledError, "NProveedorSistPens:ListaProveedorSistemaPension Method")
End Function
Public Function EditarDatosSistemaPension(ByVal psPersCod As String, ByVal pnTpoSistPens As TipoSistemaPensionProveeedor, _
                                             ByVal psCUSP As String, ByVal psPersCodAFP As String, ByVal pnTpoComision As TipoComisionAFPProveeedor) As Boolean
    On Error GoTo ErrRegistrarDatosSistemaPension
    Dim obj As New DProveedorSistPens
    Dim bTrans As Boolean
    
    obj.BeginTrans
    bTrans = True
    
    obj.EditarSistemaPension psPersCod, pnTpoSistPens
    obj.EliminaSistemaPensionAFP psPersCod
    If pnTpoSistPens = AFP Then
        obj.InsertarSistemaPensionAFP psPersCod, psCUSP, psPersCodAFP, pnTpoComision
    End If
    
    obj.CommitTrans
    bTrans = False
    EditarDatosSistemaPension = True
    
    Set obj = Nothing
    Exit Function
ErrRegistrarDatosSistemaPension:
    EditarDatosSistemaPension = False
    If bTrans Then
        obj.RollbackTrans
        Set obj = Nothing
    End If
   Call RaiseError(MyUnhandledError, "NProveedorSistPens:EditarDatosSistemaPension Method")
End Function
Public Function ListaTasasAFP() As ADODB.Recordset
    On Error GoTo ErrListaTasasAFP
    Dim obj As New DProveedorSistPens
    Set ListaTasasAFP = obj.ListaTasasAFP()
    Set obj = Nothing
    Exit Function
ErrListaTasasAFP:
   Call RaiseError(MyUnhandledError, "NProveedorSistPens:ListaTasasAFP Method")
End Function
Public Function ActualizaTasasAFP(ByVal psPersCod As String, ByVal pnTasaComiFlujo As Currency, ByVal pnTasaComiSaldo As Currency, ByVal pnTasaSeguro As Currency) As Boolean
    On Error GoTo ErrActualizaTasasAFP
    Dim obj As New DProveedorSistPens
    Dim bTrans As Boolean
    
    obj.BeginTrans
    bTrans = True
    
    obj.ActualizaTasaComisionSistPensAFPProv psPersCod, Flujo, pnTasaComiFlujo
    obj.ActualizaTasaComisionSistPensAFPProv psPersCod, Saldo, pnTasaComiSaldo
    obj.ActualizaTasaSeguroSistPensAFPProv psPersCod, pnTasaSeguro
    
    obj.CommitTrans
    bTrans = False
    ActualizaTasasAFP = True
    Set obj = Nothing
    Exit Function
ErrActualizaTasasAFP:
    ActualizaTasasAFP = False
    If bTrans Then
        obj.RollbackTrans
        Set obj = Nothing
    End If
   Call RaiseError(MyUnhandledError, "NProveedorSistPens:ActualizaTasasAFP Method")
End Function
