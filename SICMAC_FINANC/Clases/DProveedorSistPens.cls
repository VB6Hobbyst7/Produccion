VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DProveedorSistPens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim sCentralCom As String
Dim dbConec As DConecta
Dim psSql As String

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
    Dim oIni As New ClasIni
    sCentralCom = oIni.BaseComunes
    Set dbConec = New DConecta
    If Not dbConec.AbreConexion Then
        Call RaiseError(MyUnhandledError, "DRepFormula:Initialize Method. Error en Conexion a Base de datos")
    End If
   Set oIni = Nothing
End Sub
Private Sub Class_Terminate()
    If Not dbConec Is Nothing Then
        dbConec.CierraConexion
        Set dbConec = Nothing
    End If
End Sub
Public Sub BeginTrans()
    dbConec.BeginTrans
End Sub
Public Sub RollbackTrans()
    dbConec.RollbackTrans
    dbConec.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub
Public Sub CommitTrans()
    dbConec.CommitTrans
End Sub
Public Function ObtieneDatosSistemaPension(ByVal psPersCod As String) As ADODB.Recordset
    On Error GoTo ErrObtieneDatosProveedorSistPension
    psSql = "EXEC stp_sel_ERS1022014_DatosProveedorSistPension '" & psPersCod & "'"
    Set ObtieneDatosSistemaPension = dbConec.CargaRecordSet(psSql)
    Exit Function
ErrObtieneDatosProveedorSistPension:
   Call RaiseError(MyUnhandledError, "DProveedorSistPens:ObtieneDatosProveedorSistPension Method")
End Function
Public Function AplicaRetencionSistemaPension(ByVal psPersCod As String, ByVal pdFecha As Date, ByVal pnMontoBase As Currency) As Boolean
    On Error GoTo ErrObtieneDatosProveedorSistPension
    Dim rs As New ADODB.Recordset
    psSql = "EXEC stp_sel_ERS1022014_AplicaRetencion '" & psPersCod & "','" & Format(pdFecha, "yyyymmdd") & "'," & pnMontoBase
    Set rs = dbConec.CargaRecordSet(psSql)
    AplicaRetencionSistemaPension = False
    If Not rs.EOF Then
        AplicaRetencionSistemaPension = rs!bAplicaRetencion
    End If
    RSClose rs
    Exit Function
ErrObtieneDatosProveedorSistPension:
   Call RaiseError(MyUnhandledError, "DProveedorSistPens:ObtieneDatosProveedorSistPension Method")
End Function
Public Sub InsertarSistemaPension(ByVal psPersCod As String, ByVal pnTpoSistPens As TipoSistemaPensionProveeedor, ByVal psMovNroRegistro As String)
    On Error GoTo ErrInsertarNotasEstadoConfigDet
    psSql = "EXEC stp_ins_ERS1022014_ProveedorSistPens '" & psPersCod & "'," & pnTpoSistPens & ",'" & psMovNroRegistro & "'"
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrInsertarNotasEstadoConfigDet:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertarNotasEstadoConfigDet Method")
End Sub
Public Sub InsertarSistemaPensionAFP(ByVal psPersCod As String, ByVal psCUSP As String, ByVal psPersCodAFP As String, ByVal pnTpoComision As TipoComisionAFPProveeedor)
    On Error GoTo ErrInsertarSistemaPensionAFP
    psSql = "EXEC stp_ins_ERS1022014_ProveedorSistPensAFP '" & psPersCod & "','" & psCUSP & "','" & psPersCodAFP & "'," & pnTpoComision
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrInsertarSistemaPensionAFP:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertarSistemaPensionAFP Method")
End Sub
Public Function ObtenerMontosRetencSistPens(ByVal psPersCod As String, ByVal pdFecha As Date, ByVal pnMontoBase As Currency, ByVal pnMoneda As Integer) As ADODB.Recordset
    On Error GoTo ErrObtenerMontosRetencSistPens
    psSql = "EXEC stp_sel_ERS1022014_ObtieneMontosRetencSistPens '" & psPersCod & "','" & Format(pdFecha, "yyyymmdd") & "'," & pnMontoBase & "," & pnMoneda
    Set ObtenerMontosRetencSistPens = dbConec.CargaRecordSet(psSql)
    Exit Function
ErrObtenerMontosRetencSistPens:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerMontosRetencSistPens Method")
End Function
Public Function ListaAFP() As ADODB.Recordset
    On Error GoTo ErrListaAFP
    psSql = "EXEC stp_sel_ERS1022014_ListaAFP "
    Set ListaAFP = dbConec.CargaRecordSet(psSql)
    Exit Function
ErrListaAFP:
    Call RaiseError(MyUnhandledError, "DRepFormula:ListaAFP Method")
End Function
Public Function ListaProveedorSistemaPension() As ADODB.Recordset
    On Error GoTo ErrListaAFP
    psSql = "EXEC stp_sel_ERS1022014_ListaProveedorSistemaPension"
    Set ListaProveedorSistemaPension = dbConec.CargaRecordSet(psSql)
    Exit Function
ErrListaAFP:
    Call RaiseError(MyUnhandledError, "DRepFormula:ListaProveedorSistemaPension Method")
End Function
Public Sub EditarSistemaPension(ByVal psPersCod As String, ByVal pnTpoSistPens As TipoSistemaPensionProveeedor)
    On Error GoTo ErrEditarSistemaPension
    psSql = "EXEC stp_upd_ERS1022014_ProveedorSistPens '" & psPersCod & "'," & pnTpoSistPens
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrEditarSistemaPension:
    Call RaiseError(MyUnhandledError, "DRepFormula:EditarSistemaPension Method")
End Sub
Public Sub EditarSistemaPensionAFP(ByVal psPersCod As String, ByVal psCUSP As String, ByVal psPersCodAFP As String, ByVal pnTpoComision As TipoComisionAFPProveeedor)
    On Error GoTo ErrEditarSistemaPensionAFP
    psSql = "EXEC stp_upd_ERS1022014_ProveedorSistPensAFP '" & psPersCod & "','" & psCUSP & "','" & psPersCodAFP & "'," & pnTpoComision
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrEditarSistemaPensionAFP:
    Call RaiseError(MyUnhandledError, "DRepFormula:EditarSistemaPensionAFP Method")
End Sub
Public Sub EliminaSistemaPensionAFP(ByVal psPersCod As String)
    On Error GoTo ErrEliminaSistemaPensionAFP
    psSql = "EXEC stp_del_ERS1022014_ProveedorSistPensAFP '" & psPersCod & "'"
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrEliminaSistemaPensionAFP:
    Call RaiseError(MyUnhandledError, "DRepFormula:EliminaSistemaPensionAFP Method")
End Sub
Public Function ListaTasasAFP() As ADODB.Recordset
    On Error GoTo ErrListaAFP
    psSql = "EXEC stp_sel_ERS1022014_ListaTasasAFP"
    Set ListaTasasAFP = dbConec.CargaRecordSet(psSql)
    Exit Function
ErrListaAFP:
    Call RaiseError(MyUnhandledError, "DRepFormula:ListaTasasAFP Method")
End Function
Public Sub ActualizaTasaComisionSistPensAFPProv(ByVal psPersCod As String, ByVal pnTpoComision As TipoComisionAFPProveeedor, ByVal pnTasa As Currency)
    On Error GoTo ErrActualizaTasaComisionSistPensAFPProv
    psSql = "EXEC stp_upd_ERS1022014_TasaComisionSistPensAFPProv '" & psPersCod & "'," & pnTpoComision & "," & pnTasa
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrActualizaTasaComisionSistPensAFPProv:
    Call RaiseError(MyUnhandledError, "DRepFormula:ActualizaTasaComisionSistPensAFPProv Method")
End Sub
Public Sub ActualizaTasaSeguroSistPensAFPProv(ByVal psPersCod As String, ByVal pnTasa As Currency)
    On Error GoTo ErrActualizaTasaSeguroSistPensAFPProv
    psSql = "EXEC stp_upd_ERS1022014_TasaSeguroSistPensAFPProv '" & psPersCod & "'," & pnTasa
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrActualizaTasaSeguroSistPensAFPProv:
    Call RaiseError(MyUnhandledError, "DRepFormula:ActualizaTasaSeguroSistPensAFPProv Method")
End Sub

