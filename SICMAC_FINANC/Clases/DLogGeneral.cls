VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsCentralPer As String
Dim vsCentralCom As String
Dim vsCentralImg As String
Dim sSql As String

Public Enum TpoCargaUsu
    UsuUnRegistro = 0
    UsuTodosArea = 1
    UsuTodos = 2
End Enum

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
    Dim oIni As ClasIni
    
    Set oIni = New ClasIni
        vsCentralPer = oIni.BasePersonas
        vsCentralCom = oIni.BaseComunes
        vsCentralImg = oIni.BaseImagenes
    Set oIni = Nothing
End Sub

'Devuelve un String con un Nro de Mov
Public Function GeneraMov(ByVal pdFecsis As Date, ByVal psCodCmac As String, _
    ByVal psCodAge As String, ByVal psCodUser As String) As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    If oConec.AbreConexion() Then
        GeneraMov = Format(pdFecsis & " " & oConec.GetHoraServer, "yyyymmddhhmmss") & psCodCmac & Right(psCodAge, 2) & "00" & psCodUser
    Else
        GeneraMov = ""
    End If
    Set oConec = Nothing
End Function

'Devuelve un RecordSet con las Areas
Public Function CargaArea() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        'Todos las Areas
        sSql = "SELECT a.cAreaCod, a.cAreaDescripcion, a.cAreaEstruc " & _
            " FROM  Areas A " & _
            " ORDER BY a.cAreaDescripcion "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaArea = rs
End Function

'Devuelve un RecordSet con la relación usuario-area
Public Function CargaUsuario(ByVal pnTpoCarga As TpoCargaUsu, ByVal psUserArea As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If pnTpoCarga = UsuUnRegistro Then
            'Un usuario
            sSql = "SELECT rh.cAreaCod , a.cAreaDescripcion, p.cPersNombre, a.cAreaEstruc " & _
                " FROM  RRHH RH INNER JOIN Areas A ON rh.cAreaCod = a.cAreaCod " & _
                "       JOIN " & vsCentralPer & "PERSONA P ON P.cPersCod = RH.cPersCod  " & _
                " WHERE  rh.cUser = '" & psUserArea & "' "
        ElseIf pnTpoCarga = UsuTodosArea Then
            'Todos los usuarios por area
            sSql = "SELECT rh.cPersCod, rh.cUser, p.cPersNombre, p.cPersEstado " & _
                " FROM  RRHH RH INNER JOIN Areas A ON rh.cAreaCod = a.cAreaCod and rh.cAreaCod = '" & psUserArea & "'" & _
                "       JOIN " & vsCentralPer & "PERSONA P ON P.cPersCod = RH.cPersCod  " & _
                " ORDER BY p.cPersNombre "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaUsuario = rs
End Function

'Devuelve un RecordSet de acuerdo a la constante deseada
Public Function CargaConstante(ByVal psConstante As String, Optional ByVal pbPrimeroValor As Boolean = True) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = IIf(pbPrimeroValor, "SELECT c.nConsValor, c.cConsDescripcion ", "SELECT c.cConsDescripcion, c.nConsValor ")
        sSql = sSql & _
            " FROM Constante C " & _
            " WHERE nCONSCOD = '" & psConstante & "' AND Not nConsValor = '" & psConstante & "'" & _
            " ORDER BY c.nConsValor "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaConstante = rs
End Function


'Devuelve un RecordSet con el Area siguiente para el trámite
Public Function CargaAreaSuperior(ByVal psAreaStru As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = " SELECT cAreaCod, cAreaDescripcion " & _
            " FROM Areas " & _
            " WHERE cAreaEstruc = substring('" & psAreaStru & "',1,len('" & psAreaStru & "') - 2) "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaAreaSuperior = rs
End Function

Public Function ExisteOperacionCtaBS(ByVal psOpeCod As String) As Boolean
Dim prs As ADODB.Recordset
Set prs = CargaOperacionCtaBS(psOpeCod)
If prs.EOF Then
    ExisteOperacionCtaBS = False
Else
    ExisteOperacionCtaBS = True
End If
RSClose prs
End Function

Public Function CargaOperacionCtaBS(ByVal psOpeCod As String) As ADODB.Recordset
Dim sSql As String
Dim prs  As ADODB.Recordset
Set prs = New ADODB.Recordset
    Dim oConec As DConecta
    Set oConec = New DConecta
    If oConec.AbreConexion() Then
        sSql = "SELECT oc.cCtaContCod, c.cCtaContDesc, b.cBSCod, b.cBSDescripcion " _
             & "FROM   CtaBS oc LEFT JOIN CtaCont c ON c.cCtaContCod = oc.cCtaContCod " _
             & "                JOIN BienesServicios b ON b.cBSCod = oc.cObjetoCod " _
             & "WHERE  oc.cOpeCod = '" & psOpeCod & "' ORDER BY b.cBSCod "
        Set CargaOperacionCtaBS = oConec.CargaRecordSet(sSql, adLockOptimistic)
    End If
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'EJVG20131113 ***
Public Function ListaComprobantesxProvision(ByVal pnMoneda As Moneda) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim sSql As String
    sSql = "Exec stp_sel_ERS0622013_ListaComprobantesxProvision '" & pnMoneda & "'"
    oConec.AbreConexion
    Set ListaComprobantesxProvision = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function ComprobanteDet(ByVal pnMovNro As Long, ByVal pnMoneda As Moneda) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim sSql As String
    sSql = "Exec stp_sel_ERS0622013_ComprobanteDet " & pnMovNro & "," & pnMoneda
    oConec.AbreConexion
    Set ComprobanteDet = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function ListaIFisxPagoProveedor() As ADODB.Recordset
    Dim oConec As New DConecta
    Dim sSql As String
    sSql = "Exec stp_sel_ERS0622013_ListaIFisxPagoProveedor"
    oConec.AbreConexion
    Set ListaIFisxPagoProveedor = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function ListaPenalidadxRegistroVenta() As ADODB.Recordset
    Dim oConec As New DConecta
    Dim sSql As String
    sSql = "Exec stp_sel_ERS0622013_ListaPenalidadxRegistroVenta"
    oConec.AbreConexion
    Set ListaPenalidadxRegistroVenta = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'END EJVG *******

'PASI 20140314 TI-ERS001-2014*********************************
Public Function ListaProvisionesdeIntangibles() As ADODB.Recordset
    Dim oConec As New DConecta
    Dim sSql As String
    sSql = "exec stp_sel_ERS0012014_ListaProvisionesdeIntangibles"
    oConec.AbreConexion
    Set ListaProvisionesdeIntangibles = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function ProvisionIntangibleDetalle(ByVal pnMovNro As Long) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim sSql As String
    sSql = "exec stp_sel_ERS0012014_ProvisionIntangibleDet " & pnMovNro
    oConec.AbreConexion
    Set ProvisionIntangibleDetalle = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'PASI END****************************************************
'EJVG20140724 ***
Public Function ObtieneAcumuladoPeriodoRetencSistPens(ByVal psPersCod As String) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim sSql As String
    sSql = "exec stp_sel_ERS1022014_ObtieneAcumuladoPeriodoRetencSistPens '" & psPersCod & "'"
    oConec.AbreConexion
    Set ObtieneAcumuladoPeriodoRetencSistPens = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'END EJVG *******
