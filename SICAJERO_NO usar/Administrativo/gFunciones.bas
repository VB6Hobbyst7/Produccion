Attribute VB_Name = "gFunciones"
Option Explicit
Global C As ADODB.Connection
Global MatCol0200(42) As Integer
Global gsCodAge As String
Global gsNomAge As String
Global gsCodCiudad As String
Global gsCodUser As String
Global gsNomUser As String
Global gsPass As String
Global gsBIN As String
Global gnTipoPinPad As Integer
Global gnPinPadPuerto As Integer
Global gsNomMaquinaUsu As String
Global gdFecSis As Date
Global LstUsuarios() As String

'**DAOR 20100526 *************************
Global gsServerName As String
Global gsDatabaseName As String
'*****************************************

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

 
Public Function ExisteTarjeta(ByVal psNumtarjeta As String) As Boolean
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 20)
    Prm.Value = psNumtarjeta
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnRes", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva 'AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_VerificaExtiteTarjeta"
    Cmd.Execute
    
    If Cmd.Parameters(1).Value > 0 Then
        ExisteTarjeta = True
    Else
        ExisteTarjeta = False
    End If
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    
    Set Cmd = Nothing
    Set Prm = Nothing

End Function



 Public Sub CargaControlGrupos(ByVal psDominio As String)
Dim container As IADsContainer
Dim containername As String
Dim group As IADsGroup
Dim nNumGrupos As Integer

    Set container = GetObject("WinNT://" & psDominio)
    container.Filter = Array("Group")
    nNumGrupos = 0
    ReDim LstUsuarios(nNumGrupos)
    For Each group In container
        If UCase(Left(group.Name, 5)) = "GRUPO" Or UCase(Left(group.Name, 3)) = "GG_" Then
            ReDim Preserve LstUsuarios(nNumGrupos + 1)
            LstUsuarios(nNumGrupos) = group.Name
            nNumGrupos = nNumGrupos + 1
        End If
    Next

    
End Sub
 
Public Function RecuperaPVV(ByVal pPAN As String) As String
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta

    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cNumTarj", adVarChar, adParamInput, 50, pPAN)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psPVV", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva 'AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaPVV"
    
    Cmd.Execute
    
    RecuperaPVV = Cmd.Parameters(1).Value
        
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
        
    Set Cmd = Nothing
    Set Prm = Nothing
    
End Function

Public Sub ActualizaPVV(ByVal pssPVV As String, ByVal psPan As String)
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta

    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psPan", adVarChar, adParamInput, 16, psPan)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psPVV", adVarChar, adParamInput, 10, pssPVV)
    Cmd.Parameters.Append Prm

    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ActualizaPVV"
    
    Cmd.Execute
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
 
    Set Cmd = Nothing
    Set Prm = Nothing
    
End Sub

Public Function GetMaquinaUsuario() As String   'Para obtener la Maquina del Usuario
    Dim buffMaq As String
    Dim lSizeMaq As Long
    buffMaq = Space(255)
    lSizeMaq = Len(buffMaq)
    GetComputerName buffMaq, lSizeMaq
    GetMaquinaUsuario = Trim(Left$(buffMaq, lSizeMaq))
End Function

Public Sub RecuperaConfigPinPad()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cNomPC", adVarChar, adParamInput, 50, gsNomMaquinaUsu)
    Cmd.Parameters.Append Prm
     
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTipoPinPad", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
     
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnNumPuerto", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaDatosPinPad"
    Cmd.Execute
    
    gnTipoPinPad = Cmd.Parameters(1).Value
    gnPinPadPuerto = Cmd.Parameters(2).Value

    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing
End Sub

Public Sub UltimoRetiroAtm(ByVal psNumtarjeta As String, ByVal psCtaCod As String, _
            ByRef psFecha As String, ByRef pnMonto As Double)
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 20)
    Prm.Value = psNumtarjeta
    Cmd.Parameters.Append Prm
     
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 20)
    Prm.Value = psCtaCod
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pdFecha", adVarChar, adParamOutput, 30)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMonto", adCurrency, adParamOutput)
    Cmd.Parameters.Append Prm

    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_UltimoRetiroATM"
    Cmd.Execute
    
    psFecha = IIf(IsNull(Cmd.Parameters(2).Value), "", Cmd.Parameters(2).Value)
    pnMonto = Format(IIf(IsNull(Cmd.Parameters(3).Value), 0, Cmd.Parameters(3).Value), "#,0.00")

    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing

End Sub


'

Public Sub RecuperaCantidadDEUNRangoDETarjetasSalida(ByVal psRanIni As String, ByVal psRanFin As String, ByRef nValorCant As Integer)
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psRangoINI", adVarChar, adParamInput, 20)
    Prm.Value = psRanIni
    Cmd.Parameters.Append Prm
         
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psRangoFIN", adVarChar, adParamInput, 20)
    Prm.Value = psRanFin
    Cmd.Parameters.Append Prm
         
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaCantidadDEUNRangoDETarjetasSalida "
        
    Set R = Cmd.Execute

    nValorCant = IIf(IsNull(R!nNumTarj), 0, R!nNumTarj)
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing

End Sub


Public Sub RecuperaCantidadDEUNRangoDETarjetasDevol(ByVal psRanIni As String, ByVal psRanFin As String, ByRef nValorCant As Integer)
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psRangoINI", adVarChar, adParamInput, 20)
    Prm.Value = psRanIni
    Cmd.Parameters.Append Prm
         
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psRangoFIN", adVarChar, adParamInput, 20)
    Prm.Value = psRanFin
    Cmd.Parameters.Append Prm
         
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCodAge", adInteger, adParamInput)
    Prm.Value = CInt(gsCodAge)
    Cmd.Parameters.Append Prm
                      
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaCantidadDEUNRangoDETarjetasDevol  "
    
    Set R = Cmd.Execute

    nValorCant = IIf(IsNull(R!nNumTarj), 0, R!nNumTarj)
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    
    Set Cmd = Nothing
    Set Prm = Nothing

End Sub

Public Sub RecuperaRangosDETarjetasEmitidas(ByVal pnCant As Integer, _
    ByRef psRanIni As String, ByRef psRanFin As String, ByRef nValorCant As Integer)
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCantTarj", adInteger, adParamInput)
    Prm.Value = pnCant
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaRangosDETarjetasEmitidas "
    
    Set R = Cmd.Execute
    
    psRanIni = IIf(IsNull(R!sMinTarj), "", R!sMinTarj)
    psRanFin = IIf(IsNull(R!sMaxTarj), "", R!sMaxTarj)
    nValorCant = IIf(IsNull(R!nNumTarj), 0, R!nNumTarj)
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing

End Sub

Public Sub RecuperaRangosDETarjetasRemesadasDEAge(ByVal pnCant As Integer, _
    ByRef psRanIni As String, ByRef psRanFin As String, ByRef nValorCant As Integer)
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCantTarj", adInteger, adParamInput)
    Prm.Value = pnCant
    Cmd.Parameters.Append Prm
         
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCodAge", adInteger, adParamInput)
    Prm.Value = CInt(gsCodAge)
    Cmd.Parameters.Append Prm
                  
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaRangosDETarjetasRemesadasDEAge "
        
    Set R = Cmd.Execute
    
    psRanIni = IIf(IsNull(R!sMinTarj), "", R!sMinTarj)
    psRanFin = IIf(IsNull(R!sMaxTarj), "", R!sMaxTarj)
    nValorCant = IIf(IsNull(R!nNumTarj), 0, R!nNumTarj)
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing


End Sub


Public Sub RecuperaRangosDETarjetasIngresadas(ByVal pnCant As Integer, _
    ByRef psRanIni As String, ByRef psRanFin As String, ByRef nValorCant As Integer)
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCantTarj", adInteger, adParamInput)
    Prm.Value = pnCant
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaRangosDETarjetasIngresadas "
    
    Set R = Cmd.Execute
    
    psRanIni = IIf(IsNull(R!sMinTarj), "", R!sMinTarj)
    psRanFin = IIf(IsNull(R!sMaxTarj), "", R!sMaxTarj)
    nValorCant = IIf(IsNull(R!nNumTarj), 0, R!nNumTarj)
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing

End Sub

Public Function RecuperaEstadoDETarjeta(ByVal psNumtarjeta As String) As Integer
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarjeta", adVarChar, adParamInput, 20)
    Prm.Value = psNumtarjeta
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCondicion", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm

    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaEstadoDETarjeta"
    Cmd.Execute
    
    RecuperaEstadoDETarjeta = Cmd.Parameters(1).Value
   
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing
    
End Function

Public Sub RecuperaDatosDETarjetas(ByVal psNumtarjeta As String, ByRef pnCond As Integer, _
    ByRef pnRetenTar As Integer, ByRef pnNOOperMonExt As Integer, ByRef nSuspOper As Integer, _
    ByRef dFecVenc As Date, ByRef psDescEstado As String)
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@PAN", adVarChar, adParamInput, 20)
    Prm.Value = psNumtarjeta
    Cmd.Parameters.Append Prm
     
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCondicion", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nRetenerTarjeta", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nNOOperMonExt", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSuspOper", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dfecVenc", adDBDate, adParamOutput)
    Cmd.Parameters.Append Prm
    
     Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psDescEstado", adVarChar, adParamOutput, 100)
    Cmd.Parameters.Append Prm
            
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaDatosTarjeta"
    Cmd.Execute
    
    pnCond = Cmd.Parameters(1).Value
    pnRetenTar = Cmd.Parameters(2).Value
    pnNOOperMonExt = Cmd.Parameters(3).Value
    nSuspOper = Cmd.Parameters(4).Value
    dFecVenc = Cmd.Parameters(5).Value
    psDescEstado = Cmd.Parameters(6).Value
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing
    
End Sub

Public Sub ConsultaTarjetaCuenta(ByVal psNumtarjeta As String, ByVal psCtaCod As String, _
        ByRef nPrio As Integer, ByRef pnRelacionada As Integer, ByRef pnConsulta As Integer, ByRef pnRetiro As Integer)
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 20)
    Prm.Value = psNumtarjeta
    Cmd.Parameters.Append Prm
     
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 20)
    Prm.Value = psCtaCod
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnPrior", adSmallInt, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnRelacionada", adSmallInt, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@bConsulta", adSmallInt, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@bRetiro", adSmallInt, adParamOutput)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ConsultaTarjetaCuenta"
    Cmd.Execute
    

    nPrio = IIf(IsNull(Cmd.Parameters(2).Value), 1, Cmd.Parameters(2).Value)
    pnRelacionada = IIf(IsNull(Cmd.Parameters(3).Value), 0, Cmd.Parameters(3).Value)
    pnConsulta = IIf(IsNull(Cmd.Parameters(4).Value), 0, Cmd.Parameters(4).Value)
    pnRetiro = IIf(IsNull(Cmd.Parameters(5).Value), 0, Cmd.Parameters(5).Value)
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing

End Sub

Public Function CuentaVinculada(ByVal psNumtarjeta As String, ByVal psCtaCod As String) As Boolean
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarjeta", adVarChar, adParamInput, 20)
    Prm.Value = psNumtarjeta
    Cmd.Parameters.Append Prm
     
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCodCta", adVarChar, adParamInput, 20)
    Prm.Value = psCtaCod
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnResp", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm

    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_VerificaCuentaVinculada"
    Cmd.Execute
    
    If Cmd.Parameters(2).Value = 1 Then
         CuentaVinculada = True
    Else
         CuentaVinculada = False
    End If
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing

End Function

Public Function TarjetaActiva(ByVal psNumtarjeta As String) As Boolean
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 20)
    Prm.Value = psNumtarjeta
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnRes", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm

    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_VerificaTarjetaActiva"
    Cmd.Execute
    
    If Cmd.Parameters(1).Value = 1 Then
         TarjetaActiva = True
    Else
         TarjetaActiva = False
    End If
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing


End Function

Public Function ExisteTarjetaEmitida(ByVal psNumtarjeta As String) As Boolean
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim loConec As New DConecta

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 20)
    Prm.Value = psNumtarjeta
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnRes", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm

    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_VerificaTarjetaEMITIDA"
    Cmd.Execute
    
    If Cmd.Parameters(1).Value > 0 Then
         ExisteTarjetaEmitida = True
    Else
         ExisteTarjetaEmitida = False
    End If
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing

End Function


'Public Function LeerTarjeta()
'
'End Function

Public Sub DefinirColumnas0200()
    MatCol0200(0) = 4
    MatCol0200(1) = 1
    MatCol0200(2) = 4
    MatCol0200(3) = 1
    MatCol0200(4) = 6
    MatCol0200(5) = 1
    MatCol0200(6) = 1
    MatCol0200(7) = 3
    MatCol0200(8) = 3
    MatCol0200(9) = 3
    MatCol0200(10) = 13
    MatCol0200(11) = 2
    MatCol0200(12) = 13
    MatCol0200(13) = 13
    MatCol0200(14) = 2
    MatCol0200(15) = 13
    MatCol0200(16) = 13
    MatCol0200(17) = 2
    MatCol0200(18) = 13
    MatCol0200(19) = 13
    MatCol0200(20) = 2
    MatCol0200(21) = 13
    MatCol0200(22) = 15
    MatCol0200(23) = 15
    MatCol0200(24) = 15
    MatCol0200(25) = 5
    MatCol0200(26) = 25
    MatCol0200(27) = 25
    MatCol0200(28) = 15
    MatCol0200(29) = 15
    MatCol0200(30) = 20
    MatCol0200(31) = 10
    MatCol0200(32) = 14
    MatCol0200(33) = 14
    MatCol0200(34) = 2
    MatCol0200(35) = 1
    MatCol0200(36) = 6
    MatCol0200(37) = 1
    MatCol0200(38) = 9
    MatCol0200(39) = 40
    MatCol0200(40) = 1
    MatCol0200(41) = 26

    
End Sub

Public Function DigitoChequeo(ByVal psNum As String) As String
Dim N As String
Dim nValor As Integer
Dim i As Integer
Dim Dig1 As String
Dim Dig2 As String
Dim nSum As Integer
N = psNum

nSum = 0
For i = 1 To 15
    nValor = CInt(Mid(N, i, 1))
    'Si es par multiplica por 1
    If i Mod 2 = 0 Then
        nValor = nValor * 1
    Else
    'Multiplica por 2
        nValor = nValor * 2
        If nValor >= 10 Then
            Dig1 = Mid(Trim(Str(nValor)), 1, 1)
            Dig2 = Mid(Trim(Str(nValor)), 2, 1)
            nValor = CInt(Dig1) + CInt(Dig2)
        End If
        
    End If
    nSum = nSum + nValor
Next i

DigitoChequeo = Trim(Str(((nSum - CInt(Right(Trim(Str(nSum)), 1))) + 10) - nSum))
If DigitoChequeo = 10 Then DigitoChequeo = 0

End Function
Public Function GeneraTrama0200(ByVal pdFecAfil As Date, ByVal psNombre As String, ByVal psApePat As String, _
    ByVal psApeMat As String, ByVal psDirecc As String, ByVal psTelef As String, ByVal psSexo As String, _
    ByVal pdFecNac As Date, ByVal psEstCiv As String, ByVal psAge As String, psDNI As String, _
    ByVal psCiudad As String) As String
    
Dim i As Integer
Dim sCad As String

    sCad = ""
    
        sCad = sCad & "0200"
        sCad = sCad & "0"
        sCad = sCad & "0000"
        sCad = sCad & "N"
        sCad = sCad & Format(pdFecAfil, "yymmdd")
        sCad = sCad & "Y"
        sCad = sCad & "Y"
        sCad = sCad & "SPN"
        sCad = sCad & "604"
        sCad = sCad & "840"
        sCad = sCad & "C000000000000"
        sCad = sCad & "00"
        sCad = sCad & "C000000000000"
        sCad = sCad & "C000000000000"
        sCad = sCad & "00"
        sCad = sCad & "C000000000000"
        sCad = sCad & "C000000000000"
        sCad = sCad & "00"
        sCad = sCad & "C000000000000"
        sCad = sCad & "C000000000000"
        sCad = sCad & "00"
        sCad = sCad & "C000000000000"
        sCad = sCad & Left(psNombre & Space(15), 15)
        sCad = sCad & Left(psApePat & Space(15), 15)
        sCad = sCad & Left(psApeMat & Space(15), 15)
        sCad = sCad & Space(5)
        sCad = sCad & Left(psDirecc & Space(25), 25)
        sCad = sCad & Space(25)
        sCad = sCad & Left(psCiudad & Space(15), 15)
        sCad = sCad & Left(psCiudad & Space(15), 15)
        sCad = sCad & Left("PERU" & Space(20), 20)
        sCad = sCad & Space(10)
        sCad = sCad & Left(psTelef & Space(14), 14)
        sCad = sCad & Space(14)
        sCad = sCad & "01"
        sCad = sCad & psSexo
        sCad = sCad & Format(pdFecNac, "yymmdd")
        sCad = sCad & psEstCiv
        sCad = sCad & Right("000000000" & psAge, 9)
        sCad = sCad & Left("LE" & psDNI & Space(40), 40)
        sCad = sCad & "N"
        sCad = sCad & Left("INNOMINADA" & Space(26), 26)
        
    GeneraTrama0200 = sCad
    
End Function

Public Function GeneraTrama0201() As String
    
Dim i As Integer
Dim sCad As String

        sCad = ""
        sCad = sCad & "0201"
        sCad = sCad & "0"
        sCad = sCad & "0000"
        sCad = sCad & "N"
                
    GeneraTrama0201 = sCad
    
End Function

Public Function GeneraTrama0203() As String
    
Dim i As Integer
Dim sCad As String

    sCad = ""
    
        sCad = sCad & "0203"
        sCad = sCad & "0"
        sCad = sCad & "0000"
        sCad = sCad & "N"
        
        
    GeneraTrama0203 = sCad
    
End Function

Public Function GeneraTrama0224() As String
    
Dim i As Integer
Dim sCad As String

    sCad = ""
    
        sCad = sCad & "0224"
        sCad = sCad & "0"
        sCad = sCad & "0000"
        sCad = sCad & "N"
        
        
    GeneraTrama0224 = sCad
    
End Function

Public Function GeneraTrama0223() As String
    
Dim i As Integer
Dim sCad As String

    sCad = ""
    
        sCad = sCad & "0223"
        sCad = sCad & "0"
        sCad = sCad & "0000"
        sCad = sCad & "N"
        
        
    GeneraTrama0223 = sCad
    
End Function

Public Function RecuperaValorXML(ByVal pINXml As String, ByVal psEtiqueta As String)
Dim i As Integer
Dim sCadTempo As String
Dim sCadValor As String
Dim bIniCad As Boolean

    i = 1
    Do While i <= Len(pINXml)
        
        Do While Mid(pINXml, i, 1) <> "<" And i <= Len(pINXml)
            i = i + 1
        Loop
        i = i + 1
        sCadTempo = ""
        Do While Mid(pINXml, i, 1) <> " " And Mid(pINXml, i, 1) <> ">" And i <= Len(pINXml)
            sCadTempo = sCadTempo & Mid(pINXml, i, 1)
            i = i + 1
        Loop
        
        If UCase(Trim(sCadTempo)) = UCase(Trim(psEtiqueta)) Then
                Do While Mid(pINXml, i, 1) <> "=" And Mid(pINXml, i, 1) <> ">" And i <= Len(pINXml)
                    i = i + 1
                Loop
                i = i + 1
                sCadValor = ""
                Do While Mid(pINXml, i, 1) <> "/" And Mid(pINXml, i, 1) <> ">" And i <= Len(pINXml) And Mid(pINXml, i, 1) <> "<"
                    sCadValor = sCadValor & Mid(pINXml, i, 1)
                    i = i + 1
                Loop
                sCadValor = Trim(Replace(sCadValor, """", ""))
                Exit Do
        End If
        
    Loop
    
    RecuperaValorXML = sCadValor

End Function

Public Function InstanciaProxy(ByVal TipoTxn As String, ByVal PAN As String, ByVal Trama As String, ByRef ADD_DATA_RESP As String, _
    Optional ByVal psTrack2 As String = "", Optional ByVal psClave As String = "", Optional ByVal psClaveNew As String = "") As String
    
Dim RESP_CODE As String
Dim TRACE As Integer
Dim Result As Long

    'Result = ubaobject.TXNADM_JPAU(gsCodUser, gsPass, TipoTxn, PAN, psTrack2, Trama, psClave, psClaveNew, RESP_CODE, TRACE, ADD_DATA_RESP)
    
    'InstanciaProxy = ubaobject.GetResponse("UBASOFT2007")
    InstanciaProxy = RESP_CODE

    'Set ubaobject = Nothing
    
End Function
Public Function GeneraTrama0216(ByVal pdFecAfil As Date, ByVal psNombre As String, ByVal psApePat As String, _
    ByVal psApeMat As String, ByVal psDirecc As String, ByVal psTelef As String, ByVal psSexo As String, _
    ByVal pdFecNac As Date, ByVal psEstCiv As String, ByVal psAge As String, psDNI As String, _
    ByVal psCiudad As String) As String
    
Dim i As Integer
Dim sCad As String

    sCad = ""
    
        sCad = sCad & "0216"
        sCad = sCad & "0"
        sCad = sCad & "0000"
        sCad = sCad & "N"
        sCad = sCad & "Y"
        sCad = sCad & "Y"
        sCad = sCad & "SPN"
        sCad = sCad & Left(psNombre & Space(15), 15)
        sCad = sCad & Left(psApePat & Space(15), 15)
        sCad = sCad & Left(psApeMat & Space(15), 15)
        sCad = sCad & Space(5)
        sCad = sCad & Left(psDirecc & Space(25), 25)
        sCad = sCad & Space(25)
        sCad = sCad & Left(psCiudad & Space(15), 15)
        sCad = sCad & Left(psCiudad & Space(15), 15)
        sCad = sCad & Left("PERU" & Space(20), 20)
        sCad = sCad & Space(10)
        sCad = sCad & Left(psTelef & Space(14), 14)
        sCad = sCad & Space(14)
        sCad = sCad & "01"
        sCad = sCad & psSexo
        sCad = sCad & Format(pdFecNac, "yymmdd")
        sCad = sCad & psEstCiv
        sCad = sCad & Right("000000000" & psAge, 9)
        sCad = sCad & Left("LE" & psDNI & Space(40), 40)
        sCad = sCad & "N"
        sCad = sCad & Left("INNOMINADA" & Space(26), 26)
        
    GeneraTrama0216 = sCad
    
End Function

Public Function GeneraTrama0220(ByVal psCtaCod As String, ByVal psBIN As String, _
    ByVal psAge As String, ByVal psMoneda As String) As String
    
Dim sCad As String

        sCad = ""
    
        sCad = sCad & "0220"
        sCad = sCad & "0"
        sCad = sCad & "0000"
        sCad = sCad & "N"
        sCad = sCad & Mid(psCtaCod, 6, 13) & "0" & Left(psBIN & Space(11), 11) & Right("000" & psAge, 3)
        sCad = sCad & "1"
        sCad = sCad & IIf(psMoneda = "1", "604", "840")
        sCad = sCad & "000"
        sCad = sCad & "C000000000000"
        sCad = sCad & "00"
        sCad = sCad & "C000000000000"
        sCad = sCad & "C000000000000"
        sCad = sCad & "00"
        sCad = sCad & "C000000000000"
        sCad = sCad & "C000000000000"
        sCad = sCad & "00"
        sCad = sCad & "C000000000000"
        sCad = sCad & "C000000000000"
        sCad = sCad & "00"
        sCad = sCad & "C000000000000"
        sCad = sCad & "Y"
        sCad = sCad & "C000000000000"
                      
    GeneraTrama0220 = sCad
    
End Function

Public Function GeneraTrama0227(ByVal psCtaCod As String, ByVal psBIN As String, _
    ByVal psAge As String) As String
    
Dim sCad As String

        sCad = ""
    
        sCad = sCad & "0227"
        sCad = sCad & "0"
        sCad = sCad & "0000"
        sCad = sCad & "N"
        sCad = sCad & Mid(psCtaCod, 6, 13) & "0" & Left(psBIN & Space(11), 11) & Right("000" & psAge, 3)
                      
    GeneraTrama0227 = sCad
    
End Function
    
'Public Function GeneraTrama0205(ByVal psCtaCod As String, ByVal psBIN As String, _
'    ByVal psAge As String) As String
'
'Dim sCad As String
'
'        sCad = ""
'        sCad = sCad & "0205"
'        sCad = sCad & "0"
'        sCad = sCad & "0000"
'        sCad = sCad & "N"
'        sCad = sCad & Mid(psCtaCod, 6, 13) & "0" & Left(psBIN & Space(11), 11) & Right("000" & psAge, 3)
'
'    GeneraTrama0205 = sCad
'
'End Function
    
Public Function GeneraTrama0221(ByVal psCtaCod As String, ByVal psBIN As String, _
    ByVal psAge As String) As String
    
Dim sCad As String

        sCad = ""
    
        sCad = sCad & "0221"
        sCad = sCad & "0"
        sCad = sCad & "0000"
        sCad = sCad & "N"
        sCad = sCad & Mid(psCtaCod, 6, 13) & "0" & Left(psBIN & Space(11), 11) & Right("000" & psAge, 3)
    GeneraTrama0221 = sCad
    
End Function
    
Public Function GeneraTrama0222(ByVal pbEstadoNew As String) As String
    
Dim sCad As String

        sCad = ""
    
        sCad = sCad & "0222"
        sCad = sCad & "0"
        sCad = sCad & "0000"
        sCad = sCad & "N"
        'H -bloqueada C - Cancelada
        sCad = sCad & pbEstadoNew
                      
    GeneraTrama0222 = sCad
    
End Function
    
'LSDO 20080612 SE AGREGO LOS CAMPOS DE LINEAS 1 Y 2 PARA LA DESCRIPCION DE LA CUENTA EN EL ATM
Public Function GeneraTrama0204(ByVal psCtaCod As String, ByVal psBIN As String, _
    ByVal psAge As String, ByVal psCtaDef As String, ByVal psMoneda As String, _
    ByVal psPerCons As String, ByVal psPerRet As String, ByVal psPerDep As String, _
    ByVal psPerTraHacia As String, ByVal psPerTraDesde As String, ByVal psDesc1 As String, ByVal psDesc2 As String) As String
    
Dim sCad As String

        sCad = ""
    
        sCad = sCad & "0204"
        sCad = sCad & "0"
        sCad = sCad & "0000"
        sCad = sCad & "N"
        sCad = sCad & Left(Mid(psCtaCod, 6, 13) & "0" & Left(psBIN & Space(11), 11) & Right("000" & psAge, 3) & Space(28), 28)
        sCad = sCad & psCtaDef
        sCad = sCad & psDesc1 & Space(15 - Len(psDesc1)) '"12345"
        sCad = sCad & psDesc2 & Space(15 - Len(psDesc2)) '"54321"
        sCad = sCad & IIf(psMoneda = "1", "604", "840")
        sCad = sCad & psPerCons
        sCad = sCad & psPerRet
        sCad = sCad & psPerDep
        sCad = sCad & psPerTraHacia
        sCad = sCad & psPerTraDesde
                
    GeneraTrama0204 = sCad
    
End Function
    
Public Function GeneraTrama0225(ByVal psCtaCod As String, ByVal psBIN As String, _
    ByVal psAge As String, ByVal psCtaDef As String, ByVal psMoneda As String, _
    ByVal psPerCons As String, ByVal psPerRet As String, ByVal psPerDep As String, _
    ByVal psPerTraHacia As String, ByVal psPerTraDesde As String) As String
    
Dim sCad As String

        sCad = ""
    
        sCad = sCad & "0225"
        sCad = sCad & "0"
        sCad = sCad & "0000"
        sCad = sCad & "N"
        sCad = sCad & Left(Mid(psCtaCod, 6, 13) & "0" & Left(psBIN & Space(11), 11) & Right("000" & psAge, 3) & Space(28), 28)
        sCad = sCad & psCtaDef
        sCad = sCad & "CUSCO" & Space(10)
        sCad = sCad & "CUSCO" & Space(10)
        'sCad = sCad & IIf(psMoneda = "1", "604", "840")
        sCad = sCad & psPerCons
        sCad = sCad & psPerRet
        sCad = sCad & psPerDep
        sCad = sCad & psPerTraHacia
        sCad = sCad & psPerTraDesde
                
    GeneraTrama0225 = sCad
    
End Function

Public Function GeneraTrama0205(ByVal psCtaCod As String, ByVal psBIN As String, _
    ByVal psAge As String) As String
    
Dim sCad As String

        sCad = ""
    
        sCad = sCad & "0205"
        sCad = sCad & "0"
        sCad = sCad & "0000"
        sCad = sCad & "N"
        sCad = sCad & Left(Mid(psCtaCod, 6, 13) & "0" & Left(psBIN & Space(11), 11) & Right("000" & psAge, 3) & Space(28), 28)
                
    GeneraTrama0205 = sCad
    
End Function

'Public Function AbrirConexion() As ADODB.Connection
'Dim sCadCadConex As String
'
'
'    sCadCadConex = "Provider=SQLOLEDB.1;Password=desarrollomaynas;User ID=sa;Initial Catalog=DBTarjetaP;Data Source=192.168.15.25"
'    'sCadCadConex = "Provider=SQLOLEDB.1;Password=1234;User ID=sa;Initial Catalog=DBTarjetaP;Data Source=00TI02\SQLEXPRESS"
'
'
'
'    Set C = New ADODB.Connection
'    C.Open sCadCadConex
'
'    Set AbrirConexion = C
'
'End Function
'
'Public Sub CerrarConexion()
'    C.Close
'    Set C = Nothing
'
'End Sub

'***************************************************
Public Function ArmaFecha(dtmFechas As Date) As String
    Dim txtMeses As String
    txtMeses = Choose(Month(dtmFechas), "Enero", "Febrero", "Marzo", "Abril", _
                                        "Mayo", "Junio", "Julio", "Agosto", _
                                        "Setiembre", "Octubre", "Noviembre", "Diciembre")
                                        
    ArmaFecha = Right("00" & Trim(Str(Day(dtmFechas))), 2) & " de " & txtMeses & " del " & Year(dtmFechas)

End Function

Public Sub CargaArchivo(lsArchivo As String, lsRutaArchivo As String)
    Dim X As Long
    Dim Temp As String
    Temp = GetActiveWindow()
    X = ShellExecute(Temp, "open", lsArchivo, "", lsRutaArchivo, 1)
    If X <= 32 Then
        If X = 2 Then
            MsgBox "No se encuentra el Archivo adjunto, " & vbCr & " verifique el servidor de archivos", vbInformation, " Aviso "
        ElseIf X = 8 Then
            MsgBox "Memoria insuficiente ", vbInformation, " Aviso "
        Else
            MsgBox "No se pudo abrir el Archivo adjunto", vbInformation, " Aviso "
        End If
    End If
  
End Sub

'***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Public Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
        
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox Err.Description, vbInformation, "Aviso"
  ExcelBegin = False
End Function
'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Public Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox Err.Description, vbInformation, "Aviso"
End Sub


'********************************
' Adiciona Hoja a LibroExcel
'********************************
Public Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional pbActivaHoja As Boolean = True)
Dim lbExisteHoja As Boolean
Dim lbBorrarRangos As Boolean
On Error Resume Next
lbExisteHoja = False
lbBorrarRangos = False
activaHoja:
For Each xlHoja1 In xlLibro.Worksheets
    If UCase(xlHoja1.Name) = UCase(psHojName) Then
        If Not pbActivaHoja Then
            SendKeys "{ENTER}"
            xlHoja1.Delete
        Else
            xlHoja1.Activate
            If lbBorrarRangos Then xlHoja1.Range("A1:BZ1").EntireColumn.Delete
            lbExisteHoja = True
        End If
       Exit For
    End If
Next
If Not lbExisteHoja Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = psHojName
    If Err Then
        Err.Clear
        pbActivaHoja = True
        lbBorrarRangos = True
        GoTo activaHoja
    End If
End If
End Sub

'DAOR 20091202
Public Sub fEnfoque(ctrControl As Control)
    ctrControl.SelStart = 0
    ctrControl.SelLength = Len(ctrControl.Text)
End Sub

'DAOR 20091202
Public Function Letras(intTecla As Integer, Optional lbMayusculas As Boolean = True) As Integer
If lbMayusculas Then
    Letras = Asc(UCase(Chr(intTecla)))
Else
    Letras = Asc(LCase(Chr(intTecla)))
End If
End Function

'DAOR 20091202
Public Function NumerosEnteros(intTecla As Integer, Optional pbNegativos As Boolean = False) As Integer
Dim cValidar As String
    If pbNegativos = False Then
        cValidar = "0123456789"
    Else
        cValidar = "0123456789-"
    End If
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) = 0 Then
            intTecla = 0
            Beep
        End If
    End If
    NumerosEnteros = intTecla
End Function

'DAOR 20100113, pnAlineamiento: 0 (Izquierda), 1 (Derecha) y 2 (Centro)
Public Function fijarTamanoTexto(psTexto As String, pnTamano As Integer, Optional pnAlineamiento As Integer = 0, Optional psCaracterDeRelleno As String = " ") As String
Dim lsTemp As String
Dim lnPosIni As Long

    Select Case pnAlineamiento
        Case 0
            fijarTamanoTexto = Left(psTexto & String(pnTamano, psCaracterDeRelleno), pnTamano)
        Case 1
            fijarTamanoTexto = Right(String(pnTamano, psCaracterDeRelleno) & psTexto, pnTamano)
        Case 2
            psTexto = Trim(psTexto)
            If Len(psTexto) > pnTamano Then
                psTexto = Left(psTexto, pnTamano)
            Else
                lnPosIni = Int(pnTamano / 2) - Int(Len(psTexto) / 2)
                psTexto = String(lnPosIni, psCaracterDeRelleno) & psTexto & String(lnPosIni, psCaracterDeRelleno)
                fijarTamanoTexto = psTexto
            End If
        End Select
End Function



Public Function CentrarCadena(psCadena As String, pnNroLineas As Long, Optional lnEspaciosIzq As Integer = 0, Optional lsCarImp As String = " ") As String
    Dim psNinf As Long
    Dim lnPosIni As Long
    
    psCadena = Trim(psCadena)
    If Len(psCadena) > pnNroLineas Then
        'psCadena = Left(psCadena, pnNroLineas)
        'MsgBox "EL valor de la Cadena enviada es mayor al espacio destinado", vbInformation, "Aviso"
        psCadena = Left(psCadena, pnNroLineas)
    End If
    'Else
    psNinf = Len(psCadena) / 2
    lnPosIni = Int(pnNroLineas / 2) - Int(Len(psCadena) / 2)
    
    'psCadena = String((pnNroLineas / 2) - psNinf, " ") & psCadena & String(pnNroLineas - Len(psCadena), " ")
    psCadena = String(lnEspaciosIzq, " ") & String(lnPosIni, lsCarImp) & psCadena & String(lnPosIni, lsCarImp)
    CentrarCadena = psCadena
   'End If
End Function
