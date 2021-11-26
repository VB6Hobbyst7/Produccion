Attribute VB_Name = "gDeclare"
Private Sub Main()
On Error GoTo ErrMain

Dim Tamano   As Long
Dim BuferArchivo

'    Set dbCmact = New ADODB.Connection
'    dbCmact.CommandTimeout = 30
'    dbCmact.ConnectionTimeout = 30
'    dbCmact.CursorLocation = adUseClient
'    If Dir(App.Path & "\Contabilidad.ini") <> "" Then
'      Open App.Path & "\Contabilidad.ini" For Input As #1
'      Tamano = LOF(1)
'      BuferArchivo = Input(Tamano, #1)   ' Establece la posición.
'      gsConnection = Trim(BuferArchivo)
'      Close #1
'    Else
'
      Dim oconect As DConecta
      Set oconect = New DConecta
      
      If oconect.AbreConexion() = True Then
           gsConnection = oconect.CadenaConexion
           oconect.CierraConexion
           Set oconect = Nothing
      End If
      
'      If gsConnection = "" Then
'         'gsConnection = "DSN=DSNCMACT;uid=DBAccess;pwd=cmact;"
'      End If
'    End If
    
    'dbCmact.Close
    'dbCmact.Open gsConnection
    'Para Impresión
    'frmMdiMain.Show
      
    'frmPLogin.Show 1
    'If frmPLogin.LoginSucceeded Then
    '   If Not GetTipCambio(gdFecSis) Then
    '      frmTipoCambio.Show 1
    '      AbreConexion
    '      GetTipCambio gdFecSis
    '   End If
   '    CierraConexion

    '   frmSplash.Show 1
   
    gsCodUser = "CASL"
       frmMdiMain.Show
    'Else
    '   CierraConexion
    '   End
    'End If
    Exit Sub
ErrMain:
    MsgBox "Error de Apertura de dbCmact de Datos. " & Err.Description, vbCritical, "Error"
    Set dbCmact = Nothing
    End
End Sub

Public Function SuperBye(Optional pbMensaje As Boolean = True) As Boolean
Dim lsSql As String
On Error GoTo ERROR
'AbreConexion
'If frmPLogin.LoginSucceeded = True Then
'    If pbMensaje Then
'        If MsgBox("Está seguro que desea salir del sistema?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
''            lsSql = "UPDATE LOGIN SET dfin = '" & FechaHora(gdFecSis) & "', cEstado = 'I', bAtenPub = 0 " _
'                & " WHERE dinicio = (select max(dinicio) from login WHERE cCodUsu = '" & gsCodUser & "' and cEstado = 'A') " _
'                & " and cCodUsu = '" & gsCodUser & "'"
'            dbCmact.BeginTrans
'            dbCmact.Execute lsSql
'            dbCmact.CommitTrans
'            dbCmact.Close
'            Set dbCmact = Nothing
'            frmPLogin.Hide
'            SuperBye = True
'            End
'        End If
'    Else
'        lsSql = "UPDATE LOGIN SET dfin = '" & FechaHora(gdFecSis) & "', cEstado = 'I', bAtenPub = 0 " _
'                & " WHERE dinicio = (select max(dinicio) from login WHERE cCodUsu = '" & gsCodUser & "' and cEstado = 'A') " _
'                & " and cCodUsu = '" & gsCodUser & "'"
'        dbCmact.BeginTrans
'        dbCmact.Execute lsSql
'        dbCmact.CommitTrans
'        dbCmact.Close
'        Set dbCmact = Nothing
'        frmPLogin.Hide
'        End
'    End If
'Else
'    If MsgBox("Desea Salir del Sistema", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
'        frmPLogin.Hide
'        dbCmact.Close
'        Set dbCmact = Nothing
        End
'    End If
'End If

Exit Function
ERROR:
     MsgBox Err.Description, vbExclamation, "Error"
     dbCmact.RollbackTrans
End Function

Public Function CierreRealizado(Optional pnTipo As Integer = 1) As Boolean
Dim rsVarSis As New ADODB.Recordset
Dim pdCieDia As Date
Dim pdCieMes As Date
VSQL = "select cNomVar,cValorVar from VarSistema where cNomVar in ('dFecCierre','dFecCierreMes') and cCodProd='ADM'"
rsVarSis.Open VSQL, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
Do While Not rsVarSis.EOF
    If Trim(rsVarSis!cNomVar) = "dFecCierre" Then
        pdCieDia = CDate(Trim(rsVarSis!cValorVar))
    ElseIf Trim(rsVarSis!cNomVar) = "dFecCierreMes" Then
        pdCieMes = CDate(Trim(rsVarSis!cValorVar))
    End If
    rsVarSis.MoveNext
Loop
rsVarSis.Close
Set rsVarSis = Nothing
If pnTipo = 1 Then
    CierreRealizado = IIf(pdCieDia = gdFecSis, True, False)
ElseIf pnTipo = 2 Then
    CierreRealizado = IIf(pdCieMes = gdFecSis, True, False)
End If
End Function



'Private Function GetConexion(psApp As String) As String
'Dim sConex As String
'sConex = GetSetting(psApp, "Init", "Provider")
'If sConex <> "" Then
'   'GetConexion = "PROVIDER=" & Encripta(sConex, False) & ";"
'   sConex = GetSetting(psApp, "Init", "Server")
'   If sConex <> "" Then
'      GetConexion = GetConexion & "SERVER=" & Encripta(sConex, False) & ";"
'      sConex = GetSetting(psApp, "Init", "DataBase")
'      If sConex <> "" Then
'         GetConexion = GetConexion & "DATABASE=" & Encripta(sConex, False) & ";"
'         sConex = GetSetting(psApp, "Init", "Uid")
'         If sConex <> "" Then
'            GetConexion = GetConexion & "UID=" & Encripta(sConex, False) & ";"
'            sConex = GetSetting(psApp, "Init", "Pwd")
'            If sConex <> "" Then
'               GetConexion = GetConexion & "PWD=" & Encripta(sConex, False) & ";"
'            End If
'         End If
'      End If
'   End If
'End If
'End Function
'
'ARLO20161221-------------------------

Public Function GetMovNro(psCodUser As String, psCodAge As String, Optional psCorrelativo As String = "00") As String
    Dim oCon As NConstSistemas
    Set oCon = New NConstSistemas
    
    If Len(psCodAge) = 2 Then
        GetMovNro = Format(gdFecSis, "yyyymmdd") & Format(Time, "hhmmss") & oCon.LeeConstSistema(gConstSistCodCMAC) & psCodAge & psCorrelativo & psCodUser
    Else
        GetMovNro = Format(gdFecSis, "yyyymmdd") & Format(Time, "hhmmss") & psCodAge & psCorrelativo & psCodUser
    End If
End Function

