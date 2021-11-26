VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   4725
   ClientLeft      =   3150
   ClientTop       =   2340
   ClientWidth     =   8595
   ControlBox      =   0   'False
   ForeColor       =   &H80000006&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin NegForms.Usuario ctlUsuario 
      Left            =   600
      Top             =   3720
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.TextBox TxtClave 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C16A0B&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3360
      MaxLength       =   25
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Ingrese su Clave Secreta"
      Top             =   3735
      Width           =   2430
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   3360
      TabIndex        =   1
      Top             =   4200
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   4440
      TabIndex        =   0
      Top             =   4200
      Width           =   1000
   End
   Begin VB.Label LblUsu 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C16A0B&
      Height          =   315
      Left            =   3360
      TabIndex        =   4
      Top             =   3360
      Width           =   2430
   End
   Begin VB.Image Image1 
      Height          =   3285
      Left            =   0
      Picture         =   "frmLogin.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8580
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2640
      TabIndex        =   5
      Top             =   3375
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clave     "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2640
      TabIndex        =   3
      Top             =   3750
      Width           =   690
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oAcceso As COMDPersona.UCOMAcceso
Dim oSeguridad As COMManejador.Pista 'JUEZ 20160216
Dim UsRf As Boolean 'JUEZ 20160216
'Para el poder cargar los Datos de la Maquina Cliente
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Public gsSistDescripcion As String 'ALPA20160722
Public gnSistCod As String 'ALPA20160722
Public gsFechaVersion As String 'ALPA20160722

Private Function UbicaMenuActivos(ByVal psName As String, ByVal psIndex As String, ByVal nPunt As Integer) As Boolean
   
    'If MatMenuItems(nPunt).sName = "M170000000000" Then
    '    MatMenuItems(nPunt).bCheck = MatMenuItems(nPunt).bCheck
    'End If
    
       'Si lo encuentro
    If Left(MatMenuItems(nPunt).sName, 11) = psName And MatMenuItems(nPunt).sIndex = psIndex Then
        If MatMenuItems(nPunt).bCheck Then
            UbicaMenuActivos = True
        Else
            UbicaMenuActivos = False
        End If
        Exit Function
    End If
    
    'Si Tiene mas Nodos Hijos
    If MatMenuItems(nPunt).nPuntDer <> -1 Then
        UbicaMenuActivos = UbicaMenuActivos(psName, psIndex, MatMenuItems(nPunt).nPuntDer)
        If UbicaMenuActivos Then
            Exit Function
        End If
    End If
    
    'Si Tiene mas Nodos Paralelos
    If MatMenuItems(nPunt).nPuntAbajo <> -1 Then
        UbicaMenuActivos = UbicaMenuActivos(psName, psIndex, MatMenuItems(nPunt).nPuntAbajo)
    End If
    
    'Si es Nodo Final
    If MatMenuItems(nPunt).nPuntDer = -1 And MatMenuItems(nPunt).nPuntAbajo = -1 Then
        UbicaMenuActivos = False
    End If
    
End Function


Private Function ActualizaMenuActivos(ByRef nPunt As Integer) As Integer
    'If MatMenuItems(nPunt).sName = "M170100000021" Then
    '    MatMenuItems(nPunt).bCheck = MatMenuItems(nPunt).bCheck
    'End If
    'Si Es Nodo Final
        
    
    If MatMenuItems(nPunt).nPuntDer = -1 And MatMenuItems(nPunt).nPuntAbajo = -1 Then
        If oAcceso.TienePermiso(Left(MatMenuItems(nPunt).sName, 11), MatMenuItems(nPunt).sIndex) Then
            ActualizaMenuActivos = 1
            MatMenuItems(nPunt).bCheck = True
        Else
            ActualizaMenuActivos = 0
            MatMenuItems(nPunt).bCheck = False
        End If
    End If
    
    If MatMenuItems(nPunt).nPuntDer = 0 Then
        MatMenuItems(nPunt).bCheck = MatMenuItems(nPunt).bCheck
    End If
    'Si Tiene mas Nodos Hijos
    If MatMenuItems(nPunt).nPuntDer <> -1 Then
        ActualizaMenuActivos = ActualizaMenuActivos(MatMenuItems(nPunt).nPuntDer)
        If ActualizaMenuActivos > 0 Then
            MatMenuItems(nPunt).bCheck = True
            ActualizaMenuActivos = 1
        Else
            ActualizaMenuActivos = 0
        End If
    End If
    
    'If nPunt = 441 Then
    '    MatMenuItems(nPunt).bCheck = MatMenuItems(nPunt).bCheck
    'End If
    
    'Si Tiene mas Nodos Paralelos
    If MatMenuItems(nPunt).nPuntAbajo <> -1 Then
        If oAcceso.TienePermiso(Left(MatMenuItems(nPunt).sName, 11), MatMenuItems(nPunt).sIndex) Then
            MatMenuItems(nPunt).bCheck = True
            ActualizaMenuActivos = 1
        End If
        ActualizaMenuActivos = ActualizaMenuActivos + ActualizaMenuActivos(MatMenuItems(nPunt).nPuntAbajo)
    End If
    
    'If nPunt = 441 Then
    '    MatMenuItems(nPunt).bCheck = MatMenuItems(nPunt).bCheck
    'End If
End Function

Private Sub CargaMenu(ByVal poAcceso As COMDPersona.UCOMAcceso)
Dim Ctl As Control
Dim sTipo As String

On Error Resume Next
    For Each Ctl In MDISicmact.Controls
        sTipo = TypeName(Ctl)
        If sTipo = "Menu" Then
            If InStr(poAcceso.sCadMenu, Ctl.Name) > 0 Then
                Ctl.Visible = True
            Else
                Ctl.Visible = False
            End If
        End If
        DoEvents
    Next
End Sub


Private Sub CargaMenuMDIMain()
Dim Ctl As Control
Dim sTipo As String
Dim nPos As Integer
Dim sCadMenuTemp As String
Dim nPunt As Integer
    
    
Call ActualizaMenuActivos(nPunt)

On Error Resume Next
    For Each Ctl In MDISicmact.Controls
        sTipo = TypeName(Ctl)
        If sTipo = "Menu" Then
            'If Ctl.Name = "M070100000000" Then
            '    sTipo = sTipo
            'End If
            If UbicaMenuActivos(Ctl.Name, Format(Ctl.Index, "00"), 0) Then
                Ctl.Visible = True
            Else
                Ctl.Visible = False
            End If
        End If
    Next
End Sub

Private Sub cmdAceptar_Click()
    
Dim i As Integer
Dim Y As Integer
Dim bInterconexCorrecta As Boolean
Dim bClaveIncorrecta As Boolean
Dim rsITF As ADODB.Recordset
Dim rsVar As ADODB.Recordset
Dim psServerName As String
Dim psDBName As String
Dim psCadConexion As String
Dim RsMenu As ADODB.Recordset
Dim bTieneAlgunPermiso As Boolean
Dim bTienePermisoRetiroSinFirma As Boolean
Dim bIniciarNuevoDia As Boolean
Dim pMatMenuPer As Variant
'Dim oSeguridad As New COMManejador.Pista

    Screen.MousePointer = 11

    Call oAcceso.CargarLogin_NEW(Trim(lblUsu.Caption), Trim(TxtClave.Text), bInterconexCorrecta, bClaveIncorrecta, rsITF, rsVar, psServerName, psDBName, psCadConexion, _
                              pMatMenuPer, bTieneAlgunPermiso, bTienePermisoRetiroSinFirma, bIniciarNuevoDia, gRsOpeF2, gRsExtornos, gRsOpeCMACRecep, gRsOpeCMACLlam, gRsOpeRepo, 1)
    

    'Set oAcceso = Nothing
    
    'AGREGADO POR ARLO 20170331
    Call CargaVarSistema(False, rsVar, psServerName, psDBName, psCadConexion)
    gsSistDescripcion = "SICMACM Negocio"
    gnSistCod = 1
    gsFechaVersion = "20210429"             'CAMBIAR LA FECHA CADA VEZ QUE SE COMPILA"
    'AGREGADO POR ARLO 20170331
    
    If Not bInterconexCorrecta Then
        MsgBox "No se puede Establecer la Conexion con el Servidor" & Chr(10) & "Consulte con el Area de Sistemas", vbCritical, "Conexion SICMACT"
        'Set oAcceso = Nothing
        TxtClave.Text = ""
        TxtClave.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If

    If Not bClaveIncorrecta Then
        MsgBox "Clave Incorrecta, Ingrese su Clave Nuevamente ", vbCritical, "Conexion SICMACT"
        'Set oAcceso = Nothing
        'AGREGADO POR ARLO 20170331
        Set oSeguridad = New COMManejador.Pista 'JUEZ 20160216
            Call oSeguridad.InsertarPista(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, "Intento Fallido al Ingreso al " & gsSistDescripcion & " Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-" & gsFechaVersion)
            If UsRf Then
                Call oSeguridad.InsertarPistaSesion(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, gnSistCod)
            End If
        Set oSeguridad = Nothing
        'AGREGADO POR ARLO 20170331
        fEnfoque TxtClave
        TxtClave.Text = ""
        TxtClave.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = False

    '************************************************************************

    'Call CargaVarSistema(False, rsVar, psServerName, psDBName, psCadConexion) 'COMENTADO POR ARLO2017031
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    
    Call fgITFParametros(rsITF)

    '***Modificaion MPBR
    Vusuario = lblUsu.Caption
    'EJVG20120419
    gsGruposUser = oAcceso.CargaUsuarioGrupo(gsCodUser, gsDominio)
    gsGruposUser = Replace(gsGruposUser, "'", "")
    If Not bTieneAlgunPermiso Then
        MsgBox "Usted No Tiene Acceso a Ninguna Opcion del Sistema, Avise a Sistemas", vbInformation, "Aviso"
        'Set oAcceso = Nothing
        Screen.MousePointer = 0
        End
    End If

    Call CargaMenuMDIMain_NEW(pMatMenuPer)

    'Permiso Retiro sin tener en cuenta las firmas
    gbRetiroSinFirma = bTienePermisoRetiroSinFirma
    
    'Permiso Para Inicio de Dia
    If bIniciarNuevoDia Then
        Screen.MousePointer = 0
        frmInicioDia.Show 1
        'ALPA 20080721
        TxtClave.SetFocus
        Screen.MousePointer = 11
    End If
      
    'Inicializa DMONT para lectura de tarjetas magnéticas

    Dim X As Double
    ChDrive App.Path
    ChDir App.Path

    'Trasladar para uso de otro equipo   -- ppoa

'    Dim nRetVal As Long
'    nRetVal = MakeWord(PpDvcSignal())
'    If nRetVal = ERR_DMON_OFF Then
'        X = Shell(App.path & "\Dmonnt.exe", vbMinimizedNoFocus)
'    End If


    'Obtiene la impresora predeterminada
    Dim sImpresora As String
    Dim lnPos As Long
    sImpresora = Printer.DeviceName
    If Left(sImpresora, 2) <> "\\" Then
        lnPos = InStr(1, Printer.port, ":", vbTextCompare)
        If lnPos > 0 Then
            sLpt = Mid(Printer.port, 1, lnPos - 1)
        Else
            sLpt = "LPT1"
        End If
    Else
        sLpt = frmImpresora.EliminaEspacios(sImpresora)
    End If

    MsgBox "Por favor Configure su Impresora antes de Empezar sus operaciones", vbInformation, "Aviso"
    frmImpresora.Show 1
    
    ValidaAccesoSesion 'JUEZ 20160329
    
    '**DAOR 20081125, Tipo de PinPad ******************
    Call RecuperaConfigPinPad
    '**************************************************
    
    '***** GITU 20-04-2001 Tiempo de Espera PINPADS*****'
    Call RecuperaTimeOutPinPadAG
    
    'NRLO 20180319 ERS027-2017
    Dim rsValidaVisita As ADODB.Recordset
    Dim obDCredVD As COMDCredito.DCOMCredito
    Set obDCredVD = New COMDCredito.DCOMCredito
    Set rsValidaVisita = obDCredVD.ValidarVisitasMinimasPostDesembolso(gsCodUser, gsCodCargo)
    If Not (rsValidaVisita.EOF And rsValidaVisita.BOF) Then
        If rsValidaVisita!nEstado <> 1 Then
            MsgBox rsValidaVisita!cMensaje, vbInformation, "Aviso"
            If rsValidaVisita!nEstado = 0 Then
                rsValidaVisita.Close
                Set obDCredVD = Nothing
                End
            End If
            rsValidaVisita.Close
            Set obDCredVD = Nothing
        End If
    Set obDCredVD = Nothing
    End If
    'NRLO FIN 20180319 ERS027-2017
        
    '**DAOR 20090203, Regitro de ingreso al sistema****
    'Bitacora Version 201011
    Set oSeguridad = New COMManejador.Pista 'JUEZ 20160216
    Call oSeguridad.InsertarPista(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, "Ingreso al " & gsSistDescripcion & " Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-" & gsFechaVersion)
    If UsRf Then
        Call oSeguridad.InsertarPistaSesion(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, gnSistCod)
    End If
    Set oSeguridad = Nothing
    '**************************************************
    
    
    Screen.MousePointer = 0
    Unload Me
    IniciarVerDocsPendiente 'EJVG20140523
    'frmVideo.Show 1
    'MADM 20091112 - MDISicmact.Caption = "SICMACT " & Space(15) & Trim(gsNomAge) & Space(10) & gsCodUser & Space(2) & gsServerName & "\" & gsDBName & Space(5) & Format(gdFecSis, "dd/mm/yyyy")
    MDISicmact.Caption = "SICMACM - " & Trim(gsNomAge) & " - " & Trim(gsCodUser) & Space(10) & gsServerName & "\" & gsDBName & Space(5) & Format(gdFecSis, "dd/mm/yyyy")
    
    '**DAOR 20100406, Control de versión ********************************
    MDISicmact.Caption = MDISicmact.Caption & Space(5) & " - Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & " - " & gsFechaVersion 'Cambiar la fecha cada vez que se compila"
    '********************************************************************
    '**************************
    MDISicmact.SBBarra.Panels(2).Text = Format(gdFecSis, "dddd - dd - mmmm - yyyy") & Space(3) & Format(Time, "hh:mm AMPM")
    MDISicmact.Show

End Sub


Private Sub CargaMenuMDIMain_NEW(pMatMenuPer As Variant)
Dim Ctl As Control
Dim sTipo As String
Dim nPos As Integer
Dim sCadMenuTemp As String
Dim nPunt As Integer
Dim i As Integer
    
On Error Resume Next
    i = 0
    For Each Ctl In MDISicmact.Controls
        sTipo = TypeName(Ctl)
        If sTipo = "Menu" Then
            
            Ctl.Visible = IIf(pMatMenuPer(i) = 1, True, False)
            i = i + 1
        End If
    Next
End Sub

'***********************************************************

Private Sub cmdCancelar_Click()
    If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        End
    End If
End Sub

'**************************************************
'Para Poder Cargar el Usuario de la Maquina Cliente
'**************************************************
Public Function ObtenerUsuarioCliente() As String
Dim buffUsu As String
Dim lSizeUsu As Long

    buffUsu = String(255, " ")
    lSizeUsu = Len(buffUsu)
    Call GetUserName(buffUsu, lSizeUsu)
    ObtenerUsuarioCliente = Mid(Trim(buffUsu), 1, Len(Trim(buffUsu)) - 1)
       
End Function


'Private Sub Form_Activate()
'    TxtClave.SetFocus
'End Sub

Private Sub Form_Load()
   
    'Dim oAcceso As COMDPersona.UCOMAcceso
    Dim lsRutaUltActualiz As String
    Dim lsFlagActualizaEXE As String
''    'ALPA 20080721
''    TxtClave.SetFocus
    
    If App.PrevInstance Then
            MsgBox "Ud. solo puede cargar el sistema una solo vez.", vbInformation, "Aviso"
            End
    End If
 

    
    Me.Icon = LoadPicture(App.Path & "\BMPS\cm.ico")
    If Not ValidaConfiguracionRegional Then
        MsgBox "Su actual CONFIGURACIÓN REGIONAL NO ES CORRECTA. Revísela y reinicie.", vbInformation, "Aviso"
        End
    End If
    
     
    Dim oConst As COMDConstSistema.NCOMConstSistema 'JUEZ 20160405
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Me.Caption = "Acceso"
    Set oImp = New COMDConstSistema.DCOMImpresoras
    oImpresora.inicia oImp.GetImpreSetup(GetMaquinaUsuario) 'oImp.GetMaquina)

    gImpresora = oImp.GetImpreSetup(GetMaquinaUsuario)  'oImp.GetMaquina)

    If gImpresora = -1 Then
        MsgBox "Ud. debe asignar los caracteres de impresion por defecto para esta maquina.", vbInformation, "Aviso"
        frmCaracImpresion.Show 1

        gImpresora = oImp.GetImpreSetup(GetMaquinaUsuario) 'oImp.GetMaquina)

        If gImpresora = -1 Then
            MsgBox "Como Ud. no ha elegido los caracteres de impresion para esta maquina, se esta procediendo a asignarle el tipo EPSON, si Ud. desea luego puede modificarlo.", vbInformation, "Aviso"
            oImp.SetImpreSetup GetMaquinaUsuario, gEPSON
        End If
    End If
    Set oImp = Nothing
    
    lblUsu.Caption = UCase(ObtenerUsuarioCliente) 'UCase(oAcceso.ObtenerUsuario)
    
    Set oAcceso = New COMDPersona.UCOMAcceso
    
    If Not oAcceso.VerificarUsuarioExistaEnRRHH(lblUsu.Caption) Then
        MsgBox "Su Codigo de Usuario no a sido Registrado por el Area Recursos Humanos", vbInformation, "Aviso"
'        Set oAcceso = Nothing
        End
    End If
'    Set oAcceso = Nothing
    'RotateText 90, Picture1, "Times New Roman", 15, 25, 1700, "NEGOCIO"
    Call ctlUsuario.Inicio(Trim(lblUsu.Caption))
    gsCodAge = ctlUsuario.CodAgeAct
    gsCodUser = Trim(lblUsu.Caption)
    gsNomAge = Trim(ctlUsuario.DescAgeAct)
    gsCodArea = Trim(ctlUsuario.AreaCod)
    gsCodCargo = Trim(ctlUsuario.PersCargoCod)
    gsCodPersUser = ctlUsuario.PersCod
    'JUEZ 20160405 ********************
    gsNomPersUser = PstaNombre(ctlUsuario.UserNom)
    gsNomArea = ctlUsuario.AreaNom
    gsNomCargo = ctlUsuario.PersCargo
    'END JUEZ *************************
        
    'JUEZ 20121219 **********************************************
    Dim oNCredito As COMNCredito.NCOMCredito
    Set oNCredito = New COMNCredito.NCOMCredito
    'gnAgenciaCredEval = oNCredito.ObtieneAgenciaCredEval(gsCodAge)
    gnAgenciaCredEval = oNCredito.ObtieneAgenciaHabNivelesApr(gsCodAge) 'JUEZ 20130401
    Set oNCredito = Nothing
    'END JUEZ ***************************************************
    
    'JUEZ 20160405 ****************************************
    Set oConst = New COMDConstSistema.NCOMConstSistema
    gsCorreoHost = oConst.LeeConstSistema(90)
    gsCorreoEnvia = oConst.LeeConstSistema(91)
    Set oConst = Nothing
    'END JUEZ ****-****************************************
    
    ValidaAccesoSesion 'JUEZ 20160329
        
    'WIOR 20151109 ***
    Dim oDHojaRuta As DCOMhojaRuta
    Set oDHojaRuta = New DCOMhojaRuta

    gnAgenciaHojaRutaNew = oDHojaRuta.ObtieneAgenciaHojaRutaNew(gsCodAge)
    Set oDHojaRuta = Nothing
    'WIOR FIN ********
    
    '************************************************
    'AMARRE MAQUINA A USUARIO
    '************************************************
    Dim sSqlVal As String
    Dim o As COMConecta.DCOMConecta
    Set o = New COMConecta.DCOMConecta
    Dim sCodAgeMaq As String
    Dim nValorSis As Integer
    Dim R As ADODB.Recordset
    
    o.AbreConexion
    
    'Dim oConst As COMDConstSistema.NCOMConstSistema
    Set oConst = New COMDConstSistema.NCOMConstSistema
    
    'gcEmpresa = oConst.LeeConstSistema(gConstSistCMACNombreCompleto)
    Call oConst.ObtenerValoresCargaLogin(ctlUsuario.CodAgeAct, gcEmpresa, lsRutaUltActualiz, lsFlagActualizaEXE, gsProyectoActual)
    
    Set oConst = Nothing
    
    If fgActualizaUltVersionEXE(ctlUsuario.CodAgeAct, lsRutaUltActualiz, lsFlagActualizaEXE) = True Then ' Verifica si existe una actualizacion
       End
    End If
    '
    'RotateText 90, Picture1, "Tahoma", 13, 25, 1500, "Negocio"
End Sub

Private Function EstacionExcluida(ByVal psEstacion As String) As Boolean
Dim ssql As String
Dim o As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
    
    Set o = New COMConecta.DCOMConecta
        
ssql = "Select cEstacion from AgenciaEstacionesExcl where cEstacion = '" & GetMaquinaUsuario & "'"

o.AbreConexion

Set R = o.CargaRecordSet(ssql)
If R.RecordCount > 0 Then
    EstacionExcluida = True
Else
    EstacionExcluida = False
End If
Set R = Nothing
o.CierraConexion


End Function

Private Sub Form_Unload(Cancel As Integer)
    Set oAcceso = Nothing
End Sub



Private Sub TxtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAceptar_Click
    End If
End Sub

Private Function GetMaquinaUsuario() As String  'Para obtener la Maquina del Usuario
    Dim buffMaq As String
    Dim lSizeMaq As Long
    buffMaq = Space(255)
    lSizeMaq = Len(buffMaq)
    GetComputerName buffMaq, lSizeMaq
    GetMaquinaUsuario = Trim(Left$(buffMaq, lSizeMaq))
End Function

'**DAOR 20081125, Obtenido del módulo Admin de Tarjetas*****************************
Public Sub RecuperaConfigPinPad()
    Dim lrs As ADODB.Recordset
    Dim loCn As COMConecta.DCOMConecta
    Set loCn = New COMConecta.DCOMConecta
    
     
    Set lrs = New ADODB.Recordset

    loCn.AbreConexion
        Set lrs = loCn.ConexionActiva.Execute(" exec ATM_RecuperaDatosPinPad '" & GetMaquinaUsuario & "'")
    
    If Not (lrs.EOF And lrs.BOF) Then
        gnTipoPinPad = lrs("nTipoPinPad")
        gnPinPadPuerto = lrs("nNumPuerto")
    End If
    
    loCn.CierraConexion
    Set loCn = Nothing
    
End Sub
'******************************************************************
'JUEZ 20160329 ****************************************************************
Private Sub ValidaAccesoSesion()
Dim acceso As Boolean
Dim oConst As COMDConstSistema.NCOMConstSistema 'JUEZ 20160405
Dim sMaquinaActiva As String 'JUEZ 20160405

    UsRf = False
    acceso = False
    
    Set oSeguridad = New COMManejador.Pista
    If oSeguridad.ValidaAccesoPistaRF(gsCodUser) Then
        UsRf = True
        'acceso = oSeguridad.ValidaIniciaSessionRF(gsCodUser, gdFecSis, GetMaquinaUsuario)
        acceso = oSeguridad.ValidaIniciaSessionRF(gsCodPersUser, gdFecSis, GetMaquinaUsuario, 1, sMaquinaActiva) 'JUEZ 20160125
        If acceso Then
            MsgBox "No puede Acceder al Sistema porque tiene una Sesion Abierta en otra PC. Consulte con Help Desk", vbExclamation, "Ingreso SICMACM Negocio"

            'Envio Correo - JUEZ 20160405 ****************************************
            Dim lsCorreoDestino As String, lsContenido As String
            Set oConst = New COMDConstSistema.NCOMConstSistema
            lsCorreoDestino = oConst.LeeConstSistema(92)

            lsContenido = "El usuario " & gsCodUser & " está intentando acceder a la PC " & GetMaquinaUsuario & _
                          ", sin embargo tiene una sesión abierta en la PC " & sMaquinaActiva & "<p><p>" & _
                          "<b>Nombre Usuario:</b> " & gsNomPersUser & "<br>" & _
                          "<b>Agencia:</b> " & gsNomAge & "<br>" & _
                          "<b>Cargo:</b> " & gsNomCargo

            EnviarMail gsCorreoHost, gsCorreoEnvia, lsCorreoDestino, "Validación Acceso SICMACM Negocio", lsContenido
            'Fin Envio Correo ****************************************************

            TxtClave.Text = ""
            'TxtClave.SetFocus
            Screen.MousePointer = 0
            End
        End If
    End If
    Set oSeguridad = Nothing
End Sub
'END JUEZ *********************************************************************
