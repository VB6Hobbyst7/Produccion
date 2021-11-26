VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Acceso SICMAC I  Administrativo"
   ClientHeight    =   4770
   ClientLeft      =   4005
   ClientTop       =   1890
   ClientWidth     =   8535
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmLogin.frx":030A
   ScaleHeight     =   4770
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   1000
   End
   Begin VB.TextBox TxtClave 
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
      ForeColor       =   &H00C00000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3360
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   8
      ToolTipText     =   "Ingrese su Clave Secreta"
      Top             =   3735
      Width           =   2430
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      ForeColor       =   &H8000000D&
      Height          =   1785
      Left            =   1080
      ScaleHeight     =   1755
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   5040
      Width           =   465
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -240
         Picture         =   "frmLogin.frx":249B2
         Stretch         =   -1  'True
         Top             =   360
         Width           =   390
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   4305
      TabIndex        =   3
      Top             =   4200
      Width           =   1140
   End
   Begin Sicmact.Usuario CtlUsuario 
      Left            =   120
      Top             =   5040
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1890
      Left            =   6000
      TabIndex        =   0
      Top             =   5880
      Width           =   3315
      Begin VB.Label lblIni 
         Alignment       =   2  'Center
         Caption         =   "Sicmac I Administrativo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   1920
         Width           =   3060
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   2520
      TabIndex        =   6
      Top             =   3855
      Width           =   705
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   2520
      TabIndex        =   5
      Top             =   3480
      Width           =   870
   End
   Begin VB.Label LblUsu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   3360
      TabIndex        =   4
      Top             =   3360
      Width           =   2430
   End
   Begin VB.Label lblVer 
      Caption         =   "Label2"
      Height          =   210
      Left            =   600
      TabIndex        =   2
      Top             =   5280
      Width           =   3825
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oAcceso As UAcceso
Dim lnIndiceActual As Integer
'ALPA 20090122******************
Dim objPista As COMManejador.Pista
'*******************************



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

Private Sub CargaMenu(ByVal poAcceso As UAcceso)
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


Private Sub CargaMenuMDIMain(ByVal poAcceso As UAcceso)
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
            If UbicaMenuActivos(Ctl.Name, Format(Ctl.index, "00"), 0) Then
                Ctl.Visible = True
            Else
                Ctl.Visible = False
            End If
        End If
    Next
End Sub

Private Sub CmdAceptar_Click()
    Dim sTitulo As String
    Dim R As ADODB.Recordset
    Dim I As Integer
    Dim Y As Integer
    Dim oConec As DConecta
    Dim sSql As String
    Dim sImpresora As String
    Dim lnPos As Long
    '***Agregado por ELRO el 20111121, según Acta 270-2011/TI-D
    Dim oPersona As New DPersonas
    Dim rsPersona As New ADODB.Recordset
    '***Fin Agregado por ELRO


    Screen.MousePointer = 11
    'sLpt = Mid(Printer.Port, 1, InStr(1, Printer.Port, ":", vbTextCompare) - 1)
    Call CtlUsuario.Inicio(Trim(LblUsu.Caption))
    gsCodAge = CtlUsuario.CodAgeAct
    gsCodArea = CtlUsuario.cAreaCodAct
    gsCodUser = Trim(LblUsu.Caption)
    
    
    Call CargaVarSistema(False)
    Set oAcceso = New UAcceso
    
    'ARLO20161221
    gsCodUser = UCase(oAcceso.ObtenerUsuario)
    glsMovNro = GetMovNro(gsCodUser, gsCodAge)
    gsopecod = LogPistaIngresarSalirSistema
    gsFechaVersion = "20210825" 'CAMBIAR LA FECHA CADA VEZ QUE SE COMPILA
    '*****************************
    Dim psGrupoUsu As String '->***** LUCV20190323, Agregó Según RO-1000373
    
    If Not oAcceso.InterconexionCorrecta Then
        MsgBox "No se puede Establecer la Interconexion con el Servidor" & Chr(10) & "Consulte con el Area de Sistemas", vbCritical, "Conexion SICMACT"
        Set oAcceso = Nothing
        TxtClave.Text = ""
        TxtClave.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    If Not oAcceso.ClaveIncorrectaNT(Trim(LblUsu.Caption), Trim(TxtClave.Text), gsDominio) Then
        'ARLO20161221
        Set objPista = New COMManejador.Pista
        gsopecod = LogPistaIngresoSistema
        objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, "Intento Fallido al Ingreso al Sicmac Administrativo" & " Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-" & gsFechaVersion
        Set objPista = Nothing
        '*****************************
        MsgBox "Clave Incorrecta, Ingrese su Clave Nuevamente ", vbCritical, "Conexion SICMACT"
        Set oAcceso = Nothing
        fEnfoque TxtClave
        TxtClave.Text = ""
        TxtClave.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    '***Modificado por ELRO el 20111014, según Acta 270-2011/TI-D
    Set rsPersona = oPersona.BuscaCliente(CtlUsuario.PersCod, BusquedaEmpleadoCodigo)
    If Not rsPersona.EOF Then
    gnPersPersoneria = rsPersona!nPersPersoneria
    End If
    '***Fin Modificado por ELRO**********************************
    
    CmdAceptar.Enabled = False
    cmdCancelar.Enabled = False
    
        
        Screen.MousePointer = 0
        sImpresora = Printer.DeviceName
        If Left(sImpresora, 2) <> "\\" Then
        lnPos = InStr(1, Printer.Port, ":", vbTextCompare)
        If lnPos > 0 Then
            sLpt = Mid(Printer.Port, 1, lnPos - 1)
        Else
            sLpt = "LPT1"
        End If
        Else
        sLpt = frmImpresora.EliminaEspacios(sImpresora)
        End If
        MsgBox "Por favor Configure su Impresora antes de Empezar sus operaciones", vbInformation, "Aviso"
        frmImpresora.Show 1
        Screen.MousePointer = 11
    
    '************************************************************************
    
    Set R = oAcceso.DameItemsMenu
    I = 0
    ReDim MatMenuItems(0)
    ReDim Preserve MatMenuItems(I + 1)
    MatMenuItems(I).nId = I
    MatMenuItems(I).sCodigo = Trim(R!cCodigo)
    MatMenuItems(I).sCaption = Trim(R!cDescrip)
    MatMenuItems(I).sName = Trim(R!cname)
    MatMenuItems(I).sIndex = Right(R!cname, 2)
    MatMenuItems(I).bCheck = False
    MatMenuItems(I).nNivel = 1
    MatMenuItems(I).nPuntDer = -1
    MatMenuItems(I).nPuntAbajo = -1
    I = I + 1
    Y = I
    R.MoveNext
    Call CargaMenuArbol(R, I, Y)
    R.Close
    
    Call oAcceso.CargaMenu(gsDominio, LblUsu.Caption, , psGrupoUsu) '->***** LUCV20190323, Agregó psGrupoUsu Según RO-1000373
    'Call CargaMenu(oAcceso)
    gsGrupoUsu = psGrupoUsu '->***** LUCV20190323, Según RO-1000373
    Call CargaMenuMDIMain(oAcceso)
    
    'Habilita Permiso para Operaciones y Reportes
    Set oConec = New DConecta
    oConec.AbreConexion
    sSql = "Select cOpeCod, cOpeDesc, cOpeGruCod, cOpeVisible, nOpeNiv from OpeTpo Order by cOpeCod"
    Set R = oConec.CargaRecordSet(sSql)
    Y = 0
    Do While Not R.EOF
        If oAcceso.TienePermiso(R!cOpeCod, "", True) Then
            Y = Y + 1
            MatOperac(Y - 1, 0) = R!cOpeCod
            MatOperac(Y - 1, 1) = R!cOpeDesc
            MatOperac(Y - 1, 2) = IIf(IsNull(R!cOpeGruCod), "", R!cOpeGruCod)
            MatOperac(Y - 1, 3) = R!cOpeVisible
            MatOperac(Y - 1, 4) = R!nOpeNiv
        End If
        R.MoveNext
    Loop
    NroRegOpe = Y
    oConec.CierraConexion
    Set oAcceso = Nothing
    
    Screen.MousePointer = 0
    Unload Me
    'frmVideo.Show 1
    'COMENTADO ARLO20161221
    'glsMovNro = GetMovNro(gsCodUser, gsCodAge)
    'ALPA 20090122 **********************************************************
    'gsopecod = LogPistaIngresarSalirSistema
    'objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gIngresarSistema
    '************************************************************************
    
    'TORE 20190724
    AlertaVencimientoCartaFianza
    
    'ARLO20161221
    Set objPista = New COMManejador.Pista
    gsopecod = LogPistaIngresoSistema
    objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, "Ingreso al Sicmac Administrativo" & " Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-" & gsFechaVersion
    Set objPista = Nothing
    '***********************
    

    '**Modificado por DAOR 20100406, Control de versión ********************************
    'MDISicmact.Caption = "SICMACT " & Space(10) & UCase(gsCodUser) & Space(7) & gsServerName & "\" & gsDBName & Space(5) & Format(gdFecSis, "dd/mm/yyyy")
    MDISicmact.Caption = "SICMACM ADMINISTRATIVO " & Space(10) & UCase(gsCodUser) & Space(7) & gsServerName & "\" & gsDBName & Space(5) & Format(gdFecSis, "dd/mm/yyyy")
    MDISicmact.Caption = MDISicmact.Caption & Space(5) & " - Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-" & gsFechaVersion  'Cambiar la fecha cada vez que se compila
    '***********************************************************************************
    
    MDISicmact.staMain.Panels(2).Text = Format(gdFecSis, "dddd - dd - mmmm - yyyy") & Space(3) & Format(Time, "hh:mm AMPM")
    MDISicmact.Show
    Exit Sub
AceptarErr:
    MsgBox Err.Description, vbInformation, "Aviso!"
End Sub

'TORE RCF1902190004 - 20190724
Private Sub AlertaVencimientoCartaFianza()
    Dim oConst As NConstSistemas
    Dim oCF As DCartaFianza
    Dim oRS As ADODB.Recordset
    Dim sMaquinaActiva As String
    Dim lsCorreoDestino As String, _
    lsContenidoH As String, _
    lsContenidoB As String, _
    lsContenidoF As String, _
    lsContenidoFinal As String
    
    Dim nItem As Integer
    
    Set oCF = New DCartaFianza
    Set oRS = New ADODB.Recordset
    
    Set oRS = oCF.EnvioCorreoCartaFianza(gsCodUser, Format(gdFecSis, "yyyyMMdd"))
    If Not (oRS.BOF And oRS.EOF) Then
        Set oConst = New NConstSistemas
        gsCorreoHost = oConst.LeeConstSistema(90)
        gsCorreoEnvia = oConst.LeeConstSistema(91)
        lsCorreoDestino = oConst.LeeConstSistema(99)
    
        lsContenidoH = "<html>" & _
                      "<head> <style> table,th, td { border: 0.5px outset #D4D0C8; background: #2121212; font-family: Calibri Light; font-size: 14px}</style> </head>" & _
                      "<body>" & _
                      "<h4>Listado de Cartas Fianzas vencidas</h4>" & _
                      "<p></p>" & _
                      "<table style='width:70%' cellpadding='0' cellspacing='0'>" & _
                      "<tr style='background: #981B1B; color: #FFFFFF' > <th style='width:20%'>Carta Fianza</th> <th style='width:20%'>Contrato</th> <th>Proveedor</th> <th style='width:10%'>Fecha de Vencimiento</th> </tr>"
        lsContenidoF = "</table> </body> </html>"
        lsContenidoB = ""
        For nItem = 1 To oRS.RecordCount
            lsContenidoB = lsContenidoB & "<tr style='padding:10px'> <td>" & oRS!cCartaFianza & "</td> <td>" & oRS!cContrato & "</td>" & _
                           "<td>" & oRS!cNombreProveedor & "</td> <td>" & oRS!dFechaVencimiento & "</td> </tr>"
            oRS.MoveNext
        Next
    
        lsContenidoFinal = lsContenidoH & lsContenidoB & lsContenidoF
        
        EnviarMail gsCorreoHost, gsCorreoEnvia, lsCorreoDestino, "Alerta Carta Fianza Vencidas", lsContenidoFinal
        'Fin Envio Correo ****************************************************
        Screen.MousePointer = 0
    End If
End Sub
'END TORE *********************************************************************


Private Sub cmdCancelar_Click()
    If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        End
    End If
End Sub

Private Sub Form_Load()
    Dim oAcceso As UAcceso
    Dim oCon As NConstSistemas
    Dim oImp As DImpresoras
    On Error GoTo ErrLoad
    
    Me.lblVer.Caption = "Versión : " & App.Major & "." & App.Minor & "." & App.Revision
     
    Set oAcceso = New UAcceso
    Set oCon = New NConstSistemas
    Set oImp = New DImpresoras
    
    ChDrive App.path
    ChDir App.path
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    gImpresora = oImp.GetImpreSetup(oImp.GetMaquina)
    
    If gImpresora = -1 Then
        MsgBox "Ud. debe asignar los caracteres de impresion por defecto para esta maquina.", vbInformation, "Aviso"
        frmCaracImpresion.Show 1
        
        gImpresora = oImp.GetImpreSetup(oImp.GetMaquina)
        
        If gImpresora = -1 Then
            MsgBox "Como Ud. no ha elegido los caracteres de impresion para esta maquina, se esta procediendo a asignarle el tipo EPSON, si Ud. desea luego puede modificarlo.", vbInformation, "Aviso"
            oImp.SetImpreSetup oImp.GetMaquina, gEPSON
        End If
    End If
    
    gcPDC = oCon.LeeConstSistema(gConstSistPDC)
    gcDominio = oCon.LeeConstSistema(gConstSistDominio)
    gsEmpresa = oCon.LeeConstSistema(gConstSistNombreAbrevCMAC)
    gsEmpresaCompleto = oCon.LeeConstSistema(gConstSistCMACNombreCompleto)
    gsEmpresaDireccion = oCon.LeeConstSistema(gConstSistCMACDireccion)
    gcCtaIGV = oCon.LeeConstSistema(gConstCtaIGV)
    gsRUC = oCon.LeeConstSistema(gConstSistCMACRuc)
    gbBitCentral = IIf(oCon.LeeConstSistema(gConstSistBitCentral) = "1", True, False)
    gbBitTCPonderado = IIf(oCon.LeeConstSistema(gConstSistBitTCPonderado) = "1", True, False)
    gbBitIGVCredFiscal = IIf(oCon.LeeConstSistema(gConstSistBitIGVxActivoCredFiscal) = "1", True, False)
    gbBitIGVCredFiscal = False
    gnIGV = oCon.GetValorImpuesto(gcCtaIGV)
    
    gsConexion = ""
    vsServerComunes = ""
    vsServerPersona = ""
    vsServerAdministracion = ""
    vsServerNegocio = ""
    vsServerImagenes = ""
    gcWINNT = "WinNT://"
    
    LblUsu.Caption = UCase(oAcceso.ObtenerUsuario)

    If LblUsu.Caption = "" Then
        MsgBox "El sistema operativo no puede reconocer su suario. Coordinar con el area de Sistemas.", vbInformation, "Aviso"
        End
    End If
    
    Call CtlUsuario.Inicio(Trim(LblUsu.Caption))
    gsCodAge = CtlUsuario.CodAgeAct
    gsNomAge = CtlUsuario.DescAgeAct
    gsCodCargo = Trim(CtlUsuario.PersCargoCod) 'EJVG20111217
    gsCodPersUser = CtlUsuario.PersCod
    
    If gsCodAge = "" Then
        MsgBox "El usuario no esta asignado a ninguna Agencia. Coordinar con el area de rrhh.", vbInformation, "Aviso"
        End
    End If
    
    Set oAcceso = Nothing
    
    If Not ValidaConfiguracionRegional Then
        MsgBox "Su actual CONFIGURACIÓN REGIONAL NO ES CORRECTA. Revísela y reinicie.", vbInformation, "Aviso"
        End
    End If
    
    If fgActualizaUltVersionEXE(CtlUsuario.CodAgeAct) = True Then ' Verifica si existe una actualizacion
       End
    End If
    
    RotateText 90, Picture1, "Times New Roman", 15, 25, 1700, "SICMACT"
     'ALPA 20090122 ***************************************************************************
    Set objPista = New COMManejador.Pista
    '*****************************************************************************************
    
    'ARLO2017066
    gsFechaVersion = "20191004" 'Cambiar la fecha cada vez que se compila
   
Exit Sub
ErrLoad:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CmdAceptar_Click
    End If
End Sub

