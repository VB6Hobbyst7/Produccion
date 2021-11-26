VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SICMAC Finanzas"
   ClientHeight    =   2250
   ClientLeft      =   4530
   ClientTop       =   4965
   ClientWidth     =   3930
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   3930
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H005CB3DC&
      ForeColor       =   &H00FF0000&
      Height          =   1785
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   120
      Width           =   465
      Begin VB.Image Image1 
         Height          =   390
         Left            =   15
         Picture         =   "frmLogin.frx":030A
         Stretch         =   -1  'True
         Top             =   -15
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1890
      Left            =   555
      TabIndex        =   3
      Top             =   0
      Width           =   3315
      Begin Sicmact.Usuario CtlUsuario 
         Left            =   60
         Top             =   150
         _ExtentX        =   820
         _ExtentY        =   820
      End
      Begin VB.TextBox TxtClave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1095
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   0
         ToolTipText     =   "Ingrese su Clave Secreta"
         Top             =   900
         Width           =   2070
      End
      Begin VB.CommandButton CmdCancelar 
         BackColor       =   &H00D0CECC&
         Caption         =   "&Salir"
         Height          =   390
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1365
         Width           =   1140
      End
      Begin VB.CommandButton CmdAceptar 
         BackColor       =   &H00D0CECC&
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1365
         Width           =   1140
      End
      Begin VB.Label frmLogin 
         Caption         =   "Acceso al Sistema"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1215
         TabIndex        =   8
         Top             =   210
         Width           =   1875
      End
      Begin VB.Label Label3 
         Caption         =   "Clave :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         TabIndex        =   6
         Top             =   900
         Width           =   825
      End
      Begin VB.Label LblUsu 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   1095
         TabIndex        =   5
         Top             =   555
         Width           =   2070
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   135
         TabIndex        =   4
         Top             =   555
         Width           =   990
      End
   End
   Begin VB.Label lblVer 
      Caption         =   "Label2"
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   3705
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oAcceso As UAcceso

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
    For Each Ctl In frmMdiMain.Controls
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


'Private Sub CargaMenuMDIMain(ByVal poAcceso As UAcceso)
'Dim Ctl As Control
'Dim sTipo As String
'Dim nPos As Integer
'Dim sCadMenuTemp As String
'Dim nPunt As Integer
'
'    Call ActualizaMenuActivos(nPunt)
'On Error GoTo ErrCarga
'    For Each Ctl In frmMdiMain.Controls
'        sTipo = TypeName(Ctl)
'        If sTipo = "Menu" Then
'            If UbicaMenuActivos(Ctl.Name, Format(Ctl.Index, "00"), 0) Then
'                Ctl.Visible = True
'            Else
'                Ctl.Visible = False
'            End If
'        End If
'    Next
'
'Exit Sub
'ErrCarga:
'  Err.Clear
'End Sub

Private Sub CargaMenuMDIMain(ByVal poAcceso As UAcceso)
Dim Ctl As Control
Dim sTipo As String
Dim nPos As Integer
Dim sCadMenuTemp As String
Dim nPunt As Integer

    Call ActualizaMenuActivos(nPunt)
On Error Resume Next
    For Each Ctl In frmMdiMain.Controls
        sTipo = TypeName(Ctl)
        If sTipo = "Menu" Then
            If UbicaMenuActivos(Ctl.Name, Format(Ctl.Index, "00"), 0) Then
                Ctl.Visible = True
            Else
               Ctl.Visible = False
            End If
        End If
    Next

Exit Sub
ErrCarga:
  Err.Clear
End Sub
Private Sub cmdAceptar_Click()
Dim sTitulo As String
Dim R As ADODB.Recordset
Dim i As Integer
Dim Y As Integer
Dim oConec As COMConecta.DCOMConecta
Dim sSql As String

Dim oImpuesto As DImpuesto
Set oImpuesto = New DImpuesto


On Error GoTo AceptarErr

    Screen.MousePointer = 11
    If InStr(1, Printer.Port, ":", vbTextCompare) > 0 Then
        sLPT = Mid(Printer.Port, 1, InStr(1, Printer.Port, ":", vbTextCompare) - 1)
    Else
        sLPT = "LPT1"
    End If

    
    Call CargaVarSistema(True)
    Set oAcceso = New UAcceso
    
    If Not oAcceso.InterconexionCorrecta Then
        MsgBox "No se puede Establecer la Interconexion con el Servidor" & oImpresora.gPrnSaltoLinea & "Consulte con el Area de Sistemas", vbCritical, "Conexion SICMACT"
        Set oAcceso = Nothing
        TxtClave.Text = ""
        TxtClave.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If Not oAcceso.ClaveIncorrectaNT(Trim(LblUsu.Caption), Trim(TxtClave.Text), gsDominio) Then
        MsgBox "Clave Incorrecta, Ingrese su Clave Nuevamente ", vbCritical, "Conexion SICMACT"
        Set oAcceso = Nothing
        fEnfoque TxtClave
        TxtClave.Text = ""
        TxtClave.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = False
    
    
    gnIGVValor = oImpuesto.CargaImpuestoFechaValor(gcCtaIGV, gdFecSis) / 100
    
    '************************************************************************
    Set R = oAcceso.DameItemsMenu
    
    i = 0
    ReDim MatMenuItems(0)
    ReDim Preserve MatMenuItems(i + 1)
    MatMenuItems(i).nId = i
    MatMenuItems(i).sCodigo = Trim(R!cCodigo)
    MatMenuItems(i).sCaption = Trim(R!cDescrip)
    MatMenuItems(i).sName = Trim(R!cname)
    MatMenuItems(i).sIndex = Right(R!cname, 2)
    MatMenuItems(i).bCheck = False
    MatMenuItems(i).nNivel = 1
    MatMenuItems(i).nPuntDer = -1
    MatMenuItems(i).nPuntAbajo = -1
    i = i + 1
    Y = i
    R.MoveNext
    Call CargaMenuArbol(R, i, Y)
    R.Close
    
    Call oAcceso.CargaMenu(gsDominio, LblUsu.Caption)
    'Call CargaMenu(oAcceso)
    Call CargaMenuMDIMain(oAcceso)
    
    'Habilita Permiso para Operaciones y Reportes
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "Select * from OpeTpo Order by cOpeCod"
    Set R = oConec.CargaRecordSet(sSql)
    Y = 0
    Do While Not R.EOF
        If oAcceso.TienePermiso(R!cOpeCod, "", True) Then
            Y = Y + 1
            MatOperac(Y - 1, 0) = R!cOpeCod
            MatOperac(Y - 1, 1) = R!cOpeDesc
            MatOperac(Y - 1, 2) = R!cOpeGruCod
            MatOperac(Y - 1, 3) = R!cOpeVisible
            MatOperac(Y - 1, 4) = R!nOpeNiv
        End If
        R.MoveNext
    Loop
    NroRegOpe = Y
    oConec.CierraConexion
    Set oAcceso = Nothing
  
  ' Obtiene la impresora predeterminada
    Dim sImpresora As String
    Dim lnPos As Long
    sImpresora = Printer.DeviceName
    If Left(sImpresora, 2) <> "\\" Then
        lnPos = InStr(1, Printer.Port, ":", vbTextCompare)
        If lnPos > 0 Then
            sLPT = Mid(Printer.Port, 1, lnPos - 1)
        Else
            sLPT = "LPT1"
        End If
 '   Else
 '       sLPT = frmImpresora.EliminaEspacios(sImpresora)
    End If
    
'    DeshabilitaOpeacionesPendientes
    
'    MsgBox "Por favor Configure su Impresora antes de Empezar sus operaciones", vbInformation, "Aviso"
'   frmImpresora.Show 1
  
    
    
    Screen.MousePointer = 0
    Unload Me
    'frmVideo.Show 1
    'frmMdiMain.Caption = "SICMACT " & Space(15) & Trim(gsNomAge) & Space(10) & gsCodUser & Space(2) & gsServerName & "\" & gsDBName & Space(5) & Format(gdFecSis, "dd/mm/yyyy")
    frmMdiMain.Show
Exit Sub
AceptarErr:
    MsgBox Err.Description, vbInformation, "Aviso!"
    
End Sub

Private Sub cmdCancelar_Click()
    If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        End
    End If
End Sub

Private Sub Form_Load()
Dim oAcceso As UAcceso
On Error GoTo ErrLogin
    Me.lblVer.Caption = "Vesrion : " & App.Major & "." & App.Minor & "." & App.Revision
    
    
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    Dim oCon As NConstSistemas
    Set oCon = New NConstSistemas
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    gImpresora = oImp.GetImpreSetup(oImp.GetMaquina)
    If gImpresora = 0 Then
        MsgBox "Ud. debe asignar los caracteres de impresion por defecto para esta maquina.", vbInformation, "Aviso"
        frmCaracImpresion.Show 1
        
        gImpresora = oImp.GetImpreSetup(oImp.GetMaquina)
        If gImpresora = 0 Then
            MsgBox "Como Ud. no ha elegido los caracteres de impresion para esta maquina, se esta procediendo a asignarle el tipo EPSON, si Ud. desea luego puede modificarlo.", vbInformation, "Aviso"
            oImp.SetImpreSetup oImp.GetMaquina, gEPSON
        End If
    End If
    Set oAcceso = New UAcceso
    LblUsu.Caption = UCase(oAcceso.ObtenerUsuario)
    Set oAcceso = Nothing
    If Not ValidaConfiguracionRegional Then
        MsgBox "Su actual CONFIGURACION REGIONAL NO ES CORRECTA. Revísela y reinicie.", vbCritical, "Error"
        End
    End If
    
    gcPDC = oCon.LeeConstSistema(gConstSistPDC)
    gcDominio = oCon.LeeConstSistema(gConstSistDominio)
    gsEmpresa = oCon.LeeConstSistema(gConstSistNombreAbrevCMAC)
    gsEmpresaCompleto = oCon.LeeConstSistema(gConstSistCMACNombreCompleto)
    gsEmpresaDireccion = oCon.LeeConstSistema(gConstSistCMACDireccion)
    gsRUC = oCon.LeeConstSistema(gConstSistCMACRuc)
    gbBitCentral = IIf(oCon.LeeConstSistema(gConstSistBitCentral) = "1", True, False)
    gbBitTCPonderado = IIf(oCon.LeeConstSistema(gConstSistBitTCPonderado) = "1", True, False)
    
    cCtaDetraccionProvision = oCon.LeeConstSistema(166)
'    gnDocCuentaPendiente = 80
    gnTipoCambioEuro = 4.052456 'Tipo de Cambio a Euros para Adeudados
    gsProyectoActual = oCon.LeeConstSistema(300)
    gsMesCerrado = oCon.LeeConstSistema(10)
    Call CtlUsuario.Inicio(Trim(LblUsu.Caption))
    gsCodAge = CtlUsuario.CodAgeAct
    gsCodAgeAsig = CtlUsuario.CodAgeAsig
    gsCodArea = CtlUsuario.cAreaCodAct
    gsNomAge = CtlUsuario.DescAgeAct
    gsCodUser = Trim(LblUsu.Caption)
    
    If gsCodAge = "" Then
        MsgBox "El usuario no esta asignado a ninguna Agencia. Coordinar con el area de RRHH.", vbInformation, "Aviso"
        End
    End If
    
'    If fgActualizaUltVersionEXE(CtlUsuario.CodAgeAct) = True Then  ' Verifica si existe una actualizacion
'       End
'    End If
    
    'RotateText 90, Picture1, "Times New Roman", 15, 25, 1700, "NURIA"
    RotateText 90, Picture1, "Times New Roman", 15, 25, 1700, "CMAC-T"
    
Exit Sub
ErrLogin:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        Call cmdAceptar_Click
    End If
End Sub
