VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGarantiaTramiteLegal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRÁMITE LEGAL"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
   Icon            =   "frmGarantiaTramiteLegal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtZonaRegistral 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3960
      Width           =   6240
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4185
      TabIndex        =   11
      ToolTipText     =   "Cancelar"
      Top             =   3160
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      ToolTipText     =   "Aceptar"
      Top             =   3160
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   80
      TabIndex        =   12
      Top             =   80
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Trámite Legal"
      TabPicture(0)   =   "frmGarantiaTramiteLegal.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtInscripcionFecha"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtNotariaCod"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtNroPartidaRegistral"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbTramiteLegalTpo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmbZonaRegistral"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtGravamen"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtNotariaNombre"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmbTramiteLegalEstado"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmbTramiteLegalCobertura"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtNroAsiento"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtMoneda"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmbOficinaRegistral"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      Begin VB.ComboBox cmbOficinaRegistral 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1500
         Width           =   6255
      End
      Begin VB.TextBox txtMoneda 
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "txtPrincipal"
         Top             =   840
         Width           =   450
      End
      Begin VB.TextBox txtNroAsiento 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2560
         Width           =   1680
      End
      Begin VB.ComboBox cmbTramiteLegalCobertura 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2190
         Width           =   1455
      End
      Begin VB.ComboBox cmbTramiteLegalEstado 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2190
         Width           =   1695
      End
      Begin VB.TextBox txtNotariaNombre 
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "txtPrincipal"
         Top             =   1860
         Width           =   4425
      End
      Begin VB.TextBox txtGravamen 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   6600
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   840
         Width           =   1440
      End
      Begin VB.ComboBox cmbZonaRegistral 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1170
         Width           =   6255
      End
      Begin VB.ComboBox cmbTramiteLegalTpo 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   810
         Width           =   3255
      End
      Begin VB.TextBox txtNroPartidaRegistral 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   0
         Top             =   480
         Width           =   1680
      End
      Begin SICMACT.TxtBuscar txtNotariaCod 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   1860
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin MSMask.MaskEdBox txtInscripcionFecha 
         Height          =   285
         Left            =   6960
         TabIndex        =   23
         Top             =   2560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inscripción :"
         Height          =   195
         Left            =   4845
         TabIndex        =   22
         Top             =   2600
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tipo :"
         Height          =   195
         Left            =   6045
         TabIndex        =   21
         Top             =   2240
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "N° de Asiento :"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   2600
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estado :"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   2240
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Notaría :"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Oficina Registral :"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1575
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gravamen :"
         Height          =   195
         Left            =   5205
         TabIndex        =   16
         Top             =   855
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N° Partida Registral :"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   495
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Trámite :"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Zona Registral :"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1215
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmGarantiaTramiteLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************
'** Nombre : frmGarantiaTramiteLegal
'** Descripción : Para registro/consulta de Trámite Legal creado segun TI-ERS063-2014
'** Creación : EJVG, 20130220 11:25:00 AM
'************************************************************************************
Option Explicit
Dim fvMatOfiRegistral As Variant
Dim fsNumGarant As String
Dim fbPrimero As Boolean
Dim fnMoneda As Moneda

Dim fbRegistrar As Boolean
Dim fbEditar As Boolean
Dim fbConsultar As Boolean

Dim fvTramiteLegal As tTramiteLegal
Dim fvTramiteLegal_ULT As tTramiteLegal

Dim fbOk As Boolean

Public Function Registrar(ByVal psNumGarant As String, ByVal pbPrimero As Boolean, ByVal pnMoneda As Moneda, ByRef pvTramiteLegal As tTramiteLegal, ByRef pvTramiteLegal_ULT As tTramiteLegal) As Boolean
    fbRegistrar = True
    fsNumGarant = psNumGarant
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fvTramiteLegal = pvTramiteLegal
    fvTramiteLegal_ULT = pvTramiteLegal_ULT
    Show 1
    pvTramiteLegal = fvTramiteLegal

    Registrar = fbOk
End Function
Public Sub Consultar(ByVal psNumGarant As String, ByVal pnMoneda As Moneda, ByRef pvTramiteLegal As tTramiteLegal)
    fbConsultar = True
    fsNumGarant = psNumGarant
    fnMoneda = pnMoneda
    fvTramiteLegal = pvTramiteLegal
    Show 1
End Sub
Public Function Editar(ByVal psNumGarant As String, ByVal pbPrimero As Boolean, ByVal pnMoneda As Moneda, ByRef pvTramiteLegal As tTramiteLegal, ByRef pvTramiteLegal_ULT As tTramiteLegal) As Boolean
    fbEditar = True
    fsNumGarant = psNumGarant
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fvTramiteLegal = pvTramiteLegal
    fvTramiteLegal_ULT = pvTramiteLegal_ULT
    Show 1
    pvTramiteLegal = fvTramiteLegal

    Editar = fbOk
End Function

Private Sub cmbZonaRegistral_Click()
    Dim lnZonaID As Integer
    lnZonaID = val(Trim(Right(cmbZonaRegistral.Text, 3)))
    If lnZonaID > 0 Then
        RecuperaOficinaRegistralxZona lnZonaID
    End If
End Sub

Private Sub cmbZonaRegistral_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmbOficinaRegistral
    End If
End Sub
Private Sub Form_Load()
    fbOk = False
    Screen.MousePointer = 11
    
    CargarControles
    LimpiaControles

    If fbEditar Or fbConsultar Then
        txtNroPartidaRegistral.Text = fvTramiteLegal.sNroPartidaRegistral
        cmbTramiteLegalTpo.ListIndex = IndiceListaCombo(cmbTramiteLegalTpo, fvTramiteLegal.nTipoTramite)
        '''txtMoneda.Text = IIf(fnMoneda = gMonedaNacional, "S/.", IIf(fnMoneda = gMonedaExtranjera, "US$", "")) 'MARG ERS044
        txtMoneda.Text = IIf(fnMoneda = gMonedaNacional, gcPEN_SIMBOLO, IIf(fnMoneda = gMonedaExtranjera, "US$", ""))  'MARG ERS044
        txtGravamen.Text = Format(fvTramiteLegal.nGravamen, "#,##0.00")
        cmbZonaRegistral.ListIndex = IndiceListaCombo(cmbZonaRegistral, fvTramiteLegal.nZonaRegistralID)
        cmbOficinaRegistral.ListIndex = IndiceListaCombo(cmbOficinaRegistral, fvTramiteLegal.nOficinaRegistralID)
        txtNotariaCod.Text = fvTramiteLegal.sNotariaCod
        txtNotariaCod.psCodigoPersona = fvTramiteLegal.sNotariaCod
        txtNotariaNombre.Text = fvTramiteLegal.sNotariaNombre
        cmbTramiteLegalCobertura.ListIndex = IndiceListaCombo(cmbTramiteLegalCobertura, fvTramiteLegal.nTipoCobertura)
        cmbTramiteLegalEstado.ListIndex = IndiceListaCombo(cmbTramiteLegalEstado, fvTramiteLegal.nEstado)
        
        If fvTramiteLegal.nEstado = Inscrita Then
            txtNroAsiento.Text = fvTramiteLegal.sNroAsiento
            txtInscripcionFecha.Text = Format(fvTramiteLegal.dInscripcion, gsFormatoFechaView)
        End If
         
        If fbConsultar Then
            SSTab1.Enabled = False
            cmdAceptar.Enabled = False
        End If
    End If
    
    If Not fbPrimero Then
        txtNroPartidaRegistral.Enabled = False
    End If
    
    If fbRegistrar Then
        txtNroPartidaRegistral.Text = fvTramiteLegal_ULT.sNroPartidaRegistral
        cmbTramiteLegalTpo.ListIndex = IndiceListaCombo(cmbTramiteLegalTpo, fvTramiteLegal_ULT.nTipoTramite)
        'txtMoneda.Text = IIf(fnMoneda = gMonedaNacional, "S/.", IIf(fnMoneda = gMonedaExtranjera, "US$", ""))
        txtGravamen.Text = Format(fvTramiteLegal_ULT.nGravamen, "#,##0.00")
        cmbZonaRegistral.ListIndex = IndiceListaCombo(cmbZonaRegistral, fvTramiteLegal_ULT.nZonaRegistralID)
        cmbOficinaRegistral.ListIndex = IndiceListaCombo(cmbOficinaRegistral, fvTramiteLegal_ULT.nOficinaRegistralID)
        txtNotariaCod.Text = fvTramiteLegal_ULT.sNotariaCod
        txtNotariaCod.psCodigoPersona = fvTramiteLegal_ULT.sNotariaCod
        txtNotariaNombre.Text = fvTramiteLegal_ULT.sNotariaNombre
        cmbTramiteLegalCobertura.ListIndex = IndiceListaCombo(cmbTramiteLegalCobertura, fvTramiteLegal_ULT.nTipoCobertura)
        cmbTramiteLegalEstado.ListIndex = IndiceListaCombo(cmbTramiteLegalEstado, fvTramiteLegal_ULT.nEstado)
        
        If fvTramiteLegal_ULT.nEstado = Inscrita Then
            txtNroAsiento.Text = fvTramiteLegal_ULT.sNroAsiento
            txtInscripcionFecha.Text = Format(fvTramiteLegal_ULT.dInscripcion, gsFormatoFechaView)
        End If
        
        If fvTramiteLegal_ULT.bMigrado Then
            txtNroPartidaRegistral.Enabled = True
        End If
    End If
    
    If fbRegistrar Then
        Caption = "TRÁMITE LEGAL [ NUEVO ]"
    End If
    If fbConsultar Then
        Caption = "TRÁMITE LEGAL [ CONSULTAR ]"
    End If
    If fbEditar Then
        Caption = "TRÁMITE LEGAL [ EDITAR ]"
    End If
    
    Call CambiaTamañoCombo(cmbTramiteLegalTpo, 320)
    
    Screen.MousePointer = 0
End Sub
Private Sub CargarControles()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Dim oGar As New COMDCredito.DCOMGarantia
    Dim rs As New ADODB.Recordset
    
    'Cargar Tipo Tramite Legal
    Set rs = oGar.RecuperaProxTramiteLegal(fsNumGarant, fvTramiteLegal_ULT.dFecha, IIf(fbRegistrar, 1, IIf(fbConsultar, 2, IIf(fbEditar, 3, 0))))
    cmbTramiteLegalTpo.Clear
    Do While Not rs.EOF
        cmbTramiteLegalTpo.AddItem rs!cdescripcion & Space(100) & rs!nCodigo
        rs.MoveNext
    Loop
    'Cargar Moneda
    '''txtMoneda.Text = IIf(fnMoneda = gMonedaNacional, "S/.", IIf(fnMoneda = gMonedaExtranjera, "US$", "")) 'MARG ERS044
    txtMoneda.Text = IIf(fnMoneda = gMonedaNacional, gcPEN_SIMBOLO, IIf(fnMoneda = gMonedaExtranjera, "US$", "")) 'MARG ERS044
    If fnMoneda = gMonedaNacional Then
        txtGravamen.BackColor = &H80000005
    ElseIf fnMoneda = gMonedaExtranjera Then
        txtGravamen.BackColor = &HC0FFC0
    End If
    
    'Carga Zona Registral
    cmbZonaRegistral.Clear
    Set rs = oGar.RecuperaZonaRegistral()
    Do While Not rs.EOF
        cmbZonaRegistral.AddItem rs!cZona & Space(100) & rs!nId
        rs.MoveNext
    Loop
    
    'Cargar Oficina Registral
    'Set rs = oGar.RecuperaOficinaZonaRegistral()
    'ReDim fvMatOfiRegistral(4, 0)
    cmbOficinaRegistral.Clear
    'Do While Not rs.EOF
    '    ReDim Preserve fvMatOfiRegistral(4, 1 To rs.Bookmark)
    '    fvMatOfiRegistral(1, rs.Bookmark) = rs!nOficinaRegistralID
    '    fvMatOfiRegistral(2, rs.Bookmark) = rs!nZonaRegistralID
    '    fvMatOfiRegistral(3, rs.Bookmark) = rs!cOficinaRegistralDesc
    '    fvMatOfiRegistral(4, rs.Bookmark) = rs!cZonaRegistralDesc
    '    cmbOficinaRegistral.AddItem rs!cOficinaRegistralDesc & Space(100) & rs!nOficinaRegistralID
    '    rs.MoveNext
    'Loop
    'Cargar Estado Tramite Legal
    Set rs = oCons.RecuperaConstantes(gGarantiaTramiteLegalEstado)
    Call Llenar_Combo_con_Recordset(rs, cmbTramiteLegalEstado)
    'Cargar Cobertura Tramite Legal
    Set rs = oCons.RecuperaConstantes(gGarantiaTramiteLegalCobertura)
    Call Llenar_Combo_con_Recordset(rs, cmbTramiteLegalCobertura)
    
    RSClose rs
    Set oCons = Nothing
End Sub
Private Sub LimpiaControles()
    txtNroPartidaRegistral.Text = ""
    cmbTramiteLegalTpo.ListIndex = -1
    '''txtMoneda.Text = IIf(fnMoneda = gMonedaNacional, "S/.", IIf(fnMoneda = gMonedaExtranjera, "US$", "")) 'MARG ERS044
    txtMoneda.Text = IIf(fnMoneda = gMonedaNacional, gcPEN_SIMBOLO, IIf(fnMoneda = gMonedaExtranjera, "US$", "")) 'MARG ERS044
    txtGravamen.Text = "0.00"
    cmbZonaRegistral.ListIndex = -1
    cmbOficinaRegistral.ListIndex = -1
    txtZonaRegistral.Text = ""
    txtNotariaCod.Text = ""
    txtNotariaNombre.Text = ""
    cmbTramiteLegalEstado.ListIndex = -1
    cmbTramiteLegalCobertura.ListIndex = -1
    txtNroAsiento.Text = ""
    txtInscripcionFecha.Text = "__/__/____"
End Sub
Private Sub cmbTramiteLegalCobertura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not EnfocaControl(txtNroAsiento) Then
            If Not EnfocaControl(txtInscripcionFecha) Then
                EnfocaControl cmdAceptar
            End If
        End If
    End If
End Sub
Private Sub cmbTramiteLegalEstado_Click()
    If CInt(Trim(Right(cmbTramiteLegalEstado.Text, 3))) = eGarantiaTramiteLegalEstado.Pendiente Then
        txtNroAsiento.Text = ""
        txtNroAsiento.Enabled = False
        txtInscripcionFecha.Text = "__/__/____"
        txtInscripcionFecha.Enabled = False
    Else
        txtNroAsiento.Enabled = True
        txtInscripcionFecha.Text = Format(gdFecSis, gsFormatoFechaView)
        txtInscripcionFecha.Enabled = True
    End If
End Sub
Private Sub cmbTramiteLegalEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not EnfocaControl(cmbTramiteLegalCobertura) Then
            If Not EnfocaControl(txtNroAsiento) Then
                If Not EnfocaControl(txtInscripcionFecha) Then
                    EnfocaControl cmdAceptar
                End If
            End If
        End If
    End If
End Sub
Private Sub cmbTramiteLegalTpo_Click()
    If CInt(Trim(Right(cmbTramiteLegalTpo.Text, 3))) = eGarantiaTramiteLegalTipo.BloqueoRegistral Then
        cmbTramiteLegalCobertura.ListIndex = -1
        cmbTramiteLegalCobertura.Enabled = False
    Else
        cmbTramiteLegalCobertura.Enabled = True
    End If
End Sub
Private Sub cmdAceptar_Click()
    On Error GoTo ErrAceptar
    If Not validarTramite Then Exit Sub
    
    fvTramiteLegal.sUltimaActualizacion = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    fvTramiteLegal.dFecha = fgFechaHoraMovDate(fvTramiteLegal.sUltimaActualizacion)
    fvTramiteLegal.sNroPartidaRegistral = txtNroPartidaRegistral.Text
    fvTramiteLegal.nTipoTramite = CInt(Trim(Right(cmbTramiteLegalTpo.Text, 3)))
    fvTramiteLegal.cTipoTramite = Trim(Mid(cmbTramiteLegalTpo.Text, 1, Len(cmbTramiteLegalTpo) - 3))
    fvTramiteLegal.nGravamen = CCur(txtGravamen.Text)
    fvTramiteLegal.nZonaRegistralID = CInt(Trim(Right(cmbZonaRegistral.Text, 3)))
    fvTramiteLegal.cZonaRegistralNombre = Trim(Mid(cmbZonaRegistral.Text, 1, Len(cmbZonaRegistral) - 3))
    fvTramiteLegal.nOficinaRegistralID = CInt(Trim(Right(cmbOficinaRegistral.Text, 3)))
    fvTramiteLegal.cOficinaRegistralNombre = Trim(Mid(cmbOficinaRegistral.Text, 1, Len(cmbOficinaRegistral) - 3))
    fvTramiteLegal.sNotariaCod = txtNotariaCod.psCodigoPersona
    fvTramiteLegal.sNotariaNombre = txtNotariaNombre.Text
    fvTramiteLegal.nEstado = CInt(Trim(Right(cmbTramiteLegalEstado.Text, 3)))
    fvTramiteLegal.cEstado = Trim(Mid(cmbTramiteLegalEstado.Text, 1, Len(cmbTramiteLegalEstado) - 3))
    fvTramiteLegal.nTipoCobertura = IIf(cmbTramiteLegalCobertura.ListIndex = -1, 0, val(Trim(Right(cmbTramiteLegalCobertura.Text, 3))))
    If fvTramiteLegal.nEstado = Inscrita Then
        fvTramiteLegal.sNroAsiento = txtNroAsiento.Text
        fvTramiteLegal.dInscripcion = CDate(IIf(IsDate(txtInscripcionFecha.Text), txtInscripcionFecha.Text, "1900-01-01"))
    End If
        
    fbOk = True
    Unload Me
    Exit Sub
ErrAceptar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdCancelar_Click()
    fbOk = False
    Unload Me
End Sub
Private Sub txtGravamen_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtGravamen, KeyAscii, 15)
    If KeyAscii = 13 Then
        EnfocaControl cmbZonaRegistral
    End If
End Sub
Private Sub txtGravamen_LostFocus()
    txtGravamen.Text = Format(txtGravamen, "#,##0.00")
End Sub
Private Sub txtInscripcionFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdAceptar
    End If
End Sub
Private Sub txtNotariaCod_EmiteDatos()
    txtNotariaNombre.Text = ""
    If txtNotariaCod.Text <> "" Then
        txtNotariaNombre.Text = txtNotariaCod.psDescripcion
        EnfocaControl cmbTramiteLegalEstado
    End If
End Sub
Private Sub txtNotariaNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmbTramiteLegalEstado
    End If
End Sub
Private Sub txtNroAsiento_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl txtInscripcionFecha
    End If
End Sub
Private Sub txtNroAsiento_LostFocus()
    txtNroAsiento.Text = Trim(txtNroAsiento.Text)
End Sub
Private Sub txtNroPartidaRegistral_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl cmbTramiteLegalTpo
    End If
End Sub
Private Sub txtNroPartidaRegistral_LostFocus()
    txtNroPartidaRegistral.Text = Trim(UCase(txtNroPartidaRegistral.Text))
End Sub
Private Sub cmbTramiteLegalTpo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtGravamen
    End If
End Sub
Private Sub cmbOficinaRegistral_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtNotariaCod
    End If
End Sub
'Private Sub cmbOficinaRegistral_Click()
    'Dim lnOfiRegistralID As Integer
    'txtZonaRegistral.Text = ""
    'If IsArray(fvMatOfiRegistral) Then
    '    txtZonaRegistral.Text = fvMatOfiRegistral(4, cmbOficinaRegistral.ListIndex + 1)
    'End If
'End Sub
Private Sub txtRegistroFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdAceptar
    End If
End Sub
Private Sub txtZonaRegistral_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtNotariaCod
    End If
End Sub
Private Function validarTramite() As Boolean
    Dim oDGar As COMDCredito.DCOMGarantia
    Dim lsFecha As String
    Dim lnTpoTramite As eGarantiaTramiteLegalTipo
    Dim lnGravamen As Currency
    
    If Len(Trim(txtNroPartidaRegistral.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Nro. de Partida Registral", vbInformation, "Aviso"
        EnfocaControl txtNroPartidaRegistral
        Exit Function
    End If
    If cmbTramiteLegalTpo.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el Tipo de Trámite", vbInformation, "Aviso"
        EnfocaControl cmbTramiteLegalTpo
        Exit Function
    Else
        lnTpoTramite = CInt(Right(cmbTramiteLegalTpo.Text, 3))
    End If
    If Len(txtMoneda.Text) = 0 Then
        MsgBox "No se ha podido definir la moneda de la Garantía", vbInformation, "Aviso"
        Exit Function
    End If
    If Not IsNumeric(txtGravamen.Text) Then
        MsgBox "Ud. debe especificar el Gravamen", vbInformation, "Aviso"
        EnfocaControl txtGravamen
        Exit Function
    Else
        If CCur(txtGravamen.Text) <= 0 Then
            MsgBox "El monto de Gravamen debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtGravamen
            Exit Function
        Else
            lnGravamen = CCur(txtGravamen.Text)
        End If
        'Valida el monto según Tipo de Trámite [ampliado o modificado]
        If Not fvTramiteLegal_ULT.bMigrado Then
            If lnTpoTramite = AmpliacionGarantia Then
                If lnGravamen <= fvTramiteLegal_ULT.nGravamen Then
                    MsgBox "Por ser Ampliación el monto de Gravamen debe ser mayor a " & Format(fvTramiteLegal_ULT.nGravamen, "#,##0.00"), vbInformation, "Aviso"
                    EnfocaControl txtGravamen
                    Exit Function
                End If
            ElseIf lnTpoTramite = ModificacionGarantia Then
                If lnGravamen >= fvTramiteLegal_ULT.nGravamen Then
                    MsgBox "Por ser Modificación el monto de Gravamen debe ser menor a " & Format(fvTramiteLegal_ULT.nGravamen, "#,##0.00"), vbInformation, "Aviso"
                    EnfocaControl txtGravamen
                    Exit Function
                End If
            ElseIf lnTpoTramite = LevantamientoGarantia Then
                If lnGravamen <> fvTramiteLegal_ULT.nGravamen Then
                    MsgBox "Por ser Levantamiento el monto de Gravamen debe ser igual al del último Trámite Legal: " & Format(fvTramiteLegal_ULT.nGravamen, "#,##0.00"), vbInformation, "Aviso"
                    EnfocaControl txtGravamen
                    Exit Function
                End If
            End If
        End If
    End If
    If cmbZonaRegistral.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la Zona Registral", vbInformation, "Aviso"
        EnfocaControl cmbZonaRegistral
        Exit Function
    End If
    If cmbOficinaRegistral.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la Oficina Registral", vbInformation, "Aviso"
        EnfocaControl cmbOficinaRegistral
        Exit Function
    End If
    If Len(txtNotariaCod.Text) <> 13 Or Len(txtNotariaCod.psCodigoPersona) <> 13 Then
        MsgBox "Ud. debe seleccionar a la Notaría", vbInformation, "Aviso"
        EnfocaControl txtNotariaCod
        Exit Function
    End If
    If cmbTramiteLegalEstado.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el Estado del presente trámite", vbInformation, "Aviso"
        EnfocaControl cmbTramiteLegalEstado
        Exit Function
    End If
    If cmbTramiteLegalCobertura.Enabled Then
        If cmbTramiteLegalCobertura.ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar el Tipo de Cobertura", vbInformation, "Aviso"
            EnfocaControl cmbTramiteLegalCobertura
            Exit Function
        End If
    End If
    If txtNroAsiento.Enabled Then
        If Len(txtNroAsiento.Text) = 0 Then
            MsgBox "Ud. debe especificar el Nro. de Asiento", vbInformation, "Aviso"
            EnfocaControl txtNroAsiento
            Exit Function
        End If
        If fbRegistrar Then
            Set oDGar = New DCOMGarantia
            If oDGar.ExisteNroAsientoGarantia(fsNumGarant, fvTramiteLegal.dFecha, txtNroAsiento.Text) Then
                MsgBox "El Nro. de Asiento de la Garantía ya fue registrado anteriormente." & Chr(13) & Chr(13) & "Por favor modifiquelo e inténtelo de nuevo.", vbInformation, "Aviso"
                EnfocaControl txtNroAsiento
                Set oDGar = Nothing
                Exit Function
            End If
        End If
    End If
    If txtInscripcionFecha.Enabled Then
        lsFecha = ValidaFecha(txtInscripcionFecha.Text)
        If Len(lsFecha) > 0 Then
            MsgBox lsFecha, vbInformation, "Aviso"
            EnfocaControl txtInscripcionFecha
            Exit Function
        End If
        If Not fvTramiteLegal_ULT.bMigrado Then
            If CDate(txtInscripcionFecha.Text) <= fvTramiteLegal_ULT.dInscripcion And lnTpoTramite <> LevantamientoGarantia Then
                MsgBox "La actual fecha de Inscripción no puede ser menor o igual a la última Inscripción: " & Format(fvTramiteLegal_ULT.dInscripcion, gsFormatoFechaView), vbInformation, "Aviso"
                EnfocaControl txtInscripcionFecha
                Exit Function
            End If
        End If
        If CDate(txtInscripcionFecha.Text) > gdFecSis Then
            MsgBox "La fecha de Inscripción no debe ser mayor a la fecha del Sistema", vbInformation, "Aviso"
            EnfocaControl txtInscripcionFecha
            Exit Function
        End If
    End If
    
    Set oDGar = Nothing
    validarTramite = True
End Function
Private Sub RecuperaOficinaRegistralxZona(ByVal pnZonaID As Integer)
    Dim oGar As New COMDCredito.DCOMGarantia
    Dim rs As New ADODB.Recordset
    
    Screen.MousePointer = 11
    Set rs = oGar.RecuperaOficinaRegistralxZona(pnZonaID)
    
    cmbOficinaRegistral.Clear
    Do While Not rs.EOF
        cmbOficinaRegistral.AddItem rs!cOficinaNombre & Space(200) & rs!nOficinaID
        rs.MoveNext
    Loop
    RSClose rs
    Screen.MousePointer = 0
    Set oGar = Nothing
End Sub


