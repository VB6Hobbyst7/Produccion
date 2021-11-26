VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRemesaAgenciaToIFiAgencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REMESA ENTRE AGENCIAS - DEVOLUCIÓN A BANCO"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   Icon            =   "frmRemesaAgenciaToIFiAgencia.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   5865
      Left            =   80
      TabIndex        =   16
      Top             =   0
      Width           =   8055
      Begin VB.Frame fraMoneda 
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   795
         Left            =   120
         TabIndex        =   29
         Top             =   180
         Width           =   2115
         Begin VB.OptionButton optMoneda 
            Caption         =   "Moneda Nacional"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   1680
         End
         Begin VB.OptionButton optMoneda 
            Caption         =   "Moneda Extranjera"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   2
            Top             =   500
            Width           =   1695
         End
      End
      Begin VB.Frame fraDestino 
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2175
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   7815
         Begin VB.Frame fraLinea 
            Height          =   1935
            Left            =   1650
            TabIndex        =   30
            Top             =   60
            Width           =   15
         End
         Begin VB.OptionButton optDestino 
            Caption         =   "Inst. Financiera"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   645
            Width           =   1455
         End
         Begin VB.OptionButton optDestino 
            Caption         =   "Agencia"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Frame fraDocumento 
            Caption         =   "Emisión de Documento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   735
            Left            =   1800
            TabIndex        =   24
            Top             =   1200
            Width           =   5895
            Begin VB.ComboBox cboDocumento 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   280
               Width           =   2400
            End
            Begin VB.TextBox txtDocumentoNro 
               Height          =   300
               Left            =   3000
               TabIndex        =   8
               Top             =   280
               Width           =   2280
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "N° :"
               Height          =   195
               Left            =   2640
               TabIndex        =   25
               Top             =   300
               Width           =   270
            End
         End
         Begin SICMACT.TxtBuscar txtDestinoCod 
            Height          =   300
            Left            =   1800
            TabIndex        =   6
            Top             =   480
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   529
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
         End
         Begin VB.Label lblDestinoDesc1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4440
            TabIndex        =   31
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label lblDestino 
            AutoSize        =   -1  'True
            Caption         =   "@Destino"
            Height          =   195
            Left            =   1800
            TabIndex        =   27
            Top             =   255
            Width           =   705
         End
         Begin VB.Label lblDestinoDesc2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1800
            TabIndex        =   26
            Top             =   840
            Width           =   5895
         End
      End
      Begin VB.Frame fraOrigen 
         Caption         =   "Origen"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   7815
         Begin SICMACT.TxtBuscar txtAreaAgeCod 
            Height          =   300
            Left            =   1395
            TabIndex        =   3
            Top             =   240
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   529
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
            sTitulo         =   ""
         End
         Begin VB.Label Label6 
            Caption         =   "Area - Agencia :"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label lblAreaAgeDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2925
            TabIndex        =   21
            Top             =   240
            Width           =   4755
         End
      End
      Begin VB.Frame fraGlosa 
         Caption         =   "Glosa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1095
         Left            =   120
         TabIndex        =   19
         Top             =   4200
         Width           =   3855
         Begin VB.TextBox txtGlosa 
            Height          =   660
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   240
            Width           =   3480
         End
      End
      Begin VB.Frame fraTransporte 
         Caption         =   "Transporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1095
         Left            =   4080
         TabIndex        =   18
         Top             =   4200
         Width           =   3855
         Begin VB.OptionButton optTransporte 
            Caption         =   "Propio"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   315
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optTransporte 
            Caption         =   "Terceros"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cboTerceros 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   600
            Width           =   2520
         End
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6180
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   5415
         Width           =   1710
      End
      Begin VB.Frame fraFecha 
         Caption         =   "Fecha"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   795
         Left            =   6345
         TabIndex        =   17
         Top             =   180
         Width           =   1575
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   330
            Left            =   240
            TabIndex        =   0
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Monto :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   5475
         TabIndex        =   28
         Top             =   5460
         Width           =   660
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   5355
         Top             =   5400
         Width           =   2565
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   350
      Left            =   7080
      TabIndex        =   15
      Top             =   5895
      Width           =   1050
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   350
      Left            =   6000
      TabIndex        =   14
      Top             =   5895
      Width           =   1050
   End
End
Attribute VB_Name = "frmRemesaAgenciaToIFiAgencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'** Nombre : frmRemesaAgenciaToIFiAgencia
'** Descripción : Formulario para el remesas de Agencias a Agencias o a Inst. Financieras
'** Creación : EJVG, 20140630 11:00:00 AM
'****************************************************************************************
Option Explicit
Dim fnMoneda As Moneda
Dim rsAreaAgencia As ADODB.Recordset
Dim fsCtaContCodD As String
Dim fsCtaContCodH As String
Dim fbAceptar As Boolean
Dim fsCtaContCodEfecTrans As String

Private Sub CargarControles()
    Dim oPer As New DPersonas
    Dim oNCont As New NConstSistemas
    Dim oOpe As New COMNCajaGeneral.NCOMCajaGeneral
    Dim oDOperacion As New clases.DOperacion
    Dim rsPer As New ADODB.Recordset
    Dim rsDocumento As New ADODB.Recordset
    
    On Error GoTo ErrCargarControles
    Screen.MousePointer = 11
    'Carga Cuentas Contables
    fsCtaContCodD = oNCont.LeeConstSistema(474)
    fsCtaContCodH = oNCont.LeeConstSistema(475)
    fsCtaContCodEfecTrans = fsCtaContCodD
    'Carga Area-Agencia
    Set rsAreaAgencia = New ADODB.Recordset
    Set rsAreaAgencia = oOpe.GetOpeObj(gsOpeCod, "1")
    'Carga Documentos
    Set rsDocumento = oDOperacion.CargaOpeDoc(gsOpeCod)
    cboDocumento.Clear
    Do While Not rsDocumento.EOF
        cboDocumento.AddItem Mid(rsDocumento!cDocDesc & Space(100), 1, 100) & Space(200) & rsDocumento!nDocTpo
        rsDocumento.MoveNext
    Loop
    'Transporte Terceros
    cboTerceros.Clear
    Set rsPer = oPer.ListaPersonaxRol(13)
    If Not rsPer.EOF Then
        Do While Not rsPer.EOF
            cboTerceros.AddItem rsPer!cPersNombre & Space(200) & rsPer!cPersCod
            rsPer.MoveNext
        Loop
        CambiaTamañoCombo cboTerceros, 300
    End If
    
    Set oDOperacion = Nothing
    Set rsDocumento = Nothing
    Set oOpe = Nothing
    Set rsPer = Nothing
    Set oPer = Nothing
    Set oNCont = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrCargarControles:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub Limpiar()
    Dim oOpe As New COMNCajaGeneral.NCOMCajaGeneral
    fbAceptar = False
    txtFecha.Text = Format(gdFecSis, gsFormatoFechaView)
    optMoneda.iTem(0).value = True
    OptMoneda_Click (0)
    txtAreaAgeCod.Text = "026" & Right(gsCodAge, 2)
    lblAreaAgeDesc.Caption = gsNomAge
    optDestino.iTem(0).value = True
    optDestino_Click (0)
    cboDocumento.ListIndex = -1
    txtDocumentoNro.Text = ""
    txtGlosa.Text = ""
    optTransporte.iTem(0).value = True
    optTransporte_Click (0)
    cboTerceros.ListIndex = -1
    txtMonto.Text = "0.00"
    Set oOpe = Nothing
End Sub
Private Sub cboTerceros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtMonto
    End If
End Sub

Private Sub cmdSalir_Click()
    fbAceptar = False
    Unload Me
End Sub
Private Sub Form_Load()
    CargarControles
    Limpiar
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not fbAceptar Then
        If MsgBox("¿Está seguro de Salir sin grabar la Operación?", vbInformation + vbYesNo, "Aviso") = vbNo Then
           Cancel = 1
           Exit Sub
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'fbAceptar = True
End Sub
Private Sub OptMoneda_Click(Index As Integer)
    Dim lsColor As String
    On Error GoTo ErrMoneda
    optDestino.iTem(0).value = True
    optDestino_Click (0)
    If Index = 0 Then
        fnMoneda = gMonedaNacional
        lsColor = &H80000005
    Else
        fnMoneda = gMonedaExtranjera
        lsColor = &HC0FFC0
    End If
    txtMonto.BackColor = lsColor
    Exit Sub
ErrMoneda:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub optMoneda_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If fraOrigen.Visible And fraOrigen.Enabled Then
            EnfocaControl txtAreaAgeCod
        End If
    End If
End Sub
Private Sub optTransporte_Click(Index As Integer)
    cboTerceros.Locked = True
    If Index = 0 Then
        cboTerceros.ListIndex = -1
    Else
        cboTerceros.Locked = False
    End If
End Sub
Private Sub optTransporte_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If optTransporte.iTem(0).value Then
            EnfocaControl txtMonto
        ElseIf optTransporte.iTem(1).value Then
            EnfocaControl cboTerceros
        End If
    End If
End Sub
Private Sub txtAreaAgeCod_EmiteDatos()
    lblAreaAgeDesc.Caption = ""
    If txtAreaAgeCod.Text <> "" Then
        lblAreaAgeDesc.Caption = txtAreaAgeCod.psDescripcion
    End If
End Sub
Private Sub optDestino_Click(Index As Integer)
    On Error GoTo ErrDestino
    Dim oOpe As New clases.DOperacion
    
    txtDestinoCod.Text = ""
    txtDestinoCod_EmiteDatos
    cboDocumento.ListIndex = -1
    txtDocumentoNro.Text = ""
    If Index = 0 Then
        lblDestino.Caption = "Agencia :"
        txtDestinoCod.psRaiz = "Agencias"
        lblDestinoDesc2.Visible = False
        txtDestinoCod.rs = rsAreaAgencia
        fraDocumento.Enabled = False
    ElseIf Index = 1 Then
        lblDestino.Caption = "Instituciones Financieras :"
        txtDestinoCod.psRaiz = "Instituciones Financieras"
        lblDestinoDesc2.Visible = True
        txtDestinoCod.rs = oOpe.listarCuentasEntidadesFinacieras("_1_[12]" & CStr(fnMoneda) & "%", CStr(fnMoneda))
        fraDocumento.Enabled = True
    End If
    Set oOpe = Nothing
    Exit Sub
ErrDestino:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub optDestino_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtDestinoCod
    End If
End Sub
Private Sub txtAreaAgeCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If optDestino.iTem(0).value Then
            EnfocaControl optDestino.iTem(0)
        ElseIf optDestino.iTem(1).value Then
            EnfocaControl optDestino.iTem(1)
        End If
    End If
End Sub
Private Sub txtDestinoCod_EmiteDatos()
    Dim oCtaIf As New NCajaCtaIF
    lblDestinoDesc1.Caption = ""
    lblDestinoDesc2.Caption = ""
    If txtDestinoCod.Text <> "" Then
        If optDestino.iTem(0).value Then
            lblDestinoDesc1.Caption = txtDestinoCod.psDescripcion
        ElseIf optDestino.iTem(1).value Then
            lblDestinoDesc1.Caption = oCtaIf.NombreIF(Mid(txtDestinoCod.Text, 4, 13))
            lblDestinoDesc2.Caption = oCtaIf.EmiteTipoCuentaIF(Mid(txtDestinoCod.Text, 18, Len(txtDestinoCod.Text))) & " " & txtDestinoCod.psDescripcion
        End If
    End If
    Set oCtaIf = Nothing
End Sub
Private Sub txtDestinoCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If fraDocumento.Visible And fraDocumento.Enabled Then
            EnfocaControl cboDocumento
        End If
    End If
End Sub
Private Sub txtDocumentoNro_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtGlosa
    End If
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If optTransporte.iTem(0).value Then
            EnfocaControl optTransporte.iTem(0)
        ElseIf optTransporte.iTem(1).value Then
            EnfocaControl optTransporte.iTem(1)
        End If
    End If
End Sub
Private Sub cboDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtDocumentoNro
    End If
End Sub
Private Sub cmdAceptar_Click()
    Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
    Dim oNContFunciones As clases.NContFunciones
    Dim oContImp As COMNContabilidad.NCOMContImprimir
    Dim rsBill As ADODB.Recordset
    Dim rsMon As ADODB.Recordset
    
    Dim lsMovNro As String, lsMovNroAprob As String
    Dim lsCadImp As String
    Dim nFicSal As Integer
    Dim lbOk As Boolean
    
    Dim ldFecha As Date
    Dim lnMonto As Currency, lnSaldoDisp As Currency
    Dim lsCtaCont As String, lsSubCtaCont As String, lsCtaContD As String, lsCtaContH As String
    Dim lsPersCodIF As String, lsCtaIFCod As String, lsIFTpo As String
    Dim lsAreaCodOrig As String, lsAgeCodOrig As String, lsAreaCodDest As String, lsAgeCodDest As String
    Dim lnTpoDoc As TpoDoc, lsNroDoc As String
    Dim lbPropioTransp As Boolean, lsPersCodTransp As String
    Dim lsGlosa As String
    Dim lbAgenciaDest As Boolean
    Dim lnIDSolicitud As Long
    Dim lbExito As Boolean
    Dim lnAprobacion As Integer
    Dim lsCtaContTemp As String
    
    If Not ValidaInterfaz Then Exit Sub
    
    ldFecha = CDate(txtFecha.Text)
    lnMonto = CCur(txtMonto.Text)
    lsAreaCodOrig = Left(txtAreaAgeCod.Text, 3)
    lsAgeCodOrig = Mid(txtAreaAgeCod.Text, 4, 2)
    lnTpoDoc = val(Trim(Right(cboDocumento.Text, 5)))
    lsNroDoc = Trim(txtDocumentoNro.Text)
    lbPropioTransp = IIf(optTransporte.iTem(0).value, True, False)
    lsPersCodTransp = Trim(Right(cboTerceros.Text, 15))
    lsGlosa = Trim(txtGlosa.Text)
    
    Set oNContFunciones = New clases.NContFunciones
    If optDestino.iTem(0).value Then
        lbAgenciaDest = True
        lsAreaCodDest = Left(txtDestinoCod.Text, 3)
        lsAgeCodDest = Mid(txtDestinoCod.Text, 4, 2)
    ElseIf optDestino.iTem(1).value Then
        lbAgenciaDest = False
        lsIFTpo = Left(txtDestinoCod.Text, 2)
        lsPersCodIF = Mid(txtDestinoCod.Text, 4, 13)
        lsCtaIFCod = Mid(txtDestinoCod.Text, 18, Len(txtDestinoCod.Text))
        lsCtaCont = "11" & fnMoneda & IIf(lsPersCodIF = "1090100822183", "2", "3") & "01"
        lsSubCtaCont = oNContFunciones.GetFiltroObjetos(1, lsCtaCont, txtDestinoCod.Text, False)
    End If
    
    Set oNContFunciones = New NContFunciones
    If lbAgenciaDest Then
        lsCtaContD = ReemplazaCaracterCtaCont(fsCtaContCodEfecTrans, fnMoneda, lsAgeCodDest)
        lsCtaContH = ReemplazaCaracterCtaCont(fsCtaContCodEfecTrans, fnMoneda, lsAgeCodOrig)
    Else
        lsCtaContD = lsCtaCont & lsSubCtaCont
        lsCtaContH = ReemplazaCaracterCtaCont(fsCtaContCodEfecTrans, fnMoneda, lsAgeCodOrig)
    End If
    'Verifica que las cuenta tengan puentes
    lsCtaContTemp = oNContFunciones.BuscaCtaEquivalente(lsCtaContD)
    If Len(lsCtaContTemp) > 0 Then lsCtaContD = lsCtaContTemp
    lsCtaContTemp = oNContFunciones.BuscaCtaEquivalente(lsCtaContH)
    If Len(lsCtaContTemp) > 0 Then lsCtaContH = lsCtaContTemp
        
    If Not oNContFunciones.verificarUltimoNivelCta(lsCtaContD) Then
        MsgBox "La cuenta contable " & lsCtaContD & " no es de Ultimo Nivel, comunicarse con el Dpto. de Contabilidad", vbInformation, "Aviso"
        EnfocaControl txtDestinoCod
        Set oNContFunciones = Nothing
        Exit Sub
    End If
    If Not oNContFunciones.verificarUltimoNivelCta(lsCtaContH) Then
        MsgBox "La cuenta contable " & lsCtaContH & " no es de Ultimo Nivel, comunicarse con el Dpto. de Contabilidad", vbInformation, "Aviso"
        EnfocaControl txtDestinoCod
        Set oNContFunciones = Nothing
        Exit Sub
    End If
    
    'lnSaldoDisp = oNContFunciones.GetSaldoCtaCont(gsFormatoMovFecha, ldFecha, lsCtaContH, fnMoneda)
    'If lnMonto > lnSaldoDisp Then
    '    MsgBox "No posee saldo suficiente para realizar la operación", vbInformation, "Aviso"
    '    Set oNContFunciones = Nothing
    '    Exit Sub
    'End If
    
    Set rsBill = New ADODB.Recordset
    Set rsMon = New ADODB.Recordset
    
    frmCajaGenEfectivo.Inicio gsOpeCod, gsOpeDesc, lnMonto, fnMoneda, False
    If frmCajaGenEfectivo.lbOk Then
        Set rsBill = frmCajaGenEfectivo.rsBilletes
        Set rsMon = frmCajaGenEfectivo.rsMonedas
    Else
        MsgBox "Ud. debe registrar correctamente la descomposición de efectivo", vbInformation, "Aviso"
        Set frmCajaGenEfectivo = Nothing
        Exit Sub
    End If
    Set frmCajaGenEfectivo = Nothing
    If (rsBill Is Nothing And rsMon Is Nothing) Then
        MsgBox "Error en ingreso de Descomposición de Efectivo, no se puede continuar..", vbInformation, "Aviso"
        RSClose rsBill: RSClose rsMon
        Exit Sub
    End If
    
    If MsgBox("¿Está seguro de realizar la operación?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
    
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    If lbAgenciaDest Then
        lsMovNroAprob = oNContFunciones.GeneraMovNro(ldFecha, Right(gsCodAge, 2), gsCodUser)
        lnIDSolicitud = oCaja.InsertaSolicitudAprobacionHabRemesa(lsMovNroAprob, lsAreaCodOrig, lsAgeCodOrig, lsAreaCodDest, lsAgeCodDest, fnMoneda, lnMonto)
        lnAprobacion = -1
        Do While lnAprobacion = -1
            MsgBox "La solicitud de Habilitación de remesa ya fue registrada." & Chr(13) & "Comuniquese con su Administrador para la aprobación/rechazo de la misma." & Chr(13), vbInformation, "Aviso"
            lnAprobacion = oCaja.ConsultaSolicitudAprobacionHabRemesa(lnIDSolicitud)
            If lnAprobacion = 0 Then
                MsgBox "Su solicitud de Habilitación de remesa fue rechazado, no puede continuar..", vbExclamation, "Aviso"
                Set oCaja = Nothing
                Exit Sub
            ElseIf lnAprobacion = 1 Then
                MsgBox "Su solicitud de Habilitación de remesa fue aprobado", vbInformation, "Aviso"
            End If
        Loop
    End If
    
    Screen.MousePointer = 11
    cmdAceptar.Enabled = False
    
    lsMovNro = oNContFunciones.GeneraMovNro(ldFecha, Right(gsCodAge, 2), gsCodUser)
    lbExito = oCaja.GrabaHabRemesaoDevolucion(lsMovNro, gsOpeCod, lsGlosa, fnMoneda, lsAreaCodOrig, lsAgeCodOrig, _
                                            lbAgenciaDest, lsAreaCodDest, lsAgeCodDest, lsIFTpo, lsPersCodIF, lsCtaIFCod, _
                                            lnTpoDoc, lsNroDoc, ldFecha, lbPropioTransp, lsPersCodTransp, lnMonto, rsBill, rsMon, lsCtaContD, lsCtaContH, lnIDSolicitud)
    Screen.MousePointer = 0
    If lbExito Then
        MsgBox "Se ha realizado la operación satisfactoriamente", vbInformation, "Aviso"
        ImprimeAsientoContable lsMovNro, , , , False, False, , , , , , True, 1
        If MsgBox("¿Desea registrar otra operación?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            Limpiar
        Else
            fbAceptar = True
            Unload Me
            Exit Sub
        End If
    Else
        MsgBox "Ha sucedido un error al realizar la operación, si el error persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    
    RSClose rsBill: RSClose rsMon
    Set oNContFunciones = Nothing
    Set oCaja = Nothing
    cmdAceptar.Enabled = True
    Exit Sub
ErrAceptar:
    Screen.MousePointer = 0
    cmdAceptar.Enabled = True
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Function ValidaInterfaz() As Boolean
    Dim lsValFecha As String
    ValidaInterfaz = True
    
    lsValFecha = ValidaFecha(txtFecha.Text)
    If Len(lsValFecha) > 0 Then
        MsgBox lsValFecha, vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl txtFecha
        Exit Function
    End If
    If optMoneda.iTem(0).value = False And optMoneda.iTem(1).value = False Then
        MsgBox "Ud. debe seleccionar la Moneda de la Operación", vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl optMoneda.iTem(0)
        Exit Function
    End If
    If Len(Trim(txtAreaAgeCod.Text)) <> 5 Then
        MsgBox "Ud. debe seleccionar el Área-Agencia Origen", vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl txtAreaAgeCod
        Exit Function
    End If
    If optDestino.iTem(0).value = False And optDestino.iTem(1).value = False Then
        MsgBox "Ud. debe seleccionar el Destino de la remesa", vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl optDestino.iTem(0)
        Exit Function
    Else
        If optDestino.iTem(0).value Then
            If Len(txtDestinoCod.Text) <> 5 Then
                MsgBox "Ud. debe seleccionar la Agencia Destino para la habilitación", vbInformation, "Aviso"
                ValidaInterfaz = False
                EnfocaControl txtDestinoCod
                Exit Function
            Else
                If txtDestinoCod.Text = txtAreaAgeCod.Text Then
                    MsgBox "Ud. debe seleccionar la Agencia Destino diferente a la Agencia Origen", vbInformation, "Aviso"
                    ValidaInterfaz = False
                    EnfocaControl txtDestinoCod
                    Exit Function
                End If
            End If
        ElseIf optDestino.iTem(1).value Then
            If Len(txtDestinoCod.Text) = 0 Then
                MsgBox "Ud. debe seleccionar la cuenta de la Institución Financiera para la Devolución", vbInformation, "Aviso"
                ValidaInterfaz = False
                EnfocaControl txtDestinoCod
                Exit Function
            End If
            If Len(Trim(cboDocumento.Text)) = 0 Then
                MsgBox "Ud. debe seleccionar el Documento de la remesa", vbInformation, "Aviso"
                ValidaInterfaz = False
                EnfocaControl cboDocumento
                Exit Function
            End If
            If Len(Trim(txtDocumentoNro.Text)) = 0 Then
                MsgBox "Ud. debe especificar el Nro. de Documento de la remesa", vbInformation, "Aviso"
                ValidaInterfaz = False
                EnfocaControl txtDocumentoNro
                Exit Function
            End If
        End If
    End If
    If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Glosa", vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl txtGlosa
        Exit Function
    End If
    If optTransporte.iTem(0).value = False And optTransporte.iTem(1).value = False Then
        MsgBox "Ud. debe seleccionar el Modo de Transporte", vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl optTransporte.iTem(0)
        Exit Function
    Else
        If optTransporte.iTem(1).value And cboTerceros.ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar la empresa de Transporte respectiva", vbInformation, "Aviso"
            ValidaInterfaz = False
            EnfocaControl cboTerceros
            Exit Function
        End If
    End If
    If Not IsNumeric(txtMonto.Text) Then
        MsgBox "Ud. debe ingresar un monto mayor a cero", vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl txtMonto
        Exit Function
    Else
        If CCur(txtMonto.Text) <= 0 Then
            MsgBox "Ud. debe ingresar un monto mayor a cero", vbInformation, "Aviso"
            ValidaInterfaz = False
            EnfocaControl txtMonto
            Exit Function
        End If
    End If
End Function
Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 20, 2)
    If KeyAscii = 13 Then
        EnfocaControl cmdAceptar
    End If
End Sub
Private Sub txtMonto_LostFocus()
    If Len(Trim(txtMonto.Text)) = 0 Then txtMonto.Text = "0"
    txtMonto.Text = Format(txtMonto.Text, gsFormatoNumeroView)
End Sub
Private Function ReemplazaCaracterCtaCont(ByVal psCtaContCod As String, ByVal pnMoneda As Moneda, ByVal psAgeCod As String) As String
    ReemplazaCaracterCtaCont = psCtaContCod
    ReemplazaCaracterCtaCont = Replace(ReemplazaCaracterCtaCont, "M", pnMoneda)
    ReemplazaCaracterCtaCont = Replace(ReemplazaCaracterCtaCont, "AG", Format(psAgeCod, "00"))
End Function
