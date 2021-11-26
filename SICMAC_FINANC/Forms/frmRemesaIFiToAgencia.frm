VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRemesaIFiToAgencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REMESA DE BANCO A AGENCIA"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   Icon            =   "frmRemesaIFiToAgencia.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   350
      Left            =   6390
      TabIndex        =   11
      Top             =   5055
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   350
      Left            =   7500
      TabIndex        =   12
      Top             =   5055
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Height          =   5025
      Left            =   80
      TabIndex        =   13
      Top             =   0
      Width           =   8535
      Begin VB.Frame fraFecha 
         BorderStyle     =   0  'None
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
         Height          =   375
         Left            =   6470
         TabIndex        =   26
         Top             =   180
         Width           =   1935
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   330
            Left            =   720
            TabIndex        =   0
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha :"
            Height          =   195
            Left            =   0
            TabIndex        =   27
            Top             =   45
            Width           =   540
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
         Left            =   6300
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   4575
         Width           =   2070
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
         Left            =   4440
         TabIndex        =   22
         Top             =   3360
         Width           =   3975
         Begin VB.ComboBox cboTerceros 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   600
            Width           =   2520
         End
         Begin VB.OptionButton optTransporte 
            Caption         =   "Terceros"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton optTransporte 
            Caption         =   "Propio"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   315
            Value           =   -1  'True
            Width           =   855
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
         TabIndex        =   21
         Top             =   3360
         Width           =   3495
         Begin VB.TextBox txtGlosa 
            Height          =   660
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   240
            Width           =   3240
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
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   8295
         Begin Sicmact.TxtBuscar txtAreaAgeCod 
            Height          =   300
            Left            =   1515
            TabIndex        =   5
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
         Begin VB.Label lblAreaAgeDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3120
            TabIndex        =   25
            Top             =   240
            Width           =   5055
         End
         Begin VB.Label Label6 
            Caption         =   "Area - Agencia :"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.Frame fraOrigen 
         Caption         =   "Origen"
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
         Height          =   1935
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   8295
         Begin VB.Frame Frame1 
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
            ForeColor       =   &H80000006&
            Height          =   855
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   8055
            Begin VB.TextBox txtDocumentoNro 
               Height          =   300
               Left            =   2040
               TabIndex        =   4
               Top             =   360
               Width           =   2280
            End
            Begin VB.OptionButton optDocumento 
               Caption         =   "Cheque"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   3
               Top             =   480
               Width           =   1215
            End
            Begin VB.OptionButton optDocumento 
               Caption         =   "Carta"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   2
               Top             =   225
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.Label lblNro 
               AutoSize        =   -1  'True
               Caption         =   "N° :"
               Height          =   195
               Left            =   1560
               TabIndex        =   17
               Top             =   360
               Width           =   270
            End
         End
         Begin Sicmact.TxtBuscar txtIFiCtaCod 
            Height          =   300
            Left            =   840
            TabIndex        =   1
            Top             =   240
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
         Begin VB.Label lblCtaIFDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   840
            TabIndex        =   24
            Top             =   600
            Width           =   7335
         End
         Begin VB.Label lblIFNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3480
            TabIndex        =   23
            Top             =   240
            Width           =   4695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta :"
            Height          =   195
            Left            =   140
            TabIndex        =   18
            Top             =   260
            Width           =   600
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
         TabIndex        =   14
         Top             =   4620
         Width           =   660
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   5355
         Top             =   4560
         Width           =   3045
      End
   End
End
Attribute VB_Name = "frmRemesaIFiToAgencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************************
'** Nombre : frmRemesaIFiToAgencia
'** Descripción : Formulario para el remesas de Instituciones Financieras a nuestras Agencias
'** Creación : EJVG, 20140626 15:30:00 PM
'********************************************************************************************
Dim fsopecod As String
Dim fsCtaContCodD As String
Dim objPista As COMManejador.Pista 'ARLO20170217

Public Sub Inicia(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    Dim lsColor As String
    fsopecod = psOpeCod
    fsCtaContCodD = ""
    glAceptar = False
    Caption = psOpeDesc
    If Mid(fsopecod, 3, 1) = Moneda.gMonedaNacional Then
        lsColor = &H80000005
    Else
        lsColor = &HC0FFC0
    End If
    txtIFiCtaCod.BackColor = lsColor
    txtMonto.BackColor = lsColor
    Show 1
End Sub
Private Sub Form_Load()
    CargarControles
    Limpiar
End Sub
Private Sub cmdSalir_Click()
    glAceptar = False
    Unload Me
End Sub
Private Sub CargarControles()
    Dim oOpe As New DOperacion
    Dim oPer As New DPersonas
    Dim rsPer As New ADODB.Recordset
    
    On Error GoTo ErrCargarControles
    Screen.MousePointer = 11
    'Origen: Cuentas Instituciones
    txtIFiCtaCod.psRaiz = "Cuentas de Instituciones Financieras"
    txtIFiCtaCod.rs = oOpe.GetOpeObj(fsopecod, "1")
    fsCtaContCodD = oOpe.EmiteOpeCta(fsopecod, "D", "0")
    'Destino: Area Agencia
    txtAreaAgeCod.psRaiz = "Área-Agencia"
    txtAreaAgeCod.rs = GetObjetosOpeCta(fsopecod, "2", fsCtaContCodD, "")
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
    Set rsPer = Nothing
    Set oPer = Nothing
    Set oOpe = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrCargarControles:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not glAceptar Then
        If MsgBox("¿Está seguro de Salir sin grabar la Operación?", vbInformation + vbYesNo, "Aviso") = vbNo Then
           Cancel = 1
           Exit Sub
        End If
    End If
End Sub
Private Sub optDocumento_Click(Index As Integer)
    txtDocumentoNro.Text = ""
    If Index = 0 Then
        lblNro.Caption = "N° :"
        lblNro.Width = 270
        txtDocumentoNro.Left = 2040
    ElseIf Index = 1 Then
        lblNro.Caption = "N° de Planilla:"
        lblNro.Width = 990
        txtDocumentoNro.Left = 2640
    End If
End Sub
Private Sub optDocumento_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtDocumentoNro
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
        If Index = 0 Then
            EnfocaControl txtMonto
        ElseIf Index = 1 Then
            EnfocaControl cboTerceros
        End If
    End If
End Sub
Private Sub txtAreaAgeCod_EmiteDatos()
    lblAreaAgeDesc.Caption = ""
    If txtAreaAgeCod.Text <> "" Then
        lblAreaAgeDesc.Caption = txtAreaAgeCod.psDescripcion
        EnfocaControl txtAreaAgeCod
    End If
End Sub
Private Sub txtAreaAgeCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtGlosa
    End If
End Sub
Private Sub txtDocumentoNro_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtAreaAgeCod
    End If
End Sub
Private Sub txtDocumentoNro_LostFocus()
    txtDocumentoNro.Text = Trim(txtDocumentoNro.Text)
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If optTransporte.Item(0).value Then
            EnfocaControl optTransporte.Item(0)
        ElseIf optTransporte.Item(1).value Then
            EnfocaControl optTransporte.Item(1)
        End If
    End If
End Sub
Private Sub txtIFiCtaCod_EmiteDatos()
    Dim oCtaIf As New NCajaCtaIF
    lblIFNombre.Caption = ""
    lblCtaIFDesc.Caption = ""
    If txtIFiCtaCod.Text <> "" Then
        lblIFNombre.Caption = oCtaIf.NombreIF(Mid(txtIFiCtaCod.Text, 4, 13))
        lblCtaIFDesc = oCtaIf.EmiteTipoCuentaIF(Mid(txtIFiCtaCod.Text, 18, Len(txtIFiCtaCod.Text))) & " " & txtIFiCtaCod.psDescripcion
    End If
    Set oCtaIf = Nothing
End Sub
Private Sub txtIFiCtaCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If optDocumento.Item(0).value Then
            EnfocaControl optDocumento.Item(0)
        ElseIf optDocumento.Item(1).value Then
            EnfocaControl optDocumento.Item(1)
        End If
    End If
End Sub
Private Sub cboTerceros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtMonto
    End If
End Sub
Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 20, 2)
    If KeyAscii = 13 Then
        EnfocaControl cmdAceptar
    End If
End Sub
Private Sub cmdAceptar_Click()
    Dim oOpe As DOperacion
    Dim oCaja As nCajaGeneral
    Dim oFun As NContFunciones
    Dim oImp As NContImprimir
    Dim oNIFi As NCajaCtaIF
    Dim bExito As Boolean
    Dim lnTpoDoc As TpoDoc
    Dim lsNroDoc As String
    Dim lsCtaContCodD As String
    Dim lsCtaContCodH As String
    Dim lsSubCtaContCodH As String
    Dim lsIFTpo As String
    Dim lsPersCodIf As String
    Dim lsCtaIFCod As String
    Dim lbPropioTransp As Boolean
    Dim lsPersCodTransp As String
    Dim lsMovNro As String
    Dim lnIFiSaldo As Currency
    Dim lnImporte As Currency
    Dim lsAreaAgeCodOrig As String

    On Error GoTo ErrAceptar
    If Not ValidaInterfaz() Then Exit Sub
    Set oOpe = New DOperacion
    lsCtaContCodH = oOpe.EmiteOpeCta(fsopecod, "H", "0", txtIFiCtaCod.Text, ObjEntidadesFinancieras)
    If Not oOpe.ValidaCtaCont(lsCtaContCodH) Then
        MsgBox "Falta definir Cuenta Contable de Institución Financiera [HABER], no se puede continuar..", vbInformation, "Aviso"
        Set oOpe = Nothing
        Exit Sub
    End If
    Set oFun = New NContFunciones
    lsSubCtaContCodD = oFun.GetFiltroObjetos(ObjCMACAgenciaArea, fsCtaContCodD, txtAreaAgeCod.Text, False)
    lsCtaContCodD = fsCtaContCodD + lsSubCtaContCodD
    If Not oOpe.ValidaCtaCont(lsCtaContCodD) Then
        MsgBox "Falta definir Cuenta Contable [DEBE], no se puede continuar..", vbInformation, "Aviso"
        Set oOpe = Nothing
        Set oFun = Nothing
        Exit Sub
    End If
    Set oOpe = Nothing
    Set oFun = Nothing
    
    Set oNIFi = New NCajaCtaIF
    lsIFTpo = Left(txtIFiCtaCod, 2)
    lsPersCodIf = Mid(txtIFiCtaCod, 4, 13)
    lsCtaIFCod = Mid(txtIFiCtaCod, 18, Len(txtIFiCtaCod))
    lnIFiSaldo = oNIFi.GetSaldoCtaIf(lsPersCodIf, lsIFTpo, lsCtaIFCod, gdFecSis, lsCtaContCodH, IIf(Mid(lsCtaContCodH, 3, 1) = "1", gMonedaNacional, gMonedaExtranjera))
    lnImporte = CCur(txtMonto)
    lsAreaAgeCodOrig = Left(gsCodArea, 3) & Right(gsCodAge, 2)
    Set oNIFi = Nothing
    
    If lnIFiSaldo < lnImporte Then
        MsgBox "El saldo de la Institución Financiera no cubre la remesa a realizar, no se puede continuar..", vbExclamation, "Aviso"
        EnfocaControl txtMonto
        Exit Sub
    End If
    
    If MsgBox("¿Está seguro de realizar la remesa? ", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    cmdAceptar.Enabled = False
    
    If optDocumento(0).value Then
        lnTpoDoc = TpoDocCarta
    Else
        lnTpoDoc = TpoDocCheque
    End If
    lsNroDoc = Trim(txtDocumentoNro.Text)
    lbPropioTransp = IIf(optTransporte.Item(0).value, 1, 0)
    lsPersCodTransp = Right(cboTerceros.Text, 13)
    
    Set oCaja = New nCajaGeneral
    bExito = oCaja.GrabarRemesaIFiAAgencia(CDate(txtFecha), Right(gsCodAge, 2), gsCodUser, fsopecod, lsIFTpo, lsPersCodIf, lsCtaIFCod, lnTpoDoc, lsNroDoc, lsAreaAgeCodOrig, txtAreaAgeCod, txtGlosa, lnImporte, lbPropioTransp, lsPersCodTransp, lsCtaContCodD, lsCtaContCodH, lsMovNro)
    Set oCaja = Nothing
    Screen.MousePointer = 0
    
    If bExito Then
        MsgBox "Se ha registrado satisfactoriamente la remesa", vbInformation, "Aviso"
        Set oImp = New NContImprimir
        EnviaPrevio oImp.ImprimeAsientoContable(lsMovNro, gnLinPage, gnColPage, "REMESA DE INSTITUCIÓN FINANCIERA A AGENCIAS"), "REMESA DE INSTITUCIÓN FINANCIERA A AGENCIAS", gnLinPage, False
        Set oImp = Nothing
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Remesa de Banco a Agencia"
        Set objPista = Nothing
        '****
        If MsgBox("¿Desea registrar otra remesa? ", vbInformation + vbYesNo, "Aviso") = vbYes Then
            Limpiar
        Else
            glAceptar = True
            Unload Me
            Exit Sub
        End If
    Else
        MsgBox "Ha sucedido un error al realizar la remesa, si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    cmdAceptar.Enabled = True
    Exit Sub
ErrAceptar:
    Screen.MousePointer = 0
    cmdAceptar.Enabled = True
    MsgBox Err.Description, vbInformation, "Aviso"
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
    If Len(Trim(txtIFiCtaCod.Text)) = 0 Then
        MsgBox "Ud. debe seleccionar la Institución Financiera", vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl txtIFiCtaCod
        Exit Function
    End If
    If optDocumento.Item(0).value = False And optDocumento.Item(1).value = False Then
        MsgBox "Ud. debe seleccionar un Tipo de Documento", vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl optDocumento.Item(0)
        Exit Function
    End If
    If Len(Trim(txtDocumentoNro.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Nro. de Documento", vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl txtDocumentoNro
        Exit Function
    End If
    If Len(Trim(txtAreaAgeCod.Text)) = 0 Then
        MsgBox "Ud. debe seleccionar el Área-Agencia destino", vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl txtAreaAgeCod
        Exit Function
    End If
    If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Glosa", vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl txtGlosa
        Exit Function
    End If
    If optTransporte.Item(0).value = False And optTransporte.Item(1).value = False Then
        MsgBox "Ud. debe seleccionar el Modo de Transporte", vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl optTransporte.Item(0)
        Exit Function
    End If
    If optTransporte.Item(1).value And cboTerceros.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la empresa de Transporte respectiva", vbInformation, "Aviso"
        ValidaInterfaz = False
        EnfocaControl cboTerceros
        Exit Function
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
Private Sub Limpiar()
    txtFecha.Text = Format(gdFecSis, gsFormatoFechaView)
    txtIFiCtaCod.Text = ""
    txtIFiCtaCod_EmiteDatos
    optDocumento.Item(0).value = True
    optDocumento_Click (0)
    txtAreaAgeCod.Text = ""
    txtAreaAgeCod_EmiteDatos
    txtGlosa.Text = ""
    optTransporte.Item(0).value = True
    optTransporte_Click (0)
    txtMonto.Text = "0.00"
End Sub
Private Sub txtMonto_LostFocus()
    If Len(Trim(txtMonto.Text)) = 0 Then txtMonto.Text = "0"
    txtMonto.Text = Format(txtMonto.Text, gsFormatoNumeroView)
End Sub
