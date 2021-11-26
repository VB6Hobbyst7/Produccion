VERSION 5.00
Begin VB.Form frmPigPagoSobrante 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pago de Sobrante"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   Icon            =   "frmPigPagoSobrante.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraContenedor 
      Height          =   4155
      Index           =   0
      Left            =   30
      TabIndex        =   3
      Top             =   -45
      Width           =   7275
      Begin VB.CommandButton cmdBuscar 
         Height          =   390
         Left            =   6540
         Picture         =   "frmPigPagoSobrante.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Buscar ..."
         Top             =   465
         Width           =   420
      End
      Begin VB.Frame fraDatos 
         Enabled         =   0   'False
         Height          =   2880
         Left            =   75
         TabIndex        =   4
         Top             =   1200
         Width           =   7125
         Begin SICMACT.EditMoney txtMonto 
            Height          =   270
            Left            =   5685
            TabIndex        =   9
            Top             =   2445
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   476
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   12648447
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.FlexEdit feCte 
            Height          =   765
            Left            =   90
            TabIndex        =   5
            Top             =   210
            Width           =   6930
            _ExtentX        =   12224
            _ExtentY        =   1349
            Cols0           =   3
            FixedCols       =   0
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "Codigo-Nombre-Doc.Iden"
            EncabezadosAnchos=   "1300-3800-1500"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X"
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C"
            FormatosEdit    =   "0-0-0"
            TextArray0      =   "Codigo"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   1305
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feContratos 
            Height          =   1695
            Left            =   75
            TabIndex        =   15
            Top             =   1050
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   2990
            Cols0           =   4
            FixedCols       =   0
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Contrato-Remate-Sobrante"
            EncabezadosAnchos=   "300-2000-650-1600"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X"
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-R"
            FormatosEdit    =   "0-0-3-2"
            TextArray0      =   "#"
            lbFlexDuplicados=   0   'False
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label Label3 
            Caption         =   "PAGO DE SOBRANTE"
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
            Height          =   210
            Left            =   4995
            TabIndex        =   16
            Top             =   1200
            Width           =   1980
         End
         Begin VB.Label lblSobrante 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   5700
            TabIndex        =   14
            Top             =   1665
            Width           =   1260
         End
         Begin VB.Label Label2 
            Caption         =   "Sobrante"
            Height          =   210
            Left            =   4875
            TabIndex        =   13
            Top             =   1680
            Width           =   780
         End
         Begin VB.Label lblInteres 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   5700
            TabIndex        =   12
            Top             =   2040
            Width           =   1260
         End
         Begin VB.Label Label1 
            Caption         =   "Interés"
            Height          =   210
            Left            =   4905
            TabIndex        =   11
            Top             =   2070
            Width           =   780
         End
         Begin VB.Label Label12 
            Caption         =   "Total"
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
            Height          =   195
            Left            =   4890
            TabIndex        =   8
            Top             =   2490
            Width           =   615
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   450
         Left            =   165
         TabIndex        =   7
         Top             =   270
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   794
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin SICMACT.ActXCodCta ActXCodCtaAhorro 
         Height          =   450
         Left            =   150
         TabIndex        =   10
         Top             =   735
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   794
         Texto           =   "Cta Ahorro"
         EnabledCta      =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4035
      TabIndex        =   2
      Top             =   4215
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   6195
      TabIndex        =   1
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5130
      TabIndex        =   0
      Top             =   4215
      Width           =   975
   End
End
Attribute VB_Name = "frmPigPagoSobrante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim pnInteres As Currency
'Dim pnSaldo As Currency
'Dim lnRemate As Long
'Dim lsPersCod As String
'
'Private Sub ActXCodCtaAhorro_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then BuscaContrato (ActXCodCtaAhorro.NroCuenta)
'End Sub
'
'Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
'End Sub
'
'Private Sub cmdBuscar_Click()
'
'Dim loPers As UPersona
'Dim lsPersNombre As String
'Dim lsEstados As String
'Dim loPersContrato As DColPContrato
'Dim lrContratos As ADODB.Recordset
'Dim loCuentas As UProdPersona
'
'On Error GoTo ControlError
'
'Set loPers = New UPersona
'    Set loPers = frmBuscaPersona.Inicio
'    If loPers Is Nothing Then Exit Sub
'    lsPersCod = loPers.sPersCod
'    lsPersNombre = loPers.sPersNombre
'Set loPers = Nothing
'
''Selecciona Estados
'lsEstados = gPigEstRematFact & "," & gPigEstAdjud
'
'If Trim(lsPersCod) <> "" Then
'    Set loPersContrato = New DColPContrato
'        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
'    Set loPersContrato = Nothing
'End If
'
'    Set loCuentas = New UProdPersona
'    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos)
'    If loCuentas.sCtaCod <> "" Then
'        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
'        AXCodCta.SetFocusCuenta
'    End If
'    Set loCuentas = Nothing
'    cmdGrabar.Enabled = True
'
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'
'End Sub
'
'Private Sub cmdCancelar_Click()
'    Limpia
'End Sub
'
'Private Sub cmdGrabar_Click()
'Dim oPigGraba As NPigRemate
'Dim oImpre As NPigImpre
'Dim oCont As NContFunciones
'Dim lsFechaHoraGrab As String
'Dim lsMovNro As String
'Dim rs As Recordset
'Dim lsCtaAho As String
'
'
'    Set oPigGraba = New NPigRemate
'
'    Set oCont = New NContFunciones
'    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'    Set oCont = Nothing
'    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
'
'    Set rs = Me.feContratos.GetRsNew
'    lsCtaAho = ActXCodCtaAhorro.NroCuenta
'    oPigGraba.nPigPagoSobrante rs, lsCtaAho, lsMovNro, CCur(txtMonto), gdFecSis, gsCodAge, gsCodUser
'    Set oPigGraba = Nothing
'
'    Set oImpre = New NPigImpre
'    Call oImpre.ImpreReciboCancelacionSobrante(gsInstCmac, gsNomAge, lsFechaHoraGrab, rs, '                ActXCodCtaAhorro.NroCuenta, feCte.TextMatrix(1, 1), txtMonto, pnSaldo, pnInteres, gsCodUser, sLpt, "")
'
'    Do While MsgBox("Desea Reimprimir Comprobante de Pago de Sobrante? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes
'        Call oImpre.ImpreReciboCancelacionSobrante(gsInstCmac, gsNomAge, lsFechaHoraGrab, rs, '                    ActXCodCtaAhorro.NroCuenta, feCte.TextMatrix(1, 1), txtMonto, pnSaldo, pnInteres, gsCodUser, sLpt, "")
'
'    Loop
'
'    Set oImpre = Nothing
'    Limpia
'
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub BuscaContrato(ByVal psNroContrato As String)
'Dim rs As ADODB.Recordset
'Dim oValContrato As nPigValida
'Dim oPigContrato As DPigContrato
'Dim oPigRemate As DPigRemate
'Dim lsCtaAho As String
'Dim lnSobrante As Currency
'
'On Error GoTo ControlError
'
'    'Valida Contrato
'    lnSobrante = 0
'    Set rs = New ADODB.Recordset
'    Set oValContrato = New nPigValida
'        Set rs = oValContrato.ValidaCancelacionSobrante(psNroContrato, lsPersCod)
'    Set oValContrato = Nothing
'
'    If (rs.EOF And rs.BOF) Then
'        MsgBox "Contrato no posee Sobrante de Remate Pendiente de Cancelación"
'        Limpia
'        Exit Sub
'    Else
'        lsCtaAho = rs!cCtaAbono
'        '== Muestro los datos del contrato
'        If feCte.TextMatrix(1, 0) = "" Then
'            feCte.TextMatrix(1, 0) = rs!cPersCod
'            feCte.TextMatrix(1, 1) = PstaNombre(rs!cPersNombre)
'            feCte.TextMatrix(1, 2) = IIf(IsNull(rs!NroDNI), " ", rs!NroDNI)
'        End If
'        Do While Not rs.EOF
'            feContratos.AdicionaFila
'            feContratos.TextMatrix(feContratos.Rows - 1, 1) = rs!cCtaCod
'            feContratos.TextMatrix(feContratos.Rows - 1, 2) = rs!nRemate
'            feContratos.TextMatrix(feContratos.Rows - 1, 3) = Format(rs!nSobrante, "#####,###.00")
'            lnSobrante = lnSobrante + CCur(rs!nSobrante)
'            rs.MoveNext
'        Loop
'    End If
'
'    'Cargar los datos de la Cta para su cancelacion
'    Set oPigRemate = New DPigRemate
'
'    txtMonto = oPigRemate.GetSaldoCancelacion(lsCtaAho, gdFecSis, gsCodAge, pnInteres, pnSaldo)
'    lblSobrante = Format(lnSobrante, "#####,###.0000")
'    pnSaldo = lnSobrante
'    lblInteres = Format(CCur(txtMonto) - lnSobrante, "#####,###.0000")
'    pnInteres = lblInteres
'
'    ActXCodCtaAhorro.NroCuenta = lsCtaAho
'
'    Set oPigRemate = Nothing
'
'    ActXCodCtaAhorro.Enabled = False
'    AXCodCta.Enabled = False
'    cmdBuscar.Enabled = False
'    cmdGrabar.Enabled = True
'    cmdGrabar.SetFocus
'
'Exit Sub
'
'ControlError:
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
'        Dim sCuenta As String
'        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
'        If sCuenta <> "" Then
'            AXCodCta.NroCuenta = sCuenta
'            AXCodCta.SetFocusCuenta
'        End If
'    End If
'End Sub
'
'Private Sub Form_Load()
'
'   AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
'   AXCodCta.Age = ""
'   ActXCodCtaAhorro.CMAC = AXCodCta.CMAC
'   ActXCodCtaAhorro.Prod = "232"
'
'End Sub
'
'Private Sub Limpia()
'
'    cmdBuscar.Enabled = True
'    cmdGrabar.Enabled = False
'    ActXCodCtaAhorro.Age = ""
'    ActXCodCtaAhorro.Cuenta = ""
'    feCte.Clear
'    feCte.Rows = 2
'    feCte.FormaCabecera
'    lblInteres = ""
'    lblSobrante = ""
'    txtMonto = ""
'    AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
'    AXCodCta.Age = ""
'    AXCodCta.Cuenta = ""
'    ActXCodCtaAhorro.Enabled = True
'    AXCodCta.Enabled = True
'    lsPersCod = ""
'    feContratos.Clear
'    feContratos.Rows = 2
'    feContratos.FormaCabecera
'
'End Sub
