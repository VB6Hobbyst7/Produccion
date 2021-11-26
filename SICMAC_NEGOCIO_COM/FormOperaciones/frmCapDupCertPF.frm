VERSION 5.00
Begin VB.Form frmCapDupCertPF 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   Icon            =   "frmCapDupCertPF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ForeColor       =   &H00000080&
      Height          =   915
      Left            =   60
      TabIndex        =   28
      Top             =   4320
      Width           =   7935
      Begin VB.TextBox txtGlosa 
         Height          =   555
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6900
      TabIndex        =   4
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   5700
      TabIndex        =   3
      Top             =   5340
      Width           =   1100
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1770
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   7920
      Begin VB.Frame fraDatos 
         Height          =   945
         Left            =   105
         TabIndex        =   12
         Top             =   660
         Width           =   7680
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   135
            TabIndex        =   24
            Top             =   255
            Width           =   690
         End
         Begin VB.Label lblApertura 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   1140
            TabIndex        =   23
            Top             =   180
            Width           =   1905
         End
         Begin VB.Label lblEtqUltCnt 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (días) :"
            Height          =   195
            Left            =   3240
            TabIndex        =   22
            Top             =   255
            Width           =   930
         End
         Begin VB.Label lblPlazo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4320
            TabIndex        =   21
            Top             =   180
            Width           =   855
         End
         Begin VB.Label lblDuplicados 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4320
            TabIndex        =   20
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "# Duplicados :"
            Height          =   195
            Left            =   3180
            TabIndex        =   19
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label lblVencimiento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   1140
            TabIndex        =   18
            Top             =   540
            Width           =   1905
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento :"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   960
         End
         Begin VB.Label lblTasa 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   6180
            TabIndex        =   16
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tasa (%) :"
            Height          =   195
            Left            =   5400
            TabIndex        =   15
            Top             =   255
            Width           =   705
         End
         Begin VB.Label lblDias 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   6180
            TabIndex        =   14
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "# Días :"
            Height          =   195
            Left            =   5340
            TabIndex        =   13
            Top             =   600
            Width           =   585
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   3840
         TabIndex        =   25
         Top             =   240
         Width           =   3960
      End
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2475
      Left            =   60
      TabIndex        =   6
      Top             =   1785
      Width           =   7950
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1755
         Left            =   105
         TabIndex        =   1
         Top             =   225
         Width           =   7680
         _ExtentX        =   13811
         _ExtentY        =   3096
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre-Relacion"
         EncabezadosAnchos=   "250-1700-3800-1500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         EncabezadosAlineacion=   "C-C-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   255
         RowHeight0      =   300
         TipoBusPersona  =   1
      End
      Begin VB.Label lblFormaRetiro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   5505
         TabIndex        =   27
         Top             =   2085
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Forma Retiro :"
         Height          =   195
         Left            =   4425
         TabIndex        =   26
         Top             =   2145
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   2138
         Width           =   960
      End
      Begin VB.Label lblTipoCuenta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   1155
         TabIndex        =   9
         Top             =   2085
         Width           =   1560
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "# Firmas :"
         Height          =   195
         Left            =   2910
         TabIndex        =   8
         Top             =   2145
         Width           =   690
      End
      Begin VB.Label lblFirmas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   3735
         TabIndex        =   7
         Top             =   2085
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmCapDupCertPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nMoneda As Moneda
Dim nTipoCuenta As ProductoCuentaTipo
Dim nOperacion As CaptacOperacion
Dim bDocumento As Boolean
Dim nDocumento As TpoDoc
Dim nCapitalInicial As Double
Dim nTasa As Double

Dim ldRenovacion As Date

Private Function GetDireccionCliente() As String
Dim sDireccion As String
Dim i As Integer
Dim clsMant As NCapMantenimiento
Dim rsPers As Recordset
Set clsMant = New NCapMantenimiento
sDireccion = ""
For i = 1 To grdCliente.Rows - 1
    If CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4))) = gCapRelPersTitular Then
        Set rsPers = clsMant.GetDatosPersona(grdCliente.TextMatrix(i, 1))
        sDireccion = Trim(rsPers("Direccion"))
        Exit For
    End If
Next i
GetDireccionCliente = sDireccion
End Function

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As NCapMantenimiento
Dim clsCap As NCapMovimientos
Dim rsCta As Recordset, rsRel As Recordset
Dim nEstado As CaptacEstado
Dim nRow As Long
Dim sMsg As String, sMoneda As String, sPersona As String
Dim dUltRetInt As Date
Set clsCap = New NCapMovimientos
sMsg = clsCap.ValidaCuentaOperacion(sCuenta)
If sMsg = "" Then
    Set clsMant = New NCapMantenimiento
    Set rsCta = New Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        nEstado = rsCta("nPrdEstado")
        lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm")
        nMoneda = CLng(Mid(sCuenta, 9, 1))
        If nMoneda = gMonedaNacional Then
            lblMensaje = "DEPOSITO A PLAZO FIJO" & Chr$(13) & "MONEDA NACIONAL"
        ElseIf nMoneda = gMonedaExtranjera Then
            lblMensaje = "DEPOSITO A PLAZO FIJO" & Chr$(13) & "MONEDA EXTRANJERA"
        End If
        lblPlazo = Format$(rsCta("nPlazo"), "#,##0")
        lblTipoCuenta = UCase(rsCta("cTipoCuenta"))
        lblVencimiento = Format(DateAdd("d", rsCta("nPlazo"), rsCta("dRenovacion")), "dd mmm yyyy")
        ldRenovacion = rsCta("dRenovacion")
        nTasa = rsCta("nTasaInteres")
        lblTasa = Format$(ConvierteTNAaTEA(nTasa), "#0.00")
        nTipoCuenta = rsCta("nPrdCtaTpo")
        lblFirmas = Format$(rsCta("nFirmas"), "#0")
        dUltRetInt = clsCap.GetFechaUltimoRetiroIntPF(sCuenta)
        lblDias = Format$(DateDiff("d", dUltRetInt, gdFecSis), "#0")
        lblDuplicados = rsCta("nDuplicado")
        nCapitalInicial = rsCta("nApertura")
        lblFormaRetiro = rsCta("cRetiro")
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        sPersona = ""
        Do While Not rsRel.EOF
            If sPersona <> rsRel("cPersCod") Then
                grdCliente.AdicionaFila
                nRow = grdCliente.Rows - 1
                grdCliente.TextMatrix(nRow, 1) = rsRel("cPersCod")
                grdCliente.TextMatrix(nRow, 2) = UCase(PstaNombre(rsRel("Nombre")))
                grdCliente.TextMatrix(nRow, 3) = UCase(rsRel("Relacion")) & Space(100) & rsRel("nPrdPersRelac")
                sPersona = rsRel("cPersCod")
            End If
            rsRel.MoveNext
        Loop
        rsRel.Close
        Set rsRel = Nothing
        fraCliente.Enabled = True
        fraCuenta.Enabled = False
        fraGlosa.Enabled = True
        cmdImprimir.Enabled = True
        cmdCancelar.Enabled = True
    End If
Else
    MsgBox sMsg, vbInformation, "Operacion"
    txtCuenta.SetFocus
End If
End Sub

Private Sub LimpiaControles()
grdCliente.Clear
grdCliente.Rows = 2
grdCliente.FormaCabecera
cmdImprimir.Enabled = False
txtCuenta.Age = ""
txtCuenta.Cuenta = ""
cmdImprimir.Enabled = False
cmdCancelar.Enabled = False
lblApertura = ""
lblFirmas = ""
lblTipoCuenta = ""
lblPlazo = ""
lblDias = ""
lblVencimiento = ""
lblDuplicados = ""
lblTasa = ""
lblFormaRetiro = ""
fraCliente.Enabled = False
fraCuenta.Enabled = True
fraGlosa.Enabled = False
txtCuenta.SetFocus
nTasa = 0
End Sub

Private Sub cmdCancelar_Click()
LimpiaControles
End Sub

Private Sub cmdImprimir_Click()
Dim bReImp As Boolean
Dim sNomTit As String, sDirCli As String
Dim nMonto As Double
Dim sFormaRetiro As String, sCuenta As String
Dim nNumDuplicado As Integer
sCuenta = txtCuenta.NroCuenta
If MsgBox("¿ Desea Imprimir el Certificado de PF ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim clsMant As NCapMantenimiento
    Dim clsCap As NCapMovimientos
    Dim clsMov As NContFunciones
    Dim sLetras As String, sMovNro As String
    Set clsMov = New NContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    Set clsCap = New NCapMovimientos
    clsCap.EmiteDuplicadoCertPF sCuenta, sMovNro, Trim(txtGlosa)
    Set clsCap = Nothing
    Set clsMant = New NCapMantenimiento
    sNomTit = clsMant.GetNombreTitulares(sCuenta, True, 2, 0)
    Set clsMant = Nothing
    sDirCli = GetDireccionCliente
    nMonto = nCapitalInicial
    sLetras = ConversNL(nMoneda, nMonto)
    sFormaRetiro = Trim(lblFormaRetiro)
    bReImp = False
    nNumDuplicado = CInt(lblDuplicados) + 1
'    Do
'        ImprimeCertificadoPlazoFijo ldRenovacion, sNomTit, sDirCli, sCuenta, "1", CLng(lblPlazo), _
'                    nMonto, CDbl(nTasa), sFormaRetiro, sLetras, nNumDuplicado
'        If MsgBox("Desea reimprimir Certificado PF?? ", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
'            bReImp = True
'        Else
'            bReImp = False
'        End If
'    Loop Until Not bReImp
    LimpiaControles
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gCapPlazoFijo, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
Me.Caption = "Captaciones - Plazo Fijo - Duplicado Certificado Plazo Fijo"
Me.Icon = LoadPicture(App.path & gsRutaIcono)
txtCuenta.Prod = Trim(gCapPlazoFijo)
txtCuenta.EnabledProd = False
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledCMAC = False
cmdImprimir.Enabled = False
cmdCancelar.Enabled = False
fraCliente.Enabled = False
fraGlosa.Enabled = False
End Sub

Private Sub grdCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGlosa.SetFocus
End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCta As String
    sCta = txtCuenta.NroCuenta
    ObtieneDatosCuenta sCta
End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdImprimir.SetFocus
End If
End Sub




