VERSION 5.00
Begin VB.Form frmCapBusqCredPendPF 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   825
      Left            =   135
      TabIndex        =   10
      Top             =   2745
      Width           =   3435
      Begin VB.TextBox txtGlosa 
         Height          =   510
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   3210
      End
   End
   Begin VB.Frame fraClientes 
      Height          =   2115
      Left            =   135
      TabIndex        =   8
      Top             =   570
      Width           =   7845
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1755
         Left            =   90
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
   End
   Begin VB.Frame fraBusqueda 
      Height          =   870
      Left            =   3645
      TabIndex        =   7
      Top             =   2730
      Width           =   4335
      Begin VB.CheckBox chkBusqueda 
         Caption         =   "Buscar Créditos Pendientes de Pago en Cancelación ?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   570
         TabIndex        =   3
         Top             =   180
         Width           =   3315
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5805
      TabIndex        =   4
      Top             =   3660
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   135
      TabIndex        =   6
      Top             =   3660
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6945
      TabIndex        =   5
      Top             =   3660
      Width           =   1035
   End
   Begin SICMACT.ActXCodCta txtCuenta 
      Height          =   435
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   3660
      _ExtentX        =   5106
      _ExtentY        =   767
      Texto           =   "Cuenta N°"
      EnabledCta      =   -1  'True
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3855
      TabIndex        =   9
      Top             =   90
      Width           =   4110
   End
End
Attribute VB_Name = "frmCapBusqCredPendPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClearScreen()
txtCuenta.EnabledCMAC = False
txtCuenta.EnabledAge = False
txtCuenta.EnabledProd = False
txtCuenta.Prod = gCapPlazoFijo
txtCuenta.Age = gsCodAge
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Cuenta = ""
txtCuenta.Enabled = True
lblMensaje = ""
grdCliente.Clear
grdCliente.Rows = 2
grdCliente.FormaCabecera
txtGlosa.Text = ""
chkBusqueda.value = 0
End Sub

Private Sub cmdCancelar_Click()
txtCuenta.Enabled = True
ClearScreen
cmdCancelar.Enabled = False
CmdGrabar.Enabled = False
fraBusqueda.Enabled = False
fraGlosa.Enabled = False
txtCuenta.SetFocus
End Sub

Private Sub cmdGrabar_Click()
If Trim(txtGlosa.Text) = "" Then
    MsgBox "Debe digitar la glosa correspondiente.", vbInformation, "Aviso"
    txtGlosa.SetFocus
    Exit Sub
End If
If MsgBox("¿Desea grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim oCap As nCapDefinicion
    Dim sCuenta As String, sGlosa As String
    Dim bValor As Boolean
    Dim sMovNro As String
    Dim oCont As NContFunciones
    Set oCont = New NContFunciones
    sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oCont = Nothing
    sGlosa = Trim(txtGlosa.Text)
    sCuenta = txtCuenta.NroCuenta
    bValor = IIf(chkBusqueda.value = 1, True, False)
    Set oCap = New nCapDefinicion
    oCap.ActualizaCredPendPagoPF sCuenta, bValor, sMovNro, sGlosa
    Set oCap = Nothing
    cmdCancelar_Click
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        Dim clsGen As DGeneral
        Set clsGen = New DGeneral
        sCuenta = frmValTarCodAnt.Inicia(gCapPlazoFijo, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
Me.Caption = "Captaciones - Plazo Fijo - Búsqueda Créditos Pendientes de Pago"
ClearScreen
Me.Icon = LoadPicture(App.path & gsRutaIcono)
cmdCancelar.Enabled = False
CmdGrabar.Enabled = False
fraBusqueda.Enabled = False
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As NCapMantenimiento
Dim clsCap As NCapMovimientos
Dim rsCta As Recordset, rsRel As Recordset
Dim nEstado As CaptacEstado
Dim nRow As Long
Dim sMsg As String, sMoneda As String, sPersona As String
Dim dUltRetInt As Date
Dim nMoneda As Moneda

Set clsCap = New NCapMovimientos
sMsg = clsCap.ValidaCuentaOperacion(sCuenta)
If sMsg = "" Then
    Set clsMant = New NCapMantenimiento
    Set rsCta = New Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        nEstado = rsCta("nPrdEstado")
        nMoneda = CLng(Mid(sCuenta, 9, 1))
        If nMoneda = gMonedaNacional Then
            lblMensaje = "DEPOSITO A PLAZO FIJO" & Chr$(13) & "MONEDA NACIONAL"
        ElseIf nMoneda = gMonedaExtranjera Then
            lblMensaje = "DEPOSITO A PLAZO FIJO" & Chr$(13) & "MONEDA EXTRANJERA"
        End If
        chkBusqueda.value = IIf(rsCta("bBusCredPend"), 1, 0)
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
        fraBusqueda.Enabled = True
        txtCuenta.Enabled = False
        cmdCancelar.Enabled = True
        CmdGrabar.Enabled = True
        fraGlosa.Enabled = True
        txtGlosa.SetFocus
    End If
Else
    MsgBox sMsg, vbInformation, "Operacion"
    txtCuenta.SetFocus
End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCuenta As String
    sCuenta = txtCuenta.NroCuenta
    ObtieneDatosCuenta sCuenta
End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    chkBusqueda.SetFocus
End If
End Sub
