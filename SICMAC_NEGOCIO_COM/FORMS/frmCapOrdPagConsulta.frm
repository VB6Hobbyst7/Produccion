VERSION 5.00
Begin VB.Form frmCapOrdPagConsulta 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   Icon            =   "frmCapOrdPagConsulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOrdenPago 
      Caption         =   "Orden Pago"
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
      Height          =   2235
      Left            =   3870
      TabIndex        =   10
      Top             =   2760
      Width           =   6045
      Begin VB.TextBox txtOrdPag 
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
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   4
         Top             =   300
         Width           =   1215
      End
      Begin SICMACT.FlexEdit grdOrdPag 
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   2355
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Orden-Estado-Monto-Fecha-Usu-cEstado"
         EncabezadosAnchos=   "250-850-1700-1000-1000-600-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden N° :"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   390
         Width           =   750
      End
   End
   Begin VB.Frame fraCuentas 
      Height          =   2745
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   9855
      Begin VB.Frame Frame1 
         Height          =   1995
         Left            =   4380
         TabIndex        =   23
         Top             =   660
         Width           =   5355
         Begin SICMACT.FlexEdit grdCliente 
            Height          =   1635
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   2884
            Cols0           =   3
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Persona-Relación"
            EncabezadosAnchos=   "250-3500-1200"
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
            ColumnasAEditar =   "X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L"
            FormatosEdit    =   "0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   255
            RowHeight0      =   300
         End
      End
      Begin VB.Frame fraDatosCuenta 
         Height          =   1995
         Left            =   120
         TabIndex        =   12
         Top             =   660
         Width           =   4215
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   1320
            TabIndex        =   22
            Top             =   885
            Width           =   2055
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cuenta :"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   945
            Width           =   960
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   1320
            TabIndex        =   20
            Top             =   1260
            Width           =   2055
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Tasa :"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   1680
            Width           =   810
         End
         Begin VB.Label lblTipoTasa 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   1320
            TabIndex        =   18
            Top             =   1620
            Width           =   2055
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Firmas :"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   540
         End
         Begin VB.Label lblEstado 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   1320
            TabIndex        =   16
            Top             =   525
            Width           =   2055
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Estado :"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   585
            Width           =   585
         End
         Begin VB.Label lblApertura 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   1320
            TabIndex        =   14
            Top             =   165
            Width           =   2055
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   225
            Width           =   690
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   350
         Left            =   3780
         TabIndex        =   1
         Top             =   240
         Width           =   435
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   105
         TabIndex        =   0
         Top             =   210
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   767
         Texto           =   "Cuenta N°"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin VB.Frame fraOrdPagEmi 
      Caption         =   "Orden Pago Emitidas"
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
      Height          =   2220
      Left            =   60
      TabIndex        =   8
      Top             =   2760
      Width           =   3795
      Begin SICMACT.FlexEdit grdOrdPagEmi 
         Height          =   1905
         Left            =   105
         TabIndex        =   3
         Top             =   210
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   3360
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Inicial-Final-Fecha-Usu"
         EncabezadosAnchos=   "250-800-800-1000-550"
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
         ColumnasAEditar =   "X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   300
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5100
      TabIndex        =   7
      Top             =   5100
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3780
      TabIndex        =   6
      Top             =   5100
      Width           =   1170
   End
End
Attribute VB_Name = "frmCapOrdPagConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LimpiaPantalla()
grdCliente.Clear
grdCliente.Rows = 2
grdCliente.FormaCabecera
grdOrdPagEmi.Clear
grdOrdPagEmi.Rows = 2
grdOrdPagEmi.FormaCabecera
grdOrdPag.Clear
grdOrdPag.Rows = 2
grdOrdPag.FormaCabecera
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Age = Right(gsCodAge, 2)
txtCuenta.Prod = Trim(Str(gCapAhorros))
txtCuenta.Cuenta = ""
txtCuenta.EnabledCMAC = False
txtCuenta.EnabledAge = False
txtCuenta.EnabledProd = False
cmdCancelar.Enabled = False
txtCuenta.Enabled = True
cmdBuscar.Enabled = True
txtOrdPag = ""
lblApertura = ""
lblEstado = ""
lblFirmas = ""
lblTipoTasa = ""
lblTipoCuenta = ""
fraOrdenPago.Enabled = False
End Sub

Private Sub ObtieneDatosOrdenPago(ByVal sCuenta As String, ByVal nNumOP As Long)
Dim rsOrd As ADODB.Recordset
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsOrd = clsMant.GetDatosOrdenPago(sCuenta, nNumOP)
If rsOrd.EOF And rsOrd.BOF Then
    MsgBox "Orden de Pago no registra movimientos.", vbInformation, "Aviso"
    txtOrdPag.SetFocus
Else
    Set grdOrdPag.Recordset = rsOrd
    grdOrdPag.SetFocus
End If
rsOrd.Close
Set clsMant = Nothing
Set rsOrd = Nothing
End Sub

Private Sub ObtieneDatosOrdenPagoEmi(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsCta As ADODB.Recordset
Set rsCta = New ADODB.Recordset
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsCta = clsMant.GetOrdenPagoEmitidas(sCuenta)
Set clsMant = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    Set grdOrdPagEmi.Recordset = rsCta
End If
rsCta.Close
Set rsCta = Nothing
End Sub

Private Sub ObtieneDatosPersona(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsCta As ADODB.Recordset
Dim nIndex As Long
Dim sPers As String
Set rsCta = New ADODB.Recordset
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsCta = clsMant.GetPersonaCuenta(sCuenta)
If Not (rsCta.EOF And rsCta.BOF) Then
    sPers = ""
    Do While Not rsCta.EOF
        If sPers <> rsCta("cPersCod") Then
            grdCliente.AdicionaFila
            nIndex = grdCliente.Rows - 1
            grdCliente.TextMatrix(nIndex, 1) = rsCta("Nombre")
            grdCliente.TextMatrix(nIndex, 2) = UCase(rsCta("Relacion"))
            sPers = rsCta("cPersCod")
        End If
        rsCta.MoveNext
    Loop
Else
    MsgBox "Cuenta no posee relacion con Persona", vbExclamation, "Aviso"
    txtCuenta.SetFocusCuenta
End If
rsCta.Close
Set clsMant = Nothing
Set rsCta = Nothing
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsCta As ADODB.Recordset, rsRel As ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim sSQL As String

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsCta = New ADODB.Recordset
Set rsCta = clsMant.GetDatosCuenta(sCuenta)
Set clsMant = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    nEstado = rsCta("nPrdEstado")
    If nEstado <> gCapEstAnulada And nEstado <> gCapEstCancelada Then
        If rsCta("bOrdPag") Then
            lblApertura = UCase(Format$(rsCta("dApertura"), "dd-mmm-yyyy"))
            lblEstado = UCase(rsCta("cEstado"))
            lblTipoCuenta = UCase(rsCta("cTipoCuenta"))
            lblFirmas = rsCta("nFirmas")
            lblTipoTasa = UCase(rsCta("cTipoTasa"))
            ObtieneDatosPersona sCuenta
            ObtieneDatosOrdenPagoEmi sCuenta
            cmdCancelar.Enabled = True
            txtCuenta.Enabled = False
            cmdBuscar.Enabled = False
            fraOrdenPago.Enabled = True
            txtOrdPag.SetFocus
        Else
            MsgBox "Cuenta no fue aperturada para emitir Ordenes de Pago.", vbInformation, "Aviso"
            txtCuenta.SetFocusCuenta
        End If
    Else
        MsgBox "Cuenta Anulada o Cancelada", vbInformation, "Aviso"
        txtCuenta.SetFocusCuenta
    End If
Else
    MsgBox "Cuenta no existe", vbInformation, "Aviso"
    txtCuenta.SetFocusCuenta
End If
rsCta.Close
Set rsCta = Nothing
End Sub

Private Sub cmdBuscar_Click()
Dim clsPers As COMDPersona.UCOMPersona
Set clsPers = New COMDPersona.UCOMPersona
Set clsPers = frmBuscaPersona.Inicio
If Not clsPers Is Nothing Then
    Dim sPers As String
    Dim rsPers As New ADODB.Recordset
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim sCta As String
    Dim sRelac As String * 15
    Dim sEstado As String
    Dim clsCuenta As UCapCuenta
    sPers = clsPers.sPerscod
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsPers = clsCap.GetCuentasPersona(sPers, gCapAhorros, True, , , True)
    Set clsCap = Nothing
    If Not (rsPers.EOF And rsPers.EOF) Then
        Do While Not rsPers.EOF
            sCta = rsPers("cCtaCod")
            sRelac = rsPers("cRelacion")
            sEstado = Trim(rsPers("cEstado"))
            frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
            rsPers.MoveNext
        Loop
        Set clsCuenta = New UCapCuenta
        Set clsCuenta = frmCapMantenimientoCtas.Inicia
        If clsCuenta.sCtaCod <> "" Then
            txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
            txtCuenta.Prod = Mid(clsCuenta.sCtaCod, 6, 3)
            txtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
        End If
        Set clsCuenta = Nothing
    Else
        MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
    End If
    rsPers.Close
    Set rsPers = Nothing
End If
txtCuenta.SetFocusCuenta
Set clsPers = Nothing
End Sub

Private Sub cmdCancelar_Click()
LimpiaPantalla
txtCuenta.SetFocusCuenta
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gCapAhorros, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
Me.Caption = "Captaciones - Orden Pago - Consulta"
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Age = Right(gsCodAge, 2)
txtCuenta.Prod = Trim(Str(gCapAhorros))
txtCuenta.EnabledCMAC = False
txtCuenta.EnabledAge = False
txtCuenta.EnabledProd = False
cmdCancelar.Enabled = False
fraOrdenPago.Enabled = False
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCta As String
    sCta = txtCuenta.NroCuenta
    ObtieneDatosCuenta sCta
End If
End Sub

Private Sub txtOrdPag_GotFocus()
With txtOrdPag
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtOrdPag_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCuenta As String
    Dim nNumOP As Long
    sCuenta = txtCuenta.NroCuenta
    nNumOP = CLng(Trim(txtOrdPag))
    ObtieneDatosOrdenPago sCuenta, nNumOP
Else
    KeyAscii = NumerosEnteros(KeyAscii)
End If
End Sub
