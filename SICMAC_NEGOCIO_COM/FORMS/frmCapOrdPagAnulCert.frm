VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCapOrdPagAnulCert 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "frmCapOrdPagAnulCert.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAnulCert 
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
      Height          =   2325
      Left            =   3780
      TabIndex        =   16
      Top             =   2835
      Width           =   4110
      Begin VB.TextBox txtFin 
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
         Height          =   330
         Left            =   2520
         MaxLength       =   8
         ScrollBars      =   1  'Horizontal
         TabIndex        =   8
         Top             =   570
         Width           =   1380
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "&Rango"
         Height          =   330
         Index           =   1
         Left            =   2520
         TabIndex        =   6
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "&Individual"
         Height          =   330
         Index           =   0
         Left            =   735
         TabIndex        =   5
         Top             =   210
         Width           =   1065
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   330
         Left            =   735
         TabIndex        =   9
         Top             =   990
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         Text            =   "0"
      End
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
         Height          =   330
         Left            =   735
         MaxLength       =   8
         TabIndex        =   7
         Top             =   570
         Width           =   1380
      End
      Begin VB.TextBox txtGlosa 
         Height          =   750
         Left            =   735
         TabIndex        =   10
         Top             =   1410
         Width           =   3270
      End
      Begin VB.Label lblFin 
         Caption         =   "Al :"
         Height          =   225
         Left            =   2205
         TabIndex        =   20
         Top             =   630
         Width           =   225
      End
      Begin VB.Label lblIni 
         AutoSize        =   -1  'True
         Caption         =   "Orden :"
         Height          =   195
         Left            =   105
         TabIndex        =   19
         Top             =   630
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   1410
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   1058
         Width           =   540
      End
   End
   Begin VB.Frame fraCuentas 
      Height          =   2745
      Left            =   105
      TabIndex        =   15
      Top             =   0
      Width           =   7785
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   350
         Left            =   3780
         TabIndex        =   1
         Top             =   240
         Width           =   500
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCliente 
         Height          =   1170
         Left            =   105
         TabIndex        =   3
         Top             =   1470
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2064
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
      Begin SICMACT.FlexEdit grdCuentas 
         Height          =   645
         Left            =   105
         TabIndex        =   2
         Top             =   735
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   1138
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Cuenta-Apertura-Estado-Tipo Cuenta-Firm-Tipo Tasa"
         EncabezadosAnchos=   "250-1900-1000-1200-1200-400-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColor       =   16777215
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   240
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
      Height          =   2325
      Left            =   135
      TabIndex        =   14
      Top             =   2835
      Width           =   3585
      Begin SICMACT.FlexEdit grdOrdPagEmi 
         Height          =   2010
         Left            =   105
         TabIndex        =   4
         Top             =   210
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3545
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Inicial-Final-Fecha-Usu"
         EncabezadosAnchos=   "250-725-725-1000-550"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
         RowHeight0      =   285
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6645
      TabIndex        =   12
      Top             =   5250
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   135
      TabIndex        =   13
      Top             =   5250
      Width           =   1170
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5385
      TabIndex        =   11
      Top             =   5250
      Width           =   1170
   End
End
Attribute VB_Name = "frmCapOrdPagAnulCert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nOperacion As COMDConstantes.CaptacOperacion
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by


Public Function EsOrdenEmitida(ByVal nNumOP As Long) As Boolean
Dim i As Integer
Dim nInicio As Long, nFin As Long, nFila As Long, nCol As Long
Dim bexiste As Boolean
bexiste = False
If grdOrdPagEmi.TextMatrix(1, 1) <> "" Then
    nFila = grdOrdPagEmi.Row
    nCol = grdOrdPagEmi.Col
    For i = 1 To grdOrdPagEmi.Rows - 1
        nInicio = CLng(grdOrdPagEmi.TextMatrix(i, 1))
        nFin = CLng(grdOrdPagEmi.TextMatrix(i, 2))
        If nNumOP >= nInicio And nNumOP <= nFin Then
            bexiste = True
            Exit For
        End If
    Next i
    grdOrdPagEmi.Row = nFila
    grdOrdPagEmi.Col = nCol
End If
EsOrdenEmitida = bexiste
End Function

Public Sub Inicia(ByVal nOpe As COMDConstantes.CaptacOperacion)
nOperacion = nOpe
If nOperacion = gAhoOPAnulacion Then
    fraAnulCert.Caption = "Anular Orden Pago"
    Me.Caption = "Captaciones - Orden Pago - Anular"
    optTipo(0).Visible = True
    optTipo(1).Visible = True
    optTipo(0).value = True
    'By Capi 20012009
    Set objPista = New COMManejador.Pista
    gsOpeCod = gAhoAnulacionOrdPago
    'End By

ElseIf nOperacion = gAhoOPCertificacion Then
    fraAnulCert.Caption = "Certificar Orden Pago"
    Me.Caption = "Captaciones - Orden Pago - Certificar"
    optTipo(0).Visible = False
    optTipo(1).Visible = False
    lblFin.Visible = False
    txtFin.Visible = False
    'By Capi 20012009
    Set objPista = New COMManejador.Pista
    gsOpeCod = gAhoCertificacionOrdPago
    'End By

End If
SetupGridCliente
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Age = Right(gsCodAge, 2)
txtCuenta.Prod = Trim(Str(gCapAhorros))
txtCuenta.EnabledCMAC = False
txtCuenta.EnabledAge = False
txtCuenta.EnabledProd = False
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
fraAnulCert.Enabled = False
Me.Show 1
End Sub

Private Sub LimpiaPantalla()
grdCliente.Clear
grdCliente.Rows = 2
SetupGridCliente
grdCuentas.Clear
grdCuentas.Rows = 2
grdCuentas.FormaCabecera
grdOrdPagEmi.Clear
grdOrdPagEmi.Rows = 2
grdOrdPagEmi.FormaCabecera
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Age = Right(gsCodAge, 2)
txtCuenta.Prod = Trim(Str(gCapAhorros))
txtCuenta.cuenta = ""
txtCuenta.EnabledCMAC = False
txtCuenta.EnabledAge = False
txtCuenta.EnabledProd = False
txtMonto.value = 0
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
txtCuenta.Enabled = True
cmdBuscar.Enabled = True
txtOrdPag = ""
txtFin = ""
txtGlosa = ""
fraAnulCert.Enabled = False
End Sub

Private Sub SetupGridCliente()
Dim i As Integer
For i = 1 To grdCliente.Rows - 1
    grdCliente.MergeCol(i) = True
Next i
grdCliente.MergeCells = flexMergeFree
grdCliente.Cols = 12
grdCliente.ColWidth(0) = 100
grdCliente.ColWidth(1) = 3500
grdCliente.ColWidth(2) = 400
grdCliente.ColWidth(3) = 3500
grdCliente.ColWidth(4) = 1500
grdCliente.ColWidth(5) = 1000
grdCliente.ColWidth(6) = 600
grdCliente.ColWidth(7) = 1500
grdCliente.ColWidth(8) = 0
grdCliente.ColWidth(9) = 0
grdCliente.ColWidth(10) = 0
grdCliente.ColWidth(11) = 0

grdCliente.TextMatrix(0, 1) = "Nombre"
grdCliente.TextMatrix(0, 2) = "RE"
grdCliente.TextMatrix(0, 3) = "Direccion"
grdCliente.TextMatrix(0, 4) = "Zona"
grdCliente.TextMatrix(0, 5) = "Fono"
grdCliente.TextMatrix(0, 6) = "ID"
grdCliente.TextMatrix(0, 7) = "ID N°"
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

Set rsCta = New ADODB.Recordset
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsCta = clsMant.GetPersonaCuenta(sCuenta)
If Not (rsCta.EOF And rsCta.BOF) Then
    Set grdCliente.DataSource = rsCta
    SetupGridCliente
Else
    MsgBox "Cuenta no posee relacion con Persona", vbExclamation, "Aviso"
    txtCuenta.SetFocusCuenta
    grdCuentas.Clear
    grdCuentas.Rows = 2
    grdCuentas.FormaCabecera
End If
rsCta.Close
Set clsMant = Nothing
Set rsCta = Nothing
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsCta As ADODB.Recordset, rsRel As ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim ssql As String

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsCta = New ADODB.Recordset
Set rsCta = clsMant.GetDatosCuenta(sCuenta)
Set clsMant = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    nEstado = rsCta("nPrdEstado")
    If nEstado <> gCapEstAnulada And nEstado <> gCapEstCancelada Then
        If rsCta("bOrdPag") Then
            Dim nItem As Long
            grdCuentas.Clear
            grdCuentas.FormaCabecera
            grdCuentas.Rows = 2
            grdCuentas.AdicionaFila
            nItem = grdCuentas.Row
            grdCuentas.TextMatrix(nItem, 1) = rsCta("cCtaCod")
            grdCuentas.TextMatrix(nItem, 2) = Format$(rsCta("dApertura"), "dd-mm-yyyy")
            grdCuentas.TextMatrix(nItem, 3) = rsCta("cEstado")
            grdCuentas.TextMatrix(nItem, 4) = rsCta("cTipoCuenta")
            grdCuentas.TextMatrix(nItem, 5) = rsCta("nFirmas")
            grdCuentas.TextMatrix(nItem, 6) = rsCta("cTipoTasa")
            
            ObtieneDatosPersona sCuenta
            ObtieneDatosOrdenPagoEmi sCuenta
            Dim clsGen As COMDConstSistema.DCOMGeneral
            Set clsGen = New COMDConstSistema.DCOMGeneral
            If CLng(Mid(sCuenta, 9, 1)) = gMonedaNacional Then
                txtMonto.BackColor = &HC0FFFF
            Else
                txtMonto.BackColor = &HC0FFC0
            End If
            Set clsGen = Nothing
            fraAnulCert.Enabled = True
            cmdCancelar.Enabled = True
            txtCuenta.Enabled = False
            cmdBuscar.Enabled = False
            cmdGrabar.Enabled = True
            If nOperacion = gAhoOPCertificacion Then
                txtOrdPag.SetFocus
            ElseIf nOperacion = gAhoOPAnulacion Then
                optTipo(0).SetFocus
            End If
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
    Dim rsPers As ADODB.Recordset
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim sCta As String
    Dim sRelac As String * 15
    Dim sEstado As String
    Dim clsCuenta As UCapCuenta
    sPers = clsPers.sPersCod
    Set clsCap = New COMDCaptaGenerales.DCOMCaptaGenerales
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
            txtCuenta.cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
        End If
        Set clsCuenta = Nothing
    Else
        MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
    End If
    rsPers.Close
    Set rsPers = Nothing
End If
Set clsPers = Nothing
txtCuenta.SetFocusCuenta
End Sub

Private Sub cmdCancelar_Click()
LimpiaPantalla
txtCuenta.SetFocusCuenta
End Sub

Private Sub CmdGrabar_Click()
Dim nMonto As Double
Dim sCuenta As String
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim bOPExiste As Boolean
sCuenta = txtCuenta.NroCuenta
nMonto = txtMonto.value
If txtOrdPag = "" Then
    MsgBox "Debe digitar una orden de pago válida.", vbInformation, "Aviso"
    txtOrdPag.SetFocus
    Exit Sub
End If
If nOperacion = gAhoOPCertificacion Then
    If nMonto = 0 Then
        MsgBox "Monto debe ser mayor a cero", vbInformation, "Aviso"
        txtOrdPag.SetFocus
        Exit Sub
    End If
    If Not EsOrdenEmitida(CLng(txtOrdPag)) Then
        MsgBox "Orden de Pago no Emitida", vbInformation, "Aviso"
        txtOrdPag.SetFocus
        Exit Sub
    End If
    Dim rsOP As ADODB.Recordset
    Dim nEstadoOP As CaptacOrdPagoEstado
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsOP = clsMant.GetDatosOrdenPago(sCuenta, CLng(Val(txtOrdPag)))
    Set clsMant = Nothing
    If Not (rsOP.EOF And rsOP.BOF) Then
        nEstadoOP = rsOP("nEstado")
        If nEstadoOP = gCapOPEstAnulada Or nEstadoOP = gCapOPEstCobrada Or nEstadoOP = gCapOPEstAnulada Or _
            nEstadoOP = gCapOPEstExtraviada Or nEstadoOP = gCapOPEstRechazada Or nEstadoOP = gCapOPEstCertifiCada Then
            MsgBox "Orden de Pago N° " & Trim(txtOrdPag.Text) & " " & rsOP("cDescripcion"), vbInformation, "Aviso"
            rsOP.Close
            Set rsOP = Nothing
            txtOrdPag.SetFocus
            Exit Sub
        End If
        rsOP.Close
        Set rsOP = Nothing
        bOPExiste = True
    Else
        bOPExiste = False
    End If
End If
If Trim(txtGlosa) = "" Then
    MsgBox "Debe escribir la glosa correspondiente al movimiento.", vbInformation, "Aviso"
    txtGlosa.SetFocus
    Exit Sub
End If
If optTipo(1).value And Trim(txtFin) = "" Then
    MsgBox "Debe colocar un valor válido en el rango final.", vbInformation, "Aviso"
    txtFin.SetFocus
    Exit Sub
End If

If MsgBox("¿Desea grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim clsMov As COMNContabilidad.NCOMContFunciones  'NContFunciones
    Dim sMovNro As String
    Dim nInicio As Long, nFin As Long
    Dim lsMensaje As String
    Dim lsBoleta As String
    Dim lsBoletaITF As String
    Dim nFicSal As Integer
    
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    sCuenta = txtCuenta.NroCuenta
    nInicio = CLng(txtOrdPag)
    On Error GoTo ErrGraba
    Select Case nOperacion
        Case gAhoOPCertificacion
            Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
            Dim nSaldo As Double
            Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
            If Not clsCap.ValidaSaldoCuenta(sCuenta, nMonto) Then
                MsgBox "La cuenta NO posee SALDO SUFICIENTE.", vbInformation, "Aviso"
                Set clsCap = Nothing
                Exit Sub
            End If
            nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, gAhoOPCertificacion, sMovNro, Trim(txtGlosa.Text), TpoDocOrdenPago, txtOrdPag.Text, , bOPExiste, True, False, , , gsNomAge, sLpt, , , , , , , , , , , , , , , lsMensaje, lsBoleta, lsBoletaITF)
            'By Capi 21012009
              objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Certificacion OP", txtCuenta.NroCuenta, gcodigocuenta
            'End by

            Set clsCap = Nothing
            If Trim(lsMensaje) <> "" Then
                MsgBox lsMensaje, vbInformation, "Aviso"
            End If
            
            If Trim(lsBoleta) <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, lsBoleta
                    Print #nFicSal, ""
                Close #nFicSal
            End If
            
            If Trim(lsBoletaITF) <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, lsBoletaITF
                    Print #nFicSal, ""
                Close #nFicSal
            End If
        Case gAhoOPAnulacion
            
            Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
            If optTipo(1).value Then
                nFin = CLng(txtFin)
                clsMant.AnulaRangoOrdenPago sCuenta, sMovNro, nInicio, nFin, txtGlosa.Text
            Else
                clsMant.AnulaOrdenPago sCuenta, sMovNro, nInicio, txtGlosa.Text, nMonto
            End If
             'By Capi 21012009
              objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Anulacion OP", txtCuenta.NroCuenta, gcodigocuenta
            'End by
    End Select
    Set clsMant = Nothing
    Set clsCap = Nothing
    cmdCancelar_Click
End If
Exit Sub
ErrGraba:
    MsgBox Err.Description, vbExclamation, "Aviso"
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
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub optTipo_Click(Index As Integer)
Select Case Index
    Case 0
        txtFin.Visible = False
        lblFin.Visible = False
        lblIni.Caption = "Orden :"
        txtMonto.Visible = True
    Case 1
        txtFin.Visible = True
        lblFin.Visible = True
        lblIni.Caption = "Del :"
        txtFin.Text = ""
        txtMonto.Visible = False
End Select
End Sub

Private Sub optTipo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    txtOrdPag.SetFocus
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
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGlosa.SetFocus
End If
End Sub

Private Sub txtFin_GotFocus()
With txtFin
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtOrdPag_GotFocus()
With txtOrdPag
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtOrdPag_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtFin.Visible Then
        txtFin.SetFocus
    Else
        txtMonto.Enabled = True
        txtMonto.SetFocus
    End If
    Exit Sub
End If
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub TxtFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtMonto.Visible Then
        txtMonto.SetFocus
    Else
        txtGlosa.SetFocus
    End If
    Exit Sub
End If
KeyAscii = NumerosEnteros(KeyAscii)
End Sub
