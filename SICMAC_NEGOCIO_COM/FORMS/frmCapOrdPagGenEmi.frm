VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCapOrdPagGenEmi 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   Icon            =   "frmCapOrdPagGenEmi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEmitir 
      Caption         =   "&Emitir"
      Height          =   375
      Left            =   5355
      TabIndex        =   7
      Top             =   5145
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   105
      TabIndex        =   9
      Top             =   5145
      Width           =   1170
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6615
      TabIndex        =   8
      Top             =   5145
      Width           =   1170
   End
   Begin VB.Frame fraEmision 
      Caption         =   "Emisión"
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
      Left            =   3780
      TabIndex        =   12
      Top             =   2835
      Width           =   4005
      Begin VB.CheckBox chkDctoImp 
         Alignment       =   1  'Right Justify
         Caption         =   "Descontar??"
         Height          =   330
         Left            =   2415
         TabIndex        =   5
         Top             =   840
         Width           =   1380
      End
      Begin VB.TextBox txtInicio 
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
         Left            =   840
         MaxLength       =   7
         TabIndex        =   6
         Top             =   1470
         Width           =   1065
      End
      Begin VB.Label lblOrdPagTal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2415
         TabIndex        =   19
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Num. Ord Pag. x Talonario"
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   375
         Width           =   1875
      End
      Begin VB.Label lblFin 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   2415
         TabIndex        =   17
         Top             =   1470
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Costo Emisión"
         Height          =   195
         Left            =   210
         TabIndex        =   16
         Top             =   908
         Width           =   990
      End
      Begin VB.Label lblDctoImp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1365
         TabIndex        =   15
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al :"
         Height          =   195
         Left            =   2100
         TabIndex        =   14
         Top             =   1575
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del :"
         Height          =   195
         Left            =   420
         TabIndex        =   13
         Top             =   1575
         Width           =   330
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
      Left            =   105
      TabIndex        =   11
      Top             =   2835
      Width           =   3585
      Begin SICMACT.FlexEdit grdOrdPagEmi 
         Height          =   1905
         Left            =   105
         TabIndex        =   4
         Top             =   210
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3360
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Inicial-Final-Fecha-Usu"
         EncabezadosAnchos=   "250-725-725-1000-550"
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
         ColumnasAEditar =   "X-X-X-X-X"
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
   Begin VB.Frame fraCuentas 
      Height          =   2745
      Left            =   80
      TabIndex        =   10
      Top             =   0
      Width           =   7785
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
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   350
         Left            =   3780
         TabIndex        =   1
         Top             =   240
         Width           =   500
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
End
Attribute VB_Name = "frmCapOrdPagGenEmi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nOrdPagTal As Double

Public Sub ImprimeOP(ByVal nCantOP As Integer, ByVal sCuenta As String)
Dim intLinBla As Integer
Dim intContOP As Integer
Dim nFicSal As Integer
Dim k As Integer, i As Integer
Dim lsNomCli As String, sCtaBco As String
Dim lnLngNom As Integer
Dim clsMant As NCapMantenimiento
Dim clsGen As nCapDefinicion
intLinBla = 8
intContOP = 0
Set clsMant = New NCapMantenimiento
lsNomCli = clsMant.GetNombreTitulares(sCuenta)
Set clsMant = Nothing
Set clsGen = New nCapDefinicion
sCtaBco = clsGen.GetCapParametroDesc(gCtaBacoOrdPag)
Set clsGen = Nothing
nFicSal = FreeFile
Open sLpt For Output As nFicSal
Print #nFicSal, Chr$(15);                           'Establece tipo de letra condensada
Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(1);     'Tipo de Letra Sans Serif
Print #nFicSal, Chr$(27) + Chr$(77);                'Tamaño 10.5 - 12 CPI
Print #nFicSal, Chr$(27) + Chr$(50);                'Espaciado entre lineas 1/16
Print #nFicSal, Chr$(27) + Chr$(108) + Chr$(44);     'Margen Izquierdo  44
Print #nFicSal, Chr$(27) + Chr$(67) + Chr$(66);     'Longitud de página a 16 líneas'
For k = 1 To nCantOP
    intContOP = intContOP + 1
    If intContOP > 4 Then
       intContOP = 1
       Print #nFicSal, Chr$(12);                                      'Avance de Página
    End If
    Select Case intContOP
        Case 1
             intLinBla = 8
        Case 2
             intLinBla = 14
        Case 3
             intLinBla = 14
        Case 4
             intLinBla = 14
    End Select
    For i = 1 To intLinBla - 1
        Print #nFicSal, ""
    Next i
    Print #nFicSal, Chr$(27) + Chr$(69);                           'Establece tipo de letra negrita
    Print #nFicSal, ImpreCarEsp(lsNomCli)
    Print #nFicSal, sCuenta
    'Print #nFicSal, psCtaBco                                       '      "Cta. 710-1631918 Bco. WIESE"
    Print #nFicSal, gsNomAge                                       '      "Nombre de Agencia"
    Print #nFicSal, Chr$(27) + Chr$(70);                           'Desactiva tipo de letra negrita
Next k
Print #nFicSal, Chr$(18);                           'Retorna al tipo de letra normal
Print #nFicSal, Chr$(27) + Chr$(108) + Chr$(6);     'Margen Izquierdo  44
Close nFicSal
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
txtCuenta.Cuenta = ""
chkDctoImp.value = 1
lblDctoImp = ""
cmdEmitir.Enabled = False
cmdCancelar.Enabled = False
txtCuenta.Enabled = True
cmdBuscar.Enabled = True
txtInicio = ""
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
Dim clsMant As NCapMantenimiento
Dim rsCta As Recordset
Set rsCta = New Recordset
Set clsMant = New NCapMantenimiento
Set rsCta = clsMant.GetOrdenPagoEmitidas(sCuenta)
Set clsMant = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    Set grdOrdPagEmi.Recordset = rsCta
End If
rsCta.Close
Set rsCta = Nothing
End Sub

Private Sub ObtieneDatosPersona(ByVal sCuenta As String)
Dim clsMant As NCapMantenimiento
Dim rsCta As Recordset

Set rsCta = New Recordset
Set clsMant = New NCapMantenimiento
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
Set rsCta = Nothing
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As NCapMantenimiento
Dim rsCta As Recordset, rsRel As Recordset
Dim nEstado As CaptacEstado
Dim sSQL As String

Set clsMant = New NCapMantenimiento
Set rsCta = New Recordset
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
            Dim clsGen As nCapDefinicion
            Set clsGen = New nCapDefinicion
            If CLng(Mid(sCuenta, 9, 1)) = gMonedaNacional Then
                lblDctoImp.BackColor = &HC0FFFF
                lblDctoImp = "S/. " & Format$(clsGen.GetCapParametro(gCostoChqMN), "#,##0.00")
            Else
                lblDctoImp.BackColor = &HC0FFC0
                lblDctoImp = "$ " & Format$(clsGen.GetCapParametro(gCostoChqME), "#,##0.00")
            End If
            Set clsGen = Nothing
            chkDctoImp.value = 1
            cmdCancelar.Enabled = True
            txtCuenta.Enabled = False
            cmdBuscar.Enabled = False
            txtInicio.SetFocus
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
Dim clsPers As UPersona
Set clsPers = New UPersona
Set clsPers = frmBuscaPersona.Inicio
If Not clsPers Is Nothing Then
    Dim sPers As String
    Dim rsPers As Recordset
    Dim clsCap As NCapMantenimiento
    Dim sCta As String
    Dim sRelac As String * 15
    Dim sEstado As String
    Dim clsCuenta As UCapCuentas
    sPers = clsPers.sPersCod
    Set clsCap = New NCapMantenimiento
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
        Set clsCuenta = New UCapCuentas
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
End Sub

Private Sub cmdCancelar_Click()
LimpiaPantalla
txtCuenta.SetFocusCuenta
End Sub

Private Sub cmdEmitir_Click()
If MsgBox("¿Desea emitir las ordenes de pago?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim clsMant As NCapMantenimiento
    Dim clsMov As NContFunciones
    Dim sCuenta As String, sMovNro As String, sMonto As String
    Dim nInicio As Long, nFin As Long
    Dim bDescuento As Boolean
    Dim nMontoDcto As Double
    Dim nNumOP As Integer
    Set clsMov = New NContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Set clsMov = Nothing
    sCuenta = txtCuenta.NroCuenta
    bDescuento = IIf(chkDctoImp.value = 1, True, False)
    If bDescuento Then
        sMonto = Replace(lblDctoImp, "S/.", "", 1, , vbTextCompare)
        sMonto = Replace(lblDctoImp, "$", "", 1, , vbTextCompare)
        nMontoDcto = CDbl(Trim(sMonto))
    Else
        nMontoDcto = 0
    End If
    nNumOP = CInt(lblOrdPagTal)
    Set clsMant = New NCapMantenimiento
    On Error GoTo ErrGraba
    nInicio = CLng(txtInicio)
    nFin = CLng(lblFin)
    clsMant.EmiteRangoOrdenPago sCuenta, nInicio, nFin, sMovNro, bDescuento, nMontoDcto
    'ImprimeOP nNumOp, sCuenta
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
Me.Caption = "Captaciones - Orden Pago - Generacion y Emisión"
SetupGridCliente
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Age = Right(gsCodAge, 2)
txtCuenta.Prod = Trim(Str(gCapAhorros))
txtCuenta.EnabledCMAC = False
txtCuenta.EnabledAge = False
txtCuenta.EnabledProd = False
chkDctoImp.value = 1
Dim clsGen As nCapDefinicion
Set clsGen = New nCapDefinicion
nOrdPagTal = clsGen.GetCapParametro(gNumOrdPagTal)
lblOrdPagTal = Trim(nOrdPagTal)
Set clsGen = Nothing
cmdEmitir.Enabled = False
cmdCancelar.Enabled = False
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCta As String
    sCta = txtCuenta.NroCuenta
    ObtieneDatosCuenta sCta
End If
End Sub

Private Sub txtInicio_Change()
If txtInicio <> "" Then
    lblFin = Trim(Str(CLng(txtInicio) + nOrdPagTal) - 1)
    cmdEmitir.Enabled = True
Else
    lblFin = ""
    cmdEmitir.Enabled = False
End If
End Sub

Private Sub txtInicio_GotFocus()
With txtInicio
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtInicio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdEmitir.Enabled Then
        cmdEmitir.SetFocus
        Exit Sub
    End If
    
End If
KeyAscii = NumerosEnteros(KeyAscii)

End Sub
