VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogRetasacion 
   Caption         =   "Retasacion de Bienes"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   9360
   Begin VB.Frame frmBotones 
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   7680
      Width           =   9180
      Begin VB.CommandButton cmdImprimirTodos 
         Caption         =   "Imprimir todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1600
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdVender 
         Caption         =   "&Vender"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
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
         Height          =   495
         Left            =   3070
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4430
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frmVendido 
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   5760
      Width           =   9180
      Begin VB.TextBox TxtTCambio 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   26
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox chVendido 
         Caption         =   "Vendido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin Sicmact.TxtBuscar txtCodPersona 
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Enabled         =   0   'False
         Appearance      =   0
         TipoBusqueda    =   3
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin MSMask.MaskEdBox TxFecVenta 
         Height          =   345
         Left            =   5040
         TabIndex        =   18
         Top             =   1440
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "T.Cambio  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha       :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblNombreComprador 
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   16
         Top             =   720
         Width           =   5295
      End
      Begin VB.Label lblComprador 
         Caption         =   "Comprador :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame frmDatos 
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   9180
      Begin VB.ComboBox cboTipoValorizacion 
         Height          =   315
         Left            =   1800
         TabIndex        =   28
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox cboPeriodo 
         Height          =   315
         Left            =   1800
         TabIndex        =   23
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txDesc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   4935
      End
      Begin VB.CommandButton cmdAAgregar 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtMonto 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4920
         TabIndex        =   7
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Valorización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo       :  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblLugar 
         Caption         =   "Descripcion :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lbVaolr 
         Caption         =   "Valor     :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame frmFEdit 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   9180
      Begin Sicmact.FlexEdit FERemates 
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3836
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Suma-DesTipo-DesValor-V.Comer.-V.Reali.-Comprador-FechaCompra-Periodo"
         EncabezadosAnchos=   "400-600-1800-2000-1000-1000-3200-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-L-R-R-R-R-C-C"
         FormatosEdit    =   "0-3-0-4-4-4-0-0-3"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame FraBuscaPers 
      Height          =   1275
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9180
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Aplicar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   2
         ToolTipText     =   "Pulse este Boton para Mostrar los Datos de la Garantia"
         Top             =   675
         Width           =   1425
      End
      Begin VB.CommandButton CmdBuscaPersona 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6945
         TabIndex        =   1
         ToolTipText     =   "Busca Documentos de Persona"
         Top             =   225
         Width           =   1440
      End
      Begin MSComctlLib.ListView LstGaratias 
         Height          =   975
         Left            =   90
         TabIndex        =   3
         Top             =   165
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   1720
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuenta"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Garantia"
            Object.Width           =   5468
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "cNombre"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "nomemi"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "tipodoc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "cnumdoc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "cCtaCod"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "nEstado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "dFechaAd"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "nEstadoAd"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "cUsuarioAd"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmLogRetasacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pgcCtaCod As String

Dim bEstadoCargando As Boolean
Dim cmdEjecutar As Integer
Enum TGarantiaTipoInicio
    RegistroGarantia = 1
    MantenimientoGarantia = 2
    ConsultaGarant = 3
End Enum
Dim vTipoInicio As TGarantiaTipoInicio
Dim sNumgarant As String
Dim sCtaCod As String
Dim nEstadoA As Integer
Dim bCarga As Boolean
Dim bAsignadoACredito As Boolean

Dim bCreditoCF As Boolean
Dim bValdiCCF As Boolean

Dim gcPermiteModificar As Boolean
Dim lcGar As String
Dim MatrixGarantias() As String
Dim nPos As Integer
Dim nDat As Integer
Dim nEstadoAdju As Integer
Dim dEstadoAdju As Date
Dim nEstado As Integer
Dim cUsuariAdju As String
Dim nVendido As Integer
Dim nTpoCambio As Double
Dim cPropietario As String

Private Sub chVendido_Click()
 If chVendido.value = 1 Then
        txtCodPersona.Enabled = True
    Else
        txtCodPersona.Enabled = False
    End If
End Sub

Private Sub cmdAAgregar_Click()
  Dim J As Integer
    Dim i As Integer
    Dim NCaDAr As Integer
    NCaDAr = 0
    Dim nMoneda As Integer
    nMoneda = 0
   
If Val(txtMonto.Text) = "0" Or Trim(cboTipoValorizacion.Text) = "" Then
        If Left(Trim(cboTipoValorizacion.Text), 1) <> "1" Then
            MsgBox "Ingrese Datos Correctamente...", vbInformation, "Aviso"
            Exit Sub
        End If
End If
If nDat = 1 Then
For i = 0 To nPos
    If Trim(MatrixGarantias(2, i)) = Right(cboPeriodo.Text, 4) And Trim(MatrixGarantias(8, i)) = Trim(Right(cboTipoValorizacion.Text, 4)) Then
        MsgBox "Este dato ya fue registrado...", vbInformation, "Aviso"
        txtMonto.Text = 0#
        Exit Sub
    End If
Next i
End If
'20080811******************
    nDat = 1
    FERemates.AdicionaFila
     If FERemates.Row = 1 Then
        ReDim MatrixGarantias(1 To 12, 0 To 0)
     End If
     nPos = FERemates.Row - 1
     MatrixGarantias(1, nPos) = FERemates.Row
     ReDim Preserve MatrixGarantias(1 To 12, 0 To UBound(MatrixGarantias, 2) + 1)
     'If nPos >= 1 Then
        If Mid(Trim(Right(cboTipoValorizacion.Text, 4)), 1, 1) = 1 Then
            MatrixGarantias(1, nPos) = 1
            'MatrixGarantias(5, nPos) = gnTipCambio * CDbl(txtMonto.Text)
            MatrixGarantias(5, nPos) = 0.8 * CDbl(txtMonto.Text)
        Else
            MatrixGarantias(1, nPos) = 0
            MatrixGarantias(5, nPos) = ""
        End If
        MatrixGarantias(2, nPos) = Right(cboPeriodo.Text, 4)
        MatrixGarantias(3, nPos) = txDesc.Text
        MatrixGarantias(4, nPos) = txtMonto.Text
        MatrixGarantias(6, nPos) = ""
        MatrixGarantias(7, nPos) = ""
        MatrixGarantias(8, nPos) = Trim(Right(cboTipoValorizacion.Text, 4))
        MatrixGarantias(9, nPos) = gnTipCambio
        MatrixGarantias(10, nPos) = 0
        MatrixGarantias(11, nPos) = gsCodPersUser
        MatrixGarantias(12, nPos) = Trim(Left(cboTipoValorizacion.Text, 100))
   ' End If

    For i = 0 To nPos
        FERemates.EliminaFila (1)
    Next i
    For i = 0 To nPos
        FERemates.AdicionaFila
        FERemates.TextMatrix(FERemates.Row, 1) = MatrixGarantias(1, i)
        FERemates.TextMatrix(FERemates.Row, 2) = MatrixGarantias(12, i)
        FERemates.TextMatrix(FERemates.Row, 3) = MatrixGarantias(3, i)
        FERemates.TextMatrix(FERemates.Row, 4) = MatrixGarantias(4, i)
        FERemates.TextMatrix(FERemates.Row, 5) = MatrixGarantias(5, i)
        FERemates.TextMatrix(FERemates.Row, 6) = MatrixGarantias(6, i)
        FERemates.TextMatrix(FERemates.Row, 7) = MatrixGarantias(7, i)
        FERemates.TextMatrix(FERemates.Row, 8) = MatrixGarantias(2, i)
        NCaDAr = 1
    Next
    txtMonto.Text = "0.00"
    CmdEliminar.Enabled = True
End Sub

Private Sub CmdBuscaPersona_Click()
 Call cmdCancelarInicio
    ObtieneDocumPersona
    If vTipoInicio = ConsultaGarant Then
        CmdEliminar.Enabled = False
    End If
End Sub
Private Sub ObtieneDocumPersona()
Dim oGaran As DGarantia
Dim R As ADODB.Recordset
'Dim oPers As COMDPersona.UCOMPersona
Dim oPers As UPersona
Dim L As ListItem
    
    LstGaratias.ListItems.Clear
    Set oPers = New UPersona
    Set oPers = frmBuscaPersona.Inicio
    Set oGaran = New DGarantia
    
    If oPers Is Nothing Then
        Exit Sub
    End If
    Set R = oGaran.RecuperaGarantiasPersonaLogistica(oPers.sPersCod, True, False, True)
    Set oGaran = Nothing
    If R.RecordCount > 0 Then
        Me.Caption = "Garantias de Cliente : " & oPers.sPersNombre
    End If
    LstGaratias.ListItems.Clear
    Set oPers = Nothing
    Do While Not R.EOF
        Set L = LstGaratias.ListItems.Add(, , IIf(IsNull(R!cCtaCod), "", R!cCtaCod))
               
        L.SubItems(1) = Trim(R!cDescripcion)
        L.Bold = True
        If R!nMoneda = gMonedaExtranjera Then
            L.ForeColor = RGB(0, 125, 0)
        Else
            L.ForeColor = vbBlack
        End If
        L.SubItems(2) = Trim(R!cNumGarant)
        L.SubItems(3) = Trim(R!cPersCodEmisor)
        L.SubItems(4) = PstaNombre(R!cPersNombre)
        L.SubItems(5) = Trim(R!cTpoDoc)
        L.SubItems(6) = Trim(R!cNroDoc)
        L.SubItems(7) = Trim(R!cCtaCod)
        L.SubItems(8) = Trim(R!nEstadoAdju)
        L.SubItems(9) = R!dEstadoAdju
        L.SubItems(10) = R!nEstado
        L.SubItems(11) = Trim(R!cUsuariAdju)
        nEstadoAdju = Trim(R!nEstadoAdju)
        dEstadoAdju = R!dEstadoAdju
        nEstado = R!nEstado
        cUsuariAdju = R!cUsuariAdju
        R.MoveNext
    Loop
End Sub
Private Sub cmdCancelarInicio()
    Call LimpiaPantalla
    cmdEjecutar = -1
End Sub
Private Sub LimpiaPantalla()
    bCarga = True
    Call LimpiaControles(Me)
    Call InicializaCombos(Me)
    CmdEliminar.Enabled = False
    bCarga = False
End Sub

Private Sub cmdBuscar_Click()
 bAsignadoACredito = False
    
    If Me.LstGaratias.ListItems.Count = 0 Then
        MsgBox "No Existe Garantia que Mostrar ", vbInformation, "Aviso"
        Exit Sub
    End If
    
    sNumgarant = Trim(Me.LstGaratias.SelectedItem.SubItems(2))
    sCtaCod = Trim(Me.LstGaratias.SelectedItem.SubItems(7))
    Call ObtenerArreglo(sNumgarant, sCtaCod)
    nEstadoA = CInt(Trim(Me.LstGaratias.SelectedItem.SubItems(10)))
    nEstadoAdju = CInt(Trim(Me.LstGaratias.SelectedItem.SubItems(8)))
    If vTipoInicio = ConsultaGarant Then
         CmdEliminar.Enabled = False
    End If
 
End Sub
Public Sub ObtenerArreglo(ByVal sNumGarantia As String, ByVal sCodCta As String)
    Dim oGaran As DGarantia
    Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    Set oGaran = New DGarantia
    
    Dim i As Integer
    
    If Trim(cboPeriodo.Text) = "" Then
        MsgBox "Seleccionar periodo"
    Exit Sub
    End If
    
''    If gnTipCambio = 0 Then
''            GetTipCambio gdFecSis, Not gbBitCentral
''       End If
    TxtTCambio.Enabled = True
    TxtTCambio = Format(gnTipCambio, "##,###,##0.0000")
    
    For i = 0 To nPos
        FERemates.EliminaFila (1)
    Next i
    nPos = 0
    
    Set R = oGaran.RecuperaDatosGarantiaLogistica(sNumGarantia, CInt(Right(Trim(cboPeriodo.Text), 4)))
    If R.RecordCount > 0 Then
    nEstadoAdju = 0
            If Not R.EOF And Not R.BOF Then
                R.MoveFirst
            End If
    Do Until R.EOF
        FERemates.AdicionaFila
        nPos = FERemates.Row - 1
       ' MatrixHojaEval(1, nPos) = FEHojaEval.Row
        ReDim Preserve MatrixGarantias(1 To 12, 0 To nPos + 1)
        FERemates.AdicionaFila
        If Left(R!nTValor, 1) = 1 Then
            MatrixGarantias(1, nPos) = 1
            If R!nVendido = 1 Then
                MatrixGarantias(5, nPos) = R!nValor * 0.8
                nTpoCambio = R!nTipoCambio
            Else
                MatrixGarantias(5, nPos) = R!nValor * 0.8
                nTpoCambio = gnTipCambio
            End If
        Else
            MatrixGarantias(1, nPos) = 0
            MatrixGarantias(5, nPos) = 0
        End If
        
        MatrixGarantias(2, nPos) = R!nperiodo
        MatrixGarantias(3, nPos) = R!cDesValor
        MatrixGarantias(4, nPos) = R!nValor
        
        MatrixGarantias(6, nPos) = R!cDescCodComprador
        MatrixGarantias(7, nPos) = Format(R!dFechaCompra, "YYYY/MM/DD")
        MatrixGarantias(8, nPos) = R!nTValor
        
        MatrixGarantias(10, nPos) = R!nVendido
        MatrixGarantias(11, nPos) = R!cPersCodUsuario
        MatrixGarantias(12, nPos) = R!cConsDescripcion
        nEstadoAdju = R!nEstadoAdju
        nVendido = R!nVendido
                
        FERemates.TextMatrix(FERemates.Row, 1) = MatrixGarantias(1, nPos)
        FERemates.TextMatrix(FERemates.Row, 2) = MatrixGarantias(12, nPos)
        FERemates.TextMatrix(FERemates.Row, 3) = MatrixGarantias(3, nPos)
        FERemates.TextMatrix(FERemates.Row, 4) = MatrixGarantias(4, nPos)
        FERemates.TextMatrix(FERemates.Row, 5) = MatrixGarantias(5, nPos)
        FERemates.TextMatrix(FERemates.Row, 6) = MatrixGarantias(6, nPos)
        FERemates.TextMatrix(FERemates.Row, 7) = MatrixGarantias(7, nPos)
        FERemates.TextMatrix(FERemates.Row, 8) = MatrixGarantias(2, nPos)
        CmdEliminar.Enabled = True
        
        R.MoveNext
        
    Loop
    End If
    If nEstadoAdju = 10 Then
        CmdEliminar.Enabled = False
        cmdGrabar.Enabled = False
        cmdVender.Enabled = False
    Else
        CmdEliminar.Enabled = True
        cmdGrabar.Enabled = True
        cmdVender.Enabled = True
    End If
End Sub

Private Sub CmdEliminar_Click()
 Dim nXPos As Integer
    nXPos = FERemates.Row
    If nPos >= 1 Then
    If MsgBox("Esta Seguro de Eliminar este registro.", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        FERemates.EliminaFila (FERemates.Row)
        If nPos >= 1 Then
            Dim J As Integer
            For J = nXPos - 1 To nPos
                MatrixGarantias(1, J) = MatrixGarantias(1, J + 1)
                MatrixGarantias(2, J) = MatrixGarantias(2, J + 1)
                MatrixGarantias(3, J) = MatrixGarantias(3, J + 1)
                MatrixGarantias(4, J) = MatrixGarantias(4, J + 1)
                MatrixGarantias(5, J) = MatrixGarantias(5, J + 1)
                MatrixGarantias(6, J) = MatrixGarantias(6, J + 1)
                MatrixGarantias(7, J) = MatrixGarantias(7, J + 1)
                MatrixGarantias(8, J) = MatrixGarantias(8, J + 1)
                MatrixGarantias(9, J) = MatrixGarantias(9, J + 1)
                MatrixGarantias(10, J) = MatrixGarantias(10, J + 1)
                MatrixGarantias(11, J) = MatrixGarantias(11, J + 1)
            Next J
            nPos = nPos - 1
        Else
            nPos = nPos - 1
            nDat = 0
        End If
    End If
    Else
        If FERemates.Row >= 1 Then
        FERemates.EliminaFila (1)
        End If
        nPos = -1
        nDat = 0
    End If
End Sub
'*******
Private Sub CmdGrabar_Click()
Dim i As Integer
Dim J As Integer
Dim oGaran As NGarantia
Set oGaran = New NGarantia
Dim nCont As Integer
If nEstadoAdju >= 7 Then
    For i = 0 To nPos
        nCont = nCont + 1
        Call oGaran.InsertarGarantiaLogistica(sNumgarant, CInt(MatrixGarantias(2, i)), CInt(MatrixGarantias(8, i)), CDbl(MatrixGarantias(4, i)), MatrixGarantias(3, i), 1, nEstadoAdju, MatrixGarantias(11, i), i, gdFecSis, gsCodUser, gsCodAge)
    Next i
    If nCont > 0 Then
         MsgBox "Datos se registraron correctamente...", vbInformation, "Aviso"
    Else
        MsgBox "Ingrese Datos Correctamente...", vbInformation, "Aviso"
    End If
Else
    MsgBox "Verificar datos...", vbInformation, "Aviso"
End If
End Sub
Private Sub Impresion()
Dim oPrev As Previo.clsPrevio
Dim sCad As String
Dim i As Integer
Dim sPropietario As String
Dim loImpre As NGarantia
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
   With rs
        'Crear RecordSet****
        .Fields.Append "nSuma", adInteger
        .Fields.Append "nPeriodo", adInteger
        .Fields.Append "cDesTipo", adVarChar, 50
        .Fields.Append "cDescripcion", adVarChar, 100
        .Fields.Append "nMonto", adCurrency
        .Fields.Append "nMontoRea", adCurrency
        .Fields.Append "nTipo", adInteger
        .Open
        'Llenar Recordset****
        For i = 1 To FERemates.Rows - 1
            FERemates.Row = i
            FERemates.Col = 1
                .AddNew
                .Fields("nSuma") = IIf(IsNull(FERemates.TextMatrix(i, 1)), "", FERemates.TextMatrix(i, 1))
                .Fields("nPeriodo") = IIf(IsNull(FERemates.TextMatrix(i, 8)), "", FERemates.TextMatrix(i, 8))
                .Fields("cDesTipo") = IIf(IsNull(FERemates.TextMatrix(i, 2)), "", FERemates.TextMatrix(i, 2))
                .Fields("cDescripcion") = IIf(IsNull(FERemates.TextMatrix(i, 3)), "", FERemates.TextMatrix(i, 3))
                .Fields("nMonto") = IIf(IsNull(FERemates.TextMatrix(i, 4)), "", FERemates.TextMatrix(i, 4))
                If IIf(IsNull(FERemates.TextMatrix(i, 1)), "", FERemates.TextMatrix(i, 1)) = 1 Then
                .Fields("nMontoRea") = IIf(IsNull(FERemates.TextMatrix(i, 5)), 0, FERemates.TextMatrix(i, 5))
                Else
                .Fields("nMontoRea") = 0
                End If
                sPropietario = IIf(IsNull(FERemates.TextMatrix(i, 6)), 0, FERemates.TextMatrix(i, 6))
                .Fields("nTipo") = MatrixGarantias(8, i - 1)
        Next i
    End With
    Set loImpre = New NGarantia
        sCad = loImpre.ImpresionRetasacionGarantia(rs, gsNomCmac, gdFecSis, gsNomAge, gsCodUser, gImpresora, nTpoCambio, nVendido, sPropietario)
    Set loImpre = Nothing
    rs.Close
    Set oPrev = New Previo.clsPrevio
    oPrev.Show sCad, "Transferencia A Recuperaciones"
    Set oPrev = Nothing
    
    
End Sub

Private Sub CmdImprimir_Click()
Call Impresion
End Sub
'**
Private Sub cmdImprimirTodos_Click()
Dim oGaran As DGarantia
Dim R As ADODB.Recordset
    Dim lsNombreArchivo As String
    Dim lMatCabecera As Variant
    Dim lsmensaje As String

    lsNombreArchivo = "reporteGarantiasAdjudicadas"

    ReDim lMatCabecera(14, 2)

    lMatCabecera(0, 0) = "Cod.Cliente"
    lMatCabecera(1, 0) = "Cliente"
    lMatCabecera(2, 0) = "Ncredito"
    lMatCabecera(3, 0) = "NroGarantia"
    lMatCabecera(4, 0) = "DesGarantia"
    lMatCabecera(5, 0) = "Dirección"
    lMatCabecera(6, 0) = "Monto Saneado"
    lMatCabecera(7, 0) = "Valor Comercial"
    lMatCabecera(8, 0) = "Valor Realización"
    lMatCabecera(9, 0) = "cfrente"
    lMatCabecera(10, 0) = "cDerecha"
    lMatCabecera(11, 0) = "cIzquierda"
    lMatCabecera(12, 0) = "cFondo"
    lMatCabecera(13, 0) = "Estado"
    
        
    Set oGaran = New DGarantia
    Set R = oGaran.ReporteGarantiasLogisticaxPeriodo(gdFecSis)
    Set oGaran = Nothing

    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Garantias Adjudicados", " Al " & gdFecSis, lsNombreArchivo, lMatCabecera, R, 2, , , True)
 
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdVender_Click()
Dim i As Integer
Dim oGaran As NGarantia
Set oGaran = New NGarantia
Dim nCont As Integer
If chVendido.value = 1 Then
        If txtCodPersona.Text = "" Or Val(TxtTCambio.Text) = "0" Or Trim(TxFecVenta.Text) = "__/__/____" Then
                MsgBox "Ingrese Datos Correctamente...", vbInformation, "Aviso"
                Exit Sub
        End If
        
        For i = 0 To nPos
            nCont = nCont + 1
        Next i
        If nCont > 0 Then
             Call oGaran.ActualizargarantiasxVentaLogistica(sNumgarant, 10, gdFecSis, gsCodUser, CDate(TxFecVenta.Text), 1, Val(TxtTCambio.Text), txtCodPersona.Text, 1, gsCodAge)
             MsgBox "Datos se registraron correctamente...", vbInformation, "Aviso"
        Else
            MsgBox "Ingrese Datos Correctamente...", vbInformation, "Aviso"
        End If
        chVendido.Enabled = False
        cmdVender.Enabled = False
Else
        MsgBox "Debe activar Vendido ", vbInformation, "Aviso"
End If
End Sub

Private Sub Form_Load()

Dim Conn As DConecta
Dim oCons As DConstante
Dim Res As ADODB.Recordset

Set Res = New ADODB.Recordset
Set oCons = New DConstante
Set Conn = New DConecta

Dim i As Integer

nVendido = 0
nTpoCambio = 0

If gnTipCambio = 0 Then
        GetTipCambio gdFecSis, Not gbBitCentral
End If

For i = CInt(Format(gdFecSis, "YYYY")) - 10 To CInt(Format(gdFecSis, "YYYY"))
    cboPeriodo.AddItem i & Space(200) & Trim(i)
Next i
 
Conn.AbreConexion

Set Res = oCons.RecuperaConstantes(9074)
Call Llenar_Combo_con_Recordset(Res, cboTipoValorizacion)
Set Res = Nothing
Conn.CierraConexion

Set Conn = Nothing


End Sub

Private Sub txtCodPersona_EmiteDatos()
    Me.lblNombreComprador.Caption = Trim(txtCodPersona.psDescripcion)
End Sub
Sub Llenar_Combo_con_Recordset(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cConsDescripcion) & Space(100) & Trim(Str(pRs!nConsValor))
    pRs.MoveNext
Loop
pRs.Close
End Sub
