VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form frmLogContReajAdenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Titulo"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385
   Icon            =   "frmLogContReajAdenda.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstReajuste 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos de Reajuste"
      TabPicture(0)   =   "frmLogContReajAdenda.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblNombreArchivo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dtpFecHasta"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtNContrato"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtProveedor"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtNReajuste"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboTpoReajuste"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "dtpFecDesde"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtTpoMoneda"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtMonto"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdExaminar"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "chbreajuste"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Bienes del Contrato"
      TabPicture(1)   =   "frmLogContReajAdenda.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraItemContrato"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraItemContrato 
         Caption         =   "Bienes relacionados al contrato"
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
         Height          =   3375
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   7935
         Begin VB.CommandButton cmdQuitarItemCont 
            Caption         =   "Quitar"
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
            Left            =   1440
            TabIndex        =   29
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CommandButton cmdAgregarItemCont 
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
            Height          =   375
            Left            =   240
            TabIndex        =   28
            Top             =   2880
            Width           =   1215
         End
         Begin Sicmact.FlexEdit feOrden 
            Height          =   2535
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   4471
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Ag.Des.-Objeto-Descripcion-Solic.-P.Unitario-SubTotal-CtaContCod"
            EncabezadosAnchos=   "0-800-900-3000-700-1100-1100-0"
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
            ColumnasAEditar =   "X-1-2-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-1-1-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-L-R-R-R-L"
            FormatosEdit    =   "0-0-0-0-3-2-2-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.CheckBox chbreajuste 
         Height          =   255
         Left            =   -72600
         TabIndex        =   18
         Top             =   3120
         Width           =   255
      End
      Begin VB.CommandButton cmdExaminar 
         Appearance      =   0  'Flat
         Caption         =   "E&xaminar"
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
         Left            =   -68040
         TabIndex        =   22
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -72960
         TabIndex        =   20
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtTpoMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73560
         TabIndex        =   19
         Top             =   2640
         Width           =   495
      End
      Begin MSComCtl2.DTPicker dtpFecDesde 
         Height          =   375
         Left            =   -70920
         TabIndex        =   14
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   74186753
         CurrentDate     =   41876
      End
      Begin VB.ComboBox cboTpoReajuste 
         Height          =   315
         Left            =   -70920
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtNReajuste 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -72840
         TabIndex        =   8
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtProveedor 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -72840
         TabIndex        =   5
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox txtNContrato 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -72840
         TabIndex        =   3
         Top             =   840
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker dtpFecHasta 
         Height          =   375
         Left            =   -68760
         TabIndex        =   16
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   74186753
         CurrentDate     =   41876
      End
      Begin VB.Label lblNombreArchivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -72840
         TabIndex        =   25
         Tag             =   "txtnombre"
         Top             =   3480
         Width           =   4695
      End
      Begin VB.Label Label12 
         Caption         =   "Reajuste Digital"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74760
         TabIndex        =   21
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Archivo Reajuste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -69480
         TabIndex        =   15
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -71640
         TabIndex        =   13
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74280
         TabIndex        =   12
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74760
         TabIndex        =   11
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -71760
         TabIndex        =   9
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Nº Reajuste:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74280
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Reajuste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74760
         TabIndex        =   6
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74280
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Contrato :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74280
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   7200
      TabIndex        =   24
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "&Registrar"
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
      Left            =   6120
      TabIndex        =   23
      Top             =   4200
      Width           =   1095
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgFile 
      Height          =   495
      Left            =   7680
      TabIndex        =   26
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Filtro          =   ""
      Altura          =   0
   End
End
Attribute VB_Name = "frmLogContReajAdenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fsNContrato As String
Dim fnContRef As Integer
Dim fnAdenda As Integer
Dim fntpodocorigen As Integer
Dim psRutaContrato As String
Dim pbActivaArchivo As Boolean
Dim fsPathFile As String
Dim fsNomFile As String
Dim fsRuta As String
Dim Datoscontrato() As TContratoBS
Dim fRsAgencia As New ADODB.Recordset
Dim fRsServicio As New ADODB.Recordset
Dim fRsCompra As New ADODB.Recordset
Dim fnTpoCambio As Currency
Dim rsLog As ADODB.Recordset
'Dim sObjCod    As String, sObjDesc As String, sObjUnid As String
'Dim sCtaCod    As String, sCtaDesc Asf String
Public Sub Inicio(ByVal psNContrato As String, ByVal pnContRef As Integer, ByVal pnTipoCont As Integer, Optional ByVal pnNAdenda As Integer = 0)
    fsNContrato = psNContrato
    fnContRef = pnContRef
    fnAdenda = pnNAdenda
    fntpodocorigen = pnTipoCont
    Me.Show 1
End Sub
Private Sub CargaVariables()
    Dim oArea As New DActualizaDatosArea
    Dim oalmacen As New DLogAlmacen
    Dim rs As New ADODB.Recordset
    
    If gbBitTCPonderado Then
        fnTpoCambio = gnTipCambioPonderado
    Else
        fnTpoCambio = gnTipCambioC
    End If
    
    Set fRsAgencia = oArea.GetAgencias(, , True)
    Set fRsCompra = oalmacen.GetBienesAlmacen(, "11','12','13")
    'Set fRsServicio = OrdenServicio()

    Set rs = Nothing
    Set oalmacen = Nothing
    Set oArea = Nothing
End Sub
Private Sub cboTpoReajuste_Click()
    Dim dLog As DLogGeneral
    Dim rsBS As New ADODB.Recordset
    Dim i As Integer
    Dim row As Integer
    
    LimpiarControles
    If cboTpoReajuste.Text <> "" Then
    If fntpodocorigen <> LogTipoContrato.ContratoObra Then
    Select Case CInt((Right(cboTpoReajuste.Text, 4)))
        Case LogtipoReajusteAdenda.Complementaria
               DesHabilitaControles True, False, False
               'dtpFecDesde.SetFocus
        Case LogtipoReajusteAdenda.Adicional
               DesHabilitaControles True, True, True
'               If fnTpoDocOrigen = LogTipoContrato.ContratoAdqBienes Or fnTpoDocOrigen = LogTipoContrato.ContratoSuministro Then
                    gcOpeCod = IIf(txtTpoMoneda.Text = gcPEN_SIMBOLO, "501215", "502215") 'marg ers044-2016
'               End If
               txtMonto.Text = ""
               txtMonto.SetFocus
               cmdAgregarItemCont.Enabled = True
         Case LogtipoReajusteAdenda.Reduccion
               DesHabilitaControles True, True, True
               txtMonto.Enabled = False
               If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or fntpodocorigen = LogTipoContrato.ContratoSuministro Then
                    gcOpeCod = IIf(txtTpoMoneda.Text = gcPEN_SIMBOLO, "501215", "502215") 'marg ers044-2016
                    i = 0
                    Set dLog = New DLogGeneral
                    Set rsBS = dLog.ListaBienesContrato(fsNContrato, fnContRef)
                    If Not rsBS.EOF Then
                        Do While Not rsBS.EOF
                            i = i + 1
                            ReDim Preserve Datoscontrato(i)
                            feOrden.AdicionaFila
                            row = feOrden.row
                            Datoscontrato(i).sAgeCod = rsBS!cAgeDest
                            feOrden.TextMatrix(row, 1) = rsBS!cAgeDest
                            Datoscontrato(i).sObjeto = rsBS!cBSCod
                            feOrden.TextMatrix(row, 2) = rsBS!cBSCod
                            Datoscontrato(i).sDescripcion = rsBS!cBSDescripcion
                             feOrden.TextMatrix(row, 3) = rsBS!cBSDescripcion
                            Datoscontrato(i).nCantidad = rsBS!nCant
                            feOrden.TextMatrix(row, 4) = rsBS!nCant
                            Datoscontrato(i).nPrecUnit = rsBS!PrecUnit
                            feOrden.TextMatrix(row, 5) = rsBS!PrecUnit
                            Datoscontrato(i).nTotal = rsBS!nMovImporte
                            feOrden.TextMatrix(row, 6) = rsBS!nMovImporte
                            Datoscontrato(i).sCtaContCod = rsBS!cCtaContCod
                            feOrden.TextMatrix(row, 7) = rsBS!cCtaContCod
                            Datoscontrato(i).nMovItem = rsBS!nMovItem
                            rsBS.MoveNext
                        Loop
                    End If
                    feOrden.ColumnasAEditar = "X-X-X-X-4-5-X-X"
               End If
               Set rsBS = Nothing
               txtMonto.Text = ""
               cmdAgregarItemCont.Enabled = False
    End Select
    Else
        Select Case CInt((Right(cboTpoReajuste.Text, 4)))
            Case LogtipoReajusteAdenda.Complementaria
               DesHabilitaControles True, False, False
               dtpFecDesde.SetFocus
            Case LogtipoReajusteAdenda.Adicional
                DesHabilitaControles True, True, True
                txtMonto.Text = ""
               txtMonto.SetFocus
            Case LogtipoReajusteAdenda.Reduccion
                DesHabilitaControles True, True, True
                txtMonto.Text = ""
                txtMonto.SetFocus
        End Select
    End If
    End If
End Sub
Private Sub chbreajuste_Click()
    cmdExaminar.Enabled = IIf(chbreajuste.value = 1, True, False)
End Sub
Private Sub cmdAgregarItemCont_Click()
If Not validaBusqueda Then Exit Sub
    If feOrden.TextMatrix(1, 0) <> "" Then
        If Not validaIngresoRegistros Then Exit Sub
    End If
    feOrden.AdicionaFila
'    If fnTpoDocOrigen = LogTipoContrato.ContratoAdqBienes Or _
'       fnTpoDocOrigen = LogTipoContrato.ContratoSuministro Then
        feOrden.ColumnasAEditar = "X-1-2-X-4-5-X-X"
        feOrden.TextMatrix(feOrden.row, 4) = "0"
        feOrden.TextMatrix(feOrden.row, 5) = "0.00"
        feOrden.TextMatrix(feOrden.row, 6) = "0.00"
'    ElseIf fnTpoDocOrigen = LogTipoContrato.ContratoServicio Then
'        feOrden.ColumnasAEditar = "X-1-2-X-X-X-6-X"
'    End If
    feOrden.TextMatrix(feOrden.row, 6) = "0.00"
    feOrden.col = 2
    feOrden.SetFocus
    feOrden_RowColChange
End Sub
Private Function validaIngresoRegistros() As Boolean 'PASI20140110
    Dim i As Long, j As Long
    Dim col As Integer
    Dim Columnas() As String
    Dim lsColumnas As String
    
    lsColumnas = "1,2,6"
    Columnas = Split(lsColumnas, ",")
        
    validaIngresoRegistros = True
    If feOrden.TextMatrix(1, 0) <> "" Then
        For i = 1 To feOrden.Rows - 1
            For j = 1 To feOrden.Cols - 1
                For col = 0 To UBound(Columnas)
                    If j = Columnas(col) Then
                        If Len(Trim(feOrden.TextMatrix(i, j))) = 0 And feOrden.ColWidth(j) <> 0 Then
                            MsgBox "Ud. debe especificar el campo " & feOrden.TextMatrix(0, j), vbInformation, "Aviso"
                            validaIngresoRegistros = False
                            feOrden.TopRow = i
                            feOrden.row = i
                            feOrden.col = j
                            feOrden_RowColChange
                            Exit Function
                        End If
                    End If
                Next
            Next
            If IsNumeric(feOrden.TextMatrix(i, 6)) Then
                If CCur(feOrden.TextMatrix(i, 6)) <= 0 Then
                    MsgBox "El Importe Total debe ser mayor a cero", vbInformation, "Aviso"
                    validaIngresoRegistros = False
                    feOrden.TopRow = i
                    feOrden.row = i
                    feOrden.col = 6
                    Exit Function
                End If
            Else
                MsgBox "El Importe Total debe ser númerico", vbInformation, "Aviso"
                validaIngresoRegistros = False
                feOrden.TopRow = i
                feOrden.row = i
                feOrden.col = 6
                Exit Function
            End If
            If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or _
                fntpodocorigen = LogTipoContrato.ContratoSuministro Then
                If Len(Trim(feOrden.TextMatrix(i, 7))) = 0 Then
                    MsgBox "El Objeto " & feOrden.TextMatrix(i, 3) & Chr(10) & "no tiene configurado Plantilla Contable, consulte con el Dpto. de Contabilidad", vbInformation, "Aviso"
                    feOrden.TopRow = i
                    feOrden.row = i
                    feOrden.col = 2
                    validaIngresoRegistros = False
                    Exit Function
                End If
            End If
        Next
    Else
        MsgBox "Ud. debe agregar los Bienes/Servicios a dar Conformidad", vbInformation, "Aviso"
        validaIngresoRegistros = False
    End If
End Function
Private Function validaBusqueda()
    validaBusqueda = True
    If Len(txtMonto.Text) = 0 Then
        MsgBox "Ud. primero debe Ingresar el Monto.", vbInformation, "Aviso"
        validaBusqueda = False
        Exit Function
    End If
End Function
Private Sub cmdExaminar_Click()
Dim i As Integer
    CdlgFile.nHwd = Me.hwnd
    CdlgFile.Filtro = "Contratos Digital (*.pdf)|*.pdf"
    Me.CdlgFile.Altura = 300
    CdlgFile.Show

    fsPathFile = CdlgFile.Ruta
    fsRuta = fsPathFile
        If fsPathFile <> Empty Then
            For i = Len(fsPathFile) - 1 To 1 Step -1
                    If Mid(fsPathFile, i, 1) = "\" Then
                        fsPathFile = Mid(CdlgFile.Ruta, 1, i)
                        fsNomFile = Mid(CdlgFile.Ruta, i + 1, Len(CdlgFile.Ruta) - i)
                        Exit For
                    End If
             Next i
          Screen.MousePointer = 11
          
            If pbActivaArchivo Then
                lblNombreArchivo.Caption = UCase(Trim(Me.txtNContrato.Text)) & ".pdf"
            Else
                lblNombreArchivo.Caption = ""
            End If
            Me.cmdRegistrar.SetFocus
        Else
           MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
           Exit Sub
        End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdExaminar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdExaminar_Click
    End If
    If KeyAscii = 9 Then
        Me.cmdRegistrar.SetFocus
    End If
End Sub

Private Sub cmdQuitarItemCont_Click()
    feOrden.EliminaFila feOrden.row
End Sub
Private Sub cmdRegistrar_Click()
    On Error GoTo ErrorRegistrarAdenda
    Dim oLog As New DLogGeneral
    Dim bTrans As Boolean
    Dim lsNContrato As String
    Dim lsNAdenda As Integer
    Dim lnUltItem As Integer
    Dim Datoscontrato() As TContratoBS
    Dim Index As Integer
    Dim lnMovItem As Integer
    Dim lnImporte As Currency
    Dim j As Integer
    
    If Not ValidaAdenda Then Exit Sub
    
    lsNContrato = Trim(txtNContrato.Text)
    lnNAdenda = CInt(Trim(txtNReajuste.Text))
    
    If MsgBox("Esta seguro de grabar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Set oLog = New DLogGeneral
    oLog.dBeginTrans
    bTrans = True
    
    If chbreajuste.value = 1 Then
        GrabarArchivo
    End If
    If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or _
    fntpodocorigen = LogTipoContrato.ContratoSuministro Then
        If Trim(Right(cboTpoReajuste.Text, 100)) <> LogtipoReajusteAdenda.Complementaria Then
            For Index = 1 To feOrden.Rows - 1
                ReDim Preserve Datoscontrato(Index)
                Datoscontrato(Index).sAgeCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 1))))
                If fntpodocorigen = LogTipoContrato.ContratoAdqBienes _
                    Or fntpodocorigen = LogTipoContrato.ContratoSuministro Then
                    Datoscontrato(Index).sCtaContCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 7))))
                End If
                Datoscontrato(Index).sObjeto = CStr(Trim(feOrden.TextMatrix(Index, 2)))
                Datoscontrato(Index).sDescripcion = CStr(Trim(feOrden.TextMatrix(Index, 3)))
                Datoscontrato(Index).nCantidad = Val(feOrden.TextMatrix(Index, 4))
                Datoscontrato(Index).nTotal = feOrden.TextMatrix(Index, 6)
            Next
        End If
    End If
    oLog.ActualizarContratoProveedorNew lsNContrato, fnContRef, lnNAdenda
        If fntpodocorigen = LogTipoContrato.ContratoObra Then
            Select Case Trim(Right(cboTpoReajuste.Text, 100))
                Case LogtipoReajusteAdenda.Complementaria
                     oLog.RegistraContratoAdendaReajuste lsNContrato, fnContRef, lnNAdenda, Trim(Right(cboTpoReajuste.Text, 100)), IIf(txtTpoMoneda.Text = gcPEN_SIMBOLO, 1, 2), 0, CDate(dtpFecDesde.value), CDate(dtpFecHasta.value), Trim(lblNombreArchivo.Caption) 'marg ers044-2016
                Case LogtipoReajusteAdenda.Adicional
                    If CInt(txtNReajuste.Text) = 1 Then
                        oLog.RegistraContratoAdendaReajuste lsNContrato, fnContRef, lnNAdenda, Trim(Right(cboTpoReajuste.Text, 100)), IIf(txtTpoMoneda.Text = gcPEN_SIMBOLO, 1, 2), Trim(txtMonto.Text), CDate(dtpFecDesde.value), CDate(dtpFecHasta.value), Trim(lblNombreArchivo.Caption) 'marg ers044-2016
                    Else
                        oLog.MigraContratoAdendaReajuste lsNContrato, fnContRef, lnNAdenda, Trim(Right(cboTpoReajuste.Text, 100))
                        oLog.ActualizaMontoAdendaReajuste lsNContrato, fnContRef, lnNAdenda, Trim(txtMonto.Text)
                    End If
                    oLog.ActualizaSaldoContrato lsNContrato, fnContRef, Trim(txtMonto.Text), lnNAdenda
                Case LogtipoReajusteAdenda.Reduccion
                    If CInt(txtNReajuste.Text) = 1 Then
                        oLog.RegistraContratoAdendaReajuste lsNContrato, fnContRef, lnNAdenda, Trim(Right(cboTpoReajuste.Text, 100)), IIf(txtTpoMoneda.Text = gcPEN_SIMBOLO, 1, 2), Trim(txtMonto.Text), CDate(dtpFecDesde.value), CDate(dtpFecHasta.value), Trim(lblNombreArchivo.Caption) 'marg ers044-2016
                    Else
                        oLog.MigraContratoAdendaReajuste lsNContrato, fnContRef, lnNAdenda, Trim(Right(cboTpoReajuste.Text, 100))
                        oLog.ActualizaMontoAdendaReajuste lsNContrato, fnContRef, lnNAdenda, Trim(txtMonto.Text) * -1
                    End If
                    oLog.ActualizaSaldoContrato lsNContrato, fnContRef, Trim(txtMonto.Text) * -1, lnNAdenda
            End Select
        End If
        If (fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or fntpodocorigen = LogTipoContrato.ContratoSuministro) Then
            Select Case Trim(Right(cboTpoReajuste.Text, 100))
                Case LogtipoReajusteAdenda.Complementaria
                    oLog.RegistraContratoAdendaReajuste lsNContrato, fnContRef, lnNAdenda, Trim(Right(cboTpoReajuste.Text, 100)), IIf(txtTpoMoneda.Text = gcPEN_SIMBOLO, 1, 2), 0, CDate(dtpFecDesde.value), CDate(dtpFecHasta.value), Trim(lblNombreArchivo.Caption) 'marg ers044-2016
                    lnUltItem = oLog.MigrarBienesContratoxAdenda(lsNContrato, lnNAdenda, fnContRef)
                Case LogtipoReajusteAdenda.Adicional
                    oLog.RegistraContratoAdendaReajuste lsNContrato, fnContRef, lnNAdenda, Trim(Right(cboTpoReajuste.Text, 100)), IIf(txtTpoMoneda.Text = gcPEN_SIMBOLO, 1, 2), Trim(txtMonto.Text), CDate(dtpFecDesde.value), CDate(dtpFecHasta.value), Trim(lblNombreArchivo.Caption) 'marg ers044-2016
                    lnUltItem = oLog.MigrarBienesContratoxAdenda(lsNContrato, lnNAdenda, fnContRef)
                    If lnUltItem <= 0 Then
                        oLog.dRollbackTrans
                        Set oLog = Nothing
                         MsgBox "No se ha podido registrar la Adenda Adicional", vbCritical, "Aviso"
                        Exit Sub
                    End If
                    If UBound(Datoscontrato) Then
                        For lnMovItem = 1 To UBound(Datoscontrato)
                            lnImporte = oLog.RegistrarContratoBienes(Trim(txtNContrato.Text), fnContRef, lnNAdenda, Datoscontrato(lnMovItem).sAgeCod, Datoscontrato(lnMovItem).sCtaContCod, Datoscontrato(lnMovItem).sDescripcion, lnUltItem + lnMovItem, Datoscontrato(lnMovItem).sObjeto, Datoscontrato(lnMovItem).nCantidad, Datoscontrato(lnMovItem).nTotal, 1, "", "")
                            oLog.ActualizaSaldoContrato lsNContrato, fnContRef, lnImporte, lnNAdenda
                        Next lnMovItem
                    End If
                    For j = 1 To feOrden.Rows - 1
                        oLog.RegistrarContratoAdendaBienRel lsNContrato, fnContRef, lnNAdenda, feOrden.TextMatrix(j, 1), feOrden.TextMatrix(j, 7), feOrden.TextMatrix(j, 3), feOrden.TextMatrix(j, 2), feOrden.TextMatrix(j, 4), feOrden.TextMatrix(j, 5), feOrden.TextMatrix(j, 6)
                    Next
                Case LogtipoReajusteAdenda.Reduccion
                If UBound(Datoscontrato) Then
                    lnImporte = 0
                    For lnMovItem = 1 To UBound(Datoscontrato)
                        lnImporte = lnImporte + Datoscontrato(lnMovItem).nTotal
                        oLog.RegistrarContratoBienes lsNContrato, fnContRef, lnNAdenda, Datoscontrato(lnMovItem).sAgeCod, Datoscontrato(lnMovItem).sCtaContCod, Datoscontrato(lnMovItem).sDescripcion, lnMovItem, Datoscontrato(lnMovItem).sObjeto, Datoscontrato(lnMovItem).nCantidad, Datoscontrato(lnMovItem).nTotal, 1, "", ""
                    Next lnMovItem
                    'lnImporte = oLog.ActualizaSaldoContrato(lsNContrato, fnContRef, lnImporte, lnNAdenda)
                    oLog.ActualizaSaldoContrato lsNContrato, fnContRef, lnImporte, lnNAdenda
                    oLog.RegistraContratoAdendaReajuste lsNContrato, fnContRef, lnNAdenda, Trim(Right(cboTpoReajuste.Text, 100)), IIf(txtTpoMoneda.Text = gcPEN_SIMBOLO, 1, 2), lnImporte, CDate(dtpFecDesde.value), CDate(dtpFecHasta.value), Trim(lblNombreArchivo.Caption) 'marg ers044-2016
                    For j = 1 To feOrden.Rows - 1
                        oLog.RegistrarContratoAdendaBienRel lsNContrato, fnContRef, lnNAdenda, feOrden.TextMatrix(j, 1), feOrden.TextMatrix(j, 7), feOrden.TextMatrix(j, 3), feOrden.TextMatrix(j, 2), feOrden.TextMatrix(j, 4), feOrden.TextMatrix(j, 5), feOrden.TextMatrix(j, 6)
                    Next
                End If
                
            End Select
        End If
    oLog.dCommitTrans
    bTrans = False
    Set oLog = Nothing
    MsgBox "Adenda registrada satisfactoriamente", vbInformation, "Aviso"
    Unload Me
    Exit Sub
ErrorRegistrarAdenda:
      If bTrans Then
        oLog.dRollbackTrans
        Set oLog = Nothing
    End If
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Function ValidaAdenda() As Boolean
Dim i As Integer
    If fntpodocorigen = LogTipoContrato.ContratoObra Then
        If cboTpoReajuste.ListIndex = -1 Then
            MsgBox "No se ha Seleccionado el tipo de Adenda.", vbInformation, "Aviso"
            cboTpoReajuste.SetFocus
            ValidaAdenda = False
            Exit Function
        End If
        If Trim(Right(cboTpoReajuste.Text, 100)) = LogtipoReajusteAdenda.Adicional Or Trim(Right(cboTpoReajuste.Text, 100)) = LogtipoReajusteAdenda.Reduccion Then
            If Len(txtMonto.Text) = 0 Then
                MsgBox "No se ha Ingresado el Monto de la Adenda.", vbInformation, "Aviso"
                txtMonto.SetFocus
                ValidaAdenda = False
                Exit Function
            End If
        End If
        If chbreajuste.value = 1 Then
            If Len(lblNombreArchivo.Caption) = 0 Then
                MsgBox "No se ha Seleccionado ningun Archivo (.pdf).", vbInformation, "Aviso"
                cmdExaminar.SetFocus
                ValidaAdenda = False
                Exit Function
            End If
        End If
    End If
    If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or fntpodocorigen = LogTipoContrato.ContratoSuministro Then
         If cboTpoReajuste.ListIndex = -1 Then
            MsgBox "No se ha Seleccionado el tipo de Adenda.", vbInformation, "Aviso"
            cboTpoReajuste.SetFocus
            ValidaAdenda = False
            Exit Function
        End If
        If Trim(Right(cboTpoReajuste.Text, 100)) = LogtipoReajusteAdenda.Adicional Then
            If Len(txtMonto.Text) = 0 Then
                MsgBox "No se ha Ingresado el Monto de la Adenda.", vbInformation, "Aviso"
                txtMonto.SetFocus
                ValidaAdenda = False
                Exit Function
            End If
            If feOrden.TextMatrix(1, 1) = "" Then
                MsgBox "Ud. Debe Asegurarse de haber ingresado correctamente los Bienes/Servicios.", vbInformation, "Aviso"
                ValidaAdenda = False
            Exit Function
            End If
            Dim nMonto As Double
            nMonto = 0
            For i = 1 To feOrden.Rows - 1
                nMonto = nMonto + (feOrden.TextMatrix(i, 6))
            Next i
            If CDbl(txtMonto.Text) <> nMonto Then
                MsgBox "El monto total de los Bienes del Contrato no coincide con el total ingresado, verifique", vbInformation, "Aviso"
                ValidaDatos = False
            Exit Function
        End If
        End If
        If Trim(Right(cboTpoReajuste.Text, 100)) = LogtipoReajusteAdenda.Reduccion Then
            For i = 1 To feOrden.Rows - 1
                    If feOrden.TextMatrix(i, 1) = Datoscontrato(i).sAgeCod And feOrden.TextMatrix(i, 2) = Datoscontrato(i).sObjeto Then
                        If feOrden.TextMatrix(i, 6) > Datoscontrato(i).nTotal Then
                             MsgBox "El Monto de los Bienes/Servicios modificados no debe ser superior a los Items Actuales.", vbInformation, "Aviso"
                            ValidaAdenda = False
                             Exit Function
                        End If
                    End If
            Next
        End If
    End If
    ValidaAdenda = True
End Function
Private Sub feOrden_OnCellChange(pnRow As Long, pnCol As Long)
    On Error GoTo ErrfeOrden_OnCellChange
    If feOrden.TextMatrix(1, 0) <> "" Then
        If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or _
            fntpodocorigen = LogTipoContrato.ContratoSuministro Then
            If pnCol = 4 Or pnCol = 5 Then
                feOrden.TextMatrix(pnRow, 6) = Format(Val(feOrden.TextMatrix(pnRow, 4)) * feOrden.TextMatrix(pnRow, 5), gsFormatoNumeroView)
            End If
        End If
    End If
    Exit Sub
ErrfeOrden_OnCellChange:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub AsignaObjetosSerItem(ByVal sCtaCod As String)
    Dim nNiv As Integer
    Dim nObj As Integer
    Dim nObjs As Integer
    Dim oCon As New DConecta
    Dim oCtaCont As New DCtaCont
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim oRHAreas As New DActualizaDatosArea
    Dim oCtaIf As New NCajaCtaIF
    Dim oEfect As New Defectivo
    Dim oDescObj As New ClassDescObjeto
    Dim oContFunct As New NContFunciones
    Dim lsRaiz As String, lsFiltro As String, sSql As String
        
    oDescObj.lbUltNivel = True
    oCon.AbreConexion
    EliminaObjeto feOrden.row

    sSql = "SELECT MAX(nCtaObjOrden) as nNiveles FROM CtaObj WHERE cCtaContCod = '" & sCtaCod & "' and cObjetoCod <> '00' "
    Set rs = oCon.CargaRecordSet(sSql)
    nObjs = IIf(IsNull(rs!nNiveles), 0, rs!nNiveles)
      
    Set rs1 = oCtaCont.CargaCtaObj(sCtaCod, , True)
    If Not rs1.EOF And Not rs1.BOF Then
        Do While Not rs1.EOF
            lsRaiz = ""
            lsFiltro = ""
            Set rs = New ADODB.Recordset
            Select Case Val(rs1!cObjetoCod)
                Case ObjCMACAgencias
                    Set rs = oRHAreas.GetAgencias()
                Case ObjCMACAgenciaArea
                    lsRaiz = "Unidades Organizacionales"
                    Set rs = oRHAreas.GetAgenciasAreas()
                Case ObjCMACArea
                    Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
                Case ObjEntidadesFinancieras
                    lsRaiz = "Cuentas de Entidades Financieras"
                    Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, sCtaCod)
                Case ObjDescomEfectivo
                    Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
                Case ObjPersona
                    Set rs = Nothing
                Case Else
                    lsRaiz = "Varios"
                    Set rs = GetObjetos(rs1!cObjetoCod)
            End Select
            If Not rs Is Nothing Then
                If rs.State = adStateOpen Then
                    If Not rs.EOF And Not rs.BOF Then
                        If rs.RecordCount > 1 Then
                            oDescObj.Show rs, "", lsRaiz
                            If oDescObj.lbOk Then
                                lsFiltro = oContFunct.GetFiltroObjetos(Trim(rs1!cObjetoCod), sCtaCod, oDescObj.gsSelecCod, False)
                                'AdicionaObjeto feOrden.TextMatrix(feOrden.row, 0), IIf(IsNull(rs1!nCtaObjOrden), "", rs1!nCtaObjOrden), oDescObj.gsSelecCod, oDescObj.gsSelecDesc, lsFiltro, IIf(IsNull(rs1!cObjetoCod), "", rs1!cObjetoCod)
                            Else
                                'EliminaObjeto feOrden.row
                                Exit Do
                            End If
                        Else
                            'AdicionaObjeto feOrden.TextMatrix(feOrden.row, 0), IIf(IsNull(rs1!nCtaObjOrden), "", rs1!nCtaObjOrden), oDescObj.gsSelecCod, oDescObj.gsSelecDesc, lsFiltro, IIf(IsNull(rs1!cObjetoCod), "", rs1!cObjetoCod) 'Comentado PASIERS0772014
                        End If
                    End If
                End If
            End If
            rs1.MoveNext
        Loop
    End If

    Set rs = Nothing
    Set rs1 = Nothing
    Set oDescObj = Nothing
    Set oCon = Nothing
    Set oCtaCont = Nothing
    Set oCtaIf = Nothing
    Set oEfect = Nothing
    Set oContFunct = Nothing
    Set oContFunct = Nothing
    Exit Sub
End Sub
Private Function DameCtaCont(ByVal psObjeto As String, nNiv As Integer, psAgeCod As String) As String 'PASI20140110ERS0772014
    Dim oCon As New DConecta
    Dim oForm As New frmLogOCompra
    Dim rs As New ADODB.Recordset
    Dim sSql As String
    
    sSql = oForm.FormaSelect(gcOpeCod, psObjeto, 0, psAgeCod)
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    If Not rs.EOF Then
        DameCtaCont = rs!cObjetoCod
    End If
    Set rs = Nothing
    Set oForm = Nothing
    Set oCon = Nothing
End Function
Private Sub feOrden_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean) 'PASI20140110 ERS0772014
    If psDataCod <> "" Then
'        If pnCol = 2 Then
'            If fnTpoDocOrigen = LogTipoContrato.ContratoServicio Then
'                AsignaObjetosSerItem psDataCod
'            End If
'        End If
'        If pnCol = 1 Or pnCol = 2 Then
            '*** Si esta vacio el campo de la cuenta contable y si ya eligió agencia y objeto
            If Len(Trim(feOrden.TextMatrix(pnRow, 1))) <> 0 And Len(Trim(feOrden.TextMatrix(pnRow, 2))) <> 0 Then
                feOrden.TextMatrix(pnRow, 7) = DameCtaCont(feOrden.TextMatrix(pnRow, 2), 0, Trim(feOrden.TextMatrix(pnRow, 1)))
            End If
            '***
'        End If
    End If
End Sub
Private Sub feOrden_RowColChange()  'PASI20140110 ERS0772014
    If feOrden.col = 1 Then
        feOrden.rsTextBuscar = fRsAgencia
    ElseIf feOrden.col = 2 Then
        If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or _
            fntpodocorigen = LogTipoContrato.ContratoSuministro Then
            feOrden.rsTextBuscar = fRsCompra
        End If
    End If
End Sub
Private Sub feOrden_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean) 'PASI20140110 ERS0772014
    Dim sColumnas() As String
    sColumnas = Split(feOrden.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    If fnAdenda = 0 Then
        CargaDatos
    Else
        CargaDatosDetalle
    End If
End Sub
Private Sub CargaDatosDetalle()
    Dim oLog As DLogGeneral
    Dim oConst As DConstantes
    Set oConst = New DConstantes
    Dim rsBienes As ADODB.Recordset
    
    Set oLog = New DLogGeneral
    CargaCombo oConst.GetConstante(gsLogContTipoAdendas), Me.cboTpoReajuste
    
     If fntpodocorigen = LogTipoContrato.ContratoObra Then
        frmLogContReajAdenda.Caption = "Reajuste de Contrato de Obra"
        sstReajuste.TabCaption(0) = "Datos de Reajuste"
    Else
        frmLogContReajAdenda.Caption = "Adenda de Contrato de Suministros y Bienes"
        sstReajuste.TabCaption(0) = "Datos de la Adenda"
    End If
    'Datos de Reajuste
    Set rsLog = oLog.ListaContratoAdendaReajusteDet(fsNContrato, fnContRef, fnAdenda)
    If rsLog.RecordCount > 0 Then
        Me.txtNContrato.Text = rsLog!NContrato
        Me.txtProveedor.Text = rsLog!Proveedor
        Me.txtNReajuste.Text = CInt(rsLog!UltAdenda)
        txtNReajuste.Enabled = False
        cboTpoReajuste.ListIndex = IndiceListaCombo(cboTpoReajuste, rsLog!nTipo)
        txtTpoMoneda.Text = IIf(rsLog!nMoneda = 1, gcPEN_SIMBOLO, "$.") 'marg ers044-2016
        txtMonto.Text = Format(rsLog!nMonto, "#,#0.00")
        dtpFecDesde.value = CDate(rsLog!dFecIni)
        dtpFecHasta.value = CDate(rsLog!dFecFin)
        If rsLog!cNombreArchivo <> "" Then
            Me.chbreajuste.value = True
            Me.lblNombreArchivo.Caption = rs!cNombreArchivo
        End If
    End If
    'Datos Bienes
    If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or fntpodocorigen = LogTipoContrato.ContratoSuministro Then
        If rsLog!nTipo <> 1 Then
            Set rsBienes = oLog.ListaContratoBienAdendaDet(fsNContrato, fnContRef, fnAdenda)
            If Not rsBienes.EOF Then
                Do While Not rsBienes.EOF
                    feOrden.TextMatrix(feOrden.row, 1) = rsBienes!cAgeDest
                    feOrden.TextMatrix(feOrden.row, 2) = rsBienes!cBSCod
                    feOrden.TextMatrix(feOrden.row, 3) = rsBienes!cBSDescripcion
                    feOrden.TextMatrix(feOrden.row, 4) = rsBienes!nCant
                    feOrden.TextMatrix(feOrden.row, 5) = rsBienes!nPrec
                    feOrden.TextMatrix(feOrden.row, 6) = rsBienes!nImporte
                    rsBienes.MoveNext
                Loop
            End If
            DesHabilitaTab True
        Else
            DesHabilitaTab False
        End If
    ElseIf fntpodocorigen = LogTipoContrato.ContratoObra Then
          DesHabilitaTab False
    End If
    cmdAgregarItemCont.Enabled = False
    cmdQuitarItemCont.Enabled = False
    Me.cboTpoReajuste.Enabled = False
    txtMonto.Enabled = False
    dtpFecDesde.Enabled = False
    dtpFecHasta.Enabled = False
    chbreajuste.Enabled = False
    cmdExaminar.Enabled = False
    cmdRegistrar.Enabled = False
    cmdCancelar.Enabled = False
End Sub
Private Sub CargaDatos()
    Dim oLog As DLogGeneral
    Dim oConst As DConstantes
    Set oConst = New DConstantes

    Set oLog = New DLogGeneral
    CargaCombo oConst.GetConstante(gsLogContTipoAdendas), Me.cboTpoReajuste
    
    If fntpodocorigen = LogTipoContrato.ContratoObra Then
        frmLogContReajAdenda.Caption = "Reajuste de Contrato de Obra"
        sstReajuste.TabCaption(0) = "Datos de Reajuste"
    Else
        frmLogContReajAdenda.Caption = "Adenda de Contrato de Suministros y Bienes"
        sstReajuste.TabCaption(0) = "Datos de la Adenda"
    End If
    
    Set rsLog = oLog.ObtenerContratoAdendaReajuste(fsNContrato, fnContRef)
    
    If rsLog.RecordCount > 0 Then
            Me.txtNContrato.Text = rsLog!NContrato
            Me.txtProveedor.Text = rsLog!Proveedor
        'If fnAdenda = 0 Then
            Me.txtNReajuste.Text = CInt(rsLog!UltAdenda) + 1
            txtNReajuste.Enabled = False
            cboTpoReajuste.ListIndex = -1
            txtTpoMoneda.Text = IIf(rsLog!nMoneda = 1, gcPEN_SIMBOLO, "$.") 'marg ers044-2016
            txtMonto.Text = Format(rsLog!nMonto, "#,#0.00")
            dtpFecDesde.value = CDate(rsLog!dFecIni)
            dtpFecHasta.value = CDate(rsLog!dFecFin)
            
'        Else
'            Me.txtNReajuste.Text = CInt(rsLog!UltAdenda)
'            cboTpoReajuste.ListIndex = IndiceListaCombo(cboTpoReajuste, rsLog!nTipo)
'            txtTpoMoneda.Text = IIf(rsLog!nMoneda = 1, "S/.", "$.")
'            txtMonto.Text = Format(rsLog!nMonto, "#,#0.00")
'            dtpFecDesde.value = CDate(rsLog!dFecIni)
'            dtpFecHasta.value = CDate(rsLog!dFecFin)
'            If Len(rsLog!cNombreArchivo) > 0 Then
'                chbreajuste.value = 1
'                lblNombreArchivo.Caption = rsLog!cNombreArchivo
'            End If
'        End If
        If fntpodocorigen = LogTipoContrato.ContratoObra Then
            DesHabilitaTab False
        Else
            DesHabilitaTab True
            fraItemContrato.Enabled = False
        End If
        DesHabilitaControles False, False, False
        txtNContrato.Enabled = False
        txtProveedor.Enabled = False
    End If
    dtpFecDesde.Enabled = False
    dtpFecHasta.Enabled = False
    CargaVariables
End Sub
Private Sub DesHabilitaTab(ByVal phEstado As Boolean)
    sstReajuste.TabVisible(1) = phEstado
End Sub
Private Sub DesHabilitaControles(ByVal pbEstado As Boolean, ByVal pbEstadMonto As Boolean, Optional ByVal pbItem As Boolean)
    Dim oConstSist As NConstSistemas
    Set oConst = New DConstantes
    txtMonto.Enabled = pbEstadMonto
    dtpFecDesde.Enabled = Not pbEstadMonto
    dtpFecHasta.Enabled = Not pbEstadMonto
    lblNombreArchivo.Enabled = pbEstado
    cmdRegistrar.Enabled = pbEstado
    cmdCancelar.Enabled = pbEstado
    
    'If Trim(Mid(GetMaquinaUsuario, 1, 2)) = "01" And phEstado = True Then 'Activar PASI20140826
    If pbEstado = True Then
        chbreajuste.Enabled = pbEstado
        Me.cmdExaminar.Enabled = Not pbEstado
        pbActivaArchivo = pbEstado
    
        'OBTENER RUTA DE CONTRATOS
        Set oConstSist = New NConstSistemas
        psRutaContrato = Trim(oConstSist.LeeConstSistema(gsLogContRutaContratos))
    Else
        chbreajuste.Enabled = False
        Me.cmdExaminar.Enabled = False
        pbActivaArchivo = False
        psRutaContrato = ""
    End If
        'HABILITAMOS LOS ITEMS DE CONTRATO
    If pbItem = False Then
        sstReajuste.TabEnabled(1) = pbItem
        fraItemContrato.Enabled = pbItem
    Else
        sstReajuste.TabEnabled(1) = pbItem
        fraItemContrato.Enabled = pbItem
    End If
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 10, 3)
    If KeyAscii = 13 Then
        'Me.dtpFecDesde.SetFocus
    End If
End Sub
Sub GrabarArchivo()
    If Trim(fsRuta) <> "" Then
        Dim RutaFinal As String
        RutaFinal = psRutaContrato
        Dim a As New Scripting.FileSystemObject
    
    If a.FolderExists(RutaFinal) = False Then
        a.CreateFolder (RutaFinal)
    End If
    Copiar fsRuta, RutaFinal & Trim(lblNombreArchivo.Caption)
    Else
        MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
    End If
End Sub
Private Sub Copiar(Archivo As String, Destino As String)
Dim a As New Scripting.FileSystemObject
If a.FileExists(Destino) = False Then
    a.CopyFile Archivo, Destino
Else
    MsgBox "Archivo ya existe", vbInformation, "Aviso"
End If
End Sub
Private Sub LimpiarControles()
    txtMonto.Text = ""
    'dtpFecDesde.value = CDate(gdFecSis)
    'dtpFecHasta.value = CDate(gdFecSis)
    chbreajuste.value = 0
    lblNombreArchivo.Caption = ""
    ReDim Datoscontrato(0)
    Call LimpiaFlex(feOrden)
End Sub
Public Function FormaSelect(psOpeCod As String, sObj As String, nNiv As Integer, psAgeCod As String) As String
    Dim sText As String
    sText = " SELECT b.cCtaContCod cObjetoCod, b.cCtaContDesc cObjetoDesc, e.cBSCod cObjCod," _
          & " upper(e.cBSDescripcion) as cObjDesc, 1 nObjetoNiv, CO.cConsDescripcion " _
          & " FROM  CtaCont b " _
          & " Inner JOIN CtaBS  c ON Replace(c.cCtaContCod,'AG','" & psAgeCod & "') = b.cCtaContCod And cOpeCod = '" & psOpeCod & "'" _
          & " Inner JOIN BienesServicios e ON e.cBSCod like c.cObjetoCod + '%'" _
          & " Inner Join Constante CO On nBSunidad = CO.nConsValor And nConsCod = '1019'" _
          & " "
    If nNiv > 0 Then
       sText = sText & "WHERE d.nObjetoNiv = " & nNiv & " "
    End If
    FormaSelect = sText & IIf(sObj <> "", " And e.cBSCod = '" & sObj & "' ", sObj) _
                & "ORDER BY e.cBSCod"
End Function

