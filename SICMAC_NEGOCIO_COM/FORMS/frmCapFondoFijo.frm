VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCapFondoFijo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmCapFondoFijo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   105
      TabIndex        =   5
      Top             =   3090
      Width           =   975
   End
   Begin VB.Frame fraMoneda 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   960
      Left            =   75
      TabIndex        =   8
      Top             =   75
      Width           =   2595
      Begin VB.OptionButton optMoneda 
         Caption         =   "Moneda Extranjera"
         Height          =   315
         Index           =   1
         Left            =   285
         TabIndex        =   1
         Top             =   540
         Width           =   1680
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "Moneda Nacional"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   0
         Top             =   270
         Width           =   1680
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4935
      TabIndex        =   4
      Top             =   3090
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   3090
      Width           =   975
   End
   Begin VB.Frame fraOrden 
      Caption         =   "Orden Pago"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1890
      Left            =   90
      TabIndex        =   7
      Top             =   1125
      Width           =   5835
      Begin SICMACT.TxtBuscar txtBanco 
         Height          =   345
         Left            =   1005
         TabIndex        =   18
         Top             =   1080
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
      End
      Begin VB.Label lblBanco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   990
         TabIndex        =   19
         Top             =   1470
         Width           =   4350
      End
      Begin VB.Label lblEtqBanco 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   1140
         Width           =   555
      End
      Begin VB.Label lblMonto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3810
         TabIndex        =   16
         Top             =   690
         Width           =   1560
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   3180
         TabIndex        =   15
         Top             =   765
         Width           =   540
      End
      Begin VB.Label lblFecDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1020
         TabIndex        =   14
         Top             =   690
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   765
         Width           =   540
      End
      Begin VB.Label lblProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1020
         TabIndex        =   12
         Top             =   300
         Width           =   4350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor :"
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   345
         Width           =   825
      End
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Buscar..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   945
      Left            =   2760
      TabIndex        =   6
      Top             =   83
      Width           =   3165
      Begin MSMask.MaskEdBox txtOrden 
         Height          =   330
         Left            =   150
         TabIndex        =   9
         Top             =   495
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "########"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   1740
         TabIndex        =   2
         Top             =   465
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Orden N°"
         Height          =   210
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmCapFondoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nMoneda As COMDConstantes.Moneda
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim nMovRef As Long

Public Sub Inicia(ByVal nOpe As CaptacOperacion, ByVal sDescOperacion As String)
nOperacion = nOpe
Me.Caption = "Captaciones - Retiro - " & sDescOperacion
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
optMoneda(0).value = True
If nOperacion = gAhoRetFondoFijoCanje Then
    txtBanco.Visible = True
    lblBanco.Visible = True
    lblEtqBanco.Visible = True
    Dim clsBanco As COMNCajaGeneral.NCOMCajaCtaIF 'NCajaCtaIF
    Dim rsBanco As New ADODB.Recordset
    Set clsBanco = New COMNCajaGeneral.NCOMCajaCtaIF
    Set rsBanco = clsBanco.CargaCtasIF(gMonedaNacional, "_1%", MuestraInstituciones, "1")
    Set clsBanco = Nothing
    txtBanco.rs = rsBanco
Else
    txtBanco.Visible = False
    lblBanco.Visible = False
    lblEtqBanco.Visible = False
End If
Me.Show 1
End Sub

Private Sub cmdBuscar_Click()
Dim oGeneral As COMNCajaGeneral.NCOMCajaGeneral 'nCajaGeneral
Dim rsOP As New ADODB.Recordset
Dim dInicio As Date, dFin As Date

If Trim(txtOrden) = "" Then
    MsgBox "Número de Orden No válido", vbInformation, "Aviso"
    txtOrden.SetFocus
    Exit Sub
End If
Set oGeneral = New COMNCajaGeneral.NCOMCajaGeneral
Set rsOP = oGeneral.GetOrdPagFondoFijoEntregado(nMoneda, txtOrden)
'Set oGeneral = Nothing
If Not (rsOP.EOF And rsOP.BOF) Then
    'lblProveedor = PstaNombre(rsOP("cPersNombre"))
    lblProveedor = PstaNombre(rsOP("cNomPers"))
    lblMonto = Format$(Abs(rsOP("nDocImporte")), "#,##0.00")
    lblFecDoc = Format$(rsOP("dDocFecha"), "dd/mm/yyyy")
    nMovRef = rsOP("nMovNro")
    cmdGrabar.Enabled = True
    txtOrden.Enabled = False
    cmdCancelar.Enabled = True
    frmSegSepelioAfiliacion.Inicio rsOP("cObjetoCod"), , gSegTpoBusPersCod
Else
    lblProveedor = ""
    lblMonto = ""
    lblFecDoc = ""
    nMovRef = 0
    MsgBox "No existen Ordenes de Pago de Fondo de Fijo con las condiciones de búsqueda", vbInformation, "Aviso"
    cmdGrabar.Enabled = False
End If

rsOP.Close
Set rsOP = Nothing
Set oGeneral = Nothing
End Sub

Private Sub cmdCancelar_Click()
txtOrden.Text = "________"
txtOrden.Enabled = True
cmdGrabar.Enabled = False
lblProveedor = ""
lblMonto = ""
lblFecDoc = ""
nMovRef = 0
txtOrden.SetFocus
cmdCancelar.Enabled = False
End Sub

Private Sub cmdGrabar_Click()
    Dim sMovNro As String
    Dim sDocNro As String
    Dim nMonto As Double
    Dim oCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    Dim oMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
    Dim dFecDoc As Date
    Dim lsBoleta As String
    Dim nFicSal As Integer
    
    If nOperacion = gAhoRetFondoFijoCanje Then
        If Trim(txtBanco.Text) = "" Then
            MsgBox "Debe registrar el Banco del canje.", vbInformation, "Aviso"
            txtBanco.SetFocus
            Exit Sub
        End If
    End If
    If MsgBox("¿Desea Grabar la Operación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oMov = Nothing
    sDocNro = Trim(txtOrden)
    nMonto = CDbl(lblMonto)
    dFecDoc = CDate(lblFecDoc)
    Set oCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    oCap.CapCargoFondoFijo sDocNro, nMonto, nMovRef, nOperacion, sMovNro, dFecDoc, lblProveedor.Caption, gsNomAge, sLpt, IIf(optMoneda(0).value = True, Moneda.gMonedaNacional, Moneda.gMonedaExtranjera), lsBoleta, gbImpTMU
    
    If Trim(lsBoleta) <> "" Then
        nFicSal = FreeFile
        Open sLpt For Output As nFicSal
            Print #nFicSal, lsBoleta
            Print #nFicSal, ""
        Close #nFicSal
    End If
                            
    Set oCap = Nothing
    cmdCancelar_Click
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub optMoneda_Click(Index As Integer)
Dim oGeneral As COMNCajaGeneral.NCOMCajaGeneral 'nCajaGeneral
Dim rsOP As ADODB.Recordset
Select Case Index
    Case 0
        nMoneda = gMonedaNacional
    Case 1
        nMoneda = gMonedaExtranjera
End Select
End Sub

Private Sub optMoneda_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    txtOrden.SetFocus
End If
End Sub

Private Sub txtBanco_EmiteDatos()
lblBanco = Trim(txtBanco.psDescripcion)
End Sub

Private Sub txtOrden_GotFocus()
With txtOrden
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtOrden_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar.SetFocus
End If
End Sub
