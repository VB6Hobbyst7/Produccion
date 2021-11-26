VERSION 5.00
Begin VB.Form FrmPigRemateAdjud 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas en Remate - Verificador y Bloqueo de Piezas"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "FrmPigRemateAdjud.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Datos de Venta(Piezas)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   2505
      Left            =   90
      TabIndex        =   29
      Top             =   2895
      Width           =   9045
      Begin SICMACT.EditMoney txtValorVenta 
         Height          =   285
         Left            =   7830
         TabIndex        =   30
         Top             =   2160
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         BackColor       =   -2147483624
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.FlexEdit fepiezas 
         Height          =   1860
         Left            =   75
         TabIndex        =   31
         Top             =   240
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   3281
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Pieza-Tipo-Material-Descripcion-PNeto-ValBase-ValVenta-Cliente-prueba"
         EncabezadosAnchos=   "350-500-1100-1200-1250-700-850-850-2000-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-7-8-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-C-C-C-C-C-R-R-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-2-2-1-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbOrdenaCol     =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label10 
         Caption         =   "Valor Vta."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   6840
         TabIndex        =   32
         Top             =   2220
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdLote 
      Caption         =   "VENTAS POR LOTE "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      Left            =   105
      TabIndex        =   28
      Top             =   2895
      Width           =   9030
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
      Height          =   330
      Left            =   4170
      TabIndex        =   26
      Top             =   915
      Width           =   330
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
      Left            =   6705
      TabIndex        =   25
      Top             =   5535
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
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
      Height          =   375
      Left            =   7890
      TabIndex        =   24
      Top             =   5535
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Height          =   375
      Left            =   5535
      TabIndex        =   23
      Top             =   5520
      Width           =   1095
   End
   Begin SICMACT.EditMoney emvalrem 
      Height          =   315
      Left            =   7815
      TabIndex        =   22
      Top             =   1635
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      Text            =   "0"
      Enabled         =   -1  'True
   End
   Begin SICMACT.EditMoney emdeuda 
      Height          =   315
      Left            =   7815
      TabIndex        =   21
      Top             =   1275
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      Text            =   "0"
      Enabled         =   -1  'True
   End
   Begin SICMACT.EditMoney emreta 
      Height          =   315
      Left            =   7815
      TabIndex        =   20
      Top             =   915
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      Text            =   "0"
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtpneto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1395
      TabIndex        =   19
      Top             =   1650
      Width           =   660
   End
   Begin VB.TextBox txtpieza 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1395
      TabIndex        =   18
      Top             =   1320
      Width           =   660
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos de Venta(Lote)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   705
      Left            =   2940
      TabIndex        =   11
      Top             =   2160
      Width           =   6180
      Begin VB.TextBox txtcompra 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3465
         TabIndex        =   15
         Top             =   255
         Width           =   2520
      End
      Begin SICMACT.EditMoney emvalvtalote 
         Height          =   285
         Left            =   1125
         TabIndex        =   13
         Top             =   315
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         BackColor       =   -2147483624
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label11 
         Caption         =   "Comprador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2460
         TabIndex        =   14
         Top             =   330
         Width           =   930
      End
      Begin VB.Label Label10 
         Caption         =   "Valor Vta."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rematado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   645
      Left            =   135
      TabIndex        =   5
      Top             =   2160
      Width           =   2460
      Begin VB.OptionButton oppieza 
         Caption         =   "Piezas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1290
         TabIndex        =   7
         Top             =   330
         Width           =   915
      End
      Begin VB.OptionButton oplote 
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   330
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   8790
      Begin VB.TextBox txtubica 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3825
         TabIndex        =   4
         Top             =   165
         Width           =   4815
      End
      Begin VB.TextBox txtremate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Left            =   975
         TabIndex        =   3
         Top             =   180
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "UBICACION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   2730
         TabIndex        =   2
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "REMATE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   180
         TabIndex        =   1
         Top             =   225
         Width           =   780
      End
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   375
      Left            =   330
      TabIndex        =   27
      Top             =   870
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Texto           =   "Crédito"
      EnabledCta      =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Label lblEstadoCont 
      Height          =   330
      Left            =   3315
      TabIndex        =   33
      Top             =   5190
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   150
      X2              =   9000
      Y1              =   2085
      Y2              =   2085
   End
   Begin VB.Label Label6 
      Caption         =   "P.Neto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   390
      TabIndex        =   17
      Top             =   1725
      Width           =   705
   End
   Begin VB.Label Label5 
      Caption         =   "Piezas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   390
      TabIndex        =   16
      Top             =   1380
      Width           =   690
   End
   Begin VB.Label Label9 
      Caption         =   "Valor de Remate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   6000
      TabIndex        =   10
      Top             =   1710
      Width           =   1425
   End
   Begin VB.Label Label8 
      Caption         =   "Deuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6000
      TabIndex        =   9
      Top             =   1365
      Width           =   645
   End
   Begin VB.Label Label7 
      Caption         =   "Retasacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6000
      TabIndex        =   8
      Top             =   1005
      Width           =   1035
   End
End
Attribute VB_Name = "FrmPigRemateAdjud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim lnTipoProceso As Integer
'
'Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
'Dim oContrato As DPigContrato
'Dim rsDatos As Recordset
'Dim psCuenta As String
'Dim rs As Recordset
'
'    If KeyAscii = 13 Then
'
'        psCuenta = Me.AXCodCta.NroCuenta
'        Set oContrato = New DPigContrato
'
'        Set rsDatos = oContrato.dObtieneContratosLotes(psCuenta, lnTipoProceso)
'        If Not (rsDatos.EOF And rsDatos.BOF) Then
'            txtpieza.Text = rsDatos!npiezas
'            txtpneto.Text = rsDatos!pesoneto
'            emreta.Text = rsDatos!retasacion
'            emdeuda.Text = rsDatos!valordeuda
'            emvalrem.Text = rsDatos!valorproceso
'            lblEstadoCont = rsDatos!nPrdEstado
'            cmdGrabar.Enabled = True
'        End If
'        Set rsDatos = Nothing
'
'        Set rs = oContrato.dObtieneContratosPiezas(psCuenta, lnTipoProceso)
'
'        If Not (rs.EOF And rs.BOF) Then
'          fepiezas.Clear
'          fepiezas.Rows = 2
'          fepiezas.FormaCabecera
'
'          Do While Not rs.EOF
'                fepiezas.AdicionaFila
'                fepiezas.TextMatrix(fepiezas.Rows - 1, 1) = rs!nItemPieza
'                fepiezas.TextMatrix(fepiezas.Rows - 1, 2) = rs!Descritipo
'                fepiezas.TextMatrix(fepiezas.Rows - 1, 3) = rs!cConsDescripcion
'                fepiezas.TextMatrix(fepiezas.Rows - 1, 4) = rs!cDescripcion
'                fepiezas.TextMatrix(fepiezas.Rows - 1, 5) = rs!npesoneto
'                fepiezas.TextMatrix(fepiezas.Rows - 1, 6) = rs!nValorProceso
'                rs.MoveNext
'          Loop
'        End If
'
'        Set rs = Nothing
'        Set oContrato = Nothing
'        oplote.SetFocus
'
'    End If
'End Sub
'
'Private Sub cmdBuscar_Click()
'    FrmPigContratosRem.Show 1
'End Sub
'
'Private Sub cmdCancelar_Click()
'    Limpia
'End Sub
'
'Private Sub cmdGrabar_Click()
'Dim oGraba As NPigRemate
'Dim rs As Recordset
'Dim lnTipoVenta As Integer
'Dim lsCliente As String
'Dim psCuenta As String
'Dim lnValorVta As Currency
'Dim oCont As NContFunciones
'Dim lsMovNro As String
'Dim lsFechaHoraGrab As String
'
'    psCuenta = Me.AXCodCta.NroCuenta
'    Set rs = fepiezas.GetRsNew
'
'    Set oCont = New NContFunciones
'    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'    Set oCont = Nothing
'
'    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
'
'    If oplote Then 'EN CASO DE QUE LA VENTA SEA POR LOTE
'        lnTipoVenta = 1
'        lnValorVta = emvalvtalote
'        lsCliente = txtcompra.Text
'
'    ElseIf oppieza Then
'        lnTipoVenta = 2
'        lnValorVta = txtValorVenta
'        lsCliente = ""
'    End If
'
'    Set oGraba = New NPigRemate
'    oGraba.nPigVentaRemate psCuenta, lnValorVta, lnTipoVenta, lsFechaHoraGrab, rs, lnTipoProceso, lblEstadoCont, lsCliente
'    Set oGraba = Nothing
'
'    Limpia
'
'End Sub
'
'Private Sub Limpia()
'    txtpieza.Text = ""
'    txtpneto.Text = ""
'    emreta.Text = ""
'    emdeuda.Text = ""
'    emvalrem.Text = ""
'    emvalvtalote.Text = ""
'    txtcompra.Text = ""
'    oplote.value = True
'    AXCodCta.Cuenta = ""
'    AXCodCta.SetFocusCuenta
'    cmdLote.Visible = True
'    fepiezas.Clear
'    fepiezas.Rows = 2
'    fepiezas.FormaCabecera
'    cmdGrabar.Enabled = False
'End Sub
'Private Sub cmdsalir_Click()
'    Unload Me
'End Sub
'Private Sub emvalvtalote_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If emvalvtalote <> "" Then
'            If CCur(emvalrem) <= CCur(emvalvtalote) Then
'                  Me.txtcompra.SetFocus
'            Else
'                  MsgBox "El Valor de Compra no puede ser menor al Valor del Remate"
'                  emvalvtalote.Text = 0
'                  emvalvtalote.SetFocus
'            End If
'        End If
'    End If
'End Sub
'Private Sub fepiezas_RowColChange()
'If fepiezas.Col = 8 Then
'    If fepiezas.TextMatrix(fepiezas.Row, 7) <> "" Then
'        If CCur(fepiezas.TextMatrix(fepiezas.Row, 6)) <= CCur(fepiezas.TextMatrix(fepiezas.Row, 7)) Then
'              txtValorVenta = txtValorVenta + fepiezas.TextMatrix(fepiezas.Row, 8)
'         Else
'             fepiezas.TextMatrix(fepiezas.Row, 7) = 0
'             MsgBox "El Valor de la Venta no debe ser menor al Valor de la Base "
'         End If
'    End If
'    suma_pieza
'End If
'End Sub
'Private Sub suma_pieza()
'    Dim Total As Integer
'    Dim CantFila As Integer
'    Dim i As Integer
'    Total = 0
'    i = 1
'    CantFila = Me.fepiezas.Rows - 1
'    Do While CantFila >= i
'             If IsNull(Me.fepiezas.TextMatrix(i, 7)) Or Me.fepiezas.TextMatrix(i, 7) = "" Then
'                Me.fepiezas.TextMatrix(i, 7) = 0
'             End If
'           Total = Total + Me.fepiezas.TextMatrix(i, 7)
'           i = i + 1
'      Loop
'      txtValorVenta = Total
'End Sub
'
'Private Sub Form_Activate()
'Me.AXCodCta.SetFocusCuenta
'
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
' Dim nRemate As DPigContrato
' Dim nDatosRemate As ADODB.Recordset
' Set nRemate = New DPigContrato
' Set nDatosRemate = nRemate.dObtieneDatosRemate(nRemate.dObtieneMaxRemate() - 1)
' AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
' AXCodCta.Age = ""
' If Not (nDatosRemate.EOF And nDatosRemate.BOF) Then
'    txtremate.Text = nDatosRemate!nRemate
'    txtubica.Text = nDatosRemate!cConsDescripcion
'    lnTipoProceso = nDatosRemate!nTipoProceso
' End If
' Set nDatosRemate = Nothing
' Set nRemate = Nothing
' Me.Icon = LoadPicture(App.path & gsRutaIcono)
'End Sub
'
'Private Sub oplote_Click()
'
'    fepiezas.Enabled = False
'    cmdLote.Visible = True
'    emvalvtalote.Enabled = True
'    emvalvtalote.SetFocus
'
'End Sub
'Private Sub oppieza_Click()
'
'    fepiezas.Enabled = True
'    cmdLote.Visible = False
'    emvalvtalote.Enabled = False
'
'End Sub
