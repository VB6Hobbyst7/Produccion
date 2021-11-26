VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmPigFacturaRemate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturacion de Piezas Venta en Remate"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPigFacturaRemate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1830
      Left            =   75
      TabIndex        =   10
      Top             =   -15
      Width           =   8790
      Begin VB.TextBox txtDirPers 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1335
         TabIndex        =   27
         Top             =   1380
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.TextBox txtDocJur 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7005
         TabIndex        =   22
         Top             =   990
         Width           =   1290
      End
      Begin VB.TextBox txtPersCod 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1335
         TabIndex        =   20
         Top             =   990
         Width           =   1980
      End
      Begin VB.TextBox txtnombre 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1335
         TabIndex        =   13
         Top             =   1380
         Width           =   6975
      End
      Begin VB.TextBox txtDocId 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4830
         TabIndex        =   12
         Top             =   1020
         Width           =   1185
      End
      Begin VB.CommandButton cmdBuscar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7830
         Picture         =   "FrmPigFacturaRemate.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Buscar ..."
         Top             =   285
         Width           =   465
      End
      Begin MSMask.MaskEdBox txtNumDoc 
         Height          =   300
         Left            =   1890
         TabIndex        =   24
         Top             =   225
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   16711680
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtSerie 
         Height          =   300
         Left            =   1350
         TabIndex        =   25
         Top             =   225
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   16711680
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         Caption         =   "Número :"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label9 
         Caption         =   "Doc Jur"
         Height          =   180
         Left            =   6225
         TabIndex        =   23
         Top             =   1050
         Width           =   750
      End
      Begin VB.Label Label8 
         Caption         =   "Cliente"
         Height          =   180
         Left            =   255
         TabIndex        =   21
         Top             =   1425
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   225
         Left            =   225
         TabIndex        =   17
         Top             =   1035
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Doc ID"
         Height          =   180
         Left            =   4110
         TabIndex        =   16
         Top             =   1065
         Width           =   675
      End
      Begin VB.Label Label6 
         Caption         =   "Comprador"
         Height          =   225
         Left            =   225
         TabIndex        =   15
         Top             =   660
         Width           =   930
      End
      Begin VB.Label lbalias 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1335
         TabIndex        =   14
         Top             =   615
         Width           =   2925
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   330
      Left            =   3180
      TabIndex        =   9
      Top             =   6915
      Width           =   1035
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   4335
      TabIndex        =   8
      Top             =   6900
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   5745
      TabIndex        =   1
      Top             =   5895
      Width           =   3135
      Begin SICMACT.EditMoney txtSubTotal 
         Height          =   255
         Left            =   1515
         TabIndex        =   19
         Top             =   165
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   450
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
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtComMart 
         Height          =   255
         Left            =   1515
         TabIndex        =   5
         Top             =   450
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   450
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
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtIGV 
         Height          =   240
         Left            =   1515
         TabIndex        =   6
         Top             =   735
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   423
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
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtTotal 
         Height          =   255
         Left            =   1530
         TabIndex        =   7
         Top             =   1140
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         BackColor       =   12648447
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1485
         X2              =   3060
         Y1              =   1065
         Y2              =   1050
      End
      Begin VB.Label Label7 
         Caption         =   "Sub.Total"
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   210
         Width           =   1125
      End
      Begin VB.Label Label5 
         Caption         =   "TOTAL"
         Height          =   165
         Left            =   135
         TabIndex        =   4
         Top             =   1185
         Width           =   1140
      End
      Begin VB.Label Label4 
         Caption         =   "IGV"
         Height          =   180
         Left            =   135
         TabIndex        =   3
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "Com.Martillero"
         Height          =   180
         Left            =   105
         TabIndex        =   2
         Top             =   495
         Width           =   1365
      End
   End
   Begin SICMACT.FlexEdit fefacturadet 
      Height          =   4050
      Left            =   75
      TabIndex        =   0
      Top             =   1860
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   7144
      Cols0           =   14
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Item-Contrato-Pieza-Tipo-Material-Observacion-PesoNeto-Importe-ComMart-Tasacion-Estado-NumRemate-TipoProceso-IGV"
      EncabezadosAnchos=   "400-1800-500-1100-1100-2000-800-1000-0-0-0-0-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-L-L-R-R-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-2-0-0-0-0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "FrmPigFacturaRemate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim lnPorcComMart As Double
'Dim lnRemate As Integer
'
'Private Sub cmdBuscar_Click()
'Dim loPers As UPersona
'Dim lsPersCod As String, lsPersNombre As String
'Dim lsEstados As String
'Dim lspieza As Integer
'Dim loPersContrato As DColPContrato
'Dim loPersCredito As dPigContrato
'Dim lrContratos As ADODB.Recordset
'Dim lsDocId As String
'Dim lsDocJur As String
'Dim lsDirPers As String
'Dim loCuentas As UProdPersona
'Dim I As Integer
'Dim liEvalCli As Integer
'
'    On Error GoTo ControlError
'
'    Set loPers = New UPersona
'        Set loPers = frmBuscaPersona.Inicio
'        If Not loPers Is Nothing Then
'            lsPersCod = loPers.sPersCod
'            lsPersNombre = PstaNombre(loPers.sPersNombre)
'            lsDocId = loPers.sPersIdnroDNI
'            lsDocJur = loPers.sPersIdnroRUC
'            lsDirPers = loPers.sPersDireccDomicilio
'        End If
'    Set loPers = Nothing
'    txtPersCod = lsPersCod
'    txtnombre.Text = lsPersNombre
'    txtDocId = lsDocId
'    txtDocJur = lsDocJur
'    txtDirPers = lsDirPers
'
'    Exit Sub
'
'ControlError:
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'End Sub
'
'
'Private Sub cmdGrabar_Click()
'Dim oCont As NContFunciones
'Dim lsMovNro As String
'Dim lsFechaHoraGrab As String
'Dim oGraba As NPigRemate
'Dim rs As Recordset
'Dim lsPoliza As String
'
'Dim lsCuenta As String
'Dim I As Integer
'Dim lspieza As Integer
'Dim lnMonto As Currency
'Dim lnCapital As Currency, lnComision As Currency
'Dim lnInteresComp As Currency, lnImpuesto As Currency
'Dim lnMontoEntregar As Currency
'Dim lnConcepto As Integer
'Dim oImpre As NPigImpre
'Dim oPrevio As Previo.clsPrevio
'Dim lsImpre As String
'Dim lsDocId As String
'
'If txtPersCod = "" Then
'    MsgBox "Debe ingresar los datos del Cliente", vbInformation, "Aviso"
'    Exit Sub
'End If
'
'If MsgBox(" Grabar Facturacion de Piezas ? ", vbYesNo + vbQuestion + vbDefaultButton1, "Aviso") = vbYes Then
'
'    CmdGrabar.Enabled = False
'    lsPoliza = Trim(txtSerie) & Trim(txtNumDoc)
'
'    Set oCont = New NContFunciones
'    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'    Set oCont = Nothing
'
'    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
'    Set rs = fefacturadet.GetRsNew
'
'    Set oGraba = New NPigRemate
'    oGraba.nPigFacturaVentaRemate lsMovNro, lsFechaHoraGrab, gPigOpeVentaRemate, rs, lsPoliza, CCur(txtTotal), '                txtPersCod, lnRemate
'    Set oGraba = Nothing
'
'    '***** IMPRESION DE LA POLIZA
'    Set oImpre = New NPigImpre
'    If txtDocId <> "" Then
'        lsDocId = txtDocId
'    Else
'        lsDocId = txtDocJur
'    End If
'
'    Set rs = fefacturadet.GetRsNew
'
'    lsImpre = oImpre.ImprePoliza(lsPoliza, txtPersCod, txtnombre, txtDirPers, gdFecSis, txtSubTotal, txtComMart, txtIGV, txtTotal, lsDocId, rs, lsPoliza)
'    Set oImpre = Nothing
'
'    Set oPrevio = New Previo.clsPrevio
'    oPrevio.Show lsImpre, "Poliza", True, 66
'    Set oPrevio = Nothing
'
'    Limpiar
'    FrmPigClienteRemate.Inicia
'
'Else
'    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
'End If
'
'Exit Sub
'
'ControlError:
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'Dim nDatosPiezas As Recordset
'Dim psNombre As String
'Dim I As Integer
'Dim lnTVenta As Currency, lnTComMart As Currency, lnTIGV As Currency
'Dim oParam As DPigFunciones
'Dim lsSerPoliza As String
'Dim lsNumPoliza As String
'
'    Set oParam = New DPigFunciones
'    'Busca Numero de Poliza
'    lsSerPoliza = oParam.GetSerDocumento(Val(gsCodAge), gPigTipoPoliza)
'    lsNumPoliza = oParam.GetNumDocumento(gPigTipoPoliza, lsSerPoliza)
'    txtSerie = lsSerPoliza
'    txtNumDoc = lsNumPoliza
'    'Comision del Martillero
'    lnPorcComMart = oParam.GetParamValor(gPigParamPorcComisMartillero)
'
'    Set oParam = Nothing
'    'LLama al cliente de la pantalla anterior
'    lbalias.Caption = FrmPigClienteRemate.feclienteremate.TextMatrix(FrmPigClienteRemate.feclienteremate.Row, 1)
'    psNombre = lbalias.Caption
'
'    Set nDatosPiezas = FrmPigClienteRemate.fepiezasrem.GetRsNew
'    lnTVenta = 0: lnTComMart = 0: lnTIGV = 0
'    fefacturadet.Clear
'    fefacturadet.Rows = 2
'    fefacturadet.FormaCabecera
'
'    nDatosPiezas.MoveFirst
'    Do While Not (nDatosPiezas.EOF)
'         fefacturadet.AdicionaFila
'         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 1) = nDatosPiezas!Contrato
'         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 2) = nDatosPiezas!Pieza
'         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 3) = nDatosPiezas!Tipo
'         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 4) = nDatosPiezas!Material
'         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 5) = nDatosPiezas!Observacion
'         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 6) = nDatosPiezas!pNeto
'         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 7) = Format(nDatosPiezas!pVenta, "#####,###.00")
'         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 8) = (CCur(nDatosPiezas!pVenta) * lnPorcComMart / 100)
'         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 13) = (CCur(nDatosPiezas!pVenta) * lnPorcComMart / 100) * 0.18
'         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 9) = nDatosPiezas!Tasacion
'         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 10) = nDatosPiezas!Estado
'         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 11) = nDatosPiezas!NumRemate
'         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 12) = nDatosPiezas!TipoProceso
'         lnTVenta = lnTVenta + nDatosPiezas!pVenta
'         lnTComMart = lnTComMart + (CCur(nDatosPiezas!pVenta) * lnPorcComMart / 100)
'         lnTIGV = lnTIGV + ((CCur(nDatosPiezas!pVenta) * lnPorcComMart / 100) * 0.18)
'         lnRemate = nDatosPiezas!NumRemate
'         nDatosPiezas.MoveNext
'     Loop
'     Set nDatosPiezas = Nothing
'
'     txtSubTotal = Format(lnTVenta, "#####,###.00")
'     txtComMart = Format(lnTComMart, "#####,###.00")
'     txtIGV = Format(lnTIGV, "######,###.00")
'     txtTotal.Text = Format((CCur(txtSubTotal.Text) + CCur(txtComMart.Text) + CCur(txtIGV.Text)), "#####,###.00")
'
'End Sub
'
'Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = 13 Then
'        txtNumDoc = Right("00000000" & CStr(CLng(txtNumDoc)), 8)
'    End If
'
'End Sub
'
'Private Sub Limpiar()
'
'    txtDirPers = ""
'    txtComMart = ""
'    txtDocId = ""
'    txtDocJur = ""
'    txtIGV = ""
'    txtnombre = ""
'    txtNumDoc = ""
'    txtPersCod = ""
'    txtSerie = ""
'    txtSubTotal = ""
'    txtTotal = ""
'    fefacturadet.Clear
'    fefacturadet.FormaCabecera
'    lnRemate = 0
'
'End Sub
