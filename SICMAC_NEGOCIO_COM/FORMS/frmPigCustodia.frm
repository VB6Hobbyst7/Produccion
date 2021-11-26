VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPigCustodia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rescate de Joya"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "frmPigCustodia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4875
      TabIndex        =   5
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6090
      TabIndex        =   4
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3675
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame fraContenedor 
      Height          =   6150
      Index           =   1
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   7185
      Begin VB.Frame Frame1 
         Height          =   1410
         Left            =   75
         TabIndex        =   9
         Top             =   4665
         Width           =   7035
         Begin VB.Label lblNumDuplicado 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1695
            TabIndex        =   21
            Top             =   1005
            Width           =   675
         End
         Begin VB.Label Label5 
            Caption         =   "Nro. Duplicado"
            Height          =   240
            Left            =   150
            TabIndex        =   20
            Top             =   1035
            Width           =   1350
         End
         Begin VB.Label lblPiezas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1695
            TabIndex        =   19
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label6 
            Caption         =   "Total de Piezas"
            Height          =   225
            Left            =   135
            TabIndex        =   18
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Cancelación"
            Height          =   240
            Left            =   135
            TabIndex        =   15
            Top             =   675
            Width           =   1470
         End
         Begin VB.Label lblFecCancel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1695
            TabIndex        =   14
            Top             =   615
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "Dias Custodia"
            Height          =   240
            Left            =   4110
            TabIndex        =   13
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label lblDiasCustodia 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5445
            TabIndex        =   12
            Top             =   225
            Width           =   630
         End
         Begin VB.Label Label2 
            Caption         =   "Costo Custodia"
            Height          =   210
            Left            =   4095
            TabIndex        =   11
            Top             =   675
            Width           =   1215
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5445
            TabIndex        =   10
            Top             =   615
            Width           =   1275
         End
      End
      Begin VB.Frame fraContenedor 
         Height          =   3975
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   690
         Width           =   7035
         Begin SICMACT.FlexEdit feJoyas 
            Height          =   2880
            Left            =   60
            TabIndex        =   16
            Top             =   1005
            Width           =   6930
            _ExtentX        =   12224
            _ExtentY        =   5080
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "Item-Tipo-SubTipo-Material-Observacion"
            EncabezadosAnchos=   "400-1200-1100-1150-2900"
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
            TextArray0      =   "Item"
            SelectionMode   =   1
            lbUltimaInstancia=   -1  'True
            lbPuntero       =   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin MSComctlLib.ListView lstClientes 
            Height          =   810
            Left            =   60
            TabIndex        =   7
            Top             =   150
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   1429
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483627
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Codigo"
               Object.Width           =   2470
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cliente"
               Object.Width           =   5468
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Doc Ident."
               Object.Width           =   1765
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Tipo Cliente"
               Object.Width           =   2293
            EndProperty
         End
         Begin VB.Label Label4 
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
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   120
            TabIndex        =   17
            Top             =   1170
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   405
         Left            =   6480
         Picture         =   "frmPigCustodia.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Buscar ..."
         Top             =   225
         Width           =   465
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   165
         TabIndex        =   2
         Top             =   270
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin VB.Label lblTasacion 
      Height          =   210
      Left            =   885
      TabIndex        =   8
      Top             =   3630
      Visible         =   0   'False
      Width           =   1410
   End
End
Attribute VB_Name = "frmPigCustodia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Sub cmdBuscar_Click()
'Dim loPers As UPersona
'Dim lsPersCod As String, lsPersNombre As String
'Dim lsEstados As String
'Dim loPersContrato As DColPContrato
'Dim lrContratos As ADODB.Recordset
'Dim loCuentas As UProdPersona
'
'On Error GoTo ControlError
'
'Set loPers = New UPersona
'    Set loPers = frmBuscaPersona.Inicio
'    If Not loPers Is Nothing Then
'        lsPersCod = loPers.sPersCod
'        lsPersNombre = loPers.sPersNombre
'    Else
'        Set loPers = Nothing
'        Exit Sub
'    End If
'
'lsEstados = gPigEstCancelPendRes
'
'If Trim(lsPersCod) <> "" Then
'    Set loPersContrato = New DColPContrato
'        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
'    Set loPersContrato = Nothing
'End If
'
'Set loCuentas = New UProdPersona
'    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos)
'    If loCuentas.sCtaCod <> "" Then
'        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
'        AXCodCta.SetFocusCuenta
'        BuscaContrato (AXCodCta.NroCuenta)
'        cmdGrabar.Enabled = True
'    End If
'
'Set loCuentas = Nothing
'If cmdGrabar.Enabled Then cmdGrabar.SetFocus
'Exit Sub
'
'ControlError:
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'End Sub
'Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
'End Sub
'
'Private Sub BuscaContrato(ByVal psNroContrato As String)
'Dim lbOk As Boolean
'Dim lrValida As Recordset, rs As Recordset
'Dim loValContrato As nPigValida
'Dim lsEstados As String
'Dim lrCredPigPersonas As Recordset
'Dim lrCredPig As Recordset
'Dim oDatos As DPigContrato
'Dim lnDiasCustodia As Integer
'Dim lnCostoCust As Currency, lnIGV As Currency
'Dim lstTmpCliente As ListItem
'
'    On Error GoTo ControlError
'
'    Set loValContrato = New nPigValida
'    'lsEstados = gPigEstCancelPendRes & "," & gPigEstAnula
'    lsEstados = gPigEstCancelPendRes
'    Set lrValida = loValContrato.nValidaComision(psNroContrato, lsEstados)
'    If Not (lrValida.EOF And lrValida.BOF) Then
'        lblTasacion = lrValida!Tasacion
'        Set loValContrato = Nothing
'    Else
'        Set loValContrato = Nothing
'        Exit Sub
'    End If
'
'    If lrValida Is Nothing Then
'        Set lrValida = Nothing
'        Exit Sub
'    End If
'
'     'Mostrar Clientes
'     Set oDatos = New DPigContrato
'         Set lrCredPigPersonas = oDatos.dObtieneDatosCreditoPignoraticioPersonas(psNroContrato)
'
'     If Not (lrCredPigPersonas.BOF And lrCredPigPersonas.EOF) Then
'        lstClientes.ListItems.Clear
'        Set lstTmpCliente = lstClientes.ListItems.Add(, , Trim(lrCredPigPersonas!cPersCod))
'        lstTmpCliente.SubItems(1) = Trim(PstaNombre(lrCredPigPersonas!cPersNombre, False))
'        lstTmpCliente.SubItems(2) = Trim(IIf(IsNull(lrCredPigPersonas!NroDNI), "", lrCredPigPersonas!NroDNI))
'        lstTmpCliente.SubItems(3) = Trim(IIf(IsNull(lrCredPigPersonas!DescCalif), "", lrCredPigPersonas!DescCalif))
'    Else
'        Exit Sub
'    End If
'
'    Set lrCredPigPersonas = Nothing
'
'    Set rs = oDatos.dObtieneDetalleJoyas(psNroContrato)
'
'    Do While Not rs.EOF
'        feJoyas.AdicionaFila
'        feJoyas.TextMatrix(rs!Item, 0) = rs!Item
'        feJoyas.TextMatrix(rs!Item, 1) = rs!Tipo
'        If rs!nSubTipo <> 0 Or Not IsNull(rs!nSubTipo) Then
'            feJoyas.TextMatrix(rs!Item, 2) = rs!SubTipo
'        End If
'        feJoyas.TextMatrix(rs!Item, 3) = rs!Material
'        feJoyas.TextMatrix(rs!Item, 4) = IIf(IsNull(rs!Observacion), "", rs!Observacion)
'        rs.MoveNext
'    Loop
'    Set rs = Nothing
'
'    Set lrCredPig = oDatos.dObtieneCreditoPigno(psNroContrato)
'
'    lblPiezas = lrCredPig!npiezas
'    lblFecCancel = Format(lrCredPig!dPrdEstado, "dd/mm/yyyy") ' CAMBIO CMCPL
'    lnDiasCustodia = DateDiff("d", Format(lrCredPig!dPrdEstado, "dd/mm/yyyy"), gdFecSis)
'    lblDiasCustodia = lnDiasCustodia
'    lblNumDuplicado = IIf(IsNull(lrCredPig!nNroDuplic), 0, lrCredPig!nNroDuplic)
'
'    lnCostoCust = Round(calculo(lnDiasCustodia, gColPigConceptoCodCustodia, gPiParamDiasMinComision), 2)
'
'    lblTotal = Format(lnCostoCust, "###,###.00")
'
'    Set lrCredPig = Nothing
'    Set oDatos = Nothing
'    AXCodCta.Enabled = False
'    cmdGrabar.Enabled = True
'    cmdGrabar.SetFocus
'
'    Exit Sub
'
'ControlError:
'    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & '        "Avise al Area de Sistemas ", vbInformation, " Aviso "
'End Sub
'Private Sub cmdCancelar_Click()
'    cmdGrabar.Enabled = False
'    Limpiar
'    cmdGrabar.Enabled = False
'    AXCodCta.Enabled = True
'    cmdBuscar.SetFocus
'End Sub
'
'Private Sub Limpiar()
'
'    AXCodCta.Age = ""
'    AXCodCta.Cuenta = ""
'    lstClientes.ListItems.Clear
'    feJoyas.Clear
'    feJoyas.FormaCabecera
'    feJoyas.Rows = 2
'    lblPiezas = ""
'    lblDiasCustodia = ""
'    lblFecCancel = ""
'    lblTotal = ""
'    lblTasacion = ""
'
'End Sub
'
'Private Sub cmdGrabar_Click()
'Dim oCont As NContFunciones
'Dim oPigGraba As NPigContrato
'Dim lsFechaHoraGrab As String
'Dim lsMovNro As String
'Dim lsCadImprimir As String
'Dim oPrevio As Previo.clsPrevio
'Dim oImpre As NPigImpre
'Dim lsCuenta As String
'Dim lsPersNombre As String
'
'lsCuenta = AXCodCta.NroCuenta
'lsPersNombre = lstClientes.ListItems(1)
'
'On Error GoTo ControlError
'If MsgBox(" Desea Rescatar la Joya ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'
'    Set oCont = New NContFunciones
'    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'    Set oCont = Nothing
'
'    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
'
'    Set oPigGraba = New NPigContrato
'    oPigGraba.nRescateLotes lsCuenta, lsMovNro, CCur(lblTotal), CCur(lblTasacion), lsFechaHoraGrab, gsCodAge
'    Set oPigGraba = Nothing
'
'    If CCur(lblTotal) > 0 Then
'
'        Set oImpre = New NPigImpre
'        lsCadImprimir = oImpre.ImpreReciboCustodia(gsNomAge, gdFecSis, lsCuenta, lsPersNombre, CCur(lblTotal), gsCodUser, "")
'        Set oImpre = Nothing
'
'        Set oPrevio = New Previo.clsPrevio
'        oPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
'
'        Do While True
'            If MsgBox("Desea Reimprimir Comprobante de Rescate? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'                oPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
'            Else
'                Set oPrevio = Nothing
'                Exit Do
'            End If
'        Loop
'
'        Set oPrevio = Nothing
'    End If
'
'    cmdGrabar.Enabled = False
'    AXCodCta.Enabled = True
'    cmdBuscar.Enabled = True
'
'    Limpiar
'End If
'
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'End Sub
'
'Private Sub cmdSalir_Click()
'Unload Me
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
'    AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
'    AXCodCta.Age = ""
'    Me.Icon = LoadPicture(App.path & "\bmps\cm.ico")
'End Sub
'
