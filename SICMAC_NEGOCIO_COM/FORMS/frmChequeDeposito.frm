VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChequeDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DEPOSITO DE CHEQUES"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10275
   Icon            =   "frmChequeDeposito.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5490
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   9684
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cheques"
      TabPicture(0)   =   "frmChequeDeposito.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feCheque"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraVoucher"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkTodos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSalir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OptMoneda(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "OptMoneda(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdGrabar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   345
         Left            =   8760
         TabIndex        =   10
         Top             =   4560
         Width           =   1290
      End
      Begin VB.OptionButton OptMoneda 
         Caption         =   "&Dolares"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   9060
         TabIndex        =   9
         Top             =   420
         Width           =   975
      End
      Begin VB.OptionButton OptMoneda 
         Caption         =   "&Soles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   8040
         TabIndex        =   8
         Top             =   420
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   345
         Left            =   8760
         TabIndex        =   7
         Top             =   4920
         Width           =   1290
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         Caption         =   "Todos"
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
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   555
         Width           =   855
      End
      Begin VB.Frame fraVoucher 
         Caption         =   "Depósito de Cheque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1050
         Left            =   80
         TabIndex        =   1
         Top             =   4350
         Width           =   8610
         Begin SICMACT.TxtBuscar txtBcoCtaIFCod 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   270
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   503
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            sTitulo         =   ""
         End
         Begin VB.Label lblBcoNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   8295
         End
         Begin VB.Label lblBcoCtaIFDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3240
            TabIndex        =   3
            Top             =   270
            Width           =   5175
         End
      End
      Begin SICMACT.FlexEdit feCheque 
         Height          =   3450
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   6085
         Cols0           =   8
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-nID--N° de Cheque-Girador-Banco Emisor-Monto-Aux"
         EncabezadosAnchos=   "0-0-400-2500-2700-2700-1000-0"
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
         ColumnasAEditar =   "X-X-2-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-4-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-L-L-L-R-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmChequeDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'** Nombre : frmChequeDeposito
'** Descripción : Para Deposito de Cheques creado segun TI-ERS126-2013
'** Creación : EJVG, 20130124 11:00:00 AM
'**********************************************************************
Option Explicit

Private Sub Form_Load()
    optMoneda_Click (0)
End Sub
Private Sub optMoneda_Click(Index As Integer)
    Dim obj As New clases.DOperacion
    Dim lnMoneda As Integer
    Dim lsColor As String
    
    On Error GoTo ErrOptMoneda
    Screen.MousePointer = 11
    txtBcoCtaIFCod.psRaiz = "Cuentas de Instituciones Financieras"
    txtBcoCtaIFCod.Text = ""
    lblBcoCtaIFDesc.Caption = ""
    lblBcoNombre.Caption = ""
    chkTodos.value = 0
    If optMoneda(0).value = True Then
        lnMoneda = 1
        lsColor = &H80000005
    Else
        lnMoneda = 2
        lsColor = &HC0FFC0
    End If
    txtBcoCtaIFCod.BackColor = lsColor
    lblBcoCtaIFDesc.BackColor = lsColor
    lblBcoNombre.BackColor = lsColor
    txtBcoCtaIFCod.rs = obj.listarCuentasEntidadesFinacieras("_1_[12]" & CStr(lnMoneda) & "%", CStr(lnMoneda))
    cargar_cheques
    Screen.MousePointer = 0
    Exit Sub
ErrOptMoneda:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub txtBcoCtaIFCod_EmiteDatos()
    Dim oNCajaCtaIF As New clases.NCajaCtaIF
    Dim oDOperacion As New clases.DOperacion
    
    lblBcoNombre.Caption = ""
    lblBcoCtaIFDesc.Caption = ""
    If txtBcoCtaIFCod.Text <> "" Then
        lblBcoNombre.Caption = oNCajaCtaIF.NombreIF(Mid(txtBcoCtaIFCod.Text, 4, 13))
        lblBcoCtaIFDesc.Caption = oDOperacion.recuperaTipoCuentaEntidadFinaciera(Mid(txtBcoCtaIFCod.Text, 18, 10)) & " " & txtBcoCtaIFCod.psDescripcion
    End If
    Set oNCajaCtaIF = Nothing
    Set oDOperacion = Nothing
End Sub
Private Function cargar_cheques() As Boolean
    Dim oDR As New NCOMDocRec
    Dim oRs As New ADODB.Recordset
    Dim row As Long
    Dim lnMoneda As Integer
    
    lnMoneda = IIf(optMoneda(0).value, 1, 2)
    Set oRs = oDR.ListaChequexDeposito(Right(gsCodAge, 2), lnMoneda)
    FormateaFlex feCheque
    If Not oRs.EOF Then
        Do While Not oRs.EOF
            feCheque.AdicionaFila
            row = feCheque.row
            feCheque.TextMatrix(row, 1) = oRs!nId
            feCheque.TextMatrix(row, 3) = oRs!cNroDoc
            feCheque.TextMatrix(row, 4) = oRs!cGiradorNombre
            feCheque.TextMatrix(row, 5) = oRs!cIFiNombre
            feCheque.TextMatrix(row, 6) = Format(oRs!nMonto, gsFormatoNumeroView)
            oRs.MoveNext
        Loop
        cargar_cheques = True
    Else
        cargar_cheques = False
    End If
    Set oDR = Nothing
    Set oRs = Nothing
End Function
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub chkTodos_Click()
    Dim row As Long
    Dim lsCheck As String
    
    If feCheque.TextMatrix(1, 0) = "" Then
        chkTodos.value = 0
        Exit Sub
    End If
    If chkTodos.value = 1 Then
        lsCheck = "1"
    Else
        lsCheck = ""
    End If
    If feCheque.TextMatrix(1, 0) <> "" Then
        For row = 1 To feCheque.Rows - 1
            feCheque.TextMatrix(row, 2) = lsCheck
        Next
    End If
End Sub
Private Sub txtBcoCtaIFCod_LostFocus()
    If txtBcoCtaIFCod.Text = "" Then
        lblBcoNombre.Caption = ""
        lblBcoCtaIFDesc.Caption = ""
    End If
End Sub
'Private Sub CmdGrabar_Click()
'    Dim oDocRec As NCOMDocRec
'    Dim oCont As COMNContabilidad.NCOMContFunciones
'    Dim bExito As Boolean
'    Dim lsMovNro As String
'    Dim row As Long
'    Dim MatDatos() As Long
'    Dim i As Long
'    On Error GoTo ErrCmdGrabar
'
'    If feCheque.TextMatrix(1, 0) = "" Then
'        MsgBox "No existen cheques para realizar depósito", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    ReDim MatDatos(0)
'    For row = 1 To feCheque.Rows - 1
'        If feCheque.TextMatrix(row, 2) = "." Then
'            i = UBound(MatDatos) + 1
'            ReDim Preserve MatDatos(i)
'            MatDatos(i) = CLng(feCheque.TextMatrix(row, 1))
'        End If
'    Next
'    If UBound(MatDatos) = 0 Then
'        MsgBox "Ud. debe seleccionar un registro por lo menos para continuar", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    If Len(Trim(txtBcoCtaIFCod.Text)) = 0 Then
'        MsgBox "Ud. debe seleccionar la Cuenta a la que se realizó el Deposito", vbInformation, "Aviso"
'        If txtBcoCtaIFCod.Visible And txtBcoCtaIFCod.Enabled Then txtBcoCtaIFCod.SetFocus
'        Exit Sub
'    End If
'
'    If MsgBox("¿Esta seguro de Registrar los depósitos de Cheque?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
'    Screen.MousePointer = 11
'    Set oDocRec = New NCOMDocRec
'    Set oCont = New COMNContabilidad.NCOMContFunciones
'    lsMovNro = oCont.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
'    bExito = oDocRec.RegistrarDepositoCheque(lsMovNro, MatDatos, gdFecSis, Left(txtBcoCtaIFCod.Text, 2), Mid(txtBcoCtaIFCod.Text, 4, 13), Mid(txtBcoCtaIFCod.Text, 18, Len(txtBcoCtaIFCod.Text)))
'    If bExito Then
'        MsgBox "Se ha registrado con éxito el Depósito de Cheque", vbInformation, "Aviso"
'        For i = 1 To UBound(MatDatos)
'            feCheque.EliminaFila MatDatos(i)
'        Next
'    Else
'        MsgBox "Ha ocurrido un error al realizar la operación, si el mismo persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
'    End If
'    Screen.MousePointer = 0
'    Set oDocRec = Nothing
'    Set oCont = Nothing
'
'    Exit Sub
'ErrCmdGrabar:
'    Screen.MousePointer = 0
'    MsgBox err.Description, vbCritical, "Aviso"
'End Sub
Private Sub cmdGrabar_Click()
    Dim oDocRec As NCOMDocRec
    Dim oCont As NContFunciones
    Dim oSis As NConstSistemas
    Dim oImpre As NContImprimir
    Dim bExito As Boolean
    Dim row As Long
    Dim MatDatos() As Long
    Dim i As Long
    Dim lsCtaContDepositoD As String, lsCtaContDepositoH As String
    Dim lsIFTpo As String, lsPersCod As String, lsCtaIFCod As String
    Dim lsMovNroImpre As String, lsCadImpre As String
    Dim MatMovNro() As String
    Dim lnMoneda As Moneda
    Dim lsBcoCtaCont As String, lsBcoSubCtaCont As String
    Dim oPrevio As clsprevio
    
    On Error GoTo ErrCmdGrabar
    If feCheque.TextMatrix(1, 0) = "" Then
        MsgBox "No existen cheques para realizar depósito", vbInformation, "Aviso"
        Exit Sub
    End If
    ReDim MatDatos(0)
    For row = 1 To feCheque.Rows - 1
        If feCheque.TextMatrix(row, 2) = "." Then
            i = UBound(MatDatos) + 1
            ReDim Preserve MatDatos(i)
            MatDatos(i) = CLng(feCheque.TextMatrix(row, 1))
        End If
    Next
    If UBound(MatDatos) = 0 Then
        MsgBox "Ud. debe seleccionar un registro por lo menos para continuar", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtBcoCtaIFCod.Text)) = 0 Then
        MsgBox "Ud. debe seleccionar la Cuenta a la que se realizó el Deposito", vbInformation, "Aviso"
        If txtBcoCtaIFCod.Visible And txtBcoCtaIFCod.Enabled Then txtBcoCtaIFCod.SetFocus
        Exit Sub
    End If
    lsIFTpo = Mid(txtBcoCtaIFCod.Text, 1, 2)
    lsPersCod = Mid(txtBcoCtaIFCod.Text, 4, 13)
    lsCtaIFCod = Mid(txtBcoCtaIFCod.Text, 18, Len(txtBcoCtaIFCod.Text))
    lnMoneda = IIf(optMoneda(0).value = True, gMonedaNacional, gMonedaExtranjera)
    
    Set oSis = New NConstSistemas
    Set oCont = New NContFunciones
    lsBcoCtaCont = "11" & CStr(lnMoneda) & IIf(lsPersCod = "1090100822183", "2", "3") & lsIFTpo
    lsBcoSubCtaCont = oCont.GetFiltroObjetos(1, lsBcoCtaCont, txtBcoCtaIFCod.Text, False)
    If Len(lsBcoSubCtaCont) = 0 Then
        MsgBox "Esta cuenta contable " & lsBcoCtaCont & " no esta registrado en CtaIFFiltro, comunicarse con el Dpto. de TI", vbInformation, "Aviso"
        Exit Sub
    End If
    lsCtaContDepositoD = lsBcoCtaCont & lsBcoSubCtaCont
    lsCtaContDepositoH = oSis.LeeConstSistema(468)
    lsCtaContDepositoH = Replace(Replace(lsCtaContDepositoH, "M", lnMoneda), "AG", Right(gsCodAge, 2))
    Set oSis = Nothing
    If Not oCont.verificarUltimoNivelCta(lsCtaContDepositoD) Then
       MsgBox "La Cuenta Contable " & lsCtaContDepositoD & " no es de Ultimo Nivel, comunicarse con el Dpto. de Contabilidad", vbInformation, "Aviso"
       Set oCont = Nothing
       Exit Sub
    End If
    If Not oCont.verificarUltimoNivelCta(lsCtaContDepositoH) Then
       MsgBox "La Cuenta Contable " & lsCtaContDepositoH & " no es de Ultimo Nivel, comunicarse con el Dpto. de Contabilidad", vbInformation, "Aviso"
       Set oCont = Nothing
       Exit Sub
    End If
    Set oCont = Nothing

    If MsgBox("¿Esta seguro de registrar los Depósitos de Cheque?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    cmdGrabar.Enabled = False
    Screen.MousePointer = 11
    Set oDocRec = New NCOMDocRec
    bExito = oDocRec.RegistrarDepositoCheque(MatDatos, lnMoneda, lsIFTpo, lsPersCod, lsCtaIFCod, gdFecSis, Right(gsCodAge, 2), gsCodUser, lsCtaContDepositoD, lsCtaContDepositoH, lsMovNroImpre)
    If bExito Then
        MsgBox "Se ha registrado con éxito el Depósito de Cheque", vbInformation, "Aviso"
        'Set oImpre = New NContImprimir
        'MatMovNro = Split(lsMovNroImpre, ",")
        'For i = 0 To UBound(MatMovNro)
        '    lsCadImpre = lsCadImpre & oImpre.ImprimeAsientoContable(MatMovNro(i), gnLinPage, gnColPage, "DEPÓSITO DE CHEQUE", , "179") & oImpresora.gPrnSaltoPagina
        'Next
        'Set oImpre = Nothing
        'Set oPrevio = New clsprevio
        'oPrevio.Show lsCadImpre, "DEPÓSITO DE CHEQUE", False, gnLinPage
        'Set oPrevio = Nothing
        optMoneda_Click (0)
    Else
        MsgBox "Ha ocurrido un error al realizar la operación, si el mismo persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    Screen.MousePointer = 0
    cmdGrabar.Enabled = True
    Set oDocRec = Nothing
    Exit Sub
ErrCmdGrabar:
    Screen.MousePointer = 0
    cmdGrabar.Enabled = True
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
