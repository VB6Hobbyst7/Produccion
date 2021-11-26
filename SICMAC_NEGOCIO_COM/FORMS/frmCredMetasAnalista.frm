VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCredMetasAnalista 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Metas de Analista"
   ClientHeight    =   6945
   ClientLeft      =   855
   ClientTop       =   1725
   ClientWidth     =   10665
   Icon            =   "frmCredMetasAnalista.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmAnalista 
      Caption         =   "Analista"
      Height          =   2145
      Left            =   1950
      TabIndex        =   10
      Top             =   0
      Width           =   6675
      Begin MSDataGridLib.DataGrid DGMetas 
         Height          =   1410
         Left            =   210
         TabIndex        =   17
         Top             =   600
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   2487
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cTipoMeta"
            Caption         =   "Tipo Meta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "dInicial"
            Caption         =   "Fecha Inicial"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "dFinal"
            Caption         =   "Fecha Final"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "sMoneda"
            Caption         =   "Moneda"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "nMoneda"
            Caption         =   "Moneda"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   14.74
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox CmbAnalista 
         Height          =   315
         Left            =   1035
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Seleccione y Presione Enter para Mostrar los Datos"
         Top             =   195
         Width           =   5055
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   390
      Left            =   4455
      TabIndex        =   5
      Top             =   6375
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   5340
      TabIndex        =   4
      Top             =   6375
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   6240
      TabIndex        =   3
      Top             =   6375
      Width           =   870
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   7140
      TabIndex        =   2
      Top             =   6375
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   390
      Left            =   2655
      TabIndex        =   1
      Top             =   6375
      Width           =   870
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   390
      Left            =   3570
      TabIndex        =   0
      Top             =   6375
      Width           =   870
   End
   Begin SICMACT.FlexEdit FECredProd 
      Height          =   2775
      Left            =   150
      TabIndex        =   18
      Top             =   3285
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   4895
      Cols0           =   11
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "-Producto-Meta-Codigo-Totales-Nuevos-Total Nuevos-Paralelos-Total Paralelos-Recurrentes-Total Recurrentes"
      EncabezadosAnchos=   "500-3500-1200-0-1200-1200-1200-1200-1200-1200-1200"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   1
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-R-C-R-C-R-C-R-C-R"
      FormatosEdit    =   "0-0-4-0-4-4-3-4-3-4-3"
      AvanceCeldas    =   1
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame frmDetalle 
      Enabled         =   0   'False
      Height          =   4095
      Left            =   60
      TabIndex        =   6
      Top             =   2145
      Width           =   10575
      Begin VB.ComboBox CboMoneda 
         Height          =   315
         Left            =   7845
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   510
         Width           =   1920
      End
      Begin VB.Frame Frame1 
         Height          =   900
         Left            =   1725
         TabIndex        =   12
         Top             =   120
         Width           =   5265
         Begin MSMask.MaskEdBox txtFecFin 
            Height          =   315
            Left            =   1830
            TabIndex        =   13
            Top             =   525
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFecIni 
            Height          =   315
            Left            =   1830
            TabIndex        =   14
            Top             =   180
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicial"
            Height          =   195
            Left            =   90
            TabIndex        =   16
            Top             =   225
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Final"
            Height          =   195
            Left            =   90
            TabIndex        =   15
            Top             =   555
            Width           =   825
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo de Meta"
         Height          =   900
         Left            =   135
         TabIndex        =   7
         Top             =   120
         Width           =   1470
         Begin VB.OptionButton optMetaMensual 
            Caption         =   "Mensual"
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   9
            Top             =   255
            Value           =   -1  'True
            Width           =   1080
         End
         Begin VB.OptionButton optMetaMensual 
            Caption         =   "Anual"
            Height          =   255
            Index           =   1
            Left            =   210
            TabIndex        =   8
            Top             =   540
            Width           =   975
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda :"
         Height          =   210
         Left            =   7110
         TabIndex        =   19
         Top             =   540
         Width           =   705
      End
   End
   Begin VB.Label LblBuscar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Buscando ........."
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
      Height          =   270
      Left            =   90
      TabIndex        =   21
      Top             =   6450
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "frmCredMetasAnalista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RCab As ADODB.Recordset
Private RDet As ADODB.Recordset
Private nTipoAct As Integer
Private bValidando As Boolean

'Para el tema de los Componentes
Private MatProdCred As Variant
Private bCargoAnalistaDet As Boolean

Private Function ValidaDatos() As Boolean
Dim nFilaDG As Integer
Dim oCred As COMDCredito.DCOMCredito

    ValidaDatos = True
    If ValidaFecha(txtFecIni.Text) <> "" Then
        MsgBox ValidaFecha(txtFecIni.Text), vbInformation, "Aviso"
        ValidaDatos = False
        txtFecIni.SetFocus
        Exit Function
    End If
    
    If ValidaFecha(Me.txtFecFin.Text) <> "" Then
        MsgBox ValidaFecha(txtFecFin.Text), vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    If nTipoAct = 1 Then
        If Not RCab Is Nothing Then
            If RCab.RecordCount > 0 Then
                Set oCred = New COMDCredito.DCOMCredito
                If oCred.ExisteMetaAnalista(RCab!cPersCod, CDate(txtFecIni.Text), CDate(txtFecFin.Text), CInt(Right(CboMoneda.Text, 1)), IIf(Me.optMetaMensual(0).value, 1, 2)) Then
                
                    MsgBox "Metas para esa fecha ya existen", vbInformation, "Aviso"
                    ValidaDatos = False
                    Exit Function
                Else
                    ValidaDatos = True
                    Exit Function
                End If
                Set oCred = Nothing
            End If
        End If
    End If
End Function

Private Sub LimpiaPantalla()
Dim i As Integer
    i = Me.CmbAnalista.ListIndex
    optMetaMensual(0).value = True
    txtFecIni.Text = "__/__/____"
    txtFecFin.Text = "__/__/____"
    LimpiaControles Me, , True
    Me.CmbAnalista.ListIndex = i
    LimpiaFlex FECredProd
    Call CargaProdCred
End Sub

Private Sub CargaProdCred()
'Dim oCredDatos As COMDCredito.DCOMCredito
'Dim R As ADODB.Recordset
Dim i As Integer
    
    LimpiaFlex FECredProd
'    Set oCredDatos = New COMDCredito.DCOMCredito
'    Set R = oCredDatos.RecuperaProductosDeCredito
'    Set oCredDatos = Nothing
'    Do While Not R.EOF
'        FECredProd.AdicionaFila
'        FECredProd.TextMatrix(R.Bookmark, 0) = Trim(Str(R.Bookmark))
'        FECredProd.TextMatrix(R.Bookmark, 1) = Trim(R!cConsDescripcion)
'        FECredProd.TextMatrix(R.Bookmark, 2) = "0.00"
'        FECredProd.TextMatrix(R.Bookmark, 3) = Trim(Str(R!nConsValor))
'        FECredProd.TextMatrix(R.Bookmark, 4) = "0.00"
'        FECredProd.TextMatrix(R.Bookmark, 5) = "0.00"
'        FECredProd.TextMatrix(R.Bookmark, 6) = "0.00"
'        FECredProd.TextMatrix(R.Bookmark, 7) = "0.00"
'        R.MoveNext
'    Loop
'    R.Close
i = 1
Do While i <= UBound(MatProdCred)
    FECredProd.AdicionaFila
    FECredProd.TextMatrix(i, 0) = MatProdCred(i, 0)
    FECredProd.TextMatrix(i, 1) = MatProdCred(i, 1)
    FECredProd.TextMatrix(i, 2) = MatProdCred(i, 2)
    FECredProd.TextMatrix(i, 3) = MatProdCred(i, 3)
    FECredProd.TextMatrix(i, 4) = MatProdCred(i, 4)
    FECredProd.TextMatrix(i, 5) = MatProdCred(i, 5)
    FECredProd.TextMatrix(i, 6) = MatProdCred(i, 6)
    FECredProd.TextMatrix(i, 7) = MatProdCred(i, 7)
    
    '20060504
    'modificado para mejorar la estetica del formulario
    FECredProd.TextMatrix(i, 8) = MatProdCred(i, 8)
    FECredProd.TextMatrix(i, 9) = MatProdCred(i, 9)
    FECredProd.TextMatrix(i, 10) = MatProdCred(i, 10)
    i = i + 1
Loop

End Sub

Private Sub CargaControles()
Dim oDCred As COMDCredito.DCOMCredito
'Dim R As ADODB.Recordset
Dim rsAnalis As ADODB.Recordset
Dim rsMoneda As ADODB.Recordset
Dim rsProCred As ADODB.Recordset

Set oDCred = New COMDCredito.DCOMCredito

Call oDCred.CargarControlesMetaAnalista(rsAnalis, rsMoneda, rsProCred)

Set oDCred = Nothing

    'Carga Analistas
'    Set oDCred = New COMDCredito.DCOMCredito
'    Set R = oDCred.CargaAnalistas
'    Set oDCred = Nothing
    Do While Not rsAnalis.EOF
        CmbAnalista.AddItem PstaNombre(rsAnalis!cPersNombre) & Space(150) & Trim(rsAnalis!cPersCod)
        rsAnalis.MoveNext
    Loop
    rsAnalis.Close
    Set rsAnalis = Nothing
    
Call Llenar_Combo_con_Recordset(rsMoneda, CboMoneda)
'    Call CargaComboConstante(gMoneda, CboMoneda)
    CboMoneda.ListIndex = 0
    
'    Call CargaProdCred
'redimiensionamos esta matriz
'de 8 a 11
ReDim MatProdCred(rsProCred.RecordCount, 11)

While Not rsProCred.EOF
    MatProdCred(rsProCred.Bookmark, 0) = Trim(Str(rsProCred.Bookmark))
    MatProdCred(rsProCred.Bookmark, 1) = Trim(rsProCred!cConsDescripcion)
    MatProdCred(rsProCred.Bookmark, 2) = "0.00"
    MatProdCred(rsProCred.Bookmark, 3) = Trim(Str(rsProCred!nConsValor))
    MatProdCred(rsProCred.Bookmark, 4) = "0.00"
    MatProdCred(rsProCred.Bookmark, 5) = "0.00"
    MatProdCred(rsProCred.Bookmark, 6) = "0.00"
    MatProdCred(rsProCred.Bookmark, 7) = "0.00"
    
    '20060504
    'mejoramos estetica
    MatProdCred(rsProCred.Bookmark, 8) = "0.00"
    MatProdCred(rsProCred.Bookmark, 9) = "0.00"
    MatProdCred(rsProCred.Bookmark, 10) = "0.00"
    rsProCred.MoveNext
Wend

Call CargaProdCred

    If CmbAnalista.ListCount > 0 Then
        CmbAnalista.ListIndex = 0
        'cmbAnalista_KeyPress 13
    End If
End Sub

Private Sub HabilitaActualizacion(ByVal pbHabilita As Boolean)
    frmAnalista.Enabled = Not pbHabilita
    frmDetalle.Enabled = pbHabilita
    cmdNuevo.Enabled = Not pbHabilita
    cmdModificar.Enabled = Not pbHabilita
    CmdGrabar.Visible = pbHabilita
    CmdCancelar.Visible = pbHabilita
    CmdSalir.Enabled = Not pbHabilita
    CmdImprimir.Enabled = Not pbHabilita
    FECredProd.lbEditarFlex = pbHabilita
End Sub

Private Sub CargaDatosAnalista(ByVal psPersCod As String)
Dim oDCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset

    On Error GoTo ErrorCargaDatosAnalista
    LimpiaPantalla
    Set oDCred = New COMDCredito.DCOMCredito
    Set RCab = Nothing
    Set RCab = oDCred.RecuperaDatosMetasAnalistaCab(psPersCod)
    Set oDCred = Nothing
    Set DGMetas.DataSource = RCab
    Call DGMetas_RowColChange(DGMetas.Row, DGMetas.Col)
    If Not RCab Is Nothing Then
        If RCab.RecordCount > 0 Then
            cmdModificar.Enabled = True
        End If
    End If
    
    If FECredProd.Rows > 2 Then
        FECredProd.TopRow = 1
    End If
    Exit Sub

ErrorCargaDatosAnalista:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub CmbAnalista_Click()
    bCargoAnalistaDet = False
    cmbAnalista_KeyPress 13
End Sub

Private Sub cmbAnalista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CargaDatosAnalista(Trim(Right(CmbAnalista.Text, 20)))
    End If
End Sub

Private Sub cmdCancelar_Click()
    HabilitaActualizacion False
End Sub

Private Sub cmdGrabar_Click()
Dim MatMontos() As String
Dim oNegCred As COMNCredito.NCOMCredito
Dim i As Integer

    If Not ValidaDatos Then
        Exit Sub
    End If
    If MsgBox("Se va a Grabar los Datos, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    ReDim MatMontos(FECredProd.Rows - 1, 3)
    For i = 1 To Me.FECredProd.Rows - 1
        MatMontos(i - 1, 0) = FECredProd.TextMatrix(i, 2)
        MatMontos(i - 1, 1) = FECredProd.TextMatrix(i, 3)
        MatMontos(i - 1, 2) = FECredProd.TextMatrix(i, 4)
    Next i
        
    Set oNegCred = New COMNCredito.NCOMCredito
    Call oNegCred.ActualizaMetasAnalista(Trim(Right(CmbAnalista.Text, 20)), IIf(optMetaMensual(0).value, 1, 2) _
                    , CDate(txtFecIni.Text), CDate(txtFecFin.Text), nTipoAct, CInt(Right(CboMoneda.Text, 1)), MatMontos)
    Set oNegCred = Nothing
    HabilitaActualizacion False
    Call CmbAnalista_Click
End Sub

Private Sub cmdImprimir_Click()
Dim oCredDoc As COMNCredito.NCOMCredDoc
Dim Prev As previo.clsPrevio

    On Error GoTo ErrorCmdImprimir_Click
    Set oCredDoc = New COMNCredito.NCOMCredDoc
    'Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
    'Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(70);  'Longitud de página a 70 líneas'
    'Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
    'Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(1);     'Tipo de Letra Sans Serif
    'Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
    
    Set Prev = New clsPrevio
    Prev.Show oCredDoc.ImprimirMetasAnalistas(gsNomAge, gdFecSis, gsCodUser, gsNomCmac), "", True
    Set Prev = Nothing
    
    Set oCredDoc = Nothing
    
    Exit Sub

ErrorCmdImprimir_Click:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub CmdModificar_Click()
    HabilitaActualizacion True
    FECredProd.SetFocus
    nTipoAct = 2
    Frame4.Enabled = False
    Frame1.Enabled = False
End Sub

Private Sub cmdNuevo_Click()
    LimpiaPantalla
    HabilitaActualizacion True
    nTipoAct = 1
    Frame4.Enabled = True
    Frame1.Enabled = True
    txtFecIni.SetFocus
    Call CargaProdCred
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub DGMetas_Click()
    'Call DGMetas_RowColChange(DGMetas.Row, DGMetas.Col)
End Sub

Private Sub DGMetas_KeyPress(KeyAscii As Integer)
    Call DGMetas_RowColChange(DGMetas.Row, DGMetas.Col)
End Sub

Private Sub DGMetas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim oDCred As COMDCredito.DCOMCredito
Dim MatMontos() As String
Dim i As Integer

    '20060406
    
    'verificamos si la fila anterior es la misma fila
    If LastRow = Empty Then Exit Sub

    '20060405
    'código de la caja que para mi ya no tiene valor pero lo dejo
    'para analizarlo mejor.
'    If bCargoAnalistaDet Then Exit Sub
'    bCargoAnalistaDet = True
    
    If bValidando Then
        Exit Sub
    End If
    'i = 0
    'Do While i <= 500
        LblBuscar.Caption = "Buscando ......."
        LblBuscar.Visible = True
        Screen.MousePointer = 11
    '     i = i + 1
    '     DoEvents
    'Loop
    If RCab.RecordCount = 0 Then
        cmdModificar.Enabled = False
        LblBuscar.Visible = False
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Call CargaProdCred
    
    
    Set oDCred = New COMDCredito.DCOMCredito
    Call oDCred.RecuperaDatosMetasAnalistaDetalle_Montos(RCab!cPersCod, RCab!nTipoMeta, RCab!dInicial, RCab!dFinal, RCab!nMoneda, RDet, MatMontos)
    Set oDCred = Nothing
    
    Do While Not RDet.EOF
        optMetaMensual(IIf(IsNull(RCab!nTipoMeta), 0, RCab!nTipoMeta - 1)).value = True
        txtFecIni.Text = Format(RCab!dInicial, "dd/mm/yyyy")
        txtFecFin.Text = Format(RCab!dFinal, "dd/mm/yyyy")
        For i = 1 To FECredProd.Rows - 1
            If Trim(FECredProd.TextMatrix(i, 3)) = Trim(Str(RDet!nTipoCred)) Then
                FECredProd.TextMatrix(i, 2) = Format(RDet!nMonto, "#0.00")
                'FECredProd.TextMatrix(i, 4) = Format(RDet!nMontoAlc, "#0.00")
                FECredProd.TextMatrix(i, 4) = Format(MatMontos(RDet.Bookmark - 1, 0), "#0.00")
                FECredProd.TextMatrix(i, 5) = Format(MatMontos(RDet.Bookmark - 1, 1), "#0.00")
                FECredProd.TextMatrix(i, 6) = Format(MatMontos(RDet.Bookmark - 1, 2), "#0")
                FECredProd.TextMatrix(i, 7) = Format(MatMontos(RDet.Bookmark - 1, 3), "#0.00")
                FECredProd.TextMatrix(i, 8) = Format(MatMontos(RDet.Bookmark - 1, 4), "#0")
                FECredProd.TextMatrix(i, 9) = Format(MatMontos(RDet.Bookmark - 1, 5), "#0.00")
                FECredProd.TextMatrix(i, 10) = Format(MatMontos(RDet.Bookmark - 1, 6), "#0")
                Exit For
            End If
        Next i
        RDet.MoveNext
    Loop
    RDet.Close
    
    Set RDet = Nothing
    If FECredProd.Rows > 2 Then
        FECredProd.TopRow = 1
    End If
    
    Screen.MousePointer = 0
    LblBuscar.Visible = False
End Sub

Private Sub Form_Load()
    CentraSdi Me
    CargaControles
    nTipoAct = -1
    
End Sub

Private Sub txtFecFin_GotFocus()
    fEnfoque txtFecFin
End Sub

Private Sub txtFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.FECredProd.SetFocus
    End If
End Sub

Private Sub txtFecIni_GotFocus()
    fEnfoque txtFecIni
End Sub

Private Sub txtFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecFin.SetFocus
    End If
End Sub


