VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogSaneamiento 
   Caption         =   "Saneamiento"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9225
   LinkTopic       =   "Form2"
   ScaleHeight     =   6315
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   660
      Left            =   0
      TabIndex        =   17
      Top             =   5640
      Width           =   9180
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
         Height          =   390
         Left            =   3720
         TabIndex        =   21
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton CmdEliminar 
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
         Height          =   390
         Left            =   2520
         TabIndex        =   20
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1320
         TabIndex        =   19
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
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
         Height          =   390
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   180
         Width           =   1125
      End
   End
   Begin VB.Frame frmSaneamiento 
      Height          =   4335
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   9180
      Begin VB.Frame frmSaneOpcion 
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8895
         Begin VB.ComboBox cboTipoSaneamiento 
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Text            =   "cboTipoSaneamiento"
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox cboPeriodoSaneamiento 
            Height          =   315
            Left            =   3000
            TabIndex        =   11
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtMontoSaneamiento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5040
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdAgregar 
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
            Left            =   7560
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkMoneda 
            Caption         =   "Dolar"
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
            Left            =   6480
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Saneamiento"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo"
            Height          =   255
            Left            =   3000
            TabIndex        =   15
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "Monto"
            Height          =   255
            Left            =   5040
            TabIndex        =   14
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Moneda"
            Height          =   255
            Left            =   6480
            TabIndex        =   13
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.CheckBox ckAdjudicado 
         Caption         =   "Adjudicado"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3840
         Width           =   1455
      End
      Begin Sicmact.FlexEdit FESaneamiento 
         Height          =   2535
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4471
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Credito-cGarantia-Sanemiento-Periodo-Monto-Moneda"
         EncabezadosAnchos=   "400-2200-1300-2600-900-1200-1200"
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
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-C-R-C"
         FormatosEdit    =   "0-0-0-0-0-4-0"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame FraBuscaPers 
      Height          =   1275
      Left            =   0
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
Attribute VB_Name = "frmLogSaneamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pgcCtaCod As String

Dim bEstadoCargando As Boolean
Dim cmdEjecutar As Integer

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

Private Sub CmdAceptar_Click()
Dim i As Integer
Dim J As Integer
Dim oGaran As NGarantia
Set oGaran = New NGarantia
Dim nCont As Integer
For i = 0 To nPos
    nCont = nCont + 1
    Call oGaran.InsertarGarantiaSaneamiento(sNumgarant, sCtaCod, MatrixGarantias(6, i), MatrixGarantias(4, i), MatrixGarantias(5, i), gdFecSis, gsCodUser, i, nEstadoAdju, dEstadoAdju, nEstadoAdju, gdFecSis, cUsuariAdju, gsCodAge, CInt(MatrixGarantias(7, i)), 2)
Next i
If nCont > 0 Then
     MsgBox "Datos se registraron correctamente...", vbInformation, "Aviso"
     CmdAceptar.Enabled = False
Else
    MsgBox "Ingrese Datos Correctamente...", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdAgregar_Click()
    Dim J As Integer
    Dim i As Integer
    Dim NCaDAr As Integer
    Dim NTipoMoneda As Integer
    NTipoMoneda = 0
    NCaDAr = 0
    
If Val(txtMontoSaneamiento.Text) = 0 Or Trim(Left(Me.cboTipoSaneamiento.Text, 30)) = "" _
    Or Trim(Left(Me.cboPeriodoSaneamiento.Text, 30)) = "" Then
    MsgBox "Ingrese Datos Correctamente...", vbInformation, "Aviso"
    Exit Sub
End If
If nDat = 1 Then
For i = 0 To nPos
    If Trim(MatrixGarantias(3, i)) = Val(Trim(Right(cboTipoSaneamiento.Text, 8))) And MatrixGarantias(4, i) = Val(Trim(Right(cboPeriodoSaneamiento.Text, 8))) Then
        MsgBox "Este dato ya fue registrado...", vbInformation, "Aviso"
        txtMontoSaneamiento.Text = "0.00"
        Exit Sub
    End If
Next i
End If
'20080811************
        If chkMoneda.value = 1 Then
            NTipoMoneda = 1
        End If
 nDat = 1
    FESaneamiento.AdicionaFila
     If FESaneamiento.Row = 1 Then
        ReDim MatrixGarantias(1 To 7, 0 To 0)
     End If
     nPos = FESaneamiento.Row - 1
     MatrixGarantias(1, nPos) = FESaneamiento.Row
     ReDim Preserve MatrixGarantias(1 To 7, 0 To UBound(MatrixGarantias, 2) + 1)
     If nPos >= 1 Then
             For J = 0 To nPos - 1
                If Trim(Right(cboPeriodoSaneamiento.Text, 8)) > MatrixGarantias(4, nPos - 1) Then
                    i = nPos
                    Exit For
                End If
                If Trim(Right(cboPeriodoSaneamiento.Text, 8)) < MatrixGarantias(4, 0) Then
                    i = 0
                    Exit For
                End If
                If Trim(Right(cboPeriodoSaneamiento.Text, 8)) > MatrixGarantias(4, J) And Trim(Right(cboPeriodoSaneamiento.Text, 8)) <= MatrixGarantias(4, J + 1) Then
                    i = J + 1
                    Exit For
                    '***********
                End If
            Next J
            For J = nPos - 1 To i Step -1
                    MatrixGarantias(1, J + 1) = MatrixGarantias(1, J)
                    MatrixGarantias(2, J + 1) = MatrixGarantias(2, J)
                    MatrixGarantias(3, J + 1) = MatrixGarantias(3, J)
                    MatrixGarantias(4, J + 1) = MatrixGarantias(4, J)
                    MatrixGarantias(5, J + 1) = MatrixGarantias(5, J)
                    MatrixGarantias(6, J + 1) = MatrixGarantias(6, J)
                    MatrixGarantias(7, J + 1) = MatrixGarantias(7, J)
                    
             Next J
                    MatrixGarantias(1, i) = sCtaCod
                    MatrixGarantias(2, i) = sNumgarant
                    MatrixGarantias(3, i) = Left(cboTipoSaneamiento.Text, 40)
                    MatrixGarantias(4, i) = Left(cboPeriodoSaneamiento.Text, 4)
                    MatrixGarantias(5, i) = Me.txtMontoSaneamiento.Text
                    MatrixGarantias(6, i) = Trim(Right(cboTipoSaneamiento.Text, 4))
                    MatrixGarantias(7, i) = NTipoMoneda
    Else
                    MatrixGarantias(1, nPos) = sCtaCod
                    MatrixGarantias(2, nPos) = sNumgarant
                    MatrixGarantias(3, nPos) = Left(cboTipoSaneamiento.Text, 40)
                    MatrixGarantias(4, nPos) = Left(cboPeriodoSaneamiento.Text, 4)
                    MatrixGarantias(5, nPos) = Me.txtMontoSaneamiento.Text
                    MatrixGarantias(6, nPos) = Trim(Right(cboTipoSaneamiento.Text, 4))
                    MatrixGarantias(7, nPos) = NTipoMoneda
    End If
    
    For i = 0 To nPos
        FESaneamiento.EliminaFila (1)
    Next i
    For i = 0 To nPos
        FESaneamiento.AdicionaFila
        FESaneamiento.TextMatrix(FESaneamiento.Row, 1) = MatrixGarantias(1, i)
        FESaneamiento.TextMatrix(FESaneamiento.Row, 2) = MatrixGarantias(2, i)
        FESaneamiento.TextMatrix(FESaneamiento.Row, 3) = MatrixGarantias(3, i)
        FESaneamiento.TextMatrix(FESaneamiento.Row, 4) = MatrixGarantias(4, i)
        FESaneamiento.TextMatrix(FESaneamiento.Row, 5) = MatrixGarantias(5, i)
        FESaneamiento.TextMatrix(FESaneamiento.Row, 6) = MatrixGarantias(7, i)
        NCaDAr = 1
    Next
    txtMontoSaneamiento.Text = "0.00"
    CmdEliminar.Enabled = True
End Sub

Private Sub CmdBuscaPersona_Click()
    Call cmdCancelarInicio
    ObtieneDocumPersona
    If vTipoInicio = ConsultaGarant Then
        CmdEliminar.Enabled = False
    End If
End Sub

Private Sub cmdBuscar_Click()
    bAsignadoACredito = False
    
    If Me.LstGaratias.ListItems.Count = 0 Then
        MsgBox "No Existe Garantia que Mostrar ", vbInformation, "Aviso"
        Exit Sub
    End If
    
    sNumgarant = Trim(Me.LstGaratias.SelectedItem.SubItems(2))
    sCtaCod = Trim(Me.LstGaratias.SelectedItem.SubItems(7))
    nEstadoA = CInt(Trim(Me.LstGaratias.SelectedItem.SubItems(8)))
    Call ObtenerArreglo(sNumgarant, sCtaCod, 2)
    If vTipoInicio = ConsultaGarant Then
        CmdEliminar.Enabled = False
    End If
    If nEstadoA = 10 Then
        ckAdjudicado.value = 1
        CmdCancelar.Enabled = False
        ckAdjudicado.Enabled = False
    End If
End Sub
Public Sub ObtenerArreglo(ByVal sNumGarantia As String, ByVal sCodCta As String, ByVal pnTESan As Integer)
    Dim oGaran As DGarantia
    Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    Set oGaran = New DGarantia
    
    Dim i As Integer
    For i = 0 To nPos
        FESaneamiento.EliminaFila (1)
    Next i
    nPos = 0
    Set R = oGaran.RecuperaDatosGarantiaSanemiento(sNumGarantia, sCodCta, pnTESan)
    If R.RecordCount > 0 Then
            If Not R.EOF And Not R.BOF Then
                R.MoveFirst
            End If
    Do Until R.EOF
        FESaneamiento.AdicionaFila
        nPos = FESaneamiento.Row - 1
        ReDim Preserve MatrixGarantias(1 To 7, 0 To nPos + 1)
        FESaneamiento.AdicionaFila
        MatrixGarantias(1, nPos) = R!cCtaCod
        MatrixGarantias(2, nPos) = R!cNumGarant
        MatrixGarantias(3, nPos) = R!cConsDescripcion
        MatrixGarantias(4, nPos) = R!nPeriSan
        MatrixGarantias(5, nPos) = R!nMontSan
        MatrixGarantias(6, nPos) = R!nTipoSan
        MatrixGarantias(7, nPos) = R!nMoneda
        
        FESaneamiento.TextMatrix(FESaneamiento.Row, 1) = MatrixGarantias(1, nPos)
        FESaneamiento.TextMatrix(FESaneamiento.Row, 2) = MatrixGarantias(2, nPos)
        FESaneamiento.TextMatrix(FESaneamiento.Row, 3) = MatrixGarantias(3, nPos)
        FESaneamiento.TextMatrix(FESaneamiento.Row, 4) = MatrixGarantias(4, nPos)
        FESaneamiento.TextMatrix(FESaneamiento.Row, 5) = MatrixGarantias(5, nPos)
        If R!nMoneda = 0 Then
            FESaneamiento.TextMatrix(FESaneamiento.Row, 6) = "Soles"
        Else
            FESaneamiento.TextMatrix(FESaneamiento.Row, 6) = "Dolares"
        End If
        R.MoveNext
        CmdEliminar.Enabled = True
    Loop
    End If
End Sub
Private Sub cmdCancelarInicio()
    Call LimpiaPantalla
    cmdEjecutar = -1
End Sub

Private Sub cmdCancelar_Click()
    CmdAceptar.Enabled = True
    CmdAceptar.Visible = True
    CmdCancelar.Enabled = False
    CmdCancelar.Caption = "Cancelar"
End Sub

Private Sub LimpiaPantalla()
    bCarga = True
    Call LimpiaControles(Me)
    Call InicializaCombos(Me)
    CmdEliminar.Enabled = False
    bCarga = False
End Sub
Private Sub ObtieneDocumPersona()
Dim oGaran As DGarantia
Dim R As ADODB.Recordset
Dim oPers As UPersona
Dim L As ListItem
    
    LstGaratias.ListItems.Clear
    Set oPers = New UPersona
    Set oPers = frmBuscaPersona.Inicio
    Set oGaran = New DGarantia
    
    If oPers Is Nothing Then
        Exit Sub
    End If
    Set R = oGaran.RecuperaGarantiasPersonaLogistica(oPers.sPersCod, True, , True)
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


Private Sub CmdEliminar_Click()
   Dim nXPos As Integer
    nXPos = FESaneamiento.Row
    If nPos >= 1 Then
    If MsgBox("Esta Seguro de Eliminar este registro.", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        FESaneamiento.EliminaFila (FESaneamiento.Row)
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
            Next J
            nPos = nPos - 1
        Else
            nPos = nPos - 1
            nDat = 0
        End If
    End If
    Else
        If FESaneamiento.Row >= 1 Then
        FESaneamiento.EliminaFila (1)
        End If
        nPos = -1
        nDat = 0
    End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Conn As DConecta
Dim oCons As DConstante
Dim Res As ADODB.Recordset
Set Res = New ADODB.Recordset
Set oCons = New DConstante
Set Conn = New DConecta
 
Dim i As Integer
For i = CInt(Format(gdFecSis, "YYYY")) - 10 To CInt(Format(gdFecSis, "YYYY"))
    cboPeriodoSaneamiento.AddItem i & Space(200) & Trim(i)
Next i
Conn.AbreConexion
Set Res = oCons.RecuperaConstantes(9073)
Call Llenar_Combo_con_Recordset(Res, cboTipoSaneamiento)
Set Res = Nothing
Conn.CierraConexion
Set Conn = Nothing
End Sub

Sub Llenar_Combo_con_Recordset(pRs As ADODB.Recordset, pcboObjeto As ComboBox)

pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cConsDescripcion) & Space(100) & Trim(Str(pRs!nConsValor))
    pRs.MoveNext
Loop
pRs.Close
End Sub
