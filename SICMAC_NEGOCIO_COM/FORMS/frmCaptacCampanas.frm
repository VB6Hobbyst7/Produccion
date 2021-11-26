VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCaptacCampanas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Campañas de Captaciones"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "frmCaptacCampanas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDineroNuevo 
      Caption         =   "Dinero Nuevo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   6240
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Top             =   2520
      Width           =   8295
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6960
         TabIndex        =   12
         Top             =   240
         Width           =   1170
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5760
         TabIndex        =   11
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.Frame FraLista 
      Height          =   2895
      Left            =   120
      TabIndex        =   27
      Top             =   3240
      Width           =   8295
      Begin SICMACT.FlexEdit feCamp 
         Height          =   2475
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4366
         Cols0           =   16
         HighLight       =   1
         AllowUserResizing=   1
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   $"frmCaptacCampanas.frx":030A
         EncabezadosAnchos=   "500-0-2500-0-1000-0-1200-0-1800-1000-1200-1200-1200-500-800-1400"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-13-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-4-0-0"
         EncabezadosAlineacion=   "C-C-L-L-L-C-L-C-L-R-C-C-C-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         AvanceCeldas    =   1
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
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
      Left            =   7320
      TabIndex        =   26
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   24
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdDesactivar 
      Caption         =   "Desactivar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   25
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame fraDatos 
      Enabled         =   0   'False
      Height          =   2535
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   8295
      Begin VB.CheckBox chkPJ 
         Caption         =   "Pers. Jurídica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CheckBox chkTodosAgencia 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5680
         TabIndex        =   9
         Top             =   690
         Width           =   1215
      End
      Begin VB.ListBox LstAgencias 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         ItemData        =   "frmCaptacCampanas.frx":039B
         Left            =   5640
         List            =   "frmCaptacCampanas.frx":03A2
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtCampana 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   700
         Width           =   4455
      End
      Begin VB.TextBox txtPeriodoVigMeses 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   4680
         MaxLength       =   4
         TabIndex        =   5
         Text            =   "0"
         Top             =   1155
         Width           =   735
      End
      Begin VB.ComboBox cboSubProducto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox cboProducto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cboMoneda 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCaptacCampanas.frx":03B3
         Left            =   960
         List            =   "frmCaptacCampanas.frx":03B5
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin SICMACT.EditMoney txtMontoMin 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   1155
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox txtFechaIni 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   1620
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   2880
         TabIndex        =   7
         Top             =   1620
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Campaña:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   22
         Top             =   760
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ini:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2040
         TabIndex        =   20
         Top             =   1680
         Width           =   750
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Monto Minimo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   1195
         Width           =   1005
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Periodo Vigencia Meses: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2760
         TabIndex        =   18
         Top             =   1195
         Width           =   1830
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sub Producto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4800
         TabIndex        =   17
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2400
         TabIndex        =   16
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCaptacCampanas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************************
'** Nombre : frmCaptacCampanas
'** Descripción : Formulario para administrar las campañas de captaciones
'** Creación : JUEZ, 20160415 09:00:00 AM
'*******************************************************************************************

Option Explicit

Dim oNCapDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rs As ADODB.Recordset
Dim fnCampanaCod As Long
Dim bEdita As Boolean
Dim bCheckTodosAge As Boolean, bCheckLista As Boolean

Public Sub inicia()
    fnCampanaCod = 0
    bEdita = False
    CargarConstantes
    CargarCampanas
    Me.Show 1
End Sub

Private Sub CargarConstantes()
    CargarCombos CboMoneda, gMoneda
    CargarCombos cboProducto, gProducto
    CargarAgencias
End Sub

Private Sub HabilitaControlesRegistro(ByVal pbHabilita As Boolean)
    FraDatos.Enabled = pbHabilita
    cmdGrabar.Enabled = pbHabilita
    cmdCancelar.Enabled = pbHabilita
    FraLista.Enabled = Not pbHabilita
    cmdNuevo.Enabled = Not pbHabilita
    cmdEditar.Enabled = Not pbHabilita
    cmdDesactivar.Enabled = Not pbHabilita
End Sub

Private Sub cboMoneda_Click()
    If cboProducto.ListCount > 0 Then cboProducto.SetFocus
End Sub

Private Sub cboProducto_Click()
Dim nConstSubProd As ConstanteCabecera

    txtPeriodoVigMeses.Text = "0"
    If Trim(Right(cboProducto.Text, 4)) <> "" Then
        Select Case Trim(Right(cboProducto.Text, 4))
            Case gCapAhorros
                nConstSubProd = gCaptacSubProdAhorros
                txtPeriodoVigMeses.Enabled = True
            Case gCapPlazoFijo
                nConstSubProd = gCaptacSubProdPlazoFijo
                txtPeriodoVigMeses.Enabled = False
            Case gCapCTS
                nConstSubProd = gCaptacSubProdCTS
                txtPeriodoVigMeses.Enabled = True
        End Select
        CargarCombos cboSubProducto, nConstSubProd
        'cboSubProducto.SetFocus
    End If
End Sub

Private Sub cboSubProducto_Click()
    If FraDatos.Enabled And txtCampana.Enabled Then txtCampana.SetFocus
End Sub

Private Sub chkTodosAgencia_Click()
Dim i As Integer
    If Not bCheckLista Then
        bCheckTodosAge = True
        For i = 0 To LstAgencias.ListCount - 1
            LstAgencias.Selected(i) = IIf(chkTodosAgencia.value = 1, True, False)
        Next i
        bCheckTodosAge = False
    End If
End Sub

Private Sub lstAgencias_Click()
Dim i As Integer
    bCheckLista = True
    For i = 0 To LstAgencias.ListCount - 1
        If Not LstAgencias.Selected(i) Then
            If Not bCheckTodosAge Then chkTodosAgencia.value = 0
            bCheckLista = False
            Exit Sub
        End If
    Next i
    bCheckLista = False
End Sub

Private Sub cmdCancelar_Click()
    fnCampanaCod = 0
    txtCampana.Text = ""
    txtMontoMin.Text = "0"
    txtPeriodoVigMeses.Text = "0"
    txtFechaIni.Text = "__/__/____"
    txtFechaFin.Text = "__/__/____"
    CargarConstantes
    CargarCampanas
    chkTodosAgencia.value = 0
    chkPJ.value = 0
    bEdita = False
    txtCampana.Locked = False
    HabilitaControlesRegistro False
    chkDineroNuevo.value = 0 'APRI20210621 ERS031-2021
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub CargarCombos(ByRef combo As ComboBox, ByVal nConstante As ConstanteCabecera)
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsGen As ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsGen = clsGen.GetConstante(nConstante, , "", " ")
    Set clsGen = Nothing

    combo.Clear
    Do While Not rsGen.EOF
        If nConstante <> gProducto Or (nConstante = gProducto And Left(rsGen("nConsValor"), 2) = "23" And Right(rsGen("nConsValor"), 1) <> "0") Then
            combo.AddItem rsGen("cDescripcion") & Space(100) & rsGen("nConsValor")
        End If
        rsGen.MoveNext
    Loop

    combo.ListIndex = IIf(combo.ListCount <= 0, -1, 0)
    rsGen.Close
    Set rsGen = Nothing
End Sub

Private Sub CargarAgencias()
Dim oAge As COMDConstantes.DCOMAgencias
Dim rsAgencias As ADODB.Recordset
    
    Set oAge = New COMDConstantes.DCOMAgencias
        Set rsAgencias = oAge.ObtieneAgencias()
    Set oAge = Nothing
    If rsAgencias Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        LstAgencias.Clear
        With rsAgencias
            Do While Not rsAgencias.EOF
                LstAgencias.AddItem rsAgencias!nConsValor & " " & Trim(rsAgencias!cConsDescripcion)
                rsAgencias.MoveNext
            Loop
        End With
    End If
End Sub

Private Sub CargarCampanas()
Dim rsCamp As ADODB.Recordset
    
    Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set rsCamp = oNCapDef.GetCaptacCampanas()
    Set oNCapDef = Nothing
    
    If Not rsCamp.EOF And Not rsCamp.BOF Then
        Call LimpiaFlex(feCamp)
        
        Do While Not rsCamp.EOF
            feCamp.AdicionaFila
            feCamp.TextMatrix(feCamp.row, 1) = rsCamp("nCampanaCod")
            feCamp.TextMatrix(feCamp.row, 2) = rsCamp("cCampanaDesc")
            feCamp.TextMatrix(feCamp.row, 3) = rsCamp("nMoneda")
            feCamp.TextMatrix(feCamp.row, 4) = rsCamp("cMoneda")
            feCamp.TextMatrix(feCamp.row, 5) = rsCamp("nProducto")
            feCamp.TextMatrix(feCamp.row, 6) = rsCamp("cProducto")
            feCamp.TextMatrix(feCamp.row, 7) = rsCamp("nSubProducto")
            feCamp.TextMatrix(feCamp.row, 8) = rsCamp("cSubProducto")
            feCamp.TextMatrix(feCamp.row, 9) = Format(rsCamp("nMontoMin"), "#,##0.00")
            feCamp.TextMatrix(feCamp.row, 10) = rsCamp("nPeriodoVigMeses")
            feCamp.TextMatrix(feCamp.row, 11) = Format(rsCamp("dFechaIni"), "dd/MM/yyyy")
            feCamp.TextMatrix(feCamp.row, 12) = Format(rsCamp("dFechaFin"), "dd/MM/yyyy")
            feCamp.TextMatrix(feCamp.row, 13) = IIf(rsCamp("bPersJur"), "1", "")
            feCamp.TextMatrix(feCamp.row, 14) = "Ver"
            feCamp.TextMatrix(feCamp.row, 15) = rsCamp("cDineroNuevo") 'APRI20210621 ERS031-2021
            rsCamp.MoveNext
        Loop
        feCamp.TopRow = 1
        feCamp.Col = 1
    End If
    
End Sub

Private Sub cmdEditar_Click()
    If feCamp.TextMatrix(feCamp.row, 0) = "" Then
        MsgBox "Debe seleccionar al menos un registro", vbInformation, "Aviso"
        Exit Sub
    End If
    HabilitaControlesRegistro (True)
    If Not CargaDatos Then
        MsgBox "No se pudieron cargar los datos", vbInformation, "Aviso"
    End If
End Sub

Private Function CargaDatos() As Boolean
Dim i As Integer
    CargaDatos = False
    fnCampanaCod = CLng(feCamp.TextMatrix(feCamp.row, 1))
    bEdita = True
    txtCampana.Locked = True
    
    CboMoneda.ListIndex = IndiceListaCombo(CboMoneda, CStr(feCamp.TextMatrix(feCamp.row, 3)))
    cboProducto.ListIndex = IndiceListaCombo(cboProducto, CStr(feCamp.TextMatrix(feCamp.row, 5)))
    cboSubProducto.ListIndex = IndiceListaCombo(cboSubProducto, CStr(feCamp.TextMatrix(feCamp.row, 7)))
        
    txtCampana.Text = feCamp.TextMatrix(feCamp.row, 2)
    txtMontoMin.Text = Format(feCamp.TextMatrix(feCamp.row, 9), "#,##0.00")
    txtPeriodoVigMeses.Text = feCamp.TextMatrix(feCamp.row, 10)
    txtFechaIni.Text = Format(feCamp.TextMatrix(feCamp.row, 11), "dd/MM/yyyy")
    txtFechaFin.Text = Format(feCamp.TextMatrix(feCamp.row, 12), "dd/MM/yyyy")
    
    Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set rs = oNCapDef.GetCaptacCampanasAge(fnCampanaCod)
    Set oNCapDef = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        Do While Not rs.EOF
            For i = 0 To LstAgencias.ListCount - 1
                If rs("cAgeCod") = Left(LstAgencias.List(i), 2) Then
                    LstAgencias.Selected(i) = True
                End If
            Next i
            rs.MoveNext
        Loop
    End If
    
    chkTodosAgencia.value = 1
    For i = 0 To LstAgencias.ListCount - 1
        If Not LstAgencias.Selected(i) Then
            chkTodosAgencia.value = 0
            Exit For
        End If
    Next i
    Set rs = Nothing
    chkDineroNuevo = CInt(IIf(feCamp.TextMatrix(feCamp.row, 15) = "SI", 1, 0)) 'APRI20210621 ERS031-2021
    CargaDatos = True
End Function

Private Sub cmdDesactivar_Click()
Dim objPista As COMManejador.Pista

    If feCamp.TextMatrix(feCamp.row, 0) = "" Then
        MsgBox "Debe seleccionar al menos un registro", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("¿Está seguro de desactivar la campaña '" + CStr(feCamp.TextMatrix(feCamp.row, 2)) + "'?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
            oNCapDef.EliminaCaptacCampanas feCamp.TextMatrix(feCamp.row, 1)
        Set oNCapDef = Nothing
        Set objPista = New COMManejador.Pista
            objPista.InsertarPista gCaptacCampanaEdit, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Desactivar Campaña Codigo " & feCamp.TextMatrix(feCamp.row, 1) & " - " & CStr(feCamp.TextMatrix(feCamp.row, 2))
        Set objPista = Nothing
        MsgBox "Campaña desactivada", vbInformation, "Aviso"
        feCamp.EliminaFila feCamp.row
    End If
End Sub

Private Sub cmdNuevo_Click()
    HabilitaControlesRegistro True
    CboMoneda.SetFocus
End Sub

Private Sub feCamp_Click()
Dim oAge As COMDConstantes.DCOMAgencias
    If feCamp.TextMatrix(feCamp.row, feCamp.Col) <> "" Then
        Dim rsLista As ADODB.Recordset, rsDatos As ADODB.Recordset
        If feCamp.TextMatrix(feCamp.row, 0) <> "" Then
            If feCamp.Col = 14 Then
                Set oAge = New COMDConstantes.DCOMAgencias
                    Set rsLista = oAge.ObtieneAgencias()
                Set oAge = Nothing
                Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
                    Set rsDatos = oNCapDef.GetCaptacCampanasAge(feCamp.TextMatrix(feCamp.row, 1))
                Set oNCapDef = Nothing
                frmCredListaDatos.Inicio "Agencias", rsDatos, rsLista, 1
            End If
        End If
    End If
End Sub

Private Sub txtCampana_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        txtMontoMin.SetFocus
    End If
End Sub

Private Sub txtCampana_LostFocus()
    txtCampana.Text = UCase(txtCampana.Text)
End Sub

Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkPJ.SetFocus
    End If
End Sub

Private Sub txtFechaIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFechaFin.SetFocus
    End If
End Sub

Private Sub txtMontoMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPeriodoVigMeses.Enabled Then
            txtPeriodoVigMeses.SetFocus
        Else
            txtFechaIni.SetFocus
        End If
    End If
End Sub

Private Sub txtPeriodoVigMeses_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        txtFechaIni.SetFocus
    End If
End Sub

Private Sub cmdGrabar_Click()
Dim objPista As COMManejador.Pista
Dim lnCampanaCod As Long
Dim MatAgencias As Variant
Dim i As Integer, j As Integer
Dim lsMovNro As String
Dim bExito As Boolean

    If ValidaDatos Then
        j = 0
        ReDim MatAgencias(j)
        For i = 0 To LstAgencias.ListCount - 1
            If LstAgencias.Selected(i) Then
                ReDim Preserve MatAgencias(j)
                MatAgencias(j) = CStr(Left(LstAgencias.List(i), 2))
                j = j + 1
            End If
        Next i
        
        lsMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        
        If bEdita Then
            If fnCampanaCod <> 0 Then
                If MsgBox("Se van a actualizar los datos de la campaña." & Chr(13) & Chr(13) & _
                          "*  Los cambios aplicarán sólo para las cuentas nuevas." & Chr(13) & _
                          "*  Para las cuentas vigentes aperturadas dentro de esta campaña los cambios aplicarán a partir del siguiente cierre de mes." & Chr(13) & Chr(13) & _
                          "Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
                
                bExito = oNCapDef.ActualizaCaptacCampanas(fnCampanaCod, Trim(txtCampana.Text), Trim(Right(CboMoneda.Text, 1)), Trim(Right(cboProducto.Text, 4)), _
                                                          Trim(Right(cboSubProducto.Text, 4)), CDbl(txtMontoMin.Text), CInt(txtPeriodoVigMeses.Text), _
                                                          CDate(Me.txtFechaIni.Text), CDate(Me.txtFechaFin.Text), IIf(chkPJ.value = 1, True, False), lsMovNro, MatAgencias, IIf(chkDineroNuevo.value = 1, True, False))
                                                          'APRI20210621 ERSO31-2021 Add chkDineroNuevo
                If bExito Then
                    Set objPista = New COMManejador.Pista
                        objPista.InsertarPista gCaptacCampanaEdit, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Editar Campaña Codigo " & fnCampanaCod & " - " & Trim(txtCampana.Text)
                    Set objPista = Nothing
                    MsgBox "Se actualizaron los datos", vbInformation, "Aviso"
                    cmdCancelar_Click
                Else
                    MsgBox "Hubo un inconveniente en la actualización", vbInformation, "Aviso"
                End If
            Else
                MsgBox "Hubo un inconveniente en la actualización", vbInformation, "Aviso"
            End If
            
        Else
            If MsgBox("Se van a registrar los datos de la campaña, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
            
            fnCampanaCod = oNCapDef.InsertaCaptacCampanas(Trim(txtCampana.Text), Trim(Right(CboMoneda.Text, 1)), Trim(Right(cboProducto.Text, 4)), _
                                                          Trim(Right(cboSubProducto.Text, 4)), CDbl(txtMontoMin.Text), CInt(txtPeriodoVigMeses.Text), _
                                                          CDate(Me.txtFechaIni.Text), CDate(Me.txtFechaFin.Text), IIf(chkPJ.value = 1, True, False), lsMovNro, MatAgencias, IIf(chkDineroNuevo.value = 1, True, False))
                                                          'APRI20210621 ERSO31-2021 Add chkDineroNuevo
            If fnCampanaCod <> 0 Then
                Set objPista = New COMManejador.Pista
                    objPista.InsertarPista gCaptacCampanaReg, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Registrar Campaña Codigo " & fnCampanaCod & " - " & Trim(txtCampana.Text)
                Set objPista = Nothing
                MsgBox "Se registró la campaña", vbInformation, "Aviso"
                cmdCancelar_Click
            Else
                MsgBox "Hubo un inconveniente en el registro", vbInformation, "Aviso"
            End If
        End If
        Set oNCapDef = Nothing
    End If
End Sub

Private Function ValidaDatos() As Boolean
Dim i As Integer, bCheckAgencia As Boolean
    
    ValidaDatos = False
    
    If Trim(CboMoneda.Text) = "" Then
        MsgBox "Debe seleccionar la Moneda", vbInformation, "Aviso"
        CboMoneda.SetFocus
        Exit Function
    End If
    If Trim(cboProducto.Text) = "" Then
        MsgBox "Debe seleccionar el Producto", vbInformation, "Aviso"
        cboProducto.SetFocus
        Exit Function
    End If
    If Trim(cboSubProducto.Text) = "" Then
        MsgBox "Debe seleccionar el Sub Producto", vbInformation, "Aviso"
        cboSubProducto.SetFocus
        Exit Function
    End If
    If Not bEdita Then
        Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
            Set rs = oNCapDef.GetCaptacCampanas(, Trim(Right(cboProducto.Text, 4)), Trim(Right(cboSubProducto.Text, 4)), Trim(Right(CboMoneda.Text, 1)))
        Set oNCapDef = Nothing
        If Not (rs.EOF And rs.BOF) Then
            MsgBox "Ya existe una campaña para el Sub Producto y la Moneda seleccionada", vbInformation, "Aviso"
            cboSubProducto.SetFocus
            Exit Function
        End If
    End If
    If Trim(txtCampana.Text) = "" Then
        MsgBox "Debe ingresar el nombre de la Campaña", vbInformation, "Aviso"
        txtCampana.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtMontoMin.Text) Then
        MsgBox "Ingrese correctamente el Monto Mínimo", vbInformation, "Aviso"
        txtMontoMin.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtPeriodoVigMeses.Text) Then
        MsgBox "Ingrese correctamente el Periodo de Vigencia", vbInformation, "Aviso"
        txtPeriodoVigMeses.SetFocus
        Exit Function
    End If
    If ValidaFecha(txtFechaIni.Text) <> "" Then
        MsgBox "Ingrese correctamente la Fecha de Inicio de la campaña", vbInformation, "Aviso"
        txtFechaIni.SetFocus
        Exit Function
    End If
    If ValidaFecha(txtFechaFin.Text) <> "" Then
        MsgBox "Ingrese correctamente la Fecha Final de la campaña", vbInformation, "Aviso"
        txtFechaFin.SetFocus
        Exit Function
    End If
    If CDate(txtFechaFin.Text) < CDate(txtFechaIni.Text) Then
        MsgBox "La fecha final no puede ser menor que la fecha inicial", vbInformation, "Aviso"
        txtFechaIni.SetFocus
        Exit Function
    End If
    
    For i = 0 To LstAgencias.ListCount - 1
        If LstAgencias.Selected(i) Then bCheckAgencia = True
    Next i
    
    If Not bCheckAgencia Then
        MsgBox "Debe seleccionar al menos una agencia", vbInformation, "Aviso"
        LstAgencias.SetFocus
        Exit Function
    End If
    
    ValidaDatos = True
End Function
