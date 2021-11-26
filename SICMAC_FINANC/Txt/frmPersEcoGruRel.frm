VERSION 5.00
Begin VB.Form frmPersEcoGruRel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Relacion de Grupos Economicos"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   Icon            =   "frmPersEcoGruRel.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   7125
      TabIndex        =   10
      Top             =   6390
      Width           =   1065
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   330
      Left            =   45
      TabIndex        =   8
      Top             =   6390
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   45
      TabIndex        =   7
      Top             =   6390
      Width           =   1065
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   330
      Left            =   7110
      TabIndex        =   1
      Top             =   210
      Width           =   1065
   End
   Begin Sicmact.TxtBuscar txtGE 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
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
      Appearance      =   0
      TipoBusqueda    =   2
      sTitulo         =   ""
   End
   Begin VB.Frame fraGE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Grupo Economico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5760
      Left            =   45
      TabIndex        =   3
      Top             =   585
      Width           =   8145
      Begin VB.CommandButton cmdEliminarEmp 
         Caption         =   "&Eliminar"
         Height          =   330
         Left            =   6975
         TabIndex        =   12
         Top             =   1875
         Width           =   1065
      End
      Begin VB.CommandButton cmdNuevoEmp 
         Caption         =   "N&uevo"
         Height          =   330
         Left            =   5865
         TabIndex        =   11
         Top             =   1875
         Width           =   1065
      End
      Begin VB.Frame fraGERel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Relacion de Grupo Economico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3450
         Left            =   105
         TabIndex        =   5
         Top             =   2205
         Width           =   7935
         Begin VB.CheckBox chkVerTodos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Ver Todos"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   3105
            TabIndex        =   15
            Top             =   3075
            Width           =   2400
         End
         Begin VB.CommandButton cmdEliminarPers 
            Caption         =   "E&liminar"
            Height          =   330
            Left            =   6750
            TabIndex        =   14
            Top             =   3015
            Width           =   1065
         End
         Begin VB.CommandButton cmdNuevoPers 
            Caption         =   "Nue&vo"
            Height          =   330
            Left            =   5640
            TabIndex        =   13
            Top             =   3015
            Width           =   1065
         End
         Begin Sicmact.FlexEdit FlexPers 
            Height          =   2670
            Left            =   90
            TabIndex        =   6
            Top             =   270
            Width           =   7740
            _ExtentX        =   13653
            _ExtentY        =   4710
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Cod Persona-Nombre-Cargo-Particip-Rel"
            EncabezadosAnchos=   "300-1200-3500-1200-1000-1200"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-X-3-4-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-1-0-3-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-R-C"
            FormatosEdit    =   "0-0-0-0-2-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin Sicmact.FlexEdit flexEmp 
         Height          =   1605
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   2831
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cod Persona-Nombre-Relacion"
         EncabezadosAnchos=   "300-1200-3500-2400"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3"
         TextStyleFixed  =   4
         ListaControles  =   "0-1-0-3"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbPuntero       =   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   330
      Left            =   1155
      TabIndex        =   9
      Top             =   6390
      Width           =   1065
   End
   Begin VB.Label lblGE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1785
      TabIndex        =   2
      Top             =   233
      Width           =   5250
   End
End
Attribute VB_Name = "frmPersEcoGruRel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
    Activa False
    Form_Load
End Sub

Private Sub cmdEditar_Click()
    If Me.txtGE.Text = "" Then
        MsgBox "Debe Elegir a un grupo economico para poder agregar a las empresa juridica.", vbInformation, "Aviso"
        Me.txtGE.SetFocus
        Exit Sub
    End If
    
    Activa True
End Sub

Private Sub cmdEliminarEmp_Click()
    Dim lnI As Integer
    
    For lnI = 1 To Me.FlexPers.Rows - 1
        If FlexPers.TextMatrix(lnI, 5) = flexEmp.TextMatrix(flexEmp.Row, 1) Then
            flexEmp.EliminaFila lnI
            lnI = lnI - 1
        End If
    Next lnI
    
    flexEmp.EliminaFila flexEmp.Row
End Sub

Private Sub cmdEliminarPers_Click()
    FlexPers.EliminaFila FlexPers.Row
End Sub

Private Sub cmdGrabar_Click()
    Dim oGE As DGrupoEco
    Set oGE = New DGrupoEco
    
    Me.chkVerTodos.value = 1
    flexEmp_RowColChange
    
    If MsgBox("Desea guardar los cambios ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    If Not Valida Then Exit Sub
    
    oGE.ActulizaGERel flexEmp.GetRsNew, FlexPers.GetRsNew, Me.txtGE.Text
    
    cmdCancelar_Click
End Sub

Private Sub cmdNuevo_Click()
    frmPersEcoGru.Show 1
    Form_Load
End Sub

Private Sub cmdNuevoEmp_Click()
    If Me.txtGE.Text = "" Then
        MsgBox "Debe Elegir a un grupo economico para poder agregar a las empresa juridica.", vbInformation, "Aviso"
        Me.txtGE.SetFocus
        Exit Sub
    End If
    
    flexEmp.AdicionaFila
End Sub

Private Sub cmdNuevoPers_Click()
    If Me.flexEmp.TextMatrix(flexEmp.Row, 1) = "" Then
        MsgBox "Debe Elegir a una empresa juridica que pertenesca al grupo economico.", vbInformation, "Aviso"
        Me.cmdNuevoEmp.SetFocus
        Exit Sub
    End If
    
    FlexPers.AdicionaFila
    FlexPers.TextMatrix(FlexPers.Rows - 1, 5) = Me.flexEmp.TextMatrix(flexEmp.Row, 1)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub flexEmp_RowColChange()
    Dim lnI As Integer
    
    For lnI = 1 To Me.FlexPers.Rows - 1
        If flexEmp.TextMatrix(flexEmp.Row, 1) = FlexPers.TextMatrix(lnI, 5) Or chkVerTodos.value Then
            FlexPers.RowHeight(lnI) = 285
        Else
            FlexPers.RowHeight(lnI) = 0
        End If
    Next lnI
End Sub

Private Sub Form_Load()
    Dim oGE As DGrupoEco
    Set oGE = New DGrupoEco
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Me.Icon = LoadPicture(App.path & "\Graficos\cm.ico")
    flexEmp.CargaCombo oCon.GetConstante(4028, , , , , , True)
    FlexPers.CargaCombo oCon.GetConstante(4029, , , , , , True)
    
    Me.txtGE.rs = oGE.GetGE
    
    Activa False
    CentraForm Me
End Sub

Private Sub GetData(psGECod As String)
    Dim oGE As DGrupoEco
    Set oGE = New DGrupoEco
    
    Me.flexEmp.Clear
    Me.flexEmp.Rows = 2
    Me.flexEmp.FormaCabecera
    
    Me.FlexPers.Clear
    Me.FlexPers.Rows = 2
    Me.FlexPers.FormaCabecera
    
    Me.flexEmp.rsFlex = oGE.GetGEEmp(psGECod)
    Me.FlexPers.rsFlex = oGE.GetGEPers(psGECod)
    
End Sub

Private Sub txtGE_EmiteDatos()
    Me.lblGE.Caption = txtGE.psDescripcion
    
    GetData txtGE.Text
End Sub

Private Sub Activa(pbEditar As Boolean)
    Me.cmdCancelar.Visible = pbEditar
    Me.cmdEditar.Visible = Not pbEditar
    Me.cmdGrabar.Enabled = pbEditar
    Me.txtGE.Enabled = Not pbEditar
    Me.fraGE.Enabled = pbEditar
End Sub

Private Function Valida() As Boolean
    Dim lnI As Integer
    
    For lnI = 1 To Me.flexEmp.Rows - 1
        flexEmp.Row = lnI
        If flexEmp.TextMatrix(lnI, 1) = "" Then
            MsgBox "Debe ingresa un codigo persona. En el registro :" & Str(lnI), vbInformation, "Aviso"
            flexEmp.Col = 1
            flexEmp.SetFocus
            Valida = False
            Exit Function
        ElseIf flexEmp.TextMatrix(lnI, 3) = "" Then
            MsgBox "Debe ingresa una relacion de la persona con el grupo economico. En el registro :" & Str(lnI), vbInformation, "Aviso"
            flexEmp.Col = 3
            flexEmp.SetFocus
            Valida = False
            Exit Function
        Else
            Valida = True
        End If
    Next lnI
    
    For lnI = 1 To Me.FlexPers.Rows - 1
        FlexPers.Row = lnI
        If FlexPers.TextMatrix(lnI, 1) = "" Then
            MsgBox "Debe ingresa un codigo persona. En el registro :" & Str(lnI), vbInformation, "Aviso"
            FlexPers.Col = 1
            FlexPers.SetFocus
            Valida = False
            Exit Function
        ElseIf FlexPers.TextMatrix(lnI, 3) = "" Then
            MsgBox "Debe ingresa un relacionado a la persona. En el registro :" & Str(lnI), vbInformation, "Aviso"
            FlexPers.Col = 3
            FlexPers.SetFocus
            Valida = False
            Exit Function
        ElseIf FlexPers.TextMatrix(lnI, 4) = "" Then
            MsgBox "Debe ingresa un porcentaje de participación. En el registro :" & Str(lnI), vbInformation, "Aviso"
            FlexPers.Col = 4
            FlexPers.SetFocus
            Valida = False
            Exit Function
        Else
            Valida = True
        End If
    Next lnI
    
End Function


