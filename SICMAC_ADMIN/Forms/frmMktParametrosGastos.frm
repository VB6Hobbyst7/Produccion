VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMktParametrosGastos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Gastos Marketing"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9165
   Icon            =   "frmMktParametrosGastos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalir 
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
      Height          =   320
      Left            =   8040
      TabIndex        =   4
      Top             =   3705
      Width           =   1020
   End
   Begin VB.CommandButton btnGrabar 
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
      Height          =   320
      Left            =   6960
      TabIndex        =   3
      Top             =   3705
      Width           =   1020
   End
   Begin TabDlg.SSTab TabGasto 
      Height          =   3620
      Left            =   40
      TabIndex        =   0
      Top             =   40
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   6376
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Tipos de Actividades"
      TabPicture(0)   =   "frmMktParametrosGastos.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feTipoActividad"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "btnEditarActividad"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "btnNuevaActividad"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   " &Categoría de Gastos"
      TabPicture(1)   =   "frmMktParametrosGastos.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "feCategoria"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "btnNuevaCategoria"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "btnEditarCategoria"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton btnEditarCategoria 
         Caption         =   "&Editar"
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
         Height          =   300
         Left            =   -73800
         TabIndex        =   13
         Top             =   3240
         Width           =   1000
      End
      Begin VB.CommandButton btnNuevaCategoria 
         Caption         =   "&Nuevo"
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
         Height          =   300
         Left            =   -74880
         TabIndex        =   12
         Top             =   3240
         Width           =   1000
      End
      Begin VB.CommandButton btnNuevaActividad 
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
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   3240
         Width           =   1000
      End
      Begin VB.CommandButton btnEditarActividad 
         Caption         =   "&Editar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   11
         Top             =   3240
         Width           =   1000
      End
      Begin MSComctlLib.ListView lstAhorros 
         Height          =   2790
         Left            =   -74910
         TabIndex        =   5
         Top             =   495
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   4921
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Producto"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Agencia"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Nro. Cuenta"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Nro. Cta Antigua"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Estado"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Participación"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "SaldoCont"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "SaldoDisp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Motivo de Bloque"
            Object.Width           =   7231
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Moneda"
            Object.Width           =   2540
         EndProperty
      End
      Begin Sicmact.FlexEdit feTipoActividad 
         Height          =   2730
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   4815
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Tipo-Descripción-Activo-Id"
         EncabezadosAnchos=   "350-2500-5000-800-0"
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
         ColumnasAEditar =   "X-1-2-3-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-4-0"
         EncabezadosAlineacion=   "C-L-L-C-C"
         FormatosEdit    =   "0-1-1-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit feCategoria 
         Height          =   2730
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   4815
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Categoría-Descripción-Activo-Id"
         EncabezadosAnchos=   "350-2500-5000-800-0"
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
         ColumnasAEditar =   "X-1-2-3-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-4-0"
         EncabezadosAlineacion=   "C-L-L-C-C"
         FormatosEdit    =   "0-1-1-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   10
         Top             =   3465
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL AHORROS"
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
         Left            =   -73185
         TabIndex        =   9
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   8
         Top             =   3465
         Width           =   765
      End
      Begin VB.Label lblSolesAho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   -70815
         TabIndex        =   7
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label lblDolaresAho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   -67680
         TabIndex        =   6
         Top             =   3375
         Width           =   2145
      End
   End
End
Attribute VB_Name = "frmMktParametrosGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnOpeCod As Integer

Private Sub Form_Load()
    CentraForm Me
End Sub
Public Sub Inicio(ByVal pnInicio As Integer)
    fnOpeCod = pnInicio
    If fnOpeCod = 1 Then
        Me.btnNuevaActividad.Enabled = True
        Me.btnNuevaCategoria.Enabled = True
        Me.btnEditarActividad.Enabled = False
        Me.btnEditarCategoria.Enabled = False
        Me.btnGrabar.Enabled = False
        Me.feTipoActividad.lbEditarFlex = False
        Me.feCategoria.lbEditarFlex = False
    ElseIf fnOpeCod = 2 Then
        Me.btnNuevaActividad.Enabled = False
        Me.btnNuevaCategoria.Enabled = False
        Me.btnEditarActividad.Enabled = True
        Me.btnEditarCategoria.Enabled = True
        Me.btnGrabar.Enabled = False
        Me.feTipoActividad.lbEditarFlex = False
        Me.feCategoria.lbEditarFlex = False
    ElseIf fnOpeCod = 3 Then
        Me.btnNuevaActividad.Enabled = False
        Me.btnNuevaCategoria.Enabled = False
        Me.btnEditarActividad.Enabled = False
        Me.btnEditarCategoria.Enabled = False
        Me.btnGrabar.Enabled = False
        Me.feTipoActividad.lbEditarFlex = False
        Me.feCategoria.lbEditarFlex = False
    End If
    MuestraTipoActividad
    MuestraCategoriaGasto
    Me.Show 1
End Sub
Private Sub btnNuevaActividad_Click()
    Me.feTipoActividad.AdicionaFila
    Me.feTipoActividad.SetFocus
    SendKeys "{Enter}"
    feTipoActividad.lbEditarFlex = True
End Sub
Private Sub btnNuevaCategoria_Click()
    Me.feCategoria.AdicionaFila
    Me.feCategoria.SetFocus
    SendKeys "{Enter}"
    feCategoria.lbEditarFlex = True
End Sub
Private Sub btnEditarActividad_Click()
    feTipoActividad.lbEditarFlex = True
End Sub
Private Sub btnEditarCategoria_Click()
    feCategoria.lbEditarFlex = True
End Sub
Private Sub btnGrabar_Click()
    Dim i As Integer
    Dim oGastoMkt As DGastosMarketing
    
    If validaGrabar = False Then Exit Sub
    
    If MsgBox("Esta seguro de guardar los parametros de Gastos?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    Set oGastoMkt = New DGastosMarketing
    'Mantenimiento Tipo Actividad
    If Not (Me.feTipoActividad.Rows - 1 = 1 And Me.feTipoActividad.TextMatrix(1, 0) = "") Then
        For i = 1 To Me.feTipoActividad.Rows - 1
            If Len(Trim(Me.feTipoActividad.TextMatrix(i, 4))) = 0 Then
                Call oGastoMkt.InsertaTipoActividad(UCase(Trim(Me.feTipoActividad.TextMatrix(i, 1))), UCase(Trim(Me.feTipoActividad.TextMatrix(i, 2))), IIf(Me.feTipoActividad.TextMatrix(i, 3) = ".", True, False))
            Else
                Call oGastoMkt.ActualizaTipoActividad(CLng(Me.feTipoActividad.TextMatrix(i, 4)), UCase(Trim(Me.feTipoActividad.TextMatrix(i, 1))), UCase(Trim(Me.feTipoActividad.TextMatrix(i, 2))), IIf(Me.feTipoActividad.TextMatrix(i, 3) = ".", True, False))
            End If
        Next
    End If
    'Mantenimiento Categoría Gastos
    If Not (Me.feCategoria.Rows - 1 = 1 And Me.feCategoria.TextMatrix(1, 0) = "") Then
        For i = 1 To Me.feCategoria.Rows - 1
            If Len(Trim(Me.feCategoria.TextMatrix(i, 4))) = 0 Then
                Call oGastoMkt.Insertacategoriagasto(UCase(Trim(Me.feCategoria.TextMatrix(i, 1))), UCase(Trim(Me.feCategoria.TextMatrix(i, 2))), IIf(Me.feCategoria.TextMatrix(i, 3) = ".", True, False))
            Else
                Call oGastoMkt.Actualizacategoriagasto(CLng(Me.feCategoria.TextMatrix(i, 4)), UCase(Trim(Me.feCategoria.TextMatrix(i, 1))), UCase(Trim(Me.feCategoria.TextMatrix(i, 2))), IIf(Me.feCategoria.TextMatrix(i, 3) = ".", True, False))
            End If
        Next
    End If
    
    MuestraTipoActividad
    MuestraCategoriaGasto
    feTipoActividad.lbEditarFlex = False
    feCategoria.lbEditarFlex = False
    btnGrabar.Enabled = False
    TabGasto.Tab = 0
    
    MsgBox "Se ha grabado con éxito los Parámetros de Gastos", vbInformation, "Aviso"
    Set oGastoMkt = Nothing
End Sub

Private Sub feTipoActividad_OnCellChange(pnRow As Long, pnCol As Long)
    Dim i As Integer, J As Integer
    If pnCol = 1 Then
        For i = 1 To feTipoActividad.Rows - 1
            For J = 1 To feTipoActividad.Rows - 1
                If i <> J Then
                    If UCase(Trim(feTipoActividad.TextMatrix(i, 1))) = UCase(Trim(feTipoActividad.TextMatrix(J, 1))) Then
                        MsgBox "Tipo de Actividad ya existe", vbInformation, "Aviso"
                        feTipoActividad.Row = J
                        feTipoActividad.Col = 1
                        feTipoActividad.TextMatrix(J, 1) = ""
                        Exit Sub
                    End If
                End If
            Next J
        Next i
    End If
    feTipoActividad.TextMatrix(pnRow, pnCol) = UCase(Trim(feTipoActividad.TextMatrix(pnRow, pnCol)))
    btnGrabar.Enabled = True
End Sub
Private Sub feTipoActividad_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    btnGrabar.Enabled = True
End Sub
Private Sub feCategoria_OnCellChange(pnRow As Long, pnCol As Long)
    Dim i As Integer, J As Integer
    If pnCol = 1 Then
        For i = 1 To feCategoria.Rows - 1
            For J = 1 To feCategoria.Rows - 1
                If i <> J Then
                    If UCase(Trim(feCategoria.TextMatrix(i, 1))) = UCase(Trim(feCategoria.TextMatrix(J, 1))) Then
                        MsgBox "Categoría de Gasto ya existe", vbInformation, "Aviso"
                        feCategoria.Row = J
                        feCategoria.Col = 1
                        feCategoria.TextMatrix(J, 1) = ""
                        Exit Sub
                    End If
                End If
            Next J
        Next i
    End If
    feCategoria.TextMatrix(pnRow, pnCol) = UCase(Trim(feCategoria.TextMatrix(pnRow, pnCol)))
    btnGrabar.Enabled = True
End Sub
Private Sub feCategoria_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    btnGrabar.Enabled = True
End Sub
Private Sub MuestraTipoActividad()
    Dim oGasto As DGastosMarketing
    Dim rsTpoAct As ADODB.Recordset
    Set oGasto = New DGastosMarketing
    Set rsTpoAct = New ADODB.Recordset
    
    Call FormateaFlex(feTipoActividad)
    Set rsTpoAct = oGasto.RecuperaTipoActividad()
    If Not RSVacio(rsTpoAct) Then
        Do While Not rsTpoAct.EOF
            Me.feTipoActividad.AdicionaFila
            Me.feTipoActividad.TextMatrix(Me.feTipoActividad.Row, 1) = rsTpoAct!cNombre
            Me.feTipoActividad.TextMatrix(Me.feTipoActividad.Row, 2) = rsTpoAct!cDescripcion
            Me.feTipoActividad.TextMatrix(Me.feTipoActividad.Row, 3) = IIf(rsTpoAct!bEstado, "1", "")
            Me.feTipoActividad.TextMatrix(Me.feTipoActividad.Row, 4) = rsTpoAct!nId
            rsTpoAct.MoveNext
        Loop
    End If
    Set oGasto = Nothing
    Set rsTpoAct = Nothing
End Sub
Private Sub MuestraCategoriaGasto()
    Dim oGasto As DGastosMarketing
    Dim rsCatGast As ADODB.Recordset
    Set oGasto = New DGastosMarketing
    Set rsCatGast = New ADODB.Recordset

    Call FormateaFlex(feCategoria)
    Set rsCatGast = oGasto.RecuperaCategoriaGasto()
    If Not RSVacio(rsCatGast) Then
        Do While Not rsCatGast.EOF
            Me.feCategoria.AdicionaFila
            Me.feCategoria.TextMatrix(Me.feCategoria.Row, 1) = rsCatGast!cNombre
            Me.feCategoria.TextMatrix(Me.feCategoria.Row, 2) = rsCatGast!cDescripcion
            Me.feCategoria.TextMatrix(Me.feCategoria.Row, 3) = IIf(rsCatGast!bEstado, "1", "")
            Me.feCategoria.TextMatrix(Me.feCategoria.Row, 4) = rsCatGast!nId
            rsCatGast.MoveNext
        Loop
    End If
End Sub
Private Function validaGrabar() As Boolean
    Dim i As Integer
    validaGrabar = True
    If Not (Me.feTipoActividad.Rows - 1 = 1 And Me.feTipoActividad.TextMatrix(1, 0) = "") Then
        For i = 1 To Me.feTipoActividad.Rows - 1
            If Len(Trim(Me.feTipoActividad.TextMatrix(i, 1))) = 0 Then
                MsgBox "Falta ingresar el Tipo de Actividad", vbInformation, "Aviso"
                Me.TabGasto.Tab = 0
                Me.feTipoActividad.SetFocus
                validaGrabar = False
                Exit Function
            End If
            If Len(Trim(Me.feTipoActividad.TextMatrix(i, 2))) = 0 Then
                MsgBox "Falta ingresar la Descripción del Tipo de Actividad", vbInformation, "Aviso"
                Me.TabGasto.Tab = 0
                Me.feTipoActividad.SetFocus
                validaGrabar = False
                Exit Function
            End If
        Next
    End If
    If Not (Me.feCategoria.Rows - 1 = 1 And Me.feCategoria.TextMatrix(1, 0) = "") Then
        For i = 1 To Me.feCategoria.Rows - 1
            If Len(Trim(Me.feCategoria.TextMatrix(i, 1))) = 0 Then
                MsgBox "Falta ingresar la Categoría de Gasto", vbInformation, "Aviso"
                Me.TabGasto.Tab = 1
                Me.feCategoria.SetFocus
                validaGrabar = False
                Exit Function
            End If
            If Len(Trim(Me.feCategoria.TextMatrix(i, 2))) = 0 Then
                MsgBox "Falta ingresar la Descripción de la Categoría de Gasto", vbInformation, "Aviso"
                Me.TabGasto.Tab = 1
                Me.feCategoria.SetFocus
                validaGrabar = False
                Exit Function
            End If
        Next
    End If
End Function
Private Sub btnSalir_Click()
    Unload Me
End Sub
