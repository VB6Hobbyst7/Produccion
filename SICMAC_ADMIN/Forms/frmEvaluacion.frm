VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRHEvaluacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Evaluación"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "frmEvaluacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   4410
      TabIndex        =   38
      Top             =   6165
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame fraProp 
      Caption         =   "Propiedades"
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
      Height          =   660
      Left            =   30
      TabIndex        =   22
      Top             =   435
      Width           =   7905
      Begin VB.ComboBox cmbEstado 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   255
         Width           =   2610
      End
      Begin VB.ComboBox cmbModalidad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4665
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2970
      End
      Begin VB.Label lblEstado 
         Caption         =   "Estado :"
         Height          =   255
         Left            =   135
         TabIndex        =   24
         Top             =   270
         Width           =   735
      End
      Begin VB.Label lblModalidad 
         Caption         =   "Modalidad :"
         Height          =   255
         Left            =   3780
         TabIndex        =   23
         Top             =   270
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   6405
      Top             =   6165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "N&uevo"
      Height          =   375
      Left            =   60
      TabIndex        =   13
      Top             =   6165
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   60
      TabIndex        =   14
      Top             =   6165
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6960
      TabIndex        =   19
      Top             =   6165
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3315
      TabIndex        =   18
      Top             =   6165
      Width           =   975
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2220
      TabIndex        =   17
      Top             =   6165
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1140
      TabIndex        =   15
      Top             =   6165
      Width           =   975
   End
   Begin VB.ComboBox cmbEval 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1140
      TabIndex        =   16
      Top             =   6165
      Width           =   975
   End
   Begin VB.Frame fraEva 
      Caption         =   "Datos"
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
      Height          =   5040
      Left            =   30
      TabIndex        =   21
      Top             =   1080
      Width           =   7905
      Begin VB.Frame fraCom 
         Caption         =   "Comite"
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
         Height          =   2550
         Left            =   4035
         TabIndex        =   35
         Top             =   2385
         Width           =   3780
         Begin VB.CommandButton cmdMantenimientoCom 
            Caption         =   "&Mantenimiento >>"
            Height          =   375
            Left            =   1020
            TabIndex        =   37
            Top             =   2085
            Width           =   1755
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexComite 
            Height          =   1800
            Left            =   105
            TabIndex        =   36
            Top             =   225
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   3175
            _Version        =   393216
            FixedCols       =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraDatosEva 
         Caption         =   "Evaluación"
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
         Height          =   1530
         Left            =   135
         TabIndex        =   30
         Top             =   165
         Width           =   7680
         Begin VB.ComboBox cmbTipo 
            Height          =   315
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   735
            Width           =   2550
         End
         Begin VB.TextBox txtDes 
            Appearance      =   0  'Flat
            Height          =   525
            Left            =   1050
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   180
            Width           =   6495
         End
         Begin VB.ComboBox cmbArea 
            Height          =   315
            Left            =   4545
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   735
            Width           =   3000
         End
         Begin VB.ComboBox cmbCargo 
            Height          =   315
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1095
            Width           =   6525
         End
         Begin VB.Label lblDes 
            Caption         =   "Descripción :"
            Height          =   255
            Left            =   90
            TabIndex        =   34
            Top             =   195
            Width           =   855
         End
         Begin VB.Label lblCargo 
            Caption         =   "Cargo :"
            Height          =   255
            Left            =   90
            TabIndex        =   33
            Top             =   1125
            Width           =   735
         End
         Begin VB.Label lblArea 
            Caption         =   "Area :"
            Height          =   255
            Left            =   3825
            TabIndex        =   32
            Top             =   810
            Width           =   495
         End
         Begin VB.Label lblTipo 
            Caption         =   "Tipo"
            Height          =   255
            Left            =   90
            TabIndex        =   31
            Top             =   780
            Width           =   615
         End
      End
      Begin VB.Frame fraFechas 
         Caption         =   "Fechas"
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
         Height          =   645
         Left            =   90
         TabIndex        =   25
         Top             =   1725
         Width           =   7725
         Begin MSMask.MaskEdBox mskFF 
            Height          =   315
            Left            =   4650
            TabIndex        =   8
            Top             =   255
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFI 
            Height          =   315
            Left            =   1020
            TabIndex        =   7
            Top             =   225
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblFI 
            Caption         =   "Inicio :"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   270
            Width           =   855
         End
         Begin VB.Label lblFF 
            Caption         =   "Fin :"
            Height          =   255
            Left            =   3930
            TabIndex        =   26
            Top             =   270
            Width           =   735
         End
      End
      Begin VB.Frame fratexto 
         Caption         =   "Texto"
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
         Height          =   2550
         Left            =   90
         TabIndex        =   28
         Top             =   2385
         Width           =   3795
         Begin TabDlg.SSTab SSTab1 
            Height          =   2280
            Left            =   90
            TabIndex        =   9
            Top             =   180
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   4022
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            BackColor       =   0
            TabCaption(0)   =   "Escrito"
            TabPicture(0)   =   "frmEvaluacion.frx":030A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "REscrito"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "cmdCargarE"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Psicologico"
            TabPicture(1)   =   "frmEvaluacion.frx":0326
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "cmdCargarP"
            Tab(1).Control(1)=   "RPsicologico"
            Tab(1).ControlCount=   2
            Begin VB.CommandButton cmdCargarP 
               Caption         =   "&Cargar"
               Height          =   375
               Left            =   -73657
               TabIndex        =   12
               Top             =   1800
               Width           =   975
            End
            Begin VB.CommandButton cmdCargarE 
               Caption         =   "&Cargar"
               Height          =   375
               Left            =   1343
               TabIndex        =   10
               Top             =   1800
               Width           =   975
            End
            Begin RichTextLib.RichTextBox REscrito 
               Height          =   1365
               Left            =   120
               TabIndex        =   29
               Top             =   390
               Width           =   3420
               _ExtentX        =   6033
               _ExtentY        =   2408
               _Version        =   393217
               Enabled         =   -1  'True
               Appearance      =   0
               TextRTF         =   $"frmEvaluacion.frx":0342
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin RichTextLib.RichTextBox RPsicologico 
               Height          =   1350
               Left            =   -74880
               TabIndex        =   11
               Top             =   405
               Width           =   3420
               _ExtentX        =   6033
               _ExtentY        =   2381
               _Version        =   393217
               Enabled         =   -1  'True
               Appearance      =   0
               TextRTF         =   $"frmEvaluacion.frx":040A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
      End
   End
   Begin VB.Label lbl 
      Caption         =   "Evaluación :"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "frmRHEvaluacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnTipo As TipoOpe
Dim lnModalidad As RHProcesoSeleccionModal
Dim lsCodigo As String
Dim lbEditado As Boolean
Dim loVar As RHProcesoSeleccionTipo
Dim lbCerrar As Boolean

Private Sub CargaDatos(pbTodos As Boolean)
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    Set rsE = oEva.GetProcesosSelecion(Trim(Str(lnModalidad)), lbCerrar)
    CargaCombo rsE, Me.cmbEval, 200, 1, 0
    rsE.Close
    Set oEva = Nothing
       
    CargaComite
        
    If pbTodos Then
        Dim oCargo As DActualizadatosCargo
        Set oCargo = New DActualizadatosCargo
        Set rsE = oCargo.GetCargos(False)
        CargaCombo rsE, Me.cmbCargo, 150, 2, 1
        rsE.Close
        Set rsE = Nothing
        Set oCargo = Nothing
        
        Dim oArea As DActualizaDatosArea
        Set oArea = New DActualizaDatosArea
        Set rsE = New ADODB.Recordset
        Set rsE = oArea.GetAreasOrg
        CargaCombo rsE, Me.cmbArea
        rsE.Close
        Set oArea = Nothing
        
        Dim oCons As DConstantes
        Set oCons = New DConstantes
        Set rsE = New ADODB.Recordset
        Set rsE = oCons.GetConstante(gRHProcesoSeleccionTipo)
        CargaCombo rsE, Me.cmbTipo
        rsE.Close
        Set oCons = Nothing
    
        Set oCons = New DConstantes
        Set rsE = New ADODB.Recordset
        Set rsE = oCons.GetConstante(gRHProcesoSeleccionModal)
        CargaCombo rsE, Me.cmbModalidad
        rsE.Close
        Set oCons = Nothing
    
        Set oCons = New DConstantes
        Set rsE = New ADODB.Recordset
        Set rsE = oCons.GetConstante(gRHProcesoSeleccionEstado)
        CargaCombo rsE, Me.cmbEstado
        rsE.Close
        Set oCons = Nothing
    End If
End Sub

Private Sub cmbArea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmbCargo.SetFocus
    End If
End Sub

Private Sub cmbCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFI.SetFocus
    End If
End Sub

Private Sub cmbEval_Click()
    Dim oEva As DActualizaProcesoSeleccion
    Dim rsE As New ADODB.Recordset
    Set oEva = New DActualizaProcesoSeleccion
    
    Set rsE = oEva.GetProcesoSelecion(Trim(Right(Me.cmbEval.Text, 6)))
    If Not (rsE.EOF And rsE.BOF) Then
        Me.txtDes.Text = rsE!Comentario
        UbicaCombo Me.cmbArea, rsE!Area
        UbicaCombo Me.cmbTipo, rsE!Tipo
        UbicaCombo Me.cmbTipo, rsE!Modo
        UbicaCombo Me.cmbCargo, rsE!Cargo
        UbicaCombo Me.cmbEstado, rsE!Estado
        Me.mskFI.Text = Format(rsE!fi, gsFormatoFechaView)
        Me.mskFF.Text = Format(rsE!FF, gsFormatoFechaView)
        Me.REscrito.Text = rsE!escrito
        Me.RPsicologico.Text = rsE!Psico
        CargaComite
    Else
        Me.FlexComite.Rows = 1
        Me.FlexComite.Rows = 2
        Me.FlexComite.FixedRows = 1
    End If
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmbArea.SetFocus
    End If
End Sub

Private Sub cmdCerrar_Click()
    lbEditado = True
    If MsgBox("Desea Cerrar proceso Seleccion ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    UbicaCombo Me.cmbEstado, RHProcesoSeleccionEstado.gRHProcSelEstFinalizado
    cmdGrabar_Click
End Sub

Private Sub cmdMantenimientoCom_Click()
    If Me.cmbEval.Text = "" Then
        MsgBox "No puede agregar a los miembros del comite sin haber creado la evaluacion.", vbInformation, "Aviso"
        Exit Sub
    End If
    frmRHComite.Ini (Right(Me.cmbEval.Text, 6))
    CargaComite
End Sub

Private Sub cmdCancelar_Click()
    Limpia
    Activa False, lnTipo
    If lnTipo = gTipoOpeRegistro Then
        Unload Me
    End If
End Sub

Private Sub cmdCargarE_Click()
    CDialog.CancelError = False
    CDialog.Flags = cdlOFNHideReadOnly
    CDialog.Filter = "Archivos txt(*.txt)|*.txt"
    CDialog.FilterIndex = 2
    CDialog.ShowOpen
    Me.REscrito.LoadFile CDialog.FileName, 1
End Sub

Private Sub cmdCargarP_Click()
    CDialog.CancelError = False
    CDialog.Flags = cdlOFNHideReadOnly
    CDialog.Filter = "Archivos txt(*.txt)|*.txt"
    CDialog.FilterIndex = 2
    CDialog.ShowOpen
    Me.RPsicologico.LoadFile CDialog.FileName, 1
End Sub

Private Sub cmdEditar_Click()
    If Me.cmbEval.Text = "" Then
        Me.cmbEval.SetFocus
        Exit Sub
    End If
    lbEditado = True
    Activa True, lnTipo
End Sub

Private Sub cmbEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbModalidad.Enabled Then
            Me.cmbModalidad.SetFocus
        Else
            Me.txtDes.SetFocus
        End If
    End If
End Sub

Private Sub cmdEliminar_Click()
    Dim oEva As NActualizaProcesoSeleccion
    Set oEva = New NActualizaProcesoSeleccion
    
    If MsgBox("Desea Elimiar la Evaluación. " & Trim(Left(Me.cmbEval.Text, 50)) & Chr(13) & "Se eliminaran todas las personas relacionadas con esta Evaluacion.", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    oEva.EliminaProSelec Trim(Right(Me.cmbEval.List(cmbEval.ListCount - 1), 7))
    CargaDatos False
    Limpia
End Sub

Private Sub cmdGrabar_Click()
    Dim oEval As NActualizaProcesoSeleccion
    Set oEval = New NActualizaProcesoSeleccion
    If Not Valida() Then Exit Sub
        
    If lbEditado Then
        oEval.ModificaProSelec Right(Me.cmbEval, 6), Right(Me.cmbTipo, 1), Trim(Str(lnModalidad)), Right(Me.cmbArea, 3), Right(Me.cmbCargo, 6), Format(Me.mskFI.Text, gsFormatoFecha), Format(Me.mskFF.Text, gsFormatoFecha), Right(cmbEstado, 1), Me.txtDes.Text, Me.RPsicologico.Text, Me.REscrito.Text, GetMovNro(gsCodUser, gsCodAge)
    Else
        oEval.AgregaProSelec lsCodigo, Right(Me.cmbTipo, 1), Trim(Str(lnModalidad)), Right(Me.cmbArea, 3), Right(Me.cmbCargo, 6), Format(Me.mskFI.Text, gsFormatoFecha), Format(Me.mskFF.Text, gsFormatoFecha), Right(cmbEstado, 1), Me.txtDes.Text, Me.RPsicologico.Text, Me.REscrito.Text, GetMovNro(gsCodUser, gsCodAge)
    End If
        
    Limpia
    Activa False, lnTipo
    CargaDatos False

    If lnTipo = gTipoOpeRegistro Then
        Unload Me
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim oEval As NActualizaProcesoSeleccion
    Dim lsCadena As String
    Dim lsCadenaTemp As String
    Dim lbRep(1) As Boolean
    Dim oPrevio As Previo.clsPrevio
    Set oEval = New NActualizaProcesoSeleccion
    Set oPrevio = New Previo.clsPrevio
    
    frmImpreRRHH.Ini "Lista de Examenes;", "Evaluaicon", lbRep, gdFecSis, gdFecSis, False
    
    If lbRep(1) Then
        lsCadena = oEval.GetReporte(gsNomAge, gsEmpresa, gdFecSis)
    End If
    
    'If lbRep(2) Then
    '
    'End If
    
    If lsCadena <> "" Then oPrevio.Show lsCadena, " Evaluaciones ", True, 66
    Set oEval = Nothing
    Set oPrevio = Nothing
End Sub

Private Sub cmdNuevo_Click()
    If Trim(Left(Me.cmbEval.List(cmbEval.ListCount - 1), 7)) = "" Then
        lsCodigo = "000001"
    Else
        lsCodigo = Format(CCur(Trim(Right(Me.cmbEval.List(cmbEval.ListCount - 1), 7))) + 1, "000000")
    End If
    Me.cmbEval.ListIndex = -1
    lbEditado = False
    Limpia
    Activa True, lnTipo
    
    If lnTipo <> gTipoOpeRegistro Then Me.cmbEstado.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If lnTipo = gTipoOpeRegistro Then Me.cmbEstado.SetFocus
End Sub

Private Sub Form_Load()
    CargaDatos True
    Activa False, gTipoOpeRegistro
    
    If lnTipo = gTipoOpeRegistro Then
        cmdNuevo_Click
        Me.cmdEliminar.Enabled = False
        Me.cmdEditar.Enabled = False
        Me.cmdImprimir.Enabled = False
    ElseIf lnTipo = gTipoOpeConsulta Then
        Me.cmdEliminar.Enabled = False
        Me.cmdNuevo.Enabled = False
        Me.cmdEditar.Enabled = False
        Me.cmdImprimir.Enabled = False
        Me.fraDatosEva.Enabled = False
        fraFechas.Enabled = False
        Me.fraEva.Enabled = True
        Me.cmdCargarE.Enabled = False
        Me.cmdCargarP.Enabled = False
        Me.cmdCerrar.Visible = lbCerrar
    ElseIf lnTipo = gTipoOpeReporte Then
        Me.cmdEliminar.Enabled = False
        Me.cmdNuevo.Enabled = False
        Me.cmdEditar.Enabled = False
        Me.fraDatosEva.Enabled = False
        fraFechas.Enabled = False
    End If
    
    UbicaCombo Me.cmbModalidad, Str(lnModalidad)
End Sub

Public Sub Ini(pnTipo As TipoOpe, pnModalidad As RHProcesoSeleccionModal)
    lbCerrar = False
    lnTipo = pnTipo
    lnModalidad = pnModalidad
    Me.Show 1
End Sub

Private Sub mskFF_GotFocus()
    mskFF.SelStart = 0
    mskFF.SelLength = 10
End Sub

Private Sub mskFF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.SSTab1.SetFocus
    End If
End Sub

Private Sub mskFI_GotFocus()
    mskFI.SelStart = 0
    mskFI.SelLength = 10
End Sub

Private Sub Activa(pbValor As Boolean, pnTipo As TipoOpe)
    'Objetos
    Me.cmbEval.Enabled = Not pbValor
    
    If pnTipo = gTipoOpeRegistro Or pnTipo = gTipoOpeMantenimiento Then
        Me.fraEva.Enabled = pbValor
        Me.fraProp.Enabled = pbValor
    ElseIf pnTipo = gTipoOpeConsulta Or pnTipo = gTipoOpeReporte Then
        Me.txtDes.Enabled = pbValor
        Me.cmbArea.Enabled = pbValor
        Me.cmbCargo.Enabled = pbValor
        Me.cmbTipo.Enabled = pbValor
        Me.mskFI.Enabled = pbValor
        Me.mskFF.Enabled = pbValor
    End If
    
    Me.cmdNuevo.Visible = Not pbValor
    Me.cmdEditar.Visible = Not pbValor
    Me.cmdGrabar.Visible = pbValor
    Me.cmdCancelar.Visible = pbValor
End Sub

Private Sub mskFI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFF.SetFocus
    End If
End Sub

Private Sub REscrito_DblClick()
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    oPrevio.Show REscrito.Text, "Examen Escrito", True, 66
    Set oPrevio = Nothing
End Sub

Private Sub RPsicologico_DblClick()
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    oPrevio.Show RPsicologico.Text, "Examen Psicologico", True, 66
    Set oPrevio = Nothing
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.SSTab1.Tab = 0 Then
            Me.cmdCargarE.SetFocus
        Else
            Me.cmdCargarP.SetFocus
        End If
    End If

End Sub

Private Sub txtDes_GotFocus()
    txtDes.SelStart = 0
    txtDes.SelLength = 50
End Sub

Private Sub txtDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmbTipo.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub Limpia()
    Me.SSTab1.Tab = 0
    Me.cmbEval.ListIndex = -1
    Me.cmbArea.ListIndex = -1
    Me.cmbCargo.ListIndex = -1
    Me.cmbTipo.ListIndex = -1
    Me.cmbEstado.ListIndex = -1
    Me.txtDes.Text = ""
    Me.mskFI.Text = "__/__/____"
    Me.mskFF.Text = "__/__/____"
    Me.REscrito.Text = ""
    Me.RPsicologico.Text = ""
End Sub

Private Function Valida() As Boolean
    If Me.cmbEstado.Text = "" Then
        MsgBox "Debe Elegir un Estado.", vbInformation, "Aviso"
        cmbEstado.SetFocus
        Valida = False
    ElseIf Me.txtDes.Text = "" Then
        MsgBox "Debe Ingresar una descripción.", vbInformation, "Aviso"
        txtDes.SetFocus
        Valida = False
    ElseIf Me.cmbTipo.Text = "" Then
        MsgBox "Debe Elegir un Tipo.", vbInformation, "Aviso"
        cmbTipo.SetFocus
        Valida = False
    ElseIf Me.cmbArea.Text = "" Then
        MsgBox "Debe Elegir un Area.", vbInformation, "Aviso"
        cmbArea.SetFocus
        Valida = False
    ElseIf Not IsDate(Me.mskFI.Text) Then
        MsgBox "Debe Ingresar una Fecha Valida.", vbInformation, "Aviso"
        mskFI.SetFocus
        Valida = False
    ElseIf Not IsDate(Me.mskFF.Text) Then
        MsgBox "Debe Ingresar una Fecha Valida.", vbInformation, "Aviso"
        mskFF.SetFocus
        Valida = False
    ElseIf Me.cmbCargo.Text = "" Then
        MsgBox "Debe Elegir un Cargo.", vbInformation, "Aviso"
        cmbCargo.SetFocus
        Valida = False
    ElseIf Me.REscrito.Text = "" Then
        MsgBox "Debe Ingresar el texto del examen escrito.", vbInformation, "Aviso"
        Me.SSTab1.Tab = 0
        Me.cmdCargarE.SetFocus
        Valida = False
    ElseIf Me.RPsicologico.Text = "" Then
        MsgBox "Debe Ingresar el texto del examen Psocilogico.", vbInformation, "Aviso"
        Me.SSTab1.Tab = 1
        Me.cmdCargarP.SetFocus
        Valida = False
    Else
        Valida = True
    End If
End Function

Private Sub CargaComite()
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    Set Me.FlexComite.DataSource = oEva.GetNomPersonasComite(Right(Me.cmbEval.Text, 6))
    Me.FlexComite.ColWidth(0) = 1
    Me.FlexComite.ColWidth(1) = 3000
    Me.FlexComite.ColWidth(2) = 3000
    Set oEva = Nothing
End Sub

Public Sub IniCerrar(pnModalidad As RHProcesoSeleccionModal)
    lbCerrar = True
    lnTipo = gTipoOpeConsulta
    lnModalidad = pnModalidad
    Me.Show 1
End Sub
