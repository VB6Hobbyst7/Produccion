VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRHConceptoAsigPla 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   Icon            =   "frmRHConceptoAsigPla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbTipConcepto 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2280
      Width           =   5055
   End
   Begin VB.Frame fraEstadoValidos 
      Appearance      =   0  'Flat
      Caption         =   "Estados Validos"
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
      Height          =   2580
      Left            =   6990
      TabIndex        =   33
      Top             =   15
      Width           =   2325
      Begin MSComctlLib.ListView lvwG 
         Height          =   2235
         Left            =   120
         TabIndex        =   9
         Top             =   225
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   3942
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   45
      TabIndex        =   16
      Top             =   6180
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1245
      TabIndex        =   17
      Top             =   6180
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2445
      TabIndex        =   20
      Top             =   6180
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3645
      TabIndex        =   21
      Top             =   6180
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   380
      Left            =   8265
      TabIndex        =   22
      Top             =   6180
      Width           =   1050
   End
   Begin VB.Frame fraG 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   7035
      TabIndex        =   23
      Top             =   70
      Width           =   2175
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   15
      TabIndex        =   18
      Top             =   6180
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1215
      TabIndex        =   19
      Top             =   6180
      Width           =   1095
   End
   Begin VB.Frame fraDatos 
      Appearance      =   0  'Flat
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
      Height          =   3570
      Left            =   15
      TabIndex        =   29
      Top             =   2580
      Width           =   9300
      Begin VB.CommandButton cmdIzqUno 
         Caption         =   "&<"
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
         Left            =   4455
         TabIndex        =   15
         ToolTipText     =   "Mueve un Conceto a la Izquierda (Alt + <)"
         Top             =   2205
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdIzqTodos 
         Caption         =   "<<"
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
         Left            =   4455
         TabIndex        =   30
         ToolTipText     =   "Mueve todos los Conceptos a la Izquierda"
         Top             =   1845
         Width           =   375
      End
      Begin VB.CommandButton cmdDerTodos 
         Caption         =   ">>"
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
         Left            =   4455
         TabIndex        =   14
         ToolTipText     =   "Mueve todos los Conceptos a la Derecha"
         Top             =   1485
         Width           =   375
      End
      Begin VB.CommandButton cmdDerUno 
         Caption         =   "&>"
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
         Left            =   4455
         TabIndex        =   13
         ToolTipText     =   "Mueve un Conceto a la Derecha (Alt + >)"
         Top             =   1125
         Width           =   375
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexLibres 
         Height          =   3240
         Left            =   75
         TabIndex        =   11
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5715
         _Version        =   393216
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexAdd 
         Height          =   3240
         Left            =   4875
         TabIndex        =   12
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5715
         _Version        =   393216
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin Sicmact.TxtBuscar TxtPlanilla 
      Height          =   315
      Left            =   1335
      TabIndex        =   39
      Top             =   90
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Appearance      =   0
      BackColor       =   -2147483624
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
      sTitulo         =   ""
   End
   Begin VB.Frame fraPlanilla 
      Appearance      =   0  'Flat
      Caption         =   "Planillas"
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
      Height          =   1845
      Left            =   30
      TabIndex        =   24
      Top             =   390
      Width           =   6855
      Begin VB.TextBox txtOpeCont 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5580
         TabIndex        =   3
         Top             =   855
         Width           =   1125
      End
      Begin VB.TextBox txtOpeEst 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3585
         TabIndex        =   2
         Top             =   855
         Width           =   1125
      End
      Begin VB.TextBox txtOpeAbono 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1275
         TabIndex        =   1
         Top             =   855
         Width           =   1125
      End
      Begin Sicmact.TxtBuscar TxtTipoPlanilla 
         Height          =   315
         Left            =   1275
         TabIndex        =   36
         Top             =   195
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Appearance      =   0
         BackColor       =   -2147483624
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
         sTitulo         =   ""
      End
      Begin VB.TextBox txtParametro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1290
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1500
         Width           =   4410
      End
      Begin VB.TextBox txtPlaDes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1275
         TabIndex        =   0
         Top             =   540
         Width           =   5430
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Activo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5940
         TabIndex        =   8
         Top             =   1545
         Width           =   765
      End
      Begin MSMask.MaskEdBox mskFecRef 
         Height          =   270
         Left            =   1290
         TabIndex        =   4
         Top             =   1185
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskIni 
         Height          =   270
         Left            =   3600
         TabIndex        =   5
         Top             =   1185
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFin 
         Height          =   270
         Left            =   5985
         TabIndex        =   6
         Top             =   1185
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblOpeCon 
         Caption         =   "Ope Cont :"
         Height          =   255
         Left            =   4800
         TabIndex        =   42
         Top             =   900
         Width           =   810
      End
      Begin VB.Label lblOpeAbono 
         Caption         =   "Ope Abono :"
         Height          =   255
         Left            =   105
         TabIndex        =   41
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label lblOpeEst 
         Caption         =   "Ope Estable :"
         Height          =   255
         Left            =   2610
         TabIndex        =   40
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lblTipoPlanillaRes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2730
         TabIndex        =   37
         Top             =   210
         Width           =   3960
      End
      Begin VB.Label lblParametro 
         Caption         =   "Parametro :"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   1545
         Width           =   1200
      End
      Begin VB.Label lblInicio 
         Caption         =   "Inicio :"
         Height          =   255
         Left            =   2610
         TabIndex        =   31
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label lblPlaDes 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label lblTipCon 
         Caption         =   "Tipo Contrato"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   225
         Width           =   1125
      End
      Begin VB.Label lblFecRef 
         Caption         =   "Fecha Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1193
         Width           =   1335
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Fin :"
         Height          =   255
         Left            =   5250
         TabIndex        =   32
         Top             =   1193
         Width           =   420
      End
   End
   Begin VB.Label lblPlanillaRes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2775
      TabIndex        =   38
      Top             =   105
      Width           =   4110
   End
   Begin VB.Label lblTipConceptos 
      Caption         =   "Tipo de Concepto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   105
      TabIndex        =   28
      Top             =   2332
      Width           =   1695
   End
   Begin VB.Label lblPlaCod 
      Caption         =   "Codigo Planilla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   45
      TabIndex        =   35
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmRHConceptoAsigPla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbEditado As Boolean
Dim lnTipo As TipoOpe

Private Sub chkActivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.lvwG.SetFocus
    End If
End Sub

Private Sub cmbPlaCod_Change()
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    Dim llAux As ListItem
    Dim lsCodEmp As String
    
    
    If TxtPlanilla.Text = "" Then Exit Sub
    'ClearScreen
    
    Set rsP = oPla.GetPlanilla(Right(TxtPlanilla, 3))
    TxtTipoPlanilla = Left(Right(TxtPlanilla, 3), 1)
    Me.lblTipoPlanillaRes.Caption = TxtTipoPlanilla.psDescripcion
    
    If Not (rsP.BOF And rsP.BOF) Then
        txtPlaDes = rsP!Descrip & ""
        
        If Not rsP!Estado Then
            chkActivo.value = 0
        Else
            chkActivo.value = 1
        End If
        
        mskFecRef = Format(rsP!Pago, "dd")
        Me.mskIni.Text = Format(rsP!Inicio, "dd")
        Me.mskFin.Text = Format(rsP!Fin, "dd")
        Me.txtParametro.Text = rsP!parametro
        
        Me.txtOpeAbono.Text = rsP!cOpeAbono & ""
        Me.txtOpeCont.Text = rsP!cOpeCodCont & ""
        Me.txtOpeEst.Text = rsP!cOpeCodEst & ""
        
        'Me.txtCta.Text = rsP!SubCuenta
        rsP.Close
    End If
    
    'Libres
    Set rsP = oPla.GetPlanillaConcepLibres(Me.TxtPlanilla.Text, Right(Me.cmbTipConcepto, 1))
    Set FlexLibres.DataSource = rsP
    rsP.Close
    
    Set rsP = oPla.GetPlanillaConcepUsados(TxtPlanilla.Text, Right(Me.cmbTipConcepto, 1))
    Set FlexAdd.DataSource = rsP
    rsP.Close

    Set rsP = oPla.GetPlanillaConcepUsados(TxtPlanilla.Text, Right(Me.cmbTipConcepto, 1))
    Set FlexAdd.DataSource = rsP
    rsP.Close

    Set rsP = oPla.GetPlanillaEstadosUsados(TxtPlanilla.Text)

    lvwG.HideColumnHeaders = False
    lvwG.ColumnHeaders.Clear
    lvwG.ColumnHeaders.Add , , "Tipo de Persona", 3500
    lvwG.ListItems.Clear
    If Not RSVacio(rsP) Then
        While Not rsP.EOF
            Set llAux = lvwG.ListItems.Add(, , Trim(rsP!Descrip & Space(30) & rsP!Valor))
            If Not IsNull(rsP!Estado) Then
                llAux.Checked = True
            Else
                llAux.Checked = False
            End If
            rsP.MoveNext
        Wend
    End If

    rsP.Close
    Set rsP = Nothing
End Sub

Private Sub cmbPlaCod_Click()
    cmbPlaCod_Change
End Sub

Private Sub cmbPlaCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbTipConcepto.Enabled Then
            cmbTipConcepto.SetFocus
        Else
            Me.cmbTipConcepto.SetFocus
        End If
    End If
End Sub

Private Sub cmbTipCon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtPlaDes.SetFocus
End Sub

Private Sub cmbTipConcepto_Change()
    cmbPlaCod_Change
End Sub

Private Sub cmbTipConcepto_Click()
    cmbPlaCod_Change
End Sub

Private Sub cmbTipConcepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdDerUno.Enabled Then
            Me.cmdDerUno.SetFocus
        Else
            Me.cmdEditar.SetFocus
        End If
    End If
End Sub

Private Sub CmdCancelar_Click()
    Activa False
    fraG.Enabled = False
    Me.TxtPlanilla.Text = ""
    Me.lblPlanillaRes.Caption = ""
    Me.TxtPlanilla.Enabled = True
    ClearScreen
End Sub

Private Sub cmdDerTodos_Click()
    Dim i As Integer
    Dim sqlI As String
    Dim lsHoy As String
    Dim lnTope As Integer
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    
    lsHoy = FechaHora(gdFecSis)
    lnTope = FlexLibres.Rows - 1
    
    For i = 1 To lnTope
        FlexLibres.Row = 1
        
        If FlexLibres.TextMatrix(1, 0) <> "" Then
            oPla.AgregaPlanillaConcepto TxtPlanilla.Text, FlexLibres.TextMatrix(1, 0), GetMovNro(gsCodUser, gsCodAge)
            
            If FlexAdd.TextMatrix(FlexAdd.Rows - 1, 0) <> "" Then FlexAdd.Rows = FlexAdd.Rows + 1
            FlexAdd.RowHeight(FlexAdd.Rows - 1) = 240
            FlexAdd.TextMatrix(FlexAdd.Rows - 1, 0) = FlexLibres.TextMatrix(1, 0)
            FlexAdd.TextMatrix(FlexAdd.Rows - 1, 1) = FlexLibres.TextMatrix(1, 1)
            
            If FlexLibres.Rows <> 2 Then
                FlexLibres.RemoveItem 1
            Else
                FlexLibres.TextMatrix(1, 0) = ""
                FlexLibres.TextMatrix(1, 1) = ""
            End If
        End If
    Next i
    FlexLibres.Refresh
End Sub

Private Sub cmdDerUno_Click()
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    
    If FlexLibres.TextMatrix(FlexLibres.Row, 1) = "" Then Exit Sub
    
    oPla.AgregaPlanillaConcepto TxtPlanilla.Text, FlexLibres.TextMatrix(FlexLibres.Row, 0), GetMovNro(gsCodUser, gsCodAge)
    
    If FlexAdd.TextMatrix(FlexAdd.Rows - 1, 0) <> "" Then FlexAdd.Rows = FlexAdd.Rows + 1
    FlexAdd.RowHeight(FlexAdd.Rows - 1) = 240
    FlexAdd.TextMatrix(FlexAdd.Rows - 1, 0) = FlexLibres.TextMatrix(FlexLibres.Row, 0)
    FlexAdd.TextMatrix(FlexAdd.Rows - 1, 1) = FlexLibres.TextMatrix(FlexLibres.Row, 1)
    
    If FlexLibres.Rows <> 2 Then
        FlexLibres.RemoveItem FlexLibres.Row
    Else
        FlexLibres.TextMatrix(FlexLibres.Row, 0) = ""
        FlexLibres.TextMatrix(FlexLibres.Row, 1) = ""
    End If
    FlexLibres.Refresh
End Sub

Private Sub cmdEditar_Click()
    If TxtPlanilla.Text = "" Then Exit Sub
    Activa True
    fraG.Enabled = True
    Me.TxtPlanilla.Enabled = False
    UbicaCombo Me.cmbTipConcepto, "5"
    lbEditado = True
    Me.cmbTipConcepto.Enabled = True
    Me.txtPlaDes.SetFocus
End Sub

Private Sub cmdEliminar_Click()
    Dim sqlI As String
    Dim lsOpc As String
    Dim oPlan As DActualizaDatosConPlanilla
    Set oPlan = New DActualizaDatosConPlanilla
    
    If MsgBox("Desea Eliminar una Planilla, esta se eliminara logicamente.", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    lsOpc = IIf(chkActivo.value = 1, "1", "0")
    oPlan.EliminaConceptoPlanilla TxtPlanilla.Text
    cmbPlaCod_Change
End Sub

Private Sub cmdGrabar_Click()
    Dim sqlI As String
    Dim lsOpc As String
    Dim oPla As NActualizaDatosConPlanilla
    Dim lsCod As String
    Dim i As Integer
    Set oPla = New NActualizaDatosConPlanilla
    Dim oPlaT As DActualizaDatosConPlanilla
    Set oPlaT = New DActualizaDatosConPlanilla
    Dim rsPlaT As ADODB.Recordset
    Set rsPlaT = New ADODB.Recordset
    
    lsOpc = IIf(chkActivo.value = 1, "1", "0")
    
    If Me.TxtTipoPlanilla.Text = "" Then
        MsgBox "Debe un tipo de Planilla.", vbInformation, "Aviso"
        Me.TxtTipoPlanilla.SetFocus
        Exit Sub
    ElseIf Me.txtPlaDes.Text = "" Then
        MsgBox "Debe ingrear una referencia de la Planilla.", vbInformation, "Aviso"
        Me.txtPlaDes.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(mskFecRef) Then
        MsgBox "Debe Ingresar una Fecha de Referencia Valida.", vbInformation, "Aviso"
        Me.mskFecRef.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(Me.mskIni.Text) Then
        MsgBox "Debe Ingresar una Fecha Valida.", vbInformation, "Aviso"
        Me.mskIni.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(Me.mskFin.Text) Then
        MsgBox "Debe Ingresar una Fecha Valida.", vbInformation, "Aviso"
        Me.mskFin.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(Me.txtParametro.Text) Then
        MsgBox "Debe Ingresar un numero valido.", vbInformation, "Aviso"
        Me.txtParametro.SetFocus
        Exit Sub
    ElseIf Me.txtOpeAbono.Text = "" Then
        MsgBox "Debe Ingresar una operacion de Abono.", vbInformation, "Aviso"
        Me.txtOpeAbono.SetFocus
        Exit Sub
    ElseIf Me.txtOpeEst.Text = "" Then
        MsgBox "Debe Ingresar una operacion de Planilla de Estables.", vbInformation, "Aviso"
        Me.txtOpeEst.SetFocus
        Exit Sub
    ElseIf Me.txtOpeCont.Text = "" Then
        MsgBox "Debe Ingresar una operacion de Planilla de Contratados.", vbInformation, "Aviso"
        Me.txtOpeCont.SetFocus
        Exit Sub
    End If
    
    If Not lbEditado Then
        lsCod = Right(Me.TxtTipoPlanilla.Text, 1)
        oPla.AgregaConceptoPlanilla lsCod, Me.txtPlaDes.Text, GetMovNro(gsCodUser, gsCodAge), Format(CDate(mskFecRef & "/" & Format(gdFecSis, "mm/yyyy")), gsFormatoFecha), Me.txtPlaDes.Text, Format(CDate(Me.mskIni.Text & "/" & Format(gdFecSis, "mm/yyyy")), gsFormatoFecha), Format(CDate(Me.mskFin.Text & "/" & Format(gdFecSis, "mm/yyyy")), gsFormatoFecha), CDbl(Me.txtParametro.Text), "", lsOpc, Me.txtOpeEst.Text, Me.txtOpeCont.Text, Me.txtOpeAbono.Text
        
        For i = 1 To lvwG.ListItems.Count
            If lvwG.ListItems(i).Checked Then oPla.AgregaPlanillaEstadosRH lsCod, Trim(Right(lvwG.ListItems(i), 3))
        Next i
        
        Set rsPlaT = oPlaT.GetPlanillas
        Me.TxtPlanilla.rs = rsPlaT
        
        TxtPlanilla.Text = ""
        Me.lblPlanillaRes.Caption = TxtPlanilla.psDescripcion
    Else
        Dim sactualiza As String
        sactualiza = GetMovNro(gsCodUser, gsCodAge)
        oPla.ModificaConceptoPlanilla Me.TxtPlanilla.Text, Me.txtPlaDes.Text, sactualiza, Format(Format(gdFecSis, "yyyy/mm") & "/" & mskFecRef, "yyyy/mm/dd"), Me.txtPlaDes.Text, Format(Me.mskIni.Text & "/" & Format(gdFecSis, "mm/yyyy"), gsFormatoFecha), Format(Me.mskFin.Text & "/" & Format(gdFecSis, "mm/yyyy"), gsFormatoFecha), CDbl(Me.txtParametro.Text), "", lsOpc, Me.txtOpeEst.Text, Me.txtOpeCont.Text, Me.txtOpeAbono.Text
        
        oPla.EliminaPlanillaEstadosRH Me.TxtPlanilla.Text
        
        For i = 1 To lvwG.ListItems.Count
            If lvwG.ListItems(i).Checked Then oPla.AgregaPlanillaEstadosRH Me.TxtPlanilla.Text, Trim(Right(lvwG.ListItems(i), 3))
        Next i
    End If
    TxtPlanilla.Enabled = True
    
    lbEditado = False
    Activa False
    fraG.Enabled = False
End Sub

Private Sub cmdImprimir_Click()
    If Me.TxtPlanilla.Text = "" Then Exit Sub
    
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    Dim oPla As NActualizaDatosConPlanilla
    Set oPla = New NActualizaDatosConPlanilla
    Dim lsCadena As String
    
    
    lsCadena = oPla.GetReporte(Me.TxtPlanilla.Text & " " & Me.lblPlanillaRes.Caption, Me.cmbTipConcepto.Text, Me.Caption, gsNomAge, gsEmpresa, gdFecSis)
        
    oPrevio.Show lsCadena, Caption, True, 66
    
    Set oPrevio = Nothing
    Set oPla = Nothing
End Sub

Private Sub cmdIzqTodos_Click()
    Dim i As Integer
    Dim lnTope As Integer
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla

    lnTope = FlexAdd.Rows - 1
    
    For i = 1 To lnTope
        FlexAdd.Row = 1
        
        If FlexAdd.TextMatrix(1, 0) <> "" Then
            oPla.EliminaPlanillaConcepto Me.TxtPlanilla.Text, FlexAdd.TextMatrix(1, 0)
            
            If FlexLibres.TextMatrix(FlexLibres.Rows - 1, 0) <> "" Then FlexLibres.Rows = FlexLibres.Rows + 1
            FlexLibres.RowHeight(FlexLibres.Rows - 1) = 240
            FlexLibres.TextMatrix(FlexLibres.Rows - 1, 0) = FlexAdd.TextMatrix(1, 0)
            FlexLibres.TextMatrix(FlexLibres.Rows - 1, 1) = FlexAdd.TextMatrix(1, 1)
            
            If FlexAdd.Rows <> 2 Then
                FlexAdd.RemoveItem 1
            Else
                FlexAdd.TextMatrix(1, 0) = ""
                FlexAdd.TextMatrix(1, 1) = ""
            End If
        End If
    Next i
    FlexAdd.Refresh
End Sub

Private Sub cmdIzquno_Click()
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    
    If FlexAdd.TextMatrix(FlexAdd.Row, 0) = "" Then Exit Sub
    oPla.EliminaPlanillaConcepto Me.TxtPlanilla.Text, FlexAdd.TextMatrix(FlexAdd.Row, 0)
    
    If FlexLibres.TextMatrix(FlexLibres.Rows - 1, 0) <> "" Then FlexLibres.Rows = FlexLibres.Rows + 1
    FlexLibres.RowHeight(FlexLibres.Rows - 1) = 240
    FlexLibres.TextMatrix(FlexLibres.Rows - 1, 0) = FlexAdd.TextMatrix(FlexAdd.Row, 0)
    FlexLibres.TextMatrix(FlexLibres.Rows - 1, 1) = FlexAdd.TextMatrix(FlexAdd.Row, 1)
    
    If FlexAdd.Rows <> 2 Then
        FlexAdd.RemoveItem FlexAdd.Row
    Else
        FlexAdd.TextMatrix(FlexAdd.Row, 0) = ""
        FlexAdd.TextMatrix(FlexAdd.Row, 1) = ""
    End If
    FlexAdd.Refresh
End Sub

Private Sub cmdNuevo_Click()
    TxtPlanilla.Enabled = False
    TxtPlanilla.Text = ""
    lbEditado = False
    Activa True
    ClearScreen
    Me.TxtTipoPlanilla.SetFocus
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub FlexAdd_DblClick()
    cmdIzquno_Click
End Sub

Private Sub FlexLibres_DblClick()
    cmdDerUno_Click
End Sub

Private Sub Form_Load()
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Dim rsPla As ADODB.Recordset
    Set rsPla = New ADODB.Recordset
    Dim rsPlaT As ADODB.Recordset
    Set rsPlaT = New ADODB.Recordset
    Dim llAux As ListItem
    
    Set rsPlaT = oPla.GetPlanillas
    Me.TxtPlanilla.rs = rsPlaT
   
    Set rsP = oCon.GetConstante(6011)
    CargaCombo rsP, Me.cmbTipConcepto
    
    rsP.Close
    Set rsPla = oCon.GetRHTipoContrato
    TxtTipoPlanilla.rs = rsPla
    
    FlexLibres.Rows = 1
    FlexLibres.Cols = 2
    FlexLibres.Rows = 2
    FlexLibres.FixedCols = 0
    FlexLibres.FixedRows = 1
    
    FlexLibres.TextMatrix(0, 0) = "Nemonico"
    FlexLibres.TextMatrix(0, 1) = "Nombre"
    
    FlexLibres.ColWidth(0) = 500
    FlexLibres.ColWidth(1) = 3900
    
    '******************Con Grupo
    FlexAdd.Rows = 1
    FlexAdd.Cols = 2
    FlexAdd.Rows = 2
    FlexAdd.FixedCols = 0
    FlexAdd.FixedRows = 1
    
    FlexAdd.TextMatrix(0, 0) = "Nemonico"
    FlexAdd.TextMatrix(0, 1) = "Nombre"
    
    FlexAdd.ColWidth(0) = 1
    FlexAdd.ColWidth(1) = 3900
    
    Me.fraDatos.Enabled = False
    
    Set rsP = oPla.GetPlanillaEstadosUsados(Me.TxtPlanilla.Text)

    lvwG.HideColumnHeaders = False
    lvwG.ColumnHeaders.Clear
    lvwG.ColumnHeaders.Add , , "Tipo de Persona", 3500
    lvwG.ListItems.Clear
    If Not RSVacio(rsP) Then
        While Not rsP.EOF
            Set llAux = lvwG.ListItems.Add(, , Trim(rsP!Descrip & Space(30) & rsP!Valor))
            If Not IsNull(rsP!Estado) Then
                llAux.Checked = True
            Else
                llAux.Checked = False
            End If
            rsP.MoveNext
        Wend
    End If

    rsP.Close
    Set rsP = Nothing
    
    UbicaCombo Me.cmbTipConcepto, "5"
    Activa False
End Sub

Private Sub lvwG_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.cmdGrabar.Enabled Then
            cmdGrabar.SetFocus
        End If
    End If
End Sub

Private Sub mskIni_GotFocus()
    mskIni.SelStart = 0
    mskIni.SelLength = 5
End Sub

Private Sub mskIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFin.SetFocus
    End If
End Sub

Private Sub mskFin_GotFocus()
    mskFin.SelStart = 0
    mskFin.SelLength = 10
End Sub

Private Sub mskFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtParametro.SetFocus
    End If
End Sub

Private Sub mskFecRef_GotFocus()
    mskFecRef.SelStart = 0
    mskFecRef.SelLength = 10
End Sub

Private Sub mskFecRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskIni.SetFocus
    End If
End Sub

Private Sub txtCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.chkActivo.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtOpeAbono_GotFocus()
    txtOpeAbono.SelStart = 0
    txtOpeAbono.SelLength = 50
End Sub

Private Sub txtOpeAbono_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtOpeEst.SetFocus
    End If
End Sub

Private Sub txtOpeCOnt_GotFocus()
    txtOpeCont.SelStart = 0
    txtOpeCont.SelLength = 50
End Sub

Private Sub txtOpeCOnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskFecRef.SetFocus
    End If
End Sub

Private Sub txtOpeEst_GotFocus()
    txtOpeEst.SelStart = 0
    txtOpeEst.SelLength = 50
End Sub

Private Sub txtOpeEst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtOpeCont.SetFocus
    End If
End Sub

Private Sub txtParametro_GotFocus()
    txtParametro.SelStart = 0
    txtParametro.SelLength = 50
End Sub

Private Sub txtParametro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Me.txtCta.SetFocus
        Me.chkActivo.SetFocus
    Else
        NumerosDecimales txtParametro, KeyAscii
    End If
End Sub

Private Sub txtPlaDes_GotFocus()
    txtPlaDes.SelStart = 0
    txtPlaDes.SelLength = 300
End Sub

Private Sub txtPlaDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtOpeAbono.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub ClearScreen(Optional pbTodos As Boolean)
    Dim i As Integer
    
    Me.txtPlaDes.Text = ""
    'Me.txtCta.Text = ""
    Me.mskIni.Text = "__"
    Me.mskFin.Text = "__"
    Me.mskFecRef.Text = "__"
    Me.chkActivo.value = 1
    Me.TxtTipoPlanilla.Text = ""
    Me.txtParametro.Text = ""
    
    For i = 1 To Me.lvwG.ListItems.Count
        lvwG.ListItems(i).Checked = False
    Next i
    
    Me.fraDatos.Enabled = False
    
    If lnTipo = gTipoOpeRegistro Then
        Me.cmbTipConcepto.ListIndex = -1
    End If
    
    Me.FlexLibres.Rows = 1
    Me.FlexLibres.Rows = 2
    Me.FlexLibres.FixedRows = 1
    Me.FlexAdd.Rows = 1
    Me.FlexAdd.Rows = 2
    Me.FlexAdd.FixedRows = 1
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.cmdNuevo.Visible = Not pbValor
    Me.cmdEditar.Visible = Not pbValor
    Me.cmdGrabar.Visible = pbValor
    Me.cmdCancelar.Visible = pbValor
    Me.cmdSalir.Enabled = Not pbValor
    Me.fraDatos.Enabled = pbValor
    Me.cmbTipConcepto.Enabled = Not pbValor
    Me.fraEstadoValidos.Enabled = pbValor
    Me.fraPlanilla.Enabled = pbValor
    Me.fraG.Enabled = Not pbValor
    If lnTipo = gTipoOpeRegistro Then
        cmdEliminar.Enabled = False
        cmdEditar.Enabled = False
        cmdImprimir.Enabled = False
        Me.TxtPlanilla.Enabled = False
        Me.cmbTipConcepto.Enabled = False
    ElseIf lnTipo = gTipoOpeMantenimiento Then
        Me.cmdNuevo.Enabled = False
        Me.cmdEliminar.Enabled = Not pbValor
        Me.TxtTipoPlanilla.Enabled = False
    ElseIf lnTipo = gTipoOpeConsulta Then
        cmdEliminar.Enabled = False
        cmdEditar.Enabled = False
        cmdImprimir.Enabled = False
        cmdNuevo.Enabled = False
        Me.fraEstadoValidos.Enabled = True
    ElseIf lnTipo = gTipoOpeReporte Then
        cmdEliminar.Enabled = False
        cmdEditar.Enabled = False
        cmdNuevo.Enabled = False
        Me.fraEstadoValidos.Enabled = True
    End If
End Sub

Public Sub Ini(pnTipo As TipoOpe, psCaption As String)
    lnTipo = pnTipo
    Caption = psCaption
    Me.Show 1
End Sub

Private Sub TxtPlanilla_EmiteDatos()
    Me.lblPlanillaRes.Caption = TxtPlanilla.psDescripcion
    cmbPlaCod_Change
    Me.txtPlaDes.Text = TxtPlanilla.psDescripcion
End Sub

Private Sub TxtTipoPlanilla_EmiteDatos()
    Me.lblTipoPlanillaRes.Caption = TxtTipoPlanilla.psDescripcion
    Me.txtPlaDes.SetFocus
End Sub
