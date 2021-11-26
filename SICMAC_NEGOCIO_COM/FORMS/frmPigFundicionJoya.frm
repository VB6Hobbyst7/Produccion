VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPigFundicionJoya 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proceso de Fundición"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   2745
      TabIndex        =   15
      Top             =   6975
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   4185
      TabIndex        =   14
      Top             =   6975
      Width           =   1200
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Procesar"
      Height          =   345
      Left            =   1305
      TabIndex        =   13
      Top             =   6975
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   6690
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6075
      Begin VB.CommandButton cmdEjecutar 
         Caption         =   "&Ejecutar"
         Height          =   360
         Left            =   4080
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.Frame frmCriterios 
         Caption         =   "Busquedad Por Remate"
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
         Height          =   840
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3735
         Begin MSDataListLib.DataCombo cboRemDesde 
            Height          =   315
            Left            =   1305
            TabIndex        =   4
            Top             =   285
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cboRemHasta 
            Height          =   315
            Left            =   2550
            TabIndex        =   5
            Top             =   270
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl1 
            Caption         =   "Remate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   330
            Width           =   975
         End
         Begin VB.Label lblText1 
            Alignment       =   2  'Center
            Caption         =   "AL"
            Height          =   210
            Left            =   2190
            TabIndex        =   7
            Top             =   330
            Width           =   405
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Joyas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   5340
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   5835
         Begin SICMACT.FlexEdit feDisponibles 
            Height          =   3255
            Left            =   120
            TabIndex        =   6
            Top             =   570
            Width           =   5475
            _ExtentX        =   9657
            _ExtentY        =   5741
            Cols0           =   7
            ScrollBars      =   1
            HighLight       =   2
            EncabezadosNombres=   "Item-Remate-nPiezas-PBruto-PNeto-nTasacion-nTasacionAdicional"
            EncabezadosAnchos=   "0-800-700-900-900-1350-1350"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "R-R-L-L-R-R"
            FormatosEdit    =   "3-3-0-0-0-2-3-2-2-2-0-1"
            AvanceCeldas    =   1
            TextArray0      =   "Item"
            SelectionMode   =   1
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin VB.Label Label6 
            Caption         =   "Tot. Piezas :"
            Height          =   210
            Left            =   120
            TabIndex        =   22
            Top             =   4020
            Width           =   1275
         End
         Begin VB.Label lblTotPiezas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   345
            Left            =   1440
            TabIndex        =   21
            Top             =   3960
            Width           =   1140
         End
         Begin VB.Label Label7 
            Caption         =   "Tot. Tasacion :"
            Height          =   210
            Left            =   120
            TabIndex        =   20
            Top             =   4965
            Width           =   1155
         End
         Begin VB.Label lblTotTasacion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   345
            Left            =   1440
            TabIndex        =   19
            Top             =   4920
            Width           =   1140
         End
         Begin VB.Label Label4 
            Caption         =   "Tot. Tasacion Adic.:"
            Height          =   210
            Left            =   2640
            TabIndex        =   18
            Top             =   4965
            Width           =   1545
         End
         Begin VB.Label lblTotTasacionAdicional 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   345
            Left            =   4200
            TabIndex        =   17
            Top             =   4920
            Width           =   1305
         End
         Begin VB.Label lblTotPNeto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   345
            Left            =   4200
            TabIndex        =   12
            Top             =   4410
            Width           =   1305
         End
         Begin VB.Label Label5 
            Caption         =   "Tot. Peso Neto :"
            Height          =   210
            Left            =   2640
            TabIndex        =   11
            Top             =   4470
            Width           =   1200
         End
         Begin VB.Label lblTotPBruto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   345
            Left            =   1440
            TabIndex        =   10
            Top             =   4410
            Width           =   1140
         End
         Begin VB.Label Label3 
            Caption         =   "Tot. Peso Bruto :"
            Height          =   210
            Left            =   120
            TabIndex        =   9
            Top             =   4470
            Width           =   1275
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Disponibles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   4845
         End
      End
   End
End
Attribute VB_Name = "frmPigFundicionJoya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'* frmPigSeleccionVentaFundicion : Seleccion de Joyas Para Venta ó Fundición
'* EAFA - 22/10/2002
'**************************************************************
Dim fnVarCriterio As Integer
Dim psCondicionAdic As String
Dim fmJoyas() As String
Dim fnTipoProceso As Integer

Private Sub cmdCancelar_Click()
limpiaInicia
End Sub

Private Sub cmdEjecutar_Click()
Dim lrDatos As ADODB.Recordset
Dim lrPigContrato As DPigContrato
    
If Not (Not IsNull(cboRemDesde.Text) And Not IsNull(cboRemHasta.Text)) Then
    MsgBox "Debe Especificar correctamente los Números de Remate", vbInformation, "Aviso"
    Exit Sub
End If

Set lrDatos = New ADODB.Recordset
Set lrPigContrato = New DPigContrato
    Set lrDatos = lrPigContrato.dObtieneDatosFundicion("T", cboRemDesde.Text, cboRemHasta.Text)
Set lrPigContrato = Nothing

If lrDatos Is Nothing Or (lrDatos.BOF And lrDatos.EOF) Then
    MsgBox "No se Encuentro Información", vbInformation, "Aviso"
    Exit Sub
Else
    feDisponibles.Clear
    feDisponibles.Rows = 2
    feDisponibles.FormaCabecera
    Do While Not lrDatos.EOF
          feDisponibles.AdicionaFila
          If Not IsNull(lrDatos!Remate) Then _
          feDisponibles.TextMatrix(feDisponibles.Row, 1) = lrDatos!Remate
          If Not IsNull(lrDatos!Piezas) Then _
          feDisponibles.TextMatrix(feDisponibles.Row, 2) = lrDatos!Piezas
          If Not IsNull(lrDatos!PesoBruto) Then _
          feDisponibles.TextMatrix(feDisponibles.Row, 3) = lrDatos!PesoBruto
          If Not IsNull(lrDatos!pesoneto) Then _
          feDisponibles.TextMatrix(feDisponibles.Row, 4) = lrDatos!pesoneto
          If Not IsNull(lrDatos!Tasacion) Then _
          feDisponibles.TextMatrix(feDisponibles.Row, 5) = lrDatos!Tasacion
          If Not IsNull(lrDatos!TasacionAdicional) Then _
          feDisponibles.TextMatrix(feDisponibles.Row, 6) = lrDatos!TasacionAdicional
          lrDatos.MoveNext
    Loop
    ActivaDesactivaControles
    SumaColumnas
    frmCriterios.Enabled = False
    cmdEjecutar.Enabled = False
End If
End Sub

Private Sub cmdGrabar_Click()
Dim lrDatos As ADODB.Recordset
Dim lrPigContrato As DPigContrato
Dim loRegPig As DPigActualizaBD
Dim mbTrans As Boolean

If MsgBox("Desea Realizar el Proceso de Fundición? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
    'Genera el Mov Nro
    Set loContFunct = New NContFunciones
         lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
Set lrDatos = New ADODB.Recordset
Set lrPigContrato = New DPigContrato
    Set lrDatos = lrPigContrato.dObtieneDatosFundicion("A", cboRemDesde.Text, cboRemHasta.Text)
Set lrPigContrato = Nothing

If lrDatos Is Nothing Or (lrDatos.BOF And lrDatos.EOF) Then
    MsgBox "No se Encuentro Información", vbInformation, "Aviso"
    Exit Sub
Else
    Do While Not lrDatos.EOF
       Set loRegPig = New DPigActualizaBD
       loRegPig.dBeginTrans
       mbTrans = True
       Call loRegPig.dUpdateColocPigProceso(lrDatos!Remate, lrDatos!nEstadoJoya, lrDatos!cCtaCod, lrDatos!nItemPieza, lsMovNro, , , gPigSituacionFundido, , , , False)
       loRegPig.dCommitTrans
       mbTrans = False
       Set loRegPig = Nothing
       lrDatos.MoveNext
    Loop
    ActivaDesactivaControles
    SumaColumnas
    frmCriterios.Enabled = False
    cmdEjecutar.Enabled = False
End If
    limpiaInicia
End If
End Sub


Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSelTodos_Click()
Do While feDisponibles.SumaRow(5) > 0
    feSeleccionadas.AdicionaFila
    For c = 1 To feDisponibles.Cols - 1
        feSeleccionadas.TextMatrix(feSeleccionadas.Row, c) = feDisponibles.TextMatrix(feDisponibles.Row, c)
    Next
    feSeleccionadas.TextMatrix(feSeleccionadas.Row, 5) = Format(feSeleccionadas.TextMatrix(feSeleccionadas.Row, 5), "##,##0.00")
    feDisponibles.EliminaFila (feDisponibles.Row)
Loop
ActivaDesactivaControles
SumaColumnas
End Sub



Private Sub Form_Activate()
limpiaInicia
ActivaDesactivaControles
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub CargaCombos()
CargaRemates

End Sub

Private Sub CargaRemates()
Dim lrDatos As ADODB.Recordset
Dim lrPigFunciones As DPigFunciones

Set lrDatos = New ADODB.Recordset
Set lrPigFunciones = New DPigFunciones
      Set lrDatos = lrPigFunciones.GetListaRemates()
Set lrPigFunciones = Nothing

Set cboRemDesde.RowSource = lrDatos
       cboRemDesde.ListField = "nRemate"
       cboRemDesde.BoundColumn = "nRemate"

Set cboRemHasta.RowSource = lrDatos
       cboRemHasta.ListField = "nRemate"
       cboRemHasta.BoundColumn = "nRemate"

Set lrDatos = Nothing
End Sub


Private Sub ActivaDesactivaControles()
Dim X As Integer
If feDisponibles.SumaRow(4) = 0 Then
    cmdGrabar.Enabled = False
Else
    feDisponibles.SetFocus
    cmdGrabar.Enabled = True
End If
End Sub

Private Sub SumaColumnas()
Dim i As Integer

lblTotPiezas.Caption = Format(feDisponibles.SumaRow(2), "###,##0")
lblTotPBruto.Caption = Format(feDisponibles.SumaRow(3), "###,##0.00 ")
lblTotPNeto.Caption = Format(feDisponibles.SumaRow(4), "###,##0.00 ")
lblTotTasacion.Caption = Format(feDisponibles.SumaRow(5), "###,##0.00 ")
lblTotTasacionAdicional.Caption = Format(feDisponibles.SumaRow(6), "###,##0.00 ")

End Sub



Private Sub limpiaInicia()

lblTotPiezas.Caption = Format("0.00", "###,##0 ")
lblTotPNeto.Caption = Format("0.00", "###,##0.00 ")
lblTotPBruto.Caption = Format("0.00", "###,##0.00 ")
lblTotTasacion.Caption = Format("0.00", "###,##0.00 ")
lblTotTasacionAdicional.Caption = Format("0.00", "###,##0.00 ")

feDisponibles.Clear
feDisponibles.Rows = 2
feDisponibles.FormaCabecera
CargaCombos
ActivaDesactivaControles
frmCriterios.Enabled = True
cmdEjecutar.Enabled = True
End Sub

Private Sub txtCtaCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdEjecutar_Click
End If
End Sub

