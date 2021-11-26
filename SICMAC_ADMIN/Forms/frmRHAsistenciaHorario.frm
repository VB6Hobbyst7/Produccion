VERSION 5.00
Begin VB.Form frmRHAsistenciaHorario 
   Caption         =   "Mantenimeinto de Horarios"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   Icon            =   "frmRHAsistenciaHorario.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   8040
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3885
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   2670
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgregarDia 
      Caption         =   "&Agregar Dia"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminarDia 
      Caption         =   "&Eliminar Dia"
      Height          =   375
      Left            =   1455
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ComboBox cmbTipoHorario 
      Height          =   315
      ItemData        =   "frmRHAsistenciaHorario.frx":030A
      Left            =   240
      List            =   "frmRHAsistenciaHorario.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin Sicmact.FlexEdit FlexDia 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   6600
      _ExtentX        =   9102
      _ExtentY        =   5741
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "#-Cod Dia-Dia-Turno-Hora Ini-Hora Fin-Existe"
      EncabezadosAnchos=   "300-800-800-1200-800-800-0"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-X-3-4-5-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-1-0-3-2-2-0"
      EncabezadosAlineacion=   "C-L-L-L-R-R-C"
      FormatosEdit    =   "0-0-0-6-6-6-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label lblHoraSemana 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   6000
      TabIndex        =   5
      Top             =   3360
      Width           =   945
   End
   Begin VB.Label lblHora 
      Caption         =   "Horas S."
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   3360
      Width           =   705
   End
End
Attribute VB_Name = "frmRHAsistenciaHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnRowAnt As Integer
Dim lnTipo As TipoOpe
Dim lsCaption As String

Private Sub cmdAgregarDia_Click()
 'If Not IsDate(Me.FlexHor.TextMatrix(FlexHor.Row, 1)) Then Exit Sub
    Me.FlexDia.AdicionaFila
    Me.FlexDia.TextMatrix(FlexDia.Row, 6) = ""
    FlexDia.SetFocus
End Sub

Private Sub cmdEliminarDia_Click()
    'If FlexDia.TextMatrix(FlexDia.Row, 6) = "" Then
    '    MsgBox "No puede borrar un registro grabado.", vbInformation, "Aviso"
    '    Exit Sub
    'End If
    
    FlexDia.EliminaFila FlexDia.Row
    
    Me.lblHoraSemana.Caption = Format(GetSuma(), "0#.00")

End Sub


Private Sub cmdGrabar_Click()
'Dim cRHHorarioTurno As Integer
'Dim cRHHorarioDias As Integer
'Dim nGrupoHorario As Integer
'Dim dRHHorarioInicio As String
'Dim dRHHorarioFin As String
''Dim i As Integer
'Dim n As Integer
'For i = 1 To FlexDia.Rows - 1
'    For n = 1 To FlexDia.Cols - 1
'        cRHHorarioTurno = FlexDia.TextMatrix
'        cRHHorarioDias = 2
'        nGrupoHorario = 3
'        dRHHorarioInicio = 4
 '       dRHHorarioFin = 5
'    Next
'Next

End Sub

Private Sub FlexDia_RowColChange()
If Not IsNumeric(Me.FlexDia.TextMatrix(FlexDia.Row, 1)) Then
        Me.FlexDia.TextMatrix(FlexDia.Row, 1) = ""
    End If
    
    Me.lblHoraSemana.Caption = Format(GetSuma(), "0#.00")
End Sub

Private Sub Form_Load()
    Dim oCon As DConstantes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConstantes
    
    Set rs = oCon.GetConstante(1022, False, False, True)
    Me.FlexDia.rsTextBuscar = rs
    Me.FlexDia.CargaCombo oCon.GetConstante(gRHEmpleadoTurno)
    
    CargaCombo oCon.GetConstante(6041), Me.cmbTipoHorario
    cmbTipoHorario.ListIndex = 0
    'Activa False
    'Limpia
    Caption = lsCaption

End Sub
Private Function GetSuma() As Double
    Dim lnI As Integer
    Dim lnSuma As Double
    
    lnSuma = 0
    For lnI = 1 To Me.FlexDia.Rows - 1
        If IsDate(Me.FlexDia.TextMatrix(lnI, 4)) And IsDate(Me.FlexDia.TextMatrix(lnI, 5)) Then
            lnSuma = lnSuma + DateDiff("n", CDate(Me.FlexDia.TextMatrix(lnI, 4)), CDate(Me.FlexDia.TextMatrix(lnI, 5)))
        End If
    Next lnI
    
    GetSuma = lnSuma / 60
End Function
