VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmColocEvalConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Calificacion"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "FrmColocEvalConsulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
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
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   7335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshCalif 
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      FillStyle       =   1
      GridLinesUnpopulated=   1
      SelectionMode   =   1
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).BandIndent=   5
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   300
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
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
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton OptBusca 
         Caption         =   "Por Cliente"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton OptBusca 
         Caption         =   "Todos"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   4920
         TabIndex        =   4
         Top             =   240
         Width           =   450
      End
   End
End
Attribute VB_Name = "FrmColocEvalConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CargaDatos(sCod As String, nOpt As Integer)
Dim rs As New ADODB.Recordset
Dim oEval As nColocEvalCal
Set oEval = New nColocEvalCal
Dim Server As nRcdReportes
Set Server = New nRcdReportes

Set rs = oEval.nobtineEvalconsulta(sCod, nOpt, mskFecha, Server.GetServerConsol)
MshCalif.Clear
Marco
While Not rs.EOF
    With MshCalif
        .TextMatrix(.Rows - 1, 0) = rs!cPersCod
        .TextMatrix(.Rows - 1, 2) = PstaNombre(rs!nombre)
        .TextMatrix(.Rows - 1, 3) = rs!cCalGen
        .TextMatrix(.Rows - 1, 4) = rs!cCtaCod
        .TextMatrix(.Rows - 1, 5) = Format(IIf(IsNull(rs!nSaldoCap), 0, rs!nSaldoCap), "0.00")
        .TextMatrix(.Rows - 1, 6) = rs!nDiasAtraso
        .Rows = .Rows + 1
    End With
    rs.MoveNext
Wend
If MshCalif.Rows > 2 Then MshCalif.Rows = MshCalif.Rows - 1
Set rs = Nothing
Set oEval = Nothing
Set Server = Nothing
End Sub
Private Sub cmdBuscar_Click()
Dim loPers As UPersona
Dim lsPersCod As String
Set loPers = New UPersona
    
    If Me.OptBusca(1).Value Then Call CargaDatos("", 1): Exit Sub
    If Me.OptBusca(0).Value Then
        Set loPers = frmBuscaPersona.Inicio
        If loPers Is Nothing Then Exit Sub
        lsPersCod = loPers.sPersCod
        Call CargaDatos(lsPersCod, 0)
    End If
        
Set loPers = Nothing
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Me.Icon = LoadPicture(App.path & gsRutaIcono)
mskFecha = gdFecSis
Marco
End Sub

Private Sub Marco()
With MshCalif
    MshCalif.Rows = 2
    .TextMatrix(0, 0) = "Cod. Persona"
    .TextMatrix(0, 1) = "Nombre"
    .TextMatrix(0, 2) = "Cal. Gen."
    .TextMatrix(0, 3) = "Nro. Credito"
    .TextMatrix(0, 4) = "Saldo Cap."
    .TextMatrix(0, 5) = "Dias Atraso"
    .ColWidth(0) = 1000
    .ColWidth(1) = 4500
    .ColWidth(2) = 700
    .ColWidth(3) = 1500
    .ColWidth(4) = 1800
    .ColWidth(5) = 1000
End With
End Sub

Private Sub MshCalif_Click()
'If MshCalif.Rows > 1 And MshCalif.TextMatrix(MshCalif.Row, MshCalif.Col) <> "" Then
' FrmColocEvalConsultaCailf.Inicio (MshCalif.TextMatrix(MshCalif.Row, 0))
'End If

End Sub

Private Sub OptBusca_Click(Index As Integer)

MshCalif.Clear
'MshCalif.Rows = 2
Marco
End Sub
