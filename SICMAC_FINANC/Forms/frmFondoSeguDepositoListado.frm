VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmFondoSeguDepositoListado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
   Icon            =   "frmFondoSeguDepositoListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Búsqueda"
      Height          =   1335
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   8190
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   315
         Left            =   480
         TabIndex        =   8
         Top             =   840
         Width           =   2340
      End
      Begin VB.CommandButton cmdbuscar 
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
         Height          =   360
         Left            =   5760
         TabIndex        =   7
         ToolTipText     =   "Buscar Credito"
         Top             =   840
         Width           =   1200
      End
      Begin VB.OptionButton optFechas 
         Caption         =   "Por Fecha de Registro"
         Height          =   315
         Left            =   450
         TabIndex        =   1
         Top             =   330
         Width           =   2340
      End
      Begin MSMask.MaskEdBox mskInicio 
         Height          =   345
         Left            =   3510
         TabIndex        =   3
         Top             =   270
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFin 
         Height          =   345
         Left            =   5610
         TabIndex        =   4
         Top             =   270
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Del:"
         Height          =   240
         Left            =   3060
         TabIndex        =   6
         Top             =   345
         Width           =   315
      End
      Begin VB.Label Label6 
         Caption         =   "Al:"
         Height          =   240
         Left            =   5235
         TabIndex        =   5
         Top             =   345
         Width           =   240
      End
   End
   Begin Sicmact.FlexEdit FePolizas 
      Height          =   3270
      Left            =   105
      TabIndex        =   2
      Top             =   1425
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   4921
      Cols0           =   8
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Año-Mes-Nº Trim.-Tasa %-Tipo Cambio-Tot. Dep. MN-Tot. Dep. ME"
      EncabezadosAnchos=   "400-800-800-800-800-1000-1600-1600"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-R-R-R-R"
      FormatosEdit    =   "0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmFondoSeguDepositoListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum TipoOperacion
    nBusqueda = 0
    nConsulta = 1
End Enum

Dim nTipoOperacion As TipoOperacion

'Public sNumPoliza As String
'Public nTipoPoliza As Integer
'Public sPersCodContr As String

Public nAnio As Integer
Public nMes As Integer
'Dim nEstadoPoliza As Integer

Public Sub Inicio(ByVal pnTipoBusqueda As TipoOperacion)
    nTipoOperacion = pnTipoBusqueda
'    nEstadoPoliza = pnEstadoPoliza
'    sNumPoliza = ""
    Me.Show 1
    
End Sub

'Private Sub cmdBuscaCont_Click()
'
'Dim oPers As UPersona
'    Set oPers = frmBuscaPersona.Inicio
'    If Not oPers Is Nothing Then
'        LblContPersCod.Caption = oPers.sPersCod
'        LblContPersNombre.Caption = oPers.sPersNombre
'    End If
'    Set oPers = Nothing
'    cmdbuscar.SetFocus
'End Sub

Private Sub cmdBuscar_Click()

Dim oPol As DOperacion

Dim rs As ADODB.Recordset
Set oPol = New DOperacion

If Me.optFechas.value Then
    If Len(ValidaFecha(mskInicio.Text)) = 0 And Len(ValidaFecha(mskFin.Text)) = 0 Then
        Set rs = oPol.CargaFondoSeguDepo(1, Format(CDate(mskInicio.Text), "yyyymmdd"), Format(CDate(mskFin.Text), "yyyymmdd"))
    Else
        MsgBox "Ingrese fecha correctas.", vbOKOnly, "Atención"
        Exit Sub
    End If
ElseIf Me.optTodos.value Then
    Set rs = oPol.CargaFondoSeguDepo(2)
End If

If rs.EOF Then MsgBox "No se encontraron registros", vbInformation, "Mensaje"

FePolizas.Clear
FePolizas.FormaCabecera
FePolizas.Rows = 2
FePolizas.rsFlex = rs
FePolizas.SetFocus
Set oPol = Nothing
End Sub


Private Sub FePolizas_DblClick()
If nTipoOperacion = 2 Or nTipoOperacion = 0 Then
    nAnio = FePolizas.TextMatrix(FePolizas.Row, 1)
    nMes = FePolizas.TextMatrix(FePolizas.Row, 2)
    
'    sNumPoliza = FePolizas.TextMatrix(FePolizas.Row, 1)
'    nTipoPoliza = FePolizas.TextMatrix(FePolizas.Row, 6)
'    sPersCodContr = FePolizas.TextMatrix(FePolizas.Row, 5)
End If

Unload Me

End Sub

Private Sub FePolizas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call FePolizas_DblClick
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
End Sub

Private Sub LblContPersCod_Click()

End Sub

Private Sub mskFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdbuscar.SetFocus
End Sub

Private Sub mskInicio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then mskFin.SetFocus
End Sub

'Private Sub optContrat_Click()
'If optContrat.value Then
'    CmdBuscaCont.Enabled = True
'    mskInicio.Enabled = False
'    mskInicio.Text = "__/__/____"
'    mskFin.Enabled = False
'    mskFin.Text = "__/__/____"
'End If
'End Sub

Private Sub optFechas_Click()
If optFechas.value Then
    
'    CmdBuscaCont.Enabled = False
    mskInicio.Enabled = True
    mskInicio.Text = "__/__/____"
    mskFin.Enabled = True
    mskFin.Text = "__/__/____"
    
    'Me.LblContPersCod.Caption = ""
    'Me.LblContPersNombre.Caption = ""
    
End If
End Sub

