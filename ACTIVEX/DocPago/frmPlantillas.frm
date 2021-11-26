VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmPlantillas 
   Caption         =   "Edición de Plantilla de Cartas"
   ClientHeight    =   6690
   ClientLeft      =   435
   ClientTop       =   1455
   ClientWidth     =   11100
   Icon            =   "frmPlantillas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   11100
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "&Seleccionar"
      Height          =   390
      Left            =   9570
      TabIndex        =   10
      Top             =   6112
      Width           =   1275
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Plantilla"
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
      Height          =   5910
      Left            =   2310
      TabIndex        =   1
      Top             =   90
      Width           =   8640
      Begin VB.ComboBox cboCampos 
         Height          =   315
         ItemData        =   "frmPlantillas.frx":030A
         Left            =   6270
         List            =   "frmPlantillas.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1980
      End
      Begin VB.TextBox txtPlantillaId 
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
         Height          =   315
         Left            =   1275
         MaxLength       =   10
         TabIndex        =   2
         Top             =   225
         Width           =   1470
      End
      Begin VB.TextBox txtReferencia 
         Height          =   315
         Left            =   1275
         MaxLength       =   150
         TabIndex        =   3
         Top             =   630
         Width           =   6975
      End
      Begin RichTextLib.RichTextBox rtfPlantilla 
         Height          =   4770
         Left            =   300
         TabIndex        =   4
         Top             =   975
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   8414
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmPlantillas.frx":030E
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " Campos"
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
         Left            =   5505
         TabIndex        =   17
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Referencia :"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   660
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   345
         TabIndex        =   15
         Top             =   285
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Márgenes de Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   75
      TabIndex        =   11
      Top             =   6000
      Width           =   4695
      Begin Spinner.uSpinner txtMgSup 
         Height          =   360
         Left            =   735
         TabIndex        =   20
         Top             =   210
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin Spinner.uSpinner txtMgIzq 
         Height          =   360
         Left            =   2310
         TabIndex        =   19
         Top             =   180
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin Spinner.uSpinner txtMgDer 
         Height          =   360
         Left            =   3825
         TabIndex        =   18
         Top             =   180
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Superior"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   255
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Izquierdo"
         Height          =   195
         Left            =   1590
         TabIndex        =   13
         Top             =   255
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Derecho"
         Height          =   195
         Left            =   3150
         TabIndex        =   12
         Top             =   255
         Width           =   615
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgPlantillas 
      Height          =   5745
      Left            =   135
      TabIndex        =   0
      Top             =   195
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   10134
      _Version        =   393216
      BackColorFixed  =   -2147483637
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   390
      Left            =   4935
      TabIndex        =   6
      Top             =   6112
      Width           =   1275
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   390
      Left            =   6210
      TabIndex        =   7
      Top             =   6112
      Width           =   1275
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   390
      Left            =   4935
      TabIndex        =   8
      Top             =   6112
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   6210
      TabIndex        =   9
      Top             =   6112
      Width           =   1275
   End
   Begin VB.Menu mnuCamposGen 
      Caption         =   "Campos"
      Visible         =   0   'False
      Begin VB.Menu mnuCampos 
         Caption         =   "Nro. Movimiento"
         Index           =   0
      End
      Begin VB.Menu mnuCampos 
         Caption         =   "Fecha"
         Index           =   1
      End
      Begin VB.Menu mnuCampos 
         Caption         =   "Entidad Origen"
         Index           =   2
      End
      Begin VB.Menu mnuCampos 
         Caption         =   "Entidad Destino"
         Index           =   3
      End
      Begin VB.Menu mnuCampos 
         Caption         =   "Cuenta Origen"
         Index           =   4
      End
      Begin VB.Menu mnuCampos 
         Caption         =   "Cuenta Destino"
         Index           =   5
      End
      Begin VB.Menu mnuCampos 
         Caption         =   "Importe"
         Index           =   6
      End
      Begin VB.Menu mnuCampos 
         Caption         =   "Nro.Documento"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmPlantillas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim nContFunc As NContFunciones
Dim nTipCarta As Integer
Public lsFile As String
Dim nLastPos As Integer
Dim sCampo As String
Dim lsOpeCod  As String
Dim lbNuevo As Boolean
Dim oNPlant As NPlantilla
Dim oPlant As dPlantilla

Public lsPlantillaID As String
Public lnMagDer As Integer
Public lnMagIzq As Integer
Public lnMagSup As Integer

Private Sub cboCampos_DblClick()
InsertaCampo

End Sub

Private Sub cboCampos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   InsertaCampo
End If
End Sub
Private Sub InsertaCampo(Optional lbMenu As Boolean = False, Optional Item As Integer = 0)
Dim sTxt As String
If lbMenu = False Then
    Select Case cboCampos.ListIndex
        Case 0: sCampo = "gcMovNro"
        Case 1: sCampo = "gdFecha"
        Case 2: sCampo = "gcEntiOrig"
        Case 3: sCampo = "gcEntiDest"
        Case 4: sCampo = "gcCtaEntiOrig"
        Case 5: sCampo = "gcCtaEntiDest"
        Case 6: sCampo = "gnImporte"
        Case 7: sCampo = "gcDocNro"
    End Select
Else
    Select Case Item
        Case 0: sCampo = "gcMovNro"
        Case 1: sCampo = "gdFecha"
        Case 2: sCampo = "gcEntiOrig"
        Case 3: sCampo = "gcEntiDest"
        Case 4: sCampo = "gcCtaEntiOrig"
        Case 5: sCampo = "gcCtaEntiDest"
        Case 6: sCampo = "gnImporte"
        Case 7: sCampo = "gcDocNro"
    End Select
End If
If nLastPos > 0 Then
    rtfPlantilla.Text = Mid(rtfPlantilla.Text, 1, nLastPos) & "<<" & sCampo & ">>" & Mid(rtfPlantilla.Text, nLastPos, Len(rtfPlantilla.Text))
End If

rtfPlantilla.SetFocus
rtfPlantilla.SelStart = nLastPos + Len("<<" & sCampo & ">>")
End Sub

Private Sub cmdCancelar_Click()
lbNuevo = False
fraDatos.Enabled = False
fgPlantillas.Enabled = True
cmdNuevo.Visible = True
cmdEditar.Visible = True
cmdCancelar.Visible = False
cmdGrabar.Visible = False
cmdSeleccionar.Enabled = True
Me.fgPlantillas.SetFocus
End Sub
Private Sub cmdEditar_Click()
lbNuevo = False
cmdNuevo.Visible = False
cmdEditar.Visible = False
cmdCancelar.Visible = True
cmdGrabar.Visible = True
cmdSeleccionar.Enabled = False
fraDatos.Enabled = True
fgPlantillas.Enabled = False
txtReferencia.SetFocus
End Sub

Private Sub cmdGrabar_Click()
If MsgBox("¿ Está seguro de grabar la Información ? ", vbYesNo + vbQuestion, "Mensaje de confirmación") = vbYes Then
    Grabar
    lbNuevo = False
    fraDatos.Enabled = False
    fgPlantillas.Enabled = True
    cmdNuevo.Visible = True
    cmdEditar.Visible = True
    cmdCancelar.Visible = False
    cmdGrabar.Visible = False
    cmdSeleccionar.Enabled = True
    CargaPlantillas
    fgPlantillas.SetFocus
End If
Exit Sub
ErrHandler:
   MsgBox " Grabación Cancelada ", vbInformation, " Aviso "
   
End Sub
Private Sub Grabar()
    oPlant.GrabaPlantilla txtPlantillaId, Trim(txtReferencia), rtfPlantilla.Text
End Sub
Private Sub cmdNuevo_Click()
lbNuevo = True
cmdNuevo.Visible = False
cmdEditar.Visible = False
cmdCancelar.Visible = True
cmdGrabar.Visible = True
cmdSeleccionar.Enabled = False

fraDatos.Enabled = True
fgPlantillas.Enabled = False
txtPlantillaId = oNPlant.GetNewCodPlantilla(lsOpeCod)
txtReferencia = ""
rtfPlantilla.Text = ""
txtReferencia.SetFocus

End Sub

Private Sub cmdSeleccionar_Click()
If txtPlantillaId = "" Then
    MsgBox "Plantilla No válida", vbInformation, "Aviso"
    Exit Sub
End If
lsPlantillaID = Trim(txtPlantillaId)
lnMagDer = Me.txtMgDer.Valor
lnMagIzq = Me.txtMgIzq.Valor
lnMagSup = Me.txtMgSup.Valor
Unload Me
End Sub

Private Sub fgPlantillas_Click()
txtPlantillaId = fgPlantillas.TextMatrix(fgPlantillas.Row, 1)
txtReferencia = fgPlantillas.TextMatrix(fgPlantillas.Row, 2)
rtfPlantilla = fgPlantillas.TextMatrix(fgPlantillas.Row, 3)
End Sub

Private Sub fgPlantillas_GotFocus()
txtPlantillaId = fgPlantillas.TextMatrix(fgPlantillas.Row, 1)
txtReferencia = fgPlantillas.TextMatrix(fgPlantillas.Row, 2)
rtfPlantilla = fgPlantillas.TextMatrix(fgPlantillas.Row, 3)
End Sub

Private Sub fgPlantillas_RowColChange()
txtPlantillaId = fgPlantillas.TextMatrix(fgPlantillas.Row, 1)
txtReferencia = fgPlantillas.TextMatrix(fgPlantillas.Row, 2)
rtfPlantilla = fgPlantillas.TextMatrix(fgPlantillas.Row, 3)
End Sub
Private Sub Form_Load()
Dim sText As String
Set nContFunc = New NContFunciones
Set oNPlant = New NPlantilla
Set oPlant = New dPlantilla
Dim oConst As NConstSistemas

Set oConst = New NConstSistemas

Me.fgPlantillas.Enabled = True
Me.fraDatos.Enabled = False
txtMgDer.Valor = lnMagDer
txtMgIzq.Valor = lnMagIzq
txtMgSup.Valor = lnMagSup

Set oConst = Nothing
CargaPlantillas
LlenaCboCampos
End Sub
Public Sub Inicio(psPlantillaID As String, ByVal psOpeCod As String)
lsPlantillaID = psPlantillaID
lsOpeCod = psOpeCod
Me.Show 1
End Sub
Private Sub LlenaCboCampos()
cboCampos.Clear
cboCampos.AddItem "Nro. Movimiento"
cboCampos.AddItem "Fecha          "
cboCampos.AddItem "Entidad Origen "
cboCampos.AddItem "Entidad Destino"
cboCampos.AddItem "Cuenta Origen  "
cboCampos.AddItem "Cuenta Destino "
cboCampos.AddItem "Importe        "
cboCampos.AddItem "Nro. Documento "
cboCampos.ListIndex = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set nContFunc = Nothing
Set oNPlant = Nothing
Set oPlant = Nothing

End Sub
Private Sub mnuCampos_Click(Index As Integer)
nLastPos = rtfPlantilla.SelStart
InsertaCampo True, Index
End Sub
Private Sub rtfPlantilla_LostFocus()
nLastPos = rtfPlantilla.SelStart
End Sub
Private Sub rtfPlantilla_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu mnuCamposGen
End Sub
Private Sub txtMgDer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdGrabar.SetFocus
End If
End Sub
Private Sub txtMgIzq_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtMgDer.SetFocus
End If
End Sub

Private Sub txtMgSup_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtMgIzq.SetFocus
End If
End Sub
Private Sub CargaPlantillas()
Dim rs As ADODB.Recordset

Set rs = oNPlant.GetPlantillas(lsOpeCod)  'GetPlantillas

Set fgPlantillas.DataSource = rs
If rs.EOF Then
    fgPlantillas.Rows = 2
End If
fgPlantillas.ColWidth(0) = 180
fgPlantillas.ColWidth(1) = 0
fgPlantillas.ColWidth(2) = 1800
fgPlantillas.ColWidth(3) = 0
fgPlantillas.Row = 0
fgPlantillas.Col = 0
fgPlantillas.Text = ""
fgPlantillas.Col = 2
fgPlantillas.Text = "Plantilla"
fgPlantillas.Row = 1
If fgPlantillas.Enabled And fgPlantillas.Visible Then
    fgPlantillas.SetFocus
End If
End Sub
Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    rtfPlantilla.SetFocus
End If
End Sub
