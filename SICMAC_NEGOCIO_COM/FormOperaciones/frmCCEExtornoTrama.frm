VERSION 5.00
Begin VB.Form frmCCEExtornoTrama 
   BackColor       =   &H80000016&
   Caption         =   "Extorno De Trama CCE"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12975
   Icon            =   "frmCCEExtornoTrama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   12975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   10935
      Begin VB.TextBox txtGlosa 
         Height          =   405
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   10695
      End
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "Extornar"
      Height          =   375
      Left            =   11280
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin SICMACT.FlexEdit flxTramaExt 
      Height          =   2775
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   4895
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Item-Fecha-Sesion-Registros-Moneda-ID"
      EncabezadosAnchos=   "1200-2200-2200-2200-2200-2200"
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
      ColumnasAEditar =   "X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   1200
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame FraInstrumento 
      BackColor       =   &H80000016&
      Caption         =   "Extorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin VB.ComboBox cboInstrumento 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdReporte 
         BackColor       =   &H80000016&
         Caption         =   "Buscar Trama"
         Height          =   495
         Left            =   3840
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblCargaFecha 
         BackColor       =   &H80000016&
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTipoSesion 
         BackColor       =   &H80000016&
         Caption         =   "Sesión"
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCCEExtornoTrama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCCE As COMNCajaGeneral.NCOMCCE
Dim lsAplicacion As String
Dim bLogicoDev As Boolean
Dim bLogicoCon As Boolean
Dim btrama As Boolean 'VAPA20170721
Private Sub cboInstrumento_Click()
lsAplicacion = cboInstrumento.Text
cmdReporte.SetFocus
End Sub

Private Sub CargaDatosflxTramaExt(ByVal psSesion As String, ByVal pdFecha As Date)
    Dim rsExtorno As ADODB.Recordset
    Set rsExtorno = oCCE.CCE_ExtornoTrama(psSesion, pdFecha)
    LimpiaFlex flxTramaExt 'VAPA20170720
    If Not rsExtorno.EOF Then
    Do While Not rsExtorno.EOF
    
    flxTramaExt.AdicionaFila
    flxTramaExt.TextMatrix(flxTramaExt.row, 1) = rsExtorno!dFechaArchivo
    flxTramaExt.TextMatrix(flxTramaExt.row, 2) = rsExtorno!cCodAplicacion
    flxTramaExt.TextMatrix(flxTramaExt.row, 3) = rsExtorno!nRegistros
    flxTramaExt.TextMatrix(flxTramaExt.row, 4) = rsExtorno!cCodMoneda
    flxTramaExt.TextMatrix(flxTramaExt.row, 5) = rsExtorno!nIdIT
    rsExtorno.MoveNext
    cmdExtornar.Enabled = True 'VAPA20170720
    btrama = True 'VAPA20170721
    Loop
    Else
    Set Me.flxTramaExt.Recordset = rsExtorno
    btrama = False 'VAPA20170721
    End If
   
End Sub

Private Sub cmdExtornar_Click()
 If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Glosa del extorno", vbInformation, "¡Aviso!"
        txtGlosa.SetFocus
        Exit Sub
    End If
    'VAPA20170721
    If btrama = False Then
        MsgBox "Ud. debe seleccionar  una trama para realizar el extorno", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    
    oCCE.CCE_ExtornaTramaPre gdFecSis, gsCodAge, gsCodUser, "940002", Trim(txtGlosa.Text), flxTramaExt.TextMatrix(flxTramaExt.row, 5)
    Call CargaDatosflxTramaExt(lsAplicacion, gdFecSis)
    MsgBox "Se ha extornado con éxito la trama.", vbInformation, "¡Aviso!"
    txtGlosa.Text = ""
End Sub

Private Sub cmdReporte_Click()
Dim lbTrama As Boolean
If lsAplicacion = "" Then
 MsgBox "Debe elegir una sesión a Extornar", vbInformation, "Aviso"
 Exit Sub
End If
lbTrama = oCCE.CCE_SeEnvioTrama(lsAplicacion, gdFecSis)
If lbTrama = True Then
 Call CargaDatosflxTramaExt(lsAplicacion, gdFecSis)
 Else
 MsgBox "No existe trama enviada en esta sesión ", vbInformation, "Aviso"
 Call CargaDatosflxTramaExt(lsAplicacion, gdFecSis)
 Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
lsAplicacion = "" 'VAPA20170719
cmdExtornar.Enabled = False 'VAPA20170720
    Set oCCE = New COMNCajaGeneral.NCOMCCE
    Set rs = oCCE.CCE_ObtieneAplicacionReporte
    Do While Not rs.EOF
        cboInstrumento.AddItem rs!cCodAplicacion
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

