VERSION 5.00
Begin VB.Form frmLogOrdenDet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Orden"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13320
   Icon            =   "frmLogOrdenDet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   10950
      TabIndex        =   18
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   12120
      TabIndex        =   17
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame fraOrden 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Orden"
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
      Height          =   1815
      Left            =   50
      TabIndex        =   2
      Top             =   0
      Width           =   13215
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Glosa:"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblGlosa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   1440
         TabIndex        =   15
         Top             =   1080
         Width           =   11655
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha:"
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
         Left            =   5040
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fecha"
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
         Height          =   270
         Left            =   5760
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo:"
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
         Left            =   3360
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblTipoOrden 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OCS"
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
         Height          =   270
         Left            =   3960
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nro. Orden:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblNroOrden 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NOrden"
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
         Height          =   270
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Proveedor:"
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
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Proveedor"
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
         Height          =   270
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   11655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Moneda:"
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
         Left            =   6960
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Moneda"
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
         Height          =   270
         Left            =   7800
         TabIndex        =   5
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Monto:"
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
         Left            =   9000
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblMonto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Monto"
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
         Height          =   270
         Left            =   9720
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin Sicmact.FlexEdit dgOrden 
      Height          =   3375
      Left            =   45
      TabIndex        =   1
      Top             =   2160
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   5953
      Cols0           =   11
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Agencia-Descripción-Moneda-Monto-Estado-Fecha-Tpo Doc-Nro Doc-User-nMovNro"
      EncabezadosAnchos=   "300-1000-3500-800-800-2200-800-1500-1400-800-1000"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-C-R-L-L-L-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label lblCronograma 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CRONOGRAMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1850
      Width           =   1900
   End
End
Attribute VB_Name = "frmLogOrdenDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gnMovNro As Long
Dim gsDocNro As String
Dim gnDocTpo As Integer
Public Function Inicio(ByVal pnMovNro As Long, ByVal psNroOrden As String, ByVal psTpoOrden As String, ByVal psFecha As String, ByVal psMoneda As String, ByVal pnMonto As Currency, ByVal psProveedor As String, ByVal psGlosa As String)
    gnMovNro = pnMovNro
    gsDocNro = psNroOrden
    gnDocTpo = IIf(psTpoOrden = "O/C", 90, 97)
    lblNroOrden.Caption = psNroOrden
    lblTipoOrden.Caption = psTpoOrden
    lblFecha.Caption = psFecha
    lblMoneda.Caption = psMoneda
    lblMonto.Caption = Format(pnMonto, "#,#0.00")
    lblProveedor.Caption = UCase(psProveedor)
    lblGlosa.Caption = psGlosa
    Me.Show 1
End Function
Private Sub cmdActualizar_Click()
    Dim rs As ADODB.Recordset
    Dim row As Integer
    Dim oProv As DLogGeneral
    Set oProv = New DLogGeneral
    lblCronograma.Caption = "Cargando..."
    LimpiaFlex dgOrden
    cmdActualizar.Enabled = False
    DoEvents
    Set rs = oProv.ObtieneOrdenCab(Trim(lblNroOrden.Caption), gnDocTpo)
    Do While Not rs.EOF
        lblNroOrden.Caption = rs!cDocNro
        lblTipoOrden.Caption = rs!Tipo
        lblFecha.Caption = Format(rs!dDocFecha, "dd/MM/YYYY")
        lblMoneda.Caption = rs!Moneda
        lblMonto.Caption = Format(rs!nMonto, "#,#0.00")
        lblProveedor.Caption = UCase(rs!cPersNombre)
        lblGlosa.Caption = rs!cMovDesc
        rs.MoveNext
    Loop
    gsDocNro = lblNroOrden.Caption
    gnDocTpo = IIf(lblTipoOrden.Caption = "O/C", 90, 97)
    CargaOrdenDet
    cmdActualizar.Enabled = True
    lblCronograma.Caption = "CRONOGRAMA"
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    CargaOrdenDet
End Sub
Private Sub CargaOrdenDet()
    Dim rs As ADODB.Recordset
    Dim row As Integer
    Dim oProv As DLogGeneral
    Set oProv = New DLogGeneral
    Set rs = oProv.ObtieneOrdenDet(gsDocNro, gnDocTpo)
    Do While Not rs.EOF
        dgOrden.AdicionaFila
        row = dgOrden.row
        dgOrden.TextMatrix(row, 1) = rs!Agencia
        dgOrden.TextMatrix(row, 2) = rs!descripcion
        dgOrden.TextMatrix(row, 3) = rs!Moneda
        dgOrden.TextMatrix(row, 4) = Format(rs!monto, "#,#0.00")
        dgOrden.TextMatrix(row, 5) = rs!Estado
        dgOrden.TextMatrix(row, 6) = rs!fecha
        dgOrden.TextMatrix(row, 7) = rs!TpoDoc
        dgOrden.TextMatrix(row, 8) = rs!NroDoc
        dgOrden.TextMatrix(row, 9) = rs!usuario
        dgOrden.TextMatrix(row, 10) = rs!nMovNro
        rs.MoveNext
    Loop
End Sub

