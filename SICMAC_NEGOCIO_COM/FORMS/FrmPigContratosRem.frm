VERSION 5.00
Begin VB.Form FrmPigContratosRem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contratos Actualmente en Remate"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   Icon            =   "FrmPigContratosRem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   5445
      TabIndex        =   2
      Top             =   4665
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
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
      Height          =   375
      Left            =   6615
      TabIndex        =   1
      Top             =   4665
      Width           =   1095
   End
   Begin SICMACT.FlexEdit fecontrato 
      Height          =   4485
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   7911
      Cols0           =   8
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Item-Nro Contrato-Piezas-Estado-Ubicacion-Dias Atraso-CodEstado-nUbica"
      EncabezadosAnchos=   "400-1800-600-1600-2200-1000-1200-1200"
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
      EncabezadosAlineacion=   "C-C-C-L-L-C-R-R"
      FormatosEdit    =   "0-0-0-0-0-0-0-0"
      TextArray0      =   "Item"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label lblNumRemate 
      Height          =   225
      Left            =   4110
      TabIndex        =   3
      Top             =   4215
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "FrmPigContratosRem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Sub CmdAceptar_Click()
'    Dim oDatos As DPigRemate
'    Dim lsCtaCod As String
'    Dim lsEstado As Long
'    Dim lnUbicaLote As Integer, lnUbicaRemate As Integer
'    Set oDatos = New DPigRemate
'    lnUbicaRemate = oDatos.GetUbicaLoteRemate(FrmPigVentaRemate.txtRemate)
'    ' Cta Codigo del contrato
'    lsCtaCod = fecontrato.TextMatrix(fecontrato.Row, 1)
'    lsEstado = fecontrato.TextMatrix(fecontrato.Row, 6)
'    lnUbicaLote = fecontrato.TextMatrix(fecontrato.Row, 7)
'    'Validacion de Contratos - que solo se encuentren remates
'     If lsEstado = 2808 Or lsEstado = 2807 Or lsEstado = 2809 Then
'      If lnUbicaLote = lnUbicaRemate Then
'        FrmPigVentaRemate.AXCodCta.NroCuenta = lsCtaCod
'        Unload Me
'        FrmPigVentaRemate.AXCodCta.SetFocusCuenta
'      Else
'        MsgBox "La ubicacion del Lote es distinta a la del Remate", vbExclamation, "Ubicacion"
'      End If
'    Else
'        MsgBox "El contrato no puede ser Rematado su estado es " + fecontrato.TextMatrix(fecontrato.Row, 3), vbExclamation, "No se puede rematar"
'    End If
'End Sub
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub fecontrato_Click()
'    cmdAceptar.SetFocus
'End Sub
'
'Private Sub fecontrato_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    Call CmdAceptar_Click
'End If
'End Sub
'
'Private Sub Form_Load()
'Dim oDatos As DPigRemate
'Dim lrDatosContrato As ADODB.Recordset
'Dim i As Integer
'    Set oDatos = New DPigRemate
'    'REVISAR QUE SOLO VISUALIZEN CONTRATOS
'     Set lrDatosContrato = oDatos.GetContratosRemate(FrmPigVentaRemate.txtRemate)
'    Set oDatos = Nothing
'    i = 1
'    If Not (lrDatosContrato.EOF) Then
'        Do While Not lrDatosContrato.EOF
'            Me.fecontrato.AdicionaFila
'            Me.fecontrato.TextMatrix(i, 1) = lrDatosContrato!cCtaCod
'            Me.fecontrato.TextMatrix(i, 2) = lrDatosContrato!npiezas
'            Me.fecontrato.TextMatrix(i, 3) = lrDatosContrato!Estado 'lrDatosContrato!nestadocont
'            Me.fecontrato.TextMatrix(i, 4) = lrDatosContrato!UbicaLote 'nubicalote
'            Me.fecontrato.TextMatrix(i, 5) = lrDatosContrato!nDiasAtraso
'            Me.fecontrato.TextMatrix(i, 6) = lrDatosContrato!nPrdEstado
'            Me.fecontrato.TextMatrix(i, 7) = lrDatosContrato!nUbicaLote
'            lrDatosContrato.MoveNext
'            i = i + 1
'        Loop
'    Set lrDatosContrato = Nothing
'    End If
'End Sub
