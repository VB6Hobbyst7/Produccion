VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   Caption         =   "Cambia"
   ClientHeight    =   4155
   ClientLeft      =   1890
   ClientTop       =   1950
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   6225
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   360
      Left            =   3660
      TabIndex        =   4
      Top             =   1245
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   19
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin Sicmact.FlexEdit FlexEdit1 
      Height          =   2085
      Left            =   255
      TabIndex        =   3
      Top             =   1755
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   3678
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      VisiblePopMenu  =   -1  'True
      EncabezadosNombres=   "-Check-Option-texto"
      EncabezadosAnchos=   "350-500-500-1200"
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
      ColumnasAEditar =   "X-1-2-3"
      ListaControles  =   "0-4-5-0"
      EncabezadosAlineacion=   "C-C-C-L"
      FormatosEdit    =   "0-0-0-0"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   930
      Left            =   390
      TabIndex        =   2
      Top             =   1110
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   1640
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   2490
      TabIndex        =   1
      Top             =   195
      Width           =   1365
   End
   Begin Sicmact.Usuario Usuario1 
      Left            =   2655
      Top             =   885
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin Sicmact.TxtBuscar TxtBuscar1 
      Height          =   375
      Left            =   435
      TabIndex        =   0
      Top             =   210
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   661
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      sTitulo         =   ""
      EnabledText     =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
'gsCodAge = TxtBuscar1.Text
Ejemplo
End Sub
Sub Ejemplo()
Me.FlexEdit1.AdicionaFila , , True
'Dim oconect As DConecta
'Set oconect = New DConecta
'oconect.AbreConexion
''Me.MSHFlexGrid1.RowHeight(0) = 150
'For i = 1 To 10
    
'    If i = 5 Then
'        FlexEdit1.TextMatrix(FlexEdit1.Row, 2) = "1"
'    End If
'    FlexEdit1.TextMatrix(FlexEdit1.Row, 1) = ""
    
'Next
'CentraForm Me
'Dim oArea As DActualizaDatosArea
'Set oArea = New DActualizaDatosArea
'TxtBuscar1.rs = oArea.GetAgencias
'Set oArea = Nothing
End Sub

