VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPigSeleccionRemate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proceso de Seleccion de Contratos"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   Icon            =   "frmPigSeleccionRemate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRemate 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1470
      TabIndex        =   7
      Top             =   210
      Width           =   705
   End
   Begin VB.CommandButton cbsalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   4695
      TabIndex        =   6
      Top             =   4845
      Width           =   1260
   End
   Begin VB.CommandButton cbCargar 
      Caption         =   "&Cargar"
      Height          =   345
      Left            =   3375
      TabIndex        =   5
      Top             =   4845
      Width           =   1260
   End
   Begin VB.CommandButton cmdSeleccion 
      Caption         =   "&Seleccion"
      Height          =   345
      Left            =   4785
      TabIndex        =   4
      Top             =   645
      Width           =   1065
   End
   Begin SICMACT.FlexEdit feContrato 
      Height          =   3600
      Left            =   60
      TabIndex        =   3
      Top             =   1125
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6350
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Item-Contratos-Nro.Piezas-Dias Venc"
      EncabezadosAnchos=   "400-3200-900-1200"
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
      ColumnasAEditar =   "X-X-X-X"
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "R-L-C-R"
      FormatosEdit    =   "3-0-0-3"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin MSMask.MaskEdBox MskIniRemate 
      Height          =   315
      Left            =   4755
      TabIndex        =   2
      Top             =   180
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Inicio Remate"
      Height          =   210
      Left            =   3090
      TabIndex        =   1
      Top             =   255
      Width           =   1470
   End
   Begin VB.Label Label1 
      Caption         =   "Nro. de Remate"
      Height          =   180
      Left            =   210
      TabIndex        =   0
      Top             =   270
      Width           =   1200
   End
End
Attribute VB_Name = "frmPigSeleccionRemate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim lnDiasVencido As Integer
'
'Private Sub cbCargar_Click()
'Dim lrInserta As DPigActualizaBD
'Dim i As Integer
'Dim lsCtaCod As String
'Dim lsMovAct As String
'Dim loContFunct As NContFunciones
'
'Set loContFunct = New NContFunciones
'    lsMovAct = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'Set loContFunct = Nothing
'
'Set lrInserta = New DPigActualizaBD
'Call lrInserta.dInsertaRegistroSel(MskIniRemate.Text, lnDiasVencido, lsMovAct)
'Call lrInserta.dUpdateColocPignosel(3, MskIniRemate.Text, lnDiasVencido)
'
'MsgBox "Seleccion de Contratos Incluidos"
'Limpia
'
'End Sub
'
'Private Sub cbsalir_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdSeleccion_Click()
'Dim oParam As DPigFunciones
'Dim loContratos As DPigContrato
'Dim rs As Recordset
'
'    feContrato.Clear
'    feContrato.Rows = 2
'    feContrato.FormaCabecera
'
'    Set oParam = New DPigFunciones
'        lnDiasVencido = oParam.GetParamValor(gPigParamDiasAtrasoSelecRemate)
'    Set oParam = Nothing
'
'    Set loContratos = New DPigContrato
'
'    Set rs = loContratos.dObtieneSeleccionContratos(lnDiasVencido, MskIniRemate)
'    Set loContratos = Nothing
'
'    If Not rs.EOF And Not rs.BOF Then Set feContrato.Recordset = rs
'
'    Set rs = Nothing
'    cbCargar.SetFocus
'
'End Sub
'
'Private Sub Form_Load()
'Me.Icon = LoadPicture(App.path & gsRutaIcono)
'End Sub
'
'Private Sub txtRemate_KeyPress(KeyAscii As Integer)
'Dim oRemate As DPigRemate
'Dim rs As Recordset
'
'    If KeyAscii = 13 Then
'
'        Set oRemate = New DPigRemate
'        Set rs = oRemate.GetNumRemate(txtRemate)
'
'        If Not (rs.EOF And rs.BOF) Then
'            txtRemate = rs!NumRemate
'            MskIniRemate = rs!dInicio
'        Else
'            MsgBox "Número de Remate no Existe"
'            txtRemate.Text = ""
'            txtRemate.SetFocus
'        End If
'
'        Set rs = Nothing
'        Set oRemate = Nothing
'    End If
'
'End Sub
'
'Private Sub Limpia()
'    MskIniRemate.Mask = ""
'    feContrato.Clear
'    feContrato.FormaCabecera
'    feContrato.Rows = 2
'    txtRemate = ""
'End Sub
