VERSION 5.00
Begin VB.Form frmLogSelListadoBienes 
   Caption         =   "Listado de Bienes"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   Icon            =   "frmLogSelListadoBienes.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5085
   ScaleWidth      =   10575
   Begin VB.Frame s 
      Caption         =   "Proceso de  Seleccion"
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtdescripcion 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   5895
      End
      Begin VB.TextBox txttipo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   5895
      End
      Begin Sicmact.TxtBuscar txtSeleccion 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
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
         Enabled         =   0   'False
         Enabled         =   0   'False
         TipoBusqueda    =   2
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin VB.Label Label7 
         Caption         =   "Numero"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
   Begin Sicmact.FlexEdit fgeBienesPlantilla 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   5318
      Cols0           =   9
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "Item-Código-Descripción-Unidad-ValorUnidad-Descripcion Adicional-Cantidad-Precio Ref-Sub Total"
      EncabezadosAnchos=   "450-1200-3500-700-0-1000-1000-1000-1200"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "R-L-L-L-R-L-R-C-R"
      FormatosEdit    =   "0-0-0-0-3-0-3-2-2"
      CantEntero      =   6
      CantDecimales   =   1
      AvanceCeldas    =   1
      TextArray0      =   "Item"
      lbEditarFlex    =   -1  'True
      Enabled         =   0   'False
      lbFlexDuplicados=   0   'False
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   450
      RowHeight0      =   300
   End
End
Attribute VB_Name = "frmLogSelListadoBienes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDGnral As DLogGeneral
Dim clsDGAdqui As DLogAdquisi
Dim rs As New ADODB.Recordset
Dim bcontBienes As Boolean


Private Sub CmdAceptar_Click()
If bcontBienes = False Then Exit Sub
Set frmLogSelSeleccionBienes.fgeBienesConfig.Recordset = clsDGAdqui.CargaSelDetalle(frmLogSelSeleccionBienes.txtSeleccion.Text, 1)
frmLogSelSeleccionBienes.fgeBienesConfig.AdicionaFila
frmLogSelSeleccionBienes.fgeBienesConfig.TextMatrix(frmLogSelSeleccionBienes.fgeBienesConfig.Rows - 1, 1) = " --------------------------- "
frmLogSelSeleccionBienes.fgeBienesConfig.TextMatrix(frmLogSelSeleccionBienes.fgeBienesConfig.Rows - 1, 2) = " --------------------- TOTAL ------------- "
frmLogSelSeleccionBienes.fgeBienesConfig.TextMatrix(frmLogSelSeleccionBienes.fgeBienesConfig.Rows - 1, 3) = " ------------ "
frmLogSelSeleccionBienes.fgeBienesConfig.TextMatrix(frmLogSelSeleccionBienes.fgeBienesConfig.Rows - 1, 5) = " --------------- " 'fgeBienesConfig.SumaRow(5)
frmLogSelSeleccionBienes.fgeBienesConfig.TextMatrix(frmLogSelSeleccionBienes.fgeBienesConfig.Rows - 1, 6) = " ------------ "
frmLogSelSeleccionBienes.fgeBienesConfig.TextMatrix(frmLogSelSeleccionBienes.fgeBienesConfig.Rows - 1, 7) = Format(frmLogSelSeleccionBienes.fgeBienesConfig.SumaRow(7), "########.00")
Unload Me
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
Set clsDGAdqui = New DLogAdquisi
CentraSdi Me
txtSeleccion.Text = frmLogSelSeleccionBienes.txtSeleccion.Text
bcontBienes = False
Set rs = clsDGAdqui.CargaSelDetalle(frmLogSelSeleccionBienes.txtSeleccion.Text, 1)
If Not rs.EOF = True Then
    Set fgeBienesPlantilla.Recordset = rs
    bcontBienes = True
Else
    fgeBienesPlantilla.Clear
    fgeBienesPlantilla.FormaCabecera
    fgeBienesPlantilla.Rows = 2
End If
mostrar_descripcion (frmLogSelSeleccionBienes.txtSeleccion.Text)

End Sub
Sub mostrar_descripcion(nLogSelProceso As Long)
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs = clsDGAdqui.CargaLogSelDescripcionProceso(nLogSelProceso)
If rs.EOF = True Then
    txttipo.Text = ""
    txtdescripcion.Text = ""
    Else
    txttipo.Text = UCase(rs!cTipo)
    txtdescripcion.Text = rs!cDescripcionProceso
End If
End Sub

