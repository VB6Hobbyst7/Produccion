VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmPigAnularVentaJoyas 
   Caption         =   "Anulacion de Facturas/Boletas"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7335
      TabIndex        =   15
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8775
      TabIndex        =   14
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsulta 
      Caption         =   "&Detalle"
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   13
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsulta 
      Caption         =   "&Procesar"
      Height          =   375
      Index           =   0
      Left            =   5895
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   5760
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   9930
      Begin VB.Frame frEmitidos 
         Caption         =   "Documentos Emitidos"
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
         Height          =   4860
         Left            =   150
         TabIndex        =   2
         Top             =   780
         Width           =   9645
         Begin MSComctlLib.ListView lstDocumentos 
            Height          =   4590
            Left            =   90
            TabIndex        =   3
            Top             =   180
            Width           =   9465
            _ExtentX        =   16695
            _ExtentY        =   8096
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nro.Documento"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre o Razon Social"
               Object.Width           =   6262
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Monto Venta"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Tipo Documento"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Tipo de Venta"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Fec.Venta"
               Object.Width           =   2028
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Nro.Mvto."
               Object.Width           =   2028
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Estado"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Registro"
               Object.Width           =   4322
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "nTipDoc"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.Frame frCriterios 
         Caption         =   "Criterios de Busquedad"
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
         Height          =   660
         Left            =   150
         TabIndex        =   1
         Top             =   120
         Width           =   9645
         Begin MSDataListLib.DataCombo cboTipoVenta 
            Height          =   315
            Left            =   4380
            TabIndex        =   12
            Top             =   210
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cboDocumento 
            Height          =   315
            Left            =   1695
            TabIndex        =   11
            Top             =   210
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSMask.MaskEdBox mskFechaDesde 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   6960
            TabIndex        =   8
            Top             =   210
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFechaHasta 
            Height          =   315
            Left            =   8355
            TabIndex        =   10
            Top             =   210
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            Caption         =   "Al"
            Height          =   225
            Left            =   8130
            TabIndex        =   9
            Top             =   270
            Width           =   150
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha: Del"
            Height          =   210
            Left            =   6075
            TabIndex        =   7
            Top             =   285
            Width           =   810
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Venta :"
            Height          =   225
            Left            =   3225
            TabIndex        =   6
            Top             =   270
            Width           =   1110
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Documento :"
            Height          =   225
            Left            =   150
            TabIndex        =   5
            Top             =   270
            Width           =   1530
         End
      End
   End
End
Attribute VB_Name = "frmPigAnularVentaJoyas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboDocumento_Change()
frCriterios_Click
End Sub

Private Sub cboDocumento_Click(Area As Integer)
If cboDocumento.Text = "POLIZA" Then
   cboTipoVenta.Text = "A TERCEROS"
   cboTipoVenta.Enabled = False
Else
   cboTipoVenta.Text = ""
   cboTipoVenta.Enabled = True
End If
End Sub

Private Sub cboTipoVenta_Change()
frCriterios_Click
End Sub

Private Sub cmdCancelar_Click()
lstDocumentos.ListItems.Clear
Limpiar
mskFechaDesde.SelLength = 10
mskFechaDesde.SetFocus
End Sub

Private Sub cmdConsulta_Click(Index As Integer)
Select Case Index
            Case 0
                      lstDocumentos.ListItems.Clear
                      MuestraVentaJoyas
            Case 1
                      If cboDocumento.BoundText = gPigTipoPoliza Then
                          frmPigDetalleRemate.Inicio lstDocumentos.SelectedItem.SubItems(9), lstDocumentos.SelectedItem, _
                                                                          lstDocumentos.SelectedItem.SubItems(5), lstDocumentos.SelectedItem.SubItems(7)
                      Else
                          frmPigDetalleVenta.Inicio lstDocumentos.SelectedItem.SubItems(9), lstDocumentos.SelectedItem, _
                                                                       lstDocumentos.SelectedItem.SubItems(5)
                      End If
End Select
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub CargaDocumentos()
Dim lrDatos As ADODB.Recordset
Dim lrDocumento As DPigContrato

Set lrDatos = New ADODB.Recordset
Set lrDocumento = New DPigContrato
      Set lrDatos = lrDocumento.dEvaluacion(gColocPigTipoDocumento)
Set lrDocumento = Nothing

Set cboDocumento.RowSource = lrDatos
       cboDocumento.ListField = "cConsDescripcion"
       cboDocumento.BoundColumn = "nConsValor"

Set lrDatos = Nothing
End Sub

Private Sub CargaTipoVenta()
Dim lrDatos As ADODB.Recordset
Dim lrTipoVenta As DPigContrato

Set lrDatos = New ADODB.Recordset
Set lrTipoVenta = New DPigContrato
      Set lrDatos = lrTipoVenta.dEvaluacion(gColocPigTipoVentaJoya)
Set lrTipoVenta = Nothing

Set cboTipoVenta.RowSource = lrDatos
       cboTipoVenta.ListField = "cConsDescripcion"
       cboTipoVenta.BoundColumn = "nConsValor"

Set lrDatos = Nothing
End Sub

Private Sub Form_Activate()
mskFechaDesde.SelLength = 10
mskFechaDesde.SetFocus
End Sub

Private Sub Form_Load()
Limpiar
End Sub

Private Sub MuestraVentaJoyas()
Dim lrDatos As ADODB.Recordset
Dim lrPigContrato As DPigContrato
Dim lstTmpDocumentos As ListItem
Dim lsFechaDesde As String
Dim lsFechaHasta As String

lsFechaDesde = Mid(mskFechaDesde.Text, 7, 4) & Mid(mskFechaDesde.Text, 4, 2) & Mid(mskFechaDesde.Text, 1, 2)
lsFechaHasta = Mid(mskFechaHasta.Text, 7, 4) & Mid(mskFechaHasta.Text, 4, 2) & Mid(mskFechaHasta.Text, 1, 2)

Set lrDatos = New ADODB.Recordset
Set lrPigContrato = New DPigContrato
       Set lrDatos = lrPigContrato.dObtieneDocumentosVentasJoyas(lsFechaDesde, lsFechaHasta, _
                                                                                                                      Val(cboDocumento.BoundText), Val(cboTipoVenta.BoundText))
Set lrPigContrato = Nothing

If lrDatos.EOF And lrDatos.BOF Then
    Exit Sub
Else
    lstDocumentos.ListItems.Clear
    Do While Not lrDatos.EOF
          Set lstTmpDocumentos = lstDocumentos.ListItems.Add(, , Mid(lrDatos!cDocumento, 1, 4) & "-" & Mid(lrDatos!cDocumento, 5, 8))
                 lstTmpDocumentos.SubItems(1) = PstaNombre(Trim(lrDatos!Cliente))
                 lstTmpDocumentos.SubItems(2) = Format(lrDatos!nMonto, "###,##0.00")
                 lstTmpDocumentos.SubItems(3) = Trim(lrDatos!DesTipDoc)
                 lstTmpDocumentos.SubItems(4) = Trim(lrDatos!DesTipVta)
                 lstTmpDocumentos.SubItems(5) = Mid(lrDatos!dFechaVta, 7, 2) & "/" & Mid(lrDatos!dFechaVta, 5, 2) & "/" & Mid(lrDatos!dFechaVta, 1, 4)
                 lstTmpDocumentos.SubItems(6) = Format(lrDatos!nNroMov, "00000000")
                 If lrDatos!nFlag = gMovFlagExtornado Then
                     lstTmpDocumentos.SubItems(7) = "ANULADA"
                 Else
                     lstTmpDocumentos.SubItems(7) = "ACTIVA"
                 End If
                 lstTmpDocumentos.SubItems(8) = Trim(lrDatos!cUltimaActualizacion)
                 lstTmpDocumentos.SubItems(9) = Trim(lrDatos!nCodTipo)
          lrDatos.MoveNext
    Loop
    Set lrDatos = Nothing
    cmdConsulta(0).Visible = False
    cmdConsulta(1).Visible = True
End If
End Sub

Private Sub Limpiar()
cboTipoVenta.Enabled = True
cboDocumento.Text = ""
cboTipoVenta.Text = ""

mskFechaDesde.Text = Format(gdFecSis, "dd/mm/yyyy")
mskFechaHasta.Text = Format(gdFecSis, "dd/mm/yyyy")
lstDocumentos.ListItems.Clear
cmdConsulta(1).Visible = False
cmdConsulta(0).Visible = True
CargaDocumentos
CargaTipoVenta
End Sub

Private Sub frCriterios_Click()
cmdConsulta(0).Visible = True
cmdConsulta(1).Visible = False


End Sub

Private Sub mskFechaDesde_Change()
frCriterios_Click
End Sub

Private Sub mskFechaDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    mskFechaHasta.SelLength = 10
    mskFechaHasta.SetFocus
End If
End Sub

Private Sub mskFechaHasta_Change()
frCriterios_Click
End Sub

Private Sub mskFechaHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdConsulta_Click (0)
End If
End Sub
