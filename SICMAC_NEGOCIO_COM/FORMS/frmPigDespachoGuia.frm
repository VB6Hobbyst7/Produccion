VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPigDespachoGuia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Despacho de Valija"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   Icon            =   "frmPigDespachoGuia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
      Height          =   345
      Left            =   10590
      TabIndex        =   5
      Top             =   150
      Width           =   360
   End
   Begin VB.TextBox txtTransporte 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   6930
      TabIndex        =   4
      Top             =   165
      Width           =   3525
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9795
      TabIndex        =   3
      Top             =   5790
      Width           =   1200
   End
   Begin VB.CommandButton cmdDespachar 
      Caption         =   "&Despachar"
      Height          =   375
      Left            =   8460
      TabIndex        =   2
      Top             =   5790
      Width           =   1200
   End
   Begin SICMACT.FlexEdit feGuias 
      Height          =   4680
      Left            =   105
      TabIndex        =   0
      Top             =   1035
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   8255
      Cols0           =   11
      FixedCols       =   2
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-No-C-Nro Guia-Destino-Motivo-Cant-Clase-SSunat-NumSunat-nMotivo"
      EncabezadosAnchos=   "0-400-400-1200-2400-2900-700-900-700-1300-0"
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
      ColumnasAEditar =   "X-X-2-X-X-X-X-X-8-9-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-R-C-L-L-L-R-L-L-C-C"
      FormatosEdit    =   "0-3-0-0-1-0-2-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
      CellForeColor   =   -2147483630
      CellBackColor   =   -2147483633
   End
   Begin MSDataListLib.DataCombo cboMotivo 
      Height          =   315
      Left            =   750
      TabIndex        =   6
      Top             =   165
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboOrigen 
      Height          =   315
      Left            =   750
      TabIndex        =   9
      Top             =   600
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Origen "
      Height          =   225
      Left            =   150
      TabIndex        =   8
      Top             =   660
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Motivo"
      Height          =   255
      Left            =   165
      TabIndex        =   7
      Top             =   225
      Width           =   555
   End
   Begin VB.Label Label3 
      Caption         =   "Transportista"
      Height          =   270
      Left            =   5850
      TabIndex        =   1
      Top             =   240
      Width           =   1005
   End
End
Attribute VB_Name = "frmPigDespachoGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim lsPersCod As String, lsPersNombre As String
'
'Private Sub cboMotivo_Click(Area As Integer)
'
'    feGuias.Clear
'    feGuias.Rows = 2
'    feGuias.FormaCabecera
'
'    cboOrigen.Enabled = True
'    cboOrigen.BoundText = -1
'
'    If cboMotivo.BoundText <> "" Then
'        Select Case cboMotivo.BoundText
'        Case 1  'Contratos Nuevos
'            cboOrigen.BoundText = CInt(gsCodAge)
'            cboOrigen.Enabled = False
'        Case 2  'Pendientes de Rescate > 30 dias
'            cboOrigen.BoundText = CInt(gsCodAge)
'            cboOrigen.Enabled = False
'        Case 3  'Pendientes de Rescate
'            cboOrigen.BoundText = 99
'        Case 4  'Lotes de la 1era Subasta
'            cboOrigen.BoundText = 99
'            cboOrigen.Enabled = False
'        Case 5  'Devolucion de Remate
'            cboOrigen.Enabled = True
'        Case 6  'Joyas para Venta en OP
'            cboOrigen.BoundText = 99
'            cboOrigen.Enabled = False
'        Case 7  'Devolucion de Joyas de OP
'            cboOrigen.Enabled = True
'        Case 8  'Joyas para Ventas y Remates
'            cboOrigen.BoundText = 99
'            cboOrigen.Enabled = False
'        Case 9  'Devolucion de Joyas de Ventas y Remates
'            cboOrigen.Enabled = True
'        End Select
'    End If
'
'    If cboMotivo.Text <> "" And cboOrigen.Text <> "" Then
'        CargaGuias
'    End If
'
'End Sub
'Private Sub CargaGuias()
'Dim oPigCont As DPigContrato
'Dim RS As Recordset
'
'feGuias.Clear
'feGuias.FormaCabecera
'feGuias.Rows = 2
'
'Set oPigCont = New DPigContrato
'Set RS = oPigCont.dObtieneGuias(1, cboOrigen.BoundText, , cboMotivo.BoundText)
'
'Do While Not RS.EOF
'    feGuias.AdicionaFila
'    feGuias.TextMatrix(feGuias.Rows - 1, 1) = feGuias.Rows - 1
'    feGuias.TextMatrix(feGuias.Rows - 1, 3) = RS!cNumDoc
'    feGuias.TextMatrix(feGuias.Rows - 1, 4) = RS!Destino
'    feGuias.TextMatrix(feGuias.Rows - 1, 5) = RS!Motivo
'    feGuias.TextMatrix(feGuias.Rows - 1, 6) = RS!nTotItem
'    feGuias.TextMatrix(feGuias.Rows - 1, 7) = RS!TipoGuia
'    feGuias.TextMatrix(feGuias.Rows - 1, 10) = RS!nMotivo
'    RS.MoveNext
'Loop
'
'Set RS = Nothing
'Set oPigCont = Nothing
'
'End Sub
'
'Private Sub cboOrigen_Click(Area As Integer)
'    If cboMotivo.Text <> "" And cboOrigen.Text <> "" Then
'        CargaGuias
'    End If
'End Sub
'
'Private Sub cboOrigen_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If cboMotivo.Text <> "" And cboOrigen <> "" Then
'            CargaGuias
'        End If
'    End If
'End Sub
'
'Private Sub cmdBuscar_Click()
'Dim oPers As UPersona
'
'Set oPers = New UPersona
'    Set oPers = frmBuscaPersona.Inicio
'    If oPers Is Nothing Then Exit Sub
'    lsPersCod = oPers.sPersCod
'    lsPersNombre = oPers.sPersNombre
'    txtTransporte = lsPersNombre
'Set oPers = Nothing
'
'End Sub
'
'Private Sub cmdDespachar_Click()
'Dim oContFunc As NContFunciones
'Dim lsNumGuia As String
'Dim lsSerieSunat As String
'Dim lsNumSunat As String
'Dim lsMovNro As String
'Dim lnUbicTransito As Integer
'Dim oPigImpre As NPigImpre
'Dim lsCadImp As String
'Dim lnTipoGuia As Integer
'Dim oPrevio As Previo.clsPrevio
'Dim lbGuia As Boolean
'Dim oDatos As DPigContrato
'Dim oPigRemesa As NPigContrato
'Dim RS As Recordset
'
'If ValidaGrabar Then
'
'    lnUbicTransito = 16     'Ubicacion del Lote o de la Pieza en Transito
'
'    Set oContFunc = New NContFunciones
'    Set oPigRemesa = New NPigContrato
'    Set oPigImpre = New NPigImpre
'
'    For i = 1 To feGuias.Rows - 1
'
'        If feGuias.TextMatrix(i, 2) = "." Then 'Guia seleccionada para despacho
'
'            lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'
'            lsNumGuia = feGuias.TextMatrix(i, 3)
'            lsSerieSunat = Right(feGuias.TextMatrix(i, 8), 3)
'            lsNumSunat = Right(feGuias.TextMatrix(i, 9), 12)
'
'            If feGuias.TextMatrix(i, 7) = "LOTE" Then
'                lnTipoGuia = 1
'            ElseIf feGuias.TextMatrix(i, 7) = "PIEZA" Then
'                lnTipoGuia = 2
'            End If
'
'            Set oDatos = New DPigContrato
'            Set RS = oDatos.dObtieneValorRemesa(lsNumGuia)
'            Set oDatos = Nothing
'
'            oPigRemesa.nDespachoRemesa lsMovNro, lsNumGuia, lsSerieSunat, lsNumSunat, lsPersCod, lnTipoGuia, lnUbicTransito, RS, feGuias.TextMatrix(i, 10)
'            Set RS = Nothing
'
'            lsCadImp = lsCadImp & oPigImpre.ImpreGuia(lsNumGuia, lnTipoGuia, gsInstCmac, gdFecSis)
'
'        End If
'
'    Next i
'
'    If lsCadImp <> "" Then  'ACA FALTA VERIFICAR QUE SE VISUALICE EL PREVIO PARA LA IMPRESION DE LA GUIA
'        Set oPrevio = New Previo.clsPrevio
'        'oPrevio.PrintSpool sLpt, lsCadImpre, False, 66
'        clsPrevio.Show lsCadImp, "Despacho de Remesa", False, 66
'    End If
'
'    Set oPrevio = Nothing
'    Set oPigImpre = Nothing
'    Set oContFunc = Nothing
'    Set oPigRemesa = Nothing
'
'    feGuias.Clear
'    feGuias.Rows = 2
'    feGuias.FormaCabecera
'    txtTransporte = ""
'
'End If
'
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'Dim RS As Recordset
'
'Call CargaCombo(cboMotivo, gColocPigMotivoRem)
'Call CargaCombo(cboOrigen, gColocPigUbicacion)
'
'End Sub
'
'Private Function ValidaGrabar() As Boolean
'
'    ValidaGrabar = True
'
'    If txtTransporte = "" Then
'        MsgBox "Seleccione una empresa de Transporte de la Remesa", vbInformation, "Aviso"
'        ValidaGrabar = False
'        Exit Function
'    End If
'
'
'End Function
'
'Private Sub CargaCombo(Combo As DataCombo, ByVal psConsCod As String)
'Dim oPigFunc As DPigFunciones
'Dim RS As Recordset
'
'Set oPigFunc = New DPigFunciones
'
'    Set RS = oPigFunc.GetConstante(psConsCod)
'    Set Combo.RowSource = RS
'    Combo.ListField = "cConsDescripcion"
'    Combo.BoundColumn = "nConsValor"
'
'    Set RS = Nothing
'
'    Set oPigFunc = Nothing
'
'End Sub
'
