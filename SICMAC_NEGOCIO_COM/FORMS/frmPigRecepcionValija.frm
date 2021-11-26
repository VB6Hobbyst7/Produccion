VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPigRecepcionValija 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepción de Valija"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9825
   Icon            =   "frmPigRecepcionValija.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRecepcion 
      Caption         =   "&Recepcion"
      Height          =   375
      Left            =   7110
      TabIndex        =   1
      Top             =   5355
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8430
      TabIndex        =   0
      Top             =   5340
      Width           =   1200
   End
   Begin SICMACT.FlexEdit feGuias 
      Height          =   4680
      Left            =   60
      TabIndex        =   2
      Top             =   570
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   8255
      Cols0           =   9
      FixedCols       =   2
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-No-R-Nro Guia-Origen-Motivo-Cant-Clase-nMotivo"
      EncabezadosAnchos=   "0-400-400-1500-2500-3200-700-900-0"
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
      ColumnasAEditar =   "X-X-2-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-4-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-R-C-L-L-L-R-C-C"
      FormatosEdit    =   "0-3-0-0-1-1-3-1-0"
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
      Left            =   735
      TabIndex        =   4
      Top             =   165
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboDestino 
      Height          =   315
      Left            =   5850
      TabIndex        =   5
      Top             =   135
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Motivo"
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   240
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Destino"
      Height          =   270
      Left            =   5205
      TabIndex        =   3
      Top             =   210
      Width           =   630
   End
End
Attribute VB_Name = "frmPigRecepcionValija"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Sub cboDestino_Click(Area As Integer)
'    If cboMotivo.Text <> "" And cboDestino.Text <> "" Then
'        CargaGuias
'    End If
'End Sub
'
'Private Sub cboDestino_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If cboMotivo.Text <> "" And cboDestino.Text <> "" Then
'        CargaGuias
'    End If
'End If
'End Sub
'
'Private Sub cboMotivo_Click(Area As Integer)
'    feGuias.Clear
'    feGuias.Rows = 2
'    feGuias.FormaCabecera
'
'    cboDestino.Enabled = True
'    cboDestino.BoundText = -1
'
'    If cboMotivo.BoundText <> "" Then
'        Select Case cboMotivo.BoundText
'        Case 1  'Contratos Nuevos
'            cboDestino.BoundText = 99
'            If Not ValidaAgBovedaValores Then
'                MsgBox "Usuario no se encuentra en la Agencia de la Boveda de Valores", vbInformation, "Aviso"
'                cboDestino.BoundText = 1
'                Exit Sub
'            End If
'            cboDestino.Enabled = False
'        Case 2  'Pendientes de Rescate > 30 dias
'            cboDestino.BoundText = CInt(gsCodAge)
'            cboDestino.Enabled = False
'        Case 3  'Pendientes de Rescate
'            cboDestino.BoundText = CInt(gsCodAge)
'            cboDestino.BoundText = False
'        Case 4  'Lotes de la 1era Subasta
'            cboDestino.BoundText = CInt(gsCodAge)
'            cboDestino.Enabled = True
'        Case 5  'Devolucion de Remate
'            cboDestino.BoundText = 99
'            If Not ValidaAgBovedaValores Then
'                MsgBox "Usuario no se encuentra en la Agencia de la Boveda de Valores", vbInformation, "Aviso"
'                cboDestino.BoundText = 1
'                Exit Sub
'            End If
'        Case 6  'Joyas para Venta en OP
'            cboDestino.BoundText = CInt(gsCodAge)
'            cboDestino.Enabled = True
'        Case 7  'Devolucion de Joyas de OP
'            cboDestino.BoundText = 99
'            If Not ValidaAgBovedaValores Then
'                MsgBox "Usuario no se encuentra en la Agencia de la Boveda de Valores", vbInformation, "Aviso"
'                cboDestino.BoundText = 1
'                Exit Sub
'            End If
'            cboDestino.Enabled = False
'        Case 8  'Joyas para Ventas y Remates
'            cboDestino.BoundText = CInt(gsCodAge)
'            cboDestino.Enabled = True
'        Case 9  'Devolucion de Joyas de Ventas y Remates
'            cboDestino.BoundText = 99
'            If Not ValidaAgBovedaValores Then
'                MsgBox "Usuario no se encuentra en la Agencia de la Boveda de Valores", vbInformation, "Aviso"
'                cboDestino.BoundText = 1
'                Exit Sub
'            End If
'            cboDestino.Enabled = False
'        End Select
'    End If
'
'    If cboMotivo.Text <> "" And cboDestino.Text <> "" Then
'        CargaGuias
'    End If
'
'End Sub
'
'Private Sub cmdRecepcion_Click()
'Dim I As Integer
'Dim lsNumDoc As String
'Dim lsMovNro As String
'Dim oContFunc As NContFunciones
'Dim oPigGraba As NPigContrato
'Dim lnTipoGuia As Integer
'Dim rs As Recordset
'Dim oDatos As DPigContrato
'Dim oPigImpre As NPigImpre
'Dim lsCadImp As String
'Dim oPrevio As previo.clsPrevio
'
'    Set oContFunc = New NContFunciones
'    Set oPigGraba = New NPigContrato
'    Set oPigImpre = New NPigImpre
'
'    For I = 1 To feGuias.Rows - 1
'
'        If feGuias.TextMatrix(I, 2) = "." Then  'Realiza la recepcion
'
'            lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'
'            lsNumDoc = feGuias.TextMatrix(I, 3)
'
'            If feGuias.TextMatrix(I, 7) = "LOTE" Then
'                lnTipoGuia = 1
'            ElseIf feGuias.TextMatrix(I, 7) = "PIEZA" Then
'                lnTipoGuia = 2
'            End If
'
'            Set oDatos = New DPigContrato
'            Set rs = oDatos.dObtieneValorRemesa(lsNumDoc)
'            Set oDatos = Nothing
'
'            oPigGraba.nRecepcionRemesa lsNumDoc, lsMovNro, lnTipoGuia, cboDestino.BoundText, rs, feGuias.TextMatrix(I, 8)
'            Set rs = Nothing
'
'            lsCadImp = lsCadImp & oPigImpre.ImpreGuiaRecep(lsNumDoc, lnTipoGuia, gdFecSis, gsInstCmac)
'
'        End If
'
'    Next I
'
'    If lsCadImp <> "" Then
'        Set oPrevio = New previo.clsPrevio
'        clsPrevio.Show lsCadImp, "Constancia de Recepcion", False, 66
'    End If
'
'    Set oPrevio = Nothing
'    Set oPigImpre = Nothing
'    Set oContFunc = Nothing
'    Set oPigGraba = Nothing
'
'    MsgBox "La Recepción de la Valija finalizó satisfactoriamente", vbInformation, "Aviso"
'
'    feGuias.Clear
'    feGuias.Rows = 2
'    feGuias.FormaCabecera
'
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'
'Call CargaCombo(cboMotivo, gColocPigMotivoRem)
'Call CargaCombo(cboDestino, gColocPigUbicacion)
'
'
'End Sub
'
'Private Sub CargaGuias()
'Dim oPigCont As DPigContrato
'Dim rs As Recordset
'
'feGuias.Clear
'feGuias.FormaCabecera
'feGuias.Rows = 2
'
'Set oPigCont = New DPigContrato
'Set rs = oPigCont.dObtieneGuias(2, , cboDestino.BoundText, cboMotivo.BoundText)
'
'Do While Not rs.EOF
'    feGuias.AdicionaFila
'    feGuias.TextMatrix(feGuias.Rows - 1, 1) = feGuias.Rows - 1
'    feGuias.TextMatrix(feGuias.Rows - 1, 3) = rs!cNumDoc
'    feGuias.TextMatrix(feGuias.Rows - 1, 4) = rs!Origen
'    feGuias.TextMatrix(feGuias.Rows - 1, 5) = rs!Motivo
'    feGuias.TextMatrix(feGuias.Rows - 1, 6) = rs!nTotItem
'    feGuias.TextMatrix(feGuias.Rows - 1, 7) = rs!TipoGuia
'    feGuias.TextMatrix(feGuias.Rows - 1, 8) = rs!nMotivo
'    rs.MoveNext
'Loop
'
'Set rs = Nothing
'Set oPigCont = Nothing
'
'End Sub
'
'Private Sub CargaCombo(Combo As DataCombo, ByVal psConsCod As String)
'Dim oPigFunc As DPigFunciones
'Dim rs As Recordset
'
'Set oPigFunc = New DPigFunciones
'
'    Set rs = oPigFunc.GetConstante(psConsCod)
'    Set Combo.RowSource = rs
'    Combo.ListField = "cConsDescripcion"
'    Combo.BoundColumn = "nConsValor"
'
'    Set rs = Nothing
'
'    Set oPigFunc = Nothing
'
'End Sub
'
'Private Function ValidaAgBovedaValores() As Boolean
'Dim oParam As DPigFunciones
'Dim lsAgeBovVal As String
'
'    Set oParam = New DPigFunciones
'        lsAgeBovVal = CStr(oParam.GetParamValor(8040))
'        If Right(lsAgeBovVal, 2) <> Right(gsCodAge, 2) Then
'            ValidaAgBovedaValores = False
'        Else
'            ValidaAgBovedaValores = True
'        End If
'    Set oParam = Nothing
'
'End Function
'
