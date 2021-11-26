VERSION 5.00
Begin VB.Form frmLogSelPaseContra 
   Caption         =   "Pase Contratacion"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11670
   Icon            =   "frmLogSelPaseContra.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6765
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtcotizacion 
      Height          =   375
      Left            =   7320
      TabIndex        =   17
      Top             =   2280
      Width           =   1695
   End
   Begin VB.ComboBox cboperiodo 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox txtDescripcionProveedor 
      Height          =   315
      Left            =   3120
      TabIndex        =   9
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Frame s 
      Caption         =   "Proceso de  Seleccion"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11415
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
         Height          =   555
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   8895
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
         TabIndex        =   1
         Top             =   600
         Width           =   5895
      End
      Begin Sicmact.TxtBuscar txtSeleccionA 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
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
         TipoBusqueda    =   2
         sTitulo         =   ""
      End
      Begin VB.Label Label5 
         Caption         =   "Numero"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   315
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C00000&
         BorderStyle     =   4  'Dash-Dot
         FillColor       =   &H8000000D&
         Height          =   495
         Left            =   7440
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblestado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
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
         Height          =   195
         Left            =   7560
         TabIndex        =   5
         Top             =   480
         Width           =   660
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Proceso"
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
         Height          =   195
         Index           =   7
         Left            =   7440
         TabIndex        =   4
         Top             =   120
         Width           =   1350
      End
   End
   Begin Sicmact.TxtBuscar txtProveedor 
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
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
      sTitulo         =   ""
   End
   Begin Sicmact.FlexEdit fgeBienesConfig 
      Height          =   3495
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   6165
      Cols0           =   9
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "Item-Código-Descripción Bien-Unidad-Valor Unidad-Descripcion Adicional-Cantidad-Precio Ref-Sub Total"
      EncabezadosAnchos=   "450-1200-3500-700-0-2500-1000-1000-1000"
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
      ColumnasAEditar =   "X-1-X-X-X-5-6-7-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-1-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "R-L-L-L-L-L-R-C-R"
      FormatosEdit    =   "0-0-0-0-0-0-3-2-2"
      CantEntero      =   10
      CantDecimales   =   4
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Periodo"
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
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   660
   End
   Begin VB.Label lblproveedor 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   885
   End
End
Attribute VB_Name = "frmLogSelPaseContra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim clsDGnral As DLogGeneral
Dim clsDGAdqui As DLogAdquisi
Dim ClsNAdqui As NActualizaProcesoSelecLog
Dim oCons As DConstantes
Public sAccionSelBienes As String
Dim bpuntaje As Boolean
Dim clsDBS As DLogBieSer
Dim saccion As String
Dim psTposel As String
Dim rs_cta As ADODB.Recordset


Private Sub cboPeriodo_Click()
txtSeleccionA.Text = ""
txttipo.Text = ""
txtdescripcion.Text = ""
'501205 Compra      Soles
'502205 Compra      Dollares
'501207 Servicio    Soles
'502207 Servicio    Dollares
'mid(gcOpeCod,2,1)
Me.txtSeleccionA.rs = clsDGAdqui.LogSeleccionListaConPara(cboperiodo.Text, 4, Mid(gcOpeCod, 3, 1))
End Sub

Private Sub cmdAceptar_Click()
Dim i  As Integer
Dim oCon As DConecta
Set oCon = New DConecta
Dim sSqlO As String
Dim sctaProvision As String
If txtProveedor.Text = "" Then
    MsgBox "Antes debe seleccionar un proveedor", vbInformation, "Seleccione un proveedor"
    Exit Sub
End If

Set rs = clsDGAdqui.CargaSelDetalle(txtSeleccionA.Text, 4, txtProveedor.Text)

If rs.EOF = True Then Exit Sub

i = 1
Set rs_cta = CargaOpeCta(gcOpeCod, "H")
            If rs_cta.EOF Then
                MsgBox "Falta definir Cuenta de Provisión en Operación", vbInformation, "¡Aviso!"
                Exit Sub
            End If
sctaProvision = rs_cta!cCtaContCod
frmLogOCompra.txtCotNro = txtcotizacion.Text
Limpia_Grilla

Do While Not rs.EOF
        frmLogOCompra.fgDetalle.TextMatrix(i, 1) = rs!cBSCod     'Codigo
        frmLogOCompra.fgDetalle.TextMatrix(i, 2) = rs!cBSDescripcion     'Descripción
        frmLogOCompra.fgDetalle.TextMatrix(i, 3) = rs!cConsDescripcion     'unidad
        frmLogOCompra.fgDetalle.TextMatrix(i, 4) = rs!nLogSelCotDetCantidad     'Solicitado
        frmLogOCompra.fgDetalle.TextMatrix(i, 5) = rs!nLogSelCotDetPrecio       'P.Unitario
        frmLogOCompra.fgDetalle.TextMatrix(i, 6) = rs!nLogSelCotDetPrecio       'Saldo
        frmLogOCompra.fgDetalle.TextMatrix(i, 7) = rs!nLogSelCotDetPrecio * rs!nLogSelCotDetCantidad    'SubTotal
        
        sSqlO = frmLogOCompra.FormaSelect(gcOpeCod, rs!cBSCod, 0, "01")
        oCon.AbreConexion
        Set rs_cta = oCon.CargaRecordSet(sSqlO)
            If RSVacio(rs_cta) Then
                MsgBox "Objeto no asignado a Operación  " + rs!cBSDescripcion, vbInformation, "¡Aviso ! Definir Cuentas Contables"
                Limpia_Grilla
                Exit Sub
            End If
            If rs_cta.RecordCount = 1 Then
                frmLogOCompra.fgDetalle.TextMatrix(i, 8) = Trim(rs_cta!cObjetoCod)
            End If
            frmLogOCompra.fgDetalle.TextMatrix(i, 9) = sctaProvision
            frmLogOCompra.fgDetalle.TextMatrix(i, 11) = frmLogOCompra.txtPlazo.value 'plazo de entrega
            i = i + 1
        rs.MoveNext
Loop
frmLogOCompra.txtPersona.Text = txtProveedor.Text
frmLogOCompra.txtProvNom.Text = txtDescripcionProveedor.Text
frmLogOCompra.txtPersona.psCodigoPersona = txtProveedor.Text
frmLogOCompra.txtPersona.psDescripcion = txtDescripcionProveedor.Text

Unload Me

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub Form_Load()
Me.Width = 11790
Me.Height = 7275
Set clsDBS = New DLogBieSer
fgeBienesConfig.BackColorBkg = 16777215
Set rs = New ADODB.Recordset
Set clsDGnral = New DLogGeneral
Set clsDGAdqui = New DLogAdquisi
Set clsDBS = New DLogBieSer
Set rs = clsDGnral.CargaPeriodo
Set ClsNAdqui = New NActualizaProcesoSelecLog
Call CargaCombo(rs, cboperiodo)
ubicar_ano Year(gdFecSis), cboperiodo

Set rs_cta = New ADODB.Recordset

End Sub
Sub ubicar_ano(codigo As String, combo As ComboBox)
Dim i As Integer
For i = 0 To combo.ListCount
If combo.List(i) = codigo Then
    combo.ListIndex = i
    Exit For
    End If
Next
End Sub
Private Sub txtProveedor_EmiteDatos()
txtDescripcionProveedor.Text = txtProveedor.psDescripcion
If txtProveedor.Text = "" Or txtProveedor.Text = "0" Then
          MsgBox "Seleccione el Proveedor", vbInformation, "Seleccione el Proveedor"
          Exit Sub
End If
'obtener el Numero de RUC del Proveedor
Set rs = clsDGAdqui.CargaSelDetalle(txtSeleccionA.Text, 4, txtProveedor.Text)
    If Not rs.EOF = True Then
        Set fgeBienesConfig.Recordset = rs
            fgeBienesConfig.AdicionaFila
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 1) = " --------------------------- "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 2) = " --------------------- TOTAL ------------- "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 3) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 5) = " --------------- " 'fgeBienesConfig.SumaRow(5)
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 6) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 7) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 8) = Format(fgeBienesConfig.SumaRow(8), "########.00")
        Else
        MsgBox "El Proveedor no gano en ningun Bien  ", vbInformation, "El Proveedor no gano en ningun Bien"
        fgeBienesConfig.Clear
        fgeBienesConfig.FormaCabecera
        fgeBienesConfig.Rows = 2
    End If
End Sub
Private Sub txtProveedor_GotFocus()
If txtSeleccionA.Text = "" Then
    txtProveedor.Text = ""
    txtDescripcionProveedor.Text = ""
    Exit Sub
End If
Me.txtProveedor.rs = clsDGAdqui.LogSeleccionListaProveedores(txtSeleccionA.Text)
End Sub
Private Sub txtSelecciona_EmiteDatos()
If txtSeleccionA.Text = "" Then Exit Sub
       fgeBienesConfig.Clear
       fgeBienesConfig.FormaCabecera
       fgeBienesConfig.Rows = 2
       txtProveedor.Text = ""
       txtDescripcionProveedor.Text = ""
       mostrar_descripcion txtSeleccionA.Text
End Sub

Sub Mostrar_Config_Bienes(pnNumseleccion As Long)
'Mostrar la Referencia
    Dim nSuma As Currency
    If txtSeleccion.Text = "" Then Exit Sub
    Set rs = clsDGAdqui.CargaSelReferencia(pnNumseleccion)
    If Not rs.EOF = True Then
        txtperiodo.Text = rs!nLogSelPeriodo
        If rs!nLogSelTpoReq = ReqTipoRegular Then
            txtRequerimiento.Text = "Regular" & Space(10) & rs!nLogSelTpoReq
           ElseIf rs!nLogSelTpoReq = ReqTipoExtemporaneo Then
           txtRequerimiento.Text = "Extemporaneo" & Space(10) & rs!nLogSelTpoReq
        End If
        txtconsolidado.Text = Str(rs!nConsolidado) + " - " + rs!SDescripcionConsol
        txtmesini.Text = rs!nMesIni
        txtmesfin.Text = rs!nMesFin
        txtCategoria.Text = rs!sCategoriaBien
        'Mostrar el Detalle
    End If
    Set rs = clsDGAdqui.CargaSelDetalle(pnNumseleccion, 1)
    If Not rs.EOF = True Then
        Set fgeBienesConfig.Recordset = rs
            fgeBienesConfig.AdicionaFila
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 1) = " --------------------------- "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 2) = " --------------------- TOTAL ------------- "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 3) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 4) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 5) = " --------------- " 'fgeBienesConfig.SumaRow(5)
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 6) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 7) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 8) = Format(fgeBienesConfig.SumaRow(8), "########.00")
        Else
        fgeBienesConfig.Clear
        fgeBienesConfig.FormaCabecera
        fgeBienesConfig.Rows = 2
    End If
End Sub

Sub Mostrar_config_Bienes_Cotiza(pnNumseleccion As Long, psCodproveedor As String)
    Dim nSuma As Currency
    If txtSeleccion.Text = "" Then Exit Sub
    If txtProveedor.Text = "" Or txtProveedor.Text = "0" Then
         MsgBox "Seleccione un Proveedor,Asegurese de que el proceso tenga Proveedores Configurados", vbInformation, "Seleccione un Proveedor"
         Exit Sub
    End If
    txtestado.Text = clsDGAdqui.CargaEstadoProveedor(txtSeleccion.Text, txtProveedor.Text)
    Set rs = clsDGAdqui.CargaSelReferencia(pnNumseleccion)
    If Not rs.EOF = True Then
        txtperiodo.Text = rs!nLogSelPeriodo
        If rs!nLogSelTpoReq = ReqTipoRegular Then
            txtRequerimiento.Text = "Regular" & Space(10) & rs!nLogSelTpoReq
           ElseIf rs!nLogSelTpoReq = ReqTipoExtemporaneo Then
           txtRequerimiento.Text = "Extemporaneo" & Space(10) & rs!nLogSelTpoReq
        End If
        txtconsolidado.Text = Str(rs!nConsolidado) + " - " + rs!SDescripcionConsol
        txtmesini.Text = rs!nMesIni
        txtmesfin.Text = rs!nMesFin
        txtCategoria.Text = rs!sCategoriaBien
        'Mostrar el Detalle
    End If
    Set rs = clsDGAdqui.CargaSelDetalle(pnNumseleccion, 3, psCodproveedor)
    If Not rs.EOF = True Then
        Set fgeBienesConfig.Recordset = rs
            fgeBienesConfig.AdicionaFila
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 1) = " --------------------------- "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 2) = " --------------------- TOTAL ------------- "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 3) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 5) = " --------------- " 'fgeBienesConfig.SumaRow(5)
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 6) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 7) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 8) = Format(fgeBienesConfig.SumaRow(8), "########.00")
        Else
        fgeBienesConfig.Clear
        fgeBienesConfig.FormaCabecera
        fgeBienesConfig.Rows = 2
    End If
End Sub


Private Function ValidaGrilla() As Boolean
    Dim nBs As Integer, nBSMes As Integer, nCant As Integer
    'Validación de BienesServicios
    ValidaGrilla = True
    For nBs = 1 To fgeBienesConfig.Rows - 1
        If fgeBienesConfig.TextMatrix(nBs, 1) = "" Then
            MsgBox "Falta determinar el Bien/Servicio en el Item " & nBs, vbInformation, "Seleccione un Codigo de Bien"
            ValidaGrilla = False
            Exit Function
        End If
       
        If fgeBienesConfig.TextMatrix(nBs, 6) = "" Then
            MsgBox "Falta determinar La Cantidad en el Item " & nBs, vbInformation, "Ingrese la Cantidad en el Item"
            ValidaGrilla = False
            Exit Function
        End If
        If fgeBienesConfig.TextMatrix(nBs, 7) = "" Then
            MsgBox "Falta determinar el Precio Referencial en el Item " & nBs, vbInformation, "Seleccione un Codigo de Bien"
            ValidaGrilla = False
            Exit Function
        End If
    Next
End Function

Sub mostrar_descripcion(nLogSelProceso As Long)
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs = clsDGAdqui.CargaLogSelDescripcionProceso(nLogSelProceso)
If rs.EOF = True Then
    txttipo.Text = ""
    txtdescripcion.Text = ""
    lblestado.Caption = ""
    txtcotizacion.Text = ""
    Else
    txttipo.Text = UCase(rs!cTipo)
    txtcotizacion.Text = rs!nLogSelNumeroCot
    txtdescripcion.Text = "COTIZACION Nº: " + rs!nLogSelNumeroCot + " - " + rs!cDescripcionProceso + " - TIPO PROCESO: " + rs!nLogSelDescProceso + Space(300) + Str(rs!nLogSelTipoProceso)
    'lblcotiza.Caption = rs!nLogSelNumeroCot
    If rs!nLogSelEstado = SelEstProcesoIniciado Then
        lblestado.Caption = "INICIADO"
    ElseIf rs!nLogSelEstado = SelEstProcesoEvaluacionTec Then
        lblestado.Caption = "EVALUACION TECNICA"
    ElseIf rs!nLogSelEstado = SelEstProcesoEvaluacionEco Then
        lblestado.Caption = "EVALUACION ECONOMICA"
    ElseIf rs!nLogSelEstado = SelEstProcesoFinEvaluacion Then
        lblestado.Caption = "FIN DE EVALUACION"
    ElseIf rs!nLogSelEstado = SelEstProcesoCerrado Then
        lblestado.Caption = "CERRADO"
    ElseIf rs!nLogSelEstado = SelEstProcesoCancelado Then
        lblestado.Caption = "CANCELADO"
    End If
End If
End Sub

Public Sub Inicio(ByVal psTipoSel As String, ByVal psFormTpo As String, Optional ByVal psSeleccionNro As String = "")
psTposel = psTipoSel
psFrmTpo = psFormTpo
psReqNro = psSeleccionNro
Me.Show
End Sub


Sub Limpia_Grilla()
Dim i As Integer
i = 1
For i = 1 To frmLogOCompra.fgDetalle.Rows - 1
        frmLogOCompra.fgDetalle.TextMatrix(i, 1) = ""
        frmLogOCompra.fgDetalle.TextMatrix(i, 2) = ""
        frmLogOCompra.fgDetalle.TextMatrix(i, 3) = ""
        frmLogOCompra.fgDetalle.TextMatrix(i, 4) = ""
        frmLogOCompra.fgDetalle.TextMatrix(i, 5) = ""
        frmLogOCompra.fgDetalle.TextMatrix(i, 6) = ""
        frmLogOCompra.fgDetalle.TextMatrix(i, 7) = ""
        frmLogOCompra.fgDetalle.TextMatrix(i, 8) = ""
        
        frmLogOCompra.fgDetalle.TextMatrix(i, 11) = "" 'plazo de entrega
Next


End Sub

