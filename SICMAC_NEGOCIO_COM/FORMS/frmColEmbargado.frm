VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmColEmbargado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Embargo"
   ClientHeight    =   7065
   ClientLeft      =   2805
   ClientTop       =   3360
   ClientWidth     =   11940
   Icon            =   "frmColEmbargado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   11940
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   6240
      Width           =   11775
      Begin VB.CheckBox chkExcel 
         Caption         =   "Excel"
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
         Left            =   1560
         TabIndex        =   34
         Top             =   300
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   10320
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   8880
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   7440
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraInfoCredito 
      Caption         =   "Informacion del Credito"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   11775
      Begin VB.CommandButton cmdBuscarTitularBien 
         Caption         =   "Buscar"
         Height          =   285
         Left            =   10680
         TabIndex        =   41
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscarDepo 
         Caption         =   "Buscar"
         Height          =   285
         Left            =   4800
         TabIndex        =   38
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtDocSalida 
         Height          =   285
         Left            =   10200
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtExpediente 
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtResolucion 
         Height          =   285
         Left            =   4080
         TabIndex        =   16
         Top             =   615
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   300
         Left            =   7080
         TabIndex        =   22
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskSalida 
         Height          =   300
         Left            =   10200
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label txtTitularBien 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7080
         TabIndex        =   40
         Top             =   960
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label lblTitularBien 
         Caption         =   "Titular Bien:"
         Height          =   255
         Left            =   6000
         TabIndex        =   39
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label txtDepositario 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   37
         Top             =   960
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label lblDepositario 
         Caption         =   "Depositario:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblDocSalida 
         Caption         =   "Doc Salida:"
         Height          =   255
         Left            =   9120
         TabIndex        =   30
         Top             =   255
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblSalida 
         Caption         =   "Fec Salida:"
         Height          =   255
         Left            =   9120
         TabIndex        =   25
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Fec Embargo:"
         Height          =   255
         Left            =   6000
         TabIndex        =   23
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Nº Expediente"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   615
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Nº Resolucion"
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblGasto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7080
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Gastos:"
         Height          =   255
         Left            =   6000
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblInteres 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Interes:"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblCapital 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Capital:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   3960
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   3615
         _extentx        =   6376
         _extenty        =   661
         texto           =   "Crédito"
         enabledcta      =   -1
         enabledprod     =   -1
         enabledage      =   -1
         cmac            =   "109"
      End
      Begin VB.Label lbltitular 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4080
         TabIndex        =   2
         Top             =   720
         Width           =   5895
      End
      Begin VB.Label Label4 
         Caption         =   "Titular Cred:"
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblPersCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame fraBien 
      Caption         =   "Bienes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3495
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   11775
      Begin SICMACT.FlexEdit fgBienes 
         Height          =   2655
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   11535
         _extentx        =   20346
         _extenty        =   4683
         cols0           =   21
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   $"frmColEmbargado.frx":030A
         encabezadosanchos=   "400-0-1100-1100-1250-1250-1200-800-1200-1200-800-1200-1200-1200-1200-1200-1200-1200-1200-0-0"
         font            =   "frmColEmbargado.frx":03B4
         font            =   "frmColEmbargado.frx":03E0
         font            =   "frmColEmbargado.frx":040C
         font            =   "frmColEmbargado.frx":0438
         font            =   "frmColEmbargado.frx":0464
         fontfixed       =   "frmColEmbargado.frx":0490
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-L-L-L-L-L-C-L-L-L-L-L-L-L-L-L-L-R-L-L"
         formatosedit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         textarray0      =   "Nº"
         colwidth0       =   405
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   960
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdAgregarBien 
         Caption         =   "Agregar Bien"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar Bien"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   21
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdEliminarBien 
         Caption         =   "Eliminar Bien"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   20
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox cmbTipoSalida 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3030
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblTipoSalida 
         Caption         =   "Tipo Salida:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3060
         Visible         =   0   'False
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmColEmbargado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim lsCtaCod As String
'Dim lnMovNro As Long
'Dim lnOpcion As Integer
'Dim lbresultado As Boolean
'
'Dim lnNroCalen As Integer
'Dim lnNroCuota As Integer
'Dim lnDiasAtraso As Integer
'Dim lsMetLiqui As String
'Dim lnPlazo As Integer
'Dim lnPrdEstado As Integer
'
'Dim lsCodDepositario As String
'Dim lsCodTitBien As String
'Dim lsDniDepositario, lsDniTitular
'
'Dim lnInteresComp As Currency
'Dim lnInteresMora As Currency
'
'Const consNuevoEmbargo = 0
'Const consSalidaEmbargo = 1
'Const consModificarEmbargo = 2
'Const consConsultarEmbargo = 3
'Public Function Inicio(Optional pnMovNro As Long = 0, Optional psCtaCod As String = "", Optional pnOpcion As Integer = 0) As Boolean
'    lnMovNro = pnMovNro
'    lsCtaCod = psCtaCod
'    lnOpcion = pnOpcion
'    Me.Show 1
'    Inicio = lbresultado
'End Function
'Private Sub chkTodos_Click()
'    Dim i As Integer
'    For i = 1 To fgBienes.Rows - 1
'        fgBienes.TextMatrix(i, 1) = Me.chkTodos.value
'    Next i
'End Sub
'Private Sub cmbTipoSalida_Click()
'    If Right(Me.cmbTipoSalida.Text, 1) <> "2" Then
'        Me.lblDepositario.Visible = True
'        Me.txtDepositario.Visible = True
'        Me.cmdBuscarDepo.Visible = True
'        Me.lblTitularBien.Visible = True
'        Me.txtTitularBien.Visible = True
'        Me.cmdBuscarTitularBien.Visible = True
'        Me.txtTitularBien = ""
'        Me.txtDepositario = ""
'
'         lsCodDepositario = ""
'         lsCodTitBien = ""
'    Else
'        Me.lblDepositario.Visible = False
'        Me.txtDepositario.Visible = False
'        Me.cmdBuscarDepo.Visible = False
'        Me.lblTitularBien.Visible = False
'        Me.txtTitularBien.Visible = False
'        Me.cmdBuscarTitularBien.Visible = False
'
'    End If
'End Sub
'
'Private Sub cmdBuscar_Click()
'    Dim loPers As COMDPersona.UCOMPersona 'UPersona
'    Dim lsPersCod As String, lsPersNombre As String
'    Dim lsEstados As String
'    Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
'    Dim lrCreditos As New ADODB.Recordset
'    Dim loCuentas As COMDPersona.UCOMProdPersona
'
'On Error GoTo ControlError
'
'    Set loPers = New COMDPersona.UCOMPersona
'    Set loPers = frmBuscaPersona.Inicio
'    If loPers Is Nothing Then Exit Sub
'    lsPersCod = loPers.sPersCod
'    lsPersNombre = loPers.sPersNombre
'    Set loPers = Nothing
'
'    ' Selecciona Estados
'    lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast
'
'    If Trim(lsPersCod) <> "" Then
'        Set loPersCredito = New COMDColocRec.DCOMColRecCredito
'            Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
'        Set loPersCredito = Nothing
'    End If
'
'    Set loCuentas = New COMDPersona.UCOMProdPersona
'        Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
'        If loCuentas.sCtaCod <> "" Then
'            AXCodCta.Enabled = True
'            AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
'            AXCodCta.SetFocusCuenta
'        End If
'    Set loCuentas = Nothing
'
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
'        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'End Sub
'
'Private Sub cmdBuscarDepo_Click()
'    Dim loPers As COMDPersona.UCOMPersona
'
'On Error GoTo ControlError
'
'    Set loPers = New COMDPersona.UCOMPersona
'    Set loPers = frmBuscaPersona.Inicio
'    If loPers Is Nothing Then Exit Sub
'    Me.txtDepositario.BackColor = vbWhite
'    lsCodDepositario = loPers.sPersCod
'    Me.txtDepositario = loPers.sPersNombre
'    lsDniDepositario = loPers.sPersIdnroDNI
'    Set loPers = Nothing
'
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
'        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'End Sub
'
'Private Sub cmdBuscarTitularBien_Click()
'     Dim loPers As COMDPersona.UCOMPersona
'
'On Error GoTo ControlError
'
'    Set loPers = New COMDPersona.UCOMPersona
'    Set loPers = frmBuscaPersona.Inicio
'    If loPers Is Nothing Then Exit Sub
'    Me.txtTitularBien.BackColor = vbWhite
'    lsCodTitBien = loPers.sPersCod
'    Me.txtTitularBien = loPers.sPersNombre
'    lsDniTitular = loPers.sPersIdnroDNI
'
'    Set loPers = Nothing
'
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
'        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'End Sub
'
'Private Sub cmdImprimir_Click()
'   If Me.chkExcel.value = 0 Then
'        ImprimirEmbargo
'   Else
'        ImprimirExcel
'   End If
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    lbresultado = False
'    If lnMovNro <> 0 And lsCtaCod <> "" Then
'        Me.AXCodCta.NroCuenta = lsCtaCod
'        ObtieneDatosCuenta lsCtaCod
'        HabilitarControles
'        obtenerDatosEmbargo
'        'HabilitarControles
'    End If
'End Sub
'Private Sub obtenerDatosEmbargo()
'    Dim oColRec As COMNColocRec.NCOMColRecCredito
'    Dim rsEmbargo As ADODB.Recordset
'    Dim rsEmbargoDetalle As ADODB.Recordset
'
'    Set oColRec = New COMNColocRec.NCOMColRecCredito
'    Set rsEmbargo = New ADODB.Recordset
'    Set rsEmbargoDetalle = New ADODB.Recordset
'
'    Set rsEmbargo = oColRec.obtenerEmbargo(lsCtaCod)
'    If Not (rsEmbargo.EOF And rsEmbargo.BOF) Then
'        Me.txtExpediente.Text = rsEmbargo!cNroExpediente
'        Me.txtResolucion.Text = rsEmbargo!cNroResolucion
'        Me.mskFecha.Text = rsEmbargo!dFecEmbargo
'        Set rsEmbargoDetalle = oColRec.obtenerEmbargoDetalle(lsCtaCod)
'        If Not (rsEmbargoDetalle.EOF And rsEmbargoDetalle.BOF) Then
'            Set fgBienes.Recordset = rsEmbargoDetalle
'        End If
'    End If
'
'End Sub
'Private Sub HabilitarControles()
'    If lnOpcion = consNuevoEmbargo Or lnOpcion = consModificarEmbargo Then
'        If lnOpcion = consModificarEmbargo Then
'            Me.Caption = "Modificar Embargo"
'        End If
'        Me.fgBienes.EncabezadosNombres = "Nº-OK-Estado-Fecha-Tpo_Bien-Sub_Tpo_Bien-Bien-Cant-Descripcion-Color-Marca-Modelo-Serie-Motor-Almacen-Estado_Bien-Placa-Part_Elect-Tasacion-Fec_Tasacion-Mone_Tasacion"
'        Me.fgBienes.EncabezadosAnchos = "400-0-1100-1100-1250-1250-1200-800-1200-1200-800-1200-1200-1200-1200-1200-1200-1200-1200"
'    ElseIf lnOpcion = consConsultarEmbargo Then
'        Me.Caption = "Consulta de Embargo"
'        Me.fgBienes.EncabezadosNombres = "Nº-OK-Estado-Fecha-Tpo_Bien-Sub_Tpo_Bien-Bien-Cant-Descripcion-Color-Marca-Modelo-Serie-Motor-Almacen-Estado_Bien-Placa-Part_Elect-Tasacion-Fec_Tasacion-Mone_Tasacion"
'        Me.fgBienes.EncabezadosAnchos = "400-0-1100-1100-1250-1250-1200-800-1200-1200-800-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200"
'        Me.txtExpediente.Enabled = False
'        Me.txtResolucion.Enabled = False
'        Me.cmdAgregarBien.Enabled = False
'        Me.cmdGrabar.Enabled = False
'        Me.mskFecha.Enabled = False
'        Me.AXCodCta.Enabled = False
'        Me.cmdImprimir.Visible = True
'        Me.chkExcel.Visible = True
'    ElseIf lnOpcion = consSalidaEmbargo Then
'        Me.Caption = "Salida de Embargo"
'        Me.fgBienes.EncabezadosAnchos = "400-300-1100-1100-1250-1250-1200-800-1200-1200-800-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200"
'        Me.fgBienes.ColumnasAEditar = "X-1-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
'        Me.fgBienes.ListaControles = "0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
'        Me.cmdAgregarBien.Visible = False
'        Me.cmdEliminarBien.Visible = False
'        Me.cmdModificar.Visible = False
'
'        Me.AXCodCta.Enabled = False
'        Me.txtExpediente.Enabled = False
'        Me.txtResolucion.Enabled = False
'        Me.mskFecha.Enabled = False
'        Me.cmdAgregarBien.Enabled = False
'
'
'        Me.lblSalida.Visible = True
'        Me.lblTipoSalida.Visible = True
'        Me.lblDocSalida.Visible = True
'        Me.mskSalida.Visible = True
'        Me.cmbTipoSalida.Visible = True
'        Me.txtDocSalida.Visible = True
'        Me.chkTodos.Visible = True
'
'
'        If Me.txtDocSalida.Visible = True Then
'            Me.txtDocSalida.SetFocus
'        End If
'        cargarCombo
'    End If
'End Sub
'Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If Not verificarEmbargo Then
'            ObtieneDatosCuenta AXCodCta.NroCuenta
'
'        Else
'            MsgBox "La Cuenta ya se Encuentra Registrada en la Lista de Embargos", vbInformation, "AVISO"
'        End If
'    End If
'End Sub
'Private Function verificarEmbargo() As Boolean
'    Dim oColRec As COMNColocRec.NCOMColRecCredito
'    Dim rsCta As ADODB.Recordset
'
'    Set oColRec = New COMNColocRec.NCOMColRecCredito
'    Set rsCta = New Recordset
'    Set rsCta = oColRec.verificarEmbargo(Me.AXCodCta.NroCuenta)
'    If Not (rsCta.EOF And rsCta.BOF) Then
'        verificarEmbargo = True
'    Else
'        verificarEmbargo = False
'    End If
'
'End Function
'Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
'    Dim oColRec As COMNColocRec.NCOMColRecCredito
'    Dim rsCta As ADODB.Recordset
'    Dim nCta As String
'    Set oColRec = New COMNColocRec.NCOMColRecCredito
'    Set rsCta = New ADODB.Recordset
'
'    Set rsCta = oColRec.ObtenerCuentaEmbargo(sCuenta)
'
'    If Not (rsCta.EOF And rsCta.BOF) Then
'        nCta = Me.AXCodCta.NroCuenta
'        LimpiarPantalla
'        Me.AXCodCta.NroCuenta = nCta
'        lblPersCod.Caption = rsCta("cPersCod")
'        lbltitular.Caption = UCase(PstaNombre(rsCta("Nombre")))
'        Me.lblCapital = Format(rsCta("Capital"), "##,##0.00")
'        Me.lblInteres = Format(rsCta("Intereses"), "##,##0.00")
'        Me.lblGasto = Format(rsCta("gastos"), "##,##0.00")
'
'        lnInteresComp = rsCta("nSaldoIntComp")
'        lnInteresMora = rsCta("nSaldoIntMor")
'
'        lnNroCalen = rsCta("nNroCalen")
'        lnNroCuota = rsCta("nNroProxCuota")
'        lnDiasAtraso = rsCta("nDiasAtraso")
'        lsMetLiqui = rsCta("cMetLiquidacion")
'        lnPlazo = rsCta("nPlazo")
'        lnPrdEstado = rsCta("nPrdEstado")
'
'        Me.fraInfoCredito.Enabled = True
'        Me.fgBienes.Enabled = True
'        Me.fgBienes.lbEditarFlex = True
'        Me.cmdAgregarBien.Enabled = True
'        If lnMovNro = 0 Then
'            Me.txtExpediente.SetFocus
'        End If
'    Else
'        MsgBox "No se ha encontrado información de la cuenta ingresada"
'        AXCodCta.SetFocus
'    End If
'End Sub
'Private Sub LimpiarPantalla()
' If lnOpcion = consNuevoEmbargo Or lnOpcion = consModificarEmbargo Then
'    Me.lblPersCod = ""
'    Me.lbltitular = ""
'    Me.lblCapital = ""
'    Me.lblInteres = ""
'    Me.lblGasto = ""
'    Me.txtExpediente.Text = ""
'    Me.txtResolucion.Text = ""
'    If lnOpcion = consNuevoEmbargo Then
'       Call LimpiaFlex(Me.fgBienes)
'    End If
'    Me.cmdAgregarBien.Enabled = False
'    Me.cmdModificar.Enabled = False
'    Me.cmdEliminarBien.Enabled = False
'    Me.AXCodCta.NroCuenta = "109"
'    Me.mskFecha.Text = "__/__/____"
' ElseIf lnOpcion = consSalidaEmbargo Then
'    Me.txtDocSalida.Text = ""
'    Me.mskSalida.Text = "__/__/____"
'    Me.cmbTipoSalida.ListIndex = -1
'    Me.txtDepositario = ""
'    Me.txtTitularBien = ""
'    chkTodos.value = 0
'    chkTodos_Click
'
' End If
'End Sub
'
'Private Sub cmdAgregarBien_Click()
'   cmdModificar.Enabled = False
'   AgregarBien (frmColEmbargoBien.Inicio)
'
'End Sub
'Private Sub AgregarBien(ByVal pbBien As Boolean, Optional ByVal pnFila As Integer = 0)
'    If pbBien Then
'            If pnFila = 0 Then
'                fgBienes.AdicionaFila
'            End If
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 1) = 0
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 2) = IIf(lnMovNro = 0 Or fgBienes.TextMatrix(fgBienes.row, 2) = "", "Embargado" + Space(100) + "1", fgBienes.TextMatrix(fgBienes.row, 2))
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 3) = IIf(lnOpcion = consModificarEmbargo And fgBienes.TextMatrix(fgBienes.row, 3) <> "", fgBienes.TextMatrix(fgBienes.row, 3), Me.mskFecha.Text) 'IIf(lnMovNro = 0, Me.mskFecha.Text, "01/01/1900")
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 4) = frmColEmbargoBien.TpoBien
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 5) = frmColEmbargoBien.SubTpoBien
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 6) = frmColEmbargoBien.Bien
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 7) = frmColEmbargoBien.Cantidad
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 8) = frmColEmbargoBien.Descripcion
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 9) = frmColEmbargoBien.Color
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 10) = frmColEmbargoBien.Marca
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 11) = frmColEmbargoBien.Modelo
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 12) = frmColEmbargoBien.Serie
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 13) = frmColEmbargoBien.Motor
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 14) = frmColEmbargoBien.Almacen
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 15) = frmColEmbargoBien.Estado
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 16) = frmColEmbargoBien.Placa
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 17) = frmColEmbargoBien.Partida
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 18) = frmColEmbargoBien.Tasacion
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 19) = frmColEmbargoBien.FecTasacion
'            fgBienes.TextMatrix(IIf(pnFila = 0, fgBienes.Rows - 1, fgBienes.row), 20) = frmColEmbargoBien.MonTasacion
'
'   End If
'End Sub
'Private Sub cmdCancelar_Click()
'    LimpiarPantalla
'End Sub
'
'Private Sub cmdEliminarBien_Click()
'        fgBienes.EliminaFila fgBienes.row
'        Me.cmdModificar.Enabled = False
'End Sub
'
'Private Sub cmdGrabar_Click()
'    If Not validaDatos Then
'        Exit Sub
'    End If
'
'    If MsgBox("Seguro de Registrar los Datos", vbYesNo, "Aviso") = vbYes Then
'
'        Dim oColRec As COMNColocRec.NCOMColRecCredito
'        Dim oBase As COMDCredito.DCOMCredActBD
'        Dim clsMant As COMDCaptaGenerales.DCOMCaptaMovimiento
'        Dim clsMov As COMNContabilidad.NCOMContFunciones
'        Dim sMovNro As String
'        Dim nMovNro As Long
'        Dim nNroFilas As Integer
'        Dim i As Integer
'        Set clsMov = New COMNContabilidad.NCOMContFunciones
'        Set oColRec = New COMNColocRec.NCOMColRecCredito
'        Set clsMant = New COMDCaptaGenerales.DCOMCaptaMovimiento
'        Set oBase = New COMDCredito.DCOMCredActBD
'        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'
'        If lnOpcion = consNuevoEmbargo Then
'            clsMant.AgregaMov sMovNro, 136302, "BIENES EMBARGADOS", gMovEstContabMovContable, gMovFlagVigente
'            nMovNro = clsMant.GetnMovNro(sMovNro)
'            'Call oBase.dInsertMovCol(nMovNro, "136302", Me.AXCodCta.NroCuenta, lnNroCalen, CCur(Me.lblCapital) + CCur(Me.lblInteres) + CCur(Me.lblGasto), lnDiasAtraso, lsMetLiqui, lnPlazo, Me.lblCapital, lnPrdEstado, False)
'            Call oBase.dInsertMovCol(nMovNro, "136302", Me.AXCodCta.NroCuenta, lnNroCalen, CCur(Me.lblCapital), lnDiasAtraso, lsMetLiqui, lnPlazo, Me.lblCapital, lnPrdEstado, False)
'            Call oBase.dInsertMovColDet(nMovNro, "136302", Me.AXCodCta.NroCuenta, CLng(lnNroCalen), 3000, lnNroCuota, CCur(Me.lblCapital), False)
'
'            oColRec.guardarEmbargo nMovNro, Me.AXCodCta.NroCuenta, Me.lblCapital, Me.lblInteres, Me.lblGasto, Me.txtExpediente, Me.txtResolucion, Me.mskFecha.Text, sMovNro
'
'        ElseIf lnOpcion = consModificarEmbargo Then
'            If lnMovNro <> 0 And lsCtaCod <> "" Then
'                oColRec.modificarEmbargo lsCtaCod
'            End If
'            oColRec.guardarEmbargo lnMovNro, Me.AXCodCta.NroCuenta, Me.lblCapital, Me.lblInteres, Me.lblGasto, Me.txtExpediente, Me.txtResolucion, Me.mskFecha.Text, sMovNro
'
'        ElseIf lnOpcion = consSalidaEmbargo Then
'            'Si la opcion es Adjudicar debe valorizar los bienes
'            nNroFilas = obtenerNroFilas
'            If nNroFilas = 0 Then
'                MsgBox "Debe seleccionar un Bien en estado Embargado", vbInformation, "AVISO"
'                Exit Sub
'            End If
'            If Right(Me.cmbTipoSalida.Text, 1) = "2" Then
'                MsgBox "Debe Valorizar los Bienes a Adjudicar", vbOKOnly, "Valorizar Bienes"
'
'                If Not frmColEmbargoAdjudicaBienes.Inicio(Me.AXCodCta.NroCuenta, Me.fgBienes, Me.mskSalida.Text, sMovNro, nNroFilas) Then
'                    MsgBox "Se ha Cancelado la Adjudicacion", vbInformation, "AVISO"
'                    Exit Sub
'                End If
'            End If
'
'        End If
'
'
'        With Me.fgBienes
'            For i = 1 To Me.fgBienes.Rows - 1
'                If lnOpcion = consNuevoEmbargo Or lnOpcion = consModificarEmbargo Then   'Nuevo Bien o Modificar
'
'                oColRec.guardarEmbargoDetalle Me.AXCodCta.NroCuenta, i, IIf(Me.txtDocSalida.Visible = True, Me.txtDocSalida.Text, ""), CDate(Me.mskFecha.Text), _
'                                              Right(.TextMatrix(i, 5), 5), Right(.TextMatrix(i, 6), 4), Right(.TextMatrix(i, 15), 2), _
'                                              CInt(Right(.TextMatrix(i, 2), 1)), Right(.TextMatrix(i, 14), 5), .TextMatrix(i, 8), .TextMatrix(i, 7), .TextMatrix(i, 9), _
'                                              .TextMatrix(i, 10), .TextMatrix(i, 11), .TextMatrix(i, 12), .TextMatrix(i, 13), IIf(.TextMatrix(i, 18) = "", 0, .TextMatrix(i, 18)), .TextMatrix(i, 16), .TextMatrix(i, 17), sMovNro, _
'                                              IIf(Right(.TextMatrix(i, 2), 1) <> 1, .TextMatrix(i, 3), ""), IIf(.TextMatrix(i, 19) <> "", .TextMatrix(i, 19), ""), IIf(.TextMatrix(i, 20) <> "", Right(.TextMatrix(i, 20), 1), "")
'
'                ElseIf lnOpcion = consSalidaEmbargo Then  'Salida del Bien
'
'
'                    If .TextMatrix(i, 1) = "." Then ' Guardar Salida
'                        oColRec.modificarEmbargo lsCtaCod, i
'                        oColRec.guardarEmbargoDetalle Me.AXCodCta.NroCuenta, i, IIf(Me.txtDocSalida.Visible = True, Me.txtDocSalida.Text, ""), CDate(Me.mskFecha.Text), _
'                                                  Right(.TextMatrix(i, 5), 5), Right(.TextMatrix(i, 6), 4), Right(.TextMatrix(i, 15), 2), _
'                                                  CInt(Right(Me.cmbTipoSalida.Text, 1)), Right(.TextMatrix(i, 14), 5), .TextMatrix(i, 8), .TextMatrix(i, 7), .TextMatrix(i, 9), _
'                                                  .TextMatrix(i, 10), .TextMatrix(i, 11), .TextMatrix(i, 12), .TextMatrix(i, 13), IIf(.TextMatrix(i, 18) = "", 0, .TextMatrix(i, 18)), .TextMatrix(i, 16), .TextMatrix(i, 17), sMovNro, _
'                                                  Me.mskSalida.Text, IIf(.TextMatrix(i, 19) <> "", .TextMatrix(i, 19), ""), IIf(.TextMatrix(i, 20) <> "", Right(.TextMatrix(i, 20), 1), ""), _
'                                                  lsCodDepositario, lsCodTitBien
'
'                    End If
'                End If
'
'            Next i
'            If lnOpcion = consSalidaEmbargo Then 'Se guarda en MOV si la CTA no tiene Bienes Embargados
'                Dim nNroBienes As Integer
'                nNroBienes = oColRec.ObtenerNroBienesEmbargados(Me.AXCodCta.NroCuenta)
'                If nNroBienes = 0 Then
'                    clsMant.AgregaMov sMovNro, 136303, "SALIDA DE EMBARGO", gMovEstContabMovContable, gMovFlagVigente
'                    nMovNro = clsMant.GetnMovNro(sMovNro)
'                    Call oBase.dInsertMovCol(nMovNro, "136303", Me.AXCodCta.NroCuenta, lnNroCalen, CCur(Me.lblCapital), lnDiasAtraso, lsMetLiqui, lnPlazo, Me.lblCapital, lnPrdEstado, False)
'                    Call oBase.dInsertMovColDet(nMovNro, "136303", Me.AXCodCta.NroCuenta, CLng(lnNroCalen), 3000, lnNroCuota, CCur(Me.lblCapital), False)
'                End If
'            End If
'        End With
'
'
'
'        lbresultado = True
'        MsgBox "Se han Guardado los Datos", vbInformation, "Aviso"
'
'        If lnOpcion = consNuevoEmbargo Or lnOpcion = consModificarEmbargo Then   'Nuevo Bien o Modificar
'            ImprimirEmbargo
'            If MsgBox("Desea Guardar un Archivo Excel", vbYesNo, "AVISO") = vbYes Then
'                ImprimirExcel
'            End If
'        Else
'            If Right(Me.cmbTipoSalida.Text, 1) <> "2" Then
'                ImprimirSalidaEmbargo
'            Else
'                ImprimirEmbargo
'            End If
'        End If
'        Unload Me
'    End If
'
'End Sub
'Private Function obtenerNroFilas() As Integer
'    Dim i As Integer
'    obtenerNroFilas = 0
'    For i = 1 To Me.fgBienes.Rows - 1
'        If fgBienes.TextMatrix(i, 1) = "." Then
'            If Right(fgBienes.TextMatrix(i, 2), 1) = "1" Then
'            obtenerNroFilas = obtenerNroFilas + 1
'            End If
'        End If
'    Next i
'End Function
'Private Sub ImprimirEmbargo()
'    Dim sCadImp As String
'    Dim lsCabe01 As String
'    Dim lsCabe02 As String
'    Dim lsCabe03 As String
'    Dim lsCabe04 As String
'
'    Dim oPrevio As previo.clsprevio
'    Dim oImpre As COMFunciones.FCOMImpresion
'    Dim oImp As COMFunciones.FCOMVarImpresion
'
'
'  Set oImpre = New COMFunciones.FCOMImpresion
'  Set oImp = New COMFunciones.FCOMVarImpresion
'  oImp.Inicia gEPSON
'
'  sCadImp = sCadImp & Chr(10)
'   ' Cabecera 1
'  sCadImp = oImpre.FillText(Trim(UCase(gsNomCmac)), 45, " ")
'  sCadImp = sCadImp & Space(105 - 45 - 25)
'  sCadImp = sCadImp & "Pag.  : " & str(1) & "  -  " & gsCodUser & Chr(10)
'  sCadImp = sCadImp & oImpre.FillText(Trim(UCase(gsNomAge)), 25, " ")
'  sCadImp = sCadImp & Space(105 - 25 - 25)
'  sCadImp = sCadImp & "Fecha : " & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & Chr(10)
'
'  ' Titulo
'  If lnOpcion <> consSalidaEmbargo Then
'    sCadImp = sCadImp & Centra("BIENES EMBARGADOS", 105) & Chr(10)
'  ElseIf lnOpcion = consSalidaEmbargo Then
'    sCadImp = sCadImp & Centra("SALIDA BIENES EMBARGADOS", 105) & Chr(10)
'  End If
'  sCadImp = sCadImp & Centra(String(26, "="), 105) & Chr(10)
'  'INFORMACION
'  sCadImp = sCadImp & Space(3) & "CLIENTE        :" & ImpreFormat(Me.lbltitular, 20) & Chr(10)
'  sCadImp = sCadImp & Space(3) & "CUENTA         :" & ImpreFormat(Me.AXCodCta.NroCuenta, 20) & Chr(10)
'
'  sCadImp = sCadImp & Space(3) & "Nro EXPEDIENTE :" & ImpreFormat(Me.txtExpediente.Text, 20)
'  sCadImp = sCadImp & ImpreFormat("Nro RESOLUCION :", 17, 0) & ImpreFormat(Me.txtResolucion.Text, 15, 0)
'  sCadImp = sCadImp & ImpreFormat("FEC. EMBARGO:", 14, 0) & ImpreFormat(Me.mskFecha.Text, 15, 0) & Chr(10)
'
'  sCadImp = sCadImp & Space(3) & "CAPITAL        :" & ImpreFormat(Format(Me.lblCapital, "##,##0.00"), 20)
'  sCadImp = sCadImp & ImpreFormat("INTERES        :", 17, 0) & ImpreFormat(Format(Me.lblInteres, "##,##0.00"), 15, 0)
'  sCadImp = sCadImp & ImpreFormat("GASTOS      :", 14, 0) & ImpreFormat(Format(Me.lblGasto, "##,##0.00"), 15, 0) & Chr(10)
'
'  If lnOpcion = consSalidaEmbargo Then
'  sCadImp = sCadImp & Space(3) & "TPO SALIDA     :" & ImpreFormat(Trim(Left(Me.cmbTipoSalida.Text, 20)), 20)
'  sCadImp = sCadImp & ImpreFormat("DOC. SALIDA    :", 17, 0) & ImpreFormat(Me.txtDocSalida.Text, 15, 0)
'  sCadImp = sCadImp & ImpreFormat("FEC. SALIDA :", 14, 0) & ImpreFormat(Me.mskSalida.Text, 15, 0) & Chr(10)
'  End If
'
'  sCadImp = sCadImp & Chr(10)
'  sCadImp = sCadImp & String(105, "-") & Chr(10)
'  sCadImp = sCadImp & Chr(10)
'
'  sCadImp = sCadImp & "Nro    BIEN           ESTADO   CANT   ALMACEN                   DESCRIPCION"
'
'  sCadImp = sCadImp & Chr(10)
'  sCadImp = sCadImp & String(105, "-") & Chr(10)
'  sCadImp = sCadImp & Chr(10)
'
'  Dim i As Integer
'  Dim nNroItem As Integer
'  nNroItem = 0
'  For i = 1 To Me.fgBienes.Rows - 1
'
'    If lnOpcion <> consSalidaEmbargo Then
'        If Right(Me.fgBienes.TextMatrix(i, 2), 1) = 1 Then
'            nNroItem = nNroItem + 1
'            sCadImp = sCadImp & ImpreFormat(nNroItem, 3, 0)
'            sCadImp = sCadImp & ImpreFormat(Trim(Left(Me.fgBienes.TextMatrix(i, 6), 50)), 15) 'BIEN
'            sCadImp = sCadImp & ImpreFormat(Trim(Left(Me.fgBienes.TextMatrix(i, 15), 50)), 8) 'CONDICION
'            sCadImp = sCadImp & ImpreFormat(Trim(Left(Me.fgBienes.TextMatrix(i, 7), 50)), 4) 'CANTIDAD
'            sCadImp = sCadImp & ImpreFormat(Trim(Left(Me.fgBienes.TextMatrix(i, 14), 50)), 23, 0) 'ALMACEN
'            sCadImp = sCadImp & ImpreFormat(Trim(Me.fgBienes.TextMatrix(i, 8)), 40) & Chr(10) 'DESCRIPCION
'        End If
'    ElseIf lnOpcion = consSalidaEmbargo Then
'        If fgBienes.TextMatrix(i, 1) = "." Then
'            nNroItem = nNroItem + 1
'            sCadImp = sCadImp & ImpreFormat(nNroItem, 3, 0)
'            sCadImp = sCadImp & ImpreFormat(Trim(Left(Me.fgBienes.TextMatrix(i, 6), 50)), 15) 'BIEN
'            sCadImp = sCadImp & ImpreFormat(Trim(Left(Me.fgBienes.TextMatrix(i, 15), 50)), 8) 'CONDICION
'            sCadImp = sCadImp & ImpreFormat(Trim(Left(Me.fgBienes.TextMatrix(i, 7), 50)), 4) 'CANTIDAD
'            sCadImp = sCadImp & ImpreFormat(Trim(Left(Me.fgBienes.TextMatrix(i, 14), 50)), 23, 0) 'ALMACEN
'            sCadImp = sCadImp & ImpreFormat(Trim(Me.fgBienes.TextMatrix(i, 8)), 40) & Chr(10) 'DESCRIPCION
'        End If
'    End If
'  Next i
'
'
'    Set oPrevio = New previo.clsprevio
'    oPrevio.Show oImpresora.gPrnTpoLetraSansSerif1PDef & oImpresora.gPrnTamLetra10CPIDef & sCadImp, "BIENES EMBARGADOS"
'    Set oPrevio = Nothing
'End Sub
'Private Sub ImprimirSalidaEmbargo()
'    Dim sCadImp As String
'    Dim sCebecera As String
'    Dim FirmaDni As String
'    Dim sFirmaDni As String
'    Dim sFirmaNombre As String
'    Dim oColRec As COMNColocRec.NCOMColRecCredito
'    Dim sCiudad As String
'    Dim i As Integer
'    Dim nNroItem As Integer
'
'
'    Dim oWord As Word.Application
'    Dim oDoc As Word.Document
'    Dim oRange As Word.Range
'
'    Set oColRec = New COMNColocRec.NCOMColRecCredito
'    Set oWord = CreateObject("Word.Application")
'    'oWord.Visible = True
'
'
'    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\ACTA_SALIDA_EMBARGOS.doc")
'     With oWord.Selection.Find
'        .Text = "<<Hora>>"
'        .Replacement.Text = Format(GetHoraServer(), "hh:mm ") & IIf(CInt(Mid(GetHoraServer(), 1, 2)) >= 12, " pm", " am")
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'    With oWord.Selection.Find
'        .Text = "<<Dia>>"
'        .Replacement.Text = Format(Day(gdFecSis), "00")
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'     With oWord.Selection.Find
'        .Text = "<<Mes>>"
'        .Replacement.Text = Format(gdFecSis, "MMMM")
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    With oWord.Selection.Find
'        .Text = "<<Anio>>"
'        .Replacement.Text = Year(gdFecSis)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'    With oWord.Selection.Find
'        .Text = "<<Depositario>>"
'        .Replacement.Text = PstaNombre(Trim(Left(Me.txtDepositario, 70)))
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'    With oWord.Selection.Find
'        .Text = "<< DniDepositario >>"
'        .Replacement.Text = "DNI: " & lsDniDepositario
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'    With oWord.Selection.Find
'        .Text = "<<Expediente>>"
'        .Replacement.Text = Me.txtExpediente
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'    With oWord.Selection.Find
'        .Text = "<<Credito>>"
'        .Replacement.Text = Me.AXCodCta.NroCuenta
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'
'    sCiudad = oColRec.ObtenerAgenciaCiudad(gsCodAge)
'    With oWord.Selection.Find
'        .Text = "<<Ciudad>>"
'        .Replacement.Text = sCiudad
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'    sFirmaDni = "DNI: " & lsDniDepositario + Space(80) + "DNI: " & lsDniTitular
'    With oWord.Selection.Find
'        .Text = "<<FirmaDni>>"
'        .Replacement.Text = sFirmaDni
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    sFirmaNombre = Me.txtDepositario & ImpreFormat(Me.txtTitularBien, 50, 30)
'    With oWord.Selection.Find
'        .Text = "<<FirmaNombre>>"
'        .Replacement.Text = sFirmaNombre
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    oWord.Selection.GoTo What:=wdGoToBookmark, Name:="Listado"
'
'    nNroItem = 0
'    With oWord
'         .ActiveDocument.Tables.Add Range:=.Selection.Range, NumRows:=Me.fgBienes.Rows, NumColumns:=6
'         .ActiveDocument.Tables(1).Borders.Enable = True
'         .ActiveDocument.Tables(1).Cell(1, 1).Range.InsertAfter "Nro"
'         .ActiveDocument.Tables(1).Cell(1, 2).Range.InsertAfter "Bien"
'         .ActiveDocument.Tables(1).Cell(1, 3).Range.InsertAfter "Estado"
'         .ActiveDocument.Tables(1).Cell(1, 4).Range.InsertAfter "Cant"
'         .ActiveDocument.Tables(1).Cell(1, 5).Range.InsertAfter "Almacen"
'         .ActiveDocument.Tables(1).Cell(1, 6).Range.InsertAfter "Descripcion"
'
'         .ActiveDocument.Tables(1).Cell(1, 1).Column.Width = 25
'         .ActiveDocument.Tables(1).Cell(1, 2).Column.Width = 90
'         .ActiveDocument.Tables(1).Cell(1, 3).Column.Width = 45
'         .ActiveDocument.Tables(1).Cell(1, 4).Column.Width = 33
'         .ActiveDocument.Tables(1).Cell(1, 5).Column.Width = 90
'         .ActiveDocument.Tables(1).Cell(1, 6).Column.Width = 140
'         .ActiveDocument.Tables(1).Rows.iTem(1).Range.Font.Bold = True
'         .ActiveDocument.Tables(1).Rows.iTem(1).Range.Font.Size = 10
'
'
'        For i = 1 To Me.fgBienes.Rows - 1
'            If fgBienes.TextMatrix(i, 1) = "." Then
'                nNroItem = nNroItem + 1
'
'                .ActiveDocument.Tables(1).Cell(nNroItem + 1, 1).Range.InsertAfter nNroItem
'                .ActiveDocument.Tables(1).Cell(nNroItem + 1, 2).Range.InsertAfter Trim(Left(Me.fgBienes.TextMatrix(i, 6), 100))
'                .ActiveDocument.Tables(1).Cell(nNroItem + 1, 3).Range.InsertAfter Trim(Left(Me.fgBienes.TextMatrix(i, 15), 20))
'                .ActiveDocument.Tables(1).Cell(nNroItem + 1, 4).Range.InsertAfter Trim(Left(Me.fgBienes.TextMatrix(i, 7), 10))
'                .ActiveDocument.Tables(1).Cell(nNroItem + 1, 5).Range.InsertAfter Trim(Left(Me.fgBienes.TextMatrix(i, 14), 150))
'                .ActiveDocument.Tables(1).Cell(nNroItem + 1, 6).Range.InsertAfter Trim(Left(Me.fgBienes.TextMatrix(i, 8), 200))
'                .ActiveDocument.Tables(1).Rows.iTem(nNroItem + 1).Range.Font.Size = 10
'
'            End If
'        Next i
'    End With
'
'    Dim sArchivo As String
'
'   If Right(Me.cmbTipoSalida.Text, 1) = "3" Then
'        sArchivo = App.Path & "\SPOOLER\DevDesafectacion_" & Me.AXCodCta.NroCuenta + "_" + Format(gdFecSis, "yyyymmdd") & ".doc"
'        oDoc.SaveAs sArchivo
'   ElseIf Right(Me.cmbTipoSalida.Text, 1) = "4" Then
'         sArchivo = App.Path & "\SPOOLER\DevPago_" & Me.AXCodCta.NroCuenta + "_" + Format(gdFecSis, "yyyymmdd") & ".doc"
'         oDoc.SaveAs sArchivo
'   End If
'
'    oDoc.Close
'    Set oDoc = Nothing
'
'    Set oWord = CreateObject("Word.Application")
'    oWord.Visible = True
'
'    Set oDoc = oWord.Documents.Open(sArchivo)
'
'    oWord.Visible = True
'    Set oDoc = Nothing
'    Set oWord = Nothing
'
'End Sub
'Private Sub ImprimirExcel()
'    MousePointer = vbHourglass
'
'    Dim sPathEmbargo As String
'    Dim fs As New Scripting.FileSystemObject
'    Dim obj_Excel As Object, Libro As Object, Hoja As Object
'
'    sPathEmbargo = App.Path & "\FormatoCarta\PLANTILLA_EMBARGO.xls"
'
'    If Len(Dir(sPathEmbargo)) = 0 Then
'           MsgBox "No se Pudo Encontrar el Archivo:" & sPathEmbargo, vbCritical
'           Me.MousePointer = vbDefault
'           Exit Sub
'    End If
'
'    Set obj_Excel = CreateObject("Excel.Application")
'    obj_Excel.DisplayAlerts = False
'    Set Libro = obj_Excel.Workbooks.Open(sPathEmbargo)
'    Set Hoja = Libro.ActiveSheet
'    Dim celda As Excel.Range
'
'    Set celda = obj_Excel.Range("A3") ' NOMBRE AGENCIA
'    celda.value = gsNomAge
'
'    Set celda = obj_Excel.Range("H2") 'USUARIO
'    celda.value = "Usuario : " & gsCodUser
'
'    Set celda = obj_Excel.Range("H3") 'FECHA
'    celda.value = "Fecha     : " & gdFecSis
'
'    Set celda = obj_Excel.Range("C7") 'Cliente
'    celda.value = Me.lbltitular
'    Set celda = obj_Excel.Range("C8") 'Cuenta
'    celda.value = Me.AXCodCta.NroCuenta
'    Set celda = obj_Excel.Range("C9") 'Expediente
'    celda.value = Me.txtExpediente.Text
'    Set celda = obj_Excel.Range("E9") 'Resolucion
'    celda.value = Me.txtResolucion.Text
'    Set celda = obj_Excel.Range("H9") 'Fecha
'    celda.value = Me.mskFecha.Text
'    Set celda = obj_Excel.Range("C10") 'Capital
'    celda.value = Format(Me.lblCapital, "##,##0.00")
'    Set celda = obj_Excel.Range("E10") 'Interes
'    celda.value = Format(Me.lblInteres, "##,##0.00")
'    Set celda = obj_Excel.Range("H10") 'Gastos
'    celda.value = Format(Me.lblGasto, "##,##0.00")
'
'    Dim nNroItem As Integer
'    Dim i As Integer
'    nNroItem = 0
'    For i = 1 To Me.fgBienes.Rows - 1
'
'        If Right(Me.fgBienes.TextMatrix(i, 2), 1) = 1 Then
'            nNroItem = nNroItem + 1
'             Set celda = obj_Excel.Range("A" + CStr(12 + nNroItem)) 'Nº
'             celda.value = nNroItem
'
'             Set celda = obj_Excel.Range("B" + CStr(12 + nNroItem), "C" + CStr(12 + nNroItem)) 'BIEN
'             celda.value = Trim(Left(Me.fgBienes.TextMatrix(i, 6), 50))
'             celda.MergeCells = True
'
'             Set celda = obj_Excel.Range("D" + CStr(12 + nNroItem)) 'ESTADO
'             celda.value = Trim(Left(Me.fgBienes.TextMatrix(i, 15), 50))
'
'             Set celda = obj_Excel.Range("E" + CStr(12 + nNroItem)) 'CANTIDAD
'             celda.value = Trim(Left(Me.fgBienes.TextMatrix(i, 7), 50))
'
'            Set celda = obj_Excel.Range("F" + CStr(12 + nNroItem), "G" + CStr(12 + nNroItem)) 'ALMACEN
'             celda.value = Trim(Left(Me.fgBienes.TextMatrix(i, 14), 50))
'             celda.MergeCells = True
'
'             Set celda = obj_Excel.Range("H" + CStr(12 + nNroItem), "J" + CStr(12 + nNroItem)) 'DESCRIPCION
'             celda.value = Trim(Me.fgBienes.TextMatrix(i, 8))
'             celda.MergeCells = True
'
'
'        End If
'
'    Next i
'
'     Set celda = obj_Excel.Range("A13", "J" + CStr(12 + nNroItem))
'
'     celda.Borders.LineStyle = 7
'
'    sPathEmbargo = App.Path & "\Spooler\Embargo_" + Me.AXCodCta.NroCuenta + "_" + Format(gdFecSis, "yyyymmdd") + ".xls"
'
'     Hoja.SaveAs sPathEmbargo
'     Libro.Close
'     obj_Excel.Quit
'     Set Hoja = Nothing
'     Set Libro = Nothing
'     Set obj_Excel = Nothing
'     Me.MousePointer = vbDefault
'     Dim m_excel As New Excel.Application
'     m_excel.Workbooks.Open (sPathEmbargo)
'     m_excel.Visible = True
'
'
'End Sub
'Private Function validaDatos() As Boolean
'    validaDatos = True
'    If Me.AXCodCta.Cuenta = "" Then
'        validaDatos = False
'        MsgBox "Debe Ingresar una Cuenta Valida", vbExclamation, "Aviso"
'        Me.AXCodCta.SetFocus
'        Exit Function
'    End If
'    If Me.txtExpediente.Text = "" Then
'        validaDatos = False
'        MsgBox "Debe Ingresar el Nro de Expediente", vbExclamation, "Aviso"
'        Me.txtExpediente.SetFocus
'        Exit Function
'    End If
'    If Me.txtResolucion.Text = "" Then
'        validaDatos = False
'        MsgBox "Debe Ingresar el Nro de Resolucion", vbExclamation, "Aviso"
'        Me.txtResolucion.SetFocus
'        Exit Function
'    End If
'    If Me.fgBienes.TextMatrix(1, 4) = "" Then
'        validaDatos = False
'        MsgBox "Debe Ingresar al menos un Bien", vbExclamation, "Aviso"
'        Me.cmdAgregarBien.SetFocus
'        Exit Function
'    End If
'    Dim sFecha As String
'    sFecha = ValidaFecha(mskFecha.Text)
'    If sFecha <> "" Then
'        validaDatos = False
'        MsgBox sFecha, vbExclamation, "Aviso"
'        Me.mskFecha.SetFocus
'        Exit Function
'    End If
'
'    If lnOpcion = 1 Then
'        If Me.txtDocSalida.Text = "" Then
'            validaDatos = False
'            MsgBox "Debe Ingresar el Documento de Salida", vbExclamation, "Aviso"
'            Me.txtDocSalida.SetFocus
'            Exit Function
'        End If
'        sFecha = ValidaFecha(Me.mskSalida.Text)
'        If sFecha <> "" Then
'            validaDatos = False
'            MsgBox "Fecha de Salida:" + sFecha, vbExclamation, "Aviso"
'            Me.mskSalida.SetFocus
'            Exit Function
'        End If
'        If Me.cmbTipoSalida.ListIndex = -1 Then
'            validaDatos = False
'            MsgBox "Debe Seleccionar el Tipo de Salida", vbExclamation, "Aviso"
'            Me.cmbTipoSalida.SetFocus
'            Exit Function
'        End If
'        If Right(Me.cmbTipoSalida.Text, 1) <> "2" Then
'            If Me.txtDepositario = "" Then
'                validaDatos = False
'                MsgBox "Debe Ingresar el Depositario", vbExclamation, "Aviso"
'                Me.cmdBuscarDepo.SetFocus
'                Exit Function
'            End If
'            If Me.txtTitularBien = "" Then
'                validaDatos = False
'                MsgBox "Debe Ingresar el Titular del Bien", vbExclamation, "Aviso"
'                Me.cmdBuscarTitularBien.SetFocus
'                Exit Function
'            End If
'        End If
'
'        Dim i As Integer
'        Dim valor As Integer
'        valor = 0
'        With Me.fgBienes
'            For i = 1 To Me.fgBienes.Rows - 1
'                If .TextMatrix(i, 1) <> "" Then
'                    valor = 1
'                    Exit For
'                End If
'            Next i
'            If valor = 0 Then
'                validaDatos = False
'                MsgBox "Debe Seleccionar al menos un bien", vbExclamation, "Aviso"
'                Me.fgBienes.SetFocus
'                Exit Function
'            End If
'        End With
'    End If
'
'End Function
'Private Sub cmdModificar_Click()
'    ModificarBien
'End Sub
'
'Private Sub fgBienes_Click()
'    If fgBienes.TextMatrix(fgBienes.row, 4) <> "" Then
'      If lnOpcion = consNuevoEmbargo Or lnOpcion = consModificarEmbargo Then
'        cmdModificar.Enabled = True
'        Me.cmdEliminarBien.Enabled = True
'      End If
'    End If
'End Sub
'
'Private Sub fgBienes_DblClick()
'            ModificarBien
'End Sub
'Private Sub ModificarBien()
'            If fgBienes.TextMatrix(fgBienes.row, 4) <> "" Then
'                frmColEmbargoBien.TpoBien = fgBienes.TextMatrix(fgBienes.row, 4)
'                frmColEmbargoBien.SubTpoBien = fgBienes.TextMatrix(fgBienes.row, 5)
'                frmColEmbargoBien.Bien = IIf(fgBienes.TextMatrix(fgBienes.row, 6) = "", "[Ingrese Bien]" + Space(100) + "0", fgBienes.TextMatrix(fgBienes.row, 6))
'                frmColEmbargoBien.Cantidad = fgBienes.TextMatrix(fgBienes.row, 7)
'                frmColEmbargoBien.Descripcion = fgBienes.TextMatrix(fgBienes.row, 8)
'                frmColEmbargoBien.Color = fgBienes.TextMatrix(fgBienes.row, 9)
'                frmColEmbargoBien.Marca = fgBienes.TextMatrix(fgBienes.row, 10)
'                frmColEmbargoBien.Modelo = fgBienes.TextMatrix(fgBienes.row, 11)
'                frmColEmbargoBien.Serie = fgBienes.TextMatrix(fgBienes.row, 12)
'                frmColEmbargoBien.Motor = fgBienes.TextMatrix(fgBienes.row, 13)
'                frmColEmbargoBien.Almacen = IIf(fgBienes.TextMatrix(fgBienes.row, 14) = "", "[Ingrese Almacen]" + Space(100) + "0", fgBienes.TextMatrix(fgBienes.row, 14))
'                frmColEmbargoBien.Estado = fgBienes.TextMatrix(fgBienes.row, 15)
'                frmColEmbargoBien.Placa = fgBienes.TextMatrix(fgBienes.row, 16)
'                frmColEmbargoBien.Partida = fgBienes.TextMatrix(fgBienes.row, 17)
'                frmColEmbargoBien.Tasacion = fgBienes.TextMatrix(fgBienes.row, 18)
'                frmColEmbargoBien.FecTasacion = fgBienes.TextMatrix(fgBienes.row, 19)
'                frmColEmbargoBien.MonTasacion = CInt(IIf(fgBienes.TextMatrix(fgBienes.row, 20) = "", 0, Right(fgBienes.TextMatrix(fgBienes.row, 20), 1)) - 1)
'                'frmColEmbargoBien.Inicio fgBienes.Row
'                AgregarBien frmColEmbargoBien.Inicio(fgBienes.row), fgBienes.row
'            End If
'
'End Sub
'Private Sub cargarCombo()
'    Dim rsCombo As Recordset
'    Dim oColRec As COMNColocRec.NCOMColRecCredito
'
'    Set oColRec = New COMNColocRec.NCOMColRecCredito
'    Set rsCombo = oColRec.ObtenerConsValorEmbargo(9989, "[^1]", "%")
'
'    If Not (rsCombo.BOF And rsCombo.EOF) Then
'       Llenar_Combo_con_Recordset rsCombo, Me.cmbTipoSalida
'       Set rsCombo = Nothing
'    End If
'
'End Sub
'Private Sub fgBienes_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        ModificarBien
'    End If
'End Sub
'
'Private Sub mskFecha_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.cmdAgregarBien.SetFocus
'    End If
'End Sub
'
'Private Sub mskSalida_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.cmbTipoSalida.SetFocus
'    End If
'End Sub
'Private Sub txtDocSalida_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.mskSalida.SetFocus
'    End If
'End Sub
'
'Private Sub txtExpediente_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.txtResolucion.SetFocus
'    End If
'End Sub
'
'Private Sub txtResolucion_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.mskFecha.SetFocus
'    End If
'End Sub
