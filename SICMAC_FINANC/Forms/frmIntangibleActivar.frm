VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIntangibleActivar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intangibles - Activación"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15600
   Icon            =   "frmIntangibleActivar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   15600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   1320
      TabIndex        =   22
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   2520
      TabIndex        =   21
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   14280
      TabIndex        =   20
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   13560
      TabIndex        =   19
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   13560
      TabIndex        =   18
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame fraDatosIntang 
      Caption         =   "Datos de Intangibles"
      Height          =   2775
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.TextBox txtProveedorA 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1200
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1200
         TabIndex        =   9
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtProveedorB 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2880
         TabIndex        =   8
         Top             =   840
         Width           =   5775
      End
      Begin VB.TextBox txtValorA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1200
         TabIndex        =   7
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtValorB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2280
         TabIndex        =   6
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtPeriodoAmort 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   6960
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtTipo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6720
         TabIndex        =   1
         Top             =   1320
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtFechaComprob 
         Height          =   375
         Left            =   7440
         TabIndex        =   2
         Top             =   2280
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaAct 
         Height          =   345
         Left            =   3240
         TabIndex        =   3
         Top             =   2280
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   4680
         TabIndex        =   16
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Valor:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Periodo Amortizable (mes):"
         Height          =   255
         Left            =   4680
         TabIndex        =   13
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha de Activación :"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha de Comprobante:"
         Height          =   255
         Left            =   4680
         TabIndex        =   11
         Top             =   2400
         Width           =   1815
      End
   End
   Begin Sicmact.FlexEdit feIntangActiv 
      Height          =   2535
      Left            =   120
      TabIndex        =   24
      Top             =   3120
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   4471
      Cols0           =   12
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Cod-Proveedor-Descripcion-Tipo-Moneda-Valor-Valor MN-Periodo-F. Act.-nMovNro-cMovNro"
      EncabezadosAnchos=   "300-1200-2800-4500-1200-800-1200-1200-800-1200-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-L-C-R-C-L-C-R-L"
      FormatosEdit    =   "0-1-1-1-1-1-2-2-1-1-3-1"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmIntangibleActivar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**-------------------------------------------------------------------------------------**'
'** Formulario : frmIntangibleActivar                                                   **'
'** Finalidad  : Este formulario permite activar las intangibles que                    **'
'**                posteriormente sera amortizadas                                      **'
'** Programador: Paolo Hector Sinti Cabrera - PASI                                      **'
'** Fecha/Hora : 20140305 11:50 AM                                                      **'
'**-------------------------------------------------------------------------------------**'
Option Explicit
Dim oIntang As dIntangible
Dim fRsComprobante As ADODB.Recordset
Dim lnInserModif As Integer
Dim fsDocOrigenNro As String
Dim fnComprobanteMovNro As Long
Dim nMovNroInt As Long
Dim sCtaCont As String
Dim sCodInt As String
Dim lsMsgErr As String

Private Sub Command6_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lnInserModif = 0
    Des_HabilitarControles False
    ActualizaProvisiones
    fsDocOrigenNro = ""
    fnComprobanteMovNro = 0
    sCtaCont = ""
    sCodInt = ""
    CargarDatosIntangible
End Sub
Private Sub CargarDatosIntangible()
    Dim Row As Long
    Dim rsIntangActiv As ADODB.Recordset
    Set rsIntangActiv = New ADODB.Recordset
    FormateaFlex feIntangActiv
    Set oIntang = New dIntangible
    Set rsIntangActiv = oIntang.ListaIntangiblesActivadas()
    
    Do While Not rsIntangActiv.EOF
        feIntangActiv.AdicionaFila
        Row = feIntangActiv.Row
        feIntangActiv.TextMatrix(Row, 1) = rsIntangActiv!Codigo
        feIntangActiv.TextMatrix(Row, 2) = rsIntangActiv!Proveedor
        feIntangActiv.TextMatrix(Row, 3) = rsIntangActiv!Descripcion
        feIntangActiv.TextMatrix(Row, 4) = UCase(rsIntangActiv!Tipo)
        feIntangActiv.TextMatrix(Row, 5) = rsIntangActiv!Moneda
        feIntangActiv.TextMatrix(Row, 6) = Format(rsIntangActiv!Valor, "#,#0.00")
        feIntangActiv.TextMatrix(Row, 7) = Format(rsIntangActiv!ValorMN, "#,#0.00")
        feIntangActiv.TextMatrix(Row, 8) = rsIntangActiv!periodo
        feIntangActiv.TextMatrix(Row, 9) = rsIntangActiv!FecActiv
        feIntangActiv.TextMatrix(Row, 10) = rsIntangActiv!nMovNro
        feIntangActiv.TextMatrix(Row, 11) = rsIntangActiv!cMovNro
        rsIntangActiv.MoveNext
    Loop
End Sub
Private Sub Des_HabilitarControles(phestado As Boolean)
    cmdSeleccionar.Enabled = phestado
    txtProveedorA.Enabled = phestado
    txtProveedorB.Enabled = phestado
    txtDescripcion.Enabled = phestado
    txtTipo.Enabled = phestado
    txtValorA.Enabled = phestado
    txtValorB.Enabled = phestado
    txtPeriodoAmort.Enabled = phestado
    txtFechaAct.Enabled = phestado
    txtFechaComprob.Enabled = phestado
    cmdGrabar.Enabled = phestado
    cmdCancelar.Enabled = phestado
End Sub
Private Sub ActualizaProvisiones()
    Dim oLog As New DLogGeneral
    Set fRsComprobante = New ADODB.Recordset
    Set fRsComprobante = oLog.ListaProvisionesdeIntangibles()
    Set oLog = Nothing
End Sub
Private Sub cmdSeleccionar_Click()
    Dim lsDocNro As String
    Dim lnMovNro As Long
    Dim frmSel As New frmIntangibleProvision
    LimpiaControles
    frmSel.Inicio fRsComprobante, lsDocNro, lnMovNro
    If lsDocNro <> "" And lnMovNro > 0 Then
        fsDocOrigenNro = lsDocNro
        fnComprobanteMovNro = lnMovNro
        CargarDatosProvision
    End If
End Sub
Private Sub CargarDatosProvision()
    On Error GoTo ErrCargarDatosComprobante
    Dim oLog As New DLogGeneral
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = oLog.ProvisionIntangibleDetalle(fnComprobanteMovNro)
    If Not rs.EOF Then
        Salvar rs
    End If
    Set rs = Nothing
    Set oLog = Nothing
    Exit Sub
ErrCargarDatosComprobante:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub Salvar(ByVal rs As ADODB.Recordset)
    txtProveedorA.Text = rs!RucProveedor
    txtProveedorB.Text = rs!Proveedor
    'txtDescripcion.Text = rs!Descripcion
    txtTipo.Text = UCase(rs!Tipo)
    txtValorA.Text = rs!Moneda
    txtValorB.Text = Format(rs!Valor, "#,#0.00")
    txtFechaComprob.Text = rs!FechaDoc
    sCtaCont = rs!cCtaContCod
    If lnInserModif = 2 Then
        txtPeriodoAmort.Text = rs!periodo
        txtFechaAct.Text = rs!FecActv
        nMovNroInt = rs!nMovNro
        sCodInt = rs!cIntgCod
        txtDescripcion.Text = rs!Descripcion
    End If
    txtPeriodoAmort.Enabled = True
    txtFechaAct.Enabled = True
    txtDescripcion.Enabled = True
    Set rs = Nothing
End Sub
Private Sub InserModif(ByVal phestado As Boolean)
    cmdNuevo.Enabled = phestado
    cmdModificar.Enabled = phestado
    cmdQuitar.Enabled = phestado
    cmdGrabar.Enabled = Not phestado
    cmdCancelar.Enabled = Not phestado
End Sub
Private Sub cmdNuevo_Click()
    lnInserModif = 1
    cmdSeleccionar.Enabled = True
    InserModif False
    cmdSeleccionar.SetFocus
End Sub
Private Sub LimpiaControles()
    fsDocOrigenNro = ""
    fnComprobanteMovNro = 0
    nMovNroInt = 0
    sCtaCont = ""
    sCodInt = ""
    txtProveedorA.Text = ""
    txtProveedorB.Text = ""
    txtDescripcion.Text = ""
    txtTipo.Text = ""
    txtValorA.Text = ""
    txtValorB.Text = ""
    txtPeriodoAmort.Text = ""
    txtFechaAct.Text = gdFecSis
    txtFechaComprob.Text = gdFecSis
    lnInserModif = 1
End Sub
Private Sub cmdCancelar_Click()
    LimpiaControles
    InserModif True
    Des_HabilitarControles False
End Sub
Private Sub cmdModificar_Click()
    If feIntangActiv.TextMatrix(feIntangActiv.Row, 1) <> "" Then
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        Set oIntang = New dIntangible
        Set rs = oIntang.ObtenerIntangibleActiv(feIntangActiv.TextMatrix(feIntangActiv.Row, 10))
        lnInserModif = 2
        Salvar rs
        cmdSeleccionar.Enabled = False
        InserModif False
        Set rs = Nothing
    Else
        MsgBox "No hay Datos para Modificar", vbInformation, "Aviso!!!"
    End If
End Sub
Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    If KeyAscii = 8 Then SoloNumeros = KeyAscii  'borrado atras
    If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function
Private Sub txtPeriodoAmort_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        txtFechaAct.SetFocus
    End If
End Sub
Private Sub txtFechaAct_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub
Private Sub cmdGrabar_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If ValidaDatos Then
        Dim bexito As Boolean
        Set oIntang = New dIntangible
        If lnInserModif = 1 Then
            'Codigo para la Insercion de una Activacion
            If MsgBox("Está seguro de registrar la Activación de la Intangible", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
                Dim oMov As DMov

                Set oMov = New DMov
                gdFecha = gdFecSis
                'gsMovNro = oMov.GeneraMovNro(gdFecha, gsCodAge, gsCodUser)
                gsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                
                bexito = oIntang.RegistraIntangibleActiv(fnComprobanteMovNro, gsMovNro, Trim(Me.txtDescripcion.Text), gRegistroActivacionIntangible, IIf(Mid(sCtaCont, 6, 3) = "301", "1", IIf(Mid(sCtaCont, 6, 3) = "302", "2", 3)), Trim(txtPeriodoAmort.Text), txtFechaAct.Text, lsMsgErr)
                If bexito Then
                    MsgBox "Se ha registrado con éxito la Activación de la Intangible", vbInformation, "Aviso"
                Else
                    MsgBox lsMsgErr, vbInformation, "Aviso!!!"
                End If
        Else
            'Codigo para la Modificacion de una Activacion
            If MsgBox("Está seguro de guardar los datos modificados", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
                Set rs = oIntang.IntangibleTieneAmortizacion(nMovNroInt, sCodInt)
                If Not rs.EOF Then
                    MsgBox "No se puede Actualizar por que la intangible ya ha sido amortizada", vbInformation, "Aviso!!!"
                    cmdCancelar_Click
                    Exit Sub
                End If
                bexito = oIntang.ActualizaIntangibleActiv(nMovNroInt, Trim(Me.txtDescripcion.Text), sCodInt, Trim(txtPeriodoAmort.Text), txtFechaAct.Text, lsMsgErr)
                If bexito Then
                    MsgBox "Se ha realizado la modificación con éxito", vbInformation, "Aviso"
                Else
                    MsgBox lsMsgErr, vbInformation, "Aviso!!!"
                End If
        End If
        LimpiaControles
        InserModif True
        Des_HabilitarControles False
        CargarDatosIntangible
    End If
End Sub
Private Function ValidaDatos() As Boolean
    If lnInserModif = 1 Then
        If fnComprobanteMovNro = 0 Then
            MsgBox "Ud. debe seleccionar primero la provisión realizada", vbInformation, "Aviso"
            cmdSeleccionar.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
    If Len(txtPeriodoAmort.Text) = 0 Then
        MsgBox "Ud. tiene que ingresar el Periodo de Amortización", vbInformation, "Aviso"
        txtPeriodoAmort.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Not (IsDate(txtFechaAct.Text)) Then
        MsgBox "La Fecha de Activación no es válida", vbInformation, "Aviso"
        txtFechaAct.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    ValidaDatos = True
End Function
Private Sub cmdQuitar_Click()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim bexito As Boolean
    If feIntangActiv.TextMatrix(feIntangActiv.Row, 1) <> "" Then
        
        If MsgBox("Está seguro de Eliminar el registro seleccionado", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Set oIntang = New dIntangible
        Set rs = oIntang.IntangibleTieneAmortizacion(feIntangActiv.TextMatrix(feIntangActiv.Row, 10), feIntangActiv.TextMatrix(feIntangActiv.Row, 1))
        If Not rs.EOF And Not rs.BOF Then
            MsgBox "No se puede Eliminar por que la intangible ya ha sido amortizada", vbInformation, "Aviso!!!"
            Exit Sub
        End If
        bexito = oIntang.EliminarIntangible(feIntangActiv.TextMatrix(feIntangActiv.Row, 11), lsMsgErr)
        If bexito Then
            MsgBox "El registro ha sido eliminado con exito", vbInformation, "Aviso"
        Else
            MsgBox lsMsgErr, vbInformation, "Aviso!!!"
        End If
        CargarDatosIntangible
        ActualizaProvisiones
    Else
        MsgBox "No hay Datos para Quitar", vbInformation, "Aviso!!!"
    End If
End Sub

