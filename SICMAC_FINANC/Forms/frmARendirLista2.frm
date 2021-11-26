VERSION 5.00
Begin VB.Form frmARendirLista2 
   Caption         =   "A rendir Cuenta: Pendientes de regularizar"
   ClientHeight    =   5115
   ClientLeft      =   885
   ClientTop       =   2145
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   10245
   Begin VB.CheckBox chkTodos 
      Caption         =   "Incluir Arendir Cuenta Sustentados"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   7230
      TabIndex        =   16
      Top             =   4860
      Width           =   2805
   End
   Begin VB.CheckBox chkSelec 
      Caption         =   "&Todos"
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
      Height          =   210
      Left            =   165
      TabIndex        =   1
      Top             =   645
      Value           =   1  'Checked
      Width           =   900
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8610
      TabIndex        =   9
      Top             =   4410
      Width           =   1380
   End
   Begin VB.CommandButton cmdRegulariza 
      Caption         =   "S&ustentación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7245
      TabIndex        =   8
      ToolTipText     =   "Regularizar con Documentos sustentatorios"
      Top             =   4410
      Width           =   1380
   End
   Begin VB.Frame FraSeleccion 
      Enabled         =   0   'False
      Height          =   945
      Left            =   90
      TabIndex        =   14
      Top             =   615
      Width           =   9885
      Begin Sicmact.TxtBuscar txtBuscarAgenciaArea 
         Height          =   330
         Left            =   1425
         TabIndex        =   2
         Top             =   180
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
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
         lbUltimaInstancia=   0   'False
      End
      Begin VB.Label lblAgeDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1410
         TabIndex        =   4
         Top             =   525
         Width           =   6420
      End
      Begin VB.Label lblAgenciaArea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   3
         Top             =   195
         Width           =   5310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Area/Agencia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   15
         Top             =   225
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Glosa"
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
      Height          =   915
      Left            =   90
      TabIndex        =   13
      Top             =   4095
      Width           =   5640
      Begin VB.TextBox txtMovDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   195
         Width           =   5370
      End
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   105
      TabIndex        =   10
      Top             =   0
      Width           =   9885
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8220
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
      Begin Sicmact.TxtBuscar TxtBuscarArendir 
         Height          =   330
         Left            =   1425
         TabIndex        =   0
         Top             =   180
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
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
         lbUltimaInstancia=   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "A rendir de... :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   12
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label lblDescArendir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   11
         Top             =   195
         Width           =   5280
      End
   End
   Begin Sicmact.Usuario usu 
      Left            =   885
      Top             =   5760
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdRendicion 
      Caption         =   "&Rendicion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      ToolTipText     =   "Ingresar Saldo a Caja General"
      Top             =   4410
      Visible         =   0   'False
      Width           =   1380
   End
   Begin Sicmact.FlexEdit fgAtenciones 
      Height          =   2385
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   4207
      Cols0           =   19
      HighLight       =   2
      AllowUserResizing=   3
      EncabezadosNombres=   $"frmARendirLista2.frx":0000
      EncabezadosAnchos=   "350-450-1000-900-1100-900-2500-1000-0-0-0-0-1000-2000-0-0-0-2000-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-C-L-C-L-R-C-L-L-L-R-L-L-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-1-2-0-0-0-0-2-0-0-0-0-0-0"
      TextArray0      =   "N°"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
   End
End
Attribute VB_Name = "frmARendirLista2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cOpeCod As String, cOpeDesc As String
Dim lTransActiva As Boolean
Dim lRindeCajaG As Boolean, lEsChica As Boolean
Dim lRindeCajaCh As Boolean, lRindeViaticos As Boolean
Dim sObjRendir As String
Dim sDocTpoRecibo As String
Dim sCtaPendiente As String
'************************************************************************************
'************************************************************************************
Dim oContFunc As NContFunciones
Dim oAreas As DActualizaDatosArea
Dim oNArendir As NARendir
Dim oOperacion As DOperacion

Dim lnTipoArendir As ArendirTipo
Dim lbEsChica  As Boolean
Dim lsCtaArendir As String
Dim lsCtaPendiente As String
Dim lsDocTpoRecibo As String

Dim lsTpoDocVoucher  As String
Dim lSalir As Boolean
Dim lsMovNroSolicitud As String
Dim lnArendirFase As ARendirFases

Public Sub Inicio(ByVal pnTipoArendir As ArendirTipo, ByVal pnArendirFase As ARendirFases, Optional pbEsCajaChica As Boolean = False)
lnArendirFase = pnArendirFase
lnTipoArendir = pnTipoArendir
lbEsChica = pbEsCajaChica
Me.Show 1
End Sub

Private Function GetReciboEgreso() As Boolean
Dim lnFila As Long
Dim rs As ADODB.Recordset
GetReciboEgreso = False
lSalir = False
Set rs = New ADODB.Recordset
fgAtenciones.Clear
fgAtenciones.FormaCabecera
fgAtenciones.Rows = 2
If TxtBuscarArendir = "" Then
    MsgBox "Seleccione el Area/Agencia a quien solicitó el Arendir", vbInformation, "Aviso"
    TxtBuscarArendir.SetFocus
    Exit Function
End If
If chkSelec.value = 0 Then
    If txtBuscarAgenciaArea = "" Then
        MsgBox "Ingrese el Area a la cual Pertenece el Arendir", vbInformation, "Aviso"
        txtBuscarAgenciaArea.SetFocus
        Exit Function
    End If
End If
Me.MousePointer = 11
Set rs = oNArendir.GetAtencionPendArendir(chkSelec.value, Mid(txtBuscarAgenciaArea.Text, 4, 2), Mid(txtBuscarAgenciaArea.Text, 1, 3), lnTipoArendir, lsCtaArendir, Mid(gsOpeCod, 3, 1), Mid(TxtBuscarArendir, 1, 3), Mid(TxtBuscarArendir, 4, 2), chkTodos.value = vbChecked)
If Not rs.EOF And Not rs.BOF Then
   Set fgAtenciones.Recordset = rs
   fgAtenciones.FormatoPersNom 6
Else
   If lnTipoArendir = gArendirTipoCajaChica Then
      MsgBox "Caja Chica sin egresos pendientes de A rendir", vbInformation, "Aviso"
   Else
      MsgBox "Area funcional sin A rendir Cuenta Pendientes", vbInformation, "Aviso"
   End If
End If
rs.Close: Set rs = Nothing
GetReciboEgreso = True
Me.MousePointer = 0
End Function

Private Sub chkSelec_Click()
If chkSelec.value = 0 Then
    FraSeleccion.Enabled = True
    txtBuscarAgenciaArea.SetFocus
Else
    FraSeleccion.Enabled = False
    txtBuscarAgenciaArea.Text = ""
    lblAgenciaArea = ""
    lblAgeDesc = ""
  End If
End Sub

Private Sub chkSelec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdBuscar.SetFocus
End If
End Sub

Private Sub cmdBuscar_Click()
If GetReciboEgreso Then
   fgAtenciones.SetFocus
Else
   lblAgenciaArea = ""
   lblAgenciaArea = ""
End If
End Sub

Private Sub cmdRegulariza_Click()
Dim sRecEstado As String
Dim lsNroArendir As String
Dim lsNroDoc As String
Dim lsFechaDoc As String
Dim lsPersCod As String
Dim lsPersNomb As String
Dim lsAreaCod As String
Dim lsAreaDesc As String
Dim lsDescDoc As String
Dim lnImporte As Currency
Dim lnSaldo As Currency
Dim lsMovNroAtenc As String
Dim lsMovNroSolicitud As String
Dim lsAgeCod As String
Dim lsAgeDesc As String

If fgAtenciones.TextMatrix(1, 0) = "" Then
   Exit Sub
End If
lsNroArendir = fgAtenciones.TextMatrix(fgAtenciones.Row, 4)
lsNroDoc = fgAtenciones.TextMatrix(fgAtenciones.Row, 2)
lsFechaDoc = fgAtenciones.TextMatrix(fgAtenciones.Row, 5)
lsPersCod = fgAtenciones.TextMatrix(fgAtenciones.Row, 9)
lsPersNomb = fgAtenciones.TextMatrix(fgAtenciones.Row, 6)
lsAreaCod = fgAtenciones.TextMatrix(fgAtenciones.Row, 14)
lsAreaDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 13)

lsDescDoc = fgAtenciones.TextMatrix(fgAtenciones.Row, 15)
lnImporte = CCur(fgAtenciones.TextMatrix(fgAtenciones.Row, 7))
lnSaldo = CCur(fgAtenciones.TextMatrix(fgAtenciones.Row, 12))
lsMovNroAtenc = fgAtenciones.TextMatrix(fgAtenciones.Row, 10)
lsMovNroSolicitud = fgAtenciones.TextMatrix(fgAtenciones.Row, 16)
If lnTipoArendir = gArendirTipoViaticos Then
    lsAgeDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 17)
    lsAgeCod = fgAtenciones.TextMatrix(fgAtenciones.Row, 18)
Else
    lsAgeDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 17)
    lsAgeCod = fgAtenciones.TextMatrix(fgAtenciones.Row, 17)
End If
frmOpeRegDocs.Inicio lnArendirFase, lnTipoArendir, False, lsNroArendir, lsNroDoc, lsFechaDoc, lsPersCod, _
                     lsPersNomb, lsAreaCod, lsAreaDesc, lsAgeCod, lsAgeDesc, lsDescDoc, lsMovNroAtenc, lnImporte, lsCtaArendir, _
                     lsCtaPendiente, lnSaldo, lsMovNroSolicitud, , , , chkTodos.value = vbChecked

fgAtenciones.TextMatrix(fgAtenciones.Row, 12) = Format(frmOpeRegDocs.lnSaldo, gsFormatoNumeroView)
'If Val(fgAtenciones.TextMatrix(fgAtenciones.Row, 12)) = 0 Then  'Doc. Regularizado
'    fgAtenciones.EliminaFila fgAtenciones.Row
'End If
fgAtenciones.SetFocus
End Sub
Private Sub cmdRendicion_Click()
Dim sRecEstado As String
Dim lsNroArendir As String
Dim lsNroDoc As String
Dim lsFechaDoc As String
Dim lsPersCod As String
Dim lsPersNomb As String
Dim lsAreaCod As String
Dim lsAreaDesc As String
Dim lsDescDoc As String
Dim lnImporte As Currency
Dim lnSaldo As Currency
Dim lsMovNroAtenc As String
Dim lsMovNroSolicitud As String
Dim lsAgeCod As String
Dim lsAgeDesc As String
Dim lsAbrevDoc As String

If fgAtenciones.TextMatrix(1, 0) = "" Then
    MsgBox "No existen Atenciones Pendientes", vbInformation, "Aviso"
    Exit Sub
End If

lsNroArendir = fgAtenciones.TextMatrix(fgAtenciones.Row, 4)
lsNroDoc = fgAtenciones.TextMatrix(fgAtenciones.Row, 2)
lsFechaDoc = fgAtenciones.TextMatrix(fgAtenciones.Row, 5)
lsPersCod = fgAtenciones.TextMatrix(fgAtenciones.Row, 9)
lsPersNomb = fgAtenciones.TextMatrix(fgAtenciones.Row, 6)
lsAreaCod = fgAtenciones.TextMatrix(fgAtenciones.Row, 14)
lsAreaDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 13)
lsAbrevDoc = fgAtenciones.TextMatrix(fgAtenciones.Row, 1)

lsDescDoc = fgAtenciones.TextMatrix(fgAtenciones.Row, 15)
lnImporte = CCur(fgAtenciones.TextMatrix(fgAtenciones.Row, 7))
lnSaldo = CCur(fgAtenciones.TextMatrix(fgAtenciones.Row, 12))
lsMovNroAtenc = fgAtenciones.TextMatrix(fgAtenciones.Row, 10)
lsMovNroSolicitud = fgAtenciones.TextMatrix(fgAtenciones.Row, 16)
lsAgeDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 17)
lsAgeCod = fgAtenciones.TextMatrix(fgAtenciones.Row, 17)

frmArendirRendicion.Inicio lnArendirFase, lnTipoArendir, lsNroArendir, lsNroDoc, lsFechaDoc, lsPersCod, _
                     lsPersNomb, lsAreaCod, lsAreaDesc, lsAgeCod, lsAgeDesc, lsDescDoc, lsMovNroAtenc, lsAbrevDoc, lnImporte, lsCtaArendir, _
                     lsCtaPendiente, lnSaldo, lsMovNroSolicitud, txtMovDesc

If frmArendirRendicion.vbOk Then
    fgAtenciones.EliminaFila fgAtenciones.Row
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgAtenciones_Click()
If lnArendirFase <> ArendirExtornoAtencion And lnArendirFase <> ArendirExtornoRendicion Then
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 8)
    Else
        txtMovDesc = ""
    End If
End Sub

Private Sub fgAtenciones_GotFocus()
If lnArendirFase <> ArendirExtornoAtencion And lnArendirFase <> ArendirExtornoRendicion Then
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 8)
    Else
        txtMovDesc = ""
    End If

End Sub

Private Sub fgAtenciones_OnRowChange(pnRow As Long, pnCol As Long)
If fgAtenciones.TextMatrix(1, 0) <> "" Then
    If lnArendirFase <> ArendirExtornoAtencion And lnArendirFase <> ArendirExtornoRendicion Then
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 8)
    Else
        txtMovDesc = ""
    End If
End If

End Sub

Private Sub Form_Activate()
If lSalir Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim lvItem As ListItem
Dim rsPer As New ADODB.Recordset
Dim sOpeCod As String

Set oContFunc = New NContFunciones
Set oNArendir = New NARendir
Set oAreas = New DActualizaDatosArea
Set oOperacion = New DOperacion
lSalir = False
Me.Caption = gsOpeDesc

CentraForm Me

chkTodos.Visible = False

If Mid(gsOpeCod, 3, 1) = gMonedaNacional Then
   gsSimbolo = gcMN
Else
   gsSimbolo = gcME
End If
lsTpoDocVoucher = oOperacion.EmiteDocOpe(gsOpeCod, OpeDocEstOpcionalDebeExistir, OpeDocMetAutogenerado)
lsCtaArendir = oOperacion.EmiteOpeCta(gsOpeCod, "H", "0")
If lsCtaArendir = "" Then
   MsgBox "Faltan asignar Cuentas Contables a Operación." & oImpresora.gPrnSaltoLinea & "Por favor consultar con Sistemas", vbInformation, "Aviso"
   lSalir = True
   Exit Sub
End If
Select Case gsOpeCod
    Case gCGArendirCtaRendMN, gCGArendirCtaRendME, gCGArendirViatRendMN, gCGArendirViatRendME
        lsCtaPendiente = oOperacion.EmiteOpeCta(gsOpeCod, "D", "1")
    Case Else
        lsCtaPendiente = oOperacion.EmiteOpeCta(gsOpeCod, "H", "1")
End Select

If lsCtaPendiente = "" Then
   MsgBox "Falta asignar Cuenta de Pendiente a Operación." & oImpresora.gPrnSaltoLinea & "Por favor consultar con Sistemas", vbInformation, "Aviso"
   lSalir = True
   Exit Sub
End If

TxtBuscarArendir.psRaiz = "A Rendir de..."
If lnTipoArendir = gArendirTipoAgencias Then
    Set rsPer = oOperacion.CargaOpeObj(gCGArendirCtaSolMNAge, 1)
Else
    Set rsPer = oOperacion.CargaOpeObj(gCGArendirCtaSolMN, 1)
End If
If Not rsPer.EOF Then
    TxtBuscarArendir.rs = oAreas.GetAgenciasAreas(rsPer!cOpeObjFiltro, 1)
End If
txtBuscarAgenciaArea.rs = oAreas.GetAgenciasAreas

Select Case lnArendirFase
    Case ArendirSustentacion
        cmdRendicion.Visible = False
        
    Case ArendirRendicion
        cmdRendicion.Visible = True
    Case ArendirExtornoAtencion, ArendirExtornoRendicion
        cmdRendicion.Visible = False
        cmdRegulariza.Visible = False
        txtMovDesc.Locked = False
End Select
Select Case lnTipoArendir
    Case gArendirTipoCajaChica
         Me.Height = 5550
         cmdRendicion.Top = 5030 - cmdRendicion.Height
         cmdRegulariza.Top = 5030 - cmdRegulariza.Height
         cmdSalir.Top = 5030 - cmdSalir.Height
         cmdRendicion.Visible = True
         lsDocTpoRecibo = oOperacion.EmiteDocOpe(gsOpeCod, OpeDocEstObligatorioDebeExistir, OpeDocMetDigitado)
         If lsDocTpoRecibo = "" Then
            MsgBox "No se asignó Tipo de Documento Recibo de A rendir a Operación", vbCritical, "Error"
            lSalir = True
            Exit Sub
         End If
         fgAtenciones.ColWidth(2) = fgAtenciones.ColWidth(2) - 300
         fgAtenciones.ColWidth(3) = fgAtenciones.ColWidth(3) + 300
    Case gArendirTipoViaticos
         fgAtenciones.EncabezadosAnchos = "350-600-0-0-1200-900-3000-1200-0-0-0-0-1200-2000-0-0-0-2000-0"
         'fgAtenciones.FormaCabecera
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set oContFunc = Nothing
Set oNArendir = Nothing
Set oAreas = Nothing
End Sub
Public Property Get sPendiente() As String
sPendiente = sCtaPendiente
End Property
Public Property Let sPendiente(ByVal vNewValue As String)
sCtaPendiente = sPendiente
End Property
Private Sub txtBuscarAgenciaArea_EmiteDatos()
lblAgenciaArea = oAreas.GetNombreAreas(Mid(txtBuscarAgenciaArea, 1, 3))
lblAgeDesc = oAreas.GetNombreAgencia(Mid(txtBuscarAgenciaArea, 4, 2))
If txtBuscarAgenciaArea <> "" Then
   cmdBuscar.SetFocus
Else
   txtBuscarAgenciaArea.SetFocus
End If
End Sub

Private Sub TxtBuscarArendir_EmiteDatos()
lblDescArendir = Trim(TxtBuscarArendir.psDescripcion)
If TxtBuscarArendir.Enabled Then
    chkSelec.SetFocus
End If
End Sub
