VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAnalisisRegulaPend 
   Caption         =   "Operaciones con Pendientes: Regulariza"
   ClientHeight    =   5715
   ClientLeft      =   1155
   ClientTop       =   1785
   ClientWidth     =   10995
   Icon            =   "frmAnalisisRegulaPend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10995
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   8730
      TabIndex        =   4
      Top             =   0
      Width           =   2175
      Begin VB.CheckBox chkFechaRegula 
         Caption         =   "Fecha Regularización"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   15
         Width           =   1860
      End
      Begin MSMask.MaskEdBox txtFecRegula 
         Height          =   345
         Left            =   750
         TabIndex        =   5
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   105
         TabIndex        =   18
         Top             =   300
         Width           =   645
      End
   End
   Begin Sicmact.FlexEdit lvPend 
      Height          =   3705
      Left            =   120
      TabIndex        =   15
      Top             =   780
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   6535
      Cols0           =   15
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Ord-Ok-Tipo-Numero-Persona-Fecha-Importe-cMovDesc-cCodPers-nMovNro-nDocTpo-Saldo-Rendicion-cMovNro-cCodOpe"
      EncabezadosAnchos=   "0-400-800-1600-3200-1200-1100-2000-0-0-0-1100-1100-0-0"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X-X-X-12-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-4-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-L-C-R-L-L-L-L-R-R-L-L"
      FormatosEdit    =   "0-0-0-0-0-0-2-0-0-0-0-2-2-0-0"
      TextArray0      =   "Ord"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbOrdenaCol     =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CheckBox chkAge 
      Caption         =   "Operaciones de Agencias"
      Height          =   225
      Left            =   4710
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cuenta a Regularizar"
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
      Height          =   645
      Left            =   120
      TabIndex        =   13
      Top             =   30
      Width           =   6885
      Begin Sicmact.TxtBuscar txtCtaPend 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1545
         _ExtentX        =   2725
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
      Begin VB.Label txtCtaPendDes 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   1650
         TabIndex        =   1
         Top             =   225
         Width           =   5085
      End
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   345
      Left            =   7110
      TabIndex        =   2
      Top             =   240
      Width           =   1155
   End
   Begin VB.CheckBox chkTodo 
      Caption         =   "&Incluir Pendientes Regularizadas"
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   5250
      Width           =   2685
   End
   Begin VB.CommandButton cmdSaldo 
      Caption         =   "&Rendición"
      Height          =   345
      Left            =   8220
      TabIndex        =   10
      ToolTipText     =   "Ingresar Saldo a Caja Chica"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   585
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4560
      Width           =   10785
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   9450
      TabIndex        =   11
      Top             =   5280
      Width           =   1185
   End
   Begin MSComctlLib.ImageList imgRec 
      Left            =   150
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalisisRegulaPend.frx":030A
            Key             =   "recibo"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraAge 
      Height          =   495
      Left            =   2970
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   4995
      Begin Sicmact.TxtBuscar txtAgeCod 
         Height          =   315
         Left            =   930
         TabIndex        =   16
         Top             =   150
         Width           =   735
         _ExtentX        =   1296
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
      Begin VB.Label lblAgeDesc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1650
         TabIndex        =   17
         Top             =   150
         Width           =   3285
      End
      Begin VB.Label Label1 
         Caption         =   "Agencias"
         Height          =   225
         Left            =   180
         TabIndex        =   9
         Top             =   180
         Width           =   765
      End
   End
   Begin RichTextLib.RichTextBox rtxtAsiento 
      Height          =   375
      Left            =   6270
      TabIndex        =   12
      Top             =   5250
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmAnalisisRegulaPend.frx":0404
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAnalisisRegulaPend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs      As New ADODB.Recordset
Dim sSql    As String

Dim lSalir  As Boolean
Dim lTransActiva As Boolean
Dim sCtaPendiente As String
Dim sTpoCta       As String

Dim lbActCheque As Boolean, lbActCarta     As Boolean, lbActOrdenP As Boolean
Dim lbActNotaAC As Boolean, lbActEfectivo  As Boolean, lbActOtros  As Boolean
Dim lbSeleAge   As Boolean, lnTpoRendicion As Integer
Dim lsClaseCta  As String, lsCtaAgePend    As String
Dim lbActPagVent As Boolean

Public Sub Inicio(psClaseCta As String, pbActCheque As Boolean, pbActOrdenP As Boolean, pbActNotaAC As Boolean, pbActEfectivo As Boolean, pbActOtros As Boolean, Optional pbActCarta As Boolean = False, Optional pbSeleAge As Boolean = False, Optional pnTpoRendicion As Integer = 0, Optional psCtaAgePend As String = "", Optional pbActPagVent As Boolean = False)
lsClaseCta = psClaseCta
lbActCheque = pbActCheque
lbActOrdenP = pbActOrdenP
lbActNotaAC = pbActNotaAC
lbActEfectivo = pbActEfectivo
lbActOtros = pbActOtros
lbActCarta = pbActCarta
lbSeleAge = pbSeleAge
lnTpoRendicion = pnTpoRendicion
lsCtaAgePend = psCtaAgePend
lbActPagVent = pbActPagVent
Me.Show 1
End Sub

Private Function ValidaDatos() As Boolean
Dim CadTemp As String
   ValidaDatos = False
    CadTemp = ValidaFecha(txtFecRegula.Text)
    If Len(Trim(CadTemp)) > 0 Then
        MsgBox CadTemp, vbInformation, "Aviso"
        txtFecRegula.SetFocus
        Exit Function
    End If
    ValidaDatos = True
End Function

Private Sub chkAge_Click()
If chkAge.value = vbChecked Then
    fraAge.Visible = True
    txtAgeCod.SetFocus
Else
    fraAge.Visible = False
    txtAgeCod = ""
End If
End Sub

Private Sub chkFechaRegula_Click()
If Me.chkFechaRegula.value = 1 Then
    Me.txtFecRegula.Enabled = True
Else
    Me.txtFecRegula.Enabled = False
End If
End Sub

Private Sub cmdProcesar_Click()
Dim lvItem As ListItem
Dim sPlaCod As String
Dim oCon As New DConecta
Dim oAna As New NAnalisisCtas
Dim sOpeCod As String

If txtCtaPend = "" Then
    MsgBox "Debe seleccionar Cuenta Pendiente", vbInformation, "¡Aviso!"
    Exit Sub
End If
lsCtaAgePend = Mid(txtCtaPend, 1, 2) & "_" & Mid(txtCtaPend, 4, 22) & "%"

lvPend.Clear
lvPend.Rows = 2
lvPend.FormaCabecera
If Not ValidaDatos() Then
   Exit Sub
End If
Select Case gsOpeCod
    Case gAnaSubsidioPrePostAbonoSeguro, gAnaSubsidioPrePostAbonoSeguroME, gAnaSubsidioEnfermAbonoSeguro, gAnaSubsidioEnfermAbonoSeguroME, gsRHPlanillaSubsidio
        If Mid(gsOpeCod, 5, 1) = "0" Then
            sPlaCod = gsRHPlanillaSubsidio
            sOpeCod = "622601"
        Else
            sPlaCod = gsRHPlanillaSubsidioEnfermedad
            sOpeCod = "622602"
        End If
        Set rs = oAna.GetSubsidiosPendientes(sPlaCod, txtFecRegula, sOpeCod, gsOpeCod, sCtaPendiente)
    'JACA 20110819**********************************************
    Case "741123", "742123"
        Set rs = oAna.GetOpePendientesRegAsienCont(txtFecRegula, Mid(gsOpeCod, 3, 1), lsCtaAgePend)
    'JACA END************************************************
    Case Else
        If fraAge.Visible Then
            If txtAgeCod = "" Then
                MsgBox "Seleccionar Agencia donde Buscar Pendiente a Regularizar", vbInformation, "¡Aviso!"
                Exit Sub
            End If
            Set rs = oAna.GetOpePendientesNegocio(gbBitCentral, txtFecRegula, Mid(gsOpeCod, 3, 1), Me.txtAgeCod, lsCtaAgePend)
        Else
            Set rs = oAna.GetOpePendientesMov(gbBitCentral, txtFecRegula, Mid(gsOpeCod, 3, 1), lsCtaAgePend, sTpoCta)
        End If
End Select
If rs Is Nothing Then
    Exit Sub
End If
If rs.EOF Then
   MsgBox "No existen Pendientes por Regularizar", vbInformation, "¡Aviso!"
   RSClose rs
   Exit Sub
End If
Set lvPend.Recordset = rs
lvPend.FormateaColumnas
lvPend.Row = 1
End Sub

Private Sub cmdSaldo_Click()
Dim lbOk As Boolean
Dim N As Integer
Dim lsPersCod As String
Dim lnSaldo   As Currency
If lvPend.TextMatrix(1, 0) <> "" Then
   gnImporte = 0
   lnSaldo = 0
   For N = 1 To lvPend.Rows - 1
      If lvPend.TextMatrix(N, 1) = "." Then
         gnImporte = gnImporte + nVal(lvPend.TextMatrix(N, 12))
         lnSaldo = lnSaldo + nVal(lvPend.TextMatrix(N, 11))
         gnMovNro = lvPend.TextMatrix(N, 9)
         gsMovNro = lvPend.TextMatrix(N, 13)
         gsPersNombre = lvPend.TextMatrix(N, 4)
         lsPersCod = lvPend.TextMatrix(N, 8)
         gsDocNro = lvPend.TextMatrix(N, 3)
      End If
   Next
   If gnImporte = 0 Then
        MsgBox "Falta seleccionar Pendiente a Regularizar", vbInformation, "¡Aviso!"
        Exit Sub
   End If
   gsGlosa = txtMovDesc
   gdFecha = txtFecRegula
    lbOk = False
    If lnTpoRendicion = 0 Then
      If sTpoCta = "D" Then
         frmAnalRegulaPendIngreso.Inicio lbActCheque, lbActOrdenP, lbActNotaAC, lbActEfectivo, lbActOtros, lbActPagVent, lsPersCod, lnSaldo
         lbOk = frmAnalRegulaPendIngreso.lOk
      Else
         frmAnalRegulaPendSalida.Inicio lbActCheque, lbActCarta, lbActOrdenP, lbActNotaAC, lbActEfectivo, lbActOtros, lsPersCod, lnSaldo
         lbOk = frmAnalRegulaPendSalida.lOk
      End If
    End If
    If lnTpoRendicion = 1 Then
       If gsOpeCod = gAnaOtraProvisRegulaProvision Or gsOpeCod = gAnaOtraProvisRegulaProvisionME Then
          frmLogProvisionPago.Inicio False, False, lsPersCod, True, True, True, True
       Else
          frmLogProvisionPago.Inicio False, False, lsPersCod, True, lsClaseCta = "D", True, True
       End If
       lbOk = frmLogProvisionPago.lOk
    End If
    If lnTpoRendicion = 2 Then
       frmAsientoRegistro.Inicio "", 0, , True, True, False, True, txtAgeCod, Me.lvPend.GetRsNew
       lbOk = frmAsientoRegistro.lOk
    End If
    If lbOk Then
       N = 1
       Do While N <= lvPend.Rows - 1
          If lvPend.TextMatrix(N, 1) = "." Then
            If nVal(lvPend.TextMatrix(N, 11)) <> nVal(lvPend.TextMatrix(N, 12)) Then
                lvPend.TextMatrix(N, 11) = nVal(lvPend.TextMatrix(N, 11)) - nVal(lvPend.TextMatrix(N, 12))
                lvPend.TextMatrix(N, 12) = 0
                lvPend.TextMatrix(N, 1) = ""
                N = N + 1
            Else
                lvPend.EliminaFila N
            End If
          Else
             N = N + 1
          End If
       Loop
    End If
   lvPend.SetFocus
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
If lSalir Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim lvItem As ListItem
Dim rsPer As New ADODB.Recordset
Dim sOpeCod As String
Dim oOpe As New DOperacion

CentraForm Me

AbreConexion
lSalir = False
Me.Caption = gsOpeDesc
CentraForm Me
lTransActiva = False
If Mid(gsOpeCod, 3, 1) = "1" Then
   gsSimbolo = gcMN
Else
   gsSimbolo = gcME
End If
txtCtaPend.rs = oOpe.CargaOpeCta(gsOpeCod, lsClaseCta, "0")
txtFecRegula = gdFecSis

gnDocTpo = 0
gsDocNro = ""
gsGlosa = ""
Cmdsaldo.Visible = True

Dim oCla As New DCtaCont
Set rs = oCla.CargaCtaContClase(txtCtaPend)
sTpoCta = Trim(rs!cCtaCaracter)
RSClose rs
If lbSeleAge Then
    chkAge.Visible = True
    Dim oAge As New DActualizaDatosArea
    txtAgeCod.rs = oAge.GetAgencias(, True)
    Set oAge = Nothing
End If
End Sub

Public Property Get sPendiente() As String
sPendiente = sCtaPendiente
End Property

Public Property Let sPendiente(ByVal vNewValue As String)
sCtaPendiente = sPendiente
End Property


Private Sub lvPend_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
If lvPend.TextMatrix(pnRow, 1) = "." Then
    lvPend.TextMatrix(pnRow, 12) = lvPend.TextMatrix(pnRow, 11)
    Me.txtMovDesc = lvPend.TextMatrix(lvPend.Row, 7) 'JACA 20110829
Else
    lvPend.TextMatrix(pnRow, 12) = "0.00"
    Me.txtMovDesc = "" 'JACA 20110829
End If
End Sub

Private Sub lvPend_RowColChange()
If lvPend.TextMatrix(1, 0) = "" Then
   txtMovDesc = ""
Else
   txtMovDesc = lvPend.TextMatrix(lvPend.Row, 7)
End If
End Sub

Private Sub txtAgeCod_EmiteDatos()
Me.lblAgeDesc = txtAgeCod.psDescripcion
If lblAgeDesc <> "" And cmdProcesar.Visible Then
    cmdProcesar.SetFocus
End If
End Sub

Private Sub txtCtaPend_EmiteDatos()
txtCtaPendDes = txtCtaPend.psDescripcion
sCtaPendiente = txtCtaPend
If sCtaPendiente <> "" Then
    Dim oCla As New DCtaCont
    Set rs = oCla.CargaCtaContClase(txtCtaPend)
    sTpoCta = Trim(rs!cCtaCaracter)
    RSClose rs

    Me.lvPend.Rows = 2
    Me.lvPend.Clear
    Me.lvPend.FormaCabecera
    Me.lvPend.FormateaColumnas
    If cmdProcesar.Visible Then
        cmdProcesar.SetFocus
    End If
End If
End Sub

Private Sub txtFecRegula_GotFocus()
fEnfoque txtFecRegula
End Sub

Private Sub txtFecRegula_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFecRegula) <> "" Then
      MsgBox "Fecha no válida", vbInformation, "¡Aviso!"
      Exit Sub
   End If
   cmdProcesar.SetFocus
End If
End Sub

