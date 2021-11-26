VERSION 5.00
Begin VB.Form frmCapServConvPlanPag 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   Icon            =   "frmCapServConvPlanPag.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   90
      TabIndex        =   9
      Top             =   5355
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7335
      TabIndex        =   5
      Top             =   5355
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6255
      TabIndex        =   4
      Top             =   5355
      Width           =   975
   End
   Begin VB.Frame fraPlanPagos 
      Caption         =   "Plan Pagos"
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
      Height          =   4335
      Left            =   60
      TabIndex        =   2
      Top             =   960
      Width           =   8280
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   1020
         TabIndex        =   3
         Top             =   3840
         Width           =   795
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3840
         Width           =   795
      End
      Begin SICMACT.FlexEdit grdPlanPago 
         Height          =   3555
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6271
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Cuota-Vencimiento-Monto-Gasto-Afecto Mora-Afecto Feriado-Flag"
         EncabezadosAnchos=   "600-1300-1200-1200-1300-1300-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-4-5-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-2-0-0-3-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-R-R-C-C-C"
         FormatosEdit    =   "0-0-2-2-0-0-0"
         TextArray0      =   "Cuota"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraConvenio 
      Caption         =   "Institucion Covnenio"
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
      Height          =   855
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   8265
      Begin SICMACT.TxtBuscar txtCodigo 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   661
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
      Begin VB.Label lblInstitucion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   300
         Width           =   6120
      End
   End
End
Attribute VB_Name = "frmCapServConvPlanPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sPersonaCod As String

Private Function GetRespuestaAfecto() As ADODB.Recordset
Dim rsSI As ADODB.Recordset
Set rsSI = New ADODB.Recordset
rsSI.CursorType = adOpenStatic
rsSI.LockType = adLockOptimistic
rsSI.Fields.Append "Respuesta", adVarChar, 2
rsSI.Open
rsSI.AddNew
rsSI("Respuesta") = "SI"
rsSI.Update
rsSI.AddNew
rsSI("Respuesta") = "NO"
rsSI.Update
rsSI.MoveFirst
Set GetRespuestaAfecto = rsSI
End Function


Public Sub Inicia(Optional sPersona As String = "", Optional sPersonaDesc As String = "")
sPersonaCod = sPersona
txtCodigo.sTitulo = "Instituciones - Convenio"
If sPersonaCod = "" Then
    Dim clsCap As COMNCaptaServicios.NCOMCaptaServicios
    Dim rsCap As New ADODB.Recordset
    Set clsCap = New COMNCaptaServicios.NCOMCaptaServicios
    Set rsCap = clsCap.GetServConveniosArbol()
    txtCodigo.rs = rsCap
    fraPlanPagos.Enabled = False
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    fraConvenio.Enabled = True
Else
    txtCodigo.Text = sPersona
    lblInstitucion = sPersonaDesc
    fraConvenio.Enabled = False
    txtCodigo_EmiteDatos
End If
grdPlanPago.CargaCombo GetRespuestaAfecto
Me.Caption = "Captaciones - Servicio - Convenios - Plan Pagos"
Me.Show 1
End Sub

Private Sub CmdAgregar_Click()
Dim nFila As Long
grdPlanPago.AdicionaFila
grdPlanPago.Col = 1
nFila = grdPlanPago.Rows - 1
grdPlanPago.TextMatrix(nFila, 1) = Format$(gdFecSis, "dd/mm/yyyy")
grdPlanPago.TextMatrix(nFila, 2) = "0.00"
grdPlanPago.TextMatrix(nFila, 3) = "0.00"
grdPlanPago.SetFocus
cmdEliminar.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
grdPlanPago.Clear
grdPlanPago.Rows = 2
grdPlanPago.FormaCabecera
txtCodigo.Text = ""
lblInstitucion = ""
fraPlanPagos.Enabled = False
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
fraConvenio.Enabled = True
txtCodigo.SetFocus
End Sub

Private Sub cmdeliminar_Click()
grdPlanPago.EliminaFila grdPlanPago.Row
grdPlanPago.Col = 1
grdPlanPago.SetFocus
If grdPlanPago.Rows = 2 And grdPlanPago.TextMatrix(1, 1) = "" Then
    cmdEliminar.Enabled = False
End If
cmdAgregar.SetFocus
End Sub

Private Sub cmdGrabar_Click()
If MsgBox("¿Desea Grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim rsPlan As New ADODB.Recordset
    Set rsPlan = grdPlanPago.GetRsNew
    Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
    clsServ.ActualizaServConvPlanPagos sPersonaCod, rsPlan
    Set clsServ = Nothing
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub txtCodigo_EmiteDatos()
Dim sCodigo As String, sCuenta As String
sCodigo = txtCodigo.Text
    If sCodigo <> "" Then
        Dim clsCap As COMNCaptaServicios.NCOMCaptaServicios
        Dim rsCtas As New ADODB.Recordset
        sPersonaCod = sCodigo
        If txtCodigo.psDescripcion <> "" Then lblInstitucion = txtCodigo.psDescripcion
        Set clsCap = New COMNCaptaServicios.NCOMCaptaServicios
        Set rsCtas = clsCap.GetServPlanPagos(sCodigo)
        If Not (rsCtas.EOF And rsCtas.BOF) Then
            Set grdPlanPago.Recordset = rsCtas
        End If
        cmdGrabar.Enabled = True
        fraPlanPagos.Enabled = True
        fraConvenio.Enabled = False
        cmdCancelar.Enabled = True
    End If
End Sub
