VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmEmiteCheques 
   Caption         =   "Emisión de Cheques"
   ClientHeight    =   3465
   ClientLeft      =   1350
   ClientTop       =   2835
   ClientWidth     =   7320
   Icon            =   "frmEmiteCheques.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7320
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   2055
      Begin VB.OptionButton OptDolar 
         Caption         =   "Dolar"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptSoles 
         Caption         =   "Soles"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.ComboBox cboOpeTpo 
      Height          =   315
      ItemData        =   "frmEmiteCheques.frx":030A
      Left            =   1560
      List            =   "frmEmiteCheques.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   345
      Left            =   3240
      TabIndex        =   8
      Top             =   2880
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   4800
      TabIndex        =   10
      Top             =   2880
      Width           =   1365
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cliente"
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
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   7125
      Begin Sicmact.TxtBuscar txtProvCod 
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
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
         TipoBusqueda    =   3
         sTitulo         =   ""
         TipoBusPers     =   1
      End
      Begin VB.CommandButton cmdExaCab 
         Caption         =   "..."
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1770
         TabIndex        =   13
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txtProvNom 
         Height          =   345
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   4965
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   1290
      Width           =   7125
      Begin VB.TextBox txtDescripcion 
         Height          =   330
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "txtPrincipal"
         Top             =   240
         Width           =   3915
      End
      Begin MSMask.MaskEdBox txtDocFecha 
         Height          =   315
         Left            =   780
         TabIndex        =   3
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad Emis.:"
         Height          =   195
         Left            =   2040
         TabIndex        =   18
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         Height          =   225
         Left            =   210
         TabIndex        =   11
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.Frame Frame6 
      Height          =   650
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   2955
      Begin VB.TextBox txtVVenta 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   1500
         TabIndex        =   7
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label lblSTot 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   210
         Width           =   1155
      End
      Begin VB.Shape ShapeIGV 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   465
         Left            =   60
         Top             =   120
         Width           =   2835
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Entidad Financiera:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "frmEmiteCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql   As String
Dim rs     As New ADODB.Recordset
Dim lNuevo As Boolean
Dim nTasaIGV As Currency
Dim oReg As DRegVenta
Dim sNumDoc As String
Dim lsDocRef As String
Dim lsDocFecRef As Date
Dim cmdPersRefComercialEjecutado As Integer
Dim FERefComPersNoMoverdeFila As Integer
Dim lnNumRefCom As Integer
Dim MatrixHojaEval() As String
Dim nPos As Integer
Dim nDat As Integer
Dim i As Integer
Dim nTotSubTotal As Currency, nTotSubIgv As Currency, nTotFinal As Currency

Public Sub Inicio(plNuevo As Boolean, pnTasaIgv As Currency)
lNuevo = plNuevo
nTasaIGV = pnTasaIgv
Me.Show 1
End Sub

Private Sub cboOpeTpo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.OptSoles.SetFocus
End If
End Sub

Private Function datosOk() As Boolean
datosOk = False
If cboOpeTpo.ListIndex = -1 Then
   MsgBox "Tipo de Operación no definido...!", vbInformation, "! Aviso !"
   cboOpeTpo.SetFocus
   Exit Function
End If
If ValidaFecha(txtDocFecha) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "! Aviso !"
   txtDocFecha.SetFocus
   Exit Function
End If
If txtProvCod.Text = "" Then
   MsgBox "Proveedor no identificado...!", vbInformation, "! Aviso !"
   txtProvCod.SetFocus
   Exit Function
End If
datosOk = True
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

'Private Sub fg_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
'    fg.TextMatrix(pnRow, 4) = fg.TextMatrix(pnRow, 1) * fg.TextMatrix(pnRow, 3)
'End Sub

Private Sub cmdImprimir_Click()

Dim oImp As NContImprimir
Dim lsImpresion As String, lnMoneda As Integer
Dim MatDatos() As String

Set oImp = New NContImprimir

If Val(Me.txtVVenta.Text) = 0 Then
    MsgBox "No tiene monto el cheque...", vbInformation, "Aviso"
    Me.txtVVenta.SetFocus
    Exit Sub
End If

If Me.txtProvNom.Text = "" Or Me.txtDocFecha.Text = "__/__/____" Then
    MsgBox "Falta ingresar datos, por favor revise...", vbInformation, "Aviso"
    Exit Sub
End If

ReDim Preserve MatDatos(1 To 6)

MatDatos(1) = Me.txtDocFecha.Text
MatDatos(2) = Me.txtProvNom.Text
MatDatos(3) = Me.txtVVenta.Text
MatDatos(4) = IIf(Me.OptSoles.value = True, 1, 2)
MatDatos(5) = Right(Me.cboOpeTpo.Text, 13)
MatDatos(6) = Trim(Me.txtDescripcion.Text)

lsImpresion = ImprimirCheque(MatDatos)
If Len(Trim(lsImpresion)) = 0 Then
    MsgBox "Este cheque no tiene plantilla.", vbOKOnly, "Atención"
    Exit Sub
End If

EnviaPrevio lsImpresion, "CHEQUE", gnLinPage
Set oImp = Nothing

End Sub

Private Sub Form_Load()
Dim oPer As New UPersona
CentraForm Me
Set oReg = New DRegVenta
Set rs = oReg.CargaEntidades()
RSLlenaCombo rs, cboOpeTpo

Dim oOpe As New DOperacion

CentraForm Me

Me.txtDescripcion.Text = "IQUITOS"

If Not lNuevo Then
   Set rs = oReg.CargaRegistro(gnDocTpo, gsDocNro, gdFecha, gdFecha)
   If Not rs.EOF Then
      txtDocFecha = Format(rs!dDocFecha, "dd/mm/yyyy")
      txtVVenta = Format(rs!nVVenta, gsFormatoNumeroView)
      oPer.ObtieneClientexCodigo rs!cPersCod
      txtProvCod.Tag = oPer.sPersCod
      txtProvNom = oPer.sPersNombre
      txtProvCod = oPer.sPersIdnroRUC
      If txtProvCod = "" Then
         txtProvCod = oPer.sPersIdnroDNI
      End If
   End If
End If
rs.Close: Set rs = Nothing
End Sub

Private Sub txtDescripcion_GotFocus()
    fEnfoque txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtDocFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtDocFecha) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "! Aviso !"
      Exit Sub
   End If
   txtProvCod.SetFocus
End If
End Sub

Private Sub txtProvCod_EmiteDatos()
txtProvCod.Tag = txtProvCod.Text
txtProvNom = txtProvCod.psDescripcion
txtProvCod.Text = txtProvCod.sPersNroDoc
End Sub

Private Sub txtProvNom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If

End Sub

Private Sub txtVVenta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtVVenta, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   txtVVenta = Format(txtVVenta, gsFormatoNumeroView)
End If
End Sub

Public Sub RSLlenaCombo(prs As ADODB.Recordset, psCombo As ComboBox, Optional pnPosCod As Integer = 0, Optional pnPosDes As Integer = 1, Optional pbPresentaCodigo As Boolean = True)
If Not prs Is Nothing Then
   If Not prs.EOF Then
      psCombo.Clear
      Do While Not prs.EOF
        psCombo.AddItem Trim(prs!cpersnombre) & Space(100) & Trim(prs!cPersCod)
        prs.MoveNext
      Loop
   End If
End If
End Sub

