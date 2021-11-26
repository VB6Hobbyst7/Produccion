VERSION 5.00
Begin VB.Form frmIngNotaAbonoCargo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Notas de Abono y Cargo"
   ClientHeight    =   2970
   ClientLeft      =   2475
   ClientTop       =   2820
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIngNotaAbonoCargo.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5280
      TabIndex        =   5
      Top             =   2565
      Width           =   1230
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   4065
      TabIndex        =   4
      Top             =   2565
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   180
      TabIndex        =   6
      Top             =   390
      Width           =   6330
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   4695
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1680
         Width           =   1425
      End
      Begin SICMACT.TxtBuscar txtBuscarMotivo 
         Height          =   330
         Left            =   870
         TabIndex        =   0
         Top             =   255
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   16711680
      End
      Begin VB.Frame fraObjeto 
         Height          =   960
         Left            =   60
         TabIndex        =   9
         Top             =   630
         Width           =   6135
         Begin SICMACT.TxtBuscar txtBuscarObj 
            Height          =   330
            Left            =   780
            TabIndex        =   1
            Top             =   150
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
            ForeColor       =   16711680
         End
         Begin SICMACT.TxtBuscar txtBuscarDet 
            Height          =   330
            Left            =   780
            TabIndex        =   2
            Top             =   495
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
            ForeColor       =   16711680
         End
         Begin VB.Label lblDetalleDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1845
            TabIndex        =   15
            Top             =   510
            Width           =   4110
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Detalle:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   105
            TabIndex        =   14
            Top             =   555
            Width           =   600
         End
         Begin VB.Label lblObjetoDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1860
            TabIndex        =   11
            Top             =   165
            Width           =   4110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Objeto :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   10
            Top             =   210
            Width           =   630
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3930
         TabIndex        =   12
         Top             =   1732
         Width           =   615
      End
      Begin VB.Label lblMotivoDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1950
         TabIndex        =   8
         Top             =   270
         Width           =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Motivo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   150
         TabIndex        =   7
         Top             =   315
         Width           =   645
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   345
         Left            =   3675
         Top             =   1665
         Width           =   2460
      End
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "NOTA DE ABONO "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2228
      TabIndex        =   13
      Top             =   75
      Width           =   2145
   End
End
Attribute VB_Name = "frmIngNotaAbonoCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oDocRec As nDocRec
Public lsNroNota As String
Public lsMotivoDesc As String
Dim lnDocTpo As TpoDoc
Public lnMonto As Double
Public lnMotivo As Long
Dim lsObjetoCodPadre As String
Dim lsObjetoCod As String
Public lbOk As Boolean
Public lnMotivoEsp As MotivoNotaAbonoCargo
Public bPermiteMontoCero As Boolean

Private Sub cmdAceptar_Click()
Dim lsNroNotaNCNA As String
Dim lsMovNro As String
Dim oCont As NContFunciones
Set oCont = New NContFunciones
If Valida = False Then Exit Sub

If MsgBox("Desea Registrar " & lblTitulo & "???", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsNroNotaNCNA = oDocRec.GetNroNotaCargoAbono(lnDocTpo)
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    oDocRec.RegistroNotasAbonoCargo lnDocTpo, lsNroNotaNCNA, gNCNARegistrado, txtBuscarMotivo, CCur(txtImporte), lsMovNro, _
                                    txtBuscarObj, txtBuscarDet
    
    lbOk = True
    lsNroNota = lsNroNotaNCNA
    lnMonto = CDbl(txtImporte)
    lnMotivo = CLng(txtBuscarMotivo)
    lsMotivoDesc = lblMotivoDesc
    lsObjetoCodPadre = Trim(txtBuscarObj)
    lsObjetoCod = Trim(txtBuscarDet)
    Unload Me
End If
End Sub

Function Valida() As Boolean
Valida = True
If Len(Trim(txtBuscarMotivo)) = 0 Then
    MsgBox "Motivo no ha sido ingresado", vbInformation, "Aviso"
    txtBuscarMotivo.SetFocus
    Valida = False
    Exit Function
End If

If fraObjeto.Visible And fraObjeto.Enabled Then
    If Len(Trim(txtBuscarObj)) = 0 Then
        MsgBox "Objeto no ha sido ingresado", vbInformation, "Aviso"
        txtBuscarObj.SetFocus
        Valida = False
        Exit Function
    End If
    If Len(Trim(txtBuscarDet)) = 0 Then
        MsgBox "Detalle de Objeto no ha sido ingresado", vbInformation, "Aviso"
        txtBuscarDet.SetFocus
        Valida = False
        Exit Function
    End If
End If

If Val(txtImporte) = 0 And Not bPermiteMontoCero Then
    MsgBox "Importe de documento no ingresado", vbInformation, "Aviso"
    txtImporte.SetFocus
    Valida = False
    Exit Function
End If
End Function

Public Sub Inicio(ByVal pnTpoDoc As TpoDoc, ByVal pnMonto As Currency, Optional pnMotivoEsp As MotivoNotaAbonoCargo = -1, _
            Optional ByVal bMontoCero As Boolean = False)
lnDocTpo = pnTpoDoc
lnMonto = pnMonto
lnMotivoEsp = pnMotivoEsp
bPermiteMontoCero = bMontoCero
Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
lbOk = False
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Set oDocRec = New nDocRec
txtBuscarMotivo.psRaiz = "Motivos"
txtBuscarMotivo.rs = oDocRec.GetMotivosNivel(lnDocTpo, lnMotivoEsp)
lbOk = False
If lnDocTpo = TpoDocNotaCargo Then
    lblTitulo = "NOTA DE CARGO"
Else
    lblTitulo = "NOTA DE ABONO"
End If
txtImporte = Format(lnMonto, "#,#0.00")
txtImporte.Locked = False
If Val(txtImporte) > 0 Then
    txtImporte.Locked = True
End If
End Sub
Public Property Get psNroNota() As String
psNroNota = lsNroNota
End Property
Public Property Let psNroNota(ByVal vNewValue As String)
lsNroNota = vNewValue
End Property

Public Property Get pnDocTpo() As TpoDoc
pnDocTpo = lnDocTpo
End Property

Public Property Let pnDocTpo(ByVal vNewValue As TpoDoc)
lnDocTpo = vNewValue
End Property

Public Property Get pnMonto() As Currency
pnMonto = lnMonto
End Property

Public Property Let pnMonto(ByVal vNewValue As Currency)
lnMonto = vNewValue
End Property

Public Property Get pnMotivo() As Long
pnMotivo = lnMotivo
End Property

Public Property Let pnMotivo(ByVal vNewValue As Long)
lnMotivo = vNewValue
End Property

Public Property Get psObjetoCodPadre() As String
psObjetoCodPadre = lsObjetoCodPadre
End Property

Public Property Let psObjetoCodPadre(ByVal vNewValue As String)
lsObjetoCodPadre = vNewValue
End Property

Public Property Get psObjetoCod() As String
psObjetoCod = lsObjetoCod
End Property

Public Property Let psObjetoCod(ByVal vNewValue As String)
lsObjetoCod = vNewValue
End Property

Private Sub txtBuscarDet_EmiteDatos()
lblDetalleDesc = txtBuscarDet.psDescripcion
txtImporte.SetFocus
End Sub
Private Sub txtBuscarDet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtImporte.SetFocus
End If
End Sub

Private Sub txtBuscarMotivo_EmiteDatos()
fraObjeto.Visible = True
lblMotivoDesc = txtBuscarMotivo.psDescripcion
txtBuscarObj = ""
lblObjetoDesc = ""
If txtBuscarMotivo <> "" Then
    txtBuscarObj.psRaiz = "Objetos"
    txtBuscarObj.rs = oDocRec.GetMotivosObjNivel(Val(txtBuscarMotivo))
    If Not txtBuscarObj.rs.EOF And Not txtBuscarObj.rs.BOF Then
        If txtBuscarObj.Enabled And txtBuscarObj.Visible Then txtBuscarObj.SetFocus
    Else
        fraObjeto.Visible = False
        If txtImporte.Visible And txtImporte.Enabled Then txtImporte.SetFocus
    End If
End If
End Sub

Private Sub txtBuscarMotivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If fraObjeto.Visible Then
        txtBuscarObj.SetFocus
    Else
        txtImporte.SetFocus
    End If
End If
End Sub

Private Sub txtBuscarObj_EmiteDatos()
Dim lsFiltro As String
lblObjetoDesc = txtBuscarObj.psDescripcion
txtBuscarDet = ""
lblDetalleDesc = ""
If txtBuscarObj <> "" Then
    txtBuscarDet.psRaiz = "Detalle Objetos"
    lsFiltro = oDocRec.GetFiltroMotivoObj(Val(txtBuscarMotivo), Trim(txtBuscarObj))
    txtBuscarDet.rs = oDocRec.GetDetalleObjetos(Val(txtBuscarObj), lsFiltro, 1)
End If
End Sub

Public Property Get pbOk() As Boolean
pbOk = lbOk
End Property

Public Property Let pbOk(ByVal vNewValue As Boolean)
lbOk = vNewValue
End Property

Private Sub txtBuscarObj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBuscarDet.SetFocus
End If
End Sub

Private Sub txtImporte_GotFocus()
fEnfoque txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtImporte, KeyAscii)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub

Private Sub txtImporte_LostFocus()
If Len(Trim(txtImporte)) = 0 Then txtImporte = 0
txtImporte = Format(txtImporte, "#,#0.00")
End Sub


