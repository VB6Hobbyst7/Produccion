VERSION 5.00
Begin VB.Form frmMntSaldosIniTransf 
   Caption         =   "Saldos Iniciales: Transferencia"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   Icon            =   "frmMntSaldosIniTransf.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Cuenta Destino"
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
      Height          =   1275
      Left            =   60
      TabIndex        =   11
      Top             =   1530
      Width           =   7185
      Begin VB.TextBox txtDestinoDesc 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2220
         TabIndex        =   3
         Top             =   300
         Width           =   4770
      End
      Begin Sicmact.TxtBuscar txtDestino 
         Height          =   345
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   2055
         _ExtentX        =   3625
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
         sTitulo         =   ""
      End
      Begin VB.Label Label8 
         Caption         =   "Moneda Extranjera"
         Height          =   390
         Left            =   4710
         TabIndex        =   17
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblDImporteME 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5580
         TabIndex        =   16
         Top             =   780
         Width           =   1365
      End
      Begin VB.Label lblFechaAlD 
         AutoSize        =   -1  'True
         Caption         =   "Saldos al  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   810
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Moneda Nacional"
         Height          =   390
         Left            =   2280
         TabIndex        =   14
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblDImporte 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         TabIndex        =   13
         Top             =   780
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdAceptaTrans 
      Caption         =   "&Transferir"
      Height          =   345
      Left            =   2287
      TabIndex        =   4
      Top             =   2970
      Width           =   1290
   End
   Begin VB.CommandButton cmdCanTr 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3727
      TabIndex        =   5
      Top             =   2970
      Width           =   1290
   End
   Begin VB.Frame FraTransfer 
      Caption         =   "Cuenta Origen"
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
      Height          =   1410
      Left            =   60
      TabIndex        =   6
      Top             =   30
      Width           =   7185
      Begin Sicmact.TxtBuscar txtOrigen 
         Height          =   345
         Left            =   180
         TabIndex        =   0
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
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
         Enabled         =   0   'False
         Enabled         =   0   'False
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin VB.TextBox TxtOrigenDesc 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2220
         TabIndex        =   1
         Top             =   360
         Width           =   4770
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda Extranjera"
         Height          =   390
         Left            =   4710
         TabIndex        =   12
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblOImporteME 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5580
         TabIndex        =   10
         Top             =   900
         Width           =   1365
      End
      Begin VB.Label lblFechaAlO 
         AutoSize        =   -1  'True
         Caption         =   "Saldos al  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   930
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda Nacional"
         Height          =   390
         Left            =   2280
         TabIndex        =   8
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblOImporte 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3090
         TabIndex        =   7
         Top             =   900
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmMntSaldosIniTransf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCta As ADODB.Recordset
Dim sDestino As String

Dim sCtaCod  As String
Dim sFecha   As String
Dim sFechaAl As String
Dim nImporte As Currency
Dim nImporteME As Currency

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(psCtaCod As String, psFecha As String, pnImporte As Currency, pnImporteME As Currency, psFechaAl As String)
sCtaCod = psCtaCod
sFecha = psFecha
sFechaAl = psFechaAl
nImporte = pnImporte
nImporteME = pnImporteME
Me.Show 1
End Sub

Private Sub cmdCanTr_Click()
glAceptar = False
Unload Me
End Sub

Private Sub Form_Activate()
Me.TxtOrigenDesc = txtOrigen.psDescripcion
End Sub

Private Sub Form_Load()
Dim clsCta As New DCtaCont
Set rsCta = clsCta.CargaCtaCont("cCtaContCod LIKE '__[126]%'", gsCentralCom & "CtaCont", adLockReadOnly)
Set clsCta = Nothing

CentraForm Me
txtOrigen.rs = rsCta
txtOrigen.TipoBusqueda = BuscaGrid
txtOrigen.lbUltimaInstancia = False
txtOrigen.Text = sCtaCod
lblOImporte = Format(nImporte, gsFormatoNumeroView)
lblOImporteME = Format(nImporteME, gsFormatoNumeroView)

txtDestino.rs = rsCta
txtDestino.TipoBusqueda = BuscaGrid
txtDestino.lbUltimaInstancia = True
lblFechaAlO = lblFechaAlO & " " & sFechaAl
lblFechaAlD = lblFechaAlD & " " & sFechaAl
End Sub

Private Sub cmdAceptaTrans_Click()
Dim nMontoS  As Currency
Dim nTpoCam  As Double
Dim sCta     As String
Dim sFechaCta As String
   If MsgBox(" ¿ Esta Seguro que desea realizar Transferencia ? ", vbInformation + vbYesNo, "Aviso") = vbNo Then
       Exit Sub
   End If
   If Not ValidaDatos Then
       Exit Sub
   End If
   nTpoCam = LeeTpoCambio(sFecha, TCFijoDia)
   sDestino = txtDestino.Text
   
   Dim clsSdo As New DCtaSaldo
   Dim prs    As ADODB.Recordset
   Set prs = clsSdo.CargaCtaSaldo(sDestino, sFechaAl)
   If prs.EOF Then
     clsSdo.InsertaCtaSaldo sDestino, sFecha, nImporte, nImporteME
   Else
     sFechaCta = Format(prs!dCtaSaldofecha, gsFormatoFecha)
     clsSdo.ActualizaCtaSaldo sDestino, sFechaCta, nImporte + CCur(Format(lblDImporte, gsFormatoNumeroDato)), nImporteME + CCur(Format(lblDImporteME, gsFormatoNumeroDato))
   End If
   
   Set prs = clsSdo.CargaCtaSaldo(sDestino, sFechaAl)
   If prs.EOF Then
    clsSdo.InsertaCtaSaldo txtOrigen, sFecha, 0, 0
   Else
    clsSdo.ActualizaCtaSaldo txtOrigen, sFecha, 0, 0
   End If
    
    'ARLO20170208
    Set objPista = New COMManejador.Pista
    gsOpeCod = LogMantSaldoProducto
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se Transferio Saldo |Origen  " & txtOrigen.Text & "|" & TxtOrigenDesc.Text & "|Monto Soles " & lblOImporte & " |Monto Dolares " & lblOImporteME & _
    " a Destino" & txtDestino & " |" & txtDestinoDesc & "|Monto Soles " & lblDImporte & " |Monto Dolares " & lblDImporteME
    Set objPista = Nothing
    '*******
   Set clsSdo = Nothing
   RSClose prs
   glAceptar = True
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsCta
End Sub

Private Sub txtDestino_EmiteDatos()
Dim prs As ADODB.Recordset
Dim clsSdo As New DCtaSaldo
   If txtDestino.psDescripcion <> "" Then
      txtDestinoDesc.Text = txtDestino.psDescripcion
      Set prs = clsSdo.CargaCtaSaldo(txtDestino, sFechaAl)
      If Not prs.EOF Then
         lblDImporte = Format(prs!nCtaSaldoImporte, gsFormatoNumeroView)
         lblDImporteME = Format(0, gsFormatoNumeroView)
      End If
      cmdAceptaTrans.SetFocus
   End If
   Set clsSdo = Nothing
End Sub

Private Sub TxtDestino_GotFocus()
   fEnfoque txtDestino
End Sub

Private Sub txtDestino_KeyPress(KeyAscii As Integer)
   KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtOrigen_EmiteDatos()
   If txtOrigen.psDescripcion <> "" Then
      TxtOrigenDesc.Text = txtOrigen.psDescripcion
      If txtDestino.Visible Then
         txtDestino.SetFocus
      End If
   End If
End Sub

Public Property Get psDestino() As String
psDestino = sDestino
End Property

Public Property Let psDestino(ByVal vNewValue As String)
sDestino = vNewValue
End Property

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If txtDestino.Text = "" Then
   MsgBox "Falta indicar Cuenta destino de Transferencia", vbInformation, "¡Aviso!"
   txtDestino.SetFocus
End If
ValidaDatos = True
End Function
