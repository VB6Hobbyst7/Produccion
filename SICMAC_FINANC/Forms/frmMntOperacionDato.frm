VERSION 5.00
Begin VB.Form frmMntOperacionDato 
   ClientHeight    =   2955
   ClientLeft      =   2310
   ClientTop       =   5265
   ClientWidth     =   7785
   Icon            =   "frmMntOperacionDato.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2550
      TabIndex        =   6
      Top             =   2370
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   3990
      TabIndex        =   7
      Top             =   2370
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2205
      Left            =   90
      TabIndex        =   8
      Top             =   30
      Width           =   7635
      Begin VB.TextBox txtOpeGruDesc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2580
         MaxLength       =   60
         TabIndex        =   5
         Top             =   1620
         Width           =   4890
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "&Visible"
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
         Left            =   2580
         TabIndex        =   3
         Top             =   1170
         Width           =   945
      End
      Begin VB.TextBox txtOpeDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1245
         MaxLength       =   120
         TabIndex        =   1
         Top             =   690
         Width           =   6240
      End
      Begin VB.TextBox txtOpeCod 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1245
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox txtOpeNiv 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1245
         MaxLength       =   1
         TabIndex        =   2
         Top             =   1140
         Width           =   765
      End
      Begin Sicmact.TxtBuscar txtOpeTpo 
         Height          =   345
         Left            =   1245
         TabIndex        =   4
         Top             =   1620
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label4 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   12
         Top             =   1710
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Nivel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   11
         Top             =   1215
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
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
         TabIndex        =   10
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   9
         Top             =   315
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmMntOperacionDato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lpVisible As String
Dim lnOpeNiv As Integer
Dim cVisible As String, nOpeNiv As Integer
Dim sOpeCod  As String, sOpeDesc As String
Dim sOpeTpo  As String
Dim lNuevo   As Boolean

Dim rsGru  As ADODB.Recordset
'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(vOpeCod As String, vOpeDesc As String, vVisible As String, vOpeNiv As Integer, vOpeGruCod As String, plNuevo As Boolean)
cVisible = vVisible
nOpeNiv = vOpeNiv
sOpeCod = vOpeCod
sOpeDesc = vOpeDesc
sOpeTpo = vOpeGruCod
lNuevo = plNuevo
Me.Show 1
End Sub
Public Property Get pVisible() As String
pVisible = lpVisible
End Property
Public Property Let pVisible(ByVal vNewValue As String)
lpVisible = vNewValue
End Property
Public Property Get pOpeDesc() As String
pOpeDesc = sOpeDesc
End Property
Public Property Let pOpeDesc(ByVal vNewValue As String)
sOpeDesc = vNewValue
End Property

Public Property Get pOpeTpo() As String
pOpeTpo = sOpeTpo
End Property
Public Property Let pOpeTpo(ByVal vNewValue As String)
sOpeTpo = vNewValue
End Property

Private Sub chkVisible_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtOpeTpo.SetFocus
End If
End Sub

Private Sub Form_Load()
Me.Caption = "Operaciones:Mantenimiento:" & IIf(lNuevo, "Nuevo", "Modificar")
CentraForm Me
If lNuevo Then
   txtOpeCod.Enabled = True
Else
   txtOpeCod.Text = sOpeCod
   txtOpeDesc.Text = sOpeDesc
   txtOpeCod.Enabled = False
   txtOpeDesc.SelStart = 0
   txtOpeDesc.SelLength = Len(txtOpeDesc.Text)
End If
chkVisible.value = IIf(cVisible = "0", 0, 1)
txtOpeNiv = nOpeNiv

Dim clsOpe As DOperacion
Set clsOpe = New DOperacion
txtOpeTpo.rs = clsOpe.CargaOpeGru
Set clsOpe = Nothing

txtOpeTpo.EditFlex = False
txtOpeTpo.TipoBusqueda = BuscaArbol
txtOpeTpo.psRaiz = "Tipos de Operaciones"
If Not lNuevo Then
   txtOpeTpo.Text = sOpeTpo
   txtOpeGruDesc = txtOpeTpo.psDescripcion
End If
End Sub
Private Sub txtOpeCod_Validate(Cancel As Boolean)
Cancel = False
If Len(Trim(txtOpeCod.Text)) < 6 Then
   MsgBox "El código de operación está incompleto...", vbCritical, "Aviso!"
   Cancel = True
End If
End Sub

Private Sub txtOpeCod_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtOpeDesc.SetFocus
End If
End Sub
Private Sub cmdAceptar_Click()
If Len(Trim(txtOpeDesc.Text)) <> 0 And Len(Trim(txtOpeCod.Text)) <> 0 Then
   If Len(Trim(txtOpeCod.Text)) < 6 Then
      MsgBox "El código de operación está incompleto...", vbCritical, "Advertencia!"
      txtOpeCod.SetFocus
      glAceptar = False
      Exit Sub
   End If
  
   If MsgBox(" ¿ Esta seguro que desea grabar ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
      gsOpeCod = txtOpeCod.Text
      sOpeDesc = txtOpeDesc.Text
      sOpeTpo = txtOpeTpo.Text
      lpVisible = IIf(chkVisible.value, "1", "0")
      lnOpeNiv = txtOpeNiv
      glAceptar = True
   Else
      glAceptar = False
   End If
   
Else
   MsgBox "No se pueden grabar campos en blanco...", vbCritical, "Aviso"
   glAceptar = False
   Exit Sub
End If
Unload Me
End Sub
Private Sub cmdCancelar_Click()
glAceptar = False
Unload Me
End Sub

Private Sub txtOpeDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtOpeNiv.SetFocus
End If
End Sub

Private Sub txtOpeDesc_Validate(Cancel As Boolean)
Cancel = False
If Len(Trim(txtOpeCod.Text)) = 0 Then
   Cancel = True
End If
End Sub

Private Sub txtOpeNiv_GotFocus()
fEnfoque txtOpeNiv
End Sub

Private Sub txtOpeNiv_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   chkVisible.SetFocus
End If
End Sub

Public Property Get pOpeNiv() As Integer
pOpeNiv = lnOpeNiv
End Property

Public Property Let pOpeNiv(ByVal vNewValue As Integer)
lnOpeNiv = vNewValue
End Property

Private Sub txtOpeTpo_EmiteDatos()
txtOpeGruDesc = txtOpeTpo.psDescripcion
If txtOpeGruDesc <> "" And cmdAceptar.Visible Then
   cmdAceptar.SetFocus
End If
End Sub
