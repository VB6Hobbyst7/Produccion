VERSION 5.00
Begin VB.Form frmPersComentario 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1755
      TabIndex        =   4
      Top             =   4950
      Width           =   1100
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7830
      TabIndex        =   6
      Top             =   4950
      Width           =   1100
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6660
      TabIndex        =   5
      Top             =   4950
      Width           =   1100
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo Comentario"
      Height          =   375
      Left            =   90
      TabIndex        =   3
      Top             =   4950
      Width           =   1635
   End
   Begin VB.Frame fraComentario 
      Caption         =   "Nuevo Comentario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1275
      Left            =   90
      TabIndex        =   16
      Top             =   3600
      Width           =   8835
      Begin VB.TextBox txtComentario 
         Height          =   870
         Left            =   90
         MaxLength       =   235
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   8610
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1080
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   8865
      Begin VB.CommandButton CmdBuscar 
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
         Left            =   7530
         TabIndex        =   0
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   80
         TabIndex        =   15
         Top             =   315
         Width           =   570
      End
      Begin VB.Label lblDocNatural 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Identificación :"
         Height          =   195
         Left            =   80
         TabIndex        =   14
         Top             =   660
         Width           =   1425
      End
      Begin VB.Label lblDocJuridico 
         AutoSize        =   -1  'True
         Caption         =   "RUC:"
         Height          =   195
         Left            =   3045
         TabIndex        =   13
         Top             =   660
         Width           =   390
      End
      Begin VB.Label lblNomPers 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2760
         TabIndex        =   12
         Top             =   262
         Width           =   4620
      End
      Begin VB.Label lblDocNat 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1580
         TabIndex        =   11
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblDocJur 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4200
         TabIndex        =   10
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label LblPersCod 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   870
         TabIndex        =   9
         Top             =   262
         Width           =   1755
      End
   End
   Begin VB.Frame fraHistoria 
      Caption         =   "Historia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2220
      Left            =   90
      TabIndex        =   7
      Top             =   1305
      Width           =   8865
      Begin SICMACT.FlexEdit grdHistoria 
         Height          =   1905
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   3360
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Usu-Fecha-Comentario"
         EncabezadosAnchos=   "350-500-2000-5300"
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
         ColumnasAEditar =   "X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         Appearance      =   0
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmPersComentario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sPersCod As String

Private Sub LimpiaControles()
grdHistoria.Clear
grdHistoria.FormaCabecera
grdHistoria.Rows = 2
txtComentario.Text = ""
sPersCod = ""
fraComentario.Enabled = False
cmdgrabar.Enabled = False
cmdcancelar.Enabled = False
fraHistoria.Enabled = False
cmdNuevo.Enabled = False
lblPersCod.Caption = ""
lblNomPers.Caption = ""
lblDocnat.Caption = ""
lblDocJur.Caption = ""
End Sub

Private Sub BuscaComentarios(ByVal sPersona As String)
'Dim oCreditos As DCreditos
Dim oCreditos As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim sComentario As String, sFecha As String, sUsuario As String
Dim i As Integer

    'Set oCreditos = New DCreditos
    Set oCreditos = New COMDCredito.DCOMCredito
    Set R = oCreditos.DatosPosicionClienteComentarios(sPersona)
    Set oCreditos = Nothing
    i = 0
    Do While Not R.EOF
        sComentario = Trim(R("cComentario"))
        sFecha = Mid(R("cMovNro"), 7, 2) & "/" & Mid(R("cMovNro"), 5, 2) & "/" & Mid(R("cMovNro"), 1, 4)
        sFecha = sFecha & " " & Mid(R("cMovNro"), 9, 2) & ":" & Mid(R("cMovNro"), 11, 2) & ":" & Mid(R("cMovNro"), 13, 2)
        sUsuario = Right(R("cMovNro"), 4)
        i = i + 1
        grdHistoria.AdicionaFila
        grdHistoria.TextMatrix(i, 1) = sUsuario
        grdHistoria.TextMatrix(i, 2) = sFecha
        grdHistoria.TextMatrix(i, 3) = sComentario
        R.MoveNext
    Loop
    fraHistoria.Enabled = True
    cmdNuevo.Enabled = True
    cmdcancelar.Enabled = True
End Sub

Private Sub cmdBuscar_Click()
'Dim oPersona As UPersona
Dim oPersona As COMDPersona.UCOMPersona

Set oPersona = frmBuscaPersona.Inicio
If Not oPersona Is Nothing Then
    lblPersCod.Caption = oPersona.sPersCod
    lblNomPers.Caption = oPersona.sPersNombre
    lblDocnat.Caption = Trim(oPersona.sPersIdnroDNI)
    lblDocJur.Caption = Trim(oPersona.sPersIdnroRUC)
Else
    Exit Sub
End If

sPersCod = oPersona.sPersCod
Set oPersona = Nothing
Call BuscaComentarios(sPersCod)
    
End Sub

Private Sub CmdCancelar_Click()
If fraComentario.Enabled = True Then
    txtComentario.Text = ""
    fraComentario.Enabled = False
    cmdgrabar.Enabled = False
    cmdNuevo.Enabled = True
Else
    LimpiaControles
    cmdbuscar.SetFocus
End If

End Sub

Private Sub cmdGrabar_Click()
If Trim(txtComentario.Text) = "" Then
    MsgBox "Debe escribir un comentario válido", vbInformation, "Aviso"
    txtComentario.SetFocus
End If

If MsgBox("¿Desea Grabar el Nuevo Comentario?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim oCred As COMDCredito.DCOMCredito
    Set oCred = New COMDCredito.DCOMCredito
    
    Call oCred.Grabar_Comentarios(sPersCod, txtComentario.Text, gdFecSis, gsCodAge, gsCodUser)
    
    Set oCred = Nothing
    LimpiaControles
    cmdbuscar.SetFocus
End If
End Sub

Private Sub cmdNuevo_Click()
fraComentario.Enabled = True
txtComentario.Text = ""
txtComentario.SetFocus
cmdgrabar.Enabled = True
cmdNuevo.Enabled = False
cmdcancelar.Enabled = True
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Persona - Comentarios"
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub grdHistoria_RowColChange()
Dim nFila As Long
nFila = grdHistoria.Row
txtComentario.Text = Trim(grdHistoria.TextMatrix(nFila, 3))
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdgrabar.SetFocus
Else
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End If
End Sub
