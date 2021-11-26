VERSION 5.00
Begin VB.Form frmMntRepColumnasNuevo 
   ClientHeight    =   2160
   ClientLeft      =   1890
   ClientTop       =   3315
   ClientWidth     =   7335
   Icon            =   "frmMntRepColumnasNuevo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   7335
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2745
      TabIndex        =   3
      Top             =   1725
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4020
      TabIndex        =   4
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   90
      TabIndex        =   5
      Top             =   15
      Width           =   7080
      Begin Sicmact.TxtBuscar txtOpeCod 
         Height          =   315
         Left            =   1230
         TabIndex        =   0
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
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
         TipoBusqueda    =   2
         sTitulo         =   ""
      End
      Begin VB.CheckBox chkTotal 
         Caption         =   "&Columna Totalizada"
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
         Left            =   1230
         TabIndex        =   10
         Top             =   1170
         Width           =   1995
      End
      Begin VB.TextBox txtNroCol 
         Alignment       =   2  'Center
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
         Left            =   1230
         MaxLength       =   2
         TabIndex        =   2
         Top             =   720
         Width           =   555
      End
      Begin VB.TextBox txtDesCol 
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
         Left            =   3045
         MaxLength       =   60
         TabIndex        =   1
         Top             =   720
         Width           =   3825
      End
      Begin VB.Label lblDescOpe 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2595
         TabIndex        =   9
         Top             =   270
         Width           =   4275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº Columna "
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
         Left            =   135
         TabIndex        =   8
         Top             =   780
         Width           =   1065
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
         Left            =   1935
         TabIndex        =   7
         Top             =   780
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
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmMntRepColumnasNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOpeCod  As String
Dim sOpeDes  As String
Dim nNroCol  As Integer
Dim sColDes  As String
Dim lTotal   As Boolean
Dim lNuevo   As Boolean

Dim rsOpe As ADODB.Recordset


Public Sub Inicio(psOpeCod As String, psOpeDesc As String, pnNroCol As Integer, psColDes As String, plTotal As Boolean)
sOpeCod = psOpeCod
sOpeDes = psOpeDesc
nNroCol = pnNroCol
sColDes = psColDes
lTotal = plTotal
Me.Show 1
End Sub

Private Sub Form_Load()
CentraForm Me
lNuevo = False
If sColDes = "" Then
   lNuevo = True
End If
Me.Caption = "Operaciones: Mantenimiento: " & IIf(lNuevo, "Nuevo", "Modificar")
txtOpeCod.Text = sOpeCod
lblDescOpe.Caption = sOpeDes
txtNroCol.Text = nNroCol

If lNuevo Then
   txtOpeCod.Enabled = True
   txtNroCol.Enabled = True
Else
   txtDesCol = sColDes
   chkTotal.value = IIf(lTotal, 1, 0)
   txtOpeCod.Enabled = False
   txtNroCol.Enabled = False
End If

Dim clsOpe As New DOperacion
Set rsOpe = clsOpe.CargaOpeTpo("")
txtOpeCod.rs = rsOpe
txtOpeCod.TipoBusqueda = BuscaGrid
txtOpeCod.EditFlex = False
Set clsOpe = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsOpe
End Sub

Private Sub txtDesCol_GotFocus()
fEnfoque txtDesCol
End Sub

Private Sub txtDesCol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdAceptar.SetFocus
End If
End Sub

Private Sub txtNroCol_GotFocus()
fEnfoque txtNroCol
End Sub

Private Sub txtNroCol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.txtDesCol.SetFocus
End If
End Sub

Private Sub txtOpeCod_EmiteDatos()
lblDescOpe = txtOpeCod.psDescripcion
If lblDescOpe <> "" Then
    If txtNroCol.Enabled Then
        txtNroCol.SetFocus
    Else
        txtDesCol.SetFocus
    End If
End If
End Sub

Private Sub txtOpeCod_GotFocus()
fEnfoque txtOpeCod
End Sub

Private Sub txtOpeCod_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If Len(Trim(txtOpeCod.Text)) = 0 Then
   MsgBox "Falta ingresar Código de Operación", vbInformation, "¡Aviso!"
   txtOpeCod.SetFocus
   Exit Function
End If
If Len(Trim(txtNroCol)) = 0 Then
   MsgBox "Falta Ingresar Número de Columna", vbInformation, "¡Aviso!"
   txtNroCol.SetFocus
   Exit Function
End If
If Len(Trim(txtDesCol)) = 0 Then
   MsgBox "Falta Ingresar Descripción de Columna", vbInformation, "¡Aviso!"
   txtDesCol.SetFocus
   Exit Function
End If
ValidaDatos = True
End Function

Private Sub cmdAceptar_Click()
On Error GoTo AceptarErr
If Not ValidaDatos() Then
   Exit Sub
End If
   
If MsgBox(" ¿ Seguro que desea Grabar Datos ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbYes Then
   sOpeCod = txtOpeCod.Text
   nNroCol = txtNroCol.Text
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   
   Dim clsRep As New DRepCtaColumna
   If lNuevo Then
      clsRep.InsertaRepColumna sOpeCod, nNroCol, txtDesCol, chkTotal.value, gsMovNro
   Else
      clsRep.ActualizaRepColumna sOpeCod, nNroCol, txtDesCol, chkTotal.value, gsMovNro
   End If
   Set clsRep = Nothing
   glAceptar = True
Else
   glAceptar = False
End If
Unload Me
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub
Private Sub cmdCancelar_Click()
glAceptar = False
Unload Me
End Sub

Public Property Get pOpeCod() As String
pOpeCod = sOpeCod
End Property

Public Property Let pOpeCod(ByVal vNewValue As String)
sOpeCod = vNewValue
End Property


Public Property Get pNroCol() As Integer
pNroCol = nNroCol
End Property

Public Property Let pNroCol(ByVal vNewValue As Integer)
nNroCol = vNewValue
End Property
