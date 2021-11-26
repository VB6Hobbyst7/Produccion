VERSION 5.00
Begin VB.Form frmCapMovCancelacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Motivo de Cancelación"
   ClientHeight    =   2775
   ClientLeft      =   8385
   ClientTop       =   5295
   ClientWidth     =   3195
   Icon            =   "frmCapMovCancelacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2700
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2970
      Begin VB.CommandButton cmdContinuar 
         Caption         =   "&Continuar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   3
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtGlosa 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   750
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbMotivos 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmCapMovCancelacion.frx":030A
         Left            =   240
         List            =   "frmCapMovCancelacion.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblExtCmb 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Glosa:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCapMovCancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lnValida As Boolean
Private lnMotivo As Integer
Private lcGlosa As String
Property Get RegistraMotivo() As Boolean
    RegistraMotivo = lnValida
End Property
Property Get Motivo() As Integer
    Motivo = lnMotivo
End Property
Property Get Glosa() As String
    Glosa = lcGlosa
End Property
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Set oCons = New COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset
Set R = New ADODB.Recordset

Set R = oCons.ObtenerMovCancelacion
cmbMotivos.Clear

If Not (R.BOF And R.EOF) Then
 Do While Not R.EOF
     cmbMotivos.AddItem R!cConsDescripcion & space(100) & R!nConsValor
     R.MoveNext
 Loop
End If
Set R = Nothing
Set oCons = Nothing

End Sub
Public Sub Inicio()
    CargaControles
    Me.Show 1
End Sub

Public Sub GuardarMotivoCancelacion(ByVal pcCtaCod As String, pnMontivo As Integer, ByVal pcGlosa As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Call clsMant.GuardarMotivoCancelacion(pcCtaCod, pnMontivo, pcGlosa)
    Set clsMant = Nothing
End Sub

Private Sub cmbMotivos_Click()
    If Right(Trim(cmbMotivos), 1) = 4 Then
        txtGlosa.Enabled = True
    Else
        txtGlosa.Enabled = False
    End If
End Sub
Private Sub cmdContinuar_Click()
    lnValida = False
    If Trim(cmbMotivos) = "" Then
        MsgBox "Debe seleccionar el motivo de la cancelación.", vbInformation, "Aviso"
        cmbMotivos.SetFocus
        Exit Sub
    End If
    lnMotivo = CInt(Right(Trim(cmbMotivos), 1))
    lcGlosa = txtGlosa.Text
    lnValida = True
    Unload Me
End Sub

Private Sub cmdContinuar_KeyDown(KeyCode As Integer, Shift As Integer)
 If Trim(cmbMotivos) = "" Then
        MsgBox "Debe seleccionar el motivo de la cancelación.", vbInformation, "Aviso"
        cmbMotivos.SetFocus
        Exit Sub
    End If
End Sub
