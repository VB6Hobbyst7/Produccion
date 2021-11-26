VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmContabMayor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad: Libro Mayor"
   ClientHeight    =   3735
   ClientLeft      =   1050
   ClientTop       =   3405
   ClientWidth     =   6045
   Icon            =   "frmContabMayor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Impresión"
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
      Height          =   1215
      Left            =   4320
      TabIndex        =   16
      Top             =   60
      Width           =   1695
      Begin VB.OptionButton optImpresion 
         Caption         =   "Resumido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optImpresion 
         Caption         =   "Sin Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optImpresion 
         Caption         =   "Detallado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CheckBox chkResDet 
      Caption         =   "Imprime en forma resumida"
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CheckBox chkOperacion 
      Caption         =   "Todas las Operaciones"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   210
      TabIndex        =   13
      Top             =   2880
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operación"
      Enabled         =   0   'False
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
      Height          =   885
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   5835
      Begin Sicmact.TxtBuscar txtOpeCod 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1545
         _ExtentX        =   2725
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
      End
      Begin VB.TextBox txtOpeDes 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   4005
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Agencia"
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
      Height          =   885
      Left            =   90
      TabIndex        =   11
      Top             =   1320
      Width           =   5835
      Begin VB.CheckBox chkAgencia 
         Caption         =   "Todas las Agencias"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   585
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin Sicmact.TxtBuscar txtAgeCod 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1545
         _ExtentX        =   2725
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
      End
      Begin VB.TextBox txtAgeDes 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   4005
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4560
      TabIndex        =   7
      Top             =   3240
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   3300
      TabIndex        =   6
      Top             =   3240
      Width           =   1200
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rango de Fechas"
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
      Height          =   855
      Left            =   90
      TabIndex        =   8
      Top             =   240
      Width           =   4170
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   345
         Left            =   660
         TabIndex        =   0
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   345
         Left            =   2760
         TabIndex        =   1
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DEL"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   390
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "AL"
         Height          =   195
         Left            =   2370
         TabIndex        =   9
         Top             =   390
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmContabMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs   As New ADODB.Recordset
Dim sSql As String
Dim sAge As String
Dim sFec As String
Dim sTpo As String
Dim WithEvents oImp As NContImprimir
Attribute oImp.VB_VarHelpID = -1
Dim oBarra As New clsProgressBar

Private Sub chkOperacion_Click()
    If chkOperacion.value = 1 Then
        txtOpeCod.Text = ""
        txtOpeDes.Text = ""
        txtOpeCod.Enabled = False
        txtOpeDes.Enabled = False
    Else
        txtOpeCod.Enabled = True
        txtOpeDes.Enabled = True
    End If
End Sub

Private Sub chkAgencia_Click()
    If chkAgencia.value = 1 Then
        txtAgeCod.Text = ""
        txtAgeDes.Text = ""
        txtAgeCod.Enabled = False
        txtAgeDes.Enabled = False
    Else
        txtAgeCod.Enabled = True
        txtAgeDes.Enabled = True
    End If
End Sub

Private Sub cmdImprimir_Click()
Dim sImpre As String
On Error GoTo ImprimeErr
Screen.MousePointer = 11
Me.Enabled = False
Set oImp = New NContImprimir
'Modfificado PASI TI-ERS086-2014
'sImpre = oImp.ImprimeLibroMayor(CDate(txtFechaDel), CDate(txtFechaAl), txtAgeCod, txtOpeCod, gnLinPage, Me.chkResDet.value)
sImpre = oImp.ImprimeLibroMayor(CDate(txtFechaDel.Text), CDate(txtFechaAl.Text), txtAgeCod.Text, txtOpeCod.Text, gnLinPage, OptionTrueImpresion(Me.optImpresion))
'end PASI
Me.Enabled = True
Screen.MousePointer = 0
   Select Case MsgBox("Desea Enviar el Reporte a la Impresora", vbInformation + vbYesNoCancel, "Aviso")
      Case vbNo
         If sImpre <> "" Then
            EnviaPrevio sImpre, "LIBRO MAYOR", gnLinPage, False
         Else
            MsgBox "No existe información.", vbInformation, "Aviso"
         End If
         
      Case vbYes
         EnviaImpresion sImpre, gnLinPage, False
   End Select
   Exit Sub
ImprimeErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
   Me.Enabled = True
   Screen.MousePointer = 0
End Sub
'PASI TI-ERS086-2014
Private Function OptionTrueImpresion(Botones As Object) As Integer
    'Recibe matriz Option Button y devuelve cual esta activado o -1 si ninguno
    Dim Boton As Object
    For Each Boton In Botones
        If Boton.value Then
            OptionTrueImpresion = Boton.Index
            Exit Function
        End If
    Next
    OptionTrueImpresion = -1
End Function
'end PASI
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim oAge As New DActualizaDatosArea
Dim oOpe As New DOperacion
frmReportes.Enabled = False
CentraForm Me
txtAgeCod.rs = oAge.GetAgencias(, False)
txtOpeCod.rs = oOpe.CargaOpeGru()

If gsOpeCod = "760002" Then '*** PEAC 20110406
    chkResDet.Visible = True
Else
    chkResDet.Visible = False
End If

'PASI TI-ERS086-2014
    chkResDet.Enabled = False
    chkResDet.Visible = False
'end PASI

Set oAge = Nothing
Set oOpe = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oImp = Nothing
Set oBarra = Nothing
frmReportes.Enabled = True
End Sub

Private Sub oImp_BarraClose()
oBarra.CloseForm Me
End Sub

Private Sub oImp_BarraProgress(value As Variant, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
oBarra.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Sub

Private Sub oImp_BarraShow(pnMax As Variant)
oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = pnMax
End Sub

Private Sub txtAgeCod_EmiteDatos()
txtAgeDes = txtAgeCod.psDescripcion
If txtAgeCod.Text <> "" Then
   'txtOpeCod.SetFocus
End If
End Sub

Private Sub txtFechaAl_Validate(Cancel As Boolean)
If ValidaFecha(txtFechaAl) <> "" Then
   Cancel = True
End If
End Sub

Private Sub txtFechaDel_GotFocus()
txtFechaDel.SelStart = 0
txtFechaDel.SelLength = Len(txtFechaDel)
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFechaDel) <> "" Then
      MsgBox "Fecha no válida...", vbInformation, "Aviso"
   Else
      txtFechaAl.SetFocus
   End If
End If
End Sub

Private Sub txtFechaAl_GotFocus()
txtFechaAl.SelStart = 0
txtFechaAl.SelLength = Len(txtFechaAl)
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFechaAl) <> "" Then
      MsgBox "Fecha no Válida...", vbInformation, "Aviso"
      Exit Sub
   End If
   If txtAgeCod.Enabled = True Then
        txtAgeCod.SetFocus
   Else
        chkAgencia.SetFocus
   End If
    
End If
End Sub

Private Sub txtFechaDel_Validate(Cancel As Boolean)
If ValidaFecha(txtFechaDel) <> "" Then
   Cancel = True
End If
End Sub

Private Sub txtOpeCod_EmiteDatos()
txtOpeDes = txtOpeCod.psDescripcion
If txtOpeCod.Text <> "" Then
   CmdImprimir.SetFocus
End If
End Sub

