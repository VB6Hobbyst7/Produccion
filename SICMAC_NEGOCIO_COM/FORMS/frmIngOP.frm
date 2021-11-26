VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIngOP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Ordenes de Pago ANTIGUAS"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "frmIngOP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Talonario ANTIGUO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2340
      Left            =   60
      TabIndex        =   8
      Top             =   960
      Width           =   5235
      Begin VB.TextBox TxtFin 
         Height          =   285
         Left            =   1905
         MaxLength       =   7
         TabIndex        =   2
         Top             =   915
         Width           =   1560
      End
      Begin VB.TextBox TxtIni 
         Height          =   285
         Left            =   1905
         MaxLength       =   7
         TabIndex        =   1
         Top             =   495
         Width           =   1560
      End
      Begin MSMask.MaskEdBox txtNro 
         Height          =   315
         Left            =   1905
         TabIndex        =   3
         Top             =   1395
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   1
         Format          =   "#"
         Mask            =   "#"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   1905
         TabIndex        =   4
         Top             =   1845
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         HideSelection   =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         Height          =   210
         Left            =   570
         TabIndex        =   12
         Top             =   1950
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Talonarios"
         Height          =   195
         Left            =   570
         TabIndex        =   11
         Top             =   1485
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin Serie"
         Height          =   195
         Left            =   570
         TabIndex        =   10
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio Serie"
         Height          =   195
         Left            =   570
         TabIndex        =   9
         Top             =   525
         Width           =   780
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   3960
      TabIndex        =   7
      Top             =   3360
      Width           =   1185
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   345
      Left            =   2640
      TabIndex        =   5
      Top             =   3360
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   915
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   5235
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   1110
         TabIndex        =   0
         Top             =   330
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmIngOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By capi 21012009
Option Explicit
Dim objPista As COMManejador.Pista
'End by


Private Sub CmdGrabar_Click()
Dim ssql As String, bexiste As Boolean
Dim rstemp As New ADODB.Recordset
Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
   
Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
   
   If TxtIni.Text = "" Or TxtFin.Text = "" Or val(txtNro.Text) = 0 Or Not IsDate(txtFecha.Text) Then
        MsgBox "La información registrada es incompleta, verifique los datos.", vbInformation, "AVISO"
        Exit Sub
   End If

   Set rstemp = oCap.ObtenerOrdenIngOp(txtCuenta.NroCuenta, TxtIni.Text, TxtFin.Text)
   If rstemp.State = 1 Then
        bexiste = IIf(rstemp.Fields(0).value > 0, True, False)
   End If
   
   If bexiste Then
        MsgBox "Este rango de Ordenes de Pago ya fueron ingresados para la cuenta."
   Else
        oCap.InsertaOrdenIngOp txtCuenta.NroCuenta, txtFecha.Text, TxtIni.Text, TxtFin.Text, txtNro.Text
        'By Capi 21012009
        objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, , txtCuenta.NroCuenta, gCodigoCuenta
        'End by


        txtCuenta.NroCuenta = ""
        TxtIni.Text = ""
        TxtFin.Text = ""
        txtNro.Text = "_"
        txtFecha.Text = "__/__/____"
    End If
    Set oCap = Nothing
    'rstemp.Close
End Sub


Private Sub cmdsalir_Click()
 Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nProducto As Producto
If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(nProducto, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            TxtIni.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
       SendKeys "{Tab}"
End If
'If KeyAscii = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
'        Dim sCuenta As String
'        Set clsGen = New DGeneral
'        sCuenta = frmValTarCodAnt.Inicia(nProducto, True)
'        If sCuenta <> "" Then
'            txtCuenta.NroCuenta = sCuenta
'            txtCuenta.SetFocusCuenta
'        End If
'    End If
End Sub

Private Sub Form_Load()
        SendKeys "{Tab}"
        SendKeys "{ENTER}"
        'By Capi 20012009
        Set objPista = New COMManejador.Pista
        gsOpeCod = gAhoRegTalonarioOrdPago
        'End By

End Sub

Private Sub TxtFin_KeyPress(KeyAscii As Integer)
'Or KeyAscii = 32
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
End If

End Sub

Private Sub TxtIni_KeyPress(KeyAscii As Integer)
'Or KeyAscii = 32
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
End If

End Sub

Private Sub txtNro_KeyPress(KeyAscii As Integer)
 If Chr(KeyAscii) = "0" Then
    KeyAscii = 0
 End If
End Sub
