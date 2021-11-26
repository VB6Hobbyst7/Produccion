VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMntSaldosMov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia de Cuentas Contables"
   ClientHeight    =   3555
   ClientLeft      =   2625
   ClientTop       =   2160
   ClientWidth     =   5625
   Icon            =   "frmMntSaldosMov.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "&Mostrar"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2880
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   2880
      Width           =   1275
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "&Cambiar"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2880
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   210
      TabIndex        =   10
      Top             =   1860
      Width           =   5235
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   345
         Left            =   780
         TabIndex        =   2
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFecha2 
         Height          =   345
         Left            =   3210
         TabIndex        =   3
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Caption         =   "AL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2760
         TabIndex        =   12
         Top             =   390
         Width           =   435
      End
      Begin VB.Label Label3 
         Caption         =   "DEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   210
      TabIndex        =   7
      Top             =   120
      Width           =   5235
      Begin Sicmact.TxtBuscar txtCtaContNew 
         Height          =   345
         Left            =   210
         TabIndex        =   1
         Top             =   1170
         Width           =   1665
         _ExtentX        =   2937
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
      Begin Sicmact.TxtBuscar txtCtaContAnt 
         Height          =   345
         Left            =   210
         TabIndex        =   0
         Top             =   480
         Width           =   1665
         _ExtentX        =   2937
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
      Begin VB.TextBox txtCtaDesNew 
         Height          =   345
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1170
         Width           =   3135
      End
      Begin VB.TextBox txtCtaDesAnt 
         Height          =   345
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta Nueva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   9
         Top             =   930
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   8
         Top             =   240
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmMntSaldosMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCta   As ADODB.Recordset
Dim clsCta  As DCtaCont

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If txtCtaContAnt.Text = txtCtaContNew.Text Then
    MsgBox "La transferencia debe ser entre Cuentas diferentes", vbInformation, "¡Aviso!"
    txtCtaContAnt.SetFocus
    Exit Function
End If
If txtCtaContAnt.Text = "" Then
   MsgBox "Falta especificar Cuenta a Cambiar...", vbInformation, "¡Aviso!"
   txtCtaContAnt.SetFocus
   Exit Function
End If
If txtCtaContNew.Text = "" Then
   MsgBox "Falta especificar Cuenta Destino...", vbInformation, "¡Aviso!"
   txtCtaContNew.SetFocus
   Exit Function
End If
If Trim(txtFecha) = "/  /" Or txtFecha2 = "/  /" Then
   MsgBox "Falta especificar fechas...", vbInformation, "Aviso"
   txtCtaContAnt.SetFocus
   Exit Function
End If
ValidaDatos = True
End Function

Private Sub cmdCambiar_Click()
On Error GoTo ErrCambio
If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox(" ¿ Seguro que desea realizar cambio de Cuenta Contable ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbNo Then
   Exit Sub
End If
Dim clsMov As New DMov
clsMov.ActualizaMovCta_Cuenta txtCtaContAnt.Text, txtCtaContNew.Text, Format(txtFecha, gsFormatoMovFecha), Format(txtFecha2, gsFormatoMovFecha)
Set clsMov = Nothing
MsgBox "Transferencia realizada satisfactoriamente", vbInformation, "¡Aviso!"
    
    'ARLO20170208
    Set objPista = New COMManejador.Pista
    gsOpeCod = LogMantSaldoProducto
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, "Se Transferio Cuenta Contable |Actual  " & txtCtaContAnt.Text & "|" & txtCtaDesAnt.Text & " a Cuenta Nueva :" & txtCtaContNew.Text & " |" & txtCtaDesNew _
    & " del " & txtFecha & " al " & txtFecha2
    Set objPista = Nothing
    '*******
txtCtaContAnt = ""
txtCtaContNew = ""
txtCtaDesAnt = ""
txtCtaDesNew = ""
txtCtaContAnt.SetFocus
Exit Sub
ErrCambio:
   MsgBox TextErr(Err.Description), vbInformation, "Aviso"

End Sub

Private Sub cmdMostrar_Click()
Dim clsImp As New NContImprimir
clsImp.Inicio gsNomCmac, gsCodAge, Format(gdFecSis, gsFormatoFechaView)
EnviaPrevio clsImp.ImprimeMovCta(Me.txtCtaContAnt, Format(txtFecha, gsFormatoMovFecha), Format(txtFecha2, gsFormatoMovFecha)), "Movimiento Cambiados", gnLinPage, False
Set clsImp = Nothing
cmdCambiar.SetFocus
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set clsCta = New DCtaCont
Set rsCta = clsCta.CargaCtaCont
txtCtaContAnt.rs = rsCta
txtCtaContAnt.lbUltimaInstancia = False
txtCtaContAnt.EditFlex = False
txtCtaContAnt.TipoBusqueda = BuscaGrid

txtCtaContNew.rs = rsCta
txtCtaContNew.lbUltimaInstancia = True
txtCtaContNew.EditFlex = False
txtCtaContNew.TipoBusqueda = BuscaGrid
CentraForm Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsCta
Set clsCta = Nothing
End Sub

Private Sub txtCtaContAnt_EmiteDatos()
txtCtaDesAnt = txtCtaContAnt.psDescripcion
If txtCtaDesAnt <> "" And txtCtaContNew.Visible Then
   txtCtaContNew.SetFocus
End If
End Sub

Private Sub txtCtaContAnt_GotFocus()
fEnfoque txtCtaContAnt
End Sub

Private Sub txtCtaContAnt_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtCtaContNew_EmiteDatos()
txtCtaDesNew.Text = txtCtaContNew.psDescripcion
If txtCtaDesNew <> "" And txtFecha.Visible Then
   txtFecha.SetFocus
End If
End Sub

Private Sub txtCtaContNew_GotFocus()
fEnfoque txtCtaContNew
End Sub

Private Sub txtCtaContNew_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFecha) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Aviso"
      txtFecha.SetFocus
   Else
      txtFecha2.SetFocus
   End If
End If
End Sub

Private Sub txtFecha_Validate(Cancel As Boolean)
   If ValidaFecha(txtFecha) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Aviso"
      Cancel = True
   End If
End Sub

Private Sub txtFecha2_GotFocus()
fEnfoque txtFecha2
End Sub

Private Sub txtFecha2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFecha2) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Aviso"
      txtFecha2.SetFocus
   Else
      cmdCambiar.SetFocus
   End If
End If
End Sub

Private Sub txtFecha2_Validate(Cancel As Boolean)
   If ValidaFecha(txtFecha2) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Aviso"
      Cancel = True
   End If
End Sub

