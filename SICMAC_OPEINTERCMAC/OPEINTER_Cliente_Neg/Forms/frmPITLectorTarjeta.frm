VERSION 5.00
Object = "{E828906B-5DC7-427F-ABFD-2B2352E1D6E5}#1.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmPITLectorTarjeta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lector de Tarjeta"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmPITLectorTarjeta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OCXTarjeta.CtrlTarjeta Tarjeta 
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   3360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
   End
   Begin VB.Frame fraTarjeta 
      Caption         =   " Datos Tarjeta "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox mskTarjeta 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   16
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.TextBox txtClave 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3840
         MaxLength       =   16
         TabIndex        =   13
         Text            =   "0000000000000000"
         Top             =   360
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.TextBox txtTrack2 
         Height          =   315
         Left            =   1080
         MaxLength       =   37
         TabIndex        =   12
         Text            =   "0000000000000000=00000000000000000000"
         Top             =   840
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "N° Tarjeta:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lblEtqDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Clave :"
         Height          =   195
         Left            =   3240
         TabIndex        =   15
         Top             =   420
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Track: "
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   510
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   390
      Left            =   4320
      TabIndex        =   9
      Top             =   3360
      Width           =   1485
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   2760
      TabIndex        =   10
      Top             =   3360
      Width           =   1485
   End
   Begin VB.Frame fraClave 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5895
      Begin VB.CommandButton cmdPedClave 
         Caption         =   "Pedir Clave"
         Height          =   360
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "Clave  :"
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblClave 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NO INGRESADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   255
         Width           =   3165
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5865
      Begin VB.CommandButton CmdLecTarj 
         Caption         =   "Leer Tarjeta"
         Height          =   345
         Left            =   4410
         TabIndex        =   2
         Top             =   255
         Width           =   1290
      End
      Begin VB.TextBox TxtNumTarj 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   795
         MaxLength       =   16
         TabIndex        =   1
         Top             =   225
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.Label lblNumTarjeta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   795
         TabIndex        =   4
         Top             =   240
         Width           =   3225
      End
      Begin VB.Label Label2 
         Caption         =   "Tarjeta :"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPITLectorTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private fsPAN As String
Private fsTrack As String
Private fsPINBlock As String
Private fsOpeCod As String
Private fbClaveIngresado As Boolean
Dim lIpPuertoPinVerifyPOS As String

Public Property Get PAN() As String
    PAN = fsPAN
End Property

Public Property Get TRACK() As String
    TRACK = fsTrack
End Property

Public Property Get pinblock() As String
    pinblock = fsPINBlock
End Property

Public Function Inicio(Optional ByVal psOpeCod As String = "") As String
    fsPAN = ""
    fsOpeCod = psOpeCod
    Me.Show 1
    
    Inicio = fsPAN
End Function

Private Sub cmdPedClave_Click()
    
    'lIpPuertoPinVerifyPOS = "192.168.15.35:81"
    
    lIpPuertoPinVerifyPOS = RecuperaIpPuertoPinVerifyPOS()
    
    If fsPAN = "" Then
        Call MsgBox("Necesita pasar su tarjeta por la lectora", vbCritical)
    Else
        fsPINBlock = ""
        'fsPINBlock = Tarjeta.PedirPinEnc(Me.lblNumTarjeta.Caption, gNMKPOS, gWKPOS)
        'fsPINBlock = Tarjeta.PITPedirPinEnc(Me.lblNumTarjeta.Caption, gNMKPOS, gWKPOS, gIpPuertoPinVerifyPOS, gCanalIdPOS, gCanalIdATM, gnTipoPinPad, gnPinPadPuerto)
        fsPINBlock = Tarjeta.PITPedirPinEnc(Me.lblNumTarjeta.Caption, gNMKPOS, gWKPOS, lIpPuertoPinVerifyPOS, gCanalIdPOS, gCanalIdATM, gnTipoPinPad, gnPinPadPuerto)
    
        If fsPINBlock <> "" And Len(fsPINBlock) = 16 Then
            Me.lblClave.Caption = "CLAVE INGRESADA"
            fbClaveIngresado = True
        Else
            Me.lblClave.Caption = "NO INGRESADO"
            fbClaveIngresado = False
        End If
    End If
End Sub


Private Sub CmdAceptar_Click()
Dim sResp As String
Dim i As Integer

    Call LectorTemporal 'Esto tiene que eliminarse
    
    If fsOpeCod = "" Then 'si viene sin codigo de operación no necesita validar PIN
        'sCodCta = Me.LstCta.SelectedItem.Text
        Unload Me
    Else 'Verificar que haya ingresado el PIN
        Select Case fsOpeCod
            Case "261001", "261004"
                If Not fbClaveIngresado Then
                    Call MsgBox("Error en la clave, reintente por favor, y si persiste el error comunicar al Area de Sistemas", vbInformation, "Aviso")
                    Me.lblClave.Caption = "NO INGRESADO"
                End If
        End Select
        Unload Me
    End If
    
End Sub







Private Sub mskTarjeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(mskTarjeta.Text)) = 16 Then
            txtTrack2.Text = Trim(mskTarjeta.Text) & Right(txtTrack2.Text, 21)
        End If
    End If
End Sub

Private Sub TxtNumTarj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            lblNumTarjeta.Caption = TxtNumTarj.Text
            TxtNumTarj.Visible = False
            Me.lblNumTarjeta.Visible = True
    End If
End Sub



Private Sub CmdLecTarj_Click()
Dim lsCard As String
    lsCard = Tarjeta.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto, gnTimeOutAg)
    lblNumTarjeta.Caption = Mid(lsCard, 2, 16)
    mskTarjeta.Text = lblNumTarjeta.Caption
    fsPAN = Mid(lsCard, 2, 16)
    fsTrack = Mid(lsCard, 2, 37)
End Sub

Private Sub cmdSalir_Click()
    fsPAN = ""
    Unload Me
End Sub

Sub LectorTemporal()
    If fsPAN = "" Then
        lblNumTarjeta.Caption = Replace(mskTarjeta.Text, "-", "")
        fsPAN = Replace(mskTarjeta.Text, "-", "")
        fsTrack = txtTrack2.Text
    End If
    
    If Not fbClaveIngresado Then
        fsPINBlock = TxtClave.Text
        lblClave.Caption = "CLAVE INGRESADA"
        fbClaveIngresado = True
    End If
End Sub

