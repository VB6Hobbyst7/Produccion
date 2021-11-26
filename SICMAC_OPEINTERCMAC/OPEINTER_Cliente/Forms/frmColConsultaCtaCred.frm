VERSION 5.00
Begin VB.Form FrmColConsultaCtaCred 
   Caption         =   "Consulta de Cuentas de Credito"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   Icon            =   "frmColConsultaCtaCred.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2655
      TabIndex        =   8
      Top             =   3150
      Width           =   1080
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1215
      TabIndex        =   7
      Top             =   3150
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Tarjeta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4695
      Begin VB.TextBox txtTarjeta 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtTrack2 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtClave 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Tarjeta:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Track2"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Clave:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   990
         Width           =   495
      End
   End
   Begin VB.TextBox txtDNI 
      Height          =   315
      Left            =   840
      MaxLength       =   11
      TabIndex        =   0
      Top             =   180
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "DNI:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   210
      Width           =   375
   End
End
Attribute VB_Name = "FrmColConsultaCtaCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOperacion As CaptacOperacion
Dim sPersCodCMAC As String
Dim sNombreCMAC As String
Dim sDescOperacion As String
Dim nmoneda As Moneda
Public Sub Inicia(ByVal nOpe As CaptacOperacion, ByVal sDescOpe As String, _
        ByVal sCodCmac As String, ByVal sNomCmac As String)

    sDescOperacion = sDescOpe
    Me.Caption = "Colocaciones - Operaciones InterCMACs - " & sDescOperacion
    nOperacion = nOpe
    sPersCodCMAC = sCodCmac
    sNombreCMAC = Trim(sNomCmac)
    lblMensaje = sNombreCMAC & Chr$(13) & sDescOperacion

    gsOpeCod = CStr(nOperacion)

    Me.Show 1
End Sub

Private Sub cmdAceptar_Click()

Dim sDNI As String
Dim sClave As String
Dim sNumTarj As String
Dim sMovNro As String
Dim sTrack2 As String
Dim clsFun As DFunciones.dFuncionesNeg
'Dim loLavDinero As SICMACMOPEINTER.frmMovLavDinero
'Set loLavDinero = New SICMACMOPEINTER.frmMovLavDinero

sDNI = Trim(txtDNI.Text)
sNumTarj = Trim(txtTarjeta.Text)
sClave = Trim(txtClave.Text)
sTrack2 = Trim(txtTrack2.Text)

If sDNI = "" Then
    MsgBox "Debe digitar un número de DNI válido", vbInformation, "Aviso"
    txtCuenta.SetFocus
    Exit Sub
End If
If sNumTarj = "" Then
    MsgBox "Debe digitar el numero de tarjeta", vbInformation, "Aviso"
    txtTarjeta.SetFocus
    Exit Sub
End If
If sClave = "" Then
    MsgBox "Debe digitar su clave", vbInformation, "Aviso"
    txtClave.SetFocus
    Exit Sub
End If
If sTrack2 = "" Then
    MsgBox "Track2 esta vacio", vbInformation, "Aviso"
    txtTrack2.SetFocus
    Exit Sub
End If

If MsgBox("Desea realizar la consulta??", vbQuestion + vbYesNo, "Aviso") = vbYes Then

    Set clsFun = New DFunciones.dFuncionesNeg

    sMovNro = clsFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

    Dim lsBoleta As String
    Dim nFicSal As Integer

    Call RegistrarOperacionInterCMAC(sNumTarj, sClave, "", nOperacion, sTrack2, nmoneda, sDNI, sPersCodCMAC, sMovNro, sLpt, sDescOperacion, sNombreCMAC, , , 0)
    
    'gVarPublicas.LimpiaVarLavDinero
    'Set loLavDinero = Nothing
    Set clsFun = Nothing
    Unload Me
End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
