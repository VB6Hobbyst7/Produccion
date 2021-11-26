VERSION 5.00
Begin VB.Form frmCapConsultaSaldos 
   Caption         =   "Consulta de Saldos y Movimientos"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5490
   Icon            =   "frmCapConsultaSaldos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   5490
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCuenta 
      Height          =   315
      Left            =   1020
      MaxLength       =   18
      TabIndex        =   2
      Top             =   720
      Width           =   3735
   End
   Begin VB.ComboBox cboMoneda 
      Height          =   315
      ItemData        =   "frmCapConsultaSaldos.frx":030A
      Left            =   1020
      List            =   "frmCapConsultaSaldos.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Frame frmDatTarj 
      Caption         =   "Datos Tarjeta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   5295
      Begin VB.TextBox txtClave 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtTrack2 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtTarjeta 
         Height          =   315
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   4
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Clave:"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Track2:"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Tarjeta:"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox txtGlosa 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   5295
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2981
      TabIndex        =   7
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1534
      TabIndex        =   0
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Glosa"
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
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblcuenta 
      Caption         =   "Cuenta"
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
      Left            =   120
      TabIndex        =   13
      Top             =   780
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Moneda :"
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
      Left            =   120
      TabIndex        =   12
      Top             =   270
      Width           =   675
   End
End
Attribute VB_Name = "frmCapConsultaSaldos"
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
    Me.Caption = "Captaciones - Operaciones InterCMACs - " & sDescOperacion
    nOperacion = nOpe
    sPersCodCMAC = sCodCmac
    sNombreCMAC = Trim(sNomCmac)
    lblMensaje = sNombreCMAC & Chr$(13) & sDescOperacion

    gsOpeCod = CStr(nOperacion)

    cboMoneda.AddItem "SOLES" & Space(100) & gMonedaNacional
    cboMoneda.AddItem "DOLARES" & Space(100) & gMonedaExtranjera
    'Me.cboMoneda.ListIndex = 0

    Me.Show 1
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub cboMoneda_Click()
    nmoneda = Right(cboMoneda.Text, 2)
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCuenta.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim sCuenta As String
Dim sClave As String
Dim sNumTarj As String
Dim sGlosa As String
Dim sMovNro As String
Dim sCtaAbono As String
Dim sTrack2 As String
Dim clsFun As DFunciones.dFuncionesNeg

sCuenta = Trim(txtCuenta.Text)
sNumTarj = Trim(txtTarjeta.Text)
sClave = Trim(txtClave.Text)
sTrack2 = Trim(txtTrack2.Text)


If sCuenta = "" And nOperacion = 260506 Then
    MsgBox "Debe digitar un número de cuenta válido", vbInformation, "Aviso"
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
        
    sGlosa = txtGlosa.Text

    sMovNro = clsFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

    Dim lsBoleta As String
    Dim nFicSal As Integer
       
    Call RegistrarOperacionInterCMAC(sNumTarj, sClave, sCuenta, nOperacion, sTrack2, nmoneda, "", sPersCodCMAC, sMovNro, sLpt, sDescOperacion, sNombreCMAC, , , 0, sGlosa)

    Set clsFun = Nothing
    Unload Me
End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


