VERSION 5.00
Begin VB.Form frmPITColOpeInterCajaConsultasEnvio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Cuentas de Credito"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   Icon            =   "frmPITColOpeInterCajaConsultasEnvio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLector 
      Caption         =   "&Lector Tarjeta"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Datos Cuenta"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6765
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmPITColOpeInterCajaConsultasEnvio.frx":030A
         Left            =   1140
         List            =   "frmPITColOpeInterCajaConsultasEnvio.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtDNI 
         Height          =   315
         Left            =   1140
         MaxLength       =   11
         TabIndex        =   4
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Comision :"
         Height          =   195
         Left            =   3960
         TabIndex        =   12
         Top             =   2580
         Width           =   720
      End
      Begin VB.Label lblComision 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4920
         TabIndex        =   11
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
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
         Left            =   240
         TabIndex        =   10
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Tarjeta: "
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1980
         Width           =   930
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1140
         TabIndex        =   7
         Top             =   1920
         Width           =   3435
      End
      Begin VB.Image imagen 
         Height          =   480
         Index           =   0
         Left            =   6240
         Picture         =   "frmPITColOpeInterCajaConsultasEnvio.frx":030E
         Top             =   840
         Width           =   480
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
         TabIndex        =   5
         Top             =   1110
         Width           =   375
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   795
         Left            =   3720
         TabIndex        =   3
         Top             =   120
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   3390
      Width           =   1080
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4095
      TabIndex        =   0
      Top             =   3390
      Width           =   1080
   End
End
Attribute VB_Name = "frmPITColOpeInterCajaConsultasEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nOperacion As Long
Dim sPersCodCMAC As String
Dim sNombreCMAC As String
Dim sDescOperacion As String

Dim fsPAN As String, fsTrack2 As String, fsPINBlock As String

Public Sub Inicia(ByVal nOpe As CaptacOperacion, ByVal sDescOpe As String, _
        ByVal sCodCmac As String, ByVal sNomCmac As String, _
        Optional ByVal nComision As Double = 0)

    fsPAN = ""
    fsTrack2 = ""
    fsPINBlock = ""

    sDescOperacion = sDescOpe
    Me.Caption = "Colocaciones - Operaciones InterCMACs - " & sDescOperacion
    nOperacion = nOpe
    sPersCodCMAC = sCodCmac
    sNombreCMAC = Trim(sNomCmac)
    lblMensaje.Caption = sNombreCMAC & Chr$(13) & sDescOperacion

    gsOpeCod = CStr(nOperacion)
    
    lblComision.Caption = Format$(nComision, "#,##0.00")
    
    cboMoneda.AddItem "TODAS" & Space(100) & "0"
    cboMoneda.AddItem "SOLES" & Space(100) & gMonedaNacional
    cboMoneda.AddItem "DOLARES" & Space(100) & gMonedaExtranjera
    Me.cboMoneda.ListIndex = 0
    
    Me.Show 1
End Sub

Private Sub CmdAceptar_Click()
Dim lsDNI As String
Dim lnMoneda As Integer
Dim lnComision As Currency

    lsDNI = Trim(txtDNI.Text)
    lnMoneda = Right(cboMoneda.Text, 2)
    lnComision = CDbl(lblComision.Caption)
    
    
    If lsDNI = "" Then
        MsgBox "Debe digitar un número de DNI válido", vbInformation, "Aviso"
        txtDNI.SetFocus
        Exit Sub
    End If
    If fsPAN = "" Then
        fsPAN = "0000000000000000"
        fsPINBlock = "0000000000000000"
        fsTrack2 = "0000000000000000=00000000000000000000"
        'MsgBox "Debe digitar el numero de tarjeta", vbInformation, "Aviso"
        'cmdLector.SetFocus
        'Exit Sub
    End If
    'If fsPINBlock = "" Then
    '    MsgBox "Debe digitar su clave", vbInformation, "Aviso"
    '    cmdLector.SetFocus
    '    Exit Sub
    'End If
    'If fsTrack2 = "" Then
    '    MsgBox "Track2 esta vacio", vbInformation, "Aviso"
    '    cmdLector.SetFocus
    '    Exit Sub
    'End If
    
    If MsgBox("Desea realizar la consulta??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call RegistrarOperacionInterCMAC(fsPAN, fsPINBlock, "", nOperacion, fsTrack2, lnMoneda, lsDNI, sPersCodCMAC, sLpt, sDescOperacion, sNombreCMAC, gdFecSis, gsCodAge, gsCodUser, , , , False, 0, lnComision)
        Unload Me
    End If

End Sub

Private Sub cmdLector_Click()
    Call LectorTarjeta
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF10 Then
        Call LectorTarjeta
    End If
End Sub

Sub LectorTarjeta()
Dim objLector As New frmPITLectorTarjeta
    fsPAN = objLector.Inicio(CStr(nOperacion))
    If fsPAN <> "" Then
        fsTrack2 = objLector.TRACK
        fsPINBlock = objLector.pinblock
    End If
    Set objLector = Nothing
    lblTarjeta.Caption = getTarjetaFormateado(fsPAN)
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
        Exit Sub
    End If
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub
