VERSION 5.00
Begin VB.Form frmPITCapOpeInterCajaConsultasEnvio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Saldos y Movimientos"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "frmPITCapOpeInterCajaConsultasEnvio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLector 
      Caption         =   "&Lector Tarjeta"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2760
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
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmPITCapOpeInterCajaConsultasEnvio.frx":030A
         Left            =   1200
         List            =   "frmPITCapOpeInterCajaConsultasEnvio.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   2535
      End
      Begin VB.TextBox txtCuenta 
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
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   3
         Top             =   840
         Width           =   3255
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
         TabIndex        =   12
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Comision :"
         Height          =   195
         Left            =   4080
         TabIndex        =   11
         Top             =   2100
         Width           =   720
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
         Left            =   1200
         TabIndex        =   10
         Top             =   1500
         Width           =   3570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Tarjeta: "
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label lblCuenta 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta :"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   960
         Width           =   600
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
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
         Height          =   675
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   2835
      End
      Begin VB.Label lblMoneda 
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   330
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   2760
      Width           =   975
   End
End
Attribute VB_Name = "frmPITCapOpeInterCajaConsultasEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nOperacion As Long
Dim fnMoneda As Integer

Dim sPersCodCMAC As String
Dim sNombreCMAC As String
Dim sDescOperacion As String

Dim fsPAN As String, fsTrack2 As String, fsPINBlock As String

Public Sub Inicia(ByVal psOpeCod As String, ByVal sDescOpe As String, _
        ByVal sCodCmac As String, ByVal sNomCmac As String, _
        Optional ByVal nComision As Double = 0)

    fsPAN = ""
    fsTrack2 = ""
    fsPINBlock = ""

    sDescOperacion = sDescOpe
    Me.Caption = "Captaciones - Operaciones InterCMACs - " & sDescOperacion
    nOperacion = CLng(psOpeCod)
    sPersCodCMAC = sCodCmac
    sNombreCMAC = Trim(sNomCmac)
    gsOpeCod = psOpeCod
    lblMensaje.Caption = sNombreCMAC & Chr$(13) & sDescOperacion


    Select Case psOpeCod
        Case "261003"
            lblCuenta.Visible = False
            txtCuenta.Visible = False
            
            cboMoneda.AddItem "TODAS" & Space(100) & "0"
        Case "261004"
            'lblMoneda.Visible = False
            'cboMoneda.Visible = False
    End Select
    
    lblComision.Caption = Format$(nComision, "#,##0.00")
    
    'cboMoneda.AddItem "TODAS" & Space(100) & "0"
    cboMoneda.AddItem "SOLES" & Space(100) & gMonedaNacional
    cboMoneda.AddItem "DOLARES" & Space(100) & gMonedaExtranjera
    Me.cboMoneda.ListIndex = 0
    
    Me.Show 1
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub cboMoneda_Click()
    fnMoneda = Right(cboMoneda.Text, 2)
End Sub


Private Sub CmdAceptar_Click()
Dim sCuenta As String
Dim lnComision As Currency


    sCuenta = Trim(txtCuenta.Text)
    lnComision = CDbl(lblComision.Caption)

    fnMoneda = Right(cboMoneda.Text, 2)

    If txtCuenta.Visible Then
        If sCuenta = "" Then
            MsgBox "Debe digitar un número de cuenta válido", vbInformation, "Aviso"
            txtCuenta.SetFocus
            Exit Sub
        End If
    End If
        
'    If CStr(fnMoneda) <> Mid(sCuenta, 9, 1) And nOperacion = "261004" Then
'        MsgBox "El tipo de moneda seleccionada es diferente al de la cuenta", vbInformation, "Aviso"
'        cboMoneda.SetFocus
'        Exit Sub
'    End If
    
    If fsPAN = "" Then
        MsgBox "Debe digitar el numero de tarjeta", vbInformation, "Aviso"
        cmdLector.SetFocus
        Exit Sub
    End If
    If fsPINBlock = "" Then
        MsgBox "Debe digitar su clave", vbInformation, "Aviso"
        cmdLector.SetFocus
        Exit Sub
    End If
    If fsTrack2 = "" Then
        MsgBox "Track2 esta vacio", vbInformation, "Aviso"
        cmdLector.SetFocus
        Exit Sub
    End If
        
    
    If MsgBox("Desea realizar la consulta??", vbQuestion + vbYesNo, "Aviso") = vbYes Then

        Call RegistrarOperacionInterCMAC(fsPAN, fsPINBlock, sCuenta, nOperacion, fsTrack2, fnMoneda, "", sPersCodCMAC, sLpt, sDescOperacion, sNombreCMAC, gdFecSis, gsCodAge, gsCodUser, 0, "", , False, 0, lnComision)
    
        Unload Me
    End If
End Sub

Private Sub cmdLector_Click()
    Call LectorTarjeta
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

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF10 And txtCuenta.Enabled Then
        Call LectorTarjeta
    End If
    
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
        Exit Sub
    End If
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub
