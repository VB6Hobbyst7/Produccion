VERSION 5.00
Begin VB.Form frmPITColOpeInterCajaEnvio 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "frmPITColOpeInterCajaEnvio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLector 
      Caption         =   "&Lector Tarjeta"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame fraMovimiento 
      Caption         =   "Movimiento"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   6735
      Begin VB.Frame fraGlosa 
         Caption         =   "Glosa"
         Height          =   1245
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3525
         Begin VB.TextBox txtGlosa 
            Height          =   885
            Left            =   75
            TabIndex        =   13
            Top             =   240
            Width           =   3330
         End
      End
      Begin VB.Frame fraMonto 
         Height          =   1245
         Left            =   3720
         TabIndex        =   9
         Top             =   240
         Width           =   2850
         Begin VB.TextBox txtMonto 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Left            =   1080
            MaxLength       =   14
            TabIndex        =   10
            Top             =   360
            Width           =   1575
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
            Left            =   1080
            TabIndex        =   18
            Top             =   810
            Width           =   1575
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Comision :"
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
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
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
            Left            =   120
            TabIndex        =   11
            Top             =   420
            Width           =   660
         End
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   3720
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   3720
      Width           =   990
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
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   6765
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
         MaxLength       =   20
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmPITColOpeInterCajaEnvio.frx":030A
         Left            =   1200
         List            =   "frmPITColOpeInterCajaEnvio.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Tarjeta: "
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1200
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
         TabIndex        =   15
         Top             =   1140
         Width           =   3555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   330
         Width           =   585
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
         TabIndex        =   6
         Top             =   120
         Width           =   2955
      End
      Begin VB.Image imagen 
         Height          =   480
         Index           =   0
         Left            =   6240
         Picture         =   "frmPITColOpeInterCajaEnvio.frx":030E
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   720
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmPITColOpeInterCajaEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nOperacion As Long
Dim sOpeDesc As String
Dim sPersCodCMAC As String
Dim sNombreCMAC As String
Dim nmoneda As Integer
Dim fsPAN As String, fsTrack2 As String, fsPINBlock As String

Public Sub Inicio(ByVal pnOpeCod As Long, ByVal psOpeDesc As String, _
        ByVal psPersCodCMAC As String, ByVal psNomCmac As String, _
        Optional nComision As Double = 0)

    fsPAN = ""
    fsTrack2 = ""
    fsPINBlock = ""

    nOperacion = pnOpeCod
    sOpeDesc = psOpeDesc
    sPersCodCMAC = psPersCodCMAC
    sNombreCMAC = psNomCmac
    lblMensaje = sNombreCMAC & Chr$(13) & sOpeDesc

    Me.Caption = "Créditos CMAC Llamada - " & psOpeDesc
    gsOpeCod = CStr(pnOpeCod)

    txtMonto.Text = 0#
    txtMonto.Text = Format$(txtMonto, "#,##0.00")
    
    lblComision.Caption = Format$(nComision, "#,##0.00")

    cboMoneda.AddItem "SOLES" & Space(100) & gMonedaNacional
    cboMoneda.AddItem "DOLARES" & Space(100) & gMonedaExtranjera

    Me.Show 1
End Sub

Private Sub cboMoneda_Change()
Dim vNComision As Double
Dim vMonto As Double

    If Val(txtMonto.Text) = 0 Then
        cmdGrabar.Enabled = False
    Else
        vMonto = GetLimiteMonto(nmoneda)
        cmdGrabar.Enabled = True
    End If

End Sub

Private Sub cboMoneda_Click()
    nmoneda = Right(cboMoneda.Text, 2)
    txtMonto.BackColor = IIf(nmoneda = gMonedaNacional, &HC0FFFF, &HC0FFC0)

    If Val(txtMonto.Text) = 0 Then
        cmdGrabar.Enabled = False
    Else
        cmdGrabar.Enabled = True
    End If

End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCuenta.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
Dim lsCuenta As String
Dim lnMonto As Currency
Dim lnMoneda As Integer
Dim lsIFTipo As String
Dim lsGlosa As String
Dim lnComision As Currency

    lsCuenta = Trim(txtCuenta)
    
    lnMonto = Val(txtMonto.Text)
    lnComision = CDbl(lblComision.Caption)

    If lsCuenta = "" Then
        MsgBox "Debe digitar un número de cuenta válido", vbInformation, "Aviso"
        txtCuenta.SetFocus
        Exit Sub
    End If

    If lnMonto = 0 Then
        MsgBox "Debe colocar un monto mayor a cero", vbInformation, "Aviso"
        txtMonto.SetFocus
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
    
'    If fsPINBlock = "" Then
'        MsgBox "Debe digitar una clave válida", vbInformation, "Aviso"
'        cmdLector.SetFocus
'        Exit Sub
'    End If
'
'    If fsTrack2 = "" Then
'        MsgBox "Debe digitar un Track2", vbInformation, "Aviso"
'        cmdLector.SetFocus
'        Exit Sub
'    End If


    lsIFTipo = Format$(gTpoIFCmac, "00")
    lsGlosa = Trim(txtGlosa.Text)
    lnMoneda = Right(cboMoneda.Text, 2)
    
    If MsgBox(" Desea Grabar la Operación - " & sOpeDesc & " ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then

        cmdGrabar.Enabled = False
            
        Call RegistrarOperacionInterCMAC(fsPAN, fsPINBlock, lsCuenta, nOperacion, fsTrack2, lnMoneda, "", sPersCodCMAC, sLpt, sOpeDesc, sNombreCMAC, gdFecSis, gsCodAge, gsCodUser, lnMonto, lsGlosa, lsIFTipo, False, 0, lnComision)
                
        Unload Me
    End If
 
End Sub

Private Sub cmdLector_Click()
    Call LectorTarjeta
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF10 And txtCuenta.Enabled Then
        Call LectorTarjeta
    End If
End Sub




Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGlosa.SetFocus
        Exit Sub
    End If
    NumerosEnteros (KeyAscii)
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        Me.txtMonto.SetFocus
    End If
End Sub

Private Function GetLimiteMonto(ByVal nmoneda As Moneda) As Double
Dim nValor As Double

    nValor = 0


    GetLimiteMonto = nValor

End Function

Private Function GetComisionMayor() As Double
Dim rsPar As New ADODB.Recordset
Dim clsFun As COMOpeInterCMAC.dFuncionesNeg

    Set clsFun = New COMOpeInterCMAC.dFuncionesNeg

    Set rsPar = clsFun.GetTarifaParametro(nOperacion, gMonedaNacional, 2077)
    Set clsFun = Nothing
    If rsPar.EOF And rsPar.BOF Then
        GetComisionMayor = 0
    Else
        GetComisionMayor = rsPar("nParValor")
    End If
    rsPar.Close
    Set rsPar = Nothing

End Function

Private Function GetValorComision() As Double
Dim rsPar As New ADODB.Recordset
Dim clsFun As COMOpeInterCMAC.dFuncionesNeg

    Set clsFun = New COMOpeInterCMAC.dFuncionesNeg

    Set rsPar = clsFun.GetTarifaParametro(nOperacion, gMonedaNacional, gCostoOperacionCMACLlam)
    Set clsFun = Nothing
    If rsPar.EOF And rsPar.BOF Then
        GetValorComision = 0
    Else
        GetValorComision = rsPar("nParValor")
    End If
    rsPar.Close
    Set rsPar = Nothing
End Function

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
Dim vMonto As Double
Dim vNComision As Double

    If KeyAscii = 13 Then
        If Val(txtMonto.Text) = 0 Then
            cmdGrabar.Enabled = False
        Else
            vMonto = GetLimiteMonto(nmoneda)
            cmdGrabar.Enabled = True
        End If

        If cmdGrabar.Enabled Then cmdGrabar.SetFocus
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
