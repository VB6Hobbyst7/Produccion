VERSION 5.00
Begin VB.Form frmPITCapOpeInterCajaEnvio 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "frmPITCapOpeInterCajaEnvio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLector 
      Caption         =   "&Lector Tarjeta"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   3960
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5685
      TabIndex        =   3
      Top             =   3960
      Width           =   1000
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
      Height          =   1885
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   6735
      Begin VB.Frame fraMonto 
         Height          =   1485
         Left            =   3240
         TabIndex        =   11
         Top             =   240
         Width           =   3330
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
            Left            =   1320
            MaxLength       =   14
            TabIndex        =   12
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label LblComision 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1320
            TabIndex        =   20
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Comisión :"
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
            TabIndex        =   19
            Top             =   900
            Width           =   885
         End
         Begin VB.Label Label4 
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
            Left            =   240
            TabIndex        =   13
            Top             =   420
            Width           =   660
         End
      End
      Begin VB.Frame fraGlosa 
         Caption         =   "Glosa"
         Height          =   1485
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2925
         Begin VB.TextBox txtGlosa 
            Height          =   1125
            Left            =   75
            TabIndex        =   10
            Top             =   240
            Width           =   2730
         End
      End
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
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   6735
      Begin VB.TextBox txtDNI 
         Height          =   315
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   17
         Top             =   1020
         Visible         =   0   'False
         Width           =   2175
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
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   1
         Top             =   660
         Width           =   2535
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmPITCapOpeInterCajaEnvio.frx":030A
         Left            =   1080
         List            =   "frmPITCapOpeInterCajaEnvio.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2535
      End
      Begin VB.Label lblDNI 
         Caption         =   "DNI:"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   1050
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Tarjeta: "
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1500
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
         Top             =   1440
         Width           =   3555
      End
      Begin VB.Label Label6 
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   330
         Width           =   675
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
         TabIndex        =   7
         Top             =   240
         Width           =   2835
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta :"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   720
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmPITCapOpeInterCajaEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nOperacion As CaptacOperacion
Dim sPersCodCMAC As String, sNombreCMAC As String, sDescOperacion As String
Dim nmoneda As Integer
Dim nMontoMinRet As Double, nMontoMaxRet As Double, nMontoMinReqDNI As Double
Dim nMontoMinDep As Double, nMontoMaxDep As Double
Dim nMontoMaxOpeXDia As Double, nMontoMaxOpeXMes As Double

Dim fsPAN As String, fsTrack2 As String, fsPINBlock As String

Public Sub inicia(ByVal nOpe As CaptacOperacion, ByVal sDescOpe As String, _
        ByVal sCodCmac As String, ByVal sNomCmac As String, Optional nComision As Double = 0)
    
    fsPAN = ""
    fsTrack2 = ""
    fsPINBlock = ""
    
    sDescOperacion = sDescOpe
    Me.Caption = "Captaciones - Operaciones InterCMACs - " & sDescOperacion
    nOperacion = nOpe
    sPersCodCMAC = sCodCmac
    sNombreCMAC = sNomCmac
    lblMensaje = sNombreCMAC & Chr$(13) & sDescOperacion
    
    If nOperacion = "261002" Then 'Depósito
        txtMonto.ForeColor = &HC00000
    ElseIf nOperacion = "261001" Then 'Retiro
        txtMonto.ForeColor = &HC0&
    End If
    
    txtMonto.Text = 0#
    txtMonto.Text = Format(txtMonto, "#,##0.00")
    gsOpeCod = CStr(nOperacion)
    
    LblComision = Format$(nComision, "#,##0.00")
    
    cboMoneda.AddItem "SOLES" & Space(100) & gMonedaNacional
    cboMoneda.AddItem "DOLARES" & Space(100) & gMonedaExtranjera

    Me.Show 1
End Sub

Private Sub cboMoneda_Change()
Dim vMonto As Double
Dim vNComision As Double
Dim vTipoCambio As COMOpeInterCMAC.dFuncionesNeg

    Set vTipoCambio = New COMOpeInterCMAC.dFuncionesNeg
    If Val(txtMonto.Text) = 0 Then
        cmdGrabar.Enabled = False
    Else
        vMonto = GetLimiteMonto(nmoneda)
        cmdGrabar.Enabled = True
    End If
    Set vTipoCambio = Nothing
    
End Sub

Private Sub cboMoneda_Click()
Dim vMonto As Double
Dim vNComision As Double

    nmoneda = Right(cboMoneda.Text, 2)
    txtMonto.BackColor = IIf(nmoneda = gMonedaNacional, &HC0FFFF, &HC0FFC0)
    txtMonto.Text = Format(txtMonto, "#,##0.00")
    
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

Private Sub cmdGrabar_Click()
Dim lsCuenta As String, lsDNI As String
Dim lnMonto As Double
Dim lsGlosa As String
Dim lnComision As Currency
Dim lnMonOpeXDia As Double
Dim lnMonOpeXMes As Double
Dim lnNumOpeXDia As Integer
Dim lnNumOpeXMes As Integer

    lsCuenta = Trim(txtCuenta.Text)
    lnMonto = CDbl(txtMonto.Text)
    lsGlosa = txtGlosa.Text
    lsDNI = txtDNI.Text
    lnComision = CDbl(LblComision.Caption)
    
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
    
    If nOperacion = 261001 Then 'Para retiros
        nMontoMinRet = IIf(nmoneda = 1, gnMontoMinRetMN, gnMontoMinRetME)
        nMontoMaxRet = IIf(nmoneda = 1, gnMontoMaxRetMN, gnMontoMaxRetME)
        nMontoMinReqDNI = IIf(nmoneda = 1, gnMontoMinRetMNReqDNI, gnMontoMinRetMEReqDNI)
    
        If lnMonto < nMontoMinRet Then
            MsgBox "El monto mínimo para esta operación es " & CStr(Format(nMontoMinRet, "#,##0.00")), vbInformation, "Aviso"
            txtMonto.SetFocus
            Exit Sub
        End If

        If lnMonto > nMontoMaxRet Then
            MsgBox "El monto máximo para esta operación es " & CStr(Format(nMontoMaxRet, "#,##0.00")), vbInformation, "Aviso"
            txtMonto.SetFocus
            Exit Sub
        End If
        
        If lnMonto <= nMontoMinReqDNI Then
            lblDNI.Visible = False
            txtDNI.Visible = False
            txtDNI.Text = ""
        End If
        
        If lnMonto > nMontoMinReqDNI And txtDNI.Text = "" Then
            lblDNI.Visible = True
            txtDNI.Visible = True
            MsgBox "Es necesario ingresar el Nº de DNI del cliente", vbInformation, "Aviso"
            txtDNI.SetFocus
            Exit Sub
        End If
    End If
    
    
    If nOperacion = 261002 Then 'Para depósitos
        nMontoMinDep = IIf(nmoneda = 1, gnMontoMinDepMN, gnMontoMinDepME)
        nMontoMaxDep = IIf(nmoneda = 1, gnMontoMaxDepMN, gnMontoMaxDepME)

        If lnMonto < nMontoMinDep Then
            MsgBox "El monto mínimo para esta operación es " & CStr(Format(nMontoMinDep, "#,##0.00")), vbInformation, "Aviso"
            txtMonto.SetFocus
            Exit Sub
        End If

        If lnMonto > nMontoMaxDep Then
            MsgBox "El monto máximo para esta operación es " & CStr(Format(nMontoMaxDep, "#,##0.00")), vbInformation, "Aviso"
            txtMonto.SetFocus
            Exit Sub
        End If

    End If
    
    If (nOperacion = 261001 Or nOperacion = 261002) Then
        nMontoMaxOpeXDia = IIf(nmoneda = 1, gnMontoMaxOpeMNxDia, gnMontoMaxOpeMExDia)
        nMontoMaxOpeXMes = IIf(nmoneda = 1, gnMontoMaxOpeMNxMes, gnMontoMaxOpeMExMes)
        lnMonOpeXDia = RecuperaMontoOpeXdia(lsCuenta, nmoneda, lnNumOpeXDia)
        'lnMonOpeXDia = lnMonOpeXDia + lnMonto
        lnMonOpeXMes = RecuperaMontoOpeXMes(lsCuenta, nmoneda, lnNumOpeXMes)
        'lnMonOpeXMes = lnMonOpeXMes + lnMonto
        
        If lnMonOpeXDia > nMontoMaxOpeXDia Then
            MsgBox "Supera el monto limite de operaciones por dia", vbInformation, "Aviso"
            txtMonto.SetFocus
            Exit Sub
        End If
        
        If lnNumOpeXDia > gnNumeroMaxOpeXDia Then
            MsgBox "Supera el mumero de operaciones por dia", vbInformation, "Aviso"
            txtMonto.SetFocus
            Exit Sub
        End If

        If lnMonOpeXMes > nMontoMaxOpeXMes Then
            MsgBox "Supera el monto limite de operaciones por Mes", vbInformation, "Aviso"
            txtMonto.SetFocus
            Exit Sub
        End If
        
        If lnNumOpeXMes > gnNumeroMaxOpeXMes Then
            MsgBox "Supera el mumero de operaciones por Mes", vbInformation, "Aviso"
            txtMonto.SetFocus
            Exit Sub
        End If
        
    End If
    
    If fsPAN = "" And nOperacion = 261001 Then
        MsgBox "Debe digitar un número de tarjeta válido", vbInformation, "Aviso"
        cmdLector.SetFocus
        Exit Sub
    End If
    
    If fsTrack2 = "" And nOperacion = 261001 Then
        MsgBox "Debe digitar el Tack2", vbInformation, "Aviso"
        cmdLector.SetFocus
        Exit Sub
    End If
    If fsPINBlock = "" And nOperacion = 261001 Then
        MsgBox "Debe ingresar su clave de Tarjeta", vbInformation, "Aviso"
        cmdLector.SetFocus
        Exit Sub
    End If
    
    If fsPAN = "" And nOperacion = 261002 Then
        fsPAN = "0000000000000000"
        fsPINBlock = "0000000000000000"
        fsTrack2 = "0000000000000000=00000000000000000000"
    End If

    If MsgBox("Desea Grabar la Operacion??", vbQuestion + vbYesNo, "Aviso") = vbYes Then

        Call RegistrarOperacionInterCMAC(fsPAN, fsPINBlock, lsCuenta, nOperacion, fsTrack2, nmoneda, lsDNI, sPersCodCMAC, sLpt, sDescOperacion, sNombreCMAC, gdFecSis, gsCodAge, gsCodUser, lnMonto, lsGlosa, "", gbImpTMU, , lnComision)

        Unload Me
    End If
End Sub


Private Sub cmdLector_Click()
    Call LectorTarjeta
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtMonto.Text = Format(txtMonto, "#,##0.00")
End Sub


Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGlosa.SetFocus
        Exit Sub
    End If
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub


Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGlosa.SetFocus
        Exit Sub
    End If
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        txtMonto.SetFocus
        If cmdGrabar.Enabled = True Then
            cmdGrabar.SetFocus
        End If
    End If
End Sub


Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 15, 2)
    If KeyAscii = 13 Then
        
        If nOperacion = 261001 Then 'Sólamente para retiros
        
            nMontoMinReqDNI = IIf(nmoneda = 1, gnMontoMinRetMNReqDNI, gnMontoMinRetMEReqDNI)
            
            If CDbl(txtMonto.Text) < nMontoMinReqDNI Then
                lblDNI.Visible = False
                txtDNI.Visible = False
                txtDNI.Text = ""
            End If
                    
            If CDbl(txtMonto.Text) >= nMontoMinReqDNI And txtDNI.Text = "" Then
                lblDNI.Visible = True
                txtDNI.Visible = True
                MsgBox "Es necesario ingresar el Nº de DNI del cliente", vbInformation, "Aviso"
                txtDNI.SetFocus
                Exit Sub
            End If
        End If
        
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
        
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF10 And txtCuenta.Enabled Then
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
Private Function GetLimiteMonto(ByVal nmoneda As Moneda) As Double
Dim nValor As Double

    nValor = 0


    GetLimiteMonto = nValor

End Function
