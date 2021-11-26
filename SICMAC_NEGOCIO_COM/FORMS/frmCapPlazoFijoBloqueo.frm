VERSION 5.00
Begin VB.Form frmCapPlazoFijoBloqueo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desbloqueo de Plazo Fijo"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   60
      TabIndex        =   19
      Top             =   1680
      Width           =   7815
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   280
         Width           =   570
      End
      Begin VB.Label LblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000006&
         Height          =   300
         Left            =   960
         TabIndex        =   22
         Top             =   240
         Width           =   6540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Relación :"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   645
         Width           =   720
      End
      Begin VB.Label LblRelacion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000006&
         Height          =   300
         Left            =   960
         TabIndex        =   20
         Top             =   600
         Width           =   1260
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6600
      TabIndex        =   17
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H00000080&
      Height          =   1680
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Frame fraDatos 
         Height          =   945
         Left            =   60
         TabIndex        =   1
         Top             =   620
         Width           =   7680
         Begin VB.Label lblEtqTasaCanc 
            AutoSize        =   -1  'True
            Caption         =   "TEA Canc (%) :"
            Height          =   195
            Left            =   6375
            TabIndex        =   15
            Top             =   255
            Width           =   1080
         End
         Begin VB.Label lblTasaCanc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   6600
            TabIndex        =   14
            Top             =   540
            Width           =   930
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   360
            TabIndex        =   13
            Top             =   255
            Width           =   690
         End
         Begin VB.Label lblApertura 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   1140
            TabIndex        =   12
            Top             =   202
            Width           =   1500
         End
         Begin VB.Label lblEtqUltCnt 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (días) :"
            Height          =   195
            Left            =   2730
            TabIndex        =   11
            Top             =   255
            Width           =   930
         End
         Begin VB.Label lblPlazo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   3765
            TabIndex        =   10
            Top             =   202
            Width           =   700
         End
         Begin VB.Label lblDuplicados 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   3765
            TabIndex        =   9
            Top             =   540
            Width           =   700
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "# Duplicados :"
            Height          =   195
            Left            =   2670
            TabIndex        =   8
            Top             =   593
            Width           =   1035
         End
         Begin VB.Label lblVencimiento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   1140
            TabIndex        =   7
            Top             =   540
            Width           =   1500
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento :"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   593
            Width           =   960
         End
         Begin VB.Label lblTasa 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   5385
            TabIndex        =   5
            Top             =   202
            Width           =   855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "TEA (%) :"
            Height          =   195
            Left            =   4650
            TabIndex        =   4
            Top             =   255
            Width           =   660
         End
         Begin VB.Label lblDias 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   5385
            TabIndex        =   3
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "# Días :"
            Height          =   195
            Left            =   4590
            TabIndex        =   2
            Top             =   593
            Width           =   585
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   60
         TabIndex        =   24
         Top             =   240
         Width           =   3630
         _extentx        =   6403
         _extenty        =   661
         texto           =   "Cuenta N°:"
         enabledcmac     =   -1  'True
         enabledcta      =   -1  'True
         enabledprod     =   -1  'True
         enabledage      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCapPlazoFijoBloqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by



Private Sub cmdCancelar_Click()
    LimpiaControles
End Sub

Private Sub CmdGuardar_Click()
    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
       If oCap.VerificaBloqueoPlazoFijo(Me.txtCuenta.NroCuenta) Then
            oCap.BloqueoDesPlazoFijo Me.txtCuenta.NroCuenta, 0
            'By Capi 21012009
            objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, , Me.txtCuenta.NroCuenta, gCodigoCuenta
            'End by

            MsgBox "Cuenta de Plazo Fijo Desblouqeada", vbInformation, "Aviso"
            CmdGuardar.Enabled = False
       Else
            MsgBox "Cuenta no esta Bloqueada", vbInformation, "Aviso"
            CmdGuardar.Enabled = False
       End If
    Set oCap = Nothing
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtCuenta.Prod = Trim(gCapPlazoFijo)
    txtCuenta.EnabledProd = False
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.EnabledCMAC = False
    Me.txtCuenta.Age = gsCodAge
    Me.CmdGuardar.Enabled = False
    'By Capi 20012009
    Set objPista = New COMManejador.Pista
    gsOpeCod = gPFDesbloqueo
    'End By


End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCta As String
        sCta = txtCuenta.NroCuenta
        'funcion qu verifique bloqueo
        ObtieneDatosCuenta sCta
        CmdGuardar.Enabled = True
    End If
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim nTasaNominal As Double
Dim nTasaCancelacion As Double, nMontoRetiro As Double
Dim rsCta As ADODB.Recordset, rsRel As New ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim nRow As Long
Dim sMsg As String, sMoneda As String, sPersona As String
Dim dUltRetInt As Date
Dim bGarantia As Boolean
Dim dRenovacion As Date, dApeReal As Date

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
Set clsMant = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm")
    lblPlazo = Format$(rsCta("nPlazo"), "#,##0")
    lblVencimiento = Format(DateAdd("d", rsCta("nPlazo"), rsCta("dRenovacion")), "dd mmm yyyy")
    
    nTasaNominal = rsCta("nTasaInteres")
    lbltasa = Format$(ConvierteTNAaTEA(nTasaNominal), "#0.00")
    lblDuplicados = rsCta("nDuplicado")
    
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
        dUltRetInt = clsCap.GetFechaUltimoRetiroIntPF(sCuenta)
        lblDias = Format$(DateDiff("d", dUltRetInt, gdFecSis), "#0")
        nMontoRetiro = clsCap.GetSaldoCancelacion(sCuenta, gdFecSis, gsCodAge, nTasaCancelacion)
        lblTasaCanc = Format$(ConvierteTNAaTEA(nTasaCancelacion), "#,##0.00")
    Set clsCap = Nothing
    
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
    Set clsMant = Nothing
    
    If Not (rsRel.EOF And rsRel.BOF) Then
        Me.LblCliente.Caption = UCase(PstaNombre(rsRel("Nombre")))
        Me.LblRelacion.Caption = UCase(rsRel("Relacion")) & Space(50) & Trim(rsRel("nPrdPersRelac"))
    End If
Else
    MsgBox "Nro de Cuenta Incorrecta", vbInformation, "Aviso"
End If
    
End Sub

Private Sub LimpiaControles()
    CmdGuardar.Enabled = False
    txtCuenta.Age = ""
    txtCuenta.cuenta = ""
    lblApertura = ""
    lblPlazo = ""
    lblDias = ""
    lblVencimiento = ""
    lblDuplicados = ""
    lbltasa = ""
    lblTasaCanc = ""
    LblCliente = ""
    LblRelacion = ""
End Sub

