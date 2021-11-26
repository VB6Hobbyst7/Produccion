VERSION 5.00
Begin VB.Form FrmCapPlazoPanderito 
   Caption         =   "Renovacion de Plazo Ahorro Diario"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   Icon            =   "FrmCapPlazoPanderito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   5400
      TabIndex        =   21
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Renovación"
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
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   2760
      Width           =   7815
      Begin VB.ComboBox CboRenovacion 
         Height          =   315
         ItemData        =   "FrmCapPlazoPanderito.frx":030A
         Left            =   1200
         List            =   "FrmCapPlazoPanderito.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LblNewRenovacion 
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
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Width           =   3180
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Plazo :"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nueva Renovación :"
         Height          =   195
         Left            =   2280
         TabIndex        =   17
         Top             =   285
         Width           =   1485
      End
   End
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
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   7815
      Begin VB.Label LblRelacion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000006&
         Height          =   300
         Left            =   960
         TabIndex        =   13
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Relación :"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   645
         Width           =   720
      End
      Begin VB.Label LblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000006&
         Height          =   300
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   6540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   280
         Width           =   570
      End
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
      Height          =   1800
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Frame fraDatos 
         Height          =   1065
         Left            =   60
         TabIndex        =   1
         Top             =   600
         Width           =   7680
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Abono  :"
            Height          =   195
            Left            =   6160
            TabIndex        =   25
            Top             =   600
            Width           =   600
         End
         Begin VB.Label LblMontoAbo 
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
            Left            =   6840
            TabIndex        =   24
            Top             =   555
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Renovación :"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   960
         End
         Begin VB.Label LblRenovacion 
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
            TabIndex        =   14
            Top             =   600
            Width           =   3165
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "TEA (%) :"
            Height          =   195
            Left            =   4620
            TabIndex        =   7
            Top             =   240
            Width           =   660
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
            Left            =   5355
            TabIndex        =   6
            Top             =   195
            Width           =   705
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
            Left            =   5360
            TabIndex        =   5
            Top             =   555
            Width           =   705
         End
         Begin VB.Label lblEtqUltCnt 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (días) :"
            Height          =   195
            Left            =   4360
            TabIndex        =   4
            Top             =   615
            Width           =   930
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
            TabIndex        =   3
            Top             =   195
            Width           =   3180
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   360
            TabIndex        =   2
            Top             =   255
            Width           =   690
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   3630
         _extentx        =   6403
         _extenty        =   661
         texto           =   "Cuenta N°:"
         enabledcmac     =   -1
         enabledcta      =   -1
         enabledprod     =   -1
         enabledage      =   -1
      End
   End
End
Attribute VB_Name = "FrmCapPlazoPanderito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnPlazo As Integer
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by


Private Sub CboRenovacion_Click()
    lnPlazo = CInt(CboRenovacion.Text)
End Sub

Private Sub CmdGuardar_Click()
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim lsMensaje As String
    
    Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
    lsMensaje = clsCap.GetEstadoAhoPanderito(TxtCuenta.NroCuenta, gdFecSis)
    If lsMensaje <> "" Then
        MsgBox lsMensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
       'Grabar operación
        clsCap.ActualizaCaptaAhoPanderito TxtCuenta.NroCuenta, lnPlazo, gdFecSis
        'By Capi 21012009
        objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, , Trim(TxtCuenta.NroCuenta), gCodigoCuenta
        'End by

        LimpiaControles
    End If
    
    Set clsCap = Nothing
End Sub

Private Sub cmdsalir_Click()
     Unload Me
End Sub

Private Sub Form_Load()
    TxtCuenta.Prod = Trim(gCapAhorros)
    TxtCuenta.EnabledProd = False
    TxtCuenta.CMAC = gsCodCMAC
    TxtCuenta.EnabledCMAC = False
    Me.TxtCuenta.Age = gsCodAge
    Me.CmdGuardar.Enabled = False
    CboRenovacion.ListIndex = 0
    'By Capi 20012009
    Set objPista = New COMManejador.Pista
    gsOpeCod = gAhoRenovPlazoPanderito
    'End By


End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCta As String
        sCta = TxtCuenta.NroCuenta
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
    If rsCta("nTpoPrograma") = 2 Then
        lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm")
        lblPlazo = Format$(rsCta("nPlazo"), "#,##0")
        nTasaNominal = rsCta("nTasaInteres")
        lblTasa = Format$(ConvierteTNAaTEA(nTasaNominal), "#0.00")
        LblMontoAbo = IIf(IsNull(rsCta("nMontoAbono")), 0, rsCta("nMontoAbono"))
        LblRenovacion = IIf(IsNull(rsCta("dRenoPanderito")), "", Format$(rsCta("dRenoPanderito"), "dd mmm yyyy hh:mm"))
        
        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
            Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        Set clsMant = Nothing
        
        If Not (rsRel.EOF And rsRel.BOF) Then
            Me.LblCliente.Caption = UCase(PstaNombre(rsRel("Nombre")))
            Me.LblRelacion.Caption = UCase(rsRel("Relacion")) & Space(50) & Trim(rsRel("nPrdPersRelac"))
        End If
        
        LblNewRenovacion = Format$(gdFecSis, "dd mmm yyyy hh:mm")
    Else
        MsgBox "No es un Ahorro Panderito", vbInformation, "Aviso"
        Exit Sub
    End If
Else
    MsgBox "Nro de Cuenta Incorrecta", vbInformation, "Aviso"
End If
    
End Sub

Private Sub cmdCancelar_Click()
    LimpiaControles
End Sub

Private Sub LimpiaControles()
    CmdGuardar.Enabled = False
    TxtCuenta.Age = ""
    TxtCuenta.Cuenta = ""
    lblApertura = ""
    lblPlazo = ""
    lblTasa = ""
    LblCliente = ""
    LblRelacion = ""
    LblNewRenovacion = ""
    LblRenovacion = ""
    LblMontoAbo = ""
End Sub


