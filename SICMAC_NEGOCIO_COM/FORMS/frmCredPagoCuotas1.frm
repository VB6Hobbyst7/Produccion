VERSION 5.00
Begin VB.Form frmCredPagoCuotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pago de Cuotas"
   ClientHeight    =   7110
   ClientLeft      =   3330
   ClientTop       =   2220
   ClientWidth     =   7065
   ForeColor       =   &H8000000F&
   Icon            =   "frmCredPagoCuotas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Credito"
      Height          =   1200
      Left            =   30
      TabIndex        =   58
      Top             =   -15
      Width           =   6990
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Creditos"
         Height          =   960
         Left            =   4800
         TabIndex        =   59
         Top             =   150
         Width           =   2115
         Begin VB.ListBox LstCred 
            Height          =   450
            Left            =   75
            TabIndex        =   3
            Top             =   225
            Width           =   1980
         End
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3840
         TabIndex        =   1
         Top             =   315
         Width           =   900
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   435
         Left            =   180
         TabIndex        =   0
         Top             =   285
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   767
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label LblAgencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   60
         Top             =   780
         Width           =   3465
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Pago"
      Height          =   5865
      Left            =   30
      TabIndex        =   2
      Top             =   1200
      Width           =   7005
      Begin VB.Frame Frame3 
         Height          =   1560
         Left            =   270
         TabIndex        =   10
         Top             =   3810
         Width           =   6570
         Begin VB.TextBox TxtMonPag 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1350
            MaxLength       =   15
            TabIndex        =   5
            Top             =   600
            Width           =   1380
         End
         Begin VB.ComboBox CmbForPag 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1335
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   210
            Width           =   1785
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Pag.Total. :"
            Height          =   195
            Left            =   4620
            TabIndex        =   70
            Top             =   630
            Width           =   825
         End
         Begin VB.Label lblPagoTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5475
            TabIndex        =   69
            Top             =   600
            Width           =   1020
         End
         Begin VB.Label LblItf 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3405
            TabIndex        =   68
            Top             =   600
            Width           =   1020
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "I.T.F. :"
            Height          =   195
            Left            =   2820
            TabIndex        =   67
            Top             =   630
            Width           =   465
         End
         Begin VB.Label LblNumDoc 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4425
            TabIndex        =   63
            Top             =   225
            Width           =   1665
         End
         Begin VB.Label LblEstado 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   4560
            TabIndex        =   21
            Top             =   1215
            Width           =   75
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Estado Credito"
            Height          =   195
            Left            =   3180
            TabIndex        =   20
            Top             =   1200
            Width           =   1035
         End
         Begin VB.Label LblNewCPend 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1905
            TabIndex        =   19
            Top             =   1230
            Width           =   45
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Nueva Cuota Pendiente"
            Height          =   195
            Left            =   90
            TabIndex        =   18
            Top             =   1230
            Width           =   1710
         End
         Begin VB.Label LblNewSalCap 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1890
            TabIndex        =   17
            Top             =   900
            Width           =   45
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nuevo Saldo de Capital"
            Height          =   195
            Left            =   90
            TabIndex        =   16
            Top             =   915
            Width           =   1680
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Monto a Pagar"
            Height          =   195
            Left            =   135
            TabIndex        =   15
            Top             =   615
            Width           =   1050
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Prox. fecha Pag :"
            Height          =   195
            Left            =   3180
            TabIndex        =   14
            Top             =   945
            Width           =   1230
         End
         Begin VB.Label LblProxfec 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4575
            TabIndex        =   13
            Top             =   975
            Width           =   45
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
            Height          =   195
            Left            =   3210
            TabIndex        =   12
            Top             =   255
            Width           =   1050
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Forma Pago"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton CmdPlanPagos 
         Caption         =   "&Plan Pagos"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1560
         TabIndex        =   66
         Top             =   5400
         Width           =   1275
      End
      Begin VB.Frame Frame2 
         Height          =   2055
         Left            =   240
         TabIndex        =   22
         Top             =   195
         Width           =   6570
         Begin VB.Label LblCalMiViv 
            Appearance      =   0  'Flat
            Caption         =   "Mal Pagador"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1110
            TabIndex        =   65
            Top             =   1710
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Calificacion :"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   1695
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label LblNomCli 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1110
            TabIndex        =   39
            Top             =   195
            Width           =   4950
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   210
            Width           =   480
         End
         Begin VB.Label LblMonCred 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   4290
            TabIndex        =   37
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Monto del Credito"
            Height          =   195
            Left            =   2835
            TabIndex        =   36
            Top             =   810
            Width           =   1245
         End
         Begin VB.Label LblSalCap 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1110
            TabIndex        =   35
            Top             =   1065
            Width           =   1155
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Capital"
            Height          =   195
            Left            =   105
            TabIndex        =   34
            Top             =   1095
            Width           =   930
         End
         Begin VB.Label LblLinCred 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1110
            TabIndex        =   33
            Top             =   495
            Width           =   4950
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Linea Credito"
            Height          =   195
            Left            =   105
            TabIndex        =   32
            Top             =   510
            Width           =   930
         End
         Begin VB.Label LblMoneda 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1110
            TabIndex        =   31
            Top             =   750
            Width           =   1155
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   105
            TabIndex        =   30
            Top             =   780
            Width           =   585
         End
         Begin VB.Label LblTotDeuda 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Height          =   270
            Left            =   4290
            TabIndex        =   29
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Deuda a la Fecha : "
            Height          =   195
            Left            =   2835
            TabIndex        =   28
            Top             =   1095
            Width           =   1410
         End
         Begin VB.Label Lbl2 
            AutoSize        =   -1  'True
            Caption         =   "Forma Pago"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1395
            Width           =   870
         End
         Begin VB.Label LblForma 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1110
            TabIndex        =   26
            Top             =   1365
            Width           =   480
         End
         Begin VB.Label LblTotCuo 
            AutoSize        =   -1  'True
            Caption         =   "Cuotas"
            Height          =   195
            Left            =   1650
            TabIndex        =   25
            Top             =   1410
            Width           =   495
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Monto Cuota :"
            Height          =   195
            Left            =   2835
            TabIndex        =   24
            Top             =   1410
            Width           =   1005
         End
         Begin VB.Label LblMontoCuota 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Height          =   270
            Left            =   4290
            TabIndex        =   23
            Top             =   1410
            Width           =   840
         End
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   270
         TabIndex        =   6
         Top             =   5400
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   2880
         TabIndex        =   7
         Top             =   5400
         Width           =   1275
      End
      Begin VB.CommandButton cmdmora 
         Caption         =   "&Mora"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4200
         TabIndex        =   8
         Top             =   5400
         Width           =   1275
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   345
         Left            =   5520
         TabIndex        =   9
         Top             =   5400
         Width           =   1275
      End
      Begin VB.Label LblMonCalDin 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1890
         TabIndex        =   62
         Top             =   3585
         Width           =   1155
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Monto Calen. Din:"
         Height          =   195
         Left            =   540
         TabIndex        =   61
         Top             =   3585
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota Pendiente"
         Height          =   210
         Left            =   555
         TabIndex        =   57
         Top             =   2340
         Width           =   1185
      End
      Begin VB.Label LblCPend 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1905
         TabIndex        =   56
         Top             =   2310
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Monto de Pago"
         Height          =   195
         Left            =   555
         TabIndex        =   55
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label LblMonPago 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1905
         TabIndex        =   54
         Top             =   2625
         Width           =   1155
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Mora Total"
         Height          =   195
         Left            =   3375
         TabIndex        =   53
         Top             =   2625
         Width           =   765
      End
      Begin VB.Label LblMora 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   4695
         TabIndex        =   52
         Top             =   2610
         Width           =   1275
      End
      Begin VB.Label LblGastos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4695
         TabIndex        =   51
         Top             =   2310
         Width           =   1275
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Gastos"
         Height          =   195
         Left            =   3375
         TabIndex        =   50
         Top             =   2355
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Venc."
         Height          =   195
         Left            =   3360
         TabIndex        =   49
         Top             =   2925
         Width           =   915
      End
      Begin VB.Label LblFecVec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   270
         Left            =   4695
         TabIndex        =   48
         Top             =   2910
         Width           =   1275
      End
      Begin VB.Label LblDiasAtraso 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   47
         Top             =   3240
         Width           =   45
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Dias Atrasados"
         Height          =   195
         Left            =   3360
         TabIndex        =   46
         Top             =   3225
         Width           =   1065
      End
      Begin VB.Label lblMoraC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1905
         TabIndex        =   45
         Top             =   2925
         Width           =   1155
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Mora Cuota"
         Height          =   210
         Left            =   555
         TabIndex        =   44
         Top             =   2940
         Width           =   825
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Cuotas en Mora"
         Height          =   195
         Left            =   555
         TabIndex        =   43
         Top             =   3255
         Width           =   1125
      End
      Begin VB.Label lblCuotasMora 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1905
         TabIndex        =   42
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label25 
         Caption         =   "Met Liq."
         Height          =   195
         Left            =   3360
         TabIndex        =   41
         Top             =   3495
         Width           =   735
      End
      Begin VB.Label lblMetLiq 
         Height          =   195
         Left            =   4755
         TabIndex        =   40
         Top             =   3480
         Width           =   645
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "Opciones"
      Visible         =   0   'False
      Begin VB.Menu mnuVerCredito 
         Caption         =   "Ver Credito"
      End
   End
End
Attribute VB_Name = "frmCredPagoCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nProducto As Producto
Private MatGastosCancelacion As Variant
Private nNumGastosCancel As Integer
Private MatGastosFinal As Variant
Private nNumGastosFinal As Integer

Private MatCalendTmp As Variant
Private MatCalend As Variant
Private MatCalend_2 As Variant
Private MatCalendNormalT1 As Variant
Private MatCalendParalelo As Variant
Private MatCalendMiVivResult As Variant
Private MatCalendDistribuido As Variant
Private MatCalendDistribuido_2 As Variant
Private MatCalendDistribuidoParalelo As Variant
Private MatCalendDistribuidoTempo As Variant

Private nNroTransac As Long
Private bCalenDinamic As Boolean
Private bCalenCuotaLibre As Boolean
Private bRecepcionCmact As Boolean
Private sPersCmac As String
Private vnIntPendiente As Double
Private vnIntPendientePagado As Double
Dim nCalPago As Integer
Dim bDistrib As Boolean
Dim bPrepago As Integer
Dim nCalendDinamTipo As Integer
Dim nMiVivienda As Integer
Dim MatDatos As Variant
Dim sOperacion As String
Dim sPerscod As String
Dim nInteresDesagio As Double

Dim nMontoPago As Double
Dim nITF As Double


' CMACICA_CSTS - 08/11/2003 -------------------------------------------------------------------
Dim nPrestamo As Double
Dim bCuotaCom As Integer
Dim nCalendDinamico As Integer

' RFA
Dim bRFA As Boolean
'----------------------------------------------------------------------------------------------

'Datos de Pantalla
Dim oCred As NCredDoc

Dim pLblAgencia As String
Dim plblMetLiq As String
Dim pLblMontoCuota As Double
Dim pbPrepago As Integer
Dim pnCalendDinamTipo As Integer
Dim pnCalendDinamico As Integer
Dim pbCalenDinamic As Boolean
Dim pbCalenCuotaLibre As Boolean
Dim pnMivivienda As Integer
Dim pnCalPago As Integer
Dim pLblCalMiViv As String
Dim pLblNomCli As String
Dim pLblLinCred As String
Dim pLblMoneda As String
Dim pLblMonCred As Double
Dim pLblSalCap As Double
Dim pLblForma As String
Dim pLblCPend As String
Dim pLblGastos As Double
Dim pLblMonPago As Double
Dim pLblMora As Double
Dim plblMoraC As Double
Dim pLblFecVec As String
Dim plblCuotasMora As String
Dim pLblDiasAtraso As Integer
Dim pTxtMonPag As Double
Dim pLblTotDeuda As Double
Dim pLblMonCalDin As Double
Dim pMSGJud As String

Public Sub RecepcionCmac(ByVal psPersCodCMAC As String)
    bRecepcionCmact = True
    sPersCmac = psPersCodCMAC
    Me.Show 1
End Sub

Private Function HabilitaActualizacion(ByVal pbHabilita As Boolean) As Boolean
    cmdmora.Enabled = pbHabilita
    Frame4.Enabled = Not pbHabilita
    CmbForPag.Enabled = pbHabilita
    LblNumDoc.Enabled = pbHabilita
    TxtMonPag.Enabled = pbHabilita
    If Mid(ActxCta.NroCuenta, 9, 1) = "1" Or Trim(Mid(ActxCta.NroCuenta, 9, 1)) = "" Then
        TxtMonPag.BackColor = vbWhite
        LblItf.BackColor = vbWhite
    Else
        TxtMonPag.BackColor = vbGreen
        LblItf.BackColor = vbGreen
    End If
    Frame3.Enabled = pbHabilita
    If CmbForPag.ListCount > 0 Then
        CmbForPag.ListIndex = 0
    End If
    If pbHabilita Then
        If TxtMonPag.Enabled And TxtMonPag.Visible Then
            TxtMonPag.SetFocus
        End If
    End If
End Function


Private Function ValidaDatos() As Boolean
    ValidaDatos = True
    If CInt(Trim(Right(CmbForPag, 10))) = gColocTipoPagoCheque Then
        If Trim(LblNumDoc.Caption) = "" Then
            ValidaDatos = False
            MsgBox "Ingrese Numero de Documento", vbInformation, "Aviso"
        End If
    End If
    
End Function

Private Sub LimpiaPantalla()
    Set oCred = Nothing
    LimpiaControles Me, True
    InicializaCombos Me
    Frame3.Enabled = False
    LblEstado.Caption = ""
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    LblNewSalCap.Caption = ""
    LblProxfec.Caption = ""
    LblNewCPend.Caption = ""
    LblEstado.Caption = ""
    bCalenDinamic = False
    LblAgencia.Caption = ""
    Label23.Visible = False
    LblCalMiViv.Visible = False
    
End Sub
Private Sub CargaControles()
Dim oCredGeneral As DCredGeneral
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Call CargaComboConstante(gColocTipoPago, CmbForPag)
    Exit Sub

ERRORCargaControles:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
Dim oCredito As DCredito
Dim R As ADODB.Recordset
Dim oNegCredito As NCredito
Dim oGastos As nGasto
Dim dParam As DParametro
Dim nAnios As Integer
Dim oAge As DAgencias

    On Error GoTo ErrorCargaDatos
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaDatosCreditoVigente(psCtaCod, gdFecSis)
    Set oCredito = Nothing
    Set dParam = New DParametro
    nAnios = dParam.RecuperaValorParametro(3053)
    Set dParam = Nothing
    If Not R.BOF And Not R.EOF Then
    
        If Mid(psCtaCod, 4, 2) <> gsCodAge Then
            Set oAge = New DAgencias
            LblAgencia.Caption = oAge.NombreAgencia(Mid(psCtaCod, 4, 2))
            Set oAge = Nothing
        Else
            LblAgencia.Caption = ""
        End If
                
        lblMetLiq.Caption = Trim(R!cMetLiquidacion)
        LblMontoCuota.Caption = Format(IIf(IsNull(R!CuotaAprobada), 0, R!CuotaAprobada), "#0.00")
        bPrepago = IIf(R!bPrepago = True, 1, 0)
        nCalendDinamTipo = R!nCalendDinamTipo
        Set oNegCredito = New NCredito
        
        ' CMACICA_CSTS - 08/11/2003 -------------------------------------------------------------------
        nCalendDinamico = R!nCalendDinamico
        ' ---------------------------------------------------------------------------------------------
        
        If IsNull(R!nCalendDinamico) Then
            bCalenDinamic = False
        Else
            If R!nCalendDinamico = 1 Then
                bCalenDinamic = True
            Else
                bCalenDinamic = False
            End If
        End If
        If R!nColocCalendCod = gColocCalendCodCL Then
            bCalenCuotaLibre = True
        Else
            bCalenCuotaLibre = False
        End If
        nMiVivienda = R!bMiVivienda
        
        nCalPago = IIf(IsNull(R!nCalPago), 0, R!nCalPago)
        If bPrepago = 1 Then
            If R!nPlazoTranscurrido > nAnios Then 'si es mayor a 10 años
                nCalPago = 1 'buen pagador
            Else
                nCalPago = 0 'Mal pagador
            End If
        End If
        
        If nMiVivienda Then
            Label23.Visible = True
            LblCalMiViv.Visible = True
            If nCalPago = 1 Then
                LblCalMiViv.Caption = "Buen Pagador"
            Else
                LblCalMiViv.Caption = "Mal Pagador"
            End If
            MatCalend = oNegCredito.RecuperaMatrizCalendarioPendiente(psCtaCod)
            MatCalend_2 = MatCalend
            MatCalendNormalT1 = MatCalend
            MatCalendParalelo = oNegCredito.RecuperaMatrizCalendarioPendiente(psCtaCod, True)
            '*******************************************************************************
            'Verificar si es MalPagador o Buen Pagador y Unir o mantener un solo Calendario
            '*******************************************************************************
            MatCalendMiVivResult = UnirMatricesMiViviendaAmortizacion(MatCalend, MatCalendParalelo)
            
            MatCalend = MatCalendMiVivResult
            MatCalendTmp = MatCalend
            MatCalendDistribuido = oNegCredito.CrearMatrizparaAmortizacion(MatCalend)
        Else
            Label23.Visible = False
            LblCalMiViv.Visible = False
            MatCalend = oNegCredito.RecuperaMatrizCalendarioPendiente(psCtaCod)
            MatCalendTmp = MatCalend
            MatCalendDistribuido = oNegCredito.CrearMatrizparaAmortizacion(MatCalend)
        End If
                
        CargaDatos = True
        vnIntPendiente = IIf(IsNull(R!nintPend), 0, R!nintPend)
        vnIntPendientePagado = 0
        nNroTransac = IIf(IsNull(R!nTransacc), 0, R!nTransacc)
        sPerscod = R!cPersCod
        LblNomCli.Caption = PstaNombre(R!cPersNombre)
        LblLinCred.Caption = Trim(R!cLineaCred)
        LblMoneda.Caption = Trim(R!cmoneda)
        LblMonCred.Caption = Format(R!nMontoCol, "#0.00")
        
        ' CMACICA_CSTS - 08/11/2003 -------------------------------------------------------------------
        nPrestamo = Format(R!nMontoCol, "#0.00")
        bCuotaCom = IIf(IsNull(R!bCuotaCom), 0, R!bCuotaCom)
        '----------------------------------------------------------------------------------------------
        
        LblSalCap.Caption = Format(R!nSaldo, "#0.00")
        LblForma.Caption = Trim(Str(R!nCuotasApr))
        LblCPend.Caption = MatCalend(0, 1)
                
        LblGastos.Caption = Format(oNegCredito.MatrizGastosVencidos(MatCalend, gdFecSis), "#0.00")
        If nMiVivienda Then
            LblMonPago.Caption = oNegCredito.MatrizMontoAPagarCuotaPendiente(MatCalend, gdFecSis)
        Else
            LblMonPago.Caption = oNegCredito.MatrizMontoAPagar(MatCalend, gdFecSis)
        End If
        
        LblMora.Caption = Format(oNegCredito.MatrizMoraTotal(MatCalend, gdFecSis), "#0.00")
        
        lblMoraC.Caption = Format(CDbl(MatCalend(0, 6)), "#0.00")
        LblFecVec.Caption = MatCalend(0, 0)
        lblCuotasMora.Caption = oNegCredito.MatrizCuotasEnMora(MatCalend, gdFecSis)
        LblDiasAtraso.Caption = Trim(Str(R!nDiasAtraso))
        'LblTotalAPagar.Caption = Format(oNegCredito.MatrizDeudaAlaFecha(psCtaCod, MatCalend, gdFecSis), "#0.00")
        lblMetLiq.Caption = Trim(R!cMetLiquidacion)
        TxtMonPag.Text = LblMonPago.Caption
        
        'Deuda a la Fecha
        Dim nInteresFecha As Currency
        Dim nMontoFecha As Currency
        nInteresFecha = oNegCredito.MatrizInteresGastosAFecha(psCtaCod, MatCalend, gdFecSis, True, bCalenDinamic)
        nMontoFecha = oNegCredito.MatrizCapitalAFecha(psCtaCod, MatCalend)
        
        LblTotDeuda.Caption = Format(nInteresFecha + nMontoFecha, "#0.00")
        
        nInteresDesagio = 0
        If nInteresFecha < 0 Then
            nInteresDesagio = Abs(nInteresFecha)
        End If
        '**** AGREGADO EJRS 30/09/2004 ******************
        'If CCur(TxtMonPag.Text) < CCur(LblTotDeuda.Caption) Then
        '    nInteresDesagio = 0
        'End If
        '*************************************************
        
        Set oGastos = New nGasto
        MatGastosCancelacion = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosCancel, gdFecSis, psCtaCod, 1, "CA", , , CDbl(LblTotDeuda.Caption), oNegCredito.MatrizMontoCapitalAPagar(MatCalend, gdFecSis), oNegCredito.MatrizCuotaPendiente(MatCalend, MatCalendDistribuido), , , , , R!nDiasAtraso)
        LblTotDeuda.Caption = Format(CDbl(LblTotDeuda.Caption) + MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("CA", "PA", "")), "#0.00")
        'LblTotDeuda.Caption = Format(CDbl(LblTotDeuda.Caption) + MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("PA", IIf(bPrepago = 1, "PP", ""), "")), "#0.00")
        LblGastos.Caption = Format(oNegCredito.MatrizGastosVencidos(MatCalend, gdFecSis) + MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("PA", IIf(bPrepago = 1, "PP", ""), "")), "#0.00")
        LblMonPago.Caption = Format(CDbl(LblMonPago.Caption) + MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("PA", IIf(bPrepago = 1, "PP", ""), "")), "#0.00")
        TxtMonPag.Text = LblMonPago.Caption
        Set oGastos = Nothing
        
        'Para Generar el calendario Dinamico
        'Si es mivivienda
        LblMonCalDin.Caption = Format(oNegCredito.MatrizMontoCalendDinamico(psCtaCod, MatCalend, gdFecSis, nMiVivienda) + MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("PA", "PP", "")), "#0.00")
        
        
        Set oNegCredito = Nothing
        If nMiVivienda = 1 And bPrepago = 0 Then
            TxtMonPag.Locked = True
        Else
            TxtMonPag.Locked = False
        End If
        
        Set oCredito = New DCredito
        If oCredito.NumerosCredEnJudicial(sPerscod) > 0 Then
            MsgBox "Cliente tiene Creditos en Judicial", vbInformation, "Aviso"
        End If
        Set oCredito = Nothing
        
        Call TxtMonPag_KeyPress(13)
    Else
        CargaDatos = False
    End If
    R.Close
    Set R = Nothing
    Exit Function

ErrorCargaDatos:
    MsgBox Err.Description, vbCritical, "Aviso"

End Function

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        VerificaRFA (ActxCta.NroCuenta)
'        If bRFA Then
'            MsgBox "Este creditos es RFA " & vbCrLf & _
'                   "Por favor ingrese a la opción de Pagos en RFA", vbInformation, ""
'        Else
'            Set oCred = New NCredDoc
'            If Not oCred.CargaDatosPago(ActxCta.NroCuenta, pLblAgencia, plblMetLiq, pLblMontoCuota, pbPrepago, _
'                        pnCalendDinamTipo, pnCalendDinamico, pbCalenDinamic, pbCalenCuotaLibre, pnMivivienda, pnCalPago, _
'                        pLblCalMiViv, pLblNomCli, pLblLinCred, pLblMoneda, pLblMonCred, pLblSalCap, pLblForma, pLblCPend, _
'                        pLblGastos, pLblMonPago, pLblMora, plblMoraC, pLblFecVec, plblCuotasMora, pLblDiasAtraso, pTxtMonPag, _
'                        pLblTotDeuda, pLblMonCalDin, pMSGJud, gdFecSis) Then
'                        HabilitaActualizacion False
'                        MsgBox "No se pudo encontrar el Credito, o el Credito No esta Vigente", vbInformation, "Aviso"
'                        CmdPlanPagos.Enabled = False
'
'                        Set oCred = Nothing
'            Else
'                        LblAgencia.Caption = pLblAgencia
'                        lblMetLiq.Caption = plblMetLiq
'                        LblMontoCuota.Caption = pLblMontoCuota
'                        LblCalMiViv.Caption = pLblCalMiViv
'                        LblNomCli.Caption = pLblNomCli
'                        LblLinCred.Caption = pLblLinCred
'                        LblMoneda.Caption = pLblMoneda
'                        LblMonCred.Caption = pLblMonCred
'                        LblSalCap.Caption = pLblSalCap
'                        LblForma.Caption = pLblForma
'                        LblCPend.Caption = pLblCPend
'                        LblGastos.Caption = pLblGastos
'                        LblMonPago.Caption = pLblMonPago
'                        LblMora.Caption = pLblMora
'                        lblMoraC.Caption = plblMoraC
'                        LblFecVec.Caption = pLblFecVec
'                        lblCuotasMora.Caption = plblCuotasMora
'                        LblDiasAtraso.Caption = pLblDiasAtraso
'                        TxtMonPag.Text = Format(pTxtMonPag, "#0.00")
'                        LblTotDeuda.Caption = pLblTotDeuda
'                        LblMonCalDin.Caption = pLblMonCalDin
'                        If Len(Trim(pMSGJud)) > 0 Then
'                            MsgBox pMSGJud, vbInformation, "Aviso"
'                        End If
'
'                CmdPlanPagos.Enabled = True
'                HabilitaActualizacion True
'            End If
'        End If
'    End If
End Sub

Private Sub CmbForPag_Click()

    LblNumDoc.Caption = ""
    If CmbForPag.ListIndex <> -1 Then
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCheque Then
            MatDatos = frmBuscaCheque.BuscaCheque(gChqEstEnValorizacion, CInt(Mid(ActxCta.NroCuenta, 9, 1)))
            If MatDatos(0) <> "" Then
                LblNumDoc.Caption = MatDatos(4)
                TxtMonPag.Text = MatDatos(3)
            Else
                LblNumDoc.Caption = ""
            End If
            LblNumDoc.Visible = True
        Else
            LblNumDoc.Visible = False
        End If
    End If
End Sub

Private Sub CmbForPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CInt(Trim(Right(CmbForPag, 10))) = gColocTipoPagoCheque Then
            TxtMonPag.SetFocus
        Else
            TxtMonPag.SetFocus
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As DCredito
Dim R As ADODB.Recordset
Dim oPers As UPersona

    
    LstCred.Clear
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set oCredito = New DCredito
        Set R = oCredito.RecuperaCreditosVigentes(oPers.sPerscod, , Array(gColocEstVigMor, gColocEstVigVenc, gColocEstVigNorm, gColocEstRefMor, gColocEstRefVenc, gColocEstRefNorm))
        Do While Not R.EOF
            LstCred.AddItem R!cCtaCod
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Set oCredito = Nothing
        FrmVerCredito.Inicio (oPers.sPerscod)
        Me.ActxCta.SetFocusCuenta
    End If
    If LstCred.ListCount = 0 Then
        MsgBox "El Cliente No Tiene Creditos Vigentes", vbInformation, "Aviso"
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    Set oCred = Nothing
    Call LimpiaPantalla
    Call HabilitaActualizacion(False)
    cmdGrabar.Enabled = False
    CmdPlanPagos.Enabled = False
End Sub

Private Sub cmdGrabar_Click()
'Dim oNegCred As NCredito
'Dim oDoc As NCredDoc
'Dim oConstante As DConstante
'Dim sError As String
'Dim sTipoCred As String
'Dim MatCalDinam As Variant
'Dim MatCalDinam_2 As Variant
'Dim sCad As String
'Dim sCad2 As String
'Dim vPrevio As Previo.clsPrevio
'Dim oCal As Dcalendario
'    On Error GoTo ErrorCmdGrabar_Click
'
'    If CInt(Trim(Right(CmbForPag.Text, 2))) = gColocTipoPagoCheque Then
'        If Trim(Me.LblNumDoc.Caption) = "" Then
'            MsgBox "Cheque No es Valido", vbInformation, "Aviso"
'            Me.CmbForPag.SetFocus
'            Exit Sub
'        End If
'        If IsArray(MatDatos) Then
'            If Trim(MatDatos(3)) = "" Then
'                MatDatos(3) = "0.00"
'            End If
'            If Trim(TxtMonPag.Text) = "" Then
'                TxtMonPag.Text = "0.00"
'            End If
'            If CDbl(TxtMonPag.Text) > CDbl(MatDatos(3)) Then
'                MsgBox "Monto de Pago No Puede Ser Mayor que el Monto de Cheque", vbInformation, "Aviso"
'                TxtMonPag.SetFocus
'                Exit Sub
'            End If
'        End If
'    End If
'
'    If MsgBox("Se va a Efectuar el Pago del Credito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'
'    Dim nPorcRetCTS As Double, nMontoLavDinero As Double, nTC As Double
'    Dim clsLav As DCOMLavado
'    Dim sPersLavDinero As String
'    ''''''''''''''''''''''''''''''
'    'Lavado de Dinero
'    sPersLavDinero = ""
'    Set clsLav = New DCOMLavado
'    If clsLav.EsOperacionEfectivo(Trim(sOperacion)) Then
'        If Not EsExoneradaLavadoDinero() Then
'
'            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
'            Set clsLav = Nothing
'
'            If Mid(ActxCta.NroCuenta, 9, 1) = gMonedaNacional Then
'                Dim clsTC As NCOMTipoCambio
'                Set clsTC = New NCOMTipoCambio
'                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
'                Set clsTC = Nothing
'            Else
'                nTC = 1
'            End If
'            If CDbl(TxtMonPag.Text) >= Round(nMontoLavDinero * nTC, 2) Then
'                'Falta Pasarlo a COM+
'                sPersLavDinero = IniciaLavDinero()
'                If sPersLavDinero = "" Then Exit Sub
'            End If
'        End If
'    End If
'    Set clsLav = Nothing
'
'    ''''''''''''''''''''''''''''''
'
'        Call oCred.GrabarPagoCredito(ActxCta.NroCuenta, CDbl(Me.TxtMonPag.Text), gdFecSis, Me.lblMetLiq.Caption, CInt(Trim(Right(CmbForPag.Text, 10))), _
'                gsCodAge, gsCodUser, Trim(LblNumDoc.Caption), bRecepcionCmact, sPersCmac, _
'                sPersLavDinero, nITF)
'
''        If sError <> "" Then
''            MsgBox sError, vbInformation, "Aviso"
''        Else
''            'Verifica si fue un pago para Calendario Dinamico
''            'bPrepago nCalendDinamTipo
''            If (bCalenDinamic Or bPrepago = 1) And (nMontoPago < CDbl(LblTotDeuda.Caption)) Then
''                If nMontoPago > CDbl(LblMonCalDin.Caption) Then
''                    If nMiVivienda = 1 Then
''                        MatCalDinam = oNegCred.ReprogramarCreditoenMemoriaTotalMiVivienda(ActxCta.NroCuenta, gdFecSis, MatCalDinam_2, IIf(nCalendDinamTipo = 1, True, False))
''                        'Reporgramacion 2 de otorgar un nuevo calendario en basae al saldo de capital pendiente
''                        'Como si fueera un nuevo credito bajo las cuotas pendientes
''                        oNegCred.ReprogramarCredito ActxCta.NroCuenta, MatCalDinam, 2, True, MatCalDinam_2, gdFecSis, , gsCodUser, gsCodAge
''                        Call oNegCred.ActualizarCalificacionMIVivienda(ActxCta.NroCuenta)
''                        Set oDoc = New NCredDoc
''                        sCad = oDoc.ImprimePlandePagos(ActxCta.NroCuenta, gsNomAge, gdFecSis, gsCodUser, oNegCred.MatrizCapitalCalendario(MatCalDinam) + oNegCred.MatrizCapitalCalendario(MatCalDinam_2), True)
''                        sCad2 = oDoc.ImprimePlandePagos(ActxCta.NroCuenta, gsNomAge, gdFecSis, gsCodUser, oNegCred.MatrizCapitalCalendario(MatCalDinam) + oNegCred.MatrizCapitalCalendario(MatCalDinam_2), True, True)
''                        Set vPrevio = New clsPrevio
''                        vPrevio.PrintSpool sLpt, sCad & sCad2
''                        Set vPrevio = Nothing
''                        Set oDoc = Nothing
''                    Else
''                        MatCalDinam = oNegCred.ReprogramarCreditoenMemoriaTotal(ActxCta.NroCuenta, gdFecSis)
''                        'Reporgramacion 2 de otorgar un nuevo calendario en basae al saldo de capital pendiente
''                        'Como si fueera un nuevo credito bajo las cuotas pendientes
''                        oNegCred.ReprogramarCredito ActxCta.NroCuenta, MatCalDinam, 2, , , gdFecSis, , gsCodUser, gsCodAge
''                        Set oDoc = New NCredDoc
''                        sCad = oDoc.ImprimePlandePagos(ActxCta.NroCuenta, gsNomAge, gdFecSis, gsCodUser, oNegCred.MatrizCapitalCalendario(MatCalDinam), False)
''                        Set vPrevio = New clsPrevio
''                        vPrevio.ShowImpreSpool sCad
''                        Set vPrevio = Nothing
''                        Set oDoc = Nothing
''                    End If
''
''
''                End If
''            End If
''
''            Set oConstante = New DConstante
''            sTipoCred = oConstante.DameDescripcionConstante(gProducto, CInt(ActxCta.Prod))
''            Set oConstante = Nothing
''            Set oDoc = New NCredDoc
''
''            If nMiVivienda = 1 Then
''                'Recupero para imprimir las boletas
''                MatCalendDistribuido = MatCalendDistribuidoTempo
''            End If
''
''            Set oCal = New Dcalendario
''
''            Call oDoc.ImprimeBoleta(ActxCta.NroCuenta, lblNomCli.Caption, gsNomAge, LblMoneda, _
''                oNegCred.MatrizCuotasPagadas(MatCalendDistribuido), gdFecSis, Format(FechaHora(gdFecSis), "hh:mm:ss"), nNroTransac + 1, Mid(sTipoCred, 1, 18), _
''                oNegCred.MatrizCapitalPagado(MatCalendDistribuido), oNegCred.MatrizIntCompPagado(MatCalendDistribuido), _
''                oNegCred.MatrizIntCompVencPagado(MatCalendDistribuido), _
''                oNegCred.MatrizIntMorPagado(MatCalendDistribuido), oNegCred.MatrizGastoPag(MatCalendDistribuido), _
''                oNegCred.MatrizIntGraciaPagado(MatCalendDistribuido), _
''                oNegCred.MatrizIntSuspensoPag(MatCalendDistribuido) + oNegCred.MatrizIntReprogPag(MatCalendDistribuido), _
''                oNegCred.MatrizSaldoCapital(MatCalend, MatCalendDistribuido), LblProxfec.Caption, _
''                gsCodUser, sLpt, gsInstCmac, IIf(Trim(Right(Me.CmbForPag.Text, 2)) = "2", True, False), Me.LblNumDoc.Caption, gsCodCMAC, nITF, nInteresDesagio, bRecepcionCmact)
''
''            Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes
''                Call oDoc.ImprimeBoleta(ActxCta.NroCuenta, lblNomCli.Caption, gsNomAge, LblMoneda, _
''                oNegCred.MatrizCuotasPagadas(MatCalendDistribuido), gdFecSis, Format(FechaHora(gdFecSis), "hh:mm:ss"), nNroTransac + 1, Mid(sTipoCred, 1, 18), _
''                oNegCred.MatrizCapitalPagado(MatCalendDistribuido), oNegCred.MatrizIntCompPagado(MatCalendDistribuido), _
''                oNegCred.MatrizIntCompVencPagado(MatCalendDistribuido), _
''                oNegCred.MatrizIntMorPagado(MatCalendDistribuido), oNegCred.MatrizGastoPag(MatCalendDistribuido), _
''                oNegCred.MatrizIntGraciaPagado(MatCalendDistribuido), _
''                oNegCred.MatrizIntSuspensoPag(MatCalendDistribuido) + oNegCred.MatrizIntReprogPag(MatCalendDistribuido), _
''                oNegCred.MatrizSaldoCapital(MatCalend, MatCalendDistribuido), LblProxfec.Caption, _
''                gsCodUser, sLpt, gsInstCmac, IIf(Trim(Right(Me.CmbForPag.Text, 2)) = "2", True, False), Me.LblNumDoc.Caption, gsCodCMAC, nITF, nInteresDesagio, bRecepcionCmact)
''            Loop
''            Set oDoc = Nothing
''            Set oCal = Nothing
'            Set oCred = Nothing
'            Call cmdCancelar_Click
' '       End If
'
'    End If
'    Exit Sub
'
'ErrorCmdGrabar_Click:
'    MsgBox Err.Description, vbCritical, "Aviso"
'
End Sub

Private Sub cmdmora_Click()
    Call TxtMonPag_KeyPress(13)
    Call frmCredMoraCuotas.MostarMoraDetalle(MatCalend, gdFecSis)
End Sub

Private Sub CmdPlanPagos_Click()

'Dim odCred As DCredito
'Dim oCredDoc As NCredDoc
'Dim sCadImp As String
'Dim sCadImp_2 As String
'Dim Prev As Previo.clsPrevio
'
    On Error GoTo ErrorCmdPlanPagos_Click
'            Set oCredDoc = New NCredDoc
'            Set Prev = New clsPrevio
'            sCadImp = oCredDoc.ImprimePlandePagos(ActxCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser, nPrestamo, nMiVivienda, , gsNomCmac, bCuotaCom, nCalendDinamico)
'            sCadImp_2 = ""
'            If nMiVivienda Then
'                sCadImp_2 = oCredDoc.ImprimePlandePagos(ActxCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser, nPrestamo, nMiVivienda, True, gsNomCmac, bCuotaCom, nCalendDinamico)
'            End If
'            Prev.Show sCadImp & sCadImp_2, "", False
'            Set Prev = Nothing
'            Set oCredDoc = Nothing
    
    Call frmCredHistCalendario.PagoCuotas(ActxCta.NroCuenta)

    Exit Sub

ErrorCmdPlanPagos_Click:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub Inicia(sCodOpe As String)
    bRecepcionCmact = False
    sOperacion = sCodOpe
    Me.Show 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sNumTar As String
    Dim sClaveTar As String
    Dim nErr As Integer
    Dim sCaption As String
    Dim clsGen As DGeneral
    Dim nEstado  As CaptacTarjetaEstado
    Set clsGen = New DGeneral
    
    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(nProducto, bRetSinTarjeta)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
    
'    If KeyCode = vbKeyF11 And txtCuenta.Enabled = True Then 'F11
'        Dim nPuerto As TipoPuertoSerial
'        nPuerto = clsGen.GetPuertoPeriferico(gPerifPINPAD)
'        If nPuerto < 0 Then nPuerto = gPuertoSerialCOM1
'        IniciaPinPad nPuerto
'
'        WriteToLcd "Pase su Tarjeta por la Lectora."
'        sCaption = Me.Caption
'        Me.Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
'        sNumTar = GetNumTarjeta
'        sNumTar = Trim(Replace(sNumTar, "-", "", 1, , vbTextCompare))
'        If Len(sNumTar) <> 16 Then
'            MsgBox "Error en la Lectura de Tarjeta.", vbInformation, "Aviso"
'            WriteToLcd "                                       "
'            WriteToLcd "Gracias por su  Preferencia..."
'            FinalizaPinPad
'            Me.Caption = sCaption
'            Exit Sub
'        End If
'
'        Me.Caption = "Ingrese la Clave de la Tarjeta."
'        WriteToLcd "                                       "
'        WriteToLcd "Ingrese Clave"
'        sClaveTar = GetClaveTarjeta
'        If clsGen.ValidaClaveTarjeta(sNumTar, sClaveTar) Then
'            Dim clsMant As NCapMantenimiento
'            Dim rsTarj As Recordset
'
'            Set clsMant = New NCapMantenimiento
'            Set rsTarj = clsMant.GetTarjetaCuentas(sNumTar)
'            If rsTarj.EOF And rsTarj.BOF Then
'                MsgBox "Tarjeta no posee ninguna relación con cuentas activas o Tarjeta no activa.", vbInformation, "Aviso"
'                WriteToLcd "                                       "
'                WriteToLcd "Gracias por su  Preferencia..."
'                FinalizaPinPad
'                Me.Caption = sCaption
'                Exit Sub
'            Else
'                nEstado = rsTarj("nEstado")
'                If nEstado = gCapTarjEstBloqueada Or nEstado = gCapTarjEstCancelada Then
'                    If nEstado = gCapTarjEstBloqueada Then
'                        MsgBox "Número de Tarjeta Bloqueada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
'                    ElseIf nEstado = gCapTarjEstCancelada Then
'                        MsgBox "Número de Tarjeta Cancelada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
'                    End If
'                    WriteToLcd "                                       "
'                    WriteToLcd "Gracias por su  Preferencia..."
'                    FinalizaPinPad
'                    Me.Caption = sCaption
'                    Exit Sub
'                End If
'
'                Dim rsPers As Recordset
'                Dim sCta As String, sRelac As String, sEstado As String
'                Dim clsCuenta As UCapCuentas
'
'                Set rsPers = clsMant.GetCuentasPersona(rsTarj("cPersCod"), nProducto)
'                Set clsMant = Nothing
'                If Not (rsPers.EOF And rsPers.EOF) Then
'                    Do While Not rsPers.EOF
'                        sCta = rsPers("cCtaCod")
'                        sRelac = rsPers("cRelacion")
'                        sEstado = Trim(rsPers("cEstado"))
'                        frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
'                        rsPers.MoveNext
'                    Loop
'                    Set clsCuenta = New UCapCuentas
'                    Set clsCuenta = frmCapMantenimientoCtas.Inicia
'                    If clsCuenta.sCtaCod <> "" Then
'                        txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
'                        txtCuenta.Prod = Mid(clsCuenta.sCtaCod, 6, 3)
'                        txtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
'                        txtCuenta.SetFocusCuenta
'                        SendKeys "{Enter}"
'                    End If
'                    Set clsCuenta = Nothing
'                Else
'                    MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
'                End If
'                rsPers.Close
'                Set rsPers = Nothing
'            End If
'            Set rsTarj = Nothing
'            Set clsMant = Nothing
'
'        Else
'            WriteToLcd "                                       "
'            WriteToLcd "Clave Incorrecta"
'            MsgBox "Clave Incorrecta", vbInformation, "Aviso"
'        End If
'        Set clsGen = Nothing
'        WriteToLcd "                                       "
'        WriteToLcd "Gracias por su  Preferencia..."
'        FinalizaPinPad
'    End If
End Sub


Private Sub Form_Load()
    Call CargaControles
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    bCalenDinamic = False
    CentraSdi Me
End Sub



Private Sub LstCred_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
            ActxCta.NroCuenta = LstCred.Text
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub TxtMonPag_Change()
Dim oNegCredito As NCredito

'CAMBIO A COM+
'    Set oNegCredito = New NCredito
'    MatCalendDistribuido = oNegCredito.CrearMatrizparaAmortizacion(MatCalend)
'    Set oNegCredito = Nothing
    LblNewSalCap.Caption = ""
    LblNewCPend.Caption = ""
    LblProxfec.Caption = ""
    LblEstado.Caption = ""
    'LblItf.Caption = "0.00"
    cmdGrabar.Enabled = False
End Sub

Private Sub TxtMonPag_GotFocus()
    fEnfoque TxtMonPag
End Sub

Private Sub TxtMonPag_KeyPress(KeyAscii As Integer)
'Dim oNegCredito As NCredito
'Dim oGastos As nGasto
'Dim nMontoGastoGen As Double
'Dim odCredito As DCOMCalendario
'Dim nInteresFecha As Double
'
'    KeyAscii = NumerosDecimales(TxtMonPag, KeyAscii, 15)
'    If KeyAscii = 13 Then
'        If CDbl(TxtMonPag.Text) = 0 Then
'            MsgBox "Monto de Pago Debe ser mayor que Cero", vbQuestion, "Aviso"
'            Exit Sub
'        End If
'        If CDbl(TxtMonPag.Text) > CDbl(LblTotDeuda.Caption) Then
'            Set odCredito = New DCOMCalendario
'                If odCredito.ObtieneMonto_Validate(ActxCta.NroCuenta, TxtMonPag) = True Then
'                   MsgBox "Monto de Pago es mayor que la deuda", vbQuestion, "Aviso"
'                    Exit Sub
'                Else
'                    MsgBox "Monto de pago sobrepasa el total", vbInformation, "Aviso"
'                    Exit Sub
'                End If
'        End If
'
'        TxtMonPag.Text = Format(TxtMonPag.Text, "#0.00")
'
'        nMontoPago = fgITFCalculaImpuestoNOIncluido(CDbl(TxtMonPag.Text))
'        'nITF = Format(nMontoPago - CDbl(TxtMonPag.Text), "0.00")
'        'nITF = CalculoSinRedondeo(nMontoPago - CDbl(TxtMonPag.Text))
'        If Mid(ActxCta.NroCuenta, 6, 3) = "423" Then
'            nITF = 0
'        Else
'            nITF = CalculoSinRedondeo(CDbl(TxtMonPag.Text))
'        End If
'        'LblItf.Caption = Format(nITF, "0.00")
'        LblItf.Caption = Format(nITF, "#0.00") 'CalculoSinRedondeo(nITF)
'        lblPagoTotal.Caption = Format(Val(TxtMonPag.Text) + nITF, "#0.00")
'        nMontoPago = Val(TxtMonPag.Text)
'
'        Dim plblnewsalcap As Double
'        Dim plblnewcpend As String
'        Dim plblproxfec As String
'        Dim plblestado As String
'
'        Call oCred.DistribuyeMonto(ActxCta.NroCuenta, nMontoPago, gdFecSis, lblMetLiq.Caption, LblMonCalDin.Caption, CDbl(LblTotDeuda.Caption), plblnewsalcap, plblnewcpend, plblproxfec, plblestado)
'
'        LblNewSalCap.Caption = Format(plblnewsalcap, "#0.00")
'        LblNewCPend.Caption = plblnewcpend
'        LblProxfec.Caption = plblproxfec
'        LblEstado.Caption = plblestado
'
'        If LblEstado.Caption = "CANCELADO" Then
'            LblProxfec.Caption = ""
'        End If
'        Set oNegCredito = Nothing
'
'        cmdGrabar.Enabled = True
'        cmdGrabar.SetFocus
'
'    End If
End Sub
Function CalculoSinRedondeo(ByVal pnMonto As Double) As Double
    Dim sCadena As String
    Dim intpos  As Integer
    Dim nEntera As Integer
    Dim nDecimal As Integer
    Dim lnValor As Double
    
        lnValor = pnMonto * gnITFPorcent
        lnValor = CortaDosITF(lnValor)
        lnValor = Format(lnValor, "#0.00")
        CalculoSinRedondeo = lnValor
       
End Function


Private Sub TxtMonPag_LostFocus()
    If Trim(TxtMonPag.Text) = "" Then
        TxtMonPag.Text = "0.00"
    End If

End Sub

Private Function IniciaLavDinero() As String
Dim i As Long
 
Dim oPersona As DCapMantenimiento
Dim oCta As DCredito
Dim rsPers As Recordset
Dim sPerscod As String
Dim sNombre As String
Dim sDireccion As String
Dim sDocId As String
Dim nMonto As Double
Set oCta = New DCredito
sPerscod = oCta.RecuperaTitularCredito(ActxCta.NroCuenta)
Set oCta = Nothing

Set oPersona = New DCapMantenimiento

Set rsPers = oPersona.GetDatosPersona(sPerscod)
If rsPers.BOF Then
Else

    sPerscod = sPerscod
    sNombre = rsPers!NOMBRE
    sDireccion = rsPers!Direccion
    sDocId = rsPers!id & " " & rsPers![ID N°]
End If
rsPers.Close
Set rsPers = Nothing

nMonto = CDbl(TxtMonPag.Text)

IniciaLavDinero = frmMovLavDinero.Inicia(sPerscod, sNombre, sDireccion, sDocId, True, True, nMonto, ActxCta.NroCuenta, sOperacion, , "COLOCACIONES")

End Function

Private Function EsExoneradaLavadoDinero() As Boolean
'Dim bExito As Boolean
'Dim clsExo As DCOMLavado
'bExito = True
'
'    Set clsExo = New DCOMLavado
'
'    If Not clsExo.EsPersonaExoneradaLavadoDinero(sPerscod) Then bExito = False
'
'    Set clsExo = Nothing
'    EsExoneradaLavadoDinero = bExito
    
End Function


Sub VerificaRFA(ByVal psCtaCod As String)
'    Dim objRFA As DCOMRFA
'    Set objRFA = New DCOMRFA
'    bRFA = objRFA.VerificaCreditoRFA(psCtaCod)
'    Set objRFA = Nothing
End Sub


