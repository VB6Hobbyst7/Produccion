VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColRecPagoCredRecup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocaciones - Recuperaciones : Pago de Credito"
   ClientHeight    =   6000
   ClientLeft      =   3315
   ClientTop       =   3225
   ClientWidth     =   7515
   Icon            =   "frmColRecPagoCredRecup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraSaldos 
      Caption         =   "Saldos"
      Height          =   1635
      Left            =   60
      TabIndex        =   31
      Top             =   4215
      Width           =   7395
      Begin VB.Label LblSaldNUe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   6105
         TabIndex        =   61
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Mora"
         Height          =   195
         Index           =   13
         Left            =   5040
         TabIndex        =   60
         Top             =   720
         Width           =   885
      End
      Begin VB.Label LblSaldNUe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   6105
         TabIndex        =   59
         Top             =   690
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Gastos"
         Height          =   195
         Index           =   12
         Left            =   5040
         TabIndex        =   58
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label LblSaldNUe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   6105
         TabIndex        =   57
         Top             =   435
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Interes"
         Height          =   195
         Index           =   11
         Left            =   5040
         TabIndex        =   56
         Top             =   465
         Width           =   1005
      End
      Begin VB.Label LblSaldNUe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   6105
         TabIndex        =   55
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Capital"
         Height          =   195
         Index           =   10
         Left            =   5040
         TabIndex        =   54
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Index           =   14
         Left            =   5055
         TabIndex        =   53
         Top             =   1305
         Width           =   360
      End
      Begin VB.Label lbltotalNue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6105
         TabIndex        =   52
         Top             =   1305
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Capital"
         Height          =   195
         Index           =   5
         Left            =   2880
         TabIndex        =   51
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lblDistrib 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   3615
         TabIndex        =   50
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Interes"
         Height          =   195
         Index           =   6
         Left            =   2880
         TabIndex        =   49
         Top             =   465
         Width           =   480
      End
      Begin VB.Label lblDistrib 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   3615
         TabIndex        =   48
         Top             =   435
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Gastos"
         Height          =   195
         Index           =   7
         Left            =   2880
         TabIndex        =   47
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label lblDistrib 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   3615
         TabIndex        =   46
         Top             =   690
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mora"
         Height          =   195
         Index           =   8
         Left            =   2880
         TabIndex        =   45
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblDistrib 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   3615
         TabIndex        =   44
         Top             =   945
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Index           =   9
         Left            =   2880
         TabIndex        =   43
         Top             =   1290
         Width           =   360
      End
      Begin VB.Label lbltotalDist 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3615
         TabIndex        =   42
         Top             =   1275
         Width           =   1155
      End
      Begin VB.Label lblSaldAct 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   41
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label lblSaldAct 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   40
         Top             =   435
         Width           =   1155
      End
      Begin VB.Label lblSaldAct 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   39
         Top             =   690
         Width           =   1155
      End
      Begin VB.Label lblSaldAct 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   38
         Top             =   945
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Mora"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   37
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Gastos"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   36
         Top             =   1020
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Interes"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   35
         Top             =   465
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Capital"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   34
         Top             =   210
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   33
         Top             =   1305
         Width           =   360
      End
      Begin VB.Label lblTotalAct 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   32
         Top             =   1275
         Width           =   1155
      End
   End
   Begin VB.Frame FraComandos 
      Height          =   675
      Left            =   75
      TabIndex        =   27
      Top             =   3510
      Width           =   7335
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   360
         Left            =   1755
         TabIndex        =   30
         Top             =   225
         Width           =   975
      End
      Begin VB.CommandButton cmdsalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   360
         Left            =   6000
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   540
         TabIndex        =   28
         Top             =   225
         Width           =   975
      End
      Begin VB.Label lblEstadonew 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   3080
         TabIndex        =   70
         Top             =   255
         Width           =   2600
      End
   End
   Begin VB.Frame FraPago 
      Height          =   1170
      Left            =   60
      TabIndex        =   5
      Top             =   2340
      Width           =   7395
      Begin VB.TextBox txtNumDoc 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.ComboBox cboModalidad 
         Height          =   315
         ItemData        =   "frmColRecPagoCredRecup.frx":030A
         Left            =   1020
         List            =   "frmColRecPagoCredRecup.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   1950
      End
      Begin SICMACT.EditMoney AXMontoPago 
         Height          =   285
         Left            =   1005
         TabIndex        =   26
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney TxtTotalAPagar 
         Height          =   285
         Left            =   6210
         TabIndex        =   64
         Top             =   735
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Pagar :"
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
         Left            =   5130
         TabIndex        =   65
         Top             =   765
         Width           =   1125
      End
      Begin VB.Label LblITF 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3690
         TabIndex        =   63
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "I.T.F."
         Height          =   195
         Index           =   7
         Left            =   3060
         TabIndex        =   62
         Top             =   765
         Width           =   375
      End
      Begin VB.Label LblNumDoc 
         AutoSize        =   -1  'True
         Caption         =   "No Doc"
         Height          =   195
         Left            =   3075
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pago :"
         Height          =   195
         Left            =   105
         TabIndex        =   9
         Top             =   345
         Width           =   825
      End
      Begin VB.Label lblMonMonto 
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   4680
         TabIndex        =   7
         Top             =   780
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   765
         Width           =   540
      End
   End
   Begin VB.TextBox Txttemp 
      Height          =   405
      Left            =   8430
      TabIndex        =   4
      Tag             =   "txtcodigo"
      Top             =   375
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Frame FraCredito 
      Caption         =   "Credito"
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
      Height          =   2250
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   7365
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   465
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   820
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar..."
         Height          =   360
         Left            =   6240
         TabIndex        =   0
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label lblCuotaNeg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3825
         TabIndex        =   69
         Top             =   1890
         Width           =   1365
      End
      Begin VB.Label Label7 
         Caption         =   "Cuota Negoc"
         Height          =   195
         Index           =   22
         Left            =   2610
         TabIndex        =   68
         Top             =   1935
         Width           =   1020
      End
      Begin VB.Label lblNroNeg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   900
         TabIndex        =   67
         Top             =   1890
         Width           =   1365
      End
      Begin VB.Label Label7 
         Caption         =   "Negociac"
         Height          =   195
         Index           =   21
         Left            =   90
         TabIndex        =   66
         Top             =   1905
         Width           =   1095
      End
      Begin VB.Label lblDemanda 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6480
         TabIndex        =   25
         Top             =   810
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Demanda"
         Height          =   195
         Index           =   1
         Left            =   5670
         TabIndex        =   24
         Top             =   810
         Width           =   690
      End
      Begin VB.Label lblComision 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6480
         TabIndex        =   23
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   810
         Width           =   480
      End
      Begin VB.Label lblCliente 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   900
         TabIndex        =   21
         Top             =   810
         Width           =   4305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Condicion"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   20
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label lblCondicion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   900
         TabIndex        =   19
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cobranza"
         Height          =   195
         Index           =   3
         Left            =   2610
         TabIndex        =   18
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label lblTipoCobranza 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3780
         TabIndex        =   17
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Comision"
         Height          =   195
         Index           =   6
         Left            =   5670
         TabIndex        =   16
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Abogado"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   15
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label lblEstudioJur 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   900
         TabIndex        =   14
         Top             =   1560
         Width           =   4305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Met.Liquid."
         Height          =   195
         Index           =   4
         Left            =   5670
         TabIndex        =   13
         Top             =   1290
         Width           =   780
      End
      Begin VB.Label lblMetLiquid 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6480
         TabIndex        =   12
         Top             =   1200
         Width           =   780
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   300
      Left            =   8385
      TabIndex        =   8
      Top             =   885
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   529
      _Version        =   393217
      TextRTF         =   $"frmColRecPagoCredRecup.frx":030E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmColRecPagoCredRecup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* RECUPERACIONES - PAGO DE CREDITOS EN RECUPERACIONES
'Archivo:  frmColRecPagoCredRecup.frm
'LAYG   :  15/08/2001.
'Resumen:  Nos permite registrar los pagos de Creditos en Recuperaciones

Option Explicit

Dim fnVarOpeCod As Long
Dim fsVarOpeDesc As String
Dim fsVarPersCodCMAC As String
Dim fsVarNombreCMAC As String

Dim fnSaldoCap As Currency, fnSaldoIntComp As Currency, fnSaldoIntMorat As Currency, fnSaldoGasto As Currency
Dim fnCapPag As Currency, fnIntCompPag As Currency, fnIntMoratPag As Currency, fnGastoPag As Currency
Dim fnNewSaldoCap As Currency, fnNewSaldoIntComp As Currency, fnNewSaldoIntMorat As Currency, fnNewSaldoGasto As Currency
Dim fnSaldoIntCompGen As Currency, fnNewSaldoIntCompGen As Currency
Dim fnSaldoIntMoraGen As Currency, fnNewSaldoIntMoraGen As Currency 'DAOR 20070802
Dim fnNroUltGastoCta As Integer, fnNroCalend As Integer
Dim fmMatGastos As Variant ' 1=nNroGastoCta//2=nMonto//3=nMontoPagado//4=nColocRecGastoEstado//5=Modificado

Dim fnPorcComision As Double
Dim fnComisionAbog As Currency
Dim fnIntCompGenerado As Currency, fnIntMoraGenerado As Currency

Dim fnGastoAdminAdicional As Currency
Dim fnMontoPagar As Currency
Dim fsFecUltPago As String
Dim fnTasaInt As Double, fnTasaIntMorat As Double
Dim fsCondicion As String, fsDemanda As String
Dim sPersCod As String
Dim lsFechaHoraGrab As String

Dim fnTipoCalcIntComp As Integer, fnTipoCalcIntMora As Integer
'0 --> No Calcula
'1 --> Capital
'2 --> Capital + Int Comp
'3 --> Capital + Int comp + Int Morat

Dim fnFormaCalcIntComp As Integer, fnFormaCalcIntMora As Integer
'0 INTERES SIMPLE
'1 INTERES COMPUESTO
Private Type TPlanPagosNeg
    NumCuota As Integer
    FecVenc As Date
    Monto As Double
    MontoPag As Double
    Estado As String * 1
    Modificado As Boolean
End Type
Dim MatPagos() As TPlanPagosNeg
Dim MatPagosTempo() As TPlanPagosNeg
Dim fnContCuotas As Integer
Dim fnDiasAtraso As Integer
Dim fnEstadoNew As ColocEstado
Dim fnEstadoIni As ColocEstado

Dim fnRegCancelacion As Integer '** Juez 20120423

Dim MatDatos As Variant

'**DAOR 20070416************************************************************
Dim fnCapDist As Currency, fnIntCompDist As Currency, fnIntMoratDist As Currency
Dim fnGastoDist As Currency, fnComisionAbogDist As Currency
Dim fbExisteDistribucionCIMG As Boolean
'***************************************************************************
'Dim ventana As Integer ' MADM 20090929
Dim nRedondeoITF As Double 'BRGO 20110914
Dim oDocRec As UDocRec 'EJVG20140408
'RIRO20140530 ERS017 ********************
Dim nMovNroRVD As Long
Dim nMovNroRVDPen As Long
Dim nMontoVoucher As Currency
'END RIRO *******************************

Public Sub Inicio(ByVal pnOpeCod As Long, ByVal psOpeDesc As String, _
        ByVal psPersCodCMAC As String, ByVal psNomCmac As String, ByVal pbMuestraSaldos As Boolean)

    fnVarOpeCod = pnOpeCod
    fsVarOpeDesc = psOpeDesc
    fsVarPersCodCMAC = psPersCodCMAC
    fsVarNombreCMAC = psNomCmac
    
    Select Case fnVarOpeCod
    '    Case gColPOpeCancelEFE
    '        'txtDocumento.Visible = false
    '    Case gColPOpeCancelCHQ
    '        'txtDocumento.Visible = True
    '    Case Else
    '        txtDocumento.Visible = False
    End Select
    If pbMuestraSaldos = False Then
        Me.Height = 4700
    End If
    CargaParametros
    Limpiar
    Me.Show 1
End Sub

Private Sub HabilitaControles(ByVal pbCmdGrabar As Boolean, ByVal pbCmdCancelar As Boolean, _
            ByVal pbCmdSalir As Boolean)
    cmdGrabar.Enabled = pbCmdGrabar
    cmdCancelar.Enabled = pbCmdCancelar
    cmdsalir.Enabled = pbCmdSalir
End Sub

Private Sub Limpiar()
Dim lnI As Integer
    Me.AXCodCta.Enabled = True
    Me.CmdBuscar.Enabled = True
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
    Me.lblCliente.Caption = ""
    Me.lblDemanda.Caption = ""
    Me.lblCondicion.Caption = ""
    Me.lblTipoCobranza.Caption = ""
    Me.lblMetLiquid.Caption = ""
    Me.lblEstudioJur.Caption = ""
    Me.lblComision.Caption = ""
    Me.lblEstadonew = ""
    
    cboModalidad.ListIndex = -1
    txtNumDoc.Text = ""
    Me.AXMontoPago.Text = 0
    lblNroNeg = "": fnDiasAtraso = 0
    Me.lblCuotaNeg = 0
    For lnI = 0 To 3
        Me.lblSaldAct(lnI) = 0
        Me.lblDistrib(lnI) = 0
        Me.LblSaldNUe(lnI) = 0
    Next
    Me.lblTotalAct = 0
    Me.lbltotalDist = 0
    Me.lbltotalNue = 0
    Me.TxtTotalAPagar = 0
    Me.LblITF = 0
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaCredito (AXCodCta.NroCuenta)
End Sub

Sub ReCalCulaGasto(ByVal psCtaCod As String)
    Dim lrDatGastos As New ADODB.Recordset
    Dim loValCred As COMDColocRec.DCOMColRecCredito
    Dim lsMensaje As String
    
    Set loValCred = New COMDColocRec.DCOMColRecCredito
    Set lrDatGastos = loValCred.dObtieneListaGastosxCredito(psCtaCod, lsMensaje)
    If Trim(lsMensaje) <> "" Then
        MsgBox lsMensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    If lrDatGastos Is Nothing Then   ' Hubo un Error
        'MsgBox "No se encontro el Credito ", vbInformation, "Aviso"
       ' Limpiar
       ' Set lrDatCredito = Nothing
        Exit Sub
    End If
    
    Set fmMatGastos = Nothing
    
        Dim i As Integer
        ReDim fmMatGastos(0)
        ReDim fmMatGastos(lrDatGastos.RecordCount, 11)
        
        Do While Not lrDatGastos.EOF
            If lrDatGastos!nColocRecGastoEstado = gColRecGastoEstPendiente Then
                fmMatGastos(i, 1) = lrDatGastos!nNroGastoCta
                fmMatGastos(i, 2) = lrDatGastos!nMonto
                fmMatGastos(i, 3) = lrDatGastos!nMontoPagado
                fmMatGastos(i, 4) = lrDatGastos!nColocRecGastoEstado
                fmMatGastos(i, 5) = "N" ' Estado del Gasto
                fmMatGastos(i, 6) = 0 '(fmMatGastos(i, 2) - fmMatGastos(i, 3)) 'avmm 0   ' Monto a Cubrir del Gasto
                fmMatGastos(i, 7) = lrDatGastos!nPrdConceptoCod
                'nMontoSaldo = nMontoSaldo + (fmMatGastos(i, 2) - fmMatGastos(i, 3))
                i = i + 1
            End If
            lrDatGastos.MoveNext
        Loop
    
End Sub

Private Sub BuscaCredito(ByVal psCtaCod As String)
Dim lbOk As Boolean
Dim lrDatCredito As New ADODB.Recordset
Dim lrDatGastos As New ADODB.Recordset
Dim loValCred As COMDColocRec.DCOMColRecCredito
Dim lnDiasUltTrans As Integer
Dim loCalcula As COMNColocRec.NCOMColRecCalculos
Dim lrCIMG As ADODB.Recordset 'DAOR 20070416
'----- MADM
Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset
'----- MADM
Dim lsMensaje As String

'******RECO 2013-07-22**********
Dim loGrabar As COMNColocRec.NCOMColRecCredito
'******END RECO*****************



On Error GoTo ControlError
    
    'FRHU 20150415 ERS022-2015
    If VerificarSiEsUnCreditoTransferido(psCtaCod) Then
        MsgBox "El Credito seleccionado se encuentra en estado Transferido"
        Exit Sub
    End If
    'FIN FRHU 20150415
    '**************RECO 2013-07-22********
    Set loGrabar = New COMNColocRec.NCOMColRecCredito
        If loGrabar.nValidaDiaAutorizacionPagoJud(psCtaCod, gdFecSis) = False Then
            MsgBox "La cuenta no tiene una autorización de pago valida." & Chr(13) & _
            "No es posible realizar el pago porque es necesario una coordicacíon previa con el área de Recuperaciones", vbInformation, " Aviso "
            Exit Sub
        End If
    '****************END RECO*************


    'Carga Datos
    fnRegCancelacion = 0 ' Juez 20120713
    
    Set loValCred = New COMDColocRec.DCOMColRecCredito
        Set lrDatCredito = loValCred.dObtieneDatosPagoCredRecup(psCtaCod, lsMensaje)
        If Trim(lsMensaje) <> "" Then
            MsgBox lsMensaje, vbInformation, "Aviso"
            Exit Sub
        End If
        Set lrDatGastos = loValCred.dObtieneListaGastosxCredito(psCtaCod, lsMensaje)
        If Trim(lsMensaje) <> "" Then
            MsgBox lsMensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    '**DAOR 20070416**************************************************************
    Set lrCIMG = loValCred.dObtieneDistribucionCIMGCobranza(psCtaCod, gdFecSis)
    '*****************************************************************************
    Set loValCred = Nothing
    
    If lrDatCredito Is Nothing Then   ' Hubo un Error
        MsgBox "No se encontro el Credito ", vbInformation, "Aviso"
        Limpiar
        Set lrDatCredito = Nothing
        Exit Sub
    End If
        ' Asigna Valores a las Variables
        fnSaldoCap = lrDatCredito!nSaldo
        fnSaldoIntComp = lrDatCredito!nSaldoIntComp
        fnSaldoIntMorat = lrDatCredito!nSaldoIntMor
        fnSaldoGasto = lrDatCredito!nSaldoGasto
        
        fnSaldoIntCompGen = lrDatCredito!nIntCompGen
        fnSaldoIntMoraGen = IIf(IsNull(lrDatCredito!nIntMoraGen), 0, lrDatCredito!nIntMoraGen) 'DAOR 20070809
        
        fnNroUltGastoCta = lrDatCredito!nUltGas
        fsFecUltPago = CDate(fgFechaHoraGrab(lrDatCredito!cUltimaActualizacion))
        fnNroCalend = lrDatCredito!nNroCalen
        fnEstadoNew = lrDatCredito!nPrdEstado
        fnEstadoIni = lrDatCredito!nPrdEstado
        Select Case fnEstadoIni
            Case 2201
                lblEstadonew = "Vigente Judicial"
            Case 2202
                lblEstadonew = "Vigente Castigado"
            Case 2205
                lblEstadonew = "Refinanciado Judicial"
            Case 2206
                lblEstadonew = "Refinanciado Castigado"
        End Select
        
        fnTasaInt = IIf(IsNull(lrDatCredito!nTasaInt), 0, lrDatCredito!nTasaInt)
        
        fnTasaIntMorat = lrDatCredito!nTasaIntMor
        lnDiasUltTrans = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(fsFecUltPago, "dd/mm/yyyy"))
        
        'Muestra Datos
        sPersCod = lrDatCredito!cPersCod
        Me.lblCliente.Caption = PstaNombre(Trim(lrDatCredito!cPersNombre))
        Me.lblDemanda.Caption = IIf(lrDatCredito!nDemanda = gColRecDemandaSi, "S", "N")
        Me.lblCondicion = fgCondicionColRecupDesc(lrDatCredito!nPrdEstado)
        Me.lblTipoCobranza = IIf(lrDatCredito!nTipCJ = gColRecTipCobJudicial, "Judicial", "ExtraJudicial")
        Me.lblMetLiquid = lrDatCredito!cMetLiquid
        Me.lblEstudioJur.Caption = lrDatCredito!cPersNombreAbog
        Me.lblComision = lrDatCredito!nValorCom
        fnPorcComision = lrDatCredito!nValorCom
        fsCondicion = IIf(lrDatCredito!nPrdEstado = gColocEstRecVigJud Or lrDatCredito!nPrdEstado = 2205, "J", "A")
        fsDemanda = IIf(lrDatCredito!nDemanda = gColRecDemandaSi, "S", "N")
        Me.lblNroNeg = IIf(IsNull(lrDatCredito!cNroNeg), "", lrDatCredito!cNroNeg)
        
        '*******
        'Agregado por LMMD
        Dim nMontoSaldo As Double
        '*** Carga Gastos en Matriz
        Dim i As Integer
        ReDim fmMatGastos(0)
        ReDim fmMatGastos(lrDatGastos.RecordCount, 11)
        Do While Not lrDatGastos.EOF
            If lrDatGastos!nColocRecGastoEstado = gColRecGastoEstPendiente Then
                fmMatGastos(i, 1) = lrDatGastos!nNroGastoCta
                fmMatGastos(i, 2) = lrDatGastos!nMonto
                fmMatGastos(i, 3) = lrDatGastos!nMontoPagado
                fmMatGastos(i, 4) = lrDatGastos!nColocRecGastoEstado
                fmMatGastos(i, 5) = "N" ' Estado del Gasto
                fmMatGastos(i, 6) = 0 '(fmMatGastos(i, 2) - fmMatGastos(i, 3)) 'avmm 0  ' Monto a Cubrir del Gasto
                fmMatGastos(i, 7) = lrDatGastos!nPrdConceptoCod
                nMontoSaldo = nMontoSaldo + (fmMatGastos(i, 2) - fmMatGastos(i, 3))
                i = i + 1
            End If
            lrDatGastos.MoveNext
        Loop
        fnSaldoGasto = nMontoSaldo
        
        'Calcula el Int Comp Generado
        Set loCalcula = New COMNColocRec.NCOMColRecCalculos
            If fnTipoCalcIntComp = 0 Then ' NoCalcula
                fnIntCompGenerado = 0
            ElseIf fnTipoCalcIntComp = 1 Then ' En base al capital
                If fnFormaCalcIntComp = 1 Then 'INTERES COMPUESTO
                    fnIntCompGenerado = loCalcula.nCalculaIntCompGenerado(lnDiasUltTrans, fnTasaInt, fnSaldoCap)
                Else
                    'INTERES SIMPLE
                    fnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, fnTasaInt, fnSaldoCap)
                End If
            ElseIf fnTipoCalcIntComp = 2 Then ' En base a capit + int Comp
                If fnFormaCalcIntComp = 1 Then
                    fnIntCompGenerado = loCalcula.nCalculaIntCompGenerado(lnDiasUltTrans, fnTasaInt, fnSaldoCap + fnSaldoIntComp)
                Else
                    fnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, fnTasaInt, fnSaldoCap + fnSaldoIntComp)
                End If
            ElseIf fnTipoCalcIntComp = 3 Then ' En base a capit + int Comp + int Morat
                If fnFormaCalcIntComp = 1 Then
                    fnIntCompGenerado = loCalcula.nCalculaIntCompGenerado(lnDiasUltTrans, fnTasaInt, fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat)
                Else
                    fnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, fnTasaInt, fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat)
                End If
            End If
            If fnTipoCalcIntMora = 0 Then  ' NoCalcula
                fnIntMoraGenerado = 0
            ElseIf fnTipoCalcIntMora = 1 Then ' En base al capital
                If fnFormaCalcIntMora = 1 Then 'INTERES COMPUESTO
                    fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap)
                Else
                    'INTERES SIMPLE
                    fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap)
                End If
            ElseIf fnTipoCalcIntMora = 2 Then ' En base a capit + int Comp
                If fnFormaCalcIntMora = 1 Then
                    fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap + fnSaldoIntComp)
                Else
                    fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap + fnSaldoIntComp)
                End If
            ElseIf fnTipoCalcIntMora = 3 Then ' En base a capit + int Comp + int Morat
                If fnFormaCalcIntMora = 1 Then
                    fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat)
                Else
                    fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat)
                End If
            End If
              
        Set loCalcula = Nothing
        'Agregamos el Int Calculado al Saldo Int Comp
        fnSaldoIntComp = lrDatCredito!nSaldoIntComp + fnIntCompGenerado
        fnSaldoIntMorat = lrDatCredito!nSaldoIntMor + fnIntMoraGenerado
                
        Me.lblSaldAct(0) = Format(fnSaldoCap, "##0.00")
        'Me.lblSaldAct(1) = Format(fnSaldoIntComp + fnIntCompGenerado, "##0.00")
        Me.lblSaldAct(1) = Format(fnSaldoIntComp, "##0.00")
        'Me.lblSaldAct(2) = Format(fnSaldoIntMorat + fnIntMoraGenerado, "##0.00")
        Me.lblSaldAct(2) = Format(fnSaldoIntMorat, "##0.00")
        Me.lblSaldAct(3) = Format(fnSaldoGasto, "##0.00")
        Me.lblTotalAct = Format(fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat + fnSaldoGasto, "##0.00")
        'Me.lblTotalAct = Format(fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat + fnSaldoGasto + fnIntCompGenerado + fnIntMoraGenerado, "##0.00")
        '***
        
        '**DAOR 20070416********************************************************
        Me.AXMontoPago.Text = "0.00"
        If Not lrCIMG.EOF Then
            fnCapDist = lrCIMG!nCapital
            fnIntCompDist = lrCIMG!nIntComp
            fnIntMoratDist = lrCIMG!nMora
            fnGastoDist = lrCIMG!nGasto
            fnComisionAbogDist = lrCIMG!nComiAbog
            fbExisteDistribucionCIMG = True
            AXMontoPago.Text = Format(fnCapDist + fnIntCompDist + fnIntMoratDist + fnGastoDist + fnComisionAbogDist, "#0.00")
            AXMontoPago.Enabled = False
            If Not IsNull(lrCIMG!nRegCancelacion) Then
                fnRegCancelacion = lrCIMG!nRegCancelacion
                MsgBox "Al grabar el pago, el crédito será cancelado como lo estableció en el Area de Recuperaciones", vbInformation, "Aviso"
            End If
            Call AXMontoPago_KeyPress(13)
        Else
            fbExisteDistribucionCIMG = False
        End If
        '***********************************************************************
    Set lrDatCredito = Nothing
    Set lrDatGastos = Nothing
    Set lrCIMG = Nothing
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
        
      '************ firma madm
       ' If ventana = 0 Then
         Set lafirma = New frmPersonaFirma
         Set ClsPersona = New COMDPersona.DCOMPersonas
        
         Set Rf = ClsPersona.BuscaCliente(sPersCod, BusquedaCodigo)
         If Not Rf.BOF And Not Rf.EOF Then
            If Rf!nPersPersoneria = 1 Then
            Call frmPersonaFirma.Inicio(Trim(sPersCod), Mid(sPersCod, 4, 2), False, True)
            End If
         End If
         Set Rf = Nothing
       ' End If
      '  ventana = 0
        '************ firma madm

    AXCodCta.Enabled = False
    CmdBuscar.Enabled = False ' RIRO20131102
    Call HabilitaControles(True, True, True)
        
   'Carga el Plan de Pagos Si tiene Negociacion.
   If Len(lblNroNeg.Caption) > 0 Then
        Call CargaPlanPagos(psCtaCod, lblNroNeg)
        lblCuotaNeg.Caption = Format(MontoPendienteNeg, "#.00")
    End If
    
    'Comentado por DAOR 20070416
    'Me.AXMontoPago.Text = "0.00"
    Me.cboModalidad.SetFocus
    

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AXMontoPago_KeyPress(KeyAscii As Integer)
Dim lnITF As Double
If KeyAscii = 13 Then
    If Val(AXMontoPago.Text) = 0 Then
        Exit Sub
    End If
    lnITF = gITF.fgITFCalculaImpuesto(CCur(AXMontoPago.Text))
    '*** BRGO 20110908 ************************************************
    nRedondeoITF = fgDiferenciaRedondeoITF(lnITF)
    If nRedondeoITF > 0 Then
        lnITF = lnITF - nRedondeoITF
    End If
    '*** END BRGO
    LblITF = Format(lnITF, "#,##0.00")
    TxtTotalAPagar = Format(CCur(AXMontoPago.Text) + lnITF, "#0.00")
    
    CalculaDistribucionPago
    cmdGrabar.Enabled = True
    If cmdGrabar.Enabled And cmdGrabar.Visible Then cmdGrabar.SetFocus
End If
End Sub

Private Sub CalculaDistribucionPago()
   fnNewSaldoCap = 0: fnNewSaldoIntComp = 0: fnNewSaldoIntMorat = 0: fnNewSaldoGasto = 0
    fnCapPag = 0: fnIntCompPag = 0: fnIntMoratPag = 0: fnGastoPag = 0
    ' Distribuye Monto
    Call DistribuyePago
   
    ' *****
    ' Actualiza la Deuda con la Comision
    Me.lblSaldAct(3) = fnSaldoGasto + fnComisionAbog
    Me.lblTotalAct = Format(fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat + fnSaldoGasto + fnComisionAbog, "##0.00")
    
    fnNewSaldoCap = fnSaldoCap - fnCapPag
'    If fnSaldoIntComp = fnIntCompPag Then
'        fnNewSaldoIntComp = (fnSaldoIntComp + fnIntCompGenerado) - (fnIntCompPag + fnIntCompGenerado)
'    Else
'        fnNewSaldoIntComp = (fnSaldoIntComp + fnIntCompGenerado) - fnIntCompPag
'    End If
    fnNewSaldoIntMorat = fnSaldoIntMorat - fnIntMoratPag
    fnNewSaldoGasto = fnSaldoGasto - fnGastoPag
    'fnNewSaldoIntCompGen = fnSaldoIntCompGen + fnIntCompGenerado + fnIntMoraGenerado
    'Comentado por DAOR 20070808
    'fnNewSaldoIntCompGen = fnSaldoIntCompGen + fnIntCompGenerado
    
    '**DAOR 20070808***************
    fnNewSaldoIntCompGen = fnIntCompGenerado
    fnNewSaldoIntMoraGen = fnIntMoraGenerado
    '******************************
    
    fnNewSaldoIntComp = fnSaldoIntComp - fnIntCompPag
    
    Me.lblDistrib(0) = Format(fnCapPag, "##0.00")
    If Me.lblSaldAct(1) = (fnIntCompPag + fnIntCompGenerado) Then
        Me.lblDistrib(1) = Format(fnIntCompPag + fnIntCompGenerado, "##0.00")
    Else
        Me.lblDistrib(1) = Format(fnIntCompPag, "##0.00")
    End If
    Me.lblDistrib(2) = Format(fnIntMoratPag, "##0.00")
    Me.lblDistrib(3) = Format(fnGastoPag + fnComisionAbog, "##0.00") ' (Comsion abogado)
    If Me.lblSaldAct(1) = (fnIntCompPag + fnIntCompGenerado) Then
        Me.lbltotalDist = Format(fnCapPag + fnIntCompPag + fnIntMoratPag + fnGastoPag + fnComisionAbog + fnIntCompGenerado, "##0.00")
    Else
      Me.lbltotalDist = Format(fnCapPag + fnIntCompPag + fnIntMoratPag + fnGastoPag + fnComisionAbog, "##0.00")
    End If
    Me.LblSaldNUe(0) = Format(fnNewSaldoCap, "##0.00")
    Me.LblSaldNUe(1) = Format(fnNewSaldoIntComp, "##0.00")
    Me.LblSaldNUe(2) = Format(fnNewSaldoIntMorat, "##0.00")
    Me.LblSaldNUe(3) = Format(fnNewSaldoGasto, "##0.00")
    Me.lbltotalNue = Format(fnNewSaldoCap + fnNewSaldoIntComp + fnNewSaldoIntMorat + fnNewSaldoGasto, "##0.00")
    
    '** Comentado x Juez 20120423
    'If CCur(lbltotalDist) >= CCur(lblTotalAct) Then 'actualiza el estado del credito si cancela el total de conceptos
        'Select Case fnEstadoIni
            'Case 2201, 2205
            '    fnEstadoNew = 2203
            '    lblEstadonew = "Cancelado Judicial"
            'Case 2202, 2206
            '    fnEstadoNew = 2204
            '    lblEstadonew = "Cancelado Castigado"
        'End Select
    'Else
        'fnEstadoNew = fnEstadoIni
        'Select Case fnEstadoNew
            'Case 2201
            '    lblEstadonew = "Vigente Judicial"
            'Case 2202
            '    lblEstadonew = "Vigente Castigado"
            'Case 2205
            '    lblEstadonew = "Refinanciado Judicial"
            'Case 2206
            '    lblEstadonew = "Refinanciado Castigado"
        'End Select
    'End If
    '** Juez 20120423 ************************************************
    If fnRegCancelacion = 1 Then
        Select Case fnEstadoIni
            Case 2201, 2205
                fnEstadoNew = 2203
            Case 2202, 2206
                fnEstadoNew = 2204
        End Select
    Else
        fnEstadoNew = fnEstadoIni
    End If
    '**End Juez ******************************************************
End Sub
Private Sub DistribuyePago()
Dim lnMontoDistrib As Currency
Dim lsPrio1 As String, lsPrio2 As String, lsPrio3 As String, lsPrio4 As String
Dim loCalculaComision As COMNColocRec.NCOMColRecCalculos
Dim lnComiCMI As Double
lsPrio1 = Mid(Me.lblMetLiquid, 1, 1)
lsPrio2 = Mid(Me.lblMetLiquid, 2, 1)
lsPrio3 = Mid(Me.lblMetLiquid, 3, 1)
lsPrio4 = Mid(Me.lblMetLiquid, 4, 1)

fnCapPag = 0: fnIntCompPag = 0: fnIntMoratPag = 0: fnGastoPag = 0

lnMontoDistrib = Format(Me.AXMontoPago.Text, "#0.00")
fnMontoPagar = Format(Me.AXMontoPago.Text, "#0.00")

'Comentado por DAOR 20070809***********************
'Call ReCalCulaGasto(AXCodCta.NroCuenta)
'**************************************************

'Comentado por DAOR 20070424
'Call DistribuyePagoCominAbogado
''  Calcula Comision Abogado
'    fnComisionAbog = 0
'    'Comentado por DAOR 20070424
'    'lnComiCMI = fnCapPag + fnIntCompPag + fnIntMoratPag
'
'    '**DAOR 20070424***********************************************
'    lnComiCMI = CDbl(Me.AXMontoPago.Text) - fnGastoPag
'    '**************************************************************
'    Set loCalculaComision = New COMNColocRec.NCOMColRecCalculos
'        fnComisionAbog = loCalculaComision.nCalculaComisionAbogado(fnPorcComision, lnComiCMI)
'        fnComisionAbog = Round(fnComisionAbog, 2)
'    Set loCalculaComision = Nothing
'
'' Comision abogado
'   fnMontoPagar = Format(Me.AXMontoPago.Text, "#0.00")
'   fnMontoPagar = fnMontoPagar - fnComisionAbog
'Call DistribuyePagoCominAbogado

'**DAOR 20070424***************************************************************************
If Not fbExisteDistribucionCIMG Then
    Call DistribuyePagoCominAbogado
    'Calcula Comision Abogado
    fnComisionAbog = 0
    lnComiCMI = CDbl(Me.AXMontoPago.Text) - fnGastoPag
    Set loCalculaComision = New COMNColocRec.NCOMColRecCalculos
        fnComisionAbog = loCalculaComision.nCalculaComisionAbogado(fnPorcComision, lnComiCMI)
        fnComisionAbog = Round(fnComisionAbog, 2)
    Set loCalculaComision = Nothing
    fnMontoPagar = Format(Me.AXMontoPago.Text, "#0.00")
    fnMontoPagar = fnMontoPagar - fnComisionAbog
    Call ReCalCulaGasto(AXCodCta.NroCuenta)
    Call DistribuyePagoCominAbogado
Else
    Call EstablecerCIMGPersonalizado
End If
'******************************************************************************************
    
'        If fnMontoPagar > 0 Then
'            Select Case lsPrio1
'                Case "G": Call CubrirGastos
'                Case "M": Call CubrirMora
'                Case "I": Call CubrirInteres
'                Case "C": Call CubrirCapital
'            End Select
'        End If
'        If fnMontoPagar > 0 Then
'            Select Case lsPrio2
'                Case "G": Call CubrirGastos
'                Case "M": Call CubrirMora
'                Case "I": Call CubrirInteres
'                Case "C": Call CubrirCapital
'            End Select
'        End If
'        If fnMontoPagar > 0 Then
'            Select Case lsPrio3
'                Case "G":  Call CubrirGastos
'                Case "M":  Call CubrirMora
'                Case "I":  Call CubrirInteres
'                Case "C":  Call CubrirCapital
'            End Select
'        End If
'        If fnMontoPagar > 0 Then
'            Select Case lsPrio4
'                Case "G": Call CubrirGastos
'                Case "M": Call CubrirMora
'                Case "I": Call CubrirInteres
'                Case "C": Call CubrirCapital
'            End Select
'        End If
        
    'fnMontoPagar = Format(Me.AXMontoPago.Text, "#0.00")
    
    fnGastoAdminAdicional = 0

        If lnMontoDistrib >= (fnSaldoCap + fnSaldoIntComp + fnSaldoGasto + fnSaldoIntMorat) Then
            If fnMontoPagar > 0 Then
                fnGastoAdminAdicional = Format(fnMontoPagar, "#0.00")
                'fnGastoPag = fnGastoPag + fnGastoAdminAdicional
            End If
        End If


    ' Cubre cuotas de Negociacion Si tiene Negoc
    If Len(lblNroNeg) > 0 Then
        Call CubrirCalendario(Format(Me.AXMontoPago.Text, "#0.00"))
    End If
End Sub

Private Sub CubrirCapital()
Dim loCalculaComision As COMNColocRec.NCOMColRecCalculos
Dim nComision As Double
    'Cubro Capital
    Set loCalculaComision = New COMNColocRec.NCOMColRecCalculos
   
    If fnSaldoCap > 0 And fnMontoPagar > 0 Then
        If fnMontoPagar >= fnSaldoCap Then
            fnCapPag = fnSaldoCap
'            nComision = Round(loCalculaComision.nCalculaComisionAbogado(fnPorcComision, fnCapPag), 2)
'            fnComisionAbog = fnComisionAbog + nComision
'            If fnMontoPagar >= (fnSaldoCap + nComision) Then
'                fnMontoPagar = fnMontoPagar - (fnSaldoCap + nComision)
'            Else
'                fnCapPag = fnMontoPagar - nComision
'                fnMontoPagar = 0
'            End If
             fnMontoPagar = fnMontoPagar - fnSaldoCap
        Else
'            nComision = Round(loCalculaComision.nCalculaComisionAbogado(fnPorcComision, fnMontoPagar), 2)
'            fnComisionAbog = fnComisionAbog + nComision
'            fnCapPag = fnMontoPagar - nComision
            fnCapPag = fnMontoPagar
            fnMontoPagar = 0
        End If
    End If
    
    Set loCalculaComision = Nothing

End Sub

Private Sub CubrirInteres()
Dim loCalculaComision As COMNColocRec.NCOMColRecCalculos
Dim nComision As Double

    Set loCalculaComision = New COMNColocRec.NCOMColRecCalculos
    'Cubro Interes
    If fnSaldoIntComp > 0 And fnMontoPagar > 0 Then
        If fnMontoPagar >= fnSaldoIntComp Then
            fnIntCompPag = fnSaldoIntComp
'            nComision = Round(loCalculaComision.nCalculaComisionAbogado(fnPorcComision, fnIntCompPag), 2)
'            fnComisionAbog = fnComisionAbog + nComision
'            If fnMontoPagar > (fnSaldoIntComp + nComision) Then
'                fnMontoPagar = fnMontoPagar - (fnSaldoIntComp + nComision)
'            Else
'                fnIntCompPag = fnMontoPagar - nComision
'                fnMontoPagar = 0
'            End If
            fnMontoPagar = fnMontoPagar - fnSaldoIntComp
        Else
'            nComision = Round(loCalculaComision.nCalculaComisionAbogado(fnPorcComision, fnMontoPagar), 2)
'            fnComisionAbog = fnComisionAbog + nComision
'            fnIntCompPag = fnMontoPagar - nComision
            fnIntCompPag = fnMontoPagar
            fnMontoPagar = 0
        End If
    End If
    Set loCalculaComision = Nothing
    
End Sub

Private Sub CubrirGastos()
Dim lnGastoDistrib As Currency
Dim i As Integer
    'Cubro Gastos
    If fnSaldoGasto > 0 And fnMontoPagar > 0 Then
        If fnMontoPagar >= fnSaldoGasto Then
            fnGastoPag = fnSaldoGasto
            fnMontoPagar = fnMontoPagar - fnSaldoGasto
        Else
            fnGastoPag = fnMontoPagar
            fnMontoPagar = 0
        End If
    End If
    'Actualiza la Matriz de Gastos
    lnGastoDistrib = fnGastoPag
    
    For i = 0 To UBound(fmMatGastos) - 1
        If CInt(fmMatGastos(i, 4)) = gColRecGastoEstPendiente And lnGastoDistrib > 0 _
           And (CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))) > 0 Then
            If lnGastoDistrib >= (CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))) Then
                lnGastoDistrib = lnGastoDistrib - (CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3)))
                'Actualiza el monto Pagado
                'fmMatGastos(i, 3) = CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))
                'Actualiza el estado del gasto
                fmMatGastos(i, 4) = gColRecGastoEstPagado
                fmMatGastos(i, 5) = "S" ' Si se ha modificado
                fmMatGastos(i, 6) = CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))
            Else
                'Actualiza el monto pagado
                fmMatGastos(i, 3) = Format(CDbl(fmMatGastos(i, 3)) + lnGastoDistrib, "#0.00")
                fmMatGastos(i, 4) = gColRecGastoEstPendiente
                fmMatGastos(i, 5) = "S" ' Si se ha modificado
                fmMatGastos(i, 6) = Format(lnGastoDistrib, "#0.00")
                lnGastoDistrib = 0
            End If
        End If
    Next i
   
'   Dim nMontoTemp As Double
'   For i = 0 To UBound(fmMatGastos) - 1
'    If CInt(fmMatGastos(i, 4)) = gColRecGastoEstPendiente Then
'        If lnGastoDistrib > 0 Then
'            nMontoTemp = CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))
'            If lnGastoDistrib >= nMontoTemp Then
'               fmMatGastos(i, 3) = CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))
'               lnGastoDistrib = lnGastoDistrib - (CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3)))
'               fmMatGastos(i, 4) = gColRecGastoEstPagado
'               fmMatGastos(i, 5) = "S" ' Si se ha modificado
'            Else
'              fmMatGastos(i, 3) = lnGastoDistrib
'              lnGastoDistrib = 0
'              fmMatGastos(i, 4) = gColRecGastoEstPendiente
'              fmMatGastos(i, 5) = "S" ' Si se ha modificado
'            End If
'        End If
'    End If
'   Next i
End Sub

Private Sub CubrirMora()
Dim loCalculaComision As COMNColocRec.NCOMColRecCalculos
Dim nComision As Double
    
    Set loCalculaComision = New COMNColocRec.NCOMColRecCalculos
    'Cubro Mora
    If fnSaldoIntMorat > 0 And fnMontoPagar > 0 Then
        If fnMontoPagar >= fnSaldoIntMorat Then
            fnIntMoratPag = fnSaldoIntMorat
'            nComision = Round(loCalculaComision.nCalculaComisionAbogado(fnPorcComision, fnSaldoIntMorat), 2)
'            fnComisionAbog = fnComisionAbog + nComision
'            If fnMontoPagar >= (fnSaldoIntMorat + nComision) Then
'                fnMontoPagar = fnMontoPagar - (fnSaldoIntMorat + nComision)
'            Else
'                fnIntMoratPag = fnMontoPagar - nComision
'                fnMontoPagar = 0
'            End If
            fnMontoPagar = fnMontoPagar - fnSaldoIntMorat
        Else
'            nComision = Round(loCalculaComision.nCalculaComisionAbogado(fnPorcComision, fnMontoPagar), 2)
'            fnComisionAbog = fnComisionAbog + nComision
'            fnIntMoratPag = fnMontoPagar - nComision
            fnIntMoratPag = fnMontoPagar
            fnMontoPagar = 0
        End If
    End If
End Sub


Private Sub cboModalidad_Click()
    Dim lnMontoChq As Double
    On Error GoTo ErrCboModalidad
    If Len(AXCodCta.NroCuenta) <> 18 Then Exit Sub
    txtNumDoc.Text = ""
    Set oDocRec = New UDocRec
    If Me.cboModalidad.ListIndex <> -1 Then
        If CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoCheque Then
            'By Capi 14042008 para que jale solo cheques valorizados caja general
            'MatDatos = frmBuscaCheque.BuscaCheque(gChqEstEnValorizacion, CInt(Mid(AXCodCta.NroCuenta, 9, 1)))
            'MatDatos = frmBuscaCheque.BuscaCheque(gChqEstValorizado, CInt(Mid(AXCodCta.NroCuenta, 9, 1)), 1)
            '
            '************************RECO 2013-07-23***************
            Dim lsVarCondicion As Boolean
            Dim oform As New frmChequeBusqueda 'EJVG20140228
            lsVarCondicion = False
            Do While lsVarCondicion = False
                'MatDatos = frmBuscaCheque.BuscaCheque(gChqEstValorizado, CInt(Mid(AXCodCta.NroCuenta, 9, 1)), 1)
                Set oDocRec = oform.iniciarBusqueda(Val(Mid(AXCodCta.NroCuenta, 9, 1)), TipoOperacionCheque.CRED_Pago, AXCodCta.NroCuenta)
                'If MatDatos(0) = "" Then
                If oDocRec.fsNroDoc = "" Then
                    Exit Do
                End If
                'If frmBuscaCheque.pnMontoDisponible < TxtTotalAPagar.Text Then
                lnMontoChq = DeducirMontoxITF(oDocRec.fnMonto)
                If oDocRec.fnMonto < CCur(TxtTotalAPagar.Text) Then 'EJVG20140408
                    MsgBox "No es posible realizar el pago con ese cheque porque no cuenta con saldo suficiente para realizar la operación", _
                    vbInformation, " Aviso "
                'Exit Sub
                    
                Else
                    lsVarCondicion = True
                End If
            Loop
            Set oform = Nothing
            '***********************END RECO**********************

            'If MatDatos(0) <> "" Then
            If oDocRec.fsNroDoc <> "" Then 'EJVG20140228
                'txtNumDoc.Text = MatDatos(4)
                txtNumDoc.Text = oDocRec.fsNroDoc
                'Modificado por DAOR 20070809
                'AXMontoPago.Text = MatDatos(3)
                If Not fbExisteDistribucionCIMG Then
                    'By Capi 15042008
                    'AXMontoPago.Text = MatDatos(3)
                    'AXMontoPago.Text = MatDatos(0)
                    AXMontoPago.Text = lnMontoChq
                End If
            Else
                txtNumDoc.Text = ""
            End If
            txtNumDoc.Visible = True
        
        'RIRO20140530 ERS017 ***
        ElseIf CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoVoucher Then
        
            Dim oformV As frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIf As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
            Dim bCondicion As Boolean
                        
            Set oformV = New frmCapRegVouDepBus
            lnTipMot = 11 ' Pago Credito Judicial
                
            oformV.iniciarFormularioDeposito CInt(Mid(AXCodCta.NroCuenta, 9, 1)), lnTipMot, sGlosa, sIf, sVaucher, nMontoVoucher, sPersCod, nMovNroRVD, nMovNroRVDPen, sNombre, sDireccion, sDocumento, AXCodCta.NroCuenta
            If nMontoVoucher < AXMontoPago.value + CDbl(LblITF.Caption) Then
                If nMovNroRVD > 0 Then
                    MsgBox "No es posible realizar el pago con el Voucher porque no cuenta con saldo suficiente para realizar la operación", _
                    vbExclamation, "Aviso"
                End If
                nMovNroRVD = 0
                nMovNroRVDPen = -1
                nMontoVoucher = 0
                sNombre = ""
                sDireccion = ""
                If cboModalidad.Enabled Then cboModalidad.SetFocus
                Exit Sub
            Else
                If Len(sVaucher) = 0 Then
                    txtNumDoc.Text = sVaucher
                Else
                    txtNumDoc.Text = Trim(Mid(sVaucher, 1, Len(sVaucher) - 10))
                End If
                txtNumDoc.Visible = True
            End If
        'END RIRO *****
        Else
            txtNumDoc.Visible = False
            'MatDatos(0) = ""
        End If
    End If
    Exit Sub
ErrCboModalidad:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub cboModalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If AXMontoPago.Enabled = True Then '**Juez 20120516 Agegado x observaciones en Pruebas
            AXMontoPago.SetFocus
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
Dim lrCreditos As ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If Not loPers Is Nothing Then
        lsPersCod = loPers.sPersCod
        sPersCod = loPers.sPersCod
        lsPersNombre = loPers.sPersNombre
    Else
        Exit Sub
    End If
Set loPers = Nothing

' Selecciona Estados
'lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast
lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast & "," & gColocEstTransferido 'FRHU 20150415 ERS022-2015

If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
        Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
    Set loPersCredito = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing
'ventana = 1
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    Call HabilitaControles(False, True, True)
    AXCodCta.SetFocusAge
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim loGrabar As COMNColocRec.NCOMColRecCredito 'NColRecCredito
Dim loImprime As COMNColocRec.NCOMColRecImpre 'NColRecImpre
Dim loPrevio As previo.clsprevio
Dim lnDocTpo As Integer
Dim lsNroDoc As String

Dim lsMovNro As String

Dim lsCadImprimir As String
Dim lsNombreCliente As String
Dim lsOpeCod As String
Dim lsOpeITFChequeEfec As Integer

Dim loLavDinero As frmMovLavDinero
Set loLavDinero = New frmMovLavDinero

'*****BRGO 20110914 ***********
Dim loMov As COMDMov.DCOMMov
Set loMov = New COMDMov.DCOMMov
'*** End BRGO *****************
Dim fnCondicion As Integer 'WIOR 20130301
Dim regPersonaRealizaPago As Boolean 'WIOR 20130301

If Len(Trim(lblMetLiquid)) <> 4 Then
    MsgBox "Metodo de Liquidación no válido o no definido", vbInformation, "Aviso"
    Exit Sub
End If

lsNombreCliente = Mid(Me.lblCliente.Caption, 1, 30)
If cboModalidad = "" Then
    MsgBox "Seleccione modalidad de Pago", vbInformation, "aviso"
    Me.cboModalidad.SetFocus
    Exit Sub
End If
If Left(fnVarOpeCod, 3) <> "136" Then
    If CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoCheque Then
        fnVarOpeCod = gColRecOpePagJudSDChq
        If txtNumDoc = "" Then
            MsgBox "Numero de Cheque no Válido", vbInformation, "aviso"
            Exit Sub
        End If
        lnDocTpo = TpoDocCheque
        lsNroDoc = Trim(txtNumDoc)
        'lsOpeITFChequeEfec = "990107"
    'RIRO20140610 ERS017 ****************
    ElseIf CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoVoucher Then
        lnDocTpo = 0
        lsNroDoc = Trim(txtNumDoc)
        fnVarOpeCod = gColRecOpePagJudSDVou
        If nMovNroRVD = 0 Then
            MsgBox "Debe seleccionar un voucher para continuar con la operacion", vbInformation, "Aviso"
            Exit Sub
        End If
    'END RIRO ***************************
    Else
        lnDocTpo = 0
        lsNroDoc = ""
        fnVarOpeCod = gColRecOpePagJudSDEfe
    End If
End If
lsOpeCod = AsignaCodigoOperacionPago(fnVarOpeCod, fsCondicion, fsDemanda)

If CCur(AXMontoPago.Text) > CCur(lblTotalAct) Then
   MsgBox "Monto a Pagar no debe Exceder el Total de Deuda", vbInformation, "Aviso"
   Me.AXMontoPago.SetFocus
   Exit Sub
End If

'By Capi 15042008
If CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoCheque Then
    If Not ValidaSeleccionCheque Then Exit Sub
    'If CCur(TxtTotalAPagar.Text) > MatDatos(0) Then
    If CCur(TxtTotalAPagar.Text) > oDocRec.fnMonto Then
        MsgBox "Disponible del cheque no cubre el Monto a Pagar", vbInformation, "Aviso"
        If AXMontoPago.Visible And AXMontoPago.Enabled Then Me.AXMontoPago.SetFocus
        Exit Sub
    End If

End If

'********* VERIFICAR VISTO AVMM - 13-12-2006 **********************
'Dim loVisto As COMDColocRec.DCOMColRecCredito
'Set loVisto = New COMDColocRec.DCOMColRecCredito
'    '2=Negociacion
'    If loVisto.bVerificarVisto(AXCodCta.NroCuenta, 2) = False Then
'        If loVisto.bVerificarVisto(AXCodCta.NroCuenta, 3) = False Then
'            MsgBox "No existe Visto para realizar Negociación", vbInformation, "Aviso"
'            Exit Sub
'        Else
'            MsgBox "El Crédito posee Visto para Cancelación", vbInformation, "Aviso"
'            Exit Sub
'        End If
'    End If
'Set loVisto = Nothing
'********************************************************************

   'WIOR 20121009 Clientes Observados **************************************
            Dim oDPersona As COMDPersona.DCOMPersona
            Dim rsPersona As ADODB.Recordset
            Set oDPersona = New COMDPersona.DCOMPersona
            Set rsPersona = oDPersona.ObtenerUltimaVisita(Trim(sPersCod))
         
            If rsPersona.RecordCount > 0 Then
                If Not (rsPersona.EOF And rsPersona.BOF) Then
                    If Trim(rsPersona!sUsual) = "3" Then
                        MsgBox "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                        Call frmPersona.Inicio(Trim(sPersCod), PersonaActualiza)
                    End If
                End If
            End If
    'WIOR FIN ***************************************************************

If MsgBox(" Desea Grabar Pago de Credito en Recuperaciones ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    
    ' EJRS **** RECALCULAR LA DISTRIBUCIÓN DEL MONTO PAGADO
    Call CalculaDistribucionPago
  '======= Lavado de Dinero
    Dim nMontoLavDinero As Double, nTC As Double
    Dim nmoneda As Integer
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim sPersLavDinero As String
    nmoneda = CDbl(Mid(AXCodCta.NroCuenta, 9, 1))
    sPersLavDinero = ""
    Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
    If clsLav.EsOperacionEfectivo(CStr(fnVarOpeCod)) Then
        If Not EsExoneradaLavadoDinero() Then

            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            
            If Mid(AXCodCta.NroCuenta, 9, 1) = gMonedaNacional Then
                Dim clsTC As COMDConstSistema.NCOMTipoCambio
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If CDbl(AXMontoPago.Text) >= Round(nMontoLavDinero * nTC, 2) Then
                'By Capi 1402208
                    Call IniciaLavDinero(loLavDinero)
                    'ALPA 20081009***************************************************************************
                    'sPersLavDinero = loLavDinero.Inicia(, , , , False, True, CDbl(AXMontoPago.Text), AXCodCta.NroCuenta, Mid(Me.Caption, 15), True, "", , , , , nmoneda)
                    sPersLavDinero = loLavDinero.Inicia(, , , , False, True, CDbl(AXMontoPago.Text), AXCodCta.NroCuenta, Mid(Me.Caption, 15), True, "", , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                    '****************************************************************************************
                    If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                'End

                
            End If
        End If
    End If
    Set clsLav = Nothing
  '=============================
    'WIOR 20130301 Personas Sujetas a Procedimiento Reforzado *************************************
    If loLavDinero.OrdPersLavDinero = "Exit" Then
        Dim oPersonaSPR As UPersona_Cli
        Dim oPersonaU As COMDPersona.UCOMPersona
        Dim nTipoConBN As Integer
        Dim sConPersona As String
        Dim pbClienteReforzado As Boolean
        Dim rsAgeParam As Recordset
        Dim objCred As COMNCredito.NCOMCredito
        Dim lnMonto As Double, lnTC As Double
        Dim objTC As COMDConstSistema.NCOMTipoCambio
        Dim sOpeCod As String 'RIRO20140620 ERS017
        
        Set oPersonaU = New COMDPersona.UCOMPersona
        Set oPersonaSPR = New UPersona_Cli
        
        regPersonaRealizaPago = False
        pbClienteReforzado = False
        fnCondicion = 0

        oPersonaSPR.RecuperaPersona sPersCod
                            
        If oPersonaSPR.Personeria = 1 Then
            If oPersonaSPR.Nacionalidad <> "04028" Then
                sConPersona = "Extranjera"
                fnCondicion = 1
                pbClienteReforzado = True
            ElseIf oPersonaSPR.Residencia <> 1 Then
                sConPersona = "No Residente"
                fnCondicion = 2
                pbClienteReforzado = True
            ElseIf oPersonaSPR.RPeps = 1 Then
                sConPersona = "PEPS"
                fnCondicion = 4
                pbClienteReforzado = True
            ElseIf oPersonaU.ValidaEnListaNegativaCondicion(IIf(Trim(oPersonaSPR.ObtenerDNI) = "", oPersonaSPR.ObtenerNumeroDoc(0), oPersonaSPR.ObtenerDNI), oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                If nTipoConBN = 1 Or nTipoConBN = 3 Then
                    sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                    fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                    pbClienteReforzado = True
                End If
            End If
        Else
            If oPersonaU.ValidaEnListaNegativaCondicion(oPersonaSPR.ObtenerDNI, oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                If nTipoConBN = 1 Or nTipoConBN = 3 Then
                    sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                    fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                    pbClienteReforzado = True
                End If
            End If
        End If
        
        If pbClienteReforzado Then
            MsgBox "El Cliente: " & Trim(lblCliente.Caption) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
            frmPersRealizaOpeGeneral.Inicia fsVarOpeDesc & " (Persona " & sConPersona & ")", fnVarOpeCod
            regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
            
            If Not regPersonaRealizaPago Then
                MsgBox "Se va a proceder a Anular el Pago de la Cuota", vbInformation, "Aviso"
                Exit Sub
            End If
        Else
            fnCondicion = 0
            lnMonto = CDbl(AXMontoPago.Text)
            pbClienteReforzado = False
            
            Set objTC = New COMDConstSistema.NCOMTipoCambio
            lnTC = objTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
            Set objTC = Nothing
        
        
            Set objCred = New COMNCredito.NCOMCredito
            Set rsAgeParam = objCred.obtieneCredPagoCuotasAgeParam(gsCodAge)
            Set objCred = Nothing
            
            If Mid(AXCodCta.NroCuenta, 9, 1) = 2 Then
                lnMonto = Round(lnMonto * lnTC, 2)
            End If
        
            If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                If lnMonto >= rsAgeParam!nMontoMin And lnMonto <= rsAgeParam!nMontoMax Then
                    frmPersRealizaOpeGeneral.Inicia fsVarOpeDesc, fnVarOpeCod
                    regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
                    If Not regPersonaRealizaPago Then
                        MsgBox "Se va a proceder a Anular el Pago de la Cuota", vbInformation, "Aviso"
                        Exit Sub
                    End If
                End If
            End If
            
        End If
    End If
    'WIOR FIN ***************************************************************
    cmdGrabar.Enabled = False
    If oDocRec Is Nothing Then Set oDocRec = New UDocRec 'EJVG20140408
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabar = New COMNColocRec.NCOMColRecCredito
            'ALPA 20081009***********************************************************
            'Call loGrabar.nPagoCreditoRecup(AXCodCta.NroCuenta, lsFechaHoraGrab, lsOpeCod, _
                 lsMovNro, CCur(Me.AXMontoPago.Text), Me.lblMetLiquid.Caption, fnNewSaldoCap, fnNewSaldoIntComp, _
                 fnNewSaldoIntMorat, fnNewSaldoGasto, fnNewSaldoIntCompGen, fnComisionAbog, fnNroUltGastoCta, _
                 fnGastoAdminAdicional, fnCapPag, fnIntCompPag, fnIntMoratPag, fnGastoPag, fmMatGastos, _
                 fnNroCalend, fnEstadoNew, False, sPersLavDinero, CDbl(LblITF.Caption), lnDocTpo, lsNroDoc, IIf(CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoCheque, "990107", "990105"), fsVarPersCodCMAC, fnNewSaldoIntMoraGen, , , , , , gnMovNro) 'DAOR 20070809, se agregó el campo fnNewSaldoIntMoraGen
        
            'RIRO20140620 ERS017 - COMENTADO
            'Call loGrabar.nPagoCreditoRecup(AXCodCta.NroCuenta, lsFechaHoraGrab, lsOpeCod, _
            '     lsMovNro, CCur(Me.AXMontoPago.Text), Me.lblMetLiquid.Caption, fnNewSaldoCap, fnNewSaldoIntComp, _
            '     fnNewSaldoIntMorat, fnNewSaldoGasto, fnNewSaldoIntCompGen, fnComisionAbog, fnNroUltGastoCta, _
            '     fnGastoAdminAdicional, fnCapPag, fnIntCompPag, fnIntMoratPag, fnGastoPag, fmMatGastos, _
            '     fnNroCalend, fnEstadoNew, False, sPersLavDinero, CDbl(lblITF.Caption), oDocRec.fnTpoDoc, oDocRec.fsNroDoc, IIf(CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoCheque, "990107", "990105"), fsVarPersCodCMAC, fnNewSaldoIntMoraGen, , , , , , gnMovNro, oDocRec.fsPersCod, oDocRec.fsIFTpo, oDocRec.fsIFCta)  'EJVG20140408
                             
            ' RIRO20140620 ERS017 *********
            If CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoCheque Then
                sOpeCod = "990107"
            ElseIf CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoVoucher Then
                sOpeCod = "990121"
            Else
                sOpeCod = "990105"
            End If
            Call loGrabar.nPagoCreditoRecup(AXCodCta.NroCuenta, lsFechaHoraGrab, lsOpeCod, _
                 lsMovNro, CCur(Me.AXMontoPago.Text), Me.lblMetLiquid.Caption, fnNewSaldoCap, fnNewSaldoIntComp, _
                 fnNewSaldoIntMorat, fnNewSaldoGasto, fnNewSaldoIntCompGen, fnComisionAbog, fnNroUltGastoCta, _
                 fnGastoAdminAdicional, fnCapPag, fnIntCompPag, fnIntMoratPag, fnGastoPag, fmMatGastos, _
                 fnNroCalend, fnEstadoNew, False, sPersLavDinero, CDbl(LblITF.Caption), oDocRec.fnTpoDoc, oDocRec.fsNroDoc, sOpeCod, fsVarPersCodCMAC, fnNewSaldoIntMoraGen, , , , , , gnMovNro, oDocRec.fsPersCod, oDocRec.fsIFTpo, oDocRec.fsIFCta, nMovNroRVD, nMovNroRVDPen)   'EJVG20140408
            'END RIRO *********************
                 
        If gnMovNro = 0 Then
             MsgBox "La operación no se realizó, favor intente nuevamente", vbInformation, "Aviso"
             Exit Sub
        End If
            '***************RECO 2013-07-22*******************
             loGrabar.nActualizarMetLiquidPagoJud (AXCodCta.NroCuenta)
            '***************************************************
        'ALPA20131001***********************

        '***********************************
        Set loGrabar = Nothing
        'Actualiza Negociacion
        '*****BRGO 20110914 *****************************************************
        If gbITFAplica = True And CCur(LblITF.Caption) > 0 Then
           Call loMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(LblITF.Caption) + nRedondeoITF, CCur(LblITF.Caption)) 'BRGO 20110914
        End If
        Set loMov = Nothing
        '*** End BRGO *****************
        If Len(lblNroNeg) > 0 Then
            Call ActualizaNegociacion(AXCodCta.NroCuenta, lsOpeCod)
        End If
        
        'Impresión
        Set loImprime = New COMNColocRec.NCOMColRecImpre
            'If gsCodCMAC = "102" Then
            '    lsCadImprimir = loImprime.nPrintReciboPagoCredRecupLima(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, _
            '    lsNombreCliente, CCur(Me.AXMontoPago.Text), gsCodUser, lblDistrib(0), lblDistrib(1), _
            '     lblDistrib(2), lblDistrib(3), " ")
            'Else
                lsCadImprimir = loImprime.nPrintReciboPagoCredRecup(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, _
                lsNombreCliente, CCur(Me.AXMontoPago.Text), gsCodUser, " ", CDbl(LblITF.Caption), gImpresora, gbImpTMU, lsOpeCod)
                'WIOR 20150615 AGREGO lsOpeCod
            'End If
        Set loImprime = Nothing
        Set loPrevio = New previo.clsprevio
            loPrevio.PrintSpool sLpt, lsCadImprimir, True, 22
            
            Do While True
                If MsgBox("Reimprimir Recibo de Pago de Credito en Recuperaciones ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    loPrevio.PrintSpool sLpt, lsCadImprimir, True, 22
                Else
                    Set loPrevio = Nothing
                    Exit Do
                End If
            Loop
        'WIOR 20130301 ************************************************************
        If regPersonaRealizaPago And gnMovNro > 0 Then
            frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(AXCodCta.NroCuenta), fnCondicion
            regPersonaRealizaPago = False
        End If
        'WIOR FIN *****************************************************************
        
        Limpiar
        
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
        'ALPA 20081010
        If gnMovNro > 0 Then
            'Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, , , sPersLavDinero, , , , gnTipoREU, gnMontoAcumulado, gsOrigen)
             Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110225
            
        End If
   
    
        
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
    Dim bRetSinTarjeta As Boolean
    Dim sCuenta As String
    Dim nProducto As Producto
    nProducto = gColComercAgro
    sCuenta = frmValTarCodAnt.Inicia(nProducto, bRetSinTarjeta)
    If sCuenta <> "" Then
        AXCodCta.NroCuenta = sCuenta
        AXCodCta.SetFocusCuenta
    End If
End If
End Sub

Private Sub Form_Load()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
    CentraForm Me
'    ventana = 0
End Sub

Private Sub FraComandos_DblClick()
    If Me.Height = 6210 Then
        Me.Height = 4700
    Else
        Me.Height = 6210
    End If
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub

Private Function AsignaCodigoOperacionPago(ByVal pnOperacion As String, ByVal psCondicion As String, Optional ByVal pnDemanda As String) As String
Dim lsOpe As String

Select Case pnOperacion
    Case gColRecOpePagJudSDEfe  ' Pago en Efectivo
        If psCondicion = "J" Then
            If pnDemanda = "S" Then
                lsOpe = gColRecOpePagJudCDEfe
            ElseIf pnDemanda = "N" Then
                lsOpe = gColRecOpePagJudCDEfe
            End If
        ElseIf psCondicion = "A" Then
            lsOpe = gColRecOpePagCastEfe
        End If
    
    Case gColRecOpePagJudSDChq
        If psCondicion = "J" Then
            If pnDemanda = "S" Then
                lsOpe = gColRecOpePagJudCDChq
            ElseIf pnDemanda = "N" Then
                lsOpe = gColRecOpePagJudCDChq
            End If
        ElseIf psCondicion = "A" Then
            lsOpe = gColRecOpePagCastChq
        End If
    'RIRO20140530 ER017 ************
    Case gColRecOpePagJudSDVou
        If psCondicion = "J" Then
            If pnDemanda = "S" Then
                lsOpe = gColRecOpePagJudCDVou
            ElseIf pnDemanda = "N" Then
                lsOpe = gColRecOpePagJudCDVou
            End If
        ElseIf psCondicion = "A" Then
            lsOpe = gColRecOpePagCastVou
        End If
    'END RIRO **********************
    Case gColRecOpePagJudSDEnOtCjEfe   ' En otra CMAC
        If psCondicion = "J" Then
            If pnDemanda = "S" Then
                lsOpe = gColRecOpePagJudCDEnOtCjEfe
            ElseIf pnDemanda = "N" Then
                lsOpe = gColRecOpePagJudCDEnOtCjEfe
            End If
        ElseIf psCondicion = "A" Then
            lsOpe = gColRecOpePagJudCastEnOtCjEfe
        End If
End Select
AsignaCodigoOperacionPago = lsOpe

End Function

Private Function EsExoneradaLavadoDinero() As Boolean
Dim bExito As Boolean
Dim clsExo As COMNCaptaServicios.NCOMCaptaServicios
bExito = True

    Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
    
    If Not clsExo.EsPersonaExoneradaLavadoDinero(sPersCod) Then bExito = False

    Set clsExo = Nothing
    EsExoneradaLavadoDinero = bExito
    
End Function

Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)
Dim i As Long
 
Dim oPersona As COMDCaptaGenerales.DCOMCaptaGenerales
Dim oCta As COMDCredito.DCOMCredito
Dim rsPers As New ADODB.Recordset
Dim sNombre As String
Dim sDireccion As String
Dim sDocId As String
Dim nMonto As Double
Set oCta = New COMDCredito.DCOMCredito

sPersCod = oCta.RecuperaTitularCredito(AXCodCta.NroCuenta)
Set oCta = Nothing

Set oPersona = New COMDCaptaGenerales.DCOMCaptaGenerales

Set rsPers = oPersona.GetDatosPersona(sPersCod)
If rsPers.BOF Then
Else
    poLavDinero.TitPersLavDinero = sPersCod
    poLavDinero.TitPersLavDineroNom = rsPers!Nombre
    poLavDinero.TitPersLavDineroDir = rsPers!Direccion
    poLavDinero.TitPersLavDineroDoc = rsPers!id & " " & rsPers![ID N°]
End If
rsPers.Close
Set rsPers = Nothing

 nMonto = CDbl(AXMontoPago.Text)
'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, AXCodCta.NroCuenta, CStr(fnVarOpeCod), False, "COLOCACIONES")
'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, True, nMonto, AXCodCta.NroCuenta, CStr(fnVarOpeCod), , "COLOCACIONES")

End Sub

Private Sub CargaParametros()

Dim loParam As COMDConstSistema.NCOMConstSistema
Set loParam = New COMDConstSistema.NCOMConstSistema
fnTipoCalcIntComp = loParam.LeeConstSistema(151)
fnTipoCalcIntMora = loParam.LeeConstSistema(152)
fnFormaCalcIntComp = loParam.LeeConstSistema(202) ' CMACICA
fnFormaCalcIntMora = loParam.LeeConstSistema(203) ' CMACICA

'Call CargaComboConstante(gColocTipoPago, CboModalidad) RIRO20140530 ER017 Comentado

Dim oConstante As New COMDConstSistema.DCOMGeneral
Dim rsConstante As New ADODB.Recordset

Set rsConstante = oConstante.GetConstante(gColocTipoPago, , "'[125]'")
CargaCombo cboModalidad, rsConstante

Set oConstante = Nothing
Set rsConstante = Nothing
Set loParam = Nothing

End Sub

Private Sub TxtTotalAPagar_Change()
'    If Trim(TxtTotalAPagar.Text) = "" Then
'        TxtTotalAPagar.Text = "0.00"
'    End If
'    AXMontoPago.Text = Format(fgITFCalculaImpuestoIncluido(CDbl(Me.TxtTotalAPagar.Text)), "#0.00")
'    LblITF.Caption = Format(CDbl(Me.TxtTotalAPagar.Text) - CDbl(AXMontoPago.Text), "#0.00")
End Sub

Private Sub TxtTotalAPagar_GotFocus()
    fEnfoque TxtTotalAPagar
End Sub

Private Sub TxtTotalAPagar_KeyPress(KeyAscii As Integer)
    Call AXMontoPago_KeyPress(KeyAscii)
End Sub

Private Sub CargaPlanPagos(ByVal psCodCta As String, ByVal psNroNeg As String)
Dim RegCuotas  As New ADODB.Recordset
Dim i, k As Integer

Dim lrDatCuotas As New ADODB.Recordset
Dim loValCred As COMDColocRec.DCOMColRecCredito

Dim lsMensaje As String

 fnContCuotas = 0: fnDiasAtraso = 0
'Realiza Carga de Caledario

    Set loValCred = New COMDColocRec.DCOMColRecCredito
        Set lrDatCuotas = loValCred.dObtieneDatosNegociaCuotas(psCodCta, psNroNeg, lsMensaje)
        If Trim(lsMensaje) <> "" Then
            MsgBox lsMensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    Set loValCred = Nothing
    If Not lrDatCuotas Is Nothing Then
        If Not lrDatCuotas.BOF And Not lrDatCuotas.EOF Then
            Do While Not lrDatCuotas.EOF
                If lrDatCuotas!cEstado = "P" Then
                    fnContCuotas = fnContCuotas + 1
                    ReDim Preserve MatPagos(fnContCuotas)
                    MatPagos(fnContCuotas - 1).NumCuota = lrDatCuotas!nNroCuota
                    MatPagos(fnContCuotas - 1).FecVenc = lrDatCuotas!dFecVenc
                    MatPagos(fnContCuotas - 1).Monto = Format(lrDatCuotas!nMonto - lrDatCuotas!nMontoPag, "#0.00")
                    MatPagos(fnContCuotas - 1).MontoPag = 0
                    MatPagos(fnContCuotas - 1).Estado = lrDatCuotas!cEstado
                    MatPagos(fnContCuotas - 1).Modificado = False
                End If
                lrDatCuotas.MoveNext
            Loop
            lrDatCuotas.Close
            fnDiasAtraso = gdFecSis - MatPagos(0).FecVenc
            'Salvar Matriz de Pagos Original
            MatPagosTempo = MatPagos
        End If
    End If
    Set lrDatCuotas = Nothing
End Sub

Sub CubrirCalendario(ByVal pnMontoPagado As Double)
Dim i As Integer
Dim lnMontoSaldo As Double
lnMontoSaldo = pnMontoPagado
For i = 0 To fnContCuotas - 1
    If MatPagos(i).Estado = "P" And lnMontoSaldo > 0 Then
        MatPagos(i).Modificado = True
        If CDbl(Format(lnMontoSaldo, "#0.00")) >= MatPagos(i).Monto Then
            MatPagos(i).MontoPag = MatPagos(i).Monto
            MatPagos(i).Estado = "G"
            lnMontoSaldo = lnMontoSaldo - MatPagos(i).Monto
        Else
            MatPagos(i).MontoPag = lnMontoSaldo
            lnMontoSaldo = 0#
        End If
    End If
Next i
End Sub

Function MontoPendienteNeg() As Double
Dim i As Integer
Dim lnMontoSaldo As Double
lnMontoSaldo = 0
For i = 0 To fnContCuotas - 1
    If MatPagos(i).Estado = "P" And CDate(MatPagos(i).FecVenc) <= CDate(gdFecSis) Then
            lnMontoSaldo = lnMontoSaldo + MatPagos(i).Monto - MatPagos(i).MontoPag
    End If
Next i
If fnContCuotas > 0 Then
    If lnMontoSaldo = 0 Then
            lnMontoSaldo = lnMontoSaldo + MatPagos(0).Monto - MatPagos(0).MontoPag
    End If
End If
MontoPendienteNeg = lnMontoSaldo
End Function

Sub ActualizaNegociacion(ByVal psCtaCod As String, ByVal psOpeCod As String)
Dim i As Integer
Dim SQL1 As String
Dim rs As New ADODB.Recordset
Dim loRecNeg As COMDColocRec.DCOMColRecNegociacion

With rs
    'Crear RecordSet
    .Fields.Append "nNumCuota", adInteger
    .Fields.Append "dFecVenc", adDate
    .Fields.Append "nMonto", adDecimal
    .Fields.Append "nMontoPag", adDecimal
    .Fields.Append "cEstado", adVarChar, 150
    .Fields.Append "bModificado", adBoolean
    .Open
    'Llenar Recordset
     For i = 0 To fnContCuotas - 1
        .AddNew
        .Fields("nNumCuota") = MatPagos(i).NumCuota
        .Fields("dFecVenc") = MatPagos(i).FecVenc
        .Fields("nMonto") = MatPagos(i).Monto
        .Fields("nMontoPag") = MatPagos(i).MontoPag
        .Fields("cEstado") = MatPagos(i).Estado
        .Fields("bModificado") = MatPagos(i).Modificado
    Next i
    
End With


    'Actualiza PlanPagos
    Set loRecNeg = New COMDColocRec.DCOMColRecNegociacion
        loRecNeg.ActualizarNegocPlanPagos lsFechaHoraGrab, psCtaCod, Me.lblNroNeg.Caption, rs
    Set loRecNeg = Nothing

    'Inserta Kardex
    
    Set loRecNeg = New COMDColocRec.DCOMColRecNegociacion
        loRecNeg.InsertaNegocPlanPagoskardex psCtaCod, lblNroNeg.Caption, psOpeCod, lsFechaHoraGrab, Me.AXMontoPago.Text, fnDiasAtraso, txtNumDoc.Text, gsCodAge, gsCodUser
    Set loRecNeg = Nothing
    
  
End Sub

Private Sub DistribuyePagoCominAbogado()
Dim lsPrio1 As String, lsPrio2 As String, lsPrio3 As String, lsPrio4 As String

lsPrio1 = Mid(Me.lblMetLiquid, 1, 1)
lsPrio2 = Mid(Me.lblMetLiquid, 2, 1)
lsPrio3 = Mid(Me.lblMetLiquid, 3, 1)
lsPrio4 = Mid(Me.lblMetLiquid, 4, 1)

fnCapPag = 0: fnIntCompPag = 0: fnIntMoratPag = 0:  fnGastoPag = 0

'fnMontoPagar = Format(Me.AXMontoPago.Text, "#0.00")

If fnMontoPagar > 0 Then
    Select Case lsPrio1
        Case "G": Call CubrirGastos
        Case "M": Call CubrirMora
        Case "I": Call CubrirInteres
        Case "C": Call CubrirCapital
    End Select
End If

If fnMontoPagar > 0 Then
    Select Case lsPrio2
        Case "G": Call CubrirGastos
        Case "M": Call CubrirMora
        Case "I": Call CubrirInteres
        Case "C": Call CubrirCapital
    End Select
End If

If fnMontoPagar > 0 Then
    Select Case lsPrio3
        Case "G":  Call CubrirGastos
        Case "M":  Call CubrirMora
        Case "I":  Call CubrirInteres
        Case "C":  Call CubrirCapital
    End Select
End If

If fnMontoPagar > 0 Then
    Select Case lsPrio4
        Case "G": Call CubrirGastos
        Case "M": Call CubrirMora
        Case "I": Call CubrirInteres
        Case "C": Call CubrirCapital
    End Select
End If
End Sub

'**DAOR 20070424, Procedimiento que establece los montos CIMG distribuidos de forma manual
Sub EstablecerCIMGPersonalizado()
Dim lnGastoDistrib As Currency
Dim i As Integer
    fnCapPag = fnCapDist
    fnIntCompPag = fnIntCompDist
    fnIntMoratPag = fnIntMoratDist
    fnComisionAbog = fnComisionAbogDist
    fnGastoPag = fnGastoDist
    fnMontoPagar = 0
    lnGastoDistrib = fnGastoPag
    
    Call ReCalCulaGasto(AXCodCta.NroCuenta)
    
    For i = 0 To UBound(fmMatGastos) - 1
        If CInt(fmMatGastos(i, 4)) = gColRecGastoEstPendiente And lnGastoDistrib > 0 _
           And (CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))) > 0 Then
            If lnGastoDistrib >= (CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))) Then
                lnGastoDistrib = lnGastoDistrib - (CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3)))
                'Actualiza el monto Pagado
                'fmMatGastos(i, 3) = CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))
                'Actualiza el estado del gasto
                fmMatGastos(i, 4) = gColRecGastoEstPagado
                fmMatGastos(i, 5) = "S" ' Si se ha modificado
                fmMatGastos(i, 6) = CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))
            Else
                'Actualiza el monto pagado
                fmMatGastos(i, 3) = Format(CDbl(fmMatGastos(i, 3)) + lnGastoDistrib, "#0.00")
                fmMatGastos(i, 4) = gColRecGastoEstPendiente
                fmMatGastos(i, 5) = "S" ' Si se ha modificado
                fmMatGastos(i, 6) = Format(lnGastoDistrib, "#0.00")
                lnGastoDistrib = 0
            End If
        End If
    Next i
End Sub
'EJVG20140303 ***
Private Function ValidaSeleccionCheque() As Boolean
    ValidaSeleccionCheque = True
    If oDocRec Is Nothing Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
    If Len(Trim(oDocRec.fsNroDoc)) = 0 Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
End Function
Private Sub Form_Unload(Cancel As Integer)
    Set oDocRec = Nothing
End Sub
'END EJVG *******
'FRHU 20150415 ERS022-2015
Private Function VerificarSiEsUnCreditoTransferido(ByVal psCtaCod As String) As Boolean
    Dim oCredito As COMDCredito.DCOMCredito
    
    Set oCredito = New COMDCredito.DCOMCredito
    VerificarSiEsUnCreditoTransferido = oCredito.VerificaSiEsCreditoTransferido(psCtaCod)
    Set oCredito = Nothing
End Function
'FIN FRHU 20150415
