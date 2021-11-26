VERSION 5.00
Begin VB.Form frmGestionSiniestroLiquidar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidación de Créditos con el Seguro Desgravamen"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   Icon            =   "frmGestionSiniestroLiquidar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPago 
      Caption         =   "Datos del Pago"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   40
      Top             =   6600
      Width           =   6735
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   360
         Width           =   1800
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   285
         Left            =   1800
         TabIndex        =   42
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtMontoPagar 
         Height          =   285
         Left            =   1800
         TabIndex        =   43
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtDevolucion 
         Height          =   285
         Left            =   4680
         TabIndex        =   44
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Devolución:"
         Height          =   195
         Left            =   3720
         TabIndex        =   48
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Monto a Pagar:"
         Height          =   195
         Left            =   360
         TabIndex        =   47
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Monto:"
         Height          =   195
         Left            =   360
         TabIndex        =   46
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago:"
         Height          =   195
         Left            =   360
         TabIndex        =   45
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.Frame fraDistribucion 
      Caption         =   "Distribución de Montos"
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
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   6735
      Begin SICMACT.EditMoney txtCapitalFecha 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtIntCompFecha 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtIntMoraFecha 
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtGastosFecha 
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtITFFecha 
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtCapitalPagar 
         Height          =   285
         Left            =   4320
         TabIndex        =   12
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtIntCompPagar 
         Height          =   285
         Left            =   4320
         TabIndex        =   13
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtIntMoraPagar 
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtGastosPagar 
         Height          =   285
         Left            =   4320
         TabIndex        =   15
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtITFPagar 
         Height          =   285
         Left            =   4320
         TabIndex        =   16
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtTotalFecha 
         Height          =   285
         Left            =   2280
         TabIndex        =   49
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtTotalPagar 
         Height          =   285
         Left            =   4320
         TabIndex        =   50
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtAjuste 
         Height          =   285
         Left            =   4320
         TabIndex        =   53
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.Label Label1 
         Caption         =   "Ajuste:"
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   360
         TabIndex        =   51
         Top             =   2640
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Deuda a la Fecha"
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
         Left            =   2400
         TabIndex        =   23
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "A Pagar"
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
         Left            =   4920
         TabIndex        =   22
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Capital:"
         Height          =   195
         Left            =   360
         TabIndex        =   21
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Interés Compensatorio:"
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   1620
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "ITF:"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   2280
         Width           =   285
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Gastos:"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Interés Moratorio:"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   1200
         Width           =   1230
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos del Crédito"
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
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   6735
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Cuotas:"
         Height          =   195
         Left            =   3720
         TabIndex        =   39
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Crédito:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Monto del Crédito:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Capital:"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Método Liq.: "
         Height          =   195
         Left            =   3720
         TabIndex        =   34
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label lblCuotas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4800
         TabIndex        =   33
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblMetodLiq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4800
         TabIndex        =   32
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   31
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblMontoCred 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   30
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   29
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblTpoCred 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   28
         Top             =   1080
         Width           =   5175
      End
      Begin VB.Label lblTpoProd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   720
         Width           =   5175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Producto:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label lblCodTitular 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   8280
      Width           =   1215
   End
   Begin SICMACT.ActXCodCta ActXCtaCred 
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3585
      _ExtentX        =   6535
      _ExtentY        =   688
      Texto           =   "Credito :"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
      CMAC            =   "112"
   End
End
Attribute VB_Name = "frmGestionSiniestroLiquidar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'** Nombre      : frmGestionSiniestroLiquidar
'** Descripción : Formulario para realizar la operacion de liquidacion respectiva del credito
'**               Creado según TI-ERS003-2016
'** Creación    : WIOR, 20160226 09:00:00 AM
'*********************************************************************************************

Option Explicit
Private fnIdSiniestro As Long
Private fnIdRegAtencion As Long
Private oDocRec As UDocRec
Private nMovNroRVD() As Variant 'Mov del voucher y de la pendiente
Private nMontoVoucher As Currency

Private ArrDistMontos As Variant
Private ArrDocRec As Variant
Private fsNroDoc As String
Dim vPrevio As previo.clsprevio
'MARG ERS073****************
Dim lcCodPersTitu As String
'END MARG********************

Private Sub ActXCtaCred_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CargaDatos (Trim(ActXCtaCred.NroCuenta))
End If
End Sub


Private Sub cboFormaPago_Click()
fsNroDoc = ""
txtDevolucion.Text = "0.00"
If cboFormaPago.ListIndex <> -1 Then

    If CInt(Trim(Right(cboFormaPago.Text, 10))) = gColocTipoPagoCheque Then

        Dim oform As New frmChequeBusqueda
        Set oDocRec = oform.iniciarBusqueda(Val(Mid(ActXCtaCred.NroCuenta, 9, 1)), TipoOperacionCheque.CRED_LiqCreditosSegDes, ActXCtaCred.NroCuenta)
        Set oform = Nothing
        txtMonto = Format(oDocRec.fnMonto, "#,##0.00")
        fsNroDoc = oDocRec.fsNroDoc
        ReDim nMovNroRVD(6)
    ElseIf CInt(Trim(Right(cboFormaPago.Text, 10))) = gColocTipoPagoVoucher Then
 
        Dim oformVou As New frmCapRegVouDepBus
        Dim lnTipMot As Integer
        Dim sGlosa As String
        Dim sIF As String
        Dim sVaucher As String
        Dim sPersCod As String
        Dim sNombre As String
        Dim sDireccion As String
        Dim sDocumento As String
        Dim nMovNro As Long, nMovNroPend As Long
                    
        lnTipMot = 13 'Liq Credito seg des.
        oformVou.iniciarFormularioDeposito CInt(Mid(ActXCtaCred.NroCuenta, 9, 1)), lnTipMot, sGlosa, sIF, sVaucher, nMontoVoucher, sPersCod, nMovNro, nMovNroPend, sNombre, sDireccion, sDocumento, ActXCtaCred.NroCuenta

        txtMonto.Text = Format(nMontoVoucher, "#,##0.00")
        
        ReDim nMovNroRVD(5)
        nMovNroRVD(0) = nMovNro
        nMovNroRVD(1) = nMovNroPend
        nMovNroRVD(2) = 0
        nMovNroRVD(3) = 0
        nMovNroRVD(4) = 0
        nMovNroRVD(5) = 0
        
        If Len(sVaucher) = 0 Then
            fsNroDoc = sVaucher
        Else
            fsNroDoc = Trim(Mid(sVaucher, 1, Len(sVaucher) - 10))
        End If
    End If
    
    txtDevolucion.Text = CCur(txtMonto.Text) - CCur(txtMontoPagar.Text)
    txtDevolucion.Text = Format(IIf(CCur(txtDevolucion.Text) < 0, "0.00", txtDevolucion.Text), "#,##0.00")
End If
End Sub

Private Sub cmdBuscar_Click()
Dim oPersona As COMDPersona.UCOMPersona
Dim loPersCreditos As COMDCredito.DCOMCredito
Dim lrCreditos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

LimpiarDatos

On Error GoTo ControlError

Set oPersona = Nothing
Set oPersona = New COMDPersona.UCOMPersona
Set oPersona = frmBuscaPersona.Inicio
If oPersona Is Nothing Then Exit Sub

If Trim(oPersona.sPersCod) <> "" Then
    Set loPersCreditos = New COMDCredito.DCOMCredito
    Set lrCreditos = loPersCreditos.CreditosALiquidarSegDesgravamen(oPersona.sPersCod)
    Set loPersCreditos = Nothing
End If

If Not (lrCreditos.EOF And lrCreditos.BOF) Then
Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(oPersona.sPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        ActXCtaCred.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        ActXCtaCred.Enabled = False
        Call ActXCtaCred_KeyPress(13)
    End If
Else
    MsgBox "Persona No cuenta con créditos preparados para liquidar", vbInformation, "Aviso"
End If
Set loCuentas = Nothing
Exit Sub
ControlError:
MsgBox "Error: " & err.Description, vbCritical, "Error"
End Sub

Private Sub cmdCancelar_Click()
LimpiarDatos
End Sub

Private Sub cmdGrabar_Click()
Dim lsMsjError As String
Dim lsImpreBoleta As String
Dim oNCred As COMNCredito.NCOMCredito
'MARG ERS073***
Dim nMovNro As Long
'END MARG******
Set oNCred = New COMNCredito.NCOMCredito

If Not ValidaGrabar Then Exit Sub

ReDim ArrDocRec(3)

If CInt(Trim(Right(cboFormaPago.Text, 4))) = gColocTipoPagoCheque Then
    ArrDocRec(0) = oDocRec.fnTpoDoc
    ArrDocRec(1) = oDocRec.fsPersCod
    ArrDocRec(2) = oDocRec.fsIFTpo
    ArrDocRec(3) = oDocRec.fsIFCta
End If

If MsgBox("¿Estas seguro de grabar la operacion?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'MARG ERS073********************************
Dim clsExo As New COMNCaptaServicios.NCOMCaptaServicios
If Not clsExo.EsCuentaExoneradaLavadoDinero(Trim(ActXCtaCred.NroCuenta)) Then
  Dim sPersLavDinero As String
  Dim nMontoLavDinero As Double, nTC As Double
  Dim clsLav As New COMNCaptaGenerales.NCOMCaptaDefinicion, nMoneda As Integer, nMonto As Double
  
  Dim bLavDinero As Boolean
  Dim loLavDinero As frmMovLavDinero
  Set loLavDinero = New frmMovLavDinero
    nMonto = CDbl(txtMontoPagar.Text)
    Set clsExo = Nothing
    sPersLavDinero = ""
    nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
    Set clsLav = Nothing
    nMoneda = gMonedaNacional
    If nMoneda = gMonedaNacional Then
        Dim clsTC As COMDConstSistema.NCOMTipoCambio
        Set clsTC = New COMDConstSistema.NCOMTipoCambio
        nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
        Set clsTC = Nothing
    Else
        nTC = 1
    End If
    If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
        sPersLavDinero = ""
        loLavDinero.TitPersLavDinero = Trim(lcCodPersTitu)
        sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, Trim(ActXCtaCred.NroCuenta), "LIQ.CRED/SEG. DESGRAVAMEN", True, "", , , , , Mid(Trim(ActXCtaCred.NroCuenta), 9, 1), , gnTipoREU, gnMontoAcumulado, gsOrigen, , gsOpeCod)
        bLavDinero = True
         If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
    End If
 Else
    Set clsExo = Nothing
 End If
'END MARG***********************************

lsMsjError = oNCred.GrabarLiquidacionCredSegDes(lsImpreBoleta, fnIdSiniestro, fnIdRegAtencion, CCur(txtDevolucion.Text), Trim(ActXCtaCred.NroCuenta), _
                gsCodUser, gsCodAge, gsNomAge, gdFecSis, CCur(txtMontoPagar.Text) - CCur(txtITFPagar.Text) - CCur(txtAjuste.Text), CCur(txtITFPagar.Text), ArrDistMontos, _
                CInt(Trim(Right(cboFormaPago.Text, 4))), fsNroDoc, sLpt, gsInstCmac, gbImpTMU, ArrDocRec, nMovNroRVD, nMovNro)
                'APRI20171207 ERS028-2017 AGREGÒ (- CCur(txtAjuste.Text)) AL PARAMETRO pnMontoPago
If Trim(lsMsjError) <> "" Then
    MsgBox lsMsjError, vbCritical, "Error"
    Exit Sub
End If
                        
'MARG ERS073********************************
    If bLavDinero Then
        Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , nMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4)
    End If
    Set loLavDinero = Nothing
'END MARG***********************************
Set vPrevio = New previo.clsprevio
vPrevio.PrintSpool sLpt, lsImpreBoleta

Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes
    vPrevio.PrintSpool sLpt, lsImpreBoleta
Loop

Set vPrevio = Nothing
cmdCancelar_Click
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CargaControles
End Sub

Private Sub LimpiarDatos()
ActXCtaCred.Enabled = True
ActXCtaCred.NroCuenta = ""
ActXCtaCred.CMAC = gsCodCMAC
ActXCtaCred.Age = gsCodAge
ActXCtaCred.EnabledCMAC = False

fnIdSiniestro = 0
fnIdRegAtencion = 0

lblCodTitular.Caption = ""
lblTpoProd.Caption = ""
lblTpoCred.Caption = ""
lblMoneda.Caption = ""
lblMetodLiq.Caption = ""
lblMontoCred.Caption = ""
lblCuotas.Caption = ""
lblSaldo.Caption = ""

txtCapitalFecha.Text = "0.00"
txtIntCompFecha.Text = "0.00"
txtIntMoraFecha.Text = "0.00"
txtGastosFecha.Text = "0.00"
txtITFFecha.Text = "0.00"
txtTotalFecha.Text = "0.00"

txtCapitalPagar.Text = "0.00"
txtIntCompPagar.Text = "0.00"
txtIntMoraPagar.Text = "0.00"
txtGastosPagar.Text = "0.00"
txtAjuste.Text = "0.00" 'APRI20171205 ERS028-2017
txtITFPagar.Text = "0.00"
txtTotalPagar.Text = "0.00"
cboFormaPago.ListIndex = -1
txtMonto.Text = "0.00"
txtMontoPagar.Text = "0.00"
txtDevolucion.Text = "0.00"

ActXCtaCred.Enabled = True
cmdBuscar.Enabled = True
cmdGrabar.Enabled = False
cboFormaPago.Enabled = False

Set ArrDistMontos = Nothing
fsNroDoc = ""
End Sub

Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

Set oCons = New COMDConstantes.DCOMConstantes

Set R = oCons.RecuperaConstantes(gColocTipoPago, 6, , 2)

Set oCons = Nothing
Call Llenar_Combo_con_Recordset(R, cboFormaPago)

LimpiarDatos
End Sub

Private Sub CargaDatos(ByVal psCtaCod As String)
Dim oDCred As COMDCredito.DCOMCredito
Dim rsCred As ADODB.Recordset
Dim nTotal As Currency
Dim nITF As Currency
Dim oITF As COMDConstSistema.FCOMITF
Dim nRedondeoITF  As Double

Set oDCred = New COMDCredito.DCOMCredito
Set rsCred = oDCred.ObtenerDatosALiquidarCred(psCtaCod)

If Not (rsCred.EOF And rsCred.BOF) Then
    ActXCtaCred.Enabled = False
    cmdBuscar.Enabled = False
    cmdGrabar.Enabled = True
    cboFormaPago.Enabled = True
        
    fnIdSiniestro = CLng(rsCred!nIdSiniestro)
    fnIdRegAtencion = CLng(rsCred!nIdRegAtencion)
    lblCodTitular.Caption = rsCred!Cliente
    lblTpoProd.Caption = rsCred!TpoProd
    lblTpoCred.Caption = rsCred!TpoCred
    lblMoneda.Caption = rsCred!Moneda
    lblMetodLiq.Caption = rsCred!MetoLiqui
    lblMontoCred.Caption = Format(rsCred!MontoCred, "###," & String(15, "#") & "#0.00") & " "
    lblCuotas.Caption = rsCred!Cuotas
    lblSaldo.Caption = Format(rsCred!Saldo, "###," & String(15, "#") & "#0.00") & " "
    'MARG ERS073**********************
    lcCodPersTitu = rsCred!CodPersTitu
    'END MARG*************************
    'Distribucion Montos
    'Fecha
    Call CargaDataDistMontosFecha(fnIdSiniestro, psCtaCod)
    
    'Pagar
    txtCapitalPagar.Text = Format(rsCred!nCapital, "###," & String(15, "#") & "#0.00")
    txtIntCompPagar.Text = Format(rsCred!nIntComp, "###," & String(15, "#") & "#0.00")
    txtIntMoraPagar.Text = Format(rsCred!nIntMora, "###," & String(15, "#") & "#0.00")
    txtGastosPagar.Text = Format(rsCred!nGastos, "###," & String(15, "#") & "#0.00")
    txtAjuste.Text = Format(rsCred!nAjuste, "###," & String(15, "#") & "#0.00") 'APRI201711205 ERS028-2017
    'ReDim ArrDistMontos(3)
    ReDim ArrDistMontos(4) 'APRI201711205 ERS028-2017
    ArrDistMontos(0) = CCur(rsCred!nCapital)
    ArrDistMontos(1) = CCur(rsCred!nIntComp)
    ArrDistMontos(2) = CCur(rsCred!nIntMora)
    ArrDistMontos(3) = CCur(rsCred!nGastos)
        ArrDistMontos(4) = CCur(rsCred!nAjuste)

    nTotal = CCur(txtCapitalPagar.Text) + CCur(txtIntCompPagar.Text) + CCur(txtIntMoraPagar.Text) + CCur(txtGastosPagar.Text)

    Set oITF = New COMDConstSistema.FCOMITF
    oITF.fgITFParametros
    
    nITF = nTotal * oITF.gnITFPorcent
    nITF = oITF.CortaDosITF(nITF)

    nRedondeoITF = fgDiferenciaRedondeoITF(nITF)
    If nRedondeoITF > 0 Then
        nITF = nITF - nRedondeoITF
    End If
    
    'nTotal = nTotal + nITF
    nTotal = nTotal + nITF + CCur(txtAjuste.Text) 'APRI201711205 ERS028-2017
    
    txtITFPagar.Text = Format(nITF, "###," & String(15, "#") & "#0.00")
    txtTotalPagar.Text = Format(nTotal, "###," & String(15, "#") & "#0.00")
    
    txtMontoPagar.Text = txtTotalPagar.Text
Else
    MsgBox "No se encontraron datos del crédito o ya fue liquidado.", vbInformation, "Aviso"
End If

Set oDCred = Nothing
Set rsCred = Nothing
End Sub

Private Sub CargaDataDistMontosFecha(ByVal pnCodigo As Long, ByVal pCtaCod As String)
Dim oDCredito As COMDCredito.DCOMCredito
Dim rsDatos As ADODB.Recordset
Dim nTotal As Currency
Dim nITF As Currency
Dim oITF As COMDConstSistema.FCOMITF
Dim nRedondeoITF  As Double

Set oDCredito = New COMDCredito.DCOMCredito
Set rsDatos = oDCredito.GestionSiniestroDistMontosFecha(pnCodigo, pCtaCod, CDate(gdFecSis))

nITF = 0
nTotal = 0
    
If Not (rsDatos.EOF And rsDatos.BOF) Then
    txtCapitalFecha.Text = Format(rsDatos!nCapital, "###," & String(15, "#") & "#0.00")
    txtIntCompFecha.Text = Format(rsDatos!nIntComp, "###," & String(15, "#") & "#0.00")
    txtIntMoraFecha.Text = Format(rsDatos!nIntMora, "###," & String(15, "#") & "#0.00")
    txtGastosFecha.Text = Format(rsDatos!nGastos, "###," & String(15, "#") & "#0.00")
    
    nTotal = CCur(txtCapitalFecha.Text) + CCur(txtIntCompFecha.Text) + CCur(txtIntMoraFecha.Text) + CCur(txtGastosFecha.Text)

    Set oITF = New COMDConstSistema.FCOMITF
    oITF.fgITFParametros
    
    nITF = nTotal * oITF.gnITFPorcent
    nITF = oITF.CortaDosITF(nITF)

    nRedondeoITF = fgDiferenciaRedondeoITF(nITF)
    If nRedondeoITF > 0 Then
        nITF = nITF - nRedondeoITF
    End If
    
    nTotal = nTotal + nITF
End If

txtITFFecha.Text = Format(nITF, "###," & String(15, "#") & "#0.00")
txtTotalFecha.Text = Format(nTotal, "###," & String(15, "#") & "#0.00")

Set oDCredito = Nothing
Set rsDatos = Nothing
End Sub


Private Function ValidaGrabar() As Boolean
ValidaGrabar = True
If cboFormaPago.ListIndex = -1 Then
    MsgBox "Favor de selccionar la forma de pago.", vbInformation, "Aviso"
    cboFormaPago.SetFocus
    ValidaGrabar = False
    Exit Function
End If
If CInt(Trim(Right(cboFormaPago.Text, 2))) = gColocTipoPagoCheque Then
    If Trim(fsNroDoc) = "" Then
        MsgBox "Cheque No es Valido", vbInformation, "Aviso"
        cboFormaPago.SetFocus
        ValidaGrabar = False
        Exit Function
    End If
    If Not ValidaSeleccionCheque Then
        MsgBox "Ud. debe seleccionar el Cheque para continuar", vbInformation, "Aviso"
        If cboFormaPago.Visible And cboFormaPago.Enabled Then cboFormaPago.SetFocus
        ValidaGrabar = False
        Exit Function
    End If

    Dim nDifValorCh As Double
    Dim nDifTotalCh As Double
    Dim nPagadoTotal As Double


    nDifValorCh = Format(CDbl(oDocRec.fnMonto), "0.00")
    nPagadoTotal = CDbl(txtMontoPagar.Text)
    nDifTotalCh = (CDbl(nDifValorCh) - CDbl(nPagadoTotal))
    
    If nDifTotalCh < 0 Then
        MsgBox "No se puede realizar el Pago con Cheque solo dispone de: " & Format(nDifValorCh, gsFormatoNumeroView), vbInformation, "Aviso"
        ValidaGrabar = False
        Exit Function
    End If
End If

If CInt(Trim(Right(cboFormaPago.Text, 10))) = gColocTipoPagoVoucher Then
    If Trim(fsNroDoc) = "" Then
        MsgBox "Voucher No es Valido", vbInformation, "Aviso"
        cboFormaPago.SetFocus
        ValidaGrabar = False
        Exit Function
    End If
    
    Dim nPagadoTotalV As Double
    nPagadoTotalV = CDbl(txtMontoPagar.Text)
    
    If nPagadoTotalV > nMontoVoucher Then
        MsgBox "No se puede realizar el Pago con Voucher solo dispone de: " & Format(nMontoVoucher, "#0.00"), vbInformation, "Aviso"
        ValidaGrabar = False
        Exit Function
    End If
End If
    

End Function

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

