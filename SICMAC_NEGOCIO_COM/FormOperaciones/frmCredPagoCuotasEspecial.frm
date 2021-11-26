VERSION 5.00
Begin VB.Form frmCredPagoCuotasEspecial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adelanto de Cuota / Pago Anticipado"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "frmCredPagoCuotasEspecial.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8040
      TabIndex        =   50
      Top             =   4680
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   49
      Top             =   4680
      Width           =   1050
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   48
      Top             =   4680
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   25
      Top             =   2400
      Width           =   9015
      Begin VB.Frame FraDisminucion 
         Caption         =   "Disminución :"
         Height          =   1335
         Left            =   6960
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
         Begin VB.OptionButton OptTipoCuota 
            Caption         =   "Número de cuotas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton OptTipoCuota 
            Caption         =   "Monto de cuotas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Value           =   -1  'True
            Width           =   1575
         End
      End
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
         Left            =   1305
         MaxLength       =   15
         TabIndex        =   30
         Top             =   1080
         Width           =   1380
      End
      Begin VB.ComboBox cboFormaPago 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   600
         Width           =   1785
      End
      Begin VB.ComboBox cboTipoPago 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   240
         Width           =   1785
      End
       Begin SICMACT.ActXCodCta txtCuentaCargo 
         Height = 375
Left = 3240
TabIndex = 39
Top = 480
Visible = 0   'False
Width = 3630
_extentx = 6403
_extenty = 661
texto = "Cuenta N°:"
End
      Begin VB.Label lblNewCuotaPend 
         Appearance = 0  'Flat
BackColor = &H80000004&
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 700
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 8040
TabIndex = 47
Top = 1800
Width = 495
End
      Begin VB.Label Label33 
         Caption = "Nueva Cuota Pendiente :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 6240
TabIndex = 46
Top = 1800
Width = 1815
End
      Begin VB.Label lblProxFecPago 
         Appearance = 0  'Flat
BackColor = &H80000004&
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 700
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 4800
TabIndex = 45
Top = 1800
Width = 1095
End
      Begin VB.Label Label31 
         Caption = "Prox. Fecha Pago :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 3360
TabIndex = 44
Top = 1800
Width = 1695
End
      Begin VB.Label lblNewSaldoCap 
         Appearance = 0  'Flat
BackColor = &H80000004&
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 700
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 1800
TabIndex = 43
Top = 1800
Width = 1095
End
      Begin VB.Label Label29 
         Caption = "Nuevo Saldo Capital :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 120
TabIndex = 42
Top = 1800
Width = 1695
End
      Begin VB.Label Label27 
         AutoSize = -1  'True
Caption = "Nº Documento"
Height = 195
Left = 3330
TabIndex = 41
Top = 510
Width = 1050
End
      Begin VB.Label LblNumDoc 
         Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
         BeginProperty Font
Name = "MS Sans Serif"
Size = 8.25
Charset = 0
Weight = 700
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &HC00000&
Height = 285
Left = 4545
TabIndex = 40
Top = 480
Width = 1665
End
      Begin VB.Line Line1 
         BorderColor = &H80000003&
X1 = 120
X2 = 8880
Y1 = 1560
Y2 = 1560
End
      Begin VB.Label Label25 
         AutoSize = -1  'True
Caption = "Monto a Pagar :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 210
Left = 120
TabIndex = 35
Top = 1095
Width = 1125
End
      Begin VB.Label Label26 
         AutoSize = -1  'True
Caption = "ITF :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 210
Left = 3000
TabIndex = 34
Top = 1110
Width = 300
End
      Begin VB.Label lblITF 
         Alignment = 1  'Right Justify
Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
Caption = "0.00"
         BeginProperty Font
Name = "MS Sans Serif"
Size = 8.25
Charset = 0
Weight = 700
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &HC00000&
Height = 285
Left = 3390
TabIndex = 33
Top = 1080
Width = 900
End
      Begin VB.Label lblPagoTotal 
         Alignment = 1  'Right Justify
Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
Caption = "0.00"
         BeginProperty Font
Name = "MS Sans Serif"
Size = 8.25
Charset = 0
Weight = 700
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H0&
Height = 285
Left = 5520
TabIndex = 32
Top = 1080
Width = 1260
End
      Begin VB.Label Label28 
         AutoSize = -1  'True
Caption = "Pago Total :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 210
Left = 4605
TabIndex = 31
Top = 1110
Width = 840
End
      Begin VB.Label Label24 
         AutoSize = -1  'True
Caption = "Forma de Pago :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 210
Left = 120
TabIndex = 29
Top = 630
Width = 1170
End
      Begin VB.Label Label23 
         AutoSize = -1  'True
Caption = "Tipo de Pago :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 210
Left = 120
TabIndex = 27
Top = 270
Width = 1020
End
End
   Begin VB.Frame Frame1 
      Height = 1695
Left = 120
TabIndex = 2
Top = 600
Width = 9015
      Begin VB.Label lblMonto2Cuota 
         Alignment = 1  'Right Justify
Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 7560
TabIndex = 24
Top = 1320
Width = 1215
End
      Begin VB.Label Label21 
         Caption = "Monto 2 Cuota :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 6360
TabIndex = 23
Top = 1320
Width = 1215
End
      Begin VB.Label lblMontoCuota 
         Alignment = 1  'Right Justify
Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 7560
TabIndex = 22
Top = 960
Width = 1215
End
      Begin VB.Label Label19 
         Caption = "Monto Cuota :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 6360
TabIndex = 21
Top = 960
Width = 1095
End
      Begin VB.Label lblMoneda 
         Alignment = 2  'Center
Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 7560
TabIndex = 20
Top = 600
Width = 1215
End
      Begin VB.Label Label17 
         Caption = "Moneda :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 6360
TabIndex = 19
Top = 600
Width = 735
End
      Begin VB.Label lblDOI 
         Alignment = 2  'Center
Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 7560
TabIndex = 18
Top = 240
Width = 1215
End
      Begin VB.Label Label15 
         Caption = "D.O.I. :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 6360
TabIndex = 17
Top = 240
Width = 615
End
      Begin VB.Label lblMetLiquid 
         Alignment = 2  'Center
Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 4680
TabIndex = 16
Top = 1320
Width = 1335
End
      Begin VB.Label Label13 
         Caption = "Met Liquid :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 3480
TabIndex = 15
Top = 1320
Width = 1095
End
      Begin VB.Label lblFecDesemb 
         Alignment = 2  'Center
Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 4680
TabIndex = 14
Top = 600
Width = 1335
End
      Begin VB.Label Label11 
         Caption = "Fecha Desemb :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 3480
TabIndex = 13
Top = 600
Width = 1215
End
      Begin VB.Label lblFecVenc 
         Alignment = 2  'Center
Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 4680
TabIndex = 12
Top = 960
Width = 1335
End
      Begin VB.Label Label9 
         Caption = "Fecha Venc :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 3480
TabIndex = 11
Top = 960
Width = 1095
End
      Begin VB.Label lblDeudaAct 
         Alignment = 1  'Right Justify
Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 1560
TabIndex = 10
Top = 1320
Width = 1335
End
      Begin VB.Label Label7 
         Caption = "Deuda Actual :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 120
TabIndex = 9
Top = 1320
Width = 1095
End
      Begin VB.Label lblMontoCred 
         Alignment = 1  'Right Justify
Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 1560
TabIndex = 8
Top = 600
Width = 1335
End
      Begin VB.Label Label5 
         Caption = "Monto del Crédito :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 120
TabIndex = 7
Top = 600
Width = 1335
End
      Begin VB.Label lblSaldoCap 
         Alignment = 1  'Right Justify
Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 1560
TabIndex = 6
Top = 960
Width = 1335
End
      Begin VB.Label Label3 
         Caption = "Saldo Capital :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 120
TabIndex = 5
Top = 960
Width = 1095
End
      Begin VB.Label lblNomCliente 
         Appearance = 0  'Flat
BackColor = &H80000005&
BorderStyle = 1  'Fixed Single
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ForeColor = &H80000008&
Height = 255
Left = 840
TabIndex = 4
Top = 240
Width = 5175
End
      Begin VB.Label Label1 
         Caption = "Cliente :"
         BeginProperty Font
Name = "Arial"
Size = 8.25
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 255
Left = 120
TabIndex = 3
Top = 240
Width = 735
End
End
   Begin VB.CommandButton CmdBuscar 
      Caption = "&Buscar"
      BeginProperty Font
Name = "MS Sans Serif"
Size = 8.25
Charset = 0
Weight = 700
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
Height = 360
Left = 3780
TabIndex = 0
Top = 105
Width = 900
End
   Begin SICMACT.ActXCodCta ActxCta 
      Height = 435
Left = 120
TabIndex = 1
Top = 75
Width = 3660
_extentx = 6456
_extenty = 767
texto = "Credito :"
enabledcmac = -1  'True
enabledcta = -1  'True
enabledprod = -1  'True
enabledage = -1  'True
End
End
Attribute VB_Name = "frmCredPagoCuotasEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredPagoCuotaEspecial
'** Descripción : Formulario para realizar pagos anticipados o adelelanto de cuotas a solicitud
'**               del cliente creado segun TI-ERS008-2015
'** Creación : JUEZ, 20150420 09:00:00 AM
'**********************************************************************************************

Option Explicit

Public nProducto As Producto

Private oCredito As COMNCredito.NCOMCredito

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
Dim sPersCod As String
Dim nInteresDesagio As Double

Dim nMontoPago As Double
Dim nITF As Double
Dim nCalendDinamico As Integer
Dim bRFA As Boolean
Dim lnDiasAtraso As Integer

'Lavado de Dinero
Dim bExoneradaLavado As Boolean
Dim sPerscodLav As String
Dim sNombreLav As String
Dim sDireccionLav As String
Dim sDocIdLav As String

'Variables agregadas para el uso de los Componentes
Private bOperacionEfectivo As Boolean
Private nMontoLavDinero As Double
Private nTC As Double

Private bantxtmonpag As Boolean

Private bActualizaMontoPago As Boolean

Dim pnValorChq As Double

Dim lsTemp As String
Dim lsAgeCodAct As String
Dim lsTpoProdCod As String
Dim lsTpoCredCod As String
Dim nRedondeoITF As Double

Dim nMontoPag2CuotxVenc As Double
Dim bCuotasVencidas As Boolean

Dim fnPersPersoneria As Integer
Dim oDocRec As UDocRec
Dim bInstFinanc As Boolean

Private nMovNroRVD() As Variant
Private nMontoVoucher As Currency
Dim lnMontoPendienteIntGracia As Double
Dim lnMontoPagoInicio As Double
Dim lnMontGasto As Double
Dim lnMontIntComp As Double
Dim nCuotasApr As Integer
Dim nCuotasPend As Integer
Dim oVisto As frmVistoElectronico 'ADD BY MARG ERS052-2017
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

'CTI2 20190110 ERS075-2018 - INI ****
Dim bExoneraGastoConvenio As Boolean
'CTI2 20190110 ERS075-2018 - FIN ****

Dim bValidaActualizacionLiq As Boolean 'RIRO 20200911 Actualización Liquidación

Public Sub Inicia(ByVal sCodOpe As String)
'Dim oVisto As frmVistoElectronico 'COMMENT BY MARG ERS052-2017
Dim bResultadoVisto As Boolean

    Set oVisto = New frmVistoElectronico
    bResultadoVisto = oVisto.Inicio(3)
    If Not bResultadoVisto Then
        Exit Sub
    End If
    
    sOperacion = sCodOpe
    Me.Show 1
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not CargaDatos(ActxCta.NroCuenta) Then
            HabilitaActualizacion False
        Else
            HabilitaActualizacion True
        End If
    End If
End Sub

Private Sub cboFormaPago_Click()
    LblNumDoc.Caption = ""
    txtCuentaCargo.NroCuenta = ""
    TxtMonPag.Locked = False
    If cboFormaPago.ListIndex <> -1 Then
        cmdGrabar.Enabled = False
        If CInt(Trim(Right(cboFormaPago.Text, 10))) = gColocTipoPagoCheque Then
            Dim oform As New frmChequeBusqueda
            Set oDocRec = oform.iniciarBusqueda(Val(Mid(ActxCta.NroCuenta, 9, 1)), TipoOperacionCheque.CRED_Pago, ActxCta.NroCuenta)
            Set oform = Nothing
            LblNumDoc.Caption = oDocRec.fsNroDoc
            TxtMonPag.Text = Format(DeducirMontoxITF(oDocRec.fnMonto), "#0.00") 'Restamos el ITF al disponible
            TxtMonPag.Locked = True
            LblNumDoc.Visible = True
            txtCuentaCargo.Visible = False
            ReDim nMovNroRVD(6)
        ElseIf CInt(Trim(Right(cboFormaPago.Text, 10))) = gColocTipoPagoCargoCta Then
            If Not SeleccionarCtaCargo Then
                cboFormaPago.ListIndex = 0
                Exit Sub
            End If
            LblNumDoc.Visible = False
            txtCuentaCargo.Visible = True
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
                        
            lnTipMot = 10 ' Pago Credito
            oformVou.iniciarFormularioDeposito CInt(Mid(ActxCta.NroCuenta, 9, 1)), lnTipMot, sGlosa, sIF, sVaucher, nMontoVoucher, sPersCod, nMovNro, nMovNroPend, sNombre, sDireccion, sDocumento, ActxCta.NroCuenta
            LblNumDoc.Visible = True
            ReDim nMovNroRVD(5)
            nMovNroRVD(0) = nMovNro
            nMovNroRVD(1) = nMovNroPend
            nMovNroRVD(2) = lnMontoPendienteIntGracia
            nMovNroRVD(3) = 0
            nMovNroRVD(4) = lnMontIntComp
            nMovNroRVD(5) = lnMontGasto
            If Len(sVaucher) = 0 Then
                LblNumDoc.Caption = sVaucher
            Else
                LblNumDoc.Caption = Trim(Mid(sVaucher, 1, Len(sVaucher) - 10))
            End If
            TxtMonPag.Text = Format(DeducirMontoxITF(nMontoVoucher), "#,##0.00") 'Restamos el ITF al disponible
        Else
            ReDim nMovNroRVD(6)
            nMovNroRVD(0) = nMovNro
            nMovNroRVD(1) = nMovNroPend
            nMovNroRVD(2) = lnMontoPendienteIntGracia
            nMovNroRVD(3) = 0
            nMovNroRVD(4) = lnMontIntComp
            nMovNroRVD(5) = lnMontGasto
            LblNumDoc.Visible = False
            txtCuentaCargo.Visible = False
        End If
        bActualizaMontoPago = True
        If TxtMonPag.Enabled Then TxtMonPag.SetFocus
    End If
End Sub

Private Sub cboTipoPago_Click()
If cboFormaPago.ListIndex <> -1 Then
    cmdGrabar.Enabled = False
    If CInt(Trim(Right(cboTipoPago.Text, 10))) = gPagoAnticipado Then
        FraDisminucion.Visible = True
        FraDisminucion.Enabled = True
        'JUEZ 20150626 ********************
        If nCuotasApr - (nCuotasPend - 1) = 1 Then
            OptTipoCuota(1).Value = 0
            OptTipoCuota(1).Enabled = False
        Else
            OptTipoCuota(1).Enabled = True
        End If
        'END JUEZ *************************
        'marg ers004-2017**
        Me.lblNewSaldoCap.Visible = False
        Me.lblProxFecPago.Visible = False
        Me.lblNewCuotaPend.Visible = False
        'end marg**********
    Else
        FraDisminucion.Visible = False
        OptTipoCuota(0).Value = 0
        OptTipoCuota(1).Value = 0
        'marg ers004-2017**
        Me.lblNewSaldoCap.Visible = True
        Me.lblProxFecPago.Visible = True
        Me.lblNewCuotaPend.Visible = True
        'end marg**********
    End If
    bActualizaMontoPago = True
    TxtMonPag.Text = 0
    If TxtMonPag.Enabled Then TxtMonPag.SetFocus
Else
    FraDisminucion.Enabled = False
End If
End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona
    
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set oCredito = New COMDCredito.DCOMCredito
        Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstVigMor, gColocEstVigVenc, gColocEstVigNorm, gColocEstRefMor, gColocEstRefVenc, gColocEstRefNorm))
        R.Close
        Set R = Nothing
        Set oCredito = Nothing
        Call FrmVerCredito.Inicio(oPers.sPersCod, , , True, Me.ActxCta)
        Me.ActxCta.SetFocusCuenta
    End If
    Set oPers = Nothing
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaPantalla
    Call HabilitaActualizacion(False)
    cmdGrabar.Enabled = False
    If Not (oCredito Is Nothing) Then Set oCredito = Nothing
End Sub

Private Sub cmdGrabar_Click()
Dim sError As String
Dim vPrevio As previo.clsprevio
Dim sImprePlanPago As String
Dim sImpreBoleta As String
Dim oCredD As COMDCredito.DCOMCreditos
Dim sVisPersLavDinero As String
Dim loLavDinero As frmMovLavDinero
Dim oCred As COMDCredito.DCOMCredActBD
Set loLavDinero = New frmMovLavDinero

Dim objPersona As COMDPersona.DCOMPersonas
Set objPersona = New COMDPersona.DCOMPersonas

Dim oMov As COMDMov.DCOMMov
Set oMov = New COMDMov.DCOMMov
Dim fnCondicion As Integer
Dim regPersonaRealizaPago As Boolean
Dim pnMotCancAnt As Integer
Dim oDCred As COMDCredito.DCOMCredito
Dim oSeguridad As New COMManejador.Pista

'JIPR 20180625 INICIO
Dim oCredCodFactElect As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim scPersCod As String
Dim scPersIDTpo As String
Dim sSerie As String
Dim sCorrelativo As String
Set R = New ADODB.Recordset
'JIPR 20180625 FIN

'CTI3 ERS082-2019
Dim pnCancelado(3) As Variant 'CTI3 28122018
pnCancelado(0) = 0  'CTI3 28122018
pnCancelado(1) = 0  'CTI3 28122018
pnCancelado(2) = 0  'CTI3 28122018
pnCancelado(3) = 0  'CTI3 28122018

    On Error GoTo ErrorCmdGrabar_Click
    
    If CInt(Trim(Right(cboTipoPago.Text, 10))) = gPagoAnticipado Then
        If OptTipoCuota(0).Value = 0 And OptTipoCuota(1).Value = 0 Then
            MsgBox "Debe seleccionar un tipo de Pago Anticipado", vbInformation, "Aviso"
            OptTipoCuota(0).SetFocus
            Exit Sub
        End If
    End If
        
    'RIRO20200911 VALIDA LIQUIDACION ***************
    If CInt(Trim(Right(cboTipoPago.Text, 10))) = gPagoAnticipado Then
        Dim oDCredTmp As COMDCredito.DCOMCredito
        Set oDCredTmp = New COMDCredito.DCOMCredito
        If oCredito Is Nothing Then
            Set oCredito = New COMNCredito.NCOMCredito
        End If
        bValidaActualizacionLiq = False
        bValidaActualizacionLiq = oCredito.VerificaActualizacionLiquidacion(ActxCta.NroCuenta)
        If Not bValidaActualizacionLiq Then
            MsgBox "El crédito no tiene actualizados sus datos de liquidación, no podrá realizar cancelaciones " & _
            "anticipadas ni pagos anticipados a menos que actualice estos datos. Deberá comunicarse con el área de T.I.", vbExclamation, "Aviso"
            Exit Sub
        End If
    End If
    'END RIRO **************************************
        
    Dim PerAut As COMDPersona.DCOMPersonas
    Dim rs As New ADODB.Recordset
    Dim nValorBloqueo As Boolean
    Set PerAut = New COMDPersona.DCOMPersonas
    Set rs = New ADODB.Recordset
    
    Set rs = PerAut.DevuelvePersBloqueaRecuperaCred(Trim(sPersCod), Trim(ActxCta.NroCuenta))
    Set PerAut = Nothing
    If Not (rs.EOF And rs.BOF) Then
        nValorBloqueo = IIf(rs!dVigente, True, False)
         rs.Close
         Set rs = Nothing
        If nValorBloqueo Then
            MsgBox "Ud. NO podrá continuar, persona registra Bloqueo en Recuperaciones, Comuniquese con el Area de Recuperaciones", vbCritical, "Aviso"
            Exit Sub
        End If
    End If
        
    If CInt(Trim(Right(cboFormaPago.Text, 2))) = gColocTipoPagoCheque Then
        If Trim(Me.LblNumDoc.Caption) = "" Then
            MsgBox "Cheque No es Valido", vbInformation, "Aviso"
            Me.cboFormaPago.SetFocus
            Exit Sub
        End If
        If Not ValidaSeleccionCheque Then
            MsgBox "Ud. debe seleccionar el Cheque para continuar", vbInformation, "Aviso"
            If cboFormaPago.Visible And cboFormaPago.Enabled Then cboFormaPago.SetFocus
            Exit Sub
        End If

        Dim nDifValorCh As Double
        Dim nDifTotalCh As Double
        Dim nPagadoTotal As Double
        nDifValorCh = Format(CDbl(oDocRec.fnMonto), "0.00")
        nPagadoTotal = CDbl(lblPagoTotal.Caption)
        nDifTotalCh = (CDbl(nDifValorCh) - CDbl(nPagadoTotal))
        If nDifTotalCh < 0 Then
            MsgBox "No se puede realizar el Pago con Cheque solo dispone de: " & Format(nDifValorCh, gsFormatoNumeroView), vbInformation, "Aviso"
            Exit Sub
        End If
    End If
       
    If CInt(Trim(Right(cboFormaPago.Text, 10))) = gColocTipoPagoCargoCta Then
        If Len(txtCuentaCargo.NroCuenta) <> 18 Then
            MsgBox "Favor de ingresar la cuenta de ahorros a debitar", vbInformation, "Aviso"
            Exit Sub
        End If
        
        Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, nMontoPago) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "Aviso"
            Exit Sub
        End If
        
        'Verifica actualización Persona
        Dim lsDireccionActualizada As String
        Dim oPersona As New COMNPersona.NCOMPersona
        
        If oPersona.NecesitaActualizarDatos(sPersCod, gdFecSis) Then
             MsgBox "Para continuar con la Operación Ud. debe actualizar los datos del" & Chr(13) & "Titular: " & lblNomCliente.Caption, vbInformation, "Aviso"
             Dim foPersona As New frmPersona
             If Not foPersona.realizarMantenimiento(sPersCod, lsDireccionActualizada) Then
                 MsgBox "No se ha realizado la actualización de los datos de " & lblNomCliente.Caption & "," & Chr(13) & "la Operación no puede continuar!", vbInformation, "Aviso"
                 Exit Sub
             End If
        End If
        lsDireccionActualizada = ""
    End If

    If CInt(Trim(Right(cboFormaPago.Text, 10))) = gColocTipoPagoVoucher Then
        If Trim(Me.LblNumDoc.Caption) = "" Then
            MsgBox "Voucher No es Valido", vbInformation, "Aviso"
            Me.cboFormaPago.SetFocus
            Exit Sub
        End If
        Dim nPagadoTotalV As Double
        nPagadoTotalV = CDbl(lblPagoTotal.Caption)
        If nPagadoTotalV > nMontoVoucher Then
            MsgBox "No se puede realizar el Pago con Voucher solo dispone de: " & Format(nMontoVoucher, "#0.00"), vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    'JOEP20200705 Cambio ReactivaCovid
Dim objReactCovid As COMDCredito.DCOMCredito
Dim rsReactCovid As ADODB.Recordset
Set objReactCovid = New COMDCredito.DCOMCredito
Set rsReactCovid = objReactCovid.RestrincionCovidReact(Trim(ActxCta.NroCuenta), Trim(gsCodAge), sOperacion, 0, IIf(Trim(Right(Me.cboTipoPago.Text, 2)) = 1, 1, 0), IIf(Trim(Right(Me.cboTipoPago.Text, 2)) = 2, 1, 0))
If Not (rsReactCovid.EOF And rsReactCovid.BOF) Then
    If rsReactCovid!MsgBox <> "" Then
        MsgBox rsReactCovid!MsgBox, vbInformation, "Aviso"
        Exit Sub
    End If
End If
'JOEP20200705 Cambio ReactivaCovid
    
    Dim rsPersVerifica As Recordset
    Dim i As Integer
    Set rsPersVerifica = New Recordset
    
    Set rsPersVerifica = objPersona.ObtenerDatosPersona(sPersCod)
    If rsPersVerifica!nPersIngresoProm = 0 Or rsPersVerifica!cActiGiro1 = "" Then
        If MsgBox("Necesita Registrar la Ocupacion e Ingreso Promedio de: " + lblNomCliente, vbYesNo) = vbYes Then
            frmPersOcupIngreProm.Inicio sPersCod, lblNomCliente, rsPersVerifica!cActiGiro1, rsPersVerifica!nPersIngresoProm
        End If
    End If

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

    Dim oDPersonaAct As COMDPersona.DCOMPersona
    Set oDPersonaAct = New COMDPersona.DCOMPersona
    If oDPersonaAct.VerificaExisteSolicitudDatos(sPersCod) Then
        MsgBox Trim("SE SOLICITA DATOS DEL CLIENTE: " & lblNomCliente.Caption) & "." & Chr(10), vbInformation, "Aviso"
        Call frmActInfContacto.Inicio(sPersCod)
    End If


    
    If MsgBox("Se va a Efectuar el Pago del Credito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    If CInt(Trim(Right(cboFormaPago.Text, 10))) = gColocTipoPagoEfectivo Then
'            nMovNroRVD(0) = 0
'            nMovNroRVD(1) = 0
            nMovNroRVD(2) = lnMontoPendienteIntGracia
            nMovNroRVD(3) = 0
            nMovNroRVD(4) = lnMontIntComp
            nMovNroRVD(5) = lnMontGasto
    End If
    
    If CInt(Trim(Right(cboTipoPago.Text, 10))) = gPagoAnticipado Then
        Dim sMovNroPrepago As String
        nCalendDinamico = 1
        bCalenDinamic = True
        Set oCred = New COMDCredito.DCOMCredActBD
        Call oCred.dUpdateColocacCred(ActxCta.NroCuenta, , , , , , , , , , , , 1)
         
        sMovNroPrepago = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        'oCred.dInsertCredMantPrepago ActxCta.NroCuenta, sMovNroPrepago, 1, IIf(OptTipoCuota(0).value, 1, 0), 0, nCuotasApr - nCuotasPend
        oCred.dInsertCredMantPrepago ActxCta.NroCuenta, sMovNroPrepago, 1, IIf(OptTipoCuota(0).Value, 1, 0), 0, nCuotasApr - (nCuotasPend - 1) 'JUEZ 20150625
        Set oCred = Nothing
        Call TxtMonPag_KeyPress(13)
    End If

    Dim nMonto As Double
    Dim nmoneda As Integer
    nMonto = CDbl(nMontoPago)
    Dim sPersLavDinero As String
    nmoneda = CLng(Mid(ActxCta.NroCuenta, 9, 1))
    
    sPersLavDinero = ""
    If bOperacionEfectivo Then
        If Not bExoneradaLavado Then
            If CDbl(TxtMonPag.Text) >= Round(nMontoLavDinero * nTC, 2) Then
                Call IniciaLavDinero(loLavDinero)
                sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, ActxCta.NroCuenta, Me.Caption, True, "", , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
            End If
        End If
    End If
    
    If loLavDinero.OrdPersLavDinero = "Exit" Then
        Dim oPersonaSPR As UPersona_Cli
        Dim oPersonaU As COMDPersona.UCOMPersona
        Dim nTipoConBN As Integer
        Dim sConPersona As String
        Dim pbClienteReforzado As Boolean
        Dim rsAgeParam As Recordset
        Dim objCred As COMNCredito.NCOMCredito
        Dim lnMonto As Double, lnTC As Double
        Dim ObjTc As COMDConstSistema.NCOMTipoCambio
        
        
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
            MsgBox "El Cliente: " & Trim(lblNomCliente.Caption) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
            frmPersRealizaOpeGeneral.Inicia Me.Caption & " (Persona " & sConPersona & ")", sOperacion
            regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
            
            If Not regPersonaRealizaPago Then
                MsgBox "Se va a proceder a Anular el Pago de la Cuota", vbInformation, "Aviso"
                Exit Sub
            End If
        Else
            fnCondicion = 0
            lnMonto = nMonto
            pbClienteReforzado = False
            
            Set ObjTc = New COMDConstSistema.NCOMTipoCambio
            lnTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
            Set ObjTc = Nothing
        
        
            Set objCred = New COMNCredito.NCOMCredito
            Set rsAgeParam = objCred.obtieneCredPagoCuotasAgeParam(gsCodAge)
            Set objCred = Nothing
            
            If Mid(ActxCta.NroCuenta, 9, 1) = 2 Then
                lnMonto = Round(lnMonto * lnTC, 2)
            End If
        
            If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                If lnMonto >= rsAgeParam!nMontoMin And lnMonto <= rsAgeParam!nMontoMax Then
                    frmPersRealizaOpeGeneral.Inicia Me.Caption, sOperacion
                    regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
                    If Not regPersonaRealizaPago Then
                        MsgBox "Se va a proceder a Anular el Pago de la Cuota", vbInformation, "Aviso"
                        Exit Sub
                    End If
                End If
            End If
            
        End If
    End If

    lsTemp = MDISicmact.SBBarra.Panels(1).Text
    MDISicmact.SBBarra.Panels(1).Text = "Procesando ....."
    Me.cmdGrabar.Enabled = False
    
    oCredito.pbExcluyeGastos = bExoneraGastoConvenio 'CTI2 20181215 ERS075-2018
    Call oCredito.GrabarPagoCuotas(ActxCta.NroCuenta, nMiVivienda, nCalPago, nMontoPago, _
                            gdFecSis, lblMetLiquid.Caption, CInt(Trim(Right(cboFormaPago.Text, 10))), gsCodAge, gsCodUser, gsCodCMAC, Trim(LblNumDoc.Caption), _
                            bRecepcionCmact, sPersCmac, vnIntPendiente, vnIntPendientePagado, bPrepago, sPersLavDinero, CCur(lblITF.Caption), _
                            nInteresDesagio, CDbl(lblDeudaAct.Caption), bCalenDinamic, 0, nCalendDinamTipo, gsNomAge, CInt(ActxCta.Prod), _
                            lblNomCliente.Caption, lblMoneda.Caption, nNroTransac, lblProxFecPago.Caption, sLpt, gsInstCmac, IIf(Trim(Right(Me.cboFormaPago.Text, 2)) = "2", True, False), _
                            Me.LblNumDoc.Caption, sError, sImprePlanPago, sImpreBoleta, lnDiasAtraso, gsProyectoActual, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, lsAgeCodAct, lsTpoProdCod, lsTpoCredCod, 0, , pnCancelado, IIf(Trim(Right(Me.cboTipoPago.Text, 2)) = gAdelantoCuota, 2, 0), 0, txtCuentaCargo.NroCuenta, _
                            oDocRec.fnTpoDoc, oDocRec.fsPersCod, oDocRec.fsIFTpo, oDocRec.fsIFCta, nMovNroRVD)  'CTI3  pnMotCancAnt se cambio por pnCancelado
                            
    If gnMovNro > 0 Then
        Call oMov.InsertaMovRedondeoITF("", 1, CCur(lblITF.Caption) + nRedondeoITF, CCur(lblITF.Caption), gnMovNro)
        Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4)
        'MARG ERS052-2017----
        'If loLavDinero.VisPersLavDinero = "" Then
            oVisto.RegistraVistoElectronico gnMovNro, , gsCodUser, gnMovNro
        'End If
        'END MARG- -------------
    End If
    
    Set oMov = Nothing
    Set oCred = New COMDCredito.DCOMCredActBD
    Call oCred.dUpdateColocacCred(ActxCta.NroCuenta, , , , , , , , , , , , IIf(bCalenDinamic, 0, -1))

    Call oSeguridad.InsertarPista(sOperacion, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, 8, "PAGO DE CUOTA", ActxCta.NroCuenta)
    Set oSeguridad = Nothing
                
    Dim rsPersOcu As Recordset
    Dim nAcumulado As Currency
    Dim nMontoPersOcupacion As Currency
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim sMovNro As String
    
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    Dim clsTC As COMDConstSistema.NCOMTipoCambio
    Set clsTC = New COMDConstSistema.NCOMTipoCambio
    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
    Set clsTC = Nothing
    
    Set rsPersOcu = New Recordset
                    
    Set rsPersOcu = objPersona.ObtenerDatosPersona(sPersCod)
    nAcumulado = objPersona.ObtenerPersAcumuladoMontoOpe(nTC, Mid(Format(gdFecSis, "yyyymmdd"), 1, 6), rsPersOcu!cPersCod)
    nMontoPersOcupacion = objPersona.ObtenerParamPersAgeOcupacionMonto(Mid(rsPersOcu!cPersCod, 4, 2), CInt(Mid(rsPersOcu!cPersCIIU, 2, 2)))

    If nAcumulado >= nMontoPersOcupacion Then
        If Not objPersona.ObtenerPersonaAgeOcupDatos_Verificar(rsPersOcu!cPersCod, gdFecSis) Then
            objPersona.insertarPersonaAgeOcupacionDatos gnMovNro, rsPersOcu!cPersCod, IIf(nmoneda = 1, nMonto, nMonto * nTC), nAcumulado, gdFecSis, sMovNro
        End If
    End If
    
    If regPersonaRealizaPago And gnMovNro > 0 Then
        frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(ActxCta.NroCuenta), fnCondicion
        regPersonaRealizaPago = False
    End If
    
           'INICIO JIPR 20180625
     Set oCredCodFactElect = New COMDCredito.DCOMCredito
     Call oCredCodFactElect.VerificaInteresFactElect(gnMovNro)
     'ANPS COMENTADO 20210820
'     If Not (R.BOF And R.EOF) Then
'     If R!exist = True Then
'     Set R = oCredCodFactElect.RecuperaFactElectDetalle(ActxCta.NroCuenta)
'        If R.RecordCount > 0 Then
'            If Not (R.BOF And R.EOF) Then
'         scPersCod = R!cPersCod
'         scPersIDTpo = R!cPersIDTpo
'         Call oCredCodFactElect.InsertaFactElect(gnMovNro, ActxCta.NroCuenta, scPersCod, scPersIDTpo, gdFecSis)
'
'         Set R = oCredCodFactElect.GenerarCorrelativo(scPersIDTpo, gsCodAge, ActxCta.NroCuenta)  'ANPS 16022021 add codigo de cuenta
'                If R.RecordCount > 0 Then
'                    If Not (R.BOF And R.EOF) Then
'         sSerie = R!cSerie
'         sCorrelativo = R!cNro
'
'         Call oCredCodFactElect.UpdateFactElect(gnMovNro, sSerie, sCorrelativo)
'         Call oCredCodFactElect.InsertarRegVentaFactElect(gdFecSis, scPersCod, ActxCta.NroCuenta, gnMovNro, IIf(nmoneda = 1, 1, 2), gsCodAge)
'
            'MsgBox "Coordinar con el Supervisor de Operaciones, para la Emisión de Facturación Electrónica del Pago de Crédito.", vbInformation, "Aviso" COMENTADO ANPS 20210531
          
'                    End If
'                End If
         Set oCredCodFactElect = Nothing
          
'            End If
'        End If
'     End If
'     End If
     Set oCredCodFactElect = Nothing
     'FIN JIPR 20180625
     
     
    Set vPrevio = New clsprevio
    If sImprePlanPago <> "" Then
        vPrevio.PrintSpool sLpt, sImprePlanPago
        'vPrevio.Show sImprePlanPago, "Impresión de Plan de Pagos"
    End If
    
    vPrevio.PrintSpool sLpt, sImpreBoleta
    'vPrevio.Show sImpreBoleta, "Impresión de Plan de Pagos"
    Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes
        vPrevio.PrintSpool sLpt, sImpreBoleta
        'vPrevio.Show sImpreBoleta, "Impresión de Plan de Pagos"
    Loop
    Set vPrevio = Nothing
    'INICIO JHCU ENCUESTA 16-10-2019
    Encuestas gsCodUser, gsCodAge, "ERS0292019", sOperacion
    'FIN
    gVarPublicas.LimpiaVarLavDinero
    
    Call cmdCancelar_Click
    
    MDISicmact.SBBarra.Panels(1).Text = lsTemp
    
Exit Sub

ErrorCmdGrabar_Click:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Function HabilitaActualizacion(ByVal pbHabilita As Boolean) As Boolean
    cboFormaPago.Enabled = pbHabilita
    cboTipoPago.Enabled = pbHabilita
    LblNumDoc.Enabled = pbHabilita
    TxtMonPag.Enabled = pbHabilita
    If Mid(ActxCta.NroCuenta, 9, 1) = "1" Or Trim(Mid(ActxCta.NroCuenta, 9, 1)) = "" Then
        TxtMonPag.BackColor = vbWhite
        lblITF.BackColor = vbWhite
    Else
        TxtMonPag.BackColor = vbGreen
        lblITF.BackColor = vbGreen
    End If
    If cboFormaPago.ListCount > 0 Then
        cboFormaPago.ListIndex = 0
    End If
    cboTipoPago.ListIndex = 0
    
    If pbHabilita Then
        If TxtMonPag.Enabled And TxtMonPag.Visible Then
            TxtMonPag.SetFocus
        End If
    End If
End Function

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean

Dim rsPers As ADODB.Recordset
Dim rsCredVig As ADODB.Recordset
Dim sAgencia As String
Dim nGastos As Double
Dim nMonPago As Double
Dim nMora As Double
Dim nCuotasMora As Integer
Dim nTotalDeuda As Currency
Dim nInteresDesagio As Double
Dim nMonCalDin As Double
Dim sMensaje As String

Dim nNewSalCap As Double
Dim nNewCPend As Integer
Dim dProxFec As Date
Dim sEstado As String

Dim nCuotaPendiente As Integer
Dim nMoraCalculada As Double
Dim dFechaVencimiento As Date

Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset

    On Error GoTo ErrorCargaDatos

    Set oCredito = New COMNCredito.NCOMCredito
    Call oCredito.CargaDatosPagoCuotas(psCtaCod, gdFecSis, bPrepago, gsCodAge, rsCredVig, sAgencia, nCalendDinamico, bCalenDinamic, bCalenCuotaLibre, _
                                    nMiVivienda, nCalPago, nGastos, nMonPago, nMora, nCuotasMora, nTotalDeuda, nInteresDesagio, _
                                    nMonCalDin, sMensaje, sPersCod, sOperacion, bExoneradaLavado, bRFA, rsPers, bOperacionEfectivo, nMontoLavDinero, nTC, _
                                    nMontoPago, nITF, vnIntPendientePagado, nNewSalCap, nNewCPend, dProxFec, sEstado, nCuotaPendiente, nMoraCalculada, dFechaVencimiento, _
                                    nMontoPag2CuotxVenc, bCuotasVencidas, lnMontoPendienteIntGracia, lnMontIntComp, lnMontGasto, , , , , , bExoneraGastoConvenio)
    
    
    'RIRO20200911 VALIDA LIQUIDACION ***************
    bValidaActualizacionLiq = oCredito.VerificaActualizacionLiquidacion(psCtaCod)
    If Not bValidaActualizacionLiq Then
        MsgBox "El crédito no tiene actualizados sus datos de liquidación, no podrá realizar cancelaciones " & _
        "anticipadas a menos que actualice estos datos. Deberá comunicarse con el área de T.I.", vbExclamation, "Aviso"
    End If
    'END RIRO **************************************
    
    'CTI2 ADD bExoneraGastoConvenio 20190101
    If Not rsCredVig.BOF And Not rsCredVig.EOF Then
        ''**ARLO20180712 ERS042 - 2018
        'Select Case Mid(psCtaCod, 6, 3)
        'Case "515", "516"
        '       MsgBox "Ud. debe realizar el pago de este crédito por la opción de PAGO CUOTA ARRENDAMIENTO FINANCIERO", vbInformation, "Aviso"
        '       CargaDatos = False
        '       Exit Function
        'End Select
        Set objProducto = New COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
        If objProducto.GetResultadoCondicionCatalogo("O0000014", Mid(psCtaCod, 6, 3)) Then
                MsgBox "Ud. debe realizar el pago de este crédito por la opción de PAGO CUOTA ARRENDAMIENTO FINANCIERO", vbInformation, "Aviso"
                CargaDatos = False
                Exit Function
        End If
        ''**ARLO20180712 ERS042 - 2018
        
        If rsCredVig!cTpoCredCod = "853" Or rsCredVig!cTpoCredCod = "854" Then
            MsgBox "Ud. no puede realizar pagos MiVivienda ni Techo propio por esta opción", vbInformation, "Aviso"
            CargaDatos = False
            Exit Function
        End If
        
        lblNomCliente.Caption = PstaNombre(rsCredVig!cPersNombre)
        lblMontoCred.Caption = Format(rsCredVig!nMontoCol, "#,##0.00")
        lblSaldoCap.Caption = Format(rsCredVig!nSaldo, "#,##0.00")
        lblDeudaAct.Caption = Format(nTotalDeuda, "#,##0.00")
        
        lblFecDesemb.Caption = Format(rsCredVig!dFecDesemb, "dd/MM/yyyy")
        lblFecVenc.Caption = dFechaVencimiento
        lblMetLiquid.Caption = Trim(rsCredVig!cMetLiquidacion)
        
        lblDOI.Caption = rsCredVig!nDoi
        lblMoneda.Caption = Trim(rsCredVig!cmoneda)
        lblMontoCuota.Caption = Format(nMonPago, "#,##0.00")
        lblMonto2Cuota.Caption = Format(nMontoPag2CuotxVenc, "#,##0.00")
        
        lnDiasAtraso = Trim(Str(rsCredVig!nDiasAtraso))
        
        nCalendDinamTipo = rsCredVig!nCalendDinamTipo
        
        nCuotasApr = CInt(rsCredVig!nCuotasApr)
        nCuotasPend = nCuotaPendiente
                
        CargaDatos = True

        TxtMonPag.Text = nMonPago
    
        Dim oDInstFinan As COMDPersona.DCOMInstFinac
        Set oDInstFinan = New COMDPersona.DCOMInstFinac
        bInstFinanc = oDInstFinan.VerificaEsInstFinanc(rsCredVig!cPersCod)
        Set oDInstFinan = Nothing
        
        'TxtMonPag.Text = Format(TxtMonPag.Text, "#0.00")
        TxtMonPag.Text = 0
        
        If sMensaje <> "" Then
            MsgBox sMensaje, vbInformation, "Mensaje"
            Exit Function
        End If
        
        bantxtmonpag = True
        
        If bInstFinanc Then nITF = 0
        lblITF.Caption = Format(nITF, "#0.00")
        nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
        If nRedondeoITF > 0 Then
            Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
        End If
        If Trim(Right(cboFormaPago.Text, 10)) = gColocTipoPagoCargoCta Then lblITF.Caption = "0.00"
        lblPagoTotal.Caption = Format(Val(TxtMonPag.Text) + CCur(Me.lblITF.Caption), "#0.00")
        lblNewSaldoCap.Caption = nNewSalCap
        lblNewCuotaPend.Caption = nNewCPend
        If dProxFec <> 0 Then lblProxFecPago.Caption = dProxFec
        
        bantxtmonpag = False
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
    
        bActualizaMontoPago = False
        '-----------------------
    
        If Not rsPers.EOF Then
            sPerscodLav = sPersCod
            sNombreLav = rsPers!Nombre
            sDireccionLav = rsPers!Direccion
            sDocIdLav = rsPers!id & " " & rsPers![ID N°]
        End If
            
        Set lafirma = New frmPersonaFirma
        Set ClsPersona = New COMDPersona.DCOMPersonas
            
        Set Rf = ClsPersona.BuscaCliente(sPersCod, BusquedaCodigo)
        
        If Not Rf.BOF And Not Rf.EOF Then
            If Rf!nPersPersoneria = 1 Then
                Call frmPersonaFirma.Inicio(Trim(sPersCod), Mid(sPersCod, 4, 2), False, True)
            End If
    
            Set Rf = Nothing
        End If
    Else
        CargaDatos = False
        MsgBox "No se pudo encontrar el Credito, o el Credito No esta Vigente", vbInformation, "Aviso"
    End If
    
    Exit Function

ErrorCargaDatos:
    MsgBox Err.Description, vbCritical, "Aviso"

End Function

Private Sub LimpiaPantalla()
    LimpiaControles Me, True
    InicializaCombos Me
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    lblNewSaldoCap.Caption = ""
    lblProxFecPago.Caption = ""
    lblNewCuotaPend.Caption = ""
    bCalenDinamic = False
    nRedondeoITF = 0
    txtCuentaCargo.NroCuenta = ""
    txtCuentaCargo.Visible = False
    bInstFinanc = False
    bValidaActualizacionLiq = False
End Sub

Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gColocTipoPago, , , 2)
    Call Llenar_Combo_con_Recordset(R, cboFormaPago)
    Set R = oCons.RecuperaConstantes(gColocTipoPagoEspecial)
    Call Llenar_Combo_con_Recordset(R, cboTipoPago)
    Set oCons = Nothing
    Exit Sub

ERRORCargaControles:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub Form_Load()
    Call CargaControles
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    bCalenDinamic = False
    CentraSdi Me
    bantxtmonpag = False
    Set oDocRec = New UDocRec
    bExoneraGastoConvenio = True 'CTI2 20181215 ERS075-2018
    bValidaActualizacionLiq = False
End Sub

Private Sub TxtMonPag_Change()
    If Not bantxtmonpag Then
        bActualizaMontoPago = True
        lblNewSaldoCap.Caption = ""
        lblNewCuotaPend.Caption = ""
        lblProxFecPago.Caption = ""
        cmdGrabar.Enabled = False
    End If
End Sub

Private Sub TxtMonPag_GotFocus()
    fEnfoque TxtMonPag
End Sub

Private Sub TxtMonPag_KeyPress(KeyAscii As Integer)
Dim bValorProceso As Boolean
Dim sMensaje As String
Dim nMonIntGra As Double
Dim nNewSalCap As Double
Dim nNewCPend As Integer
Dim dProxFec As Date
Dim sEstado As String

    KeyAscii = NumerosDecimales(TxtMonPag, KeyAscii, 15)
       
    If KeyAscii <> 13 Then Exit Sub
    
    If bActualizaMontoPago = False Then
        bActualizaMontoPago = True
        If cmdGrabar.Enabled Then
            cmdGrabar.SetFocus
        End If
        Exit Sub
    End If
    
    If Not IsNumeric(TxtMonPag.Text) Then
        MsgBox "Ingrese un monto válido", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    'If TxtMonPag.Text < CDbl(lblMontoCuota.Caption) Then
    If nCuotasApr - (nCuotasPend - 1) > 1 And TxtMonPag.Text < CDbl(lblMontoCuota.Caption) Then   'JUEZ 20150625
        MsgBox "El Pago Especial debe ser necesariamente más de una cuota", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'MARG 20180618 Pag. Ant.---------------------------------------------------------
    If CInt(Trim(Right(cboTipoPago.Text, 10))) = gPagoAnticipado Then
        Dim oDecisionPagAnt As COMDCredito.DCOMCredito
        Dim rsDecisionPagAnt As ADODB.Recordset
        Dim cMensajePagAnt As String
        Dim bMuestraMensajePagAnt As Boolean
        Dim bPermitePagAnt As Boolean
        Dim bAplicaValidacionPagAnt As Boolean
        
        Set oDecisionPagAnt = New COMDCredito.DCOMCredito
        Set rsDecisionPagAnt = oDecisionPagAnt.getDecisionPagoAnticipado(ActxCta.NroCuenta, lnDiasAtraso, False, True, False)
        If Not rsDecisionPagAnt.BOF And Not rsDecisionPagAnt.EOF Then
            bAplicaValidacionPagAnt = CBool(rsDecisionPagAnt!bAplicaValidacionPagAnt)
            bPermitePagAnt = CBool(rsDecisionPagAnt!bPermitePagAnt)
            bMuestraMensajePagAnt = CBool(rsDecisionPagAnt!bMuestraMensajePagAnt)
            cMensajePagAnt = rsDecisionPagAnt!cMensajePagAnt
            If bAplicaValidacionPagAnt Then
                If bMuestraMensajePagAnt Then
                    MsgBox cMensajePagAnt, vbInformation, "AVISO"
                End If
                If Not bPermitePagAnt Then
                    rsDecisionPagAnt.Close
                    Set rsDecisionPagAnt = Nothing
                    Exit Sub
                End If
            End If
        End If
        rsDecisionPagAnt.Close
        Set rsDecisionPagAnt = Nothing
    End If
    'END MARG-----------------------------------------------------------------------
    
    If Trim(Right(Me.cboTipoPago.Text, 2)) = gAdelantoCuota Then
        If TxtMonPag.Text < CDbl(lblMonto2Cuota.Caption) Then
            MsgBox "Para realizar Adelanto de Cuota con menos de dos cuotas utilice la opción de PAGO NORMAL", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    If Trim(Right(Me.cboTipoPago.Text, 2)) = gPagoAnticipado Then
        If TxtMonPag.Text > CDbl(lblMonto2Cuota.Caption) Then
            MsgBox "Para realizar Pagos Anticipados mayor a dos cuotas utilice la opción de PAGO NORMAL", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    
    
    bValorProceso = oCredito.ActualizaMontoPago(CDbl(TxtMonPag.Text), CDbl(lblDeudaAct.Caption), ActxCta.NroCuenta, gdFecSis, lblMetLiquid.Caption, vnIntPendiente, vnIntPendientePagado, _
                                        bCalenCuotaLibre, bCalenDinamic, bPrepago, nMontoPago, 0, sMensaje, nITF, _
                                        nInteresDesagio, nNewSalCap, nNewCPend, dProxFec, sEstado, nMonIntGra, , , , nMiVivienda, , lnMontoPendienteIntGracia, 0, lnMontIntComp, lnMontGasto, , , bExoneraGastoConvenio)
    'bExoneraGastoConvenio CTI2
    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
    End If
    
    If bValorProceso = False Then Exit Sub
    
    If sEstado = "CANCELADO" Then
        MsgBox "Para realizar cancelaciones utilice la opción de PAGO NORMAL", vbInformation, "Aviso"
        Exit Sub
    End If
    
    bantxtmonpag = True
    TxtMonPag.Text = Format(TxtMonPag.Text, "#0.00")
    
    If bInstFinanc Then nITF = 0
    lblITF.Caption = Format(nITF, "#0.00")
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
    If nRedondeoITF > 0 Then
        Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
    End If
    If CInt(Trim(Right(cboFormaPago.Text, 10))) = gColocTipoPagoCargoCta Then lblITF.Caption = "0.00"
    lblPagoTotal.Caption = Format(Val(TxtMonPag.Text) + CCur(Me.lblITF.Caption), "#0.00")
    lblNewSaldoCap.Caption = nNewSalCap
    lblNewCuotaPend.Caption = nNewCPend
    If dProxFec <> 0 Then lblProxFecPago.Caption = dProxFec
    
    bantxtmonpag = False
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
End Sub

Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)
Dim nMonto As Double

nMonto = CDbl(TxtMonPag.Text)
poLavDinero.TitPersLavDinero = sPersCod
poLavDinero.OrdPersLavDinero = sPersCod

End Sub

Private Function SeleccionarCtaCargo() As Boolean
Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Dim loCuentas As COMDPersona.UCOMProdPersona
Dim rsCuentas As ADODB.Recordset

    SeleccionarCtaCargo = False

    If Trim(sPersCod) <> "" Then
    Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsCuentas = oDCapGen.GetCuentasPersona(sPersCod, gCapAhorros, True, True, Mid(ActxCta.NroCuenta, 9, 1), , , , True, gPrdCtaTpoIndiv)
    Set oDCapGen = Nothing
    End If
    
    If rsCuentas.RecordCount > 0 Then
        Set loCuentas = New COMDPersona.UCOMProdPersona
        Set loCuentas = frmProdPersona.Inicio(lblNomCliente.Caption, rsCuentas)
        If loCuentas.sCtaCod <> "" Then
            SeleccionarCtaCargo = True
            txtCuentaCargo.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
            txtCuentaCargo.SetFocusCuenta
        End If
        Set loCuentas = Nothing
    Else
        MsgBox "El cliente no tiene cuentas de ahorro activas", vbInformation, "Aviso"
        SeleccionarCtaCargo = False
        Exit Function
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
