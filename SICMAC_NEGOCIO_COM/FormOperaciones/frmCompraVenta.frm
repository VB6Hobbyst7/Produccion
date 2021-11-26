VERSION 5.00
Begin VB.Form frmCompraVenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compra Venta Moneda Extranjera"
   ClientHeight    =   7845
   ClientLeft      =   1875
   ClientTop       =   2205
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTranferecia 
      Caption         =   "Transferencia"
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
      Height          =   2445
      Left            =   60
      TabIndex        =   30
      Top             =   4920
      Width           =   6945
      Begin VB.TextBox txtTransferGlosa 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   855
         MaxLength       =   255
         TabIndex        =   33
         Top             =   1410
         Width           =   5865
      End
      Begin VB.CommandButton cmdTranfer 
         Height          =   350
         Left            =   3840
         Picture         =   "frmCompraVenta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   645
         Width           =   475
      End
      Begin VB.ComboBox cboTransferMoneda 
         Height          =   330
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   255
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "TCV"
         Height          =   285
         Left            =   5355
         TabIndex        =   46
         Top             =   630
         Width           =   390
      End
      Begin VB.Label lblTTCC 
         Caption         =   "TCC"
         Height          =   285
         Left            =   5370
         TabIndex        =   45
         Top             =   270
         Width           =   390
      End
      Begin VB.Label lblTransferGlosa 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   1410
         Width           =   495
      End
      Begin VB.Label lblTransferMoneda 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   45
         TabIndex        =   43
         Top             =   315
         Width           =   585
      End
      Begin VB.Label lblTrasferND 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   855
         TabIndex        =   42
         Top             =   645
         Width           =   2880
      End
      Begin VB.Label lbltransferBcol 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   90
         TabIndex        =   41
         Top             =   1110
         Width           =   555
      End
      Begin VB.Label lbltransferN 
         AutoSize        =   -1  'True
         Caption         =   "Nro Doc :"
         Height          =   195
         Left            =   45
         TabIndex        =   40
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lbltransferBco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   855
         TabIndex        =   39
         Top             =   1020
         Width           =   5865
      End
      Begin VB.Label lblTTCCD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5985
         TabIndex        =   38
         Top             =   255
         Width           =   750
      End
      Begin VB.Label lblTTCVD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5970
         TabIndex        =   37
         Top             =   615
         Width           =   750
      End
      Begin VB.Label lblEtiMonTra 
         AutoSize        =   -1  'True
         Caption         =   "Monto Transacción"
         Height          =   195
         Left            =   840
         TabIndex        =   36
         Top             =   2100
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lblSimTra 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   2400
         TabIndex        =   35
         Top             =   2070
         Width           =   300
      End
      Begin VB.Label lblMonTra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   300
         Left            =   2760
         TabIndex        =   34
         Top             =   2040
         Width           =   1665
      End
   End
   Begin VB.Frame fraFormaPago 
      Height          =   600
      Left            =   60
      TabIndex        =   24
      Top             =   4200
      Width           =   6945
      Begin VB.ComboBox CmbForPag 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   195
         Width           =   1785
      End
      Begin SICMACT.ActXCodCta txtCuentaCargo 
         Height          =   375
         Left            =   3120
         TabIndex        =   26
         Top             =   200
         Visible         =   0   'False
         Width           =   3630
         _extentx        =   6403
         _extenty        =   661
         texto           =   "Cuenta N°:"
         enabledcta      =   -1
         enabledage      =   -1
      End
      Begin VB.Label lblFormaPago 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   255
         Width           =   855
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
         Left            =   4300
         TabIndex        =   28
         Top             =   200
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblNroDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   3105
         TabIndex        =   27
         Top             =   250
         Visible         =   0   'False
         Width           =   1050
      End
   End
   Begin VB.CommandButton CmdRechazar 
      Caption         =   "&Rechazar"
      Height          =   360
      Left            =   1200
      TabIndex        =   23
      Top             =   7440
      Width           =   1185
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   3600
      TabIndex        =   4
      Top             =   7440
      Width           =   1275
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   2400
      TabIndex        =   3
      Top             =   7440
      Width           =   1185
   End
   Begin VB.Frame fraTipoCambio 
      Height          =   1065
      Left            =   60
      TabIndex        =   9
      Top             =   3060
      Width           =   6945
      Begin VB.TextBox txtTpoCambio 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """S/."" #,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   5115
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   620
         Width           =   1140
      End
      Begin VB.CheckBox ChkTCEspecial 
         Caption         =   "Tipo Cambio Especial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtMontoPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   10
         Top             =   600
         Width           =   1695
      End
      Begin SICMACT.EditMoney txtImporte 
         Height          =   375
         Left            =   1665
         TabIndex        =   2
         Top             =   195
         Width           =   1695
         _extentx        =   2990
         _extenty        =   661
         font            =   "frmCompraVenta.frx":030A
         text            =   "0"
         enabled         =   -1
      End
      Begin VB.Label lblTipoCambio 
         Caption         =   "Monto a Cambiar:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   255
         Width           =   1575
      End
      Begin VB.Label lblMonto 
         Caption         =   "Monto a Pagar:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label lblsimbolosoles2 
         AutoSize        =   -1  'True
         Caption         =   "S/"
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
         Left            =   3435
         TabIndex        =   12
         Top             =   690
         Width           =   225
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "$."
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
         Left            =   3420
         TabIndex        =   11
         Top             =   285
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Cambio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3915
         TabIndex        =   18
         Top             =   650
         Width           =   1080
      End
      Begin VB.Label lblTpoCambioDia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   5115
         TabIndex        =   17
         Top             =   620
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   3840
         Top             =   600
         Width           =   2430
      End
   End
   Begin VB.Frame fraPersona 
      Caption         =   "Datos de la Persona"
      Height          =   2745
      Left            =   60
      TabIndex        =   7
      Top             =   315
      Width           =   6945
      Begin VB.CheckBox ChkAut 
         Caption         =   "AUTORIZACION"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin SICMACT.FlexEdit fgDocs 
         Height          =   1230
         Left            =   1515
         TabIndex        =   1
         Top             =   1335
         Width           =   4095
         _extentx        =   7223
         _extenty        =   2170
         cols0           =   4
         highlight       =   2
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "N°-Documento-N° Doc.-Tipo"
         encabezadosanchos=   "450-1500-1800-0"
         font            =   "frmCompraVenta.frx":032E
         font            =   "frmCompraVenta.frx":0356
         font            =   "frmCompraVenta.frx":037E
         font            =   "frmCompraVenta.frx":03A6
         font            =   "frmCompraVenta.frx":03CE
         fontfixed       =   "frmCompraVenta.frx":03F6
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X"
         listacontroles  =   "0-0-0-0"
         encabezadosalineacion=   "C-L-L-L"
         formatosedit    =   "0-0-0-0"
         textarray0      =   "N°"
         lbformatocol    =   -1
         lbpuntero       =   -1
         lbordenacol     =   -1
         colwidth0       =   450
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin SICMACT.TxtBuscar txtBuscaPers 
         Height          =   330
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   1815
         _extentx        =   3201
         _extenty        =   582
         appearance      =   1
         appearance      =   1
         font            =   "frmCompraVenta.frx":041C
         appearance      =   1
         tipobusqueda    =   3
         stitulo         =   ""
         tipobuspers     =   1
      End
      Begin VB.Label lblResultAut 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N° Documentos :"
         Height          =   210
         Left            =   195
         TabIndex        =   8
         Top             =   1335
         Width           =   1200
      End
      Begin VB.Label lblPersDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         TabIndex        =   6
         Top             =   930
         Width           =   5670
      End
      Begin VB.Label lblPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         TabIndex        =   5
         Top             =   585
         Width           =   5670
      End
   End
   Begin VB.Label lblPermiso 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6600
      Width           =   3735
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "COMPRA MONEDA EXTRANJERA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   1410
      TabIndex        =   15
      Top             =   75
      Width           =   4140
   End
End
Attribute VB_Name = "frmCompraVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oPers As COMDPersona.UCOMPersona
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim lsAgencia As String
Dim lbSalir As Boolean
Dim lsDocumento  As String
Dim nPersoneria As COMDConstantes.PersPersoneria
Dim lsMovNroCV As String 'ALPA 20140328*******************
Dim codigoMovimiento As String 'GIPO 04/01/2017

'***************Variabres Agregadas********************
Dim Gtitular As String
Dim GAutNivel As String
Dim GAutMontoFinSol As Double
Dim GAutMontoFinDol As Double
Dim GMontoAprobado As Double, GNroID As Long, GPersCod As String
Dim gbEstado As Boolean
Dim gnNumDec As Integer
Dim nTipoCambioCV As Currency
Public Event Change() 'ALPA 20140328
Public Event KeyPress(KeyAscii As Integer) 'ALPA 20140328

Dim bLimpiarFormulario As Boolean

'******* TORE 20190327*******'
Dim cCodOpeTmp As String
'****************************'
Private nMontoVoucher As Currency 'CTI6 ERS0112020
Dim nMovNroRVD As Long, nMovNroRVDPend As Long 'CTI6 ERS0112020
Dim sNumTarj As String 'CTI6 ERS0112020
Dim pnMoneda As Integer 'CTI6 ERS0112020
Dim psPersCodTitularAhorroCargo As String 'CTI6 ERS0112020
Dim pnITF As Double 'CTI6 ERS0112020
Dim pbEsMismoTitular As Boolean 'CTI6 ERS0112020
Dim pnMontoPagarCargo As Double 'CTI4 ERS0112020
Dim nRedondeoITF As Double ' BRGO 20110914
Dim loVistoElectronico As frmVistoElectronico 'CTI4 ERS0112020
Dim nRespuesta As Integer 'CTI4 ERS0112020
Dim lnTransferSaldo As Currency 'CTI7 OPEv2
Dim fsPersCodTransfer As String 'CTI7 OPEv2
Dim fnMovNroRVD As Long 'CTI7 OPEv2
Dim lnMovNroTransfer As Long 'CTI7 OPEv2

Private Sub ChkAut_Click()
If Me.ChkAut.value = 1 Then
    'Me.txtIdAut.Visible = True
    Me.ChkTCEspecial.Visible = False
    'GIPO ERS069-2016
    cargarResultadoAutorizacion
    
ElseIf (bLimpiarFormulario) Then
   Call LimpiarFormulario
Else
    Me.ChkTCEspecial.Visible = True
End If

End Sub

Private Sub LimpiarFormulario()
     'Me.txtIdAut.Visible = False
    Me.ChkTCEspecial.Visible = True
    txtImporte.Enabled = True
    Me.CmdGuardar.Enabled = False
    Me.txtBuscaPers.Enabled = True
    txtImporte.Text = Format(0, "#,#0.00")
    lblTpoCambioDia.Caption = Format(0, "#,#0.00")
    TxtMontoPagar = Format(0, "#,#0.00")
    lblPersNombre.Caption = ""
    lblPersDireccion.Caption = ""
    txtBuscaPers.Text = ""
    'GIPO ERS069-2016
    lblResultAut.Visible = False
    txtTpoCambio.Text = Format(0, "#,#0.00")
    CmdRechazar.Visible = False
    Me.fgDocs.Clear
    CmbForPag.ListIndex = -1 'CTI6 ERS0112020
    txtCuentaCargo.NroCuenta = "" 'CTI6 ERS0112020
    LblNumDoc.Caption = "" 'CTI6 ERS0112020
    pnMoneda = 0 'CTI6 ERS0112020
    CmbForPag.Enabled = False 'CTI6 ERS0112020
End Sub
'CTI6 ERS0112020
Private Sub EstadoFormaPago(ByVal nFormaPago As Integer)
    LblNumDoc.Caption = ""
    txtCuentaCargo.NroCuenta = ""
    Select Case nFormaPago
        Case -1
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargo.Visible = False
            CmdGuardar.Enabled = True
        Case gCVTipoPagoEfectivo
            txtCuentaCargo.Visible = False
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            lblNroDocumento.Visible = False
            CmdGuardar.Enabled = True
        Case gCVTipoPagoCargoCta
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargo.Visible = True
            txtCuentaCargo.Enabled = True
            txtCuentaCargo.CMAC = gsCodCMAC
            txtCuentaCargo.Prod = Trim(Str(gCapAhorros))
            CmdGuardar.Enabled = False
    End Select
End Sub
Private Sub CmbForPag_Click()
    EstadoFormaPago IIf(CmbForPag.ListIndex = -1, -1, CInt(Trim(Right(IIf(CmbForPag.Text = "", "-1", CmbForPag.Text), 10))))
    If CmbForPag.ListIndex <> -1 Then
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gCVTipoPagoEfectivo Then
     
            Dim oformVou As New frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIF As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
            lnTipMot = 15
            CmdGuardar.Enabled = True
            Me.fraTranferecia.Enabled = False
            Call IniciarControlesFormaPago
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gCVTipoPagoCargoCta Then
            Dim sCuenta As String
            Dim sTempOpeCod As String
            Dim lsOpeTpoPP As String
            sTempOpeCod = ""
            lsOpeTpoPP = gAhoCargoCompra
            If gsOpeCod = COMDConstSistema.gOpeCajeroMECompra Then
                sTempOpeCod = gOpeCajeroMECompra
                lsOpeTpoPP = gAhoCargoCompra
            ElseIf gsOpeCod = COMDConstSistema.gOpeCajeroMEVenta Then
                sTempOpeCod = gOpeCajeroMEVenta
                lsOpeTpoPP = gAhoCargoVenta
            End If
            
            
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(lsOpeTpoPP), sNumTarj, 232, False)
            If Val(Mid(sCuenta, 6, 3)) <> "232" And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If Len(Trim(sCuenta)) = 18 Then
            txtCuentaCargo.Age = Mid(sCuenta, 4, 2)
            txtCuentaCargo.Prod = Mid(sCuenta, 6, 3)
            txtCuentaCargo.Cuenta = Mid(sCuenta, 9, 18)
            Else
            txtCuentaCargo.Prod = "232"
            End If
            Call txtCuentaCargo_KeyPress(13)
            Call IniciarControlesFormaPago
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gCVTipoPagoVoucher Then
        
            'IniciaCombo cboTransferMoneda, gMoneda
            cboTransferMoneda.Enabled = False
            Me.fraTranferecia.Enabled = True
        
            'chkITFEfectivo.Visible = False
            'Me.chkITFEfectivo.value = 1

            Select Case gsOpeCod
                Case COMDConstSistema.gOpeCajeroMECompra
                    cboTransferMoneda.ListIndex = IndiceListaCombo(cboTransferMoneda, 2)
                Case COMDConstSistema.gOpeCajeroMEVenta
                    cboTransferMoneda.ListIndex = IndiceListaCombo(cboTransferMoneda, 1)
             End Select
            
    
        End If
    End If
End Sub

Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gCVTipoPagoBase, , , 3)
  
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPag)
    Set loVistoElectronico = New frmVistoElectronico
    
    IniciaCombo cboTransferMoneda, gMoneda
    Select Case gsOpeCod
       Case COMDConstSistema.gOpeCajeroMECompra
           cboTransferMoneda.ListIndex = IndiceListaCombo(cboTransferMoneda, 1)
       Case COMDConstSistema.gOpeCajeroMEVenta
           cboTransferMoneda.ListIndex = IndiceListaCombo(cboTransferMoneda, 2)
    End Select
    Me.fraTranferecia.Enabled = False
    Exit Sub
ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Function ValidaFormaPago() As Boolean
Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento
    ValidaFormaPago = False
    If CmbForPag.ListIndex = -1 Then
        MsgBox "No se ha seleccionado la forma de pago. Verifique.", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
    
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gCVTipoPagoCargoCta And Len(txtCuentaCargo.NroCuenta) <> 18 Then
        MsgBox "Favor de ingresar la cuenta de ahorros a debitar", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
        
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gCVTipoPagoCargoCta Then
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, CDbl(TxtMontoPagar.Text)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            Exit Function
        End If
    End If
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gCVTipoPagoVoucher Then
        If CDbl(TxtMontoPagar.Text) = 0 Then
            MsgBox "Ingresar monto de la operación.", vbInformation, "¡Aviso!"
            Exit Function
        End If
    
        If CDbl(TxtMontoPagar.Text) <> CDbl(lblMonTra.Caption) Then
            Select Case gsOpeCod
                Case COMDConstSistema.gOpeCajeroMECompra
'                    cboTransferMoneda.ListIndex = IndiceListaCombo(cboTransferMoneda, 2)
                    If CDbl(txtImporte.Text) <> CDbl(lblMonTra.Caption) Then
                        MsgBox "El monto a pagar es diferente al monto del Voucher.", vbInformation, "¡Aviso!"
                        Exit Function
                    End If
                Case COMDConstSistema.gOpeCajeroMEVenta
'                    cboTransferMoneda.ListIndex = IndiceListaCombo(cboTransferMoneda, 1)
                    If CDbl(TxtMontoPagar.Text) <> CDbl(lblMonTra.Caption) Then
                        MsgBox "El monto a pagar es diferente al monto del Voucher.", vbInformation, "¡Aviso!"
                        Exit Function
                    End If
             End Select
        End If
    End If
    
    ValidaFormaPago = True
End Function
'End CTI6 ERS0112020
Private Sub cmdGuardar_Click()
Dim fbPersonaReaOtros As Boolean 'WIOR 20130301
    Dim oCajero As COMNCajaGeneral.NCOMCajero
    Dim lsMovNro As String
    Dim oGen  As COMNContabilidad.NCOMContFunciones
    Dim lsMovNroAut As String
    'Dim ObjTcP As COMDConstSistema.DCOMTCEspPermiso 'LUCV20180222
    'Set ObjTcP = New COMDConstSistema.DCOMTCEspPermiso 'LUCV20180222
        
    Dim OCon As COMConecta.DCOMConecta 'LUCV20180224
    Set OCon = New COMConecta.DCOMConecta 'LUCV20180224
        
    Dim loLavDinero As frmMovLavDinero
    Set loLavDinero = New frmMovLavDinero

    Dim objPersona As COMDPersona.DCOMPersonas 'JACA 20110512
    Set objPersona = New COMDPersona.DCOMPersonas 'JACA 20110512
    
    Set oGen = New COMNContabilidad.NCOMContFunciones
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    Dim lsBoletaCargo  As String 'CTI6 ERS0112020
    Dim MatDatosAho(14) As String 'CTI6 ERS0112020
    Dim lsNombreClienteCargoCta As String 'CTI6 ERS0112020
    Dim lsOpeDesc As String 'CTI7 OPEv2
    'TORE 20190327: Correccion error reportado por correo - registro de codigo de operacion incorrecto
    'Call VerSiClienteActualizoAutorizoSusDatos(txtBuscaPers.Text, gsOpeCod) 'FRHU ERS077-2015 20151204
    Call VerSiClienteActualizoAutorizoSusDatos(txtBuscaPers.Text, cCodOpeTmp)  'FRHU ERS077-2015 20151204
    'END TORE
    
    'CTI6 ERS0112020
        pnMontoPagarCargo = 0#
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gCVTipoPagoCargoCta Then
            If Mid(txtCuentaCargo.NroCuenta, 9, 1) = gMonedaNacional Then
                pnMontoPagarCargo = CDbl(TxtMontoPagar.Text)
            Else
                pnMontoPagarCargo = CDbl(txtImporte.Text)
            End If
            AsignaValorITF
        End If
        If pbEsMismoTitular = False Then
            If CInt(Trim(Right(CmbForPag.Text, 10))) = 4 Then
                 MsgBox "El titular de la cuenta de ahorro no es la misma persona quien solicita la operación", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
     Dim clsCapN As New COMNCaptaGenerales.NCOMCaptaMovimiento
     lsOpeDesc = gsOpeDesc
     If CInt(Trim(Right(CmbForPag.Text, 10))) = gCVTipoPagoCargoCta Then
        If Not clsCapN.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, pnMontoPagarCargo) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            Exit Sub
        End If
        Select Case gsOpeCod
            Case COMDConstSistema.gOpeCajeroMECompra
                lblTitulo = "COMPRA MONEDA EXTRANJERA CARGO A CUENTA"
                lsOpeDesc = gsOpeDesc & "-CARGO A CUENTA"
            Case COMDConstSistema.gOpeCajeroMEVenta
                lblTitulo = "VENTA MONEDA EXTRANJERA CARGO A CUENTA"
                lsOpeDesc = gsOpeDesc & "-CARGO A CUENTA"
        End Select
     End If
     If CInt(Trim(Right(CmbForPag.Text, 10))) = gCVTipoPagoVoucher Then
        If lblTrasferND.Caption = "" Then
            MsgBox "Debe ingresar un numero de transferencia.", vbInformation, "Aviso"
            EnfocaControl cmdTranfer
            Exit Sub
        End If

'        If CDbl(TxtMontoPagar.Text) > CDbl(lblMonTra.Caption) Then
'            MsgBox "El Voucher seleccionado no cubre el valor del monto a pagar.", vbInformation, "¡Aviso!"
'            Exit Sub
'        End If
        
        Select Case gsOpeCod
            Case COMDConstSistema.gOpeCajeroMECompra
                lblTitulo = "COMPRA MONEDA EXTRANJERA VOUCHER"
                lsOpeDesc = gsOpeDesc & "-VOUCHER"
            Case COMDConstSistema.gOpeCajeroMEVenta
                lblTitulo = "VENTA MONEDA EXTRANJERA VOUCHER"
                lsOpeDesc = gsOpeDesc & "-VOUCHER"
        End Select
     End If
     Set clsCapN = Nothing
     
    'END CTI6 ERS0112020
    
    'ALPA 20140328 ********************************************************
    Dim oCred  As COMNCredito.NCOMNivelAprobacion
    Set oCred = New COMNCredito.NCOMNivelAprobacion
    Dim lnEstado As Integer
    Dim nContadorVistos As Integer
    If Len(Trim(lsMovNroCV)) > 0 Then
        nContadorVistos = oCred.ObtenerCantidadAprobacionMovCompraVenta(lsMovNroCV, lnEstado)
        If lnEstado = 3 Then
            MsgBox "El permiso de nivel de aprobación fue denegada", vbInformation, "Aviso"
            Call ActualizarNivelAprovacion
            lblTpoCambioDia.Caption = Format(nTipoCambioCV, "#,#0.0000")
            txtTpoCambio.Text = lblTpoCambioDia.Caption
            TxtMontoPagar.Text = Format(Val(txtImporte.value) * Val(txtTpoCambio.Text), "#,#0.00")
            Exit Sub
        End If
        If nContadorVistos > 0 Then
            MsgBox "Aun falta mas vistos para la operación, favor coordinar para continuar", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    '**********************************************************************
    If Not ValidaFormaPago Then Exit Sub 'CTI6 ERS0112020
    
    If ValidaInterfaz = False Then Exit Sub

    'Graba Autorizacion para Tipo de Cambio Especial
    If Me.ChkAut.value = 0 Then
        If Me.ChkTCEspecial.value = 1 Then
         If MsgBox("Desea grabar la Solicitud de Tipo de Cambio Especial", vbYesNo + vbQuestion, "Aviso") = vbYes Then
             lsMovNroAut = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            'Set ObjTcP = New COMDConstSistema.DCOMTCEspPermiso 'LUCV20180222
            'Call ObjTcP.InsertaTipoCambio(lsMovNroAut, gsOpeCod, CCur(Me.lblTpoCambioDia.Caption), gdFecSis, CCur(Me.txtImporte.Text), Me.txtBuscaPers.Text) 'LUCV20180222
            
            '->***** LUCV20180224 *****
            Dim dFechaReg As String
            Dim sql As String
            Call OCon.AbreConexion
            
            dFechaReg = Format(gdFecSis & " " & OCon.GetHoraServer, "MM/DD/YYYY hh:mm:ss AMPM")
            
            'TORE 20190327: Correccion error reportado por correo - registro de codigo de operacion incorrecto
            'sql = "EXEC stp_ins_CapAutorizacionTC '" & lsMovNroAut & "','" & gsOpeCod & "'," & CCur(Me.lblTpoCambioDia.Caption) & ",'" & dFechaReg & "'," & CCur(Me.txtImporte.Text) & ",'" & Me.txtBuscaPers.Text & "'," & 0
            sql = "EXEC stp_ins_CapAutorizacionTC '" & lsMovNroAut & "','" & cCodOpeTmp & "'," & CCur(Me.lblTpoCambioDia.Caption) & ",'" & dFechaReg & "'," & CCur(Me.txtImporte.Text) & ",'" & Me.txtBuscaPers.Text & "'," & 0
            'END TORE
            
            OCon.Ejecutar sql
            Call OCon.CierraConexion
            Set OCon = Nothing
            '<-***** Fin LUCV20180224 *****
            
            MsgBox "Se grabo la Solicitud de Tipo de Cambio Especial", vbInformation, "AVISO"
            'Set oImp = Nothing
            txtBuscaPers = ""
            lblPersDireccion = ""
            lblPersNombre = ""
            fgDocs.Clear
            fgDocs.FormaCabecera
            fgDocs.Rows = 2
            txtImporte = 0
            TxtMontoPagar = "0.00"
            txtBuscaPers.SetFocus
            Me.ChkTCEspecial.value = 0
'            Me.fraTranferecia.Enabled = False
'            cboTransferMoneda.ListIndex = -1
'            lblTrasferND.Caption = ""
'            lbltransferBco.Caption = ""
'            txtTransferGlosa.Text = ""
'            lblMonTra.Caption = ""
            'Set ObjTcP = Nothing 'LUCV20180222
            Exit Sub
         End If
           
        End If
    End If
  'JACA 20110512 *****VERIFICA SI LAS PERSONAS CUENTAN CON OCUPACION E INGRESO PROMEDIO
        Dim rsPersVerifica As Recordset
        Dim i As Integer
        Set rsPersVerifica = New Recordset
        
            Set rsPersVerifica = objPersona.ObtenerDatosPersona(txtBuscaPers.Text)
            If rsPersVerifica!nPersIngresoProm = 0 Or rsPersVerifica!cActiGiro1 = "" Then
                If MsgBox("Necesita Registrar la Ocupacion e Ingreso Promedio de: " + Me.lblPersNombre, vbYesNo) = vbYes Then
                    'frmPersona.Inicio txtBuscaPers.Text, PersonaActualiza
                    frmPersOcupIngreProm.Inicio txtBuscaPers.Text, lblPersNombre, rsPersVerifica!cActiGiro1, rsPersVerifica!nPersIngresoProm
                End If
            End If
        
    'JACA END***************************************************************************

    If MsgBox("Desea grabar la Operación de Compra/Venta??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        Dim sPersLavDinero As String, sReaPersLavDinero As String, sBenPersLavDinero As String
        Dim nMontoLavDinero As Double

        Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
        Dim nMonto As Double
        Dim bLavDinero As Boolean
        Dim oVarPublicas As New COMFunciones.FCOMVarPublicas
        Dim lsBoleta As String
        Dim nFicSal As Integer
        bLavDinero = False
        nMonto = txtImporte.value
        'Realiza la Validación para el Lavado de Dinero
        Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
        If Not clsExo.EsPersonaExoneradaLavadoDinero(txtBuscaPers) Then
            Set clsExo = Nothing
            sPersLavDinero = ""
            Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
            nMontoLavDinero = clsLav.GetCapParametro(COMDConstantes.gMonOpeLavDineroME)
            Set clsLav = Nothing
            
            If nMonto >= nMontoLavDinero Then
                loLavDinero.TitPersLavDinero = Trim(txtBuscaPers.Text)
                'By Capi 1402208
                 'Call IniciaLavDinero(loLavDinero)
                 'ALPA 20081009*************************************************
                 'sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, "", Mid(Me.Caption, 15), True, "", , , , , 2)
                 
                 'TORE 20190327: Correccion error reportado por correo - registro de codigo de operacion incorrecto
                 'sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, "", Mid(Me.Caption, 15), True, "", , , , , 2, , gnTipoREU, gnMontoAcumulado, gsOrigen, , gsOpeCod) 'WIOR 20131106 AGREGO gsOpeCod
                 sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, "", Mid(Me.Caption, 15), True, "", , , , , 2, , gnTipoREU, gnMontoAcumulado, gsOrigen, , cCodOpeTmp)  'WIOR 20131106 AGREGO gsOpeCod
                 'END TORE
                 
                 bLavDinero = True
                 If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                'End
'                If nPersoneria = COMDConstantes.gPersonaNat Then
'
'                    sReaPersLavDinero = gVarPublicas.gReaPersLavDinero
'                    sBenPersLavDinero = gVarPublicas.gBenPersLavDinero
'                    'By Capi 28012008
'                    sOrdPersLavDinero = gVarPublicas.gOrdPersLavDinero
'                    VisPersLavDinero = gVarPublicas.gVisPersLavDinero
'                    If sPersLavDinero = "" Then Exit Sub
'
'                    sPersLavDinero = Trim(txtBuscaPers.Text)
'                Else
'                    sPersLavDinero = frmMovLavDinero.Inicia(Trim(txtBuscaPers.Text), Trim(lblPersNombre), Trim(lblPersDireccion), Trim(fgDocs.TextMatrix(1, 2)), True, True, nMonto, " ", gsOpeDesc, , , , , , , 2)
'
'                    sReaPersLavDinero = gVarPublicas.gReaPersLavDinero
'                    sBenPersLavDinero = gVarPublicas.gBenPersLavDinero
'                    'By Capi 28012008
'                    sOrdPersLavDinero = gVarPublicas.gOrdPersLavDinero
'                    VisPersLavDinero = gVarPublicas.gVisPersLavDinero
'                    If sPersLavDinero = "" Then Exit Sub
'                End If
'                'If sPersLavDinero = "" Then Exit Sub
'                bLavDinero = True
'

            End If
        Else
            Set clsExo = Nothing
        End If
        'WIOR 20130301 *********************************************************
        fbPersonaReaOtros = False
        If loLavDinero.OrdPersLavDinero = "Exit" Then
        
            'TORE 20190327: Correccion error reportado por correo - registro de codigo de operacion incorrecto
            'frmPersRealizaOpeGeneral.Inicia lblTitulo.Caption, gsOpeCod
            frmPersRealizaOpeGeneral.Inicia lblTitulo.Caption, cCodOpeTmp
            'END TORE
            
            fbPersonaReaOtros = frmPersRealizaOpeGeneral.PersRegistrar
            
            If Not fbPersonaReaOtros Then
                MsgBox "Se va a proceder a Anular la Operación", vbInformation, "Aviso"
                CmdGuardar.Enabled = True
                Exit Sub
            End If
            'CTI6 ERS00112020
            Dim lnTipoPagoTemporal As Integer
            If (CInt(Trim(Right(CmbForPag.Text, 10))) = gCVTipoPagoEfectivo) Then
                lnTipoPagoTemporal = gCVTipoPagoEfectivo
            Else
                 lnTipoPagoTemporal = 2
            End If
            
          
            If frmPersRealizaOpeGeneral.OpeEfectivo <> lnTipoPagoTemporal Then
                MsgBox "No coincide la información del tipo de pago con lo seleccionado en los datos de la transacción", vbInformation, "Aviso"
                CmdGuardar.Enabled = True

                Exit Sub
            End If
            'END CTI6
        End If
        'WIOR FIN ***************************************************************
        lsMovNro = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
        'If oCajero.GrabaCompraVenta(gsFormatoFecha, lsMovNro, gsOpeCod, gsOpeDesc, nMonto, CDbl(lblTpoCambioDia), txtBuscaPers, bLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro) = 0 Then
        
        'TORE 20190327: Correccion error reportado por correo - registro de codigo de operacion incorrecto
        'If oCajero.GrabaCompraVenta(gdFecSis, gsFormatoFecha, lsMovNro, gsOpeCod, gsOpeDesc, nMonto, CDbl(lblTpoCambioDia), txtBuscaPers, bLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro) = 0 Then 'TI ERS 002-2017 JUCS 05/10/2017
        If oCajero.GrabaCompraVenta(gdFecSis, gsFormatoFecha, lsMovNro, cCodOpeTmp, gsOpeDesc, nMonto, CDbl(lblTpoCambioDia), txtBuscaPers, bLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, CInt(Trim(Right(CmbForPag.Text, 10))), txtCuentaCargo.NroCuenta, MatDatosAho, gITF.gbITFAplica, pnITF, pnMontoPagarCargo, lnMovNroTransfer, CInt(Right(IIf(Trim(Me.cboTransferMoneda.Text) = "", "001", Me.cboTransferMoneda.Text), 3)), fnMovNroRVD, CCur(IIf(Trim(lblMonTra) = "", 0, lblMonTra))) = 0 Then 'TI ERS 002-2017 JUCS 05/10/2017
        'END TORE
        
            'GIPO 2017-01-05
             Dim oCapAut As COMDConstSistema.DCOMTCEspPermiso
             Set oCapAut = New COMDConstSistema.DCOMTCEspPermiso
             Call oCapAut.AceptarTipoCambioEspecialCliente(codigoMovimiento)
             Call oCapAut.OpeTipoCambioEspecialCliente(codigoMovimiento, gnMovNro) 'APRI20180201 INC20180105004
            'ALPA 20140608****************
            Call ActualizarNivelAprovacion
            'ALPA 20081010
            If bLavDinero Then
             'Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
              Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
            End If
            'JACA 20110510***********************************************************
                
                'Dim objPersona As COMDPersona.DCOMPersonas
                Dim rsPersOcu As Recordset
                Dim nAcumulado As Currency
                Dim nMontoPersOcupacion As Currency
                 
                Set rsPersOcu = New Recordset
                'Set objPersona = New COMDPersona.DCOMPersonas
                                
                Set rsPersOcu = objPersona.ObtenerDatosPersona(txtBuscaPers.Text)
                nAcumulado = objPersona.ObtenerPersAcumuladoMontoOpe(lblTpoCambioDia, Mid(Format(gdFecSis, "yyyymmdd"), 1, 6), rsPersOcu!cperscod)
                nMontoPersOcupacion = objPersona.ObtenerParamPersAgeOcupacionMonto(Mid(rsPersOcu!cperscod, 4, 2), CInt(Mid(rsPersOcu!cPersCIIU, 2, 2)))
            
                If nAcumulado >= nMontoPersOcupacion Then
                    If Not objPersona.ObtenerPersonaAgeOcupDatos_Verificar(txtBuscaPers.Text, gdFecSis) Then
                        objPersona.insertarPersonaAgeOcupacionDatos gnMovNro, txtBuscaPers.Text, TxtMontoPagar.Text, nAcumulado, gdFecSis, lsMovNro
                    End If
                End If
               
        
            'JACA END*****************************************************************
            
            
            Dim oImp As COMNContabilidad.NCOMContImprimir
            Dim lsTexto As String
            Dim lbReimp As Boolean
            Set oImp = New COMNContabilidad.NCOMContImprimir
            
            
'            lsBoleta = oImp.ImprimeBoletaCompraVenta(lblTitulo, "", lblPersNombre, lblPersDireccion, lsDocumento, _
'                        CCur(lblTpoCambioDia), gsOpeCod, CCur(txtImporte), CCur(txtMontoPagar), gsNomAge, lsMovNro, sLpt, gsCodCMAC, gsNomCmac, gbImpTMU)
                                    
            lsBoleta = oImp.ImprimeBoletaCompraVenta(lblTitulo, "", lblPersNombre, lblPersDireccion, lsDocumento, _
                        CCur(lblTpoCambioDia), cCodOpeTmp, CCur(txtImporte), CCur(TxtMontoPagar), gsNomAge, lsMovNro, sLpt, gsCodCMAC, gsNomCmac, gbImpTMU)
                        
            'CTI6 ERS0112020
            If CInt(Trim(Right(CmbForPag.Text, 10))) = gCVTipoPagoCargoCta Then
                lsBoletaCargo = oImp.ImprimeBoletaAhorro("RETIRO AHORROS", IIf(Mid(txtCuentaCargo.NroCuenta, 9, 1) = gMonedaNacional, "COMPRA ME", "VENTA ME"), "", CStr(CDbl(pnMontoPagarCargo) + pnITF), lsNombreClienteCargoCta, txtCuentaCargo.NroCuenta, "", CDbl(MatDatosAho(10)), CDbl(MatDatosAho(3)), "", 1, CDbl(MatDatosAho(11)), , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, gsCodCMAC, 0, , , gbImpTMU)
            End If
            'END CTI6
            
            lbReimp = True
            Do While lbReimp
                 If Trim(lsBoleta) <> "" Then
                    nFicSal = FreeFile
                    Open sLpt For Output As nFicSal
                        Print #nFicSal, lsBoleta & lsBoletaCargo
                        Print #nFicSal, ""
                    Close #nFicSal
                 End If
               
            
                If MsgBox("¿Desea Reimprimir boleta de Operación?", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                    lbReimp = False
                End If
            Loop
                 
            Set oImp = Nothing

            txtBuscaPers = ""
            lblPersDireccion = ""
            lblPersNombre = ""
            fgDocs.Clear
            fgDocs.FormaCabecera
            fgDocs.Rows = 2
            txtImporte = 0
            TxtMontoPagar = "0.00"
            'txtBuscaPers.SetFocus
            txtImporte.Enabled = True
            txtTpoCambio.Text = lblTpoCambioDia.Caption
            lblResultAut.Visible = False
            Me.ChkAut.value = 0
            'INICIO JHCU ENCUESTA 16-10-2019
            Encuestas gsCodUser, gsCodAge, "ERS0292019", cCodOpeTmp
            'FIN
            CmbForPag.ListIndex = -1 'CTI6 ERS0112020
            txtCuentaCargo.NroCuenta = "" 'CTI6 ERS0112020
            LblNumDoc.Caption = "" 'CTI6 ERS0112020
            pnMoneda = 0 'CTI6 ERS0112020
            'CTI7 OPEv2*******************************
            Me.fraTranferecia.Enabled = False
            cboTransferMoneda.ListIndex = -1
            lblTrasferND.Caption = ""
            lbltransferBco.Caption = ""
            txtTransferGlosa.Text = ""
            lblMonTra.Caption = ""
            '*****************************************
            
        End If
        'WIOR 20130301 ************************************************************
        If fbPersonaReaOtros And gnMovNro > 0 Then
            frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, ""
            fbPersonaReaOtros = False
        End If
        'WIOR FIN *****************************************************************
        'CTI4 ERS0112020
        If CInt(Trim(Right(IIf(Trim(CmbForPag.Text) = "", "001", CmbForPag.Text), 10))) = gColocTipoPagoCargoCta Then
            Dim oMovOperacion As COMDMov.DCOMMov
            Dim nMovNroOperacion As Long
            Dim rsCli As New ADODB.Recordset
            Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
            Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
            Set oMovOperacion = New COMDMov.DCOMMov
            nMovNroOperacion = oMovOperacion.GetnMovNro(lsMovNro)

            loVistoElectronico.RegistraVistoElectronico nMovNroOperacion, , gsCodUser, nMovNroOperacion

            If nRespuesta = 2 Then
                Set rsCli = clsCli.GetPersonaCuenta(txtCuentaCargo.NroCuenta, gCapRelPersTitular)
                oSolicitud.ActualizarCapAutSinTarjetaVisto_nMovNro gsCodUser, gsCodAge, txtCuentaCargo.NroCuenta, rsCli!cperscod, nMovNroOperacion, CStr(gAhoCargoCtaCancelaPigno)
            End If
            Set oMovOperacion = Nothing
            nRespuesta = 0
        End If
        'CTI4 end
        gVarPublicas.LimpiaVarLavDinero
        
        Set oGen = Nothing
        Set oCajero = Nothing
        Set loLavDinero = Nothing
        
    
    End If

End Sub
Private Sub IniciarControlesFormaPago()
        Me.fraTranferecia.Enabled = False
        cboTransferMoneda.ListIndex = -1
        lblTrasferND.Caption = ""
        lbltransferBco.Caption = ""
        txtTransferGlosa.Text = ""
        lblMonTra.Caption = ""
End Sub
Function ValidaInterfaz() As Boolean
ValidaInterfaz = True
If txtBuscaPers = "" Then
    MsgBox "Persona no Ingresada", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtBuscaPers.SetFocus
    Exit Function
End If
If Val(txtImporte) = 0 Then
    MsgBox "Importe de Operación no Ingresado", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtImporte.SetFocus
    Exit Function
End If
If Val(TxtMontoPagar) = 0 Then
    MsgBox "Monto a Pagar no válido para Operación", vbInformation, "Aviso"
    ValidaInterfaz = False
    Exit Function
End If
End Function
Private Sub ActualizarNivelAprovacion()
'ALPA 20140328**********************************************************
Dim oDN  As COMDCredito.DCOMNivelAprobacion
Set oDN = New COMDCredito.DCOMNivelAprobacion
If Len(Trim(lsMovNroCV)) > 0 Then
    oDN.ActualizarAprobacionMovCompraVenta (lsMovNroCV)
End If
lsMovNroCV = "" 'ALPA 20140328**************
lblPermiso.Caption = "" 'ALPA 20140328**************
'***********************************************************************
End Sub
'GIPO
Private Sub cmdRechazar_Click()
      If MsgBox("Está seguro de proceder con el Rechazo del Tipo de Cambio Especial propuesto", vbYesNo + vbQuestion, "Aviso") = vbYes Then
          Dim oCapAut As COMDConstSistema.DCOMTCEspPermiso
          Set oCapAut = New COMDConstSistema.DCOMTCEspPermiso
          Call oCapAut.RechazarTipoCambioEspecialCliente(codigoMovimiento)
          Call LimpiarFormulario
          ' AXCodCta.Enabled = True 'CTI6 ERS0112020
          CmbForPag.Enabled = False 'CTI6 ERS0112020
         ' AXCodCta.SetFocus 'CTI6 ERS0112020
      End If
End Sub

Private Sub cmdsalir_Click()
If Len(Trim(lsMovNroCV)) > 0 Then
    If MsgBox("Desea eliminar el movimiento de tipo de cambio", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        Call ActualizarNivelAprovacion 'ALPA 20140328**************
    End If
End If
lsMovNroCV = ""
lblPermiso.Caption = ""

Unload Me
End Sub

Private Sub cmdTranfer_Click()
    Dim lsGlosa As String
    Dim lsDoc As String
    Dim lsInstit As String
    Dim oForm As frmCapRegVouDepBus
    Dim lnTipMot As Integer
    Dim lsDetalle As String

    On Error GoTo ErrTransfer
    If cboTransferMoneda.Text = "" Then
        MsgBox "Debe escoger la moneda de la transferencia.", vbInformation, "Aviso"
        If cboTransferMoneda.Visible And cboTransferMoneda.Enabled Then cboTransferMoneda.SetFocus
        Exit Sub
    End If
       
    If gsOpeCod = COMDConstSistema.gOpeCajeroMECompra Then
       lnTipMot = 14
    ElseIf gsOpeCod = COMDConstSistema.gOpeCajeroMEVenta Then
       lnTipMot = 15
    End If
  
    
    fnMovNroRVD = 0
    Set oForm = New frmCapRegVouDepBus
    'sinReglas 'EJVG20140408
    SetDatosTransferencia "", "", "", 0, -1, "" 'Limpiamos datos y variables globales
    oForm.iniciarFormulario Trim(Right(cboTransferMoneda, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, lsDetalle
    If fnMovNroRVD = 0 Then
        Exit Sub
    End If
    SetDatosTransferencia lsGlosa, lsInstit, lsDoc, lnTransferSaldo, lnMovNroTransfer, lsDetalle
'    grdCliente.row = 1
'    grdCliente.Col = 3
'    grdCliente_OnEnterTextBuscar grdCliente.TextMatrix(1, 1), 1, 1, False
'    Set oForm = Nothing
    Exit Sub
ErrTransfer:
    MsgBox "Ha sucedido un error al cargar los datos de la Transferencia", vbCritical, "Aviso"
End Sub
Private Sub SetDatosTransferencia(ByVal psGlosa As String, ByVal psInstit As String, ByVal psDoc As String, ByVal pnTransferSaldo As Currency, ByVal pnMovNroTransfer As Long, ByVal psDetalle As String)
    Dim oPersona As New DCOMPersona
    Dim rsPersona As New ADODB.Recordset
    Dim row As Integer
    
    txtTransferGlosa.Text = psGlosa
    lbltransferBco.Caption = psInstit
    lblTrasferND.Caption = psDoc
    
'    LimpiaFlex grdCliente
'    If psDetalle <> "" Then
'        Set rsPersona = oPersona.RecuperaPersonaxCapRegVouDep(psDetalle)
'        Do While Not rsPersona.EOF
'            grdCliente.AdicionaFila
'            row = grdCliente.row
'            grdCliente.TextMatrix(row, 1) = rsPersona!cperscod
'            grdCliente.TextMatrix(row, 2) = rsPersona!cPersNombre
'            grdCliente.TextMatrix(row, 4) = rsPersona!nPersPersoneria
'            rsPersona.MoveNext
'        Loop
'    End If
    
   
'    If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
'        txtMonto.Text = Format(pnTransferSaldo, "#,##0.00")
'    Else
'        txtMonto.Text = Format(pnTransferSaldo * CCur(lblTTCCD.Caption), "#,##0.00")
'    End If
'
'
'    If txtCuenta.Prod = "234" Then
'        vnMontoDOC = CDbl(txtMonto.Text)
'        lblTotTran.Caption = vnMontoDOC
'    End If
'
'    txtMonto_Change
'
    If pnMovNroTransfer <> -1 Then
        txtTransferGlosa.SetFocus
    End If
'
    txtTransferGlosa.Locked = True
'    txtMonto.Enabled = False
    lblMonTra = Format(pnTransferSaldo, "#,##0.00")
'
'    Set rsPersona = Nothing
'    Set oPersona = Nothing
End Sub
Private Sub Form_Activate()
If lbSalir Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim ObjTc As COMDConstSistema.DCOMTipoCambioEsp

Dim oOpe As COMNCajaGeneral.NCOMCajaCtaIF
Dim oTipCambio As COMDConstSistema.NCOMTipoCambio

Set ObjTc = New COMDConstSistema.DCOMTipoCambioEsp

Set oTipCambio = New COMDConstSistema.NCOMTipoCambio
gnTipCambioC = oTipCambio.EmiteTipoCambio(gdFecSis, TCCompra)
gnTipCambioV = oTipCambio.EmiteTipoCambio(gdFecSis, TCVenta)
'CTI3 ERS0032020*************************************************
Me.lblTTCCD.Caption = Format(gnTipCambioC, "#,#0.0000")
Me.lblTTCVD.Caption = Format(gnTipCambioV, "#,#0.0000")
'****************************************************************
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Set oTipCambio = Nothing

fgDocs.Enabled = False 'para evitar el error al hacer doble click en el flexgrid
bLimpiarFormulario = True 'GIPO ERS0692016

Set oOpe = New COMNCajaGeneral.NCOMCajaCtaIF

Me.Caption = gsOpeDesc
txtImporte.psSoles False
lbSalir = False


If gdFecSis <> Format(ObjTc.GetFechaIng(), "DD/MM/YYYY") Then
    MsgBox "Tipo de Cambio no ha sido Ingresado. Por favor Ingrese Tipo Cambio del Día", vbInformation, "Aviso"
    lbSalir = True
    Set ObjTc = Nothing
    Exit Sub
End If
Set ObjTc = Nothing
'lblTpoCambioDia.Caption = Format(IIf(gsOpeCod = gOpeCajeroMECompra, gnTipCambioC, gnTipCambioV), "#,#0.0000")

lblTpoCambioDia.Caption = Format(0, "#,#0.0000")

'If Val(lblTpoCambioDia) = 0 Then
'    MsgBox "Tipo de Cambio no ha sido Ingresado. Por favor Ingrese Tipo Cambio del Día", vbInformation, "Aviso"
'    lbSalir = True
'    Exit Sub
'End If
cCodOpeTmp = ""
Select Case gsOpeCod
    Case COMDConstSistema.gOpeCajeroMECompra
        lblTitulo = "COMPRA MONEDA EXTRANJERA NORMAL"
        Me.lblMonto = "Monto a Pagar"
        cCodOpeTmp = gsOpeCod 'TORE 20190327: Correccion error reportado por correo - registro de codigo de operacion incorrecto
    Case COMDConstSistema.gOpeCajeroMEVenta
        lblTitulo = "VENTA MONEDA EXTRANJERA NORMAL"
        Me.lblMonto = "Monto a Recibir"
        cCodOpeTmp = gsOpeCod 'TORE 20190327: Correccion error reportado por correo - registro de codigo de operacion incorrecto
End Select
TxtMontoPagar = "0.00"

'falta definir el objeto area agencia con que va a trabajar
'TORE 20190327: Correccion error reportado por correo - registro de codigo de operacion incorrecto
'lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", , gsCodAge, ObjCMACAgenciaArea)
'lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H", , gsCodAge, ObjCMACAgenciaArea)

lsCtaDebe = oOpe.EmiteOpeCta(cCodOpeTmp, "D", , gsCodAge, ObjCMACAgenciaArea)
lsCtaHaber = oOpe.EmiteOpeCta(cCodOpeTmp, "H", , gsCodAge, ObjCMACAgenciaArea)
'END TORE



lsMovNroCV = "" 'ALPA 20140328**************
lblPermiso.Caption = "" 'ALPA 20140328**************

'GIPO
Me.ChkAut.Enabled = False
Me.lblResultAut.Visible = False
CmdRechazar.Visible = False
cboTransferMoneda.ListIndex = -1
Call CargaControles 'CTI6 ERS0112020
Call IniciarControlesFormaPago 'CTI7 OPEv2

End Sub



'CTI3 ERS0032020***********************************************************************
Private Sub cboTransferMoneda_Click()
    If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
        'lblSimTra.Caption = "S/."
        lblSimTra.Caption = gcPEN_SIMBOLO 'APRI20191022 SUGERENCIA CALIDAD
        lblMonTra.BackColor = &HC0FFFF
    Else
        lblSimTra.Caption = "$"
        lblMonTra.BackColor = &HC0FFC0
    End If
End Sub
Private Sub cboTransferMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdTranfer.SetFocus
    End If
End Sub

Private Sub IniciaCombo(ByRef cboConst As ComboBox, nCapConst As ConstanteCabecera)
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(nCapConst)
    Set clsGen = Nothing
    Do While Not rsConst.EOF
        cboConst.AddItem rsConst("cDescripcion") & space(100) & rsConst("nConsValor")
        rsConst.MoveNext
    Loop
    cboConst.ListIndex = 0
End Sub
'*************************************************************
Private Sub txtBuscaPers_EmiteDatos()

lblPersNombre = txtBuscaPers.psDescripcion
lblPersDireccion = txtBuscaPers.sPersDireccion
nPersoneria = txtBuscaPers.PersPersoneria
fgDocs.Clear
fgDocs.FormaCabecera
fgDocs.Rows = 2
lsDocumento = ""
If txtBuscaPers <> "" Then
    lsDocumento = txtBuscaPers.sPersNroDoc
    Set fgDocs.Recordset = txtBuscaPers.rsDocPers
    Me.ChkAut.value = 0 'GIPO
    Me.ChkAut.Enabled = True 'GIPO
End If

If gsCodPersUser = txtBuscaPers Then
    txtBuscaPers.Text = ""
    MsgBox "No se puede hacer esta Operacion con su persona", vbInformation, "AVISO"
    Exit Sub
End If
'ALPA 20140606**********************************
    Dim lsMovNroFecha As String
    Dim oGen  As COMNContabilidad.NCOMContFunciones
    Set oGen = New COMNContabilidad.NCOMContFunciones

    lsMovNroFecha = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

    Dim oCont As COMNContabilidad.NCOMContFunciones
    Set oCont = New COMNContabilidad.NCOMContFunciones
    Dim oNiv As COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Dim rsMov As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    Set rsMov = New ADODB.Recordset
    
    Dim lnFila As Integer
    Set oNiv = New COMDCredito.DCOMNivelAprobacion
    
    'TORE 20190327: Correccion error reportado por correo - registro de codigo de operacion incorrecto
    'Set rs = oNiv.ObtenerAprobacionMovCompraVentaPendientexCliente(Mid(lsMovNroFecha, 1, 8), txtBuscaPers.Text, gsOpeCod)
    Set rs = oNiv.ObtenerAprobacionMovCompraVentaPendientexCliente(Mid(lsMovNroFecha, 1, 8), txtBuscaPers.Text, cCodOpeTmp)
    'EN TORE
    
    

If Not (rs.EOF Or rs.BOF) Then
    'TORE 20190327: Correccion error reportado por correo - registro de codigo de operacion incorrecto
    'lsMovNroCV = frmNivelesAprobacionCVxPendientesCliente.InicioRegistroNiveles(txtBuscaPers.Text, gsOpeCod)
    lsMovNroCV = frmNivelesAprobacionCVxPendientesCliente.InicioRegistroNiveles(txtBuscaPers.Text, cCodOpeTmp)
    'END TORE
    Set rsMov = oNiv.AprobacionMovCompraVentaxMovimiento(lsMovNroCV)
    If Not (rsMov.EOF Or rsMov.BOF) Then
        lblPermiso = rsMov!cMovNro
        'rsMov!cOpecod
        'rsMov!cNivelCod
        ChkTCEspecial.value = IIf(rsMov!nTipoEspecial = 0, 0, 1)
        txtImporte.value = rsMov!nMonto
        txtTpoCambio.Text = rsMov!nTipoCambioSolici
        TxtMontoPagar.Text = Format(Val(txtImporte.value) * Val(txtTpoCambio.Text), "#,#0.00")
        lblTpoCambioDia.Caption = rsMov!nTipoCambioNormal
        Me.CmdGuardar.Enabled = True
        CmdGuardar.SetFocus
    End If
End If
'***********************************************
fgDocs.RowHeight(-1) = 230
fgDocs.RowHeight(0) = 280
If Me.txtImporte.Enabled = True Then txtImporte.SetFocus

'GIPO ERS069-2016
    'AXCodCta.Enabled = False 'CTI6 ERS0112020
    CmbForPag.Enabled = True 'CTI6 ERS0112020
    CmbForPag.ListIndex = IndiceListaCombo(CmbForPag, 1) 'CTI6 ERS0112020
End Sub

Private Sub cargarResultadoAutorizacion()
Dim nmoneda As Integer
Dim oCapAut As COMDConstSistema.DCOMTCEspPermiso
Set oCapAut = New COMDConstSistema.DCOMTCEspPermiso
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset

 nmoneda = COMDConstantes.gMonedaExtranjera
   
 'Set rs = oCapAut.GetTc_AUT(gdFecSis, Me.txtIdAut.Text, gsCodUser, gsOpeCod, Trim(txtBuscaPers.Text))
 Set rs = oCapAut.GetTc_Autorizacion(gdFecSis, gsCodUser, gsOpeCod, Trim(txtBuscaPers.Text))
 
 If Not (rs.EOF And rs.BOF) Then
 
   If (rs.RecordCount = 1) Then
    
    Call cargarDatosAutorizacion(rs)

    Else 'En caso de que tengan más de 1 evaluacion del mismo cliente
    
    Dim cMovNroSelected As String
    cMovNroSelected = frmEvaluacionAutorizacionME.Inicia(rs)
    If (cMovNroSelected <> "") Then
        Set Rs2 = oCapAut.GetTc_AutorizacionByNroMov(cMovNroSelected)
        Call cargarDatosAutorizacion(Rs2)
    Else
        bLimpiarFormulario = False
        Me.ChkAut.value = 0
    End If
   End If
   
   
 Else
   'MsgBox "No Existe este Id de Autorización para esta operación." & vbCrLf & "Consulte las Operaciones Pendientes.", vbOKOnly + vbInformation, "Atención"
   MsgBox "No se ha solicitado ningún tipo de cambio especial para ésta persona.", vbOKOnly + vbInformation, "Atención"
   bLimpiarFormulario = False
   Me.ChkAut.value = 0
   Set oCapAut = Nothing
   Set rs = Nothing
   Exit Sub
 End If

End Sub
'GIPO ERS0692016
Private Sub cargarDatosAutorizacion(ByVal rs As ADODB.Recordset)
Dim oCapAut As COMDConstSistema.DCOMTCEspPermiso
Set oCapAut = New COMDConstSistema.DCOMTCEspPermiso
  If (rs!nEstado = 0) Then
      Me.lblResultAut.Visible = True
      Me.lblResultAut.Caption = "AUT. PENDIENTE"
      Me.lblResultAut.ForeColor = &H0&
      txtTpoCambio = Format(rs!nTCReg, "#,##0.0000")
      CmdRechazar.Visible = False
   ElseIf (rs!nEstado = 1) Then
       Me.lblResultAut.Visible = True
       Me.lblResultAut.Caption = "AUT. ANULADA"
       Me.lblResultAut.ForeColor = &HC0&
       txtTpoCambio = Format(rs!nTCReg, "#,##0.0000")
       CmdRechazar.Visible = False
   Else 'rs!nEstado = 2
       Me.lblResultAut.Visible = True
       Me.lblResultAut.Caption = "AUT. ENVIADA"
       Me.lblResultAut.ForeColor = &H8000&
       txtTpoCambio = Format(rs!nTCAprob, "#,##0.0000")
       CmdRechazar.Visible = True
   End If
   
   codigoMovimiento = rs!cMovNro
   txtBuscaPers.Text = rs!cperscod
   lblPersNombre = rs!cPersNombre
   lblPersDireccion = rs!cPersDireccDomicilio
   nPersoneria = rs!nPersPersoneria
   fgDocs.Clear
   fgDocs.FormaCabecera
   fgDocs.Rows = 2
   lsDocumento = ""
   If txtBuscaPers <> "" Then
       lsDocumento = IIf(rs!DNI = "", rs!Ruc, rs!DNI)
       Set fgDocs.Recordset = oCapAut.GetDoc_Persona(rs!cperscod)
   End If
   fgDocs.RowHeight(-1) = 230
   fgDocs.RowHeight(0) = 280
   Me.ChkTCEspecial.Visible = 0
   Me.ChkTCEspecial.Visible = False
   txtImporte.Text = Format(rs!nMontoReg, "#,#0.00")
   lblTpoCambioDia.Caption = Format(rs!nTCAprob, "#,#0.000")
   TxtMontoPagar = Format(Val(txtImporte.value) * Val(lblTpoCambioDia), "#,#0.00")
   
  
   txtImporte.Enabled = False
   Me.CmdGuardar.Enabled = True
   Me.txtBuscaPers.Enabled = False
End Sub

Private Sub txtIdAut_KeyPress(KeyAscii As Integer)
Dim nmoneda As Integer

Dim oCapAut As COMDConstSistema.DCOMTCEspPermiso
Set oCapAut = New COMDConstSistema.DCOMTCEspPermiso
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

 nmoneda = COMDConstantes.gMonedaExtranjera
   
If KeyAscii = 13 Then 'And Trim(txtIdAut.Text) <> "" Then
      'Set rs = oCapAut.GetTc_AUT(gdFecSis, Me.txtIdAut.Text, gsCodUser, gsOpeCod, Trim(txtBuscaPers.Text))
      Set rs = oCapAut.GetTc_Autorizacion(gdFecSis, gsCodUser, gsOpeCod, Trim(txtBuscaPers.Text))
      
      If Not (rs.EOF And rs.BOF) Then
        
        If (rs!nEstado = 0) Then
           Me.lblResultAut.Caption = "AUT. PENDIENTE"
        ElseIf (rs!nEstado = 1) Then
            Me.lblResultAut.Caption = "AUT. DENEGADA"
        Else
            Me.lblResultAut.Caption = "AUT. APROBADA"
        End If
        
        txtBuscaPers.Text = rs!cperscod
        lblPersNombre = rs!cPersNombre
        lblPersDireccion = rs!cPersDireccDomicilio
        nPersoneria = rs!nPersPersoneria
        fgDocs.Clear
        fgDocs.FormaCabecera
        fgDocs.Rows = 2
        lsDocumento = ""
        If txtBuscaPers <> "" Then
            lsDocumento = IIf(rs!DNI = "", rs!Ruc, rs!DNI)
            Set fgDocs.Recordset = oCapAut.GetDoc_Persona(rs!cperscod)
        End If
        fgDocs.RowHeight(-1) = 230
        fgDocs.RowHeight(0) = 280
        Me.ChkTCEspecial.Visible = 0
        Me.ChkTCEspecial.Visible = False
        'sql = sql & " cMovNro,nCodAut,cCodOpe,nTCReg,dFechaReg,nMontoReg,cPersCod,"
        'sql = sql & " nEstado , cCodUserAprob, dFechaAprob, cMovNroAprob, nTCAprob,"
        txtImporte.Text = Format(rs!nMontoReg, "#,#0.00")
        lblTpoCambioDia.Caption = Format(rs!nTCAprob, "#,#0.000")
        TxtMontoPagar = Format(Val(txtImporte.value) * Val(lblTpoCambioDia), "#,#0.00")
        txtImporte.Enabled = False
        Me.CmdGuardar.Enabled = True
        Me.txtBuscaPers.Enabled = False
      Else
        'MsgBox "No Existe este Id de Autorización para esta operación." & vbCrLf & "Consulte las Operaciones Pendientes.", vbOKOnly + vbInformation, "Atención"
        MsgBox "No se ha solicitado ningún tipo de cambio para ésta persona." & vbCrLf & "Consulte las Operaciones Pendientes.", vbOKOnly + vbInformation, "Atención"
        
        Set oCapAut = Nothing
        Set rs = Nothing
        Exit Sub
      End If
      
'      Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
'        Set Rs = oCapAut.SAA(Left(gsOpeCod, 5) & "1", Vusuario, "", Gtitular, CInt(nMoneda), CLng(txtIdAut.Text))
'      Set oCapAut = Nothing
'     If Rs.State = 1 Then
'       If Rs.RecordCount > 0 Then
'        txtImporte.Text = Rs!nMontoAprobado
'       Else
'          MsgBox "No Existe este Id de Autorización para esta operación." & vbCrLf & "Consulte las Operaciones Pendientes.", vbOKOnly + vbInformation, "Atención"
'          txtIdAut.Text = ""
'       End If
'     End If
     
     
 End If
 
 If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 13 Or KeyAscii = 8) Then
      KeyAscii = 0
 End If
End Sub
'CTI6 ERS0112020

Private Sub txtCuentaCargo_KeyPress(KeyAscii As Integer)
    Dim sMsg As String
    If KeyAscii = 13 Then sMsg = ValidaCuentaACargo(txtCuentaCargo.NroCuenta)
    If Len(sMsg) > 0 Then
        MsgBox sMsg, vbInformation, "¡Aviso!"
        txtCuentaCargo.SetFocus
        Exit Sub
    End If
    Dim sMoneda As String
    pnMoneda = 0
    If gsOpeCod = COMDConstSistema.gOpeCajeroMECompra Then
        pnMoneda = gMonedaNacional
    ElseIf gsOpeCod = COMDConstSistema.gOpeCajeroMEVenta Then
        pnMoneda = gMonedaExtranjera
    End If
    If Len(txtCuentaCargo.NroCuenta) = 18 Then
        If CInt(Mid(txtCuentaCargo.NroCuenta, 9, 1)) <> pnMoneda Then
            MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la operación de compra/venta.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
    End If
    ObtieneDatosCuenta txtCuentaCargo.NroCuenta
End Sub
Private Function ValidaCuentaACargo(ByVal psCuenta As String) As String
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMsg As String
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    sMsg = clsCap.ValidaCuentaOperacion(psCuenta)
    ValidaCuentaACargo = sMsg
End Function

Private Sub ObtieneDatosCuenta(ByVal psCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsCta As ADODB.Recordset
    Dim rsV As ADODB.Recordset
    Dim lnTpoPrograma As Integer
    Dim lsTieneTarj As String
    Dim lbVistoVal As Boolean
    Dim lsOpeAhorrCompraVentaCargoCtaAhorro As String 'CTI06
    lsOpeAhorrCompraVentaCargoCtaAhorro = "" 'CTI06
    If pnMoneda = gMonedaNacional Then
        lsOpeAhorrCompraVentaCargoCtaAhorro = gAhoCargoCompra
    ElseIf pnMoneda = gMonedaExtranjera Then
        lsOpeAhorrCompraVentaCargoCtaAhorro = gAhoCargoVenta
    End If

    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsV = New ADODB.Recordset
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(psCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
        If sNumTarj = "" Then
            If Mid(psCuenta, 6, 3) = "232" And lnTpoPrograma <> 1 Then
                Set rsV = clsMant.ValidaTarjetizacion(psCuenta, lsTieneTarj)
                If rsV.RecordCount > 0 Then
                    Dim tipoCta As Integer
                    tipoCta = rsCta("nPrdCtaTpo")
                    If tipoCta = 0 Or tipoCta = 2 Then
                        Dim rsCli As New ADODB.Recordset
                        Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
                        Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
                        Dim bExitoSol As Integer
                        Set rsCli = clsCli.GetPersonaCuenta(psCuenta, gCapRelPersTitular)
                        psPersCodTitularAhorroCargo = rsCli!cperscod ' CTI6
                        nRespuesta = oSolicitud.SolicitarVistoAtencionSinTarjeta(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(lsOpeAhorrCompraVentaCargoCtaAhorro))
                    
                        If nRespuesta = 1 Then '1:Tiene Visto de atencion sin tarjeta pendiente de autorizar
                             MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                             Unload Me
                        End If
                        If nRespuesta = 2 Then '2:Tiene visto aceptado
                            MsgBox "La solicitud de atención sin tarjeta fue Aprobada, proceda con la atención", vbInformation, "Aviso"
                        End If
                        If nRespuesta = 3 Then '3:Tiene visto rechazado
                           If MsgBox("La solicitud de atención sin tarjeta fue RECHAZADA. ¿Desea realizar una nueva solicitud?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                Call oSolicitud.RegistrarVistoDeUsuario(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(lsOpeAhorrCompraVentaCargoCtaAhorro))
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        If nRespuesta = 4 Then '4:Se permite registrar la solicitud
                            Dim mensaje As String
                            If lsTieneTarj = "SI" Then
                                mensaje = "El Cliente posee tarjeta. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                            Else
                                mensaje = "El Cliente NO posee tarjeta activa. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                            End If
                        
                            If MsgBox(mensaje, vbInformation + vbYesNo, "Aviso") = vbYes Then
                        
                                bExitoSol = frmCapMotivoAutorizacion.Inicio(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(lsOpeAhorrCompraVentaCargoCtaAhorro))
                                If bExitoSol > 0 Then
                                    MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                                End If
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        lbVistoVal = loVistoElectronico.Inicio(5, lsOpeAhorrCompraVentaCargoCtaAhorro)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema"
                            Exit Sub
                        End If
                    End If
                ElseIf lsTieneTarj = "NO" And rsV.RecordCount > 0 Then
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers 065-2017
                        lbVistoVal = loVistoElectronico.Inicio(5, lsOpeAhorrCompraVentaCargoCtaAhorro)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema" 'ADD BY MARG ERS065-2017
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            End If
        Else
            If Mid(psCuenta, 6, 3) = "232" And lnTpoPrograma <> 1 Then
                   Set rsV = clsMant.ValidaTarjetizacion(psCuenta, lsTieneTarj)
                   If rsV.RecordCount > 0 Then
                       Dim tipoCta2 As Integer
                       tipoCta2 = rsCta("nPrdCtaTpo")
                       If tipoCta2 = 0 Or tipoCta = 2 Then
                           Dim rsCli2 As New ADODB.Recordset
                           Dim clsCli2 As New COMNCaptaGenerales.NCOMCaptaGenerales
                           Dim oSolicitud2 As New COMDCaptaGenerales.DCOMCaptaGenerales
                           Dim bExitoSol2 As Integer
                           Dim nRespuesta2 As Integer
                           Set rsCli2 = clsCli2.GetPersonaCuenta(psCuenta, gCapRelPersTitular)
                           psPersCodTitularAhorroCargo = rsCli2!cperscod ' CTI6
                       End If
                   End If
            End If
        End If
        txtCuentaCargo.Enabled = False
        AsignaValorITF
        CmdGuardar.Enabled = True
        CmdGuardar.SetFocus
    End If
End Sub

Private Sub AsignaValorITF()
    If psPersCodTitularAhorroCargo = txtBuscaPers Then
        pbEsMismoTitular = True
    Else
        pbEsMismoTitular = False
    End If
    pnITF = 0#
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gCVTipoPagoCargoCta Then
        If gITF.gbITFAplica Then
            pnITF = Format(gITF.fgITFCalculaImpuesto(pnMontoPagarCargo), "#0.00")
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(pnITF))
            If nRedondeoITF > 0 Then
                  pnITF = Format(CCur(pnITF) - nRedondeoITF, "#,##0.00")
            End If
         End If
     End If
End Sub
'End CTI6
Private Sub txtImporte_GotFocus()
With txtImporte
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub


Private Sub txtImporte_KeyPress(KeyAscii As Integer)
Dim rs As ADODB.Recordset
Dim ObjTc As COMDConstSistema.DCOMTipoCambioEsp
Dim oDN  As COMDCredito.DCOMNivelAprobacion
Set oDN = New COMDCredito.DCOMNivelAprobacion

If KeyAscii = 13 Then
    If txtImporte.Text = "" Then
        MsgBox "Ingrese el monto a Cambiar", vbInformation, "AVISO"
        txtImporte.SetFocus
        Exit Sub
    End If
    
    If txtImporte.Text = 0 Then
        MsgBox "Ingrese el monto a Cambiar", vbInformation, "AVISO"
        txtImporte.SetFocus
        Exit Sub
    End If
    
    
    Set ObjTc = New COMDConstSistema.DCOMTipoCambioEsp
    Set rs = New ADODB.Recordset
    Set rs = ObjTc.GetTipoCambioCV(CCur(Me.txtImporte.Text))
    If Not (rs.EOF And rs.BOF) Then
        Do Until rs.EOF
            If Val(txtImporte.value) <= Val(rs!nHasta) Then
                'TORE 20190327: Correccion error reportado por correo - registro de codigo de operacion incorrecto
                'lblTpoCambioDia.Caption = Format(IIf(gsOpeCod = gOpeCajeroMECompra, rs!nCompra, rs!nVenta), "#,#0.0000")
                lblTpoCambioDia.Caption = Format(IIf(cCodOpeTmp = gOpeCajeroMECompra, rs!nCompra, rs!nVenta), "#,#0.0000")
                'END TORE
                txtTpoCambio.Text = Format(IIf(lblTpoCambioDia.Caption = 0#, txtTpoCambio.Text, lblTpoCambioDia.Caption), "#,#0.0000")
                nTipoCambioCV = CDbl(lblTpoCambioDia.Caption)
                'ALPA20140327******************************************************************************
                'TxtMontoPagar.Text = Format(Val(txtImporte.value) * Val(lblTpoCambioDia), "#,#0.00")
                TxtMontoPagar.Text = Format(Val(txtImporte.value) * Val(txtTpoCambio.Text), "#,#0.00")
                If Len(Trim(lsMovNroCV)) > 0 Then
                    oDN.ActualizarAprobacionMovCompraVenta (lsMovNroCV)
                    lsMovNroCV = ""
                    lblPermiso.Caption = ""
                End If
                '******************************************************************************************
                Exit Do
            End If
            rs.MoveNext
        Loop
    Else
        MsgBox "No se ha definido el Tipo de Cambio", vbCritical, "AVISO"
        Set rs = Nothing
        Set ObjTc = Nothing
        Exit Sub
    End If
    
    TxtMontoPagar = Format(Val(txtImporte.value) * Val(lblTpoCambioDia), "#,#0.00")
    Me.CmdGuardar.Enabled = True
    CmdGuardar.SetFocus
End If
End Sub

Private Function Cargousu(ByVal NomUser As String) As String
 Dim rs As New ADODB.Recordset
 Dim oCons As COMDConstSistema.DCOMUAcceso
 Set oCons = New COMDConstSistema.DCOMUAcceso
 
 Set rs = oCons.Cargousu(NomUser)
  If Not (rs.EOF And rs.BOF) Then
    Cargousu = rs(0)
  End If
 Set rs = Nothing
 'rs.Close
 Set oCons = Nothing
End Function
'ALPA 20140327***********************************************************************
Private Sub txtTpoCambio_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If (IsNumeric(Me.txtTpoCambio.Text)) Then
      Call ValidaInsertaAprobacionCV
   Else
      MsgBox "Formato incorrecto!", vbInformation
      txtTpoCambio.Text = Format(lblTpoCambioDia.Caption, "#,#0.0000")
      Exit Sub
   End If
End If

End Sub
Private Sub ValidaInsertaAprobacionCV()
    Dim oDN  As COMDCredito.DCOMNivelAprobacion
    Set oDN = New COMDCredito.DCOMNivelAprobacion
    Dim oGen  As COMNContabilidad.NCOMContFunciones
    Set oGen = New COMNContabilidad.NCOMContFunciones
    If Len(Trim(lsMovNroCV)) > 0 Then
        oDN.ActualizarAprobacionMovCompraVenta (lsMovNroCV)
    End If
    Set oDN = Nothing
    lsMovNroCV = ""
    lblPermiso.Caption = ""
    If CDbl(txtTpoCambio.Text) <> CDbl(lblTpoCambioDia.Caption) Then
        Set oDN = New COMDCredito.DCOMNivelAprobacion
        
        Dim lnActivarCambioenTipodeCambio As Integer
        Dim lnNivel As String
        
        Dim oCred  As COMNCredito.NCOMNivelAprobacion
        Set oCred = New COMNCredito.NCOMNivelAprobacion
        
        lnActivarCambioenTipodeCambio = oCred.ObtenerNivelesAprobacionCompraVentaxMonto(txtImporte.Text, CDbl(txtTpoCambio.Text) - CDbl(lblTpoCambioDia.Caption), IIf(gsOpeCod = COMDConstSistema.gOpeCajeroMECompra, 1, 2), lnNivel, IIf(ChkTCEspecial.value = 1, 1, 0))
        If lnActivarCambioenTipodeCambio = 1 Then
            MsgBox "Se generará movimiento de tipo de cambio...esperar la aprobación al nivel respectivo", vbInformation, "SICMACM Operaciones"
        Else
            MsgBox "El cambio realizado no tiene nivel de aprobacion, favor registrar otro tipo de cambio", vbInformation, "SICMACM Operaciones"
            txtTpoCambio.Text = Format(lblTpoCambioDia.Caption, "#,#0.0000")
            Exit Sub
        End If
        lsMovNroCV = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        TxtMontoPagar.Text = Format(Val(txtImporte.value) * Val(txtTpoCambio.Text), "#,#0.00")
        lblPermiso.Caption = "Mov: " & lsMovNroCV
        'inserta movimiento en la tabla AprobacionMovCompraVenta
        Call oDN.AprobacionMovCompraVenta(lsMovNroCV, gsOpeCod, lnNivel, CDbl(txtImporte.value), CDbl(lblTpoCambioDia.Caption), CDbl(txtTpoCambio.Text), IIf(ChkTCEspecial.value = 1, 1, 0), txtBuscaPers.Text)
    Else
        TxtMontoPagar.Text = Format(Val(txtImporte.value) * Val(txtTpoCambio.Text), "#,#0.00")
    End If
End Sub

Private Function TienePunto(psCadena As String) As Boolean
If InStr(1, psCadena, ".", vbTextCompare) > 0 Then
    TienePunto = True
Else
    TienePunto = False
End If
End Function

Private Function NumDecimal(psCadena As String) As Integer
Dim lnPos As Integer
lnPos = InStr(1, psCadena, ".", vbTextCompare)
If lnPos > 0 Then
    NumDecimal = Len(psCadena) - lnPos
Else
    NumDecimal = 0
End If
End Function

Private Sub txtTpoCambio_Change()
txtTpoCambio.SelStart = Len(txtTpoCambio)
gnNumDec = NumDecimal(txtTpoCambio)
If gbEstado And txtTpoCambio <> "" Then
    Select Case gnNumDec
        Case 0
                txtTpoCambio = Format(txtTpoCambio, "#,##0")
        Case 1
                txtTpoCambio = Format(txtTpoCambio, "#,##0.0")
        Case 2
                txtTpoCambio = Format(txtTpoCambio, "#,##0.00")
        Case 3
                txtTpoCambio = Format(txtTpoCambio, "#,##0.000")
        Case Else
                txtTpoCambio = Format(txtTpoCambio, "#,##0.0000")
    End Select
End If
If txtTpoCambio = "" Then
    txtTpoCambio = 0
End If
gbEstado = False
RaiseEvent Change
End Sub
'****************************************************************************************
