VERSION 5.00
Begin VB.Form frmCCETransfInterBanca 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12495
   Icon            =   "frmCCETransfInterBanca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkIncluyeComis 
      Caption         =   "Monto Incluye Comisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   28
      Top             =   3590
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame fraOrdenanteCuenta 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   12255
      Begin VB.TextBox txtOrdenanteCCIOrigen 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   4920
         MaxLength       =   20
         TabIndex        =   12
         Top             =   190
         Width           =   3495
      End
      Begin SICMACT.ActXCodCta txtOrdenanteCuenta 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   150
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
         CMAC            =   "109"
      End
      Begin VB.Label lblOrdenanteCCIOrigen 
         Caption         =   "CCI Origen:"
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
         Left            =   3840
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkOrdenanteCargoCuenta 
      Caption         =   "Con Cargo a Cuenta"
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
      Height          =   435
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.Frame FraBeneficiario 
      Caption         =   "Beneficiario"
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
      Height          =   720
      Left            =   120
      TabIndex        =   21
      Top             =   2680
      Width           =   12255
      Begin VB.TextBox txtBeneficiarioTpoDoc 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   6720
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtBeneficiarioNroDoc 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   9000
         MaxLength       =   20
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin SICMACT.TxtBuscar txtBeneficiario 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         TipoBusqueda    =   3
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin VB.Label lblBeneficiarioNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "Tpo. Doc:"
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
         Left            =   5740
         TabIndex        =   24
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Nro. Doc:"
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
         Left            =   8160
         TabIndex        =   26
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   10320
      TabIndex        =   43
      Top             =   4920
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   11400
      TabIndex        =   44
      Top             =   4920
      Width           =   1050
   End
   Begin VB.Frame fraMontos 
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
      Height          =   1095
      Left            =   120
      TabIndex        =   29
      Top             =   3680
      Width           =   12255
      Begin VB.TextBox txtITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   8160
         MaxLength       =   20
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtMontoTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   10500
         MaxLength       =   20
         TabIndex        =   41
         Text            =   "0.00"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtMontoATransferir 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   6360
         MaxLength       =   20
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtMontoComision 
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
         Left            =   4560
         MaxLength       =   20
         TabIndex        =   35
         Text            =   "0.00"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtMonto 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2760
         MaxLength       =   15
         TabIndex        =   33
         Text            =   "0.00"
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmCCETransfInterBanca.frx":030A
         Left            =   240
         List            =   "frmCCETransfInterBanca.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblITF 
         Caption         =   "ITF:"
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
         Left            =   8160
         TabIndex        =   38
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblMontoTotal 
         Caption         =   "TOTAL:"
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
         Left            =   10500
         TabIndex        =   40
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblMontoTotalATransf 
         Caption         =   "A Transferir:"
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
         Left            =   6360
         TabIndex        =   36
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblMontoComision 
         Caption         =   "Comisión:"
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
         Left            =   4560
         TabIndex        =   34
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblMonto 
         Caption         =   "Monto:"
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
         Left            =   2760
         TabIndex        =   32
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblMoneda 
         Caption         =   "Moneda:"
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
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame FraOrdenante 
      Caption         =   "Ordenante"
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
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   1860
      Width           =   12255
      Begin VB.CheckBox chkOrdenanteMismoTit 
         Caption         =   "Mismo Titular de Cuentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10680
         TabIndex        =   20
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox txtOrdenanteNroDoc 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   9000
         MaxLength       =   25
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtOrdenanteTpoDoc 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   6720
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin SICMACT.TxtBuscar txtOrdenante 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         TipoBusqueda    =   3
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin VB.Label lblOrdenanteNroDoc 
         Caption         =   "Nro. Doc:"
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
         Left            =   8160
         TabIndex        =   18
         Top             =   255
         Width           =   855
      End
      Begin VB.Label lblOrdenanteTpoDoc 
         Caption         =   "Tpo. Doc:"
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
         Left            =   5740
         TabIndex        =   16
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label lblOrdenanteNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame FraDestino 
      Caption         =   "Destino"
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12255
      Begin VB.TextBox txtNroTarjCred 
         Height          =   285
         Left            =   9360
         MaxLength       =   16
         TabIndex        =   7
         Top             =   300
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtDestinoCCI 
         Height          =   285
         Left            =   9360
         MaxLength       =   20
         TabIndex        =   5
         Top             =   300
         Width           =   2775
      End
      Begin SICMACT.TxtBuscar txtBancoDestinoCod 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   300
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Enabled         =   0   'False
         Appearance      =   0
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin VB.Label lblNroTarjCred 
         Caption         =   "Nº Tarj:"
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
         Left            =   8520
         TabIndex        =   6
         Top             =   315
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblBancoDestinoNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3105
         TabIndex        =   3
         Top             =   300
         Width           =   5055
      End
      Begin VB.Label lblCCIDestino 
         Caption         =   "CCI Destino:"
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
         Left            =   8280
         TabIndex        =   4
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label lblBancoDestino 
         Caption         =   "Banco Destino:"
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
         TabIndex        =   1
         Top             =   320
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   345
      Left            =   9240
      TabIndex        =   42
      Top             =   4920
      Width           =   1050
   End
End
Attribute VB_Name = "frmCCETransfInterBanca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'** Nombre : frmCCETransfInterBanca
'** Descripción : Para el Registro de Transferencias Interbancarias, Proyecto: Implementacion del Servicio de Compensación Electrónica Diferido de Instrumentos Compensables CCE
'** Creación : PASI, 20160613
'**********************************************************************
Option Explicit
Dim fsOpeCod As String
Dim oCCE As COMNCajaGeneral.NCOMCCE
Dim oPersonaOrd As UPersona_Cli
Dim oPersonaBenef As UPersona_Cli
Private clsprevio As New previo.clsprevio
Dim sMovNroAut As String
Dim psOpeTransf As String
Private Type TarifaCampo
    sTariCod As String
    sTariCriApl As String
    sComisSigno As String
    nCCEComisMonto As Currency
    nCMACMComisMonto As Currency
End Type
Dim tTarifaCampo As TarifaCampo
Dim fnLimMaxTrMN As Currency
Dim fnLimMaxTrME As Currency
Dim fnLimMinTrMN As Currency
Dim fnLimMinTrME As Currency
Public Sub Inicio(ByVal psOpeCod As String)
    fsOpeCod = psOpeCod
    Me.Show 1
End Sub
Private Sub cboMoneda_Click()
    txtMonto.SetFocus
    CalculaComision
End Sub
Private Sub chkIncluyeComis_Click()
    CalculaComision
End Sub
Private Sub CalculaComision()
    Dim rs As ADODB.Recordset
    
    txtMontoComision.Text = "0.00"
    txtMontoATransferir.Text = "0.00"
    txtMontoTotal.Text = "0.00"
    txtITF.Text = "0.00"
    
    If cboMoneda.ListIndex = -1 Then Exit Sub
    If nVal(IIf(Len(txtMonto.Text) = 0, 0, IIf(txtMonto.Text = ".", 0, txtMonto.Text))) = 0 Then Exit Sub
    If Len(txtBancoDestinoCod.Text) = 0 Then Exit Sub
    
    Set rs = oCCE.CCE_ObtieneComisionTransf(psOpeTransf, Trim((txtOrdenanteCuenta.Age)), txtBancoDestinoCod.Text, Mid(Trim(txtDestinoCCI.Text), 4, 3), txtMonto.Text, Right(cboMoneda, 1))
    If Not (rs.EOF And rs.BOF) Then
        tTarifaCampo.sTariCod = rs!cTariCod
        tTarifaCampo.sTariCriApl = rs!cTariCriApl
        tTarifaCampo.sComisSigno = rs!cComisSigno
        tTarifaCampo.nCCEComisMonto = rs!nCCEComisMonto
        tTarifaCampo.nCMACMComisMonto = rs!nCMACMComisMonto
        
        txtMontoComision.Text = Format(tTarifaCampo.nCCEComisMonto + tTarifaCampo.nCMACMComisMonto, "#,##0.00")
        txtMontoATransferir.Text = Format(IIf(chkIncluyeComis = 1, nVal(txtMonto.Text) - nVal(txtMontoComision.Text), nVal(txtMonto.Text)), "#,##0.00")
        If ((chkOrdenanteCargoCuenta.value = 1 And txtOrdenanteCuenta.Prod = "232") Or (chkOrdenanteCargoCuenta.value = 0)) And chkOrdenanteMismoTit.value = 0 Then 'SE AGREGO chkOrdenanteMismoTit VALUE 1 VAPA20170410
            txtITF.Text = Format(oCCE.DameMontoITF(nVal(txtMonto.Text)), "#,##0.00")
        End If
        txtMontoTotal.Text = Format(IIf(chkIncluyeComis = 1, nVal(txtMonto.Text), nVal(txtMonto.Text) + nVal(txtMontoComision.Text)) + nVal(txtITF.Text), "#,##0.00")
    Else
        MsgBox "No se pudo realizar el cálculo de la comisión. Si el problema persiste comuniquese con el Dpto. de TI", vbInformation, "¡Aviso!"
        Exit Sub
    End If
End Sub
Private Sub chkOrdenanteCargoCuenta_Click()
    Dim rsDatosPers As New ADODB.Recordset
    Dim rsCtaCCI As New ADODB.Recordset
    Dim sCuenta As String
    fraOrdenanteCuenta.Enabled = IIf(chkOrdenanteCargoCuenta.value = 0, False, True)
    ReiniciaCuenta
    If chkOrdenanteCargoCuenta.value = 1 Then txtOrdenanteCuenta.SetFocus
    txtOrdenante.Enabled = IIf(chkOrdenanteCargoCuenta.value = 0, True, False)
End Sub
'vapa 20170707
Private Sub chkOrdenanteMismoTit_Click()
CalculaComision
End Sub

Private Sub cmdCancelar_Click()
    LimpiarDatos
    ReiniciaCuenta
End Sub
Private Sub ReiniciaCuenta()
    sMovNroAut = ""
    txtOrdenanteCuenta.CMAC = "109"
    txtOrdenanteCuenta.Age = ""
    txtOrdenanteCuenta.Prod = ""
    txtOrdenanteCuenta.Cuenta = ""
    txtOrdenanteCCIOrigen.Text = ""
    txtOrdenante.Enabled = False
    txtOrdenante.Text = ""
    lblOrdenanteNombre.Caption = ""
    txtOrdenanteTpoDoc.Text = ""
    txtOrdenanteNroDoc.Text = ""
    txtOrdenanteCuenta.Enabled = True
    txtOrdenanteCuenta.EnabledCMAC = False
    txtOrdenanteCuenta.EnabledAge = True
    txtOrdenanteCuenta.EnabledCta = True
    txtOrdenanteCuenta.EnabledProd = True
    cboMoneda.Enabled = True
    cboMoneda.ListIndex = -1
    txtMonto.Text = "0.00"
End Sub
Private Sub LimpiarDatos()
    Set oPersonaOrd = New UPersona_Cli
    Set oPersonaBenef = New UPersona_Cli
    sMovNroAut = ""
    txtBancoDestinoCod = ""
    lblBancoDestinoNombre.Caption = ""
    txtDestinoCCI.Text = ""
    txtNroTarjCred.Text = ""
    'chkOrdenanteCargoCuenta.value = 0
    If FraBeneficiario.Visible Then
        txtBeneficiario.Text = ""
        lblBeneficiarioNombre.Caption = ""
        txtBeneficiarioTpoDoc.Text = ""
        txtBeneficiarioNroDoc.Text = ""
    End If
    cboMoneda.ListIndex = -1
    txtMonto.Text = "0.00"
    txtMontoComision.Text = "0.00"
    txtMontoATransferir.Text = "0.00"
    txtMontoTotal.Text = "0.00"
    chkOrdenanteMismoTit.value = 0
    chkIncluyeComis.value = 0
    tTarifaCampo.sTariCod = ""
    tTarifaCampo.sTariCriApl = ""
    tTarifaCampo.sComisSigno = ""
    tTarifaCampo.nCCEComisMonto = 0
    tTarifaCampo.nCMACMComisMonto = 0
    If txtBancoDestinoCod.Enabled Then
        txtBancoDestinoCod.SetFocus
    Else
        txtDestinoCCI.SetFocus
    End If
End Sub
Private Sub cmdGrabar_Click()
    Dim oDMov As New COMDMov.DCOMMov
    Dim oNCapMov As New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oNCapGen As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim lnMovNro As Long
    Dim bTrans As Boolean
    Dim nMonto As Double
    Dim sCuenta As String
    Dim sProd As String
    Dim sTipoCuenta As String
    Dim nTpoPrograma As Integer
    Dim rsCuenta As New ADODB.Recordset
    Dim rsVoucher As New ADODB.Recordset
    Dim sBoleta As String
    Dim fbPersonaReaAhorros As Boolean
    Dim fbPersonaReaOtros As Boolean
    Dim fnCondicion As Integer
    Dim ObjTc As New COMDConstSistema.NCOMTipoCambio
    Dim nTC As Double
    Dim nmoneda As Integer
    Dim loLavDinero As frmMovLavDinero
    Set loLavDinero = New frmMovLavDinero
        
    bTrans = False
On Error GoTo ErrGrabar
    
    Call CalculaComision 'GEMO2021-01-05
    
    If Not ValidaDatos Then Exit Sub
    nMonto = txtMontoTotal.Text
    If gsOpeCod = gCMCCETransfOrdinaria And chkOrdenanteCargoCuenta.value = 1 Then
        sCuenta = Trim(txtOrdenanteCuenta.NroCuenta)
        sProd = txtOrdenanteCuenta.Prod
        Set rsCuenta = oNCapGen.GetDatosCuenta(sCuenta)
        nTpoPrograma = rsCuenta!nTpoPrograma
        sTipoCuenta = UCase(rsCuenta!cTipoCuenta)
        
        If sProd = gCapAhorros Then
            If nTpoPrograma <> 0 And nTpoPrograma <> 5 And nTpoPrograma <> 6 And nTpoPrograma <> 8 Then
                MsgBox "El subproducto de la cuenta no es permitido para esta operación", vbInformation, "¡Aviso!"
                cmdCancelar_Click
                Exit Sub
            End If
        Else
            If nTpoPrograma <> 0 And nTpoPrograma <> 1 Then
                MsgBox "El subproducto de la cuenta no es permitido para esta operación", vbInformation, "¡Aviso!"
                cmdCancelar_Click
                Exit Sub
            End If
        End If
        If Not oNCapMov.ValidaSaldoCuenta(sCuenta, nMonto, gCMCCETransfOrdinaria) Then
           MsgBox "La Cuenta NO posee saldo suficiente", vbInformation, "¡Aviso!"
           txtMonto.SetFocus
           Exit Sub
        End If
        'If VerificarAutorizacion = False Then Exit Sub 'Solicita Autorizacion de Retiros para montos altos.
    End If
    If MsgBox("¿Esta seguro de Guardar los datos de la Transferencia?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
        Dim sPersLavDinero As String, sReaPersLavDinero As String, sBenPersLavDinero As String
        Dim nSaldo As Double, nPorcDisp As Double
        Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
        Dim nMontoLavDinero As Double
        
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
        If Not (IIf(gsOpeCod = gCMCCETransfOrdinaria And chkOrdenanteCargoCuenta.value = 1, clsExo.EsCuentaExoneradaLavadoDinero(sCuenta), clsExo.EsPersonaExoneradaLavadoDinero(txtOrdenante.Text))) Then
            Set clsExo = Nothing
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            
            If Right(cboMoneda.Text, 1) = gMonedaNacional Then
                nTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
            Else
                nTC = 1
            End If
            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                loLavDinero.TitPersLavDinero = Trim(txtOrdenante.Text)
                sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, IIf(gsOpeCod = gCMCCETransfOrdinaria And chkOrdenanteCargoCuenta.value = 1, sCuenta, ""), Mid(Me.Caption, 15), True, IIf(gsOpeCod = gCMCCETransfOrdinaria And chkOrdenanteCargoCuenta.value = 1, sTipoCuenta, ""), , , , , IIf(gsOpeCod = gCMCCETransfOrdinaria And chkOrdenanteCargoCuenta.value = 1, Mid(sCuenta, 9, 1), Right(cboMoneda, 1)), , gnTipoREU, gnMontoAcumulado, gsOrigen, , gsOpeCod)
                If loLavDinero.OrdPersLavDinero = "" Then
                    cmdCancelar_Click
                    Exit Sub
                End If
            End If
        Else
            Set clsExo = Nothing
        End If
        fbPersonaReaAhorros = False
        fbPersonaReaOtros = False
        If (loLavDinero.OrdPersLavDinero = "Exit") Then
            Dim oPersonaSPR As UPersona_Cli
            Dim oPersonaU As COMDPersona.UCOMPersona
            Dim nTipoConBN As Integer
            Dim sConPersona As String
            Dim pbClienteReforzado As Boolean
            Dim rsAgeParam As Recordset
            Dim objCap As COMNCaptaGenerales.NCOMCaptaMovimiento
            Dim lnMontoX As Double, lnTC As Double
            
            Set oPersonaU = New COMDPersona.UCOMPersona
            Set oPersonaSPR = New UPersona_Cli
            
            fbPersonaReaAhorros = False
            pbClienteReforzado = False
            fnCondicion = 0
            
            If (gsOpeCod = gCMCCETransfOrdinaria And chkOrdenanteCargoCuenta.value = 1) Then
                oPersonaSPR.RecuperaPersona Trim(txtOrdenante.Text)
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
                     MsgBox "El Cliente: " & Trim(lblOrdenanteNombre.Caption) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
                        frmPersRealizaOpeGeneral.Inicia Me.Caption & " (Persona " & sConPersona & ")", gsOpeCod
                        fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar
                    
                    If Not fbPersonaReaAhorros Then
                        MsgBox "Se va a proceder a Anular la Operacion ", vbInformation, "¡Aviso!"
                        cmdCancelar_Click
                        Exit Sub
                    End If
                End If
            End If
            If gsOpeCod = gCMCCETransfPagoTarjCred Or Not pbClienteReforzado Then
                fnCondicion = 0
                lnMontoX = nMonto
                pbClienteReforzado = False
                
                nTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set objCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                Set rsAgeParam = objCap.getCapAbonoAgeParam(gsCodAge)
                Set objCap = Nothing
                
                If Right(cboMoneda, 1) = 1 Then
                    lnMontoX = Round(lnMontoX / nTC, 2)
                End If
                If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                    If lnMontoX >= rsAgeParam!nMontoMin And lnMontoX <= rsAgeParam!nMontoMax Then
                        frmPersRealizaOpeGeneral.Inicia Me.Caption, gsOpeCod
                        fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar
                        If Not fbPersonaReaAhorros Then
                            MsgBox "Se va a proceder a Anular la Operacion", vbInformation, "¡Aviso!"
                            cmdCancelar_Click
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        
        If gsOpeCod = gCMCCETransfOrdinaria Then
            lnMovNro = oCCE.CCE_RegistrarTransferenciaOrdinaria(gdFecSis, Right(gsCodAge, 2), gsCodUser, gsOpeCod, "CCE - REGISTRO DE TRANSFERENCIA ORDINARIA - ORIGINANTE", Trim(txtBancoDestinoCod.Text), _
                                Trim(txtOrdenante.Text), Trim(lblOrdenanteNombre.Caption), Right(txtOrdenanteTpoDoc.Text, 2), Trim(txtOrdenanteNroDoc.Text), oPersonaOrd.Domicilio, IIf(oPersonaOrd.Celular = "", oPersonaOrd.Telefonos, oPersonaOrd.Celular), Right(cboMoneda.Text, 1), chkIncluyeComis.value, _
                                CCur(Trim(txtMontoATransferir.Text)), Trim(txtDestinoCCI.Text), chkOrdenanteCargoCuenta.value, IIf(Not Len(txtOrdenanteCuenta.NroCuenta) = 18, "", txtOrdenanteCuenta.NroCuenta), _
                                Trim(txtOrdenanteCCIOrigen.Text), chkOrdenanteMismoTit.value, tTarifaCampo.sTariCod, tTarifaCampo.sTariCriApl, tTarifaCampo.sComisSigno, tTarifaCampo.nCCEComisMonto, tTarifaCampo.nCMACMComisMonto, CCur(Trim(txtITF.Text)))
        ElseIf gsOpeCod = gCMCCETransfPagoTarjCred Then
            lnMovNro = oCCE.CCE_RegistraTransferenciaPagoTarjCred(gdFecSis, Right(gsCodAge, 2), gsCodUser, gsOpeCod, "CCE - REGISTRO DE TRANSFERENCIA PAGO TARJETA DE CRÈDITO - ORIGINANTE", Trim(txtBancoDestinoCod.Text), _
                                 Trim(txtOrdenante.Text), Trim(lblOrdenanteNombre.Caption), Right(txtOrdenanteTpoDoc.Text, 2), Trim(txtOrdenanteNroDoc.Text), oPersonaOrd.Domicilio, IIf(oPersonaOrd.Celular = "", oPersonaOrd.Telefonos, oPersonaOrd.Celular), Right(cboMoneda.Text, 1), chkIncluyeComis.value, _
                                CCur(Trim(txtMontoATransferir.Text)), Trim(txtNroTarjCred.Text), Trim(txtBeneficiario.Text), lblBeneficiarioNombre.Caption, Right(txtBeneficiarioTpoDoc.Text, 2), Trim(txtBeneficiarioNroDoc.Text), oPersonaBenef.Domicilio, IIf(oPersonaBenef.Celular = "", oPersonaBenef.Telefonos, oPersonaBenef.Celular), _
                                 tTarifaCampo.sTariCod, tTarifaCampo.sTariCriApl, tTarifaCampo.sComisSigno, tTarifaCampo.nCCEComisMonto, tTarifaCampo.nCMACMComisMonto)
        End If
        If lnMovNro > 0 Then
        
            'Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'COMENTADO VAPA20170510
            If pbClienteReforzado = False Then
            Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , lnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'AGREGADO VAPA20170510
            MsgBox "Se ha registrado satisfactoriamente la Operación", vbInformation, "¡Aviso!"
            Else
            Call frmPersRealizaOpeGeneral.InsertaPersonasOperacion(lnMovNro, txtOrdenanteCuenta.NroCuenta, fnCondicion)
            'InsertaPersonasOperacion
            MsgBox "Se ha registrado satisfactoriamente la Operación", vbInformation, "¡Aviso!"
            End If
            Set rsVoucher = oCCE.DatosVoucherTransferencia(lnMovNro)
            sBoleta = oNCapMov.ImprimeVoucherTransferencia(rsVoucher, gbImpTMU, gsCodUser, gsOpeCod)
            Do
                clsprevio.PrintSpool sLpt, sBoleta
            Loop While MsgBox("Desea Reimprimir el voucher?", vbInformation + vbYesNo, "¡Aviso!") = vbYes
            'INICIO JHCU ENCUESTA 16-10-2019
            Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
            'FIN
            If MsgBox("¿Desea registrar otra transferencia?", vbYesNo + vbInformation, "¡Aviso!") = vbNo Then
                Unload Me
                Exit Sub
            Else
                cmdCancelar_Click
            End If
        Else
            MsgBox "No se pudo realizar la operación de transferencia, " & Chr(13) & "si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "¡Aviso!"
        End If
    Exit Sub
ErrGrabar:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "¡Aviso!"
End Sub
Private Function ValidaDatos() As Boolean
    ValidaDatos = True
    If Len(txtBancoDestinoCod.Text) = 0 And Len(lblBancoDestinoNombre.Caption) = 0 Then
        MsgBox "No se ha encontrado el Banco de Destino.", vbInformation, "¡Aviso!"
        If txtBancoDestinoCod.Enabled Then
            txtBancoDestinoCod.SetFocus
        Else
            txtDestinoCCI.SetFocus
        End If
        ValidaDatos = False
        Exit Function
    End If
    If gsOpeCod = gCMCCETransfOrdinaria Then
        If Not IsNumeric(txtDestinoCCI.Text) Then
            MsgBox "La cuenta CCI de Destino no es válida. Verifique.", vbInformation, "¡Aviso!"
            txtDestinoCCI.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Not Len(txtDestinoCCI.Text) = 20 Then
            MsgBox "La cuenta CCI de Destino no es válida. Verifique.", vbInformation, "¡Aviso!"
            txtDestinoCCI.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Not oCCE.CCE_ValidaCuentaCCI(txtDestinoCCI.Text) Then
            MsgBox "La cuenta CCI Destino no es válida. Verifique", vbInformation, "¡Aviso!"
            txtDestinoCCI.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Not oCCE.CCE_EsEntidadReceptorTransferencia(Left(txtDestinoCCI.Text, 3), psOpeTransf) Then
            MsgBox "El banco destino asociado a la cuenta CCI no puede " & Chr(13) & "recibir este tipo de transferencia. Verifique.", vbInformation, "¡Aviso!"
            txtBancoDestinoCod.Text = ""
            lblBancoDestinoNombre.Caption = ""
            txtDestinoCCI.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If chkOrdenanteCargoCuenta.value = 1 And Not Len(txtOrdenanteCuenta.NroCuenta) = 18 Then
            MsgBox "No se ha indicado la Cuenta CMACM de Origen. Verifique", vbInformation, "¡Aviso!"
            txtOrdenanteCuenta.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If chkOrdenanteCargoCuenta.value = 1 And Len(txtOrdenanteCCIOrigen.Text) = 0 Then
            MsgBox "No se ha indicado la Cuenta CMACM de Origen. Verifique", vbInformation, "¡Aviso!"
            txtOrdenanteCuenta.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
    If gsOpeCod = gCMCCETransfPagoTarjCred Then
        If Not IsNumeric(txtNroTarjCred.Text) Then
            MsgBox "Ingrese un número de tarjeta válida.", vbInformation, "¡Aviso!"
            txtNroTarjCred.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Not Len(txtNroTarjCred.Text) = 16 Then
            MsgBox "Ingrese un número de tarjeta válida.", vbInformation, "¡Aviso!"
            txtNroTarjCred.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Len(txtNroTarjCred.Text) = 0 Then
            MsgBox "No se ha ingresado el numero de tarjeta de crédito. Verifique.", vbInformation, "¡Aviso!"
            txtNroTarjCred.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
    If Len(txtOrdenante.Text) = 0 Then
        MsgBox "No se ha seleccionado la persona ordenante de la transferencia.", vbInformation, "¡Aviso!"
        txtOrdenante.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Len(lblOrdenanteNombre.Caption) = 0 Then
        MsgBox "No se ha seleccionado la persona ordenante de la transferencia.", vbInformation, "¡Aviso!"
        txtOrdenante.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Len(txtOrdenanteTpoDoc.Text) = 0 Then
        MsgBox "No se ha encontrado un tipo de documento para la " & Chr(13) & "persona ordenante de la transferencia. Verifique.", vbInformation, "¡Aviso!"
        txtOrdenante.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Len(txtOrdenanteNroDoc.Text) = 0 Then
        MsgBox "No se ha encontrado el numero de documento para la " & Chr(13) & "persona ordenante de la transferencia. Verifique.", vbInformation, "¡Aviso!"
        txtOrdenante.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If gsOpeCod = gCMCCETransfPagoTarjCred Then
        If Len(txtBeneficiario.Text) = 0 Then
            MsgBox "No se ha seleccionado la persona beneficiaria de la transferencia.", vbInformation, "¡Aviso!"
            txtBeneficiario.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Len(lblBeneficiarioNombre.Caption) = 0 Then
            MsgBox "No se ha seleccionado la persona beneficiaria de la transferencia.", vbInformation, "¡Aviso!"
            txtBeneficiario.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Len(txtBeneficiarioTpoDoc.Text) = 0 Then
            MsgBox "No se ha encontrado un tipo de documento para la " & Chr(13) & "persona beneficiaria de la transferencia. Verifique.", vbInformation, "¡Aviso!"
            txtBeneficiario.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Len(txtBeneficiarioNroDoc.Text) = 0 Then
            MsgBox "No se ha encontrado el numero de documento para la " & Chr(13) & "persona beneficiaria de la transferencia. Verifique.", vbInformation, "¡Aviso!"
            txtBeneficiario.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
    If cboMoneda.ListIndex = -1 Then
        MsgBox "No se ha seleccionado el tipo de moneda para realizar la transferencia. Verifique.", vbInformation, "¡Aviso!"
        cboMoneda.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Len(txtMonto.Text) = 0 Then
        MsgBox "No se ha ingresado el monto para realizar la transferencia. Verifique.", vbInformation, "¡Aviso!"
        txtMonto.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If nVal(txtMonto.Text) = 0 Then
        MsgBox "El monto para la operación no puede ser 0.00. Verifique.", vbInformation, "¡Aviso!"
        txtMonto.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If nVal(txtMontoComision.Text) = 0 Then
        MsgBox "La Comisión para la operación no puede ser 0.00. Verifique.", vbInformation, "¡Aviso!"
        txtMonto.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If nVal(txtMontoATransferir.Text) <= 0 Then
        MsgBox "El monto a transferir no es correcto. Verifique.", vbInformation, "¡Aviso!"
        txtMonto.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If nVal(txtMontoTotal.Text) <= 0 Then
        MsgBox "El monto total no es correcto. Verifique.", vbInformation, "¡Aviso!"
        txtMonto.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If nVal(txtMontoATransferir.Text) > IIf(Right(cboMoneda.Text, 1) = 1, (fnLimMaxTrMN), (fnLimMaxTrME)) Then
        MsgBox "El monto a transferir no puede superar los " & IIf(Right(cboMoneda, 1) = 1, "S/ " & Format(fnLimMaxTrMN, "#,##0.00") & " SOLES", "$ " & Format(fnLimMaxTrME, "#,##0.00") & " DOLARES") & ". Verifique.", vbInformation, "¡Aviso!"
        txtMonto.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If nVal(txtMontoATransferir.Text) < IIf(Right(cboMoneda.Text, 1) = 1, (fnLimMinTrMN), (fnLimMinTrME)) Then
        MsgBox "El monto a transferir no puede ser menor de " & IIf(Right(cboMoneda, 1) = 1, "S/ " & Format(fnLimMinTrMN, "#,##0.00") & " SOLES", "$ " & Format(fnLimMinTrME, "#,##0.00") & " DOLARES") & ". Verifique.", vbInformation, "¡Aviso!"
        txtMonto.SetFocus
        ValidaDatos = False
        Exit Function
    End If
End Function
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsDatosPers As New ADODB.Recordset
    Dim rsCtaCCI As New ADODB.Recordset
    Dim sCuenta As String
    If KeyCode = vbKeyF10 And chkOrdenanteCargoCuenta.value = 1 Then
        sCuenta = frmCCEATMCargaCuentas.RecuperaCuenta(gsOpeCod)
        If sCuenta <> "" Then
            txtOrdenanteCuenta.NroCuenta = sCuenta
            txtOrdenanteCuenta.Enabled = False
            Set rsCtaCCI = oCCE.CCE_ConsultaCuentaCCI(txtOrdenanteCuenta.NroCuenta)
            If Not (rsCtaCCI.EOF And rsCtaCCI.BOF) Then
                txtOrdenanteCCIOrigen.Text = rsCtaCCI!cCCI
            End If
            Set rsDatosPers = oCCE.CCE_ObtieneDatosPersona(txtOrdenanteCuenta.NroCuenta)
            If Not (rsDatosPers.EOF And rsDatosPers.BOF) Then
                If rsDatosPers!cPersCod = gsCodPersUser Then
                    MsgBox "No se puede registrar una transferencia de si mismo", vbInformation, "¡Aviso!"
                    chkOrdenanteCargoCuenta.value = 0
                    Exit Sub
                End If
                txtOrdenante.Text = rsDatosPers!cPersCod
                txtOrdenante.Enabled = False
                Call oPersonaOrd.RecuperaPersona(rsDatosPers!cPersCod, , gsCodUser)
                If Not (Right(rsDatosPers!cPersIDTpoDesc, 1) = 6) Then
                    lblOrdenanteNombre.Caption = (oPersonaOrd.ApellidoPaterno) & "-" _
                                    & IIf(BuscaNombre(oPersonaOrd.NombreCompleto, 1) = "", oPersonaOrd.ApellidoMaterno, BuscaNombre(oPersonaOrd.NombreCompleto, 1)) & "-" _
                                    & BuscaNombre(oPersonaOrd.Nombres, 2)
                Else
                    lblOrdenanteNombre.Caption = oPersonaOrd.NombreCompleto
                End If
                txtOrdenanteTpoDoc.Text = rsDatosPers!cPersIDTpoDesc
                txtOrdenanteNroDoc.Text = rsDatosPers!cPersIDnro
            End If
        End If
        cboMoneda.ListIndex = IIf(Len(sCuenta) = 0, -1, IIf(Mid(txtOrdenanteCuenta.NroCuenta, 9, 1) = "1", 0, 1))
        cboMoneda.Enabled = False
    End If
End Sub
Private Sub Form_Load()
    Dim rsLim As ADODB.Recordset
    Set oCCE = New COMNCajaGeneral.NCOMCCE
    Set oPersonaOrd = New UPersona_Cli
    Set oPersonaBenef = New UPersona_Cli
    txtBancoDestinoCod.psRaiz = "Instituciones Financieras CCE"
    txtBancoDestinoCod.rs = oCCE.CCE_ObtieneIFIS
        Select Case fsOpeCod
            Case gCMCCETransfOrdinaria
                Me.Caption = UCase("Transferencia Interbancaria Ordinaria - Originante")
                psOpeTransf = gCCETransfOrdinaria
            Case gCMCCETransfPagoTarjCred
                Me.Caption = UCase("Pago InterBancario de Tarjeta de Crédito - Originante")
                psOpeTransf = gCCETransfPagoTarjCred
        End Select
    Set rsLim = oCCE.CCE_ObtieneLimitesTransf
    If Not (rsLim.EOF And rsLim.BOF) Then
        fnLimMaxTrMN = rsLim!nLimMaxTrMN
        fnLimMaxTrME = rsLim!nLimMaxTrME
        fnLimMinTrMN = rsLim!nLimMinTrMN
        fnLimMinTrME = rsLim!nLimMinTrME
    End If
    ReiniciaCuenta
    HabilitaCampos
End Sub
Private Sub HabilitaCampos()
    Select Case fsOpeCod
        Case gCMCCETransfOrdinaria
            FraBeneficiario.Visible = False
            chkIncluyeComis.Top = 2600
            fraMontos.Top = 2680
            cmdGrabar.Top = 3850
            cmdCancelar.Top = 3850
            cmdSalir.Top = 3850
            Me.Height = 4600
        Case gCMCCETransfPagoTarjCred
            txtBancoDestinoCod.Enabled = True
            lblNroTarjCred.Visible = True
            txtNroTarjCred.Visible = True
            lblCCIDestino.Visible = False
            txtDestinoCCI.Visible = False
            chkOrdenanteCargoCuenta.Visible = False
            fraOrdenanteCuenta.Visible = False
            FraOrdenante.Top = 840
            FraBeneficiario.Top = 1650
            chkOrdenanteMismoTit.Visible = False
            chkIncluyeComis.Top = 2350
            fraMontos.Top = 2450
            cmdGrabar.Top = 3600
            cmdCancelar.Top = 3600
            cmdSalir.Top = 3600
            Me.Height = 4400
    End Select
End Sub
Private Sub txtBancoDestinoCod_EmiteDatos()
    lblBancoDestinoNombre.Caption = txtBancoDestinoCod.psDescripcion
    If Not Len(txtBancoDestinoCod.Text) = 0 Then
        txtNroTarjCred.SetFocus
    End If
End Sub
Private Sub txtBeneficiario_EmiteDatos()
    Dim rs As ADODB.Recordset
    lblBeneficiarioNombre.Caption = ""
    txtBeneficiarioTpoDoc.Text = ""
    txtBeneficiarioNroDoc.Text = ""
    If txtBeneficiario.Text = gsCodPersUser Then
        MsgBox "No se puede registrar una transferencia de si mismo", vbInformation, "¡Aviso!"
        txtBeneficiario.Text = ""
        Exit Sub
    End If
    If Not Len(txtBeneficiario.Text) = 0 Then
        Call oPersonaBenef.RecuperaPersona(txtBeneficiario.Text, , gsCodUser)
        Set rs = oCCE.CCE_ObtieneDocumentoPersona(txtBeneficiario.psCodigoPersona)
        If rs.EOF And rs.BOF Then
            MsgBox "No se ha encontrado ningun documento de la persona. Verifique.", vbInformation + vbInformation, "Aviso"
            Exit Sub
        Else
            If Not (Right(rs!cPersIDTpoDesc, 1) = 6) Then
                lblBeneficiarioNombre.Caption = (oPersonaBenef.ApellidoPaterno) & "-" _
                                    & IIf(BuscaNombre(oPersonaBenef.NombreCompleto, 1) = "", oPersonaBenef.ApellidoMaterno, BuscaNombre(oPersonaBenef.NombreCompleto, 1)) & "-" _
                                    & BuscaNombre(oPersonaBenef.Nombres, 2)
            Else
                lblBeneficiarioNombre.Caption = oPersonaBenef.NombreCompleto
            End If
            txtBeneficiarioTpoDoc.Text = rs!cPersIDTpoDesc
            txtBeneficiarioNroDoc.Text = rs!cPersIDnro
            cboMoneda.SetFocus
        End If
    End If
End Sub
Private Sub txtDestinoCCI_Change()
    txtBancoDestinoCod.Text = ""
    lblBancoDestinoNombre.Caption = ""
End Sub
Private Sub txtDestinoCCI_GotFocus()
    fEnfoque txtDestinoCCI
End Sub
Private Sub txtDestinoCCI_LostFocus()
    Dim rsBcoDest As ADODB.Recordset
    If Not Len(txtDestinoCCI.Text) = 20 Then Exit Sub
    Set rsBcoDest = oCCE.CCE_ObtieneIFIxCodBCR(Left(txtDestinoCCI.Text, 3))
    If Not (rsBcoDest.EOF And rsBcoDest.BOF) Then
        txtBancoDestinoCod.Text = rsBcoDest!cCodBCR
        lblBancoDestinoNombre.Caption = rsBcoDest!cPersNombre
        CalculaComision
    End If
End Sub
Private Sub txtDestinoCCI_KeyPress(KeyAscii As Integer)
    Dim rsBcoDest As ADODB.Recordset
    KeyAscii = NumerosEnteros(KeyAscii)
    txtBancoDestinoCod.Text = ""
    lblBancoDestinoNombre.Caption = ""
    If KeyAscii = 13 And Not IsNumeric(txtDestinoCCI.Text) Then
        MsgBox "Ingrese una cuenta CCI válida.", vbInformation, "¡Aviso!"
        txtDestinoCCI.SetFocus
        Exit Sub
    End If
    If KeyAscii = 13 And Len(txtDestinoCCI.Text) = 20 Then
        If Not oCCE.CCE_ValidaCuentaCCI(txtDestinoCCI.Text) Then
            MsgBox "Ingrese una cuenta CCI válida.", vbInformation, "¡Aviso!"
            txtDestinoCCI.SetFocus
            Exit Sub
        End If
        Set rsBcoDest = oCCE.CCE_ObtieneIFIxCodBCR(Left(txtDestinoCCI.Text, 3))
        If Not (rsBcoDest.EOF And rsBcoDest.BOF) Then
            txtBancoDestinoCod.Text = rsBcoDest!cCodBCR
            lblBancoDestinoNombre.Caption = rsBcoDest!cPersNombre
            If Not oCCE.CCE_EsEntidadReceptorTransferencia(Left(txtDestinoCCI.Text, 3), psOpeTransf) Then
                MsgBox "El banco destino asociado a la cuenta CCI no puede " & Chr(13) & "recibir este tipo de transferencia. Verifique.", vbInformation, "¡Aviso!"
                txtBancoDestinoCod.Text = ""
                lblBancoDestinoNombre.Caption = ""
                txtDestinoCCI.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "El banco destino asociado a la cuenta CCI no " & Chr(13) & "existe para esta operaciòn. Verifique.", vbInformation, "¡Aviso!"
            txtDestinoCCI.SetFocus
            Exit Sub
        End If
        
      If Not Len(txtOrdenanteCuenta.NroCuenta) = 18 Then
        txtOrdenanteCuenta.SetFocus 'VAPA20170627 AGREGO IF
      End If
'        If chkOrdenanteCargoCuenta.Visible Then
'            chkOrdenanteCargoCuenta.SetFocus
'        ElseIf FraOrdenante.Visible Then
'            txtOrdenante.SetFocus
'        End If
    ElseIf KeyAscii = 13 And Not Len(txtDestinoCCI.Text) = 20 Then
        MsgBox "Ingrese una cuenta CCI válida.", vbInformation, "¡Aviso!"
        txtDestinoCCI.SetFocus
        Exit Sub
    End If
End Sub
Private Sub txtMonto_GotFocus()
    fEnfoque txtMonto
End Sub
Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 10, 3)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub
Private Sub txtMonto_KeyUp(KeyCode As Integer, Shift As Integer)
    CalculaComision
End Sub
Private Sub txtMonto_LostFocus()
    If Trim(txtMonto.Text) = "" Or Trim(txtMonto.Text) = "." Then
        txtMonto.Text = "0.00"
    End If
    txtMonto.Text = Format(txtMonto.Text, "#0.00")
    CalculaComision
End Sub
Private Sub txtNroTarjCred_GotFocus()
     fEnfoque txtNroTarjCred
End Sub
Private Sub txtNroTarjCred_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 And Not IsNumeric(txtNroTarjCred.Text) Then
        MsgBox "Ingrese un número de tarjeta válida.", vbInformation, "¡Aviso!"
        txtNroTarjCred.SetFocus
        Exit Sub
    End If
    If KeyAscii = 13 And Not Len(txtNroTarjCred.Text) = 16 Then
        MsgBox "Ingrese un número de tarjeta válida.", vbInformation, "¡Aviso!"
        txtNroTarjCred.SetFocus
        Exit Sub
    End If
    If KeyAscii = 13 And Len(txtNroTarjCred.Text) = 16 Then
        txtOrdenante.SetFocus
    End If
End Sub
Private Sub txtOrdenante_EmiteDatos()
    Dim rs As ADODB.Recordset
    lblOrdenanteNombre.Caption = ""
    txtOrdenanteTpoDoc.Text = ""
    txtOrdenanteNroDoc.Text = ""
    If txtOrdenante.Text = gsCodPersUser Then
        MsgBox "No se puede registrar una transferencia de si mismo", vbInformation, "¡Aviso!"
        txtOrdenante.Text = ""
        Exit Sub
    End If
    If Not Len(txtOrdenante.Text) = 0 Then
        Call oPersonaOrd.RecuperaPersona(txtOrdenante.Text, , gsCodUser)
        Set rs = oCCE.CCE_ObtieneDocumentoPersona(txtOrdenante.psCodigoPersona)
        If rs.EOF And rs.BOF Then
            MsgBox "No se ha encontrado un tipo de documento para la " & Chr(13) & "persona ordenante de la transferencia. Verifique.", vbInformation + vbInformation, "¡Aviso!"
            Exit Sub
        Else
            If Not (Right(rs!cPersIDTpoDesc, 1) = 6) Then
                lblOrdenanteNombre.Caption = (oPersonaOrd.ApellidoPaterno) & "-" _
                                    & IIf(BuscaNombre(oPersonaOrd.NombreCompleto, 1) = "", oPersonaOrd.ApellidoMaterno, BuscaNombre(oPersonaOrd.NombreCompleto, 1)) & "-" _
                                    & BuscaNombre(oPersonaOrd.Nombres, 2)
            Else
                lblOrdenanteNombre.Caption = oPersonaOrd.NombreCompleto
            End If
            txtOrdenanteTpoDoc.Text = rs!cPersIDTpoDesc
            txtOrdenanteNroDoc.Text = rs!cPersIDnro
            If FraBeneficiario.Visible Then
                txtBeneficiario.SetFocus
            Else
                cboMoneda.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txtOrdenanteCuenta_KeyPress(KeyAscii As Integer)
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim sMsg As String
Dim bEstadoCuenta As Boolean 'add pti1 23-01-19 ERS081-2018 ACTA 06-2019
Dim rsDatosPers As New ADODB.Recordset
Dim rsCtaCCI As New ADODB.Recordset
Dim sCuenta As String
Dim loVistoElectronico As New frmVistoElectronico
Dim lbVistoVal As Boolean

Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    If KeyAscii = 13 Then
    '*********************add pti1 23-01-19 ERS081-2018 ACTA 06-2019 **********************
    'comentado por pti1
    '        If Not (txtOrdenanteCuenta.Prod) = "232" And Not (txtOrdenanteCuenta.Prod) = "234" Then
    '            MsgBox "Cuenta no válida", vbInformation, "¡Aviso!"
    '            txtOrdenanteCuenta.SetFocus
    '            Exit Sub
    '        End If
        If fsOpeCod = 930001 Then
           bEstadoCuenta = clsCap.ValidaCuentaOperacionCCI(txtOrdenanteCuenta.NroCuenta)
           
           If bEstadoCuenta Then
          
           Else
            MsgBox "Este Tipo de Cuenta no está habilitada para realizar la operación", vbInformation, "¡Aviso!"
            txtOrdenanteCuenta.SetFocus
            Exit Sub
           End If
           
        End If
        '*********************END PTI1
        sMsg = clsCap.ValidaCuentaOperacion(txtOrdenanteCuenta.NroCuenta)
        If Len(sMsg) = 0 Then
'            Set loVistoElectronico = New frmVistoElectronico
'            lbVistoVal = loVistoElectronico.Inicio(5, gsOpeCod)
'            If Not lbVistoVal Then
'                MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones.", vbInformation, "Mensaje del Sistema"
'                Exit Sub
'            End If
'            loVistoElectronico.RegistraVistoElectronico (0)
            sCuenta = txtOrdenanteCuenta.Cuenta
            txtOrdenanteCuenta.Enabled = False
            Set rsCtaCCI = oCCE.CCE_ConsultaCuentaCCI(txtOrdenanteCuenta.NroCuenta)
            If Not (rsCtaCCI.EOF And rsCtaCCI.BOF) Then
                txtOrdenanteCCIOrigen.Text = rsCtaCCI!cCCI
            End If
            Set rsDatosPers = oCCE.CCE_ObtieneDatosPersona(txtOrdenanteCuenta.NroCuenta)
            If Not (rsDatosPers.EOF And rsDatosPers.BOF) Then
                If rsDatosPers!cPersCod = gsCodPersUser Then
                    MsgBox "No se puede registrar una transferencia de si mismo", vbInformation, "¡Aviso!"
                    chkOrdenanteCargoCuenta.value = 0
                    Exit Sub
                End If
                txtOrdenante.Text = rsDatosPers!cPersCod
                txtOrdenante.Enabled = False
                Call oPersonaOrd.RecuperaPersona(rsDatosPers!cPersCod, , gsCodUser)
                If Not (Right(rsDatosPers!cPersIDTpoDesc, 1) = 6) Then
                    lblOrdenanteNombre.Caption = (oPersonaOrd.ApellidoPaterno) & "-" _
                                    & IIf(BuscaNombre(oPersonaOrd.NombreCompleto, 1) = "", oPersonaOrd.ApellidoMaterno, BuscaNombre(oPersonaOrd.NombreCompleto, 1)) & "-" _
                                    & BuscaNombre(oPersonaOrd.Nombres, 2)
                Else
                    lblOrdenanteNombre.Caption = oPersonaOrd.NombreCompleto
                End If
                txtOrdenanteTpoDoc.Text = rsDatosPers!cPersIDTpoDesc
                txtOrdenanteNroDoc.Text = rsDatosPers!cPersIDnro
            End If
            cboMoneda.ListIndex = IIf(Len(sCuenta) = 0, -1, IIf(Mid(txtOrdenanteCuenta.NroCuenta, 9, 1) = "1", 0, 1))
            cboMoneda.Enabled = False
            CalculaComision
        Else
            MsgBox sMsg, vbInformation, "¡Aviso!"
            txtOrdenanteCuenta.SetFocus
        End If
    End If
End Sub
Private Function VerificarAutorizacion() As Boolean
Dim ocapaut As COMDCaptaGenerales.COMDCaptAutorizacion
Dim oCapAutN  As COMNCaptaGenerales.NCOMCaptAutorizacion
Dim oPers As COMDPersona.UCOMAcceso
Dim rs As New ADODB.Recordset
Dim lnMonTopD As Double
Dim lnMonTopS As Double
Dim lsmensaje As String
Dim gsGrupo As String
Dim sCuenta As String, sNivel As String
Dim lbEstadoApr As Boolean
Dim nMonto As Double
Dim nmoneda As Moneda

sCuenta = txtOrdenanteCuenta.NroCuenta
nMonto = txtMontoTotal.Text
nmoneda = CLng(Mid(sCuenta, 9, 1))
'Obtiene los grupos al cual pertenece el usuario
Set oPers = New COMDPersona.UCOMAcceso
    gsGrupo = oPers.CargaUsuarioGrupo(gsCodUser, gsDominio)
Set oPers = Nothing
 
'Verificar Montos
Set ocapaut = New COMDCaptaGenerales.COMDCaptAutorizacion
    Set rs = ocapaut.ObtenerMontoTopNivAutRetCan(gsGrupo, "1", gsCodAge, gsCodPersUser)
Set ocapaut = Nothing
 
If Not (rs.EOF And rs.BOF) Then
    lnMonTopD = rs("nTopDol")
    lnMonTopS = rs("nTopSol")
    sNivel = rs("cNivCod")
Else
    MsgBox "Usuario no Autorizado para realizar Operacion", vbInformation, "¡Aviso!"
    VerificarAutorizacion = False
    Exit Function
End If

If nmoneda = gMonedaNacional Then
    If nMonto <= lnMonTopS Then
        VerificarAutorizacion = True
        Exit Function
    End If
Else
    If nMonto <= lnMonTopD Then
        VerificarAutorizacion = True
        Exit Function
    End If
End If
   
Set oCapAutN = New COMNCaptaGenerales.NCOMCaptAutorizacion
If sMovNroAut = "" Then 'Si es nueva, registra
    oCapAutN.NuevaSolicitudAutorizacion sCuenta, "1", nMonto, gdFecSis, gsCodAge, gsCodUser, nmoneda, gOpeAutorizacionRetiro, sNivel, sMovNroAut
    MsgBox "Solicitud Registrada, comunique a su Admnistrador para la Aprobación..." & Chr$(10) & _
        " No salir de esta operación mientras se realice el proceso..." & Chr$(10) & _
        " Porque sino se procedera a grabar otra Solicitud...", vbInformation, "¡Aviso!"
    VerificarAutorizacion = False
Else
    'Valida el estado de la Solicitud
    If Not oCapAutN.VerificarAutorizacion(sCuenta, "1", nMonto, sMovNroAut, lsmensaje) Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        VerificarAutorizacion = False
    Else
        VerificarAutorizacion = True
    End If
End If
Set oCapAutN = Nothing
End Function
Private Function BuscaNombre(ByVal psNombre As String, ByVal pnTpoBusqueda As Integer) As String
Dim sCadTmp As String
Dim PosIni As Integer
Dim PosFin As Integer
Dim PosIni2 As Integer
Dim i As Integer
Dim n As Integer
    sCadTmp = ""
    Select Case pnTpoBusqueda
        Case 1 'Apellido de casada
           PosIni = InStr(1, psNombre, "\")
           If PosIni <> 0 Then
                PosIni2 = InStr(1, psNombre, "VDA")
                If PosIni2 <> 0 Then
                    PosIni = PosIni2
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                Else
                    PosIni = PosIni + 1
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                End If
                sCadTmp = Trim(Mid(psNombre, PosIni, PosFin - PosIni))
            Else
                sCadTmp = ""
            End If
        Case 2 'Nombres
            n = 0
            i = 1
                Do While i <= Len(psNombre)
                    If Not n = 2 Then
                        sCadTmp = sCadTmp & Mid(psNombre, i, 1)
                    End If
                    If Mid(psNombre, i, 1) = " " Then
                        n = n + 1
                    End If
                    If n = 2 Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
    End Select
    BuscaNombre = sCadTmp
End Function
