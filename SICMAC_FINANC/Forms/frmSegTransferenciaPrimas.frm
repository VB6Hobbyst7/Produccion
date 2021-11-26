VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSegTransferenciaPrimas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pago Seguro Tarjeta"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7410
   Icon            =   "frmSegTransferenciaPrimas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   9
      Top             =   6960
      Width           =   1170
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5880
      TabIndex        =   8
      Top             =   6960
      Width           =   1170
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "Extornar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   5
      Top             =   6960
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5880
      TabIndex        =   2
      Top             =   6960
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2566
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Ramo"
      TabPicture(0)   =   "frmSegTransferenciaPrimas.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraPeriodo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraProducto"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdBuscar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5640
         TabIndex        =   30
         Top             =   720
         Width           =   1170
      End
      Begin VB.Frame FraProducto 
         Caption         =   " Producto"
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
         TabIndex        =   26
         Top             =   480
         Width           =   2655
         Begin VB.ComboBox cboProducto 
            Height          =   315
            ItemData        =   "frmSegTransferenciaPrimas.frx":0326
            Left            =   240
            List            =   "frmSegTransferenciaPrimas.frx":0330
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Seleccionar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4680
            TabIndex        =   27
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   2250
         End
      End
      Begin VB.Frame FraPeriodo 
         Caption         =   " Periodo"
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
         Left            =   2880
         TabIndex        =   3
         Top             =   480
         Width           =   2655
         Begin VB.ComboBox cboMes 
            Height          =   315
            ItemData        =   "frmSegTransferenciaPrimas.frx":038B
            Left            =   240
            List            =   "frmSegTransferenciaPrimas.frx":0395
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtAño 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   1
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblAño 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1680
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblMes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   1290
         End
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   5055
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Detalle"
      TabPicture(0)   =   "frmSegTransferenciaPrimas.frx":03F0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraOperacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraCtaIF"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame FraCtaIF 
         Caption         =   " Cuenta Institución Financiera "
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
         Height          =   1695
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   6615
         Begin Sicmact.TxtBuscar txtBuscarBanco 
            Height          =   330
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   582
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
         End
         Begin VB.Label lblTipoPago 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   25
            Top             =   1170
            Width           =   1410
         End
         Begin VB.Label lblCtaBanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   120
            TabIndex        =   24
            Top             =   375
            Width           =   2250
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Pago :"
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   1200
            Width           =   825
         End
         Begin VB.Label lblDescBanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2520
            TabIndex        =   22
            Top             =   360
            Width           =   3720
         End
         Begin VB.Label lblDescCtaBanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   120
            TabIndex        =   21
            Top             =   750
            Width           =   6135
         End
      End
      Begin VB.Frame FraOperacion 
         Caption         =   " Datos de la operación "
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
         Height          =   2295
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   6495
         Begin VB.TextBox txtMovDesc 
            Height          =   720
            Left            =   720
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   840
            Width           =   5475
         End
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   330
            Left            =   720
            TabIndex        =   31
            Top             =   360
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblnMovNro 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4200
            TabIndex        =   18
            Top             =   1680
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lblsMovNro 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   5160
            TabIndex        =   17
            Top             =   1680
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lblMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   720
            TabIndex        =   16
            Top             =   1650
            Width           =   1290
         End
         Begin VB.Label Label2 
            Caption         =   "Monto :"
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Glosa :"
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   705
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha :"
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   705
         End
      End
   End
End
Attribute VB_Name = "frmSegTransferenciaPrimas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmSegTransferenciaPrimas
'** Descripción : Formulario para realizar las transferencias de primas declaradas ERS028-2017
'** Creación : APRI, 20180129 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim oDSeg As DSeguros
Dim oOpe As DOperacion
Dim nOperacion As OpeCGOpeBancos
Dim oNContFunc As NContFunciones
Dim lsCtaContBanco As String
Dim nTipoDoc As TpoDoc
Dim rs As ADODB.Recordset
Dim nTipoSeg As Integer
Dim objPista As COMManejador.Pista 'ARLO20170217
Dim bCancela As Boolean

Public Sub inicio(ByVal pnOperacion As OpeCGOpeBancos)
nOperacion = pnOperacion

If nOperacion = "421008" Or nOperacion = "422008" Then
    Me.Caption = "Transferencia de Primas de Seguros"
    txtBuscarBanco.psRaiz = "Cuentas de Bancos"
    CargaComboConstante 10084, cboProducto
    CargaComboConstante 1010, cboMes
    HabilitaControles (False)
    HabilitaControlesExtorno False
Else
    CargaComboConstante 10084, cboProducto
    Me.Caption = "Extorno Transferencia de Primas de Seguros"
    HabilitaControlesExtorno True
End If

Me.Show 1
End Sub

Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
FraProducto.Enabled = Not pbHabilita
FraPeriodo.Enabled = Not pbHabilita
FraCtaIF.Enabled = pbHabilita
FraOperacion.Enabled = pbHabilita
cmdAceptar.Enabled = pbHabilita
End Sub

Private Sub HabilitaControlesExtorno(ByVal pbHabilitaExt As Boolean)
lblMes.Visible = pbHabilitaExt
lblAño.Visible = pbHabilitaExt
cboMes.Visible = Not pbHabilitaExt
txtAño.Visible = Not pbHabilitaExt
lblCtaBanco.Visible = pbHabilitaExt
txtBuscarBanco.Visible = Not pbHabilitaExt
cmdExtornar.Visible = pbHabilitaExt
cmdAceptar.Visible = Not pbHabilitaExt
cmdCerrar.Visible = pbHabilitaExt
cmdCancelar.Visible = Not pbHabilitaExt
cmdExtornar.Enabled = Not pbHabilitaExt 'APRI20180901 MEJORA
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAño.SetFocus
    End If
End Sub


Private Sub cmdAceptar_Click()
Dim oDocPago As clsDocPago
Dim oCtasIF As NCajaCtaIF
Dim oDoc As DDocumento
Dim oNSeg As NSeguros
Dim oImp As NContImprimir

Dim lsDocVoucher As String
Dim lsPersNombre As String
Dim lsSubCuentaIF As String
Dim lsDocNro As String
Dim lsFecha As String
Dim lsDocNroTmp As String
Dim lsDocVoucherTmp As String
Dim lsPlanillaNro As String
Dim lsMovNro As String
Dim lsTpoIf As String
Dim lsCtaBanco As String
Dim lsImpre As String

If Not ValidaDatos Then Exit Sub

Set oNContFunc = New NContFunciones
Set oDocPago = New clsDocPago
Set oCtasIF = New NCajaCtaIF
Set oDoc = New DDocumento

lsTpoIf = Mid(txtBuscarBanco.Text, 1, 2)
lsCtaBanco = Mid(txtBuscarBanco.Text, 18, Len(txtBuscarBanco.Text))
lsSubCuentaIF = oCtasIF.SubCuentaIF(Mid(txtBuscarBanco.Text, 4, 13))

If nTipoDoc = TpoDocCheque Then
    lsDocVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1))
    oDocPago.InicioCheque lsDocNro, True, Mid(txtBuscarBanco.Text, 4, 13), nOperacion, lsPersNombre, "Transferencia de Primas de Seguros", Trim(txtMovDesc.Text), CCur(lblMonto.Caption), gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lblDescBanco.Caption, lblCtaBanco.Caption, lsDocVoucher, True, gsCodAge, , , lsTpoIf, Mid(txtBuscarBanco.Text, 4, 13), lsCtaBanco
    If oDocPago.vbOk Then
        lsFecha = oDocPago.vdFechaDoc
        lsDocNroTmp = oDocPago.vsNroDoc
        lsDocVoucherTmp = oDocPago.vsNroVoucher
    Else
        Exit Sub
    End If
Else
    Do While True
        lsPlanillaNro = InputBox("Ingrese el Nro. de Plantilla", "Planilla de Pagos", lsPlanillaNro)
        If lsPlanillaNro = "" Then Exit Sub
        lsPlanillaNro = Format(lsPlanillaNro, "00000000")
        If oDoc.GetValidaDocProv("", CLng(nTipoDoc), lsPlanillaNro) Then
            MsgBox "Nro. de carta ya ha sido ingresada, verifique..!", vbInformation, "Aviso"
        Else
            lsDocNroTmp = lsPlanillaNro
            lsDocVoucherTmp = ""
            gnMgIzq = 17
            gnMgDer = 0
            gnMgSup = 12
            Exit Do
        End If
    Loop
End If
Set oDoc = Nothing

If MsgBox("¿Esta seguro de realizar la operación?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub

Set oNSeg = New NSeguros
'Call oNSeg.GrabarTransferenciaPrimaSeguros(gdFecSis, gsCodAge, gsCodUser, nOperacion, txtAño.Text + IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0", "") & Trim(Right(cboMes.Text, 2)), _
'                                         nTipoDoc, lsDocNroTmp, lblFecha.Caption, Trim(txtMovDesc.Text), CDbl(lblMonto.Caption), lsCtaContBanco, Mid(txtBuscarBanco.Text, 4, 13), _
'                                         lsTpoIf, lsCtaBanco, lsMovNro, nTipoSeg)
Call oNSeg.GrabarTransferenciaPrimaSeguros(txtFecha, gsCodAge, gsCodUser, nOperacion, txtAño.Text + IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0", "") & Trim(Right(cboMes.Text, 2)), _
                                         nTipoDoc, lsDocNroTmp, Trim(txtMovDesc.Text), CDbl(lblMonto.Caption), lsCtaContBanco, Mid(txtBuscarBanco.Text, 4, 13), _
                                         lsTpoIf, lsCtaBanco, lsMovNro, nTipoSeg)
                                         
Set oNSeg = Nothing

MsgBox "La operación fue realizada", vbInformation, "Aviso"
Set oImp = New NContImprimir
    lsImpre = oImp.ImprimeAsientoContable(lsMovNro, gnLinPage, gnColPage, "TRANSFERENCIA DE PRIMAS DE SEGUROS")
    EnviaPrevio lsImpre & oImpresora.gPrnSaltoPagina, "TRANSFERENCIA DE PRIMAS DE SEGUROS", gnLinPage, False
Set oImp = Nothing
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Realizo la Operación "
        Set objPista = Nothing
        '****
Unload Me
End Sub

Private Sub cmdCancelar_Click()
    bCancela = True
    cboProducto.ListIndex = -1
    cboMes.ListIndex = -1
    txtAño.Text = ""
    txtBuscarBanco.Text = ""
    lblDescBanco.Caption = ""
    lblDescCtaBanco.Caption = ""
    lblTipoPago.Caption = ""
    nTipoDoc = 0
    'lblFecha.Caption = ""
    txtFecha.Text = "__/__/____"
    txtMovDesc.Text = ""
    lblMonto.Caption = ""
    HabilitaControles (False)
    If nOperacion = "421008" Then cboProducto.SetFocus
    bCancela = False
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdExtornar_Click()

Dim oNContFun As NContFunciones
Dim oNCaja As nCajaGeneral
Dim oImp As NContImprimir

Dim lsMovNroExt As String
Dim lbEliminaMov As Boolean
Dim lsImpre As String

On Error GoTo ExtornarErr
    
    Set oNContFun = New NContFunciones
    lsMovNroExt = oNContFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oNContFun = Nothing
    
    If MsgBox("Se va a extornar la operación, Desea continuar?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
    
    Dim oFun As New NContFunciones
    lbEliminaMov = oFun.PermiteModificarAsiento(lblsMovNro.Caption, False)
    
    Set oNCaja = New nCajaGeneral
    If oNCaja.GrabaExtornoMov(gdFecSis, gdFecSis, lsMovNroExt, CLng(lblnMovNro.Caption), nOperacion, txtMovDesc.Text, CCur(lblMonto.Caption), lsMovNroExt, lbEliminaMov, , , , , gbBitCentral) = 0 Then
        If Not lbEliminaMov Then
            Set oImp = New NContImprimir
                lsImpre = oImp.ImprimeAsientoContable(lsMovNroExt, gnLinPage, gnColPage, "EXTORNO TRANSFERENCIA DE PRIMAS DE SEGUROS")
                EnviaPrevio lsImpre & oImpresora.gPrnSaltoPagina, "EXTORNO TRANSFERENCIA DE PRIMAS DE SEGUROS", gnLinPage, False
            Set oImp = Nothing
            
        End If
        MsgBox "El extorno fue realizado", vbInformation, "Aviso"
        
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Realizo el Extorno "
        Set objPista = Nothing
        '****
        Unload Me
        Exit Sub
    End If
    cmdExtornar.Enabled = True

Exit Sub
ExtornarErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub cmdBuscar_Click()

If nOperacion = "421008" Or nOperacion = "422008" Then
    If Not ValidaSeleccionar Then Exit Sub
        nTipoSeg = CInt(Right(Trim(cboProducto.Text), 1))
        Set oOpe = New DOperacion
        txtBuscarBanco.rs = oOpe.GetOpeObj(nOperacion, "2")
        Set oOpe = Nothing
        
        HabilitaControles (True)
        txtBuscarBanco.SetFocus
        'lblFecha.Caption = gdFecSis
        txtFecha = gdFecSis
        Set oDSeg = New DSeguros
        Set rs = oDSeg.RecuperaGeneracionTramaSeguros(CInt(Right(Trim(cboMes.Text), 2)), CInt(txtAño.Text), nTipoSeg, CInt(Mid(nOperacion, 3, 1)))
        Set oDSeg = Nothing
        
        lblMonto.Caption = Format(rs!nPrimaTotal, "#,##0.00")
            
        If nTipoSeg = 1 Then
            txtMovDesc.Text = "Tranferencia de Primas por Comercialización del Seguro de Protección de Tarjeta, del periodo " & Trim(Left(Trim(cboMes.Text), 10)) & "-" & Trim(txtAño.Text)
        ElseIf nTipoSeg = 2 Then
            txtMovDesc.Text = "Tranferencia de Primas por Comercialización del Seguro de Vida y Sepelio, del periodo " & Trim(Left(Trim(cboMes.Text), 10)) & "-" & Trim(txtAño.Text)
        ElseIf nTipoSeg = 3 Then
            txtMovDesc.Text = "Tranferencia de Primas por Comercialización del Seguro Multiriesgo ContraIncendio, del periodo " & Trim(Left(Trim(cboMes.Text), 10)) & "-" & Trim(txtAño.Text)
        ElseIf nTipoSeg = 4 Then
            txtMovDesc.Text = "Tranferencia de Primas por Comercialización del Seguro de Desgravamen, del periodo " & Trim(Left(Trim(cboMes.Text), 10)) & "-" & Trim(txtAño.Text)
        'APRI20180901 ERS061-2018
        ElseIf nTipoSeg = 5 Then
            txtMovDesc.Text = "Tranferencia de Primas por Comercialización del Seguro Multiriesgo Microcrédito, del periodo " & Trim(Left(Trim(cboMes.Text), 10)) & "-" & Trim(txtAño.Text)
        End If
        'END APRI
  Else
  
    If Trim(cboProducto.Text) = "" Then
        MsgBox "Ingrese correctamente el Producto", vbInformation, "Aviso"
        cboProducto.SetFocus
        Exit Sub
    End If
    nTipoSeg = CInt(Right(Trim(cboProducto.Text), 1))
    Set oDSeg = New DSeguros
        Set rs = oDSeg.RecuperaSegurosOpeBancosExtorno(Format(gdFecSis, "yyyymmdd"), CInt(Mid(nOperacion, 3, 1)), nTipoSeg)
    Set oDSeg = Nothing
    If rs.EOF Then
        MsgBox "No se encontraron operaciones de Transferencia de Primas de Seguros para extornar.", vbInformation, "Aviso"
        Exit Sub
    Else
            cmdExtornar.Enabled = True 'APRI20180901 MEJORA
            lblAño.Caption = rs!cAnio
            lblMes.Caption = rs!cMES
            lblCtaBanco.Caption = rs!cCtaBanco
            lblDescBanco.Caption = rs!cBancoDesc
            lblDescCtaBanco.Caption = rs!cCtaBancoDesc
            lblTipoPago.Caption = rs!cTipoPago
            'lblFecha.Caption = Format(rs!dFechaPago, "dd/mm/yyyy")
            txtFecha = Format(rs!dFechaPago, "dd/mm/yyyy")
            txtMovDesc.Text = rs!cMovDesc
            txtMovDesc.Enabled = False
            lblMonto.Caption = Format(rs!nMonto, "#0.00")
            lblsMovNro.Caption = rs!cMovNro
            lblnMovNro.Caption = rs!nMovNro
            Set rs = Nothing

    End If
    
  End If
  
    
End Sub

Private Function ValidaSeleccionar() As Boolean
ValidaSeleccionar = False
If Trim(cboProducto.Text) = "" Then
    MsgBox "Ingrese correctamente el Producto", vbInformation, "Aviso"
    cboProducto.SetFocus
    Exit Function
End If
If Trim(txtAño.Text) = "" Then
    MsgBox "Ingrese correctamente el año", vbInformation, "Aviso"
    txtAño.SetFocus
    Exit Function
End If
If Val(txtAño.Text) < 1900 Or Val(txtAño.Text) > 9972 Then
    MsgBox "Ingrese correctamente el año", vbInformation, "Aviso"
    txtAño.SetFocus
    Exit Function
End If
If Trim(cboMes.Text) = "" Then
    MsgBox "Ingrese correctamente el mes", vbInformation, "Aviso"
    cboMes.SetFocus
    Exit Function
End If

Dim nAnioMesRep As Long
Dim nAnioMesSis As Long
nAnioMesRep = CLng(txtAño.Text & IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0", "") & Trim(Right(cboMes.Text, 2)))
nAnioMesSis = CLng(Right(gdFecSis, 4) & Mid(gdFecSis, 4, 2))
If nAnioMesRep >= nAnioMesSis Then
    MsgBox "El periodo seleccionado debe ser anterior al mes actual", vbInformation, "Aviso"
    Exit Function
End If

Set oDSeg = New DSeguros
Set rs = oDSeg.RecuperaGeneracionTramaSeguros(CInt(Right(Trim(cboMes.Text), 2)), CInt(txtAño.Text), CInt(Right(Trim(cboProducto.Text), 1)), CInt(Mid(nOperacion, 3, 1)))
    If rs!nPrimaTotal <= 0 Then
        MsgBox "Trama Seleccionada no ha sido declarado por el Departamento de GPS.", vbInformation, "Aviso"
        Exit Function
    End If
    If Not oDSeg.RecuperaSegurosOpeBancos(txtAño.Text, IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0", "") & Trim(Right(cboMes.Text, 2)), CInt(Mid(nOperacion, 3, 1)), CInt(Right(Trim(cboProducto.Text), 1))).EOF Then
        MsgBox "El Periodo Seleccionado ya fue Ejecutado", vbInformation, "Aviso"
        Exit Function
    End If
Set oDSeg = Nothing
Set rs = Nothing
ValidaSeleccionar = True
End Function

Private Sub txtAño_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    End If
End Sub

Private Sub txtBuscarBanco_EmiteDatos()
Dim oCtaIf As NCajaCtaIF
Dim oOpe As DOperacion
Set oOpe = New DOperacion
Set oCtaIf = New NCajaCtaIF

lblDescBanco = oCtaIf.NombreIF(Mid(txtBuscarBanco, 4, 13))
lblDescCtaBanco = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscarBanco, 18, 10)) + " " + txtBuscarBanco.psDescripcion
lsCtaContBanco = oOpe.EmiteOpeCta(nOperacion, "H", , txtBuscarBanco.Text, ObjEntidadesFinancieras)
    If lsCtaContBanco = "" And Not bCancela Then
        MsgBox "Cuentas Contables no determinadas Correctamente" & Chr(13) & "consulte con el Dpto. de Contabilidad", vbInformation, "Aviso"
        txtBuscarBanco.Text = ""
        lblDescBanco.Caption = ""
        lblDescCtaBanco.Caption = ""
        lblTipoPago.Caption = ""
        nTipoDoc = 0
        Exit Sub
    End If
    If Mid(txtBuscarBanco, 4, 13) = "1090100824640" Then
        nTipoDoc = TpoDocCarta
        lblTipoPago.Caption = "Transferencia"
    Else
        nTipoDoc = TpoDocCheque
        lblTipoPago.Caption = "Cheque"
    End If
txtMovDesc.SetFocus
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
    If lsCtaContBanco = "" Then
        MsgBox "No seleccionó la cuenta de la Institución Financiera", vbInformation, "Aviso"
        txtBuscarBanco.SetFocus
        Exit Function
    End If
    If nTipoDoc = 0 Then
        MsgBox "No se especificó el Tipo de Pago", vbInformation, "Aviso"
        txtBuscarBanco.SetFocus
        Exit Function
    End If
    If Trim(txtMovDesc.Text) = "" Then
        MsgBox "Debe ingresar la glosa", vbInformation, "Aviso"
        txtMovDesc.SetFocus
        Exit Function
    End If
    
ValidaDatos = True
End Function
