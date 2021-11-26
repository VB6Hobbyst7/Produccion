VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSegTarjRetPagoBanco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pago Seguro Tarjeta"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   Icon            =   "frmSegTarjRetPagoBanco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5400
      TabIndex        =   24
      Top             =   5760
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
      Left            =   4080
      TabIndex        =   21
      Top             =   5760
      Width           =   1170
   End
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
      Left            =   4080
      TabIndex        =   5
      Top             =   5760
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
      Left            =   5400
      TabIndex        =   6
      Top             =   5760
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Origen"
      TabPicture(0)   =   "frmSegTarjRetPagoBanco.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraPeriodo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraCtaIF"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraOperacion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
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
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   6255
         Begin VB.TextBox txtMovDesc 
            Height          =   720
            Left            =   720
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   840
            Width           =   5370
         End
         Begin VB.Label lblLabelTot 
            Caption         =   "Total :"
            Height          =   285
            Left            =   4300
            TabIndex        =   30
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label lblTotal 
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
            Left            =   4800
            TabIndex        =   29
            Top             =   1650
            Width           =   1290
         End
         Begin VB.Label lblLabelNC 
            Caption         =   "Nota Cargo :"
            Height          =   285
            Left            =   2280
            TabIndex        =   28
            Top             =   1680
            Width           =   945
         End
         Begin VB.Label lblNotaCargo 
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
            Left            =   3240
            TabIndex        =   27
            Top             =   1650
            Width           =   810
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   20
            Top             =   1650
            Width           =   1290
         End
         Begin VB.Label lblFecha 
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
            Left            =   720
            TabIndex        =   19
            Top             =   340
            Width           =   1290
         End
         Begin VB.Label Label2 
            Caption         =   "Monto :"
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   1680
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Glosa :"
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   705
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha :"
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   705
         End
      End
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
         TabIndex        =   12
         Top             =   1200
         Width           =   6255
         Begin Sicmact.TxtBuscar txtBuscarBanco 
            Height          =   330
            Left            =   120
            TabIndex        =   3
            Top             =   375
            Width           =   2070
            _ExtentX        =   3651
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
            TabIndex        =   9
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
            TabIndex        =   18
            Top             =   375
            Width           =   2010
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Pago :"
            Height          =   285
            Left            =   120
            TabIndex        =   17
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
            Left            =   2205
            TabIndex        =   7
            Top             =   375
            Width           =   3930
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
            TabIndex        =   8
            Top             =   750
            Width           =   6015
         End
      End
      Begin VB.Frame FraPeriodo 
         Caption         =   " Periodo "
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
         TabIndex        =   10
         Top             =   360
         Width           =   3855
         Begin VB.ComboBox cboMes 
            Height          =   315
            ItemData        =   "frmSegTarjRetPagoBanco.frx":0326
            Left            =   120
            List            =   "frmSegTarjRetPagoBanco.frx":0330
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtAño 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   1
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmdSeleccionar 
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
            Left            =   2520
            TabIndex        =   2
            Top             =   240
            Width           =   1170
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
            Left            =   1560
            TabIndex        =   23
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
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1290
         End
      End
   End
End
Attribute VB_Name = "frmSegTarjRetPagoBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmSegTarjRetPagoBanco
'** Descripción : Formulario para registrar el retiro por pago de seguros de tarjetas y su
'**               extorno creado segun TI-ERS029-2013
'** Creación : JUEZ, 20140711 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim oDSeg As DSeguros
Dim oOpe As DOperacion
Dim nOperacion As OpeCGOpeBancos
Dim oNContFunc As NContFunciones
Dim lsCtaContBanco As String
Dim nTipoDoc As TpoDoc
Dim rs As ADODB.Recordset
Dim objPista As COMManejador.Pista 'ARLO20170217

Public Sub Inicio(ByVal pnOperacion As OpeCGOpeBancos)
nOperacion = pnOperacion

If nOperacion = gOpeCGOpeBancosRetPagoSegTarjMN Then
    Me.Caption = "Retiro por Pago Seguro Tarjeta"
    txtBuscarBanco.psRaiz = "Cuentas de Bancos"
    Set oOpe = New DOperacion
    txtBuscarBanco.rs = oOpe.GetOpeObj(pnOperacion, "2")
    Set oOpe = Nothing
    CargaComboConstante 1010, cboMes
    HabilitaControles (False)
    HabilitaControlesExtorno False
Else
    Me.Caption = "Extorno MN: Extorno Pago Seguro Tarjeta"
    Set oDSeg = New DSeguros
        Set rs = oDSeg.RecuperaSegTarjetaOpeBancosExtorno(Format(gdFecSis, "yyyymmdd"), 1)
    Set oDSeg = Nothing
    If rs.EOF Then
        MsgBox "No se encontraron operaciones de Retiro por Pago de Seguro de Tarjetas para extornar", vbInformation, "Aviso"
        Exit Sub
    Else
        Dim rsDep As ADODB.Recordset
        Set oDSeg = New DSeguros
            Set rsDep = oDSeg.RecuperaSegTarjetaOpeBancos(rs!cAnio, rs!cMesNro, 2)
        Set oDSeg = Nothing
        If Not rsDep.EOF Then
            MsgBox "El ingreso por comisión a Caja Maynas para el periodo " & rs!cMES & "/" & rs!cAnio & " ya fue realizado. No se podrá realizar el extorno", vbInformation, "Aviso"
            Exit Sub
        Else
            lblAño.Caption = rs!cAnio
            lblMes.Caption = rs!cMES
            lblCtaBanco.Caption = rs!cCtaBanco
            lblDescBanco.Caption = rs!cBancoDesc
            lblDescCtaBanco.Caption = rs!cCtaBancoDesc
            lblTipoPago.Caption = rs!cTipoPago
            lblFecha.Caption = Format(rs!dFechaPago, "dd/mm/yyyy")
            txtMovDesc.Text = rs!cMovDesc
            txtMovDesc.Enabled = False
            lblMonto.Caption = Format(rs!nMonto, "#0.00")
            lblsMovNro.Caption = rs!cMovNro
            lblnMovNro.Caption = rs!nMovNro
            Set rs = Nothing
        End If
    End If
    HabilitaControlesExtorno True
End If

Me.Show 1
End Sub

Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
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
cmdSeleccionar.Visible = Not pbHabilitaExt
lblCtaBanco.Visible = pbHabilitaExt
txtBuscarBanco.Visible = Not pbHabilitaExt
cmdExtornar.Visible = pbHabilitaExt
cmdAceptar.Visible = Not pbHabilitaExt
cmdCerrar.Visible = pbHabilitaExt
cmdCancelar.Visible = Not pbHabilitaExt
lblLabelNC.Visible = Not pbHabilitaExt
lblNotaCargo.Visible = Not pbHabilitaExt
lblLabelTot.Visible = Not pbHabilitaExt
lblTotal.Visible = Not pbHabilitaExt
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
    'oDocPago.InicioCheque lsDocNro, True, Mid(txtBuscarBanco.Text, 4, 13), nOperacion, lsPersNombre, "Retiro Pago Seguro Tarjeta", Trim(txtMovDesc.Text), CCur(lblMonto.Caption), gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lblDescBanco.Caption, lblCtaBanco.Caption, lsDocVoucher, True, gsCodAge, , , lsTpoIf, Mid(txtBuscarBanco.Text, 4, 13), lsCtaBanco
    oDocPago.InicioCheque lsDocNro, True, Mid(txtBuscarBanco.Text, 4, 13), nOperacion, lsPersNombre, "Retiro Pago Seguro Tarjeta", Trim(txtMovDesc.Text), CCur(lblTotal.Caption), gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lblDescBanco.Caption, lblCtaBanco.Caption, lsDocVoucher, True, gsCodAge, , , lsTpoIf, Mid(txtBuscarBanco.Text, 4, 13), lsCtaBanco
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
Call oNSeg.GrabarRetiroPagoSeguroTarjeta(gdFecSis, gsCodAge, gsCodUser, nOperacion, txtAño.Text + IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0", "") & Trim(Right(cboMes.Text, 2)), _
                                         nTipoDoc, lsDocNroTmp, lblFecha.Caption, Trim(txtMovDesc.Text), CDbl(lblTotal.Caption), lsCtaContBanco, Mid(txtBuscarBanco.Text, 4, 13), _
                                         lsTpoIf, lsCtaBanco, lsMovNro)
                                         'JUEZ 20150510 Se cambió lblMonto por lblTotal
Set oNSeg = Nothing

MsgBox "La operación fue realizada", vbInformation, "Aviso"
Set oImp = New NContImprimir
    lsImpre = oImp.ImprimeAsientoContable(lsMovNro, gnLinPage, gnColPage, "RETIRO POR PAGO SEGURO DE TARJETAS")
    EnviaPrevio lsImpre & oImpresora.gPrnSaltoPagina, "RETIRO POR PAGO SEGURO DE TARJETAS", gnLinPage, False
Set oImp = Nothing
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Realizo la Operación "
        Set objPista = Nothing
        '****
Unload Me
End Sub

Private Sub cmdCancelar_Click()
    cboMes.ListIndex = -1
    txtAño.Text = ""
    txtBuscarBanco.Text = ""
    lblDescBanco.Caption = ""
    lblDescCtaBanco.Caption = ""
    lblTipoPago.Caption = ""
    nTipoDoc = 0
    lblFecha.Caption = ""
    txtMovDesc.Text = ""
    lblMonto.Caption = ""
    lblNotaCargo.Caption = ""
    lblTotal.Caption = ""
    HabilitaControles (False)
    If nOperacion = gOpeCGOpeBancosRetPagoSegTarjMN Then cboMes.SetFocus
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
                lsImpre = oImp.ImprimeAsientoContable(lsMovNroExt, gnLinPage, gnColPage, "EXTORNO RETIRO POR PAGO SEGURO DE TARJETAS")
                EnviaPrevio lsImpre & oImpresora.gPrnSaltoPagina, "EXTORNO RETIRO POR PAGO SEGURO DE TARJETAS", gnLinPage, False
            Set oImp = Nothing
            
        End If
        'JUEZ 20150510 ******************************************
        Set oDSeg = New DSeguros
            oDSeg.ExtornaSegTarjetaAnulaDevPendiente CLng(lblnMovNro.Caption)
        Set oDSeg = Nothing
        'END JUEZ ***********************************************
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

Private Sub cmdSeleccionar_Click()
Dim nNotaCargo As Double 'JUEZ 20150510
nNotaCargo = 0 'JUEZ 20150510

If Not ValidaSeleccionar Then Exit Sub
    HabilitaControles (True)
    txtBuscarBanco.SetFocus
    lblFecha.Caption = gdFecSis
    Set oDSeg = New DSeguros
    Set rs = oDSeg.RecuperaGeneracionTramaSegTarjeta(CInt(Right(Trim(cboMes.Text), 2)), CInt(txtAño.Text), "AF")
    Set oDSeg = Nothing
    lblMonto.Caption = Format(rs!nPrimaTotal, "#0.00")
    'JUEZ 20150510 ******************************************
    Set oDSeg = New DSeguros
    Set rs = oDSeg.RecuperaSegTarjetaAnulaDevPend(True, False)
    Set oDSeg = Nothing
    If Not rs.EOF And Not rs.BOF Then
        Do While Not rs.EOF
            nNotaCargo = nNotaCargo + CDbl(rs!nMontoCom)
            rs.MoveNext
        Loop
    End If
    
    lblNotaCargo.Caption = Format(nNotaCargo, "#0.00")
    lblTotal.Caption = Format(CDbl(lblMonto.Caption) - nNotaCargo, "#0.00")
    'END JUEZ ***********************************************
End Sub

Private Function ValidaSeleccionar() As Boolean
ValidaSeleccionar = False
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
'JUEZ 20150123 ************************************************************
Dim nAnioMesRep As Long
Dim nAnioMesSis As Long
nAnioMesRep = CLng(txtAño.Text & IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0", "") & Trim(Right(cboMes.Text, 2)))
nAnioMesSis = CLng(Right(gdFecSis, 4) & Mid(gdFecSis, 4, 2))
'If CInt(Trim(Right(cboMes.Text, 2))) >= CInt(Mid(gdFecSis, 4, 2)) Or val(txtAño.Text) > Right(gdFecSis, 4) Then
If nAnioMesRep >= nAnioMesSis Then
'END JUEZ *****************************************************************
    MsgBox "El periodo seleccionado debe ser anterior al mes actual", vbInformation, "Aviso"
    Exit Function
End If

Set oDSeg = New DSeguros
    If oDSeg.RecuperaGeneracionTramaSegTarjeta(CInt(Right(Trim(cboMes.Text), 2)), CInt(txtAño.Text), "AF").EOF Then
        MsgBox "No se ha generado la trama del periodo seleccionado", vbInformation, "Aviso"
        Exit Function
    End If
    If Not oDSeg.RecuperaSegTarjetaOpeBancos(txtAño.Text, IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0", "") & Trim(Right(cboMes.Text, 2)), 1).EOF Then
        MsgBox "El Retiro por Pago de Seguro de Tarjeta para este periodo ya fue registrado", vbInformation, "Aviso"
        Exit Function
    End If
Set oDSeg = Nothing

ValidaSeleccionar = True
End Function

Private Sub txtAño_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSeleccionar.SetFocus
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
    If lsCtaContBanco = "" Then
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
