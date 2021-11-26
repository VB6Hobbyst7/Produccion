VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmServCobDebitoAutoServCred 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   Icon            =   "frmServCobDebitoAutoServCred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabServ 
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4048
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Agregar Servicio"
      TabPicture(0)   =   "frmServCobDebitoAutoServCred.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblConvenio"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblCodigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDepositante"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblConvNombre"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblId"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblMonedaConv"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "spnDia2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdBuscarConv"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdBuscarCod"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "spnDia1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin SICMACT.uSpinner spnDia1 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Max             =   31
         Min             =   1
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.CommandButton cmdBuscarCod 
         Caption         =   "..."
         Height          =   300
         Left            =   3500
         TabIndex        =   1
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdBuscarConv 
         Caption         =   "..."
         Height          =   300
         Left            =   3500
         TabIndex        =   0
         Top             =   600
         Width           =   375
      End
      Begin SICMACT.uSpinner spnDia2 
         Height          =   315
         Left            =   3240
         TabIndex        =   3
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Max             =   31
         Min             =   1
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.Label lblMonedaConv 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4200
         TabIndex        =   27
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblId 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4200
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblConvNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4200
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblDepositante 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   20
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   19
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblConvenio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   18
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Dia 2: "
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   1720
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Dia 1:"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   1720
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Dia Pago:"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Depositante:"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1000
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Convenio:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   640
         Width           =   735
      End
   End
   Begin SICMACT.EditMoney txtMontoMax 
      Height          =   315
      Left            =   1080
      TabIndex        =   25
      Top             =   2550
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
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
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   2760
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   4080
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTabCred 
      Height          =   2175
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3836
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Crédito"
      TabPicture(0)   =   "frmServCobDebitoAutoServCred.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTitular"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSaldoCap"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblVencCuota"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCuenta"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdBuscarCred"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.CommandButton cmdBuscarCred 
         Caption         =   "..."
         Height          =   325
         Left            =   4080
         TabIndex        =   15
         ToolTipText     =   "Busca cliente por nombre, documento o codigo"
         Top             =   765
         Width           =   375
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   720
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblVencCuota 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1800
         TabIndex        =   24
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblSaldoCap 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   720
         TabIndex        =   21
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblTitular 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   720
         TabIndex        =   17
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label12 
         Caption         =   "Titular:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1380
         Width           =   735
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Monto Max:"
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Top             =   2595
      Width           =   855
   End
End
Attribute VB_Name = "frmServCobDebitoAutoServCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'** Nombre : frmSegTarjetaAfiliacion
'** Descripción : Formulario para afiliar seleccionar una cuenta para asociar al débito
'**               automático para pagos de servicios de recaudo o créditos creado segun
'**               TI-ERS144-2014
'** Creación : JUEZ, 20150130 09:00:00 AM
'****************************************************************************************

Option Explicit

Dim oCredito As COMDCredito.DCOMCredito
Dim fsTipoDebito As CaptacServDebitoAutomaticoTipo
Dim MatServCred As Variant
Dim RS As ADODB.Recordset

Public Sub inicia(ByVal pnTipoDebito As CaptacServDebitoAutomaticoTipo, ByRef pMatServCred As Variant)
    fsTipoDebito = pnTipoDebito
    If fsTipoDebito = gServConvenio Then
        SSTabServ.Visible = True
        SSTabCred.Visible = False
        Me.Caption = "Agregar Servicio"
    ElseIf fsTipoDebito = gServCredito Then
        SSTabServ.Visible = False
        SSTabCred.Visible = True
        Me.Caption = "Agregar Crédito"
        txtCuenta.CMAC = "109"
        txtCuenta.Age = gsCodAge
    End If
    ReDim MatServCred(0, 0)
    Me.Show 1
    pMatServCred = MatServCred
End Sub

Private Sub CmdAceptar_Click()
If ValidaDatos Then
    If fsTipoDebito = gServConvenio Then
        ReDim MatServCred(0, 7)
        MatServCred(0, 0) = lblConvNombre.Caption
        MatServCred(0, 1) = lblConvenio.Caption
        MatServCred(0, 2) = lblCodigo.Caption
        MatServCred(0, 3) = spnDia1.valor
        MatServCred(0, 4) = spnDia2.valor
        MatServCred(0, 5) = txtMontoMax.Text
        MatServCred(0, 6) = lblId.Caption
        MatServCred(0, 7) = lblMonedaConv.Caption
    ElseIf fsTipoDebito = gServCredito Then
        ReDim MatServCred(0, 4)
        MatServCred(0, 0) = txtCuenta.NroCuenta
        MatServCred(0, 1) = lblSaldoCap.Caption
        MatServCred(0, 2) = lblTitular.Caption
        MatServCred(0, 3) = txtMontoMax.Text
        MatServCred(0, 4) = lblVencCuota.Caption
    End If
    Unload Me
End If
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If fsTipoDebito = gServConvenio Then
    If lblConvenio.Caption = "" Then
        MsgBox "Ingresar correctamente el Convenio", vbInformation, "Aviso"
        cmdBuscarConv.SetFocus
        Exit Function
    End If
    If lblCodigo.Caption = "" Then
        MsgBox "Ingresar correctamente el Código", vbInformation, "Aviso"
        cmdBuscarCod.SetFocus
        Exit Function
    End If
    If CInt(spnDia1.valor) = CInt(spnDia2.valor) Then
        MsgBox "El dia de pago 1 no puede ser igual al dia de pago 2", vbInformation, "Aviso"
        spnDia2.SetFocus
        Exit Function
    End If
ElseIf fsTipoDebito = gServCredito Then
    If lblTitular.Caption = "" Then
        MsgBox "Cagar correctamente el crédito", vbInformation, "Aviso"
        txtCuenta.SetFocusCuenta
        Exit Function
    End If
End If
If val(txtMontoMax.Text) = 0 Then
    MsgBox "Ingrese el monto máximo a debitar por el servicio", vbInformation, "Aviso"
    txtMontoMax.SetFocus
    Exit Function
End If
ValidaDatos = True
End Function

Private Sub cmdBuscarCod_Click()
If lblConvenio.Caption = "" Then
    MsgBox "Ingresar correctamente el Convenio", vbInformation, "Aviso"
    Exit Sub
End If
Set RS = New Recordset
Set RS = frmBuscarUsuarioConvenio.Inicio(Trim(lblConvenio.Caption))
If Not RS Is Nothing Then
    lblCodigo.Caption = RS!cCodCliente
    lblDepositante.Caption = RS!cNomCliente
    lblId.Caption = RS!cId
    lblMonedaConv.Caption = IIf(RS!cmoneda = "SOLES", 1, 2)
    spnDia1.SetFocus
Else
    lblCodigo.Caption = ""
    lblDepositante.Caption = ""
    lblId.Caption = ""
    lblMonedaConv.Caption = ""
End If
End Sub

Private Sub cmdBuscarConv_Click()
Set RS = New ADODB.Recordset
Set RS = frmBuscarConvenio.Inicio
If Not RS Is Nothing Then
    If RS!nTipoConvenio = Convenio_VCM Then
        MsgBox "El Servicio de Debito Automático no está disponible para Convenio MYPE", vbInformation, "Aviso"
        lblConvenio.Caption = ""
        lblConvNombre.Caption = ""
    Else
        lblConvenio.Caption = RS!cCodConvenio
        lblConvNombre.Caption = RS!cPersNombre
        cmdBuscarCod.SetFocus
    End If
Else
    lblConvenio.Caption = ""
    lblConvNombre.Caption = ""
End If
lblCodigo.Caption = ""
lblDepositante.Caption = ""
lblId.Caption = ""
lblMonedaConv.Caption = ""
End Sub

Private Sub cmdBuscarCred_Click()
Dim oPers As COMDPersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Call FrmVerCredito.Inicio(oPers.sPersCod, , , True, txtCuenta)
        txtCuenta.SetFocusCuenta
    End If
    Set oPers = Nothing
End Sub

Private Sub cmdCancelar_Click()
    Set MatServCred = Nothing
    ReDim MatServCred(0, 0)
    Unload Me
End Sub

Private Sub spnDia1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spnDia2.SetFocus
    End If
End Sub

Private Sub spnDia2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMontoMax.SetFocus
    End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
If KeyAscii = 13 Then
    If Len(txtCuenta.NroCuenta) = 18 Then
        Set oCredito = New COMDCredito.DCOMCredito
        Set RS = oCredito.RecuperaDatosComunes(txtCuenta.NroCuenta)
        If Not RS.EOF And Not RS.BOF Then
            If gdFecSis > RS!dVenc_Cuota Then
                MsgBox "El Credito se encuentra con Cuota Vencida, solo se puede afiliar crédito que estén al día", vbInformation, "Aviso"
                lblTitular.Caption = ""
                lblSaldoCap.Caption = ""
                lblVencCuota.Caption = ""
                Exit Sub
            End If
            lblTitular.Caption = RS!cTitular
            lblSaldoCap.Caption = RS!nSaldo
            lblVencCuota.Caption = RS!dVenc_Cuota
            SSTabCred.Enabled = False
            txtMontoMax.SetFocus
        Else
            MsgBox "Crédito no existe", vbInformation, "Aviso"
            lblTitular.Caption = ""
            lblSaldoCap.Caption = ""
            lblVencCuota.Caption = ""
        End If
    Else
        MsgBox "Ingresar correctamente el crédito", vbInformation, "Aviso"
        lblTitular.Caption = ""
        lblSaldoCap.Caption = ""
        lblVencCuota.Caption = ""
    End If
End If
End Sub

Private Sub txtMontoMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub
