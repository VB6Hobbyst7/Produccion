VERSION 5.00
Begin VB.Form frmOpeComisionDiversasAho 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   Icon            =   "frmOpeComisionDiversasAho.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   360
      Left            =   3840
      TabIndex        =   5
      Top             =   5160
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   6120
      TabIndex        =   7
      Top             =   5160
      Width           =   1050
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   4920
      TabIndex        =   6
      Top             =   5160
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos "
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7095
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
         Left            =   80
         TabIndex        =   18
         Top             =   2400
         Width           =   6945
         Begin VB.ComboBox cboTransferMoneda 
            Height          =   315
            Left            =   855
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   255
            Width           =   2880
         End
         Begin VB.CommandButton cmdTranfer 
            Height          =   350
            Left            =   3840
            Picture         =   "frmOpeComisionDiversasAho.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   645
            Width           =   475
         End
         Begin VB.TextBox txtTransferGlosa 
            Appearance      =   0  'Flat
            Height          =   555
            Left            =   855
            MaxLength       =   255
            TabIndex        =   19
            Top             =   1410
            Width           =   5865
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
            Visible         =   0   'False
            Width           =   1665
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
            TabIndex        =   33
            Top             =   2070
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label lblEtiMonTra 
            AutoSize        =   -1  'True
            Caption         =   "Monto Transacción"
            Height          =   195
            Left            =   840
            TabIndex        =   32
            Top             =   2100
            Visible         =   0   'False
            Width           =   1380
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
            TabIndex        =   31
            Top             =   615
            Width           =   750
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
            TabIndex        =   30
            Top             =   255
            Width           =   750
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
            TabIndex        =   29
            Top             =   1020
            Width           =   5865
         End
         Begin VB.Label lbltransferN 
            AutoSize        =   -1  'True
            Caption         =   "Nro Doc :"
            Height          =   195
            Left            =   45
            TabIndex        =   28
            Top             =   720
            Width           =   690
         End
         Begin VB.Label lbltransferBcol 
            AutoSize        =   -1  'True
            Caption         =   "Banco :"
            Height          =   195
            Left            =   90
            TabIndex        =   27
            Top             =   1110
            Width           =   555
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
            TabIndex        =   26
            Top             =   645
            Width           =   2880
         End
         Begin VB.Label lblTransferMoneda 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   45
            TabIndex        =   25
            Top             =   315
            Width           =   585
         End
         Begin VB.Label lblTransferGlosa 
            AutoSize        =   -1  'True
            Caption         =   "Glosa :"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1410
            Width           =   495
         End
         Begin VB.Label lblTTCC 
            Caption         =   "TCC"
            Height          =   285
            Left            =   5370
            TabIndex        =   23
            Top             =   270
            Width           =   390
         End
         Begin VB.Label Label2 
            Caption         =   "TCV"
            Height          =   285
            Left            =   5355
            TabIndex        =   22
            Top             =   630
            Width           =   390
         End
      End
      Begin VB.CommandButton cmdBCodCom 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   290
         Left            =   2430
         TabIndex        =   0
         Top             =   382
         Width           =   350
      End
      Begin VB.ComboBox cboTipoPago 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   1850
      End
      Begin SICMACT.TxtBuscar TxtBCodPers 
         Height          =   330
         Left            =   960
         TabIndex        =   2
         Top             =   840
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
         CMAC            =   "109"
      End
      Begin VB.Label lblDOI 
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
         Height          =   315
         Left            =   5520
         TabIndex        =   17
         Top             =   2160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblCodCom 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
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
         Height          =   330
         Left            =   960
         TabIndex        =   16
         Top             =   360
         Width           =   1480
      End
      Begin VB.Label lblMontoCom 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   4720
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblMonedaCom 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   4080
         TabIndex        =   14
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Monto:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   1365
         Width           =   615
      End
      Begin VB.Label lblTipoPago 
         Caption         =   "Tipo Pago:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1370
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Persona:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblPersNombre 
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
         Height          =   315
         Left            =   2880
         TabIndex        =   10
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Comisión:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   420
         Width           =   735
      End
      Begin VB.Label lblComDesc 
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
         Height          =   330
         Left            =   2880
         TabIndex        =   8
         Top             =   360
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmOpeComisionDiversasAho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmOpeComisionDiversasAho
'** Descripción : Formulario para pagos de gastos y comisiones de Ahorros según TI-ERS097-2013
'** Creación : JUEZ, 20130829 09:00:00 AM
'*****************************************************************************************************

Option Explicit

Dim fnTipoCobro As Integer
Dim rs As ADODB.Recordset
Dim sNumTarj As String
Dim fsParProd As String 'JUEZ 20150930
Dim loVistoElectronico As frmVistoElectronico 'CTI5 ERS0112020
Dim lnTransferSaldo As Currency 'CTI7 OPEv2
Dim fsPersCodTransfer As String 'CTI7 OPEv2
Dim fnMovNroRVD As Long 'CTI7 OPEv2
Dim lnMovNroTransfer As Long 'CTI7 OPEv2

'Public Sub Inicia()
Public Sub Inicia(ByVal psParProd As String) 'JUEZ 20150930
    fsParProd = psParProd 'JUEZ 20150930
    Me.Caption = "Pago de Comisiones y Gastos de " & IIf(fsParProd = "A", "Ahorros", "Créditos") 'JUEZ 20150930
    'JUEZ 20150930 *************************
    'If fsParProd = "C" Then
    '    lblTipoPago.Visible = False
    '    cboTipoPago.Visible = False
    '    Me.Frame1.Height = 1935
    '    Me.cmdGrabar.Top = 2160
    '    Me.cmdLimpiar.Top = 2160
    '    Me.CmdSalir.Top = 2160
    '    Me.Height = 3045
    'End If
    'END JUEZ ******************************
    fnTipoCobro = 0
    sNumTarj = ""
    'CargaTipoPago
    CargaControles
    HabilitaControles (False)
    txtCuenta.Visible = False
    Me.Show 1
End Sub

'COMENTADO POR CTI5 ERS0112020
'Private Sub CargaTipoPago()
'    Dim rsConst As New ADODB.Recordset
'    Dim clsGen As New COMDConstSistema.DCOMGeneral
'    Set rsConst = clsGen.GetConstante(2037)
'    Set clsGen = Nothing
'
'    cboTipoPago.Clear
'    While Not rsConst.EOF
'        cboTipoPago.AddItem rsConst.Fields(0) & space(100) & rsConst.Fields(1)
'        rsConst.MoveNext
'    Wend
'End Sub

Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
    TxtBCodPers.Enabled = pbHabilita
    cboTipoPago.Enabled = pbHabilita
    txtCuenta.Enabled = pbHabilita
    cmdGrabar.Enabled = pbHabilita
    cmdLimpiar.Enabled = pbHabilita
End Sub

'Private Sub cboTipoPago_Change()
'    'If Trim(Right(cboTipoPago.Text, 2)) = "1" Then 'COMENTADO POR CTI5 ERS0112020
'    If Trim(Right(cboTipoPago.Text, 2)) = gColocTipoPagoCargoCta Then 'CTI5 ERS0112020
'        txtCuenta.Visible = False
'        txtCuenta.NroCuenta = ""
'    Else
'        txtCuenta.Visible = True
'        txtCuenta.CMAC = "109"
'        'txtCuenta.Age = gsCodAge
'        txtCuenta.Prod = gCapAhorros
'    End If
'End Sub

'CTI5 ERS0112020 *********************
Private Sub cboTipoPago_Click()
    'COMENTADO POR CTI5 ERS0112020
    'If Trim(Right(cboTipoPago.Text, 2)) = "1" Then
    '    txtCuenta.Visible = False
    '    txtCuenta.NroCuenta = ""
    'Else
    '    txtCuenta.Visible = True
    '    txtCuenta.NroCuenta = ""
    '    txtCuenta.CMAC = "109"
    '    'txtCuenta.Age = gsCodAge
    '    txtCuenta.Prod = gCapAhorros
    'End If
    EstadoFormaPago IIf(cboTipoPago.ListIndex = -1, -1, CInt(Trim(Right(IIf(cboTipoPago.Text = "", "-1", cboTipoPago.Text), 10))))
    If cboTipoPago.ListIndex <> -1 Then
        If CInt(Trim(Right(cboTipoPago.Text, 10))) = gColocTipoPagoVoucher Then
     
            Dim oformVou As New frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIF As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
                        
            cboTransferMoneda.Enabled = False
            Me.fraTranferecia.Enabled = True
            cboTransferMoneda.ListIndex = IndiceListaCombo(cboTransferMoneda, 1)
            cboTransferMoneda.Enabled = False
                        
            lnTipMot = 17 ' Cancelacion Credito Pignoraticio
            'oformVou.iniciarFormularioDeposito CInt(Mid(txtCuenta.NroCuenta, 9, 1)), lnTipMot, sGlosa, sIF, sVaucher, nMontoVoucher, sPersCod, nMovNroRVD, nMovNroRVDPend, sNombre, sDireccion, sDocumento, txtCuenta.NroCuenta
            'If Len(sVaucher) = 0 Then Exit Sub
            'LblNumDoc.Caption = Trim(Mid(sVaucher, 1, Len(sVaucher) - 10))
            'cmdGrabar.Enabled = True
            

            cmdGrabar.Enabled = True
            EnfocaControl cmdGrabar
        ElseIf CInt(Trim(Right(cboTipoPago.Text, 10))) = gColocTipoPagoCargoCta Then
            Dim sCuenta As String
            
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gAhoCargoCobroComDiversasAho), sNumTarj, 232)
            If Val(Mid(sCuenta, 6, 3)) <> "232" And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If sCuenta <> "" Then
                txtCuenta.NroCuenta = sCuenta
                txtCuenta.SetFocusCuenta
            End If
            If Len(sCuenta) = 18 Then
                If CInt(Mid(sCuenta, 9, 1)) <> CInt(Mid(txtCuenta.NroCuenta, 9, 1)) Then
                    MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la cuenta a amortizar.", vbOKOnly + vbInformation, App.Title
                End If
            End If
            If Len(sCuenta) = 0 Then
                txtCuenta.EnabledAge = True
                txtCuenta.EnabledCta = True
                txtCuenta.SetFocusAge
                Exit Sub
            End If
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.Enabled = False
            'AsignaValorITF
            cmdGrabar.Enabled = True
            cmdGrabar.SetFocus
            
        End If
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
Private Sub EstadoFormaPago(ByVal nFormaPago As Integer)
    'LblNumDoc.Caption = ""
    txtCuenta.NroCuenta = ""
    Select Case nFormaPago
        Case -1
            'LblNumDoc.Visible = False
            'lblNroDocumento.Visible = False
            txtCuenta.Visible = False
            cmdGrabar.Enabled = True
            fraTranferecia.Enabled = False
            Call IniciarControlesFormaPago
        Case gColocTipoPagoEfectivo
            txtCuenta.Visible = False
            'LblNumDoc.Visible = False
            'lblNroDocumento.Visible = False
            'lblNroDocumento.Visible = False
            cmdGrabar.Enabled = True
            fraTranferecia.Enabled = False
            Call IniciarControlesFormaPago
        Case gColocTipoPagoCargoCta
            'LblNumDoc.Visible = False
            'lblNroDocumento.Visible = False
            txtCuenta.Visible = True
            txtCuenta.Enabled = True
            txtCuenta.CMAC = gsCodCMAC
            txtCuenta.Prod = Trim(Str(gCapAhorros))
            cmdGrabar.Enabled = False
            fraTranferecia.Enabled = False
            Call IniciarControlesFormaPago
        Case gColocTipoPagoVoucher
            'LblNumDoc.Visible = True
            'lblNroDocumento.Visible = True
            txtCuenta.Visible = False
            cmdGrabar.Enabled = False
            fraTranferecia.Enabled = True
    End Select
End Sub
Private Function ValidaFormaPago() As Boolean
Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento
    ValidaFormaPago = False
    If cboTipoPago.ListIndex = -1 Then
        MsgBox "No se ha seleccionado la forma de pago. Verifique.", vbInformation, "¡Aviso!"
        EnfocaControl cboTipoPago
        Exit Function
    End If
    If CInt(Trim(Right(cboTipoPago.Text, 10))) = gColocTipoPagoVoucher And Len(Trim(lblTrasferND.Caption)) = 0 Then
        MsgBox "No se ha seleccionado el voucher correctamente. Verifique.", vbInformation, "¡Aviso!"
        EnfocaControl cboTipoPago
        Exit Function
    End If

    If CInt(Trim(Right(cboTipoPago.Text, 10))) = gColocTipoPagoVoucher And Len(Trim(lblTrasferND.Caption)) > 0 _
        And CCur(lblMontoCom.Caption) <> CCur(IIf(Trim(lblMonTra) = "", "0", lblMonTra)) Then
        MsgBox "El monto de la operación no es igual al monto de la tranferencia. Verifique.", vbInformation, "¡Aviso!"
        EnfocaControl cboTipoPago
        Exit Function
    End If
    
    If CInt(Trim(Right(cboTipoPago.Text, 10))) = gColocTipoPagoCargoCta And Len(txtCuenta.NroCuenta) <> 18 Then
        MsgBox "Favor de ingresar la cuenta de ahorros a debitar", vbInformation, "¡Aviso!"
        cboTipoPago.SetFocus
        Exit Function
    End If
 
    If CInt(Trim(Right(cboTipoPago.Text, 10))) = gColocTipoPagoCargoCta Then
        If Not clsCap.ValidaSaldoCuenta(txtCuenta.NroCuenta, CDbl(lblMontoCom.Caption)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
                                 
            Exit Function
        End If
    End If
   
    ValidaFormaPago = True
End Function
Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
'    COMENTADO POR CTI5 ERS0112020
'    If KeyAscii = 13 Then
'        cmdGrabar.SetFocus
'    End If
    Dim sMsg As String
    If KeyAscii = 13 Then sMsg = ValidaCuentaACargo(txtCuenta.NroCuenta)
    If Len(sMsg) > 0 Then
        MsgBox sMsg, vbInformation, "¡Aviso!"
        txtCuenta.SetFocus
        Exit Sub
    End If
    If Len(txtCuenta.NroCuenta) = 18 Then
        If CInt(Mid(txtCuenta.NroCuenta, 9, 1)) <> CInt(Mid(txtCuenta.NroCuenta, 9, 1)) Then
            MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la cuenta a amortizar.", vbOKOnly + vbInformation, App.Title
        End If
    End If
    ObtieneDatosCuenta txtCuenta.NroCuenta
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
    
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsV = New ADODB.Recordset
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(psCuenta)
    Set loVistoElectronico = New frmVistoElectronico
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
                        Dim nRespuesta As Integer
                        Set rsCli = clsCli.GetPersonaCuenta(psCuenta, gCapRelPersTitular)
                        nRespuesta = oSolicitud.SolicitarVistoAtencionSinTarjeta(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gComiCredDiversasCargoCta))
                    
                        If nRespuesta = 1 Then '1:Tiene Visto de atencion sin tarjeta pendiente de autorizar
                             MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                             Exit Sub
                        End If
                        If nRespuesta = 2 Then '2:Tiene visto aceptado
                            MsgBox "La solicitud de atención sin tarjeta fue Aprobada, proceda con la atención", vbInformation, "Aviso"
                        End If
                        If nRespuesta = 3 Then '3:Tiene visto rechazado
                           If MsgBox("La solicitud de atención sin tarjeta fue RECHAZADA. ¿Desea realizar una nueva solicitud?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                Call oSolicitud.RegistrarVistoDeUsuario(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gComiCredDiversasCargoCta))
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
                        
                                bExitoSol = frmCapMotivoAutorizacion.Inicio(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gComiCredDiversasCargoCta))
                                If bExitoSol > 0 Then
                                    MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                                End If
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        lbVistoVal = loVistoElectronico.Inicio(5, gComiCredDiversasCargoCta)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema"
                            Exit Sub
                        End If
                    End If
                ElseIf lsTieneTarj = "NO" And rsV.RecordCount > 0 Then
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers 065-2017
                        lbVistoVal = loVistoElectronico.Inicio(5, gComiCredDiversasCargoCta)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema" 'ADD BY MARG ERS065-2017
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            End If
        End If
        txtCuenta.Enabled = False
        'AsignaValorITF
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
    End If
End Sub
Private Sub Form_Load()
    Dim oTipCambio As COMDConstSistema.NCOMTipoCambio
    Set oTipCambio = New COMDConstSistema.NCOMTipoCambio
    gnTipCambioC = oTipCambio.EmiteTipoCambio(gdFecSis, TCCompra)
    gnTipCambioV = oTipCambio.EmiteTipoCambio(gdFecSis, TCVenta)
    Call CargaControles
    'CTI7 OPEv2*************************************************
    Me.lblTTCCD.Caption = Format(gnTipCambioC, "#,#0.0000")
    Me.lblTTCVD.Caption = Format(gnTipCambioV, "#,#0.0000")
    '****************************************************************
End Sub
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gColocTipoPago, , , 4)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, cboTipoPago)
    cboTransferMoneda.Clear
    IniciaCombo cboTransferMoneda, gMoneda
    cboTransferMoneda.ListIndex = IndiceListaCombo(cboTransferMoneda, 1)
    Me.fraTranferecia.Enabled = False
    
    Exit Sub
ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub
'END CTI5

Private Sub cboTipoPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Right(cboTipoPago.Text, 2)) <> "" Then
            If Trim(Right(cboTipoPago.Text, 2)) = "1" Then
                cmdGrabar.SetFocus
            Else
                txtCuenta.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmdBCodCom_Click()
    'lblCodCom.Caption = frmCapParametrosCom.BuscarComision
    lblCodCom.Caption = frmCapParametrosCom.BuscarComision(fsParProd) 'JUEZ 20150930
    If lblCodCom.Caption <> "" Then
    Dim clsCapDef As New COMNCaptaGenerales.NCOMCaptaDefinicion
        'Set rs = clsCapDef.GetParametrosComision(lblCodCom.Caption)
        Set rs = clsCapDef.GetParametrosComision(lblCodCom.Caption, , fsParProd) 'JUEZ 20150930
        Set clsCapDef = Nothing
        lblComDesc.Caption = rs!cParDesc
        lblMonedaCom.Caption = IIf(rs!nParMoneda = 1, "S/", "$")
        lblMontoCom.Caption = Format(rs!nParMonto, "#,##0.00")
        fnTipoCobro = rs!nParTipo
        HabilitaControles (True)
        TxtBCodPers.SetFocus
    Else
        cmdLimpiar_Click
    End If
End Sub

Private Sub cmdLimpiar_Click()
    lblCodCom.Caption = ""
    lblComDesc.Caption = ""
    TxtBCodPers.Enabled = True
    TxtBCodPers.Text = ""
    lblPersNombre.Caption = ""
    cboTipoPago.ListIndex = -1
    lblMonedaCom.Caption = ""
    lblMontoCom.Caption = ""
    HabilitaControles (False)
    txtCuenta.CMAC = "109" 'JUEZ 20150930
    txtCuenta.Prod = gCapAhorros 'JUEZ 20150930
    txtCuenta.Visible = False
    sNumTarj = ""
End Sub

Private Sub cmdGrabar_Click()
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII
    Dim lsOpeDescripcion As String
    Dim lsOpeCod As CaptacOperacion
    Dim lnMonto As Double, sCuenta As String, lnMoneda As Moneda, lnITF As Double, nRedondeoITF As Double
    Dim lnSaldo As Double, lnMontoMonCta As Double, lnMonedaCta As Integer, lnTCOpe As Double
    Dim oCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oCapGen As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim clsCont As COMNContabilidad.NCOMContFunciones
    Set oCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oCapGen = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set clsCont = New COMNContabilidad.NCOMContFunciones
    Dim clsTC As COMDConstSistema.NCOMTipoCambio
    Dim lsMov As String
    Dim lsBoleta As String, lsmensaje As String, lsBoletaITF As String
    Dim loMov As COMDMov.DCOMMov
    Set loMov = New COMDMov.DCOMMov
    Dim oBol As COMNCredito.NCOMCredDoc
    Dim oDCredAct As COMDCredito.DCOMCredActBD
    
    Dim loLavDinero As frmMovLavDinero
    Set loLavDinero = New frmMovLavDinero
    Dim sPersLavDinero As String, sReaPersLavDinero As String, sBenPersLavDinero As String
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim nMontoLavDinero As Double, nTC As Double

    Dim fbPersonaReaAhorros As Boolean
    Dim fnCondicion As Integer, nI As Integer
    Dim oPersonaSPR As UPersona_Cli
    Dim oPersonaU As COMDPersona.UCOMPersona
    Dim nTipoConBN As Integer
    Dim sConPersona As String
    Dim pbClienteReforzado As Boolean
    Dim rsAgeParam As Recordset
    Dim objCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim lnMontoX As Double, lnTC As Double
    Dim ObjTc As COMDConstSistema.NCOMTipoCambio
    Dim lsGlosa As String 'JUEZ 20150930
    Dim lnConceptoCod As Long 'JUEZ 20150930
    Dim MatDatosAho(14) As String 'CTI5 ERS0112020
    
    'If fsParProd = "C" Then cboTipoPago.ListIndex = IndiceListaCombo(cboTipoPago, 1) 'JUEZ 20150930 COMENTADO POR CTI5
    
    lsGlosa = "Cobro " & IIf(fnTipoCobro = 1, "Gasto", "Comisión") & " " & IIf(fsParProd = "A", "Ahorros ", "Créditos") & " " & lblCodCom.Caption & " cliente: " & lblPersNombre.Caption 'JUEZ 20150930
       
    If Not ValidaFormaPago Then Exit Sub 'CTI6 ERS0112020
    If ValidaDatos Then
        sCuenta = txtCuenta.NroCuenta
        lnMonto = lblMontoCom.Caption
        lnMoneda = IIf(lblMonedaCom.Caption = "S/", 1, 2)
        lnConceptoCod = CInt("1" & Mid(lblCodCom.Caption, 4, Len(lblCodCom.Caption) - 1)) 'JUEZ 20150930
        
        gnMovNro = 0
        lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

        If Trim(Right(cboTipoPago.Text, 2)) = "2" Then
            If Mid(sCuenta, 6, 3) = gCapAhorros Then
                lsOpeCod = gAhoCargoCobroComDiversasAho
            ElseIf Mid(sCuenta, 6, 3) = gCapCTS Then
                lsOpeCod = gCTSCargoCobroComDiversasAho
            End If
            
            If Len(Trim(sCuenta)) <> 18 Then
                MsgBox "Debe ingresar la cuenta a la que se debitará el cobro de la comisión", vbInformation, "Aviso"
                txtCuenta.SetFocus
                Exit Sub
            End If
            If Mid(sCuenta, 6, 3) = "233" Then
                MsgBox "La comisión no puede ser debitada de un Plazo Fijo", vbInformation, "Aviso"
                Exit Sub
            End If
            If Mid(sCuenta, 6, 3) <> "232" And Mid(sCuenta, 6, 3) <> "234" Then
                MsgBox "La comisión sólo puede ser debitada de una cuenta de Ahorros o CTS", vbInformation, "Aviso"
                Exit Sub
            End If

            lsmensaje = oCapMov.ValidaCuentaOperacion(sCuenta)
            If lsmensaje <> "" Then
                MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Sub
            End If
            
            Dim rsPersona As ADODB.Recordset
            Dim bPersValida As Boolean
            bPersValida = False
            Set rsPersona = oCapGen.GetPersonaCuenta(sCuenta)
            If Not rsPersona.EOF Then
                Do While Not rsPersona.EOF
                    If rsPersona!cperscod = TxtBCodPers.Text Then
                        bPersValida = True
                        Exit Do
                    End If
                    rsPersona.MoveNext
                Loop
            End If
            If Not bPersValida Then
                MsgBox "La persona no tiene ninguna relación con la cuenta ingresada", vbInformation, "Aviso"
                Exit Sub
            End If
    
            lnMonedaCta = CInt(Mid(sCuenta, 9, 1))
            'Set clsTC = New COMDConstSistema.NCOMTipoCambio
            'lnTCOpe = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
            'Set clsTC = Nothing
            GetTipCambio gdFecSis
            
            If lnMoneda = lnMonedaCta Then
                lnMontoMonCta = lnMonto
            ElseIf lnMoneda = 1 And lnMonedaCta = 2 Then
                'lnMontoMonCta = Round(lnMonto / lnTCOpe, 2)
                
                'ANDE 20170120
                'lnMontoMonCta = Round(lnMonto / gnTipCambioV, 2)
                lnMontoMonCta = Round(lnMonto / gnTipCambioC, 2)
                'End ANDE
            ElseIf lnMoneda = 2 And lnMonedaCta = 1 Then
                'lnMontoMonCta = Round(lnMonto * lnTCOpe, 2)
                'ANDE 20170120
                'lnMontoMonCta = Round(lnMonto * gnTipCambioC, 2)
                lnMontoMonCta = Round(lnMonto * gnTipCambioV, 2)
                'End ANDE
            End If


            If oCapMov.ValidaSaldoCuenta(sCuenta, lnMontoMonCta) Then
                If MsgBox("Desea Grabar la Operación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
                
                lnITF = fgITFCalculaImpuesto(lnMontoMonCta)
                nRedondeoITF = fgDiferenciaRedondeoITF(lnITF)
                If nRedondeoITF > 0 Then
                    lnITF = Format(lnITF - nRedondeoITF, "#,##0.00")
                End If
                
                'Validación Lavado de Dinero REU
                Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
                sPersLavDinero = ""
                nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
                Set clsLav = Nothing

                If lnMoneda = gMonedaNacional Then
                    Set clsTC = New COMDConstSistema.NCOMTipoCambio
                    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                    Set clsTC = Nothing
                Else
                    nTC = 1
                End If
                If lnMonto >= Round(nMontoLavDinero * nTC, 2) Then
                    sPersLavDinero = loLavDinero.Inicia(, , , , False, True, lnMonto, sCuenta, Mid(Me.Caption, 15), True, , , , , , lnMoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                    If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                End If

                'Formulario solicitud de información
                nI = 0
                fbPersonaReaAhorros = False
                If (loLavDinero.OrdPersLavDinero = "Exit") Then

                    Set oPersonaU = New COMDPersona.UCOMPersona
                    Set oPersonaSPR = New UPersona_Cli

                    fbPersonaReaAhorros = False
                    pbClienteReforzado = False
                    fnCondicion = 0

                    oPersonaSPR.RecuperaPersona Trim(TxtBCodPers.Text)

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
                        MsgBox "El Cliente: " & Trim(lblPersNombre.Caption) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
                        frmPersRealizaOpeGeneral.Inicia Left(Me.Caption, 38) & " (Persona " & sConPersona & ")", lsOpeCod
                        fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar

                        If Not fbPersonaReaAhorros Then
                            MsgBox "Se va a proceder a Anular la Operacion ", vbInformation, "Aviso"
                            Exit Sub
                        End If
                    Else
                        fnCondicion = 0
                        lnMontoX = lnMonto
                        pbClienteReforzado = False

                        Set ObjTc = New COMDConstSistema.NCOMTipoCambio
                        lnTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
                        Set ObjTc = Nothing


                        Set objCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                        Set rsAgeParam = objCap.getCapAbonoAgeParam(gsCodAge)
                        Set objCap = Nothing

                        If Mid(Trim(txtCuenta.NroCuenta), 9, 1) = 1 Then
                            lnMontoX = Round(lnMontoX / lnTC, 2)
                        End If

                        If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                            If lnMontoX >= rsAgeParam!nMontoMin And lnMontoX <= rsAgeParam!nMontoMax Then
                                frmPersRealizaOpeGeneral.Inicia Left(Me.Caption, 38), lsOpeCod
                                fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar
                                If Not fbPersonaReaAhorros Then
                                    MsgBox "Se va a proceder a Anular la Operacion", vbInformation, "Aviso"
                                    Exit Sub
                                End If
                            End If
                        End If

                    End If
                End If
                
                Select Case Mid(sCuenta, 6, 3)
                    Case gCapAhorros
                        lnSaldo = oCapMov.CapCargoCuentaAho(sCuenta, lnMontoMonCta, lsOpeCod, lsMov, "Cuenta = " & sCuenta, , , , , , , , , gsNomAge, sLpt, sReaPersLavDinero, , , , gsCodCMAC, , gsCodAge, , gbITFAplica, CCur(lnITF), gbITFAsumidoAho, gITFCobroCargo, lsOpeCod, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, , , gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, sNumTarj, , , , , , , , , lnConceptoCod)
                    Case gCapCTS
                        lnSaldo = oCapMov.CapCargoCuentaCTS(sCuenta, lnMontoMonCta, lsOpeCod, lsMov, "", , , , , , , gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, , , , , , , , , lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, , , , , , , , , , , , , , lnConceptoCod)
                End Select

                If gnMovNro > 0 Then
                    Call oCapMov.InsertaCapComisiones(gnMovNro, lblCodCom.Caption, TxtBCodPers.Text, Trim(Right(Me.cboTipoPago.Text, 2)), IIf(Len(sCuenta) = 18, sCuenta, ""))
                End If
                If gnMovNro > 0 Then 'INSERTA REU
                     Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4)
                End If
                If fbPersonaReaAhorros And gnMovNro > 0 Then 'INSERTA SOLICITUD INFORMACION OC
                    frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(sCuenta), fnCondicion
                    fbPersonaReaAhorros = False
                End If

                If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                'If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                'If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
                
                Set oBol = New COMNCredito.NCOMCredDoc
                     lsBoleta = oBol.ImprimeBoletaComision("COMISIONES VARIAS", Left(lblCodCom.Caption & " - " & lblComDesc.Caption, 36), "", Str(CDbl(lnMontoMonCta)), lblPersNombre.Caption, lblDOI.Caption, IIf(Len(sCuenta) = 18, sCuenta, "________" & lnMoneda), False, "", "", , gdFecSis, gsNomAge, gsCodUser, sLpt, , gbImpTMU, True, Trim(Left(cboTipoPago.Text, 50)), , txtCuenta.NroCuenta)
                Set oBol = Nothing
                ImprimeBoleta lsBoleta

                If gbITFAplica = True And CCur(lnITF) > 0 Then
                    Call loMov.InsertaMovRedondeoITF(lsMov, 1, CCur(lnITF) + nRedondeoITF, CCur(lnITF))
                End If

                Set oCapMov = Nothing
                Set loLavDinero = Nothing
                cmdLimpiar_Click
            Else
                MsgBox "Cuenta NO posee saldo suficiente para realizar la operación", vbInformation, "Aviso"
            End If
            lsOpeDescripcion = "Cobro comisión "
        Else
            If CInt(Trim(Right(cboTipoPago.Text, 10))) = gColocTipoPagoVoucher Then
                lsOpeCod = gComiDiversasAhoCredComVoucher
                lsOpeDescripcion = "Cobro comisión "
            Else
                If fnTipoCobro = 1 Then
                    lsOpeCod = gComiDiversasAhoCredGasto
                    lsOpeDescripcion = "Cobro comisión "
                Else
                    lsOpeCod = gComiDiversasAhoCredCom
                    lsOpeDescripcion = "Cobro comisión créditos "
                End If
            End If
            If MsgBox("Desea Grabar la Operación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
                          
            lnITF = fgITFCalculaImpuesto(lnMonto)
            nRedondeoITF = fgDiferenciaRedondeoITF(lnITF)
            If nRedondeoITF > 0 Then
                lnITF = Format(lnITF - nRedondeoITF, "#,##0.00")
            End If
            
            'Validación Lavado de Dinero REU
            Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing

            If lnMoneda = gMonedaNacional Then
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If lnMonto >= Round(nMontoLavDinero * nTC, 2) Then
                sPersLavDinero = loLavDinero.Inicia(, , , , False, True, lnMonto, sCuenta, Mid(Me.Caption, 15), True, , , , , , lnMoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
            End If

            'Formulario solicitud de información
            nI = 0
            fbPersonaReaAhorros = False
            If (loLavDinero.OrdPersLavDinero = "Exit") Then

                Set oPersonaU = New COMDPersona.UCOMPersona
                Set oPersonaSPR = New UPersona_Cli

                fbPersonaReaAhorros = False
                pbClienteReforzado = False
                fnCondicion = 0

                oPersonaSPR.RecuperaPersona Trim(TxtBCodPers.Text)

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
                    MsgBox "El Cliente: " & Trim(lblPersNombre.Caption) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
                    frmPersRealizaOpeGeneral.Inicia Left(Me.Caption, 38) & " (Persona " & sConPersona & ")", lsOpeCod
                    fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar

                    If Not fbPersonaReaAhorros Then
                        MsgBox "Se va a proceder a Anular la Operacion ", vbInformation, "Aviso"
                        Exit Sub
                    End If
                Else
                    fnCondicion = 0
                    lnMontoX = lnMonto
                    pbClienteReforzado = False

                    Set ObjTc = New COMDConstSistema.NCOMTipoCambio
                    lnTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
                    Set ObjTc = Nothing


                    Set objCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                    Set rsAgeParam = objCap.getCapAbonoAgeParam(gsCodAge)
                    Set objCap = Nothing

                    If Mid(Trim(txtCuenta.NroCuenta), 9, 1) = 1 Then
                        lnMontoX = Round(lnMontoX / lnTC, 2)
                    End If

                    If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                        If lnMontoX >= rsAgeParam!nMontoMin And lnMontoX <= rsAgeParam!nMontoMax Then
                            frmPersRealizaOpeGeneral.Inicia Left(Me.Caption, 38), lsOpeCod
                            fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar
                            If Not fbPersonaReaAhorros Then
                                MsgBox "Se va a proceder a Anular la Operacion", vbInformation, "Aviso"
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
            Dim lnTransferenciaMoneda As Moneda
            Dim lnTransferenciaMonto As Currency
            
            If Me.cboTransferMoneda.Text <> "" Then
                lnTransferenciaMoneda = CInt(Right(Me.cboTransferMoneda.Text, 3))
            Else
                lnTransferenciaMoneda = gMonedaNacional
            End If
            
            If Trim(lblMonTra) <> "" Then
                lnTransferenciaMonto = CCur(lblMonTra)
            Else
                lnTransferenciaMonto = 1
            End If
            
            Dim lsFechaHoraGrab As String               'CTI5 ERS0112020
            lsFechaHoraGrab = fgFechaHoraGrab(lsMov)    'CTI5 ERS0112020
            'gnMovNro = oCapMov.OtrasOperaciones(lsMov, lsOpeCod, CDbl(lnMonto), "", "Cobro comisión Ahorros " & lblCodCom.Caption & " cliente: " & lblPersNombre.Caption, lnMoneda, TxtBCodPers.Text, , , , , , , gnMovNro)
            'gnMovNro = oCapMov.OtrasOperaciones(lsMov, lsOpeCod, CDbl(lnMonto), "", "Cobro comisión Ahorros " & lblCodCom.Caption & " cliente: " & lblPersNombre.Caption, lnMoneda, TxtBCodPers.Text, , , , , , , gnMovNro, , , lnConceptoCod) 'JUEZ 20151019
            gnMovNro = oCapMov.OtrasOperaciones(lsMov, lsOpeCod, CDbl(lnMonto), "", lsOpeDescripcion & lblCodCom.Caption & " cliente: " & lblPersNombre.Caption, lnMoneda, TxtBCodPers.Text, , , , , , , gnMovNro, , , lnConceptoCod, sCuenta, MatDatosAho, Trim(Right(Me.cboTipoPago.Text, 2)), lsFechaHoraGrab, , , , , , lnMovNroTransfer, lnTransferenciaMoneda, fnMovNroRVD, CCur(lnTransferenciaMonto)) 'CTI5 ERS0112020
            If gnMovNro <> 0 Then
                'JUEZ 20150930 *****************************************
                'Call oCapMov.InsertaCapComisiones(gnMovNro, lblCodCom.Caption, TxtBCodPers.Text, Trim(Right(Me.cboTipoPago.Text, 2)), "")
                    If fsParProd = "A" Then
                        Call oCapMov.InsertaCapComisiones(gnMovNro, lblCodCom.Caption, TxtBCodPers.Text, Trim(Right(Me.cboTipoPago.Text, 2)), "")
                    ElseIf fsParProd = "C" Then
                        Set oDCredAct = New COMDCredito.DCOMCredActBD
                        Call oDCredAct.dInsertComision(gnMovNro, TxtBCodPers.Text, CDbl(lnMonto), 0, lblCodCom.Caption) 'JUEZ 20150930
                        Set oDCredAct = Nothing
                    End If
                    'END JUEZ **********************************************
                If lnITF > 0 Then
                    Dim lsMovITF As String
                    lsMovITF = Mid(lsMov, 1, 20) & "1" & gsCodUser
                    Call oCapMov.OtrasOperaciones(lsMovITF, gITFCobroEfectivo, CDbl(lnITF), sCuenta, "", lnMoneda, TxtBCodPers.Text)
                End If

                'INSERTA REU
                Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224

                If fbPersonaReaAhorros Then 'INSERTA SOLICITUD INFORMACION OC
                    frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(sCuenta), fnCondicion
                    fbPersonaReaAhorros = False
                End If
                
                If gbITFAplica = True And CCur(lnITF) > 0 Then
                    Call loMov.InsertaMovRedondeoITF(lsMov, 1, CCur(lnITF) + nRedondeoITF, CCur(lnITF))
                End If

                Set oBol = New COMNCredito.NCOMCredDoc
                    lsBoleta = oBol.ImprimeBoletaComision(IIf(Trim(Right(cboTipoPago.Text, 50)) = 4, "Comis.Diversas Crédito-Cargo a Cta", "COMISIONES DIVERSAS"), Left(lblCodCom.Caption & " - " & lblComDesc.Caption, 36), "", Str(CDbl(lnMonto)), lblPersNombre.Caption, lblDOI.Caption, IIf(Len(sCuenta) = 18, sCuenta, "________" & lnMoneda), False, "", "", , gdFecSis, gsNomAge, gsCodUser, sLpt, , gbImpTMU, True, Trim(Left(cboTipoPago.Text, 50)), , txtCuenta.NroCuenta)
                Set oBol = Nothing

                Do
                   If Trim(lsBoleta) <> "" Then
                        lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea
                        nFicSal = FreeFile
                        Open sLpt For Output As nFicSal
                            Print #nFicSal, oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                            Print #nFicSal, ""
                        Close #nFicSal
                  End If

                Loop While MsgBox("Desea Re Imprimir el voucher ?", vbQuestion + vbYesNo, "Aviso") = vbYes
                Set oBol = Nothing
                cmdLimpiar_Click
                'INICIO JHCU ENCUESTA 16-10-2019
                Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
                'FIN
                Call IniciarControlesFormaPago 'CTI7 OPEv2
            Else
                MsgBox "Hubo un error al grabar la operación", vbInformation, "Aviso"
            End If
        End If
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload frmCapParametrosCom
    Unload Me
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF10 And txtCuenta.Enabled Then
'        Dim lsCuenta As String
'        lsCuenta = frmATMCargaCuentas.RecuperaCuenta(, sNumTarj, txtCuenta.Prod)
'        If val(Mid(lsCuenta, 6, 3)) <> txtCuenta.Prod And lsCuenta <> "" Then
'            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
'            Exit Sub
'        End If
'        If lsCuenta <> "" Then
'            txtCuenta.NroCuenta = lsCuenta
'            txtCuenta.SetFocusCuenta
'        End If
'    End If
'End Sub

Private Sub TxtBCodPers_EmiteDatos()
    If Trim(TxtBCodPers.Text) <> "" Then
    Dim oCom As New COMDCredito.DCOMCredito
        Set rs = oCom.RecuperaDatosComision(TxtBCodPers.Text, 2)
        Set oCom = Nothing
        lblDOI.Caption = rs!cPersIDnro
        lblPersNombre.Caption = rs!cPersNombre
        TxtBCodPers.Enabled = False
        'cboTipoPago.SetFocus
        If cboTipoPago.Visible Then cboTipoPago.SetFocus 'JUEZ 20150930
    Else
        lblDOI.Caption = ""
        lblPersNombre.Caption = ""
    End If
End Sub

Private Function ValidaDatos() As Boolean
    ValidaDatos = False
    If Trim(lblPersNombre.Caption) = "" And Trim(TxtBCodPers.Text) = "" Then
        MsgBox "Debe ingresar a la persona que está relizando la operación", vbInformation, "Aviso"
        TxtBCodPers.SetFocus
        Exit Function
    End If
    If Trim(cboTipoPago.Text) = "" Then
        MsgBox "Debe seleccionar el tipo de pago", vbInformation, "Aviso"
        'cboTipoPago.SetFocus
        If cboTipoPago.Visible Then cboTipoPago.SetFocus 'JUEZ 20150930
        Exit Function
    End If
    ValidaDatos = True
End Function

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF10 And txtCuenta.Enabled Then
        Dim lsCuenta As String
        lsCuenta = frmATMCargaCuentas.RecuperaCuenta(, sNumTarj, txtCuenta.Prod)
        If Val(Mid(lsCuenta, 6, 3)) <> txtCuenta.Prod And lsCuenta <> "" Then
            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        If lsCuenta <> "" Then
            txtCuenta.NroCuenta = lsCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub ImprimeBoleta(ByVal sBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, sBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    Print #nFicSal, ""
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub

'CTI7 OPEv2***********************************************************************
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
  
    lnTipMot = 17
   
    fnMovNroRVD = 0
    Set oForm = New frmCapRegVouDepBus

    SetDatosTransferencia "", "", "", 0, -1, "" 'Limpiamos datos y variables globales
    oForm.iniciarFormulario Trim(Right(cboTransferMoneda, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, lsDetalle
    If fnMovNroRVD = 0 Then
        Exit Sub
    End If
    SetDatosTransferencia lsGlosa, lsInstit, lsDoc, lnTransferSaldo, lnMovNroTransfer, lsDetalle

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

    If pnMovNroTransfer <> -1 Then
        EnfocaControl txtTransferGlosa
    End If
    txtTransferGlosa.Locked = True

    lblMonTra = Format(pnTransferSaldo, "#,##0.00")
'

End Sub
'*************************************************************

