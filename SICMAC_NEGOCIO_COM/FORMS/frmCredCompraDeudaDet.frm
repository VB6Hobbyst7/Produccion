VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredCompraDeudaDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de crédito"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frmCredCompraDeudaDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3240
      TabIndex        =   3
      Top             =   3900
      Width           =   1000
   End
   Begin TabDlg.SSTab TabCreditoAnt 
      Height          =   1440
      Left            =   0
      TabIndex        =   4
      Top             =   4320
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2540
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Datos de crédito anterior"
      TabPicture(0)   =   "frmCredCompraDeudaDet.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Shape5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtFecCanledadoAnt"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtMontoCuotaAnt"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtMontoDesembAnt"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.TextBox txtMontoDesembAnt 
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
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   480
         Width           =   1260
      End
      Begin VB.TextBox txtMontoCuotaAnt 
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
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   960
         Width           =   1260
      End
      Begin MSMask.MaskEdBox txtFecCanledadoAnt 
         Height          =   285
         Left            =   4920
         TabIndex        =   0
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H8000000C&
         Height          =   1005
         Left            =   120
         Top             =   360
         Width           =   6285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto cuota que venia pagando :"
         Height          =   390
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Monto desembolso :"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Cancelación :"
         Height          =   195
         Left            =   3240
         TabIndex        =   5
         Top             =   510
         Width           =   1470
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   2160
      TabIndex        =   2
      Top             =   3900
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3855
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Datos de crédito a comprar"
      TabPicture(0)   =   "frmCredCompraDeudaDet.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label14"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label16"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label19"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label20"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label15"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label13"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label11"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtNroCuotaPaga"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtFecDesemb"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtNroCuotaPacta"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtMontoCuota"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtSaldoComprar"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtMontoDesemb"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmbMoneda"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtNroCredito"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "chkRecompra"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmbDestino"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmbIFI"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "chkCreditoAnt"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      Begin VB.CheckBox chkCreditoAnt 
         Caption         =   "Crédito Anterior"
         Height          =   375
         Left            =   2520
         TabIndex        =   25
         Top             =   3240
         Width           =   1815
      End
      Begin VB.ComboBox cmbIFI 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   720
         Width           =   6075
      End
      Begin VB.ComboBox cmbDestino 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1290
         Width           =   2745
      End
      Begin VB.CheckBox chkRecompra 
         Caption         =   "Es una recompra de deuda (antes también Caja Maynas)"
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   1155
         Width           =   2415
      End
      Begin VB.TextBox txtNroCredito 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   1680
         Width           =   1665
      End
      Begin VB.ComboBox cmbMoneda 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2040
         Width           =   1665
      End
      Begin VB.TextBox txtMontoDesemb 
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
         Left            =   5040
         MaxLength       =   15
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   2040
         Width           =   1260
      End
      Begin VB.TextBox txtSaldoComprar 
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   2805
         Width           =   1260
      End
      Begin VB.TextBox txtMontoCuota 
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
         Left            =   5040
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   2805
         Width           =   1260
      End
      Begin SICMACT.uSpinner txtNroCuotaPacta 
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   2430
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         Max             =   1000
         Min             =   1
         MaxLength       =   4
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
      Begin MSMask.MaskEdBox txtFecDesemb 
         Height          =   285
         Left            =   5040
         TabIndex        =   17
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin SICMACT.uSpinner txtNroCuotaPaga 
         Height          =   285
         Left            =   5520
         TabIndex        =   18
         Top             =   2400
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         Max             =   1000
         MaxLength       =   4
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Saldo a comprar :"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   2895
         Width           =   1245
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nro. cuotas pactadas :"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   2475
         Width           =   1620
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   2085
         Width           =   675
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Crédito :"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Institución Financiera :"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   510
         Width           =   1590
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Destino :"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Fecha desembolso :"
         Height          =   195
         Left            =   3360
         TabIndex        =   22
         Top             =   1710
         Width           =   1425
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Monto desembolso :"
         Height          =   195
         Left            =   3360
         TabIndex        =   21
         Top             =   2055
         Width           =   1425
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nro. cuotas pagadas :"
         Height          =   195
         Left            =   3360
         TabIndex        =   20
         Top             =   2445
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Monto cuota que viene pagando :"
         Height          =   390
         Left            =   3360
         TabIndex        =   19
         Top             =   2775
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000C&
         Height          =   3315
         Left            =   120
         Top             =   360
         Width           =   6285
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000E&
         Height          =   3300
         Left            =   120
         Top             =   435
         Width           =   6285
      End
   End
End
Attribute VB_Name = "frmCredCompraDeudaDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCredCompraDeudaDet
'** Descripción : Para registro de las deudas que se comprará a otras IFIs creado segun TI-ERS002-2016
'** Creación : EJVG, 20160129 11:00:00 AM
'*****************************************************************************************************
Option Explicit
Private Enum eCompraDeuda
    CompraDeudaRegistrar = 1
    CompraDeudaModificar = 2
End Enum

Dim fnInicio As eCompraDeuda
Dim fvCompraDeuda As TCompraDeuda

Dim fnIndexValida As Integer
Dim fvListaValida() As TCompraDeuda

Dim fbAceptar As Boolean
Dim fnTpoProducto As Integer '**ARLO20180317 ERS 070 - 2017 ANEXO 02
Dim fnFormTamanioIni As Double, fnFormTamanioActiva As Double 'ARLO20180525
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

Private Sub Form_Load()
    fnFormTamanioIni = 4650 'ARLO20180525
    fnFormTamanioActiva = 6300 'ARLO20180525
    
    fbAceptar = False
    CargaControles
    Limpiar
    
    If fnInicio = CompraDeudaModificar Then
        cmbIFI.ListIndex = IndiceListaCombo(cmbIFI, fvCompraDeuda.sIFICod)
        cmbDestino.ListIndex = IndiceListaCombo(cmbDestino, fvCompraDeuda.nDestino)
        chkRecompra.value = IIf(fvCompraDeuda.bRecompra, 1, 0)
        txtNroCredito.Text = fvCompraDeuda.sNroCredito
        txtFecDesemb.Text = Format(fvCompraDeuda.dDesembolso, "dd/mm/yyyy")
        cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fvCompraDeuda.nMoneda)
        txtMontoDesemb.Text = Format(fvCompraDeuda.nMontoDesembolso, "#,##0.00")
        txtNroCuotaPacta.valor = fvCompraDeuda.nNroCuotasPactadas
        txtNroCuotaPaga.valor = fvCompraDeuda.nNroCuotasPagadas
        txtSaldoComprar.Text = Format(fvCompraDeuda.nSaldoComprar, "#,##0.00")
        txtMontoCuota.Text = Format(fvCompraDeuda.nMontoCuota, "#,##0.00")
        chkCreditoAnt.value = IIf(fvCompraDeuda.bCreditoAnt, 1, 0) '**ARLO20180528
        txtMontoDesembAnt = Format(fvCompraDeuda.nMontoDesembolsoAnt, "#,##0.00")    '**ARLO20180528
        txtMontoCuotaAnt.Text = Format(fvCompraDeuda.nMontoCuotaAnt, "#,##0.00")     '**ARLO20180528
        txtFecCanledadoAnt.Text = Format(fvCompraDeuda.dCancelacion, "dd/mm/yyyy")     '**ARLO20180528
    End If
End Sub
Public Function Registrar(ByRef pvCompraDeuda As TCompraDeuda, ByRef pvListaValida() As TCompraDeuda, ByVal pnTpoProducto As Integer) As Boolean '**ARLO20180317 ADD pnTpoProducto
    fnInicio = CompraDeudaRegistrar
    
    fvCompraDeuda = pvCompraDeuda
    fvListaValida = pvListaValida
    fnTpoProducto = pnTpoProducto '**ARLO20180317 ERS 070 - 2017 ANEXO 02
    Show 1
    Registrar = fbAceptar
    pvCompraDeuda = fvCompraDeuda
End Function
Public Function Modificar(ByRef pvCompraDeuda As TCompraDeuda, ByVal pnIndexValida As Integer, ByRef pvListaValida() As TCompraDeuda, ByVal pnTpoProducto As Integer) As Boolean '**ARLO20180317 ADD pnTpoProducto
    fnInicio = CompraDeudaModificar
    
    fvCompraDeuda = pvCompraDeuda
    fnIndexValida = pnIndexValida
    fvListaValida = pvListaValida
    fnTpoProducto = pnTpoProducto '**ARLO20180317 ERS 070 - 2017 ANEXO 02
    Show 1
    Modificar = fbAceptar
    pvCompraDeuda = fvCompraDeuda
End Function

Private Sub CargaControles()
    Dim oIFI As New COMDPersona.DCOMInstFinac
    Dim oConstante As New COMDConstantes.DCOMConstantes
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrCargaControles
    
    Height = fnFormTamanioIni 'ARLO20180525
    
    Screen.MousePointer = 11
    cmbIFI.Clear
    Set rs = oIFI.CargaIFIxCompraDeuda()
    Do While Not rs.EOF
        cmbIFI.AddItem rs!cPersNombre & Space(200) & rs!cPersCod
        rs.MoveNext
    Loop
    
    cmbDestino.Clear
    'Set rs = oConstante.RecuperaConstantes(gCompraDeudaDestino)'Comento JOEP20190206 CP
    Set rs = oConstante.RecuperaConstantes(gCompraDeudaDestino, , , , fnTpoProducto) 'JOEP20190206 CP
    Call Llenar_Combo_con_Recordset(rs, cmbDestino)
    
    cmbMoneda.Clear
    Set rs = oConstante.RecuperaConstantes(1011)
    Call Llenar_Combo_con_Recordset(rs, cmbMoneda)
    
    RSClose rs
    Set oIFI = Nothing
    Set oConstante = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrCargaControles:
    MsgBox Err.Description, vbCritical, "Aviso"
    Screen.MousePointer = 0
End Sub
Private Sub cmbMoneda_Click()
    Dim lnMoneda As Moneda
    Dim lsColor As String
    
    lnMoneda = val(Right(cmbMoneda.Text, 3))
    
    If lnMoneda = gMonedaNacional Then
        lsColor = &H80000005
    Else
        lsColor = &HC0FFC0
    End If
    
    txtMontoDesemb.BackColor = lsColor
    txtSaldoComprar.BackColor = lsColor
    txtMontoCuota.BackColor = lsColor
End Sub
Private Sub cmdAceptar_Click()
    On Error GoTo ErrAceptar
    Dim bCreditoAnt As Boolean
    
    If Not ValidaDatos Then Exit Sub
    
    fvCompraDeuda.sIFICod = Right(cmbIFI.Text, 13)
    fvCompraDeuda.sIFINombre = Trim(Mid(cmbIFI.Text, 1, Len(cmbIFI) - 13))
    fvCompraDeuda.nDestino = CInt(Trim(Right(cmbDestino.Text, 3)))
    fvCompraDeuda.bRecompra = IIf(chkRecompra.value = 1, True, False)
    fvCompraDeuda.sNroCredito = txtNroCredito.Text
    fvCompraDeuda.dDesembolso = CDate(txtFecDesemb.Text)
    fvCompraDeuda.nMoneda = Right(cmbMoneda.Text, 1)
    fvCompraDeuda.nMontoDesembolso = CDbl(txtMontoDesemb.Text)
    fvCompraDeuda.nNroCuotasPactadas = CInt(txtNroCuotaPacta.valor)
    fvCompraDeuda.nNroCuotasPagadas = CInt(txtNroCuotaPaga.valor)
    fvCompraDeuda.nSaldoComprar = CDbl(txtSaldoComprar.Text)
    fvCompraDeuda.nMontoCuota = CDbl(txtMontoCuota.Text)
    '**ARLO20180528 INICIO
    bCreditoAnt = IIf(chkCreditoAnt.value = 1, True, False)
    If bCreditoAnt Then
    fvCompraDeuda.bCreditoAnt = IIf(chkCreditoAnt.value = 1, True, False)
    fvCompraDeuda.nMontoDesembolsoAnt = CDbl(txtMontoDesembAnt.Text)
    fvCompraDeuda.nMontoCuotaAnt = CDbl(txtMontoCuotaAnt.Text)
    fvCompraDeuda.dCancelacion = CDate(Me.txtFecCanledadoAnt.Text)
    Else
    fvCompraDeuda.bCreditoAnt = IIf(chkCreditoAnt.value = 1, True, False)
    End If
    '**ARLO20180528 FIN
    
    fbAceptar = True
    Unload Me
    Exit Sub
ErrAceptar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub Limpiar()
    cmbIFI.ListIndex = -1
    cmbDestino.ListIndex = -1
    chkRecompra.value = 0
    txtNroCredito.Text = ""
    txtFecDesemb.Text = "__/__/____"
    cmbMoneda.ListIndex = -1
    txtMontoDesemb.Text = "0.00"
    txtNroCuotaPacta.valor = 120
    txtNroCuotaPaga.valor = 1
    txtSaldoComprar.Text = "0.00"
    txtMontoCuota.Text = "0.00"
    txtFecCanledadoAnt.Text = "__/__/____" 'ARLO20180528
    txtMontoDesembAnt.Text = "0.00" 'ARLO20180528
    txtMontoCuotaAnt.Text = "0.00" 'ARLO20180528
End Sub
Private Sub txtMontoCuota_LostFocus()
    txtMontoCuota.Text = Format(txtMontoCuota, "#,##0.00")
End Sub
Private Sub TxtMontoDesemb_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMontoDesemb, KeyAscii, 15)
    
    If KeyAscii = 13 Then
        EnfocaControl txtNroCuotaPacta
    End If
    
End Sub
Private Sub txtMontoDesemb_LostFocus()
    txtMontoDesemb.Text = Format(txtMontoDesemb, "#,##0.00")
End Sub
Private Sub txtNroCredito_LostFocus()
    txtNroCredito.Text = Trim(txtNroCredito.Text)
End Sub
Private Sub txtNroCuotaPacta_LostFocus()
    If val(txtNroCuotaPacta.valor) <= val(txtNroCuotaPaga.valor) Then
        MsgBox "El Nro. de Cuotas pactadas debe ser mayor al Nro. de Cuotas Pagadas", vbInformation, "Aviso"
        'EnfocaControl txtNroCuotaPacta
        Exit Sub
    End If
End Sub
Private Sub txtNroCuotaPaga_LostFocus()
    If val(txtNroCuotaPaga.valor) >= val(txtNroCuotaPacta.valor) And val(txtNroCuotaPacta.valor) > 1 Then
        MsgBox "El Nro. de Cuotas Pagadas debe ser menor al Nro. de Cuotas Pactadas", vbInformation, "Aviso"
        'EnfocaControl txtNroCuotaPaga
        Exit Sub
    End If
End Sub
Private Sub txtSaldoComprar_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtSaldoComprar, KeyAscii, 15)
    
    If KeyAscii = 13 Then
        EnfocaControl txtMontoCuota
    End If
    
    fEnfoque txtMontoCuota 'JOEP ERS004-2016 20-08-2016
End Sub
Private Sub txtMontoCuota_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMontoCuota, KeyAscii, 15)
    
    If KeyAscii = 13 Then
        EnfocaControl cmdAceptar
    End If
End Sub
Private Sub cmbIFI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmbDestino
    End If
End Sub
Private Sub CmbDestino_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        '**ARLO20180419
        If cmbDestino.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el destino del prestamo", vbInformation, "Aviso"
        EnfocaControl cmbDestino
        Exit Sub
        End If
        '**ARLO20180419
        
        '**ARLO20171127 ERS070 - 2017
        If (CInt(Trim(Right(cmbDestino.Text, 3))) = 3) Then
            Me.Label14.Caption = "Monto de linea :"
            Me.Label11.Caption = "Saldo a comprar " & Chr(13) & " de Linea :"
        Else
            Me.Label14.Caption = "Monto desembolso :"
            Me.Label11.Caption = "Saldo a comprar :"  'ARLO20180601
        End If
        '**ARLO20171127 ERS070 - 2017
        EnfocaControl chkRecompra
    End If
End Sub
Private Sub chkRecompra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtNroCredito
    End If
End Sub
Private Sub txtNroCredito_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtFecDesemb
    End If
End Sub
Private Sub TxtFecDesemb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmbMoneda
    End If
End Sub
Private Sub CmbMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtMontoDesemb
    End If
    fEnfoque txtMontoDesemb 'JOEP ERS004-2016 20-08-2016
End Sub
Private Sub txtNroCuotaPacta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtNroCuotaPaga
    End If
End Sub
Private Sub txtNroCuotaPaga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtSaldoComprar
            'ARLO20180530
            If (txtNroCuotaPaga.valor) <= 5 Then
                chkCreditoAnt.value = 1
                chkCreditoAnt.Enabled = True
            Else
                chkCreditoAnt.value = 0
                chkCreditoAnt.Enabled = False
            End If
            'ARLO20180530
    End If
    fEnfoque txtSaldoComprar 'JOEP ERS004-2016 20-08-2016
End Sub
Private Function ValidaDatos() As Boolean
    Dim lsFecha, lsFechaCan As String '**ARLO20180528
    Dim lnIndexValida As Integer
    Dim i As Integer

    If cmbIFI.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la Institución Financiera", vbInformation, "Aviso"
        EnfocaControl cmbIFI
        Exit Function
    End If
    If cmbDestino.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el destino del prestamo", vbInformation, "Aviso"
        EnfocaControl cmbDestino
        Exit Function
    End If
    If Len(Trim(txtNroCredito.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Nro. de Crédito que tenía en la Institución Financiera", vbInformation, "Aviso"
        EnfocaControl txtNroCredito
        Exit Function
    End If
    lsFecha = ValidaFecha(txtFecDesemb)
    If Len(lsFecha) > 0 Then
        MsgBox lsFecha, vbInformation, "Aviso"
        EnfocaControl txtFecDesemb
        Exit Function
    End If
    If CDate(txtFecDesemb) > gdFecSis Then
        MsgBox "La fecha de desembolso no puede ser mayor o igual a la fecha del sistema", vbInformation, "Aviso"
        EnfocaControl txtFecDesemb
        Exit Function
    End If
    If cmbMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la moneda del crédito", vbInformation, "Aviso"
        EnfocaControl cmbMoneda
        Exit Function
    End If
    If Not IsNumeric(txtMontoDesemb.Text) Then
        MsgBox "Ud. debe ingresar el monto de desembolso", vbInformation, "Aviso"
        EnfocaControl txtMontoDesemb
        Exit Function
    Else
        If CDbl(txtMontoDesemb.Text) <= 0 Then
            MsgBox "Ud. debe ingresar el monto de desembolso", vbInformation, "Aviso"
            EnfocaControl txtMontoDesemb
            Exit Function
        End If
    End If
    If Not IsNumeric(txtNroCuotaPacta.valor) Then
        MsgBox "Ud. debe ingresar el Nro. de Cuotas pactadas", vbInformation, "Aviso"
        EnfocaControl txtNroCuotaPacta
        Exit Function
    End If
    If Not IsNumeric(txtNroCuotaPaga.valor) Then
        MsgBox "Ud. debe ingresar el Nro. de Cuotas pagadas", vbInformation, "Aviso"
        EnfocaControl txtNroCuotaPaga
        Exit Function
    End If
    If val(txtNroCuotaPaga.valor) >= val(txtNroCuotaPacta.valor) Then
        MsgBox "El Nro. de Cuotas Pagadas debe ser menor al Nro. de Cuotas Pactadas", vbInformation, "Aviso"
        EnfocaControl txtNroCuotaPaga
        Exit Function
    End If
    If Not IsNumeric(txtSaldoComprar.Text) Then
        MsgBox "Ud. debe ingresar el Saldo que se va a comprar", vbInformation, "Aviso"
        EnfocaControl txtSaldoComprar
        Exit Function
    Else
        If CDbl(txtSaldoComprar.Text) <= 0 Then
            MsgBox "Ud. debe ingresar el Saldo que se va a comprar", vbInformation, "Aviso"
            EnfocaControl txtSaldoComprar
            Exit Function
        End If
    End If
    If Not IsNumeric(txtMontoCuota.Text) Then
        MsgBox "Ud. debe ingresar el Monto de Cuota que se viene pagando del crédito a comprar", vbInformation, "Aviso"
        EnfocaControl txtMontoCuota
        Exit Function
    '**COMENTADO BY ARLO20171127 ERS070 - 2017
'    Else
'        If CDbl(txtMontoCuota.Text) <= 0 Then
'            MsgBox "Ud. debe ingresar el Monto de Cuota que se viene pagando del crédito a comprar", vbInformation, "Aviso"
'            EnfocaControl txtMontoCuota
'            Exit Function
'        End If
    '**COMENTADO BY ARLO20171127 ERS070 - 2017
    End If
    If CDbl(txtMontoCuota.Text) >= CDbl(txtSaldoComprar.Text) Then
        MsgBox "El Monto de Cuota no puede ser mayor o igual al Saldo a Comprar", vbInformation, "Aviso"
        EnfocaControl txtMontoCuota
        Exit Function
    End If
    If CDbl(txtMontoCuota.Text) >= CDbl(txtMontoDesemb.Text) Then
        MsgBox "El Monto de Cuota no puede ser mayor o igual al Monto Desembolsado", vbInformation, "Aviso"
        EnfocaControl txtMontoCuota
        Exit Function
    End If
    
    '**ARLO20180528 INICIO
        If Me.chkCreditoAnt = vbChecked Then
                lsFechaCan = ValidaFecha(Me.txtFecCanledadoAnt)
        If Len(lsFechaCan) > 0 Then
            MsgBox lsFechaCan, vbInformation, "Aviso"
            EnfocaControl txtFecCanledadoAnt
            Exit Function
        End If
        If CDate(txtFecCanledadoAnt) > gdFecSis Then
            MsgBox "La fecha de desembolso no puede ser mayor o igual a la fecha del sistema", vbInformation, "Aviso"
            EnfocaControl txtFecCanledadoAnt
            Exit Function
        End If
        If CDbl(txtMontoCuotaAnt.Text) >= CDbl(txtMontoDesembAnt.Text) Then
            MsgBox "El Monto de Cuota no puede ser mayor o igual al Monto Desembolsado", vbInformation, "Aviso"
            EnfocaControl txtMontoCuotaAnt
            Exit Function
        End If
        If Not IsNumeric(txtMontoDesembAnt.Text) Then
            MsgBox "Ud. debe ingresar el monto de desembolso", vbInformation, "Aviso"
            EnfocaControl txtMontoDesembAnt
            Exit Function
        ElseIf CDbl(txtMontoDesembAnt.Text) <= 0 Then
            MsgBox "Ud. debe ingresar el monto de desembolso", vbInformation, "Aviso"
            EnfocaControl txtMontoDesembAnt
            Exit Function
        End If
        If Not IsNumeric(txtMontoCuotaAnt.Text) Then
            MsgBox "Ud. debe ingresar el Monto de Cuota que se viene pagando del crédito a comprar", vbInformation, "Aviso"
            EnfocaControl txtMontoCuotaAnt
            Exit Function
        End If
        Call txtMontoDesembAnt_LostFocus
    End If
    '**ARLO20180528 INICIO
    
    
    '**ARLO20171127 ERS070 - 2017
    If Me.chkRecompra = vbUnchecked Then
        If (Me.txtMontoDesemb <> Me.txtSaldoComprar) Then
            '**ARLO20180712 ERS042 - 2018
            Set objProducto = New COMDCredito.DCOMCredito     '**ARLO20180712 ERS042 - 2018
            If objProducto.GetResultadoCondicionCatalogo("N0000091", fnTpoProducto) Then
            'If (fnTpoProducto <> 704) Then '**ARLO20180317 ERS 070 - 2017 ANEXO 02
            '**ARLO20180712 ERS042 - 2018
                '**ARLO20180526
                If Me.chkCreditoAnt = vbChecked Then
                    txtFecCanledadoAnt.Text = Format(txtFecCanledadoAnt.Text, "dd/mm/yyyy")
                    If DateDiff("M", txtFecCanledadoAnt.Text, gdFecSis) <= 3 Then
                        If Not ((CDbl(txtMontoDesembAnt.Text) >= CDbl(txtMontoDesemb)) Or (CDbl(txtMontoCuotaAnt) >= CDbl(txtMontoCuota))) Then
                            If Not ValidaCuotas(txtNroCuotaPaga.valor) Then
                                EnfocaControl Me.txtNroCuotaPaga
                                Exit Function
                            End If
                        End If
                    ElseIf Not ValidaCuotas(txtNroCuotaPaga.valor) Then
                                EnfocaControl Me.txtNroCuotaPaga
                                Exit Function
                    End If
                ElseIf Me.chkCreditoAnt = vbUnchecked Then
                '**ARLO20180526
                    If Not ValidaCuotas(txtNroCuotaPaga.valor) Then
                        EnfocaControl Me.txtNroCuotaPaga
                        Exit Function
                    End If
                End If
            End If '**ARLO20180317 ERS 070 - 2017 ANEXO 02
            If CDbl(txtMontoCuota.Text) <= 0 Then
                MsgBox "Ud. debe ingresar el Monto de Cuota que se viene pagando del crédito a comprar", vbInformation, "Aviso"
                EnfocaControl txtMontoCuota
                Exit Function
            End If
        End If
        
    ElseIf Me.chkRecompra = vbChecked Then
    
        If CDbl(txtMontoCuota.Text) < 0 Then
            MsgBox "Ud. debe ingresar el Monto de Cuota que se viene pagando del crédito a comprar", vbInformation, "Aviso"
            EnfocaControl txtMontoCuota
            Exit Function
        End If
    End If
    '**ARLO20171127 ERS070 - 2017
    
    'Validar que no se repite la IFI, el Nro. de Crédito y la moneda
    If fnInicio = CompraDeudaRegistrar Then
        lnIndexValida = 0
    ElseIf fnInicio = CompraDeudaModificar Then
        lnIndexValida = fnIndexValida
    End If
    
    For i = 1 To UBound(fvListaValida)
        If i <> lnIndexValida Then 'Que no se valide el actual item que se está editando
            If Right(cmbIFI, 13) = fvListaValida(i).sIFICod And _
                    Trim(txtNroCredito.Text) = fvListaValida(i).sNroCredito And _
                    CInt(Right(cmbMoneda.Text, 3)) = fvListaValida(i).nMoneda Then
                MsgBox "La Institución Financiera y el Nro. de Crédito ya se ha registrado anteriormente, favor de verificar!!", vbInformation, "Aviso"
                EnfocaControl txtNroCredito
                Exit Function
            End If
        End If
    Next
    
    ValidaDatos = True
End Function
Private Sub cmdCancelar_Click()
    fbAceptar = False
    Unload Me
End Sub
Private Sub txtSaldoComprar_LostFocus()
    txtSaldoComprar.Text = Format(txtSaldoComprar, "#,##0.00")
End Sub
'**ARLO20171127 ERS070 - 2017
Private Sub CmbDestino_Click()
    If (CInt(Trim(Right(cmbDestino.Text, 3))) = 3) Then
        Me.Label14.Caption = "Monto de linea :"
        Me.Label11.Caption = "Saldo a comprar " & Chr(13) & " de Linea :"
    Else
        Me.Label14.Caption = "Monto desembolso :"
        Me.Label11.Caption = "Saldo a comprar :"
    End If
End Sub
'**ARLO20171127 ERS070 - 2017

Private Sub chkCreditoAnt_Click()
    If Me.chkCreditoAnt = vbUnchecked Then
        Height = fnFormTamanioIni
    ElseIf Me.chkCreditoAnt = vbChecked Then
        Height = fnFormTamanioActiva
    End If
End Sub
Private Sub txtMontoDesembAnt_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMontoDesembAnt, KeyAscii, 15)
    If KeyAscii = 13 Then
        EnfocaControl txtFecCanledadoAnt
    End If
End Sub
Private Sub txtFecCanledadoAnt_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtMontoCuotaAnt
    End If
    fEnfoque txtMontoCuotaAnt
End Sub
Private Sub txtMontoCuotaAnt_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMontoCuotaAnt, KeyAscii, 15)
    If KeyAscii = 13 Then
        EnfocaControl cmdAceptar
    End If
End Sub
Private Sub txtMontoDesembAnt_LostFocus()
    txtMontoDesembAnt.Text = Format(txtMontoDesembAnt, "#,##0.00")
End Sub
Private Sub txtMontoCuotaAnt_LostFocus()
    txtMontoCuotaAnt.Text = Format(txtMontoCuotaAnt, "#,##0.00")
End Sub
Private Function ValidaCuotas(pnCuotas) As Boolean
        If CInt(pnCuotas) <= 5 Then
            MsgBox "El Nro de Cuotas pagadas debe ser minimo 6", vbInformation, "Aviso"
            EnfocaControl Me.txtNroCuotaPaga
            Exit Function
        End If
        ValidaCuotas = True
End Function
