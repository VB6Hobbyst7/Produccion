VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaGenCompraMEEfect 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3960
   ClientLeft      =   2025
   ClientTop       =   2310
   ClientWidth     =   5490
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajaGenCompraMEEfect.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   3975
      TabIndex        =   6
      Top             =   3480
      Width           =   1350
   End
   Begin VB.Frame FraSale 
      Height          =   840
      Left            =   120
      TabIndex        =   16
      Top             =   1545
      Width           =   5190
      Begin VB.CheckBox chkBillSale 
         Caption         =   "Billetaje que &Sale"
         Height          =   420
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   270
         Width           =   1785
      End
      Begin VB.Label lblSimbSale 
         AutoSize        =   -1  'True
         Caption         =   "S/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2970
         TabIndex        =   20
         Top             =   375
         Width           =   195
      End
      Begin VB.Label lblTotalSale 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00004080&
         Height          =   360
         Left            =   3300
         TabIndex        =   18
         Top             =   300
         Width           =   1740
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   225
         Left            =   2235
         TabIndex        =   17
         Top             =   345
         Width           =   510
      End
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   810
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2595
      Width           =   5160
   End
   Begin VB.Frame fraIngresa 
      Height          =   840
      Left            =   120
      TabIndex        =   12
      Top             =   705
      Width           =   5190
      Begin VB.CheckBox chkBillIngresa 
         Caption         =   "Billetaje que &Ingresa"
         Height          =   420
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label lblSimbIngresa 
         AutoSize        =   -1  'True
         Caption         =   "S/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2925
         TabIndex        =   19
         Top             =   360
         Width           =   195
      End
      Begin VB.Label lbltotalIngresa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   3300
         TabIndex        =   14
         Top             =   300
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   2250
         TabIndex        =   13
         Top             =   360
         Width           =   510
      End
   End
   Begin VB.Frame fraTipCambio 
      Caption         =   "Ti&po de Cambio "
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
      Height          =   645
      Left            =   120
      TabIndex        =   8
      Top             =   60
      Width           =   3255
      Begin VB.TextBox txtTCBanco 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2145
         TabIndex        =   0
         Text            =   "0.00"
         Top             =   225
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fijo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   285
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1530
         TabIndex        =   10
         Top             =   270
         Width           =   555
      End
      Begin VB.Label lblTCFijo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   540
         TabIndex        =   9
         Top             =   225
         Width           =   825
      End
   End
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   315
      Left            =   4185
      TabIndex        =   1
      Top             =   270
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   2640
      TabIndex        =   5
      Top             =   3480
      Width           =   1350
   End
   Begin VB.Label Label5 
      Caption         =   "Glosa :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   180
      TabIndex        =   15
      Top             =   2385
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3600
      TabIndex        =   7
      Top             =   315
      Width           =   540
   End
End
Attribute VB_Name = "frmCajaGenCompraMEEfect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCtaContDebe As String
Dim lsCtaContHaber As String
Dim oOpe As DOperacion
Dim lbSalir As Boolean
Dim rsBillSal As ADODB.Recordset
Dim rsMonSal As ADODB.Recordset
Dim rsBillIng As ADODB.Recordset
Dim rsMonIng As ADODB.Recordset

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub chkBillIngresa_Click()
Set rsBillIng = Nothing
Set rsMonIng = Nothing
If chkBillIngresa.value = 1 Then
   If gsOpeCod <> gOpeMEVentaEfec Then
        If Val(Me.txtTCBanco) = 0 Then
            MsgBox "Ingrese el Tipo de Cambio Banco por Favor", vbInformation, "Aviso"
            txtTCBanco.SetFocus
            chkBillIngresa.value = 0
            Exit Sub
        End If
    Else
        If Val(lblTotalSale) = 0 Then
            MsgBox "Billetaje de Salida no se ha Digitado", vbInformation, "Aviso"
            chkBillIngresa.value = 0
            chkBillSale.SetFocus
            Exit Sub
        End If
    End If
    frmCajaGenEfectivo.Inicio gsOpeCod, gsOpeDesc, IIf(gsOpeCod = gOpeMEVentaEfec, CCur(lbltotalIngresa), 0), Mid(lsCtaContDebe, 3, 1), False
    If frmCajaGenEfectivo.lbOk Then
        Set rsBillIng = frmCajaGenEfectivo.rsBilletes
        Set rsMonIng = frmCajaGenEfectivo.rsMonedas
        lbltotalIngresa = frmCajaGenEfectivo.lblTotal
        If gsOpeCod <> gOpeMEVentaEfec Then
            lblTotalSale = Format(CCur(lbltotalIngresa) * CCur(txtTCBanco), "#,#0.00")
            chkBillSale.SetFocus
        Else
            txtMovDesc.SetFocus
        End If
        
    Else
        chkBillIngresa.value = 0
        If gsOpeCod <> gOpeMEVentaEfec Then
            lbltotalIngresa = "0.00"
            lblTotalSale = "0.00"
            chkBillSale.value = 0
        End If
    End If
    Set frmCajaGenEfectivo = Nothing
Else
    If gsOpeCod <> gOpeMEVentaEfec Then
        lbltotalIngresa = "0.00"
        lblTotalSale = "0.00"
        chkBillSale.value = 0
    End If
End If
End Sub

Private Sub chkBillSale_Click()
Set rsBillSal = Nothing
Set rsMonSal = Nothing
If chkBillSale.value = 1 Then
    If gsOpeCod <> gOpeMEVentaEfec Then
        If Val(lbltotalIngresa) = 0 Then
            MsgBox "Billetaje de Ingreso no se ha Digitado", vbInformation, "Aviso"
            chkBillSale.value = 0
            Me.chkBillIngresa.SetFocus
            Exit Sub
        End If
    Else
        If Val(txtTCBanco) = 0 Then
            MsgBox "Ingrese el Tipo de Cambio Banco por Favor", vbInformation, "Aviso"
            txtTCBanco.SetFocus
            chkBillSale.value = 0
            Exit Sub
        End If
    End If
    frmCajaGenEfectivo.Inicio gsOpeCod, gsOpeDesc, IIf(gsOpeCod = gOpeMEVentaEfec, 0, CCur(lblTotalSale)), Mid(lsCtaContHaber, 3, 1), False
    If frmCajaGenEfectivo.lbOk Then
        Set rsBillSal = frmCajaGenEfectivo.rsBilletes
        Set rsMonSal = frmCajaGenEfectivo.rsMonedas
        lblTotalSale = frmCajaGenEfectivo.lblTotal
        If gsOpeCod = gOpeMEVentaEfec Then
            lbltotalIngresa = Format(CCur(lblTotalSale) * CCur(txtTCBanco), "#,#0.00")
            chkBillIngresa.SetFocus
        Else
            txtMovDesc.SetFocus
        End If
    Else
        chkBillSale.value = 0
        If gsOpeCod = gOpeMEVentaEfec Then
            lbltotalIngresa = "0.00"
            lblTotalSale = "0.00"
            chkBillIngresa.value = 0
        End If
    End If
    Set frmCajaGenEfectivo = Nothing
Else
    If gsOpeCod = gOpeMEVentaEfec Then
        lbltotalIngresa = "0.00"
        lblTotalSale = "0.00"
        chkBillIngresa.value = 0
    End If
End If
End Sub

Private Sub cmdAceptar_Click()
Dim lsMovNro As String
Dim oCon As NContFunciones
Dim oCaja As nCajaGeneral

Set oCaja = New nCajaGeneral
Set oCon = New NContFunciones
If Valida = False Then Exit Sub

If MsgBox("Desea grabar la Operación??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    lsMovNro = oCon.GeneraMovNro(CDate(txtFecha), gsCodAge, gsCodUser)
    If oCaja.GrabaCompraVentaEfectivo(lsMovNro, gsOpeCod, _
                        txtMovDesc, rsBillIng, rsMonIng, rsBillSal, rsMonSal, _
                        lsCtaContDebe, lsCtaContHaber, CCur(txtTCBanco), CCur(lblTCFijo)) = 0 Then
        
        ImprimeAsientoContable lsMovNro, , , , True, True, txtMovDesc
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Tipo Cambio Banco : " & txtTCBanco.Text & " Tipo Cambio Fijo : " & lblTCFijo _
            & " Billetaje que Ingresa : " & lbltotalIngresa & " Billeta que Sale : " & lblTotalSale & " Glosa  : " & txtMovDesc.Text
            Set objPista = Nothing
            '*******
            
        Set oCaja = Nothing
        Set oCon = Nothing
        If MsgBox("Desea Realizar otra Operación", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            chkBillIngresa.value = 0
            chkBillSale.value = 0
            lbltotalIngresa = "0.00"
            lblTotalSale = "0.00"
            Set rsBillIng = Nothing
            Set rsBillSal = Nothing
            Set rsMonIng = Nothing
            Set rsMonSal = Nothing
            If fraTipCambio.Visible Then
               txtTCBanco.SetFocus
            Else
               chkBillIngresa.SetFocus
            End If
        Else
            Unload Me
        End If
    End If
End If

End Sub
Function Valida() As Boolean
Valida = True
If Val(txtTCBanco) = 0 Then
    MsgBox "Tipo de Cambio no ingresado", vbInformation, "Aviso"
    txtTCBanco.SetFocus
    Valida = False
    Exit Function
End If
If ValidaFechaContab(txtFecha, gdFecSis) = False Then
    Valida = False
    Exit Function
End If
If chkBillIngresa.value = 0 Or Val(lbltotalIngresa) = 0 Then
    MsgBox "Billeta de ingreso no ha sido Digitado", vbInformation, "Aviso"
    Valida = False
    chkBillIngresa.SetFocus
    Exit Function
End If
If chkBillSale.value = 0 Or Val(lblTotalSale) = 0 Then
    MsgBox "Billeta de Salida no ha sido Digitado", vbInformation, "Aviso"
    Valida = False
    chkBillSale.SetFocus
    Exit Function
End If

End Function
Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If lbSalir Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
Set oOpe = New DOperacion
Set rsBillSal = New ADODB.Recordset
Set rsMonSal = New ADODB.Recordset
Set rsBillIng = New ADODB.Recordset
Set rsMonIng = New ADODB.Recordset
On Error GoTo CajaGenCompraVentaErr
CentraForm Me
lbSalir = False
Me.Caption = gsOpeDesc
txtMovDesc = gsOpeDesc

txtFecha = gdFecSis
lsCtaContDebe = oOpe.EmiteOpeCta(gsOpeCod, "D")
lsCtaContHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
If lsCtaContDebe = "" Or lsCtaContHaber = "" Then
    MsgBox "No se han definido Cuentas Contables corectamente", vbInformation, "Aviso"
    lbSalir = True
    Exit Sub
End If
If gsOpeCod = OpeCGOtrosOpeEfecCamb Or gsOpeCod = OpeCGOtrosOpeEfecCambme Then
   fraTipCambio.Visible = False
   txtTCBanco = 1
End If
If gsOpeCod = gOpeMEVentaEfec Then
   FraSale.Top = Me.fraTipCambio.Top + Me.fraTipCambio.Height + 20
   fraIngresa.Top = FraSale.Top + FraSale.Height + 10
End If
lblSimbIngresa = IIf(Mid(lsCtaContDebe, 3, 1) = gMonedaNacional, gcPEN_SIMBOLO, "$.") 'marg ers044-2016
lblSimbSale = IIf(Mid(lsCtaContHaber, 3, 1) = gMonedaNacional, gcPEN_SIMBOLO, "$.") 'marg ers044-2016
lbltotalIngresa.BackColor = IIf(Mid(lsCtaContDebe, 3, 1) = gMonedaNacional, vbWhite, &HC0FFC0)
lblTotalSale.BackColor = IIf(Mid(lsCtaContHaber, 3, 1) = gMonedaNacional, vbWhite, &HC0FFC0)

If gnTipCambio = 0 Then  'Verificamos el Tipo de Cambio Fijo
    Dim oTC As New nTipoCambio
    gnTipCambio = oTC.EmiteTipoCambio(txtFecha, TCFijoDia)
    If gnTipCambio = 0 Then
        lbSalir = True
        Exit Sub
    End If
End If
'lblTCFijo = Format(gnTipCambio, gsFormatoNumeroView)
lblTCFijo = Format(gnTipCambio, "#,#0.00##")
Exit Sub
CajaGenCompraVentaErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oOpe = Nothing
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFecha) = False Then Exit Sub
    If gsOpeCod <> gOpeMEVentaEfec Then
        chkBillIngresa.SetFocus
    Else
        chkBillSale.SetFocus
    End If
End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdAceptar.SetFocus
End If
End Sub
Private Sub txtTCBanco_GotFocus()
fEnfoque txtTCBanco
End Sub
Private Sub txtTCBanco_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTCBanco, KeyAscii, 8, 3)
If KeyAscii = 13 Then
    chkBillIngresa.value = 0
    chkBillSale.value = 0
    Me.lbltotalIngresa = "0.00"
    Me.lblTotalSale = "0.00"
    txtFecha.SetFocus
End If
End Sub
Private Sub txtTCBanco_LostFocus()
txtTCBanco = Format(txtTCBanco, "#,#0.000")
End Sub

