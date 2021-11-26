VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCartaFianzaIngreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartas Fianza: Ingreso"
   ClientHeight    =   4110
   ClientLeft      =   2460
   ClientTop       =   2640
   ClientWidth     =   6570
   Icon            =   "frmCartaFianzaIngreso.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTipoCambio 
      Height          =   570
      Left            =   150
      TabIndex        =   25
      Top             =   3450
      Width           =   2115
      Begin VB.Label lblTipoCambio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1035
         TabIndex        =   27
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "T.Cambio :"
         Height          =   195
         Left            =   135
         TabIndex        =   26
         Top             =   225
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   4200
      TabIndex        =   12
      Top             =   3585
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   5325
      TabIndex        =   13
      Top             =   3585
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   3075
      TabIndex        =   11
      Top             =   3585
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Carta Fianza"
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
      Height          =   1860
      Left            =   150
      TabIndex        =   18
      Top             =   630
      Width           =   6285
      Begin Sicmact.TxtBuscar txtCodBanco 
         Height          =   330
         Left            =   1020
         TabIndex        =   4
         Top             =   630
         Width           =   1515
         _ExtentX        =   2672
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
         sTitulo         =   ""
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
         Height          =   330
         Left            =   4590
         TabIndex        =   9
         Top             =   1410
         Width           =   1560
      End
      Begin MSMask.MaskEdBox txtFechaVenc 
         Height          =   330
         Left            =   990
         TabIndex        =   8
         Top             =   1410
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNumCarta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   990
         MaxLength       =   20
         TabIndex        =   2
         Top             =   240
         Width           =   2445
      End
      Begin MSMask.MaskEdBox txtFechaEmision 
         Height          =   300
         Left            =   4980
         TabIndex        =   3
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Sicmact.TxtBuscar txtCodPers 
         Height          =   330
         Left            =   1020
         TabIndex        =   6
         Top             =   1020
         Width           =   1515
         _ExtentX        =   2672
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
         sTitulo         =   ""
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Emisión :"
         Height          =   195
         Left            =   3795
         TabIndex        =   24
         Top             =   285
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
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
         Left            =   3900
         TabIndex        =   23
         Top             =   1455
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Vence el :"
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   1470
         Width           =   720
      End
      Begin VB.Label lblNomPers 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2535
         TabIndex        =   7
         Top             =   1020
         Width           =   3615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor :"
         Height          =   195
         Left            =   135
         TabIndex        =   21
         Top             =   1035
         Width           =   825
      End
      Begin VB.Label lblnombanco 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2535
         TabIndex        =   5
         Top             =   630
         Width           =   3615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   645
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Carta Nº :"
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
         Left            =   135
         TabIndex        =   19
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "&Glosa"
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
      Height          =   945
      Left            =   150
      TabIndex        =   17
      Top             =   2505
      Width           =   6285
      Begin VB.TextBox txtMovDesc 
         Height          =   585
         Left            =   150
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   225
         Width           =   6000
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operación"
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
      Height          =   630
      Left            =   165
      TabIndex        =   14
      Top             =   -15
      Width           =   6270
      Begin VB.TextBox txtOpeCod 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   810
         TabIndex        =   0
         Top             =   210
         Width           =   1080
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   4965
         TabIndex        =   1
         Top             =   210
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
         Caption         =   "Código"
         Height          =   255
         Left            =   135
         TabIndex        =   16
         Top             =   285
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha :"
         Height          =   240
         Left            =   4335
         TabIndex        =   15
         Top             =   240
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmCartaFianzaIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cmdAceptar_Click()
Dim sql As String
Dim lnMontoMov As Currency
Dim lbTrans As Boolean
Dim lsCtaContDebe  As String
Dim lsCtaContHaber As String
Dim lsDatosCarta   As String
Dim lsImpre        As String
On Error GoTo ErrorGrabacion
Dim oOpe As New DOperacion

If Valida = True Then
'   If VerificaCierreCaja(txtFecha) = False Then Exit Sub
If Not PermiteModificarAsiento(Format(txtFecha, gsFormatoMovFecha), False) Then
    MsgBox "No puede realizarse Operación con fecha de Mes ya cerrado", vbInformation, "¡Aviso!"
    Exit Sub
End If
   If MsgBox(" ¿ Desea Grabar la Operación ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
      Dim oCaja As New nCajaGeneral
      Dim oMov  As New DMov
      Dim oImp  As New NContImprimir
      lsCtaContDebe = oOpe.EmiteOpeCta(gsOpeCod, "D")
      lsCtaContHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
      gsMovNro = oMov.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
      If oCaja.GrabaCartaFianza(gsMovNro, gsOpeCod, txtMovDesc, TpoDocCartaFianza, Me.txtNumCarta, Format(txtFechaEmision, gsFormatoFecha), _
                                Mid(txtCodBanco, 4, 13), txtCodPers, Format(txtFechaVenc, gsFormatoFecha), lsCtaContDebe, lsCtaContHaber, nVal(txtMonto)) = 0 Then
         lsDatosCarta = lsDatosCarta & "  Inst.Financiera   : " & Me.lblnombanco & oImpresora.gPrnSaltoLinea
         lsDatosCarta = lsDatosCarta & "  Proveedor         : " & Me.lblNomPers & oImpresora.gPrnSaltoLinea
         lsDatosCarta = lsDatosCarta & "  Vencimiento       : " & Me.txtFechaVenc
         lsImpre = oImp.ImprimeAsientoContable(gsMovNro, gnLinPage, gnColPage, "INGRESO DE CARTA FIANZA", lsDatosCarta, "19")
         EnviaPrevio lsImpre, "INGRESO DE CARTA FIANZA", gnLinPage, False
        
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
        Set objPista = Nothing
        '****
         
         If MsgBox(" ¿ Desea Registrar nuevo Carta Fianza ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbYes Then
            cmdCancelar_Click
         Else
            Unload Me
         End If
      End If
    End If
End If
Exit Sub
ErrorGrabacion:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub cmdCancelar_Click()
txtMovDesc = ""
txtCodBanco = ""
txtNumCarta = ""
txtCodPers = ""
txtCodBanco.Enabled = True
txtFechaVenc = "__/__/____"
txtMonto = "0.00"
lblnombanco = ""
lblNomPers = ""
txtFechaVenc = gdFecSis
txtFechaEmision = gdFecSis
txtNumCarta.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
txtOpeCod = gsOpeCod
txtFecha = gdFecSis

txtCodPers.TipoBusPers = BusPersDocumento
txtCodPers.TipoBusqueda = BuscaPersona

Dim oOpe As New DOperacion
txtCodBanco.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")   '  oDCtaIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), gTpoIFBanco + gTpoCtaIFCtaCte + gTpoCtaIFCtaAho)
Set oOpe = Nothing

txtMonto = "0.00"
txtFechaVenc = gdFecSis
txtFechaEmision = gdFecSis
If Mid(gsOpeCod, 3, 1) = "1" Then
    fraTipoCambio.Visible = False
Else
    fraTipoCambio.Visible = True
    lblTipoCambio = gnTipCambio
End If
End Sub

Private Sub txtCodBanco_EmiteDatos()
lblnombanco = txtCodBanco.psDescripcion
If txtCodBanco <> "" Then
   txtCodPers.SetFocus
End If
End Sub

Private Sub txtCodBanco_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtCodPers.SetFocus
End If
End Sub

Private Sub txtCodPers_EmiteDatos()
lblNomPers = txtCodPers.psDescripcion
If txtCodPers <> "" Then
   txtFechaVenc.SetFocus
End If
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFecha) Then
        txtNumCarta.SetFocus
    End If
End If
End Sub

Private Sub txtFechaEmision_GotFocus()
fEnfoque txtFechaEmision
End Sub

Private Sub txtFechaEmision_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CDate(Me.txtFechaEmision) > gdFecSis Then
        MsgBox "Fecha de Emisión no Puede ser mayor a la Actual", vbInformation, "Aviso"
        txtFechaEmision.SetFocus
    Else
        txtCodBanco.SetFocus
    End If
End If
End Sub

Private Sub txtFechaVenc_GotFocus()
fEnfoque txtFechaVenc
End Sub

Private Sub txtFechaVenc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFechaVenc) = True Then
        If CDate(txtFechaVenc) < CDate(txtFechaEmision) Then
            MsgBox "Fecha de Vencimiento no puede ser menor a la Fecha de emisión", vbInformation, "Aviso"
            txtFechaVenc.SetFocus
        Else
            txtMonto.SetFocus
        End If
    End If
End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdAceptar.SetFocus
End If
End Sub
Private Sub txtmonto_GotFocus()
fEnfoque txtMonto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMonto, KeyAscii)
If KeyAscii = 13 Then
   txtMovDesc.SetFocus
End If
End Sub

Private Sub txtMonto_LostFocus()
txtMonto = Format(txtMonto, gsFormatoNumeroView)
End Sub

Private Sub txtNumCarta_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   txtFechaEmision.SetFocus
End If
End Sub
Private Function Valida() As Boolean
Valida = False
If Len(Trim(Me.txtNumCarta)) = 0 Then
    MsgBox "Numero de Carta Fianza no Válido", vbInformation, "Aviso"
    txtNumCarta.SetFocus
    Exit Function
End If

If ValFecha(Me.txtFechaVenc) = False Then
   Exit Function
Else
   If CDate(txtFechaEmision) > gdFecSis Then
      MsgBox "Fecha no Puede ser mayor a la Actual", vbInformation, "Aviso"
      txtFechaEmision.SetFocus
      Exit Function
   End If
End If
If Len(Trim(txtCodBanco)) = 0 Then
    MsgBox "Objeto Banco no Válido", vbInformation, "Aviso"
    txtCodBanco.SetFocus
    Exit Function
End If

If Len(Trim(Me.txtCodPers)) = 0 Then
    MsgBox "Código de Persona no Válido", vbInformation, "Aviso"
    txtCodPers.SetFocus
    Exit Function
End If
If ValFecha(txtFechaVenc) = False Then
    Exit Function
End If
If Val(Me.txtMonto) = 0 Then
    MsgBox "Monto de Operación no Válido", vbInformation, "Aviso"
    txtMonto.SetFocus
    Exit Function
End If
If Len(Trim(Me.txtMovDesc)) = 0 Then
    MsgBox "Descripcion de Operacion no Válida", vbInformation, "Aviso"
    Me.txtMovDesc.SetFocus
    Exit Function
End If
Valida = True
End Function
