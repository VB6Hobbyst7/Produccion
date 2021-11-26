VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaGenTransf 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4485
   ClientLeft      =   1815
   ClientTop       =   2250
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajaGenTransf.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   6285
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   3570
      Width           =   1905
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
      Left            =   105
      TabIndex        =   19
      Top             =   3420
      Width           =   1755
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
         TabIndex        =   21
         Top             =   285
         Width           =   330
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
         TabIndex        =   20
         Top             =   225
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6930
      TabIndex        =   9
      Top             =   4050
      Width           =   1275
   End
   Begin VB.Frame FraDestino 
      Caption         =   "D&estino de Recursos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   105
      TabIndex        =   14
      Top             =   1560
      Width           =   8115
      Begin VB.CheckBox chkDocDest 
         Caption         =   "Documen&to"
         Height          =   345
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   615
         Width           =   1140
      End
      Begin Sicmact.TxtBuscar txtBuscaCtaIFDest 
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   635
         Appearance      =   1
         BackColor       =   16777215
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
         ForeColor       =   8388608
      End
      Begin VB.Label lblDescIFDest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2745
         TabIndex        =   16
         Top             =   255
         Width           =   5100
      End
      Begin VB.Label lblDescCtaIFDest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2745
         TabIndex        =   15
         Top             =   600
         Width           =   3960
      End
   End
   Begin VB.Frame fraOrigen 
      Caption         =   "&Fuente de Recursos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1095
      Left            =   105
      TabIndex        =   11
      Top             =   435
      Width           =   8115
      Begin VB.CheckBox chkDocFuente 
         Caption         =   "&Documento"
         Height          =   345
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   630
         Width           =   1140
      End
      Begin Sicmact.TxtBuscar txtBuscaCtaIFFte 
         Height          =   360
         Left            =   150
         TabIndex        =   1
         Top             =   255
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   635
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
         ForeColor       =   4210816
      End
      Begin VB.Label lblDescIFFte 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00404080&
         Height          =   315
         Left            =   2700
         TabIndex        =   13
         Top             =   270
         Width           =   5160
      End
      Begin VB.Label lblDescCtaIFFte 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00404080&
         Height          =   315
         Left            =   2700
         TabIndex        =   12
         Top             =   615
         Width           =   3990
      End
   End
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   315
      Left            =   7080
      TabIndex        =   0
      Top             =   45
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Frame fraCambioBanco 
      Caption         =   "Cambio Banco"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   645
      Left            =   1935
      TabIndex        =   17
      Top             =   3420
      Width           =   1755
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
         Left            =   675
         TabIndex        =   6
         Text            =   "0.0000"
         Top             =   225
         Width           =   930
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
         Left            =   90
         TabIndex        =   18
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   585
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2730
      Width           =   8115
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5670
      TabIndex        =   8
      Top             =   4050
      Width           =   1275
   End
   Begin VB.CommandButton cmdtransferir 
      Caption         =   "&Transferir"
      Height          =   375
      Left            =   5670
      TabIndex        =   24
      Top             =   4050
      Width           =   1275
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
      TabIndex        =   25
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Label lblSimbolo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   5865
      TabIndex        =   23
      Top             =   3622
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monto "
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
      Height          =   210
      Left            =   5115
      TabIndex        =   22
      Top             =   3637
      Width           =   570
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6405
      TabIndex        =   10
      Top             =   90
      Width           =   555
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   405
      Left            =   4950
      Top             =   3540
      Width           =   3255
   End
End
Attribute VB_Name = "frmCajaGenTransf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim lbSalir As Boolean
Dim oCtaIf As NCajaCtaIF
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim lsTpoDoc As String
Dim lsNroDocFte As String
Dim lsDocumentoFte As String
Dim lsNroDocDest As String
Dim lsDocumentoDest As String
Dim lnTipoCambioDia As Currency 'ALPA 20140331
Dim oOpe As DOperacion
Dim lsMovNroCV As String 'ALPA 20140328*******************

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub chkDocDest_Click()
Dim oDocPago As clsDocPago
Set oDocPago = New clsDocPago
lsNroDocDest = ""
lsDocumentoDest = ""
If chkDocDest.value = 1 Then
    If ValidaInterfaz = False Then
        chkDocDest.value = 0
        Exit Sub
    End If
    Select Case Val(lsTpoDoc)
        Case TpoDocCarta
            oDocPago.InicioCarta lsNroDocDest, Mid(txtBuscaCtaIFFte, 4, 13), gsOpeCod, _
                                gsOpeDesc, txtMovDesc, "", CCur(txtImporte), gdFecSis, _
                                Trim(lblDescIFFte), Trim(lblDescCtaIFFte), Trim(lblDescIFDest), Trim(lblDescCtaIFDest), ""
            If oDocPago.vbOk Then
                lsNroDocDest = oDocPago.vsNroDoc
                lsDocumentoDest = oDocPago.vsFormaDoc
                If lsNroDocFte = lsNroDocDest Then
                    MsgBox "Nro  de documento ya ha sido ingresado", vbInformation, "Aviso"
                    chkDocDest.value = 0
                    Set oDocPago = Nothing
                    Exit Sub
                End If
                If cmdAceptar.Visible Then cmdAceptar.SetFocus
                If cmdtransferir.Visible Then cmdtransferir.SetFocus
            Else
                chkDocDest.value = 0
                Set oDocPago = Nothing
                Exit Sub
            End If
    End Select
End If
Set oDocPago = Nothing

End Sub
Function ValidaInterfaz() As Boolean
ValidaInterfaz = True
If txtBuscaCtaIFFte = "" Then
    MsgBox "Cuenta de Institucion Financiera Fuente no Ingresada ", vbInformation, "Aviso"
    txtBuscaCtaIFFte.SetFocus
    ValidaInterfaz = False
    Exit Function
End If
If txtBuscaCtaIFDest = "" Then
    MsgBox "Cuenta de Institucion Financiera Destino no Ingresada ", vbInformation, "Aviso"
    txtBuscaCtaIFDest.SetFocus
    ValidaInterfaz = False
    Exit Function
End If
If fraCambioBanco.Visible Then
    If Val(txtTCBanco) = 0 Then
        MsgBox "Tipo de Cambio de Banco no Ingresado", vbInformation, "Aviso"
        txtTCBanco.SetFocus
        ValidaInterfaz = False
        Exit Function
    End If
End If
If fraTipCambio.Visible Then
    If Val(lblTCFijo) = 0 Then
        MsgBox "Tipo de Cambio de fijo no válido", vbInformation, "Aviso"
        ValidaInterfaz = False
        Exit Function
    End If
End If
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Descripción de Operación no Ingresada", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    ValidaInterfaz = False
    Exit Function
End If
If Val(txtImporte) = 0 Then
    MsgBox "Importe de Operación no Ingresado", vbInformation, "Aviso"
    txtImporte.SetFocus
    ValidaInterfaz = False
    Exit Function
End If
If txtBuscaCtaIFFte = txtBuscaCtaIFDest Then
    MsgBox "Cuenta de Origen no puede ser igual que la de Destino", vbInformation, "Aviso"
    txtBuscaCtaIFDest.SetFocus
    ValidaInterfaz = False
    Exit Function
End If
Select Case gsOpeCod
    Case gOpeCGTransfMismoBancoMN, gOpeCGTransfMismoBancoME
        If Mid(txtBuscaCtaIFFte, 1, 16) <> Mid(txtBuscaCtaIFDest, 1, 16) Then
            MsgBox "Entidad Origen debe ser igual que la de Destino", vbInformation, "Aviso"
            txtBuscaCtaIFDest.SetFocus
            ValidaInterfaz = False
            Exit Function
        End If
End Select

End Function

Private Sub chkDocFuente_Click()
Dim oDocPago As clsDocPago
Set oDocPago = New clsDocPago
lsNroDocFte = ""
lsDocumentoFte = ""
If chkDocFuente.value = 1 Then
    If ValidaInterfaz = False Then
        chkDocFuente.value = 0
        Exit Sub
    End If
    Select Case Val(lsTpoDoc)
        Case TpoDocCarta
            oDocPago.InicioCarta lsNroDocFte, Mid(txtBuscaCtaIFFte, 4, 13), gsOpeCod, _
                                gsOpeDesc, txtMovDesc, "", CCur(txtImporte), txtFecha, _
                                Trim(lblDescIFFte), Trim(lblDescCtaIFFte), Trim(lblDescIFDest), Trim(lblDescCtaIFDest), ""
            If oDocPago.vbOk Then
                lsNroDocFte = oDocPago.vsNroDoc
                lsDocumentoFte = oDocPago.vsFormaDoc
                If lsNroDocFte = lsNroDocDest Then
                    MsgBox "Nro de documento ya ha sido ingresado", vbInformation, "Aviso"
                    chkDocFuente.value = 0
                    Set oDocPago = Nothing
                    Exit Sub
                End If
                chkDocDest.SetFocus
            Else
                chkDocFuente.value = 0
                Set oDocPago = Nothing
                Exit Sub
            End If
    End Select
End If
Set oDocPago = Nothing
End Sub

Private Sub cmdAceptar_Click()
Dim oCon As NContFunciones
Dim oCaja As nCajaGeneral
Dim lsCtaContDebe As String
Dim lsCtaContHaber As String
Dim lsMovNro As String
Dim lsDocumento As String

On Error GoTo AceptarErr

Set oCaja = New nCajaGeneral
Set oCon = New NContFunciones

'ALPA 20140328 ********************************************************
If (gsOpeCod = gOpeMECompraAInst Or gsOpeCod = gOpeMEVentaAInst) Then
    Dim oCred  As nTipoCambio
    Set oCred = New nTipoCambio
    Dim lnEstado As Integer
    Dim nContadorVistos As Integer
    If Len(Trim(lsMovNroCV)) > 0 Then
        nContadorVistos = oCred.ObtenerCantidadAprobacionMovCompraVenta(lsMovNroCV, lnEstado)
        If lnEstado = 3 Then
            MsgBox "El permiso de nivel de aprobación fue denegada", vbInformation, "Aviso"
            Call ActualizarNivelAprovacion
            txtTCBanco.Text = lnTipoCambioDia
            Exit Sub
        End If
        If nContadorVistos > 0 Then
            MsgBox "Aun falta mas vistos para la operación, favor coordinar para continuar", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
End If
'**********************************************************************

If ValidaInterfaz = False Then Exit Sub
If lsDocumentoFte = "" Or chkDocFuente.value = 0 Then
    If MsgBox("Documento de Institucion Origen o Fuente no ha sido Ingresado " & vbCrLf & "Desea Continuar??", vbQuestion + vbYesNo) = vbNo Then
        Me.chkDocFuente.SetFocus
        Exit Sub
    End If
End If
If lsDocumentoDest = "" Or chkDocDest.value = 0 Then
    If MsgBox("Documento de Institución Destino no ha sido Ingresado " & vbCrLf & "Desea Continuar??", vbQuestion + vbYesNo) = vbNo Then
        chkDocDest.SetFocus
        Exit Sub
    End If
End If
If Not ValidaFechaContab(txtFecha, gdFecSis) Then
    Exit Sub
End If

Select Case Val(lsTpoDoc)
    Case TpoDocCarta
        lsDocumento = IIf(lsDocumentoFte = "", "", lsDocumentoFte & IIf(lsDocumentoDest = "", "", oImpresora.gPrnSaltoPagina)) & lsDocumentoDest
        
End Select
lsCtaContDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", , txtBuscaCtaIFDest, CtaOBjFiltroIF)
lsCtaContHaber = oOpe.EmiteOpeCta(gsOpeCod, "H", , txtBuscaCtaIFFte, CtaOBjFiltroIF)
If lsCtaContHaber = "" Or lsCtaContDebe = "" Then
    MsgBox "Objetos Seleccionados No pueden generar Correctamente las Cuentas Contables del Movimiento" + vbCrLf + "Por favor Consulte Con Sistemas...", vbExclamation, "Aviso"
    Exit Sub
End If
If MsgBox("Desea grabar la Operación", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    cmdAceptar.Enabled = False
    lsMovNro = oCon.GeneraMovNro(CDate(txtFecha), gsCodAge, gsCodUser)
    If oCaja.GrabaCompraVentaMEIF(lsMovNro, gsOpeCod, txtMovDesc, _
                    lsCtaContDebe, lsCtaContHaber, txtBuscaCtaIFFte, txtBuscaCtaIFDest, _
                    lsTpoDoc, lsNroDocFte, lsTpoDoc, lsNroDocDest, CCur(txtImporte), CCur(txtTCBanco), gdFecSis, CCur(IIf(Trim(lblTCFijo.Caption) = "", 0, lblTCFijo.Caption))) = 0 Then
        'ALPA 20140517*****************************************************************
        Call ActualizarNivelAprovacion
        'ALPA 20140317*****************************************************************
        'Se agregó el campo CCur(IIf(Trim(lblTCFijo.Caption) = "", 0, lblTCFijo.Caption)
        ImprimeAsientoContable lsMovNro, , lsTpoDoc, lsDocumento
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, Me.Caption & "|Fuentes de Recursos : " & txtBuscaCtaIFFte.Text _
            & " Destino Recursos: " & txtBuscaCtaIFDest.Text & " Tipo Cambio " & lblTCFijo & " Importe : " & txtImporte.Text
            Set objPista = Nothing
            '*******
        If MsgBox("Desea Realizar otra Operación", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            Limpiar
        Else
            Unload Me
            Exit Sub
        End If
    End If
    cmdAceptar.Enabled = True
End If
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdCancelar_Click()
If Len(Trim(lsMovNroCV)) > 0 Then
    If MsgBox("Desea eliminar el movimiento de tipo de cambio", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        Call ActualizarNivelAprovacion 'ALPA 20140328**************
    End If
End If
    lsMovNroCV = ""
    lblPermiso.Caption = ""
Unload Me
End Sub

Private Sub cmdTransferir_Click()
Dim oCon As NContFunciones
Dim oCaja As nCajaGeneral
Dim lsCtaContDebe As String
Dim lsCtaContHaber As String
Dim lsCtaContDebeOtr As String
Dim lsMovNro As String
Dim lsMovNroCtr As String
Dim lsDocumento As String
On Error GoTo TransferirErr

Set oCaja = New nCajaGeneral
Set oCon = New NContFunciones

If ValidaInterfaz = False Then Exit Sub
If lsDocumentoFte = "" Or chkDocFuente.value = 0 Then
        If MsgBox("Documento de Institucion Origen o Fuente no ha sido Ingresado " & vbCrLf & "Desea Continuar??", vbQuestion + vbYesNo) = vbNo Then
        Me.chkDocFuente.SetFocus
        Exit Sub
    End If
End If
If lsDocumentoDest = "" Or chkDocDest.value = 0 Then
    If MsgBox("Documento de Institución Destino no ha sido Ingresado " & vbCrLf & "Desea Continuar??", vbQuestion + vbYesNo) = vbNo Then
        chkDocDest.SetFocus
        Exit Sub
    End If
End If
If ValFecha(txtFecha) = False Then
    Exit Sub
End If
Select Case Val(lsTpoDoc)
    Case TpoDocCarta
        lsDocumento = IIf(lsDocumentoFte = "", "", lsDocumentoFte & IIf(lsDocumentoDest = "", "", oImpresora.gPrnSaltoPagina)) & lsDocumentoDest
        
End Select
lsCtaContDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", , txtBuscaCtaIFDest, CtaOBjFiltroIF)
lsCtaContHaber = oOpe.EmiteOpeCta(gsOpeCod, "H", , txtBuscaCtaIFFte, CtaOBjFiltroIF)
If lsCtaContHaber = "" Or lsCtaContDebe = "" Then
    MsgBox "Objetos Seleccionados No pueden generar Correctamente las Cuentas Contables del Movimiento" + vbCrLf + "Por favor Consulte Con Sistemas...", vbExclamation, "Aviso"
    Exit Sub
End If


If MsgBox("Desea grabar la Operación", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oCon.GeneraMovNro(CDate(txtFecha), gsCodAge, gsCodUser)
    If oCaja.GrabaMovimientosCtasIF(lsMovNro, gsOpeCod, txtMovDesc, _
                    lsCtaContDebe, lsCtaContHaber, txtBuscaCtaIFFte, txtBuscaCtaIFDest, _
                    lsTpoDoc, lsNroDocFte, lsTpoDoc, lsNroDocDest, CCur(txtImporte), CCur(txtTCBanco), gdFecSis) = 0 Then
        
        ImprimeAsientoContable lsMovNro, , lsTpoDoc, lsDocumento
'        If MsgBox("Desea Realizar otra Operación", vbYesNo + vbQuestion, "Aviso") = vbYes Then
'           Limpiar
'        Else
'            Unload Me
'        End If
    End If
    
    If gsOpeCod = "401425" Or gsOpeCod = "402425" Then
        'Contra Asiento GITU 03/04/2008
        lsCtaContDebeOtr = oOpe.EmiteOpeCta(gsOpeCod, "D", "1", txtBuscaCtaIFDest, CtaOBjFiltroIF)
        lsMovNroCtr = oCon.GeneraMovNro(CDate(txtFecha), gsCodAge, gsCodUser)
        If oCaja.GrabaMovimientosCtasIF(lsMovNroCtr, gsOpeCod, txtMovDesc, _
                        lsCtaContDebeOtr, lsCtaContDebe, txtBuscaCtaIFFte, txtBuscaCtaIFDest, _
                        lsTpoDoc, lsNroDocFte, lsTpoDoc, lsNroDocDest, CCur(txtImporte), CCur(txtTCBanco), gdFecSis) = 0 Then
           ImprimeAsientoContable lsMovNroCtr, , lsTpoDoc, lsDocumento
        End If
    End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            Dim lsCaption As String
            lsCaption = Replace(Me.Caption, "'", "")
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, " ", lsCaption & "|Fuentes de Recursos : " & txtBuscaCtaIFFte.Text
            '& " Destino Recursos: " & txtBuscaCtaIFDest.Text & " Importe : " & txtImporte.Text
            Set objPista = Nothing
            '*******
    If MsgBox("Desea Realizar otra Operación", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        Limpiar
    Else
        Unload Me
    End If
End If
Set oCon = Nothing
Set oCaja = Nothing
    Exit Sub
TransferirErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    Set oCon = Nothing
    Set oCaja = Nothing
End Sub

Sub Limpiar()
'txtMovDesc = ""
'txtBuscaCtaIFDest = ""
'txtBuscaCtaIFFte = ""
'lblDescCtaIFDest = ""
'lblDescCtaIFFte = ""
'lblDescIFDest = ""
'lblDescIFFte = ""
txtImporte = "0.00"
txtTCBanco = "0.00000"
lsNroDocDest = ""
lsNroDocFte = ""
lsDocumentoDest = ""
lsDocumentoFte = ""
chkDocDest.value = 0
chkDocFuente.value = 0
txtBuscaCtaIFFte.SetFocus
End Sub

Private Sub Form_Activate()
If lbSalir Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
lbSalir = False
Set oOpe = New DOperacion
Set oCtaIf = New NCajaCtaIF
Me.Caption = gsOpeDesc
txtMovDesc = gsOpeDesc
'lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D")
'lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
lsTpoDoc = oOpe.EmiteDocOpe(gsOpeCod, OpeDocEstOpcionalNoDebeExistir, OpeDocMetDigitado)
If gnTipCambio = 0 Then
    Dim oTC As New nTipoCambio
    gnTipCambio = oTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
    If gnTipCambio = 0 And Mid(gsOpeCod, 3, 1) = "2" Then
        MsgBox "Es necesario definir Tipo de Cambio para realizar Operación", vbInformation, "!Aviso!"
        lbSalir = True
        Exit Sub
    End If
End If
lblTCFijo = Format(gnTipCambio, "#,#0.00##")
CentraForm Me
txtFecha = gdFecSis
If Mid(gsOpeCod, 3, 1) = gMonedaExtranjera Then
    txtImporte.BackColor = &HC0FFC0
    lblSimbolo = "US$"
End If
txtBuscaCtaIFFte.psRaiz = "Cuentas de Instituciones Financieras"
txtBuscaCtaIFFte.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")
txtBuscaCtaIFDest.psRaiz = "Cuentas de Instituciones Financieras"
txtBuscaCtaIFDest.rs = oOpe.GetRsOpeObj(gsOpeCod, "2")
cmdAceptar.Visible = False
cmdtransferir.Visible = False
fraCambioBanco.Visible = False
fraTipCambio.Visible = False
Select Case gsOpeCod
    Case gOpeMECompraAInst
        cmdAceptar.Visible = True
        lblSimbolo = gcME
        fraCambioBanco.Visible = True
        fraTipCambio.Visible = True
        'ALPA 20140331
        gnTipCambio = oTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
        lnTipoCambioDia = oTC.EmiteTipoCambio(gdFecSis, 10)
        txtTCBanco.Text = oTC.EmiteTipoCambio(gdFecSis, 10)
        'Compra = "400011"
        '**********************************
    Case gOpeMEVentaAInst
        cmdAceptar.Visible = True
        fraCambioBanco.Visible = True
        fraTipCambio.Visible = True
        lblSimbolo = gcME
        'ALPA 20140331*********************
        lnTipoCambioDia = oTC.EmiteTipoCambio(gdFecSis, 11)
        txtTCBanco.Text = oTC.EmiteTipoCambio(gdFecSis, 11)
        'Compra = "400012"
        '**********************************
    Case gOpeCGTransfBancosMN, gOpeCGTransfBancosCMACSMN, gOpeCGTransfCMACSBancosMN, gOpeCGTransfMismoBancoMN, "401425"
        cmdtransferir.Visible = True
        Select Case gsOpeCod
            Case gOpeCGTransfBancosMN
                FraOrigen.Caption = "Banco Origen"
                FraDestino.Caption = "Banco Destino"
            Case gOpeCGTransfBancosCMACSMN
                FraOrigen.Caption = "Banco Origen"
                FraDestino.Caption = "CMAC Destino"
            Case gOpeCGTransfCMACSBancosMN
                FraOrigen.Caption = "CMAC Origen"
                FraDestino.Caption = "Banco Destino"
            Case gOpeCGTransfMismoBancoMN
                FraOrigen.Caption = "Cuenta Banco Origen"
                FraDestino.Caption = "Cuenta Banco Destino"
            Case "401425"
                FraOrigen.Caption = "Cuenta Banco Origen"
                FraDestino.Caption = "Cuenta Banco Destino"
        End Select
    Case gOpeCGTransfBancosME, gOpeCGTransfBancosCMACSME, gOpeCGTransfCMACSBancosME, gOpeCGTransfMismoBancoME, "402425"
        fraTipCambio.Visible = True
        cmdtransferir.Visible = True
        Select Case gsOpeCod
            Case gOpeCGTransfBancosME
                FraOrigen.Caption = "Banco Origen"
                FraDestino.Caption = "Banco Destino"
            Case gOpeCGTransfBancosCMACSME
                FraOrigen.Caption = "Banco Origen"
                FraDestino.Caption = "CMAC Destino"
            Case gOpeCGTransfCMACSBancosME
                FraOrigen.Caption = "CMAC Origen"
                FraDestino.Caption = "Banco Destino"
            Case gOpeCGTransfMismoBancoME
                FraOrigen.Caption = "Cuenta Banco Origen"
                FraDestino.Caption = "Cuenta Banco Destino"
            Case "402425"
                FraOrigen.Caption = "Cuenta Banco Origen"
                FraDestino.Caption = "Cuenta Banco Destino"
        End Select
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set oOpe = Nothing
End Sub

Private Sub txtBuscaCtaIFDest_EmiteDatos()
Select Case gsOpeCod
    Case gOpeCGTransfMismoBancoMN, gOpeCGTransfMismoBancoME
        If Mid(txtBuscaCtaIFFte, 1, 16) <> Mid(txtBuscaCtaIFDest, 1, 16) Then
            MsgBox "Entidad Origen debe ser igual que la de Destino", vbInformation, "Aviso"
            Exit Sub
        End If
End Select
lblDescCtaIFDest = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaCtaIFDest, 18, 10)) + " " + txtBuscaCtaIFDest.psDescripcion
lblDescIFDest = oCtaIf.NombreIF(Mid(txtBuscaCtaIFDest, 4, 13))
txtMovDesc.SetFocus
End Sub
Private Sub txtBuscaCtaIFFte_EmiteDatos()
lblDescCtaIFFte = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaCtaIFFte, 18, 10)) + " " + txtBuscaCtaIFFte.psDescripcion
lblDescIFFte = oCtaIf.NombreIF(Mid(txtBuscaCtaIFFte, 4, 13))
txtBuscaCtaIFDest.SetFocus
'ALPA 20140606**********************************
    Dim lsMovNroFecha As String
    Dim oGen  As NContFunciones
    Set oGen = New NContFunciones

    lsMovNroFecha = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

'    Dim oCont As COMNContabilidad.NCOMContFunciones
'    Set oCont = New COMNContabilidad.NCOMContFunciones
    Dim oNiv As nTipoCambio
    Dim rs As ADODB.Recordset
    Dim rsMov As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    Set rsMov = New ADODB.Recordset
    
    Dim lnFila As Integer
    Set oNiv = New nTipoCambio
    Set rs = oNiv.ObtenerAprobacionMovCompraVentaPendientexCliente(Mid(lsMovNroFecha, 1, 8), txtBuscaCtaIFFte.Text, gsOpeCod)
    

If Not (rs.EOF Or rs.BOF) Then
    lsMovNroCV = frmNivelesAprobacionCVxPendientesCliente.InicioRegistroNiveles(txtBuscaCtaIFFte.Text, gsOpeCod)
    Set rsMov = oNiv.AprobacionMovCompraVentaxMovimiento(lsMovNroCV)
    If Not (rsMov.EOF Or rsMov.BOF) Then
        lblPermiso = rsMov!cMovNro
        'rsMov!cOpecod
        'rsMov!cNivelCod
        'ChkTCEspecial.value = IIf(rsMov!nTipoEspecial = 0, 0, 1)
        txtImporte.Text = rsMov!nMonto
        txtTCBanco.Text = rsMov!nTipoCambioSolici
        'TxtMontoPagar.Text = Format(Val(txtImporte.Text) * Val(txtTCBanco.Text), "#,#0.00")
        lnTipoCambioDia = rsMov!nTipoCambioNormal
        cmdAceptar.Enabled = True
        cmdAceptar.SetFocus
    End If
End If
'***********************************************
End Sub
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBuscaCtaIFFte.SetFocus
End If
End Sub

Private Sub txtImporte_GotFocus()
fEnfoque txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtImporte, KeyAscii, 20, 2, True)
If KeyAscii = 13 Then
    If cmdAceptar.Visible Then
        cmdAceptar.SetFocus
    ElseIf cmdtransferir.Visible Then
        cmdtransferir.SetFocus
    End If
    'ALPA 20140331*********************************************************
    If (gsOpeCod = gOpeMECompraAInst Or gsOpeCod = gOpeMEVentaAInst) Then
        Dim oDN As dTipoCambio
        Set oDN = New dTipoCambio
        If Len(Trim(lsMovNroCV)) > 0 Then
            oDN.ActualizarAprobacionMovCompraVenta (lsMovNroCV)
            txtTCBanco.Text = lnTipoCambioDia
            lsMovNroCV = ""
            lblPermiso.Caption = ""
        End If
    End If
    '**********************************************************************
End If
End Sub

Private Sub txtImporte_LostFocus()
If Len(Trim(txtImporte)) = 0 Then txtImporte = 0
txtImporte = Format(txtImporte, "#,#0.00")
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If fraCambioBanco.Visible And fraCambioBanco.Enabled Then
        txtTCBanco.SetFocus
    Else
        txtImporte.SetFocus
    End If
    
End If
End Sub

Private Sub txtTCBanco_GotFocus()
fEnfoque txtTCBanco
End Sub
Private Sub txtTCBanco_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTCBanco, KeyAscii, 8, 5, True)
If KeyAscii = 13 Then
    Call ValidaInsertaAprobacionCV 'ALPA 20140331******************************************************
    txtImporte.SetFocus
End If
End Sub
Private Sub txtTCBanco_LostFocus()
txtTCBanco = Format(txtTCBanco, "#,#0.00000")
End Sub
'ALPA 20140331************************************************************************
Private Sub ValidaInsertaAprobacionCV()
    If (gsOpeCod = gOpeMECompraAInst Or gsOpeCod = gOpeMEVentaAInst) Then
        Dim oDN As dTipoCambio
        Set oDN = New dTipoCambio
        Dim lnActivarCambioenTipodeCambio As Integer
        Dim lnNivel As String
        Dim oGen  As NContFunciones
        Set oGen = New NContFunciones
        Dim oCred  As nTipoCambio
        Set oCred = New nTipoCambio
        
        If Len(Trim(lsMovNroCV)) > 0 Then
            oDN.ActualizarAprobacionMovCompraVenta (lsMovNroCV)
        End If
        Set oDN = Nothing
        lsMovNroCV = ""
        lblPermiso.Caption = ""
        If CDbl(txtTCBanco.Text) <> CDbl(lnTipoCambioDia) Then
            Set oDN = New dTipoCambio
            lnActivarCambioenTipodeCambio = oCred.ObtenerNivelesAprobacionCompraVentaxMonto(txtImporte.Text, CDbl(txtTCBanco.Text) - CDbl(lnTipoCambioDia), IIf(gsOpeCod = gOpeMECompraAInst, 1, 2), lnNivel, 0)
            If lnActivarCambioenTipodeCambio = 1 Then
                MsgBox "Se generará movimiento de tipo de cambio...esperar la aprobación al nivel respectivo", vbCritical
            Else
                MsgBox "El cambio realizado no tiene nivel de aprobacion, favor registrar otro tipo de cambio", vbCritical
                txtTCBanco.Text = Format(lnTipoCambioDia, "#,#0.0000")
                Exit Sub
            End If
            lsMovNroCV = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            lblPermiso.Caption = "Mov: " & lsMovNroCV
            Call oDN.AprobacionMovCompraVenta(lsMovNroCV, gsOpeCod, lnNivel, CDbl(txtImporte.Text), CDbl(lnTipoCambioDia), CDbl(txtTCBanco.Text), 0, txtBuscaCtaIFFte.Text)
'        Else
'            TxtMontoPagar.Text = Format(Val(txtImporte.value) * Val(txtTCBanco.Text), "#,#0.00")
        End If
    End If
End Sub
Private Sub ActualizarNivelAprovacion()
'ALPA 20140328**********************************************************
Dim oDN  As dTipoCambio
Set oDN = New dTipoCambio
If Len(Trim(lsMovNroCV)) > 0 Then
    oDN.ActualizarAprobacionMovCompraVenta (lsMovNroCV)
End If
lsMovNroCV = "" 'ALPA 20140328**************
lblPermiso.Caption = "" 'ALPA 20140328**************
'***********************************************************************
End Sub
'*************************************************************************************
