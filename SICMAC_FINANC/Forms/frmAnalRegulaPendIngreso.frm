VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAnalRegulaPendIngreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos de Rendicion a Caja General"
   ClientHeight    =   3495
   ClientLeft      =   1020
   ClientTop       =   3075
   ClientWidth     =   7860
   Icon            =   "frmAnalRegulaPendIngreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkDif 
      Caption         =   "Ajustar Diferencia"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   150
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pendiente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   780
      Left            =   150
      TabIndex        =   14
      Top             =   150
      Width           =   7545
      Begin VB.Label txtNomPers 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   150
         TabIndex        =   17
         Top             =   300
         Width           =   4845
      End
      Begin VB.Label txtSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   315
         Left            =   5790
         TabIndex        =   16
         Top             =   315
         Width           =   1605
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldo"
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   4650
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   660
      Left            =   150
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2250
      Width           =   4410
   End
   Begin VB.Frame Frame4 
      Caption         =   "Importe "
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
      Height          =   780
      Left            =   4605
      TabIndex        =   11
      Top             =   2160
      Width           =   3060
      Begin VB.Label txtImporte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   315
         Left            =   960
         TabIndex        =   18
         Top             =   270
         Width           =   1905
      End
      Begin VB.Label lblSimbolo 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   195
         TabIndex        =   12
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5220
      TabIndex        =   7
      Top             =   3030
      Width           =   1125
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   6390
      TabIndex        =   8
      Top             =   3030
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Forma de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1065
      Left            =   150
      TabIndex        =   10
      Top             =   1020
      Width           =   7545
      Begin VB.OptionButton optFormPago 
         Caption         =   "&Ingreso en Ventanilla"
         CausesValidation=   0   'False
         Height          =   195
         Index           =   4
         Left            =   330
         TabIndex        =   5
         Top             =   660
         Width           =   1920
      End
      Begin VB.OptionButton optFormPago 
         Caption         =   "O&tros"
         CausesValidation=   0   'False
         Height          =   195
         Index           =   6
         Left            =   4965
         TabIndex        =   4
         Top             =   660
         Width           =   900
      End
      Begin VB.OptionButton optFormPago 
         Caption         =   "&Efectivo"
         CausesValidation=   0   'False
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   3
         Top             =   330
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optFormPago 
         Caption         =   "&Orden de pago "
         CausesValidation=   0   'False
         Height          =   195
         Index           =   3
         Left            =   4965
         TabIndex        =   1
         Top             =   330
         Width           =   1380
      End
      Begin VB.OptionButton optFormPago 
         Caption         =   "&Cargo a Cuenta"
         CausesValidation=   0   'False
         Height          =   195
         Index           =   5
         Left            =   2640
         TabIndex        =   2
         Top             =   660
         Width           =   1500
      End
      Begin VB.OptionButton optFormPago 
         Caption         =   "C&heque"
         CausesValidation=   0   'False
         Height          =   195
         Index           =   2
         Left            =   2640
         TabIndex        =   0
         Top             =   330
         Width           =   1230
      End
   End
   Begin RichTextLib.RichTextBox rtxt 
      Height          =   315
      Left            =   1665
      TabIndex        =   13
      Top             =   6555
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmAnalRegulaPendIngreso.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAnalRegulaPendIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OK As Boolean
Dim sMovNroRef As String    'Movimiento de Referencia
Dim sSql As String
Dim rs As New ADODB.Recordset
Dim lMN As Boolean
Dim lSalir As Boolean
Dim lTransActiva As Boolean
Dim lViaticos    As Boolean
Dim sCtaPendiente As String
Dim lsObjetos()  As String
Dim lbActCheque  As Boolean, lbActOrdenP As Boolean, lbActCargoC As Boolean, lbActEfectivo As Boolean, lbActOtros As Boolean
Dim lbActPagVent As Boolean
Dim lsPersCod    As String
Dim lnSaldo     As Currency
Dim MenuItem As Integer
Dim oArendir As NARendir
Dim oOpe As DOperacion

Public Sub Inicio(pbActCheque As Boolean, pbActOrdenP As Boolean, pbActCargoC As Boolean, pbActEfectivo As Boolean, pbActOtros As Boolean, pbActPagVent As Boolean, Optional psPersCod As String = "", Optional pnSaldo As Currency = 0)
lbActCheque = pbActCheque
lbActOrdenP = pbActOrdenP
lbActCargoC = pbActCargoC
lbActEfectivo = pbActEfectivo
lbActOtros = pbActOtros
lbActPagVent = pbActPagVent
lsPersCod = psPersCod
lnSaldo = pnSaldo
Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
Dim rs As ADODB.Recordset
Dim oContFunc As NContFunciones
Dim oDocPago As clsDocPago
Dim oDocRec As NDocRec
Dim oContImp As NContImprimir
                

Dim lsMovNro As String
Dim lsOpeCod As String
Dim lbMueveCtasCont As Boolean
Dim lsCtaPendiente As String
Dim lsCtaOperacion As String
Dim lsCtaDiferencia As String
Dim lnImporte  As Currency
Dim lnMontoDif As Currency
Dim lsDocVoucher As String
Dim lsDocumento As String
Dim lbEfectivo As Boolean
Dim lbIngreso As Boolean

Dim lsPersCodIf As String
Dim lsTipoIF As String
Dim lsCtaBanco As String

Dim lsCtaChqIf As String
Dim lnPlaza As Integer
Dim ldValchq As Date
Dim lsDocNroVoucher As String
Dim ldFechaVoucher As Date
Dim lsDocNRo As String
Dim lsEntidadOrig As String
Dim lsCtaEntidadOrig As String
Dim lsGlosa As String
Dim lsPersNombre As String
Dim lsSubCuentaIF As String
Dim lsPersCod As String
Dim lsFechaDoc As String
Dim lsCuentaAho As String
Dim lnPersoneria As PersPersoneria

Dim lsPersDireccion As String
Dim lsUbigeo As String
Dim lnMotivo As MotivoNotaAbonoCargo
Dim lsCadBol As String
Dim lsAgeCodRef As String
On Error GoTo ErrGraba

lsPersNombre = Me.txtNomPers
lsPersCod = txtNomPers.Tag
lnImporte = CCur(txtSaldo)
lsDocNRo = ""
lsDocVoucher = ""

lsCtaBanco = ""
lsPersCodIf = ""
lsTipoIF = ""
gsGlosa = txtMovDesc

Set oDocPago = New clsDocPago
Set oOpe = New DOperacion
Set oArendir = New NARendir

lbMueveCtasCont = True
If ValidaDatos = False Then Exit Sub
Set oContFunc = New NContFunciones
cmdAceptar.Enabled = False
gnDocTpo = -1
lbIngreso = True
lsCtaDiferencia = ""
lsCuentaAho = ""
lnMontoDif = 0
Set oDocRec = New NDocRec
Set oContImp = New NContImprimir

Select Case MenuItem
    Case 1    'Efectivo
        frmArendirEfectivo.Inicio -1, gsDocNro, Mid(gsOpeCod, 3, 1), "", txtSaldo, lsPersCod, lsPersNombre, ArendirRendicion
        If Not frmArendirEfectivo.lbOk Then
            Exit Sub
        End If
        Set rs = frmArendirEfectivo.rsEfectivo
        If frmArendirEfectivo.vnDiferencia <> 0 Then
            lnMontoDif = frmArendirEfectivo.vnDiferencia
        End If
    Case 2    'Cheque
        'Registro de cheque
        'EJVG20140415 ***
        'Set frmIngCheques = Nothing
        'gnDocTpo = TpoDocCheque
        'lsOpeCod = oArendir.GetOpeRendicion(Mid(gsOpeCod, 1, 5), gnDocTpo, sCtaPendiente, sCtaPendiente, lbMueveCtasCont, "D")
        'lsCtaPendiente = frmAnalisisRegulaPend.txtCtaPend
        'frmIngCheques.InicioRendirPendiente lsOpeCod, lnImporte, lsCtaPendiente, Trim(txtMovDesc), Mid(gsOpeCod, 3, 1), frmAnalisisRegulaPend.lvPend.GetRsNew()
        'If frmIngCheques.OK = False Then
        '    Exit Sub
        'Else
        '    OK = True
        '    Unload Me
        '    Exit Sub
        'End If
        Exit Sub
        'END EJVG *******
    Case 3    'Orden Pago
        gnDocTpo = TpoDocOrdenPago
        frmNotaCargoAbono.Inicio TpoDocOrdenPago, CCur(txtSaldo), gdFecSis, txtMovDesc, gsOpeCod
        If frmNotaCargoAbono.vbOk Then
            lsDocNRo = frmNotaCargoAbono.NroNotaCA
            lsFechaDoc = frmNotaCargoAbono.FechaNotaCA
            txtMovDesc = frmNotaCargoAbono.Glosa
            lsDocumento = frmNotaCargoAbono.NotaCargoAbono
            lsPersNombre = frmNotaCargoAbono.PersNombre
            lsPersDireccion = frmNotaCargoAbono.PersDireccion
            lsUbigeo = frmNotaCargoAbono.PersUbigeo
            lnPersoneria = frmNotaCargoAbono.Personeria
            ldFechaVoucher = frmNotaCargoAbono.FechaNotaCA
            lsCuentaAho = frmNotaCargoAbono.CuentaAhoNro
        Else
            Exit Sub
        End If
    
    Case 4    'Ingresos por Ventanilla
        gnDocTpo = TpoDocRecibosDiversos
        frmOpeNegVentanilla.Inicio gsDocNro, Mid(gsOpeCod, 3, 1), txtSaldo, lsPersCod, lsPersNombre
        If Not frmOpeNegVentanilla.lbOk Then
            Exit Sub
        End If
        Set rs = frmOpeNegVentanilla.rsPago
        If frmOpeNegVentanilla.vnDiferencia <> 0 Then
            lnMontoDif = frmOpeNegVentanilla.vnDiferencia
        End If
        
    Case 5    'Nota de Cargo
        gnDocTpo = TpoDocNotaCargo
        frmNotaCargoAbono.Inicio TpoDocNotaCargo, CCur(txtSaldo), gdFecSis, txtMovDesc, gsOpeCod
        If frmNotaCargoAbono.vbOk Then
            lsDocNRo = frmNotaCargoAbono.NroNotaCA
            lsFechaDoc = frmNotaCargoAbono.FechaNotaCA
            txtMovDesc = frmNotaCargoAbono.Glosa
            lsDocumento = frmNotaCargoAbono.NotaCargoAbono
            'lsDocNroVoucher = oContFunc.GeneraDocNro(TpoDocVoucherEgreso, Mid(gsOpeCod, 3, 1))
            lsPersNombre = frmNotaCargoAbono.PersNombre
            lsPersDireccion = frmNotaCargoAbono.PersDireccion
            lsUbigeo = frmNotaCargoAbono.PersUbigeo
            ldFechaVoucher = frmNotaCargoAbono.FechaNotaCA
            lsCuentaAho = frmNotaCargoAbono.CuentaAhoNro
        Else
            Exit Sub
        End If
    Case 6    'Otros Ingresos
       frmAsientoRegistro.Inicio "", 0, , True, True, False, True, "", frmAnalisisRegulaPend.lvPend.GetRsNew
       If frmAsientoRegistro.lOk Then
            OK = True
            Unload Me
       End If
       Exit Sub
End Select

lsOpeCod = oArendir.GetOpeRendicion(Mid(gsOpeCod, 1, 5), gnDocTpo, sCtaPendiente, sCtaPendiente, lbMueveCtasCont, IIf(lbIngreso = True, "D", "H"))
If MsgBox("Desea Grabar la Rendicion respectiva?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oContFunc.GeneraMovNro(frmAnalisisRegulaPend.txtFecRegula, gsCodAge, gsCodUser)
    lbEfectivo = False
    lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "D")
    lsCtaPendiente = frmAnalisisRegulaPend.txtCtaPend
    Select Case MenuItem
        Case 1    'Efectivo
            lsCtaDiferencia = oOpe.EmiteOpeCta(lsOpeCod, IIf(lnMontoDif > 0, "H", "D"), "2")
            oArendir.GrabaRendicionEfectivo -1, gsFormatoFecha, lsMovNro, lsOpeCod, txtMovDesc, _
                                        lsCtaPendiente, lsCtaOperacion, lnImporte, rs, "", 0, lbIngreso, lsCtaDiferencia, lnMontoDif, frmAnalisisRegulaPend.lvPend.GetRsNew
            lbEfectivo = True
        Case 2    'Cheque
            'Se realiza la grabación dentro del formulario de resgistro de cheques
        Case 3    'Orden Pago
            lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "D", Trim(Str(lnPersoneria)), gsCodAge, ObjCMACAgencias)
            oArendir.CapCargoCuentaAhoMov gsFormatoFecha, lsCuentaAho, lnImporte, lsOpeCod, lsMovNro, txtMovDesc, TpoDocOrdenPago, _
                    lsDocNRo, , True, , , , True, , lsCtaPendiente, lsCtaOperacion, "", _
                    "", gdFecSis, True
                    
        Case 4    'Ingreso por Ventanilla
            oArendir.GrabaRendicionVentanilla -1, gsFormatoFecha, lsMovNro, lsOpeCod, txtMovDesc, _
                                        lsCtaPendiente, lsCtaOperacion, lnImporte, rs, "", 0, lbIngreso, lsCtaDiferencia, lnMontoDif, frmAnalisisRegulaPend.txtAgeCod, frmAnalisisRegulaPend.lvPend.GetRsNew, gbBitCentral
            lbEfectivo = False
            
        Case 5    'Nota de Cargo
            lnMotivo = gNCRendirCuenta
            oArendir.GrabaRendicionGiroDocumento -1, lsMovNro, "", _
                    "", lsOpeCod, txtMovDesc, lsCtaPendiente, _
                    lsCtaOperacion, lsPersCod, lnImporte, gnDocTpo, lsDocNRo, lsFechaDoc, lsDocNroVoucher, lsPersCodIf, lsTipoIF, _
                    lsCtaBanco, , , , lnMotivo, lsCuentaAho, gbBitCentral, True
            
            Dim oDis As New NRHProcesosCierre
            lsCadBol = oDis.ImprimeBoletaCad(ldFechaVoucher, "CARGO CAJA GENERAL", "Retiro CAJA GENERAL*Nro." & lsDocNRo, "", lnImporte, lsPersNombre, lsCuentaAho, "", 0, 0, "Nota Cargo", 0, 0, False, False, , , , True, , , , False, gsNomAge) & oImpresora.gPrnSaltoPagina
            
        Case 6    'Otros Ingresos
            lsOpeCod = gsOpeCod
            oArendir.GrabaRendicionGiroDocumento -1, gsFormatoFecha, lsMovNro, "", _
                    "", lsOpeCod, txtMovDesc, lsCtaOperacion, _
                    lsCtaPendiente, lsPersCod, lnImporte, gnDocTpo, "", "", "", "", "", ""
    End Select
    
    ImprimeAsientoContable lsMovNro, lsDocVoucher, gnDocTpo, lsDocumento, lbEfectivo, lbIngreso, txtMovDesc, lsPersCod, lnImporte, , , , , , "17", , , lsCadBol
    
    Set frmArendirEfectivo = Nothing
    OK = True
    Set oDocRec = Nothing
    Set oContImp = Nothing
    cmdAceptar.Enabled = True
    Unload Me
End If
Exit Sub
ErrGraba:
  MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"""
End Sub


Private Sub cmdSalir_Click()
OK = False
Unload Me
End Sub

Private Sub Form_Load()
Dim N As Integer
CentraForm Me
If Mid(gsOpeCod, 3, 1) = "1" Then
   lMN = True
   gsSimbolo = gcMN
Else
   lMN = False
   gsSimbolo = gcME
End If
lblSimbolo = gsSimbolo
lTransActiva = False
lSalir = False
txtImporte = Format(gnImporte, gsFormatoNumeroView)
txtSaldo = Format(lnSaldo, gsFormatoNumeroView)

Me.txtMovDesc = gsGlosa
sMovNroRef = gnMovNro
sCtaPendiente = frmAnalisisRegulaPend.sPendiente
txtNomPers = gsPersNombre
txtNomPers.Tag = lsPersCod

gnDocTpo = -1
MenuItem = 1

optFormPago(1).Enabled = lbActEfectivo
optFormPago(2).Enabled = lbActCheque
optFormPago(3).Enabled = lbActOrdenP
optFormPago(4).Enabled = lbActPagVent
optFormPago(5).Enabled = lbActCargoC
optFormPago(6).Enabled = lbActOtros
End Sub

Private Sub optFormPago_Click(Index As Integer)
MenuItem = Index
    Select Case MenuItem
        Case 1    'Efectivo
            cmdAceptar_Click
        Case 2    'Cheque
            cmdAceptar_Click
        Case 3    'Orden Pago
            cmdAceptar_Click
        Case 4    'Ingreso por Ventanilla
            cmdAceptar_Click
        Case 5    'Nota de Cargo
            cmdAceptar_Click
        Case 6
            cmdAceptar_Click
    End Select
End Sub

Public Property Get lOk() As Boolean
lOk = OK
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
OK = vNewValue
End Property

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtImporte, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   KeyAscii = 0
   txtImporte = Format(Format(txtImporte, gsFormatoNumeroDato), gsFormatoNumeroView)
   If Val(Format(txtImporte, gsFormatoNumeroDato)) > 0 Then
      cmdAceptar.Enabled = True
      cmdAceptar.SetFocus
   Else
      txtImporte = ""
      cmdAceptar.Enabled = False
   End If
End If
End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   cmdAceptar.Enabled = True
   cmdAceptar.SetFocus
ElseIf Len(txtMovDesc) = 0 Then
   cmdAceptar.Enabled = False
End If
End Sub

Function ValidaDatos() As Boolean
ValidaDatos = True
cmdAceptar.Enabled = False
Select Case MenuItem
    Case 1    'Efectivo
    Case 2    'Cheque
    Case 3    'Orden Pago
    Case 4    'Nota de Cargo
End Select
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Descripción o glosa de Operacion no Ingresada ", vbInformation, "Aviso"
    ValidaDatos = False
    txtMovDesc.SetFocus
    Exit Function
End If
cmdAceptar.Enabled = True
End Function

