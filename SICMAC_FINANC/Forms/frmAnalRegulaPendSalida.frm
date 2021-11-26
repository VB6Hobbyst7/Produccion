VERSION 5.00
Begin VB.Form frmAnalRegulaPendSalida 
   Caption         =   "Documentos de Pago de Caja General"
   ClientHeight    =   4590
   ClientLeft      =   4305
   ClientTop       =   2775
   ClientWidth     =   7440
   Icon            =   "frmAnalRegulaPendSalida.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7440
   Begin VB.CheckBox chkDif 
      Caption         =   "Ajustar Diferencia"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   4290
      Width           =   1695
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   600
      Left            =   90
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3420
      Width           =   4290
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
      Left            =   90
      TabIndex        =   16
      Top             =   30
      Width           =   7275
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
         TabIndex        =   20
         Top             =   285
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
         Left            =   5700
         TabIndex        =   19
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldo"
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   4845
         TabIndex        =   17
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame FraCtaIFPagadora 
      Caption         =   "Entidad Pagadora"
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
      Height          =   1560
      Left            =   105
      TabIndex        =   15
      Top             =   1770
      Visible         =   0   'False
      Width           =   7260
      Begin Sicmact.TxtBuscar txtBuscaEntidad 
         Height          =   375
         Left            =   135
         TabIndex        =   7
         Top             =   285
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   661
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCtaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   1095
         Width           =   6975
      End
      Begin VB.Label lblIFNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         TabIndex        =   8
         Top             =   720
         Width           =   6975
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   6030
      TabIndex        =   13
      Top             =   4170
      Width           =   1125
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
      Height          =   705
      Left            =   4500
      TabIndex        =   12
      Top             =   3330
      Width           =   2865
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
         Left            =   750
         TabIndex        =   21
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
         Left            =   180
         TabIndex        =   14
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4860
      TabIndex        =   11
      Top             =   4170
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
      ForeColor       =   &H8000000D&
      Height          =   915
      Left            =   90
      TabIndex        =   0
      Top             =   810
      Width           =   7275
      Begin VB.OptionButton optFormPago 
         Caption         =   "O&tros"
         CausesValidation=   0   'False
         Height          =   195
         Index           =   6
         Left            =   2100
         TabIndex        =   6
         Top             =   570
         Width           =   900
      End
      Begin VB.OptionButton optFormPago 
         Caption         =   "A&bono en Cuenta"
         CausesValidation=   0   'False
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   5
         Top             =   540
         Width           =   1545
      End
      Begin VB.OptionButton optFormPago 
         Caption         =   "E&fectivo"
         CausesValidation=   0   'False
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optFormPago 
         Caption         =   "Car&ta"
         CausesValidation=   0   'False
         Height          =   195
         Index           =   4
         Left            =   5670
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
      Begin VB.OptionButton optFormPago 
         Caption         =   "&Orden de Pago"
         CausesValidation=   0   'False
         Height          =   195
         Index           =   3
         Left            =   3750
         TabIndex        =   3
         Top             =   270
         Width           =   1455
      End
      Begin VB.OptionButton optFormPago 
         Caption         =   "C&heque"
         CausesValidation=   0   'False
         Height          =   195
         Index           =   2
         Left            =   2100
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAnalRegulaPendSalida"
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
Dim sCtaPendiente As String
Dim lsPersCod    As String
Dim MenuItem As Integer
Dim lbActCheque As Boolean, lbActCarta As Boolean, lbActOrdenP As Boolean, lbActAbonoC As Boolean, lbActEfectivo As Boolean, lbActOtros As Boolean
Dim lnSaldo     As Currency
Dim oArendir As NARendir
Dim oCtasIF As NCajaCtaIF
Dim oOpe As DOperacion

Public Sub Inicio(pbActCheque As Boolean, pbActCarta As Boolean, pbActOrdenP As Boolean, pbActAbonoC As Boolean, pbActEfectivo As Boolean, pbActOtros As Boolean, Optional psPersCod As String = "", Optional pnSaldo As Currency = 0)
lbActCheque = pbActCheque
lbActCarta = pbActCarta
lbActOrdenP = pbActOrdenP
lbActAbonoC = pbActAbonoC
lbActEfectivo = pbActEfectivo
lbActOtros = pbActOtros
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
Dim lsFechaDoc As Date
Dim lsDocViaticos As String
Dim lsCuentaAho As String
Dim lnPersoneria As PersPersoneria

Dim lsPersDireccion As String
Dim lsUbigeo As String
Dim lnMotivo As MotivoNotaAbonoCargo
Dim lsCadBol As String

On Error GoTo ErrGraba
If ValidaDatos = False Then Exit Sub

lsEntidadOrig = lblIFNombre
lsCtaEntidadOrig = Trim(lblCtaDesc)
lsPersNombre = PstaNombre(txtNomPers, True)
Set oDocPago = New clsDocPago
Set oOpe = New DOperacion
Set oCtasIF = New NCajaCtaIF
Set oContFunc = New NContFunciones

lsSubCuentaIF = oCtasIF.SubCuentaIF(Mid(txtBuscaEntidad.Text, 1, 13))
Set oCtasIF = Nothing

lnImporte = CCur(txtImporte)
lsDocNRo = ""
lsDocVoucher = ""

lsCtaBanco = ""
lsPersCodIf = ""
lsTipoIF = ""

lbMueveCtasCont = True
cmdAceptar.Enabled = False
gnDocTpo = -1
lbIngreso = False
lsCtaDiferencia = ""
lsCuentaAho = ""
lnMontoDif = 0
Set oDocRec = New NDocRec
Set oContImp = New NContImprimir

Select Case MenuItem
    Case 1    'Efectivo
        frmArendirEfectivo.Inicio -1, "", Mid(gsOpeCod, 3, 1), "", txtImporte, lsPersCod, lsPersNombre, IIf(Me.chkDif.value = vbChecked, ArendirRendicion, -1), "Regulazarización en Efectivo"
        If Not frmArendirEfectivo.lbOk Then
            Exit Sub
        End If
        Set rs = frmArendirEfectivo.rsEfectivo
        If frmArendirEfectivo.vnDiferencia <> 0 Then
            lnMontoDif = frmArendirEfectivo.vnDiferencia
        End If
    
    Case 2    'Cheque
        Screen.MousePointer = 11
        gnDocTpo = TpoDocCheque
        lsCtaBanco = Mid(txtBuscaEntidad, 18, Len(txtBuscaEntidad))
        lsPersCodIf = Mid(txtBuscaEntidad, 4, 13)
        lsTipoIF = Mid(txtBuscaEntidad, 1, 2)
        'oDocPago.InicioCheque lsDocNRo, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeDesc, txtMovDesc, lnImporte, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocNroVoucher, True ', , lsCtaBanco
        oDocPago.InicioCheque lsDocNRo, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeDesc, txtMovDesc, lnImporte, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocNroVoucher, True, , , , lsTipoIF, lsPersCodIf, lsCtaBanco 'EJVG20121130
        Screen.MousePointer = 0
        If oDocPago.vbOk Then    'Se ingresó dato de Cheque u Orden de Pago
            lsFechaDoc = oDocPago.vdFechaDoc
            lsDocNRo = oDocPago.vsNroDoc
            lsDocNroVoucher = oDocPago.vsNroVoucher
            ldFechaVoucher = oDocPago.vdFechaDoc
            lsDocumento = oDocPago.vsFormaDoc
            txtMovDesc = oDocPago.vsGlosa
        Else
            Exit Sub
        End If
    Case 3    'Orden Pago
        Screen.MousePointer = 11
        gnDocTpo = TpoDocOrdenPago
        oDocPago.InicioOrdenPago lsDocNRo, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeCod, txtMovDesc, lnImporte, gdFecSis, lsDocNroVoucher, False ',  gsCodAge
        Screen.MousePointer = 0
        If oDocPago.vbOk Then    'Se ingresó dato de Cheque u Orden de Pago
            lsFechaDoc = oDocPago.vdFechaDoc
            lsDocNRo = oDocPago.vsNroDoc
            lsDocNroVoucher = oDocPago.vsNroVoucher
            ldFechaVoucher = oDocPago.vdFechaDoc
            lsDocumento = oDocPago.vsFormaDoc
            txtMovDesc = oDocPago.vsGlosa
        Else
            Exit Sub
        End If
    Case 4    'Carta
        gnDocTpo = TpoDocCarta
        lsCtaBanco = Mid(txtBuscaEntidad, 18, Len(txtBuscaEntidad))
        lsPersCodIf = Mid(txtBuscaEntidad, 4, 13)
        lsTipoIF = Mid(txtBuscaEntidad, 1, 2)
        oDocPago.InicioCarta lsDocNRo, lsPersCod, gsOpeCod, gsOpeCod, txtMovDesc, "", lnImporte, gdFecSis, lsEntidadOrig, lsCtaEntidadOrig, lsPersNombre, "", ""
        If oDocPago.vbOk Then    'Se ingresó datos de carta
            lsFechaDoc = oDocPago.vdFechaDoc
            lsDocNRo = oDocPago.vsNroDoc
            lsDocNroVoucher = oDocPago.vsNroVoucher
            ldFechaVoucher = oDocPago.vdFechaDoc
            lsDocumento = oDocPago.vsFormaDoc
            txtMovDesc = oDocPago.vsGlosa
        Else
            Exit Sub
        End If
    Case 5    'Nota de Abono
        Dim oImp As New NContImprimir
        gnDocTpo = TpoDocNotaAbono
        frmNotaCargoAbono.Inicio TpoDocNotaAbono, CCur(txtSaldo), gdFecSis, txtMovDesc, gsOpeCod
        If frmNotaCargoAbono.vbOk Then
            lsDocNRo = frmNotaCargoAbono.NroNotaCA
            lsFechaDoc = frmNotaCargoAbono.FechaNotaCA
            txtMovDesc = frmNotaCargoAbono.Glosa
            lsDocumento = frmNotaCargoAbono.NotaCargoAbono
            lsPersNombre = frmNotaCargoAbono.lblPersNombre
            lsPersDireccion = frmNotaCargoAbono.lblPersDireccion
            lsUbigeo = frmNotaCargoAbono.lblUbigeo
            ldFechaVoucher = frmNotaCargoAbono.FechaNotaCA
            lsCuentaAho = frmNotaCargoAbono.CuentaAhoNro
            
            lsDocumento = oImp.ImprimeNotaAbono(Format(ldFechaVoucher, gsFormatoFecha), lnImporte, txtMovDesc, lsCuentaAho, lsPersNombre)
            Dim oDis As New NRHProcesosCierre
            lsCadBol = oDis.ImprimeBoletaCad(ldFechaVoucher, "ABONO CAJA GENERAL", "Depósito CAJA GENERAL*Nro." & lsDocNRo, "", lnImporte, lsPersNombre, lsCuentaAho, "", 0, 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False, gsNomAge) & oImpresora.gPrnSaltoPagina
            
            Unload frmNotaCargoAbono
            Set frmNotaCargoAbono = Nothing
        Else
            Unload frmNotaCargoAbono
            Set frmNotaCargoAbono = Nothing
            Exit Sub
        End If
    Case 6   'Otros Egresos
        'validacion dentro del menu
End Select
Set oArendir = New NARendir
lsOpeCod = GetOpeRegulaPendienteSalida(gsOpeCod, gnDocTpo)
If MsgBox("¿ Desea Grabar la Rendicion Respectiva ? ", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    lbEfectivo = False
    lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "H")
    lsCtaPendiente = frmAnalisisRegulaPend.txtCtaPend
        Select Case MenuItem
            Case 1    'Efectivo
                lsCtaDiferencia = oOpe.EmiteOpeCta(gsOpeCod, IIf(lnMontoDif > 0, "H", "D"), "2")
                oArendir.GrabaRendicionEfectivo -1, gsFormatoFecha, lsMovNro, lsOpeCod, txtMovDesc, _
                                            lsCtaPendiente, lsCtaOperacion, lnImporte, rs, "", 0, lbIngreso, lsCtaDiferencia, lnMontoDif, frmAnalisisRegulaPend.lvPend.GetRsNew
                lbEfectivo = True
            Case 2    'Cheque
                lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "H", , txtBuscaEntidad, CtaOBjFiltroIF, True)
                oArendir.GrabaRendicionGiroDocumento -1, lsMovNro, 0, _
                    0, lsOpeCod, txtMovDesc, lsCtaPendiente, _
                    lsCtaOperacion, lsPersCod, lnImporte, gnDocTpo, lsDocNRo, Format(lsFechaDoc, gsFormatoFecha), lsDocNroVoucher, lsPersCodIf, lsTipoIF, _
                    lsCtaBanco
            
            Case 3     'Orden Pago
                ''***Modificado por ELRO el 20120723, según OYP-RFC005-2012 y OYP-RFC016-2012
                'lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "H", , frmARendirLista.TxtBuscarArendir, ObjCMACAgenciaArea)
                lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "H", , frmARendirLista2.TxtBuscarArendir, ObjCMACAgenciaArea)
                lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "H", , txtBuscaEntidad, CtaOBjFiltroIF, True)
                oArendir.GrabaRendicionGiroDocumento -1, lsMovNro, 0, _
                    0, lsOpeCod, txtMovDesc, lsCtaPendiente, _
                    lsCtaOperacion, lsPersCod, lnImporte, gnDocTpo, lsDocNRo, Format(lsFechaDoc, gsFormatoFecha), lsDocNroVoucher, lsPersCodIf, lsTipoIF, _
                    lsCtaBanco
            
            Case 4      'Carta
                lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "H", "0", txtBuscaEntidad, CtaOBjFiltroIF, True)
                oArendir.GrabaRendicionGiroDocumento -1, lsMovNro, "", _
                    "", lsOpeCod, txtMovDesc, lsCtaPendiente, _
                    lsCtaOperacion, lsPersCod, lnImporte, gnDocTpo, lsDocNRo, Format(lsFechaDoc, gsFormatoFecha), lsDocNroVoucher, lsPersCodIf, lsTipoIF, _
                    lsCtaBanco
                    
            Case 5      'Nota de Abono
                lnMotivo = gNARendirCuenta
                oArendir.GrabaRendicionGiroDocumento -1, lsMovNro, "", _
                    "", lsOpeCod, txtMovDesc, lsCtaPendiente, _
                    lsCtaOperacion, lsPersCod, lnImporte, gnDocTpo, lsDocNRo, Format(lsFechaDoc, gsFormatoFecha), lsDocNroVoucher, lsPersCodIf, lsTipoIF, _
                    lsCtaBanco, , , , lnMotivo, lsCuentaAho, gbBitCentral, True
                
            Case 6      'Otros Egresos
                oArendir.GrabaRendicionGiroDocumento -1, lsMovNro, "", _
                        "", lsOpeCod, txtMovDesc, lsCtaPendiente, _
                        lsCtaOperacion, lsPersCod, lnImporte, gnDocTpo, "", "", "", "", "", ""
                
        End Select
    ImprimeAsientoContable lsMovNro, lsDocVoucher, gnDocTpo, lsDocumento, lbEfectivo, lbIngreso, txtMovDesc, lsPersCod, lnImporte, , lsDocViaticos, , , , "17", , , lsCadBol
    Set frmArendirEfectivo = Nothing
    OK = True
    Set oDocRec = Nothing
    Set oContImp = Nothing
    cmdAceptar.Enabled = True
    Unload Me
End If

Exit Sub
ErrGraba:
  MsgBox TextErr(Err.Description), vbCritical, "Error de Actualización"
End Sub

Public Function GetOpeRegulaPendienteSalida(ByVal psOpeCod As String, ByVal psDocTpo As TpoDoc) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsFiltroCta As String
Dim oConect As DConecta
Dim lsFiltroDH As String

Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sql = " Select  O.COPECOD, OD.nDocTpo " _
    & " From    OpeTpo O " _
    & "         Left Join OpeDoc OD on OD.cOpeCod = O.cOpeCod " _
    & " Where   Substring(O.cOpeCod,1,5) = '" & Left(psOpeCod, 5) & "' " _
    & "         and O.cOpeCod > '" & psOpeCod & "' " _
    & "         and OD.nDocTpo " & IIf(psDocTpo = -1, " IS NULL ", " ='" & psDocTpo & "'") _
    & " GROUP BY O.COPECOD, OD.nDocTpo "

Set rs = oConect.CargaRecordSet(sql)
GetOpeRegulaPendienteSalida = ""
If Not rs.EOF And Not rs.BOF Then
    GetOpeRegulaPendienteSalida = rs!cOpeCod
End If
rs.Close
Set rs = Nothing
oConect.CierraConexion
Set oConect = Nothing
End Function

Private Sub cmdSalir_Click()
OK = False
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
If Mid(gsOpeCod, 3, 1) = "1" Then
   lMN = True
Else
   lMN = False
End If
lTransActiva = False
lSalir = False
FraCtaIFPagadora.Visible = False
txtImporte = Format(gnImporte, gsFormatoNumeroView)
txtSaldo = Format(lnSaldo, gsFormatoNumeroView)
sMovNroRef = gsMovNro
sCtaPendiente = frmAnalisisRegulaPend.sPendiente
txtNomPers = gsPersNombre
txtNomPers.Tag = lsPersCod

gnDocTpo = -1
MenuItem = 1

optFormPago(1).Enabled = lbActEfectivo
optFormPago(2).Enabled = lbActCheque
optFormPago(3).Enabled = lbActOrdenP
optFormPago(4).Enabled = lbActCarta
optFormPago(5).Enabled = lbActAbonoC
optFormPago(6).Enabled = lbActOtros

Set oOpe = New DOperacion
txtBuscaEntidad.rs = oOpe.GetOpeObj(gsOpeCod, "0")
Set oOpe = Nothing
End Sub

Private Sub optFormPago_Click(Index As Integer)
MenuItem = Index
FraCtaIFPagadora.Visible = False
    Select Case MenuItem
        Case 1    'Efectivo
            cmdAceptar_Click
        Case 2    'Cheque
            FraCtaIFPagadora.Visible = True
        Case 3    'Orden Pago
            cmdAceptar_Click
        Case 4    'Carta
            FraCtaIFPagadora.Visible = True
        Case 5    'Nota de Abono
            gnDocTpo = TpoDocNotaCargo
            cmdAceptar_Click
        Case 6
            If CCur(Abs(txtSaldo)) >= 1 Then
                MsgBox "Saldo no válido para realizar este tipo de Operación. Monto debe ser menor que 1", vbInformation, "Aviso"
                Exit Sub
            Else
                Me.cmdAceptar.Enabled = True
                Me.txtMovDesc.SetFocus
            End If
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

Private Sub optFormPago_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.FraCtaIFPagadora.Visible Then
        Me.txtBuscaEntidad.SetFocus
    Else
        Me.txtMovDesc.SetFocus
    End If
End If
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
Set oCtasIF = New NCajaCtaIF

If txtBuscaEntidad.Text <> "" Then
    lblIFNombre = oCtasIF.NombreIF(Mid(txtBuscaEntidad.Text, 4, 13))
    lblCtaDesc = oCtasIF.EmiteTipoCuentaIF(Mid(txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))) & " " & txtBuscaEntidad.psDescripcion
    
    lblIFNombre = oCtasIF.NombreIF(Mid(txtBuscaEntidad.Text, 4, 13))
    lblCtaDesc = oCtasIF.EmiteTipoCuentaIF(Mid(txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))) & " " & txtBuscaEntidad.psDescripcion
    txtMovDesc.SetFocus
Else
    lblIFNombre = ""
    lblCtaDesc = ""
End If
Set oCtasIF = Nothing
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   cmdAceptar.Enabled = True
   cmdAceptar.SetFocus
End If
End Sub

Function ValidaDatos() As Boolean
ValidaDatos = True
cmdAceptar.Enabled = False
Select Case MenuItem
    Case 1    'Efectivo
    Case 2, 4     ' Carta    'Cheque
        If Len(Trim(txtBuscaEntidad)) = 0 Or lblIFNombre = "" Then
            MsgBox "Cuenta de Institución Financiera no Ingresada", vbInformation, "Aviso"
            ValidaDatos = False
            txtBuscaEntidad.SetFocus
            Exit Function
        End If
    Case 3    'Orden Pago
    Case 5    'Nota de Cargo
End Select

If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Descripción o glosa de Operacion no Ingresada ", vbInformation, "Aviso"
    ValidaDatos = False
    txtMovDesc.SetFocus
    Exit Function
End If
cmdAceptar.Enabled = True
End Function

