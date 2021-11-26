VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCredExtornos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extornos de Credito"
   ClientHeight    =   4245
   ClientLeft      =   3030
   ClientTop       =   3540
   ClientWidth     =   8835
   Icon            =   "frmCredExtornos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   8835
   Begin VB.Frame frmMotExtorno 
      Caption         =   "Motivos del Extorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2700
      Left            =   3105
      TabIndex        =   15
      Top             =   645
      Visible         =   0   'False
      Width           =   2845
      Begin VB.CommandButton cmdExtContinuar 
         Caption         =   "&Continuar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   860
         TabIndex        =   20
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtDetExtorno 
         BackColor       =   &H00C0FFFF&
         Height          =   750
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbMotivos 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmCredExtornos.frx":030A
         Left            =   240
         List            =   "frmCredExtornos.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalles del Extorno"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblExtCmb 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   2970
      TabIndex        =   10
      Top             =   150
      Width           =   4035
      Begin VB.CommandButton cmdBusCli 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2805
         TabIndex        =   12
         Top             =   1155
         Width           =   1005
      End
      Begin VB.TextBox TxtUsu 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   885
         MaxLength       =   12
         TabIndex        =   11
         Tag             =   "txtcodigo"
         Top             =   285
         Width           =   1065
      End
      Begin SICMACT.ActXCodCta ActXCta 
         Height          =   405
         Left            =   135
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   714
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label LblUsu 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   270
         TabIndex        =   13
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operaciones de Extorno"
      Height          =   2415
      Left            =   165
      TabIndex        =   8
      Top             =   1770
      Width           =   8625
      Begin MSComctlLib.ListView LstOpExt 
         Height          =   1995
         Left            =   195
         TabIndex        =   9
         Top             =   255
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3519
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nª Cuenta"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Operacion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hora"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Movimiento"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "CodOpe"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "nPrePago"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "nCuota"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "nCalendActual"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "bExitePagoAnt"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Buscar Por"
      Height          =   1605
      Left            =   135
      TabIndex        =   3
      Top             =   135
      Width           =   2805
      Begin VB.OptionButton opt 
         Caption         =   "&General"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   7
         Top             =   1170
         Width           =   1485
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Cliente"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   6
         Top             =   870
         Width           =   855
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Usuario"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Nro Cuenta"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   585
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1275
      Left            =   7230
      ScaleHeight     =   1215
      ScaleWidth      =   1395
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   330
      Width           =   1455
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   2
         Top             =   720
         Width           =   1245
      End
      Begin VB.CommandButton cmdExtorno 
         Caption         =   "&Extorno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmCredExtornos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nProducto As Producto

Dim vbExtornoDesemb As Boolean
Private fbExtornoPagoHonramiento As Boolean 'WIOR 20131228
Dim nExtornoVigencia As Integer 'LUCV20160523, ERS004-2016
Dim bExtornarVigencia As Boolean 'LUCV20160530, ERS004-2016
'JOEP20190322
Dim nTpOpcionExt As Integer
Dim cCtaCodFalExt As String
Dim cNroMovFalExt As String
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
'JOEP20190322

Public Sub ExtornoDesemb()
    nTpOpcionExt = 1 'JOEP20190325 Mejora de Extorno
    nExtornoVigencia = 0 'LUCV20160523, ERS004-2016
    bExtornarVigencia = False 'LUCV20160530, ERS004-2016
    vbExtornoDesemb = True
    fbExtornoPagoHonramiento = False 'WIOR 20131228
    Me.Show 1
    '109002
End Sub

Public Sub ExtornoVigencia()
    nTpOpcionExt = 2 'JOEP20190325 Mejora de Extorno
    nExtornoVigencia = 1 'LUCV20160523, ERS004-2016
    bExtornarVigencia = True 'LUCV20160530, ERS004-2016
    vbExtornoDesemb = True
    fbExtornoPagoHonramiento = False 'LUCV20160521, ERS004-2016
    Me.Show 1
    '109010
End Sub

Public Sub ExtornoPagos()
    nTpOpcionExt = 3 'JOEP20190325 Mejora de Extorno
    nExtornoVigencia = 0 'LUCV20160523, ERS004-2016
    bExtornarVigencia = False 'LUCV20160530, ERS004-2016
    vbExtornoDesemb = False
    fbExtornoPagoHonramiento = False 'WIOR 20131228
    Me.Show 1
    '109001
End Sub
  
 'WIOR 20131228 *************************************
Public Sub ExtornoPagosHonramiento()
    nTpOpcionExt = 4 'JOEP20190325 Mejora de Extorno
    nExtornoVigencia = 0 'LUCV20160523, ERS004-2016
    bExtornarVigencia = False 'LUCV20160530, ERS004-2016
    vbExtornoDesemb = False
    fbExtornoPagoHonramiento = True
    Me.Show 1
End Sub
'WIOR FIN ******************************************

Private Sub ActxCta_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sCuenta As String
Dim bRetSinTarjeta As Boolean

If KeyCode = vbKeyF12 And ActXCta.Visible = True Then 'F12
        sCuenta = frmValTarCodAnt.Inicia(nProducto, bRetSinTarjeta)
        If sCuenta <> "" Then
            ActXCta.NroCuenta = sCuenta
            ActXCta.SetFocusCuenta
        End If
End If

End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBusCli.SetFocus
    End If
End Sub

Private Sub cmdBusCli_Click()
Dim oCredito As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona
Dim i As Long
Dim L As ListItem
'Dim L As MSComctlLib.ListItem

'JOEP20190322
Dim oExtPag As COMDCredito.DCOMCreditos

If nTpOpcionExt = 3 Then
'Verificamos si falta extornar la cuota que se pago 2 veces.
    Dim RExP As ADODB.Recordset
    Set oExtPag = New COMDCredito.DCOMCreditos
        If cCtaCodFalExt <> "" Then
            If LstOpExt.SelectedItem <> cCtaCodFalExt Then
                Set RExP = oExtPag.SelExtornoPagoNor(cCtaCodFalExt, cNroMovFalExt, 2)
                If Not (RExP.BOF And RExP.EOF) Then
                    If RExP!cMovNro <> "" Then
                        MsgBox "Falta extornar el N° Credito " & cCtaCodFalExt & ", N° Movimiento " & RExP!cMovNro & " y realizar el pago nuevamente", vbInformation, "Aviso"
                        Exit Sub
                    End If
                End If
            Set oExtPag = Nothing
            RSClose RExP
            End If
        End If
'Verificamos si falta extornar la cuota que se pago 2 veces.
End If

    Set oCredito = New COMDCredito.DCOMCreditos
    LstOpExt.ListItems.Clear
    'Busqueda por Cuenta
    If opt(0).value = True Then
        Set R = oCredito.RecuperaDatosExtornoGeneral(gdFecSis, gTipoExtornoCuenta, , ActXCta.NroCuenta, , vbExtornoDesemb, Mid(ActXCta.NroCuenta, 4, 2), , fbExtornoPagoHonramiento, nExtornoVigencia) 'gsCodAge'WIOR 20131228 AGREGO fbExtornoPagoHonramiento /LUCV20160523: bExtonoVigencia
    End If
    'Busqueda por usuario
    If opt(1).value = True Then
        Set R = oCredito.RecuperaDatosExtornoGeneral(gdFecSis, gTipoExtornoUsuario, Trim(TxtUsu.Text), , , vbExtornoDesemb, gsCodAge, , fbExtornoPagoHonramiento, nExtornoVigencia) 'WIOR 20131228 AGREGO fbExtornoPagoHonramiento /LUCV20160523: bExtonoVigencia
    End If
    'Busqueda por Cliente
    If opt(2).value = True Then
        Set oPers = frmBuscaPersona.Inicio
        If Not oPers Is Nothing Then
            Set R = oCredito.RecuperaDatosExtornoGeneral(gdFecSis, gTipoExtornoCliente, , , oPers.sPersCod, vbExtornoDesemb, gsCodAge, , fbExtornoPagoHonramiento, nExtornoVigencia) 'WIOR 20131228 AGREGO fbExtornoPagoHonramiento /LUCV20160523: bExtonoVigencia
        Else
            Exit Sub
        End If
        Set oPers = Nothing
    End If
    'Busqueda General
    If opt(3).value = True Then
        Set R = oCredito.RecuperaDatosExtornoGeneral(gdFecSis, gTipoExtornoGeneral, , , , vbExtornoDesemb, gsCodAge, , fbExtornoPagoHonramiento, nExtornoVigencia) 'WIOR 20131228 AGREGO fbExtornoPagoHonramiento /LUCV20160523: bExtonoVigencia
    End If
    
    If R.RecordCount = 0 Then
        MsgBox "No se encontraron Movimientos", vbInformation, "Aviso"
    Else
        For i = 0 To R.RecordCount - 1
            Set L = LstOpExt.ListItems.Add(, , R!cCtaCod)
            L.SubItems(1) = Trim(R!cMovDesc) & IIf(Trim(R!cMovDesc) = "", "", " - ") & Trim(R!cOpeDesc)
            'L.SubItems(1) = R!cMovDesc
            L.SubItems(2) = R!cHora
            L.SubItems(3) = R!nMovNro
            L.SubItems(4) = Format(R!nMonto, "#0.00")
            L.SubItems(5) = R!cUsuario
            L.SubItems(6) = R!cOpeCod
            L.SubItems(7) = R!nPrepago
            L.SubItems(8) = IIf(IsNull(R!nCuotas), 0, R!nCuotas) 'JOEP20190321
            R.MoveNext
        Next i
    End If
End Sub



'Private Function ValidaExtorno(ByVal psCtaCod As String, ByVal pnMovNro As Long) As Boolean
'Dim odCred As COMDCredito.DCOMCredito
'    ValidaExtorno = True
'    Set odCred = New COMDCredito.DCOMCredito
'    If odCred.PerteneceADesembolsoConCancelacion(pnMovNro, psCtaCod) Then
'        MsgBox "La Operacion es Una Cancelacion que se ha Hecho Con un Desembolso, Esta Operacion se Extornara cuando Extorne el Desembolso", vbInformation, "Aviso"
'        ValidaExtorno = False
'    End If
'    Set odCred = Nothing
'End Function

'**CTI3 (ferimoro) 27092018
Private Sub cmdExtContinuar_Click()

Dim oNCred As COMNCredito.NCOMCredito

'Dim oCredDoc As COMNCredito.NCOMCredDoc
'Dim psDescrip As String
'Dim odCred As COMDCredito.DCOMCredito
'Dim R As ADODB.Recordset
'Dim RCap As ADODB.Recordset
'Dim oBase As COMDCredito.DCOMCredActBD
'Dim nSaldoCtaAho As Double
'Dim lnMovNro As Long
'Dim sUser As String
'Dim oDCredDoc As COMDCredito.DCOMCredDoc
Dim MatDatos(8) As String
Dim sMensaje As String
Dim sImpreBoleta_1 As String
Dim sImpreBoleta_2() As String
Dim sImpreBoletaAho_1() As String
Dim sImpreBoletaAho_2() As String
Dim oPrevio As previo.clsPrevio
Dim i As Integer

'JOEP20190323 Mejora Extorno
Dim oFunFecha As New COMDConstSistema.DCOMGeneral
'JOEP20190323 Mejora Extorno

'****cti3
Dim DatosExtorna(1) As String

If cmbMotivos.ListIndex = -1 Or Len(txtDetExtorno.Text) <= 0 Then
    MsgBox "Debe ingresar el motivo y/o detalle del Extorno", vbInformation, "Aviso"
    Exit Sub
End If

'*** PEAC 20081002
Dim lbResultadoVisto As Boolean
Dim sPersVistoCod  As String
Dim sPersVistoCom As String
Dim loVistoElectronico As frmVistoElectronico
Set loVistoElectronico = New frmVistoElectronico

'*** PEAC 20081001 - visto electronico ******************************************************
'*** en estos extornos de operaciones pedirá visto electrónico

' *** RIRO SEGUN TI-ERS108-2013 ***
Dim nMovNroOperacion As Long
If (IsNumeric(LstOpExt.SelectedItem.SubItems(3))) Then
    nMovNroOperacion = LstOpExt.SelectedItem.SubItems(3)
End If
' *** Fin RIRO ***

If vbExtornoDesemb = False Then
    lbResultadoVisto = loVistoElectronico.Inicio(3, "109001", , , nMovNroOperacion)
    If Not lbResultadoVisto Then
    
        '***CTI3 (ferimoro) *****************
        frmMotExtorno.Visible = False
        Me.cmbMotivos.ListIndex = -1
        Me.txtDetExtorno.Text = ""
        Frame3.Enabled = True
        Frame1.Enabled = True
        cmdExtorno.Enabled = True
        '************************************
    
        Exit Sub
    End If
End If

'*** FIN PEAC ************************************************************

    If MsgBox("Se procederá a realizar el extorno, ¿está seguro de continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        '***CTI3 (ferimoro) *****************
        frmMotExtorno.Visible = False
        Me.cmbMotivos.ListIndex = -1
        Me.txtDetExtorno.Text = ""
        Frame3.Enabled = True
        Frame1.Enabled = True
        cmdExtorno.Enabled = True
        '************************************
        Exit Sub
    End If

'    If Not ValidaExtorno(LstOpExt.SelectedItem.Text, CLng(LstOpExt.SelectedItem.SubItems(3))) Then
'        Exit Sub
'    End If
'
'    'Valida Que CtaAhorros Tenga Saldo
'    If LstOpExt.SelectedItem.SubItems(6) = gCredDesembCtaExist Or LstOpExt.SelectedItem.SubItems(6) = gCredDesembCtaNueva Or LstOpExt.SelectedItem.SubItems(6) = gCredDesembCtaExistDOA Or LstOpExt.SelectedItem.SubItems(6) = gCredDesembCtaNuevaDOA Then
'        Set oBase = New COMDCredito.DCOMCredActBD
'        Set R = New ADODB.Recordset
'        Set RCap = New ADODB.Recordset
'        Set R = oBase.RecuperaMovimientoCapataciones(CLng(LstOpExt.SelectedItem.SubItems(3)))
'        Set RCap = oBase.GetDatosCuentaAho(R!cCtaCod)
'        nSaldoCtaAho = RCap!nSaldoDisp
'        Set oBase = Nothing
'        R.Close
'        RCap.Close
'
'        If CDbl(LstOpExt.SelectedItem.SubItems(4)) > nSaldoCtaAho Then
'            MsgBox "No Existe Saldo Suficiente en la Cuenta de Ahorros para Extornar el Desembolso", vbInformation, "Aviso"
'            Exit Sub
'        End If
'    End If
'
'    If MsgBox("Se va a Extornar la Operacion de la Cuenta : " & LstOpExt.SelectedItem.Text, vbInformation + vbYesNo, "Aviso") = vbNo Then
'        Exit Sub
'    End If
'    lnMovNro = CLng(LstOpExt.SelectedItem.SubItems(3))
'    Set oDCredDoc = New COMDCredito.DCOMCredDoc
'    sUser = oDCredDoc.GetUsuario(lnMovNro)
'     Set oDCredDoc = Nothing
'
'    Set oNCred = New COMNCredito.NCOMCredito
'    Call oNCred.ExtornarCredito(LstOpExt.SelectedItem.Text, gdFecSis, gsCodUser, gsCodAge, CLng(LstOpExt.SelectedItem.SubItems(3)), LstOpExt.SelectedItem.SubItems(6), CDbl(LstOpExt.SelectedItem.SubItems(4)), , , , , CInt(LstOpExt.SelectedItem.SubItems(7)))
'    Set oNCred = Nothing
'    Set oCredDoc = New COMNCredito.NCOMCredDoc
'    If LstOpExt.SelectedItem.SubItems(6) = gCredDesembCtaExist Or LstOpExt.SelectedItem.SubItems(6) = gCredDesembCtaNueva Or LstOpExt.SelectedItem.SubItems(6) = gCredDesembCtaExistDOA Or LstOpExt.SelectedItem.SubItems(6) = gCredDesembCtaNuevaDOA Or LstOpExt.SelectedItem.SubItems(6) = gCredDesembEfec Then
'        psDescrip = "Extorno de Desembolso de Credito"
'    Else
'        psDescrip = "Extorno de Pago de Credito"
'    End If
'
'
'    Call oCredDoc.ImprimeBoletaExtorno(psDescrip, LstOpExt.SelectedItem.Text, "", gsNomAge, _
'        IIf(Mid(LstOpExt.SelectedItem.Text, 9, 1) = "1", "SOLES", "DOLARES"), "", Format(gdFecSis, "dd/mm/yyyy"), _
'        Mid(FechaHora(gdFecSis), 12, 8), "", LstOpExt.SelectedItem.SubItems(4), "0.00", gsCodUser, sLpt, gsInstCmac, gsCodCMAC, sUser)
'
'    psDescrip = "Extorno de Cancelacion"
'    Set odCred = New COMDCredito.DCOMCredito
'    Set R = odCred.CreditosCanceladoConDesembolso(LstOpExt.SelectedItem.Text, CLng(LstOpExt.SelectedItem.SubItems(3)))
'    lnMovNro = CLng(LstOpExt.SelectedItem.SubItems(3))
'    Set oDCredDoc = New COMDCredito.DCOMCredDoc
'    sUser = oDCredDoc.GetUsuario(lnMovNro)
'    Set oDCredDoc = Nothing
'
'    Set oDCredDoc = New COMDCredito.DCOMCredDoc
'    Do While Not R.EOF
'        Call oCredDoc.ImprimeBoletaExtorno(psDescrip, R!cCtaCod, "", gsNomAge, _
'        IIf(Mid(R!cCtaCod, 9, 1) = "1", "SOLES", "DOLARES"), "", Format(gdFecSis, "dd/mm/yyyy"), _
'        Mid(FechaHora(gdFecSis), 12, 8), "", Format(R!nMonto, "#0.00"), "0.00", gsCodUser, sLpt, , , sUser)
'        R.MoveNext
'    Loop
'
'    Set R = odCred.RecuperaMovimientosAhorros(CLng(LstOpExt.SelectedItem.SubItems(3)), True)
'    Do While Not R.EOF
'        Call oCredDoc.ImprimeBoletaExtornoAhorros("Extorno de Abono", R!cCtaCod, "", gsNomAge, "", "", Format(gdFecSis, "dd/mm/yyyy"), Right(FechaHora(gdFecSis), 10), "", R!nMonto, 0, gsCodUser, sLpt, gsInstCmac, gsCodCMAC)
'        R.MoveNext
'    Loop
'
'    Set R = odCred.RecuperaMovimientosAhorros(CLng(LstOpExt.SelectedItem.SubItems(3)), False)
'    Do While Not R.EOF
'        Call oCredDoc.ImprimeBoletaExtornoAhorros("Extorno de Retiros Por Desembolso", R!cCtaCod, "", gsNomAge, "", "", Format(gdFecSis, "dd/mm/yyyy"), Right(FechaHora(gdFecSis), 10), "", R!nMonto, 0, gsCodUser, sLpt, gsInstCmac, gsCodCMAC)
'        R.MoveNext
'    Loop
'    R.Close
'    Set odCred = Nothing
'    Set oCredDoc = Nothing
        
    'ReDim MatDatos(8)
        
    MatDatos(0) = LstOpExt.SelectedItem.Text
    For i = 1 To 7
        MatDatos(i) = LstOpExt.SelectedItem.SubItems(i)
    Next i
    
    '**** cti3
    frmMotExtorno.Visible = False
    DatosExtorna(0) = cmbMotivos.Text
    DatosExtorna(1) = txtDetExtorno.Text
    
    Set oNCred = New COMNCredito.NCOMCredito
    Call oNCred.ExtornarOperacionCredito(MatDatos, gdFecSis, gsCodUser, gsCodAge, gsNomAge, sLpt, gsInstCmac, gsCodCMAC, _
                                        gsUser, sMensaje, sImpreBoleta_1, sImpreBoleta_2, sImpreBoletaAho_1, sImpreBoletaAho_2, gbImpTMU, fbExtornoPagoHonramiento, bExtornarVigencia, DatosExtorna) 'LUCV20160530. Agregó: bExtornarVigencia
    
    '*** PEAC 20081002
        loVistoElectronico.RegistraVistoElectronico (MatDatos(3))
    '*** FIN PEAC
    
    Set oNCred = Nothing
    If sMensaje <> "" Then
        'MsgBox sMensaje, vbInformation, "Mensaje"'Comento JOEP20190408 mejora de extorno de pago
        MsgBox sMensaje, vbInformation, "Aviso" 'JOEP20190408 mejora de extorno de pago
    Else
        Set oPrevio = New previo.clsPrevio
        oPrevio.Show sImpreBoleta_1, ""
        For i = 0 To UBound(sImpreBoleta_2) - 1
            oPrevio.Show sImpreBoleta_2(i), ""
        Next i
        For i = 0 To UBound(sImpreBoletaAho_1) - 1
            oPrevio.Show sImpreBoletaAho_1(i), ""
        Next i
        For i = 0 To UBound(sImpreBoletaAho_2) - 1
            oPrevio.Show sImpreBoletaAho_2(i), ""
        Next i
        Set oPrevio = Nothing
        MsgBox "Extorno Finalizado", vbInformation, "Aviso"
        
        '***CTI3 (ferimoro) *****************
        frmMotExtorno.Visible = False
        Me.cmbMotivos.ListIndex = -1
        Me.txtDetExtorno.Text = ""
        Frame3.Enabled = True
        Frame1.Enabled = True
        cmdExtorno.Enabled = True
        '************************************
        'JOEP20190322
        If nTpOpcionExt = 3 Then
            Dim oExtPag As COMDCredito.DCOMCreditos
            Dim R As ADODB.Recordset
            Set oExtPag = New COMDCredito.DCOMCreditos
            Call oExtPag.UpdExtornoPagoNor(LstOpExt.SelectedItem, LstOpExt.SelectedItem.SubItems(3), LstOpExt.SelectedItem.SubItems(8), oFunFecha.FechaHora(gdFecSis), gsCodUser, 1)
            
            Set R = oExtPag.SelExtornoPagoNor(LstOpExt.SelectedItem, LstOpExt.SelectedItem.SubItems(3), 1)
            If Not (R.BOF And R.EOF) Then
                cCtaCodFalExt = R!cCtaCod
                cNroMovFalExt = R!cMovNro
            Else
                cCtaCodFalExt = ""
                cNroMovFalExt = ""
            End If
            Set oExtPag = Nothing
            RSClose R
        End If
        'JOEP20190322
        
        Call cmdBusCli_Click
    End If
        
    Exit Sub
'ExtornarCredito
End Sub


Private Sub cmdExtorno_Click()
'JOEP20190321
Dim oExtPag As COMDCredito.DCOMCreditos
Dim oDCred As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Dim RVal As ADODB.Recordset
'JOEP20190321

    If LstOpExt.ListItems.count <= 0 Then
        MsgBox "No existen Operaciones para Extornar", vbInformation, "Aviso"
        Exit Sub
    End If

'JOEP20190322
If nTpOpcionExt = 3 Then

'Valida si existe datos para el extorno
Set oDCred = New COMDCredito.DCOMCredito
Set RVal = oDCred.RecuperaDatosExtorno(LstOpExt.SelectedItem.SubItems(3), LstOpExt.SelectedItem)
If RVal.EOF Then
    MsgBox "Actualmente no se puede realizar el extorno de ésta operación." & Chr(13) & "Por favor comuníquese con el Dpto. de TI.", vbInformation, "Aviso"
    Set oDCred = Nothing
    RSClose RVal
    Exit Sub
End If
'Valida si existe datos para el extorno

'Si el movimiento fue extornado por otro usuario
    Set oExtPag = New COMDCredito.DCOMCreditos
    Set rs = oExtPag.ValidaExtornoPagoNor(LstOpExt.SelectedItem, LstOpExt.SelectedItem.SubItems(3), LstOpExt.SelectedItem.SubItems(8), gdFecSis, gsCodUser, LstOpExt.SelectedItem.SubItems(4), gsCodAge, 0)
    If Not (rs.BOF And rs.EOF) Then
        If rs!nMovExt = 1 Then
            MsgBox rs!Mensaje, vbInformation, "Aviso"
            Call cmdBusCli_Click
            Exit Sub
        End If
    End If
    Set oExtPag = Nothing
    RSClose rs
'Si el movimiento fue extornado por otro usuario

'Verificamos si falta extornar la cuota que se pago 2 veces.
Dim R As ADODB.Recordset
Set oExtPag = New COMDCredito.DCOMCreditos
If cCtaCodFalExt <> "" Then
    If LstOpExt.SelectedItem <> cCtaCodFalExt Then
        Set R = oExtPag.SelExtornoPagoNor(cCtaCodFalExt, cNroMovFalExt, 2)
        If Not (R.BOF And R.EOF) Then
            If R!cMovNro <> "" Then
                MsgBox "Falta extornar el N° Credito " & cCtaCodFalExt & ", N° Movimiento " & R!cMovNro & " y realizar el pago nuevamente", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        Set oExtPag = Nothing
        RSClose R
    End If
End If
'Verificamos si falta extornar la cuota que se pago 2 veces.

MsgBox "Se está  extornando la Cuota N°: " & LstOpExt.SelectedItem.SubItems(8) & Chr(13) & " del Crédito " & LstOpExt.SelectedItem, vbInformation, "Aviso"
    
'Si existen varios pagos, verificamos que se extorne de mayor a menor. - 'Cuando se paga la misma cuota 2 veces.
Set oExtPag = New COMDCredito.DCOMCreditos
Set rs = oExtPag.ValidaExtornoPagoNor(LstOpExt.SelectedItem, LstOpExt.SelectedItem.SubItems(3), LstOpExt.SelectedItem.SubItems(8), gdFecSis, gsCodUser, LstOpExt.SelectedItem.SubItems(4), gsCodAge, 0)
If Not (rs.BOF And rs.EOF) Then
    If rs!Mensaje <> "" Then
        MsgBox rs!Mensaje, vbInformation, "Aviso"
        If rs!Op <> 1 Then
            Exit Sub
        End If
    End If
End If
Set oExtPag = Nothing
RSClose rs
'Si existen varios pagos, verificamos que se extorne de mayor a menor.
End If
'JOEP20190321

'******CTI3 (ferimoro) 27092018
 frmMotExtorno.Visible = True
 Frame3.Enabled = False
 Frame1.Enabled = False
 cmdExtorno.Enabled = False
 cmbMotivos.SetFocus
'******************************
End Sub

Private Sub cmdSalir_Click()
'JOEP20190322
If nTpOpcionExt = 3 Then
    Dim oExtPag As COMDCredito.DCOMCreditos
    Dim R As ADODB.Recordset
    Set oExtPag = New COMDCredito.DCOMCreditos
    
    If cCtaCodFalExt <> "" Then
        Set R = oExtPag.SelExtornoPagoNor(cCtaCodFalExt, cNroMovFalExt, 2)
        If Not (R.BOF And R.EOF) Then
            If R!cMovNro <> "" Then
                MsgBox "Falta extornar el N° Credito " & cCtaCodFalExt & ", N° Movimiento " & R!cMovNro & " y realizar el pago nuevamente", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        Set oExtPag = Nothing
        RSClose R
    End If
End If
'JOEP20190322

    Unload Me
    
End Sub
'***CTI3 (feirmoro)  18102018
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

Set oCons = New COMDConstantes.DCOMConstantes

Set R = oCons.ObtenerConstanteExtornoMotivo

Set oCons = Nothing
Call Llenar_Combo_MotivoExtorno(R, cmbMotivos)

End Sub
Private Sub Form_Load()
'JOEP20190322
    DisableCloseButton Me
    cCtaCodFalExt = ""
    cNroMovFalExt = ""
    nTpOpcionExt = 0 'JOEP20190325 Mejora de Extorno
'JOEP20190322
    CentraForm Me
    ActXCta.CMAC = gsCodCMAC
    ActXCta.Age = gsCodAge
    Call CargaControles
End Sub

Private Sub opt_Click(Index As Integer)
    Select Case Index
        Case 0
            ActXCta.Visible = True
            LblUsu.Visible = False
            TxtUsu.Visible = False
            ActXCta.NroCuenta = ""
            ActXCta.CMAC = gsCodCMAC
            ActXCta.Age = gsCodAge
        Case 1
            ActXCta.Visible = False
            LblUsu.Visible = True
            TxtUsu.Visible = True
        Case Else
            ActXCta.Visible = False
            LblUsu.Visible = False
            TxtUsu.Visible = False
    End Select
End Sub

Private Sub txtDetExtorno_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub
End Sub

Private Sub TxtUsu_GotFocus()
    fEnfoque TxtUsu
End Sub

Private Sub TxtUsu_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        cmdBusCli.SetFocus
    End If
End Sub
'JOEP20190322
Public Function DisableCloseButton(frm As Form) As Boolean
'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu
    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
    
    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)

    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)
   'Remove seperator bar
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    'Return True if both calls were successful
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)
End Function
'JOEP20190322
