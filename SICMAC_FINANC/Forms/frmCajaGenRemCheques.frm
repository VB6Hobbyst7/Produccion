VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaGenRemCheques 
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   Icon            =   "frmCajaGenRemCheques.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraCuentaDesde 
      Caption         =   "Banco:"
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
      Height          =   1035
      Left            =   90
      TabIndex        =   21
      Top             =   270
      Width           =   7965
      Begin Sicmact.TxtBuscar txtCtaIFDesde 
         Height          =   315
         Left            =   1155
         TabIndex        =   22
         Top             =   255
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
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
         ForeColor       =   -2147483635
      End
      Begin VB.Label lblDescCtaDesde 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1155
         TabIndex        =   27
         Top             =   637
         Width           =   6630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Institución :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   4035
         TabIndex        =   26
         Top             =   315
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   135
         TabIndex        =   25
         Top             =   690
         Width           =   960
      End
      Begin VB.Label lblDescIFDesde 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   4890
         TabIndex        =   24
         Top             =   262
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta IF:"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   315
         Width           =   735
      End
   End
   Begin VB.Frame FraRetiro 
      Caption         =   "Emisión de Documentos"
      Height          =   930
      Left            =   300
      TabIndex        =   15
      Top             =   2850
      Width           =   7350
      Begin VB.OptionButton OptDoc 
         Caption         =   "&Carta"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   375
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.OptionButton OptDoc 
         Caption         =   "Che&que"
         Height          =   375
         Index           =   1
         Left            =   5865
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   375
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdRetirar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4890
      TabIndex        =   14
      Top             =   3990
      Width           =   1380
   End
   Begin VB.Frame FraDeposito 
      Caption         =   "Efectivo"
      Height          =   1020
      Left            =   300
      TabIndex        =   8
      Top             =   2820
      Width           =   7350
      Begin VB.CommandButton cmdefectivo 
         Caption         =   "&Efectivo"
         Height          =   375
         Left            =   615
         TabIndex        =   13
         Top             =   315
         Width           =   1470
      End
      Begin VB.Frame FraDocumento 
         Caption         =   "Documento"
         Height          =   690
         Left            =   2820
         TabIndex        =   9
         Top             =   195
         Width           =   4425
         Begin VB.TextBox txtNroDoc 
            Height          =   315
            Left            =   2400
            TabIndex        =   11
            Top             =   240
            Width           =   1830
         End
         Begin VB.ComboBox cboDocumento 
            Height          =   315
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   210
            Width           =   1890
         End
         Begin VB.Label Label7 
            Caption         =   "N° :"
            Height          =   195
            Left            =   2100
            TabIndex        =   12
            Top             =   285
            Width           =   255
         End
      End
   End
   Begin VB.Frame FraOrigen 
      Caption         =   "Agencia:"
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
      Height          =   690
      Left            =   330
      TabIndex        =   5
      Top             =   1320
      Width           =   7335
      Begin Sicmact.TxtBuscar txtCtaOrig 
         Height          =   345
         Left            =   240
         TabIndex        =   6
         Top             =   225
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   609
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
      Begin VB.Label lblDescCtaOrig 
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
         Height          =   315
         Left            =   2145
         TabIndex        =   7
         Top             =   240
         Width           =   5025
      End
   End
   Begin VB.TextBox txtmonto 
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
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   3990
      Width           =   1740
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   780
      Left            =   300
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2070
      Width           =   7395
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6285
      TabIndex        =   1
      Top             =   3990
      Width           =   1380
   End
   Begin VB.CommandButton cmdDepositar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4905
      TabIndex        =   0
      Top             =   3990
      Width           =   1380
   End
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   345
      Left            =   6390
      TabIndex        =   4
      Top             =   30
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   609
      _Version        =   393216
      ForeColor       =   -2147483635
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Importe :"
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
      Height          =   195
      Left            =   480
      TabIndex        =   20
      Top             =   4050
      Width           =   765
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   390
      Left            =   375
      Top             =   3975
      Width           =   2790
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   5655
      TabIndex        =   19
      Top             =   45
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Importe :"
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
      Height          =   195
      Left            =   480
      TabIndex        =   18
      Top             =   4080
      Width           =   765
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   7695
      Y1              =   3870
      Y2              =   3870
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   240
      X2              =   7695
      Y1              =   3900
      Y2              =   3900
   End
End
Attribute VB_Name = "frmCajaGenRemCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oOpe As DOperacion
Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset
Dim oCtaIf As NCajaCtaIF
Dim lnTipoObj As TpoObjetos
Dim lsNroVoucher As String
Dim lsNombreBanco As String
Dim lsCuenta As String

Dim lsNroDoc As String
Dim lsDocumento As String
Dim lnTipoDoc As TpoDoc
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cboDocumento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNroDoc.SetFocus
End If
End Sub

Private Sub chkcarta1_Click(Index As Integer)

End Sub

Private Sub cmdDepositar_Click()
Dim oCajero As nCajaGeneral
Dim oCont As NContFunciones
Set oCont = New NContFunciones
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim lsCtafiltro As String
Dim lsMovNro As String
If Valida = False Then Exit Sub

Set oCajero = New nCajaGeneral
If MsgBox("Desea Grabar el movimiento respectivo??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oCont.GeneraMovNro(CDate(txtFecha), gsCodAge, gsCodUser)
    'lsCtafiltro = oCont.GetFiltroObjetos(lnTipoObj, txtCtaDest, txtObjDest, False)
    'oCajero.GrabaMovEfectivo lsMovNro, gsOpeCod, txtMovDesc, _
    '            rsBill, rsMon, txtCtaDest + lsCtafiltro, txtCtaOrig, txtmonto, lnTipoObj, _
    '            txtObjDest, Val(Right(cboDocumento, 2)), txtNroDoc, gdFecSis

'    ImprimeAsientoContable lsMovNro
'    Set frmCajaGenEfectivo = Nothing
'    If MsgBox("Desea realizar otra operación??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
'        txtCtaDest = ""
'        lblDescCtaDest = ""
'        lblDescObjDest = ""
'        txtObjDest = ""
'        txtMovDesc = ""
'        txtmonto = "0.00"
'        cboDocumento.ListIndex = -1
'        txtNroDoc = ""
'        Set rsBill = Nothing
'        Set rsMon = Nothing
'
'    Else
'        Unload Me
'
'    End If
End If
End Sub
Function Valida() As Boolean
Valida = True
If ValFecha(txtFecha) = False Then
    Valida = False
    Exit Function
End If
If Len(Trim(txtCtaOrig)) = 0 Then
    MsgBox "Cuenta de " & FraOrigen.Caption & " no seleccionada", vbInformation, "Aviso"
    Valida = False
    If txtCtaOrig.Enabled Then txtCtaOrig.SetFocus
    Exit Function
End If
'If Trim(Len(txtCtaDest)) = 0 Then
'    MsgBox "Cuenta de " & FraDestino.Caption & " no seleccionada", vbInformation, "Aviso"
'    Valida = False
'    If txtCtaDest.Enabled Then txtCtaDest.SetFocus
'    Exit Function
'End If
'If Trim(Len(txtObjDest)) = 0 And txtObjDest.Enabled Then
'    MsgBox "Objeto " & FraDestino.Caption & " no seleccionada", vbInformation, "Aviso"
'    Valida = False
'    If txtObjDest.Enabled Then txtObjDest.SetFocus
'    Exit Function
'End If
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Descripción de movimiento no ingresado", vbInformation, "Aviso"
    Valida = False
    txtMovDesc.SetFocus
    Exit Function
End If
If FraDeposito.Visible Then
    If rsBill Is Nothing And rsMon Is Nothing Then
        MsgBox "Billetaje no ha sido ingresado", vbInformation, "Aviso"
        Valida = False
        cmdDepositar.SetFocus
        Exit Function
    End If
    If Len(Trim(cboDocumento)) <> 0 Then
        If Len(Trim(txtNroDoc)) = 0 Then
            MsgBox "Nro de Documento no Ingresado", vbInformation, "Aviso"
            Valida = False
            txtNroDoc.SetFocus
            Exit Function
        End If
    Else
        If MsgBox("Documento no ha sido ingresado. Desea continuar??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            Valida = False
            cboDocumento.SetFocus
            Exit Function
        End If
    End If
End If
If Val(txtmonto) = 0 Then
    MsgBox "Importe de movimiento no ingresado", vbInformation, "Aviso"
    Valida = False
    txtmonto.SetFocus
    Exit Function
End If

If Len(Trim(txtCtaIFDesde)) = 0 Then
    MsgBox "Ingrese Cuenta de Institución Financiera Inicial", vbInformation, "Aviso"
    txtCtaIFDesde.SetFocus
    Valida = False
    Exit Function
End If


End Function
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdefectivo_Click()
    Set rsBill = New ADODB.Recordset
    Set rsMon = New ADODB.Recordset
    frmCajaGenEfectivo.Inicio gsOpeCod, gsOpeDesc, 0, Mid(gsOpeCod, 3, 1), False
    If frmCajaGenEfectivo.lbOk Then
        Set rsBill = frmCajaGenEfectivo.rsBilletes
        Set rsMon = frmCajaGenEfectivo.rsMonedas
        txtmonto = frmCajaGenEfectivo.lblTotal
        If FraDocumento.Visible Then
           cboDocumento.SetFocus
        End If
    Else
        Set rsBill = Nothing
        Set rsMon = Nothing
    End If
End Sub
Private Sub cmdRetirar_Click()
Dim oCajero As nCajaGeneral
Dim oCont As NContFunciones
Set oCont = New NContFunciones
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim lsCtafiltro As String
Dim lsMovNro As String
Dim lsCtaIFCod As String
Dim lsIFTpo As String
Dim lsPersCod As String
Dim pnMoneda As Integer
Dim oPrevio As clsPrevioFinan
Set oPrevio = New clsPrevioFinan


If Valida = False Then Exit Sub

If lsNroDoc = "" Or lsDocumento = "" Then
    MsgBox "No ha seleccionado el documento Utilizado en Operación", vbInformation, "aviso"
    OptDoc(1).SetFocus
    Exit Sub
End If
            
If Mid(gsOpeCod, 3, 1) = "1" Then
    pnMoneda = 1
Else
    pnMoneda = 2
End If

lsIFTpo = Mid(txtCtaIFDesde, 1, 2)
lsPersCod = Mid(txtCtaIFDesde, 4, 13)
lsCtaIFCod = Mid(txtCtaIFDesde, 18, 7)

Set oCajero = New nCajaGeneral
If MsgBox("Desea Grabar el movimiento respectivo??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oCont.GeneraMovNro(CDate(txtFecha), gsCodAge, gsCodUser)
    
    oCajero.GrabaMovCheques lsMovNro, gsOpeCod, txtMovDesc, CCur(txtmonto), txtCtaOrig, _
               lsIFTpo, lsPersCod, lsCtaIFCod, lnTipoDoc, lsNroDoc, gdFecSis, lsNroVoucher, pnMoneda

' psDocumento & oImpresora.gPrnSaltoPagina
  '  oPrevio.Show lsDocumento, gsOpeDesc, False, gnLinPage, gImpresora

    ImprimeAsientoNoContable lsMovNro, lsNroVoucher, lnTipoDoc, lsDocumento
    Set frmCajaGenEfectivo = Nothing
    
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
                Set objPista = Nothing
                '****
    
    If MsgBox("Desea realizar otra operación??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        txtCtaIFDesde = ""
        lblDescCtaDesde = ""
        lblDescIFDesde = ""
        txtCtaOrig = ""
        txtMovDesc = ""
        txtmonto = "0.00"
        txtNroDoc = ""
        OptDoc(0).value = False
        OptDoc(1).value = False
        lsNroDoc = ""
        lsDocumento = ""
        Set rsBill = Nothing
        Set rsMon = Nothing

    Else
        Unload Me

    End If
End If
End Sub
Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim oAgencias As DActualizaDatosArea
Set rs = New ADODB.Recordset
Set oOpe = New DOperacion
Set oCtaIf = New NCajaCtaIF
Set oAgencias = New DActualizaDatosArea

CentraForm Me
Me.Caption = gsOpeDesc
txtFecha = gdFecSis
Set rs = oOpe.CargaOpeDoc(gsOpeCod)
Do While Not rs.EOF
   cboDocumento.AddItem Mid(rs!cDocDesc & Space(100), 1, 100) & rs!nDocTpo
   rs.MoveNext
Loop
CambiaTamañoCombo cboDocumento
cmdDepositar.Visible = False
cmdRetirar.Visible = False
FraDeposito.Visible = False
FraRetiro.Visible = False
txtmonto.Locked = False
Select Case gsOpeCod
    Case gOpeCGOpeBancosDepEfecMN, gOpeCGOpeBancosDepEfecME
        txtCtaOrig.psRaiz = "Cuentas Contables"
       ' txtCtaOrig.rs = oOpe.CargaOpeCta(gsOpeCod, "H", "0")
         
         
                
        'txtCtaDest.rs = oOpe.CargaOpeCta(gsOpeCod, "D", "0")
        cmdDepositar.Visible = True
        FraDeposito.Visible = True
        FraOrigen.Caption = "Origen"
        'FraDestino.Caption = "Destino"
        txtmonto.Locked = True
    Case gOpeCGRemChequesMN, gOpeCGRemChequesME
        FraRetiro.Visible = True
        cmdRetirar.Visible = True
        FraOrigen.Caption = "Destino"
        'FraDestino.Caption = "Origen"
        
        txtCtaOrig.rs = oAgencias.GetAgencias
        'txtCtaOrig.psRaiz = "Cuentas Contables"
        'txtCtaOrig.rs = oOpe.CargaOpeCta(gsOpeCod, "D", "0")
        'txtCtaDest.psRaiz = "Cuentas Contables"
        'txtCtaDest.rs = oOpe.CargaOpeCta(gsOpeCod, "H", "0")

End Select

txtCtaIFDesde.rs = oOpe.GetRsOpeObj(gsOpeCod, "0")

End Sub









Private Sub OptDoc_Click(Index As Integer)
Dim oDocPago As clsDocPago
Dim oNContFunc As NContFunciones
Dim oCtasIF As NCajaCtaIF

Set oDocPago = New clsDocPago

Dim lsSubCuentaIF As String

lsNroDoc = ""
lsDocumento = ""
lsNroVoucher = ""

If Valida = False Then
    OptDoc(Index).value = False
    Exit Sub
End If
Select Case Index
    Case 0
        oDocPago.InicioCarta "", Mid(txtCtaIFDesde, 4, 13), gsOpeCod, _
                            gsOpeDesc, txtMovDesc, "", CCur(txtmonto), gdFecSis, _
                            lblDescCtaDesde, "", gsNomCmac, "", ""
        If oDocPago.vbOk Then
            lnTipoDoc = Val(oDocPago.vsTpoDoc)
            lsNroDoc = oDocPago.vsNroDoc
            lsDocumento = oDocPago.vsFormaDoc
            cmdRetirar.SetFocus
        Else
            lsNroDoc = ""
            lsDocumento = ""
            lsNroVoucher = ""
            OptDoc(Index).value = False
            Set oDocPago = Nothing
            Exit Sub
        End If
    Case 1
        Set oNContFunc = New NContFunciones
        Set oCtasIF = New NCajaCtaIF

        lsNroVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1), gsCodAge)
        lsSubCuentaIF = oCtasIF.SubCuentaIF(Mid(txtCtaIFDesde, 4, 13))

        oDocPago.InicioCheque "", True, Mid(txtCtaIFDesde, 4, 13), gsOpeCod, gsNomCmac, gsOpeDesc, txtMovDesc, CCur(txtmonto), _
                        gdFecSis, "", lsSubCuentaIF, lblDescIFDesde, "", lsNroVoucher, True, , , , Mid(txtCtaIFDesde, 1, 2), Mid(txtCtaIFDesde, 4, 13), Mid(txtCtaIFDesde, 18, 10) 'EJVG20121207
                        'gdFecSis, "", lsSubCuentaIF, lblDescIFDesde, "", lsNroVoucher, True ', gsCodAge, Mid(txtCtaIFDesde, 18, 10)
      
        
        'oDocPago.InicioCheque "", True, Mid(txtObjDest, 4, 13), gsOpeCod, gsNomCmac, gsOpeDesc, txtMovDesc, CCur(txtmonto), _
        '                gdFecSis, "", lsSubCuentaIF, lsNombreBanco, lsCuenta, lsNroVoucher, True, gsCodAge, Mid(txtObjDest, 18, 10)
        Set oNContFunc = Nothing
        If oDocPago.vbOk Then
            lsNroDoc = ""
            lsDocumento = ""
            lsNroVoucher = ""
            lnTipoDoc = Val(oDocPago.vsTpoDoc)
            lsNroDoc = oDocPago.vsNroDoc
            lsDocumento = oDocPago.vsFormaDoc
            lsNroVoucher = oDocPago.vsNroVoucher
            cmdRetirar.SetFocus
        Else
            OptDoc(Index).value = False
            Set oDocPago = Nothing
            Exit Sub
        End If
End Select
Set oDocPago = Nothing

End Sub

Private Sub txtCtaDest_EmiteDatos()
Dim lsRaiz As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
'txtObjDest = ""
'lblDescObjDest = ""
'lblDescCtaDest = txtCtaDest.psDescripcion
'If txtCtaDest <> "" Then
'    Set rs = AsignaCtaObj(txtCtaDest, lsRaiz)
'    txtObjDest.psRaiz = lsRaiz
'    txtObjDest.rs = rs
'    txtObjDest.SetFocus
'End If
End Sub




Private Sub txtCtaIFDesde_EmiteDatos()
lblDescCtaDesde = oCtaIf.EmiteTipoCuentaIF(Mid(txtCtaIFDesde, 18, 10)) + " " + txtCtaIFDesde.psDescripcion
lblDescIFDesde = oCtaIf.NombreIF(Mid(txtCtaIFDesde, 4, 13))
txtCtaOrig.SetFocus
End Sub


Private Sub txtCtaOrig_EmiteDatos()
lblDescCtaOrig = txtCtaOrig.psDescripcion
txtMovDesc.SetFocus
End Sub

Private Sub txtmonto_GotFocus()
fEnfoque txtmonto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtmonto, KeyAscii)
If KeyAscii = 13 Then
    If cmdDepositar.Visible Then cmdDepositar.SetFocus
    If cmdRetirar.Visible Then cmdRetirar.SetFocus
End If
End Sub

Private Sub txtMonto_LostFocus()
If Trim(Len(txtmonto)) = 0 Then txtmonto = 0
txtmonto = Format(txtmonto, "#,#0.00")
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If FraDeposito.Visible Then
        cmdefectivo.SetFocus
    End If
    If FraRetiro.Visible Then
        txtmonto.SetFocus
    End If
    
End If
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    txtmonto.SetFocus
End If
End Sub
Private Function AsignaCtaObj(ByVal psCtaContCod As String, ByRef lsRaiz As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
'Dim lsRaiz As String
Dim oDescObj As ClassDescObjeto
Dim UP As UPersona
Dim lsFiltro As String
Dim oRHAreas As DActualizaDatosArea
Dim oCtaCont As DCtaCont
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo

Set oEfect = New Defectivo
Set oCtaIf = New NCajaCtaIF
Set oRHAreas = New DActualizaDatosArea
Set oDescObj = New ClassDescObjeto
Set oCtaCont = New DCtaCont
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset

Set rs1 = oCtaCont.CargaCtaObj(psCtaContCod, , True)
If Not rs1.EOF And Not rs1.BOF Then
    lsRaiz = ""
    lsFiltro = ""
    lnTipoObj = Val(rs1!cObjetoCod)
    Select Case Val(rs1!cObjetoCod)
        Case ObjCMACAgencias
            Set rs = oRHAreas.GetAgencias(rs1!cCtaObjFiltro)
        Case ObjCMACAgenciaArea
            lsRaiz = "Unidades Organizacionales"
            Set rs = oRHAreas.GetAgenciasAreas(rs1!cCtaObjFiltro)
        Case ObjCMACArea
            Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
        Case ObjEntidadesFinancieras
'            lsRaiz = "Cuentas de Entidades Financieras"
             lsRaiz = ""
            'Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, psCtaContCod)
            Set rs = oCtaIf.CargaCtasIF(Mid(psCtaContCod, 3, 1), rs1!cCtaObjFiltro)
        Case ObjDescomEfectivo
            lsRaiz = "Denominación"
            Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
        Case ObjPersona
            Set rs = Nothing
        Case Else
            Set rs = GetObjetos(Val(rs1!cObjetoCod))
    End Select
End If
rs1.Close
Set rs1 = Nothing
Set AsignaCtaObj = rs

Set oDescObj = Nothing
Set UP = Nothing
Set oCtaCont = Nothing
Set oCtaIf = Nothing
Set oEfect = Nothing
End Function
Private Sub txtObjDest_EmiteDatos()
'lsNombreBanco = ""
'lsCuenta = ""
'lblDescObjDest = txtObjDest.psDescripcion
'If lblDescObjDest <> "" Then
'    If Len(txtObjDest) > 15 Then
'         lsNombreBanco = oCtaIf.NombreIF(Mid(txtObjDest, 4, 13))
'         lsCuenta = oCtaIf.EmiteTipoCuentaIF(Mid(txtObjDest, 18, 10)) + " " + txtObjDest.psDescripcion
'         lblDescObjDest = lsNombreBanco + " " + lsCuenta
'    End If
'    txtMovDesc.SetFocus
'End If
End Sub
Private Sub txtObjDest_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMovDesc.SetFocus
End If
End Sub


