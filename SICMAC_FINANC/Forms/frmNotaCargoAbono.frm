VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNotaCargoAbono 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nota de Cargo / Abono"
   ClientHeight    =   5385
   ClientLeft      =   2430
   ClientTop       =   2220
   ClientWidth     =   6555
   Icon            =   "frmNotaCargoAbono.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   5220
      TabIndex        =   7
      Top             =   4980
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   4020
      TabIndex        =   6
      Top             =   4980
      Width           =   1215
   End
   Begin VB.Frame fraIngNotaCargo 
      Caption         =   "Nota de Cargo"
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
      Height          =   4830
      Left            =   60
      TabIndex        =   1
      Top             =   45
      Width           =   6360
      Begin Sicmact.ActXCodCta txtCtaCaptaNroCentral 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   270
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   661
         Texto           =   "Cuenta N°"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.ListBox lstCuentas 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   3795
         TabIndex        =   14
         Top             =   780
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgTitulares 
         Height          =   1290
         Left            =   180
         TabIndex        =   3
         Top             =   2340
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   2275
         _Version        =   393216
         Cols            =   5
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin VB.TextBox txtMonto 
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
         ForeColor       =   &H00404080&
         Height          =   300
         Left            =   4635
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   4380
         Width           =   1515
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   540
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3675
         Width           =   6045
      End
      Begin MSMask.MaskEdBox txtFechaNC 
         Height          =   345
         Left            =   5040
         TabIndex        =   2
         Top             =   285
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin Sicmact.TxtBuscar txtBuscarPers 
         Height          =   345
         Left            =   135
         TabIndex        =   15
         Top             =   870
         Width           =   2010
         _ExtentX        =   3545
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
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Frame fradoc 
         Height          =   585
         Left            =   135
         TabIndex        =   16
         Top             =   4185
         Width           =   2835
         Begin VB.TextBox txtnroDoc 
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
            Height          =   330
            Left            =   1050
            MaxLength       =   8
            TabIndex        =   18
            Top             =   180
            Width           =   1620
         End
         Begin VB.Label lblDescDoc 
            AutoSize        =   -1  'True
            Caption         =   "NroDoc :"
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
            Left            =   90
            TabIndex        =   17
            Top             =   225
            Width           =   780
         End
      End
      Begin VB.Frame framotivo 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   2145
         TabIndex        =   19
         Top             =   1920
         Visible         =   0   'False
         Width           =   4020
         Begin VB.CommandButton cmdNotacargoAbono 
            Height          =   345
            Left            =   3135
            Picture         =   "frmNotaCargoAbono.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   30
            Width           =   840
         End
         Begin VB.ComboBox cboNANC 
            Height          =   315
            Left            =   345
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   60
            Width           =   2730
         End
      End
      Begin Sicmact.ActXCodCta_Ant txtCtaCaptaNro 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
      End
      Begin VB.Label lblUbigeo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3240
         TabIndex        =   13
         Top             =   1590
         Width           =   2925
      End
      Begin VB.Label lblPersDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   135
         TabIndex        =   12
         Top             =   1590
         Width           =   3075
      End
      Begin VB.Label lblPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   135
         TabIndex        =   11
         Top             =   1245
         Width           =   6030
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00404080&
         Height          =   195
         Left            =   3840
         TabIndex        =   10
         Top             =   4440
         Width           =   660
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   4530
         TabIndex        =   9
         Top             =   337
         Width           =   555
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Datos Titular(es) :"
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
         Left            =   150
         TabIndex        =   8
         Top             =   2070
         Width           =   1545
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   360
         Left            =   3615
         Top             =   4365
         Width           =   2550
      End
   End
   Begin VB.Label lblITFValor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   300
      Left            =   1185
      TabIndex        =   24
      Top             =   4980
      Width           =   1545
   End
   Begin VB.Label lblITF 
      Caption         =   "ITF"
      Height          =   240
      Left            =   135
      TabIndex        =   23
      Top             =   5010
      Width           =   1170
   End
End
Attribute VB_Name = "frmNotaCargoAbono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbOk As Boolean
Dim lsNroNotaCA As String
Dim ldFechaNotaCA As String
Dim lsPersCod As String
Dim lsPersNombre As String
Dim lsCuentaAhoNro As String
Dim lsGlosa As String
Dim lsNotaCargoAbono As String
Dim lsPersDireccion As String
Dim lsPersUbigeo As String
Dim lbMotivo As Boolean

Dim oCapta As NCapMantenimiento

Dim lnMotivo As MotivoNotaAbonoCargo
Dim lsObjetoMotivo As String
Dim lsObjetoMotivoPadre As String
Dim lbCargo As Boolean
Dim lnMonto As Currency
Dim pnMonto As Currency
Dim ldFecha As Date
Dim lnTpoDoc As TpoDoc
Dim oContFun As NContFunciones
Dim oDoc As DDocumento
Dim oDocRec As NDocRec
Dim lsOpeCod As String
Dim lnPersoneria As PersPersoneria
Dim lsCuenta  As String
Dim lbModificaMonto As Boolean
'Public lnITFValor As Currency
Public lnITFValor As Double '*** PEAC 20110331

Public Sub Inicio(ByVal pnTpoDoc As TpoDoc, ByVal pnMonto As Currency, ByVal pdFecha As Date, ByVal psGlosa As String, ByVal psOpeCod As String, Optional pbMotivo As Boolean = False, Optional psPersCod As String, Optional psPersNombre As String, Optional psCuenta As String = "", Optional pbModificaMonto As Boolean = False, Optional pnITFValor As Double) '*** PEAC 20110331 - Optional pnITFValor As Currency
lnMonto = pnMonto
ldFecha = pdFecha
lsGlosa = psGlosa
lsOpeCod = psOpeCod
lnTpoDoc = pnTpoDoc
lbMotivo = pbMotivo
lsPersCod = psPersCod
lsPersNombre = psPersNombre
lsCuenta = psCuenta
lbModificaMonto = pbModificaMonto
lnITFValor = pnITFValor

'txtBuscarPers

Me.Show 1
End Sub

Private Sub cboNANC_Click()
If cboNANC.ListCount > 0 Then
    If Trim(Right(cboNANC, 10)) = "<<Nuevo>>" Then
        cmdNotacargoAbono.value = True
    End If
End If
End Sub

Private Sub cboNANC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMonto.SetFocus
End If
End Sub

Private Sub cmdAceptar_Click()
Dim rs As ADODB.Recordset
Dim lsAgeDesc As String
Set rs = New ADODB.Recordset
If ValidaDatos = False Then Exit Sub
vbOk = True
lsNroNotaCA = Trim(txtnroDoc)
ldFechaNotaCA = txtFechaNC
lsPersCod = txtBuscarPers.Text
lsPersNombre = PstaNombre(lblPersNombre)
lsCuentaAhoNro = IIf(gbBitCentral, txtCtaCaptaNroCentral.NroCuenta, txtCtaCaptaNro.NroCuenta)
lsGlosa = Trim(txtDescripcion)
lsPersDireccion = lblPersDireccion
lsPersUbigeo = lblUbigeo
pnMonto = CCur(txtMonto)

Dim oImp As New NContImprimir
Dim oAre As New DActualizaDatosArea
If gbBitCentral Then
    lsAgeDesc = oAre.GetNombreAgencia(Me.txtCtaCaptaNroCentral.Age)
    lsNotaCargoAbono = oImp.ImprimeNotaCargoAbono(lsNroNotaCA, lsGlosa, lnMonto, lsPersNombre, lsPersDireccion, lsPersUbigeo, ldFechaNotaCA, Mid(gsOpeCod, 3, 1), lsCuentaAhoNro, lnTpoDoc, lsAgeDesc, gsCodUser)
Else
    lsNotaCargoAbono = oImp.ImprimeNotaAbono(Format(ldFechaNotaCA, gsFormatoFecha), pnMonto, lsGlosa, lsCuentaAhoNro, lsPersNombre)
End If

If framotivo.Visible Then
   lsNroNotaCA = Trim(Left(cboNANC, 10))
   Set rs = oDocRec.GetDatosNotaAC(lnTpoDoc, lsNroNotaCA, gNCNARegistrado)
   If Not rs.EOF And Not rs.BOF Then
        lnMotivo = rs!nMotivoCod
        lsObjetoMotivoPadre = rs!cObjetoCodPadre
        lsObjetoMotivo = rs!cObjetoCod
        If txtMonto.Locked = True Then
           If rs!nMonto <> CCur(txtMonto) Then
                MsgBox "Monto de Nota seleccionada [" & Format(rs!nMonto, "#,#0.00") & "] no coincide con Monto de Operación" + vbCrLf + "Registre una Nota con el Monto de Operación", vbInformation, "Aviso"
                cmdNotacargoAbono.SetFocus
                rs.Close
                Set rs = Nothing
                Exit Sub
           End If
        End If
   Else
        MsgBox "Nota Seleccionada ya ha sido confirmada", vbInformation, "Aviso"
        rs.Close
        Set rs = Nothing
        Exit Sub
   End If
   rs.Close
   Set rs = Nothing
End If
Unload Me
End Sub

Function ValidaDatos() As Boolean
ValidaDatos = True

If gbBitCentral Then
    If Len(txtCtaCaptaNroCentral.NroCuenta) <> 18 Then
        MsgBox "Cuenta de Ahorros no válida", vbInformation, "Aviso "
        txtCtaCaptaNroCentral.SetFocusCuenta
        ValidaDatos = False
        Exit Function
    End If
    If Mid(txtCtaCaptaNroCentral.Cuenta, 1, 1) <> Mid(gsOpeCod, 3, 1) Then
        MsgBox "Cuenta no válida para Operacion. Moneda diferente de Operación", vbInformation, "Aviso"
        txtCtaCaptaNroCentral.SetFocusCuenta
        ValidaDatos = False
        Exit Function
    End If
Else
    If Len(txtCtaCaptaNro.NroCuenta) <> 12 Then
        MsgBox "Cuenta de Ahorros no válida", vbInformation, "Aviso "
        txtCtaCaptaNro.SetFocusCuenta
        ValidaDatos = False
        Exit Function
    End If
    If Mid(txtCtaCaptaNro.psCuenta, 1, 1) <> Mid(gsOpeCod, 3, 1) Then
        MsgBox "Cuenta no válida para Operacion. Moneda diferente de Operación", vbInformation, "Aviso"
        txtCtaCaptaNro.SetFocusCuenta
        ValidaDatos = False
        Exit Function
    End If
End If

If txtBuscarPers = "" Or lblPersNombre = "" Then
    MsgBox "Codigo de Persona o Nombre de Titular no válidos ", vbInformation, "Aviso "
    ValidaDatos = False
    Exit Function
End If
If Len(Trim(txtDescripcion)) = 0 Then
    MsgBox "Decripcion no valida", vbInformation, "Aviso "
    ValidaDatos = False
    Exit Function
End If
If Val(txtMonto) = 0 Then
    MsgBox "Monto no válido", vbInformation, "Aviso "
    If txtMonto.Enabled Then txtMonto.SetFocus
    ValidaDatos = False
    Exit Function
End If
If fradoc.Visible Then
    If Len(Trim(txtnroDoc)) = 0 Then
        MsgBox "N° de Documento no ha sido ingresado", vbInformation, "Aviso"
        txtnroDoc.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    Select Case lnTpoDoc
        Case TpoDocOrdenPago
                If ValidaOrdenPagoCuenta(txtCtaCaptaNro.NroCuenta, txtnroDoc, txtMonto) = False Then
                    ValidaDatos = False
                    Exit Function
                End If
    End Select
End If
If framotivo.Visible Then
    If Trim(cboNANC) = "" Then
        MsgBox "Nota de Abono o Cargo no ha sido Ingresada", vbInformation, "aviso"
        cboNANC.SetFocus
        ValidaDatos = False
        Exit Function
    End If
End If

End Function
Private Sub cmdCancelar_Click()
ldFechaNotaCA = ""
lsPersCod = ""
lsPersNombre = ""
lsCuentaAhoNro = ""
lsGlosa = ""
lsNotaCargoAbono = ""
Unload Me
lbOk = False
End Sub

Private Sub cmdNotacargoAbono_Click()
frmIngNotaAbonoCargo.Inicio lnTpoDoc, txtMonto
If frmIngNotaAbonoCargo.pbOk Then
    CargaCombo cboNANC, oDocRec.GetNotasCargoAbonoEst(lnTpoDoc, gNCNARegistrado, Mid(gsOpeCod, 3, 1)), False
    cboNANC.ListIndex = cboNANC.ListCount - 1
    fgTitulares.SetFocus
End If
End Sub
Private Sub fgTitulares_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtDescripcion.SetFocus
End If
End Sub

Private Sub fgTitulares_RowColChange()
If fgTitulares.TextMatrix(1, 1) <> "" Then
    lblPersNombre = fgTitulares.TextMatrix(fgTitulares.Row, 1)
    txtBuscarPers = fgTitulares.TextMatrix(fgTitulares.Row, 8)
    lblPersDireccion = fgTitulares.TextMatrix(fgTitulares.Row, 3)
    lblUbigeo = fgTitulares.TextMatrix(fgTitulares.Row, 4)
    lnPersoneria = fgTitulares.TextMatrix(fgTitulares.Row, 9)
    
End If
End Sub
Private Sub Form_Load()
CentraForm Me
txtFechaNC = ldFecha
txtFechaNC.Enabled = False

Me.Caption = gsOpeDesc
txtMonto = Format(lnMonto, "##,#0.00")
LimpiaControles

If gbBitCentral Then
    txtCtaCaptaNroCentral.Visible = True
    txtCtaCaptaNro.Visible = False
    
    txtCtaCaptaNroCentral.CMAC = gsCodCMAC
    txtCtaCaptaNroCentral.Prod = gCapAhorros
    txtCtaCaptaNroCentral.Age = gsCodAge
    
    txtCtaCaptaNroCentral.EnabledProd = False
    txtCtaCaptaNroCentral.EnabledCMAC = False
Else
    txtCtaCaptaNroCentral.Visible = False
    txtCtaCaptaNro.Visible = True
    If lsCuenta = "" Then
        txtCtaCaptaNro.psAge = gsCodAge
        txtCtaCaptaNro.psProd = gCapAhorros
    Else
        txtCtaCaptaNro.psAge = Left(lsCuenta, 2)
        txtCtaCaptaNro.psProd = Mid(lsCuenta, 3, 3)
        txtCtaCaptaNro.psCuenta = Mid(lsCuenta, 6, 20)
    End If
End If
txtDescripcion = lsGlosa
txtBuscarPers = lsPersCod
lblPersNombre = lsPersNombre

If lnMonto = 0 Or lbModificaMonto Then txtMonto.Locked = False
Set oCapta = New NCapMantenimiento
Set oContFun = New NContFunciones
Set oDoc = New DDocumento
Set oDocRec = New NDocRec
framotivo.Visible = lbMotivo
'fraDoc.Visible = False
Me.txtnroDoc = oContFun.GeneraDocNro(lnTpoDoc, , , , True)
Select Case lnTpoDoc
    Case TpoDocNotaCargo
        fraIngNotaCargo.Caption = "OPERACION CON NOTA DE CARGO"
    Case TpoDocNotaAbono
        fraIngNotaCargo.Caption = "OPERACION CON NOTA DE ABONO"
    Case TpoDocOrdenPago
        fradoc.Visible = True
        fraIngNotaCargo.Caption = "OPERACION CON ORDEN DE PAGO"
        lblDescDoc = "Orden N° :"
End Select
'CargaCombo cboNANC, oDocRec.GetNotasCargoAbonoEst(lnTpoDoc, gNCNARegistrado, Mid(gsOpeCod, 3, 1))
CambiaTamañoCombo cboNANC, 250
End Sub
Public Property Get vbOk() As Boolean
vbOk = lbOk
End Property
Public Property Let vbOk(ByVal vNewValue As Boolean)
lbOk = vNewValue
End Property
Public Property Get NroNotaCA() As String
NroNotaCA = lsNroNotaCA
End Property
Public Property Let NroNotaCA(ByVal vNewValue As String)
NroNotaCA = vNewValue
End Property
Public Property Get FechaNotaCA() As Date
FechaNotaCA = ldFechaNotaCA
End Property
Public Property Let FechaNotaCA(ByVal vNewValue As Date)
ldFechaNotaCA = vNewValue
End Property
Public Property Get PersCod() As String
PersCod = lsPersCod
End Property
Public Property Let PersCod(ByVal vNewValue As String)
lsPersCod = vNewValue
End Property
Public Property Get PersNombre() As String
PersNombre = lsPersNombre
End Property
Public Property Let PersNombre(ByVal vNewValue As String)
lsPersNombre = vNewValue
End Property
Public Property Get PersDireccion() As String
PersDireccion = lsPersDireccion
End Property
Public Property Let PersDireccion(ByVal vNewValue As String)
lsPersDireccion = vNewValue
End Property
Public Property Get PersUbigeo() As String
PersUbigeo = lsPersUbigeo
End Property
Public Property Let PersUbigeo(ByVal vNewValue As String)
lsPersUbigeo = vNewValue
End Property
Public Property Get CuentaAhoNro() As String
CuentaAhoNro = lsCuentaAhoNro
End Property
Public Property Let CuentaAhoNro(ByVal vNewValue As String)
lsCuentaAhoNro = vNewValue
End Property
Public Property Get Glosa() As String
Glosa = lsGlosa
End Property
Public Property Let Glosa(ByVal vNewValue As String)
lsGlosa = vNewValue
End Property
Public Property Get NotaCargoAbono() As String
NotaCargoAbono = lsNotaCargoAbono
End Property
Public Property Let NotaCargoAbono(ByVal vNewValue As String)
lsNotaCargoAbono = vNewValue
End Property

Private Sub Form_Unload(Cancel As Integer)
Set oCapta = Nothing
Set oContFun = Nothing
End Sub
Private Sub lstCuentas_Click()
If lstCuentas.ListCount > 0 Then
    txtCtaCaptaNroCentral.NroCuenta = Trim(lstCuentas)
    CargaDatosPersona lstCuentas
    lstCuentas.SetFocus
End If
End Sub
Private Sub lstCuentas_DblClick()
lstCuentas.Visible = False
fgTitulares.SetFocus
End Sub
Private Sub lstCuentas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lstCuentas.ListCount > 0 Then
        If lstCuentas = "" Then
            MsgBox "Seleccione alguna cuenta por favor", vbInformation, "Aviso"
            Exit Sub
        End If
        txtCtaCaptaNroCentral.NroCuenta = Trim(lstCuentas)
        fgTitulares.SetFocus
    End If
    lstCuentas.Visible = False
End If
End Sub

Private Sub txtBuscarPers_EmiteDatos()
If txtBuscarPers.Text <> "" Then
    
    '*** PEAC 20110405
    If gsOpeCod = "421110" Or gsOpeCod = "422110" Then
        If Trim(txtBuscarPers.Text) <> Trim(lsPersCod) Then
            MsgBox "No se permite cambiar el proveedor.", vbOKOnly + vbInformation, "Atención"
            txtBuscarPers.Text = lsPersCod
            Exit Sub
        End If
    End If
    '*** FIN PEAC
    
    fgTitulares.Clear
    fgTitulares.Rows = 2
    txtCtaCaptaNro.psCuenta = ""
    txtCtaCaptaNroCentral.Cuenta = ""
    CargaCuentas txtBuscarPers.Text
End If
lblPersNombre = txtBuscarPers.psDescripcion
lblPersDireccion = txtBuscarPers.sPersDireccion
lblUbigeo = ""
If lstCuentas.Visible Then
    lstCuentas.SetFocus
End If
End Sub
Sub CargaCuentas(ByVal psPersCod As String)
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set rs = oCapta.GetCuentasPersona(psPersCod, gCapAhorros, True)
lstCuentas.Clear
If Not rs.EOF And Not rs.BOF Then
    lstCuentas.Visible = True
    Do While Not rs.EOF
        If lnTpoDoc = TpoDocOrdenPago Then
            Set rs1 = oCapta.GetDatosCuenta(rs!cCtaCod)
            If Not rs1.EOF And Not rs1.BOF Then
                If rs1!bOrdPag = True Then
                    If Mid(rs!cCtaCod, 1, 1) = Mid(gsOpeCod, 3, 1) Then
                        lstCuentas.AddItem rs!cCtaCod
                    End If
                End If
            End If
            rs1.Close
            Set rs1 = Nothing
        Else
            If Mid(rs!cCtaCod, 9, 1) = Mid(gsOpeCod, 3, 1) Then
                lstCuentas.AddItem rs!cCtaCod
            End If
        End If
        rs.MoveNext
    Loop
    If lstCuentas.ListCount = 0 Then
        lstCuentas.Visible = False
        MsgBox "No se encontraron Cuentas válidas de la Persona Ingresada  para operación", vbInformation, "Aviso"
    End If
Else
    lstCuentas.Visible = False
End If
rs.Close
Set rs = Nothing
End Sub
Sub CargaDatosPersona(ByVal psCodCta As String)
Dim rs As ADODB.Recordset
Dim sSql As String
Set rs = New ADODB.Recordset

If lnTpoDoc = TpoDocOrdenPago Then
    Set rs = oCapta.GetDatosCuenta(psCodCta)
    If Not rs.EOF And Not rs.BOF Then
        If rs!bOrdPag = False Then
            MsgBox "Cuenta Ingresada no válida para operaciones con Ordenes de Pago", vbInformation, "Aviso"
            txtCtaCaptaNro.psCuenta = ""
            txtCtaCaptaNroCentral.Cuenta = ""
            txtCtaCaptaNro.SetFocusCuenta
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    End If
    rs.Close
    Set rs = Nothing
End If

'*****************************
'Nombre, Relacion, Direccion, Zona,
'Fono, ID, [ID N°], P.cPersCod, P.nPersPersoneria,
'I.cPersIDtpo, PP.nPrdPersRelac, PP.cCtaCod
Dim oCon As New DConecta
If Not gbBitCentral Then
    If oCon.AbreConexion Then 'Remota(Left(psCodCta, 2), True, False, "01")
         sSql = "SELECT b.cNomPers Nombre, cRelaCta Relacion, b.cDirPers Direccion, b.cCodZon Zona, cTelPers FONO, ISNULL(cTidotr,cTidoci) ID, ISNULL(cNudoci,cNudotr) [ID N°], a.cCodPers cPersCod, cTipPers nPersPersoneria, ISNULL(cTidotr,cTidoci) cPersIFTpo, cRelaCta nPrdPersRelac, cCodCta cCtaCod " _
            & "FROM   PersCuenta a JOIN DBPersona..persona b ON (b.cCodPers = a.cCodPers) " _
            & "WHERE  a.cCodCta = '" & txtCtaCaptaNro.psAge & txtCtaCaptaNro.psProd & txtCtaCaptaNro.psCuenta & "'"
    
         Set rs = oCon.CargaRecordSet(sSql)
    Else
        Exit Sub
    End If
    
    oCon.CierraConexion
Else
'*** SOLO CENTRALIZADO
    Set rs = oCapta.GetPersonaCuenta(psCodCta)
'***
End If
If Not rs.EOF And Not rs.BOF Then
    txtBuscarPers.Text = gsCodCMAC & rs!cPersCod
    lnPersoneria = rs!nPersPersoneria
    lblPersNombre = rs!Nombre
    lblPersDireccion = rs!Direccion
    lblUbigeo = rs!Zona
    Set fgTitulares.Recordset = rs
    SetupGridCliente
    If txtFechaNC.Enabled Then
        txtFechaNC.SetFocus
    Else
        If txtBuscarPers.Text <> "" Then
            fgTitulares.SetFocus
        Else
            txtBuscarPers.SetFocus
        End If
    End If
Else
    LimpiaControles
    MsgBox "Nro de Cuenta no válida", vbInformation, "Aviso"
    If txtCtaCaptaNro.Visible And txtCtaCaptaNro.Enabled Then
        txtCtaCaptaNro.SetFocus
    Else
        txtCtaCaptaNroCentral.SetFocus
    End If
    Exit Sub
End If
End Sub
Sub LimpiaControles()
'txtBuscarPers.Text = ""
'lblpersNombre = ""
fgTitulares.Clear
fgTitulares.Rows = 2
SetupGridCliente
End Sub
Private Sub txtCtaCaptaNro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtCtaCaptaNro.NroCuenta <> "" Then
        If Mid(txtCtaCaptaNro.psCuenta, 1, 1) <> Mid(gsOpeCod, 3, 1) Then
            MsgBox "Cuenta no válida para Operacion. Moneda diferente de Operación", vbInformation, "Aviso"
            Exit Sub
        End If
        lstCuentas.Visible = False
        CargaDatosPersona txtCtaCaptaNro.NroCuenta
    End If
End If

End Sub


Private Function EsHABERES(ByVal sCta As String) As Boolean
 Dim sSql As String, CONEX As DConecta, RSTEMP As Recordset
 Set CONEX = New DConecta
 sSql = "select valor=count(*) from itfexoneracioncta "
 sSql = sSql & " where nexotpo=3 and cctacod='" & sCta & "'"
 CONEX.AbreConexion
 Set RSTEMP = CONEX.CargaRecordSet(sSql)
   If RSTEMP.State = 1 Then
        If RSTEMP("Valor") > 0 Then
            EsHABERES = True
        Else
            EsHABERES = False
        End If
        
   End If
 CONEX.CierraConexion
 Set CONEX = Nothing
 
 If RSTEMP.State = 1 Then RSTEMP.Close
 Set RSTEMP = Nothing
 
End Function

Private Sub txtCtaCaptaNroCentral_KeyPress(KeyAscii As Integer)
Dim clsMant As NCapMantenimiento, rsCta As Recordset, nestado As Integer
If KeyAscii = 13 Then
    Dim sCta As String
    
    
    sCta = txtCtaCaptaNroCentral.NroCuenta
    
    
    Set clsMant = New NCapMantenimiento
    Set rsCta = New Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        nestado = rsCta("nPrdEstado")
        If nestado <> 1000 Then
           MsgBox "Esta cuenta no se encuentra activa." & vbCrLf & "COMUNICARSE CON EL AREA DE OPERACIONES", vbOKOnly + vbExclamation, "AVISO"
           Exit Sub
        End If
        
    End If
    
    If txtCtaCaptaNroCentral.NroCuenta <> "" Then
        If Mid(txtCtaCaptaNroCentral.Cuenta, 1, 1) <> Mid(gsOpeCod, 3, 1) Then
            MsgBox "Cuenta no válida para Operacion. Moneda diferente de Operación", vbInformation, "Aviso"
            Exit Sub
        End If
        If EsHABERES(sCta) Then
           MsgBox "No puede utilizar esta Operación para una Cuenta de Haberes", vbOKOnly + vbExclamation, App.Title
           Exit Sub
        End If
        
        Dim oITF As COMDConstSistema.FCOMITF
        Set oITF = New COMDConstSistema.FCOMITF
        
        oITF.fgITFParametros
        
        If CtaExoneradaITF(sCta) Then
           Me.lblITFValor.Caption = "0.00"
        Else
            Me.lblITFValor.Caption = Format(oITF.fgITFCalculaImpuesto(Me.txtMonto.Text), "0.00")
            lnITFValor = Format(Me.lblITFValor.Caption, "0.00")
        End If
        
        Set oITF = Nothing
        
        lstCuentas.Visible = False
        CargaDatosPersona txtCtaCaptaNroCentral.NroCuenta
    End If
End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If fradoc.Visible Then
        txtnroDoc.SetFocus
    Else
        txtMonto.SetFocus
    End If
End If
End Sub

Private Sub txtFechaNC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFechaNC) = False Then Exit Sub
    If txtBuscarPers.Text <> "" Then
        fgTitulares.SetFocus
    Else
        txtBuscarPers.SetFocus
    End If
End If
End Sub

Private Sub txtmonto_GotFocus()
fEnfoque txtMonto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMonto, KeyAscii)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub
Private Sub SetupGridCliente()
Dim I As Integer
 
fgTitulares.Cols = 12
For I = 1 To fgTitulares.Cols - 1
    fgTitulares.MergeCol(I) = True
Next I
fgTitulares.MergeCells = flexMergeFree
fgTitulares.ColWidth(0) = 100
fgTitulares.ColWidth(1) = 3500
fgTitulares.ColWidth(2) = 400
fgTitulares.ColWidth(3) = 3500
fgTitulares.ColWidth(4) = 1500
fgTitulares.ColWidth(5) = 1000
fgTitulares.ColWidth(6) = 600
fgTitulares.ColWidth(7) = 1500
fgTitulares.ColWidth(8) = 0
fgTitulares.ColWidth(9) = 0
fgTitulares.ColWidth(10) = 0
fgTitulares.ColWidth(11) = 0

fgTitulares.TextMatrix(0, 1) = "Nombre"
fgTitulares.TextMatrix(0, 2) = "RE"
fgTitulares.TextMatrix(0, 3) = "Direccion"
fgTitulares.TextMatrix(0, 4) = "Zona"
fgTitulares.TextMatrix(0, 5) = "Fono"
fgTitulares.TextMatrix(0, 6) = "ID"
fgTitulares.TextMatrix(0, 7) = "ID N°"
End Sub
Private Sub txtNotaCargo_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    txtFechaNC.SetFocus
End If
End Sub

Private Sub txtMonto_LostFocus()
If Val(txtMonto) = 0 Then txtMonto = 0
txtMonto = Format(txtMonto, "#,#0.00")
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    txtnroDoc = Format(txtnroDoc, "00000000")
    txtMonto.SetFocus
End If
End Sub
Public Property Get Personeria() As PersPersoneria
Personeria = lnPersoneria
End Property
Public Property Let Personeria(ByVal vNewValue As PersPersoneria)
lnPersoneria = vNewValue
End Property
Public Property Get Monto() As Currency
Monto = pnMonto
End Property
Public Property Let Monto(ByVal vNewValue As Currency)
pnMonto = vNewValue
End Property
Public Property Get Motivo() As MotivoNotaAbonoCargo
Motivo = lnMotivo
End Property
Public Property Let Motivo(ByVal vNewValue As MotivoNotaAbonoCargo)
lnMotivo = vNewValue
End Property
Public Property Get ObjetoMotivoPadre() As String
ObjetoMotivoPadre = lsObjetoMotivoPadre
End Property
Public Property Let ObjetoMotivoPadre(ByVal vNewValue As String)
lsObjetoMotivoPadre = vNewValue
End Property
Public Property Get ObjetoMotivo() As String
ObjetoMotivo = lsObjetoMotivo
End Property
Public Property Let ObjetoMotivo(ByVal vNewValue As String)
lsObjetoMotivo = vNewValue
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And txtCtaCaptaNroCentral.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValCtaAhoAnt.Inicia(gCapAhorros, False)
        If sCuenta <> "" Then
            txtCtaCaptaNroCentral.NroCuenta = sCuenta
            txtCtaCaptaNroCentral.SetFocusCuenta
        End If
    End If
End Sub

Private Function CtaExoneradaITF(ByVal sCta As String) As Boolean
 Dim sSql As String, CONEX As DConecta, RSTEMP As Recordset
 Set CONEX = New DConecta
 sSql = "select valor=count(*) from itfexoneracioncta "
 sSql = sSql & " where nexotpo=3 and cctacod='" & sCta & "'"
 CONEX.AbreConexion
 Set RSTEMP = CONEX.CargaRecordSet(sSql)
   If RSTEMP.State = 1 Then
        If RSTEMP("Valor") > 0 Then
            CtaExoneradaITF = True
        Else
            CtaExoneradaITF = False
        End If
        
   End If
 CONEX.CierraConexion
 Set CONEX = Nothing
 
 If RSTEMP.State = 1 Then RSTEMP.Close
 Set RSTEMP = Nothing
 
End Function
