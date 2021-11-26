VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5F774E03-DB36-4DFC-AAC4-D35DC9379F2F}#1.1#0"; "VertMenu.ocx"
Begin VB.Form frmOpePagSunat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pago de Impuestos"
   ClientHeight    =   4950
   ClientLeft      =   1350
   ClientTop       =   1395
   ClientWidth     =   9675
   Icon            =   "frmOpePagSUNAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame fraEntidad 
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
      ForeColor       =   &H8000000D&
      Height          =   660
      Left            =   1290
      TabIndex        =   18
      Top             =   4140
      Width           =   5970
      Begin Sicmact.TxtBuscar txtBuscaEntidad 
         Height          =   360
         Left            =   105
         TabIndex        =   6
         Top             =   210
         Width           =   1875
         _ExtentX        =   3307
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
      End
      Begin VB.Label lblCtaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1995
         TabIndex        =   7
         Top             =   210
         Width           =   3825
      End
   End
   Begin Sicmact.FlexEdit fgDetalle 
      Height          =   2025
      Left            =   1290
      TabIndex        =   5
      Top             =   1770
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   3572
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Nro-Cuenta-Descripcion-Monto-DH"
      EncabezadosAnchos=   "400-1900-4300-1300-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-X-3-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-1-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-R-L"
      FormatosEdit    =   "0-0-0-2-0"
      AvanceCeldas    =   1
      TextArray0      =   "Nro"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame frameDoc 
      Caption         =   "&Formulario"
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
      Height          =   675
      Left            =   5040
      TabIndex        =   16
      Top             =   60
      Width           =   4515
      Begin VB.TextBox txtDocNro 
         Height          =   315
         Left            =   810
         TabIndex        =   2
         Top             =   240
         Width           =   1605
      End
      Begin MSMask.MaskEdBox txtFechaFormula 
         Height          =   330
         Left            =   3225
         TabIndex        =   3
         Top             =   240
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   2550
         TabIndex        =   19
         Top             =   315
         Width           =   555
      End
      Begin VB.Label lblDocNro 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.TextBox txtDebe 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
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
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   7560
      TabIndex        =   14
      Top             =   3825
      Width           =   1680
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8490
      TabIndex        =   9
      Top             =   4320
      Width           =   1065
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
      Height          =   675
      Left            =   1290
      TabIndex        =   11
      Top             =   30
      Width           =   2985
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   330
         Left            =   1725
         TabIndex        =   1
         Top             =   232
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
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
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   1170
         TabIndex        =   13
         Top             =   277
         Width           =   555
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
      Height          =   915
      Left            =   1290
      TabIndex        =   12
      Top             =   780
      Width           =   8265
      Begin VB.TextBox txtMovDesc 
         Height          =   555
         Left            =   180
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   7920
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7380
      TabIndex        =   8
      Top             =   4320
      Width           =   1065
   End
   Begin VertMenu.VerticalMenu vFormPago 
      Height          =   4695
      Left            =   90
      TabIndex        =   10
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   8281
      MenuCaption1    =   "Forma Pago"
      MenuItemsMax1   =   3
      MenuItemIcon11  =   "frmOpePagSUNAT.frx":08CA
      MenuItemCaption11=   "Efectivo"
      MenuItemIcon12  =   "frmOpePagSUNAT.frx":0BE4
      MenuItemCaption12=   "Carta"
      MenuItemIcon13  =   "frmOpePagSUNAT.frx":0EFE
      MenuItemCaption13=   "Cheque"
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6420
      TabIndex        =   15
      Top             =   3870
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   315
      Left            =   6210
      Top             =   3810
      Width           =   3045
   End
End
Attribute VB_Name = "frmOpePagSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim lTransActiva As Boolean      ' Controla si la transaccion esta activa o no
Dim rs As New ADODB.Recordset    'Rs temporal para lectura de datos

Dim lMN As Boolean
Dim lSalir As Boolean, OK As Boolean
Dim lsDocTpo As String

Dim oNContFunc As NContFunciones
Dim lsFileCarta As String

Private Function ValidaInterfaz() As Boolean
Dim nItem      As Integer, nCtas As Integer, N As Integer
Dim sOpeCod    As String, sCTAS() As String
Dim sTexto     As String, sCheque As String
Dim sAsiento   As String, sAsientoD As String
Dim sCtaCod    As String, sCtaDes  As String, sOpeCtaDH As String
Dim sDocVoucher As String
Dim nImporteD  As Currency
Dim sFile      As String, sDocAbrev As String
Dim lOk        As Boolean
Dim lsMovNro   As String
Dim oCon       As New NContFunciones

ValidaInterfaz = False
If Val(Format(txtDebe, gsFormatoNumeroDato)) = 0 Then
   MsgBox " Monto a pagar debe ser Mayor que Cero...! ", vbInformation, "¡Aviso!"
   Exit Function
End If
If lsDocTpo <> "-1" Then
   If txtBuscaEntidad = "" Then
      MsgBox "Falta definir Entidad Bancaria...", vbInformation, "¡Aviso!"
      Exit Function
   End If
End If
If txtMovDesc = "" Then
   MsgBox "Falta ingresar Glosa de Operación!", vbInformation, "¡Aviso!"
   Exit Function
End If

lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
If PermiteModificarAsiento(lsMovNro, , gdFecSis) = False Then
   ValidaInterfaz = False
   Exit Function
End If

ValidaInterfaz = True
End Function

Private Sub cmdAceptar_Click()
Dim lsEntidadOrig As String
Dim lsCtaEntidadOrig As String
Dim lsGlosa As String
Dim lnImporte As Currency
Dim lsSubCuentaIF As String
Dim lsPersCod As String
Dim lsMovNro As String
Dim lsDocumento As String
Dim lsOpeCod As String
Dim lsCtaBanco As String
Dim lsCtaContDebe As String
Dim lsCtaContHaber As String
Dim lsPersCodIf  As String
Dim lsPersNombre As String
Dim lnTrans As Integer
Dim rsBillete As ADODB.Recordset
Dim rsMoneda  As ADODB.Recordset
Dim rsDebe    As ADODB.Recordset
Dim lsRecibo As String
Dim lbEfectivo As Boolean
Dim lsTpoIf As String
Dim lsDocVoucher As String
Dim lsDocNRo     As String
Dim lsFecha      As String
Dim nItem        As Integer


Dim oDocPago As clsDocPago
Dim oCtasIF  As NCajaCtaIF
Dim oOpe     As DOperacion
Dim oCaja    As nCajaGeneral

On Error GoTo GrabarPagoErr

If ValidaInterfaz = False Then Exit Sub

Set oCtasIF = New NCajaCtaIF
Set oOpe = New DOperacion
Set oCaja = New nCajaGeneral
Set oDocPago = New clsDocPago

lsTpoIf = Mid(txtBuscaEntidad, 1, 2)
lsCtaBanco = Mid(txtBuscaEntidad, 18, Len(txtBuscaEntidad))
lsPersCodIf = Mid(txtBuscaEntidad, 4, 13)
lsEntidadOrig = oCtasIF.NombreIF(lsPersCodIf)
lsSubCuentaIF = oCtasIF.SubCuentaIF(lsPersCodIf)
lsCtaEntidadOrig = Trim(lblCtaDesc)
lsGlosa = Trim(txtMovDesc)
lsPersNombre = "SUNAT/BANCO DE LA NACION"

If lsDocTpo = "" Then
    MsgBox "Seleccione Forma de Pago del Impuesto", vbInformation, "Aviso"
    Exit Sub
End If

lnImporte = nVal(txtDebe)
lsDocVoucher = ""
lsDocNRo = ""
lsRecibo = ""
lbEfectivo = False
lsOpeCod = gsOpeCod
If lsDocTpo = "-1" Then
   frmCajaGenEfectivo.Inicio lsOpeCod, gsOpeDesc, lnImporte, Mid(gsOpeCod, 3, 1), False
   If frmCajaGenEfectivo.lbOk Then
        Set rsBillete = frmCajaGenEfectivo.rsBilletes
        Set rsMoneda = frmCajaGenEfectivo.rsMonedas
   Else
       Set frmCajaGenEfectivo = Nothing
       Exit Sub
   End If
   Set frmCajaGenEfectivo = Nothing
   If rsBillete Is Nothing And rsMoneda Is Nothing Then
       MsgBox "Error en Ingreso de Billetaje", vbInformation, "Aviso"
       Exit Sub
   End If
   lbEfectivo = True
   lsCtaContHaber = oOpe.EmiteOpeCta(lsOpeCod, "H", "1", txtBuscaEntidad, ObjEntidadesFinancieras)
Else
    If lsDocTpo = TpoDocCheque Then
        lsDocVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1), gsCodAge)
        oDocPago.InicioCheque lsDocNRo, True, lsPersCodIf, gsOpeCod, lsPersNombre, gsOpeDesc, lsGlosa, lnImporte, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, False, gsCodAge
    End If
    If lsDocTpo = TpoDocCarta Then
       oDocPago.InicioCarta lsDocNRo, lsPersCod, gsOpeCod, gsOpeDesc, lsGlosa, lsFileCarta, lnImporte, gdFecSis, lsEntidadOrig, lsCtaEntidadOrig, lsPersNombre, "", lsMovNro
    End If
    If oDocPago.vbOk Then    'Se ingresó dato de Cheque u Orden de Pago
       lsFecha = oDocPago.vdFechaDoc
       lsDocNRo = oDocPago.vsNroDoc
       lsDocVoucher = oDocPago.vsNroVoucher
       lsDocumento = oDocPago.vsFormaDoc
    Else
        Exit Sub
    End If
   lsCtaContHaber = oOpe.EmiteOpeCta(lsOpeCod, "H", "2", txtBuscaEntidad, ObjEntidadesFinancieras)
End If
If lsCtaContHaber = "" Then
    MsgBox "Cuentas Contables no determinadas correctamente en Operación." & oImpresora.gPrnSaltoLinea & "consulte con Sistemas", vbInformation, "Aviso"
    Exit Sub
End If
If MsgBox("Desea Grabar la Información", vbYesNo + vbQuestion, "Aviso") = vbYes Then
   lsMovNro = oNContFunc.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
   Set rsDebe = fgDetalle.GetRsNew
   If oCaja.GrabaPagoSunat(lsMovNro, lsOpeCod, txtMovDesc, Val(lsDocTpo), lsDocNRo, Format(lsFecha, gsFormatoFecha), rsBillete, rsMoneda, rsDebe, _
                           lsCtaContHaber, lnImporte, lsPersCodIf, lsTpoIf, lsCtaBanco, TpoDocFormSUNAT, _
                           txtDocNro, CDate(txtFechaFormula), lsDocVoucher) = 0 Then
      ImprimeAsientoContable lsMovNro, lsDocVoucher, lsDocTpo, lsDocumento, lbEfectivo, _
                              False, txtMovDesc, lsPersCodIf, lnImporte
      OK = True
      If MsgBox(" ¿ Desea registrar otra Operación ? ", vbQuestion + vbYesNo, "¡Consulta!") = vbYes Then
         For nItem = 1 To fgDetalle.Rows - 1
            fgDetalle.TextMatrix(nItem, 3) = ""
         Next
         txtDebe.Text = "0.00"
         txtMovDesc = ""
         txtDocNro = ""
         lsDocumento = ""
         txtBuscaEntidad = ""
         lblCtaDesc = ""
      Else
         Unload Me
      End If
   End If
End If
   
Exit Sub
GrabarPagoErr:
  MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub


Private Sub Sumas()
Dim nFilas As Integer
Dim nDebe As Currency, nHaber As Currency
Dim I As Integer
nFilas = fgDetalle.Rows - 1
For I = 1 To nFilas
   nDebe = nDebe + nVal(fgDetalle.TextMatrix(I, 3))
Next
txtDebe = Format(nDebe, gsFormatoNumeroView)
fgDetalle.SetFocus
End Sub

Private Sub cmdSalir_Click()
OK = False
Unload Me
End Sub

Private Sub fgDetalle_OnCellChange(pnRow As Long, pnCol As Long)
Sumas
End Sub

Private Sub Form_Activate()
If lSalir Then
   Set rs = Nothing
   Unload Me
End If
End Sub

Private Sub Form_Load()
Set oNContFunc = New NContFunciones

CentraForm Me
Me.Caption = gsOpeDesc
lSalir = False

If Mid(gsOpeCod, 3, 1) = "2" Then  'Identificación de Tipo de Moneda
   MsgBox "Operación no Implementada para Moneda Extranjera", vbInformation, "Error"
   lSalir = True
   Exit Sub
End If
gsSimbolo = gcMN

' Defino el Nro de Movimiento
txtOpeCod = gsOpeCod
txtFecha = gdFecSis
txtFechaFormula = gdFecSis

CargaFlex gsOpeCod

lsFileCarta = App.path & gsDirPlantillas & gsOpeCod & ".TXT"
txtBuscaEntidad.psRaiz = "Cuentas de Instituciones Financieras"

Dim oOpe As DOperacion
Set oOpe = New DOperacion
txtBuscaEntidad.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")  '  oDCtaIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), gTpoIFBanco + gTpoCtaIFCtaCte + gTpoCtaIFCtaAho)
Set oOpe = Nothing
txtDebe = Format(0, "0.00")
lsDocTpo = "-1"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not OK And Not lSalir Then
   If MsgBox(" ¿ Seguro de Salir sin grabar Operación ? ", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
      Cancel = True
   End If
End If
End Sub

Private Sub CargaFlex(psOpeCod As String)
Dim oOpe As New DOperacion
Dim nItem As Integer
Set rs = oOpe.CargaOpeCtaArbol(psOpeCod, "D", "0")
Do While Not rs.EOF
   fgDetalle.AdicionaFila
   nItem = fgDetalle.Row
   fgDetalle.TextMatrix(nItem, 1) = rs!cCtaContCod
   fgDetalle.TextMatrix(nItem, 2) = rs!cCtaContDesc
   fgDetalle.TextMatrix(nItem, 4) = "D"
   fgDetalle.Col = 1
   fgDetalle.CellBackColor = "&H00DBDBDB"
   fgDetalle.Col = 2
   fgDetalle.CellBackColor = "&H00DBDBDB"
   rs.MoveNext
Loop
fgDetalle.Col = 3
fgDetalle.Row = 1
RSClose rs
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oNContFunc = Nothing
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
Dim oNCtasIf As NCajaCtaIF
Set oNCtasIf = New NCajaCtaIF
If txtBuscaEntidad.Text <> "" Then
    lblCtaDesc = oNCtasIf.EmiteTipoCuentaIF(Mid(Me.txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))) & " " & txtBuscaEntidad.psDescripcion
    Set oNCtasIf = Nothing
    cmdAceptar.Enabled = True
    cmdAceptar.SetFocus
    cmdAceptar_Click
End If
End Sub

Private Sub txtDocNro_GotFocus()
fEnfoque txtDocNro
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFecha) = False Then
        Exit Sub
    Else
      txtDocNro.SetFocus
    End If
End If
End Sub

Private Sub txtFechaFormula_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFechaFormula) = False Then
        Exit Sub
    Else
      txtMovDesc.SetFocus
    End If
End If
End Sub

Private Sub txtMovDesc_GotFocus()
fEnfoque txtMovDesc
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   fgDetalle.SetFocus
End If
End Sub

Private Sub txtDocNro_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   txtDocNro = Format(txtDocNro, "00000000")
   txtFechaFormula.SetFocus
End If
End Sub

Private Sub vFormPago_MenuItemClick(MenuNumber As Long, MenuItem As Long)
Select Case MenuItem
    Case 1: lsDocTpo = "-1"
    Case 2: lsDocTpo = TpoDocCarta
   Case 3:  lsDocTpo = TpoDocCheque
End Select
cmdAceptar.SetFocus
cmdAceptar_Click
End Sub

