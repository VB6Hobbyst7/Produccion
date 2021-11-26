VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCobroPrimaSegSoat 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8325
   Icon            =   "FrmCobroPrimaSegSoat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   8325
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   3120
      Width           =   1215
   End
   Begin SICMACT.Usuario user 
      Left            =   360
      Top             =   3120
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdgrabar 
      Caption         =   "&Grabar"
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   7815
      Begin VB.ComboBox cmbTipoDoi 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Prima Cobrada"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   4800
         TabIndex        =   17
         Top             =   1680
         Width           =   2415
         Begin VB.TextBox lblmoneda 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   495
         End
         Begin SICMACT.EditMoney lblmonto 
            Height          =   375
            Left            =   840
            TabIndex        =   6
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0.00"
            Enabled         =   -1  'True
         End
      End
      Begin VB.TextBox txtformulario 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cmbUso 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmCobroPrimaSegSoat.frx":030A
         Left            =   1920
         List            =   "FrmCobroPrimaSegSoat.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtdoi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtBuscarUser 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   1210
         Width           =   735
      End
      Begin VB.CommandButton cmbBuscar 
         Caption         =   "Buscar"
         Height          =   360
         Left            =   6960
         TabIndex        =   2
         Top             =   470
         Width           =   735
      End
      Begin VB.TextBox txtcliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         TabIndex        =   1
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo DOI:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo de Uso"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Nro. Formulario"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Shape shpUsuario 
         Height          =   345
         Left            =   240
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblDescUser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1080
         TabIndex        =   14
         Top             =   1200
         Width           =   6615
      End
      Begin VB.Label lsPersCod 
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "DOI                           Cliente"
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   5295
      End
   End
   Begin MSComctlLib.TabStrip tabstrip 
      Height          =   3615
      Left            =   75
      TabIndex        =   9
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6376
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Datos del Cobro"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
End
Attribute VB_Name = "frmCobroPrimaSegSoat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim nProducto As COMDConstantes.Producto
Dim sOperacion As String

Dim oConstante As New COMDConstantes.DCOMConstantes
Dim lrsTipoDoi As New ADODB.Recordset

Dim oSegSoat As New COMNCaptaGenerales.NCOMSeguros
Dim rsDatos As New ADODB.Recordset

Private Sub cmbBuscar_Click()
Dim ClsPersona As New COMDPersona.DCOMPersonas
Dim oPersona As COMDPersona.UCOMPersona
Dim lsPersCod As String

Set ClsPersona = New COMDPersona.DCOMPersonas
Set oPersona = frmBuscaPersona.Inicio 'TxtBuscarPersona.psCodigoPersona

If Not oPersona Is Nothing Then
    lsPersCod = oPersona.sPersCod
    cmbTipoDoi.Clear
    cmbTipoDoi.AddItem Trim(oSegSoat.BuscaTpoDoc(lsPersCod))
    cmbTipoDoi.ListIndex = 0
    cmbTipoDoi.Enabled = False
    txtdoi.Text = oPersona.sPersIdnroDNI ' oPersona.sPersIdnroDNI
    txtdoi.Enabled = False
    txtcliente.Text = oPersona.sPersNombre
    txtcliente.Enabled = False
    txtformulario.SetFocus
Else
    lsPersCod = ""
    txtcliente.Text = ""
    txtdoi.Text = ""
    txtdoi.SetFocus
End If

End Sub

Private Sub cmdcancelar_Click()
 Unload Me
End Sub

Private Sub cmdgrabar_Click()

Dim bExisteF As Integer


   'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    
    If cmbTipoDoi.ListIndex = -1 Or Len(cmbTipoDoi.Text) = 0 Then cmbTipoDoi.SetFocus: MsgBox "Debe Seleccionar el Tipo de Documento", vbInformation, "Aviso": Exit Sub
    If txtdoi.Text = "" Then txtdoi.SetFocus: MsgBox "Debe Registrar el Documento del Cliente", vbInformation, "Aviso": Exit Sub
    If txtcliente = "" Then txtcliente.SetFocus: MsgBox "Debe Registrar los Datos del Cliente", vbInformation, "Aviso": Exit Sub
    If txtformulario = "" Then
     txtformulario = "": txtformulario.SetFocus: MsgBox "Debe Registrar el Número del Formulario", vbInformation, "Aviso": Exit Sub
    Else
      If CDbl(txtformulario) = 0 Then
        txtformulario = "": txtformulario.SetFocus: MsgBox "Debe Registrar el Número del Formulario", vbInformation, "Aviso": Exit Sub
      End If
      
    End If
    
    If txtBuscarUser = "" Then txtBuscarUser.SetFocus: MsgBox "Debe Registrar el Usuario", vbInformation, "Aviso": Exit Sub
    If CDbl(lblmonto) <= 0 Then lblmonto.SetFocus: MsgBox "El monto de la Prima no puede ser S/ 0.00 (Cero)", vbInformation, "Aviso": Exit Sub
    
    bExisteF = oSegSoat.BuscaFormularioExiste(txtformulario)
    
    If bExisteF = 1 Then MsgBox "El Formulario ya fue Usado, Ingrese otro formulario", vbInformation, "Aviso": txtformulario.SetFocus: Exit Sub
  
  
    Call Grabar
    
 Set oSegSoat = Nothing
 Set rsDatos = Nothing
End Sub
Private Sub Grabar()

Dim clsCapMov As New COMNCaptaGenerales.NCOMCaptaMovimiento
Dim clsCont As New COMNContabilidad.NCOMContFunciones
Dim clsCap As New COMDCaptaGenerales.DCOMCaptaMovimiento
Dim oSeg As New COMNCaptaGenerales.NCOMSeguros
Dim lsMov As String
Dim lnMovNro As Long
Dim lsBoleta As String

Dim lsDocumento As String

If MsgBox("La Prima Cobrar es: " & lblmoneda.Text & " " & lblmonto.Text, vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
If MsgBox("Desea Grabar la Información", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub

Dim lnMonto As Currency
'Dim lsDocumento As String
Dim lsNombAfiliado As String
Dim lsTipoUso As String
Dim lsUsuaAfilia As String
Dim lsDoi As String
'Dim lnMonto As Currency

On Error GoTo Error

lsDocumento = txtformulario.Text
lsNombAfiliado = txtcliente.Text
lsTipoUso = cmbUso.Text
lsUsuaAfilia = txtBuscarUser.Text
lnMonto = lblmonto.Text
lsDoi = txtdoi.Text

lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
lnMovNro = clsCapMov.OtrasOperaciones(lsMov, gServCobPriSegSoat, lnMonto, lsDocumento, "Cobro Prima Seguro Soat", gMonedaNacional, lsPersCod, , , , , , , gnMovNro)
Call clsCap.AgregaSegSoat(lsDocumento, lsNombAfiliado, lsPersCod, lsTipoUso, lsUsuaAfilia, lnMonto, lnMovNro, gsCodAge, 501, gdFecSis, lsDoi)

If gnMovNro = 0 Then
    MsgBox "La operación no se realizó, favor intentar nuevamente", vbInformation, "Aviso"
    Exit Sub
End If

Set clsCont = Nothing
Set clsCont = Nothing
Set clsCapMov = Nothing

MsgBox "Debe realizar el cobro de " & lblmoneda.Text & " " & Format(CDbl(lblmonto.Text), gsFormatoNumeroView), vbInformation, "Aviso"

    lsBoleta = oSeg.ImprimeBoletaAfiliacionSegSoat(lnMovNro, lsMov, gsNomAge, gbImpTMU)

        Do
           If Trim(lsBoleta) <> "" Then
                lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    Print #nFicSal, ""
                Close #nFicSal
          End If
        Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
        'Set oBol = Nothing

Call cmdLimpiar_Click

    Exit Sub
Error:
      MsgBox Str(err.Number) & err.Description
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdLimpiar_Click()
cmbTipoDoi.Enabled = True
cmbTipoDoi.Clear
Set lrsTipoDoi = oConstante.RecuperaDoi(gPersIdTipo)
Call Llenar_Combo_con_Recordset(lrsTipoDoi, cmbTipoDoi)

cmbTipoDoi.Enabled = True
txtcliente.Text = ""
txtcliente.Enabled = True

txtdoi.Text = ""
txtdoi.Enabled = True

user.Inicio gsCodUser
cmbUso.ListIndex = 0
txtBuscarUser.Text = gsCodUser
lblDescUser = user.UserNom

 lblmoneda = gcPEN_SIMBOLO
 lblmonto.Text = "0.00"
 lblmonto.MarcaTexto
 
 txtformulario.Text = ""
 


End Sub

'gdFecSis , nOperacion, gsCodAge, gsSimbolo, gcPEN_SIMBOLO, gsUser, gsCodUser
Private Sub Form_Load()
shpUsuario.Visible = True
shpUsuario.BorderWidth = 3
shpUsuario.Move txtBuscarUser.Left, txtBuscarUser.Top, txtBuscarUser.Width, txtBuscarUser.Height
shpUsuario.ZOrder 1

'Tipo Doi
Set lrsTipoDoi = oConstante.RecuperaDoi(gPersIdTipo)
Call Llenar_Combo_con_Recordset(lrsTipoDoi, cmbTipoDoi)
End Sub

Public Sub inicia(ByVal nOpe As CaptacOperacion)
  nOperacion = nOpe
  sOperacion = "Cobro Prima Seguro SOAT"
  Me.Caption = "Cobranza Servicios - " & sOperacion & " - " & nOperacion
  
  user.Inicio gsCodUser
  cmbUso.ListIndex = 0
  txtBuscarUser.Text = gsCodUser
  lblDescUser = user.UserNom

  lblmoneda = gcPEN_SIMBOLO
  lblmonto.Text = "0.00"
  lblmonto.MarcaTexto
  'lblmonto.
  
  Me.Show 1
End Sub
Private Sub LimpiarFormulario()
 Call cmdLimpiar_Click

End Sub


Sub cargaUsuarios()
user.Inicio txtBuscarUser
lblDescUser = user.UserNom
'txtcliente.SetFocus
End Sub

Private Sub lblmonto_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   cmdgrabar.SetFocus
 End If
End Sub

Private Sub txtBuscarUser_Change()

txtBuscarUser = UCase(txtBuscarUser)
txtBuscarUser.SelStart = Len(txtBuscarUser.Text)
  cargaUsuarios
End Sub

Private Sub txtBuscarUser_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   cargaUsuarios
 End If
End Sub

Private Sub txtcliente_Change()
txtcliente = UCase(txtcliente)
txtcliente.SelStart = Len(txtcliente.Text)
End Sub

Private Sub txtdoi_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
If KeyAscii = 13 And txtdoi <> "" Then
  
   Call verificaDigitosDNIRUC(txtdoi)
  
End If
End Sub

Private Sub txtdoi_LostFocus()
If txtdoi <> "" Then
    Call verificaDigitosDNIRUC(txtdoi)
End If
End Sub
Sub verificaDigitosDNIRUC(ByVal doi As String)

If Len(doi) < txtdoi.MaxLength Then
    MsgBox "Por favor verifique su número de documento", vbInformation, "Aviso": txtdoi.SetFocus
    Exit Sub
End If

' Select Case Len(doi)
'  Case Is = 0
'     cmdcancelar.SetFocus
'  Case Is = 1
'     txtdoi.SetFocus
'     Exit Sub
'  Case Is < 8
'        MsgBox "El DNI debe tener 8 Dígitos", vbInformation, "Aviso": txtdoi.SetFocus
'        Exit Sub
'  Case 9 To 10
'        MsgBox "El RUC debe tener 11 digitos", vbInformation, "Aviso": txtdoi.SetFocus
'End Select
End Sub
Private Sub txtformulario_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
End Sub

'------
'Funcion que llena un Combo con un recordset
Sub Llenar_Combo_con_Recordset(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
On Error Resume Next
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cConsDescripcion) & space(100) & Trim((pRs!nConsValor))
    pRs.MoveNext
Loop
pRs.Close
    
End Sub

Private Sub cmbTipoDoi_Click()
    If Trim(Right(cmbTipoDoi.Text, 2)) = "1" Then txtdoi.MaxLength = 8: txtdoi.Text = "": txtdoi.SetFocus
    If Trim(Right(cmbTipoDoi.Text, 2)) = "2" Then txtdoi.MaxLength = 11: txtdoi.Text = "": txtdoi.SetFocus
    If Trim(Right(cmbTipoDoi.Text, 2)) = "4" Then txtdoi.MaxLength = 9: txtdoi.Text = "": txtdoi.SetFocus
End Sub
Private Sub cmbTipoDoi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Right(cmbTipoDoi.Text, 2)) = "1" Then txtdoi.MaxLength = 8: txtdoi.Text = "": txtdoi.SetFocus
        If Trim(Right(cmbTipoDoi.Text, 2)) = "2" Then txtdoi.MaxLength = 11: txtdoi.Text = "": txtdoi.SetFocus
        If Trim(Right(cmbTipoDoi.Text, 2)) = "4" Then txtdoi.MaxLength = 9: txtdoi.Text = "": txtdoi.SetFocus
    End If
End Sub
