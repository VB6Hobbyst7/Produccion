VERSION 5.00
Begin VB.Form frmMntParaViaticosDat 
   Caption         =   "Parámetros de Viáticos"
   ClientHeight    =   2985
   ClientLeft      =   2850
   ClientTop       =   4665
   ClientWidth     =   6945
   Icon            =   "frmMntParaViaticosDat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4140
      TabIndex        =   4
      Top             =   2490
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   2490
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   120
      TabIndex        =   6
      Top             =   90
      Width           =   6675
      Begin VB.ComboBox cboConcepto 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1770
         Width           =   2715
      End
      Begin VB.TextBox txtTranspDesc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1410
         TabIndex        =   19
         Top             =   1290
         Width           =   2295
      End
      Begin VB.TextBox txtDestinoDesc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1410
         TabIndex        =   18
         Top             =   810
         Width           =   2295
      End
      Begin VB.TextBox txtCategoriaDesc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1410
         TabIndex        =   17
         Top             =   330
         Width           =   2295
      End
      Begin VB.TextBox txtTranspCod 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   16
         Top             =   1290
         Width           =   345
      End
      Begin VB.TextBox txtDestinoCod 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   15
         Top             =   810
         Width           =   345
      End
      Begin VB.TextBox txtCategoriaCod 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   14
         Top             =   330
         Width           =   345
      End
      Begin VB.TextBox txtConcepImporte 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   4890
         TabIndex        =   3
         Top             =   1740
         Width           =   1575
      End
      Begin VB.ComboBox cboAfectoTope 
         Height          =   315
         ItemData        =   "frmMntParaViaticosDat.frx":030A
         Left            =   5580
         List            =   "frmMntParaViaticosDat.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   780
         Width           =   915
      End
      Begin VB.ComboBox cboAfectoA 
         Height          =   315
         ItemData        =   "frmMntParaViaticosDat.frx":0356
         Left            =   5040
         List            =   "frmMntParaViaticosDat.frx":0363
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Importe"
         Height          =   255
         Left            =   4050
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Afecto a"
         Height          =   255
         Left            =   3990
         TabIndex        =   12
         Top             =   420
         Width           =   1545
      End
      Begin VB.Label Label5 
         Caption         =   "Afecto a Tope"
         Height          =   255
         Left            =   3990
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Concepto"
         Height          =   285
         Left            =   150
         TabIndex        =   10
         Top             =   1830
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Transporte"
         Height          =   285
         Left            =   150
         TabIndex        =   9
         Top             =   1350
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Destino"
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   870
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Categoría"
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   420
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMntParaViaticosDat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lNuevo As Boolean
Dim rs     As New ADODB.Recordset
Dim sCategCod     As String
Dim sDestinoCod   As String
Dim sTranspCod    As String
Dim sObjetoCod    As String
Dim sViaticoAfectoA    As String
Dim sViaticoAfectoTope As String
Dim nViaticoImporte     As Currency

Dim clsV As DParaViaticos
'ARLO2010208****
Dim objPista As COMManejador.Pista
Dim lsPalabra, lsAccion As String
'************

Public Sub Inicio(psCategoriaCod As String, psDestinoCod As String, psTranspCod As String, psObjetoCod As String, plNuevo As Boolean)
lNuevo = plNuevo
sCategCod = psCategoriaCod
sDestinoCod = psDestinoCod
sTranspCod = psTranspCod
sObjetoCod = psObjetoCod
Me.Show 1
End Sub

Private Sub cboAfectoA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboAfectoTope.SetFocus
End If
End Sub

Private Sub cboAfectoTope_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtConcepImporte.SetFocus
End If
End Sub

Private Sub cboConcepto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboAfectoA.SetFocus
End If
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo AceptarErr
If cboConcepto.ListIndex = -1 Then
   MsgBox "Necesita seleccionar Concepto de Viático", vbInformation, "¡Aviso!"""
   cboConcepto.SetFocus
   Exit Sub
End If
If cboAfectoA.ListIndex = -1 Then
   MsgBox "Necesita definir Afectación al Nro. de Días de ConceptoImporte", vbInformation, "Error"
   cboAfectoA.SetFocus
   Exit Sub
End If
If cboAfectoTope.ListIndex = -1 Then
   MsgBox "Necesita definir Afectación al Monto Tope de ConceptoImporte", vbInformation, "Error"
   cboAfectoTope.SetFocus
   Exit Sub
End If
If txtConcepImporte = "" Then
   MsgBox "Falta definir Monto de Concepto", vbInformation, "Error"
   txtConcepImporte.SetFocus
   Exit Sub
End If


If MsgBox(" ¿ Seguro de Grabar Concepto de Viaticos ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   sCategCod = txtCategoriaCod
   sDestinoCod = txtDestinoCod
   sTranspCod = txtTranspCod
   sObjetoCod = Trim(Right(cboConcepto.Text, 100))
   sViaticoAfectoA = Trim(Right(cboAfectoA.Text, 100))
   sViaticoAfectoTope = Trim(Mid(cboAfectoTope.Text, 1, 1))
   nViaticoImporte = CCur(Format(txtConcepImporte, gsFormatoNumeroDato))
   
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   If lNuevo Then
      clsV.InsertaParametros sCategCod, sDestinoCod, sTranspCod, sObjetoCod, sViaticoAfectoA, sViaticoAfectoTope, nViaticoImporte, gsMovNro
   Else
      clsV.ActualizaParametros sCategCod, sDestinoCod, sTranspCod, sObjetoCod, sViaticoAfectoA, sViaticoAfectoTope, nViaticoImporte, gsMovNro
   End If
   
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            If (lNuevo) Then
            lsPalabra = "Creo Nuevo"
            lsAccion = "1"
            Else: lsPalabra = "Modifico"
            lsAccion = "2"
            End If
            gsOpeCod = LogPistaParaMantViatico
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, "Se " & lsPalabra & " Parametro de Viatico de la Categoria : " & txtCategoriaDesc.Text & " | Concepto : " & Left(cboConcepto.Text, 15) & " | El Importe : " & txtConcepImporte.Text
            Set objPista = Nothing
            '*******
   glAceptar = True
   Unload Me
End If
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdSalir_Click()
glAceptar = False
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Conceptos de Viáticos: " & IIf(lNuevo, "Nuevo", "Modificación")

Dim clsC As New DConstantes
Set rs = clsC.CargaConstante(gViaticosCateg, sCategCod)
txtCategoriaDesc = rs!cConsDescripcion
Set rs = clsC.CargaConstante(gViaticosDestino, sDestinoCod)
txtDestinoDesc = rs!cConsDescripcion
Set rs = clsC.CargaConstante(gViaticosTransporte, sTranspCod)
txtTranspDesc = rs!cConsDescripcion
Set clsC = Nothing

CentraForm Me

txtCategoriaCod = sCategCod
txtDestinoCod = sDestinoCod
txtTranspCod = sTranspCod

LlenaComboConstante gViaticosAfectoA, cboAfectoA

Dim clsObj As DObjeto
Set clsObj = New DObjeto
Set rs = clsObj.CargaObjeto(ObjConceptosARendir & "_", True)
RSLlenaCombo rs, cboConcepto
RSClose rs
Set clsObj = Nothing

Set clsV = New DParaViaticos

If Not lNuevo Then
   Set rs = clsV.CargaParametros(sCategCod, sDestinoCod, sTranspCod, sObjetoCod)
   If Not rs.EOF Then
      sViaticoAfectoA = rs!cViaticoAfectoA
      sViaticoAfectoTope = rs!cViaticoAfectoTope
      nViaticoImporte = rs!nViaticoImporte
   End If
   RSClose rs
   cboAfectoA.ListIndex = BuscaCombo(sViaticoAfectoA, cboAfectoA)
   cboAfectoTope.ListIndex = BuscaCombo(sViaticoAfectoTope, cboAfectoTope)
   cboConcepto.ListIndex = BuscaCombo(sObjetoCod, cboConcepto)
   txtConcepImporte = Format(nViaticoImporte, gsFormatoNumeroView)
   cboConcepto.Enabled = False
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rs
Set clsV = Nothing
End Sub

Private Sub txtConcepImporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtConcepImporte = Format(txtConcepImporte, gsFormatoNumeroView)
   cmdAceptar.SetFocus
End If
End Sub

Public Property Get psCodigo() As String
psCodigo = sObjetoCod
End Property

Public Property Let psCodigo(ByVal vNewValue As String)
sObjetoCod = vNewValue
End Property
