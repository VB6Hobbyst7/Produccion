VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRegVentaDet 
   Caption         =   "Registro de Ventas: "
   ClientHeight    =   7305
   ClientLeft      =   1350
   ClientTop       =   2835
   ClientWidth     =   7560
   Icon            =   "frmRegVentaDet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   7560
   Begin VB.TextBox txtMNacional 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   43
      Text            =   "0"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtTipoCambio 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   40
      Text            =   "0"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CheckBox chkDocAnulado 
      Caption         =   "Documento Anulado"
      Height          =   315
      Left            =   4560
      TabIndex        =   37
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Frame fraReferencia 
      Caption         =   "Referencia"
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
      Height          =   615
      Left            =   120
      TabIndex        =   31
      Top             =   2400
      Width           =   7365
      Begin VB.ComboBox cboDocTpoRefe 
         Height          =   315
         Left            =   690
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtNumRefe 
         Height          =   315
         Left            =   4080
         MaxLength       =   8
         TabIndex        =   33
         Top             =   240
         Width           =   1365
      End
      Begin VB.TextBox txtSerRefe 
         Height          =   315
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   32
         Top             =   240
         Width           =   555
      End
      Begin MSMask.MaskEdBox txtDocFecRef 
         Height          =   315
         Left            =   6060
         TabIndex        =   36
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo"
         Height          =   225
         Left            =   240
         TabIndex        =   39
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha"
         Height          =   225
         Left            =   5520
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Número"
         Height          =   225
         Left            =   2760
         TabIndex        =   34
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.CheckBox ckSinContrato 
      Caption         =   "Sin Contrato"
      Height          =   495
      Left            =   3600
      TabIndex        =   30
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtDocHora 
      Height          =   315
      Left            =   7680
      TabIndex        =   28
      Top             =   1140
      Width           =   555
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   6030
      TabIndex        =   12
      Top             =   5970
      Width           =   1365
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   4620
      TabIndex        =   11
      Top             =   5970
      Width           =   1365
   End
   Begin VB.Frame Frame5 
      Caption         =   "Observaciones"
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
      Height          =   3345
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   4245
      Begin VB.TextBox txtDescrip 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2925
         Left            =   150
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   270
         Width           =   3945
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Contrato"
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
      Left            =   4680
      TabIndex        =   21
      Top             =   120
      Width           =   2805
      Begin VB.TextBox txtCodCta 
         Height          =   315
         Left            =   900
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label Label5 
         Caption         =   "Número"
         Height          =   225
         Left            =   180
         TabIndex        =   22
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cliente"
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
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   7365
      Begin Sicmact.TxtBuscar txtProvCod 
         Height          =   345
         Left            =   630
         TabIndex        =   29
         Top             =   270
         Width           =   1935
         _extentx        =   3413
         _extenty        =   609
         appearance      =   1
         appearance      =   1
         font            =   "frmRegVentaDet.frx":030A
         appearance      =   1
         tipobusqueda    =   3
         stitulo         =   ""
         tipobuspers     =   1
      End
      Begin VB.CommandButton cmdExaCab 
         Caption         =   "..."
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2250
         TabIndex        =   20
         Top             =   300
         Width           =   285
      End
      Begin VB.TextBox txtProvNom 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2580
         TabIndex        =   6
         Top             =   270
         Width           =   4605
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Documento"
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
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   810
      Width           =   7365
      Begin VB.TextBox txtDocSerie 
         Height          =   315
         Left            =   3420
         MaxLength       =   4
         TabIndex        =   3
         Top             =   270
         Width           =   555
      End
      Begin VB.TextBox txtDocNro 
         Height          =   315
         Left            =   3990
         MaxLength       =   7
         TabIndex        =   4
         Top             =   270
         Width           =   1365
      End
      Begin VB.ComboBox cboDocTpo 
         Height          =   315
         Left            =   630
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtDocFecha 
         Height          =   315
         Left            =   6060
         TabIndex        =   5
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         Height          =   225
         Left            =   5490
         TabIndex        =   18
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Número"
         Height          =   225
         Left            =   2730
         TabIndex        =   17
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         Height          =   225
         Left            =   180
         TabIndex        =   16
         Top             =   330
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
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
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3465
      Begin VB.ComboBox cboOpeTpo 
         Height          =   315
         ItemData        =   "frmRegVentaDet.frx":0336
         Left            =   630
         List            =   "frmRegVentaDet.frx":0338
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
         Height          =   225
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Width           =   435
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1185
      Left            =   4530
      TabIndex        =   24
      Top             =   4680
      Width           =   2955
      Begin VB.TextBox txtPVenta 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Height          =   285
         Left            =   1500
         TabIndex        =   10
         Top             =   840
         Width           =   1365
      End
      Begin VB.TextBox txtIGV 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Left            =   1500
         TabIndex        =   9
         Top             =   510
         Width           =   1365
      End
      Begin VB.TextBox txtVVenta 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Left            =   1500
         TabIndex        =   8
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Precio Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   27
         Top             =   870
         Width           =   1245
      End
      Begin VB.Label lblSTot 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   26
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "I.G.V."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   25
         Top             =   540
         Width           =   915
      End
      Begin VB.Shape ShapeIGV 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   60
         Top             =   150
         Width           =   2835
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   60
         Top             =   480
         Width           =   2835
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   60
         Top             =   810
         Width           =   2835
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Moneda Nacional"
      Height          =   255
      Left            =   4680
      TabIndex        =   42
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Tipo Cambio"
      Height          =   255
      Left            =   4680
      TabIndex        =   41
      Top             =   3840
      Width           =   975
   End
End
Attribute VB_Name = "frmRegVentaDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql   As String
Dim rs     As New ADODB.Recordset
Dim lNuevo As Boolean
Dim nTasaIGV As Currency
Dim oReg As DRegVenta
Dim sNumDoc As String
Dim lsDocRef As String
Dim lsDocFecRef As Date
Dim lnDocTpoRefIndex As Integer 'EJVG20140915
Dim nMoneda As Integer 'YIHU20152002-ERS181-2014. Se agregó pnMoneda
Dim cTipoAccion As String 'NAGL 20170804

Public Sub inicio(plNuevo As Boolean, pnTasaIgv As Currency, Optional pnMoneda As Integer = 0, Optional psTipoAccion As String = "")
lNuevo = plNuevo
nTasaIGV = pnTasaIgv
nMoneda = pnMoneda 'YIHU20152002-ERS181-2014. Se agregó pnMoneda
cTipoAccion = psTipoAccion 'NAGL 20170804 Agregó psTipoAccion
Me.Show 1
End Sub

Private Sub cboDocTpo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (cTipoAccion = "A") Then
       txtDocSerie.SetFocus
    Else
       txtDocNro.SetFocus
    End If '***NAGL 20170804
End If
End Sub

Private Sub cboDocTpo_Click()
'If Trim(Right(cboDocTpo.Text, 2)) = "7" Then
If Trim(Right(cboDocTpo.Text, 2)) = "7" Or Trim(Right(cboDocTpo.Text, 2)) = "8" Then 'EJVG20140915
    cboDocTpoRefe.ListIndex = lnDocTpoRefIndex
    txtSerRefe = Mid(txtSerRefe, 1, 4) 'NAGL 20170801 CAMBIO DE 3 a 4 y de lsDocRef a txtSerRefe
    txtNumRefe = Mid(txtNumRefe, 5, 12) 'NAGL 20170801 CAMBIO DE 4 a 5
    txtDocFecRef = Format(lsDocFecRef, "dd/mm/yyyy")
    fraReferencia.Enabled = True
Else
    lnDocTpoRefIndex = cboDocTpoRefe.ListIndex
    If Trim(txtSerRefe) <> "" Then
        lsDocRef = txtSerRefe & txtNumRefe
    End If
    cboDocTpoRefe.ListIndex = -1
    txtSerRefe = ""
    txtNumRefe = ""
    If txtDocFecRef <> "  /  /    " Then
    lsDocFecRef = txtDocFecRef.Text
    txtDocFecRef.Text = "  /  /    "
    End If
    fraReferencia.Enabled = False
End If
End Sub
Private Sub cboDocTpoRefe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       EnfocaControl txtSerRefe
    End If
End Sub
Private Sub cboOpeTpo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   'txtCodCta.SetFocus
   EnfocaControl txtCodCta
End If
End Sub

Private Function datosOk() As Boolean
datosOk = False
If cboOpeTpo.ListIndex = -1 Then
   MsgBox "Tipo de Operación no definido...!", vbInformation, "! Aviso !"
   cboOpeTpo.SetFocus
   Exit Function
End If

If txtCodCta = "" And txtCodCta.Enabled = True Then
   MsgBox "Codigo de Contrato no definido...!", vbInformation, "! Aviso !"
   txtCodCta.SetFocus
   Exit Function
End If
If cboDocTpo.ListIndex = -1 Then
   MsgBox "Tipo de Documento no definido...!", vbInformation, "! Aviso !"
   cboDocTpo.SetFocus
   Exit Function
End If
'If txtDocSerie = "" Then
If Val(Trim(txtDocSerie)) = 0 Then
   MsgBox "Serie de Documento no definido...!", vbInformation, "! Aviso !"
   txtDocSerie.SetFocus
   Exit Function
End If
'If txtDocNro = "" Then
If Val(Trim(txtDocNro)) = 0 Then
   MsgBox "Número de Documento no definido...!", vbInformation, "! Aviso !"
   txtDocNro.SetFocus
   Exit Function
End If
If ValidaFecha(txtDocFecha) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "! Aviso !"
   txtDocFecha.SetFocus
   Exit Function
End If
If txtProvCod.Text = "" Then
   MsgBox "Proveedor no identificado...!", vbInformation, "! Aviso !"
   txtProvCod.SetFocus
   Exit Function
End If
'If txtDescrip = "" Then
If Len(Trim(txtDescrip)) = 0 Then
   MsgBox "Descripción de Operación no definido...!", vbInformation, "! Aviso !"
   txtDescrip.SetFocus
   Exit Function
End If

If nVal(txtPVenta) = 0 And Trim(Right(cboOpeTpo.Text, 3)) <> "7" Then
   MsgBox "Monto de Operación no indicado...!", vbInformation, "! Aviso !"
   txtVVenta.SetFocus
   Exit Function
End If

Set rs = oReg.CargaSerieValida(Trim(txtDocSerie), gsCodAge, gsCodArea & gsCodAge, cTipoAccion) 'NAGL 20170801

If Not (rs.BOF And rs.EOF) Then
    If rs!ParamValidar = "NO" Then
       MsgBox "Número de Serie no válido...!", vbInformation, "! Aviso !"
       txtDocSerie.SetFocus
       Exit Function
    End If
End If 'NAGL 20170801

'YIHU 20150218-ERS181 *********
If nVal(txtTipoCambio) = 0 And nMoneda = 2 Then
   MsgBox "El Tipo de Cambio no ha sido ingresado", vbInformation, "!Aviso!"
   txtTipoCambio.SetFocus
   Exit Function
End If
'END YIHU 20150218-ERS181 *********


If Not validaReferencia Then Exit Function
datosOk = True
End Function
Private Function validaReferencia() As Boolean
      If fraReferencia.Enabled Then
        If cboDocTpoRefe.ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar el Tipo de Documento de Referencia", vbInformation, "Aviso"
            EnfocaControl cboDocTpoRefe
            Exit Function
        End If
        'If Val(Trim(txtSerRefe.Text)) = 0 Then
        If txtSerRefe.Text = "" Then 'NAGL 20180907
            MsgBox "Ud. debe especificar la Serie del Documento de Referencia", vbInformation, "Aviso"
            EnfocaControl txtSerRefe
            Exit Function
        End If
        If Val(Trim(txtNumRefe.Text)) = 0 Then
            MsgBox "Ud. debe especificar el Nro. del Documento de Referencia", vbInformation, "Aviso"
            EnfocaControl txtNumRefe
            Exit Function
        End If
    End If
    validaReferencia = True
End Function
Private Sub chkDocAnulado_Click()
    If Me.chkDocAnulado.value = 1 Then
        Me.txtDescrip = "A N U L A D O"
        txtDescrip.Locked = True
    Else
        Me.txtDescrip = ""
        txtDescrip.Locked = False
    End If
    
End Sub

'ALPA 20090924*****************************************
Private Sub ckSinContrato_Click()
    If txtCodCta.Enabled = False Then
        txtCodCta.Text = sNumDoc
        txtCodCta.Enabled = True
    Else
        txtCodCta.Text = ""
        txtCodCta.Enabled = False
    End If
End Sub
'******************************************************
Private Sub cmdAceptar_Click()
Dim nDocTpo As Long
Dim sDocNro As String
Dim dFecha  As Date
Dim sFecha  As String
Dim oMov As DMov 'EJVG20140915
Dim sMovNro As String
Dim nMovNro As Long
Dim nDocTpoRefe As Integer

Dim RSBusca As ADODB.Recordset

On Error GoTo errAcepta


'*** PEAC 20100707

'If Not datosOk Then
'   Exit Sub
'End If

If Me.chkDocAnulado.value = 0 Then
    If Not datosOk Then
       Exit Sub
    End If
Else 'verifica datos cuando está anulado
    If cboDocTpo.ListIndex = -1 Then
       MsgBox "Tipo de Documento no definido...!", vbInformation, "! Aviso !"
       cboDocTpo.SetFocus
       Exit Sub
    End If
    'If txtDocSerie = "" Then
    If Val(Trim(txtDocSerie)) = 0 Then
       MsgBox "Serie de Documento no definido...!", vbInformation, "! Aviso !"
       txtDocSerie.SetFocus
       Exit Sub
    End If
    'If txtDocNro = "" Then
    If Val(Trim(txtDocNro)) = 0 Then
       MsgBox "Número de Documento no definido...!", vbInformation, "! Aviso !"
       txtDocNro.SetFocus
       Exit Sub
    End If
    If ValidaFecha(txtDocFecha) <> "" Then
       MsgBox "Fecha no válida...!", vbInformation, "! Aviso !"
       txtDocFecha.SetFocus
       Exit Sub
    End If
    If Not validaReferencia Then Exit Sub
End If
'*** FIN PEAC

If MsgBox(" ¿ Seguro de grabar datos ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbNo Then
   Exit Sub
End If
nDocTpo = gnDocTpo
sDocNro = gsDocNro
dFecha = gdFecha
gnDocTpo = Right(cboDocTpo, 2)
gsDocNro = Trim(txtDocSerie & txtDocNro)
'ALPA 20090925************************************
'ALPA 20091128************************************
nDocTpoRefe = Val(Trim(Right(cboDocTpoRefe.Text, 2)))
lsDocRef = Trim(txtSerRefe & txtNumRefe)
If Trim(txtDocFecRef) <> "/  /" Then
    lsDocFecRef = CDate(Format(txtDocFecRef & " " & GetHoraServer, "dd/mm/yyyy hh:mm:ss"))
End If
'*************************************************
'*************************************************
If txtDocHora = "" Then
   sFecha = Format(txtDocFecha & " " & GetHoraServer, gsFormatoFechaHora)
   gdFecha = CDate(Format(txtDocFecha & " " & GetHoraServer, "dd/mm/yyyy hh:mm:ss"))
Else
   sFecha = Format(txtDocFecha & " " & txtDocHora, gsFormatoFechaHora)
   gdFecha = CDate(Format(txtDocFecha & " " & txtDocHora, "dd/mm/yyyy hh:mm:ss"))
End If

If lNuevo Then
'ALPA 20090925*************************************************
   'oReg.InsertaVenta Right(cboOpeTpo, 1), gnDocTpo, gsDocNro, gdFecha, txtProvCod.Tag, txtCodCta, txtDescrip, nVal(txtVVenta), nVal(txtIGV), nVal(txtPVenta)
   Set RSBusca = oReg.VerificaVentaExistente(gnDocTpo, gsDocNro, gdFecha, Right(cboOpeTpo, 1)) '*** PEAC 20110425
   If Not (RSBusca.EOF And RSBusca.BOF) Then
    MsgBox "Este registro ya fue ingresado.", vbOKOnly + vbExclamation, "Atención"
    Exit Sub
   End If
   Set oMov = New DMov
   sMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
   oMov.InsertaMov sMovNro, "760210", "REGISTRO DE VENTAS", gMovEstContabNoContable, gMovFlagVigente
   nMovNro = oMov.GetnMovNro(sMovNro)
   Set oMov = Nothing
   
'YIHU 20150218-ERS181.
    If nMoneda <> 1 Then ' si se va a actualizar con moneda extranjera
        oReg.InsertaVenta Right(cboOpeTpo, 1), gnDocTpo, gsDocNro, gdFecha, txtProvCod.Tag, txtCodCta, txtDescrip, IIf(Me.chkDocAnulado.value = 1, 0, nVal(txtVVenta) * nVal(txtTipoCambio)), IIf(Me.chkDocAnulado.value = 1, 0, nVal(txtIGV) * nVal(txtTipoCambio)), IIf(Me.chkDocAnulado.value = 1, 0, nVal(txtPVenta) * nVal(txtTipoCambio)), lsDocRef, lsDocFecRef, nDocTpoRefe, nMovNro, nMoneda, txtTipoCambio.Text, gsCodArea & gsCodAge 'NAGL 20170801
    Else
        oReg.InsertaVenta Right(cboOpeTpo, 1), gnDocTpo, gsDocNro, gdFecha, txtProvCod.Tag, txtCodCta, txtDescrip, IIf(Me.chkDocAnulado.value = 1, 0, nVal(txtVVenta)), IIf(Me.chkDocAnulado.value = 1, 0, nVal(txtIGV)), IIf(Me.chkDocAnulado.value = 1, 0, nVal(txtPVenta)), lsDocRef, lsDocFecRef, nDocTpoRefe, nMovNro, nMoneda, txtTipoCambio.Text, gsCodArea & gsCodAge 'NAGL 20170801
    End If
    'END YIHU
   
   
'YIHU 20150218-ERS181. Se agregó el campo txtTipoCambio, nMoneda
Else
'ALPA 20090925*************************************************
   'oReg.ActualizaVenta Right(cboOpeTpo, 1), gnDocTpo, gsDocNro, gdFecha, txtProvCod.Tag, txtCodCta, txtDescrip, nVal(txtVVenta), nVal(txtIGV), nVal(txtPVenta), nDocTpo, sDocNro, Format(dFecha, gsFormatoFechaHora)
   
   'YIHU 201502-18-ERS181
   If nMoneda <> 1 Then ' si se va a actualizar con moneda extranjera
    oReg.ActualizaVenta Right(cboOpeTpo, 1), gnDocTpo, gsDocNro, gdFecha, txtProvCod.Tag, txtCodCta, txtDescrip, IIf(Me.chkDocAnulado.value = 1, 0, nVal(txtVVenta) * nVal(txtTipoCambio)), IIf(Me.chkDocAnulado.value = 1, 0, nVal(txtIGV) * nVal(txtTipoCambio)), IIf(Me.chkDocAnulado.value = 1, 0, nVal(txtPVenta) * nVal(txtTipoCambio)), nDocTpo, sDocNro, Format(dFecha, "dd/mm/yyyy hh:mm:ss"), lsDocRef, lsDocFecRef, nDocTpoRefe, txtTipoCambio.Text
   Else
    oReg.ActualizaVenta Right(cboOpeTpo, 1), gnDocTpo, gsDocNro, gdFecha, txtProvCod.Tag, txtCodCta, txtDescrip, IIf(Me.chkDocAnulado.value = 1, 0, nVal(txtVVenta)), IIf(Me.chkDocAnulado.value = 1, 0, nVal(txtIGV)), IIf(Me.chkDocAnulado.value = 1, 0, nVal(txtPVenta)), nDocTpo, sDocNro, Format(dFecha, "dd/mm/yyyy hh:mm:ss"), lsDocRef, lsDocFecRef, nDocTpoRefe, txtTipoCambio.Text
   End If
   'END YIHU*************
    
'**************************************************************
'YIHU 20150218-ERS181. Se agregó el campo txtTipoCambio

End If
glAceptar = True
Unload Me
Exit Sub
errAcepta:
   MsgBox TextErr(Err.Description), vbInformation, "! Aviso !"
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim oPer As New UPersona
CentraForm Me
Set oReg = New DRegVenta
Set rs = oReg.CargaRegOperacion()
RSLlenaCombo rs, cboOpeTpo

If (cTipoAccion = "M") Then
   txtDocSerie.Enabled = False
End If 'NAGL 20170805

'YIHU20152002-ERS181-2014
If nMoneda = 2 Then
    txtVVenta.BackColor = &H80FF80
    txtIGV.BackColor = &H80FF80
    txtPVenta.BackColor = &H80FF80
    
End If

calcularMonedaNacional
txtTipoCambio.Text = Format(txtTipoCambio.Text, "#0.000")
If nMoneda = 1 Then
    txtTipoCambio.Enabled = False
End If
'END YIHU ***************

Dim oOpe As New DOperacion
Set rs = oOpe.CargaOpeDoc(gsOpeCod)
CentraForm Me
Do While Not rs.EOF
   cboDocTpo.AddItem rs!cDocDesc & space(50) & Trim(rs!cDocAbrev) & " " & Trim(rs!nDocTpo)
   rs.MoveNext
Loop
Set rs = oOpe.CargaOpeDoc("760211")
Do While Not rs.EOF
   cboDocTpoRefe.AddItem rs!cDocDesc & space(50) & Trim(rs!cDocAbrev) & " " & Trim(rs!nDocTpo)
   rs.MoveNext
Loop
lnDocTpoRefIndex = -1
lsDocFecRef = Format(gdFecSis, "dd/mm/yyyy")
If Not lNuevo Then
   Set rs = oReg.CargaRegistro(gnDocTpo, gsDocNro, gdFecha, gdFecha, gsCodArea & gsCodAge) 'NAGL 20170804
   If Not rs.EOF Then
      cboOpeTpo.ListIndex = BuscaCombo(rs!cOpeTpo, cboOpeTpo)
      cboDocTpo.ListIndex = BuscaCombo(rs!nDocTpo, cboDocTpo)
      txtCodCta = rs!cCtaCod
      If Trim(txtCodCta) = "" Then
        ckSinContrato.value = 1
      Else
        ckSinContrato.value = 0
      End If
      txtDocSerie = Mid(rs!cDocNroNew, 1, 4) 'NAGL ERS 012 - 2017 Se cambio la Long.Nro Serie de 3 a 4 Dígitos, además del Campo
      txtDocNro = Mid(rs!cDocNroNew, 5, 12) 'NAGL ERS 012 - 2017 Se toma desde la posición 5, además del Campo
      txtDocFecha = Format(rs!dDocFecha, "dd/mm/yyyy")
      txtDocHora = Right(Format(rs!dDocFecha, "dd/mm/yyyy hh:mm:ss"), 8)
      txtDescrip = rs!cDescrip
    'YIHU20152002-ERS181-2014 CUANDO ES MONEDA EXTRANJERA DIVIDIMOS ENTRE EL TIPO DE CAMBIO
      If nMoneda = 2 Then
        txtVVenta = Format(rs!nVVenta / rs!nTipoCambio, gsFormatoNumeroView)
        txtIGV = Format(rs!nIGV / rs!nTipoCambio, gsFormatoNumeroView)
        txtPVenta = Format(rs!nPVenta / rs!nTipoCambio, gsFormatoNumeroView)
      Else
        txtVVenta = Format(rs!nVVenta, gsFormatoNumeroView)
        txtIGV = Format(rs!nIGV, gsFormatoNumeroView)
        txtPVenta = Format(rs!nPVenta, gsFormatoNumeroView)
      End If
      'END YIHU ********************
      
      
    txtTipoCambio = IIf(IsNull(rs!nTipoCambio), 0, rs!nTipoCambio) 'YIHU20152002-ERS181-2014
      sNumDoc = txtCodCta.Text
      
    '*** PEAC 201007017
    If Len(rs!cPersCod) > 0 Then
        oPer.ObtieneClientexCodigo rs!cPersCod
        txtProvCod.Tag = oPer.sPersCod
        txtProvNom = oPer.sPersNombre
        txtProvCod = oPer.sPersIdnroRUC
    Else
        'Me.chkDocAnulado.value = 1
    End If
    '*** FIN PEAC
    'EJVG20140915 ***
    If UCase(txtDescrip.Text) = "A N U L A D O" Then
        chkDocAnulado.value = 1
    Else
        chkDocAnulado.value = 0
    End If
    'END EJVG *******
    
      If txtProvCod = "" Then
         txtProvCod = oPer.sPersIdnroDNI
      End If
      'If rs!nDocTpo = 7 Then
      If rs!nDocTpo = 7 Or rs!nDocTpo = 8 Then
        fraReferencia.Enabled = False
        cboDocTpoRefe.ListIndex = BuscaCombo(IIf(IsNull(rs!nDocTpoRefe), 0, rs!nDocTpoRefe), cboDocTpoRefe)
        txtSerRefe.Text = Mid(IIf(IsNull(rs!cDocNroRefe), "", rs!cDocNroRefe), 1, 4) 'NAGL 20170801 CAMBIO DE 3 a 4
        txtNumRefe.Text = Mid(IIf(IsNull(rs!cDocNroRefe), "", rs!cDocNroRefe), 5, 12) 'NAGL 20170801 CAMBIO DE 4 a 5
        txtDocFecRef.Text = Format(IIf(IsNull(rs!dDocRefeFec), "", rs!dDocRefeFec), "DD/MM/YYYY")
        fraReferencia.Enabled = True
      Else
        fraReferencia.Enabled = True
        cboDocTpoRefe.ListIndex = -1
        txtSerRefe.Text = ""
        txtNumRefe.Text = ""
        fraReferencia.Enabled = False
      End If
   End If
End If
rs.Close: Set rs = Nothing
End Sub
Private Sub txtDocFecRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtDescrip
    End If
End Sub
Private Sub txtCodCta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   cboDocTpo.SetFocus
End If
End Sub

Private Sub txtNumRefe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtDocFecRef
    End If
End Sub
Private Sub txtNumRefe_LostFocus()
    txtNumRefe = Format(txtNumRefe, "00000000")
End Sub
Private Sub txtSerRefe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtNumRefe
    End If
End Sub
Private Sub txtSerRefe_LostFocus()
    txtSerRefe = Format(txtSerRefe, "0000") 'NAGL 20170801 CAMBIO DE TRES CEROS A CUATRO
End Sub
Private Sub TxtDescrip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   txtVVenta.SetFocus
End If
End Sub

Private Sub txtDocFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtDocFecha) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "! Aviso !"
      Exit Sub
   End If
   txtProvCod.SetFocus
End If
End Sub

Private Sub txtDocNro_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtDocNro = Format(txtDocNro, "0000000") 'NAGL ERS 012-2017 de 8 a 7 Dígitos
   txtDocFecha.SetFocus
End If
End Sub

Private Sub txtDocNro_LostFocus()
    txtDocNro = Format(txtDocNro, "0000000") 'NAGL ERS 012-2017 de 8 a 7 Dígitos
End Sub

Private Sub txtDocSerie_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtDocSerie = Format(txtDocSerie, "0000") 'NAGL ERS 012-2017 de 3 a 4 Dígitos
   txtDocNro.SetFocus
End If
End Sub

Private Sub txtDocSerie_LostFocus()
    txtDocSerie = Format(txtDocSerie, "0000") 'NAGL ERS 012-2017 de 3 a 4 Dígitos
End Sub

Private Sub txtIgv_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtIGV, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   txtIGV = Format(txtIGV, gsFormatoNumeroView)
   txtPVenta = Format(nVal(txtVVenta) + nVal(txtIGV), gsFormatoNumeroView)
   txtPVenta.SetFocus
End If
End Sub

Private Sub txtProvCod_EmiteDatos()
txtProvCod.Tag = txtProvCod.Text
txtProvNom = txtProvCod.psDescripcion
txtProvCod.Text = txtProvCod.sPersNroDoc
Me.txtDescrip.SetFocus
End Sub

Private Sub txtPVenta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPVenta, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   txtVVenta = Format(Round(nVal(txtPVenta) / (1 + nTasaIGV), 2), gsFormatoNumeroView)
   txtIGV = Format(nVal(txtPVenta) - nVal(txtVVenta), gsFormatoNumeroView)
   txtPVenta = Format(txtPVenta, gsFormatoNumeroView)
   cmdAceptar.SetFocus
End If
End Sub

Private Sub txtVVenta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtVVenta, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   txtIGV = Format(Round(nVal(txtVVenta) * nTasaIGV, 2), gsFormatoNumeroView)
   txtPVenta = Format(nVal(txtVVenta) + nVal(txtIGV), gsFormatoNumeroView)
   txtVVenta = Format(txtVVenta, gsFormatoNumeroView)
   txtIGV.SetFocus
End If
End Sub

'YIHU20152002-ERS181-2014
Private Sub txtPVenta_Change()
    'txtMNacional.Text = Format(Val(txtPVenta) * Val(txtTipoCambio), gsFormatoNumeroView)
    calcularMonedaNacional
End Sub

Private Sub calcularMonedaNacional()
    txtMNacional.Text = Format(nVal(txtPVenta) * nVal(txtTipoCambio), gsFormatoNumeroView)
End Sub
'END YIHU

'YIHU20152002-ERS181-2014
Private Sub txtTipoCambio_Change()
    If Not IsNumeric(txtTipoCambio.Text) Then
        MsgBox "Se dede ingresar valores numéricos", vbInformation, "¡Aviso!"
    End If
    calcularMonedaNacional
    Exit Sub
End Sub
'END YIHU

'VAPI SEGUN ERS 181-2014
Private Sub txtTipoCambio_LostFocus()
If Trim(txtTipoCambio.Text) = "" Then
        txtTipoCambio.Text = "0.000"
    End If
    txtTipoCambio.Text = Format(txtTipoCambio.Text, "#0.000")
End Sub

Private Sub txtTipoCambio_GotFocus()
    fEnfoque txtTipoCambio
End Sub

'FIN VAPI

