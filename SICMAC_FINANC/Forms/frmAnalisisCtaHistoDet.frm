VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAnalisisCtaHistoDet 
   Caption         =   "Análisis de Cuentas: Datos Históricos: "
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   Icon            =   "frmAnalisisCtaHistoDet.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3540
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4140
      TabIndex        =   5
      Top             =   3090
      Width           =   1275
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5490
      TabIndex        =   6
      Top             =   3090
      Width           =   1275
   End
   Begin VB.Frame Frame1 
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
      Height          =   1935
      Left            =   180
      TabIndex        =   8
      Top             =   1020
      Width           =   6735
      Begin VB.TextBox tMovDesc 
         Height          =   675
         Left            =   1290
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   690
         Width           =   5295
      End
      Begin VB.TextBox tImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4980
         TabIndex        =   4
         Top             =   1470
         Width           =   1605
      End
      Begin MSMask.MaskEdBox tFecha 
         Height          =   285
         Left            =   1290
         TabIndex        =   2
         Top             =   270
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Importe"
         Height          =   285
         Left            =   4140
         TabIndex        =   11
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción"
         Height          =   225
         Left            =   270
         TabIndex        =   10
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   225
         Left            =   270
         TabIndex        =   9
         Top             =   300
         Width           =   645
      End
   End
   Begin VB.Frame fraCta 
      Caption         =   "Cuenta Contable "
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
      Height          =   840
      Left            =   180
      TabIndex        =   7
      Top             =   150
      Width           =   6735
      Begin VB.TextBox tCtaDesc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2580
         MaxLength       =   30
         TabIndex        =   1
         Top             =   375
         Width           =   3975
      End
      Begin Sicmact.TxtBuscar tCtaCod 
         Height          =   330
         Left            =   150
         TabIndex        =   0
         Top             =   375
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
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
   End
End
Attribute VB_Name = "frmAnalisisCtaHistoDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCtaCod  As String
Dim sMovEstado As MovEstado
Dim sMovFlag   As MovFlag
Dim nImporte As Currency
Dim lNuevo   As Boolean

Public Sub Inicio(psCtaCod As String, pnImporte As Currency, Optional psMovEstado As MovEstado = gMovEstContabMovContable, Optional psMovFlag As MovFlag = gMovFlagVigente, Optional plNuevo As Boolean = False)
sCtaCod = psCtaCod
sMovEstado = psMovEstado
sMovFlag = psMovFlag
nImporte = pnImporte
lNuevo = plNuevo
Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox(" ¿ Seguro que desea grabar datos ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
   Exit Sub
End If

Dim clsAna As New NAnalisisCtas
If lNuevo Then
   Dim oMov As New DMov
   gsMovNro = oMov.GeneraMovNro(tFecha, Right(gsCodAge, 2), gsCodUser)
   clsAna.InsertaPendienteHisto gsMovNro, tMovDesc, gsOpeCod, tCtaCod.Text, Format(tImporte, gsFormatoNumeroDato)
Else
   clsAna.ActualizaPendienteHisto gnMovNro, tMovDesc, tCtaCod, Format(tImporte, gsFormatoNumeroDato), sMovEstado, sMovFlag
End If
Set clsAna = Nothing
glAceptar = True
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Me.Caption = Me.Caption & IIf(lNuevo, "Nuevo", "Modificar")

If Mid(gsOpeCod, 3, 1) = "2" Then
   tImporte.BackColor = "&H0080FF80"
End If
If gsMovNro <> "" Then
   tFecha = GetFechaMov(gsMovNro, True)
Else
   tFecha = gdFecSis
End If
tMovDesc = gsGlosa
tImporte = Format(nImporte, gsFormatoNumeroView)

Dim oPend As New NAnalisisCtas
tCtaCod.rs = oPend.CargaCtaContPendiente(Mid(gsOpeCod, 3, 1), , True)
Set oPend = Nothing
tCtaCod.psRaiz = "Cuentas de Pendientes"
tCtaCod.TipoBusqueda = BuscaArbol
tCtaCod.EditFlex = False
tCtaCod.lbUltimaInstancia = False
tCtaCod = sCtaCod
tCtaDesc = tCtaCod.psDescripcion
If Not lNuevo Then
   tFecha.Enabled = False
End If
End Sub

Private Sub tCtaCod_EmiteDatos()
tCtaDesc = tCtaCod.psDescripcion
If tCtaDesc <> "" Then
   If tFecha.Enabled And tFecha.Visible Then
      tFecha.SetFocus
   Else
    If tMovDesc.Visible Then
        tMovDesc.SetFocus
    End If
   End If
End If
End Sub

Private Sub tFecha_GotFocus()
fEnfoque tFecha
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(tFecha) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "¡Aviso!"
      Exit Sub
   End If
   tMovDesc.SetFocus
End If
End Sub

Private Sub tImporte_GotFocus()
fEnfoque tImporte
End Sub

Private Sub tImporte_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(tImporte, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   tImporte = Format(tImporte, gsFormatoNumeroView)
   CmdAceptar.SetFocus
End If
End Sub

Private Sub tMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   tImporte.SetFocus
End If
End Sub


Private Function ValidaDatos() As Boolean
ValidaDatos = False
If tCtaCod = "" Then
   MsgBox "No ingresó Cuenta Contable!", vbInformation, "¡Aviso!"
   tCtaCod.SetFocus
   Exit Function
End If
If ValidaFecha(tFecha) <> "" Then
   MsgBox "Ingrese Fecha...!", vbInformation, "¡Aviso!"
   tFecha.SetFocus
   Exit Function
End If
If tMovDesc = "" Then
   MsgBox "No ingresó Descripción !", vbInformation, "¡Aviso!"
   tMovDesc.SetFocus
   Exit Function
End If
If Val(tImporte) = 0 Then
   MsgBox "Importe tiene q ser mayor a Cero !", vbInformation, "¡Aviso!"
   tImporte.SetFocus
   Exit Function
End If
ValidaDatos = True
End Function
