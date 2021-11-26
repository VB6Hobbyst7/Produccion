VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMntAjusteInfla 
   Caption         =   "Ajuste por Inflación : Mantenimiento Datos Históricos"
   ClientHeight    =   4680
   ClientLeft      =   2085
   ClientTop       =   2100
   ClientWidth     =   10305
   Icon            =   "frmMntAjusteInfla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEjercicio 
      Caption         =   "Nuevo E&jercicio"
      Height          =   390
      Left            =   3990
      TabIndex        =   11
      ToolTipText     =   "Transferencia de Valores Ajustados a Históricos"
      Top             =   4080
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   3885
      Left            =   90
      TabIndex        =   16
      Top             =   90
      Width           =   10125
      Begin MSDataGridLib.DataGrid grdAjuste 
         Height          =   3570
         Left            =   60
         TabIndex        =   0
         Top             =   210
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   6297
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   2
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "cCtaContCod"
            Caption         =   "Cuenta Contable"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cAjusteCod"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cAjusteDescrip"
            Caption         =   "Descripción"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "dAjusteFecha"
            Caption         =   "Fecha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "nAjusteValor1"
            Caption         =   "Valor 01"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "nAjusteValor2"
            Caption         =   "Valor 02"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "nAjusteValor3"
            Caption         =   "Valor 03"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            ScrollBars      =   2
            BeginProperty Column00 
               DividerStyle    =   6
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   6
               ColumnWidth     =   2924.788
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               DividerStyle    =   6
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               DividerStyle    =   6
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               DividerStyle    =   6
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1094.74
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   5550
         TabIndex        =   4
         Top             =   3465
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtValor3 
         Height          =   315
         Left            =   8850
         TabIndex        =   7
         Top             =   3465
         Visible         =   0   'False
         Width           =   1065
      End
      Begin Sicmact.TxtBuscar txtCta 
         Height          =   330
         Left            =   270
         TabIndex        =   1
         Top             =   3465
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
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
         sTitulo         =   ""
      End
      Begin VB.TextBox txtAjusteCod 
         Height          =   315
         Left            =   1860
         TabIndex        =   2
         Top             =   3465
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtValor2 
         Height          =   315
         Left            =   7770
         TabIndex        =   6
         Top             =   3465
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   2430
         TabIndex        =   3
         Top             =   3465
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.TextBox txtValor1 
         Height          =   315
         Left            =   6690
         TabIndex        =   5
         Top             =   3465
         Visible         =   0   'False
         Width           =   1065
      End
   End
   Begin VB.Frame fraControl 
      Height          =   585
      Left            =   90
      TabIndex        =   14
      Top             =   3930
      Width           =   10125
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   390
         Left            =   120
         TabIndex        =   8
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   390
         Left            =   1380
         TabIndex        =   9
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   390
         Left            =   2640
         TabIndex        =   10
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         CausesValidation=   0   'False
         Height          =   390
         Left            =   8370
         TabIndex        =   13
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         CausesValidation=   0   'False
         Height          =   390
         Left            =   8370
         TabIndex        =   15
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   7110
         TabIndex        =   12
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMntAjusteInfla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lConsulta As Boolean
Dim lNuevo As Boolean
Dim sSql As String
Dim rsAju As ADODB.Recordset
Dim clsAjuste As DAjusteCont

Public Sub Inicio(plConsulta As Boolean)
lConsulta = plConsulta
Me.Show 1
End Sub

Private Sub cmdEjercicio_Click()
Dim sELIM As String, cCtaImp
If MsgBox(" Los Valores Ajustados se trasladarán a Valores Históricos del siguiente Ejercicio." & Chr(10) & " ¿ Está seguro de Desea realizar Transferencia ? ", vbQuestion + vbYesNo, "Mensaje de confirmación") = vbYes Then
    Me.Enabled = False
   clsAjuste.TrasladoValorAjustado
   CargaAjuste
End If
Me.Enabled = True
grdAjuste.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim oCta As New DCtaCont
Set clsAjuste = New DAjusteCont
CargaAjuste
txtCta.rs = oCta.CargaCtaCont()
txtCta.EditFlex = False
txtCta.TipoBusqueda = BuscaGrid
txtCta.sTitulo = "Cuenta Contable"
Set oCta = Nothing
End Sub
Private Sub CargaAjuste()
Set rsAju = clsAjuste.CargaAjusteInflacion(, , , adLockOptimistic)
Set grdAjuste.DataSource = rsAju
End Sub

Private Sub cmdNuevo_Click()
txtCta.Text = ""
txtAjusteCod.Text = ""
txtDesc.Text = ""
txtFecha.Text = "  /  /    "
txtValor1.Text = ""
txtValor2.Text = ""
txtValor3.Text = ""
ActivaBotones False
txtCta.Enabled = True
txtAjusteCod.Enabled = True
txtCta.SetFocus
lNuevo = True
End Sub

Private Sub cmdModificar_Click()
txtCta.Text = rsAju!cCtaContCod
txtAjusteCod.Text = rsAju!cAjusteCod
txtDesc.Text = rsAju!cAjusteDescrip
txtFecha.Text = rsAju!dAjusteFecha
txtValor1.Text = rsAju!nAjusteValor1
txtValor2.Text = rsAju!nAjusteValor2
txtValor3.Text = rsAju!nAjusteValor3
ActivaBotones False
txtCta.Enabled = False
txtAjusteCod.Enabled = False
txtDesc.SetFocus
lNuevo = False
End Sub

Private Sub cmdEliminar_Click()
Dim sELIM As String, cCtaImp
If MsgBox(" ¿ Está seguro de eliminar el impuesto ? ", vbQuestion + vbYesNo, "Mensaje de confirmación") = vbYes Then
   clsAjuste.EliminaAjuste rsAju!cCtaContCod, rsAju!cAjusteCod
   rsAju.Delete adAffectCurrent
End If
grdAjuste.SetFocus
End Sub

Private Sub cmdCancelar_Click()
ActivaBotones True
grdAjuste.SetFocus
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If Len(txtCta.Text) = 0 Then
   MsgBox "Falta ingresar Cuenta Contable...  ", vbInformation, "¡Aviso!"
   txtCta.SetFocus
   Exit Function
End If
If Len(txtAjusteCod) <> 2 Then
   MsgBox "Falta ingresar Correlativo de Cuenta Contable", vbInformation, "¡Aviso!"
   txtAjusteCod.SetFocus
   Exit Function
End If
If Len(txtDesc) = 0 Then
   MsgBox "Falta ingresar Concepto de Ajuste", vbInformation, "¡Aviso!"
   txtDesc.SetFocus
   Exit Function
End If
If Len(txtFecha.Text) = 0 Then
   MsgBox "  Abreviatura de impuesto no válida ...  ", vbInformation, "¡Aviso!"
   txtFecha.SetFocus
   Exit Function
End If

If Val(txtValor3.Text) = 0 Then
   MsgBox " Falta Ingresar Valor Histórico del Concepto...  ", vbInformation, "¡Aviso!"
   txtValor3.SetFocus
   Exit Function
End If

ValidaDatos = True
End Function

Private Sub cmdAceptar_Click()
On Error GoTo AceptarErr
If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox(" ¿ Seguro de grabar datos del impuesto ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   Select Case lNuevo
      Case True
            clsAjuste.InsertaAjuste txtCta.Text, txtAjusteCod, txtDesc, Format(txtFecha.Text, gsFormatoFecha), Val(Format(txtValor1.Text, gsFormatoNumeroDato)), Val(Format(txtValor2.Text, gsFormatoNumeroDato)), Val(Format(txtValor3.Text, gsFormatoNumeroDato))
      Case False
            clsAjuste.ActualizaAjuste txtCta.Text, txtAjusteCod, txtDesc, Format(txtFecha.Text, gsFormatoFecha), Val(Format(txtValor1.Text, gsFormatoNumeroDato)), Val(Format(txtValor2.Text, gsFormatoNumeroDato)), Val(Format(txtValor3.Text, gsFormatoNumeroDato))
   End Select
   CargaAjuste
   rsAju.Find "cCtaContCod = '" & txtCta.Text & "'", , , 1
   rsAju.Find "cAjusteCod = '" & txtAjusteCod & "'"
End If
ActivaBotones True
grdAjuste.SetFocus
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsAju
Set clsAjuste = Nothing
End Sub

Private Sub grdAjuste_GotFocus()
grdAjuste.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdAjuste_LostFocus()
grdAjuste.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub txtCta_EmiteDatos()
If txtCta.psDescripcion <> "" Then
   txtAjusteCod.SetFocus
End If
End Sub

Private Sub txtCta_GotFocus()
txtCta.BackColor = "&H00F5FFDD"
fEnfoque txtCta
End Sub
Private Sub txtCta_LostFocus()
txtCta.BackColor = "&H80000005"
End Sub
Private Sub txtCta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtAjusteCod_GotFocus()
txtAjusteCod.BackColor = "&H00F5FFDD"
fEnfoque txtAjusteCod
End Sub
Private Sub txtAjusteCod_LostFocus()
txtAjusteCod.BackColor = "&H80000005"
End Sub
Private Sub txtAjusteCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtDesc.SetFocus
End If
End Sub

Private Sub txtDesc_GotFocus()
txtDesc.BackColor = "&H00F5FFDD"
fEnfoque txtDesc
End Sub
Private Sub txtDesc_LostFocus()
txtDesc.BackColor = "&H80000005"
End Sub
Private Sub txtdesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFecha.SetFocus
End If
End Sub

Private Sub txtFecha_GotFocus()
txtFecha.BackColor = "&H00F5FFDD"
fEnfoque txtFecha
End Sub
Private Sub txtFecha_LostFocus()
txtFecha.BackColor = "&H80000005"
End Sub
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtValor1.SetFocus
End If
End Sub
Private Sub txtValor1_GotFocus()
txtValor1.BackColor = "&H00F5FFDD"
fEnfoque txtValor1
End Sub
Private Sub txtValor1_LostFocus()
txtValor1.BackColor = "&H80000005"
End Sub
Private Sub txtValor1_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtValor1, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   txtValor2.SetFocus
End If
End Sub
Private Sub txtValor2_GotFocus()
txtValor2.BackColor = "&H00F5FFDD"
fEnfoque txtValor2
End Sub
Private Sub txtValor2_LostFocus()
txtValor2.BackColor = "&H80000005"
End Sub
Private Sub txtValor2_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtValor2, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   txtValor3.SetFocus
End If
End Sub

Private Sub txtValor3_GotFocus()
txtValor3.BackColor = "&H00F5FFDD"
fEnfoque txtValor3
End Sub
Private Sub txtValor3_LostFocus()
txtValor3.BackColor = "&H80000005"
End Sub
Private Sub txtValor3_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtValor3, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   cmdAceptar.SetFocus
End If
End Sub

Sub ActivaBotones(lActiva As Boolean)
If lActiva Then
   grdAjuste.Height = 3570
Else
   grdAjuste.Height = 3180
End If
cmdNuevo.Visible = lActiva
cmdModificar.Visible = lActiva
cmdEliminar.Visible = lActiva
cmdSalir.Visible = lActiva
cmdAceptar.Visible = Not lActiva
cmdCancelar.Visible = Not lActiva
txtCta.Visible = Not lActiva
txtAjusteCod.Visible = Not lActiva
txtDesc.Visible = Not lActiva
txtFecha.Visible = Not lActiva
txtValor1.Visible = Not lActiva
txtValor2.Visible = Not lActiva
txtValor3.Visible = Not lActiva
End Sub

