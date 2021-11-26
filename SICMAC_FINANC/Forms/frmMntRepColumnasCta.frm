VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntRepColumnasCta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes :Asignación de Cuentas Contables"
   ClientHeight    =   4065
   ClientLeft      =   1500
   ClientTop       =   3120
   ClientWidth     =   8355
   Icon            =   "frmMntRepColumnasCta.frx":0000
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2400
      TabIndex        =   1
      Top             =   3090
      Width           =   285
   End
   Begin VB.TextBox txtOpeCod 
      BackColor       =   &H00F0FFFF&
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
      Left            =   840
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtOpeDesc 
      BackColor       =   &H00F0FFFF&
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
      Left            =   2760
      TabIndex        =   6
      Top             =   240
      Width           =   5355
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   5595
      TabIndex        =   3
      Top             =   3585
      Width           =   1290
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   7035
      TabIndex        =   4
      Top             =   3585
      Width           =   1230
   End
   Begin VB.Frame Frame3 
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
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2790
      Width           =   8160
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   11
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtCta 
         Height          =   315
         Left            =   360
         MaxLength       =   12
         TabIndex        =   0
         Top             =   270
         Width           =   1965
      End
      Begin VB.TextBox txtDesc 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2610
         TabIndex        =   12
         Top             =   270
         Width           =   5295
      End
   End
   Begin VB.Frame fraOpe 
      Caption         =   "Operación "
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
      TabIndex        =   8
      Top             =   0
      Width           =   8175
      Begin VB.Label Label3 
         Caption         =   "Descripción"
         Height          =   195
         Left            =   1680
         TabIndex        =   10
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Código"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   495
      End
   End
   Begin MSDataGridLib.DataGrid grdCtas 
      CausesValidation=   0   'False
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   3334
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   2
      RowHeight       =   19
      RowDividerStyle =   6
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "cCtaContCod"
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
      BeginProperty Column01 
         DataField       =   "cCtaContDesc"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         BeginProperty Column00 
            ColumnWidth     =   2280.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5309.858
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMntRepColumnasCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim varSiGraba As Boolean
Dim sCtaCod As String
Dim sOpeCod As String, sOpeDesc As String
Dim nNroCol As Integer

Dim EnGrid As Boolean
Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim clsCtaCont As DCtaCont
Dim clsRepCtaCol As DRepCtaColumna

Public Sub Inicio(OpeCod As String, OpeDesc As String, NumCol As Long)
sOpeDesc = OpeDesc
sOpeCod = OpeCod
nNroCol = NumCol
Me.Show 1
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If Len(Trim(txtDesc.Text)) = 0 Then
   MsgBox "La Cuenta Contable no está registrada en el Plan de Cuentas...", vbInformation, "¡Aviso!"
   txtCta.Text = ""
   txtCta.SetFocus
   Exit Function
End If

If Not clsCtaCont.ctaInstancia(txtCta.Text, gsCentralCom & "CtaCont") Then
   If MsgBox(" Cuenta contable elegida NO es de Asiento... ¿ Desea Continuar ? ", vbQuestion + vbYesNo, "Advertencia") = vbNo Then
      txtCta.Text = ""
      txtCta.SetFocus
      Exit Function
   End If
End If
ValidaDatos = True
End Function

Private Sub cmdAceptar_Click()
Dim N As Integer, strGRABA As String, sQRY As String
Dim rt As ADODB.Recordset

If Not ValidaDatos() Then
   Exit Sub
End If
If MsgBox("¿ Esta seguro de asignar la Cuenta al Reporte ? ", vbYesNo, "Confirmación") = vbYes Then
   sCtaCod = txtCta.Text
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   Set clsRepCtaCol = New DRepCtaColumna
   clsRepCtaCol.InsertaRepColumnaCta sOpeCod, nNroCol, sCtaCod, gsMovNro
   Set clsRepCtaCol = Nothing
   varSiGraba = True
   Unload Me
Else
   varSiGraba = False
End If
End Sub

Private Sub cmdbuscar_Click()
frmBuscaCtaGrd.Inicia rsCta, 0, 0
If frmBuscaCtaGrd.lOk Then
   grdCtas.Refresh
End If
grdCtas.SetFocus
End Sub

Private Sub cmdCerrar_Click()
varSiGraba = False
Unload Me
End Sub

Private Sub Form_Load()
EnGrid = False
Set clsCtaCont = New DCtaCont
Set rsCta = clsCtaCont.CargaCtaCont
Set grdCtas.DataSource = rsCta

CentraForm Me
Me.Caption = "Operaciones: Mantenimiento: Asignación de Cuentas"
txtOpeCod.Text = Trim(sOpeCod)
txtOpeDesc.Text = UCase(sOpeDesc)
End Sub

Private Sub grdCtas_GotFocus()
EnGrid = True
End Sub

Private Sub grdCtas_LostFocus()
EnGrid = False
End Sub

Private Sub grdCtas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If EnGrid Then
   If Not rsCta.EOF Then
      txtCta.Text = rsCta!cCtaContCod
      txtDesc.Text = rsCta!cCtaContDesc
   End If
End If
End Sub

Private Sub txtCta_Change()
If Len(Trim(txtCta.Text)) <> 0 Then
   BuscaDato "cCtaContCod like '" & txtCta.Text & "'", rsCta, 1, False
End If
End Sub

Private Sub txtCta_GotFocus()
txtCta.BackColor = "&H00F0FFFF"
End Sub

Private Sub txtCta_LostFocus()
txtCta.BackColor = "&H80000005"
End Sub

Private Sub txtCta_Validate(Cancel As Boolean)
If Not ValidaCuenta() Then
   Cancel = True
End If
End Sub

Private Function ValidaCuenta() As Boolean
ValidaCuenta = True
If Len(Trim(txtCta.Text)) = 0 Then
   MsgBox "Debe indicar una cuenta contable de Asiento", vbCritical, "Error"
   ValidaCuenta = False
   Exit Function
End If
If clsCtaCont.ctaInstancia(txtCta.Text, gsCentralCom & "CtaCont") Then
   txtCta = rsCta!cCtaContCod
   txtDesc = rsCta!cCtaContDesc
   ValidaCuenta = True
ElseIf MsgBox(" Cuenta contable elegida NO es de Asiento... ¿ Desea Continuar ? ", vbQuestion + vbYesNo, "Advertencia") = vbNo Then
   ValidaCuenta = False
Else
   txtCta = rsCta!cCtaContCod
   txtDesc = rsCta!cCtaContDesc
End If
End Function

Private Sub txtCta_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 8 Then
   If InStr("1234567890", Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End If
If KeyAscii = 13 Then
   If ValidaCuenta Then
      cmdAceptar.SetFocus
   End If
End If
End Sub

Public Property Get pSiGraba() As Boolean
pSiGraba = varSiGraba
End Property

Public Property Let pSiGraba(ByVal vNewValue As Boolean)
varSiGraba = vNewValue
End Property

Public Property Get pOpeCod() As String
pOpeCod = sOpeCod
End Property

Public Property Let pOpeCod(ByVal vNewValue As String)
sOpeCod = vNewValue
End Property


