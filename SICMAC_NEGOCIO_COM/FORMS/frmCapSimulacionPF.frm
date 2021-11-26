VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCapSimulacionPF 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   Icon            =   "frmCapSimulacionPF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosBasicos 
      Caption         =   "Datos Básicos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2490
      Left            =   105
      TabIndex        =   19
      Top             =   0
      Width           =   6255
      Begin VB.Frame fraChqVal 
         Height          =   670
         Left            =   2835
         TabIndex        =   24
         Top             =   210
         Width           =   3270
         Begin MSMask.MaskEdBox txtFecVal 
            Height          =   330
            Left            =   1365
            TabIndex        =   2
            Top             =   225
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Valorizacion :"
            Height          =   195
            Left            =   210
            TabIndex        =   25
            Top             =   293
            Width           =   945
         End
      End
      Begin VB.Frame fraDatos 
         Height          =   1485
         Left            =   105
         TabIndex        =   21
         Top             =   840
         Width           =   6000
         Begin VB.CheckBox chkEspecial 
            Alignment       =   1  'Right Justify
            Caption         =   "Tasa Preferencial"
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
            Left            =   90
            TabIndex        =   27
            Top             =   1140
            Width           =   1845
         End
         Begin VB.TextBox txtTasa 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   3825
            TabIndex        =   7
            Top             =   1050
            Width           =   1905
         End
         Begin VB.ComboBox cboMoneda 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3825
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   630
            Width           =   1860
         End
         Begin VB.TextBox txtPlazo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3825
            MaxLength       =   4
            TabIndex        =   4
            Top             =   225
            Width           =   1905
         End
         Begin SICMACT.EditMoney txtCapital 
            Height          =   330
            Left            =   1140
            TabIndex        =   3
            Top             =   225
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox txtFecApe 
            Height          =   330
            Left            =   1155
            TabIndex        =   5
            Top             =   622
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Capital :"
            Height          =   195
            Left            =   105
            TabIndex        =   29
            Top             =   300
            Width           =   570
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   105
            TabIndex        =   28
            Top             =   690
            Width           =   690
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Efectiva Anual"
            Height          =   195
            Left            =   2175
            TabIndex        =   26
            Top             =   1155
            Width           =   1440
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   2985
            TabIndex        =   23
            Top             =   690
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Plazo :"
            Height          =   195
            Left            =   2985
            TabIndex        =   22
            Top             =   300
            Width           =   480
         End
      End
      Begin VB.Frame fraTipo 
         Height          =   670
         Left            =   105
         TabIndex        =   20
         Top             =   180
         Width           =   2640
         Begin VB.OptionButton optTipo 
            Caption         =   "C&heque"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   1365
            TabIndex        =   1
            Top             =   315
            Width           =   960
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "&Efectivo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   210
            TabIndex        =   0
            Top             =   315
            Width           =   1065
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Plazo Fijo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2025
      Left            =   105
      TabIndex        =   12
      Top             =   2520
      Width           =   6255
      Begin VB.Label lblIntAdelant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1275
         TabIndex        =   35
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Interés Adelantado"
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
         Height          =   375
         Left            =   210
         TabIndex        =   34
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Interés Mensual"
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
         Height          =   480
         Left            =   210
         TabIndex        =   33
         Top             =   920
         Width           =   735
      End
      Begin VB.Label lblIntMens 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1275
         TabIndex        =   32
         Top             =   920
         Width           =   1185
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Cancelación"
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
         Left            =   2610
         TabIndex        =   31
         Top             =   1522
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label lblFecCan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   4275
         TabIndex        =   30
         Top             =   1440
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label5 
         Caption         =   "Interés Final Plazo"
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
         Height          =   435
         Left            =   2655
         TabIndex        =   18
         Top             =   920
         Width           =   1155
      End
      Begin VB.Label lblIntFinal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   3900
         TabIndex        =   17
         Top             =   920
         Width           =   1365
      End
      Begin VB.Label lblFecVenc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   3900
         TabIndex        =   16
         Top             =   330
         Width           =   2205
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Vencimiento"
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
         Height          =   480
         Left            =   2655
         TabIndex        =   15
         Top             =   277
         Width           =   1245
      End
      Begin VB.Label lblTasaMens 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1275
         TabIndex        =   14
         Top             =   315
         UseMnemonic     =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "Tasa Efectiva Mensual"
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
         Height          =   585
         Left            =   210
         TabIndex        =   13
         Top             =   210
         UseMnemonic     =   0   'False
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4590
      Width           =   1005
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5355
      TabIndex        =   11
      Top             =   4590
      Width           =   1005
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   1260
      TabIndex        =   9
      Top             =   4590
      Width           =   1005
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4305
      TabIndex        =   10
      Top             =   4590
      Width           =   1005
   End
End
Attribute VB_Name = "frmCapSimulacionPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nMoneda As Moneda
Dim nTasa As Double
Dim nIntPriMes As Double
Dim bInicio As Boolean

Private Sub CalculaInteres()
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim clsMantA As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim nCapital As Double, nPlazo As Long
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set clsMantA = New COMNCaptaGenerales.NCOMCaptaMovimiento
nCapital = txtCapital.value
nTasa = clsMant.GetTasaNominal(CDbl(txtTasa.Text), 360)
nPlazo = CLng(txtPlazo)
If bInicio = False Then
    If optTipo(0).value Then
        lblFecVenc = Format$(DateAdd("d", nPlazo, CDate(txtFecApe)), "dd mmmm yyyy")
    Else
        lblFecVenc = Format$(DateAdd("d", CInt(txtPlazo), CDate(txtFecVal)), "dd mmmm yyyy")
    End If
    lblIntMens = Format$(clsMant.GetInteresPF(nTasa, nCapital, 30), "#,##0.00")
    
    lblIntFinal = Format$(clsMant.GetInteresPF(nTasa, nCapital, nPlazo), "#,##0.00")
    lblIntAdelant = Format$(clsMantA.GetInteres(nCapital, nTasa, nPlazo, TpoCalcIntAdelantado), "#,##0.00")
End If
Set clsMant = Nothing
Set clsMantA = Nothing
End Sub

Private Sub cboMoneda_Click()
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
nMoneda = CLng(Right(cboMoneda.Text, 1))
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
If nMoneda = gMonedaExtranjera Then
    txtCapital.BackColor = &HFF00&
Else
    txtCapital.BackColor = &H80000005
End If
If txtPlazo <> "" Then
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion 'nCapDefinicion
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    nTasa = clsDef.GetCapTasaInteres(gCapPlazoFijo, nMoneda, IIf(chkEspecial.value = vbUnchecked, gCapTasaNormal, gCapTasaPreferencial), val(txtPlazo.Text), val(txtCapital.value), gsCodAge)
    txtTasa = Format$(clsMant.GetTasaEfectiva(nTasa, 360), "#,##0.00")
    lblTasaMens = Format$(clsMant.GetTasaEfectiva(nTasa, 30), "#,##0.0000")
    CalculaInteres
    Set clsDef = Nothing
    Set clsMant = Nothing
End If
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTasa.SetFocus
End If
End Sub

Private Sub chkEspecial_Click()
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
If txtPlazo <> "" Then
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    nTasa = clsDef.GetCapTasaInteres(gCapPlazoFijo, nMoneda, IIf(chkEspecial.value = vbUnchecked, gCapTasaNormal, gCapTasaPreferencial), val(txtPlazo.Text), val(txtCapital.value), gsCodAge)
    
    'nTasa = clsMant.GetCapTasaInteres(gCapPlazoFijo, nMoneda, IIf(chkEspecial.value = vbUnchecked, gCapTasaNormal, gCapTasaPreferencial), Val(txtPlazo.Text), txtCapital.value, gsCodAge)
    txtTasa = Format$(clsMant.GetTasaEfectiva(nTasa, 360), "#,##0.00")
    lblTasaMens = Format$(clsMant.GetTasaEfectiva(nTasa, 30), "#,##0.0000")
    If Trim(txtFecApe) <> "" And IsDate(txtFecApe) And Trim(txtPlazo) <> "" Then
        lblFecVenc = Format$(DateAdd("d", CInt(txtPlazo), CDate(txtFecApe)), "dd mmmm yyyy")
        lblFecCan = Format$(DateAdd("d", CInt(txtPlazo) + 1, CDate(txtFecApe)), "dd mmmm yyyy")
    End If
    CalculaInteres
    Set clsDef = Nothing
    Set clsMant = Nothing
End If
End Sub

Private Sub cmdApply_Click()
If txtCapital.value = 0 Then
    MsgBox "Monto de capital deber ser mayor que cero", vbInformation, "Aviso"
    txtCapital.SetFocus
ElseIf Trim(txtPlazo) = "" Or CInt(txtPlazo) = 0 Then
    MsgBox "Plazo deber ser mayor que cero", vbInformation, "Aviso"
    txtPlazo.SetFocus
ElseIf Trim(txtFecApe) = "" Or Not IsDate(txtFecApe) Then
    MsgBox "Fecha de Apertura no válida", vbInformation, "Aviso"
    txtFecApe.SetFocus
Else
    CalculaInteres
    'txtCapital.SetFocus
End If
cmdPrint.Enabled = True
cmdPrint.SetFocus
fraDatosBasicos.Enabled = False
End Sub

Private Sub cmdNew_Click()
cmdPrint.Enabled = False
txtCapital.value = 0
txtPlazo = "0"
fraDatosBasicos.Enabled = True
txtFecApe = Format$(gdFecSis, "dd/mm/yyyy")
cboMoneda.ListIndex = 0
lblIntMens = ""
lblIntFinal = ""
lblFecVenc = ""
lblFecCan.Caption = ""
txtTasa = ""
lblTasaMens = ""
optTipo(0).value = True
optTipo(0).SetFocus
txtCapital.Enabled = True

End Sub

Private Sub cmdPrint_Click()
Dim clsPrev As previo.clsprevio
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim sCad As String
Dim dFecVal As Date, dFecApe As Date
Dim nDiasVal As Integer
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    If optTipo(0).value Then
        nDiasVal = 0
    Else
        dFecVal = CDate(txtFecVal)
        nDiasVal = DateDiff("d", CDate(txtFecApe), dFecVal)
    End If
'By Capi 05122007 para que se envie como parametro la tasa nominal tal como en la apertura de plazos fijos
'sCad = clsMant.GetPFPlanRetInt(CDate(txtFecApe), CDbl(lblIntMens), CLng(txtPlazo), nMoneda, CDbl(lblIntFinal), txtCapital.value, CDbl(txtTasa), nDiasVal, dFecVal)
sCad = clsMant.GetPFPlanRetInt(CDate(txtFecApe), CDbl(lblIntMens), CLng(txtPlazo), nMoneda, CDbl(lblIntFinal), txtCapital.value, nTasa, nDiasVal, dFecVal, , "10901233", , CDbl(lblIntAdelant))
Set clsMant = Nothing

Set clsPrev = New previo.clsprevio
    clsPrev.Show sCad, "Plazo Fijo", True, , gImpresora
Set clsPrev = Nothing
cmdNew.SetFocus
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Me.Caption = "Simulación Plazo Fijo"
bInicio = True
txtFecApe = Format$(gdFecSis, "dd/mm/yyyy")
cboMoneda.AddItem "Nacional" & Space(50) & gMonedaNacional
cboMoneda.AddItem "Extranjera" & Space(50) & gMonedaExtranjera
cboMoneda.ListIndex = 0
txtPlazo = "0"
lblTasaMens = "0"
cmdPrint.Enabled = False
optTipo(0).value = True
txtTasa = ""
txtCapital.Enabled = True
bInicio = False
End Sub
Private Sub optTipo_Click(Index As Integer)
Select Case Index
    Case 0
        fraChqVal.Visible = False
    Case 1
        fraChqVal.Visible = True
        txtFecVal.SetFocus
End Select
End Sub

Private Sub optTipo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 1 Then
        txtFecVal.SetFocus
    Else
        txtCapital.SetFocus
    End If
End If
End Sub

Private Sub txtCapital_GotFocus()
txtCapital.MarcaTexto
End Sub

Private Sub txtCapital_KeyPress(KeyAscii As Integer)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales

 If KeyAscii = 13 Then
        txtPlazo.SetFocus
 Else
    If txtPlazo <> "" Then
        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
        Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        nTasa = clsDef.GetCapTasaInteres(gCapPlazoFijo, nMoneda, IIf(chkEspecial.value = vbUnchecked, gCapTasaNormal, gCapTasaPreferencial), val(txtPlazo.Text), val(txtCapital.value), gsCodAge)
        
       ' nTasa = clsMant.GetCapTasaInteres(gCapPlazoFijo, nMoneda, IIf(chkEspecial.value = vbUnchecked, gCapTasaNormal, gCapTasaPreferencial), Val(txtPlazo.Text), txtCapital.value, gsCodAge)
        txtTasa = Format$(clsMant.GetTasaEfectiva(nTasa, 360), "#,##0.00")
        'txtTasa = Format$(ConvierteTNAaTEA(nTasa), "#,##0.00")
        lblTasaMens = Format$(clsMant.GetTasaEfectiva(nTasa, 30), "#,##0.0000")
'        CalculaInteres
     
        Set clsDef = Nothing
        Set clsMant = Nothing
   
   
   End If

 End If
 
 
End Sub

Private Sub txtFecApe_Change()
If Trim(txtFecApe) <> "" And IsDate(txtFecApe) And Trim(txtPlazo) <> "" Then
    lblFecVenc = Format$(DateAdd("d", CLng(txtPlazo), CDate(txtFecApe)), "dd mmmm yyyy")
    lblFecCan = Format$(DateAdd("d", CInt(txtPlazo) + 1, CDate(txtFecApe)), "dd mmmm yyyy")
End If
End Sub

Private Sub txtFecApe_GotFocus()
With txtFecApe
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtFecApe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CDate(txtFecApe) >= CDate(gdFecSis) Then
        cboMoneda.SetFocus
    Else
        MsgBox "La Fecha de Apertura no puede ser Menor a la Fecha Actual", vbInformation, "Aviso"
    End If
End If
End Sub

Private Sub txtFecVal_GotFocus()
With txtFecVal
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtFecVal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha = False Then Exit Sub
End If
End Sub

Private Sub txtPlazo_GotFocus()
With txtPlazo
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtPlazo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim nTEA As Double, nTEM As Double
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    
    If txtPlazo <> "" Then
        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        'nTasa = clsMant.GetCapTasaInteres(gCapPlazoFijo, nMoneda, IIf(chkEspecial.value = vbUnchecked, gCapTasaNormal, gCapTasaPreferencial), Val(txtPlazo.Text), txtCapital.value, gsCodAge)
        Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
        Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        If chkEspecial.value = vbUnchecked Then
            nTasa = clsDef.GetCapTasaInteres(gCapPlazoFijo, nMoneda, IIf(chkEspecial.value = vbUnchecked, gCapTasaNormal, gCapTasaPreferencial), val(txtPlazo.Text), val(txtCapital.value), gsCodAge)
            txtTasa = Format$(clsMant.GetTasaEfectiva(nTasa, 360), "#,##0.00")
        
        'txtTasa = Format$(ConvierteTNAaTEA(nTasa), "#,##0.00")
            lblTasaMens = Format$(clsMant.GetTasaEfectiva(nTasa, 30), "#,##0.0000")
            If Trim(txtFecApe) <> "" And IsDate(txtFecApe) And Trim(txtPlazo) <> "" Then
                lblFecVenc = Format$(DateAdd("d", CInt(txtPlazo), CDate(txtFecApe)), "dd mmmm yyyy")
                lblFecCan = Format$(DateAdd("d", CInt(txtPlazo) + 1, CDate(txtFecApe)), "dd mmmm yyyy")
            End If
        Else
           '***Modificado por ELRO el 20121026, según SATI INC1210240014
           nTasa = clsMant.GetTasaNominal(CDbl(txtTasa.Text), 360)
           'lblTasaMens = Format$(clsMant.GetTasaEfectiva(CDbl(txtTasa), 30), "#,##0.0000")
           lblTasaMens = Format$(clsMant.GetTasaEfectiva(CDbl(nTasa), 30), "#,##0.0000")
           '***Fin Modificado por ELRO el 20121026**********************
        End If
        CalculaInteres
        Set clsDef = Nothing
        Set clsMant = Nothing
    End If

    txtFecApe.SetFocus
    Exit Sub
End If
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtTasa_GotFocus()
With txtTasa
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub


Private Sub txtTasa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTasa.Text = Format$(CDbl(txtTasa.Text), "#,##0.00")
    'CalculaInteres
    txtPlazo_KeyPress (13)
    cmdApply.SetFocus
    Exit Sub
End If
KeyAscii = NumerosDecimales(txtTasa, KeyAscii)
End Sub

Public Function ValidaFecha() As Boolean
    ValidaFecha = True
   If IsDate(txtFecVal) Then
        If CDate(txtFecVal) > CDate(gdFecSis) Then
            txtCapital.Enabled = True
            txtCapital.SetFocus
        Else
           ValidaFecha = False
           MsgBox "La Fecha de Valorización tiene que ser Mayor a la Actual", vbInformation, "Aviso"
           Exit Function
        End If
    Else
        ValidaFecha = False
        MsgBox "Formato de fecha no válido", vbInformation, "Aviso"
        txtFecVal.SetFocus
    End If
End Function
