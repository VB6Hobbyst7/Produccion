VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGarantiaValorDirectoTotal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VALOR DIRECTO TOTAL"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   Icon            =   "frmGarantiaValorDirectoTotal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "Aceptar"
      Top             =   2480
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2265
      TabIndex        =   4
      ToolTipText     =   "Cancelar"
      Top             =   2480
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4048
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Monto Único"
      TabPicture(0)   =   "frmGarantiaValorDirectoTotal.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraDatos 
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
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3975
         Begin VB.TextBox txtGlosa 
            Height          =   570
            Left            =   840
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Tag             =   "txtPrincipal"
            Top             =   1080
            Width           =   2970
         End
         Begin VB.TextBox txtMonto 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   840
            MaxLength       =   15
            TabIndex        =   1
            Top             =   720
            Width           =   1680
         End
         Begin VB.ComboBox cmbMoneda 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   340
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Glosa:"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "VRA:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   630
         End
      End
   End
End
Attribute VB_Name = "frmGarantiaValorDirectoTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************************
'** Nombre : frmGarantiaValorDirectoTotal
'** Descripción : Para registro/edición/consulta de valuaciones Directas segun TI-ERS063-2014
'** Creación : EJVG, 20150108 05:00:00 PM
'********************************************************************************************
Option Explicit
Dim fbRegistrar As Boolean
Dim fbEditar As Boolean
Dim fbConsultar As Boolean

Dim fbOk As Boolean
Dim fbPrimero As Boolean
Dim fnMoneda As Moneda
Dim fvValorDirectoTotal As tValorDirectoTotal
Dim fvValorDirectoTotal_ULT_VAL As tValorDirectoTotal
Dim fsGlosa As String

Private Sub cmbMoneda_Click()
    Dim lnMoneda As Moneda
    lnMoneda = val(Trim(Right(cmbMoneda.Text, 3)))
    If lnMoneda = gMonedaNacional Then
        txtMonto.BackColor = &H80000005
    ElseIf lnMoneda = gMonedaExtranjera Then
        txtMonto.BackColor = &HC0FFC0
    End If
End Sub
Private Sub CmbMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtMonto
    End If
End Sub
Private Sub CmdAceptar_Click()
    On Error GoTo ErrAceptar
    If cmbMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la Moneda", vbInformation, "Aviso"
        EnfocaControl cmbMoneda
        Exit Sub
    End If
    If Not IsNumeric(Trim(txtMonto.Text)) Then
        MsgBox "Ud. debe ingresar el Monto", vbInformation, "Aviso"
        EnfocaControl txtMonto
        Exit Sub
    Else
        If CCur(Trim(txtMonto.Text)) <= 0 Then
            MsgBox "Ud. debe ingresar un Monto mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtMonto
            Exit Sub
        End If
    End If
    If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Glosa", vbInformation, "Aviso"
        EnfocaControl txtGlosa
        Exit Sub
    End If
    fnMoneda = Trim(Right(cmbMoneda.Text, 3))
    fsGlosa = Trim(txtGlosa.Text)
    fvValorDirectoTotal.nVRM = CCur(txtMonto.Text)
    fbOk = True
    Unload Me
    Exit Sub
ErrAceptar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdCancelar_Click()
    fbOk = False
    Unload Me
End Sub
Public Function Registrar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef psGlosa As String, ByRef pvValorDirectoTotal As tValorDirectoTotal, ByRef pvValorDirectoTotal_ULT_VAL As tValorDirectoTotal) As Boolean
    fbRegistrar = True
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fsGlosa = psGlosa
    fvValorDirectoTotal = pvValorDirectoTotal
    fvValorDirectoTotal_ULT_VAL = pvValorDirectoTotal_ULT_VAL
    Show 1
    pnMoneda = fnMoneda
    psGlosa = fsGlosa
    pvValorDirectoTotal = fvValorDirectoTotal
    
    Registrar = fbOk
End Function
Public Function Editar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef psGlosa As String, ByRef pvValorDirectoTotal As tValorDirectoTotal, ByRef pvValorDirectoTotal_ULT_VAL As tValorDirectoTotal) As Boolean
    fbEditar = True
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fsGlosa = psGlosa
    fvValorDirectoTotal = pvValorDirectoTotal
    fvValorDirectoTotal_ULT_VAL = pvValorDirectoTotal_ULT_VAL
    Show 1
    pnMoneda = fnMoneda
    psGlosa = fsGlosa
    pvValorDirectoTotal = fvValorDirectoTotal
    
    Editar = fbOk
End Function
Public Sub Consultar(ByVal pnMoneda As Moneda, ByVal psGlosa As String, ByRef pvValorDirectoTotal As tValorDirectoTotal)
    fbConsultar = True
    fnMoneda = pnMoneda
    fsGlosa = psGlosa
    fvValorDirectoTotal = pvValorDirectoTotal
    Show 1
End Sub
Private Sub Form_Load()
    fbOk = False
    Screen.MousePointer = 11
    
    CargarControles
    LimpiarControles

    If fbEditar Or fbConsultar Then
        cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
        txtMonto.Text = Format(fvValorDirectoTotal.nVRM, "#,##0.00")
        txtGlosa.Text = fsGlosa
        If fbConsultar Then
            FraDatos.Enabled = False
            cmdAceptar.Enabled = False
        End If
    End If
    
    If fbPrimero Then
       fnMoneda = gMonedaNacional
    End If
    cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
    cmbMoneda.Enabled = False
    
    If fbRegistrar Then
        txtMonto.Text = Format(fvValorDirectoTotal_ULT_VAL.nVRM, "#,##0.00")
    End If
    
    If fbRegistrar Then
        Caption = "VALOR DIRECTO TOTAL [ NUEVO ]"
    End If
    If fbConsultar Then
        Caption = "VALOR DIRECTO TOTAL [ CONSULTAR ]"
    End If
    If fbEditar Then
        Caption = "VALOR DIRECTO TOTAL [ EDITAR ]"
    End If
    '***JGPA SEGUN ACTA 130-2018
    Clipboard.Clear
    Clipboard.SetText ("")
    '***End JGPA
    
    Screen.MousePointer = 0
End Sub
Private Sub CargarControles()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Dim rsMoneda As New ADODB.Recordset
    
    Set rsMoneda = oCons.RecuperaConstantes(1011)
    Call Llenar_Combo_con_Recordset(rsMoneda, cmbMoneda)
    
    RSClose rsMoneda
    Set oCons = Nothing
End Sub
Private Sub LimpiarControles()
    cmbMoneda.ListIndex = -1
    txtMonto.Text = "0.00"
    txtGlosa.Text = ""
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        KeyAscii = 0
        EnfocaControl cmdAceptar
    End If
End Sub
Private Sub txtGlosa_LostFocus()
    txtGlosa.Text = UCase(txtGlosa.Text)
End Sub
Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub
Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 15)
    If KeyAscii = 13 Then
        EnfocaControl txtGlosa
    End If
End Sub
Private Sub txtMonto_LostFocus()
    txtMonto.Text = Format(txtMonto.Text, "#,##0.00")
End Sub
Private Sub txtMonto_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub
