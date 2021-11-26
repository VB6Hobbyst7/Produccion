VERSION 5.00
Object = "{F9AB04EF-FCD4-4161-99E1-9F65F8191D72}#14.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmCtaCons 
   Caption         =   "Consulta de una Cuenta  - F12 para Digitar Tarjeta"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   Icon            =   "frmCtaCons.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNumTarj 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   810
      MaxLength       =   16
      TabIndex        =   19
      Top             =   225
      Visible         =   0   'False
      Width           =   3240
   End
   Begin VB.Frame Frame4 
      Height          =   1620
      Left            =   30
      TabIndex        =   9
      Top             =   3105
      Width           =   5550
      Begin VB.Label Label9 
         Caption         =   "Monto :"
         Height          =   270
         Left            =   120
         TabIndex        =   24
         Top             =   1185
         Width           =   1035
      End
      Begin VB.Label LblMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1245
         TabIndex        =   23
         Top             =   1170
         Width           =   1740
      End
      Begin VB.Label LblEstado 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1020
         Left            =   3255
         TabIndex        =   18
         Top             =   300
         Width           =   2025
      End
      Begin VB.Label LblMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1245
         TabIndex        =   17
         Top             =   870
         Width           =   1740
      End
      Begin VB.Label Label4 
         Caption         =   "Moneda :"
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   870
         Width           =   1035
      End
      Begin VB.Label LblTipoCta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1245
         TabIndex        =   15
         Top             =   540
         Width           =   1740
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Cta :"
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label LblUltRet 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1245
         TabIndex        =   13
         Top             =   210
         Width           =   1740
      End
      Begin VB.Label Label1 
         Caption         =   "Ultimo Retiro :"
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   210
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Height          =   810
      Left            =   45
      TabIndex        =   5
      Top             =   4680
      Width           =   5580
      Begin VB.CommandButton CmdCons 
         Caption         =   "Consultar"
         Height          =   390
         Left            =   60
         TabIndex        =   11
         Top             =   225
         Width           =   1560
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "Nueva Consulta"
         Height          =   390
         Left            =   1695
         TabIndex        =   7
         Top             =   225
         Width           =   1560
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   390
         Left            =   4320
         TabIndex        =   6
         Top             =   225
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cuentas"
      Height          =   2205
      Left            =   15
      TabIndex        =   4
      Top             =   885
      Width           =   5595
      Begin OCXTarjeta.CtrlTarjeta CtrlTarjeta1 
         Height          =   480
         Left            =   4665
         TabIndex        =   10
         Top             =   2145
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   847
      End
      Begin VB.ListBox LstCtas 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   90
         TabIndex        =   8
         Top             =   450
         Width           =   5355
      End
      Begin VB.Label Label7 
         Caption         =   "Cuenta"
         Height          =   255
         Left            =   660
         TabIndex        =   22
         Top             =   210
         Width           =   660
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de Cuenta"
         Height          =   255
         Left            =   2370
         TabIndex        =   21
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Moneda"
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         Top             =   180
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   5580
      Begin VB.CommandButton CmdLecTarj 
         Caption         =   "Leer Tarjeta"
         Height          =   345
         Left            =   4065
         TabIndex        =   1
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label LblNumTarj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   795
         TabIndex        =   3
         Top             =   240
         Width           =   3225
      End
      Begin VB.Label Label2 
         Caption         =   "Tarjeta :"
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCtaCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sbIn As String
Dim sResp As String
Dim R As ADODB.Recordset
Dim oConec As DConecta

Private Sub Form_Load()
    Set oConec = New DConecta
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub TxtNumTarj_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            LblNumTarj.Caption = TxtNumTarj.Text
            TxtNumTarj.Visible = False
            Me.LblNumTarj.Visible = True
            Me.CmdLecTarj.Visible = True
            Me.Caption = "Consulta de una Cuenta  - F12 para Digitar Tarjeta"
            If Len(Trim(LblNumTarj.Caption)) > 0 Then
                Call CargaDatos
            Else

            End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 123 Then
            TxtNumTarj.Text = ""
            TxtNumTarj.Visible = True
            Me.LblNumTarj.Visible = False
            Me.CmdLecTarj.Visible = False
            Me.Caption = "Consulta de una Cuenta  - F12 para Digitar Tarjeta"
            TxtNumTarj.SetFocus
    End If
End Sub


Private Sub CmdCons_Click()
Dim sResp As String
Dim sTramaResp As String
Dim sFecha As String
Dim nMonto As Double
    
    sResp = "00"
    If sResp = "00" Then
    
        If Mid(Me.LstCtas.Text, 9, 1) = "1" Then
            LblMoneda.Caption = "SOLES"
        Else
            LblMoneda.Caption = "DOLARES"
        End If
        
        If Mid(Me.LstCtas.Text, 6, 3) = "232" Then
            LblTipoCta.Caption = "AHORROS"
        End If
        If Mid(Me.LstCtas.Text, 6, 3) = "233" Then
            LblTipoCta.Caption = "PLAZO FIJO"
        End If
        If Mid(Me.LstCtas.Text, 6, 3) = "234" Then
            LblTipoCta.Caption = "CTS"
        End If
        
        Call UltimoRetiroAtm(Me.LblNumTarj.Caption, Mid(Me.LstCtas.Text, 1, 18), sFecha, nMonto)
        LblUltRet.Caption = sFecha
        LblMonto.Caption = Format(nMonto, "#,0.00")
        
        If CuentaVinculada(Me.LblNumTarj.Caption, Mid(Me.LstCtas.Text, 1, 18)) Then
            LblEstado.Caption = "VINCULADA"
        Else
            LblEstado.Caption = "NO VINCULADA"
        End If
    Else
        Me.LblMoneda.Caption = ""
        Me.LblUltRet.Caption = ""
        Me.LblTipoCta.Caption = ""
        LblEstado.Caption = "NO REGISTRADA"
    End If
End Sub
Private Sub CargaDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter

    If Not ExisteTarjeta(Me.LblNumTarj.Caption) Then
        MsgBox "Tarjeta No Existe", vbInformation, "Aviso"
        Exit Sub
    End If

    If RecuperaEstadoDETarjeta(Me.LblNumTarj.Caption) <> 1 Then
        MsgBox "Tarjeta No Esta Activa", vbInformation, "Aviso"
        Set Cmd = Nothing
        Set Prm = Nothing
        Exit Sub
    End If

    Set R = New ADODB.Recordset
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 20, Me.LblNumTarj.Caption)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaCtasTitularTotal"
    R.CursorType = adOpenStatic
    R.LockType = adLockReadOnly
    Set R = Cmd.Execute
    LstCtas.Clear
    Do While Not R.EOF
        'LstCtas.AddItem R!Cuenta
        LstCtas.AddItem R!Cuenta & Space(7) & R!TipoCta & Space(15) & IIf(Mid(R!Cuenta, 9, 1) = "1", "SOLES", "DOLARES")
        
        R.MoveNext
    Loop
    
    CmdCons.SetFocus
    oConec.CierraConexion
End Sub
Private Sub CmdLecTarj_Click()

Me.Caption = "Consulta de una Cuenta - PASE LA TARJETA"

LblNumTarj.Caption = Mid(CtrlTarjeta1.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto), 2, 16)
Me.Caption = "Consulta de una Cuenta  - F12 para Digitar Tarjeta"

Call CargaDatos

End Sub

Private Sub cmdNuevo_Click()
   
    LblNumTarj.Caption = ""
    Me.LblMoneda.Caption = ""
    Me.LblTipoCta.Caption = ""
    Me.LblUltRet.Caption = ""
    Me.LblEstado.Caption = ""
    Me.LstCtas.Clear
    LblMonto.Caption = "0.00"
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub


