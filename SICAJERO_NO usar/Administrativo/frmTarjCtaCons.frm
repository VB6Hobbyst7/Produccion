VERSION 5.00
Object = "{F9AB04EF-FCD4-4161-99E1-9F65F8191D72}#14.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmTarjCtaCons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Relación Tarjeta - Cuenta"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   Icon            =   "frmTarjCtaCons.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Height          =   1665
      Left            =   5325
      TabIndex        =   15
      Top             =   1440
      Width           =   4200
      Begin VB.Frame Frame7 
         Height          =   555
         Left            =   60
         TabIndex        =   21
         Top             =   690
         Visible         =   0   'False
         Width           =   3975
         Begin VB.OptionButton OptCons 
            Caption         =   "No"
            Height          =   390
            Index           =   2
            Left            =   3075
            TabIndex        =   23
            Top             =   135
            Width           =   555
         End
         Begin VB.OptionButton OptCons 
            Caption         =   "Si"
            Height          =   390
            Index           =   3
            Left            =   2370
            TabIndex        =   22
            Top             =   135
            Value           =   -1  'True
            Width           =   555
         End
         Begin VB.Label Label6 
            Caption         =   "Permitir Retiros"
            Height          =   285
            Left            =   120
            TabIndex        =   24
            Top             =   195
            Width           =   1845
         End
      End
      Begin VB.Frame Frame6 
         Height          =   555
         Left            =   60
         TabIndex        =   17
         Top             =   135
         Visible         =   0   'False
         Width           =   3975
         Begin VB.OptionButton OptCons 
            Caption         =   "Si"
            Height          =   390
            Index           =   0
            Left            =   2385
            TabIndex        =   19
            Top             =   120
            Value           =   -1  'True
            Width           =   555
         End
         Begin VB.OptionButton OptCons 
            Caption         =   "No"
            Height          =   390
            Index           =   1
            Left            =   3075
            TabIndex        =   18
            Top             =   135
            Width           =   555
         End
         Begin VB.Label Label5 
            Caption         =   "Permitir Consultas"
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   195
            Width           =   1845
         End
      End
      Begin VB.CheckBox ChkActRela 
         Caption         =   "Activar Relación Tarjeta Cuenta"
         Height          =   300
         Left            =   60
         TabIndex        =   16
         Top             =   1290
         Width           =   3675
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1650
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   5295
      Begin VB.Frame Frame5 
         Caption         =   "Cuenta Defecto "
         Height          =   765
         Left            =   75
         TabIndex        =   11
         Top             =   720
         Width           =   5145
         Begin VB.Label LblPrio 
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
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   2355
            TabIndex        =   25
            Top             =   255
            Width           =   645
         End
         Begin VB.Label Label3 
            Caption         =   "Prioridad Retiro en Cascada :"
            Height          =   285
            Left            =   150
            TabIndex        =   12
            Top             =   330
            Width           =   2145
         End
      End
      Begin VB.Label LblCta 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1980
         TabIndex        =   14
         Top             =   285
         Width           =   3150
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Seleccionada :"
         Height          =   285
         Left            =   105
         TabIndex        =   13
         Top             =   360
         Width           =   1770
      End
   End
   Begin VB.Frame Frame3 
      Height          =   720
      Left            =   75
      TabIndex        =   5
      Top             =   3150
      Width           =   9615
      Begin VB.CommandButton CmdNewCons 
         Caption         =   "Nueva Consulta"
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   210
         Width           =   1320
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   360
         Left            =   8055
         TabIndex        =   6
         Top             =   210
         Width           =   1320
      End
   End
   Begin OCXTarjeta.CtrlTarjeta CtrlTarjeta1 
      Height          =   450
      Left            =   6555
      TabIndex        =   4
      Top             =   315
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   794
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5505
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
         Left            =   780
         MaxLength       =   16
         TabIndex        =   9
         Top             =   225
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.ComboBox CboCtas 
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   810
         Width           =   5175
      End
      Begin VB.CommandButton CmdLecTarj 
         Caption         =   "Leer Tarjeta"
         Height          =   345
         Left            =   4065
         TabIndex        =   1
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label Label2 
         Caption         =   "Tarjeta :"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   300
         Width           =   735
      End
      Begin VB.Label LblNumTarj 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   795
         TabIndex        =   2
         Top             =   240
         Width           =   3225
      End
   End
End
Attribute VB_Name = "frmTarjCtaCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sResp As String
Dim MatDatos() As String
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
            Me.Caption = "Consulta de Relación Tarjeta - Cuenta  - F12 para Digitar Tarjeta"
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
            Me.Caption = "Consulta de Relación Tarjeta - Cuenta  - F12 para Digitar Tarjeta"
            TxtNumTarj.SetFocus
    End If
End Sub

Private Sub CboCtas_Click()
    
    Dim nPrio As Integer
Dim nRelac As Integer
Dim nCons As Integer
Dim nRetiro As Integer

    Call ConsultaTarjetaCuenta(Me.LblNumTarj.Caption, Mid(Me.CboCtas.Text, 1, 18), nPrio, nRelac, nCons, nRetiro)
    LblCta.Caption = Mid(Me.CboCtas.Text, 1, 18)
    LblPrio.Caption = Trim(Str(nPrio))
    OptCons(0).Value = IIf(nCons = 0, False, True)
    OptCons(1).Value = IIf(nCons = 0, True, False)
    OptCons(3).Value = IIf(nRetiro = 0, False, True)
    OptCons(2).Value = IIf(nRetiro = 0, True, False)
    ChkActRela.Value = IIf(nRelac = 0, 0, 1)
    
End Sub
Private Sub CargaDatos()
 
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim R As ADODB.Recordset
    
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
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 50, LblNumTarj.Caption)
    Cmd.Parameters.Append Prm

    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaCtasRelacionadas"
    
    Set R = Cmd.Execute
    CboCtas.Clear
    Do While Not R.EOF
         CboCtas.AddItem R!Cuenta '& Space(100) & R!cAgeDesCorta
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    oConec.CierraConexion
    
End Sub
Private Sub CmdLecTarj_Click()
    Me.Caption = "Consulta de Relación Tarjeta-Cuenta -> PASE LA TARJETA"
    
    LblNumTarj.Caption = Mid(CtrlTarjeta1.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto), 2, 16)
    Me.Caption = "Consulta de Relación Tarjeta - Cuenta  - F12 para Digitar Tarjeta"
 
    Call CargaDatos
    
End Sub

Private Sub CmdNewCons_Click()
    LblNumTarj.Caption = ""
    Me.CboCtas.Clear
    LblCta.Caption = ""
    Me.LblPrio.Caption = "0"
    ChkActRela.Value = 0
End Sub

Private Sub Command1_Click()
 Unload Me
End Sub
