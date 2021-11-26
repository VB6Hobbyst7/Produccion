VERSION 5.00
Object = "{F9AB04EF-FCD4-4161-99E1-9F65F8191D72}#14.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmAdicTarjetaCta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adición de Tarjeta - Cuenta - F12 para Digitar Tarjeta"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   Icon            =   "frmAdicTarjetaCta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OCXTarjeta.CtrlTarjeta CtrlTarjeta1 
      Height          =   240
      Left            =   7380
      TabIndex        =   23
      Top             =   315
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.Frame Frame2 
      Height          =   4725
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   9705
      Begin VB.ListBox LstCtas 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1740
         Left            =   480
         TabIndex        =   21
         Top             =   480
         Width           =   8655
      End
      Begin VB.Frame Frame4 
         Height          =   1650
         Left            =   105
         TabIndex        =   17
         Top             =   2235
         Width           =   5295
         Begin VB.Frame Frame5 
            Caption         =   "Cuenta Defecto "
            Height          =   765
            Left            =   75
            TabIndex        =   18
            Top             =   720
            Width           =   5145
            Begin VB.ComboBox cboPrio 
               Height          =   315
               ItemData        =   "frmAdicTarjetaCta.frx":030A
               Left            =   2355
               List            =   "frmAdicTarjetaCta.frx":032C
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   285
               Width           =   1230
            End
            Begin VB.Label Label3 
               Caption         =   "Prioridad Retiro en Cascada :"
               Height          =   285
               Left            =   150
               TabIndex        =   24
               Top             =   330
               Width           =   2145
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Seleccionada :"
            Height          =   285
            Left            =   105
            TabIndex        =   20
            Top             =   360
            Width           =   1770
         End
         Begin VB.Label LblCta 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   1980
            TabIndex        =   19
            Top             =   285
            Width           =   3150
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1665
         Left            =   5430
         TabIndex        =   8
         Top             =   2235
         Width           =   4200
         Begin VB.CheckBox ChkActRela 
            Caption         =   "Activar Relación Tarjeta Cuenta"
            Height          =   300
            Left            =   60
            TabIndex        =   26
            Top             =   1290
            Width           =   3675
         End
         Begin VB.Frame Frame6 
            Height          =   555
            Left            =   60
            TabIndex        =   13
            Top             =   135
            Visible         =   0   'False
            Width           =   3975
            Begin VB.OptionButton OptCons 
               Caption         =   "No"
               Height          =   390
               Index           =   1
               Left            =   3075
               TabIndex        =   15
               Top             =   135
               Width           =   555
            End
            Begin VB.OptionButton OptCons 
               Caption         =   "Si"
               Height          =   390
               Index           =   0
               Left            =   2385
               TabIndex        =   14
               Top             =   120
               Value           =   -1  'True
               Width           =   555
            End
            Begin VB.Label Label5 
               Caption         =   "Permitir Consultas"
               Height          =   285
               Left            =   120
               TabIndex        =   16
               Top             =   195
               Width           =   1845
            End
         End
         Begin VB.Frame Frame7 
            Height          =   555
            Left            =   60
            TabIndex        =   9
            Top             =   690
            Visible         =   0   'False
            Width           =   3975
            Begin VB.OptionButton OptCons 
               Caption         =   "Si"
               Height          =   390
               Index           =   3
               Left            =   2370
               TabIndex        =   11
               Top             =   135
               Value           =   -1  'True
               Width           =   555
            End
            Begin VB.OptionButton OptCons 
               Caption         =   "No"
               Height          =   390
               Index           =   2
               Left            =   3075
               TabIndex        =   10
               Top             =   135
               Width           =   555
            End
            Begin VB.Label Label6 
               Caption         =   "Permitir Retiros"
               Height          =   285
               Left            =   120
               TabIndex        =   12
               Top             =   195
               Width           =   1845
            End
         End
      End
      Begin VB.Frame Frame3 
         Height          =   675
         Left            =   45
         TabIndex        =   5
         Top             =   3975
         Width           =   9585
         Begin VB.CommandButton CmdSalir 
            Caption         =   "&Salir"
            Height          =   360
            Left            =   8160
            TabIndex        =   7
            Top             =   195
            Width           =   1350
         End
         Begin VB.CommandButton CmdReg 
            Caption         =   "Actualizar"
            Enabled         =   0   'False
            Height          =   360
            Left            =   75
            TabIndex        =   6
            Top             =   210
            Width           =   1770
         End
      End
      Begin VB.Label Label9 
         Caption         =   "TIPO DE CUENTA"
         Height          =   255
         Left            =   5640
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "PRIORIDAD"
         Height          =   255
         Left            =   3840
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "ESTADO"
         Height          =   255
         Left            =   2520
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "CUENTA"
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frame1 
      Caption         =   "Adicion de Tarjeta - Cuenta"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
         TabIndex        =   22
         Top             =   225
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.CommandButton Tarjeta 
         Caption         =   "Leer Tarjeta"
         Height          =   345
         Left            =   4110
         TabIndex        =   1
         Top             =   240
         Width           =   1290
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
         ForeColor       =   &H00C00000&
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
Attribute VB_Name = "frmAdicTarjetaCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sbIn As String
Dim R As ADODB.Recordset
Dim sResp As String
Dim oConec As DConecta
 

Private Sub CargaDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
    
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
    Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaCtasRegistradas"
    
    Set R = Cmd.Execute
    LstCtas.Clear
    Do While Not R.EOF
         LstCtas.AddItem R!Cuenta & Space(15) & Left(R!cTipoPrograma & Space(20), 20)
        R.MoveNext
    Loop
    
    oConec.CierraConexion
    
End Sub
Private Sub CmdLecTarj_Click()
Me.Caption = "Adición de Tarjeta-Cuenta -> PASE LA TARJETA"

LblNumTarj.Caption = Mid(CtrlTarjeta1.LeerTarjeta("PASE LA TARJETA"), 2, 16)
Me.Caption = "Adición de Tarjeta - Cuenta - F12 para Digitar Tarjeta"
    If Not ExisteTarjeta(LblNumTarj.Caption) Then
        LblNumTarj.Caption = ""
        Exit Sub
    End If
   Call CargaDatos
   
        
End Sub

Private Sub CmdReg_Click()
Dim sResp As String
Dim sTramaResp As String
Dim sDesc1 As String
Dim sDesc2 As String
    
     Dim Cmd As New Command
     Dim Prm As New ADODB.Parameter
     
     Set Prm = New ADODB.Parameter
     Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 50, LblNumTarj.Caption)
     Cmd.Parameters.Append Prm
    
     Set Prm = New ADODB.Parameter
     Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 50, LblCta.Caption)
     Cmd.Parameters.Append Prm
     
     Set Prm = New ADODB.Parameter
     Set Prm = Cmd.CreateParameter("@pnActivar", adInteger, adParamInput, , IIf(ChkActRela.Value = 1, 1, 0))
     Cmd.Parameters.Append Prm
     
     Set Prm = New ADODB.Parameter
     Set Prm = Cmd.CreateParameter("@pnPrio", adInteger, adParamInput, , CInt(Me.cboPrio.Text))
     Cmd.Parameters.Append Prm
     
     Set Prm = New ADODB.Parameter
     Set Prm = Cmd.CreateParameter("@pnConsulta", adInteger, adParamInput, , IIf(Me.OptCons(0).Value, 1, 0))
     Cmd.Parameters.Append Prm
     
     Set Prm = New ADODB.Parameter
     Set Prm = Cmd.CreateParameter("@pnRetiro", adInteger, adParamInput, , IIf(Me.OptCons(3).Value, 1, 0))
     Cmd.Parameters.Append Prm
     
     oConec.AbreConexion
     Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
     Cmd.CommandType = adCmdStoredProc
     Cmd.CommandText = "ATM_ActivaRelacionTarjetaCta"
     
     Cmd.Execute
    
     MsgBox "Se Actualizó la Cuenta a la Tarjeta Correctamente"
     oConec.CierraConexion
    
     Call CargaDatos

End Sub

Private Sub DGDatos_Click()
    LblCta.Caption = R!Cuenta
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub LstCtas_Click()
Dim nPrio As Integer
Dim nRelac As Integer
Dim nCons As Integer
Dim nRetiro As Integer
Dim i As Integer

    Call ConsultaTarjetaCuenta(Me.LblNumTarj.Caption, Mid(Me.LstCtas.Text, 1, 18), nPrio, nRelac, nCons, nRetiro)
    LblCta.Caption = Mid(Me.LstCtas.Text, 1, 18)
    cboPrio.Clear
    
    If nRelac Then
        cboPrio.AddItem Trim(Str(nPrio))
    Else
        cboPrio.AddItem Trim(Str(nPrio + 1))
    End If
        
    'cboPrio.Text = IIf(Trim(Str(nPrio)) = "0", 1, Trim(Str(nPrio)))
    If cboPrio.ListCount > 0 Then
        cboPrio.ListIndex = 0
    End If
    
    cboPrio.Enabled = False
    
    OptCons(0).Value = IIf(nCons = 0, False, True)
    OptCons(1).Value = IIf(nCons = 0, True, False)
    OptCons(3).Value = IIf(nRetiro = 0, False, True)
    OptCons(2).Value = IIf(nRetiro = 0, True, False)
    ChkActRela.Value = IIf(nRelac = 0, 0, 1)
    Me.CmdReg.Enabled = True
    
End Sub


Private Sub Tarjeta_Click()
    Me.Caption = "Adición de Tarjeta - Cuenta - PASE LA TARJETA"
    
    LblNumTarj.Caption = Mid(CtrlTarjeta1.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto), 2, 16)
    Frame3.Enabled = True
    Me.Caption = "Adición de Tarjeta - Cuenta - F12 para Digitar Tarjeta"
    
    If Not ExisteTarjeta(LblNumTarj.Caption) Then
        LblNumTarj.Caption = ""
        Exit Sub
    End If
    
    Call CargaDatos
 
End Sub

Private Sub TxtNumTarj_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            LblNumTarj.Caption = TxtNumTarj.Text
            TxtNumTarj.Visible = False
            Me.LblNumTarj.Visible = True
            Me.Tarjeta.Visible = True
            Me.Caption = "Adición de Tarjeta - Cuenta - F12 para Digitar Tarjeta"
            If Len(Trim(TxtNumTarj.Text)) > 0 Then
                If Not ExisteTarjeta(TxtNumTarj.Text) Then
                  TxtNumTarj.Text = ""
                  Exit Sub
                End If
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
            Me.Tarjeta.Visible = False
            Me.Caption = "Adición de Tarjeta - Cuenta - F12 para Digitar Tarjeta"
            TxtNumTarj.SetFocus
    End If
End Sub
