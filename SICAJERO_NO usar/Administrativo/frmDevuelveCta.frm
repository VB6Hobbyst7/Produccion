VERSION 5.00
Object = "{F9AB04EF-FCD4-4161-99E1-9F65F8191D72}#14.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmDevuelveCta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Cuenta - F12 para Digitar Tarjeta"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmDevuelveCta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5580
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
         Left            =   825
         MaxLength       =   16
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.CommandButton CmdLecTarj 
         Caption         =   "Leer Tarjeta"
         Height          =   345
         Left            =   4155
         TabIndex        =   10
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label Label2 
         Caption         =   "Tarjeta :"
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   735
      End
      Begin VB.Label LblNumTarj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   825
         TabIndex        =   11
         Top             =   255
         Width           =   3225
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cuentas"
      Height          =   2055
      Left            =   15
      TabIndex        =   2
      Top             =   930
      Width           =   5520
      Begin VB.ListBox LstCtas 
         Height          =   1425
         Left            =   90
         TabIndex        =   3
         Top             =   450
         Width           =   5355
      End
      Begin VB.Label Label5 
         Caption         =   "Moneda"
         Height          =   255
         Left            =   4320
         TabIndex        =   6
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de Cuenta"
         Height          =   255
         Left            =   2370
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Cuenta"
         Height          =   255
         Left            =   660
         TabIndex        =   4
         Top             =   210
         Width           =   660
      End
   End
   Begin VB.Frame Frame3 
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   2985
      Width           =   5550
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   330
         Left            =   4290
         TabIndex        =   8
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   330
         Left            =   60
         TabIndex        =   7
         Top             =   195
         Width           =   1125
      End
   End
   Begin OCXTarjeta.CtrlTarjeta CtrlTarjeta1 
      Height          =   375
      Left            =   5730
      TabIndex        =   0
      Top             =   195
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmDevuelveCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCtaCod As String
Public Function SeleccionarCuenta() As String
    Me.Show 1
    SeleccionarCuenta = sCtaCod
End Function

Private Sub CargaDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
Dim loConec As New DConecta

    Set R = New ADODB.Recordset
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 20, Me.LblNumTarj.Caption)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaCtasParaCajero"
    R.CursorType = adOpenStatic
    R.LockType = adLockReadOnly
    Set R = Cmd.Execute
    LstCtas.Clear
    Do While Not R.EOF
        'LstCtas.AddItem R!Cuenta
        LstCtas.AddItem R!Cuenta & Space(7) & R!TipoCta & Space(15) & IIf(Mid(R!Cuenta, 9, 1) = "1", "SOLES", "DOLARES")
        
        R.MoveNext
    Loop
    loConec.CierraConexion
    Set loConec = Nothing
  
End Sub

Private Sub CmdAceptar_Click()
    If Me.LstCtas.ListIndex >= 0 Then
        sCtaCod = Me.LstCtas.List(Me.LstCtas.ListIndex)
    Else
        sCtaCod = ""
    End If
     Unload Me
End Sub

Private Sub CmdSalir_Click()
sCtaCod = ""
  Unload Me
End Sub

Private Sub TxtNumTarj_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            LblNumTarj.Caption = TxtNumTarj.Text
            TxtNumTarj.Visible = False
            Me.LblNumTarj.Visible = True
            Me.CmdLecTarj.Visible = True
            Me.Caption = "Seleccionar Cuenta  - F12 para Digitar Tarjeta"
            If Len(Trim(LblNumTarj.Caption)) > 0 Then
                Call CargaDatos
            Else

            End If
    End If
End Sub
Private Sub CmdLecTarj_Click()

Me.Caption = "Seleccionar Cuenta - PASE LA TARJETA"

LblNumTarj.Caption = Mid(CtrlTarjeta1.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto), 2, 16)
Me.Caption = "Seleccionar Cuenta  - F12 para Digitar Tarjeta"

Call CargaDatos

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 123 Then
            TxtNumTarj.Text = ""
            TxtNumTarj.Visible = True
            Me.LblNumTarj.Visible = False
            Me.CmdLecTarj.Visible = False
            Me.Caption = "Seleccionar Cuenta  - F12 para Digitar Tarjeta"
            TxtNumTarj.SetFocus
    End If
End Sub


