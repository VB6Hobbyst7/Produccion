VERSION 5.00
Object = "{F9AB04EF-FCD4-4161-99E1-9F65F8191D72}#14.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmTarjCtaElim 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminacion de Relacion Tarjeta - Cuenta - F12 para Digitar Tarjeta"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   Icon            =   "frmTarjCtaElim.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCuenta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   120
      TabIndex        =   11
      Top             =   1095
      Width           =   5340
   End
   Begin OCXTarjeta.CtrlTarjeta CtrlTarjeta1 
      Height          =   600
      Left            =   6675
      TabIndex        =   10
      Top             =   1935
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1058
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   30
      TabIndex        =   4
      Top             =   3720
      Width           =   5505
      Begin VB.CommandButton Command2 
         Caption         =   "&Salir"
         Height          =   390
         Left            =   4185
         TabIndex        =   8
         Top             =   225
         Width           =   1215
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar Relacion"
         Height          =   390
         Left            =   75
         TabIndex        =   5
         Top             =   225
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   15
      TabIndex        =   0
      Top             =   30
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
         Left            =   795
         MaxLength       =   16
         TabIndex        =   9
         Top             =   225
         Visible         =   0   'False
         Width           =   3240
      End
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
   Begin VB.Label Label6 
      Caption         =   "Cuenta Seleccionada :"
      Height          =   285
      Left            =   45
      TabIndex        =   7
      Top             =   3405
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   3330
      Width           =   3150
   End
End
Attribute VB_Name = "frmTarjCtaElim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R As ADODB.Recordset
Dim oConec As DConecta


Private Sub Command2_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub lstCuenta_Click()
    Me.LblCta.Caption = Me.lstCuenta.Text
End Sub

Private Sub TxtNumTarj_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            LblNumTarj.Caption = TxtNumTarj.Text
            TxtNumTarj.Visible = False
            Me.LblNumTarj.Visible = True
            Me.CmdLecTarj.Visible = True
            Me.Caption = "Eliminacion de Relacion Tarjeta - Cuenta - F12 para Digitar Tarjeta"
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
            Me.Caption = "Eliminacion de Relacion Tarjeta - Cuenta - F12 para Digitar Tarjeta"
            TxtNumTarj.SetFocus
    End If
End Sub

Private Sub CmdEliminar_Click()

Call GeneraTrama0205(LblCta.Caption, gsBIN, gsCodAge)
    
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 50, LblNumTarj.Caption)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 50, LblCta.Caption)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnActivar", adInteger, adParamInput, , 0)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ActivaRelacionTarjetaCta"
    
    Cmd.Execute
    
    oConec.CierraConexion
End Sub

Private Sub CargaDatos()
Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 50, LblNumTarj.Caption)
    Cmd.Parameters.Append Prm

    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaCtasAfiliadas"
    
    Set R = Cmd.Execute
    lstCuenta.Clear
    Do While Not R.EOF
        lstCuenta.AddItem R!Cuenta
        R.MoveNext
    Loop
    
    oConec.CierraConexion
End Sub

Private Sub CmdLecTarj_Click()
Me.Caption = "Consulta de Tarjeta - PASE LA TARJETA"

LblNumTarj.Caption = Mid(CtrlTarjeta1.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto), 2, 16)

Me.Caption = "Eliminacion de Relacion Tarjeta - Cuenta - F12 para Digitar Tarjeta"
    Call CargaDatos
    

End Sub

Private Sub DGDatos_Click()
    LblCta.Caption = R!Cuenta
    
End Sub
