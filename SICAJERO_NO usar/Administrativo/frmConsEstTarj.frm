VERSION 5.00
Object = "{F9AB04EF-FCD4-4161-99E1-9F65F8191D72}#14.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmConsEstTarj 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Estado de Tarjeta - F12 para Digitar Tarjeta"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmConsEstTarj.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1245
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5505
      Begin VB.TextBox TxtDNI 
         Height          =   360
         Left            =   780
         MaxLength       =   8
         TabIndex        =   13
         Top             =   705
         Width           =   1755
      End
      Begin VB.CommandButton CmdBuscarDNI 
         Caption         =   "Buscar Persona"
         Height          =   390
         Left            =   3795
         TabIndex        =   12
         Top             =   735
         Width           =   1545
      End
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
         TabIndex        =   11
         Top             =   225
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.CommandButton CmdLecTarj 
         Caption         =   "Leer Tarjeta"
         Height          =   345
         Left            =   4065
         TabIndex        =   4
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label Label3 
         Caption         =   "DNI :"
         Height          =   255
         Left            =   75
         TabIndex        =   14
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Tarjeta :"
         Height          =   255
         Left            =   60
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   240
         Width           =   3225
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de Tarjeta"
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   1245
      Width           =   5505
      Begin VB.Frame Frame4 
         Caption         =   "Glosa"
         Height          =   2085
         Left            =   120
         TabIndex        =   15
         Top             =   690
         Width           =   5310
         Begin VB.TextBox TxtGlosa 
            Height          =   1770
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   240
            Width           =   5115
         End
      End
      Begin OCXTarjeta.CtrlTarjeta CtrlTarjeta1 
         Height          =   390
         Left            =   4605
         TabIndex        =   10
         Top             =   225
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   688
      End
      Begin VB.Label LblEstado 
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
         Height          =   315
         Left            =   1170
         TabIndex        =   8
         Top             =   300
         Width           =   3045
      End
      Begin VB.Label Label1 
         Caption         =   "Estado :"
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   345
         Width           =   705
      End
   End
   Begin VB.Frame Frame3 
      Height          =   750
      Left            =   30
      TabIndex        =   0
      Top             =   4230
      Width           =   5505
      Begin VB.CommandButton CmdNuevaCons 
         Caption         =   "Nueva Consulta"
         Height          =   360
         Left            =   135
         TabIndex        =   9
         Top             =   255
         Width           =   1635
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   360
         Left            =   4245
         TabIndex        =   1
         Top             =   240
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmConsEstTarj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sResp As String
Dim oConec As DConecta

Private Sub CmdBuscarDNI_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter


   If Len(Trim(Me.TxtDNI.Text)) <> 8 Then
        MsgBox "DNI Incorrecto", vbInformation, "Aviso"
        Me.TxtDNI.SetFocus
        Exit Sub
   End If
                      
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psDNI", adVarChar, adParamInput, 8, Trim(Me.TxtDNI.Text))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarjeta", adVarChar, adParamOutput, 20)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaNumTarjetaDesdeDNI"
    Cmd.Execute
    
    Me.LblNumTarj.Caption = IIf(IsNull(Cmd.Parameters(1).Value), "", Cmd.Parameters(1).Value)
    
    'Call CerrarConexion
    oConec.CierraConexion
            
    If Len(Trim(Me.LblNumTarj.Caption)) = 0 Then
        MsgBox "No se Encontro Ninguna Tarjeta Asociada al DNI", vbInformation, "Aviso"
        Me.LblEstado.Caption = ""
        Exit Sub
    End If
    
    Call CargaDatos
    
End Sub

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
            Me.Caption = "Consulta de Estado de Tarjeta - F12 para Digitar Tarjeta"
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
            Me.Caption = "Consulta de Estado de Tarjeta - F12 para Digitar Tarjeta"
            TxtNumTarj.SetFocus
    End If
End Sub

Public Function RecuperaGlosa() As String
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter

    Set Prm = New ADODB.Parameter
    
    Set Prm = Cmd.CreateParameter("@psPAN", adVarChar, adParamInput, 20)
    Prm.Value = Me.TxtNumTarj.Text
    Cmd.Parameters.Append Prm
     
    Set Prm = Cmd.CreateParameter("@psGlosa", adVarChar, adParamOutput, 2000)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaGLOSACondicionTarjeta "
    Cmd.Execute
    
    RecuperaGlosa = Cmd.Parameters(1).Value


    'Call CerrarConexion
    oConec.CierraConexion

    Set Cmd = Nothing
    Set Prm = Nothing
    


End Function

Private Sub CargaDatos()
Dim nCond As Integer
Dim nRetenerTar As Integer
Dim nNOOpeMonExt As Integer
Dim nSuspOper As Integer
Dim dFecVenc As Date
Dim psDesEstado As String

If Not ExisteTarjeta(Me.LblNumTarj.Caption) Then
    MsgBox "Tarjeta No Existe", vbInformation, "Aviso"
    Exit Sub
End If

Call RecuperaDatosDETarjetas(Me.LblNumTarj.Caption, nCond, nRetenerTar, nNOOpeMonExt, nSuspOper, dFecVenc, psDesEstado)



Me.txtGlosa.Text = RecuperaGlosa()

Me.LblEstado.Caption = psDesEstado


End Sub

Private Sub CmdLecTarj_Click()

    Call CmdNuevaCons_Click
    Me.Caption = "Consulta de Estado de Tarjeta - PASE LA TARJETA"
    LblNumTarj.Caption = Mid(CtrlTarjeta1.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto), 2, 16)
    Call CargaDatos
    
End Sub

Private Sub CmdNuevaCons_Click()
    LblNumTarj.Caption = ""
    Me.LblEstado.Caption = ""
    Me.TxtDNI.Text = ""
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub
