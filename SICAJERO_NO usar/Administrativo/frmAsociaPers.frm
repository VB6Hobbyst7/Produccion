VERSION 5.00
Object = "{F9AB04EF-FCD4-4161-99E1-9F65F8191D72}#14.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmAsociaPers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asociar Persona - F12 para Digitar Tarjeta"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmAsociaPers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   15
      TabIndex        =   27
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
         TabIndex        =   31
         Top             =   225
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.CommandButton CmdLecTarj 
         Caption         =   "Leer Tarjeta"
         Height          =   345
         Left            =   4065
         TabIndex        =   28
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label Lblnumtarjeta 
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
         TabIndex        =   30
         Top             =   240
         Width           =   3225
      End
      Begin VB.Label Label2 
         Caption         =   "Tarjeta :"
         Height          =   255
         Left            =   60
         TabIndex        =   29
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   30
      TabIndex        =   24
      Top             =   4950
      Width           =   5505
      Begin VB.CommandButton CmdActTar 
         Caption         =   "Asociar Persona"
         Enabled         =   0   'False
         Height          =   375
         Left            =   150
         TabIndex        =   26
         Top             =   180
         Width           =   1305
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   4380
         TabIndex        =   25
         Top             =   150
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Persona"
      Enabled         =   0   'False
      Height          =   4035
      Left            =   0
      TabIndex        =   0
      Top             =   870
      Width           =   5535
      Begin VB.TextBox TxtDirecc 
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   1065
         TabIndex        =   8
         Top             =   3525
         Width           =   4080
      End
      Begin VB.ComboBox CboEstCiv 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmAsociaPers.frx":030A
         Left            =   1065
         List            =   "frmAsociaPers.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3060
         Width           =   1500
      End
      Begin VB.TextBox TxtFecNac 
         Enabled         =   0   'False
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
         Height          =   360
         Left            =   3615
         TabIndex        =   6
         Top             =   2565
         Width           =   1290
      End
      Begin VB.OptionButton OptSex 
         Caption         =   "Femenino"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   4350
         TabIndex        =   5
         Top             =   2175
         Width           =   1035
      End
      Begin VB.OptionButton OptSex 
         Caption         =   "Masculino"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   3150
         TabIndex        =   4
         Top             =   2160
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.TextBox TxtTelef 
         Enabled         =   0   'False
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
         Height          =   360
         Left            =   1065
         TabIndex        =   3
         Top             =   2535
         Width           =   1425
      End
      Begin VB.TextBox TxtDNI 
         Height          =   360
         Left            =   555
         TabIndex        =   2
         Top             =   285
         Width           =   1755
      End
      Begin VB.CommandButton CmdBuscarDNI 
         Caption         =   "Buscar Persona"
         Height          =   390
         Left            =   2445
         TabIndex        =   1
         Top             =   255
         Width           =   1545
      End
      Begin OCXTarjeta.CtrlTarjeta Tarjeta 
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
      End
      Begin VB.Label Label17 
         Caption         =   "Direccion   :"
         Height          =   255
         Left            =   45
         TabIndex        =   23
         Top             =   3600
         Width           =   915
      End
      Begin VB.Label Label16 
         Caption         =   "Estado Civil :"
         Height          =   255
         Left            =   75
         TabIndex        =   22
         Top             =   3090
         Width           =   945
      End
      Begin VB.Label Label14 
         Caption         =   "Fec. Nacim:"
         Height          =   255
         Left            =   2655
         TabIndex        =   21
         Top             =   2610
         Width           =   870
      End
      Begin VB.Label Label15 
         Caption         =   "Telefono :"
         Height          =   255
         Left            =   75
         TabIndex        =   20
         Top             =   2580
         Width           =   840
      End
      Begin VB.Label Label13 
         Caption         =   "Sexo :"
         Height          =   255
         Left            =   2565
         TabIndex        =   19
         Top             =   2145
         Width           =   525
      End
      Begin VB.Label LblDNI 
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
         Left            =   1065
         TabIndex        =   18
         Top             =   2070
         Width           =   1320
      End
      Begin VB.Label Label11 
         Caption         =   "DNI            :"
         Height          =   255
         Left            =   75
         TabIndex        =   17
         Top             =   2115
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "DNI :"
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   345
         Width           =   480
      End
      Begin VB.Label LblNom 
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
         Left            =   1065
         TabIndex        =   15
         Top             =   1635
         Width           =   4200
      End
      Begin VB.Label Label8 
         Caption         =   "Nombres    :"
         Height          =   255
         Left            =   75
         TabIndex        =   14
         Top             =   1710
         Width           =   915
      End
      Begin VB.Label LblApeMat 
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
         Left            =   1065
         TabIndex        =   13
         Top             =   1200
         Width           =   2370
      End
      Begin VB.Label Label6 
         Caption         =   "A. Materno :"
         Height          =   255
         Left            =   75
         TabIndex        =   12
         Top             =   1275
         Width           =   915
      End
      Begin VB.Label LblApePat 
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
         Left            =   1065
         TabIndex        =   11
         Top             =   765
         Width           =   2370
      End
      Begin VB.Label Label4 
         Caption         =   "A. Paterno :"
         Height          =   255
         Left            =   75
         TabIndex        =   10
         Top             =   840
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmAsociaPers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sPerscod As String
Dim oConec As DConecta

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub TxtNumTarj_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            Lblnumtarjeta.Caption = TxtNumTarj.Text
            TxtNumTarj.Visible = False
            Me.Lblnumtarjeta.Visible = True
            Me.CmdLecTarj.Visible = True
            Me.Caption = "Asociar Persona - F12 para Digitar Tarjeta"
            If Len(Trim(Lblnumtarjeta.Caption)) > 0 Then
                Me.Frame3.Enabled = True
                CmdActTar.Enabled = True
                CmdBuscarDNI.SetFocus
            Else
                 Me.Frame3.Enabled = False
                CmdActTar.Enabled = False
            End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 123 Then
            TxtNumTarj.Text = ""
            TxtNumTarj.Visible = True
            Me.Lblnumtarjeta.Visible = False
            Me.CmdLecTarj.Visible = False
            Me.Caption = "Asociar Persona - F12 para Digitar Tarjeta"
            TxtNumTarj.SetFocus
    End If
End Sub

Private Sub CmdActTar_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim sResp As String
Dim sTramaResp As String

    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cNumtarjeta", adVarChar, adParamInput, 16, Lblnumtarjeta.Caption)
    Cmd.Parameters.Append Prm
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cPersCod", adVarChar, adParamInput, 20, sPerscod)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_AsociaPersona"
    Cmd.Execute
    
    'Call CerrarConexion
    oConec.CierraConexion
            
    MsgBox "Tarjeta Asociada"
    
    Call LimpiaDatos
    Me.Frame3.Enabled = False
    CmdActTar.Enabled = False
End Sub

Private Sub CmdBuscarDNI_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter

    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psDNI", adVarChar, adParamInput, 50, TxtDNI.Text)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psApePat", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psApeMat", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNombres", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psSexo", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psTelef", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pdFecNac", adDate, adParamOutput)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psEstCiv", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psDirecc", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psPersCodTar", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarjeta", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaDatosPersona"
    Cmd.Execute
    If Trim(Cmd.Parameters(9).Value) <> "" Then
        Me.LblApePat.Caption = Replace(UCase(Cmd.Parameters(1).Value), "Ñ", "#")
        Me.LblApeMat.Caption = Replace(UCase(Cmd.Parameters(2).Value), "Ñ", "#")
        Me.LblNom.Caption = Replace(UCase(Cmd.Parameters(3).Value), "Ñ", "#")
        Me.OptSex(0).Value = IIf(Cmd.Parameters(4).Value = "M", True, False)
        Me.OptSex(1).Value = IIf(Cmd.Parameters(4).Value = "F", True, False)
        Me.TxtTelef.Text = Cmd.Parameters(5).Value
        Me.TxtFecNac.Text = Cmd.Parameters(6).Value
        Me.CboEstCiv.Text = IIf(Cmd.Parameters(7).Value = "1", "SOLTERO", _
                IIf(Cmd.Parameters(7).Value = "2", "CASADO", _
                IIf(Cmd.Parameters(7).Value = "3", "VIUDO", "DIVORCIADO")))
        LblDNI.Caption = Cmd.Parameters(0).Value
        Me.TxtDirecc.Text = Replace(UCase(Cmd.Parameters(8).Value), "Ñ", "#")
        sPerscod = Cmd.Parameters(9).Value
        CmdActTar.Enabled = True
    Else
        CmdActTar.Enabled = False
    End If
    
    'Call CerrarConexion
    oConec.CierraConexion

    Set Cmd = Nothing
    Set Prm = Nothing

End Sub


Private Sub LimpiaDatos()
    Me.LblApeMat.Caption = ""
    Me.LblApePat.Caption = ""
    Me.LblDNI.Caption = ""
    Me.LblNom.Caption = ""
    Me.Lblnumtarjeta.Caption = ""
    Me.TxtDirecc.Text = ""
    Me.TxtDNI.Text = ""
    Me.TxtFecNac.Text = ""
    Me.TxtTelef.Text = ""

End Sub

Private Sub CmdLecTarj_Click()
Me.Caption = "Asociar Persona - PASE LA TARJETA"

Lblnumtarjeta.Caption = Mid(Tarjeta.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto), 2, 16)
Frame3.Enabled = True
Me.Caption = "Asociar Persona - F12 para Digitar Tarjeta"

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
    
    Set oConec = New DConecta
    
    Set Cmd = New ADODB.Command
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaEstadoCivil"
    Set R = Cmd.Execute
    Me.CboEstCiv.Clear
    
    Do While Not R.EOF
        CboEstCiv.AddItem R!cConsDescripcion
        R.MoveNext
    Loop
    R.Close
    'Call CerrarConexion
    oConec.CierraConexion
End Sub


