VERSION 5.00
Object = "{F9AB04EF-FCD4-4161-99E1-9F65F8191D72}#14.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmAfilTarj 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden de Afiliación de Tarjeta"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmAfiliacionTarj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OCXTarjeta.CtrlTarjeta Tarjeta 
      Height          =   150
      Left            =   4410
      TabIndex        =   33
      Top             =   1350
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   265
   End
   Begin VB.Frame Frame3 
      Caption         =   "Persona"
      Height          =   4035
      Left            =   6195
      TabIndex        =   10
      Top             =   285
      Width           =   5340
      Begin VB.TextBox TxtDirecc 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1065
         TabIndex        =   18
         Top             =   3525
         Width           =   4080
      End
      Begin VB.ComboBox CboEstCiv 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmAfiliacionTarj.frx":030A
         Left            =   1065
         List            =   "frmAfiliacionTarj.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3060
         Width           =   1500
      End
      Begin VB.TextBox TxtFecNac 
         Enabled         =   0   'False
         Height          =   360
         Left            =   3615
         TabIndex        =   16
         Top             =   2565
         Width           =   1290
      End
      Begin VB.OptionButton OptSex 
         Caption         =   "Femenino"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   4215
         TabIndex        =   15
         Top             =   2175
         Width           =   1035
      End
      Begin VB.OptionButton OptSex 
         Caption         =   "Masculino"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   3150
         TabIndex        =   14
         Top             =   2160
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.TextBox TxtTelef 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1065
         TabIndex        =   13
         Top             =   2535
         Width           =   1425
      End
      Begin VB.TextBox TxtDNI 
         Height          =   360
         Left            =   555
         TabIndex        =   12
         Top             =   285
         Width           =   1755
      End
      Begin VB.CommandButton CmdBuscarDNI 
         Caption         =   "Buscar Persona"
         Height          =   390
         Left            =   2445
         TabIndex        =   11
         Top             =   255
         Width           =   1545
      End
      Begin VB.Label Label17 
         Caption         =   "Direccion   :"
         Height          =   255
         Left            =   45
         TabIndex        =   19
         Top             =   3600
         Width           =   915
      End
      Begin VB.Label Label16 
         Caption         =   "Estado Civil :"
         Height          =   255
         Left            =   75
         TabIndex        =   32
         Top             =   3090
         Width           =   945
      End
      Begin VB.Label Label14 
         Caption         =   "Fec. Nacim:"
         Height          =   255
         Left            =   2655
         TabIndex        =   31
         Top             =   2610
         Width           =   870
      End
      Begin VB.Label Label15 
         Caption         =   "Telefono :"
         Height          =   255
         Left            =   75
         TabIndex        =   30
         Top             =   2580
         Width           =   840
      End
      Begin VB.Label Label13 
         Caption         =   "Sexo :"
         Height          =   255
         Left            =   2565
         TabIndex        =   29
         Top             =   2145
         Width           =   525
      End
      Begin VB.Label LblDNI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1065
         TabIndex        =   28
         Top             =   2070
         Width           =   1320
      End
      Begin VB.Label Label11 
         Caption         =   "DNI            :"
         Height          =   255
         Left            =   75
         TabIndex        =   27
         Top             =   2115
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "DNI :"
         Height          =   255
         Left            =   90
         TabIndex        =   26
         Top             =   345
         Width           =   480
      End
      Begin VB.Label LblNom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1065
         TabIndex        =   25
         Top             =   1635
         Width           =   4200
      End
      Begin VB.Label Label8 
         Caption         =   "Nombres    :"
         Height          =   255
         Left            =   75
         TabIndex        =   24
         Top             =   1710
         Width           =   915
      End
      Begin VB.Label LblApeMat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1065
         TabIndex        =   23
         Top             =   1200
         Width           =   2370
      End
      Begin VB.Label Label6 
         Caption         =   "A. Materno :"
         Height          =   255
         Left            =   75
         TabIndex        =   22
         Top             =   1275
         Width           =   915
      End
      Begin VB.Label LblApePat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1065
         TabIndex        =   21
         Top             =   765
         Width           =   2370
      End
      Begin VB.Label Label4 
         Caption         =   "A. Paterno :"
         Height          =   255
         Left            =   75
         TabIndex        =   20
         Top             =   840
         Width           =   915
      End
   End
   Begin VB.CommandButton CmdLecTarj 
      Caption         =   "Leer Tarjeta"
      Height          =   345
      Left            =   10965
      TabIndex        =   6
      Top             =   4545
      Width           =   1290
   End
   Begin VB.TextBox txtPan 
      Height          =   375
      Left            =   2490
      MaxLength       =   15
      TabIndex        =   5
      Top             =   375
      Width           =   1650
   End
   Begin VB.Frame Frame2 
      Height          =   720
      Left            =   75
      TabIndex        =   1
      Top             =   1020
      Width           =   4320
      Begin VB.CommandButton Command2 
         Caption         =   "PIN"
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   195
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   405
         Left            =   1575
         TabIndex        =   3
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton CmdAfil 
         Caption         =   "Ordenar tarjeta"
         Height          =   405
         Left            =   105
         TabIndex        =   2
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Tarjeta"
      Height          =   930
      Left            =   150
      TabIndex        =   0
      Top             =   45
      Width           =   4305
      Begin VB.Label Label1 
         Caption         =   "Numero de Tarjeta a Ordenar :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Tarjeta :"
      Height          =   255
      Left            =   6960
      TabIndex        =   8
      Top             =   4590
      Width           =   735
   End
   Begin VB.Label LblNumTarj 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7695
      TabIndex        =   7
      Top             =   4530
      Width           =   3225
   End
End
Attribute VB_Name = "frmAfilTarj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sPerscod As String
Dim sResp As String

Private Function TarjetaValida(ByRef psMsgVal As String) As Boolean

    TarjetaValida = True
    'Valida Longitud
    If Len(Trim(Me.txtPan.Text)) <> 15 Then
        psMsgVal = "Longitud de Tarjeta Incorrecto"
        TarjetaValida = False
        Exit Function
    End If
    
    'Valida Caracteres Correctos
    If Not IsNumeric(Me.txtPan.Text) Then
        psMsgVal = "Numero de Tarjeta Contiene Caracteres Incorrectos"
        TarjetaValida = False
        Exit Function
    End If
    
    
End Function

Private Sub CmdAfil_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim PAN As String
Dim sMsgVal As String
Dim Trama As String
Dim sTramaResp As String
Dim loConec As New DConecta

    If Not TarjetaValida(sMsgVal) Then
            MsgBox sMsgVal
        Exit Sub
    End If

    PAN = txtPan.Text & DigitoChequeo(txtPan.Text)
    
    If ExisteTarjeta(PAN) Then
        MsgBox "Tarjeta ya Existe", vbInformation
        Exit Sub
    End If
    
    sResp = "00"
    'SI result== OK
    If sResp = "00" Then
    
        Set Cmd = New ADODB.Command
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cNumTarjeta", adVarChar, adParamInput, 50, PAN)
        Cmd.Parameters.Append Prm
            
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nCondicion", adInteger, adParamInput, , -1)
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nRetenerTarjeta", adInteger, adParamInput, , 0)
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nCodAge", adInteger, adParamInput, , gsCodAge)
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cUserAfil", adVarChar, adParamInput, 50, gsCodUser)
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@dFecAfil", adDate, adParamInput, , Now)
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cPersCod", adVarChar, adParamInput, 50, sPerscod)
        Cmd.Parameters.Append Prm
        
        loConec.AbreConexion
        Cmd.ActiveConnection = loConec.ConexionActiva 'AbrirConexion
        Cmd.CommandType = adCmdStoredProc
        Cmd.CommandText = "ATM_RegistraTarjeta"
        Cmd.Execute
        loConec.CierraConexion
        MsgBox "Orden de Afiliación de Tarjeta Con Exito"
        txtPan.Text = ""
    Else
        MsgBox "Error al realizar la Orden de Afiliación de Tarjeta, Consulte con el Area de TI", vbCritical, "AVISO CRITICO"
    End If
    Set loConec = Nothing
End Sub

Private Sub LimpiaPantalla()

    TxtDNI.Text = ""
    Me.LblApeMat.Caption = ""
    Me.LblApePat.Caption = ""
    Me.LblDNI.Caption = ""
    Me.LblNom.Caption = ""
    Me.LblNumTarj.Caption = ""
    Me.TxtDirecc.Text = ""
    Me.TxtFecNac.Text = ""
    Me.TxtTelef.Text = ""
    
End Sub

Private Sub CmdLecTarj_Click()
    Me.CmdAfil.Enabled = True
    
    LblNumTarj.Caption = Tarjeta.LeerTarjeta("PASE LA TARJETA")
        
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub


Private Sub Command2_Click()
    LblNumTarj.Caption = Tarjeta.PedirPinDes(Me.txtPan.Text, gNMKPOS, gWKPOS)
End Sub



