VERSION 5.00
Object = "{F9AB04EF-FCD4-4161-99E1-9F65F8191D72}#14.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmCambSitu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bloqueo de Tarjeta - F12 para Digitar Tarjeta"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   Icon            =   "frmCambSitu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Glosa"
      Height          =   1515
      Left            =   30
      TabIndex        =   40
      Top             =   3720
      Width           =   9270
      Begin VB.TextBox TxtGlosa 
         Height          =   1275
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   180
         Width           =   9045
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos de Tarjeta"
      Height          =   2430
      Left            =   15
      TabIndex        =   14
      Top             =   1230
      Width           =   9255
      Begin VB.Label Label8 
         Caption         =   "DNI :"
         Height          =   255
         Left            =   5490
         TabIndex        =   36
         Top             =   2025
         Width           =   1740
      End
      Begin VB.Label LblDNI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7290
         TabIndex        =   35
         Top             =   2010
         Width           =   1425
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   210
         TabIndex        =   34
         Top             =   300
         Width           =   1410
      End
      Begin VB.Label LblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2010
         TabIndex        =   33
         Top             =   285
         Width           =   3390
      End
      Begin VB.Label Label5 
         Caption         =   "Direccion :"
         Height          =   255
         Left            =   210
         TabIndex        =   32
         Top             =   630
         Width           =   1410
      End
      Begin VB.Label LblDirecc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2010
         TabIndex        =   31
         Top             =   615
         Width           =   3405
      End
      Begin VB.Label Label7 
         Caption         =   "Telefono :"
         Height          =   255
         Left            =   210
         TabIndex        =   30
         Top             =   945
         Width           =   1410
      End
      Begin VB.Label LblTelef 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2010
         TabIndex        =   29
         Top             =   930
         Width           =   1770
      End
      Begin VB.Label Label9 
         Caption         =   "Sexo :"
         Height          =   255
         Left            =   5505
         TabIndex        =   28
         Top             =   300
         Width           =   1410
      End
      Begin VB.Label LblSex 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7305
         TabIndex        =   27
         Top             =   285
         Width           =   1770
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha Nacimiento :"
         Height          =   255
         Left            =   5505
         TabIndex        =   26
         Top             =   615
         Width           =   1410
      End
      Begin VB.Label LblFecNac 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7305
         TabIndex        =   25
         Top             =   600
         Width           =   1770
      End
      Begin VB.Label Label13 
         Caption         =   "Estado Civil :"
         Height          =   255
         Left            =   5505
         TabIndex        =   24
         Top             =   975
         Width           =   1410
      End
      Begin VB.Label LblEst 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7305
         TabIndex        =   23
         Top             =   960
         Width           =   1770
      End
      Begin VB.Label LblfecExp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7305
         TabIndex        =   22
         Top             =   1335
         Width           =   1425
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha de Expiracion :"
         Height          =   255
         Left            =   5505
         TabIndex        =   21
         Top             =   1350
         Width           =   1740
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha de Activación :"
         Height          =   255
         Left            =   210
         TabIndex        =   20
         Top             =   1290
         Width           =   1740
      End
      Begin VB.Label LblFecIng 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2010
         TabIndex        =   19
         Top             =   1275
         Width           =   1425
      End
      Begin VB.Label Lblestado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1995
         TabIndex        =   18
         Top             =   1635
         Width           =   2775
      End
      Begin VB.Label Label15 
         Caption         =   "Estado :"
         Height          =   255
         Left            =   195
         TabIndex        =   17
         Top             =   1650
         Width           =   1740
      End
      Begin VB.Label LblNumTarEntr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7305
         TabIndex        =   16
         Top             =   1665
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Tarjetas Entregadas :"
         Height          =   255
         Left            =   5505
         TabIndex        =   15
         Top             =   1680
         Width           =   1740
      End
   End
   Begin VB.Frame Frame3 
      Height          =   720
      Left            =   15
      TabIndex        =   7
      Top             =   6105
      Width           =   9300
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   390
         Left            =   7875
         TabIndex        =   9
         Top             =   210
         Width           =   1335
      End
      Begin VB.CommandButton CmdAplEst 
         Caption         =   "Aplicar Estado"
         Enabled         =   0   'False
         Height          =   390
         Left            =   135
         TabIndex        =   8
         Top             =   225
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Motivo de Bloqueo: "
      Height          =   705
      Left            =   -15
      TabIndex        =   4
      Top             =   5370
      Width           =   9300
      Begin VB.OptionButton OptEst 
         Caption         =   "Perdida"
         Height          =   300
         Index           =   3
         Left            =   5040
         TabIndex        =   13
         Top             =   300
         Width           =   1125
      End
      Begin VB.OptionButton OptEst 
         Caption         =   "Robada"
         Height          =   300
         Index           =   2
         Left            =   3480
         TabIndex        =   12
         Top             =   300
         Width           =   1125
      End
      Begin OCXTarjeta.CtrlTarjeta CtrlTarjeta1 
         Height          =   390
         Left            =   3630
         TabIndex        =   10
         Top             =   225
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   688
      End
      Begin VB.OptionButton OptEst 
         Caption         =   "Cancelada"
         Height          =   300
         Index           =   1
         Left            =   2160
         TabIndex        =   6
         Top             =   300
         Width           =   1125
      End
      Begin VB.OptionButton OptEst 
         Caption         =   "Retención Cajero"
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   5505
      Begin VB.TextBox TxtDNI 
         Height          =   360
         Left            =   795
         MaxLength       =   8
         TabIndex        =   38
         Top             =   675
         Width           =   1755
      End
      Begin VB.CommandButton CmdBuscarDNI 
         Caption         =   "Buscar Persona"
         Height          =   390
         Left            =   3825
         TabIndex        =   37
         Top             =   705
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
         Top             =   210
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.CommandButton CmdLecTarj 
         Caption         =   "Leer Tarjeta"
         Height          =   345
         Left            =   4095
         TabIndex        =   1
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "DNI :"
         Height          =   255
         Left            =   105
         TabIndex        =   39
         Top             =   750
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Tarjeta :"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   300
         Width           =   735
      End
      Begin VB.Label LblTarjeta 
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
         TabIndex        =   2
         Top             =   240
         Width           =   3225
      End
   End
End
Attribute VB_Name = "frmCambSitu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As New DConecta

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
    
    Me.LblTarjeta.Caption = IIf(IsNull(Cmd.Parameters(1).Value), "", Cmd.Parameters(1).Value)
    
    'Call CerrarConexion
    oConec.CierraConexion
    
    If Len(Trim(Me.LblTarjeta.Caption)) = 0 Then
        MsgBox "No se Encontro Ninguna Tarjeta Asociada al DNI", vbInformation, "Aviso"
        Call LimpiaDatos
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
            LblTarjeta.Caption = TxtNumTarj.Text
            TxtNumTarj.Visible = False
            Me.LblTarjeta.Visible = True
            Me.CmdLecTarj.Visible = True
            Me.Caption = "Cambio de Situacion de Tarjeta - F12 para Digitar Tarjeta"
            If Len(Trim(LblTarjeta.Caption)) > 0 Then
                CmdAplEst.Enabled = True
                Call CargaDatos
            Else
                CmdAplEst.Enabled = False
            End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 123 Then
            TxtNumTarj.Text = ""
            TxtNumTarj.Visible = True
            Me.LblTarjeta.Visible = False
            Me.CmdLecTarj.Visible = False
            Me.Caption = "Cambio de Situacion de Tarjeta - F12 para Digitar Tarjeta"
            TxtNumTarj.SetFocus
    End If
End Sub

Private Sub CmdAplEst_Click()
Dim sResp As String
Dim sTramaResp As String
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
        
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psPAN", adVarChar, adParamInput, 20, LblTarjeta.Caption)
    Cmd.Parameters.Append Prm
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCond", adInteger, adParamInput, , IIf(Me.OptEst(0).Value, 3, IIf(Me.OptEst(1).Value, 50, IIf(Me.OptEst(2).Value, 10, 2))))
    Cmd.Parameters.Append Prm
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pdFecha", adDate, adParamInput, , Now)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cUserBloqCanc", adChar, adParamInput, 4, gsCodUser)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ActualizaCondicionTarjeta"
    Cmd.Execute
    
    'Call CerrarConexion
    oConec.CierraConexion
                               
    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psPAN", adVarChar, adParamInput, 20, LblTarjeta.Caption)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psGlosa", adVarChar, adParamInput, 2000, Me.txtGlosa.Text)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ActualizaGLOSACondicionTarjeta"
    Cmd.Execute
    
    'Call CerrarConexion
    oConec.CierraConexion
    
    MsgBox "Cambio de Situacion de Tarjeta Satisfactorio"
    
    LblTarjeta.Caption = ""
    CmdAplEst.Enabled = False
    Call LimpiaDatos
End Sub

Private Sub CargaDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset

    Set R = New ADODB.Recordset
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 20, Me.LblTarjeta.Caption)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaConsultaTarjeta"
    R.CursorType = adOpenStatic
    R.LockType = adLockReadOnly
    Set R = Cmd.Execute
        
    If Not R.EOF Then
        If CInt(Right(R!cDescrip, 5)) <> 1 Then
            MsgBox "Tarjeta no esta activa", vbInformation, "Aviso"
            Me.CmdAplEst.Enabled = False
            Frame2.Enabled = False
        Else
            Me.CmdAplEst.Enabled = True
            Frame2.Enabled = True
        End If
        
        LblNombre.Caption = R!cPersNombre
        LblDirecc.Caption = R!cPersDireccDomicilio
        LblTelef.Caption = R!cPersTelefono
        LblSex.Caption = IIf(R!cPersNatSexo = "F", "FEMENINO", "MASCULINO")
        LblFecNac.Caption = Format(R!dPersNacCreac, "dd/mm/yyyy")
         Me.LblEst.Caption = IIf(R!nPersNatEstCiv = 1, "SOLTERO", _
            IIf(R!nPersNatEstCiv = 2, "CASADO", _
            IIf(R!nPersNatEstCiv = 3, "VIUDO", "DIVORCIADO")))
        LblfecExp.Caption = Format(R!dFechaExp, "dd/mm/yyyy")
        LblFecIng.Caption = Format(R!dFecActivacion, "dd/mm/yyyy")
        LblEstado.Caption = R!cDescrip
        LblNumTarEntr.Caption = R!nNumTarEntreg
        LblDNI.Caption = R!cDNI
    End If
        
    R.Close
    Set R = Nothing
    
    oConec.CierraConexion
    
End Sub

Private Sub LimpiaDatos()
            LblNombre.Caption = ""
            LblDirecc.Caption = ""
            LblTelef.Caption = ""
            LblSex.Caption = ""
            LblFecNac.Caption = ""
             Me.LblEst.Caption = ""
            LblfecExp.Caption = ""
            LblFecIng.Caption = ""
            LblEstado.Caption = ""
            LblNumTarEntr.Caption = ""
            LblDNI.Caption = ""
            Me.txtGlosa.Text = ""
            
End Sub

Private Sub CmdLecTarj_Click()

    Me.Caption = "Cambio de Situacion de Tarjeta - PASE LA TARJETA"
    
    LblTarjeta.Caption = Mid(CtrlTarjeta1.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto), 2, 16)
    
    Me.Caption = "Cambio de Situacion de Tarjeta - F12 para Digitar Tarjeta"
    CmdAplEst.Enabled = True
    
    Call CargaDatos
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub
