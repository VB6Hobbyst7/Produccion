VERSION 5.00
Object = "{F9AB04EF-FCD4-4161-99E1-9F65F8191D72}#14.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmConsTarj 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Tarjeta  - F12 para Digitar Tarjeta"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frmConsTarj.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OCXTarjeta.CtrlTarjeta CtrlTarjeta1 
      Height          =   195
      Left            =   2205
      TabIndex        =   26
      Top             =   6210
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   344
   End
   Begin VB.Frame Frame3 
      Height          =   690
      Left            =   30
      TabIndex        =   5
      Top             =   5145
      Width           =   5475
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   360
         Left            =   4245
         TabIndex        =   6
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de Tarjeta"
      Height          =   3750
      Left            =   30
      TabIndex        =   4
      Top             =   1350
      Width           =   5505
      Begin VB.Label Label4 
         Caption         =   "Tarjetas Entregadas :"
         Height          =   255
         Left            =   195
         TabIndex        =   28
         Top             =   3330
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
         Left            =   1995
         TabIndex        =   27
         Top             =   3315
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Estado :"
         Height          =   255
         Left            =   195
         TabIndex        =   25
         Top             =   3015
         Width           =   1740
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
         TabIndex        =   24
         Top             =   3000
         Width           =   2775
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
         TabIndex        =   23
         Top             =   2640
         Width           =   1425
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha de Activación :"
         Height          =   255
         Left            =   210
         TabIndex        =   22
         Top             =   2655
         Width           =   1740
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha de Expiracion :"
         Height          =   255
         Left            =   210
         TabIndex        =   21
         Top             =   2295
         Width           =   1740
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
         Left            =   2010
         TabIndex        =   20
         Top             =   2280
         Width           =   1425
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
         Left            =   2010
         TabIndex        =   19
         Top             =   1905
         Width           =   1770
      End
      Begin VB.Label Label13 
         Caption         =   "Estado Civil :"
         Height          =   255
         Left            =   210
         TabIndex        =   18
         Top             =   1920
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
         Left            =   2010
         TabIndex        =   17
         Top             =   1560
         Width           =   1770
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha Nacimiento :"
         Height          =   255
         Left            =   210
         TabIndex        =   16
         Top             =   1575
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
         Left            =   2010
         TabIndex        =   15
         Top             =   1245
         Width           =   1770
      End
      Begin VB.Label Label9 
         Caption         =   "Sexo :"
         Height          =   255
         Left            =   210
         TabIndex        =   14
         Top             =   1260
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
         TabIndex        =   13
         Top             =   930
         Width           =   1770
      End
      Begin VB.Label Label7 
         Caption         =   "Telefono :"
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   945
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
         TabIndex        =   11
         Top             =   615
         Width           =   3405
      End
      Begin VB.Label Label5 
         Caption         =   "Direccion :"
         Height          =   255
         Left            =   210
         TabIndex        =   10
         Top             =   630
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
         TabIndex        =   9
         Top             =   285
         Width           =   3390
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   300
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5505
      Begin VB.CommandButton CmdBuscarDNI 
         Caption         =   "Buscar Persona"
         Height          =   390
         Left            =   3795
         TabIndex        =   30
         Top             =   750
         Width           =   1545
      End
      Begin VB.TextBox TxtDNI 
         Height          =   360
         Left            =   780
         MaxLength       =   8
         TabIndex        =   29
         Top             =   720
         Width           =   1755
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
         Left            =   780
         MaxLength       =   16
         TabIndex        =   7
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
      Begin VB.Label Label1 
         Caption         =   "DNI :"
         Height          =   255
         Left            =   75
         TabIndex        =   31
         Top             =   795
         Width           =   480
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
Attribute VB_Name = "frmConsTarj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub LimpiaDatos()
    LblNombre.Caption = ""
    LblDirecc.Caption = ""
    LblTelef.Caption = ""
    LblSex.Caption = ""
    LblFecNac.Caption = ""
    LblEst.Caption = ""
    LblfecExp.Caption = ""
    LblFecIng.Caption = ""
    Lblestado.Caption = ""
    LblNumTarEntr.Caption = ""

End Sub
Private Sub CmdBuscarDNI_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter

Dim loConec As DConecta
Set loConec = New DConecta


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
            Me.Caption = "Consulta de Tarjeta  - F12 para Digitar Tarjeta"
            If Not ExisteTarjeta(LblTarjeta.Caption) Then
                LblTarjeta.Caption = ""
                Exit Sub
            End If
            If Len(Trim(LblTarjeta.Caption)) > 0 Then
                Call CargaDatos
            Else

            End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 123 Then
            TxtNumTarj.Text = ""
            TxtNumTarj.Visible = True
            Me.LblTarjeta.Visible = False
            Me.CmdLecTarj.Visible = False
            Me.Caption = "Consulta de Tarjeta  - F12 para Digitar Tarjeta"
            TxtNumTarj.SetFocus
    End If
End Sub
Private Function VerificaSITarjetaActiva() As Boolean


If TarjetaActiva(Me.LblTarjeta.Caption) Then
    VerificaSITarjetaActiva = True
Else
    VerificaSITarjetaActiva = False
End If

End Function

Private Sub CargaDatos()

Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset

    If Not VerificaSITarjetaActiva Then
        MsgBox "Tarjeta No esta Activa, Consulte el Estado de la Tarjeta"
        Exit Sub
    End If
    
    Set R = New ADODB.Recordset
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 20, Me.LblTarjeta.Caption)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaConsultaTarjeta"
    R.CursorType = adOpenStatic
    R.LockType = adLockReadOnly
    Set R = Cmd.Execute

    If Not R.EOF Then
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
        Lblestado.Caption = R!cDescrip
        LblNumTarEntr.Caption = R!nNumTarEntreg
    End If
    
    
    R.Close
    Set R = Nothing
    
    oConec.CierraConexion
        
End Sub

Private Sub CmdLecTarj_Click()

    Me.Caption = "Consulta de Tarjeta - PASE LA TARJETA"
    
    LblTarjeta.Caption = Mid(CtrlTarjeta1.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto), 2, 16)
    
    Me.Caption = "Consulta de Tarjeta  - F12 para Digitar Tarjeta"
    
    If Not ExisteTarjeta(LblTarjeta.Caption) Then
        LblTarjeta.Caption = ""
        Exit Sub
    End If
    Call CargaDatos
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub
