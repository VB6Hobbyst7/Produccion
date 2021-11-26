VERSION 5.00
Begin VB.Form frmAdmUsu 
   Caption         =   "Administracion de Usuario"
   ClientHeight    =   5250
   ClientLeft      =   1905
   ClientTop       =   2400
   ClientWidth     =   14115
   Icon            =   "frmAdmUsu.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   14115
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12700
      TabIndex        =   17
      Top             =   4680
      Width           =   1300
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   75
      TabIndex        =   3
      Top             =   1545
      Width           =   13900
      Begin VB.Frame Frame4 
         Height          =   2835
         Left            =   2970
         TabIndex        =   10
         Top             =   135
         Width           =   10800
         Begin VB.CommandButton CmdIzq 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5100
            TabIndex        =   16
            Top             =   1320
            Width           =   510
         End
         Begin VB.CommandButton CmdDer 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5100
            TabIndex        =   15
            Top             =   720
            Width           =   510
         End
         Begin VB.ListBox LstGrupo 
            Height          =   2205
            Left            =   5700
            TabIndex        =   12
            Top             =   435
            Width           =   5000
         End
         Begin VB.ListBox LstPertenece 
            Height          =   2205
            Left            =   90
            TabIndex        =   11
            Top             =   480
            Width           =   4900
         End
         Begin VB.Label Label4 
            Caption         =   "Grupos :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   5700
            TabIndex        =   14
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label Label3 
            Caption         =   "Pertenece a :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   120
            TabIndex        =   13
            Top             =   180
            Width           =   1995
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Acceso a Maquinas"
         Height          =   2880
         Left            =   75
         TabIndex        =   5
         Top             =   120
         Width           =   2865
         Begin VB.CheckBox ChkTodasMaq 
            Caption         =   "A Todas las Maquinas"
            Height          =   225
            Left            =   105
            TabIndex        =   18
            Top             =   300
            Width           =   2550
         End
         Begin VB.CommandButton CmdDelMaq 
            Caption         =   "&Eliminar"
            Height          =   360
            Left            =   1500
            TabIndex        =   9
            Top             =   2445
            Width           =   1035
         End
         Begin VB.CommandButton CmdAddMaq 
            Caption         =   "&Agregar"
            Height          =   360
            Left            =   225
            TabIndex        =   8
            Top             =   2445
            Width           =   1035
         End
         Begin VB.ListBox LstMaquina 
            Height          =   1425
            Left            =   90
            TabIndex        =   7
            Top             =   945
            Width           =   2595
         End
         Begin VB.TextBox TxtMaquina 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   90
            TabIndex        =   6
            Top             =   585
            Width           =   2610
         End
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   13900
      Begin VB.CheckBox ChkCtaBloq 
         Caption         =   "Cuenta Bloqueada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1065
         Width           =   1950
      End
      Begin VB.ComboBox CmbUser 
         Height          =   315
         Left            =   1095
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   2415
      End
      Begin VB.Label lblNomusu 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   210
         TabIndex        =   19
         Top             =   675
         Width           =   6450
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6780
         Picture         =   "frmAdmUsu.frx":030A
         Top             =   525
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   165
         Left            =   195
         TabIndex        =   2
         Top             =   270
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmAdmUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MatGrupos() As String
Dim bBloq As Boolean
Dim objPista As COMManejador.Pista 'MAVM 20110407

Private Sub CargaMaquinasUsuario()
Dim oAcceso As UAcceso
Dim vMaquinas As Variant
Dim i As Integer

    Set oAcceso = New UAcceso
    TxtMaquina.Text = ""
    vMaquinas = oAcceso.DameMaquinasdeUsuario(CmbUser.Text, gsDominio)
    If Not IsArray(vMaquinas) Then  'Tiene acceso a todas las maquinas
        If Len(vMaquinas) > 0 Then
            ChkTodasMaq.value = 0
            LstMaquina.Clear
            LstMaquina.AddItem vMaquinas
        Else
            ChkTodasMaq.value = 1
            TxtMaquina.Enabled = False
            LstMaquina.Clear
            LstMaquina.Enabled = False
        End If
    Else
        ChkTodasMaq.value = 0
        LstMaquina.Clear
        For i = 0 To UBound(vMaquinas)
            LstMaquina.AddItem vMaquinas(i)
        Next i
    End If
    Set oAcceso = Nothing
End Sub

Private Sub RecargaGrupos()
Dim i As Integer
    LstGrupo.Clear
    For i = 0 To UBound(MatGrupos) - 1
        LstGrupo.AddItem MatGrupos(i)
    Next i
End Sub

Private Sub ChkCtaBloq_Click()
Dim i As Integer
Dim oAcceso As UAcceso
    
    Screen.MousePointer = 11
    Set oAcceso = New UAcceso
    If ChkCtaBloq.value = 1 Then
        If bBloq Then
            Exit Sub
        End If
        For i = 0 To 7
            Call oAcceso.ChangePassword(gsDominio, CmbUser.Text, "X", "X")
        Next i
    Else
        Call oAcceso.Desbloquear_Habilitar_Cuenta(gsDominio, CmbUser.Text)
    End If
    Set oAcceso = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub ChkTodasMaq_Click()
    If ChkTodasMaq.value = 1 Then
        Call AsignarAccesoATodasMaquinas(gsDominio, CmbUser.Text)
    Else
        LstMaquina.Clear
        LstMaquina.Enabled = True
        TxtMaquina.Enabled = True
    End If
End Sub

Private Sub CmbUser_Click()
Dim oAcceso As UAcceso
Dim sCad As String
Dim i As Integer
Dim bPertenece As Boolean
    
    Call RecargaGrupos
    
    Set oAcceso = New UAcceso
    If Len(Trim(CmbUser.Text)) <> 0 Then
    
        lblNomusu.Caption = oAcceso.MostarNombre(gsDominio, CmbUser.Text)
        
        LstPertenece.Clear
        Call oAcceso.CargaGruposUsuario(CmbUser.Text, gsDominio)
        sCad = oAcceso.DameGrupoUsuario
        Do While Len(sCad) > 0
            LstPertenece.AddItem sCad
            For i = 0 To LstGrupo.ListCount - 1
                If LstGrupo.List(i) = sCad Then
                    LstGrupo.RemoveItem (i)
                    Exit For
                End If
            Next i
            sCad = oAcceso.DameGrupoUsuario
        Loop
        
    End If
    
    'Chequea si cuenta Bloqueada
    If oAcceso.CuentaBloqueada(CmbUser.Text, gsDominio) Then
        bBloq = True
        ChkCtaBloq.value = 1
        bBloq = False
    Else
        ChkCtaBloq.value = 0
    End If
    Set oAcceso = Nothing
    
    Call CargaMaquinasUsuario
End Sub

Private Sub CmdAddMaq_Click()
Dim m() As String
Dim i As Integer
Dim oAcceso As UAcceso
        
    If Len(Trim(Me.TxtMaquina.Text)) > 0 Then
        LstMaquina.AddItem TxtMaquina.Text
        
        Set oAcceso = New UAcceso
        ReDim m(0)
        For i = 0 To LstMaquina.ListCount - 1
            ReDim Preserve m(i + 1)
            m(i) = LstMaquina.List(i)
        Next i
        Call oAcceso.AsignarAccesoAMaquinas(gsDominio, CmbUser.Text, m)
        Set oAcceso = Nothing
    End If
End Sub

Private Sub CmdDelMaq_Click()
Dim m() As String
Dim i As Integer
Dim oAcceso As UAcceso
        
    If Len(Trim(LstMaquina.Text)) > 0 Then
        LstMaquina.RemoveItem LstMaquina.ListIndex
        
        Set oAcceso = New UAcceso
        ReDim m(0)
        For i = 0 To LstMaquina.ListCount - 1
            m(i) = LstMaquina.List(i)
        Next i
        Call oAcceso.AsignarAccesoAMaquinas(gsDominio, CmbUser.Text, m)
        Set oAcceso = Nothing
    End If
End Sub

Private Sub CmdDer_Click()
Dim oAcceso As UAcceso
Dim oActualizaDatosrrhh As DActualizaDatosRRHH 'RECO20140220 ERS164-2013**********************
                           
    If Len(CmbUser.Text) > 0 And LstPertenece.ListCount > 0 And LstPertenece.ListIndex <> -1 Then
        Set oAcceso = New UAcceso
        oAcceso.EliminaGrupodeUsuario gsDominio, CmbUser.Text, LstPertenece.Text
        'RECO2014 ERS164-2013**********************************************************
        Set oActualizaDatosrrhh = New DActualizaDatosRRHH
        oActualizaDatosrrhh.RegistraAsigRetiroGrupo gdFecSis, LstPertenece.Text, CmbUser.Text, gsCodUser, 2
        Set oActualizaDatosrrhh = Nothing
        'RECO FIN**********************************************************************
        Set oAcceso = Nothing
        
        'MAVM 20110407 ***
        Dim clsDGnral As DLogGeneral
        Set clsDGnral = New DLogGeneral
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Se eliminó el usuario: " & CmbUser.Text & " del grupo: " & LstPertenece.Text
        '***
        
        LstGrupo.AddItem LstPertenece.Text
        LstPertenece.RemoveItem LstPertenece.ListIndex
    End If
End Sub

Private Sub CmdIzq_Click()
Dim oAcceso As UAcceso
Dim oActualizaDatosrrhh As DActualizaDatosRRHH 'RECO20140220 ERS164-2013**********************

    If Len(CmbUser.Text) > 0 And LstGrupo.ListCount > 0 And LstGrupo.ListIndex <> -1 Then
        Set oAcceso = New UAcceso
        oAcceso.AgregaGrupoAUsuario gsDominio, CmbUser.Text, LstGrupo.Text
        Set oAcceso = Nothing
        'RECO2014 ERS164-2013**********************************************************
        Set oActualizaDatosrrhh = New DActualizaDatosRRHH
        oActualizaDatosrrhh.RegistraAsigRetiroGrupo gdFecSis, LstGrupo.Text, CmbUser.Text, gsCodUser, 1
        Set oActualizaDatosrrhh = Nothing
        'RECO FIN**********************************************************************
        'MAVM 20110407 ***
        Dim clsDGnral As DLogGeneral
        Set clsDGnral = New DLogGeneral
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Se agrego el usuario: " & CmbUser.Text & " al Grupo: " & LstGrupo.Text
        '***
        
        LstPertenece.AddItem LstGrupo.Text
        LstGrupo.RemoveItem LstGrupo.ListIndex
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CargaGrupos()
Dim oAcceso As UAcceso
Dim sCad As String
Dim i As Integer

    Set oAcceso = New UAcceso
    Call oAcceso.CargaControlGrupos(gsDominio)
    LstGrupo.Clear
    sCad = oAcceso.DameGrupo
    i = 0
    ReDim MatGrupos(i)
    Do While sCad <> ""
        LstGrupo.AddItem sCad
        ReDim Preserve MatGrupos(i + 1)
        MatGrupos(i) = sCad
        sCad = oAcceso.DameGrupo
        i = i + 1
    Loop
    Set oAcceso = Nothing

End Sub

Private Sub CargaUsuarios()
Dim oAcceso As UAcceso
Dim sCad As String
    
    Set oAcceso = New UAcceso
    Call oAcceso.CargaControlUsuarios(gsDominio)
    CmbUser.Clear
    sCad = oAcceso.DameUsuario
    Do While sCad <> ""
        CmbUser.AddItem sCad
        sCad = oAcceso.DameUsuario
    Loop
    Set oAcceso = Nothing
End Sub


Private Sub Form_Load()
    CentraForm Me
    Call CargaUsuarios
    Call CargaGrupos
    gsOpeCod = LogPistaAdministracionUsuario
End Sub
