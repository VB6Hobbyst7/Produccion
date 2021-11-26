VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCredNewNivAprDelegacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delegación de Aprobación de Créditos"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   Icon            =   "frmCredNewNivAprDelegacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRetirar 
      Caption         =   "Retirar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3480
      TabIndex        =   11
      Top             =   3120
      Width           =   1170
   End
   Begin VB.CommandButton cmdDelegar 
      Caption         =   "Delegar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Usuario"
      TabPicture(0)   =   "frmCredNewNivAprDelegacion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblUsuario"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTipoNiv"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraDelegacion"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame fraDelegacion 
         Caption         =   " Delegación "
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   4335
         Begin SICMACT.TxtBuscar txtBuscarUsuario 
            Height          =   315
            Left            =   1080
            TabIndex        =   7
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox txtDesde 
            Height          =   300
            Left            =   1080
            TabIndex        =   13
            Top             =   600
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   19
            Mask            =   "##/##/#### ##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtHasta 
            Height          =   300
            Left            =   1080
            TabIndex        =   14
            Top             =   960
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   19
            Mask            =   "##/##/#### ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label6 
            Caption         =   "Hasta :"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Desde :"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   640
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Usuario :"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   280
            Width           =   735
         End
      End
      Begin VB.Label lblTipoNiv 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2520
         TabIndex        =   4
         Top             =   930
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Nivel de Aprobación :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario :"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   510
         Width           =   735
      End
      Begin VB.Label lblUsuario 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmCredNewNivAprDelegacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredNewNivAprDelegacion
'** Descripción : Formulario para la delegar una aprobación de credito a un usuario creado segun RFC110-2012
'** Creación : JUEZ, 20121204 09:00:00 AM
'**********************************************************************************************
Option Explicit
Dim oDNiv As COMDCredito.DCOMNivelAprobacion
Dim oNNiv As COMNCredito.NCOMNivelAprobacion
Dim rs As ADODB.Recordset

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdDelegar_Click()
    If ValidaDatosDelegacion Then
        If MsgBox("¿Está seguro de delegar la Aprobación al usuario " & txtBuscarUsuario & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
            Set oNNiv = New COMNCredito.NCOMNivelAprobacion
            Call oNNiv.dDelegarAprobacion(gsCodUser, Trim(txtBuscarUsuario), txtDesde.Text, txtHasta.Text)
            Set oNNiv = Nothing
        MsgBox "La delegación se hizo correctamente", vbInformation, "Aviso"
        Unload Me
    End If
End Sub

Private Sub cmdRetirar_Click()
    Set oNNiv = New COMNCredito.NCOMNivelAprobacion
    If MsgBox("¿Está seguro de retirar la delegación de Aprobación al usuario " & txtBuscarUsuario & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Call oNNiv.dRetirarDelegacionAprobacion(gsCodUser)
    MsgBox "Se retiró la delegación correctamente", vbInformation, "Aviso"
    Unload Me
    Set oNNiv = Nothing
End Sub

Public Sub Inicio()
    Dim oConst As COMDConstantes.DCOMConstantes
    Dim sNombreUsu As String, sTipoNiv As String
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    
    Set oConst = New COMDConstantes.DCOMConstantes
    txtBuscarUsuario.lbUltimaInstancia = False
    txtBuscarUsuario.psRaiz = "USUARIOS DISPONIBLES PARA LOS NIVELES DE APROBACION"
    txtBuscarUsuario.rs = oConst.ObtenerUsuariosArea
    
    If oDNiv.VerificaUsuarioSiTieneNivel(gsCodUser, sNombreUsu, sTipoNiv) Then
        lblUsuario.Caption = sNombreUsu
        lblTipoNiv.Caption = sTipoNiv
        If oDNiv.VerificaUsuarioSiTieneDelegacion(gsCodUser) Then
            RetirarDelegacion
        Else
            NuevaDelegacion
        End If
        Me.Show 1
    Else
        MsgBox "Ud. no tiene permiso para visualizar esta opción", vbInformation, "Aviso"
        Unload Me
        Exit Sub
    End If
    Set oDNiv = Nothing
End Sub

Private Sub NuevaDelegacion()
    cmdDelegar.Visible = True
    cmdRetirar.Visible = False
    
    fraDelegacion.Enabled = True
End Sub

Private Sub RetirarDelegacion()
    cmdDelegar.Visible = False
    cmdRetirar.Visible = True
    
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oDNiv.RecuperaDelegacion(gsCodUser, 1)
    Set oDNiv = Nothing
    txtBuscarUsuario = rs!cUserDelegado
    txtDesde.Text = Format(rs!dFechaDesde, "dd/mm/yyyy hh:mm:ss")
    txtHasta.Text = Format(rs!dFechaHasta, "dd/mm/yyyy hh:mm:ss")
    fraDelegacion.Enabled = False
End Sub

Private Sub txtBuscarUsuario_EmiteDatos()
    txtBuscarUsuario = Right(txtBuscarUsuario, 4)
    txtDesde.SetFocus
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtHasta.SetFocus
    End If
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdDelegar.SetFocus
    End If
End Sub

Private Function ValidaDatosDelegacion() As Boolean
    Dim sUserDelegado As String
    Dim oConst As COMDConstantes.DCOMConstantes
    ValidaDatosDelegacion = False
    Dim ldFechaHoraSist As Date
    
    Set oConst = New COMDConstantes.DCOMConstantes
    Set rs = oConst.ObtenerUsuariosArea(txtBuscarUsuario)
    Set oConst = Nothing
    If rs.EOF Then
        MsgBox "El usuario ingresado no es valido", vbInformation, "Aviso"
        ValidaDatosDelegacion = False
        Exit Function
    End If
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oDNiv.RecuperaDelegacion(txtBuscarUsuario, 2)
    Set oDNiv = Nothing
    If Not rs.EOF Then
        MsgBox "El Usuario " & txtBuscarUsuario & " ya tiene una delegacion activa", vbInformation, "Aviso"
        ValidaDatosDelegacion = False
        Exit Function
    End If
    If txtBuscarUsuario = "" Then
        MsgBox "Debe ingresar el usuario a delegar", vbInformation, "Aviso"
        txtBuscarUsuario.SetFocus
        ValidaDatosDelegacion = False
        Exit Function
    End If
    If txtDesde.Text = "__/__/____ __:__:__" Then
        MsgBox "Debe ingresar la Fecha de Inicio", vbInformation, "Aviso"
        txtDesde.SetFocus
        ValidaDatosDelegacion = False
        Exit Function
    End If
    If txtHasta.Text = "__/__/____ __:__:__" Then
        MsgBox "Debe ingresar la Fecha Final", vbInformation, "Aviso"
        txtHasta.SetFocus
        ValidaDatosDelegacion = False
        Exit Function
    End If
    If ValidaFecha(Left(txtDesde.Text, 10)) <> "" Then
        MsgBox ValidaFecha(Left(txtDesde.Text, 10)) & " en la Fecha de Inicio", vbInformation, "Aviso"
        txtDesde.SetFocus
        ValidaDatosDelegacion = False
        Exit Function
    End If
    If ValidaFecha(Left(txtHasta.Text, 10)) <> "" Then
        MsgBox ValidaFecha(Left(txtHasta.Text, 10)) & " en la Fecha Final", vbInformation, "Aviso"
        txtHasta.SetFocus
        ValidaDatosDelegacion = False
        Exit Function
    End If
    If ValidarHoras(Right(txtDesde.Text, 8)) <> "" Then
        MsgBox ValidarHoras(Right(txtDesde.Text, 8)) & " en la Fecha de Inicio", vbInformation, "Aviso"
        txtDesde.SetFocus
        ValidaDatosDelegacion = False
        Exit Function
    End If
    If ValidarHoras(Right(txtHasta.Text, 8)) <> "" Then
        MsgBox ValidarHoras(Right(txtHasta.Text, 8)) & " en la Fecha Final", vbInformation, "Aviso"
        txtHasta.SetFocus
        ValidaDatosDelegacion = False
        Exit Function
    End If
    ldFechaHoraSist = CDate(gdFecSis & " " & GetHoraServer)
    'If txtDesde.Text < gdFecSis & " " & GetHoraServer Then
    If CDate(txtDesde.Text) < ldFechaHoraSist Then
        'MsgBox "La Fecha de Inicio no puede ser menor a la Fecha Actual", vbInformation, "Aviso"
        MsgBox "La Fecha de Inicio no puede ser menor a la Fecha Actual " & Format(ldFechaHoraSist, "dd/mm/yyyy hh:mm:ss"), vbInformation, "Aviso"
        txtDesde.SetFocus
        ValidaDatosDelegacion = False
        Exit Function
    End If
    'If txtHasta.Text < txtDesde.Text Then
    If CDate(txtHasta.Text) < CDate(txtDesde.Text) Then
        MsgBox "La Fecha Final no puede ser menor a la Fecha de Inicio", vbInformation, "Aviso"
        txtHasta.SetFocus
        ValidaDatosDelegacion = False
        Exit Function
    End If
    ValidaDatosDelegacion = True
End Function

Private Function ValidarHoras(ByVal psHora As String) As String
    If Mid(psHora, 1, 2) >= 0 And Mid(psHora, 1, 2) <= 23 Then
        If Mid(psHora, 4, 2) >= 0 And Mid(psHora, 4, 2) <= 59 Then
            If Mid(psHora, 7, 2) >= 0 And Mid(psHora, 7, 2) <= 59 Then
                ValidarHoras = ""
            Else
                ValidarHoras = "Los segundos no son válidos"
                Exit Function
            End If
        Else
            ValidarHoras = "Los Minutos no son válidos"
            Exit Function
        End If
    Else
        ValidarHoras = "La Hora no es válida"
        Exit Function
    End If
End Function
