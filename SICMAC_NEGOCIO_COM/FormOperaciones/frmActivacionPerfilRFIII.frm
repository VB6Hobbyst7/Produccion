VERSION 5.00
Begin VB.Form frmActivacionPerfilRFIII 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activacion de Perfiles RFIII"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   FillColor       =   &H80000012&
   Icon            =   "frmActivacionPerfilRFIII.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTipoUsuario 
      Caption         =   "Tipo Usuario"
      Height          =   615
      Left            =   90
      TabIndex        =   11
      Top             =   3570
      Width           =   2865
      Begin VB.ComboBox cboTipoUsuario 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   210
         Width           =   2685
      End
   End
   Begin VB.CommandButton cmdActivar 
      Caption         =   "&Activar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3090
      TabIndex        =   10
      Top             =   3780
      Width           =   1425
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4590
      TabIndex        =   9
      Top             =   3780
      Width           =   1065
   End
   Begin VB.Frame fraOpcionesAdicionales 
      Caption         =   "Opciones Adicionales"
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
      Height          =   855
      Left            =   75
      TabIndex        =   3
      Top             =   2640
      Width           =   5580
      Begin VB.CheckBox optOperacionesSimultaneas 
         Caption         =   "Permitir Operaciones Simultaneas Supervisor - RFIII"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   8
         Top             =   315
         Width           =   4740
      End
   End
   Begin VB.Frame fraGruposActivar 
      Caption         =   "Grupos a Activar"
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
      Height          =   285
      Left            =   5940
      TabIndex        =   2
      Top             =   1980
      Visible         =   0   'False
      Width           =   600
      Begin VB.ListBox ListGruposActivar 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   210
         TabIndex        =   7
         Top             =   315
         Width           =   4740
      End
   End
   Begin VB.Frame fraGruposActuales 
      Caption         =   "Grupos Actuales"
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
      Height          =   1485
      Left            =   75
      TabIndex        =   1
      Top             =   1050
      Width           =   5580
      Begin VB.ListBox ListGruposActuales 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   210
         TabIndex        =   6
         Top             =   315
         Width           =   5190
      End
   End
   Begin VB.Frame fraRf3 
      Caption         =   "RFIII"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   855
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   5580
      Begin VB.TextBox txtNombreUsuario 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   315
         Width           =   4140
      End
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   315
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmActivacionPerfilRFIII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************************************
'* NOMBRE         : "frmActivacionPerfilRFIII"
'* DESCRIPCION    : Formulario creado modificar los grupos a los que pertenede el RFIII. Modificacion segun TI-ERS108-2013
'* CREACION       : RIRO, 20130921 10:00 AM
'**************************************************************************************************************************
Option Explicit

Private sUserRF3 As String
Private sNombreRF3 As String, sCargo As String
Private sPersCod As String
Private sAgenciaPermisoOperSimultaneas As String
Private sGruposActuales() As String
Private sGruposActivar() As String
Private sNombreGrupoRF3 As String
Private bPermitirOperacionesSimultaneas As Boolean
Private bModoSupervisor As Boolean
Private nEstado As Integer

Private Sub cmdActivar_Click()

    Dim sMensaje As String
    sMensaje = ""
    
On Error GoTo error
    
    If bModoSupervisor Then
        If optOperacionesSimultaneas Then
        Else
            sMensaje = "Al desactivar al RFIII con funciones de supervisor éste no tendrá acceso a ninguna de las opciones de supervisor ni podrá emitir vistos, ¿Desea continuar?"
        End If
    Else
        If optOperacionesSimultaneas Then
        Else
            sMensaje = "Al activar al RFIII con funciones de supervisor usted no podrá realizar ninguna operacion ni emitir ningun visto, ¿Desea continuar?"
        End If
    End If
    If Len(Trim(sMensaje)) > 0 Then
        If MsgBox(sMensaje, vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
    End If
    
    If MsgBox("Se procederá a " & cmdActivar.Caption & " el perfil de RFII, ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    Dim oAcceso As UAcceso
    Dim i As Integer
    Set oAcceso = New UAcceso

    'Eliminando "Grupos Actuales"
'    For i = 1 To UBound(sGruposActuales)
'        oAcceso.EliminaGrupodeUsuario gsDominio, sUserRF3, sGruposActuales(i) '///// COMENTADO SOLO POR PRUEBAS
'    Next
    'Asignando "Grupos a Activar"
'    For i = 1 To UBound(sGruposActivar)
'        oAcceso.AgregaGrupoAUsuario gsDominio, sUserRF3, sGruposActivar(i) '///// COMENTADO SOLO POR PRUEBAS
'    Next
    Set oAcceso = Nothing
    Dim oDCOMPersona As New COMDPersona.DCOMPersonas
'    If Not bModoSupervisor Then
'        oDCOMPersona.RegistrarGruposRF3 sPersCod, optOperacionesSimultaneas.value
'    Else
'        oDCOMPersona.RegistrarGruposRF3 sPersCod, optOperacionesSimultaneas.value
'    End If
    
    oDCOMPersona.RegistrarGruposRF3 sPersCod, optOperacionesSimultaneas.value, , Trim(Right(cboTipoUsuario.Text, 10))
    Set oDCOMPersona = Nothing
    bModoSupervisor = False
    ReDim sGruposActuales(0)
    'ReDim sGruposActivar(0)
    ListGruposActuales.Clear
    'ListGruposActivar.Clear
    ObtenerGruposActuales
    'ObtenerGruposActivar

    CargarDatos
    MsgBox "Proceso completado", vbInformation, "Aviso"
    'Unload Me
Exit Sub
error:
    MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
CargarDatos
End Sub

Public Sub CargarDatos()

    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(10027, , 100, "1")
    sNombreGrupoRF3 = rsConst!cDescripcion
    Set rsConst = clsGen.GetConstante(10027, , 200, "1")
    sAgenciaPermisoOperSimultaneas = rsConst!cDescripcion
    bModoSupervisor = False
    
    cboTipoUsuario.Clear
    cboTipoUsuario.AddItem "ASESOR DE CLIENTES" & Space(50) & "006024"
    cboTipoUsuario.AddItem "REPRESENTANTE FINANCIERO" & Space(50) & "007012"
    cboTipoUsuario.AddItem "TASADOR" & Space(50) & "7013"
    cboTipoUsuario.ListIndex = 0
    
    If ObtenerRF3 Then
        txtUsuario.Text = sUserRF3
        txtNombreUsuario.Text = sNombreRF3
        optOperacionesSimultaneas.value = IIf(bPermitirOperacionesSimultaneas, 1, 0)
        
        If nEstado = 0 Or nEstado = 2 Then
            cmdActivar.Caption = "Activar"
            cboTipoUsuario.Enabled = False
        ElseIf nEstado = 1 Then
            cmdActivar.Caption = "Desactivar"
            cboTipoUsuario.Enabled = True
        End If
        
        ObtenerGruposActuales
        'ObtenerGruposActivar
        
        'Seleccionando el cargo al que pertenece el rf3
        Select Case sCargo
            Case Trim(Right(cboTipoUsuario.List(0), 10))
                cboTipoUsuario.ListIndex = 0
            Case Trim(Right(cboTipoUsuario.List(1), 10))
                cboTipoUsuario.ListIndex = 1
            Case Trim(Right(cboTipoUsuario.List(2), 10))
                cboTipoUsuario.ListIndex = 2
            Case Else
                cboTipoUsuario.ListIndex = -1
        End Select
        
    Else
        txtUsuario.Text = ""
        txtNombreUsuario.Text = ""
        fraGruposActivar.Enabled = False
        fraGruposActuales.Enabled = False
        fraRf3.Enabled = False
        fraOpcionesAdicionales.Enabled = False
        cmdActivar.Enabled = False
        Exit Sub
    End If
    
    If InStr(1, sAgenciaPermisoOperSimultaneas, gsCodAge) <> 0 Then
        fraOpcionesAdicionales.Enabled = True
        optOperacionesSimultaneas.Enabled = True
    Else
        fraOpcionesAdicionales.Enabled = False
        optOperacionesSimultaneas.Enabled = False
    End If

End Sub


Public Function ObtenerGruposActuales() As Boolean

    Dim oAcceso As UAcceso
    Dim sCad As String
    Dim i As Integer
    Dim bPertenece As Boolean
    
    Set oAcceso = New UAcceso
    
    If Len(Trim(txtUsuario.Text)) <> 0 Then
        
        ListGruposActuales.Clear
        Call oAcceso.CargaGruposUsuario(txtUsuario.Text, gsDominio)
        sCad = oAcceso.DameGrupoUsuario
        ReDim Preserve sGruposActuales(i)
        
        Dim oPersona As New COMDPersona.DCOMPersonas
        Dim rsRF3 As New ADODB.Recordset
        Set rsRF3 = oPersona.RecuperarGruposRF3(Trim(txtUsuario.Text))
        
        Do While Len(sCad) > 0
            If Trim(sCad) = Trim(sNombreGrupoRF3) Then
                If Not rsRF3 Is Nothing Then
                    If Not rsRF3.BOF And Not rsRF3.EOF Then
                        If rsRF3!nEstado = 1 Then
                            bModoSupervisor = True
                        Else
                            bModoSupervisor = False
                        End If
                    Else
                        bModoSupervisor = False
                    End If
                Else
                    bModoSupervisor = False
                End If
            End If
            i = i + 1
            ReDim Preserve sGruposActuales(i)
            ListGruposActuales.AddItem sCad
            sGruposActuales(i) = sCad
            sCad = oAcceso.DameGrupoUsuario
        Loop
        
        Set oPersona = Nothing
        Set rsRF3 = Nothing
        
    End If
    
End Function

Public Function ObtenerGruposActivar() As Boolean
    If Len(Trim(txtUsuario.Text)) <> 0 Then
        If bModoSupervisor Then
            Dim oDCOMPersona As New COMDPersona.DCOMPersonas
            Dim rsRF3 As ADODB.Recordset
            Dim i As Integer
            
            ListGruposActivar.Clear
            Set rsRF3 = oDCOMPersona.RecuperarGruposRF3(sPersCod)
            ReDim Preserve sGruposActivar(i)
            If Not rsRF3 Is Nothing Then
                i = 1
                Do While Not (rsRF3.BOF Or rsRF3.EOF)
                    ReDim Preserve sGruposActivar(i)
                    ListGruposActivar.AddItem rsRF3!cGrupo
                    sGruposActivar(i) = rsRF3!cGrupo
                    i = i + 1
                    rsRF3.MoveNext
                Loop
            End If
        Else
            ListGruposActivar.Clear
            ListGruposActivar.AddItem sNombreGrupoRF3
            ReDim Preserve sGruposActivar(1)
            sGruposActivar(1) = sNombreGrupoRF3
        End If
        ObtenerGruposActivar = True
    Else
        ObtenerGruposActivar = False
    End If
End Function

'Private Sub RecargaGrupos()
'Dim i As Integer
'    LstGrupo.Clear
'    For i = 0 To UBound(MatGrupos) - 1
'        LstGrupo.AddItem MatGrupos(i)
'    Next i
'End Sub

Public Function ObtenerRF3() As Boolean
    Dim oDCOMPersona As New COMDPersona.DCOMPersonas
    Dim rsRF3 As ADODB.Recordset
    Set rsRF3 = oDCOMPersona.RecuperarDatosRF3(gsCodAge)
    
    If Not rsRF3 Is Nothing Then
        If Not rsRF3.EOF And Not rsRF3.BOF Then

            Dim oAcceso As New UAcceso
            Dim sCad, sGrupos As String
            Dim i As Integer
            
            
            sGrupos = ""
            If Trim(rsRF3!cuser) <> "" Then
                Call oAcceso.CargaGruposUsuario(Trim(rsRF3!cuser), gsDominio)
                sCad = oAcceso.DameGrupoUsuario
                Do While Len(sCad) > 0
                    i = i + 1
                    sGrupos = sGrupos & "," & sCad
                    sCad = oAcceso.DameGrupoUsuario
                Loop
            End If
            
            If InStr(1, sGrupos, sNombreGrupoRF3) > 0 Then
                sUserRF3 = rsRF3!cuser
                sNombreRF3 = rsRF3!cPersNombre
                sPersCod = rsRF3!cPersCod
                If rsRF3!nAccionesSimultaneas = 0 Or rsRF3!nAccionesSimultaneas = 2 Then
                    bPermitirOperacionesSimultaneas = False
                ElseIf rsRF3!nAccionesSimultaneas = 1 Then
                    bPermitirOperacionesSimultaneas = True
                End If
                nEstado = rsRF3!nEstado
                ObtenerRF3 = True
                sCargo = oDCOMPersona.RecuperarCargoRF3(sPersCod)
            Else
                ObtenerRF3 = False
            End If
            
        Else
            ObtenerRF3 = False
        End If
    Else
        ObtenerRF3 = False
    End If
End Function

