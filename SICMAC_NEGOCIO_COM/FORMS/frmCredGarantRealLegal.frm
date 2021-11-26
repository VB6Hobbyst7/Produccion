VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCredGarantRealLegal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bloqueo/Desbloqueo Legal"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8985
   Icon            =   "frmCredGarantRealLegal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton CmdBuscaPersona 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   3480
         TabIndex        =   5
         ToolTipText     =   "Busca Documentos de Persona"
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox txtNumGar 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   4
         Tag             =   "txtPrincipal"
         Text            =   "0"
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmdsalir 
         Cancel          =   -1  'True
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
         Height          =   320
         Left            =   7560
         TabIndex        =   3
         ToolTipText     =   "Busca Documentos de Persona"
         Top             =   2520
         Width           =   1080
      End
      Begin VB.CommandButton CmdActualizaLegal 
         Caption         =   "&Bloquear"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "No se podra Modificar"
         Top             =   2520
         Width           =   1080
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Desbloquear"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1320
         TabIndex        =   1
         ToolTipText     =   "Se podra Modificar"
         Top             =   2520
         Width           =   1320
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAnalista 
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2355
         _Version        =   393216
         Cols            =   6
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.CheckBox chkActXHistorico 
         Caption         =   "Permitir actualización garantía aunque quede descoberturada"
         Enabled         =   0   'False
         Height          =   435
         Left            =   4440
         TabIndex        =   10
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label lbltitular 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   750
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Num Garantia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbltitu 
         BackStyle       =   0  'Transparent
         Caption         =   "Titular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCredGarantRealLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fbActXHistorico As Boolean 'EJVG20160421
Dim fbLoadData As Boolean 'EJVG20160421

Private Sub chkActXHistorico_Click()
    If fbLoadData Then Exit Sub
    
    Dim oDGar As New COMDCredito.DCOMGarantia
    Dim objPista As COMManejador.Pista
    Dim bActiva As Boolean
    
    On Error GoTo ErrActivar

    bActiva = IIf(chkActXHistorico.value = 1, True, False)
    oDGar.ActivaExclusionValidaDescobertura Trim(grdAnalista.TextMatrix(1, 1)), bActiva
    Set oDGar = Nothing
    
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista "190343", GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, IIf(bActiva, gInsertar, gEliminar), IIf(bActiva, "A", "Desa") & "ctiva permitir actualizar garantía así quede descobertura", Trim(grdAnalista.TextMatrix(1, 1)), gCodigoGarantia
    Set objPista = Nothing
    
    MsgBox "Se ha " & IIf(Not bActiva, "des", "") & "activado la exclusión de la validación por descobertura de la garantía N° " & Trim(grdAnalista.TextMatrix(1, 1)), vbInformation, "Aviso"
    
    Set oDGar = Nothing
    Exit Sub
ErrActivar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmdActualizaLegal_Click()
    Dim oGarantia As COMDCredito.DCOMGarantia
    Dim objPista As COMManejador.Pista
    
    If Trim(grdAnalista.TextMatrix(1, 1)) <> "" Then
        Set oGarantia = New COMDCredito.DCOMGarantia
        oGarantia.dUpdateGarantiasLegal Trim(grdAnalista.TextMatrix(1, 1)), 1
        Set oGarantia = Nothing
        
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista "190341", GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Bloqueo Legal", Trim(grdAnalista.TextMatrix(1, 1)), gCodigoGarantia
        Set objPista = Nothing
        
        MsgBox "La Garantía N° " & Trim(grdAnalista.TextMatrix(1, 1)) & " se ha bloqueado con éxito.", vbInformation, "Aviso"
        
        CmdBuscaPersona_Click
    Else
        MsgBox "No se pudo Bloquear la Garantía, Verifique", vbExclamation, "Aviso"
    End If
End Sub
Private Sub CmdBuscaPersona_Click()
    Dim oGarantia As COMDCredito.DCOMGarantia
    Dim rsGarantia As ADODB.Recordset
    Dim lsNumGarant As String
    Dim i As Integer
    
    On Error GoTo ErrBuscar
    fbLoadData = True
    lbltitu.Visible = False
    lbltitular.Visible = False
    lbltitular.Caption = ""
    CmdActualizaLegal.Enabled = False
    chkActXHistorico.value = 0
    chkActXHistorico.Enabled = False 'EJVG20160421
    CmdEditar.Enabled = False
    Call LimpiaFlex(grdAnalista)
    ConfigurarMShComite
    
    If val(txtNumGar.Text) <= 0 Then
        MsgBox "Ud. debe de especificar el Nro. de Garantía", vbInformation, "Aviso"
        EnfocaControl txtNumGar
        Exit Sub
    End If

    lsNumGarant = txtNumGar.Text
    
    Screen.MousePointer = 11
    Set oGarantia = New COMDCredito.DCOMGarantia
    Set rsGarantia = oGarantia.RecuperaGarantiaxBloqueoLegal(lsNumGarant)
    Set oGarantia = Nothing
    Screen.MousePointer = 0
    
    If rsGarantia.EOF Then
        MsgBox "No se ha encontrado la garantía especificada, asegurese" & Chr(13) & "digitarlo correctamente y que la garantía sea Real.", vbInformation, "Aviso"
        RSClose rsGarantia
        Exit Sub
    End If
    
    If Not rsGarantia!bTramiteLegal Then
        MsgBox "El número de garantía que acaba de ingresar no es una Real", vbInformation, "Aviso"
        RSClose rsGarantia
        Exit Sub
    End If
    
    lbltitu.Visible = True
    lbltitular.Visible = True
    lbltitular.Caption = rsGarantia!cPersNombre

    grdAnalista.TextMatrix(i + 1, 0) = i + 1
    grdAnalista.row = i + 1
    grdAnalista.col = 1
                   
    grdAnalista.TextMatrix(i + 1, 1) = rsGarantia!cNumGarant
    grdAnalista.TextMatrix(i + 1, 2) = rsGarantia!cDescripcion
    grdAnalista.TextMatrix(i + 1, 3) = Format(rsGarantia!dTasacion, gsFormatoFechaView)
    grdAnalista.TextMatrix(i + 1, 4) = Format(rsGarantia!nVRM, gsFormatoNumeroView)
    grdAnalista.TextMatrix(i + 1, 5) = IIf(rsGarantia!bBloqueoLegal, "Bloqueado", "Desbloqueado")
    
    chkActXHistorico.value = IIf(rsGarantia!nActXHistorico > 0, 1, 0)
    chkActXHistorico.Enabled = True
    If rsGarantia!bBloqueoLegal Then
        CmdEditar.Enabled = True
    Else
        CmdActualizaLegal.Enabled = True
    End If
    
    fbLoadData = False
    RSClose rsGarantia
    Exit Sub
ErrBuscar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Sub ConfigurarMShComite()
 grdAnalista.Clear
    grdAnalista.Cols = 6
    grdAnalista.Rows = 2
    
    With grdAnalista
        .TextMatrix(0, 1) = "Garantia"
        .TextMatrix(0, 2) = "Descripcion"
        .TextMatrix(0, 3) = "Fecha"
        .TextMatrix(0, 4) = "Valor"
        .TextMatrix(0, 5) = "Estado"
        
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 2500
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        
    End With
End Sub

Private Sub CmdEditar_Click()
 Dim oGarantia As COMDCredito.DCOMGarantia
 Dim objPista As COMManejador.Pista
 Set oGarantia = New COMDCredito.DCOMGarantia
    
    oGarantia.dUpdateGarantiasLegal Trim(grdAnalista.TextMatrix(1, 1)), 0
    
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista "190341", GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gEliminar, "Desbloqueo Legal", Trim(grdAnalista.TextMatrix(1, 1)), gCodigoGarantia
    Set objPista = Nothing
    
    MsgBox "La Garantía N° " & Trim(grdAnalista.TextMatrix(1, 1)) & " se ha desbloqueado con éxito.", vbInformation, "Aviso"
    CmdBuscaPersona_Click
End Sub

Private Sub cmdSalir_Click()
    If lbltitular.Visible = True Then
        lbltitular.Caption = ""
        lbltitu.Visible = False
        lbltitular.Visible = False
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CentraForm Me
    ConfigurarMShComite
    CargarVariables
    
    chkActXHistorico.Visible = fbActXHistorico
    fbLoadData = False
End Sub

Private Sub txtNumGar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl CmdBuscaPersona
    End If
End Sub

Private Sub txtNumGar_LostFocus()
    txtNumGar.Text = Format(txtNumGar, "00000000")
End Sub
Private Sub CargarVariables()
    Dim oCS As New NCOMConstSistema
    fbActXHistorico = IIf(oCS.LeeConstSistema(523) = "0", False, True)
    Set oCS = Nothing
End Sub
