VERSION 5.00
Begin VB.Form frmCapCampanas 
   Caption         =   "Campañas de Ahorros"
   ClientHeight    =   5535
   ClientLeft      =   3525
   ClientTop       =   2475
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7125
   Begin VB.Frame Frame4 
      Caption         =   "Adición de premio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1935
      Left            =   3600
      TabIndex        =   21
      Top             =   3000
      Width           =   3375
      Begin VB.CommandButton cmdDesPre 
         Height          =   495
         Left            =   2640
         Picture         =   "frmCapCampanas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdRegPre 
         Height          =   495
         Left            =   1920
         Picture         =   "frmCapCampanas.frx":0482
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtCantPre 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdNuePre 
         Height          =   495
         Left            =   1200
         Picture         =   "frmCapCampanas.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cboPremio 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad en campaña: "
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Premio:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Premios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2775
      Left            =   3600
      TabIndex        =   20
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtCantidad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   960
         TabIndex        =   25
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton CmdEditPre 
         Height          =   375
         Left            =   2760
         Picture         =   "frmCapCampanas.frx":0A1E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton CmdAddPre 
         Height          =   375
         Left            =   2160
         Picture         =   "frmCapCampanas.frx":0F10
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2280
         Width           =   495
      End
      Begin VB.ListBox LstPre 
         Height          =   2010
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblCantidad 
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2400
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Campañas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2775
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton CmdEditCam 
         Height          =   375
         Left            =   2760
         Picture         =   "frmCapCampanas.frx":0F60
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Editar Campaña"
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton CmdAddCam 
         Height          =   375
         Left            =   2160
         Picture         =   "frmCapCampanas.frx":1452
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Agregar Campaña"
         Top             =   2280
         Width           =   495
      End
      Begin VB.ListBox LstCamp 
         Height          =   2010
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registro/Edición de campaña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   3375
      Begin VB.CommandButton cmdDesCam 
         Height          =   495
         Left            =   2520
         Picture         =   "frmCapCampanas.frx":1794
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdRegCam 
         Height          =   495
         Left            =   1800
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmCapCampanas.frx":1C16
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cboEstCam 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtDesCam 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Estado:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCapCampanas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnIdCamp, lnIdPre As Integer
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by


'
Public Sub Inicia()
    Dim objCapGen As COMDCaptaGenerales.DCOMCampanas
    Dim rsCam As ADODB.Recordset

    'Inicio de objetos del formulario.
    LstCamp.Enabled = True
    CmdAddCam.Enabled = True
    CmdEditCam.Enabled = False
    txtDesCam.Text = ""
    txtDesCam.Enabled = False
    With cboEstCam
        .Enabled = False
        .Clear
        .AddItem "Actva" & Space(100) & "1"
        .AddItem "Inactiva" & Space(100) & "0"
    End With
    cmdRegCam.Enabled = False
    cmdDesCam.Enabled = False

    Set objCapGen = New COMDCaptaGenerales.DCOMCampanas
    Set rsCam = New ADODB.Recordset
    Set rsCam = objCapGen.GetCapCampanas("A")
    If Not rsCam.EOF And Not rsCam.BOF Then
        With LstCamp
            .Clear
            While Not rsCam.EOF
                .AddItem rsCam.Fields("cDescripcion") & Space(100) & rsCam.Fields("bEstado") & "|" & rsCam.Fields("idCampana")
                rsCam.MoveNext
            Wend
        End With
    Else
        LstCamp.AddItem "No existen campañas." & Space(100) & "000"
    End If
    If rsCam.State = adStateOpen Then
        rsCam.Close
    End If
    Set rsCam = Nothing

    LstPre.Enabled = False
    CmdAddPre.Enabled = False
    CmdEditPre.Enabled = False
    LlenarListaPremio
    cboPremio.Enabled = False
    cmdNuePre.Enabled = False
    cmdRegPre.Enabled = False
    cmdDesPre.Enabled = False
    txtCantPre.Enabled = False
    Set objCapGen = Nothing
End Sub
'
Private Sub cboEstCam_Change()
    cmdRegCam.SetFocus
End Sub
'
Private Sub cboPremio_Change()
    txtCantPre.SetFocus
End Sub
'
Private Sub CmdAddCam_Click()
    CmdAddCam.Enabled = False
    txtDesCam.Enabled = True
    txtDesCam.Text = ""
    cboEstCam.Enabled = True
    cmdRegCam.Enabled = True
    cmdDesCam.Enabled = True
    LstCamp.Enabled = False
    txtDesCam.SetFocus
End Sub
'
Private Sub CmdAddPre_Click()
    cboPremio.Enabled = True
    txtCantPre.Enabled = True
    cmdNuePre.Enabled = True
    cmdRegPre.Enabled = True
    cmdDesPre.Enabled = True
    cboPremio.SetFocus
End Sub
'
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
'
Private Sub cmdDesCam_Click()
    LstCamp.Enabled = True
    CmdAddCam.Enabled = True
    txtDesCam.Text = ""
    txtDesCam.Enabled = False
    cboEstCam.Enabled = False
    cmdRegCam.Enabled = False
    LstCamp_Click
End Sub
'
Private Sub cmdDesPre_Click()
    LlenarListaPremio
    cboPremio.Enabled = False
    txtCantPre.Enabled = False
    txtCantPre.Enabled = False
    cmdNuePre.Enabled = False
    cmdRegPre.Enabled = False
    cmdDesPre.Enabled = False
End Sub
'
Private Sub CmdEditCam_Click()
    txtDesCam.Enabled = True
    cboEstCam.Enabled = True
    cmdRegCam.Enabled = True
    cmdDesCam.Enabled = True
End Sub
'
Private Sub cmdNuePre_Click()
    frmCapPremio.Inicia
    frmCapPremio.Show 1
    LlenarListaPremio
End Sub
'
Private Sub cmdRegCam_Click()
    Dim objCam As COMDCaptaGenerales.DCOMCampanas
    If txtDesCam.Text = "" Then
        MsgBox "Debe ingresar una nombre para la Campaña.", vbCritical, "SICMACM"
        txtDesCam.SetFocus
        Exit Sub
    End If
    If MsgBox("Esta seguro de registrar la Campaña", vbInformation + vbYesNo, "SICMACM") = vbYes Then
    'ARCV 24-01-2007
        Set objCam = New COMDCaptaGenerales.DCOMCampanas
        objCam.RegCampana Trim(txtDesCam.Text), CInt(Right(cboEstCam.Text, 1)), "A"
        'By Capi 21012009
        objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Campañas"
        'End by
        Me.Inicia
    End If
End Sub
'
Private Sub cmdRegPre_Click()
'ARCV 24-01-2007
    Dim objCamPre As COMDCaptaGenerales.DCOMCampanas
    If cboPremio.Text = "" Then
        MsgBox "Debe seleccionar un premio.", vbCritical, "SICMACM"
        cboPremio.SetFocus
        Exit Sub
    End If
    If txtCantPre.Text = "" Then
        MsgBox "Debe ingresar ls cantidad del premio para la campaña.", vbCritical, "SICMACM"
        txtCantPre.SetFocus
        Exit Sub
    End If
    If MsgBox("Esta seguro de guardar la información?", vbInformation + vbYesNo, "SICMACM") = vbYes Then
        Set objCamPre = New COMDCaptaGenerales.DCOMCampanas
        objCamPre.RegCampanaPremio lnIdCamp, CInt(Right(cboPremio.Text, 2)), CInt(txtCantPre.Text)
        'By Capi 21012009
        objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Premios"
        'End by
        LstCamp_Click
    End If
    cmdDesPre_Click
End Sub
'
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    'By Capi 20012009
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCapRegistroCampanaPremio
    'End By

End Sub
'
Private Sub LstCamp_Click()
    Dim sCodigo As String
    Dim sCod() As String
    Dim objCamPre As COMDCaptaGenerales.DCOMCampanas
    Dim rsCamPre As ADODB.Recordset
    If CInt(Trim(Right(LstCamp.List(LstCamp.ListIndex), 1))) <> 0 Then
        CmdEditCam.Enabled = True
        CmdAddPre.Enabled = True
        LstPre.Enabled = True
        txtDesCam.Text = RTrim(Mid(LstCamp.List(LstCamp.ListIndex), 1, 100))
        sCodigo = Trim(Right(LstCamp.List(LstCamp.ListIndex), 15))
        If Left(sCodigo, 1) = "V" Then
            cboEstCam.ListIndex = 0
        Else
            cboEstCam.Text = "Inactiva" & Space(100) & "0"
            cboEstCam.ListIndex = 1
        End If
        sCod = Split(sCodigo, "|", 3)
        lnIdCamp = CInt(sCod(1))
        Set objCamPre = New COMDCaptaGenerales.DCOMCampanas
        Set rsCamPre = New ADODB.Recordset
        Set rsCamPre = objCamPre.GetCapCampanaPremio(lnIdCamp)
        If Not rsCamPre.EOF And Not rsCamPre.BOF Then
            With LstPre
                .Clear
                While Not rsCamPre.EOF
                    .AddItem rsCamPre.Fields("cDescripcion") & Space(50) & "|" & rsCamPre.Fields("nCantidad") & "|" & rsCamPre.Fields("nTipoPremio")
                    rsCamPre.MoveNext
                Wend
            End With
        Else
            LstPre.Clear
            LstPre.AddItem "No hay premios en esta campaña" & Space(50) & "0"
        End If
        rsCamPre.Close
    End If
    Set rsCamPre = Nothing
    Set objCamPre = Nothing
End Sub
'
Private Sub LstPre_Click()
    Dim sPremio() As String
    sPremio = Split(LstPre.List(LstPre.ListIndex), "|")
    txtCantidad.Text = Trim(sPremio(1))
End Sub
'
Private Sub txtCantPre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdRegPre.SetFocus
    End If
End Sub
'
Private Sub txtDesCam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboEstCam.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub
'
Private Sub LlenarListaPremio()
    Dim objPre As COMDCaptaGenerales.DCOMCampanas
    Dim rsPre As ADODB.Recordset
    Set objPre = New COMDCaptaGenerales.DCOMCampanas
    Set rsPre = New ADODB.Recordset
    Set rsPre = objPre.GetCapPremios()
    If Not rsPre.EOF And Not rsPre.BOF Then
        With cboPremio
            .Clear
            While Not rsPre.EOF
                .AddItem rsPre.Fields("cDescripcion") & Space(100) & rsPre.Fields("nTipoPremio")
                rsPre.MoveNext
            Wend
        End With
    Else
        cboPremio.AddItem "No existen Premios." & Space(100) & "000"
    End If
End Sub
