VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmColRecRegVisitaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recuperaciones - Registro de Visita de Gestores"
   ClientHeight    =   6690
   ClientLeft      =   1470
   ClientTop       =   435
   ClientWidth     =   11160
   Icon            =   "FrmColRecRegVisitaCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Registro de Visitas"
      TabPicture(0)   =   "FrmColRecRegVisitaCliente.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraHistoria"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdGrabar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancelar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Visitas Realizadas al Cliente"
      TabPicture(1)   =   "FrmColRecRegVisitaCliente.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FeAdj"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdBorrarVisita"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdBorrarVisita 
         Caption         =   "Borrar Registro"
         Height          =   410
         Left            =   4440
         TabIndex        =   38
         Top             =   3755
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
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
         Height          =   375
         Left            =   -73440
         TabIndex        =   36
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   35
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Frame fraHistoria 
         Caption         =   "Datos de la Visita"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   2940
         Left            =   -74760
         TabIndex        =   15
         Top             =   480
         Width           =   10425
         Begin VB.ComboBox cboCondicion 
            Height          =   315
            ItemData        =   "FrmColRecRegVisitaCliente.frx":0342
            Left            =   6600
            List            =   "FrmColRecRegVisitaCliente.frx":0344
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1560
            Width           =   2235
         End
         Begin VB.ComboBox cboPersContactada 
            Height          =   315
            ItemData        =   "FrmColRecRegVisitaCliente.frx":0346
            Left            =   1680
            List            =   "FrmColRecRegVisitaCliente.frx":0348
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   720
            Width           =   2235
         End
         Begin VB.ComboBox cboLugarContacto 
            Height          =   315
            ItemData        =   "FrmColRecRegVisitaCliente.frx":034A
            Left            =   6600
            List            =   "FrmColRecRegVisitaCliente.frx":034C
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   720
            Width           =   2235
         End
         Begin VB.TextBox txtZona 
            Height          =   375
            Left            =   6600
            TabIndex        =   19
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox cboGestion 
            Height          =   315
            ItemData        =   "FrmColRecRegVisitaCliente.frx":034E
            Left            =   1680
            List            =   "FrmColRecRegVisitaCliente.frx":0350
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1560
            Width           =   2235
         End
         Begin VB.TextBox txtMontoCompromiso 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6600
            MaxLength       =   15
            TabIndex        =   17
            Text            =   "0"
            Top             =   1200
            Width           =   2175
         End
         Begin VB.TextBox txtComentario 
            Height          =   375
            Left            =   1680
            MaxLength       =   100
            TabIndex        =   16
            Top             =   2040
            Width           =   7215
         End
         Begin MSMask.MaskEdBox mskFecVisita 
            Height          =   375
            Left            =   1680
            TabIndex        =   23
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFecCompromiso 
            Height          =   375
            Left            =   1680
            TabIndex        =   24
            Top             =   1080
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha visita:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   315
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Compromiso:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   1395
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Observación:"
            Height          =   195
            Left            =   4920
            TabIndex        =   31
            Top             =   1620
            Width           =   945
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Persona Contactada:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   780
            Width           =   1500
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Resultado gestión:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1620
            Width           =   1320
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Lugar de Contacto:"
            Height          =   195
            Left            =   4920
            TabIndex        =   28
            Top             =   780
            Width           =   1365
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Monto Compromiso:"
            Height          =   195
            Left            =   4920
            TabIndex        =   27
            Top             =   1200
            Width           =   1395
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Zona:"
            Height          =   195
            Left            =   4920
            TabIndex        =   26
            Top             =   315
            Width           =   420
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Comentario:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   2040
            Width           =   840
         End
      End
      Begin SICMACT.FlexEdit FeAdj 
         Height          =   3375
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   5953
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         EncabezadosNombres=   "Nº-Fecha Visita-Lugar Contacto-Pers. Contac.-Resul. Gestion-Fec. Comprom.-Monto Compr.-Comentario-Registro-nMovNro"
         EncabezadosAnchos=   "400-1200-1200-1200-1200-1200-1200-2200-2000-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-L-R-L-L-C"
         FormatosEdit    =   "0-0-0-0-0-5-2-0-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "Nº"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   0
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Datos del Crédito"
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
      Height          =   1725
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Width           =   11025
      Begin VB.CommandButton cmdVerVisitas 
         Caption         =   "Ver Visitas"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   37
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   3780
         TabIndex        =   11
         Top             =   270
         Width           =   1005
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   465
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   820
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblUser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   9600
         TabIndex        =   13
         Top             =   240
         Width           =   1275
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Usuario/Gestor:"
         Height          =   195
         Left            =   8400
         TabIndex        =   12
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label lblDireCli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   5880
         TabIndex        =   9
         Top             =   720
         Width           =   4995
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDireGar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   5880
         TabIndex        =   8
         Top             =   1200
         Width           =   4995
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Left            =   5160
         TabIndex        =   7
         Top             =   1110
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Left            =   5160
         TabIndex        =   6
         Top             =   750
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Garante"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1110
         Width           =   570
      End
      Begin VB.Label lblNomGar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   1080
         Width           =   4155
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNomPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   4155
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   750
         Width           =   480
      End
   End
End
Attribute VB_Name = "FrmColRecRegVisitaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* REGISTRO DE VISITA DE GESTORES
'Archivo:  FrmColRecRegVisitaCliente.frm
'PEAC   :  02/07/2012.
'Resumen:  Permite el registro de visitas de gestores de acuerdo a las especificaciones del area.
Option Explicit

Dim frsPers As ADODB.Recordset
Public fsNuevoConsulta As String
Dim objPista As COMManejador.Pista  '' *** PEAC 20090126

Public Sub inicia(ByVal psNuevoConsulta As String, Optional ByVal psCredito As String)
'        Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
'        AXCodCta.SetFocusCuenta
'    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub
Function ValidaDatos() As Boolean
ValidaDatos = True
'If txtNroExp = "" Then
'    ValidaDatos = False
'    MsgBox "Falta Ingresar el Expediente"
'    txtNroExp.SetFocus
'    Exit Function
'End If
End Function

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaDatos(AXCodCta.NroCuenta)
End Sub

Private Sub BuscaDatos(ByVal psNroCredito As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValCredito As COMNColocRec.NColRecValida
Dim lrAct As New ADODB.Recordset
Dim lsComentario As String, lsFecha As String, lsUsuario As String
Dim lnItem As Integer

Dim intColumna As Integer
Dim lngcolor As Long

Dim lsmensaje As String

    Set lrValida = New ADODB.Recordset
    Set loValCredito = New COMNColocRec.NColRecValida
        Set lrValida = loValCredito.ObtieneDatosVisitaCliente(psNroCredito, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
        If lrValida Is Nothing Then
'            cmdNuevo.Enabled = False
            cmdCancelar.Enabled = True
        Else
            If lrValida Is Nothing Then ' Hubo un Error
                limpiar
                Set lrValida = Nothing
                Exit Sub
            End If
            
            lblNomPers.Caption = lrValida!nomtitular
            Me.lblDireCli.Caption = lrValida!diretitular
            Me.lblNomGar.Caption = lrValida!nomgarante
            Me.lblDireGar.Caption = lrValida!diregarante

'            If fsNuevoConsulta = "N" Then cmdNuevo.Enabled = True
'            cmdGrabar.Enabled = False
            cmdCancelar.Enabled = True
            AXCodCta.Enabled = False
            
        End If
        
        Set lrValida = Nothing
    Set loValCredito = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cboCondicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtComentario.SetFocus
    End If
End Sub

Private Sub cboGestion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtZona.SetFocus
    End If
End Sub

Private Sub cboLugarContacto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMontoCompromiso.SetFocus
    End If
End Sub

Private Sub cboPersContactada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskFecCompromiso.SetFocus
    End If
End Sub

Private Sub cmdBorrarVisita_Click()
    Dim loPersCredito  As COMDColocRec.DCOMColRecCredito ' DColRecCredito
    
    If MsgBox("Está seguro de borrar este registro:" + FeAdj.TextMatrix(FeAdj.Row, 1) + " " + FeAdj.TextMatrix(FeAdj.Row, 6), vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
    
    
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
    Call loPersCredito.AnulaRegVisitaGestores(FeAdj.TextMatrix(FeAdj.Row, 9))
    
    Call VerVisitas
    
End Sub

Private Sub CmdBuscar_Click()

Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCredito  As COMDColocRec.DCOMColRecCredito ' DColRecCredito
Dim lrCreditos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

lsEstados = "2020,2021,2022,2030,2031,2032,2201,2202,2203,2204,2205,2206"

limpiar ' True
'If fsNuevoConsulta = "N" Then cmdNuevo.Enabled = True

AXCodCta.Enabled = True
    
If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
        Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)

    Set loPersCredito = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
    limpiar ' True
'    cmdNuevo.Enabled = False
'    cmdGrabar.Enabled = False
    AXCodCta.Enabled = True
End Sub

Private Sub CmdGrabar_Click()
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabar As COMNColocRec.NCOMColRecCredito
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lnTipoAct As Integer
If fValidaData = False Then
    Exit Sub
End If

If MsgBox(" Grabar Registro de Visita de Gestores ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    Set loGrabar = New COMNColocRec.NCOMColRecCredito
        Call loGrabar.RegistraVisitaGestores(lsMovNro, gsOpeCod, AXCodCta.NroCuenta, Format(Me.mskFecVisita.Text, "yyyymmdd"), Trim(Right(cboPersContactada.Text, 3)), _
               IIf(IsDate(mskFecCompromiso.Text), Format(mskFecCompromiso.Text, "yyyymmdd"), ""), Trim(Right(cboGestion, 3)), Trim(txtZona.Text), Trim(IIf(Right(cboLugarContacto, 3) = "", 3, Right(cboLugarContacto, 3))), Format(txtMontoCompromiso.Text, "#,#00.00"), Trim(Right(cboCondicion, 3)), Trim(txtComentario.Text))
        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , AXCodCta.NroCuenta, gCodigoCuenta
    Set loGrabar = Nothing
    
    'BuscaDatos (AXCodCta.NroCuenta)

    limpiar

Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub limpiar()

lblNomPers = ""
cboPersContactada.ListIndex = -1
cboGestion.ListIndex = -1
txtComentario.Text = ""
txtZona.Text = ""
cboLugarContacto.ListIndex = -1
txtMontoCompromiso.Text = "0.00"
cboCondicion.ListIndex = -1

Me.lblNomPers.Caption = ""
Me.lblNomGar.Caption = ""
Me.lblDireCli.Caption = ""
Me.lblDireGar.Caption = ""

mskFecVisita.Text = "__/__/____"
mskFecCompromiso.Text = "__/__/____"

AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
AXCodCta.SetFocusCuenta

End Sub
Private Function fValidaData() As Boolean
Dim lbOk As Boolean
Dim psMensaje As String
lbOk = True

psMensaje = ValidaFecha(mskFecVisita.Text)
If psMensaje <> "" Then
    MsgBox psMensaje, vbInformation, "Aviso"
    lbOk = False
    mskFecVisita.SetFocus
    Exit Function
End If
If cboPersContactada.ListIndex = -1 Then
    MsgBox "No ingreso la persona contactada", vbInformation, "Aviso"
    lbOk = False
    cboPersContactada.SetFocus
    Exit Function
End If

'psMensaje = ValidaFecha(mskFecCompromiso.Text)
'If psMensaje <> "" Then
'    MsgBox psMensaje, vbInformation, "Aviso"
'    lbOk = False
'    mskFecCompromiso.SetFocus
'    Exit Function
'End If

psMensaje = ValidaFecha(mskFecCompromiso.Text)
If IsDate(psMensaje) Then
    If CDate(mskFecVisita.Text) > CDate(mskFecCompromiso.Text) Then
        MsgBox "La Fecha de Visita no puede se mayor a la fecha de compromiso.", vbInformation, "Aviso"
        lbOk = False
        mskFecVisita.SetFocus
        Exit Function
    End If
End If

If cboGestion.ListIndex = -1 Then
    MsgBox "No ingreso el resultado de gestion", vbInformation, "Aviso"
    lbOk = False
    cboGestion.SetFocus
    Exit Function
End If
If Len(Trim(txtZona.Text)) = 0 Then
    MsgBox "No ingreso la zona", vbInformation, "Aviso"
    lbOk = False
    txtZona.SetFocus
    Exit Function
End If

'If cboLugarContacto.ListIndex = -1 Then
'    MsgBox "No ingreso el lugar de contacto", vbInformation, "Aviso"
'    lbOk = False
'    cboLugarContacto.SetFocus
'    Exit Function
'End If

'If CDbl(txtMontoCompromiso.Text) <= 0 Then
'    MsgBox "Ingrese correctamente el monto de compromiso.", vbInformation, "Aviso"
'    lbOk = False
'    txtMontoCompromiso.SetFocus
'    Exit Function
'End If

If cboCondicion.ListIndex = -1 Then
    MsgBox "No ingreso la condicion.", vbInformation, "Aviso"
    lbOk = False
    cboCondicion.SetFocus
    Exit Function
End If
If Len(Trim(txtComentario.Text)) = 0 Then
    MsgBox "No ingreso el comentario", vbInformation, "Aviso"
    lbOk = False
    txtComentario.SetFocus
    Exit Function
End If

fValidaData = lbOk
End Function

Private Sub Command1_Click()

End Sub

Private Sub cmdVerVisitas_Click()
       
    Call VerVisitas
    
End Sub

Private Sub VerVisitas()

    Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
    Dim lrVisit As ADODB.Recordset
    
    If Len(AXCodCta.NroCuenta) = 0 Then
        MsgBox "Ingrese una cuenta", vbInformation + vbOKOnly, "Mensaje"
        Exit Sub
    End If
    
    Set lrVisit = loPersCredito.ObtieneVisitaDeGestores(AXCodCta.NroCuenta)
    Set loPersCredito = Nothing
    
    If lrVisit.EOF Then
        MsgBox "No se encontraron datos.", vbInformation + vbOKOnly, "Mensaje"
        Exit Sub
    End If
    
    Me.SSTab1.Tab = 1
    
    FeAdj.Clear
    FeAdj.FormaCabecera
    FeAdj.Rows = 2
    FeAdj.rsFlex = lrVisit

End Sub

'Private Sub FeAdj_Click(ByVal sender As System.Object, ByVal e As MouseEventArgs)
''Private Sub img_click(ByVal sender As System.Object, ByVal e As MouseEventArgs)
'If e.Button = Windows.Forms.MouseButtons.Right Then
'    '-- AQUI LO QUE HARA AL HACER CLICK DERECHO
'    MsgBox "derecho"
'Else
'    If e.Button = Windows.Forms.MouseButtons.Left Then
'    '- AQUI LO QUE HARA AL HACER CLICK IZQUIERDO
'    MsgBox "isquierdo"
'    End If
'End If
''End Sub
'
'End Sub

Private Sub Form_Load()
    CargaComboDatos cboPersContactada, 3315
    CargaComboDatos cboGestion, 3316
    CargaComboDatos cboLugarContacto, 3317
    CargaComboDatos cboCondicion, 3318

'    If gsProyectoActual = "H" Then
'       ChkAlerta.Visible = True
'    End If
'
    Set objPista = New COMManejador.Pista
    'gsOpeCod = gRecRegistrarActuaProcesal
    gsOpeCod = "191121" ' --> Recuperaciones - Registro de Visita de Gestores
    
    Me.lblUser.Caption = gsCodUser
    
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
'    AXCodCta.SetFocusCuenta
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    
    Me.SSTab1.Tab = 0
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub lbl_Click()

End Sub

'Private Sub grdHistoria_RowColChange()
'    If cmdNuevo.Enabled = True Then
'        Me.txtComentario.Text = grdHistoria.TextMatrix(grdHistoria.Row, 3)
'        Me.mskFechaAviso.Text = grdHistoria.TextMatrix(grdHistoria.Row, 4)
'        Me.mskFechaVencimiento.Text = grdHistoria.TextMatrix(grdHistoria.Row, 5)
'    End If
'End Sub

'Private Sub mskFechaAviso_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.mskFechaVencimiento.SetFocus
'    End If
'End Sub

'Private Sub mskFechaVencimiento_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
''        Me.txtComentario.SetFocus
'    End If
'End Sub

'Private Sub txtComentario_KeyPress(KeyAscii As Integer)
'     KeyAscii = fgIntfMayusculas(KeyAscii)
'     If KeyAscii = 13 Then
'        cmdGrabar.SetFocus
'     End If
'End Sub

Private Sub CargaComboDatos(ByVal combo As ComboBox, ByVal pnValor As Integer)
    Dim oConst As COMDConstantes.DCOMConstantes
    Dim rs As New ADODB.Recordset
    Set oConst = New COMDConstantes.DCOMConstantes
        Set rs = oConst.ObtenerVarRecuperaciones(pnValor)
        combo.Clear
        If Not (rs.EOF And rs.BOF) Then
            Do Until rs.EOF
                combo.AddItem rs(0)
                rs.MoveNext
            Loop
        End If
    Set oConst = Nothing
    Set rs = Nothing
End Sub

Private Sub mskFecCompromiso_GotFocus()
    fEnfoque mskFecCompromiso
End Sub

Private Sub mskFecCompromiso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboGestion.SetFocus
    End If
End Sub

Private Sub mskFecVisita_GotFocus()
    fEnfoque mskFecVisita
End Sub

Private Sub mskFecVisita_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboPersContactada.SetFocus
    End If
End Sub

Private Sub txtComentario_GotFocus()
    fEnfoque txtComentario
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub

Private Sub txtMontoCompromiso_Change()
    If txtMontoCompromiso.Text = "" Then txtMontoCompromiso.Text = "0"
End Sub

Private Sub txtMontoCompromiso_GotFocus()
    fEnfoque txtMontoCompromiso
End Sub

Private Sub txtMontoCompromiso_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMontoCompromiso, KeyAscii)
    If KeyAscii = 13 Then cboCondicion.SetFocus
End Sub

Private Sub txtZona_GotFocus()
    fEnfoque txtZona
End Sub

Private Sub txtZona_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboLugarContacto.SetFocus
    End If
End Sub
