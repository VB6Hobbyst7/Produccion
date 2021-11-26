VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmColRecActuacionesProc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recuperaciones - Actuaciones Procesales"
   ClientHeight    =   6480
   ClientLeft      =   1470
   ClientTop       =   1230
   ClientWidth     =   8370
   Icon            =   "frmColRecActuacionesProc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8370
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
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
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame fraHistoria 
      Caption         =   "Actuaciones Procesales"
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
      Height          =   2220
      Left            =   60
      TabIndex        =   21
      Top             =   1920
      Width           =   8265
      Begin SICMACT.FlexEdit grdHistoria 
         Height          =   1905
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   3360
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Usu-Fecha-Comentario-Fecha Aviso-Fecha Vencimiento-Tipo"
         EncabezadosAnchos=   "350-500-2000-5000-2000-2000-2000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraComentario 
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
      Height          =   1695
      Left            =   60
      TabIndex        =   20
      Top             =   4200
      Width           =   8265
      Begin VB.ComboBox cboTipoAct 
         Height          =   315
         ItemData        =   "frmColRecActuacionesProc.frx":030A
         Left            =   5880
         List            =   "frmColRecActuacionesProc.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   240
         Width           =   2235
      End
      Begin VB.CheckBox ChkAlerta 
         Caption         =   "Alerta de Aviso"
         Height          =   255
         Left            =   6360
         TabIndex        =   29
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtComentario 
         Height          =   630
         Left            =   120
         MaxLength       =   235
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   960
         Width           =   8010
      End
      Begin MSMask.MaskEdBox mskFechaVencimiento 
         Height          =   330
         Left            =   3960
         TabIndex        =   24
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
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
      Begin MSMask.MaskEdBox mskFechaAviso 
         Height          =   315
         Left            =   1080
         TabIndex        =   23
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
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
      Begin VB.Label Label11 
         Caption         =   "Tipo :"
         Height          =   195
         Left            =   5400
         TabIndex        =   31
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label7 
         Caption         =   "Comentario:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Vencimiento:"
         Height          =   195
         Left            =   2520
         TabIndex        =   27
         Top             =   255
         Width           =   1485
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Aviso:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   260
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdGrabar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Grabar"
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
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
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
      Left            =   2760
      TabIndex        =   2
      Top             =   6000
      Width           =   1095
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
      Left            =   7080
      TabIndex        =   3
      Top             =   6000
      Width           =   1095
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   465
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   820
      Texto           =   "Crédito"
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar ..."
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   120
      Width           =   1005
   End
   Begin VB.Frame fraCliente 
      Height          =   1365
      Left            =   60
      TabIndex        =   6
      Top             =   540
      Width           =   8265
      Begin VB.Label lblEstudioJur 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   930
         TabIndex        =   19
         Top             =   960
         Width           =   4305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Abogado"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Comision"
         Height          =   195
         Index           =   6
         Left            =   5400
         TabIndex        =   17
         Top             =   960
         Width           =   630
      End
      Begin VB.Label lblComision 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6450
         TabIndex        =   16
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label lblFecIngRecup 
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
         Left            =   6480
         TabIndex        =   15
         Top             =   600
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMontoPrestamo 
         Alignment       =   1  'Right Justify
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
         Left            =   960
         TabIndex        =   14
         Top             =   600
         Width           =   1335
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
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   5475
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCodPers 
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
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblSaldoCapital 
         Alignment       =   1  'Right Justify
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
         Left            =   3720
         TabIndex        =   11
         Top             =   600
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Ingreso Recup."
         Height          =   240
         Left            =   5400
         TabIndex        =   10
         Top             =   630
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo Capital"
         Height          =   195
         Left            =   2640
         TabIndex        =   9
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "Prestamo"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   270
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmColRecActuacionesProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* ACTUACIONES PROCESALES DE RECUPERACIONES
'Archivo:  frmColRecActuacionesProc.frm
'LAYG   :  03/02/2004.
'Resumen:  Nos permite registrar actuaciones procesales de Recuperaciones
Option Explicit

Dim frsPers As ADODB.Recordset
Public fsNuevoConsulta As String
Dim objPista As COMManejador.Pista  '' *** PEAC 20090126


Public Sub Inicia(ByVal psNuevoConsulta As String, Optional ByVal psCredito As String)
    fsNuevoConsulta = psNuevoConsulta
    If fsNuevoConsulta = "C" Then
        Me.cmdBuscar.Visible = False
        Me.cmdNuevo.Visible = False
        Me.CmdGrabar.Visible = False
        Me.cmdCancelar.Visible = False
        Me.AXCodCta.NroCuenta = psCredito
        BuscaDatos (psCredito)
        Me.Show 1
    Else
       Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
       cmdNuevo.Enabled = False
       Me.Show 1
    End If
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
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
'On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValCredito = New COMNColocRec.NColRecValida
        Set lrValida = loValCredito.nValidaExpediente(psNroCredito, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
        If lrValida Is Nothing Then
            cmdNuevo.Enabled = False
            CmdGrabar.Enabled = False
            cmdCancelar.Enabled = True
        Else
            If lrValida Is Nothing Then ' Hubo un Error
                Limpiar
                Set lrValida = Nothing
                Exit Sub
            End If
            'Muestra Datos
            lblCodPers.Caption = lrValida!cCodClie
            lblNomPers.Caption = lrValida!cNomClie
            lblMontoPrestamo = Format(lrValida!nMontoCol, "#,##0.00")
            lblSaldoCapital = Format(lrValida!nSaldo, "#,##0.00")
            lblFecIngRecup = Format(lrValida!dIngRecup, "dd/mm/yyyy")
            lblEstudioJur.Caption = lrValida!cNomAbog
            If Len(lrValida!nTipComis) > 0 Then
                If lrValida!nTipComis = 1 Then
                    lblComision.Caption = Format(lrValida!nComisionValor, "#0.00") & " Mon"
                Else
                    lblComision.Caption = Format(lrValida!nComisionValor, "#0.00") & " % "
                End If
            Else
                lblComision.Caption = ""
            End If
            If fsNuevoConsulta = "N" Then cmdNuevo.Enabled = True
            CmdGrabar.Enabled = False
            cmdCancelar.Enabled = True
            AXCodCta.Enabled = False
            
        End If
        Set lrValida = Nothing
        'Actuaciones Procesales
        Set lrAct = loValCredito.nValidaActuacionesProc(psNroCredito, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
        grdHistoria.Clear
        grdHistoria.FormaCabecera
        grdHistoria.Rows = 2
        txtComentario.Text = ""
        If lrAct Is Nothing Then
            MsgBox "Credito NO tiene Actuaciones Procesales registradas", vbInformation, "Aviso"
        Else
            lnItem = 0
            Do While Not lrAct.EOF
                lsComentario = Trim(lrAct!cComenta)
                lsFecha = Mid(lrAct!cMovNro, 7, 2) & "/" & Mid(lrAct!cMovNro, 5, 2) & "/" & Mid(lrAct!cMovNro, 1, 4)
                lsFecha = lsFecha & " " & Mid(lrAct!cMovNro, 9, 2) & ":" & Mid(lrAct!cMovNro, 11, 2) & ":" & Mid(lrAct!cMovNro, 13, 2)
                lsUsuario = Right(lrAct!cMovNro, 4)
                lnItem = lnItem + 1
                grdHistoria.AdicionaFila
                grdHistoria.TextMatrix(lnItem, 1) = lsUsuario
                grdHistoria.TextMatrix(lnItem, 2) = lsFecha
                grdHistoria.TextMatrix(lnItem, 3) = lsComentario
                grdHistoria.TextMatrix(lnItem, 4) = Format(lrAct!dFechaAviso, "dd/mm/yyyy")
                grdHistoria.TextMatrix(lnItem, 5) = Format(lrAct!dFechaVencimiento, "dd/mm/yyyy")
                '**DAOR 200701254, ********************************************
                grdHistoria.TextMatrix(lnItem, 6) = lrAct!TipoAct
                '**************************************************************
                
                'primero se debe establecer la Fila: .row = intFila y el Color:             lngColor = vbBlue
                
                If CDate(lrAct!dFechaAviso) = CDate(gdFecSis) Then
                    grdHistoria.Row = lnItem
                    lngcolor = RGB(248, 247, 199)
                    For intColumna = 1 To grdHistoria.Cols - 1
                        grdHistoria.Col = intColumna
                        grdHistoria.CellBackColor = lngcolor
                    Next
                End If
                
                lrAct.MoveNext
                
            Loop
        End If
        
        Set lrAct = Nothing
    Set loValCredito = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdBuscar_Click()

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

' Selecciona Estados
lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast

Limpiar ' True
If fsNuevoConsulta = "N" Then cmdNuevo.Enabled = True
CmdGrabar.Enabled = False

'Me.fraDatos.Enabled = False
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
    Limpiar ' True
    cmdNuevo.Enabled = False
    CmdGrabar.Enabled = False
    AXCodCta.Enabled = True
End Sub

Private Sub cmdGrabar_Click()
Dim loContFunct As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim loGrabar As COMNColocRec.NCOMColRecCredito 'NColRecCredito
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lnTipoAct As Integer
'On Error GoTo ControlError
' Valida Datos a Grabar
If fValidaData = False Then
    Exit Sub
End If

If MsgBox(" Grabar Registro de Actuacion Procesal ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        
    'Genera el Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    'DAOR 20070124, *****************************************
    lnTipoAct = Right(cboTipoAct.Text, 2)
    '********************************************************
    
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    If gsProyectoActual = "H" Then
        Set loGrabar = New COMNColocRec.NCOMColRecCredito
            Call loGrabar.nRegistraActProcesales(AXCodCta.NroCuenta, lsFechaHoraGrab, lsMovNro, _
                   Trim(txtComentario.Text), False, Me.mskFechaAviso.Text, Me.mskFechaVencimiento.Text, gsProyectoActual, ChkAlerta.value)
        Set loGrabar = Nothing
    Else
        Set loGrabar = New COMNColocRec.NCOMColRecCredito
            Call loGrabar.nRegistraActProcesales(AXCodCta.NroCuenta, lsFechaHoraGrab, lsMovNro, _
                   Trim(txtComentario.Text), False, Me.mskFechaAviso.Text, Me.mskFechaVencimiento.Text, gsProyectoActual, , lnTipoAct)
            '' *** PEAC 20090126
            objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , AXCodCta.NroCuenta, gCodigoCuenta
                                      
        Set loGrabar = Nothing
    End If
    BuscaDatos (AXCodCta.NroCuenta)
    cmdNuevo.Enabled = True
    CmdGrabar.Enabled = False
    mskFechaAviso.Text = "__/__/____"
    mskFechaVencimiento.Text = "__/__/____"
    Me.txtComentario.Text = ""
    cboTipoAct.ListIndex = -1
    ChkAlerta.value = 0
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdNuevo_Click()
    fraComentario.Enabled = True
    txtComentario.Text = ""
    mskFechaAviso.Text = "__/__/____"
    mskFechaVencimiento.Text = "__/__/____"
    'TxtComentario.SetFocus
    mskFechaAviso.SetFocus
    cmdNuevo.Enabled = False
    CmdGrabar.Enabled = True
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub Limpiar()
lblCodPers = ""
lblNomPers = ""
lblFecIngRecup = ""
lblMontoPrestamo = ""
lblSaldoCapital = ""
lblEstudioJur.Caption = ""
lblComision.Caption = ""

grdHistoria.Clear
grdHistoria.FormaCabecera
grdHistoria.Rows = 2
txtComentario.Text = ""
mskFechaAviso.Text = "__/__/____"
mskFechaVencimiento.Text = "__/__/____"

AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
AXCodCta.SetFocusCuenta

End Sub
Private Function fValidaData() As Boolean
Dim lbOk As Boolean
Dim psMensaje As String
lbOk = True

If Len(Trim(txtComentario.Text)) = 0 Then
    MsgBox "No ingreso el comentario", vbInformation, "Aviso"
    lbOk = False
    Exit Function
End If

psMensaje = ValidaFecha(mskFechaAviso.Text)
If psMensaje <> "" Then
    MsgBox psMensaje, vbInformation, "Aviso"
    lbOk = False
    Exit Function
End If

psMensaje = ValidaFecha(mskFechaVencimiento.Text)
If psMensaje <> "" Then
    MsgBox psMensaje, vbInformation, "Aviso"
    lbOk = False
    Exit Function
End If

If CDate(mskFechaAviso.Text) > CDate(mskFechaVencimiento.Text) Then
    MsgBox "La Fecha de Aviso no puede se mayor a la fecha de Vencimiento", vbInformation, "Aviso"
    lbOk = False
    Exit Function
End If


fValidaData = lbOk
End Function

Private Sub Form_Load()
    CargaComboDatos cboTipoAct, gColocRecTipoActProcesal
    If gsProyectoActual = "H" Then
       ChkAlerta.Visible = True
    End If
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gRecRegistrarActuaProcesal
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub grdHistoria_RowColChange()
    If cmdNuevo.Enabled = True Then
        Me.txtComentario.Text = grdHistoria.TextMatrix(grdHistoria.Row, 3)
        Me.mskFechaAviso.Text = grdHistoria.TextMatrix(grdHistoria.Row, 4)
        Me.mskFechaVencimiento.Text = grdHistoria.TextMatrix(grdHistoria.Row, 5)
    End If
End Sub

Private Sub mskFechaAviso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFechaVencimiento.SetFocus
    End If
End Sub

Private Sub mskFechaVencimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtComentario.SetFocus
    End If
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
     KeyAscii = fgIntfMayusculas(KeyAscii)
     If KeyAscii = 13 Then
        CmdGrabar.SetFocus
     End If
End Sub

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
