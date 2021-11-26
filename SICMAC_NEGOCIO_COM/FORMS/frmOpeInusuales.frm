VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOpeInusuales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operaciones Inusuales"
   ClientHeight    =   6345
   ClientLeft      =   3960
   ClientTop       =   2475
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8100
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   6600
      TabIndex        =   16
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton cmbBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
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
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
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
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
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
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtNroROS 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   3240
         Width           =   2415
      End
      Begin VB.ComboBox cmbResultado 
         Height          =   315
         ItemData        =   "frmOpeInusuales.frx":0000
         Left            =   1080
         List            =   "frmOpeInusuales.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2840
         Width           =   2415
      End
      Begin VB.TextBox txtMotivo 
         Height          =   855
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1903
         Width           =   5055
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   330
         Left            =   1080
         TabIndex        =   1
         Top             =   1094
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dcAgencia 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   1506
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1080
         TabIndex        =   18
         Top             =   712
         Width           =   5055
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1080
         TabIndex        =   17
         Top             =   345
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Nº ROS"
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
         Left            =   120
         TabIndex        =   15
         Top             =   3255
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Resultado"
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
         Left            =   120
         TabIndex        =   14
         Top             =   2870
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Motivo"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1903
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre"
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
         Left            =   120
         TabIndex        =   12
         Top             =   742
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Codigo"
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
         Left            =   120
         TabIndex        =   11
         Top             =   375
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Agencia"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1536
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1132
         Width           =   615
      End
   End
   Begin SICMACT.FlexEdit grdHistoria 
      Height          =   2145
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   3784
      Cols0           =   9
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Fecha-Nombre-Agencia-Resultado-Nro_ROS-Motivo-cUltimaActualizacion-cPersCod"
      EncabezadosAnchos=   "350-1200-3000-1300-1200-1200-3000-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-C-L-L-L-C-L-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   345
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmOpeInusuales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnEstado  As Integer '0=inactivo;1=Guardar;2=Modificar
Dim lnBuscar As Integer
Dim lnFilaAnt As Integer
Dim lnFilaAct As Integer
Private Sub cmbBuscar_Click()
    Dim oPers As COMDpersona.UCOMPersona
    Dim rsGrilla As Recordset
    LimpiarPantalla
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        
        If ObtenerOpeInusualPersona_Lista(oPers.sPersCod) Then
           lnBuscar = 1
           lnFilaAnt = 0
           CargarTamanioFormulario 6850
           'cmdGrabar.Enabled = True
           CmdCancelar.Enabled = True
           'cmdModificar.Enabled = True
           cmdNuevo.Enabled = False
            
        Else
            CargarTamanioFormulario 4350
            MsgBox "La Persona No se Encuentra en la Lista de Operaciones Inusuales", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Sub cmbResultado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNroROS.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
    LimpiarPantalla
End Sub
Private Sub LimpiarPantalla()
    Me.lblCodigo.Caption = ""
    Me.lblNombre.Caption = ""
    Me.txtFecha.Text = "__/__/____"
    Me.dcAgencia.BoundText = "0"
    Me.txtMotivo.Text = ""
    Me.cmbResultado.ListIndex = -1
    Me.txtNroROS.Text = ""
    cmdGrabar.Enabled = False
    CmdCancelar.Enabled = False
    cmdNuevo.Enabled = True
    'If lnEstado = 1 Then
        CargarTamanioFormulario 4350
    'Else
    If lnEstado = 2 Then
        cmdModificar.Enabled = False
    End If
    lnEstado = 0
    lnBuscar = 0
End Sub


Private Sub cmdGrabar_Click()
    Dim oServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim sMovNro As String
    
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set clsMov = New COMNContabilidad.NCOMContFunciones
        
    If Not validaDatos Then
        Exit Sub
    End If
    If MsgBox("Esta Seguro de Guardar los Datos", vbYesNo, "Aviso") = vbYes Then
        If lnEstado = 2 Then
            oServ.modificarOperacionInusualPersona grdHistoria.TextMatrix(grdHistoria.Row, 7)
        End If
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        oServ.agregarOperacionInusual Trim(lblCodigo.Caption), CDate(txtFecha.Text), dcAgencia.BoundText, txtMotivo.Text, cmbResultado.ItemData(cmbResultado.ListIndex), txtNroROS.Text, sMovNro
        'LimpiarPantalla
        lnEstado = 0
        cmdGrabar.Enabled = False
        CmdCancelar.Enabled = False
        CargarTamanioFormulario 6850
        ObtenerOpeInusualPersona_Lista lblCodigo.Caption
    End If
    Set clsMov = Nothing
    Set oServ = Nothing
    
End Sub
Private Function ObtenerOpeInusualPersona_Lista(ByVal psPersCod As String) As Boolean
     ObtenerOpeInusualPersona_Lista = False
     Dim oServ As COMNCaptaServicios.NCOMCaptaServicios
     Dim rs As Recordset
     Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
     Set rs = New Recordset
     Set rs = oServ.obtenerOperacionInusualPersona(psPersCod, "")
     If Not rs.BOF Or Not rs.EOF Then
        Set Me.grdHistoria.Recordset = rs
        ObtenerOpeInusualPersona_Lista = True
     End If
End Function
Private Sub ObtenerOpeInusualPersona()
     
     Dim oServ As COMNCaptaServicios.NCOMCaptaServicios
     Dim rs As Recordset
     Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
     Set rs = New Recordset
     Set rs = oServ.obtenerOperacionInusualPersona("", grdHistoria.TextMatrix(grdHistoria.Row, 7))
     If Not rs.BOF Or Not rs.EOF Then
        lblCodigo.Caption = rs!cPersCod
        lblNombre.Caption = PstaNombre(rs!Nombre)
        txtFecha.Text = rs!Fecha
        dcAgencia.BoundText = Right(rs!Agencia, 2)
        txtMotivo.Text = rs!motivo
        cmbResultado.ListIndex = CInt(Right(rs!Resultado, 1)) - 1
        txtNroROS.Text = rs!Nro_ROS
     End If
End Sub
Private Function validaDatos() As Boolean
    validaDatos = True
    If ValidaFecha(txtFecha.Text) <> "" Then
        MsgBox "Ingrese una Fecha Valida", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
    End If
    If dcAgencia.BoundText = "0" Then
        MsgBox "Seleccione la Agencia", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
    End If
    If txtMotivo.Text = "" Then
        MsgBox "Ingrese el Motivo", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
    End If
    If cmbResultado.ListIndex = -1 Then
        MsgBox "Seleccione un Resultado", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
    End If
    If txtNroROS.Text = "" Then
        MsgBox "Ingrese el Nro ROS", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
    End If
End Function

Private Sub cmdModificar_Click()
    lnEstado = 2
    cmdGrabar.Enabled = True
    CmdCancelar.Enabled = True
    ObtenerOpeInusualPersona
    cmdModificar.Enabled = False
    
End Sub

Private Sub cmdNuevo_Click()
    Dim oPers As COMDpersona.UCOMPersona
    Dim rsGrilla As Recordset
    LimpiarPantalla
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        lnEstado = 1
        lblCodigo.Caption = oPers.sPersCod
        lblNombre.Caption = oPers.sPersNombre
        'txtFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
        If ObtenerOpeInusualPersona_Lista(oPers.sPersCod) Then
           CargarTamanioFormulario 6850
        Else
            CargarTamanioFormulario 4350
        End If
        cmdGrabar.Enabled = True
        CmdCancelar.Enabled = True
        cmdModificar.Enabled = False
        txtFecha.SetFocus
    End If
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub dcAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMotivo.SetFocus
    End If
End Sub

Private Sub Form_Load()
    CargarTamanioFormulario 4350
    CargarAgencias
    lnEstado = 0
    lnBuscar = 0
End Sub
Private Sub CargarTamanioFormulario(ByVal altura As Integer)
    Me.Height = altura
End Sub
Private Sub CargarAgencias()
    Dim rsAgencia As New ADODB.Recordset
    Dim objCOMNCredito As COMNCredito.NCOMBPPR
    Set objCOMNCredito = New COMNCredito.NCOMBPPR
    Set rsAgencia.DataSource = objCOMNCredito.getCargarAgencias
    dcAgencia.BoundColumn = "cAgeCod"
    dcAgencia.DataField = "cAgeCod"
    Set dcAgencia.RowSource = rsAgencia
    dcAgencia.ListField = "cAgeDescripcion"
    dcAgencia.BoundText = 0
End Sub

Private Sub grdHistoria_Click()
     If lnEstado = 2 Then
     cmdModificar.Enabled = True
     End If
End Sub

Private Sub grdHistoria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdModificar.SetFocus
    End If
End Sub
Private Sub grdHistoria_OnRowChange(pnRow As Long, pnCol As Long)
    If lnBuscar = 1 Then
        lnFilaAct = pnRow
        If lnFilaAnt <> 0 Then
            grdHistoria.Row = lnFilaAnt
            ColoreaCelda vbWhite, vbBlack
        End If
        grdHistoria.Row = lnFilaAct
        lnFilaAnt = lnFilaAct
        ColoreaCelda &HC0FFC0, vbBlack
        cmdModificar.Enabled = True
        
    End If
End Sub
Private Sub ColoreaCelda(ByVal colorCelda As OLE_COLOR, ByVal colorFuente As OLE_COLOR)
   
    With grdHistoria
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Col = i
            .CellBackColor = colorCelda
            .CellForeColor = colorFuente
        Next i
        .Col = 1
   End With
    
End Sub



Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dcAgencia.SetFocus
    End If
End Sub
Private Sub txtMotivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbResultado.SetFocus
    End If
End Sub
Private Sub txtNroROS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdGrabar.Enabled = True Then
            cmdGrabar.SetFocus
        End If
    End If
End Sub
