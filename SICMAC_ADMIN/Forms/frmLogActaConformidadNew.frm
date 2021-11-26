VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogActaConformidadNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Titulo"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13185
   Icon            =   "frmLogActaConformidadNew.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstActaConformidad 
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Acta de Conformidad"
      TabPicture(0)   =   "frmLogActaConformidadNew.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmActaDatoArea"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmActaDatoMoneda"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmActaDatoDocRef"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmActaDatoNActa"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmActaDatoProveedor"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmActaDatoCompra"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdConforme"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCancelar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraComprobante"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.Frame fraComprobante 
         Caption         =   "Datos del Comprobante"
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
         Left            =   240
         TabIndex        =   40
         Top             =   1080
         Width           =   12495
         Begin VB.Label lblEmisionComp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   10080
            TabIndex        =   46
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label5 
            Caption         =   "Emisión :"
            Height          =   255
            Left            =   9360
            TabIndex        =   45
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblNumeroComprobante 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5640
            TabIndex        =   44
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label lblTpoComprobante 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   600
            TabIndex        =   43
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label4 
            Caption         =   "Nº :"
            Height          =   255
            Left            =   5160
            TabIndex        =   42
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo :"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11160
         TabIndex        =   39
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CommandButton cmdConforme 
         Caption         =   "&Confome"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11160
         TabIndex        =   38
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Frame frmActaDatoCompra 
         Caption         =   "Datos de Compra"
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
         Height          =   1215
         Left            =   240
         TabIndex        =   31
         Top             =   3000
         Width           =   10785
         Begin VB.TextBox txtCompraNGuia 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   9240
            MaxLength       =   1225
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtCompraDescripcion 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   1080
            MaxLength       =   1225
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   240
            Width           =   7335
         End
         Begin VB.TextBox txtCompraObservacion 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   1080
            MaxLength       =   1225
            ScrollBars      =   2  'Vertical
            TabIndex        =   37
            Top             =   720
            Width           =   9495
         End
         Begin VB.Label Label2 
            Caption         =   "Nº Guía:"
            Height          =   255
            Left            =   8520
            TabIndex        =   34
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label79 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   330
            Width           =   975
         End
         Begin VB.Label Label76 
            Caption         =   "Observa.:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.Frame frmActaDatoProveedor 
         Caption         =   "Datos del Proveedor"
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
         Height          =   1035
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   12500
         Begin VB.TextBox txtProveedorDocTpo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   7920
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtProveedorDocNro 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   10320
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtProveedorCod 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtProveedorNombre 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   4095
         End
         Begin VB.TextBox txtProveedorCtaNro 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox txtProveedorCtaMoneda 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtProveedorCtaInstitucionNombre 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   600
            Width           =   3850
         End
         Begin VB.TextBox txtProveedorCtaInstitucionCod 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label68 
            Caption         =   "Proveedor:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label69 
            Caption         =   "Tipo Doc.:"
            Height          =   255
            Left            =   7080
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label70 
            Caption         =   "N° Doc.:"
            Height          =   255
            Left            =   9600
            TabIndex        =   22
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label71 
            Caption         =   "Cuenta:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   620
            Width           =   615
         End
         Begin VB.Label Label72 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   4320
            TabIndex        =   26
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label73 
            Caption         =   "Institución:"
            Height          =   255
            Left            =   6240
            TabIndex        =   28
            Top             =   615
            Width           =   855
         End
      End
      Begin VB.Frame frmActaDatoNActa 
         Caption         =   "N° Acta"
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
         Height          =   615
         Left            =   10680
         TabIndex        =   14
         Top             =   480
         Width           =   2055
         Begin VB.TextBox txtActaNro 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame frmActaDatoDocRef 
         Caption         =   "Doc. Referencia"
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
         Height          =   615
         Left            =   8040
         TabIndex        =   12
         Top             =   480
         Width           =   2505
         Begin VB.TextBox txtDocReferencia 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   500
            TabIndex        =   13
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame frmActaDatoMoneda 
         Caption         =   "Moneda"
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
         Height          =   615
         Left            =   6840
         TabIndex        =   10
         Top             =   480
         Width           =   1185
         Begin VB.TextBox txtMoneda 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   95
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame frmActaDatoArea 
         Caption         =   "Datos de Área"
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
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   6495
         Begin VB.TextBox txtAreaAgeCod 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtAreaAgeNombre 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1250
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtSubAreaDescripcion 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2760
            MaxLength       =   235
            TabIndex        =   9
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Área:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   270
            Width           =   375
         End
         Begin VB.Label Label67 
            Caption         =   "-"
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
            Left            =   2640
            TabIndex        =   8
            Top             =   240
            Width           =   135
         End
      End
   End
   Begin TabDlg.SSTab sstCompPendiente 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Comprobantes Pendientes"
      TabPicture(0)   =   "frmLogActaConformidadNew.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feComprobante"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdDarConformidad"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.CommandButton cmdDarConformidad 
         Caption         =   "&Dar Conformidad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   3000
         Width           =   1935
      End
      Begin Sicmact.FlexEdit feComprobante 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   12735
         _extentx        =   22463
         _extenty        =   4260
         cols0           =   11
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Proveedor-Tipo-Numero-Emisión-Area Usuaria-Observaciones-Moneda-Monto-nMovNro-DocOrigen"
         encabezadosanchos=   "300-1800-1200-1200-1200-1500-3000-1200-1200-0-0"
         font            =   "frmLogActaConformidadNew.frx":0342
         font            =   "frmLogActaConformidadNew.frx":036A
         font            =   "frmLogActaConformidadNew.frx":0392
         font            =   "frmLogActaConformidadNew.frx":03BA
         font            =   "frmLogActaConformidadNew.frx":03E2
         fontfixed       =   "frmLogActaConformidadNew.frx":040A
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-C-C-C-C-C-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0-0-0-0-0"
         textarray0      =   "#"
         colwidth0       =   300
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmLogActaConformidadNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************
'Nombre : frmLogActaConformidadNew
'Descripcion:Formulario para el Registrode Actas de conformidad
'Creacion: PASIERS0772014
'*****************************
Option Explicit
Dim gsOpeCod As String
Dim fnMoneda As Integer
Dim fnMovNro As Long
Dim fsAreaAgeCod As String
Dim fntpodocorigen As Integer
Dim fnDocTpo As Integer
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Public Sub Inicio(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    gsOpeCod = psOpeCod
    fnMoneda = CInt(Mid(psOpeCod, 3, 1))
    Me.Caption = UCase(psOpeDesc)
    Me.Show 1
End Sub
Private Sub cmdCancelar_Click()
    EstadoControles 0
    LimpiaControles
    cmdDarConformidad.SetFocus
End Sub
Private Sub cmdConforme_Click()
    Dim olog As NLogGeneral
    Dim lnMovNro As Long
    Dim lsNroActaConformidad As String
    Dim lsMovNro As String
    
    On Error GoTo ErrCmdConforme
    If Len(Trim(txtAreaAgeCod.Text)) = 0 Then
        MsgBox "La presente conformidad no cuenta con Área Agencia", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtActaNro.Text)) = 0 Then
        MsgBox "No se ha conseguido el correlativo del Nro de Acta", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtProveedorCod.Text)) = 0 Then
        MsgBox "No se cuenta con Proveedor en la presente operación", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtCompraDescripcion.Text)) = 0 Then
        MsgBox "Ud. debe ingresar una descripción", vbInformation, "Aviso"
        txtCompraDescripcion.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtCompraObservacion.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la observación respectiva", vbInformation, "Aviso"
        txtCompraObservacion.SetFocus
        Exit Sub
    End If
    lsNroActaConformidad = txtActaNro.Text
    If MsgBox("¿Esta seguro de guardar el Acta de Conformidad Digital?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set olog = New NLogGeneral
    Screen.MousePointer = 11
    cmdConforme.Enabled = False 'PASI20151228
        lnMovNro = olog.GrabarActaConformidadNew(gdFecSis, Right(gsCodAge, 2), gsCodUser, gsOpeCod, fntpodocorigen, fnDocTpo, txtDocReferencia.Text, _
                                                txtAreaAgeCod.Text, Trim(txtSubAreaDescripcion.Text), fnMoneda, lsNroActaConformidad, Trim(Replace(Replace(Me.txtCompraDescripcion.Text, Chr(10), ""), Chr(13), "")), _
                                                Trim(Replace(Replace(txtCompraObservacion.Text, Chr(10), ""), Chr(13), "")), lsMovNro, Trim(txtCompraNGuia.Text), fnMovNro)
    Screen.MousePointer = 0
    
    If lnMovNro = 0 Then
        MsgBox "Ha ocurrido un error al registrar el Acta de Conformidad", vbCritical, "Aviso"
        olog = Nothing
        Exit Sub
    End If
    
    MsgBox "Se ha registrado el Acta de Conformidad Nro. " & lsNroActaConformidad & " con éxito", vbInformation, "Aviso"
    ImprimeActaConformidadPDFNew lnMovNro, fnMoneda, fntpodocorigen
    CargaCompPendientes
    cmdConforme.Enabled = True 'PASI20151228
    EstadoControles 0
    'ARLO 20160126 ***
    Dim lsMonedas As String
    Set objPista = New COMManejador.Pista
    If (fnMoneda = 1) Then
    lsMonedas = "SOLES"
    Else
    lsMonedas = "DOLARES"
    End If
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se ha registrado el Acta de Conformidad Nro. " & lsNroActaConformidad & " en Moneda " & lsMonedas
    Set objPista = Nothing
    '***
     If MsgBox("¿Desea registrar otra Acta de Conformidad?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
        cmdCancelar_Click
    Else
        Unload Me
    End If
    
    Exit Sub
ErrCmdConforme:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdDarConformidad_Click()
    Dim olog As DLogGeneral
    Dim oMov As DMov
    Dim lsCorrelativo As String
    Dim rs As ADODB.Recordset
    Dim rsNActa As ADODB.Recordset
    On Error GoTo ErrcmdDarConformidad
    If feComprobante.TextMatrix(feComprobante.row, 1) <> "" Then
        fntpodocorigen = feComprobante.TextMatrix(feComprobante.row, 10)
        fnMovNro = feComprobante.TextMatrix(feComprobante.row, 9)
        Screen.MousePointer = 11
        Select Case fntpodocorigen
            Case LogTipoDocOrigenComprobante.OrdenCompra _
                , LogTipoDocOrigenComprobante.ContratoAdqBienes _
                , LogTipoDocOrigenComprobante.CompraLibre
                fnDocTpo = LogTipoActaConformidad.gActaRecepcionBienes
            Case LogTipoDocOrigenComprobante.OrdenServicio _
                , LogTipoDocOrigenComprobante.ContratoServicio _
                , LogTipoDocOrigenComprobante.ContratoArrendamiento _
                , LogTipoDocOrigenComprobante.ContratoObra _
                , LogTipoDocOrigenComprobante.Serviciolibre
                fnDocTpo = LogTipoActaConformidad.gActaConformidadServicio
        End Select
        EstadoControles 1
        Set olog = New DLogGeneral
        Set oMov = New DMov
        Set rs = New ADODB.Recordset
        
        If fntpodocorigen = LogTipoDocOrigenComprobante.OrdenCompra _
            Or fntpodocorigen = LogTipoDocOrigenComprobante.OrdenServicio Then
            Set rs = olog.ListaOrdenxActaConformidad(fnMovNro)
        ElseIf fntpodocorigen = LogTipoDocOrigenComprobante.ContratoAdqBienes _
            Or fntpodocorigen = LogTipoDocOrigenComprobante.ContratoArrendamiento _
            Or fntpodocorigen = LogTipoDocOrigenComprobante.ContratoObra _
            Or fntpodocorigen = LogTipoDocOrigenComprobante.ContratoServicio Then
            Set rs = olog.ListaContratoxActaConformidad(fnMovNro, fsAreaAgeCod)
        ElseIf fntpodocorigen = LogTipoDocOrigenComprobante.CompraLibre Or fntpodocorigen = LogTipoDocOrigenComprobante.Serviciolibre Then
            Set rs = olog.ListaOrdenLibrexActaConformidad(fnMovNro)
        End If
        
         lsCorrelativo = oMov.GetExisteNActaConformidadxContrato(fnMovNro)
        If lsCorrelativo = "" Then
            lsCorrelativo = oMov.GetCorrelativoActaConformidad(fnDocTpo, Right(gsCodAge, 2), CStr(Year(gdFecSis)))
        'Else
            'lsCorrelativo = oMov.GetCorrelativoActaConformidad(fnDocTpo, Right(gsCodAge, 2), CStr(Year(gdFecSis)))
        End If
        If Not rs.EOF Then
            EstablecerDatosActaConformidad rs!cAreaAgeCod, rs!cAreaAgeDesc, rs!cMoneda, rs!cDocReferencia, rs!DocDesc, rs!NroDoc, Format(rs!Emision, "dd/mm/yyyy"), lsCorrelativo, rs!cProveedorCod, rs!cProveedorNombre, rs!cDocTpo, rs!cDocNro, rs!cCtaCodAhorro, IIf(rs!cCtaCodAhorro <> "", rs!cMoneda, ""), rs!cInstitucionCod, rs!cInstitucionNombre, rs!cMovDesc
        End If
        txtSubAreaDescripcion.SetFocus
        Screen.MousePointer = 0
        Set rs = Nothing
        Set olog = Nothing
        Set oMov = Nothing
    Else
        MsgBox "No Hay Comprobantes para dar Conformidad.", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
 Exit Sub
ErrcmdDarConformidad:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub EstablecerDatosActaConformidad(Optional ByVal psAreaAgeCod As String = "", Optional ByVal psAreaAgeNombre As String = "", _
                                            Optional ByVal psMoneda As String = "", Optional ByVal psDocReferencia As String = "", _
                                            Optional ByVal psDocCompDesc As String = "", Optional ByVal psNroComp As String = "", Optional ByVal psEmision As String = "", _
                                            Optional ByVal psActaNro As String = "", Optional ByVal psProveedorCod As String = "", _
                                            Optional ByVal psProveedorNombre As String = "", Optional ByVal psProveedorDocTpo As String = "", _
                                            Optional ByVal psProveedorDocNro As String = "", Optional ByVal psProveedorCtaNro As String = "", _
                                            Optional ByVal psProveedorCtaMoneda As String = "", Optional ByVal psProveedorCtaInstitucionCod As String = "", _
                                            Optional ByVal psProveedorCtaInstitucionNombre As String = "", Optional ByVal psCompraDescripcion As String = "", _
                                            Optional ByVal psCompraObservacion As String = "")
    If psAreaAgeCod <> "" Then
        txtAreaAgeCod.Text = psAreaAgeCod
    End If
    If psAreaAgeNombre <> "" Then
        txtAreaAgeNombre.Text = psAreaAgeNombre
    End If
    If psMoneda <> "" Then
        txtMoneda.Text = psMoneda
    End If
    If psDocReferencia <> "" Then
        txtDocReferencia.Text = psDocReferencia
    End If
    If psDocCompDesc <> "" Then
        lblTpoComprobante.Caption = psDocCompDesc
    End If
    If psNroComp <> "" Then
        lblNumeroComprobante.Caption = psNroComp
    End If
    If psEmision <> "" Then
        lblEmisionComp.Caption = psEmision
    End If
    If psActaNro <> "" Then
        txtActaNro.Text = psActaNro
    End If
    If psProveedorCod <> "" Then
        txtProveedorCod.Text = psProveedorCod
    End If
    If psProveedorNombre <> "" Then
        txtProveedorNombre.Text = psProveedorNombre
    End If
    If psProveedorDocTpo <> "" Then
        txtProveedorDocTpo.Text = psProveedorDocTpo
    End If
    If psProveedorDocNro <> "" Then
        txtProveedorDocNro.Text = psProveedorDocNro
    End If
    If psProveedorCtaNro <> "" Then
        txtProveedorCtaNro.Text = psProveedorCtaNro
    End If
    If psProveedorCtaMoneda <> "" Then
        txtProveedorCtaMoneda.Text = psProveedorCtaMoneda
    End If
    If psProveedorCtaInstitucionCod <> "" Then
        txtProveedorCtaInstitucionCod.Text = psProveedorCtaInstitucionCod
    End If
    If psProveedorCtaInstitucionNombre <> "" Then
        txtProveedorCtaInstitucionNombre.Text = psProveedorCtaInstitucionNombre
    End If
    If psCompraDescripcion <> "" Then
        txtCompraDescripcion.Text = psCompraDescripcion
    End If
    If psCompraObservacion <> "" Then
        txtCompraObservacion.Text = psCompraObservacion
    End If
    If fnDocTpo = LogTipoActaConformidad.gActaRecepcionBienes Then
        txtCompraNGuia.Enabled = True
    Else
        txtCompraNGuia.Enabled = False
    End If
'    If fnTpoDocOrigen = LogTipoDocOrigenComprobante.CompraLibre Then
'        txtDocReferencia.Enabled = True
'        txtDocReferencia.BackColor = RGB(252, 250, 207)
'    Else
'        txtDocReferencia.BackColor = vbWhite
'        txtDocReferencia.Enabled = False
'    End If
End Sub
Private Sub Form_Load()
    fsAreaAgeCod = gsCodArea & Right(gsCodAge, 2)
    CargaCompPendientes
    EstadoControles 0
End Sub
Private Sub LimpiaControles()
    fnMovNro = 0
    fntpodocorigen = 0
    fnDocTpo = 0
    txtAreaAgeCod.Text = ""
    txtAreaAgeNombre.Text = ""
    txtSubAreaDescripcion.Text = ""
    txtMoneda.Text = ""
    txtDocReferencia.Text = ""
    txtActaNro.Text = ""
    lblTpoComprobante.Caption = ""
    lblNumeroComprobante.Caption = ""
    lblEmisionComp.Caption = ""
    txtProveedorCod.Text = ""
    txtProveedorNombre.Text = ""
    txtProveedorDocTpo.Text = ""
    txtProveedorDocNro.Text = ""
    txtProveedorCtaNro.Text = ""
    txtProveedorCtaMoneda.Text = ""
    txtProveedorCtaInstitucionCod.Text = ""
    txtProveedorCtaInstitucionNombre.Text = ""
    txtCompraDescripcion = ""
    txtCompraObservacion.Text = ""
End Sub
Private Sub EstadoControles(ByVal pnEstado As Integer)
    Select Case pnEstado
        Case 0
            feComprobante.Enabled = True
            cmdDarConformidad.Enabled = True
            cmdConforme.Enabled = False
            cmdCancelar.Enabled = False
            txtSubAreaDescripcion.BackColor = vbWhite
            txtCompraDescripcion.BackColor = vbWhite
            txtCompraNGuia.BackColor = vbWhite
            txtCompraObservacion.BackColor = vbWhite
            If gsOpeCod = gnAlmaActaConformidadLibreMN Or gsOpeCod = gnAlmaActaConformidadLibreMN Then
                txtDocReferencia.Enabled = True
                txtDocReferencia.BackColor = vbWhite
            Else
                txtDocReferencia.BackColor = vbWhite
                txtDocReferencia.Enabled = False
            End If
        Case 1
            feComprobante.Enabled = False
            cmdDarConformidad.Enabled = False
            cmdConforme.Enabled = True
            cmdCancelar.Enabled = True
            txtSubAreaDescripcion.BackColor = RGB(252, 250, 207)
            txtCompraDescripcion.BackColor = RGB(252, 250, 207)
            txtCompraNGuia.BackColor = RGB(252, 250, 207)
            txtCompraObservacion.BackColor = RGB(252, 250, 207)
            If gsOpeCod = gnAlmaActaConformidadLibreMN Or gsOpeCod = gnAlmaActaConformidadLibreMN Then
                txtDocReferencia.Enabled = True
                txtDocReferencia.BackColor = RGB(252, 250, 207)
            Else
                txtDocReferencia.BackColor = vbWhite
                txtDocReferencia.Enabled = False
            End If
    End Select
End Sub
Private Sub CargaCompPendientes()
Dim oDLog As DLogGeneral
Dim rs As ADODB.Recordset
Dim row As Integer
Set oDLog = New DLogGeneral
    Select Case gsOpeCod
        Case gnAlmaActaConformidadMN, gnAlmaActaConformidadME
            Set rs = oDLog.ListaComprobantexActaConformidad(fsAreaAgeCod, fnMoneda)
        Case gnAlmaActaConformidadLibreMN, gnAlmaActaConformidadLibreME
        Set rs = oDLog.ListaComprobanteLibrexActaConformidad(fsAreaAgeCod, fnMoneda)
    End Select
    Call LimpiaFlex(feComprobante)
    Do While Not rs.EOF
        feComprobante.AdicionaFila
        row = feComprobante.row
        feComprobante.TextMatrix(row, 1) = rs!Proveedor
        feComprobante.TextMatrix(row, 2) = rs!Tipo
        feComprobante.TextMatrix(row, 3) = rs!Numero
        feComprobante.TextMatrix(row, 4) = rs!Emision
        feComprobante.TextMatrix(row, 5) = rs!AreaUsuaria
        feComprobante.TextMatrix(row, 6) = rs!Observacion
        feComprobante.TextMatrix(row, 7) = rs!Moneda
        feComprobante.TextMatrix(row, 8) = rs!monto
        feComprobante.TextMatrix(row, 9) = rs!nMovNro
        feComprobante.TextMatrix(row, 10) = rs!TpoDocOri
        rs.MoveNext
    Loop
    Set rs = Nothing
    Set oDLog = Nothing
    Exit Sub
End Sub
Private Sub txtCompraDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCompraNGuia.Enabled = True Then
            txtCompraNGuia.SetFocus
        Else
            txtCompraObservacion.SetFocus
        End If
    End If
End Sub
Private Sub txtCompraNGuia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCompraObservacion.SetFocus
    End If
End Sub
Private Sub txtCompraObservacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdConforme.SetFocus
    End If
End Sub
Private Sub txtDocReferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCompraDescripcion.SetFocus
    End If
End Sub

Private Sub txtSubAreaDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtDocReferencia.Enabled Then
            txtDocReferencia.SetFocus
        Else
            txtCompraDescripcion.SetFocus
        End If
    End If
End Sub
