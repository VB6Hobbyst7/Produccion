VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogComprobanteLibreRegistro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Comprobante"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   Icon            =   "frmLogComprobanteLibreRegistro.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstComprobanteLibre 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Comprobante"
      TabPicture(0)   =   "frmLogComprobanteLibreRegistro.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmInformaGeneral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmComprobanteTipo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmDetalleComprobante"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAgregar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdQuitar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCancelar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdRegistrar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame Frame2 
         Caption         =   "Documento Origen"
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
         TabIndex        =   24
         Top             =   480
         Width           =   9945
         Begin VB.ComboBox cboTpoDocOrigen 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Doc. Origen:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   270
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdRegistrar 
         Caption         =   "&Registrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   8160
         TabIndex        =   22
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Ca&ncelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   9120
         TabIndex        =   23
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   1200
         TabIndex        =   21
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   240
         TabIndex        =   20
         Top             =   5520
         Width           =   975
      End
      Begin VB.Frame frmDetalleComprobante 
         Caption         =   "Detalle del Comprobante"
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
         Height          =   2835
         Left            =   240
         TabIndex        =   18
         Top             =   2640
         Width           =   9945
         Begin Sicmact.FlexEdit feOrden 
            Height          =   2535
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   9615
            _ExtentX        =   16933
            _ExtentY        =   4366
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Ag.Destino-Objeto-Descripcion-Solicitado-P.Unitario-SubTotal-CtaContCod"
            EncabezadosAnchos=   "0-1000-1400-3500-900-1100-1100-0"
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
            ColumnasAEditar =   "X-1-2-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-1-1-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-L-R-R-R-L"
            FormatosEdit    =   "0-0-0-0-3-2-2-0"
            CantEntero      =   12
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame frmComprobanteTipo 
         Caption         =   "Comprobante"
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
         Height          =   1350
         Left            =   6120
         TabIndex        =   10
         Top             =   1200
         Width           =   4050
         Begin VB.ComboBox cboTpoComprobante 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   240
            Width           =   3015
         End
         Begin VB.TextBox txtComprobanteNro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   15
            Top             =   600
            Width           =   2145
         End
         Begin VB.TextBox txtComprobanteSerie 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   14
            Top             =   600
            Width           =   795
         End
         Begin MSComCtl2.DTPicker txtComprobanteFecEmision 
            Height          =   285
            Left            =   840
            TabIndex        =   17
            Top             =   915
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   240058369
            CurrentDate     =   41586
         End
         Begin VB.Label Label5 
            Caption         =   "Emisión:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "N°:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label11 
            Caption         =   "Tipo:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame frmInformaGeneral 
         Caption         =   "Información General"
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
         Height          =   1350
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   5850
         Begin VB.TextBox txtObservacion 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   9
            Top             =   960
            Width           =   4400
         End
         Begin Sicmact.TxtBuscar txtPersona 
            Height          =   315
            Left            =   1320
            TabIndex        =   3
            Top             =   240
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            TipoBusqueda    =   3
            sTitulo         =   ""
         End
         Begin Sicmact.TxtBuscar txtArea 
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Top             =   600
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
         End
         Begin VB.Label Label1 
            Caption         =   "Proveedor:"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Área Usuaria:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Observaciones:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblProveedorNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2520
            TabIndex        =   4
            Top             =   240
            Width           =   3210
         End
         Begin VB.Label lblAreaAgeNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2520
            TabIndex        =   7
            Top             =   600
            Width           =   3210
         End
      End
   End
   Begin Sicmact.FlexEdit feObj 
      Height          =   1575
      Left            =   10440
      TabIndex        =   27
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2778
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Id-Objeto Orden-CtaContCod-CtaContDesc-Filtro-CodObjeto"
      EncabezadosAnchos=   "0-400-800-800-800-800-800"
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
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmLogComprobanteLibreRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************
'Nombre : frmLogComprobanteLibreRegistro
'Descripcion:Formulario para registro de Comprobante Libre
'Creacion: PASIERS0772014
'*****************************
Option Explicit
Dim fntpodocorigen As Integer
Dim fnMoneda As Integer
Dim fsAreaAgeCod As String
Dim fnTpoCambio As Currency
Dim fsCtaContCodProv As String
Dim fnTipoPago As Integer
Dim lsCodPers As String
Dim fRsCompra As New ADODB.Recordset
Dim fRsServicio As New ADODB.Recordset
Dim fRsAgencia As New ADODB.Recordset
Private Type TypeProveedor
    DocTpo As String
    DocNro As String
    IFICtaCod As String
    CtaMoneda As String
    IFICod As String
    CtaIFINombre As String
    CompraDesc As String
    CompraObs As String
End Type
Dim objProveedor As TypeProveedor
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Public Function Inicio(ByVal psOpeCod)
    gsopecod = psOpeCod
    fnMoneda = Mid(psOpeCod, 3, 1)
    Me.Show 1
End Function
Private Sub cboTpoComprobante_Click()
    txtComprobanteSerie.SetFocus
End Sub
Private Sub cboTpoDocOrigen_Click()
   Dim lnTpoDoc As Integer
    LimpiarDatos
    If Trim(Right(cboTpoDocOrigen.Text, 4)) <> "" Then
        lnTpoDoc = CInt(Trim(Right(cboTpoDocOrigen.Text, 4)))
    End If
    fntpodocorigen = lnTpoDoc ' + 2
    If fntpodocorigen = LogTipoDocOrigenComprobante.CompraLibre Then
        feOrden.lbUltimaInstancia = True
    ElseIf fntpodocorigen = LogTipoDocOrigenComprobante.Serviciolibre Then
        feOrden.lbUltimaInstancia = False
    End If
End Sub
Private Function validaBusquedaLibre() As Boolean
    validaBusquedaLibre = True
    If cboTpoDocOrigen.ListIndex = -1 Then
        MsgBox "Ud. primero debe de seleccionar el Tipo de Documento Origen", vbInformation, "Aviso"
        validaBusquedaLibre = False
        cboTpoDocOrigen.SetFocus
        Exit Function
    End If
End Function
Private Sub cmdAgregar_Click()
   If Not validaBusquedaLibre Then Exit Sub
    If feOrden.TextMatrix(1, 0) <> "" Then
        If Not validaIngresoRegistros Then Exit Sub
    End If
    feOrden.AdicionaFila
    
    If fntpodocorigen = LogTipoDocOrigenComprobante.CompraLibre Then
        feOrden.ColumnasAEditar = "X-1-2-3-4-5-X-X"
        feOrden.TextMatrix(feOrden.row, 4) = "0"
        feOrden.TextMatrix(feOrden.row, 5) = "0.00"
        feOrden.TextMatrix(feOrden.row, 6) = "0.00"
    ElseIf fntpodocorigen = LogTipoDocOrigenComprobante.Serviciolibre Then
        feOrden.ColumnasAEditar = "X-1-2-3-X-X-6-X"
    End If
    feOrden.TextMatrix(feOrden.row, 6) = "0.00"
    feOrden.col = 2
    feOrden.SetFocus
    feOrden_RowColChange
End Sub
Private Function validaIngresoRegistros() As Boolean
    Dim I As Long, J As Long
    Dim col As Integer
    Dim Columnas() As String
    Dim lsColumnas As String
    
    lsColumnas = "1,2,6"
    Columnas = Split(lsColumnas, ",")
        
    validaIngresoRegistros = True
    If feOrden.TextMatrix(1, 0) <> "" Then
        For I = 1 To feOrden.Rows - 1
            For J = 1 To feOrden.Cols - 1
                For col = 0 To UBound(Columnas)
                    If J = Columnas(col) Then
                        If Len(Trim(feOrden.TextMatrix(I, J))) = 0 And feOrden.ColWidth(J) <> 0 Then
                            MsgBox "Ud. debe especificar el campo " & feOrden.TextMatrix(0, J), vbInformation, "Aviso"
                            validaIngresoRegistros = False
                            feOrden.TopRow = I
                            feOrden.row = I
                            feOrden.col = J
                            feOrden_RowColChange
                            Exit Function
                        End If
                    End If
                Next
            Next
            If IsNumeric(feOrden.TextMatrix(I, 6)) Then
                If CCur(feOrden.TextMatrix(I, 6)) <= 0 Then
                    MsgBox "El Importe Total debe ser mayor a cero", vbInformation, "Aviso"
                    validaIngresoRegistros = False
                    feOrden.TopRow = I
                    feOrden.row = I
                    feOrden.col = 6
                    Exit Function
                End If
            Else
                MsgBox "El Importe Total debe ser númerico", vbInformation, "Aviso"
                validaIngresoRegistros = False
                feOrden.TopRow = I
                feOrden.row = I
                feOrden.col = 6
                Exit Function
            End If
            If fntpodocorigen = LogTipoDocOrigenActaConformidad.CompraLibre Then
                If Len(Trim(feOrden.TextMatrix(I, 7))) = 0 Then
                    MsgBox "El Objeto " & feOrden.TextMatrix(I, 3) & Chr(10) & "no tiene configurado Plantilla Contable, consulte con el Dpto. de Contabilidad", vbInformation, "Aviso"
                    feOrden.TopRow = I
                    feOrden.row = I
                    feOrden.col = 2
                    validaIngresoRegistros = False
                    Exit Function
                End If
            End If
        Next
    Else
        MsgBox "Ud. debe agregar los Detalles del Comprobante.", vbInformation, "Aviso"
        validaIngresoRegistros = False
    End If
End Function
Private Function ValidarComprobante() As Boolean
Dim oProv As New DLogProveedor
    If Len(txtPersona.Text) = 0 Then
        MsgBox "Ud. debe seleccionar el Proveedor para el presente Comprobante a Registrar.", vbInformation, "Aviso"
        txtPersona.SetFocus
        ValidarComprobante = False
        Exit Function
    End If
    If Not oProv.IsExisProveedor(txtPersona.Text) Then
        MsgBox "Ud. debe verificar que el Proveedor " & UCase(lblProveedorNombre.Caption) & Chr(10) & "se encuentre registrado en la BD de Proveedores de Logística, además tenga" & Chr(10) & "configurado una cuenta en " & UCase(fnMoneda) & " para poder continuar con el proceso.", vbInformation, "Aviso"
        txtPersona.SetFocus
        ValidarComprobante = False
        Exit Function
    End If
    If Len(txtArea.Text) = 0 Then
        MsgBox "Ud. debe seleccionar el Área Agencia para el presente Comprobante a Registrar.", vbInformation, "Aviso"
        txtArea.SetFocus
        ValidarComprobante = False
        Exit Function
    End If
    If Len(txtObservacion.Text) = 0 Then
        MsgBox "No se ha ingresado la Observación para el presente Comprobante.", vbInformation, "Aviso"
        txtObservacion.SetFocus
        ValidarComprobante = False
        Exit Function
    End If
    If cboTpoComprobante.ListIndex = -1 Then
        MsgBox "No se ha seleccionado el tipo de comprobante.", vbInformation, "Aviso"
        cboTpoComprobante.SetFocus
        ValidarComprobante = False
        Exit Function
    End If
    If Len(txtComprobanteSerie.Text) = 0 Then
        MsgBox "No se ha completado el numero de comprobante.", vbInformation, "Aviso"
        txtComprobanteSerie.SetFocus
        ValidarComprobante = False
        Exit Function
    End If
    If Len(txtComprobanteNro.Text) = 0 Then
        MsgBox "No se ha completado el numero de comprobante.", vbInformation, "Aviso"
        txtComprobanteNro.SetFocus
        ValidarComprobante = False
        Exit Function
    End If
    If feOrden.TextMatrix(1, 1) = "" Then
        MsgBox "Favor verifique del Detalle del Comprobante.", vbInformation, "Aviso"
        cmdAgregar.SetFocus
        ValidarComprobante = False
        Exit Function
    End If
    ValidarComprobante = True
End Function
Private Sub LimpiarDatos()
    txtPersona.Text = ""
    lblProveedorNombre.Caption = ""
    txtArea.Text = ""
    lblAreaAgeNombre.Caption = ""
    txtObservacion.Text = ""
    cboTpoComprobante.ListIndex = -1
    txtComprobanteSerie.Text = ""
    txtComprobanteNro.Text = ""
    txtComprobanteFecEmision.value = CDate(gdFecSis)
    objProveedor.DocTpo = ""
    objProveedor.DocNro = ""
    objProveedor.IFICtaCod = ""
    objProveedor.CtaMoneda = ""
    objProveedor.IFICod = ""
    objProveedor.CtaIFINombre = ""
    objProveedor.CompraDesc = ""
    objProveedor.CompraObs = ""
    Call LimpiaFlex(feOrden)
    Call FormateaFlex(feObj) 'PASI20150130
    txtPersona.SetFocus
End Sub
Private Sub cmdCancelar_Click()
    LimpiarDatos
End Sub
Private Sub cmdQuitar_Click()
    feOrden.EliminaFila feOrden.row
End Sub

Private Sub cmdRegistrar_Click()
    Dim olog As NLogGeneral
    Dim oDLog As DLogGeneral
    Dim lnMovNro As Long
    Dim lsMovNro As String
    Dim DatosOrden() As TComprobanteOrden
    Dim Index As Integer, indexObj As Integer
    Dim lsSubCta As String
    Dim lsValidaMovReg As String 'vapa20161110
    
    On Error GoTo ErrCmdRegistrar
    If Not ValidarComprobante Then Exit Sub
    
    ReDim DatosOrden(Index)
    For Index = 1 To feOrden.Rows - 1
        ReDim Preserve DatosOrden(Index)
        If fntpodocorigen = LogTipoDocOrigenComprobante.CompraLibre Then
            DatosOrden(Index).sCtaContCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 7))))
        ElseIf fntpodocorigen = LogTipoDocOrigenComprobante.Serviciolibre Then
            lsSubCta = ""
            For indexObj = 1 To feObj.Rows - 1
                If feObj.TextMatrix(indexObj, 1) = feOrden.TextMatrix(Index, 0) Then
                    lsSubCta = lsSubCta & feObj.TextMatrix(indexObj, 5)
                End If
            Next
            DatosOrden(Index).sCtaContCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 2)))) & lsSubCta
        End If
        DatosOrden(Index).sObjeto = CStr(Trim(feOrden.TextMatrix(Index, 2)))
        DatosOrden(Index).sDescripcion = CStr(Trim(feOrden.TextMatrix(Index, 3)))
        DatosOrden(Index).nCantidad = Val(feOrden.TextMatrix(Index, 4))
        DatosOrden(Index).nTotal = feOrden.TextMatrix(Index, 6)
    Next
    If UBound(DatosOrden) = 0 Then
        MsgBox "No existen Items a dar conformidad", vbCritical, "Aviso"
        Exit Sub
    End If
    If fntpodocorigen = LogTipoDocOrigenActaConformidad.CompraLibre Then
        fsCtaContCodProv = "25" & Mid(gsopecod, 3, 1) & "601"
    ElseIf fntpodocorigen = LogTipoDocOrigenActaConformidad.Serviciolibre Then
        fsCtaContCodProv = "25" & Mid(gsopecod, 3, 1) & "60202"
    End If
    If fsCtaContCodProv = "" Then
        MsgBox "No se ha definido cuenta contable de Proveedor, consulte al Dpto. de TI", vbInformation, "Aviso"
        Exit Sub
    End If
    
'     If MsgBox("¿Esta seguro de guardar el Comprobante Libre?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
'        Exit Sub
'    End If
    Set olog = New NLogGeneral
    'Screen.MousePointer = 11
    lsValidaMovReg = olog.ValidaComprobanteReg(Trim(Right(cboTpoComprobante.Text, 3)), txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, txtPersona.Text) ' CDate(txtComprobanteFecEmision.value)) 'vapa
    If lsValidaMovReg = "no" Then
    
        If MsgBox("¿Esta seguro de guardar el Comprobante Libre?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
        End If
    
    
    Screen.MousePointer = 11
    lnMovNro = olog.GrabarComprobanteOrden(gdFecSis, Right(gsCodAge, 2), gsCodUser, gsopecod, "Registro del Comprobante Nº " & txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, txtPersona.Text, txtArea.Text, fnMoneda, fntpodocorigen, txtObservacion.Text, Trim(Right(cboTpoComprobante.Text, 3)), txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, CDate(txtComprobanteFecEmision.value), fnTipoPago, objProveedor.IFICod, objProveedor.IFICtaCod, DatosOrden, fsCtaContCodProv, fnTpoCambio, lsMovNro)
    Else
         MsgBox "El número de comprobante ya existe, por favor ingrese otro número ", vbInformation, "Aviso"
        Exit Sub
    End If
    Screen.MousePointer = 0
    If lnMovNro = 0 Then
        MsgBox "Ha ocurrido un error al registrar el Comprobante", vbCritical, "Aviso"
        Set olog = Nothing
        Exit Sub
    End If
    Set oDLog = New DLogGeneral
    oDLog.RegistraActaPendxComprobante lnMovNro, txtArea.Text
    MsgBox "Se ha registrado el Comprobante de Nro. " & txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text & " con éxito", vbInformation, "Aviso"
    Set olog = Nothing
    'ARLO 20160126 ***
    Dim lsMoneda As String
    If (fnMoneda = 1) Then
    lsMoneda = "SOLES"
    gsopecod = LogPistaRegistroComprobanteMN
    Else
    lsMoneda = "DOLARES"
    gsopecod = LogPistaRegistroComprobanteME
    End If
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, "Registro del Comprobante Libre Nº " & txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text & " En Moneda " & lsMoneda
    Set objPista = Nothing
    '***
    If MsgBox("¿Desea registrar otro Comprobante?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
        cmdCancelar_Click
    Else
        Unload Me
    End If
    Exit Sub
ErrCmdRegistrar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub feOrden_OnCellChange(pnRow As Long, pnCol As Long)
    On Error GoTo ErrfeOrden_OnCellChange
    If feOrden.TextMatrix(1, 0) <> "" Then
        If fntpodocorigen = LogTipoDocOrigenComprobante.CompraLibre Then
            If pnCol = 4 Or pnCol = 5 Then
                feOrden.TextMatrix(pnRow, 6) = Format(Val(feOrden.TextMatrix(pnRow, 4)) * feOrden.TextMatrix(pnRow, 5), gsFormatoNumeroView)
            End If
        End If
    End If
    Exit Sub
ErrfeOrden_OnCellChange:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub feOrden_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    If psDataCod <> "" Then
        If pnCol = 2 Then
            If fntpodocorigen = LogTipoDocOrigenComprobante.Serviciolibre Then
                AsignaObjetosSer psDataCod
            End If
        End If
        If pnCol = 1 Or pnCol = 2 Then
            '*** Si esta vacio el campo de la cuenta contable y si ya eligió agencia y objeto
            If Len(Trim(feOrden.TextMatrix(pnRow, 1))) <> 0 And Len(Trim(feOrden.TextMatrix(pnRow, 2))) <> 0 Then
                feOrden.TextMatrix(pnRow, 7) = DameCtaCont(feOrden.TextMatrix(pnRow, 2), 0, Trim(feOrden.TextMatrix(pnRow, 1)))
            End If
            '***
        End If
    End If
End Sub
Private Function DameCtaCont(ByVal psObjeto As String, nNiv As Integer, psAgeCod As String) As String
    Dim oCon As New DConecta
    Dim oForm As New frmLogOCompra
    Dim rs As New ADODB.Recordset
    Dim sSql As String
    
    sSql = oForm.FormaSelect(gsopecod, psObjeto, 0, psAgeCod)
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    If Not rs.EOF Then
        DameCtaCont = rs!cObjetoCod
    End If
    Set rs = Nothing
    Set oForm = Nothing
    Set oCon = Nothing
End Function
Private Sub feOrden_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(feOrden.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
Private Sub feOrden_RowColChange()
    If feOrden.col = 1 Then
        feOrden.rsTextBuscar = fRsAgencia
    ElseIf feOrden.col = 2 Then
        If fntpodocorigen = LogTipoDocOrigenComprobante.CompraLibre Then
            feOrden.rsTextBuscar = fRsCompra
        ElseIf fntpodocorigen = LogTipoDocOrigenComprobante.Serviciolibre Then
            feOrden.rsTextBuscar = fRsServicio
        End If
    End If
End Sub
Private Sub Form_Load()
    fsAreaAgeCod = gsCodArea & Right(gsCodArea, 2)
    '''Me.Caption = "Registro de Comprobantes Libres en " & IIf(Mid(gsOpeCod, 3, 1) = 1, "SOLES", "DOLARES") 'marg ers044-2016
    Me.Caption = "Registro de Comprobantes Libres en " & IIf(Mid(gsopecod, 3, 1) = 1, StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES") 'marg ers044-2016
    LLenaControles
    LLenaVariables
End Sub
Private Sub LLenaControles()
    Dim oArea As DActualizaDatosArea
    Set oArea = New DActualizaDatosArea
    Dim oDoc As New DOperacion
    Dim rs As New ADODB.Recordset
    Dim rsOrigen As ADODB.Recordset
    Dim olog As New DLogGeneral
     
    Set rsOrigen = olog.ListaTpoDocOrigenComprobanteLibre
    cboTpoDocOrigen.Clear
    CargaCombo rsOrigen, cboTpoDocOrigen, , 1, 0
    
    txtArea.rs = oArea.GetAgenciasAreas
    Set rs = oDoc.CargaOpeDoc(gnAlmaComprobanteLibreRegistroMN, OpeDocMetDigitado)
    cboTpoComprobante.Clear
    Do While Not rs.EOF
        cboTpoComprobante.AddItem Format(rs!nDocTpo, "00") & " " & Mid(rs!cDocDesc & Space(100), 1, 100) & rs!nDocTpo
        rs.MoveNext
    Loop
    Set rs = Nothing
    Set oDoc = Nothing
End Sub
Private Sub LLenaVariables()
    Dim oDoc As New DOperacion
    Dim oArea As New DActualizaDatosArea
    Dim oAlmacen As New DLogAlmacen
    Dim rs As New ADODB.Recordset
    
    fntpodocorigen = LogTipoDocOrigenComprobante.CompraLibre
    If gbBitTCPonderado Then
        fnTpoCambio = gnTipCambioPonderado
    Else
        fnTpoCambio = gnTipCambioC
    End If
    Set rs = oDoc.CargaOpeCta(gsopecod, "H")
    fsCtaContCodProv = rs!cCtaContCod
    Set fRsAgencia = oArea.GetAgencias(, , True)
    Set fRsCompra = oAlmacen.GetBienesAlmacen(, "11','12','13")
    Set fRsServicio = OrdenServicio()

    Set rs = Nothing
    Set oArea = Nothing
    Set oAlmacen = Nothing
    Set oDoc = Nothing
End Sub
Private Function OrdenServicio() As ADODB.Recordset
    Dim oCon As New DConecta
    Dim sSqlO As String
    Dim lnMoneda As Integer
    If fnMoneda <> 0 Then
        oCon.AbreConexion
        sSqlO = "SELECT DISTINCT a.cCtaContCod as cObjetoCod, b.cCtaContDesc, 2 as nObjetoNiv " _
              & "FROM  " & gcCentralCom & "OpeCta a,  " & gcCentralCom & "CtaCont b " _
              & "WHERE b.cCtaContCod = a.cCtaContCod AND (a.cOpeCod='" & IIf(fnMoneda = 1, "501207", "502207") & "' AND (a.cOpeCtaDH='D'))"
        Set OrdenServicio = oCon.CargaRecordSet(sSqlO)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function
Private Sub TxtArea_EmiteDatos()
    Me.lblAreaAgeNombre.Caption = txtArea.psDescripcion
    If lblAreaAgeNombre.Caption <> "" Then
        Me.txtObservacion.SetFocus
    End If
End Sub

Private Sub cboTpoComprobante_LostFocus()
    If Trim(Left(cboTpoComprobante, 2)) = "05" Then
        txtComprobanteSerie = Trim(Str("3"))
    End If
End Sub '***NAGL ERS012-2017 20170710

Private Sub txtComprobanteNro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAgregar.SetFocus
    End If
End Sub
Private Sub txtComprobanteSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = LetrasNumeros(KeyAscii)
    If KeyAscii = 13 Then
        If Trim(Left(cboTpoComprobante, 2)) = "05" Then
           txtComprobanteSerie = Trim(Str("3"))
        Else
            txtComprobanteSerie = Right(String(4, "0") & txtComprobanteSerie, 4)
        End If 'NAGL ERS012-2017 20170710
        txtComprobanteNro.SetFocus
    End If
End Sub
Private Sub txtComprobanteSerie_LostFocus()
    If Trim(Left(cboTpoComprobante, 2)) = "05" Then
        txtComprobanteSerie = Trim(Str("3"))
    Else
        txtComprobanteSerie = Right(String(4, "0") & txtComprobanteSerie, 4)
    End If
End Sub '***NAGL ERS012-2017 20170710
Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboTpoComprobante.SetFocus
    End If
End Sub
Private Sub txtPersona_EmiteDatos()
    Dim olog As New DLogGeneral
    Dim oProv As New DLogProveedor
    Dim rs As New ADODB.Recordset
    Dim bExiste As Boolean
    
    lblProveedorNombre.Caption = ""
    bExiste = oProv.IsExisProveedor(txtPersona.Text)
    If Not bExiste Then
        MsgBox "El proveedor seleccionado no se encuentra registrado en la BD de proveedores" & Chr(10) & "del Departamento de Logistica, esto sera necesario para el Pago.", vbInformation, "Aviso"
    End If
    lblProveedorNombre.Caption = txtPersona.psDescripcion
    Set rs = olog.ObtenerProveedorxRegistroComprobanteLibre(txtPersona.Text, fnMoneda)
    If Not RSVacio(rs) Then
        EstablecerDatosRegistroComprobante rs!cDocTpo, rs!cDocNro, rs!cIFiCtaCod, IIf(rs!cIFiCtaCod = "", "", rs!cMoneda), rs!cIFiCod, rs!cIFiNombre
    End If
    Set rs = Nothing
    Set oProv = Nothing
    Set olog = Nothing
    Exit Sub
ErrorPersona:
    MsgBox Err.Description, vbInformation, "Error"
End Sub
Private Sub EstablecerDatosRegistroComprobante(Optional ByVal psProveedorDocTpo As String = "", _
                                                Optional ByVal psProveedorDocNro As String = "", _
                                                Optional ByVal psProveedorIFICtaCod As String = "", _
                                                Optional ByVal psProveedorCtaMoneda As String = "", _
                                                Optional ByVal psProveedorCtaIFICod As String = "", _
                                                Optional ByVal psProveedorCtaIFINombre As String = "", _
                                                Optional ByVal psCompraDescripcion As String = "", _
                                                Optional ByVal psCompraObservacion As String = "")

If psProveedorDocTpo <> "" Then
    objProveedor.DocTpo = psProveedorDocTpo
End If
If psProveedorDocNro <> "" Then
    objProveedor.DocNro = psProveedorDocNro
End If
If psProveedorIFICtaCod <> "" Then
    objProveedor.IFICtaCod = psProveedorIFICtaCod
End If
If psProveedorCtaMoneda <> "" Then
    objProveedor.CtaMoneda = psProveedorCtaMoneda
End If
If psProveedorCtaIFICod <> "" Then
    objProveedor.IFICod = psProveedorCtaIFICod
End If
If psProveedorCtaIFINombre <> "" Then
   objProveedor.CtaIFINombre = psProveedorCtaIFINombre
End If
If psCompraDescripcion <> "" Then
    objProveedor.CompraDesc = psCompraDescripcion
End If
If psCompraObservacion <> "" Then
    objProveedor.CompraObs = psCompraObservacion
End If

    If objProveedor.IFICod = "1090100012521" Then 'CMACMAYNAS
        fnTipoPago = LogTipoPagoComprobante.gPagoCuentaCMAC
    Else
        If objProveedor.IFICod = "1090100824640" Then 'BCP
            fnTipoPago = LogTipoPagoComprobante.gPagoTransferencia
        Else 'OTRO BANCO
            fnTipoPago = LogTipoPagoComprobante.gPagoCheque
        End If
    End If
End Sub
'PASI20150130
Private Sub AsignaObjetosSer(ByVal sCtaCod As String)
    Dim nNiv As Integer
    Dim nObj As Integer
    Dim nObjs As Integer
    Dim oCon As New DConecta
    Dim oCtaCont As New DCtaCont
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim oRHAreas As New DActualizaDatosArea
    Dim oCtaIf As New NCajaCtaIF
    Dim oEfect As New Defectivo
    Dim oDescObj As New ClassDescObjeto
    Dim oContFunct As New NContFunciones
    Dim lsRaiz As String, lsFiltro As String, sSql As String
        
    oDescObj.lbUltNivel = True
    oCon.AbreConexion
    EliminaObjeto feOrden.row

    sSql = "SELECT MAX(nCtaObjOrden) as nNiveles FROM CtaObj WHERE cCtaContCod = '" & sCtaCod & "' and cObjetoCod <> '00' "
    Set rs = oCon.CargaRecordSet(sSql)
    nObjs = IIf(IsNull(rs!nNiveles), 0, rs!nNiveles)
      
    Set rs1 = oCtaCont.CargaCtaObj(sCtaCod, , True)
    If Not rs1.EOF And Not rs1.BOF Then
        Do While Not rs1.EOF
            lsRaiz = ""
            lsFiltro = ""
            Set rs = New ADODB.Recordset
            Select Case Val(rs1!cObjetoCod)
                Case ObjCMACAgencias
                    Set rs = oRHAreas.GetAgencias()
                Case ObjCMACAgenciaArea
                    lsRaiz = "Unidades Organizacionales"
                    Set rs = oRHAreas.GetAgenciasAreas()
                Case ObjCMACArea
                    Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
                Case ObjEntidadesFinancieras
                    lsRaiz = "Cuentas de Entidades Financieras"
                    Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, sCtaCod)
                Case ObjDescomEfectivo
                    Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
                Case ObjPersona
                    Set rs = Nothing
                Case Else
                    lsRaiz = "Varios"
                    Set rs = GetObjetos(rs1!cObjetoCod)
            End Select
            If Not rs Is Nothing Then
                If rs.State = adStateOpen Then
                    If Not rs.EOF And Not rs.BOF Then
                        If rs.RecordCount > 1 Then
                            oDescObj.Show rs, "", lsRaiz
                            If oDescObj.lbOK Then
                                lsFiltro = oContFunct.GetFiltroObjetos(Trim(rs1!cObjetoCod), sCtaCod, oDescObj.gsSelecCod, False)
                                AdicionaObjeto feOrden.TextMatrix(feOrden.row, 0), IIf(IsNull(rs1!nCtaObjOrden), "", rs1!nCtaObjOrden), oDescObj.gsSelecCod, oDescObj.gsSelecDesc, lsFiltro, IIf(IsNull(rs1!cObjetoCod), "", rs1!cObjetoCod)
                            Else
                                EliminaObjeto feOrden.row
                                Exit Do
                            End If
                        Else
                            AdicionaObjeto feOrden.TextMatrix(feOrden.row, 0), IIf(IsNull(rs1!nCtaObjOrden), "", rs1!nCtaObjOrden), oDescObj.gsSelecCod, oDescObj.gsSelecDesc, lsFiltro, IIf(IsNull(rs1!cObjetoCod), "", rs1!cObjetoCod)
                        End If
                    End If
                End If
            End If
            rs1.MoveNext
        Loop
    End If

    Set rs = Nothing
    Set rs1 = Nothing
    Set oDescObj = Nothing
    Set oCon = Nothing
    Set oCtaCont = Nothing
    Set oCtaIf = Nothing
    Set oEfect = Nothing
    Set oContFunct = Nothing
    Set oContFunct = Nothing
    Exit Sub
End Sub
Private Sub AdicionaObjeto(ByVal pnItem As Integer, ByVal psCtaObjOrden As String, ByVal psCodigo As String, ByVal psDesc As String, ByVal psFiltro As String, ByVal psObjetoCod As String)
    feObj.AdicionaFila
    feObj.TextMatrix(feObj.row, 1) = pnItem
    feObj.TextMatrix(feObj.row, 2) = psCtaObjOrden
    feObj.TextMatrix(feObj.row, 3) = psCodigo
    feObj.TextMatrix(feObj.row, 4) = psDesc
    feObj.TextMatrix(feObj.row, 5) = psFiltro
    feObj.TextMatrix(feObj.row, 6) = psObjetoCod
End Sub
Private Sub EliminaObjeto(ByVal pnItem As Integer)
    Dim I As Long
    Dim bEncuentra As Boolean
    If feObj.TextMatrix(1, 0) <> "" Then
        For I = 1 To feObj.Rows - 1
            If Val(feObj.TextMatrix(I, 1)) = pnItem Then
                bEncuentra = True
                Exit For
            End If
        Next
    End If
    If bEncuentra Then
        feObj.EliminaFila I
        EliminaObjeto pnItem
    End If
End Sub
'END PASI



