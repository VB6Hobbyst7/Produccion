VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRiesgoConfgCodigosSobEnd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Sobreendeudados"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8595
   Icon            =   "frmCredRiesgoConfgCodigosSobEnd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGlosa 
      Height          =   480
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   4720
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      ToolTipText     =   "Grabar"
      Top             =   4755
      Width           =   1000
   End
   Begin VB.Frame fraTipoConfiguracion 
      Caption         =   "Tipo de configuración"
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
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8415
      Begin VB.ComboBox cmbTpoConfiguracion 
         Height          =   315
         ItemData        =   "frmCredRiesgoConfgCodigosSobEnd.frx":030A
         Left            =   1320
         List            =   "frmCredRiesgoConfgCodigosSobEnd.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   310
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Configuración:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      ToolTipText     =   "Cancelar"
      Top             =   4755
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      ToolTipText     =   "Grabar"
      Top             =   4755
      Width           =   1000
   End
   Begin TabDlg.SSTab TabSegConfig 
      Height          =   3720
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   6562
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Capacidad de Pago"
      TabPicture(0)   =   "frmCredRiesgoConfgCodigosSobEnd.frx":03DD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTramoCP"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Nro Entidades"
      TabPicture(1)   =   "frmCredRiesgoConfgCodigosSobEnd.frx":03F9
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraEntidad"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Matriz de configuración"
      TabPicture(2)   =   "frmCredRiesgoConfgCodigosSobEnd.frx":0415
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "feMatriz"
      Tab(2).ControlCount=   2
      Begin VB.Frame fraTramoCP 
         Caption         =   "Tramos Capacidad de Pago"
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
         Height          =   3135
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   8175
         Begin VB.CommandButton cmdTramoCPAceptar 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   350
            Left            =   2880
            TabIndex        =   16
            ToolTipText     =   "Aceptar"
            Top             =   2700
            Width           =   900
         End
         Begin VB.CommandButton cmdTramoCPCancelar 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   350
            Left            =   3840
            TabIndex        =   15
            ToolTipText     =   "Cancelar"
            Top             =   2700
            Width           =   900
         End
         Begin VB.CommandButton cmdTramoCPEliminar 
            Caption         =   "Elimina&r"
            Height          =   350
            Left            =   7155
            TabIndex        =   14
            ToolTipText     =   "Eliminar Tramo CP"
            Top             =   960
            Width           =   900
         End
         Begin VB.CommandButton cmdTramoCPNuevo 
            Caption         =   "&Nuevo"
            Height          =   350
            Left            =   7155
            TabIndex        =   13
            ToolTipText     =   "Nuevo Tramo CP"
            Top             =   230
            Width           =   900
         End
         Begin VB.CommandButton cmdTramoCPEditar 
            Caption         =   "&Editar"
            Height          =   350
            Left            =   7155
            TabIndex        =   12
            ToolTipText     =   "Editar Tramo CP"
            Top             =   600
            Width           =   900
         End
         Begin SICMACT.FlexEdit feTramoCP 
            Height          =   2415
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   6960
            _ExtentX        =   12277
            _ExtentY        =   4260
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   1
            VisiblePopMenu  =   -1  'True
            EncabezadosNombres=   "#-CP Inicial-CP Final-Descripción-Aux"
            EncabezadosAnchos=   "400-1250-1250-3300-0"
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
            ColumnasAEditar =   "X-1-2-3-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-R-L-C"
            FormatosEdit    =   "0-2-2-0-0"
            TextArray0      =   "#"
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
         End
      End
      Begin VB.Frame fraEntidad 
         Caption         =   "Nro de Entidades"
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
         Height          =   3135
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   8175
         Begin VB.CommandButton cmdEntidadEliminar 
            Caption         =   "Elimina&r"
            Height          =   350
            Left            =   7155
            TabIndex        =   22
            ToolTipText     =   "Eliminar Tramo CP"
            Top             =   960
            Width           =   900
         End
         Begin VB.CommandButton cmdEntidadEditar 
            Caption         =   "&Editar"
            Height          =   350
            Left            =   7155
            TabIndex        =   21
            ToolTipText     =   "Editar Tramo CP"
            Top             =   600
            Width           =   900
         End
         Begin VB.CommandButton cmdEntidadNuevo 
            Caption         =   "&Nuevo"
            Height          =   350
            Left            =   7155
            TabIndex        =   20
            ToolTipText     =   "Nuevo Tramo CP"
            Top             =   230
            Width           =   900
         End
         Begin VB.CommandButton cmdEntidadAceptar 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   350
            Left            =   2880
            TabIndex        =   9
            ToolTipText     =   "Aceptar"
            Top             =   2700
            Width           =   900
         End
         Begin VB.CommandButton cmdEntidadCancelar 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   350
            Left            =   3840
            TabIndex        =   8
            ToolTipText     =   "Cancelar"
            Top             =   2700
            Width           =   900
         End
         Begin SICMACT.FlexEdit feEntidad 
            Height          =   2415
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   6960
            _ExtentX        =   12277
            _ExtentY        =   4260
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   1
            VisiblePopMenu  =   -1  'True
            EncabezadosNombres=   "#-Nro Ent. Inicio-Nro Ent. Fin-Descripción-Aux"
            EncabezadosAnchos=   "400-1250-1250-3300-0"
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
            ColumnasAEditar =   "X-1-2-3-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-R-L-C"
            FormatosEdit    =   "0-3-3-0-0"
            TextArray0      =   "#"
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
         End
      End
      Begin SICMACT.FlexEdit feMatriz 
         Height          =   2880
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   5080
         FixedCols       =   0
         HighLight       =   1
         AllowUserResizing=   1
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
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   ""
         FormatosEdit    =   ""
         lbFlexDuplicados=   0   'False
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   1200
         RowHeight0      =   300
      End
      Begin VB.Label Label1 
         Caption         =   "Para cambiar el RSE dar doble clic sobre la celda"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   3405
         Width           =   3615
      End
   End
   Begin VB.Label lblGlosa 
      Caption         =   "Glosa :"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Width           =   495
   End
End
Attribute VB_Name = "frmCredRiesgoConfgCodigosSobEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'** Nombre : frmCredRiesgoConfigSobEnd
'** Descripción : Configuración de Seguimiento y Adminisión de RSE segun TI-ERS003-2020
'** Creación : EJVG, 20210119 12:00:00 PM
'****************************************************************************************
Option Explicit

Private Enum eAccionRiesgoConfig
    Nuevo = 1
    Editar = 2
    Eliminar = 3
End Enum
Private Enum eTipoConfiguracion
    Seguimiento = 1
    Admision = 2
End Enum
Private Type tCPConfig
    Index As Integer
    Inicio As Double
    Fin As Double
    Descripcion As String
    TipoRegistro As eAccionRiesgoConfig
End Type
Private Type tEntidadConfig
    Index As Integer
    Inicio As Integer
    Fin As Integer
    Descripcion As String
    TipoRegistro As eAccionRiesgoConfig
End Type
Private Type tMatrizConfig
    Index As Integer
    IndexCP As Integer
    IndexNroEntidad As Integer
    RSE As Integer
    TipoRegistro As eAccionRiesgoConfig
End Type

Dim fnTramoCPAccion As eAccionRiesgoConfig
Dim fnTramoCPNoMoverFila As Integer
Dim fnEntidadAccion As eAccionRiesgoConfig
Dim fnEntidadNoMoverFila As Integer
Dim fnTipoConfiguracion As eTipoConfiguracion
Dim TramosCP() As tCPConfig
Dim NroEntidades() As tEntidadConfig
Dim MatrizConfiguracion() As tMatrizConfig
Dim fbRealizaCambios As Boolean

Private Sub listarTramosCP()
    Dim oDCredito As COMNCredito.NCOMCredito
    Set oDCredito = New COMNCredito.NCOMCredito
    
    Dim rs As New ADODB.Recordset
    Set rs = oDCredito.ListaTramosCapacidadPagoSobreEnd(fnTipoConfiguracion)
    ReDim TramosCP(0)
    Do While Not rs.EOF
        ReDim Preserve TramosCP(rs.Bookmark)
        TramosCP(rs.Bookmark).Index = rs!nOrdenTramoCP
        TramosCP(rs.Bookmark).Inicio = rs!nTramoCPInicial
        TramosCP(rs.Bookmark).Fin = rs!nTramoCPFinal
        TramosCP(rs.Bookmark).Descripcion = rs!cDescripcionTramoCP
        rs.MoveNext
    Loop
    
    SetFlexTramosCP
End Sub

Private Sub listarNroEntidades()
    Dim oDCredito As COMNCredito.NCOMCredito
    Set oDCredito = New COMNCredito.NCOMCredito
    
    Dim rs As New ADODB.Recordset
    Set rs = oDCredito.ListaNroEntidadesPagoSobreEnd(fnTipoConfiguracion)
    ReDim NroEntidades(0)
    Do While Not rs.EOF
        ReDim Preserve NroEntidades(rs.Bookmark)
        NroEntidades(rs.Bookmark).Index = rs!nOrdenNroEntidad
        NroEntidades(rs.Bookmark).Inicio = rs!nNroEntidadIni
        NroEntidades(rs.Bookmark).Fin = rs!nNroEntidadFin
        NroEntidades(rs.Bookmark).Descripcion = rs!cDescripcionNroEntidad
        rs.MoveNext
    Loop
    
    SetFlexNroEntidades
End Sub

Private Sub listarMatrizConfiguracion()
    Dim oDCredito As COMNCredito.NCOMCredito
    Set oDCredito = New COMNCredito.NCOMCredito
    
    Dim rs As New ADODB.Recordset
    Set rs = oDCredito.ListaMatrizConfiguracionPagoSobreEnd(fnTipoConfiguracion)
    ReDim MatrizConfiguracion(0)
    Do While Not rs.EOF
        ReDim Preserve MatrizConfiguracion(rs.Bookmark)
        MatrizConfiguracion(rs.Bookmark).IndexCP = rs!nOrdenTramoCP
        MatrizConfiguracion(rs.Bookmark).IndexNroEntidad = rs!nOrdenNroEntidad
        MatrizConfiguracion(rs.Bookmark).RSE = rs!nRSE
        rs.MoveNext
    Loop
    
    SetFlexMatrizConfiguracion
End Sub

Private Sub cmbTpoConfiguracion_Click()
    If cmbTpoConfiguracion.ListIndex > -1 Then
        fnTipoConfiguracion = CInt(Right(cmbTpoConfiguracion.Text, 3))
        TabSegConfig.Tab = 2
        
        fbRealizaCambios = False
        Call cmdCancelar_Click
    End If
End Sub

Private Sub cmdCancelar_Click()
    Dim lsTipoConfig As String
    
    If fbRealizaCambios Then
        lsTipoConfig = Trim(Mid(cmbTpoConfiguracion.Text, 1, Len(cmbTpoConfiguracion.Text) - 3))
        If MsgBox("Ud. ha realizado cambios en el formulario para " & lsTipoConfig + "." & Chr(13) & "¿Está seguro de cancelar todos los cambios?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
    End If
    
    fbRealizaCambios = False
    HabilitarBotonera
        
    Call listarTramosCP
    Call listarNroEntidades
    Call listarMatrizConfiguracion
End Sub

Private Function validarGrabar() As Boolean
    Dim lsGlosa As String
    Dim i As Integer, j As Integer
    Dim objCPAnterior As tCPConfig, objCPActual As tCPConfig
    Dim totalCP As Integer, totalEntidades As Integer
    Dim objNroEntidadAnterior As tEntidadConfig, objNroEntidadActual As tEntidadConfig
    
    lsGlosa = Trim(txtGlosa.Text)
    
    If (Len(lsGlosa) <= 9) Then
        MsgBox "Ud. debe ingresar una glosa con un texto mayor a 9 caracteres", vbExclamation, "Aviso"
        txtGlosa.SetFocus
        Exit Function
    End If
    
    totalCP = UBound(TramosCP)
    For i = 1 To totalCP
        objCPActual = TramosCP(i)
        
        If (i > 1) Then
            If (objCPAnterior.Fin + 0.01 <> objCPActual.Inicio) Then
                TabSegConfig.Tab = 0
                EnfocaControl feTramoCP
                feTramoCP.row = i
                feTramoCP.col = 1
                MsgBox "Registro #" & i & ": La CP Final anterior [" & Format(objCPAnterior.Fin, "#0.00") & "] y la CP Inicial [" & Format(objCPActual.Inicio, "#0.00") & "] no son consecutivos, por favor verifique.", vbExclamation, "Aviso"
                Exit Function
            End If
        End If
        
        If (i = totalCP) Then
            If (objCPActual.Fin <> 999.99) Then
                TabSegConfig.Tab = 0
                EnfocaControl feTramoCP
                feTramoCP.row = i
                feTramoCP.col = 1
                MsgBox "La CP Final del último registro debe tener valor [999.00], por favor verifique.", vbExclamation, "Aviso"
                Exit Function
            End If
        End If
        
        objCPAnterior = TramosCP(i)
    Next i
    
    totalEntidades = UBound(NroEntidades)
    For i = 1 To totalEntidades
        objNroEntidadActual = NroEntidades(i)
        
        If (i > 1) Then
            If (objNroEntidadAnterior.Fin + 1 <> objNroEntidadActual.Inicio) Then
                TabSegConfig.Tab = 1
                EnfocaControl feEntidad
                feEntidad.row = i
                feEntidad.col = 1
                MsgBox "Registro #" & i & ": El Nro. de Entidad Final anterior [" & objNroEntidadAnterior.Fin & "] y el Nro. de Entidad Inicial [" & objNroEntidadActual.Inicio & "] no son consecutivos, por favor verifique.", vbExclamation, "Aviso"
                Exit Function
            End If
        End If
        
        If (i = totalEntidades) Then
            If (objNroEntidadActual.Fin <> 999) Then
                TabSegConfig.Tab = 1
                EnfocaControl feEntidad
                feEntidad.row = i
                feEntidad.col = 1
                MsgBox "La Nro. de Entidad Final del último registro debe tener valor [999], por favor verifique.", vbExclamation, "Aviso"
                Exit Function
            End If
        End If
        
        objNroEntidadAnterior = NroEntidades(i)
    Next i
    
    validarGrabar = True
End Function

Private Sub cmdGrabar_Click()
    Dim lsTipoConfig As String
    Dim obj As COMNCredito.NCOMCredito
    Dim MatTramosCP As Variant
    Dim MatNroEntidad As Variant
    Dim MatConfig As Variant
    Dim i As Integer
    Dim exito As Boolean
    Dim lsGlosa As String
    
    If Not validarGrabar Then Exit Sub
    
    lsTipoConfig = Trim(Mid(cmbTpoConfiguracion.Text, 1, Len(cmbTpoConfiguracion.Text) - 3))
    If MsgBox("¿Está seguro de grabar los cambios para " & lsTipoConfig & "?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrGrabar
    
    ReDim MatTramosCP(1 To 4, 0 To 0)
    For i = 1 To UBound(TramosCP)
        ReDim Preserve MatTramosCP(1 To 4, 0 To i)
        MatTramosCP(1, i) = TramosCP(i).Index
        MatTramosCP(2, i) = TramosCP(i).Inicio
        MatTramosCP(3, i) = TramosCP(i).Fin
        MatTramosCP(4, i) = TramosCP(i).Descripcion
    Next
    
    ReDim MatNroEntidad(1 To 4, 0 To 0)
    For i = 1 To UBound(NroEntidades)
        ReDim Preserve MatNroEntidad(1 To 4, 0 To i)
        MatNroEntidad(1, i) = NroEntidades(i).Index
        MatNroEntidad(2, i) = NroEntidades(i).Inicio
        MatNroEntidad(3, i) = NroEntidades(i).Fin
        MatNroEntidad(4, i) = NroEntidades(i).Descripcion
    Next
    
    ReDim MatConfig(1 To 4, 0 To 0)
    For i = 1 To UBound(MatrizConfiguracion)
        ReDim Preserve MatConfig(1 To 4, 0 To i)
        MatConfig(1, i) = MatrizConfiguracion(i).Index
        MatConfig(2, i) = MatrizConfiguracion(i).IndexCP
        MatConfig(3, i) = MatrizConfiguracion(i).IndexNroEntidad
        MatConfig(4, i) = MatrizConfiguracion(i).RSE
    Next
    
    Set obj = New COMNCredito.NCOMCredito
    lsGlosa = Trim(txtGlosa.Text)
    
    exito = obj.GrabarConfiguracionSobreEnd(fnTipoConfiguracion, MatTramosCP, MatNroEntidad, MatConfig, gsCodUser, lsGlosa)
    
    If Not exito Then
        MsgBox "Ha sucedido un error al grabar la configuración de Sobreendeudado, volver a intentar.", vbCritical, "Aviso"
        Exit Sub
    End If
    
    txtGlosa.Text = ""
    MsgBox "Se ha grabado con éxito la configuración de Sobreendeudado", vbInformation, "Aviso"
    fbRealizaCambios = False
    cmdCancelar_Click
    Exit Sub
ErrGrabar:
    Screen.MousePointer = 0
    cmdGrabar.Enabled = True
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub cmdEntidadCancelar_Click()
    SetFlexNroEntidades
    EditarEntidad False
    HabilitarBotonera
End Sub

Private Sub cmdEntidadEditar_Click()
    If feEntidad.TextMatrix(1, 0) = "" Then Exit Sub
       
    EditarEntidad True
    fnEntidadAccion = Editar
    fnEntidadNoMoverFila = feEntidad.row
    feEntidad.col = 1
    feEntidad.SetFocus
End Sub

Private Sub cmdEntidadEliminar_Click()
    If feEntidad.TextMatrix(1, 0) = "" Then Exit Sub
    
    Dim fila As Integer
    Dim Z As Integer
    Dim iMatConfig As Integer
    
    fila = UBound(NroEntidades)
    
    feEntidad.row = fila
    feEntidad.col = 1
    
    If MsgBox("Se va a eliminar el último registro." & Chr(13) & "¿Está seguro de continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    iMatConfig = UBound(MatrizConfiguracion)
    
    'Validar en la matriz exista RSE para que usuario se dé cuenta
    For Z = 1 To iMatConfig
        If (MatrizConfiguracion(Z).IndexNroEntidad = fila And MatrizConfiguracion(Z).RSE = 1) Then
            If MsgBox("El tramo de Nro. de Entidades seleccionado tiene configuración en la matriz con RSE (SI)." & Chr(13) & "¿Está seguro de eliminar el registro?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                Exit Sub
            Else
                Exit For
            End If
        End If
    Next Z
    
    'Eliminar de la matriz las configuraciones de este tramo de Nro. de Entidades
    Dim MatrizConfiguracionTmp() As tMatrizConfig
    Dim iMC As Integer, iMCTmp As Integer
    
    ReDim MatrizConfiguracionTmp(0)
        
    For iMC = 1 To iMatConfig
        ReDim Preserve MatrizConfiguracionTmp(iMC)
        MatrizConfiguracionTmp(iMC) = MatrizConfiguracion(iMC)
    Next
    ReDim MatrizConfiguracion(0)
    iMC = 0
    For iMCTmp = 1 To iMatConfig
        If MatrizConfiguracionTmp(iMCTmp).IndexNroEntidad <> fila Then
            iMC = iMC + 1
            ReDim Preserve MatrizConfiguracion(iMC)
            MatrizConfiguracion(iMC) = MatrizConfiguracionTmp(iMCTmp)
        End If
    Next
    Erase MatrizConfiguracionTmp
    
    'Eliminar tramo de Nro. de Entidades
    Dim NroEntidadesTmp() As tEntidadConfig
    Dim Index As Integer, indexTmp As Integer
    ReDim NroEntidadesTmp(0)
        
    For Index = 1 To fila
        ReDim Preserve NroEntidadesTmp(Index)
        NroEntidadesTmp(Index) = NroEntidades(Index)
    Next
    ReDim NroEntidades(0)
    Index = 0
    For indexTmp = 1 To fila
        If indexTmp <> fila Then
            Index = Index + 1
            ReDim Preserve NroEntidades(Index)
            NroEntidades(Index) = NroEntidadesTmp(indexTmp)
        End If
    Next
    Erase NroEntidadesTmp

    SetFlexNroEntidades
    SetFlexMatrizConfiguracion
    
    fbRealizaCambios = True
    HabilitarBotonera
End Sub

Private Sub feEntidad_RowColChange()
    If feEntidad.lbEditarFlex Then
        feEntidad.row = fnEntidadNoMoverFila
    End If
End Sub

Private Sub feMatriz_DblClick()
    Dim row As Integer, col As Integer
    
    row = feMatriz.row
    col = feMatriz.col
    
    If col > 1 And row > 0 Then
        Select Case feMatriz.TextMatrix(row, col)
           Case "NO":
                'Call estableceRseMatrizConfig(row, col - 1, 1)
                If (estableceRseMatrizConfig(row, col - 1, 1)) Then
                    feMatriz.TextMatrix(row, col) = "SI"
                    feMatriz.CellBackColor = &H8080FF
                End If
           Case "SI":
                'Call estableceRseMatrizConfig(row, col - 1, 0)
                If (estableceRseMatrizConfig(row, col - 1, 0)) Then
                    feMatriz.TextMatrix(row, col) = "NO"
                    feMatriz.CellBackColor = &H80000005
                End If
        End Select
    End If
End Sub

Private Sub Form_Load()
    fnTipoConfiguracion = Seguimiento
    cmbTpoConfiguracion.ListIndex = -1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If fbRealizaCambios Then
        Dim lsTipoConfig As String
        lsTipoConfig = Trim(Mid(cmbTpoConfiguracion.Text, 1, Len(cmbTpoConfiguracion.Text) - 3))
        
        If MsgBox("Ud. ha realizado cambios en el formulario para " & lsTipoConfig + "." & Chr(13) & "De continuar los cambios realizados no se grabarán." & Chr(13) & Chr(13) & "¿Está seguro de salir de la opción?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdEntidadAceptar_Click()
    If Not validaEntidad Then Exit Sub
    
    Dim Index As Integer
    Dim i As Integer
    
    If fnEntidadAccion = Nuevo Then
        Index = UBound(NroEntidades) + 1
        ReDim Preserve NroEntidades(Index)
        NroEntidades(Index).Index = Index
        NroEntidades(Index).Inicio = CInt(feEntidad.TextMatrix(fnEntidadNoMoverFila, 1))
        NroEntidades(Index).Fin = CInt(feEntidad.TextMatrix(fnEntidadNoMoverFila, 2))
        NroEntidades(Index).Descripcion = Trim(feEntidad.TextMatrix(fnEntidadNoMoverFila, 3))
        NroEntidades(Index).TipoRegistro = Nuevo
        
        'Agregar a matriz de configuración
        Dim iMatConfig As Integer
        iMatConfig = UBound(MatrizConfiguracion)
        
        For i = 1 To UBound(TramosCP)
            ReDim Preserve MatrizConfiguracion(iMatConfig + i)
            MatrizConfiguracion(iMatConfig + i).IndexCP = i
            MatrizConfiguracion(iMatConfig + i).IndexNroEntidad = Index
            MatrizConfiguracion(iMatConfig + i).RSE = 0
        Next i
    ElseIf fnEntidadAccion = Editar Then
        NroEntidades(fnEntidadNoMoverFila).Index = fnEntidadNoMoverFila
        NroEntidades(fnEntidadNoMoverFila).Inicio = CInt(feEntidad.TextMatrix(fnEntidadNoMoverFila, 1))
        NroEntidades(fnEntidadNoMoverFila).Fin = CInt(feEntidad.TextMatrix(fnEntidadNoMoverFila, 2))
        NroEntidades(fnEntidadNoMoverFila).Descripcion = Trim(feEntidad.TextMatrix(fnEntidadNoMoverFila, 3))
        NroEntidades(fnEntidadNoMoverFila).TipoRegistro = Editar
    End If
    
    SetFlexNroEntidades
    SetFlexMatrizConfiguracion
    EditarEntidad False
    fnEntidadAccion = -1
    fnEntidadNoMoverFila = -1
    
    fbRealizaCambios = True
    HabilitarBotonera
End Sub

Private Sub HabilitarBotonera()
    cmdGrabar.Enabled = fbRealizaCambios
    cmdCancelar.Enabled = fbRealizaCambios
    cmdSalir.Enabled = Not fbRealizaCambios
    fraTipoConfiguracion.Enabled = Not fbRealizaCambios
End Sub

Private Sub cmdEntidadNuevo_Click()
    feEntidad.AdicionaFila
    feEntidad.SetFocus
    SendKeys "{ENTER}"
    
    EditarEntidad True
    fnEntidadAccion = eAccionRiesgoConfig.Nuevo
    fnEntidadNoMoverFila = feEntidad.row
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = False
End Sub

Private Sub EditarEntidad(ByVal pbEditar As Boolean)
    cmdEntidadNuevo.Enabled = Not pbEditar
    cmdEntidadEditar.Enabled = Not pbEditar
    cmdEntidadEliminar.Enabled = Not pbEditar
    cmdEntidadAceptar.Enabled = pbEditar
    cmdEntidadCancelar.Enabled = pbEditar
    
    feEntidad.lbEditarFlex = pbEditar
    fraTipoConfiguracion.Enabled = Not pbEditar
End Sub

Private Function validaEntidad(Optional ByVal pbGrabar As Boolean = False) As Boolean
    Dim i As Integer
    Dim j As Integer

    If feEntidad.TextMatrix(1, 0) = "" Then
        MsgBox "No se ha ingresado información para el registro de Nro. de Entidades", vbExclamation, "Aviso"
        TabSegConfig.Tab = 0
        EnfocaControl feEntidad
        Exit Function
    End If

    For i = 1 To feEntidad.rows - 1
        If (i = fnEntidadNoMoverFila) Then
            For j = 1 To feEntidad.cols - 1
                If feEntidad.ColWidth(j) > 0 Then
                    If Len(Trim(feEntidad.TextMatrix(i, j))) = 0 Then
                        MsgBox "El campo " & UCase(feEntidad.TextMatrix(0, j)) & " está vacio, verifique..", vbExclamation, "Aviso"
                        feEntidad.TabIndex = 1
                        EnfocaControl feEntidad
                        feEntidad.TopRow = i
                        feEntidad.row = i
                        feEntidad.col = j
                        Exit Function
                    End If
                End If
            Next

            Dim lnTramoInicioRegActual As Double, lnTramoFinRegActual As Double
            lnTramoInicioRegActual = CInt(feEntidad.TextMatrix(fnEntidadNoMoverFila, 1))
            lnTramoFinRegActual = CInt(feEntidad.TextMatrix(fnEntidadNoMoverFila, 2))

            'Validar Nro. Entidades Inicio con Nro. Entidades Fin
            If (lnTramoFinRegActual < lnTramoInicioRegActual) Then
                MsgBox "El Nro. Entidad Final [" & lnTramoFinRegActual & "] debe ser mayor o igual al Nro. Entidad Inicio [" & lnTramoInicioRegActual & "]", vbExclamation, "Aviso"
                feEntidad.TabIndex = 1
                EnfocaControl feEntidad
                feEntidad.row = fnEntidadNoMoverFila
                feEntidad.col = 1
                Exit Function
            End If

            If (fnEntidadNoMoverFila > 1) Then
                Dim lnTramoFinRegAnterior As Double
                lnTramoFinRegAnterior = NroEntidades(fnEntidadNoMoverFila - 1).Fin

                'Validar Nro. Entidades Inicio con Nro. Entidades Fin anterior
                If (lnTramoInicioRegActual <= lnTramoFinRegAnterior) Then
                    MsgBox "El Nro. Entidad Inicial [" & lnTramoInicioRegActual & "] debe ser mayor al Nro. Entidad Final anterior [" & lnTramoFinRegAnterior & "]", vbExclamation, "Aviso"
                    feEntidad.TabIndex = 1
                    EnfocaControl feEntidad
                    feEntidad.row = fnEntidadNoMoverFila
                    feEntidad.col = 1
                    Exit Function
                End If
            End If

            If (fnEntidadAccion = Editar) Then
                If (fnEntidadNoMoverFila < UBound(NroEntidades)) Then
                    Dim lnTramoInicioRegPosterior As Double
                    lnTramoInicioRegPosterior = NroEntidades(fnEntidadNoMoverFila + 1).Inicio

                    If (lnTramoInicioRegPosterior <= lnTramoFinRegActual) Then
                        MsgBox "El Nro. de Entidades Fin [" & lnTramoFinRegActual & "] debe ser menor al Nro. Entidades Inicio posterior [" & lnTramoInicioRegPosterior & "]", vbExclamation, "Aviso"
                        feEntidad.TabIndex = 1
                        EnfocaControl feEntidad
                        feEntidad.row = fnEntidadNoMoverFila
                        feEntidad.col = 1
                        Exit Function
                    End If
                End If
            End If
            
            'Nro. Entidades no puede ser mayor a 999
            If (lnTramoInicioRegActual > 999) Then
                MsgBox "El Nro. Entidades Inicial [" & lnTramoInicioRegActual & "] no puede ser mayor a [999]", vbExclamation, "Aviso"
                feEntidad.TabIndex = 0
                EnfocaControl feEntidad
                feEntidad.row = fnEntidadNoMoverFila
                feEntidad.col = 1
                Exit Function
            End If
            If (lnTramoFinRegActual > 999) Then
                MsgBox "El Nro. Entidades Fin [" & lnTramoFinRegActual & "] no puede ser mayor a [999]", vbExclamation, "Aviso"
                feEntidad.TabIndex = 0
                EnfocaControl feEntidad
                feEntidad.row = fnEntidadNoMoverFila
                feEntidad.col = 1
                Exit Function
            End If
        End If
    Next

    validaEntidad = True
End Function
Private Sub SetFlexNroEntidades()
    Dim Index As Integer, IndexFlex As Integer
    
    FormateaFlex feEntidad
    For Index = 1 To UBound(NroEntidades)
        Dim objEntidad As tEntidadConfig
        objEntidad = NroEntidades(Index)
        
        'If objEntidad.TipoRegistro <> Eliminar Then
        feEntidad.AdicionaFila
        IndexFlex = feEntidad.row
        feEntidad.TextMatrix(IndexFlex, 1) = objEntidad.Inicio
        feEntidad.TextMatrix(IndexFlex, 2) = objEntidad.Fin
        feEntidad.TextMatrix(IndexFlex, 3) = objEntidad.Descripcion
        'End If
    Next
End Sub
Private Sub SetFlexMatrizConfiguracion()
    Dim lsEncabezadosNombres As String
    Dim lsColumnasAEditar As String
    Dim lsEncabezadosAlineacion As String
    Dim lsEncabezadosAnchos As String
    Dim lsFormatosEdit As String
    Dim lsListaControles As String
    Dim i As Integer, j As Integer
    Dim iNroEntidades As Integer, iNroTramoCP As Integer
    Dim objEnt As tEntidadConfig
    Dim objTramoCP As tCPConfig
    Dim ixFlex As Integer
    Dim nRSE As Integer
    
    FormateaFlex feMatriz
    feMatriz.cols = 2
    
    lsEncabezadosNombres = "#-Capacidad de Pago"
    lsColumnasAEditar = "X-X"
    lsEncabezadosAlineacion = "C-L"
    lsEncabezadosAnchos = "0-1800"
    lsFormatosEdit = "0-0"
    lsListaControles = "0-0"
        
    iNroEntidades = UBound(NroEntidades)
    iNroTramoCP = UBound(TramosCP)

    For j = 1 To iNroEntidades
        objEnt = NroEntidades(j)
        lsEncabezadosNombres = lsEncabezadosNombres & "-" & objEnt.Descripcion
        lsColumnasAEditar = lsColumnasAEditar & "-X"
        lsEncabezadosAlineacion = lsEncabezadosAlineacion & "-C"
        lsEncabezadosAnchos = lsEncabezadosAnchos & "-1800"
        lsFormatosEdit = lsFormatosEdit & "-0"
        lsListaControles = lsListaControles & IIf(i = 1, "-0", "-0")
    Next

    feMatriz.EncabezadosNombres = lsEncabezadosNombres
    feMatriz.ColumnasAEditar = lsColumnasAEditar
    feMatriz.EncabezadosAlineacion = lsEncabezadosAlineacion
    feMatriz.EncabezadosAnchos = lsEncabezadosAnchos
    feMatriz.FormatosEdit = lsFormatosEdit
    feMatriz.ListaControles = lsListaControles

    For i = 1 To iNroTramoCP
        feMatriz.AdicionaFila , , True
        ixFlex = feMatriz.rows - 1
        
        feMatriz.row = i
        objTramoCP = TramosCP(i)
        
        feMatriz.TextMatrix(ixFlex, 1) = objTramoCP.Descripcion
        
        'Se empieza de la columna 2 porque el flexedit al empezar de la columna 1 y si se daba click en la cabecera quitaba la primera columna
        For j = 2 To iNroEntidades + 1
            feMatriz.col = j
            
            nRSE = obtieneRseMatrizConfig(i, j - 1)
            
            Select Case nRSE
               Case 0:
                   feMatriz.TextMatrix(ixFlex, j) = "NO"
                   feMatriz.CellBackColor = &H80000005
               Case 1:
                   feMatriz.TextMatrix(ixFlex, j) = "SI"
                   feMatriz.CellBackColor = &H8080FF
            End Select
        Next
    Next

    feMatriz.FixedCols = IIf(feMatriz.cols >= 3, 2, 0)
End Sub

Private Function estableceRseMatrizConfig(pnIndexCP As Integer, pnIndexNroEnt As Integer, pnRSE As Integer) As Boolean
    Dim iNroMatrizConfig As Integer
    Dim i As Integer, j As Integer
    Dim obj As tMatrizConfig
    Dim retorno As Boolean

    iNroMatrizConfig = UBound(MatrizConfiguracion)
    
    For i = 1 To iNroMatrizConfig
        If MatrizConfiguracion(i).IndexCP = pnIndexCP And MatrizConfiguracion(i).IndexNroEntidad = pnIndexNroEnt Then
            MatrizConfiguracion(i).RSE = pnRSE
            fbRealizaCambios = True
            HabilitarBotonera
            retorno = True
            Exit For
        End If
    Next
    
    estableceRseMatrizConfig = retorno
End Function

Private Function obtieneRseMatrizConfig(pnIndexCP As Integer, pnIndexNroEnt As Integer)
    Dim iNroMatrizConfig As Integer
    Dim nRSE As Integer, i As Integer, j As Integer
    Dim obj As tMatrizConfig
    
    nRSE = 0
    iNroMatrizConfig = UBound(MatrizConfiguracion)
    
    For i = 1 To iNroMatrizConfig
        obj = MatrizConfiguracion(i)
        If obj.IndexCP = pnIndexCP And obj.IndexNroEntidad = pnIndexNroEnt Then
            nRSE = obj.RSE
            Exit For
        End If
    Next
    
    obtieneRseMatrizConfig = nRSE
End Function

Private Sub cmdTramoCPCancelar_Click()
    SetFlexTramosCP
    EditarTramoCP False
    HabilitarBotonera
End Sub

Private Sub cmdTramoCPEditar_Click()
    If feTramoCP.TextMatrix(1, 0) = "" Then Exit Sub
       
    EditarTramoCP True
    fnTramoCPAccion = Editar
    fnTramoCPNoMoverFila = feTramoCP.row
    feTramoCP.col = 1
    feTramoCP.SetFocus
End Sub

Private Sub cmdTramoCPEliminar_Click()
    If feTramoCP.TextMatrix(1, 0) = "" Then Exit Sub
    
    Dim fila As Integer
    Dim Z As Integer
    Dim iMatConfig As Integer
    
    fila = UBound(TramosCP)
    
    feTramoCP.row = fila
    feTramoCP.col = 1
       
    If MsgBox("Se va a eliminar el último registro." & Chr(13) & "¿Está seguro de continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    iMatConfig = UBound(MatrizConfiguracion)
    
    'Validar en la matriz exista RSE para que usuario se dé cuenta
    For Z = 1 To iMatConfig
        If (MatrizConfiguracion(Z).IndexCP = fila And MatrizConfiguracion(Z).RSE = 1) Then
            If MsgBox("El tramo CP seleccionado tiene configuración en la matriz con RSE (SI)." & Chr(13) & "¿Está seguro de eliminar el registro?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                Exit Sub
            Else
                Exit For
            End If
        End If
    Next Z
    
    'Eliminar de la matriz las configuraciones de este tramo CP
    Dim MatrizConfiguracionTmp() As tMatrizConfig
    Dim iMC As Integer, iMCTmp As Integer
    
    ReDim MatrizConfiguracionTmp(0)
        
    For iMC = 1 To iMatConfig
        ReDim Preserve MatrizConfiguracionTmp(iMC)
        MatrizConfiguracionTmp(iMC) = MatrizConfiguracion(iMC)
    Next
    ReDim MatrizConfiguracion(0)
    iMC = 0
    For iMCTmp = 1 To iMatConfig
        If MatrizConfiguracionTmp(iMCTmp).IndexCP <> fila Then
            iMC = iMC + 1
            ReDim Preserve MatrizConfiguracion(iMC)
            MatrizConfiguracion(iMC) = MatrizConfiguracionTmp(iMCTmp)
        End If
    Next
    Erase MatrizConfiguracionTmp
    
    'Eliminar tramo de Capacidad de Pago
    Dim TramosCPTmp() As tCPConfig
    Dim Index As Integer, indexTmp As Integer
    ReDim TramosCPTmp(0)
        
    For Index = 1 To fila
        ReDim Preserve TramosCPTmp(Index)
        TramosCPTmp(Index) = TramosCP(Index)
    Next
    ReDim TramosCP(0)
    Index = 0
    For indexTmp = 1 To fila
        If indexTmp <> fila Then
            Index = Index + 1
            ReDim Preserve TramosCP(Index)
            TramosCP(Index) = TramosCPTmp(indexTmp)
        End If
    Next
    Erase TramosCPTmp
    
    SetFlexTramosCP
    SetFlexMatrizConfiguracion
    
    fbRealizaCambios = True
    HabilitarBotonera
End Sub

Private Sub feTramoCP_RowColChange()
    If feTramoCP.lbEditarFlex Then
        feTramoCP.row = fnTramoCPNoMoverFila
    End If
End Sub

Private Sub cmdTramoCPAceptar_Click()
    If Not validaTramoCP Then Exit Sub
    
    Dim Index As Integer
    Dim j As Integer
    
    If fnTramoCPAccion = Nuevo Then
        Index = UBound(TramosCP) + 1
        ReDim Preserve TramosCP(Index)
        TramosCP(Index).Index = Index
        TramosCP(Index).Inicio = CDbl(feTramoCP.TextMatrix(fnTramoCPNoMoverFila, 1))
        TramosCP(Index).Fin = CDbl(feTramoCP.TextMatrix(fnTramoCPNoMoverFila, 2))
        TramosCP(Index).Descripcion = Trim(feTramoCP.TextMatrix(fnTramoCPNoMoverFila, 3))
        TramosCP(Index).TipoRegistro = Nuevo
        
        'Agregar a matriz de configuración
        Dim iMatConfig As Integer
        iMatConfig = UBound(MatrizConfiguracion)
        
        For j = 1 To UBound(NroEntidades)
            ReDim Preserve MatrizConfiguracion(iMatConfig + j)
            MatrizConfiguracion(iMatConfig + j).IndexCP = Index
            MatrizConfiguracion(iMatConfig + j).IndexNroEntidad = j
            MatrizConfiguracion(iMatConfig + j).RSE = 0
        Next j
    ElseIf fnTramoCPAccion = Editar Then
        TramosCP(fnTramoCPNoMoverFila).Index = fnTramoCPNoMoverFila
        TramosCP(fnTramoCPNoMoverFila).Inicio = CDbl(feTramoCP.TextMatrix(fnTramoCPNoMoverFila, 1))
        TramosCP(fnTramoCPNoMoverFila).Fin = CDbl(feTramoCP.TextMatrix(fnTramoCPNoMoverFila, 2))
        TramosCP(fnTramoCPNoMoverFila).Descripcion = Trim(feTramoCP.TextMatrix(fnTramoCPNoMoverFila, 3))
        TramosCP(fnTramoCPNoMoverFila).TipoRegistro = Editar
    End If
    
    SetFlexTramosCP
    SetFlexMatrizConfiguracion
    EditarTramoCP False
    fnTramoCPAccion = -1
    fnTramoCPNoMoverFila = -1
    
    fbRealizaCambios = True
    HabilitarBotonera
End Sub

Private Sub cmdTramoCPNuevo_Click()
    feTramoCP.AdicionaFila
    feTramoCP.SetFocus
    SendKeys "{ENTER}"
    
    EditarTramoCP True
    fnTramoCPAccion = eAccionRiesgoConfig.Nuevo
    fnTramoCPNoMoverFila = feTramoCP.row
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = False
End Sub

Private Sub EditarTramoCP(ByVal pbEditar As Boolean)
    cmdTramoCPNuevo.Enabled = Not pbEditar
    cmdTramoCPEditar.Enabled = Not pbEditar
    cmdTramoCPEliminar.Enabled = Not pbEditar
    cmdTramoCPAceptar.Enabled = pbEditar
    cmdTramoCPCancelar.Enabled = pbEditar
    
    feTramoCP.lbEditarFlex = pbEditar
    fraTipoConfiguracion.Enabled = Not pbEditar
End Sub

Private Function validaTramoCP(Optional ByVal pbGrabar As Boolean = False) As Boolean
    Dim i As Integer
    Dim j As Integer

    If feTramoCP.TextMatrix(1, 0) = "" Then
        MsgBox "No se ha ingresado ningún Tramo de Capacidad de Pago", vbExclamation, "Aviso"
        TabSegConfig.Tab = 0
        EnfocaControl feTramoCP
        Exit Function
    End If

    For i = 1 To feTramoCP.rows - 1
        If (i = fnTramoCPNoMoverFila) Then
            For j = 1 To feTramoCP.cols - 1
                If feTramoCP.ColWidth(j) > 0 Then
                    If Len(Trim(feTramoCP.TextMatrix(i, j))) = 0 Then
                        MsgBox "El campo " & UCase(feTramoCP.TextMatrix(0, j)) & " está vacio, verifique..", vbExclamation, "Aviso"
                        feTramoCP.TabIndex = 0
                        EnfocaControl feTramoCP
                        feTramoCP.TopRow = i
                        feTramoCP.row = i
                        feTramoCP.col = j
                        Exit Function
                    End If
                End If
            Next
            
            Dim lnTramoInicioRegActual As Double, lnTramoFinRegActual As Double
            lnTramoInicioRegActual = CDbl(feTramoCP.TextMatrix(fnTramoCPNoMoverFila, 1))
            lnTramoFinRegActual = CDbl(feTramoCP.TextMatrix(fnTramoCPNoMoverFila, 2))
            
            'Validar CP Inicio con CP Fin actual
            If (lnTramoFinRegActual <= lnTramoInicioRegActual) Then
                MsgBox "La CP Final [" & Format(lnTramoFinRegActual, "#0.00") & "] debe ser mayor a la CP Inicio [" & Format(lnTramoInicioRegActual, "#0.00") & "]", vbExclamation, "Aviso"
                feTramoCP.TabIndex = 0
                EnfocaControl feTramoCP
                feTramoCP.row = fnTramoCPNoMoverFila
                feTramoCP.col = 1
                Exit Function
            End If
            
            If (fnTramoCPNoMoverFila > 1) Then
                Dim lnTramoFinRegAnterior As Double
                lnTramoFinRegAnterior = TramosCP(fnTramoCPNoMoverFila - 1).Fin
                
                'Validar CP Inicio actual con CP Fin anterior
                If (lnTramoInicioRegActual <= lnTramoFinRegAnterior) Then
                    MsgBox "La CP Inicial [" & Format(lnTramoInicioRegActual, "#0.00") & "] debe ser mayor a la CP Final anterior [" & Format(lnTramoFinRegAnterior, "#0.00") & "]", vbExclamation, "Aviso"
                    feTramoCP.TabIndex = 0
                    EnfocaControl feTramoCP
                    feTramoCP.row = fnTramoCPNoMoverFila
                    feTramoCP.col = 1
                    Exit Function
                End If
            End If
            
            If (fnTramoCPAccion = Editar) Then
                If (fnTramoCPNoMoverFila < UBound(TramosCP)) Then
                    Dim lnTramoInicioRegPosterior As Double
                    lnTramoInicioRegPosterior = TramosCP(fnTramoCPNoMoverFila + 1).Inicio
                    
                    If (lnTramoInicioRegPosterior <= lnTramoFinRegActual) Then
                        MsgBox "La CP Fin [" & Format(lnTramoFinRegActual, "#0.00") & "] debe ser menor a la CP Inicio posterior [" & Format(lnTramoInicioRegPosterior, "#0.00") & "]", vbExclamation, "Aviso"
                        feTramoCP.TabIndex = 0
                        EnfocaControl feTramoCP
                        feTramoCP.row = fnTramoCPNoMoverFila
                        feTramoCP.col = 1
                        Exit Function
                    End If
                End If
            End If
            
            'Tramo no puede ser mayor a 999.99
            If (lnTramoInicioRegActual > 999.99) Then
                MsgBox "La CP Inicial [" & Format(lnTramoInicioRegActual, "#0.00") & "] no puede ser mayor a [999.99]", vbExclamation, "Aviso"
                feTramoCP.TabIndex = 0
                EnfocaControl feTramoCP
                feTramoCP.row = fnTramoCPNoMoverFila
                feTramoCP.col = 1
                Exit Function
            End If
            If (lnTramoFinRegActual > 999.99) Then
                MsgBox "La CP Fin [" & Format(lnTramoFinRegActual, "#0.00") & "] no puede ser mayor a [999.99]", vbExclamation, "Aviso"
                feTramoCP.TabIndex = 0
                EnfocaControl feTramoCP
                feTramoCP.row = fnTramoCPNoMoverFila
                feTramoCP.col = 1
                Exit Function
            End If
        End If
    Next

    validaTramoCP = True
End Function

Private Sub SetFlexTramosCP()
    Dim Index As Integer, IndexFlex As Integer
    
    FormateaFlex feTramoCP
    For Index = 1 To UBound(TramosCP)
        Dim objTramoCP As tCPConfig
        objTramoCP = TramosCP(Index)
        
        'If objTramoCP.TipoRegistro <> Eliminar Then
        feTramoCP.AdicionaFila
        IndexFlex = feTramoCP.row
        feTramoCP.TextMatrix(IndexFlex, 1) = Format(objTramoCP.Inicio, "#0.00")
        feTramoCP.TextMatrix(IndexFlex, 2) = Format(objTramoCP.Fin, "#0.00")
        feTramoCP.TextMatrix(IndexFlex, 3) = objTramoCP.Descripcion
        'End If
    Next
End Sub
