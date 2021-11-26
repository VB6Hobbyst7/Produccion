VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersBusqueda 
   Caption         =   "Busqueda de Cliente por Dirección"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13425
   Icon            =   "frmPersBusqueda.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   13425
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Filtro de Búsqueda"
      TabPicture(0)   =   "frmPersBusqueda.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTotal"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraUbigeo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "feCliente"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDireccion"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSalir"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCancelar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdVer"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.CommandButton cmdVer 
         Caption         =   "Ver"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   8400
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   10560
         TabIndex        =   16
         Top             =   8400
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   495
         Left            =   11880
         TabIndex        =   15
         Top             =   8400
         Width           =   1215
      End
      Begin VB.TextBox txtDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   14
         Top             =   1920
         Width           =   4455
      End
      Begin SICMACT.FlexEdit feCliente 
         Height          =   5895
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   10398
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cod. Cliente-Nombre-DOI-Dirección-Total-Credito"
         EncabezadosAnchos=   "500-1300-3700-1200-4200-600-1000"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-6"
         ListaControles  =   "0-0-0-0-0-0-1"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbOrdenaCol     =   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Frame fraUbigeo 
         Caption         =   "Ubigeo"
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
         Height          =   1335
         Left            =   2640
         TabIndex        =   1
         Top             =   480
         Width           =   8535
         Begin VB.CommandButton cmdSeleccionar 
            Caption         =   "Seleccionar"
            Height          =   495
            Left            =   7320
            TabIndex        =   18
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Index           =   3
            ItemData        =   "frmPersBusqueda.frx":0326
            Left            =   5025
            List            =   "frmPersBusqueda.frx":0328
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   855
            Width           =   2190
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Index           =   2
            ItemData        =   "frmPersBusqueda.frx":032A
            Left            =   1080
            List            =   "frmPersBusqueda.frx":032C
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   855
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Index           =   4
            ItemData        =   "frmPersBusqueda.frx":032E
            Left            =   3390
            List            =   "frmPersBusqueda.frx":0330
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   135
            Visible         =   0   'False
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Index           =   1
            ItemData        =   "frmPersBusqueda.frx":0332
            Left            =   5040
            List            =   "frmPersBusqueda.frx":0334
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   495
            Width           =   2175
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Index           =   0
            ItemData        =   "frmPersBusqueda.frx":0336
            Left            =   1080
            List            =   "frmPersBusqueda.frx":0338
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Distrito :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3615
            TabIndex        =   11
            Top             =   855
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Zona : "
            Height          =   195
            Left            =   2550
            TabIndex        =   10
            Top             =   150
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pais : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   9
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Departamento :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3615
            TabIndex        =   8
            Top             =   480
            Width           =   1320
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Provincia :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   7
            Top             =   840
            Width           =   930
         End
      End
      Begin VB.Label lblTotal 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   2880
         TabIndex        =   21
         Top             =   8400
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   19
         Top             =   8400
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Dirección:"
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
         Left            =   3960
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   20
      Top             =   8520
      Width           =   855
   End
End
Attribute VB_Name = "frmPersBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmPersBusqueda
'***     Descripcion:       Permite realizar la busqueda de personas por su direccion
'***     Creado por:        FRHU
'***     Fecha-Tiempo:         24/04/2014 01:00:00 PM
'*****************************************************************************************
Option Explicit
Dim nPos As Integer
Dim i As Integer
Dim bEstadoCargando As Boolean
Dim oPersona As New UPersona_Cli   ' COMDPersona.DCOMPersona
'Seleccionar
Dim distrito As String

Private Sub CmdSeleccionar_Click()
    If Not ValidarSeleccion Then
        Exit Sub
    End If
    'Distrito
    distrito = Trim(Right(cmbPersUbiGeo(3).Text, 15))
    '***
    Me.fraUbigeo.Enabled = False
    Me.txtDireccion.Enabled = True
    Me.cmdSeleccionar.Enabled = False
    Me.cmdCancelar.Enabled = True
    Me.cmdVer.Enabled = True
    Me.txtDireccion.SetFocus
End Sub
Private Sub cmdCancelar_Click()
    Me.fraUbigeo.Enabled = True
    Me.txtDireccion.Enabled = False
    FormateaFlex feCliente
    Me.txtDireccion.Text = ""
    Me.cmdSeleccionar.Enabled = True
    Me.cmdCancelar.Enabled = False
    Me.cmdVer.Enabled = False
    Call LimpiarPantalla
    lblTotal.Caption = "0"
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub CmdSeleccionar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDireccion.SetFocus
End Sub

Private Sub cmdVer_Click()
    Dim CodPer As String
    CodPer = Me.feCliente.TextMatrix(feCliente.row, 1)
    If feCliente.Col <> 5 And feCliente.Col <> 6 Then
        If CodPer = "" Then
            MsgBox "No hay datos que mostrar", vbInformation
            Exit Sub
        Else
            Call frmPersona.ConsultarPorPersona(CodPer)
        End If
    End If
End Sub

'Private Sub feCliente_DblClick()
'    Dim CodPer As String
'    If feCliente.row > 0 Then
'        CodPer = Me.feCliente.TextMatrix(feCliente.row, 1)
'        If feCliente.Col <> 5 And feCliente.Col <> 6 Then
'            If CodPer = "" Then
'                MsgBox "No hay datos que mostrar", vbInformation
'                Exit Sub
'            Else
'                Call frmPersona.ConsultarPorPersona(CodPer)
'            End If
'        End If
'    End If
'End Sub

Private Sub feCliente_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim sTotal As String
    Dim sPersCod As String
    
    'sTotal = feCliente.TextMatrix(pnRow, 6)
    sPersCod = Trim(feCliente.TextMatrix(pnRow, 1))
    If sPersCod <> "" Then
        frmPosicionCli.iniciarFormulario sPersCod
    End If
    'feCliente.TextMatrix(pnRow, 6) = sTotal
End Sub
Private Sub Form_Load()
   Call CargarControles
End Sub
Private Sub LimpiarPantalla()
    cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), "04028")
    cmbPersUbiGeo(1).ListIndex = -1
    cmbPersUbiGeo(2).ListIndex = -1
    cmbPersUbiGeo(3).ListIndex = -1
    cmbPersUbiGeo(4).ListIndex = -1
End Sub
Private Function ValidarSeleccion() As Boolean
    If Me.cmbPersUbiGeo(0).ListIndex = -1 Then
        MsgBox "Seleccione Un Pais", vbInformation
        ValidarSeleccion = False
        Exit Function
    ElseIf Me.cmbPersUbiGeo(1).ListIndex = -1 Then
        MsgBox "Seleccione el Departamento", vbInformation
        ValidarSeleccion = False
        Exit Function
    ElseIf Me.cmbPersUbiGeo(2).ListIndex = -1 Then
        MsgBox "Seleccione la Provincia", vbInformation
        ValidarSeleccion = False
        Exit Function
    ElseIf Me.cmbPersUbiGeo(3).ListIndex = -1 Then
        MsgBox "Seleccione el Distrito", vbInformation
        ValidarSeleccion = False
        Exit Function
    End If
    ValidarSeleccion = True
End Function
Private Sub CargarControles()
    '***** UBIGEO
    bEstadoCargando = True
    Dim oPersonas As New COMDPersona.DCOMPersonas
    Dim lrsUbiGeo As ADODB.Recordset
    Set lrsUbiGeo = oPersonas.CargarUbicacionesGeograficas(True, 0)
    While Not lrsUbiGeo.EOF
        Me.cmbPersUbiGeo(0).AddItem Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
        lrsUbiGeo.MoveNext
    Wend
    If lrsUbiGeo.RecordCount > 0 Then lrsUbiGeo.MoveFirst
    For i = 0 To lrsUbiGeo.RecordCount
        If Trim(lrsUbiGeo!cUbiGeoCod) = "04028" Then
            nPos = i
        End If
    Next i
    Me.cmbPersUbiGeo(0).ListIndex = nPos
    Call LimpiarPantalla
    bEstadoCargando = False
End Sub
Private Sub cmbPersUbiGeo_Change(Index As Integer)
     If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.UbicacionGeografica = Trim(Right(cmbPersUbiGeo(4).Text, 15))
     End If
End Sub
Private Sub cmbPersUbiGeo_Click(Index As Integer)
    Dim oUbic As COMDPersona.DCOMPersonas
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    If Index <> 4 Then
    
        Set oUbic = New COMDPersona.DCOMPersonas
    
        Set rs = oUbic.CargarUbicacionesGeograficas(True, Index + 1, Trim(Right(cmbPersUbiGeo(Index).Text, 15)))
    
        If Trim(Right(cmbPersUbiGeo(0).Text, 12)) <> "04028" Then
         'MADM 20101228
             If Index = 0 Then
                For i = 1 To cmbPersUbiGeo.Count - 1
                    cmbPersUbiGeo(i).Clear
                    cmbPersUbiGeo(i).AddItem Trim(Trim(cmbPersUbiGeo(0).Text)) & Space(50) & Trim(Right(cmbPersUbiGeo(0).Text, 12))
                Next i
             End If
        'END MADM
        Else
            For i = Index + 1 To cmbPersUbiGeo.Count - 1
            cmbPersUbiGeo(i).Clear
            Next
            
            While Not rs.EOF
                cmbPersUbiGeo(Index + 1).AddItem Trim(rs!cUbiGeoDescripcion) & Space(50) & Trim(rs!cUbiGeoCod)
                rs.MoveNext
            Wend
        End If
        Set oUbic = Nothing
    End If
    
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.UbicacionGeografica = Trim(Right(cmbPersUbiGeo(4).Text, 15))
    End If
End Sub
Private Sub cmbPersUbiGeo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index < 3 Then
            cmbPersUbiGeo(Index + 1).SetFocus
        Else
            Me.cmdSeleccionar.SetFocus
        End If
    End If
End Sub
Private Sub txtDireccion_Change()
    Dim Fila As Integer
    Dim Palabra As String
    Dim oCred As New COMDCredito.DCOMCredito
    Dim rs As ADODB.Recordset
    Palabra = txtDireccion.Text
    If Palabra = "" Then
        Call FormateaFlex(feCliente)
    Else
        Set rs = oCred.ObtenerClientePorDireccion(Palabra, distrito)
        Fila = 0
        Call FormateaFlex(feCliente)
        Do While Not rs.EOF
            Fila = Fila + 1
            feCliente.AdicionaFila
            'fila = feCliente.row
            feCliente.TextMatrix(Fila, 1) = rs!cPersCod
            feCliente.TextMatrix(Fila, 2) = rs!cPersNombre
            feCliente.TextMatrix(Fila, 3) = rs!cPersIDnro
            feCliente.TextMatrix(Fila, 4) = rs!cPersDireccDomicilio
            feCliente.TextMatrix(Fila, 5) = rs!Total
            rs.MoveNext
        Loop
        lblTotal.Caption = CStr(Fila)
    End If
End Sub

