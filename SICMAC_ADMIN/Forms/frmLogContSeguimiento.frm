VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogContSeguimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contratación: Seguimiento de Contratos"
   ClientHeight    =   7635
   ClientLeft      =   2220
   ClientTop       =   1200
   ClientWidth     =   12630
   Icon            =   "frmLogContSeguimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTContratos 
      Height          =   7500
      Left            =   80
      TabIndex        =   0
      Top             =   80
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   13229
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lista de Contratos Vigentes"
      TabPicture(0)   =   "frmLogContSeguimiento.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraArea"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "feCrontratos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin Sicmact.FlexEdit feCrontratos 
         Height          =   6255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   12135
         _ExtentX        =   21378
         _ExtentY        =   10927
         Cols0           =   13
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nº Contrato-Proveedor-Moneda-Monto-Desde-Hasta-Nº Cuota-Estado-cNAdenda-cTpoAdenda-nTipoContrato-cNContRef"
         EncabezadosAnchos=   "500-2000-3500-1000-1200-1200-1200-1200-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-C-C-C-C-C-L-L-R-R"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-3-3"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   7
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Frame Frame2 
         Caption         =   "Opciones"
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
         Height          =   660
         Left            =   8760
         TabIndex        =   11
         Top             =   6720
         Width           =   3600
         Begin VB.CommandButton cmdActualizar 
            Caption         =   "&Actualizar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1250
            TabIndex        =   14
            Top             =   240
            Width           =   1110
         End
         Begin VB.CommandButton cmdCerrar 
            Cancel          =   -1  'True
            Caption         =   "&Cerrar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2380
            TabIndex        =   13
            Top             =   240
            Width           =   1110
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1110
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Contrato"
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
         Height          =   660
         Left            =   4440
         TabIndex        =   8
         Top             =   6720
         Width           =   3720
         Begin VB.CommandButton cmdExtornarCont 
            Caption         =   "&Extornar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1350
            TabIndex        =   16
            Top             =   240
            Width           =   1110
         End
         Begin VB.CommandButton cmdVerPDF 
            Caption         =   "&Ver PDF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2480
            TabIndex        =   10
            Top             =   240
            Width           =   1110
         End
         Begin VB.CommandButton cmdCronograma 
            Caption         =   "&Cronograma"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraArea 
         Caption         =   "Adendas"
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
         Height          =   660
         Left            =   240
         TabIndex        =   5
         Top             =   6720
         Width           =   3600
         Begin VB.CommandButton cmdExtornar 
            Caption         =   "&Extornar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2335
            TabIndex        =   15
            Top             =   240
            Width           =   1110
         End
         Begin VB.CommandButton cmdVer 
            Caption         =   "&Ver"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1200
            TabIndex        =   7
            Top             =   240
            Width           =   1110
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
            Height          =   360
            Left            =   80
            TabIndex        =   6
            Top             =   240
            Width           =   1110
         End
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3120
         TabIndex        =   2
         Top             =   2160
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "_"
         ForeColor       =   &H80000008&
         Height          =   75
         Left            =   3240
         TabIndex        =   1
         Top             =   2760
         Width           =   90
      End
   End
   Begin VB.PictureBox CdlgFile 
      Height          =   615
      Left            =   7440
      ScaleHeight     =   555
      ScaleWidth      =   735
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   800
   End
End
Attribute VB_Name = "frmLogContSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Integer
Dim psRutaContrato As String

Private Sub cmdActualizar_Click()
Screen.MousePointer = 11
Call CargarGrid
Screen.MousePointer = 0
End Sub

Private Sub cmdAgregar_Click()
Dim row As Long
If ValidarSeleccion Then
    row = feCrontratos.row
    If feCrontratos.TextMatrix(row, 11) >= 5 Then   'PASIERS0772014
        frmLogContReajAdenda.Inicio Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 1)), Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 12)), Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 11)) ', feCrontratos.TextMatrix(Me.feCrontratos.row, 9)
    Else
        If feCrontratos.TextMatrix(row, 11) = 3 Then
            frmLogContTipoAdenda.Inicio Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 1)), Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 12))
        Else
            frmLogContRegAdendas.Inicio Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 1)), , , LogTipoContrato.ContratoArrendamiento, Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 12)) 'Modificado pasi20140823 ti-ers077-2014
            feCrontratos.row = row
            feCrontratos.TopRow = row
        End If
    End If
      cmdActualizar_Click
End If
End Sub
Private Sub cmdCerrar_Click()
Unload Me
End Sub
Sub CargarGrid()
Dim oLog As DLogGeneral
Dim rsLog As ADODB.Recordset

Set oLog = New DLogGeneral
Set rsLog = oLog.ListarContratosPorEstado("4,5")
Call LimpiaFlex(Me.feCrontratos)
If rsLog.RecordCount > 0 Then
    For i = 0 To rsLog.RecordCount - 1
        feCrontratos.AdicionaFila
        Me.feCrontratos.TextMatrix(i + 1, 0) = i + 1
        Me.feCrontratos.TextMatrix(i + 1, 1) = rsLog!NContrato
        Me.feCrontratos.TextMatrix(i + 1, 2) = rsLog!Proveedor
        Me.feCrontratos.TextMatrix(i + 1, 3) = rsLog!Moneda
        Me.feCrontratos.TextMatrix(i + 1, 4) = rsLog!monto
        Me.feCrontratos.TextMatrix(i + 1, 5) = Format(rsLog!Desde, "dd/mm/yyyy")
        Me.feCrontratos.TextMatrix(i + 1, 6) = Format(rsLog!Hasta, "dd/mm/yyyy")
        Me.feCrontratos.TextMatrix(i + 1, 7) = oLog.ObtenerUltCuotaContratos(Trim(rsLog!NContrato), rsLog!cNContRef)
        Me.feCrontratos.TextMatrix(i + 1, 8) = Trim(rsLog!nEstado)
        'EJVG20131204 ***
        feCrontratos.TextMatrix(i + 1, 9) = rsLog!cNAdenda
        'END EJVG *******
        'PASI20140822 TI-ERS077-2014
        feCrontratos.TextMatrix(i + 1, 10) = rsLog!nTipo
        feCrontratos.TextMatrix(i + 1, 11) = rsLog!nTipoContrato
        feCrontratos.TextMatrix(i + 1, 12) = rsLog!cNContRef
        'end PASI
        rsLog.MoveNext
    Next i
Else
    MsgBox "No hay datos.", vbInformation, "Aviso"
End If
End Sub
Private Sub cmdCronograma_Click()
If ValidarSeleccion Then
    frmLogContCronograma.Inicio Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 1)), Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 12)) 'PASI20140823 TI-ERS077-2014
End If
End Sub

Private Sub cmdExtornar_Click()
Dim row As Long
Dim lnNAdenda As Integer
Dim lnTpoAdenda As Integer
If ValidarSeleccion Then
    'frmLogContExtAdendas.Inicio (Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 1)))
    row = feCrontratos.row
    lnNAdenda = feCrontratos.TextMatrix(Me.feCrontratos.row, 9)
    lnTpoAdenda = feCrontratos.TextMatrix(Me.feCrontratos.row, 10)
    If lnNAdenda <= 0 Then
        MsgBox "El Presente contrato no cuenta con Adendas", vbInformation, "Aviso"
        Exit Sub
    End If
    frmLogContExtAdendas.Inicio (Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 1))), lnNAdenda, lnTpoAdenda, feCrontratos.TextMatrix(feCrontratos.row, 11), feCrontratos.TextMatrix(feCrontratos.row, 12)
    cmdActualizar_Click
    feCrontratos.row = row
    feCrontratos.TopRow = row
End If
End Sub

Private Sub cmdExtornarCont_Click()
Dim oLog As New DLogGeneral 'PASI20141821 TI-ERS077-2014
Dim nExisCuotasPagad As Integer 'PASI20141821 TI-ERS077-2014
Set oLog = New DLogGeneral
Dim row As Long
Dim lsNContrato As String
If ValidarSeleccion Then
'    Select Case Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 8))
'    Case "1": frmLogContExtorno.Inicio (Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 1)))
'    Case "2": MsgBox "Contrato se encuentra Vigente", vbInformation, "Aviso"
'    Case "3": MsgBox "Contrato se encuentra Vencido", vbInformation, "Aviso"
'    End Select
    'PASI20141821 TI-ERS077-2014
    nExisCuotasPagad = oLog.ExistenCuotasPagadas(Trim(Me.feCrontratos.TextMatrix(Me.feCrontratos.row, 1)), Trim(Me.feCrontratos.TextMatrix(Me.feCrontratos.row, 12)))
    If nExisCuotasPagad = 1 Then
        MsgBox "El extorno no se puede realizar por que ya existen cuotas Pagadas.", vbInformation, "Aviso"
        Exit Sub
    End If
    'end PASI

    row = feCrontratos.row
    lsNContrato = Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 1))
    frmLogContExtorno.Inicio lsNContrato, Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 12))
    cmdActualizar_Click
    feCrontratos.row = row
    feCrontratos.TopRow = row
End If
End Sub

Private Sub cmdImprimir_Click()
Dim oImpre As New COMFunciones.FCOMImpresion

If Me.feCrontratos.TextMatrix(1, 1) = "" Then
    MsgBox "No existen contratos vigentes.", vbInformation, "Aviso"
    Me.cmdActualizar.SetFocus
    Exit Sub
End If

Dim lsCadena As String
Dim lnPagina As Long
Dim lnItem As Long
Dim lnI As Long
Dim oPrevio As clsPrevio
    
    
Set oPrevio = New clsPrevio
    
Dim lsNContrato As String * 25
Dim lsProveedor As String * 45
Dim lsMoneda As String * 8
Dim lsMonto As String * 20
Dim lsDesde As String * 15
Dim lsHasta As String * 15
Dim lsNCuotas As String * 8
    

    
Dim oCon As DConecta
Set oCon = New DConecta
    
lsCadena = ""

lsCadena = lsCadena & oImpresora.gPrnCondensadaON
lsCadena = lsCadena & CabeceraPagina1("C O N T R A T O S     V I G E N T E S", lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "")
lsCadena = lsCadena & Encabezado("Nº Contratos;20; ;5;Proveedor;30; ;15;Moneda;8;Monto;15; ;6;Desde;10; ;2;Hasta;10; ;2;Cuotas;8;", lnItem)
    
    For lnI = 1 To Me.feCrontratos.Rows - 1
        RSet lsNContrato = Me.feCrontratos.TextMatrix(lnI, 1)
        RSet lsProveedor = Me.feCrontratos.TextMatrix(lnI, 2)
        RSet lsMoneda = Me.feCrontratos.TextMatrix(lnI, 3)
        RSet lsMonto = Me.feCrontratos.TextMatrix(lnI, 4)
        RSet lsDesde = Me.feCrontratos.TextMatrix(lnI, 5)
        RSet lsHasta = Me.feCrontratos.TextMatrix(lnI, 6)
        RSet lsNCuotas = Me.feCrontratos.TextMatrix(lnI, 7)


        lsCadena = lsCadena & Space(1) & Trim(lsNContrato) & Space(22 - Len(Trim(lsNContrato)))
        lsCadena = lsCadena & Space(1) & IIf(Len(Trim(lsProveedor)) < 45, (Trim(lsProveedor) & Space(Abs(45 - Len(Trim(lsProveedor))))), Mid(Trim(lsProveedor), 1, 45))
        lsCadena = lsCadena & oImpre.CentrarCadena(Trim(lsMoneda), 8)
        lsCadena = lsCadena & Space(5) & Trim(lsMonto) & Space(15 - Len(Trim(lsMonto)))
        lsCadena = lsCadena & oImpre.CentrarCadena(Trim(lsDesde), 12)
        lsCadena = lsCadena & oImpre.CentrarCadena(Trim(lsHasta), 12)
        lsCadena = lsCadena & oImpre.CentrarCadena(Trim(lsNCuotas), 8) & oImpresora.gPrnSaltoLinea
        
        If lnItem > 52 Then
            lnItem = 0
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            lsCadena = lsCadena & CabeceraPagina1("C O N T R A T O S     V I G E N T E S", lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "")
            lsCadena = lsCadena & Encabezado("Nº Contratos;20; ;5;Proveedor;30; ;20;Moneda;8;Monto;15; ;6;Desde;13; ;1;Hasta;13; ;3;Cuotas;8;", lnItem)
        End If
        
        lnItem = lnItem + 1
    Next lnI
     
    
    oPrevio.Show lsCadena, "CONTRATOS VIGENTES", True, 66
    Set oPrevio = Nothing
End Sub

Private Sub cmdVer_Click()
Dim row As Long
If ValidarSeleccion Then
    row = feCrontratos.row
    frmLogContSegAdendas.Inicio (Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 1))), (Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 11))), (Trim(feCrontratos.TextMatrix(Me.feCrontratos.row, 12)))
    cmdActualizar_Click
    feCrontratos.row = row
    feCrontratos.TopRow = row
End If
End Sub

Private Sub cmdVerPDF_Click()
If ValidarSeleccion Then
    Dim NContrato As String
    Dim Ruta As String
    Dim oLog As DLogGeneral
    Dim Archivo As New Scripting.FileSystemObject
    
    Set oLog = New DLogGeneral
    NContrato = oLog.NombreArchivoNContrato(Trim(Me.feCrontratos.TextMatrix(Me.feCrontratos.row, 1)), CInt(Me.feCrontratos.TextMatrix(Me.feCrontratos.row, 12))) 'PASI Agrego las Columna 12
    
    If NContrato = "" Then
        MsgBox "Contrato no Cuenta con Archivo Digital.", vbInformation, "Aviso"
    Else
        Ruta = psRutaContrato & Trim(NContrato)
        If Archivo.FileExists(Ruta) = False Then
            MsgBox "Archivo fue eliminado.", vbCritical, "Aviso"
        Else
            ShellExecute Me.hwnd, "open", Ruta, "", "", 4
        End If
    End If
End If
End Sub
Private Function ValidarSeleccion() As Boolean
If Trim(Me.feCrontratos.TextMatrix(1, 1)) = "" Then
    MsgBox "No hay datos.", vbInformation, "Aviso"
    ValidarSeleccion = False
    Exit Function
Else
    If Trim(Me.feCrontratos.TextMatrix(Me.feCrontratos.row, 1)) = "" Then
        MsgBox "Seleccione correctamente el Registro.", vbInformation, "Aviso"
        ValidarSeleccion = False
        Exit Function
    End If
End If
ValidarSeleccion = True
End Function
Private Sub Form_Load()
Dim oConstSist As NConstSistemas

If Trim(Mid(GetMaquinaUsuario, 1, 2)) = "01" Then
    Me.cmdVerPDF.Enabled = True
    'OBTENER RUTA DE CONTRATOS
    Set oConstSist = New NConstSistemas
    psRutaContrato = Trim(oConstSist.LeeConstSistema(gsLogContRutaContratos))
Else
    Me.cmdVerPDF.Enabled = False
    psRutaContrato = ""
End If
Call CargarGrid
End Sub
