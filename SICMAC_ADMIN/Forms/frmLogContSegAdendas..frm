VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogContSegAdendas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contratación: Consulta de Adendas"
   ClientHeight    =   7635
   ClientLeft      =   2220
   ClientTop       =   1200
   ClientWidth     =   11205
   Icon            =   "frmLogContSegAdendas..frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTContratos 
      Height          =   7500
      Left            =   80
      TabIndex        =   0
      Top             =   80
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   13229
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lista de Adendas"
      TabPicture(0)   =   "frmLogContSegAdendas..frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "feAdendas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCerrar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdImprimir"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdExtornar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdVerDetalle"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdActualizar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
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
         Left            =   8400
         TabIndex        =   14
         Top             =   6960
         Width           =   1485
      End
      Begin VB.CommandButton cmdVerDetalle 
         Caption         =   "&Ver Detalle"
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
         Left            =   3720
         TabIndex        =   13
         Top             =   6960
         Width           =   1485
      End
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
         Left            =   5280
         TabIndex        =   12
         Top             =   6960
         Width           =   1485
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
         Left            =   6840
         TabIndex        =   11
         Top             =   6960
         Width           =   1485
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
         Left            =   9960
         TabIndex        =   10
         Top             =   6960
         Width           =   885
      End
      Begin VB.Frame Frame3 
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
         Height          =   1005
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   10560
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Contrato:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   270
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   780
         End
         Begin VB.Label lblNContrato 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   7
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblProveedor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   6
            Tag             =   "txtnombre"
            Top             =   600
            Width           =   6975
         End
      End
      Begin Sicmact.FlexEdit feAdendas 
         Height          =   5115
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   10560
         _ExtentX        =   18627
         _ExtentY        =   9022
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nº Adenda-Tipo-Glosa-NTipo-NEstado"
         EncabezadosAnchos=   "500-1200-2000-6000-0-0"
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
         EncabezadosAlineacion=   "C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   7
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
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
Attribute VB_Name = "frmLogContSegAdendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim fsNContrato As String
Dim fnContRef As Integer  'Pasi20140823 ti-ers077-2014
Dim fdFecIni As Date
Dim fdFecFin As Date
Dim i As Integer
Dim fntpodocorigen As Integer
Private Sub cmdActualizar_Click()
If CargaDatos Then
    
End If
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub
Public Sub Inicio(ByVal psNContrato As String, Optional ByVal pnTipoDocOrigen As Integer = 0, Optional ByVal pnContRef As Integer = 0) 'pnContRef agregado  Pasi20140823 ti-ers077-2014
fsNContrato = psNContrato
fnContRef = pnContRef
fntpodocorigen = pnTipoDocOrigen
If CargaDatos Then
    Me.Show 1
End If
End Sub

Private Sub cmdExtornar_Click()
Dim lnNAdenda As Integer
Dim lnTpoAdenda As Integer
Dim row As Long
If ValidarSeleccion Then
    'Call frmLogContExtAdendas.Inicio(Trim(lblNContrato.Caption), Trim(feAdendas.TextMatrix(Me.feAdendas.row, 1)), CInt(Trim(feAdendas.TextMatrix(Me.feAdendas.row, 4))), CInt(Trim(feAdendas.TextMatrix(Me.feAdendas.row, 5))))
    row = feAdendas.row
    lnNAdenda = CInt(Trim(feAdendas.TextMatrix(Me.feAdendas.row, 1)))
    lnTpoAdenda = CInt(Trim(feAdendas.TextMatrix(Me.feAdendas.row, 4)))
    If lnNAdenda <= 0 Then
        MsgBox "El Presente contrato no cuenta con Adendas", vbInformation, "Aviso"
        Exit Sub
    End If
    frmLogContExtAdendas.Inicio Trim(lblNContrato.Caption), lnNAdenda, lnTpoAdenda, fntpodocorigen, fnContRef 'EJVG20131204 'fncontRef agregado pasi20140823 ti-ers077-2014
    cmdActualizar_Click
    feAdendas.row = row
    feAdendas.TopRow = row
End If
End Sub
Private Function ValidarSeleccion() As Boolean
If Trim(Me.feAdendas.TextMatrix(1, 1)) = "" Then
    MsgBox "No hay datos.", vbInformation, "Aviso"
    ValidarSeleccion = False
    Exit Function
Else
    If Trim(Me.feAdendas.TextMatrix(Me.feAdendas.row, 1)) = "" Then
        MsgBox "Seleccione correctamente el Registro.", vbInformation, "Aviso"
        ValidarSeleccion = False
        Exit Function
    End If
End If
ValidarSeleccion = True
End Function

Private Sub cmdImprimir_Click()
Dim oImpre As New COMFunciones.FCOMImpresion


Dim lsCadena As String
Dim lnPagina As Long
Dim lnItem As Long
Dim lnI As Long
Dim oPrevio As clsPrevio
    
    
Set oPrevio = New clsPrevio
    
Dim lsNContrato As String
Dim lsProveedor As String
Dim lsNAdenda As String * 4
Dim lsTipo As String
Dim lsGlosa As String

    
Dim oCon As DConecta
Set oCon = New DConecta
    
lsCadena = ""
lsNContrato = Trim(Me.lblNContrato.Caption)
lsProveedor = PstaNombre(Trim(Me.lblProveedor.Caption), False)

lsCadena = lsCadena & oImpresora.gPrnCondensadaON
lsCadena = lsCadena & CabeceraPagina1("A D E N D A S   A L   C O N T R A T O  Nº " & fsNContrato, lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "")

'DATOS
lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
lsCadena = lsCadena & Space(10) & "Nro CONTRATO : " & lsNContrato & oImpresora.gPrnSaltoLinea
lsCadena = lsCadena & Space(10) & "PROVEEDOR    : " & lsProveedor & oImpresora.gPrnSaltoLinea
lsCadena = lsCadena & oImpresora.gPrnSaltoLinea

lsCadena = lsCadena & Space(6) & "LISTA DE ADENDAS" & oImpresora.gPrnSaltoLinea
lsCadena = lsCadena & Encabezado1("Nº Adenda;10; ;5;Tipo;8; ;5;Glosa;20; ;80;", lnItem)


    For lnI = 1 To feAdendas.Rows - 1
        RSet lsNAdenda = Me.feAdendas.TextMatrix(lnI, 1)
         lsTipo = Me.feAdendas.TextMatrix(lnI, 2)
         lsGlosa = Me.feAdendas.TextMatrix(lnI, 3)

        lsCadena = lsCadena & Space(6) & lsNAdenda
        lsCadena = lsCadena & oImpre.CentrarCadena(Trim(lsTipo), 30)
        lsCadena = lsCadena & Mid(Trim(lsGlosa), 1, 200) & oImpresora.gPrnSaltoLinea

        
        If lnItem > 52 Then
            lnItem = 0
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            lsCadena = lsCadena & oImpresora.gPrnCondensadaON
            lsCadena = lsCadena & CabeceraPagina1("A D E N D A S   A L   C O N T R A T O  Nº " & fsNContrato, lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "")
            
            'DATOS
            lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Space(10) & "Nro CONTRATO : " & lsNContrato & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Space(10) & "PROVEEDOR    : " & lsProveedor & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
            
            lsCadena = lsCadena & Space(6) & "LISTA DE ADENDAS" & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Encabezado1("Nº Adenda;10; ;5;Tipo;8; ;5;Glosa;20; ;80;", lnItem)
        End If
        
        lnItem = lnItem + 1
    Next lnI
     
    
    oPrevio.Show lsCadena, "ADENDAS AL CONTRATO Nº " & lsNContrato, True, 66
    Set oPrevio = Nothing
End Sub
Private Sub cmdVerDetalle_Click()
If ValidarSeleccion Then
    If fntpodocorigen = 0 Then
        Call frmLogContRegAdendas.Inicio(Trim(Me.lblNContrato.Caption), CInt(Trim(feAdendas.TextMatrix(feAdendas.row, 1))), CInt(Trim(feAdendas.TextMatrix(feAdendas.row, 4)))) 'fncontRef agregado pasi20140823 ti-ers077-2014
    ElseIf fntpodocorigen = LogTipoContrato.ContratoServicio Or fntpodocorigen = LogTipoContrato.ContratoArrendamiento Then
        Call frmLogContRegAdendas.Inicio(Trim(Me.lblNContrato.Caption), CInt(Trim(feAdendas.TextMatrix(feAdendas.row, 1))), CInt(Trim(feAdendas.TextMatrix(feAdendas.row, 4))), fntpodocorigen, fnContRef) 'fncontRef agregado pasi20140823 ti-ers077-2014
    Else
        Call frmLogContReajAdenda.Inicio(Trim(lblNContrato.Caption), fnContRef, fntpodocorigen, CInt(Trim(feAdendas.TextMatrix(feAdendas.row, 1))))
    End If
End If
End Sub
Public Function CargaDatos() As Boolean
Dim oLog As DLogGeneral
Dim rsLog As ADODB.Recordset

Set oLog = New DLogGeneral
Set rsLog = oLog.ListarDatosContratos(fsNContrato, fnContRef) 'fncontRef agregado pasi20140823 ti-ers077-2014

If rsLog.RecordCount > 0 Then
    'DATOS
    Me.lblNContrato.Caption = Space(1) & rsLog!NContrato
    Me.lblProveedor.Caption = Space(1) & rsLog!Proveedor

    fdFecIni = CDate(rsLog!Desde)
    fdFecFin = CDate(rsLog!Hasta)
    
    'ADENDAS
    Set rsLog = oLog.ListarDatosAdendasPorContrato(fsNContrato, fnContRef) 'fncontRef agregado pasi20140823 ti-ers077-2014
     Call LimpiaFlex(feAdendas)
     If rsLog.RecordCount > 0 Then
        For i = 0 To rsLog.RecordCount - 1
            feAdendas.AdicionaFila
            feAdendas.TextMatrix(i + 1, 0) = i + 1
            feAdendas.TextMatrix(i + 1, 1) = Trim(rsLog!cNAdenda)
            feAdendas.TextMatrix(i + 1, 2) = Trim(rsLog!TpoDesc)
            feAdendas.TextMatrix(i + 1, 3) = Trim(rsLog!cGlosa)
            feAdendas.TextMatrix(i + 1, 4) = Trim(rsLog!nTipo)
            feAdendas.TextMatrix(i + 1, 5) = Trim(rsLog!nEstado)
            rsLog.MoveNext
        Next i
        CargaDatos = True
    Else
        MsgBox "Contrato no cuenta con Adendas", vbInformation, "Aviso"
        CargaDatos = False
    End If
Else
    MsgBox "Contrato no cuenta con Adendas", vbInformation, "Aviso"
    CargaDatos = False
End If

End Function

