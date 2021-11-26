VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogContCronograma 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contratación: Cronograma"
   ClientHeight    =   6435
   ClientLeft      =   2220
   ClientTop       =   1200
   ClientWidth     =   10950
   Icon            =   "frmLogContCronograma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTContratos 
      Height          =   6300
      Left            =   80
      TabIndex        =   0
      Top             =   80
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   11113
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cronograma"
      TabPicture(0)   =   "frmLogContCronograma.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCronograma"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCerrar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdImprimir"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
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
         Left            =   8480
         TabIndex        =   11
         Top             =   5760
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
         Left            =   9600
         TabIndex        =   10
         Top             =   5760
         Width           =   1110
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
         Height          =   1125
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   10560
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
            TabIndex        =   9
            Tag             =   "txtnombre"
            Top             =   720
            Width           =   6495
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
            TabIndex        =   8
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   720
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Contrato:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   270
            Width           =   870
         End
      End
      Begin VB.Frame fraCronograma 
         Caption         =   "Cronograma"
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
         Height          =   4125
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   10560
         Begin Sicmact.FlexEdit feCronograma 
            Height          =   3435
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   10320
            _ExtentX        =   18203
            _ExtentY        =   6059
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Nº Pago-Fecha de Pago-Moneda-Monto-Tipo-Estado"
            EncabezadosAnchos=   "500-1000-1200-1000-1200-3200-2000"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0"
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
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   800
   End
End
Attribute VB_Name = "frmLogContCronograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Integer
Dim fsNContrato As String
Dim fncontRef As Integer 'PASI20140823 TI-ERS077-2014
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Public Sub Inicio(ByVal psNContrato As String, ByVal pnContRef As Integer) 'pnConRef Agregado PASI20140823 TI-ERS077-2014
    fsNContrato = psNContrato
    fncontRef = pnContRef
    If CargarGrid(fsNContrato, fncontRef) Then
        Me.Show 1
    End If
End Sub
Private Function CargarGrid(ByVal psNContrato As String, pnContRef As Integer) As Boolean 'pnConRef Agregado PASI20140823 TI-ERS077-2014
Dim oLog As DLogGeneral
Dim rsLog As ADODB.Recordset

Set oLog = New DLogGeneral
Set rsLog = oLog.ListarCronogramaDeContratos(psNContrato, pnContRef)

Call LimpiaFlex(feCronograma)
If rsLog.RecordCount > 0 Then
    For i = 0 To rsLog.RecordCount - 1
        feCronograma.AdicionaFila
        If i = 0 Then
            Me.lblNContrato.Caption = Space(1) & Trim(rsLog!NContrato)
            Me.lblProveedor.Caption = Space(1) & Trim(rsLog!Proveedor)
        End If
        Me.feCronograma.TextMatrix(i + 1, 0) = i + 1
        Me.feCronograma.TextMatrix(i + 1, 1) = rsLog!NPago
        Me.feCronograma.TextMatrix(i + 1, 2) = Format(rsLog!FechaPago, "dd/mm/yyyy")
        Me.feCronograma.TextMatrix(i + 1, 3) = rsLog!Moneda
        Me.feCronograma.TextMatrix(i + 1, 4) = Format(rsLog!monto, gsFormatoNumeroView)
        Me.feCronograma.TextMatrix(i + 1, 5) = rsLog!Tipo
        Me.feCronograma.TextMatrix(i + 1, 6) = rsLog!Estado
        rsLog.MoveNext
    Next i
    CargarGrid = True
Else
    MsgBox "No hay datos.", vbInformation, "Aviso"
    CargarGrid = False
End If
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
Dim lsNPago As String * 8
Dim lsFechaPago As String * 18
Dim lsMoneda As String * 8
Dim lsMonto As String * 20
Dim lsTipo As String
Dim lsEstado As String

    
Dim oCon As DConecta
Set oCon = New DConecta
    
lsCadena = ""
lsNContrato = Trim(Me.lblNContrato.Caption)
lsProveedor = PstaNombre(Trim(Me.lblProveedor.Caption), False)

lsCadena = lsCadena & oImpresora.gPrnCondensadaON
lsCadena = lsCadena & CabeceraPagina1("C R O N O G R A M A   D E   C O N T R A T O  Nº " & fsNContrato, lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "")

'DATOS
lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
lsCadena = lsCadena & Space(10) & "Nro CONTRATO : " & lsNContrato & oImpresora.gPrnSaltoLinea
lsCadena = lsCadena & Space(10) & "PROVEEDOR    : " & lsProveedor & oImpresora.gPrnSaltoLinea
lsCadena = lsCadena & oImpresora.gPrnSaltoLinea

lsCadena = lsCadena & Space(10) & "DETALLE CRONOGRAMA" & oImpresora.gPrnSaltoLinea
lsCadena = lsCadena & Encabezado("Nº Pago;10; ;5;Fecha de Pago;15; ;5;Moneda;8;Monto;15; ;6;Tipo;8; ;15;Estado;10; ;5;", lnItem)


    For lnI = 1 To feCronograma.Rows - 1
        RSet lsNPago = Me.feCronograma.TextMatrix(lnI, 1)
        RSet lsFechaPago = Me.feCronograma.TextMatrix(lnI, 2)
        RSet lsMoneda = Me.feCronograma.TextMatrix(lnI, 3)
        RSet lsMonto = Me.feCronograma.TextMatrix(lnI, 4)
        lsTipo = Me.feCronograma.TextMatrix(lnI, 5)
        lsEstado = Me.feCronograma.TextMatrix(lnI, 6)

        lsCadena = lsCadena & lsNPago & Space(6)
        lsCadena = lsCadena & oImpre.CentrarCadena(Trim(lsFechaPago), 20)
        lsCadena = lsCadena & oImpre.CentrarCadena(Trim(lsMoneda), 8) 'oImpre.CentrarCadena(Trim(lsMonto), 20)
        lsCadena = lsCadena & IIf(Len(Trim(lsMonto)) > 15, Mid(Trim(lsMonto), 1, 15), Space(Abs(15 - Len(Trim(lsMonto)))) & Trim(lsMonto))
        lsCadena = lsCadena & Space(3) & IIf(Len(Trim(lsTipo)) > 30, Mid(Trim(lsTipo), 1, 30), Trim(lsTipo) & Space(Abs(30 - Len(Trim(lsTipo)))))
        lsCadena = lsCadena & Space(1) & Trim(lsEstado) & oImpresora.gPrnSaltoLinea
        
        If lnItem > 52 Then
            lnItem = 0
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            lsCadena = lsCadena & CabeceraPagina1("C R O N O G R A M A   D E   C O N T R A T O  Nº " & fsNContrato, lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "")
            
            'DATOS
            lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Space(10) & "Nro CONTRATO : " & lsNContrato & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Space(10) & "PROVEEDOR    : " & lsProveedor & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
            
            lsCadena = lsCadena & Space(10) & "DETALLE CRONOGRAMA" & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Encabezado("Nº Pago;10; ;5;Fecha de Pago;15; ;5;Moneda;8;Monto;15; ;6;Tipo;8; ;15;Estado;10; ;5;", lnItem)
        End If
        
        lnItem = lnItem + 1
    Next lnI
     
    
    oPrevio.Show lsCadena, "CRONOGRAMA DE CONTRATO Nº " & lsNContrato, True, 66
    Set oPrevio = Nothing
End Sub

