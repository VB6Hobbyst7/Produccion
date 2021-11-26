VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOCOcupacionCIIU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ocupacion / CIIU"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   Icon            =   "frmOCOcupacionCIIU.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
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
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ocupación"
      TabPicture(0)   =   "frmOCOcupacionCIIU.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feOcupacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "CIIU"
      TabPicture(1)   =   "frmOCOcupacionCIIU.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "feCIIU"
      Tab(1).ControlCount=   1
      Begin SICMACT.FlexEdit feOcupacion 
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6376
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Ocupación-Total"
         EncabezadosAnchos=   "0-7000-700"
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
         ColumnasAEditar =   "X-X-X"
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit feCIIU 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6376
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-CIIU-Total"
         EncabezadosAnchos=   "0-7000-700"
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
         ColumnasAEditar =   "X-X-X"
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmOCOcupacionCIIU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************
'*** REQUERIMIENTO: TI-ERS106-2014 Y ANEXO-01
'*** USUARIO: FRHU
'*** FECHA CREACION: 17/09/2014
'********************************************
Option Explicit
Private Sub Form_Load()
    Call CargarOcupaciones
    Call CargarCIIU
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub CargarOcupaciones()
    Dim oConstante As New COMDConstantes.DCOMConstantes
    Dim fila As Integer
    Dim rs As ADODB.Recordset
    Set rs = oConstante.GetOficCumpTotalPersonasPorOcupacion
    
    fila = 0
    Do While Not rs.EOF
        fila = fila + 1
        feOcupacion.AdicionaFila
        feOcupacion.TextMatrix(fila, 1) = rs!Descripcion
        feOcupacion.TextMatrix(fila, 2) = rs!Total
        rs.MoveNext
    Loop
End Sub
Private Sub CargarCIIU()
    Dim oConstante As New COMDConstantes.DCOMConstantes
    Dim fila As Integer
    Dim rs As ADODB.Recordset
    Set rs = oConstante.GetOficCumpTotalPersonasPorCIIU
    
    fila = 0
    Do While Not rs.EOF
        fila = fila + 1
        feCIIU.AdicionaFila
        feCIIU.TextMatrix(fila, 1) = rs!cCIIUdescripcion
        feCIIU.TextMatrix(fila, 2) = rs!Total
        rs.MoveNext
    Loop
End Sub
