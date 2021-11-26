VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAudDesarrolloProcedimientoVerificar 
   Caption         =   "AUDITORIA:DESARROLLO DE PROCEDIMIENTOS"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   Icon            =   "frmAudDesarrolloProcedimientoVerificar.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   6345
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Procedimientos :"
      Height          =   5415
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   10935
      Begin VB.CommandButton cmdHistorial 
         Caption         =   "Ver Historial"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   4920
         Width           =   1455
      End
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "Verificar"
         Height          =   315
         Left            =   8640
         TabIndex        =   4
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   315
         Left            =   9720
         TabIndex        =   3
         Top             =   4920
         Width           =   975
      End
      Begin SICMACT.FlexEdit grdProcedimientos 
         Height          =   4335
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7646
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Actividad-Procedimiento-Descripción-Usuario-Fecha-ID"
         EncabezadosAnchos=   "450-2000-2000-3500-1500-1200-450"
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
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11245
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Desarrollo de Procedimientos"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAudDesarrolloProcedimientoVerificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public nProcedimientoID As Integer
'
'Private Sub cmdHistorial_Click()
'    If grdProcedimientos.TextMatrix(grdProcedimientos.row, 6) <> "" Then
'        frmAudHistorialSolicitudVal.pnProcedimientoID = grdProcedimientos.TextMatrix(grdProcedimientos.row, 6)
'        frmAudHistorialSolicitudVal.Show 1
'    Else
'        MsgBox "Dato vacio. Debe seleccionar una opción válida", vbCritical, "Aviso"
'    End If
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdVerificar_Click()
'    If grdProcedimientos.TextMatrix(grdProcedimientos.row, 6) <> "" Then
'        nProcedimientoID = grdProcedimientos.TextMatrix(grdProcedimientos.row, 6)
'        frmAudVerificarProcedimiento.Show 1
'        CargarGrilla
'    Else
'        MsgBox "Dato vacio. Seleccione opción válida.", vbCritical, "Aviso"
'    End If
'End Sub
'
'Public Sub CargarGrilla()
'    grdProcedimientos.Clear
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'    grdProcedimientos.FormaCabecera
'    grdProcedimientos.rsFlex = objCOMNAuditoria.ListarProcedimientoVerificacion
'End Sub
'
'Private Sub Form_Load()
'    CargarGrilla
'End Sub
