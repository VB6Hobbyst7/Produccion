VERSION 5.00
Begin VB.Form frmRHAreas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Areas"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmRHAreas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraComite 
      Caption         =   "Comite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4260
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6810
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   5745
         TabIndex        =   4
         Top             =   3810
         Width           =   975
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "N&uevo"
         Height          =   375
         Left            =   4710
         TabIndex        =   3
         Top             =   3810
         Width           =   975
      End
      Begin SicmactAdmin.FlexEdit lFlexEdit 
         Height          =   3540
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6630
         _ExtentX        =   16828
         _ExtentY        =   6244
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Area"
         EncabezadosAnchos=   "350-1600-4000"
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
         ColumnasAEditar =   "X-1-X"
         ListaControles  =   "0-1-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   240
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   4305
      Width           =   975
   End
End
Attribute VB_Name = "frmRHAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbEditado As Boolean
Dim lsCodigo As String
Dim lsEvaCod As String
Dim rs As ADODB.Recordset
Dim rsIni As ADODB.Recordset
Dim lbIni As Boolean

Public Function Ini(psEvalCod As String, rsI As ADODB.Recordset) As ADODB.Recordset
    lsEvaCod = psEvalCod
    lbIni = True
    Set rsIni = rsI
    Show 1
    Set Ini = rs
End Function


Private Sub cmdEliminar_Click()
    Me.lFlexEdit.EliminaFila Me.lFlexEdit.Row
End Sub

Private Sub cmdNuevo_Click()
    Dim oPersona As UPersona
    If Me.lFlexEdit.TextMatrix(Me.lFlexEdit.Rows - 1, 0) = "#" Then
        lFlexEdit.Rows = 2
    End If
    
    If Me.lFlexEdit.TextMatrix(Me.lFlexEdit.Rows - 1, 0) = "" Then
        lFlexEdit.AdicionaFila 1
    Else
        lFlexEdit.AdicionaFila CLng(Me.lFlexEdit.TextMatrix(lFlexEdit.Rows - 1, 0)) + 1
    End If
    
    Me.lFlexEdit.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsE As ADODB.Recordset
    Dim oCons As DConstantes
    
    Set oCons = New DConstantes
    Set rsE = New ADODB.Recordset
    Set rsE = oCons.GetConstante(gRHEvaluacionComite)
    Me.lFlexEdit.CargaCombo rsE
    Set oCons = Nothing
    If Not lbIni Then
        Activa False
    Else
        Set Me.lFlexEdit.Recordset = rsIni
        Me.fraComite.Enabled = True
    End If
    
    Dim oDatoAreas As DActualizaDatosArea
    Set oDatoAreas = New DActualizaDatosArea
    
    Me.lFlexEdit.rsTextBuscar = oDatoAreas.GetAreas
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.fraComite.Enabled = pbValor
    Me.cmdSalir.Enabled = Not pbValor
    If Not lbIni Then Me.cmdGrabar.Enabled = pbValor
    Me.cmdCancelar.Visible = pbValor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Me.lFlexEdit.GetRsNew
End Sub

