VERSION 5.00
Begin VB.Form frmLogSelParaMant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Parámetros"
   ClientHeight    =   4080
   ClientLeft      =   1575
   ClientTop       =   2340
   ClientWidth     =   7095
   Icon            =   "frmLogSelParaMant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optParametro 
      Caption         =   "Económico"
      Height          =   330
      Index           =   1
      Left            =   3450
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   525
      Width           =   1320
   End
   Begin VB.OptionButton optParametro 
      Caption         =   "Técnico"
      Height          =   330
      Index           =   0
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   525
      Value           =   -1  'True
      Width           =   1320
   End
   Begin VB.CommandButton cmdPara 
      Caption         =   "Agregar"
      Height          =   390
      Index           =   0
      Left            =   2070
      TabIndex        =   6
      Top             =   3600
      Width           =   1230
   End
   Begin VB.CommandButton cmdPara 
      Caption         =   "Modificar"
      Height          =   390
      Index           =   1
      Left            =   3630
      TabIndex        =   5
      Top             =   3600
      Width           =   1230
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   5415
      TabIndex        =   0
      Top             =   3600
      Width           =   1230
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   0
      Top             =   3585
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin Sicmact.FlexEdit fgeParametro 
      Height          =   2415
      Left            =   225
      TabIndex        =   3
      Top             =   1020
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   4260
      Cols0           =   3
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Codigo-Descripción"
      EncabezadosAnchos=   "500-0-5500"
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
      ColumnasAEditar =   "X-X-X"
      ListaControles  =   "0-0-0"
      EncabezadosAlineacion=   "C-L-L"
      FormatosEdit    =   "0-0-0"
      CantDecimales   =   0
      AvanceCeldas    =   1
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   495
      RowHeight0      =   285
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Parámetros :"
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
      Height          =   210
      Index           =   1
      Left            =   375
      TabIndex        =   4
      Top             =   585
      Width           =   1185
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1260
      TabIndex        =   2
      Top             =   90
      Width           =   3435
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Area :"
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
      Height          =   210
      Index           =   0
      Left            =   420
      TabIndex        =   1
      Top             =   135
      Width           =   555
   End
End
Attribute VB_Name = "frmLogSelParaMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPara_Click(Index As Integer)
    Dim sCodigo As String, sDescripcion As String
    Dim nResult As Integer
    Dim nIndOptPar As Integer
    Dim sCodSec As Integer
    'Carga el indice del OPTPARAMETRO
    nIndOptPar = IIf(optParametro(0).Value = True, 0, 1)
    sCodSec = IIf(nIndOptPar = 0, gLogSelParTec, gLogSelParEco)
    Select Case Index
        Case 0:
            'Agregar
            nResult = frmLogMantOpc.Inicio("2", "1", , sCodSec)
            If nResult = 0 Then
                Call optParametro_Click(nIndOptPar)
            End If
        Case 1:
            'Modificar
            sCodigo = fgeParametro.TextMatrix(fgeParametro.Row, 1)
            If sCodigo = "" Then
                MsgBox "Falta identificar el código", vbInformation, " Aviso"
                Exit Sub
            End If
            nResult = frmLogMantOpc.Inicio("2", "2", sCodigo, sCodSec)
            If nResult = 0 Then
                Call optParametro_Click(nIndOptPar)
            End If
        Case Else
            MsgBox "Indice de Comando no reconocido", vbInformation, " Aviso "
    End Select
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    lblAreaDes.Caption = Usuario.AreaNom
    
    Call optParametro_Click(0)
End Sub

Private Sub optParametro_Click(Index As Integer)
    Dim clsDGnral As DLogGeneral
    Dim rs As ADODB.Recordset
    If Index > 1 Then
        MsgBox "Opción no reconocida", vbInformation, " Aviso "
        Exit Sub
    End If
    'Limpiar
    fgeParametro.Clear
    fgeParametro.FormaCabecera
    fgeParametro.Rows = 2
    'Carga Parametros
    Set clsDGnral = New DLogGeneral
    Set rs = New ADODB.Recordset
    If Index = 0 Then
        Set rs = clsDGnral.CargaConstante(gLogSelParTec)
        If rs.RecordCount > 0 Then Set fgeParametro.Recordset = rs
    ElseIf Index = 1 Then
        Set rs = clsDGnral.CargaConstante(gLogSelParEco)
        If rs.RecordCount > 0 Then Set fgeParametro.Recordset = rs
    End If
    Set rs = Nothing
    Set clsDGnral = Nothing
End Sub

