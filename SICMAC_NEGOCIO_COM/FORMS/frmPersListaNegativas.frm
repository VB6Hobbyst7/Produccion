VERSION 5.00
Begin VB.Form frmPersListaNegativas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista clientes negativos"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14265
   Icon            =   "frmPersListaNegativas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   14265
   StartUpPosition =   3  'Windows Default
   Begin SICMACT.FlexEdit FePolizas 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   9551
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Tipo Persona-Tipo Doc.-Num.Doc.-Nombre / Razón Social-Justificación-Fuente-Condición-nTipoPers-nTipoDocID"
      EncabezadosAnchos=   "400-1200-1400-1500-3000-2800-2500-1200-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-L-L-L-L-L-L"
      FormatosEdit    =   "0-0-0-0-1-1-1-1-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Para dar mantenimiento a un determinado Cliente, selecciónelo y de doble clic."
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
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmPersListaNegativas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim nTipoOperacion As TipoOperacion

Public sAgencia As String
Public nInmueble As Integer
Public ntna  As Double
Public nPriMinima As Double
Public nDereEmi As Double
Public nalternativa As Integer
Public dfecha  As Date

Public snumdoc As String
Public sNombres As String
Public sJustificacion As String
Public sFuente As String
Public lnTipoDocId As Integer

Public sNumPoliza As String
Public sPersCodContr As String
Dim nEstadoPoliza As Integer

Public Sub Inicio()

    Dim oPers As COMDPersona.DCOMPersonas
    Dim rs As ADODB.Recordset
    
    snumdoc = ""
    lnTipoDocId = 0
    
    
    Set oPers = New COMDPersona.DCOMPersonas
    Set rs = oPers.RecuperaPersListaNegativa()
    Set oPers = Nothing
    
    If (rs.EOF And rs.BOF) Then
        MsgBox "No se encontraron datos.", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    FePolizas.Clear
    FePolizas.FormaCabecera
    FePolizas.rows = 2
    FePolizas.rsFlex = rs
    'FePolizas.SetFocus
    Set oPers = Nothing

    Me.Show 1

End Sub

Private Sub FePolizas_DblClick()
    If FePolizas.rows > 1 Then
    
        If Len(FePolizas.TextMatrix(FePolizas.row, 3)) > 0 Then
            snumdoc = FePolizas.TextMatrix(FePolizas.row, 3)
            sNombres = FePolizas.TextMatrix(FePolizas.row, 4)
            sJustificacion = FePolizas.TextMatrix(FePolizas.row, 5)
            sFuente = FePolizas.TextMatrix(FePolizas.row, 6)
            lnTipoDocId = FePolizas.TextMatrix(FePolizas.row, 8)
        End If
'        sAgencia = FePolizas.TextMatrix(FePolizas.Row, 8)
'        nInmueble = FePolizas.TextMatrix(FePolizas.Row, 9)
'        ntna = FePolizas.TextMatrix(FePolizas.Row, 3)
'        nPriMinima = FePolizas.TextMatrix(FePolizas.Row, 4)
'        nDereEmi = FePolizas.TextMatrix(FePolizas.Row, 5)
'        nalternativa = FePolizas.TextMatrix(FePolizas.Row, 10)
'        dfecha = FePolizas.TextMatrix(FePolizas.Row, 7)
             
        Unload Me
    End If
End Sub

Private Sub FePolizas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call FePolizas_DblClick
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
End Sub


