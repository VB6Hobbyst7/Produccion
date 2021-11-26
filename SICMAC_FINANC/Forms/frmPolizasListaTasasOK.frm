VERSION 5.00
Begin VB.Form frmPolizasListaTasasOK 
   Caption         =   "Porcentajes de Seguros"
   ClientHeight    =   5730
   ClientLeft      =   465
   ClientTop       =   945
   ClientWidth     =   11235
   Icon            =   "frmPolizasListaTasasOK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Porcentaje por agencias y tipo de seguro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11055
      Begin Sicmact.FlexEdit FEGasAge 
         Height          =   4455
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7858
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-CodAge-Desc. Agencia-Cod.Seguro-Desc. Seguro-Porcentaje"
         EncabezadosAnchos=   "400-800-3500-1000-3500-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-5"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-R-L-R"
         FormatosEdit    =   "0-0-0-3-0-2"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      Top             =   5280
      Width           =   1335
   End
End
Attribute VB_Name = "frmPolizasListaTasasOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nPost As Integer
Dim J As Integer
Dim FENoMoverdeFila As Integer
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim obDAgencia As DAgencia
    Set obDAgencia = New DAgencia
    
    nPost = Me.FEGasAge.Rows - 1
    
    If nPost > 0 Then
        For J = 1 To nPost
            If FEGasAge.TextMatrix(J, 1) <> "" Then
                obDAgencia.ActualizarAgenciaPorcenPoliSeguPatri FEGasAge.TextMatrix(J, 1), FEGasAge.TextMatrix(J, 3), FEGasAge.TextMatrix(J, 5)
            End If
        Next J
        MsgBox "Datos se registraron correctamente", vbApplicationModal
        
         Call llenadatos
        
    End If
End Sub

Private Sub FEGasAge_click()
   Call FEGasAge_KeyPress(13)
End Sub

Private Sub FEGasAge_KeyPress(KeyAscii As Integer)
 FENoMoverdeFila = FEGasAge.Row
 FEGasAge.lbEditarFlex = True
End Sub

Private Sub Form_Load()

    Call llenadatos

    
End Sub


Private Sub cmdCargarArch(rs As ADODB.Recordset)
    
    
        Dim I As Integer
        If nPost > 0 Then
            For I = 1 To nPost
                FEGasAge.EliminaFila (1)
            Next I
        End If
        nPost = 0
        If (rs.EOF Or rs.BOF) Then
            MsgBox "No existen porcenctajes de gastos de Agencias"
            Exit Sub
        End If
        rs.MoveFirst
        nPost = 0
        
        Me.FEGasAge.Clear
        
        Me.FEGasAge.rsFlex = rs
        
'        Do While Not (rs.EOF Or rs.BOF)
'            nPost = nPost + 1
'            FEGasAge.AdicionaFila
'            FEGasAge.TextMatrix(nPost, 0) = nPost ''"1"
'            FEGasAge.TextMatrix(nPost, 1) = rs!cAgecod
'            FEGasAge.TextMatrix(nPost, 2) = rs!cAgeDescripcion
'            FEGasAge.TextMatrix(nPost, 3) = rs!nTipoSeguro
'            FEGasAge.TextMatrix(nPost, 4) = rs!cConsDescripcion
'            FEGasAge.TextMatrix(nPost, 5) = rs!nAgePorcentaje
'            rs.MoveNext
'        Loop
        
End Sub

Private Sub llenadatos()
    Dim obDAgencia As DAgencia
    Set obDAgencia = New DAgencia
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = obDAgencia.CargaPorcenPoliSeguro
    Call cmdCargarArch(rs)
End Sub

