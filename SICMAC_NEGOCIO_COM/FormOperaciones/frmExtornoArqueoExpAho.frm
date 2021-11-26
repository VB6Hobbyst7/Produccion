VERSION 5.00
Begin VB.Form frmExtornoArqueoExpAho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extorno de arqueo de expediente de ahorro"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10635
   Icon            =   "frmExtornoArqueoExpAho.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   9240
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "Extornar"
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   10335
      Begin VB.TextBox txtGlosa 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   10095
      End
   End
   Begin VB.Frame fraArqueos 
      Caption         =   "Arqueo(s) del día"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin SICMACT.FlexEdit flxArqueos 
         Height          =   2055
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3625
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Código Arqueo-Usuario Arqueado-Usuario Líder-Fecha-Hora"
         EncabezadosAnchos=   "400-1800-2700-2700-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "R-C-C-C-C-C"
         FormatosEdit    =   "3-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmExtornoArqueoExpAho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ANDE 20170706 ERS021-2017
Option Explicit

Dim lcIdArqueo As String

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdExtornar_Click()
    If Validar Then
        Dim bGuardado As Boolean
        If lcIdArqueo <> "" Then
            bGuardado = GuardarExtorno(lcIdArqueo)
            If bGuardado Then
                MsgBox "El extorno se realizó con éxito.", vbInformation + vbOKOnly, "Aviso"
                Call LimpiarFlxArqueos
            Else
                MsgBox "Error al extornar el arqueo.", vbError + vbOKOnly, "Error"
            End If
            Call CargarArqueosDelDia
        Else
            MsgBox "Debe seleccionar un arqueo.", vbInformation + vbOKOnly, "Aviso"
        End If
    End If
End Sub

Public Function LimpiarFlxArqueos()
    Dim i, nRows As Integer
    nRows = flxArqueos.Rows - 1
    For i = 1 To nRows
        flxArqueos.EliminaFila (i)
    Next i
    
End Function

Private Sub flxArqueos_Click()
    Dim nRow, nCol As Integer
    nRow = flxArqueos.row
    lcIdArqueo = flxArqueos.TextMatrix(nRow, 1)
End Sub

Private Sub Form_Load()
    flxArqueos.Enabled = False
    txtGlosa.Enabled = False
    txtGlosa.BackColor = vbGrayed
    cmdExtornar.Enabled = False
    
    Call CargarArqueosDelDia
    
End Sub

Public Sub CargarArqueosDelDia()
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rArqueos As ADODB.Recordset
    Dim i As Integer
    Set rArqueos = oCaptaLN.ObtenerArqueoDelDia(gdFecSis)
    
    If Not (rArqueos.BOF And rArqueos.EOF) Then
        flxArqueos.Enabled = True
        i = 1
        Do While Not rArqueos.EOF
            flxArqueos.AdicionaFila
            flxArqueos.TextMatrix(i, 1) = rArqueos!cIdArqueo
            flxArqueos.TextMatrix(i, 2) = rArqueos!Arqueado
            flxArqueos.TextMatrix(i, 3) = rArqueos!Arqueador
            flxArqueos.TextMatrix(i, 4) = rArqueos!Fecha
            flxArqueos.TextMatrix(i, 5) = rArqueos!Hora
            i = i + 1
            rArqueos.MoveNext
        Loop
        
        txtGlosa.Enabled = True
        txtGlosa.BackColor = vbWhite
        cmdExtornar.Enabled = True
    Else
        MsgBox "No se iniciaron arqueos el día de hoy.", vbInformation + vbOKOnly, "Aviso"
    End If
    
End Sub

Public Function Validar() As Boolean
    'validando glosa
    Validar = True
    If txtGlosa.Text = "" Then
        MsgBox "Por favor ingrese una glosa, la glosa es necesario para extornar.", vbInformation + vbOKOnly, "Aviso"
        txtGlosa.SetFocus
        Validar = False
        Exit Function
    End If
End Function

Public Function GuardarExtorno(ByVal pcIdArqueo As String) As Boolean
    Dim oCaptaLN As New COMDCaptaGenerales.DCOMCaptaGenerales
    On Error GoTo ErrorAlGuardarExtorno
    
    Call oCaptaLN.ExtornarArqueoExpAho(pcIdArqueo)
    GuardarExtorno = True
    Set oCaptaLN = Nothing
    Exit Function
ErrorAlGuardarExtorno:
    GuardarExtorno = False
End Function
'END ANDE ERS021-2101
