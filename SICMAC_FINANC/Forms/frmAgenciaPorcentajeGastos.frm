VERSION 5.00
Begin VB.Form frmAgenciaPorcentajeGastos 
   Caption         =   "Porcentaje de depreciación"
   ClientHeight    =   5730
   ClientLeft      =   465
   ClientTop       =   945
   ClientWidth     =   9810
   Icon            =   "frmAgenciaPorcentajeGastos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Porcentaje de gastos  por agencias"
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
      Width           =   9615
      Begin Sicmact.FlexEdit FEGasAge 
         Height          =   4455
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7858
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-CodAge-DescAge-Porcentaje-% Cartera Aho.-% Ing. Financ."
         EncabezadosAnchos=   "400-1200-4000-1200-1200-1200"
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
         ColumnasAEditar =   "X-X-X-3-4-5"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-R-R-R"
         FormatosEdit    =   "0-0-0-2-2-2"
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
Attribute VB_Name = "frmAgenciaPorcentajeGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nPost As Integer
Dim j As Integer
Dim FENoMoverdeFila As Integer
'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim obDAgencia As DAgencia
    Set obDAgencia = New DAgencia
    
    Dim nDescuadreEncontrado As Integer, i As Integer, j As Integer, K As Integer
    Dim nSumPorcen As Currency
    
    If nPost > 0 Then
    
    '*********************** BEGIN valida los porcentajes al 100% - PEAC 20100708
    '-- 1 creditos 2 ahorros 3 financiero
    nDescuadreEncontrado = 0
    For i = 3 To 5
        nSumPorcen = 0
        For K = 1 To nPost
            nSumPorcen = nSumPorcen + FEGasAge.TextMatrix(K, i)
        Next K
            If nSumPorcen <> 100 Then
                MsgBox "Los porcentajes de la columna ''" + FEGasAge.TextMatrix(0, i) + "'' no dan al 100%, por favor revise. Suman " + Trim(CStr(nSumPorcen)) + " Dif. " + Trim(CStr(Abs(100 - nSumPorcen))), vbOKOnly + vbExclamation, "Atencion"
                nDescuadreEncontrado = 1
            End If
    Next i
    If nDescuadreEncontrado = 1 Then Exit Sub
    '*********************** END valida los porcentajes al 100%
        
    For j = 1 To nPost
        
        '*** PEAC 20100708
        'obDAgencia.ActualizarAgenciaPorcentajeGastos FEGasAge.TextMatrix(J, 1), FEGasAge.TextMatrix(J, 3)
        obDAgencia.ActualizarAgenciaPorcentajeGastos FEGasAge.TextMatrix(j, 1), FEGasAge.TextMatrix(j, 3), FEGasAge.TextMatrix(j, 4), FEGasAge.TextMatrix(j, 5)
        
    Next j
    MsgBox "Datos se registraron correctamente", vbApplicationModal
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = ""
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo Correctamente "
            Set objPista = Nothing
            '*******
    End If
End Sub

Private Sub FEGasAge_click()
   Call FEGasAge_KeyPress(13)
End Sub

Private Sub FEGasAge_KeyPress(KeyAscii As Integer)
 FENoMoverdeFila = FEGasAge.row
 FEGasAge.lbEditarFlex = True
End Sub

Private Sub Form_Load()
    Dim obDAgencia As DAgencia
    Set obDAgencia = New DAgencia
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = obDAgencia.GetAgenciaPorcentajeGastos
    Call cmdCargarArch(rs)
End Sub


Private Sub cmdCargarArch(rs As ADODB.Recordset)
    
    
        Dim i As Integer
        If nPost > 0 Then
            For i = 1 To nPost
                FEGasAge.EliminaFila (1)
            Next i
        End If
        nPost = 0
        If (rs.EOF Or rs.BOF) Then
            MsgBox "No existen porcenctajes de gastos de Agencias"
            Exit Sub
        End If
        rs.MoveFirst
'        If (rs.EOF Or rs.BOF) Then
'            MsgBox "No existen porcenctajes de gastos de Agencias"
'            Exit Sub
'        End If
        nPost = 0
        Do While Not (rs.EOF Or rs.BOF)
            nPost = nPost + 1
            FEGasAge.AdicionaFila
            FEGasAge.TextMatrix(nPost, 0) = "1"
            FEGasAge.TextMatrix(nPost, 1) = rs!cAgeCod
            FEGasAge.TextMatrix(nPost, 2) = rs!cAgeDescripcion
            FEGasAge.TextMatrix(nPost, 3) = rs!nAgePorcentaje
            FEGasAge.TextMatrix(nPost, 4) = rs!nPorcenCarteraAho '*** PEAC 20100708
            FEGasAge.TextMatrix(nPost, 5) = rs!nPorcenIngFinan '*** PEAC 20100708
            rs.MoveNext
        Loop
        
End Sub
