VERSION 5.00
Begin VB.Form frmCredEvalExtornoVerif 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno de Verificación"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   Icon            =   "frmCredEvalExtornoVerif.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar Verificación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   2010
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7920
      TabIndex        =   4
      Top             =   4080
      Width           =   1170
   End
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "&Mostrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1170
   End
   Begin VB.ComboBox cmbAgencia 
      Height          =   315
      ItemData        =   "frmCredEvalExtornoVerif.frx":030A
      Left            =   960
      List            =   "frmCredEvalExtornoVerif.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin SICMACT.FlexEdit feCredVerif 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   5318
      Cols0           =   4
      HighLight       =   1
      EncabezadosNombres=   "-Crédito-Titular-Fecha Verificación"
      EncabezadosAnchos=   "300-2500-3500-2000"
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X"
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C"
      FormatosEdit    =   "0-0-0-0"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   3
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
   End
   Begin VB.Label lblAgencia 
      AutoSize        =   -1  'True
      Caption         =   "Agencia:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "frmCredEvalExtornoVerif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsCodAgencias, fsCtaCod As String
Dim objPista As COMManejador.Pista

Private Sub Cargar_Objetos_Controles()
Dim oAgencia  As COMDConstantes.DCOMAgencias
Dim rsAgencias As ADODB.Recordset
Set oAgencia = New COMDConstantes.DCOMAgencias
Set rsAgencias = oAgencia.ObtieneAgencias()
Call Llenar_Combo_con_Recordset(rsAgencias, cmbAgencia)
End Sub

Private Sub cmbAgencia_Click()
fsCodAgencias = Trim(Right(Me.cmbAgencia.Text, 4))
fsCodAgencias = IIf(Len(fsCodAgencias) < 2, "0" & fsCodAgencias, fsCodAgencias)
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdExtornar_Click()
fsCtaCod = Trim(feCredVerif.TextMatrix(feCredVerif.Row, 1))
If MsgBox("Estas seguro de Extornar la Verificación del Crédito Nº " & fsCtaCod & "?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    Dim oCredito As COMDCredito.DCOMCredito
    Set oCredito = New COMDCredito.DCOMCredito
    
    Call oCredito.ExtornarVerificacion(fsCtaCod)
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gCredExtornoVerificacionEvaluacionCred, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Extorno de Formato de Evaluacón de Credito ", fsCtaCod, gCodigoCuenta
    
    Call LlenarGrid
    
    MsgBox "Verificación Extornada con exito.", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdMostrar_Click()
If Trim(Me.cmbAgencia.Text) = "" Then
    MsgBox "Seleccione Agencia.", vbInformation, "Aviso"
Else
    Call LlenarGrid
End If
End Sub

Private Sub Form_Load()
Call Cargar_Objetos_Controles
End Sub

Private Sub LlenarGrid()
Dim oCredito As COMDCredito.DCOMCredito
    Dim rsCredito As ADODB.Recordset
    Dim i As Integer
    Set oCredito = New COMDCredito.DCOMCredito
    Set rsCredito = oCredito.RecuperaCredVerif(fsCodAgencias)
    Call LimpiaFlex(feCredVerif)
    If rsCredito.RecordCount > 0 Then
        If Not (rsCredito.EOF And rsCredito.BOF) Then
            For i = 0 To rsCredito.RecordCount - 1
                feCredVerif.AdicionaFila
                feCredVerif.TextMatrix(i + 1, 0) = i + 1
                feCredVerif.TextMatrix(i + 1, 1) = rsCredito!cCtaCod
                feCredVerif.TextMatrix(i + 1, 2) = Trim(rsCredito!Titular)
                feCredVerif.TextMatrix(i + 1, 3) = rsCredito!Fecha
                rsCredito.MoveNext
            Next i
        End If
    Else
        MsgBox "No hay datos.", vbInformation, "Aviso"
    End If
End Sub
