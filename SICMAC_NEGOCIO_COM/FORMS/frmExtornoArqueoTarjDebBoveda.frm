VERSION 5.00
Begin VB.Form frmExtornoArqueoTarjDebBoveda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "frmExtornoArqueoTarjDebBoveda.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   9615
      Begin VB.TextBox txtGlosa 
         Height          =   405
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9375
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   7150
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin SICMACT.FlexEdit feExtorno 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4471
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Usuario Arqueado-Usuario Superviza-Fecha-nIdArqueo"
      EncabezadosAnchos=   "300-1500-1500-1200-0"
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
      ColumnasAEditar =   "X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmExtornoArqueoTarjDebBoveda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'** Nombre : frmExtornoArqueoTarjDebBoveda
'** Descripción : Formulario para Extornar el Arqueo de Stock de Tarjetas de Debito  - Boveda o Ventanilla
'** Creación : PASI, 20151221
'** Referencia : TI-ERS069-2015
'***************************************************************************
Dim bResultadoVisto As Boolean
Dim oVisto As frmVistoElectronico
Dim cUsuVisto As String
Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
Dim gTpoArqueo As Integer
Public Sub Inicia(ByVal pTpoArqueo As Integer)
    gTpoArqueo = pTpoArqueo
    Set oVisto = New frmVistoElectronico
    bResultadoVisto = oVisto.Inicio(15)
    If Not bResultadoVisto Then
        Exit Sub
    End If
    cUsuVisto = oVisto.ObtieneUsuarioVisto
    Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
    If feExtorno.TextMatrix(feExtorno.row, 1) = "" Then
        MsgBox "No Existen Arqueos a Extornar", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ud. Debe Ingresar la Glosa del Extorno", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("¿Esta seguro de extornar el Arqueo?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    oCaja.ExtornaArqueoTarjDeb feExtorno.TextMatrix(feExtorno.row, 4), cUsuVisto, gdFecSis, Trim(Replace(Replace((txtGlosa.Text), Chr(10), ""), Chr(13), ""))
    MsgBox "Se ha extornado con éxito el arqueo.", vbInformation, "Aviso"
    Unload Me
End Sub
Private Sub Form_Load()
    CargaDatos
    If gTpoArqueo = 0 Then
        Me.Caption = "Extorno de Arqueo de Bóveda"
    Else
        Me.Caption = "Extorno de Arqueo de Ventanilla"
    End If
End Sub
Private Sub CargaDatos()
    Dim rs As ADODB.Recordset
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    If gTpoArqueo = 0 Then
        Set rs = oCaja.ObtieneArqueoBovxExtorno(gdFecSis, gsCodUser)
    Else
        Set rs = oCaja.ObtieneArqueoVentxExtorno(gdFecSis, gsCodUser)
    End If
    Do While Not rs.EOF
        feExtorno.AdicionaFila
        feExtorno.TextMatrix(feExtorno.row, 1) = rs!cUserArqueado
        feExtorno.TextMatrix(feExtorno.row, 2) = rs!cUserSuperviza
        feExtorno.TextMatrix(feExtorno.row, 3) = rs!dFecha
        feExtorno.TextMatrix(feExtorno.row, 4) = rs!nIdArqueo
        rs.MoveNext
    Loop
End Sub
