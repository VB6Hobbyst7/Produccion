VERSION 5.00
Begin VB.Form frmHojaRutaAnalistaDarVisto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vistos Pendientes Para la Hoja de Ruta"
   ClientHeight    =   3870
   ClientLeft      =   10350
   ClientTop       =   5145
   ClientWidth     =   6495
   Icon            =   "frmHojaRutaAnalistaDarVisto.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6495
   Begin VB.CommandButton cmbVisto 
      Caption         =   "Dar Visto"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmbCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame framVistos 
      Caption         =   "Vistos Pendientes"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin SICMACT.FlexEdit flxVistos 
         Height          =   2535
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4471
         Cols0           =   5
         EncabezadosNombres=   "#-Analista Solicita-Fecha Solicitud-Fecha de Atraso-nIdVisto"
         EncabezadosAnchos=   "500-1500-1500-1500-0"
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
         EncabezadosAlineacion=   "C-C-C-C-C"
         FormatosEdit    =   "0-0-5-5-0"
         TextArray0      =   "#"
         ColWidth0       =   495
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmHojaRutaAnalistaDarVisto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsVistos As ADODB.Recordset
Dim dHojaRuta As New DCOMhojaRuta

Private Sub cmbCerrar_Click()
    Unload Me
End Sub

Private Sub cmbVisto_Click()
    Dim cGlosa As String
    cGlosa = InputBox("Ingrese una glosa para el visto:", "Glosa")
    If cGlosa <> "" Then
        Dim nIdVisto As Integer: nIdVisto = CInt(flxVistos.TextMatrix(flxVistos.row, 4))
        dHojaRuta.darVisto nIdVisto, cGlosa, gsCodUser
        flxVistos.EliminaFila flxVistos.row, True
    End If
End Sub

Private Sub Form_Load()
    llenarVistos
End Sub

Private Function llenarVistos()
    Set rsVistos = dHojaRuta.obtenerVistosPendientes(gsCodAge)
    Dim nRow As Integer
    LimpiaFlex flxVistos
    flxVistos.Rows = 2
    Do While Not rsVistos.EOF
        flxVistos.AdicionaFila
        nRow = flxVistos.Rows - 1
        flxVistos.TextMatrix(nRow, 1) = rsVistos!cUserAnalista
        flxVistos.TextMatrix(nRow, 2) = rsVistos!dFechaSolicita
        flxVistos.TextMatrix(nRow, 3) = rsVistos!dFechaAtraso
        flxVistos.TextMatrix(nRow, 4) = rsVistos!nIdVisto
        rsVistos.MoveNext
    Loop
End Function
