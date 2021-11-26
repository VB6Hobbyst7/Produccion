VERSION 5.00
Begin VB.Form frmChequeEditSel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecciona Cheque"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11985
   Icon            =   "frmChequeEditSel.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "&Actualizar"
      Height          =   345
      Left            =   8690
      TabIndex        =   3
      Top             =   4630
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   10810
      TabIndex        =   1
      Top             =   4630
      Width           =   1050
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   345
      Left            =   9750
      TabIndex        =   0
      Top             =   4630
      Width           =   1050
   End
   Begin SICMACT.FlexEdit feDetalle 
      Height          =   4410
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   7779
      Cols0           =   7
      HighLight       =   2
      AllowUserResizing=   3
      EncabezadosNombres=   "N°-nID-Banco-N° Cheque-Girador-Moneda-Importe"
      EncabezadosAnchos=   "350-0-3200-3000-3200-1200-1500"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-L-L-C-R"
      FormatosEdit    =   "0-0-0-0-0-0-0"
      TextArray0      =   "N°"
      lbFlexDuplicados=   0   'False
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   3
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   11800
   End
End
Attribute VB_Name = "frmChequeEditSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'** Nombre : frmChequeEditSel
'** Descripción : Selección de Cheques para editar creado segun TI-ERS126-2013
'** Creación : EJVG, 20131220 11:00:00 AM
'*****************************************************************************

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Function cargar_datos() As Boolean
    Dim oDR As New COMNCajaGeneral.NCOMDocRec
    Dim oRS As New ADODB.Recordset
    Dim row As Long
    
    On Error GoTo Errcargar_datos
    Screen.MousePointer = 11
    Set oRS = oDR.ListaChequexMantenimiento(Right(gsCodAge, 2))
    If Not oRS.EOF Then
        FormateaFlex feDetalle
        Do While Not oRS.EOF
            feDetalle.AdicionaFila
            row = feDetalle.row
            feDetalle.TextMatrix(row, 1) = oRS!nId
            feDetalle.TextMatrix(row, 2) = oRS!cIFiNombre
            feDetalle.TextMatrix(row, 3) = oRS!cNroDoc
            feDetalle.TextMatrix(row, 4) = oRS!cGiradorNombre
            feDetalle.TextMatrix(row, 5) = oRS!cMoneda
            feDetalle.TextMatrix(row, 6) = Format(oRS!nMonto, gsFormatoNumeroView)
            oRS.MoveNext
        Loop
        feDetalle.row = 1
        feDetalle.TopRow = 1
        cmdEditar.Default = True
        cargar_datos = True
    Else
        cargar_datos = False
    End If
    Set oRS = Nothing
    Set oDR = Nothing
    Screen.MousePointer = 0
    Exit Function
Errcargar_datos:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Function
Private Sub feDetalle_DblClick()
    If feDetalle.row > 0 Then
        cmdEditar_Click
    End If
End Sub
Private Sub cmdActualizar_Click()
    cargar_datos
End Sub
Private Sub cmdEditar_Click()
    If feDetalle.TextMatrix(1, 0) = "" Then
        MsgBox "No existen cheques", vbInformation, "Aviso"
        Exit Sub
    End If
    Dim frm As New frmCheque
    frm.Editar (CLng(feDetalle.TextMatrix(feDetalle.row, 1)))
    Set frm = Nothing
End Sub
Private Sub Form_Load()
    If Not cargar_datos Then
        MsgBox "No existen Cheques para realizar el Mantenimiento", vbInformation, "Aviso"
    End If
End Sub
