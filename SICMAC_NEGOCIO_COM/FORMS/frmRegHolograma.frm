VERSION 5.00
Begin VB.Form frmRegHolograma 
   Caption         =   "Historial de Holograma"
   ClientHeight    =   3270
   ClientLeft      =   8745
   ClientTop       =   4725
   ClientWidth     =   6060
   Icon            =   "frmRegHolograma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6060
   Begin SICMACT.FlexEdit FlexHolog 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3413
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Fecha. Reg-Holog. Inicio-Holog. Fin-Contador-Estado-Codigo"
      EncabezadosAnchos=   "300-1200-1000-1000-900-1000-0"
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
      ColumnasAEditar =   "X-X-X-X-4-X-X"
      ListaControles  =   "0-0-0-0-1-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-L-C-C"
      FormatosEdit    =   "5-5-0-0-0-1-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   6
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdModalRegHolog 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmRegHolograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'* HISTORIAL DE HOLOGRAMAS
'*Archivo:  frmPigHolograma.frm
'*JUCS   :  28/11/2017
'*Resumen:  Nos permite listar los hologramas y ver el inventario disponible
'***************************************************************************************
Option Explicit
'APRI20190515 ERS005-2019
Dim gnContador As Integer
Dim gcEstado As String
'Dim cboAgencias As Integer
Private Sub CmdEditar_Click()
If Not IsNumeric(FlexHolog.TextMatrix(FlexHolog.row, 6)) Then
    MsgBox "Seleccione un registro", vbInformation, "Aviso"
    Exit Sub
ElseIf Trim(FlexHolog.TextMatrix(FlexHolog.row, 5)) = "INACTIVO" Then
    MsgBox "No se puede editar el registro porque está en estado Inactivo", vbInformation, "Aviso"
    Exit Sub
'COMENTADO POR APRI20190515 ERS005-2019
'ElseIf CInt(FlexHolog.TextMatrix(FlexHolog.row, 4)) > 0 Then
'    MsgBox "No se puede editar porque ya existe registro con este rango de Holograma", vbInformation, "Aviso"
'    Exit Sub
Else
    'frmPigRegHolograma.Inicio 2, FlexHolog.TextMatrix(FlexHolog.row, 6), FlexHolog.TextMatrix(FlexHolog.row, 2), FlexHolog.TextMatrix(FlexHolog.row, 3)
    frmPigRegHolograma.Inicio 2, FlexHolog.TextMatrix(FlexHolog.row, 6), FlexHolog.TextMatrix(FlexHolog.row, 2), FlexHolog.TextMatrix(FlexHolog.row, 3), CInt(FlexHolog.TextMatrix(FlexHolog.row, 4)) 'APRI20190515 ERS005-2019
    Call CargarDatosHologramas
End If
End Sub

Private Sub cmdModalRegHolog_Click()
    If Not VerificaExisteHologramaActivo(gsCodAge) Then
        frmPigRegHolograma.Inicio 1
        Call CargarDatosHologramas
    Else
        MsgBox "Ya existe un registro de Holograma Activo para esta Agencia", vbInformation, "Aviso"
        Exit Sub
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub FlexHolog_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    gnContador = FlexHolog.TextMatrix(FlexHolog.row, 4)
    gcEstado = FlexHolog.TextMatrix(FlexHolog.row, 5)
    If FlexHolog.TextMatrix(FlexHolog.row, 5) <> "INACTIVO" Then
        frmPigHistHolograma.Inicio FlexHolog.TextMatrix(FlexHolog.row, 6)
    End If
End Sub

Private Sub FlexHolog_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    FlexHolog.TextMatrix(pnRow, 5) = gcEstado
    FlexHolog.TextMatrix(pnRow, 4) = gnContador
End Sub

'Inicializa el formulario
Private Sub Form_Load()
Call CargarDatosHologramas
End Sub
Private Sub CargarDatosHologramas()
    Dim obj As New COMDColocPig.DCOMColPContrato
    Dim rs As ADODB.Recordset
    Dim i As Integer
  
    FlexHolog.Clear
    FormateaFlex FlexHolog
        Set rs = obj.ListaHistHologramas(gsCodAge)
        If Not (rs.EOF And rs.BOF) Then
            For i = 1 To rs.RecordCount
                FlexHolog.AdicionaFila
                'Para dar color al estado
                If rs!Estado = "ACTIVO" Then
                FlexHolog.Col = 5
                FlexHolog.CellForeColor = vbBlue
                Else
                FlexHolog.Col = 5
                FlexHolog.CellForeColor = vbRed
                End If
                FlexHolog.TextMatrix(i, 6) = rs!iHologramaID
                FlexHolog.TextMatrix(i, 5) = rs!Estado
                FlexHolog.TextMatrix(i, 4) = rs!contador
                FlexHolog.TextMatrix(i, 3) = rs!HOLOG_FIN
                FlexHolog.TextMatrix(i, 2) = rs!HOLOG_INI
                FlexHolog.TextMatrix(i, 1) = rs!Fecha
                rs.MoveNext
            Next i
            cmdEditar.Enabled = True
        Else
            cmdEditar.Enabled = False
        End If
End Sub




