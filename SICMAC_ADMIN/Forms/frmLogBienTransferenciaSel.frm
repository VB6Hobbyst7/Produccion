VERSION 5.00
Begin VB.Form frmLogBienTransferenciaSel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Bien a Transferir"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10140
   Icon            =   "frmLogBienTransferenciaSel.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   9075
      TabIndex        =   2
      Top             =   2280
      Width           =   1050
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   7995
      TabIndex        =   1
      Top             =   2280
      Width           =   1050
   End
   Begin Sicmact.FlexEdit feBien 
      Height          =   2205
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   10080
      _ExtentX        =   15452
      _ExtentY        =   3889
      Cols0           =   11
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "#-Cod. Inventario-Nombre-Marca-Persona-Tipo Activación-Fecha Actv.-nUnico-nMovNro-nId-cPersCod"
      EncabezadosAnchos=   "350-1700-2200-1500-1400-1400-1200-0-0-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-C-L-L-C-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0"
      CantEntero      =   9
      TextArray0      =   "#"
      SelectionMode   =   1
      TipoBusqueda    =   0
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
   End
End
Attribute VB_Name = "frmLogBienTransferenciaSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'** Nombre : frmLogBienTransferenciaSel
'** Descripción : Selección de Bienes a Transferir creado segun ERS059-2013
'** Creación : EJVG, 20130618 03:30:00 AM
'***************************************************************************
Option Explicit
Dim fbAceptar As Boolean
Dim fsCodigo As String, fsDescripcion As String
Dim fsAreaAgeCod As String, fsBSCodCate As String

Private Sub Form_Load()
    CentraForm Me
    ListarSeries
End Sub
Public Sub Inicio(ByRef psCodigo As String, ByRef psDescripcion As String, ByVal psAreaAgeCod As String, ByVal psBSCodCate As String)
    fsCodigo = psCodigo
    fsDescripcion = psDescripcion
    fsAreaAgeCod = psAreaAgeCod
    fsBSCodCate = psBSCodCate
    fbAceptar = False
    Show 1
    If fbAceptar Then
        psCodigo = fsCodigo
        psDescripcion = fsDescripcion
    End If
End Sub
Private Sub ListarSeries()
    Dim oBien As New DBien
    Dim rs As New ADODB.Recordset
    Dim fila As Long
    Set rs = oBien.RecuperaSeriesConActivaciones(fsAreaAgeCod, fsBSCodCate)
    Call LimpiaFlex(feBien)
    Do While Not rs.EOF
        feBien.AdicionaFila
        fila = feBien.Row
        feBien.TextMatrix(fila, 1) = rs!cInventarioCod
        feBien.TextMatrix(fila, 2) = rs!cNombre
        feBien.TextMatrix(fila, 3) = rs!cMarca
        feBien.TextMatrix(fila, 4) = rs!cPersNombre
        feBien.TextMatrix(fila, 5) = rs!cTipoActivacion
        feBien.TextMatrix(fila, 6) = Format(rs!dActivacion, "dd/mm/yyyy")
        feBien.TextMatrix(fila, 7) = rs!nUnico
        feBien.TextMatrix(fila, 8) = rs!nMovNro
        feBien.TextMatrix(fila, 9) = rs!nId
        feBien.TextMatrix(fila, 10) = rs!cPersCod
        rs.MoveNext
    Loop
    feBien.TopRow = 1
    feBien.Row = 1
    Set oBien = Nothing
    Set rs = Nothing
End Sub
'Private Sub feBien_DblClick()
'    If Not FlexVacio(feBien) Then
'        If feBien.Row > 0 Then
'            Call cmdAceptar_Click
'        End If
'    End If
'End Sub
Private Sub cmdAceptar_Click()
    If Not FlexVacio(feBien) Then
        'Envio nUnico,nMovNro,nId,cPersCod y cPersNombre
        fsCodigo = feBien.TextMatrix(feBien.Row, 1)
        fsDescripcion = feBien.TextMatrix(feBien.Row, 2) & Space(500) & _
                        feBien.TextMatrix(feBien.Row, 7) & "," & _
                        feBien.TextMatrix(feBien.Row, 8) & "," & _
                        feBien.TextMatrix(feBien.Row, 9) & "," & _
                        feBien.TextMatrix(feBien.Row, 10) & "," & _
                        Replace(feBien.TextMatrix(feBien.Row, 4), ",", " ")
    Else
        MsgBox "Ud. primero debe seleccionar el Bien a Transferir", vbInformation, "Aviso"
        Exit Sub
    End If
    fbAceptar = True
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    fbAceptar = False
    Unload Me
End Sub
