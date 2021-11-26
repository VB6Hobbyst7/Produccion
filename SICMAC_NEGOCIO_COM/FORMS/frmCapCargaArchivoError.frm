VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCapCargaArchivoError 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Errores"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14280
   Icon            =   "frmCapCargaArchivoError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   14280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SICMACT.FlexEdit feError 
      Height          =   3900
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   6879
      Cols0           =   11
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-CodCliente-Observacion-Id-Servicio-Concepto-Doi-Cliente-Importe-TipoDoi-CodConvenio"
      EncabezadosAnchos=   "500-0-5000-800-1500-1500-1500-2500-1200-1200-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-C-L-L-L-L-L-L-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   11655
      Top             =   5670
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   360
      Left            =   12075
      TabIndex        =   2
      Top             =   4230
      Width           =   960
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   13125
      TabIndex        =   1
      Top             =   4230
      Width           =   960
   End
   Begin SICMACT.FlexEdit grdError 
      Height          =   1590
      Left            =   8085
      TabIndex        =   0
      Top             =   5355
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   2805
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Detalle Error-ID-Codigo-Tipo DOI-DOI-Cliente-Servicio-Concepto-Importe"
      EncabezadosAnchos=   "500-7000-700-1000-850-1000-3500-0-0-0"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-C-C-L-L-C-R-R"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmCapCargaArchivoError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************************************************************************************************
'* NOMBRE         : "frmCapCargaArchivoError"
'* DESCRIPCION    : Formulario creado para mostrar los errores al momento de la carga de trama segun proyecto: "Mejora del Sistema y Automatizacion de Ahorros y Servicios"
'* CREACION       : RIRO, 20121213 10:00 AM
'*****************************************************************************************************************************************************************************************

Option Explicit

Private rsErrores As ADODB.Recordset
Private sNombreEmpresa As String
Private sCodConvenio As String

Public Sub Inicia(ByVal pRSError As ADODB.Recordset, Optional pNombreEmpresa As String = "", Optional pCodConvenio As String = "")

    Dim i As Integer
    Dim nIndice As Integer
    
    LimpiaFlex grdError
    pRSError.MoveFirst
    sNombreEmpresa = pNombreEmpresa
    sCodConvenio = pCodConvenio
    nIndice = 1
    
    For i = 1 To pRSError.RecordCount
    
        feError.AdicionaFila
        
        feError.TextMatrix(i, 0) = nIndice
        feError.TextMatrix(i, 1) = "--"
        feError.TextMatrix(i, 2) = IIf(pRSError!cObservacion = "", " ", pRSError!cObservacion)
        feError.TextMatrix(i, 3) = IIf(pRSError!cId = "", " ", pRSError!cId)
        feError.TextMatrix(i, 4) = IIf(pRSError!cServicio = "", " ", pRSError!cServicio)
        feError.TextMatrix(i, 5) = IIf(pRSError!cConcepto = "", " ", pRSError!cConcepto)
        feError.TextMatrix(i, 6) = IIf(pRSError!cDOI = "", " ", pRSError!cDOI)
        feError.TextMatrix(i, 7) = IIf(pRSError!cNomCliente = "", " ", pRSError!cNomCliente)
        feError.TextMatrix(i, 8) = Format(pRSError!nImporte, "##,0.00")
        feError.TextMatrix(i, 9) = IIf(pRSError!nTipoDOI = 1, "DNI", IIf(pRSError!nTipoDOI = 2, "RUC", "OTRO"))
        feError.TextMatrix(i, 10) = IIf(pRSError!cCodConvenio = "", " ", pRSError!cCodConvenio)
        
        pRSError.MoveNext
        nIndice = nIndice + 1
        
    Next i
    
    Me.Show 1
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdExportar_Click()

    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim nFila, i As Double
    Dim campos() As String

    On Error GoTo error

    lsArchivo = "FormatoError"
    lsNomHoja = "ListadoErrores"
    nFila = 16
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    'campos = Split(sCabecera, ",")

    dlgArchivo.Filename = Empty
    dlgArchivo.Filter = "Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx"
    dlgArchivo.ShowSave

    If dlgArchivo.Filename = Empty Then
        MsgBox "No seleccionó ningun destino donde guardar", vbExclamation, "Aviso"
       Exit Sub
    End If

    lsArchivo1 = dlgArchivo.Filename
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
        If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
        End If
    Next
   
    xlHoja1.Cells(7, 3) = sNombreEmpresa ' empresa convenio
    xlHoja1.Cells(5, 3) = sCodConvenio ' numero convenio

    Dim k, j, c As Integer
    
    For k = 1 To feError.Rows - 1
                
        xlHoja1.Cells(k + 11, 3) = feError.TextMatrix(k, 2) 'Observacion
        xlHoja1.Cells(k + 11, 4) = feError.TextMatrix(k, 3) 'id
        xlHoja1.Cells(k + 11, 5) = feError.TextMatrix(k, 1) 'cod cliente
        xlHoja1.Cells(k + 11, 6) = feError.TextMatrix(k, 9) 'tipoDoi
        xlHoja1.Cells(k + 11, 7) = feError.TextMatrix(k, 6) 'doi
        xlHoja1.Cells(k + 11, 8) = feError.TextMatrix(k, 7) 'nombre
        xlHoja1.Cells(k + 11, 9) = feError.TextMatrix(k, 4) ' Servicio
        xlHoja1.Cells(k + 11, 10) = feError.TextMatrix(k, 5) ' concepto
        xlHoja1.Cells(k + 11, 11) = feError.TextMatrix(k, 8) ' Importe
           
    Next

'    Dim listFilas() As String
'    Dim listMensaje() As String
'    Dim listTrama() As String
'    Dim sError As String
'    Dim j, item As Double
'    Dim c, v, d, e As Variant
'    i = 12
'    j = 3
'    item = 0
'
'    For Each c In listError
'
'        If c <> "" Then
'            item = item + 1
'            listFilas = Split(c, "|")
'            listMensaje = Split(listFilas(0), ";")
'            listTrama = Split(listFilas(1), ";")
'            sError = ""
'
'            Dim col As Integer
'
'            For Each v In listMensaje
'                col = val(Right(v, 5))
'                If col <> 0 Then
'                    col = col + 3
'                    With xlHoja1.Cells(i, col).Interior
'                        .Pattern = xlSolid
'                        .PatternColorIndex = xlAutomatic
'                        .Color = 65535
'                        .TintAndShade = 0
'                        .PatternTintAndShade = 0
'                    End With
'                    sError = sError & " " & Trim(Mid(v, 1, Len(v) - 5)) & ","
'                Else
'                    sError = v & " "
'                End If
'            Next
'
'            sError = Mid(sError, 1, Len(sError) - 1)
'            xlHoja1.Cells(i, 3) = sError
'
'            For Each d In listTrama
'                j = j + 1
'                xlHoja1.Cells(i, j) = d
'            Next
'
'            xlHoja1.Cells(i, 2) = item
'            j = 3
'            i = i + 1
'
'        End If
'
'    Next

    xlHoja1.SaveAs lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

    Exit Sub

error:
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    MsgBox err.Description, vbCritical, "Aviso"

End Sub













