VERSION 5.00
Begin VB.Form frmCapMigracionAhorrosError 
   Caption         =   "Error de Migración"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2895
   Icon            =   "frmCapMigracionAhorrosError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin SICMACT.FlexEdit FEError 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   5530
      Cols0           =   3
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-N° Cuenta-Glosa"
      EncabezadosAnchos=   "500-1800-0"
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
      ColumnasAEditar =   "X-X-X"
      ListaControles  =   "0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L"
      FormatosEdit    =   "0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "frmCapMigracionAhorrosError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCapMigracionAhorrosError
'*** Descripción : Formulario para migrar cuentas de ahorros.
'*** Creación : ELRO, 20130401 07:09:58 PM, según TI-ERS011-2013
'********************************************************************
Option Explicit

Public Sub cargarDatos(ByVal psCtaCod As String, ByVal psGlosa As String, ByVal pbError As Boolean)
    If pbError = False Then
    Call LimpiaFlex(FEError)
    End If
    FEError.AdicionaFila
    FEError.lbEditarFlex = True
    FEError.TextMatrix(FEError.Row, 1) = psCtaCod
    FEError.TextMatrix(FEError.Row, 2) = psGlosa
    FEError.lbEditarFlex = False
End Sub

Private Sub cmdExportar_Click()
'Variable de tipo Aplicación de Excel
Dim oExcel As Excel.Application
Dim lnTipoDOI, lnFila1, lnFila2, lnFilasFormato As Integer
Dim lsDOI As String
Dim lsMoneda As String
Dim lbExisteError As Boolean
Dim i, lnFilas As Integer
'Una variable de tipo Libro de Excel
Dim oLibro As Excel.Workbook
Dim oHoja As Excel.Worksheet

'***Para verificar la existencia del archivo xls
Dim fs As Scripting.FileSystemObject
Dim lsArchivo, lsArchivo1, lsArchivo2 As String
Set fs = New Scripting.FileSystemObject
lsArchivo = "ErrorMigracion"
lsArchivo1 = Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS")
lsArchivo2 = "\SPOOLER\ErrorMigracion_" & lsArchivo1 & ".xls"
'***Fin Para verificar la existencia del archivo xls
 

'creamos un nuevo objeto excel
Set oExcel = New Excel.Application

'Usamos el método open para abrir el archivo que está en el directorio del programa llamado archivo.xls
If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
    Set oLibro = oExcel.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")

 Else
    MsgBox "No existe la plantilla ErrorMigracion.xls en la carpeta FormatoCarta, Consulte con el Area de  TI", vbCritical, "Advertencia"
    Set oHoja = Nothing
    Set oLibro = Nothing
    Set oExcel = Nothing
    Exit Sub
 End If

'Hacemos referencia a la Hoja
Set oHoja = oLibro.Sheets(1)

lnFilas = FEError.Rows

With oHoja
    For i = 1 To lnFilas - 1
        .Cells(i + 1, 1) = FEError.TextMatrix(i, 1)
        .Cells(i + 1, 2) = FEError.TextMatrix(i, 2)
    Next i
End With

oHoja.Activate

oLibro.SaveAs (App.path & lsArchivo2)
oExcel.Visible = True

Set oHoja = Nothing
Set oLibro = Nothing
Set oExcel = Nothing

End Sub
