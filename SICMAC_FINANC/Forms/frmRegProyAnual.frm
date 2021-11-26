VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmRegProyAnual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Proyecciones Anuales"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16095
   Icon            =   "frmRegProyAnual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   16095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraProyGasto 
      Caption         =   "Proyección Gastos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15930
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   5880
         Width           =   1095
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   5880
         Width           =   1095
      End
      Begin Sicmact.FlexEdit fgRegConcepto 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   15735
         _ExtentX        =   27755
         _ExtentY        =   7858
         Cols0           =   16
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Nivel-Concepto-Enero-Febrero-Marzo-Abril-Mayo-Junio-Julio-Agosto-Setiembre-Octubre-Noviembre-Diciembre-nOrden"
         EncabezadosAnchos=   "0-500-3100-1000-1000-1000-1000-1000-1000-1000-1000-1000-1000-1000-1000-0"
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
         ColumnasAEditar =   "X-X-X-3-4-5-6-7-8-9-10-11-12-13-14-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-R-L-R-C-C-C-C-C-C-C-C-C-C-C-R"
         FormatosEdit    =   "0-3-0-4-4-4-4-4-4-4-4-4-4-4-4-3"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
      Begin VB.Frame fraRegConcepto 
         Caption         =   "Registro de Conceptos"
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   15675
         Begin VB.CommandButton cmdMostrar 
            Caption         =   "Mostrar"
            Height          =   375
            Left            =   2040
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
         Begin Spinner.uSpinner txtAnio 
            Height          =   255
            Left            =   840
            TabIndex        =   2
            Top             =   285
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   450
            Max             =   9999
            Min             =   1990
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
         End
         Begin VB.Label lblAnio 
            Alignment       =   2  'Center
            Caption         =   "Año:"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   280
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "frmRegProyAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'***Nombre:         frmConfigRepGastoxNiveles
'***Descripción:    Formulario que permite la configuracion
'                   de reporte de gastos por niveles.
'***Creación:       MIOL el 20130529 según ERS033-2013 OBJ B
'************************************************************
Option Explicit
Dim oRepCtaColumna As DRepCtaColumna
Dim nAnio As Integer
Dim nEditar As Integer

Private Sub cmdExportar_Click()
    Me.MousePointer = vbHourglass
        Dim sPathProyAnual As String
       
        Dim fs As New Scripting.FileSystemObject
        Dim obj_Excel As Object, Libro As Object, Hoja As Object
        
        Dim convert As Double
        
        On Error GoTo error_sub
          
        sPathProyAnual = App.path & "\Spooler\ProyAnual_" + Me.txtAnio.Valor + ".xls"
        
        If fs.FileExists(sPathProyAnual) Then
            
            If ArchivoEstaAbierto(sPathProyAnual) Then
                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(sPathProyAnual) + " para continuar", vbRetryCancel) = vbCancel Then
                   Me.MousePointer = vbDefault
                   Exit Sub
                End If
                Me.MousePointer = vbHourglass
            End If
    
            fs.DeleteFile sPathProyAnual, True
        End If
        
        sPathProyAnual = App.path & "\FormatoCarta\Plantilla_ProyAnual.xls"

        If Len(Dir(sPathProyAnual)) = 0 Then
           MsgBox "No se Pudo Encontrar el Archivo:" & sPathProyAnual, vbCritical
           Me.MousePointer = vbDefault
           Exit Sub
        End If
        
        Set obj_Excel = CreateObject("Excel.Application")
        obj_Excel.DisplayAlerts = False
        Set Libro = obj_Excel.Workbooks.Open(sPathProyAnual)
        Set Hoja = Libro.ActiveSheet
        
        Dim celda As Excel.Range
        Set oRepCtaColumna = New DRepCtaColumna
        Dim rsCtaColumna As ADODB.Recordset
        
       ' Fecha del ANEXO
        Set celda = obj_Excel.Range("ProyAnual!D1")
        celda.value = Me.txtAnio.Valor

        '****************************CARGAR DATOS *******************************
        Set rsCtaColumna = oRepCtaColumna.GetConfGastoNivelProyAnual(Me.txtAnio.Valor)
        Dim nFilas As Integer
        Dim nItem As Integer
        
        nFilas = 4
        nItem = 1
        If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Set celda = obj_Excel.Range("ProyAnual!A" & nFilas)
               celda.value = nItem
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!B" & nFilas)
               celda.value = Format(rsCtaColumna(1), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!C" & nFilas)
               celda.value = Format(rsCtaColumna(2), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!D" & nFilas)
               celda.value = Format(rsCtaColumna(3), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!E" & nFilas)
               celda.value = Format(rsCtaColumna(4), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!F" & nFilas)
               celda.value = Format(rsCtaColumna(5), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!G" & nFilas)
               celda.value = Format(rsCtaColumna(6), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!H" & nFilas)
               celda.value = Format(rsCtaColumna(7), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!I" & nFilas)
               celda.value = Format(rsCtaColumna(8), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!J" & nFilas)
               celda.value = Format(rsCtaColumna(9), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!K" & nFilas)
               celda.value = Format(rsCtaColumna(10), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!L" & nFilas)
               celda.value = Format(rsCtaColumna(11), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!M" & nFilas)
               celda.value = Format(rsCtaColumna(12), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!N" & nFilas)
               celda.value = Format(rsCtaColumna(13), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               Set celda = obj_Excel.Range("ProyAnual!O" & nFilas)
               celda.value = Format(rsCtaColumna(14), "#,###0.00")
               celda.Cells(1, 1).Borders.LineStyle = 1
               nItem = nItem + 1
               nFilas = nFilas + 1
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        'verifica si existe el archivo
        sPathProyAnual = App.path & "\Spooler\PROYANUAL_" + Me.txtAnio.Valor + ".xls"
        If fs.FileExists(sPathProyAnual) Then
            
            If ArchivoEstaAbierto(sPathProyAnual) Then
                MsgBox "Debe Cerrar el Archivo:" + fs.GetFileName(sPathProyAnual)
            End If
            fs.DeleteFile sPathProyAnual, True
        End If
        'guarda el archivo
        Hoja.SaveAs sPathProyAnual

        Libro.Close
        obj_Excel.Quit
        Set Hoja = Nothing
        Set Libro = Nothing
        Set obj_Excel = Nothing
        Me.MousePointer = vbDefault
        'abre y muestra el archivo
        Dim m_excel As New Excel.Application
        m_excel.Workbooks.Open (sPathProyAnual)
        m_excel.Visible = True
Exit Sub
error_sub:
        MsgBox TextErr(Err.Description), vbInformation, "Aviso"
        Set Libro = Nothing
        Set obj_Excel = Nothing
        Set Hoja = Nothing
        Me.MousePointer = vbDefault
End Sub

Private Sub cmdMostrar_Click()
cargarDatosEstructura (txtAnio.Valor)
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub fgRegConcepto_OnCellChange(pnRow As Long, pnCol As Long)
Set oRepCtaColumna = New DRepCtaColumna
Dim rsProyAnual As ADODB.Recordset
Set rsProyAnual = New ADODB.Recordset
Dim nOrden As Integer

nOrden = fgRegConcepto.TextMatrix(fgRegConcepto.Row, 15)

Set rsProyAnual = oRepCtaColumna.GetProyeccionAnualxOrden(nOrden)
    If rsProyAnual.RecordCount > 0 Then
        Call oRepCtaColumna.ActualizarProyeccionAnual(nOrden, fgRegConcepto.TextMatrix(fgRegConcepto.Row, 3), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 4), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 5), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 6), _
                                                    fgRegConcepto.TextMatrix(fgRegConcepto.Row, 7), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 8), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 9), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 10), _
                                                    fgRegConcepto.TextMatrix(fgRegConcepto.Row, 11), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 12), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 13), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 14))
    Else
        Call oRepCtaColumna.InsertarProyeccionAnual(nOrden, fgRegConcepto.TextMatrix(fgRegConcepto.Row, 3), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 4), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 5), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 6), _
                                                    fgRegConcepto.TextMatrix(fgRegConcepto.Row, 7), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 8), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 9), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 10), _
                                                    fgRegConcepto.TextMatrix(fgRegConcepto.Row, 11), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 12), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 13), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 14))
    End If

Set rsProyAnual = Nothing
Set oRepCtaColumna = Nothing
End Sub

Private Sub fgRegConcepto_OnRowChange(pnRow As Long, pnCol As Long)
Set oRepCtaColumna = New DRepCtaColumna
Dim rsProyAnual As ADODB.Recordset
Set rsProyAnual = New ADODB.Recordset
Dim nOrden As Integer

nOrden = fgRegConcepto.TextMatrix(fgRegConcepto.Row, 15)

Set rsProyAnual = oRepCtaColumna.GetProyeccionAnualxOrden(nOrden)
    If rsProyAnual.RecordCount > 0 Then
        Call oRepCtaColumna.ActualizarProyeccionAnual(nOrden, fgRegConcepto.TextMatrix(fgRegConcepto.Row, 3), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 4), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 5), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 6), _
                                                    fgRegConcepto.TextMatrix(fgRegConcepto.Row, 7), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 8), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 9), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 10), _
                                                    fgRegConcepto.TextMatrix(fgRegConcepto.Row, 11), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 12), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 13), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 14))
    Else
        Call oRepCtaColumna.InsertarProyeccionAnual(nOrden, fgRegConcepto.TextMatrix(fgRegConcepto.Row, 3), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 4), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 5), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 6), _
                                                    fgRegConcepto.TextMatrix(fgRegConcepto.Row, 7), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 8), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 9), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 10), _
                                                    fgRegConcepto.TextMatrix(fgRegConcepto.Row, 11), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 12), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 13), fgRegConcepto.TextMatrix(fgRegConcepto.Row, 14))
    End If

Set rsProyAnual = Nothing
Set oRepCtaColumna = Nothing
End Sub

Private Sub fgRegConcepto_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)

    Dim scolumnas() As String
    scolumnas = Split(fgRegConcepto.ColumnasAEditar, "-")
    
    If scolumnas(fgRegConcepto.Col) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If

End Sub

Private Sub Form_Load()
    nAnio = Right(CStr(gdFecSis), 4)
    Me.txtAnio.Valor = nAnio
    cargarDatosEstructura (nAnio)
End Sub

Private Sub cargarDatosEstructura(ByVal psAnio As String)
 Set oRepCtaColumna = New DRepCtaColumna
 Dim rsConfGastoNivel As ADODB.Recordset
 Set rsConfGastoNivel = New ADODB.Recordset
 Dim i As Integer

   Call LimpiaFlex(fgRegConcepto)

   Set rsConfGastoNivel = oRepCtaColumna.GetConfGastoNivelProyAnual(psAnio)
        If Not rsConfGastoNivel.BOF And Not rsConfGastoNivel.EOF Then
            i = 1
            fgRegConcepto.lbEditarFlex = True
            Do While Not rsConfGastoNivel.EOF
                fgRegConcepto.AdicionaFila
                fgRegConcepto.TextMatrix(i, 1) = rsConfGastoNivel!nNivel
                fgRegConcepto.TextMatrix(i, 2) = rsConfGastoNivel!cConcepto
                fgRegConcepto.TextMatrix(i, 3) = Format(rsConfGastoNivel!nEnero, "#,###0.00")
                fgRegConcepto.TextMatrix(i, 4) = Format(rsConfGastoNivel!nFebrero, "#,###0.00")
                fgRegConcepto.TextMatrix(i, 5) = Format(rsConfGastoNivel!nMarzo, "#,###0.00")
                fgRegConcepto.TextMatrix(i, 6) = Format(rsConfGastoNivel!nAbril, "#,###0.00")
                fgRegConcepto.TextMatrix(i, 7) = Format(rsConfGastoNivel!nMayo, "#,###0.00")
                fgRegConcepto.TextMatrix(i, 8) = Format(rsConfGastoNivel!nJunio, "#,###0.00")
                fgRegConcepto.TextMatrix(i, 9) = Format(rsConfGastoNivel!nJulio, "#,###0.00")
                fgRegConcepto.TextMatrix(i, 10) = Format(rsConfGastoNivel!nAgosto, "#,###0.00")
                fgRegConcepto.TextMatrix(i, 11) = Format(rsConfGastoNivel!nSetiembre, "#,###0.00")
                fgRegConcepto.TextMatrix(i, 12) = Format(rsConfGastoNivel!nOctubre, "#,###0.00")
                fgRegConcepto.TextMatrix(i, 13) = Format(rsConfGastoNivel!nNoviembre, "#,###0.00")
                fgRegConcepto.TextMatrix(i, 14) = Format(rsConfGastoNivel!nDiciembre, "#,###0.00")
                fgRegConcepto.TextMatrix(i, 15) = rsConfGastoNivel!nOrden
                i = i + 1
                rsConfGastoNivel.MoveNext
            Loop
        End If
    Set rsConfGastoNivel = Nothing
    Set oRepCtaColumna = Nothing
End Sub

Private Function ArchivoEstaAbierto(ByVal Ruta As String) As Boolean
On Error GoTo HayErrores
Dim f As Integer
   f = FreeFile
   Open Ruta For Append As f
   Close f
   ArchivoEstaAbierto = False
   Exit Function
HayErrores:
   If Err.Number = 70 Then
      ArchivoEstaAbierto = True
   Else
      Err.Raise Err.Number
   End If
End Function
