VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNIIFBaseFormulasEEFFProy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estados Financieros Proyectados"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10125
   Icon            =   "frmNIIFBaseFormulasEEFFProy.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOpciones 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   675
      Left            =   80
      TabIndex        =   5
      Top             =   5280
      Width           =   9975
      Begin VB.CommandButton cmdGenFormato 
         Caption         =   "Generar &Formato"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   200
         Width           =   1410
      End
      Begin VB.TextBox txtRuta 
         Height          =   280
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   220
         Width           =   3660
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5325
         TabIndex        =   9
         Top             =   200
         Width           =   450
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "&Cargar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5805
         TabIndex        =   8
         Top             =   200
         Width           =   930
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7480
         TabIndex        =   7
         Top             =   200
         Width           =   1170
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8685
         TabIndex        =   6
         Top             =   200
         Width           =   1170
      End
   End
   Begin VB.Frame fraPeriodo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   675
      Left            =   80
      TabIndex        =   2
      Top             =   0
      Width           =   9975
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "&Seleccionar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         TabIndex        =   1
         Top             =   220
         Width           =   1170
      End
      Begin VB.TextBox txtAnio 
         Alignment       =   2  'Center
         Height          =   280
         Left            =   600
         MaxLength       =   4
         TabIndex        =   0
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   375
      End
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   4485
      Left            =   80
      TabIndex        =   4
      Top             =   735
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   7911
      Cols0           =   17
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "#-Nota-Nivel-Concepto-Enero-Febrero-Marzo-Abril-Mayo-Junio-Julio-Agosto-Septiembre-Octubre-Noviembre-Diciembre-Aux"
      EncabezadosAnchos=   "0-0-1300-2250-1000-1000-1000-1000-1000-1000-1000-1000-1100-1000-1000-1000-0"
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
      ColumnasAEditar =   "X-X-X-X-4-5-6-7-8-9-10-11-12-13-14-15-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-C-L-L-R-R-R-R-R-R-R-R-R-R-R-R-C"
      FormatosEdit    =   "0-0-0-0-2-2-2-2-2-2-2-2-2-2-2-2-0"
      CantEntero      =   9
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      TipoBusqueda    =   0
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      RowHeight0      =   300
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   120
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmNIIFBaseFormulasEEFFProy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
'** Nombre : frmNIIFBaseFormulasEEFFProy
'** Descripción : Configuración de información proyectada creado segun ERS057-2014
'** Creación : EJVG, 20140908 09:00:00 AM
'*********************************************************************************
Option Explicit
Dim fsOpeCod As String

Private Sub cmdBuscar_Click()
    On Error GoTo ErrBuscar
    txtRuta.Text = ""

    dlgArchivo.InitDir = "C:\"
    dlgArchivo.Filter = "Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx"
    dlgArchivo.FileName = ""
    dlgArchivo.ShowOpen
    
    Screen.MousePointer = 11
    If dlgArchivo.FileName <> "" Then
        txtRuta.Text = dlgArchivo.FileName
        cmdCargar.Enabled = True
    Else
        cmdCargar.Enabled = False
        MsgBox "No se selecciono ningún archivo", vbInformation, "Aviso"
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBuscar:
    If Err.Number = 32755 Then
        cmdCargar.Enabled = False
    ElseIf Err.Number = 1004 Then
        MsgBox "Archivo en uso. Ciérrelo y luego proceda a reemplazar.", vbExclamation, "Aviso"
    Else
        MsgBox "Error al momento de seleccionar el archivo", vbCritical, "Aviso"
    End If
End Sub
Private Sub cmdCargar_Click()
    On Error GoTo errcargar
    
    Dim oRS As New ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim lsAddress As String
    
    Dim objExcel As Excel.Application
    Dim xLibro As Excel.Workbook
    Dim Col As Integer, fila As Integer
    Dim psArchivoAGrabar As String
    Dim psArchivoAGrabarMenores As String
    Dim psArchivoAGrabarPersJurid As String
    
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object

    If Len(Trim(txtRuta.Text)) = 0 Then
        MsgBox "No selecciono ningun archivo", vbExclamation, "Aviso"
        EnfocaControl cmdBuscar
        Exit Sub
    End If
    
    Set objExcel = New Excel.Application
    Set xLibro = objExcel.Workbooks.Open(txtRuta.Text)
    
    'Valida cabecera
    i = 0
    For j = 1 To fg.Cols - 2
        If fg.TextMatrix(i, j) <> xLibro.Sheets(1).Cells(i + 1, j) Then
            MsgBox "La estructura del archivo seleccionado no es la correcta, verifique..", vbExclamation, "Aviso"
            objExcel.Quit
            Set xLibro = Nothing
            Set objExcel = Nothing
            Exit Sub
        End If
    Next
    'Valida detalle
    For i = 1 To fg.Rows - 1
        For j = 1 To fg.Cols - 2
            If j = 1 Then 'Valida código de Nota con primera columna excel
                If Val(fg.TextMatrix(i, j)) <> xLibro.Sheets(1).Cells(i + 1, j) Then
                    MsgBox "Los códigos de las Notas en el archivo seleccionado no corresponden a los de la Estructura actual, verifique..", vbExclamation, "Aviso"
                    objExcel.Quit
                    Set xLibro = Nothing
                    Set objExcel = Nothing
                    Exit Sub
                End If
            ElseIf j = 4 Or j = 5 Or j = 6 Or j = 7 Or j = 8 Or j = 9 Or j = 10 Or j = 11 Or j = 12 Or j = 13 Or j = 14 Or j = 15 Then
                If Not IsNumeric(xLibro.Sheets(1).Cells(i + 1, j)) Then 'Valida los montos de todos los meses
                    lsAddress = xLibro.Sheets(1).Cells(i + 1, j).Address
                    lsAddress = Replace(lsAddress, "$", "")
                    MsgBox "La celda " & lsAddress & " del archivo seleccionado no es un dato numerico, verifique..", vbExclamation, "Aviso"
                    objExcel.Quit
                    Set xLibro = Nothing
                    Set objExcel = Nothing
                    Exit Sub
                End If
            End If
        Next
    Next
    
    If MsgBox("¿Esta seguro de subir los datos del archivo excel a la grilla?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        objExcel.Quit
        Set xLibro = Nothing
        Set objExcel = Nothing
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    'Sube los datos cargados del excel al grid
    For i = 1 To fg.Rows - 1
        For j = 4 To fg.Cols - 2
            If j = 4 Or j = 5 Or j = 6 Or j = 7 Or j = 8 Or j = 9 Or j = 10 Or j = 11 Or j = 12 Or j = 13 Or j = 14 Or j = 15 Then
                fg.TextMatrix(i, j) = Format(CCur(xLibro.Sheets(1).Cells(i + 1, j)), gsFormatoNumeroView)
            End If
        Next
    Next
    Screen.MousePointer = 0
    MsgBox "Se ha subido los datos del archivo seleccionado a la grilla, ahora puede grabar los datos", vbInformation, "Aviso"
    
    txtRuta.Text = ""
    cmdCargar.Enabled = False
    Exit Sub
errcargar:
    Screen.MousePointer = 0
    If Not objExcel Is Nothing Then
        objExcel.Quit
        Set xLibro = Nothing
        Set objExcel = Nothing
    End If
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdCancelar_Click()
    Limpiar
    cmdCargar.Enabled = False
    fraOpciones.Enabled = False
    fraPeriodo.Enabled = True
End Sub
Private Sub cmdGenFormato_Click()
    On Error GoTo ErrGenerarFormato
    If Not ValidaDatos Then Exit Sub
    
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet, xlHoja1 As Excel.Worksheet
    Dim rs As New ADODB.Recordset
    Dim lnFilaActual As Integer, lnColumnaActual As Integer
    Dim i As Integer
    Dim lsArchivo As String
    Dim bFileRepet As Boolean
    
    bFileRepet = True
    dlgArchivo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx| Archivos de Excel (*.xls)|*.xls"
    dlgArchivo.FileName = "FormatoProyeccion" & Format(Now, "yyyyMMddhhnnss") & ".xlsx"
    Do While bFileRepet
        dlgArchivo.ShowSave
        Set fs = New Scripting.FileSystemObject
        Set xlsAplicacion = New Excel.Application
        If fs.FileExists(dlgArchivo.FileName) Then
            MsgBox "El archivo '" & dlgArchivo.FileTitle & "' ya existe, debe asignarle un nombre diferente", vbExclamation, "Aviso"
        Else
            bFileRepet = False
        End If
    Loop
    
    Screen.MousePointer = 11
    
    Set xlsLibro = xlsAplicacion.Workbooks.Add
    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "PROYECCION"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    
    Set rs = fg.GetRsNew
    lnFilaActual = 1
    lnColumnaActual = 1

    For i = 0 To rs.Fields.Count - 1
        xlsHoja.Cells(lnFilaActual, i + 1) = rs.Fields(i).Name
    Next i

    xlsHoja.Range(xlsHoja.Cells(lnFilaActual, 1), xlsHoja.Cells(lnFilaActual, i)).Cells.Interior.Color = RGB(220, 220, 220)
    xlsHoja.Range(xlsHoja.Cells(lnFilaActual, 1), xlsHoja.Cells(lnFilaActual, i)).HorizontalAlignment = xlCenter
    xlsHoja.Range("A2").CopyFromRecordset rs
    RSClose rs
    
    xlsHoja.Range("P:P").Delete
    xlsHoja.Range("D:O").NumberFormat = "#,##0.00"
    
    For Each xlHoja1 In xlsLibro.Worksheets
        If UCase(xlHoja1.Name) = "HOJA1" Or UCase(xlHoja1.Name) = "HOJA2" Or UCase(xlHoja1.Name) = "HOJA3" Then
            xlHoja1.Delete
        End If
    Next
    
    xlsHoja.SaveAs dlgArchivo.FileName
    MsgBox "Se ha generado satisfactoriamente el Formato" & Chr(10) & Chr(10) & "Si desea lo edita para volver a subir la información" & Chr(10) & "Columnas a modificar (meses): D-E-F-G-H-I-J-K-L-M-N-O", vbInformation, "Aviso"
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    
    Set xlsHoja = Nothing
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrGenerarFormato:
    Screen.MousePointer = 0
    If Not xlsAplicacion Is Nothing Then
        Set xlsHoja = Nothing
        Set xlsLibro = Nothing
        Set xlsAplicacion = Nothing
    End If
    If Err.Number = 32755 Then
    Else
        MsgBox Err.Description, vbCritical, "Aviso"
    End If
End Sub
Private Sub cmdGrabar_Click()
    Dim obj As NRepFormula
    Dim rs As ADODB.Recordset
    Dim bExito As Boolean
    
    If Not ValidaDatos Then Exit Sub
    
    If MsgBox("¿Esta seguro de grabar la información?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    On Error GoTo ErrGrabar
    Screen.MousePointer = 11
    Set obj = New NRepFormula
    Set rs = New ADODB.Recordset
    
    Set rs = fg.GetRsNew
    bExito = obj.GrabarProyeccionxOperacion(fsOpeCod, CInt(txtAnio.Text), rs)
    If bExito Then
        MsgBox "Se ha registrado con éxito los cambios", vbInformation, "Aviso"
    Else
        MsgBox "Ha sucedido un error al grabar la operación, si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrGrabar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Function ValidaDatos() As Boolean
    ValidaDatos = False
    If Len(txtAnio) <= 3 Then
        MsgBox "Ud. debe especificar un año válido", vbInformation, "Aviso"
        EnfocaControl txtAnio
        Exit Function
    End If
    If fg.TextMatrix(1, 0) = "" Then
        MsgBox "No existen datos", vbInformation, "Aviso"
        Exit Function
    End If
    ValidaDatos = True
End Function
Private Sub cmdSeleccionar_Click()
    Dim obj As DRepFormula
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    
    On Error GoTo ErrSeleccionar
    If Len(txtAnio) <= 3 Then
        MsgBox "Ud. debe especificar un año valido", vbInformation, "Aviso"
        EnfocaControl txtAnio
        Exit Sub
    End If
    Screen.MousePointer = 11
    Set obj = New DRepFormula
    Set rs = New ADODB.Recordset
    
    FormateaFlex fg
    Set rs = obj.ListaProyeccionxOperacion(fsOpeCod, CInt(txtAnio))
    If Not rs.EOF Then
        Do While Not rs.EOF
            fg.AdicionaFila
            fila = fg.row
            fg.TextMatrix(fila, 1) = rs!nCorreInt
            fg.TextMatrix(fila, 2) = rs!cNivel
            fg.TextMatrix(fila, 3) = rs!cConceptoDesc
            fg.TextMatrix(fila, 4) = Format(rs!nEnero, gsFormatoNumeroView)
            fg.TextMatrix(fila, 5) = Format(rs!nFebrero, gsFormatoNumeroView)
            fg.TextMatrix(fila, 6) = Format(rs!nMarzo, gsFormatoNumeroView)
            fg.TextMatrix(fila, 7) = Format(rs!nAbril, gsFormatoNumeroView)
            fg.TextMatrix(fila, 8) = Format(rs!nMayo, gsFormatoNumeroView)
            fg.TextMatrix(fila, 9) = Format(rs!nJunio, gsFormatoNumeroView)
            fg.TextMatrix(fila, 10) = Format(rs!nJulio, gsFormatoNumeroView)
            fg.TextMatrix(fila, 11) = Format(rs!nAgosto, gsFormatoNumeroView)
            fg.TextMatrix(fila, 12) = Format(rs!nSeptiembre, gsFormatoNumeroView)
            fg.TextMatrix(fila, 13) = Format(rs!nOctubre, gsFormatoNumeroView)
            fg.TextMatrix(fila, 14) = Format(rs!nNoviembre, gsFormatoNumeroView)
            fg.TextMatrix(fila, 15) = Format(rs!nDiciembre, gsFormatoNumeroView)
            rs.MoveNext
        Loop
        fg.row = 1
        fg.TopRow = 1
        fraOpciones.Enabled = True
        fraPeriodo.Enabled = False
    Else
        fraOpciones.Enabled = False
        fraPeriodo.Enabled = True
    End If
    
    Screen.MousePointer = 0
    Exit Sub
ErrSeleccionar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub Form_Load()
    Limpiar
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl cmdSeleccionar
    Else
        FormateaFlex fg
    End If
End Sub
Private Sub Limpiar()
    txtAnio.Text = Year(gdFecSis)
    FormateaFlex fg
    txtRuta.Text = ""
End Sub
Public Sub Inicio(ByVal psOpeCod As String)
    fsOpeCod = psOpeCod
    Show 1
End Sub
