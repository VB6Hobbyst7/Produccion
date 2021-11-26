VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmFlujoCajaProyectado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flujo de Caja Proyectado: Registro Anual"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10410
   Icon            =   "frmFlujoCajaProyectado.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   40
      TabIndex        =   4
      Top             =   5690
      Width           =   1200
   End
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   9150
      TabIndex        =   6
      Top             =   5690
      Width           =   1200
   End
   Begin VB.CommandButton btnProcesar 
      Caption         =   "&Procesar"
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
      Height          =   350
      Left            =   1250
      TabIndex        =   5
      Top             =   5690
      Width           =   1200
   End
   Begin VB.Frame fraRegistro 
      Caption         =   "Carga de Información"
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
      Height          =   5610
      Left            =   40
      TabIndex        =   7
      Top             =   40
      Width           =   10335
      Begin VB.Frame fraArchivo 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   90
         TabIndex        =   9
         Top             =   190
         Width           =   10095
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            ItemData        =   "frmFlujoCajaProyectado.frx":030A
            Left            =   6960
            List            =   "frmFlujoCajaProyectado.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   150
            Width           =   1815
         End
         Begin VB.TextBox txtArchivo 
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   720
            TabIndex        =   10
            Top             =   120
            Width           =   4575
         End
         Begin VB.CommandButton btnCargar 
            Caption         =   "&Cargar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   8880
            TabIndex        =   3
            Top             =   120
            Width           =   1200
         End
         Begin VB.CommandButton btnArchivo 
            Caption         =   "..."
            Height          =   350
            Left            =   5280
            TabIndex        =   0
            Top             =   120
            Width           =   405
         End
         Begin Spinner.uSpinner usAnio 
            Height          =   300
            Left            =   5760
            TabIndex        =   1
            Top             =   150
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Max             =   9999
            Min             =   1901
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
         End
         Begin VB.Label Label3 
            Caption         =   "Archivo"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   180
            Width           =   615
         End
      End
      Begin Sicmact.FlexEdit feFlujoCaja 
         Height          =   4815
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   10095
         _ExtentX        =   17754
         _ExtentY        =   8414
         Cols0           =   16
         HighLight       =   1
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   $"frmFlujoCajaProyectado.frx":030E
         EncabezadosAnchos=   "0-4500-0-0-1800-1800-1800-1800-1800-1800-1800-1800-1800-1800-1800-1800"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-R-R-R-R-R-R-R-R-R-R-R-R"
         FormatosEdit    =   "0-1-0-0-2-2-2-2-2-2-2-2-2-2-2-2"
         TextArray0      =   "#"
         lbFlexDuplicados=   0   'False
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   0
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmFlujoCajaProyectado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'** Nombre : frmFlujoCajaProyectado
'** Descripción : Subir información al Sistema  creado según RFC088-2012
'** Creación : EJVG, 20121022 09:00:00 AM
'***********************************************************************

Option Explicit
Dim fbExisteFCP As Boolean
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub Form_Load()
    CentraForm Me
    btnCancelar_Click
End Sub
Private Sub btnSalir_Click()
    Unload Me
End Sub
Private Sub btnArchivo_Click()
    dlgArchivo.InitDir = "C:\"
    dlgArchivo.Filter = "Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx"
    dlgArchivo.ShowOpen
    If dlgArchivo.FileName <> Empty Then
        txtArchivo.Text = dlgArchivo.FileName
    Else
        txtArchivo.Text = ""
    End If
End Sub

Private Sub btnCargar_Click()
    Dim objExcel As Excel.Application
    Dim xLibro As Excel.Workbook
    Dim oSheet As Excel.Worksheet
    Dim oCaja As New DCajaGeneral
    Dim rsCaja As New ADODB.Recordset
    
    Dim n As Integer, i As Integer, j As Integer
    Dim lnAnio As Integer, lnFCPId As Integer, lnFilaExcel As Integer
    Dim lnMoneda As Moneda
    
    If Not validaCargar Then Exit Sub
    
    lnAnio = usAnio.Valor
    lnMoneda = CInt(Trim(Right(cboMoneda.Text, 3)))
    
    Set rsCaja = oCaja.RecuperaFlujoCajaProyectadoxAnio(lnAnio, lnMoneda)
    
    If Not RSVacio(rsCaja) Then
        n = MsgBox("Flujo de Caja Proyectado ya fue procesado." & Chr(10) & "SI para cargar la información del archivo al Sistema" & Chr(10) & "NO para visualizar la información procesada anteriormente", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Aviso")
        If n = vbCancel Then
            Exit Sub
        ElseIf n = vbNo Then
            Call CargarGrillaFlujoCajaProyectado(rsCaja)
            txtArchivo.Text = ""
            fraArchivo.Enabled = False
            Exit Sub
        ElseIf n = vbYes Then
            fbExisteFCP = True
        End If
    End If

    Set objExcel = New Excel.Application
    Set xLibro = objExcel.Workbooks.Open(txtArchivo.Text)
    Set oSheet = xLibro.Worksheets(1)

    For i = 1 To feFlujoCaja.Rows - 1
        lnFCPId = feFlujoCaja.TextMatrix(i, 2)
        lnFilaExcel = feFlujoCaja.TextMatrix(i, 3)
        For j = 1 To 12
            feFlujoCaja.TextMatrix(i, j + 3) = Format(Val(oSheet.Cells(lnFilaExcel, j + 1)), gsFormatoNumeroView)
        Next
    Next
    
    fraArchivo.Enabled = False
    Me.btnProcesar.Enabled = True
    Me.btnProcesar.SetFocus
    MsgBox "Se ha cargado satisfactoriamente la información al Sistema", vbInformation, "Aviso"

    objExcel.Quit
    Set oCaja = Nothing
    Set objExcel = Nothing
    Set xLibro = Nothing
End Sub
Private Sub btnCancelar_Click()
    fbExisteFCP = False
    usAnio.Valor = Year(Now)
    MostrarGrillaFlujoCajaProyectadoDefault
    CargaMoneda
    txtArchivo.Text = ""
    btnProcesar.Enabled = False
    fraArchivo.Enabled = True
End Sub
Private Sub btnProcesar_Click()
    Dim oFlujoCaja As New nCajaGeneral
    Dim rs As New ADODB.Recordset
    Dim lnAnio As Integer, i As Integer
    Dim lnMoneda As Moneda
    Dim ldFecha As Date
    
    On Error GoTo ErrProcesar
    
    lnAnio = usAnio.Valor
    lnMoneda = CInt(Trim(Right(cboMoneda.Text, 3)))
    
    If fbExisteFCP Then
        If MsgBox("¿ Esta seguro de reemplazar la información del Flujo de Caja Proyectado año " & lnAnio & " ?", vbYesNo + vbInformation, "Aviso") = vbNo Then
            Exit Sub
        End If
    Else
        If MsgBox("¿ Esta seguro de procesar la información del Flujo de Caja Proyectado año " & lnAnio & " ?", vbYesNo + vbInformation, "Aviso") = vbNo Then
            Exit Sub
        End If
    End If
    
    ldFecha = CDate(gdFecSis & " " & Format(Time, "hh:mm:ss"))
    Set rs = feFlujoCaja.GetRsNew
    Call oFlujoCaja.GrabarFlujoCajaProyectado(lnAnio, lnMoneda, gsCodUser, ldFecha, rs)
    MsgBox "Se ha procesado satisfactoriamente el Flujo de Caja Proyectado año " & lnAnio, vbInformation, "Aviso"
    
    btnCancelar_Click
    Set oFlujoCaja = Nothing
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Proceso la Información "
                Set objPista = Nothing
                '****
    Exit Sub
ErrProcesar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub txtArchivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        txtArchivo.Text = ""
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub usAnio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboMoneda.SetFocus
    End If
End Sub
Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnCargar.SetFocus
    End If
End Sub
Private Function validaCargar() As Boolean
    Dim fs As New Scripting.FileSystemObject

    validaCargar = True
    If Len(Trim(txtArchivo.Text)) = 0 Then
        MsgBox "Ud. debe de seleccionar el archivo Flujo de Caja Proyectado", vbInformation, "Aviso"
        btnArchivo.SetFocus
        validaCargar = False
        Exit Function
    End If
    If UCase(Right(Trim(txtArchivo.Text), 5)) <> ".XLSX" And UCase(Right(Trim(txtArchivo.Text), 4)) <> ".XLS" Then
        MsgBox "Ud. debe de seleccionar el archivo de Flujo de Caja Proyectado en formato excel", vbInformation, "Aviso"
        btnArchivo.SetFocus
        validaCargar = False
        Exit Function
    End If
    If Not fs.FileExists(txtArchivo.Text) Then
        MsgBox "El archivo especificado NO existe, verifique", vbInformation, "Aviso"
        btnArchivo.SetFocus
        validaCargar = False
        Exit Function
    End If
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la Moneda del Flujo de Caja Proyectado", vbInformation, "Aviso"
        cboMoneda.SetFocus
        validaCargar = False
        Exit Function
    End If
End Function
Private Sub MostrarGrillaFlujoCajaProyectadoDefault()
    Dim oCaja As New DCajaGeneral
    Dim rs As New ADODB.Recordset
    Set rs = oCaja.RecuperaFlujoCajaProyectado()
    Call CargarGrillaFlujoCajaProyectado(rs)
End Sub
Private Sub CargarGrillaFlujoCajaProyectado(ByVal rs As ADODB.Recordset)
    Dim lnFila As Integer
    Call LimpiaFlex(feFlujoCaja)
    If Not RSVacio(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            feFlujoCaja.AdicionaFila
            lnFila = feFlujoCaja.row
            feFlujoCaja.TextMatrix(lnFila, 1) = IIf(rs!nOrden <> 0, Space(2) & rs!cFCPDescripcion, rs!cFCPDescripcion)
            feFlujoCaja.TextMatrix(lnFila, 2) = Format(rs!nFCPId, gsFormatoNumeroView)
            feFlujoCaja.TextMatrix(lnFila, 3) = rs!nFilaExcel
            feFlujoCaja.TextMatrix(lnFila, 4) = Format(rs!Enero, gsFormatoNumeroView)
            feFlujoCaja.TextMatrix(lnFila, 5) = Format(rs!Febrero, gsFormatoNumeroView)
            feFlujoCaja.TextMatrix(lnFila, 6) = Format(rs!Marzo, gsFormatoNumeroView)
            feFlujoCaja.TextMatrix(lnFila, 7) = Format(rs!Abril, gsFormatoNumeroView)
            feFlujoCaja.TextMatrix(lnFila, 8) = Format(rs!Mayo, gsFormatoNumeroView)
            feFlujoCaja.TextMatrix(lnFila, 9) = Format(rs!Junio, gsFormatoNumeroView)
            feFlujoCaja.TextMatrix(lnFila, 10) = Format(rs!Julio, gsFormatoNumeroView)
            feFlujoCaja.TextMatrix(lnFila, 11) = Format(rs!Agosto, gsFormatoNumeroView)
            feFlujoCaja.TextMatrix(lnFila, 12) = Format(rs!Septiembre, gsFormatoNumeroView)
            feFlujoCaja.TextMatrix(lnFila, 13) = Format(rs!Octubre, gsFormatoNumeroView)
            feFlujoCaja.TextMatrix(lnFila, 14) = Format(rs!Noviembre, gsFormatoNumeroView)
            feFlujoCaja.TextMatrix(lnFila, 15) = Format(rs!Diciembre, gsFormatoNumeroView)

            If rs!nOrden = 0 Then
                feFlujoCaja.BackColorRow &HE0E0E0, True
            End If
            rs.MoveNext
        Loop
    End If
End Sub
Private Sub CargaMoneda()
    Dim oCon As New DConstante
    CargaCombo cboMoneda, oCon.RecuperaConstantes(gMoneda)
    Set oCon = Nothing
End Sub
