VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmNIIFNotasEstado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas de Estado"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   Icon            =   "frmNIIFNotasEstado.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPeriodoCompara 
      Caption         =   "Periodo a Comparar"
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar BarraProgreso 
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   2640
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar EstadoBarra 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   2595
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Moneda"
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
      TabIndex        =   11
      Top             =   1440
      Width           =   5055
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmNIIFNotasEstado.frx":030A
         Left            =   1080
         List            =   "frmNIIFNotasEstado.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   270
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   6
      ToolTipText     =   "Generar Reporte Nota Estado"
      Top             =   2210
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
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
      Height          =   360
      Left            =   2640
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   2210
      Width           =   1575
   End
   Begin VB.Frame fraPerEvaluar 
      Caption         =   "Periodo a Evaluar"
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
      TabIndex        =   8
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   280
         Left            =   720
         MaxLength       =   4
         TabIndex        =   0
         Top             =   240
         Width           =   780
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmNIIFNotasEstado.frx":030E
         Left            =   2760
         List            =   "frmNIIFNotasEstado.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   2280
         TabIndex        =   9
         Top             =   270
         Width           =   390
      End
   End
   Begin VB.Frame fraPerComparar 
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
      TabIndex        =   15
      Top             =   720
      Width           =   5055
      Begin VB.ComboBox cboMesCompara 
         Height          =   315
         ItemData        =   "frmNIIFNotasEstado.frx":0312
         Left            =   2760
         List            =   "frmNIIFNotasEstado.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtAnioCompara 
         Alignment       =   1  'Right Justify
         Height          =   280
         Left            =   720
         MaxLength       =   4
         TabIndex        =   3
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   2280
         TabIndex        =   17
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   270
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmNIIFNotasEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmNIIFNotasEstado
'** Descripción : Generación de Reportes Notas Estado creado segun ERS052-2013
'** Creación : EJVG, 20130413 09:00:00 AM
'********************************************************************
Option Explicit

Dim fsOpeCod As String
Private Type TCtaCont
    CuentaContable As String
    Saldo As Currency
End Type

Private Sub cboMes_LostFocus()
    cboMesCompara.ListIndex = cboMes.ListIndex
End Sub
Private Sub Form_Load()
    CentraForm Me
    cargarMoneda
    cargarMes
    txtAnio.Text = Year(gdFecSis)
    txtAnioCompara.Text = Year(gdFecSis) - 1
    cboMes.ListIndex = IndiceListaCombo(cboMes, Month(gdFecSis))
    cboMesCompara.ListIndex = IndiceListaCombo(cboMesCompara, Month(gdFecSis))
End Sub
Public Sub Inicio(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    fsOpeCod = psOpeCod
    Caption = UCase(psOpeDesc)
    Show 1
End Sub
Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdGenerar.SetFocus
    End If
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cboMes.SetFocus
    End If
End Sub
Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkPeriodoCompara.SetFocus
    End If
End Sub
Private Sub txtAnio_LostFocus()
    If Val(txtAnio.Text) > 0 Then
        txtAnioCompara.Text = Val(txtAnio.Text) - 1
    End If
End Sub
Private Sub txtAnioCompara_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cboMesCompara.SetFocus
    End If
End Sub
Private Sub cboMesCompara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CboMoneda.SetFocus
    End If
End Sub
Private Sub chkMuestraResultAnioAnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub
Private Sub chkPeriodoCompara_Click()
    If chkPeriodoCompara.value = vbChecked Then
        fraPerComparar.Enabled = True
    Else
        fraPerComparar.Enabled = False
    End If
End Sub
Private Sub chkPeriodoCompara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkPeriodoCompara.value = vbChecked Then
            txtAnioCompara.SetFocus
        Else
            CboMoneda.SetFocus
        End If
    End If
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub cmdGenerar_Click()
    Dim ldFechaEvalua As Date
    Dim ldFechaCompara As Date
    Dim lnMoneda As Integer
        
    On Error GoTo ErrGenerar
    
    If validaGenerar = False Then Exit Sub
    Screen.MousePointer = 11
    ldFechaEvalua = obtenerFechaFinMes(cboMes.ListIndex + 1, txtAnio.Text)
    ldFechaCompara = obtenerFechaFinMes(cboMesCompara.ListIndex + 1, txtAnioCompara.Text)
    lnMoneda = CInt(Trim(Right(CboMoneda.Text, 2)))
    
    Select Case fsOpeCod
        Case gContRepBaseNotasEstadoSitFinan, gContRepBaseNotasEstadoResultado
            Call generarReporteNotasEstado(fsOpeCod, ldFechaEvalua, IIf(chkPeriodoCompara.value = vbChecked, True, False), ldFechaCompara, lnMoneda)
        Case Else
            MsgBox "Lo sentimos esta opción no puede generar el presente reporte", vbCritical, "Aviso"
            Exit Sub
    End Select
    Screen.MousePointer = 0
    Exit Sub
ErrGenerar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Function validaGenerar() As Boolean
    validaGenerar = True
    If Val(txtAnio.Text) <= 1900 Then
        MsgBox "Ud. debe especificar el año del Periodo a Evaluar", vbInformation, "Aviso"
        txtAnio.SetFocus
        validaGenerar = False
        Exit Function
    End If
    If cboMes.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el mes del Periodo a Evaluar", vbInformation, "Aviso"
        cboMes.SetFocus
        validaGenerar = False
        Exit Function
    End If
    If chkPeriodoCompara.value = vbChecked Then
        If Val(txtAnioCompara.Text) <= 1900 Then
            MsgBox "Ud. debe especificar el año del Periodo a Comparar", vbInformation, "Aviso"
            txtAnioCompara.SetFocus
            validaGenerar = False
            Exit Function
        End If
        If cboMesCompara.ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar el mes del Periodo a Comparar", vbInformation, "Aviso"
            cboMesCompara.SetFocus
            validaGenerar = False
            Exit Function
        End If
        If (Val(txtAnio.Text) = Val(txtAnioCompara.Text)) And (cboMes.ListIndex = cboMesCompara.ListIndex) Then
            MsgBox "Ud. debe seleccionar periodos diferentes para poder comparar", vbInformation, "Aviso"
            validaGenerar = False
            Exit Function
        End If
    End If
    If CboMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el tipo de moneda", vbInformation, "Aviso"
        CboMoneda.SetFocus
        validaGenerar = False
        Exit Function
    End If
End Function
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
Private Sub cargarMoneda()
    CboMoneda.AddItem "UNIFICADO" & Space(200) & "0"
    CboMoneda.AddItem "SOLES" & Space(200) & "1"
    CboMoneda.AddItem "DOLARES" & Space(200) & "2"
End Sub
Private Sub cargarMes()
    cboMes.AddItem "ENERO" & Space(200) & "1"
    cboMes.AddItem "FEBRERO" & Space(200) & "2"
    cboMes.AddItem "MARZO" & Space(200) & "3"
    cboMes.AddItem "ABRIL" & Space(200) & "4"
    cboMes.AddItem "MAYO" & Space(200) & "5"
    cboMes.AddItem "JUNIO" & Space(200) & "6"
    cboMes.AddItem "JULIO" & Space(200) & "7"
    cboMes.AddItem "AGOSTO" & Space(200) & "8"
    cboMes.AddItem "SEPTIEMBRE" & Space(200) & "9"
    cboMes.AddItem "OCTUBRE" & Space(200) & "10"
    cboMes.AddItem "NOVIEMBRE" & Space(200) & "11"
    cboMes.AddItem "DICIEMBRE" & Space(200) & "12"
    
    cboMesCompara.AddItem "ENERO" & Space(200) & "1"
    cboMesCompara.AddItem "FEBRERO" & Space(200) & "2"
    cboMesCompara.AddItem "MARZO" & Space(200) & "3"
    cboMesCompara.AddItem "ABRIL" & Space(200) & "4"
    cboMesCompara.AddItem "MAYO" & Space(200) & "5"
    cboMesCompara.AddItem "JUNIO" & Space(200) & "6"
    cboMesCompara.AddItem "JULIO" & Space(200) & "7"
    cboMesCompara.AddItem "AGOSTO" & Space(200) & "8"
    cboMesCompara.AddItem "SEPTIEMBRE" & Space(200) & "9"
    cboMesCompara.AddItem "OCTUBRE" & Space(200) & "10"
    cboMesCompara.AddItem "NOVIEMBRE" & Space(200) & "11"
    cboMesCompara.AddItem "DICIEMBRE" & Space(200) & "12"
End Sub
Private Sub generarReporteNotasEstado(ByVal psOpeCod As String, ByVal pdFechaEvalua As Date, ByVal pbComparaPeriodo As Boolean, pdFechaCompara As Date, ByVal pnMoneda As Integer)
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim fs As New Scripting.FileSystemObject
    Dim oRep As New NRepFormula
    Dim rsNotas As New ADODB.Recordset
    Dim rsNotasDet As New ADODB.Recordset
    Dim iCab As Long, iDet As Long
    Dim lsPath As String, lsArchivo As String
    Dim lbAbierto As Boolean
    Dim lnFilaActual As Integer, lnColumnaActual As Integer
    Dim lnNivelMax As Integer, lnUltimaColumna As Integer
    Dim lsNombreMesEvalua As String, lsNombreMesCompara As String
    Dim lnAnioEvalua As Integer, lnAnioCompara As Integer
    Dim lsFormula1Evalua As String, lsFormula1Compara As String
    Dim lsFormula2Evalua As String, lsFormula2Compara As String
    Dim lsFormula3Evalua As String, lsFormula3Compara As String
    Dim lsFormula4Evalua As String, lsFormula4Compara As String
    Dim lsFormula5Evalua As String, lsFormula5Compara As String
    Dim lnMontoFormula1Evalua As Currency, lnMontoFormula1Compara As Currency
    Dim lnMontoFormula2Evalua As Currency, lnMontoFormula2Compara As Currency
    Dim lnMontoFormula3Evalua As Currency, lnMontoFormula3Compara As Currency
    Dim lnMontoFormula4Evalua As Currency, lnMontoFormula4Compara As Currency
    Dim lnMontoFormula5Evalua As Currency, lnMontoFormula5Compara As Currency
    Dim lnFilaCabecera As Integer, lbPintoCabecera As Boolean 'EJVG20131230
    
    If psOpeCod = gContRepBaseNotasEstadoSitFinan Then
        lsPath = App.path & "\FormatoCarta\NIIF_NotasEstado_SitFinanc.xls"
        lsArchivo = "NIIF_NotasEstado_SitFinac_" & UCase(gsCodUser) & "_" & Format(gdFecSis, "yyyymmdd") & Format(Now, "hhmmss") & "_" & Format(pdFechaEvalua, "yyyymmdd") & ".xls"
    ElseIf psOpeCod = gContRepBaseNotasEstadoResultado Then
        lsPath = App.path & "\FormatoCarta\NIIF_NotasEstado_Resultado.xls"
        lsArchivo = "NIIF_NotasEstado_Resultado_" & UCase(gsCodUser) & "_" & Format(gdFecSis, "yyyymmdd") & Format(Now, "hhmmss") & "_" & Format(pdFechaEvalua, "yyyymmdd") & ".xls"
    End If
    
    'valida formato carta
    If Len(Dir(lsPath)) = 0 Then
        MsgBox "No se pudo encontrar el archivo: " & lsPath & "," & Chr(10) & "comuniquese con el Dpto. de TI", vbCritical, "Aviso"
        Exit Sub
    End If
    'verifica formato carta abierto
    If fs.FileExists(lsPath) Then
        lbAbierto = True
        Do While lbAbierto
            If ArchivoEstaAbierto(lsPath) Then
                lbAbierto = True
                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(lsPath) + " para continuar", vbRetryCancel, "Aviso") = vbCancel Then
                    Exit Sub
                End If
            Else
                lbAbierto = False
            End If
        Loop
    End If
    
    Set rsNotas = oRep.RecuperaConfigRepNotasEstado(psOpeCod)
    If Not RSVacio(rsNotas) Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(lsPath)
        Set xlsHoja = xlsLibro.ActiveSheet
        
        BarraProgreso.value = 0
        BarraProgreso.Min = 0
        BarraProgreso.Max = rsNotas.RecordCount
        BarraProgreso.value = 0
        EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
        
        lsNombreMesEvalua = dameNombreMes(Month(pdFechaEvalua), True)
        lsNombreMesCompara = dameNombreMes(Month(pdFechaCompara), True)
        lnAnioEvalua = Year(pdFechaEvalua)
        lnAnioCompara = Year(pdFechaCompara)
        
        lnNivelMax = oRep.ObtenerUltimoNivelConfig(psOpeCod)
        lnUltimaColumna = lnNivelMax + IIf(pbComparaPeriodo, 10, 5) '5 por las formulas evalua y 10 con todo las comparativas
        
        lnFilaActual = 8
        lnColumnaActual = 1
        
        xlsHoja.Range(xlsHoja.Cells(lnFilaActual, lnColumnaActual), xlsHoja.Cells(lnFilaActual, lnNivelMax - 1)).ColumnWidth = 5 'Niveles
        xlsHoja.Range(xlsHoja.Cells(lnFilaActual, lnNivelMax), xlsHoja.Cells(lnFilaActual, lnNivelMax)).ColumnWidth = 50 'Ult Nivel
        xlsHoja.Range(Mid(xlsHoja.Cells(lnFilaActual, lnColumnaActual).AddressLocal, 2, 1) & ":" & Mid(xlsHoja.Cells(lnFilaActual, lnNivelMax).AddressLocal, 2, 1)).HorizontalAlignment = xlLeft
        xlsHoja.Range(xlsHoja.Cells(lnFilaActual, lnNivelMax + 1), xlsHoja.Cells(lnFilaActual, lnNivelMax + 6)).ColumnWidth = 15 'Formulas
        xlsHoja.Range(Mid(xlsHoja.Cells(lnFilaActual, lnNivelMax + 1).AddressLocal, 2, 1) & ":" & Mid(xlsHoja.Cells(lnFilaActual, lnNivelMax + 6).AddressLocal, 2, 1)).HorizontalAlignment = xlRight
        xlsHoja.Range("A:" & IIf(pbComparaPeriodo, "O", "J")).NumberFormat = "#,##0"
        
        xlsHoja.Cells(5, 1) = "AL " & Format(Day(pdFechaEvalua), "00") & " DE " & lsNombreMesEvalua & " " & lnAnioEvalua
        For iCab = 1 To rsNotas.RecordCount
            'Notas
            xlsHoja.Cells(lnFilaActual, lnColumnaActual) = "NOTA N° " & Format(iCab, "00") & " " & rsNotas!cDescripcion
            If rsNotas!bPeriodo Then
                xlsHoja.Cells(lnFilaActual, lnUltimaColumna - IIf(pbComparaPeriodo, 6, 1)) = lsNombreMesEvalua & " - " & lnAnioEvalua
                If pbComparaPeriodo Then
                    xlsHoja.Cells(lnFilaActual, lnUltimaColumna - 1) = lsNombreMesCompara & " - " & lnAnioCompara
                End If
            End If
            If rsNotas!cFormula <> "" Then
                xlsHoja.Cells(lnFilaActual, lnUltimaColumna - IIf(pbComparaPeriodo, 4, 0)) = Format(ObtenerResultadoFormula(pdFechaEvalua, rsNotas!cFormula, pnMoneda), gsFormatoNumeroView)
                If pbComparaPeriodo Then
                    xlsHoja.Cells(lnFilaActual, lnUltimaColumna) = Format(ObtenerResultadoFormula(pdFechaCompara, rsNotas!cFormula, pnMoneda), gsFormatoNumeroView)
                End If
            End If
            xlsHoja.Cells(lnFilaActual, lnColumnaActual).RowHeight = 20
            xlsHoja.Range(xlsHoja.Cells(lnFilaActual, lnColumnaActual), xlsHoja.Cells(lnFilaActual, lnUltimaColumna)).Interior.Color = RGB(255, 0, 0)
            xlsHoja.Range(xlsHoja.Cells(lnFilaActual, lnColumnaActual), xlsHoja.Cells(lnFilaActual, lnUltimaColumna)).Font.Color = RGB(255, 255, 255)
            xlsHoja.Range(xlsHoja.Cells(lnFilaActual, lnColumnaActual), xlsHoja.Cells(lnFilaActual, lnUltimaColumna)).Font.Bold = True
            
            'EJVG20131230 ***
            lnFilaCabecera = lnFilaActual
            lbPintoCabecera = False
            'END EJVG *******
            lnFilaActual = lnFilaActual + 2
            'Detalle
            Set rsNotasDet = oRep.RecuperaConfigRepNotasEstadoDetalle(fsOpeCod, rsNotas!nId, rsNotas!nNotaEstado)
            For iDet = 1 To rsNotasDet.RecordCount
                lsFormula1Evalua = "": lsFormula1Compara = ""
                lsFormula2Evalua = "": lsFormula2Compara = ""
                lsFormula3Evalua = "": lsFormula3Compara = ""
                lsFormula4Evalua = "": lsFormula4Compara = ""
                lsFormula5Evalua = "": lsFormula5Compara = ""
                lnMontoFormula1Evalua = 0: lnMontoFormula1Compara = 0
                lnMontoFormula2Evalua = 0: lnMontoFormula2Compara = 0
                lnMontoFormula3Evalua = 0: lnMontoFormula3Compara = 0
                lnMontoFormula4Evalua = 0: lnMontoFormula4Compara = 0
                lnMontoFormula5Evalua = 0: lnMontoFormula5Compara = 0
            
                lnColumnaActual = rsNotasDet!nNivel
                If rsNotasDet!nTpoDetalle = 2 Then 'Formula
                    'Formula 1
                    lsFormula1Evalua = Trim(IIf(lnAnioEvalua <= 2012, rsNotasDet!cFormula1_2012, rsNotasDet!cFormula1))
                    If lsFormula1Evalua <> "" Then
                        lnMontoFormula1Evalua = ObtenerResultadoFormula(pdFechaEvalua, lsFormula1Evalua, pnMoneda)
                    End If
                    If pbComparaPeriodo Then
                        lsFormula1Compara = Trim(IIf(lnAnioCompara <= 2012, rsNotasDet!cFormula1_2012, rsNotasDet!cFormula1))
                        If lsFormula1Compara <> "" Then
                            lnMontoFormula1Compara = ObtenerResultadoFormula(pdFechaCompara, lsFormula1Compara, pnMoneda)
                        End If
                    End If
                    'Formula 2
                    lsFormula2Evalua = Trim(IIf(lnAnioEvalua <= 2012, rsNotasDet!cFormula2_2012, rsNotasDet!cFormula2))
                    If lsFormula2Evalua <> "" Then
                        lnMontoFormula2Evalua = ObtenerResultadoFormula(pdFechaEvalua, lsFormula2Evalua, pnMoneda)
                    End If
                    If pbComparaPeriodo Then
                        lsFormula2Compara = Trim(IIf(lnAnioCompara <= 2012, rsNotasDet!cFormula2_2012, rsNotasDet!cFormula2))
                        If lsFormula2Compara <> "" Then
                            lnMontoFormula2Compara = ObtenerResultadoFormula(pdFechaCompara, lsFormula2Compara, pnMoneda)
                        End If
                    End If
                    'Formula 3
                    lsFormula3Evalua = Trim(IIf(lnAnioEvalua <= 2012, rsNotasDet!cFormula3_2012, rsNotasDet!cFormula3))
                    If lsFormula3Evalua <> "" Then
                        lnMontoFormula3Evalua = ObtenerResultadoFormula(pdFechaEvalua, lsFormula3Evalua, pnMoneda)
                    End If
                    If pbComparaPeriodo Then
                        lsFormula3Compara = Trim(IIf(lnAnioCompara <= 2012, rsNotasDet!cFormula3_2012, rsNotasDet!cFormula3))
                        If lsFormula3Compara <> "" Then
                            lnMontoFormula3Compara = ObtenerResultadoFormula(pdFechaCompara, lsFormula3Compara, pnMoneda)
                        End If
                    End If
                    'Formula 4
                    lsFormula4Evalua = Trim(IIf(lnAnioEvalua <= 2012, rsNotasDet!cFormula4_2012, rsNotasDet!cFormula4))
                    If lsFormula4Evalua <> "" Then
                        lnMontoFormula4Evalua = ObtenerResultadoFormula(pdFechaEvalua, lsFormula4Evalua, pnMoneda)
                    End If
                    If pbComparaPeriodo Then
                        lsFormula4Compara = Trim(IIf(lnAnioCompara <= 2012, rsNotasDet!cFormula4_2012, rsNotasDet!cFormula4))
                        If lsFormula4Compara <> "" Then
                            lnMontoFormula4Compara = ObtenerResultadoFormula(pdFechaCompara, lsFormula4Compara, pnMoneda)
                        End If
                    End If
                    'Formula 5
                    lsFormula5Evalua = Trim(IIf(lnAnioEvalua <= 2012, rsNotasDet!cFormula5_2012, rsNotasDet!cFormula5))
                    If lsFormula5Evalua <> "" Then
                        lnMontoFormula5Evalua = ObtenerResultadoFormula(pdFechaEvalua, lsFormula5Evalua, pnMoneda)
                    End If
                    If pbComparaPeriodo Then
                        lsFormula5Compara = Trim(IIf(lnAnioCompara <= 2012, rsNotasDet!cFormula5_2012, rsNotasDet!cFormula5))
                        If lsFormula5Compara <> "" Then
                            lnMontoFormula5Compara = ObtenerResultadoFormula(pdFechaCompara, lsFormula5Compara, pnMoneda)
                        End If
                    End If
                    '***********************
                    If lnMontoFormula1Evalua <> 0 Or lnMontoFormula1Compara <> 0 Or lnMontoFormula2Evalua <> 0 Or lnMontoFormula2Compara <> 0 Or lnMontoFormula3Evalua <> 0 Or lnMontoFormula3Compara <> 0 Or lnMontoFormula4Evalua <> 0 Or lnMontoFormula4Compara <> 0 Or lnMontoFormula5Evalua <> 0 Or lnMontoFormula5Compara <> 0 Then
                        If lnMontoFormula1Evalua <> 0 Then
                            xlsHoja.Cells(lnFilaActual, lnUltimaColumna - IIf(pbComparaPeriodo, 8, 4)) = Format(lnMontoFormula1Evalua, gsFormatoNumeroView)
                        End If
                        If pbComparaPeriodo And lnMontoFormula1Compara <> 0 Then
                            xlsHoja.Cells(lnFilaActual, lnUltimaColumna - 4) = Format(lnMontoFormula1Compara, gsFormatoNumeroView)
                        End If
                        If lnMontoFormula2Evalua <> 0 Then
                            xlsHoja.Cells(lnFilaActual, lnUltimaColumna - IIf(pbComparaPeriodo, 7, 3)) = Format(lnMontoFormula2Evalua, gsFormatoNumeroView)
                        End If
                        If pbComparaPeriodo And lnMontoFormula2Compara <> 0 Then
                            xlsHoja.Cells(lnFilaActual, lnUltimaColumna - 3) = Format(lnMontoFormula2Compara, gsFormatoNumeroView)
                        End If
                        If lnMontoFormula3Evalua <> 0 Then
                            xlsHoja.Cells(lnFilaActual, lnUltimaColumna - IIf(pbComparaPeriodo, 6, 2)) = Format(lnMontoFormula3Evalua, gsFormatoNumeroView)
                        End If
                        If pbComparaPeriodo And lnMontoFormula3Compara <> 0 Then
                            xlsHoja.Cells(lnFilaActual, lnUltimaColumna - 2) = Format(lnMontoFormula3Compara, gsFormatoNumeroView)
                        End If
                        If lnMontoFormula4Evalua <> 0 Then
                            xlsHoja.Cells(lnFilaActual, lnUltimaColumna - IIf(pbComparaPeriodo, 5, 1)) = Format(lnMontoFormula4Evalua, gsFormatoNumeroView)
                        End If
                        If pbComparaPeriodo And lnMontoFormula4Compara <> 0 Then
                            xlsHoja.Cells(lnFilaActual, lnUltimaColumna - 1) = Format(lnMontoFormula4Compara, gsFormatoNumeroView)
                        End If
                        If lnMontoFormula5Evalua <> 0 Then
                            xlsHoja.Cells(lnFilaActual, lnUltimaColumna - IIf(pbComparaPeriodo, 4, 0)) = Format(lnMontoFormula5Evalua, gsFormatoNumeroView)
                        End If
                        If pbComparaPeriodo And lnMontoFormula5Compara <> 0 Then
                            xlsHoja.Cells(lnFilaActual, lnUltimaColumna) = Format(lnMontoFormula5Compara, gsFormatoNumeroView)
                        End If
                        
                        'EJVG20131230 ***
                        xlsHoja.Cells(lnFilaActual, lnUltimaColumna + 1) = "'" & Trim(lsFormula1Evalua & " " & lsFormula2Evalua & " " & lsFormula3Evalua & " " & lsFormula4Evalua & " " & lsFormula5Evalua)
                        If pbComparaPeriodo Then
                            xlsHoja.Cells(lnFilaActual, lnUltimaColumna + 2) = "'" & Trim(lsFormula1Compara & " " & lsFormula2Compara & " " & lsFormula3Compara & " " & lsFormula4Compara & " " & lsFormula5Compara)
                        End If
                        If Not lbPintoCabecera Then
                            xlsHoja.Cells(lnFilaCabecera, lnUltimaColumna + 1) = "FORMULA " & lsNombreMesEvalua & " - " & lnAnioEvalua
                            If pbComparaPeriodo Then
                                xlsHoja.Cells(lnFilaCabecera, lnUltimaColumna + 2) = "FORMULA " & lsNombreMesCompara & " - " & lnAnioCompara
                            End If
                            xlsHoja.Range(xlsHoja.Cells(lnFilaCabecera, lnUltimaColumna + 1), IIf(pbComparaPeriodo, xlsHoja.Cells(lnFilaCabecera, lnUltimaColumna + 2), xlsHoja.Cells(lnFilaCabecera, lnUltimaColumna + 1))).Interior.Color = RGB(255, 0, 0)
                            xlsHoja.Range(xlsHoja.Cells(lnFilaCabecera, lnUltimaColumna + 1), IIf(pbComparaPeriodo, xlsHoja.Cells(lnFilaCabecera, lnUltimaColumna + 2), xlsHoja.Cells(lnFilaCabecera, lnUltimaColumna + 1))).Font.Color = RGB(255, 255, 255)
                            xlsHoja.Range(xlsHoja.Cells(lnFilaCabecera, lnUltimaColumna + 1), IIf(pbComparaPeriodo, xlsHoja.Cells(lnFilaCabecera, lnUltimaColumna + 2), xlsHoja.Cells(lnFilaCabecera, lnUltimaColumna + 1))).Font.Bold = True
                            lbPintoCabecera = True
                        End If
                        'END EJVG *******
                        
                        xlsHoja.Cells(lnFilaActual, lnColumnaActual) = rsNotasDet!cDescripcion
                        If rsNotasDet!bNegrita Then xlsHoja.Cells(lnFilaActual, lnColumnaActual).Font.Bold = True
                        lnFilaActual = lnFilaActual + 1
                    End If
                Else
                    xlsHoja.Range(xlsHoja.Cells(lnFilaActual, lnColumnaActual), xlsHoja.Cells(lnFilaActual, lnUltimaColumna)).MergeCells = True
                    xlsHoja.Cells(lnFilaActual, lnColumnaActual) = rsNotasDet!cDescripcion
                    If rsNotasDet!bNegrita Then xlsHoja.Cells(lnFilaActual, lnColumnaActual).Font.Bold = True
                    lnFilaActual = lnFilaActual + 1
                End If
                rsNotasDet.MoveNext
            Next
            'Comentario
            lnColumnaActual = 1
            If rsNotas!bComentario Then
                xlsHoja.Cells(lnFilaActual, lnColumnaActual) = "Comentario"
                xlsHoja.Cells(lnFilaActual, lnColumnaActual).Font.Bold = True
                lnFilaActual = lnFilaActual + 1
                xlsHoja.Range(xlsHoja.Cells(lnFilaActual, lnColumnaActual), xlsHoja.Cells(lnFilaActual, lnUltimaColumna)).MergeCells = True
                xlsHoja.Range(xlsHoja.Cells(lnFilaActual, lnColumnaActual), xlsHoja.Cells(lnFilaActual, lnUltimaColumna)).Borders.Weight = xlThin
            End If
            rsNotas.MoveNext
            lnFilaActual = lnFilaActual + 2
            
            BarraProgreso.value = iCab
            EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
        Next
        xlsHoja.SaveAs App.path & "\Spooler\" & lsArchivo
        xlsAplicacion.Visible = True
        xlsAplicacion.Windows(1).Visible = True
        EstadoBarra.Panels(1) = "Proceso Terminado"
    Else
        MsgBox "No existe la configuración respectiva para generar el presente Reporte", vbInformation, "Aviso"
    End If
    
    Set rsNotas = Nothing
    Set rsNotasDet = Nothing
    Set oRep = Nothing
    Set xlsHoja = Nothing
    Set xlsLibro = Nothing
    Set xlsAplicacion = Nothing
End Sub
Public Function ObtenerResultadoFormula(ByVal pdFecha As Date, ByVal psFormula As String, ByVal pnMoneda As Integer) As Currency
    Dim oBal As New DbalanceCont
    Dim oFormula As New NInterpreteFormula
    Dim lsFormula As String, lsTmp As String, lsTmp1 As String, lsCadFormula As String
    Dim MatDatos() As TCtaCont
    Dim i As Long, j As Long, nCtaCont As Long
    
    lsFormula = Trim(psFormula)
    ReDim MatDatos(0)
    nCtaCont = 0

    For i = 1 To Len(lsFormula)
        If (Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Then
            lsTmp = lsTmp + Mid(Trim(lsFormula), i, 1)
        Else
            If Len(lsTmp) > 0 Then
                nCtaCont = nCtaCont + 1
                ReDim Preserve MatDatos(nCtaCont)
                MatDatos(nCtaCont).CuentaContable = lsTmp
                MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual2(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True)
            End If
            lsTmp = ""
        End If
    Next i
    If Len(lsTmp) > 0 Then
        nCtaCont = nCtaCont + 1
        ReDim Preserve MatDatos(nCtaCont)
        MatDatos(nCtaCont).CuentaContable = lsTmp
        MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual2(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True)
    End If
    'Genero la formula en cadena
    lsTmp = ""
    lsCadFormula = ""
    For i = 1 To Len(lsFormula)
        If (Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Then
            lsTmp = lsTmp + Mid(Trim(lsFormula), i, 1)
        Else
            If Len(lsTmp) > 0 Then
                For j = 1 To nCtaCont
                    If MatDatos(j).CuentaContable = lsTmp Then
                        lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
                        Exit For
                    End If
                Next j
            End If
            lsTmp = ""
            lsCadFormula = lsCadFormula & Mid(Trim(lsFormula), i, 1)
        End If
    Next
    If Len(lsTmp) > 0 Then
        For j = 1 To nCtaCont
           If MatDatos(j).CuentaContable = lsTmp Then
               lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
               Exit For
           End If
        Next j
    End If
    
    ObtenerResultadoFormula = oFormula.ExprANum(lsCadFormula)
    Set oBal = Nothing
    Set oFormula = Nothing
End Function
