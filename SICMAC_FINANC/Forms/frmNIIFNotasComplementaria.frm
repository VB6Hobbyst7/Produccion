VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmNIIFNotasComplementaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas Complementarias"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   Icon            =   "frmNIIFNotasComplementaria.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
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
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.Frame fraReportes 
      Caption         =   "Reportes"
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
      Height          =   2355
      Left            =   80
      TabIndex        =   24
      Top             =   1440
      Width           =   5055
      Begin VB.CheckBox chkReporte 
         Caption         =   "Concentración de Riesgos por Sector"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   3255
      End
      Begin VB.CheckBox chkReporte 
         Caption         =   "Créditos vencidos por Días de atraso"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   3255
      End
      Begin VB.CheckBox chkReporte 
         Caption         =   "Calidad Crediticia de Activos Financieros"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CheckBox chkReporte 
         Caption         =   "Exposición al Riesgo de Liquidez"
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CheckBox chkReporte 
         Caption         =   "Valor Razonable y Valor en Libro"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2775
      End
      Begin VB.CheckBox chkReporte 
         Caption         =   "Exposición Riesgo Cambiario"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox chkReporte 
         Caption         =   "Exposición Máx. Riesgo Crediticio"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox chkReporte 
         Caption         =   "Instrumentos Financieros"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   0
         Width           =   855
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
      TabIndex        =   21
      Top             =   720
      Width           =   5055
      Begin VB.TextBox txtAnioCompara 
         Alignment       =   1  'Right Justify
         Height          =   280
         Left            =   720
         MaxLength       =   4
         TabIndex        =   3
         Top             =   240
         Width           =   780
      End
      Begin VB.ComboBox cboMesCompara 
         Height          =   315
         ItemData        =   "frmNIIFNotasComplementaria.frx":030A
         Left            =   2760
         List            =   "frmNIIFNotasComplementaria.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   2280
         TabIndex        =   22
         Top             =   270
         Width           =   390
      End
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
      TabIndex        =   18
      Top             =   0
      Width           =   5055
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmNIIFNotasComplementaria.frx":030E
         Left            =   2760
         List            =   "frmNIIFNotasComplementaria.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   280
         Left            =   720
         MaxLength       =   4
         TabIndex        =   0
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   2280
         TabIndex        =   20
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   270
         Width           =   375
      End
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
      Left            =   2760
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   3885
      Width           =   1455
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
      Left            =   1200
      TabIndex        =   14
      ToolTipText     =   "Generar Reporte"
      Top             =   3885
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar BarraProgreso 
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   4320
      Width           =   3730
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar EstadoBarra 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   4290
      Width           =   5205
      _ExtentX        =   9181
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
End
Attribute VB_Name = "frmNIIFNotasComplementaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************
'** Nombre : frmNIIFNotasComplementaria
'** Descripción : Generación de Reportes Notas Complementaria creado segun ERS149-2013
'** Creación : EJVG, 20140103 11:00:00 AM
'*************************************************************************************
Option Explicit
Private Type NotaComplementaType
    nSeccion As Integer
    nNotaEstado As Integer
    cDescripcion As String
    nColumna As Integer
    cFormulaEvalua As String
    nSaldoEvaluaUnificado As Currency
    nSaldoEvaluaSoles As Currency
    nSaldoEvaluaDolares As Currency
    cFormulaCompara As String
    nSaldoComparaUnificado As Currency
    nSaldoComparaSoles As Currency
    nSaldoComparaDolares As Currency
End Type
Dim Notas() As NotaComplementaType


Private Sub Form_Load()
    cargarMes
    txtAnio.Text = Year(gdFecSis)
    txtAnioCompara.Text = Year(gdFecSis) - 1
    cboMes.ListIndex = IndiceListaCombo(cboMes, Month(gdFecSis))
    cboMesCompara.ListIndex = IndiceListaCombo(cboMesCompara, Month(gdFecSis))
End Sub
Private Sub chkTodos_Click()
    Dim i As Integer
    For i = 0 To chkReporte.Count - 1
        If chkReporte.Item(i).Visible And chkReporte.Item(i).Enabled Then
            chkReporte.Item(i).value = chkTodos.value
        End If
    Next
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Function validaGenerar() As Boolean
    Dim i As Integer
    Dim bSelecciona As Boolean
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
    For i = 0 To chkReporte.Count - 1
        If chkReporte.Item(i).value = 1 Then
            bSelecciona = True
            Exit For
        End If
    Next
    If Not bSelecciona Then
        MsgBox "Ud. debe seleccionar al menos un reporte para continuar", vbInformation, "Aviso"
        validaGenerar = False
        chkTodos.SetFocus
        Exit Function
    End If
End Function
Private Sub cargarMes()
    cboMes.Clear
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
    
    cboMesCompara.Clear
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
Private Sub chkPeriodoCompara_Click()
    fraPerComparar.Enabled = IIf(chkPeriodoCompara.value = 1, True, False)
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cboMes.SetFocus
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
Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkPeriodoCompara.SetFocus
    End If
End Sub
Private Sub cboMes_LostFocus()
    If cboMes.ListIndex > -1 Then
        cboMesCompara.ListIndex = cboMes.ListIndex
    End If
End Sub
Private Sub chkPeriodoCompara_KeyPress(KeyAscii As Integer)
    If chkPeriodoCompara.value = vbChecked Then
        txtAnioCompara.SetFocus
    Else
        chkTodos.SetFocus
    End If
End Sub
Private Sub cboMesCompara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkTodos.SetFocus
    End If
End Sub
Private Sub chkTodos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub
Private Sub cmdGenerar_Click()
    If Not validaGenerar Then Exit Sub
    Dim oNRep As New NRepFormula
    Dim objNotaEstado As New frmNIIFNotasEstado
    Dim rsRep As New ADODB.Recordset
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet, xlHoja1 As Excel.Worksheet
    Dim lbFechaCompara As Boolean
    Dim ldFechaEvalua As Date, ldFechaCompara As Date
    Dim oConstSis As New DConstSistemas
    Dim lnTpoCambio As Currency
    Dim lsArchivo As String
    Dim iNota As Integer
    Dim bBuscaSoles As Boolean, bBuscaDolares As Boolean
    Dim bBuscaNotaComplementaria As Boolean
    
    On Error GoTo ErrGenerar
    
    Screen.MousePointer = 11
    lsArchivo = "\spooler\RptNotaComplementaria" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    ldFechaEvalua = obtenerFechaFinMes(cboMes.ListIndex + 1, txtAnio.Text)
    lbFechaCompara = IIf(chkPeriodoCompara.value = 1, True, False)
    ldFechaCompara = obtenerFechaFinMes(cboMesCompara.ListIndex + 1, txtAnioCompara.Text)
    
    If chkReporte.Item(0).value = vbChecked Or chkReporte.Item(1).value = vbChecked Or chkReporte.Item(2).value = vbChecked Or chkReporte.Item(3).value = vbChecked Then
        bBuscaNotaComplementaria = True
    End If
    If chkReporte.Item(2).value = vbChecked Then
        bBuscaSoles = True
        bBuscaDolares = True
    End If
    
    EstadoBarra.Panels(1) = "Load Data 0.00%"
    
    If bBuscaNotaComplementaria Then
        'Cargamos los datos en memoria para agilizar los reportes
        Set rsRep = oNRep.ListaNotaComplementariasAll(gContRepBaseNotasEstadoSitFinan)
        BarraProgreso.value = 0
        BarraProgreso.Min = 0
        BarraProgreso.Max = rsRep.RecordCount
        BarraProgreso.value = 0
        ReDim Notas(0)
        Do While Not rsRep.EOF
            iNota = UBound(Notas) + 1
            ReDim Preserve Notas(iNota)
            Notas(iNota).nSeccion = rsRep!nSeccion
            Notas(iNota).nNotaEstado = rsRep!nNotaEstado
            Notas(iNota).cDescripcion = rsRep!cDescripcion
            Notas(iNota).nColumna = rsRep!nColumna
            Notas(iNota).cFormulaEvalua = IIf(Year(ldFechaEvalua) <= 2012, rsRep!cFormula_2012, rsRep!cFormula)
            Notas(iNota).nSaldoEvaluaUnificado = objNotaEstado.ObtenerResultadoFormula(ldFechaEvalua, Notas(iNota).cFormulaEvalua, 0)
            If bBuscaSoles Then
                Notas(iNota).nSaldoEvaluaSoles = objNotaEstado.ObtenerResultadoFormula(ldFechaEvalua, Notas(iNota).cFormulaEvalua, 1)
            End If
            If bBuscaDolares Then
                Notas(iNota).nSaldoEvaluaDolares = objNotaEstado.ObtenerResultadoFormula(ldFechaEvalua, Notas(iNota).cFormulaEvalua, 2)
            End If
            If lbFechaCompara Then
                Notas(iNota).cFormulaCompara = IIf(Year(ldFechaCompara) <= 2012, rsRep!cFormula_2012, rsRep!cFormula)
                Notas(iNota).nSaldoComparaUnificado = objNotaEstado.ObtenerResultadoFormula(ldFechaCompara, Notas(iNota).cFormulaCompara, 0)
                If bBuscaSoles Then
                    Notas(iNota).nSaldoComparaSoles = objNotaEstado.ObtenerResultadoFormula(ldFechaCompara, Notas(iNota).cFormulaCompara, 1)
                End If
                If bBuscaDolares Then
                    Notas(iNota).nSaldoComparaDolares = objNotaEstado.ObtenerResultadoFormula(ldFechaCompara, Notas(iNota).cFormulaCompara, 2)
                End If
            End If
            rsRep.MoveNext
            DoEvents
            BarraProgreso.value = iNota
            EstadoBarra.Panels(1) = "Load Data: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
        Loop
    End If
    
    EstadoBarra.Panels(1) = "Generando Reportes"
    BarraProgreso.value = 0
    BarraProgreso.Min = 0
    BarraProgreso.Max = 8
    BarraProgreso.value = 0
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    
    Set xlsLibro = xlsAplicacion.Workbooks.Add
    'Instrumentos Financieros
    If chkReporte.Item(0).value = vbChecked Then
        Set xlsHoja = xlsLibro.Worksheets.Add
        xlsHoja.Name = "Instrumentos Financieros"
        generaHojaInstrumentosFinancieros xlsHoja, ldFechaEvalua, lbFechaCompara, ldFechaCompara
    End If
    BarraProgreso.value = 1
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    DoEvents
    'Exposición Máx. Riesgo Crediticio
    If chkReporte.Item(1).value = vbChecked Then
        Set xlsHoja = xlsLibro.Worksheets.Add
        xlsHoja.Name = "Exposicion Max. Riesgo Credit."
        generaHojaExposicionMaxRiesgoCrediticio xlsHoja, ldFechaEvalua, lbFechaCompara, ldFechaCompara
    End If
    BarraProgreso.value = 2
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    DoEvents
    'Exposición Riesgo Cambiario
    If chkReporte.Item(2).value = vbChecked Then
        Set xlsHoja = xlsLibro.Worksheets.Add
        xlsHoja.Name = "Exposición Riesgo Cambiario"
        generaHojaExposicionRiesgoCambiario xlsHoja, ldFechaEvalua, lbFechaCompara, ldFechaCompara
    End If
    BarraProgreso.value = 3
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    DoEvents
    'Valor Razonable y Valor en Libro
    If chkReporte.Item(3).value = vbChecked Then
        Set xlsHoja = xlsLibro.Worksheets.Add
        xlsHoja.Name = "Valor Raz. y Valor en Libro"
        generaHojaValorRazonableValorLibro xlsHoja, ldFechaEvalua, lbFechaCompara, ldFechaCompara
    End If
    BarraProgreso.value = 4
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    DoEvents
    'Exposición al Riesgo de Liquidez
    'If chkReporte.Item(4).value = vbChecked Then
    '    Set xlsHoja = xlsLibro.Worksheets.Add
    '    xlsHoja.Name = "Exposic. al Riesgo de Liquidez"
    '    generaHojaExposicionRiesgoLiquidez xlsHoja, ldFechaEvalua, lbFechaCompara, ldFechaCompara
    'End If
    BarraProgreso.value = 5
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    DoEvents
    'Calidad Crediticia de Activos Financieros
    If chkReporte.Item(5).value = vbChecked Then
        Set xlsHoja = xlsLibro.Worksheets.Add
        xlsHoja.Name = "Calidad Credit. de Act. Financ."
        If MsgBox("¿Desea generar el Resumen o Detalle del Reporte de Calidad Crediticia de Activos Financieros?" & Chr(13) & "Presione [SI] para generar el Resumen" & Chr(13) & "Presione [NO] para generar en Detalle", vbInformation + vbYesNo, "Aviso") = vbYes Then
            generaHojaCalidadCrediticiaActivosFinancieros xlsHoja, ldFechaEvalua, lbFechaCompara, ldFechaCompara, True
        Else
            generaHojaCalidadCrediticiaActivosFinancieros xlsHoja, ldFechaEvalua, lbFechaCompara, ldFechaCompara, False
        End If
    End If
    BarraProgreso.value = 6
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    DoEvents
    'Créditos venidos por Días de atraso
    If chkReporte.Item(6).value = vbChecked Then
        Set xlsHoja = xlsLibro.Worksheets.Add
        xlsHoja.Name = "Cred. venc. por dias de atraso"
        generaHojaCalidadCreditosVencidosxDiasAtraso xlsHoja, ldFechaEvalua, lbFechaCompara, ldFechaCompara
    End If
    BarraProgreso.value = 7
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    DoEvents
    'Concentración de Riesgos por Sector
    If chkReporte.Item(7).value = vbChecked Then
        Set xlsHoja = xlsLibro.Worksheets.Add
        xlsHoja.Name = "Concentrac. de Riesgos x Sector"
        generaHojaConcentracionRiesgoxSector xlsHoja, ldFechaEvalua, lbFechaCompara, ldFechaCompara
    End If
    BarraProgreso.value = 8
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    
    For Each xlHoja1 In xlsLibro.Worksheets
        If UCase(xlHoja1.Name) = "HOJA1" Or UCase(xlHoja1.Name) = "HOJA2" Or UCase(xlHoja1.Name) = "HOJA3" Then
            xlHoja1.Delete
        End If
    Next
    
    MsgBox "Se ha generado satisfactoriamente el reporte de Notas Complementarias Info. Anual", vbInformation, "Aviso"
    xlsHoja.SaveAs App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True

    Screen.MousePointer = 0
    Set objNotaEstado = Nothing
    Set rsRep = Nothing
    Set oNRep = Nothing
    Set xlHoja1 = Nothing
    Set xlsHoja = Nothing
    Set xlsLibro = Nothing
    Set xlsAplicacion = Nothing
    Exit Sub
ErrGenerar:
    Screen.MousePointer = 0
    BarraProgreso.value = 0
    EstadoBarra.Panels(1) = "Proceso: " & "0.00%"
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub generaHojaInstrumentosFinancieros(ByRef xlsHoja As Worksheet, ByVal pdFechaEvalua As Date, ByVal pbFechaCompara As Boolean, ByVal pdFechaCompara As Date)
    Dim iNota As Integer
    Dim lnFilaActual As Integer, lnFilaAnterior As Integer
    Dim lsColEvalua As String, lsColCompara As String
  
    xlsHoja.Cells.EntireColumn.WrapText = True
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 8
    xlsHoja.Cells.VerticalAlignment = xlCenter
    
    xlsHoja.Columns.ColumnWidth = 12
    xlsHoja.Columns("A:A").ColumnWidth = 50
    xlsHoja.Columns("B:B").ColumnWidth = 0
    
    xlsHoja.Range("A5:A6").RowHeight = 45.75
    xlsHoja.Range("C3", "I3").MergeCells = True
    xlsHoja.Range("C3").NumberFormat = "mmm-yy"
    xlsHoja.Range("C3") = pdFechaEvalua
    xlsHoja.Range("C4", "I4").MergeCells = True
    xlsHoja.Range("C4") = "Activos Financieros"
    xlsHoja.Range("C5") = "A valor razonable con cambios en resultados"
    xlsHoja.Range("C5", "D5").MergeCells = True
    xlsHoja.Range("C6") = "Por Negociación"
    xlsHoja.Range("D6") = "Designado al momento Inicial"
    xlsHoja.Range("E5") = "Préstamos y Partidas por Cobrar"
    xlsHoja.Range("E5", "E6").MergeCells = True
    xlsHoja.Range("F5") = "Disponible para la Venta"
    xlsHoja.Range("F5", "G5").MergeCells = True
    xlsHoja.Range("F6") = "Al costo Amortizado"
    xlsHoja.Range("G6") = "Al Valor Razonable"
    xlsHoja.Range("H5") = "Mantenido hasta su vencimiento"
    xlsHoja.Range("H5", "H6").MergeCells = True
    xlsHoja.Range("I5") = "Derivados de cobertura"
    xlsHoja.Range("I5", "I6").MergeCells = True
    
    If pbFechaCompara Then
        xlsHoja.Range("C3", "I6").Copy
        xlsHoja.Range("J3").PasteSpecial
    End If
    xlsHoja.Range("C3:P6").HorizontalAlignment = xlCenter
    
    xlsHoja.Cells(7, 1) = "Activos"
    xlsHoja.Cells(7, 1).Font.Bold = True
    xlsHoja.Cells(7, 1).Font.Underline = True
    lnFilaActual = 8
    lnFilaAnterior = lnFilaActual

    For iNota = 1 To UBound(Notas)
        If Notas(iNota).nSeccion = 1 Then
            lsColEvalua = "": lsColCompara = ""
            xlsHoja.Cells(lnFilaActual, 1) = Notas(iNota).cDescripcion
            Select Case Notas(iNota).nColumna
                Case 1
                    lsColEvalua = "C"
                    lsColCompara = "J"
                Case 2
                    lsColEvalua = "D"
                    lsColCompara = "K"
                Case 3
                    lsColEvalua = "E"
                    lsColCompara = "L"
                Case 4
                    lsColEvalua = "F"
                    lsColCompara = "M"
                Case 5
                    lsColEvalua = "G"
                    lsColCompara = "N"
                Case 6
                    lsColEvalua = "H"
                    lsColCompara = "O"
                Case 7
                    lsColEvalua = "I"
                    lsColCompara = "P"
            End Select
            If Len(lsColEvalua) > 0 Or Len(lsColCompara) > 0 Then
                xlsHoja.Range(lsColEvalua & lnFilaActual) = Notas(iNota).nSaldoEvaluaUnificado
                If pbFechaCompara Then
                    xlsHoja.Range(lsColCompara & lnFilaActual) = Notas(iNota).nSaldoComparaUnificado
                End If
            End If
            lnFilaActual = lnFilaActual + 1
        End If
    Next

    xlsHoja.Cells(lnFilaActual, 1) = "Total"
    xlsHoja.Range("C" & lnFilaActual).Formula = "=SUM(C" & 7 & ":C" & lnFilaActual - 1 & ")"
    xlsHoja.Range("C" & lnFilaActual).Copy
    xlsHoja.Range("D" & lnFilaActual, "I" & lnFilaActual).PasteSpecial
    If pbFechaCompara Then
        xlsHoja.Range("C" & lnFilaActual).Copy
        xlsHoja.Range("J" & lnFilaActual, "P" & lnFilaActual).PasteSpecial
    End If
    xlsHoja.Range("C7:P" & lnFilaActual).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFilaActual = lnFilaActual + 1
    lnFilaAnterior = lnFilaActual
    xlsHoja.Range("C" & lnFilaActual, "I" & lnFilaActual).MergeCells = True
    xlsHoja.Range("C" & lnFilaActual) = "Pasivos Financieros"
    lnFilaActual = lnFilaActual + 1
    xlsHoja.Range("A" & lnFilaActual & ":A" & lnFilaActual + 1).RowHeight = 45.75
    xlsHoja.Range("C" & lnFilaActual) = "A valor razonable con cambios en resultados"
    xlsHoja.Range("C" & lnFilaActual + 1) = "Por Negociación"
    xlsHoja.Range("D" & lnFilaActual + 1) = "Designado al momento Inicial"
    xlsHoja.Range("C" & lnFilaActual, "D" & lnFilaActual).MergeCells = True
    xlsHoja.Range("E" & lnFilaActual) = "Al Costo Amortizado"
    xlsHoja.Range("E" & lnFilaActual, "E" & lnFilaActual + 1).MergeCells = True
    xlsHoja.Range("F" & lnFilaActual) = "Disponible para la Venta"
    xlsHoja.Range("F" & lnFilaActual, "F" & lnFilaActual + 1).MergeCells = True
    xlsHoja.Range("F" & lnFilaActual) = "Otros Pasivos"
    xlsHoja.Range("G" & lnFilaActual, "H" & lnFilaActual + 1).MergeCells = True
    xlsHoja.Range("I" & lnFilaActual) = "Derivados de cobertura"
    xlsHoja.Range("I" & lnFilaActual, "I" & lnFilaActual + 1).MergeCells = True
    lnFilaActual = lnFilaActual + 1
    xlsHoja.Range("G" & lnFilaActual) = "Al Valor Razonable"
    
    If pbFechaCompara Then
        xlsHoja.Range("C" & lnFilaAnterior, "I" & lnFilaActual).Copy
        xlsHoja.Range("J" & lnFilaAnterior).PasteSpecial
    End If
    xlsHoja.Range("C" & lnFilaAnterior & ":P" & lnFilaActual).HorizontalAlignment = xlCenter
    
    lnFilaActual = lnFilaActual + 1
    xlsHoja.Cells(lnFilaActual, 1) = "Pasivos"
    xlsHoja.Cells(lnFilaActual, 1).Font.Bold = True
    xlsHoja.Cells(lnFilaActual, 1).Font.Underline = True
    lnFilaActual = lnFilaActual + 1
    lnFilaAnterior = lnFilaActual
    
    For iNota = 1 To UBound(Notas)
        If Notas(iNota).nSeccion = 2 Then
            lsColEvalua = "": lsColCompara = ""
            xlsHoja.Cells(lnFilaActual, 1) = Notas(iNota).cDescripcion
            Select Case Notas(iNota).nColumna
                Case 1
                    lsColEvalua = "C"
                    lsColCompara = "J"
                Case 2
                    lsColEvalua = "D"
                    lsColCompara = "K"
                Case 3
                    lsColEvalua = "E"
                    lsColCompara = "L"
                Case 4
                    lsColEvalua = "F"
                    lsColCompara = "M"
                Case 5
                    lsColEvalua = "I"
                    lsColCompara = "P"
            End Select
            If Len(lsColEvalua) > 0 Or Len(lsColCompara) > 0 Then
                xlsHoja.Range(lsColEvalua & lnFilaActual) = Notas(iNota).nSaldoEvaluaUnificado
                If pbFechaCompara Then
                    xlsHoja.Range(lsColCompara & lnFilaActual) = Notas(iNota).nSaldoComparaUnificado
                End If
            End If
            lnFilaActual = lnFilaActual + 1
        End If
    Next

    xlsHoja.Cells(lnFilaActual, 1) = "Total"
    xlsHoja.Range("C" & lnFilaActual).Formula = "=SUM(C" & lnFilaAnterior - 1 & ":C" & lnFilaActual - 1 & ")"
    xlsHoja.Range("C" & lnFilaActual).Copy
    xlsHoja.Range("D" & lnFilaActual, "I" & lnFilaActual).PasteSpecial
    If pbFechaCompara Then
        xlsHoja.Range("C" & lnFilaActual).Copy
        xlsHoja.Range("J" & lnFilaActual, "P" & lnFilaActual).PasteSpecial
    End If
    xlsHoja.Range("C" & lnFilaAnterior & ":P" & lnFilaActual).NumberFormat = "#,##0.00;-#,##0.00"
    
    xlsHoja.Range("A3", IIf(pbFechaCompara, "P", "I") & lnFilaActual).Borders.LineStyle = xlContinuous
End Sub
Private Sub generaHojaExposicionMaxRiesgoCrediticio(ByRef xlsHoja As Worksheet, ByVal pdFechaEvalua As Date, ByVal pbFechaCompara As Boolean, ByVal pdFechaCompara As Date)
    Dim iNota As Integer
    Dim lnFilaActual As Integer, lnFilaAnterior As Integer
    Dim lnMontoEvalua As Currency, lnMontoCompara As Currency
    Dim lnNroPosNota As Integer

    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 8
    xlsHoja.Cells.VerticalAlignment = xlCenter
    
    xlsHoja.Columns.ColumnWidth = 12
    xlsHoja.Columns("A:A").ColumnWidth = 50
   
    xlsHoja.Range("A1") = "Exposición Máxima al Riesgo de Créditos"
    xlsHoja.Range("B1") = "Notas"
    xlsHoja.Range("C1").NumberFormat = "mmm-yy"
    xlsHoja.Range("C1") = pdFechaEvalua
    
    If pbFechaCompara Then
        xlsHoja.Range("D1").NumberFormat = "mmm-yy"
        xlsHoja.Range("D1") = pdFechaCompara
    End If
    
    xlsHoja.Range("A2") = "ACTIVO"
    xlsHoja.Range("A2").Font.Underline = True
    
    lnFilaActual = 3
    lnFilaAnterior = lnFilaActual
    lnNroPosNota = 0
    For iNota = 1 To UBound(Notas)
        If Notas(iNota).nSeccion = 1 Then
            lnMontoEvalua = 0: lnMontoCompara = 0
            xlsHoja.Range("A" & lnFilaActual) = Notas(iNota).cDescripcion
            lnMontoEvalua = Notas(iNota).nSaldoEvaluaUnificado
            xlsHoja.Range("C" & lnFilaActual) = lnMontoEvalua
            If pbFechaCompara Then
                lnMontoCompara = Notas(iNota).nSaldoComparaUnificado
                xlsHoja.Range("D" & lnFilaActual) = lnMontoCompara
            End If
            If lnMontoEvalua > 0 Or lnMontoCompara > 0 Then
                lnNroPosNota = lnNroPosNota + 1
                xlsHoja.Range("B" & lnFilaActual) = lnNroPosNota
            End If
            lnFilaActual = lnFilaActual + 1
        End If
    Next
    xlsHoja.Cells(lnFilaActual, 1) = "Total"
    xlsHoja.Range("C" & lnFilaActual).Formula = "=SUM(C" & 2 & ":C" & lnFilaActual - 1 & ")"
    If pbFechaCompara Then
        xlsHoja.Range("D" & lnFilaActual).Formula = "=SUM(D" & 2 & ":D" & lnFilaActual - 1 & ")"
    End If
    xlsHoja.Range("C2:D" & lnFilaActual).NumberFormat = "#,##0.00;-#,##0.00"
    xlsHoja.Range("A1", "D1").HorizontalAlignment = xlCenter
    xlsHoja.Range("A1", "D1").Font.Bold = True

    xlsHoja.Range("A1", IIf(pbFechaCompara, "D", "C") & "1").Borders.LineStyle = xlContinuous
    xlsHoja.Range("A" & lnFilaActual, "B" & lnFilaActual).MergeCells = True
    xlsHoja.Range("A2", "A" & lnFilaActual - 1).BorderAround xlContinuous
    xlsHoja.Range("B2", "B" & lnFilaActual - 1).BorderAround xlContinuous
    xlsHoja.Range("C2", "C" & lnFilaActual - 1).BorderAround xlContinuous
    If pbFechaCompara Then
        xlsHoja.Range("D2", "D" & lnFilaActual - 1).BorderAround xlContinuous
    End If
    xlsHoja.Range("A" & lnFilaActual, IIf(pbFechaCompara, "D", "C") & lnFilaActual).Borders.LineStyle = xlContinuous
End Sub
Private Sub generaHojaExposicionRiesgoCambiario(ByRef xlsHoja As Worksheet, ByVal pdFechaEvalua As Date, ByVal pbFechaCompara As Boolean, ByVal pdFechaCompara As Date)
    Dim iNota As Integer
    Dim lnFilaActual As Integer, lnFilaAnterior As Integer
    Dim lnNroPosNota As Integer

    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 8
    xlsHoja.Cells.VerticalAlignment = xlCenter
    
    xlsHoja.Columns.ColumnWidth = 12
    xlsHoja.Columns("A:A").ColumnWidth = 50
    xlsHoja.Range("A2").RowHeight = 33
   
    xlsHoja.Range("A2") = "Exposición al Riesgo de Mercado Riesgo de Tipo de cambio"
    xlsHoja.Range("B1").NumberFormat = "mmm-yy"
    xlsHoja.Range("B1") = pdFechaEvalua
    xlsHoja.Range("B1", "E1").MergeCells = True
    xlsHoja.Range("B2") = "Dólar Estadounidense"
    '''xlsHoja.Range("C2") = "Nuevos Soles" 'marg ers044-2016
    xlsHoja.Range("C2") = StrConv(gcPEN_PLURAL, vbProperCase) 'marg ers044-2016
    xlsHoja.Range("D2") = "Otras Monedas"
    xlsHoja.Range("E2") = "TOTAL"
    xlsHoja.Range("A3") = "Activos Monetarios"
    '''xlsHoja.Range("B3") = "s/.(000)" 'marg ers044-2016
    xlsHoja.Range("B3") = StrConv(gcPEN_SIMBOLO, vbLowerCase) & " (000)"
    '''xlsHoja.Range("C3") = "s/.(000)" 'marg ers044-2016
    xlsHoja.Range("C3") = StrConv(gcPEN_SIMBOLO, vbLowerCase) & " (000)" 'marg ers044-2016
    '''xlsHoja.Range("D3") = "s/.(000)" 'marg ers044-2016
    xlsHoja.Range("D3") = StrConv(gcPEN_SIMBOLO, vbLowerCase) & " (000)" 'marg ers044-2016
    '''xlsHoja.Range("E3") = "s/.(000)" 'marg ers044-2016
    xlsHoja.Range("E3") = StrConv(gcPEN_SIMBOLO, vbLowerCase) & " (000)" 'marg ers044-2016
    xlsHoja.Range("B1", "E3").HorizontalAlignment = xlCenter
    xlsHoja.Range("B1", "E3").VerticalAlignment = xlCenter
    xlsHoja.Range("B1", "E3").WrapText = True
    xlsHoja.Range("A2", "E3").Borders.LineStyle = xlContinuous
    xlsHoja.Range("B1", "E1").Borders.LineStyle = xlContinuous
    
    If pbFechaCompara Then
        xlsHoja.Range("B1", "E3").Copy
        xlsHoja.Range("F1").PasteSpecial
        xlsHoja.Range("F1") = pdFechaCompara
    End If
    
    lnFilaActual = 4
    lnFilaAnterior = lnFilaActual
    For iNota = 1 To UBound(Notas)
        If Notas(iNota).nSeccion = 1 Then
            xlsHoja.Range("A" & lnFilaActual) = Notas(iNota).cDescripcion
            xlsHoja.Range("B" & lnFilaActual) = Notas(iNota).nSaldoEvaluaDolares
            xlsHoja.Range("C" & lnFilaActual) = Notas(iNota).nSaldoEvaluaSoles
            xlsHoja.Range("D" & lnFilaActual) = 0#
            xlsHoja.Range("E" & lnFilaActual).Formula = "=B" & lnFilaActual & "+C" & lnFilaActual & "+D" & lnFilaActual
            If pbFechaCompara Then
                xlsHoja.Range("F" & lnFilaActual) = Notas(iNota).nSaldoComparaDolares
                xlsHoja.Range("G" & lnFilaActual) = Notas(iNota).nSaldoComparaSoles
                xlsHoja.Range("H" & lnFilaActual) = 0#
                xlsHoja.Range("I" & lnFilaActual).Formula = "=F" & lnFilaActual & "+G" & lnFilaActual & "+H" & lnFilaActual
            End If
            lnFilaActual = lnFilaActual + 1
        End If
    Next
    If lnFilaAnterior = lnFilaActual Then
        lnFilaActual = lnFilaActual + 1
    End If
    xlsHoja.Range("A" & lnFilaAnterior, "A" & lnFilaActual - 1).BorderAround xlContinuous
    xlsHoja.Range("A" & lnFilaAnterior, "A" & lnFilaActual - 1).Copy
    xlsHoja.Range("B" & lnFilaAnterior, IIf(pbFechaCompara, "I", "E") & lnFilaActual - 1).PasteSpecial xlPasteFormats
    xlsHoja.Range("A" & lnFilaActual) = "Total Activos Monetarios"
    xlsHoja.Range("A" & lnFilaActual, "B" & lnFilaActual).Borders.LineStyle = xlContinuous
    xlsHoja.Range("A" & lnFilaActual, "B" & lnFilaActual).Interior.Color = RGB(192, 192, 192)
    xlsHoja.Range("B" & lnFilaActual).Formula = "=SUM(B" & lnFilaAnterior & ":B" & lnFilaActual - 1 & ")"
    xlsHoja.Range("B" & lnFilaActual).Copy
    xlsHoja.Range("C" & lnFilaActual, IIf(pbFechaCompara, "I", "E") & lnFilaActual).PasteSpecial
    
    lnFilaActual = lnFilaActual + 1
    xlsHoja.Range("A" & lnFilaActual) = "Pasivos Monetarios"
    xlsHoja.Range("A" & lnFilaActual, "E" & lnFilaActual).Borders.LineStyle = xlContinuous
    If pbFechaCompara Then
        xlsHoja.Range("B" & lnFilaActual, "E" & lnFilaActual).Copy
        xlsHoja.Range("F" & lnFilaActual).PasteSpecial
    End If
    
    lnFilaActual = lnFilaActual + 1
    lnFilaAnterior = lnFilaActual
    For iNota = 1 To UBound(Notas)
        If Notas(iNota).nSeccion = 2 Then
            xlsHoja.Range("A" & lnFilaActual) = Notas(iNota).cDescripcion
            xlsHoja.Range("B" & lnFilaActual) = Notas(iNota).nSaldoEvaluaDolares
            xlsHoja.Range("C" & lnFilaActual) = Notas(iNota).nSaldoEvaluaSoles
            xlsHoja.Range("D" & lnFilaActual) = 0#
            xlsHoja.Range("E" & lnFilaActual).Formula = "=B" & lnFilaActual & "+C" & lnFilaActual & "+D" & lnFilaActual
            If pbFechaCompara Then
                xlsHoja.Range("F" & lnFilaActual) = Notas(iNota).nSaldoComparaDolares
                xlsHoja.Range("G" & lnFilaActual) = Notas(iNota).nSaldoComparaSoles
                xlsHoja.Range("H" & lnFilaActual) = 0#
                xlsHoja.Range("I" & lnFilaActual).Formula = "=F" & lnFilaActual & "+G" & lnFilaActual & "+H" & lnFilaActual
            End If
            lnFilaActual = lnFilaActual + 1
        End If
    Next
    If lnFilaAnterior = lnFilaActual Then
        lnFilaActual = lnFilaActual + 1
    End If
    xlsHoja.Range("A" & lnFilaAnterior, "A" & lnFilaActual - 1).BorderAround xlContinuous
    xlsHoja.Range("A" & lnFilaAnterior, "A" & lnFilaActual - 1).Copy
    xlsHoja.Range("B" & lnFilaAnterior, IIf(pbFechaCompara, "I", "E") & lnFilaActual - 1).PasteSpecial xlPasteFormats
    xlsHoja.Range("A" & lnFilaActual) = "Total Pasivos Monetarios"
    xlsHoja.Range("A" & lnFilaActual, "B" & lnFilaActual).Borders.LineStyle = xlContinuous
    xlsHoja.Range("A" & lnFilaActual, "B" & lnFilaActual).Interior.Color = RGB(192, 192, 192)
    xlsHoja.Range("B" & lnFilaActual).Formula = "=SUM(B" & lnFilaAnterior & ":B" & lnFilaActual - 1 & ")"
    xlsHoja.Range("B" & lnFilaActual).Copy
    xlsHoja.Range("C" & lnFilaActual, IIf(pbFechaCompara, "I", "E") & lnFilaActual).PasteSpecial
    
    lnFilaActual = lnFilaActual + 1
    xlsHoja.Range("A" & lnFilaActual) = "Cuentas Fuera de Balance (Instrumentos Derivados)"
    xlsHoja.Range("A" & lnFilaActual, "E" & lnFilaActual).Borders.LineStyle = xlContinuous
    If pbFechaCompara Then
        xlsHoja.Range("B" & lnFilaActual, "E" & lnFilaActual).Copy
        xlsHoja.Range("F" & lnFilaActual).PasteSpecial
    End If
    lnFilaActual = lnFilaActual + 1
    xlsHoja.Range("A" & lnFilaActual) = "Instrumentos Derivados Activos"
    lnFilaActual = lnFilaActual + 1
    xlsHoja.Range("A" & lnFilaActual) = "Instrumentos Derivados Pasivos"
    xlsHoja.Range("A" & lnFilaActual - 1, "A" & lnFilaActual + 1).BorderAround xlContinuous
    xlsHoja.Range("A" & lnFilaActual - 1, "A" & lnFilaActual + 1).Copy
    xlsHoja.Range("B" & lnFilaActual - 1, IIf(pbFechaCompara, "I", "E") & lnFilaActual + 1).PasteSpecial xlPasteFormats
    lnFilaActual = lnFilaActual + 2
    xlsHoja.Range("A" & lnFilaActual) = "Posición Monetaria Neta"
    xlsHoja.Range("A" & lnFilaActual, IIf(pbFechaCompara, "I", "E") & lnFilaActual).Borders.LineStyle = xlContinuous
    xlsHoja.Range("B4", IIf(pbFechaCompara, "I", "E") & lnFilaActual).NumberFormat = "#,##0.00;-#,##0.00"
End Sub
Private Sub generaHojaValorRazonableValorLibro(ByRef xlsHoja As Worksheet, ByVal pdFechaEvalua As Date, ByVal pbFechaCompara As Boolean, ByVal pdFechaCompara As Date)
    Dim iNota As Integer
    Dim lnFilaActual As Integer, lnFilaAnterior As Integer
    Dim lnMontoEvalua As Currency, lnMontoCompara As Currency
    Dim lnNroPosNota As Integer

    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 8
    xlsHoja.Cells.VerticalAlignment = xlCenter
    
    xlsHoja.Columns.ColumnWidth = 12
    xlsHoja.Columns("A:A").ColumnWidth = 50
    xlsHoja.Range("A2").RowHeight = 33
   
    xlsHoja.Range("A1") = "Valor Razonable y Valor en Libros"
    xlsHoja.Range("A1", "A2").MergeCells = True
    xlsHoja.Range("B1") = "Notas"
    xlsHoja.Range("B1", "B2").MergeCells = True
    xlsHoja.Range("C1").NumberFormat = "mmm-yy"
    xlsHoja.Range("C1") = pdFechaEvalua
    xlsHoja.Range("C1", "D1").MergeCells = True
    xlsHoja.Range("C2") = "Valor en Libros"
    xlsHoja.Range("D2") = "Valor Razonable"
    If pbFechaCompara Then
        xlsHoja.Range("E1").NumberFormat = "mmm-yy"
        xlsHoja.Range("E1") = pdFechaCompara
        xlsHoja.Range("E1", "F1").MergeCells = True
        xlsHoja.Range("E2") = "Valor en Libros"
        xlsHoja.Range("F2") = "Valor Razonable"
    End If
    xlsHoja.Range("A1", IIf(pbFechaCompara, "F2", "D2")).HorizontalAlignment = xlCenter
    xlsHoja.Range("A1", IIf(pbFechaCompara, "F2", "D2")).VerticalAlignment = xlCenter
    xlsHoja.Range("A1", IIf(pbFechaCompara, "F2", "D2")).WrapText = True
    xlsHoja.Range("A1", IIf(pbFechaCompara, "F2", "D2")).Borders.LineStyle = xlContinuous
    
    xlsHoja.Range("A3") = "Activo"
    xlsHoja.Range("A3").Font.Bold = True
    xlsHoja.Range("A3").Font.Underline = True
   
    lnFilaActual = 4
    lnFilaAnterior = lnFilaActual
    For iNota = 1 To UBound(Notas)
        If Notas(iNota).nSeccion = 1 Then
            lnMontoEvalua = 0: lnMontoCompara = 0
            xlsHoja.Range("A" & lnFilaActual) = Notas(iNota).cDescripcion
            lnMontoEvalua = Notas(iNota).nSaldoEvaluaUnificado
            xlsHoja.Range("C" & lnFilaActual) = lnMontoEvalua
            xlsHoja.Range("D" & lnFilaActual) = 0#
            If pbFechaCompara Then
                lnMontoCompara = Notas(iNota).nSaldoComparaUnificado
                xlsHoja.Range("E" & lnFilaActual) = lnMontoCompara
                xlsHoja.Range("F" & lnFilaActual) = 0#
            End If
            If lnMontoEvalua > 0 Or lnMontoCompara > 0 Then
                lnNroPosNota = lnNroPosNota + 1
                xlsHoja.Range("B" & lnFilaActual) = lnNroPosNota
            End If
            lnFilaActual = lnFilaActual + 1
        End If
    Next
    If lnFilaAnterior = lnFilaActual Then
        lnFilaActual = lnFilaActual + 1
    End If
    xlsHoja.Range("A" & lnFilaAnterior - 1, "A" & lnFilaActual - 1).BorderAround xlContinuous
    xlsHoja.Range("A" & lnFilaAnterior - 1, "A" & lnFilaActual - 1).Copy
    xlsHoja.Range("B" & lnFilaAnterior - 1, IIf(pbFechaCompara, "F", "D") & lnFilaActual - 1).PasteSpecial xlPasteFormats
    xlsHoja.Range("A" & lnFilaActual) = "Total"
    xlsHoja.Range("A" & lnFilaActual).Font.Bold = True
    xlsHoja.Range("A" & lnFilaActual, IIf(pbFechaCompara, "F", "D") & lnFilaActual).Borders.LineStyle = xlContinuous
    xlsHoja.Range("C" & lnFilaActual).Formula = "=SUM(C" & lnFilaAnterior & ":C" & lnFilaActual - 1 & ")"
    xlsHoja.Range("C" & lnFilaActual).Copy
    xlsHoja.Range("D" & lnFilaActual, IIf(pbFechaCompara, "F", "D") & lnFilaActual).PasteSpecial
    
    lnFilaActual = lnFilaActual + 1
    xlsHoja.Range("A" & lnFilaActual, IIf(pbFechaCompara, "F", "D") & lnFilaActual).BorderAround xlContinuous
    lnFilaActual = lnFilaActual + 1
    xlsHoja.Range("A" & lnFilaActual) = "Pasivo"
    xlsHoja.Range("A" & lnFilaActual).Font.Bold = True
    xlsHoja.Range("A" & lnFilaActual).Font.Underline = True
    
    lnFilaActual = lnFilaActual + 1
    lnFilaAnterior = lnFilaActual
    lnNroPosNota = 0
    For iNota = 1 To UBound(Notas)
        If Notas(iNota).nSeccion = 2 Then
            lnMontoEvalua = 0: lnMontoCompara = 0
            xlsHoja.Range("A" & lnFilaActual) = Notas(iNota).cDescripcion
            lnMontoEvalua = Notas(iNota).nSaldoEvaluaUnificado
            xlsHoja.Range("C" & lnFilaActual) = lnMontoEvalua
            xlsHoja.Range("D" & lnFilaActual) = 0#
            If pbFechaCompara Then
                lnMontoCompara = Notas(iNota).nSaldoComparaUnificado
                xlsHoja.Range("E" & lnFilaActual) = lnMontoCompara
                xlsHoja.Range("F" & lnFilaActual) = 0#
            End If
            If lnMontoEvalua > 0 Or lnMontoCompara > 0 Then
                lnNroPosNota = lnNroPosNota + 1
                xlsHoja.Range("B" & lnFilaActual) = lnNroPosNota
            End If
            lnFilaActual = lnFilaActual + 1
        End If
    Next
    If lnFilaAnterior = lnFilaActual Then
        lnFilaActual = lnFilaActual + 1
    End If
    xlsHoja.Range("A" & lnFilaAnterior - 1, "A" & lnFilaActual - 1).BorderAround xlContinuous
    xlsHoja.Range("A" & lnFilaAnterior - 1, "A" & lnFilaActual - 1).Copy
    xlsHoja.Range("B" & lnFilaAnterior - 1, IIf(pbFechaCompara, "F", "D") & lnFilaActual - 1).PasteSpecial xlPasteFormats
    xlsHoja.Range("A" & lnFilaActual) = "Total"
    xlsHoja.Range("A" & lnFilaActual).Font.Bold = True
    xlsHoja.Range("A" & lnFilaActual, IIf(pbFechaCompara, "F", "D") & lnFilaActual).Borders.LineStyle = xlContinuous
    xlsHoja.Range("C" & lnFilaActual).Formula = "=SUM(C" & lnFilaAnterior & ":C" & lnFilaActual - 1 & ")"
    xlsHoja.Range("C" & lnFilaActual).Copy
    xlsHoja.Range("D" & lnFilaActual, IIf(pbFechaCompara, "F", "D") & lnFilaActual).PasteSpecial
    
    xlsHoja.Range("C4", IIf(pbFechaCompara, "F", "D") & lnFilaActual).NumberFormat = "#,##0.00;-#,##0.00"
End Sub
Private Sub generaHojaExposicionRiesgoLiquidez(ByRef xlsHoja As Worksheet, ByVal pdFechaEvalua As Date, ByVal pbFechaCompara As Boolean, ByVal pdFechaCompara As Date)
    Dim oNRep As New NRepFormula
    Dim rsRep As New ADODB.Recordset
    Dim iNota As Integer
    Dim lnFilaActual As Integer, lnFilaAnterior As Integer
    Dim lnMontoEvalua As Currency, lnMontoCompara As Currency
    Dim lnNroPosNota As Integer

    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 8
    xlsHoja.Cells.VerticalAlignment = xlCenter
    
    xlsHoja.Columns.ColumnWidth = 12
    xlsHoja.Columns("A:A").ColumnWidth = 75
    xlsHoja.Range("A1").RowHeight = 33
   
    xlsHoja.Range("A1") = "Exposición al Riesgo de Liquidez"
    xlsHoja.Range("A1", "A2").MergeCells = True
    xlsHoja.Range("B1") = "A la Vista"
    xlsHoja.Range("B2") = "s/.(000)"
    xlsHoja.Range("C1") = "Hasta 1 Mes"
    xlsHoja.Range("D1") = "Mas de 1 hasta 3 meses"
    xlsHoja.Range("E1") = "Mas de 3 hasta 12 meses"
    xlsHoja.Range("F1") = "Mas de 1 año"
    xlsHoja.Range("B2").Copy
    xlsHoja.Range("C2", "F2").PasteSpecial
    xlsHoja.Range("A1", "F2").Borders.LineStyle = xlContinuous
    xlsHoja.Range("A3") = "Riesgo de Balance"
    xlsHoja.Range("A3").Borders.LineStyle = xlContinuous
    xlsHoja.Range("B3", "F3").BorderAround xlContinuous
    xlsHoja.Range("A1", "F3").WrapText = True
    xlsHoja.Range("A1", "F3").HorizontalAlignment = xlCenter
    xlsHoja.Range("A1", "F3").VerticalAlignment = xlCenter
    
    xlsHoja.Range("A4") = "Pasivo"
    xlsHoja.Range("A4").Font.Bold = True
    xlsHoja.Range("A4").Font.Underline = True
    
    lnFilaActual = 5
    lnFilaAnterior = lnFilaActual
    
    xlsHoja.Range("A5") = "Obligaciones con el Público"
    
    xlsHoja.Range("A6") = "Fondos Interbancarios"
    xlsHoja.Range("A7") = "Depósitos de Empresas del Sistema Financiero y Organismos Financieros Internacionales"
    xlsHoja.Range("A8") = "Adeudos y Obligaciones Financieros"
    xlsHoja.Range("A9") = "Derivados para Negociación"
    xlsHoja.Range("A10") = "Derivados de Cobertura"
    xlsHoja.Range("A11") = "Cuentas por Pagar"
    xlsHoja.Range("A11") = "Otros Pasivos"
    
    'xlsHoja.Range("C4", IIf(pbFechaCompara, "F", "D") & lnFilaActual).NumberFormat = "#,##0.00;-#,##0.00"
    Set oNRep = Nothing
    Set rsRep = Nothing
End Sub
Private Sub generaHojaCalidadCrediticiaActivosFinancieros(ByRef xlsHoja As Worksheet, ByVal pdFechaEvalua As Date, ByVal pbFechaCompara As Boolean, ByVal pdFechaCompara As Date, ByVal pbResumen As Boolean)
    Dim lnFilaActual As Long, lnFilaAnterior As Long

    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 8
    xlsHoja.Cells.VerticalAlignment = xlCenter
    xlsHoja.Columns.ColumnWidth = 14
    xlsHoja.Columns("A:A").ColumnWidth = 50
    xlsHoja.Range("A2").RowHeight = 40
    xlsHoja.Range("A:A").NumberFormat = "@"
    xlsHoja.Range("B:G").NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFilaActual = 1
    generaSeccionHojaCalidadCrediticiaActivosFinancieros xlsHoja, pdFechaEvalua, lnFilaActual, pbResumen
    If pbFechaCompara Then
        lnFilaActual = lnFilaActual + 1
        generaSeccionHojaCalidadCrediticiaActivosFinancieros xlsHoja, pdFechaCompara, lnFilaActual, pbResumen
    End If
End Sub
Private Sub generaSeccionHojaCalidadCrediticiaActivosFinancieros(ByRef xlsHoja As Worksheet, ByVal pdFecha As Date, ByRef pnFilaActual As Long, ByVal pbResumen As Boolean)
    Dim oNRep As New NRepFormula
    Dim rsRep As New ADODB.Recordset
    Dim lnFilaActual As Long, lnFilaAnterior As Long
    Dim lnProvisionNoMinorista As Currency, lnProvisionMiPe As Currency, lnProvisionConsumo As Currency, lnProvisionHipotecario As Currency
    Dim lnSaldoCapNoMinorista As Currency, lnSaldoCapMiPe As Currency, lnSaldoCapConsumo As Currency, lnSaldoCapHipotecario As Currency
    
    lnFilaActual = pnFilaActual
    xlsHoja.Range("A" & lnFilaActual).NumberFormat = "mmmm-yyyy"
    xlsHoja.Range("A" & lnFilaActual) = pdFecha
    xlsHoja.Range("A" & lnFilaActual, "A" & lnFilaActual + 1).MergeCells = True
    xlsHoja.Range("B" & lnFilaActual) = "Cartera de Créditos"
    xlsHoja.Range("B" & lnFilaActual, "G" & lnFilaActual).MergeCells = True
    xlsHoja.Range("B" & lnFilaActual + 1) = "Créditos No Minoristas"
    xlsHoja.Range("C" & lnFilaActual + 1) = "Créditos pequeña y micro empresa"
    xlsHoja.Range("D" & lnFilaActual + 1) = "Créditos de consumo"
    xlsHoja.Range("E" & lnFilaActual + 1) = "Créditos hipotecarios para vivienda"
    xlsHoja.Range("F" & lnFilaActual + 1) = "Total"
    xlsHoja.Range("G" & lnFilaActual + 1) = "%"
    xlsHoja.Range("A" & lnFilaActual, "G" & lnFilaActual + 1).VerticalAlignment = xlCenter
    xlsHoja.Range("A" & lnFilaActual, "G" & lnFilaActual + 1).HorizontalAlignment = xlCenter
    xlsHoja.Range("A" & lnFilaActual, "G" & lnFilaActual + 1).WrapText = True
    xlsHoja.Range("A" & lnFilaActual, "G" & lnFilaActual + 1).Font.Bold = True

    lnFilaActual = lnFilaActual + 2
    lnFilaAnterior = lnFilaActual
    Set rsRep = oNRep.ListaCalidadCrediticiaActivosFinancieros(pdFecha, pbResumen)
    Do While Not rsRep.EOF
        xlsHoja.Cells(lnFilaActual, 1) = rsRep!cAgrupa
        xlsHoja.Cells(lnFilaActual, 2) = rsRep!nSaldoCapNoMinorista
        xlsHoja.Cells(lnFilaActual, 3) = rsRep!nSaldoCapMiPe
        xlsHoja.Cells(lnFilaActual, 4) = rsRep!nSaldoCapConsumo
        xlsHoja.Cells(lnFilaActual, 5) = rsRep!nSaldoCapHipotecario
        xlsHoja.Cells(lnFilaActual, 6).Formula = "=SUM(B" & lnFilaActual & ":E" & lnFilaActual & ")"

        If rsRep!nOrden = 0 Then
            lnProvisionNoMinorista = lnProvisionNoMinorista + (rsRep!nProvisionNoMinorista * -1)
            lnProvisionMiPe = lnProvisionMiPe + (rsRep!nProvisionMiPe * -1)
            lnProvisionConsumo = lnProvisionConsumo + (rsRep!nProvisionConsumo * -1)
            lnProvisionHipotecario = lnProvisionHipotecario + (rsRep!nProvisionHipotecario * -1)
            lnSaldoCapNoMinorista = lnSaldoCapNoMinorista + rsRep!nSaldoCapNoMinorista
            lnSaldoCapMiPe = lnSaldoCapMiPe + rsRep!nSaldoCapMiPe
            lnSaldoCapConsumo = lnSaldoCapConsumo + rsRep!nSaldoCapConsumo
            lnSaldoCapHipotecario = lnSaldoCapHipotecario + rsRep!nSaldoCapHipotecario
        End If
        lnFilaActual = lnFilaActual + 1
        rsRep.MoveNext
    Loop
    xlsHoja.Cells(lnFilaActual, 1) = "Cartera Bruta"
    xlsHoja.Cells(lnFilaActual, 2) = lnSaldoCapNoMinorista
    xlsHoja.Cells(lnFilaActual, 3) = lnSaldoCapMiPe
    xlsHoja.Cells(lnFilaActual, 4) = lnSaldoCapConsumo
    xlsHoja.Cells(lnFilaActual, 5) = lnSaldoCapHipotecario
    xlsHoja.Cells(lnFilaActual, 6).Formula = "=SUM(B" & lnFilaActual & ":E" & lnFilaActual & ")"

    lnFilaActual = lnFilaActual + 1
    xlsHoja.Cells(lnFilaActual, 1) = "Provisiones"
    xlsHoja.Cells(lnFilaActual, 2) = lnProvisionNoMinorista
    xlsHoja.Cells(lnFilaActual, 3) = lnProvisionMiPe
    xlsHoja.Cells(lnFilaActual, 4) = lnProvisionConsumo
    xlsHoja.Cells(lnFilaActual, 5) = lnProvisionHipotecario
    xlsHoja.Cells(lnFilaActual, 6).Formula = "=SUM(B" & lnFilaActual & ":E" & lnFilaActual & ")"

    lnFilaActual = lnFilaActual + 1
    xlsHoja.Cells(lnFilaActual, 1) = "Total Neto"
    xlsHoja.Cells(lnFilaActual, 2).Formula = "=SUM(B" & lnFilaActual - 2 & ":B" & lnFilaActual - 1 & ")"
    xlsHoja.Cells(lnFilaActual, 3).Formula = "=SUM(C" & lnFilaActual - 2 & ":C" & lnFilaActual - 1 & ")"
    xlsHoja.Cells(lnFilaActual, 4).Formula = "=SUM(D" & lnFilaActual - 2 & ":D" & lnFilaActual - 1 & ")"
    xlsHoja.Cells(lnFilaActual, 5).Formula = "=SUM(E" & lnFilaActual - 2 & ":E" & lnFilaActual - 1 & ")"
    xlsHoja.Cells(lnFilaActual, 6).Formula = "=SUM(B" & lnFilaActual & ":E" & lnFilaActual & ")"

    xlsHoja.Cells(lnFilaAnterior, 7).Formula = "=(F" & lnFilaAnterior & "/$F$" & lnFilaActual & ")*100"
    xlsHoja.Cells(lnFilaAnterior, 7).Copy
    xlsHoja.Range("G" & lnFilaAnterior & ":" & "G" & lnFilaActual).PasteSpecial

    xlsHoja.Range("A" & pnFilaActual, "G" & lnFilaActual).Borders.LineStyle = xlContinuous
    pnFilaActual = lnFilaActual
    Set rsRep = Nothing
    Set oNRep = Nothing
End Sub
Private Sub generaHojaCalidadCreditosVencidosxDiasAtraso(ByRef xlsHoja As Worksheet, ByVal pdFechaEvalua As Date, ByVal pbFechaCompara As Boolean, ByVal pdFechaCompara As Date)
    Dim oNRep As New NRepFormula
    Dim rsRep As New ADODB.Recordset
    Dim lnFilaActual As Integer

    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 8
    xlsHoja.Cells.VerticalAlignment = xlCenter
    
    xlsHoja.Columns.ColumnWidth = 10
    xlsHoja.Columns("A:A").ColumnWidth = 50
    xlsHoja.Range("A2").RowHeight = 40

    xlsHoja.Range("A2") = "Tipo de Crédito"
    xlsHoja.Range("A3") = "Días de Atraso"
    xlsHoja.Range("B1").NumberFormat = "mmmm-yyyy"
    xlsHoja.Range("B1") = pdFechaEvalua
    xlsHoja.Range("B1", "F1").MergeCells = True
    '''xlsHoja.Range("B2") = "Créditos Vencidos y No deteriorados (En Nuevos Soles)" 'MARG ERS044-2016
    xlsHoja.Range("B2") = "Créditos Vencidos y No deteriorados " & StrConv(gcPEN_PLURAL, vbProperCase) 'MARG ERS044-2016
    xlsHoja.Range("B2", "E2").MergeCells = True
    xlsHoja.Range("B3") = "'1-15"
    xlsHoja.Range("C3") = "'16-30"
    xlsHoja.Range("D3") = "'31-60"
    xlsHoja.Range("E3") = "'> 60"
    xlsHoja.Range("F2") = "'Total"
    xlsHoja.Range("F2", "F3").MergeCells = True
    xlsHoja.Range("A1", "F3").Font.Bold = True
    xlsHoja.Range("A1", "F3").VerticalAlignment = xlCenter
    xlsHoja.Range("A1", "F3").HorizontalAlignment = xlCenter
    xlsHoja.Range("A1", "F3").WrapText = True
    
    xlsHoja.Range("B4", "F14") = 0#
    xlsHoja.Range("A4") = "Corporativos"
    xlsHoja.Range("A5") = "Grande Empresa"
    xlsHoja.Range("A6") = "Mediana Empresa"
    xlsHoja.Range("A7") = "Sub Total"
    xlsHoja.Range("B7").Formula = "=SUM(B4:B6)"
    xlsHoja.Range("B7").Copy
    xlsHoja.Range("C7", "F7").PasteSpecial
    xlsHoja.Range("A7").Font.Bold = True
    xlsHoja.Range("A8") = "Pequeña Empresa"
    xlsHoja.Range("A9") = "Microempresa"
    xlsHoja.Range("A10") = "Consumo Revolvente"
    xlsHoja.Range("A11") = "Consumo no Revolvente"
    xlsHoja.Range("A12") = "Hipotecario"
    xlsHoja.Range("A13") = "Sub Total"
    xlsHoja.Range("F4").Formula = "=SUM(B4:E4)"
    xlsHoja.Range("F4").Copy
    xlsHoja.Range("F5", "F14").PasteSpecial
    xlsHoja.Range("B13").Formula = "=SUM(B8:B12)"
    xlsHoja.Range("B14").Formula = "=B7+B13"
    xlsHoja.Range("B13", "B14").Copy
    xlsHoja.Range("C13", "F13").PasteSpecial
    
    xlsHoja.Range("A14") = "Total"
    xlsHoja.Range("A13", "A14").Font.Bold = True
    xlsHoja.Range("A1", "F14").Borders.LineStyle = xlContinuous
    
    If pbFechaCompara Then
        xlsHoja.Range("B1", "F14").Copy
        xlsHoja.Range("G1").PasteSpecial
        xlsHoja.Range("G1") = pdFechaCompara
        Set rsRep = oNRep.ListaCreditosVencidosxDiasAtraso(pdFechaCompara)
        Do While Not rsRep.EOF
            lnFilaActual = 0
            Select Case rsRep!cTpoCredCod
                Case "150": lnFilaActual = 4
                Case "250": lnFilaActual = 5
                Case "350": lnFilaActual = 6
                Case "450": lnFilaActual = 8
                Case "550": lnFilaActual = 9
                Case "650": lnFilaActual = 10
                Case "750": lnFilaActual = 11
                Case "850": lnFilaActual = 12
            End Select
            If lnFilaActual > 0 Then
                xlsHoja.Cells(lnFilaActual, 7) = rsRep!nSaldoCapTramo1
                xlsHoja.Cells(lnFilaActual, 8) = rsRep!nSaldoCapTramo2
                xlsHoja.Cells(lnFilaActual, 9) = rsRep!nSaldoCapTramo3
                xlsHoja.Cells(lnFilaActual, 10) = rsRep!nSaldoCapTramo4
            End If
            rsRep.MoveNext
        Loop
    End If
    
    Set rsRep = oNRep.ListaCreditosVencidosxDiasAtraso(pdFechaEvalua)
    Do While Not rsRep.EOF
        lnFilaActual = 0
        Select Case rsRep!cTpoCredCod
            Case "150": lnFilaActual = 4
            Case "250": lnFilaActual = 5
            Case "350": lnFilaActual = 6
            Case "450": lnFilaActual = 8
            Case "550": lnFilaActual = 9
            Case "650": lnFilaActual = 10
            Case "750": lnFilaActual = 11
            Case "850": lnFilaActual = 12
        End Select
        If lnFilaActual > 0 Then
            xlsHoja.Cells(lnFilaActual, 2) = rsRep!nSaldoCapTramo1
            xlsHoja.Cells(lnFilaActual, 3) = rsRep!nSaldoCapTramo2
            xlsHoja.Cells(lnFilaActual, 4) = rsRep!nSaldoCapTramo3
            xlsHoja.Cells(lnFilaActual, 5) = rsRep!nSaldoCapTramo4
        End If
        rsRep.MoveNext
    Loop
    xlsHoja.Range("B4", IIf(pbFechaCompara, "K", "F") & "14").NumberFormat = "#,##0.00;-#,##0.00"
    
    Set rsRep = Nothing
    Set oNRep = Nothing
End Sub
Private Sub generaHojaConcentracionRiesgoxSector(ByRef xlsHoja As Worksheet, ByVal pdFechaEvalua As Date, ByVal pbFechaCompara As Boolean, ByVal pdFechaCompara As Date)
    Dim oNRep As New NRepFormula
    Dim rsRep As New ADODB.Recordset
    Dim lnFilaActual As Integer, lnFilaAnterior As Integer
    Dim lnNroReg As Integer
    Dim lnTotal As Double

    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 8
    xlsHoja.Cells.VerticalAlignment = xlCenter
    
    xlsHoja.Columns.ColumnWidth = 10
    xlsHoja.Columns("A:A").ColumnWidth = 80
    xlsHoja.Range("A1", "A2").RowHeight = 20

    xlsHoja.Range("A1") = "Crédito"
    xlsHoja.Range("A1", "A2").MergeCells = True
    xlsHoja.Range("B1") = "Año " & CStr(Year(pdFechaEvalua))
    xlsHoja.Range("B1", "C1").MergeCells = True
    xlsHoja.Range("B2") = "Total"
    xlsHoja.Range("C2") = "Porcentaje"
    xlsHoja.Range("A1", "C2").Font.Bold = True
    xlsHoja.Range("A1", "C2").VerticalAlignment = xlCenter
    xlsHoja.Range("A1", "C2").HorizontalAlignment = xlCenter
    xlsHoja.Range("A1", "C2").WrapText = True
    
    lnFilaActual = 3
    lnFilaAnterior = lnFilaActual
    
    Set rsRep = oNRep.ListaConcentracionRiesgoxSector(pdFechaEvalua)
    lnNroReg = rsRep.RecordCount
    Do While Not rsRep.EOF
        xlsHoja.Range("A" & lnFilaActual) = rsRep!cDescrip
        xlsHoja.Range("B" & lnFilaActual) = rsRep!nSaldoCap
        xlsHoja.Range("C" & lnFilaActual).Formula = "=B" & lnFilaActual & "/B" & (lnNroReg + 3)
        If rsRep!nDetalle = 0 Then
            lnTotal = lnTotal + rsRep!nSaldoCap
        End If
        rsRep.MoveNext
        lnFilaActual = lnFilaActual + 1
    Loop
    If lnFilaActual = lnFilaAnterior Then
        lnFilaActual = lnFilaActual + 1
    End If
    xlsHoja.Range("C" & lnFilaActual).Formula = "=B" & lnFilaActual & "/B" & (lnNroReg + 3)
    xlsHoja.Range("B3", "B" & lnFilaActual).NumberFormat = "#,##0.00;-#,##0.00"
    xlsHoja.Range("B" & lnFilaActual) = lnTotal
    xlsHoja.Range("C3", "C" & lnFilaActual).NumberFormat = "0.00%"
    xlsHoja.Range("A1", "C" & lnFilaActual - 1).Borders.LineStyle = xlContinuous
    
    If pbFechaCompara Then
        xlsHoja.Range("B1", "C" & lnFilaActual).Copy
        xlsHoja.Range("D1").PasteSpecial
        xlsHoja.Range("D1") = "Año " & CStr(Year(pdFechaCompara))
        lnTotal = 0
        lnFilaActual = 3
        lnFilaAnterior = lnFilaActual
        Set rsRep = oNRep.ListaConcentracionRiesgoxSector(pdFechaCompara)
        lnNroReg = rsRep.RecordCount
        Do While Not rsRep.EOF
            xlsHoja.Range("D" & lnFilaActual) = rsRep!nSaldoCap
            xlsHoja.Range("E" & lnFilaActual).Formula = "=D" & lnFilaActual & "/D" & (lnNroReg + 3)
            If rsRep!nDetalle = 0 Then
                lnTotal = lnTotal + rsRep!nSaldoCap
            End If
            rsRep.MoveNext
            lnFilaActual = lnFilaActual + 1
        Loop
        If lnFilaActual = lnFilaAnterior Then
            lnFilaActual = lnFilaActual + 1
        End If
        xlsHoja.Range("E" & lnFilaActual).Formula = "=D" & lnFilaActual & "/D" & (lnNroReg + 3)
        xlsHoja.Range("D" & lnFilaActual) = lnTotal
    End If
    
    Set rsRep = Nothing
    Set oNRep = Nothing
End Sub

