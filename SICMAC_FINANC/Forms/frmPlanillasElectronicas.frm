VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPlanillasElectronicas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programa de Libros Electrónicos"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5115
   Icon            =   "frmPlanillasElectronicas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmPlanillasElectronicas.frx":030A
   ScaleHeight     =   2475
   ScaleWidth      =   5115
   Begin MSComctlLib.ProgressBar BarraProgreso 
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   2255
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
      TabIndex        =   9
      Top             =   2220
      Width           =   5115
      _ExtentX        =   9022
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
   Begin VB.Frame fraPerEvaluar 
      Caption         =   "Periodo"
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
      TabIndex        =   6
      Top             =   80
      Width           =   4935
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmPlanillasElectronicas.frx":0614
         Left            =   2760
         List            =   "frmPlanillasElectronicas.frx":0616
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   2280
         TabIndex        =   8
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año  :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   270
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Programa de Libros Electrónicos"
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
      Height          =   975
      Left            =   80
      TabIndex        =   5
      Top             =   800
      Width           =   4935
      Begin VB.ComboBox cboLibro 
         Height          =   315
         ItemData        =   "frmPlanillasElectronicas.frx":0618
         Left            =   720
         List            =   "frmPlanillasElectronicas.frx":061A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Libro :"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   390
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
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
      Left            =   2595
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   1830
      Width           =   1200
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
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
      Left            =   1350
      TabIndex        =   3
      ToolTipText     =   "Generar"
      Top             =   1830
      Width           =   1200
   End
End
Attribute VB_Name = "frmPlanillasElectronicas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    CentraForm Me
    Call ListarMes
    Call ListarPrograma
    txtAnio.Text = Year(gdFecSis)
    cboMes.ListIndex = IndiceListaCombo(cboMes, Month(gdFecSis))
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cboMes.SetFocus
    End If
End Sub
Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboLibro.SetFocus
    End If
End Sub
Private Sub cboLibro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub
Private Sub optPLE_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub
Private Sub ListarMes()
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
End Sub
Private Sub ListarPrograma()
    cboLibro.Clear
    cboLibro.AddItem "REGISTRO DE COMPRAS Y GASTOS" & Space(200) & "1"
    cboLibro.AddItem "REGISTRO DE VENTAS" & Space(200) & "2"
    cboLibro.AddItem "LIBRO DIARIO" & Space(200) & "3"
    cboLibro.AddItem "LIBRO MAYOR" & Space(200) & "4"
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Public Function validarPlanilla() As Boolean
    validarPlanilla = True
    If Val(txtAnio.Text) <= 1900 Then
        MsgBox "Ud. debe seleccionar el año a generar la información", vbInformation, "Aviso"
        validarPlanilla = False
        txtAnio.SetFocus
        Exit Function
    End If
    If cboMes.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el mes a generar la información", vbInformation, "Aviso"
        validarPlanilla = False
        cboMes.SetFocus
        Exit Function
    End If
    If cboLibro.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el Libro respectivo", vbInformation, "Aviso"
        validarPlanilla = False
        cboLibro.SetFocus
        Exit Function
    End If
End Function
Private Sub cmdGenerar_Click()
    Dim sCadena As String
    Dim sCadenaDetCta As String 'PASI20140531
    Dim fs As New Scripting.FileSystemObject
    Dim psArchivoAGrabar As String
    Dim psArchivoAGrabarDetCta As String 'PASI20140531
    Dim sCad As String
    Dim sCadDetCta As String 'PASI20140531
    Dim ArcSal As Integer
    Dim ArcSalDetCta As Integer 'PASI20140531
    Dim lsTpoLibro As Integer
    Dim ldFecha As Date
    Dim ArchTxtCompGast As String 'NAGL ERS 012-2017 20170710
    If Not validarPlanilla Then Exit Sub
    
  
    On Error GoTo ErrGenerar
    Screen.MousePointer = 11
    lsTpoLibro = CInt(Right(cboLibro.Text, 2))
    ldFecha = obtenerFechaFinMes(CInt(Right(cboMes.Text, 2)), Val(txtAnio))
    cmdGenerar.Enabled = False
    sCad = ""
    sCadDetCta = "" 'PASI20140531
    ArcSal = FreeFile
    ArcSalDetCta = FreeFile
    Select Case lsTpoLibro
        Case 1
            'sCadena = RecuperaCadenaRegistroCompras(ldFecha)
            'psArchivoAGrabar = App.path & "\SPOOLER\LE" & gsRUC & Format(ldFecha, "YYYYMM00") & "080100001111_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".txt"
            'Comentado by NAGL
            If Not RecuperaCadenaRegistroComprasNew(ldFecha, ArchTxtCompGast) Then
                 MsgBox "Al parecer existen inconvenientes. Por el momento no se ha podido generar el Archivo txt, favor comuniquese con el Area de TI.", vbInformation, "Aviso"
                 Exit Sub
            End If 'NAGL ERS 012-2017 20170710
        Case 2
            sCadena = RecuperaCadenaRegistroVentas(ldFecha)
            psArchivoAGrabar = App.path & "\SPOOLER\LE" & gsRUC & Format(ldFecha, "YYYYMM00") & "140100001111_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".txt"
        Case 3
            sCadena = RecuperaCadenaLibroDiario(ldFecha)
            sCadenaDetCta = RecuperaCadenaLibroDiarioDetCtaCont(ldFecha) 'PASI20140531
            psArchivoAGrabar = App.path & "\SPOOLER\LE" & gsRUC & Format(ldFecha, "YYYYMM00") & "050100001111_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".txt"
            psArchivoAGrabarDetCta = App.path & "\SPOOLER\DETCTA_LE" & gsRUC & Format(ldFecha, "YYYYMM00") & "050100001111_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".txt" 'PASI20140531
        Case 4
            sCadena = RecuperaCadenaLibroMayor(ldFecha)
            psArchivoAGrabar = App.path & "\SPOOLER\LE" & gsRUC & Format(ldFecha, "YYYYMM00") & "060100001111_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".txt"
    End Select
    
    If lsTpoLibro <> 1 Then 'NAGL ERS 012-2017 20170710
        If sCadena <> "" Then
            Open psArchivoAGrabar For Output As ArcSal
            Print #1, sCad; sCadena
            MsgBox "Archivo generado correctamente en " & psArchivoAGrabar, vbInformation, "Aviso"
        Else
            MsgBox "No se encontró información para generar en este periodo", vbInformation, "Aviso"
        End If
        Close ArcSal
        
        'PASI20140531
        If lsTpoLibro = 3 Then
            Open psArchivoAGrabarDetCta For Output As ArcSalDetCta
            Print #1, sCadDetCta; sCadenaDetCta
            MsgBox "Archivo Detalle de Cta Contable generado correctamente en " & psArchivoAGrabarDetCta, vbInformation, "Aviso"
        End If
         Close ArcSalDetCta
        'end PASI
    End If
    
    cmdGenerar.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
ErrGenerar:
    cmdGenerar.Enabled = True
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Public Function RecuperaCadenaRegistroCompras(ByVal pdFecha As Date) As String
    Dim oReg As New NContImpreReg
    Dim rs As New ADODB.Recordset
    Dim lsSeparador As String, lsCadena As String
    Dim lnNroRegistros As Long, lnLineaActual As Long, i As Long
    Dim N As Integer
    
    N = vbYes
    If oReg.ExisteRegistroComprasPLE(pdFecha) Then
        N = MsgBox("Reporte de Registro de Compras ya fue generado. ¿Desea volver a Procesar?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Aviso")
        If N = vbCancel Then Exit Function
    End If
    If N = vbYes Then
        Call oReg.InsertaRegistroComprasPLE(pdFecha, CDate(gdFecSis & " " & Format(Time, "hh:mm:ss")), gsCodUser)
    End If

    lsSeparador = "|"
    BarraProgreso.value = 0
    BarraProgreso.Min = 0
    BarraProgreso.value = 0
    EstadoBarra.Panels(1) = "Generando..."

    Set rs = oReg.RecuperaDatosRegistroComprasPLE(pdFecha)
    BarraProgreso.Max = rs.RecordCount
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    
    lnNroRegistros = rs.RecordCount
    For i = 1 To rs.RecordCount
        lnLineaActual = lnLineaActual + 1
        lsCadena = lsCadena & rs!cPeriodo & lsSeparador
        lsCadena = lsCadena & rs!nMovNro & lsSeparador 'EJVG20140520
        lsCadena = lsCadena & rs!cMovItem & lsSeparador 'EJVG20140520
        'lsCadena = lsCadena & rs!nCorrelativo & lsSeparador
        lsCadena = lsCadena & Format(rs!dFechaEmision, "dd/mm/yyyy") & lsSeparador
        lsCadena = lsCadena & IIf(DateDiff("D", rs!dFechaPago, CDate("1900-01-01")) = 0, "", Format(rs!dFechaPago, "dd/mm/yyyy")) & lsSeparador
        lsCadena = lsCadena & rs!cTpoDoc & lsSeparador
        lsCadena = lsCadena & rs!cNroSerie & lsSeparador
        lsCadena = lsCadena & rs!cAnioEmisionDua & lsSeparador
        lsCadena = lsCadena & rs!cNroDocumento & lsSeparador
        lsCadena = lsCadena & rs!cNroFinalOpeDiarias & lsSeparador
        lsCadena = lsCadena & rs!cProveedorTpoDocId & lsSeparador
        lsCadena = lsCadena & rs!cProveedorNroDocId & lsSeparador
        lsCadena = lsCadena & rs!cProveedorNombre & lsSeparador
        lsCadena = lsCadena & Format(rs!nBaseImponible1, "#0.00") & lsSeparador
        lsCadena = lsCadena & Format(rs!nIGV1, "#0.00") & lsSeparador
        lsCadena = lsCadena & Format(rs!nBaseImponible2, "#0.00") & lsSeparador
        lsCadena = lsCadena & Format(rs!nIGV2, "#0.00") & lsSeparador
        lsCadena = lsCadena & Format(rs!nBaseImponible3, "#0.00") & lsSeparador
        lsCadena = lsCadena & Format(rs!nIGV3, "#0.00") & lsSeparador
        lsCadena = lsCadena & Format(rs!nValorAdqNOGrab, "#0.00") & lsSeparador
        lsCadena = lsCadena & Format(rs!nMontoImpSelect, "#0.00") & lsSeparador
        lsCadena = lsCadena & Format(rs!nOtroTrib, "#0.00") & lsSeparador
        lsCadena = lsCadena & Format(rs!nImporteTotal, "#0.00") & lsSeparador
        'PASI20160407 Cod.Moneda
        lsCadena = lsCadena & Format(rs!nTpoCambio, "#0.000") & lsSeparador
        lsCadena = lsCadena & IIf(DateDiff("D", rs!dFechaEmisionMod, CDate("1900-01-01")) = 0, "01/01/0001", Format(rs!dFechaEmisionMod, "dd/mm/yyyy")) & lsSeparador
        lsCadena = lsCadena & rs!cTpoDocMod & lsSeparador
        lsCadena = lsCadena & rs!cNroSerieMod & lsSeparador
        lsCadena = lsCadena & rs!cDepAdua & lsSeparador 'EJVG20140520
        lsCadena = lsCadena & rs!cNroDocumentoMod & lsSeparador
        lsCadena = lsCadena & rs!cNroPagoSND & lsSeparador
        lsCadena = lsCadena & IIf(DateDiff("D", rs!dDetraccionFecha, CDate("1900-01-01")) = 0, "01/01/0001", Format(rs!dDetraccionFecha, "dd/mm/yyyy")) & lsSeparador
        lsCadena = lsCadena & rs!cDetraccionNroDoc & lsSeparador
        lsCadena = lsCadena & rs!cMarcaSujetoReten & lsSeparador
        lsCadena = lsCadena & rs!cEstado & lsSeparador
        If lnLineaActual <> lnNroRegistros Then
            lsCadena = lsCadena & Chr(10)
        End If
        BarraProgreso.value = i
        EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
        rs.MoveNext
    Next
    RecuperaCadenaRegistroCompras = lsCadena
    EstadoBarra.Panels(1) = "Proceso Finalizado"
    Set rs = Nothing
    Set oReg = Nothing
End Function
'NAGL ERS 012-2017 20170710 *******************************************
Public Function RecuperaCadenaRegistroComprasNew(ByVal pdFecha As Date, ByRef ArchTxtCompGast As String) As Boolean
    Dim oReg As New NContImpreReg
    Dim rs As New ADODB.Recordset
    Dim lsSeparador As String, lsCadena As String
    Dim lnNroRegistros As Long, lnLineaActual As Long, i As Long
    Dim N As Integer
    
    N = vbYes
    If oReg.ExisteRegistroComprasPLE(pdFecha) Then
        N = MsgBox("Reporte de Registro de Compras ya fue generado. ¿Desea volver a Procesar?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Aviso")
        If N = vbCancel Then Exit Function
    End If
    If N = vbYes Then
        Call oReg.InsertaRegistroComprasPLE(pdFecha, CDate(gdFecSis & " " & Format(Time, "hh:mm:ss")), gsCodUser)
    End If

    lsSeparador = "|"
    BarraProgreso.value = 0
    BarraProgreso.Min = 0
    BarraProgreso.value = 0
    EstadoBarra.Panels(1) = "Generando..."

    Dim ArcSal1 As Integer
    ArcSal1 = FreeFile
    ArchTxtCompGast = App.path & "\SPOOLER\LE" & gsRUC & Format(pdFecha, "YYYYMM00") & "080100001111_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".txt"
    Open App.path & "\SPOOLER\LE" & gsRUC & Format(pdFecha, "YYYYMM00") & "080100001111_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".txt" For Output As ArcSal1

    Set rs = oReg.RecuperaDatosRegistroComprasPLE(pdFecha)
    BarraProgreso.Max = rs.RecordCount
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    
    lnNroRegistros = rs.RecordCount
    For i = 1 To rs.RecordCount
        lnLineaActual = lnLineaActual + 1
        lsCadena = rs!cDatoComp
        Print #1, lsCadena
        BarraProgreso.value = i
        EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
        rs.MoveNext
    Next
    Close ArcSal1
    RecuperaCadenaRegistroComprasNew = True
    EstadoBarra.Panels(1) = "Proceso Finalizado"
    Set rs = Nothing
    Set oReg = Nothing
End Function
'NAGL ERS 012-2017 20170710 FIN*******************************************
Public Function RecuperaCadenaRegistroVentas(ByVal pdFecha As Date) As String
    Dim oReg As New NContImpreReg
    Dim rs As New ADODB.Recordset
    Dim lsSeparador As String, lsCadena As String
    Dim lnNroRegistros As Long, lnLineaActual As Long
    Dim i As Long
    Dim N As Integer
    
    N = vbYes
    If oReg.ExisteRegistroVentasPLE(pdFecha) Then
        N = MsgBox("Reporte de Registro de Ventas ya fue generado. ¿Desea volver a Procesar?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Aviso")
        If N = vbCancel Then Exit Function
    End If
    If N = vbYes Then
        Call oReg.InsertaRegistroVentasPLE(pdFecha, CDate(gdFecSis & " " & Format(Time, "hh:mm:ss")), gsCodUser)
    End If
    
    lsSeparador = "|"
    BarraProgreso.value = 0
    BarraProgreso.Min = 0
    BarraProgreso.value = 0
    EstadoBarra.Panels(1) = "Generando..."
    
    Set rs = oReg.RecuperaDatosRegistroVentasPLE(pdFecha)
    BarraProgreso.Max = rs.RecordCount
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    
    lnNroRegistros = rs.RecordCount
    For i = 1 To rs.RecordCount
        lnLineaActual = lnLineaActual + 1
        'Comentado xPASI20160411*****************************************************
        'lsCadena = lsCadena & rs!cPeriodo & lsSeparador
        'lsCadena = lsCadena & rs!nMovNro & lsSeparador 'EJVG20140520
        'lsCadena = lsCadena & rs!cMovItem & lsSeparador 'EJVG20140520
        ''lsCadena = lsCadena & rs!nCorrelativo & lsSeparador
        'lsCadena = lsCadena & Format(rs!dFechaEmision, "dd/mm/yyyy") & lsSeparador
        'lsCadena = lsCadena & Format(rs!dFechaPago, "dd/mm/yyyy") & lsSeparador
        'lsCadena = lsCadena & rs!cTpoDoc & lsSeparador
        'lsCadena = lsCadena & rs!cNroSerie & lsSeparador
        'lsCadena = lsCadena & rs!cNroDocumento & lsSeparador
        'lsCadena = lsCadena & rs!cNroFinalRegTicket & lsSeparador
        'lsCadena = lsCadena & rs!cClienteTpoDocId & lsSeparador
        'lsCadena = lsCadena & rs!cClienteNroDocId & lsSeparador
        'lsCadena = lsCadena & rs!cClienteNombre & lsSeparador
        'lsCadena = lsCadena & Format(rs!nValFactExporta, "#0.00") & lsSeparador
        'lsCadena = lsCadena & Format(rs!nBaseImponible, "#0.00") & lsSeparador
        'lsCadena = lsCadena & Format(rs!nImpTotOpeExonerada, "#0.00") & lsSeparador
        'lsCadena = lsCadena & Format(rs!nOperaInafecta, "#0.00") & lsSeparador
        'lsCadena = lsCadena & Format(rs!nImpSelectConsumo, "#0.00") & lsSeparador
        'lsCadena = lsCadena & Format(rs!nIGV, "#0.00") & lsSeparador
        'lsCadena = lsCadena & Format(rs!nBaseImponibleArrozPillado, "#0.00") & lsSeparador
        'lsCadena = lsCadena & Format(rs!nImpVentaArrozPillado, "#0.00") & lsSeparador
        'lsCadena = lsCadena & Format(rs!nOtrosTrib, "#0.00") & lsSeparador
        'lsCadena = lsCadena & Format(rs!nImporteTotal, "#0.00") & lsSeparador
        'lsCadena = lsCadena & Format(rs!nTpoCambio, "#0.000") & lsSeparador
        'lsCadena = lsCadena & IIf(DateDiff("D", rs!dFechaPagoMod, CDate("1900-01-01")) = 0, "01/01/0001", Format(rs!dFechaPagoMod, "dd/mm/yyyy")) & lsSeparador
        'lsCadena = lsCadena & rs!cTpoDocMod & lsSeparador
        'lsCadena = lsCadena & rs!cNroSerieMod & lsSeparador
        'lsCadena = lsCadena & rs!cNroDocumentoMod & lsSeparador
        'lsCadena = lsCadena & Format(rs!nFOB, "#0.00") & lsSeparador 'EJVG20140520
        'lsCadena = lsCadena & rs!cEstado & lsSeparador
        
        'PASI20160411********************************
        lsCadena = lsCadena & Replace(Replace(Replace(rs!cDato, vbCrLf, ""), Chr$(13), ""), Chr$(10), "")
        lsCadena = lsCadena & lsSeparador
        
        If lnLineaActual <> lnNroRegistros Then
            lsCadena = lsCadena & Chr(10)
        End If
        BarraProgreso.value = i
        EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
        rs.MoveNext
    Next
    RecuperaCadenaRegistroVentas = lsCadena
    EstadoBarra.Panels(1) = "Proceso Finalizado"
    Set rs = Nothing
    Set oReg = Nothing
End Function
'PASI 20140531
Public Function RecuperaCadenaLibroDiarioDetCtaCont(ByVal pdFecha As Date) As String
    Dim oReg As New NContImpreReg
    Dim rs As New ADODB.Recordset
    Dim lsSeparador As String, lsCadena As String
    Dim lnNroRegistros As Long, lnLineaActual As Long
    Dim i As Long
    Dim N As Integer
    
    N = vbYes
    lsSeparador = "|"
    
    Set rs = oReg.RecuperaDatosLibroDiarioDetCtaContPLE(pdFecha)
    
    lnNroRegistros = rs.RecordCount
    For i = 1 To rs.RecordCount
        lnLineaActual = lnLineaActual + 1
        lsCadena = lsCadena & Replace(Replace(Replace(rs!cDato, vbCrLf, ""), Chr$(13), ""), Chr$(10), "")
        lsCadena = lsCadena & lsSeparador
        If lnLineaActual <> lnNroRegistros Then
            lsCadena = lsCadena & Chr(10)
        End If
        rs.MoveNext
    Next
    RecuperaCadenaLibroDiarioDetCtaCont = lsCadena
    Set rs = Nothing
    Set oReg = Nothing
End Function
'END PASI
'ALPA 20130529*********************************
Public Function RecuperaCadenaLibroDiario(ByVal pdFecha As Date) As String
    Dim oReg As New NContImpreReg
    Dim rs As New ADODB.Recordset
    Dim lsSeparador As String, lsCadena As String
    Dim lnNroRegistros As Long, lnLineaActual As Long
    Dim i As Long
    Dim N As Integer
    
    N = vbYes
    If oReg.ExisteLibroDiarioPLE(pdFecha) Then
        N = MsgBox("Reporte de Libro de Ventas ya fue generado. ¿Desea volver a Procesar?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Aviso")
        If N = vbCancel Then Exit Function
    End If
    If N = vbYes Then
        Call oReg.InsertaLibroDiarioPLE(pdFecha, CDate(gdFecSis & " " & Format(Time, "hh:mm:ss")), gsCodUser)
    End If
    
    lsSeparador = "|"
    BarraProgreso.value = 0
    BarraProgreso.Min = 0
    BarraProgreso.value = 0
    EstadoBarra.Panels(1) = "Generando..."
    
    Set rs = oReg.RecuperaDatosLibroDiarioPLE(pdFecha)
    BarraProgreso.Max = rs.RecordCount
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    
    lnNroRegistros = rs.RecordCount
    For i = 1 To rs.RecordCount
        lnLineaActual = lnLineaActual + 1
        lsCadena = lsCadena & Replace(Replace(Replace(rs!cDato, vbCrLf, ""), Chr$(13), ""), Chr$(10), "")
        lsCadena = lsCadena & lsSeparador
        If lnLineaActual <> lnNroRegistros Then
            lsCadena = lsCadena & Chr(10)
        End If
        BarraProgreso.value = i
        EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
        rs.MoveNext
    Next
    RecuperaCadenaLibroDiario = lsCadena
    EstadoBarra.Panels(1) = "Proceso Finalizado"
    Set rs = Nothing
    Set oReg = Nothing
End Function
Public Function RecuperaCadenaLibroMayor(ByVal pdFecha As Date) As String
    Dim oReg As New NContImpreReg
    Dim rs As New ADODB.Recordset
    Dim lsSeparador As String, lsCadena As String
    Dim lnNroRegistros As Long, lnLineaActual As Long
    Dim i As Long
    Dim N As Integer
    
    N = vbYes
    If oReg.ExisteLibroMayorPLE(pdFecha) Then
        N = MsgBox("Reporte de Libro de Ventas ya fue generado. ¿Desea volver a Procesar?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Aviso")
        If N = vbCancel Then Exit Function
    End If
    If N = vbYes Then
        Call oReg.InsertaLibroMayorPLE(pdFecha, CDate(gdFecSis & " " & Format(Time, "hh:mm:ss")), gsCodUser)
    End If
    
    lsSeparador = "|"
    BarraProgreso.value = 0
    BarraProgreso.Min = 0
    BarraProgreso.value = 0
    EstadoBarra.Panels(1) = "Generando..."
    
    Set rs = oReg.RecuperaDatosLibroMayorPLE(pdFecha)
    BarraProgreso.Max = rs.RecordCount
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    
    lnNroRegistros = rs.RecordCount
    For i = 1 To rs.RecordCount
        lnLineaActual = lnLineaActual + 1
        lsCadena = lsCadena & Replace(Replace(Replace(rs!cDato, vbCrLf, ""), Chr$(13), ""), Chr$(10), "")
        lsCadena = lsCadena & lsSeparador
        If lnLineaActual <> lnNroRegistros Then
            lsCadena = lsCadena & Chr(10)
        End If
        BarraProgreso.value = i
        EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
        rs.MoveNext
    Next
    RecuperaCadenaLibroMayor = lsCadena
    EstadoBarra.Panels(1) = "Proceso Finalizado"
    Set rs = Nothing
    Set oReg = Nothing
End Function
'**********************************************
