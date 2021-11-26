VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRRHHRep 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "frmRRHHRep.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flex 
      Height          =   270
      Left            =   1620
      TabIndex        =   8
      Top             =   5940
      Visible         =   0   'False
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   476
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CheckBox chkComprimido 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Comprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   45
      TabIndex        =   7
      Top             =   5940
      Width           =   1575
   End
   Begin MSMask.MaskEdBox mskFecIni 
      Height          =   300
      Left            =   1095
      TabIndex        =   0
      Top             =   52
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   4215
      TabIndex        =   4
      Top             =   5880
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskFecFin 
      Height          =   300
      Left            =   3945
      TabIndex        =   1
      Top             =   52
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSComctlLib.ListView lvwImp 
      Height          =   5415
      Left            =   0
      TabIndex        =   2
      Top             =   420
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   9551
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImaLis"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImaLis 
      Left            =   240
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRRHHRep.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRRHHRep.frx":079C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRRHHRep.frx":0B62
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   3060
      TabIndex        =   3
      Top             =   5880
      Width           =   1095
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   240
      Left            =   60
      SizeMode        =   1  'Stretch
      TabIndex        =   9
      Top             =   6225
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblFecFin 
      Caption         =   "Fecha Fin:"
      Height          =   255
      Left            =   3045
      TabIndex        =   6
      Top             =   75
      Width           =   855
   End
   Begin VB.Label lblFecINi 
      Caption         =   "Fecha Ini:"
      Height          =   255
      Left            =   135
      TabIndex        =   5
      Top             =   75
      Width           =   735
   End
End
Attribute VB_Name = "frmRRHHRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnNumRep As Integer
Dim lsCadenaBuscar As String
Dim lsRep() As String

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim WithEvents oRepEvento As NRHReportes
Attribute oRepEvento.VB_VarHelpID = -1
Dim Progress As clsProgressBar

Dim lsNombreCabecera As String

Private Sub oRepEvento_CloseProgress()
    Progress.CloseForm Me
End Sub

Private Sub oRepEvento_Progress(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Generando Reporte"
End Sub

Private Sub oRepEvento_ShowProgress()
    Progress.ShowForm Me
End Sub

Private Sub IniRep()
    Dim i As Integer
    Dim llAux As ListItem
    ReDim lsRep(lnNumRep)
    
    lsRep(1) = "Reporte 5TA CATEGORIA"
    lsRep(2) = "Listado para Contabilidad "
    lsRep(3) = "Vencimiento de Contratos/Adendas"
    lsRep(4) = "Relacion de Personal"
    lsRep(5) = "Relacion de Personal Por Agencias"
    'John lsRep(6) = "Relacion de Personal Solo Agencias"
    '----lsRep(7) = "Relacion de Personal Sede por Areas"
    lsRep(6) = "Tardanzas,Dias Vacaciones de Empleados"
    lsRep(7) = "Tardanzas de Empleados por Agencias"
    lsRep(8) = "Detalle Tardanzas de Emplados"
    '----lsRep(7) = "Ingresos de Empleados por fechas"
    '----lsRep(8) = "Descuentos de Empleados por fechas"
    '----lsRep(9) = "Valida Archivo para Planilla de AFP"
    '----lsRep(10) = "Archivo para Planilla de AFP"
    '---lsRep(8) = "Ingresos de Empleados x 3 Meses"
    lsRep(9) = "Certificados de 5ta Categoria"
    '---lsRep(13) = "Estad Detallada de RRHH Numero"
    '---lsRep(14) = "Estad Detallada de RRHH Monto"
    'agregado
    lsRep(10) = "Archivo de Texto para el PDT"
    lsRep(11) = "Boletas Consolidadas Planilla"
    lsRep(12) = "Reportes RRHH"
    '-----
    lsRep(13) = "Ingresos de Empleados por fechas"
    lsRep(14) = "Descuentos de Empleados por fechas"
    lsRep(15) = "Valida Archivo para Planilla de AFP"
    lsRep(16) = "Archivo para Planilla de AFP"
    lsRep(17) = "Ingresos de Empleados x 3 Meses"
    lsRep(18) = "PDT-Jornada Laboral"
    lsRep(19) = "PDT-Lugar de Trabajo"
    
    lvwImp.ListItems.Clear
    lvwImp.HideColumnHeaders = False
    
    lvwImp.ColumnHeaders.Clear
    lvwImp.ListItems.Clear
    
    lvwImp.ColumnHeaders.Add , , "Reporte", 3500
    lvwImp.View = lvwReport
        
    For i = 1 To lnNumRep
        Set llAux = lvwImp.ListItems.Add(, , lsRep(i), , 2)
    Next i
End Sub

Private Sub chkComprimido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAceptar.SetFocus
    End If
End Sub

Private Sub CmdAceptar_Click()
    Dim i As Integer
    Dim lsCadena As String
    Dim lsCad As String
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    Dim oRep As NRHReportes
    Set oRep = New NRHReportes
    Set oRepEvento = New NRHReportes
    If Not ValidaFecha Then Exit Sub
    For i = 1 To lvwImp.ListItems.Count
        If lvwImp.ListItems(i).Checked Then
            If i = 1 Then
                lsCad = oRepEvento.Rep5taRRHH(gdFecSis, Me.mskFecIni.Text, Me.mskFecFin.Text, gsEmpresa)
            ElseIf i = 2 Then
                lsCad = oRepEvento.RepContabilidadRRHH(gdFecSis, Me.mskFecIni.Text, Me.mskFecFin.Text, gsEmpresa)
            ElseIf i = 3 Then
                lsCad = oRepEvento.GetContratosVencidos(CDate(Me.mskFecIni.Text), DateAdd("d", 1, CDate(Me.mskFecFin.Text)), gsNomAge, gsEmpresa, gdFecSis)
            ElseIf i = 4 Then
                lsCad = oRepEvento.GetListaPersonal(Me.mskFecIni.Text, Me.mskFecFin.Text, gsNomAge, gsEmpresa, gdFecSis)
            ElseIf i = 5 Then
                lsCad = oRepEvento.GetListaPersonalPorAgencias(gsNomAge, gsEmpresa, gdFecSis)
            'John ElseIf i = 6 Then
            'John    lsCad = oRepEvento.GetListaPersonalSoloAgencias(Me.mskFecIni.Text, Me.mskFecFin.Text, gsNomAge, gsEmpresa, gdFecSis)
            '---ElseIf I = 7 Then
            '---    lsCad = oRepEvento.GetListaPersonalSoloSedeAreas(Me.mskFecIni.Text, Me.mskFecFin.Text, gsNomAge, gsEmpresa, gdFecSis)
            ElseIf i = 6 Then
                lsCad = oRepEvento.GetTardanzasEmpleados(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), gsNomAge, gsEmpresa, gdFecSis)
            ElseIf i = 7 Then
                lsCad = oRepEvento.GetTardanzasEmpleadosPorAgencias(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), gsNomAge, gsEmpresa, gdFecSis, "")
            ElseIf i = 8 Then
                Set oRRHH = New DActualizaDatosRRHH
                Set oPersona = New UPersona
                Set oPersona = frmBuscaPersona.Inicio(True)
                If oPersona Is Nothing Then Exit Sub
                lsCadenaBuscar = oPersona.sPersCod
                lsCad = oRepEvento.GetTardanzasEmpleadosDetalle(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), lsCadenaBuscar, gsNomAge, gsEmpresa, gdFecSis)
            ElseIf i = 9 Then
                 lsCadena = frmRHRep5taCat.Ini(Me)
            'ElseIf I = 13 Then
                 'GetEstadDetRRHHNum
            'ElseIf I = 14 Then
                 'GetEstadDetRRHHMonto
            ElseIf i = 10 Then  'PDT
                 'frmRHformatoPDT.Show vbModal
                 lsCad = oRepEvento.GetRepPlaPDT(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text))
            ElseIf i = 11 Then  'BOLETAS CONSOLIDADAS
                 frmRHRepConsolBoleta.Show
            ElseIf i = 12 Then  'BOLETAS CONSOLIDADAS
                 frmRHReportes.Show
                 Unload Me
            ElseIf i = 13 Then
                GetIngresosConceptos
            ElseIf i = 14 Then
                 GetDescuentosConceptos
            ElseIf i = 15 Then
                lsCad = oRepEvento.GetRepPlaAFPValida(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), gsNomAge, gsEmpresa, gdFecSis)
            ElseIf i = 16 Then
                lsCad = oRepEvento.GetRepPlaAFP(gsRUC, CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text))
            ElseIf i = 17 Then
                GetIngresosConceptos_x_3
             ElseIf i = 18 Then                 '
                 lsCad = oRepEvento.GetJorLabPDT(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text))
            ElseIf i = 19 Then                 '
                 lsCad = oRepEvento.GetEstabPDT(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text))
            End If
            
            If lsCad <> "" Then
                If lsCadena <> "" Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & lsCad
            Else
               MsgBox "No existen datos para Imprimir", vbInformation, "Aviso"
            End If
        End If
    Next i
    If lsCadena <> "" Then
        oPrevio.Show lsCadena, "Reportes de RRHH", Me.chkComprimido.value, 66, gImpresora
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sqlAge As String
    Dim rsAge As New ADODB.Recordset
    Set Progress = New clsProgressBar
    
    lnNumRep = 19
    mskFecIni.Text = Format(gdFecSis, gsFormatoFechaView)
    mskFecFin.Text = Format(gdFecSis, gsFormatoFechaView)
    IniRep
End Sub

Private Sub mskFecFin_GotFocus()
    mskFecFin.SelStart = 0
    mskFecFin.SelLength = 10
End Sub

Private Sub mskFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkComprimido.SetFocus
    End If
End Sub

Private Sub mskFecIni_GotFocus()
    mskFecIni.SelStart = 0
    mskFecIni.SelLength = 10
End Sub

Private Sub lvwImp_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.SmallIcon = 1
    Else
        Item.SmallIcon = 2
    End If
End Sub


Private Sub mskFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFecFin.SetFocus
    End If
End Sub

Private Function ValidaFecha() As Boolean
    If Not IsDate(mskFecIni) Then
        MsgBox "La fecha de inicio no es correcta.", vbInformation, "Aviso"
        mskFecIni.SetFocus
        ValidaFecha = False
    ElseIf Not IsDate(mskFecFin) Then
        MsgBox "La fecha de fin no es correcta.", vbInformation, "Aviso"
        mskFecFin.SetFocus
        ValidaFecha = False
    Else
        ValidaFecha = True
    End If
End Function


Private Sub GetEstadDetRRHHNum()
    Dim rsD As ADODB.Recordset
    Set rsD = New ADODB.Recordset
    Dim rsI As ADODB.Recordset
    Set rsI = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean

    Call oRepEvento.GetEstadDetRRHHNum(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), rsI, rsD)
    
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFecFin.Text), "yyyymmdd") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis & "  " & Time, "yyyymmddmmhhss"), xlLibro, xlHoja1
       Call GeneraReportePlanea(rsD, rsI, False)
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
End Sub

Private Sub GetEstadDetRRHHMonto()
    Dim rsD As ADODB.Recordset
    Set rsD = New ADODB.Recordset
    Dim rsI As ADODB.Recordset
    Set rsI = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean

    Call oRepEvento.GetEstadDetRRHHMonto(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), rsI, rsD)
    
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFecFin.Text), "yyyymmdd") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis & "  " & Time, "yyyymmddmmhhss"), xlLibro, xlHoja1
       Call GeneraReportePlanea(rsD, rsI, True)
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
End Sub



Private Sub GetIngresosConceptos()
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean

    Set rsE = oRepEvento.GetIngresosConceptos(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text))
    
    If rsE Is Nothing Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If rsE.EOF And rsE.BOF Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
    Else
        lsArchivoN = App.path & "\Spooler\Ingresos_" & Format(CDate(Me.mskFecFin.Text) & " " & Time, "yyyymmddhhmmss") & ".xls"
        OleExcel.Class = "ExcelWorkSheet"
        lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
        If lbLibroOpen Then
           Set xlHoja1 = xlLibro.Worksheets(1)
           ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
           Call GeneraReporte(rsE)
           OleExcel.Class = "ExcelWorkSheet"
           ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
           OleExcel.SourceDoc = lsArchivoN
           OleExcel.Verb = 1
           OleExcel.Action = 1
           OleExcel.DoVerb -1
        End If
        MousePointer = 0
    End If
End Sub

Private Sub GetIngresosConceptos_x_3()
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean

    Set rsE = oRepEvento.GetIngresosConceptos(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), 3)
    
    If rsE Is Nothing Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If rsE.EOF And rsE.BOF Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
    Else
        lsArchivoN = App.path & "\Spooler\Ingresos_x_3" & Format(CDate(Me.mskFecFin.Text) & " " & Time, "yyyymmddhhmmss") & ".xls"
        OleExcel.Class = "ExcelWorkSheet"
        lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
        If lbLibroOpen Then
           Set xlHoja1 = xlLibro.Worksheets(1)
           ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
           Call GeneraReporte(rsE)
           OleExcel.Class = "ExcelWorkSheet"
           ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
           OleExcel.SourceDoc = lsArchivoN
           OleExcel.Verb = 1
           OleExcel.Action = 1
           OleExcel.DoVerb -1
        End If
        MousePointer = 0
    End If
End Sub


Private Sub GetDescuentosConceptos()
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean

    Set rsE = oRepEvento.GetDescuentosConceptos(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text))
    
    If rsE Is Nothing Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If rsE.EOF And rsE.BOF Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
    Else
        lsArchivoN = App.path & "\Spooler\Descuentos_" & Format(CDate(Me.mskFecFin.Text) & " " & Time, "yyyymmddhhmmss") & ".xls"
        OleExcel.Class = "ExcelWorkSheet"
        lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
        If lbLibroOpen Then
           Set xlHoja1 = xlLibro.Worksheets(1)
           ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
           Call GeneraReporte(rsE)
           OleExcel.Class = "ExcelWorkSheet"
           ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
           OleExcel.SourceDoc = lsArchivoN
           OleExcel.Verb = 1
           OleExcel.Action = 1
           OleExcel.DoVerb -1
        End If
        MousePointer = 0
    End If
End Sub


Private Function BuscaCol(pflex As MSHFlexGrid, psCon As String) As Integer
    Dim lnI As Integer
    For lnI = 2 To pflex.Cols - 1
        If pflex.TextMatrix(0, lnI) = psCon Then
            BuscaCol = lnI
            lnI = pflex.Cols - 1
            Exit Function
        End If
    Next lnI
End Function

Private Sub GeneraReporte(prRs As ADODB.Recordset)
    Dim i As Integer
    Dim K As Integer
    Dim J As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    Dim lnPosX As Integer
    Dim lnPosY As Integer
    
    i = -1
    prRs.MoveFirst
    While Not prRs.EOF
        i = i + 1
        For J = 1 To prRs.Fields.Count
            xlHoja1.Cells(i + 1, J + 1) = Format(prRs.Fields(J - 1), "#,##0.00")
            If IsNumeric(prRs.Fields(J - 1)) And i > 2 Then
                lnAcum = lnAcum + CCur(prRs.Fields(J - 1))
            End If
        Next J
        If i > 2 Then
            xlHoja1.Cells(i + 1, 1) = i - 2
            lsSuma = Format(lnAcum, "#,##0.00")
            xlHoja1.Cells(i + 1, prRs.Fields.Count + 1 + 1) = lsSuma
            lnAcum = 0
        End If
        prRs.MoveNext
    Wend
        
    xlHoja1.Range("A1:A" & Trim(Str(i + 1))).Font.Bold = True
    xlHoja1.Range("B1:B" & Trim(Str(i + 1))).Font.Bold = True
    
    For lnPosX = 4 To prRs.Fields.Count + 2
        lnAcum = 0
        For lnPosY = 4 To i + 1
            lnAcum = lnAcum + xlHoja1.Cells(lnPosY, lnPosX)
        Next lnPosY
        xlHoja1.Cells(i + 2, lnPosX) = lnAcum
    Next lnPosX
    
    With xlHoja1.PageSetup
        .LeftHeader = gsEmpresaCompleto
        .CenterHeader = "&""Arial,Negrita""&18" & lsNombreCabecera & "(" & Format(gdFecSis) & ")" & "  " & Me.mskFecIni.Text & " - " & Me.mskFecFin.Text
        .RightHeader = "&P"
        .LeftFooter = ""
        .CenterFooter = ""
'        .RightFooter = ""
'        .LeftMargin = Application.InchesToPoints(0.748031496062992)
'        .RightMargin = Application.InchesToPoints(0.748031496062992)
'        .TopMargin = Application.InchesToPoints(0.984251968503937)
'        .BottomMargin = Application.InchesToPoints(0.984251968503937)
'        .HeaderMargin = Application.InchesToPoints(0.511811023622047)
'        .FooterMargin = Application.InchesToPoints(0.511811023622047)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 65
        '.PrintErrors = xlPrintErrorsDisplayed
    End With
    
    xlHoja1.Range("C1:U3").Font.Bold = True
    xlHoja1.Range("C4:" & ExcelColumnaString(prRs.Fields.Count) & Trim(Str(i + 1))).NumberFormat = "#,##0.00"
    
    xlHoja1.Columns.AutoFit
    
    'xlHoja1.Range("A1:B5").Merge
    
    
End Sub


Public Sub Ini(psCaption As String, pForm As Form)
    Caption = psCaption
    Show 0, pForm
End Sub

Private Sub GeneraReportePlanea(prD As ADODB.Recordset, prI As ADODB.Recordset, pbMonto As Boolean)
    Dim i As Integer
    Dim K As Integer
    Dim J As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    
    Set rsC = oGen.GetConstante(6042)
    
    i = -1
    prD.MoveFirst
    
    i = i + 1
    
    If Not pbMonto Then
        xlHoja1.Cells(i + 1, 3) = "NUMERO DE PERSONAL"
    Else
        xlHoja1.Cells(i + 1, 3) = "REMUNERACION DE PERSONAL"
    End If
    xlHoja1.Range("C1:" & ExcelColumnaString(prD.RecordCount - 4) & "1").Merge
    xlHoja1.Range("C1:" & ExcelColumnaString(prD.RecordCount - 4) & "1").HorizontalAlignment = xlCenter
    xlHoja1.Range("C1:" & ExcelColumnaString(prD.RecordCount - 4) & "2").Font.Bold = True
    
    i = i + 1
    xlHoja1.Cells(i + 1, 1) = "AGE/AREA"
    xlHoja1.Cells(i + 1, 2) = "UBICACION"
    
    xlHoja1.Range("A1:A2").Merge
    xlHoja1.Range("A1:A2").HorizontalAlignment = xlCenter
    xlHoja1.Range("B1:B2").Merge
    xlHoja1.Range("B1:B2").HorizontalAlignment = xlCenter
    
    J = 2
    While Not rsC.EOF
        J = J + 1
        xlHoja1.Cells(i + 1, J) = Trim(Left(rsC.Fields(0), 20))
        rsC.MoveNext
    Wend
    
    i = i + 1
    xlHoja1.Cells(i + 1, 2) = "DIRECTOS"
    xlHoja1.Range("A" & i + 1 & ":B" & i + 1).Merge
    xlHoja1.Range("A" & i + 1 & ":B" & i + 1).HorizontalAlignment = xlCenter
 
    For J = 2 To prD.Fields.Count - 1
        xlHoja1.Range("" & ExcelColumnaString(J + 1) & "3").FormulaR1C1 = "=SUM(R[1]C:R[" & prD.RecordCount & "]C)"
        xlHoja1.Range("" & ExcelColumnaString(J + 1) & "3").Interior.ColorIndex = 34
        xlHoja1.Range("" & ExcelColumnaString(J + 1) & "3").Font.Color = &H80&
        xlHoja1.Range("" & ExcelColumnaString(J + 1) & "3").Font.Bold = True
    
    Next J
     
    While Not prD.EOF
        i = i + 1
        xlHoja1.Cells(i + 1, 1) = prD.Fields(0)
        xlHoja1.Cells(i + 1, 2) = prD.Fields(1)
        
        For J = 2 To prD.Fields.Count - 1
            If pbMonto Then
                xlHoja1.Cells(i + 1, J + 1) = Format(prD.Fields(J), "#,##0.00")
            Else
                xlHoja1.Cells(i + 1, J + 1) = Format(prD.Fields(J), "#,##0")
            End If
        Next J
        prD.MoveNext
    Wend
        
    i = i + 1
    xlHoja1.Cells(i + 1, 2) = "INDIRECTOS"
    xlHoja1.Range("A" & i + 1 & ":B" & i + 1).Merge
    xlHoja1.Range("A" & i + 1 & ":B" & i + 1).HorizontalAlignment = xlCenter
        
    For J = 2 To prD.Fields.Count - 1
        xlHoja1.Range("" & ExcelColumnaString(J + 1) & prD.RecordCount + 4).FormulaR1C1 = "=SUM(R[1]C:R[" & prI.RecordCount & "]C)"
        xlHoja1.Range("" & ExcelColumnaString(J + 1) & prD.RecordCount + 4).Interior.ColorIndex = 34
        xlHoja1.Range("" & ExcelColumnaString(J + 1) & prD.RecordCount + 4).Font.Color = &H80&
        xlHoja1.Range("" & ExcelColumnaString(J + 1) & prD.RecordCount + 4).Font.Bold = True
    Next J
        
    While Not prI.EOF
        i = i + 1
        xlHoja1.Cells(i + 1, 1) = prI.Fields(0)
        xlHoja1.Cells(i + 1, 2) = prI.Fields(1)
        
        For J = 2 To prI.Fields.Count - 3
            If prI.Fields(3) < 3 Then
                xlHoja1.Cells(i + 1, J + 1) = ""
            Else
                If pbMonto Then
                    xlHoja1.Cells(i + 1, J + 1) = Format(prI.Fields(J + 2), "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, J + 1) = Format(prI.Fields(J + 2), "#,##0")
                End If
            End If
        Next J
        prI.MoveNext
    Wend
        
    xlHoja1.Range("A1:A" & Trim(Str(i + 1))).Font.Bold = True
    xlHoja1.Range("B1:B" & Trim(Str(i + 1))).Font.Bold = True
    
    xlHoja1.Columns.AutoFit
    
    xlHoja1.Range("A1:" & ExcelColumnaString(prD.Fields.Count) & prD.RecordCount + prI.RecordCount + 4).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A1:" & ExcelColumnaString(prD.Fields.Count) & prD.RecordCount + prI.RecordCount + 4).Borders(xlDiagonalUp).LineStyle = xlNone
    With xlHoja1.Range("A1:" & ExcelColumnaString(prD.Fields.Count) & prD.RecordCount + prI.RecordCount + 4).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:" & ExcelColumnaString(prD.Fields.Count) & prD.RecordCount + prI.RecordCount + 4).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:" & ExcelColumnaString(prD.Fields.Count) & prD.RecordCount + prI.RecordCount + 4).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:" & ExcelColumnaString(prD.Fields.Count) & prD.RecordCount + prI.RecordCount + 4).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:" & ExcelColumnaString(prD.Fields.Count) & prD.RecordCount + prI.RecordCount + 4).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:" & ExcelColumnaString(prD.Fields.Count) & prD.RecordCount + prI.RecordCount + 4).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = _
        "&""Arial,Negrita""&14ESTADISTICA DETALLADA DE RECURSOS HUMANOS"
        .RightHeader = "MES DE : " & Format(CDate(Me.mskFecIni.Text), "MMMM") & " - " & Format(CDate(Me.mskFecIni.Text), "YYYY")
        .LeftFooter = ""
        .CenterFooter = ""
'        .RightFooter = ""
'        .LeftMargin = Application.InchesToPoints(0)
'        .RightMargin = Application.InchesToPoints(0)
'        .TopMargin = Application.InchesToPoints(0.86)
'        .BottomMargin = Application.InchesToPoints(0)
'        .HeaderMargin = Application.InchesToPoints(0.41)
'        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 75
    End With
    
    If pbMonto Then
        xlHoja1.Range("C3:" & ExcelColumnaString(prD.Fields.Count) & prD.RecordCount + prI.RecordCount + 4).NumberFormat = "#,##0.00"
    Else
        xlHoja1.Range("C3:" & ExcelColumnaString(prD.Fields.Count) & prD.RecordCount + prI.RecordCount + 4).NumberFormat = "#,##0"
    End If
End Sub
