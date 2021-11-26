VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReporteEstadFoncodes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información estadística Fideicomiso FONCODES"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   Icon            =   "frmReporteEstadFoncodes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   195
      Width           =   1215
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "Reporte"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   195
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblRep 
      Caption         =   "Reporte al:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblFechaReporte 
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmReporteEstadFoncodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************
'MIOL 20121102, SEGUN RFC103- 2012
'************************************************

Private Sub cmdReporte_Click()
    If Mid(gdFecSis, 4, 2) = "01" Or Mid(gdFecSis, 4, 2) = "04" Or Mid(gdFecSis, 4, 2) = "07" Or Mid(gdFecSis, 4, 2) = "10" Then
        Call ReporteTrimestral(gdFecSis, 3)
    Else
        Call ReporteSaldoConsolidado
    End If
    Unload Me
End Sub

Public Sub ReporteSaldoConsolidado()
    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja  As String
    Dim lsNomHoja1  As String
    Dim lsNomHoja2  As String
    Dim lsArchivo As String
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim xlHoja2 As Excel.Worksheet
    Dim xlHoja3 As Excel.Worksheet
    
    Dim rsCreditos As ADODB.Recordset
    Dim rsCreditosMes As ADODB.Recordset
    Dim oCreditos As New DCreditos
       
    Dim sFecha As String
    Dim nPase As Integer
    
    PB1.Min = 0
    PB1.Max = 10
    PB1.value = 0
    PB1.Visible = True
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ReporteCOFO"
    'Primera Hoja ******************************************************
    lsNomHoja = "ReporteCOFO"
    'Segunda Hoja ******************************************************
    lsNomHoja1 = "BaseDatos"
    'Tercera Hoja ******************************************************
    lsNomHoja2 = "InfFinanciera"
    '*******************************************************************
    lsArchivo1 = "\spooler\LineaCredito" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & "FONCODES" & ".xls"
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
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    sFecha = DateAdd("D", 1, Format(lblFechaReporte.Caption, "DD/MM/YYYY"))
    xlHoja1.Cells(10, 4) = DateAdd("M", -1, Format(sFecha, "DD/MM/YYYY")) & " Al " & lblFechaReporte.Caption
    xlHoja1.Cells(10, 6) = lblFechaReporte.Caption
    
    PB1.value = 1
    'Creditos por Personeria
    'Saldo
    Set rsCreditos = oCreditos.ReporteLineaCredito_Personeria(gdFecSis, "04")
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If

    If nPase = 1 Then
        Do While Not rsCreditos.EOF
            If rsCreditos!Concepto = "PerNaturalFemenino" Then
                xlHoja1.Cells(15, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(15, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "PerNaturalMasculino" Then
                xlHoja1.Cells(16, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(16, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "PerJuridica" Then
                xlHoja1.Cells(17, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(17, 7) = rsCreditos!Monto
            End If
            rsCreditos.MoveNext
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    PB1.value = 2
    'Colocaciones
    Set rsCreditosMes = oCreditos.ReporteLineaCredito_PersoneriaMes(gdFecSis, "04")

    nPase = 1
    If (rsCreditosMes Is Nothing) Then
        nPase = 0
    End If

    If nPase = 1 Then
        Do While Not rsCreditosMes.EOF
            If rsCreditosMes!Concepto = "PerNaturalFemenino" Then
                xlHoja1.Cells(15, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(15, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "PerNaturalMasculino" Then
                xlHoja1.Cells(16, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(16, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "PerJuridica" Then
                xlHoja1.Cells(17, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(17, 5) = rsCreditosMes!Monto
            End If
            rsCreditosMes.MoveNext
            If rsCreditosMes.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditosMes.Close
    End If
    Set rsCreditosMes = Nothing

    PB1.value = 3
    'Creditos por Sector Economico
    Set rsCreditos = oCreditos.ReporteLineaCredito_SectorEconomico(gdFecSis, "04")
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If

    If nPase = 1 Then
        Do While Not rsCreditos.EOF
            If rsCreditos!Concepto = "Agricola" Then
                xlHoja1.Cells(19, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(19, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "Pecuario" Then
                xlHoja1.Cells(20, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(20, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "Comercio" Then
                xlHoja1.Cells(21, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(21, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "Produccion" Then
                xlHoja1.Cells(22, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(22, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "Pesca" Then
                xlHoja1.Cells(23, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(23, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "Servicios" Then
                xlHoja1.Cells(24, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(24, 7) = rsCreditos!Monto
            End If
            rsCreditos.MoveNext
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing

    PB1.value = 4
    Set rsCreditosMes = oCreditos.ReporteLineaCredito_SectorEconomicoMes(gdFecSis, "04")
    nPase = 1
    If (rsCreditosMes Is Nothing) Then
        nPase = 0
    End If

    If nPase = 1 Then
        Do While Not rsCreditosMes.EOF
            If rsCreditosMes!Concepto = "Agricola" Then
                xlHoja1.Cells(19, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(19, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "Pecuario" Then
                xlHoja1.Cells(20, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(20, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "Comercio" Then
                xlHoja1.Cells(21, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(21, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "Produccion" Then
                xlHoja1.Cells(22, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(22, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "Pesca" Then
                xlHoja1.Cells(23, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(23, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "Servicios" Then
                xlHoja1.Cells(24, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(24, 5) = rsCreditosMes!Monto
            End If
            rsCreditosMes.MoveNext
            If rsCreditosMes.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditosMes.Close
    End If
    Set rsCreditosMes = Nothing

    'Creditos por Ubicacion Geografica
    'Saldos
    PB1.value = 5
    Set rsCreditos = oCreditos.ReporteLineaCredito_UbicaGeografica(gdFecSis, "04")
     nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If

    If nPase = 1 Then
        Do While Not rsCreditos.EOF
            If rsCreditos!Concepto = "BAGUA" Then
                xlHoja1.Cells(31, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(31, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "RODRIGUEZ DE MENDOZA" Then
                xlHoja1.Cells(32, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(32, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "CHACHAPOYAS" Then
                xlHoja1.Cells(33, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(33, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "CAJAMARCA" Then
                xlHoja1.Cells(34, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(34, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "SAN MIGUEL" Then
                xlHoja1.Cells(35, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(35, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "AMBO" Then
                xlHoja1.Cells(36, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(36, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "DOS DE MAYO" Then
                xlHoja1.Cells(37, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(37, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "HUAMALIES" Then
                xlHoja1.Cells(38, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(38, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "HUANUCO" Then
                xlHoja1.Cells(39, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(39, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "LEONCIO PRADO" Then
                xlHoja1.Cells(40, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(40, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "PACHITEA" Then
                xlHoja1.Cells(41, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(41, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "PUERTO INCA" Then
                xlHoja1.Cells(42, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(42, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "MARAÑON" Then
                xlHoja1.Cells(43, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(43, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "JUNIN" Then
                xlHoja1.Cells(44, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(44, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "ALTO AMAZONAS" Then
                xlHoja1.Cells(45, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(45, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "DATEM DEL MARAÑON" Then
                xlHoja1.Cells(46, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(46, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "LORETO" Then
                xlHoja1.Cells(47, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(47, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "MARISCAL RAMON CASTILLA" Then
                xlHoja1.Cells(48, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(48, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "MAYNAS" Then
                xlHoja1.Cells(49, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(49, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "REQUENA" Then
                xlHoja1.Cells(50, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(50, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "UCAYALI" Then
                xlHoja1.Cells(51, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(51, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "DANIEL ALCIDES CARRION" Then
                xlHoja1.Cells(52, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(52, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "PASCO" Then
                xlHoja1.Cells(53, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(53, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "EL DORADO" Then
                xlHoja1.Cells(54, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(54, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "LAMAS" Then
                xlHoja1.Cells(55, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(55, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "MARISCAL CACERES" Then
                xlHoja1.Cells(56, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(56, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "RIOJA" Then
                xlHoja1.Cells(57, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(57, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "SAN MARTIN" Then
                xlHoja1.Cells(58, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(58, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "BELLAVISTA" Then
                xlHoja1.Cells(59, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(59, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "HUALLAGA" Then
                xlHoja1.Cells(60, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(60, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "MOYOBAMBA" Then
                xlHoja1.Cells(61, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(61, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "PICOTA" Then
                xlHoja1.Cells(62, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(62, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "TOCACHE" Then
                xlHoja1.Cells(63, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(63, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "ATALAYA" Then
                xlHoja1.Cells(64, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(64, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "CORONEL PORTILLO" Then
                xlHoja1.Cells(65, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(65, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "PADRE ABAD" Then
                xlHoja1.Cells(66, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(66, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "Callao" Then
                xlHoja1.Cells(67, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(67, 7) = rsCreditos!Monto
            End If
            If rsCreditos!Concepto = "Lima" Then
                xlHoja1.Cells(68, 6) = rsCreditos!Cantidad
                xlHoja1.Cells(68, 7) = rsCreditos!Monto
            End If
            rsCreditos.MoveNext
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing

    'Colocaciones
    PB1.value = 6
    Set rsCreditosMes = oCreditos.ReporteLineaCredito_UbicaGeograficaMes(gdFecSis, "04")
    nPase = 1
    If (rsCreditosMes Is Nothing) Then
        nPase = 0
    End If

    If nPase = 1 Then
        Do While Not rsCreditosMes.EOF
            If rsCreditosMes!Concepto = "BAGUA" Then
                xlHoja1.Cells(31, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(31, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "RODRIGUEZ DE MENDOZA" Then
                xlHoja1.Cells(32, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(32, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "CHACHAPOYAS" Then
                xlHoja1.Cells(33, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(33, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "CAJAMARCA" Then
                xlHoja1.Cells(34, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(34, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "SAN MIGUEL" Then
                xlHoja1.Cells(35, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(35, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "AMBO" Then
                xlHoja1.Cells(36, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(36, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "DOS DE MAYO" Then
                xlHoja1.Cells(37, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(37, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "HUAMALIES" Then
                xlHoja1.Cells(38, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(38, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "HUANUCO" Then
                xlHoja1.Cells(39, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(39, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "LEONCIO PRADO" Then
                xlHoja1.Cells(40, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(40, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "PACHITEA" Then
                xlHoja1.Cells(41, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(41, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "PUERTO INCA" Then
                xlHoja1.Cells(42, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(42, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "MARAÑON" Then
                xlHoja1.Cells(43, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(43, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "JUNIN" Then
                xlHoja1.Cells(44, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(44, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "ALTO AMAZONAS" Then
                xlHoja1.Cells(45, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(45, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "DATEM DEL MARAÑON" Then
                xlHoja1.Cells(46, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(46, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "LORETO" Then
                xlHoja1.Cells(47, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(47, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "MARISCAL RAMON CASTILLA" Then
                xlHoja1.Cells(48, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(48, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "MAYNAS" Then
                xlHoja1.Cells(49, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(49, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "REQUENA" Then
                xlHoja1.Cells(50, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(50, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "UCAYALI" Then
                xlHoja1.Cells(51, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(51, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "DANIEL ALCIDES CARRION" Then
                xlHoja1.Cells(52, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(52, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "PASCO" Then
                xlHoja1.Cells(53, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(53, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "EL DORADO" Then
                xlHoja1.Cells(54, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(54, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "LAMAS" Then
                xlHoja1.Cells(55, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(55, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "MARISCAL CACERES" Then
                xlHoja1.Cells(56, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(56, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "RIOJA" Then
                xlHoja1.Cells(57, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(57, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "SAN MARTIN" Then
                xlHoja1.Cells(58, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(58, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "BELLAVISTA" Then
                xlHoja1.Cells(59, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(59, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "HUALLAGA" Then
                xlHoja1.Cells(60, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(60, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "MOYOBAMBA" Then
                xlHoja1.Cells(61, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(61, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "PICOTA" Then
                xlHoja1.Cells(62, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(62, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "TOCACHE" Then
                xlHoja1.Cells(63, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(63, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "ATALAYA" Then
                xlHoja1.Cells(64, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(64, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "CORONEL PORTILLO" Then
                xlHoja1.Cells(65, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(65, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "PADRE ABAD" Then
                xlHoja1.Cells(66, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(66, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "Callao" Then
                xlHoja1.Cells(67, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(67, 5) = rsCreditosMes!Monto
            End If
            If rsCreditosMes!Concepto = "Lima" Then
                xlHoja1.Cells(68, 4) = rsCreditosMes!Cantidad
                xlHoja1.Cells(68, 5) = rsCreditosMes!Monto
            End If
            rsCreditosMes.MoveNext
            If rsCreditosMes.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditosMes.Close
    End If
    Set rsCreditosMes = Nothing

    PB1.value = 7
    'REPORTE BASE DE DATOS ************************************************
    Dim X As Integer
    For Each xlHoja2 In xlsLibro.Worksheets
       If xlHoja2.Name = lsNomHoja1 Then
            xlHoja2.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja2 = xlsLibro.Worksheets
        xlHoja2.Name = lsNomHoja1
    End If
    xlHoja2.Cells(5, 1) = lblFechaReporte.Caption

    PB1.value = 8
    Set rsCreditos = oCreditos.ReporteLineaCredito_Consolidado(gdFecSis, "04")
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If

    X = 8
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
                xlHoja2.Cells(X, 1) = rsCreditos!cAgeDescripcion
                xlHoja2.Cells(X, 1).Borders.LineStyle = 1
                xlHoja2.Cells(X, 2) = rsCreditos!Cliente
                xlHoja2.Cells(X, 2).Borders.LineStyle = 1
                xlHoja2.Cells(X, 3) = rsCreditos!cPersIDnro
                xlHoja2.Cells(X, 3).Borders.LineStyle = 1
                xlHoja2.Cells(X, 4) = rsCreditos!nEdad
                xlHoja2.Cells(X, 4).Borders.LineStyle = 1
                xlHoja2.Cells(X, 5) = rsCreditos!PerSexo
                xlHoja2.Cells(X, 5).Borders.LineStyle = 1
                xlHoja2.Cells(X, 6) = rsCreditos!Dpto
                xlHoja2.Cells(X, 6).Borders.LineStyle = 1
                xlHoja2.Cells(X, 7) = rsCreditos!Provincia
                xlHoja2.Cells(X, 7).Borders.LineStyle = 1
                xlHoja2.Cells(X, 8) = rsCreditos!Distrito
                xlHoja2.Cells(X, 8).Borders.LineStyle = 1
                xlHoja2.Cells(X, 9) = rsCreditos!SectorEconomico
                xlHoja2.Cells(X, 9).Borders.LineStyle = 1
                xlHoja2.Cells(X, 10) = rsCreditos!cCIIUdescripcion
                xlHoja2.Cells(X, 10).Borders.LineStyle = 1
                xlHoja2.Cells(X, 11) = rsCreditos!ProdRubro
                xlHoja2.Cells(X, 11).Borders.LineStyle = 1
                xlHoja2.Cells(X, 12) = rsCreditos!cTipoProd
                xlHoja2.Cells(X, 12).Borders.LineStyle = 1
                xlHoja2.Cells(X, 13) = rsCreditos!dFecVig
                xlHoja2.Cells(X, 13).Borders.LineStyle = 1
                xlHoja2.Cells(X, 14) = rsCreditos!cMoneda
                xlHoja2.Cells(X, 14).Borders.LineStyle = 1
                xlHoja2.Cells(X, 15) = rsCreditos!nMontoApr
                xlHoja2.Cells(X, 15).Borders.LineStyle = 1
                xlHoja2.Cells(X, 16) = rsCreditos!nSaldoCap
                xlHoja2.Cells(X, 16).Borders.LineStyle = 1
                xlHoja2.Cells(X, 17) = rsCreditos!nTasaInt
                xlHoja2.Cells(X, 17).Borders.LineStyle = 1
                xlHoja2.Cells(X, 18) = rsCreditos!nCuotasApr
                xlHoja2.Cells(X, 18).Borders.LineStyle = 1
                xlHoja2.Cells(X, 19) = rsCreditos!nGraciaApr
                xlHoja2.Cells(X, 19).Borders.LineStyle = 1
                xlHoja2.Cells(X, 20) = rsCreditos!cEstado
                xlHoja2.Cells(X, 20).Borders.LineStyle = 1
                xlHoja2.Cells(X, 21) = rsCreditos!cCalifActual
                xlHoja2.Cells(X, 21).Borders.LineStyle = 1

            X = X + 1
            rsCreditos.MoveNext
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    'END ****************************************************************
    
    'REPORTE FINANCIERO FONCODES ****************************************
    Dim dFechaFinanciera As Date
    Dim nMes As Integer
    Dim Y As Integer
    For Each xlHoja3 In xlsLibro.Worksheets
       If xlHoja3.Name = lsNomHoja2 Then
            xlHoja3.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja3 = xlsLibro.Worksheets
        xlHoja3.Name = lsNomHoja2
    End If
        
    PB1.value = 9
    nMes = Month(gdFecSis)
    dFechaFinanciera = gdFecSis
    If nMes > 1 Then
        For Y = 2 To nMes
            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                If Not rsCreditos.EOF Then
                       Dim nMesAnterior As Integer
                       nMesAnterior = Month(dFechaFinanciera) - 1
                            Select Case nMesAnterior
                                    Case 11:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 13) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 13) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 13) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 13) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 13) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 13) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 13) = rsCreditos!Monto
                                    Case 10:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 12) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 12) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 12) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 12) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 12) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 12) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 12) = rsCreditos!Monto
                                    Case 9:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 11) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 11) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 11) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 11) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 11) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 11) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 11) = rsCreditos!Monto
                                    Case 8:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 10) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 10) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 10) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 10) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 10) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 10) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 10) = rsCreditos!Monto
                                    Case 7:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 9) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 9) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 9) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 9) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 9) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 9) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 9) = rsCreditos!Monto
                                    Case 6:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 8) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 8) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 8) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 8) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 8) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 8) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 8) = rsCreditos!Monto
                                    Case 5:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 7) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 7) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 7) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 7) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 7) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 7) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 7) = rsCreditos!Monto
                                    Case 4:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 6) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 6) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 6) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 6) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 6) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 6) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 6) = rsCreditos!Monto
                                    Case 3:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 5) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 5) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 5) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 5) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 5) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 5) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 5) = rsCreditos!Monto
                                    Case 2:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 4) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 4) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 4) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 4) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 4) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 4) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 4) = rsCreditos!Monto
                                    Case 1:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 3) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 3) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 3) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 3) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 3) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 3) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 3) = rsCreditos!Monto
                            End Select
                End If
            dFechaFinanciera = DateAdd("M", -1, Format(dFechaFinanciera, "DD/MM/YYYY"))
            Set oCreditos = Nothing
            If nPase = 1 Then
                rsCreditos.Close
            End If
            Set rsCreditos = Nothing
        Next
    ElseIf nMes = 1 Then
        nMesAnterior = 12
        For Y = 1 To nMes + 11
                Select Case nMesAnterior
                        Case 12:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 14) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 14) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 14) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 14) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 14) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 14) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 11:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 13) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 13) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 13) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 13) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 13) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 13) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 10:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 12) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 12) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 12) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 12) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 12) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 12) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 9:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 11) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 11) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 11) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 11) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 11) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 11) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 8:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 10) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 10) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 10) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 10) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 10) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 10) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 7:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 9) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 9) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 9) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 9) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 9) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 9) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 6:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 8) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 8) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 8) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 8) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 8) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 8) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 5:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 7) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 7) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 7) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 7) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 7) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 7) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 4:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 6) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 6) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 6) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 6) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 6) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 6) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 3:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 5) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 5) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 5) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 5) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 5) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 5) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 2:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 4) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 4) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 4) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 4) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 4) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 4) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 1:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 3) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 3) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 3) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 3) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 3) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 3) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                End Select
            dFechaFinanciera = DateAdd("M", -1, Format(dFechaFinanciera, "DD/MM/YYYY"))
            nMesAnterior = nMesAnterior - 1
            Set oCreditos = Nothing
            If nPase = 1 Then
                rsCreditos.Close
            End If
            Set rsCreditos = Nothing
        Next
    End If
    'END ****************************************************************
    PB1.value = 10
    
    MsgBox "La comprobación se realizo en forma correcta ", vbInformation, "Aviso"
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
End Sub

Public Sub ReporteTrimestral(ByVal dfechSist As Date, ByVal nParTri As Integer)
    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja  As String
    Dim lsNomHoja1  As String
    Dim lsNomHoja2  As String
    Dim lsFecha As String
    Dim lsFechaIni As String
    Dim lsFechaFin As String
    Dim lsArchivo As String
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim xlHoja2 As Excel.Worksheet
    Dim xlHoja3 As Excel.Worksheet
    
    Dim rsCreditos As ADODB.Recordset
    Dim rsCreditosMes As ADODB.Recordset
    
    Dim oCreditos As New DCreditos
       
    Dim sdFecha As String
    Dim sFecha As String
    Dim sFechaIni As String
    Dim sFechaFin As String
    
    Dim pnLinPage As Integer
    Dim nPase As Integer
    
    Dim X As Integer
    Dim n As Integer
    Dim m As Integer
    Dim nColCant As Integer
    Dim nColSaldo As Integer
    Dim nColSaldoCant As Integer
    Dim nColSaldoMont As Integer
    Dim nColSMesCant As Integer
    Dim nColMesMont As Integer
    Dim nColMont As Integer
    
    n = 0
    m = 23
    PB1.Min = n
    PB1.Max = m
    PB1.value = n
    PB1.Visible = True
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ReporteCOFO"
    'Primera Hoja ******************************************************
    lsNomHoja = "ReporteCOFOTrimestral"
    '*******************************************************************
    'Segunda Hoja ******************************************************
    lsNomHoja1 = "BaseDatos"
    'Tercera Hoja ******************************************************
    lsNomHoja2 = "InfFinanciera"
    '*******************************************************************
    lsArchivo1 = "\spooler\LineaCredito" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & "FONCODES" & ".xls"
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
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If

    X = 1
    nColCant = 12
    nColSal = 14
    nColSaldoCant = 14
    nColSaldoMont = 15
    nColMesCant = 12
    nColMesMont = 13

    sdFecha = gdFecSis
    sFecha = DateAdd("D", 1, Format(lblFechaReporte.Caption, "DD/MM/YYYY"))
    sFechaIni = DateAdd("M", -1, Format(sFecha, "DD/MM/YYYY"))
    sFechaFin = Format(lblFechaReporte.Caption, "dd/MM/yyyy")
    n = n + 1
    PB1.value = n + 1
    Do While X <= nParTri
        'Creditos por Personeria
        n = n + 1
        PB1.value = n
        Set rsCreditos = oCreditos.ReporteLineaCredito_Personeria(sdFecha, "04")

        nPase = 1
        If (rsCreditos Is Nothing) Then
            nPase = 0
        End If

        If nPase = 1 Then
            Do While Not rsCreditos.EOF
                If rsCreditos!Concepto = "PerNaturalFemenino" Then
                    xlHoja1.Cells(15, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(15, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "PerNaturalMasculino" Then
                    xlHoja1.Cells(16, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(16, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "PerJuridica" Then
                    xlHoja1.Cells(17, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(17, nColSaldoMont) = rsCreditos!Monto
                End If
                rsCreditos.MoveNext
                If rsCreditos.EOF Then
                   Exit Do
                End If
            Loop
        End If
        Set oCreditos = Nothing
        If nPase = 1 Then
            rsCreditos.Close
        End If
        Set rsCreditos = Nothing

        n = n + 1
        PB1.value = n
        'Colocaciones
        Set rsCreditosMes = oCreditos.ReporteLineaCredito_PersoneriaMes(sdFecha, "04")

        nPase = 1
        If (rsCreditosMes Is Nothing) Then
            nPase = 0
        End If

        If nPase = 1 Then
            Do While Not rsCreditosMes.EOF
                If rsCreditosMes!Concepto = "PerNaturalFemenino" Then
                    xlHoja1.Cells(15, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(15, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "PerNaturalMasculino" Then
                    xlHoja1.Cells(16, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(16, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "PerJuridica" Then
                    xlHoja1.Cells(17, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(17, nColMesMont) = rsCreditosMes!Monto
                End If
                rsCreditosMes.MoveNext
                If rsCreditosMes.EOF Then
                   Exit Do
                End If
            Loop
        End If
        Set oCreditos = Nothing
        If nPase = 1 Then
            rsCreditosMes.Close
        End If
        Set rsCreditosMes = Nothing

        n = n + 1
        PB1.value = n
        'Creditos por Sector Economico
        Set rsCreditos = oCreditos.ReporteLineaCredito_SectorEconomico(sdFecha, "04")
        nPase = 1
        If (rsCreditos Is Nothing) Then
            nPase = 0
        End If

        If nPase = 1 Then
            Do While Not rsCreditos.EOF
                If rsCreditos!Concepto = "Agricola" Then
                    xlHoja1.Cells(19, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(19, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "Pecuario" Then
                    xlHoja1.Cells(20, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(20, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "Comercio" Then
                    xlHoja1.Cells(21, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(21, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "Produccion" Then
                    xlHoja1.Cells(22, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(22, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "Pesca" Then
                    xlHoja1.Cells(23, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(23, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "Servicios" Then
                    xlHoja1.Cells(24, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(24, nColSaldoMont) = rsCreditos!Monto
                End If
                rsCreditos.MoveNext
                If rsCreditos.EOF Then
                   Exit Do
                End If
            Loop
        End If
        Set oCreditos = Nothing
        If nPase = 1 Then
            rsCreditos.Close
        End If
        Set rsCreditos = Nothing

        n = n + 1
        PB1.value = n
        Set rsCreditosMes = oCreditos.ReporteLineaCredito_SectorEconomicoMes(sdFecha, "04")
        nPase = 1
        If (rsCreditosMes Is Nothing) Then
            nPase = 0
        End If

        If nPase = 1 Then
            Do While Not rsCreditosMes.EOF
                If rsCreditosMes!Concepto = "Agricola" Then
                    xlHoja1.Cells(19, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(19, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "Pecuario" Then
                    xlHoja1.Cells(20, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(20, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "Comercio" Then
                    xlHoja1.Cells(21, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(21, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "Produccion" Then
                    xlHoja1.Cells(22, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(22, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "Pesca" Then
                    xlHoja1.Cells(23, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(23, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "Servicios" Then
                    xlHoja1.Cells(24, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(24, nColMesMont) = rsCreditosMes!Monto
                End If
                rsCreditosMes.MoveNext
                If rsCreditosMes.EOF Then
                   Exit Do
                End If
            Loop
        End If
        Set oCreditos = Nothing
        If nPase = 1 Then
            rsCreditosMes.Close
        End If
        Set rsCreditosMes = Nothing

        'Creditos por Ubicacion Geografica
        n = n + 1
        PB1.value = n
        Set rsCreditos = oCreditos.ReporteLineaCredito_UbicaGeografica(sdFecha, "04")
        nPase = 1
        If (rsCreditos Is Nothing) Then
            nPase = 0
        End If

        If nPase = 1 Then
            Do While Not rsCreditos.EOF
                If rsCreditos!Concepto = "BAGUA" Then
                    xlHoja1.Cells(31, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(31, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "RODRIGUEZ DE MENDOZA" Then
                    xlHoja1.Cells(32, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(32, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "CHACHAPOYAS" Then
                    xlHoja1.Cells(33, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(33, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "CAJAMARCA" Then
                    xlHoja1.Cells(34, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(34, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "SAN MIGUEL" Then
                    xlHoja1.Cells(35, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(35, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "AMBO" Then
                    xlHoja1.Cells(36, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(36, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "DOS DE MAYO" Then
                    xlHoja1.Cells(37, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(37, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "HUAMALIES" Then
                    xlHoja1.Cells(38, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(38, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "HUANUCO" Then
                    xlHoja1.Cells(39, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(39, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "LEONCIO PRADO" Then
                    xlHoja1.Cells(40, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(40, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "PACHITEA" Then
                    xlHoja1.Cells(41, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(41, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "PUERTO INCA" Then
                    xlHoja1.Cells(42, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(42, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "MARAÑON" Then
                    xlHoja1.Cells(43, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(43, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "JUNIN" Then
                    xlHoja1.Cells(44, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(44, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "ALTO AMAZONAS" Then
                    xlHoja1.Cells(45, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(45, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "DATEM DEL MARAÑON" Then
                    xlHoja1.Cells(46, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(46, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "LORETO" Then
                    xlHoja1.Cells(47, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(47, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "MARISCAL RAMON CASTILLA" Then
                    xlHoja1.Cells(48, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(48, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "MAYNAS" Then
                    xlHoja1.Cells(49, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(49, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "REQUENA" Then
                    xlHoja1.Cells(50, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(50, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "UCAYALI" Then
                    xlHoja1.Cells(51, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(51, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "DANIEL ALCIDES CARRION" Then
                    xlHoja1.Cells(52, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(52, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "PASCO" Then
                    xlHoja1.Cells(53, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(53, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "EL DORADO" Then
                    xlHoja1.Cells(54, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(54, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "LAMAS" Then
                    xlHoja1.Cells(55, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(55, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "MARISCAL CACERES" Then
                    xlHoja1.Cells(56, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(56, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "RIOJA" Then
                    xlHoja1.Cells(57, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(57, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "SAN MARTIN" Then
                    xlHoja1.Cells(58, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(58, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "BELLAVISTA" Then
                    xlHoja1.Cells(59, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(59, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "HUALLAGA" Then
                    xlHoja1.Cells(60, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(60, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "MOYOBAMBA" Then
                    xlHoja1.Cells(61, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(61, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "PICOTA" Then
                    xlHoja1.Cells(62, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(62, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "TOCACHE" Then
                    xlHoja1.Cells(63, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(63, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "ATALAYA" Then
                    xlHoja1.Cells(64, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(64, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "CORONEL PORTILLO" Then
                    xlHoja1.Cells(65, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(65, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "PADRE ABAD" Then
                    xlHoja1.Cells(66, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(66, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "Callao" Then
                    xlHoja1.Cells(67, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(67, nColSaldoMont) = rsCreditos!Monto
                End If
                If rsCreditos!Concepto = "Lima" Then
                    xlHoja1.Cells(68, nColSaldoCant) = rsCreditos!Cantidad
                    xlHoja1.Cells(68, nColSaldoMont) = rsCreditos!Monto
                End If
                rsCreditos.MoveNext
                If rsCreditos.EOF Then
                   Exit Do
                End If
            Loop
        End If
        Set oCreditos = Nothing
        If nPase = 1 Then
            rsCreditos.Close
        End If
        Set rsCreditos = Nothing

        n = n + 1
        PB1.value = n
        Set rsCreditosMes = oCreditos.ReporteLineaCredito_UbicaGeograficaMes(sdFecha, "04")
        nPase = 1
        If (rsCreditosMes Is Nothing) Then
            nPase = 0
        End If

        If nPase = 1 Then
            Do While Not rsCreditosMes.EOF
                If rsCreditosMes!Concepto = "BAGUA" Then
                    xlHoja1.Cells(31, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(31, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "RODRIGUEZ DE MENDOZA" Then
                    xlHoja1.Cells(32, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(32, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "CHACHAPOYAS" Then
                    xlHoja1.Cells(33, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(33, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "CAJAMARCA" Then
                    xlHoja1.Cells(34, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(34, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "SAN MIGUEL" Then
                    xlHoja1.Cells(35, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(35, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "AMBO" Then
                    xlHoja1.Cells(36, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(36, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "DOS DE MAYO" Then
                    xlHoja1.Cells(37, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(37, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "HUAMALIES" Then
                    xlHoja1.Cells(38, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(38, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "HUANUCO" Then
                    xlHoja1.Cells(39, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(39, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "LEONCIO PRADO" Then
                    xlHoja1.Cells(40, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(40, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "PACHITEA" Then
                    xlHoja1.Cells(41, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(41, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "PUERTO INCA" Then
                    xlHoja1.Cells(42, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(42, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "MARAÑON" Then
                    xlHoja1.Cells(43, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(43, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "JUNIN" Then
                    xlHoja1.Cells(44, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(44, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "ALTO AMAZONAS" Then
                    xlHoja1.Cells(45, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(45, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "DATEM DEL MARAÑON" Then
                    xlHoja1.Cells(46, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(46, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "LORETO" Then
                    xlHoja1.Cells(47, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(47, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "MARISCAL RAMON CASTILLA" Then
                    xlHoja1.Cells(48, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(48, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "MAYNAS" Then
                    xlHoja1.Cells(49, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(49, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "REQUENA" Then
                    xlHoja1.Cells(50, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(50, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "UCAYALI" Then
                    xlHoja1.Cells(51, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(51, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "DANIEL ALCIDES CARRION" Then
                    xlHoja1.Cells(52, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(52, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "PASCO" Then
                    xlHoja1.Cells(53, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(53, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "EL DORADO" Then
                    xlHoja1.Cells(54, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(54, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "LAMAS" Then
                    xlHoja1.Cells(55, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(55, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "MARISCAL CACERES" Then
                    xlHoja1.Cells(56, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(56, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "RIOJA" Then
                    xlHoja1.Cells(57, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(57, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "SAN MARTIN" Then
                    xlHoja1.Cells(58, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(58, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "BELLAVISTA" Then
                    xlHoja1.Cells(59, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(59, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "HUALLAGA" Then
                    xlHoja1.Cells(60, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(60, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "MOYOBAMBA" Then
                    xlHoja1.Cells(61, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(61, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "PICOTA" Then
                    xlHoja1.Cells(62, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(62, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "TOCACHE" Then
                    xlHoja1.Cells(63, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(63, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "ATALAYA" Then
                    xlHoja1.Cells(64, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(64, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "CORONEL PORTILLO" Then
                    xlHoja1.Cells(65, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(65, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "PADRE ABAD" Then
                    xlHoja1.Cells(66, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(66, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "Callao" Then
                    xlHoja1.Cells(67, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(67, nColMesMont) = rsCreditosMes!Monto
                End If
                If rsCreditosMes!Concepto = "Lima" Then
                    xlHoja1.Cells(68, nColMesCant) = rsCreditosMes!Cantidad
                    xlHoja1.Cells(68, nColMesMont) = rsCreditosMes!Monto
                End If
                rsCreditosMes.MoveNext
                If rsCreditosMes.EOF Then
                   Exit Do
                End If
            Loop
        End If
        Set oCreditos = Nothing
        If nPase = 1 Then
            rsCreditosMes.Close
        End If
        Set rsCreditosMes = Nothing
    X = X + 1
    nColSaldoCant = nColSaldoCant - 4
    nColSaldoMont = nColSaldoMont - 4
    nColMesCant = nColMesCant - 4
    nColMesMont = nColMesMont - 4

    xlHoja1.Cells(10, nColCant) = sFechaIni & " Al " & sFechaFin
    xlHoja1.Cells(10, nColSal) = sFechaFin

    sdFecha = DateAdd("M", -1, Format(sdFecha, "DD/MM/YYYY"))
    sFechaIni = DateAdd("M", -1, Format(sFechaIni, "DD/MM/YYYY"))
    sFechaFin = obtenerFechaFinMes(Month(sFechaIni), Year(sFechaIni))

    nColCant = nColCant - 4
    nColSal = nColSal - 4
Loop

    'REPORTE BASE DE DATOS **************************************************
    Dim Y As Integer
    For Each xlHoja2 In xlsLibro.Worksheets
       If xlHoja2.Name = lsNomHoja1 Then
            xlHoja2.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja2 = xlsLibro.Worksheets
        xlHoja2.Name = lsNomHoja1
    End If
    xlHoja2.Cells(5, 1) = lblFechaReporte.Caption

    n = n + 1
    PB1.value = n
    Set rsCreditos = oCreditos.ReporteLineaCredito_Consolidado(gdFecSis, "04")
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If

    Y = 8
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
                xlHoja2.Cells(Y, 1) = rsCreditos!cAgeDescripcion
                xlHoja2.Cells(Y, 1).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 2) = rsCreditos!Cliente
                xlHoja2.Cells(Y, 2).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 3) = rsCreditos!cPersIDnro
                xlHoja2.Cells(Y, 3).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 4) = rsCreditos!nEdad
                xlHoja2.Cells(Y, 4).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 5) = rsCreditos!PerSexo
                xlHoja2.Cells(Y, 5).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 6) = rsCreditos!Dpto
                xlHoja2.Cells(Y, 6).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 7) = rsCreditos!Provincia
                xlHoja2.Cells(Y, 7).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 8) = rsCreditos!Distrito
                xlHoja2.Cells(Y, 8).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 9) = rsCreditos!SectorEconomico
                xlHoja2.Cells(Y, 9).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 10) = rsCreditos!cCIIUdescripcion
                xlHoja2.Cells(Y, 10).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 11) = rsCreditos!ProdRubro
                xlHoja2.Cells(Y, 11).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 12) = rsCreditos!cTipoProd
                xlHoja2.Cells(Y, 12).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 13) = rsCreditos!dFecVig
                xlHoja2.Cells(Y, 13).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 14) = rsCreditos!cMoneda
                xlHoja2.Cells(Y, 14).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 15) = rsCreditos!nMontoApr
                xlHoja2.Cells(Y, 15).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 16) = rsCreditos!nSaldoCap
                xlHoja2.Cells(Y, 16).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 17) = rsCreditos!nTasaInt
                xlHoja2.Cells(Y, 17).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 18) = rsCreditos!nCuotasApr
                xlHoja2.Cells(Y, 18).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 19) = rsCreditos!nGraciaApr
                xlHoja2.Cells(Y, 19).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 20) = rsCreditos!cEstado
                xlHoja2.Cells(Y, 20).Borders.LineStyle = 1
                xlHoja2.Cells(Y, 21) = rsCreditos!cCalifActual
                xlHoja2.Cells(Y, 21).Borders.LineStyle = 1

            Y = Y + 1
            rsCreditos.MoveNext
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    'End *****************************************************************
    n = n + 1
    PB1.value = n
    'REPORTE FINANCIERO FONCODES ****************************************
    Dim dFechaFinanciera As Date
    Dim nMes As Integer
    Dim Z As Integer
    For Each xlHoja3 In xlsLibro.Worksheets
       If xlHoja3.Name = lsNomHoja2 Then
            xlHoja3.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja3 = xlsLibro.Worksheets
        xlHoja3.Name = lsNomHoja2
    End If
    
    n = n + 1
    PB1.value = n
    nMes = Month(gdFecSis)
    dFechaFinanciera = gdFecSis
    If nMes > 1 Then
        For Z = 2 To nMes
            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                If Not rsCreditos.EOF Then
                       Dim nMesAnterior As Integer
                       nMesAnterior = Month(dFechaFinanciera) - 1
                            Select Case nMesAnterior
                                    Case 11:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 13) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 13) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 13) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 13) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 13) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 13) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
                                        'xlHoja3.Cells(15, 13) = rsCreditos!Monto
                                    Case 10:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 12) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 12) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 12) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 12) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 12) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 12) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 12) = rsCreditos!Monto
                                    Case 9:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 11) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 11) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 11) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 11) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 11) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 11) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 11) = rsCreditos!Monto
                                    Case 8:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 10) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 10) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 10) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 10) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 10) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 10) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 10) = rsCreditos!Monto
                                    Case 7:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 9) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 9) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 9) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 9) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 9) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 9) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 9) = rsCreditos!Monto
                                    Case 6:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 8) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 8) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 8) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 8) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 8) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 8) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 8) = rsCreditos!Monto
                                    Case 5:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 7) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 7) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 7) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 7) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 7) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 7) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 7) = rsCreditos!Monto
                                    Case 4:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 6) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 6) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 6) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 6) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 6) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 6) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 6) = rsCreditos!Monto
                                    Case 3:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 5) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 5) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 5) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 5) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 5) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 5) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 5) = rsCreditos!Monto
                                    Case 2:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 4) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 4) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 4) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 4) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 4) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 4) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 4) = rsCreditos!Monto
                                    Case 1:
                                        If Not rsCreditos.EOF Then
                                            Do While Not rsCreditos.EOF
                                                If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                                    xlHoja3.Cells(15, 3) = rsCreditos!Monto
                                                    xlHoja3.Cells(23, 3) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "1 CPP" Then
                                                    xlHoja3.Cells(24, 3) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                                    xlHoja3.Cells(25, 3) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                                    xlHoja3.Cells(26, 3) = rsCreditos!Monto
                                                End If
                                                If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                                    xlHoja3.Cells(27, 3) = rsCreditos!Monto
                                                End If
                                                rsCreditos.MoveNext
                                                If rsCreditos.EOF Then
                                                   Exit Do
                                                End If
                                            Loop
                                        End If
'                                        xlHoja3.Cells(15, 3) = rsCreditos!Monto
                            End Select
                End If
            dFechaFinanciera = DateAdd("M", -1, Format(dFechaFinanciera, "DD/MM/YYYY"))
            Set oCreditos = Nothing
            If nPase = 1 Then
                rsCreditos.Close
            End If
            Set rsCreditos = Nothing
        Next
    ElseIf nMes = 1 Then
        nMesAnterior = 12
        For Y = 1 To nMes + 11
                Select Case nMesAnterior
                        Case 12:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 14) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 14) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 14) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 14) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 14) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 14) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 11:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 13) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 13) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 13) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 13) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 13) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 13) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 10:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 12) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 12) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 12) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 12) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 12) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 12) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 9:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 11) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 11) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 11) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 11) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 11) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 11) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 8:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 10) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 10) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 10) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 10) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 10) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 10) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 7:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 9) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 9) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 9) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 9) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 9) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 9) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 6:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 8) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 8) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 8) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 8) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 8) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 8) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 5:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 7) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 7) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 7) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 7) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 7) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 7) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 4:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 6) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 6) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 6) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 6) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 6) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 6) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 3:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 5) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 5) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 5) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 5) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 5) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 5) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 2:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 4) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 4) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 4) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 4) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 4) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 4) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                        Case 1:
                            Set rsCreditos = oCreditos.ReporteFinancieroFoncodes(dFechaFinanciera, "04")
                            If Not rsCreditos.EOF Then
                                Do While Not rsCreditos.EOF
                                    If rsCreditos!cCalifAnterior = "0 NORMAL" Then
                                        xlHoja3.Cells(15, 3) = rsCreditos!Monto
                                        xlHoja3.Cells(23, 3) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "1 CPP" Then
                                        xlHoja3.Cells(24, 3) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "2 DEFICIENTE" Then
                                        xlHoja3.Cells(25, 3) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "3 DUDOSO" Then
                                        xlHoja3.Cells(26, 3) = rsCreditos!Monto
                                    End If
                                    If rsCreditos!cCalifAnterior = "4 PERDIDA" Then
                                        xlHoja3.Cells(27, 3) = rsCreditos!Monto
                                    End If
                                    rsCreditos.MoveNext
                                    If rsCreditos.EOF Then
                                       Exit Do
                                    End If
                                Loop
                            End If
                End Select
            dFechaFinanciera = DateAdd("M", -1, Format(dFechaFinanciera, "DD/MM/YYYY"))
            nMesAnterior = nMesAnterior - 1
            Set oCreditos = Nothing
            If nPase = 1 Then
                rsCreditos.Close
            End If
            Set rsCreditos = Nothing
        Next
    End If
    'END ****************************************************************
n = n + 1
PB1.value = n

m = n
PB1.value = m
MsgBox "La comprobación se realizo en forma correcta ", vbInformation, "Aviso"
xlHoja1.SaveAs App.path & lsArchivo1

xlsAplicacion.Visible = True
xlsAplicacion.Windows(1).Visible = True
Set xlsAplicacion = Nothing
Set xlsLibro = Nothing
Set xlHoja1 = Nothing
Set xlHoja2 = Nothing
Set xlHoja3 = Nothing

Exit Sub
End Sub

Private Function obtenerFechaFinMes(ByVal pnMes As Integer, ByVal pnAnio As Integer) As Date
    Dim sFecha  As Date
    sFecha = CDate("01/" & Format(pnMes, "00") & "/" & pnAnio)
    sFecha = DateAdd("m", 1, sFecha)
    sFecha = sFecha - 1
    obtenerFechaFinMes = sFecha
End Function

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim loConstS As DConstSistemas
    CentraForm Me
    Set loConstS = New DConstSistemas
    lblFechaReporte.Caption = CDate(loConstS.LeeConstSistema(gConstSistCierreMesNegocio))
End Sub


