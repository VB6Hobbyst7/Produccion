VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmParaEncDiario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros para encaje diario"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   Icon            =   "frmParaEncDiario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFecha 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   8775
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
   End
   Begin VB.Frame fraParaEncDiaro 
      Caption         =   "Parámetros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8775
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   375
         Left            =   5280
         TabIndex        =   3
         Top             =   4200
         Width           =   1215
      End
      Begin Sicmact.FlexEdit FEParaEncDiario 
         Height          =   3855
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8535
         _extentx        =   15055
         _extenty        =   6800
         cols0           =   6
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "Nro-Parámetro-Tipo-Valor-cCodParamEncDiario-nEstado"
         encabezadosanchos=   "500-4070-2000-1500-0-0"
         font            =   "frmParaEncDiario.frx":030A
         font            =   "frmParaEncDiario.frx":0336
         font            =   "frmParaEncDiario.frx":0362
         font            =   "frmParaEncDiario.frx":038E
         font            =   "frmParaEncDiario.frx":03BA
         fontfixed       =   "frmParaEncDiario.frx":03E6
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-X-3-X-X"
         listacontroles  =   "0-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-R-C-C"
         formatosedit    =   "0-0-0-0-0-0"
         avanceceldas    =   1
         textarray0      =   "Nro"
         colwidth0       =   495
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   6840
         TabIndex        =   1
         Top             =   4200
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   4200
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblAvance 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   4560
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmParaEncDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmParaEncDiario
'*** Descripción : Formulario para el encaje legal diario.
'*** Creación : MIOL el 20120827, según OYP-RFC091-2012
'********************************************************************
Option Explicit
Dim cuSaldoEncDiaUltMN As Currency
Dim cuSaldoEncDiaUltME As Currency
Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    Dim oNEncajeBCR As nEncajeBCR
    Set oNEncajeBCR = New nEncajeBCR
    Dim rsEncDiario As ADODB.Recordset
    Set rsEncDiario = New ADODB.Recordset
    Dim f As Integer
    f = 0
    Set rsEncDiario = oNEncajeBCR.ObtenerParamEncajeDiario()
    
    Do While f < rsEncDiario.RecordCount
        f = f + 1
        If FEParaEncDiario.TextMatrix(f, 4) = "02" Or FEParaEncDiario.TextMatrix(f, 4) = "03" Or FEParaEncDiario.TextMatrix(f, 4) = "12" Or FEParaEncDiario.TextMatrix(f, 4) = "13" Then
            Call oNEncajeBCR.UpdateParamEncajeDiario(Format(FEParaEncDiario.TextMatrix(f, 3), "#,##0.0000"), FEParaEncDiario.TextMatrix(f, 4))
        ElseIf FEParaEncDiario.TextMatrix(f, 4) = "10" Then
            Call oNEncajeBCR.UpdateParamEncajeDiario(FEParaEncDiario.TextMatrix(f, 3), FEParaEncDiario.TextMatrix(f, 4))
        Else
            Call oNEncajeBCR.UpdateParamEncajeDiario(Format(FEParaEncDiario.TextMatrix(f, 3), "#,##0.00"), FEParaEncDiario.TextMatrix(f, 4))
        End If
        gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        Call oNEncajeBCR.InsParamEncajeDiario(FEParaEncDiario.TextMatrix(f, 4), gsMovNro, Format(FEParaEncDiario.TextMatrix(f, 3), "#,##0.00"), gdFecSis)
    Loop
    MsgBox "Los Parametros de Encaje Diario se Actualizaron correctamente", vbInformation, "Aviso"
    Me.cmdGuardar.Enabled = False
    Set rsEncDiario = Nothing
    'Procedimiento para generar el reporte
    generarAnexo
    Unload Me
End Sub

Private Sub generarAnexo()
        Me.MousePointer = vbHourglass
        Dim sPathAnexoDiario As String
       
        Dim fs As New Scripting.FileSystemObject
        Dim obj_excel As Object, Libro As Object, Hoja As Object
        
        On Error GoTo error_sub
          
        PB1.Min = 0
        PB1.Max = 16
        PB1.value = 0
        PB1.Visible = True
        sPathAnexoDiario = App.path & "\Spooler\ANEXO_DIARIO_" + Format(Me.txtFecha.Text, "yyyymmdd") + ".xls"
        
        If fs.FileExists(sPathAnexoDiario) Then
            
            If ArchivoEstaAbierto(sPathAnexoDiario) Then
                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(sPathAnexoDiario) + " para continuar", vbRetryCancel) = vbCancel Then
                   Me.MousePointer = vbDefault
                   Exit Sub
                End If
                Me.MousePointer = vbHourglass
            End If
    
            fs.DeleteFile sPathAnexoDiario, True
        End If
        PB1.value = 1
        lblAvance.Caption = "Abriendo archivo a copiar"
        'sPathAnexoDiario = App.path & "\FormatoCarta\PLANILLAANEXODIARIO.xls" 'Comentado por pasi 20140409
        sPathAnexoDiario = App.path & "\FormatoCarta\Nuevo_PLANILLAANEXODIARIO.xls" '*** PASI 20140409
        lblAvance.Caption = "cargando archivo..."
        If Len(Dir(sPathAnexoDiario)) = 0 Then
           MsgBox "No se Pudo Encontrar el Archivo:" & sPathAnexoDiario, vbCritical
           Me.MousePointer = vbDefault
           lblAvance.Caption = ""
            PB1.Visible = False
           Exit Sub
        End If
        
        Set obj_excel = CreateObject("Excel.Application")
        obj_excel.DisplayAlerts = False
        Set Libro = obj_excel.Workbooks.Open(sPathAnexoDiario)
        Set Hoja = Libro.ActiveSheet
        
        Dim celda As Excel.Range
        Dim oCtaCont As DbalanceCont
        Dim rsCtaCont As ADODB.Recordset
        
        Set oCtaCont = New DbalanceCont
        Set rsCtaCont = New ADODB.Recordset
       ' Fecha del ANEXO
        FechaAnioMesDia obj_excel
        
        PB1.value = 2
        lblAvance.Caption = "Cargando Datos..."
        '*************************OBLIGACIONES INMEDIATAS*****************************
        cargarObligacionesInmediatas obj_excel

        '***************SALDOS X ENCAJE: AHORROS, PLAZO FIJO Y CTS********************
        PB1.value = 3
        cargarAhorrosPlazoFijoCTS obj_excel
        
        PB1.value = 4
        cargarChequeAhorrosPlazoFijoCTS obj_excel

        '******************************* OTRAS CTAS **********************************
        PB1.value = 5
        cargarBCRP obj_excel 'BCRP
        
        PB1.value = 6
        cargarCmacsCracs obj_excel 'BCRP
        
        PB1.value = 7
        cargarCajaObligExoneradas obj_excel 'CAJA
        
        PB1.value = 8
        cargarIntDevengado obj_excel 'INT DEVENGADO
        
        '******************************* PF X RANGOS *********************************
        PB1.value = 9
        cargarPlazoFijoxRango obj_excel, 1 'PFxRANGOS MN
        
        PB1.value = 10
        cargarPlazoFijoxRango obj_excel, 2 'PFxRANGOS ME
        
        '******************************* ANEXO DIARIO  *******************************
        PB1.value = 11
        cargarAnexoDiario obj_excel

        PB1.value = 12 '*** Agregado por PASI 20140409
        CargarProyecciones obj_excel
        
        'PB1.value = 12 '*** Modificado por Pasi 20140409
        PB1.value = 13
        
        'Verifica si Existe el Archivo
        lblAvance.Caption = "Verificando Archivo..."
        sPathAnexoDiario = App.path & "\Spooler\ANEXO_DIARIO_" + Format(Me.txtFecha.Text, "yyyymmdd") + ".xls"
        If fs.FileExists(sPathAnexoDiario) Then
            
            If ArchivoEstaAbierto(sPathAnexoDiario) Then
                MsgBox "Debe Cerrar el Archivo:" + fs.GetFileName(sPathAnexoDiario)
            End If
            fs.DeleteFile sPathAnexoDiario, True
        End If
        PB1.value = 13
        'Guarda el Archivo
        lblAvance.Caption = "Guardando Archivo..."
        Hoja.SaveAs sPathAnexoDiario
        Libro.Close
        obj_excel.Quit
        PB1.value = 14
        Set Hoja = Nothing
        Set Libro = Nothing
        Set obj_excel = Nothing
        Me.MousePointer = vbDefault
        PB1.value = 15
        'Abre y Muestra el Archivo
        lblAvance.Caption = "Abriendo Archivo..."
        Dim m_excel As New Excel.Application
        m_excel.Workbooks.Open (sPathAnexoDiario)
        m_excel.Visible = True
        PB1.value = 16
        PB1.Visible = False
        lblAvance.Caption = ""
Exit Sub
'Unload Me
error_sub:
        MsgBox TextErr(Err.Description), vbInformation, "Aviso"
        Set Libro = Nothing
        Set obj_excel = Nothing
        Set Hoja = Nothing
        PB1.Visible = False
        lblAvance.Caption = ""
        Me.MousePointer = vbDefault
End Sub

Private Sub cargarObligacionesInmediatas(ByVal pobj_Excel As Excel.Application)
Dim nFil As Integer
Dim nFilCaj As Integer
Dim TotalMN As Currency
Dim TotalME As Currency
Dim lnTCF As Currency
Dim dFecIni As Date
Dim TotalDias As Integer
Dim celda As Excel.Range
dFecIni = CDate("01/" & Format(Month(Me.txtFecha.Text), "00") & "/" & Format(Year(Me.txtFecha.Text), "0000"))
TotalDias = DateDiff("d", dFecIni, Me.txtFecha.Text) + 1
nFil = 12
nFilCaj = 8
Dim oTC As New nTipoCambio
lnTCF = Format(oTC.EmiteTipoCambio(dFecIni, TCFijoMes), "#,##0.00###")

Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsEncDiario As ADODB.Recordset
Set rsEncDiario = New ADODB.Recordset
Dim nEfectivoCaja As Currency
Set rsEncDiario = oNEncajeBCR.ObtenerParamEncajeDiarioxCod("05")
nEfectivoCaja = rsEncDiario!nValor

Do While dFecIni <= Me.txtFecha.Text
  'OBLIGACIONES INMEDIATAS
   'MONEDA NACIONAL
        TotalMN = SaldoCtas(1, "761201", dFecIni, Me.txtFecha.Text, lnTCF, lnTCF)
        Set celda = pobj_Excel.Range("ENCAJELEGALMN!F" & nFil)
        celda.value = TotalMN
   'MONEDA EXTRAJERA
        TotalME = SaldoCtas(1, "762201", dFecIni, Me.txtFecha.Text, lnTCF, lnTCF)
        Set celda = pobj_Excel.Range("ENCAJELEGALME!F" & nFil)
        celda.value = TotalME
   'EFECTIVO CAJA
        'Set celda = pobj_Excel.Range("ENCAJELEGALMN!L" & nFil) 'Comentado por PASI
        Set celda = pobj_Excel.Range("ENCAJELEGALMN!M" & nFil) '*** PASI20140409
        celda.value = nEfectivoCaja
    
    nFil = nFil + 1
    nFilCaj = nFilCaj + 1
    dFecIni = DateAdd("d", 1, dFecIni)
Loop
Set rsEncDiario = Nothing
End Sub

Private Sub cargarAhorrosPlazoFijoCTS(ByVal pobj_Excel As Excel.Application)
Dim nFil As Integer
Dim TotalMN As Currency
Dim TotalME As Currency
Dim TotalAhoExo As Currency
Dim lnTCF As Currency
Dim dFecIni As Date
Dim TotalDias As Integer
Dim celda As Excel.Range
dFecIni = CDate("01/" & Format(Month(Me.txtFecha.Text), "00") & "/" & Format(Year(Me.txtFecha.Text), "0000"))
TotalDias = DateDiff("d", dFecIni, Me.txtFecha.Text) + 1
nFil = 8
Dim oTC As New nTipoCambio
lnTCF = Format(oTC.EmiteTipoCambio(dFecIni, TCFijoMes), "#,##0.00###")
Do While dFecIni <= Me.txtFecha.Text
    'AHORROS *******************************************************
        'MONEDA NACIONAL
         TotalMN = SaldoAhoPlaFijCTS(Format(dFecIni, "yyyymmdd"), 1, "232")
         Set celda = pobj_Excel.Range("SALDOSxENCAJE!B" & nFil)
         celda.value = TotalMN
        'MONEDA EXTRAJERA
         TotalME = SaldoAhoPlaFijCTS(Format(dFecIni, "yyyymmdd"), 2, "232")
         Set celda = pobj_Excel.Range("SALDOSxENCAJE!H" & nFil)
         celda.value = TotalME
     'PLAZO FIJO ***************************************************
         'MONEDA NACIONAL
         TotalMN = SaldoAhoPlaFijCTS(Format(dFecIni, "yyyymmdd"), 1, "233")
         Set celda = pobj_Excel.Range("SALDOSxENCAJE!D" & nFil)
         celda.value = TotalMN
        'MONEDA EXTRAJERA
         TotalME = SaldoAhoPlaFijCTS(Format(dFecIni, "yyyymmdd"), 2, "233")
         TotalAhoExo = SaldoAhoExoEnc()
         Set celda = pobj_Excel.Range("SALDOSxENCAJE!J" & nFil)
         celda.value = TotalME
        'CARGAR DATOS PARA CUADRE
         If dFecIni = Me.txtFecha.Text Then
            cuSaldoEncDiaUltMN = TotalMN
            Set celda = pobj_Excel.Range("PFxRangos!C42")
            celda.value = TotalMN
            Set celda = pobj_Excel.Range("PFxRangos!C43")
            celda.value = TotalME
         End If
     'CTS **********************************************************
         'MONEDA NACIONAL
         TotalMN = SaldoAhoPlaFijCTS(Format(dFecIni, "yyyymmdd"), 1, "234")
         Set celda = pobj_Excel.Range("SALDOSxENCAJE!F" & nFil)
         celda.value = TotalMN
        'MONEDA EXTRAJERA
         TotalME = SaldoAhoPlaFijCTS(Format(dFecIni, "yyyymmdd"), 2, "234")
         Set celda = pobj_Excel.Range("SALDOSxENCAJE!L" & nFil)
         celda.value = TotalME
         
     nFil = nFil + 1
     dFecIni = DateAdd("d", 1, dFecIni)
Loop
End Sub

Private Sub cargarChequeAhorrosPlazoFijoCTS(ByVal pobj_Excel As Excel.Application)
Dim nFil As Integer
Dim TotalMN As Currency
Dim TotalME As Currency
Dim TotalAhoExo As Currency
Dim lnTCF As Currency
Dim dFecIni As Date
Dim TotalDias As Integer
Dim celda As Excel.Range
dFecIni = CDate("01/" & Format(Month(Me.txtFecha.Text), "00") & "/" & Format(Year(Me.txtFecha.Text), "0000"))
TotalDias = DateDiff("d", dFecIni, Me.txtFecha.Text) + 1
nFil = 8
Dim oTC As New nTipoCambio
lnTCF = Format(oTC.EmiteTipoCambio(dFecIni, TCFijoMes), "#,##0.00###")
Do While dFecIni <= Me.txtFecha.Text
    'AHORROS *******************************************************
        'MONEDA NACIONAL
         TotalMN = SaldoChequeAhoPlaFijCTS(Format(dFecIni, "yyyymmdd"), 1, "232")
         Set celda = pobj_Excel.Range("SALDOSxENCAJE!C" & nFil)
         celda.value = TotalMN
        'MONEDA EXTRAJERA
         TotalME = SaldoChequeAhoPlaFijCTS(Format(dFecIni, "yyyymmdd"), 2, "232")
         Set celda = pobj_Excel.Range("SALDOSxENCAJE!I" & nFil)
         celda.value = TotalME
     'PLAZO FIJO ***************************************************
         'MONEDA NACIONAL
         TotalMN = SaldoChequeAhoPlaFijCTS(Format(dFecIni, "yyyymmdd"), 1, "233")
         Set celda = pobj_Excel.Range("SALDOSxENCAJE!E" & nFil)
         celda.value = TotalMN
        'MONEDA EXTRAJERA
         TotalME = SaldoChequeAhoPlaFijCTS(Format(dFecIni, "yyyymmdd"), 2, "233")
         Set celda = pobj_Excel.Range("SALDOSxENCAJE!K" & nFil)
         celda.value = TotalME
     'CTS **********************************************************
         'MONEDA NACIONAL
         TotalMN = SaldoChequeAhoPlaFijCTS(Format(dFecIni, "yyyymmdd"), 1, "234")
         Set celda = pobj_Excel.Range("SALDOSxENCAJE!G" & nFil)
         celda.value = TotalMN
        'MONEDA EXTRAJERA
         TotalME = SaldoChequeAhoPlaFijCTS(Format(dFecIni, "yyyymmdd"), 2, "234")
         Set celda = pobj_Excel.Range("SALDOSxENCAJE!M" & nFil)
         celda.value = TotalME
         
     nFil = nFil + 1
     dFecIni = DateAdd("d", 1, dFecIni)
Loop
End Sub

Private Sub cargarBCRP(ByVal pobj_Excel As Excel.Application)
Dim nFil As Integer
Dim TotalMN As Currency
Dim TotalME As Currency
Dim lnTCF As Currency
Dim dFecIni As Date
Dim TotalDias As Integer
Dim celda As Excel.Range
dFecIni = CDate("01/" & Format(Month(Me.txtFecha.Text), "00") & "/" & Format(Year(Me.txtFecha.Text), "0000"))
TotalDias = DateDiff("d", dFecIni, Me.txtFecha.Text) + 1
nFil = 8
Dim oTC As New nTipoCambio
lnTCF = Format(oTC.EmiteTipoCambio(dFecIni, TCFijoMes), "#,##0.00###")
Do While dFecIni <= Me.txtFecha.Text
    'BCRP **************************************************************************
        'MONEDA NACIONAL
         TotalMN = SaldoBCRPAnexoDiario(Format(dFecIni, "yyyymmdd"), 1)
         Set celda = pobj_Excel.Range("OtrasCTAS!B" & nFil)
         celda.value = TotalMN
        'MONEDA EXTRAJERA
         TotalME = SaldoBCRPAnexoDiario(Format(dFecIni, "yyyymmdd"), 2)
         'Set celda = pobj_Excel.Range("OtrasCTAS!Q" & nFil) 'Comentado Por PASI20140409
         'Set celda = pobj_Excel.Range("OtrasCTAS!R" & nfil) '*** PASI 20140409
         Set celda = pobj_Excel.Range("OtrasCTAS!T" & nFil) 'NAGL 20190713 Según RFC1907080005
         celda.value = TotalME
     nFil = nFil + 1
     dFecIni = DateAdd("d", 1, dFecIni)
Loop
End Sub

Private Sub cargarCmacsCracs(ByVal pobj_Excel As Excel.Application)
Dim nFil As Integer
Dim TotalMN As Currency
Dim TotalME As Currency
Dim lnTCF As Currency
Dim dFecIni As Date
Dim TotalDias As Integer
Dim celda As Excel.Range
dFecIni = CDate("01/" & Format(Month(Me.txtFecha.Text), "00") & "/" & Format(Year(Me.txtFecha.Text), "0000"))
TotalDias = DateDiff("d", dFecIni, Me.txtFecha.Text) + 1
nFil = 8
Dim oTC As New nTipoCambio
lnTCF = Format(oTC.EmiteTipoCambio(dFecIni, TCFijoMes), "#,##0.00###")
Do While dFecIni <= Me.txtFecha.Text
    'CAJAS AHORROS **************************************************
        'MONEDA NACIONAL
         TotalMN = SaldoCajasCracsAnexoDiario(Format(dFecIni, "yyyymmdd"), 1, "232")
         Set celda = pobj_Excel.Range("OtrasCTAS!D" & nFil)
         celda.value = TotalMN
        'MONEDA EXTRAJERA
         TotalME = SaldoCajasCracsAnexoDiario(Format(dFecIni, "yyyymmdd"), 2, "232")
         'Set celda = pobj_Excel.Range("OtrasCTAS!S" & nFil) 'Comentado por PASI 20140409
         Set celda = pobj_Excel.Range("OtrasCTAS!V" & nFil) '*** PASI 20140409 'NAGL 20190713 Cambio de T a V
         celda.value = TotalME
     'CAJAS PLAZO FIJO ***************************************************
        'MONEDA NACIONAL
         TotalMN = SaldoCajasCracsAnexoDiario(Format(dFecIni, "yyyymmdd"), 1, "233")
         Set celda = pobj_Excel.Range("OtrasCTAS!C" & nFil)
         celda.value = TotalMN
        'MONEDA EXTRAJERA
         TotalME = SaldoCajasCracsAnexoDiario(Format(dFecIni, "yyyymmdd"), 2, "233")
         'Set celda = pobj_Excel.Range("OtrasCTAS!R" & nFil) 'Comentado por PASI 20140409
         Set celda = pobj_Excel.Range("OtrasCTAS!U" & nFil) '*** PASI 20140409 'NAGL 20190713 Cambio de S a U
         celda.value = Format(TotalME, "#,##0.00")

     'CRACS AHORROS NAGL 20190713 Según RFC1907080005**************************
        'MONEDA NACIONAL
         TotalMN = SaldoCracsAnexoDiario(Format(dFecIni, "yyyymmdd"), 1, "232")
         Set celda = pobj_Excel.Range("OtrasCTAS!F" & nFil)
         celda.value = TotalMN
        'MONEDA EXTRAJERA
         TotalME = SaldoCracsAnexoDiario(Format(dFecIni, "yyyymmdd"), 2, "232")
         Set celda = pobj_Excel.Range("OtrasCTAS!X" & nFil)
         celda.value = TotalME
         
     'CRACS PLAZO FIJO **************************************************
        'MONEDA NACIONAL
         TotalMN = SaldoCracsAnexoDiario(Format(dFecIni, "yyyymmdd"), 1, "233")
         Set celda = pobj_Excel.Range("OtrasCTAS!E" & nFil)
         celda.value = TotalMN
        'MONEDA EXTRAJERA
         TotalME = SaldoCracsAnexoDiario(Format(dFecIni, "yyyymmdd"), 2, "233")
         'Set celda = pobj_Excel.Range("OtrasCTAS!T" & nFil) 'Comentado por PASI 20140409
         Set celda = pobj_Excel.Range("OtrasCTAS!W" & nFil) '*** PASI 20140409 'NAGL 20190713 Cambio de U a W
         celda.value = TotalME
         
     'COPERATIVAS AHORROS**************************************************
        'MONEDA NACIONAL
         TotalMN = SaldoCoopAhoAnexoDiario(Format(dFecIni, "yyyymmdd"), 1, "232")
         Set celda = pobj_Excel.Range("OtrasCTAS!H" & nFil) 'NAGL 20190713 Cambió de F a H
         celda.value = TotalMN
     
        'MONEDA EXTRANJERA
         TotalME = SaldoCoopAhoAnexoDiario(Format(dFecIni, "yyyymmdd"), 2, "232")
         Set celda = pobj_Excel.Range("OtrasCTAS!Z" & nFil)
         celda.value = TotalME 'NAGL 20190712 Según RFC1907080005
         
    'COPERATIVAS PLAZO FIJO NAGL 20190713 Según RFC1907080005*****************************
        'MONEDA NACIONAL
         TotalMN = SaldoCoopAhoAnexoDiario(Format(dFecIni, "yyyymmdd"), 1, "233")
         Set celda = pobj_Excel.Range("OtrasCTAS!G" & nFil)
         celda.value = TotalMN
     
        'MONEDA EXTRANJERA
         TotalME = SaldoCoopAhoAnexoDiario(Format(dFecIni, "yyyymmdd"), 2, "233")
         Set celda = pobj_Excel.Range("OtrasCTAS!Y" & nFil)
         celda.value = TotalME
         
    'EDYPIMES AHORROS ***********************************************************
        'MONEDA NACIONAL
        TotalMN = SaldoEdypimesAnexoDiario(dFecIni, 1, "232")
        Set celda = pobj_Excel.Range("OtrasCTAS!J" & nFil) 'NAGL 20190713 Cambió de H a J
        celda.value = TotalMN
        
        'MONEDA EXTRANJERA
        TotalME = SaldoEdypimesAnexoDiario(dFecIni, 2, "232")
        Set celda = pobj_Excel.Range("OtrasCTAS!AB" & nFil)
        celda.value = TotalME 'NAGL 20190712 Según RFC1907080005
        
    'EDYPIMES PLAZO FIJO ***********************************************************
        'MONEDA NACIONAL
        TotalMN = SaldoEdypimesAnexoDiario(dFecIni, 1, "233")
        Set celda = pobj_Excel.Range("OtrasCTAS!I" & nFil) 'NAGL 20190713 Cambió de G a I
        celda.value = TotalMN
        
        'MONEDA EXTRANJERA
        TotalME = SaldoEdypimesAnexoDiario(dFecIni, 2, "233")
        Set celda = pobj_Excel.Range("OtrasCTAS!AA" & nFil)
        celda.value = TotalME 'NAGL 20190712 Según RFC1907080005
     
    'END PASI
     
     nFil = nFil + 1
     dFecIni = DateAdd("d", 1, dFecIni)
Loop
End Sub

Private Sub cargarCajaObligExoneradas(ByVal pobj_Excel As Excel.Application)
Dim nFil As Integer
Dim TotalMN As Currency
Dim TotalME As Currency
Dim lnTCF As Currency
Dim dFecIni As Date
Dim dFecAnt As Date
Dim dFecAntME As Date
Dim TotalDias As Integer
Dim celda As Excel.Range

Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsEncDiario As ADODB.Recordset
Set rsEncDiario = New ADODB.Recordset
Dim nCajaChica As Currency
Set rsEncDiario = oNEncajeBCR.ObtenerParamEncajeDiarioxCod("04")
nCajaChica = rsEncDiario!nValor

dFecIni = CDate("01/" & Format(Month(Me.txtFecha.Text), "00") & "/" & Format(Year(Me.txtFecha.Text), "0000"))
TotalDias = DateDiff("d", dFecIni, Me.txtFecha.Text) + 1
nFil = 8
Dim oTC As New nTipoCambio
lnTCF = Format(oTC.EmiteTipoCambio(dFecIni, TCFijoMes), "#,##0.00###")
Do While dFecIni <= Me.txtFecha.Text
    'CAJAS **************************************************************************
        'MONEDA NACIONAL
         TotalMN = SaldoCajasObligExoneradas(Format(dFecIni, "yyyymmdd"), 1)
         'Set celda = pobj_Excel.Range("OtrasCTAS!H" & nFil) 'Comentado por PASI 20140409
         'Set celda = pobj_Excel.Range("OtrasCTAS!I" & nfil) 'PASI 20140409
         Set celda = pobj_Excel.Range("OtrasCTAS!K" & nFil) 'Agregado by NAGL 20190713 Según RFC1907080005
         If TotalMN <> 0 Then
            celda.value = TotalMN + nCajaChica
         Else
            dFecAnt = DateAdd("d", -1, dFecIni)
            TotalMN = SaldoCajasObligExoneradas(Format(dFecAnt, "yyyymmdd"), 1)
            If TotalMN <> 0 Then
                celda.value = TotalMN + nCajaChica
            Else
                dFecAnt = DateAdd("d", -1, dFecAnt)
                TotalMN = SaldoCajasObligExoneradas(Format(dFecAnt, "yyyymmdd"), 1)
                If TotalMN <> 0 Then
                    celda.value = TotalMN + nCajaChica
                Else
                    dFecAnt = DateAdd("d", -1, dFecAnt)
                    TotalMN = SaldoCajasObligExoneradas(Format(dFecAnt, "yyyymmdd"), 1)
                    celda.value = TotalMN + nCajaChica
                End If
            End If
            'Reporte5MN
            'If dFecIni = Me.txtFecha.Text Then
            '    Set celda = pobj_Excel.Range("Reporte5MN!F49")
            '    celda.value = TotalMN + nCajaChica
            'End If 'Comentado por PASI 20140409
         End If
        'MONEDA EXTRANJERA
         TotalME = SaldoCajasObligExoneradas(Format(dFecIni, "yyyymmdd"), 2)
         'Set celda = pobj_Excel.Range("OtrasCTAS!U" & nFil) 'Comentado por PASI20140409
         'Set celda = pobj_Excel.Range("OtrasCTAS!V" & nfil) '*** PASI 20140409 'Comentado by NAGL 20190713
         Set celda = pobj_Excel.Range("OtrasCTAS!AC" & nFil) 'Agregado by NAGL 20190713 Según RFC1907080005
         
         If TotalME <> 0 Then
            celda.value = TotalME
         Else
            dFecAnt = DateAdd("d", -1, dFecIni)
            TotalME = SaldoCajasObligExoneradas(Format(dFecAnt, "yyyymmdd"), 2)
            If TotalME <> 0 Then
                celda.value = TotalME
            Else
                dFecAnt = DateAdd("d", -1, dFecAnt)
                TotalME = SaldoCajasObligExoneradas(Format(dFecAnt, "yyyymmdd"), 2)
                If TotalME <> 0 Then
                    celda.value = TotalME
                Else
                    dFecAnt = DateAdd("d", -1, dFecAnt)
                    TotalME = SaldoCajasObligExoneradas(Format(dFecAnt, "yyyymmdd"), 2)
                    celda.value = TotalME
                End If
            End If
         End If
         
     nFil = nFil + 1
     dFecIni = DateAdd("d", 1, dFecIni)
Loop
    Set rsEncDiario = Nothing
    Set oNEncajeBCR = Nothing
End Sub

Private Sub cargarIntDevengado(ByVal pobj_Excel As Excel.Application)
Dim nFil As Integer
Dim TotalMN As Currency
Dim TotalME As Currency
Dim lnTCF As Currency
Dim dFecIni As Date
Dim TotalDias As Integer
Dim celda As Excel.Range
dFecIni = CDate("01/" & Format(Month(Me.txtFecha.Text), "00") & "/" & Format(Year(Me.txtFecha.Text), "0000"))
TotalDias = DateDiff("d", dFecIni, Me.txtFecha.Text) + 1
nFil = 8
Dim oTC As New nTipoCambio
lnTCF = Format(oTC.EmiteTipoCambio(dFecIni, TCFijoMes), "#,##0.00###")
Do While dFecIni <= Me.txtFecha.Text
    'INT DEVENGADO **************************************************
        'MONEDA NACIONAL
         TotalMN = SaldoIntDevAnexoDiario(Format(dFecIni, "yyyymmdd"), 1)
         'Set celda = pobj_Excel.Range("OtrasCTAS!V" & nFil) 'Comentado PASI20140523
         'Set celda = pobj_Excel.Range("OtrasCTAS!W" & nfil) 'Agregado PASI20140523 'Comentado by NAGL 20190713
         Set celda = pobj_Excel.Range("OtrasCTAS!AD" & nFil) 'NAGL 20190713 Según RFC1907080005
         
         celda.value = TotalMN
        'MONEDA EXTRANJERA
         TotalME = SaldoIntDevAnexoDiario(Format(dFecIni, "yyyymmdd"), 2)
         'Set celda = pobj_Excel.Range("OtrasCTAS!W" & nFil) 'Comentado PASI20140523
         'Set celda = pobj_Excel.Range("OtrasCTAS!X" & nfil) 'Comentado by NAGL 20190713
         Set celda = pobj_Excel.Range("OtrasCTAS!AE" & nFil) 'NAGL 20190713 Según RFC1907080005
         celda.value = TotalME
         
         nFil = nFil + 1
         dFecIni = DateAdd("d", 1, dFecIni)
Loop
End Sub

Private Sub cargarPlazoFijoxRango(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String)
Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsPFxRango As ADODB.Recordset
Set rsPFxRango = New ADODB.Recordset
Dim pcelda As Excel.Range
Dim nFil As Integer
Dim TotalMN As Currency
Dim TotalME As Currency
Dim lnTCF As Currency
Dim dFecIni As Date
Dim TotalDias As Integer
Dim celda As Excel.Range
dFecIni = CDate("01/" & Format(Month(Me.txtFecha.Text), "00") & "/" & Format(Year(Me.txtFecha.Text), "0000"))
TotalDias = DateDiff("d", dFecIni, Me.txtFecha.Text) + 1
If cMoneda = "1" Then
Set pcelda = pobj_Excel.Range("PFxRangos!F8")
    pcelda.value = Me.txtFecha.Text
End If
nFil = 8
Dim oTC As New nTipoCambio
lnTCF = Format(oTC.EmiteTipoCambio(dFecIni, TCFijoMes), "#,##0.000###")

        Set rsPFxRango = oNEncajeBCR.ObtenerPlazoFijoxRangoEncajeDiario(Format(Me.txtFecha.Text, "yyyymmdd"), cMoneda)
        nFil = 13
        If Not rsPFxRango.EOF Or rsPFxRango.BOF Then
            If cMoneda = 1 Then
                Do While Not rsPFxRango.EOF
                    Set pcelda = pobj_Excel.Range("PFxRangos!C" & nFil)
                    pcelda.value = IIf(cMoneda = 1, rsPFxRango(1), 0)
                    nFil = nFil + 1
                    rsPFxRango.MoveNext
                Loop
                Set celda = pobj_Excel.Range("PFxRangos!C32")
                celda.value = Format(lnTCF, gsFormatoNumeroView3Dec)
            ElseIf cMoneda = 2 Then
                Do While Not rsPFxRango.EOF
                    Set pcelda = pobj_Excel.Range("PFxRangos!F" & nFil)
                    pcelda.value = IIf(cMoneda = 2, rsPFxRango(1), 0)
                    nFil = nFil + 1
                    rsPFxRango.MoveNext
                Loop
            End If
        End If
        Set rsPFxRango = Nothing
End Sub

Private Function cargarAnexoDiario(ByVal pobj_Excel As Excel.Application)
Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsEncDiario As ADODB.Recordset
Set rsEncDiario = New ADODB.Recordset
Dim rsEncDiarioCod As ADODB.Recordset
Set rsEncDiarioCod = New ADODB.Recordset
Dim pcelda As Excel.Range
Dim celda As Excel.Range
Dim nMonto As Currency
Dim nParFil As Integer
Dim dFecDia As Date
Dim nDiaMes As Integer
Set rsEncDiario = oNEncajeBCR.ObtenerParamEncajeDiario
Do While Not rsEncDiario.EOF
    Select Case rsEncDiario!cCodParamEncDiario
        Case "01"
                    Set pcelda = pobj_Excel.Range("ANEXODiario!H50")
                    pcelda.value = rsEncDiario!nValor
        Case "02"
                    Set pcelda = pobj_Excel.Range("ANEXODiario!D29")
                    pcelda.value = rsEncDiario!nValor / 100
        Case "03"
                    Set pcelda = pobj_Excel.Range("ANEXODiario!E29")
                    pcelda.value = rsEncDiario!nValor / 100
        Case "04"
                    Set pcelda = pobj_Excel.Range("ANEXODiario!D45")
                    pcelda.value = rsEncDiario!nValor
        Case "05"
                    Set pcelda = pobj_Excel.Range("ANEXODiario!D46")
                    pcelda.value = rsEncDiario!nValor
        Case "06"
                    Set pcelda = pobj_Excel.Range("ANEXODiario!H35")
                    pcelda.value = rsEncDiario!nValor
        Case "07"
                    Set pcelda = pobj_Excel.Range("ANEXODiario!H44")
                    pcelda.value = rsEncDiario!nValor
        Case "08"
                    Set rsEncDiarioCod = oNEncajeBCR.ObtenerParamEncajeDiarioxCod("10")
                    nDiaMes = rsEncDiarioCod!nValor
                    Set pcelda = pobj_Excel.Range("ANEXODiario!H37") 'NAGL Cambió de H38 A H37
                    pcelda.value = rsEncDiario!nValor '/ nDiaMes 'NAGL Comentó esta parte
        'Comentado por PASI 20140409
'        Case "09"
'                    Set pcelda = pobj_Excel.Range("ANEXODiario!H48")
'                    pcelda.value = rsEncDiario!nValor
        
        'end PASI
        
        Case "10"
                    'Modificado por PASI20140409
                    
                    'Set pcelda = pobj_Excel.Range("ANEXODiario!N36")
                    'pcelda.value = rsEncDiario!nValor
                    
                    Set pcelda = pobj_Excel.Range("AnexoDiario!H36")
                    pcelda.Formula = "=(H35/" & Round(rsEncDiario!nValor, 2) & ")"
                    
                    'end PASI
                    
        Case "11"
                    Set pcelda = pobj_Excel.Range("ANEXODiario!H43")
                    pcelda.value = rsEncDiario!nValor
        Case "12"
                    Set pcelda = pobj_Excel.Range("ANEXODiario!H54")
                    pcelda.value = rsEncDiario!nValor / 100
        Case "13"
                    Set pcelda = pobj_Excel.Range("ANEXODiario!N46")
                    pcelda.value = rsEncDiario!nValor
                    
        Case "30" '***PASI 20140409
                    Set pcelda = pobj_Excel.Range("ANEXODiario!H60")
                    pcelda.value = rsEncDiario!nValor / 100
        
        Case "31" '***PASI 20140409
                    Set pcelda = pobj_Excel.Range("ANEXODiario!H49")
                    pcelda.value = rsEncDiario!nValor / 100
        Case "32" '*** PASI 20140409
                    Set pcelda = pobj_Excel.Range("ANEXODiario!H43")
                    pcelda.value = rsEncDiario!nValor / 100
    End Select
    rsEncDiario.MoveNext
Loop
    dFecDia = CDate(Format(Day(Me.txtFecha.Text), "00") & "/" & Format(Month(Me.txtFecha.Text), "00") & "/" & Format(Year(Me.txtFecha.Text), "0000"))
    nParFil = Day(dFecDia) + 11
'Modificado PASI20140409
    'Set celda = pobj_Excel.Range("Reporte5MN!F50")
    'celda.Formula = "=ENCAJELEGALMN!O" & nParFil
    
    'Set celda = pobj_Excel.Range("Reporte5ME!F47")
    'celda.Formula = "=ENCAJELEGALME!M" & nParFil
    
    Set celda = pobj_Excel.Range("Reporte5MN!F51")
    celda.Formula = "=ENCAJELEGALMN!P" & nParFil
    
    Set celda = pobj_Excel.Range("Reporte5ME!F51")
    celda.Formula = "=ENCAJELEGALME!N" & nParFil
    
    Set celda = pobj_Excel.Range("Reporte5MN!F50")
    celda.Formula = "=OtrasCtas!I" & (Day(dFecDia) + 7)
    
    'END PASI
End Function

Private Function SaldoCtas(lnNroCol As Long, lsOpeCod As String, ByVal ldFecha As Date, ByVal pdFecFin As Date, ByVal pnTCF As Currency, ByVal pnTCFF As Currency) As Currency
Dim rsC As New ADODB.Recordset
Dim oRep As New DRepCtaColumna
Set rsC = oRep.GetRepColumnaCtaSaldo(lsOpeCod, lnNroCol, ldFecha, gbBitCentral)
SaldoCtas = 0
If Not RSVacio(rsC) Then
    If Mid(lsOpeCod, 3, 1) = "1" Then
        SaldoCtas = IIf(IsNull(rsC!Total), 0, rsC!Total)
    Else
        If lnNroCol = 29 Then
            SaldoCtas = Round(IIf(IsNull(rsC!TotalME), 0, rsC!TotalME), 1)
        Else
            If pdFecFin = ldFecha Then
               SaldoCtas = IIf(IsNull(rsC!Total), 0, Round(rsC!Total / pnTCFF, 2))
            Else
               SaldoCtas = IIf(IsNull(rsC!Total), 0, Round(rsC!Total / pnTCF, 2))
            End If
        End If
    End If
End If
Set oRep = Nothing
RSClose rsC
End Function

Private Function SaldoAhoPlaFijCTS(ByVal ldFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Currency
Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerSaldoAhoPlaFijCTS(ldFecha, nMoneda, cProducto)
SaldoAhoPlaFijCTS = 0
    SaldoAhoPlaFijCTS = Round(rsC!nSaldo, 2)
Set oNEncajeBCR = Nothing
RSClose rsC
End Function

Private Function SaldoChequeAhoPlaFijCTS(ByVal ldFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Currency
Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerSaldoChequeAhoPlaFijCTS(ldFecha, nMoneda, cProducto)
SaldoChequeAhoPlaFijCTS = 0
    SaldoChequeAhoPlaFijCTS = Round(rsC!nSaldo, 2)
Set oNEncajeBCR = Nothing
RSClose rsC
End Function

Private Function FechaAnioMesDia(ByVal obj_excel As Excel.Application)
Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsC As New ADODB.Recordset
Dim celda As Excel.Range

Set rsC = oNEncajeBCR.ObtenerFechaAnioMesDia(Format(Me.txtFecha.Text, "yyyymmdd"))
        Set celda = obj_excel.Range("ANEXODiario!H2")
        celda.value = rsC!Anio
        Set celda = obj_excel.Range("ANEXODiario!H3")
        celda.value = rsC!MES
        Set celda = obj_excel.Range("ANEXODiario!H4")
        celda.value = rsC!Dia
        '*****************************************************************
        Set celda = obj_excel.Range("'FORMATO CALCULO ENCAJE MN BCRP'!B7")
        celda.value = rsC!MES & " " & rsC!Anio
        '******************'Agregado by NAGL 20190918*********************
        Set celda = obj_excel.Range("'FORMATO CALCULO ENCAJE MN BCRP'!J9")
        celda.value = "TOTAL CAJA PERIODO ACTUAL " & rsC!Anio
        '*****************************************************************
        Set celda = obj_excel.Range("'FORMATO CALCULO ENCAJE ME BCRP'!D7")
        celda.value = Day(DateAdd("D", -Day(DateAdd("M", 1, CDate(txtFecha.Text))), DateAdd("M", 1, CDate(txtFecha.Text))))
        '*****************************************************************
Set oNEncajeBCR = Nothing
RSClose rsC
End Function

Private Function SaldoBCRPAnexoDiario(ByVal ldFecha As String, ByVal nMoneda As Integer) As Currency
Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerSaldoBCRPEncajeDiario(ldFecha, nMoneda)
SaldoBCRPAnexoDiario = 0
If rsC.RecordCount > 0 Then
    SaldoBCRPAnexoDiario = Round(rsC!nImporte, 2)
Else
    SaldoBCRPAnexoDiario = 0
End If
    
Set oNEncajeBCR = Nothing
RSClose rsC
End Function

Private Function SaldoCajasCracsAnexoDiario(ByVal ldFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Currency
Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerSaldoCajasCracsEncajeDiario(ldFecha, nMoneda, cProducto)
SaldoCajasCracsAnexoDiario = 0
If rsC.RecordCount > 0 Then
    SaldoCajasCracsAnexoDiario = Round(rsC!nSaldo, 2)
Else
    SaldoCajasCracsAnexoDiario = 0
End If
    
Set oNEncajeBCR = Nothing
RSClose rsC
End Function

Private Function SaldoCajasObligExoneradas(ByVal ldFecha As String, ByVal nMoneda As Integer) As Currency
Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerSaldoCajasObligExoneradas(ldFecha, nMoneda)
SaldoCajasObligExoneradas = 0
If rsC.RecordCount > 0 Then
    SaldoCajasObligExoneradas = Round(rsC!nMonto, 2)
Else
    SaldoCajasObligExoneradas = 0
End If
    
Set oNEncajeBCR = Nothing
RSClose rsC
End Function

Private Function SaldoIntDevAnexoDiario(ByVal ldFecha As String, ByVal nMoneda As Integer) As Currency
Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerIntDevEncajeDiario(ldFecha, nMoneda)
SaldoIntDevAnexoDiario = 0
If rsC.RecordCount > 0 Then
    SaldoIntDevAnexoDiario = Round(rsC!nSaldo, 2)
Else
    SaldoIntDevAnexoDiario = 0
End If
    
Set oNEncajeBCR = Nothing
RSClose rsC
End Function

Private Function SaldoCoopAhoAnexoDiario(ByVal ldFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Currency
Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerSaldoCoopAhoEncajeDiario(ldFecha, nMoneda, cProducto)
SaldoCoopAhoAnexoDiario = 0
If rsC.RecordCount > 0 Then
    SaldoCoopAhoAnexoDiario = Round(rsC!nSaldo, 2)
Else
    SaldoCoopAhoAnexoDiario = 0
End If

Set oNEncajeBCR = Nothing
RSClose rsC
End Function

Private Function SaldoCracsAnexoDiario(ByVal ldFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Currency
Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerSaldoCracsEncajeDiario(ldFecha, nMoneda, cProducto)
SaldoCracsAnexoDiario = 0
If rsC.RecordCount > 0 Then
    SaldoCracsAnexoDiario = Round(rsC!nSaldo, 2)
Else
    SaldoCracsAnexoDiario = 0
End If

Set oNEncajeBCR = Nothing
RSClose rsC
End Function

'***PASI 20140409
Private Function SaldoEdypimesAnexoDiario(ByVal pdFecha As Date, ByVal pnMoneda As Integer, ByVal pnProducto As Integer) As Currency
    Dim onEncaje As nEncajeBCR
    Set onEncaje = New nEncajeBCR
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = onEncaje.ObtenerSaldoEdypimesEncajeDiario(pdFecha, pnMoneda, pnProducto)
    If Not rs.EOF And Not rs.BOF Then
        SaldoEdypimesAnexoDiario = Round(rs!nSaldo, 2)
    Else
        SaldoEdypimesAnexoDiario = 0
    End If
    Set onEncaje = Nothing
    RSClose rs
End Function
'***END PASI

Private Function SaldoAhoExoEnc() As Currency
Dim oNEncajeBCR As nEncajeBCR
Set oNEncajeBCR = New nEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerSaldoAhoExoEnc()
SaldoAhoExoEnc = 0
    SaldoAhoExoEnc = Round(rsC!nValor, 2)
Set oNEncajeBCR = Nothing
RSClose rsC
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

Private Sub FEParaEncDiario_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim lnValor As Currency
    lnValor = CCur(IIf(FEParaEncDiario.TextMatrix(pnRow, pnCol) = "", "0", FEParaEncDiario.TextMatrix(pnRow, pnCol)))
    If pnRow = 2 Or pnRow = 3 Or pnRow = 12 Or pnRow = 13 Then
        FEParaEncDiario.TextMatrix(pnRow, 3) = Format(lnValor, "#,##0.0000")
    Else
        FEParaEncDiario.TextMatrix(pnRow, 3) = Format(lnValor, "#,##0.00")
    End If
    
End Sub

Private Sub Form_Load()
    cargarParamEncDiario
    txtFecha = gdFecSis
End Sub

Private Sub cargarParamEncDiario()
    Dim oNEncajeBCR As nEncajeBCR
    Set oNEncajeBCR = New nEncajeBCR
    Dim rsEncDiario As ADODB.Recordset
    Set rsEncDiario = New ADODB.Recordset
    Dim i As Integer
    Set rsEncDiario = oNEncajeBCR.ObtenerParamEncajeDiario()
    If Not rsEncDiario.BOF And Not rsEncDiario.EOF Then
        i = 1
        FEParaEncDiario.lbEditarFlex = True
        Do While Not rsEncDiario.EOF
            FEParaEncDiario.AdicionaFila
            FEParaEncDiario.TextMatrix(i, 1) = rsEncDiario!cParamEncDiario
            FEParaEncDiario.TextMatrix(i, 2) = rsEncDiario!cTipoEncDiario
            If rsEncDiario!cCodParamEncDiario = "02" Or rsEncDiario!cCodParamEncDiario = "03" Or rsEncDiario!cCodParamEncDiario = "12" Or rsEncDiario!cCodParamEncDiario = "13" Then
                FEParaEncDiario.TextMatrix(i, 3) = Format(rsEncDiario!nValor, "#,##0.0000")
            ElseIf rsEncDiario!cCodParamEncDiario = "10" Then
                FEParaEncDiario.TextMatrix(i, 3) = Format(rsEncDiario!nValor, "#,##0")
            Else
                FEParaEncDiario.TextMatrix(i, 3) = Format(rsEncDiario!nValor, "#,##0.00")
            End If
            FEParaEncDiario.TextMatrix(i, 4) = rsEncDiario!cCodParamEncDiario
            FEParaEncDiario.TextMatrix(i, 5) = rsEncDiario!nEstado
            i = i + 1
            rsEncDiario.MoveNext
        Loop
    End If
    Set rsEncDiario = Nothing
    Set oNEncajeBCR = Nothing
End Sub

Private Sub txtFecha_GotFocus()
    fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If ValidaFecha(txtFecha.Text) <> "" Then
          MsgBox "Fecha no válida...!", vbInformation, "¡Aviso!"
          Exit Sub
       End If
    End If
End Sub
'****PASI 20140409
Private Sub CargarProyecciones(ByVal pobj_Excel As Excel.Application)
Dim TotalMN As Currency
Dim TotalME As Currency
Dim dFecIni As Date
Dim ldFechaRef As Date
Dim X As Integer, nDias As Integer, nFil As Integer
Dim celda As Excel.Range
Dim lnProy As Currency

dFecIni = CDate("01/" & Format(Month(Me.txtFecha.Text), "00") & "/" & Format(Year(Me.txtFecha.Text), "0000"))
ldFechaRef = DateAdd("D", -1, DateAdd("M", 1, dFecIni))
nDias = DateDiff("d", dFecIni, ldFechaRef) + 1

nFil = 12
X = 0
Do While dFecIni <= Me.txtFecha.Text
    X = X + 1
    nFil = nFil + 1
    dFecIni = DateAdd("D", 1, dFecIni)
Loop

    If X < nDias Then
        Do While dFecIni <= ldFechaRef
        'DEPOSITOS A PLAZO FIJO
            'MONEDA NACIONAL
            Set celda = pobj_Excel.Range("ENCAJEPROYMN!C" & (nFil - 1))
            TotalMN = celda.value
            lnProy = SaldoDepPlaFijCalenProy(dFecIni, 1)
            Set celda = pobj_Excel.Range("ENCAJEPROYMN!C" & (nFil))
            celda.value = TotalMN + lnProy

            'MONEDA EXTRANJERA
            Set celda = pobj_Excel.Range("ENCAJEPROYME!C" & (nFil - 1))
            TotalMN = celda.value
            lnProy = SaldoDepPlaFijCalenProy(dFecIni, 2)
            Set celda = pobj_Excel.Range("ENCAJEPROYME!C" & (nFil))
            celda.value = TotalMN + lnProy

        'DEPOSITOS DE AHORRO
            'MONEDA NACIONAL
            Set celda = pobj_Excel.Range("ENCAJEPROYMN!D" & (nFil - 1))
            TotalMN = celda.value
            lnProy = SaldoDepAhoCalenProy(dFecIni, 1)
            Set celda = pobj_Excel.Range("ENCAJEPROYMN!D" & (nFil))
            celda.value = TotalMN + lnProy

            'MONEDA EXTRANJERA
            Set celda = pobj_Excel.Range("ENCAJEPROYME!D" & (nFil - 1))
            TotalMN = celda.value
            lnProy = SaldoDepAhoCalenProy(dFecIni, 2)
            Set celda = pobj_Excel.Range("ENCAJEPROYME!D" & (nFil))
            celda.value = TotalMN + lnProy

        'DEPOSITOS BCRP
            'MONEDA NACIONAL
            Set celda = pobj_Excel.Range("ENCAJEPROYMN!N" & (nFil - 1))
            TotalMN = celda.value
            lnProy = SaldoDepBcrCalenProy(dFecIni, 1)
            Set celda = pobj_Excel.Range("ENCAJEPROYMN!N" & (nFil))
            celda.value = TotalMN + lnProy

            'MONEDA EXTRANJERA
            Set celda = pobj_Excel.Range("ENCAJEPROYME!L" & (nFil - 1))
            TotalMN = celda.value
            lnProy = SaldoDepBcrCalenProy(dFecIni, 2)
            Set celda = pobj_Excel.Range("ENCAJEPROYME!L" & (nFil))
            celda.value = TotalMN + lnProy

        'OBLIGACIONES INMEDIATAS
            'MONEDA NACIONAL
            Set celda = pobj_Excel.Range("ENCAJEPROYMN!F" & (nFil - 1))
            TotalMN = celda.value
            lnProy = SaldoOblInmCalenProy(dFecIni, 1)
            Set celda = pobj_Excel.Range("ENCAJEPROYMN!F" & (nFil))
            celda.value = TotalMN + lnProy

            'MONEDA EXTRANJERA
            Set celda = pobj_Excel.Range("ENCAJEPROYME!F" & (nFil - 1))
            TotalMN = celda.value
            lnProy = SaldoOblInmCalenProy(dFecIni, 2)
            Set celda = pobj_Excel.Range("ENCAJEPROYME!F" & (nFil))
            celda.value = TotalMN + lnProy

        'EFECTIVO CAJA
            'MONEDA NACIONAL
             Set celda = pobj_Excel.Range("ENCAJEPROYMN!M" & (nFil - 1))
            TotalMN = celda.value
            lnProy = SaldoEfecCajaCalenProy(dFecIni, 1)
            Set celda = pobj_Excel.Range("ENCAJEPROYMN!M" & (nFil))
            celda.value = TotalMN + lnProy

            'MONEDA EXTRANJERA
             Set celda = pobj_Excel.Range("ENCAJEPROYME!K" & (nFil - 1))
            TotalMN = celda.value
            lnProy = SaldoEfecCajaCalenProy(dFecIni, 2)
            Set celda = pobj_Excel.Range("ENCAJEPROYME!K" & (nFil))
            celda.value = TotalMN + lnProy

            nFil = nFil + 1
            dFecIni = DateAdd("D", 1, dFecIni)
        Loop
    End If
End Sub
Private Function SaldoDepPlaFijCalenProy(ByVal pdFecha As Date, ByVal pnMoneda As Integer) As Currency
    Dim oEncaje As nEncajeBCR
    Set oEncaje = New nEncajeBCR
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = oEncaje.ObtenerSaldoDepPlaCalenProy(pdFecha, pnMoneda)
    If Not rs.EOF And Not rs.BOF Then
        SaldoDepPlaFijCalenProy = Round(rs!nSaldoProy, 2)
    Else
        SaldoDepPlaFijCalenProy = 0
    End If
    Set oEncaje = Nothing
    RSClose rs
End Function
Private Function SaldoDepAhoCalenProy(ByVal pdFecha As Date, ByVal pnMoneda As Integer) As Currency
    Dim oEncaje As nEncajeBCR
    Set oEncaje = New nEncajeBCR
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = oEncaje.ObtenerSaldoDepAhoCalenProy(pdFecha, pnMoneda)
    If Not rs.EOF And Not rs.BOF Then
        SaldoDepAhoCalenProy = Round(rs!nSaldoProy, 2)
    Else
        SaldoDepAhoCalenProy = 0
    End If
    Set oEncaje = Nothing
    RSClose rs
End Function
Private Function SaldoDepBcrCalenProy(ByVal pdFecha As Date, ByVal pnMoneda As Integer) As Currency
    Dim oEncaje As nEncajeBCR
    Set oEncaje = New nEncajeBCR
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = oEncaje.ObtenerSaldoDepBcrpCalenProy(pdFecha, pnMoneda)
    If Not rs.EOF And Not rs.BOF Then
        SaldoDepBcrCalenProy = Round(rs!nSaldoProy, 2)
    Else
        SaldoDepBcrCalenProy = 0
    End If
    Set oEncaje = Nothing
    RSClose rs
End Function
Private Function SaldoOblInmCalenProy(ByVal pdFecha As Date, ByVal pnMoneda As Integer) As Currency
    Dim oEncaje As nEncajeBCR
    Set oEncaje = New nEncajeBCR
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = oEncaje.ObtenerSaldoOblInmCalenProy(pdFecha, pnMoneda)
    If Not rs.EOF And Not rs.BOF Then
        SaldoOblInmCalenProy = Round(rs!nSaldoProy, 2)
    Else
        SaldoOblInmCalenProy = 0
    End If
    Set oEncaje = Nothing
    RSClose rs
End Function
Private Function SaldoEfecCajaCalenProy(ByVal pdFecha As Date, ByVal pnMoneda As Integer) As Currency
    Dim oEncaje As nEncajeBCR
    Set oEncaje = New nEncajeBCR
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = oEncaje.ObtenerSaldoEfecCajaCalenProy(pdFecha, pnMoneda)
    If Not rs.EOF And Not rs.BOF Then
        SaldoEfecCajaCalenProy = Round(rs!nSaldoProy, 2)
    Else
        SaldoEfecCajaCalenProy = 0
    End If
    Set oEncaje = Nothing
    RSClose rs
End Function
'END PASI

