VERSION 5.00
Begin VB.Form frmRepResponsabilityConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta - Reporte Responsability"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   Icon            =   "frmRepResponsabilityConsulta.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Extornar"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin Sicmact.FlexEdit fgReporte 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4260
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Mes-Año-cMovNro-nIdRep"
      EncabezadosAnchos=   "300-800-1200-2500-0"
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
      ColumnasAEditar =   "X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmRepResponsabilityConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************
'Nombre : frmRepResponsabilityConsulta
'Descripcion:Formulario para Consultar los Reportes de Responsability
'Creacion: PASI TI-ERS087-2014
'*****************************Dim ldFecPeriodo As Date
Option Explicit
Dim ldFecPeriodo As Date
Private Type TCtaCont
    CuentaContable As String
    Saldo As Currency
    bSaldoA As Boolean
    bSaldoD As Boolean
End Type
Private Sub cmdEliminar_Click()
     Dim oDResponblt As DResponsability
     Set oDResponblt = New DResponsability
     
    If fgReporte.TextMatrix(fgReporte.row, 1) = "" Then
        MsgBox "No existen Datos para Eliminar. Verifique", vbInformation + vbOKOnly, "Aviso"
        Exit Sub
    End If
    Dim row As Integer
    row = fgReporte.row
    ldFecPeriodo = DateAdd("D", -1, DateAdd("M", 1, CDate("01/" & fgReporte.TextMatrix(row, 1) & "/" & fgReporte.TextMatrix(row, 2))))
    
    If Not (ldFecPeriodo = DateAdd("D", -1, "01/" & Format(DatePart("M", gdFecSis), "00") & "/" & Format(DatePart("YYYY", gdFecSis), "0000"))) Then
        MsgBox "No se puede Eliminar los Reportes anteriores a " & dameNombreMes(DatePart("M", DateAdd("D", -1, "01/" & Format(DatePart("M", gdFecSis), "00") & "/" & Format(DatePart("YYYY", gdFecSis), "0000")))) & " del " & DatePart("YYYY", DateAdd("D", -1, "01/" & Format(DatePart("M", gdFecSis), "00") & "/" & Format(DatePart("YYYY", gdFecSis), "0000"))) & ". Verifique", vbExclamation + vbOKOnly, "Aviso"
        Exit Sub
    End If
    If MsgBox("Esta seguro de Extornar el Reporte Seleccionado: " & dameNombreMes(fgReporte.TextMatrix(row, 1)) & " del " & fgReporte.TextMatrix(row, 2) & " ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    oDResponblt.ExtornaReporte (fgReporte.TextMatrix(row, 4))
    CargaReportes
End Sub
Private Sub cmdGenerar_Click()
    If fgReporte.TextMatrix(fgReporte.row, 1) = "" Then
        MsgBox "No existen Datos para generar. Verifique", vbInformation + vbOKOnly, "Aviso"
        Exit Sub
    End If
    GeneraReporte
End Sub
Private Sub GeneraReporte()
    Dim oDResponblt As DResponsability
    Dim celda As Excel.Range
    Dim row As Integer
    row = fgReporte.row
    ldFecPeriodo = DateAdd("D", -1, DateAdd("M", 1, CDate("01/" & fgReporte.TextMatrix(row, 1) & "/" & fgReporte.TextMatrix(row, 2))))

    Set oDResponblt = New DResponsability
    If Not oDResponblt.ExisteConfigRepResponsability(Format(DatePart("M", ldFecPeriodo), "00"), fgReporte.TextMatrix(row, 2)) Then
        MsgBox "Aún no se ha generado la información para este periodo. Verifique.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    Dim sPathFormatoResponsability As String
    
    Dim fs As New Scripting.FileSystemObject
    Dim obj_excel As Object, Libro As Object, Hoja As Object
    
    On Error GoTo error_sub
    
    sPathFormatoResponsability = App.path & "\Spooler\RepResponsability_" + Format(ldFecPeriodo, "yyyymmdd") + ".xlsx"
    If fs.FileExists(sPathFormatoResponsability) Then
        If ArchivoEstaAbierto(sPathFormatoResponsability) Then
            If MsgBox("Debe Cerrar el Archivo: " + fs.GetFileName(sPathFormatoResponsability) + " para continuar", vbRetryCancel) = vbCancel Then
                Me.MousePointer = vbDefault
                Exit Sub
            End If
            Me.MousePointer = vbHourglass
        End If
        fs.DeleteFile sPathFormatoResponsability, True
    End If
    sPathFormatoResponsability = App.path & "\FormatoCarta\FormatoResponsabiility.xlsx"
    If Len(Dir(sPathFormatoResponsability)) = 0 Then
           MsgBox "No se Pudo Encontrar el Archivo:" & sPathFormatoResponsability, vbCritical
           Me.MousePointer = vbDefault
           Exit Sub
    End If
    
    Set obj_excel = CreateObject("Excel.Application")
    obj_excel.DisplayAlerts = False
    Set Libro = obj_excel.Workbooks.Open(sPathFormatoResponsability)
    Set Hoja = Libro.ActiveSheet
    
    CargaData obj_excel
    
    sPathFormatoResponsability = App.path & "\Spooler\RepResponsability_" + Format(ldFecPeriodo, "yyyymmdd") + ".xlsx"
    If fs.FileExists(sPathFormatoResponsability) Then
        If ArchivoEstaAbierto(sPathFormatoResponsability) Then
            MsgBox "Debe Cerrar el Archivo:" + fs.GetFileName(sPathFormatoResponsability)
        End If
        fs.DeleteFile sPathFormatoResponsability, True
    End If
    Hoja.SaveAs sPathFormatoResponsability
    Libro.Close
    obj_excel.Quit
    Set Hoja = Nothing
    Set Libro = Nothing
    Set obj_excel = Nothing
    Me.MousePointer = vbDefault
    
    Dim m_excel As New Excel.Application
    m_excel.Workbooks.Open (sPathFormatoResponsability)
    m_excel.Visible = True
    Exit Sub
error_sub:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
        Set Libro = Nothing
        Set obj_excel = Nothing
        Set Hoja = Nothing
        Me.MousePointer = vbDefault
End Sub
Private Sub CargaData(ByVal pobj_Excel As Excel.Application)
   Dim oDResponblt As DResponsability
   Set oDResponblt = New DResponsability
   
   Dim nfil As Integer
   Dim celdaCampo As Excel.Range
   Dim celdaValor As Excel.Range
   
   'Nombre de la Institucion
   nfil = 2
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = "CMAC Maynas"
   'Cifras al
   nfil = 3
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = DatePart("D", ldFecPeriodo) & "-" & Left(dameNombreMes(DatePart("M", ldFecPeriodo)), 3) & Right(DatePart("YYYY", ldFecPeriodo), 2)
   'Moneda Local
   nfil = 4
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = "PEN"
   'Número de prestatarios
   nfil = 5
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = oDResponblt.ObtieneValorNumeroPrestatario(ldFecPeriodo)
   'Número de ahorrantes (excl. ahorros forzosos)
    nfil = 6
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = oDResponblt.ObtieneValorNumeroAhorrantes(ldFecPeriodo)
   'Caja y Bancos
   nfil = 7
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 2), 0)
   'Inversiones Financieras
   nfil = 8
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 10), 0)
   'Cartera bruta (vigentes y de largo plazo)
   nfil = 10
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   Dim nValVig, nValRefinan, nValVenc, nValCobJudi As Currency
   nValVig = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 15), 0)
   nValRefinan = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 17), 0)
   nValVenc = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 18), 0)
   nValCobJudi = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 19), 0)
   celdaValor.value = CCur(Round(nValVig + nValRefinan + nValVenc + nValCobJudi, 0))
   'Reservas créditos vencidos
   nfil = 11
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = CCur(Round(Abs(ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 20), 0)), 0))
   'Otras cuentas corrientes
   nfil = 12
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   Dim nValCtaxCobrar, nValBienReal, nValorActivIntangible, nValorImpCte, nValorImpDif, nValOtroActiv As Currency
   nValCtaxCobrar = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 24), 0)
   nValBienReal = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 25), 0)
   nValorActivIntangible = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 31), 0)
   nValorImpCte = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 32), 0)
   nValorImpDif = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 33), 0)
   nValOtroActiv = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 35), 0)
   celdaValor.value = nValCtaxCobrar + nValBienReal + nValorActivIntangible + nValorImpCte + nValorImpDif + nValOtroActiv
   'Activos fijos netos(y otros activos no corrientes)
   nfil = 13
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   Dim nValorParticip, nValorInmueble As Currency
   nValorParticip = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 28), 0)
   nValorInmueble = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 29), 0)
   celdaValor.value = nValorParticip + nValorInmueble
   'Ahorros y Depositos a Plazo (excl. Ahorros forzosos)
   nfil = 16
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   Dim nValorOblPub, nValorDepESF As Currency
   nValorOblPub = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 37), 0)
   nValorDepESF = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 43), 0)
   celdaValor.value = nValorOblPub + nValorDepESF
   'Creditos en moneda local /de IFIs, bancos y otros
   nfil = 17
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = CCur(oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Créditos en Moneda Local de IFIs, Bancos y otros:"))
   'Creditos en divisas (expresado en moneda local)
   nfil = 18
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = CCur(oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Créditos en divisas (expresado en moneda local):"))
   'Otros pasivos
   nfil = 19
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = Abs(ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 62), 0))
   'Créditos subordinados en moneda local
   nfil = 20
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = CCur(oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Créditos subordinados en moneda local:"))
   'Créditos subordinados en divisas (expresado en moneda local)
   nfil = 21
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = CCur(oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Créditos subordinados en divisas (expresado en moneda local):"))
   'Patrimonio
   nfil = 22
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = Abs(ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 64), 0))
   'Monto total mantenido en bancos u otras instituciones financieras para financiacion (back to back)
   nfil = 24
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = 0
   'Activos Dado en garantia hacia los refinanciados
   nfil = 25
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = Abs(oDResponblt.ObtieneValorActivGarantia(ldFecPeriodo))
   'Cartera en administracion (Ver Hoja 'Definitions' Linea 26)
   nfil = 26
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = 0
   'Activos denominados o indexados en USD
   nfil = 28
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 1), 2)
   'Pasivos denominados o indexados en USD (no cubiertos)
   nfil = 29
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 63), 2)
   'Ratio de suficiencia de capital
   nfil = 30
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = CCur((oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Ratio de suficiencia del capital:"))) / 100
   'Cantidad de meses
   nfil = 31
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = DatePart("M", ldFecPeriodo)
   'Ingresos por intereses(cartera de créditos)
   nfil = 32
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 8), 0)
   'Ingresos por comisiones (cartera de crédito)
   nfil = 33
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = 0
   'Gastos operativos
   nfil = 34
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
            'Gastos de personal
            Dim nValorGastoPers As Currency
            nValorGastoPers = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 51), 0)
            'Gastos Administrativos
            Dim nValorGastosAdmin As Currency
            nValorGastosAdmin = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 52), 0)
            'Otros Gastos Operativos
            Dim nValorOtrosGastosOpe, nValorPrimaFondo, nValorGastosDiv, nValorImpContri, nValorDeprecAmort, nValorValuaActivProv As Currency
            nValorPrimaFondo = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 37), 0)
            nValorGastosDiv = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 38), 0)
            nValorImpContri = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 53), 0)
            nValorDeprecAmort = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 54), 0)
            nValorValuaActivProv = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 56), 0)
            nValorOtrosGastosOpe = nValorPrimaFondo + nValorGastosDiv + nValorImpContri + nValorDeprecAmort + (nValorValuaActivProv * 2)
    celdaValor.value = nValorGastoPers + nValorGastosAdmin + nValorOtrosGastosOpe
    'Resultado neto del ejercicio despues de impuestos
    nfil = 35
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = Abs(oDResponblt.ObtieneValorCalCartera_ResNetoEjercicioByVal(ldFecPeriodo))
    
    ' PAR 1-30 Dias
    nfil = 36
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = Abs(oDResponblt.ObtieneValorCalCartera_PAR1and30Dias(ldFecPeriodo))
    'PAR > 30 dias
    nfil = 37
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = Abs(oDResponblt.ObtieneValorCalCartera_PARMayor30Dias(ldFecPeriodo))
    'Creditos reestructurados/reprogramados/refinanciados(no incluidos en PAR > 30)
    nfil = 38
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = Abs(oDResponblt.ObtieneValorCalCartera_CredReesctruc(ldFecPeriodo))
    'Total de Castigos
    nfil = 39
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = Abs(oDResponblt.ObtieneValorCalCartera_TotalCastigos(ldFecPeriodo))
    'Informe emtido/recibido en el periodo (Si/No)
    nfil = 40
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = (oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Informe emitido/recibido en el periodo (Si/No):"))
    'Agencia de Calificacion
    nfil = 41
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = (oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Agencia de Calificación:"))
    'Fecha del Reporte
    nfil = 42
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = (oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Fecha del Reporte:"))
    'Calificación:
    nfil = 43
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = (oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Calificación:"))
    'Esta su institucion incumpliendo algun covenant:
    nfil = 44
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = (oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "¿Está su institución rompiendo algún covenant y/o en default con algún acreedor (si/no)?"))
    'Si la respuesta es afirmativa, favor de explicar
    nfil = 45
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = (oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Si la respuesta es afirmativa, favor explicar"))
End Sub
Private Sub Form_Load()
    CargaReportes
End Sub
Private Sub CargaReportes()
    Dim rs As ADODB.Recordset
    Dim oDRep As DResponsability
    Dim row As Integer
    Set oDRep = New DResponsability
    
    Set rs = oDRep.ListaRepResponsability
'    If (rs.EOF And rs.BOF) Then
'        MsgBox "No existen Datos.", vbYesNo + vbInformation, "Aviso"
'        Exit Sub
'    End If
    LimpiaFlex fgReporte
    Do While Not rs.EOF
        fgReporte.AdicionaFila
            row = fgReporte.row
            fgReporte.TextMatrix(row, 1) = rs!cMES
            fgReporte.TextMatrix(row, 2) = rs!cAnio
            fgReporte.TextMatrix(row, 3) = rs!cMovNro
            fgReporte.TextMatrix(row, 4) = rs!nIdRep
        rs.MoveNext
    Loop
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
Private Function DepuraSaldoAD(ByVal sCta As String) As String
Dim i As Integer
Dim Cad As String
    Cad = ""
    For i = 1 To Len(sCta)
        If Mid(sCta, i, 1) >= "0" And Mid(sCta, i, 1) <= "9" Then
            Cad = Cad + Mid(sCta, i, 1)
        End If
    Next i
    DepuraSaldoAD = Cad
End Function
Private Function ObtenerResultadoFormula(ByVal pdFecha As Date, ByVal psFormula As String, ByVal pnMoneda As Integer, Optional psAgencia As String = "") As Currency
    Dim oBal As New DbalanceCont
    Dim oNBal As New NBalanceCont
    Dim oFormula As New NInterpreteFormula
    Dim lsFormula As String, lsTmp As String, lsTmp1 As String, lsCadFormula As String
    Dim MatDatos() As TCtaCont
    Dim i As Long, j As Long, nCtaCont As Long
    Dim sTempAD As String
    Dim nPosicion As Integer
    Dim signo As String
    Dim LsSigno As String
    lsFormula = Trim(psFormula)
    ReDim MatDatos(0)
    nCtaCont = 0
    lsTmp = ""
    lsFormula = Replace(lsFormula, "M", pnMoneda)
    sTempAD = ""
    For i = 1 To Len(lsFormula)
        If (Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Then
            lsTmp = lsTmp + Mid(Trim(lsFormula), i, 1)
        Else
            If Len(lsTmp) > 0 Then
                nCtaCont = nCtaCont + 1
                ReDim Preserve MatDatos(nCtaCont)
                
                MatDatos(nCtaCont).CuentaContable = lsTmp
                
                If MatDatos(nCtaCont).CuentaContable = "100" Or MatDatos(nCtaCont).CuentaContable = "1000" Then
                    MatDatos(nCtaCont).Saldo = MatDatos(nCtaCont).CuentaContable
                Else
                    If Trim(psAgencia) = "" Then
                        MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual2(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True)
                    Else
                        MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensualxAgencia(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True, psAgencia)
                    End If
                End If
                
                If nCtaCont > 1 Then
                    If Mid(Trim(lsFormula), i, 1) = ")" Then
                        nPosicion = 0
                    Else
                        nPosicion = i
                    End If
                End If
                    If sTempAD = "" Then
                        If nCtaCont = 1 Then
                            If ((i - Len(Trim(lsTmp))) - 3) > 1 Then
                                sTempAD = Mid(Trim(lsFormula), (i - Len(Trim(lsTmp))) - 3, 2)
                            Else
                                sTempAD = ""
                            End If
                        Else
                            sTempAD = Mid(Trim(lsFormula), (i - Len(MatDatos(nCtaCont).CuentaContable)) - 3, 2)
                        End If
                    End If
                
                If sTempAD = "SA" Or sTempAD = "SD" Then
                    MatDatos(nCtaCont).CuentaContable = DepuraSaldoAD(MatDatos(nCtaCont).CuentaContable)
                    If sTempAD = "SA" Then
                        MatDatos(nCtaCont).bSaldoA = True
                        MatDatos(nCtaCont).bSaldoD = False
                    Else
                        MatDatos(nCtaCont).bSaldoA = False
                        MatDatos(nCtaCont).bSaldoD = True
                    End If
                    Else
                        MatDatos(nCtaCont).bSaldoA = False
                        MatDatos(nCtaCont).bSaldoD = False
                End If
            End If
            If nPosicion = 0 Then
               sTempAD = ""
            End If
            lsTmp = ""
        End If
    Next i
    If Len(lsTmp) > 0 Then
        nCtaCont = nCtaCont + 1
        ReDim Preserve MatDatos(nCtaCont)
        MatDatos(nCtaCont).CuentaContable = lsTmp
        If MatDatos(nCtaCont).CuentaContable = "100" Or MatDatos(nCtaCont).CuentaContable = "1000" Then
            MatDatos(nCtaCont).Saldo = MatDatos(nCtaCont).CuentaContable
        Else
            If Trim(psAgencia) = "" Then
                MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual2(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True)
            Else
                MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensualxAgencia(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True, psAgencia)
            End If
        End If
    End If
    lsTmp = ""
    lsCadFormula = ""
    Dim nEncontrado As Integer
    For i = 1 To Len(lsFormula)
        If (Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Then
            lsTmp = lsTmp + Mid(Trim(lsFormula), i, 1)
        Else
            If Len(lsTmp) > 0 Then
                For j = 1 To nCtaCont
                    nEncontrado = 0
                    If MatDatos(j).CuentaContable = lsTmp Then
                            
                            If MatDatos(j).bSaldoA = True Or MatDatos(j).bSaldoD = True Then
                                MatDatos(j).Saldo = oNBal.CalculaSaldoBECuentaAD(MatDatos(j).CuentaContable, pnMoneda, MatDatos(j).bSaldoA, CStr(pnMoneda), Trim(psAgencia), Format(pdFecha, "YYYY"), Format(pdFecha, "MM"))
                                nEncontrado = 1
                            End If
                                If Left(Format(MatDatos(j).Saldo, "#0.00"), 1) = "-" And (Right(lsCadFormula, 1) = "-" Or Right(lsCadFormula, 1) = "+") Then
                                    
                                    If Right(Trim(lsCadFormula), 1) = "-" Or Right(Trim(lsCadFormula), 1) = "+" Then
                                        If Right(Trim(lsCadFormula), 1) = "-" And MatDatos(j).Saldo >= 0 Then
                                            LsSigno = "-"
                                        ElseIf Right(Trim(lsCadFormula), 1) = "-" And MatDatos(j).Saldo < 0 Then
                                            LsSigno = "+"
                                        ElseIf Right(Trim(lsCadFormula), 1) = "+" And MatDatos(j).Saldo >= 0 Then
                                            LsSigno = "+"
                                        ElseIf Right(Trim(lsCadFormula), 1) = "+" And MatDatos(j).Saldo < 0 Then
                                            LsSigno = "-"
                                        End If
                                    Else
                                        LsSigno = ""
                                    End If
                                    If LsSigno = "" Then
                                        lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & Format(MatDatos(j).Saldo, "#0.00")
                                    Else
                                        lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & LsSigno & Format(Abs(MatDatos(j).Saldo), "#0.00")
                                    End If
                                    nEncontrado = 1
                                Else
                                    lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
                                    nEncontrado = 1
                                End If
                        Exit For
                    End If
                Next j
            End If
            lsTmp = ""
            If nEncontrado = 1 Or (Mid(Trim(lsFormula), i, 1) = "S" Or Mid(Trim(lsFormula), i, 1) = "A" Or Mid(Trim(lsFormula), i, 1) = "D") Then
            lsCadFormula = lsCadFormula & Mid(Trim(lsFormula), i, 1)
            Else
            lsCadFormula = lsCadFormula & "" & Mid(Trim(lsFormula), i, 1)
            End If
        End If
    Next
    If Len(lsTmp) > 0 Then
        For j = 1 To nCtaCont
           If MatDatos(j).CuentaContable = lsTmp Then
               If Left(Format(MatDatos(j).Saldo, "#0.00"), 1) = "-" And (Right(lsCadFormula, 1) = "-" Or Right(lsCadFormula, 1) = "+") Then
                    lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & Format(MatDatos(j).Saldo, "#0.00")
                    nEncontrado = 1
                Else
                    lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
                    nEncontrado = 1
                End If
               Exit For
           End If
        Next j
    End If
    lsCadFormula = Replace(Replace(lsCadFormula, "SA", ""), "SD", "")
    ObtenerResultadoFormula = oFormula.ExprANum(lsCadFormula)
    Set oBal = Nothing
    Set oFormula = Nothing
End Function
