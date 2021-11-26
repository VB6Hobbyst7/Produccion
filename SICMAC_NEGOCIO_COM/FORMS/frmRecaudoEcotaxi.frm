VERSION 5.00
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form frmRecaudoEcotaxi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DEPOSITO A CTA X RECAUDO ECOTAXI (Formato IV)"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   Icon            =   "frmrecaudoecotaxi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   12000
   Begin VB.CommandButton cmdListado 
      Caption         =   "3. Listado                 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   4275
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "ECOTAXI  "
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
      Height          =   4095
      Left            =   80
      TabIndex        =   3
      Top             =   40
      Width           =   11895
      Begin SICMACT.FlexEdit feRecaudo 
         Height          =   3330
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   5874
         Cols0           =   13
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   $"frmrecaudoecotaxi.frx":030A
         EncabezadosAnchos=   "350-3000-1800-1000-1500-1200-1500-1200-1800-0-0-0-0"
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
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-C-C-C-R-L-C-C-R-L-R-C"
         FormatosEdit    =   "0-0-0-5-2-0-0-0-0-3-0-2-2"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin VB.Label lblNumRegistros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   300
         Left            =   1800
         TabIndex        =   7
         Top             =   3720
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total de Registros:"
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
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   3765
         Width           =   1635
      End
   End
   Begin VB.CommandButton CmdCargaArch 
      Caption         =   "1. Cargar Archivo al Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   80
      TabIndex        =   2
      Top             =   4275
      Width           =   2745
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "2. Depositar Recaudos"
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
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   4275
      Width           =   2415
   End
   Begin VB.CommandButton CmdSalir 
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
      Height          =   375
      Left            =   10680
      TabIndex        =   0
      Top             =   4275
      Width           =   1275
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgFile 
      Height          =   30
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   700
      _ExtentX        =   1244
      _ExtentY        =   53
      Filtro          =   "Archivos de Texto (*.pagos)|*.pagos|Archivos de Texto (*.cobros)|*.cobros"
      Altura          =   0
   End
End
Attribute VB_Name = "frmRecaudoEcotaxi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oBarra As clsProgressBar
Dim fsNomFile As String
Dim fsPathFile As String
Dim fsruta As String
Dim ldFechaCargaG As Date
Dim i As Integer
Dim f As Integer

Private Sub cmdListado_Click()
    Dim objCredito As COMDCredito.DCOMCredito
    Dim oRsCredito As ADODB.Recordset
    Set oRsCredito = New ADODB.Recordset
    Set objCredito = New COMDCredito.DCOMCredito
    'ldFechaCargaG = "2012/12/12"
    Set oRsCredito = objCredito.RecuperaRecaudosParaReporte(ldFechaCargaG)
    Call ReporteRecaudo(ldFechaCargaG, oRsCredito)
    Set objCredito = Nothing
'    RSVacio (oRsCredito)
End Sub
Private Sub ReporteRecaudo(ByVal pdFecha As Date, ByVal poRs As ADODB.Recordset)
    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja  As String
    Dim lsArchivo As String
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
'    Dim sFecha As String

'On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ReAbReEco"
   
    lsNomHoja = "ReAbReEco"
   
    lsArchivo1 = "\spooler\Reporte_2A1_" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
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
    
    If Not (poRs.BOF Or poRs.EOF) Then
    xlHoja1.Cells(1, 7) = pdFecha
    xlHoja1.Range("E:E").NumberFormat = "#,##0.00"
    Do While Not poRs.EOF
        xlHoja1.Range(xlHoja1.Cells(5 + poRs!nId, 1), xlHoja1.Cells(5 + poRs!nId, 7)).Borders.LineStyle = 1
        xlHoja1.Cells(5 + poRs!nId, 1) = poRs!nId
        xlHoja1.Cells(5 + poRs!nId, 2) = poRs!cPersCod
        xlHoja1.Cells(5 + poRs!nId, 3) = poRs!cPersNombre
        xlHoja1.Cells(5 + poRs!nId, 4) = poRs!cCtaCodAbono
        xlHoja1.Cells(5 + poRs!nId, 5) = CDbl(poRs!nIFIRecaudoNeto) + CDbl(poRs!nCOFIDEPorcentajeComision)
        xlHoja1.Cells(5 + poRs!nId, 6) = Format(poRs!dFechaCarga, "YYYY/MM/DD hh:mm:ss")
        xlHoja1.Cells(5 + poRs!nId, 7) = poRs!cUser
        poRs.MoveNext
    Loop
    End If
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
End Sub

Private Sub Form_Load()
    CentraForm Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End Sub
Public Sub Inicio()
    If SeSubioArchivoRecaudo(gdFecSis) Then
        If MostrarRecaudos(gdFecSis) Then
            MsgBox "Quedan pendientes por Depositar los siguientes Recaudos", vbInformation, "Aviso"
            CmdCargaArch.Enabled = False
            cmdGrabar.Enabled = True
            cmdListado.Enabled = True 'ALPA 20130114
        Else
            MsgBox "El Día de hoy ya se realizó el Déposito x Recaudo EcoTaxi", vbInformation, "Aviso"
            Unload Me
            Exit Sub
        End If
    Else
        CmdCargaArch.Enabled = True
        cmdGrabar.Enabled = False
        cmdListado.Enabled = False 'ALPA 20130114
    End If
    Me.Show
End Sub
Private Sub CmdCargaArch_Click()
    If MsgBox("Recuerde que solo una vez al día podrá cargar la información de los recaudos" & Chr(10) & "¿Desea cargar los recaudos en el Sistema?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If
    CdlgFile.nHwd = Me.hwnd
    CdlgFile.Filtro = "Archivos (*.txt)|*.txt"
    CdlgFile.altura = 300
    CdlgFile.TipoVentana = Normal
    CdlgFile.Show
    
    fsPathFile = CdlgFile.Ruta
    fsruta = fsPathFile
    If fsPathFile <> Empty Then
        cmdGrabar.Enabled = False
        cmdListado.Enabled = False 'ALPA 20130114
        Screen.MousePointer = 11
        Leer_Lineas (fsruta)
    Else
        MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
        Exit Sub
    End If
    Screen.MousePointer = 0
End Sub
Public Sub Leer_Lineas(ByVal strTextFile As String)
    Dim oBase As COMDCredito.DCOMCredActBD
    Dim oCap As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsCap As ADODB.Recordset, rsRecaudo As ADODB.Recordset
    Dim Datos() As String, lsCtaCodAbono As String, str_Linea As String
    Dim Linea As Long, lnNroRecaudos As Long
    Dim ldFechaCarga As Date, ldFechaRecaudo As Date
    Dim lsMsgRecaudosRepetidos As String
    Dim vPrevio As previo.clsprevio
    Dim lnRecaudoId As Long
    Dim oFun As COMFunciones.FCOMImpresion

    f = FreeFile

    On Error GoTo ErrLeerLineas
    FormatearGrillaRecaudo
    Open strTextFile For Input As #f

    Set oCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set oBase = New COMDCredito.DCOMCredActBD
    Set oFun = New COMFunciones.FCOMImpresion
    
    ldFechaCarga = CDate(gdFecSis & " " & Time)
    lsMsgRecaudosRepetidos = ""
    lnNroRecaudos = 0

    'Limpia Detalle Temporal
    oBase.dLimpiaRecaudoDetTemp
    Linea = 0

    'Inserta Detalle de Recaudo Temporal
    Do
        Line Input #f, str_Linea
        Datos = Split(str_Linea, "|")
        Linea = Linea + 1
        Set rsCap = New ADODB.Recordset

        If UBound(Datos) = 14 Then
            lsCtaCodAbono = ""
            Set rsRecaudo = New ADODB.Recordset
            Set rsCap = oCap.GetCuentasPersona(Datos(12), gCapAhorros, True, False, CInt(Mid(Datos(11), 9, 1)), , , 7)
            
            If Not RSVacio(rsCap) Then
                lsCtaCodAbono = IIf(IsNull(rsCap!cCtaCod), "", rsCap!cCtaCod)
            End If

            If lsCtaCodAbono = "" Then
                MsgBox "La Cta Ecotaxi Nro " & Datos(11) & " no tiene Cta de Ahorro para Recaudo, Se tiene que vincular al Crédito", vbOKOnly + vbCritical, "Aviso"
                Exit Sub
            End If
            ldFechaRecaudo = CDate(Datos(1)) & " " & Datos(2)
            Set rsRecaudo = oBase.DameRecaudo(Datos(11), Datos(14), Datos(4), ldFechaRecaudo)
            'Para que no se repitan los recaudos con dias anteriores
            If RSVacio(rsRecaudo) Then
                Call oBase.dRegistrarRecaudoDetTemp(Datos(0), ldFechaRecaudo, Datos(3), Datos(4), CCur(Datos(5)), CCur(Datos(6)), CCur(Datos(7)), CCur(Datos(8)), CCur(Datos(9)), CCur(Datos(10)), Datos(11), lsCtaCodAbono, Datos(12), Datos(13), Datos(14), True)
                lnNroRecaudos = lnNroRecaudos + 1
            Else
                Call oBase.dRegistrarRecaudoDetTemp(Datos(0), ldFechaRecaudo, Datos(3), Datos(4), CCur(Datos(5)), CCur(Datos(6)), CCur(Datos(7)), CCur(Datos(8)), CCur(Datos(9)), CCur(Datos(10)), Datos(11), lsCtaCodAbono, Datos(12), Datos(13), Datos(14), False)
                lsMsgRecaudosRepetidos = lsMsgRecaudosRepetidos & DevolverCadenaRecaudosRepetidos(Linea, Datos, rsRecaudo)
            End If
        Else
            MsgBox "La Linea N° " & Format(Linea, "0000") & " no tiene la estructura correcta, comunicar al Dpto de TI", vbOKOnly + vbCritical, "Aviso"
            Close #f
            Exit Sub
        End If
    Loop While Not EOF(f)
    Close #f
    
    If Len(lsMsgRecaudosRepetidos) > 0 Then
        lsMsgRecaudosRepetidos = oFun.CabeceraPagina("Recaudos EcoTaxi Repetidos", 0, 1, gsNomAge, gsInstCmac, gdFecSis, , False) & Chr(10) & lsMsgRecaudosRepetidos
        MsgBox "Se han encontrado Recaudos con Duplicidad, favor verifique", vbExclamation, "Aviso"
        Set vPrevio = New previo.clsprevio
        vPrevio.Show lsMsgRecaudosRepetidos, "Recaudos EcoTaxi Repetidos", True
        If lnNroRecaudos > 0 Then
            If MsgBox("¿Desea cargar los Recaudos que no se repiten al Sistema el día de Hoy?" & Chr(10) & "Recuerde que solo una vez al día puede subir este archivo", vbYesNo + vbInformation, "Aviso") = vbNo Then
                Unload Me
                Exit Sub
            End If
        Else
            MsgBox "No se va a poder subir el Archivo al Sistema ya que todos los Recaudos están duplicados", vbExclamation, "Aviso"
            Unload Me
            Exit Sub
        End If
    End If

    'Migra el Detalle Temporal
    oBase.dMigraRecaudosDet (ldFechaCarga)
    ldFechaCargaG = ldFechaCarga
    CmdCargaArch.Enabled = False
    cmdGrabar.Enabled = False
    cmdListado.Enabled = False 'ALPA 20130114
    If MostrarRecaudos(gdFecSis) Then
        cmdGrabar.Enabled = True
        cmdListado.Enabled = True 'ALPA 20130114
    End If

    Set oBase = Nothing
    Set oCap = Nothing
    Set oFun = Nothing
    Set vPrevio = Nothing
    Exit Sub
ErrLeerLineas:
    CmdCargaArch.Enabled = False
    cmdGrabar.Enabled = False
    cmdListado.Enabled = False 'ALPA 20130114
    MsgBox TextErr(err.Description), vbCritical, "Aviso"
End Sub
Private Function MostrarRecaudos(ByVal pdFecha As Date) As Boolean
    Dim oCred As COMDCredito.DCOMCredito
    Dim rsRecaudo As ADODB.Recordset
    
    Set oCred = New COMDCredito.DCOMCredito
    Set rsRecaudo = New ADODB.Recordset
    
    Set rsRecaudo = oCred.RecuperaRecaudosParaAbono(pdFecha)
    FormatearGrillaRecaudo
    If Not RSVacio(rsRecaudo) Then
        Do While Not rsRecaudo.EOF
            feRecaudo.AdicionaFila
            feRecaudo.TextMatrix(feRecaudo.row, 1) = rsRecaudo!cPersNombre 'NOM CLIENTE
            feRecaudo.TextMatrix(feRecaudo.row, 2) = rsRecaudo!cCtaCodCredito 'COD SOLICITUD
            feRecaudo.TextMatrix(feRecaudo.row, 3) = rsRecaudo!cPlacaNum 'PLACA
            feRecaudo.TextMatrix(feRecaudo.row, 4) = Format(rsRecaudo!dEESSFechaRecaudo, "dd/mm/yyyy hh:mm AMPM") 'FECHA RECAUDO
            'feRecaudo.TextMatrix(feRecaudo.row, 5) = Format(rsRecaudo!nIFIRecaudoNeto, "##,##0.00") 'RECAUDO IFI
            feRecaudo.TextMatrix(feRecaudo.row, 5) = Format(rsRecaudo!nIFIRecaudoNeto + rsRecaudo!nCOFIDEPorcentajeComision, "##,##0.00") 'Recaudo IFI + Comision COFIDE = Recaudo Bruto
            feRecaudo.TextMatrix(feRecaudo.row, 6) = rsRecaudo!cEESSNombre 'NOMBRE EESS
            feRecaudo.TextMatrix(feRecaudo.row, 7) = rsRecaudo!cEESSNroTicket 'TICKET EESS
            feRecaudo.TextMatrix(feRecaudo.row, 8) = rsRecaudo!cCtaCodAbono 'CTA ABONO
            feRecaudo.TextMatrix(feRecaudo.row, 9) = rsRecaudo!nId 'RECAUDO ID
            feRecaudo.TextMatrix(feRecaudo.row, 10) = rsRecaudo!nMovNro 'nMovNro
            feRecaudo.TextMatrix(feRecaudo.row, 11) = Format(rsRecaudo!nIFIRecaudoNeto, "##,##0.00") 'Recaudo Neto
            feRecaudo.TextMatrix(feRecaudo.row, 12) = Format(rsRecaudo!nCOFIDEPorcentajeComision, "##,##0.00") 'Comision COFIDE
            rsRecaudo.MoveNext
        Loop
        lblNumRegistros.Caption = rsRecaudo.RecordCount
        MostrarRecaudos = True
    Else
        lblNumRegistros.Caption = 0
        MostrarRecaudos = False
    End If
    Set oCred = Nothing
    Set rsRecaudo = Nothing
End Function
Private Sub FormatearGrillaRecaudo()
    feRecaudo.Clear
    feRecaudo.FormaCabecera
    feRecaudo.Rows = 2
End Sub
Private Sub CmdGrabar_Click()
    Dim oBase As COMDCredito.DCOMCredActBD
    Dim oITF As COMDConstSistema.FCOMITF
    Dim sMovNro As String
    Dim bTransac As Boolean
    Dim pMatDatosAhoAbo As Variant
    Dim nITFAbono As Double, nRedondeoITF As Double
    Dim nMovNro As Long, tam As Long
    Dim i As Integer
    
    Dim lsCtaCredito As String
    Dim lsCtaAhorro As String
    Dim lnMontoAbono As Double
    Dim lnId As Long
    Dim lnMonto As Double, lnComision As Double

    If MsgBox("¿Esta seguro de depositar los recuados en las Ctas de Ahorro de Ecotaxi?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If
    
On Error GoTo ErrAbonoRecaudos
    'Inicializa Datos de Ahorros
    ReDim pMatDatosAhoAbo(14)
    pMatDatosAhoAbo(0) = "" 'Cuenta de Ahorros
    pMatDatosAhoAbo(1) = "0.00" 'Monto de Apertura
    pMatDatosAhoAbo(2) = "0.00" 'Interes Ganado de Abono
    pMatDatosAhoAbo(3) = "0.00" 'Interes Ganado de Retiro Gastos
    pMatDatosAhoAbo(4) = "0.00" 'Interes Ganado de Retiro Cancelaciones
    pMatDatosAhoAbo(5) = "0.00" 'Monto de Abono
    pMatDatosAhoAbo(6) = "0.00" 'Monto de Retiro de Gastos
    pMatDatosAhoAbo(7) = "0.00" 'Monto de Retiro de Cancelaciones
    pMatDatosAhoAbo(8) = "0.00" 'Saldo Disponible Abono
    pMatDatosAhoAbo(9) = "0.00" 'Saldo Contable Abono
    pMatDatosAhoAbo(10) = "0.00" 'Saldo Disponible Retiro de Gastos
    pMatDatosAhoAbo(11) = "0.00" 'Saldo Contable Retiro de Gastos
    pMatDatosAhoAbo(12) = "0.00" 'Saldo Disponible Retiro de Cancelaciones
    pMatDatosAhoAbo(13) = "0.00" 'Saldo Contable Retiro de Cancelaciones

    Set oBarra = New clsProgressBar
    Set oITF = New COMDConstSistema.FCOMITF
    
    tam = feRecaudo.Rows - 1
    oBarra.ShowForm Me
    oBarra.CaptionSyle = eCap_CaptionPercent
    oBarra.Max = tam
    oBarra.Progress 0, "Proceso de Depósito a Cta x Recaudo EcoTaxi", "Preparando Abono...", "Recaudo Ecotaxi", vbBlue
    
    oITF.fgITFParametros

    For i = 1 To feRecaudo.Rows - 1
        sMovNro = ""
        nMovNro = 0
        nITFAbono = 0
        nRedondeoITF = 0

        lsCtaCredito = CStr(feRecaudo.TextMatrix(i, 2))
        lnId = CLng(feRecaudo.TextMatrix(i, 9))
'        lnRecaudo = lnId
        lsCtaAhorro = CStr(feRecaudo.TextMatrix(i, 8))
        'lnMontoAbono = CDbl(feRecaudo.TextMatrix(i, 5))
        lnMonto = CDbl(feRecaudo.TextMatrix(i, 11))
        lnComision = CDbl(feRecaudo.TextMatrix(i, 12))
        lnMontoAbono = lnMonto + lnComision
        
        Set oBase = New COMDCredito.DCOMCredActBD
        bTransac = False
        Call oBase.dBeginTrans
        bTransac = True

        sMovNro = oBase.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Call oBase.InsertaMov(sMovNro, gAhoDepCtaRecaudoEcotaxi, "Depósito x Recaudo Cta EcoTaxi Nro. " & lsCtaCredito, gMovEstContabMovContable, gMovFlagVigente)
        nMovNro = oBase.GetnMovNro(sMovNro)
        

        If sMovNro <> "" Or nMovNro > 0 Then
            'Abona a Ctas de Ahorro de los Créditos Ecotaxi
            nITFAbono = oITF.fgTruncar(oITF.fgITFCalculaImpuesto(lnMontoAbono), 2)
            nRedondeoITF = fgDiferenciaRedondeoITF(nITFAbono)
            nITFAbono = IIf(nRedondeoITF > 0, nITFAbono - nRedondeoITF, nITFAbono)

            oBase.CapAbonoCuentaAho pMatDatosAhoAbo, lsCtaAhorro, lnMontoAbono, gAhoDepCtaRecaudoEcotaxi, sMovNro, "Depósito x Recaudo a la Cta Ecotaxi Nro. " & lsCtaCredito, , , , , , , gdFecSis, "", True, nITFAbono, False, gITFCobroCargo, lnComision

            If nITFAbono + nRedondeoITF > 0 Then
               Call oBase.InsertaMovRedondeoITF(sMovNro, 1, nITFAbono + nRedondeoITF, nITFAbono)
            End If
            'Actualiza nMovNro en RecaudoEcoTaxiDet
            Call oBase.dAsignaNroMovDepEcoTaxi(lnId, nMovNro)

            oBarra.Progress i, "Proceso de Depósito a Cta x Recaudo EcoTaxi", "Efectuando Depósito Nro " & i & " Cuenta: " & lsCtaAhorro, "Recaudo EcoTaxi", vbBlue
            Call oBase.dCommitTrans
            bTransac = False
        Else
            Call oBase.dRollbackTrans
        End If
    Next

    oBarra.CloseForm Me
    
    Set oBase = Nothing
    Set oITF = Nothing
    Set oBarra = Nothing
    
    Call cmdListado_Click 'ALPA 20130114
    MsgBox "Se ha realizado el Depósito a las Ctas de Ahorro de los Créditos Ecotaxi con éxito", vbInformation, "Aviso"
    Unload Me
    Exit Sub
ErrAbonoRecaudos:
    If bTransac Then
        Call oBase.dRollbackTrans
        Set oBase = Nothing
    End If
    err.Raise err.Number, "Error En Proceso", err.Description
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Function SeSubioArchivoRecaudo(ByVal pdFecha As Date) As Boolean
    Dim oCredito As COMDCredito.DCOMCredito
    Set oCredito = New COMDCredito.DCOMCredito
    SeSubioArchivoRecaudo = oCredito.SeSubioArchivoRecaudo(pdFecha)
    Set oCredito = Nothing
End Function
Private Function DevolverCadenaRecaudosRepetidos(ByVal NroLinea As Long, ByVal Recaudo As Variant, ByVal rsRecaudosRepetido As ADODB.Recordset) As String
    Dim Cadena As String
    Dim lsCab As String * 50
    Dim lsIdentificador As String * 50
    Dim lsNroCredito As String * 50
    Dim lsCliente As String * 50
    Dim lsFechaRecaudo As String * 50
    Dim lsMontoRecaudo As String * 50
    Dim lsFechaCarga As String * 50

    Do While Not rsRecaudosRepetido.EOF
        lsCab = Space(2) & "Archivo Recibido"
        Cadena = Cadena & lsCab
        lsCab = "Archivo Repetido"
        Cadena = Cadena & lsCab & Chr(10)
        lsCab = Space(2) & String(16, "-")
        Cadena = Cadena & lsCab
        lsCab = String(16, "-")
        Cadena = Cadena & lsCab & Chr(10)
        lsIdentificador = Space(2) & "Nro Linea: " & NroLinea
        Cadena = Cadena & lsIdentificador
        lsIdentificador = "Id Recaudo: " & rsRecaudosRepetido!nId
        Cadena = Cadena & lsIdentificador & Chr(10)
        lsNroCredito = Space(2) & "Nro Crédito: " & Recaudo(11)
        Cadena = Cadena & lsNroCredito
        lsNroCredito = "Nro Crédito: " & rsRecaudosRepetido!cCtaCodCredito
        Cadena = Cadena & lsNroCredito & Chr(10)
        lsFechaRecaudo = Space(2) & "Fecha Recaudo: " & Format(CDate(Recaudo(1)) & " " & Recaudo(2), "dd/mm/yyyy hh:mm:ss AMPM")
        Cadena = Cadena & lsFechaRecaudo
        lsFechaRecaudo = "Fecha de Recaudo: " & Format(rsRecaudosRepetido!dEESSFechaRecaudo, "dd/mm/yyyy hh:mm:ss AMPM")
        Cadena = Cadena & lsFechaRecaudo & Chr(10)
        lsMontoRecaudo = Space(2) & "Monto Recaudo: " & Format(Recaudo(10), "##,##0.00")
        Cadena = Cadena & lsMontoRecaudo
        lsMontoRecaudo = "Monto Recaudo: " & Format(rsRecaudosRepetido!nIFIRecaudoNeto, "##,##0.00")
        Cadena = Cadena & lsMontoRecaudo & Chr(10)
        lsFechaCarga = "Fecha Carga Sistema: " & Format(rsRecaudosRepetido!dFechaCarga, "dd/mm/yyyy hh:mm:ss AMPM")
        Cadena = Cadena & Space(50) & lsFechaCarga & Chr(10) & Chr(10)
        rsRecaudosRepetido.MoveNext
    Loop
    DevolverCadenaRecaudosRepetidos = Cadena
End Function
