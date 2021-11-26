VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmINFOGASLecturaArchivos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   Icon            =   "frmINFOGASLecturaArchivos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   9240
   Begin MSComDlg.CommonDialog CdlgFile2 
      Left            =   3600
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCargaArch 
      Caption         =   "&Cargar Archivo"
      Height          =   375
      Left            =   80
      TabIndex        =   6
      Top             =   3315
      Width           =   1545
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   3315
      Width           =   1755
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   3315
      Width           =   1035
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
      Height          =   3255
      Left            =   80
      TabIndex        =   0
      Top             =   40
      Width           =   9135
      Begin SICMACT.FlexEdit feRecaudo 
         Height          =   2610
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   4604
         Cols0           =   16
         FixedCols       =   2
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   $"frmINFOGASLecturaArchivos.frx":030A
         EncabezadosAnchos=   "350-3000-1800-1000-1500-1200-1500-1200-0-0-0-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-C-C-C-R-L-C-C-C-C-C-R-C-C-C"
         FormatosEdit    =   "0-0-0-5-2-0-0-0-0-0-0-2-2-2-2-2"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         CellBackColor   =   -2147483633
      End
      Begin SICMACT.FlexEdit feCreditosAprobados 
         Height          =   2610
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   4604
         Cols0           =   8
         FixedCols       =   2
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Cliente-N° Cuenta-Placa-Taller-Fecha-CodIFI-Cod Cliente"
         EncabezadosAnchos=   "350-2500-2500-1000-900-1500-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-5-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         CellBackColor   =   -2147483633
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
         TabIndex        =   8
         Top             =   2925
         Width           =   1635
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
         Top             =   2880
         Width           =   825
      End
   End
   Begin VB.PictureBox CdlgFile 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   645
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   700
   End
End
Attribute VB_Name = "frmINFOGASLecturaArchivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim oBarra As clsProgressBar
Dim fsNomFile As String
Dim fsPathFile As String
Dim fsruta As String
Dim i As Integer
Dim f As Integer
Dim fsOpeCod As String

Private Sub Form_Load()
    CentraForm Me
End Sub

Public Sub Inicio(ByVal psOpeCod As String, ByVal psOpeDescripcion As String)
    fsOpeCod = psOpeCod
    Me.Caption = psOpeDescripcion
    Select Case fsOpeCod
        Case "05000"
            feCreditosAprobados.Visible = True
            feRecaudo.Visible = False
            FormatearGrillaAprobacion
            cmdGrabar.Caption = "&Grabar Activación"
        'Case gAhoDepCtaRecaudoEcotaxi
        '    If SeRealizoAbonoRecaudo(gdFecSis) Then
        '        MsgBox "Solo se puede realizar un Abono de Recaudo en el día", vbInformation, "Aviso"
        '        Unload Me
        '        Exit Sub
        '    End If
        '    feCreditosAprobados.Visible = False
        '    feRecaudo.Visible = True
        '    FormatearGrillaRecaudo
        '    cmdGrabar.Caption = "&Grabar Recaudos"
    End Select
    Me.Show 0
End Sub
Private Sub CmdCargaArch_Click()
    'CdlgFile.nHwd = Me.hwnd
'    CdlgFile.Filtro = "Archivos (*.txt)|*.txt"
'    CdlgFile.altura = 300
'    CdlgFile.TipoVentana = Normal
'    CdlgFile.Show
    
    'CdlgFile2.hwnd = Me.hwnd
    CdlgFile2.InitDir = "C:\"
    CdlgFile2.Filter = "Archivos (*.txt)|*.txt"
    'CdlgFile2. = 300
    'CdlgFile. = Normal
    CdlgFile2.ShowOpen
    
    
    fsPathFile = CdlgFile2.Filename
    fsruta = fsPathFile
    If fsPathFile <> Empty Then
        cmdGrabar.Enabled = False
        Screen.MousePointer = 11
        Leer_Lineas (fsruta)
    Else
        MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
        Exit Sub
    End If
    Screen.MousePointer = 0
End Sub

Public Sub Leer_Lineas(ByVal strTextFile As String)
    Dim str_Linea As String
    Dim oPersona As New COMDPersona.DCOMPersonas
    Dim rsPersona As New ADODB.Recordset
    Dim Datos() As String
    Dim Linea As Long

    Linea = 0
    f = FreeFile
    
    On Error GoTo ErrLeerLineas
    Select Case fsOpeCod
        Case "05000"
            FormatearGrillaAprobacion
            Open strTextFile For Input As #f
            Do
                Line Input #f, str_Linea
                Linea = Linea + 1
                Datos = Split(str_Linea, "|")
                If UBound(Datos) = 5 Then
                    Set rsPersona = oPersona.BuscaCliente(Datos(2), BusquedaCodigo)
                    feCreditosAprobados.AdicionaFila
                    feCreditosAprobados.TextMatrix(feCreditosAprobados.Row, 1) = PstaNombre(rsPersona!cPersNombre)  'NOM CLIENTE
                    feCreditosAprobados.TextMatrix(feCreditosAprobados.Row, 2) = Datos(0)  'SOLICITUD
                    feCreditosAprobados.TextMatrix(feCreditosAprobados.Row, 3) = Datos(1)  'PLACA
                    feCreditosAprobados.TextMatrix(feCreditosAprobados.Row, 4) = Datos(5)  'TALLER
                    feCreditosAprobados.TextMatrix(feCreditosAprobados.Row, 5) = Datos(3)  'FECHA GENERACION
                    feCreditosAprobados.TextMatrix(feCreditosAprobados.Row, 6) = Datos(4)  'IFI
                    feCreditosAprobados.TextMatrix(feCreditosAprobados.Row, 7) = Datos(2)  'COD CLIENTE
                Else
                    MsgBox "La Linea N° " & Format(Linea, "0000") & " no tiene la estructura correcta, comunicar al Dpto de TI", vbOKOnly + vbCritical, "Aviso"
                    lblNumRegistros.Caption = 0
                    FormatearGrillaAprobacion
                    cmdGrabar.Enabled = False
                    Close #f
                    Exit Sub
                End If
            Loop While Not EOF(f)
            Close #f
            lblNumRegistros.Caption = Linea
            cmdGrabar.Enabled = True
        Case gAhoDepCtaRecaudoEcotaxi
            FormatearGrillaRecaudo
            Open strTextFile For Input As #f
            Do
                Line Input #f, str_Linea
                Linea = Linea + 1
                Datos = Split(str_Linea, "|")
                If UBound(Datos) = 14 Then
                    Set rsPersona = oPersona.BuscaCliente(Datos(12), BusquedaCodigo)
                    feRecaudo.AdicionaFila
                    feRecaudo.TextMatrix(feRecaudo.Row, 1) = PstaNombre(rsPersona!cPersNombre) 'NOM CLIENTE
                    feRecaudo.TextMatrix(feRecaudo.Row, 2) = Datos(11) 'COD SOLICITUD
                    feRecaudo.TextMatrix(feRecaudo.Row, 3) = Datos(0) 'PLACA
                    feRecaudo.TextMatrix(feRecaudo.Row, 4) = Datos(1) & " " & Datos(2) 'FECHA RECAUDO
                    feRecaudo.TextMatrix(feRecaudo.Row, 5) = Format(CCur(Datos(10)), "##,##0.00") 'RECAUDO IFI
                    feRecaudo.TextMatrix(feRecaudo.Row, 6) = Datos(3) 'NOMBRE EESS
                    feRecaudo.TextMatrix(feRecaudo.Row, 7) = Datos(4) 'TICKET EESS
                    feRecaudo.TextMatrix(feRecaudo.Row, 8) = Datos(12) 'COD CLIENTE
                    feRecaudo.TextMatrix(feRecaudo.Row, 9) = Datos(14) 'COD EESS
                    feRecaudo.TextMatrix(feRecaudo.Row, 10) = Datos(13) 'COD FINANCIERA
                    feRecaudo.TextMatrix(feRecaudo.Row, 11) = Format(CCur(Datos(5)), "##,##0.00") 'RECAUDO BRUTO
                    feRecaudo.TextMatrix(feRecaudo.Row, 12) = Format(CCur(Datos(8)), "##,##0.00") 'RECAUDO REAL
                    feRecaudo.TextMatrix(feRecaudo.Row, 13) = Format(CCur(Datos(9)), "##,##0.00") '% COMISION COFIDE
                    feRecaudo.TextMatrix(feRecaudo.Row, 14) = Format(CCur(Datos(6)), "##,##0.00") 'ITF ENTRADA
                    feRecaudo.TextMatrix(feRecaudo.Row, 15) = Format(CCur(Datos(7)), "##,##0.00") 'ITF SALIDA
                Else
                    MsgBox "La Linea N° " & Format(Linea, "0000") & " no tiene la estructura correcta, comunicar al Dpto de TI", vbOKOnly + vbCritical, "Aviso"
                    lblNumRegistros.Caption = 0
                    FormatearGrillaRecaudo
                    cmdGrabar.Enabled = False
                    Close #f
                    Exit Sub
                End If
            Loop While Not EOF(f)
            Close #f
            lblNumRegistros.Caption = Linea
            cmdGrabar.Enabled = True
        Case Else
            MsgBox "La Operación Actual No existe", vbInformation, "Aviso"
            Exit Sub
    End Select
    Exit Sub
ErrLeerLineas:
    MsgBox TextErr(err.Description), vbCritical, "Aviso"
    lblNumRegistros.Caption = 0
    cmdGrabar.Enabled = False
End Sub
Private Sub FormatearGrillaAprobacion()
    feCreditosAprobados.Clear
    feCreditosAprobados.FormaCabecera
    feCreditosAprobados.Rows = 2
End Sub
Private Sub FormatearGrillaRecaudo()
    feRecaudo.Clear
    feRecaudo.FormaCabecera
    feRecaudo.Rows = 2
End Sub
Private Sub cmdGrabar_Click()
    Dim MatRecaudos() As Recaudo
    Dim MatHabilitaciones() As HabilitaVehiculo
    Dim bExistenCtasAbonoRecaudo As Boolean
    Dim bExito As Boolean
    Dim oCredito As COMNCredito.NCOMCredito 'BRGO 20120707
    Dim oImpre As New FCOMImpresion
    Dim nTotalActivados As Integer 'BRGO 20120707
    Dim oPrevio As previo.clsprevio
    Dim lsCadena As String
    
    If MsgBox("¿Esta seguro de grabar los datos?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If

    Dim oCap As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsCap As ADODB.Recordset
    Dim i As Integer
    
    Set oCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    i = 0

    Select Case fsOpeCod
        Case "05000"
            
            nTotalActivados = 0
            Set oCredito = New COMNCredito.NCOMCredito
            ReDim MatHabilitaciones(feCreditosAprobados.Rows - 2)
            For i = 1 To feCreditosAprobados.Rows - 1
                MatHabilitaciones(i - 1).PersNombre = feCreditosAprobados.TextMatrix(i, 1)
                MatHabilitaciones(i - 1).CtaCod = feCreditosAprobados.TextMatrix(i, 2)
                MatHabilitaciones(i - 1).Placa = feCreditosAprobados.TextMatrix(i, 3)
                MatHabilitaciones(i - 1).TallerCod = feCreditosAprobados.TextMatrix(i, 4)
                MatHabilitaciones(i - 1).FechaHora = Trim(feCreditosAprobados.TextMatrix(i, 5))
                MatHabilitaciones(i - 1).IFICod = feCreditosAprobados.TextMatrix(i, 6)
                MatHabilitaciones(i - 1).PersCod = feCreditosAprobados.TextMatrix(i, 7)
                lsCadena = lsCadena & MatHabilitaciones(i - 1).CtaCod
                lsCadena = lsCadena & Chr$(10)
            Next
            
            bExito = oCredito.RegistrarHabilitacionesInfoGas(MatHabilitaciones)

            If bExito Then
                MsgBox "Se ha registrado satisfactoriamente las activaciones de Crédito INFOGAS", vbInformation, "Aviso"
                
                lsCadena = oCredito.ImprimeVehiculosActivados(MatHabilitaciones)
                Set oPrevio = New previo.clsprevio
                    oPrevio.Show lsCadena, "Vehículos Habilitados Ecotaxi"
                Set oPrevio = Nothing
            Else
                MsgBox "Se presentaron errores en el proceso", vbInformation, "Aviso"
            End If
            FormatearGrillaAprobacion
'        Case gAhoDepCtaRecaudoEcotaxi
'            CmdCargaArch.Enabled = False
'            cmdGrabar.Enabled = False
'            bExistenCtasAbonoRecaudo = True
'            ReDim MatRecaudos(feRecaudo.Rows - 2)
'            For i = 1 To feRecaudo.Rows - 1
'                MatRecaudos(i - 1).CtaCodCredito = Trim(feRecaudo.TextMatrix(i, 2))
'                MatRecaudos(i - 1).Placa = Trim(feRecaudo.TextMatrix(i, 3))
'                MatRecaudos(i - 1).PersCod = Trim(feRecaudo.TextMatrix(i, 8))
'                MatRecaudos(i - 1).EESSFechaHora = CDate(Trim(feRecaudo.TextMatrix(i, 4)))
'                MatRecaudos(i - 1).IFICod = Trim(feRecaudo.TextMatrix(i, 10))
'                MatRecaudos(i - 1).EESSNombre = Trim(feRecaudo.TextMatrix(i, 6))
'                MatRecaudos(i - 1).EESSNroTicket = Trim(feRecaudo.TextMatrix(i, 7))
'                MatRecaudos(i - 1).EESSRecaudoBruto = CDbl(Trim(feRecaudo.TextMatrix(i, 12)))
'                MatRecaudos(i - 1).EESSITFValorEntrada = CDbl(Trim(feRecaudo.TextMatrix(i, 14)))
'                MatRecaudos(i - 1).EESSITFValorSalida = CDbl(Trim(feRecaudo.TextMatrix(i, 15)))
'                MatRecaudos(i - 1).EESSRecaudoNeto = CDbl(Trim(feRecaudo.TextMatrix(i, 12)))
'                MatRecaudos(i - 1).COFIDEPorcentajeComision = CDbl(Trim(feRecaudo.TextMatrix(i, 13)))
'                MatRecaudos(i - 1).IFIRecaudoNeto = CDbl(feRecaudo.TextMatrix(i, 5)) 'Recaudo Transferido de COFIDE a la CAJA
'                MatRecaudos(i - 1).EESSId = Trim(feRecaudo.TextMatrix(i, 9))
'
'                'Obtener cta abono
'                Set rsCap = oCap.GetCuentasPersona(MatRecaudos(i - 1).PersCod, gCapAhorros, True, False, CInt(Mid(MatRecaudos(i - 1).CtaCodCredito, 9, 1)), , , 7)
'                MatRecaudos(i - 1).CtaCodAbono = IIf(IsNull(rsCap!cCtaCod), "", rsCap!cCtaCod)
'
'                If MatRecaudos(i - 1).CtaCodAbono = "" Or Len(MatRecaudos(i - 1).CtaCodAbono) <> 18 Then
'                    feRecaudo.Row = i
'                    feRecaudo.SetFocus
'                    Call feRecaudo.BackColorRow(vbYellow, True)
'                    bExistenCtasAbonoRecaudo = False
'                End If
'            Next
'
'            If bExistenCtasAbonoRecaudo = False Then
'                MsgBox "Verifique, Los registros resaltados no cuentan con Cta Abono Ecotaxi", vbExclamation, "Aviso"
'                Exit Sub
'            End If
'
'            bExito = RegistrarRecaudos(MatRecaudos, gdFecSis, gsCodAge, gsCodUser)
'
'            If bExito Then
'                MsgBox "Se ha registrado con éxito los recaudos de Ecotaxi", vbInformation, "Aviso"
'            Else
'                MsgBox "No se ha podido terminar el proceso de Abono de Recaudo" & Chr(10) & "Consulte con el Dpto de TI", vbCritical, "Aviso"
'            End If
        Case Else
            MsgBox "La Operación Actual No existe", vbInformation, "Aviso"
    End Select
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
'Private Function SeRealizoAbonoRecaudo(ByVal pdFecha As Date) As Boolean
'    Dim oCredito As COMDCredito.DCOMCredito
'    Set oCredito = New COMDCredito.DCOMCredito
'    SeRealizoAbonoRecaudo = oCredito.SeRealizoAbonoRecuado(pdFecha)
'    Set oCredito = Nothing
'End Function
'Public Function RegistrarRecaudos(ByRef pMatRecaudos() As Recaudo, ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String) As Boolean
'    Dim oBase As COMDCredito.DCOMCredActBD
'    Dim oITF As COMDConstSistema.FCOMITF
'    Dim oRecaudo As Recaudo
'    Dim MatCtasAbono() As AbonoEcoTaxi
'    Dim sMovNro As String
'    Dim bTransac As Boolean, bExisteCtaAbono As Boolean
'    Dim pMatDatosAhoAbo As Variant
'    Dim nITFAbono As Double, nRedondeoITF As Double
'    Dim nMovNro As Long
'    Dim i As Integer, j As Integer, lnRegistros As Integer, tam As Integer, iContMovITF As Integer
'
'On Error GoTo ErrRegistrarRecaudos
'    Set oBase = New COMDCredito.DCOMCredActBD
'    Set oBarra = New clsProgressBar
'
'    oBarra.ShowForm Me
'    oBarra.CaptionSyle = eCap_CaptionPercent
'
'    sMovNro = oBase.GeneraMovNro(pdFecSis, Right(psCodAge, 2), psCodUser)
'
'    bTransac = False
'    Call oBase.dBeginTrans
'    bTransac = True
'
'    lnRegistros = UBound(pMatRecaudos) + 1
'
'    Call oBase.InsertaMov(sMovNro, gAhoDepCtaRecaudoEcotaxi, "Abono de Racaudo EcoTaxi: " & lnRegistros & " registro(s)", gMovEstContabMovContable, gMovFlagVigente)
'    nMovNro = oBase.GetnMovNro(sMovNro)
'
'    'Guarda Detalle de Recaudo
'    ReDim MatCtasAbono(0)
'    For i = 0 To lnRegistros - 1
'        oRecaudo = pMatRecaudos(i)
'        Call oBase.dRegistrarRecaudo(nMovNro, oRecaudo.Placa, oRecaudo.EESSFechaHora, oRecaudo.EESSNombre, oRecaudo.EESSNroTicket, oRecaudo.EESSRecaudoBruto, oRecaudo.EESSITFValorEntrada, oRecaudo.EESSITFValorSalida, oRecaudo.EESSRecaudoNeto, oRecaudo.COFIDEPorcentajeComision, oRecaudo.IFIRecaudoNeto, oRecaudo.CtaCodCredito, oRecaudo.CtaCodAbono, oRecaudo.PersCod, oRecaudo.IFICod, oRecaudo.EESSId)
'        If i = 0 Then
'            MatCtasAbono(0).CtaCodAbono = oRecaudo.CtaCodAbono
'            MatCtasAbono(0).Abono = oRecaudo.IFIRecaudoNeto
'        Else
'            tam = UBound(MatCtasAbono) + 1
'            bExisteCtaAbono = False
'            For j = 0 To tam - 1
'                If oRecaudo.CtaCodAbono = MatCtasAbono(j).CtaCodAbono Then
'                    bExisteCtaAbono = True
'                    Exit For
'                End If
'            Next
'            If bExisteCtaAbono Then
'                MatCtasAbono(j).Abono = CDbl(MatCtasAbono(j).Abono) + oRecaudo.IFIRecaudoNeto
'            Else
'                ReDim Preserve MatCtasAbono(tam)
'                MatCtasAbono(tam).CtaCodAbono = oRecaudo.CtaCodAbono
'                MatCtasAbono(tam).Abono = oRecaudo.IFIRecaudoNeto
'            End If
'        End If
'    Next
'    'Inicializa Datos de Ahorros
'    ReDim pMatDatosAhoAbo(14)
'    pMatDatosAhoAbo(0) = "" 'Cuenta de Ahorros
'    pMatDatosAhoAbo(1) = "0.00" 'Monto de Apertura
'    pMatDatosAhoAbo(2) = "0.00" 'Interes Ganado de Abono
'    pMatDatosAhoAbo(3) = "0.00" 'Interes Ganado de Retiro Gastos
'    pMatDatosAhoAbo(4) = "0.00" 'Interes Ganado de Retiro Cancelaciones
'    pMatDatosAhoAbo(5) = "0.00" 'Monto de Abono
'    pMatDatosAhoAbo(6) = "0.00" 'Monto de Retiro de Gastos
'    pMatDatosAhoAbo(7) = "0.00" 'Monto de Retiro de Cancelaciones
'    pMatDatosAhoAbo(8) = "0.00" 'Saldo Disponible Abono
'    pMatDatosAhoAbo(9) = "0.00" 'Saldo Contable Abono
'    pMatDatosAhoAbo(10) = "0.00" 'Saldo Disponible Retiro de Gastos
'    pMatDatosAhoAbo(11) = "0.00" 'Saldo Contable Retiro de Gastos
'    pMatDatosAhoAbo(12) = "0.00" 'Saldo Disponible Retiro de Cancelaciones
'    pMatDatosAhoAbo(13) = "0.00" 'Saldo Contable Retiro de Cancelaciones
'
'    oBarra.Max = UBound(MatCtasAbono) + 1
'    oBarra.Progress 0, "Proceso de Abono Cta x Recaudo EcoTaxi", "Preparando Abono...", "Recaudo Ecotaxi", vbBlue
'
'    Set oITF = New COMDConstSistema.FCOMITF
'    oITF.fgITFParametros
'    iContMovITF = 1
'
'    'Abona Ctas de Ahorro Ecotaxi
'    For i = 0 To UBound(MatCtasAbono)
'        nITFAbono = oITF.fgTruncar(oITF.fgITFCalculaImpuesto(MatCtasAbono(i).Abono), 2)
'        nRedondeoITF = fgDiferenciaRedondeoITF(nITFAbono)
'        nITFAbono = IIf(nRedondeoITF > 0, nITFAbono - nRedondeoITF, nITFAbono)
'
'        oBase.CapAbonoCuentaAho pMatDatosAhoAbo, MatCtasAbono(i).CtaCodAbono, CDbl(MatCtasAbono(i).Abono), gAhoDepCtaRecaudoEcotaxi, sMovNro, "Abono x Recaudo a la Cta Abono de Ecotaxi N°: " & MatCtasAbono(i).CtaCodAbono, , , , , , , pdFecSis, "", True, nITFAbono, False, gITFCobroCargo
'
'        If nITFAbono + nRedondeoITF > 0 Then
'           Call oBase.InsertaMovRedondeoITF(sMovNro, iContMovITF, nITFAbono + nRedondeoITF, nITFAbono)
'           iContMovITF = iContMovITF + 1
'        End If
'        oBarra.Progress (i + 1), "Proceso de Abono Cta x Recaudo EcoTaxi", "Efectuando Abono Nro " & (i + 1) & " Cuenta: " & MatCtasAbono(i).CtaCodAbono, "Recaudo EcoTaxi", vbBlue
'    Next
'
'    Call oBase.dRollbackTrans
'    RegistrarRecaudos = bTransac
'
'    oBarra.CloseForm Me
'    Set oBarra = Nothing
'    Exit Function
'ErrRegistrarRecaudos:
'    If bTransac Then
'        Call oBase.dRollbackTrans
'        Set oBarra = Nothing
'        Set oBase = Nothing
'    End If
'    RegistrarRecaudos = False
'    err.Raise err.Number, "Error En Proceso", err.Description
'End Function
