VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmCapCargaArchivo 
   BackColor       =   &H80000016&
   Caption         =   "Servicio de Recaudo - Carga de Archivo"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5865
   Icon            =   "frmCapCargaArchivo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCargaArchivo 
      BackColor       =   &H80000016&
      Caption         =   "Carga de Trama"
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
      Height          =   4335
      Left            =   158
      TabIndex        =   0
      Top             =   165
      Width           =   5535
      Begin MSWinsockLib.Winsock wsRecaudo 
         Left            =   3570
         Top             =   2625
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton cmdSelecArchivo 
         BackColor       =   &H80000016&
         Caption         =   "1. Selección de Archivo"
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
         Left            =   225
         TabIndex        =   3
         Top             =   405
         Width           =   2325
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "2. Carga de DATA"
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
         Height          =   360
         Left            =   210
         TabIndex        =   2
         Top             =   2520
         Width           =   2220
      End
      Begin VB.CommandButton cmdCerrar 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
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
         Left            =   13155
         TabIndex        =   1
         Top             =   7695
         Width           =   960
      End
      Begin MSComDlg.CommonDialog dlgArchivo 
         Left            =   2940
         Top             =   315
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ProgressBar oBarra 
         Height          =   300
         Left            =   4305
         TabIndex        =   4
         Top             =   210
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Max             =   3.00000e5
      End
      Begin SICMACT.FlexEdit grdConvenio 
         Height          =   795
         Left            =   14175
         TabIndex        =   13
         Top             =   7350
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   1402
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-ID-Cod. Cliente-Tipo DOI-DOI-Cliente-Servicio-Concepto-Importe"
         EncabezadosAnchos=   "0-700-1500-850-1500-3500-3500-3500-1000"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColor       =   -2147483628
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-L-C-C-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
         CellBackColor   =   -2147483628
      End
      Begin VB.Label lblCargando 
         Caption         =   "Cargando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2190
         TabIndex        =   12
         Top             =   3255
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Empresa: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   11
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Convenio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   1430
         Width           =   870
      End
      Begin VB.Label lblEmpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1785
         TabIndex        =   9
         Top             =   930
         Width           =   3525
      End
      Begin VB.Label lblConvenio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1785
         TabIndex        =   8
         Top             =   1365
         Width           =   3510
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   1830
         Width           =   750
      End
      Begin VB.Label lblCodigoConvenio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1785
         TabIndex        =   6
         Top             =   1785
         Width           =   3510
      End
      Begin VB.Label lblerror 
         BackColor       =   &H80000016&
         Caption         =   "errores durante la carga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1680
         MouseIcon       =   "frmCapCargaArchivo.frx":030A
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   3675
         Visible         =   0   'False
         Width           =   2220
      End
   End
End
Attribute VB_Name = "frmCapCargaArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************************************************************************************
'* NOMBRE         : "frmCapCargaArchivo"
'* DESCRIPCION    : Formulario creado para el pago de servicios de convenios segun proyecto: "Mejora del Sistema y Automatizacion de Ahorros y Servicios"
'* CREACION       : RIRO, 20121213 10:00 AM
'********************************************************************************************************************************************************

Option Explicit

Private sEmpresa As String
Private sConvenio As String
Private sCodigoConvenio As String
Private sTipoValidacion As String
Private nRegistros As Double
Private sFechaPrescripcion As String
Private listError() As String
Private nListError As Double
Private Mensaje As String
Dim rsRecaudo As ADODB.Recordset
Dim rsRecaudoError As ADODB.Recordset
Dim objExcel As Excel.Application
Dim xLibro As Excel.Workbook
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim sCadena As String
Dim campos() As String
Dim nFila As Double
Dim sRuta, sRutaLog, sRutaFormato As String
Dim oRecaudo As COMDCaptaServicios.DCOMServicioRecaudo

Private Sub cmdCargar_Click()

    Dim sCadena As String
    Dim campos() As String
    Dim sIp As String
    Dim bResultado As Boolean
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Dim sDireccionTemporales, sDireccionErrores, sDireccionFormatos As String
    Dim sInput, sRutaTemporal As String
    Dim rsRespuestaCargaTemporal As ADODB.Recordset
    Dim dInicio, dFin As Date
    
    On Error GoTo error
    
    If MsgBox("Está seguro de cargar la trama?, " _
                        & "se reemplazarán los datos de la ultima trama cargada de este convenio", _
                        vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    
    End If
        
    'Borrando data: Verificando que exista el convenio citado
    Set oRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
    Set oCont = New COMNContabilidad.NCOMContFunciones
    dInicio = Now

    Open dlgArchivo.Filename For Input As #1

        Line Input #1, sCadena
        'Llenando cabecera
        campos = Split(sCadena, "|")
        sCodigoConvenio = campos(0)
        sTipoValidacion = campos(1)
        nRegistros = CDbl(campos(2))
        sFechaPrescripcion = campos(3)

    Close #1

    Set rsRecaudo = oRecaudo.getBuscaConvenioXCodigo(sCodigoConvenio)

    If rsRecaudo.EOF Then
        Set rsRecaudo = Nothing
        MsgBox "El convenio seleccionado no se encuentra registrado en la Base de Datos", vbInformation, "Aviso"
        Exit Sub
    End If
    
    sIp = wsRecaudo.LocalIP
    sDireccionTemporales = "\\" & wsRecaudo.LocalIP & "\Trama\Temporales"
    sDireccionErrores = "\\" & wsRecaudo.LocalIP & "\Trama\Errores"
    sDireccionFormatos = "\\" & wsRecaudo.LocalIP & "\Trama\Formatos"
    
    ' *** Crendo Directorios en caso no existan
    If Dir$(sDireccionTemporales, vbDirectory) = vbNullString Then Call MkDir(sDireccionTemporales)
    If Dir$(sDireccionErrores, vbDirectory) = vbNullString Then Call MkDir(sDireccionErrores)
    If Dir$(sDireccionFormatos, vbDirectory) = vbNullString Then Call MkDir(sDireccionFormatos)
    ' *** Fin de Creacion
    
    sRutaTemporal = sDireccionTemporales & "\TEMP" & Year(gdFecSis) _
                                                   & Format(Month(gdFecSis), "##,00") _
                                                   & Format(Day(gdFecSis), "##,00") _
                                                   & Format(Hour(Now), "##,00") _
                                                   & Format(Minute(Now), "##,00") _
                                                   & Format(Second(Now), "##,00") _
                                                   & dlgArchivo.FileTitle
    
    sRutaLog = "\\" & sIp & "\Trama\Errores\ERROR" & Year(gdFecSis) _
                                                   & Format(Month(gdFecSis), "##,00") _
                                                   & Format(Day(gdFecSis), "##,00") _
                                                   & Format(Hour(Now), "##,00") _
                                                   & Format(Minute(Now), "##,00") _
                                                   & Format(Second(Now), "##,00") _
                                                   & ".Log"
    
    Call FileSystem.FileCopy(sRuta, sRutaTemporal) ' Generando copia de trama

    Open sRutaTemporal For Input As #1
        Line Input #1, sInput
        Line Input #1, sInput
    Close #1

    Open sRutaTemporal For Append As #1
        Print #1, sInput
    Close #1
    FileSystem.Reset
    bResultado = True

    Set rsRecaudo = Nothing

    'Proceso de Carga Temporal
    lblCargando.Visible = True
    DoEvents
    Set rsRespuestaCargaTemporal = oRecaudo.CargaTemporal(sCodigoConvenio, _
                                                          sRutaTemporal, _
                                                          sRutaLog, _
                                                          sRutaFormato)

    
    If Not rsRespuestaCargaTemporal.EOF And Not rsRespuestaCargaTemporal.BOF Then
        If rsRespuestaCargaTemporal!valor = 2 Then
            MsgBox "Error en el proceso de carga temporal, verificar la etructura de la trama y el tipo de dato que contenga.", vbExclamation, "Aviso"
            Set rsRecaudoError = Nothing
            lblCargando.Visible = False
            Exit Sub
        End If
    End If

    Set rsRecaudoError = Nothing
    
    'Proceso de Validacion
    DoEvents
    Set rsRecaudoError = oRecaudo.ValidarTrama(sCodigoConvenio, _
                         IIf(Len(sTipoValidacion) = 1, sTipoValidacion, _
                         Right(sTipoValidacion, 1)), nRegistros, _
                         sFechaPrescripcion, oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), _
                         sRutaTemporal, sRutaLog, sRutaFormato)
    
    DoEvents
    If Not rsRecaudoError.EOF And Not rsRecaudoError.BOF Then
        If rsRecaudoError.RecordCount > 0 Then
            
            MsgBox "Error en el proceso de validacion, verificar el detalle de los errores", vbExclamation, "Aviso"
            lblerror.Caption = rsRecaudoError.RecordCount & " errores durante la carga"
            lblerror.Visible = True
            lblCargando.Visible = False
            Exit Sub
            
        End If
    End If
  
    'Proceso de Carga de Trama
    DoEvents
    bResultado = oRecaudo.CargarTrama(sCodigoConvenio, _
                 IIf(Len(sTipoValidacion) = 1, sTipoValidacion, _
                 Right(sTipoValidacion, 1)), nRegistros, _
                 sFechaPrescripcion, oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), _
                 sRutaTemporal, sRutaLog)
    
    ' *** Eliminando Trama
    Dim oSys As Scripting.FileSystemObject
    Set oSys = New Scripting.FileSystemObject
    If oSys.FileExists(sRutaTemporal) = True Then
         Kill sRutaTemporal
    End If
    Set oSys = Nothing
    ' *** Fin Elimnacion
    
    dFin = Now
    lblCargando.Visible = False
    Dim sTiempo As String
    sTiempo = Round((DateDiff("s", dInicio, dFin) / 60), 2) & " Minutos"
    DoEvents
    If bResultado Then
        Limpiar
        DoEvents
        MsgBox "La Carga Concluyó Exitosamente " & vbNewLine _
                                                  & "Tiempo de carga: " _
                                                  & sTiempo _
                                                  & vbNewLine _
                                                  & "(" & DateDiff("s", dInicio, dFin) & " Segundos)", vbInformation, "Aviso"
    Else
        MsgBox "Error en el proceso carga de trama", vbExclamation, "Aviso"
        
    End If
        
    Exit Sub
    
error:
    lblCargando.Visible = False
    MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdSelecArchivo_Click()
    
    Dim sCadena, sIp As String
    Dim oSys As Scripting.FileSystemObject
    Dim campos() As String
    
    sIp = wsRecaudo.LocalIP
    Set oRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
    Set oSys = New Scripting.FileSystemObject

    dlgArchivo.Filename = Empty
    dlgArchivo.Filter = "Archivos de Texto (*.txt)|*.txt|Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx"
    dlgArchivo.ShowOpen
    If dlgArchivo.Filename = Empty Then
        MsgBox "No se abrió ningun archivo", vbQuestion, "Aviso"
        cmdCargar.Enabled = False
       Exit Sub
    End If
    Limpiar
    
    sRuta = "\\" & sIp & "\Trama\" & dlgArchivo.FileTitle
    sRutaFormato = "\\" & sIp & "\Trama\Formatos\Formato.xml"
    
    ' *** Verificando que el archivo seleccionado se encuentre compartido en la carpeta trama
    If oSys.FileExists(sRuta) = False Then
        MsgBox "Es necesario crear una carpeta compartida con el nombre ""Trama""," & vbNewLine & _
        "darle los permisos y copiar la trama para la carga. No deberá crear sub carpetas ", vbExclamation, "Corregir!!!"
        cmdCargar.Enabled = False
        Set oSys = Nothing
        Exit Sub
    End If
        
    ' *** Verificando que exista el formato para la trama
    If oSys.FileExists(sRutaFormato) = False Then
        If oSys.FileExists(App.path & "\FormatoCarta\Formato.xml") = True Then
            Dim sDirFormatoDirectorio As String
            sDirFormatoDirectorio = "\\" & wsRecaudo.LocalIP & "\Trama\Formatos"
            If Dir$(sDirFormatoDirectorio, vbDirectory) = vbNullString Then Call MkDir(sDirFormatoDirectorio)
            Call FileSystem.FileCopy(App.path & "\FormatoCarta\Formato.xml", sRutaFormato)  ' Copiando Formato de tabla a carpeta
        Else
            MsgBox "No se pudo encontrar el archivo: Formato.xml, es necesario copiarlo en la carpeta " & App.path & "\FormatoCarta\", vbExclamation, "Aviso"
            cmdCargar.Enabled = False
            Set oSys = Nothing
            Exit Sub
        End If
    End If
    ' *** Fin verificacion.

    Open sRuta For Input As #1
        Line Input #1, sCadena
        campos = Split(sCadena, "|")
        sCodigoConvenio = campos(0)
    Close #1

    Set rsRecaudo = oRecaudo.getBuscaConvenioXCodigo(sCodigoConvenio)
    
    If rsRecaudo.EOF Then
        MsgBox "El convenio seleccionado no se encuentra registrado en la Base de Datos", vbInformation, "Aviso"
        Set rsRecaudo = Nothing
        cmdCargar.Enabled = False
        Exit Sub
    Else
        lblCodigoConvenio.Caption = rsRecaudo!cCodConvenio
        lblEmpresa.Caption = rsRecaudo!cPersNombre
        lblConvenio.Caption = rsRecaudo!cNombreConvenio
        cmdCargar.Enabled = True
    End If

    Exit Sub
        
End Sub

'Public Sub cargarListaError( _
'                        ByVal nFila As Double, _
'                        ByVal sMensajeError As String _
'                        )
'    Dim sFila As String
'    Dim i As Double
'
'    nListError = nListError + 1
'    ReDim Preserve listError(nListError)
'
'    For i = 1 To grdConvenio.Cols - 1
'
'        If nFila = 0 Then
'            sFila = sFila & " - ;"
'        Else
'            sFila = sFila & grdConvenio.TextMatrix(nFila, i) & ";"
'
'        End If
'
'    Next
'
'    sFila = Mid(sFila, 1, Len(sFila) - 1)
'    listError(nListError) = sMensajeError & "|" & sFila
'
'End Sub

Private Sub Limpiar()

    Dim i As Double
    
    grdConvenio.Clear
    grdConvenio.FormaCabecera
    grdConvenio.Rows = 2
    ReDim listError(0)
    listError(0) = ""
    sConvenio = ""
    sCodigoConvenio = ""
    sTipoValidacion = ""
    sFechaPrescripcion = ""
    sEmpresa = ""
    lblerror.Caption = ""
    lblerror.Visible = False
    
    lblEmpresa.Caption = ""
    lblConvenio.Caption = ""
    lblCodigoConvenio.Caption = ""
    nRegistros = 0
    nListError = 0
    cmdCargar.Enabled = False
    cmdSelecArchivo.SetFocus
    oBarra.Visible = False
    
End Sub

' Cargando archivo Excel
'Private Sub cargarExcel()
'
'    Dim bUltimaFila As Boolean
'    Dim filaExcel, Fila As Double
'
'    Set objExcel = New Excel.Application
'    Set xLibro = objExcel.Workbooks.Open(dlgArchivo.Filename)
'    Set oExcel = CreateObject("Excel.Application")
'    Set oBook = oExcel.Workbooks.Add
'    Set oSheet = oBook.Worksheets(1)
'    Set oRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
'
'    On Error GoTo error
'
'    filaExcel = 25
'    bUltimaFila = True
'    Fila = filaExcel
'
'        With xLibro
'            With .Sheets(1)
'
'                nRegistros = CDbl(Trim(.Cells(16, 4)))
'                sCodigoConvenio = Trim(Trim(.Cells(12, 4)) & Trim(.Cells(12, 5)) & Trim(.Cells(12, 6)))
'                sTipoValidacion = Trim(.Cells(12, 5))
'                sFechaPrescripcion = Trim(.Cells(18, 4))
'                Set rsRecaudo = oRecaudo.getBuscaConvenioXCodigo(sCodigoConvenio)
'                Set oRecaudo = Nothing
'
'                If Not rsRecaudo.EOF Then
'
'                    sConvenio = rsRecaudo!cNombreConvenio
'                    sEmpresa = rsRecaudo!cPersNombre
'                    sCodigoConvenio = rsRecaudo!cCodConvenio
'
'                    lblConvenio.Caption = sConvenio
'                    lblCodigoConvenio.Caption = sCodigoConvenio
'                    lblEmpresa.Caption = sEmpresa
'
'                    If Not UCase(sTipoValidacion) = UCase(rsRecaudo!nTipo) Then
'                        Call cargarListaError(0, _
'                        "El tipo de validacion definida en la trama no corresponde al tipo de validacion del convenio")
'                    End If
'
'                Else
'                    Call cargarListaError(0, _
'                    "El Codigo enviado en la trama no corresponde a un convenio válido")
'                End If
'
'                oBarra.Max = nRegistros + Fila
'                oBarra.Min = filaExcel
'                oBarra.value = filaExcel
'
'                oBarra.Visible = True
'                lblEmpresa.Caption = sEmpresa
'                lblConvenio.Caption = sConvenio
'                lblCodigoConvenio.Caption = sCodigoConvenio
'
'                Do While bUltimaFila
'
'                    DoEvents
'                    grdConvenio.AdicionaFila
'                    nFila = grdConvenio.Rows - 1
'
'                    If Fila <= oBarra.Max Then
'                        oBarra.value = Fila
'                    End If
'
'                    Select Case sTipoValidacion
'
'                        Case "VC", "VI"
'
'                            'Campo Id
'                            If Trim(.Cells(Fila, 3)) = "" Then
'                                grdConvenio.TextMatrix(nFila, 1) = "- -"
'                                Mensaje = Mensaje & "Campo 'ID' vacio " & Space(10) & 1 & ";"
'                            Else
'                                grdConvenio.TextMatrix(nFila, 1) = .Cells(Fila, 3)
'                            End If
'
'                            'Campo CodCliente
'                            If Trim(.Cells(Fila, 4)) = "" Then
'                                grdConvenio.TextMatrix(nFila, 2) = "- -"
'                                Mensaje = Mensaje & "Campo 'CODCLIENTE' vacio " & Space(10) & 2 & ";"
'                            Else
'                                grdConvenio.TextMatrix(nFila, 2) = .Cells(Fila, 4)
'                            End If
'
'                            'Campo TipoDoi
'                            If Trim(.Cells(Fila, 5)) = "1" Then
'                                grdConvenio.TextMatrix(nFila, 3) = "DNI"
'                            ElseIf .Cells(Fila, 5) = "2" Then
'                                grdConvenio.TextMatrix(nFila, 3) = "RUC"
'                            Else
'                                grdConvenio.TextMatrix(nFila, 3) = "- -"
'                                Mensaje = Mensaje & "Campo 'TIPO DOI' SIN DATOS VALIDOS " & Space(10) & 3 & ";"
'                            End If
'
'                            'Cmpo Doi
'                            If Trim(.Cells(Fila, 6)) = "" Then
'                                grdConvenio.TextMatrix(nFila, 4) = "- -"
'                                Mensaje = Mensaje & "Campo 'DOI' vacio " & Space(10) & 4 & ";" ' --
'                            Else
'                                If Trim(.Cells(Fila, 5)) = "1" Then
'                                    If Len(Trim(.Cells(Fila, 6))) <> 8 Then
'                                        Mensaje = Mensaje & "Nro documento no tiene 8 digitos " & Space(10) & 4 & ";" ' --
'                                    End If
'                                ElseIf Trim(.Cells(Fila, 5)) = "2" Then
'                                    If Len(Trim(.Cells(Fila, 6))) <> 11 Then
'                                        Mensaje = Mensaje & "Nro documento no tiene 11 digitos " & Space(10) & 4 & ";" ' --
'                                    End If
'                                End If
'                                grdConvenio.TextMatrix(nFila, 4) = .Cells(Fila, 6)
'                            End If
'
'                            'Campo Cliente
'                            If Trim(.Cells(Fila, 7)) = "" Then
'                                grdConvenio.TextMatrix(nFila, 5) = "- -"
'                                Mensaje = Mensaje & "Campo 'Cliente' vacio " & Space(10) & 5 & ";" ' --
'                            Else
'                                grdConvenio.TextMatrix(nFila, 5) = .Cells(Fila, 7)
'                            End If
'
'                            'Campo Servicio
'                            If Trim(.Cells(Fila, 8)) = "" Then
'                                If Trim(.Cells(Fila, 9)) = "" Then
'                                    grdConvenio.TextMatrix(nFila, 6) = "- -"
'                                    Mensaje = Mensaje & "Campo 'Servicio' vacio " & Space(10) & 6 & ";" ' --
'                                End If
'                            Else
'                                grdConvenio.TextMatrix(nFila, 6) = Trim(.Cells(Fila, 8))
'
'                            End If
'
'                            'Campo Concepto
'                            If Trim(.Cells(Fila, 9)) = "" Then
'                                If Trim(.Cells(Fila, 8)) = "" Then
'                                    grdConvenio.TextMatrix(nFila, 7) = "- -"
'                                    Mensaje = Mensaje & "Campo 'Concepto' vacio " & Space(10) & 7 & ";" ' --
'                                End If
'                            Else
'                                grdConvenio.TextMatrix(nFila, 7) = Trim(.Cells(Fila, 9))
'
'                            End If
'
'                            'Campo Importe
'                            If Trim(.Cells(Fila, 10)) = "" Then
'                                If sTipoValidacion = "VC" Then
'                                    grdConvenio.TextMatrix(nFila, 8) = "- -"
'                                    Mensaje = Mensaje & "Campo 'Importe' vacio " & Space(10) & 8 & ";"
'                                ElseIf sTipoValidacion = "VI" Then
'                                    grdConvenio.TextMatrix(nFila, 8) = "- -"
'                                End If
'                            Else
'                                grdConvenio.TextMatrix(nFila, 8) = Format(.Cells(Fila, 10), "#,##0.00")
'                            End If
'
'                        Case "VP"
'
'                            'Campo Id
'                            If Trim(.Cells(Fila, 3)) = "" Then
'                                grdConvenio.TextMatrix(nFila, 1) = "- -"
'                                Mensaje = Mensaje & "Campo 'ID' vacio " & Space(10) & 1 & ";" ' --
'                            Else
'                                grdConvenio.TextMatrix(nFila, 1) = .Cells(Fila, 3)
'                            End If
'
'                            'Campo Concepto
'                            If Trim(.Cells(Fila, 9)) <> "" Then
'                                grdConvenio.TextMatrix(nFila, 7) = .Cells(Fila, 9)
'                            Else
'                                grdConvenio.TextMatrix(nFila, 7) = "- -"
'                                Mensaje = Mensaje & "Campo 'Concepto' vacio " & Space(10) & 7 & ";" ' --
'                            End If
'
'                            'Campo Importe
'                            If Trim(.Cells(Fila, 10)) = "" Then
'                                grdConvenio.TextMatrix(nFila, 8) = "- -"
'                                Mensaje = Mensaje & "Campo 'Importe' vacio " & Space(10) & 8 & ";" ' --
'                            Else
'                                grdConvenio.TextMatrix(nFila, 8) = Format(.Cells(Fila, 10), "#,##0.00")
'                            End If
'
'                    End Select
'
'                    If Trim(Mensaje) <> "" Then
'                        Mensaje = Mid(Mensaje, 1, Len(Mensaje) - 1)
'                        Call cargarListaError(nFila, Mensaje)
'                        Mensaje = ""
'                    End If
'
'                    If Trim(.Cells(Fila + 1, 3)) = "" And _
'                        Trim(.Cells(Fila + 1, 4)) = "" And _
'                        Trim(.Cells(Fila + 1, 5)) = "" And _
'                        Trim(.Cells(Fila + 1, 6)) = "" And _
'                        Trim(.Cells(Fila + 1, 7)) = "" And _
'                        Trim(.Cells(Fila + 1, 8)) = "" And _
'                        Trim(.Cells(Fila + 1, 9)) = "" And _
'                        Trim(.Cells(Fila + 1, 10)) = "" Then
'                        bUltimaFila = False
'                    End If
'
'                    lblTotalRegistros.Caption = (Fila - 24)
'                    Fila = Fila + 1
'
'                Loop
'
'            End With
'
'        End With
'
'        objExcel.Quit
'        oExcel.Quit
'        Set rsRecaudo = Nothing
'        Set objExcel = Nothing
'        Set xLibro = Nothing
'        Set oBook = Nothing
'        oBarra.Visible = False
'
'        If grdConvenio.Rows - 1 <> nRegistros Then
'
'            Call cargarListaError(0, _
'            "La cantidad de registros de la trama cabecera no corresponde a la cantidad total de registros de la sección detalles")
'
'        End If
'
'        If nListError > 0 Then
'
'            lblerror.Left = 5880
'            lblerror.Top = 7740
'            lblerror.Caption = nListError & " Errores durante la carga"
'            lblerror.Visible = True
''            cmdCargar.Enabled = False
'
'        Else
'           ' cmdCargar.Enabled = True
'
'        End If
'
'        Exit Sub
'
'error:
'        err.Raise err.Number, err.Source, err.Description
'        objExcel.Quit
'        oExcel.Quit
'        Set rsRecaudo = Nothing
'        Set objExcel = Nothing
'        Set xLibro = Nothing
'        Set oBook = Nothing
'        oBarra.Visible = False
'
'End Sub

'Private Sub cargarTxt()
'
'    Dim sCadena As String
'    Dim campos() As String
'    Dim c As Variant
'    Dim i As Double
'    Dim j As Double
'    Dim h As Double
'
'    On Error GoTo error
'
'    i = 1
'    j = 1
'
'    Set oRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
'    Open dlgArchivo.Filename For Input As #1
'
'            Line Input #1, sCadena
'
'            'Llenando cabecera
'            campos = Split(sCadena, "|")
'            sCodigoConvenio = campos(0)
'            sTipoValidacion = campos(1)
'            nRegistros = CDbl(campos(2))
'            sFechaPrescripcion = campos(3)
'
''            oBarra.Min = 0
''            oBarra.value = 0
''            oBarra.Visible = True
''            oBarra.Max = nRegistros
'
'            Set rsRecaudo = oRecaudo.getBuscaConvenioXCodigo(sCodigoConvenio)
'
'            If Not rsRecaudo.EOF Then
'                sEmpresa = rsRecaudo!cPersNombre
'                sConvenio = rsRecaudo!cNombreConvenio
'
'                If Not UCase(sTipoValidacion) = UCase(Right(rsRecaudo!nTipo, 1)) Then
'                    Call cargarListaError(0, _
'                    "El tipo de validacion definida en la trama no corresponde al tipo de validacion del convenio")
'                End If
'
'            Else
'                Call cargarListaError(0, _
'                "El Codigo enviado en la trama no corresponde a un convenio válido")
'            End If
'
'            Set rsRecaudo = Nothing
'            Set oRecaudo = Nothing
'
'            lblEmpresa.Caption = sEmpresa
'            lblConvenio.Caption = sConvenio
'            lblCodigoConvenio.Caption = sCodigoConvenio
'            grdConvenio.AdicionaFila
'
'            Do While Not EOF(1)
'                Line Input #1, sCadena
'                campos = Split(sCadena, "|")
'                DoEvents
'                Select Case UCase(sTipoValidacion)
'
'                    Case "C", "I"
'                            For Each c In campos
'                                Select Case j
'
'                                    Case 1
'                                            If Trim(c) = "" Then
'                                                Mensaje = Mensaje & "Campo 'ID' vacio " & Space(10) & j & ";"
'                                                grdConvenio.TextMatrix(i, j) = " "
'                                            Else
'                                                grdConvenio.TextMatrix(i, j) = c
'                                            End If
'                                    Case 2
'                                            If Trim(c) = "" Then
'                                               Mensaje = Mensaje & "Campo 'Codigo' vacio " & Space(10) & j & ";"
'                                               grdConvenio.TextMatrix(i, j) = " "
'                                            Else
'                                                grdConvenio.TextMatrix(i, j) = c
'                                            End If
'                                    Case 3
'                                            If Trim(c) = "1" Then
'                                               grdConvenio.TextMatrix(i, j) = "DNI"
'                                            ElseIf c = "2" Then
'                                                grdConvenio.TextMatrix(i, j) = "RUC"
'                                            ElseIf c = "" Then
'                                                Mensaje = Mensaje & "Campo 'Tipo DOI' vacio " & Space(10) & j & ";"
'                                                grdConvenio.TextMatrix(i, j) = c
'                                            Else
'                                                Mensaje = Mensaje & "No se identifica tipo DOI " & Space(10) & j & ";"
'                                                grdConvenio.TextMatrix(i, j) = "NO IDENT"
'                                            End If
'                                    Case 4
'                                            If Trim(c) = "" Then
'                                                Mensaje = Mensaje & "Campo 'DOI' vacio " & Space(10) & j & ";"
'                                                grdConvenio.TextMatrix(i, j) = " "
'                                            Else
'                                                If campos(2) = "1" Then
'                                                    If Len(Trim(c)) <> 8 Then
'                                                        Mensaje = Mensaje & "Nro Documento no tiene 8 digitos" & Space(10) & j & ";"
'                                                    End If
'                                                ElseIf campos(2) = "2" Then
'                                                    If Len(Trim(c)) <> 11 Then
'                                                        Mensaje = Mensaje & "Nro Documento no tiene 11" & Space(10) & j & ";"
'                                                    End If
'
'                                                End If
'
'                                                grdConvenio.TextMatrix(i, j) = c
'
'                                            End If
'                                    Case 5
'                                            If Trim(c) = "" Then
'                                                Mensaje = Mensaje & "No especifico nombre de cliente " & Space(10) & j & ";"
'                                                grdConvenio.TextMatrix(i, j) = " "
'                                            Else
'                                                grdConvenio.TextMatrix(i, j) = c
'                                            End If
'                                    Case 6
'                                           If Trim(c) = "" Then
'                                               Mensaje = Mensaje & "campo 'servicio' vacio " & Space(10) & j & ";"
'                                               grdConvenio.TextMatrix(i, j) = " "
'                                            Else
'                                                grdConvenio.TextMatrix(i, j) = c
'                                            End If
'                                    Case 7
'                                           If Trim(c) = "" Then
'                                                Mensaje = Mensaje & "Campo 'Concepto' vacio " & Space(10) & j & ";"
'                                                grdConvenio.TextMatrix(i, j) = " "
'                                           Else
'                                                grdConvenio.TextMatrix(i, j) = c
'                                           End If
'                                    Case 8
'                                           If Trim(c) = "" Then
'                                                If UCase(sTipoValidacion) = "C" Then
'                                                    Mensaje = Mensaje & "Campo 'Importe' vacio " & Space(10) & j & ";"
'                                                    grdConvenio.TextMatrix(i, j) = " "
'                                                Else
'                                                    grdConvenio.TextMatrix(i, j) = " "
'                                                End If
'                                           Else
'                                                grdConvenio.TextMatrix(i, j) = Format(c, "#,##0.00")
'
'                                           End If
'                                End Select
'
'                                j = j + 1
'
'                            Next
'
'                    Case "P"
'                            For Each c In campos
'
'                                Select Case j
'
'                                    Case 1
'                                            If Trim(c) = "" Then
'                                                Mensaje = Mensaje & "Campo 'ID' vacio " & Space(10) & j & ";"
'                                                grdConvenio.TextMatrix(i, j) = " "
'                                            Else
'                                                grdConvenio.TextMatrix(i, j) = c
'                                            End If
'                                    Case 7
'                                           If Trim(c) = "" Then
'                                                Mensaje = Mensaje & "Campo 'Concepto' se encuentra vacio " & Space(10) & j & ";"
'                                                grdConvenio.TextMatrix(i, j) = " "
'                                           Else
'                                                grdConvenio.TextMatrix(i, j) = c
'                                           End If
'                                    Case 8
'                                           If Trim(c) = "" Then
'                                                Mensaje = Mensaje & "Campos 'Importe' vacio " & Space(10) & j & ";"
'                                                grdConvenio.TextMatrix(i, j) = " "
'                                           Else
'                                                grdConvenio.TextMatrix(i, j) = Format(c, "#,##0.00")
'                                           End If
'
'                                End Select
'                                j = j + 1
'                            Next
'                End Select
'
'                If Mensaje <> "" Then
'
'                    Mensaje = Mid(Mensaje, 1, Len(Mensaje) - 1)
'
'                    Call cargarListaError(i, Mensaje)
'
'                End If
'                Mensaje = Empty
'                grdConvenio.AdicionaFila
'                lblTotalRegistros.Caption = i
'
'                If i <= oBarra.Max Then
'                    oBarra.value = i
'                End If
'
'                i = i + 1
'                j = 1
'            Loop
'        Close #1
'
'        i = 1
'        j = 1
'
'        oBarra.Visible = False
'
'        If grdConvenio.Rows - 2 <> nRegistros Then
'               Call cargarListaError(0, _
'               "La cantidad de registros de la trama cabecera no corresponde a la cantidad total de registros de la sección detalles")
'        End If
'
'        If nListError > 0 Then
'            lblerror.Left = 5880
'            lblerror.Top = 7740
'            lblerror.Caption = nListError & " Errores durante la carga"
'            lblerror.Visible = True
'            ' cmdCargar.Enabled = False
'        Else
'            'cmdCargar.Enabled = True
'        End If
'
'        Exit Sub
'
'error:
'    Close #1
'    MsgBox err.Description, vbCritical, "Aviso"
'    limpiar
'
'End Sub

Private Sub Form_Load()
Set rsRecaudoError = New ADODB.Recordset
lblCargando.Visible = False
End Sub

Private Sub fraCargaArchivo_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lblerror_Click()
    
    Dim ofrmCapCargaArchivoError As frmCapCargaArchivoError
    Set ofrmCapCargaArchivoError = New frmCapCargaArchivoError
    
    On Error GoTo error
    
    Call ofrmCapCargaArchivoError.inicia(rsRecaudoError, Trim(lblEmpresa.Caption), lblCodigoConvenio.Caption)
    
    Exit Sub
    
error:
    MsgBox err.Description, vbCritical, "Aviso"
    
End Sub













