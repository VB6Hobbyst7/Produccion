VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form frmArqueoPagare 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arqueo de Pagarés de Créditos"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   Icon            =   "frmArqueoPagare.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Selección de Archivo"
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
      Height          =   1335
      Left            =   7560
      TabIndex        =   21
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   720
         TabIndex        =   25
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscarArchivo 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   24
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   23
         Top             =   720
         Width           =   1170
      End
      Begin VB.TextBox txtNomArchivo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   280
         Width           =   2415
      End
      Begin PryOcxExplorer.OcxCdlgExplorer CdlgFile 
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Filtro          =   "Contratos Digital (*.pdf)|*.pdf"
         Altura          =   0
      End
   End
   Begin VB.Frame fraEstado 
      Caption         =   "Estado"
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
      Left            =   4320
      TabIndex        =   13
      Top             =   600
      Width           =   3135
      Begin VB.OptionButton optEstado 
         Caption         =   "Cancelado"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optEstado 
         Caption         =   "Vigente"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         Height          =   15
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   135
      End
   End
   Begin VB.Frame fraPeriodo 
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
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   3975
      Begin MSComCtl2.DTPicker dtpPerIni 
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   132972545
         CurrentDate     =   36161
      End
      Begin MSComCtl2.DTPicker dtpPerFin 
         Height          =   300
         Left            =   2160
         TabIndex        =   10
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   132972545
         CurrentDate     =   42535
      End
      Begin VB.Label Label5 
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "-"
         Height          =   15
         Left            =   1560
         TabIndex        =   11
         Top             =   480
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdCancelar 
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
      Height          =   375
      Left            =   9480
      TabIndex        =   5
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdRegArqueo 
      Caption         =   "Registrar Arqueo"
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
      Left            =   7560
      TabIndex        =   4
      Top             =   6720
      Width           =   1700
   End
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
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
      TabIndex        =   2
      Top             =   5880
      Width           =   10695
      Begin VB.TextBox txtGlosa 
         Height          =   405
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   10455
      End
   End
   Begin SICMACT.FlexEdit feArqueoPagare 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7646
      Cols0           =   8
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Tipo de Crédito-Cant. Sistema-Cant. Físico Agencia-Cant. Físico BN-Cant. Físico-Detalle-cTpoCredCod"
      EncabezadosAnchos=   "300-3400-1200-1700-1200-1200-1200-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-6-X"
      ListaControles  =   "0-0-0-0-0-0-1-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-L-L-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      TipoBusqueda    =   6
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin MSComctlLib.ProgressBar BarraProgreso 
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   6720
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblUsuSupValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4320
      TabIndex        =   20
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblAgeDesc 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   240
      Width           =   3435
   End
   Begin VB.Label lblFechaValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   6000
      TabIndex        =   18
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   17
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label lblAgencia 
      Caption         =   "Agencia:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblAgeCod 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
      Height          =   315
      Left            =   3720
      TabIndex        =   6
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblUsuSupField 
      Caption         =   "Usuario Arqueo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmArqueoPagare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'** Nombre : frmArqueoPagare
'** Descripción : Formulario que realiza el arqueo de pagares de creditos
'** Creación : marg, 20160610
'** Referencia : TI-ERS041-2016
'***************************************************************************
Option Explicit
Dim bResultadoVisto As Boolean
Dim oVisto As frmVistoElectronico
Dim cUsuVisto As String
Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral

Dim nMatDetFaltante() As Variant '' excluir
Dim MatPagareFaltante As Variant

Dim MatPagares() As TPagare
Private Type TPagare
    TipoCredito As String
    CantSistema As String
    CantFisicoAge As String
    CantFisicoBN As String
    CantFisico As String
    CodTipoCredito As String
End Type
Dim MatConteo() As TConteo
Private Type TConteo
    contEsFaltante As Boolean
    contEsDesembolsoBN As Boolean
    
    contCredito As String
    contCreditoAntiguo As String
    contCliente As String
    contMoneda As String
    contMAprobado As String
    contSCapital As String
    contDesembolo As String
    contTasa As String
    contAnalista As String
    contAtraso As String
    contTipoCredito As String
    contCondicion As String
    contEstado As String
    contDesembolsoBN As String
    contOficinaBN As String
    contDistrito As String
    contZona As String
    contDireccion As String
    contCodTipoCred As String
    contNroc As String
    contPlazo As String
    contVigencia As String
    contCalifInterna As String
    contNroEntidades As String
    contDirectoSoles As String
    contDirectoDolares As String
    contIndirectoSoles As String
    contIndirectoDolares As String
    contCalifSistFinan As String
    contMDesembolso As String
    
End Type

Dim MatArqueoPagareCredito() As TArqueoPagareCredito
Private Type TArqueoPagareCredito
    nIdArqueoPagCred As Integer
    nIdArqueoPag As Integer
    cTpoCredCod As String
    nCantSistema As Integer
    nCantFisicoAgen As Integer
    nCantFisicoBN As Integer
    nCantFisico As Integer
    cConsDescripcion As String
End Type
Dim MatArqueoPagareCreditoDet() As TArqueoPagareCreditoDet
Private Type TArqueoPagareCreditoDet
    nIdArqueoPagCredDet As Integer
    nIdArqueoPagCred As Integer
    cNumeroPagare As String
    nFaltante As Integer
    cCreditoAntiguo As String
    cCliente As String
    cMoneda As String
    cMAprobado As String
    cSCapital As String
    cDesembolo As String
    cTasa As String
    cAnalista As String
    cAtraso As String
    cTipoCredito As String
    cCondicion As String
    cEstado As String
    cDesembolsoBN As String
    cOficinaBN As String
    cDistrito As String
    cZona As String
    cDireccion As String
    cCodTipoCred As String
    cNroc As String
    cPlazo As String
    cVigencia As String
    cCalifInterna As String
    cNroEntidades As String
    cDirectoSoles As String
    cDirectoDolares As String
    cIndirectoSoles As String
    cIndirectoDolares As String
    cCalifSistFinan As String
    cMDesembolso As String
End Type

Dim nEstado As Integer
Dim CodTipoCred As String
Dim cFechaIni As String ''fecha de consulta
Dim cFechaFin As String ''fecha de consulta
Dim FechaIni As String ''fecha de registro
Dim FechaFin As String ''fecha de registro

Dim esCompleto As Boolean
Dim rsArqueoPagare As ADODB.Recordset
Dim rsArqueoPagareCredito As ADODB.Recordset
Dim rsArqueoPagareCreditoDet As ADODB.Recordset

Private fsPathFile As String
Private fsRuta As String
Private fsNomFile As String
Private nExcel As Integer
Private nFlex As Integer
Private i As Integer

Public Sub Inicia()
Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
If oCaja.PermitirArqueoPagare(gsCodUser, gsCodAge, gsCodArea, gsCodCargo) Then
    Me.Show 1
Else
    MsgBox "Ud. No puede realizar el Arqueo", vbInformation, "Mensaje"
End If
End Sub


Private Sub cmdBuscarArchivo_Click()
LimpiaFlex feArqueoPagare
Dim i As Integer
CdlgFile.nHwd = Me.hwnd
CdlgFile.Filtro = "Archivos Excel (*.xls)|*.xls"
CdlgFile.altura = 300
CdlgFile.Show

fsPathFile = CdlgFile.Ruta
fsRuta = fsPathFile
        If fsPathFile <> Empty Then
            For i = Len(fsPathFile) - 1 To 1 Step -1
                    If Mid(fsPathFile, i, 1) = "\" Then
                        fsPathFile = Mid(CdlgFile.Ruta, 1, i)
                        fsNomFile = Mid(CdlgFile.Ruta, i + 1, Len(CdlgFile.Ruta) - i)
                        Exit For
                    End If
             Next i
          Screen.MousePointer = 11
          txtNomArchivo.Text = fsNomFile
        Else
           MsgBox "No se Selecciono Ningun Archivo", vbInformation, "Aviso"
           txtNomArchivo.Text = ""
           LimpiaFlex feArqueoPagare
           Exit Sub
        End If
    Screen.MousePointer = 0
End Sub
Private Function ValidaDatos() As Boolean
If Trim(txtNomArchivo.Text) = "" Then
    MsgBox "Seleccione el Archivo a cargar", vbInformation, "Aviso"
    Exit Function
    ValidaDatos = False
End If

ValidaDatos = True
End Function
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCargar_Click()
If Not ValidaDatos Then Exit Sub
If MsgBox("Estas seguro de cargar el archivo adjuntado?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
On Error GoTo ErrorCargaArchivo

cmdRegArqueo.Enabled = False

Dim oConstante As COMDConstantes.DCOMConstantes
Dim rsTipoCredito As ADODB.Recordset

Set oConstante = New COMDConstantes.DCOMConstantes
Set rsTipoCredito = oConstante.RecuperaConstantes(3034)
Set oConstante = Nothing

BarraProgreso.Visible = True
BarraProgreso.value = 0
BarraProgreso.Min = 0
BarraProgreso.value = 0
    
''Variables de Arqueo
Dim TipoCredito As String
Dim CantSistema As Integer
Dim CantFisicoAgencia As Integer
Dim CantFisicoBN As Integer
Dim CantFisico As Integer
Dim EsFaltante As Boolean
Dim EsDesembolsoBN As Boolean
''Variables de Pagares vigentes faltantes
Dim lsCredito As String
Dim lsCreditoAntiguo As String
Dim lsCliente As String
Dim lsMoneda As String
Dim lsMAprobado As String
Dim lsSCapital As String
Dim lsDesembolo As String
Dim lsTasa As String
Dim lsAnalista As String
Dim lsAtraso As String
Dim lsTipoCredito As String
Dim lsCondicion As String
Dim lsEstado As String
Dim lsDesembolsoBN As String
Dim lsOficinaBN As String
Dim lsDistrito As String
Dim lsZona As String
Dim lsDireccion As String
Dim lsNroc As String
Dim lsPlazo As String
Dim lsVigencia As String
Dim lsCalifInterna As String
Dim lsNroEntidades As String
Dim lsDirectoSoles As String
Dim lsDirectoDolares As String
Dim lsIndirectoSoles As String
Dim lsIndirectoDolares As String
Dim lsCalifSistFinan As String
Dim lsMDesembolso As String

Dim xlsAplicacion As Excel.Application
Dim xlsLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim lsHoja As String
Dim lbExisteHoja As Boolean
Dim lbHayDatos As Boolean

Set xlsAplicacion = New Excel.Application
Set xlsLibro = xlsAplicacion.Workbooks.Open(fsRuta)

If getOptEstado = 1 Then
    lsHoja = "108203 Creditos Vigentes Arqueo"
Else
    lsHoja = "108109 Creditos cancelados"
End If

''Activa la hoja correspondiente
For Each xlHoja In xlsLibro.Worksheets
   If UCase(Trim(xlHoja.Name)) = UCase(Trim(lsHoja)) Then
        xlHoja.Activate
        lbExisteHoja = True
    Exit For
   End If
Next


If lbExisteHoja = False Then
    MsgBox "El Nombre de la Hoja debe ser ''" & lsHoja & "''", vbCritical, "Aviso"
    xlsAplicacion.Quit
    xlsAplicacion.Visible = False
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja = Nothing
    Exit Sub
End If

Dim cellFecha As String
If nEstado = 1 Then
    cellFecha = Trim(xlHoja.Cells(2, 4))
    FechaIni = Mid(cellFecha, 41, 10)
    FechaFin = Mid(cellFecha, 55, 10)
    nExcel = 7
Else
    cellFecha = Trim(xlHoja.Cells(2, 3))
    FechaIni = Mid(cellFecha, 34, 10)
    FechaFin = Mid(cellFecha, 48, 10)
    nExcel = 0
End If
dtpPerIni.value = CDate(FechaIni)
dtpPerFin.value = CDate(FechaFin)

Dim lbHayProducto As Boolean
Dim lbHayCreditoSiguiente As Boolean
Dim lbEsUltimoProducto As Boolean
Dim nFilaProducto As Integer
nFlex = 1
lbHayDatos = True
lbHayProducto = False
lbHayCreditoSiguiente = False
lbEsUltimoProducto = False
nFilaProducto = 4
Do While lbHayDatos
    
    ''CREDITOS VIGENTES
    If nEstado = 1 Then
        EsFaltante = IIf(Trim(xlHoja.Cells(nExcel, 5)) = "", False, True)
        TipoCredito = UCase(Trim(xlHoja.Cells(nExcel, 17)))
        EsDesembolsoBN = IIf(Trim(xlHoja.Cells(nExcel, 20)) = "NO", False, True)
        If EsFaltante Then
            lsCredito = Trim(xlHoja.Cells(nExcel, 2))
            lsCreditoAntiguo = Trim(xlHoja.Cells(nExcel, 6))
            lsCliente = Trim(xlHoja.Cells(nExcel, 7))
            lsMoneda = Trim(xlHoja.Cells(nExcel, 9))
            lsMAprobado = Trim(xlHoja.Cells(nExcel, 10))
            lsSCapital = Trim(xlHoja.Cells(nExcel, 11))
            lsDesembolo = Trim(xlHoja.Cells(nExcel, 12))
            lsTasa = Trim(xlHoja.Cells(nExcel, 14))
            lsAnalista = Trim(xlHoja.Cells(nExcel, 15))
            lsAtraso = Trim(xlHoja.Cells(nExcel, 16))
            lsTipoCredito = Trim(xlHoja.Cells(nExcel, 17))
            lsCondicion = Trim(xlHoja.Cells(nExcel, 18))
            lsEstado = Trim(xlHoja.Cells(nExcel, 19))
            lsDesembolsoBN = Trim(xlHoja.Cells(nExcel, 20))
            lsOficinaBN = Trim(xlHoja.Cells(nExcel, 21))
            lsDistrito = Trim(xlHoja.Cells(nExcel, 22))
            lsZona = Trim(xlHoja.Cells(nExcel, 23))
            lsDireccion = Trim(xlHoja.Cells(nExcel, 24))
        End If
        
        If Trim(TipoCredito) = "" Then
            lbHayDatos = False
            Exit Do
        End If
        
        ReDim Preserve MatConteo(nFlex)
    
        MatConteo(nFlex).contEsFaltante = EsFaltante
        MatConteo(nFlex).contTipoCredito = TipoCredito
        MatConteo(nFlex).contEsDesembolsoBN = EsDesembolsoBN
        If EsFaltante Then
                MatConteo(nFlex).contCredito = lsCredito
                MatConteo(nFlex).contCreditoAntiguo = lsCreditoAntiguo
                MatConteo(nFlex).contCliente = lsCliente
                MatConteo(nFlex).contMoneda = lsMoneda
                MatConteo(nFlex).contMAprobado = lsMAprobado
                MatConteo(nFlex).contSCapital = lsSCapital
                MatConteo(nFlex).contDesembolo = lsDesembolo
                MatConteo(nFlex).contTasa = lsTasa
                MatConteo(nFlex).contAnalista = lsAnalista
                MatConteo(nFlex).contAtraso = lsAtraso
                MatConteo(nFlex).contTipoCredito = lsTipoCredito
                MatConteo(nFlex).contCondicion = lsCondicion
                MatConteo(nFlex).contEstado = lsEstado
                MatConteo(nFlex).contDesembolsoBN = lsDesembolsoBN
                MatConteo(nFlex).contOficinaBN = lsOficinaBN
                MatConteo(nFlex).contDistrito = lsDistrito
                MatConteo(nFlex).contZona = lsZona
                MatConteo(nFlex).contDireccion = lsDireccion
        End If
        lbHayDatos = True
        nExcel = nExcel + 1
        nFlex = nFlex + 1
    End If
    
    ''CREDITOS CANCELADOS
    If nEstado = 0 Then
        
        lbHayProducto = IIf(UCase(Mid(Trim(xlHoja.Cells(nFilaProducto, 1)), 1, 8)) = "PRODUCTO", True, False)
        If lbHayProducto Then
            If lbHayCreditoSiguiente Then
                nExcel = nExcel + 1
            Else
                 nExcel = nFilaProducto + 3
            End If
           
            EsFaltante = IIf(Trim(xlHoja.Cells(nExcel, 5)) = "", False, True)
            TipoCredito = UCase(Mid(Trim(xlHoja.Cells(nFilaProducto, 1)), 11))
            EsDesembolsoBN = False
            If EsFaltante Then
                lsCredito = Trim(xlHoja.Cells(nExcel, 1))
                lsCreditoAntiguo = Trim(xlHoja.Cells(nExcel, 4))
                lsCondicion = Trim(xlHoja.Cells(nExcel, 6))
                lsEstado = Trim(xlHoja.Cells(nExcel, 7))
                lsCliente = Trim(xlHoja.Cells(nExcel, 8))
                lsMDesembolso = Trim(xlHoja.Cells(nExcel, 9))
                lsNroc = Trim(xlHoja.Cells(nExcel, 11))
                lsPlazo = Trim(xlHoja.Cells(nExcel, 12))
                lsVigencia = Trim(xlHoja.Cells(nExcel, 13))
                lsAnalista = Trim(xlHoja.Cells(nExcel, 14))
                lsCalifInterna = Trim(xlHoja.Cells(nExcel, 15))
                lsNroEntidades = Trim(xlHoja.Cells(nExcel, 16))
                lsDirectoSoles = Trim(xlHoja.Cells(nExcel, 17))
                lsDirectoDolares = Trim(xlHoja.Cells(nExcel, 18))
                lsIndirectoSoles = Trim(xlHoja.Cells(nExcel, 19))
                lsIndirectoDolares = Trim(xlHoja.Cells(nExcel, 20))
                lsCalifSistFinan = Trim(xlHoja.Cells(nExcel, 20))
            End If
                   
            ReDim Preserve MatConteo(nFlex)
    
            MatConteo(nFlex).contEsFaltante = EsFaltante
            MatConteo(nFlex).contTipoCredito = TipoCredito
            MatConteo(nFlex).contEsDesembolsoBN = EsDesembolsoBN
            If EsFaltante Then
                    MatConteo(nFlex).contCredito = lsCredito
                    MatConteo(nFlex).contCreditoAntiguo = lsCreditoAntiguo
                    MatConteo(nFlex).contCondicion = lsCondicion
                    MatConteo(nFlex).contEstado = lsEstado
                    MatConteo(nFlex).contCliente = lsCliente
                    MatConteo(nFlex).contMDesembolso = lsMDesembolso
                    MatConteo(nFlex).contNroc = lsNroc
                    MatConteo(nFlex).contPlazo = lsPlazo
                    MatConteo(nFlex).contVigencia = lsVigencia
                    MatConteo(nFlex).contAnalista = lsAnalista
                    MatConteo(nFlex).contCalifInterna = lsCalifInterna
                    MatConteo(nFlex).contNroEntidades = lsNroEntidades
                    MatConteo(nFlex).contDirectoSoles = lsDirectoSoles
                    MatConteo(nFlex).contDirectoDolares = lsDirectoDolares
                    MatConteo(nFlex).contIndirectoSoles = lsIndirectoSoles
                    MatConteo(nFlex).contIndirectoDolares = lsIndirectoDolares
                    MatConteo(nFlex).contCalifSistFinan = lsCalifSistFinan
            End If
            lbHayCreditoSiguiente = IIf(Mid(Trim(xlHoja.Cells(nExcel + 1, 1)), 1, 8) = "", False, True)
            lbEsUltimoProducto = IIf(Trim(xlHoja.Cells(nExcel + 2, 1)) = "", True, False)
            If lbHayCreditoSiguiente Then
                lbHayDatos = True
            End If
            If Not lbHayCreditoSiguiente Then
                If lbEsUltimoProducto Then
                    lbHayDatos = False
                    Exit Do
                End If
                If Not lbEsUltimoProducto Then
                    nFilaProducto = nExcel + 2
                    lbHayDatos = True
                End If
            End If
            nFlex = nFlex + 1
        End If
        If Not lbHayProducto Then
            nFilaProducto = nFilaProducto + 1
        End If
    End If
Loop


Dim iMat As Long
Dim realizoConteo As Boolean

If nEstado = 1 Then
    ReDim MatPagareFaltante(19, 0)
End If
If nEstado = 0 Then
    ReDim MatPagareFaltante(19, 0)
End If
ReDim Preserve MatPagares(0)

''CUENTA LOS PAGARES POR TIPO DE CREDITO
Do While Not rsTipoCredito.EOF
    BarraProgreso.Max = UBound(MatConteo)
    For i = 1 To UBound(MatConteo)
        If UCase(Trim(MatConteo(i).contTipoCredito)) = UCase(Trim(rsTipoCredito!cConsDescripcion)) Then
            '<CONTEO FISICO DE PAGARES EXISTENTES>
            CantSistema = CantSistema + 1
            If Not MatConteo(i).contEsFaltante Then
                CantFisico = CantFisico + 1
                If Not MatConteo(i).contEsDesembolsoBN Then
                    CantFisicoAgencia = CantFisicoAgencia + 1
                Else
                    CantFisicoBN = CantFisicoBN + 1
                End If
            End If
            '</>
            '<REGISTRO DE BITACORA DE PAGARES FALTANTES
            'VIGENTES
            If nEstado = 1 Then
                If MatConteo(i).contEsFaltante Then
                   iMat = UBound(MatPagareFaltante, 2) + 1
                   ReDim Preserve MatPagareFaltante(19, 0 To iMat)
                   MatPagareFaltante(1, iMat) = MatConteo(i).contCredito
                   MatPagareFaltante(2, iMat) = MatConteo(i).contCreditoAntiguo
                   MatPagareFaltante(3, iMat) = MatConteo(i).contCliente
                   MatPagareFaltante(4, iMat) = MatConteo(i).contMoneda
                   MatPagareFaltante(5, iMat) = MatConteo(i).contMAprobado
                   MatPagareFaltante(6, iMat) = MatConteo(i).contSCapital
                   MatPagareFaltante(7, iMat) = MatConteo(i).contDesembolo
                   MatPagareFaltante(8, iMat) = MatConteo(i).contTasa
                   MatPagareFaltante(9, iMat) = MatConteo(i).contAnalista
                   MatPagareFaltante(10, iMat) = MatConteo(i).contAtraso
                   MatPagareFaltante(11, iMat) = MatConteo(i).contTipoCredito
                   MatPagareFaltante(12, iMat) = MatConteo(i).contCondicion
                   MatPagareFaltante(13, iMat) = MatConteo(i).contEstado
                   MatPagareFaltante(14, iMat) = MatConteo(i).contDesembolsoBN
                   MatPagareFaltante(15, iMat) = MatConteo(i).contOficinaBN
                   MatPagareFaltante(16, iMat) = MatConteo(i).contDistrito
                   MatPagareFaltante(17, iMat) = MatConteo(i).contZona
                   MatPagareFaltante(18, iMat) = MatConteo(i).contDireccion
                   MatPagareFaltante(19, iMat) = rsTipoCredito!nConsValor ''codigo tipo credito
                End If
            End If
            'CANCELADOS
            If nEstado = 0 Then
                If MatConteo(i).contEsFaltante Then
                    iMat = UBound(MatPagareFaltante, 2) + 1
                    ReDim Preserve MatPagareFaltante(19, 0 To iMat)
                    MatPagareFaltante(1, iMat) = MatConteo(i).contCredito
                    MatPagareFaltante(2, iMat) = MatConteo(i).contCreditoAntiguo
                    MatPagareFaltante(3, iMat) = MatConteo(i).contCondicion
                    MatPagareFaltante(4, iMat) = MatConteo(i).contEstado
                    MatPagareFaltante(5, iMat) = MatConteo(i).contCliente
                    MatPagareFaltante(6, iMat) = MatConteo(i).contTipoCredito
                    MatPagareFaltante(7, iMat) = MatConteo(i).contMDesembolso
                    MatPagareFaltante(8, iMat) = MatConteo(i).contNroc
                    MatPagareFaltante(9, iMat) = MatConteo(i).contPlazo
                    MatPagareFaltante(10, iMat) = MatConteo(i).contVigencia
                    MatPagareFaltante(11, iMat) = MatConteo(i).contAnalista
                    MatPagareFaltante(12, iMat) = MatConteo(i).contCalifInterna
                    MatPagareFaltante(13, iMat) = MatConteo(i).contNroEntidades
                    MatPagareFaltante(14, iMat) = MatConteo(i).contDirectoSoles
                    MatPagareFaltante(15, iMat) = MatConteo(i).contDirectoDolares
                    MatPagareFaltante(16, iMat) = MatConteo(i).contIndirectoSoles
                    MatPagareFaltante(17, iMat) = MatConteo(i).contIndirectoDolares
                    MatPagareFaltante(18, iMat) = MatConteo(i).contCalifSistFinan
                    MatPagareFaltante(19, iMat) = rsTipoCredito!nConsValor ''codigo tipo credito
                End If
            End If
            realizoConteo = True
        End If
        BarraProgreso.value = i
    Next
    If realizoConteo Then
        ReDim Preserve MatPagares(UBound(MatPagares) + 1)
        MatPagares(UBound(MatPagares)).TipoCredito = rsTipoCredito!cConsDescripcion
        MatPagares(UBound(MatPagares)).CantSistema = CantSistema
        MatPagares(UBound(MatPagares)).CantFisicoAge = CantFisicoAgencia
        MatPagares(UBound(MatPagares)).CantFisicoBN = CantFisicoBN
        MatPagares(UBound(MatPagares)).CantFisico = CantFisico
        MatPagares(UBound(MatPagares)).CodTipoCredito = rsTipoCredito!nConsValor
    End If
    CantSistema = 0
    CantFisico = 0
    CantFisicoAgencia = 0
    CantFisicoBN = 0
    realizoConteo = False
    rsTipoCredito.MoveNext
Loop

''MUESTRA LOS PAGARES CONTABILIZADOS POR TIPO DE CREDITO
LimpiaFlex Me.feArqueoPagare
For i = 1 To UBound(MatPagares)
    feArqueoPagare.AdicionaFila
    feArqueoPagare.TextMatrix(i, 1) = MatPagares(i).TipoCredito
    feArqueoPagare.TextMatrix(i, 2) = MatPagares(i).CantSistema
    feArqueoPagare.TextMatrix(i, 3) = MatPagares(i).CantFisicoAge
    feArqueoPagare.TextMatrix(i, 4) = MatPagares(i).CantFisicoBN
    feArqueoPagare.TextMatrix(i, 5) = MatPagares(i).CantFisico
    feArqueoPagare.TextMatrix(i, 6) = "OK" 'adicional
    feArqueoPagare.TextMatrix(i, 7) = MatPagares(i).CodTipoCredito
    
Next
MsgBox "Datos Cargados Correctamente", vbInformation, "Aviso"
BarraProgreso.Visible = False
cmdRegArqueo.Enabled = True

feArqueoPagare.TopRow = 1
xlsAplicacion.Quit
    
xlsAplicacion.Visible = False
Set xlsAplicacion = Nothing
Set xlsLibro = Nothing
Set xlHoja = Nothing
Exit Sub

ErrorCargaArchivo:
xlsAplicacion.Quit
Set xlsAplicacion = Nothing
Set xlsLibro = Nothing
Set xlHoja = Nothing

MsgBox Err.Description & " - Carga de Archivo incorrecto", vbCritical, "Aviso"

End Sub

Private Function getOptEstado() As Integer
 If optEstado(1).value = True Then
        nEstado = 1
    End If
    If optEstado(0).value = True Then
        nEstado = 0
    End If
    getOptEstado = nEstado
End Function
Private Sub cmdConsultar_Click()
    If optEstado(1).value = True Then
        nEstado = 1
    End If
    If optEstado(0).value = True Then
        nEstado = 0
    End If
    ConsultarDatos nEstado
End Sub

Private Sub GenerarReporteArqueoPagares()

Dim fs As Scripting.FileSystemObject
Dim lsArchivo As String
Dim lsArchivoMostrar As String

Dim lsRuta As String

Dim xlsAplicacion As Excel.Application
Dim xlsLibro As Excel.Workbook

    lsArchivo = "ArqueoPagares"
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    
    lsRuta = App.Path & "\FormatoCarta\" & lsArchivo & ".xlsx"

    
    lsArchivoMostrar = "\Spooler\" & lsArchivo & "_" & gsCodUser & "_" & gsCodAge & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xlsx"
    If fs.FileExists(lsRuta) Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(lsRuta)
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    lsArchivoMostrar = App.Path & lsArchivoMostrar
    
    xlsLibro.SaveAs lsArchivoMostrar

'-------------exportar-----
Dim xlsHoja As Excel.Worksheet
Dim xlsHojaDel As Excel.Worksheet
Dim lsHojaDescarte As String

Dim lbExisteHoja As Boolean
Dim i As Long
Dim j As Long
Dim ldFecha As Date
Dim lnExcel As Integer
Dim x1 As Integer
Dim x2 As Integer
Dim x3 As Integer
Dim x4 As Integer

On Error GoTo ErrExportarExcelError

'Activa la hoja correspondiente
For Each xlsHoja In xlsLibro.Worksheets
   If UCase(Trim(xlsHoja.Name)) = UCase(Trim("HOJA DE TRABAJO")) Then
        xlsHoja.Activate
    Exit For
   End If
Next


xlsHoja.Range("B2") = IIf(nEstado = 1, "ARQUEO DE PAGARES VIGENTES DEL " & Format(CDate(FechaIni), "dd/mm/YYYY") & " AL " & Format(CDate(FechaFin), "dd/mm/YYYY"), "ARQUEO DE PAGARES CANCELADOS DEL " & Format(CDate(FechaIni), "dd/mm/YYYY") & " AL " & Format(CDate(FechaFin), "dd/mm/YYYY"))
xlsHoja.Cells(5, 3) = UCase(gsNomAge)
xlsHoja.Cells(5, 4) = UCase(gsNomAge)
xlsHoja.Cells(5, 5) = "NACION"
xlsHoja.Cells(5, 6) = UCase(gsNomAge)

lnExcel = 6
lnExcel = lnExcel + 1
For i = 1 To UBound(MatArqueoPagareCredito)
    xlsHoja.Cells(lnExcel, 2) = MatArqueoPagareCredito(i).cConsDescripcion
    xlsHoja.Cells(lnExcel, 3) = MatArqueoPagareCredito(i).nCantSistema
    x1 = x1 + MatArqueoPagareCredito(i).nCantSistema
    xlsHoja.Cells(lnExcel, 4) = MatArqueoPagareCredito(i).nCantFisicoAgen
    x2 = x2 + MatArqueoPagareCredito(i).nCantFisicoAgen
    xlsHoja.Cells(lnExcel, 5) = MatArqueoPagareCredito(i).nCantFisicoBN
    x3 = x3 + MatArqueoPagareCredito(i).nCantFisicoBN
    xlsHoja.Cells(lnExcel, 6) = MatArqueoPagareCredito(i).nCantFisico
    x4 = x4 + MatArqueoPagareCredito(i).nCantFisico
    xlsHoja.Range("B" & lnExcel + 1, "F" & lnExcel + 1).EntireRow.Insert
    lnExcel = lnExcel + 1
Next
xlsHoja.Range("B" & lnExcel, "F" & lnExcel).EntireRow.Delete

xlsHoja.Cells(lnExcel, 3) = x1
xlsHoja.Cells(lnExcel, 4) = x2
xlsHoja.Cells(lnExcel, 5) = x3
xlsHoja.Cells(lnExcel, 6) = x4

If nEstado = 1 Then
    'Activa la hoja correspondiente
    For Each xlsHoja In xlsLibro.Worksheets
       If UCase(Trim(xlsHoja.Name)) = UCase(Trim("DETALLE DE VIGENTES FALTANTES")) Then
            xlsHoja.Activate
            lsHojaDescarte = UCase(Trim("DETALLE DE CANCELADOS FALTANTES"))
        Exit For
       End If
    Next
End If

If nEstado = 0 Then
    'Activa la hoja correspondiente
    For Each xlsHoja In xlsLibro.Worksheets
       If UCase(Trim(xlsHoja.Name)) = UCase(Trim("DETALLE DE CANCELADOS FALTANTES")) Then
            xlsHoja.Activate
            lsHojaDescarte = UCase(Trim("DETALLE DE VIGENTES FALTANTES"))
        Exit For
       End If
    Next
End If

'Oculta lo que no corresponde
For Each xlsHojaDel In xlsLibro.Worksheets
    If UCase(Trim(xlsHojaDel.Name)) = UCase(Trim(lsHojaDescarte)) Then
        xlsHojaDel.Visible = xlSheetHidden
    End If
Next


If nEstado = 1 Then
    xlsHoja.Range("B2") = "DETALLE DE PAGARES VIGENTES FALTANTES DEL " & Format(CDate(FechaIni), "dd/mm/YYYY") & " AL " & Format(CDate(FechaFin))
    lnExcel = 4
    lnExcel = lnExcel + 1
    For j = 1 To UBound(MatArqueoPagareCreditoDet)
        xlsHoja.Cells(lnExcel, 2) = MatArqueoPagareCreditoDet(j).cNumeroPagare
        xlsHoja.Cells(lnExcel, 3) = MatArqueoPagareCreditoDet(j).cCreditoAntiguo
        xlsHoja.Cells(lnExcel, 4) = MatArqueoPagareCreditoDet(j).cCliente
        xlsHoja.Cells(lnExcel, 5) = MatArqueoPagareCreditoDet(j).cMoneda
        xlsHoja.Cells(lnExcel, 6) = MatArqueoPagareCreditoDet(j).cMAprobado
        xlsHoja.Cells(lnExcel, 7) = MatArqueoPagareCreditoDet(j).cSCapital
        xlsHoja.Cells(lnExcel, 8) = MatArqueoPagareCreditoDet(j).cDesembolo
        xlsHoja.Cells(lnExcel, 9) = MatArqueoPagareCreditoDet(j).cTasa
        xlsHoja.Cells(lnExcel, 10) = MatArqueoPagareCreditoDet(j).cAnalista
        xlsHoja.Cells(lnExcel, 11) = MatArqueoPagareCreditoDet(j).cAtraso
        xlsHoja.Cells(lnExcel, 12) = MatArqueoPagareCreditoDet(j).cTipoCredito
        xlsHoja.Cells(lnExcel, 13) = MatArqueoPagareCreditoDet(j).cCondicion
        xlsHoja.Cells(lnExcel, 14) = MatArqueoPagareCreditoDet(j).cEstado
        xlsHoja.Cells(lnExcel, 15) = MatArqueoPagareCreditoDet(j).cDesembolsoBN
        xlsHoja.Cells(lnExcel, 16) = MatArqueoPagareCreditoDet(j).cOficinaBN
        xlsHoja.Cells(lnExcel, 17) = MatArqueoPagareCreditoDet(j).cDistrito
        xlsHoja.Cells(lnExcel, 18) = MatArqueoPagareCreditoDet(j).cZona
        xlsHoja.Cells(lnExcel, 19) = MatArqueoPagareCreditoDet(j).cDireccion
        
        xlsHoja.Range("B" & lnExcel + 1, "C" & lnExcel + 1).EntireRow.Insert
        lnExcel = lnExcel + 1
    Next
    xlsHoja.Range("B" & lnExcel, "C" & lnExcel).EntireRow.Delete
    xlsHoja.Range("B" & lnExcel + 1, "C" & lnExcel + 1).EntireRow.Delete
End If
If nEstado = 0 Then
    ''IMPLEMENTAR
    xlsHoja.Range("B2") = "DETALLE DE PAGARES CANCELADOS FALTANTES DEL " & Format(CDate(FechaIni), "dd/mm/YYYY") & " AL " & Format(CDate(FechaFin))
    lnExcel = 4
    lnExcel = lnExcel + 1
    For j = 1 To UBound(MatArqueoPagareCreditoDet)
        xlsHoja.Cells(lnExcel, 2) = MatArqueoPagareCreditoDet(j).cNumeroPagare
        xlsHoja.Cells(lnExcel, 3) = MatArqueoPagareCreditoDet(j).cCreditoAntiguo
        xlsHoja.Cells(lnExcel, 4) = MatArqueoPagareCreditoDet(j).cCondicion
        xlsHoja.Cells(lnExcel, 5) = MatArqueoPagareCreditoDet(j).cEstado
        xlsHoja.Cells(lnExcel, 6) = MatArqueoPagareCreditoDet(j).cCliente
        xlsHoja.Cells(lnExcel, 7) = MatArqueoPagareCreditoDet(j).cTipoCredito
        xlsHoja.Cells(lnExcel, 8) = MatArqueoPagareCreditoDet(j).cNroc
        xlsHoja.Cells(lnExcel, 9) = MatArqueoPagareCreditoDet(j).cPlazo
        xlsHoja.Cells(lnExcel, 10) = MatArqueoPagareCreditoDet(j).cVigencia
        xlsHoja.Cells(lnExcel, 11) = MatArqueoPagareCreditoDet(j).cAnalista
        xlsHoja.Cells(lnExcel, 12) = MatArqueoPagareCreditoDet(j).cCalifInterna
        xlsHoja.Cells(lnExcel, 13) = MatArqueoPagareCreditoDet(j).cNroEntidades
        xlsHoja.Cells(lnExcel, 14) = MatArqueoPagareCreditoDet(j).cDirectoSoles
        xlsHoja.Cells(lnExcel, 15) = MatArqueoPagareCreditoDet(j).cDirectoDolares
        xlsHoja.Cells(lnExcel, 16) = MatArqueoPagareCreditoDet(j).cIndirectoSoles
        xlsHoja.Cells(lnExcel, 17) = MatArqueoPagareCreditoDet(j).cIndirectoDolares
        xlsHoja.Cells(lnExcel, 18) = MatArqueoPagareCreditoDet(j).cCalifSistFinan
        
        xlsHoja.Range("B" & lnExcel + 1, "C" & lnExcel + 1).EntireRow.Insert
        lnExcel = lnExcel + 1
    Next
    xlsHoja.Range("B" & lnExcel, "C" & lnExcel).EntireRow.Delete
    xlsHoja.Range("B" & lnExcel + 1, "C" & lnExcel + 1).EntireRow.Delete
End If

xlsLibro.SaveAs lsArchivoMostrar
xlsAplicacion.Visible = True
xlsAplicacion.Windows(1).Visible = True
    
Set xlsAplicacion = Nothing
Set xlsLibro = Nothing
Set xlsHoja = Nothing
    

Exit Sub
ErrExportarExcelError:
    MsgBox Err.Description, vbCritical, "Aviso"
    Exit Sub
End Sub


Private Sub cmdRegArqueo_Click()
Dim oDMov As DMov
Dim i As Integer
Dim x As Integer
Dim lsMovNro As String
Dim nIdArqueoPag As Integer
Dim nIdArqueoPagCred As Integer
Dim bTrans As Boolean

Set oDMov = New DMov
On Error GoTo ErrorRegistra
If feArqueoPagare.TextMatrix(1, 1) = "" Then
    MsgBox "No Existen Datos para realizar el Arqueo. ", vbInformation, "Mensaje"
    Exit Sub
End If

If feArqueoPagare.TextMatrix(1, 1) <> "" Then
    For i = 1 To feArqueoPagare.Rows - 1
        If Trim(feArqueoPagare.TextMatrix(i, 6)) = "" Then
            MsgBox "Favor registrar el detalle del registro (" & i & "). Verifique.", vbInformation, "Mensaje"
            Exit Sub
        End If
    Next
End If
    If MsgBox("¿Está seguro de realizar el Arqueo?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    lsMovNro = oDMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    oCaja.dBeginTrans
    bTrans = True
    nIdArqueoPag = oCaja.RegistrarArqueoPagare(lblUsuSupValue.Caption, CDate(lblFechaValue.Caption), lsMovNro, Trim(Replace(Replace((txtGlosa.Text), Chr(10), ""), Chr(13), "")), CDate(FechaIni), CDate(FechaFin), nEstado, 1)
    For i = 1 To feArqueoPagare.Rows - 1
    nIdArqueoPagCred = oCaja.RegistrarArqueoPagareCredito(nIdArqueoPag, MatPagares(i).CodTipoCredito, feArqueoPagare.TextMatrix(i, 2), feArqueoPagare.TextMatrix(i, 3), feArqueoPagare.TextMatrix(i, 4), feArqueoPagare.TextMatrix(i, 5))
        For x = 1 To UBound(MatPagareFaltante, 2)
            If nEstado = 1 Then 'VIGENTES
                If MatPagareFaltante(19, x) = MatPagares(i).CodTipoCredito Then
                oCaja.RegistrarArqueoPagareCreditoDet nIdArqueoPagCred _
                                                    , MatPagareFaltante(1, x) _
                                                    , 1 _
                                                    , MatPagareFaltante(2, x) _
                                                    , MatPagareFaltante(3, x) _
                                                    , MatPagareFaltante(4, x) _
                                                    , MatPagareFaltante(5, x) _
                                                    , MatPagareFaltante(6, x) _
                                                    , MatPagareFaltante(7, x) _
                                                    , MatPagareFaltante(8, x) _
                                                    , MatPagareFaltante(9, x) _
                                                    , MatPagareFaltante(10, x) _
                                                    , MatPagareFaltante(11, x) _
                                                    , MatPagareFaltante(12, x) _
                                                    , MatPagareFaltante(13, x) _
                                                    , MatPagareFaltante(14, x) _
                                                    , MatPagareFaltante(15, x) _
                                                    , MatPagareFaltante(16, x) _
                                                    , MatPagareFaltante(17, x) _
                                                    , MatPagareFaltante(18, x) _
                                                    , MatPagareFaltante(19, x) _
                                                    , "", "", "", "", "", "", "", "", "", "", ""
                End If
            End If
            If nEstado = 0 Then 'CANCELADOS
                If MatPagareFaltante(19, x) = MatPagares(i).CodTipoCredito Then
                oCaja.RegistrarArqueoPagareCreditoDet nIdArqueoPagCred _
                                                    , MatPagareFaltante(1, x) _
                                                    , 1 _
                                                    , MatPagareFaltante(2, x) _
                                                    , MatPagareFaltante(5, x) _
                                                    , "", "", "", "", "" _
                                                    , MatPagareFaltante(11, x) _
                                                    , "" _
                                                    , MatPagareFaltante(6, x) _
                                                    , MatPagareFaltante(3, x) _
                                                    , MatPagareFaltante(4, x) _
                                                    , "", "", "", "", "" _
                                                    , MatPagareFaltante(19, x) _
                                                    , MatPagareFaltante(8, x) _
                                                    , MatPagareFaltante(9, x) _
                                                    , MatPagareFaltante(10, x) _
                                                    , MatPagareFaltante(12, x) _
                                                    , MatPagareFaltante(13, x) _
                                                    , MatPagareFaltante(14, x) _
                                                    , MatPagareFaltante(15, x) _
                                                    , MatPagareFaltante(16, x) _
                                                    , MatPagareFaltante(17, x) _
                                                    , MatPagareFaltante(18, x) _
                                                    , MatPagareFaltante(7, x)
                End If
            End If
        Next
    Next
    oCaja.dCommitTrans
    MsgBox "El Arqueo ha sido realizado correctamente.", vbInformation, "Aviso"
    'GeneraExcel
        ObtenerDatosArqueo
        GenerarReporteArqueoPagares
    bTrans = False
    Unload Me
Exit Sub
ErrorRegistra:
    If bTrans Then
        oCaja.dRollbackTrans
        Set oCaja = Nothing
    End If
    MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error"
End Sub

Private Function getCantFisico(pnRow As Long) As Integer
    Dim cant_fisico As Integer
    Dim cant_age As Integer
    Dim cant_bn As Integer
    If feArqueoPagare.TextMatrix(pnRow, 3) = "" Then
        cant_age = 0
    Else
        cant_age = CInt(feArqueoPagare.TextMatrix(pnRow, 3))
    End If
        
    If feArqueoPagare.TextMatrix(pnRow, 4) = "" Then
        cant_bn = 0
        Else
        cant_bn = CInt(feArqueoPagare.TextMatrix(pnRow, 4))
    End If

    cant_fisico = CInt(cant_age + cant_bn)
    
    getCantFisico = cant_fisico
End Function

Private Sub feArqueoPagare_OnCellChange(pnRow As Long, pnCol As Long)
    Dim i As Long
    
    If pnCol = 3 Or pnCol = 4 Then
        feArqueoPagare.TextMatrix(pnRow, 5) = CInt(getCantFisico(pnRow))
        
        If Trim(feArqueoPagare.TextMatrix(pnRow, 2)) = Trim(feArqueoPagare.TextMatrix(pnRow, 5)) Then
            feArqueoPagare.TextMatrix(pnRow, 6) = "OK"
        Else
            feArqueoPagare.TextMatrix(pnRow, 6) = ""
        End If
        ''SendKeys "{Tab}", True 'marg esto desactiva bloqnum
    End If
End Sub
Private Sub feArqueoPagare_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    Dim row As Long
    Dim i As Long
    Dim bHab As Boolean
    Dim x As Long
    Dim CodTipoCred As String
    Dim DesTipoCred As String
    row = feArqueoPagare.row
    If Trim(feArqueoPagare.TextMatrix(row, 2)) = Trim(feArqueoPagare.TextMatrix(row, 5)) Then
        MsgBox "No existen pagarés faltantes para este tipo de crédito. Continue.", vbInformation, "Mensaje"
        psCodigo = "OK"
        Exit Sub
    End If
    If Trim(feArqueoPagare.TextMatrix(row, 5)) = "" Then
        MsgBox "No se ha ingresado la Cantidad Fisica. Verifique.", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    bHab = False
    
    CodTipoCred = MatPagares(row).CodTipoCredito
    DesTipoCred = feArqueoPagare.TextMatrix(row, 1)


    frmArqueoPagareFaltante.Inicio MatPagareFaltante, CodTipoCred, DesTipoCred, nEstado
    psCodigo = "OK"
'    nMatDetFaltante = frmArqueoPagareDetFal.Inicio(esCompleto, CLng(Trim(feArqueoPagare.TextMatrix(row, 2))) - CLng(getCantFisico(row)), nEstado, CodTipoCred, cFechaIni, cFechaFin, DesTipoCred, nMatDetFaltante)
'
'    For x = 1 To UBound(nMatDetFaltante, 2)
'        If nMatDetFaltante(3, x) = 1 Then
'            bHab = True
'        End If
'    Next
'    psCodigo = IIf(bHab, "OK", "")

End Sub
Private Sub feArqueoPagare_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
      If pnCol = 3 Or pnCol = 4 Then
        If IsNumeric(feArqueoPagare.TextMatrix(feArqueoPagare.row, pnCol)) = False Then
            Cancel = False
            SendKeys "{Tab}", True
            Exit Sub
        End If
        If feArqueoPagare.TextMatrix(feArqueoPagare.row, pnCol) < 0 Then
            Cancel = False
            SendKeys "{Tab}", True
            Exit Sub
        End If
        If CLng(feArqueoPagare.TextMatrix(feArqueoPagare.row, 2)) < CLng(getCantFisico(pnRow)) Then
            Cancel = False
            SendKeys "{Tab}", True
            Exit Sub
        End If
    End If
End Sub
Private Sub Form_Load()
    CargaDatos
End Sub
Private Sub ObtenerDatosArqueo()
    Dim i As Long
    Set rsArqueoPagare = New ADODB.Recordset
    Set rsArqueoPagareCredito = New ADODB.Recordset
    Set rsArqueoPagareCreditoDet = New ADODB.Recordset
    
    Set rsArqueoPagare = oCaja.ObtieneArqueoPagare(CDate(FechaIni), CDate(FechaFin), nEstado, 1)
    Set rsArqueoPagareCredito = oCaja.ObtieneArqueoPagareCredito(rsArqueoPagare!nIdArqueoPag)
    
    ReDim Preserve MatArqueoPagareCredito(0)
    Do While Not rsArqueoPagareCredito.EOF
        ReDim Preserve MatArqueoPagareCredito(UBound(MatArqueoPagareCredito) + 1)
        MatArqueoPagareCredito(UBound(MatArqueoPagareCredito)).nIdArqueoPagCred = rsArqueoPagareCredito!nIdArqueoPagCred
        MatArqueoPagareCredito(UBound(MatArqueoPagareCredito)).nIdArqueoPag = rsArqueoPagareCredito!nIdArqueoPag
        MatArqueoPagareCredito(UBound(MatArqueoPagareCredito)).nCantSistema = rsArqueoPagareCredito!nCantSistema
        MatArqueoPagareCredito(UBound(MatArqueoPagareCredito)).nCantFisicoAgen = rsArqueoPagareCredito!nCantFisicoAgen
        MatArqueoPagareCredito(UBound(MatArqueoPagareCredito)).nCantFisicoBN = rsArqueoPagareCredito!nCantFisicoBN
        MatArqueoPagareCredito(UBound(MatArqueoPagareCredito)).nCantFisico = rsArqueoPagareCredito!nCantFisico
        MatArqueoPagareCredito(UBound(MatArqueoPagareCredito)).cConsDescripcion = rsArqueoPagareCredito!cConsDescripcion
        rsArqueoPagareCredito.MoveNext
    Loop
    ReDim Preserve MatArqueoPagareCreditoDet(0)
    For i = 1 To UBound(MatArqueoPagareCredito)
        Set rsArqueoPagareCreditoDet = oCaja.ObtieneArqueoPagareCreditoDet(MatArqueoPagareCredito(i).nIdArqueoPagCred)
        Do While Not rsArqueoPagareCreditoDet.EOF
            ReDim Preserve MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet) + 1)
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).nIdArqueoPagCredDet = rsArqueoPagareCreditoDet!nIdArqueoPagCredDet
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).nIdArqueoPagCred = rsArqueoPagareCreditoDet!nIdArqueoPagCred
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cNumeroPagare = rsArqueoPagareCreditoDet!cNumeroPagare
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).nFaltante = rsArqueoPagareCreditoDet!nFaltante
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cCreditoAntiguo = rsArqueoPagareCreditoDet!cCreditoAntiguo
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cCliente = rsArqueoPagareCreditoDet!cCliente
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cMoneda = rsArqueoPagareCreditoDet!cMoneda
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cMAprobado = rsArqueoPagareCreditoDet!cMAprobado
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cSCapital = rsArqueoPagareCreditoDet!cSCapital
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cDesembolo = rsArqueoPagareCreditoDet!cDesembolo
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cTasa = rsArqueoPagareCreditoDet!cTasa
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cAnalista = rsArqueoPagareCreditoDet!cAnalista
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cAtraso = rsArqueoPagareCreditoDet!cAtraso
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cTipoCredito = rsArqueoPagareCreditoDet!cTipoCredito
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cCondicion = rsArqueoPagareCreditoDet!cCondicion
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cEstado = rsArqueoPagareCreditoDet!cEstado
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cDesembolsoBN = rsArqueoPagareCreditoDet!cDesembolsoBN
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cOficinaBN = rsArqueoPagareCreditoDet!cOficinaBN
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cDistrito = rsArqueoPagareCreditoDet!cDistrito
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cZona = rsArqueoPagareCreditoDet!cZona
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cDireccion = rsArqueoPagareCreditoDet!cDireccion
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cCliente = rsArqueoPagareCreditoDet!cCliente
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cCodTipoCred = rsArqueoPagareCreditoDet!cCodTipoCred
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cNroc = rsArqueoPagareCreditoDet!cNroc
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cPlazo = rsArqueoPagareCreditoDet!cPlazo
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cVigencia = rsArqueoPagareCreditoDet!cVigencia
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cCalifInterna = rsArqueoPagareCreditoDet!cCalifInterna
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cNroEntidades = rsArqueoPagareCreditoDet!cNroEntidades
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cDirectoSoles = rsArqueoPagareCreditoDet!cDirectoSoles
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cDirectoDolares = rsArqueoPagareCreditoDet!cDirectoDolares
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cIndirectoSoles = rsArqueoPagareCreditoDet!cIndirectoSoles
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cIndirectoDolares = rsArqueoPagareCreditoDet!cIndirectoDolares
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cCalifSistFinan = rsArqueoPagareCreditoDet!cCalifSistFinan
            MatArqueoPagareCreditoDet(UBound(MatArqueoPagareCreditoDet)).cMDesembolso = rsArqueoPagareCreditoDet!cMDesembolso
            rsArqueoPagareCreditoDet.MoveNext
        Loop
    Next
End Sub
Private Sub ConsultarDatos(Optional Estado As Integer = 1)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    
    Me.lblAgeDesc.Caption = UCase(gsNomAge)
    Me.lblAgeCod.Caption = gsCodAge
    Me.lblUsuSupValue.Caption = UCase(gsCodUser)
    Me.lblFechaValue.Caption = CDate(gdFecSis)

    cFechaIni = Format(dtpPerIni.value, "YYYYmmdd")
    cFechaFin = Format(dtpPerFin.value, "YYYYmmdd")
    FechaIni = dtpPerIni.value
    FechaFin = dtpPerFin.value
    ReDim nMatDetFaltante(3, 0)
    Set rs = oCaja.ObtieneArqueoPagare(CDate(FechaIni), CDate(FechaFin), Estado, 1)
    
    If Not (rs.EOF And rs.BOF) Then
        MsgBox "El Arqueo ya fue realizado el " & Format(rs!dfecha, "dd/mm/YYYY") & " por el usuario " & rs!cUserSuperviza, vbInformation, "Mensaje"
        ObtenerDatosArqueo
        GenerarReporteArqueoPagares
    Else
         MsgBox "El Arqueo no fue realizado,por favor carga el reporte contabilizado de Creditos Vigentes o Cancelados para procesar los datos del Arqueo", vbInformation, "Mensaje"
'        If (Estado = 0) Then
'            Set rs = oCaja.ObtieneCreditosCanceladosArqueoPagare(cFechaIni, cFechaFin, "1,2", gsCodAge, "", "2002", "2050", "3013", "20", "28", "100105", "100101", "1001")
'        End If
'        If (Estado = 1) Then
'            Set rs = oCaja.ObtieneCreditosVigentesArqueoPagare(CDbl("0.01"), CDbl("999999999"), cFechaIni, cFechaFin, "1,2", "1,2,3,4,5,6,7", "", gsCodAge, 1, "")
'        End If
'
'        ReDim Preserve MatPagares(0)
'        Do While Not rs.EOF
'
'            feArqueoPagare.AdicionaFila
'            feArqueoPagare.TextMatrix(feArqueoPagare.row, 1) = rs!cConsDescripcion
'            feArqueoPagare.TextMatrix(feArqueoPagare.row, 2) = rs!CantSistema
'            feArqueoPagare.TextMatrix(feArqueoPagare.row, 7) = rs!cTpoCredCod
'
'            ReDim Preserve MatPagares(UBound(MatPagares) + 1)
'            MatPagares(UBound(MatPagares)).TipoCredito = rs!cConsDescripcion
'            MatPagares(UBound(MatPagares)).CantSistema = rs!CantSistema
'            MatPagares(UBound(MatPagares)).CodTipoCredito = rs!cTpoCredCod
'            rs.MoveNext
'        Loop
    End If

End Sub
Private Sub CargaDatos()

    Me.lblAgeDesc.Caption = UCase(gsNomAge)
    Me.lblAgeCod.Caption = gsCodAge
    Me.lblUsuSupValue.Caption = UCase(gsCodUser)
    Me.lblFechaValue.Caption = CDate(gdFecSis)
    dtpPerFin.value = CDate(gdFecSis)

End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdRegArqueo.SetFocus
    End If
End Sub
