VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogImpComprobantes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logística: Impresión de Comprobantes Pendientes"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16275
   Icon            =   "frmLogImpComprobantes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   16275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Comprobantes"
      TabPicture(0)   =   "frmLogImpComprobantes.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feComprobantes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSalir"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdImprimir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdExportar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdActualizar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "&Actualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   13320
         TabIndex        =   5
         Top             =   6120
         Width           =   1335
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "&Exportar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   4
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   3
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "S&alir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   14760
         TabIndex        =   2
         Top             =   6120
         Width           =   1215
      End
      Begin Sicmact.FlexEdit feComprobantes 
         Height          =   5235
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   15720
         _extentx        =   27728
         _extenty        =   9234
         cols0           =   12
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Tipo Documento-Nº Doc.-F. Emisión-Proveedor-Moneda-Importe-Origen-Nº Doc. Origen-Glosa-Cuenta-Entidad Cuenta"
         encabezadosanchos=   "500-1600-1600-1200-3200-1000-1200-1600-1500-2000-1800-3000"
         font            =   "frmLogImpComprobantes.frx":0326
         font            =   "frmLogImpComprobantes.frx":034E
         font            =   "frmLogImpComprobantes.frx":0376
         font            =   "frmLogImpComprobantes.frx":039E
         font            =   "frmLogImpComprobantes.frx":03C6
         fontfixed       =   "frmLogImpComprobantes.frx":03EE
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         tipobusqueda    =   7
         columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
         textstylefixed  =   3
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-C-C-L-C-C-L-C-L-L-L"
         formatosedit    =   "0-0-0-5-0-0-0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbbuscaduplicadotext=   -1
         appearance      =   0
         colwidth0       =   495
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmLogImpComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub cmdlSalir_Click()
Unload Me
End Sub

Public Sub Inicio()
Me.Show 1
End Sub

Private Sub CargarGrid()
Dim oLog As DLogGeneral
Dim rsLog As ADODB.Recordset

Set oLog = New DLogGeneral

'Set rsLog = oLog.ObtenerComprobantePend()
Set rsLog = oLog.ListaComprobantesxImpresion() 'EJVG20131112
Call LimpiaFlex(feComprobantes)

If rsLog.RecordCount > 0 Then
    For i = 0 To rsLog.RecordCount - 1
        feComprobantes.AdicionaFila
        feComprobantes.TextMatrix(i + 1, 0) = i + 1
        feComprobantes.TextMatrix(i + 1, 1) = rsLog!TpoDoc
        feComprobantes.TextMatrix(i + 1, 2) = rsLog!NDoc
        feComprobantes.TextMatrix(i + 1, 3) = Format(rsLog!FEmision, "dd/mm/yyyy")
        feComprobantes.TextMatrix(i + 1, 4) = rsLog!Proveedor
        feComprobantes.TextMatrix(i + 1, 5) = rsLog!Moneda
        feComprobantes.TextMatrix(i + 1, 6) = rsLog!Importe
        feComprobantes.TextMatrix(i + 1, 7) = rsLog!Origen
        feComprobantes.TextMatrix(i + 1, 8) = rsLog!DocOrigen
        feComprobantes.TextMatrix(i + 1, 9) = rsLog!Glosa
        feComprobantes.TextMatrix(i + 1, 10) = rsLog!Cuenta
        feComprobantes.TextMatrix(i + 1, 11) = rsLog!EntidadCuenta
        rsLog.MoveNext
    Next i
Else
    MsgBox "No hay datos.", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdActualizar_Click()
Call CargarGrid
End Sub

Private Sub cmdExportar_Click()
Dim lnI As Long
Screen.MousePointer = 11

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
        
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Range("B2", "D2").MergeCells = True
    ApExcel.Cells(2, 13).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "IMPRESIÓN DE COMPROBANTES PENDIENTES"
    
    ApExcel.Cells(5, 2).Formula = UCase(Trim(gsNomAge))
    
    ApExcel.Range("B4", "M4").MergeCells = True
    ApExcel.Range("B5", "M5").MergeCells = True
    ApExcel.Range("B4", "M5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "ITEM"
    ApExcel.Cells(8, 3).Formula = "TIPO DOC."
    ApExcel.Cells(8, 4).Formula = "Nº DOCUMENTO"
    ApExcel.Cells(8, 5).Formula = "FECHA EMISION"
    ApExcel.Cells(8, 6).Formula = "PROVEEDOR"
    ApExcel.Cells(8, 7).Formula = "MONEDA"
    ApExcel.Cells(8, 8).Formula = "IMPORTE"
    ApExcel.Cells(8, 9).Formula = "ORIGEN"
    ApExcel.Cells(8, 10).Formula = "DOC. ORIGEN"
    ApExcel.Cells(8, 11).Formula = "GLOSA"
    ApExcel.Cells(8, 12).Formula = "CUENTA"
    ApExcel.Cells(8, 13).Formula = "ENTIDAD CUENTA"
    
    ApExcel.Range("B2", "M8").Font.Bold = True
    
    ApExcel.Range("B8", "M8").Interior.Color = RGB(219, 219, 219)
    ApExcel.Range("B8", "M8").HorizontalAlignment = 3
    ApExcel.Range("B8", "M8").VerticalAlignment = 3
    
    
    i = 8
    For lnI = 1 To feComprobantes.Rows - 1
        i = i + 1
        ApExcel.Cells(i, 2).Formula = lnI
        ApExcel.Cells(i, 3).Formula = feComprobantes.TextMatrix(lnI, 1)
        ApExcel.Cells(i, 4).Formula = feComprobantes.TextMatrix(lnI, 2)
        ApExcel.Cells(i, 5).Formula = feComprobantes.TextMatrix(lnI, 3)
        ApExcel.Cells(i, 6).Formula = feComprobantes.TextMatrix(lnI, 4)
        ApExcel.Cells(i, 7).Formula = feComprobantes.TextMatrix(lnI, 5)
        ApExcel.Cells(i, 8).Formula = feComprobantes.TextMatrix(lnI, 6)
        ApExcel.Cells(i, 9).Formula = feComprobantes.TextMatrix(lnI, 7)
        ApExcel.Cells(i, 10).Formula = feComprobantes.TextMatrix(lnI, 8)
        ApExcel.Cells(i, 11).Formula = feComprobantes.TextMatrix(lnI, 9)
        ApExcel.Cells(i, 12).Formula = "'" & feComprobantes.TextMatrix(lnI, 10)
        ApExcel.Cells(i, 13).Formula = feComprobantes.TextMatrix(lnI, 11)
    Next lnI
    
    ApExcel.Range(ApExcel.Cells(8, 2), ApExcel.Cells(i, 13)).Borders.LineStyle = 1
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub

Private Sub cmdImprimir_Click()
    Dim oImpre As New COMFunciones.FCOMImpresion

If Me.feComprobantes.TextMatrix(1, 1) = "" Then
    MsgBox "No existen Comprobantes Pendientes.", vbInformation, "Aviso"
    Me.cmdActualizar.SetFocus
    Exit Sub
End If

Dim lsCadena As String
Dim lnPagina As Long
Dim lnItem As Long
Dim lnI As Long
Dim oPrevio As clsPrevio
    
    
Set oPrevio = New clsPrevio

Dim lsItem As String * 6
Dim lsTpoDoc As String * 10
Dim lsNDoc As String * 13
Dim lsFemision As String * 11
Dim lsProveedor As String * 25
Dim lsMoneda As String * 8
Dim lsImporte As String * 12
Dim lsOrigen As String * 5
Dim lsDocOrigen As String * 13
Dim lsGlosa As String * 20
Dim lsCuenta As String * 18
Dim lsCuentaEnt As String * 20
    

Dim oCon As DConecta
Set oCon = New DConecta
    
lsCadena = ""

lsCadena = lsCadena & oImpresora.gPrnCondensadaON
lsCadena = lsCadena & CabeceraPagina1("IMPRESIÓN DE COMPROBANTES PENDIENTES", lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "")
lsCadena = lsCadena & Encabezado("Item;6;Tipo Doc.;10;Nº Doc;13;FEmision;11;Proveedor;20;;7;Moneda;8;Importe;10;Origen;7;Doc.Origen;13;Glosa;8;Cuenta;18;Entidad Cuenta;20;", lnItem)
    
    For lnI = 1 To feComprobantes.Rows - 1
        RSet lsItem = lnI
        RSet lsTpoDoc = feComprobantes.TextMatrix(lnI, 1)
        RSet lsNDoc = feComprobantes.TextMatrix(lnI, 2)
        RSet lsFemision = feComprobantes.TextMatrix(lnI, 3)
        RSet lsProveedor = feComprobantes.TextMatrix(lnI, 4)
        RSet lsMoneda = feComprobantes.TextMatrix(lnI, 5)
        RSet lsImporte = feComprobantes.TextMatrix(lnI, 6)
        RSet lsOrigen = feComprobantes.TextMatrix(lnI, 7)
        RSet lsDocOrigen = feComprobantes.TextMatrix(lnI, 8)
        RSet lsGlosa = feComprobantes.TextMatrix(lnI, 9)
        RSet lsCuenta = feComprobantes.TextMatrix(lnI, 10)
        RSet lsCuentaEnt = feComprobantes.TextMatrix(lnI, 11)
        
        lsCadena = lsCadena & lsItem
        lsCadena = lsCadena & oImpre.CentrarCadena(Trim(lsTpoDoc), 10)
        lsCadena = lsCadena & oImpre.CentrarCadena(Trim(lsNDoc), 13)
        lsCadena = lsCadena & lsFemision
        lsCadena = lsCadena & Mid(Space(1) & Trim(lsProveedor) & Space(50), 1, 24) & Space(1)
        lsCadena = lsCadena & lsMoneda
        lsCadena = lsCadena & lsImporte
        lsCadena = lsCadena & Mid(Space(2) & Trim(lsOrigen) & Space(10), 1, 6) & Space(1)
        lsCadena = lsCadena & oImpre.CentrarCadena(Trim(lsDocOrigen), 13)
        lsCadena = lsCadena & Mid(Space(1) & Trim(lsGlosa) & Space(50), 1, 15)
        lsCadena = lsCadena & Mid(Space(1) & Trim(lsCuenta) & Space(20), 1, 18)
        lsCadena = lsCadena & Mid(Space(1) & Trim(lsCuentaEnt), 1, 20) & oImpresora.gPrnSaltoLinea
        
        If lnItem > 52 Then
            lnItem = 0
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            lsCadena = lsCadena & CabeceraPagina1("IMPRESIÓN DE COMPROBANTES PENDIENTES", lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "")
            lsCadena = lsCadena & Encabezado("Item;6;Tipo Doc.;10;Nº Doc;13;FEmision;11;Proveedor;20;;7;Moneda;8;Importe;10;Origen;7;Doc.Origen;13;Glosa;8;Cuenta;18;Entidad Cuenta;20;", lnItem)
        End If
        
        lnItem = lnItem + 1
    Next lnI
     
    
    oPrevio.Show lsCadena, "Impresión de Comprobantes Pendientes", True, 66
    Set oPrevio = Nothing
    
    'ARLO 20160126 ***
    gsOpeCod = LogPistaComprobantes
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Imprimio Comprobantes "
    Set objPista = Nothing
    '***

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call CargarGrid
End Sub
