VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHojaRutaAnalistaGenera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hoja de ruta"
   ClientHeight    =   6585
   ClientLeft      =   6540
   ClientTop       =   3675
   ClientWidth     =   16125
   ControlBox      =   0   'False
   Icon            =   "frmHojaRutaAnalistaGenera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   16125
   Begin VB.CommandButton cmbCerrar 
      Caption         =   "Cerrar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   14400
      TabIndex        =   4
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmbGenerar 
      Caption         =   "Generar Excel"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6000
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Clientes Promocionales"
      TabPicture(0)   =   "frmHojaRutaAnalistaGenera.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flxClientesPromocionales"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Clientes en Mora"
      TabPicture(1)   =   "frmHojaRutaAnalistaGenera.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxClientesMora"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Web, Kiosko y Facebook"
      TabPicture(2)   =   "frmHojaRutaAnalistaGenera.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "flxClientesWeb"
      Tab(2).ControlCount=   1
      Begin SICMACT.FlexEdit flxClientesMora 
         Height          =   4695
         Left            =   -74760
         TabIndex        =   3
         Top             =   600
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   8281
         Cols0           =   14
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmHojaRutaAnalistaGenera.frx":035E
         EncabezadosAnchos=   "600-1500-2500-2000-2000-1200-1200-1000-1200-1200-1200-1200-1200-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-L-L-R-C-R-C-R-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-3-5-2-2-3-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit flxClientesPromocionales 
         Height          =   4695
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   8281
         Cols0           =   14
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cod. Cliente-Cliente-DOI-Dirección-Zona-Actividad-Teléfono-Móvil-ult.Cred-Endeudamiento-N° IFIs-Observaciones-cCtaCod"
         EncabezadosAnchos=   "600-1500-2500-1200-2500-2000-2000-1200-1200-2000-1200-1000-1200-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-L-L-L-C-C-R-C-R-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-2-2-3-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit flxClientesWeb 
         Height          =   4695
         Left            =   -74760
         TabIndex        =   5
         Top             =   600
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   8281
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cliente-DOI-Dirección-Profesión-Teléfono-Móvil-Monto Solicitado-N°Cuotas-Observaciones"
         EncabezadosAnchos=   "600-2500-1200-2500-2000-1200-1200-2000-1000-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-C-C-R-R-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-2-3-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmHojaRutaAnalistaGenera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dHojaRuta As New DCOMhojaRuta
Dim rsClientesPromocionales As ADODB.Recordset
Dim rsClientesMora As ADODB.Recordset
Dim bGeneraDiaSiguiente As Integer
Public Sub inicio(ByVal pbGeneraDiaSiguiente As Integer)
    bGeneraDiaSiguiente = pbGeneraDiaSiguiente
    Me.Show 1
End Sub

Private Sub cmbCerrar_Click()
    Unload Me
End Sub

Private Sub cmbGenerar_Click()
    ImprimeHojaRuta
    cmbCerrar.Enabled = True
End Sub


Private Sub Form_Load()
    'cargar la lista de clientes
    Dim cFecha As String: cFecha = Format(gdFecSis, "YYYYMM")
    Set rsClientesPromocionales = dHojaRuta.GeneraCarteraDiariaPromocion(cFecha, gsCodUser, bGeneraDiaSiguiente)
    llenarClientesPromocionales
    Set rsClientesMora = dHojaRuta.GeneraCarteraDiariaMora(cFecha, gsCodUser, bGeneraDiaSiguiente)
    llenarClientesMora
End Sub
Private Function llenarClientesPromocionales()
    Dim nRow As Integer
    Dim nRow2 As Integer
    
    LimpiaFlex flxClientesPromocionales
    flxClientesPromocionales.Rows = 2
    
    LimpiaFlex flxClientesWeb
    flxClientesWeb.Rows = 2
    
    Do While Not rsClientesPromocionales.EOF
    
        If rsClientesPromocionales!bWeb = 0 Then
            flxClientesPromocionales.AdicionaFila
            nRow = flxClientesPromocionales.Rows - 1
            flxClientesPromocionales.TextMatrix(nRow, 1) = rsClientesPromocionales!cPersCodCliente
            flxClientesPromocionales.TextMatrix(nRow, 2) = rsClientesPromocionales!cPersNombre
            flxClientesPromocionales.TextMatrix(nRow, 3) = rsClientesPromocionales!cDOI
            flxClientesPromocionales.TextMatrix(nRow, 4) = rsClientesPromocionales!cDireccion
            flxClientesPromocionales.TextMatrix(nRow, 5) = rsClientesPromocionales!cZona
            flxClientesPromocionales.TextMatrix(nRow, 6) = rsClientesPromocionales!cActividad
            flxClientesPromocionales.TextMatrix(nRow, 7) = rsClientesPromocionales!cTelefono
            flxClientesPromocionales.TextMatrix(nRow, 8) = rsClientesPromocionales!cMovil
            flxClientesPromocionales.TextMatrix(nRow, 9) = rsClientesPromocionales!nMontoUltCred
            flxClientesPromocionales.TextMatrix(nRow, 10) = rsClientesPromocionales!nEndeudamiento
            flxClientesPromocionales.TextMatrix(nRow, 11) = rsClientesPromocionales!nNIFIS
        Else
            flxClientesWeb.AdicionaFila
            nRow2 = flxClientesWeb.Rows - 1
            flxClientesWeb.TextMatrix(nRow2, 1) = rsClientesPromocionales!cPersNombre
            flxClientesWeb.TextMatrix(nRow2, 2) = rsClientesPromocionales!cDOI
            flxClientesWeb.TextMatrix(nRow2, 3) = rsClientesPromocionales!cDireccion
            flxClientesWeb.TextMatrix(nRow2, 4) = rsClientesPromocionales!cActividad
            flxClientesWeb.TextMatrix(nRow2, 5) = rsClientesPromocionales!cTelefono
            flxClientesWeb.TextMatrix(nRow2, 6) = rsClientesPromocionales!cMovil
            flxClientesWeb.TextMatrix(nRow2, 7) = rsClientesPromocionales!nMontoUltCred
            flxClientesWeb.TextMatrix(nRow2, 8) = rsClientesPromocionales!nNIFIS
        End If
        

        rsClientesPromocionales.MoveNext
        
    Loop
End Function
Private Function llenarClientesMora()
    Dim nRow As Integer
    LimpiaFlex flxClientesMora
    flxClientesMora.Rows = 2
    Do While Not rsClientesMora.EOF
        flxClientesMora.AdicionaFila
        nRow = flxClientesMora.Rows - 1
        flxClientesMora.TextMatrix(nRow, 1) = rsClientesMora!cPersCodCliente
        flxClientesMora.TextMatrix(nRow, 2) = rsClientesMora!cPersNombre
        flxClientesMora.TextMatrix(nRow, 3) = rsClientesMora!cDireccion
        flxClientesMora.TextMatrix(nRow, 4) = rsClientesMora!cDireccionNegocio
        flxClientesMora.TextMatrix(nRow, 5) = rsClientesMora!cTelefono
        flxClientesMora.TextMatrix(nRow, 6) = rsClientesMora!cMovil
        flxClientesMora.TextMatrix(nRow, 7) = rsClientesMora!nCuota
        flxClientesMora.TextMatrix(nRow, 8) = rsClientesMora!dFechaVenc
        flxClientesMora.TextMatrix(nRow, 9) = rsClientesMora!nMontoCuota
        flxClientesMora.TextMatrix(nRow, 10) = rsClientesMora!nMora
        flxClientesMora.TextMatrix(nRow, 11) = rsClientesMora!nDiasAtraso
        'flxClientesMora.TextMatrix(nRow, 13) = rsClientesMora!cCtaCod
        rsClientesMora.MoveNext
    Loop
End Function

Public Sub ImprimeHojaRuta()
    Dim oExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet
    Dim iFila As Long, i As Integer
    
    Set oExcel = New Excel.Application
    Set oWBook = oExcel.Workbooks.Add
    Set oSheet = oWBook.Worksheets(1)
    Screen.MousePointer = vbHourglass
    
    
    oSheet.Cells(1, 1) = "CLIENTES EN MORA"
    With oSheet.Range("A1:M1")
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    
    
    iFila = 2
    For i = 0 To flxClientesMora.Cols - 2
        ' pone el nombre de los campos en la primera fila
        oSheet.Cells(iFila, i + 1) = flxClientesMora.TextMatrix(0, i)
    Next
    
    'llenamos los datos de las visitas mora
    For i = 1 To flxClientesMora.Rows - 1
        oSheet.Cells(i + iFila, 1) = flxClientesMora.TextMatrix(i, 0)
        oSheet.Cells(i + iFila, 2) = "'" & flxClientesMora.TextMatrix(i, 1)
        oSheet.Cells(i + iFila, 3) = "'" & flxClientesMora.TextMatrix(i, 2)
        oSheet.Cells(i + iFila, 4) = "'" & flxClientesMora.TextMatrix(i, 3)
        oSheet.Cells(i + iFila, 5) = "'" & flxClientesMora.TextMatrix(i, 4)
        oSheet.Cells(i + iFila, 6) = "'" & flxClientesMora.TextMatrix(i, 5)
        oSheet.Cells(i + iFila, 7) = "'" & flxClientesMora.TextMatrix(i, 6)
        oSheet.Cells(i + iFila, 8) = flxClientesMora.TextMatrix(i, 7)
        oSheet.Cells(i + iFila, 9) = "'" & flxClientesMora.TextMatrix(i, 8)
        oSheet.Cells(i + iFila, 10) = flxClientesMora.TextMatrix(i, 9)
        oSheet.Cells(i + iFila, 11) = flxClientesMora.TextMatrix(i, 10)
        oSheet.Cells(i + iFila, 12) = flxClientesMora.TextMatrix(i, 11)
        oSheet.Cells(i + iFila, 13) = flxClientesMora.TextMatrix(i, 12)
    Next
    
    oSheet.Range(oSheet.Cells(iFila, 1), oSheet.Cells(i + iFila - 1, 13)).Borders.Weight = xlThin
    
    iFila = iFila + i + 3 'margen de separacion para la otra tabla
    
    oSheet.Cells(iFila, 1) = "CLIENTES PROMOCIONALES"
    With oSheet.Range(oSheet.Cells(iFila, 1), oSheet.Cells(iFila, 13))
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    
    iFila = iFila + 1
    For i = 0 To flxClientesPromocionales.Cols - 2
        ' pone el nombre de los campos en la primera fila
        oSheet.Cells(iFila, i + 1) = flxClientesPromocionales.TextMatrix(0, i)
    Next
    
    
    'llenamos los datos de las visitas promocion
    For i = 1 To flxClientesPromocionales.Rows - 1
        oSheet.Cells(i + iFila, 1) = flxClientesPromocionales.TextMatrix(i, 0)
        oSheet.Cells(i + iFila, 2) = "'" & flxClientesPromocionales.TextMatrix(i, 1)
        oSheet.Cells(i + iFila, 3) = "'" & flxClientesPromocionales.TextMatrix(i, 2)
        oSheet.Cells(i + iFila, 4) = "'" & flxClientesPromocionales.TextMatrix(i, 3)
        oSheet.Cells(i + iFila, 5) = "'" & flxClientesPromocionales.TextMatrix(i, 4)
        oSheet.Cells(i + iFila, 6) = "'" & flxClientesPromocionales.TextMatrix(i, 5)
        oSheet.Cells(i + iFila, 7) = "'" & flxClientesPromocionales.TextMatrix(i, 6)
        oSheet.Cells(i + iFila, 8) = flxClientesPromocionales.TextMatrix(i, 7)
        oSheet.Cells(i + iFila, 9) = "'" & flxClientesPromocionales.TextMatrix(i, 8)
        oSheet.Cells(i + iFila, 10) = flxClientesPromocionales.TextMatrix(i, 9)
        oSheet.Cells(i + iFila, 11) = flxClientesPromocionales.TextMatrix(i, 10)
        oSheet.Cells(i + iFila, 12) = flxClientesPromocionales.TextMatrix(i, 11)
        oSheet.Cells(i + iFila, 13) = flxClientesPromocionales.TextMatrix(i, 12)
    Next
    
    oSheet.Range(oSheet.Cells(iFila, 1), oSheet.Cells(i + iFila - 1, 13)).Borders.Weight = xlThin
    
    '------------------------------------------------------------------------------------------------
    iFila = iFila + i + 3 'margen de separacion para la otra tabla
    
    oSheet.Cells(iFila, 1) = "SOLICITUDES WEB, KIOSKO Y FACEBOOK"
    With oSheet.Range(oSheet.Cells(iFila, 1), oSheet.Cells(iFila, 13))
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    
    iFila = iFila + 1
    For i = 0 To flxClientesWeb.Cols - 2
        ' pone el nombre de los campos en la primera fila
        oSheet.Cells(iFila, i + 1) = flxClientesWeb.TextMatrix(0, i)
    Next
    
    'llenamos los datos de las visitas promocion
    For i = 1 To flxClientesWeb.Rows - 1
        oSheet.Cells(i + iFila, 1) = flxClientesWeb.TextMatrix(i, 0)
        oSheet.Cells(i + iFila, 2) = "'" & flxClientesWeb.TextMatrix(i, 1)
        oSheet.Cells(i + iFila, 3) = "'" & flxClientesWeb.TextMatrix(i, 2)
        oSheet.Cells(i + iFila, 4) = "'" & flxClientesWeb.TextMatrix(i, 3)
        oSheet.Cells(i + iFila, 5) = "'" & flxClientesWeb.TextMatrix(i, 4)
        oSheet.Cells(i + iFila, 6) = "'" & flxClientesWeb.TextMatrix(i, 5)
        oSheet.Cells(i + iFila, 7) = "'" & flxClientesWeb.TextMatrix(i, 6)
        oSheet.Cells(i + iFila, 8) = flxClientesWeb.TextMatrix(i, 7)
        oSheet.Cells(i + iFila, 9) = "'" & flxClientesWeb.TextMatrix(i, 8)
    Next
    
    oSheet.Range(oSheet.Cells(iFila, 1), oSheet.Cells(i + iFila - 1, 13)).Borders.Weight = xlThin
    
    '------------------------------------------------------------------------------------------------
    oSheet.Range("A2").ColumnWidth = 2
    oSheet.Range("B2").ColumnWidth = 13
    oSheet.Range("C2").ColumnWidth = 27
    oSheet.Range("D2").ColumnWidth = 20
    oSheet.Range("E2").ColumnWidth = 30
    oSheet.Range("F2").ColumnWidth = 25
    oSheet.Range("G2").ColumnWidth = 25
    oSheet.Range("H2").ColumnWidth = 10
    oSheet.Range("I2").ColumnWidth = 10
    oSheet.Range("J2").ColumnWidth = 12
    oSheet.Range("K2").ColumnWidth = 13
    oSheet.Range("L2").ColumnWidth = 10
    oSheet.Range("M2").ColumnWidth = 20
    
    With oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(i + iFila - 1, 13))
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 9
    End With
    

    'oSheet.Columns("A:M").EntireColumn.AutoFit
    oExcel.Visible = True
    Set oExcel = Nothing
    Screen.MousePointer = vbDefault
End Sub

