VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmColRecPagGestor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos de Gestores"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   345
      Left            =   9720
      TabIndex        =   4
      Top             =   3930
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Busqueda"
      Height          =   765
      Left            =   30
      TabIndex        =   1
      Top             =   3660
      Width           =   10875
      Begin VB.CommandButton CmdExportar 
         Caption         =   "EX¨P"
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
         Left            =   2520
         TabIndex        =   10
         Top             =   330
         Width           =   585
      End
      Begin MSComCtl2.DTPicker DPFechaFinal 
         Height          =   285
         Left            =   6690
         TabIndex        =   9
         Top             =   330
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         Format          =   95027201
         CurrentDate     =   38450
      End
      Begin MSComCtl2.DTPicker DPFechaInicial 
         Height          =   285
         Left            =   4140
         TabIndex        =   7
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   95027201
         CurrentDate     =   38450
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8070
         Top             =   180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Archivo de Excel (*.xls)|*.xls"
         FilterIndex     =   1
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   8550
         TabIndex        =   5
         Top             =   270
         Width           =   1065
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   345
         Left            =   1350
         TabIndex        =   3
         Top             =   300
         Width           =   1065
      End
      Begin VB.CommandButton CmdArchivo 
         Caption         =   "Archivo"
         Height          =   345
         Left            =   90
         TabIndex        =   2
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   5700
         TabIndex        =   8
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         Height          =   195
         Left            =   3180
         TabIndex        =   6
         Top             =   360
         Width           =   870
      End
   End
   Begin SICMACT.FlexEdit FlexEdit1 
      Height          =   3615
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   6376
      Cols0           =   11
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Movimiento-Cuenta-Saldo K-Capital-Interes-Mora-Gastos-Monto-Moneda-Titular-Fecha Pag"
      EncabezadosAnchos=   "1200-1700-1200-1200-1200-1200-1200-1200-1200-3500-1800"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-R-R-R-R-R-R-C-L-C"
      FormatosEdit    =   "0-0-2-2-2-2-2-2-0-0-5"
      TextArray0      =   "Movimiento"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   1200
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "FrmColRecPagGestor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sFilename As String
Private Sub CmdArchivo_Click()
    CommonDialog1.ShowOpen
    sFilename = CommonDialog1.FileName
End Sub

Private Sub cmdBuscar_Click()
    Dim oDColRecup As COMDColocRec.DCOMColRecRConsulta
    Dim rs As New ADODB.Recordset
    Dim fso As Scripting.FileSystemObject
    Dim nFila As Integer
    
    On Error GoTo ErrHandler
        If sFilename = "" Then
            MsgBox "Debe seleccionar un archivo", vbInformation, "AVISO"
            Exit Sub
        End If
        
        Set fso = New Scripting.FileSystemObject
        If Not fso.FileExists(sFilename) Then
            MsgBox "El archivo no existe", vbInformation, "AVISO"
            Exit Sub
        End If
        Set fso = Nothing
        
        If DPFechaInicial.value >= DPFechaFinal.value Then
            MsgBox "Las fechas estan mal dada", vbInformation, "AVISO"
        End If
        
        Set oDColRecup = New COMDColocRec.DCOMColRecRConsulta
        Set rs = oDColRecup.Recup_PagosGestores(sFilename, Format(DPFechaInicial.value, "dd/mm/yyyy"), Format(DPFechaFinal.value, "dd/mm/yyyy"))
        Set oDColRecup = Nothing
    
        If rs.EOF Or rs.BOF Then
            MsgBox "No se encontraron informacion solicitada", vbInformation, "Aviso"
            Exit Sub
        End If
    
        'Cargando la Informacion
        Do Until rs.EOF
            FlexEdit1.AdicionaFila
            nFila = Me.FlexEdit1.Rows
            
            With FlexEdit1
                .TextMatrix(nFila - 1, 0) = rs!nMovNro
                .TextMatrix(nFila - 1, 1) = rs!cCtaCod
                .TextMatrix(nFila - 1, 2) = Format(IIf(IsNull(rs!nSaldo), 0, rs!nSaldo), "#0.00")
                .TextMatrix(nFila - 1, 3) = Format(IIf(IsNull(rs!Capital), 0, rs!Capital), "#0.00")
                .TextMatrix(nFila - 1, 4) = Format(IIf(IsNull(rs!Interes), 0, rs!Interes), "#0.00")
                .TextMatrix(nFila - 1, 5) = Format(IIf(IsNull(rs!Mora), 0, rs!Mora), "#0.00")
                .TextMatrix(nFila - 1, 6) = Format(IIf(IsNull(rs!Gastos), 0, rs!Gastos), "0.00")
                .TextMatrix(nFila - 1, 7) = Format(IIf(IsNull(rs!Capital), 0, rs!Capital) + IIf(IsNull(rs!Interes), 0, rs!Interes) + IIf(IsNull(rs!Mora), 0, rs!Mora) + IIf(IsNull(rs!Gastos), 0, rs!Gastos), "#0.00")
                .TextMatrix(nFila - 1, 8) = rs!Moneda
                .TextMatrix(nFila - 1, 9) = rs!cPersNombre
                .TextMatrix(nFila - 1, 10) = rs!Fechapago
            End With
            rs.MoveNext
        Loop
        Set rs = Nothing
         Exit Sub
ErrHandler:
    MsgBox "Se ha producido un error en la consulta", vbInformation, "AVISO"
End Sub

Private Sub CmdCancelar_Click()
    Me.FlexEdit1.Clear
    Me.FlexEdit1.FormaCabecera
End Sub

Private Sub CmdExportar_Click()
    If MsgBox("Esta seguro que desea exportar esta consulta", vbQuestion + vbYesNo) = vbYes Then
        ExportarArchivo
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DPFechaInicial.value = Date
    DPFechaFinal.value = Date
End Sub

Sub ExportarArchivo()
    Dim m_Excel As Excel.Application
    Dim oLibroExcel As Excel.Workbook
    Dim oHojaExcel As Excel.Worksheet
    Dim i As Integer
    Dim nFila As Integer
    On Error GoTo ErrHandler
    
    Set m_Excel = New Excel.Application
    m_Excel.Visible = True
    If Me.FlexEdit1.Rows <= 1 Then
        MsgBox "No existe datos para la consulta", vbInformation, "AVISO"
        Exit Sub
    End If
    
    Set oLibroExcel = m_Excel.Workbooks.Add
    Set oHojaExcel = oLibroExcel.Worksheets(1)
    oHojaExcel.Visible = xlSheetVisible
    
    oHojaExcel.Activate
    oHojaExcel.PageSetup.Zoom = 75
    oHojaExcel.PageSetup.Orientation = xlLandscape
    
    m_Excel.Range("A1:R1000").Font.Size = 9
    
    m_Excel.Selection.NumberFormat = "#,##0.00"
    
    'creando el encabezado del Informe
   oHojaExcel.Range("A2:D2").Merge
   oHojaExcel.Range("A2.D2").value = "LISTA DE CREDITOS PAGADOS"
   oHojaExcel.Range("A2:D2").Font.Italic = True
   oHojaExcel.Range("A2:D2").Font.Size = 13
   
   
   nFila = 5
   
   With oHojaExcel
     .Cells(nFila, 1) = "Movimiento" 'a
     .Cells(nFila, 1).Font.Bold = True
     
     .Cells(nFila, 2) = "Cuenta" 'B
     .Cells(nFila, 2).Font.Bold = True
     
     .Cells(nFila, 3) = "Saldo K" 'C
     .Cells(nFila, 3).Font.Bold = True
     
     .Cells(nFila, 4) = "Capital" 'D
     .Cells(nFila, 4).Font.Bold = True
     
     .Cells(nFila, 5) = "Interes" 'E
     .Cells(nFila, 5).Font.Bold = True
     
     .Cells(nFila, 6) = "Mora" 'F
     .Cells(nFila, 6).Font.Bold = True
     
     .Cells(nFila, 7) = "Gastos" 'G
     .Cells(nFila, 7).Font.Bold = True
     
     .Cells(nFila, 8) = "Monto Pagado" 'H
     .Cells(nFila, 8).Font.Bold = True
     
     .Cells(nFila, 9) = "Moneda" 'I
     .Cells(nFila, 9).Font.Bold = True
     
     .Cells(nFila, 10) = "Titular" 'J
     .Cells(nFila, 10).Font.Bold = True
     
     .Cells(nFila, 11) = "Fecha Pago" 'K
     .Cells(nFila, 11).Font.Bold = True
     
     .Range("B1").EntireColumn.NumberFormat = "@"
     .Range("C1").EntireColumn.NumberFormat = "0.00"
     .Range("D1").EntireColumn.NumberFormat = "0.00"
     .Range("E1").EntireColumn.NumberFormat = "0.00"
     .Range("F1").EntireColumn.NumberFormat = "0.00"
     .Range("G1").EntireColumn.NumberFormat = "0.00"
     .Range("H1").EntireColumn.NumberFormat = "0.00"
     .Range("K1").EntireColumn.NumberFormat = "dd/mm/yyyy"
     
     
      nFila = nFila + 1
      For i = 1 To FlexEdit1.Rows - 1
          .Cells(nFila, 1) = FlexEdit1.TextMatrix(i, 0)
          .Cells(nFila, 2) = FlexEdit1.TextMatrix(i, 1)
          .Cells(nFila, 3) = FlexEdit1.TextMatrix(i, 2)
          .Cells(nFila, 4) = FlexEdit1.TextMatrix(i, 3)
          .Cells(nFila, 5) = FlexEdit1.TextMatrix(i, 4)
          .Cells(nFila, 6) = FlexEdit1.TextMatrix(i, 5)
          .Cells(nFila, 7) = FlexEdit1.TextMatrix(i, 6)
          .Cells(nFila, 8) = FlexEdit1.TextMatrix(i, 7)
          .Cells(nFila, 9) = FlexEdit1.TextMatrix(i, 8)
          .Cells(nFila, 10) = FlexEdit1.TextMatrix(i, 9)
          .Cells(nFila, 11) = FlexEdit1.TextMatrix(i, 10)
          nFila = nFila + 1
      Next i
   End With
   oHojaExcel.Range("A:A").EntireColumn.AutoFit
   oHojaExcel.Range("B:B").EntireColumn.AutoFit
   oHojaExcel.Range("J:J").EntireColumn.AutoFit
   
   oLibroExcel.PrintPreview
    Exit Sub
ErrHandler:
    MsgBox "Se ha producido un error " & Err.Description, vbInformation, "AVISO"
End Sub
