VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptFoncodes 
   Caption         =   "Convenio Foncodes"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   10620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   7200
      TabIndex        =   7
      Top             =   4800
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   8700
      TabIndex        =   6
      Top             =   4800
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5115
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   345
         Left            =   3720
         TabIndex        =   1
         Top             =   180
         Width           =   1275
      End
      Begin MSMask.MaskEdBox txtFechaini 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   210
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechafin 
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   210
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Fechas"
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
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "al"
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
         Height          =   315
         Left            =   2100
         TabIndex        =   4
         Top             =   240
         Width           =   765
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FgridFoncodes 
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   8
      AllowUserResizing=   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblcapacitacion 
      Alignment       =   1  'Right Justify
      Caption         =   "0.0"
      Height          =   255
      Left            =   8760
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblfdo 
      Alignment       =   1  'Right Justify
      Caption         =   "0.0"
      Height          =   255
      Left            =   7080
      TabIndex        =   13
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblgastos 
      Alignment       =   1  'Right Justify
      Caption         =   "0.0"
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblinteres 
      Alignment       =   1  'Right Justify
      Caption         =   "0.0"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblcapital 
      Alignment       =   1  'Right Justify
      Caption         =   "0.0"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Total:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "frmRptFoncodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConect    As New DConecta



Private Sub cmdProcesar_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim I As Integer
On Error GoTo errores
    FgridFoncodes.Clear
    
    If ValidaFecha(txtFechaini.Text) <> "" Then
       MsgBox "Fecha Inicial no válida...!", vbInformation, "Aviso"
       txtFechaini.SetFocus
       Exit Sub
    End If

    If ValidaFecha(txtFechafin.Text) <> "" Then
       MsgBox "Fecha Final no válida...!", vbInformation, "Aviso"
       txtFechafin.SetFocus
       Exit Sub
    End If

   
If oConect.AbreConexion() Then
    sql = " Cnt_SelRptFoncodes_sp '" & Format(txtFechaini.Text, "mm/dd/yyyy") & "','" & Format(txtFechafin.Text, "mm/dd/yyyy") & "'"
    Set rs = oConect.Ejecutar(sql)
   
   End If

If rs.BOF Then
    Set rs = Nothing
    oConect.CierraConexion: Set oConect = Nothing
    cmdImprimir.Enabled = False
    MsgBox "No existen datos...", vbExclamation, "Aviso!!!"
    Exit Sub
Else
    Set FgridFoncodes.DataSource = rs
    For I = 0 To rs.RecordCount - 1
        lblcapital.Caption = CDbl(lblcapital.Caption) + rs!Capital
        lblinteres.Caption = CDbl(lblinteres.Caption) + rs!Interes
        lblgastos.Caption = CDbl(lblgastos.Caption) + rs!gastos
        lblfdo.Caption = CDbl(lblfdo.Caption) + rs!FdoRotat
        lblcapacitacion.Caption = CDbl(lblcapacitacion.Caption) + rs!Capacit
    Next I
    Set rs = Nothing
    oConect.CierraConexion: Set oConect = Nothing
    cmdImprimir.Enabled = True
End If
FormatoGrid
Exit Sub
errores:
 MsgBox Err.Description

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    txtFechaini.Text = gdFecSis
    txtFechafin.Text = gdFecSis
    cmdImprimir.Enabled = False
    CentraForm Me
End Sub

Private Sub FormatoGrid()
    With FgridFoncodes
        .TextMatrix(0, 0) = "#"
        .TextMatrix(0, 1) = "Mes"
        .TextMatrix(0, 2) = "Año"
        .TextMatrix(0, 3) = "Capital"
        .TextMatrix(0, 4) = "Interés"
        .TextMatrix(0, 5) = "Gastos Operativos"
        .TextMatrix(0, 6) = "Capitaliz. Fdo. Rotat."
        .TextMatrix(0, 7) = "Capacit. Asist Técn."
        
        .ColWidth(0) = 250
        .ColWidth(1) = 800
        .ColWidth(2) = 800
        .ColWidth(3) = 1600
        .ColWidth(4) = 1600
        .ColWidth(5) = 1600
        .ColWidth(6) = 1600
        .ColWidth(7) = 1600
   
        .ColAlignmentFixed(0) = 0
        .ColAlignmentFixed(1) = 0
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 9
     End With
End Sub

Private Sub cmdImprimir_Click()
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo As String
    Dim N As Integer

    lsArchivo = App.path & "\Spooler\FONCODES" & "_" & Left(Format(txtFechafin.Text, gsFormatoMovFecha), 6) & ".xls"
    
    ExcelBegin lsArchivo, xlAplicacion, xlLibro, True
    ExcelAddHoja Replace(txtFechafin.Text, "/", "-"), xlLibro, xlHoja1, True
    Call CabeceraExcel(xlHoja1)
    For N = 1 To FgridFoncodes.Rows - 1
        xlHoja1.Cells(N + 7, 1) = FgridFoncodes.TextMatrix(N, 1)
        xlHoja1.Cells(N + 7, 2) = FgridFoncodes.TextMatrix(N, 2)
        xlHoja1.Cells(N + 7, 3) = FgridFoncodes.TextMatrix(N, 3)
        xlHoja1.Cells(N + 7, 4) = FgridFoncodes.TextMatrix(N, 4)
        xlHoja1.Cells(N + 7, 5) = FgridFoncodes.TextMatrix(N, 5)
        xlHoja1.Cells(N + 7, 6) = FgridFoncodes.TextMatrix(N, 6)
        xlHoja1.Cells(N + 7, 7) = FgridFoncodes.TextMatrix(N, 7)
    Next
    ExcelCuadro xlHoja1, 1, 7, 7, N + 7
    xlHoja1.Cells(N + 7, 2) = "TOTALES"
    xlHoja1.Cells(N + 7, 3) = lblcapital.Caption
    xlHoja1.Cells(N + 7, 4) = lblinteres.Caption
    xlHoja1.Cells(N + 7, 5) = lblgastos.Caption
    xlHoja1.Cells(N + 7, 6) = lblfdo.Caption
    xlHoja1.Cells(N + 7, 7) = lblcapacitacion.Caption
    
        
     xlHoja1.Range(xlHoja1.Cells(N + 7, 2), xlHoja1.Cells(N + 7, 7)).Font.Bold = True
     ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    CargaArchivo lsArchivo, ""

End Sub


Private Sub CabeceraExcel(hoja_actual As Excel.Worksheet)
    Dim nCol As Integer
    Dim sCol As String
    hoja_actual.Cells(1, 1) = gsNomCmac
    hoja_actual.Cells(1, 2) = gsUser
    hoja_actual.Cells(3, 1) = "CONVENIO FONCODES"
    hoja_actual.Cells(4, 2) = "Del " & txtFechaini.Text & " al " & txtFechafin.Text
    hoja_actual.Cells(3, 1).Font.Bold = True
    
    hoja_actual.Range("A3:G3").Merge
    hoja_actual.Range("A4:G4").Merge
    hoja_actual.Range("A3:G4").HorizontalAlignment = xlHAlignCenter
    hoja_actual.Range("A3:G3").Font.Size = 13
    
    hoja_actual.Cells(7, 1) = "Mes"
    hoja_actual.Cells(7, 2) = "Año"
    hoja_actual.Cells(7, 3) = "Capital Recuperado"
    hoja_actual.Cells(7, 4) = "Intereses"
    hoja_actual.Cells(7, 5) = "Gastos Operativos"
    hoja_actual.Cells(7, 6) = "Capitaliz. Fdo. Rotat."
    hoja_actual.Cells(7, 7) = "Capacit. Asist. Tecn"
    
    hoja_actual.Range("A1:A1").ColumnWidth = 8
    hoja_actual.Range("B1:B1").ColumnWidth = 8
    hoja_actual.Range("C1:C1").ColumnWidth = 19
    hoja_actual.Range("D1:D1").ColumnWidth = 19
    hoja_actual.Range("E1:E1").ColumnWidth = 19
    hoja_actual.Range("F1:F1").ColumnWidth = 19
    hoja_actual.Range("G1:G1").ColumnWidth = 19
    
    hoja_actual.Columns("C:G").Select
    Selection.NumberFormat = "0.0000"
    
    hoja_actual.Columns("A:B").Select
    Selection.NumberFormat = "General"
    
End Sub

Private Sub txtFechafin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        cmdProcesar_Click
    End If
End Sub

Private Sub txtFechaini_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
         txtFechafin.SetFocus
    End If
End Sub
