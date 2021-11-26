VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLibroInventBalanc 
   Caption         =   "Libro de Inventario y Balances al:"
   ClientHeight    =   1170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4170
   Icon            =   "frmLibroInventBalanc.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1170
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin MSMask.MaskEdBox txtfecha 
      Height          =   315
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Libro de Inventario y Balances al:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2400
   End
End
Attribute VB_Name = "frmLibroInventBalanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs   As New ADODB.Recordset
Dim oConect    As New DConecta
Dim cmoneda As String
Dim nMoneda As String

Private Sub cmdAceptar_Click()
Dim sql As String


On Error GoTo errores
    If ValidaFecha(txtFecha.Text) <> "" Then
       MsgBox "Fecha no válida...!", vbInformation, "Aviso"
       txtFecha.SetFocus
       Exit Sub
    End If


   
If oConect.AbreConexion() Then
    sql = " Cnt_SelInventBalance_sp '" & nMoneda & "','" & Format(txtFecha.Text, "mm/dd/yyyy") & "'"
    Set rs = oConect.Ejecutar(sql)
    
   End If

If rs.BOF Then
    Set rs = Nothing
    oConect.CierraConexion: Set oConect = Nothing
    MsgBox "No existen datos para generar el reporte", vbExclamation, "Aviso!!!"
    Exit Sub
Else
    rs.MoveFirst
    cmdAceptar.Enabled = False
    Call Imprime_libro
End If

Exit Sub
errores:
 MsgBox Err.Description


End Sub
Private Sub Imprime_libro()
    Dim Row As Integer
    Dim startRow As Integer
    Dim filas_Count As Integer
    Dim Fila As Integer
    Dim matriz()


    Dim Total_SUBCUENTAS As Double
    Dim Total_CUENTAS As Double
    Dim Total_CLASE As Double
    Dim Cuenta As String
    Dim clase As String
    Dim cuenta_tmp As String
    Dim clase_tmp As String
    
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim excelSheet As Excel.Worksheet
    Dim lsArchivo As String
    Dim lbLibroOpen As Boolean
     
On Error Resume Next
       
    startRow = 3
    filas_Count = rs.RecordCount
    Total_SUBCUENTAS = 0
    Total_CUENTAS = 0
    Total_CLASE = 0
    clase = "ACTIVO"
    Cuenta = ""
    clase_tmp = ""
    cuenta_tmp = ""
    
    ReDim Preserve matriz(filas_Count, 3)
        For Row = 0 To filas_Count
            matriz(Row, 0) = CStr(rs.Fields(0).value)
            matriz(Row, 1) = rs.Fields(1).value
            matriz(Row, 2) = rs.Fields(2).value
            rs.MoveNext
     Next Row
     
    Set rs = Nothing
    oConect.CierraConexion: Set oConect = Nothing

' ----------
'exportando a excell

    lsArchivo = App.path & "\Spooler\LibInventario" & Format(txtFecha.Text, "mmyyyy") & gsCodUser & nMoneda & ".XLS"
    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
    If lbLibroOpen Then
        ExcelAddHoja "Hoja1", xlLibro, excelSheet
        
       Call CabeceraExcel(excelSheet)
        
        ' llena soles
        Row = 0
        For Fila = startRow To filas_Count + startRow + 27
           Select Case (Len(Trim(matriz(Row + 1, 0))))
                Case 1
                 clase_tmp = clase
                 clase = matriz(Row + 1, 1)
                 If Fila > 4 Then
                     excelSheet.Cells(Fila, 1) = matriz(Row, 0)
                     excelSheet.Cells(Fila, 2) = matriz(Row, 1)
                     excelSheet.Cells(Fila, 3) = matriz(Row, 2)
                     If Len(Trim(matriz(Row + 2, 0))) = 2 Then
                        cuenta_tmp = Cuenta
                        Cuenta = matriz(Row + 1, 1)
                     
                         excelSheet.Cells(Fila + 1, 1) = ""
                         excelSheet.Cells(Fila + 1, 2) = ""
                         excelSheet.Cells(Fila + 1, 3) = "-------------------------"
                         excelSheet.Cells(Fila + 2, 3) = Total_SUBCUENTAS
                         excelSheet.Cells(Fila + 3, 3) = ""
                         
                         excelSheet.Cells(Fila + 4, 1) = ""
                         excelSheet.Cells(Fila + 4, 2) = "TOTAL " & cuenta_tmp
                         excelSheet.Cells(Fila + 4, 3) = Total_CUENTAS
        
                         Fila = Fila + 5
                         filas_Count = filas_Count + 5
                     End If
        
                     excelSheet.Cells(Fila + 2, 2).Font.Bold = True
                     excelSheet.Cells(Fila + 2, 3).Font.Bold = True
                     excelSheet.Cells(Fila + 2, 1) = ""
                     excelSheet.Cells(Fila + 2, 2) = "TOTAL " & clase_tmp
                     excelSheet.Cells(Fila + 2, 3) = Total_CLASE
    
                     Fila = Fila + 3
                     filas_Count = filas_Count + 3
                     Total_CLASE = 0
                  Else
                     
                     excelSheet.Cells(Fila, 1) = matriz(Row, 0)
                     excelSheet.Cells(Fila, 2) = matriz(Row, 1)
                     excelSheet.Cells(Fila, 3) = matriz(Row, 2)
                End If
                   
                   
                Case 2
                cuenta_tmp = Cuenta
                Cuenta = matriz(Row + 1, 1)
                  If Fila > 4 Then
                     excelSheet.Cells(Fila, 1) = matriz(Row, 0)
                     excelSheet.Cells(Fila, 2) = matriz(Row, 1)
                     excelSheet.Cells(Fila, 3) = matriz(Row, 2)
    
                 
                        If Len(Trim(matriz(Row, 0))) <> 1 Then
                             Total_SUBCUENTAS = Total_SUBCUENTAS + CDbl(matriz(Row, 2))
                             Total_CUENTAS = Total_CUENTAS + Total_SUBCUENTAS
                             Total_CLASE = Total_CLASE + Total_CUENTAS
                             excelSheet.Cells(Fila + 1, 1) = ""
                             excelSheet.Cells(Fila + 1, 2) = ""
                             excelSheet.Cells(Fila + 1, 3) = "-------------------------"
                             excelSheet.Cells(Fila + 2, 3) = Total_SUBCUENTAS
                             excelSheet.Cells(Fila + 3, 3) = ""
                             
                             excelSheet.Cells(Fila + 4, 1) = ""
                             excelSheet.Cells(Fila + 4, 2) = "TOTAL " & cuenta_tmp
                             excelSheet.Cells(Fila + 4, 3) = Total_CUENTAS
            
                             Fila = Fila + 5
                             filas_Count = filas_Count + 5
                             Total_CUENTAS = 0
                             Total_SUBCUENTAS = 0
                      End If
                
                  Else
    
                      excelSheet.Cells(Fila, 1) = matriz(Row, 0)
                      excelSheet.Cells(Fila, 2) = matriz(Row, 1)
                      excelSheet.Cells(Fila, 3) = matriz(Row, 2)
                End If
    
                Case 4
                    If Fila >= 4 Then
                            
                     excelSheet.Cells(Fila, 1) = matriz(Row, 0)
                     excelSheet.Cells(Fila, 2) = matriz(Row, 1)
                     excelSheet.Cells(Fila, 3) = matriz(Row, 2)
                     If Len(Trim(matriz(Row, 0))) > 4 Then
                                Total_SUBCUENTAS = Total_SUBCUENTAS + CDbl(matriz(Row, 2))
                                Total_CUENTAS = Total_CUENTAS + Total_SUBCUENTAS
                                excelSheet.Cells(Fila + 1, 1) = ""
                                excelSheet.Cells(Fila + 1, 2) = ""
                                excelSheet.Cells(Fila + 1, 3) = "-------------------------"
                                excelSheet.Cells(Fila + 2, 3) = Total_SUBCUENTAS
                                excelSheet.Cells(Fila + 3, 3) = ""
                                Fila = Fila + 3
                                filas_Count = filas_Count + 3
                                Total_SUBCUENTAS = 0
                     End If
                    Else
                        excelSheet.Cells(Fila, 1) = matriz(Row, 0)
                        excelSheet.Cells(Fila, 2) = matriz(Row, 1)
                        excelSheet.Cells(Fila, 3) = matriz(Row, 2)
                    End If
                 Case Else
                     excelSheet.Cells(Fila, 1) = matriz(Row, 0)
                     excelSheet.Cells(Fila, 2) = matriz(Row, 1)
                     excelSheet.Cells(Fila, 3) = matriz(Row, 2)
                     Total_SUBCUENTAS = Total_SUBCUENTAS + CDbl(matriz(Row, 2))
            End Select
            Row = Row + 1
        Next Fila
                    Total_CUENTAS = Total_CUENTAS + Total_SUBCUENTAS
                    Total_CLASE = Total_CLASE + Total_CUENTAS
                    excelSheet.Cells(Fila, 1) = ""
                    excelSheet.Cells(Fila, 2) = ""
                    excelSheet.Cells(Fila, 3) = "-------------------------"
                    excelSheet.Cells(Fila + 1, 3) = Total_SUBCUENTAS
                    excelSheet.Cells(Fila + 2, 3) = ""
                    excelSheet.Cells(Fila + 3, 1) = ""
                    excelSheet.Cells(Fila + 3, 2) = "TOTAL " & Cuenta
                    excelSheet.Cells(Fila + 3, 3) = Total_CUENTAS
                    excelSheet.Cells(Fila + 5, 2).Font.Bold = True
                    excelSheet.Cells(Fila + 5, 3).Font.Bold = True
                    excelSheet.Cells(Fila + 5, 1) = ""
                    excelSheet.Cells(Fila + 5, 2) = "TOTAL " & clase
                    excelSheet.Cells(Fila + 5, 3) = Total_CLASE
                    
                    
                    ' llena dolares
         ExcelEnd lsArchivo, xlAplicacion, xlLibro, excelSheet
         CargaArchivo "LibInventario" & Format(txtFecha.Text, "mmyyyy") & gsCodUser & nMoneda & ".XLS", App.path & "\Spooler"
    End If
    
    cmdAceptar.Enabled = True
    
End Sub

Private Sub CabeceraExcel(pxhoja As Excel.Worksheet)
 With pxhoja.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    pxhoja.PageSetup.PrintArea = ""

 With pxhoja.PageSetup
        .LeftHeader = "&""Arial,Negrita""" & gsNomCmac
        .CenterHeader = _
        "&""Arial,Negrita""&12LIBRO DE INVENTARIO Y BALANCES" & Chr(10) & "AL " & txtFecha.Text 'Format(txtfecha.Text, "d mmm aaaa")
        .RightHeader = "&""Arial,Negrita""&P" & Chr(10) & cmoneda
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.787401575)
        .RightMargin = Application.InchesToPoints(0.787401575)
        .TopMargin = Application.InchesToPoints(0.984251969)
        .BottomMargin = Application.InchesToPoints(0.984251969)
        .HeaderMargin = Application.InchesToPoints(0.61)
        .FooterMargin = Application.InchesToPoints(0)
        .FirstPageNumber = xlAutomatic
    End With
           
    pxhoja.Columns("A:A").ColumnWidth = 14.53
    pxhoja.Columns("B:B").ColumnWidth = 48
    pxhoja.Columns("C:C").ColumnWidth = 16.05
    
    pxhoja.Columns("A:A").Select
    Selection.NumberFormat = "0"
    pxhoja.Columns("B:B").Select
    Selection.NumberFormat = "General"
    pxhoja.Columns("C:C").Select
    Selection.NumberFormat = "0.00"
 

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

txtFecha.Text = gdFecSis
If gsSimbolo = gcME Then
    cmoneda = "Moneda Extranjera"
    nMoneda = "2"
Else
   cmoneda = "Moneda Nacional"
    nMoneda = "1"
End If
CentraForm Me

End Sub



Private Sub txtFecha_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   'cmdAceptar.SetFocus
   cmdAceptar_Click
End If
End Sub
