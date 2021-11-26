VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptDiaLiquidez 
   Caption         =   "Reporte Diario de Liquidez"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpatrimonio 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   4200
      TabIndex        =   13
      Text            =   "0"
      Top             =   240
      Width           =   1155
   End
   Begin VB.CommandButton cmdImprimirCaja 
      Caption         =   "Saldos &Agencia"
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtTipCambioFD 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   2280
      Width           =   1155
   End
   Begin VB.TextBox txtTipCambioFM 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   1800
      Width           =   1155
   End
   Begin VB.TextBox txtTipCambioV 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   1320
      Width           =   1155
   End
   Begin VB.TextBox txtTipCambioC 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   840
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimirBco 
      Caption         =   "Saldos &Banco"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin MSMask.MaskEdBox txtFechaini 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   -2147483635
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
   Begin VB.Label Label6 
      Caption         =   "Patrimonio:"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   3240
      TabIndex        =   14
      Top             =   240
      Width           =   885
   End
   Begin VB.Label Label5 
      Caption         =   "Tipo de Cambio Fijo del Día:"
      Height          =   315
      Left            =   960
      TabIndex        =   11
      Top             =   2280
      Width           =   2205
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo de Cambio Fijo del Mes:"
      Height          =   315
      Left            =   960
      TabIndex        =   10
      Top             =   1800
      Width           =   2205
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de Cambio Venta:"
      Height          =   315
      Left            =   960
      TabIndex        =   9
      Top             =   1320
      Width           =   2205
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo de Cambio Compra:"
      Height          =   315
      Left            =   960
      TabIndex        =   8
      Top             =   840
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "Saldos al:"
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
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   885
   End
End
Attribute VB_Name = "frmRptDiaLiquidez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ntipcamb As New nTipoCambio
Dim oConect    As New DConecta
Dim sql As String
Dim cOper  As String

    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
Private Sub cmdImprimirBco_Click()
    Dim rsBco As ADODB.Recordset
    On Error GoTo errores
    If ValidaFecha(txtFechaini.Text) <> "" Then
       MsgBox "Fecha no válida...!", vbInformation, "Aviso"
       txtFecha.SetFocus
       Exit Sub
    End If
    
    If CDbl(txtpatrimonio.Text) = 0# Then
       MsgBox "Patrimonio No Válido...!", vbInformation, "Aviso"
       Exit Sub
    End If
    
    If CDbl(txtTipCambioFD.Text) = 0# Then
       Respuesta = MsgBox("Tipo de Cambio del día No Válido...!. Desea Continuar", vbOKCancel, "Aviso")
       If Respuesta <> 1 Then
            Exit Sub
        End If
    End If
        
    If oConect.AbreConexion() Then
       sql = " Cnt_SelRptDiaLiquidBco_sp '" & Format(txtFechaini.Text, "mm/dd/yyyy") & "'," & txtTipCambioFD.Text
       Set rsBco = oConect.Ejecutar(sql)
    End If
        
    If rsBco.BOF Then
       Set rsBco = Nothing
       oConect.CierraConexion: Set oConect = Nothing
       MsgBox "No existen datos para generar el reporte", vbExclamation, "Aviso!!!"
       Exit Sub
    Else
       Call Imprime_SaldosBco(rsBco)
    End If
        
Exit Sub
errores:
         MsgBox Err.Description
End Sub

Private Sub cmdImprimirCaja_Click()
    Dim rsCaja As ADODB.Recordset
   

    On Error GoTo errores
    If ValidaFecha(txtFechaini.Text) <> "" Then
       MsgBox "Fecha no válida...!", vbInformation, "Aviso"
       txtFecha.SetFocus
       Exit Sub
    End If
    
    
    If CDbl(txtTipCambioFD.Text) = 0# Then
       Respuesta = MsgBox("Tipo de Cambio del día No Válido...!. Desea Continuar", vbOKCancel, "Aviso")
       If Respuesta <> 1 Then
            Exit Sub
        End If
    End If
    
    If oConect.AbreConexion() Then
            sql = "Cnt_InsRptDiaLiquidCaja_sp '" & Format(txtFechaini.Text, "yyyymmdd") & "'," & txtTipCambioFD.Text
            Set rsCaja = oConect.CargaRecordSet(sql)
            
       sql = " Cnt_SelRptDiaLiquidCaja_sp '" & Format(txtFechaini.Text, "yyyymmdd") & "'," & txtTipCambioFD.Text
       Set rsCaja = oConect.Ejecutar(sql)
    End If
    If rsCaja.BOF Then
       Set rsCaja = Nothing
       oConect.CierraConexion: Set oConect = Nothing
       MsgBox "No existen datos para generar el reporte", vbExclamation, "Aviso!!!"
       Exit Sub
    Else
       Call Imprime_SaldosCaja(rsCaja)
    End If
Exit Sub
errores:
         MsgBox Err.Description
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Public Sub Inicio(Optional Operacion As String)
    cOper = Operacion
    If Operacion = OpeCGRepSaldoBcos Then
        cmdImprimirBco.Enabled = True
        cmdImprimirCaja.Enabled = False
        
        cmdImprimirBco.Visible = True
        cmdImprimirCaja.Visible = False
        
        txtpatrimonio.Enabled = True
    Else
        cmdImprimirBco.Enabled = False
        cmdImprimirCaja.Enabled = True
        
        cmdImprimirBco.Visible = False
        cmdImprimirCaja.Visible = True
        
        txtpatrimonio.Enabled = False
End If
    Me.Show 1
End Sub
Private Sub Form_Load()
    'Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CentraForm Me
    txtFechaini.Text = DateAdd("d", -1, gdFecSis)
    TipoCambio
End Sub
Private Sub TipoCambio()
    Dim rs As ADODB.Recordset
    If ValidaFecha(txtFechaini.Text) <> "" Then
       MsgBox "Fecha no válida...!", vbInformation, "Aviso"
       txtFechaini.SetFocus
       Exit Sub
    End If
    
    Set rs = New ADODB.Recordset
    Set rs = ntipcamb.SelecTipoCambio(txtFechaini.Text)
    
    If Not rs.EOF And Not rs.BOF Then
        txtTipCambioC.Text = rs!nValComp
        txtTipCambioV.Text = rs!nValVent
        txtTipCambioFM.Text = rs!nValFijo
        txtTipCambioFD.Text = rs!nValFijoDia
    End If
    rs.Close
    Set rs = Nothing
End Sub


Private Sub txtFechaini_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        TipoCambio
        If cOper = OpeCGRepSaldoBcos Then
            txtpatrimonio.SetFocus
        Else
            cmdImprimirCaja.SetFocus
        End If
    End If
    
End Sub

Private Sub txtpatrimonio_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        
            cmdImprimirBco.SetFocus
    End If
End Sub
Private Sub txtFechaini_LostFocus()
        TipoCambio
End Sub

Private Sub Imprime_SaldosBco(rs As Recordset)
    Dim filas_Count As Integer
    Dim Row As Integer
    Dim Fila As Integer
    Dim cod_Bco As String
    Dim matriz_bco()
    Dim I As Integer
    
    Dim total_soles As Double
    Dim total_dolares As Double
    Dim total_solesTC As Double
    
    Dim T_soles As Double
    Dim T_dolares As Double
    Dim T_solesTC As Double
    
    Dim lsArchivo As String
    Dim lbLibroOpen As Boolean
    Dim Hoja1 As Excel.Worksheet

       
On Error Resume Next
'exportando a excell
    I = 0
    lsArchivo = App.path & "\Spooler\RptDiaLiquidez" & Format(txtFechaini.Text, "ddmmyyyy") & ".XLS"
    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
    If lbLibroOpen Then
        Set Hoja1 = xlLibro.Worksheets(1)
        ExcelAddHoja "Saldos_Banco", xlLibro, Hoja1
        filas_Count = rs.RecordCount
        Row = 11
        ReDim matriz_bco(filas_Count, 2)

cod_Bco = rs.Fields(0).value
 For Fila = 0 To filas_Count
        If CStr(rs.Fields(0).value) = CStr(cod_Bco) Then
            If linc_reglon = 0 Then
                linc_reglon = 1
                Hoja1.Cells(Fila + Row, 1) = CStr(rs.Fields(0).value)
                Hoja1.Cells(Fila + Row, 2) = rs.Fields(1).value
                fila_total = Fila + Row
                Row = Row + 1
                matriz_bco(I, 0) = rs.Fields(1).value
            End If
                Hoja1.Cells(Fila + Row, 2) = rs.Fields(2).value
                Hoja1.Cells(Fila + Row, 3) = rs.Fields(3).value
                Hoja1.Cells(Fila + Row, 4) = rs.Fields(4).value
                Hoja1.Cells(Fila + Row, 5) = rs.Fields(5).value
                Hoja1.Cells(Fila + Row, 6) = rs.Fields(6).value
                
            total_soles = total_soles + CDbl(rs.Fields(4).value)
            total_dolares = total_dolares + CDbl(rs.Fields(5).value)
            total_solesTC = total_solesTC + CDbl(rs.Fields(6).value)
            
            T_soles = T_soles + CDbl(rs.Fields(4).value)
            T_dolares = T_dolares + CDbl(rs.Fields(5).value)
            T_solesTC = T_solesTC + CDbl(rs.Fields(6).value)
       
            
        Else
            
            Hoja1.Cells(fila_total, 4) = total_soles
            Hoja1.Cells(fila_total, 5) = total_dolares
            Hoja1.Cells(fila_total, 6) = total_solesTC
            matriz_bco(I + 1, 0) = rs.Fields(1).value
            matriz_bco(I, 1) = total_solesTC
            I = I + 1
                        
            total_soles = 0
            total_dolares = 0
            total_solesTC = 0
            fila_total = Fila + Row
            
            Hoja1.Cells(Fila + Row, 1) = CStr(rs.Fields(0).value)
            Hoja1.Cells(Fila + Row, 2) = rs.Fields(1).value
            Row = Row + 1
            Hoja1.Cells(Fila + Row, 2) = rs.Fields(2).value
            Hoja1.Cells(Fila + Row, 3) = rs.Fields(3).value
            Hoja1.Cells(Fila + Row, 4) = rs.Fields(4).value
            Hoja1.Cells(Fila + Row, 5) = rs.Fields(5).value
            Hoja1.Cells(Fila + Row, 6) = rs.Fields(6).value
            
            total_soles = total_soles + CDbl(rs.Fields(4).value)
            total_dolares = total_dolares + CDbl(rs.Fields(5).value)
            total_solesTC = total_solesTC + CDbl(rs.Fields(6).value)
            
            T_soles = T_soles + CDbl(rs.Fields(4).value)
            T_dolares = T_dolares + CDbl(rs.Fields(5).value)
            T_solesTC = T_solesTC + CDbl(rs.Fields(6).value)


            cod_Bco = CStr(rs.Fields(0).value)
        End If
            rs.MoveNext
        Next Fila
            
            Hoja1.Cells(fila_total, 4) = total_soles
            Hoja1.Cells(fila_total, 5) = total_dolares
            Hoja1.Cells(fila_total, 6) = total_solesTC
            
            Hoja1.Cells(Fila + Row - 1, 2) = "TOTAL BANCOS"
            Hoja1.Cells(Fila + Row - 1, 4) = T_soles
            Hoja1.Cells(Fila + Row - 1, 5) = T_dolares
            Hoja1.Cells(Fila + Row - 1, 6) = T_solesTC
            
            matriz_bco(I, 0) = rs.Fields(1).value
            matriz_bco(I, 1) = total_solesTC
           
            Call CabeceraSaldoBco_Excell(Hoja1, Fila + Row - 1)
          
        Set rs = Nothing
           'imprime limite patrimonial
             Fila = Fila + Row + 1
             Hoja1.Range(Hoja1.Cells(Fila, 1), Hoja1.Cells(Fila, 6)).Merge
             Hoja1.Range(Hoja1.Cells(Fila, 1), Hoja1.Cells(Fila, 6)).Font.Bold = True
             Hoja1.Range(Hoja1.Cells(Fila, 1), Hoja1.Cells(Fila, 6)).Font.Name = "Arial"
             Hoja1.Range(Hoja1.Cells(Fila, 1), Hoja1.Cells(Fila, 6)).HorizontalAlignment = xlCenter
             Hoja1.Range(Hoja1.Cells(Fila, 1), Hoja1.Cells(Fila, 6)) = "LIMITE PATRIMONIAL"
             Fila = Fila + 2
             
            Hoja1.Cells(Fila, 2) = "BANCOS"
            Hoja1.Cells(Fila, 3) = "PATRIMONIO EFECT. =" & txtpatrimonio.Text & "x30%"
            Hoja1.Cells(Fila, 4) = "SALDOS"
            Hoja1.Cells(Fila, 5) = "DIFERENCIA"
            Hoja1.Cells(Fila, 3).WrapText = True
            Hoja1.Range(Hoja1.Cells(Fila, 2), Hoja1.Cells(Fila, 5)).Font.Bold = True
            Hoja1.Range(Hoja1.Cells(Fila, 2), Hoja1.Cells(Fila, 5)).Font.Name = "Arial"
            Hoja1.Range(Hoja1.Cells(Fila, 2), Hoja1.Cells(Fila, 5)).HorizontalAlignment = xlCenter
            Hoja1.Range(Hoja1.Cells(Fila, 2), Hoja1.Cells(Fila, 5)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic

            Fila = Fila + 1
            For N = 0 To I
                Hoja1.Cells(Fila + N, 2) = matriz_bco(N, 0)
                Hoja1.Cells(Fila + N, 3) = CDbl(txtpatrimonio.Text) * 0.3
                Hoja1.Cells(Fila + N, 4) = matriz_bco(N, 1)
                Hoja1.Cells(Fila + N, 5) = Hoja1.Cells(Fila + N, 3) - Hoja1.Cells(Fila + N, 4)
                Hoja1.Range(Hoja1.Cells(Fila + N, 2), Hoja1.Cells(Fila + N, 5)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
                Hoja1.Range("B" & Fila + N & ":E" & Fila + N).Borders(xlInsideVertical).LineStyle = xlContinuous
            Next
            
        oConect.CierraConexion: Set oConect = Nothing
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, Hoja1
        CargaArchivo "RptDiaLiquidez" & Format(txtFechaini.Text, "ddmmyyyy") & ".XLS", App.path & "\Spooler"
    End If
End Sub

Private Sub CabeceraSaldoBco_Excell(pxhoja As Excel.Worksheet, Fila As Integer)
'para la cabecera
    pxhoja.Cells(1, 1) = gsNomCmac
    pxhoja.Cells(1, 1).Font.Bold = True

     
     pxhoja.Range(pxhoja.Cells(3, 1), pxhoja.Cells(3, 6)).Merge
     pxhoja.Range(pxhoja.Cells(3, 1), pxhoja.Cells(3, 6)).Font.Bold = True
     pxhoja.Range(pxhoja.Cells(3, 1), pxhoja.Cells(3, 6)).Font.Name = "Arial"
     pxhoja.Range(pxhoja.Cells(3, 1), pxhoja.Cells(3, 6)).HorizontalAlignment = xlCenter
     pxhoja.Range(pxhoja.Cells(3, 1), pxhoja.Cells(3, 6)) = "REPORTE DIARIO DE LIQUIDEZ"
    

    pxhoja.Cells(8, 1) = "Saldos al: " & txtFechaini.Text
    pxhoja.Cells(8, 1).Font.Bold = True

    pxhoja.Cells(4, 1) = "T.C. Compra : " & txtTipCambioC
    pxhoja.Cells(5, 1) = "T.C. Venta   : " & txtTipCambioV
    pxhoja.Cells(6, 1) = "T.C. Fijo      : " & txtTipCambioFD
            
    pxhoja.Cells(10, 2) = "Banco - Cta Bancaria"
    pxhoja.Cells(10, 3) = "Tipo de Cuenta""Soles"
    pxhoja.Cells(10, 4) = "Soles"
    pxhoja.Cells(10, 5) = "Dolares"
    pxhoja.Cells(10, 6) = "Total en Soles"
    
    pxhoja.Range(pxhoja.Cells(10, 1), pxhoja.Cells(10, 6)).Font.Bold = True
     pxhoja.Range(pxhoja.Cells(10, 1), pxhoja.Cells(10, 6)).Font.Name = "Arial"
     pxhoja.Range(pxhoja.Cells(10, 1), pxhoja.Cells(10, 6)).HorizontalAlignment = xlCenter

    pxhoja.Columns("A:F").Font.Name = "ARIAL"
    pxhoja.Columns("A:F").Font.Size = 8

'para el contenido
    pxhoja.Columns("A:A").ColumnWidth = 4
    pxhoja.Columns("B:B").ColumnWidth = 17
    pxhoja.Columns("C:C").ColumnWidth = 20
    pxhoja.Columns("D:D").ColumnWidth = 12
    pxhoja.Columns("E:E").ColumnWidth = 12
    pxhoja.Columns("F:F").ColumnWidth = 12



 pxhoja.Range("D:F").NumberFormat = "#,##0.0000_);(#,##0.0000)"
 pxhoja.Range("A:C").NumberFormat = "General"
    
    ' haciendo cuadros
   pxhoja.Range(pxhoja.Cells(10, 1), pxhoja.Cells(10, 6)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic

'    pxhoja.Range("A10:" & "F" & fila).Select
'    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    With Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlInsideVertical)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlInsideHorizontal)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
  
    pxhoja.Range("A" & Fila & ":" & "F" & Fila).Select
    Selection.Font.Bold = True
    
    'IMPRIME LA FECHA DE HOY
    pxhoja.Cells(1, 6) = gdFecSis
    pxhoja.Cells(1, 6).NumberFormat = "m/d/yyyy"
    
     
End Sub

Private Sub Imprime_SaldosCaja(rsCaja As Recordset)
    Dim filas_Count As Integer
    Dim Row As Integer
    Dim Fila As Integer
    Dim cod_agencia As Integer
    Dim Row_ini  As Integer
    
    Dim total_soles As Double
    Dim total_dolares As Double
    Dim total_solesTC As Double
    
    Dim T_solesA As Double
    Dim T_dolaresA As Double
    Dim T_solesTCA As Double
    
    Dim lsArchivo As String
    Dim lbLibroOpen As Boolean
   Dim Hoja_caja As Excel.Worksheet
    

On Error Resume Next
'exportando a excell

cod_agencia = 1
 MousePointer = 11
       
          
       
    lsArchivo = App.path & "\Spooler\RptDiaLiquidez" & Format(txtFechaini.Text, "ddmmyyyy") & ".XLS"
    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
    If lbLibroOpen Then
        Set Hoja_caja = xlLibro.Worksheets(1)
        ExcelAddHoja "Saldos_Caja", xlLibro, Hoja_caja
        filas_Count = rsCaja.RecordCount
         Row = 11
         Row_ini = Row
        For Fila = 0 To filas_Count
        If rsCaja.Fields(0).value = cod_agencia Then
            If linc_reglon = 0 Then
                linc_reglon = 1
                Hoja_caja.Cells(Fila + Row, 1) = rsCaja.Fields(1).value
                Row = Row + 1
            End If
                Hoja_caja.Cells(Fila + Row, 1) = rsCaja.Fields(2).value
                Hoja_caja.Cells(Fila + Row, 2) = rsCaja.Fields(3).value
                Hoja_caja.Cells(Fila + Row, 3) = rsCaja.Fields(4).value
                Hoja_caja.Cells(Fila + Row, 4) = rsCaja.Fields(5).value
                
            total_soles = total_soles + CDbl(rsCaja.Fields(3).value)
            total_dolares = total_dolares + CDbl(rsCaja.Fields(4).value)
            total_solesTC = total_solesTC + CDbl(rsCaja.Fields(5).value)
          
            T_solesA = T_solesA + CDbl(rsCaja.Fields(3).value)
            T_dolaresA = T_dolaresA + CDbl(rsCaja.Fields(4).value)
            T_solesTCA = T_solesTCA + CDbl(rsCaja.Fields(5).value)

        Else
            Hoja_caja.Cells(Fila + Row, 1) = "Saldo Efectivo"
            Hoja_caja.Cells(Fila + Row, 2) = total_soles
            Hoja_caja.Cells(Fila + Row, 3) = total_dolares
            Hoja_caja.Cells(Fila + Row, 4) = total_solesTC
            Row = Row + 1
            
            total_soles = 0
            total_dolares = 0
            total_solesTC = 0
            Hoja_caja.Cells(Fila + Row, 1) = rsCaja.Fields(1).value
            Row = Row + 1
            Hoja_caja.Cells(Fila + Row, 1) = rsCaja.Fields(2).value
            Hoja_caja.Cells(Fila + Row, 2) = rsCaja.Fields(3).value
            Hoja_caja.Cells(Fila + Row, 3) = rsCaja.Fields(4).value
            Hoja_caja.Cells(Fila + Row, 4) = rsCaja.Fields(5).value
            
            total_soles = total_soles + CDbl(rsCaja.Fields(3).value)
            total_dolares = total_dolares + CDbl(rsCaja.Fields(4).value)
            total_solesTC = total_solesTC + CDbl(rsCaja.Fields(5).value)
     
            T_solesA = T_solesA + CDbl(rsCaja.Fields(3).value)
            T_dolaresA = T_dolaresA + CDbl(rsCaja.Fields(4).value)
            T_solesTCA = T_solesTCA + CDbl(rsCaja.Fields(5).value)

            cod_agencia = rsCaja.Fields(0).value
        End If
            rsCaja.MoveNext
        Next Fila
            Hoja_caja.Cells(Fila + Row - 1, 1) = "Saldo Efectivo"
            Hoja_caja.Cells(Fila + Row - 1, 2) = total_soles
            Hoja_caja.Cells(Fila + Row - 1, 3) = total_dolares
            Hoja_caja.Cells(Fila + Row - 1, 4) = total_solesTC
        
            Hoja_caja.Cells(Fila + Row, 1) = "TOTAL AGENCIAS"
            Hoja_caja.Cells(Fila + Row, 2) = T_solesA
            Hoja_caja.Cells(Fila + Row, 3) = T_dolaresA
            Hoja_caja.Cells(Fila + Row, 4) = T_solesTCA
            
    Call CabeceraSaldoCaja_Excell(Hoja_caja, Row_ini, Fila + Row)
      

    Set rsCaja = Nothing
        oConect.CierraConexion: Set oConect = Nothing
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, Hoja_caja
        CargaArchivo "RptDiaLiquidez" & Format(txtFechaini.Text, "ddmmyyyy") & ".XLS", App.path & "\Spooler"
    End If



   MousePointer = 0

End Sub

Private Sub CabeceraSaldoCaja_Excell(pxhoja As Excel.Worksheet, fila_min As Integer, Fila As Integer)
'para la cabecera

    pxhoja.Cells(1, 1) = gsNomCmac
    pxhoja.Cells(1, 1).Font.Bold = True

     
     pxhoja.Range(pxhoja.Cells(2, 1), pxhoja.Cells(2, 4)).Merge
     pxhoja.Range(pxhoja.Cells(2, 1), pxhoja.Cells(2, 4)).Font.Bold = True
     pxhoja.Range(pxhoja.Cells(2, 1), pxhoja.Cells(2, 4)).Font.Name = "Arial"
     pxhoja.Range(pxhoja.Cells(2, 1), pxhoja.Cells(2, 4)).HorizontalAlignment = xlCenter
     pxhoja.Range(pxhoja.Cells(2, 1), pxhoja.Cells(2, 4)) = "SALDOS DE CAJA AGENCIAS"
    

    pxhoja.Cells(8, 1) = "Saldos al: " & txtFechaini.Text
    pxhoja.Cells(8, 1).Font.Bold = True

    pxhoja.Cells(4, 1) = "T.C. Compra : " & txtTipCambioC
    pxhoja.Cells(5, 1) = "T.C. Venta   : " & txtTipCambioV
    pxhoja.Cells(6, 1) = "T.C. Fijo      : " & txtTipCambioFD
            
    pxhoja.Cells(10, 1) = "Movimientos Diarios"
    pxhoja.Cells(10, 2) = "Soles"
    pxhoja.Cells(10, 3) = "Dolares"
    pxhoja.Cells(10, 4) = "Total en Soles"
    
    pxhoja.Range(pxhoja.Cells(10, 1), pxhoja.Cells(10, 4)).Font.Bold = True
     pxhoja.Range(pxhoja.Cells(10, 1), pxhoja.Cells(10, 4)).Font.Name = "Arial"
     pxhoja.Range(pxhoja.Cells(10, 1), pxhoja.Cells(10, 4)).HorizontalAlignment = xlCenter
    

'PARA EL CUERPO
 
 pxhoja.Columns("A:A").ColumnWidth = 25
 pxhoja.Range("B:D").ColumnWidth = 17.29
 pxhoja.Range("B:D").NumberFormat = "#,##0.0000_);(#,##0.0000)"


' graficando las lineas
  'LINEAS PARA LA CABECERA
     pxhoja.Range(pxhoja.Cells(10, 1), pxhoja.Cells(10, 4)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
   
While fila_min <= Fila - 6
    'CUADRO PARA LAS OPERACIONES POR AGENCIA
    pxhoja.Range(pxhoja.Cells(fila_min, 1), pxhoja.Cells(fila_min + 6, 4)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic

    ' CUADRITO PARA SALDO EN EFECTIVO
    pxhoja.Range(pxhoja.Cells(fila_min + 6, 1), pxhoja.Cells(fila_min + 6, 4)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    
    'LINEAS VERTICALES
     pxhoja.Range("A" & fila_min & ":D" & Fila - 1).Borders(xlInsideVertical).LineStyle = xlContinuous

    'TITULO DE CADA AGENCIA EN NEGRITA
    pxhoja.Cells(fila_min, 1).Font.Bold = True
    pxhoja.Cells(fila_min, 1).HorizontalAlignment = xlCenter
    'SALDO EFECTIVO CENTRADO
    pxhoja.Cells(fila_min + 6, 1).HorizontalAlignment = xlCenter

        fila_min = fila_min + 7
Wend
    
    
'se adiciona fecha
    pxhoja.Cells(1, 4) = gdFecSis
    pxhoja.Cells(1, 4).NumberFormat = "m/d/yyyy"
    


End Sub



