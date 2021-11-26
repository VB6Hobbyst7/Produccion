VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAnx15B_PromedioMensRatioCobertLiqu 
   Caption         =   "Anexo: Promedio Mensual de RCL"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3795
   Icon            =   "frmAnx15BM_PromedioMensRatioCobertLiqu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   3795
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboMes 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame fraRango 
      Caption         =   "Rango"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3270
      Begin VB.TextBox txtAnio 
         Height          =   330
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblGuion 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1900
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1140
      Visible         =   0   'False
      Width           =   3275
      _ExtentX        =   5768
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmAnx15B_PromedioMensRatioCobertLiqu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*** Nombre : frmAnx15B_PromedioMensRatioCobertLiqu
'*** Descripción : Formulario para generar el Anexo del Promedio Mensual de Ratio de Cobertura de Liquidez (15B-M)
'*** Creación : NAGL el 20171226
'********************************************************************************
Dim pdFechaCentral As Date

Public Sub inicio()
    Dim pdFechaMesAnt As Date
    Dim oGen As New DGeneral
    Dim rs As New ADODB.Recordset
    Set rs = oGen.GetConstante(1010)
    While Not rs.EOF
       cboMes.AddItem rs.Fields(0) & space(50) & rs.Fields(1)
       rs.MoveNext
    Wend
    pdFechaMesAnt = DateAdd("d", -Day(gdFecSis), gdFecSis)
    Me.txtAnio.Text = Year(pdFechaMesAnt)
    cboMes.ListIndex = CInt(Month(pdFechaMesAnt)) - 1
    CentraForm Me
    Me.Show 1
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtAnio.SetFocus
End If
End Sub

Public Sub txtAnio_GotFocus()
    fEnfoque txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    cmdGenerar.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Function CalculaFechaFinMes() As Date
Dim sFechaParam  As Date
Dim pdFechaFinMes As Date
sFechaParam = "01/" & IIf(Len(Trim(cboMes.ListIndex + 2)) = 1, "0" & Trim(Str(cboMes.ListIndex + 2)), Trim(IIf(cboMes.ListIndex + 2 = 13, 1, cboMes.ListIndex + 2))) & "/" & IIf(cboMes.ListIndex + 2 = 13, Trim(CInt(txtAnio.Text) + 1), Trim(txtAnio.Text))
pdFechaFinMes = DateAdd("d", -1, sFechaParam)
CalculaFechaFinMes = pdFechaFinMes
End Function

Private Function ValidaFechaParam() As Boolean
If txtAnio <> "" And txtAnio <= Year(gdFecSis) And CInt(txtAnio) > 2000 Then
    pdFechaCentral = CalculaFechaFinMes
    If pdFechaCentral < "28/02/2014" Or pdFechaCentral > gdFecSis Then
        MsgBox "No existe información con el Rango Ingresado..!!", vbInformation, "Atención"
        txtAnio.SetFocus
        Exit Function
    End If
Else
    MsgBox "Debe Ingresar el año correspondiente..!!", vbInformation, "Aviso"
    txtAnio.SetFocus
    Exit Function
End If
ValidaFechaParam = True
End Function
Private Sub cmdGenerar_Click()
    If ValidaFechaParam() Then
        pdFechaCentral = CalculaFechaFinMes
        Call GeneraAnexo15BMPromedioMensualRCL(pdFechaCentral)
    End If
End Sub

Private Sub GeneraAnexo15BMPromedioMensualRCL(pdFecha As Date)
Dim fs As Scripting.FileSystemObject
Dim lbExisteHoja As Boolean
Dim lsArchivo1 As String
Dim lsNomHoja  As String
Dim lsArchivo As String
Dim xlsAplicacion As Excel.Application
Dim xlsLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim rsFec As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim DAnxRies As New DAnexoRiesgos
Dim FecParam As Date
Dim liLineas As Long, liInicio As Long
Dim lilineasCol As Long
Dim CantDay As Integer
Dim cantSep1 As Integer, cantSep2 As Integer '-->Para lineas de Separación entre la cabecera y Datos
Dim VarMN As Double, VarME As Double
Dim psNivRiesgoMN As String, psNivRiesgoME As String
Dim psNivRiesgoMNSinInterLiqu As String, psNivRiesgoMESinInterLiqu As String

On Error GoTo GeneraExcelErr

    PB1.Min = 0
    PB1.Max = 23
    PB1.value = 0
    cmdGenerar.Visible = False
    cmdCancelar.Visible = False
    PB1.Visible = True

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ANEXO_RCL_15BMENSUAL"
    'Primera Hoja ******************************************************
    'CON RESPECTO
    lsNomHoja = "CONTROL_RCL"
    '*******************************************************************
    lsArchivo1 = "\spooler\ANEXO_15BMPromMens_RCL_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
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
   
    PB1.value = 2
    
    xlHoja1.Cells(1, 1) = "PROMEDIO DE RCL AL " & Format(pdFecha, "dd/mm/yyyy")
    lilineasCol = 2
    CantDay = 1
    Set rsFec = DAnxRies.DevuelveAnexo15BRatioLiquidezDiario(pdFecha, "", "Fec")
    PB1.value = 5
    If Not (rsFec.EOF And rsFec.BOF) Then
        Do While Not rsFec.EOF
            FecParam = rsFec!dFecha
            liLineas = 8
            Set rs = DAnxRies.DevuelveAnexo15BRatioLiquidezDiario(FecParam, "15B")
                If Not (rs.BOF And rs.EOF) Then
                    Do While Not rs.EOF
                         xlHoja1.Range(xlHoja1.Cells(4, lilineasCol), xlHoja1.Cells(52, lilineasCol + 1)).Font.Name = "Arial Narrow"
                         xlHoja1.Range(xlHoja1.Cells(4, lilineasCol), xlHoja1.Cells(52, lilineasCol + 1)).Font.Size = 10
                         xlHoja1.Cells(liLineas, lilineasCol) = Format(rs!nSaldoMN, "#,##0.00")
                         xlHoja1.Cells(liLineas, lilineasCol + 1) = Format(rs!nSaldoME, "#,##0.00")
                         If liLineas = 8 Or liLineas = 26 Then 'NAGL Cambió de liLineas = 23 a liLineas = 26
                            If liLineas = 8 Then
                                xlHoja1.Cells(liLineas - 4, lilineasCol) = CantDay 'Nro Correlativo
                                xlHoja1.Range(xlHoja1.Cells(liLineas - 4, lilineasCol), xlHoja1.Cells(liLineas - 4, lilineasCol + 1)).Merge True
                                xlHoja1.Range(xlHoja1.Cells(liLineas - 4, lilineasCol), xlHoja1.Cells(liLineas - 4, lilineasCol + 1)).HorizontalAlignment = xlCenter
                                
                                xlHoja1.Cells(liLineas - 3, lilineasCol) = Format(FecParam, "mm/dd/yyyy") 'Fecha
                                xlHoja1.Range(xlHoja1.Cells(liLineas - 3, lilineasCol), xlHoja1.Cells(liLineas - 3, lilineasCol + 1)).Merge True
                                xlHoja1.Range(xlHoja1.Cells(liLineas - 3, lilineasCol), xlHoja1.Cells(liLineas - 3, lilineasCol + 1)).HorizontalAlignment = xlCenter
                                'ExcelCuadro xlHoja1, lilineasCol, liLineas - 3, lilineasCol + 1, CCur(liLineas - 3)
                                cantSep1 = 2
                                cantSep2 = 1
                            Else
                                cantSep1 = 3
                                cantSep2 = 2
                            End If
                            xlHoja1.Cells(liLineas - cantSep1, lilineasCol) = "Importe Ajustado"
                            xlHoja1.Range(xlHoja1.Cells(liLineas - cantSep1, lilineasCol), xlHoja1.Cells(liLineas - cantSep1, lilineasCol)).Font.Bold = True
                            xlHoja1.Range(xlHoja1.Cells(liLineas - cantSep1, lilineasCol), xlHoja1.Cells(liLineas - cantSep1, lilineasCol + 1)).Merge True
                            xlHoja1.Range(xlHoja1.Cells(liLineas - cantSep1, lilineasCol), xlHoja1.Cells(liLineas - cantSep1, lilineasCol + 1)).HorizontalAlignment = xlCenter
                            ExcelCuadro xlHoja1, lilineasCol, liLineas - cantSep1, lilineasCol + 1, CCur(liLineas - cantSep1)
                            
                            xlHoja1.Cells(liLineas - cantSep2, lilineasCol) = "MN (en PEN)"
                            xlHoja1.Cells(liLineas - cantSep2, lilineasCol + 1) = "MN (en USD)"
                            xlHoja1.Range(xlHoja1.Cells(liLineas - cantSep2, lilineasCol), xlHoja1.Cells(liLineas - cantSep2, lilineasCol + 1)).Font.Bold = True
                            ExcelCuadro xlHoja1, lilineasCol, liLineas - cantSep2, lilineasCol + 1, CCur(liLineas - cantSep2)
                            
                            xlHoja1.Range(xlHoja1.Cells(liLineas - cantSep1, lilineasCol), xlHoja1.Cells(liLineas - cantSep2, lilineasCol + 1)).Interior.Color = 16764057
                            If cantSep1 = 3 Then
                                xlHoja1.Range(xlHoja1.Cells(liLineas - 1, lilineasCol), xlHoja1.Cells(liLineas - 1, lilineasCol + 1)).Interior.ColorIndex = 15
                                'ExcelCuadro xlHoja1, lilineasCol, liLineas - 1, lilineasCol + 1, CCur(liLineas - 1)
                            End If
                       ElseIf liLineas = 20 Then
                                liLineas = liLineas + 1
                                xlHoja1.Cells(liLineas, lilineasCol).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(liLineas - 13, lilineasCol), xlHoja1.Cells(liLineas - 1, lilineasCol)).Address(False, False) & ")" 'SumaActivosLíquidosMN
                                xlHoja1.Cells(liLineas, lilineasCol + 1).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(liLineas - 13, lilineasCol + 1), xlHoja1.Cells(liLineas - 1, lilineasCol + 1)).Address(False, False) & ")" 'SumaActivosLíquidosME
                                xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol), xlHoja1.Cells(liLineas, lilineasCol + 1)).Font.Color = vbBlue
                                xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol), xlHoja1.Cells(liLineas, lilineasCol + 1)).Font.Bold = True
                                ExcelCuadro xlHoja1, lilineasCol, liLineas, lilineasCol + 1, CCur(liLineas)
                                liLineas = liLineas + 4
                                xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol), xlHoja1.Cells(liLineas, lilineasCol + 1)).Interior.ColorIndex = 15
                        ElseIf liLineas = 38 Then
                            liLineas = liLineas + 1
                            xlHoja1.Cells(liLineas, lilineasCol).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(liLineas - 13, lilineasCol), xlHoja1.Cells(liLineas - 1, lilineasCol)).Address(False, False) & ")" 'SumaFlujosEntrantesMN
                            xlHoja1.Cells(liLineas, lilineasCol + 1).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(liLineas - 13, lilineasCol + 1), xlHoja1.Cells(liLineas - 1, lilineasCol + 1)).Address(False, False) & ")" 'SumaFlujosEntrantesME
                            xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol), xlHoja1.Cells(liLineas, lilineasCol + 1)).Font.Color = vbBlue
                            xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol), xlHoja1.Cells(liLineas, lilineasCol + 1)).Font.Bold = True
                            ExcelCuadro xlHoja1, lilineasCol, liLineas, lilineasCol + 1, CCur(liLineas)
                            liLineas = liLineas + 1
                            xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol), xlHoja1.Cells(liLineas, lilineasCol + 1)).Interior.ColorIndex = 15
                        ElseIf liLineas = 61 Then
                            liLineas = liLineas + 1
                            xlHoja1.Cells(liLineas, lilineasCol).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(liLineas - 21, lilineasCol), xlHoja1.Cells(liLineas - 1, lilineasCol)).Address(False, False) & ")" 'SumaFlujosSalientesMN
                            xlHoja1.Cells(liLineas, lilineasCol + 1).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(liLineas - 21, lilineasCol + 1), xlHoja1.Cells(liLineas - 1, lilineasCol + 1)).Address(False, False) & ")" 'SumaFlujosSalientesME
                            xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol), xlHoja1.Cells(liLineas, lilineasCol + 1)).Font.Color = vbBlue
                            xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol), xlHoja1.Cells(liLineas, lilineasCol + 1)).Font.Bold = True
                            ExcelCuadro xlHoja1, lilineasCol, liLineas, lilineasCol + 1, CCur(liLineas)
                        ElseIf liLineas = 63 Then
                            xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol), xlHoja1.Cells(liLineas, lilineasCol + 1)).Font.Bold = True
                            ExcelCuadro xlHoja1, lilineasCol, liLineas, lilineasCol + 1, CCur(liLineas)
                            xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol), xlHoja1.Cells(liLineas, lilineasCol)).Interior.Color = RGB(153, 153, 255)
                        End If
                         'ElseIf liLineas = 34 Then
                            'xlHoja1.Range(xlHoja1.Cells(liLineas - 1, lilineasCol), xlHoja1.Cells(liLineas - 1, lilineasCol + 1)).Interior.ColorIndex = 15
                         'ElseIf liLineas = 46 Then
                            'liLineas = liLineas + 1
                            'xlHoja1.Cells(liLineas, lilineasCol) = Format(0, "#,##0.00")
                            'xlHoja1.Cells(liLineas, lilineasCol + 1) = Format(0, "#,##0.00")
                            'xlHoja1.Cells(liLineas + 1, lilineasCol) = Format(0, "#,##0.00")
                            'xlHoja1.Cells(liLineas + 1, lilineasCol + 1) = Format(0, "#,##0.00")
                            'xlHoja1.Cells(liLineas + 2, lilineasCol) = Format(0, "#,##0.00")
                            'xlHoja1.Cells(liLineas + 2, lilineasCol + 1) = Format(0, "#,##0.00")
                            'xlHoja1.Cells(liLineas + 3, lilineasCol) = Format(0, "#,##0.00")
                            'xlHoja1.Cells(liLineas + 3, lilineasCol + 1) = Format(0, "#,##0.00")
                            'liLineas = liLineas + 4
                        liLineas = liLineas + 1
                        rs.MoveNext
                    Loop
                End If
            ExcelCuadro xlHoja1, lilineasCol, 8, lilineasCol + 1, 18
            ExcelCuadro xlHoja1, lilineasCol, 19, lilineasCol + 1, 20
            ExcelCuadro xlHoja1, lilineasCol, 26, lilineasCol + 1, 38
            ExcelCuadro xlHoja1, lilineasCol, 41, lilineasCol + 1, 61
            xlHoja1.Range(xlHoja1.Cells(4, lilineasCol), xlHoja1.Cells(63, lilineasCol)).EntireColumn.AutoFit
            xlHoja1.Range(xlHoja1.Cells(4, lilineasCol + 1), xlHoja1.Cells(63, lilineasCol + 1)).EntireColumn.AutoFit
            lilineasCol = lilineasCol + 2
            CantDay = CantDay + 1
            rsFec.MoveNext
         Loop
    End If
    
    PB1.value = 10
    Set rs = Nothing
    liLineas = 68 'NAGL Cambió de 58 a 68
    liInicio = liLineas
    lilineasCol = 2
    CantDay = 1
    Set rs = DAnxRies.DevuelveAnexo15BRatioLiquidezDiario(FecParam, "", "Ratios")
        If Not (rs.BOF And rs.EOF) Then
            xlHoja1.Cells(liLineas - 1, lilineasCol) = "Fecha"
            xlHoja1.Cells(liLineas - 1, lilineasCol + 1) = "Porc."
            xlHoja1.Cells(liLineas - 1, lilineasCol + 2) = "RCL MN"
            xlHoja1.Cells(liLineas - 1, lilineasCol + 3) = "RCL ME"
            xlHoja1.Cells(liLineas - 1, lilineasCol + 4) = "Variac. MN"
            xlHoja1.Cells(liLineas - 1, lilineasCol + 5) = "Variac. ME"
            
            ExcelCuadro xlHoja1, lilineasCol, liLineas - 1, lilineasCol + 5, CCur(liLineas - 1)
            
            xlHoja1.Range(xlHoja1.Cells(liLineas - 1, lilineasCol), xlHoja1.Cells(liLineas - 1, lilineasCol + 5)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(liLineas - 1, lilineasCol), xlHoja1.Cells(liLineas - 1, lilineasCol + 5)).Interior.Color = RGB(153, 153, 255)
        
            Do While Not rs.EOF
                xlHoja1.Cells(liLineas, lilineasCol - 1) = CantDay
                xlHoja1.Cells(liLineas, lilineasCol) = Format(rs!dFecha, "mm/dd/yyyy")
                xlHoja1.Cells(liLineas, lilineasCol + 1) = "95%"
                xlHoja1.Cells(liLineas, lilineasCol + 2) = Format(rs!RatioMN, "#,##0.00")
                xlHoja1.Cells(liLineas, lilineasCol + 3) = Format(rs!RatioME, "#,##0.00")
                If liLineas >= 69 Then
                    xlHoja1.Cells(liLineas, lilineasCol + 4).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol + 2), xlHoja1.Cells(liLineas, lilineasCol + 2)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(liLineas - 1, lilineasCol + 2), xlHoja1.Cells(liLineas - 1, lilineasCol + 2)).Address(False, False)
                    VarMN = xlHoja1.Cells(liLineas, lilineasCol + 4)
                    xlHoja1.Cells(liLineas, lilineasCol + 5).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol + 3), xlHoja1.Cells(liLineas, lilineasCol + 3)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(liLineas - 1, lilineasCol + 3), xlHoja1.Cells(liLineas - 1, lilineasCol + 3)).Address(False, False)
                    VarME = xlHoja1.Cells(liLineas, lilineasCol + 5)
                    If VarMN < 0 Then
                       xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol + 4), xlHoja1.Cells(liLineas, lilineasCol + 4)).Font.Color = vbRed
                    Else
                       xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol + 4), xlHoja1.Cells(liLineas, lilineasCol + 4)).Font.Color = vbBlue
                    End If
                    If VarME < 0 Then
                       xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol + 5), xlHoja1.Cells(liLineas, lilineasCol + 5)).Font.Color = vbRed
                    Else
                       xlHoja1.Range(xlHoja1.Cells(liLineas, lilineasCol + 5), xlHoja1.Cells(liLineas, lilineasCol + 5)).Font.Color = vbBlue
                    End If
                End If
                
                ExcelCuadro xlHoja1, lilineasCol, liLineas, lilineasCol + 5, CCur(liLineas)
                CantDay = CantDay + 1
                liLineas = liLineas + 1
                rs.MoveNext
            Loop
        End If
        xlHoja1.Range(xlHoja1.Cells(liInicio, lilineasCol), xlHoja1.Cells(liInicio, lilineasCol)).EntireColumn.AutoFit
        liLineas = liLineas - 1
        xlHoja1.Cells(liLineas + 2, lilineasCol + 1) = "Min"
        xlHoja1.Cells(liLineas + 3, lilineasCol + 1) = "Max"
        xlHoja1.Cells(liLineas + 4, lilineasCol + 1) = "Prom"
        xlHoja1.Range(xlHoja1.Cells(liLineas + 2, lilineasCol + 1), xlHoja1.Cells(liLineas + 4, lilineasCol + 1)).Interior.Color = RGB(153, 153, 255)
        xlHoja1.Range(xlHoja1.Cells(liLineas + 2, lilineasCol + 1), xlHoja1.Cells(liLineas + 4, lilineasCol + 1)).Font.Bold = True
        
        xlHoja1.Cells(liLineas + 2, lilineasCol + 2).Formula = "=" & "Min" & "(" & xlHoja1.Range(xlHoja1.Cells(liInicio, lilineasCol + 2), xlHoja1.Cells(liLineas, lilineasCol + 2)).Address(False, False) & ")" 'Min MN
        xlHoja1.Cells(liLineas + 2, lilineasCol + 3).Formula = "=" & "Min" & "(" & xlHoja1.Range(xlHoja1.Cells(liInicio, lilineasCol + 3), xlHoja1.Cells(liLineas, lilineasCol + 3)).Address(False, False) & ")" 'Min ME
        ExcelCuadro xlHoja1, lilineasCol + 1, liLineas + 2, lilineasCol + 3, CCur(liLineas + 2)
        xlHoja1.Cells(liLineas + 3, lilineasCol + 2).Formula = "=" & "Max" & "(" & xlHoja1.Range(xlHoja1.Cells(liInicio, lilineasCol + 2), xlHoja1.Cells(liLineas, lilineasCol + 2)).Address(False, False) & ")" 'Max MN
        xlHoja1.Cells(liLineas + 3, lilineasCol + 3).Formula = "=" & "Max" & "(" & xlHoja1.Range(xlHoja1.Cells(liInicio, lilineasCol + 3), xlHoja1.Cells(liLineas, lilineasCol + 3)).Address(False, False) & ")" 'Max ME
        ExcelCuadro xlHoja1, lilineasCol + 1, liLineas + 3, lilineasCol + 3, CCur(liLineas + 3)
        xlHoja1.Cells(liLineas + 4, lilineasCol + 2).Formula = "=" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(liInicio, lilineasCol + 2), xlHoja1.Cells(liLineas, lilineasCol + 2)).Address(False, False) & ")" 'Prom MN
        xlHoja1.Cells(liLineas + 4, lilineasCol + 3).Formula = "=" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(liInicio, lilineasCol + 3), xlHoja1.Cells(liLineas, lilineasCol + 3)).Address(False, False) & ")" 'Prom ME
        ExcelCuadro xlHoja1, lilineasCol + 1, liLineas + 4, lilineasCol + 3, CCur(liLineas + 4)
        
        xlHoja1.Range(xlHoja1.Cells(liInicio, lilineasCol + 2), xlHoja1.Cells(liLineas + 4, lilineasCol + 5)).NumberFormat = "#,###0.00"
        xlHoja1.Range(xlHoja1.Cells(liInicio - 1, lilineasCol), xlHoja1.Cells(liLineas + 4, lilineasCol + 5)).HorizontalAlignment = xlCenter
        
        PB1.value = 15
        Set rs = Nothing
        
        'SIGUIENTE HOJA DE CÁLCULO - RCL_HIST
        lsNomHoja = "RCL_HIST"
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
        
     liLineas = 4 'NAGL Cambio de 3 a 4 20190521
     Set rs = DAnxRies.ObtieneControlDatosRCLHistorico(pdFecha)
     PB1.value = 20
     If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            xlHoja1.Cells(liLineas, 2) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha
            xlHoja1.Cells(liLineas, 3) = Format(rs!LimReg, "#,##0") 'LimReg
            xlHoja1.Cells(liLineas, 4) = Format(rs!LimInt, "#,##0") 'LimInt
            xlHoja1.Cells(liLineas, 5) = Format(rs!Alert, "#,##0") 'LimAlert
            xlHoja1.Cells(liLineas, 6) = Format(rs!RatioMNSinIntLiq, "#,##0.00") 'RatioMNSinInterCambLiqu
            xlHoja1.Cells(liLineas, 7) = Format(rs!RatioMESinIntLiq, "#,##0.00") 'RatioMESinInterCambLiqu
            '******Agregado by NAGL 20190521*******
            xlHoja1.Cells(liLineas, 8) = rs!cNivelRgoAsumMNSinIntLiq 'Nivel de Riesgo MN Sin InterCambLiqu
            xlHoja1.Cells(liLineas, 9) = rs!cNivelRgoAsumMESinIntLiq 'Nivel de Riesgo ME Sin InterCambLiqu
            xlHoja1.Cells(liLineas, 10) = Format(rs!nMontoInterCambLiqMN, "#,##0.00") 'LimInt
            xlHoja1.Cells(liLineas, 11) = Format(rs!nMontoInterCambLiqME, "#,##0.00") 'LimAlert
            xlHoja1.Cells(liLineas, 12) = Format(rs!RatioMN, "#,##0.00") 'RatioMN
            xlHoja1.Cells(liLineas, 13) = Format(rs!RatioME, "#,##0.00") 'RatioME
            xlHoja1.Cells(liLineas, 14) = rs!cNivelRgoAsumMN 'Nivel de Riesgo MN
            xlHoja1.Cells(liLineas, 15) = rs!cNivelRgoAsumME 'Nivel de Riesgo ME
            xlHoja1.Cells(liLineas, 16) = Format(rs!nTipoCambio, "#,##0.000") 'Tipo de Cambio
            
            psNivRiesgoMN = rs!cNivelRgoAsumMN
            psNivRiesgoME = rs!cNivelRgoAsumME
            psNivRiesgoMNSinInterLiqu = rs!cNivelRgoAsumMNSinIntLiq
            psNivRiesgoMESinInterLiqu = rs!cNivelRgoAsumMESinIntLiq
            '************END NAGL 20190521**********
            
            'If rs!CellMN = "R" Then
                'xlHoja1.Range(xlHoja1.Cells(liLineas, 9), xlHoja1.Cells(liLineas, 9)).Font.Color = vbRed
            'End If
            'If rs!CellME = "R" Then
                'xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 10)).Font.Color = vbRed
            'End If
            'If liLineas >= 4 Then
                'xlHoja1.Cells(liLineas, 11).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(liLineas, 9), xlHoja1.Cells(liLineas, 9)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(liLineas - 1, 9), xlHoja1.Cells(liLineas - 1, 9)).Address(False, False)
                'xlHoja1.Cells(liLineas, 12).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 10)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(liLineas - 1, 10), xlHoja1.Cells(liLineas - 1, 10)).Address(False, False)
                'xlHoja1.Range(xlHoja1.Cells(liLineas, 11), xlHoja1.Cells(liLineas, 12)).NumberFormat = "#,###0.00"
            'End If 'Comentado by NAGL 20190521
            
            '*************NAGL 20190521
            xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 11)).Font.Color = vbRed
            
            If psNivRiesgoMN = "Bajo" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 14), xlHoja1.Cells(liLineas, 14)).Interior.ColorIndex = 43
            ElseIf psNivRiesgoMN = "Moderado" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 14), xlHoja1.Cells(liLineas, 14)).Interior.ColorIndex = 6
            ElseIf psNivRiesgoMN = "Alto" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 14), xlHoja1.Cells(liLineas, 14)).Interior.ColorIndex = 44
            ElseIf psNivRiesgoMN = "Extremo" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 14), xlHoja1.Cells(liLineas, 14)).Interior.ColorIndex = 3
            End If
            
            If psNivRiesgoME = "Bajo" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 15), xlHoja1.Cells(liLineas, 15)).Interior.ColorIndex = 43
            ElseIf psNivRiesgoME = "Moderado" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 15), xlHoja1.Cells(liLineas, 15)).Interior.ColorIndex = 6
            ElseIf psNivRiesgoME = "Alto" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 15), xlHoja1.Cells(liLineas, 15)).Interior.ColorIndex = 44
            ElseIf psNivRiesgoME = "Extremo" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 15), xlHoja1.Cells(liLineas, 15)).Interior.ColorIndex = 3
            End If
            
              If psNivRiesgoMNSinInterLiqu = "Bajo" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 8), xlHoja1.Cells(liLineas, 8)).Interior.ColorIndex = 43
            ElseIf psNivRiesgoMNSinInterLiqu = "Moderado" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 8), xlHoja1.Cells(liLineas, 8)).Interior.ColorIndex = 6
            ElseIf psNivRiesgoMNSinInterLiqu = "Alto" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 8), xlHoja1.Cells(liLineas, 8)).Interior.ColorIndex = 44
            ElseIf psNivRiesgoMNSinInterLiqu = "Extremo" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 8), xlHoja1.Cells(liLineas, 8)).Interior.ColorIndex = 3
            End If
            
            If psNivRiesgoMESinInterLiqu = "Bajo" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 9), xlHoja1.Cells(liLineas, 9)).Interior.ColorIndex = 43
            ElseIf psNivRiesgoMESinInterLiqu = "Moderado" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 9), xlHoja1.Cells(liLineas, 9)).Interior.ColorIndex = 6
            ElseIf psNivRiesgoMESinInterLiqu = "Alto" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 9), xlHoja1.Cells(liLineas, 9)).Interior.ColorIndex = 44
            ElseIf psNivRiesgoMESinInterLiqu = "Extremo" Then
                xlHoja1.Range(xlHoja1.Cells(liLineas, 9), xlHoja1.Cells(liLineas, 9)).Interior.ColorIndex = 3
            End If
            '********************
            'xlHoja1.Range(xlHoja1.Cells(liLineas, 9), xlHoja1.Cells(liLineas, 10)).NumberFormat = "#,###0"
            xlHoja1.Range(xlHoja1.Cells(liLineas, 8), xlHoja1.Cells(liLineas, 9)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 14), xlHoja1.Cells(liLineas, 16)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 5)).HorizontalAlignment = xlCenter
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 16)).Font.Size = 9
            xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 16)).Font.Name = "Arial"
            ExcelCuadro xlHoja1, 2, liLineas, 16, CCur(liLineas)
            liLineas = liLineas + 1
            rs.MoveNext
         Loop
    End If
        
    PB1.value = 23
    Set rs = Nothing
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    PB1.Visible = False
    cmdGenerar.Visible = True
    cmdCancelar.Visible = True
Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
        
End Sub


