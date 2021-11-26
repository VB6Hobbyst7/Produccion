VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAnxCalculoRCambiario 
   Caption         =   "Cálculo Riesgo Cambiario"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   345
      Left            =   240
      TabIndex        =   2
      Top             =   1230
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   345
      Left            =   3600
      TabIndex        =   3
      Top             =   1230
      Width           =   1000
   End
   Begin MSMask.MaskEdBox txtFecFin 
      Height          =   300
      Left            =   3240
      TabIndex        =   1
      Top             =   570
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
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
   Begin VB.Frame frmRangFec 
      Caption         =   "Rango de Fechas"
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
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   4335
      Begin MSMask.MaskEdBox txtFecIni 
         Height          =   300
         Left            =   960
         TabIndex        =   0
         Top             =   330
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
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
      Begin VB.Label lblDesde 
         Caption         =   "Desde :"
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
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblHasta 
         Caption         =   "Hasta :"
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
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha "
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
      Left            =   720
      TabIndex        =   4
      Top             =   645
      Width           =   855
   End
End
Attribute VB_Name = "frmAnxCalculoRCambiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*** Nombre : frmAnxCalculoRCambiario
'*** Descripción : Formulario para generar el Reporte de Posiciones Afecta a Riesgo Cambiario
'*** Creación : NAGL el 20170823
'********************************************************************************

Dim rs As New ADODB.Recordset
Dim DAnxRies As New DAnexoRiesgos

Public Sub Inicio()
Me.txtFecIni = gdFecSis
Me.txtFecFin = gdFecSis
CentraForm Me
Me.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub txtFecIni_GotFocus()
    fEnfoque txtFecIni
End Sub
Private Sub txtFecIni_KeyPress(keyAscii As Integer)
  If keyAscii = 13 Then
     If ValFecha(txtFecIni) Then
        txtFecFin.SetFocus
     End If
  End If
End Sub
Private Sub txtFecFin_GotFocus()
    fEnfoque txtFecFin
End Sub
Private Sub txtFecFin_KeyPress(keyAscii As Integer)
  If keyAscii = 13 Then
     If ValFecha(txtFecFin) Then
        cmdGenerar.SetFocus
     End If
  End If
End Sub
Private Function ValidaFecha(pdFecIni As Date, pdFecFin As Date) As Boolean
If pdFecIni > gdFecSis Then
   MsgBox "La Fecha de Inicio es mayor a la Fecha de Sistema", vbInformation, "Atención"
   txtFecIni.SetFocus
   Exit Function
ElseIf pdFecFin > gdFecSis Then
   MsgBox "La Fecha de Término es mayor a la Fecha de Sistema", vbInformation, "Atención"
   txtFecFin.SetFocus
   Exit Function
End If
If pdFecIni > pdFecFin Then
   MsgBox "La Fecha de Inicio es mayor a la Fecha de Término", vbInformation, "Atención"
   txtFecFin.SetFocus
   Exit Function
End If
ValidaFecha = True
End Function

Private Sub cmdGenerar_Click()
Dim pdFechaIni As Date
Dim pdFechaFin As Date
    If ValFecha(txtFecIni) And ValFecha(txtFecFin) Then
        pdFechaIni = txtFecIni
        pdFechaFin = txtFecFin
        If ValidaFecha(pdFechaIni, pdFechaFin) Then 'Valida Datos con respecto al Rango Ingresado
            Call GenerarAnxCalculoRiesgoCambiario(pdFechaIni, pdFechaFin)
        End If
    End If
End Sub

Private Sub GenerarAnxCalculoRiesgoCambiario(pdFechaIni As Date, pdFechaFin As Date)
Dim fs As Scripting.FileSystemObject
Dim lbExisteHoja As Boolean
Dim lsArchivo1 As String
Dim lsNomHoja  As String
Dim lsArchivo As String
Dim xlsAplicacion As Excel.Application
Dim xlsLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim TituloProgress As String
Dim MensajeProgress As String
Dim oBarra As clsProgressBar
Dim nprogress As Integer

Dim lilineas As Long
Dim liInicio As Long
Dim nCorrelativo As Long, Cant As Long
Dim lilineasCol As Long
Dim lsCadena() As String
Dim ActME As String, PasME As String, PosCam As String, ReqPExRC As String, PatrimEfect As String
Dim PosCambBal As String
Dim CantPosCamb As Long
Dim cCtaCnt As String
Dim rsCtaCnt As New ADODB.Recordset
Dim rsLim As New ADODB.Recordset
Dim pdFecFinMesNew As Date
Dim pdFecIniMesNew As Date
Dim VarI() As String, VarII() As String, VarIII() As String, VarIIIDH() As String, VARH As String, VarIV() As String
ReDim lsCadena(2)
ReDim VarI(2)
ReDim VarII(2)
ReDim VarIII(2)
ReDim VarIIIDH(3)
ReDim VarIV(2)
On Error GoTo GeneraExcelErr

 Set oBarra = New clsProgressBar
    Unload Me
    oBarra.ShowForm frmReportes
    oBarra.Max = 100
    nprogress = 0
    oBarra.Progress nprogress, "Anexo: Posiciones Afecta a Riesgo Cambiario", "GENERANDO EL ARCHIVO", "", vbBlue
    TituloProgress = "Anexo: Posiciones Afecta a Riesgo Cambiario"
    MensajeProgress = "GENERANDO EL ARCHIVO"

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "AnxCalculoRCambiario"
    'Primera Hoja ******************************************************
    'CON RESPECTO
    lsNomHoja = "POS_CAM"
    '*******************************************************************
    lsArchivo1 = "\spooler\ANEXO_CalculoRiesgoCamb_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    
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
    
    lilineas = 4
    liInicio = lilineas
    oBarra.Progress 10, TituloProgress, MensajeProgress, "", vbBlue
    nprogress = 10
    Set rs = DAnxRies.DevueleReporteCalculoRCambiario(pdFechaIni, pdFechaFin, "POSCAM")
    CantPosCamb = rs.RecordCount
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            'IIICUADRO
            xlHoja1.Cells(lilineas, 16) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha
            xlHoja1.Cells(lilineas, 17) = Format(rs!ActivosME, "#,##0.00") 'Activos ME
            ActME = xlHoja1.Range(xlHoja1.Cells(lilineas, 17), xlHoja1.Cells(lilineas, 17)).Address(False, False)
            
            xlHoja1.Cells(lilineas, 18) = Format(rs!PasivosME, "#,##0.00") 'Pasivos ME
            PasME = xlHoja1.Range(xlHoja1.Cells(lilineas, 18), xlHoja1.Cells(lilineas, 18)).Address(False, False)
            
            xlHoja1.Cells(lilineas, 19).Formula = "=" & "ABS" & "(" & ActME & "-" & PasME & ")" 'PosCam
            xlHoja1.Cells(lilineas, 19).NumberFormat = "#,##0.00"
            PosCam = xlHoja1.Range(xlHoja1.Cells(lilineas, 19), xlHoja1.Cells(lilineas, 19)).Address(False, False)
            
            xlHoja1.Cells(lilineas, 21) = Format(rs!Factor, "#,##0.00%") 'Factor
            xlHoja1.Cells(lilineas, 20).Formula = "=" & "+" & PosCam & "*" & xlHoja1.Range(xlHoja1.Cells(lilineas, 21), xlHoja1.Cells(lilineas, 21)).Address(False, False)
            xlHoja1.Cells(lilineas, 20).NumberFormat = "#,##0.00"
            ReqPExRC = xlHoja1.Range(xlHoja1.Cells(lilineas, 20), xlHoja1.Cells(lilineas, 20)).Address(False, False)
            '/IIIC
            
            xlHoja1.Cells(lilineas, 23) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha IVC
            xlHoja1.Cells(lilineas, 24).Formula = "=" & "+" & ReqPExRC & "/" & "1000"
            xlHoja1.Cells(lilineas, 24).NumberFormat = "#,##0" 'ReqPExRC IVC
            xlHoja1.Cells(lilineas, 25) = Format(rs!TipoCambio, "#,##0.000") 'TipoCamb IVC
            
            'ICUADRO
            xlHoja1.Cells(lilineas, 2) = Format(rs!dFecha, "mm/dd/yyyy")
            xlHoja1.Cells(lilineas, 3) = Round(xlHoja1.Range(xlHoja1.Cells(lilineas, 17), xlHoja1.Cells(lilineas, 17)) / 1000)
            xlHoja1.Cells(lilineas, 3).NumberFormat = "#,##0"
            
            xlHoja1.Cells(lilineas, 4) = Round(xlHoja1.Range(xlHoja1.Cells(lilineas, 18), xlHoja1.Cells(lilineas, 18)) / 1000)
            xlHoja1.Cells(lilineas, 4).NumberFormat = "#,##0"
            
            xlHoja1.Cells(lilineas, 5).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 3)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas, 4)).Address(False, False)
            xlHoja1.Cells(lilineas, 5).NumberFormat = "#,##0" 'PosCambBal
            PosCambBal = xlHoja1.Range(xlHoja1.Cells(lilineas, 5), xlHoja1.Cells(lilineas, 5)).Address(False, False)
            '/IC
            
            'IICuadro
            xlHoja1.Cells(lilineas, 12) = Format(rs!PatrimEfectivo, "#,##0.00")
            PatrimEfect = xlHoja1.Range(xlHoja1.Cells(lilineas, 12), xlHoja1.Cells(lilineas, 12)).Address(False, False)
            
            xlHoja1.Cells(lilineas, 13) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha
            'IICuadro
            
            'ICUADRO
            xlHoja1.Cells(lilineas, 6).Formula = "=" & PatrimEfect & "/" & "1000"
            xlHoja1.Cells(lilineas, 6).NumberFormat = "#,##0" 'PatrimEfect en Soles
            
            If rs!PatrimEfectivo = 0 Then
                xlHoja1.Cells(lilineas, 7) = 0
            Else
                xlHoja1.Cells(lilineas, 7).Formula = "=" & PosCambBal & "/" & xlHoja1.Range(xlHoja1.Cells(lilineas, 6), xlHoja1.Cells(lilineas, 6)).Address(False, False)
            End If
            xlHoja1.Cells(lilineas, 7).NumberFormat = "#,##0.00%" 'Pos/PE
            xlHoja1.Cells(lilineas, 8) = Format(rs!TipoCambio, "#,##0.000") 'TipoCamb
            
            If (lilineas = 4) Then
                xlHoja1.Cells(lilineas, 9) = Format(rs!VariacionActivo_mil, "#,##0")
                xlHoja1.Cells(lilineas, 10) = Format(rs!VariacionPasivo_mil, "#,##0")
            Else
                xlHoja1.Cells(lilineas, 9).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 3)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 3), xlHoja1.Cells(lilineas - 1, 3)).Address(False, False)
                xlHoja1.Cells(lilineas, 9).NumberFormat = "#,##0" 'VarPas
                xlHoja1.Cells(lilineas, 10).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas, 4)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 4), xlHoja1.Cells(lilineas - 1, 4)).Address(False, False)
                xlHoja1.Cells(lilineas, 10).NumberFormat = "#,##0" 'VarAct
            End If
            
            'IIC
                xlHoja1.Cells(lilineas, 14).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas, 7), xlHoja1.Cells(lilineas, 7)).Address(False, False)
                xlHoja1.Cells(lilineas, 14).NumberFormat = "#,##0.00%" 'PosCamMaynas
            '/IIC
            ExcelCuadro xlHoja1, 2, lilineas, 10, CCur(lilineas)
            ExcelCuadro xlHoja1, 12, lilineas, 14, CCur(lilineas)
            ExcelCuadro xlHoja1, 16, lilineas, 21, CCur(lilineas)
            ExcelCuadro xlHoja1, 23, lilineas, 25, CCur(lilineas)
            nCorrelativo = nCorrelativo + 1
            lilineas = lilineas + 1
            rs.MoveNext
         Loop
     
    End If
    Set rs = Nothing
    lsNomHoja = "NEGOCIO_CAMB"
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
    
    lilineas = 4
    liInicio = lilineas
    nCorrelativo = 1
    oBarra.Progress 30, TituloProgress, MensajeProgress, "", vbBlue
    
    Set rs = DAnxRies.DevueleReporteCalculoRCambiario(pdFechaIni, pdFechaFin, "NEGCAMB")
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            xlHoja1.Cells(lilineas, 3) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha
            xlHoja1.Cells(lilineas, 4) = Format(rs!Ganancias, "#,##0.00") 'Ganancias
            xlHoja1.Cells(lilineas, 5) = Format(rs!Perdidas, "#,##0.00") 'Perdidas
            
            xlHoja1.Cells(lilineas, 6).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas, 4)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas, 5), xlHoja1.Cells(lilineas, 5)).Address(False, False)
            xlHoja1.Cells(lilineas, 6).NumberFormat = "#,##0.00" 'Neto
            
            If lilineas >= 16 Then
                xlHoja1.Cells(lilineas, 7).Formula = "=" & "ABS" & "(" & xlHoja1.Range(xlHoja1.Cells(lilineas, 6), xlHoja1.Cells(lilineas, 6)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 6), xlHoja1.Cells(lilineas - 1, 6)).Address(False, False) & "-" & "1" & ")"
                xlHoja1.Cells(lilineas, 7).NumberFormat = "#,##0%" 'VariacionMensual
                xlHoja1.Cells(lilineas, 8).Formula = "=" & "ABS" & "(" & xlHoja1.Range(xlHoja1.Cells(lilineas, 6), xlHoja1.Cells(lilineas, 6)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(lilineas - 12, 6), xlHoja1.Cells(lilineas - 12, 6)).Address(False, False) & "-" & "1" & ")"
                xlHoja1.Cells(lilineas, 8).NumberFormat = "#,##0%" 'VariacionAnual
            End If
            
            If lilineas >= 40 Then 'IIC
               xlHoja1.Cells(lilineas, 10) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha
               xlHoja1.Cells(lilineas, 11).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas, 6), xlHoja1.Cells(lilineas, 6)).Address(False, False) & "/" & "1000"
               xlHoja1.Cells(lilineas, 11).NumberFormat = "#,##0" 'Ganancias en el Mes
               xlHoja1.Cells(lilineas, 13) = Format(rs!GananciaAcumTotal, "#,##0.00") 'AcumuladoExacto
               xlHoja1.Cells(lilineas, 12).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas, 13), xlHoja1.Cells(lilineas, 13)).Address(False, False) & "/" & "1000"
               xlHoja1.Cells(lilineas, 12).NumberFormat = "#,##0" 'Ganancias Acumuladas
               ExcelCuadro xlHoja1, 10, lilineas, 13, CCur(lilineas)
            End If
            
            xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 7)).Interior.ColorIndex = 15
            ExcelCuadro xlHoja1, 3, lilineas, 8, CCur(lilineas)
            lilineas = lilineas + 1
            rs.MoveNext
         Loop
     
    End If
    
    Set rs = Nothing
    'Cuadro 6 Resultados Acumulados: por Diferencia de Cambio
    lilineas = 45
    lilineasCol = 17
    liInicio = 0
    nCorrelativo = 1
    oBarra.Progress 50, TituloProgress, MensajeProgress, "", vbBlue

    Set rs = DAnxRies.DevueleReporteCalculoRCambiario(pdFechaIni, pdFechaFin, "ResulDifCamb")
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            If nCorrelativo = rs!Correlativo Then
                If nCorrelativo = 1 Then
                    xlHoja1.Cells(44, lilineasCol - 1 + nCorrelativo) = rs!Anio
                    ExcelCuadro xlHoja1, lilineasCol - 1 + nCorrelativo, lilineas - 1, lilineasCol - 1 + nCorrelativo, CCur(lilineas - 1)
                End If
                xlHoja1.Cells(lilineas, lilineasCol) = Format(rs!SaldoAcumulado, "#,##0.00")
                If rs!SaldoAcumulado < 0 Then
                   xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol), xlHoja1.Cells(lilineas, lilineasCol)).Font.Color = vbRed
                Else
                   xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol), xlHoja1.Cells(lilineas, lilineasCol)).Font.Color = vbBlack
                End If
                lilineasCol = lilineasCol + 1
            Else
                If xlHoja1.Cells(44, lilineasCol) = "" Then
                    xlHoja1.Cells(44, lilineasCol) = "Var Anual %"
                    liInicio = lilineasCol
                End If
                If xlHoja1.Cells(lilineas, liInicio - 1) <> "" Then
                    xlHoja1.Cells(lilineas, liInicio).Formula = "=" & "ABS" & "(" & xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol - 1), xlHoja1.Cells(lilineas, lilineasCol - 1)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol - 2), xlHoja1.Cells(lilineas, lilineasCol - 2)).Address(False, False) & "-" & "1" & ")"
                    xlHoja1.Cells(lilineas, liInicio).NumberFormat = "#,##0.00%" 'Var Anual(%)
                End If
                lilineas = lilineas + 1
                nCorrelativo = nCorrelativo + 1
                lilineasCol = 17
                xlHoja1.Cells(lilineas, lilineasCol) = Format(rs!SaldoAcumulado, "#,##0.00")
                If rs!SaldoAcumulado < 0 Then
                   xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol), xlHoja1.Cells(lilineas, lilineasCol)).Font.Color = vbRed
                Else
                   xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol), xlHoja1.Cells(lilineas, lilineasCol)).Font.Color = vbBlack
                End If
                lilineasCol = lilineasCol + 1
            End If
                If liInicio <> 0 Then
                    ExcelCuadro xlHoja1, 16, lilineas, liInicio, CCur(lilineas)
                Else
                    ExcelCuadro xlHoja1, 16, lilineas, lilineasCol, CCur(lilineas)
                End If
            rs.MoveNext
        Loop
    End If
    
    ExcelCuadro xlHoja1, 16, 43, liInicio, 43
    xlHoja1.Cells(43, 16) = "Resultados Acumulados: Diferencia de Cambio"
    xlHoja1.Range(xlHoja1.Cells(43, 16), xlHoja1.Cells(43, liInicio)).Merge True
    xlHoja1.Range(xlHoja1.Cells(43, 16), xlHoja1.Cells(44, liInicio)).Font.Color = vbWhite
    xlHoja1.Range(xlHoja1.Cells(43, 16), xlHoja1.Cells(56, liInicio)).Font.Name = "Arial"
    xlHoja1.Range(xlHoja1.Cells(43, 16), xlHoja1.Cells(56, liInicio)).Font.Size = 8
    xlHoja1.Range(xlHoja1.Cells(43, 16), xlHoja1.Cells(44, liInicio)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(43, 16), xlHoja1.Cells(44, liInicio)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(43, 16), xlHoja1.Cells(44, liInicio)).Interior.ColorIndex = 5
    Set rs = Nothing
    'Cuadro 7 Resultados Acumulados: Operaciones de Compra - Venta Spot
    lilineas = 64
    lilineasCol = 17
    liInicio = 0
    nCorrelativo = 1
    oBarra.Progress 60, TituloProgress, MensajeProgress, "", vbBlue
    
    Set rs = DAnxRies.DevueleReporteCalculoRCambiario(pdFechaIni, pdFechaFin, "ComVentSpot")
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            If nCorrelativo = rs!Correlativo Then
                If nCorrelativo = 1 Then
                    xlHoja1.Cells(63, lilineasCol - 1 + nCorrelativo) = rs!Anio
                    ExcelCuadro xlHoja1, lilineasCol - 1 + nCorrelativo, lilineas - 1, lilineasCol - 1 + nCorrelativo, CCur(lilineas - 1)
                End If
                xlHoja1.Cells(lilineas, lilineasCol) = Format(rs!SaldoAcumulado, "#,##0.00")
                If rs!SaldoAcumulado < 0 Then
                   xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol), xlHoja1.Cells(lilineas, lilineasCol)).Font.Color = vbRed
                Else
                   xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol), xlHoja1.Cells(lilineas, lilineasCol)).Font.Color = vbBlack
                End If
                lilineasCol = lilineasCol + 1
            Else
                If xlHoja1.Cells(63, lilineasCol) = "" Then
                    xlHoja1.Cells(63, lilineasCol) = "Var Anual %"
                    liInicio = lilineasCol
                End If
                If xlHoja1.Cells(lilineas, liInicio - 1) <> "" Then
                    xlHoja1.Cells(lilineas, liInicio).Formula = "=" & "ABS" & "(" & xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol - 1), xlHoja1.Cells(lilineas, lilineasCol - 1)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol - 2), xlHoja1.Cells(lilineas, lilineasCol - 2)).Address(False, False) & "-" & "1" & ")"
                    xlHoja1.Cells(lilineas, liInicio).NumberFormat = "#,##0.00%" 'Var Anual(%)
                End If
                lilineas = lilineas + 1
                nCorrelativo = nCorrelativo + 1
                lilineasCol = 17
                xlHoja1.Cells(lilineas, lilineasCol) = Format(rs!SaldoAcumulado, "#,##0.00")
                If rs!SaldoAcumulado < 0 Then
                   xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol), xlHoja1.Cells(lilineas, lilineasCol)).Font.Color = vbRed
                Else
                   xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol), xlHoja1.Cells(lilineas, lilineasCol)).Font.Color = vbBlack
                End If
                lilineasCol = lilineasCol + 1
            End If
                If liInicio <> 0 Then
                    ExcelCuadro xlHoja1, 16, lilineas, liInicio, CCur(lilineas)
                Else
                    ExcelCuadro xlHoja1, 16, lilineas, lilineasCol, CCur(lilineas)
                End If
            rs.MoveNext
        Loop
    End If
    
    ExcelCuadro xlHoja1, 16, 62, liInicio, 62
    xlHoja1.Cells(62, 16) = "Resultados Acumulados: Operaciones de Compra - Venta Spot"
    xlHoja1.Range(xlHoja1.Cells(62, 16), xlHoja1.Cells(62, liInicio)).Merge True
    xlHoja1.Range(xlHoja1.Cells(62, 16), xlHoja1.Cells(63, liInicio)).Font.Color = vbWhite
    xlHoja1.Range(xlHoja1.Cells(62, 16), xlHoja1.Cells(75, liInicio)).Font.Name = "Arial"
    xlHoja1.Range(xlHoja1.Cells(62, 16), xlHoja1.Cells(75, liInicio)).Font.Size = 8
    xlHoja1.Range(xlHoja1.Cells(62, 16), xlHoja1.Cells(63, liInicio)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(62, 16), xlHoja1.Cells(63, liInicio)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(62, 16), xlHoja1.Cells(63, liInicio)).Interior.ColorIndex = 5
    Set rs = Nothing

    lsNomHoja = "RESULTADOS_BALANCE" 'RESULTADO BALANCE
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
    
    lilineas = 1
    liInicio = lilineas
    Cant = 1
    oBarra.Progress 80, TituloProgress, MensajeProgress, "", vbBlue
    
    Set rsLim = DAnxRies.DevueleDetalleResultBalance(pdFechaIni, pdFechaFin, "", "SI")
    If Not (rsLim.EOF And rsLim.BOF) Then
        Do While Not rsLim.EOF
            pdFecIniMesNew = Format(rsLim!FecIni, "dd/mm/yyyy")
            pdFecFinMesNew = Format(rsLim!FecFin, "dd/mm/yyyy")
            
            xlHoja1.Cells(lilineas, 2) = "CMAC MAYNAS S.A."
            xlHoja1.Cells(lilineas, 6) = "Fecha: " & " " & gdFecSis
            lilineas = lilineas + 1
            xlHoja1.Cells(lilineas, 2) = "BALANCE DE COMPROBACIÓN (HISTORICO) CONSOLIDADO"
            xlHoja1.Range(xlHoja1.Cells(lilineas, 2), xlHoja1.Cells(lilineas, 7)).Merge True
            xlHoja1.Range(xlHoja1.Cells(lilineas, 2), xlHoja1.Cells(lilineas, 7)).HorizontalAlignment = xlCenter
            lilineas = lilineas + 1
            xlHoja1.Cells(lilineas, 2) = "Del" & " " & pdFecIniMesNew & " " & "Al" & " " & pdFecFinMesNew
            xlHoja1.Range(xlHoja1.Cells(lilineas, 2), xlHoja1.Cells(lilineas, 7)).Merge True
            xlHoja1.Range(xlHoja1.Cells(lilineas, 2), xlHoja1.Cells(lilineas, 7)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lilineas - 2, 2), xlHoja1.Cells(lilineas, 7)).Font.Bold = True
            lilineas = lilineas + 1
            xlHoja1.Cells(lilineas, 2) = "BALANCE DE COMPROBACIÓN (HISTORICO) CONSOLIDADO al" & " " & pdFecFinMesNew
            xlHoja1.Range(xlHoja1.Cells(lilineas, 2), xlHoja1.Cells(lilineas, 7)).Merge True
            xlHoja1.Range(xlHoja1.Cells(lilineas, 2), xlHoja1.Cells(lilineas, 7)).HorizontalAlignment = xlCenter
            ExcelCuadro xlHoja1, 2, lilineas, 7, CCur(lilineas)
            lilineas = lilineas + 1
            xlHoja1.Cells(lilineas, 2) = "DESCRIPCION"
            xlHoja1.Range(xlHoja1.Cells(lilineas, 2), xlHoja1.Cells(lilineas + 1, 2)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(lilineas, 2), xlHoja1.Cells(lilineas + 1, 2)).VerticalAlignment = xlCenter
            xlHoja1.Cells(lilineas, 3) = "CUENTA CONTABLE"
            xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas + 1, 3)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas + 1, 3)).VerticalAlignment = xlJustify
            xlHoja1.Cells(lilineas, 4) = "SALDO INICIAL"
            xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas + 1, 4)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas + 1, 4)).VerticalAlignment = xlJustify
            xlHoja1.Cells(lilineas, 5) = "MOVIMIENTO"
            xlHoja1.Range(xlHoja1.Cells(lilineas, 5), xlHoja1.Cells(lilineas, 6)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(lilineas, 5), xlHoja1.Cells(lilineas, 6)).VerticalAlignment = xlJustify
            xlHoja1.Cells(lilineas, 7) = "SALDO ACUMULADO"
            xlHoja1.Range(xlHoja1.Cells(lilineas, 7), xlHoja1.Cells(lilineas + 1, 7)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(lilineas, 7), xlHoja1.Cells(lilineas + 1, 7)).VerticalAlignment = xlJustify
            ExcelCuadro xlHoja1, 2, lilineas, 7, CCur(lilineas)
            lilineas = lilineas + 1
            xlHoja1.Cells(lilineas, 5) = "DEBE"
            xlHoja1.Cells(lilineas, 6) = "HABER"
            ExcelCuadro xlHoja1, 2, lilineas - 2, 7, CCur(lilineas)
            xlHoja1.Range(xlHoja1.Cells(lilineas - 2, 2), xlHoja1.Cells(lilineas, 7)).Interior.Color = RGB(153, 153, 255)
            xlHoja1.Range(xlHoja1.Cells(lilineas - 2, 2), xlHoja1.Cells(lilineas, 7)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(lilineas - 2, 2), xlHoja1.Cells(lilineas, 7)).HorizontalAlignment = xlCenter
            lilineas = lilineas + 1
            xlHoja1.Cells(lilineas, 2) = "PERDIDA"
            xlHoja1.Range(xlHoja1.Cells(lilineas, 2), xlHoja1.Cells(lilineas, 2)).Font.Bold = True
            lilineas = lilineas + 1
            
            Set rsCtaCnt = DAnxRies.ObtieneCuentaCntCalculoRCamb
            If Not (rsCtaCnt.EOF And rsCtaCnt.BOF) Then
                Do While Not (rsCtaCnt.EOF)
                    cCtaCnt = rsCtaCnt!cCtaContCod
                    Set rs = DAnxRies.DevueleDetalleResultBalance(pdFecIniMesNew, pdFecFinMesNew, cCtaCnt, "NO")
                    If Not (rs.BOF And rs.EOF) Then
                       xlHoja1.Cells(lilineas, 2) = rs!Descripcion
                       xlHoja1.Cells(lilineas, 3) = rs!CtaCnt
                       xlHoja1.Cells(lilineas, 4) = Format(rs!SaldoInicial, "#,##0.00")
                       xlHoja1.Cells(lilineas, 5) = Format(rs!nDebe, "#,##0.00")
                       xlHoja1.Cells(lilineas, 6) = Format(rs!nHaber, "#,##0.00")
                       xlHoja1.Cells(lilineas, 7) = Format(rs!SaldoAcum, "#,##0.00")
                       lilineas = lilineas + 1
                       
                       If cCtaCnt = "4108" Then
                          VarI(1) = xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 7), xlHoja1.Cells(lilineas - 1, 7)).Address(False, False)
                       ElseIf cCtaCnt = "5108" Then
                          VarI(2) = xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 7), xlHoja1.Cells(lilineas - 1, 7)).Address(False, False)
                          xlHoja1.Cells(lilineas - 9, 8).Formula = "=" & "+" & VarI(2) & "-" & VarI(1)
                          xlHoja1.Cells(lilineas - 9, 8).NumberFormat = "#,##0.00"
                       End If
                       
                       If cCtaCnt = "410801" Then
                          VarII(1) = xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 7), xlHoja1.Cells(lilineas - 1, 7)).Address(False, False)
                       ElseIf cCtaCnt = "510801" Then
                          VarII(2) = xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 7), xlHoja1.Cells(lilineas - 1, 7)).Address(False, False)
                          xlHoja1.Cells(lilineas - 9, 8).Formula = "=" & "+" & VarII(2) & "-" & VarII(1)
                          xlHoja1.Cells(lilineas - 9, 8).NumberFormat = "#,##0.00"
                       End If
                       
                       If cCtaCnt = "41080101" Then
                          xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 5), xlHoja1.Cells(lilineas - 1, 5)).Interior.ColorIndex = 44
                          VarIII(1) = xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 7), xlHoja1.Cells(lilineas - 1, 7)).Address(False, False)
                          VarIIIDH(1) = xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 5), xlHoja1.Cells(lilineas - 1, 5)).Address(False, False)
                          VarIIIDH(2) = xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 6), xlHoja1.Cells(lilineas - 1, 6)).Address(False, False)
                          xlHoja1.Cells(lilineas, 8).Formula = "=" & "+" & VarIIIDH(1) & "+" & VarIIIDH(2)
                          xlHoja1.Cells(lilineas, 8).NumberFormat = "#,##0.00"
                          xlHoja1.Range(xlHoja1.Cells(lilineas, 8), xlHoja1.Cells(lilineas, 8)).Interior.ColorIndex = 44
                          VarIIIDH(3) = xlHoja1.Range(xlHoja1.Cells(lilineas, 8), xlHoja1.Cells(lilineas, 8)).Address(False, False)
                       ElseIf cCtaCnt = "51080101" Then
                            xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 6), xlHoja1.Cells(lilineas - 1, 6)).Interior.ColorIndex = 44
                            VarIII(2) = xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 7), xlHoja1.Cells(lilineas - 1, 7)).Address(False, False)
                            xlHoja1.Cells(lilineas - 9, 8).Formula = "=" & "+" & VarIII(2) & "-" & VarIII(1)
                            xlHoja1.Cells(lilineas - 9, 8).NumberFormat = "#,##0.00"
                            VARH = xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 6), xlHoja1.Cells(lilineas - 1, 6)).Address(False, False)
                            xlHoja1.Cells(lilineas - 7, 8).Formula = "=" & "+" & VARH & "-" & VarIIIDH(3)
                            xlHoja1.Cells(lilineas - 7, 8).NumberFormat = "#,##0.00"
                            xlHoja1.Range(xlHoja1.Cells(lilineas - 7, 8), xlHoja1.Cells(lilineas - 7, 8)).Interior.ColorIndex = 44
                       ElseIf cCtaCnt = "41080409" Then
                            xlHoja1.Cells(lilineas, 2) = "GANANCIA"
                            xlHoja1.Range(xlHoja1.Cells(lilineas, 2), xlHoja1.Cells(lilineas, 2)).Font.Bold = True
                            lilineas = lilineas + 1
                       End If
                       
                      If cCtaCnt = "41080403" Then
                          VarIV(1) = xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 7), xlHoja1.Cells(lilineas - 1, 7)).Address(False, False)
                       ElseIf cCtaCnt = "51080403" Then
                          VarIV(2) = xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 7), xlHoja1.Cells(lilineas - 1, 7)).Address(False, False)
                          xlHoja1.Cells(lilineas - 10, 8).Formula = "=" & "+" & VarIV(2) & "-" & VarIV(1)
                          xlHoja1.Cells(lilineas - 10, 8).NumberFormat = "#,##0.00"
                       End If
                       
                    End If
                   rsCtaCnt.MoveNext
                Loop
            End If
            ExcelCuadro xlHoja1, 2, lilineas - 18, 7, CCur(lilineas - 1)
            lilineas = lilineas + 1
            rsLim.MoveNext
        Loop
    End If
    oBarra.Progress 90, TituloProgress, MensajeProgress, "", vbBlue
    Set rsLim = Nothing
    oBarra.Progress 100, "Anexo: Posiciones Afecta a Riesgo Cambiario", "Generación Terminada", "", vbBlue
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing

    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub






