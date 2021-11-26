VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReporteSegTOSE 
   Caption         =   "Seguimiento TOSE"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   2155
      TabIndex        =   9
      Top             =   1560
      Width           =   1000
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   345
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1000
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
      ForeColor       =   &H8000000D&
      Height          =   1080
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2895
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cboMes 
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
         ItemData        =   "frmReporteSegTOSE.frx":0000
         Left            =   0
         List            =   "frmReporteSegTOSE.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtTC 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Top             =   240
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   285
         TabIndex        =   6
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio"
         Height          =   195
         Left            =   285
         TabIndex        =   5
         Top             =   720
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmReporteSegTOSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*** Nombre : frmReporteSegTOSE
'*** Descripción : Formulario para generar el Reporte Seguimiento TOSE
'*** Creación : NAGL el 20170912
'********************************************************************************

Dim rs As New ADODB.Recordset
Dim DAnxRies As New DAnexoRiesgos
Dim oDbalanceCont As New DbalanceCont
Dim TipoCamb As Currency
Public Sub Inicio()
Me.txtFecha = gdFecSis
txtTC = oDbalanceCont.ObtenerTipoCambioCierreNew(txtFecha, "TipoAct")
TipoCamb = txtTC
CentraForm Me
Me.Show 1
End Sub

Private Function ValidaFecha(pdFecha As Date) As Boolean
If pdFecha > gdFecSis Then
   MsgBox "La Fecha Ingresada es Incorrecta", vbInformation, "Atención"
   txtFecha.SetFocus
   Exit Function
End If
ValidaFecha = True
End Function

Private Sub cmdProcesar_Click()
Dim pdFecha As Date
    If ValFecha(txtFecha) Then
        pdFecha = txtFecha
        If ValidaFecha(pdFecha) Then 'Valida Datos con respecto a la Fecha Ingresada
            Call GenerarReporteSeguimientoTOSE(pdFecha)
        End If
    End If
End Sub

Private Sub GenerarReporteSeguimientoTOSE(pdFecha As Date)
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
Dim CountReg As Integer
Dim nCorrelativo As Long, Cant As Long
Dim lilineasCol As Long
Dim lsCadenaMN() As String
Dim lsCadenaME() As String
Dim ToseMN As String, ToseME As String, TipoCamb As String
Dim rsCtaCnt As New ADODB.Recordset

On Error GoTo GeneraExcelErr

 Set oBarra = New clsProgressBar
    Unload Me
    oBarra.ShowForm frmReportes
    oBarra.Max = 100
    nprogress = 0
    oBarra.Progress nprogress, "Anexo: Seguimiento TOSE", "GENERANDO EL ARCHIVO", "", vbBlue
    TituloProgress = "Anexo: Seguimiento TOSE"
    MensajeProgress = "GENERANDO EL ARCHIVO"

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "AnxSeguimientoTOSE"
    'Primera Hoja ******************************************************
    'CON RESPECTO
    lsNomHoja = "CONTROL_TOSE"
    '*******************************************************************
    lsArchivo1 = "\spooler\ANEXO_SeguimientoTOSE_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    
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
    
    ReDim lsCadenaMN(4)
    ReDim lsCadenaME(4)
    lilineas = 3
    liInicio = lilineas
    oBarra.Progress 10, TituloProgress, MensajeProgress, "", vbBlue
    nprogress = 10
    Set rs = DAnxRies.ObtieneReporteControlTOSEFinMes(pdFecha, "DatoFinMes")
    
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            'I CUADRO
            xlHoja1.Cells(lilineas, 2) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha
            xlHoja1.Cells(lilineas, 3) = Format(rs!TOSE_MN, "#,##0.00") 'TOSEMN
            ToseMN = xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 3)).Address(False, False)
            
            xlHoja1.Cells(lilineas, 4) = Format(rs!TOSE_ME, "#,##0.00") 'TOSEME
            ToseME = xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas, 4)).Address(False, False)
            '/I CUADRO
            
            xlHoja1.Cells(lilineas, 5) = Format(rs!TipoCambio, "#,##0.000") 'TipoCambio
            xlHoja1.Range(xlHoja1.Cells(lilineas, 5), xlHoja1.Cells(lilineas, 5)).HorizontalAlignment = xlRight
            TipoCamb = xlHoja1.Range(xlHoja1.Cells(lilineas, 5), xlHoja1.Cells(lilineas, 5)).Address(False, False)
            
            'II Cuadro
            xlHoja1.Cells(lilineas, 6).Formula = "=" & "+" & ToseMN & "+" & ToseME & "*" & TipoCamb 'Expr.Soles
            xlHoja1.Cells(lilineas, 6).NumberFormat = "#,###0.00"
            If lilineas >= 4 Then
                xlHoja1.Cells(lilineas, 7).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas, 6), xlHoja1.Cells(lilineas, 6)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 6), xlHoja1.Cells(lilineas - 1, 6)).Address(False, False)
                If xlHoja1.Range(xlHoja1.Cells(lilineas, 7), xlHoja1.Cells(lilineas, 7)) < 0 Then
                    xlHoja1.Range(xlHoja1.Cells(lilineas, 7), xlHoja1.Cells(lilineas, 7)).Font.Color = vbRed
                End If
                xlHoja1.Cells(lilineas, 7).NumberFormat = "#,###0.00"
            End If
            '/II Cuadro
            
             ExcelCuadro xlHoja1, 3, lilineas, 4, CCur(lilineas)
             ExcelCuadro xlHoja1, 6, lilineas, 7, CCur(lilineas)
            lilineas = lilineas + 1
            rs.MoveNext
         Loop
    End If
    
    Set rs = Nothing
    lilineas = 4
    oBarra.Progress 30, TituloProgress, MensajeProgress, "", vbBlue
    Set rs = DAnxRies.ObtieneReporteControlTOSEFinMes(pdFecha, "PromMesIII")
    xlHoja1.Cells(3, 8) = Format("31/10/2015", "mm/dd/yyyy")
    xlHoja1.Cells(3, 13) = Format("31/10/2015", "mm/dd/yyyy")
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            'III CUADRO PROMEDIO DEL MES
            xlHoja1.Cells(lilineas, 8) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha
            xlHoja1.Range(xlHoja1.Cells(lilineas, 8), xlHoja1.Cells(lilineas, 8)).HorizontalAlignment = xlRight
            
            xlHoja1.Cells(lilineas, 9) = Format(rs!MonedaNacional, "#,##0.00") 'TOSEMN Promedio
            ToseMN = xlHoja1.Range(xlHoja1.Cells(lilineas, 9), xlHoja1.Cells(lilineas, 9)).Address(False, False)
            
            xlHoja1.Cells(lilineas, 10) = Format(rs!MonedaExtranjera, "#,##0.00") 'TOSEME Promedio
            ToseME = xlHoja1.Range(xlHoja1.Cells(lilineas, 10), xlHoja1.Cells(lilineas, 10)).Address(False, False)
            
            xlHoja1.Cells(lilineas, 11).Formula = "=" & "+" & ToseMN & "+" & ToseME & "*" & xlHoja1.Range(xlHoja1.Cells(lilineas, 5), xlHoja1.Cells(lilineas, 5)).Address(False, False) 'Expr.Soles
            xlHoja1.Cells(lilineas, 11).NumberFormat = "#,###0.00"
            '/III CUADRO
            
            'IVCuadro Prom. Ult. tres meses
            xlHoja1.Cells(lilineas, 13) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha '/IV Cuadro
            
            If lilineas >= 5 Then
                xlHoja1.Cells(lilineas, 12).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas, 10), xlHoja1.Cells(lilineas, 10)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 10), xlHoja1.Cells(lilineas - 1, 10)).Address(False, False)
                xlHoja1.Cells(lilineas, 12).NumberFormat = "#,###0.00"
                If xlHoja1.Range(xlHoja1.Cells(lilineas, 10), xlHoja1.Cells(lilineas, 10)) < 0 Then
                    xlHoja1.Range(xlHoja1.Cells(lilineas, 10), xlHoja1.Cells(lilineas, 10)).Font.Color = vbRed
                End If
                'IVCuadro
                xlHoja1.Cells(lilineas, 14).Formula = "=" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(lilineas - 2, 11), xlHoja1.Cells(lilineas, 11)).Address(False, False) & ")" 'Expr.Soles
                xlHoja1.Cells(lilineas, 15).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas, 11), xlHoja1.Cells(lilineas, 11)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas, 14), xlHoja1.Cells(lilineas, 14)).Address(False, False) 'Variación
                xlHoja1.Range(xlHoja1.Cells(lilineas, 14), xlHoja1.Cells(lilineas, 15)).NumberFormat = "#,###0.00"
                '/IV Cuadro
            End If
            
            If lilineas >= 6 Then
               xlHoja1.Cells(lilineas, 16).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas, 11), xlHoja1.Cells(lilineas, 11)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 14), xlHoja1.Cells(lilineas - 1, 14)).Address(False, False) 'Diferencia
               If xlHoja1.Range(xlHoja1.Cells(lilineas, 16), xlHoja1.Cells(lilineas, 16)) < 0 Then
                  xlHoja1.Range(xlHoja1.Cells(lilineas, 16), xlHoja1.Cells(lilineas, 16)).Font.Color = vbRed
               End If
               xlHoja1.Cells(lilineas, 17).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas, 16), xlHoja1.Cells(lilineas, 16)).Address(False, False) & "/" & "1000000"
               xlHoja1.Range(xlHoja1.Cells(lilineas, 17), xlHoja1.Cells(lilineas, 17)).HorizontalAlignment = xlRight
               xlHoja1.Range(xlHoja1.Cells(lilineas, 16), xlHoja1.Cells(lilineas, 17)).NumberFormat = "#,###0.0000"
            End If
            
             ExcelCuadro xlHoja1, 9, lilineas, 11, CCur(lilineas)
             ExcelCuadro xlHoja1, 14, lilineas, 16, CCur(lilineas)
            lilineas = lilineas + 1
            rs.MoveNext
         Loop
    End If
    
    'SEGUIMIENTO TOSE EN EL DÍA
    Set rs = Nothing
    Cant = lilineas
    lilineas = lilineas + 3
    lilineasCol = 5
    nCorrelativo = 1
    Set rs = DAnxRies.DevuelveReporteControlTOSEDia(pdFecha)
    oBarra.Progress 50, TituloProgress, MensajeProgress, "", vbBlue
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            If nCorrelativo = rs!nSeccion Then
                     If nCorrelativo >= 2 Then
                        'Cuadro Adicional - Variación TOSE
                        CountReg = CountReg - 1
                        lilineasCol = lilineasCol - 2
                        xlHoja1.Cells(lilineas - 2, 57) = xlHoja1.Range(xlHoja1.Cells(lilineas - 1, lilineasCol), xlHoja1.Cells(lilineas - 1, lilineasCol)) 'Fecha Princ
                        xlHoja1.Range(xlHoja1.Cells(lilineas - 2, 57), xlHoja1.Cells(lilineas - 2, 57)).NumberFormat = "mmm-yy"
                        xlHoja1.Range(xlHoja1.Cells(lilineas - 2, 57), xlHoja1.Cells(lilineas - 2, 58)).Merge True
                        xlHoja1.Range(xlHoja1.Cells(lilineas - 2, 57), xlHoja1.Cells(lilineas - 2, 58)).HorizontalAlignment = xlCenter
                        
                        xlHoja1.Cells(lilineas - 1, 57) = "Sumatoria de TOSE por Moneda" 'Sumatoria TOSE
                        xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 57), xlHoja1.Cells(lilineas - 1, 58)).Merge True
                        xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 57), xlHoja1.Cells(lilineas - 1, 58)).HorizontalAlignment = xlCenter
                        ExcelCuadro xlHoja1, 57, lilineas - 1, 58, CCur(lilineas - 1)
                        xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 57), xlHoja1.Cells(lilineas - 1, 58)).Interior.ColorIndex = 44
                        xlHoja1.Cells(lilineas, 57) = "Moneda Nacional"
                        ExcelCuadro xlHoja1, 57, lilineas, 57, CCur(lilineas)
                        xlHoja1.Cells(lilineas + 1, 57).Formula = "=" & "+" & Mid(lsCadenaMN(1), 5, Len(lsCadenaMN(1)) - 5)
                        xlHoja1.Cells(lilineas + 2, 57).Formula = "=" & "+" & Mid(lsCadenaMN(2), 5, Len(lsCadenaMN(2)) - 5)
                        xlHoja1.Cells(lilineas + 3, 57).Formula = "=" & "+" & Mid(lsCadenaMN(3), 5, Len(lsCadenaMN(3)) - 5)
                        xlHoja1.Cells(lilineas + 4, 57).Formula = "=" & "+" & Mid(lsCadenaMN(4), 5, Len(lsCadenaMN(4)) - 5)
                        ExcelCuadro xlHoja1, 57, lilineas + 1, 57, CCur(lilineas + 4)
                        
                        xlHoja1.Cells(lilineas, 58) = "Moneda Extranjera"
                        ExcelCuadro xlHoja1, 58, lilineas, 58, CCur(lilineas)
                        xlHoja1.Cells(lilineas + 1, 58).Formula = "=" & "+" & Mid(lsCadenaME(1), 5, Len(lsCadenaME(1)) - 5)
                        xlHoja1.Cells(lilineas + 2, 58).Formula = "=" & "+" & Mid(lsCadenaME(2), 5, Len(lsCadenaME(2)) - 5)
                        xlHoja1.Cells(lilineas + 3, 58).Formula = "=" & "+" & Mid(lsCadenaME(3), 5, Len(lsCadenaME(3)) - 5)
                        xlHoja1.Cells(lilineas + 4, 58).Formula = "=" & "+" & Mid(lsCadenaME(4), 5, Len(lsCadenaME(4)) - 5)
                        ExcelCuadro xlHoja1, 58, lilineas + 1, 58, CCur(lilineas + 4)
                        
                        xlHoja1.Cells(lilineas - 1, 55) = "Variación del Mes" 'Variación MES
                        xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 55), xlHoja1.Cells(lilineas - 1, 56)).Merge True
                        xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 55), xlHoja1.Cells(lilineas - 1, 56)).HorizontalAlignment = xlCenter
                        xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 55), xlHoja1.Cells(lilineas - 1, 56)).Interior.Color = 5296274
                        ExcelCuadro xlHoja1, 55, lilineas - 1, 56, CCur(lilineas - 1)
                        xlHoja1.Cells(lilineas, 55) = "Moneda Nacional"
                        ExcelCuadro xlHoja1, 55, lilineas, 55, CCur(lilineas)
                        xlHoja1.Cells(lilineas + 1, 55).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, lilineasCol), xlHoja1.Cells(lilineas + 1, lilineasCol)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 7), xlHoja1.Cells(lilineas + 1, 7)).Address(False, False)
                        xlHoja1.Cells(lilineas + 2, 55).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, lilineasCol), xlHoja1.Cells(lilineas + 2, lilineasCol)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 7), xlHoja1.Cells(lilineas + 2, 7)).Address(False, False)
                        xlHoja1.Cells(lilineas + 3, 55).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, lilineasCol), xlHoja1.Cells(lilineas + 3, lilineasCol)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 7), xlHoja1.Cells(lilineas + 3, 7)).Address(False, False)
                        xlHoja1.Cells(lilineas + 4, 55).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, lilineasCol), xlHoja1.Cells(lilineas + 4, lilineasCol)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 7), xlHoja1.Cells(lilineas + 4, 7)).Address(False, False)
                        ExcelCuadro xlHoja1, 55, lilineas + 1, 55, CCur(lilineas + 4)
                        
                        xlHoja1.Cells(lilineas, 56) = "Moneda Extranjera"
                        ExcelCuadro xlHoja1, 56, lilineas, 56, CCur(lilineas)
                        xlHoja1.Cells(lilineas + 1, 56).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, lilineasCol + 1), xlHoja1.Cells(lilineas + 1, lilineasCol + 1)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 8), xlHoja1.Cells(lilineas + 1, 8)).Address(False, False)
                        xlHoja1.Cells(lilineas + 2, 56).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, lilineasCol + 1), xlHoja1.Cells(lilineas + 2, lilineasCol + 1)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 8), xlHoja1.Cells(lilineas + 2, 8)).Address(False, False)
                        xlHoja1.Cells(lilineas + 3, 56).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, lilineasCol + 1), xlHoja1.Cells(lilineas + 3, lilineasCol + 1)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 8), xlHoja1.Cells(lilineas + 3, 8)).Address(False, False)
                        xlHoja1.Cells(lilineas + 4, 56).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, lilineasCol + 1), xlHoja1.Cells(lilineas + 4, lilineasCol + 1)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 8), xlHoja1.Cells(lilineas + 4, 8)).Address(False, False)
                        ExcelCuadro xlHoja1, 56, lilineas + 1, 56, CCur(lilineas + 4)
                        
                        xlHoja1.Cells(lilineas - 1, 59) = "Promedio del Mes" 'Promedio del Mes
                        xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 59), xlHoja1.Cells(lilineas - 1, 60)).Merge True
                        xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 59), xlHoja1.Cells(lilineas - 1, 60)).HorizontalAlignment = xlCenter
                        xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 59), xlHoja1.Cells(lilineas - 1, 60)).Interior.Color = 16776960
                        ExcelCuadro xlHoja1, 59, lilineas - 1, 60, CCur(lilineas - 1)
                        xlHoja1.Cells(lilineas, 59) = "Moneda Nacional"
                        ExcelCuadro xlHoja1, 59, lilineas, 59, CCur(lilineas)
                        xlHoja1.Cells(lilineas + 1, 59).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 57), xlHoja1.Cells(lilineas + 1, 57)).Address(False, False) & "/" & CountReg
                        xlHoja1.Cells(lilineas + 2, 59).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 57), xlHoja1.Cells(lilineas + 2, 57)).Address(False, False) & "/" & CountReg
                        xlHoja1.Cells(lilineas + 3, 59).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 57), xlHoja1.Cells(lilineas + 3, 57)).Address(False, False) & "/" & CountReg
                        xlHoja1.Cells(lilineas + 4, 59).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 57), xlHoja1.Cells(lilineas + 4, 57)).Address(False, False) & "/" & CountReg
                        ExcelCuadro xlHoja1, 59, lilineas + 1, 59, CCur(lilineas + 4)
                        
                        xlHoja1.Cells(lilineas, 60) = "Moneda Extranjera"
                        ExcelCuadro xlHoja1, 60, lilineas, 60, CCur(lilineas)
                        xlHoja1.Cells(lilineas + 1, 60).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 58), xlHoja1.Cells(lilineas + 1, 58)).Address(False, False) & "/" & CountReg
                        xlHoja1.Cells(lilineas + 2, 60).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 58), xlHoja1.Cells(lilineas + 2, 58)).Address(False, False) & "/" & CountReg
                        xlHoja1.Cells(lilineas + 3, 60).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 58), xlHoja1.Cells(lilineas + 3, 58)).Address(False, False) & "/" & CountReg
                        xlHoja1.Cells(lilineas + 4, 60).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 58), xlHoja1.Cells(lilineas + 4, 58)).Address(False, False) & "/" & CountReg
                        ExcelCuadro xlHoja1, 60, lilineas + 1, 60, CCur(lilineas + 4)
                        
                        xlHoja1.Range(xlHoja1.Cells(lilineas, 61), xlHoja1.Cells(lilineas, 61)) = "III. Encaje"
                        xlHoja1.Range(xlHoja1.Cells(lilineas, 61), xlHoja1.Cells(lilineas, 64)).Merge True
                        xlHoja1.Range(xlHoja1.Cells(lilineas, 61), xlHoja1.Cells(lilineas, 61)).Font.Bold = True
                        ExcelCuadro xlHoja1, 61, lilineas, 64, CCur(lilineas)
                        xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 61), xlHoja1.Cells(lilineas + 1, 61)) = "1. Total de obligaciones sujetas a encaje - TOSE  (23)"
                        xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 61), xlHoja1.Cells(lilineas + 1, 64)).Merge True
                        xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 61), xlHoja1.Cells(lilineas + 2, 61)) = "1.1 Obligaciones inmediatas y a plazo hasta 30 días"
                        xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 61), xlHoja1.Cells(lilineas + 2, 64)).Merge True
                        xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 61), xlHoja1.Cells(lilineas + 3, 61)) = "1.2 Obligaciones a plazo mayor a 30 días"
                        xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 61), xlHoja1.Cells(lilineas + 3, 64)).Merge True
                        xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 61), xlHoja1.Cells(lilineas + 4, 61)) = "1.3 Ahorros"
                        xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 61), xlHoja1.Cells(lilineas + 4, 64)).Merge True
                        ExcelCuadro xlHoja1, 61, lilineas + 1, 64, CCur(lilineas + 4)
                        
                        xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 55), xlHoja1.Cells(lilineas + 4, 60)).NumberFormat = "#,###0.00"
                        xlHoja1.Range(xlHoja1.Cells(lilineas, 55), xlHoja1.Cells(lilineas, 60)).Interior.Color = RGB(153, 153, 255)
                        xlHoja1.Range(xlHoja1.Cells(lilineas - 2, 55), xlHoja1.Cells(lilineas, 60)).Font.Bold = True
                        xlHoja1.Range(xlHoja1.Cells(lilineas, 55), xlHoja1.Cells(lilineas, 60)).EntireColumn.AutoFit
                     End If
                    'Para el Inicio del Cuadro del Mes Correspondiente
                     lilineas = Cant + 3
                     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 1)) = "III. Encaje"
                     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 4)).Merge True
                     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 1)).Font.Bold = True
                     ExcelCuadro xlHoja1, 1, lilineas, 4, CCur(lilineas)
                     xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 1), xlHoja1.Cells(lilineas + 1, 1)) = "1. Total de obligaciones sujetas a encaje - TOSE  (23)"
                     xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 1), xlHoja1.Cells(lilineas + 1, 4)).Merge True
                     xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 1), xlHoja1.Cells(lilineas + 2, 1)) = "1.1 Obligaciones inmediatas y a plazo hasta 30 días"
                     xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 1), xlHoja1.Cells(lilineas + 2, 4)).Merge True
                     xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 1), xlHoja1.Cells(lilineas + 3, 1)) = "1.2 Obligaciones a plazo mayor a 30 días"
                     xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 1), xlHoja1.Cells(lilineas + 3, 4)).Merge True
                     xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 1), xlHoja1.Cells(lilineas + 4, 1)) = "1.3 Ahorros"
                     xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 1), xlHoja1.Cells(lilineas + 4, 4)).Merge True
                     ExcelCuadro xlHoja1, 1, lilineas + 1, 4, CCur(lilineas + 4)
                     nCorrelativo = nCorrelativo + 1
                     lilineasCol = 5
                     CountReg = 0
                     lsCadenaMN(1) = ""
                     lsCadenaMN(2) = ""
                     lsCadenaMN(3) = ""
                     lsCadenaMN(4) = ""
                     lsCadenaME(1) = ""
                     lsCadenaME(2) = ""
                     lsCadenaME(3) = ""
                     lsCadenaME(4) = ""
            End If
               
               xlHoja1.Cells(lilineas - 1, lilineasCol) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha
               xlHoja1.Range(xlHoja1.Cells(lilineas - 1, lilineasCol), xlHoja1.Cells(lilineas - 1, lilineasCol + 1)).Merge True
               xlHoja1.Range(xlHoja1.Cells(lilineas - 1, lilineasCol), xlHoja1.Cells(lilineas - 1, lilineasCol + 1)).HorizontalAlignment = xlCenter
               xlHoja1.Cells(lilineas, lilineasCol) = "Moneda Nacional"
               ExcelCuadro xlHoja1, lilineasCol, lilineas, lilineasCol, CCur(lilineas)
               xlHoja1.Cells(lilineas, lilineasCol + 1) = "Moneda Extranjera"
               ExcelCuadro xlHoja1, lilineasCol + 1, lilineas, lilineasCol + 1, CCur(lilineas)
               
               If lilineasCol <> 5 Then
                    xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol), xlHoja1.Cells(lilineas, lilineasCol + 1)).Interior.Color = RGB(153, 153, 255)
               Else
                    xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol), xlHoja1.Cells(lilineas, lilineasCol + 1)).Interior.ColorIndex = 44
               End If
               xlHoja1.Range(xlHoja1.Cells(lilineas, lilineasCol), xlHoja1.Cells(lilineas, lilineasCol + 1)).Font.Bold = True
               
               xlHoja1.Cells(lilineas + 1, lilineasCol) = Format(rs!TotalTOSEMN, "#,##0.00")
               lsCadenaMN(1) = lsCadenaMN(1) & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, lilineasCol), xlHoja1.Cells(lilineas + 1, lilineasCol)).Address(False, False) & "+"
               
               xlHoja1.Cells(lilineas + 1, lilineasCol + 1) = Format(rs!TotalTOSEME, "#,##0.00")
               lsCadenaME(1) = lsCadenaME(1) & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, lilineasCol + 1), xlHoja1.Cells(lilineas + 1, lilineasCol + 1)).Address(False, False) & "+"
               
               xlHoja1.Cells(lilineas + 2, lilineasCol) = Format(rs!ObligInmPlazhas30MN, "#,##0.00")
               lsCadenaMN(2) = lsCadenaMN(2) & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, lilineasCol), xlHoja1.Cells(lilineas + 2, lilineasCol)).Address(False, False) & "+"
               
               xlHoja1.Cells(lilineas + 2, lilineasCol + 1) = Format(rs!ObligInmPlazhas30ME, "#,##0.00")
               lsCadenaME(2) = lsCadenaME(2) & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, lilineasCol + 1), xlHoja1.Cells(lilineas + 2, lilineasCol + 1)).Address(False, False) & "+"
               
               xlHoja1.Cells(lilineas + 3, lilineasCol) = Format(rs!ObligPlazmay30MN, "#,##0.00")
               lsCadenaMN(3) = lsCadenaMN(3) & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, lilineasCol), xlHoja1.Cells(lilineas + 3, lilineasCol)).Address(False, False) & "+"
               
               xlHoja1.Cells(lilineas + 3, lilineasCol + 1) = Format(rs!ObligPlazmay30ME, "#,##0.00")
               lsCadenaME(3) = lsCadenaME(3) & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, lilineasCol + 1), xlHoja1.Cells(lilineas + 3, lilineasCol + 1)).Address(False, False) & "+"
               
               xlHoja1.Cells(lilineas + 4, lilineasCol) = Format(rs!AhorrosMN, "#,##0.00")
               lsCadenaMN(4) = lsCadenaMN(4) & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, lilineasCol), xlHoja1.Cells(lilineas + 4, lilineasCol)).Address(False, False) & "+"
               
               xlHoja1.Cells(lilineas + 4, lilineasCol + 1) = Format(rs!AhorrosME, "#,##0.00")
               lsCadenaME(4) = lsCadenaME(4) & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, lilineasCol + 1), xlHoja1.Cells(lilineas + 4, lilineasCol + 1)).Address(False, False) & "+"
               
               xlHoja1.Range(xlHoja1.Cells(lilineas + 1, lilineasCol), xlHoja1.Cells(lilineas + 4, lilineasCol)).EntireColumn.AutoFit
               xlHoja1.Range(xlHoja1.Cells(lilineas + 1, lilineasCol + 1), xlHoja1.Cells(lilineas + 4, lilineasCol + 1)).EntireColumn.AutoFit
               ExcelCuadro xlHoja1, lilineasCol, lilineas + 1, lilineasCol, CCur(lilineas + 4)
               ExcelCuadro xlHoja1, lilineasCol + 1, lilineas + 1, lilineasCol + 1, CCur(lilineas + 4)
               
               If lilineasCol >= 7 Then
                    xlHoja1.Cells(lilineas + 5, lilineasCol).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, lilineasCol), xlHoja1.Cells(lilineas + 1, lilineasCol)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, lilineasCol - 2), xlHoja1.Cells(lilineas + 1, lilineasCol - 2)).Address(False, False)
                    xlHoja1.Cells(lilineas + 5, lilineasCol).NumberFormat = "#,###0.00"
                    xlHoja1.Cells(lilineas + 6, lilineasCol).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, lilineasCol), xlHoja1.Cells(lilineas + 2, lilineasCol)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, lilineasCol - 2), xlHoja1.Cells(lilineas + 2, lilineasCol - 2)).Address(False, False)
                    xlHoja1.Cells(lilineas + 6, lilineasCol).NumberFormat = "#,###0.00"
                    xlHoja1.Cells(lilineas + 7, lilineasCol).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, lilineasCol), xlHoja1.Cells(lilineas + 3, lilineasCol)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, lilineasCol - 2), xlHoja1.Cells(lilineas + 3, lilineasCol - 2)).Address(False, False)
                    xlHoja1.Cells(lilineas + 7, lilineasCol).NumberFormat = "#,###0.00"
                    xlHoja1.Cells(lilineas + 8, lilineasCol).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, lilineasCol), xlHoja1.Cells(lilineas + 4, lilineasCol)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, lilineasCol - 2), xlHoja1.Cells(lilineas + 4, lilineasCol - 2)).Address(False, False)
                    xlHoja1.Cells(lilineas + 8, lilineasCol).NumberFormat = "#,###0.00"
                    'Diferencia Part2
                    xlHoja1.Cells(lilineas + 5, lilineasCol + 1).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, lilineasCol + 1), xlHoja1.Cells(lilineas + 1, lilineasCol + 1)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, lilineasCol - 1), xlHoja1.Cells(lilineas + 1, lilineasCol - 1)).Address(False, False)
                    xlHoja1.Cells(lilineas + 5, lilineasCol + 1).NumberFormat = "#,###0.00"
                    xlHoja1.Cells(lilineas + 6, lilineasCol + 1).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, lilineasCol + 1), xlHoja1.Cells(lilineas + 2, lilineasCol + 1)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, lilineasCol - 1), xlHoja1.Cells(lilineas + 2, lilineasCol - 1)).Address(False, False)
                    xlHoja1.Cells(lilineas + 6, lilineasCol + 1).NumberFormat = "#,###0.00"
                    xlHoja1.Cells(lilineas + 7, lilineasCol + 1).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, lilineasCol + 1), xlHoja1.Cells(lilineas + 3, lilineasCol + 1)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, lilineasCol - 1), xlHoja1.Cells(lilineas + 3, lilineasCol - 1)).Address(False, False)
                    xlHoja1.Cells(lilineas + 7, lilineasCol + 1).NumberFormat = "#,###0.00"
                    xlHoja1.Cells(lilineas + 8, lilineasCol + 1).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, lilineasCol + 1), xlHoja1.Cells(lilineas + 4, lilineasCol + 1)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, lilineasCol - 1), xlHoja1.Cells(lilineas + 4, lilineasCol - 1)).Address(False, False)
                    xlHoja1.Cells(lilineas + 8, lilineasCol + 1).NumberFormat = "#,###0.00"
                    
                    xlHoja1.Range(xlHoja1.Cells(lilineas + 5, lilineasCol), xlHoja1.Cells(lilineas + 8, lilineasCol + 1)).Font.Color = vbRed
                    
                    ExcelCuadro xlHoja1, lilineasCol, lilineas + 5, CCur(lilineasCol), lilineas + 8
                    ExcelCuadro xlHoja1, lilineasCol + 1, lilineas + 5, CCur(lilineasCol + 1), lilineas + 8
               End If
               lilineasCol = lilineasCol + 2
               Cant = lilineas + 8
               CountReg = CountReg + 1
            rs.MoveNext
         Loop
    End If
    
    '*******************************************PARA LA ULTIMA BANDA*********************************************'
    CountReg = CountReg - 1
    lilineasCol = lilineasCol - 2
    xlHoja1.Cells(lilineas - 2, 57) = xlHoja1.Range(xlHoja1.Cells(lilineas - 1, lilineasCol), xlHoja1.Cells(lilineas - 1, lilineasCol)) 'Fecha Princ
    xlHoja1.Range(xlHoja1.Cells(lilineas - 2, 57), xlHoja1.Cells(lilineas - 2, 57)).NumberFormat = "mmm-yy"
    xlHoja1.Range(xlHoja1.Cells(lilineas - 2, 57), xlHoja1.Cells(lilineas - 2, 58)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lilineas - 2, 57), xlHoja1.Cells(lilineas - 2, 58)).HorizontalAlignment = xlCenter
    
    xlHoja1.Cells(lilineas - 1, 57) = "Sumatoria de TOSE por Moneda" 'Sumatoria TOSE
    xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 57), xlHoja1.Cells(lilineas - 1, 58)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 57), xlHoja1.Cells(lilineas - 1, 58)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 57, lilineas - 1, 58, CCur(lilineas - 1)
    xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 57), xlHoja1.Cells(lilineas - 1, 58)).Interior.ColorIndex = 44
    xlHoja1.Cells(lilineas, 57) = "Moneda Nacional"
    ExcelCuadro xlHoja1, 57, lilineas, 57, CCur(lilineas)
    xlHoja1.Cells(lilineas + 1, 57).Formula = "=" & "+" & Mid(lsCadenaMN(1), 5, Len(lsCadenaMN(1)) - 5)
    xlHoja1.Cells(lilineas + 2, 57).Formula = "=" & "+" & Mid(lsCadenaMN(2), 5, Len(lsCadenaMN(2)) - 5)
    xlHoja1.Cells(lilineas + 3, 57).Formula = "=" & "+" & Mid(lsCadenaMN(3), 5, Len(lsCadenaMN(3)) - 5)
    xlHoja1.Cells(lilineas + 4, 57).Formula = "=" & "+" & Mid(lsCadenaMN(4), 5, Len(lsCadenaMN(4)) - 5)
    ExcelCuadro xlHoja1, 57, lilineas + 1, 57, CCur(lilineas + 4)
    
    xlHoja1.Cells(lilineas, 58) = "Moneda Extranjera"
    ExcelCuadro xlHoja1, 58, lilineas, 58, CCur(lilineas)
    xlHoja1.Cells(lilineas + 1, 58).Formula = "=" & "+" & Mid(lsCadenaME(1), 5, Len(lsCadenaME(1)) - 5)
    xlHoja1.Cells(lilineas + 2, 58).Formula = "=" & "+" & Mid(lsCadenaME(2), 5, Len(lsCadenaME(2)) - 5)
    xlHoja1.Cells(lilineas + 3, 58).Formula = "=" & "+" & Mid(lsCadenaME(3), 5, Len(lsCadenaME(3)) - 5)
    xlHoja1.Cells(lilineas + 4, 58).Formula = "=" & "+" & Mid(lsCadenaME(4), 5, Len(lsCadenaME(4)) - 5)
    ExcelCuadro xlHoja1, 58, lilineas + 1, 58, CCur(lilineas + 4)
    
    xlHoja1.Cells(lilineas - 1, 55) = "Variación del Mes" 'Variación MES
    xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 55), xlHoja1.Cells(lilineas - 1, 56)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 55), xlHoja1.Cells(lilineas - 1, 56)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 55), xlHoja1.Cells(lilineas - 1, 56)).Interior.Color = 5296274
    ExcelCuadro xlHoja1, 55, lilineas - 1, 56, CCur(lilineas - 1)
    xlHoja1.Cells(lilineas, 55) = "Moneda Nacional"
    ExcelCuadro xlHoja1, 55, lilineas, 55, CCur(lilineas)
    xlHoja1.Cells(lilineas + 1, 55).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, lilineasCol), xlHoja1.Cells(lilineas + 1, lilineasCol)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 7), xlHoja1.Cells(lilineas + 1, 7)).Address(False, False)
    xlHoja1.Cells(lilineas + 2, 55).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, lilineasCol), xlHoja1.Cells(lilineas + 2, lilineasCol)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 7), xlHoja1.Cells(lilineas + 2, 7)).Address(False, False)
    xlHoja1.Cells(lilineas + 3, 55).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, lilineasCol), xlHoja1.Cells(lilineas + 3, lilineasCol)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 7), xlHoja1.Cells(lilineas + 3, 7)).Address(False, False)
    xlHoja1.Cells(lilineas + 4, 55).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, lilineasCol), xlHoja1.Cells(lilineas + 4, lilineasCol)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 7), xlHoja1.Cells(lilineas + 4, 7)).Address(False, False)
    ExcelCuadro xlHoja1, 55, lilineas + 1, 55, CCur(lilineas + 4)
    
    xlHoja1.Cells(lilineas, 56) = "Moneda Extranjera"
    ExcelCuadro xlHoja1, 56, lilineas, 56, CCur(lilineas)
    xlHoja1.Cells(lilineas + 1, 56).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, lilineasCol + 1), xlHoja1.Cells(lilineas + 1, lilineasCol + 1)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 8), xlHoja1.Cells(lilineas + 1, 8)).Address(False, False)
    xlHoja1.Cells(lilineas + 2, 56).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, lilineasCol + 1), xlHoja1.Cells(lilineas + 2, lilineasCol + 1)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 8), xlHoja1.Cells(lilineas + 2, 8)).Address(False, False)
    xlHoja1.Cells(lilineas + 3, 56).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, lilineasCol + 1), xlHoja1.Cells(lilineas + 3, lilineasCol + 1)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 8), xlHoja1.Cells(lilineas + 3, 8)).Address(False, False)
    xlHoja1.Cells(lilineas + 4, 56).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, lilineasCol + 1), xlHoja1.Cells(lilineas + 4, lilineasCol + 1)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 8), xlHoja1.Cells(lilineas + 4, 8)).Address(False, False)
    ExcelCuadro xlHoja1, 56, lilineas + 1, 56, CCur(lilineas + 4)
    
    xlHoja1.Cells(lilineas - 1, 59) = "Promedio del Mes" 'Promedio del Mes
    xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 59), xlHoja1.Cells(lilineas - 1, 60)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 59), xlHoja1.Cells(lilineas - 1, 60)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(lilineas - 1, 59), xlHoja1.Cells(lilineas - 1, 60)).Interior.Color = 16776960
    ExcelCuadro xlHoja1, 59, lilineas - 1, 60, CCur(lilineas - 1)
    xlHoja1.Cells(lilineas, 59) = "Moneda Nacional"
    ExcelCuadro xlHoja1, 59, lilineas, 59, CCur(lilineas)
    xlHoja1.Cells(lilineas + 1, 59).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 57), xlHoja1.Cells(lilineas + 1, 57)).Address(False, False) & "/" & CountReg
    xlHoja1.Cells(lilineas + 2, 59).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 57), xlHoja1.Cells(lilineas + 2, 57)).Address(False, False) & "/" & CountReg
    xlHoja1.Cells(lilineas + 3, 59).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 57), xlHoja1.Cells(lilineas + 3, 57)).Address(False, False) & "/" & CountReg
    xlHoja1.Cells(lilineas + 4, 59).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 57), xlHoja1.Cells(lilineas + 4, 57)).Address(False, False) & "/" & CountReg
    ExcelCuadro xlHoja1, 59, lilineas + 1, 59, CCur(lilineas + 4)
    
    xlHoja1.Cells(lilineas, 60) = "Moneda Extranjera"
    ExcelCuadro xlHoja1, 60, lilineas, 60, CCur(lilineas)
    xlHoja1.Cells(lilineas + 1, 60).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 58), xlHoja1.Cells(lilineas + 1, 58)).Address(False, False) & "/" & CountReg
    xlHoja1.Cells(lilineas + 2, 60).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 58), xlHoja1.Cells(lilineas + 2, 58)).Address(False, False) & "/" & CountReg
    xlHoja1.Cells(lilineas + 3, 60).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 58), xlHoja1.Cells(lilineas + 3, 58)).Address(False, False) & "/" & CountReg
    xlHoja1.Cells(lilineas + 4, 60).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 58), xlHoja1.Cells(lilineas + 4, 58)).Address(False, False) & "/" & CountReg
    ExcelCuadro xlHoja1, 60, lilineas + 1, 60, CCur(lilineas + 4)
    
    xlHoja1.Range(xlHoja1.Cells(lilineas, 61), xlHoja1.Cells(lilineas, 61)) = "III. Encaje"
    xlHoja1.Range(xlHoja1.Cells(lilineas, 61), xlHoja1.Cells(lilineas, 64)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lilineas, 61), xlHoja1.Cells(lilineas, 61)).Font.Bold = True
    ExcelCuadro xlHoja1, 61, lilineas, 64, CCur(lilineas)
    xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 61), xlHoja1.Cells(lilineas + 1, 61)) = "1. Total de obligaciones sujetas a encaje - TOSE  (23)"
    xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 61), xlHoja1.Cells(lilineas + 1, 64)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 61), xlHoja1.Cells(lilineas + 2, 61)) = "1.1 Obligaciones inmediatas y a plazo hasta 30 días"
    xlHoja1.Range(xlHoja1.Cells(lilineas + 2, 61), xlHoja1.Cells(lilineas + 2, 64)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 61), xlHoja1.Cells(lilineas + 3, 61)) = "1.2 Obligaciones a plazo mayor a 30 días"
    xlHoja1.Range(xlHoja1.Cells(lilineas + 3, 61), xlHoja1.Cells(lilineas + 3, 64)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 61), xlHoja1.Cells(lilineas + 4, 61)) = "1.3 Ahorros"
    xlHoja1.Range(xlHoja1.Cells(lilineas + 4, 61), xlHoja1.Cells(lilineas + 4, 64)).Merge True
    ExcelCuadro xlHoja1, 61, lilineas + 1, 64, CCur(lilineas + 4)
    
    xlHoja1.Range(xlHoja1.Cells(lilineas + 1, 55), xlHoja1.Cells(lilineas + 4, 60)).NumberFormat = "#,###0.00"
    xlHoja1.Range(xlHoja1.Cells(lilineas, 55), xlHoja1.Cells(lilineas, 60)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(lilineas - 2, 55), xlHoja1.Cells(lilineas, 60)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lilineas, 55), xlHoja1.Cells(lilineas, 60)).EntireColumn.AutoFit
    '********************************************************FIN***************************************************
    Set rs = Nothing
'SIGUIENTE HOJA DE CÁLCULO - TOSE
    lsNomHoja = "TOSE_HIST"
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
    
    lilineas = 3
    oBarra.Progress 80, TituloProgress, MensajeProgress, "", vbBlue
    Set rs = DAnxRies.ObtieneReporteDatosHistTOSE(pdFecha)
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            xlHoja1.Cells(lilineas, 3) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha
            xlHoja1.Cells(lilineas, 4) = Format(rs!TOSE_MNIng, "#,##0.00") 'TOSEMN
            xlHoja1.Cells(lilineas, 5) = Format(rs!TOSE_MEIng, "#,##0.00") 'TOSEME
            xlHoja1.Cells(lilineas, 6) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha
            xlHoja1.Cells(lilineas, 7).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas, 4)).Address(False, False) & "/" & "1000000"
            xlHoja1.Cells(lilineas, 7).NumberFormat = "#,##0"
            xlHoja1.Cells(lilineas, 8) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha
            xlHoja1.Cells(lilineas, 9).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(lilineas, 5), xlHoja1.Cells(lilineas, 5)).Address(False, False) & "/" & "1000000"
            xlHoja1.Cells(lilineas, 9).NumberFormat = "#,##0"
            xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 9)).HorizontalAlignment = xlCenter
            ExcelCuadro xlHoja1, 3, lilineas, 9, CCur(lilineas)
            lilineas = lilineas + 1
            rs.MoveNext
         Loop
    End If
    
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

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_LostFocus()
If ValFecha(txtFecha) Then
    txtTC = oDbalanceCont.ObtenerTipoCambioCierreNew(txtFecha, "TipoAct")
    TipoCamb = txtTC
Else
    TipoCamb = txtTC
End If
End Sub

Private Sub txtFecha_KeyPress(keyAscii As Integer)
    If keyAscii = 13 Then
       If ValFecha(txtFecha) Then
            txtTC = oDbalanceCont.ObtenerTipoCambioCierreNew(txtFecha, "TipoAct")
            TipoCamb = txtTC
            txtTC.SetFocus
       End If
    End If
End Sub

Private Sub txtTC_GotFocus()
fEnfoque txtTC
End Sub

Private Sub txtTC_KeyPress(keyAscii As Integer)
    If keyAscii = 13 Then
       cmdProcesar.SetFocus
    End If
End Sub
