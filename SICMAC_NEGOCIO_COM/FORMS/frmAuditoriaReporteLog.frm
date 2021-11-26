VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAuditoriaReporteLog 
   Caption         =   "Reporte de Log"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8295
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   6840
         TabIndex        =   11
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   960
         TabIndex        =   8
         Top             =   270
         Width           =   930
      End
      Begin MSComCtl2.DTPicker Hasta 
         Height          =   345
         Left            =   3120
         TabIndex        =   7
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
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
         Format          =   64094209
         CurrentDate     =   40233
      End
      Begin MSComCtl2.DTPicker Desde 
         Height          =   345
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
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
         Format          =   64094209
         CurrentDate     =   40233
      End
      Begin VB.TextBox txtUsuario 
         Height          =   350
         Left            =   960
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1440
         Width           =   975
      End
      Begin SICMACT.TxtBuscar TxtAgencia 
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   855
         _ExtentX        =   1296
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin VB.OLE OleExcel 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         Enabled         =   0   'False
         Height          =   240
         Left            =   0
         SizeMode        =   1  'Stretch
         TabIndex        =   12
         Top             =   2160
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta:"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAuditoriaReporteLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oInventario As COMNAuditoria.NCOMAF
Dim oArea As DActualizaDatosArea
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Private Sub cmdAceptar_Click()
    Dim rsDatos As ADODB.Recordset
    Set rsDatos = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    Set oInventario = New COMNAuditoria.NCOMAF
    
    Set rsDatos = oInventario.ObtenerReporteLog(IIf(chkTodos.value = 1, "", TxtAgencia.Text), Desde.value, Hasta.value, Trim(txtUsuario))
    
        If rsDatos Is Nothing Then
            MsgBox "No existen datos.", vbInformation, "Aviso"
            Exit Sub
        End If
    
    lsArchivoN = App.path & "\Spooler\" & "ReporteLog" & Format(CDate(gdFecSis), "yyyymmdd") & Format(Time, "hhmmss") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       ReporteLogCabeceraExcel xlHoja1
       GeneraReporte rsDatos
       OleExcel.Class = "ExcelWorkSheet"
       gFunGeneral.ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
    MousePointer = 0
    
End Sub

Public Function ReporteLogCabeceraExcel(Optional xlHoja1 As Excel.Worksheet) As String
    xlHoja1.PageSetup.LeftMargin = 1.5
    xlHoja1.PageSetup.RightMargin = 0
    xlHoja1.PageSetup.BottomMargin = 1
    xlHoja1.PageSetup.TopMargin = 1
    xlHoja1.PageSetup.Zoom = 70
    
    xlHoja1.Cells(2, 3) = "REPORTE DE LOG" & " " & Desde.value & " " & Hasta.value
    
    
    xlHoja1.Cells(9, 1) = "FECHA"
    
    xlHoja1.Cells(9, 2) = "AGENCIA"
    xlHoja1.Cells(9, 3) = "USUARIO"
    xlHoja1.Cells(9, 4) = "COD CTA"
    xlHoja1.Cells(9, 5) = "DESCRIPCION"
    xlHoja1.Cells(9, 6) = "MAQUINA"
    xlHoja1.Cells(9, 7) = "COMENTARIO"
    xlHoja1.Cells(9, 8) = "OPERACION"
    
       
    xlHoja1.Range("A3:G6").Font.Bold = True
    xlHoja1.Range("A8:H9").Font.Bold = True
    
    xlHoja1.Range("A4:B6").Font.Size = 9
    xlHoja1.Range("A3:G3").Font.Size = 9
    xlHoja1.Range("A8:H9").Font.Size = 7
    
    xlHoja1.Range("A8:H9").BorderAround xlContinuous, xlMedium
    xlHoja1.Range("A8:H9").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range("A8:H9").Borders(xlInsideVertical).LineStyle = xlContinuous
    
    'Border y Merge del Titulo
    xlHoja1.Range("C2:F2").MergeCells = True
    xlHoja1.Range("C2:F2").Font.Bold = True
     
    xlHoja1.Range("A4:A4").ColumnWidth = 15
    xlHoja1.Range("B4:B4").ColumnWidth = 15
    xlHoja1.Range("C4:C4").ColumnWidth = 15
    xlHoja1.Range("D4:D4").ColumnWidth = 12
    xlHoja1.Range("E4:E4").ColumnWidth = 12
    xlHoja1.Range("F4:F4").ColumnWidth = 12
    xlHoja1.Range("G4:G4").ColumnWidth = 12
    
    xlHoja1.Range("H4:H4").ColumnWidth = 14
    xlHoja1.Range("I4:I4").ColumnWidth = 11
    xlHoja1.Range("J4:J4").ColumnWidth = 11
    xlHoja1.Range("K4:K4").ColumnWidth = 15
    xlHoja1.Range("L4:L4").ColumnWidth = 15
    
    xlHoja1.Range("M4:M4").ColumnWidth = 9
    xlHoja1.Range("N4:N4").ColumnWidth = 15
    xlHoja1.Range("O4:O4").ColumnWidth = 13
    xlHoja1.Range("P4:P4").ColumnWidth = 15
    
    xlHoja1.Range("Q4:Q4").ColumnWidth = 10
    xlHoja1.Range("R4:R4").ColumnWidth = 15
    xlHoja1.Range("S4:S4").ColumnWidth = 13
    
    xlHoja1.Range("T4:T4").ColumnWidth = 23
    xlHoja1.Range("U4:U4").ColumnWidth = 13
    
    xlHoja1.Range("V4:V4").ColumnWidth = 25
    xlHoja1.Range("W4:W4").ColumnWidth = 16
    xlHoja1.Range("X4:X4").ColumnWidth = 16
    xlHoja1.Range("Y4:Y4").ColumnWidth = 16
    xlHoja1.Range("Z4:Z4").ColumnWidth = 20
    
    xlHoja1.Range("B4:B4").HorizontalAlignment = xlLeft
    xlHoja1.Range("B5:B5").HorizontalAlignment = xlLeft
    xlHoja1.Range("B6:B6").HorizontalAlignment = xlLeft
    xlHoja1.Range("A8:H9").HorizontalAlignment = xlCenter
    xlHoja1.Range("A8:H9").VerticalAlignment = xlCenter
    
End Function

Private Sub GeneraReporte(prRs As ADODB.Recordset)
    Dim i As Integer
    Dim J As Integer
    
    i = 8
    While Not prRs.EOF
        i = i + 1
        For J = 0 To prRs.Fields.Count - 1
            xlHoja1.Cells(i + 1, J + 1) = prRs.Fields(J)
        Next J
        prRs.MoveNext
    Wend
End Sub

Private Sub Form_Load()
    Set oArea = New DActualizaDatosArea
    Me.TxtAgencia.rs = oArea.getAgencias
End Sub

Private Sub chktodos_Click()
    If Me.chkTodos.value = 1 Then
        Me.TxtAgencia.Text = ""
        Me.lblAgencia.Caption = ""
    End If
End Sub

Private Sub TxtAgencia_EmiteDatos()
    lblAgencia.Caption = TxtAgencia.psDescripcion
    chkTodos.value = 0
End Sub
