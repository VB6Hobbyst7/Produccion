VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRHReportesSubsidio 
   Caption         =   "Reporte de Subsidio"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5745
   Icon            =   "frmRHReportesSubsidio.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraFechas 
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
      ForeColor       =   &H00800000&
      Height          =   1320
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5520
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   4320
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskFF 
         Height          =   315
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFI 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFI 
         Caption         =   "Inicio :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lblFF 
         Caption         =   "Fin :"
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmRHReportesSubsidio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
Dim RHPer As DPeriodoNoLaborado
Dim rs As New ADODB.Recordset
Dim fs As Scripting.FileSystemObject
Dim lsArchivo As String

Dim lsNomHoja As String
Dim i, Y As Integer
Dim lbExisteHoja As Boolean

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Set RHPer = New DPeriodoNoLaborado
Set rs = RHPer.CargarSubsidioConsol(Format(mskFI.Text, "yyyymmdd"), Format(mskFF.Text, "yyyymmdd"))

Screen.MousePointer = 11
lsArchivo = "ReporteSubsidio" & Format(Now, "yyyymm") & "_" & Format(Time(), "HHMMSS") & ".xls"

Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\Spooler\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\Spooler\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

lsNomHoja = Format(gdFecSis, "YYYYMM")
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = lsNomHoja Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
    End If
Next
If lbExisteHoja = False Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = lsNomHoja
End If

xlHoja1.Range("B1") = "CAJA MAYNAS"
xlHoja1.Range("B1:C1").MergeCells = True
xlHoja1.Range("B1").Font.Bold = True
xlHoja1.Range("G1") = gdFecSis
xlHoja1.Range("G2") = gsCodUser
xlHoja1.Range("G2").HorizontalAlignment = xlRight
xlHoja1.Range("H1:H2").Font.Bold = True

xlHoja1.Range("B5") = "Nombre"
xlHoja1.Range("C5") = "Agencia"
xlHoja1.Range("D5") = "Motivo"
xlHoja1.Range("E5") = "Fecha Inicio"
xlHoja1.Range("F5") = "Fecha Fin"
xlHoja1.Range("G5") = "Dias D/Medico"

xlHoja1.Range("B4:G4").MergeCells = True
xlHoja1.Range("B4") = "REPORTE DE SUBSIDIO" & " DEL " & Format(gdFecSis, "DD/MM/YYYY")
xlHoja1.Range("B4").Font.Bold = True
xlHoja1.Range("B4").HorizontalAlignment = xlCenter

xlHoja1.Range("B5:G5").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:G5").Interior.ColorIndex = 35
xlHoja1.Range("B5:G5").Font.Bold = True

xlHoja1.Range("B1").ColumnWidth = 40
xlHoja1.Range("C1").ColumnWidth = 20
xlHoja1.Range("D1").ColumnWidth = 25
xlHoja1.Range("E1").ColumnWidth = 15
xlHoja1.Range("F1").ColumnWidth = 15
xlHoja1.Range("G1").ColumnWidth = 15

xlHoja1.Application.ActiveWindow.Zoom = 80
Y = 6

For i = 1 To rs.RecordCount
   
    xlHoja1.Range("B" & Y) = rs!cPersNombre
    xlHoja1.Range("C" & Y) = rs!cAgeDescripcion
    
    xlHoja1.Range("D" & Y) = rs!Motivo
    xlHoja1.Range("E" & Y) = rs!SolIni
    xlHoja1.Range("F" & Y) = rs!SolFin
    xlHoja1.Range("G" & Y) = DateDiff("D", rs!SolIni, rs!SolFin)
    
    rs.MoveNext
    Y = Y + 1
Next i

xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.

Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Screen.MousePointer = 0
MsgBox "Se ha Generado el Archivo " & lsArchivo & " Satisfactoriamente en la carpeta Spooler de SICMACT ADM", vbInformation, "Aviso"

CargaArchivo lsArchivo, App.path & "\SPOOLER\"
Exit Sub
End Sub

Private Sub Form_Load()
    CentraForm Me
    mskFI.Text = "01" & Mid(gdFecSis, 3, 8)
    mskFF.Text = gdFecSis
End Sub

Private Sub mskFF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    End If
End Sub

Private Sub mskFI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskFF.SetFocus
    End If
End Sub
