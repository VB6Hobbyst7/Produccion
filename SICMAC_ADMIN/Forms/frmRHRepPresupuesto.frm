VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRHRepPresupuesto 
   Caption         =   "Presupuesto Trimestral"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11640
   Icon            =   "frmRHRepPresupuesto.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   11640
   Begin VB.Frame fmeses 
      Caption         =   "Meses"
      Height          =   735
      Left            =   5880
      TabIndex        =   12
      Top             =   600
      Width           =   2295
      Begin VB.ComboBox cmbmes 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fano 
      Caption         =   "Año"
      Height          =   735
      Left            =   840
      TabIndex        =   10
      Top             =   600
      Width           =   1575
      Begin VB.ComboBox cmbano 
         Height          =   315
         ItemData        =   "frmRHRepPresupuesto.frx":030A
         Left            =   120
         List            =   "frmRHRepPresupuesto.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   9720
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.Frame fInterinac 
      Height          =   735
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   3255
      Begin VB.ComboBox cmbgrupo 
         Height          =   315
         ItemData        =   "frmRHRepPresupuesto.frx":032D
         Left            =   720
         List            =   "frmRHRepPresupuesto.frx":0337
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.ComboBox cmbReportes 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   150
      Width           =   7095
   End
   Begin VB.CommandButton cmdexporta 
      Caption         =   "Exporta >>>"
      Height          =   375
      Left            =   9000
      TabIndex        =   1
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10320
      TabIndex        =   0
      Top             =   6240
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msgreporte 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8281
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Reporte"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   570
   End
   Begin VB.OLE OleExcel 
      Class           =   "Excel.Sheet.8"
      Height          =   255
      Left            =   3930
      OleObjectBlob   =   "frmRHRepPresupuesto.frx":0447
      TabIndex        =   3
      Top             =   6210
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmRHRepPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oPla As DActualizaDatosConPlanilla
Dim oReporte As DRHReportes
Dim rs As ADODB.Recordset



Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
 
  

Private Sub cmbReportes_Click()

Select Case Val(Right(cmbReportes, 2))

Case 50
     fInterinac.Visible = False
     fano.Visible = True
     fmeses.Visible = True
     fmeses.Left = 2520
Case 51
     fInterinac.Visible = False
     fano.Visible = False
     fmeses.Visible = False
     

Case 52
     fInterinac.Visible = True
     fano.Visible = True
     fmeses.Visible = False

End Select
End Sub

Private Sub cmdexporta_Click()
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    If Me.msgreporte.TextMatrix(1, 1) = "" Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.cmbano.Text), "yyyy") & Format(Time, "hhmmss") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       GeneraReporte msgreporte, xlHoja1
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
    MousePointer = 0
End Sub

Private Sub cmdProcesar_Click()
Dim ano As String
Dim mes As String


If cmbano.Text = "" Then Exit Sub
If cmbmes.Text = "" Then Exit Sub



Select Case Val(Right(cmbReportes, 2))

Case 50
        
        Set rs = oPla.GetRHListaMontoCargo(Trim(cmbano.Text) + Format(Right(cmbmes.Text, 2), "00"))
        Set msgreporte.Recordset = rs

Case 51


Case 52
           If cmbgrupo.Text = "" Then Exit Sub
            If cmbano.Text = "" Then Exit Sub
            Set rs = oPla.GetRHPresupuesto(Right(cmbgrupo.Text, 2), cmbano.Text)
            Set msgreporte.Recordset = rs
            msgreporte.MergeCol(1) = True
            
            

End Select


End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Width = 11745
Me.Height = 7170

Me.cmbano.ListIndex = 0
Me.cmbgrupo.ListIndex = 0
Set rs = New ADODB.Recordset
Set oPla = New DActualizaDatosConPlanilla
Set oReporte = New DRHReportes


Set oReporte = New DRHReportes
Set rs = oReporte.GetRHReportesPre
CargaCombo rs, cmbReportes
cmbReportes.ListIndex = 0




msgreporte.Cols = 16
msgreporte.MergeCells = flexMergeRestrictColumns
msgreporte.ColWidth(0) = 1000
msgreporte.ColWidth(1) = 2500
msgreporte.ColWidth(2) = 2500

msgreporte.MergeCol(1) = True



cmbmes.AddItem "ENERO" & Space(20) & "01"
cmbmes.AddItem "FEBRERO" & Space(20) & "02"
cmbmes.AddItem "MARZO" & Space(20) & "03"
cmbmes.AddItem "ABRIL" & Space(20) & "04"
cmbmes.AddItem "MAYO" & Space(20) & "05"
cmbmes.AddItem "JUNIO" & Space(20) & "06"
cmbmes.AddItem "JULIO" & Space(20) & "07"
cmbmes.AddItem "AGOSTO" & Space(20) & "08"
cmbmes.AddItem "SETIEMBRE" & Space(20) & "09"
cmbmes.AddItem "OCTUBRE" & Space(20) & "10"
cmbmes.AddItem "NOVIEMBRE" & Space(20) & "11"
cmbmes.AddItem "DICIEMBRE" & Space(20) & "12"

cmbmes.ListIndex = 0




End Sub

