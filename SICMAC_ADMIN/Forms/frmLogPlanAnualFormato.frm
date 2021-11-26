VERSION 5.00
Begin VB.Form frmLogPlanAnualFormato 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3000
   ClientLeft      =   1605
   ClientTop       =   3195
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   8130
   Begin VB.Frame Frame1 
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7875
      Begin VB.CommandButton cmdConsucode 
         Caption         =   "Generar Plan Anual en formato CONSUCODE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1860
         Width           =   4635
      End
      Begin VB.TextBox txtAnio 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6840
         TabIndex        =   1
         Text            =   "2005"
         Top             =   105
         Width           =   615
      End
      Begin VB.Label lblPlan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plan Anual de Adquisiciones y Contrataciones del Año  "
         BeginProperty Font 
            Name            =   "Helvetica"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   420
         TabIndex        =   2
         Top             =   120
         Width           =   6375
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00EAFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   525
         Left            =   240
         Top             =   0
         Width           =   7395
      End
   End
End
Attribute VB_Name = "frmLogPlanAnualFormato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConsucode_Click()
Dim appExcel As New Excel.Application
Dim wbExcel As Excel.Workbook
Dim nFil As Integer, nCol As Integer, cCelda As String
Dim i As Integer, j As Integer, k As Integer, n As Integer

Set wbExcel = appExcel.Workbooks.Add
wbExcel.Worksheets(1).Range("A1:EZ200").Font.Name = ""
wbExcel.Worksheets(1).Range("A1:EZ200").Font.Size = 7
wbExcel.Worksheets(1).Range("A1:T11").Font.Bold = True
wbExcel.Worksheets(1).Range("A10:T11").WrapText = 1
'wbExcel.Worksheets(1).Range("A1:" + cUltLetra + "5").BorderAround 1, 1
'wbExcel.Worksheets(1).Range("C1:C" + CStr(nNroFilas)).Value = ""

wbExcel.Worksheets(1).Range("A1").ColumnWidth = 3.86
wbExcel.Worksheets(1).Range("B1").ColumnWidth = 6.14
wbExcel.Worksheets(1).Range("C1").ColumnWidth = 19.43
wbExcel.Worksheets(1).Range("D1").ColumnWidth = 13
wbExcel.Worksheets(1).Range("E1").ColumnWidth = 38.43
wbExcel.Worksheets(1).Range("F1").ColumnWidth = 5.14
wbExcel.Worksheets(1).Range("G1").ColumnWidth = 9.57
wbExcel.Worksheets(1).Range("H1").ColumnWidth = 8.71
wbExcel.Worksheets(1).Range("I1").ColumnWidth = 8.71
wbExcel.Worksheets(1).Range("J1").ColumnWidth = 9.14
wbExcel.Worksheets(1).Range("K1").ColumnWidth = 10.14
wbExcel.Worksheets(1).Range("L1").ColumnWidth = 9
wbExcel.Worksheets(1).Range("M1").ColumnWidth = 13.14
wbExcel.Worksheets(1).Range("N1").ColumnWidth = 12.14
wbExcel.Worksheets(1).Range("O1").ColumnWidth = 29.43
wbExcel.Worksheets(1).Range("P1").ColumnWidth = 17.71
wbExcel.Worksheets(1).Range("Q1").ColumnWidth = 7.57
wbExcel.Worksheets(1).Range("R1").ColumnWidth = 6.86
wbExcel.Worksheets(1).Range("S1").ColumnWidth = 8.86
wbExcel.Worksheets(1).Range("T1").ColumnWidth = 34.57

wbExcel.Worksheets(1).Range("B4").Font.Color = RGB(0, 0, 255)
wbExcel.Worksheets(1).Range("B4").Value = "B) NOMBRE DE LA ENTIDAD :"
wbExcel.Worksheets(1).Range("B6").Value = "C) SIGLAS :"
wbExcel.Worksheets(1).Range("B8").Value = "E) PLIEGO :"

wbExcel.Worksheets(1).Range("H6").Value = "F) UNIDAD EJECUTORA :"
wbExcel.Worksheets(1).Range("H8").Value = "G) INSTRUMENTO QUE APRUEBA O MODIFICA EL PAAC :"

wbExcel.Worksheets(1).Range("R4").Value = "A) AÑO :"
wbExcel.Worksheets(1).Range("R6").Value = "D) RUC :"

wbExcel.Worksheets(1).Range("A3:T3").MergeCells = True
wbExcel.Worksheets(1).Range("D4:O4").MergeCells = True:  wbExcel.Worksheets(1).Range("D4:O4").BorderAround 1, 2
wbExcel.Worksheets(1).Range("D6:E6").MergeCells = True:  wbExcel.Worksheets(1).Range("D6:E6").BorderAround 1, 2
wbExcel.Worksheets(1).Range("D8:E8").MergeCells = True:  wbExcel.Worksheets(1).Range("D8:E8").BorderAround 1, 2
wbExcel.Worksheets(1).Range("L6:O6").MergeCells = True:  wbExcel.Worksheets(1).Range("L6:O6").BorderAround 1, 2
wbExcel.Worksheets(1).Range("L8:T8").MergeCells = True:  wbExcel.Worksheets(1).Range("L8:T8").BorderAround 1, 2
wbExcel.Worksheets(1).Range("T4:T4").MergeCells = True:  wbExcel.Worksheets(1).Range("T4:T4").BorderAround 1, 2
wbExcel.Worksheets(1).Range("T6:T6").MergeCells = True:  wbExcel.Worksheets(1).Range("T6:T6").BorderAround 1, 2


wbExcel.Worksheets(1).Range("A10").Value = "N. REF"
wbExcel.Worksheets(1).Range("B10").Value = "PRECEDENTE"
wbExcel.Worksheets(1).Range("C10").Value = "TIPO DE PROCESO"
wbExcel.Worksheets(1).Range("D10").Value = "OBJETO"
wbExcel.Worksheets(1).Range("E10").Value = "SÍNTESIS DE ESPECIFICACIONES TÉCNICAS"
wbExcel.Worksheets(1).Range("F10").Value = "CIIU"
wbExcel.Worksheets(1).Range("G10").Value = "CATÁLOGO DE BIENES Y SERVICIOS"
wbExcel.Worksheets(1).Range("H10").Value = "VALOR ESTIMADO"
wbExcel.Worksheets(1).Range("I10").Value = "TIPO DE MONEDA"
wbExcel.Worksheets(1).Range("J10").Value = "UNIDAD DE MEDIDA"
wbExcel.Worksheets(1).Range("K10").Value = "CANTIDAD"
wbExcel.Worksheets(1).Range("L10").Value = "FUENTE DE FINANCIAMIENTO"
wbExcel.Worksheets(1).Range("M10").Value = "FECHA PROBABLE DE CONVOCATORIA"
wbExcel.Worksheets(1).Range("N10").Value = "COMPRA CORPORATIVA O POR ENCARGO"
wbExcel.Worksheets(1).Range("O10").Value = "NOMBRE DE LA ENTIDAD CONVOCANTE DE LA COMPRA CORPORATIVA O ENCARGADA"

wbExcel.Worksheets(1).Range("P10").Value = "ÓRGANO ENCARGADO DE LA ADQUISICIÓN O CONTRATACIÓN"
wbExcel.Worksheets(1).Range("Q10").Value = "CÓDIGO DE UBICACIÓN GEOGRÁFICA"
wbExcel.Worksheets(1).Range("Q11").Value = "DEPA"
wbExcel.Worksheets(1).Range("R11").Value = "PROV"
wbExcel.Worksheets(1).Range("S11").Value = "DIST"
wbExcel.Worksheets(1).Range("T10").Value = "OBSERVACIONES"

wbExcel.Worksheets(1).Range("A1:A8").RowHeight = 11.25
wbExcel.Worksheets(1).Range("A9").RowHeight = 12
wbExcel.Worksheets(1).Range("A10").RowHeight = 28.5
wbExcel.Worksheets(1).Range("A11").RowHeight = 13.5

For i = 1 To 16
    cCelda = ExcelColumnaString(i)
    wbExcel.Worksheets(1).Range(cCelda + "10:" + cCelda + "11").MergeCells = True
    wbExcel.Worksheets(1).Range(cCelda + "10:" + cCelda + "11").BorderAround 1, 3
    If i = 1 Or i = 5 Or i = 8 Or i = 11 Or i = 16 Or i = 20 Then
    Else
       wbExcel.Worksheets(1).Range(cCelda + "10").Font.Color = RGB(0, 0, 255)
    End If
Next
wbExcel.Worksheets(1).Range("Q10:S10").MergeCells = True
wbExcel.Worksheets(1).Range("T10:T11").MergeCells = True
wbExcel.Worksheets(1).Range("Q10:S10").BorderAround 1, 3
wbExcel.Worksheets(1).Range("Q11:S11").BorderAround 1, 3
wbExcel.Worksheets(1).Range("T10:T11").BorderAround 1, 3
wbExcel.Worksheets(1).Range("Q10:S11").Font.Color = RGB(0, 0, 255)
wbExcel.Worksheets(1).Range("A10:T11").HorizontalAlignment = 7
wbExcel.Worksheets(1).Range("A10:T11").VerticalAlignment = 2
'Para visualización
appExcel.Application.Visible = True
appExcel.Windows(1).Visible = True
End Sub


Private Sub Form_Load()
CentraForm Me
End Sub
