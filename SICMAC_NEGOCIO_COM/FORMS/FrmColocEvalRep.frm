VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmColocEvalRep 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de Evaluacion de Colocaciones"
   ClientHeight    =   7230
   ClientLeft      =   5010
   ClientTop       =   4095
   ClientWidth     =   8610
   Icon            =   "FrmColocEvalRep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      Height          =   2655
      Left            =   5280
      TabIndex        =   16
      Top             =   720
      Width           =   3255
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1920
         TabIndex        =   21
         Text            =   "8000"
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox CboSigno 
         Height          =   315
         ItemData        =   "FrmColocEvalRep.frx":030A
         Left            =   1320
         List            =   "FrmColocEvalRep.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   720
         Width           =   615
      End
      Begin VB.Frame FraMoneda2 
         Caption         =   "Moneda"
         Height          =   855
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
         Begin VB.OptionButton OptMoneda 
            Caption         =   "Soles"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptMoneda 
            Caption         =   "Dolares"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame FraEstado 
         Caption         =   "Estado"
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   3015
         Begin VB.OptionButton OptEstado 
            Caption         =   "Refinanciado"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   24
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton OptEstado 
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame FraTipo 
         Caption         =   "Tipo"
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   3015
         Begin VB.OptionButton OptTipo 
            Caption         =   "Mes"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   28
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Comercial"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   22
         Top             =   480
         Width           =   660
      End
   End
   Begin VB.Frame frmPatrimonio 
      Caption         =   "Patrimonio"
      Height          =   2175
      Left            =   5280
      TabIndex        =   51
      Top             =   4440
      Width           =   3255
      Begin VB.TextBox txtVPRiesgoAnterior 
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   1070
         Width           =   3015
      End
      Begin VB.TextBox txtVPRiesgo 
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   450
         Width           =   3015
      End
      Begin VB.Label Label8 
         Caption         =   "Patrimonio efectivo del mes anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   890
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "Valor patrimonial en riesgo"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame FraCF 
      Height          =   500
      Left            =   5280
      TabIndex        =   41
      Top             =   3360
      Width           =   3255
      Begin VB.CheckBox ChkCF 
         Caption         =   "Carta Fianza"
         Height          =   195
         Left            =   1000
         TabIndex        =   42
         Top             =   200
         Width           =   2000
      End
   End
   Begin VB.Frame FraIntervalos 
      Caption         =   "Intervalos de Montos"
      Height          =   1575
      Left            =   8640
      TabIndex        =   36
      Top             =   4800
      Width           =   3255
      Begin VB.TextBox TxtHasta 
         Height          =   285
         Left            =   1320
         TabIndex        =   40
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox TxtDesde 
         Height          =   285
         Left            =   1320
         TabIndex        =   39
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   480
         TabIndex        =   38
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   480
         TabIndex        =   37
         Top             =   480
         Width           =   555
      End
   End
   Begin VB.Frame FraCalificaciones 
      Caption         =   "Calificaciones"
      Height          =   2535
      Left            =   8640
      TabIndex        =   35
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Frame FraMayores 
      Height          =   735
      Left            =   8640
      TabIndex        =   32
      Top             =   3840
      Width           =   3255
      Begin VB.TextBox TxtMayores 
         Height          =   285
         Left            =   1560
         TabIndex        =   33
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Totales Mayores"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Frame FraCambio 
      Height          =   525
      Left            =   5295
      TabIndex        =   29
      Top             =   3840
      Width           =   3255
      Begin VB.TextBox TxtCambio 
         Height          =   285
         Left            =   1395
         TabIndex        =   31
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio"
         Height          =   195
         Left            =   330
         TabIndex        =   30
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.Frame FraMoneda 
      Caption         =   "Moneda"
      Height          =   615
      Left            =   5280
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CheckBox ChkMoneda 
         Caption         =   "Nacional"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox ChkMoneda 
         Caption         =   "Extranjera"
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraFecha 
      Height          =   735
      Left            =   5820
      TabIndex        =   10
      Top             =   0
      Width           =   2175
      Begin MSMask.MaskEdBox mskPeriodo1Del 
         Height          =   315
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   450
      End
   End
   Begin VB.Frame fraCredito 
      Caption         =   "Creditos"
      Height          =   1695
      Left            =   5280
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Hipotecario."
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   46
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Cons.no Revol."
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Cons.Revolvente."
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   44
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Microempresa."
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Pequeña Emp."
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Mediana Emp."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Grande Emp."
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Corporativo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   360
      Left            =   5280
      TabIndex        =   3
      Top             =   6840
      Width           =   1500
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   6960
      TabIndex        =   2
      Top             =   6840
      Width           =   1500
   End
   Begin VB.Frame FrameOperaciones 
      Caption         =   "Lista de Reportes"
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin MSComctlLib.TreeView tvwReporte 
         Height          =   6675
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   11774
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imglstFiguras"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   240
         Top             =   4320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmColocEvalRep.frx":031F
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmColocEvalRep.frx":0639
               Key             =   "Hijo"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraFuente 
      Height          =   1695
      Left            =   5280
      TabIndex        =   47
      Top             =   5055
      Visible         =   0   'False
      Width           =   3255
      Begin VB.OptionButton Opt 
         Caption         =   "Ninguno"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   50
         Top             =   130
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Todo"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   49
         Top             =   130
         Width           =   735
      End
      Begin VB.ListBox List1 
         Height          =   960
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   48
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.OLE OleExcel 
      Class           =   "Excel.Sheet.8"
      Height          =   255
      Left            =   3840
      OleObjectBlob   =   "FrmColocEvalRep.frx":0953
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmColocEvalRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim nMicro As Integer
'Dim nComercial As Integer
'Dim nConsumo As Integer
'Dim nHipo As Integer

' Inicio BRGO 20100608 BASILEA II Nuevos Tipos de créditos
    Dim nCorporativo As Integer
    Dim nGrandeEmp As Integer
    Dim nMedianaEmp As Integer
    Dim nPequenaEmp As Integer
    Dim nMicroEmp As Integer
    Dim nConsumoRev As Integer
    Dim nConsumoNoRev As Integer
    Dim nHipotecario As Integer
' Fin BRGO

'ALPA 20120118********
Dim lnTpoCambio As Currency
Dim oCon As DConecta
Dim R As New ADODB.Recordset
'*********************

Dim sLineasCred As String
Dim Co As New nColocEvalReporte
Dim FechaFinMes As Date
Dim fnRepoSelec As Long
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
'********NAGL 20180509
Dim xlsAplicacion As Excel.Application
Dim xlsLibro As Excel.Workbook
'********END NAGL
Dim xlHoja1 As Excel.Worksheet  'Micro
Dim xlHoja2 As Excel.Worksheet  'Consumo
Dim xlHoja3 As Excel.Worksheet  'Comercial
Dim xlHoja4 As Excel.Worksheet  'Hipotecario
Dim lnFF As Integer
Dim Progress As clsProgressBar
Dim WithEvents loRep As nColocEvalReporte
Attribute loRep.VB_VarHelpID = -1
Dim Index178100 As Integer ' *** MAVM:Auditoria
Dim Index178108 As Integer ' *** MAVM:Auditoria


'Private Function ObtieneDescripcionFuenteFinanciamiento(ByVal psFF As String) As String
'Dim lsSQL As String
'Dim rs As New ADODB.Recordset
'Dim lsDesc As String
'Dim Co As DConecta
'Set Co = New DConecta
'lsDesc = ""
'lsSQL = "Select * from ColocLineaCredito where cLineaCred like '__' and LTRIM(RTRIM(cValor)) ='" & psFF & "'"
'
'Co.AbreConexion
'Set rs = Co.CargaRecordSet(lsSQL)
'Co.CierraConexion
'
''rs.Open lsSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'If rs.BOF And rs.EOF Then
'    lsDesc = ""
'Else
'    lsDesc = rs!cDescripcion
'End If
'rs.Close
'Set rs = Nothing
'ObtieneDescripcionFuenteFinanciamiento = lsDesc
'End Function

Private Sub ImpHojaWProvisionCab(pnFila As Integer, psParteCab As String, Optional pbConsumo As Boolean = False)
Dim j As Integer
Dim lsNombre As String


    xlHoja1.Cells(pnFila, 1) = gsNomCmac
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 2) = "INFORME DE CLASIFICACION DE LA CARTERA DE CREDITOS, CONTINGENTES Y ARRENDAMIENTOS FINANCIEROS"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).HorizontalAlignment = xlCenter
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 7) = "AL " & Me.mskPeriodo1Del
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).HorizontalAlignment = xlCenter
    pnFila = pnFila + 1
    If psParteCab = "1" Then
        xlHoja1.Cells(pnFila, 6) = "(  MONEDA NACIONAL  ) "
    Else
        xlHoja1.Cells(pnFila, 6) = "( MONEDA EXTRANJERA ) "
    End If
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).HorizontalAlignment = xlCenter
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 1) = " CREDITOS "
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Cells(pnFila, 2) = "TOTAL CARTERA"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 2), xlHoja1.Cells(pnFila, 3)).Merge True
    xlHoja1.Cells(pnFila + 1, 2) = "Nro"
    xlHoja1.Cells(pnFila + 1, 3) = "Monto"
    xlHoja1.Cells(pnFila, 4) = "NORMALES"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 4), xlHoja1.Cells(pnFila, 5)).Merge True
    xlHoja1.Cells(pnFila + 1, 4) = "Nro"
    xlHoja1.Cells(pnFila + 1, 5) = "Monto"
    xlHoja1.Cells(pnFila, 6) = "POTENCIALES"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 6), xlHoja1.Cells(pnFila, 7)).Merge True
    xlHoja1.Cells(pnFila + 1, 6) = "Nro"
    xlHoja1.Cells(pnFila + 1, 7) = "Monto"
    xlHoja1.Cells(pnFila, 8) = "DEFICIENTES"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 8), xlHoja1.Cells(pnFila, 9)).Merge True
    xlHoja1.Cells(pnFila + 1, 8) = "Nro"
    xlHoja1.Cells(pnFila + 1, 9) = "Monto"
    xlHoja1.Cells(pnFila, 10) = "DUDOSOS"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 10), xlHoja1.Cells(pnFila, 11)).Merge True
    xlHoja1.Cells(pnFila + 1, 10) = "Nro"
    xlHoja1.Cells(pnFila + 1, 11) = "Monto"
    xlHoja1.Cells(pnFila, 12) = "PERDIDA"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 12), xlHoja1.Cells(pnFila, 13)).Merge True
    xlHoja1.Cells(pnFila + 1, 12) = "Nro"
    xlHoja1.Cells(pnFila + 1, 13) = "Monto"
    xlHoja1.Cells(pnFila, 14) = "PROVISION"
    xlHoja1.Cells(pnFila + 1, 14) = Format(Me.mskPeriodo1Del, "dd/mm/yyyy")
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 1, 14)).HorizontalAlignment = xlCenter
    CuadroExcel 1, pnFila, 14, pnFila + 1, True

    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 2), xlHoja1.Cells(350, 2)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 3), xlHoja1.Cells(350, 3)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 4), xlHoja1.Cells(350, 4)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 5), xlHoja1.Cells(350, 5)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 6), xlHoja1.Cells(350, 6)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 7), xlHoja1.Cells(350, 7)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 8), xlHoja1.Cells(350, 8)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 9), xlHoja1.Cells(350, 9)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 10), xlHoja1.Cells(350, 10)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 11), xlHoja1.Cells(350, 11)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 12), xlHoja1.Cells(350, 12)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 13), xlHoja1.Cells(350, 13)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 14), xlHoja1.Cells(350, 14)).NumberFormat = "#,#0.00;#,#0.00"


    xlHoja1.Range("A1").ColumnWidth = 18
    xlHoja1.Range("B1").ColumnWidth = 7
    xlHoja1.Range("C1").ColumnWidth = 15
    xlHoja1.Range("D1").ColumnWidth = 7
    xlHoja1.Range("E1").ColumnWidth = 15
    xlHoja1.Range("F1").ColumnWidth = 7
    xlHoja1.Range("G1").ColumnWidth = 15
    xlHoja1.Range("H1").ColumnWidth = 7
    xlHoja1.Range("I1").ColumnWidth = 15
    xlHoja1.Range("J1").ColumnWidth = 7
    xlHoja1.Range("K1").ColumnWidth = 15
    xlHoja1.Range("L1").ColumnWidth = 7
    xlHoja1.Range("M1").ColumnWidth = 15
    xlHoja1.Range("N1").ColumnWidth = 16

    ' Formato de celdas
    xlHoja1.Range("C1").NumberFormat = "#,#0.00"
    xlHoja1.Range("E1").NumberFormat = "#,#0.00"
    xlHoja1.Range("G1").NumberFormat = "#,#0.00"
    xlHoja1.Range("IC1").NumberFormat = "#,#0.00"
    xlHoja1.Range("K1").NumberFormat = "#,#0.00"
    xlHoja1.Range("M1").NumberFormat = "#,#0.00"
    xlHoja1.Range("N1").NumberFormat = "#,#0.00"

    pnFila = pnFila + 2


    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlAplicacion.Range("A1:Z100").Font.Size = 9

End Sub

Private Sub CuadroExcel(x1 As Integer, Y1 As Integer, x2 As Integer, Y2 As Integer, Optional lbLineasVert As Boolean = False)
Dim i, j As Integer

For i = x1 To x2
    xlHoja1.Range(xlHoja1.Cells(Y1, i), xlHoja1.Cells(Y1, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(Y2, i), xlHoja1.Cells(Y2, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Next i
If lbLineasVert = False Then
    For i = x1 To x2
        For j = Y1 To Y2
            xlHoja1.Range(xlHoja1.Cells(j, i), xlHoja1.Cells(j, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Next j
    Next i
End If
If lbLineasVert Then
    For j = Y1 To Y2
        xlHoja1.Range(xlHoja1.Cells(j, x1), xlHoja1.Cells(j, x1)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Next j
End If

For j = Y1 To Y2
    xlHoja1.Range(xlHoja1.Cells(j, x2), xlHoja1.Cells(j, x2)).Borders(xlEdgeRight).LineStyle = xlContinuous
Next j
End Sub

Public Function VerificaCheckCred(ope As Long) As Boolean
Dim Flag As Boolean
Select Case ope
'Case 178101, 178102: Flag = IIf(ChkCreditos(0).value = 1 Or _
'                        ChkCreditos(1).value = 1 Or _
'                        ChkCreditos(2).value = 1 Or _
'                        ChkCreditos(3).value = 1, True, False)

'Case 178102
'Case 178103
End Select
VerificaCheckCred = Flag
End Function

Public Function VerificaCheckMoneda(ope As Long) As Boolean
Dim Flag As Boolean
Select Case ope
Case 178101, 178102: Flag = IIf(ChkMoneda(0).value = 1 Or _
                            ChkMoneda(1).value = 1, True, False)
'Case 178102
'Case 178103
End Select
VerificaCheckMoneda = Flag
End Function



'***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Private Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean

Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox Err.Description, vbInformation, "Aviso"
  ExcelBegin = False
End Function

'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Private Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox Err.Description, vbInformation, "Aviso"
End Sub

'********************************
' Adiciona Hoja a LibroExcel
'********************************
Private Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet)
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = psHojName Then
       xlHoja1.Delete
       Exit For
    End If
Next
Set xlHoja1 = xlLibro.Worksheets.Add
xlHoja1.Name = psHojName
End Sub

Private Sub chkTipo_Click(Index As Integer)
Select Case Index
'    Case 0: nMicro = Index
'    Case 1: nComercial = Index
'    Case 2: nConsumo = Index
'    Case 3: nHipo = Index
    Case 0: nCorporativo = Index
    Case 1: nGrandeEmp = Index
    Case 2: nMedianaEmp = Index
    Case 3: nPequenaEmp = Index
    Case 4: nMicroEmp = Index
    Case 5: nConsumoRev = Index
    Case 6: nConsumoNoRev = Index
    Case 7: nHipotecario = Index
End Select
End Sub

Private Sub cmdImprimir_Click()
Dim Moneda As String
Dim Estado As Boolean
Dim tipo As Boolean
Dim Flag As Boolean
Dim s As String
Dim i As Integer
Dim Pag As Integer
Dim loPrevio As previo.clsprevio
Dim lsCadena As String
Dim Rcd As nRcdReportes
Dim lnMonto As Currency

Set Rcd = New nRcdReportes
Pag = 1

If ValFecha(mskPeriodo1Del) = False Then
    Exit Sub
End If 'NAGL 20180509

Select Case fnRepoSelec

Case 178101:
    If ChkTipo(0).value = 0 And ChkTipo(1).value = 0 And ChkTipo(2).value = 0 And ChkTipo(3).value = 0 And ChkTipo(4).value = 0 And ChkTipo(5).value = 0 And ChkTipo(6).value = 0 And ChkTipo(7).value = 0 Then
        MsgBox "Seleccione un tipo de crédito", vbInformation, "Aviso"
        Exit Sub
    End If
    Call loRep.Genera_Reporte178101(Rcd.GetServerConsol, mskPeriodo1Del, IIf(ChkTipo(0) = 1, 1, 0), IIf(ChkTipo(1) = 1, 1, 0), IIf(ChkTipo(2) = 1, 1, 0), IIf(ChkTipo(3) = 1, 1, 0), IIf(ChkTipo(4) = 1, 1, 0), IIf(ChkTipo(5) = 1, 1, 0), IIf(ChkTipo(6) = 1, 1, 0), IIf(ChkTipo(7) = 1, 1, 0))

Case 178102:
    If Me.ChkCF.value = 0 Then
        If ChkTipo(0).value = 0 And ChkTipo(1).value = 0 And ChkTipo(2).value = 0 And ChkTipo(3).value = 0 And ChkTipo(4).value = 0 And ChkTipo(5).value = 0 And ChkTipo(6).value = 0 And ChkTipo(7).value = 0 Then
            MsgBox "Seleccione un tipo de crédito", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    Call loRep.Genera_Reporte178102(Me.mskPeriodo1Del, IIf(ChkTipo(0) = 1, 1, 0), IIf(ChkTipo(1) = 1, 1, 0), IIf(ChkTipo(2) = 1, 1, 0), IIf(ChkTipo(3) = 1, 1, 0), IIf(Me.ChkCF.value = 1, True, False))
    
Case 178103:
    If ChkTipo(0).value = 0 And ChkTipo(1).value = 0 And ChkTipo(2).value = 0 And ChkTipo(3).value = 0 And ChkTipo(4).value = 0 And ChkTipo(5).value = 0 And ChkTipo(6).value = 0 And ChkTipo(7).value = 0 Then
        MsgBox "Seleccione un tipo de crédito", vbInformation, "Aviso"
        Exit Sub
    End If
    Call loRep.Genera_Reporte178103(Me.mskPeriodo1Del, IIf(ChkTipo(0) = 1, 1, 0), IIf(ChkTipo(1) = 1, 1, 0), IIf(ChkTipo(2) = 1, 1, 0), gsCodCMAC, IIf(ChkTipo(3) = 1, 1, 0))

Case 178104:
    Estado = IIf(OptEstado(0).value, False, True)
    
    Call loRep.Genera_Reporte178104(mskPeriodo1Del, Estado)

Case 178105:
    If ChkTipo(0).value = 0 And ChkTipo(1).value = 0 And ChkTipo(2).value = 0 And ChkTipo(3).value = 0 And ChkTipo(4).value = 0 And ChkTipo(5).value = 0 And ChkTipo(6).value = 0 And ChkTipo(7).value = 0 Then
        MsgBox "Seleccione un tipo de crédito", vbInformation, "Aviso"
        Exit Sub
    End If
    Call loRep.Genera_Reporte178105(Me.mskPeriodo1Del, IIf(ChkTipo(0) = 1, 1, 0), IIf(ChkTipo(1) = 1, 1, 0), IIf(ChkTipo(2) = 1, 1, 0), IIf(ChkTipo(3) = 1, 1, 0))
    
Case 178106:
    Estado = IIf(OptEstado(0).value, False, True)
    
    Call loRep.Genera_Reporte178106(mskPeriodo1Del, Estado, False, IIf(ChkCF.value, True, False))

Case 178107: 'TOTAL DE CARTERA A EXCELL
    Call loRep.Genera_Reporte178106(mskPeriodo1Del, Estado, True)

Case gColCalifCalificacionCartera: '**DAOR 20071124, Reporte de Calificacion de Cartera
    Call MostrarReporteCalificacionCartera(mskPeriodo1Del, TxtCambio.Text)
'ALPA 20080604, Anexo 5D Cartera Cofide
Case 178109:
    Call MostrarReporteAnexo5DCofide(mskPeriodo1Del, TxtCambio.Text)
Case 178110:
    Call MostrarReporteReclasificacionCreditos(Format(mskPeriodo1Del, "YYYYMMDD"))
Case 178111:
    'If MsgBox("Imprime nuevo reporte", vbQuestion + vbYesNo, "Atención") = vbYes Then
        'Call Reporte2A1Nuevo '*** PEAC 20170710
    'Else
        Call Reporte2A1 'Se realizaron implementaciones en este método, considerarlo en el método Reporte2A1Nuevo NAGL ERS020-2018
    'End If
Case 178112
    Call ReporteComparativoCartera
Case 178201:
              If CboSigno.ListIndex = -1 Then
                   MsgBox "Elija una de las Opciones de la Lista", vbInformation, "AVISO"
                   Exit Sub
              End If
              If Not IsNumeric(TxtMonto) And TxtMonto <> "" Then
                   MsgBox "Ingrese una cantidad numérica", vbInformation, "AVISO"
                   Exit Sub
              End If

             Moneda = IIf(OptMoneda(0).value, gMonedaNacional, gMonedaExtranjera)
             tipo = IIf(OptTipo(0).value, True, False)
             Estado = IIf(OptEstado(0).value, True, False)
             lnMonto = CCur(IIf(TxtMonto = "", 0, TxtMonto))
             lsCadena = loRep.Genera_Reporte178201(Rcd.GetServerConsol, Moneda, CboSigno.Text, lnMonto, tipo, Estado, gsCodCMAC, gsNomCmac, gsNomAge)
             Set loPrevio = New previo.clsprevio
                 loPrevio.Show lsCadena, "PROVISIONES", True
             Set loPrevio = Nothing


Case 178202:
            If CboSigno.ListIndex = -1 Then
                   MsgBox "Elija una de las Opiones de la Lista", vbInformation, "AVISO"
                   Exit Sub
              End If
              If Not IsNumeric(TxtMonto) And TxtMonto <> "" Then
                   MsgBox "Ingrese una cantidad numerica", vbInformation, "AVISO"
                   Exit Sub
              End If
              If Not ValFecha(mskPeriodo1Del) Then
                Exit Sub
               End If
              Moneda = IIf(OptMoneda(0).value, gMonedaNacional, gMonedaExtranjera)
              tipo = IIf(OptTipo(0).value, True, False)
              Estado = IIf(OptEstado(0).value, True, False)
              lnMonto = CCur(IIf(TxtMonto = "", 0, TxtMonto))
              lsCadena = loRep.Genera_Reporte178202(Moneda, CboSigno.Text, lnMonto, tipo, Estado, gsNomCmac, gsNomAge)
              Set loPrevio = New previo.clsprevio
                    loPrevio.Show lsCadena, "PROVISIONES", True
              Set loPrevio = Nothing

Case 178203:
              If CboSigno.ListIndex = -1 Then
                   MsgBox "Elija una de las Opiones de la Lista", vbInformation, "AVISO"
                   Exit Sub
              End If
              If Not IsNumeric(TxtMonto) And TxtMonto <> "" Then
                   MsgBox "Ingrese una cantidad numerica", vbInformation, "AVISO"
                   Exit Sub
              End If
              Moneda = IIf(OptMoneda(0).value, gMonedaNacional, gMonedaExtranjera)
              tipo = IIf(OptTipo(0).value, True, False)
              Estado = IIf(OptEstado(0).value, True, False)
              lnMonto = CCur(IIf(TxtMonto = "", 0, TxtMonto))
               lsCadena = loRep.Genera_Reporte178203(Moneda, CboSigno.Text, lnMonto, tipo, Estado, gsCodCMAC, gsNomCmac, gsNomAge)
               Set loPrevio = New previo.clsprevio
                    loPrevio.Show lsCadena, "PROVISIONES", True
                Set loPrevio = Nothing

Case 178204:  'ReporteGeneralProvisionesPignoraticio
              If Not ValFecha(mskPeriodo1Del) Then
                Exit Sub
               End If
               lsCadena = loRep.Genera_Reporte178204(gsNomCmac, gsNomAge, gdFecSis)
               Set loPrevio = New previo.clsprevio
                    loPrevio.Show lsCadena, "PROVISIONES", True
                Set loPrevio = Nothing

Case 178205:
                ' Falta Hacer Validacion de Entrada
                If Not IsNumeric(TxtCambio) And TxtMonto <> "" Then
                    MsgBox "Ingrese una cantidad numerica", vbInformation, "AVISO"
                    Exit Sub
              End If

              If CDbl(TxtCambio.Text) = 0 Then
                    MsgBox "Ingrese Tipo de Cambio", vbInformation, "AVISO"
                    Exit Sub
               End If

              lnMonto = CCur(IIf(TxtMonto = "", 0, TxtMonto))
              Call loRep.Genera_Reporte178205(Me.TxtCambio, lnMonto, Me.mskPeriodo1Del)

Case 178206:

                If Not IsNumeric(Me.TxtCambio) Then
                        MsgBox "Ingrese Tipo de Cambio", vbInformation, "AVISO"
                        Exit Sub
                End If
                If CDbl(TxtCambio.Text) = 0 Then
                    MsgBox "Ingrese Tipo de Cambio", vbInformation, "AVISO"
                    Exit Sub
               End If
                If Not ValFecha(mskPeriodo1Del) Then
                    Exit Sub
                End If

                sLineasCred = RecupFuente
                If VerificaFuente = 0 Then
                    sLineasCred = "("
                    For i = 0 To List1.ListCount - 1
                        sLineasCred = sLineasCred & "'" & Trim(Right(List1.List(i), 10)) & "',"
                    Next i
                    sLineasCred = Mid(sLineasCred, 1, Len(sLineasCred) - 1) & ")"
                End If
                lnMonto = CCur(IIf(TxtMonto = "", 0, TxtMonto))
                Call loRep.Genera_Reporte178206(val(Me.TxtCambio), lnMonto, sLineasCred, gdFecSis, gsNomCmac, gsCodCMAC)

Case 178207:
              If Not IsNumeric(TxtCambio) Then
                   MsgBox "Ingrese Tipo de Cambio", vbInformation, "AVISO"
                   Exit Sub
              End If
              If CDbl(TxtCambio.Text) = 0 Then
                    MsgBox "Ingrese Tipo de Cambio", vbInformation, "AVISO"
                    Exit Sub
               End If

              lnMonto = CCur(IIf(TxtMonto = "", 0, TxtMonto))
              Call loRep.Genera_Reporte178207(Rcd.GetServerConsol, val(Me.TxtCambio), gsNomCmac, gdFecDataFM)

Case 178208:
                'CAAU
                'Anexo 5
                If Not IsNumeric(Me.TxtCambio) Then
                        MsgBox "Ingrese Tipo de Cambio", vbInformation, "AVISO"
                        Exit Sub
                End If
                If CDbl(TxtCambio.Text) = 0 Then
                    MsgBox "Ingrese Tipo de Cambio", vbInformation, "AVISO"
                    Exit Sub
               End If
                If Not ValFecha(mskPeriodo1Del) Then
                    Exit Sub
                End If
                Call loRep.Genera_Reporte178208(val(Me.TxtCambio), Me.mskPeriodo1Del.Text, ChkCF.value, gConstSistServCentralRiesgos, gdFecData, gsNomCmac, gsCodCMAC)

Case 178209:
                'CAAU
                'Anexo 5 - D
                If Not IsNumeric(Me.TxtCambio) Then
                        MsgBox "Ingrese Tipo de Cambio", vbInformation, "AVISO"
                        Exit Sub
                End If
                If CDbl(TxtCambio.Text) = 0 Then
                    MsgBox "Ingrese Tipo de Cambio", vbInformation, "AVISO"
                    Exit Sub
               End If
                If Not ValFecha(mskPeriodo1Del) Then
                    Exit Sub
                End If

                If VerificaFuente = 0 Then
                    MsgBox "Marcar almenos una funete de ingreso", vbInformation, "AVISO"
                    Exit Sub
                End If
                Call loRep.Genera_Reporte178209(val(Me.TxtCambio), mskPeriodo1Del.Text, Me.ChkCF.value, gConstSistServCentralRiesgos, gdFecData, RecupFuente, gsNomCmac, gsCodCMAC)

Case 178210:
                'CAAU
                'Hoja de Trabajo
                If Not IsNumeric(Me.TxtCambio) Then
                        MsgBox "Ingrese Tipo de Cambio", vbInformation, "AVISO"
                        Exit Sub
                End If

                If CDbl(TxtCambio.Text) = 0 Then
                    MsgBox "Ingrese Tipo de Cambio", vbInformation, "AVISO"
                    Exit Sub
               End If

                If Not ValFecha(mskPeriodo1Del) Then
                    Exit Sub
                End If
                Call loRep.Genera_Reporte178210(val(Me.TxtCambio), mskPeriodo1Del.Text, IIf(Me.ChkCF.value = "1", True, False), gConstSistServCentralRiesgos, gdFecData, gsNomCmac, gsCodCMAC)
    
Case 178211
            Call loRep.ReporteSaldosProvxProd(val(TxtCambio))
'ALPA 20090615*****************************************************
Case 178212:
                'CAAU
                'Anexo 5 - D-COFIDE
                If Not IsNumeric(Me.TxtCambio) Then
                        MsgBox "Ingrese Tipo de Cambio", vbInformation, "AVISO"
                        Exit Sub
                End If
                If CDbl(TxtCambio.Text) = 0 Then
                    MsgBox "Ingrese Tipo de Cambio", vbInformation, "AVISO"
                    Exit Sub
               End If
                If Not ValFecha(mskPeriodo1Del) Then
                    Exit Sub
                End If

'                If VerificaFuente = 0 Then
'                    MsgBox "Marcar almenos una funete de ingreso", vbInformation, "AVISO"
'                    Exit Sub
'                End If
                Call loRep.Genera_Reporte178209(val(Me.TxtCambio), mskPeriodo1Del.Text, Me.ChkCF.value, gConstSistServCentralRiesgos, gdFecData, "", gsNomCmac, gsCodCMAC, 1)
'*******************************************************************
''ALPA 20090615*****************************************************
Case 178213:

                Call loRep.Genera_Reporte178213(mskPeriodo1Del.Text, Me.ChkCF.value, gConstSistServCentralRiesgos, gdFecData, "", gsNomCmac, gsCodCMAC)
'*******************************************************************
Case 178214
                Call Genera_Reporte178214(True, Right(mskPeriodo1Del, 2), Mid(mskPeriodo1Del, 4, 2), nVal(TxtCambio), UCase(Format(mskPeriodo1Del, "MMMM")), 1)
'ALPA 20120215******************************************************
                Call Genera_Reporte178214(True, Right(mskPeriodo1Del, 2), Mid(mskPeriodo1Del, 4, 2), nVal(TxtCambio), UCase(Format(mskPeriodo1Del, "MMMM")), 1)
'*******************************************************************
Case 178215
                If Trim(txtVPRiesgo.Text) = "" Then
                    MsgBox "Ingresar Valor Patrimonial en Riesgo", vbCritical
                    txtVPRiesgo.SetFocus
                End If
                If Trim(txtVPRiesgoAnterior.Text) = "" Then
                    MsgBox "Ingresar Valor Patrimonial en Riesgo Anterior", vbCritical
                    txtVPRiesgoAnterior.SetFocus
                End If
                Call Reporte178215(txtVPRiesgo.Text, txtVPRiesgoAnterior.Text)
End Select

Set Rcd = Nothing
End Sub

Public Sub HabilitaFrame(ByRef v1 As Boolean, ByRef v2 As Boolean, _
                         ByRef v3 As Boolean, ByRef v4 As Boolean, _
                         ByRef v5 As Boolean, ByRef v6 As Boolean, _
                         ByRef v7 As Boolean, ByRef v8 As Boolean, _
                         ByRef v9 As Boolean, ByRef v10 As Boolean, _
                         ByRef v11 As Boolean, ByRef v12 As Boolean, _
                         Optional v13 As Boolean = False, _
                         Optional v14 As Boolean = False)
With Me
    .fraCredito.Visible = v1
    .FraMoneda.Visible = v2
    .Frame.Visible = v3
    .FraTipo.Visible = v4
    .FraEstado.Visible = v5
    .FraMoneda2.Visible = v6
    .FraFecha.Visible = v7
    .FraCambio.Visible = v8
    .FraCalificaciones.Visible = v9
    .FraMayores.Visible = v10
    .FraIntervalos.Visible = v11
    .FraFuente.Visible = v12
    .FraCF.Visible = v13
    .frmPatrimonio.Visible = v14
End With
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Function VerificaFuente() As Integer
Dim i As Integer
Dim n As Integer
n = 0
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) = True Then
        n = n + 1
    End If
Next i
VerificaFuente = n
End Function

Private Sub Form_Load()
Dim LineaCred As nCredRepoFinMes
Dim RsCbo As ADODB.Recordset
Dim oTipCambio As nTipoCambio 'DAOR 20071124

Set LineaCred = New nCredRepoFinMes
Set RsCbo = New ADODB.Recordset
Set RsCbo = LineaCred.ObtieneLineaCredito
Set loRep = New nColocEvalReporte
Set Progress = New clsProgressBar
'Set Co = New nColocEvalReporte
'FechaFinMes = loRep.GetUltimaFechaCierre
Me.mskPeriodo1Del = gdFecData
FechaFinMes = gdFecData
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
CargaMenu
Call CargaListLineaCred(RsCbo)

'**DAOR 20071124***********************
Set oTipCambio = New nTipoCambio
    TxtCambio.Text = Format(oTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "#0.000")
Set oTipCambio = Nothing
'**************************************
Set LineaCred = Nothing
Set RsCbo = Nothing
Set oCon = New DConecta 'ALPA 20120118
CentraForm Me 'NAGL 20180509
End Sub
Public Function RecupFuente() As String
Dim i As Integer
    RecupFuente = "("
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
          RecupFuente = RecupFuente & "'" & Trim(Right(List1.List(i), 10)) & "',"
        End If
    Next i
    RecupFuente = Mid(RecupFuente, 1, Len(RecupFuente) - 1) & ")"
End Function

Public Sub CargaListLineaCred(ByRef rs As ADODB.Recordset)
Dim LineaCred As String
 If rs Is Nothing Then Exit Sub
While Not rs.EOF
    LineaCred = rs(1) & Space(50) & rs(0)
    List1.AddItem LineaCred
    rs.MoveNext
Wend
End Sub

Public Function ValFecha(lsControl As Control) As Boolean
   If Mid(lsControl, 1, 2) > 0 And Mid(lsControl, 1, 2) <= 31 Then
        If Mid(lsControl, 4, 2) > 0 And Mid(lsControl, 4, 2) <= 12 Then
            If Mid(lsControl, 7, 4) >= 1900 And Mid(lsControl, 7, 4) <= 9999 Then
               If IsDate(lsControl) = False Then
                    ValFecha = False
                    MsgBox "Formato de fecha no es válido", vbInformation, "Aviso"
                    lsControl.SetFocus
                    Exit Function
               Else
                    ValFecha = True
               End If
            Else
                ValFecha = False
                MsgBox "Año de Fecha no es válido", vbInformation, "Aviso"
                lsControl.SetFocus
                lsControl.SelStart = 6
                lsControl.SelLength = 4
                Exit Function
            End If
        Else
            ValFecha = False
            MsgBox "Mes de Fecha no es válido", vbInformation, "Aviso"
            lsControl.SetFocus
            lsControl.SelStart = 3
            lsControl.SelLength = 2
            Exit Function
        End If
    Else
        ValFecha = False
        MsgBox "Dia de Fecha no es válido", vbInformation, "Aviso"
        lsControl.SetFocus
        lsControl.SelStart = 0
        lsControl.SelLength = 2
        Exit Function
    End If
End Function

Private Sub CargaMenu()
Dim i As Integer ' ***MAVM:Auditoria
Dim clsGen As DGeneral
Dim rsUsu As Recordset
Dim sOperacion As String
Dim sOpeCod As String
Dim sOpePadre As String
Dim sOpeHijo As String
Dim sOpeHijito As String
Dim nodOpe As Node
Dim lsTipREP As String
lsTipREP = "178"
Set clsGen = New DGeneral

'ARCV 20-07-2006
'Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, lsTipREP, MatOperac, NroRegOpe)
Set rsUsu = clsGen.GetOperacionesUsuario_NEW(lsTipREP, , gRsOpeRepo)

Set clsGen = Nothing
Do While Not rsUsu.EOF
    sOpeCod = rsUsu("cOpeCod")
    sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
    Select Case rsUsu("nOpeNiv")
        Case "1"
            sOpePadre = "P" & sOpeCod
         Set nodOpe = tvwReporte.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "H" & sOpeCod
            Set nodOpe = tvwReporte.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
            nodOpe.Tag = sOpeCod
    End Select
    i = i + 1 ' ***MAVM:Auditoria
    If sOpeCod = "178100" Then Index178100 = i ' ***MAVM:Auditoria
    If sOpeCod = "178108" Then Index178108 = i ' ***MAVM:Auditoria
    rsUsu.MoveNext
Loop
rsUsu.Close
Set rsUsu = Nothing
End Sub

Private Sub loRep_CloseProgress()
Progress.CloseForm Me
End Sub

Private Sub loRep_Progress(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Generando Reporte", "Procesando ..."
End Sub

Private Sub loRep_ShowProgress()
    Progress.ShowForm Me
End Sub

Private Sub opt_Click(Index As Integer)
Dim bCheck As Boolean
Dim i As Integer
    If Index = 0 Then
        bCheck = True
    Else
        bCheck = False
    End If
    If List1.ListCount <= 0 Then
        Exit Sub
    End If
    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = bCheck
    Next i

End Sub

'*** MAVM: Modulo de Auditoria 21/08/2008
' Para Mostrar seleccioando El Reporte de Calificacion de Cartera
' por defecto en el Modulo de Auditoria
Public Sub Inicializar_operacion()
'    tvwReporte.Nodes(Index178100).Selected = True
'    tvwReporte_NodeClick tvwReporte.SelectedItem
'    tvwReporte.Nodes(Index178100).Expanded = True
'
'    tvwReporte.Nodes(Index178108).Selected = True
'    tvwReporte_NodeClick tvwReporte.SelectedItem
'    tvwReporte.Nodes(Index178108).Expanded = True
'
'    tvwReporte.Enabled = False
'    tvwReporte.HideSelection = False
End Sub
'*** MAVM: Modulo de Auditoria 21/08/2008

Private Sub tvwReporte_NodeClick(ByVal Node As MSComctlLib.Node)
Dim NodRep  As Node
Dim lsDesc As String

Set NodRep = tvwReporte.SelectedItem

If NodRep Is Nothing Then
   Exit Sub
End If

lsDesc = Mid(NodRep.Text, 8, Len(NodRep.Text) - 7)
fnRepoSelec = CLng(NodRep.Tag)
CmdImprimir.Enabled = True
Me.CboSigno.Visible = True
Me.TxtMonto.Visible = True
Me.Label2.Visible = True
Select Case fnRepoSelec
    'REPORTES EVALUACION DE CARTERA
    Case 178101
        Call HabilitaFrame(True, False, False, False, False, False, True, False, False, False, False, False)
        Me.mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    Case 178102
        Call HabilitaFrame(True, False, False, False, False, False, True, False, False, False, False, False, True)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    Case 178103
        Call HabilitaFrame(True, False, False, False, False, False, True, False, False, False, False, False)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    Case 178104
        Call HabilitaFrame(False, False, True, False, True, False, True, False, False, False, False, False)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
        Me.CboSigno.Visible = False
        Me.TxtMonto.Visible = False
        Me.Label2.Visible = False
        
    Case 178105
        Call HabilitaFrame(True, False, False, False, True, False, True, False, False, False, False, False)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
        Me.CboSigno.Visible = False
        Me.TxtMonto.Visible = False
        Me.Label2.Visible = False
        
    Case 178106 'CARTERA REFINANCIADA
        Call HabilitaFrame(False, False, True, False, True, False, True, False, False, False, False, False, True)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
        Me.CboSigno.Visible = False
        Me.TxtMonto.Visible = False
        Me.Label2.Visible = False
    
    
    
    Case 178107 'TOTAL DE CARTERA
        Call HabilitaFrame(False, False, False, False, False, False, True, False, False, False, False, False)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
            
    Case gColCalifCalificacionCartera  '**DAOR 20071124, Reporte de calificación de cartera
        Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, False)
        Me.mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    'ALPA 20090604**************
    Case 178109
        Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, False)
        Me.mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    Case 178110
        Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, False)
        Me.mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    Case 178111
        Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, False)
        Me.mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    Case 178112
        Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, False)
        Me.mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")

    '***************************
    'REPORTES DE PROVISIONES
    Case 178201
        Call HabilitaFrame(False, False, True, True, True, True, False, False, False, False, False, False)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    Case 178202
        Call HabilitaFrame(False, False, True, False, True, True, False, False, False, False, False, False)
    Case 178203
        Call HabilitaFrame(False, False, True, False, True, True, False, False, False, False, False, False)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    Case 178204
        Call HabilitaFrame(False, False, False, False, False, False, False, False, False, False, False, False)
        
    Case 178205
        Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, False)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    Case 178206
        Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, True)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")

    Case 178207
    Call HabilitaFrame(False, False, False, False, False, False, False, True, False, False, False, False)
    TxtCambio = TCFijoMes
    CmdImprimir.Enabled = True

    'CAAU
    Case 178208 'Anexo 5
        Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, False, True)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")

    Case 178209
    'CAAU
    'ANEXO 5-D
        Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, True, True)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")

    Case 178210
    Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, False, True)
    mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    
    Case 178211
    Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, False, True)
    mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    'ALPA 20090615*****************************************************
     Case 178212
        Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, True, True)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    '******************************************************************
    'ALPA 20090703*****************************************************
    Case 178213
        Call HabilitaFrame(False, False, False, False, False, False, True, False, False, False, False, False, False)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
    'ALPA 20120118*****************************************************
    Case 178214
        Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, False, False)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")
        
    '******************************************************************
    Case 178215
        Call HabilitaFrame(False, False, False, False, False, False, True, True, False, False, False, False, False, True)
        mskPeriodo1Del.Text = Format(FechaFinMes, "dd/mm/yyyy")

Case Else
    Call HabilitaFrame(False, False, False, False, False, False, False, False, False, False, False, False)
    CmdImprimir.Enabled = False
    Me.mskPeriodo1Del.Text = "__/__/____"
    Me.mskPeriodo1Del.Enabled = True
End Select

Set NodRep = Nothing
End Sub

'**DAOR 20071124
'**Muestra el reporte de calificación de cartera en formato Excel
Public Sub MostrarReporteCalificacionCartera(pdFechaProc As Date, pnTipCamb As Double)
Dim oNCalif As COMNCredito.NCOMColocEval
Dim R As ADODB.Recordset
Dim lMatCabecera As Variant
Dim lsmensaje As String
Dim lsNombreArchivo As String

'*****************NAGL Según 20180509************************
Dim oBarra As New clsProgressBar
Dim nprogress As Integer
Dim TituloProgress As String
Dim MensajeProgress As String
Dim RutaReport As String
oBarra.ShowForm FrmColocEvalRep
oBarra.Max = 10
    
oBarra.Progress 1, "178108 - Reporte de Calificación de Cartera", "GENERANDO EL ARCHIVO", "", vbBlue
TituloProgress = "178108 - Reporte de Calificación de Cartera"
MensajeProgress = "GENERANDO EL ARCHIVO"
'***********************END NAGL*****************************

    lsNombreArchivo = "CalificacionCartera"
    
    ReDim lMatCabecera(47, 2) 'NAGL Agregó(47, 2), Antes -> (46, 2)

    lMatCabecera(0, 0) = "cCtaCod": lMatCabecera(0, 1) = ""
    lMatCabecera(1, 0) = "Agencia": lMatCabecera(1, 1) = ""
    lMatCabecera(2, 0) = "Destino": lMatCabecera(2, 1) = ""
    lMatCabecera(3, 0) = "Codigo Cliente": lMatCabecera(3, 1) = ""
    lMatCabecera(4, 0) = "Cliente": lMatCabecera(4, 1) = ""
    lMatCabecera(5, 0) = "Documento": lMatCabecera(5, 1) = ""
    lMatCabecera(6, 0) = "Monto Aprob.": lMatCabecera(6, 1) = "N"
    lMatCabecera(7, 0) = "Estado": lMatCabecera(7, 1) = ""
    lMatCabecera(8, 0) = "Cuotas": lMatCabecera(8, 1) = "N"
    lMatCabecera(9, 0) = "Dia Fijo": lMatCabecera(9, 1) = "N"
    lMatCabecera(10, 0) = "Analista": lMatCabecera(10, 1) = ""
    lMatCabecera(11, 0) = "Tasa": lMatCabecera(11, 1) = "N"
    lMatCabecera(12, 0) = "Linea": lMatCabecera(12, 1) = ""
    lMatCabecera(13, 0) = "Fec. Desemb": lMatCabecera(13, 1) = "D"
    lMatCabecera(14, 0) = "Saldo Cap.": lMatCabecera(14, 1) = "N"
    lMatCabecera(15, 0) = "Cuota Actual": lMatCabecera(15, 1) = "N"
    lMatCabecera(16, 0) = "Tipo Personeria": lMatCabecera(16, 1) = ""
    lMatCabecera(17, 0) = "CIIU": lMatCabecera(17, 1) = ""
    lMatCabecera(18, 0) = "Direccion": lMatCabecera(18, 1) = ""
    lMatCabecera(19, 0) = "Calif. Anterior": lMatCabecera(19, 1) = ""
    lMatCabecera(20, 0) = "Calif. Actual": lMatCabecera(20, 1) = ""
    lMatCabecera(21, 0) = "Dias Atraso": lMatCabecera(21, 1) = "N"
    lMatCabecera(22, 0) = "Linea Credito": lMatCabecera(22, 1) = ""
    lMatCabecera(23, 0) = "Plazo": lMatCabecera(23, 1) = ""
    lMatCabecera(24, 0) = "Tipo Prod.": lMatCabecera(24, 1) = ""
    lMatCabecera(25, 0) = "Moneda": lMatCabecera(25, 1) = ""
    lMatCabecera(26, 0) = "Fec. Vcto": lMatCabecera(26, 1) = "D"
    lMatCabecera(27, 0) = "Int. Deveng.": lMatCabecera(27, 1) = "N"
    lMatCabecera(28, 0) = "Int. Suspen.": lMatCabecera(28, 1) = "N"
    lMatCabecera(29, 0) = "Por. Prov.": lMatCabecera(29, 1) = "N"
    lMatCabecera(30, 0) = "Prov. Con RCC": lMatCabecera(30, 1) = "N"
    lMatCabecera(31, 0) = "Prov. Sin RCC": lMatCabecera(31, 1) = "N"
    lMatCabecera(32, 0) = "Prov.Ant.Sin RCC": lMatCabecera(32, 1) = "N"
    lMatCabecera(33, 0) = "Prov.Ant Con RCC": lMatCabecera(33, 1) = "N"
    lMatCabecera(34, 0) = "Saldo Deudor": lMatCabecera(34, 1) = "N"
    lMatCabecera(35, 0) = "GAR. PREF": lMatCabecera(35, 1) = "N"
    lMatCabecera(36, 0) = "GAR.NO PREF": lMatCabecera(36, 1) = "N"
    lMatCabecera(37, 0) = "GAR. AUTOL": lMatCabecera(37, 1) = "N"
    lMatCabecera(38, 0) = "TIPO GAR CALIF": lMatCabecera(38, 1) = "N"
    lMatCabecera(39, 0) = "ALINEADO": lMatCabecera(39, 1) = ""
    lMatCabecera(40, 0) = "Condicion": lMatCabecera(40, 1) = ""
    lMatCabecera(41, 0) = "Calif.Sin Alin.": lMatCabecera(41, 1) = ""
    lMatCabecera(42, 0) = "Calif.Sist.F.": lMatCabecera(42, 1) = ""
    lMatCabecera(43, 0) = "Prov.Sin Alin.": lMatCabecera(43, 1) = "N"
    lMatCabecera(44, 0) = "Prov.Sist.F.": lMatCabecera(44, 1) = "N"
    lMatCabecera(45, 0) = "Cliente Unico CMACM": lMatCabecera(45, 1) = ""
    lMatCabecera(46, 0) = "Var.Cron.Pag": lMatCabecera(46, 1) = "" 'NAGL 20180509
    
    oBarra.Progress 3, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180509
   
    Set oNCalif = New COMNCredito.NCOMColocEval
     'CTI3 ERS003-2020 **********************************************************
'    'WIOR 20160623 ***
'    If Not oNCalif.SobreEndVerificaCartera Then
'        MsgBox "Se procederá a generar los codigos de SobreEndeudamiento para la Cartera Actual", vbInformation, "Aviso"
'        oNCalif.CalcularCapPagSobreEndEjecutarCartera (pdFechaProc) 'CTI1 20180706 'LUCV20190601, Cambio posición del método. Según JOEP
'        If oNCalif.SobreEndEjecutarCartera Then
'            MsgBox "Generación de Codigos de SobreEndeudamiento culminado satisfactoriamente", vbInformation, "Aviso"
'            'oNCalif.CalcularCapPagSobreEndEjecutarCartera (pdFechaProc) 'CTI1 20180706 'LUCV20190601, Comentó. Según JOEP
'        Else
'            MsgBox "Ocurió un error. Favor de volver a intentar, en caso que de persistir el error comunicate con el Departamento de T.I.", vbInformation, "Aviso"
'            Exit Sub
'        End If
'    Else
'        oNCalif.CalcularCapPagSobreEndEjecutarCartera (pdFechaProc) 'CTI1 20180706
'    End If
    'WIOR FIN *******
    '****************************************************************************
    oBarra.Progress 4, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180509
    
    Set R = oNCalif.darCalificacionCartera(pdFechaProc, pnTipCamb, lsmensaje)
    oBarra.Progress 7, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180509
    Set oNCalif = Nothing
    'JACA 20110708***********************************************
    oBarra.Progress 8, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180509
      
    If Not R Is Nothing Then
        'Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Calificación de Cartera", "", lsNombreArchivo, lMatCabecera, R, 2, , , True)Comentado by NAGL 20180509
        RutaReport = GeneraReporteEnArchivoExcelRptCalifCart(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Calificación de Cartera", "", lsNombreArchivo, lMatCabecera, R, 2, , , True) 'NAGL 20180509 ERS020-2018
    Else
        oBarra.Progress 10, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180509
        oBarra.CloseForm FrmColocEvalRep
        Set oBarra = Nothing 'NAGL 20180519
        MsgBox lsmensaje, vbInformation, "AVISO"
    End If
    'JACA END***************************************************
    '******NAGL BEGIN***************
    oBarra.Progress 10, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180509
    oBarra.CloseForm FrmColocEvalRep
    Set oBarra = Nothing
    MsgBox "Se ha generado el Archivo en " & RutaReport
    '******NAGL END 20180509*********
End Sub
Public Function GeneraReporteEnArchivoExcelRptCalifCart(ByVal psNomCmac As String, _
                                                        ByVal psNomAge As String, _
                                                        ByVal psCodUser As String, _
                                                        ByVal pdFecSis As Date, _
                                                        ByVal psTitulo As String, _
                                                        ByVal psSubTitulo As String, _
                                                        ByVal psNomArchivo As String, _
                                                        ByVal pMatCabeceras As Variant, _
                                                        ByVal prRegistros As ADODB.Recordset, _
                                                        Optional pnNumDecimales As Integer, _
                                                        Optional Visible As Boolean = False, _
                                                        Optional psNomHoja As String = "", _
                                                        Optional pbSinFormatDeReg As Boolean = False, _
                                                        Optional pbUsarCabecerasDeRS As Boolean = False, _
                                                        Optional psRuta As String = "") As String
    Dim rs As ADODB.Recordset
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim liLineas As Integer, i As Integer
    Dim fs As Scripting.FileSystemObject
    Dim lnNumColumns As Integer
    Dim psRutaReporte As String
    
    If Not (prRegistros.EOF And prRegistros.BOF) Then
        If pbUsarCabecerasDeRS = True Then
            lnNumColumns = prRegistros.Fields.count
        Else
            lnNumColumns = UBound(pMatCabeceras)
            lnNumColumns = IIf(prRegistros.Fields.count < lnNumColumns, prRegistros.Fields.count, prRegistros.Fields.count)
        End If

        If psNomHoja = "" Then psNomHoja = psNomArchivo
        psNomArchivo = psNomArchivo & "_" & psCodUser & ".xls"

        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application

        If psRuta = "" Then
            If fs.FileExists(App.Path & "\Spooler\" & psNomArchivo) Then
            fs.DeleteFile (App.Path & "\Spooler\" & psNomArchivo)
            End If
        Else
            If fs.FileExists(psRuta & psNomArchivo) Then
                fs.DeleteFile (psRuta & psNomArchivo)
            End If
        End If

        Set xlLibro = xlAplicacion.Workbooks.Add
        Set xlHoja1 = xlLibro.Worksheets.Add

        xlHoja1.Name = psNomHoja
        xlHoja1.Cells.Select

        'Cabeceras
        xlHoja1.Cells(2, 1) = psNomCmac
        xlHoja1.Cells(2, lnNumColumns) = Trim(Format(pdFecSis, "dd/mm/yyyy hh:mm:ss"))
        xlHoja1.Cells(2, 1) = psNomAge
        xlHoja1.Cells(2, lnNumColumns) = psCodUser
        xlHoja1.Cells(4, 1) = psTitulo
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, lnNumColumns)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(5, lnNumColumns)).HorizontalAlignment = xlCenter

        liLineas = 6
        If pbUsarCabecerasDeRS = True Then
            For i = 0 To prRegistros.Fields.count - 1
                xlHoja1.Cells(liLineas, i + 1) = prRegistros.Fields(i).Name
            Next i
        Else
            For i = 0 To lnNumColumns - 1
                If (i + 1) > UBound(pMatCabeceras) Then
                    xlHoja1.Cells(liLineas, i + 1) = prRegistros.Fields(i).Name
                Else
                    xlHoja1.Cells(liLineas, i + 1) = pMatCabeceras(i, 0)
                End If
            Next i
        End If

        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).HorizontalAlignment = xlCenter

        If pbSinFormatDeReg = False Then
            liLineas = liLineas + 1
            While Not prRegistros.EOF
                For i = 0 To lnNumColumns - 1
                    If pMatCabeceras(i, 1) = "" Then  'Verificamos si tiene tipo
                        xlHoja1.Cells(liLineas, i + 1) = prRegistros(i)
                    Else
                        Select Case pMatCabeceras(i, 1)
                            Case "S"
                                xlHoja1.Cells(liLineas, i + 1) = prRegistros(i)
                            Case "N"
                                xlHoja1.Cells(liLineas, i + 1) = Format(prRegistros(i), "#0.00")
                            Case "D"
                                xlHoja1.Cells(liLineas, i + 1) = IIf(Format(prRegistros(i), "yyyymmdd") = "19000101", "", Format(prRegistros(i), "dd/mm/yyyy"))
                        End Select
                    End If
                Next i
                liLineas = liLineas + 1
                prRegistros.MoveNext
            Wend
        Else
            xlHoja1.Range("A7").CopyFromRecordset prRegistros 'Copia el contenido del recordset a excel
        End If
        'If psRuta = "" Then
            'xlHoja1.SaveAs App.Path & "\Spooler\" & psNomArchivo
            'MsgBox "Se ha generado el Archivo en " & App.Path & "\Spooler\" & psNomArchivo
        'Else
            'xlHoja1.SaveAs psRuta & psNomArchivo
            'MsgBox "Se ha generado el Archivo en " & psRuta & psNomArchivo
        'End If
        
        xlHoja1.SaveAs App.Path & "\Spooler\" & psNomArchivo
        psRutaReporte = App.Path & "\Spooler\" & psNomArchivo
        If Visible Then
            xlAplicacion.Visible = True
            xlAplicacion.Windows(1).Visible = True
        Else
            xlLibro.Close
            xlAplicacion.Quit
        End If

        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
    End If
     GeneraReporteEnArchivoExcelRptCalifCart = psRutaReporte
End Function 'NAGL ERS020-2018

Public Sub MostrarReporteReclasificacionCreditos(psFechaProc As String)
Dim oNCalif As COMNCredito.NCOMColocEval
Dim R As ADODB.Recordset
Dim lMatCabecera As Variant
Dim lsmensaje As String
Dim lsNombreArchivo As String

    lsNombreArchivo = "ReclaCreditos"
    
    ReDim lMatCabecera(14, 2)

    lMatCabecera(0, 0) = "cAgeCod": lMatCabecera(0, 1) = ""
    lMatCabecera(1, 0) = "cAgeDescripcion": lMatCabecera(1, 1) = ""
    lMatCabecera(2, 0) = "cPersCod": lMatCabecera(2, 1) = ""
    lMatCabecera(3, 0) = "cPersNombre": lMatCabecera(3, 1) = ""
    lMatCabecera(4, 0) = "cCtaCod": lMatCabecera(4, 1) = ""
    lMatCabecera(5, 0) = "cTpoCredCod": lMatCabecera(5, 1) = ""
    lMatCabecera(6, 0) = "cTpcDesc": lMatCabecera(6, 1) = ""
    lMatCabecera(7, 0) = "cTpoCredCodAnt": lMatCabecera(7, 1) = ""
    lMatCabecera(8, 0) = "cTpcDesAnt": lMatCabecera(8, 1) = ""
    lMatCabecera(9, 0) = "nSaldo": lMatCabecera(9, 1) = ""
    lMatCabecera(10, 0) = "CapPago": lMatCabecera(10, 1) = ""
    
    Set oNCalif = New COMNCredito.NCOMColocEval
    Set R = oNCalif.darReclasificacionCartera(psFechaProc, lsmensaje)
    Set oNCalif = Nothing
           
    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Reclasificación de Cartera", "", lsNombreArchivo, lMatCabecera, R, 2, , , True)
End Sub
Public Sub MostrarReporteAnexo5DCofide(pdFechaProc As Date, pnTipCamb As Double)
Dim oNCalif As COMNCredito.NCOMColocEval
Dim R As ADODB.Recordset
Dim lMatCabecera As Variant
Dim lsmensaje As String
Dim lsNombreArchivo As String

    lsNombreArchivo = "Anexo5DCofide"
    
    ReDim lMatCabecera(40, 2)

    lMatCabecera(0, 0) = "CODIGO CLIENTE": lMatCabecera(0, 1) = ""
    lMatCabecera(1, 0) = "PRODUCTO": lMatCabecera(1, 1) = ""
    lMatCabecera(2, 0) = "LINEA COFIDE": lMatCabecera(2, 1) = ""
    lMatCabecera(3, 0) = "NOMBRE LINEA": lMatCabecera(3, 1) = ""
    lMatCabecera(4, 0) = "TIPO DE CREDITO": lMatCabecera(4, 1) = ""
    lMatCabecera(5, 0) = "CLIENTE": lMatCabecera(5, 1) = ""
    lMatCabecera(6, 0) = "CALIFICACION": lMatCabecera(6, 1) = ""
    lMatCabecera(7, 0) = "MONEDA": lMatCabecera(7, 1) = ""
    lMatCabecera(8, 0) = "MONTO": lMatCabecera(8, 1) = "N"
    lMatCabecera(9, 0) = "SALDO": lMatCabecera(9, 1) = "N"
    lMatCabecera(10, 0) = "SALDO S/.": lMatCabecera(10, 1) = "N"
    lMatCabecera(11, 0) = "PRV REQ S/.": lMatCabecera(11, 1) = "N"
    lMatCabecera(12, 0) = "PRV REQ PROC S/.": lMatCabecera(12, 1) = "N"
    lMatCabecera(13, 0) = "PRV RCC S/. ": lMatCabecera(13, 1) = "N"
    lMatCabecera(14, 0) = "TOTAL PRV": lMatCabecera(14, 1) = "N"
    lMatCabecera(15, 0) = "SALDO VENC S/.": lMatCabecera(15, 1) = "N"
    lMatCabecera(16, 0) = "TIPO PERSONA": lMatCabecera(16, 1) = ""
    lMatCabecera(17, 0) = "SEXO": lMatCabecera(17, 1) = ""
    lMatCabecera(18, 0) = "FEC.NAC.": lMatCabecera(18, 1) = "D"
    lMatCabecera(19, 0) = "DIRECCION": lMatCabecera(19, 1) = ""
    lMatCabecera(20, 0) = "DISTRITO": lMatCabecera(20, 1) = ""
    lMatCabecera(21, 0) = "TIPO DE DOCUMENTO": lMatCabecera(21, 1) = ""
    lMatCabecera(22, 0) = "NRO DE DOCUMENTO": lMatCabecera(22, 1) = ""
    lMatCabecera(23, 0) = "CODIGO SBS": lMatCabecera(23, 1) = ""
    lMatCabecera(24, 0) = "SECTOR ECONOMICO": lMatCabecera(24, 1) = ""
    lMatCabecera(25, 0) = "TASA DE INTERESES": lMatCabecera(25, 1) = "N"
    lMatCabecera(26, 0) = "FEC.APROBA": lMatCabecera(26, 1) = "D"
    lMatCabecera(27, 0) = "FEC.EJECUCION": lMatCabecera(27, 1) = "D"
    lMatCabecera(28, 0) = "FEC.VCMTO": lMatCabecera(28, 1) = "D"
    lMatCabecera(29, 0) = "CARTERA": lMatCabecera(29, 1) = ""
    lMatCabecera(30, 0) = "REFINANCIADO": lMatCabecera(30, 1) = ""
    lMatCabecera(31, 0) = "TOTAL CUOTAS": lMatCabecera(31, 1) = "N"
    lMatCabecera(32, 0) = "CUOTAS PAGADAS": lMatCabecera(32, 1) = "N"
    lMatCabecera(33, 0) = "CUOTAS PENDIENTES": lMatCabecera(33, 1) = "N"
    lMatCabecera(34, 0) = "CUOTAS VENCIDAS": lMatCabecera(34, 1) = "N"
    lMatCabecera(35, 0) = "CODIGO CIIU": lMatCabecera(35, 1) = ""
    lMatCabecera(36, 0) = "GARANTIAS": lMatCabecera(36, 1) = "N"
    lMatCabecera(37, 0) = "MAGNITUD": lMatCabecera(37, 1) = "N"
    lMatCabecera(38, 0) = "INDICADOR RIESGO CAMBIARIO CREDITICIO": lMatCabecera(38, 1) = "N"
    lMatCabecera(39, 0) = "CONDICION EN DIAS": lMatCabecera(39, 1) = "N"
    
    
    Set oNCalif = New COMNCredito.NCOMColocEval
    Set R = oNCalif.darAnexo5DCOFIDE(pdFechaProc, pnTipCamb, lsmensaje)
    Set oNCalif = Nothing
           
    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Anexo 5D de Cartera de COFIDE", "", lsNombreArchivo, lMatCabecera, R, 2, True, , True)
End Sub

Public Sub Reporte2A1()
    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja  As String
    Dim lsNombreAgencia As String
    Dim lsCodAgencia As String
    Dim lsMes As String
    Dim lnContador As Integer
    Dim lsArchivo As String
    'Dim xlsAplicacion As Excel.Application
    'Dim xlsLibro As Excel.Workbook
    'Dim xlHoja1 As Excel.Worksheet 'Comentado by NAGL 20180509, para que se tome en cuenta las variables declaradas para todos los métodos
    
    Dim rsCreditos As ADODB.Recordset
    
    Dim oCreditos As New COMDCredito.DCOMCredito
    Dim oCredCtaCont As COMDContabilidad.DCOMCtaCont
    Dim nSaldoCtaContMens As Currency
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Double
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim n As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Double
    Dim nPase As Integer
    Dim pnUIT As Currency
    Dim pdFecha As Date 'NAGL 20180515
    Dim pnTpoCambio As Currency 'NAGL 20180515
'On Error GoTo GeneraExcelErr

    '*****************NAGL 20180915************************
    Dim oBarra As New clsProgressBar
    Dim nprogress As Integer
    Dim TituloProgress As String
    Dim MensajeProgress As String
    Dim RutaReport As String
    oBarra.ShowForm FrmColocEvalRep
    oBarra.Max = 10
        
    oBarra.Progress 1, "178111 - 2A1 Activos y Contingencias ponderadas", "GENERANDO EL ARCHIVO", "", vbBlue
    TituloProgress = "178111 - 2A1 Activos y Contingencias ponderadas"
    MensajeProgress = "GENERANDO EL ARCHIVO"
    '***********************END NAGL*****************************
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "Reporte_2A1"
    'Primera Hoja ******************************************************
    lsNomHoja = "dep_efinan"
    '*******************************************************************
    lsArchivo1 = "\spooler\Reporte_2A1_" & gsCodUser & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx" 'Antes .xls 20180509
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then 'Antes .xls 20180509
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsArchivo & ".xlsx") 'Antes .xls 20180509
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
    
    nSaltoContador = 8
    
    oBarra.Progress 1, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180515
    
    'nMES = cboMes.ListIndex + 1
    Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = oCreditos.RecuperaBancosCajasClasificacion
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
    xlHoja1.Cells(3, 2) = Format(mskPeriodo1Del.Text, "DD/MM/YYYY")
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
            xlHoja1.Cells(rsCreditos!cBancosCajasCod, 9) = rsCreditos!cBancosCajasClasificacion
            xlHoja1.Cells(rsCreditos!cBancosCajasCod, 10) = rsCreditos!cBancosCajasClasificacionValor
            xlHoja1.Cells(rsCreditos!cBancosCajasCod, 11) = rsCreditos!cBancosCajasPorcentaje
            nSaltoContador = nSaltoContador + 1
            rsCreditos.MoveNext
            nContTotal = nContTotal + 1
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    '11030102
    Set oCredCtaCont = New COMDContabilidad.DCOMCtaCont
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030102", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(9, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030103", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(10, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030104", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(11, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030105", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(12, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030106", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(13, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030121", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(14, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030108", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(15, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030129", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(16, 4) = nSaldoCtaContMens
    
    '03
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030301", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(20, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030303", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(21, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030304", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(22, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030306", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(23, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030308", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(24, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030309", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(25, 4) = nSaldoCtaContMens + oCredCtaCont.ObtenerCtaContBalanceMensual("1108030309", mskPeriodo1Del.Text, "0", 1)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030310", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(26, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030311", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(27, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030312", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(28, 4) = nSaldoCtaContMens + oCredCtaCont.ObtenerCtaContBalanceMensual("1108030312", mskPeriodo1Del.Text, "0", 1)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030305", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(29, 4) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030313", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(30, 4) = nSaldoCtaContMens
        
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110304", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(32, 4) = nSaldoCtaContMens + oCredCtaCont.ObtenerCtaContBalanceMensual("11080304", mskPeriodo1Del.Text, "0", 1)
    
    oBarra.Progress 2, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180515
    
    'Segunda Hoja ******************************************************
    lsNomHoja = "Cta del Balance"
    '*******************************************************************
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
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1102", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(4, 3) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1103", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(12, 3) = nSaldoCtaContMens
    '*****************
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1108", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(12, 6) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1108030309", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(13, 6) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1108030312", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(14, 6) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11080304", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(15, 6) = nSaldoCtaContMens
    '*****************
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110301", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(13, 3) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110302", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(14, 3) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110303", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(15, 3) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110304", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(16, 3) = nSaldoCtaContMens

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110306", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(17, 3) = nSaldoCtaContMens

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14010906070503", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(18, 3) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14010906070510", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(19, 3) = nSaldoCtaContMens

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14010906070511", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(20, 3) = nSaldoCtaContMens

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1408090503", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(18, 7) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1408090510", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(19, 7) = nSaldoCtaContMens

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1408090511", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(20, 7) = nSaldoCtaContMens

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("13", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(21, 3) = nSaldoCtaContMens
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1302051903", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(22, 3) = nSaldoCtaContMens

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1302051905", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(23, 3) = nSaldoCtaContMens

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1305181905", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(24, 3) = nSaldoCtaContMens

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1308051803", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(25, 3) = nSaldoCtaContMens

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1308051805", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(26, 3) = nSaldoCtaContMens

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1308051829", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(27, 3) = nSaldoCtaContMens

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1302051929", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(28, 3) = nSaldoCtaContMens

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("170401", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(29, 3) = nSaldoCtaContMens
    
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1101", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(87, 3) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11010103", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(88, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110701", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(89, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11070901", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(90, 3) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1106", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(92, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1902", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(93, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1907", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(94, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1505", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(95, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("150701", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(96, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1901", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(97, 3) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1903", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(98, 3) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1906", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(99, 3) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1801", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(100, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1802", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(101, 3) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1803", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(102, 3) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1804", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(103, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1806", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(104, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1807", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(105, 3) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1904", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(106, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1602", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(107, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("7109", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(108, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1507", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(109, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("2903", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(110, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1908", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(111, 3) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("180902", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(101, 4) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("180903", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(102, 4) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("180904", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(103, 4) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("180906", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(104, 4) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("180907", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(105, 4) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("190409", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(106, 4) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1609", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(107, 4) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1509", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(109, 4) = Abs(nSaldoCtaContMens)
    
    'COMPROBACION DE CUENTAS DEL ACTIVO
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(137, 4) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("7102", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(138, 4) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("7109", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(139, 4) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("2903", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(140, 4) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1809", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(141, 4) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("190409", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(142, 4) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1609", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(143, 4) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1509", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(144, 4) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1409", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(145, 4) = Abs(nSaldoCtaContMens)
    
    oBarra.Progress 3, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180515
    'Cartera CF
    Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = New ADODB.Recordset
    Set rsCreditos = oCreditos.RecuperaCarteraCF(TxtCambio.Text, mskPeriodo1Del.Text)
    
    Do While Not rsCreditos.EOF
        'ALPA 20110629*******************************************************
        If rsCreditos!nTipoCredito = "381" Then
                xlHoja1.Cells(129, 3) = rsCreditos!nSaldo
        ElseIf rsCreditos!nTipoCredito = "481" Then
                xlHoja1.Cells(130, 3) = rsCreditos!nSaldo
        ElseIf rsCreditos!nTipoCredito = "581" Then
                xlHoja1.Cells(131, 3) = rsCreditos!nSaldo
        End If
        '********************************************************************
        rsCreditos.MoveNext
        If rsCreditos.EOF Then
            Exit Do
        End If
    Loop
    
    Set oCreditos = Nothing
    Set rsCreditos = Nothing
    
    Set oCreditos = New COMDCredito.DCOMCredito
    pnUIT = oCreditos.ObtnerUIT(CInt(Format(mskPeriodo1Del.Text, "YYYY")))
    
    Set oCreditos = Nothing
    Set rsCreditos = New ADODB.Recordset
    Set oCreditos = New COMDCredito.DCOMCredito
    
    Set rsCreditos = oCreditos.RecuperaCarteraAutoliquidable(TxtCambio.Text, pnUIT, mskPeriodo1Del.Text)
    
    Do While Not rsCreditos.EOF

        If rsCreditos!nTipoCredito = "7" Then
                xlHoja1.Cells(43, 4) = rsCreditos!nSaldo
        ElseIf rsCreditos!nTipoCredito = "3" Then
                xlHoja1.Cells(66, 4) = rsCreditos!nSaldo
        ElseIf rsCreditos!nTipoCredito = "4" Then
                xlHoja1.Cells(71, 4) = rsCreditos!nSaldo
        ElseIf rsCreditos!nTipoCredito = "5" Then
                xlHoja1.Cells(78, 4) = rsCreditos!nSaldo
        End If

        rsCreditos.MoveNext
        If rsCreditos.EOF Then
            Exit Do
        End If
    Loop
    
    oBarra.Progress 4, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180515
    'Segunda Hoja ******************************************************
    lsNomHoja = "Creditos"
    '*******************************************************************
    
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
    
    xlHoja1.Cells(1, 2) = UCase(fgDameNombreMes(CInt(Mid(mskPeriodo1Del.Text, 4, 2)))) 'JUEZ 20130208
    
    Set oCreditos = Nothing
    Set rsCreditos = Nothing
    
    Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = New ADODB.Recordset
    
    Set rsCreditos = oCreditos.RecuperaCarteraCreditoPorTramos(TxtCambio.Text, mskPeriodo1Del.Text)
    
    Do While Not rsCreditos.EOF

        If rsCreditos!cTpoCred = "3" And rsCreditos!cOrden = "1" Then
                xlHoja1.Cells(5, 4) = rsCreditos!K
                xlHoja1.Cells(5, 6) = rsCreditos!id
                xlHoja1.Cells(5, 8) = rsCreditos!P
                xlHoja1.Cells(5, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCred = "3" And rsCreditos!cOrden = "2" Then
                xlHoja1.Cells(6, 4) = rsCreditos!K
                xlHoja1.Cells(6, 6) = rsCreditos!id
                xlHoja1.Cells(6, 8) = rsCreditos!P
                xlHoja1.Cells(6, 11) = rsCreditos!PRCC
        End If
        
        If rsCreditos!cTpoCred = "4" And rsCreditos!cOrden = "1" Then
                xlHoja1.Cells(8, 4) = rsCreditos!K
                xlHoja1.Cells(8, 6) = rsCreditos!id
                xlHoja1.Cells(8, 8) = rsCreditos!P
                xlHoja1.Cells(8, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCred = "4" And rsCreditos!cOrden = "2" Then
                xlHoja1.Cells(9, 4) = rsCreditos!K
                xlHoja1.Cells(9, 6) = rsCreditos!id
                xlHoja1.Cells(9, 8) = rsCreditos!P
                xlHoja1.Cells(9, 11) = rsCreditos!PRCC
        End If
        
        If rsCreditos!cTpoCred = "5" And rsCreditos!cOrden = "1" Then
                xlHoja1.Cells(11, 4) = rsCreditos!K
                xlHoja1.Cells(11, 6) = rsCreditos!id
                xlHoja1.Cells(11, 8) = rsCreditos!P
                xlHoja1.Cells(11, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCred = "5" And rsCreditos!cOrden = "2" Then
                xlHoja1.Cells(12, 4) = rsCreditos!K
                xlHoja1.Cells(12, 6) = rsCreditos!id
                xlHoja1.Cells(12, 8) = rsCreditos!P
                xlHoja1.Cells(12, 11) = rsCreditos!PRCC
        End If
        
        If rsCreditos!cTpoCred = "7" And rsCreditos!cOrden = "1" Then
                xlHoja1.Cells(15, 4) = rsCreditos!K
                xlHoja1.Cells(15, 6) = rsCreditos!id
                xlHoja1.Cells(15, 8) = rsCreditos!P
                xlHoja1.Cells(15, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCred = "7" And rsCreditos!cOrden = "2" Then
                xlHoja1.Cells(16, 4) = rsCreditos!K
                xlHoja1.Cells(16, 6) = rsCreditos!id
                xlHoja1.Cells(16, 8) = rsCreditos!P
                xlHoja1.Cells(16, 11) = rsCreditos!PRCC
        End If
        'JUEZ 20131211 ************************************************
        If rsCreditos!cTpoCred = "7" And rsCreditos!cOrden = "5" Then
                xlHoja1.Cells(17, 4) = rsCreditos!K
                xlHoja1.Cells(17, 6) = rsCreditos!id
                xlHoja1.Cells(17, 8) = rsCreditos!P
                xlHoja1.Cells(17, 11) = rsCreditos!PRCC
        End If
        
        If rsCreditos!cTpoCred = "8" And rsCreditos!cOrden = "1" Then
                'xlHoja1.Cells(18, 4) = rsCreditos!K
                'xlHoja1.Cells(18, 6) = rsCreditos!id
                'xlHoja1.Cells(18, 8) = rsCreditos!P
                'xlHoja1.Cells(18, 11) = rsCreditos!PRCC
                xlHoja1.Cells(19, 4) = rsCreditos!K
                xlHoja1.Cells(19, 6) = rsCreditos!id
                xlHoja1.Cells(19, 8) = rsCreditos!P
                xlHoja1.Cells(19, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCred = "8" And rsCreditos!cOrden = "3" Then
                'xlHoja1.Cells(19, 4) = rsCreditos!K
                'xlHoja1.Cells(19, 6) = rsCreditos!id
                'xlHoja1.Cells(19, 8) = rsCreditos!P
                'xlHoja1.Cells(19, 11) = rsCreditos!PRCC
                xlHoja1.Cells(20, 4) = rsCreditos!K
                xlHoja1.Cells(20, 6) = rsCreditos!id
                xlHoja1.Cells(20, 8) = rsCreditos!P
                xlHoja1.Cells(20, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCred = "8" And rsCreditos!cOrden = "2" Then
                'xlHoja1.Cells(20, 4) = rsCreditos!K
                'xlHoja1.Cells(20, 6) = rsCreditos!id
                'xlHoja1.Cells(20, 8) = rsCreditos!P
                'xlHoja1.Cells(20, 11) = rsCreditos!PRCC
                xlHoja1.Cells(21, 4) = rsCreditos!K
                xlHoja1.Cells(21, 6) = rsCreditos!id
                xlHoja1.Cells(21, 8) = rsCreditos!P
                xlHoja1.Cells(21, 11) = rsCreditos!PRCC
        End If
        'END JUEZ *****************************************************
        
        nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("27010112", mskPeriodo1Del.Text, "0", 1)
        xlHoja1.Cells(6, 9) = Abs(nSaldoCtaContMens)

        nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("27010113", mskPeriodo1Del.Text, "0", 1)
        xlHoja1.Cells(9, 9) = Abs(nSaldoCtaContMens)

        nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("27010102", mskPeriodo1Del.Text, "0", 1)
        xlHoja1.Cells(12, 9) = Abs(nSaldoCtaContMens)

        rsCreditos.MoveNext
        If rsCreditos.EOF Then
            Exit Do
        End If
    Loop
    
    Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = New ADODB.Recordset
    
    Set rsCreditos = oCreditos.RecuperaSaldosInteresDiferidos(TxtCambio.Text, mskPeriodo1Del.Text)
    
    Do While Not rsCreditos.EOF

        If rsCreditos!cTpoCredCod = "3" Then
                xlHoja1.Cells(5, 10) = rsCreditos!nConAtraso
                xlHoja1.Cells(6, 10) = rsCreditos!nSinAtraso
        ElseIf rsCreditos!cTpoCredCod = "4" Then
                xlHoja1.Cells(8, 10) = rsCreditos!nConAtraso
                xlHoja1.Cells(9, 10) = rsCreditos!nSinAtraso
        ElseIf rsCreditos!cTpoCredCod = "5" Then
                xlHoja1.Cells(11, 10) = rsCreditos!nConAtraso
                xlHoja1.Cells(12, 10) = rsCreditos!nSinAtraso
        ElseIf rsCreditos!cTpoCredCod = "7" Then
                xlHoja1.Cells(15, 10) = rsCreditos!nConAtraso
                xlHoja1.Cells(16, 10) = rsCreditos!nSinAtraso
        End If

        rsCreditos.MoveNext
        If rsCreditos.EOF Then
            Exit Do
        End If
    Loop
    
    oBarra.Progress 5, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180515
    
    'JUEZ 20131211 *****************************************************
    Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = New ADODB.Recordset
    'EXPOSICIONES DE CONSUMO NO REVOLVENTE
    Set rsCreditos = oCreditos.RecuperaCarteraCreditoNoRevolvPorTramos(TxtCambio.Text, mskPeriodo1Del.Text)


    Do While Not rsCreditos.EOF
        'Convenio descuento por planilla no revolvente
        If rsCreditos!cTpoCredCod = "1" And rsCreditos!cOrden = "1" Then
                xlHoja1.Cells(34, 4) = rsCreditos!K
                xlHoja1.Cells(34, 6) = rsCreditos!id
                xlHoja1.Cells(34, 8) = rsCreditos!P
                xlHoja1.Cells(34, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCredCod = "1" And rsCreditos!cOrden = "2" Then
                xlHoja1.Cells(35, 4) = rsCreditos!K
                xlHoja1.Cells(35, 6) = rsCreditos!id
                xlHoja1.Cells(35, 8) = rsCreditos!P
                xlHoja1.Cells(35, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCredCod = "1" And rsCreditos!cOrden = "3" Then
                xlHoja1.Cells(36, 4) = rsCreditos!K
                xlHoja1.Cells(36, 6) = rsCreditos!id
                xlHoja1.Cells(36, 8) = rsCreditos!P
                xlHoja1.Cells(36, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCredCod = "1" And rsCreditos!cOrden = "4" Then
                xlHoja1.Cells(37, 4) = rsCreditos!K
                xlHoja1.Cells(37, 6) = rsCreditos!id
                xlHoja1.Cells(37, 8) = rsCreditos!P
                xlHoja1.Cells(37, 11) = rsCreditos!PRCC
        End If

        'Otras exposiciones de consumo no revolvente
        If rsCreditos!cTpoCredCod = "2" And rsCreditos!cOrden = "1" Then
                xlHoja1.Cells(44, 4) = rsCreditos!K
                xlHoja1.Cells(44, 6) = rsCreditos!id
                xlHoja1.Cells(44, 8) = rsCreditos!P
                xlHoja1.Cells(44, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCredCod = "2" And rsCreditos!cOrden = "2" Then
                xlHoja1.Cells(45, 4) = rsCreditos!K
                xlHoja1.Cells(45, 6) = rsCreditos!id
                xlHoja1.Cells(45, 8) = rsCreditos!P
                xlHoja1.Cells(45, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCredCod = "2" And rsCreditos!cOrden = "3" Then
                xlHoja1.Cells(46, 4) = rsCreditos!K
                xlHoja1.Cells(46, 6) = rsCreditos!id
                xlHoja1.Cells(46, 8) = rsCreditos!P
                xlHoja1.Cells(46, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCredCod = "2" And rsCreditos!cOrden = "4" Then
                xlHoja1.Cells(47, 4) = rsCreditos!K
                xlHoja1.Cells(47, 6) = rsCreditos!id
                xlHoja1.Cells(47, 8) = rsCreditos!P
                xlHoja1.Cells(47, 11) = rsCreditos!PRCC
        End If
        
        'Pignoraticio 'JUEZ 20160415
        If rsCreditos!cTpoCredCod = "3" And rsCreditos!cOrden = "1" Then
                xlHoja1.Cells(49, 4) = rsCreditos!K
                xlHoja1.Cells(49, 6) = rsCreditos!id
                xlHoja1.Cells(49, 8) = rsCreditos!P
                xlHoja1.Cells(49, 11) = rsCreditos!PRCC
        End If

        rsCreditos.MoveNext
        If rsCreditos.EOF Then
            Exit Do
        End If
    Loop
    
    oBarra.Progress 6, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180515
    
    Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = New ADODB.Recordset
    'CON ATRASO MAYOR A 90 DÍAS
    Set rsCreditos = oCreditos.RecuperaCarteraCreditoPorTramosMayor90Dias(TxtCambio.Text, mskPeriodo1Del.Text)

    Do While Not rsCreditos.EOF
        'Hipotecarios
        If rsCreditos!cTpoCredCod = "8" And rsCreditos!cOrden = "1" Then
            xlHoja1.Cells(53, 4) = rsCreditos!K
            xlHoja1.Cells(53, 6) = rsCreditos!id
            xlHoja1.Cells(53, 8) = rsCreditos!P
            xlHoja1.Cells(53, 11) = rsCreditos!PRCC
        End If

        'PROVISIÓN ESPECIFICA >=20% AL SALDO CAPITAL
        If rsCreditos!cTpoCredCod = "3" And rsCreditos!cOrden = "2" Then
            xlHoja1.Cells(55, 4) = rsCreditos!K
            xlHoja1.Cells(55, 6) = rsCreditos!id
            xlHoja1.Cells(55, 8) = rsCreditos!P
            xlHoja1.Cells(55, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCredCod = "4" And rsCreditos!cOrden = "2" Then
            xlHoja1.Cells(56, 4) = rsCreditos!K
            xlHoja1.Cells(56, 6) = rsCreditos!id
            xlHoja1.Cells(56, 8) = rsCreditos!P
            xlHoja1.Cells(56, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCredCod = "5" And rsCreditos!cOrden = "2" Then
            xlHoja1.Cells(57, 4) = rsCreditos!K
            xlHoja1.Cells(57, 6) = rsCreditos!id
            xlHoja1.Cells(57, 8) = rsCreditos!P
            xlHoja1.Cells(57, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCredCod = "7" And rsCreditos!cOrden = "2" Then
            xlHoja1.Cells(58, 4) = rsCreditos!K
            xlHoja1.Cells(58, 6) = rsCreditos!id
            xlHoja1.Cells(58, 8) = rsCreditos!P
            xlHoja1.Cells(58, 11) = rsCreditos!PRCC
        End If

        'PROVISIÓN ESPECIFICA < 20%AL SALDO CAPITAL
        If rsCreditos!cTpoCredCod = "3" And rsCreditos!cOrden = "3" Then
            xlHoja1.Cells(60, 4) = rsCreditos!K
            xlHoja1.Cells(60, 6) = rsCreditos!id
            xlHoja1.Cells(60, 8) = rsCreditos!P
            xlHoja1.Cells(60, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCredCod = "4" And rsCreditos!cOrden = "3" Then
            xlHoja1.Cells(61, 4) = rsCreditos!K
            xlHoja1.Cells(61, 6) = rsCreditos!id
            xlHoja1.Cells(61, 8) = rsCreditos!P
            xlHoja1.Cells(61, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCredCod = "5" And rsCreditos!cOrden = "3" Then
            xlHoja1.Cells(62, 4) = rsCreditos!K
            xlHoja1.Cells(62, 6) = rsCreditos!id
            xlHoja1.Cells(62, 8) = rsCreditos!P
            xlHoja1.Cells(62, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cTpoCredCod = "7" And rsCreditos!cOrden = "3" Then
            xlHoja1.Cells(63, 4) = rsCreditos!K
            xlHoja1.Cells(63, 6) = rsCreditos!id
            xlHoja1.Cells(63, 8) = rsCreditos!P
            xlHoja1.Cells(63, 11) = rsCreditos!PRCC
        End If

        rsCreditos.MoveNext
        If rsCreditos.EOF Then
            Exit Do
        End If
    Loop
    
    Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = New ADODB.Recordset
    'EXPOSICIONES DE HIPOTECARIOS PARA VIVIENDA ( No Fondo Mivivienda y Techo Propio)
    Set rsCreditos = oCreditos.RecuperaCarteraCreditoHipotecarioPorTramos(TxtCambio.Text, mskPeriodo1Del.Text, False)
    
    Do While Not rsCreditos.EOF
        'por debajo del indicador prudencial
        If rsCreditos!cIndPrud = "1" And rsCreditos!cOrden = "1" Then
            xlHoja1.Cells(77, 4) = rsCreditos!K
            xlHoja1.Cells(77, 6) = rsCreditos!id
            xlHoja1.Cells(77, 8) = rsCreditos!P
            xlHoja1.Cells(77, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cIndPrud = "1" And rsCreditos!cOrden = "2" Then
            xlHoja1.Cells(78, 4) = rsCreditos!K
            xlHoja1.Cells(78, 6) = rsCreditos!id
            xlHoja1.Cells(78, 8) = rsCreditos!P
            xlHoja1.Cells(78, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cIndPrud = "1" And rsCreditos!cOrden = "3" Then
            xlHoja1.Cells(79, 4) = rsCreditos!K
            xlHoja1.Cells(79, 6) = rsCreditos!id
            xlHoja1.Cells(79, 8) = rsCreditos!P
            xlHoja1.Cells(79, 11) = rsCreditos!PRCC
        End If
        'excede el indicador prudencial
        If rsCreditos!cIndPrud = "2" And rsCreditos!cOrden = "1" Then
            xlHoja1.Cells(81, 4) = rsCreditos!K
            xlHoja1.Cells(81, 6) = rsCreditos!id
            xlHoja1.Cells(81, 8) = rsCreditos!P
            xlHoja1.Cells(81, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cIndPrud = "2" And rsCreditos!cOrden = "2" Then
            xlHoja1.Cells(82, 4) = rsCreditos!K
            xlHoja1.Cells(82, 6) = rsCreditos!id
            xlHoja1.Cells(82, 8) = rsCreditos!P
            xlHoja1.Cells(82, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cIndPrud = "2" And rsCreditos!cOrden = "3" Then
            xlHoja1.Cells(83, 4) = rsCreditos!K
            xlHoja1.Cells(83, 6) = rsCreditos!id
            xlHoja1.Cells(83, 8) = rsCreditos!P
            xlHoja1.Cells(83, 11) = rsCreditos!PRCC
        End If
        
        rsCreditos.MoveNext
        If rsCreditos.EOF Then
            Exit Do
        End If
    Loop
    
    Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = New ADODB.Recordset
    'EXPOSICIONES DE HIPOTECARIOS PARA VIVIENDA (Solo Fondo Mivivienda y Techo Propio)
    Set rsCreditos = oCreditos.RecuperaCarteraCreditoHipotecarioPorTramos(TxtCambio.Text, mskPeriodo1Del.Text, True)
    
    Do While Not rsCreditos.EOF
        'por debajo del indicador prudencial
        If rsCreditos!cIndPrud = "1" And rsCreditos!cOrden = "1" Then
            xlHoja1.Cells(92, 4) = rsCreditos!K
            xlHoja1.Cells(92, 6) = rsCreditos!id
            xlHoja1.Cells(92, 8) = rsCreditos!P
            xlHoja1.Cells(92, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cIndPrud = "1" And rsCreditos!cOrden = "2" Then
            xlHoja1.Cells(93, 4) = rsCreditos!K
            xlHoja1.Cells(93, 6) = rsCreditos!id
            xlHoja1.Cells(93, 8) = rsCreditos!P
            xlHoja1.Cells(93, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cIndPrud = "1" And rsCreditos!cOrden = "3" Then
            xlHoja1.Cells(94, 4) = rsCreditos!K
            xlHoja1.Cells(94, 6) = rsCreditos!id
            xlHoja1.Cells(94, 8) = rsCreditos!P
            xlHoja1.Cells(94, 11) = rsCreditos!PRCC
        End If
        'excede el indicador prudencial
        If rsCreditos!cIndPrud = "2" And rsCreditos!cOrden = "1" Then
            xlHoja1.Cells(96, 4) = rsCreditos!K
            xlHoja1.Cells(96, 6) = rsCreditos!id
            xlHoja1.Cells(96, 8) = rsCreditos!P
            xlHoja1.Cells(96, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cIndPrud = "2" And rsCreditos!cOrden = "2" Then
            xlHoja1.Cells(97, 4) = rsCreditos!K
            xlHoja1.Cells(97, 6) = rsCreditos!id
            xlHoja1.Cells(97, 8) = rsCreditos!P
            xlHoja1.Cells(97, 11) = rsCreditos!PRCC
        End If
        If rsCreditos!cIndPrud = "2" And rsCreditos!cOrden = "3" Then
            xlHoja1.Cells(98, 4) = rsCreditos!K
            xlHoja1.Cells(98, 6) = rsCreditos!id
            xlHoja1.Cells(98, 8) = rsCreditos!P
            xlHoja1.Cells(98, 11) = rsCreditos!PRCC
        End If
        
        rsCreditos.MoveNext
        If rsCreditos.EOF Then
            Exit Do
        End If
    Loop
    'END JUEZ **********************************************************
    oBarra.Progress 7, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180515
    'Tercera Hoja ******************************************************
    lsNomHoja = "prov_gen"
    '*******************************************************************
    
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
    
     Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = New ADODB.Recordset
    
    Set rsCreditos = oCreditos.RecuperaProvisionCartera(TxtCambio.Text, mskPeriodo1Del.Text)
    
    Do While Not rsCreditos.EOF
        xlHoja1.Cells(5, 3) = rsCreditos!nProvision
        rsCreditos.MoveNext
        If rsCreditos.EOF Then
            Exit Do
        End If
    Loop
    'Cuarta Hoja ******************************************************
    lsNomHoja = "Exceso"
    '*******************************************************************
    
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
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("270102", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(2, 3) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14090902", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(3, 3) = Abs(nSaldoCtaContMens)
    
    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14091202", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(4, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14091302", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(5, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14090202", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(6, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14090302", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(7, 3) = Abs(nSaldoCtaContMens)

    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14090402", mskPeriodo1Del.Text, "0", 1)
    xlHoja1.Cells(8, 3) = Abs(nSaldoCtaContMens)
    
    oBarra.Progress 8, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180515
    '************NAGL BEGIN Según ERS020-2018***********************
    pdFecha = mskPeriodo1Del.Text
    pnTpoCambio = TxtCambio.Text
    
    lsNomHoja = "DetHip.NFMiVi"
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
    CargaDetalleHipotecariosNFMV xlHoja1.Application, pdFecha, pnTpoCambio
    
    '**********END********************************************
    oBarra.Progress 9, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180515
    
    'Hoja ******************************************************
    lsNomHoja = "Reporte"
    '*******************************************************************
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
    'ALPA 20120328********************************
    Dim oEvalColoc As COMNCredito.NCOMColocEval
    Set oEvalColoc = New COMNCredito.NCOMColocEval
    
    xlHoja1.Cells(5, 3) = "Correspondiente al " & Format(mskPeriodo1Del.Text, "DD") & " DE " & UCase(Format(mskPeriodo1Del.Text, "MMMM")) & " DEL  " & Format(mskPeriodo1Del.Text, "YYYY")
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 1, CDbl(xlHoja1.Range("W191")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 2, CDbl(xlHoja1.Range("W192")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 3, CDbl(xlHoja1.Range("W193")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 4, CDbl(xlHoja1.Range("W194")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 5, CDbl(xlHoja1.Range("W195")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 6, CDbl(xlHoja1.Range("W196")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 7, CDbl(xlHoja1.Range("W197")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 8, CDbl(xlHoja1.Range("W198")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 9, CDbl(xlHoja1.Range("W199")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 10, CDbl(xlHoja1.Range("W200")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 11, CDbl(xlHoja1.Range("W201")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 12, CDbl(xlHoja1.Range("W202")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 13, CDbl(xlHoja1.Range("W203")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 14, CDbl(xlHoja1.Range("W204")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 15, CDbl(xlHoja1.Range("W205")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 16, CDbl(xlHoja1.Range("W206")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 17, CDbl(xlHoja1.Range("W207")))
    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 18, CDbl(xlHoja1.Range("W208")))
    
    '*********************************************
    '******NAGL 20180515***************
    oBarra.Progress 10, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20180509
    oBarra.CloseForm FrmColocEvalRep
    Set oBarra = Nothing
    '******NAGL END********************
    
    xlHoja1.SaveAs App.Path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    
Exit Sub
'GeneraExcelErr:
'    MsgBox Err.Description, vbInformation, "Aviso"
'    Exit Sub
End Sub

Public Sub CargaDetalleHipotecariosNFMV(ByVal xlApplication As Excel.Application, ByVal pdFecha As Date, ByVal pnTpoCambio As Currency)
Dim lnFila As Integer, lnParam As Integer, lnx As Integer
Dim psCtaCod As String
Dim rs As New ADODB.Recordset
Dim oCred As New COMDCredito.DCOMCredito
lnx = 0
lnFila = 15
Set rs = oCred.DetalleCarteraCrediHipotecarioporTramos(pdFecha, pnTpoCambio, False)
If Not (rs.EOF And rs.BOF) Then
    Do While Not rs.EOF
        If psCtaCod <> Trim(rs!cCtaCod) Then
            xlHoja1.Cells(lnFila, 1) = rs!cPersCod
            xlHoja1.Cells(lnFila, 2) = rs!cPersNombre
            xlHoja1.Cells(lnFila, 3) = rs!cCtaCod
            xlHoja1.Cells(lnFila, 4) = Format(rs!dVenc, "mm/dd/yyyy")
            xlHoja1.Cells(lnFila, 5) = Format(rs!K, "#,##0.00")
            xlHoja1.Cells(lnFila, 6) = Format(rs!P, "#,##0.00")
            xlHoja1.Cells(lnFila, 7) = Format(rs!PRCC, "#,##0.00")
            xlHoja1.Cells(lnFila, 8) = Format(rs!id, "#,##0.00")
            
            xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 8)).NumberFormat = "#,##0.00;-#,##0.00"
            lnParam = lnFila
            lnx = 0
        Else
            xlHoja1.Range(xlHoja1.Cells(lnParam, 1), xlHoja1.Cells(lnParam + lnx, 1)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(lnParam, 1), xlHoja1.Cells(lnParam + lnx, 1)).VerticalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lnParam, 2), xlHoja1.Cells(lnParam + lnx, 2)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(lnParam, 2), xlHoja1.Cells(lnParam + lnx, 2)).VerticalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lnParam, 3), xlHoja1.Cells(lnParam + lnx, 3)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(lnParam, 3), xlHoja1.Cells(lnParam + lnx, 3)).VerticalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lnParam, 4), xlHoja1.Cells(lnParam + lnx, 4)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(lnParam, 4), xlHoja1.Cells(lnParam + lnx, 4)).VerticalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lnParam, 5), xlHoja1.Cells(lnParam + lnx, 5)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(lnParam, 5), xlHoja1.Cells(lnParam + lnx, 5)).VerticalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lnParam, 6), xlHoja1.Cells(lnParam + lnx, 6)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(lnParam, 6), xlHoja1.Cells(lnParam + lnx, 6)).VerticalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lnParam, 7), xlHoja1.Cells(lnParam + lnx, 7)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(lnParam, 7), xlHoja1.Cells(lnParam + lnx, 7)).VerticalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lnParam, 8), xlHoja1.Cells(lnParam + lnx, 8)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(lnParam, 8), xlHoja1.Cells(lnParam + lnx, 8)).VerticalAlignment = xlCenter
            
            'xlHoja1.Range(xlHoja1.Cells(lnParam, 13), xlHoja1.Cells(lnParam + lnx, 13)).MergeCells = True
            'xlHoja1.Range(xlHoja1.Cells(lnParam, 13), xlHoja1.Cells(lnParam + lnx, 13)).VerticalAlignment = xlCenter
            'xlHoja1.Range(xlHoja1.Cells(lnParam, 14), xlHoja1.Cells(lnParam + lnx, 14)).MergeCells = True
            'xlHoja1.Range(xlHoja1.Cells(lnParam, 14), xlHoja1.Cells(lnParam + lnx, 14)).VerticalAlignment = xlCenter
            'xlHoja1.Range(xlHoja1.Cells(lnParam, 13), xlHoja1.Cells(lnParam + lnx, 13)).Font.Color = vbRed
            'xlHoja1.Range(xlHoja1.Cells(lnParam, 14), xlHoja1.Cells(lnParam + lnx, 14)).Font.Color = vbBlue
            'xlHoja1.Range(xlHoja1.Cells(lnParam, 13), xlHoja1.Cells(lnParam + lnx, 14)).Font.Bold = True
            
            ExcelCuadro xlHoja1, 1, lnParam, 14, CCur(lnParam + lnx)
        End If
        xlHoja1.Cells(lnFila, 9) = rs!NroGarant
        xlHoja1.Cells(lnFila, 10) = rs!TpoGarant
        xlHoja1.Cells(lnFila, 11) = Format(rs!Cobertura, "#,##0.00")
        xlHoja1.Cells(lnFila, 12) = Format(rs!nMontoGarCob, "#,##0.00")
        xlHoja1.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 12)).NumberFormat = "#,##0.00;-#,##0.00"
        xlHoja1.Cells(lnFila, 13) = rs!cIndPrud
        xlHoja1.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(lnFila, 13)).Font.Color = vbRed
        xlHoja1.Cells(lnFila, 14) = rs!cOrden
        xlHoja1.Range(xlHoja1.Cells(lnFila, 14), xlHoja1.Cells(lnFila, 14)).Font.Color = vbBlue
        xlHoja1.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(lnFila, 14)).Font.Bold = True
        
        ExcelCuadro xlHoja1, 1, lnFila, 14, CCur(lnFila)
        psCtaCod = Trim(rs!cCtaCod)
        lnx = lnx + 1
        lnFila = lnFila + 1
    rs.MoveNext
    Loop
End If
xlHoja1.Range(xlHoja1.Cells(15, 1), xlHoja1.Cells(lnFila - 1, 14)).Font.Name = "Arial"
xlHoja1.Range(xlHoja1.Cells(15, 1), xlHoja1.Cells(lnFila - 1, 14)).Font.Size = 9
xlHoja1.Range(xlHoja1.Cells(15, 1), xlHoja1.Cells(lnFila - 1, 1)).EntireColumn.AutoFit
xlHoja1.Range(xlHoja1.Cells(15, 4), xlHoja1.Cells(lnFila - 1, 4)).EntireColumn.AutoFit
xlHoja1.Range(xlHoja1.Cells(15, 5), xlHoja1.Cells(lnFila - 1, 5)).EntireColumn.AutoFit
xlHoja1.Range(xlHoja1.Cells(15, 6), xlHoja1.Cells(lnFila - 1, 6)).EntireColumn.AutoFit
xlHoja1.Range(xlHoja1.Cells(15, 7), xlHoja1.Cells(lnFila - 1, 7)).EntireColumn.AutoFit
xlHoja1.Range(xlHoja1.Cells(15, 8), xlHoja1.Cells(lnFila - 1, 8)).EntireColumn.AutoFit
xlHoja1.Range(xlHoja1.Cells(15, 9), xlHoja1.Cells(lnFila - 1, 9)).EntireColumn.AutoFit
xlHoja1.Range(xlHoja1.Cells(15, 10), xlHoja1.Cells(lnFila - 1, 10)).EntireColumn.AutoFit
xlHoja1.Range(xlHoja1.Cells(15, 9), xlHoja1.Cells(lnFila - 1, 10)).HorizontalAlignment = xlCenter
xlHoja1.Range(xlHoja1.Cells(15, 11), xlHoja1.Cells(lnFila - 1, 11)).EntireColumn.AutoFit
xlHoja1.Range(xlHoja1.Cells(15, 12), xlHoja1.Cells(lnFila - 1, 12)).EntireColumn.AutoFit
xlHoja1.Range(xlHoja1.Cells(15, 13), xlHoja1.Cells(lnFila - 1, 14)).HorizontalAlignment = xlCenter

End Sub '****NAGL Según ERS020-2018

Public Sub ExcelCuadro(xlHoja1 As Excel.Worksheet, ByVal x1 As Currency, ByVal Y1 As Currency, ByVal x2 As Currency, ByVal Y2 As Currency, Optional lbLineasVert As Boolean = True, Optional lbLineasHoriz As Boolean = False)
xlHoja1.Range(xlHoja1.Cells(Y1, x1), xlHoja1.Cells(Y2, x2)).BorderAround xlContinuous, xlThin
If lbLineasVert Then
   If x2 <> x1 Then
     xlHoja1.Range(xlHoja1.Cells(Y1, x1), xlHoja1.Cells(Y2, x2)).Borders(xlInsideVertical).LineStyle = xlContinuous
   End If
End If
If lbLineasHoriz Then
    If Y1 <> Y2 Then
        xlHoja1.Range(xlHoja1.Cells(Y1, x1), xlHoja1.Cells(Y2, x2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End If
End If
End Sub '****NAGL 20180509

Public Sub ReporteComparativoCartera()
    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja  As String
    Dim lsNombreAgencia As String
    Dim lsCodAgencia As String
    Dim lsMes As String
    Dim lnContador As Integer
    Dim lsArchivo As String
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    
    Dim rsCreditos As ADODB.Recordset
    
    Dim oCreditos As New COMNCredito.NCOMColocEval
    
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Double
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim n As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Double
    Dim nPase As Integer
    Dim dFechaCP As Date
    Dim lsCelda As String
    Dim lsFechaAnterior As String
'On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "CarteraComparativa"
    lsNomHoja = "CarteraComparativa"
    lsArchivo1 = "\spooler\ComparativoCartera" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx" 'LUCV20190502. Modificó *.xls por *.xlsx
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then 'LUCV20190502. Modificó *.xls por *.xlsx
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsArchivo & ".xlsx") 'LUCV20190502. Modificó *.xls por *.xlsx
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
    
    nSaltoContador = 3
     
    xlHoja1.Cells(1, 17) = UCase(Format(mskPeriodo1Del.Text, "MMMM")) & " DEL  " & Format(mskPeriodo1Del.Text, "YYYY")
    lsFechaAnterior = DateAdd("d", -Day(mskPeriodo1Del.Text), mskPeriodo1Del.Text)
    xlHoja1.Cells(1, 10) = UCase(Format(lsFechaAnterior, "MMMM")) & " DEL  " & Format(lsFechaAnterior, "YYYY")
    Set rsCreditos = oCreditos.ObtenerCarteraComparativa(mskPeriodo1Del.Text, TxtCambio.Text)
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
               
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 25)).Borders.LineStyle = 1
                xlHoja1.Cells(nSaltoContador, 1) = rsCreditos!cCtaCod
                xlHoja1.Cells(nSaltoContador, 2) = rsCreditos!cMoneda
                xlHoja1.Cells(nSaltoContador, 3) = rsCreditos!cAgeDescripcion
                xlHoja1.Cells(nSaltoContador, 4) = rsCreditos!Destino
                xlHoja1.Cells(nSaltoContador, 5) = rsCreditos!cPersCod
                xlHoja1.Cells(nSaltoContador, 6) = rsCreditos!cPersNombre
                xlHoja1.Cells(nSaltoContador, 7) = rsCreditos!cCodDoc
                xlHoja1.Cells(nSaltoContador, 8) = rsCreditos!nMontoApr
                xlHoja1.Cells(nSaltoContador, 9) = rsCreditos!cPersNombreInst
                xlHoja1.Cells(nSaltoContador, 10) = rsCreditos!nSaldoCap05
                xlHoja1.Cells(nSaltoContador, 11) = rsCreditos!nProvision05
                xlHoja1.Cells(nSaltoContador, 12) = rsCreditos!nProvisionRCC05
                xlHoja1.Cells(nSaltoContador, 13) = rsCreditos!nProvisionProciclica05
                xlHoja1.Cells(nSaltoContador, 14) = rsCreditos!nGaran05
                xlHoja1.Cells(nSaltoContador, 15) = rsCreditos!cCalGen05
                xlHoja1.Cells(nSaltoContador, 16) = rsCreditos!cTpoCredDescripcion05
                xlHoja1.Cells(nSaltoContador, 17) = rsCreditos!nSaldoCap06
                xlHoja1.Cells(nSaltoContador, 18) = rsCreditos!nProvision06
                xlHoja1.Cells(nSaltoContador, 19) = rsCreditos!nProvisionRCC06
                xlHoja1.Cells(nSaltoContador, 20) = rsCreditos!nProvisionProciclica06
                xlHoja1.Cells(nSaltoContador, 21) = rsCreditos!nGaran06
                xlHoja1.Cells(nSaltoContador, 22) = rsCreditos!cCalGen06
                xlHoja1.Cells(nSaltoContador, 23) = rsCreditos!cTpoCredDescripcion06
                xlHoja1.Cells(nSaltoContador, 24) = rsCreditos!LogCastigado
                xlHoja1.Cells(nSaltoContador, 25) = "'" & rsCreditos!FechaCastigo
                

                nSaltoContador = nSaltoContador + 1
            rsCreditos.MoveNext
            nContTotal = nContTotal + 1
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    
    xlHoja1.SaveAs App.Path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
End Sub

'ALPA 20120118*****************************************************************************
Public Sub Genera_Reporte178214(ByVal pnBitCentral As Boolean, pnAnio As Integer, pnMes As Integer, pnTpoCambio As Currency, psMes As String, Optional pnBandera As Integer = 1)
Dim nCol  As Integer
Dim sCol  As String

Dim lsArchivo   As String
Dim lbLibroOpen As Boolean
Dim n           As Integer
Dim ldFechaRep As Date
 
On Error GoTo ErrImprimeRiesgos
 
'MousePointer = 11
lnTpoCambio = pnTpoCambio
 
If pnBandera = 1 Then
    lsArchivo = App.Path & "\Spooler\Anx03_" & pnAnio & IIf(Len(Trim(pnMes)) = 1, "0" & Trim(str(pnMes)) & gsCodUser, Trim(str(pnMes))) & ".xls"
    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
    If lbLibroOpen Then
        ExcelAddHoja psMes, xlLibro, xlHoja1
        ldFechaRep = DateAdd("m", 1, CDate("01/" & Format(pnMes, "00") & "/" & Format(pnAnio, "0000"))) - 1
        Call GeneraReporteAnexo3Riesgos(ldFechaRep, pnTpoCambio, psMes)
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
        gFunGeneral.CargaArchivo lsArchivo, App.Path & "\Spooler"
    End If
    'MousePointer = 0
    MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "Aviso"
End If

Exit Sub
ErrImprimeRiesgos:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   If lbLibroOpen Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
      lbLibroOpen = False
   End If
   'MousePointer = 0
End Sub

Private Sub GeneraReporteAnexo3Riesgos(ByVal pdFecha As Date, ByVal pnTipCambio As Double, psMes As String)   ' Flujo Crediticio por Tipo de Credito
Dim i As Integer
Dim nFila As Integer
Dim nIni  As Integer
Dim lNegativo As Boolean
Dim sConec As String
Dim lsSQL As String
Dim rsRang As New ADODB.Recordset
Dim lsCodRangINI() As String * 2
Dim lsCodRangFIN() As String * 2
Dim lsCodRango() As String * 2

Dim lsDesRang() As String

Dim nTempoFila(1 To 3) As Integer

Dim lnRangos As Integer
Dim Reg9 As New ADODB.Recordset
Dim lnNroDeudores As Long
Dim lnSaldoMesAntSol As Currency, lnSaldoMesAntDol As Currency
Dim lnSaldoSol As Currency, lnSaldoDol As Currency
Dim lnNumeroDesembNue As Long ' BRGO BASILEA II
Dim lnDesembNueSol As Currency, lnDesembNueDol As Currency
Dim lnDesembRefSol As Currency, lnDesembRefDol As Currency
Dim ldFechaMesAnt As Date
Dim CIIUReg As String
Dim lnTipCambMesAnt As Currency
Dim j As Integer
Dim lnProduc As Integer
'Dim nFil As Integer

Dim matFinMes(2, 4) As Currency
Dim regTemp As New ADODB.Recordset
Dim oConLocal As DConecta
Dim nFilTemp As Integer
Dim nTFilTemp As Integer
Dim nTotalTemp(9) As Currency
Dim nTTotalTemp(9) As Currency
Dim nTmp As Integer
Dim nTemp As Integer
    
   ldFechaMesAnt = DateAdd("d", pdFecha, -1 * Day(pdFecha))
   Dim oTC As New nTipoCambio
   lnTipCambMesAnt = oTC.EmiteTipoCambio(ldFechaMesAnt + 1, TCFijoMes)
    
   CabeceraExcelAnexo3Riesgos pdFecha, psMes
       
   If Not oCon.AbreConexion Then 'Remota(Right(gsCodAge, 2), True, False, "03")
      Exit Sub
   End If

    lsSQL = " select cCodRango, nDesde, nHasta, cDescrip from anxriesgosrango where copecod='770030'"
    Set oConLocal = New DConecta
    oConLocal.AbreConexion
    Set rsRang = oConLocal.CargaRecordSet(lsSQL)
      
    If Not (rsRang.BOF And rsRang.EOF) Then
        rsRang.MoveLast
        ReDim lsCodRangINI(rsRang.RecordCount)
        ReDim lsCodRangFIN(rsRang.RecordCount)
        ReDim lsCodRango(rsRang.RecordCount)
        ReDim lsDesRang(rsRang.RecordCount)
        
        lnRangos = rsRang.RecordCount
        rsRang.MoveFirst
        i = 0
        CIIUReg = "("
        Do While Not rsRang.EOF
                lsDesRang(i) = rsRang!cDescrip
                lsCodRangINI(i) = FillNum(str(rsRang!nDesde), 2, "0")
                lsCodRangFIN(i) = FillNum(str(rsRang!nHasta), 2, "0")
                lsCodRango(i) = FillNum(rsRang!cCodRango, 2, "0")
                
                If lsCodRangINI(i) = lsCodRangFIN(i) Then
                    CIIUReg = CIIUReg & "'" & lsCodRangINI(i) & "',"
                Else
                    For j = lsCodRangINI(i) To lsCodRangFIN(i)
                        CIIUReg = CIIUReg & "'" & FillNum(str(j), 2, "0") & "',"
                    Next
                End If
            i = i + 1
            rsRang.MoveNext
        Loop
        CIIUReg = Left(CIIUReg, Len(CIIUReg) - 1) & ")"
    End If
 
    For i = 0 To lnRangos - 1
        If i = -1 Then
            xlHoja1.Cells(i + 14, 1) = lsDesRang(i)
            If Trim(lsCodRangINI(i)) = Trim(lsCodRangFIN(i)) Then
                xlHoja1.Cells(i + 14, 2) = "'" & Trim(lsCodRangINI(i))
            Else
                xlHoja1.Cells(i + 14, 2) = "'" & Trim(lsCodRangINI(i)) & " a " & Trim(lsCodRangFIN(i))
            End If
            nTFilTemp = i + 14
        Else
            xlHoja1.Cells(i + 14, 1) = lsDesRang(i)

            If Trim(lsCodRangINI(i)) = Trim(lsCodRangFIN(i)) Then
                xlHoja1.Cells(i + 14, 2) = "'" & Trim(lsCodRangINI(i))
            Else
                If Trim(lsCodRangINI(i)) = "01" Or Trim(lsCodRango(i)) = "05" Or Trim(lsCodRangINI(i)) = "27" Or Trim(lsCodRangINI(i)) = "34" Or Trim(lsCodRangINI(i)) = "40" Or (Trim(lsCodRangINI(i)) = "70" And Trim(lsCodRangFIN(i)) = "71") Or Trim(lsCodRangINI(i)) = "95" Then
                    xlHoja1.Cells(i + 14, 2) = "'" & Trim(lsCodRangINI(i)) & " y " & Trim(lsCodRangFIN(i))
                ElseIf Trim(lsCodRangINI(i)) = "36" Then
                    xlHoja1.Cells(i + 14, 2) = "'23, " & Trim(lsCodRangINI(i)) & " y " & Trim(lsCodRangFIN(i))
                Else
                    xlHoja1.Cells(i + 14, 2) = "'" & Trim(lsCodRangINI(i)) & " a " & Trim(lsCodRangFIN(i))
                End If
            End If
        End If
    Next i
    lsSQL = "exec stp_sel_ObtenerActEconomica "
    oCon.CargaRecordSet (lsSQL)
    lsSQL = "exec B2_stp_sel_Anexo3_StockFlujoCrediticioRiesgos '" & Format(pdFecha, "YYYYmmdd") & "','" & Format("01/" & Month(pdFecha) & "/" & Year(pdFecha), "YYYYmmdd") & "'," & lnTpoCambio
    Set Reg9 = oCon.CargaRecordSet(lsSQL)
    xlHoja1.Range("K12:K52").NumberFormat = "0.00%"
    Do While Not Reg9.EOF
        xlHoja1.Cells(Reg9!TipoCIIU, 3) = Reg9!Numero
        xlHoja1.Cells(Reg9!TipoCIIU, 4) = Reg9!SaldoCapSol
        xlHoja1.Cells(Reg9!TipoCIIU, 5) = Reg9!SaldoCapDol
        xlHoja1.Cells(Reg9!TipoCIIU, 6) = Reg9!SaldoCarteTotal
        xlHoja1.Cells(Reg9!TipoCIIU, 7) = Reg9!NumeroDesembNue
        xlHoja1.Cells(Reg9!TipoCIIU, 8) = Reg9!MontoDesembNueSol
        xlHoja1.Cells(Reg9!TipoCIIU, 9) = Reg9!MontoDesembNueDol
        xlHoja1.Cells(Reg9!TipoCIIU, 10) = Reg9!nMoraCalc
        xlHoja1.Range("K" & Reg9!TipoCIIU).NumberFormat = "0.00%"
        xlHoja1.Cells(Reg9!TipoCIIU, 11) = ((Reg9!nMoraCalc / Reg9!SaldoCarteTotal))
    Reg9.MoveNext
    Loop
    
    xlHoja1.Cells(14, 1) = "A. Agricultura, Ganaderia, Caza y Silvicultura"
    xlHoja1.Range("A12:I12").Font.Bold = True
    xlHoja1.Cells(12, 1) = "1. CRÉDITOS CORPORATIVOS, A GRANDES, A MEDIANAS, A PEQUEÑAS Y A MICROEMPRESAS"
    xlHoja1.Range("A46:I46").Font.Bold = True
    xlHoja1.Cells(46, 1) = "2. CREDITOS HIPOTECARIOS PARA VIVIENDA"
    xlHoja1.Range("A48:I48").Font.Bold = True
    xlHoja1.Cells(48, 1) = "3. CREDITOS DE CONSUMO"
    xlHoja1.Cells(52, 1) = "TOTAL"
    xlHoja1.Range("A52").HorizontalAlignment = xlCenter
    gFunContab.ExcelCuadro xlHoja1, 1, 52, 11, 52
    
    xlHoja1.Range("C12:C12").Formula = "=+C14+C15+C16+C17+C28+C29+C30+C34+C35+C36+C37+C40+C41+C42+C43+C44"
    xlHoja1.Range("C17:C17").Formula = "=SUM(C18:C27)"
    xlHoja1.Range("C30:C30").Formula = "=SUM(C31:C33)"
    xlHoja1.Range("C37:C37").Formula = "=SUM(C38:C39)"
    xlHoja1.Range("C52:C52").Formula = "=+C12+C48+C46"

    xlHoja1.Range("D12:D12").Formula = "=+D14+D15+D16+D17+D28+D29+D30+D34+D35+D36+D37+D40+D41+D42+D43+D44"
    xlHoja1.Range("D17:D17").Formula = "=SUM(D18:D27)"
    xlHoja1.Range("D30:D30").Formula = "=SUM(D31:D33)"
    xlHoja1.Range("D37:D37").Formula = "=SUM(D38:D39)"
    xlHoja1.Range("D52:D52").Formula = "=+D12+D48+D46"
    
    xlHoja1.Range("E12:E12").Formula = "=+E14+E15+E16+E17+E28+E29+E30+E34+E35+E36+E37+E40+E41+E42+E43+E44"
    xlHoja1.Range("E17:E17").Formula = "=SUM(E18:E27)"
    xlHoja1.Range("E30:E30").Formula = "=SUM(E31:E33)"
    xlHoja1.Range("E37:E37").Formula = "=SUM(E38:E39)"
    xlHoja1.Range("E52:E52").Formula = "=+E12+E48+E46"
    
    xlHoja1.Range("F12:F12").Formula = "=+E12+D12"
    xlHoja1.Range("F17:F17").Formula = "=SUM(F18:F27)"
    xlHoja1.Range("F30:F30").Formula = "=SUM(F31:F33)"
    xlHoja1.Range("F37:F37").Formula = "=SUM(F38:F39)"
    xlHoja1.Range("F52:F52").Formula = "=+F12+F48+F46"
    
    xlHoja1.Range("G12:G12").Formula = "=+G14+G15+G16+G17+G28+G29+G30+G34+G35+G36+G37+G40+G41+G42+G43+G44"
    xlHoja1.Range("G17:G17").Formula = "=SUM(G18:G27)"
    xlHoja1.Range("G30:G30").Formula = "=SUM(G31:G33)"
    xlHoja1.Range("G37:G37").Formula = "=SUM(G38:G39)"
    xlHoja1.Range("G52:G52").Formula = "=+G12+G48+G46"
    
    xlHoja1.Range("H12:H12").Formula = "=+H14+H15+H16+H17+H28+H29+H30+H34+H35+H36+H37+H40+H41+H42+H43+H44"
    xlHoja1.Range("H17:H17").Formula = "=SUM(H18:H27)"
    xlHoja1.Range("H30:H30").Formula = "=SUM(H31:H33)"
    xlHoja1.Range("H37:H37").Formula = "=SUM(H38:H39)"
    xlHoja1.Range("H52:H52").Formula = "=+H12+H48+H46"
    
    xlHoja1.Range("I12:I12").Formula = "=+I14+I15+I16+I17+I28+I29+I30+I34+I35+I36+I37+I40+I41+I42+I43+I44"
    xlHoja1.Range("I17:I17").Formula = "=SUM(I18:I27)"
    xlHoja1.Range("I30:I30").Formula = "=SUM(I31:I33)"
    xlHoja1.Range("I37:I37").Formula = "=SUM(I38:I39)"
    xlHoja1.Range("I52:I52").Formula = "=+I12+I48+I46"
    
    xlHoja1.Range("J12:J12").Formula = "=+J14+J15+J16+J17+J28+J29+J30+J34+J35+J36+J37+J40+J41+J42+J43+J44"
    xlHoja1.Range("J17:J17").Formula = "=SUM(J18:J27)"
    xlHoja1.Range("J30:J30").Formula = "=SUM(J31:J33)"
    xlHoja1.Range("J37:J37").Formula = "=SUM(J38:J39)"
    xlHoja1.Range("J52:J52").Formula = "=+J12+J48+J46"
    
    xlHoja1.Range("K12:K12").Formula = "=((J12/F12))"
    xlHoja1.Range("K17:K17").Formula = "=((J17/F17))"
    xlHoja1.Range("K30:K30").Formula = "=((J30/F30))"
    xlHoja1.Range("K37:K37").Formula = "=((J37/F37))"
    xlHoja1.Range("K52:K52").Formula = "=((J52/F52))"

    
    
    xlHoja1.Range("A52:J52").Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(52, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A7:K52").BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
   
    xlHoja1.Cells(53, 1) = "Periodicidad Mensual"
    xlHoja1.Cells(55, 1) = "(1) Clasificación industrial uniforme de todas las Actividades económicas. Tercera Revisión. Naciones Unidas. Consignar la actividad económica que genera el mayor valor añadido de la entidad deudora"
    xlHoja1.Cells(56, 1) = "(2) El total de créditos directos debe coincidir con la suma de las cuentas 1401+1403+1404+1405+1406+1407 del Manual de Contabilidad"
    xlHoja1.Cells(57, 1) = "(3) El total de créditos indirectos debe corresponder a la suma de los saldos de las cuentas 7101+7102+7103+7104+7105 del Manual de Contabilidad"
      
    xlHoja1.Range("A55:A57").Font.Bold = True
    xlHoja1.Range("A56:A60").Font.Size = 8
      
    xlHoja1.Range("D59:E59").MergeCells = True
    xlHoja1.Range("I59:J59").MergeCells = True
    xlHoja1.Range("O59:P59").MergeCells = True
     
    xlHoja1.Range("D60:E60").MergeCells = True
    xlHoja1.Range("I60:J60").MergeCells = True
    xlHoja1.Range("O60:P60").MergeCells = True
    
    xlHoja1.Range("I61:J61").MergeCells = True
      
    xlHoja1.Cells(60, 4) = "Gerente General"
    xlHoja1.Cells(60, 9) = "Contador General"
    xlHoja1.Cells(61, 9) = "Matricula Nro"
    xlHoja1.Cells(60, 15) = "Hecho Por"
     
    xlHoja1.Range("A60:T61").Font.Bold = True
    xlHoja1.Range("A60:T61").Font.Size = 8
     
    xlHoja1.Range("D60:P61").HorizontalAlignment = xlCenter
    xlHoja1.Range("D60:P61").HorizontalAlignment = xlCenter
    xlHoja1.Range("B13:B61").HorizontalAlignment = xlCenter
     
    xlHoja1.Range("C11:C52").NumberFormat = "#,###,##0"
    xlHoja1.Range("G11:G52").NumberFormat = "#,###,##0"
   
    If gbBitCentral = True Then
        oConLocal.CierraConexion
    Else
        oCon.CierraConexion
    End If
   'MDISicmact.staMain.Panels(1).Text = ""
   RSClose R
End Sub

Private Sub CabeceraExcelAnexo3Riesgos(ByVal pdFecha As Date, psMes As String)
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.Zoom = 46
    xlHoja1.Cells(1, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    xlHoja1.Cells(4, 1) = "EMPRESA: " & gsNomCmac
    xlHoja1.Cells(5, 11) = "Codigo: " & gsCodCMAC
    xlHoja1.Cells(3, 5) = "STOCK Y FLUJO CREDITICIO POR TIPO DE CREDITO Y SECTOR ECONOMICO"
    xlHoja1.Cells(4, 4) = "Al " & Mid(pdFecha, 1, 2) & " de " & Trim(psMes) & " de " & Year(pdFecha)
    xlHoja1.Cells(5, 4) = "( En Nuevos Soles )"
    xlHoja1.Cells(1, 11) = "ANEXO 3"
      
    xlHoja1.Range("C7:F7").Merge
    xlHoja1.Range("D8:F8").Merge
    xlHoja1.Range("G7:I7").Merge
    xlHoja1.Range("H8:I8").Merge
    
    xlHoja1.Range("A1:S9").HorizontalAlignment = xlHAlignCenter
    
    xlHoja1.Range("A1:A50").ColumnWidth = 40
    xlHoja1.Range("B1:B50").ColumnWidth = 10
    xlHoja1.Range("C1:C50").ColumnWidth = 10
    xlHoja1.Range("D1:D50").ColumnWidth = 13
    xlHoja1.Range("E1:E50").ColumnWidth = 13
    xlHoja1.Range("F1:F50").ColumnWidth = 13
    xlHoja1.Range("G1:G50").ColumnWidth = 18
    xlHoja1.Range("H1:H50").ColumnWidth = 18
    xlHoja1.Range("I1:K50").ColumnWidth = 18
   
    xlHoja1.Range("B7:B10").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    xlHoja1.Range("A7:K10").BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
    xlHoja1.Range("A7:B10").Borders(xlInsideVertical).LineStyle = xlContinuous
    'xlHoja1.Range("C7:I10").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range("C7:K7").Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range("D8:F8").Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range("H8:I8").Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range("D7:K10").Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A7:K10").HorizontalAlignment = xlHAlignCenter
    
    'xlHoja1.Cells(8, 1) = "Tipo de Credito"
    xlHoja1.Cells(8, 2) = "Division"
    xlHoja1.Cells(9, 2) = "CIIU (*)"
    xlHoja1.Cells(7, 3) = "STOCK AL CIERRE DEL MES"
    xlHoja1.Cells(8, 3) = "Numero"
    xlHoja1.Cells(9, 3) = "de"
    xlHoja1.Cells(10, 3) = "Deudores"
    xlHoja1.Cells(8, 4) = "Saldo"
    xlHoja1.Cells(9, 4) = "M.N."
    xlHoja1.Cells(10, 4) = "(Miles de N.S.)"
    xlHoja1.Cells(9, 5) = "M.E."
    xlHoja1.Cells(10, 5) = "(Miles de N.S.)"
    xlHoja1.Cells(9, 6) = "Total"
    xlHoja1.Cells(10, 6) = "(Miles de N.S.)"
   
    
    xlHoja1.Cells(7, 7) = "FLUJO DESEMBOLSADO EN EL MES"
    xlHoja1.Cells(8, 7) = "Número de nuevos"
    xlHoja1.Cells(9, 7) = "créditos"
    xlHoja1.Cells(10, 7) = "desembolsados"
    xlHoja1.Cells(8, 8) = "Monto de nuevos créditos desemb."
    xlHoja1.Cells(9, 8) = "M.N."
    xlHoja1.Cells(10, 8) = "(Miles de N.S.)"
    xlHoja1.Cells(9, 9) = "M.E."
    xlHoja1.Cells(10, 9) = "(Miles de US$)"
    
    xlHoja1.Cells(9, 10) = " Monto "
    xlHoja1.Cells(10, 10) = "Mora"
    
    xlHoja1.Cells(9, 11) = " % "
    xlHoja1.Cells(10, 11) = "Mora"
    
    xlHoja1.Range("D12:F50").NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range("H12:K50").NumberFormat = "#,##0.00;-#,##0.00"
End Sub
'*********************************************************************************************

'ALPA 20120118*****************************************************************************
Public Sub Reporte178215(ByVal pnVPRiesgo As Currency, ByVal pnVPRiesgoAnterior As Currency)
    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja  As String
    Dim lsNombreAgencia As String
    Dim lsCodAgencia As String
    Dim lsMes As String
    Dim lnContador As Integer
    Dim lsArchivo As String
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    
    Dim rsCreditos As ADODB.Recordset
    
    Dim oCreditos As New COMDCredito.DCOMCredito
    Dim oCredCtaCont As COMDContabilidad.DCOMCtaCont
    Dim nSaldoCtaContMens As Currency
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Double
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim n As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Double
    Dim nPase As Integer
    Dim nProvision As Currency
    Dim nTotalCicloEno As Currency
    Dim nTotalCicloEnoxFactor As Currency
    Dim nTotalReqPExRiesCred As Currency
    Dim obOCred As COMNCredito.NCOMColocEval
    Dim pnPatrimonioEfectivo As Currency
    Dim nTotal2b1 As Currency
    
    'Dim nSaldo12 As Currency
    Dim nFactorAjuste As Currency
    Dim nFactorReque As Currency
    Dim nAnio1, nAnio2, nAnio3, nRiesMerc As Currency
    Dim nFactAjusRM As Currency
    Dim nFactRequRM As Currency
    Dim nSaldoPE As Currency
'On Error GoTo GeneraExcelErr
    
    Set oCredCtaCont = New COMDContabilidad.DCOMCtaCont

    nTotal2b1 = (oCredCtaCont.ObtenerCtaContBalanceMensual("1", mskPeriodo1Del.Text, "0", 1) - oCredCtaCont.ObtenerCtaContBalanceMensual("1", mskPeriodo1Del.Text, "0", 2)) * 0.98 * 1

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "AnexosyReportesEstimados"
    'Primera Hoja ******************************************************
    lsNomHoja = "Reporte 4-A1"
    '*******************************************************************
    lsArchivo1 = "\spooler\AnexRepEs_" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsArchivo & ".xls")
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
    
    nSaltoContador = 8
    xlHoja1.Range("B6") = "AL " & Format(mskPeriodo1Del.Text, "DD") & " DE  " & UCase(Format(mskPeriodo1Del.Text, "MMMM")) & " DEL  " & Format(mskPeriodo1Del.Text, "YYYY")
    xlHoja1.Cells(2, 6) = Format(mskPeriodo1Del.Text, "YYYY")
    
    Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = New ADODB.Recordset
    Set rsCreditos = oCreditos.ObtenerCicloEconomico(mskPeriodo1Del.Text, TxtCambio.Text)
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
    xlHoja1.Cells(3, 2) = Format(mskPeriodo1Del.Text, "DD/MM/YYYY")
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
            If rsCreditos!nPos <> 0 Then
                xlHoja1.Cells(rsCreditos!nPos, 4) = rsCreditos!nVal1_1 + rsCreditos!nVal1_2 + rsCreditos!nVal1_3 + rsCreditos!nVal1_4
                nTotalCicloEno = rsCreditos!nVal1_1 + rsCreditos!nVal1_2 + rsCreditos!nVal1_3 + rsCreditos!nVal1_4
                nSaltoContador = nSaltoContador + 1
                nContTotal = nContTotal + 1
            End If
                rsCreditos.MoveNext
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    
    nTotalCicloEnoxFactor = nTotalCicloEno * 0.0125
    
    Set rsCreditos = Nothing
    Set rsCreditos = New ADODB.Recordset
    Set rsCreditos = oCreditos.RecuperaProvisionCartera(TxtCambio.Text, mskPeriodo1Del.Text)
     Do While Not rsCreditos.EOF
        nProvision = rsCreditos!nProvision
        rsCreditos.MoveNext
        If rsCreditos.EOF Then
            Exit Do
        End If
    Loop
    'nTotalReqPExRiesCred = (nTotalCicloEno - (nProvision - nTotalCicloEnoxFactor)) * 0.1
    
    Set rsCreditos = Nothing
    Set rsCreditos = New ADODB.Recordset
    Set oCreditos = New COMDCredito.DCOMCredito
    
    
    Set obOCred = New COMNCredito.NCOMColocEval
        
    Set rsCreditos = obOCred.CargarExposicionesAjustadas2A1(mskPeriodo1Del.Text)
    
    
    If Not RSVacio(rsCreditos) Then
        nTotalReqPExRiesCred = rsCreditos!nValor * 0.1
    End If
    
    Set obOCred = Nothing
    Set rsCreditos = Nothing
    Set rsCreditos = New ADODB.Recordset
    Set obOCred = New COMNCredito.NCOMColocEval
    
    Set rsCreditos = obOCred.CargaBalanceGenxCodigo(mskPeriodo1Del.Text)
    If Not RSVacio(rsCreditos) Then
        pnPatrimonioEfectivo = rsCreditos!nValor1 + rsCreditos!nValor2
    End If
    
    Set obOCred = Nothing
    Set rsCreditos = Nothing
    Set rsCreditos = New ADODB.Recordset
    Set obOCred = New COMNCredito.NCOMColocEval
    '****
     Set rsCreditos = obOCred.ReporteRiesgoCambiario(Format(mskPeriodo1Del.Text, "YYYY"), Format(mskPeriodo1Del.Text, "MM"))
    
    Do While Not rsCreditos.EOF
        DoEvents
        nAnio1 = (rsCreditos!Saldo_Anualizado1_51 + rsCreditos!Saldo_Anualizado1_52_57) - (rsCreditos!Saldo_Anualizado1_41 + rsCreditos!Saldo_Anualizado1_42_49)
               
        nFactAjusRM = rsCreditos!nFacAjuste
        nFactRequRM = (20 - rsCreditos!nFacPonRequer)
        
        nAnio2 = (rsCreditos!Saldo_Anualizado2_51 + rsCreditos!Saldo_Anualizado2_52_57) - (rsCreditos!Saldo_Anualizado2_41 + rsCreditos!Saldo_Anualizado2_42_49)
        
        nAnio3 = (rsCreditos!Saldo_Anualizado3_51 + rsCreditos!Saldo_Anualizado3_52_57) - (rsCreditos!Saldo_Anualizado3_41 + rsCreditos!Saldo_Anualizado3_42_49)
                
        rsCreditos.MoveNext
        
        If rsCreditos.EOF Then
           Exit Do
        End If
    Loop
    
    Set obOCred = Nothing
    Set rsCreditos = Nothing
    Set rsCreditos = New ADODB.Recordset
    Set obOCred = New COMNCredito.NCOMColocEval
    
    nRiesMerc = ((nAnio1 + nAnio2 + nAnio3) * 15 / 300) * nFactAjusRM * nFactRequRM
    
    Set rsCreditos = obOCred.CargaDatosPatrimonio(Format(mskPeriodo1Del.Text, "YYYY"), Format(mskPeriodo1Del.Text, "MM"), 2)
    nSaldo12 = 0
    Do While Not rsCreditos.EOF
        DoEvents
        
        If rsCreditos!cCtaContCod = "1" Then
            nSaldo12 = nSaldo12 + rsCreditos!nSaldoFinImporte
        End If
        If rsCreditos!cCtaContCod = "2" Then
            nSaldo12 = nSaldo12 - rsCreditos!nSaldoFinImporte
        End If
        rsCreditos.MoveNext
        
        If rsCreditos.EOF Then
           Exit Do
        End If
    Loop
    Set obOCred = Nothing
    Set rsCreditos = Nothing
    Set rsCreditos = New ADODB.Recordset
    Set obOCred = New COMNCredito.NCOMColocEval
    
    Set rsCreditos = obOCred.FactorAjusteRiesgoOperac(Format(mskPeriodo1Del.Text, "YYYY"), Format(mskPeriodo1Del.Text, "MM"), 2)
    
    
    
    Do While Not rsCreditos.EOF
        'xlHoja1.Cells(53, 3) = rsCreditos!nFacAjuste
        'xlHoja1.Cells(54, 3) = IIf(IIf(IsNull(rsCreditos!nFacRequerimiento), 0, rsCreditos!nFacRequerimiento) = 0, 0, rsCreditos!nFacRequerimiento)
        nFactorAjuste = rsCreditos!nFacAjuste
        nFactorReque = IIf(IIf(IsNull(rsCreditos!nFacRequerimiento), 0, rsCreditos!nFacRequerimiento) = 0, 0, rsCreditos!nFacRequerimiento)
        rsCreditos.MoveNext
        
        If rsCreditos.EOF Then
           Exit Do
        End If
    Loop
    
    Set obOCred = Nothing
    Set rsCreditos = Nothing
    Set rsCreditos = New ADODB.Recordset
    nSaldo12 = ((IIf(nSaldo12 < 0, nSaldo12 * -1, nSaldo12) * 1) / 10) * nFactorAjuste * (20 - nFactorReque)
    nSaldoPE = Round((nSaldo12 * 0.1 + nRiesMerc * 0.1 + nTotalReqPExRiesCred), 2)
    '****
    'Segunda Hoja ******************************************************
    lsNomHoja = "Reporte 4-B1"
    '*******************************************************************
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
    xlHoja1.Cells(2, 5) = Format(mskPeriodo1Del.Text, "YYYY")
    
    Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = oCreditos.ObtenerRiesgoConcentracionIndividual(TxtCambio.Text)
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
    nContTotal = 11
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
            xlHoja1.Cells(nContTotal, 3) = rsCreditos!cPersNombre
            xlHoja1.Cells(nContTotal, 4) = rsCreditos!nSaldo
            rsCreditos.MoveNext
            nContTotal = nContTotal + 1
            If rsCreditos.EOF Or nContTotal = 31 Then
               Exit Do
            End If
        Loop
    End If
    xlHoja1.Cells(31, 5) = pnVPRiesgoAnterior
    xlHoja1.Cells(31, 8) = Round((nTotalReqPExRiesCred * Round((((xlHoja1.Cells(31, 4) / xlHoja1.Cells(31, 5)) * 100) * 0.01), 2)) / 100, 2) '0.0047
    'Segunda Hoja ******************************************************
    lsNomHoja = "Deudores100"
    '*******************************************************************
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
    rsCreditos.MoveFirst
    
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
    nContTotal = 4
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
            xlHoja1.Cells(nContTotal, 1) = (nContTotal - 3)
            xlHoja1.Cells(nContTotal, 2) = rsCreditos!cPersNombre
            xlHoja1.Cells(nContTotal, 3) = rsCreditos!nSaldo
            xlHoja1.Range(xlHoja1.Cells(nContTotal, 1), xlHoja1.Cells(nContTotal, 3)).Borders.LineStyle = 1
            rsCreditos.MoveNext
            nContTotal = nContTotal + 1
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
   
    
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    
    'Tercera Hoja ******************************************************
    lsNomHoja = "Reporte 4-B2"
    '*******************************************************************
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
    xlHoja1.Cells(2, 5) = Format(mskPeriodo1Del.Text, "YYYY")
    xlHoja1.Range("B5") = "AL " & Format(mskPeriodo1Del.Text, "DD") & " DE  " & UCase(Format(mskPeriodo1Del.Text, "MMMM")) & " DEL  " & Format(mskPeriodo1Del.Text, "YYYY")
    Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = oCreditos.ObtenerRiesgoConcentracionSectorial(mskPeriodo1Del.Text, mskPeriodo1Del.Text, TxtCambio.Text)
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
    nContTotal = 11
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
            xlHoja1.Cells(rsCreditos!nPos, 4) = rsCreditos!SaldoCarteTotal
            rsCreditos.MoveNext
            nContTotal = nContTotal + 1
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    
    xlHoja1.Cells(36, 4) = nTotalReqPExRiesCred
    'Cuarta Hoja *******************************************************
    lsNomHoja = "Reporte 4-B3"
    '*******************************************************************
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    Set rsCreditos = New ADODB.Recordset
    Set oCreditos = New COMDCredito.DCOMCredito
       
    
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
    xlHoja1.Cells(2, 5) = Format(mskPeriodo1Del.Text, "YYYY")
    xlHoja1.Range("B4") = "AL " & Format(mskPeriodo1Del.Text, "DD") & " DE  " & UCase(Format(mskPeriodo1Del.Text, "MMMM")) & " DEL  " & Format(mskPeriodo1Del.Text, "YYYY")
    Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = oCreditos.ObtenerRiesgoConcentracionRegional(TxtCambio.Text)
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
    nContTotal = 11
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
            xlHoja1.Cells(rsCreditos!nPos, 4) = rsCreditos!nSaldo
            rsCreditos.MoveNext
            nContTotal = nContTotal + 1
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    xlHoja1.Cells(19, 4) = nTotalReqPExRiesCred
    
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    'Cuarta Hoja *******************************************************
    lsNomHoja = "Reporte 4-C"
    '*******************************************************************
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
    xlHoja1.Range("B5") = "AL " & Format(mskPeriodo1Del.Text, "DD") & " DE  " & UCase(Format(mskPeriodo1Del.Text, "MMMM")) & " DEL  " & Format(mskPeriodo1Del.Text, "YYYY")
    xlHoja1.Cells(11, 4) = pnVPRiesgo
    xlHoja1.Cells(12, 4) = pnVPRiesgoAnterior
    'Cuarta Hoja *******************************************************
    lsNomHoja = "Reporte 4-D"
    '*******************************************************************
    Set rsCreditos = Nothing
    Set rsCreditos = New ADODB.Recordset
    Set oCreditos = New COMDCredito.DCOMCredito
       
    
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
    
    Set oCreditos = Nothing
    Set rsCreditos = Nothing
    xlHoja1.Range("B4") = "AL " & Format(mskPeriodo1Del.Text, "DD") & " DE  " & UCase(Format(mskPeriodo1Del.Text, "MMMM")) & " DEL  " & Format(mskPeriodo1Del.Text, "YYYY")
    
    Set rsCreditos = New ADODB.Recordset
    Set oCreditos = New COMDCredito.DCOMCredito
    Set rsCreditos = oCreditos.ObtenerReporteRiesgoCambiario(Format(mskPeriodo1Del.Text, "YYYY"), Format(mskPeriodo1Del.Text, "MM"))
    Dim nB, nC, nD, nE As Currency
    Dim nFactor1, nFactor2 As Currency
    Dim nTotalROperacional As Currency

    While Not rsCreditos.EOF
        nB = nB + rsCreditos!Saldo_Anualizado1_51
        nC = nC + rsCreditos!Saldo_Anualizado1_52_57
        nD = nD + rsCreditos!Saldo_Anualizado1_41
        nE = nE + rsCreditos!Saldo_Anualizado1_42_49
               
        nFactor1 = rsCreditos!nFacAjuste
        nFactor2 = 20 - rsCreditos!nFacPonRequer
        
        nB = nB + rsCreditos!Saldo_Anualizado2_51
        nC = nC + rsCreditos!Saldo_Anualizado2_52_57
        nD = nD + rsCreditos!Saldo_Anualizado2_41
        nE = nE + rsCreditos!Saldo_Anualizado2_42_49
                
        nB = nB + rsCreditos!Saldo_Anualizado3_51
        nC = nC + rsCreditos!Saldo_Anualizado3_52_57
        nD = nD + rsCreditos!Saldo_Anualizado3_41
        nE = nE + rsCreditos!Saldo_Anualizado3_42_49
                        
        rsCreditos.MoveNext
    Wend
    
    nTotalROperacional = ((((nB + nC) - (nD + nE)) * 15) / 300) * nFactor1 * nFactor2
    
    
    xlHoja1.Cells(10, 4) = nTotalReqPExRiesCred
    'xlHoja1.Cells(11, 4) = nTotalReqPExRiesCred + nTotal2b1 + nTotalROperacional
    xlHoja1.Cells(23, 4) = pnPatrimonioEfectivo
    xlHoja1.Cells(11, 4) = nSaldoPE
    
    xlHoja1.SaveAs App.Path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    

'Exit Function
'GeneraExcelErr:
'    MsgBox Err.Description, vbInformation, "Aviso"
'    Exit Sub
End Sub

Private Sub txtVPRiesgo_Change()
    If Not IsNumeric(txtVPRiesgo.Text) Then
        txtVPRiesgo.Text = ""
        txtVPRiesgo.SetFocus
    End If
End Sub

Private Sub txtVPRiesgoAnterior_Change()
    If Not IsNumeric(txtVPRiesgoAnterior.Text) Then
        txtVPRiesgoAnterior.Text = ""
        txtVPRiesgoAnterior.SetFocus
    End If
End Sub

'*** PEAC 20170710
'Public Sub Reporte2A1Nuevo()
'    Dim fs As Scripting.FileSystemObject
'    Dim lbExisteHoja As Boolean
'    Dim lsArchivo1 As String
'    Dim lsNomHoja  As String
'    Dim lsNombreAgencia As String
'    Dim lsCodAgencia As String
'    Dim lsMes As String
'    Dim lnContador As Integer
'    Dim lsArchivo As String
'    Dim xlsAplicacion As Excel.Application
'    Dim xlsLibro As Excel.Workbook
'    Dim xlHoja1 As Excel.Worksheet
'
'    Dim rsCreditos As ADODB.Recordset
'
'    Dim oCreditos As New COMDCredito.DCOMCredito
'    Dim oCredCtaCont As COMDContabilidad.DCOMCtaCont
'    Dim nSaldoCtaContMens As Currency
'    Dim sTexto As String
'    Dim sDocFecha As String
'    Dim nSaltoContador As Double
'    Dim sFecha As String
'    Dim sMov As String
'    Dim sDoc As String
'    Dim n As Integer
'    Dim pnLinPage As Integer
'    Dim nMes As Integer
'    Dim nSaldo12 As Currency
'    Dim nContTotal As Double
'    Dim nPase As Integer
'    Dim pnUIT As Currency
'    Dim i As Integer
''On Error GoTo GeneraExcelErr
'
'    Set fs = New Scripting.FileSystemObject
'    Set xlsAplicacion = New Excel.Application
'    lsArchivo = "Reporte_2A1Nuevo"
'    'Primera Hoja ******************************************************
'    lsNomHoja = "dep_efinan"
'    '*******************************************************************
'    lsArchivo1 = "\spooler\Reporte_2A1Nuevo_" & gsCodUser & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
'    If fs.FileExists(App.Path & "\FormatoCarta\" & lsArchivo & ".xls") Then
'        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsArchivo & ".xls")
'    Else
'        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
'        Exit Sub
'    End If
'
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    nSaltoContador = 8
'
'    'nMES = cboMes.ListIndex + 1
'    Set oCreditos = New COMDCredito.DCOMCredito
'    Set rsCreditos = oCreditos.RecuperaBancosCajasClasificacion
'    nPase = 1
'    If (rsCreditos Is Nothing) Then
'        nPase = 0
'    End If
'    xlHoja1.Cells(3, 2) = Format(mskPeriodo1Del.Text, "DD/MM/YYYY")
'    If nPase = 1 Then
'        Do While Not rsCreditos.EOF
'            xlHoja1.Cells(rsCreditos!cBancosCajasCod, 9) = rsCreditos!cBancosCajasClasificacion
'            xlHoja1.Cells(rsCreditos!cBancosCajasCod, 10) = rsCreditos!cBancosCajasClasificacionValor
'            xlHoja1.Cells(rsCreditos!cBancosCajasCod, 11) = rsCreditos!cBancosCajasPorcentaje
'            nSaltoContador = nSaltoContador + 1
'            rsCreditos.MoveNext
'            nContTotal = nContTotal + 1
'            If rsCreditos.EOF Then
'               Exit Do
'            End If
'        Loop
'    End If
'    Set oCreditos = Nothing
'    If nPase = 1 Then
'        rsCreditos.Close
'    End If
'    Set rsCreditos = Nothing
'    '11030102
'    Set oCredCtaCont = New COMDContabilidad.DCOMCtaCont
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030102", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(9, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030103", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(10, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030104", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(11, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030105", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(12, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030106", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(13, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030121", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(14, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030108", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(15, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030129", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(16, 4) = nSaldoCtaContMens
'
'    '03
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030301", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(20, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030303", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(21, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030304", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(22, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030306", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(23, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030308", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(24, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030309", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(25, 4) = nSaldoCtaContMens + oCredCtaCont.ObtenerCtaContBalanceMensual("1108030309", mskPeriodo1Del.Text, "0", 1)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030310", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(26, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030311", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(27, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030312", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(28, 4) = nSaldoCtaContMens + oCredCtaCont.ObtenerCtaContBalanceMensual("1108030312", mskPeriodo1Del.Text, "0", 1)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030305", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(29, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11030313", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(30, 4) = nSaldoCtaContMens
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110304", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(32, 4) = nSaldoCtaContMens + oCredCtaCont.ObtenerCtaContBalanceMensual("11080304", mskPeriodo1Del.Text, "0", 1)
'
'
'    'Segunda Hoja ******************************************************
'    lsNomHoja = "Cta del Balance"
'    '*******************************************************************
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    '---- Soberanas ----
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1102", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(7, 3) = Abs(nSaldoCtaContMens)
'
'    'empresas del sistema financiero nuevo
'    '-------------------------------------------------
'    Set oCreditos = New COMDCredito.DCOMCredito
'    Set rsCreditos = New ADODB.Recordset
'    Dim nMontEntFin As Double
'    Set rsCreditos = oCreditos.RecuperaSaldosCtasEmpreSistFinanc(Format(mskPeriodo1Del.Text, "yyyyMMdd"))
'    i = 18
'    Do While Not rsCreditos.EOF
'        i = i + 1
'
'        xlHoja1.Cells(i, 1) = rsCreditos!cDescrip
'        xlHoja1.Cells(i, 2) = rsCreditos!cCuenta
'        xlHoja1.Cells(i, 3) = Format(rsCreditos!nSaldoFinImporte, "#,#0.00")
'        If rsCreditos!cCuenta = "1103" Then
'            nMontEntFin = nMontEntFin + rsCreditos!nSaldoFinImporte
'        End If
'        xlHoja1.Cells(i, 4) = rsCreditos!cCuenta1
'        xlHoja1.Cells(i, 5) = Format(rsCreditos!nSaldoFinImporte1, "#,#0.00")
'        xlHoja1.Cells(i, 6) = Format(rsCreditos!nSaldoFinImporte + rsCreditos!nSaldoFinImporte1, "#,#0.00")
'
'        rsCreditos.MoveNext
'        If rsCreditos.EOF Then
'            Exit Do
'        End If
'    Loop
'    Set rsCreditos = Nothing
'    '-------------------------------------------------
'    Set rsCreditos = oCreditos.RecuperaSaldosCtasCnt110803(Format(mskPeriodo1Del.Text, "yyyyMMdd"))
'    i = i + 2
'    Dim nTotal As Double
'    nTotal = 0
'    Do While Not rsCreditos.EOF
'        i = i + 1
'
'        xlHoja1.Cells(i, 1) = rsCreditos!cCtaContDesc
'        xlHoja1.Cells(i, 2) = rsCreditos!cCtaContCod
'        xlHoja1.Cells(i, 3) = Format(rsCreditos!nSaldoFinImporte, "#,#0.00")
'        nTotal = nTotal + rsCreditos!nSaldoFinImporte
'
'        rsCreditos.MoveNext
'        If rsCreditos.EOF Then
'            Exit Do
'        End If
'    Loop
'    Set rsCreditos = Nothing
'    '-------------------------------------------------
'    i = i + 1
'    xlHoja1.Cells(i, 2) = "Total"
'    xlHoja1.Cells(i, 3) = Format(nTotal, "#,#0.00")
'    i = i + 1
'    xlHoja1.Cells(i, 2) = "Total Ent. Finan."
'    xlHoja1.Cells(i, 3) = Format(nTotal + nMontEntFin, "#,#0.00")
'
'    'empresas del sistema financiero
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1103", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(12, 3) = nSaldoCtaContMens
''    '*****************
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1108", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(12, 6) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1108030309", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(13, 6) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1108030312", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(14, 6) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11080304", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(15, 6) = nSaldoCtaContMens
''    '*****************
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110301", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(13, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110302", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(14, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110303", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(15, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110304", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(16, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110306", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(17, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14010906070503", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(18, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14010906070510", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(19, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14010906070511", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(20, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1408090503", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(18, 7) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1408090510", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(19, 7) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1408090511", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(20, 7) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("13", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(21, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1302051903", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(22, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1302051905", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(23, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1305181905", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(24, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1308051803", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(25, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1308051805", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(26, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1308051829", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(27, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1302051929", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(28, 3) = nSaldoCtaContMens
''
''    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("170401", mskPeriodo1Del.Text, "0", 1)
''    xlHoja1.Cells(29, 3) = nSaldoCtaContMens
'
'    '-------------------------------------------------
'    Set rsCreditos = oCreditos.RecuperaSaldosCtasCntAccionariales(Format(mskPeriodo1Del.Text, "yyyyMMdd"))
'    i = 6
'    Dim nTotal1 As Double
'    Dim nTotal2 As Double
'    nTotal1 = 0
'    nTotal2 = 0
'    Do While Not rsCreditos.EOF
'        i = i + 1
'
'        xlHoja1.Cells(i, 14) = rsCreditos!cCtaContDesc1
'        xlHoja1.Cells(i, 15) = rsCreditos!cCtaContCod1
'        xlHoja1.Cells(i, 16) = Format(rsCreditos!nSaldoFinImporte1, "#,#0.00")
'        xlHoja1.Cells(i, 17) = Format(rsCreditos!nSaldoFinImporte2, "#,#0.00")
'        xlHoja1.Cells(i, 18) = rsCreditos!cCtaContCod2
'        xlHoja1.Cells(i, 19) = rsCreditos!nSaldoFinImporte1 - rsCreditos!nSaldoFinImporte2
'
'        nTotal1 = nTotal1 + rsCreditos!nSaldoFinImporte1
'        nTotal2 = nTotal2 + rsCreditos!nSaldoFinImporte2
'
'        rsCreditos.MoveNext
'        If rsCreditos.EOF Then
'            Exit Do
'        End If
'    Loop
'    Set rsCreditos = Nothing
'    '-------------------------------------------------
'
'
'    'accionariales
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1101", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(101, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11010103", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(102, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("110701", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(103, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("11070901", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(104, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1106", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(106, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1902", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(107, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1907", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(108, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1505", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(109, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("150701", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(110, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1901", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(111, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1903", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(112, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1906", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(113, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1801", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(114, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1802", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(115, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1803", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(116, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1804", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(117, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1806", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(118, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1807", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(119, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1904", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(120, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1602", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(121, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("7109", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(122, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1507", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(123, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("2903", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(124, 11) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1908", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(125, 11) = Abs(nSaldoCtaContMens)
'
'    'columna de provisiones
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("180902", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(115, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("180903", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(116, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("180904", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(117, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("180906", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(118, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("180907", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(119, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("190409", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(120, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1609", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(121, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1509", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(123, 12) = Abs(nSaldoCtaContMens)
'
'    'COMPROBACION DE CUENTAS DEL ACTIVO
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(149, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("7102", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(150, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("7109", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(151, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("2903", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(152, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1809", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(153, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("190409", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(154, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1609", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(155, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1509", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(156, 12) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("1409", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(157, 12) = Abs(nSaldoCtaContMens)
'
'    'este si va, cambiar posicion
'    'Cartera CF
'    Set oCreditos = New COMDCredito.DCOMCredito
'    Set rsCreditos = New ADODB.Recordset
'    Set rsCreditos = oCreditos.RecuperaCarteraCF(TxtCambio.Text, mskPeriodo1Del.Text)
'
'    Do While Not rsCreditos.EOF
'        'ALPA 20110629*******************************************************
'        If rsCreditos!nTipoCredito = "381" Then
'                xlHoja1.Cells(141, 11) = rsCreditos!nSaldo
'        ElseIf rsCreditos!nTipoCredito = "481" Then
'                xlHoja1.Cells(142, 11) = rsCreditos!nSaldo
'        ElseIf rsCreditos!nTipoCredito = "581" Then
'                xlHoja1.Cells(143, 11) = rsCreditos!nSaldo
'        End If
'        '********************************************************************
'        rsCreditos.MoveNext
'        If rsCreditos.EOF Then
'            Exit Do
'        End If
'    Loop
'
'
''este va cambiar posicion
''    Set oCreditos = Nothing
''    Set rsCreditos = Nothing
''
''    Set oCreditos = New COMDCredito.DCOMCredito
''    pnUIT = oCreditos.ObtnerUIT(CInt(Format(mskPeriodo1Del.Text, "YYYY")))
''
''    Set oCreditos = Nothing
''    Set rsCreditos = New ADODB.Recordset
''    Set oCreditos = New COMDCredito.DCOMCredito
''
''    Set rsCreditos = oCreditos.RecuperaCarteraAutoliquidable(TxtCambio.Text, pnUIT, mskPeriodo1Del.Text)
''
''    Do While Not rsCreditos.EOF
''
''        If rsCreditos!nTipoCredito = "7" Then
''                xlHoja1.Cells(43, 4) = rsCreditos!nSaldo
''        ElseIf rsCreditos!nTipoCredito = "3" Then
''                xlHoja1.Cells(66, 4) = rsCreditos!nSaldo
''        ElseIf rsCreditos!nTipoCredito = "4" Then
''                xlHoja1.Cells(71, 4) = rsCreditos!nSaldo
''        ElseIf rsCreditos!nTipoCredito = "5" Then
''                xlHoja1.Cells(78, 4) = rsCreditos!nSaldo
''        End If
''
''        rsCreditos.MoveNext
''        If rsCreditos.EOF Then
''            Exit Do
''        End If
''    Loop
'
'    'Segunda Hoja ******************************************************
'    lsNomHoja = "Creditos"
'    '*******************************************************************
'
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    xlHoja1.Cells(1, 2) = UCase(fgDameNombreMes(CInt(Mid(mskPeriodo1Del.Text, 4, 2)))) 'JUEZ 20130208
'
'    Set oCreditos = Nothing
'    Set rsCreditos = Nothing
'
'    Set oCreditos = New COMDCredito.DCOMCredito
'    Set rsCreditos = New ADODB.Recordset
'
'    Set rsCreditos = oCreditos.RecuperaCarteraCreditoPorTramos(TxtCambio.Text, mskPeriodo1Del.Text)
'
'    Do While Not rsCreditos.EOF
'
'        '*** PEAC 20170710
'        If rsCreditos!cTpoCred = "2" And rsCreditos!cOrden = "2" Then
'                xlHoja1.Cells(5, 5) = rsCreditos!K
'                xlHoja1.Cells(5, 7) = rsCreditos!Id
'                xlHoja1.Cells(5, 9) = rsCreditos!P
'                xlHoja1.Cells(5, 12) = rsCreditos!PRCC
'        End If
''        If rsCreditos!cTpoCred = "2" And rsCreditos!cOrden = "4" Then
''                xlHoja1.Cells(6, 5) = rsCreditos!K
''                xlHoja1.Cells(6, 7) = rsCreditos!id
''                xlHoja1.Cells(6, 9) = rsCreditos!P
''                xlHoja1.Cells(6, 12) = rsCreditos!PRCC
''        End If
'        '*** FIN PEAC
'
'        If rsCreditos!cTpoCred = "3" And rsCreditos!cOrden = "1" Then
'                xlHoja1.Cells(7, 5) = rsCreditos!K
'                xlHoja1.Cells(7, 7) = rsCreditos!Id
'                xlHoja1.Cells(7, 9) = rsCreditos!P
'                xlHoja1.Cells(7, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCred = "3" And rsCreditos!cOrden = "2" Then
'                xlHoja1.Cells(8, 5) = rsCreditos!K
'                xlHoja1.Cells(8, 7) = rsCreditos!Id
'                xlHoja1.Cells(8, 9) = rsCreditos!P
'                xlHoja1.Cells(8, 12) = rsCreditos!PRCC
'        End If
'
'        If rsCreditos!cTpoCred = "4" And rsCreditos!cOrden = "1" Then
'                xlHoja1.Cells(14, 5) = rsCreditos!K
'                xlHoja1.Cells(14, 7) = rsCreditos!Id
'                xlHoja1.Cells(14, 9) = rsCreditos!P
'                xlHoja1.Cells(14, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCred = "4" And rsCreditos!cOrden = "2" Then
'                xlHoja1.Cells(15, 5) = rsCreditos!K
'                xlHoja1.Cells(15, 7) = rsCreditos!Id
'                xlHoja1.Cells(15, 9) = rsCreditos!P
'                xlHoja1.Cells(15, 12) = rsCreditos!PRCC
'        End If
'
'        If rsCreditos!cTpoCred = "5" And rsCreditos!cOrden = "1" Then
'                xlHoja1.Cells(18, 5) = rsCreditos!K
'                xlHoja1.Cells(18, 7) = rsCreditos!Id
'                xlHoja1.Cells(18, 9) = rsCreditos!P
'                xlHoja1.Cells(18, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCred = "5" And rsCreditos!cOrden = "2" Then
'                xlHoja1.Cells(19, 5) = rsCreditos!K
'                xlHoja1.Cells(19, 7) = rsCreditos!Id
'                xlHoja1.Cells(19, 9) = rsCreditos!P
'                xlHoja1.Cells(19, 12) = rsCreditos!PRCC
'        End If
'
''        If rsCreditos!cTpoCred = "7" And rsCreditos!cOrden = "1" Then
''                xlHoja1.Cells(22, 5) = rsCreditos!K
''                xlHoja1.Cells(22, 7) = rsCreditos!id
''                xlHoja1.Cells(22, 9) = rsCreditos!P
''                xlHoja1.Cells(22, 12) = rsCreditos!PRCC
''        End If
''        If rsCreditos!cTpoCred = "7" And rsCreditos!cOrden = "2" Then
''                xlHoja1.Cells(23, 5) = rsCreditos!K
''                xlHoja1.Cells(23, 7) = rsCreditos!id
''                xlHoja1.Cells(23, 9) = rsCreditos!P
''                xlHoja1.Cells(23, 12) = rsCreditos!PRCC
''        End If
'        'JUEZ 20131211 ************************************************
''        If rsCreditos!cTpoCred = "7" And rsCreditos!cOrden = "5" Then
''                xlHoja1.Cells(28, 5) = rsCreditos!K
''                xlHoja1.Cells(28, 7) = rsCreditos!id
''                xlHoja1.Cells(28, 9) = rsCreditos!P
''                xlHoja1.Cells(28, 12) = rsCreditos!PRCC
''        End If
'
''        If rsCreditos!cTpoCred = "8" And rsCreditos!cOrden = "1" Then
''                xlHoja1.Cells(31, 5) = rsCreditos!K
''                xlHoja1.Cells(31, 7) = rsCreditos!id
''                xlHoja1.Cells(31, 9) = rsCreditos!P
''                xlHoja1.Cells(31, 12) = rsCreditos!PRCC
''        End If
''        If rsCreditos!cTpoCred = "8" And rsCreditos!cOrden = "3" Then
''                xlHoja1.Cells(32, 5) = rsCreditos!K
''                xlHoja1.Cells(32, 7) = rsCreditos!id
''                xlHoja1.Cells(32, 9) = rsCreditos!P
''                xlHoja1.Cells(32, 12) = rsCreditos!PRCC
''        End If
''        If rsCreditos!cTpoCred = "8" And rsCreditos!cOrden = "2" Then
''                xlHoja1.Cells(33, 5) = rsCreditos!K
''                xlHoja1.Cells(33, 7) = rsCreditos!id
''                xlHoja1.Cells(33, 9) = rsCreditos!P
''                xlHoja1.Cells(33, 12) = rsCreditos!PRCC
''        End If
'        'END JUEZ *****************************************************
'
'        nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("27010112", mskPeriodo1Del.Text, "0", 1)
'        xlHoja1.Cells(8, 10) = Abs(nSaldoCtaContMens)
'
'        nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("27010113", mskPeriodo1Del.Text, "0", 1)
'        xlHoja1.Cells(15, 10) = Abs(nSaldoCtaContMens)
'
'        nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("27010102", mskPeriodo1Del.Text, "0", 1)
'        xlHoja1.Cells(19, 10) = Abs(nSaldoCtaContMens)
'
'        rsCreditos.MoveNext
'        If rsCreditos.EOF Then
'            Exit Do
'        End If
'    Loop
'
'    '-------------------------------------------------
'    Set oCreditos = New COMDCredito.DCOMCredito
'    Set rsCreditos = New ADODB.Recordset
'
'    '*** PEAC 230170718 - se incluye las comisiones de carta fianza
'    Set rsCreditos = oCreditos.RecuperaSaldosInteresDiferidosNuevo(TxtCambio.Text, mskPeriodo1Del.Text)
'
'    Do While Not rsCreditos.EOF
'
'        If rsCreditos!cTpoCredCod = "2" Then
'                'xlHoja1.Cells(5, 11) = rsCreditos!conAtraso
'                xlHoja1.Cells(5, 11) = rsCreditos!sinAtraso
'        ElseIf rsCreditos!cTpoCredCod = "3" Then
'                xlHoja1.Cells(7, 11) = rsCreditos!conAtraso
'                xlHoja1.Cells(8, 11) = rsCreditos!sinAtraso
'        ElseIf rsCreditos!cTpoCredCod = "4" Then
'                xlHoja1.Cells(14, 11) = rsCreditos!conAtraso
'                xlHoja1.Cells(15, 11) = rsCreditos!sinAtraso
'        ElseIf rsCreditos!cTpoCredCod = "5" Then
'                xlHoja1.Cells(18, 11) = rsCreditos!conAtraso
'                xlHoja1.Cells(19, 11) = rsCreditos!sinAtraso
'        ElseIf rsCreditos!cTpoCredCod = "7" Then
'                xlHoja1.Cells(22, 11) = rsCreditos!conAtraso
'                xlHoja1.Cells(23, 11) = rsCreditos!sinAtraso
'        ElseIf rsCreditos!cTpoCredCod = "8" Then '*** PEAC 20170717
'                xlHoja1.Cells(31, 11) = rsCreditos!conAtraso
'                xlHoja1.Cells(33, 11) = rsCreditos!sinAtraso
'        End If
'
'        rsCreditos.MoveNext
'        If rsCreditos.EOF Then
'            Exit Do
'        End If
'    Loop
'    '-------------------------------------------------
'
'    '*** PEAC 20170711
'    Set oCreditos = New COMDCredito.DCOMCredito
'    Set rsCreditos = New ADODB.Recordset
'
'    Set rsCreditos = oCreditos.RecuperaSaldosCartaFianzaCartera(TxtCambio.Text, mskPeriodo1Del.Text)
'
'    Do While Not rsCreditos.EOF
'
'        If Left(rsCreditos!cTpoCredCod, 1) = "2" Then
'                xlHoja1.Cells(5, 6) = rsCreditos!nMonto
'        ElseIf Left(rsCreditos!cTpoCredCod, 1) = "3" Then
'                xlHoja1.Cells(8, 6) = rsCreditos!nMonto
'        ElseIf Left(rsCreditos!cTpoCredCod, 1) = "4" Then
'                xlHoja1.Cells(15, 6) = rsCreditos!nMonto
'        ElseIf Left(rsCreditos!cTpoCredCod, 1) = "5" Then
'                xlHoja1.Cells(19, 6) = rsCreditos!nMonto
'        End If
'
'        rsCreditos.MoveNext
'        If rsCreditos.EOF Then
'            Exit Do
'        End If
'    Loop
'    '*** FIN PEAC
'
'
'    'JUEZ 20131211 *****************************************************
'    Set oCreditos = New COMDCredito.DCOMCredito
'    Set rsCreditos = New ADODB.Recordset
'    'EXPOSICIONES DE CONSUMO NO REVOLVENTE
'    Set rsCreditos = oCreditos.RecuperaCarteraCreditoNoRevolvPorTramos(TxtCambio.Text, mskPeriodo1Del.Text)
'
'
'    Do While Not rsCreditos.EOF
'        'Convenio descuento por planilla no revolvente
'        If rsCreditos!cTpoCredCod = "1" And rsCreditos!cOrden = "1" Then
'                xlHoja1.Cells(53, 5) = rsCreditos!K
'                xlHoja1.Cells(53, 7) = rsCreditos!Id
'                xlHoja1.Cells(53, 9) = rsCreditos!P
'                xlHoja1.Cells(53, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCredCod = "1" And rsCreditos!cOrden = "2" Then
'                xlHoja1.Cells(54, 5) = rsCreditos!K
'                xlHoja1.Cells(54, 7) = rsCreditos!Id
'                xlHoja1.Cells(54, 9) = rsCreditos!P
'                xlHoja1.Cells(54, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCredCod = "1" And rsCreditos!cOrden = "3" Then
'                xlHoja1.Cells(55, 5) = rsCreditos!K
'                xlHoja1.Cells(55, 7) = rsCreditos!Id
'                xlHoja1.Cells(55, 9) = rsCreditos!P
'                xlHoja1.Cells(55, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCredCod = "1" And rsCreditos!cOrden = "4" Then
'                xlHoja1.Cells(56, 5) = rsCreditos!K
'                xlHoja1.Cells(56, 7) = rsCreditos!Id
'                xlHoja1.Cells(56, 9) = rsCreditos!P
'                xlHoja1.Cells(56, 12) = rsCreditos!PRCC
'        End If
'
'        'Otras exposiciones de consumo no revolvente
'        If rsCreditos!cTpoCredCod = "2" And rsCreditos!cOrden = "1" Then
'                xlHoja1.Cells(63, 5) = rsCreditos!K
'                xlHoja1.Cells(63, 7) = rsCreditos!Id
'                xlHoja1.Cells(63, 9) = rsCreditos!P
'                xlHoja1.Cells(63, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCredCod = "2" And rsCreditos!cOrden = "2" Then
'                xlHoja1.Cells(64, 5) = rsCreditos!K
'                xlHoja1.Cells(64, 7) = rsCreditos!Id
'                xlHoja1.Cells(64, 9) = rsCreditos!P
'                xlHoja1.Cells(64, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCredCod = "2" And rsCreditos!cOrden = "3" Then
'                xlHoja1.Cells(65, 5) = rsCreditos!K
'                xlHoja1.Cells(65, 7) = rsCreditos!Id
'                xlHoja1.Cells(65, 9) = rsCreditos!P
'                xlHoja1.Cells(65, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCredCod = "2" And rsCreditos!cOrden = "4" Then
'                xlHoja1.Cells(66, 5) = rsCreditos!K
'                xlHoja1.Cells(66, 7) = rsCreditos!Id
'                xlHoja1.Cells(66, 9) = rsCreditos!P
'                xlHoja1.Cells(66, 12) = rsCreditos!PRCC
'        End If
'
'        'Pignoraticio 'JUEZ 20160415
'        If rsCreditos!cTpoCredCod = "3" And rsCreditos!cOrden = "1" Then
'                xlHoja1.Cells(68, 5) = rsCreditos!K
'                xlHoja1.Cells(68, 7) = rsCreditos!Id
'                xlHoja1.Cells(68, 9) = rsCreditos!P
'                xlHoja1.Cells(68, 12) = rsCreditos!PRCC
'        End If
'
'        rsCreditos.MoveNext
'        If rsCreditos.EOF Then
'            Exit Do
'        End If
'    Loop
'
'    Set oCreditos = New COMDCredito.DCOMCredito
'    Set rsCreditos = New ADODB.Recordset
'    'CON ATRASO MAYOR A 90 DÍAS
'    Set rsCreditos = oCreditos.RecuperaCarteraCreditoPorTramosMayor90Dias(TxtCambio.Text, mskPeriodo1Del.Text)
'
'    Do While Not rsCreditos.EOF
'        'Hipotecarios
'        If rsCreditos!cTpoCredCod = "8" And rsCreditos!cOrden = "1" Then
'            xlHoja1.Cells(72, 5) = rsCreditos!K
'            xlHoja1.Cells(72, 7) = rsCreditos!Id
'            xlHoja1.Cells(72, 9) = rsCreditos!P
'            xlHoja1.Cells(72, 12) = rsCreditos!PRCC
'        End If
'
'        'PROVISIÓN ESPECIFICA >=20% AL SALDO CAPITAL
'        If rsCreditos!cTpoCredCod = "3" And rsCreditos!cOrden = "2" Then
'            xlHoja1.Cells(74, 5) = rsCreditos!K
'            xlHoja1.Cells(74, 7) = rsCreditos!Id
'            xlHoja1.Cells(74, 9) = rsCreditos!P
'            xlHoja1.Cells(74, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCredCod = "4" And rsCreditos!cOrden = "2" Then
'            xlHoja1.Cells(75, 5) = rsCreditos!K
'            xlHoja1.Cells(75, 7) = rsCreditos!Id
'            xlHoja1.Cells(75, 9) = rsCreditos!P
'            xlHoja1.Cells(75, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCredCod = "5" And rsCreditos!cOrden = "2" Then
'            xlHoja1.Cells(76, 5) = rsCreditos!K
'            xlHoja1.Cells(76, 7) = rsCreditos!Id
'            xlHoja1.Cells(76, 9) = rsCreditos!P
'            xlHoja1.Cells(76, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCredCod = "7" And rsCreditos!cOrden = "2" Then
'            xlHoja1.Cells(77, 5) = rsCreditos!K
'            xlHoja1.Cells(77, 7) = rsCreditos!Id
'            xlHoja1.Cells(77, 9) = rsCreditos!P
'            xlHoja1.Cells(77, 12) = rsCreditos!PRCC
'        End If
'
'        'PROVISIÓN ESPECIFICA < 20%AL SALDO CAPITAL
'        If rsCreditos!cTpoCredCod = "3" And rsCreditos!cOrden = "3" Then
'            xlHoja1.Cells(79, 5) = rsCreditos!K
'            xlHoja1.Cells(79, 7) = rsCreditos!Id
'            xlHoja1.Cells(79, 9) = rsCreditos!P
'            xlHoja1.Cells(79, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCredCod = "4" And rsCreditos!cOrden = "3" Then
'            xlHoja1.Cells(80, 5) = rsCreditos!K
'            xlHoja1.Cells(80, 7) = rsCreditos!Id
'            xlHoja1.Cells(80, 9) = rsCreditos!P
'            xlHoja1.Cells(80, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCredCod = "5" And rsCreditos!cOrden = "3" Then
'            xlHoja1.Cells(81, 5) = rsCreditos!K
'            xlHoja1.Cells(81, 7) = rsCreditos!Id
'            xlHoja1.Cells(81, 9) = rsCreditos!P
'            xlHoja1.Cells(81, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cTpoCredCod = "7" And rsCreditos!cOrden = "3" Then
'            xlHoja1.Cells(82, 5) = rsCreditos!K
'            xlHoja1.Cells(82, 7) = rsCreditos!Id
'            xlHoja1.Cells(82, 9) = rsCreditos!P
'            xlHoja1.Cells(82, 12) = rsCreditos!PRCC
'        End If
'
'        rsCreditos.MoveNext
'        If rsCreditos.EOF Then
'            Exit Do
'        End If
'    Loop
'
'    Set oCreditos = New COMDCredito.DCOMCredito
'    Set rsCreditos = New ADODB.Recordset
'    'EXPOSICIONES DE HIPOTECARIOS PARA VIVIENDA ( No Fondo Mivivienda y Techo Propio)
'    Set rsCreditos = oCreditos.RecuperaCarteraCreditoHipotecarioPorTramos(TxtCambio.Text, mskPeriodo1Del.Text, False)
'
'    Do While Not rsCreditos.EOF
'        'por debajo del indicador prudencial
'        If rsCreditos!cIndPrud = "1" And rsCreditos!cOrden = "1" Then
'            xlHoja1.Cells(96, 5) = rsCreditos!K
'            xlHoja1.Cells(96, 7) = rsCreditos!Id
'            xlHoja1.Cells(96, 9) = rsCreditos!P
'            xlHoja1.Cells(96, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cIndPrud = "1" And rsCreditos!cOrden = "2" Then
'            xlHoja1.Cells(97, 8) = rsCreditos!K
'            xlHoja1.Cells(97, 7) = rsCreditos!Id
'            xlHoja1.Cells(97, 9) = rsCreditos!P
'            xlHoja1.Cells(97, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cIndPrud = "1" And rsCreditos!cOrden = "3" Then
'            xlHoja1.Cells(98, 5) = rsCreditos!K
'            xlHoja1.Cells(98, 7) = rsCreditos!Id
'            xlHoja1.Cells(98, 9) = rsCreditos!P
'            xlHoja1.Cells(98, 12) = rsCreditos!PRCC
'        End If
'        'excede el indicador prudencial
'        If rsCreditos!cIndPrud = "2" And rsCreditos!cOrden = "1" Then
'            xlHoja1.Cells(100, 5) = rsCreditos!K
'            xlHoja1.Cells(100, 7) = rsCreditos!Id
'            xlHoja1.Cells(100, 9) = rsCreditos!P
'            xlHoja1.Cells(100, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cIndPrud = "2" And rsCreditos!cOrden = "2" Then
'            xlHoja1.Cells(101, 5) = rsCreditos!K
'            xlHoja1.Cells(101, 7) = rsCreditos!Id
'            xlHoja1.Cells(101, 9) = rsCreditos!P
'            xlHoja1.Cells(101, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cIndPrud = "2" And rsCreditos!cOrden = "3" Then
'            xlHoja1.Cells(102, 5) = rsCreditos!K
'            xlHoja1.Cells(102, 7) = rsCreditos!Id
'            xlHoja1.Cells(102, 9) = rsCreditos!P
'            xlHoja1.Cells(102, 12) = rsCreditos!PRCC
'        End If
'
'        rsCreditos.MoveNext
'        If rsCreditos.EOF Then
'            Exit Do
'        End If
'    Loop
'
'    Set oCreditos = New COMDCredito.DCOMCredito
'    Set rsCreditos = New ADODB.Recordset
'    'EXPOSICIONES DE HIPOTECARIOS PARA VIVIENDA (Solo Fondo Mivivienda y Techo Propio)
'    Set rsCreditos = oCreditos.RecuperaCarteraCreditoHipotecarioPorTramos(TxtCambio.Text, mskPeriodo1Del.Text, True)
'
'    Do While Not rsCreditos.EOF
'        'por debajo del indicador prudencial
'        If rsCreditos!cIndPrud = "1" And rsCreditos!cOrden = "1" Then
'            xlHoja1.Cells(111, 5) = rsCreditos!K
'            xlHoja1.Cells(111, 7) = rsCreditos!Id
'            xlHoja1.Cells(111, 9) = rsCreditos!P
'            xlHoja1.Cells(111, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cIndPrud = "1" And rsCreditos!cOrden = "2" Then
'            xlHoja1.Cells(112, 5) = rsCreditos!K
'            xlHoja1.Cells(112, 7) = rsCreditos!Id
'            xlHoja1.Cells(112, 9) = rsCreditos!P
'            xlHoja1.Cells(112, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cIndPrud = "1" And rsCreditos!cOrden = "3" Then
'            xlHoja1.Cells(113, 5) = rsCreditos!K
'            xlHoja1.Cells(113, 7) = rsCreditos!Id
'            xlHoja1.Cells(113, 9) = rsCreditos!P
'            xlHoja1.Cells(113, 12) = rsCreditos!PRCC
'        End If
'        'excede el indicador prudencial
'        If rsCreditos!cIndPrud = "2" And rsCreditos!cOrden = "1" Then
'            xlHoja1.Cells(115, 5) = rsCreditos!K
'            xlHoja1.Cells(115, 7) = rsCreditos!Id
'            xlHoja1.Cells(115, 9) = rsCreditos!P
'            xlHoja1.Cells(115, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cIndPrud = "2" And rsCreditos!cOrden = "2" Then
'            xlHoja1.Cells(116, 5) = rsCreditos!K
'            xlHoja1.Cells(116, 7) = rsCreditos!Id
'            xlHoja1.Cells(116, 9) = rsCreditos!P
'            xlHoja1.Cells(116, 12) = rsCreditos!PRCC
'        End If
'        If rsCreditos!cIndPrud = "2" And rsCreditos!cOrden = "3" Then
'            xlHoja1.Cells(117, 5) = rsCreditos!K
'            xlHoja1.Cells(117, 7) = rsCreditos!Id
'            xlHoja1.Cells(117, 9) = rsCreditos!P
'            xlHoja1.Cells(117, 12) = rsCreditos!PRCC
'        End If
'
'        rsCreditos.MoveNext
'        If rsCreditos.EOF Then
'            Exit Do
'        End If
'    Loop
'    'END JUEZ **********************************************************
'
'    'Tercera Hoja ******************************************************
'    lsNomHoja = "prov_gen"
'    '*******************************************************************
'
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    Set oCreditos = New COMDCredito.DCOMCredito
'    Set rsCreditos = New ADODB.Recordset
'
'    Set rsCreditos = oCreditos.RecuperaProvisionCartera(TxtCambio.Text, mskPeriodo1Del.Text)
'
'    Do While Not rsCreditos.EOF
'        xlHoja1.Cells(5, 3) = rsCreditos!nProvision
'        rsCreditos.MoveNext
'        If rsCreditos.EOF Then
'            Exit Do
'        End If
'    Loop
'    'Cuarta Hoja ******************************************************
'    lsNomHoja = "Exceso"
'    '*******************************************************************
'
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("270102", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(2, 3) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14090902", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(3, 3) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14091202", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(4, 3) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14091302", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(5, 3) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14090202", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(6, 3) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14090302", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(7, 3) = Abs(nSaldoCtaContMens)
'
'    nSaldoCtaContMens = oCredCtaCont.ObtenerCtaContBalanceMensual("14090402", mskPeriodo1Del.Text, "0", 1)
'    xlHoja1.Cells(8, 3) = Abs(nSaldoCtaContMens)
'
'    'Hoja ******************************************************
'    lsNomHoja = "Reporte"
'    '*******************************************************************
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'    'ALPA 20120328********************************
'    Dim oEvalColoc As COMNCredito.NCOMColocEval
'    Set oEvalColoc = New COMNCredito.NCOMColocEval
'
'    xlHoja1.Cells(5, 3) = "Correspondiente al " & Format(mskPeriodo1Del.Text, "DD") & " DE " & UCase(Format(mskPeriodo1Del.Text, "MMMM")) & " DEL  " & Format(mskPeriodo1Del.Text, "YYYY")
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 1, CDbl(xlHoja1.Range("W191")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 2, CDbl(xlHoja1.Range("W192")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 3, CDbl(xlHoja1.Range("W193")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 4, CDbl(xlHoja1.Range("W194")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 5, CDbl(xlHoja1.Range("W195")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 6, CDbl(xlHoja1.Range("W196")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 7, CDbl(xlHoja1.Range("W197")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 8, CDbl(xlHoja1.Range("W198")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 9, CDbl(xlHoja1.Range("W199")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 10, CDbl(xlHoja1.Range("W200")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 11, CDbl(xlHoja1.Range("W201")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 12, CDbl(xlHoja1.Range("W202")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 13, CDbl(xlHoja1.Range("W203")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 14, CDbl(xlHoja1.Range("W204")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 15, CDbl(xlHoja1.Range("W205")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 16, CDbl(xlHoja1.Range("W206")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 17, CDbl(xlHoja1.Range("W207")))
'    Call oEvalColoc.InsertaExposiciones2A1(mskPeriodo1Del.Text, 18, CDbl(xlHoja1.Range("W208")))
'
'    '*********************************************
'    xlHoja1.SaveAs App.Path & lsArchivo1
'    xlsAplicacion.Visible = True
'    xlsAplicacion.Windows(1).Visible = True
'    Set xlsAplicacion = Nothing
'    Set xlsLibro = Nothing
'    Set xlHoja1 = Nothing
'
'
'Exit Sub
''GeneraExcelErr:
''    MsgBox Err.Description, vbInformation, "Aviso"
''    Exit Sub
'End Sub

