VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCredCalendCOFIDE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar Calendario COFIDE"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   Icon            =   "frmCredCalendCOFIDE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   3585
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   10155
      Begin SICMACT.FlexEdit FECalend 
         Height          =   3225
         Left            =   90
         TabIndex        =   15
         Top             =   195
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   5689
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Fecha Venc.-Cuota-Cuotas-Capital-Interes-Int. Gracia-Gastos-Saldo Capital-Cuota + ITF"
         EncabezadosAnchos=   "400-1000-600-1200-1000-1000-1000-1000-1200-1000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-R-R-R-R-R-R-C"
         FormatosEdit    =   "0-0-0-2-3-2-2-2-2-2"
         SelectionMode   =   1
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7740
      Top             =   4740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtruta 
      Enabled         =   0   'False
      Height          =   330
      Left            =   4620
      TabIndex        =   11
      Top             =   135
      Width           =   3945
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4980
      TabIndex        =   10
      ToolTipText     =   "Generar el Calendario de Pagos"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      ToolTipText     =   "Generar el Calendario de Pagos"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "&Cargar Archivo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   8
      ToolTipText     =   "Desembolsos Parciales"
      Top             =   120
      Width           =   1575
   End
   Begin SICMACT.ActXCodCta ActxCta 
      Height          =   420
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   741
      Texto           =   "Credito"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Ruta :"
      Height          =   195
      Left            =   4080
      TabIndex        =   12
      Top             =   210
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Interes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2805
      TabIndex        =   7
      Top             =   4320
      Width           =   630
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6180
      TabIndex        =   6
      Top             =   4320
      Width           =   1410
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5670
      TabIndex        =   5
      Top             =   4320
      Width           =   435
   End
   Begin VB.Label lblInteres 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3480
      TabIndex        =   4
      Top             =   4320
      Width           =   1410
   End
   Begin VB.Label lblCapital 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   795
      TabIndex        =   3
      Top             =   4320
      Width           =   1410
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Capital"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   4320
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total+ITF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7980
      TabIndex        =   1
      Top             =   4320
      Width           =   840
   End
   Begin VB.Label lblTotalCONITF 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   8865
      TabIndex        =   0
      Top             =   4320
      Width           =   1410
   End
End
Attribute VB_Name = "frmCredCalendCOFIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MatFechas As Variant
Dim MatCapital As Variant
Dim MatInteres As Variant
Dim nNroCuotas As Integer
Dim MatCalend As Variant
Dim nSaldoCapital As Double
Dim nMonto As Double
Dim nNroCalen As Integer
Dim nTasaInteres As Double
Dim dDesembolso As Date

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CargarDatosCuenta(ActxCta.NroCuenta)
End If
End Sub

Private Sub cmdAplicar_Click()
Dim oCred As COMNCredito.NCOMCredito
Set oCred = New COMNCredito.NCOMCredito

If MsgBox("Esta seguro de reemplazar los Datos?", vbQuestion + vbYesNo, "Conformación") = vbNo Then Exit Sub

Call oCred.ActualizaCalendarioCofide(MatFechas, MatCapital, MatInteres, nNroCalen, ActxCta.NroCuenta)
MsgBox "Se actualizo el calendario satisfactoriamente", vbInformation, "Mensaje"
Set oCred = Nothing
cmdAplicar.Enabled = False
End Sub

Private Sub cmdCargar_Click()
 ' Establecer CancelError a True
   
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Establecer los indicadores
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.InitDir = App.path
    ' Establecer los filtros
    CommonDialog1.FileName = "*.xls"
    CommonDialog1.Filter = "*.xls" '"Data Rcc (*.ope)|*.ope|"
    ' Especificar el filtro predeterminado
    'cmdlOpen.FilterIndex = 2
    
    ' Presentar el cuadro de diálogo Abrir
    CommonDialog1.ShowOpen
    ' Presentar el nombre del archivo seleccionado
    txtruta = CommonDialog1.FileName
    
    Call CargarArchivoExcelCOFIDE
    Call ReemplazarDatos
    Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    MsgBox Err.Description, vbCritical, "Mensaje"
    Exit Sub

End Sub

Private Sub ReemplazarDatos()
Dim i As Integer
Dim nTotalCapital As Double
Dim nTotalInteres As Double
Dim nTotalcuotasCONItF As Double

nSaldoCapital = nMonto
For i = 0 To nNroCuotas - 1 'UBound(MatFechas) - 1
        FECalend.TextMatrix(i + 1, 1) = Trim(MatFechas(i))
        FECalend.TextMatrix(i + 1, 4) = Format(MatCapital(i), "#0.00")
        FECalend.TextMatrix(i + 1, 5) = Format(MatInteres(i), "#0.00")
        FECalend.TextMatrix(i + 1, 3) = Format(CDbl(FECalend.TextMatrix(i + 1, 4)) + CDbl(FECalend.TextMatrix(i + 1, 5)), "#0.00")
        nSaldoCapital = nSaldoCapital - CDbl(FECalend.TextMatrix(i + 1, 4))
        FECalend.TextMatrix(i + 1, 8) = Format(nSaldoCapital, "#0.00")
        nTotalCapital = nTotalCapital + CDbl(FECalend.TextMatrix(i + 1, 4))
        nTotalInteres = nTotalInteres + CDbl(FECalend.TextMatrix(i + 1, 5))
        FECalend.TextMatrix(i + 1, 9) = Format(FECalend.TextMatrix(i + 1, 3) + fgITFCalculaImpuesto(CDbl(FECalend.TextMatrix(i + 1, 3))), "#0.00")
        nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(i + 1, 9))
Next
    lblCapital.Caption = Format(nTotalCapital, "#0.00")
    lblInteres.Caption = Format(nTotalInteres, "#0.00")
    lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")

cmdAplicar.Enabled = True
End Sub

Private Sub CargarArchivoExcelCOFIDE()
    
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim i As Integer
    Dim nNumReg As Integer
    
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(txtruta.Text, True, True, , "")
    
    Set xlHoja1 = xlLibro.Worksheets("Hoja2")
            
    'Cargar Matrices
    i = 2
    nNumReg = 0
    nSaldoCapital = nMonto
    ReDim MatFechas(0)
    ReDim MatCapital(0)
    ReDim MatInteres(0)
    While xlHoja1.Cells(i, 1) <> ""
        If CStr(xlHoja1.Cells(i, 2)) = "2" Then   'Secuencia=2 (Concesional)
            ReDim Preserve MatFechas(nNumReg + 1)
            ReDim Preserve MatCapital(nNumReg + 1)
            ReDim Preserve MatInteres(nNumReg + 1)
            MatFechas(nNumReg) = CDate(Format(xlHoja1.Cells(i, 4), "mm-dd-yyyy"))
            MatCapital(nNumReg) = Format(xlHoja1.Cells(i, 7), "#0.00")
            If nNumReg = 0 Then
                MatInteres(nNumReg) = Format(MontoIntPerDias(nTasaInteres, MatFechas(nNumReg) - dDesembolso, nSaldoCapital), "#0.00")
            Else
                MatInteres(nNumReg) = Format(MontoIntPerDias(nTasaInteres, MatFechas(nNumReg) - MatFechas(nNumReg - 1), nSaldoCapital), "#0.00")
            End If
            nSaldoCapital = nSaldoCapital - CDbl(MatCapital(nNumReg))
            nNumReg = nNumReg + 1
        End If
        i = i + 1
    Wend
    
    xlLibro.Close
    xlAplicacion.Quit
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing
End Sub

Private Sub CargarDatosCuenta(ByVal psCtaCod As String)
Dim oCred As COMNCredito.NCOMCredito
Dim sMensaje As String
Dim i As Integer
Dim nTotalCapital As Double
Dim nTotalInteres As Double
Dim nTotalcuotasCONItF As Double
Dim nSaldoCapital As Double

Set oCred = New COMNCredito.NCOMCredito
Call oCred.CargarDatosCalendCOFIDE(psCtaCod, nNroCuotas, nMonto, nNroCalen, nTasaInteres, dDesembolso, MatCalend, sMensaje)
Set oCred = Nothing

If sMensaje <> "" Then
    MsgBox sMensaje, vbInformation, "Mensaje"
    Exit Sub
End If
    nSaldoCapital = nMonto
    Call LimpiaFlex(FECalend)
    
    For i = 0 To UBound(MatCalend) - 1
        FECalend.AdicionaFila
        FECalend.TextMatrix(i + 1, 1) = Trim(MatCalend(i, 0))
        FECalend.TextMatrix(i + 1, 2) = Trim(MatCalend(i, 1))
        FECalend.TextMatrix(i + 1, 3) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)), "#0.00")
        FECalend.Row = i + 1
        FECalend.Col = 3
        FECalend.CellForeColor = vbBlue
        FECalend.TextMatrix(i + 1, 4) = Trim(MatCalend(i, 3))
        FECalend.TextMatrix(i + 1, 5) = Trim(MatCalend(i, 4))
        FECalend.TextMatrix(i + 1, 6) = Trim(MatCalend(i, 5))
        FECalend.TextMatrix(i + 1, 7) = Trim(MatCalend(i, 6))
        nSaldoCapital = nSaldoCapital - CDbl(MatCalend(i, 3))
        FECalend.TextMatrix(i + 1, 8) = Format(nSaldoCapital, "#0.00") 'Trim(MatCalend(i, 7))
        nTotalCapital = nTotalCapital + CDbl(Trim(MatCalend(i, 3)))
        nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(i, 4)))
        FECalend.TextMatrix(i + 1, 9) = Format(FECalend.TextMatrix(i + 1, 3) + fgITFCalculaImpuesto(CDbl(FECalend.TextMatrix(i + 1, 3))), "#0.00")
        nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(i + 1, 9))
            
    Next i
    lblCapital.Caption = Format(nTotalCapital, "#0.00")
    lblInteres.Caption = Format(nTotalInteres, "#0.00")
    lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
    lblTotalCONITF.Caption = Format(nTotalcuotasCONItF, "#0.00")
    FECalend.Row = 1
    FECalend.TopRow = 1


End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    ActxCta.Prod = gColHipoMiVivienda
End Sub

'Para el Calculo del Interes sin tener que conectar al componente
Private Function MontoIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer, ByVal pnMonto As Double) As Double
    MontoIntPerDias = (((1 + pnTasaInter / 100) ^ (pnDiasTrans / 30)) - 1) * pnMonto
    MontoIntPerDias = CDbl(Format(MontoIntPerDias, "#0.00"))
End Function
