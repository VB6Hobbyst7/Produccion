VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReporteMonedaFalsa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Moneda Falsificada"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   Icon            =   "frmReporteMonedaFalsa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraRangoFechas 
      Height          =   975
      Left            =   1530
      TabIndex        =   5
      Top             =   15
      Width           =   2160
      Begin MSMask.MaskEdBox txtFechaDesde 
         Height          =   315
         Left            =   900
         TabIndex        =   6
         Top             =   180
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaHasta 
         Height          =   315
         Left            =   915
         TabIndex        =   7
         Top             =   510
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "Desde :"
         Height          =   240
         Left            =   135
         TabIndex        =   9
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label7 
         Caption         =   "Hasta :"
         Height          =   240
         Left            =   135
         TabIndex        =   8
         Top             =   555
         Width           =   600
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1905
      TabIndex        =   4
      Top             =   1095
      Width           =   1125
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   660
      TabIndex        =   3
      Top             =   1095
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   1365
      Begin VB.OptionButton OptTipo 
         Caption         =   "Billetes"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   585
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Monedas"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   1
         Top             =   255
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmReporteMonedaFalsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lsArchivo As String
Dim sql As String
Dim rs As New ADODB.Recordset
Dim fs As New Scripting.FileSystemObject
Dim I As Integer
Dim lsNomHoja As String
Dim lbExisteHoja As Boolean
Dim sMoneda As String
Dim sTipo As String


If Not ValFecha(txtFechaDesde) Then Exit Sub
If Not ValFecha(txtFechaHasta) Then Exit Sub

sMoneda = "1"
sTipo = IIf(OptTipo(0).value = True, "1", "2")

lsArchivo = "RepMonFalsa_" & IIf(Me.OptTipo(0) = True, "Bille", "Mone") & Format(gdFecSis, "mmyyyy") & ".XLS"
Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If
Set xlHoja1 = xlLibro.Worksheets.Add

lsNomHoja = "MonedaFalsa"
 For Each xlHoja1 In xlLibro.Worksheets
 If xlHoja1.Name = lsNomHoja Then
    xlHoja1.Activate
    xlHoja1.Range("A1", "AZ10000") = ""
    lbExisteHoja = True
    Exit For
 End If
 Next
If lbExisteHoja = False Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = lsNomHoja
End If

sql = " Select cAgencia, dFecha, cDenominacion, cSerie, isnull(nCantidad,0) nCantidad"
sql = sql & " from monedafalsa MF"
sql = sql & " Inner Join detmonedafalsa  DMF on MF.cItem = DMF.cItem"
sql = sql & " where datediff(day,dFecha,'" & Format(txtFechaDesde, "YYYY/MM/DD") & "')>=0 and  datediff(day,dFecha,'" & Format(txtFechaHasta, "YYYY/MM/DD") & "')<=0"
sql = sql & " and cMoneda='1' and DMF.cTipo='" & sTipo & "'"
sql = sql & " order by MF.dFecha, cAgencia"

Dim oCon As New DConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sql)

xlHoja1.Application.ActiveWindow.Zoom = 80
xlHoja1.PageSetup.Zoom = 80
xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.Range("A1") = gsNomCmac
xlHoja1.Range("A2").HorizontalAlignment = xlCenter
xlHoja1.Range("A2") = "REMISION DE LAS PRESUNTAS FALSIFICACIONES EN MONEDA " & IIf(sMoneda = "1", "NACIONAL", "DOLARES") & "  N°"
xlHoja1.Range("A1:A2").Font.Bold = True
xlHoja1.Range("A2:D2").MergeCells = True

xlHoja1.Range("A4") = "Ica " & Format(gdFecSis, "dddd, D MMMM YYYY")
xlHoja1.Range("A5") = "Señores"
xlHoja1.Range("A6") = "Banco Central de Reserva del Peru"
xlHoja1.Range("A7") = "Sección Caja"
xlHoja1.Range("A8") = "Presente.-"

xlHoja1.Range("A10") = "De acuerdo con lo reglamentado por esta institución pública mediante circular No"
xlHoja1.Range("A11") = "remitimos el siguiente numerario expresado en Moneda Nacional , que hemos retenido"
xlHoja1.Range("A12") = "bajo la presuncion de ser falso:"

xlHoja1.Range("A14") = IIf(OptTipo(0).value = True, "Billetes", "Monedas")

xlHoja1.Range("A15") = "DENOMINACION"
xlHoja1.Range("B15") = IIf(OptTipo(0).value = True, "SERIE", "CANTIDAD")
xlHoja1.Range("C15") = "LUGAR DE PROCEDENCIA"
xlHoja1.Range("A1").ColumnWidth = 16
xlHoja1.Range("B1").ColumnWidth = 20
xlHoja1.Range("C1").ColumnWidth = 35

xlHoja1.Range("A15:D15").HorizontalAlignment = xlCenter
xlHoja1.Range("A15:D15").Font.Bold = True
xlHoja1.Range("A14").Font.Bold = True

xlHoja1.Range("A15:C15").Interior.ColorIndex = 15
xlHoja1.Range("A15:C15").Interior.Pattern = xlSolid
                                    
I = 15
While Not rs.EOF
    xlHoja1.Cells(I, 1) = rs!cDenominacion
    xlHoja1.Cells(I, 2) = IIf(OptTipo(0).value = True, rs!cSerie, rs!nCantidad)
    'xlHoja1.Cells(i, 3) = NombreAgencia( rs!cAgencia)
    rs.MoveNext
    I = I + 1
Wend

xlHoja1.Cells(I + 1, 1) = "En espera de la calificación, nos suscribimos"
xlHoja1.Cells(I + 2, 1) = "Atentamente."

xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Set rs = Nothing
'CargaArchivo lsArchivo, App.path & "\SPOOLER\"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtFechaDesde = gdFecSis
txtFechaHasta = gdFecSis
CentraForm Me
End Sub


