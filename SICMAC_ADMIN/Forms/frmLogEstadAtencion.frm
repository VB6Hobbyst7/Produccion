VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmLogEstadAtencion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadísitcas "
   ClientHeight    =   4425
   ClientLeft      =   2610
   ClientTop       =   2535
   ClientWidth     =   6375
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6375
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   -60
      Width           =   6135
      Begin VB.CommandButton cmdGenVal 
         Caption         =   "Generar"
         Height          =   315
         Left            =   4380
         TabIndex        =   12
         Top             =   300
         Width           =   1515
      End
      Begin VB.CommandButton cmdGrafico 
         Caption         =   "Grafico"
         Height          =   315
         Left            =   4380
         TabIndex        =   11
         Top             =   660
         Width           =   1515
      End
      Begin VB.TextBox txtFecFin 
         Height          =   315
         Left            =   2880
         TabIndex        =   3
         Top             =   660
         Width           =   1035
      End
      Begin VB.TextBox txtFecIni 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   660
         Width           =   1035
      End
      Begin VB.ComboBox cboDoc 
         Height          =   315
         ItemData        =   "frmLogEstadAtencion.frx":0000
         Left            =   1140
         List            =   "frmLogEstadAtencion.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Left            =   780
         TabIndex        =   5
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Documento"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame fraDatos 
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   1020
      Width           =   6135
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRes 
         Height          =   2835
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5001
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
   End
   Begin VB.Frame fraGrafico 
      Height          =   3315
      Left            =   120
      TabIndex        =   9
      Top             =   1020
      Width           =   6135
      Begin MSChart20Lib.MSChart Graf 
         Height          =   3135
         Left            =   0
         OleObjectBlob   =   "frmLogEstadAtencion.frx":0032
         TabIndex        =   10
         Top             =   90
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmLogEstadAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nMaxValor As Integer
Dim EnGrafico As Boolean
Dim EnDetalle As Boolean
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub cmdGrafico_Click()
Dim i As Integer, n As Integer
If EnGrafico Then
   cmdGrafico.Caption = "Grafico"
   fraDatos.Visible = True
   fraGrafico.Visible = False
   EnGrafico = False
   Exit Sub
End If

EnGrafico = True
cmdGrafico.Caption = "Datos"

n = flxRes.Rows - 1

Graf.RowCount = n
Graf.ColumnCount = 1
 
Graf.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
Graf.Plot.Axis(VtChAxisIdY).ValueScale.Auto = True
Graf.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = nMaxValor + 5

   For i = 1 To n
       Graf.Column = 1
       Graf.row = i
       Graf.Data = flxRes.TextMatrix(i, 2)
       Graf.RowLabel = flxRes.TextMatrix(i, 3)
   Next
   
fraDatos.Visible = False
fraGrafico.Visible = True
End Sub


Private Sub Form_Load()
EnGrafico = False
txtFecIni = DateSerial(Year(Date), Month(Date) + 0, 1)
txtFecFin = Date
flxRes.FocusRect = flexFocusNone
flxRes.SelectionMode = flexSelectionFree
cboDoc.ListIndex = 0
End Sub

Private Sub cmdGenVal_Click()
If Len(txtFecIni) = 10 And Len(txtFecFin) = 10 Then
   GeneraValores
Else
   MsgBox "Falta indicar ..." + Space(10), vbInformation
End If
End Sub

Private Sub cboDoc_Click()
If Len(txtFecIni) = 10 And Len(txtFecFin) = 10 Then
   GeneraValores
End If
End Sub

Private Sub GeneraValores()
Dim rs As New ADODB.Recordset, dFechaIni As String, dFechaFin As String
Dim oCon As DConecta
Set oCon = New DConecta
Dim sFechaIni As String, sFechaFin As String
Dim nDocTpo As Integer, i As Integer

dFechaIni = CDate(txtFecIni)
dFechaFin = CDate(txtFecFin)

nDocTpo = IIf(cboDoc.ListIndex = 0, 90, 97)
sFechaIni = CStr(Year(dFechaIni)) + Format(Month(dFechaIni), "00") + Format(Day(dFechaIni), "00")
sFechaFin = CStr(Year(dFechaFin)) + Format(Month(dFechaFin), "00") + Format(Day(dFechaFin), "00")
cmdGrafico.Enabled = True
oCon.AbreConexion
Set rs = oCon.Ejecutar("paGetEstadAtencion '" & sFechaIni & "','" & sFechaFin & "'," & nDocTpo & " ")

FlexResultado 0
If Not rs.EOF Then
   i = 0
   nMaxValor = 0
   Do While Not rs.EOF
      i = i + 1
      InsRow flxRes, i
      flxRes.TextMatrix(i, 1) = "Atendido en " + Format(rs!Dias, "00") + " dias "
      flxRes.TextMatrix(i, 2) = rs!nro
      If rs!nro > nMaxValor Then nMaxValor = rs!nro
      flxRes.TextMatrix(i, 3) = rs!Dias
      rs.MoveNext
   Loop
End If
fraDatos.Visible = True
fraGrafico.Visible = False
cmdGrafico.Caption = "Grafico"
EnGrafico = False
oCon.CierraConexion
        'ARLO 20160126 ***
        Dim lsPalabras As String
        If (nDocTpo = 90) Then
        lsPalabras = "Orden de Compra"
        ElseIf (nDocTpo = 97) Then
        lsPalabras = "Orden de Servicio"
        End If
        gsOpeCod = LogPistaReporteEstadistico
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Genero el Reporte Estadistico de Atencion del Documento " & lsPalabras & " del " & dFechaIni & " al " & sFechaFin
        Set objPista = Nothing
        '**************
End Sub

Sub FlexResultado(nFilas As Integer)
flxRes.Clear
flxRes.RowHeight(0) = 300
Select Case nFilas
    Case 0
         flxRes.Rows = 2
         flxRes.RowHeight(1) = 10
    Case 1
         flxRes.Rows = 2
         flxRes.RowHeight(1) = 260
    Case Is > 1
         flxRes.Rows = nFilas - 1
         flxRes.RowHeight(1) = 260
End Select
flxRes.ColWidth(0) = 0
flxRes.ColWidth(1) = 4000
flxRes.ColWidth(2) = 1580: flxRes.ColAlignment(2) = 4
flxRes.ColWidth(3) = 0
End Sub

Private Sub txtFecIni_Change()
If Len(txtFecIni) < 10 Then
   FlexResultado 0
End If
End Sub

Private Sub txtFecFin_Change()
If Len(txtFecFin) < 10 Then
   FlexResultado 0
End If
End Sub

Private Sub txtFecIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Len(txtFecIni) = 10 And Len(txtFecFin) = 10 Then
      GeneraValores
   End If
   txtFecFin.SetFocus
End If
End Sub

Private Sub txtFecFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Len(txtFecIni) = 10 And Len(txtFecFin) = 10 Then
      GeneraValores
   End If
End If
End Sub

Private Sub flxRes_DblClick()
flxRes_KeyPress 13
End Sub

Private Sub flxRes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   VerDetalle flxRes.row
End If
If KeyAscii = 27 And EnDetalle Then
   GeneraValores
   EnDetalle = False
End If
End Sub

Sub VerDetalle(k As Integer)
Dim rs As New ADODB.Recordset, dFechaIni As String, dFechaFin As String

Dim oConn As DConecta
Set oConn = New DConecta

Dim sFechaIni As String, sFechaFin As String
Dim nDocTpo As Integer, i As Integer, nDias As Integer

dFechaIni = CDate(txtFecIni)
dFechaFin = CDate(txtFecFin)

nDias = flxRes.TextMatrix(k, 3)

nDocTpo = IIf(cboDoc.ListIndex = 0, 33, 34)
sFechaIni = CStr(Year(dFechaIni)) + Format(Month(dFechaIni), "00") + Format(Day(dFechaIni), "00")
sFechaFin = CStr(Year(dFechaFin)) + Format(Month(dFechaFin), "00") + Format(Day(dFechaFin), "00")

oConn.AbreConexion
Set rs = oConn.Ejecutar("paGetEstadAtencion '" & sFechaIni & "','" & sFechaFin & "'," & nDocTpo & "," & nDias & " ")


FlexDetalle 0, nDias
If Not rs.EOF Then
   i = 0
   nMaxValor = 0
   Do While Not rs.EOF
      i = i + 1
      InsRow flxRes, i
      flxRes.TextMatrix(i, 0) = rs!nMovNro
      flxRes.TextMatrix(i, 1) = rs!cMovDesc
      flxRes.TextMatrix(i, 2) = rs!dDocFecha
      flxRes.TextMatrix(i, 3) = rs!dFechaAte
      rs.MoveNext
   Loop
   EnDetalle = True
End If
cmdGrafico.Enabled = False
oConn.CierraConexion
End Sub

Sub FlexDetalle(nFilas As Integer, nNro As Integer)
flxRes.Clear
flxRes.RowHeight(0) = 300
Select Case nFilas
    Case 0
         flxRes.Rows = 2
         flxRes.RowHeight(1) = 10
    Case 1
         flxRes.Rows = 2
         flxRes.RowHeight(1) = 260
    Case Is > 1
         flxRes.Rows = nFilas - 1
         flxRes.RowHeight(1) = 260
End Select
flxRes.ColWidth(0) = 0
flxRes.ColWidth(1) = 3550: flxRes.TextMatrix(0, 1) = "Detalle de atención en " + CStr(nNro) + " dias"
flxRes.ColWidth(2) = 1000: flxRes.TextMatrix(0, 2) = "Requerim": flxRes.ColAlignment(2) = 4
flxRes.ColWidth(3) = 1000: flxRes.TextMatrix(0, 3) = "Atendido": flxRes.ColAlignment(3) = 4
End Sub

