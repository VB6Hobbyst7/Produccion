VERSION 5.00
Begin VB.Form frmPropuestaDistriUtilidades 
   Caption         =   "Configuracion de Anexo 01"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   Icon            =   "frmPropuestaDistriUtilidades.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin Sicmact.FlexEdit FEParametroUtilidad 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3625
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Nro-Parametro-Valor %-Anio-ParamVar"
      EncabezadosAnchos=   "0-6000-1200-0-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-2-X-X"
      ListaControles  =   "-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-C-C"
      FormatosEdit    =   "-0-0-0-0"
      TextArray0      =   "Nro"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label Label1 
      Caption         =   "Parametros de Generación de Reporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmPropuestaDistriUtilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnSemestre As Integer
Dim lnTipo As Integer
'Dim lcParamUtilidad As String 'Descripcion
Dim lnAnio As Integer
Private Sub Form_Load()
    Call CargarFEParametroUtilidad(lnAnio)
End Sub
Private Sub CargarFEParametroUtilidad(pnAnio As Integer)
Dim I As Integer
Dim panAnio As Integer
panAnio = pnAnio
Dim oDbalanceCont As DbalanceCont
Set oDbalanceCont = New DbalanceCont
Dim rsUtilidad As ADODB.Recordset
Set rsUtilidad = New ADODB.Recordset
Set rsUtilidad = oDbalanceCont.recuperarConfParametroUtilidad(panAnio)
Call LimpiaFlex(FEParametroUtilidad)

If Not rsUtilidad.BOF And Not rsUtilidad.EOF Then
    I = 1
    Do While Not rsUtilidad.EOF
        Me.FEParametroUtilidad.AdicionaFila
        Me.FEParametroUtilidad.TextMatrix(I, 1) = rsUtilidad!cParamUtilidad
        Me.FEParametroUtilidad.TextMatrix(I, 2) = rsUtilidad!nValor
        Me.FEParametroUtilidad.TextMatrix(I, 3) = rsUtilidad!nAnio
        Me.FEParametroUtilidad.TextMatrix(I, 4) = rsUtilidad!nParamVar
        rsUtilidad.MoveNext
        I = I + 1
    Loop
End If

End Sub
Public Sub Inicio(nAnio As Integer, nSemestre As Integer, nTipo As Integer)
    lnAnio = nAnio
    lnSemestre = nSemestre
    lnTipo = nTipo
    Me.Show 1
End Sub
Private Sub cmdCerrar_Click()
Unload Me
End Sub
Private Sub cmdGenerar_Click()
    Dim oDbalanceCont As DbalanceCont
    Dim lnAnioFE As Integer
    Dim lnParamVarFE As Integer
    Dim lnValorFE As Integer
    Dim I As Integer
    Set oDbalanceCont = New DbalanceCont
    If MsgBox("Desea generar el reporte con los parametros Asignados", vbInformation + vbYesNo, "Reporte Anexo 01") = vbYes Then
        For I = 1 To 4 Step 1
            lnAnioFE = Me.FEParametroUtilidad.TextMatrix(I, 3)
            If lnAnioFE = 0 Then
                lnValorFE = Me.FEParametroUtilidad.TextMatrix(I, 2)
                lnParamVarFE = Me.FEParametroUtilidad.TextMatrix(I, 4)
                Call oDbalanceCont.ModificarInsertarConfUtilidad(lnParamVarFE, lnAnio, lnValorFE, lnAnioFE)
            Else
                lnValorFE = Me.FEParametroUtilidad.TextMatrix(I, 2)
                lnParamVarFE = Me.FEParametroUtilidad.TextMatrix(I, 4)
                Call oDbalanceCont.ModificarInsertarConfUtilidad(lnParamVarFE, lnAnioFE, lnValorFE, lnAnioFE)
            End If
        Next
        Call ReporteAnexo01PropuestaDistribucionUtilidades
    End If
End Sub
   Public Sub ReporteAnexo01PropuestaDistribucionUtilidades()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim lsArchivo As String
    Dim rsUtilidad As ADODB.Recordset
    Dim oDbalanceCont As New DbalanceCont
    Dim nPase As Integer
    Dim lbExisteHoja As Boolean
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "Anexo01PropuestaDistribucionUtilidades"
    'Primera Hoja
    lsNomHoja = "Anexo01"
    '************
    lsArchivo1 = "\spooler\RepAnexo01ProDisUtilidad" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
    
    Set rsUtilidad = oDbalanceCont.recuperarReporteAnexo01(CStr(lnAnio), lnSemestre, lnTipo)
    nPase = 1
    If (rsUtilidad Is Nothing) Then
        nPase = 0
    End If
    
    xlHoja1.Range("B3", "N3").HorizontalAlignment = xlCenter
    xlHoja1.Range("M9", "M24").HorizontalAlignment = xlCenter
    xlHoja1.Range("B3", "N3").MergeCells = True
    If lnTipo = 1 Then
    xlHoja1.Cells(3, 2) = "PROPUESTA DE DISTRIBUCION DE UTILIDADES " & CStr(lnAnio) & " (EN SOLES)"
    Else
    xlHoja1.Cells(3, 2) = "PROPUESTA DE DISTRIBUCION DE UTILIDADES " & CStr(lnAnio) & " (EN MILES DE SOLES)"
    End If
    xlHoja1.Range("B3", "N3").Font.Bold = True
    xlHoja1.Range("C9", "D9").Font.Bold = True
    
    xlHoja1.Range("N9", "N10").Font.Bold = True
    xlHoja1.Range("N13", "N14").Font.Bold = True
    xlHoja1.Range("N17", "N18").Font.Bold = True
    xlHoja1.Range("N21", "N22").Font.Bold = True
    
    xlHoja1.Range("B3", "N3").Font.Size = 11
    
    'xlHoja1.Cells(2, 7) = Format(txtfecha.Text, "DD") & " DE " & UCase(Format(txtfecha.Text, "MMMM")) & " DEL  " & Format(txtfecha.Text, "YYYY")
    If nPase = 1 Then
        If Not rsUtilidad.BOF And Not rsUtilidad.EOF Then
                'xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 13)).Borders.LineStyle = 1
                'xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 15)).Borders.LineStyle = 1 'FRHU20131014
                xlHoja1.Cells(9, 14) = Format(rsUtilidad!UtilidadNeta, "#,###")
                xlHoja1.Cells(11, 14) = Format(rsUtilidad!AfecPorAcota, "#,###")
                xlHoja1.Cells(13, 14) = Format(rsUtilidad!UtilNetaEjer, "#,###")
                xlHoja1.Cells(15, 14) = Format(rsUtilidad!ReservaLegal, "#,###")
                xlHoja1.Cells(17, 14) = Format(rsUtilidad!UtilidadReal, "#,###")
                xlHoja1.Cells(19, 14) = Format(rsUtilidad!ReseLegalEspe, "#,###")
                xlHoja1.Cells(21, 14) = Format(rsUtilidad!UtilReLiDispo, "#,###")
                xlHoja1.Cells(23, 14) = Format(rsUtilidad!UtiComproCapi, "#,###")
                xlHoja1.Cells(24, 14) = Format(rsUtilidad!UtiDividenMPM, "#,###")
                
                xlHoja1.Cells(9, 3) = "Utilidad Neta al " & Format(rsUtilidad!Fecha, "dd/mm/yyyy")
                xlHoja1.Cells(15, 3) = "Menos: Reserva Legal (" & rsUtilidad!nValorNeta & "% de la Utilidad Neta, Art. 67 de la ley de Bancos Nª 26702) "
                xlHoja1.Cells(19, 3) = "Menos: Reserva Legal especial (" & rsUtilidad!nValorReal & "% de la utilidad real Art. 4. DS Nª 157-90 EF)"
                
                xlHoja1.Cells(9, 13) = "100%"
                'xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 14), xlHoja1.Cells(nSaltoContador, 15)).NumberFormat = "dd/mm/yyyy"
                xlHoja1.Cells(15, 13) = rsUtilidad!nValorNeta & "%"
                xlHoja1.Cells(19, 13) = rsUtilidad!nValorReal & "%"
                xlHoja1.Cells(23, 13) = rsUtilidad!nValorCapi & "%"
                xlHoja1.Cells(24, 13) = rsUtilidad!nValorMuni & "%"
        End If
    End If
    
    Set oDbalanceCont = Nothing
    If nPase = 1 Then
        rsUtilidad.Close
    End If
    Set rsUtilidad = Nothing
'
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

    Exit Sub

End Sub



















