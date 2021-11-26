VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReporteSustViaticos 
   Caption         =   "A Rendir Cuenta Viaticos"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboArea 
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Text            =   "cboArea"
      Top             =   1200
      Width           =   4215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ComboBox cboAgencia 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Text            =   "cboAgencia"
      Top             =   1920
      Width           =   4215
   End
   Begin VB.ComboBox cboUsuario 
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Text            =   "cboUsuario"
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Frame fraFechaRango 
      Caption         =   "Rango de Fechas"
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
      Height          =   675
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3360
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   300
         Left            =   510
         TabIndex        =   1
         Top             =   255
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
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   255
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   1770
         TabIndex        =   4
         Top             =   315
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   315
         Width           =   240
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Agencia"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblArea 
      Caption         =   "Area"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblUsuario 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   615
   End
End
Attribute VB_Name = "frmReporteSustViaticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim xlHoja1 As Excel.Worksheet
Dim lsNomMesP As String
Dim lsNomMesLL As String
Dim nLin As Long
Dim lsArchivo As String
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cmdAceptar_Click()
Dim lsUsuario As String
Dim lsAgencia As String
Dim lsArea As String
Dim ldFechaIni As String
Dim ldFechaFin As String
Dim nU As Integer

ldFechaIni = txtFechaDel.Text
ldFechaFin = txtFechaAl.Text

If ldFechaIni > ldFechaFin Then
    MsgBox "Fecha final debe ser mayor", vbOKOnly, "Error"
    Exit Sub
End If

If cboUsuario.ListIndex >= 0 Then
   If Left(cboUsuario, 4) = "XXXX" Then
      lsUsuario = Left(Me.cboUsuario.List(0), 4)
      For nU = 1 To cboUsuario.ListCount - 1
         lsUsuario = lsUsuario & "','" & Left(Me.cboUsuario.List(nU), 4)
      Next
   Else
      lsUsuario = Left(Me.cboUsuario, 4)
   End If
End If

If cboAgencia.ListIndex >= 0 Then
   If Left(cboAgencia, 4) = "XXXX" Then
      lsAgencia = Left(Me.cboAgencia.List(0), 2)
      For nU = 1 To cboAgencia.ListCount - 1
         lsAgencia = lsAgencia & "','" & Left(Me.cboAgencia.List(nU), 2)
      Next
   Else
      lsAgencia = Left(Me.cboAgencia, 2)
   End If
End If

If cboArea.ListIndex >= 0 Then
   If Left(cboArea, 4) = "XXXX" Then
      lsArea = Left(Me.cboArea.List(0), 3)
      For nU = 1 To cboArea.ListCount - 1
         lsArea = lsArea & "','" & Left(Me.cboArea.List(nU), 3)
      Next
   Else
      lsArea = Left(Me.cboArea, 3)
   End If
End If

lsArchivo = App.path & "\SPOOLER\RSARViaticos_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLSX"

lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
If Not lbLibroOpen Then
    Exit Sub
End If

If cboUsuario.ListIndex >= 0 Or cboAgencia.ListIndex >= 0 Or cboAgencia.ListIndex >= 0 Then
   GeneraReporteSustentacionViaticos lsUsuario, lsAgencia, lsArea, ldFechaIni, ldFechaFin
Else
   MsgBox "Seleccione un Usuario/Agencia/Area", vbInformation, "¡Aviso!"
   cboUsuario.SetFocus
End If
 
ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
CargaArchivo lsArchivo, App.path & "\SPOOLER\"

                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", Me.Caption & " Se Genero Excel "
                Set objPista = Nothing
                '****
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsUser As ADODB.Recordset
    Dim rsAgency As ADODB.Recordset
    Dim rsArea As ADODB.Recordset
    Dim oCon As DConecta
    Dim sSql As String
    Set oCon = New DConecta
    oCon.AbreConexion

    sSql = "Select  rh.cUser, P.cPersNombre" _
         & " From rrhh rh JOIN Persona P ON P.cPersCod = RH.cPersCod" _
         & " Where (rh.cuser not in ('XXXX','')  and rh.nRHEstado='201') or" _
         & " (rh.nRHEstado in ('803','802','801')" _
         & " and dcese between getdate()-7 and getdate()+7)" _
         & " Order By cUser"

    Set rsUser = oCon.CargaRecordSet(sSql)
    
    RSLlenaCombo rsUser, cboUsuario, , , False
    cboUsuario.AddItem "XXXX  TODOS LOS USUARIOS"
    RSClose rsUser
    
    sSql = "Select cAreaCod, cAreaDescripcion From Areas"
    Set rsArea = oCon.CargaRecordSet(sSql)
    
    RSLlenaCombo rsArea, cboArea, , , False
    cboArea.AddItem "XXXX  TODAS LAS AREAS"
    RSClose rsArea

    sSql = "Select cAgecod, cAgeDescripcion From Agencias"
    Set rsAgency = oCon.CargaRecordSet(sSql)
    
    RSLlenaCombo rsAgency, cboAgencia, , , False
    cboAgencia.AddItem "XXXX  TODAS LOS AGENCIAS"
    RSClose rsAgency
    
End Sub

Function GeneraReporteSustentacionViaticos(psUsuario As String, psAgencia As String, _
                                           psArea As String, pdFechaIni As String, pdFechaFin As String) As Integer
    Dim rsMay As New ADODB.Recordset
    Dim oOperacion As DOperacion
    Set oOperacion = New DOperacion
    Dim n As Integer
    Dim lsCtaArendir As String
    Dim lsCtaPendiente As String
    Dim lsUsuario As String
    Dim lsFechaER As String
    Dim lsFechaRen As String
    Dim lnImporte As Double
    Dim lsArea    As String
    Dim lsLugar   As String
    Dim lsPeriodo As String
    Dim lsMotivo  As String
    Dim lnTotDev As Double
    Dim lnTotRend As Double
    Dim lnTotal As Double
   
    Dim sImpre      As String
    Dim sImpreAge   As String
    Dim lsRepTitulo As String
    Dim lsHoja      As String
    
    Dim nPosIni As Long
    Dim nPosFin As Long
    
    Dim w As Long
      
    'lsRepTitulo = "REPORTE DE SUSTENTACION DE A RENDIR CUENTAS VIATICOS"
    lsCtaArendir = oOperacion.EmiteOpeCta("401250", "H", "0")
    lsCtaPendiente = oOperacion.EmiteOpeCta("401240", "H", "1")
    
    Dim oARend As New NARendir
    
    Set rsMay = oARend.GetAtencionPendArendirViaticos(psUsuario, psAgencia, psArea, lsCtaArendir, 25, pdFechaIni, pdFechaFin)
    
    nLin = 1
    lsHoja = "SustARendirViaticos"
    ExcelAddHoja lsHoja, xlLibro, xlHoja1
    'Formato texto a una Columna
    'xlHoja1.Range(xlHoja1.Cells(1, 6), xlHoja1.Cells(700, 6)).NumberFormat = "@"
        
    Do While Not rsMay.EOF
          'N = N + 1
          'w = w + 1
          Dim rsDocRend As ADODB.Recordset
          Dim lnMovNro As Long
          
          lsUsuario = rsMay!cPersNombre
          lsFechaER = rsMay!dDocFecha
          lnImporte = CStr(rsMay!MONTOATENDIDO)
          lsArea = rsMay!cAreaDescripcion
          lsLugar = rsMay!cDestinoDesc
          lsFechaRen = rsMay!FechaRend
          'lsPeriodo = "Del " + rsMay!dPartida + " Al " + rsMay!dLlegada
          
          lsNomMesP = Devuelvemes(Month(rsMay!dPartida))
          lsNomMesLL = Devuelvemes(Month(rsMay!dllegada))
          
          If Year(rsMay!dPartida) = Year(rsMay!dllegada) Then
            If Month(rsMay!dPartida) = Month(rsMay!dllegada) Then
                lsPeriodo = CStr(Day(rsMay!dPartida)) & " al " & CStr(Day(rsMay!dllegada)) & " de " & lsNomMesLL & " del " & CStr(Year(rsMay!dllegada))
            Else
                lsPeriodo = CStr(Day(rsMay!dPartida)) & " de " & lsNomMesP & " al " & CStr(Day(rsMay!dllegada)) & " de " & lsNomMesLL & " del " & CStr(Year(rsMay!dllegada))
            End If
          Else
            lsPeriodo = CStr(Day(rsMay!dPartida)) & " de " & lsNomMesP & " del " & Year(rsMay!dPartida) & " al " & CStr(Day(rsMay!dllegada)) & " de " & lsNomMesLL & " del " & CStr(Year(rsMay!dllegada))
          End If
          lsMotivo = rsMay!cMovDesc
          lnMovNro = rsMay!nMovNro
          
          Set rsDocRend = oARend.GetDocSustentariosArendirViaticos(lnMovNro, lsCtaArendir, lsCtaPendiente)
          
          If rsDocRend.RecordCount <> 0 Then
          '   If Not rsDocRend.EOF Then
          '      rsDocRend.MoveNext
          '   End If
          '   Loop
          'End If
          'If rsMay.Bookmark = 1 Then
          ImprimeSustARendirViaticosExcelCab lsUsuario, lsFechaER, lnImporte, lsArea, _
                                             lsLugar, lsPeriodo, lsMotivo, pdFechaIni, _
                                             pdFechaFin, nLin, lsFechaRen
                                             
          xlHoja1.Range("A" & 0 + nLin & ":H" & 0 + nLin).Font.Bold = True
          'xlHoja1.Range("A" & 4 + nPosFin & ":Z" & 7 + nPosFin).Font.Bold = True
          'nPosIni = 8
          'End If
          nLin = nLin + 1
          
          
          Dim nVal As Double
          Dim nItem As Integer
          Dim lsFechaSust As String
          Dim lsDocAbr As String
          Dim lsDocNro As String
          Dim lsProvSust As String
          Dim lsDetalleSust As String
          Dim lnImporteSust As Double

          nItem = 1
          nVal = 0
          Do While Not rsDocRend.EOF
             lsFechaSust = CDate(rsDocRend!dDocFecha)
             lsDocAbr = rsDocRend!cDocAbrev
             lsDocNro = rsDocRend!cDocNro
             lsProvSust = rsDocRend!cPersNombre
             lsDetalleSust = rsDocRend!cMovDesc
             lnImporteSust = rsDocRend!nDocImporte

             ImprimeDetalleArendirViaticos nItem, lsFechaSust, lsDocAbr, lsDocNro, lsProvSust, lsDetalleSust, lnImporteSust
             nVal = nVal + lnImporteSust
             nItem = nItem + 1
             rsDocRend.MoveNext
             If rsDocRend.EOF Then
                Exit Do
             End If
             
          Loop
          
          'xlHoja1.Range("H" & nLin & ":H" & nLin).BorderAround xlContinuous, xlMedium, xlColorIndexAutomatic, 0
          xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).Borders(xlEdgeTop).LineStyle = xlContinuous
          xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).Borders(xlEdgeTop).Weight = xlMedium
          xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).NumberFormat = "#,##0.00"
          xlHoja1.Cells(nLin, 8) = nVal
          lnTotRend = lnTotRend + nVal
          nLin = nLin + 1
          xlHoja1.Cells(nLin, 6) = "Devolucion a Caja"
          xlHoja1.Cells(nLin, 8) = rsMay!MONTOATENDIDO - nVal
          lnTotDev = lnTotDev + (rsMay!MONTOATENDIDO - nVal)
          'xlHoja1.Range("H" & 0 + nLin & ":H" & 0 + nLin).BorderAround xlContinuous, xlMedium, xlColorIndexAutomatic, 0
          'xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).Borders(xlDiagonalDown).LineStyle = xlContinuous
          nLin = nLin + 1
          xlHoja1.Cells(nLin, 8) = rsMay!MONTOATENDIDO
          lnTotal = lnTotal + rsMay!MONTOATENDIDO
          'xlHoja1.Range("H" & 0 + nLin & ":H" & 0 + nLin).BorderAround xlContinuous, xlMedium, xlColorIndexAutomatic, 0
          xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).Borders(xlEdgeTop).LineStyle = xlContinuous
          xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).Borders(xlEdgeTop).Weight = xlMedium
          xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).NumberFormat = "#,##0.00"
          nLin = nLin + 2
          End If
          
          rsMay.MoveNext
          If rsMay.EOF Then
             Exit Do
          End If
       Loop
       nLin = nLin + 1
       xlHoja1.Cells(nLin, 6) = "Total de Rendiciones"
       xlHoja1.Cells(nLin, 8) = lnTotRend
       nLin = nLin + 1
       xlHoja1.Cells(nLin, 6) = "Total de Devoluciones a Caja"
       xlHoja1.Cells(nLin, 8) = lnTotDev
       nLin = nLin + 1
       xlHoja1.Cells(nLin, 6) = "TOTAL"
       xlHoja1.Cells(nLin, 8) = lnTotal
    
    RSClose rsMay
    
    Exit Function
ErrImprime:
     MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
       If lbLibroOpen Then
          xlLibro.Close
          xlAplicacion.Quit
       End If
       Set xlAplicacion = Nothing
       Set xlLibro = Nothing
       Set xlHoja1 = Nothing

End Function

Private Sub ImprimeSustARendirViaticosExcelCab(psUsuario As String, psFechaER As String, pnImporte As Double, psArea As String, _
                                               psLugarViaje As String, psPeriodo As String, psMotivo As String, _
                                               pdFecha As String, pdFecha2 As String, lnLin As Long, psFechaRen As String)

    nLin = lnLin
    
    xlHoja1.Range("A1:S1").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    'xlHoja1.Range("A11:Z1").EntireColumn.Font.Size = 8
    'xlHoja1.Range("B1").EntireColumn.HorizontalAlignment = xlHAlignCenter
    
    xlHoja1.Range("A1:A1").RowHeight = 17
    xlHoja1.Range("A1:A1").ColumnWidth = 8  'Item
    xlHoja1.Range("B1:B1").ColumnWidth = 12 'Fecha
    xlHoja1.Range("C1:C1").ColumnWidth = 15 'Documento
    xlHoja1.Range("D1:D1").ColumnWidth = 8  'Serie
    xlHoja1.Range("E1:E1").ColumnWidth = 12 'Numero
    xlHoja1.Range("F1:F1").ColumnWidth = 60 'Proveedor
    xlHoja1.Range("G1:G1").ColumnWidth = 60 'Detalle
    xlHoja1.Range("H1:H1").ColumnWidth = 12 'Importe
        
    'xlHoja1.Range("B1:B1").Font.Size = 12
    'xlHoja1.Range("A2:B4").Font.Size = 10
    xlHoja1.Cells(nLin, 2) = "Sustentación de a Rendir Cuenta por Viaticos"
    xlHoja1.Range("A" & nLin & ":I" & nLin).Merge True
    xlHoja1.Range("A" & nLin & ":I" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":I" & nLin).HorizontalAlignment = xlHAlignCenter
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 2) = "Del " & pdFecha & " Al " & pdFecha2
    xlHoja1.Range("A" & nLin & ":I" & nLin).Font.Bold = True
    'xlHoja1.Cells(nLin, 2) = "( Del " & pdFecha & " Al " & pdFecha2 & " )"

    xlHoja1.Range("A" & nLin & ":I" & nLin).Merge True
    'xlHoja1.Range("A" & nLin & ":I" & nLin).Merge True
    xlHoja1.Range("A" & nLin & ":I" & nLin).HorizontalAlignment = xlHAlignCenter
    
    nLin = nLin + 2
    
    xlHoja1.Cells(nLin, 1) = "Usuario"
    xlHoja1.Cells(nLin, 3) = psUsuario
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).Font.Bold = True
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = "Fecha de E/R"
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).NumberFormat = "dd/mm/yyyy;@"
    xlHoja1.Cells(nLin, 3) = psFechaER
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = "Importe Otorgado"
    xlHoja1.Cells(nLin, 3) = pnImporte
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).NumberFormat = "#,###0.00"
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = "Fecha de Rendición"
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).NumberFormat = "dd/mm/yyyy;@"
    xlHoja1.Cells(nLin, 3) = Trim(psFechaRen)
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = "Area Funcional"
    xlHoja1.Cells(nLin, 3) = psArea
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = "Lugar de Viaje"
    xlHoja1.Cells(nLin, 3) = psLugarViaje
    xlHoja1.Range("A" & 0 + nLin & ":B" & 0 + nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = "Periodo de Comision"
    xlHoja1.Cells(nLin, 3) = psPeriodo
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = "Motivo"
    xlHoja1.Cells(nLin, 3) = psMotivo
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft
    
    nLin = nLin + 2
    
    xlHoja1.Cells(nLin, 1) = "Item"
        
    xlHoja1.Cells(nLin, 2) = "Fecha"
    
    xlHoja1.Cells(nLin, 3) = "Documento"
    
    xlHoja1.Cells(nLin, 4) = "Serie"
      
    xlHoja1.Cells(nLin, 5) = "Número"
    
    xlHoja1.Cells(nLin, 6) = "Proveedor"
    
    xlHoja1.Cells(nLin, 7) = "Detalle"
    
    xlHoja1.Cells(nLin, 8) = "Importe"
    
    
    'xlHoja1.Range("D6:E6").Merge True
    'xlHoja1.Range("Q6:S6").Merge True
    
    'xlHoja1.Range("L6:N7").HorizontalAlignment = xlHAlignCenter
    
    xlHoja1.Range("A" & nLin & ":I" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":I" & nLin).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A" & nLin & ":H" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A" & nLin & ":H" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A" & nLin & ":H" & nLin).Borders(xlInsideVertical).Color = vbBlack
    'xlHoja1.Range("D6:E6").Borders(xlEdgeBottom).LineStyle = xlContinuous
    'xlHoja1.Range("Q6:S6").Borders(xlEdgeBottom).LineStyle = xlContinuous


    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""

        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
    End With
End Sub

Private Sub ImprimeDetalleArendirViaticos(pnItem As Integer, psFecha As String, psDocumento As String, psDocNro As String, _
                                          psProveedor As String, psDetalle As String, pnImporte As Double)
    Dim Item As Integer
    
    Item = pnItem
    xlHoja1.Cells(nLin, 1) = Item
    xlHoja1.Cells(nLin, 2) = psFecha
    'If IsDate(psFecha) Then
    '    xlHoja1.Cells(nLin, 2) = CDate(psFecha)
    'Else
    '    xlHoja1.Cells(nLin, 2) = CDate(Mid(psFecha, 7, 2) & "/" & Mid(sFec, 5, 2) & "/" & Left(sFec, 4))
    'End If
    
    xlHoja1.Cells(nLin, 3) = psDocumento
    
    If InStr(1, psDocNro, "-") <> 0 Then
        xlHoja1.Cells(nLin, 4) = "'" & Format(Left(psDocNro, InStr(1, psDocNro, "-") - 1), "000")
        xlHoja1.Cells(nLin, 5) = "'" & Format(Mid(psDocNro, InStr(1, psDocNro, "-") + 1), "000000000")
    Else
        xlHoja1.Cells(nLin, 5) = "'" & psDocNro
    End If
    
    xlHoja1.Cells(nLin, 6) = psProveedor
    xlHoja1.Cells(nLin, 7) = psDetalle
    xlHoja1.Cells(nLin, 8) = pnImporte
    xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).NumberFormat = "#,##0.00"
    nLin = nLin + 1
End Sub

Private Function Devuelvemes(PsMes As Integer) As String
    Dim lsNomMes As String
    Select Case PsMes
        Case 1
            lsNomMes = "Enero"
        Case 2
            lsNomMes = "Febrero"
        Case 3
            lsNomMes = "Marzo"
        Case 4
            lsNomMes = "Abril"
        Case 5
            lsNomMes = "Mayo"
        Case 6
            lsNomMes = "Junio"
        Case 7
            lsNomMes = "Julio"
        Case 8
            lsNomMes = "Agosto"
        Case 9
            lsNomMes = "Setiembre"
        Case 10
            lsNomMes = "Octubre"
        Case 11
            lsNomMes = "Noviembre"
        Case 12
            lsNomMes = "Diciembre"
    End Select
    
    Devuelvemes = lsNomMes
End Function
