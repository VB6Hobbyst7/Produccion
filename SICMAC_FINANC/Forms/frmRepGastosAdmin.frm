VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepGastosAdmin 
   Caption         =   "Reporte de gastos administrativos y operativos"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5445
   Icon            =   "frmRepGastosAdmin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1103
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3023
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.ComboBox cboAreas 
      Height          =   315
      Left            =   323
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   4815
   End
   Begin VB.ComboBox cboAgencias 
      Height          =   315
      Left            =   323
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin MSMask.MaskEdBox txtAl 
      Height          =   330
      Left            =   3259
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDel 
      Height          =   330
      Left            =   1219
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Del :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   746
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Al :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2786
      TabIndex        =   9
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblAreas 
      Caption         =   "Areas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1000
      Width           =   615
   End
   Begin VB.Label lblAgencia 
      Caption         =   "Agencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmRepGastosAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim xlHoja1 As Excel.Worksheet
Dim nLin As Long
Dim lsArchivo As String
Dim oCon As DConecta
Dim rsAgencia As New ADODB.Recordset
Dim rsArea As New ADODB.Recordset
Dim sSql As String

Private Sub cboAgencias_LostFocus()
    Set oCon = New DConecta
    oCon.AbreConexion
    
    sSql = "Select Distinct A.cAreaCod, A.cAreaDescripcion From Areas A " _
         & "Join AreaAgencia AA On A.cAreaCod = AA.cAreaCod " _
         & " Where AA.cUbicaCod = '" & Left(cboAgencias, 2) & "' or AA.cAgeCod = '" & Left(cboAgencias, 2) & "'"
    Set rsArea = oCon.CargaRecordSet(sSql)
    
    RSLlenaCombo rsArea, cboAreas, , , False
    cboAreas.AddItem "XXXX  TODAS LAS AREAS"
    RSClose rsArea
End Sub

Private Sub cmdAceptar_Click()
Dim ldFechaIni As String
Dim ldFechaFin As String
Dim lsAgencia As String
Dim lsArea As String
Dim nU As Integer

ldFechaIni = txtDel.Text
ldFechaFin = txtAl.Text

If ldFechaIni > ldFechaFin Then
    MsgBox "Fecha final debe ser mayor", vbOKOnly, "Error"
    Exit Sub
End If

If cboAgencias.ListIndex >= 0 Then
   If Left(cboAgencias, 4) = "XXXX" Then
      lsAgencia = Left(Me.cboAgencias.List(0), 2)
      For nU = 1 To cboAgencias.ListCount - 1
         lsAgencia = lsAgencia & "','" & Left(Me.cboAgencias.List(nU), 2)
      Next
   Else
      lsAgencia = Left(Me.cboAgencias, 2)
   End If
End If

If cboAreas.ListIndex >= 0 Then
   If Left(cboAreas, 4) = "XXXX" Then
      lsArea = Left(Me.cboAreas.List(0), 3)
      For nU = 1 To cboAreas.ListCount - 1
         lsArea = lsArea & "','" & Left(Me.cboAreas.List(nU), 3)
      Next
   Else
      lsArea = Left(Me.cboAreas, 3)
   End If
End If

lsArchivo = App.path & "\SPOOLER\RGastosAdmin_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"

lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
If Not lbLibroOpen Then
    Exit Sub
End If

If cboAgencias.ListIndex >= 0 Or cboAreas.ListIndex >= 0 Then
   GeneraReporteGastosAdministrativos lsAgencia, lsArea, ldFechaIni, ldFechaFin
Else
   MsgBox "Seleccione una Agencia/Area", vbInformation, "¡Aviso!"
   cboAgencias.SetFocus
End If
 
ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
CargaArchivo lsArchivo, App.path & "\SPOOLER\"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'****************************************************************************'
'** Nombre       : frmRepGastosAdmin
'** Descripción  : formulario para Generar Reporte de gastos administrativos
'** Creación     : GITU, 20080922 09:00 AM
'****************************************************************************'
    Set oCon = New DConecta
    oCon.AbreConexion
    
    sSql = "Select cAgecod, cAgeDescripcion From Agencias"
    Set rsAgencia = oCon.CargaRecordSet(sSql)
    
    RSLlenaCombo rsAgencia, cboAgencias, , , False
    cboAgencias.AddItem "XXXX  Todas Las Agencias"
    RSClose rsAgencia
    
    sSql = "Select cAreaCod, cAreaDescripcion From Areas"
    Set rsArea = oCon.CargaRecordSet(sSql)
    
    RSLlenaCombo rsArea, cboAreas, , , False
    cboAreas.AddItem "XXXX  TODAS LAS AREAS"
    RSClose rsArea
    
End Sub

Function GeneraReporteGastosAdministrativos(psAgencia As String, psAreas As String, psFechaIni As String, _
                                            psFechaFin As String) As Integer
    Dim rsMay As New ADODB.Recordset
    Dim oOperacion As DOperacion
    Set oOperacion = New DOperacion
    Dim lnImporte As Double
    Dim lsArea    As String
    Dim lsAgencia As String
    Dim lnTotal As Double
    Dim lnSalAntTot As Double
    Dim lnTotalAcum As Double
    Dim lsCtaS As String
    Dim lsDocs As String
    Dim lnTotalCta As Double
    Dim sImpre      As String
    Dim sImpreAge   As String
    Dim lsRepTitulo As String
    Dim lsHoja      As String
    Dim lsCta As String
    Dim oSdo  As New NCtasaldo
    Dim oImpreg As New NContImpreReg
    Dim lnSalAnt As Double
    Dim c As Integer
    
    lsCtaS = oOperacion.CargaListaCuentasOperacion("760200", "45")
    
    lsDocs = oOperacion.CargaListaDocsOperacion("760200")
    
    Set rsMay = oImpreg.GetGastosAdministrativosDatos(psAgencia, psAreas, psFechaIni, psFechaFin, lsCtaS, lsDocs)
    
    If rsMay.RecordCount = 0 Then
        MsgBox "No Exiten datos para el reporte", vbInformation, "ATENCION!"
        Exit Function
    End If
    nLin = 1
    lsArea = Mid(Me.cboAreas.Text, 5, 50)
    lsAgencia = Mid(Me.cboAgencias.Text, 4, 50)
    
    lsHoja = "SustARendirViaticos"
    ExcelAddHoja lsHoja, xlLibro, xlHoja1
    'Formato texto a una Columna
    xlHoja1.Range(xlHoja1.Cells(10, 6), xlHoja1.Cells(5000, 6)).NumberFormat = "###,#00.00"
    xlHoja1.Range(xlHoja1.Cells(10, 2), xlHoja1.Cells(5000, 2)).NumberFormat = "m/d/yyyy"
    ImprimeGastosAdminiExcelCab lsArea, lsAgencia, psFechaIni, psFechaFin, nLin
    
    lsCta = rsMay!cCtaContCod
    lnSalAnt = rsMay!nSalAnt
    c = 0
    
    Do Until rsMay.EOF
        nLin = nLin + 1
        
        If rsMay!cCtaContCod = lsCta And rsMay!nSalAnt <> lnSalAnt And c = 0 Then
           lnSalAnt = lnSalAnt + rsMay!nSalAnt
           c = 1
        End If
        
        If rsMay!cCtaContCod <> lsCta Then
            xlHoja1.Range("E" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
            xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
            xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
            xlHoja1.Cells(nLin, 5) = "Total " & lsCta
            xlHoja1.Cells(nLin, 6) = lnTotalCta
            nLin = nLin + 1
            
            xlHoja1.Range("E" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
            xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
            xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
            xlHoja1.Cells(nLin, 5) = "Saldo Anterior " & lsCta
            'lnSalAnt = oSdo.GetCtaSaldo(lsCta, Format(CDate(psFechaIni) - 1, gsFormatoFecha))
            lnSalAntTot = lnSalAntTot + lnSalAnt
            xlHoja1.Cells(nLin, 6) = lnSalAnt
            nLin = nLin + 1
            
            xlHoja1.Range("E" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
            xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
            xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
            xlHoja1.Cells(nLin, 5) = "Total Acumulado" & " " & lsCta
            xlHoja1.Cells(nLin, 6) = lnTotalCta + lnSalAnt
            lnTotalAcum = lnTotalAcum + (lnTotalCta + lnSalAnt)
            nLin = nLin + 2
            lsCta = rsMay!cCtaContCod
            lnSalAnt = rsMay!nSalAnt
            lnTotalCta = 0
            c = 0
        End If
         
        xlHoja1.Range("A" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
        xlHoja1.Range("A" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range("A" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
        
        xlHoja1.Cells(nLin, 1) = rsMay!cCtaContCod
        xlHoja1.Cells(nLin, 2) = rsMay!dDocFecha
        xlHoja1.Cells(nLin, 3) = rsMay!cDocNro
        xlHoja1.Cells(nLin, 4) = rsMay!cPersNombre
        xlHoja1.Cells(nLin, 5) = rsMay!cMovDesc
        xlHoja1.Cells(nLin, 6) = rsMay!nPV
                
        lnTotalCta = lnTotalCta + rsMay!nPV
        lnTotal = lnTotal + rsMay!nPV

        rsMay.MoveNext
        
        If rsMay.EOF Then
            nLin = nLin + 1
            xlHoja1.Range("E" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
            xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
            xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
            xlHoja1.Cells(nLin, 5) = "Total " & lsCta
            xlHoja1.Cells(nLin, 6) = lnTotalCta
            nLin = nLin + 1
            
            xlHoja1.Range("E" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
            xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
            xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
            xlHoja1.Cells(nLin, 5) = "Saldo Anterior " & lsCta
            'lnSalAnt = oSdo.GetCtaSaldo(lsCta, Format(CDate(psFechaIni) - 1, gsFormatoFecha))
            lnSalAntTot = lnSalAntTot + lnSalAnt
            xlHoja1.Cells(nLin, 6) = lnSalAnt
            nLin = nLin + 1
            
            xlHoja1.Range("E" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
            xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
            xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
            xlHoja1.Cells(nLin, 5) = "Total Acumulado" & " " & lsCta
            xlHoja1.Cells(nLin, 6) = lnTotalCta + lnSalAnt
            nLin = nLin + 2
            'lsCta = rsMay!cCtaContCod
            lnTotalCta = 0
        End If
        'If rsMay.EOF Then
        '    Exit Do
        'End If
    Loop
       nLin = nLin + 1
       xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
       xlHoja1.Range("E" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
       xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
       xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
       xlHoja1.Cells(nLin, 5) = "TOTAL"
       xlHoja1.Cells(nLin, 6) = lnTotal
       nLin = nLin + 1
       xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
       xlHoja1.Range("E" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
       xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
       xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
       xlHoja1.Cells(nLin, 5) = "Saldo Anterior"
       xlHoja1.Cells(nLin, 6) = lnSalAntTot
       nLin = nLin + 1
       xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
       xlHoja1.Range("E" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
       xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
       xlHoja1.Range("E" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
       xlHoja1.Cells(nLin, 5) = "Total Acumulado"
       xlHoja1.Cells(nLin, 6) = lnTotal + lnSalAntTot

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

Private Sub ImprimeGastosAdminiExcelCab(psArea As String, psAgencia As String, pdFechaIni As String, pdFechaFin As String, lnLin As Long)

    nLin = lnLin
    
    xlHoja1.Range("A1:F1").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    xlHoja1.Range("A1:Z20000").EntireColumn.Font.Size = 8
    'xlHoja1.Range("B1").EntireColumn.HorizontalAlignment = xlHAlignCenter
    
    xlHoja1.Range("A3:A3").RowHeight = 17
    xlHoja1.Range("A1:A1").ColumnWidth = 12
    xlHoja1.Range("B1:B1").ColumnWidth = 12 'Fecha
    xlHoja1.Range("C1:C1").ColumnWidth = 15 'Documento
    xlHoja1.Range("D1:D1").ColumnWidth = 30 'Proveedor
    xlHoja1.Range("E1:E1").ColumnWidth = 50 'Descripcion
    xlHoja1.Range("F1:F1").ColumnWidth = 12 'Total
        
    xlHoja1.Cells(nLin, 1) = gsNomCmac
    xlHoja1.Cells(nLin, 6) = Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time(), "HH:MM")
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = gsNomAge
    xlHoja1.Cells(nLin, 6) = gsCodUser
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 2) = "REPORTE DE GASTOS ADMINISTRATIVOS Y OPERATIVOS"
    xlHoja1.Range("A" & nLin & ":F" & nLin).Merge True
    xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":F" & nLin).HorizontalAlignment = xlHAlignCenter
    nLin = nLin + 1
    
    xlHoja1.Cells(nLin, 1) = "Agencia :" & psAgencia
    xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = "Area : " & psArea
    xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = "Periodo : " & pdFechaIni & " - " & pdFechaFin
    xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
    'xlHoja1.Cells(nLin, 2) = "( Del " & pdFecha & " Al " & pdFecha2 & " )"

    'xlHoja1.Range("A" & nLin & ":I" & nLin).Merge True
    'xlHoja1.Range("A" & nLin & ":I" & nLin).Merge True
    'xlHoja1.Range("A" & nLin & ":I" & nLin).HorizontalAlignment = xlHAlignCenter
    
    nLin = nLin + 2
       
    'xlHoja1.Cells(nLin, 1) = "Item"
    xlHoja1.Range(xlHoja1.Cells(nLin, 1), xlHoja1.Cells(nLin, 6)).Interior.ColorIndex = 33 '.Color = RGB(159, 206, 238)
    xlHoja1.Cells(nLin, 2) = "Documento"
    xlHoja1.Range("B" & nLin & ":C" & nLin).Merge True
    
    xlHoja1.Cells(nLin, 4) = "Proveedor"
    xlHoja1.Range("D8:D9").Merge True
    
    xlHoja1.Cells(nLin, 5) = "Descripción"
    xlHoja1.Range("E8:E9").Merge True
    
    xlHoja1.Cells(nLin, 6) = "Total"
    xlHoja1.Range("F8:F9").Merge True
    
    xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":F" & nLin).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("B" & nLin & ":C" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("B" & nLin & ":C" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("B" & nLin & ":C" & nLin).Borders(xlInsideVertical).Color = vbBlack
    
    nLin = nLin + 1
    xlHoja1.Range(xlHoja1.Cells(nLin, 1), xlHoja1.Cells(nLin, 6)).Interior.ColorIndex = 33 '.Color = RGB(159, 206, 238)
    xlHoja1.Cells(nLin, 2) = "Fecha"
    
    xlHoja1.Cells(nLin, 3) = "Nº Documento"
    
    xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":F" & nLin).HorizontalAlignment = xlHAlignCenter
    
    
    'xlHoja1.Range("D6:E6").Merge True
    'xlHoja1.Range("Q6:S6").Merge True
    
    'xlHoja1.Range("L6:N7").HorizontalAlignment = xlHAlignCenter
    
    xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":F" & nLin - 1).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A" & nLin & ":F" & nLin - 1).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A" & nLin & ":F" & nLin - 1).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A" & nLin & ":F" & nLin - 1).Borders(xlInsideVertical).Color = vbBlack
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
