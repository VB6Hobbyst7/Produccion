VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReportePosCambiaria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REQUERIMIENTO PATRIMONIAL DE POSICION CAMBIARIA"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   Icon            =   "frmReportePosCambiaria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPatrimonio 
      Caption         =   "Patrimonio"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame fraPerGeneracion 
      Caption         =   "Reporte Posición Cambiaria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.CheckBox chkBalance 
         Caption         =   "Balance"
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
         Left            =   5160
         TabIndex        =   23
         Top             =   950
         Width           =   1018
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   4680
         TabIndex        =   19
         Top             =   960
         Width           =   2415
         Begin VB.TextBox txtAnio 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1530
            MaxLength       =   4
            TabIndex        =   22
            Top             =   360
            Width           =   735
         End
         Begin VB.ComboBox cboMes 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "-"
            Height          =   375
            Left            =   1320
            TabIndex        =   21
            Top             =   390
            Width           =   135
         End
      End
      Begin VB.TextBox txtMes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtTpoCambio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtPatrimonioDolares 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtPatrimonioSoles 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2160
         Width           =   1695
      End
      Begin VB.ComboBox cboDig 
         Height          =   315
         ItemData        =   "frmReportePosCambiaria.frx":030A
         Left            =   3360
         List            =   "frmReportePosCambiaria.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   675
         Width           =   645
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   675
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblPatEfecDolares 
         Caption         =   "Patrimonio Efectivo $:"
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblTpoCambioMes 
         Caption         =   "T/C Fijo Contable Mes:"
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblPatriEfectSoles 
         Caption         =   "Patrimonio Efectivo S/.:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblMesPerComparativo 
         Caption         =   "Mes Periodo Comparativo:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblInfPeriodo 
         Caption         =   "Información del Periodo"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblNivDet 
         Caption         =   "Nivel Detalle:"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDatGen 
         Caption         =   "Datos Generación"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   2940
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmReportePosCambiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '************************************************
'MIOL 20130115, SEGUN RFC138- 2013
'************************************************
Dim oDbalanceCont As DbalanceCont
Dim rsPatrimonio As ADODB.Recordset
Dim oRepCtaColumna As DRepCtaColumna
Dim pdFechaBalance As Date 'NAGL 20170719
Dim pdFechaControl As Date 'NAGL 20170718
Dim oGen As New DGeneral 'NAGL 20170715
Dim pnTipoIng As Integer

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim nDolares As Currency
    Dim pdFecha As Date 'NAGL 20170718
    pnTipoIng = 0
    Me.txtFecha.Text = gdFecSis - 1
    pdFecha = Me.txtFecha.Text
    cboDig.ListIndex = 0
    Set rs = oGen.GetConstante(1010)
    While Not rs.EOF
       cboMes.AddItem rs.Fields(0) & space(50) & rs.Fields(1)
       rs.MoveNext
    Wend
    Call CalculaPatrimonioEfectivo(pnTipoIng, pdFecha) 'NAGL 20170718
End Sub

Private Sub CalculaPatrimonioEfectivo(ByVal pnTipo As Integer, ByVal pdFecha As Date)
    Set oDbalanceCont = New DbalanceCont
    Set rsPatrimonio = New ADODB.Recordset
    Dim pdFechaFinMesAnt As Date 'NAGL 20170718
    Dim rs As New ADODB.Recordset 'NAGL 20170715

'******NAGL 20170718
If pnTipo = 0 Or pnTipo = 1 Then
    If Day(pdFecha) >= 15 Then
        pdFechaFinMesAnt = DateAdd("d", -Day(pdFecha), pdFecha)
    Else
        pdFechaFinMesAnt = DateAdd("d", -Day(DateAdd("m", -1, pdFecha)), DateAdd("m", -1, pdFecha))
    End If
ElseIf pnTipo = 3 Then  '3 -> Balance
   pdFechaFinMesAnt = DateAdd("m", -2, pdFecha)
End If
'******END 20170718

    Set rsPatrimonio = oDbalanceCont.recuperarPatrimonioEfectivoEleccion(pnTipo, CInt(Month(pdFechaFinMesAnt)), CInt(Year(pdFechaFinMesAnt)), pdFecha, gdFecSis)
    
    If Not rsPatrimonio.BOF And Not rsPatrimonio.EOF Then
        CantReg = rsPatrimonio.RecordCount
        txtPatrimonioSoles = Format(rsPatrimonio!nSaldo, "#,##0.00")
        txtMes = rsPatrimonio!sMes
            'If rsPatrimonio.Fields(4) = Month(pdFechaFinMesAnt) Then 'NAGL 20170718
                'txtPatrimonioSoles = Format(rsPatrimonio.Fields(1), "#,##0.00###")
                'txtMes = UCase(rsPatrimonio.Fields(2))
            'Else
                'Dim oForm As New frmRegPatrimonioEfectivo
                'CentraForm oForm 'Formulario al Centro
                'oForm.Show 1, Me
               
                'Set rsPatrimonio = oDbalanceCont.recuperarPatrimonioEfectivoEleccion(pnTipo, CInt(Month(pdFechaFinMesAnt)), CInt(Year(pdFechaFinMesAnt)), pdFecha, gdFecSis)
                    'If Not rsPatrimonio.BOF And Not rsPatrimonio.EOF Then
                            'txtPatrimonioSoles = Format(rsPatrimonio.Fields(1), "#,##0.00###")
                            'txtMes = UCase(rsPatrimonio.Fields(2))
                    'End If
            'End If 'Comentado by NAGL 20170914
    Else
        txtPatrimonioSoles = ""
        txtMes = ""
        txtPatrimonioDolares = ""
        txtTpoCambio = ""
        'MsgBox "No existe patrimonio efectivo del mes ingresado..!, Favor de Registrarlo", vbInformation + vbOKOnly
        Exit Sub
    End If
    
 'txtTpoCambio = TipoCambioCierre(Year(DateAdd("m", -1, pdFecha)), Month(DateAdd("m", -1, pdFecha))) Comentado by NAGL 20170914
 txtTpoCambio = oDbalanceCont.ObtenerTipoCambioCierreNew(pdFecha)
 txtTpoCambio = IIf(txtTpoCambio = 0, 0, Format(txtTpoCambio, "#,##.000"))
 If txtTpoCambio <> 0 Then
    nDolares = txtPatrimonioSoles / txtTpoCambio
    txtPatrimonioDolares = Format(nDolares, "#,##0.00")
 End If
  '*******NAGL 20170718
     If pnTipo = 0 Or pnTipo = 1 Then
        cboMes.ListIndex = CInt(Month(pdFechaFinMesAnt)) - 1
        txtAnio.Text = CInt(Year(pdFechaFinMesAnt))
     'Else
        'cboMes.ListIndex = CInt(Month(DateAdd("d", -1, pdFecha))) - 1
        'txtAnio.Text = CInt(Year(DateAdd("d", -1, pdFecha)))
     End If
'*****END NAGL 20170718
 
 Set rsPatrimonio = Nothing
 Set oDbalanceCont = Nothing
End Sub '************NAGL incluido en un método con los agregados respectivos

Private Function ValidaFecha(pdFecha As Date) As Boolean
If pdFecha > gdFecSis Then
   MsgBox "La Fecha Ingresada es Incorrecta", vbInformation, "Atención"
   txtFecha.SetFocus
   Exit Function
End If
ValidaFecha = True
End Function

Public Function CalculaBalance() As Date
Dim psBalance As Date
    If (CInt(cboMes.ListIndex) + 2) < 10 Then
        psBalance = CDate("01" & "/" & "0" & CStr(CInt(cboMes.ListIndex) + 2) & "/" & CStr(txtAnio.Text))
        ElseIf (CInt(cboMes.ListIndex) + 1) = 12 Then
        psBalance = CDate("01" & "/" & "01" & "/" & CStr(CInt((txtAnio.Text)) + 1))
        Else
        psBalance = CDate("01" & "/" & CStr(CInt(cboMes.ListIndex) + 2) & "/" & CStr(txtAnio.Text))
    End If
    CalculaBalance = psBalance
End Function '*****NAGL 20170719

Private Sub chkBalance_Click()
pdFechaBalance = CalculaBalance
    If chkBalance.value = 1 Then
        pnTipoIng = 3
        cboMes.Enabled = True
        txtAnio.Enabled = True
        txtAnio.SetFocus
        txtFecha.Text = "__/__/____"
        txtFecha.Enabled = False
        txtMes.Enabled = False
        txtPatrimonioSoles.Enabled = False
        txtPatrimonioDolares.Enabled = False
        txtTpoCambio.Enabled = False
        Call CalculaPatrimonioEfectivo(pnTipoIng, pdFechaBalance)
    Else
        pnTipoIng = 1
        cboMes.Enabled = False
        txtAnio.Enabled = False
        txtFecha.Text = gdFecSis - 1
        txtFecha.Enabled = True
        txtFecha.SetFocus
        txtMes.Enabled = True
        txtPatrimonioSoles.Enabled = True
        txtPatrimonioDolares.Enabled = True
        txtTpoCambio.Enabled = True
        Call CalculaPatrimonioEfectivo(pnTipoIng, txtFecha)
    End If
End Sub '******NAGL 20170718

Private Sub CboMes_Click()
pnTipoIng = 3
If (txtFecha = "" Or txtFecha = "__/__/____") Then
    pdFechaBalance = CalculaBalance
    txtAnio.SetFocus
    Call CalculaPatrimonioEfectivo(pnTipoIng, pdFechaBalance)
End If
End Sub '*****NAGL 20170718

Public Function ValFecRegPatrimonioEfectivo()

    If chkBalance.value = 1 Then
        pdFechaControl = CalculaBalance
        pdFechaControl = DateAdd("d", -1, pdFechaControl)
         If pdFechaControl > gdFecSis Then
            MsgBox " No existe el Balance con el mes Ingresado ...! ", vbInformation, "Aviso"
            cboMes.SetFocus
            Exit Function
        End If
        'If txtPatrimonioSoles = "" Or txtPatrimonioDolares = "" Then
            'MsgBox "No existe Patrimonio Efectivo con el mes ingresado...!, Favor de Registrarlo", vbInformation, "Aviso"
            'Exit Function
        'End If
    Else
        pdFechaControl = txtFecha
        If pdFechaControl > gdFecSis Then
            MsgBox " Fecha Ingresada es Mayor a la Fecha Actual ...!", vbInformation, "Aviso"
            txtFecha.SetFocus
            Exit Function
        End If
        'If txtPatrimonioSoles = "" Or txtPatrimonioDolares = "" Then
            'MsgBox "No existe Patrimonio Efectivo con el mes ingresado..!, Favor de Registrarlo", vbInformation, "Aviso"
            'txtfecha.SetFocus
            'Exit Function
        'End If
   End If
ValFecRegPatrimonioEfectivo = True
End Function '*****NAGL 20170718


Private Sub cmdGenerar_Click()

Dim pdFecha As Date ', psMesBalanceDiario As String, psAnioBalanceDiario As String
Dim psOptBal As Integer, pnNivelDet As Integer
Dim pnPatriSoles As Double, pnPatriDolares As Double, pnTipoCambio As Double

pnNivelDet = CInt(cboDig.ListIndex) * 2
psOptBal = chkBalance.value

        If (psOptBal = 1) Then
            'psAnioBalanceDiario = txtAnio.Text
            'If (CInt(CboMes.ListIndex) + 1) < 10 Then
                'psMesBalanceDiario = "0" & CStr(CInt(CboMes.ListIndex) + 1)
            'Else
                'psMesBalanceDiario = CStr(CInt(CboMes.ListIndex) + 1)
            'End If
               If ValFecRegPatrimonioEfectivo Then
                   pnPatriSoles = txtPatrimonioSoles
                   pnPatriDolares = txtPatrimonioDolares
                   pnTipoCambio = txtTpoCambio
                   Call GenerarReportePosCambBalance(pnNivelDet, pnTipoCambio, pnPatriSoles, pnPatriDolares)
               End If
        Else
            If ValFecha(txtFecha) = True Then
                If ValidaFecha(txtFecha) Then
                    pdFecha = txtFecha.Text
                    If ValFecRegPatrimonioEfectivo Then 'Valida Datos con Respecto a la Fecha del Sistema y el Patrimonio Efectivo
                       pnPatriSoles = txtPatrimonioSoles
                       pnPatriDolares = txtPatrimonioDolares
                       pnTipoCambio = txtTpoCambio
                       Call GenerarReportePosCambiaria(pnNivelDet, pnTipoCambio, pnPatriSoles, pnPatriDolares, pdFecha)
                    End If
                End If
            End If
        End If
End Sub '***NAGL 20170719

Private Sub GenerarReportePosCambBalance(pnNivelDet As Integer, pnTipoCambio As Double, pnPatriSoles As Double, pnPatriDolares As Double)
Dim oCtaIf As NCajaCtaIF
Set oCtaIf = New NCajaCtaIF
Dim oclsCtaCont As DCtaCont
Set oclsCtaCont = New DCtaCont
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim RSTEMP As ADODB.Recordset
Set RSTEMP = New ADODB.Recordset
Dim rsLim As ADODB.Recordset
Set rsLim = New ADODB.Recordset
Dim lsMoneda As String
Dim oCont As New NContFunciones
Set oRepCtaColumna = New DRepCtaColumna


Dim fs              As Scripting.FileSystemObject
Dim xlAplicacion    As Excel.Application
Dim xlLibro         As Excel.Workbook
Dim xlHoja1         As Excel.Worksheet
Dim lbExisteHoja    As Boolean
Dim lilineas        As Integer
Dim i               As Integer
Dim glsArchivo      As String
Dim lsNomHoja       As String
Dim Poscamb         As Currency
Dim cuActivo        As String
Dim cuPasivo        As String
Dim cuActivPasiv    As String, cuTipCamb As String, cuPromEfectAnt As String
Dim lsTotal()       As String
Dim lsCadena()      As String
Dim liLineasInicio   As Long
Dim Cant            As Long
Dim lenCell         As Integer
Dim CellParam       As String
Dim cCtaCont        As String
Dim lsMovNro        As String
Dim pdFechaBalAnt   As Date
Dim nActivo As Double, nPasivo As Double
ReDim lsTotal(2)
ReDim lsCadena(2)

PB1.Min = 0
PB1.Max = 12
PB1.value = 0
PB1.Visible = True

If (pnNivelDet = 0) Then
    pnNivelDet = 1
End If
PB1.value = 2
Set RSTEMP = oclsCtaCont.ListarCtaContPosCamb(pnNivelDet)
    
    If RSTEMP Is Nothing Then
        MsgBox "No exite informacion para imprimir", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If rs Is Nothing Then
        MsgBox "No exite informacion para imprimir", vbInformation, "Aviso"
        Exit Sub
    End If
    
    glsArchivo = "Reporte Posición Cambiaria" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 100
    xlHoja1.PageSetup.Orientation = xlLandscape

     lbExisteHoja = False
     lsNomHoja = "PosCam"
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
     
     xlAplicacion.Range("A1:A1").ColumnWidth = 15
     xlAplicacion.Range("B1:B1").ColumnWidth = 20
     xlAplicacion.Range("c1:c1").ColumnWidth = 15
     xlAplicacion.Range("D1:D1").ColumnWidth = 25
     xlAplicacion.Range("E1:E1").ColumnWidth = 15
     xlAplicacion.Range("F1:F1").ColumnWidth = 15

     xlAplicacion.Range("A1:Z10000").Font.Size = 9
     xlAplicacion.Range("A1:Z10000").Font.Name = "Century Gothic"

     xlHoja1.Cells(1, 1) = "REPORTE DE POSICIÓN CAMBIARIA"
     xlHoja1.Cells(2, 1) = "CMAC MAYNAS"
     xlHoja1.Cells(5, 1) = "CONTROL LIMITE INTERNO"
     
     xlHoja1.Cells(2, 4) = "FECHA REPORTE:"
     xlHoja1.Cells(2, 5) = Format(DateAdd("d", -1, pdFechaBalance), "dd/mm/yyyy")
     xlHoja1.Cells(3, 4) = "PE mes anterior:"
     xlHoja1.Cells(3, 5) = txtPatrimonioSoles
     xlHoja1.Cells(4, 4) = "Activo - Pasivo:"
     xlHoja1.Cells(4, 4).Font.Bold = True
     xlHoja1.Cells(5, 4) = "Posición Cambiaria %:"
     xlHoja1.Cells(5, 4).Font.Bold = True
     xlHoja1.Cells(6, 4) = "T.C. SBS:"
     xlHoja1.Cells(6, 5) = Format(pnTipoCambio, "#,##0.000")

     xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 5)).HorizontalAlignment = xlCenter 'Titulo Princ.
     xlHoja1.Range(xlHoja1.Cells(2, 4), xlHoja1.Cells(6, 4)).HorizontalAlignment = xlRight
     xlHoja1.Range(xlHoja1.Cells(2, 5), xlHoja1.Cells(6, 5)).HorizontalAlignment = xlCenter
                 
     xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 5)).Font.Bold = True
     xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 3)).Font.Bold = True
     xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 3)).Merge True
     xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 5)).Merge True
          
     xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 3)).Font.Bold = True
     
     xlHoja1.Cells(2, 7) = "Limite"
     xlHoja1.Cells(2, 7).ColumnWidth = 12
     xlHoja1.Cells(2, 8) = "SobreCompra"
     xlHoja1.Cells(2, 8).ColumnWidth = 13
     xlHoja1.Cells(2, 9) = "SobreVenta"
     xlHoja1.Cells(2, 9).ColumnWidth = 13
     
     xlHoja1.Range(xlHoja1.Cells(2, 7), xlHoja1.Cells(2, 9)).HorizontalAlignment = xlCenter
     xlHoja1.Range(xlHoja1.Cells(2, 7), xlHoja1.Cells(2, 9)).Font.Bold = True
     xlHoja1.Range(xlHoja1.Cells(2, 7), xlHoja1.Cells(2, 9)).Borders.LineStyle = 1
     xlHoja1.Range(xlHoja1.Cells(2, 7), xlHoja1.Cells(2, 9)).Interior.ColorIndex = 15
     
     PB1.value = 4
     Set rsLim = oRepCtaColumna.GetLimitePosCamb()
     lilineas = 3
     Do Until rsLim.EOF
         xlHoja1.Cells(lilineas, 7) = rsLim(1)
         xlHoja1.Range(xlHoja1.Cells(lilineas, 7), xlHoja1.Cells(lilineas, 7)).HorizontalAlignment = xlCenter
         xlHoja1.Cells(lilineas, 7).Borders.LineStyle = 1
         xlHoja1.Cells(lilineas, 7).Font.Bold = True
         xlHoja1.Cells(lilineas, 8) = rsLim(2)
         xlHoja1.Range(xlHoja1.Cells(lilineas, 8), xlHoja1.Cells(lilineas, 8)).HorizontalAlignment = xlCenter
         xlHoja1.Cells(lilineas, 8).NumberFormat = "#,###0.00"
         xlHoja1.Cells(lilineas, 8).Borders.LineStyle = 1
         xlHoja1.Cells(lilineas, 9) = rsLim(3)
         xlHoja1.Range(xlHoja1.Cells(lilineas, 9), xlHoja1.Cells(lilineas, 9)).HorizontalAlignment = xlCenter
         xlHoja1.Cells(lilineas, 9).NumberFormat = "#,###0.00"
         xlHoja1.Cells(lilineas, 9).Borders.LineStyle = 1
         lilineas = lilineas + 1
         rsLim.MoveNext
     Loop
     Set rsLim = Nothing

     lilineas = 8
     
     xlHoja1.Cells(lilineas, 1) = "CUENTA CONTABLE"
     xlHoja1.Cells(lilineas, 2) = "DESCRIPCION"
     
     pdFechaBalance = DateAdd("d", -1, CalculaBalance())
     xlHoja1.Cells(lilineas, 3) = "SALDO ME ACUMULADO al (" & Format(pdFechaBalance, "dd/mm/yyyy") & ")"
     pdFechaBalAnt = DateAdd("d", -Day(pdFechaBalance), pdFechaBalance)
     xlHoja1.Cells(lilineas, 4) = "SALDO ME ACUMULADO al (" & Format(pdFechaBalAnt, "dd/mm/yyyy") & ")"
     xlHoja1.Cells(lilineas, 5) = "DIFERENCIA"
     
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).HorizontalAlignment = xlCenter
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).VerticalAlignment = xlCenter
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas + 5, 1)).Merge True
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).EntireRow.AutoFit
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).WrapText = True

     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).Font.Bold = True
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).Borders.LineStyle = 1
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).Interior.ColorIndex = 35
     
    lilineas = lilineas + 1
    PB1.value = 6
    Set oCtaIf = New NCajaCtaIF
    Do Until RSTEMP.EOF
            xlHoja1.Cells(lilineas, 1) = RSTEMP!cCtaContCod
            xlHoja1.Cells(lilineas, 1).HorizontalAlignment = xlRight
            xlHoja1.Cells(lilineas, 2) = RSTEMP!cCtaContDesc
            xlHoja1.Cells(lilineas, 2).ColumnWidth = 90

            If Len(RSTEMP!cCtaContCod) = pnNivelDet Then
                Set rs = oCtaIf.GetSaldoMEPosCambiariaNewYBalanc("PCBAL", pdFechaBalance, RSTEMP!cCtaContCod)
                
                     xlHoja1.Cells(lilineas, 3) = rs!nSaldoMEBalIng
                     xlHoja1.Cells(lilineas, 3).NumberFormat = "#,###0.00"
                     lsTotal(1) = xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 3)).Address(False, False)
                     
                     xlHoja1.Cells(lilineas, 4) = rs!nSaldoMEBalAnt
                     xlHoja1.Cells(lilineas, 4).NumberFormat = "#,###0.00"
                     lsTotal(2) = xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas, 4)).Address(False, False)
                     
                     xlHoja1.Cells(lilineas, 5).Formula = "=" & lsTotal(1) & "-" & lsTotal(2)
                     xlHoja1.Cells(lilineas, 5).NumberFormat = "#,###0.00"
            End If
            
           If RSTEMP!cCtaContCod = "1" Or RSTEMP!cCtaContCod = "2" Then
                xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).Interior.ColorIndex = 44
           End If
        lilineas = lilineas + 1
        RSTEMP.MoveNext
     Loop
     
    ExcelCuadro xlHoja1, 1, 9, 5, CCur(lilineas - 1)
    lsTotal(1) = ""
    lsTotal(2) = ""
    lilineas = lilineas - 1
    liLineasInicio = 9
    PB1.value = 8
    Do While lilineas >= liLineasInicio
            If xlHoja1.Cells(lilineas, 3) = "" Then
                    lenCell = Len(xlHoja1.Cells(lilineas, 1))
                    lenCell = Len(xlHoja1.Cells(lilineas, 1)) + IIf(Len(xlHoja1.Cells(lilineas, 1)) = 1, 1, 2)
                    CellParam = xlHoja1.Cells(lilineas, 1) 'Parametro de Inicio para comparar la celda en cuestión
                    Cant = 0
                    Do While (Mid(xlHoja1.Cells(lilineas + Cant, 1), 1, Len(CellParam)) = CellParam)
                            If (Len(xlHoja1.Cells(lilineas + Cant, 1)) = lenCell) Then
                                lsTotal(1) = xlHoja1.Range(xlHoja1.Cells(lilineas + Cant, 3), xlHoja1.Cells(lilineas + Cant, 3)).Address(False, False)
                                lsCadena(1) = lsCadena(1) & lsTotal(1) & ","
                                lsTotal(2) = xlHoja1.Range(xlHoja1.Cells(lilineas + Cant, 4), xlHoja1.Cells(lilineas + Cant, 4)).Address(False, False)
                                lsCadena(2) = lsCadena(2) & lsTotal(2) & ","
                            End If
                            Cant = Cant + 1
                    Loop
                    If (lsCadena(1) <> "") Then
                             lsCadena(1) = "(" & Mid(lsCadena(1), 1, Len(lsCadena(1)) - 1) & ")"
                             lsCadena(2) = "(" & Mid(lsCadena(2), 1, Len(lsCadena(2)) - 1) & ")"
                             xlHoja1.Cells(lilineas, 3).Formula = "=" & "Sum" & lsCadena(1)
                             xlHoja1.Cells(lilineas, 3).NumberFormat = "#,###0.00"
                             xlHoja1.Cells(lilineas, 4).Formula = "=" & "Sum" & lsCadena(2)
                             xlHoja1.Cells(lilineas, 4).NumberFormat = "#,###0.00"
                     Else
                             cCtaCont = CellParam
                             Set rs = oCtaIf.GetSaldoMEPosCambiariaNewYBalanc("PCBAL", pdFechaBalance, cCtaCont)
                             xlHoja1.Cells(lilineas, 3) = rs!nSaldoMEBalIng
                             xlHoja1.Cells(lilineas, 3).NumberFormat = "#,###0.00"
                             xlHoja1.Cells(lilineas, 4) = rs!nSaldoMEBalAnt
                             xlHoja1.Cells(lilineas, 4).NumberFormat = "#,###0.00"
                     End If
                             xlHoja1.Cells(lilineas, 5).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 3)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas, 4)).Address(False, False)
                             xlHoja1.Cells(lilineas, 5).NumberFormat = "#,###0.00"
                             
                     If CellParam = "1" Then
                          cuActivo = xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 3)).Address(False, False)
                     End If
                     If CellParam = "2" Then
                          cuPasivo = xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 3)).Address(False, False)
                     End If
                     lsTotal(1) = ""
                     lsCadena(1) = ""
                     lsTotal(2) = ""
                     lsCadena(2) = ""
            End If
        lilineas = lilineas - 1
    Loop
        PB1.value = 10
        Set rs = Nothing
        If (cuActivo <> "" And cuPasivo <> "") Then
            xlHoja1.Cells(4, 5).Formula = "=" & "+" & cuActivo & "-" & cuPasivo
            xlHoja1.Cells(4, 5).NumberFormat = "#,###0.00"
        Else
            Set rs = oCtaIf.GetSaldoMEPosCambiariaNewYBalanc("PCBAL", pdFechaBalance, "1")
            nActivo = rs!nSaldoMEBalIng
            Set rs = oCtaIf.GetSaldoMEPosCambiariaNewYBalanc("PCBAL", pdFechaBalance, "2")
            nPasivo = rs!nSaldoMEBalIng
            xlHoja1.Cells(4, 5) = nActivo - nPasivo
            xlHoja1.Cells(4, 5).NumberFormat = "#,###0.00"
        End If
        
        cuActivPasiv = xlHoja1.Range(xlHoja1.Cells(4, 5), xlHoja1.Cells(4, 5)).Address(False, False)
        cuTipCamb = xlHoja1.Range(xlHoja1.Cells(6, 5), xlHoja1.Cells(6, 5)).Address(False, False)
        cuPromEfectAnt = xlHoja1.Range(xlHoja1.Cells(3, 5), xlHoja1.Cells(3, 5)).Address(False, False)
        
        xlHoja1.Cells(5, 5).Formula = "=" & "(" & cuActivPasiv & "*" & cuTipCamb & ")" & "/" & cuPromEfectAnt & "*" & "100"
        xlHoja1.Cells(5, 5).NumberFormat = "#,###0.00"
        Poscamb = xlHoja1.Cells(5, 5)
   
        If Poscamb < 0 Then 'SobreVenta
        
            Dim nRegSC As Currency
            Dim nIntSC As Currency
            Dim nTemSC As Currency

            nRegSC = xlHoja1.Cells(3, 8)
            nIntSC = xlHoja1.Cells(4, 8)
            nTemSC = xlHoja1.Cells(5, 8)

            Poscamb = Poscamb * -1
            If nTemSC - Poscamb > 0 And nTemSC - Poscamb <= 1 Then
                xlHoja1.Cells(6, 1) = "CERCANO A ALERTA TEMPRANA"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nTemSC = Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE DE ALERTA TEMPRANA"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf (nTemSC - Poscamb < 0 And nTemSC - Poscamb > -1) Or (nTemSC < Poscamb And Poscamb < nIntSC) Then
                xlHoja1.Cells(6, 1) = "LIMITE DE ALERTA TEMPRANA SUPERADO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nIntSC = Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE INTERNO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nIntSC < Poscamb And nIntSC < nRegSC And Poscamb < nRegSC Then
                xlHoja1.Cells(6, 1) = "LIMITE INTERNO SUPERADO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nRegSC = Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE REGULATORIO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nRegSC < Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE REGULATORIO SUPERADO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            Else
                xlHoja1.Cells(6, 1) = "ACEPTABLE"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbBlue
            End If
        Else 'SobreCompra
            Dim nRegSV As Currency
            Dim nIntSV As Currency
            Dim nTemSV As Currency

            nRegSV = xlHoja1.Cells(3, 9)
            nIntSV = xlHoja1.Cells(4, 9)
            nTemSV = xlHoja1.Cells(5, 9)
            
            If nTemSV - Poscamb > 0 And nTemSV - Poscamb <= 1 Then
                xlHoja1.Cells(6, 1) = "CERCANO A ALERTA TEMPRANA"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nTemSV = Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE DE ALERTA TEMPRANA"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf (nTemSV - Poscamb < 0 And nTemSV - Poscamb > -1) Or (nTemSV < Poscamb And Poscamb < nIntSV) Then
                xlHoja1.Cells(6, 1) = "LIMITE DE ALERTA TEMPRANA SUPERADO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nIntSV = Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE INTERNO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nIntSV < Poscamb And nIntSV < nRegSV And Poscamb < nRegSV Then
                xlHoja1.Cells(6, 1) = "LIMITE INTERNO SUPERADO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nRegSV = Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE REGULATORIO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nRegSV < Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE REGULATORIO SUPERADO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            Else
                xlHoja1.Cells(6, 1) = "ACEPTABLE"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbBlue
            End If
        End If
        
        PB1.value = 12
        Set RSTEMP = Nothing
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        
        'Set xlAplicacion = Nothing
        'Set xlLibro = Nothing
        'Set xlHoja1 = Nothing
        'MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        'Call CargaArchivo(glsArchivo, App.path & "\SPOOLER\")
        
        xlAplicacion.Visible = True
        xlAplicacion.Windows(1).Visible = True
        Set xlAplicacion = Nothing
        Set xlHoja1 = Nothing
        
        PB1.Visible = False
        Set oCtaIf = Nothing
        Exit Sub
ReporteAdeudadosVinculadosErr:
        MsgBox Err.Description, vbInformation, "Aviso"
        Exit Sub
End Sub '********NAGL 20170725

Private Sub GenerarReportePosCambiaria(pnNivelDet As Integer, pnTipoCambio As Double, pnPatriSoles As Double, pnPatriDolares As Double, pdFecha As Date)
Dim oCtaIf As NCajaCtaIF
Set oCtaIf = New NCajaCtaIF
Dim oclsCtaCont As DCtaCont
Set oclsCtaCont = New DCtaCont
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim RSTEMP As ADODB.Recordset
Set RSTEMP = New ADODB.Recordset
Dim rsLim As ADODB.Recordset
Set rsLim = New ADODB.Recordset
Dim lsMoneda As String
Dim oCont As New NContFunciones 'NAGL 20170725
Set oRepCtaColumna = New DRepCtaColumna

Dim fs              As Scripting.FileSystemObject
Dim xlAplicacion    As Excel.Application
Dim xlLibro         As Excel.Workbook
Dim xlHoja1         As Excel.Worksheet
Dim lbExisteHoja    As Boolean
Dim lilineas        As Integer
Dim i               As Integer
Dim glsArchivo      As String
Dim lsNomHoja       As String
Dim Poscamb         As Currency
Dim cuActivo        As String 'NAGL Cambio de Currency a String
Dim cuPasivo        As String 'NAGL Cambio de Currency a String
Dim cuActivPasiv    As String, cuTipCamb As String, cuPromEfectAnt As String 'NAGL 20170718
Dim lsTotal()       As String 'NAGL 20170718
Dim lsCadena()      As String 'NAGL 20170718
Dim liLineasInicio   As Long 'NAGL 20170718
Dim Cant            As Long 'NAGL 20170718
Dim lenCell         As Integer 'NAGL 20170718
Dim CellParam       As String 'NAGL 20170718
Dim cCtaCont        As String 'NAGL 20170718
Dim lsMovNro        As String 'NAGL 20170725
Dim nActivo As Double, nPasivo As Double 'NAGL 20170912

ReDim lsTotal(2)
ReDim lsCadena(2)

PB1.Min = 0
PB1.Max = 14
PB1.value = 0
PB1.Visible = True

If (pnNivelDet = 0) Then
    pnNivelDet = 1
End If

PB1.value = 1
Set RSTEMP = oclsCtaCont.ListarCtaContPosCamb(pnNivelDet)

    If RSTEMP Is Nothing Then
        MsgBox "No exite informacion para imprimir", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If rs Is Nothing Then
        MsgBox "No exite informacion para imprimir", vbInformation, "Aviso"
        Exit Sub
    End If
    
    glsArchivo = "Reporte Posición Cambiaria" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 100
    xlHoja1.PageSetup.Orientation = xlLandscape

     lbExisteHoja = False
     lsNomHoja = "PosCam"
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

     xlAplicacion.Range("A1:A1").ColumnWidth = 15
     xlAplicacion.Range("B1:B1").ColumnWidth = 20
     xlAplicacion.Range("c1:c1").ColumnWidth = 15
     xlAplicacion.Range("D1:D1").ColumnWidth = 25
     xlAplicacion.Range("E1:E1").ColumnWidth = 15
     xlAplicacion.Range("F1:F1").ColumnWidth = 15

     xlAplicacion.Range("A1:Z10000").Font.Size = 9
     xlAplicacion.Range("A1:Z10000").Font.Name = "Century Gothic"

     xlHoja1.Cells(1, 1) = "REPORTE DE POSICIÓN CAMBIARIA"
     xlHoja1.Cells(2, 1) = "CMAC MAYNAS"
     xlHoja1.Cells(5, 1) = "CONTROL LIMITE INTERNO"
     
     xlHoja1.Cells(2, 4) = "FECHA REPORTE:"
     xlHoja1.Cells(2, 5) = pdFecha 'Format(pdFecha, "dd/mm/yyyy")
     xlHoja1.Cells(3, 4) = "PE mes anterior:"
     xlHoja1.Cells(3, 5) = txtPatrimonioSoles
     xlHoja1.Cells(4, 4) = "Activo - Pasivo:"
     xlHoja1.Cells(4, 4).Font.Bold = True
     xlHoja1.Cells(5, 4) = "Posición Cambiaria %:"
     xlHoja1.Cells(5, 4).Font.Bold = True
     xlHoja1.Cells(6, 4) = "T.C. SBS:"
     xlHoja1.Cells(6, 5) = Format(pnTipoCambio, "#,##0.000")
'            xlHoja1.Cells(7, 4) = "Limite: SobreVenta - 12.50%; SobreCompra 50%"
     
     PB1.value = 3
     xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 5)).HorizontalAlignment = xlCenter 'Titulo Princ.
     xlHoja1.Range(xlHoja1.Cells(2, 4), xlHoja1.Cells(6, 4)).HorizontalAlignment = xlRight
     xlHoja1.Range(xlHoja1.Cells(2, 5), xlHoja1.Cells(6, 5)).HorizontalAlignment = xlCenter
                 
     xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 5)).Font.Bold = True
     xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 3)).Font.Bold = True
     xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 3)).Merge True
     xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 5)).Merge True
          
     xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 3)).Font.Bold = True
     
     'MIOL 20130723; SEGUN ERS088-2013 OBJ. B *******************************
     xlHoja1.Cells(2, 7) = "Limite"
     xlHoja1.Cells(2, 7).ColumnWidth = 12
     xlHoja1.Cells(2, 8) = "SobreCompra"
     xlHoja1.Cells(2, 8).ColumnWidth = 13
     xlHoja1.Cells(2, 9) = "SobreVenta"
     xlHoja1.Cells(2, 9).ColumnWidth = 13
     
     xlHoja1.Range(xlHoja1.Cells(2, 7), xlHoja1.Cells(2, 9)).HorizontalAlignment = xlCenter
     xlHoja1.Range(xlHoja1.Cells(2, 7), xlHoja1.Cells(2, 9)).Font.Bold = True
     xlHoja1.Range(xlHoja1.Cells(2, 7), xlHoja1.Cells(2, 9)).Borders.LineStyle = 1
     xlHoja1.Range(xlHoja1.Cells(2, 7), xlHoja1.Cells(2, 9)).Interior.ColorIndex = 15 'NAGL 20170718
     PB1.value = 5
     Set rsLim = oRepCtaColumna.GetLimitePosCamb()
     lilineas = 3
     Do Until rsLim.EOF
         xlHoja1.Cells(lilineas, 7) = rsLim(1)
         xlHoja1.Range(xlHoja1.Cells(lilineas, 7), xlHoja1.Cells(lilineas, 7)).HorizontalAlignment = xlCenter 'NAGL 20170718
         xlHoja1.Cells(lilineas, 7).Borders.LineStyle = 1
         xlHoja1.Cells(lilineas, 7).Font.Bold = True
         xlHoja1.Cells(lilineas, 8) = rsLim(2)
         xlHoja1.Range(xlHoja1.Cells(lilineas, 8), xlHoja1.Cells(lilineas, 8)).HorizontalAlignment = xlCenter 'NAGL 20170718
         xlHoja1.Cells(lilineas, 8).NumberFormat = "#,###0.00"
         xlHoja1.Cells(lilineas, 8).Borders.LineStyle = 1
         xlHoja1.Cells(lilineas, 9) = rsLim(3)
         xlHoja1.Range(xlHoja1.Cells(lilineas, 9), xlHoja1.Cells(lilineas, 9)).HorizontalAlignment = xlCenter 'NAGL 20170718
         xlHoja1.Cells(lilineas, 9).NumberFormat = "#,###0.00"
         xlHoja1.Cells(lilineas, 9).Borders.LineStyle = 1
         lilineas = lilineas + 1
         rsLim.MoveNext
     Loop
     Set rsLim = Nothing
     'END MIOL **************************************************************
     
     lilineas = 8
     
     xlHoja1.Cells(lilineas, 1) = "CUENTA CONTABLE"
     xlHoja1.Cells(lilineas, 2) = "DESCRIPCION"
     xlHoja1.Cells(lilineas, 3) = "SALDO ME ACUMULADO"
     xlHoja1.Cells(lilineas, 4) = "SALDO ME ACUMULADO (" & Format(DateAdd("d", -1, pdFecha), "dd/mm/yyyy") & ")"
     xlHoja1.Cells(lilineas, 5) = "DIFERENCIA"
                 
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).HorizontalAlignment = xlCenter
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).VerticalAlignment = xlCenter
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas + 5, 1)).Merge True
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).EntireRow.AutoFit
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).WrapText = True

     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).Font.Bold = True
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).Borders.LineStyle = 1
     xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).Interior.ColorIndex = 35
    
    PB1.value = 7
    lilineas = lilineas + 1
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oCtaIf = New NCajaCtaIF
    Do Until RSTEMP.EOF
            xlHoja1.Cells(lilineas, 1) = RSTEMP!cCtaContCod
            xlHoja1.Cells(lilineas, 1).HorizontalAlignment = xlRight
            xlHoja1.Cells(lilineas, 2) = RSTEMP!cCtaContDesc
            xlHoja1.Cells(lilineas, 2).ColumnWidth = 90

            If Len(RSTEMP!cCtaContCod) = pnNivelDet Then
                Set rs = oCtaIf.GetSaldoMEPosCambiariaNewYBalanc("PC", pdFecha, RSTEMP!cCtaContCod, lsMovNro)

                     xlHoja1.Cells(lilineas, 3) = rs!nCtaSaldoImporteMEActual
                     xlHoja1.Cells(lilineas, 3).NumberFormat = "#,###0.00"
                     lsTotal(1) = xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 3)).Address(False, False) 'NAGL 20170718
                     
                     xlHoja1.Cells(lilineas, 4) = rs!nCtaSaldoImporteMEAnterior
                     xlHoja1.Cells(lilineas, 4).NumberFormat = "#,###0.00"
                     lsTotal(2) = xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas, 4)).Address(False, False) 'NAGL 20170718
                     
                     xlHoja1.Cells(lilineas, 5).Formula = "=" & lsTotal(1) & "-" & lsTotal(2) 'NAGL 20170718
                     xlHoja1.Cells(lilineas, 5).NumberFormat = "#,###0.00"
                     'xlHoja1.Cells(liLineas, 5) = rs(0) - rs(1)
                     'xlHoja1.Cells(liLineas, 5).NumberFormat = "#,###0.00"
            End If
            
           If RSTEMP!cCtaContCod = "1" Or RSTEMP!cCtaContCod = "2" Then
                xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 5)).Interior.ColorIndex = 44
           End If
        lilineas = lilineas + 1
        RSTEMP.MoveNext
     Loop
     
     ExcelCuadro xlHoja1, 1, 9, 5, CCur(lilineas - 1) 'NAGL 20170719
    '****************************************************************************NAGL 20170719
 
    lsTotal(1) = ""
    lsTotal(2) = ""
    lilineas = lilineas - 1
    liLineasInicio = 9
    
    PB1.value = 9
    Do While lilineas >= liLineasInicio
            If xlHoja1.Cells(lilineas, 3) = "" Then
                    lenCell = Len(xlHoja1.Cells(lilineas, 1))
                    lenCell = Len(xlHoja1.Cells(lilineas, 1)) + IIf(Len(xlHoja1.Cells(lilineas, 1)) = 1, 1, 2)
                    CellParam = xlHoja1.Cells(lilineas, 1) 'Parametro de Inicio para comparar la celda en cuestión
                    Cant = 0
                    Do While (Mid(xlHoja1.Cells(lilineas + Cant, 1), 1, Len(CellParam)) = CellParam)
                            If (Len(xlHoja1.Cells(lilineas + Cant, 1)) = lenCell) Then
                                lsTotal(1) = xlHoja1.Range(xlHoja1.Cells(lilineas + Cant, 3), xlHoja1.Cells(lilineas + Cant, 3)).Address(False, False)
                                lsCadena(1) = lsCadena(1) & lsTotal(1) & ","
                                lsTotal(2) = xlHoja1.Range(xlHoja1.Cells(lilineas + Cant, 4), xlHoja1.Cells(lilineas + Cant, 4)).Address(False, False)
                                lsCadena(2) = lsCadena(2) & lsTotal(2) & ","
                            End If
                            Cant = Cant + 1
                    Loop
                    If (lsCadena(1) <> "") Then
                             lsCadena(1) = "(" & Mid(lsCadena(1), 1, Len(lsCadena(1)) - 1) & ")"
                             lsCadena(2) = "(" & Mid(lsCadena(2), 1, Len(lsCadena(2)) - 1) & ")"
                             xlHoja1.Cells(lilineas, 3).Formula = "=" & "Sum" & lsCadena(1)
                             xlHoja1.Cells(lilineas, 3).NumberFormat = "#,###0.00"
                             xlHoja1.Cells(lilineas, 4).Formula = "=" & "Sum" & lsCadena(2)
                             xlHoja1.Cells(lilineas, 4).NumberFormat = "#,###0.00"
                             Call oCtaIf.setSaldoMEPosCambiariaFormulas(pdFecha, CellParam, xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 4), lsMovNro)
                     Else
                             cCtaCont = CellParam
                             Set rs = oCtaIf.GetSaldoMEPosCambiariaNewYBalanc("PC", pdFecha, cCtaCont, lsMovNro)
                             xlHoja1.Cells(lilineas, 3) = rs!nCtaSaldoImporteMEActual
                             xlHoja1.Cells(lilineas, 3).NumberFormat = "#,###0.00"
                             xlHoja1.Cells(lilineas, 4) = rs!nCtaSaldoImporteMEAnterior
                             xlHoja1.Cells(lilineas, 4).NumberFormat = "#,###0.00"
                     End If
                             xlHoja1.Cells(lilineas, 5).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 3)).Address(False, False) & "-" & xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas, 4)).Address(False, False)
                             xlHoja1.Cells(lilineas, 5).NumberFormat = "#,###0.00"
                             
                     If CellParam = "1" Then
                          cuActivo = xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 3)).Address(False, False)
                     End If
                     If CellParam = "2" Then
                          cuPasivo = xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 3)).Address(False, False)
                     End If
                     lsTotal(1) = ""
                     lsCadena(1) = ""
                     lsTotal(2) = ""
                     lsCadena(2) = ""
            End If
        lilineas = lilineas - 1
    Loop
    
        Set rs = Nothing
        If (cuActivo <> "" And cuPasivo <> "") Then
            xlHoja1.Cells(4, 5).Formula = "=" & "+" & cuActivo & "-" & cuPasivo
            xlHoja1.Cells(4, 5).NumberFormat = "#,###0.00"
        Else
            Set rs = oCtaIf.GetSaldoMEPosCambiariaNewYBalanc("PC", pdFecha, "1", lsMovNro)
            nActivo = rs!nCtaSaldoImporteMEActual
            Set rs = oCtaIf.GetSaldoMEPosCambiariaNewYBalanc("PC", pdFecha, "2", lsMovNro)
            nPasivo = rs!nCtaSaldoImporteMEActual
            xlHoja1.Cells(4, 5) = nActivo - nPasivo
            xlHoja1.Cells(4, 5).NumberFormat = "#,###0.00"
        End If 'NAGL 20170913
        
        cuActivPasiv = xlHoja1.Range(xlHoja1.Cells(4, 5), xlHoja1.Cells(4, 5)).Address(False, False)
        cuTipCamb = xlHoja1.Range(xlHoja1.Cells(6, 5), xlHoja1.Cells(6, 5)).Address(False, False)
        cuPromEfectAnt = xlHoja1.Range(xlHoja1.Cells(3, 5), xlHoja1.Cells(3, 5)).Address(False, False)
        
        xlHoja1.Cells(5, 5).Formula = "=" & "(" & cuActivPasiv & "*" & cuTipCamb & ")" & "/" & cuPromEfectAnt & "*" & "100"
        xlHoja1.Cells(5, 5).NumberFormat = "#,###0.00"
        Poscamb = xlHoja1.Cells(5, 5)
        '**********************************************************************END NAGL
        PB1.value = 12
        'MIOL 20130723, SEGUN ERS088-2013 OBJ. B *******************************
        If Poscamb < 0 Then 'SobreVenta
        
            Dim nRegSC As Currency
            Dim nIntSC As Currency
            Dim nTemSC As Currency

            nRegSC = xlHoja1.Cells(3, 8)
            nIntSC = xlHoja1.Cells(4, 8)
            nTemSC = xlHoja1.Cells(5, 8)

            Poscamb = Poscamb * -1
            If nTemSC - Poscamb > 0 And nTemSC - Poscamb <= 1 Then
                xlHoja1.Cells(6, 1) = "CERCANO A ALERTA TEMPRANA"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nTemSC = Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE DE ALERTA TEMPRANA"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf (nTemSC - Poscamb < 0 And nTemSC - Poscamb > -1) Or (nTemSC < Poscamb And Poscamb < nIntSC) Then
                xlHoja1.Cells(6, 1) = "LIMITE DE ALERTA TEMPRANA SUPERADO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nIntSC = Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE INTERNO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nIntSC < Poscamb And nIntSC < nRegSC And Poscamb < nRegSC Then
                xlHoja1.Cells(6, 1) = "LIMITE INTERNO SUPERADO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nRegSC = Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE REGULATORIO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nRegSC < Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE REGULATORIO SUPERADO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            Else
                xlHoja1.Cells(6, 1) = "ACEPTABLE"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbBlue
            End If
        Else 'SobreCompra
            Dim nRegSV As Currency
            Dim nIntSV As Currency
            Dim nTemSV As Currency

            nRegSV = xlHoja1.Cells(3, 9)
            nIntSV = xlHoja1.Cells(4, 9)
            nTemSV = xlHoja1.Cells(5, 9)
            
            If nTemSV - Poscamb > 0 And nTemSV - Poscamb <= 1 Then
                xlHoja1.Cells(6, 1) = "CERCANO A ALERTA TEMPRANA"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nTemSV = Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE DE ALERTA TEMPRANA"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf (nTemSV - Poscamb < 0 And nTemSV - Poscamb > -1) Or (nTemSV < Poscamb And Poscamb < nIntSV) Then
                xlHoja1.Cells(6, 1) = "LIMITE DE ALERTA TEMPRANA SUPERADO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nIntSV = Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE INTERNO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nIntSV < Poscamb And nIntSV < nRegSV And Poscamb < nRegSV Then
                xlHoja1.Cells(6, 1) = "LIMITE INTERNO SUPERADO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nRegSV = Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE REGULATORIO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            ElseIf nRegSV < Poscamb Then
                xlHoja1.Cells(6, 1) = "LIMITE REGULATORIO SUPERADO"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbRed
            Else
                xlHoja1.Cells(6, 1) = "ACEPTABLE"
                xlHoja1.Cells(6, 1).Font.Bold = True
                xlHoja1.Cells(6, 1).Cells.Font.Color = vbBlue
            End If
        End If
'        If Poscamb < -12.5 Then
'            xlHoja1.Cells(7, 1) = "EXCESO LIMITE SOBRE VENTA"
'            xlHoja1.Cells(7, 1).Font.Bold = True
'            xlHoja1.Cells(7, 1).Cells.Font.Color = vbRed
'        ElseIf Poscamb > 50# Then
'            xlHoja1.Cells(7, 1) = "EXCESO LIMITE SOBRE COMPRA"
'            xlHoja1.Cells(7, 1).Font.Bold = True
'            xlHoja1.Cells(7, 1).Cells.Font.Color = vbRed
'        Else
'            xlHoja1.Cells(7, 1) = "ACEPTABLE"
'            xlHoja1.Cells(7, 1).Font.Bold = True
'            xlHoja1.Cells(7, 1).Cells.Font.Color = vbBlue
'        End If
        'END MIOL **************************************************************

       
        Set RSTEMP = Nothing
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        'ExcelEnd App.path & "\Spooler\" & glsArchivo, xlAplicacion, xlLibro, xlHoja1
        PB1.value = 14
        'Set xlAplicacion = Nothing
        'Set xlLibro = Nothing
        'Set xlHoja1 = Nothing
        'MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        'Call CargaArchivo(glsArchivo, App.path & "\SPOOLER\")
        xlAplicacion.Visible = True
        xlAplicacion.Windows(1).Visible = True
        Set xlAplicacion = Nothing
        Set xlHoja1 = Nothing
        PB1.Visible = False
Set oCtaIf = Nothing
    Exit Sub
ReporteAdeudadosVinculadosErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub

Private Sub cmdPatrimonio_Click()
Set oDbalanceCont = New DbalanceCont
Set rsPatrimonio = New ADODB.Recordset

Dim oForm As New frmRegPatrimonioEfectivo
CentraForm oForm
oForm.Show 1, Me
'*****NAGL 20170718
If txtFecha.Text = "__/__/____" Or txtFecha.Text = "" Then
    pnTipoIng = 3 'Balance
    pdFechaControl = CalculaBalance
    Call CalculaPatrimonioEfectivo(pnTipoIng, pdFechaControl)
Else
    pnTipoIng = 1 'Con Fecha Ingresada
    pdFechaControl = txtFecha
    Call CalculaPatrimonioEfectivo(pnTipoIng, pdFechaControl)
End If 'NAGL 20170718
'Se cambio de txtFecha a pdFechaControl
'    If Year(pdFechaControl) < Year(gdFecSis) Then
'            Set rsPatrimonio = oDbalanceCont.recuperarPatrimonioEfectivoMesAnio(Month(pdFechaControl), Year(pdFechaControl))
'            If Not rsPatrimonio.BOF And Not rsPatrimonio.EOF Then
'                    txtPatrimonioSoles = Format(rsPatrimonio.Fields(1), "#,##0.00###")
'                    txtMes = UCase(rsPatrimonio.Fields(2))
'            End If
'            txtTpoCambio = TipoCambioCierre(Year(DateAdd("m", -1, pdFechaControl)), Month(DateAdd("m", -1, pdFechaControl)))
'            nDolares = txtPatrimonioSoles / txtTpoCambio
'            txtPatrimonioDolares = Format(nDolares, "#,##0.00]")
'    ElseIf Year(pdFechaControl) = Year(gdFecSis) Then
'        If Month(pdFechaControl) <= Month(gdFecSis) Then
'            Set rsPatrimonio = oDbalanceCont.recuperarPatrimonioEfectivoMesAnio(Month(pdFechaControl), Year(pdFechaControl))
'            If Not rsPatrimonio.BOF And Not rsPatrimonio.EOF Then
'                    txtPatrimonioSoles = Format(rsPatrimonio.Fields(1), "#,##0.00###")
'                    txtMes = UCase(rsPatrimonio.Fields(2))
'            End If
'            txtTpoCambio = TipoCambioCierre(Year(DateAdd("m", -1, pdFechaControl)), Month(DateAdd("m", -1, pdFechaControl)))
'            nDolares = txtPatrimonioSoles / txtTpoCambio
'            txtPatrimonioDolares = Format(nDolares, "#,##0.00")
'        Else
'            MsgBox " Fecha Ingresada es Mayor a la Fecha Actual ...! ", vbCritical, "Error Fecha"
'        End If
'    End If Comentado by NAGL 20170918
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub 'NAGL 20170719

Private Sub txtFecha_KeyPress(keyAscii As Integer)
pnTipoIng = 1
If keyAscii = 13 Then
      If ValFecha(txtFecha) = True Then
        If ValidaFecha(txtFecha) Then
           cboDig.SetFocus
           Call CalculaPatrimonioEfectivo(pnTipoIng, txtFecha)
        End If
      End If
End If
End Sub 'NAGL 20170719

Private Sub txtFecha_LostFocus()
    pnTipoIng = 1
    If ValFecha(txtFecha) = True Then
        If ValidaFecha(txtFecha) Then
            Call CalculaPatrimonioEfectivo(pnTipoIng, txtFecha)
        End If
    End If
End Sub 'Para que calcule automaticamente el patrimonio y balance sin hacer enter

Private Sub cboDig_KeyPress(keyAscii As Integer)
    If keyAscii = 13 Then
         cmdGenerar.SetFocus
    End If
End Sub 'NAGL 20170719

Private Sub txtAnio_GotFocus()
    fEnfoque txtAnio
End Sub 'NAGL 20170719

Private Sub txtAnio_KeyPress(keyAscii As Integer)
    pnTipoIng = 3
    keyAscii = NumerosEnteros(keyAscii)
    If keyAscii = 13 Then
         cmdGenerar.SetFocus
         pdFechaBalance = CalculaBalance
         Call CalculaPatrimonioEfectivo(pnTipoIng, pdFechaBalance)
    End If
End Sub

Private Sub txtAnio_LostFocus()
    pnTipoIng = 3
    pdFechaBalance = CalculaBalance
    Call CalculaPatrimonioEfectivo(pnTipoIng, pdFechaBalance)
End Sub 'NAGL 20170719
Private Sub cboMes_KeyPress(keyAscii As Integer)
    If keyAscii = 13 Then
         txtAnio.SetFocus
    End If
End Sub 'NAGL 20170719

Private Sub txtPatrimonioDolares_GotFocus()
    fEnfoque txtPatrimonioDolares
End Sub

Private Sub txtPatrimonioDolares_KeyPress(keyAscii As Integer)
    keyAscii = NumerosDecimales(txtPatrimonioDolares, keyAscii)
End Sub

Private Sub txtPatrimonioDolares_LostFocus()
    txtPatrimonioDolares = Format(txtPatrimonioDolares, gsFormatoNumeroView)
End Sub

Private Sub txtPatrimonioSoles_GotFocus()
    fEnfoque txtPatrimonioSoles
End Sub

Private Sub txtPatrimonioSoles_KeyPress(keyAscii As Integer)
    keyAscii = NumerosDecimales(txtPatrimonioSoles, keyAscii)
End Sub

Private Sub txtPatrimonioSoles_LostFocus()
    txtPatrimonioSoles = Format(txtPatrimonioSoles, gsFormatoNumeroView)
End Sub

Private Sub txtTpoCambio_GotFocus()
    fEnfoque txtTpoCambio
End Sub

Private Sub txtTpoCambio_KeyPress(keyAscii As Integer)
    keyAscii = NumerosDecimales(txtTpoCambio, keyAscii)
End Sub

Private Sub txtTpoCambio_LostFocus()
    txtTpoCambio = Format(txtTpoCambio, gsFormatoNumeroView)
End Sub

'Private Sub txtFecha_LostFocus()
'Set oDbalanceCont = New DbalanceCont
'Set rsPatrimonio = New ADODB.Recordset
'    If Year(txtfecha) < Year(gdFecSis) Then
'            Set rsPatrimonio = oDbalanceCont.recuperarPatrimonioEfectivoMesAnio(Month(txtfecha), Year(txtfecha))
'            If Not rsPatrimonio.BOF And Not rsPatrimonio.EOF Then
'                    txtPatrimonioSoles = Format(rsPatrimonio.Fields(1), "#,##0.00###")
'                    txtMes = UCase(rsPatrimonio.Fields(2))
'            End If
'            txtTpoCambio = TipoCambioCierre(Year(DateAdd("m", -1, txtfecha)), Month(DateAdd("m", -1, txtfecha)))
'            nDolares = txtPatrimonioSoles / txtTpoCambio
'            txtPatrimonioDolares = Format(nDolares, "#,##0.00")
'    ElseIf Year(txtfecha) = Year(gdFecSis) Then
'        If Month(txtfecha) <= Month(gdFecSis) Then
'            Set rsPatrimonio = oDbalanceCont.recuperarPatrimonioEfectivoMesAnio(Month(txtfecha), Year(txtfecha))
'            If Not rsPatrimonio.BOF And Not rsPatrimonio.EOF Then
'                    txtPatrimonioSoles = Format(rsPatrimonio.Fields(1), "#,##0.00###")
'                    txtMes = UCase(rsPatrimonio.Fields(2))
'
'                    txtTpoCambio = TipoCambioCierre(Year(DateAdd("m", -1, txtfecha)), Month(DateAdd("m", -1, txtfecha)))
'                    nDolares = txtPatrimonioSoles / txtTpoCambio
'                    txtPatrimonioDolares = Format(nDolares, "#,##0.00")
'            Else 'AGREGADO MIOL 20130502, SOLICITUD DEL USUARIO KARU
'                    txtTpoCambio = TipoCambioCierre(Year(DateAdd("m", -1, txtfecha)), Month(DateAdd("m", -1, txtfecha)))
'                    nDolares = txtPatrimonioSoles / txtTpoCambio
'                    txtPatrimonioDolares = Format(nDolares, "#,##0.00")
'            End If
'        End If
'    Else
'            MsgBox " Fecha Ingresada es Mayor a la Fecha Actual ...! ", vbCritical, "Error Fecha"
'    End If
'End Sub

