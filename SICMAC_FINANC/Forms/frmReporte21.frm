VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReporte21 
   Caption         =   "Reporte 21"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   Icon            =   "frmReporte21.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   7230
      TabIndex        =   16
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8085
      Begin VB.ComboBox cboMes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmReporte21.frx":030A
         Left            =   720
         List            =   "frmReporte21.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   330
         Width           =   1815
      End
      Begin VB.TextBox txtAnio 
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
         Height          =   330
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   7
         Top             =   325
         Width           =   1095
      End
      Begin VB.TextBox txtTC 
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
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   7050
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox txtPatrEfec 
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
         Left            =   1560
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtPatrDol 
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   390
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   2760
         TabIndex        =   13
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio:"
         Height          =   195
         Left            =   5765
         TabIndex        =   12
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label lblPatr 
         Caption         =   "Patr.Efectivo S/:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1245
         Width           =   1335
      End
      Begin VB.Label lblPatrDol 
         Caption         =   "Patr.Efectivo $:"
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   1245
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Información del Patrimonio"
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
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vinculados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   8085
      Begin VB.CommandButton cmdActualizarDataRRHH 
         Caption         =   "RRHH"
         Height          =   375
         Left            =   6960
         TabIndex        =   19
         Top             =   1360
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   6960
         TabIndex        =   2
         Top             =   320
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   6960
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
      Begin Sicmact.FlexEdit FERelVin 
         Height          =   3015
         Left            =   150
         TabIndex        =   18
         Top             =   315
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   5318
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Código-Nombres-Vínculo-VinculoId"
         EncabezadosAnchos=   "350-1300-3300-1300-0"
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3-X"
         ListaControles  =   "0-1-0-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-C"
         FormatosEdit    =   "0-0-0-0-0"
         CantEntero      =   20
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   5580
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmReporte21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim nItem As Integer
Dim lsPersCod As String
Dim lsPersCodAnt As String
Dim lsVinculo As String
Dim FERelVinNoMoverdeFila As Integer
Dim oCon As DConecta
Dim sSql As String
Dim nLin As Integer
Dim NumeroVincu As Integer
Dim lnTipo As Integer
Dim oDbalanceCont As New DbalanceCont 'NAGL ERS074-2017 20171209

Private Sub Form_Load()
'************Agregado by NAGL ERS074-2017 20171209************'
    Dim pdFechaMesAnt As Date
    Dim oGen As New DGeneral
    Dim pnTipo As Integer
    pnTipo = 3
    FERelVin.lbEditarFlex = False
    Set rs = oGen.GetConstante(1010)
    While Not rs.EOF
       cboMes.AddItem rs.Fields(0) & space(50) & rs.Fields(1)
       rs.MoveNext
    Wend
    pdFechaMesAnt = DateAdd("d", -Day(gdFecSis), gdFecSis)
    Me.txtAnio.Text = Year(pdFechaMesAnt)
    cboMes.ListIndex = CInt(Month(pdFechaMesAnt)) - 1
    
    txtTC.Text = oDbalanceCont.ObtenerTipoCambioCierreNew(pdFechaMesAnt) 'Se quitó "TipoAct", para que tome T.C cierre NAGL 20180125
    txtTC.Text = IIf(txtTC = 0, 0, Format(txtTC, "#,##0.000"))
    txtPatrEfec.Text = Format(CalculaPatrimonioEfectivo(pdFechaMesAnt), "#,##0.00")
    CalculaPatrimonioDolares ("Sist")
    Call CargarRelInst
    Call MuestraVinculados(pdFechaMesAnt)
'**************END NAGL ERS074-2017*****************************'
'Set rs = RecuperaListaVinculados()
'If rs.RecordCount = 0 Then
'    Exit Sub
'End If
'nLin = 1
'Do While Not rs.EOF
'    FERelVin.AdicionaFila
'    FERelVin.TextMatrix(nLin, 0) = nLin
'    FERelVin.TextMatrix(nLin, 1) = rs!cPersCod
'    FERelVin.TextMatrix(nLin, 2) = rs!cPersNombre
'    FERelVin.TextMatrix(nLin, 3) = rs!cVinculo
'    rs.MoveNext
'    nLin = nLin + 1
'Loop 'Comentado by NAGL ERS074-2017 20171209
End Sub
Private Sub MuestraVinculados(pdFecha As Date, Optional psTipo As String = "") 'NAGL 20190705 Agregó psTipo
Dim rs As New ADODB.Recordset
Dim DGrp As New DGrupoEco
Dim i As Integer

    Set rs = DGrp.ObtenerDatosVinculadosRpte21(pdFecha, psTipo) 'NAGL 20190705 Agregó psTipo
    FERelVin.lbEditarFlex = True
    FERelVin.Clear
    FormateaFlex FERelVin
    
    If Not (rs.BOF And rs.EOF) Then
        For i = 1 To rs.RecordCount
            FERelVin.AdicionaFila
            FERelVin.TextMatrix(i, 1) = rs!cPersCod
            FERelVin.TextMatrix(i, 2) = PstaNombre(rs!cPersNombre)
            FERelVin.TextMatrix(i, 3) = rs!cVinculo
            FERelVin.TextMatrix(i, 4) = rs!cVinculoId
            rs.MoveNext
        Next i
    End If
End Sub 'NAGL ERS074-2017 20171209

Private Function LlenarRsFERelVinDet(ByVal feControl As FlexEdit) As ADODB.Recordset
 Dim rsVD As New ADODB.Recordset
 Dim nIndex As Integer
  If feControl.Rows >= 2 Then
        If feControl.TextMatrix(nIndex, 1) = "" Then
            Exit Function
        End If
            rsVD.CursorType = adOpenStatic
            rsVD.Fields.Append "cPersCod", adVarChar, 13, adFldIsNullable
            rsVD.Fields.Append "cPersNombre", adVarChar, 300, adFldIsNullable
            rsVD.Fields.Append "cVinculo", adVarChar, 2, adFldIsNullable
            rsVD.Open
            
        For nIndex = 1 To feControl.Rows - 1
            rsVD.AddNew
            rsVD.Fields("cPersCod") = feControl.TextMatrix(nIndex, 1)
            rsVD.Fields("cPersNombre") = feControl.TextMatrix(nIndex, 2)
            rsVD.Fields("cVinculo") = Trim(Right(feControl.TextMatrix(nIndex, 4), 2))
            rsVD.Update
            rsVD.MoveFirst
        Next
    End If
    Set LlenarRsFERelVinDet = rsVD
End Function 'NAGL ERS074-2017

Private Sub cmdGenerar_Click()
    Dim pnTipoCambio As Double
    Dim pnPatrimonio As Currency
    Dim pdFecha As Date
    Dim i, nRows As Integer
    Dim Dgrup As New DGrupoEco
    '****NAGL20190118***********************
    Dim oNContFunciones As New NContFunciones
    Dim oDbalanceCont As New DbalanceCont
    Dim lsMovNro As String
    Dim lbPatrimonioReg As Boolean
    '****END NAGL***************************
    
    If FERelVin.Rows - 1 >= 2 Or FERelVin.TextMatrix(1, 1) <> "" Then
        nRows = FERelVin.Rows - 1
        Do While i <= nRows
            If FERelVin.TextMatrix(i, 1) = "" Or FERelVin.TextMatrix(i, 2) = "" Then
                FERelVin.EliminaFila i
                i = i - 1
                nRows = nRows - 1
            End If
            i = i + 1
        Loop
    pdFecha = CalculaFechaFinMes 'Subido by NAGL 20190709
    If MsgBox("¿Desea Registrar los Vinculados Ingresados al " & pdFecha & " ..", vbInformation + vbYesNo, "Atención") = vbYes Then 'NAGL 20190709 Agregó pdFecha
            If ValidaDatosRep21("Save") Then
                pnTipoCambio = txtTC.Text
                pnPatrimonio = txtPatrEfec.Text
                lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser) 'NAGL 20190118
                lbPatrimonioReg = oDbalanceCont.registrarPatrimonioEfectivo(pnPatrimonio, CInt(Month(pdFecha)), CInt(Year(pdFecha)), lsMovNro, "Rep21")
                If ValidaListaVinc Then
                   Call Dgrup.GuardaDatosVinculadosRpte21(pdFecha, LlenarRsFERelVinDet(FERelVin), lsMovNro) 'NAGL Agregó lsMovNro 20190705
                   MsgBox "Los datos se guardaron satisfactoriamente.", vbOKOnly + vbInformation, "Atención"
                   If MsgBox("¿Desea Generar el Reporte 21 - Financiamiento a Vinculados de la Empresa..", vbInformation + vbYesNo, "Atención") = vbYes Then
                        Call MuestraVinculados(pdFecha)
                        Call GeneraRep21_ClientesVinculados(pdFecha, pnTipoCambio, pnPatrimonio)
                   Else
                        Call MuestraVinculados(pdFecha)
                   End If
                End If
            End If
        End If
    ElseIf FERelVin.TextMatrix(1, 1) = "" Then
        MsgBox "Falta Ingresar Vinculados..!!", vbOKOnly + vbInformation, "Atención"
        Exit Sub
    End If
'Dim lsArchivo   As String
'Dim lbLibroOpen As Boolean
    'lsArchivo = App.path & "\Spooler\Reporte21_" & lsFec & ".xls"
    'lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
    'If lbLibroOpen Then
        'Call GeneraRep21_ClientesVinculados
        'ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
        'CargaArchivo "Reporte21_" & lsFec & ".xls", App.path & "\Spooler"
    'End If Trasladado al Método GeneraRep21_ClientesVinculados by NAGL ERS074-2017
End Sub

Public Function ValidaDatosRep21(Optional psFiltro As String) As Boolean
Dim pdFechaParam As Date
If (txtAnio.Text = "") Then
    MsgBox "Debe Ingresar el año correspondiente..!!", vbInformation, "Aviso"
    txtAnio.SetFocus
    Exit Function
End If
If psFiltro = "Opt" Then
   If (CInt(txtAnio.Text) > Year(gdFecSis) Or (CInt(txtAnio.Text) < 2012)) Then
        MsgBox "Debe Ingresar el año correspondiente..!!", vbInformation, "Aviso"
        txtAnio.SetFocus
        Exit Function
   End If
Else
    If CInt(txtAnio.Text) > Year(gdFecSis) Or CInt(txtAnio.Text) < 2012 Then
       MsgBox "Debe Ingresar el año correspondiente..!!", vbInformation, "Aviso"
       txtAnio.SetFocus
       Exit Function
    ElseIf txtPatrEfec.Text = "" Or txtPatrEfec.Text = "0.00" Or txtPatrEfec.Text = "." Then
       MsgBox "Falta ingresar el Patrimonio Efectivo..!!", vbInformation, "Aviso"
       txtPatrEfec.SetFocus
       Exit Function
    ElseIf txtTC.Text = "" Or txtTC.Text = "0.00" Then
       MsgBox "El Tipo de Cambio es Incorecto..!!", vbInformation, "Aviso"
       txtAnio.SetFocus
       Exit Function
    End If
    If psFiltro = "Save" Then
        pdFechaParam = CalculaFechaFinMes
        If pdFechaParam > gdFecSis Then
             MsgBox "No existe información de Financiamientos con respecto al Periodo Ingresado..!!", vbInformation, "Aviso"
             cboMes.SetFocus
             Exit Function
        End If
    End If
End If
ValidaDatosRep21 = True
End Function

Public Function ValidaListaVinc(Optional optAdd As String)
Dim i, X, Cant, nItem As Integer
Cant = 0
If optAdd = "optAgr" Then
    For i = 1 To CInt(FERelVin.Rows) - 1 'Para Determinar si existen Nombres repetidos con respecto al ultimo Registro Ingresado en la lista
        If i <> CInt(FERelVin.Rows) - 1 Then
            If Trim(FERelVin.TextMatrix(i, 1)) = Trim(FERelVin.TextMatrix(FERelVin.Rows - 1, 1)) Then
                nItem = CInt(FERelVin.Rows) - 1
                MsgBox "El Vinculado " & Trim(FERelVin.TextMatrix(FERelVin.Row, 2)) & ", ya ha sido ingresado !!", vbInformation, "Aviso"
                FERelVin.EliminaFila nItem
                FERelVin.SetFocus
                Exit Function
            End If
        End If
    Next i
Else
For i = 1 To CInt(FERelVin.Rows) - 1 'Para Determinar si existen Nombres repetidos en la lista
        X = i + 1
        ' CInt(FERelVin.Rows) - 1
        Do While X <= CInt(FERelVin.Rows) - 1
            If Trim(FERelVin.TextMatrix(i, 1)) = Trim(FERelVin.TextMatrix(X, 1)) Then
                Cant = Cant + 1
            End If
            If Cant >= 1 Then
                If optAdd <> "Busc" Then
                    MsgBox "No se puede Continuar.. El Vinculado " & Trim(FERelVin.TextMatrix(i, 2)) & ", se ha registrado !!" & CStr(Cant + 1) & " veces !!", vbInformation, "Aviso"
                Else
                    If FERelVin.TextMatrix(i, 2) <> "" Then
                        MsgBox "El Vinculado " & Trim(FERelVin.TextMatrix(i, 2)) & ", ya ha sido ingresado ", vbInformation, "Aviso"
                    End If
                    If X <> CInt(FERelVin.Rows) - 1 Then
                        FERelVin.TextMatrix(X, 1) = ""
                        FERelVin.TextMatrix(X, 2) = ""
                    Else
                        FERelVin.EliminaFila X
                    End If
                End If
                FERelVin.SetFocus
                Exit Function
            End If
            X = X + 1
        Loop
Next i
End If
ValidaListaVinc = True
End Function 'ERS 074-2017 20171209

Private Sub GeneraRep21_ClientesVinculados(Optional pdFecha As Date, Optional pnTipoCambio As Double, Optional nPatrEfectivo As Currency) 'Parámetros agregados by NAGL ERS074-2017 20171209
    Dim lsArchivo   As String 'NAGL ERS074-2017 20171209
    Dim lbLibroOpen As Boolean 'NAGL ERS074-2017 20171209
    
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim nFila As Integer
    Dim nCon As Integer
    Dim nFilaTotal1 As Integer
    Dim nFilaTotal2 As Integer
    Dim TotalArt202 As Double
    Dim nInicio As Integer
    Dim nFilasParam As Integer
    'Dim nPatrEfectivo As Currency
    Dim DGrp As New DGrupoEco 'NAGL ERS074-2017 20171209
    'Dim oAnx As New NEstadisticas
    
    PB1.Min = 0
    PB1.Max = 25
    PB1.value = 1

    lsArchivo = App.path & "\Spooler\Reporte21_" & CStr(Format(pdFecha, "yyyymmdd")) & ".xls"
    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
    
If lbLibroOpen Then
    PB1.Visible = True
    'Calculo el nPatrEfectivo
    'nPatrEfectivo = RecuperaPatrimonioEfectivo(Format(cboMes.ListIndex + 1, "00"), Format(txtAnio.Text, "0000")) 'Val(txtPatrimonio.Text) 'Val(oAnx.GetImporteEstadAnexosMax(gdFecSis, "TOTALREP03", "1"))
    'Set oCon = New DConecta
    'oCon.AbreConexion 'Comentado by NAGL 20171209
    'Adiciona una hoja
    PB1.value = 2
    
    ExcelAddHoja "Rep. N° 21", xlLibro, xlHoja1, True
               
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
        
    xlHoja1.Cells(2, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(2, 10) = "REPORTE Nº 21"
    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(4, 2) = "FINANCIAMIENTOS A VINCULADOS A LA EMPRESA"
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).HorizontalAlignment = xlCenter
   
    xlHoja1.Cells(6, 2) = "Empresa que remite la información: " & gsNomCmac
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(7, 2) = "INFORMACION AL " & Format(pdFecha, "DD MMMM YYYY")
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(9, 2) = "1. Vinculados por el Artículo 202º de la Ley General"
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(9, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(9, 6)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(9, 11)).Font.Bold = True
          
    'If Me.txtTC.Text = "" Then
        'MsgBox "Debe ingresar el tipo de cambio"
        'Exit Sub
    'End If
    'sSql = "Exec stp_sel_DatosReporte21 " & Me.txtTC.Text
    'pnTipoCambio = CDbl(Me.txtTC.Text)
     '************************'Comentado by NAGL 20171209
    PB1.value = 5
    Set rs = DGrp.ObtenerDatosReporte21(pnTipoCambio, pdFecha)
    PB1.value = 7
    
    If Not (rs.BOF And rs.EOF) Then 'Cambio by NAGL 20171209
        nFila = 10
        nInicio = nFila + 1
        xlHoja1.Cells(nFila, 1) = "Nº"
        xlHoja1.Cells(nFila, 2) = "Cod"
        xlHoja1.Cells(nFila, 3) = "Nombre/razón/"
        xlHoja1.Cells(nFila, 4) = "CIIU"
        xlHoja1.Cells(nFila, 5) = "Domicilio"
        xlHoja1.Cells(nFila, 6) = "Tipo de"
        xlHoja1.Cells(nFila, 7) = "Tipo de doc."
        xlHoja1.Cells(nFila, 8) = "Num."
        xlHoja1.Cells(nFila, 9) = "RUC"
        xlHoja1.Cells(nFila, 10) = "Descripcion de la vinculacion"
        xlHoja1.Range(xlHoja1.Cells(nFila, 10), xlHoja1.Cells(nFila, 12)).MergeCells = True
        xlHoja1.Cells(nFila, 13) = "Créditos"
        xlHoja1.Cells(nFila, 14) = "Inversiones"
        xlHoja1.Cells(nFila, 15) = "Contingentes"
        xlHoja1.Cells(nFila, 16) = "Arrendamiento"
        xlHoja1.Cells(nFila, 17) = "Derivados" 'JOEP
        xlHoja1.Cells(nFila, 18) = "Otros" 'JOEP
        xlHoja1.Cells(nFila, 19) = "Total"
        'xlHoja1.Cells(nFila, 17) = "Total" 'Comento JOEP
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 2) = "SBS"
        xlHoja1.Cells(nFila, 3) = "denominación"
        xlHoja1.Cells(nFila, 6) = "persona"
        xlHoja1.Cells(nFila, 7) = "de indentidad"
        xlHoja1.Cells(nFila, 8) = "Documento de"
        xlHoja1.Cells(nFila, 10) = "Propiedad"
        xlHoja1.Cells(nFila, 11) = "Propiedad"
        xlHoja1.Cells(nFila, 12) = "Gestion"
        xlHoja1.Cells(nFila, 13) = "Directos" 'JOEP
        xlHoja1.Cells(nFila, 16) = "financiero"
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 3) = "Social"
        xlHoja1.Cells(nFila, 8) = "identidad"
        xlHoja1.Cells(nFila, 10) = "Directa"
        xlHoja1.Cells(nFila, 11) = "Indirecta"
        
        'xlHoja1.Range(xlHoja1.Cells(nFila - 2, 1), xlHoja1.Cells(nFila, 17)).HorizontalAlignment = xlCenter'Comentado JOEP
        'xlHoja1.Range(xlHoja1.Cells(nFila - 2, 1), xlHoja1.Cells(nFila, 17)).Font.Bold = True 'Comentado JOEP
        xlHoja1.Range(xlHoja1.Cells(nFila - 2, 1), xlHoja1.Cells(nFila, 19)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(nFila - 2, 1), xlHoja1.Cells(nFila, 19)).Font.Bold = True
        
        'ExcelCuadro xlHoja1, 1, nFila - 2, 17, nFila 'Comentado JOEP
        ExcelCuadro xlHoja1, 1, nFila - 2, 19, nFila
        
        nCon = 1
        Do While Not rs.EOF
            If rs!Creditos + rs!CartaFianza <> 0 Then
                nFila = nFila + 1
                xlHoja1.Cells(nFila, 1) = nCon
                xlHoja1.Cells(nFila, 2) = rs!cPersCodSBS
                xlHoja1.Cells(nFila, 3) = rs!cPersNombre
                xlHoja1.Cells(nFila, 4) = rs!cPersCIIU
                xlHoja1.Cells(nFila, 5) = rs!cPersDireccDomicilio
                xlHoja1.Cells(nFila, 6) = rs!cConsDescripcion
                xlHoja1.Cells(nFila, 7) = rs!TipDoc
                xlHoja1.Cells(nFila, 8) = rs!Documento
                xlHoja1.Cells(nFila, 9) = rs!RUC
                xlHoja1.Cells(nFila, 13) = rs!Creditos
                xlHoja1.Cells(nFila, 15) = rs!CartaFianza
                'xlHoja1.Cells(nFila, 17) = rs!nDerivados 'JOEP
                'xlHoja1.Cells(nFila, 18) = rs!nOtros 'JOEP
                'xlHoja1.Cells(nFila, 17) = rs!Creditos + rs!CartaFianza 'Comentado JOEP
                xlHoja1.Cells(nFila, 19) = rs!Creditos + rs!CartaFianza '+ rs!nDerivados + rs!nOtros
                xlHoja1.Range(xlHoja1.Cells(nFila, 13), xlHoja1.Cells(nFila, 13)).NumberFormat = "#,00.00"
                xlHoja1.Range(xlHoja1.Cells(nFila, 15), xlHoja1.Cells(nFila, 15)).NumberFormat = "#,00.00"
                xlHoja1.Range(xlHoja1.Cells(nFila, 17), xlHoja1.Cells(nFila, 17)).NumberFormat = "#,00.00" 'JOEP
                xlHoja1.Range(xlHoja1.Cells(nFila, 18), xlHoja1.Cells(nFila, 18)).NumberFormat = "#,00.00" 'JOEP
                xlHoja1.Range(xlHoja1.Cells(nFila, 19), xlHoja1.Cells(nFila, 19)).NumberFormat = "#,00.00"
    
                nCon = nCon + 1
            End If
            rs.MoveNext
        Loop
        
        'ExcelCuadro xlHoja1, 1, nFila - 2, 17, nFila, , True'Comentado JOEP
        If nCon >= 2 Then
            ExcelCuadro xlHoja1, 1, nFila - (nFila - 13), 19, nFila, , True
        End If
        
        xlHoja1.Cells.Select
        xlHoja1.Cells.Font.Name = "Arial"
        xlHoja1.Cells.Font.Size = 9
        xlHoja1.Cells.EntireColumn.AutoFit
    End If
    PB1.value = 12
    nFilasParam = nFila 'NAGL 20171209
    nFila = nFila + 1
    
    xlHoja1.Cells(nFila, 1) = "Total Vinculados por el Articulo 202º LG"
    xlHoja1.Range(xlHoja1.Cells(nFila, 1), xlHoja1.Cells(nFila, 12)).MergeCells = True
    xlHoja1.Range("M" & nFila & ":M" & nFila).Formula = "=SUM(M" & nInicio & ":M" & nFila - 1 & ")"
    'xlHoja1.Range("Q" & nFila & ":Q" & nFila).Formula = "=SUM(Q" & nInicio & ":Q" & nFila - 1 & ")"'Comentado JOEP
    xlHoja1.Range("S" & nFila & ":S" & nFila).Formula = "=SUM(S" & nInicio & ":S" & nFila - 1 & ")"
    xlHoja1.Range("M" & nFila & ":M" & nFila).NumberFormat = "#,000.00"
    'xlHoja1.Range("M" & nFila & ":Q" & nFila).NumberFormat = "#,000.00"'Comentado JOEP
    xlHoja1.Range("M" & nFila & ":S" & nFila).NumberFormat = "#,000.00"
    'TotalArt202 = xlHoja1.Range("Q" & nFila & ":Q" & nFila) / 1000 'total vinculados Art 202° en miles de nuevos soles ' Comentado JOEP
    If nFilasParam <> 0 Then
        'TotalArt202 = xlHoja1.Range("S" & nFila & ":S" & nFila) / 1000 'total vinculados Art 202° en miles de nuevos soles 'Comento JOEP20180409 Adecuacion Segun SBS
        TotalArt202 = xlHoja1.Range("S" & nFila & ":S" & nFila) 'total vinculados Art 202° en miles de nuevos soles 'Agrego JOEP20180409 Adecuacion Segun SBS
    End If 'NAGL 20171209
    'xlHoja1.Range(xlHoja1.Cells(nFila, 1), xlHoja1.Cells(nFila, 17)).Font.Bold = True'Comentado JOEP
    xlHoja1.Range(xlHoja1.Cells(nFila, 1), xlHoja1.Cells(nFila, 19)).Font.Bold = True
    
    'JOEP
    nFila = nFila + 3
    
    xlHoja1.Cells(nFila, 2) = "2. Vinculados del Artículo 204º de la Ley General excluidos del calculo del limite del Articulo 202° de la Ley General"
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 6)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 11)).Font.Bold = True
    'JOEP
    
    nFilaTotal1 = nFila
    'xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(9, 11)).Font.Bold = True
    
        nFila = nFila + 1
        
        i = 0
        
        xlHoja1.Cells(nFila, 1) = "Nº"
        xlHoja1.Cells(nFila, 2) = "Cod"
        xlHoja1.Cells(nFila, 3) = "Nombre/Razon/"
        xlHoja1.Cells(nFila, 4) = "CIIU"
        xlHoja1.Cells(nFila, 5) = "Domicilio"
        xlHoja1.Cells(nFila, 6) = "Tipo de"
        xlHoja1.Cells(nFila, 7) = "Tipo de Doc"
        xlHoja1.Cells(nFila, 8) = "Num."
        xlHoja1.Cells(nFila, 9) = "RUC"
        xlHoja1.Cells(nFila, 10) = "Descripción de la Vinculación"
        xlHoja1.Range(xlHoja1.Cells(nFila, 10), xlHoja1.Cells(nFila, 12)).MergeCells = True
        xlHoja1.Cells(nFila, 13) = "Financiamiento"
        xlHoja1.Cells(nFila, 14) = "Depositos"
        xlHoja1.Cells(nFila, 15) = "Avales y"
        xlHoja1.Cells(nFila, 16) = "Otras garantías"
        xlHoja1.Cells(nFila, 17) = "Total"
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 2) = "SBS"
        xlHoja1.Cells(nFila, 3) = "denominación"
        xlHoja1.Cells(nFila, 6) = "persona"
        xlHoja1.Cells(nFila, 7) = "de identidad"
        xlHoja1.Cells(nFila, 8) = "Documento de"
        xlHoja1.Cells(nFila, 10) = "Propiedad"
        xlHoja1.Cells(nFila, 11) = "Propiedad"
        xlHoja1.Cells(nFila, 12) = "Gestión"
        xlHoja1.Cells(nFila, 15) = "fianzas"
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 3) = "social"
        xlHoja1.Cells(nFila, 8) = "identidad"
        xlHoja1.Cells(nFila, 10) = "Directa"
        xlHoja1.Cells(nFila, 11) = "Indirecta"
               
        xlHoja1.Range(xlHoja1.Cells(nFila - 2, 1), xlHoja1.Cells(nFila, 17)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(nFila - 2, 1), xlHoja1.Cells(nFila, 17)).Font.Bold = True
        
        ExcelCuadro xlHoja1, 1, (nFila - 2), 17, nFila
        
        nInicio = nFila + 1
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 1) = "Total Vinculados por el Articulo 204º LG"
        xlHoja1.Range(xlHoja1.Cells(nFila, 1), xlHoja1.Cells(nFila, 12)).MergeCells = True
        xlHoja1.Range(xlHoja1.Cells(nFila, 1), xlHoja1.Cells(nFila, 17)).Font.Bold = True
        
        nFilaTotal1 = nFila
        
        ExcelCuadro xlHoja1, 1, nInicio, 17, nFila, , True
        
        PB1.value = 14
   'JOEP
    nFila = nFila + 3
    xlHoja1.Cells(nFila, 2) = "3. Vinculados del Artículo 205º de la Ley General excluidos del calculo del limite del Articulo 202° de la Ley General"
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 6)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 11)).Font.Bold = True
    
    nFilaTotal1 = nFila
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(9, 11)).Font.Bold = True
    
        nFila = nFila + 1
        
        i = 0
        
        xlHoja1.Cells(nFila, 1) = "Nº"
        xlHoja1.Cells(nFila, 2) = "Cod"
        xlHoja1.Cells(nFila, 3) = "Nombre/Razon/"
        xlHoja1.Cells(nFila, 4) = "CIIU"
        xlHoja1.Cells(nFila, 5) = "Domicilio"
        xlHoja1.Cells(nFila, 6) = "Tipo de"
        xlHoja1.Cells(nFila, 7) = "Tipo de Doc"
        xlHoja1.Cells(nFila, 8) = "Num."
        xlHoja1.Cells(nFila, 9) = "RUC"
        xlHoja1.Cells(nFila, 10) = "Descripción de la Vinculación"
        xlHoja1.Range(xlHoja1.Cells(nFila, 10), xlHoja1.Cells(nFila, 12)).MergeCells = True
        xlHoja1.Cells(nFila, 13) = "Financiamiento"
        xlHoja1.Cells(nFila, 14) = "Depositos"
        xlHoja1.Cells(nFila, 15) = "Avales y"
        xlHoja1.Cells(nFila, 16) = "Otras garantías"
        xlHoja1.Cells(nFila, 17) = "Total"
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 2) = "SBS"
        xlHoja1.Cells(nFila, 3) = "denominación"
        xlHoja1.Cells(nFila, 6) = "persona"
        xlHoja1.Cells(nFila, 7) = "de identidad"
        xlHoja1.Cells(nFila, 8) = "Documento de"
        xlHoja1.Cells(nFila, 10) = "Propiedad"
        xlHoja1.Cells(nFila, 11) = "Propiedad"
        xlHoja1.Cells(nFila, 12) = "Gestión"
        xlHoja1.Cells(nFila, 15) = "fianzas"
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 3) = "social"
        xlHoja1.Cells(nFila, 8) = "identidad"
        xlHoja1.Cells(nFila, 10) = "Directa"
        xlHoja1.Cells(nFila, 11) = "Indirecta"
        
        xlHoja1.Range(xlHoja1.Cells(nFila - 2, 1), xlHoja1.Cells(nFila, 17)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(nFila - 2, 1), xlHoja1.Cells(nFila, 17)).Font.Bold = True
        
        ExcelCuadro xlHoja1, 1, nFila - 2, 17, nFila
        
        nInicio = nFila + 1
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 1) = "Total Vinculados por el Articulo 205º LG"
        xlHoja1.Range(xlHoja1.Cells(nFila, 1), xlHoja1.Cells(nFila, 12)).MergeCells = True
        xlHoja1.Range(xlHoja1.Cells(nFila, 1), xlHoja1.Cells(nFila, 17)).Font.Bold = True
        
        nFilaTotal1 = nFila
        
        ExcelCuadro xlHoja1, 1, nInicio, 17, nFila, , True
    'JOEP
         
         
    nFila = nFila + 2
    xlHoja1.Cells(nFila, 2) = "4. Exposición a Vinculados"
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 6)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 18)).Font.Bold = True
    
    nFila = nFila + 2
    xlHoja1.Cells(nFila, 2) = "Artículo 202 de la Ley General"
    'xlHoja1.Cells(nFila, 5) = "Artículo 204 de la Ley General" 'Comento JOEP20180409 Adecuacion Segun SBS
    'xlHoja1.Cells(nFila, 8) = "Total Exposición a Vinculados" 'Comento JOEP20180409 Adecuacion Segun SBS
    'xlHoja1.Range(xlHoja1.Cells(nFila, 9), xlHoja1.Cells(nFila, 10)).MergeCells = True 'Comento JOEP20180409 Adecuacion Segun SBS
    
    'JOEP
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 2)).Font.Bold = True
    'xlHoja1.Range(xlHoja1.Cells(nFila, 5), xlHoja1.Cells(nFila, 5)).Font.Bold = True'Comento JOEP20180409 Adecuacion Segun SBS
    'xlHoja1.Range(xlHoja1.Cells(nFila, 8), xlHoja1.Cells(nFila, 8)).Font.Bold = True'Comento JOEP20180409 Adecuacion Segun SBS
    'JOEP
    
    PB1.value = 18
    
    nInicio = nFila + 1
    
    nFila = nFila + 1
    xlHoja1.Cells(nFila, 2) = "Total Financiero a vinculados 202 LG(A)"
    If nFilaTotal1 <> 0 Then
    'xlHoja1.Range(xlHoja1.Cells(nFila, 4), xlHoja1.Cells(nFila, 4)).Formula = "=+" & xlHoja1.Cells(nFilaTotal1, 14)
    End If
    'xlHoja1.Cells(nFila, 5) = "Total Financiamiento a vinculados 204 LG(B)" 'Comento JOEP20180409 Adecuacion Segun SBS
    'xlHoja1.Range(xlHoja1.Cells(nFila, 7), xlHoja1.Cells(nFila, 7)).Formula = "=+" & xlHoja1.Cells(nFilaTotal2, 15)
    'xlHoja1.Cells(nFila, 8) = "Total Financiamiento a vinculados (A+B)" 'Comento JOEP20180409 Adecuacion Segun SBS
    If nFilaTotal1 <> 0 Then
    'xlHoja1.Range(xlHoja1.Cells(nFila, 10), xlHoja1.Cells(nFila, 10)).Formula = "=+" & xlHoja1.Cells(nFilaTotal1, 14) & "+" & xlHoja1.Cells(nFilaTotal2, 15)
    End If
    
    xlHoja1.Cells(nFila, 2) = "Total financ.a vinculados 202° LG (A)"
    'xlHoja1.Cells(nFila, 3) = Format(TotalArt202, "0")'Comento JOEP20180409 Adecuacion Segun SBS
    xlHoja1.Cells(nFila, 3) = TotalArt202 'Agrego JOEP20180409 Adecuacion Segun SBS
    'xlHoja1.Cells(nFila, 5) = "Total financ.a vinculados 204° LG (A)" 'Comento JOEP20180409 Adecuacion Segun SBS
    'xlHoja1.Cells(nFila, 8) = "Total financ.a vinculados (A+B)" 'Comento JOEP20180409 Adecuacion Segun SBS
    'xlHoja1.Cells(nFila, 9) = Format(TotalArt202, "0") 'Comento JOEP20180409 Adecuacion Segun SBS
    
    nFila = nFila + 1
    xlHoja1.Cells(nFila, 2) = "Patrimonio Efectivo(D)"
    xlHoja1.Cells(nFila, 3) = Format(nPatrEfectivo / 1000, "0")
    'xlHoja1.Cells(nFila, 5) = "Patrimonio Efectivo(C)" 'Comento JOEP20180409 Adecuacion Segun SBS
    'xlHoja1.Cells(nFila, 8) = "Patrimonio Efectivo(C)" 'Comento JOEP20180409 Adecuacion Segun SBS
    'xlHoja1.Cells(nFila, 9) = Format(nPatrEfectivo / 1000, "0") 'Comento JOEP20180409 Adecuacion Segun SBS
    
    nFila = nFila + 1
    
    xlHoja1.Cells(nFila, 2) = "Exposición (A)/(D)*100%"
    'xlHoja1.Cells(nFila, 3) = TotalArt202 / (nPatrEfectivo) / 1000) 'peac 20130405
    xlHoja1.Cells(nFila, 3) = TotalArt202 / (IIf(nPatrEfectivo = 0, 1, nPatrEfectivo / 1000)) 'peac 20130405
    'xlHoja1.Cells(nFila, 5) = "Exposición (B)/(C)*100%" 'Comento JOEP20180409 Adecuacion Segun SBS
    'xlHoja1.Cells(nFila, 8) = "Exposición (A+B)/(C)*100%"''Comento JOEP20180409 Adecuacion Segun SBS
    
    'xlHoja1.Cells(nFila, 9) = TotalArt202 / (nPatrEfectivo / 1000)
     
    'xlHoja1.Cells(nFila, 9) = TotalArt202 / (IIf(nPatrEfectivo = 0, 1, nPatrEfectivo / 1000)) 'peac 20130405 'Comento JOEP20180409 Adecuacion Segun SBS
    'xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 10)) = "=DECIMAL(((C" & 47 & "/" & "C" & 48 & ")" & "*" & 100 & "),2"
    xlHoja1.Range("C" & nFila & ":C" & nFila).Formula = "=round(((C" & nInicio & "/" & "C" & nInicio + 1 & ")" & "*" & 100 & "%" & "),2)" 'Agrego JOEP20180409 Adecuacion Segun SBS
    'Round((Format(TotalArt202, "0") / Format(nPatrEfectivo / 1000, "0")) * 100, 2)
     
    
    xlHoja1.Cells.Select
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 9
    xlHoja1.Cells.EntireColumn.AutoFit

    
    ExcelCuadro xlHoja1, 2, nInicio, 3, nFila, , True
    'ExcelCuadro xlHoja1, 5, nInicio, 6, nFila, , True 'Comento JOEP20180409 Adecuacion Segun SBS
    'ExcelCuadro xlHoja1, 8, nInicio, 9, nFila, , True 'Comento JOEP20180409 Adecuacion Segun SBS

'************ Nueva Hoja *******************
    PB1.value = 23
'JOEP
    'Nombre de la Hoja
    ExcelAddHoja "Rep. N° 21-A", xlLibro, xlHoja1
    
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    
    'Cabecera
    xlHoja1.Cells(2, 10) = "REPORTE Nº 21 - A"
    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).HorizontalAlignment = xlLeft

    xlHoja1.Cells(4, 2) = "INFORMACION DE LAS PERSONAS JURIDICAS Y ENTES JURIDICOS VINCULADOS A LA EMPRESA"
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).HorizontalAlignment = xlCenter

    xlHoja1.Cells(6, 2) = "Empresa que remite la información: " & " " & gsNomCmac
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 8)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 8)).HorizontalAlignment = xlLeft

    xlHoja1.Cells(7, 2) = "información al:" & " " & Format(pdFecha, "DD") & " de " & Format(pdFecha, "MMMM") & " de " & Format(pdFecha, "YYYY")
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).HorizontalAlignment = xlLeft

    xlHoja1.Cells(9, 2) = "Razón o denominación social de la persona jurídicas o ente jurídico: "
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(9, 11)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(9, 11)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(9, 11)).Font.Bold = True
    
    'Tabla
    nFila = 11
            nInicio = nFila + 1
            xlHoja1.Cells(nFila, 1) = "Nº"
            xlHoja1.Cells(nFila, 2) = "Nombre"
            xlHoja1.Cells(nFila, 3) = "Cod"
            xlHoja1.Cells(nFila, 4) = "Tipo de"
            xlHoja1.Cells(nFila, 5) = "Numero del"
            xlHoja1.Cells(nFila, 6) = "RUC"
            xlHoja1.Cells(nFila, 7) = "Residencia"
            xlHoja1.Cells(nFila, 8) = "Accionista"
            xlHoja1.Cells(nFila, 9) = "Cargo"
            xlHoja1.Cells(nFila, 10) = "Otro"
                        
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 3) = "SBS"
            xlHoja1.Cells(nFila, 4) = "documento"
            xlHoja1.Cells(nFila, 5) = "documento"
            xlHoja1.Cells(nFila, 8) = "o"
            xlHoja1.Cells(nFila, 10) = "Cargo"
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 8) = "equivalente"
            
            xlHoja1.Range(xlHoja1.Cells(nFila - 2, 1), xlHoja1.Cells(nFila, 10)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(nFila - 2, 1), xlHoja1.Cells(nFila, 10)).Font.Bold = True
                        
            ExcelCuadro xlHoja1, 1, nFila - 2, 10, nFila
            
    'Contenido de la Tabla
    Dim l As Integer
     l = 3
     nCon = 1
     
            Do While Not l
                nFila = nFila + 1
                    xlHoja1.Cells(nFila, 1) = nCon
                    nCon = nCon + 1
                l = l - 1
            Loop
                        
            ExcelCuadro xlHoja1, 1, nFila - 3, 10, nFila, , True
            
            xlHoja1.Cells.Select
            xlHoja1.Cells.Font.Name = "Arial"
            xlHoja1.Cells.Font.Size = 9
            xlHoja1.Cells.EntireColumn.AutoFit
    
'JOEP
    PB1.value = 25
    PB1.Visible = False
    MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "Aviso"
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
    CargaArchivo "Reporte21_" & CStr(Format(pdFecha, "yyyymmdd")) & ".xls", App.path & "\Spooler"
End If
End Sub

Private Sub cmdAgregar_Click()
Dim nItem As Integer
FERelVin.lbEditarFlex = True
If FERelVin.Row = 1 Then
        FERelVin.AdicionaFila
ElseIf FERelVin.TextMatrix(FERelVin.Row, 1) = "" Then
        nItem = FERelVin.Row
        FERelVin.EliminaFila nItem
ElseIf ValidaListaVinc("optAgr") Then
    FERelVin.AdicionaFila
End If
End Sub 'NAGL ERS074-2017 20171209

Private Sub cmdEliminar_Click()
Dim rs As ADODB.Recordset
Dim psFechaCons As Date
Dim psPersCod As String
Dim Dgrup As New DGrupoEco

If FERelVin.TextMatrix(FERelVin.Row, 1) <> "" Then
    psFechaCons = CalculaFechaFinMes
    If MsgBox("Esta seguro de eliminar al Vinculado " & Trim(FERelVin.TextMatrix(FERelVin.Row, 2)) & " ..!!", vbYesNo + vbInformation, "Atención") = vbNo Then Exit Sub
          psPersCod = Trim(FERelVin.TextMatrix(FERelVin.Row, 1))
          Call Dgrup.EliminarVinculadoRpte21(psPersCod, psFechaCons)
          Call FERelVin.EliminaFila(FERelVin.Row)
Else
    Call FERelVin.EliminaFila(FERelVin.Row)
End If
'lsPersCod = FERelVin.TextMatrix(FERelVin.Row, 1)
'oCon.AbreConexion
'sSql = "Delete from VinculadosReporte21 " _
'     & "Where cPerscod = '" & lsPersCod & "'"
'oCon.Ejecutar (sSql)
'oCon.CierraConexion
'Call MuestraVinculados 'Comentado by NAGL
End Sub

Private Sub CargarRelInst()
    Dim oDGeneral As New DGeneral
    Dim rsRel As New ADODB.Recordset
    Set rsRel = oDGeneral.GetTipoRelInst()
    FERelVin.CargaCombo rsRel
    Set rsRel = Nothing
    Set oDGeneral = Nothing
End Sub 'NAGL ERS075-2017 20171209

Private Sub FERelVin_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim oPersona As New UPersona
Set oPersona = frmBuscaPersona.inicio(False) 'NAGL 20190610 Cambió de True a False Según Correo
If Not oPersona Is Nothing Then
    FERelVin.TextMatrix(pnRow, 1) = oPersona.sPersCod
    FERelVin.TextMatrix(pnRow, 2) = oPersona.sPersNombre
    If ValidaListaVinc("Busc") = False Then Exit Sub
End If
End Sub 'NAGL ERS075-2017 20171209

Private Sub FERelVin_OnChangeCombo()
    Dim cValVinc As String
    Dim cValidar As String
    cValidar = "0123456789"
    cValVinc = Trim(Right(FERelVin.TextMatrix(FERelVin.Row, 4), 2))
    FERelVin.TextMatrix(FERelVin.Row, 4) = IIf(InStr(cValidar, Trim(Right(FERelVin.TextMatrix(FERelVin.Row, 3), 2))) <> 0, Trim(Right(FERelVin.TextMatrix(FERelVin.Row, 3), 2)), cValVinc)
End Sub 'NAGL ERS075-2017 20171209

Public Function CalculaPatrimonioEfectivo(psFechaCons As Date) As Currency
Dim rs As New ADODB.Recordset
Dim pnTipo As Integer
Dim pnPatrEfec As Currency
pnTipo = 4 '3 Cambiado by NAGL 20190118
    Set rs = oDbalanceCont.recuperarPatrimonioEfectivoEleccion(pnTipo, CInt(Month(psFechaCons)), CInt(Year(psFechaCons)), psFechaCons, gdFecSis)
    If Not (rs.BOF And rs.EOF) Then
        pnPatrEfec = rs!nSaldo
    Else
        pnPatrEfec = 0
    End If
    CalculaPatrimonioEfectivo = pnPatrEfec
End Function 'NAGL ERS074-2017 20171209

Private Function CalculaFechaFinMes() As Date
Dim sFecha  As Date
Dim sFechaFinMesAnterior As Date
sFecha = "01/" & IIf(Len(Trim(cboMes.ListIndex + 1)) = 1, "0" & Trim(Str(cboMes.ListIndex + 1)), Trim(cboMes.ListIndex + 1)) & "/" & Trim(txtAnio.Text)
sFechaFinMesAnterior = DateAdd("d", -1, DateAdd("m", 1, sFecha))
CalculaFechaFinMes = sFechaFinMesAnterior
End Function 'NAGL ERS074-2017 20171209

Public Sub CalculaPatrimonioDolares(Optional psFiltro As String)
Dim psFechaCons As Date
Dim lnTipoCambioFC As Double
If txtAnio.Text <> "" Then
    If psFiltro = "Sist" Then
        psFechaCons = DateAdd("d", -Day(gdFecSis), gdFecSis)
    Else
        psFechaCons = CalculaFechaFinMes
    End If
    lnTipoCambioFC = oDbalanceCont.ObtenerTipoCambioCierreNew(psFechaCons)
    If txtPatrEfec <> "" And txtPatrEfec <> "0.00" And txtPatrEfec <> "." And lnTipoCambioFC <> 0 Then
       txtPatrDol.Text = Format(CCur(txtPatrEfec.Text) / lnTipoCambioFC, "#,##0.00")
    Else
       txtPatrDol.Text = "0.00"
    End If
End If
End Sub 'NAGL ERS074-2017 20171209

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
Dim psFechaCons As Date
KeyAscii = NumerosEnteros(KeyAscii)

If KeyAscii = 13 Then
    'If Len(Trim(cboMes.Text)) > 0 And Val(txtAnio.Text) > 1900 Then
    'txtTC.Text = Format(oCambio.EmiteTipoCambio(sFecha, TCFijoDia), "#,##0.00000") '*** PEAC 20130220 - se aumento un decimal
     If ValidaDatosRep21("Opt") Then
        psFechaCons = CalculaFechaFinMes
        txtTC.Text = oDbalanceCont.ObtenerTipoCambioCierreNew(psFechaCons) 'Se quitó "TipoAct", para que tome T.C cierre NAGL 20180125
        txtTC.Text = IIf(txtTC = 0, 0, Format(txtTC, "#,##0.000"))
        txtPatrEfec.Text = Format(CalculaPatrimonioEfectivo(psFechaCons), "#,##0.00")
        CalculaPatrimonioDolares
        txtPatrEfec.SetFocus
    End If '**NAGL ERS074-2017 20171209
End If
End Sub

Private Sub cmdActualizarDataRRHH_Click()
    Dim pdFechaCons As Date
    pdFechaCons = CalculaFechaFinMes
    Call MuestraVinculados(pdFechaCons, "RH")
End Sub 'NAGL 20190709 RFC1907050001

'************************NAGL ERS074-2017 20171209******************************'
Private Sub CboMes_Click()
Dim psFechaCons As Date
Dim rs As New ADODB.Recordset
Dim DGrp As New DGrupoEco 'NAGL 20190705

psFechaCons = CalculaFechaFinMes 'Subido by NAGL
If DGrp.ObtieneDataRRHH(gdFecSis, psFechaCons) = True Then
    cmdActualizarDataRRHH.Enabled = True
Else
    cmdActualizarDataRRHH.Enabled = False
End If 'Agregado by NAGL 20190705

If ValidaDatosRep21("Opt") Then
    'psFechaCons = CalculaFechaFinMes
    txtTC.Text = oDbalanceCont.ObtenerTipoCambioCierreNew(psFechaCons) 'Se quitó "TipoAct", para que tome T.C cierre NAGL 20180125
    txtTC.Text = IIf(txtTC = 0, 0, Format(txtTC, "#,##0.000"))
    txtPatrEfec.Text = Format(CalculaPatrimonioEfectivo(psFechaCons), "#,##0.00")
    CalculaPatrimonioDolares
    Call MuestraVinculados(psFechaCons)
End If
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         txtAnio.SetFocus
    End If
End Sub

Private Sub txtAnio_GotFocus()
    fEnfoque txtAnio
End Sub

Private Sub txtAnio_LostFocus()
    Dim psFechaCons As Date
    Dim rs As New ADODB.Recordset
    Dim pnTipo As Integer
    Dim DGrp As New DGrupoEco 'NAGL 20190705
    pnTipo = 3
    psFechaCons = CalculaFechaFinMes 'Subido by NAGL
    If DGrp.ObtieneDataRRHH(gdFecSis, psFechaCons) = True Then
        cmdActualizarDataRRHH.Enabled = True
    Else
        cmdActualizarDataRRHH.Enabled = False
    End If 'Agregado by NAGL 20190705
    
    If ValidaDatosRep21("Opt") Then
        'psFechaCons = CalculaFechaFinMes
        txtTC.Text = oDbalanceCont.ObtenerTipoCambioCierreNew(psFechaCons) 'Se quitó "TipoAct", para que tome T.C cierre NAGL 20180125
        txtTC.Text = IIf(txtTC = 0, 0, Format(txtTC, "#,##0.000"))
        txtPatrEfec.Text = Format(CalculaPatrimonioEfectivo(psFechaCons), "#,##0.00")
        CalculaPatrimonioDolares
        Call MuestraVinculados(psFechaCons)
    End If
End Sub

'Private Sub txtTC_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtTC, KeyAscii)
'End Sub
Private Sub txtPatrEfec_GotFocus()
      fEnfoque txtPatrEfec
End Sub

Private Sub txtPatrEfec_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPatrEfec, KeyAscii)
If KeyAscii = 13 Then
    txtPatrEfec = Format(txtPatrEfec, "#,##0.00")
    CalculaPatrimonioDolares
    txtPatrDol.SetFocus
End If
End Sub
Private Sub txtPatrEfec_LostFocus()
    CalculaPatrimonioDolares
End Sub

Private Sub txtPatrDol_GotFocus()
      fEnfoque txtPatrDol
End Sub

Private Sub txtPatrDol_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPatrEfec, KeyAscii)
If KeyAscii = 13 Then
      cmdGenerar.SetFocus
End If
End Sub
'*************************END NAGL ERS072-2017*****************************************'

Private Sub cmdSalir_Click()
Unload Me
End Sub










'Private Sub cmdCancelar_Click()
'    Call MuestraVinculados
'    'Habilitar Controles
'    FERelVinNoMoverdeFila = -1
'    FERelVin.lbEditarFlex = False
'    cmdNuevo.Enabled = True
'    cmdEditar.Enabled = True
'    cmdEliminar.Enabled = True
'
'    cmdGuardar.Enabled = False
'    cmdCancelar.Enabled = False
'
'    FERelVin.SetFocus
'
'End Sub

'Private Sub CmdEditar_Click()
'
'Call MuestraVinculados
'
'If NumeroVincu > 0 Then
'        FERelVinNoMoverdeFila = FERelVin.Row
'        FERelVin.lbEditarFlex = True
'        FERelVin.SetFocus
'        cmdNuevo.Enabled = False
'        cmdEditar.Enabled = False
'        cmdEliminar.Enabled = False
'        lnTipo = 1
'        lsPersCodAnt = FERelVin.TextMatrix(FERelVin.Row, 1)
'        cmdGuardar.Enabled = True
'        cmdCancelar.Enabled = True
'Else
'    MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
'End If
'End Sub

'Private Sub cmdGuardar_Click()
'    oCon.AbreConexion
'
'    If lnTipo = 0 Then
'        lsPersCod = FERelVin.TextMatrix(FERelVin.Row, 1)
'        lsVinculo = UCase(FERelVin.TextMatrix(FERelVin.Row, 3))
'        sSql = "Insert VinculadosReporte21 " _
'             & "Values('" & lsPersCod & "', '" & lsVinculo & "')"
'    Else
'        lsPersCod = FERelVin.TextMatrix(FERelVin.Row, 1)
'        lsVinculo = UCase(FERelVin.TextMatrix(FERelVin.Row, 3))
'        sSql = "Update VinculadosReporte21 Set " _
'             & "cPerscod = '" & lsPersCod & "'," _
'             & "cVinculo = '" & lsVinculo & "'" _
'             & "Where cPerscod = '" & lsPersCodAnt & "'"
'    End If
'    oCon.Ejecutar (sSql)
'    oCon.CierraConexion
'
'    FERelVin.lbEditarFlex = False
'    cmdNuevo.Enabled = True
'    cmdEditar.Enabled = True
'    cmdEliminar.Enabled = True
'
'    cmdGuardar.Enabled = False
'    cmdCancelar.Enabled = False
'
'    FERelVin.SetFocus
'End Sub

'Private Sub cmdNuevo_Click()
'    FERelVin.lbEditarFlex = True
'    FERelVin.AdicionaFila
'    FERelVinNoMoverdeFila = FERelVin.Rows - 1
'    FERelVin.SetFocus
'
'    cmdNuevo.Enabled = False
'    cmdEditar.Enabled = False
'    cmdEliminar.Enabled = False
'    cmdGuardar.Enabled = True
'    cmdCancelar.Enabled = True
'    lnTipo = 0
'End Sub
'
''Private Function RecuperaListaVinculados() As ADODB.Recordset
''Set oCon = New DConecta
''oCon.AbreConexion
''sSql = "Select V.cPerscod, P.cPersnombre, V.cVinculo " _
''     & "From VinculadosReporte21 V " _
''     & "Inner Join Persona P On V.cPerscod = P.cPerscod"
''
''Set RecuperaListaVinculados = oCon.CargaRecordSet(sSql)
''NumeroVincu = RecuperaListaVinculados.RecordCount
''oCon.CierraConexion
''
''End Function
'Comentado by NAGL 20171209

'Private Function RecuperaPatrimonioEfectivo(sBalMes As String, sBalAnio As String) As Currency
'    Set oCon = New DConecta
'    Dim rs As ADODB.Recordset
'    Dim Total As Currency
'    oCon.AbreConexion
'    sSql = "exec stp_sel_Reporte3Patrimonio '" & sBalAnio & "', '" & sBalMes & "', '0'"
'    Set rs = oCon.CargaRecordSet(sSql)
'
'
'    Total = 0
'    Do While Not rs.EOF
'        DoEvents
'
'        If rs!cCtaContCod = "310101" Or rs!cCtaContCod = "310102" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "3201" Or rs!cCtaContCod = "320201" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "320202" Or rs!cCtaContCod = "320301" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "320302" Or rs!cCtaContCod = "3301" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "3302" Or rs!cCtaContCod = "330301" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "380101" Or rs!cCtaContCod = "390101" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "380104" Or rs!cCtaContCod = "280202" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280203" Or rs!cCtaContCod = "280204" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280601" Or rs!cCtaContCod = "280602" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280603" Or rs!cCtaContCod = "380201" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "380203" Or rs!cCtaContCod = "380204" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "380205" Or rs!cCtaContCod = "3902" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "190407" Or rs!cCtaContCod = "19040907" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "17" Or rs!cCtaContCod = "130308" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13080308" Or rs!cCtaContCod = "13090308" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13040801" Or rs!cCtaContCod = "1308040801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1309040801" Or rs!cCtaContCod = "330302" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "310103" Or rs!cCtaContCod = "310104" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "310109" Or rs!cCtaContCod = "320203" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "320204" Or rs!cCtaContCod = "320209" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "320303" Or rs!cCtaContCod = "320304" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "320309" Or rs!cCtaContCod = "260202010201" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "2602020202" Or rs!cCtaContCod = "2602020302" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "2603020102" Or rs!cCtaContCod = "2603020202" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "2603020302" Or rs!cCtaContCod = "2604020102" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "2604020202" Or rs!cCtaContCod = "2604020302" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "2605010102" Or rs!cCtaContCod = "2605010202" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "2605010302" Or rs!cCtaContCod = "2606020102" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "2606020202" Or rs!cCtaContCod = "2606020302" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "2607010102" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "2607010202" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "2607010302" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280202" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280203" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280204" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280601" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280602" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280603" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "310104" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "310109" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "320204" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "320209" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "320304" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "320309" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "2607010202" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "2607010302" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280202" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280203" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280204" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280601" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280602" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280603" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090202" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090203" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090206" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090302" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090303" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090306" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090402" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090403" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090502" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090503" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090602" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090603" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090702" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090703" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090802" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090803" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090902" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14090903" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14091002" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14091003" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14091102" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14091103" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14091202" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14091203" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14091302" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14091303" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14091306" Then
'            Total = Total + rs!nSaldoFinImporte * (1.25) / 100
'        ElseIf rs!cCtaContCod = "1509040101" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "270102" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "270203" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "270204" Then
'            Total = Total + rs!nSaldoFinImporte * (1.25) / 100
'        ElseIf rs!cCtaContCod = "130308" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13080308" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13090308" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13040801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1308040801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1309040801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "310104" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "310109" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "320204" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "320209" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "320304" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "320309" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280203" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280204" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280601" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280602" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "280603" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "130308" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13080308" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13090308" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13040801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1308040801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1309040801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "130105" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "130106" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13020510" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13020610" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "130305" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "130306" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "130309" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1303180105" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1303180205" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13040510" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13040610" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13040910" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1304180105" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1304180205" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13050510" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "13050610" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1305180105" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1305180106" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1305180109" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1305180205" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1305180206" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1305180209" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1701" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1702" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "170701" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "170702" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1401071801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1401081801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1401090606" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1401091801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14011006" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1401101801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1401111801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1401121801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1401131801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1404071801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1404081801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1404090606" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1404091801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14041006" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1404101801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1404111801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1404121801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1404131801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405071801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405071918" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405072218" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405081801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405081918" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405090606" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405091801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405091908" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14051006" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405101801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405101906" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405101918" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405111801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405111918" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405112218" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405121801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405121918" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405131801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1405131918" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406071801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406071918" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406072218" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406081801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406081918" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406090606" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406091801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406091918" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "14061006" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406101801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406101906" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406101918" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406111801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406111918" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406112218" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406121801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406121918" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406131801" Then
'            Total = Total + rs!nSaldoFinImporte
'        ElseIf rs!cCtaContCod = "1406131918" Then
'            Total = Total + rs!nSaldoFinImporte
'        End If
'
'        rs.MoveNext
'        If rs.EOF Then
'           Exit Do
'        End If
'    Loop
'    RecuperaPatrimonioEfectivo = Total
'    oCon.CierraConexion
'End Function



