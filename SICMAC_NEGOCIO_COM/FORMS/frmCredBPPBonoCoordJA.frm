VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPBonoCoordJA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Generar Bono Coordinador de Créditos y Jefe de Agencia"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17295
   Icon            =   "frmCredBPPBonoCoordJA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   17295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTBusqueda 
      Height          =   7770
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   13705
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Generar Bono Mensual"
      TabPicture(0)   =   "frmCredBPPBonoCoordJA.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdExportar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraBusqueda"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCerrar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15600
         TabIndex        =   11
         Top             =   7200
         Width           =   1170
      End
      Begin VB.Frame fraBusqueda 
         Caption         =   "Filtro"
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
         Height          =   6735
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   16815
         Begin VB.CommandButton cmdGenerar 
            Caption         =   "Generar"
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
            Left            =   4080
            TabIndex        =   4
            Top             =   960
            Width           =   1170
         End
         Begin VB.ComboBox cmbAgencias 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1035
            Width           =   3735
         End
         Begin SICMACT.FlexEdit feAnalistas 
            Height          =   4815
            Left            =   120
            TabIndex        =   5
            Top             =   1800
            Width           =   16575
            _extentx        =   29236
            _extenty        =   8493
            cols0           =   12
            highlight       =   1
            encabezadosnombres=   "#-Agencia-Usuario-Comite-% Bonif.-Meta-Cierre-Meta-Cierre-Caja-Agencia-Total"
            encabezadosanchos=   "0-2000-1000-1500-1200-1500-1500-1500-1500-1500-1500-1200"
            font            =   "frmCredBPPBonoCoordJA.frx":0326
            font            =   "frmCredBPPBonoCoordJA.frx":034E
            font            =   "frmCredBPPBonoCoordJA.frx":0376
            font            =   "frmCredBPPBonoCoordJA.frx":039E
            font            =   "frmCredBPPBonoCoordJA.frx":03C6
            fontfixed       =   "frmCredBPPBonoCoordJA.frx":03EE
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-C-L-R-R-R-R-R-R-R-R"
            formatosedit    =   "0-0-0-0-2-2-2-2-2-2-2-2"
            cantentero      =   15
            textarray0      =   "#"
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rendimiento de Cartera"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   11840
            TabIndex        =   12
            Top             =   1485
            Width           =   3010
         End
         Begin VB.Label lblFiltroSelect 
            AutoSize        =   -1  'True
            Caption         =   "Filtro de Generación:"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   795
            Width           =   1470
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo de Cartera"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5840
            TabIndex        =   9
            Top             =   1485
            Width           =   3010
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo Cartera Vencida y Judicial"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   8840
            TabIndex        =   8
            Top             =   1485
            Width           =   3010
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mes a generar:"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   1065
         End
         Begin VB.Label lblMes 
            AutoSize        =   -1  'True
            Caption         =   "@Mes"
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
            Left            =   1440
            TabIndex        =   6
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   7200
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmCredBPPBonoCoordJA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private fMatCoordJA() As CoordinadorJABPP
'Private i As Integer
'Private nIndex As Integer
'Private fgFecActual As Date
'
'Private Sub CargaControles()
'cmdExportar.Enabled = False
'MesActual
'CargaComboAgenciasLocal cmbAgencias
'lblMes.Caption = MesAnio(fgFecActual)
'End Sub
'
'Public Sub CargaComboAgenciasLocal(ByRef combo As ComboBox)
'Dim oConst As COMDConstantes.DCOMAgencias
'Dim R As ADODB.Recordset
'    On Error GoTo ERRORCargaComboAgencias
'    combo.Clear
'    Set oConst = New COMDConstantes.DCOMAgencias
'    Set R = oConst.ObtieneAgencias()
'    Set oConst = Nothing
'    combo.AddItem "Todas" & Space(250) & "%"
'    Do While Not R.EOF
'        combo.AddItem R!cConsDescripcion & Space(250) & R!nConsValor
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'    Exit Sub
'
'ERRORCargaComboAgencias:
'    MsgBox err.Description, vbCritical, "Aviso"
'End Sub
'
'Private Function ValidaDatos() As Boolean
'ValidaDatos = True
'If Trim(cmbAgencias.Text) = "" Then
'    MsgBox "Seleccione la Agencia", vbInformation, "Aviso"
'    ValidaDatos = False
'    Exit Function
'End If
'
'If Trim(Right(cmbAgencias.Text, 2)) = "%" Then
'    If MsgBox("Estas seguro de procesar el BPP para Todas la Agencias?", vbInformation + vbYesNo, "Aviso") = vbNo Then
'        ValidaDatos = False
'        Exit Function
'    End If
'End If
'
'End Function
'
'Private Sub MesActual()
'Dim oConsSist As COMDConstSistema.NCOMConstSistema
'Set oConsSist = New COMDConstSistema.NCOMConstSistema
'fgFecActual = oConsSist.LeeConstSistema(gConstSistFechaBPP)
'Set oConsSist = Nothing
'End Sub
'
'Private Function MesAnio(ByVal dFecha As Date) As String
'Dim sFechaDesc As String
'sFechaDesc = ""
'
'Select Case Month(dFecha)
'    Case 1: sFechaDesc = "Enero"
'    Case 2: sFechaDesc = "Febrero"
'    Case 3: sFechaDesc = "Marzo"
'    Case 4: sFechaDesc = "Abril"
'    Case 5: sFechaDesc = "Mayo"
'    Case 6: sFechaDesc = "Junio"
'    Case 7: sFechaDesc = "Julio"
'    Case 8: sFechaDesc = "Agosto"
'    Case 9: sFechaDesc = "Septiembre"
'    Case 10: sFechaDesc = "Octubre"
'    Case 11: sFechaDesc = "Noviembre"
'    Case 12: sFechaDesc = "Diciembre"
'End Select
'
'sFechaDesc = sFechaDesc & " " & CStr(Year(dFecha))
'MesAnio = UCase(sFechaDesc)
'End Function
'
'Private Sub cmdExportar_Click()
'GenerarExcel
'End Sub
'
'Private Sub cmdGenerar_Click()
'If ValidaDatos Then
'    If MsgBox("Estás Seguro de Generar el BPP?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'    Dim sCodAge As String
'    sCodAge = Trim(Right(cmbAgencias.Text, 2))
'
'    LimpiaFlex feAnalistas
'    CargaDatos sCodAge
'
'    If fMatCoordJA(0).Usuario <> "" Then
'        For i = 0 To UBound(fMatCoordJA)
'            feAnalistas.AdicionaFila
'            feAnalistas.TextMatrix(i + 1, 1) = fMatCoordJA(i).Agencia
'            feAnalistas.TextMatrix(i + 1, 2) = fMatCoordJA(i).Usuario
'            feAnalistas.TextMatrix(i + 1, 3) = fMatCoordJA(i).comite
'            feAnalistas.TextMatrix(i + 1, 4) = Format(fMatCoordJA(i).PorBonificacion * 100, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 5) = Format(fMatCoordJA(i).SaldoCartera, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 6) = Format(fMatCoordJA(i).SaldoCarteraCierre, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 7) = Format(fMatCoordJA(i).PorSalCartVencJud * 100, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 8) = Format(fMatCoordJA(i).PorcSalVenJud * 100, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 9) = Format(fMatCoordJA(i).RendCaja * 100, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 10) = Format(fMatCoordJA(i).RendAG * 100, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 11) = Format(fMatCoordJA(i).BonoTotal, "###," & String(15, "#") & "#0.00")
'        Next i
'        feAnalistas.TopRow = 1
'        cmdExportar.Enabled = True
'    Else
'        MsgBox "No hay Datos", vbInformation, "Aviso"
'    End If
'
'End If
'End Sub
'
'Private Sub Form_Load()
'CargaControles
'End Sub
'Private Sub CargaDatos(ByVal psCodAge As String)
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
' On Error GoTo ErrorCargaDatos
'Set oBPP = New COMNCredito.NCOMBPPR
'Set rsBPP = oBPP.GenerarBPPCoordJA(psCodAge)
'
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    ReDim fMatCoordJA(rsBPP.RecordCount - 1)
'    For i = 0 To rsBPP.RecordCount - 1
'        fMatCoordJA(i).cCodAge = rsBPP!cCodAge
'        fMatCoordJA(i).Agencia = rsBPP!Agencia
'        fMatCoordJA(i).nCargo = rsBPP!nCargo
'        fMatCoordJA(i).Cargo = rsBPP!Cargo
'        fMatCoordJA(i).cPersCod = rsBPP!cPersCod
'        fMatCoordJA(i).comite = rsBPP!comite
'        fMatCoordJA(i).Usuario = rsBPP!Usuario
'        fMatCoordJA(i).Nombre = rsBPP!Nombre
'        fMatCoordJA(i).CantAnalista = rsBPP!CantAnalista
'        fMatCoordJA(i).PorBonificacion = rsBPP!PorBonificacion
'        fMatCoordJA(i).SaldoCartera = rsBPP!SaldoCartera
'        fMatCoordJA(i).SaldoCarteraCierre = rsBPP!SaldoCarteraCierre
'        fMatCoordJA(i).PorSalCartVencJud = rsBPP!PorSalCartVencJud
'        fMatCoordJA(i).PorcSalVenJud = rsBPP!PorcSalVenJud
'        fMatCoordJA(i).RendCaja = rsBPP!RendCaja
'        fMatCoordJA(i).RendAG = rsBPP!RendAG
'        fMatCoordJA(i).Tope = rsBPP!nTope
'        fMatCoordJA(i).BonoTotal = rsBPP!BonoTotal
'        fMatCoordJA(i).PorBonificacionConf = rsBPP!PorBonificacionConf
'        fMatCoordJA(i).AnalistaBoni = rsBPP!AnalistaBoni
'        fMatCoordJA(i).SaldoCapital = rsBPP!SaldoCapital
'        fMatCoordJA(i).SaldoVencJud = rsBPP!SaldoVencJud
'        fMatCoordJA(i).IntCob = rsBPP!IntCob
'        rsBPP.MoveNext
'    Next i
'Else
'    ReDim fMatCoordJA(0)
'End If
'
'Set rsBPP = Nothing
'Set oBPP = Nothing
'
'Exit Sub
'ErrorCargaDatos:
'ReDim fMatCoordJA(0)
'MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub GenerarExcel()
'    Dim fs As Scripting.FileSystemObject
'    Dim xlsAplicacion As Excel.Application
'    Dim lsArchivo As String
'    Dim lsFile As String
'    Dim lsNomHoja As String
'    Dim xlsLibro As Excel.Workbook
'    Dim xlHoja1 As Excel.Worksheet
'    Dim lbExisteHoja As Boolean
'    Dim psArchivoAGrabarC As String
'    Dim lnExcel As Long
'
'    On Error GoTo ErrorGeneraExcelFormato
'
'    Set fs = New Scripting.FileSystemObject
'    Set xlsAplicacion = New Excel.Application
'
'    lsNomHoja = "BPP"
'    lsFile = "FormatoBPPCoordinadorJefeAgencia"
'
'    lsArchivo = "\spooler\" & "BPPCoordinadorJefeAgenciaGeneradoAlCierre" & Format(fgFecActual, "yyyymmdd") & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
'    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
'        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
'    Else
'        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
'        Exit Sub
'    End If
'
'    'Activar Hoja
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    lnExcel = 4
'    Dim sFormatoConta As String
'    sFormatoConta = "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ *  - ??_ ;_ @_ "
'
'    For i = 0 To UBound(fMatCoordJA)
'        xlHoja1.Cells(lnExcel + i, 1) = i + 1
'        xlHoja1.Cells(lnExcel + i, 2) = fMatCoordJA(i).Agencia
'        xlHoja1.Cells(lnExcel + i, 3).NumberFormat = "@"
'        xlHoja1.Cells(lnExcel + i, 3) = fMatCoordJA(i).cPersCod
'        xlHoja1.Cells(lnExcel + i, 4) = fMatCoordJA(i).comite
'        xlHoja1.Cells(lnExcel + i, 5) = fMatCoordJA(i).Usuario
'        xlHoja1.Cells(lnExcel + i, 6) = fMatCoordJA(i).Nombre
'        xlHoja1.Cells(lnExcel + i, 7) = fMatCoordJA(i).Cargo
'        xlHoja1.Cells(lnExcel + i, 8) = fMatCoordJA(i).CantAnalista
'        xlHoja1.Cells(lnExcel + i, 9) = fMatCoordJA(i).AnalistaBoni
'        xlHoja1.Cells(lnExcel + i, 10).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 10) = fMatCoordJA(i).PorBonificacion
'        xlHoja1.Cells(lnExcel + i, 11).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 11) = fMatCoordJA(i).SaldoCartera
'        xlHoja1.Cells(lnExcel + i, 12).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 12) = fMatCoordJA(i).SaldoCarteraCierre
'        xlHoja1.Cells(lnExcel + i, 13).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 13) = fMatCoordJA(i).PorSalCartVencJud
'        xlHoja1.Cells(lnExcel + i, 14).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 14) = fMatCoordJA(i).PorcSalVenJud
'        xlHoja1.Cells(lnExcel + i, 15).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 15) = fMatCoordJA(i).RendCaja
'        xlHoja1.Cells(lnExcel + i, 16).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 16) = fMatCoordJA(i).RendAG
'        xlHoja1.Cells(lnExcel + i, 17).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 17) = fMatCoordJA(i).Tope
'        xlHoja1.Cells(lnExcel + i, 18).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 18) = fMatCoordJA(i).PorBonificacionConf
'        xlHoja1.Cells(lnExcel + i, 19).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 19) = fMatCoordJA(i).BonoTotal
'
'        xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 1), xlHoja1.Cells(lnExcel + i, 19)).Borders.LineStyle = 1
'    Next i
'    xlHoja1.Range(xlHoja1.Cells(4, 19), xlHoja1.Cells(lnExcel + i - 1, 19)).Interior.Color = RGB(255, 255, 0)
'    xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 1), xlHoja1.Cells(lnExcel + i, 19)).EntireColumn.AutoFit
'
'    xlHoja1.SaveAs App.path & lsArchivo
'    psArchivoAGrabarC = App.path & lsArchivo
'    xlsAplicacion.Visible = True
'    xlsAplicacion.Windows(1).Visible = True
'    Set xlsAplicacion = Nothing
'    Set xlsLibro = Nothing
'    Set xlHoja1 = Nothing
'
'    MsgBox "Fromato Generado Satisfactoriamente en la ruta: " & psArchivoAGrabarC, vbInformation, "Aviso"
'
'    Exit Sub
'ErrorGeneraExcelFormato:
'    MsgBox err.Description, vbCritical, "Error a Generar El Formato Excel"
'End Sub
