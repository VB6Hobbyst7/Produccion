VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPBonoJT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Generar Bono Jefes Territoriales"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17160
   Icon            =   "frmCredBPPBonoJT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   17160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTBusqueda 
      Height          =   7770
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   13705
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Generar Bono Mensual"
      TabPicture(0)   =   "frmCredBPPBonoJT.frx":030A
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
         Left            =   15240
         TabIndex        =   6
         Top             =   7200
         Width           =   1170
      End
      Begin VB.Frame fraBusqueda 
         Caption         =   "Resultado"
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
         Width           =   16695
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
            Left            =   3480
            TabIndex        =   3
            Top             =   360
            Width           =   1170
         End
         Begin SICMACT.FlexEdit feAnalistas 
            Height          =   4815
            Left            =   120
            TabIndex        =   7
            Top             =   1515
            Width           =   16335
            _extentx        =   28813
            _extenty        =   8493
            cols0           =   13
            highlight       =   1
            encabezadosnombres=   "#-Zona-Usuario-Meta-Cierre-Meta-Cierre-Caja-Agencia-Agencia-Ag. Bonif.-Porc.-Total"
            encabezadosanchos=   "0-2000-1000-1500-1500-1500-1500-1200-1200-1000-1000-1000-1200"
            font            =   "frmCredBPPBonoJT.frx":0326
            font            =   "frmCredBPPBonoJT.frx":034E
            font            =   "frmCredBPPBonoJT.frx":0376
            font            =   "frmCredBPPBonoJT.frx":039E
            font            =   "frmCredBPPBonoJT.frx":03C6
            fontfixed       =   "frmCredBPPBonoJT.frx":03EE
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-C-R-R-R-R-R-R-R-R-R-R"
            formatosedit    =   "0-0-0-2-2-2-2-2-2-3-3-2-2"
            cantentero      =   15
            textarray0      =   "#"
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
         Begin VB.Label Label4 
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
            Left            =   11540
            TabIndex        =   11
            Top             =   1200
            Width           =   3030
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
            Left            =   9135
            TabIndex        =   10
            Top             =   1200
            Width           =   2410
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
            Left            =   6140
            TabIndex        =   9
            Top             =   1200
            Width           =   3015
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
            Left            =   3130
            TabIndex        =   8
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mes a generar:"
            Height          =   195
            Left            =   240
            TabIndex        =   5
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
            TabIndex        =   4
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
Attribute VB_Name = "frmCredBPPBonoJT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private fMatJefesTerritorial() As JefeTerritorialBPP
'Private i As Integer
'Private nIndex As Integer
'Private fgFecActual As Date
'
'Private Sub CargaControles()
'cmdExportar.Enabled = False
'MesActual
'lblMes.Caption = MesAnio(fgFecActual)
'End Sub
'
'Private Function ValidaDatos() As Boolean
'ValidaDatos = True
'End Function
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
'
'Private Sub cmdExportar_Click()
'GenerarExcel
'End Sub
'
'Private Sub cmdGenerar_Click()
'If ValidaDatos Then
'    If MsgBox("Estás Seguro de Generar el BPP?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'
'    LimpiaFlex feAnalistas
'    CargaDatos
'
'    If fMatJefesTerritorial(0).Usuario <> "" Then
'        For i = 0 To UBound(fMatJefesTerritorial)
'            feAnalistas.AdicionaFila
'            feAnalistas.TextMatrix(i + 1, 1) = fMatJefesTerritorial(i).Zona
'            feAnalistas.TextMatrix(i + 1, 2) = fMatJefesTerritorial(i).Usuario
'            feAnalistas.TextMatrix(i + 1, 3) = Format(fMatJefesTerritorial(i).SaldoCartera, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 4) = Format(fMatJefesTerritorial(i).SaldoCarteraCierre, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 5) = Format(fMatJefesTerritorial(i).PorSalCartVencJud * 100, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 6) = Format(fMatJefesTerritorial(i).PorcSalVenJud * 100, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 7) = Format(fMatJefesTerritorial(i).RendCaja * 100, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 8) = Format(fMatJefesTerritorial(i).Rend * 100, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 9) = fMatJefesTerritorial(i).CantAge
'            feAnalistas.TextMatrix(i + 1, 10) = fMatJefesTerritorial(i).AgeBoni
'            feAnalistas.TextMatrix(i + 1, 11) = Format(fMatJefesTerritorial(i).PorBoni * 100, "###," & String(15, "#") & "#0.00")
'            feAnalistas.TextMatrix(i + 1, 12) = Format(fMatJefesTerritorial(i).BonoTotal, "###," & String(15, "#") & "#0.00")
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
'
'Private Sub CargaDatos()
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
' On Error GoTo ErrorCargaDatos
'Set oBPP = New COMNCredito.NCOMBPPR
'Set rsBPP = oBPP.GenerarBPPJT
'
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    ReDim fMatJefesTerritorial(rsBPP.RecordCount - 1)
'    For i = 0 To rsBPP.RecordCount - 1
'        fMatJefesTerritorial(i).nZona = rsBPP!nZona
'        fMatJefesTerritorial(i).Zona = rsBPP!Zona
'        fMatJefesTerritorial(i).cPersCod = rsBPP!cPersCod
'        fMatJefesTerritorial(i).Usuario = rsBPP!Usuario
'        fMatJefesTerritorial(i).Nombre = rsBPP!Nombre
'        fMatJefesTerritorial(i).SaldoCartera = rsBPP!SaldoCartera
'        fMatJefesTerritorial(i).SaldoCarteraCierre = rsBPP!SaldoCarteraCierre
'        fMatJefesTerritorial(i).PorSalCartVencJud = rsBPP!PorSalCartVencJud
'        fMatJefesTerritorial(i).PorcSalVenJud = rsBPP!PorcSalVenJud
'        fMatJefesTerritorial(i).RendCaja = rsBPP!RendCaja
'        fMatJefesTerritorial(i).Rend = rsBPP!Rend
'        fMatJefesTerritorial(i).CantAge = rsBPP!CantAge
'        fMatJefesTerritorial(i).AgeBoni = rsBPP!AgeBoni
'        fMatJefesTerritorial(i).PorBoni = rsBPP!PorBoni
'        fMatJefesTerritorial(i).PorConfJF = rsBPP!PorConfJF
'        fMatJefesTerritorial(i).Tope = rsBPP!nTope
'        fMatJefesTerritorial(i).BonoTotal = rsBPP!BonoTotal
'        rsBPP.MoveNext
'    Next i
'Else
'    ReDim fMatJefesTerritorial(0)
'End If
'
'Set rsBPP = Nothing
'Set oBPP = Nothing
'
'Exit Sub
'ErrorCargaDatos:
'ReDim fMatJefesTerritorial(0)
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
'    lsFile = "FormatoBPPJefeTerritorial"
'
'    lsArchivo = "\spooler\" & "BPPJefeTerritorialGeneradoAlCierre" & Format(fgFecActual, "yyyymmdd") & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
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
'    For i = 0 To UBound(fMatJefesTerritorial)
'        xlHoja1.Cells(lnExcel + i, 1) = i + 1
'        xlHoja1.Cells(lnExcel + i, 2).NumberFormat = "@"
'        xlHoja1.Cells(lnExcel + i, 2) = fMatJefesTerritorial(i).cPersCod
'        xlHoja1.Cells(lnExcel + i, 3) = fMatJefesTerritorial(i).Zona
'        xlHoja1.Cells(lnExcel + i, 4) = fMatJefesTerritorial(i).Usuario
'        xlHoja1.Cells(lnExcel + i, 5) = fMatJefesTerritorial(i).Nombre
'        xlHoja1.Cells(lnExcel + i, 6).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 6) = fMatJefesTerritorial(i).SaldoCartera
'        xlHoja1.Cells(lnExcel + i, 7).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 7) = fMatJefesTerritorial(i).SaldoCarteraCierre
'        xlHoja1.Cells(lnExcel + i, 8).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 8) = fMatJefesTerritorial(i).PorSalCartVencJud
'        xlHoja1.Cells(lnExcel + i, 9).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 9) = fMatJefesTerritorial(i).PorcSalVenJud
'        xlHoja1.Cells(lnExcel + i, 10).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 10) = fMatJefesTerritorial(i).RendCaja
'        xlHoja1.Cells(lnExcel + i, 11).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 11) = fMatJefesTerritorial(i).Rend
'        xlHoja1.Cells(lnExcel + i, 12) = fMatJefesTerritorial(i).CantAge
'        xlHoja1.Cells(lnExcel + i, 13) = fMatJefesTerritorial(i).AgeBoni
'        xlHoja1.Cells(lnExcel + i, 14).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 14) = fMatJefesTerritorial(i).PorBoni
'        xlHoja1.Cells(lnExcel + i, 15).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 15) = fMatJefesTerritorial(i).Tope
'        xlHoja1.Cells(lnExcel + i, 16).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 16) = fMatJefesTerritorial(i).PorConfJF
'        xlHoja1.Cells(lnExcel + i, 17).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 17) = fMatJefesTerritorial(i).BonoTotal
'
'        xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 1), xlHoja1.Cells(lnExcel + i, 17)).Borders.LineStyle = 1
'    Next i
'    xlHoja1.Range(xlHoja1.Cells(4, 17), xlHoja1.Cells(lnExcel + i - 1, 17)).Interior.Color = RGB(255, 255, 0)
'    xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 1), xlHoja1.Cells(lnExcel + i, 17)).EntireColumn.AutoFit
'
'
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
