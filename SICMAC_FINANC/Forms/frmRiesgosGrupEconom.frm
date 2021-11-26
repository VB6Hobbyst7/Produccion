VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRiesgosGrupEconom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "+"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "frmRiesgosGrupEconom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraPeriodo 
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
      ForeColor       =   &H8000000D&
      Height          =   1080
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   2790
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
         Height          =   330
         Left            =   1575
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
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
         ItemData        =   "frmRiesgosGrupEconom.frx":030A
         Left            =   0
         List            =   "frmRiesgosGrupEconom.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Visible         =   0   'False
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
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   1560
         TabIndex        =   13
         Top             =   240
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio"
         Height          =   195
         Left            =   285
         TabIndex        =   12
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   285
         TabIndex        =   10
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   825
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4950
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
         Height          =   390
         Left            =   2235
         TabIndex        =   5
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
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
         Left            =   870
         TabIndex        =   4
         Top             =   255
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1080
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   2145
      Begin VB.TextBox txtPatrimonio 
         Height          =   285
         Left            =   135
         MaxLength       =   15
         TabIndex        =   2
         Top             =   570
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Patrimonio Efectivo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   1650
      End
   End
End
Attribute VB_Name = "frmRiesgosGrupEconom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim fsCodReport As String
Dim fdFechaDataRep As Date

Dim oCon As DConecta

Public Sub Inicio(ByVal psCodReport As String, ByVal pdFechaDataRep As Date)
    fsCodReport = psCodReport
    fdFechaDataRep = pdFechaDataRep
End Sub

Private Sub cmdProcesar_Click()
    Dim lsArchivo   As String
    Dim lbLibroOpen As Boolean
    Dim n           As Integer
    Dim lsFechaRep  As String
    Dim lsFec As String
    Dim lsFecIniMes As String
    
    lsFechaRep = txtFecha.Text 'Format(DateAdd("d", gdFecSis, -1 * Day(gdFecSis)), "mm/dd/yyyy")
    lsFec = Format(lsFechaRep, "yyyymmdd")
    
    Select Case fsCodReport
        Case gRiesgoSBSA190  ' Informe sobre el Grupo Economico de la Empresa
'   ALPA 20090929***************************************
'            lsArchivo = App.path & "\Spooler\Reporte19_" & lsFec & ".xls"
'            lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
'            If lbLibroOpen Then
'                Call GeneraRep19_GrupoEconomicoEmpresa(lsFechaRep)
'                ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
'                CargaArchivo "Reporte19_" & lsFec & ".xls", App.path & "\Spooler"
'            End If
          Call Reporte19GruposEconomicos(lsFechaRep)
'    ALPA 20090929*********************************************************
'        Case gRiesgoSBSA191  ' Informe sobre el Grupo Economico de la Empresa
'            lsArchivo = App.path & "\Spooler\Reporte19A_" & lsFec & ".xls"
'            lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
'            If lbLibroOpen Then
'                Call GeneraRep19A_GrupoEconomicoEmpresa(lsFechaRep)
'                ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
'                CargaArchivo "Reporte19A_" & lsFec & ".xls", App.path & "\Spooler"
'            End If
'   *************************************************************************
        Case gRiesgoSBSA200   ' Informacion de Clientes que representan riesgo unico
'            lsArchivo = App.path & "\Spooler\Reporte20_" & lsFec & ".xls"
'            lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
'            If lbLibroOpen Then
'
'                Set xlHoja1 = xlLibro.Worksheets(1)
'                'xlHoja1.Name = "Clientes"
'                Call GeneraRep20_ClientesRepresRiesgoUnico(lsFechaRep)
'                ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
'                CargaArchivo "Reporte20_" & lsFec & ".xls", App.path & "\Spooler"
'
'            End If
         Call Reporte20RiesgoUnico(lsFechaRep)
'        Case gRiesgoSBSA201   ' Informacion de Clientes que representan riesgo unico
'            lsArchivo = App.path & "\Spooler\Reporte20A_" & lsFec & ".xls"
'            lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
'            If lbLibroOpen Then
'                Call GeneraRep20A_ClientesRepresRiesgoUnico(lsFechaRep)
'                ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
'                CargaArchivo "Reporte20A_" & lsFec & ".xls", App.path & "\Spooler"
'            End If
            
        Case gRiesgoSBSA210  ' Financiamientos a Vinculados a la Empresa
            If Val(txtPatrimonio.Text) = 0 Then
                MsgBox "Ingrese Patrimonio Efectivo", vbInformation, "Aviso"
            Else
'                If Len(Trim(cboMes.Text)) = 0 Then
'                    MsgBox "Seleccione Mes", vbInformation, "Aviso"
'                Else
'                    If Val(txtAnio.Text) = 0 Then
'                        MsgBox "Ingrese año", vbInformation, "Aviso"
'                    Else
                        If Val(txtTC.Text) = 0 Then
                            MsgBox "Ingrese Tipo de Cambio", vbInformation, "Aviso"
                        Else
                            lsArchivo = App.path & "\Spooler\Reporte21_" & lsFec & ".xls"
                            lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
                            If lbLibroOpen Then
                                Call GeneraRep21_ClientesRepresRiesgoUnico
                                ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
                                CargaArchivo "Reporte21_" & lsFec & ".xls", App.path & "\Spooler"
                            End If
                        End If
                    End If
'                End If
'            End If
    End Select
End Sub

Private Sub cmdSalir_Click()
    CierraConexion
    Unload Me
End Sub

Private Sub Form_Load()
    Set oCon = New DConecta
    Me.Caption = "Reporte Riesgos"
    Select Case fsCodReport
        Case gRiesgoSBSA210
            FraPeriodo.Visible = True
        Case gRiesgoSBSA190
            FraPeriodo.Visible = True
            Frame1.Visible = False
             Me.Caption = "Reporte 19 y 19 A: Vinculados y Grupos Económicos de la Empresa"
        Case gRiesgoSBSA191
            FraPeriodo.Visible = True
            Frame1.Visible = False
        Case gRiesgoSBSA200
            Frame1.Visible = False
            'ALPA 20090929****************************
            FraPeriodo.Visible = True
            '*****************************************
            Me.Caption = "Reporte 20 y 20 A: Vinculados y Grupos Económicos"
        Case gRiesgoSBSA201
            Frame1.Visible = False
        Case Else
            FraPeriodo.Visible = False
    End Select
    CentraForm Me
End Sub
' ALPA 20090929************************************************
Private Sub Reporte19GruposEconomicos(ByVal pdFecha As Date)
 
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet

    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja1  As String
    Dim lsNomHoja2  As String
    Dim lsNomHoja3  As String
    Dim lsNomHoja4  As String
    Dim lsNombreAgencia As String
    Dim lsCodAgencia As String
    Dim lsTipoGarantia As Integer
    Dim lsMes As String
    Dim lnContador As Integer
    Dim lnContadorMatrix As Integer
    Dim lnPosY1 As Integer
    Dim lnPosY2 As Integer
    Dim lnPosYInicial19 As Integer
    Dim lnPosYVincula19 As Integer
    Dim lsArchivo As String
    Dim lnTipoSalto As Integer
    Dim oGrupEc As DGrupoEco
    Set oGrupEc = New DGrupoEco
    Dim rs As ADODB.Recordset
    Dim lnNumColumns As Integer
    Dim sMatrixVinculados() As String
    Dim cTipPer As String
    Dim i As Integer
    Dim j, nContPerso As Integer
    Dim lsPersCodEmpresa As String
    Set rs = New ADODB.Recordset
    Set rs = oGrupEc.ListarDatosGrupoEconomicoxGrupoRepo19y20(1)
    Set oGrupEc = Nothing
    Dim nx As Integer
    Dim ny As Integer
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "Reporte19"
    lsNomHoja1 = "Reporte19"
    lsArchivo1 = "\spooler\Reporte19" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja1 Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja1
    End If
    
    xlHoja1.Cells(7, 3) = Format(pdFecha, "DD.MM.YYYY")
    lnPosYInicial19 = 10
    i = 0
    Dim nEncontrado As Integer
    If Not (rs.EOF And rs.BOF) Then
            lnContadorMatrix = 1
            While Not rs.EOF
                nEncontrado = 0
                If i > 0 Then
                    For j = 1 To i
                        If sMatrixVinculados(1, j) = rs!cPersCod Then
                            nEncontrado = 1
                        End If
                    Next j
                End If
                If nEncontrado = 0 Then
                    i = i + 1
                    ReDim Preserve sMatrixVinculados(1 To 1, 1 To i)
                    sMatrixVinculados(1, i) = rs!cPersCod
                    xlHoja1.Cells(lnPosYInicial19, 1) = i
                    xlHoja1.Cells(lnPosYInicial19, 2) = rs!cPersEmpresa
                    xlHoja1.Cells(lnPosYInicial19, 3) = IIf(Mid(rs!P1RUC, 1, 1) = "2", "JUR", "NAT")
                    xlHoja1.Cells(lnPosYInicial19, 4) = rs!cPersCIIU
                    xlHoja1.Cells(lnPosYInicial19, 5) = rs!cDirecciEmpresa
                    xlHoja1.Cells(lnPosYInicial19, 6) = IIf(Trim(rs!P1DNI) = "-", "DNI", rs!P1DNI)
                    xlHoja1.Cells(lnPosYInicial19, 7) = IIf(Trim(rs!P1DNI) = "", "-", rs!P1DNI)
                    xlHoja1.Cells(lnPosYInicial19, 8) = IIf(Trim(rs!P1RUC) = "", "-", rs!P1RUC)
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 1), xlHoja1.Cells(lnPosYInicial19, 9)).Borders.LineStyle = 1
                    lnPosYInicial19 = lnPosYInicial19 + 1
                    'nContPerso = 0
                End If
                rs.MoveNext
            Wend
    End If
   
    
   
    lsNomHoja1 = "Reporte19A"
    For Each xlHoja1 In xlsLibro.Worksheets
    If xlHoja1.Name = lsNomHoja1 Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
    End If
    Next
   
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja1
    End If
    nContPerso = 0
    lnPosYInicial19 = 8
    lnPosYVincula19 = 15
    xlHoja1.Cells(7, 3) = Format(pdFecha, "DD.MM.YYYY")
    If i >= 1 Then
    rs.MoveFirst
    End If
    If Not (rs.EOF And rs.BOF) Then
            lnContadorMatrix = 1
            While Not rs.EOF
                nEncontrado = 0
                If i > 0 Then
                    For j = 1 To i
                        If lsPersCodEmpresa = rs!cPersCod Then
                            nEncontrado = 1
                        End If
                    Next j
                End If
                If nEncontrado = 0 Then
                    If lnPosYVincula19 = 15 Then
                        lnPosYInicial19 = lnPosYInicial19 + 1
                    Else
                        lnPosYInicial19 = lnPosYVincula19 + 15
                        ny = lnPosYInicial19 + 4
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 1), xlHoja1.Cells(ny, 7)).Borders.LineStyle = 1
                        ny = lnPosYInicial19 + 6
                        xlHoja1.Cells(ny, 1) = xlHoja1.Cells(15, 1)
                        xlHoja1.Cells(ny, 2) = xlHoja1.Cells(15, 2)
                        xlHoja1.Cells(ny, 3) = xlHoja1.Cells(15, 3)
                        xlHoja1.Cells(ny, 4) = xlHoja1.Cells(15, 4)
                        xlHoja1.Cells(ny, 5) = xlHoja1.Cells(15, 5)
                        xlHoja1.Cells(ny, 6) = xlHoja1.Cells(15, 6)
                        xlHoja1.Cells(ny, 7) = xlHoja1.Cells(15, 7)
                        xlHoja1.Cells(ny, 8) = xlHoja1.Cells(15, 8)
                        xlHoja1.Cells(ny, 9) = xlHoja1.Cells(15, 9)
                        xlHoja1.Cells(ny, 10) = xlHoja1.Cells(15, 10)
                        xlHoja1.Cells(ny, 11) = xlHoja1.Cells(15, 11)
                        xlHoja1.Range(xlHoja1.Cells(ny, 1), xlHoja1.Cells(ny, 11)).Borders.LineStyle = 1
                    End If
             
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 3), xlHoja1.Cells(lnPosYInicial19 + 6, 3)).HorizontalAlignment = xlLeft
                   
                    lsPersCodEmpresa = rs!cPersCod
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 3), xlHoja1.Cells(lnPosYInicial19, 7)).Merge True
                    xlHoja1.Cells(lnPosYInicial19, 1) = xlHoja1.Cells(9, 1)
                    xlHoja1.Cells(lnPosYInicial19, 2) = xlHoja1.Cells(9, 2)
                    xlHoja1.Cells(lnPosYInicial19, 3) = rs!cPersEmpresa
                    lnPosYInicial19 = lnPosYInicial19 + 1
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 3), xlHoja1.Cells(lnPosYInicial19, 7)).Merge True
                    xlHoja1.Cells(lnPosYInicial19, 1) = xlHoja1.Cells(10, 1)
                    xlHoja1.Cells(lnPosYInicial19, 2) = xlHoja1.Cells(10, 2)
                    xlHoja1.Cells(lnPosYInicial19, 3) = rs!cPersCodSbs1
                    lnPosYInicial19 = lnPosYInicial19 + 1
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 3), xlHoja1.Cells(lnPosYInicial19, 7)).Merge True
                    xlHoja1.Cells(lnPosYInicial19, 1) = xlHoja1.Cells(11, 1)
                    xlHoja1.Cells(lnPosYInicial19, 2) = xlHoja1.Cells(11, 2)
                    xlHoja1.Cells(lnPosYInicial19, 3) = IIf(rs!P1RUC = "", "-", rs!P1RUC)
                    lnPosYInicial19 = lnPosYInicial19 + 1
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 3), xlHoja1.Cells(lnPosYInicial19, 7)).Merge True
                    xlHoja1.Cells(lnPosYInicial19, 1) = xlHoja1.Cells(12, 1)
                    xlHoja1.Cells(lnPosYInicial19, 2) = xlHoja1.Cells(12, 2)
                    xlHoja1.Cells(lnPosYInicial19, 3) = rs!cDirecciEmpresa
                    lnPosYInicial19 = lnPosYInicial19 + 1
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 3), xlHoja1.Cells(lnPosYInicial19, 7)).Merge True
                    xlHoja1.Cells(lnPosYInicial19, 1) = xlHoja1.Cells(13, 1)
                    xlHoja1.Cells(lnPosYInicial19, 2) = xlHoja1.Cells(13, 2)
                    If rs!nRepresentanteLegal = 1 Then
                        xlHoja1.Cells(lnPosYInicial19, 3) = rs!cPersVinculado
                    End If
                    nContPerso = 0
                End If
                    nContPerso = nContPerso + 1
                    If nEncontrado = 0 And lnPosYVincula19 > 15 Then
                        lnPosYVincula19 = lnPosYVincula19 + 22
                    Else
                        lnPosYVincula19 = lnPosYVincula19 + 1
                    End If
                    xlHoja1.Cells(lnPosYVincula19, 1) = nContPerso
                    xlHoja1.Cells(lnPosYVincula19, 2) = rs!cPersVinculado
                    xlHoja1.Cells(lnPosYVincula19, 3) = rs!cPersCodSbs2
                    xlHoja1.Cells(lnPosYVincula19, 4) = IIf(Mid(rs!P2RUC, 1, 1) = "2", "JUR", "NAT")
                    xlHoja1.Cells(lnPosYVincula19, 5) = rs!PT2DNI
                    xlHoja1.Cells(lnPosYVincula19, 6) = IIf(Trim(rs!P2DNI) = "", "-", rs!P2DNI)
                    xlHoja1.Cells(lnPosYVincula19, 7) = IIf(Trim(rs!P2RUC) = "", "-", rs!P2RUC)
                    xlHoja1.Cells(lnPosYVincula19, 8) = rs!cDirecciVinculado
                    xlHoja1.Cells(lnPosYVincula19, 10) = IIf(rs!nPorcenOtro > 0, CStr(rs!nPorcenOtro) & "%", "")
                    xlHoja1.Cells(lnPosYVincula19, 10) = rs!nCargo
                    xlHoja1.Cells(lnPosYVincula19, 11) = rs!nCargoOtro
                    xlHoja1.Range(xlHoja1.Cells(lnPosYVincula19, 1), xlHoja1.Cells(lnPosYVincula19, 11)).Borders.LineStyle = 1
                    
                rs.MoveNext
            Wend
    End If
    
    rs.Close
    Set rs = Nothing
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

End Sub
Private Sub Reporte20RiesgoUnico(ByVal pdFecha As Date)
 
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet

    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja1  As String
    Dim lsNomHoja2  As String
    Dim lsNomHoja3  As String
    Dim lsNomHoja4  As String
    Dim lsNombreAgencia As String
    Dim lsCodAgencia As String
    Dim lsTipoGarantia As Integer
    Dim lsMes As String
    Dim lnContador As Long
    Dim lnContadorMatrix As Long
    Dim lnPosY1 As Integer
    Dim lnPosY2 As Integer
    Dim lnPosYInicial20 As Long
    Dim lnPosYVincula19 As Long
    Dim lsArchivo As String
    Dim lnTipoSalto As Long
    Dim oGrupEc As DGrupoEco
    Set oGrupEc = New DGrupoEco
    Dim rs As ADODB.Recordset
    Dim lnNumColumns As Long
    Dim sMatrixVinculados() As String
    Dim sMatrixVinculados2() As String
    Dim sMatrixPropiedadDirecIn() As String
    Dim sMatrixEmpresas() As String
    Dim sMatrixEmpresas2() As String
    Dim sMatrixEmpresas3() As String
    Dim sMatrixEmpresas4() As String
    Dim cTipPer As String
    Dim i, ix, ij, ipdi As Integer
    Dim j, k, H, L As Integer
    Dim m, n As Integer
    Dim lsPersCodEmpresa As String
    Dim lsPersCodEmpresaAnterior As String
    Dim nPosEmpresa As Long
    Dim encontrado20 As Long
    Dim encontrado20A As Long
    Dim encontrado20AS As Long
    Dim lnOrden As Long
    Dim lnContadorOrden As Long
    Dim lsPersCodTemp As String
    Dim lsPropiedadDirecta, lsPropiedadIndirecta As String
    Set rs = New ADODB.Recordset
    'Call oGrupEc.InsertarPersGrupoEconomicoSaldo(pdFecha, CDbl(txtTC.Text))
    Set rs = oGrupEc.ListarDatosGrupoEconomicoxRURepo20Nuevo(1, pdFecha, CDbl(txtTC.Text))
    Set oGrupEc = Nothing
    Dim nx As Long
    Dim ny As Long
    Dim nTJ As Long
    Dim oPerPDI As DGrupoEco
    Dim RsProDI As ADODB.Recordset
    Dim P, R As Long
    Set RsProDI = New ADODB.Recordset
    Set oPerPDI = New DGrupoEco
    Set RsProDI = oPerPDI.ObtenerDatosPropiedadDirectaEIndirectaNuevo(pdFecha)
    If Not (RsProDI.EOF And RsProDI.BOF) Then
            While Not RsProDI.EOF
                ipdi = ipdi + 1
                ReDim Preserve sMatrixPropiedadDirecIn(1 To 4, 1 To ipdi)
                sMatrixPropiedadDirecIn(1, ipdi) = RsProDI!nPorcenOtro
                sMatrixPropiedadDirecIn(2, ipdi) = RsProDI!cPersCod
                sMatrixPropiedadDirecIn(3, ipdi) = RsProDI!cPersCodOtro
                sMatrixPropiedadDirecIn(4, ipdi) = RsProDI!nTipo
                RsProDI.MoveNext
            Wend
    End If
    
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "Reporte20"
    lsNomHoja1 = "Reporte20"
    lsArchivo1 = "\spooler\Reporte20" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja1 Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja1
    End If
    
    xlHoja1.Cells(7, 3) = Format(pdFecha, "DD.MM.YYYY")
    lnPosYInicial20 = 8
    i = 0
    Dim nEncontrado As Integer
    Dim nEncontradoVin As Integer
    nTJ = 0
    If Not (rs.EOF And rs.BOF) Then
            lnContadorMatrix = 1
            While Not rs.EOF
            If (((rs!SalCreEmpre > 0 And rs!nCantGrupo > 0 And rs!nTipo = 1) Or (rs!nSaldo > 0 And rs!nCantGrupo > 0 And rs!nTipo = 2) Or (rs!nCantClienteSaldoGrupoEconomico > 1 And rs!nTipo = 1)) And ((rs!nOrden = 1) Or (rs!nOrden = 2))) Or (rs!nSaldo > 0 And rs!SalCreEmpre > 0 And rs!nOrden = 3) Then
            'If ((rs!SalCreEmpre > 0 And rs!nCantGrupo > 0 And rs!nTipo = 1) Or (rs!nSaldo > 0 And rs!nCantGrupo > 0 And rs!nTipo = 2) Or ((rs!nSaldo > 0 And rs!SalCreEmpre > 0 And rs!nOrden = 3))) Then
            'If ((rs!salCreEmpre > 0 And rs!nSalCreEmpreTotal > 0)) Then
            'If rs!SalCreEmpre > 0 And (rs!nSaldoGen > 0 Or (rs!nPorcenOtro > 0 Or Trim(rs!nRepresentanteLegal) = 1)) Then
                nEncontrado = 0
                nEncontradoVin = 0
                If i > 0 Then
                    For j = 1 To i
                        If Trim(sMatrixVinculados(1, j)) = Trim(rs!cPersCod) And nEncontrado = 0 Then
                            nEncontrado = 1
                        End If
                        If Trim(sMatrixVinculados(1, j)) = Trim(rs!cPersCodOtro) And nEncontradoVin = 0 Then
                            nEncontradoVin = 1
                        End If
                        If nEncontrado = 1 And nEncontradoVin = 1 Then
                            Exit For
                        End If
                    Next j
                End If
                'ALPA 20141210
                'If ((nEncontrado = 0 Or nEncontradoVin = 0) And ((rs!nSaldoGen > 0 Or ((rs!nSaldo > 0 And rs!SalCreEmpre > 0 And rs!nOrden = 3) And nEncontrado = 0)) Or (rs!SalCreEmpre > 0 And rs!nSalCreEmpreTotal > 0))) Or ((nEncontrado = 0 Or nEncontradoVin = 0) And ((rs!nSaldoGen > 0 Or (rs!SalCreEmpre > 0 And rs!nSalCreEmpreTotal > 0)) And (rs!nPorcenOtro > 0 Or Trim(rs!nRepresentanteLegal) = 1))) Then
                If (((nEncontrado = 0 Or nEncontradoVin = 0) And ((rs!nSaldoGen > 0 Or (rs!nCantClienteSaldoGrupoEconomico > 1 And nEncontrado = 0)) Or (rs!SalCreEmpre > 0 And rs!nSalCreEmpreTotal > 0))) Or ((nEncontrado = 0 Or nEncontradoVin = 0) And ((rs!nSaldoGen > 0 Or (rs!SalCreEmpre > 0 And rs!nSalCreEmpreTotal > 0)) And (rs!nPorcenOtro > 0 Or Trim(rs!nRepresentanteLegal) = 1 Or rs!nCargo <> 0))) And ((rs!nOrden = 1) Or (rs!nOrden = 2))) Or ((rs!nSaldo > 0 And rs!SalCreEmpre > 0 And rs!nOrden = 3) And (nEncontrado = 0 Or nEncontradoVin = 0)) Then
                'If ((nEncontrado = 0 Or nEncontradoVin = 0) And ((rs!nSaldoGen > 0 ) Or (rs!salCreEmpre > 0 And rs!nSalCreEmpreTotal > 0))) Or ((nEncontrado = 0 Or nEncontradoVin = 0) And ((rs!nSaldoGen > 0 Or (rs!salCreEmpre > 0 And rs!nSalCreEmpreTotal > 0)) And (rs!nPorcenOtro > 0 Or Trim(rs!nRepresentanteLegal) = 1))) Then
                        'If nEncontrado = 0 And ((Trim(rs!P1DNI) <> "" And rs!nSaldo > 0 And rs!SalCreEmpre > 0)) Then
                        'If (nEncontrado = 0 And ((Trim(rs!P1DNI) <> "" And rs!nSaldo > 0) Or Trim(rs!P1DNI) = "")) Then
                        'If ((rs!salCreEmpre > 0 And rs!nSalCreEmpreTotal > 0) And nEncontrado = 0) Then
                        If ((((rs!SalCreEmpre > 0 And rs!nSalCreEmpreTotal > 0) And nEncontrado = 0) Or (rs!nSaldo > 0 And rs!SalCreEmpre > 0 And nEncontrado = 0 And rs!nTipo = 1) And ((rs!nOrden = 1) Or (rs!nOrden = 2)))) Or ((rs!nSaldo > 0 And rs!SalCreEmpre > 0 And rs!nOrden = 3) And (nEncontrado = 0)) Then
                        'If ((rs!SalCreEmpre > 0 And rs!nSalCreEmpreTotal > 0) And nEncontrado = 0) Or ((rs!nSaldo > 0 And rs!SalCreEmpre > 0 And rs!nOrden = 3) > 1 And nEncontrado = 0 And rs!nTipo = 1) Then
                        i = i + 1
                        ix = ix + 1
                        If nTJ = 1 Then
                            lnPosYInicial20 = lnPosYInicial20 + 2
                        Else
                            lnPosYInicial20 = lnPosYInicial20 + 1
                        End If
                        nTJ = 0
                        ReDim Preserve sMatrixVinculados(1 To 2, 1 To i)
                        ReDim Preserve sMatrixEmpresas4(1 To 2, 1 To i)
                        sMatrixEmpresas4(1, i) = rs!cPersCod
                        sMatrixEmpresas4(2, i) = ix
                        
                        sMatrixVinculados(1, i) = rs!cPersCod
                        sMatrixVinculados(2, i) = ix
                        If IIf(Mid(rs!P1RUC, 1, 1) = "2", "JUR", "NAT") = "NAT" Then
                            lsPersCodEmpresa = ""
                            lsPersCodEmpresaAnterior = ""
                        Else
                            lsPersCodEmpresaAnterior = Trim(rs!cPersCod)
                        End If
                        xlHoja1.Cells(lnPosYInicial20, 1) = ix & ".-"
                        xlHoja1.Cells(lnPosYInicial20, 2) = "INFORMACION DEL CLIENTE"
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "Nombre, razon o denominación Social"
                        xlHoja1.Cells(lnPosYInicial20, 3) = rs!cPersEmpresa
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "Codigo SBS"
                        xlHoja1.Cells(lnPosYInicial20, 3) = rs!cPersCodSbs1
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "Tipo de Persona"
                        xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Mid(rs!P1RUC, 1, 1) = "2", "JUR", "NAT")
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "Documento de Identidad y Número"
                        xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Trim(rs!P1DNI) = "", "", "DNI - ") & IIf(Trim(rs!P1DNI) = "", "", rs!P1DNI)
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "RUC"
                        xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Trim(rs!P1RUC) = "3", "-", IIf(Trim(rs!P1RUC) = "", "-", rs!P1RUC))
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "Direccción"
                        xlHoja1.Cells(lnPosYInicial20, 3) = rs!cDirecciEmpresa
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "Representante Legal"
                        xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Trim(rs!P1DNI) = "", rs!cPersNombreReprLegal, "")
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        End If
                         
                        'If (IIf(Mid(rs!P1RUC, 1, 1) = "2", "JUR", "NAT")) = "NAT" And nEncontradoVin = 0 And nEncontrado = 0 And Trim(rs!cPersCodOtro) <> Trim(rs!cPersCod) Then
                        If (IIf(Mid(rs!P1RUC, 1, 1) = "2", "JUR", "NAT")) = "NAT" And Trim(rs!cPersCodOtro) <> Trim(rs!cPersCod) Then
                            If rs!nSaldo > 0 And nEncontradoVin = 0 And nEncontrado = 0 Then
                                i = i + 1
                                ix = ix + 1
                                ReDim Preserve sMatrixVinculados(1 To 2, 1 To i)
                                
                                If IIf(Mid(rs!P1RUC, 1, 1) = "2", "JUR", "NAT") = "NAT" Then
                                    lsPersCodEmpresa = ""
                                    lsPersCodEmpresaAnterior = ""
                                Else
                                    lsPersCodEmpresaAnterior = Trim(rs!cPersCodOtro)
                                End If
                                
                                ReDim Preserve sMatrixEmpresas4(1 To 2, 1 To i)
                                sMatrixEmpresas4(1, i) = rs!cPersCod
                                sMatrixEmpresas4(2, i) = ix
                                
                                sMatrixVinculados(1, i) = rs!cPersCodOtro
                                sMatrixVinculados(2, i) = ix
                                xlHoja1.Cells(lnPosYInicial20, 1) = ix & ".-"
                                xlHoja1.Cells(lnPosYInicial20, 2) = "INFORMACION DEL CLIENTE"
                                lnPosYInicial20 = lnPosYInicial20 + 1
                                xlHoja1.Cells(lnPosYInicial20, 2) = "Nombre, razon o denominación Social"
                                xlHoja1.Cells(lnPosYInicial20, 3) = rs!cPersVinculado
                                xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                                lnPosYInicial20 = lnPosYInicial20 + 1
                                xlHoja1.Cells(lnPosYInicial20, 2) = "Codigo SBS"
                                xlHoja1.Cells(lnPosYInicial20, 3) = rs!cPersCodSbs2
                                xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                                lnPosYInicial20 = lnPosYInicial20 + 1
                                xlHoja1.Cells(lnPosYInicial20, 2) = "Tipo de Persona"
                                xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Mid(rs!P2RUC, 1, 1) = "2", "JUR", "NAT")
                                xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                                lnPosYInicial20 = lnPosYInicial20 + 1
                                xlHoja1.Cells(lnPosYInicial20, 2) = "Documento de Identidad y Número"
                                xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Trim(rs!P2DNI) = "", "", "DNI - ") & IIf(Trim(rs!P2DNI) = "", "", rs!P2DNI)
                                xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                                lnPosYInicial20 = lnPosYInicial20 + 1
                                xlHoja1.Cells(lnPosYInicial20, 2) = "RUC"
                                xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Trim(rs!P2RUC) = "3", "-", IIf(Trim(rs!P2RUC) = "", "-", rs!P2RUC))
                                xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                                lnPosYInicial20 = lnPosYInicial20 + 1
                                xlHoja1.Cells(lnPosYInicial20, 2) = "Dirección"
                                xlHoja1.Cells(lnPosYInicial20, 3) = rs!cDirecciVinculado
                                xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                                lnPosYInicial20 = lnPosYInicial20 + 1
                                xlHoja1.Cells(lnPosYInicial20, 2) = "Representante Legal"
                                xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Trim(rs!P1DNI) = "", rs!cPersNombreReprLegal, "")
                                xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                                lnPosYInicial20 = lnPosYInicial20 + 1
                            End If
                        ElseIf (IIf(Mid(rs!P1RUC, 1, 1) = "2", "JUR", "NAT")) = "JUR" And (rs!nPorcenOtro > 0 Or Trim(rs!nRepresentanteLegal) = 1 Or rs!nCargo <> 0) And Trim(rs!P1DNI) = "" Then
                        'ElseIf (IIf(Mid(rs!P1RUC, 1, 1) = "2", "JUR", "NAT")) = "JUR" Then
                            If (lsPersCodEmpresaAnterior) = Trim(rs!cPersCod) Then
                            If Trim(lsPersCodEmpresa) <> Trim(rs!cPersCod) Then
                                xlHoja1.Cells(lnPosYInicial20, 2) = "2. ACCIONISTAS, DIRECTORES, GERENTES, PRINCIPALES FUNCIONARIOS Y ASESORES"
                                lnPosYInicial20 = lnPosYInicial20 + 2
                                xlHoja1.Cells(lnPosYInicial20, 2) = "Nombre"
                                xlHoja1.Cells(lnPosYInicial20, 3) = "Cod Sbs"
                                xlHoja1.Cells(lnPosYInicial20, 4) = "TIP PER"
                                xlHoja1.Cells(lnPosYInicial20, 5) = "TIP DOC"
                                xlHoja1.Cells(lnPosYInicial20, 6) = "NRO DOC"
                                xlHoja1.Cells(lnPosYInicial20, 7) = "RUC"
                                xlHoja1.Cells(lnPosYInicial20, 8) = "Residencia"
                                xlHoja1.Cells(lnPosYInicial20, 9) = "ACC"
                                xlHoja1.Cells(lnPosYInicial20, 10) = "CAR"
                                xlHoja1.Cells(lnPosYInicial20, 11) = "OTR CAR"
                                xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 11)).Borders.LineStyle = 1
                             End If
                                lsPersCodEmpresa = rs!cPersCod
                                'i = i + 1
                                lnPosYInicial20 = lnPosYInicial20 + 1
'                                ReDim Preserve sMatrixVinculados(1 To 2, 1 To i)
                                xlHoja1.Cells(lnPosYInicial20, 2) = rs!cPersVinculado
                                xlHoja1.Cells(lnPosYInicial20, 3) = rs!cPersCodSbs2
                                xlHoja1.Cells(lnPosYInicial20, 4) = IIf(Mid(rs!P2RUC, 1, 1) = "2", "JUR", "NAT")
                                xlHoja1.Cells(lnPosYInicial20, 5) = IIf(Trim(rs!P2DNI) = "", "", "DNI")
                                xlHoja1.Cells(lnPosYInicial20, 6) = IIf(Trim(rs!P2DNI) = "", "", rs!P2DNI)
                                xlHoja1.Cells(lnPosYInicial20, 7) = IIf(Trim(rs!P2RUC) = "", "-", rs!P2RUC)
                                xlHoja1.Cells(lnPosYInicial20, 8) = rs!cDirecciVinculado
                                xlHoja1.Cells(lnPosYInicial20, 9) = IIf(rs!nPorcenOtro > 0, CStr(rs!nPorcenOtro) & "%", "")
                                xlHoja1.Cells(lnPosYInicial20, 10) = rs!nCargo
                                xlHoja1.Cells(lnPosYInicial20, 11) = rs!nCargoOtro
                                xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 11)).Borders.LineStyle = 1
                                nTJ = 1
                       End If
                       End If
                    'End If
                    End If
                    End If
                rs.MoveNext
            Wend
    End If
    
    
    lsNomHoja1 = "Reporte20A"
    For Each xlHoja1 In xlsLibro.Worksheets
    If xlHoja1.Name = lsNomHoja1 Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
    End If
    Next
 
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja1
    End If
    If i >= 1 Then
        rs.MoveFirst
    End If
    ij = i
    m = 0
    H = 0
    xlHoja1.Cells(7, 3) = Format(pdFecha, "DD.MM.YYYY")
     If Not (rs.EOF And rs.BOF) Then
            lnContadorMatrix = 1
            While Not rs.EOF
                nEncontrado = 0
                nEncontradoVin = 0
                If m > 0 Then
                    For j = 1 To m
                        If Trim(sMatrixVinculados(1, j)) = Trim(rs!cPersCod) And nEncontrado = 1 Then
                            nEncontrado = 1
                        End If
                        If sMatrixVinculados(1, j) = Trim(rs!cPersCodOtro) And nEncontradoVin = 1 Then
                            nEncontradoVin = 1
                        End If
                        If nEncontrado = 1 And nEncontradoVin = 1 Then
                            Exit For
                        End If
                    Next j
                End If
              
                If nEncontrado = 0 Or nEncontradoVin = 0 Then
   
                            If Trim(lsPersCodEmpresa) <> Trim(rs!cPersCod) Then
                                 H = H + 1
                                 ReDim Preserve sMatrixEmpresas(1 To 2, 1 To H)
                                  sMatrixEmpresas(1, H) = Trim(rs!cPersCod)
                                  sMatrixEmpresas(2, H) = H
                            End If
                            lsPersCodEmpresa = rs!cPersCod
                End If
                rs.MoveNext
            Wend
    End If
 lsPersCodEmpresa = ""
 lsPersCodTemp = ""
 lnContadorOrden = 0
    If i >= 1 Then
        rs.MoveFirst
    End If
    lnPosYInicial20 = 10
    i = 0
    ix = 0
    H = 0
    If Not (rs.EOF And rs.BOF) Then
            lnContadorMatrix = 1
            While Not rs.EOF
                nEncontrado = 0
                nEncontradoVin = 0
                If H > 0 Then
                    For j = 1 To H
                        If Trim(sMatrixVinculados2(1, j)) = Trim(rs!cPersCod) And nEncontrado = 1 Then
                            nEncontrado = 1
                        End If
                        If Trim(sMatrixVinculados2(1, j)) = Trim(rs!cPersCodOtro) And nEncontradoVin = 1 Then
                            nEncontradoVin = 1
                        End If
                        If nEncontrado = 1 And nEncontradoVin = 1 Then
                            Exit For
                        End If
                    Next j
                End If
                encontrado20 = 0
                encontrado20A = 0
                encontrado20AS = 0
                
                For R = 1 To ij
                    If Trim(sMatrixVinculados(1, R)) = Trim(rs!cPersCod) Then
                        encontrado20 = 1
                        lnOrden = CInt(sMatrixVinculados(2, R))
                        Exit For
                    End If
                Next R
                If encontrado20 = 1 Then
                    For R = 1 To i
                        If Trim(sMatrixVinculados2(3, R)) = Trim(rs!cPersCod) Then
                            If Trim(sMatrixVinculados2(1, R)) = Trim(rs!cPersCodOtro) Then
                                encontrado20A = 1
                                Exit For
                            End If
                        End If
                    Next R
                End If
              
                If (nEncontrado = 0 Or nEncontradoVin = 0) And encontrado20 = 1 And encontrado20A = 0 And encontrado20AS = 0 Then
                    
                        If Trim(rs!cPersCodOtro) <> Trim(rs!cPersCod) Then
                            lnPosYInicial20 = lnPosYInicial20 + 1
                            i = i + 1
                            ReDim Preserve sMatrixVinculados2(1 To 3, 1 To i)
                            sMatrixVinculados2(1, i) = Trim(rs!cPersCodOtro)
                            sMatrixVinculados2(2, i) = i
                            sMatrixVinculados2(3, i) = Trim(rs!cPersCod)
                            lsPersCodTemp = Trim(rs!cPersCod)
                            xlHoja1.Cells(lnPosYInicial20, 1) = lnOrden
                            If Trim(lsPersCodEmpresa) <> Trim(rs!cPersCod) Then
                                'ix = ix + 1
                                'xlHoja1.Cells(lnPosYInicial20, 1) = lnOrden 'ix
                                xlHoja1.Cells(lnPosYInicial20, 2) = rs!cPersCodSbs1
                            End If

                            xlHoja1.Cells(lnPosYInicial20, 3) = rs!cPersCodSbs2
                            xlHoja1.Cells(lnPosYInicial20, 4) = rs!cPersVinculado
                            xlHoja1.Cells(lnPosYInicial20, 5) = "CIIU"
                            xlHoja1.Cells(lnPosYInicial20, 6) = rs!cDirecciVinculado
                            xlHoja1.Cells(lnPosYInicial20, 7) = IIf(Mid(rs!P2RUC, 1, 1) = "2", "JUR", "NAT")
                            xlHoja1.Cells(lnPosYInicial20, 8) = IIf(Trim(rs!P2DNI) = "", "", "DNI")
                            xlHoja1.Cells(lnPosYInicial20, 9) = IIf(Trim(rs!P2DNI) = "", "", rs!P2DNI)
                            xlHoja1.Cells(lnPosYInicial20, 10) = IIf(Trim(rs!P2RUC) = "", "-", rs!P2RUC)
                            
                            lsPropiedadDirecta = ""
                            lsPropiedadIndirecta = ""
                            
                            If nEncontrado = 0 Or nEncontradoVin = 0 Then
                                If Trim(lsPersCodEmpresa) <> Trim(rs!cPersCod) Then
                                     H = H + 1
                                      ReDim Preserve sMatrixEmpresas2(1 To 2, 1 To H)
                                      sMatrixEmpresas2(1, H) = Trim(rs!cPersCod)
                                      sMatrixEmpresas2(2, H) = lnOrden 'H
                                End If
                            lsPersCodEmpresa = Trim(rs!cPersCod)
                            End If
                            ReDim Preserve sMatrixEmpresas3(1 To 3, 1 To lnPosYInicial20)
                            sMatrixEmpresas3(1, lnPosYInicial20) = Trim(rs!cPersCod)
                            sMatrixEmpresas3(2, lnPosYInicial20) = lnPosYInicial20
                            sMatrixEmpresas3(3, lnPosYInicial20) = Trim(rs!cPersCodOtro)
                            
                            xlHoja1.Cells(lnPosYInicial20, 13) = rs!cGestion
                            xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 1), xlHoja1.Cells(lnPosYInicial20, 16)).Borders.LineStyle = 1
                            xlHoja1.Cells(lnPosYInicial20, 14) = rs!SalCreEmpre
                            xlHoja1.Cells(lnPosYInicial20, 15) = rs!nSaldo
                            xlHoja1.Cells(lnPosYInicial20, 16) = rs!cAgeDescripcion
                            'xlHoja1.Cells(lnPosYInicial20, 17) = i
                       End If
                    End If
                rs.MoveNext

            Wend
    End If
    rs.Close
    
    'If ix > 0 Then
    For P = 1 To lnPosYInicial20
       lsPropiedadDirecta = ""
       lsPropiedadIndirecta = ""
        For k = 1 To ipdi
            If Trim(sMatrixPropiedadDirecIn(3, k)) = sMatrixEmpresas3(3, P) Then
                For L = 1 To H
                    If Trim(sMatrixEmpresas2(1, L)) = Trim(sMatrixPropiedadDirecIn(2, k)) Then
                        nPosEmpresa = sMatrixEmpresas2(2, L)
                        Exit For
                    End If
                Next L
                If sMatrixPropiedadDirecIn(4, k) = 1 Then
                lsPropiedadDirecta = lsPropiedadDirecta & " " & sMatrixPropiedadDirecIn(1, k) & "% A(" & nPosEmpresa & ")"
                Else
                lsPropiedadIndirecta = lsPropiedadIndirecta & " " & sMatrixPropiedadDirecIn(1, k) & "% A(" & nPosEmpresa & ")"
                End If
                xlHoja1.Cells(sMatrixEmpresas3(2, P), 11) = lsPropiedadDirecta
                xlHoja1.Cells(sMatrixEmpresas3(2, P), 12) = lsPropiedadIndirecta
            End If
        Next k
    Next P
    'End If
    
    
    Set rs = Nothing
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

End Sub
'Private Sub XXXGeneraRep19_GrupoEconomicoEmpresa(ByVal pdfecha As Date)
'    Dim lsConec As String
'    Dim lsSql As String
'    Dim lrReg As New ADODB.Recordset
'    Dim I As Integer, lnIIni As Integer
'    Dim lnContador As Long
'    Dim lrRegCab As New ADODB.Recordset
'
'    oCon.AbreConexion
'
'   ExcelAddHoja "Rep_19", xlLibro, xlHoja1
'   xlHoja1.PageSetup.Orientation = xlLandscape
'   xlHoja1.PageSetup.CenterHorizontally = True
'   xlHoja1.PageSetup.Zoom = 75
'
'   xlHoja1.Cells(1, 1) = "REPORTE 19"
'   xlHoja1.Cells(2, 1) = "INFORME SOBRE EL GRUPO ECONOMICO DE LA EMPRESA"
'   xlHoja1.Cells(4, 1) = gsNomCmac
'   xlHoja1.Cells(5, 1) = "AL " & Format(pdfecha, "dd/mm/yyyy")
'
'   xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 9)).Font.Bold = True
'   xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 9)).Merge True
'   xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 9)).HorizontalAlignment = xlCenter
'
'   xlHoja1.Range("A1:A150").ColumnWidth = 5
'   xlHoja1.Range("B1:B150").ColumnWidth = 40
'   xlHoja1.Range("C1:C150").ColumnWidth = 20
'   xlHoja1.Range("D1:D150").ColumnWidth = 20
'   xlHoja1.Range("E1:E150").ColumnWidth = 30
'   xlHoja1.Range("F1:F150").ColumnWidth = 10
'   xlHoja1.Range("G1:G150").ColumnWidth = 15
'   xlHoja1.Range("H1:H150").ColumnWidth = 15
'   xlHoja1.Range("I1:I150").ColumnWidth = 20
'
'   I = 7
'
'   xlHoja1.Cells(I, 1) = "Nro"
'   xlHoja1.Cells(I, 2) = "Nombre, razon o "
'   xlHoja1.Cells(I + 1, 2) = "Denominacion"
'   xlHoja1.Cells(I + 2, 2) = "Social"
'   xlHoja1.Cells(I, 3) = "Tipo de"
'   xlHoja1.Cells(I + 1, 3) = "Persona"
'   xlHoja1.Cells(I, 4) = "CIIU"
'   xlHoja1.Cells(I, 5) = "Domicilio"
'   xlHoja1.Cells(I, 6) = "Tipo de"
'   xlHoja1.Cells(I + 1, 6) = "Documento"
'   xlHoja1.Cells(I, 7) = "Numero del"
'   xlHoja1.Cells(I + 1, 7) = "documento"
'   xlHoja1.Cells(I, 8) = "RUC"
'   xlHoja1.Cells(I, 9) = "Persona juridica sobre la"
'   xlHoja1.Cells(I + 1, 9) = "cual se ejerce control"
'   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 2, 9)).HorizontalAlignment = xlCenter
'   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 2, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
'   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 2, 9)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
'   I = I + 3
'
'    lsSql = " Select ger.cPersCodRel, IsNull(P.cPersCIIU,'') cPersCIIU, P.cPersNombre, P.nPersPersoneria,  " _
'          & " IsNull((Select Top 1 cPersIDnro + '*' + Case cPersIDTPO When 1 Then 'DNI' When 4 Then 'CE' When 11 Then 'PAS' End  From PersId PID Where PID.cPersCod = P.cPersCod And cPersIDTPO In (1,4,11) order by cPersIDTPO),'*') Doc" _
'          & " , ISNUll(PID.cPersIDnro,'') RUC, P.cPersDireccDomicilio, ger.nPrdPersRelac," _
'          & " CO.cConsDescripcion" _
'          & " From PersGrupoEcon ge" _
'          & " inner Join PersGERelacion ger on ge.cGECod = ger.cGECod" _
'          & " Inner Join Persona p on p.cPersCod = ger.cPersCodRel" _
'          & " Inner Join Constante CO On CO.nConsCod = 4028 And CO.nConsValor = ger.nPrdPersRelac" _
'          & " Left Join PersID PID On PID.cPersCod = P.cPersCod And cPersIDTpo = 2" _
'          & " Where ge.cGECod ='000001' "
'    lrReg.CursorLocation = adUseClient
'    Set lrReg = oCon.CargaRecordSet(lsSql)
'    Set lrReg.ActiveConnection = Nothing
'    lnContador = 1
'    lnIIni = I
'    Do While Not lrReg.EOF
'        xlHoja1.Cells(I, 1) = lnContador
'        xlHoja1.Cells(I, 2) = lrReg!cPersNombre
'        xlHoja1.Cells(I, 3) = IIf(lrReg!nPersPersoneria = PersPersoneria.gPersonaNat, "NAT", "JUR")
'        xlHoja1.Cells(I, 4) = "'" & lrReg!cPersCIIU
'        xlHoja1.Cells(I, 5) = lrReg!cPersDireccDomicilio
'        xlHoja1.Cells(I, 6) = Mid(lrReg!Doc, InStr(1, lrReg!Doc, "*") + 1)
'        xlHoja1.Cells(I, 7) = "'" & Left(lrReg!Doc, InStr(1, lrReg!Doc, "*") - 1)
'        xlHoja1.Cells(I, 8) = "'" & lrReg!Ruc
'        xlHoja1.Cells(I, 9) = Trim(Str(lnContador)) & "-" & Mid(lrReg!cConsDescripcion, InStr(1, lrReg!cConsDescripcion, "[") + 1, 1)
'        lnContador = lnContador + 1
'        I = I + 1
'        lrReg.MoveNext
'    Loop
'    lrReg.Close
'
'    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
'    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 9)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
'
'' **************************************************************************************
''  ANEXO 19-A - Informacion sobre personas Juridicas Integrantes del Grupo Economico
''***************************************************************************************
'
'   'Para cada Persona Jurid integrante del Grupo Economico
'
'    lsSql = " Select ger.cPersCodRel, P.cPersNombre, P.nPersPersoneria, IsNull((Select Top 1 cPersIDnro + '*' + Case cPersIDTPO When 1 Then 'DNI' When 4 Then 'CE' When 11 Then 'PAS' End  From PersId PID Where PID.cPersCod = P.cPersCod And cPersIDTPO In (1,4,11) order by cPersIDTPO),'*') Doc " & _
'            " , ISNUll(PID.cPersIDnro,'') RUC, P.cPersDireccDomicilio, ger.nPrdPersRelac " & _
'            " From PersGrupoEcon ge " & _
'            " Inner join PersGERelacion ger on ge.cGECod = ger.cGECod " & _
'            " Inner join Persona p on p.cPersCod = ger.cPersCodRel" & _
'            " Left Join PersID PID On PID.cPersCod = P.cPersCod And cPersIDTpo = 2" & _
'            " Where ge.cGECod = '000001' and P.nPersPersoneria <> " & PersPersoneria.gPersonaNat
'
'    lrRegCab.CursorLocation = adUseClient
'    Set lrRegCab = oCon.CargaRecordSet(lsSql)
'    Set lrRegCab.ActiveConnection = Nothing
'    lnContador = 1
'    Do While Not lrRegCab.EOF
'        ' Adiciona una hoja
'        ExcelAddHoja "Rep_19B-" & Str(lnContador), xlLibro, xlHoja1
'        xlHoja1.PageSetup.Orientation = xlLandscape
'        xlHoja1.PageSetup.CenterHorizontally = True
'        xlHoja1.PageSetup.Zoom = 75
'
'       xlHoja1.Cells(1, 1) = "REPORTE 19-A"
'       xlHoja1.Cells(2, 1) = "INFORME SOBRE PERSONAS JURIDICAS INTEGRANTES DEL GRUPO ECONOMICO"
'       xlHoja1.Cells(4, 1) = gsNomCmac
'       xlHoja1.Cells(5, 1) = "AL " & Format(pdfecha, "dd/mm/yyyy")
'
'       xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 9)).Font.Bold = True
'       xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 9)).Merge True
'       xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 9)).HorizontalAlignment = xlCenter
'
'       xlHoja1.Range("A1:A150").ColumnWidth = 40
'       xlHoja1.Range("B1:B150").ColumnWidth = 15
'       xlHoja1.Range("C1:C150").ColumnWidth = 15
'       xlHoja1.Range("D1:D150").ColumnWidth = 10
'       xlHoja1.Range("E1:E150").ColumnWidth = 15
'       xlHoja1.Range("F1:F150").ColumnWidth = 15
'       xlHoja1.Range("G1:G150").ColumnWidth = 15
'       xlHoja1.Range("H1:H150").ColumnWidth = 15
'       xlHoja1.Range("I1:I150").ColumnWidth = 15
'       xlHoja1.Range("I1:I150").ColumnWidth = 15
'
'       I = 7
'
'       xlHoja1.Cells(I, 1) = "Razon o Denominacion Social"
'       xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).Merge True
'       xlHoja1.Cells(I, 4) = lrRegCab!cPersNombre
'       xlHoja1.Range(xlHoja1.Cells(I, 4), xlHoja1.Cells(I, 10)).Merge True
'       xlHoja1.Cells(I + 1, 1) = "Codigo SBS"
'       xlHoja1.Range(xlHoja1.Cells(I + 1, 1), xlHoja1.Cells(I + 1, 3)).Merge True
'       xlHoja1.Cells(I + 1, 4) = ""
'       xlHoja1.Range(xlHoja1.Cells(I + 1, 4), xlHoja1.Cells(I + 1, 10)).Merge True
'       xlHoja1.Cells(I + 2, 1) = "R.U.C."
'       xlHoja1.Range(xlHoja1.Cells(I + 2, 1), xlHoja1.Cells(I + 2, 3)).Merge True
'       xlHoja1.Cells(I + 2, 4) = lrRegCab!Ruc
'       xlHoja1.Range(xlHoja1.Cells(I + 2, 4), xlHoja1.Cells(I + 2, 10)).Merge True
'       xlHoja1.Cells(I + 3, 1) = "Direccion"
'       xlHoja1.Range(xlHoja1.Cells(I + 3, 1), xlHoja1.Cells(I + 3, 3)).Merge True
'       xlHoja1.Cells(I + 3, 4) = lrRegCab!cPersDireccDomicilio
'       xlHoja1.Range(xlHoja1.Cells(I + 3, 4), xlHoja1.Cells(I + 3, 10)).Merge True
'       xlHoja1.Cells(I + 4, 1) = "Representante Legal"
'       xlHoja1.Range(xlHoja1.Cells(I + 4, 1), xlHoja1.Cells(I + 4, 3)).Merge True
'       xlHoja1.Cells(I + 4, 4) = ""
'       xlHoja1.Range(xlHoja1.Cells(I + 4, 4), xlHoja1.Cells(I + 4, 10)).Merge True
'       xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 4, 10)).Cells.Borders.LineStyle = xlOutside
'       xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 4, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
'       xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 4, 10)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
'
'       I = I + 6
'       xlHoja1.Cells(I, 1) = "Nombre"
'       xlHoja1.Cells(I, 2) = "Cod"
'       xlHoja1.Cells(I + 1, 2) = "SBS"
'       xlHoja1.Cells(I, 3) = "Tipo de"
'       xlHoja1.Cells(I + 1, 3) = "Persona"
'       xlHoja1.Cells(I, 4) = "Tipo de"
'       xlHoja1.Cells(I + 1, 4) = "Documento"
'       xlHoja1.Cells(I, 5) = "Numero de"
'       xlHoja1.Cells(I + 1, 5) = "Documento"
'       xlHoja1.Cells(I, 6) = "RUC"
'       xlHoja1.Cells(I, 7) = "Residencia"
'       xlHoja1.Cells(I, 8) = "Accionista"
'       xlHoja1.Cells(I, 9) = "Cargo"
'       xlHoja1.Cells(I, 10) = "Otro"
'       xlHoja1.Cells(I + 1, 10) = "Cargo"
'
'       xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 2, 10)).HorizontalAlignment = xlCenter
'       'xlHoja1.Range(xlHoja1.Cells(i, 1), xlHoja1.Cells(i + 2, 10)).Cells.Borders.LineStyle = xlOutside
'       xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 2, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
'       xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 2, 10)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
'       I = I + 3
'
'
'
'       lsSql = " Select ger.cPersCodRel, P.cPersNombre cNomPers, P.nPersPersoneria," _
'             & " IsNull((Select Top 1 cPersIDnro + '*' + Case cPersIDTPO When 1 Then 'DNI' When 4 Then 'CE' When 11 Then 'PAS' End  From PersId PID Where PID.cPersCod = P.cPersCod And cPersIDTPO In (1,4,11) order by cPersIDTPO),'*') Doc," _
'             & " ISNUll(PID.cPersIDnro,'') RUC, ger.nPrdPersRelac, gePV.cPersCodRel," _
'             & " PerVinc.cPersNombre as cNomPersVinc, PerVinc.nPersPersoneria," _
'             & " IsNUll((Select Top 1 cPersIDnro + '*' + Case cPersIDTPO When 1 Then 'DNI' When 4 Then 'CE' When 11 Then 'PAS' End  From PersId PID Where PID.cPersCod = PerVinc.cPersCod And cPersIDTPO In (1,4,11) order by cPersIDTPO),'*') DocVin," _
'             & " ISNUll(PID.cPersIDnro,'') RUCV, PerVinc.cPersDireccDomicilio as cDirPersVinc,  gePV.nCargo, gePV.nParticip" _
'             & " From PersGrupoEcon ge" _
'             & " inner join PersGERelacion ger on ge.cGECod = ger.cGECod" _
'             & " Inner join Persona p on p.cPersCod = ger.cPersCodRel" _
'             & " Inner join PersGEPersVinc gePV on gePV.cPersCodRel = ger.cPersCodRel" _
'             & " Inner join Persona PerVinc on PerVinc.cPersCod = gePV.cPersCodVinc" _
'             & " Left Join PersID PID On PID.cPersCod = P.cPersCod And PID.cPersIDTpo = 2" _
'             & " Left Join PersID PIDV On PIDV.cPersCod = P.cPersCod And PIDV.cPersIDTpo = 2" _
'             & " Where ge.cGECod ='000001' and p.nPersPersoneria <> '1' and gePV.cPersCodRel ='" & lrRegCab!cPersCodRel & "'"
'
'        lrReg.CursorLocation = adUseClient
'        Set lrReg = oCon.CargaRecordSet(lsSql)
'        Set lrReg.ActiveConnection = Nothing
'        lnIIni = I
'        Do While Not lrReg.EOF
'            xlHoja1.Cells(I, 1) = lrReg!cNomPersVinc
'            xlHoja1.Cells(I, 2) = ""
'            xlHoja1.Cells(I, 3) = IIf(lrReg!nPersPersoneria = PersPersoneria.gPersonaNat, "NAT", "JUR")
'            xlHoja1.Cells(I, 4) = Mid(lrReg!Doc, InStr(1, lrReg!Doc, "*") + 1)
'            xlHoja1.Cells(I, 5) = Left(lrReg!Doc, InStr(1, lrReg!Doc, "*") - 1)
'            xlHoja1.Cells(I, 6) = IIf(lrReg!nPersPersoneria = PersPersoneria.gPersonaNat, "", lrReg!Ruc)
'            xlHoja1.Cells(I, 7) = lrReg!cDirPersVinc
'            xlHoja1.Cells(I, 8) = lrReg!nParticip
'            xlHoja1.Cells(I, 9) = lrReg!nCargo
'            xlHoja1.Cells(I, 10) = ""
'            I = I + 1
'            lrReg.MoveNext
'        Loop
'        lrReg.Close
'
'        xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 10)).Cells.Borders.LineStyle = xlOutside
'        xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
'        xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 10)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
'
'        lrRegCab.MoveNext
'        lnContador = lnContador + 1
'   Loop
'   lrRegCab.Close
'
'   MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "Aviso"
'End Sub

''-**  ANEXO 20 - Clientes que Representan Riesgo Unico
'Private Sub xxxGeneraRep20_ClientesRepresRiesgoUnico(ByVal pdfecha As Date)
'    Dim lsConec As String
'    Dim lsSql As String
'    Dim lrReg As New ADODB.Recordset
'    Dim I As Integer, lnIIni As Integer
'    Dim lnContador As Long
'    Dim lrRegCab As New ADODB.Recordset
'
'    oCon.AbreConexion
'
'   'Para cada Cliente que representa Riesgo Unico
'
'    'lsSQL = " Select Distinct(ger.cCodPersRel) cCodPersRel, P.cNomPers, P.cTipPers," & _
'            " P.cTidoci, P.cNudoci, P.cTidotr, P.cNudotr, P.cDirPers, ger.cRelacion  " & _
'            " From RGrupoEcon ge inner join rGERelacion ger on ge.cCodGE = ger.cCodGe " & _
'            " Inner join dbPersona..Persona p on p.cCodPers = ger.cCodPersRel " & _
'            " Where ge.cCodGE <>'000001' "
'    lsSql = " Select IsNull(P.cPersCodSBS,'') cPersCodSBS, IsNull(P.cPersCIIU,'') cPersCIIU, ger.cPersCodRel, P.cPersNombre, P.nPersPersoneria,  " _
'          & " (Select Top 1 cPersIDnro + '*' + Case cPersIDTPO When 1 Then 'DNI' When 4 Then 'CE' When 11 Then 'PAS' End  From PersId PID Where PID.cPersCod = P.cPersCod And cPersIDTPO In (1,4,11) order by cPersIDTPO) Doc" _
'          & " , ISNUll(PID.cPersIDnro,'') RUC, P.cPersDireccDomicilio, ger.nPrdPersRelac" _
'          & " From PersGrupoEcon ge" _
'          & " inner join PersGERelacion ger on ge.cGECod = ger.cGECod" _
'          & " Inner join Persona p on p.cPersCod = ger.cPersCodRel" _
'          & " Left Join PersID PID On PID.cPersCod = P.cPersCod And cPersIDTpo = 2" _
'          & " Where ge.cGECod <>'000001' "
'
'    lrRegCab.CursorLocation = adUseClient
'    Set lrRegCab = oCon.CargaRecordSet(lsSql)
'    Set lrRegCab.ActiveConnection = Nothing
'    lnContador = 1
'    If lrRegCab.BOF And lrRegCab.EOF Then
'        lrRegCab.Close
'        MsgBox "No existen datos para el Reporte", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    Do While Not lrRegCab.EOF
'        ' Adiciona una hoja
'        ExcelAddHoja "Rep_20-" & Str(lnContador), xlLibro, xlHoja1
'        xlHoja1.PageSetup.Orientation = xlLandscape
'        xlHoja1.PageSetup.CenterHorizontally = True
'        xlHoja1.PageSetup.Zoom = 75
'
'       xlHoja1.Cells(1, 1) = "REPORTE 20"
'       xlHoja1.Cells(2, 1) = "INFORMACION DE CLIENTES QUE REPRESENTAN RIESGO UNICO"
'       xlHoja1.Cells(4, 1) = gsNomCmac
'       xlHoja1.Cells(5, 1) = "AL " & Format(pdfecha, "dd/mm/yyyy")
'
'       xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 10)).Font.Bold = True
'       xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 10)).Merge True
'       xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 10)).HorizontalAlignment = xlCenter
'
'       xlHoja1.Range("A1:A150").ColumnWidth = 40
'       xlHoja1.Range("B1:B150").ColumnWidth = 15
'       xlHoja1.Range("C1:C150").ColumnWidth = 15
'       xlHoja1.Range("D1:D150").ColumnWidth = 10
'       xlHoja1.Range("E1:E150").ColumnWidth = 15
'       xlHoja1.Range("F1:F150").ColumnWidth = 45
'       xlHoja1.Range("G1:G150").ColumnWidth = 15
'       xlHoja1.Range("H1:H150").ColumnWidth = 15
'       xlHoja1.Range("I1:I150").ColumnWidth = 15
'       xlHoja1.Range("I1:I150").ColumnWidth = 15
'
'       I = 7
'       xlHoja1.Cells(I, 1) = "1. INFORMACION DEL CLIENTE"
'
'       xlHoja1.Cells(I, 1) = "Razon o Denominacion Social"
'       xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).Merge True
'       xlHoja1.Cells(I, 4) = lrRegCab!cPersNombre
'       xlHoja1.Range(xlHoja1.Cells(I, 4), xlHoja1.Cells(I, 10)).Merge True
'       xlHoja1.Cells(I + 1, 1) = "Codigo SBS"
'       xlHoja1.Range(xlHoja1.Cells(I + 1, 1), xlHoja1.Cells(I + 1, 3)).Merge True
'       xlHoja1.Cells(I + 1, 4) = ""
'       xlHoja1.Range(xlHoja1.Cells(I + 1, 4), xlHoja1.Cells(I + 1, 10)).Merge True
'       xlHoja1.Cells(I + 2, 1) = "R.U.C."
'       xlHoja1.Range(xlHoja1.Cells(I + 2, 1), xlHoja1.Cells(I + 2, 3)).Merge True
'       xlHoja1.Cells(I + 2, 4) = lrRegCab!Ruc
'       xlHoja1.Range(xlHoja1.Cells(I + 2, 4), xlHoja1.Cells(I + 2, 10)).Merge True
'       xlHoja1.Cells(I + 3, 1) = "Direccion"
'       xlHoja1.Range(xlHoja1.Cells(I + 3, 1), xlHoja1.Cells(I + 3, 3)).Merge True
'       xlHoja1.Cells(I + 3, 4) = lrRegCab!cPersDireccDomicilio
'       xlHoja1.Range(xlHoja1.Cells(I + 3, 4), xlHoja1.Cells(I + 3, 10)).Merge True
'       xlHoja1.Cells(I + 4, 1) = "Representante Legal"
'       xlHoja1.Range(xlHoja1.Cells(I + 4, 1), xlHoja1.Cells(I + 4, 3)).Merge True
'       xlHoja1.Cells(I + 4, 4) = ""
'       xlHoja1.Range(xlHoja1.Cells(I + 4, 4), xlHoja1.Cells(I + 4, 10)).Merge True
'       xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 4, 10)).Cells.Borders.LineStyle = xlOutside
'       xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 4, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
'       xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 4, 10)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
'
'       ' Si es una persona Juridica se aplica lo siguiente
'       If lrRegCab!nPersPersoneria <> PersPersoneria.gPersonaNat Then
'
'            I = I + 6
'            xlHoja1.Cells(I, 1) = "Nombre"
'            xlHoja1.Cells(I, 2) = "Cod"
'            xlHoja1.Cells(I + 1, 2) = "SBS"
'            xlHoja1.Cells(I, 3) = "Tipo de"
'            xlHoja1.Cells(I + 1, 3) = "Persona"
'            xlHoja1.Cells(I, 4) = "Tipo de"
'            xlHoja1.Cells(I + 1, 4) = "Documento"
'            xlHoja1.Cells(I, 5) = "Numero de"
'            xlHoja1.Cells(I + 1, 5) = "Documento"
'            xlHoja1.Cells(I, 6) = "RUC"
'            xlHoja1.Cells(I, 7) = "Residencia"
'            xlHoja1.Cells(I, 8) = "Accionista"
'            xlHoja1.Cells(I, 9) = "Cargo"
'            xlHoja1.Cells(I, 10) = "Otro"
'            xlHoja1.Cells(I + 1, 10) = "Cargo"
'
'            xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 2, 10)).HorizontalAlignment = xlCenter
'            'xlHoja1.Range(xlHoja1.Cells(i, 1), xlHoja1.Cells(i + 2, 10)).Cells.Borders.LineStyle = xlOutside
'            xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 2, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
'            xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 2, 10)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
'            I = I + 3
'
''             lsSQL = " Select ger.cCodPersRel, P.cNomPers, P.cTipPers, " & _
''                     " P.cTidoci, P.cNudoci, P.cTidotr, P.cNudotr, ger.cRelacion, " & _
''                     " gePV.cCodPersRel, PerVinc.cNomPers as cNomPersVinc, PerVinc.cTipPers, " & _
''                     " PerVinc.cTidoci, PerVinc.cNudoci, PerVinc.cTidotr, PerVinc.cNudotr, PerVinc.cDirPers as cDirPersVinc, " & _
''                     " gePV.cRelacion, gePV.cCargo, gePV.nParticip " & _
''                     " From RGrupoecon ge inner join rGERelacion ger on ge.cCodGE = ger.cCodGe " & _
''                     " Inner join dbPersona..Persona p on p.cCodPers = ger.cCodPersRel " & _
''                     " Inner join rGEPersVinc gePV on gePV.cCodPersRel = ger.cCodPersRel " & _
''                     " Inner join dbPersona..Persona PerVinc on PerVinc.cCodPers = gePV.cCodPersVinc " & _
''                     " Where ge.cCodGE <>'000001' and gePV.cCodPersRel ='" & lrRegCab!cCodPersRel & "' "
''
''             lrReg.CursorLocation = adUseClient
''             Set lrReg = oCon.CargaRecordSet(lsSQL)
''             Set lrReg.ActiveConnection = Nothing
''             lnIIni = I
''             Do While Not lrReg.EOF
''                 xlHoja1.Cells(I, 1) = lrReg!cNomPersVinc
''                 xlHoja1.Cells(I, 2) = ""
''                 xlHoja1.Cells(I, 3) = IIf(lrReg!cTipPers = "1", "NAT", "JUR")
''                 xlHoja1.Cells(I, 4) = IIf(lrReg!cTipPers = "1", "DNI", "0")
''                 xlHoja1.Cells(I, 5) = IIf(lrReg!cTipPers = "1", lrReg!cNudoci, "")
''                 xlHoja1.Cells(I, 6) = IIf(lrReg!cTipPers = "1", "", lrReg!cNudoTr)
''                 xlHoja1.Cells(I, 7) = lrReg!cDirPersVinc
''                 xlHoja1.Cells(I, 8) = lrReg!nParticip
''                 xlHoja1.Cells(I, 9) = lrReg!cCargo
''                 xlHoja1.Cells(I, 10) = ""
''                 I = I + 1
''                 lrReg.MoveNext
''             Loop
''             lrReg.Close
'                'lsSQL = " Select ger.cPersCodRel, P.cPersNombre cNomPers, P.nPersPersoneria," _
'                      & " (Select Top 1 cPersIDnro + '*' + Case cPersIDTPO When 1 Then 'DNI' When 4 Then 'CE' When 11 Then 'PAS' End  From PersId PID Where PID.cPersCod = P.cPersCod And cPersIDTPO In (1,4,11) order by cPersIDTPO) Doc," _
'                      & " ISNUll(PID.cPersIDnro,'') RUC, ger.nPrdPersRelac, gePV.cPersCodRel," _
'                      & " PerVinc.cPersNombre as cNomPersVinc, PerVinc.nPersPersoneria," _
'                      & " (Select Top 1 cPersIDnro + '*' + Case cPersIDTPO When 1 Then 'DNI' When 4 Then 'CE' When 11 Then 'PAS' End  From PersId PID Where PID.cPersCod = PerVinc.cPersCod And cPersIDTPO In (1,4,11) order by cPersIDTPO) DocVin," _
'                      & " ISNUll(PID.cPersIDnro,'') RUCV, PerVinc.cPersDireccDomicilio as cDirPersVinc,  gePV.nCargo, gePV.nParticip" _
'                      & " From PersGrupoEcon ge" _
'                      & " inner join PersGERelacion ger on ge.cGECod = ger.cGECod" _
'                      & " Inner join Persona p on p.cPersCod = ger.cPersCodRel" _
'                      & " Inner join PersGEPersVinc gePV on gePV.cPersCodRel = ger.cPersCodRel" _
'                      & " Inner join Persona PerVinc on PerVinc.cPersCod = gePV.cPersCodVinc" _
'                      & " Left Join PersID PID On PID.cPersCod = P.cPersCod And PID.cPersIDTpo = 2" _
'                      & " Left Join PersID PIDV On PIDV.cPersCod = P.cPersCod And PIDV.cPersIDTpo = 2" _
'                      & " Where ge.cGECod ='000001' and p.nPersPersoneria <> '1' and gePV.cPersCodRel ='" & lrRegCab!cPersCodRel & "'"
'
'                lsSql = " Select IsNull(PerVinc.cPersCodSBS,'') cPersCodSBS, IsNull(PerVinc.cPersCIIU,'') cPersCIIU, PIDV.cPersCod PV , ger.cPersCodRel, P.cPersNombre cNomPers, P.nPersPersoneria," _
'                      & " (Select Top 1 cPersIDnro + '*' + Case cPersIDTPO When 1 Then 'DNI' When 4 Then 'CE' When 11 Then 'PAS' End  From PersId PID Where PID.cPersCod = P.cPersCod And cPersIDTPO In (1,4,11) order by cPersIDTPO) Doc," _
'                      & " ISNUll(PID.cPersIDnro,'') RUC, ger.nPrdPersRelac, gePV.cPersCodRel," _
'                      & " PerVinc.cPersNombre as cNomPersVinc, PerVinc.nPersPersoneria," _
'                      & " IsNull((Select Top 1 cPersIDnro + '*' + Case cPersIDTPO When 1 Then 'DNI' When 4 Then 'CE' When 11 Then 'PAS' End  From PersId PID Where PID.cPersCod = PerVinc.cPersCod And cPersIDTPO In (1,4,11) order by cPersIDTPO),'*') DocVin," _
'                      & " ISNUll(PID.cPersIDnro,'') RUCV, PerVinc.cPersDireccDomicilio as cDirPersVinc,  gePV.nCargo, gePV.nParticip" _
'                      & " From PersGrupoEcon ge" _
'                      & " inner join PersGERelacion ger on ge.cGECod = ger.cGECod" _
'                      & " Inner join Persona p on p.cPersCod = ger.cPersCodRel" _
'                      & " Inner join PersGEPersVinc gePV on gePV.cPersCodRel = ger.cPersCodRel" _
'                      & " Inner join Persona PerVinc on PerVinc.cPersCod = gePV.cPersCodVinc" _
'                      & " Left Join PersID PID On PID.cPersCod = P.cPersCod And PID.cPersIDTpo = 2" _
'                      & " Left Join PersID PIDV On PIDV.cPersCod = PerVinc.cPersCod And PIDV.cPersIDTpo = 2" _
'                      & " Where ge.cGECod <> '000001' and gePV.cPersCodRel ='" & lrRegCab!cPersCodRel & "'"
'
'                 lrReg.CursorLocation = adUseClient
'                 Set lrReg = oCon.CargaRecordSet(lsSql)
'                 Set lrReg.ActiveConnection = Nothing
'                 lnIIni = I
'                 Do While Not lrReg.EOF
'                     xlHoja1.Cells(I, 1) = lrReg!cNomPersVinc
'                     xlHoja1.Cells(I, 2) = lrReg!cPersCodSBS
'                     xlHoja1.Cells(I, 3) = IIf(lrReg!nPersPersoneria = PersPersoneria.gPersonaNat, "NAT", "JUR")
'                     xlHoja1.Cells(I, 4) = IIf(InStr(1, lrReg!DocVin, "*") <> 0, Mid(lrReg!DocVin, InStr(1, lrReg!DocVin, "*") + 1), "")
'                     xlHoja1.Cells(I, 5) = "'" & IIf(InStr(1, lrReg!DocVin, "*") <> 0, Left(lrReg!DocVin, InStr(1, lrReg!DocVin, "*") - 1), "")
'                     xlHoja1.Cells(I, 6) = "'" & lrReg!RUCv
'                     xlHoja1.Cells(I, 7) = lrReg!cDirPersVinc
'                     xlHoja1.Cells(I, 8) = lrReg!nParticip
'                     xlHoja1.Cells(I, 9) = lrReg!nCargo
'                     xlHoja1.Cells(I, 10) = ""
'                     I = I + 1
'                     lrReg.MoveNext
'                 Loop
'                 lrReg.Close
'             xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 10)).Cells.Borders.LineStyle = xlOutside
'             xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
'             xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 10)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
'
'             xlHoja1.Cells.Select
'             xlHoja1.Cells.EntireColumn.AutoFit
'
'       End If
'       lrRegCab.MoveNext
'       lnContador = lnContador + 1
'   Loop
'   lrRegCab.Close
'
''*******************************************************************************
''** Reporte 20-A INFORMACION DE LAS PERSONAS QUE REPRESNTAN RIESGO UNICO CLIENTES
''*******************************************************************************
'
'    ' Adiciona una hoja
'   ExcelAddHoja "Rep_20-A", xlLibro, xlHoja1
'   xlHoja1.PageSetup.Orientation = xlLandscape
'   xlHoja1.PageSetup.CenterHorizontally = True
'   xlHoja1.PageSetup.Zoom = 75
'
'   xlHoja1.Cells(1, 1) = "REPORTE 20-A"
'   xlHoja1.Cells(2, 1) = "INFORMACION DE LAS PERSONAS QUE REPRESENTAN RIESGO UNICO CLIENTES"
'   xlHoja1.Cells(4, 1) = gsNomCmac
'   xlHoja1.Cells(5, 1) = "AL " & Format(pdfecha, "dd/mm/yyyy")
'
'   xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 13)).Font.Bold = True
'   xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 13)).Merge True
'   xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 13)).HorizontalAlignment = xlCenter
'
'   xlHoja1.Range("A1:A150").ColumnWidth = 40
'   xlHoja1.Range("B1:B150").ColumnWidth = 15
'   xlHoja1.Range("C1:C150").ColumnWidth = 15
'   xlHoja1.Range("D1:D150").ColumnWidth = 10
'   xlHoja1.Range("E1:E150").ColumnWidth = 15
'   xlHoja1.Range("F1:F150").ColumnWidth = 15
'   xlHoja1.Range("G1:G150").ColumnWidth = 15
'   xlHoja1.Range("H1:H150").ColumnWidth = 15
'   xlHoja1.Range("I1:I150").ColumnWidth = 15
'   xlHoja1.Range("I1:I150").ColumnWidth = 15
'
'   I = 7
'
'    xlHoja1.Cells(I, 1) = "Nro de"
'    xlHoja1.Cells(I + 1, 1) = "Vinculado"
'    xlHoja1.Cells(I, 2) = "Codigo"
'    xlHoja1.Cells(I + 1, 2) = "SBS del"
'    xlHoja1.Cells(I + 2, 2) = "Cliente"
'    xlHoja1.Cells(I, 3) = "Codigo"
'    xlHoja1.Cells(I + 1, 3) = "SBS"
'    xlHoja1.Cells(I, 4) = "Nombre/Razon/"
'    xlHoja1.Cells(I + 1, 4) = "Denominacion Social"
'    xlHoja1.Cells(I, 5) = "CIIU"
'    xlHoja1.Cells(I, 6) = "Domicilio"
'    xlHoja1.Cells(I, 7) = "Tipo de"
'    xlHoja1.Cells(I + 1, 7) = "Persona"
'    xlHoja1.Cells(I, 8) = "Tipo de"
'    xlHoja1.Cells(I + 1, 8) = "Documento"
'    xlHoja1.Cells(I, 9) = "Numero de"
'    xlHoja1.Cells(I + 1, 9) = "Documento"
'    xlHoja1.Cells(I, 10) = "RUC"
'    xlHoja1.Cells(I, 11) = "Descripcion de la vinculacion"
'    xlHoja1.Cells(I + 1, 11) = "Propiedad"
'    xlHoja1.Cells(I + 2, 11) = "Directa"
'    xlHoja1.Cells(I + 1, 12) = "Propiedad"
'    xlHoja1.Cells(I + 2, 12) = "Indirecta"
'    xlHoja1.Cells(I + 1, 13) = "Gestion"
'
'
'    xlHoja1.Range(xlHoja1.Cells(I, 11), xlHoja1.Cells(I, 13)).Merge
'    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 2, 13)).HorizontalAlignment = xlCenter
'    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 2, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
'    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 2, 13)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
'    I = I + 3
'
'       lsSql = " Select IsNull(PerVinc.cPersCodSBS,'') cPersCodSBS, IsNull(PerVinc.cPersCIIU,'') cPersCIIU, ger.cPersCodRel, P.cPersNombre cNomPers, P.nPersPersoneria," _
'             & " ISNULL((Select Top 1 cPersIDnro + '*' + Case cPersIDTPO When 1 Then 'DNI' When 4 Then 'CE' When 11 Then 'PAS' End  From PersId PID Where PID.cPersCod = P.cPersCod And cPersIDTPO In (1,4,11) order by cPersIDTPO),'*') Doc," _
'             & " ISNUll(PID.cPersIDnro,'') RUC, ger.nPrdPersRelac, gePV.cPersCodRel," _
'             & " PerVinc.cPersNombre as cNomPersVinc, PerVinc.nPersPersoneria," _
'             & " ISNULL((Select Top 1 cPersIDnro + '*' + Case cPersIDTPO When 1 Then 'DNI' When 4 Then 'CE' When 11 Then 'PAS' End  From PersId PID Where PID.cPersCod = PerVinc.cPersCod And cPersIDTPO In (1,4,11) order by cPersIDTPO),'*') DocVin," _
'             & " ISNUll(PID.cPersIDnro,'') RUCV, PerVinc.cPersDireccDomicilio as cDirPersVinc,  gePV.nCargo, gePV.nParticip" _
'             & " From PersGrupoEcon ge" _
'             & " inner join PersGERelacion ger on ge.cGECod = ger.cGECod" _
'             & " Inner join Persona p on p.cPersCod = ger.cPersCodRel" _
'             & " Inner join PersGEPersVinc gePV on gePV.cPersCodRel = ger.cPersCodRel" _
'             & " Inner join Persona PerVinc on PerVinc.cPersCod = gePV.cPersCodVinc" _
'             & " Left Join PersID PID On PID.cPersCod = P.cPersCod And PID.cPersIDTpo = 2" _
'             & " Left Join PersID PIDV On PIDV.cPersCod = P.cPersCod And PIDV.cPersIDTpo = 2" _
'             & " Where ge.cGECod <> '000001' "
'
'     'lsSQL = " Select ger.cCodPersRel, P.cNomPers, P.cTipPers, " & _
'             " P.cTidoci, P.cNudoci, P.cTidotr, P.cNudotr, ger.cRelacion, " & _
'             " gePV.cCodPersRel, PerVinc.cNomPers as cNomPersVinc, PerVinc.cTipPers, " & _
'             " PerVinc.cTidoci, PerVinc.cNudoci, PerVinc.cTidotr, PerVinc.cNudotr, PerVinc.cDirPers as cDirPersVinc, " & _
'             " gePV.cRelacion, gePV.cCargo, gePV.nParticip " & _
'             " From RGrupoecon ge inner join rGERelacion ger on ge.cCodGE = ger.cCodGe " & _
'             " Inner join dbPersona..Persona p on p.cCodPers = ger.cCodPersRel " & _
'             " Inner join rGEPersVinc gePV on gePV.cCodPersRel = ger.cCodPersRel " & _
'             " Inner join dbPersona..Persona PerVinc on PerVinc.cCodPers = gePV.cCodPersVinc " & _
'             " Where ge.cCodGE <>'000001' "
'
'     lrReg.CursorLocation = adUseClient
'     Set lrReg = oCon.CargaRecordSet(lsSql)
'     Set lrReg.ActiveConnection = Nothing
'     lnIIni = I
'     lnContador = 1
'     Do While Not lrReg.EOF
'         xlHoja1.Cells(I, 1) = lnContador
'         xlHoja1.Cells(I, 2) = lrReg!cPersCodRel ' Cambiar por el CodSBS
'         xlHoja1.Cells(I, 3) = lrReg!cPersCodSBS
'         xlHoja1.Cells(I, 4) = lrReg!cNomPersVinc
'         xlHoja1.Cells(I, 5) = lrReg!cPersCIIU
'         xlHoja1.Cells(I, 6) = lrReg!cDirPersVinc
'         xlHoja1.Cells(I, 7) = IIf(lrReg!nPersPersoneria = PersPersoneria.gPersonaNat, "NAT", "JUR")
'         xlHoja1.Cells(I, 8) = IIf(InStr(1, lrReg!DocVin, "*") <> 0, Mid(lrReg!DocVin, InStr(1, lrReg!DocVin, "*") + 1), "")
'         xlHoja1.Cells(I, 9) = IIf(InStr(1, lrReg!DocVin, "*") <> 0, Left(lrReg!DocVin, InStr(1, lrReg!DocVin, "*") - 1), "")
'         xlHoja1.Cells(I, 10) = lrReg!RUCv
'         xlHoja1.Cells(I, 11) = "" & lrReg!nParticip & "% A"
'         xlHoja1.Cells(I, 12) = "" 'Prop Indirecta
'         xlHoja1.Cells(I, 13) = "" 'Gestion
'         I = I + 1
'         lrReg.MoveNext
'     Loop
'     lrReg.Close
'
'     xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 13)).Cells.Borders.LineStyle = xlOutside
'     xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
'     xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 13)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
'
'     xlHoja1.Cells.Select
'     xlHoja1.Cells.EntireColumn.AutoFit
'
'   MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "Aviso"
'End Sub
'   ALPA 20090929*************************************************************
'Private Sub GeneraRep19_GrupoEconomicoEmpresa(ByVal pdFecha As Date)
'    Dim sSql As String
'    Dim rs As New ADODB.Recordset
'    Dim I As Integer
'    Dim nFila As String
'    Dim nFilaTotal1 As Integer
'    Dim nFilaTotal2 As Integer
'
'    Dim nInicio As Integer
'
'    Dim nPatrEfectivo As Currency
'
'    Dim pdfecha1 As Date
'
'    pdfecha1 = DateAdd("m", 1, "01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Format(txtAnio.Text, "0000"))
'    pdfecha1 = DateAdd("d", -1, pdfecha1)
'
'    'Calculo el nPatrEfectivo
'    nPatrEfectivo = Val(txtPatrimonio.Text) 'Val(oAnx.GetImporteEstadAnexosMax(gdFecSis, "TOTALREP03", "1"))
'
'    Set oCon = New DConecta
'    oCon.AbreConexion
'
'    'Adiciona una hoja
'
'    ExcelAddHoja "Reporte 19", xlLibro, xlHoja1
'
'    xlHoja1.PageSetup.Orientation = xlLandscape
'    xlHoja1.PageSetup.CenterHorizontally = True
'    xlHoja1.PageSetup.Zoom = 60
'
'    xlHoja1.Cells(2, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
'    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).MergeCells = True
'    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).HorizontalAlignment = xlLeft
'
'    xlHoja1.Cells(2, 10) = "REPORTE NRO 19"
'    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).MergeCells = True
'    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).HorizontalAlignment = xlLeft
'
'    xlHoja1.Cells(4, 2) = "INFORME SOBRE EL GRUPO ECONOMICO DE LA EMPRESA"
'    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).MergeCells = True
'    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).HorizontalAlignment = xlCenter
'
'    xlHoja1.Cells(6, 2) = "Empresa que remite la información: " & gsNomCmac
'    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).MergeCells = True
'    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).HorizontalAlignment = xlLeft
'
'    xlHoja1.Cells(7, 2) = "INFORMACION AL " & Format(pdfecha1, "DD MMMM YYYY")
'    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).MergeCells = True
'    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).HorizontalAlignment = xlLeft
'
'    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(7, 10)).Font.Bold = True
'
'    sSql = "Select  PGE.CGECOD, PGR.cPersCodRel, P.cPersNombre, Con.cConsDescripcion, P.cPersCIIU, P.cPersDireccDomicilio,  "
'    sSql = sSql & " PID.cPersIDNro , PGR.cTexto "
'    sSql = sSql & " From PersGrupoEcon PGE "
'    sSql = sSql & " Inner Join PersGERelacion PGR "
'    sSql = sSql & "     On PGE.cGECod=PGR.cGECod "
'    sSql = sSql & " Inner Join Persona P "
'    sSql = sSql & "     On PGR.cPersCodRel = P.cPersCod "
'    sSql = sSql & " Inner Join Constante Con "
'    sSql = sSql & "     on P.nPersPersoneria=Con.nConsValor "
'    sSql = sSql & " Left Join PersID PID "
'    sSql = sSql & "     On P.cPersCod=PID.cPersCod "
'    sSql = sSql & " Where PGE.cCodReporte='19'  And Con.nConsCod=1002 "
'    sSql = sSql & " And isnull(PID.cPersIDTpo, 0) <>99 "
'    sSql = sSql & " Order By PGR.nOrden "
'
'
'    Set rs = oCon.CargaRecordSet(sSql)
'    If rs.BOF Then
'    Else
'        I = 0
'
'        nFila = 9
'
'        xlHoja1.Cells(nFila, 2) = "Nro"
'        xlHoja1.Cells(nFila, 3) = "Nombre, Razón o Denominación Social"
'        xlHoja1.Cells(nFila, 4) = "Tipo de"
'        xlHoja1.Cells(nFila, 5) = "CIIU"
'        xlHoja1.Cells(nFila, 6) = "Domicilio"
'        xlHoja1.Cells(nFila, 7) = "Tipo de"
'        xlHoja1.Cells(nFila, 8) = "Número del"
'        xlHoja1.Cells(nFila, 9) = "RUC"
'        xlHoja1.Cells(nFila, 10) = "Persona Jurídica sobre la cual se"
'
'        nFila = nFila + 1
'        xlHoja1.Cells(nFila, 4) = "Persona"
'        xlHoja1.Cells(nFila, 7) = "Documento"
'        xlHoja1.Cells(nFila, 8) = "Documento"
'        xlHoja1.Cells(nFila, 10) = "ejerce control"
'
'        xlHoja1.Range(xlHoja1.Cells(nFila - 1, 2), xlHoja1.Cells(nFila, 10)).HorizontalAlignment = xlCenter
'        xlHoja1.Range(xlHoja1.Cells(nFila - 1, 2), xlHoja1.Cells(nFila, 10)).Font.Bold = True
'
'        ExcelCuadro xlHoja1, 2, nFila - 1, 10, nFila
'
'        Do While Not rs.EOF
'            nFila = nFila + 1
'            I = I + 1
'            xlHoja1.Cells(nFila, 2) = I
'            xlHoja1.Cells(nFila, 3) = rs!cPersNombre
'            xlHoja1.Cells(nFila, 4) = rs!cConsDescripcion
'            xlHoja1.Cells(nFila, 5) = rs!cPersCIIU
'            xlHoja1.Cells(nFila, 6) = rs!cPersDireccDomicilio
'            xlHoja1.Cells(nFila, 7) = "0"
'            xlHoja1.Cells(nFila, 8) = ""
'            xlHoja1.Cells(nFila, 9) = rs!cPersIDnro
'            xlHoja1.Cells(nFila, 10) = rs!cTexto
'            rs.MoveNext
'        Loop
'
'        ExcelCuadro xlHoja1, 2, 11, 10, nFila, , True
'
'        xlHoja1.Cells.Select
'        xlHoja1.Cells.Font.Name = "Arial"
'        xlHoja1.Cells.Font.Size = 9
'        xlHoja1.Cells.EntireColumn.AutoFit
'    End If
'
'    oCon.CierraConexion
'    Set oCon = Nothing
'
'    MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "Aviso"
'End Sub
'   ALPA 20090929*********************************************************
'Private Sub GeneraRep19A_GrupoEconomicoEmpresa(ByVal pdFecha As Date)
'    Dim sSql As String
'    Dim rs As New ADODB.Recordset
'    Dim rs1 As New ADODB.Recordset
'    Dim I As Integer
'    Dim nFila As Integer
'    Dim nFilaTotal1 As Integer
'    Dim nFilaTotal2 As Integer
'
'    Dim nInicio As Integer
'
'    Dim nPatrEfectivo As Currency
'
'    Dim pdfecha1 As Date
'
'    pdfecha1 = DateAdd("m", 1, "01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Format(txtAnio.Text, "0000"))
'    pdfecha1 = DateAdd("d", -1, pdfecha1)
'
'    'Calculo el nPatrEfectivo
'    nPatrEfectivo = Val(txtPatrimonio.Text) 'Val(oAnx.GetImporteEstadAnexosMax(gdFecSis, "TOTALREP03", "1"))
'
'    Set oCon = New DConecta
'    oCon.AbreConexion
'
'    'Adiciona una hoja
'
'    ExcelAddHoja "Reporte 19A", xlLibro, xlHoja1
'
'    xlHoja1.PageSetup.Orientation = xlLandscape
'    xlHoja1.PageSetup.CenterHorizontally = True
'    xlHoja1.PageSetup.Zoom = 60
'
'    xlHoja1.Cells(2, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
'    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).MergeCells = True
'    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).HorizontalAlignment = xlLeft
'
'    xlHoja1.Cells(2, 10) = "REPORTE NRO 19-A"
'    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).MergeCells = True
'    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).HorizontalAlignment = xlLeft
'
'    xlHoja1.Cells(4, 2) = "INFORME SOBRE EL GRUPO ECONOMICO DE LA EMPRESA"
'    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).MergeCells = True
'    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).HorizontalAlignment = xlCenter
'
'    xlHoja1.Cells(6, 2) = "Empresa que remite la información: " & gsNomCmac
'    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).MergeCells = True
'    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).HorizontalAlignment = xlLeft
'
'    xlHoja1.Cells(7, 2) = "INFORMACION AL " & Format(pdfecha1, "DD MMMM YYYY")
'    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).MergeCells = True
'    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).HorizontalAlignment = xlLeft
'
'    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(7, 10)).Font.Bold = True
'
'    sSql = "Select PGE.CGECOD, PGR.cPersCodRel, P.cPersNombre, Con.cConsDescripcion, P.cPersCodSBS, P.cPersDireccDomicilio, "
'    sSql = sSql & " ISNULL(PID.cPersIDNro, '') as cPersIDNro "
'    sSql = sSql & " From PersGrupoEcon PGE "
'    sSql = sSql & "     Inner Join PersGERelacion PGR "
'    sSql = sSql & "         On PGE.cGECod=PGR.cGECod "
'    sSql = sSql & "     Inner Join Persona P "
'    sSql = sSql & "         On PGR.cPersCodRel = P.cPersCod "
'    sSql = sSql & "     Inner Join Constante Con "
'    sSql = sSql & "         on P.nPersPersoneria=Con.nConsValor "
'    sSql = sSql & "     Left Join PersID PID "
'    sSql = sSql & "         On PGR.cPersCodRel=PID.cPersCod "
'    sSql = sSql & " Where PGE.cCodReporte='19A' "
'    sSql = sSql & "     And Con.nConsCod=1002 "
'    sSql = sSql & "     And ISNULL(PID.cPersIDTpo,0)<>99 "
'    sSql = sSql & " Order By PGR.nOrden, P.cPersNombre "
'
'    Set rs = oCon.CargaRecordSet(sSql)
'    If rs.BOF Then
'    Else
'
'        nFila = 7
'
'        Do While Not rs.EOF
'
'            nFila = nFila + 3
'            xlHoja1.Cells(nFila, 2) = "1"
'            xlHoja1.Cells(nFila, 3) = "Razon o Denominación Social"
'            xlHoja1.Cells(nFila, 6) = PstaNombre(Trim(rs!cPersNombre))
'
'            xlHoja1.Range(xlHoja1.Cells(nFila, 3), xlHoja1.Cells(nFila, 4)).MergeCells = True
'            xlHoja1.Range(xlHoja1.Cells(nFila, 5), xlHoja1.Cells(nFila, 12)).MergeCells = True
'
'            nFila = nFila + 1
'            xlHoja1.Cells(nFila, 2) = "2"
'            xlHoja1.Cells(nFila, 3) = "Codigo SBS"
'            xlHoja1.Cells(nFila, 6) = rs!cPersCodSBS
'
'            xlHoja1.Range(xlHoja1.Cells(nFila, 3), xlHoja1.Cells(nFila, 4)).MergeCells = True
'            xlHoja1.Range(xlHoja1.Cells(nFila, 5), xlHoja1.Cells(nFila, 12)).MergeCells = True
'
'            nFila = nFila + 1
'            xlHoja1.Cells(nFila, 2) = "3"
'            xlHoja1.Cells(nFila, 3) = "RUC"
'            xlHoja1.Cells(nFila, 6) = rs!cPersIDnro
'
'            xlHoja1.Range(xlHoja1.Cells(nFila, 3), xlHoja1.Cells(nFila, 4)).MergeCells = True
'            xlHoja1.Range(xlHoja1.Cells(nFila, 5), xlHoja1.Cells(nFila, 12)).MergeCells = True
'
'            nFila = nFila + 1
'            xlHoja1.Cells(nFila, 2) = "4"
'            xlHoja1.Cells(nFila, 3) = "Dirección"
'            xlHoja1.Cells(nFila, 6) = rs!cPersDireccDomicilio
'
'            xlHoja1.Range(xlHoja1.Cells(nFila, 3), xlHoja1.Cells(nFila, 4)).MergeCells = True
'            xlHoja1.Range(xlHoja1.Cells(nFila, 5), xlHoja1.Cells(nFila, 12)).MergeCells = True
'
'            sSql = "Select Distinct P.cPersCod, P.cPersNombre "
'            sSql = sSql & " From productopersona PP "
'            sSql = sSql & "     Inner Join Persona P "
'            sSql = sSql & "         On P.cPersCod=PP.cPersCod "
'            sSql = sSql & " Where cCtaCod in ( "
'            sSql = sSql & "             select cCtaCod "
'            sSql = sSql & "             From Productopersona "
'            sSql = sSql & "             where cperscod='" & rs!cPersCodRel & "' "
'            sSql = sSql & "         )  and nprdpersrelac=12 "
'
'            Set rs1 = oCon.CargaRecordSet(sSql)
'                sSql = ""
'                Do While Not rs1.EOF
'                    If Len(Trim(sSql)) = 0 Then
'                        sSql = PstaNombre(Trim(rs1!cPersNombre))
'                    Else
'                        sSql = sSql & " y " & PstaNombre(Trim(rs1!cPersNombre))
'                    End If
'                    rs1.MoveNext
'                Loop
'
'            Set rs1 = Nothing
'
'            nFila = nFila + 1
'            xlHoja1.Cells(nFila, 2) = "5"
'            xlHoja1.Cells(nFila, 3) = "Representante Legal"
'            xlHoja1.Cells(nFila, 6) = sSql
'
'            sSql = ""
'
'            xlHoja1.Range(xlHoja1.Cells(nFila, 3), xlHoja1.Cells(nFila, 4)).MergeCells = True
'            xlHoja1.Range(xlHoja1.Cells(nFila, 5), xlHoja1.Cells(nFila, 12)).MergeCells = True
'
'            xlHoja1.Range(xlHoja1.Cells(nFila - 4, 2), xlHoja1.Cells(nFila, 12)).HorizontalAlignment = xlLeft
'            xlHoja1.Range(xlHoja1.Cells(nFila - 4, 2), xlHoja1.Cells(nFila, 3)).Font.Bold = True
'
'            ExcelCuadro xlHoja1, 2, nFila - 4, 12, nFila, , True
'
'            'Fin de Cabecera
'
'            nFila = nFila + 2
'
'            xlHoja1.Cells(nFila, 3) = "Nombre"
'            xlHoja1.Cells(nFila, 4) = "Cod SBS"
'            xlHoja1.Cells(nFila, 5) = "Tipo de"
'            xlHoja1.Cells(nFila, 6) = "Tipo de"
'            xlHoja1.Cells(nFila, 7) = "Numero del"
'            xlHoja1.Cells(nFila, 8) = "RUC"
'            xlHoja1.Cells(nFila, 9) = "Residencia"
'            xlHoja1.Cells(nFila, 10) = "Accionista"
'            xlHoja1.Cells(nFila, 11) = "Cargo"
'            xlHoja1.Cells(nFila, 12) = "Otro"
'
'            nFila = nFila + 1
'            xlHoja1.Cells(nFila, 5) = "Persona"
'            xlHoja1.Cells(nFila, 6) = "Documento"
'            xlHoja1.Cells(nFila, 7) = "Documento"
'            xlHoja1.Cells(nFila, 12) = "Cargo"
'
'            xlHoja1.Range(xlHoja1.Cells(nFila - 1, 3), xlHoja1.Cells(nFila, 12)).HorizontalAlignment = xlCenter
'            xlHoja1.Range(xlHoja1.Cells(nFila - 1, 3), xlHoja1.Cells(nFila, 12)).Font.Bold = True
'
'            ExcelCuadro xlHoja1, 3, nFila - 1, 12, nFila
'
'            sSql = "select PGV.cPersCodVinc, P.cPersNombre, P.cPersCodSBS, Con.cConsDescripcion, "
'            sSql = sSql & " ISNULL(PID.cPersIDNro, '') as cPersIDNro, Con1.cConsDescripcion as cConsDescripcion1, P.cPersDireccDomicilio, PGV.nCargo "
'            sSql = sSql & " from PersGEPersVinc PGV "
'            sSql = sSql & "     Inner Join Persona P "
'            sSql = sSql & "         On PGV.cPersCodVinc=P.cPersCod "
'            sSql = sSql & "     Inner Join Constante Con "
'            sSql = sSql & "         on P.nPersPersoneria=Con.nConsValor "
'            sSql = sSql & "     Left Join PersID PID "
'            sSql = sSql & "         On PGV.cPersCodVinc=PID.cPersCod "
'            sSql = sSql & "     Left Join Constante Con1 "
'            sSql = sSql & "         ON convert(int, PID.cPersIDTpo) = Con1.nConsValor "
'            sSql = sSql & " where PGV.cgecod='" & rs!cGECod & "' And PGV.cperscodrel='" & rs!cPersCodRel & "' "
'            sSql = sSql & "     And Con.nConsCod=1002 "
'            sSql = sSql & " And Con1.nConsCod=1003 "
'            sSql = sSql & " Order By PGV.nCargo, P.cPersNombre "
'
'            nFilaTotal1 = 0
'            Set rs1 = oCon.CargaRecordSet(sSql)
'            If rs1.BOF Then
'            Else
'                nFilaTotal1 = nFila + 1
'            End If
'
'            Do While Not rs1.EOF
'
'                nFila = nFila + 1
'
'                xlHoja1.Cells(nFila, 3) = PstaNombre(Trim(rs1!cPersNombre))
'                xlHoja1.Cells(nFila, 4) = rs1!cPersCodSBS
'                xlHoja1.Cells(nFila, 5) = rs1!cConsDescripcion
'                xlHoja1.Cells(nFila, 6) = rs1!cConsDescripcion1
'                xlHoja1.Cells(nFila, 7) = rs1!cPersIDnro
'                xlHoja1.Cells(nFila, 8) = ""
'                xlHoja1.Cells(nFila, 9) = rs1!cPersDireccDomicilio
'                xlHoja1.Cells(nFila, 10) = "0"
'                xlHoja1.Cells(nFila, 11) = rs1!nCargo
'                xlHoja1.Cells(nFila, 12) = "0"
'
'                rs1.MoveNext
'            Loop
'
'            Set rs1 = Nothing
'
'            If nFilaTotal1 > 0 Then
'                ExcelCuadro xlHoja1, 3, nFilaTotal1, 12, nFila, , True
'                xlHoja1.Range(xlHoja1.Cells(nFilaTotal1, 3), xlHoja1.Cells(nFila, 12)).Font.Bold = False
'            End If
'
'            rs.MoveNext
'        Loop
'
'    End If
'
'    oCon.CierraConexion
'    Set oCon = Nothing
'
'    xlHoja1.Cells.Select
'    xlHoja1.Cells.Font.Name = "Arial"
'    xlHoja1.Cells.Font.Size = 9
'    xlHoja1.Cells.EntireColumn.AutoFit
'
'    MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "Aviso"
'End Sub

Private Sub GeneraRep20_ClientesRepresRiesgoUnico(ByVal pdFecha As Date)
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim i As Integer
    Dim nFila As Integer
    Dim nFilaTotal1 As Integer
    Dim nFilaTotal2 As Integer
    
    Dim ntempo As Integer
    Dim nInicio As Integer
    
    Dim nPatrEfectivo As Currency
    
    Dim pdfecha1 As Date
    
    'pdfecha1 = DateAdd("m", 1, "01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Format(txtAnio.Text, "0000"))
    'pdfecha1 = DateAdd("d", -1, pdfecha1)
    
    'Calculo el nPatrEfectivo
    'nPatrEfectivo = Val(txtPatrimonio.Text) 'Val(oAnx.GetImporteEstadAnexosMax(gdFecSis, "TOTALREP03", "1"))
    
    Set oCon = New DConecta
    oCon.AbreConexion

    
    
    sSql = "Select PGE.CGECOD, PGR.cPersCodRel, P.cPersNombre, Con.cConsDescripcion, P.cPersCodSBS, P.cPersDireccDomicilio, "
    sSql = sSql & " ISNULL(PID.cPersIDNro, '') as cPersIDNro, P.nPersPersoneria "
    sSql = sSql & " From PersGrupoEcon PGE "
    sSql = sSql & "     Inner Join PersGERelacion PGR "
    sSql = sSql & "         On PGE.cGECod=PGR.cGECod "
    sSql = sSql & "     Inner Join Persona P "
    sSql = sSql & "         On PGR.cPersCodRel = P.cPersCod "
    sSql = sSql & "     Inner Join Constante Con "
    sSql = sSql & "         on P.nPersPersoneria=Con.nConsValor "
    sSql = sSql & "     Left Join PersID PID "
    sSql = sSql & "         On PGR.cPersCodRel=PID.cPersCod "
    sSql = sSql & "     Where PGE.cCodReporte='20' "
    sSql = sSql & "         And Con.nConsCod=1002 "
    sSql = sSql & "         And ISNULL(PID.cPersIDTpo,0)<>99 "
    sSql = sSql & "     Order By PGR.nOrden desc"
         
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.BOF Then
    Else
        i = rs.RecordCount
        
        Do While Not rs.EOF
            i = i - 1
            'Adiciona una hoja
    
            ExcelAddHoja "Registro" & Right("0" & Trim(Str(i)), 2), xlLibro, xlHoja1
            
            xlHoja1.PageSetup.Orientation = xlLandscape
            xlHoja1.PageSetup.CenterHorizontally = True
            xlHoja1.PageSetup.Zoom = 60
                 
            xlHoja1.Cells(2, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).HorizontalAlignment = xlLeft
             
            xlHoja1.Cells(2, 10) = "REPORTE NRO 20"
            xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).HorizontalAlignment = xlLeft
             
            xlHoja1.Cells(4, 2) = "INFORMACION DE CLIENTES QUE REPRESENTAN RIESGO UNICO"
            xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).HorizontalAlignment = xlCenter
            
            xlHoja1.Cells(6, 2) = "Empresa que remite la información: " & gsNomCmac
            xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).HorizontalAlignment = xlLeft
             
            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(6, 10)).Font.Bold = True
        
            nFila = 6
        
            sSql = "Select PGV.cPersCodVinc, P.cPersNombre, P.cPersCodSBS, Con.cConsDescripcion, "
            sSql = sSql & " ISNULL(PID.cPersIDNro, '') as cPersIDNro, Con1.cConsDescripcion as cConsDescripcion1, P.cPersDireccDomicilio, PGV.nCargo, Con2.cConsDescripcion as cConsDescripcion2 "
            sSql = sSql & " from PersGEPersVinc PGV "
            sSql = sSql & "      Inner Join Persona P "
            sSql = sSql & "         On PGV.cPersCodVinc=P.cPersCod "
            sSql = sSql & "      Inner Join Constante Con "
            sSql = sSql & "         on P.nPersPersoneria=Con.nConsValor "
            sSql = sSql & "      Inner Join Constante Con2 "
            sSql = sSql & "         on PGV.nCargo=Con2.nConsValor "
            sSql = sSql & "      Left Join PersID PID "
            sSql = sSql & "         On PGV.cPersCodVinc=PID.cPersCod "
            sSql = sSql & "      Left Join Constante Con1 "
            sSql = sSql & "         ON convert(int, PID.cPersIDTpo) = Con1.nConsValor "
            sSql = sSql & "      where PGV.cgecod='" & rs!cGECod & "' and PGV.cperscodrel='" & rs!cPersCodRel & "' "
            sSql = sSql & "         And Con.nConsCod=1002 "
            sSql = sSql & "         And Con1.nConsCod=1003 "
            sSql = sSql & "         And Con2.nConsCod=4029 "
            sSql = sSql & "         And ISNULL(PID.cPersIDTpo,0)<>99 "
            sSql = sSql & "      Order By PGV.nCargo "

            Set rs1 = oCon.CargaRecordSet(sSql)
        
            Do While Not rs1.EOF
                nFila = nFila + 3
                
                xlHoja1.Cells(nFila, 2) = "Nombre, Razon o Denominación Social"
                xlHoja1.Cells(nFila, 3) = PstaNombre(Trim(rs1!cPersNombre))
        
                nFila = nFila + 1
                xlHoja1.Cells(nFila, 2) = "Codigo SBS"
                xlHoja1.Cells(nFila, 3) = rs1!cPersCodSBS
            
                nFila = nFila + 1
                xlHoja1.Cells(nFila, 2) = "Tipo de Persona"
                xlHoja1.Cells(nFila, 3) = rs1!cConsDescripcion
                
                nFila = nFila + 1
                xlHoja1.Cells(nFila, 2) = "Documento de Identidad"
                xlHoja1.Cells(nFila, 3) = rs1!cConsDescripcion1 & " " & rs1!cPersIDnro
                
                nFila = nFila + 1
                xlHoja1.Cells(nFila, 2) = "RUC"
                xlHoja1.Cells(nFila, 3) = ""
                
                nFila = nFila + 1
                xlHoja1.Cells(nFila, 2) = "Direccion"
                xlHoja1.Cells(nFila, 3) = rs1!cPersDireccDomicilio
                
                nFila = nFila + 1
                xlHoja1.Cells(nFila, 2) = "Representante Legal"
                xlHoja1.Cells(nFila, 3) = ""
                
                ExcelCuadro xlHoja1, 2, nFila - 6, 3, nFila, , True
                
                xlHoja1.Range(xlHoja1.Cells(nFila - 6, 2), xlHoja1.Cells(nFila, 2)).Font.Bold = True
                
                rs1.MoveNext
            Loop
            
            nFila = nFila + 3
            xlHoja1.Cells(nFila, 2) = "Nombre, Razon o Denominación Social"
            xlHoja1.Cells(nFila, 3) = PstaNombre(Trim(rs!cPersNombre))
    
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 2) = "Codigo SBS"
            xlHoja1.Cells(nFila, 3) = rs!cPersCodSBS
        
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 2) = "Tipo de Persona"
            xlHoja1.Cells(nFila, 3) = rs!cConsDescripcion
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 2) = "Documento de Identidad"
            xlHoja1.Cells(nFila, 3) = ""
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 2) = "RUC"
            xlHoja1.Cells(nFila, 3) = rs!cPersIDnro
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 2) = "Direccion"
            xlHoja1.Cells(nFila, 3) = rs!cPersDireccDomicilio
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 2) = "Representante Legal"
            
            'Representante Legal
            
            sSql = "Select Distinct P.cPersCod, P.cPersNombre "
            sSql = sSql & " From productopersona PP "
            sSql = sSql & "     Inner Join Persona P "
            sSql = sSql & "         On P.cPersCod=PP.cPersCod "
            sSql = sSql & " Where cCtaCod in ( "
            sSql = sSql & "             select cCtaCod "
            sSql = sSql & "             From Productopersona "
            sSql = sSql & "             where cperscod='" & rs!cPersCodRel & "' "
            sSql = sSql & "         )  and nprdpersrelac=12 "
        
            Set rs2 = oCon.CargaRecordSet(sSql)
                sSql = ""
                Do While Not rs2.EOF
                    If Len(Trim(sSql)) = 0 Then
                        sSql = PstaNombre(Trim(rs2!cPersNombre))
                    Else
                        sSql = sSql & " y " & PstaNombre(Trim(rs2!cPersNombre))
                    End If
                    rs2.MoveNext
                Loop
                
            Set rs2 = Nothing
            
            xlHoja1.Cells(nFila, 3) = sSql
            ExcelCuadro xlHoja1, 2, nFila - 6, 3, nFila, , True
            
            xlHoja1.Range(xlHoja1.Cells(nFila - 6, 2), xlHoja1.Cells(nFila, 2)).Font.Bold = True
            
            sSql = ""
            
            'Fin de Representante Legal
            
            nFila = nFila + 2
            
            xlHoja1.Cells(nFila, 2) = "ACCIONISTAS, DIRECTORES, GERENTES, PRINCIPALES FUNCIONARIOS Y ASESORES"
            xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 4)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 4)).Font.Bold = True
            
            nFila = nFila + 2
            
            xlHoja1.Cells(nFila, 2) = "Nombre, Razon o Denominación Social"
            xlHoja1.Cells(nFila, 3) = "Cod SBS"
            xlHoja1.Cells(nFila, 4) = "Tipo de"
            xlHoja1.Cells(nFila, 5) = "Tipo de"
            xlHoja1.Cells(nFila, 6) = "Numero del"
            xlHoja1.Cells(nFila, 7) = "RUC"
            xlHoja1.Cells(nFila, 8) = "Residencia"
            xlHoja1.Cells(nFila, 9) = "Accionista"
            xlHoja1.Cells(nFila, 10) = "Cargo"
            xlHoja1.Cells(nFila, 11) = "Otro"
        
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 4) = "Persona"
            xlHoja1.Cells(nFila, 5) = "Documento"
            xlHoja1.Cells(nFila, 6) = "Documento"
            xlHoja1.Cells(nFila, 11) = "Cargo"
        
            xlHoja1.Range(xlHoja1.Cells(nFila - 1, 2), xlHoja1.Cells(nFila, 11)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(nFila - 1, 2), xlHoja1.Cells(nFila, 11)).Font.Bold = True
        
            ExcelCuadro xlHoja1, 2, nFila - 1, 11, nFila
        
            ''
            
            If rs!nPersPersoneria <> 1 Then
                rs1.MoveFirst
                ntempo = 0
                nFilaTotal1 = nFila + 1
                Do While Not rs1.EOF
                    
                    nFila = nFila + 1
                    
                    If ntempo = 0 Then
                        xlHoja1.Cells(nFila, 2) = PstaNombre(Trim(rs!cPersNombre))
                        xlHoja1.Cells(nFila, 3) = rs!cPersCodSBS
                        xlHoja1.Cells(nFila, 4) = rs!cConsDescripcion
                        xlHoja1.Cells(nFila, 5) = ""
                        xlHoja1.Cells(nFila, 6) = ""
                        xlHoja1.Cells(nFila, 7) = rs!cPersIDnro
                        xlHoja1.Cells(nFila, 8) = "Peru"
                        ntempo = 1
                    End If
                    xlHoja1.Cells(nFila, 9) = PstaNombre(Trim(rs1!cPersNombre))
                    xlHoja1.Cells(nFila, 10) = rs1!cConsDescripcion2
                    xlHoja1.Cells(nFila, 11) = ""
                
                    rs1.MoveNext
                    
                Loop
            Else
                nFilaTotal1 = nFila + 1
                nFila = nFila + 1
            End If
            
            ExcelCuadro xlHoja1, 2, nFilaTotal1, 11, nFila
            
            rs1.Close
            
            rs.MoveNext
            
            xlHoja1.Cells.Select
            xlHoja1.Cells.Font.Name = "Arial"
            xlHoja1.Cells.Font.Size = 9
            xlHoja1.Cells.EntireColumn.AutoFit
            
        Loop
    
    End If
    
    Set rs1 = Nothing
    Set rs = Nothing
    
    oCon.CierraConexion
    Set oCon = Nothing
    
'    xlHoja1.Cells.Select
'    xlHoja1.Cells.Font.Name = "Arial"
'    xlHoja1.Cells.Font.Size = 9
'    xlHoja1.Cells.EntireColumn.AutoFit
    
    MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "Aviso"
End Sub

Private Sub GeneraRep20A_ClientesRepresRiesgoUnico(ByVal pdFecha As Date)
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim i As Integer
    Dim nFila As Integer
    Dim nFilaTotal1 As Integer
    Dim nFilaTotal2 As Integer
    
    Dim ntempo As Integer
    Dim nInicio As Integer
    
    Dim nPatrEfectivo As Currency
    
    Dim pdfecha1 As Date
    
    'pdfecha1 = DateAdd("m", 1, "01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Format(txtAnio.Text, "0000"))
    'pdfecha1 = DateAdd("d", -1, pdfecha1)
    
    'Calculo el nPatrEfectivo
    'nPatrEfectivo = Val(txtPatrimonio.Text) 'Val(oAnx.GetImporteEstadAnexosMax(gdFecSis, "TOTALREP03", "1"))
    
    Set oCon = New DConecta
    oCon.AbreConexion

    ExcelAddHoja "Reporte20A", xlLibro, xlHoja1

    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    
    xlHoja1.Cells(2, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).HorizontalAlignment = xlLeft
     
    xlHoja1.Cells(2, 10) = "REPORTE NRO 20-A"
    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).HorizontalAlignment = xlLeft
      
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(6, 14)).Font.Bold = True

    nFila = 2
    
    sSql = "Select PGE.CGECOD, PGR.cPersCodRel, P.cPersNombre, Con.cConsDescripcion, P.cPersCodSBS, P.cPersDireccDomicilio, P.cPersCIIU, "
    sSql = sSql & " ISNULL(PID.cPersIDNro, '') as cPersIDNro, PGR.nRela1, PGR.nRela2, PGR.nRela3 "
    sSql = sSql & " From PersGrupoEcon PGE "
    sSql = sSql & "     Inner Join PersGERelacion PGR "
    sSql = sSql & "         On PGE.cGECod=PGR.cGECod "
    sSql = sSql & "     Inner Join Persona P "
    sSql = sSql & "         On PGR.cPersCodRel = P.cPersCod "
    sSql = sSql & "     Inner Join Constante Con "
    sSql = sSql & "         on P.nPersPersoneria=Con.nConsValor "
    sSql = sSql & "     Left Join PersID PID "
    sSql = sSql & "         On PGR.cPersCodRel=PID.cPersCod "
    sSql = sSql & "     Where PGE.cCodReporte='20' "
    sSql = sSql & "         And Con.nConsCod=1002 "
    sSql = sSql & "         And ISNULL(PID.cPersIDTpo,0)<>99 "
    sSql = sSql & "     Order By PGR.nOrden"
         
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.BOF Then
    Else
        
        Do While Not rs.EOF
        
            i = 1
        
            nFila = nFila + 3
        
            xlHoja1.Cells(nFila, 2) = "INFORMACION DE LAS PERSONAS QUE REPRESENTAN RIESGO UNICO CLIENTES"
            xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 14)).MergeCells = True
            xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 14)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 14)).Font.Bold = True
        
            nFila = nFila + 1
                
            'Cabecera
            xlHoja1.Cells(nFila, 2) = "Nro"
            xlHoja1.Cells(nFila, 3) = "Código"
            xlHoja1.Cells(nFila, 4) = "Código"
            xlHoja1.Cells(nFila, 5) = "Nombre/Razón/"
            xlHoja1.Cells(nFila, 6) = "CIIU"
            xlHoja1.Cells(nFila, 7) = "Domicilio"
            xlHoja1.Cells(nFila, 8) = "Tipo"
            xlHoja1.Cells(nFila, 9) = "Tipo de"
            xlHoja1.Cells(nFila, 10) = "Múmero"
            xlHoja1.Cells(nFila, 11) = "RUC"
            xlHoja1.Cells(nFila, 12) = "Descripcion de la Vinculación"
            xlHoja1.Range(xlHoja1.Cells(nFila, 12), xlHoja1.Cells(nFila, 14)).MergeCells = True
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 2) = "de"
            xlHoja1.Cells(nFila, 3) = "SBS del"
            xlHoja1.Cells(nFila, 4) = "SBS"
            xlHoja1.Cells(nFila, 5) = "Denominación"
            xlHoja1.Cells(nFila, 8) = "de"
            xlHoja1.Cells(nFila, 9) = "Doc. de"
            xlHoja1.Cells(nFila, 10) = "Documento"
            xlHoja1.Cells(nFila, 12) = "Propiedad"
            xlHoja1.Cells(nFila, 13) = "Propiedad"
            xlHoja1.Cells(nFila, 14) = "Gestión"
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 2) = "Vinculado"
            xlHoja1.Cells(nFila, 3) = "Cliente"
            xlHoja1.Cells(nFila, 5) = "Social"
            xlHoja1.Cells(nFila, 8) = "persona"
            xlHoja1.Cells(nFila, 9) = "Identidad"
            xlHoja1.Cells(nFila, 10) = "Identidad"
            xlHoja1.Cells(nFila, 12) = "Directa"
            xlHoja1.Cells(nFila, 13) = "Indirecta"
            
            xlHoja1.Range(xlHoja1.Cells(nFila - 2, 2), xlHoja1.Cells(nFila, 14)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(nFila - 2, 2), xlHoja1.Cells(nFila, 14)).HorizontalAlignment = xlCenter
            
            ExcelCuadro xlHoja1, 2, nFila - 2, 14, nFila
        
            'Fin Cabecera
            
            nFila = nFila + 1
            nFilaTotal1 = nFila
            
            xlHoja1.Cells(nFila, 2) = i
            xlHoja1.Cells(nFila, 3) = rs!cPersCodSBS
            xlHoja1.Cells(nFila, 4) = rs!cPersCodSBS
            xlHoja1.Cells(nFila, 5) = PstaNombre(Trim(rs!cPersNombre))
            xlHoja1.Cells(nFila, 6) = rs!cPersCIIU
            xlHoja1.Cells(nFila, 7) = rs!cPersDireccDomicilio
            xlHoja1.Cells(nFila, 8) = rs!cConsDescripcion
            xlHoja1.Cells(nFila, 9) = ""
            xlHoja1.Cells(nFila, 10) = ""
            xlHoja1.Cells(nFila, 11) = rs!cPersIDnro
            xlHoja1.Cells(nFila, 12) = IIf(rs!nRela1 = 1, "X", "")
            xlHoja1.Cells(nFila, 13) = IIf(rs!nRela2 = 1, "X", "")
            xlHoja1.Cells(nFila, 14) = IIf(rs!nRela3 = 1, "X", "")
            
        
            sSql = "Select PGV.cPersCodVinc, P.cPersNombre, P.cPersCodSBS, Con.cConsDescripcion, P.cPersCIIU, "
            sSql = sSql & " ISNULL(PID.cPersIDNro, '') as cPersIDNro, Con1.cConsDescripcion as cConsDescripcion1, P.cPersDireccDomicilio, PGV.nCargo, Con2.cConsDescripcion as cConsDescripcion2, PGV.nRela1, PGV.nRela2, PGV.nRela3 "
            sSql = sSql & " from PersGEPersVinc PGV "
            sSql = sSql & "      Inner Join Persona P "
            sSql = sSql & "         On PGV.cPersCodVinc=P.cPersCod "
            sSql = sSql & "      Inner Join Constante Con "
            sSql = sSql & "         on P.nPersPersoneria=Con.nConsValor "
            sSql = sSql & "      Inner Join Constante Con2 "
            sSql = sSql & "         on PGV.nCargo=Con2.nConsValor "
            sSql = sSql & "      Left Join PersID PID "
            sSql = sSql & "         On PGV.cPersCodVinc=PID.cPersCod "
            sSql = sSql & "      Left Join Constante Con1 "
            sSql = sSql & "         ON convert(int, PID.cPersIDTpo) = Con1.nConsValor "
            sSql = sSql & "      where PGV.cgecod='" & rs!cGECod & "' and PGV.cperscodrel='" & rs!cPersCodRel & "' "
            sSql = sSql & "         And Con.nConsCod=1002 "
            sSql = sSql & "         And Con1.nConsCod=1003 "
            sSql = sSql & "         And Con2.nConsCod=4029 "
            sSql = sSql & "         And ISNULL(PID.cPersIDTpo,0)<>99 "
            sSql = sSql & "      Order By PGV.nCargo "

            Set rs1 = oCon.CargaRecordSet(sSql)
        
            Do While Not rs1.EOF
                nFila = nFila + 1
                i = i + 1
                    
                xlHoja1.Cells(nFila, 2) = i
                xlHoja1.Cells(nFila, 3) = rs1!cPersCodSBS
                xlHoja1.Cells(nFila, 4) = rs1!cPersCodSBS
                xlHoja1.Cells(nFila, 5) = PstaNombre(Trim(rs1!cPersNombre))
                xlHoja1.Cells(nFila, 6) = rs1!cPersCIIU
                xlHoja1.Cells(nFila, 7) = rs1!cPersDireccDomicilio
                xlHoja1.Cells(nFila, 8) = rs1!cConsDescripcion
                xlHoja1.Cells(nFila, 9) = rs1!cConsDescripcion1
                xlHoja1.Cells(nFila, 10) = rs1!cPersIDnro
                xlHoja1.Cells(nFila, 11) = ""
                xlHoja1.Cells(nFila, 12) = IIf(rs1!nRela1 = 1, "X", "")
                xlHoja1.Cells(nFila, 13) = IIf(rs1!nRela2 = 1, "X", "")
                xlHoja1.Cells(nFila, 14) = IIf(rs1!nRela3 = 1, "X", "")
                
                rs1.MoveNext
            Loop
            
            
            ExcelCuadro xlHoja1, 2, nFilaTotal1, 14, nFila
            
            rs1.Close
            
            rs.MoveNext
            
            
            
        Loop
    
    End If
    
    Set rs1 = Nothing
    Set rs = Nothing
    
    
    oCon.CierraConexion
    Set oCon = Nothing
    
    xlHoja1.Cells.Select
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 9
    xlHoja1.Cells.EntireColumn.AutoFit
    
    MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "Aviso"
End Sub


Private Sub GeneraRep21_ClientesRepresRiesgoUnico()
    
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim nFila As String
    Dim nFilaTotal1 As Integer
    Dim nFilaTotal2 As Integer
    
    Dim nInicio As Integer
    
    Dim nPatrEfectivo As Currency
    
    'Dim oAnx As New NEstadisticas
    
    Dim pdFecha As Date
    
    pdFecha = DateAdd("m", 1, "01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Format(txtAnio.Text, "0000"))
    pdFecha = DateAdd("d", -1, pdFecha)
    
    'Calculo el nPatrEfectivo
    nPatrEfectivo = Val(txtPatrimonio.Text) 'Val(oAnx.GetImporteEstadAnexosMax(gdFecSis, "TOTALREP03", "1"))
    
    Set oCon = New DConecta
    oCon.AbreConexion

    'Adiciona una hoja
    
    ExcelAddHoja "Parte A", xlLibro, xlHoja1
    
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
        
    xlHoja1.Cells(2, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(2, 10) = "REPORTE NRO 21-A"
    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(4, 2) = "INFORMACION DE LAS PERSONAS JURIDICAS VINCULADAS A LA EMPRESA"
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).HorizontalAlignment = xlCenter
   
    xlHoja1.Cells(6, 2) = "Empresa que remite la información: " & gsNomCmac
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(7, 2) = "INFORMACION AL " & Format(pdFecha, "DD MMMM YYYY")
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(9, 2) = "Razón o denominación social de la persona jurídica: Municipalidad Provincial de Ica"
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(9, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(9, 6)).HorizontalAlignment = xlLeft
    
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(9, 11)).Font.Bold = True
    
    sSql = "Select  ger.cPersCodRel, P.cPersNombre cNomPers, P.nPersPersoneria, "
    sSql = sSql & " IsNull((    Select Top 1 cPersIDnro + '*' + Case cPersIDTPO "
    sSql = sSql & "             When 1 Then 'DNI'"
    sSql = sSql & "             When 4 Then 'CE'"
    sSql = sSql & "             When 11 Then 'PAS'"
    sSql = sSql & "             End"
    sSql = sSql & "         From PersId PID"
    sSql = sSql & "         Where PID.cPersCod = P.cPersCod And cPersIDTPO In (1,4,11)"
    sSql = sSql & "         order by cPersIDTPO"
    sSql = sSql & "        ),'*') Doc,"
    sSql = sSql & " ISNUll(PID.cPersIDnro,'') RUC, ger.nPrdPersRelac, gePV.cPersCodRel, PerVinc.cPersNombre as cNomPersVinc,"
    sSql = sSql & " PerVinc.nPersPersoneria,"
    sSql = sSql & " IsNUll((    Select Top 1 cPersIDnro + '*' + Case cPersIDTPO"
    sSql = sSql & "             When 1 Then 'DNI' When 4 Then 'C.E' When 11 Then 'PAS' End "
    sSql = sSql & "         From PersId PID"
    sSql = sSql & "         Where PID.cPersCod = PerVinc.cPersCod And cPersIDTPO In (1,4,11)"
    sSql = sSql & "         order by cPersIDTPO"
    sSql = sSql & "     ),'*') DocVin,"
    sSql = sSql & " ISNUll(PID.cPersIDnro,'') RUCV, PerVinc.cPersDireccDomicilio as cDirPersVinc,"
    sSql = sSql & " gePV.nCargo, gePV.nParticip, Co.cConsDescripcion, dbo.RiesgoGetCreditos (PerVinc.cPersCod, '" & Format(pdFecha, "MM/dd/YYYY") & "', " & Val(txtTC.Text) & ") as nSaldoCredito,"
    sSql = sSql & " CON.cConsDescripcion as cConsDescripcion1"
    sSql = sSql & " From PersGrupoEcon ge"
    sSql = sSql & " inner join PersGERelacion ger on ge.cGECod = ger.cGECod"
    sSql = sSql & " Inner join Persona p on p.cPersCod = ger.cPersCodRel"
    sSql = sSql & " Inner join PersGEPersVinc gePV on gePV.cPersCodRel = ger.cPersCodRel and gepv.cgecod=ger.cgecod "
    sSql = sSql & " Inner join Persona PerVinc on PerVinc.cPersCod = gePV.cPersCodVinc"
    sSql = sSql & " Inner Join Constante Co on gePV.ncargo=Co.nConsValor"
    sSql = sSql & " Inner Join Constante Con on PerVinc.nPersPersoneria=Con.nConsValor"
    sSql = sSql & " Left Join PersID PID On PID.cPersCod = P.cPersCod And PID.cPersIDTpo = 2"
    sSql = sSql & " Left Join PersID PIDV On PIDV.cPersCod = P.cPersCod And PIDV.cPersIDTpo = 2"

    sSql = sSql & " Where ge.cCodReporte='21' and p.nPersPersoneria <> '1' " 'and gePV.cPersCodRel ='1089800000272'"
    sSql = sSql & " And Co.nConsCod=4029"
    sSql = sSql & " And Con.nConsCod=1002"
    sSql = sSql & " And dbo.RiesgoGetCreditos (PerVinc.cPersCod, '" & Format(pdFecha, "MM/dd/YYYY") & "', " & Val(txtTC.Text) & ")>0 "
     
         
    Set rs = oCon.CargaRecordSet(sSql)
    If rs.BOF Then
    Else
        nFila = 11
        xlHoja1.Cells(nFila, 2) = "Nombre"
        xlHoja1.Cells(nFila, 3) = "Cod"
        xlHoja1.Cells(nFila, 4) = "Tipo de"
        xlHoja1.Cells(nFila, 5) = "Tipo de"
        xlHoja1.Cells(nFila, 6) = "Número del"
        xlHoja1.Cells(nFila, 7) = "RUC"
        xlHoja1.Cells(nFila, 8) = "Residencia"
        xlHoja1.Cells(nFila, 9) = "Accionista"
        xlHoja1.Cells(nFila, 10) = "Cargo"
        xlHoja1.Cells(nFila, 11) = "Otro"
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 3) = "SBS"
        xlHoja1.Cells(nFila, 4) = "Persona"
        xlHoja1.Cells(nFila, 5) = "Documento"
        xlHoja1.Cells(nFila, 6) = "Documento"
        xlHoja1.Cells(nFila, 11) = "Cargo"
        
        xlHoja1.Range(xlHoja1.Cells(nFila - 1, 2), xlHoja1.Cells(nFila, 11)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(nFila - 1, 2), xlHoja1.Cells(nFila, 11)).Font.Bold = True
        
        ExcelCuadro xlHoja1, 2, nFila - 1, 11, nFila
        
        Do While Not rs.EOF
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 2) = rs!cNomPersVinc
            xlHoja1.Cells(nFila, 3) = "-"
            xlHoja1.Cells(nFila, 4) = rs!cConsDescripcion1
            xlHoja1.Cells(nFila, 5) = Right(rs!DocVin, 3)
            xlHoja1.Cells(nFila, 6) = Trim(Mid(rs!DocVin, 1, Len(rs!DocVin) - 4))
            xlHoja1.Cells(nFila, 7) = "-"
            xlHoja1.Cells(nFila, 8) = "Perú"
            xlHoja1.Cells(nFila, 9) = "-"
            xlHoja1.Cells(nFila, 10) = "-"
            xlHoja1.Cells(nFila, 11) = rs!cConsDescripcion
            rs.MoveNext
        Loop
        
        ExcelCuadro xlHoja1, 2, 11, 11, nFila, , True
        
        xlHoja1.Cells.Select
        xlHoja1.Cells.Font.Name = "Arial"
        xlHoja1.Cells.Font.Size = 9
        xlHoja1.Cells.EntireColumn.AutoFit
    End If
    

    ' Adiciona una hoja
    ExcelAddHoja "Parte B", xlLibro, xlHoja1
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
       
    nFila = 0
       
    xlHoja1.Cells(2, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 6)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(2, 10) = "REPORTE NRO 21-A"
    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(2, 10), xlHoja1.Cells(2, 11)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(4, 2) = "INFORMACION DE LAS PERSONAS JURIDICAS VINCULADAS A LA EMPRESA"
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).HorizontalAlignment = xlCenter
      
    xlHoja1.Cells(6, 2) = "Empresa que remite la información: " & gsNomCmac
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, 6)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(7, 2) = "INFORMACION AL " & Format(pdFecha, "DD MMMM YYYY")
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).HorizontalAlignment = xlLeft
      
    xlHoja1.Cells(9, 2) = "1. Vinculados por el Artículo 202 de la Ley General"
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(9, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(9, 6)).HorizontalAlignment = xlLeft
    
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(9, 11)).Font.Bold = True
    
    If rs.BOF Then
    Else
        rs.MoveFirst
        
        nFila = 11
        
        i = 0
        
        xlHoja1.Cells(nFila, 2) = "Nro"
        xlHoja1.Cells(nFila, 3) = "Cod"
        xlHoja1.Cells(nFila, 4) = "Nombre/Razon/"
        xlHoja1.Cells(nFila, 5) = "CIIU"
        xlHoja1.Cells(nFila, 6) = "Domicilio"
        xlHoja1.Cells(nFila, 7) = "Tipo de"
        xlHoja1.Cells(nFila, 8) = "Tipo de Doc"
        xlHoja1.Cells(nFila, 9) = "Num."
        xlHoja1.Cells(nFila, 10) = "RUC"
        xlHoja1.Cells(nFila, 11) = "Descripción de la Vinculación"
        xlHoja1.Range(xlHoja1.Cells(nFila, 11), xlHoja1.Cells(nFila, 13)).MergeCells = True
        xlHoja1.Cells(nFila, 14) = "Créditos"
        xlHoja1.Cells(nFila, 15) = "Inversiones"
        xlHoja1.Cells(nFila, 16) = "Contingentes"
        xlHoja1.Cells(nFila, 17) = "Arrendamiento"
        xlHoja1.Cells(nFila, 18) = "Total"
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 3) = "SBS"
        xlHoja1.Cells(nFila, 4) = "denominación"
        xlHoja1.Cells(nFila, 7) = "persona"
        xlHoja1.Cells(nFila, 8) = "de identidad"
        xlHoja1.Cells(nFila, 9) = "Documento de"
        xlHoja1.Cells(nFila, 11) = "Propiedad"
        xlHoja1.Cells(nFila, 12) = "Propiedad"
        xlHoja1.Cells(nFila, 13) = "Gestión"
        xlHoja1.Cells(nFila, 17) = "Financiero"
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 4) = "social"
        xlHoja1.Cells(nFila, 9) = "identidad"
        xlHoja1.Cells(nFila, 11) = "Directa"
        xlHoja1.Cells(nFila, 12) = "Indirecta"
        
        xlHoja1.Range(xlHoja1.Cells(nFila - 2, 2), xlHoja1.Cells(nFila, 18)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(nFila - 2, 2), xlHoja1.Cells(nFila, 18)).Font.Bold = True
        
        ExcelCuadro xlHoja1, 2, nFila - 2, 18, nFila
        
        nInicio = nFila + 1
        
        Do While Not rs.EOF
            nFila = nFila + 1
            i = i + 1
            xlHoja1.Cells(nFila, 2) = i
            xlHoja1.Cells(nFila, 3) = "-"
            xlHoja1.Cells(nFila, 4) = rs!cNomPersVinc
            xlHoja1.Cells(nFila, 5) = "-"
            xlHoja1.Cells(nFila, 6) = rs!cDirPersVinc
            xlHoja1.Cells(nFila, 7) = rs!cConsDescripcion1
            xlHoja1.Cells(nFila, 8) = Right(rs!DocVin, 3)
            xlHoja1.Cells(nFila, 9) = Trim(Mid(rs!DocVin, 1, Len(rs!DocVin) - 4))
            xlHoja1.Cells(nFila, 10) = "-"
            xlHoja1.Cells(nFila, 11) = "-"
            xlHoja1.Cells(nFila, 12) = "-"
            xlHoja1.Cells(nFila, 13) = "-"
            xlHoja1.Cells(nFila, 14) = Format(rs!nSaldoCredito / 1000, "0.00")
            xlHoja1.Cells(nFila, 15) = "-"
            xlHoja1.Cells(nFila, 16) = "-"
            xlHoja1.Cells(nFila, 17) = "-"
            xlHoja1.Cells(nFila, 18) = Format(rs!nSaldoCredito / 1000, "0.00")
            rs.MoveNext
        Loop
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 2) = "Total Vinculados por el Articulo 202"
        xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 13)).MergeCells = True
        xlHoja1.Range("N" & nFila & ":N" & nFila).Formula = "=SUM(N" & nInicio & ":N" & nFila - 1 & ")"
        xlHoja1.Range("R" & nFila & ":R" & nFila).Formula = "=SUM(R" & nInicio & ":R" & nFila - 1 & ")"
        
        xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 18)).Font.Bold = True
        
        nFilaTotal1 = nFila
        
        ExcelCuadro xlHoja1, 2, nInicio, 18, nFila, , True
        
    End If
    rs.Close
    
    sSql = "Select C1.cPersCod, C1.cIFTpo, C1.cPersNombre, "
    sSql = sSql & " C1.cPersIDNro, C1.cConsDescripcion, C1.cPersCIIU, C1.cPersDireccDomicilio, ISNULL(dbo.RiesgoGetCreditos (C1.cPersCod, '" & Format(pdFecha, "MM/dd/YYYY") & "', " & Val(txtTC.Text) & "),0) as nSaldoCredito, "
    sSql = sSql & " SUM(nSaldoCaptacion) As nSaldoCaptacion "
    sSql = sSql & " From "
    sSql = sSql & " ( "
    sSql = sSql & " Select CIF.cPersCod, CIF.cIFTpo, P.cPersNombre, "
    sSql = sSql & " PID.cPersIDNro, CON.cConsDescripcion, P.cPersCIIU, P.cPersDireccDomicilio , "
    sSql = sSql & " nSaldoCaptacion = "
    sSql = sSql & " dbo.GetSaldoCtaIF ('" & Format(pdFecha, "MM/dd/YYYY") & "',  CIFF.cCtaContCod + CIFF.cCtaIFSubCta, "
    sSql = sSql & "             CIF.cPersCod, CIF.cIFTpo, CIF.cCtaIFCod, Substring(CIF.cCtaIFCod, 3,1)) "
    sSql = sSql & " * case  when Substring(cIF.cCtaIFCod, 3,1)='1' then 1 "
    sSql = sSql & "         when Substring(cIF.cCtaIFCod, 3,1)='2' then " & Val(txtTC.Text) & " "
    sSql = sSql & "       End"
    sSql = sSql & " From CtaIF CIF "
    sSql = sSql & " Inner Join Persona P "
    sSql = sSql & "     On CIF.cPersCod=P.cPersCod"
    sSql = sSql & " Inner Join CtaIFFiltro CIFF"
    sSql = sSql & "     ON CIF.cPErsCod=CIFF.cPersCod and CIF.cIFTpo=CIFF.cIFTpo And CIF.cCtaIFCod=CIFF.cCtaIFCod"
    sSql = sSql & " Inner Join PersID PID"
    sSql = sSql & "     On P.cPersCod=PID.cPersCod"
    sSql = sSql & " Inner Join Constante Con"
    sSql = sSql & "     on P.nPersPersoneria=Con.nConsValor"
    sSql = sSql & " Where CIF.cIFTpo in ('01', '03')"
    sSql = sSql & " And CIF.cCtaIFCod like '0[123]%'"
    sSql = sSql & " and pid.cpersidtpo='2' And Con.nConsCod=1002"
    sSql = sSql & " ) c1"
    sSql = sSql & " Group By"
    sSql = sSql & " C1.cPersCod, C1.cIFTpo, C1.cPersNombre,"
    sSql = sSql & " C1.cPersIDNro , C1.cConsDescripcion, C1.cPersCIIU, C1.cPersDireccDomicilio, dbo.RiesgoGetCreditos (C1.cPersCod, '" & Format(pdFecha, "MM/dd/YYYY") & "', " & Val(txtTC.Text) & ")  "
    sSql = sSql & " Order by C1.cIFTpo, C1.cPersCod"

    Set rs = oCon.CargaRecordSet(sSql)
    
    nFila = nFila + 2
    xlHoja1.Cells(nFila, 2) = "2. Vinculados por el Artículo 204 de la Ley General"
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 6)).HorizontalAlignment = xlLeft
    
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 18)).Font.Bold = True
    
    If rs.BOF Then
    Else
    
        nFila = nFila + 1
        
        i = 0
        
        xlHoja1.Cells(nFila, 2) = "Nro"
        xlHoja1.Cells(nFila, 3) = "Cod"
        xlHoja1.Cells(nFila, 4) = "Nombre/Razon/"
        xlHoja1.Cells(nFila, 5) = "CIIU"
        xlHoja1.Cells(nFila, 6) = "Domicilio"
        xlHoja1.Cells(nFila, 7) = "Tipo de"
        xlHoja1.Cells(nFila, 8) = "Tipo de Doc"
        xlHoja1.Cells(nFila, 9) = "Num."
        xlHoja1.Cells(nFila, 10) = "RUC"
        xlHoja1.Cells(nFila, 11) = "Descripción de la Vinculación"
        xlHoja1.Range(xlHoja1.Cells(nFila, 11), xlHoja1.Cells(nFila, 13)).MergeCells = True
        xlHoja1.Cells(nFila, 14) = "Financiamiento"
        xlHoja1.Cells(nFila, 15) = "Depósitos"
        xlHoja1.Cells(nFila, 16) = "Contingentes"
        xlHoja1.Cells(nFila, 17) = "Arrendamiento"
        xlHoja1.Cells(nFila, 18) = "Total"
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 3) = "SBS"
        xlHoja1.Cells(nFila, 4) = "denominación"
        xlHoja1.Cells(nFila, 7) = "persona"
        xlHoja1.Cells(nFila, 8) = "de identidad"
        xlHoja1.Cells(nFila, 9) = "Documento de"
        xlHoja1.Cells(nFila, 11) = "Propiedad"
        xlHoja1.Cells(nFila, 12) = "Propiedad"
        xlHoja1.Cells(nFila, 13) = "Gestión"
        xlHoja1.Cells(nFila, 17) = "Financiero"
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 4) = "social"
        xlHoja1.Cells(nFila, 9) = "identidad"
        xlHoja1.Cells(nFila, 11) = "Directa"
        xlHoja1.Cells(nFila, 12) = "Indirecta"
        
        xlHoja1.Range(xlHoja1.Cells(nFila - 2, 2), xlHoja1.Cells(nFila, 18)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(nFila - 2, 2), xlHoja1.Cells(nFila, 18)).Font.Bold = True
        
        ExcelCuadro xlHoja1, 2, nFila - 2, 18, nFila
        
        nInicio = nFila + 1
        
        Do While Not rs.EOF
            nFila = nFila + 1
            i = i + 1
            xlHoja1.Cells(nFila, 2) = i
            xlHoja1.Cells(nFila, 3) = "-"
            xlHoja1.Cells(nFila, 4) = rs!cPersNombre
            xlHoja1.Cells(nFila, 5) = "-"
            xlHoja1.Cells(nFila, 6) = rs!cPersDireccDomicilio
            xlHoja1.Cells(nFila, 7) = rs!cConsDescripcion
            xlHoja1.Cells(nFila, 8) = "-"
            xlHoja1.Cells(nFila, 9) = "-"
            xlHoja1.Cells(nFila, 10) = rs!cPersIDnro
            xlHoja1.Cells(nFila, 11) = "-"
            xlHoja1.Cells(nFila, 12) = "-"
            xlHoja1.Cells(nFila, 13) = "-"
            xlHoja1.Cells(nFila, 14) = Format(rs!nSaldoCredito / 1000, "0.00")
            xlHoja1.Cells(nFila, 15) = Format(rs!nSaldoCaptacion / 1000, "0.00")
            xlHoja1.Cells(nFila, 16) = "-"
            xlHoja1.Cells(nFila, 17) = "-"
            xlHoja1.Cells(nFila, 18) = Format(rs!nSaldoCaptacion / 1000, "0.00")
            rs.MoveNext
        Loop
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 2) = "Total Vinculados por el Articulo 202"
        xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 14)).MergeCells = True
        xlHoja1.Range("O" & nFila & ":O" & nFila).Formula = "=SUM(O" & nInicio & ":O" & nFila - 1 & ")"
        xlHoja1.Range("R" & nFila & ":R" & nFila).Formula = "=SUM(R" & nInicio & ":R" & nFila - 1 & ")"
        
        xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 18)).Font.Bold = True
        
        nFilaTotal2 = nFila
        
        ExcelCuadro xlHoja1, 2, nInicio, 18, nFila, , True
        
    End If
    rs.Close

    
    nFila = nFila + 2
    xlHoja1.Cells(nFila, 2) = "3. Exposición a vinculados"
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 6)).HorizontalAlignment = xlLeft
    
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 18)).Font.Bold = True
    
    nFila = nFila + 2
    xlHoja1.Cells(nFila, 3) = "Artículo 202 de la Ley General"
    xlHoja1.Cells(nFila, 6) = "Artículo 204 de la Ley General"
    xlHoja1.Cells(nFila, 9) = "Total Exposición a Vinculados"
    xlHoja1.Range(xlHoja1.Cells(nFila, 9), xlHoja1.Cells(nFila, 10)).MergeCells = True
    
    nInicio = nFila + 1
    
    nFila = nFila + 1
    xlHoja1.Cells(nFila, 3) = "Total Financiero a vinculados 202 LG(A)"
    If nFilaTotal1 <> 0 Then
    xlHoja1.Range(xlHoja1.Cells(nFila, 4), xlHoja1.Cells(nFila, 4)).Formula = "=+" & xlHoja1.Cells(nFilaTotal1, 14)
    End If
    xlHoja1.Cells(nFila, 6) = "Total Financiamiento a vinculados 204 LG(B)"
    xlHoja1.Range(xlHoja1.Cells(nFila, 7), xlHoja1.Cells(nFila, 7)).Formula = "=+" & xlHoja1.Cells(nFilaTotal2, 15)
    xlHoja1.Cells(nFila, 9) = "Total Financiamiento a vinculados (A+B)"
    If nFilaTotal1 <> 0 Then
    xlHoja1.Range(xlHoja1.Cells(nFila, 10), xlHoja1.Cells(nFila, 10)).Formula = "=+" & xlHoja1.Cells(nFilaTotal1, 14) & "+" & xlHoja1.Cells(nFilaTotal2, 15)
    End If
    
    nFila = nFila + 1
    xlHoja1.Cells(nFila, 3) = "Patrimonio Efectivo(C)"
    xlHoja1.Cells(nFila, 4) = Format(nPatrEfectivo / 1000, "0.00")
    xlHoja1.Cells(nFila, 6) = "Patrimonio Efectivo(C)"
    xlHoja1.Cells(nFila, 7) = Format(nPatrEfectivo / 1000, "0.00")
    xlHoja1.Cells(nFila, 9) = "Patrimonio Efectivo(C)"
    xlHoja1.Cells(nFila, 10) = Format(nPatrEfectivo / 1000, "0.00")
    
    nFila = nFila + 1
    xlHoja1.Cells(nFila, 3) = "Exposición (A)/(C)*100%"
    xlHoja1.Range(xlHoja1.Cells(nFila, 4), xlHoja1.Cells(nFila, 4)).Formula = "=+" & xlHoja1.Cells(nFila - 2, 4) & "/" & xlHoja1.Cells(nFila - 1, 4) & ""
    xlHoja1.Cells(nFila, 6) = "Exposición (B)/(C)*100%"
    xlHoja1.Range(xlHoja1.Cells(nFila, 7), xlHoja1.Cells(nFila, 7)).Formula = "=+" & xlHoja1.Cells(nFila - 2, 7) & "/" & xlHoja1.Cells(nFila - 1, 7) & ""
    xlHoja1.Cells(nFila, 9) = "Exposición (A+B)/(C)*100%"
    xlHoja1.Range(xlHoja1.Cells(nFila, 10), xlHoja1.Cells(nFila, 10)).Formula = "=+" & xlHoja1.Cells(nFila - 2, 10) & "/" & xlHoja1.Cells(nFila - 1, 10) & ""
    
'    xlHoja1.Range(xlHoja1.Cells(nInicio, 2), xlHoja1.Cells(nfila, 10)).Font.Size = 6
    xlHoja1.Range(xlHoja1.Cells(nFila, 2), xlHoja1.Cells(nFila, 10)).NumberFormat = "0.00%"
    
    xlHoja1.Cells.Select
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 9
    xlHoja1.Cells.EntireColumn.AutoFit

    
    ExcelCuadro xlHoja1, 3, nInicio, 4, nFila, , True
    ExcelCuadro xlHoja1, 6, nInicio, 7, nFila, , True
    ExcelCuadro xlHoja1, 9, nInicio, 10, nFila, , True
    
    MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "Aviso"
End Sub


Private Sub txtAnio_KeyPress(KeyAscii As Integer)
Dim oCambio As nTipoCambio
Dim sFecha  As Date
Dim sFecha2  As Date
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    
    sFecha = "01/" & IIf(Len(Trim(cboMes.ListIndex + 1)) = 1, "0" & Trim(Str(cboMes.ListIndex + 1)), Trim(cboMes.ListIndex + 1)) & "/" & Trim(txtAnio.Text)
    sFecha2 = "01/" & IIf(Len(Trim(cboMes.ListIndex + 1)) = 1, "0" & Trim(Str(cboMes.ListIndex + 2)), Trim(cboMes.ListIndex + 1)) & "/" & Trim(txtAnio.Text)
    Set oCambio = New nTipoCambio
    If Len(Trim(cboMes.Text)) > 0 And Val(txtAnio.Text) > 1900 Then
        txtTC.Text = Format(oCambio.EmiteTipoCambio(sFecha, TCFijoDia), "#,##0.0000")
    End If
    txtTC.SetFocus
    txtFecha.Text = CDate(sFecha2) - 1
End If
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFecha) = True Then
        txtTC.SetFocus
    End If
End If
End Sub

Private Sub txtFecha_LostFocus()
Dim oCambio As nTipoCambio
Dim sFecha As String
    If Not IsDate(txtFecha) Then
        Me.txtFecha.SetFocus
        Exit Sub
    End If
    Set oCambio = New nTipoCambio
    txtTC.Text = Format(oCambio.EmiteTipoCambio(Trim(txtFecha.Text), TCFijoDia), "#,##0.0000")
End Sub
Private Sub txtTC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdProcesar.SetFocus
    Else
        KeyAscii = NumerosDecimales(Me.txtTC, KeyAscii, , 3)
    End If
End Sub
