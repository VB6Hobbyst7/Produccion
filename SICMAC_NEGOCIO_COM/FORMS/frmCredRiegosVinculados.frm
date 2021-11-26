VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRiesgosVinculados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vinculados al Titular"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Créditos"
      TabPicture(0)   =   "frmCredRiegosVinculados.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTitular"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "feVinculados"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSalir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdExportar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton cmdExportar 
         Caption         =   "&Exportar Excel"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   8520
         TabIndex        =   2
         Top             =   5520
         Width           =   1215
      End
      Begin SICMACT.FlexEdit feVinculados 
         Height          =   4020
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   7091
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Vinculado-Tipo-Precedente"
         EncabezadosAnchos=   "400-3500-1800-3500"
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         CantEntero      =   10
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblTitular 
         AutoSize        =   -1  'True
         Caption         =   "@titular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmCredRiesgosVinculados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsCtaCod As String
Dim fnMonto As Double 'WIOR 20120928 AGREGO MONTO SUGERIDO
Private Sub cmdExportar_Click()
Dim fs As Scripting.FileSystemObject
Dim xlsAplicacion As Excel.Application
Dim xlsLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lbExisteHoja As Boolean
Dim lsArchivo1 As String
Dim lsArchivo2 As String

Dim lsNomHoja  As String
Dim lsNombreArchivo As String
Dim oPersona As COMDPersona.DCOMPersona
Dim rsPersona As ADODB.Recordset
Dim nRiesgo As Integer
Dim oCredito As COMDCredito.DCOMCredito
Set oCredito = New COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset

Set rsCredito = oCredito.ObtenerInformeRiesgo(fsCtaCod)
If rsCredito.RecordCount > 0 Then
    nRiesgo = CInt(rsCredito!nRiesgo)
End If
    
    Set oPersona = New COMDPersona.DCOMPersona
    Set rsPersona = oPersona.VinculadosACuenta(Trim(fsCtaCod))

    If rsPersona.EOF And rsPersona.BOF Then
        MsgBox "No existen Datos para este Exportar.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    lsArchivo1 = "VinculadosTitular"
    lsNomHoja = "Vinculados"
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application

    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo1 & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo1 & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    lsArchivo2 = lsArchivo1 & "_" & gsCodUser & "_" & Format$(gdFecSis, "yyyymmdd") & "_" & Format$(Time(), "HHMMSS")

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

    Call ExportarExcel(rsPersona, xlHoja1, nRiesgo)

    xlHoja1.SaveAs App.path & "\Spooler\" & lsArchivo2 & ".xls"
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

End Sub

Public Sub ExportarExcel(ByRef pR As ADODB.Recordset, ByRef xlHoja1 As Excel.Worksheet, ByVal pnRiesgo As Integer)
Dim oPersona As COMDPersona.DCOMPersona
Set oPersona = New COMDPersona.DCOMPersona
Dim rsPersona As ADODB.Recordset
'WIOR 20130903 ********************************
Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
Dim nTC As Double
Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
nTC = oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoMes)
'WIOR FIN *************************************

Dim i As Integer
If pR.RecordCount > 0 Then
    For i = 0 To pR.RecordCount - 1
        If Trim(pR!Tipo) = "TITULAR" Then
            xlHoja1.Cells(7, 4) = Trim(pR!cPersCod)
            xlHoja1.Cells(8, 4) = PstaNombre(Trim(pR!Nombre), True)
            Set rsPersona = oPersona.DeterminarSaldoPorPersona(Trim(pR!cPersCod), pnRiesgo, gdFecSis)
            If rsPersona.RecordCount > 0 Then
                'WIOR 20121017 ********************************************************
                    'EN CASO DE EXISTIR AMPLIACIONES
                    Dim oAmpliadoVal  As COMDCredito.DCOMAmpliacion
                    Dim bAmpliadoVal As Boolean
                    Dim rsAmpliadosVal As ADODB.Recordset
                    Dim nMontoAmpliado As Double
                    Dim IAmp As Integer
                    Dim oDCreditoAmp As COMDCredito.DCOMCredito
                    Dim rsDCreditoAmp As ADODB.Recordset
                    Set oAmpliadoVal = New COMDCredito.DCOMAmpliacion
                    bAmpliadoVal = oAmpliadoVal.ValidaCreditoaAmpliar(fsCtaCod)
                    nMontoAmpliado = 0
                    If bAmpliadoVal Then
                        Set rsAmpliadosVal = oAmpliadoVal.ListaCreditosBycCtaCodNew(Trim(fsCtaCod))
                        If rsAmpliadosVal.RecordCount > 0 Then
                            If Not (rsAmpliadosVal.BOF And rsAmpliadosVal.EOF) Then
                            Set oDCreditoAmp = New COMDCredito.DCOMCredito
                                For IAmp = 0 To rsAmpliadosVal.RecordCount - 1
                                    Set rsDCreditoAmp = oDCreditoAmp.RecuperaProducto(Trim(rsAmpliadosVal!cCtaCodAmp))
                                    If rsDCreditoAmp.RecordCount > 0 Then
                                        If Not (rsDCreditoAmp.BOF And rsDCreditoAmp.EOF) Then
                                            nMontoAmpliado = nMontoAmpliado + CDbl(rsDCreditoAmp!nSaldo) * CDbl(IIf(Mid(Trim(Trim(rsAmpliadosVal!cCtaCodAmp)), 9, 1) = "1", 1, nTC)) 'WIOR 20130903 AGREGO  * CDbl(IIf(Mid(Trim(Trim(rsAmpliadosVal!cCtaCodAmp)), 9, 1) = "1", 1, nTC))
                                            xlHoja1.Cells(8, 6) = Trim(rsAmpliadosVal!cCtaCodAmp) + " - " + xlHoja1.Cells(9, 4)
                                        End If
                                    End If
                                    Set rsDCreditoAmp = Nothing
                                    rsAmpliadosVal.MoveNext
                                Next IAmp
                            End If
                        End If
                        xlHoja1.Cells(8, 5) = "Creditos que Fueron Ampliados:"
                        xlHoja1.Cells(8, 5).Font.Bold = True
                        xlHoja1.Cells(8, 6).Font.Bold = True
                        xlHoja1.Cells(9, 5) = "Monto Anterior Credito:"
                        xlHoja1.Cells(9, 5).Font.Bold = True
                        xlHoja1.Cells(9, 6) = nMontoAmpliado
                        xlHoja1.Cells(9, 6).Font.Bold = True
                        xlHoja1.Cells(8, 6) = "'" & Mid(xlHoja1.Cells(8, 6), 1, Len(xlHoja1.Cells(8, 6)) - 3)
                    
                    End If
                'WIOR FIN ***************************************************************
                xlHoja1.Cells(9, 4) = CDbl(rsPersona!TotSaldoFinal) + fnMonto - nMontoAmpliado 'WIOR 20120928 AGREGO MONTO SUGERIDO
            Else
                xlHoja1.Cells(9, 4) = CDbl("0.00") + fnMonto 'WIOR 20120928 AGREGO MONTO SUGERIDO
            End If
        Else
            xlHoja1.Cells(13 + i, 2) = i
            xlHoja1.Cells(13 + i, 3) = pR!cPersCodVin
            xlHoja1.Cells(13 + i, 4) = PstaNombre(pR!Vinculado, True)
            xlHoja1.Cells(13 + i, 5) = pR!Tipo
            
            Set rsPersona = oPersona.DeterminarSaldoPorPersona(Trim(pR!cPersCodVin), pnRiesgo, gdFecSis)
            If rsPersona.RecordCount > 0 Then
                xlHoja1.Cells(13 + i, 6) = Trim(rsPersona!TotSaldoFinal)
            Else
                xlHoja1.Cells(13 + i, 6) = "0.00"
            End If
            
            xlHoja1.Cells(13 + i, 7) = PstaNombre(Trim(pR!Nombre), True)
            xlHoja1.Range(xlHoja1.Cells(13 + i, 3), xlHoja1.Cells(13 + i, 7)).Borders.LineStyle = 1
        End If
        pR.MoveNext
    Next i
    xlHoja1.Cells(13 + i, 6) = "=SUM(F13:F" & CStr(12 + i) & ")+D9"
    xlHoja1.Cells(13 + i, 6).Borders.LineStyle = 1
    xlHoja1.Cells(13 + i, 6).Font.Bold = True
    xlHoja1.Cells(13 + i, 6).Interior.Color = RGB(255, 255, 0)
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call CentraForm(Me)
Me.Icon = LoadPicture(App.path & gsRutaIcono)
lblTitular.Caption = ""
Call LlenarGrilla
End Sub
Public Sub Inicio(ByVal psCtaCod As String, ByVal pnMonto As Double)
fsCtaCod = psCtaCod
fnMonto = pnMonto 'WIOR 20120928 AGREGO MONTO SUGERIDO
Me.Show 1
End Sub

Public Sub LlenarGrilla()
Dim oPersona As COMDPersona.DCOMPersona
Dim rsPersona As ADODB.Recordset
Dim i As Integer

Set oPersona = New COMDPersona.DCOMPersona

Set rsPersona = oPersona.VinculadosACuenta(Trim(fsCtaCod))
Call LimpiaFlex(feVinculados)

If rsPersona.RecordCount > 0 Then
    For i = 0 To rsPersona.RecordCount - 1
        If Trim(rsPersona!Tipo) = "TITULAR" Then
            lblTitular.Caption = PstaNombre(Trim(rsPersona!Nombre), True)
        Else
            feVinculados.AdicionaFila
            feVinculados.TextMatrix(i, 0) = i
            feVinculados.TextMatrix(i, 1) = PstaNombre(rsPersona!Vinculado, True)
            feVinculados.TextMatrix(i, 2) = rsPersona!Tipo
             feVinculados.TextMatrix(i, 3) = PstaNombre(Trim(rsPersona!Nombre), True)
        End If
        rsPersona.MoveNext
    Next i
End If
End Sub

