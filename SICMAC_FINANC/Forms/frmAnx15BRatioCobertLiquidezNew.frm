VERSION 5.00
Begin VB.Form frmAnx15BRatioCobertLiquidezNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anexo 15B: Ratio de Cobertura de Liquidez"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboBCRP 
      Height          =   315
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Frame fraRango 
      Caption         =   "Balance Diario"
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
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   3240
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   230
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtAnio 
         Height          =   330
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblGuion 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1900
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtFechaReporte15B 
      Height          =   285
      Left            =   5400
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Reporte BCRP N°1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   400
      Width           =   1095
   End
End
Attribute VB_Name = "frmAnx15BRatioCobertLiquidezNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*** Nombre : frmAnx15BRatioCobertLiquidezNew
'*** Descripción : Formulario para generar el Anexo 15B - Ratio de Cobertura de Líquidez.
'*** Creación : NAGL el 20170421
'********************************************************************************
Dim oDAnexos As New DAnexoRiesgos

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAnio.SetFocus
    End If
End Sub
Private Sub cboBCRP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub
Private Sub txtAnio_GotFocus()
fEnfoque txtAnio
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
         cmdGenerar.SetFocus
    End If
End Sub
Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Function ValidaDatos(pdFecha As Date) As Boolean
If (Month(pdFecha) > 12) Or (Month(pdFecha) <= 2) Then
        If txtAnio.Text = "" Or IIf(CInt(CboMes.ListIndex) = 0, txtAnio.Text > Year(pdFecha) Or txtAnio.Text <= Year(DateAdd("yyyy", -1, pdFecha)), txtAnio.Text >= Year(pdFecha) Or txtAnio.Text < Year(DateAdd("yyyy", -1, pdFecha))) Then
            MsgBox "Debe ingresar el año correspondiente", vbInformation, "Aviso"
            txtAnio.SetFocus
            Exit Function
        End If
        
        If (CInt(CboMes.ListIndex) + 1) >= Month(pdFecha) And (CInt(CboMes.ListIndex) + 1) < 11 Then
                MsgBox "El mes ingresado es incorrecto", vbInformation, "Aviso"
                CboMes.SetFocus
                Exit Function
        End If
            If Day(pdFecha) < 15 Then
                If (CInt(CboMes.ListIndex) + 1) = Month(DateAdd("m", -1, pdFecha)) Or (CInt(CboMes.ListIndex) + 1) < Month(DateAdd("m", -2, pdFecha)) Then
                        If MsgBox("Desea actualizar las cuentas con el Balance del Mes Ingresado?", vbInformation + vbYesNo, "Atención") = vbNo Then
                            CboMes.SetFocus
                            Exit Function
                        End If
                End If
            Else
                If (CInt(CboMes.ListIndex) + 1) < Month(DateAdd("m", -1, pdFecha)) Then
                    If MsgBox("Desea actualizar las cuentas con el Balance del Mes Ingresado?", vbInformation + vbYesNo, "Atención") = vbNo Then
                        CboMes.SetFocus
                        Exit Function
                    End If
                End If
             End If
       
Else
        If txtAnio.Text = "" Or txtAnio.Text > Year(pdFecha) Or txtAnio.Text < Year(pdFecha) Then
            MsgBox "Debe ingresar el año correspondiente", vbInformation, "Aviso"
            txtAnio.SetFocus
            Exit Function
        End If
        
        If (CInt(CboMes.ListIndex) + 1) >= Month(pdFecha) Then
                MsgBox "El mes ingresado es incorrecto", vbInformation, "Aviso"
                CboMes.SetFocus
                Exit Function
        Else
            If Day(pdFecha) >= 15 Then
                If (CInt(CboMes.ListIndex) + 1) < Month(DateAdd("m", -1, pdFecha)) Then
                    If MsgBox("Desea actualizar las cuentas con el Balance del Mes Ingresado?", vbInformation + vbYesNo, "Atención") = vbNo Then
                        CboMes.SetFocus
                        Exit Function
                    End If
                    'MsgBox "El mes ingresado es incorrecto", vbInformation, "Aviso"
                    'cboMes.SetFocus
                    'Exit Function
                End If
            Else
                If (CInt(CboMes.ListIndex) + 1) < Month(DateAdd("m", -2, pdFecha)) Or (CInt(CboMes.ListIndex) + 1) = Month(DateAdd("m", -1, pdFecha)) Then
                    If MsgBox("Desea actualizar las cuentas con el Balance del Mes Ingresado?", vbInformation + vbYesNo, "Atención") = vbNo Then
                        CboMes.SetFocus
                        Exit Function
                    End If
                    'MsgBox "El mes ingresado es incorrecto", vbInformation, "Aviso"
                    'cboMes.SetFocus
                    'Exit Function
                End If
            End If
        End If
End If
If Day(pdFecha) < 15 Then
    If CInt(cboBCRP.ListIndex) + 1 <> "2" Then
       If MsgBox("Desea utilizar el Reporte BCRP N°1?", vbInformation + vbYesNo, "Atención") = vbNo Then
          cboBCRP.SetFocus
          Exit Function
        End If
    End If
Else
    If CInt(cboBCRP.ListIndex) + 1 <> "1" Then
       If MsgBox("Seguro de que no desea utilizar el Reporte BCRP N°1?", vbInformation + vbYesNo, "Atención") = vbNo Then
          cboBCRP.SetFocus
          Exit Function
       End If
    End If
End If
    ValidaDatos = True
End Function

Private Function ValRegValorizacionDiaria(pdFecha As Date) As Boolean
Dim DAnxVal As New DAnexoRiesgos
Dim psRegistro As String
 psRegistro = DAnxVal.ObtenerRegValorizacionDiaria(pdFecha)
    If psRegistro = "0" Then
            If MsgBox("La Valorización Diaria no ha sido ingresada, Desea continuar?", vbInformation + vbYesNo, "Atención") = vbNo Then
                CboMes.SetFocus
                Exit Function
            End If
    End If
ValRegValorizacionDiaria = True
End Function
 
Public Sub ImprimeAnexo15BRatioCoberLiqudz(psOpeCod As String, psMoneda As String, pdFecha As Date, Optional ByVal pnTipoCambio As Currency = 1)  'NAGL 20170415
Dim rs As New ADODB.Recordset
Dim oGen As New DGeneral
Dim psMesBalanceDiario As String
Dim psAnioBalanceDiario As String
Dim pdFechaFinMesAnt As String, psRepBCRP As String
    
    If Day(pdFecha) >= 15 Then
        pdFechaFinMesAnt = DateAdd("d", -Day(pdFecha), pdFecha)
        psRepBCRP = "1"
    Else
        pdFechaFinMesAnt = DateAdd("d", -Day(DateAdd("m", -1, pdFecha)), DateAdd("m", -1, pdFecha))
        psRepBCRP = "2"
    End If
    
    psMesBalanceDiario = Month(pdFechaFinMesAnt)
    psAnioBalanceDiario = Year(pdFechaFinMesAnt)
    
    Set rs = oGen.GetConstante(1010)
    While Not rs.EOF
        CboMes.AddItem rs.Fields(0) & space(50) & rs.Fields(1)
        rs.MoveNext
    Wend
    cboBCRP.AddItem "SI" & space(50) & "1"
    cboBCRP.AddItem "NO" & space(50) & "2"
    txtFechaReporte15B.Text = Format(pdFecha, "dd/MM/YYYY")
    txtAnio.Text = psAnioBalanceDiario
    CboMes.ListIndex = CInt(psMesBalanceDiario) - 1
    cboBCRP.ListIndex = CInt(psRepBCRP) - 1
    CentraForm Me
    Me.Show 1
End Sub

Private Function ActivaConsolidaIntDevCreditosFinMes(pdFecha As Date, lsPermitGenfromBitac As String) As Boolean
Dim psConsolida As String
    psConsolida = 0
    If lsPermitGenfromBitac = "NO" Then
        If pdFecha = DateAdd("d", -Day(gdFecSis), gdFecSis) Then
            If oDAnexos.ObtieneDisponibilidadConsolCalif(pdFecha) = True Then
                  psConsolida = 1
                  Call oDAnexos.EjecutaConsolidaSaldoCapCredFinMes(pdFecha) '***NAGL 2091213 Según RFC1912050003
            Else
               If MsgBox("La calificación general al " & pdFecha & " aún no se encuentra disponible. Desea Continuar?", vbInformation + vbYesNo, "Atención") = vbYes Then
                  psConsolida = 0
               Else
                  ActivaConsolidaIntDevCreditosFinMes = False
                  Exit Function
               End If
            End If
            Call oDAnexos.EjecutaConsolidaIntDevCredFinMes(pdFecha, psConsolida)
        End If
    End If
    ActivaConsolidaIntDevCreditosFinMes = True
End Function 'NAGL 20191116 Según Anx03-ERS006-2019

Private Sub cmdGenerar_Click()
Dim psMoneda As String, psMesBalanceDiario As String, psAnioBalanceDiario As String
Dim pdFecha As Date
Dim psRepBCRP As String
Dim pdFechaSist As Date 'NAGL 20170904
Dim lsPermitGenfromBitac As String 'NAGL 20170904
Dim oDbalanceCont As New DbalanceCont  'NAGL 20170904

    psMoneda = "1"
    pdFecha = txtFechaReporte15B.Text
    psAnioBalanceDiario = txtAnio.Text
    psRepBCRP = CInt(cboBCRP.ListIndex) + 1
    pdFechaSist = Format(gdFecSis, "dd/mm/yyyy") 'NAGL 20170904+
    
    If ValRegValorizacionDiaria(pdFecha) Then
        If (CInt(CboMes.ListIndex) + 1) <= 9 Then
            psMesBalanceDiario = "0" & CStr(CInt(CboMes.ListIndex) + 1)
        Else
            psMesBalanceDiario = CStr(CInt(CboMes.ListIndex) + 1)
        End If
        If ValidaDatos(pdFecha) Then 'Valida Datos con respecto a la fecha ingresada
           lsPermitGenfromBitac = oDbalanceCont.ObtenerPermiteGenerarfromBitacora15A_15B(pdFecha, pdFechaSist) 'NAGL 20170904
           If ActivaConsolidaIntDevCreditosFinMes(pdFecha, lsPermitGenfromBitac) = True Then
                '***NAGL 20170904
                If lsPermitGenfromBitac = "NO" Then
                    Call ReporteAnexo15B(psMoneda, pdFecha, psMesBalanceDiario, psAnioBalanceDiario, psRepBCRP)
                Else
                    Call ReporteAnexo15BfromBitacora(psMoneda, pdFecha, pdFechaSist, psMesBalanceDiario, psAnioBalanceDiario, psRepBCRP)
                End If
                '***FIN NAGL 20170904
           End If 'NAGL 20191107 Según Anx03_ERS006-2019
        End If
    End If
End Sub

Public Sub ReporteAnexo15B(psMoneda As String, pdFecha As Date, psMesBalanceDiario As String, psAnioBalanceDiario As String, psRepBCRP As String) 'NAGL 20170415
    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja  As String
    Dim lsArchivo As String
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim TituloProgress As String 'NAGL 20170407
    Dim MensajeProgress As String 'NAGL 20170407
    Dim oBarra As clsProgressBar 'NAGL 20170407
    Dim nprogress As Integer 'NAGL 20170407
    

    Dim oDbalanceCont As dBalanceCont
    Dim nSaldoDiario1 As Currency
    Dim nSaldoDiario2 As Currency
    Dim pdFechaFinDeMes As Date
    Dim pdFechaFinDeMesMA As Date
    Dim dFechaAnte As Date
    Dim ldFechaPro As Date
    Dim pdFechaBalanceDiario As Date
    Dim nDia As Integer
    Dim oCambio As nTipoCambio
    Dim lnTipoCambioFC As Currency, pnTipoCambio As Currency 'NAGL20170420
    Dim lnTipoCambioProceso As Currency
    'Dim lnTipoCambioBalance As Currency
    Dim lnTipoCambioBalanceAnterior As Currency
    Dim nTipoCambioAn As Currency
    Dim loRs As ADODB.Recordset
    Dim lnSubastasMN As Currency
    Dim lnSubastasME As Currency
    Dim lnOtrosDepMND1y2_C0 As Currency
    Dim lnOtrosDepMND4aM_C0 As Currency
    Dim lnOtrosDepMND1y2_C1 As Currency
    Dim lnOtrosDepMED1y2_C0 As Currency
    Dim lnOtrosDepMED4aM_C0 As Currency
    Dim lnOtrosDepMED1y2_C1 As Currency
    Dim lnSubastasMND1y2_C1 As Currency
    Dim lnSubastasMND1y2_C0 As Currency
    Dim lnSubastasMND4aM_C0 As Currency
    Dim lnSubastasMED1y2_C1 As Currency
    Dim lnSubastasMED1y2_C0 As Currency
    Dim lnSubastasMED4aM_C0 As Currency
    
    Dim lnSubastasMED3o3_C0 As Currency
    Dim lnSubastasMND3o3_C0 As Currency
    Dim lnSubastasMED3o3_C1 As Currency
    Dim lnSubastasMND3o3_C1 As Currency
    
    Dim lnOtrosDepMND3o3_C0 As Currency
    Dim lnOtrosDepMED3o3_C0 As Currency
    Dim lnOtrosDepMND3o3_C1 As Currency
    Dim lnOtrosDepMED3o3_C1 As Currency
    
    Dim lnSubastasMND4aM_C1 As Currency
    Dim lnSubastasMED4aM_C1 As Currency
    
    Dim nTotalObligSugEncajMN As Currency
    Dim nTotalTasaBaseEncajMN As Currency
    Dim nTotalObligSugEncajME As Currency
    Dim nTotalTasaBaseEncajME  As Currency
    Dim lnTotalObligacionesAlDiaMN As Currency
    Dim lnTotalObligacionesAlDiaME  As Currency
    Dim nTotalTasaBaseEncajMNDiario As Currency
    Dim nTotalTasaBaseEncajMEDiario  As Currency
    Dim nTotalTasaBaseEncajMN_DADiario As Currency
    Dim nTotalTasaBaseEncajME_DADiario As Currency
    Dim nTotalTasaBaseEncajMN_DA  As Currency
    Dim nTotalObligSugEncajMN_DA As Currency
    Dim nTotalTasaBaseEncajME_DA As Currency
    Dim nTotalObligSugEncajME_DA As Currency
    Dim ix As Integer, rx As Integer
    Dim nSubValor1 As Currency
    Dim nSubValor2 As Currency
    Dim DAnxVal As New DAnexoRiesgos 'NAGL ERS079-2017 20180128
    
    'VAPA 20171101 ***********************************
    Dim lnTotal1MN As Double
    Dim lnTotal1ME As Double
    Dim lnTotal2MN As Double
    Dim lnTotal2ME As Double
    Dim lnTotal3MN As Double
    Dim lnTotal3ME As Double
    Dim lnRatioCoberturaLiquidezMN As Double
    Dim lnRatioCoberturaLiquidezME As Double
    Dim lnAuxMN As Double
    Dim lnAuxME As Double
    Dim pdFechaAlerta As Date
    'VAPA END ****************************************
    Dim pnFila As Integer 'NAGL 20190430
    Dim iFilaAnx As Integer 'NAGL 20190430
    Dim rsCta As New ADODB.Recordset 'NAGL 20190430
    Dim rs As New ADODB.Recordset 'NAGL 20190430
    Dim nSaldoDiarioParam As Currency 'NAGL 20190509
    Dim iFilaAnxMain As Integer, iFilaCtaLiqu As Integer, iFilaAnxParam As Integer 'NAGL 20190509
    Dim rsCtasAdic As New ADODB.Recordset 'NAGL 20190509
    
On Error GoTo GeneraExcelErr

    Set oBarra = New clsProgressBar
    Unload Me 'NAGL 2070422
    oBarra.ShowForm frmReportes
    oBarra.Max = 10
    nprogress = 0
    oBarra.Progress nprogress, "ANEXO 15B: Ratio de Cobertura de Liquidez", "GENERANDO EL ARCHIVO", "", vbBlue
    TituloProgress = "ANEXO 15B: Ratio de Cobertura de Liquidez"
    MensajeProgress = "GENERANDO EL ARCHIVO"
    'NAGL20170407
    
    pdFechaFinDeMes = DateAdd("d", -Day(DateAdd("m", 1, pdFecha)), DateAdd("m", 1, pdFecha))
    pdFechaFinDeMesMA = DateAdd("d", -Day(pdFechaFinDeMes), pdFechaFinDeMes)
    nDia = Day(pdFecha)
    Set oDbalanceCont = New dBalanceCont
    Set oCambio = New nTipoCambio
    pdFechaAlerta = pdFecha 'vapa20171117
    
    If CInt(psMesBalanceDiario) < 9 Then
        pdFechaBalanceDiario = CDate("01" & "/" & "0" & CStr(CInt(psMesBalanceDiario) + 1) & "/" & psAnioBalanceDiario)
    ElseIf CInt(psMesBalanceDiario) = 12 Then
        pdFechaBalanceDiario = CDate("01" & "/" & "01" & "/" & CStr(CInt(psAnioBalanceDiario) + 1))
    Else
        pdFechaBalanceDiario = CDate("01" & "/" & CStr(CInt(psMesBalanceDiario) + 1) & "/" & psAnioBalanceDiario)
    End If
    
    lnTipoCambioBalanceAnterior = Format(oCambio.EmiteTipoCambio(pdFechaBalanceDiario, TCFijoDia), "#,##0.0000")  ''NAGL ERS079-2016 20170407
    
    If Month(pdFecha) = Month(DateAdd("d", 1, pdFecha)) Then
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(pdFecha, TCFijoDia), "#,##0.0000")
    Else
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(DateAdd("d", -1, pdFecha), TCFijoDia), "#,##0.0000")
    End If
    pnTipoCambio = lnTipoCambioFC
    
    oBarra.Progress 1, TituloProgress, MensajeProgress, "", vbBlue
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ANEXO_15B"
    'Primera Hoja ******************************************************
    lsNomHoja = "Anx15B"
    '*******************************************************************
    lsArchivo1 = "\spooler\ANEXO_15B_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
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
    xlHoja1.Cells(4, 2) = "AL " & Format(pdFecha, "YYYY/MM/DD")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(1, pdFecha, "1", 0, 1, "100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(1, pdFecha, "2", 0, 1, "100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(1, pdFecha, "3", 0, 1, "100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(1, pdFecha, "4", 0, 1, "100", "B1")
    
    'NAGL ERS079-2016 20170407 (Saldo en Caja)
    'oDbalanceCont.ObtenerCtaContSaldoDiario("1111", pdFecha)
    nSaldoDiario1 = ObtieneEfectivoDiaxRptConcentracionGastos(pdFecha, 1) + oDAnexos.ObtieneSaldoEfectTransitoTotal(pdFecha, 1)  'NAGL 20181002 Agregó el Efectivo en Tránsito
    'oDbalanceCont.ObtenerCtaContSaldoDiario("1121", pdFecha)
    nSaldoDiario2 = ObtieneEfectivoDiaxRptConcentracionGastos(pdFecha, 2) + oDAnexos.ObtieneSaldoEfectTransitoTotal(pdFecha, 2) 'NAGL 20181002 Agregó el Efectivo en Tránsito
    xlHoja1.Cells(10, 3) = nSaldoDiario1
    xlHoja1.Cells(10, 4) = nSaldoDiario2
    
    Call oDbalanceCont.InsertaDetallaReporte15A(2, pdFecha, "1", nSaldoDiario1, 1, "200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(2, pdFecha, "2", nSaldoDiario2, 1, "200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(2, pdFecha, "3", xlHoja1.Cells(10, 6), 1, "200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(2, pdFecha, "4", xlHoja1.Cells(10, 7), 1, "200", "B1")
    
    xlHoja1.Range(xlHoja1.Cells(10, 3), xlHoja1.Cells(10, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnTotal1MN = xlHoja1.Cells(10, 3) 'VAPA20171101
    lnTotal1ME = xlHoja1.Cells(10, 4) 'VAPA20171101
    
    oBarra.Progress 2, TituloProgress, MensajeProgress, "", vbBlue
    
    'Inicio TOSE
        nTotalObligSugEncajMN = 0
        nTotalTasaBaseEncajMN = 0
        nTotalObligSugEncajME = 0
        nTotalTasaBaseEncajME = 0
        lnTotalObligacionesAlDiaMN = 0
        lnTotalObligacionesAlDiaME = 0
        nTotalTasaBaseEncajMNDiario = 0
        nTotalTasaBaseEncajMEDiario = 0
        nTotalTasaBaseEncajMN_DADiario = 0
        nTotalTasaBaseEncajME_DADiario = 0
        nTotalTasaBaseEncajMN_DA = 0
        nTotalTasaBaseEncajMN = 0
        nTotalTasaBaseEncajME_DA = 0
        nTotalTasaBaseEncajME = 0
        ldFechaPro = DateAdd("d", -Day(pdFecha), pdFecha)
        ldFechaPro = DateAdd("d", -Day(ldFechaPro), ldFechaPro)
        
    If psRepBCRP = "1" Then
    
        For ix = 1 To Day(DateAdd("d", -Day(pdFecha), pdFecha))
            ldFechaPro = DateAdd("d", 1, ldFechaPro)
                If Month(ldFechaPro) = Month(DateAdd("d", 1, ldFechaPro)) Then
                    lnTipoCambioProceso = Format(oCambio.EmiteTipoCambio(ldFechaPro, TCFijoDia), "#,##0.0000")
                Else
                    lnTipoCambioProceso = Format(oCambio.EmiteTipoCambio(DateAdd("d", 1, ldFechaPro), TCFijoDia), "#,##0.0000")
                End If
                
                'SOLES
                nTotalObligSugEncajMN_DA = oDbalanceCont.SaldoCtas(1, "761201", ldFechaPro, pdFechaFinDeMes, lnTipoCambioProceso, lnTipoCambioProceso) 'Obligaciones Inmediatas
                nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "232") - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "232") 'Ahorros
                nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "233") + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "234")
                nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "233") - oDbalanceCont.SaldoCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "233") 'Depositos a plazo fijo
                'nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA - (oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "232") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "233") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "234"))
                lnTotalObligacionesAlDiaMN = lnTotalObligacionesAlDiaMN + nTotalObligSugEncajMN_DA '*************NAGL ERS079-2016 20170407
                
                'DOLARES
                nTotalObligSugEncajME_DA = oDbalanceCont.SaldoCtas(1, "762201", ldFechaPro, pdFechaFinDeMes, lnTipoCambioProceso, lnTipoCambioProceso) 'Obligaciones Inmediatas
                nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "232") - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "232") 'Ahorros
                nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "233") + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "234")
                nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "233") - oDbalanceCont.SaldoCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "233") 'Depositos a plazo fijo
                'nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA - (oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "232") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "233") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "234"))
                lnTotalObligacionesAlDiaME = lnTotalObligacionesAlDiaME + nTotalObligSugEncajME_DA '*************NAGL ERS079-2016 20170407
        Next ix
    Else
            'Soles
            'lnTotalObligacionesAlDiaMN = oDbalanceCont.ObtenerPosicionAcumMesAnterior(pdFecha, "1") / 0.01 JIPR20200610
            lnTotalObligacionesAlDiaMN = oDbalanceCont.ObtenerPosicionAcumMesAnterior(pdFecha, "1") / (oDbalanceCont.ObtenerParamEncDiarioxCodigo("34") / 100)
            'Dolares
            'lnTotalObligacionesAlDiaME = oDbalanceCont.ObtenerPosicionAcumMesAnterior(pdFecha, "2") / 0.03 JIPR20200610
            lnTotalObligacionesAlDiaME = oDbalanceCont.ObtenerPosicionAcumMesAnterior(pdFecha, "2") / (oDbalanceCont.ObtenerParamEncDiarioxCodigo("35") / 100)
    
    End If ' NAGL 20170422
    
    oBarra.Progress 3, TituloProgress, MensajeProgress, "", vbBlue
    
    CargaValidacionCtaContIntDeveng15B xlHoja1.Application, pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio, psMesBalanceDiario, psAnioBalanceDiario  '***NAGL ERS 079-2017 20180123
    
    nSaldoDiario1 = oDbalanceCont.ObtenerDispobiblesenSFN("1090100822183", "1112", pdFecha, 1) + DAnxVal.ObtenerSaldEstadistAnx15Ay15B("111802", "1", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio) 'NAGL ERS079-2017 20180128 'ObtenerCtaContSaldoBalanceDiario("111802", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
    nSaldoDiario2 = oDbalanceCont.ObtenerDispobiblesenSFN("1090100822183", "1122", pdFecha, 1) + Round(DAnxVal.ObtenerSaldEstadistAnx15Ay15B("112802", "2", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio), 2) 'NAGL ERS079-2017 20180128 'Round(ObtenerCtaContSaldoBalanceDiario("112802", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) / lnTipoCambioBalanceAnterior, 2)
    
    nSaldoDiario1 = IIf(nSaldoDiario1 > 0, nSaldoDiario1, 0)
    nSaldoDiario2 = IIf(nSaldoDiario2 > 0, nSaldoDiario2, 0)
    xlHoja1.Cells(11, 3) = nSaldoDiario1
    xlHoja1.Cells(11, 4) = nSaldoDiario2
    
    Call oDbalanceCont.InsertaDetallaReporte15A(3, pdFecha, "1", nSaldoDiario1, 1, "300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(3, pdFecha, "2", nSaldoDiario2, 1, "300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(3, pdFecha, "3", xlHoja1.Cells(11, 6), 1, "300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(3, pdFecha, "4", xlHoja1.Cells(11, 7), 1, "300", "B1")
    
    lnTotal1MN = lnTotal1MN + xlHoja1.Cells(11, 3) 'VAPA20171101
    lnTotal1ME = lnTotal1ME + xlHoja1.Cells(11, 4) 'VAPA20171101
    
    '*********NAGL ERS079-2016 20170407 Ajuste por Encaje Exigible
    nSaldoDiario1 = (lnTotalObligacionesAlDiaMN / Day(DateAdd("d", -Day(pdFecha), pdFecha))) * (oDbalanceCont.ObtenerParamEncDiarioxCodigo("32") / 100) * 0.25 * -1
    nSaldoDiario2 = (lnTotalObligacionesAlDiaME / Day(DateAdd("d", -Day(pdFecha), pdFecha))) * (oDbalanceCont.ObtenerParamEncDiarioxCodigo("33") / 100) * -1
    
    xlHoja1.Cells(12, 3) = nSaldoDiario1
    xlHoja1.Cells(12, 4) = nSaldoDiario2
    
    lnTotal1MN = lnTotal1MN + xlHoja1.Cells(12, 3) 'VAPA20171101
    lnTotal1ME = lnTotal1ME + xlHoja1.Cells(12, 4) 'VAPA20171101
    '***************NAGL ERS079-2016 20170407
    
    Call oDbalanceCont.InsertaDetallaReporte15A(300, pdFecha, "1", nSaldoDiario1, 1, "301", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(300, pdFecha, "2", nSaldoDiario2, 1, "301", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(300, pdFecha, "3", xlHoja1.Cells(12, 6), 1, "301", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(300, pdFecha, "4", xlHoja1.Cells(12, 7), 1, "301", "B1")
    
    'Encaje Liberado por los Flujos Salientes
    'xlHoja1.Cells(13, 3) --> Cálcula en Plantilla, y es considerado al final para la Bitácora
    
    '***************NAGL ERS079-2016 20170407 VALORES REPRESENTATIVOS
    nSaldoDiario1 = oDbalanceCont.ObtenerSumaValorRepresentativos(pdFecha, "C[BD]")
    nSaldoDiario2 = oDbalanceCont.ObtenerSumaValorRepresentativos(pdFecha, "LT")
    
    xlHoja1.Cells(14, 3) = nSaldoDiario1
    xlHoja1.Cells(15, 3) = nSaldoDiario2
    lnTotal1MN = lnTotal1MN + xlHoja1.Cells(14, 3) 'VAPA20171101
    lnTotal1MN = lnTotal1MN + xlHoja1.Cells(15, 3) 'VAPA20171101
    
    '***************NAGL ERS079-2016 20170407
    xlHoja1.Range(xlHoja1.Cells(11, 3), xlHoja1.Cells(11, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(12, 3), xlHoja1.Cells(12, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(14, 3), xlHoja1.Cells(14, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(15, 3), xlHoja1.Cells(15, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnTotal1ME = lnTotal1ME + xlHoja1.Cells(14, 4) 'VAPA20171101
    lnTotal1ME = lnTotal1ME + xlHoja1.Cells(15, 4) 'VAPA20171101
    
    
    '*********ANPS Valores representativos de deuda del BCRP y Gobierno Central recibidos en operaciones de reporte (22)

    nSaldoDiario1 = oDbalanceCont.ObtenerSumaValorBCRP(pdFecha, "1", 0)
    nSaldoDiario2 = oDbalanceCont.ObtenerSumaValorBCRP(pdFecha, "2", 0)
    
    xlHoja1.Cells(20, 3) = nSaldoDiario1 'SOLES
    xlHoja1.Cells(20, 4) = nSaldoDiario2 'DOLARES
    
    
    'Valores representativos de deuda emitidos por el BCRP (5)
    Call oDbalanceCont.InsertaDetallaReporte15A(5, pdFecha, "1", xlHoja1.Cells(14, 3), 1, "500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(5, pdFecha, "2", xlHoja1.Cells(14, 4), 1, "500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(5, pdFecha, "3", xlHoja1.Cells(14, 6), 1, "500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(5, pdFecha, "4", xlHoja1.Cells(14, 7), 1, "500", "B1")
    'Valores representativos de deuda emitidos por el Gobierno Central (5)
    Call oDbalanceCont.InsertaDetallaReporte15A(6, pdFecha, "1", xlHoja1.Cells(15, 3), 1, "600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(6, pdFecha, "2", xlHoja1.Cells(15, 4), 1, "600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(6, pdFecha, "3", xlHoja1.Cells(15, 6), 1, "600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(6, pdFecha, "4", xlHoja1.Cells(15, 7), 1, "600", "B1")
    'Valores representativos de deuda emitidos por Gobiernos del Exterior  (6)
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "1", xlHoja1.Cells(16, 3), 1, "700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "2", xlHoja1.Cells(16, 4), 1, "700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "3", xlHoja1.Cells(16, 6), 1, "700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "4", xlHoja1.Cells(16, 7), 1, "700", "B1")
    'Bonos corporativos emitidos por empresas privadas del sector no financiero (19) ANPS
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "1", xlHoja1.Cells(17, 3), 1, "750", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "2", xlHoja1.Cells(17, 4), 1, "750", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "3", xlHoja1.Cells(17, 6), 1, "750", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "4", xlHoja1.Cells(17, 7), 1, "750", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "1", xlHoja1.Cells(18, 3), 1, "750", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "2", xlHoja1.Cells(18, 4), 1, "750", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "3", xlHoja1.Cells(18, 6), 1, "750", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "4", xlHoja1.Cells(18, 7), 1, "750", "B1")
    'Valores representativos de deuda del BCRP y Gobierno Central recibidos en operaciones de reporte (22) ANPS COMENTADO  cuenta 15
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "1", xlHoja1.Cells(18, 3), 1, "760", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "2", xlHoja1.Cells(18, 4), 1, "760", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "3", xlHoja1.Cells(18, 6), 1, "760", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "4", xlHoja1.Cells(18, 7), 1, "760", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "1", xlHoja1.Cells(20, 3), 1, "760", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "2", xlHoja1.Cells(20, 4), 1, "760", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "3", xlHoja1.Cells(20, 6), 1, "760", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "4", xlHoja1.Cells(20, 7), 1, "760", "B1")  'ANPS
    'Valores representativos de deuda de Gobiernos del Exterior recibidos en operaciones de reporte (22) ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "1", xlHoja1.Cells(19, 3), 1, "770", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "2", xlHoja1.Cells(19, 4), 1, "770", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "3", xlHoja1.Cells(19, 6), 1, "770", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "4", xlHoja1.Cells(19, 7), 1, "770", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "1", xlHoja1.Cells(21, 3), 1, "770", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "2", xlHoja1.Cells(21, 4), 1, "770", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "3", xlHoja1.Cells(21, 6), 1, "770", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "4", xlHoja1.Cells(21, 7), 1, "770", "B1")  'ANPS

    'Bonos corporativos emitidos por empresas privadas del sector no financiero recibidos en operaciones de reporte (22) ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "1", xlHoja1.Cells(20, 3), 1, "780", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "2", xlHoja1.Cells(20, 4), 1, "780", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "3", xlHoja1.Cells(20, 6), 1, "780", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "4", xlHoja1.Cells(20, 7), 1, "780", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "1", xlHoja1.Cells(23, 3), 1, "780", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "2", xlHoja1.Cells(23, 4), 1, "780", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "3", xlHoja1.Cells(23, 6), 1, "780", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "4", xlHoja1.Cells(23, 7), 1, "780", "B1")  'ANPS
    'FLUJOS ENTRANTES 30 DÍAS
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "1", 0, 1, "1100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "2", 0, 1, "1100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "3", 0, 1, "1100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "4", 0, 1, "1100", "B1")
    
    '*******************DISPONIBLE**********************

    '***NAGL Según Anx03_ERS006-2019
    iFilaAnx = 26 'ANPS COMENTADO
    ' iFilaAnx = 30 'ANPS
     Set rsCtasAdic = ObtieneCtaSaldoDiarioAnx15B_Det("1114,1115,1116,111804", "111401", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio, "Det")
     If Not rsCtasAdic.BOF And Not rsCtasAdic.EOF Then
         Do While Not rsCtasAdic.EOF
             xlHoja1.Cells(iFilaAnx, 20) = Mid(rsCtasAdic!cCtaCnt, 1, 2) & "0" & Mid(rsCtasAdic!cCtaCnt, 4, Len(RTrim(rsCtasAdic!cCtaCnt)))
             xlHoja1.Cells(iFilaAnx, 21) = Format(rsCtasAdic!nSaldo, "#,##0.00")
             xlHoja1.Cells(iFilaAnx, 23) = rsCtasAdic!cDescrip
             iFilaAnx = iFilaAnx + 1
             rsCtasAdic.MoveNext
         Loop
     End If
     
     Set rsCtasAdic = Nothing
     iFilaAnx = 26 'ANPS COMENTADO
   ' iFilaAnx = 30 'ANPS
     Set rsCtasAdic = ObtieneCtaSaldoDiarioAnx15B_Det("1124,1125,1126,112804", "112401", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio, "Det")
     If Not rsCtasAdic.BOF And Not rsCtasAdic.EOF Then
         Do While Not rsCtasAdic.EOF
             'xlHoja1.Cells(iFilaAnx, 20) = rsCtasAdic!cCtaCnt
             xlHoja1.Cells(iFilaAnx, 22) = Format(rsCtasAdic!nSaldo, "#,##0.00")
             xlHoja1.Cells(iFilaAnx, 23) = rsCtasAdic!cDescrip
             iFilaAnx = iFilaAnx + 1
             rsCtasAdic.MoveNext
         Loop
     End If
     Set rsCtasAdic = Nothing
     '*************
'    nSaldoDiario1 = xlHoja1.Cells(28, 3) ANPS COMENTADO
'    nSaldoDiario2 = xlHoja1.Cells(28, 4) ANPS COMENTADO
    nSaldoDiario1 = xlHoja1.Cells(32, 3)    'ANPS
    nSaldoDiario2 = xlHoja1.Cells(32, 4)    'ANPS
    
'    xlHoja1.Range(xlHoja1.Cells(28, 3), xlHoja1.Cells(28, 4)).NumberFormat = "#,##0.00;-#,##0.00" ANPS COMENTADO
    xlHoja1.Range(xlHoja1.Cells(32, 3), xlHoja1.Cells(32, 4)).NumberFormat = "#,##0.00;-#,##0.00"   'ANPS
    
'    lnTotal2MN = Round(xlHoja1.Cells(28, 3), 2) 'VAPA20171101 ANPS COMENTADO
'    lnTotal2ME = Round(xlHoja1.Cells(28, 4), 2) 'VAPA20171101 ANPS COMENTADO
    lnTotal2MN = Round(xlHoja1.Cells(32, 3), 2) 'ANPS
    lnTotal2ME = Round(xlHoja1.Cells(32, 4), 2) 'ANPS
        
    Call oDbalanceCont.InsertaDetallaReporte15A(12, pdFecha, "1", nSaldoDiario1, 1, "1200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(12, pdFecha, "2", nSaldoDiario2, 1, "1200", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(12, pdFecha, "3", xlHoja1.Cells(28, 6), 1, "1200", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(12, pdFecha, "4", xlHoja1.Cells(28, 7), 1, "1200", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(12, pdFecha, "3", xlHoja1.Cells(32, 6), 1, "1200", "B1")    'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(12, pdFecha, "4", xlHoja1.Cells(32, 7), 1, "1200", "B1")    'ANPS
    oBarra.Progress 4, TituloProgress, MensajeProgress, "", vbBlue
    
    'NAGL ERS079-2016 20170407 *************************************** FONDOS DISPONIBLES EN EL SISTEMA FINANCIERO NACIONAL
    nSaldoDiario1 = (oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondos(pdFecha, 1, "01,02,04", "") + oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondos(pdFecha, 1, "03", "1090100012521") - oDbalanceCont.ObtenerDispobiblesenSFN("1090100822183", "1112", pdFecha, 1)) + DAnxVal.ObtenerSaldEstadistAnx15Ay15B("111803", "1", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio) - DAnxVal.ObtieneRestringidosSFN(pdFecha, "1") 'NAGL ERS079-2017 20180128
                    '+ ObtenerCtaContSaldoBalanceDiario("111803", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
    nSaldoDiario2 = (oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondos(pdFecha, 2, "01,02,04", "") + oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondos(pdFecha, 2, "03", "1090100012521") - oDbalanceCont.ObtenerDispobiblesenSFN("1090100822183", "1122", pdFecha, 1)) + Round(DAnxVal.ObtenerSaldEstadistAnx15Ay15B("112803", "2", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio), 2) - DAnxVal.ObtieneRestringidosSFN(pdFecha, "2") 'NAGL ERS079-2017 20180128
                    '+ Round(ObtenerCtaContSaldoBalanceDiario("112803", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) / lnTipoCambioBalanceAnterior, 2)
                    
    'xlHoja1.Cells(29, 3) = nSaldoDiario1 ANPS COMENTADO
    'xlHoja1.Cells(29, 4) = nSaldoDiario2 ANPS COMENTADO
    xlHoja1.Cells(33, 3) = nSaldoDiario1    'ANPS
    xlHoja1.Cells(33, 4) = nSaldoDiario2    'ANPS
    
    'lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(29, 3), 2) 'VAPA 20171101 ANPS COMENTADO
    'lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(29, 4), 2) 'VAPA 20171101 ANPS COMENTADO
    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(33, 3), 2)    'ANPS
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(33, 4), 2)    'ANPS
    
'    xlHoja1.Range(xlHoja1.Cells(29, 3), xlHoja1.Cells(29, 4)).NumberFormat = "#,##0.00;-#,##0.00" ANPS COMENTADO
    xlHoja1.Range(xlHoja1.Cells(33, 3), xlHoja1.Cells(29, 4)).NumberFormat = "#,##0.00;-#,##0.00" 'ANPS
    
    Call oDbalanceCont.InsertaDetallaReporte15A(13, pdFecha, "1", nSaldoDiario1, 1, "1300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(13, pdFecha, "2", nSaldoDiario2, 1, "1300", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(13, pdFecha, "3", xlHoja1.Cells(29, 6), 1, "1300", "B1") 'NAGL ERS079-2016 20170407 antes xlHoja1.Cells(25, 6)ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(13, pdFecha, "4", xlHoja1.Cells(29, 7), 1, "1300", "B1") 'NAGL ERS079-2016 20170407 antes xlHoja1.Cells(25, 7) ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(13, pdFecha, "3", xlHoja1.Cells(33, 6), 1, "1300", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(13, pdFecha, "4", xlHoja1.Cells(33, 7), 1, "1300", "B1") 'ANPS
    'Fondos disponibles en bancos del exterior de primera categoría (8)
    Call oDbalanceCont.InsertaDetallaReporte15A(14, pdFecha, "1", 0, 1, "1400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(14, pdFecha, "2", 0, 1, "1400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(14, pdFecha, "3", 0, 1, "1400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(14, pdFecha, "4", 0, 1, "1400", "B1")
    'Fondos interbancarios netos activos (9)
    Call oDbalanceCont.InsertaDetallaReporte15A(15, pdFecha, "1", 0, 1, "1500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(15, pdFecha, "2", 0, 1, "1500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(15, pdFecha, "3", 0, 1, "1500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(15, pdFecha, "4", 0, 1, "1500", "B1")
    
    'CRÉDITOS
    '***Comentado by NAGL 20191014
    'nSaldoDiario1 = oDbalanceCont.ObtenerColocacSaldosDiarioTramos(pdFecha, "1", "1") ' -oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141102", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141103", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141104", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141109", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141112", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141113", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141111", pdFecha, "1")
    'nSaldoDiario2 = oDbalanceCont.ObtenerColocacSaldosDiarioTramos(pdFecha, "1", "2") 'Round((oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("142102", pdFecha, "2") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141203", pdFecha, "2") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141204", pdFecha, "2") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141209", pdFecha, "2") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141212", pdFecha, "2") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141213", pdFecha, "2") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141211", pdFecha, "2")) / pnTipoCambio, 2)
    '***END NAGL 20191014
    
    '********** NAGL 20191014 Según Anx03_ERS006-2019*****************
    nSaldoDiario1 = oDbalanceCont.ObtenerColocacSaldosDiarioTramos(pdFecha, 1, "1", "C")
    nSaldoDiario2 = oDbalanceCont.ObtenerColocacSaldosDiarioTramos(pdFecha, 1, "2", "C")
    xlHoja1.Cells(15, 21) = nSaldoDiario1
    xlHoja1.Cells(15, 22) = nSaldoDiario2
    
    nSaldoDiario1 = oDbalanceCont.ObtenerColocacSaldosDiarioTramos(pdFecha, 1, "1", "I")
    nSaldoDiario2 = oDbalanceCont.ObtenerColocacSaldosDiarioTramos(pdFecha, 1, "2", "I")
    xlHoja1.Cells(16, 21) = nSaldoDiario1
    xlHoja1.Cells(16, 22) = nSaldoDiario2

    'nSaldoDiario1 = xlHoja1.Cells(32, 3) ANPS COMENTADO
    'nSaldoDiario2 = xlHoja1.Cells(32, 4) ANPS COMENTADO
    nSaldoDiario1 = xlHoja1.Cells(36, 3) 'ANPS
    nSaldoDiario2 = xlHoja1.Cells(36, 4) 'ANPS
    
    'lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(32, 3) * 0.5, 2) 'VAPA 20171101 ANPS COMENTADO
    'lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(32, 4) * 0.5, 2) 'VAPA 20171101 ANPS COMENTADO
    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(36, 3) * 0.5, 2) 'ANPS
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(36, 4) * 0.5, 2) 'ANPS
        
'    xlHoja1.Range(xlHoja1.Cells(32, 3), xlHoja1.Cells(32, 4)).NumberFormat = "#,##0.00;-#,##0.00" ANPS COMENTADO
    xlHoja1.Range(xlHoja1.Cells(32, 3), xlHoja1.Cells(36, 4)).NumberFormat = "#,##0.00;-#,##0.00" 'ANPS
    xlHoja1.Range(xlHoja1.Cells(15, 21), xlHoja1.Cells(16, 22)).NumberFormat = "#,##0.00;-#,##0.00" 'NAGL 20191107
    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "1", nSaldoDiario1, 1, "1600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "2", nSaldoDiario2, 1, "1600", "B1")
'    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "3", xlHoja1.Cells(32, 6), 1, "1600", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "4", xlHoja1.Cells(32, 7), 1, "1600", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "3", xlHoja1.Cells(36, 6), 1, "1600", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "4", xlHoja1.Cells(36, 7), 1, "1600", "B1") 'ANPS
    
    'Cuentas por cobrar - derivados para negociación (11)
    Call oDbalanceCont.InsertaDetallaReporte15A(17, pdFecha, "1", 0, 1, "1700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(17, pdFecha, "2", 0, 1, "1700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(17, pdFecha, "3", 0, 1, "1700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(17, pdFecha, "4", 0, 1, "1700", "B1")
    
 
    '**********NAGL ERS 006-2019 SECCIÓN CUENTAS POR COBRAR**************
    Set rsCtasAdic = DAnxVal.ObtieneCuentasContablesToAnexosLiqu("", "", "CtasMain", "xCobrar")
    iFilaAnxMain = 34 'ANPS
    iFilaCtaLiqu = 38 'ANPS
    
    If Not (rsCtasAdic.BOF Or rsCtasAdic.EOF) Then
    Do While Not rsCtasAdic.EOF
        'MN
        Set rs = DAnxVal.ObtieneCuentasContablesToAnexosLiqu(rsCtasAdic!cTpoCuentas, "1")
        nSaldoDiario1 = 0
        If rs!cCtaContCod <> "" Then
            iFilaAnx = iFilaAnxMain
            Do While Not rs.EOF
                If rs!TipoDato = "Diario" Then
                    nSaldoDiarioParam = DAnxVal.ObtieneCtaSaldoDiario(rs!cCtaContCod, pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio) * (IIf(rs!cSigno = "+", 1, -1))
                Else
                    nSaldoDiarioParam = ObtenerCtaContSaldoBalanceDiario(rs!cCtaContCod, pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario) * (IIf(rs!cSigno = "+", 1, -1))
                End If
                xlHoja1.Cells(iFilaAnx, 13) = rs!TipoDato
                xlHoja1.Cells(iFilaAnx, 14) = Mid(rs!cCtaContCod, 1, 2) & "0" & Mid(rs!cCtaContCod, 4, Len(rs!cCtaContCod))
                xlHoja1.Cells(iFilaAnx, 15) = nSaldoDiarioParam
                xlHoja1.Cells(iFilaAnx, 17) = rs!cDescrip
                xlHoja1.Range(xlHoja1.Cells(iFilaAnx, 13), xlHoja1.Cells(iFilaAnx, 17)).Font.Name = "Arial Narrow"
                xlHoja1.Range(xlHoja1.Cells(iFilaAnx, 13), xlHoja1.Cells(iFilaAnx, 17)).Font.Size = 10
                ExcelCuadro xlHoja1, 13, iFilaAnx, 17, CCur(iFilaAnx)
                nSaldoDiario1 = IIf(rs!cSigno = "+", nSaldoDiarioParam, 0) + nSaldoDiario1
                iFilaAnx = iFilaAnx + 1
                rs.MoveNext
            Loop
        End If
        Set rs = Nothing
        'ME
        Set rs = DAnxVal.ObtieneCuentasContablesToAnexosLiqu(rsCtasAdic!cTpoCuentas, "2")
        nSaldoDiario2 = 0
        If rs!cCtaContCod <> "" Then
            iFilaAnx = iFilaAnxMain
            Do While Not rs.EOF
                If rs!TipoDato = "Diario" Then
                    nSaldoDiarioParam = DAnxVal.ObtieneCtaSaldoDiario(rs!cCtaContCod, pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio) * (IIf(rs!cSigno = "+", 1, -1))
                Else
                    nSaldoDiarioParam = Round(ObtenerCtaContSaldoBalanceDiario(rs!cCtaContCod, pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) / lnTipoCambioBalanceAnterior, 2) * (IIf(rs!cSigno = "+", 1, -1))
                End If
                xlHoja1.Cells(iFilaAnx, 16) = nSaldoDiarioParam
                nSaldoDiario2 = IIf(rs!cSigno = "+", nSaldoDiarioParam, 0) + nSaldoDiario2
                iFilaAnx = iFilaAnx + 1
                rs.MoveNext
            Loop
        End If
        Set rs = Nothing
        iFilaAnxMain = iFilaAnx
        If iFilaCtaLiqu = 41 Then
        iFilaCtaLiqu = 42
        End If
        If iFilaCtaLiqu = 43 Then
        iFilaCtaLiqu = 44
        End If
        If iFilaCtaLiqu = 39 Then 'ANPS
'*********ANPS Cuentas por cobrar - operaciones de reporte con valores del BCRP y Gobierno Central  ? 30 d (23)

            nSaldoDiario1 = oDbalanceCont.ObtenerSumaValorBCRP(pdFecha, "1", 1)
            nSaldoDiario2 = oDbalanceCont.ObtenerSumaValorBCRP(pdFecha, "2", 1)
        End If
            xlHoja1.Cells(iFilaCtaLiqu, 3) = nSaldoDiario1
            xlHoja1.Cells(iFilaCtaLiqu, 4) = nSaldoDiario2
             iFilaCtaLiqu = iFilaCtaLiqu + 1
            rsCtasAdic.MoveNext

    Loop
    End If
    Set rsCtasAdic = Nothing
    
'   xlHoja1.Range(xlHoja1.Cells(34, 15), xlHoja1.Cells(iFilaAnx - 1, 16)).NumberFormat = "#,##0.00;-#,##0.00"  ANPS COMENTADO
    xlHoja1.Range(xlHoja1.Cells(38, 15), xlHoja1.Cells(iFilaAnx - 1, 16)).NumberFormat = "#,##0.00;-#,##0.00"  'ANPS
        
'    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(34, 3) * 0.8, 2) 'VAPA 20171101 ANPS COMENTADO
'    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(34, 4) * 0.8, 2) 'VAPA 20171101 ANPS COMENTADO
    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(38, 3) * 0.8, 2) 'ANPS
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(38, 4) * 0.8, 2) 'ANPS
    
'    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(35, 3) * 0, 2) 'NAGL 20190509 ANPS COMENTADO
'    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(35, 4) * 0, 2) 'NAGL 20190509 ANPS COMENTADO
    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(39, 3) * 0, 2) 'ANPS
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(39, 4) * 0, 2) 'ANPS
        
'    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(36, 3) * 0.15, 2) 'NAGL 20190509 ANPS COMENTADO
'    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(36, 4) * 0.15, 2) 'NAGL 20190509 ANPS COMENTADO
    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(40, 3) * 0.15, 2) 'ANPS
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(40, 4) * 0.15, 2) 'ANPS

'ADICIONA UN CAMBIO FILa 41 - 43
    '*********ANPS Valores representativos de deuda del BCRP y Gobierno Central recibidos en operaciones de reporte (22)

'    nSaldoDiario1 = oDbalanceCont.ObtenerSumaValorBCRP(pdFecha, "1", 1)
'    nSaldoDiario2 = oDbalanceCont.ObtenerSumaValorBCRP(pdFecha, "2", 1)
'
'    xlHoja1.Cells(41, 3) = nSaldoDiario1 'SOLES
'    xlHoja1.Cells(43, 4) = nSaldoDiario2 'DOLARES
'

'    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(37, 3) * 0.25, 2) 'NAGL 20190509 ANPS COMENTADO
'    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(37, 4) * 0.25, 2) 'NAGL 20190509 ANPS COMENTADO
    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(42, 3) * 0.25, 2) 'ANPS
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(42, 4) * 0.25, 2) 'ANPS
        
'    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(38, 3), 2) 'NAGL 20190509 ANPS COMENTADO
'    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(38, 4), 2) 'NAGL 20190509 ANPS COMENTADO
    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(44, 3), 2) 'ANPS
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(44, 4), 2) 'ANPS
       
'    xlHoja1.Range(xlHoja1.Cells(34, 3), xlHoja1.Cells(38, 4)).NumberFormat = "#,##0.00;-#,##0.00" ANPS COMENTADO
    xlHoja1.Range(xlHoja1.Cells(38, 3), xlHoja1.Cells(44, 4)).NumberFormat = "#,##0.00;-#,##0.00" 'ANPS
        
    'Cuentas por cobrar - otros (11)
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "1", xlHoja1.Cells(34, 3), 1, "1800", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "2", xlHoja1.Cells(34, 4), 1, "1800", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "3", xlHoja1.Cells(34, 6), 1, "1800", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "4", xlHoja1.Cells(34, 7), 1, "1800", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "1", xlHoja1.Cells(38, 3), 1, "1800", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "2", xlHoja1.Cells(38, 4), 1, "1800", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "3", xlHoja1.Cells(38, 6), 1, "1800", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "4", xlHoja1.Cells(38, 7), 1, "1800", "B1") 'ANPS
    'Cuentas por cobrar - operaciones de reporte con valores del BCRP y Gobierno Central (23)
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "1", xlHoja1.Cells(35, 3), 1, "1810", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "2", xlHoja1.Cells(35, 4), 1, "1810", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "3", xlHoja1.Cells(35, 6), 1, "1810", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "4", xlHoja1.Cells(35, 7), 1, "1810", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "1", xlHoja1.Cells(39, 3), 1, "1810", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "2", xlHoja1.Cells(39, 4), 1, "1810", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "3", xlHoja1.Cells(39, 6), 1, "1810", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "4", xlHoja1.Cells(39, 7), 1, "1810", "B1") 'ANPS
    'Cuentas por cobrar - operaciones de reporte con valores de Gobiernos del Exterior (23)
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "1", xlHoja1.Cells(36, 3), 1, "1820", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "2", xlHoja1.Cells(36, 4), 1, "1820", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "3", xlHoja1.Cells(36, 6), 1, "1820", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "4", xlHoja1.Cells(36, 7), 1, "1820", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "1", xlHoja1.Cells(40, 3), 1, "1820", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "2", xlHoja1.Cells(40, 4), 1, "1820", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "3", xlHoja1.Cells(40, 6), 1, "1820", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "4", xlHoja1.Cells(40, 7), 1, "1820", "B1") 'ANPS
 
 'ADICIONA UN CAMBIO FILa 41 - 43
    
    'Cuentas por cobrar - operaciones de reporte con bonos corporativos emitidos por empresas privadas del sector no financiero (23)
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "1", xlHoja1.Cells(37, 3), 1, "1830", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "2", xlHoja1.Cells(37, 4), 1, "1830", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "3", xlHoja1.Cells(37, 6), 1, "1830", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "4", xlHoja1.Cells(37, 7), 1, "1830", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "1", xlHoja1.Cells(42, 3), 1, "1830", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "2", xlHoja1.Cells(42, 4), 1, "1830", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "3", xlHoja1.Cells(42, 6), 1, "1830", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "4", xlHoja1.Cells(42, 7), 1, "1830", "B1") 'ANPS
    'Cuentas por cobrar - operaciones de reporte con otros valores (24)
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "1", xlHoja1.Cells(38, 3), 1, "1840", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "2", xlHoja1.Cells(38, 4), 1, "1840", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "3", xlHoja1.Cells(38, 6), 1, "1840", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "4", xlHoja1.Cells(38, 7), 1, "1840", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "1", xlHoja1.Cells(44, 3), 1, "1840", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "2", xlHoja1.Cells(44, 4), 1, "1840", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "3", xlHoja1.Cells(44, 6), 1, "1840", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "4", xlHoja1.Cells(44, 7), 1, "1840", "B1") 'ANPS
    
    '***************END NAGL 20190428****************************
    'Operaciones por liquidar (20)
'    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "1", xlHoja1.Cells(21, 3), 1, "1850", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "2", xlHoja1.Cells(21, 4), 1, "1850", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "3", xlHoja1.Cells(21, 6), 1, "1850", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "4", xlHoja1.Cells(21, 7), 1, "1850", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "1", xlHoja1.Cells(25, 3), 1, "1850", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "2", xlHoja1.Cells(25, 4), 1, "1850", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "3", xlHoja1.Cells(25, 6), 1, "1850", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "4", xlHoja1.Cells(25, 7), 1, "1850", "B1") 'ANPS
    
    'Posiciones activas en derivados - Delivery (12)
'    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "1", xlHoja1.Cells(22, 3), 1, "1900", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "2", xlHoja1.Cells(22, 4), 1, "1900", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "3", xlHoja1.Cells(22, 6), 1, "1900", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "4", xlHoja1.Cells(22, 7), 1, "1900", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "1", xlHoja1.Cells(26, 3), 1, "1900", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "2", xlHoja1.Cells(26, 4), 1, "1900", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "3", xlHoja1.Cells(26, 6), 1, "1900", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "4", xlHoja1.Cells(26, 7), 1, "1900", "B1") 'ANPS
    
    '****END - Restante****'
    
    'FLUJOS SALIENTES 30 DÍAS
    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "1", 0, 1, "2100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "2", 0, 1, "2100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "3", 0, 1, "2100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "4", 0, 1, "2100", "B1")
    
    'SECCIÓN FONDEO'
    oBarra.Progress 5, TituloProgress, MensajeProgress, "", vbBlue
    
    '**************NAGL Según ERS006-2019 20190424 PROCESO FONDEO***************************'
    Set loRs = oDbalanceCont.ObtenerFondeoEncajeConsolidado(pdFecha, pnTipoCambio)
        If Not (loRs.BOF Or loRs.EOF) Then
        Do While Not loRs.EOF
'            If loRs!cTipo = "FE" Then                      ANPS COMENTADO
'                    xlHoja1.Cells(43, 3) = loRs!nSaldCntMN
'                    xlHoja1.Cells(43, 4) = loRs!nSaldCntME
'            ElseIf loRs!cTipo = "FME_PN_PJSFC" Then
'                    xlHoja1.Cells(44, 3) = loRs!nSaldCntMN
'                    xlHoja1.Cells(44, 4) = loRs!nSaldCntME
'            ElseIf loRs!cTipo = "FME_PJCFC" Then
'                    xlHoja1.Cells(45, 3) = loRs!nSaldCntMN
'                    xlHoja1.Cells(45, 4) = loRs!nSaldCntME
'            ElseIf loRs!cTipo = "FGA" Then
'                    xlHoja1.Cells(47, 3) = loRs!nSaldCntMN
'                    xlHoja1.Cells(47, 4) = loRs!nSaldCntME
'            End If
       If loRs!cTipo = "FE" Then                            'ANPS
                    xlHoja1.Cells(49, 3) = loRs!nSaldCntMN  'ANPS
                    xlHoja1.Cells(49, 4) = loRs!nSaldCntME 'ANPS
            ElseIf loRs!cTipo = "FME_PN_PJSFC" Then
                    xlHoja1.Cells(50, 3) = loRs!nSaldCntMN 'ANPS
                    xlHoja1.Cells(50, 4) = loRs!nSaldCntME 'ANPS
            ElseIf loRs!cTipo = "FME_PJCFC" Then
                    xlHoja1.Cells(51, 3) = loRs!nSaldCntMN 'ANPS
                    xlHoja1.Cells(51, 4) = loRs!nSaldCntME 'ANPS
            ElseIf loRs!cTipo = "FGA" Then
                    xlHoja1.Cells(53, 3) = loRs!nSaldCntMN 'ANPS
                    xlHoja1.Cells(53, 4) = loRs!nSaldCntME 'ANPS
            End If
            loRs.MoveNext
        Loop
        End If
    '**************************END NAGL ERS006-2019***************************'
    '*****************Inclusión de Obligaciones a la Vista***************************
    nSaldoDiario1 = DAnxVal.ObtieneCtaSaldoDiario("2111", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio)
    nSaldoDiario2 = DAnxVal.ObtieneCtaSaldoDiario("2121", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio)
    xlHoja1.Cells(21, 21) = nSaldoDiario1  'ANPS COMENTADO
    xlHoja1.Cells(21, 22) = nSaldoDiario2  'ANPS COMENTADO
'   xlHoja1.Cells(25, 21) = nSaldoDiario1 'ANPS
'    xlHoja1.Cells(25, 22) = nSaldoDiario2 'ANPS
    
    '****************NAGL Según Anx03_ERS006-2019***********************************
    
    'xlHoja1.Cells(43, 3) = IIf(CCur(xlHoja1.Cells(43, 3)) < 0, 0, CCur(xlHoja1.Cells(43, 3)))
'    xlHoja1.Cells(43, 3).Formula = "=" & IIf(CCur(xlHoja1.Cells(43, 3)) < 0, 0, CCur(xlHoja1.Cells(43, 3))) & "+" & xlHoja1.Range(xlHoja1.Cells(21, 21), xlHoja1.Cells(21, 21)).Address(False, False) 'NAGL 20191014 Según Anx03_ERS006-2019  ANPS COMENTADO
'    xlHoja1.Cells(44, 3) = IIf(CCur(xlHoja1.Cells(44, 3)) < 0, 0, CCur(xlHoja1.Cells(44, 3)))   ANPS COMENTADO
'    xlHoja1.Cells(45, 3) = IIf(CCur(xlHoja1.Cells(45, 3)) < 0, 0, CCur(xlHoja1.Cells(45, 3)))   ANPS COMENTADO
    xlHoja1.Cells(49, 3).Formula = "=" & IIf(CCur(xlHoja1.Cells(49, 3)) < 0, 0, CCur(xlHoja1.Cells(49, 3))) & "+" & xlHoja1.Range(xlHoja1.Cells(21, 21), xlHoja1.Cells(21, 21)).Address(False, False)    'ANPS
    xlHoja1.Cells(50, 3) = IIf(CCur(xlHoja1.Cells(50, 3)) < 0, 0, CCur(xlHoja1.Cells(50, 3)))   'ANPS
    xlHoja1.Cells(51, 3) = IIf(CCur(xlHoja1.Cells(51, 3)) < 0, 0, CCur(xlHoja1.Cells(51, 3)))   'ANPS

    'xlHoja1.Cells(43, 4) = IIf(CCur(xlHoja1.Cells(43, 4)) < 0, 0, CCur(xlHoja1.Cells(43, 4)))
'    xlHoja1.Cells(43, 4).Formula = "=" & IIf(CCur(xlHoja1.Cells(43, 4)) < 0, 0, CCur(xlHoja1.Cells(43, 4))) & "+" & xlHoja1.Range(xlHoja1.Cells(21, 22), xlHoja1.Cells(21, 22)).Address(False, False) 'NAGL 20191014 Según Anx03_ERS006-2019 ANPS COMENTADO
'    xlHoja1.Cells(44, 4) = IIf(CCur(xlHoja1.Cells(44, 4)) < 0, 0, CCur(xlHoja1.Cells(44, 4)))
'    xlHoja1.Cells(45, 4) = IIf(CCur(xlHoja1.Cells(45, 4)) < 0, 0, CCur(xlHoja1.Cells(45, 4)))
'    xlHoja1.Cells(47, 3) = IIf(CCur(xlHoja1.Cells(47, 3)) < 0, 0, CCur(xlHoja1.Cells(47, 3)))
'    xlHoja1.Cells(47, 4) = IIf(CCur(xlHoja1.Cells(47, 4)) < 0, 0, CCur(xlHoja1.Cells(47, 4)))
    xlHoja1.Cells(49, 4).Formula = "=" & IIf(CCur(xlHoja1.Cells(49, 4)) < 0, 0, CCur(xlHoja1.Cells(49, 4))) & "+" & xlHoja1.Range(xlHoja1.Cells(21, 22), xlHoja1.Cells(21, 22)).Address(False, False)    'ANPS
    xlHoja1.Cells(50, 4) = IIf(CCur(xlHoja1.Cells(50, 4)) < 0, 0, CCur(xlHoja1.Cells(50, 4)))   'ANPS
    xlHoja1.Cells(51, 4) = IIf(CCur(xlHoja1.Cells(51, 4)) < 0, 0, CCur(xlHoja1.Cells(51, 4)))   'ANPS
    xlHoja1.Cells(53, 3) = IIf(CCur(xlHoja1.Cells(53, 3)) < 0, 0, CCur(xlHoja1.Cells(53, 3)))   'ANPS
    xlHoja1.Cells(53, 4) = IIf(CCur(xlHoja1.Cells(53, 4)) < 0, 0, CCur(xlHoja1.Cells(53, 4)))   'ANPS
        
'    xlHoja1.Range(xlHoja1.Cells(43, 3), xlHoja1.Cells(45, 4)).NumberFormat = "#,##0.00;-#,##0.00"  ANPS COMENTADO
'    xlHoja1.Range(xlHoja1.Cells(47, 3), xlHoja1.Cells(47, 4)).NumberFormat = "#,##0.00;-#,##0.00"  ANPS COMENTADO
'    xlHoja1.Range(xlHoja1.Cells(21, 21), xlHoja1.Cells(21, 22)).NumberFormat = "#,##0.00;-#,##0.00" 'NAGL 20191107 ANPS COMENTADO
    xlHoja1.Range(xlHoja1.Cells(49, 3), xlHoja1.Cells(51, 4)).NumberFormat = "#,##0.00;-#,##0.00"   'ANPS
    xlHoja1.Range(xlHoja1.Cells(53, 3), xlHoja1.Cells(53, 4)).NumberFormat = "#,##0.00;-#,##0.00"   'ANPS
    xlHoja1.Range(xlHoja1.Cells(21, 21), xlHoja1.Cells(21, 22)).NumberFormat = "#,##0.00;-#,##0.00"   'ANPS
    
'    lnTotal3MN = Round(xlHoja1.Cells(43, 3) * 0.075, 2) 'VAPA20171101  ANPS COMENTADO
'    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(44, 3) * 0.15, 2) 'VAPA20171101  ANPS COMENTADO
'    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(45, 3) * 0.15, 2) 'VAPA20171101  ANPS COMENTADO
    lnTotal3MN = Round(xlHoja1.Cells(49, 3) * 0.075, 2) 'ANPS
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(50, 3) * 0.15, 2) 'ANPS
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(51, 3) * 0.15, 2) 'ANPS
    
'    lnTotal3ME = Round(xlHoja1.Cells(43, 4) * 0.075, 2) 'VAPA20171101 ANPS COMENTADO
'    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(44, 4) * 0.15, 2) 'VAPA20171101 ANPS COMENTADO
'    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(45, 4) * 0.15, 2) 'VAPA20171101 ANPS COMENTADO
    lnTotal3ME = Round(xlHoja1.Cells(49, 4) * 0.075, 2) 'ANPS
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(50, 4) * 0.15, 2) 'ANPS
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(51, 4) * 0.15, 2) 'ANPS

'    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(47, 3) * 0.3, 2) 'VAPA20171101 ANPS COMENTADO
'    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(47, 4) * 0.3, 2) 'VAPA20171101 ANPS COMENTADO
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(53, 3) * 0.3, 2) 'ANPS
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(53, 4) * 0.3, 2) 'ANPS
    
'    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "1", xlHoja1.Cells(43, 3), 1, "2200", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "2", xlHoja1.Cells(43, 4), 1, "2200", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "3", xlHoja1.Cells(43, 6), 1, "2200", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "4", xlHoja1.Cells(43, 7), 1, "2200", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "1", xlHoja1.Cells(49, 3), 1, "2200", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "2", xlHoja1.Cells(49, 4), 1, "2200", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "3", xlHoja1.Cells(49, 6), 1, "2200", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "4", xlHoja1.Cells(49, 7), 1, "2200", "B1") 'ANPS
    
'    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "1", xlHoja1.Cells(44, 3), 1, "2300", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "2", xlHoja1.Cells(44, 4), 1, "2300", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "3", xlHoja1.Cells(44, 6), 1, "2300", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "4", xlHoja1.Cells(44, 7), 1, "2300", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "1", xlHoja1.Cells(50, 3), 1, "2300", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "2", xlHoja1.Cells(50, 4), 1, "2300", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "3", xlHoja1.Cells(50, 6), 1, "2300", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "4", xlHoja1.Cells(50, 7), 1, "2300", "B1") 'ANPS
    
'    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "1", xlHoja1.Cells(45, 3), 1, "2400", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "2", xlHoja1.Cells(45, 4), 1, "2400", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "3", xlHoja1.Cells(45, 6), 1, "2400", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "4", xlHoja1.Cells(45, 7), 1, "2400", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "1", xlHoja1.Cells(51, 3), 1, "2400", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "2", xlHoja1.Cells(51, 4), 1, "2400", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "3", xlHoja1.Cells(51, 6), 1, "2400", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "4", xlHoja1.Cells(51, 7), 1, "2400", "B1") 'ANPS
    
'    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "1", xlHoja1.Cells(46, 3), 1, "2500", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "2", xlHoja1.Cells(46, 4), 1, "2500", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "3", xlHoja1.Cells(46, 6), 1, "2500", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "4", xlHoja1.Cells(46, 7), 1, "2500", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "1", xlHoja1.Cells(52, 3), 1, "2500", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "2", xlHoja1.Cells(52, 4), 1, "2500", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "3", xlHoja1.Cells(52, 6), 1, "2500", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "4", xlHoja1.Cells(52, 7), 1, "2500", "B1") 'ANPS
    
'    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "1", xlHoja1.Cells(47, 3), 1, "2600", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "2", xlHoja1.Cells(47, 4), 1, "2600", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "3", xlHoja1.Cells(47, 6), 1, "2600", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "4", xlHoja1.Cells(47, 7), 1, "2600", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "1", xlHoja1.Cells(53, 3), 1, "2600", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "2", xlHoja1.Cells(53, 4), 1, "2600", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "3", xlHoja1.Cells(53, 6), 1, "2600", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "4", xlHoja1.Cells(53, 7), 1, "2600", "B1") 'ANPS
    
    '*********END FONDEO***************'
    
    oBarra.Progress 6, TituloProgress, MensajeProgress, "", vbBlue
    
    '************ NAGL Según ERS006-2019 20190430 Proceso Otras obligac e Inst Rec. < 30 Días**********
    pnFila = 15
    Set loRs = Nothing
    Set loRs = oDbalanceCont.ObtenerObligInstRecMenora30Dias(pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio)
        If Not (loRs.BOF Or loRs.EOF) Then
            Do While Not loRs.EOF
                xlHoja1.Cells(pnFila, 13) = loRs!cTipo
                xlHoja1.Cells(pnFila, 14) = loRs!cCtaContCod
                xlHoja1.Cells(pnFila, 15) = loRs!nSaldoMN
                xlHoja1.Cells(pnFila, 16) = loRs!nSaldoME
                xlHoja1.Range(xlHoja1.Cells(pnFila, 13), xlHoja1.Cells(pnFila, 16)).Font.Name = "Arial Narrow"
                xlHoja1.Range(xlHoja1.Cells(pnFila, 13), xlHoja1.Cells(pnFila, 16)).Font.Size = 10
                xlHoja1.Range(xlHoja1.Cells(pnFila, 14), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(pnFila, 15), xlHoja1.Cells(pnFila, 16)).NumberFormat = "#,##0.00;-#,##0.00"
                pnFila = pnFila + 1
                loRs.MoveNext
            Loop
        End If
    '************************************END NAGL ERS006-2019*******************************************'
    'xlHoja1.Cells(48, 3) = nSaldoDiario1
    'xlHoja1.Cells(48, 4) = nSaldoDiario2
'    xlHoja1.Range(xlHoja1.Cells(48, 3), xlHoja1.Cells(48, 4)).NumberFormat = "#,##0.00;-#,##0.00" ANPS COMENTADO
    xlHoja1.Range(xlHoja1.Cells(54, 3), xlHoja1.Cells(54, 4)).NumberFormat = "#,##0.00;-#,##0.00" 'ANPS
    
'    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(48, 3), 2)   'VAPA20171101 ANPS COMENTADO
'    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(48, 4), 2)  'VAPA20171101 ANPS COMENTADO
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(54, 3), 2)   'ANPS
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(54, 4), 2)   'ANPS
    
'    Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "1", xlHoja1.Cells(48, 3), 1, "2800", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "2", xlHoja1.Cells(48, 4), 1, "2800", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "3", xlHoja1.Cells(48, 6), 1, "2800", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "4", xlHoja1.Cells(48, 7), 1, "2800", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "1", xlHoja1.Cells(54, 3), 1, "2800", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "2", xlHoja1.Cells(54, 4), 1, "2800", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "3", xlHoja1.Cells(54, 6), 1, "2800", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "4", xlHoja1.Cells(54, 7), 1, "2800", "B1")  'ANPS
    
     '**************END Otras obligac e Inst Rec. < 30 Días************************
    
    'Fondos interbancarios netos pasivos (9)
'    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "1", xlHoja1.Cells(49, 3), 1, "2900", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "2", xlHoja1.Cells(49, 4), 1, "2900", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "3", xlHoja1.Cells(49, 6), 1, "2900", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "4", xlHoja1.Cells(49, 7), 1, "2900", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "1", xlHoja1.Cells(55, 3), 1, "2900", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "2", xlHoja1.Cells(55, 4), 1, "2900", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "3", xlHoja1.Cells(55, 6), 1, "2900", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "4", xlHoja1.Cells(55, 7), 1, "2900", "B1")  'ANPS
    
    oBarra.Progress 7, TituloProgress, MensajeProgress, "", vbBlue
    
     '***********NAGL ERS079-2016 20170407 Depósitos de empresas del sistema financiero y OFI
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("2312", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("2313", pdFecha, "1")
    nSaldoDiario1 = nSaldoDiario1 + DAnxVal.ObtenerSaldEstadistAnx15Ay15B("231802", "1", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio) + DAnxVal.ObtenerSaldEstadistAnx15Ay15B("231803", "1", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio) 'NAGL ERS079-2017 20180128 'ObtenerCtaContSaldoBalanceDiario("2318", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
    'NAGL 20190628 Según Anx02_ERS006-2019
    nSaldoDiario2 = Round(oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("2322", pdFecha, "2"), 2) + Round(oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("2323", pdFecha, "2"), 2)
    nSaldoDiario2 = nSaldoDiario2 + Round(DAnxVal.ObtenerSaldEstadistAnx15Ay15B("232802", "2", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio), 2) + Round(DAnxVal.ObtenerSaldEstadistAnx15Ay15B("232803", "2", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio), 2)  'NAGL ERS079-2017 20180128 'Round(ObtenerCtaContSaldoBalanceDiario("2328", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) / lnTipoCambioBalanceAnterior, 2)
    'NAGL 20190628 Según Anx02_ERS006-2019
    
'    xlHoja1.Cells(50, 3) = nSaldoDiario1 ANPS COMENTADO
'    xlHoja1.Cells(50, 4) = nSaldoDiario2 ANPS COMENTADO
'    xlHoja1.Range(xlHoja1.Cells(50, 3), xlHoja1.Cells(50, 4)).NumberFormat = "#,##0.00;-#,##0.00" '******NAGL ANPS COMENTADO
    xlHoja1.Cells(56, 3) = nSaldoDiario1  'ANPS
    xlHoja1.Cells(56, 4) = nSaldoDiario2  'ANPS
    xlHoja1.Range(xlHoja1.Cells(56, 3), xlHoja1.Cells(56, 4)).NumberFormat = "#,##0.00;-#,##0.00"  'ANPS
    
'    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(50, 3), 2) 'VAPA20171101 ANPS COMENTADO
'    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(50, 4), 2) 'VAPA20171101 ANPS COMENTADO
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(56, 3), 2)  'ANPS
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(56, 4), 2)  'ANPS
    
'    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "1", xlHoja1.Cells(50, 3), 1, "3000", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "2", xlHoja1.Cells(50, 4), 1, "3000", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "3", xlHoja1.Cells(50, 6), 1, "3000", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "4", xlHoja1.Cells(50, 7), 1, "3000", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "1", xlHoja1.Cells(56, 3), 1, "3000", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "2", xlHoja1.Cells(56, 4), 1, "3000", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "3", xlHoja1.Cells(56, 6), 1, "3000", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "4", xlHoja1.Cells(56, 7), 1, "3000", "B1")  'ANPS
    
    'Adeudos y obligaciones financieras con vencimiento <= 30 días
'    nSaldoDiario1 = oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "1", 1, 0) + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "1", 1, 1)
'    nSaldoDiario2 = oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "2", 1, 0) + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "2", 1, 1)
    
    'JIPR20200824
    nSaldoDiario1 = oDbalanceCont.MontoAdeudados15BCLP(pdFecha, "1", 1)
    nSaldoDiario2 = oDbalanceCont.MontoAdeudados15BCLP(pdFecha, "2", 1)
'    xlHoja1.Cells(51, 3) = nSaldoDiario1 ANPS COMENTADO
'    xlHoja1.Cells(51, 4) = nSaldoDiario2 ANPS COMENTADO
'    xlHoja1.Range(xlHoja1.Cells(51, 3), xlHoja1.Cells(51, 4)).NumberFormat = "#,##0.00;-#,##0.00" ANPS COMENTADO
    xlHoja1.Cells(57, 3) = nSaldoDiario1  'ANPS
    xlHoja1.Cells(57, 4) = nSaldoDiario2  'ANPS
    xlHoja1.Range(xlHoja1.Cells(57, 3), xlHoja1.Cells(57, 4)).NumberFormat = "#,##0.00;-#,##0.00"  'ANPS
    
'    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(51, 3), 2) 'VAPA20171101 ANPS COMENTADO
'    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(51, 4), 2) 'VAPA20171101 ANPS COMENTADO
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(57, 3), 2)  'ANPS
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(57, 4), 2)  'ANPS
    
'    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "1", xlHoja1.Cells(51, 3), 1, "3100", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "2", xlHoja1.Cells(51, 4), 1, "3100", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "3", xlHoja1.Cells(51, 6), 1, "3100", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "4", xlHoja1.Cells(51, 7), 1, "3100", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "1", xlHoja1.Cells(57, 3), 1, "3100", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "2", xlHoja1.Cells(57, 4), 1, "3100", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "3", xlHoja1.Cells(57, 6), 1, "3100", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "4", xlHoja1.Cells(57, 7), 1, "3100", "B1")  'ANPS
    
    'JIPR20200824
    'Adeudos y obligaciones financieras con vencimiento de 31 a 90 días
   ' nSaldoDiario1 = oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "1", 3, 0) + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "1", 3, 1)
    'nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "1", 3, 0) + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "1", 3, 1)
    'nSaldoDiario2 = oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "2", 3, 0) + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "2", 3, 1)
    'nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "2", 3, 0) + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "2", 3, 1)
    
    nSaldoDiario1 = oDbalanceCont.MontoAdeudados15BCLP(pdFecha, "1", 2)
    nSaldoDiario2 = oDbalanceCont.MontoAdeudados15BCLP(pdFecha, "2", 2)
    
'    xlHoja1.Cells(52, 3) = nSaldoDiario1 ANPS COMENTADO
'    xlHoja1.Cells(52, 4) = nSaldoDiario2 ANPS COMENTADO
'    xlHoja1.Range(xlHoja1.Cells(52, 3), xlHoja1.Cells(52, 4)).NumberFormat = "#,##0.00;-#,##0.00" ANPS COMENTADO
    xlHoja1.Cells(58, 3) = nSaldoDiario1  'ANPS
    xlHoja1.Cells(58, 4) = nSaldoDiario2  'ANPS
    xlHoja1.Range(xlHoja1.Cells(58, 3), xlHoja1.Cells(58, 4)).NumberFormat = "#,##0.00;-#,##0.00"  'ANPS
    
'    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(52, 3) * 0.3, 2) 'VAPA20171101 ANPS COMENTADO
'    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(52, 4) * 0.3, 2) 'VAPA20171101 ANPS COMENTADO
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(58, 3) * 0.3, 2)  'ANPS
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(58, 4) * 0.3, 2)  'ANPS
    
'    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "1", xlHoja1.Cells(52, 3), 1, "3200", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "2", xlHoja1.Cells(52, 4), 1, "3200", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "3", xlHoja1.Cells(52, 6), 1, "3200", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "4", xlHoja1.Cells(52, 7), 1, "3200", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "1", xlHoja1.Cells(58, 3), 1, "3200", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "2", xlHoja1.Cells(58, 4), 1, "3200", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "3", xlHoja1.Cells(58, 6), 1, "3200", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "4", xlHoja1.Cells(58, 7), 1, "3200", "B1")  'ANPS
    
    oBarra.Progress 8, TituloProgress, MensajeProgress, "", vbBlue


    
    'Valores, títulos y obligaciones en circulación <= 30 día
'    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "1", xlHoja1.Cells(53, 3), 1, "3300", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "2", xlHoja1.Cells(53, 4), 1, "3300", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "3", xlHoja1.Cells(53, 6), 1, "3300", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "4", xlHoja1.Cells(53, 7), 1, "3300", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "1", xlHoja1.Cells(59, 3), 1, "3300", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "2", xlHoja1.Cells(59, 4), 1, "3300", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "3", xlHoja1.Cells(59, 6), 1, "3300", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "4", xlHoja1.Cells(59, 7), 1, "3300", "B1")  'ANPS
    
    'Cuentas por pagar - derivados para negociación (17)
'    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "1", xlHoja1.Cells(54, 3), 1, "3400", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "2", xlHoja1.Cells(54, 4), 1, "3400", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "3", xlHoja1.Cells(54, 6), 1, "3400", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "4", xlHoja1.Cells(54, 7), 1, "3400", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "1", xlHoja1.Cells(60, 3), 1, "3400", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "2", xlHoja1.Cells(60, 4), 1, "3400", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "3", xlHoja1.Cells(60, 6), 1, "3400", "B1")  'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "4", xlHoja1.Cells(60, 7), 1, "3400", "B1")  'ANPS
    
    '**********NAGL ERS006-2019 CUENTAS POR PAGAR**************
    Set rsCtasAdic = DAnxVal.ObtieneCuentasContablesToAnexosLiqu("", "", "CtasMain", "xPagar")
    iFilaAnxMain = iFilaAnx
    iFilaAnxParam = iFilaAnx
    'iFilaCtaLiqu = 55 'ANPS COMENTADO
    iFilaCtaLiqu = 61  'ANPS
    If Not (rsCtasAdic.BOF Or rsCtasAdic.EOF) Then
    Do While Not rsCtasAdic.EOF
        'MN
        Set rs = DAnxVal.ObtieneCuentasContablesToAnexosLiqu(rsCtasAdic!cTpoCuentas, "1")
        nSaldoDiario1 = 0
        If rs!cCtaContCod <> "" Then
            iFilaAnx = iFilaAnxMain
            Do While Not rs.EOF
                If rs!TipoDato = "Diario" Then
                    nSaldoDiarioParam = DAnxVal.ObtieneCtaSaldoDiario(rs!cCtaContCod, pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio) * (IIf(rs!cSigno = "+", 1, -1))
                Else
                    nSaldoDiarioParam = ObtenerCtaContSaldoBalanceDiario(rs!cCtaContCod, pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario) * (IIf(rs!cSigno = "+", 1, -1))
                End If
                xlHoja1.Cells(iFilaAnx, 13) = rs!TipoDato
                xlHoja1.Cells(iFilaAnx, 14) = Mid(rs!cCtaContCod, 1, 2) & "0" & Mid(rs!cCtaContCod, 4, Len(rs!cCtaContCod))
                xlHoja1.Cells(iFilaAnx, 15) = nSaldoDiarioParam
                xlHoja1.Cells(iFilaAnx, 17) = rs!cDescrip
                xlHoja1.Range(xlHoja1.Cells(iFilaAnx, 13), xlHoja1.Cells(iFilaAnx, 17)).Font.Name = "Arial Narrow"
                xlHoja1.Range(xlHoja1.Cells(iFilaAnx, 13), xlHoja1.Cells(iFilaAnx, 17)).Font.Size = 10
                ExcelCuadro xlHoja1, 13, iFilaAnx, 17, CCur(iFilaAnx)
                nSaldoDiario1 = IIf(rs!cSigno = "+", nSaldoDiarioParam, 0) + nSaldoDiario1
                iFilaAnx = iFilaAnx + 1
                rs.MoveNext
            Loop
        End If
        Set rs = Nothing
        'ME
        Set rs = DAnxVal.ObtieneCuentasContablesToAnexosLiqu(rsCtasAdic!cTpoCuentas, "2")
        nSaldoDiario2 = 0
        If rs!cCtaContCod <> "" Then
            iFilaAnx = iFilaAnxMain
            Do While Not rs.EOF
                If rs!TipoDato = "Diario" Then
                    nSaldoDiarioParam = DAnxVal.ObtieneCtaSaldoDiario(rs!cCtaContCod, pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio) * (IIf(rs!cSigno = "+", 1, -1))
                Else
                    nSaldoDiarioParam = Round(ObtenerCtaContSaldoBalanceDiario(rs!cCtaContCod, pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) / lnTipoCambioBalanceAnterior, 2) * (IIf(rs!cSigno = "+", 1, -1))
                End If
                xlHoja1.Cells(iFilaAnx, 16) = nSaldoDiarioParam
                nSaldoDiario2 = IIf(rs!cSigno = "+", nSaldoDiarioParam, 0) + nSaldoDiario2
                iFilaAnx = iFilaAnx + 1
                rs.MoveNext
            Loop
        End If
        Set rs = Nothing
        iFilaAnxMain = iFilaAnx
        If iFilaCtaLiqu = 64 Then
        iFilaCtaLiqu = 65
        End If
        If iFilaCtaLiqu = 66 Then
        iFilaCtaLiqu = 67
        End If
        xlHoja1.Cells(iFilaCtaLiqu, 3) = nSaldoDiario1
        xlHoja1.Cells(iFilaCtaLiqu, 4) = nSaldoDiario2
        iFilaCtaLiqu = iFilaCtaLiqu + 1
        rsCtasAdic.MoveNext
    Loop
    End If
    Set rsCtasAdic = Nothing
    
'    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(55, 3), 2) 'VAPA20171101 ANPS COMENTADO
'    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(55, 4), 2) 'VAPA20171101 ANPS COMENTADO
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(61, 3), 2)  'ANPS
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(61, 4), 2)  'ANPS

    
    'JIPR20200824
    Dim oValor As New DAnexoRiesgos
    Dim rsvalor As New ADODB.Recordset
    Dim nTotalMN As Currency
    
    Set rsvalor = oValor.AdeudadoReactiva30(pdFecha, 1)
    If Not (rsvalor.EOF And rsvalor.BOF) Then
        nTotalMN = rsvalor!nSaldo
    Else
        nTotalMN = 0
    End If
'    xlHoja1.Cells(56, 3) = nTotalMN ANPS COMENTADO
    xlHoja1.Cells(62, 3) = nTotalMN  'ANPS
    'JIPR20200824
    
'    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(56, 4) * 0, 2) 'NAGL20190509 ANPS COMENTADO
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(62, 4) * 0, 2)  'ANPS
    
'    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(57, 3) * 0.15, 2) 'NAGL20190509 ANPS COMENTADO
'    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(57, 4) * 0.15, 2) 'NAGL20190509 ANPS COMENTADO
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(63, 3) * 0.15, 2)  'ANPS
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(63, 4) * 0.15, 2)  'ANPS

'AGREGAR FILA 64 Cuentas por pagar - operaciones de reporte con valores de Gobiernos del exterior de Riesgo II ? 30 d (26)



'    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(58, 3) * 0.25, 2) 'NAGL20190509 ANPS COMENTADO
'    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(58, 4) * 0.25, 2) 'NAGL20190509 ANPS COMENTADO
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(65, 3) * 0.25, 2)  'ANPS
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(65, 4) * 0.25, 2)  'ANPS

'    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(59, 3), 2) 'NAGL20190509 ANPS COMENTADO
'    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(59, 4), 2) 'NAGL20190509 ANPS COMENTADO
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(67, 3), 2)  'ANPS
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(67, 4), 2)  'ANPS

    xlHoja1.Range(xlHoja1.Cells(iFilaAnxParam, 15), xlHoja1.Cells(iFilaAnx - 1, 16)).NumberFormat = "#,##0.00;-#,##0.00"
'    xlHoja1.Range(xlHoja1.Cells(55, 3), xlHoja1.Cells(59, 4)).NumberFormat = "#,##0.00;-#,##0.00" ANPS COMENTADO
    xlHoja1.Range(xlHoja1.Cells(61, 3), xlHoja1.Cells(67, 4)).NumberFormat = "#,##0.00;-#,##0.00" 'ANPS
    
    'Cuentas por pagar - otros (17)
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(55, 3), 1, "3500", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(55, 4), 1, "3500", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "3", xlHoja1.Cells(55, 6), 1, "3500", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "4", xlHoja1.Cells(55, 7), 1, "3500", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(61, 3), 1, "3500", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(61, 4), 1, "3500", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "3", xlHoja1.Cells(61, 6), 1, "3500", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "4", xlHoja1.Cells(61, 7), 1, "3500", "B1") 'ANPS
    
    'Cuentas por pagar - operaciones de reporte con valores del BCRP y Gobierno Central, o con el BCRP como contraparte (25)
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(56, 3), 1, "3510", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(56, 4), 1, "3510", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "3", xlHoja1.Cells(56, 6), 1, "3510", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "4", xlHoja1.Cells(56, 7), 1, "3510", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(62, 3), 1, "3510", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(62, 4), 1, "3510", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "3", xlHoja1.Cells(62, 6), 1, "3510", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "4", xlHoja1.Cells(62, 7), 1, "3510", "B1") 'ANPS
    
    'Cuentas por pagar - operaciones de reporte con valores de Gobiernos del exterior (26)
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(57, 3), 1, "3520", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(57, 4), 1, "3520", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "3", xlHoja1.Cells(57, 6), 1, "3520", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "4", xlHoja1.Cells(57, 7), 1, "3520", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(63, 3), 1, "3520", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(63, 4), 1, "3520", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "3", xlHoja1.Cells(63, 6), 1, "3520", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "4", xlHoja1.Cells(63, 7), 1, "3520", "B1") 'ANPS
    
    
'AGREGAR FILA 64 Cuentas por pagar - operaciones de reporte con valores de Gobiernos del exterior de Riesgo II ? 30 d (26)
    
    
    
    'Cuentas por pagar - operaciones de reporte con bonos corporativos emitidos por empresas privadas del sector no financiero (26)
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(58, 3), 1, "3530", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(58, 4), 1, "3530", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "3", xlHoja1.Cells(58, 6), 1, "3530", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "4", xlHoja1.Cells(58, 7), 1, "3530", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(65, 3), 1, "3530", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(65, 4), 1, "3530", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "3", xlHoja1.Cells(65, 6), 1, "3530", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "4", xlHoja1.Cells(65, 7), 1, "3530", "B1") 'ANPS
    
    'Cuentas por pagar - operaciones de reporte con otros valores (27)
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(59, 3), 1, "3540", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(59, 4), 1, "3540", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "3", xlHoja1.Cells(59, 6), 1, "3540", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "4", xlHoja1.Cells(59, 7), 1, "3540", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(67, 3), 1, "3540", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(67, 4), 1, "3540", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "3", xlHoja1.Cells(67, 6), 1, "3540", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "4", xlHoja1.Cells(67, 7), 1, "3540", "B1") 'ANPS
    
    
    'Cuentas por pagar - operaciones de reporte con otros valores (27)
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(59, 3), 1, "3550", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(59, 4), 1, "3550", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "3", xlHoja1.Cells(59, 6), 1, "3550", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "4", xlHoja1.Cells(59, 7), 1, "3550", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(67, 3), 1, "3550", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(67, 4), 1, "3550", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "3", xlHoja1.Cells(67, 6), 1, "3550", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "4", xlHoja1.Cells(67, 7), 1, "3550", "B1") 'ANPS
    '***************END NAGL 20190428****************************
     
     'Posiciones pasivas en derivados - Delivery (12)
'    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "1", xlHoja1.Cells(61, 3), 1, "3600", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "2", xlHoja1.Cells(61, 4), 1, "3600", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "3", xlHoja1.Cells(61, 6), 1, "3600", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "4", xlHoja1.Cells(61, 7), 1, "3600", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "1", xlHoja1.Cells(69, 3), 1, "3600", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "2", xlHoja1.Cells(69, 4), 1, "3600", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "3", xlHoja1.Cells(69, 6), 1, "3600", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "4", xlHoja1.Cells(69, 7), 1, "3600", "B1") 'ANPS
    
    
    'Líneas de crédito no utilizadas y créditos concedidos no desembolsados - personas naturales y jurídicas sin fines de lucro (18)
'    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "1", xlHoja1.Cells(62, 3), 1, "3700", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "2", xlHoja1.Cells(62, 4), 1, "3700", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "3", xlHoja1.Cells(62, 6), 1, "3700", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "4", xlHoja1.Cells(62, 7), 1, "3700", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "1", xlHoja1.Cells(70, 3), 1, "3700", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "2", xlHoja1.Cells(70, 4), 1, "3700", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "3", xlHoja1.Cells(70, 6), 1, "3700", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "4", xlHoja1.Cells(70, 7), 1, "3700", "B1") 'ANPS
    
    'Líneas de crédito no utilizadas y créditos concedidos no desembolsados - personas jurídicas con fines de lucro (18)
'    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "1", xlHoja1.Cells(63, 3), 1, "3800", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "2", xlHoja1.Cells(63, 4), 1, "3800", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "3", xlHoja1.Cells(63, 6), 1, "3800", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "4", xlHoja1.Cells(63, 7), 1, "3800", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "1", xlHoja1.Cells(71, 3), 1, "3800", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "2", xlHoja1.Cells(71, 4), 1, "3800", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "3", xlHoja1.Cells(71, 6), 1, "3800", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "4", xlHoja1.Cells(71, 7), 1, "3800", "B1") 'ANPS
    
    'Créditos concedidos no desembolsados - hipoteca inversa (18A) 'Agregado by NAGL Según RFC1912050003
'    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "1", xlHoja1.Cells(64, 3), 1, "3850", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "2", xlHoja1.Cells(64, 4), 1, "3850", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "3", xlHoja1.Cells(64, 6), 1, "3850", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "4", xlHoja1.Cells(64, 7), 1, "3850", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "1", xlHoja1.Cells(72, 3), 1, "3850", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "2", xlHoja1.Cells(72, 4), 1, "3850", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "3", xlHoja1.Cells(72, 6), 1, "3850", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "4", xlHoja1.Cells(72, 7), 1, "3850", "B1") 'ANPS
    
    '**************NAGL ERS079-2016 20170407 Tasa de Encaje
    nSaldoDiario1 = (oDbalanceCont.ObtenerParamEncDiarioxCodigo("32") / 100)
    nSaldoDiario2 = (oDbalanceCont.ObtenerParamEncDiarioxCodigo("33") / 100)
    xlHoja1.Cells(10, 10) = nSaldoDiario1
    xlHoja1.Cells(10, 11) = nSaldoDiario2
    
    xlHoja1.Range(xlHoja1.Cells(10, 10), xlHoja1.Cells(10, 11)).NumberFormat = "#,##0.00;-#,##0.00"
   
    xlHoja1.Range(xlHoja1.Cells(13, 3), xlHoja1.Cells(13, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnTotal1MN = lnTotal1MN + Round(xlHoja1.Cells(13, 3), 2) 'VAPA20171116
    lnTotal1ME = lnTotal1ME + Round(xlHoja1.Cells(13, 4), 2) 'VAPA20171116
    'Encaje liberado por los flujos salientes (4)
    Call oDbalanceCont.InsertaDetallaReporte15A(4, pdFecha, "1", xlHoja1.Cells(13, 3), 1, "400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(4, pdFecha, "2", xlHoja1.Cells(13, 4), 1, "400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(4, pdFecha, "3", xlHoja1.Cells(13, 6), 1, "400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(4, pdFecha, "4", xlHoja1.Cells(13, 7), 1, "400", "B1")
    
    'Ingreso de Tipo Cambio Contable SBS
    'Cambio de Posición by NAGL 20191211 Según RFC201912050003 a Partir de la fila 63
    nSaldoDiario1 = DAnxVal.ObtieneTipoCambioContableSBS(pdFecha)
'    xlHoja1.Cells(70, 5) = IIf(nSaldoDiario1 <> 0, Format(nSaldoDiario1, "#,##0.000"), xlHoja1.Cells(70, 5)) ANPS COMENTADO
    xlHoja1.Cells(78, 5) = IIf(nSaldoDiario1 <> 0, Format(nSaldoDiario1, "#,##0.000"), xlHoja1.Cells(78, 5)) 'ANPS
    
    'Para insertar los ratios antes del posible Intercambio de Liquidez (4001)Nuevo
'    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "1", xlHoja1.Cells(66, 3), 1, "4001", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "2", xlHoja1.Cells(66, 4), 1, "4001", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "3", xlHoja1.Cells(66, 6), 1, "4001", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "4", xlHoja1.Cells(66, 7), 1, "4001", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "1", xlHoja1.Cells(74, 3), 1, "4001", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "2", xlHoja1.Cells(74, 4), 1, "4001", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "3", xlHoja1.Cells(74, 6), 1, "4001", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "4", xlHoja1.Cells(74, 7), 1, "4001", "B1") 'ANPS
    
    
    '****Para el InterCambio de Liquidez
'    Call ObtieneInterCambioLiquidezMNyME(xlHoja1.Application, pdFecha, xlHoja1.Cells(66, 6), xlHoja1.Cells(66, 7), xlHoja1.Cells(70, 5), xlHoja1.Cells(23, 6), xlHoja1.Cells(23, 7), xlHoja1.Cells(41, 6), xlHoja1.Cells(41, 7), xlHoja1.Cells(65, 6), xlHoja1.Cells(65, 7)) ANPS COMENTADO
'    lnTotal1MN = xlHoja1.Cells(23, 6) 'NAGL20190507 ANPS COMENTADO
'    lnTotal1ME = xlHoja1.Cells(23, 7) 'NAGL20190507 ANPS COMENTADO
    Call ObtieneInterCambioLiquidezMNyME(xlHoja1.Application, pdFecha, xlHoja1.Cells(74, 6), xlHoja1.Cells(74, 7), xlHoja1.Cells(78, 5), xlHoja1.Cells(27, 6), xlHoja1.Cells(27, 7), xlHoja1.Cells(47, 6), xlHoja1.Cells(47, 7), xlHoja1.Cells(73, 6), xlHoja1.Cells(73, 7)) 'ANPS
    lnTotal1MN = xlHoja1.Cells(27, 6) 'ANPS
    lnTotal1ME = xlHoja1.Cells(27, 7) 'ANPS
    
    '***NAGL ERS006-2019 20190506******
    
    oBarra.Progress 9, TituloProgress, MensajeProgress, "", vbBlue
    
   '*****TOTALES*****'
'    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "1", xlHoja1.Cells(23, 3), 1, "1000", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "2", xlHoja1.Cells(23, 4), 1, "1000", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "3", xlHoja1.Cells(23, 6), 1, "1000", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "4", xlHoja1.Cells(23, 7), 1, "1000", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "1", xlHoja1.Cells(27, 3), 1, "1000", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "2", xlHoja1.Cells(27, 4), 1, "1000", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "3", xlHoja1.Cells(27, 6), 1, "1000", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "4", xlHoja1.Cells(27, 7), 1, "1000", "B1") 'ANPS
    
'    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "1", xlHoja1.Cells(41, 3), 1, "2000", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "2", xlHoja1.Cells(41, 4), 1, "2000", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "3", xlHoja1.Cells(41, 6), 1, "2000", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "4", xlHoja1.Cells(41, 7), 1, "2000", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "1", xlHoja1.Cells(47, 3), 1, "2000", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "2", xlHoja1.Cells(47, 4), 1, "2000", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "3", xlHoja1.Cells(47, 6), 1, "2000", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "4", xlHoja1.Cells(47, 7), 1, "2000", "B1") 'ANPS

'    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "1", xlHoja1.Cells(65, 3), 1, "3900", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "2", xlHoja1.Cells(65, 4), 1, "3900", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "3", xlHoja1.Cells(65, 6), 1, "3900", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "4", xlHoja1.Cells(65, 7), 1, "3900", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "1", xlHoja1.Cells(73, 3), 1, "3900", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "2", xlHoja1.Cells(73, 4), 1, "3900", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "3", xlHoja1.Cells(73, 6), 1, "3900", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "4", xlHoja1.Cells(73, 7), 1, "3900", "B1") 'ANPS
    
    '****END TOTALES***'
   
    '*****Restante*****'
    
    'Intercambio de liquidez USD por PEN (7)
'    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "1", xlHoja1.Cells(21, 3), 1, "800", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "2", xlHoja1.Cells(21, 4), 1, "800", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "3", xlHoja1.Cells(21, 6), 1, "800", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "4", xlHoja1.Cells(21, 7), 1, "800", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "1", xlHoja1.Cells(25, 3), 1, "800", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "2", xlHoja1.Cells(25, 4), 1, "800", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "3", xlHoja1.Cells(25, 6), 1, "800", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "4", xlHoja1.Cells(25, 7), 1, "800", "B1") 'ANPS
    
    'Intercambio de liquidez PEN por USD  (7)
'    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "1", xlHoja1.Cells(22, 3), 1, "900", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "2", xlHoja1.Cells(22, 4), 1, "900", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "3", xlHoja1.Cells(22, 6), 1, "900", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "4", xlHoja1.Cells(22, 7), 1, "900", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "1", xlHoja1.Cells(26, 3), 1, "900", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "2", xlHoja1.Cells(26, 4), 1, "900", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "3", xlHoja1.Cells(26, 6), 1, "900", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "4", xlHoja1.Cells(26, 7), 1, "900", "B1") 'ANPS
    
    '****END - Restante****'

    'VAPA20171101
    If (lnTotal3MN * 0.75) > lnTotal2MN Then
        lnAuxMN = lnTotal2MN
    Else
        lnAuxMN = lnTotal3MN * 0.75
    End If
    
    If (lnTotal3ME * 0.75) > lnTotal2ME Then
        lnAuxME = lnTotal2ME
    Else
        lnAuxME = lnTotal3ME * 0.75
    End If
    
    'lnRatioCoberturaLiquidezMN = (lnTotal1MN + Min(lnTotal2MN, lnTotal3MN * 0.75) / lnTotal3MN) * 100
    lnRatioCoberturaLiquidezMN = ((lnTotal1MN + lnAuxMN) / lnTotal3MN) * 100
    'lnRatioCoberturaLiquidezME = (lnTotal1ME + Min(lnTotal2ME, lnTotal3ME * 0.75) / lnTotal3ME) * 100
    lnRatioCoberturaLiquidezME = ((lnTotal1ME + lnAuxME) / lnTotal3ME) * 100
    lnRatioCoberturaLiquidezMN = Format(Round(lnRatioCoberturaLiquidezMN, 2), "#,##0.00;-#,##0.00")
    lnRatioCoberturaLiquidezME = Format(Round(lnRatioCoberturaLiquidezME, 2), "#,##0.00;-#,##0.00")
    
    InsertaLiquidezAlertaTempranaCobertura pdFechaAlerta, lnRatioCoberturaLiquidezMN, lnRatioCoberturaLiquidezME 'VAPA20171003
    'VAPA20171101
    
    oBarra.Progress 10, "ANEXO 15B: Ratio de Cobertura de Liquidez", "Generación Terminada", "", vbBlue
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
        xlHoja1.SaveAs App.path & lsArchivo1
        xlsAplicacion.Visible = True
        xlsAplicacion.Windows(1).Visible = True
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlHoja1 = Nothing
Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub


'**************************************************BEGIN NAGL 20170904 **************************************************'
Public Sub ReporteAnexo15BfromBitacora(psMoneda As String, pdFecha As Date, pdFechaSist As Date, psMesBalanceDiario As String, psAnioBalanceDiario As String, psRepBCRP As String)   'NAGL 20170415
 Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja  As String
    Dim lsArchivo As String
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim TituloProgress As String
    Dim MensajeProgress As String
    Dim oBarra As clsProgressBar
    Dim nprogress As Integer
    
    Dim oDbalanceCont As DbalanceCont
    Dim nSaldoDiario1 As Currency
    Dim nSaldoDiario2 As Currency
    Dim pdFechaFinDeMes As Date
    Dim pdFechaFinDeMesMA As Date
    Dim dFechaAnte As Date
    Dim ldFechaPro As Date
    Dim pdFechaBalanceDiario As Date
    Dim nDia As Integer
    Dim oCambio As nTipoCambio
    Dim lnTipoCambioFC As Currency, pnTipoCambio As Currency
    Dim lnTipoCambioProceso As Currency
    Dim lnTipoCambioBalanceAnterior As Currency
    Dim nTipoCambioAn As Currency
    Dim loRs As ADODB.Recordset
    Dim lnSubastasMN As Currency
    Dim lnSubastasME As Currency
    Dim lnOtrosDepMND1y2_C0 As Currency
    Dim lnOtrosDepMND4aM_C0 As Currency
    Dim lnOtrosDepMND1y2_C1 As Currency
    Dim lnOtrosDepMED1y2_C0 As Currency
    Dim lnOtrosDepMED4aM_C0 As Currency
    Dim lnOtrosDepMED1y2_C1 As Currency
    Dim lnSubastasMND1y2_C1 As Currency
    Dim lnSubastasMND1y2_C0 As Currency
    Dim lnSubastasMND4aM_C0 As Currency
    Dim lnSubastasMED1y2_C1 As Currency
    Dim lnSubastasMED1y2_C0 As Currency
    Dim lnSubastasMED4aM_C0 As Currency
    
    Dim lnSubastasMED3o3_C0 As Currency
    Dim lnSubastasMND3o3_C0 As Currency
    Dim lnSubastasMED3o3_C1 As Currency
    Dim lnSubastasMND3o3_C1 As Currency
    
    Dim lnOtrosDepMND3o3_C0 As Currency
    Dim lnOtrosDepMED3o3_C0 As Currency
    Dim lnOtrosDepMND3o3_C1 As Currency
    Dim lnOtrosDepMED3o3_C1 As Currency
    
    Dim lnSubastasMND4aM_C1 As Currency
    Dim lnSubastasMED4aM_C1 As Currency
    
    Dim nTotalObligSugEncajMN As Currency
    Dim nTotalTasaBaseEncajMN As Currency
    Dim nTotalObligSugEncajME As Currency
    Dim nTotalTasaBaseEncajME  As Currency
    Dim lnTotalObligacionesAlDiaMN As Currency
    Dim lnTotalObligacionesAlDiaME  As Currency
    Dim nTotalTasaBaseEncajMNDiario As Currency
    Dim nTotalTasaBaseEncajMEDiario  As Currency
    Dim nTotalTasaBaseEncajMN_DADiario As Currency
    Dim nTotalTasaBaseEncajME_DADiario As Currency
    Dim nTotalTasaBaseEncajMN_DA  As Currency
    Dim nTotalObligSugEncajMN_DA As Currency
    Dim nTotalTasaBaseEncajME_DA As Currency
    Dim nTotalObligSugEncajME_DA As Currency
    Dim ix As Integer, rx As Integer
    Dim nSubValor1 As Currency
    Dim nSubValor2 As Currency
    'VAPA 20171101 ***********************************
    Dim lnTotal1MN As Double
    Dim lnTotal1ME As Double
    Dim lnTotal2MN As Double
    Dim lnTotal2ME As Double
    Dim lnTotal3MN As Double
    Dim lnTotal3ME As Double
    Dim lnRatioCoberturaLiquidezMN As Double
    Dim lnRatioCoberturaLiquidezME As Double
    Dim lnAuxMN As Double
    Dim lnAuxME As Double
    Dim pdFechaAlerta As Date
    'VAPA END ****************************************
    Dim DAnxVal As New DAnexoRiesgos '20190503
    
On Error GoTo GeneraExcelErr

    Set oBarra = New clsProgressBar
    Unload Me
    oBarra.ShowForm frmReportes
    oBarra.Max = 10
    nprogress = 0
    oBarra.Progress nprogress, "ANEXO 15B: Ratio de Cobertura de Liquidez", "GENERANDO EL ARCHIVO", "", vbBlue
    TituloProgress = "ANEXO 15B: Ratio de Cobertura de Liquidez"
    MensajeProgress = "GENERANDO EL ARCHIVO"
    
    pdFechaFinDeMes = DateAdd("d", -Day(DateAdd("m", 1, pdFecha)), DateAdd("m", 1, pdFecha))
    pdFechaFinDeMesMA = DateAdd("d", -Day(pdFechaFinDeMes), pdFechaFinDeMes)
    nDia = Day(pdFecha)
    Set oDbalanceCont = New DbalanceCont
    
    Set oCambio = New nTipoCambio
    pdFechaAlerta = pdFecha 'vapa20171117
    
    If CInt(psMesBalanceDiario) < 10 Then
        pdFechaBalanceDiario = CDate("01" & "/" & "0" & CStr(CInt(psMesBalanceDiario) + 1) & "/" & psAnioBalanceDiario)
    ElseIf CInt(psMesBalanceDiario) = 12 Then
        pdFechaBalanceDiario = CDate("01" & "/" & "01" & "/" & CStr(CInt(psAnioBalanceDiario) + 1))
    Else
        pdFechaBalanceDiario = CDate("01" & "/" & CStr(CInt(psMesBalanceDiario) + 1) & "/" & psAnioBalanceDiario)
    End If
    
    lnTipoCambioBalanceAnterior = Format(oCambio.EmiteTipoCambio(pdFechaBalanceDiario, TCFijoDia), "#,##0.0000")
    
    If Month(pdFecha) = Month(DateAdd("d", 1, pdFecha)) Then
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(pdFecha, TCFijoDia), "#,##0.0000")
    Else
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(DateAdd("d", -1, pdFecha), TCFijoDia), "#,##0.0000")
    End If
    pnTipoCambio = lnTipoCambioFC
    
    oBarra.Progress 1, TituloProgress, MensajeProgress, "", vbBlue
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    'lsArchivo = "ANEXO_15B"
    lsArchivo = "ANEXO_15B_Bitácora" 'NAGL 20191108
    'Primera Hoja ******************************************************
    lsNomHoja = "Anx15B"
    '*******************************************************************
    lsArchivo1 = "\spooler\ANEXO_15B_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
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
    xlHoja1.Cells(4, 2) = "AL " & Format(pdFecha, "YYYY/MM/DD")
    
    'ACTIVOS LÍQUIDOS DE ALTA CALIDAD
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "200")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "200")
    
    xlHoja1.Cells(10, 3) = nSaldoDiario1
    xlHoja1.Cells(10, 4) = nSaldoDiario2
    
    xlHoja1.Range(xlHoja1.Cells(10, 3), xlHoja1.Cells(10, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    lnTotal1MN = xlHoja1.Cells(10, 3) 'VAPA20171101
    lnTotal1ME = xlHoja1.Cells(10, 4) 'VAPA20171101
    oBarra.Progress 2, TituloProgress, MensajeProgress, "", vbBlue

    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "3", "B1", "300")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "4", "B1", "300")
    
    xlHoja1.Cells(11, 3) = nSaldoDiario1
    xlHoja1.Cells(11, 4) = nSaldoDiario2
    lnTotal1MN = lnTotal1MN + xlHoja1.Cells(11, 3) 'VAPA20171101
    lnTotal1ME = lnTotal1ME + xlHoja1.Cells(11, 4) 'VAPA20171101
    
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "301")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "301")
    
    xlHoja1.Cells(12, 3) = nSaldoDiario1
    xlHoja1.Cells(12, 4) = nSaldoDiario2
    lnTotal1MN = lnTotal1MN + xlHoja1.Cells(12, 3) 'VAPA20171101
    lnTotal1ME = lnTotal1ME + xlHoja1.Cells(12, 4) 'VAPA20171101
    
    oBarra.Progress 4, TituloProgress, MensajeProgress, "", vbBlue
    
    nSaldoDiario1 = oDbalanceCont.ObtenerSumaValorRepresentativos(pdFecha, "C[BD]")
    nSaldoDiario2 = oDbalanceCont.ObtenerSumaValorRepresentativos(pdFecha, "LT")
    
    xlHoja1.Cells(14, 3) = nSaldoDiario1
    xlHoja1.Cells(15, 3) = nSaldoDiario2
    
    lnTotal1MN = lnTotal1MN + xlHoja1.Cells(14, 3) 'VAPA20171101
    lnTotal1MN = lnTotal1MN + xlHoja1.Cells(15, 3) 'VAPA20171101
    
    xlHoja1.Range(xlHoja1.Cells(11, 3), xlHoja1.Cells(11, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(12, 3), xlHoja1.Cells(12, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(14, 3), xlHoja1.Cells(14, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(15, 3), xlHoja1.Cells(15, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    lnTotal1ME = lnTotal1ME + xlHoja1.Cells(14, 4) 'VAPA20171101
    lnTotal1ME = lnTotal1ME + xlHoja1.Cells(15, 4) 'VAPA20171101
    
    oBarra.Progress 5, TituloProgress, MensajeProgress, "", vbBlue
    
    'FLUJOS ENTRANTES 30 DÍAS
    'Disponible
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "1200")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "1200")
    xlHoja1.Cells(28, 3) = nSaldoDiario1
    xlHoja1.Cells(28, 4) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(28, 3), xlHoja1.Cells(28, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnTotal2MN = Round(xlHoja1.Cells(28, 3), 2) 'VAPA20171101
    lnTotal2ME = Round(xlHoja1.Cells(28, 4), 2) 'VAPA20171101
    
    'Fondos Disponibles en Empresas del SFN
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "1300")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "1300")
    xlHoja1.Cells(29, 3) = nSaldoDiario1
    xlHoja1.Cells(29, 4) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(29, 3), xlHoja1.Cells(29, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(29, 3), 2) 'VAPA 20171101
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(29, 4), 2) 'VAPA 20171101
    
    'Créditos
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "1600")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "1600")
    xlHoja1.Cells(32, 3) = nSaldoDiario1
    xlHoja1.Cells(32, 4) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(32, 3), xlHoja1.Cells(32, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(32, 3) * 0.5, 2) 'VAPA 20171101
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(32, 4) * 0.5, 2) 'VAPA 20171101
    
    '***Sección Cuentas por Cobrar
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "1800")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "1800")
    xlHoja1.Cells(34, 3) = nSaldoDiario1
    xlHoja1.Cells(34, 4) = nSaldoDiario2
    
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "1810")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "1810")
    xlHoja1.Cells(35, 3) = nSaldoDiario1
    xlHoja1.Cells(35, 4) = nSaldoDiario2

    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "1820")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "1820")
    xlHoja1.Cells(36, 3) = nSaldoDiario1
    xlHoja1.Cells(36, 4) = nSaldoDiario2

    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "1830")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "1830")
    xlHoja1.Cells(37, 3) = nSaldoDiario1
    xlHoja1.Cells(37, 4) = nSaldoDiario2
    
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "1840")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "1840")
    xlHoja1.Cells(38, 3) = nSaldoDiario1
    xlHoja1.Cells(38, 4) = nSaldoDiario2

    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(34, 3) * 0.8, 2) 'VAPA 20171101
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(34, 4) * 0.8, 2) 'VAPA 20171101
    
    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(35, 3) * 0, 2) 'NAGL 20190509
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(35, 4) * 0, 2) 'NAGL 20190509
    
    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(36, 3) * 0.15, 2) 'NAGL 20190509
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(36, 4) * 0.15, 2) 'NAGL 20190509
    
    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(37, 3) * 0.25, 2) 'NAGL 20190509
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(37, 4) * 0.25, 2) 'NAGL 20190509
    
    lnTotal2MN = lnTotal2MN + Round(xlHoja1.Cells(38, 3), 2) 'NAGL 20190509
    lnTotal2ME = lnTotal2ME + Round(xlHoja1.Cells(38, 4), 2) 'NAGL 20190509

    xlHoja1.Range(xlHoja1.Cells(34, 3), xlHoja1.Cells(38, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    oBarra.Progress 6, TituloProgress, MensajeProgress, "", vbBlue
    '***FONDEO
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "2200")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "2200")
    xlHoja1.Cells(43, 3) = nSaldoDiario1
    xlHoja1.Cells(43, 4) = nSaldoDiario2
    
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "2300")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "2300")
    xlHoja1.Cells(44, 3) = nSaldoDiario1
    xlHoja1.Cells(44, 4) = nSaldoDiario2
    
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "2400")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "2400")
    xlHoja1.Cells(45, 3) = nSaldoDiario1
    xlHoja1.Cells(45, 4) = nSaldoDiario2
    
    lnTotal3MN = Round(xlHoja1.Cells(43, 3) * 0.075, 2) 'VAPA20171101
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(44, 3) * 0.15, 2) 'VAPA20171101
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(45, 3) * 0.15, 2) 'VAPA20171101
    
    lnTotal3ME = Round(xlHoja1.Cells(43, 4) * 0.075, 2) 'VAPA20171101
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(44, 4) * 0.15, 2) 'VAPA20171101
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(45, 4) * 0.15, 2) 'VAPA20171101
    
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "2600")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "2600")
    xlHoja1.Cells(47, 3) = nSaldoDiario1
    xlHoja1.Cells(47, 4) = nSaldoDiario2
    
    xlHoja1.Range(xlHoja1.Cells(43, 3), xlHoja1.Cells(45, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(47, 3), xlHoja1.Cells(47, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(47, 3) * 0.3, 2)  'VAPA20171101
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(47, 4) * 0.3, 2) 'VAPA20171101
    
    oBarra.Progress 7, TituloProgress, MensajeProgress, "", vbBlue
    'Otras Oblig.Inst.Rec <= 30 Días
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "2800")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "2800")
    xlHoja1.Cells(48, 3) = nSaldoDiario1
    xlHoja1.Cells(48, 4) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(48, 3), xlHoja1.Cells(48, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(48, 3), 2)   'VAPA20171101
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(48, 4), 2)  'VAPA20171101
    
    'Depósitos de Empr. SF y OFI
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "3000")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "3000")
    xlHoja1.Cells(50, 3) = nSaldoDiario1
    xlHoja1.Cells(50, 4) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(50, 3), xlHoja1.Cells(50, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(50, 3), 2) 'VAPA20171101
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(50, 4), 2) 'VAPA20171101
    
    'Adeudos y Obligac.Financ <= 30 días
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "3100")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "3100")
    xlHoja1.Cells(51, 3) = nSaldoDiario1
    xlHoja1.Cells(51, 4) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(51, 3), xlHoja1.Cells(51, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(51, 3), 2) 'VAPA20171101
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(51, 4), 2) 'VAPA20171101
    
    'Adeudos y Obligac.Financ de 31 a 90 Días
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "3200")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "3200")
    xlHoja1.Cells(52, 3) = nSaldoDiario1
    xlHoja1.Cells(52, 4) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(52, 3), xlHoja1.Cells(52, 4)).NumberFormat = "#,##0.00;-#,##0.00"

    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(52, 3) * 0.3, 2) 'VAPA20171101
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(52, 4) * 0.3, 2) 'VAPA20171101

    oBarra.Progress 8, TituloProgress, MensajeProgress, "", vbBlue
    
    '***Sección Cuentas por Pagar
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "3500")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "3500")
    xlHoja1.Cells(55, 3) = nSaldoDiario1
    xlHoja1.Cells(55, 4) = nSaldoDiario2
    
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "3510")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "3510")
    xlHoja1.Cells(56, 3) = nSaldoDiario1
    xlHoja1.Cells(56, 4) = nSaldoDiario2
    
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "3520")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "3520")
    xlHoja1.Cells(57, 3) = nSaldoDiario1
    xlHoja1.Cells(57, 4) = nSaldoDiario2
    
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "3530")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "3530")
    xlHoja1.Cells(58, 3) = nSaldoDiario1
    xlHoja1.Cells(58, 4) = nSaldoDiario2
    
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "3540")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "3540")
    xlHoja1.Cells(59, 3) = nSaldoDiario1
    xlHoja1.Cells(59, 4) = nSaldoDiario2
    
    xlHoja1.Range(xlHoja1.Cells(55, 3), xlHoja1.Cells(59, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(55, 3), 2) 'VAPA20171101
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(55, 4), 2) 'VAPA20171101
        
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(56, 3) * 0, 2) 'NAGL20190509
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(56, 4) * 0, 2) 'NAGL20190509
    
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(57, 3) * 0.15, 2) 'NAGL20190509
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(57, 4) * 0.15, 2) 'NAGL20190509
    
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(58, 3) * 0.25, 2) 'NAGL20190509
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(58, 4) * 0.25, 2) 'NAGL20190509
    
    lnTotal3MN = lnTotal3MN + Round(xlHoja1.Cells(59, 3), 2) 'NAGL20190509
    lnTotal3ME = lnTotal3ME + Round(xlHoja1.Cells(59, 4), 2) 'NAGL20190509
    '*************
    
    nSaldoDiario1 = (ObtenerParamEncDiarioxCodigoxFechaIng(pdFecha, "32") / 100)
    nSaldoDiario2 = (ObtenerParamEncDiarioxCodigoxFechaIng(pdFecha, "33") / 100)
    xlHoja1.Cells(10, 10) = nSaldoDiario1
    xlHoja1.Cells(10, 11) = nSaldoDiario2
    
    xlHoja1.Range(xlHoja1.Cells(10, 10), xlHoja1.Cells(10, 11)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(13, 3), xlHoja1.Cells(13, 4)).NumberFormat = "#,##0.00;-#,##0.00"
        
    lnTotal1MN = lnTotal1MN + Round(xlHoja1.Cells(13, 3), 2) 'VAPA20171116
    lnTotal1ME = lnTotal1ME + Round(xlHoja1.Cells(13, 4), 2) 'VAPA20171116

    'Ingreso de Tipo Cambio Contable SBS
    nSaldoDiario1 = DAnxVal.ObtieneTipoCambioContableSBS(pdFecha)
    xlHoja1.Cells(70, 5) = IIf(nSaldoDiario1 <> 0, Format(nSaldoDiario1, "#,##0.000"), xlHoja1.Cells(70, 5))
    
    'Sección Intercambio de Líquidez
    nSaldoDiario1 = (oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "B1", "900")) * -1
    nSaldoDiario2 = (oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "B1", "800")) * -1
    xlHoja1.Cells(70, 3) = IIf(nSaldoDiario1 <> 0, nSaldoDiario1, xlHoja1.Cells(70, 3))
    xlHoja1.Cells(70, 4) = IIf(nSaldoDiario2 <> 0, nSaldoDiario2, xlHoja1.Cells(70, 4))
    
    lnTotal1MN = xlHoja1.Cells(23, 6) 'NAGL20190507
    lnTotal1ME = xlHoja1.Cells(23, 7) 'NAGL20190507
    'CargaValidacionCtaContIntDeveng15B xlHoja1.Application, pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio, psMesBalanceDiario, psAnioBalanceDiario  '***NAGL ERS 079-2017 20180123
    
    oBarra.Progress 9, TituloProgress, MensajeProgress, "", vbBlue
    
    If (lnTotal3MN * 0.75) > lnTotal2MN Then
        lnAuxMN = lnTotal2MN
    Else
        lnAuxMN = lnTotal3MN * 0.75
    End If
    
    If (lnTotal3ME * 0.75) > lnTotal2ME Then
        lnAuxME = lnTotal2ME
    Else
        lnAuxME = lnTotal3ME * 0.75
    End If
    
    'lnRatioCoberturaLiquidezMN = (lnTotal1MN + Min(lnTotal2MN, lnTotal3MN * 0.75) / lnTotal3MN) * 100
    lnRatioCoberturaLiquidezMN = ((lnTotal1MN + lnAuxMN) / lnTotal3MN) * 100
    'lnRatioCoberturaLiquidezME = (lnTotal1ME + Min(lnTotal2ME, lnTotal3ME * 0.75) / lnTotal3ME) * 100
    lnRatioCoberturaLiquidezME = ((lnTotal1ME + lnAuxME) / lnTotal3ME) * 100
    lnRatioCoberturaLiquidezMN = Format(Round(lnRatioCoberturaLiquidezMN, 2), "#,##0.00;-#,##0.00")
    lnRatioCoberturaLiquidezME = Format(Round(lnRatioCoberturaLiquidezME, 2), "#,##0.00;-#,##0.00")
    
    InsertaLiquidezAlertaTempranaCobertura pdFechaAlerta, lnRatioCoberturaLiquidezMN, lnRatioCoberturaLiquidezME 'VAPA20171003
    'VAPA20171101
    
    oBarra.Progress 10, "ANEXO 15B: Ratio de Cobertura de Liquidez", "Generación Terminada", "", vbBlue
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
        xlHoja1.SaveAs App.path & lsArchivo1
        xlsAplicacion.Visible = True
        xlsAplicacion.Windows(1).Visible = True
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlHoja1 = Nothing
Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub '**********************END NAGL 20170904***********************************'

Public Function ObtenerCtaContSaldoBalanceDiario(psCtaContCod As String, pdFecDiaProceso As Date, psMoneda As String, psMesBalanceDiario As String, psAnioBalanceDiario As String) As Currency
On Error GoTo ObtenerCtaContSaldoBalanceDiarioErr
   Dim oRS As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Dim psSql As String
   Set oRS = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ObtenerDatosBalanceDiarioReal '" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "', '" & psMesBalanceDiario & "', '" & psAnioBalanceDiario & "'"
   Set oRS = oConec.CargaRecordSet(psSql)
   If Not oRS.BOF And Not oRS.EOF Then
      Do While Not oRS.EOF
        nSaldo = oRS!nSaldoFinImporte
           oRS.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerCtaContSaldoBalanceDiario = nSaldo
Exit Function
ObtenerCtaContSaldoBalanceDiarioErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceDiario Method")
End Function 'NAGL 20170415

Public Function ObtenerParamEncDiarioxCodigoxFechaIng(pdFechaIng As Date, psCodigo As String) As Currency
On Error GoTo ObtenerParamEncDiarioxCodigoxFechaIngErr
   Dim oRS As ADODB.Recordset
   Dim oConec As DConecta
   Dim nValor As Currency
   Dim psSql As String
   Set oRS = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ParamEncDiarioxCodigoxFechaIng '" & Format(pdFechaIng, "yyyymmdd") & "','" & psCodigo & "'"
   Set oRS = oConec.CargaRecordSet(psSql)
   If Not oRS.BOF And Not oRS.EOF Then
      Do While Not oRS.EOF
        nValor = IIf(IsNull(oRS!nMonto), 0, oRS!nMonto)
           oRS.MoveNext
      Loop
   Else
    nValor = 0
   End If
   oConec.CierraConexion
   ObtenerParamEncDiarioxCodigoxFechaIng = nValor
Exit Function
ObtenerParamEncDiarioxCodigoxFechaIngErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerParamEncDiarioxCodigoxFechaIng Method")
End Function '***************NAGL ERS002-2017 20170904

Public Function ObtieneEfectivoDiaxRptConcentracionGastos(ByVal pdFecha As Date, ByVal pnMoneda As Moneda) As Currency
    Dim lsSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
On Error GoTo ErrobtenerPFyOvernyInvers
    If pnMoneda = gMonedaNacional Then
        lsSql = "Exec stp_sel_EfectivoDiarioMNxRptConcentraGastos '" & Format(pdFecha, "yyyymmdd") & "'"
    Else
        lsSql = "Exec stp_sel_EfectivoDiarioMExRptConcentraGastos '" & Format(pdFecha, "yyyymmdd") & "'"
    End If
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(lsSql)
    If Not rs.EOF Then
        ObtieneEfectivoDiaxRptConcentracionGastos = rs!nMonto
    End If
    oCon.CierraConexion
    Set rs = Nothing
    Set oCon = Nothing
    Exit Function
ErrobtenerPFyOvernyInvers:
    Call RaiseError(MyUnhandledError, "Recupera Efectivo Diario x Reporte Concentración Gastos")
End Function

Private Sub CargaValidacionCtaContIntDeveng15B(ByVal xlHoja1 As Excel.Application, ByVal pdFecha As Date, ByVal pdFechaBalanceDiario As Date, ByVal lnTipoCambioBalanceAnterior As Currency, ByVal pnTipoCambio As Currency, ByVal psMesBalanceDiario As String, ByVal psAnioBalanceDiario As String)
Dim pdFechaBalReal As Date
Dim DAnxRies As New DAnexoRiesgos
Dim rs As New ADODB.Recordset
Dim iFila As Integer
xlHoja1.Cells(14, 11) = "Balance a " & psMesBalanceDiario & "-" & psAnioBalanceDiario 'NAGL 20190416
Set rs = DAnxRies.CargaSaldosEstadistAnx15Ay15B(pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio)
iFila = 15
If Not rs.BOF And Not rs.EOF Then
    Do While Not rs.EOF
    xlHoja1.Cells(iFila, 9) = rs!cCtaContCod
    xlHoja1.Range(xlHoja1.Cells(iFila, 9), xlHoja1.Cells(iFila, 9)).Font.Bold = True
    xlHoja1.Cells(iFila, 10) = Format(rs!nSaldoImporte, "#,##0.00")
    xlHoja1.Cells(iFila, 11) = Round(ObtenerCtaContSaldoBalanceDiario(rs!cCtaContCod, pdFecha, Mid(rs!cCtaContCod, 3, 1), psMesBalanceDiario, psAnioBalanceDiario) / IIf(Mid(rs!cCtaContCod, 3, 1) = "1", 1, lnTipoCambioBalanceAnterior), 2)
    'ExcelCuadro xlHoja1, 9, iFila, 11, CCur(iFila)
    xlHoja1.Range(xlHoja1.Cells(iFila, 9), xlHoja1.Cells(iFila, 11)).Font.Name = "Arial Narrow"
    xlHoja1.Range(xlHoja1.Cells(iFila, 9), xlHoja1.Cells(iFila, 11)).Font.Size = 10
    xlHoja1.Range(xlHoja1.Cells(iFila, 9), xlHoja1.Cells(iFila, 11)).HorizontalAlignment = xlCenter
    iFila = iFila + 1
    rs.MoveNext
    Loop
End If 'Agregado by NAGL 20190416 ERS006-2019
xlHoja1.Range(xlHoja1.Cells(15, 10), xlHoja1.Cells(iFila, 11)).NumberFormat = "#,##0.00;-#,##0.00"

'xlHoja1.Cells(15, 10) = DAnxRies.ObtenerSaldEstadistAnx15Ay15B("111802", "1", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio)
'xlHoja1.Cells(15, 11) = ObtenerCtaContSaldoBalanceDiario("111802", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
'xlHoja1.Cells(16, 10) = Round(DAnxRies.ObtenerSaldEstadistAnx15Ay15B("112802", "2", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio), 2)
'xlHoja1.Cells(16, 11) = Round(ObtenerCtaContSaldoBalanceDiario("112802", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) / lnTipoCambioBalanceAnterior, 2)
'xlHoja1.Cells(17, 10) = DAnxRies.ObtenerSaldEstadistAnx15Ay15B("111803", "1", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio)
'xlHoja1.Cells(17, 11) = ObtenerCtaContSaldoBalanceDiario("111803", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
'xlHoja1.Cells(18, 10) = Round(DAnxRies.ObtenerSaldEstadistAnx15Ay15B("112803", "2", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio), 2)
'xlHoja1.Cells(18, 11) = Round(ObtenerCtaContSaldoBalanceDiario("112803", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) / lnTipoCambioBalanceAnterior, 2)
'xlHoja1.Cells(19, 10) = DAnxRies.ObtenerSaldEstadistAnx15Ay15B("211803 ", "1", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio)
'xlHoja1.Cells(19, 11) = ObtenerCtaContSaldoBalanceDiario("211803", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
'xlHoja1.Cells(20, 10) = Round(DAnxRies.ObtenerSaldEstadistAnx15Ay15B("212803", "2", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio), 2)
'xlHoja1.Cells(20, 11) = Round(ObtenerCtaContSaldoBalanceDiario("212803", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) / lnTipoCambioBalanceAnterior, 2)
'xlHoja1.Cells(21, 10) = DAnxRies.ObtenerSaldEstadistAnx15Ay15B("231803 ", "1", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio)
'xlHoja1.Cells(21, 11) = ObtenerCtaContSaldoBalanceDiario("231803", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
'xlHoja1.Cells(22, 10) = Round(DAnxRies.ObtenerSaldEstadistAnx15Ay15B("232803", "2", pdFecha, pdFechaBalanceDiario, lnTipoCambioBalanceAnterior, pnTipoCambio), 2)
'xlHoja1.Cells(22, 11) = Round(ObtenerCtaContSaldoBalanceDiario("232803", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) / lnTipoCambioBalanceAnterior, 2)
End Sub 'NAGL ERS079-2017 20180123

Private Sub ObtieneInterCambioLiquidezMNyME(ByVal xlHoja1 As Excel.Application, pdFecha As Date, pnRCLMN As Currency, pnRCLME As Currency, pnTpoCambSBS As Currency, pnImporteActMN As Currency, pnImporteActME As Currency, pnFlujoEntMN As Currency, pnFlujoEntME As Currency, pnFlujoSalMN As Currency, pnFlujoSalME As Currency)
Dim DAnxRies As New DAnexoRiesgos
Dim psRegValida As String
Dim rs As New ADODB.Recordset
Dim oDbalanceCont As New dBalanceCont
Dim psNivRiesgoMN As String, pnNivelRiesgoME As String
Dim pnRatioOrigMN As Currency, pnRatioOrigME As Currency

pnRatioOrigMN = xlHoja1.Cells(66, 6)
pnRatioOrigME = xlHoja1.Cells(66, 7)

psRegValida = DAnxRies.ValidaRatiosparaInterCambLiqMNyME(pnRCLMN, pnRCLME)
   Set rs = DAnxRies.ObtieneMontoInterCambioLiquidezMNyME(pnRCLMN, pnRCLME, pnTpoCambSBS, pnImporteActMN, pnImporteActME, pnFlujoEntMN, pnFlujoEntME, pnFlujoSalMN, pnFlujoSalME, psRegValida)
   If psRegValida = "1" Then
        If Not rs.BOF And Not rs.EOF Then
            Do While Not rs.EOF
                If rs!pnValor = "2" Then
'                   xlHoja1.Cells(70, 4) = Format(rs!pnInterCamLiquME, "#,##0.000") ANPS COMENTADO
'                   xlHoja1.Range(xlHoja1.Cells(70, 4), xlHoja1.Cells(70, 4)).Interior.ColorIndex = 6   ANPS COMENTADO
                   xlHoja1.Cells(78, 4) = Format(rs!pnInterCamLiquME, "#,##0.000") 'ANPS
                   xlHoja1.Range(xlHoja1.Cells(78, 4), xlHoja1.Cells(78, 4)).Interior.ColorIndex = 6 'ANPS
                Else
'                   xlHoja1.Cells(70, 3) = Format(rs!pnInterCamLiquMN, "#,##0.000") ANPS COMENTADO
'                   xlHoja1.Range(xlHoja1.Cells(70, 3), xlHoja1.Cells(70, 3)).Interior.ColorIndex = 6 ANPS COMENTADO
                   xlHoja1.Cells(78, 3) = Format(rs!pnInterCamLiquMN, "#,##0.000") 'ANPS
                   xlHoja1.Range(xlHoja1.Cells(78, 3), xlHoja1.Cells(78, 3)).Interior.ColorIndex = 6 'ANPS
                End If
'                xlHoja1.Cells(67, 6) = pnRatioOrigMN ANPS COMENTADO
'                xlHoja1.Cells(67, 7) = pnRatioOrigME ANPS COMENTADO
                xlHoja1.Cells(75, 6) = pnRatioOrigMN 'ANPS
                xlHoja1.Cells(75, 7) = pnRatioOrigME 'ANPS
                
                psNivRiesgoMN = rs!NivelRgoAsumMN
                psNivRiesgoME = rs!NivelRgoAsumME
                rs.MoveNext
            Loop
        End If
    Else
        If Not rs.BOF And Not rs.EOF Then
            Do While Not rs.EOF
                If rs!nOrden = "1" Then
                    psNivRiesgoMN = rs!cNivelRgoAsum
                Else
                    psNivRiesgoME = rs!cNivelRgoAsum
                End If
                rs.MoveNext
            Loop
        End If
    End If
'    If psNivRiesgoMN = "B" Then ANPS COMENTADO
'        xlHoja1.Range(xlHoja1.Cells(74, 7), xlHoja1.Cells(74, 7)).Interior.ColorIndex = 43
'    ElseIf psNivRiesgoMN = "M" Then
'        xlHoja1.Range(xlHoja1.Cells(74, 7), xlHoja1.Cells(74, 7)).Interior.ColorIndex = 6
'    ElseIf psNivRiesgoMN = "A" Then
'        xlHoja1.Range(xlHoja1.Cells(74, 7), xlHoja1.Cells(74, 7)).Interior.ColorIndex = 44
'    ElseIf psNivRiesgoMN = "E" Then
'        xlHoja1.Range(xlHoja1.Cells(74, 7), xlHoja1.Cells(74, 7)).Interior.ColorIndex = 3
'    End If
    If psNivRiesgoMN = "B" Then 'ANPS
        xlHoja1.Range(xlHoja1.Cells(82, 7), xlHoja1.Cells(82, 7)).Interior.ColorIndex = 43
    ElseIf psNivRiesgoMN = "M" Then
        xlHoja1.Range(xlHoja1.Cells(82, 7), xlHoja1.Cells(82, 7)).Interior.ColorIndex = 6
    ElseIf psNivRiesgoMN = "A" Then
        xlHoja1.Range(xlHoja1.Cells(82, 7), xlHoja1.Cells(82, 7)).Interior.ColorIndex = 44
    ElseIf psNivRiesgoMN = "E" Then
        xlHoja1.Range(xlHoja1.Cells(82, 7), xlHoja1.Cells(82, 7)).Interior.ColorIndex = 3
    End If
    
'    If psNivRiesgoME = "B" Then ANPS COMENTADO
'        xlHoja1.Range(xlHoja1.Cells(75, 7), xlHoja1.Cells(75, 7)).Interior.ColorIndex = 43
'    ElseIf psNivRiesgoME = "M" Then
'        xlHoja1.Range(xlHoja1.Cells(75, 7), xlHoja1.Cells(75, 7)).Interior.ColorIndex = 6
'    ElseIf psNivRiesgoME = "A" Then
'        xlHoja1.Range(xlHoja1.Cells(75, 7), xlHoja1.Cells(75, 7)).Interior.ColorIndex = 44
'    ElseIf psNivRiesgoME = "E" Then
'        xlHoja1.Range(xlHoja1.Cells(75, 7), xlHoja1.Cells(75, 7)).Interior.ColorIndex = 3
'    End If
    If psNivRiesgoME = "B" Then 'ANPS
        xlHoja1.Range(xlHoja1.Cells(83, 7), xlHoja1.Cells(83, 7)).Interior.ColorIndex = 43
    ElseIf psNivRiesgoME = "M" Then
        xlHoja1.Range(xlHoja1.Cells(83, 7), xlHoja1.Cells(83, 7)).Interior.ColorIndex = 6
    ElseIf psNivRiesgoME = "A" Then
        xlHoja1.Range(xlHoja1.Cells(83, 7), xlHoja1.Cells(83, 7)).Interior.ColorIndex = 44
    ElseIf psNivRiesgoME = "E" Then
        xlHoja1.Range(xlHoja1.Cells(83, 7), xlHoja1.Cells(83, 7)).Interior.ColorIndex = 3
    End If
    'Ratios Considerando el InterCambio de Liquidez, y por ende el Ratio Final
'    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "1", xlHoja1.Cells(66, 3), 1, "4000", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "2", xlHoja1.Cells(66, 4), 1, "4000", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "3", xlHoja1.Cells(66, 6), 1, "4000", "B1") ANPS COMENTADO
'    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "4", xlHoja1.Cells(66, 7), 1, "4000", "B1") ANPS COMENTADO
    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "1", xlHoja1.Cells(74, 3), 1, "4000", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "2", xlHoja1.Cells(74, 4), 1, "4000", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "3", xlHoja1.Cells(74, 6), 1, "4000", "B1") 'ANPS
    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "4", xlHoja1.Cells(74, 7), 1, "4000", "B1") 'ANPS
End Sub 'NAGL ERS006-2019 20190502


Public Sub InsertaLiquidezAlertaTempranaCobertura(ByVal pdFecha As Date, ByVal pnRatioLiquidezMN As Double, ByVal pnRatioLiquidezME As Double)
On Error GoTo InsertaLiquidezAlertaTempranaErr
   Dim oConec As DConecta
   Dim psSql As String
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_ins_ReporteCoberturaLiquidez '" & Format(pdFecha, "YYYY/MM/DD") & "'," & pnRatioLiquidezMN & "," & pnRatioLiquidezME & ""
   oConec.Ejecutar (psSql)
   oConec.CierraConexion
Exit Sub
InsertaLiquidezAlertaTempranaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceDiario Method")
End Sub

Public Function ObtieneCtaSaldoDiarioAnx15B_Det(psCtaContCodMas As String, psCtaContCodMenos As String, pdFecha As Date, pdFechaBalanceDiario, pnTipoCambBalance As Currency, pnTipoCambioMes As Currency, Optional psTipo As String = "") As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
    psSql = "Exec stp_sel_ObtieneCtaSaldoDiarioAnx15A '" & psCtaContCodMas & "','" & psCtaContCodMenos & "','" & Format(pdFecha, "yyyymmdd") & "','" & Format(pdFechaBalanceDiario, "yyyymmdd") & "'," & pnTipoCambBalance & "," & pnTipoCambioMes & ", '" & psTipo & "'"
oConec.AbreConexion
Set ObtieneCtaSaldoDiarioAnx15B_Det = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '****NAGL Según Anexo03 ERS006-2019 20191015
