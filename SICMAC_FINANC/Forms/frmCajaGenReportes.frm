VERSION 5.00
Begin VB.Form frmCajaGenReportes 
   Caption         =   "Reportes Caja General"
   ClientHeight    =   795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmCajaGenReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oCaja As nCajaGenImprimir
Attribute oCaja.VB_VarHelpID = -1
Dim oBarra As clsProgressBar

Dim lsArchivo As String
Dim lbExcel As Boolean
Dim xlAplicacion As excel.Application
Dim xlLibro As excel.Workbook
Dim xlHoja1 As excel.Worksheet

Public Sub ConsolidaSdoEnc(psOpeCod As String, psFecIni As String, psFecFin As String, Optional pnTipo As Integer = 1)
Dim lsImpre As String
On Error GoTo GeneraEstadError
   lsArchivo = App.path & "\SPOOLER\" & "Pla_Enc_" & Format(psFecFin, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & "_" & IIf(Mid(psOpeCod, 3, 1) = gMonedaNacional, "MN", "ME") & gsCodUser & ".XLS"
   lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
   If Not lbExcel Then
      Exit Sub
   End If

    Set oCaja = New nCajaGenImprimir
    
    Set xlAplicacion = New excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    Set xlHoja1 = xlLibro.Worksheets.Add
    
    
    oCaja.Inicio gsInstCmac, gdFecSis
       
    lsImpre = oCaja.ImprimeConsolidaSdoEnc(gbBitCentral, psOpeCod, psFecIni, psFecFin, pnTipo, xlAplicacion, xlLibro, xlHoja1)
    Set oCaja = Nothing
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    If lsArchivo <> "" Then
       CargaArchivo lsArchivo, App.path & "\SPOOLER\"
    End If
   
Exit Sub
GeneraEstadError:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    If lbExcel = True Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub

Private Sub Form_Load()
CentraForm Me
End Sub

Private Sub oCaja_BarraClose()
oBarra.CloseForm Me
Set oBarra = Nothing
End Sub

Private Sub oCaja_BarraProgress(value As Variant, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
oBarra.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Sub

Private Sub oCaja_BarraShow(pnMax As Variant)
Set oBarra = New clsProgressBar
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.ShowForm Me
oBarra.Max = pnMax
End Sub

