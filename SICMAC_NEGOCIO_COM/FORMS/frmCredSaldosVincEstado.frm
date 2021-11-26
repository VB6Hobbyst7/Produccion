VERSION 5.00
Begin VB.Form frmCredSaldosVincEstado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado Actual de Saldo para Asignación de Crédito Trabajadores"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "frmCredSaldosVincEstado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkUltFinMes 
      Caption         =   "Ultimo fin de mes"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdDetalleSaldoActual 
      Caption         =   "Detalle Saldo Actual Prestamo"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información Patrimonial"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin VB.Label lblMMGE 
         AutoSize        =   -1  'True
         Caption         =   "Monto Máx. x Grupo Econ.(5%):"
         Height          =   195
         Left            =   5040
         TabIndex        =   13
         Top             =   1560
         Width           =   2250
      End
      Begin VB.Label lblMontoMaxGE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7320
         TabIndex        =   12
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Reserva Créditos Vinculados:"
         Height          =   195
         Left            =   5040
         TabIndex        =   11
         Top             =   480
         Width           =   2085
      End
      Begin VB.Label lblReservaCred 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7320
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Disponible a Asignación:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   2190
      End
      Begin VB.Label lblSaldoDisponible 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Actual de Préstamos:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1950
      End
      Begin VB.Label lblSaldoActual 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblMMPE 
         AutoSize        =   -1  'True
         Caption         =   "Monto Máximo a asignar(7%):"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   2070
      End
      Begin VB.Label lblMontoMax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblPatrimonioActualDesc 
         AutoSize        =   -1  'True
         Caption         =   "Patrimonio Actual"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label lblPatrimonioActual 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmCredSaldosVincEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ORCR20140314**************************
'Private fgFecActual As Date
Private ffFecActual As Date
Private fnPatrimonioEfec As Double
'END ORCR20140314**************************

Private Sub cmdCerrar_Click()
    Unload Me
End Sub
'ORCR20140314**************************
Private Sub cmdDetalleSaldoActual_Click()
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    
    Dim fs  As New Scripting.FileSystemObject
    
    Dim lsArchivo As String
    Dim lsFile As String
    
    Dim lsNomHoja As String
    Dim lbExisteHoja As Boolean
    
    Dim fila_ini As Integer
    Dim colu_ini As Integer
    
    Dim i As Integer
    'ORCR20140617**************************
    Dim oPatrimonioEfectivo As New COMNCredito.NCOMPatrimonioEfectivo
    
    Dim ffFecReporte As Date
    
    If chkUltFinMes.value = 1 Then
        ffFecReporte = DateAdd("m", -1, ffFecActual)
        fnPatrimonioEfec = oPatrimonioEfectivo.ObtenerPatrimonioEfectivo(Year(ffFecReporte), Format(Month(ffFecReporte), "00"))
    Else
        ffFecReporte = ffFecActual
    End If
    
    If fnPatrimonioEfec = 0 Then
            MsgBox "favor de definir el Patrimonio Efectivo para continuar", vbInformation, "Aviso"
            Exit Sub
    End If
    '*******************************************************
    lsNomHoja = "Hoja1"
    lsFile = "Rpt_SAPT.xlsx"
    
    lsArchivo = "\spooler\" & "Rpt_SAPT" & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xlsx"
    
    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile) Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile)
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & "), Consulte con el Area de TI", vbInformation, "Advertencia"
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
    '*******************************************************
    fila_ini = 8
    colu_ini = 2
    i = 0
    '*******************************************************
    xlHoja1.Cells(4, 12) = "Fecha : " & gdFecSis
    xlHoja1.Cells(5, 12) = "Usuario : " & gsCodUser
    '*******************************************************
    '*******************************************************
    Dim n7Porc As Double
    Dim n5Porc As Double
    
    Dim oPar As New COMDCredito.DCOMParametro
    
    n7Porc = oPar.RecuperaValorParametro(102752)
    n5Porc = oPar.RecuperaValorParametro(102753)
    '*******************************************************
    '*******************************************************
    xlHoja1.Cells(3, 13) = "PE " & MesAnio(ffFecReporte) & " :"
    xlHoja1.Cells(3, 14) = fnPatrimonioEfec
    
    xlHoja1.Cells(4, 13) = "L (" & n7Porc & "%):"
    xlHoja1.Cells(4, 14) = n7Porc / 100
    xlHoja1.Range(xlHoja1.Cells(4, 15).Address).Formula = "=" & xlHoja1.Cells(4, 14).Address & "*" & xlHoja1.Cells(3, 14).Address
    
    xlHoja1.Cells(5, 13) = "LxGE (" & n5Porc & "%):"
    xlHoja1.Cells(5, 14) = n5Porc / 100
    xlHoja1.Range(xlHoja1.Cells(5, 15).Address).Formula = "=" & xlHoja1.Cells(5, 14).Address & "*" & xlHoja1.Cells(4, 14).Address & "*" & xlHoja1.Cells(3, 14).Address
    '*******************************************************
    '*******************************************************
    Dim oNCOMCredito As New COMNCredito.NCOMCredito
    Dim rs As ADODB.Recordset
    
    Set rs = oNCOMCredito.ReporteSaldosTrabDirVinc(chkUltFinMes.value)
    Dim Ancla As Integer
    Ancla = fila_ini
    Do While Not rs.EOF
        xlHoja1.Cells(fila_ini + i, colu_ini + 0) = i + 1
        xlHoja1.Cells(fila_ini + i, colu_ini + 1) = rs!cCtaCod
        xlHoja1.Cells(fila_ini + i, colu_ini + 2) = rs!dVigencia 'Format(rs!dVigencia, "dd/mm/yyyy") 'Format(rs!dVigencia, "dd/mm/yyyy")
        xlHoja1.Cells(fila_ini + i, colu_ini + 3) = rs!cPersNombre
        
        xlHoja1.Cells(fila_ini + i, colu_ini + 4) = rs!ccalgen '-------------
        
        xlHoja1.Cells(fila_ini + i, colu_ini + 5) = rs!Moneda
        
        xlHoja1.Cells(fila_ini + i, colu_ini + 6) = rs!nMontoCol  '-------------
        
        xlHoja1.Cells(fila_ini + i, colu_ini + 7) = rs!nSaldo
        xlHoja1.Cells(fila_ini + i, colu_ini + 8) = rs!nSaldoMN
        xlHoja1.Cells(fila_ini + i, colu_ini + 9) = rs!Relac
        xlHoja1.Cells(fila_ini + i, colu_ini + 10) = rs!Vinculado
        '*******************************************************
        xlHoja1.Cells(fila_ini + i, colu_ini + 11) = rs!nSaldoMN
        xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i, colu_ini + 11)).Formula = "=SUM(" & xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 8), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 8)).Address & ")"
        
        xlHoja1.Range(xlHoja1.Cells(fila_ini + i, colu_ini + 11).Address).Formula = "=SUM(" & xlHoja1.Cells(fila_ini + i, colu_ini + 8).Address & ")" 'xlHoja1.Cells(Ancla, colu_ini + 6), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 6)).Address & ")"
        xlHoja1.Range(xlHoja1.Cells(fila_ini + i, colu_ini + 12).Address).Formula = "=" & xlHoja1.Cells(fila_ini + i, colu_ini + 11).Address & "/" & xlHoja1.Cells(4, 15).Address
        '*******************************************************
        If Not xlHoja1.Cells(fila_ini + i, colu_ini + 10) = xlHoja1.Cells(fila_ini + i - 1, colu_ini + 10) Then
            If Ancla <> (fila_ini + i) Then
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 11)).value = ""
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 11)).Merge
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 11)).Formula = "=SUM(" & xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 8), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 8)).Address & ")"
                
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 12), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 12)).value = ""
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 12), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 12)).Merge
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 12), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 12)).Formula = "=" & xlHoja1.Cells(Ancla, colu_ini + 11).Address & "/" & xlHoja1.Cells(4, 15).Address
            Else
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i, colu_ini + 11)).value = ""
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i, colu_ini + 11)).Merge
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i, colu_ini + 11)).Formula = "=SUM(" & xlHoja1.Range(xlHoja1.Cells(Ancla - 0, colu_ini + 8), xlHoja1.Cells(fila_ini + i - 0, colu_ini + 8)).Address & ")"
                
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 12), xlHoja1.Cells(fila_ini + i, colu_ini + 12)).value = ""
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 12), xlHoja1.Cells(fila_ini + i, colu_ini + 12)).Merge
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 12), xlHoja1.Cells(fila_ini + i, colu_ini + 12)).Formula = "=" & xlHoja1.Cells(Ancla, colu_ini + 11).Address & "/" & xlHoja1.Cells(4, 15).Address
            End If
            Ancla = fila_ini + i
        End If
        '*******************************************************
        rs.MoveNext
        i = i + 1
    Loop
    rs.Close
    '*******************************************************
    xlHoja1.Range(xlHoja1.Cells(fila_ini, colu_ini), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 13 - 1)).Borders.LineStyle = 1
    '*******************************************************
    i = 0
    lsNomHoja = "Hoja2"
    '*******************************************************
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
    
    xlHoja1.Cells(4, 12) = "Fecha : " & gdFecSis
    xlHoja1.Cells(5, 12) = "Usuario : " & gsCodUser
    '*******************************************************
    '*******************************************************
    xlHoja1.Cells(3, 13) = "PE " & MesAnio(ffFecReporte) & " :"
    xlHoja1.Cells(3, 14) = fnPatrimonioEfec
    
    xlHoja1.Cells(4, 13) = "L (" & n7Porc & "%):"
    xlHoja1.Cells(4, 14) = n7Porc / 100
    xlHoja1.Range(xlHoja1.Cells(4, 15).Address).Formula = "=" & xlHoja1.Cells(4, 14).Address & "*" & xlHoja1.Cells(3, 14).Address
    
    xlHoja1.Cells(5, 13) = "LxGE (" & n5Porc & "%):"
    xlHoja1.Cells(5, 14) = n5Porc / 100
    xlHoja1.Range(xlHoja1.Cells(5, 15).Address).Formula = "=" & xlHoja1.Cells(5, 14).Address & "*" & xlHoja1.Cells(4, 14).Address & "*" & xlHoja1.Cells(3, 14).Address
    '*******************************************************
    '*******************************************************
    
    Set rs = oNCOMCredito.ReporteSaldoAsignaEstado(2, chkUltFinMes.value)
     
    Ancla = fila_ini
    Do While Not rs.EOF
        xlHoja1.Cells(fila_ini + i, colu_ini + 0) = i + 1
        xlHoja1.Cells(fila_ini + i, colu_ini + 1) = rs!cCtaCod
        xlHoja1.Cells(fila_ini + i, colu_ini + 2) = Format(rs!dVigencia, "dd/mm/yyyy")
        xlHoja1.Cells(fila_ini + i, colu_ini + 3) = rs!cPersNombre
        
        xlHoja1.Cells(fila_ini + i, colu_ini + 4) = rs!ccalgen '-------------
        
        xlHoja1.Cells(fila_ini + i, colu_ini + 5) = rs!Moneda
        
        xlHoja1.Cells(fila_ini + i, colu_ini + 6) = rs!nMontoCol  '-------------
        
        xlHoja1.Cells(fila_ini + i, colu_ini + 7) = rs!nSaldo
        xlHoja1.Cells(fila_ini + i, colu_ini + 8) = rs!nSaldoMN
        xlHoja1.Cells(fila_ini + i, colu_ini + 9) = rs!Relac
        xlHoja1.Cells(fila_ini + i, colu_ini + 10) = rs!Vinculado
        '*******************************************************
        xlHoja1.Cells(fila_ini + i, colu_ini + 11) = rs!nSaldoMN
        xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i, colu_ini + 11)).Formula = "=SUM(" & xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 8), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 8)).Address & ")"
        
        xlHoja1.Range(xlHoja1.Cells(fila_ini + i, colu_ini + 11).Address).Formula = "=SUM(" & xlHoja1.Cells(fila_ini + i, colu_ini + 8).Address & ")" 'xlHoja1.Cells(Ancla, colu_ini + 6), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 6)).Address & ")"
        xlHoja1.Range(xlHoja1.Cells(fila_ini + i, colu_ini + 12).Address).Formula = "=" & xlHoja1.Cells(fila_ini + i, colu_ini + 11).Address & "/" & xlHoja1.Cells(4, 15).Address
        '*******************************************************
        If Not xlHoja1.Cells(fila_ini + i, colu_ini + 10) = xlHoja1.Cells(fila_ini + i - 1, colu_ini + 10) Then
            If Ancla <> (fila_ini + i) Then
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 11)).value = ""
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 11)).Merge
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 11)).Formula = "=SUM(" & xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 8), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 8)).Address & ")"
                
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 12), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 12)).value = ""
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 12), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 12)).Merge
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 12), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 12)).Formula = "=" & xlHoja1.Cells(Ancla, colu_ini + 11).Address & "/" & xlHoja1.Cells(4, 15).Address
            Else
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i, colu_ini + 11)).value = ""
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i, colu_ini + 11)).Merge
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 11), xlHoja1.Cells(fila_ini + i, colu_ini + 11)).Formula = "=SUM(" & xlHoja1.Range(xlHoja1.Cells(Ancla - 0, colu_ini + 8), xlHoja1.Cells(fila_ini + i - 0, colu_ini + 8)).Address & ")"
                
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 12), xlHoja1.Cells(fila_ini + i, colu_ini + 12)).value = ""
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 12), xlHoja1.Cells(fila_ini + i, colu_ini + 12)).Merge
                xlHoja1.Range(xlHoja1.Cells(Ancla, colu_ini + 12), xlHoja1.Cells(fila_ini + i, colu_ini + 12)).Formula = "=" & xlHoja1.Cells(Ancla, colu_ini + 11).Address & "/" & xlHoja1.Cells(4, 15).Address
            End If
            Ancla = fila_ini + i
        End If
        '*******************************************************
        rs.MoveNext
        i = i + 1
    Loop
    rs.Close
    '*******************************************************
    xlHoja1.Range(xlHoja1.Cells(fila_ini, colu_ini), xlHoja1.Cells(fila_ini + i - 1, colu_ini + 13 - 1)).Borders.LineStyle = 1
    '*******************************************************
    'FIN ORCR20140617**************************
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.path & lsArchivo
    psArchivoAGrabarC = App.path & lsArchivo
    
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
End Sub
Public Sub Inicio()
    Dim oConsSist As New COMDConstSistema.NCOMConstSistema
    Dim oPatrimonioEfectivo As New COMNCredito.NCOMPatrimonioEfectivo
    
    ffFecActual = oConsSist.LeeConstSistema(gConstSistFechaInicioDia)
    
    fnPatrimonioEfec = oPatrimonioEfectivo.ObtenerPatrimonioEfectivo(Year(ffFecActual), Format(Month(ffFecActual), "00"))
    
    If fnPatrimonioEfec = 0 Then
        ffFecActual = DateAdd("m", -1, ffFecActual)
        fnPatrimonioEfec = oPatrimonioEfectivo.ObtenerPatrimonioEfectivo(Year(ffFecActual), Format(Month(ffFecActual), "00"))
        
        If fnPatrimonioEfec = 0 Then
            MsgBox "favor de definir el Patrimonio Efectivo para continuar", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    Me.Show 1
End Sub
'END ORCR20140314**************************

Private Sub Form_Load()
    CargaValoresParametros
End Sub

Private Sub CargaValoresParametros()
    'ORCR20140314**************************
'    Dim sAnio As String
'    Dim sMes As String
'    Dim nPatrimonioEfec As Double
    Dim nReservaCred As Double
    Dim n7PorcPatriEfec As Double
    Dim n5PorcDel7PorcPatriEfec As Double
    Dim nSaldosVinculados As Double
    Dim nSaldosAceptados As Double
'    Dim nSaldoActual As Double
    Dim nMontoMaxGE As Double
      
    Dim n7Porc As Double
    Dim n5Porc As Double
    
    Dim oPar As New COMDCredito.DCOMParametro
    Dim oCredito As New COMNCredito.NCOMCredito
'    Dim oConsSist As COMDConstSistema.NCOMConstSistema
    
    'Set oConsSist = New COMDConstSistema.NCOMConstSistema
    
    'fgFecActual = oConsSist.LeeConstSistema(gConstSistCierreMesNegocio)
    'sAnio = Year(fgFecActual)
    'sMes = Format(Month(fgFecActual), "00")
    
'    sAnio = Year(ffFecActual)
'    sMes = Format(Month(ffFecActual), "00")
    
    'nPatrimonioEfec = ObtenerSaldos(sAnio, sMes)
    
    'Set oPar = New COMDCredito.DCOMParametro
    nReservaCred = oPar.RecuperaValorParametro(102751)
    
    'ORCR20140617
    n7Porc = oPar.RecuperaValorParametro(102752)
    n5Porc = oPar.RecuperaValorParametro(102753)
      
    'n7PorcPatriEfec = oPar.RecuperaValorParametro(102752) / 100
    'n5PorcDel7PorcPatriEfec = oPar.RecuperaValorParametro(102753) / 100
    
    n7PorcPatriEfec = n7Porc / 100
    n5PorcDel7PorcPatriEfec = n5Porc / 100
        
    lblMMPE.Caption = "Monto Máximo a asignar(" & n7Porc & "%):"
    lblMMGE.Caption = "Monto Máx. x Grupo Econ.(" & n5Porc & "%):"
    'FIN ORCR20140617
    
    'Set oCredito = New COMNCredito.NCOMCredito
    nSaldosVinculados = oCredito.ObtenerSaldoTrabDirVinc
    nSaldosAceptados = oCredito.ObtenerSaldoAsignaEstado(2)
    
'    nSaldoActual = nSaldosVinculados + nSaldosAceptados
    nMontoMaxGE = n5PorcDel7PorcPatriEfec * n7PorcPatriEfec * fnPatrimonioEfec
    
    'lblPatrimonioActualDesc.Caption = lblPatrimonioActualDesc.Caption & "(" & MesAnio(fgFecActual) & "):"
    'Me.lblPatrimonioActual.Caption = Format(nPatrimonioEfec, "###," & String(15, "#") & "#0.00") & " "
    'Me.lblMontoMax.Caption = Format(n7PorcPatriEfec * nPatrimonioEfec, "###," & String(15, "#") & "#0.00") & " "
    lblPatrimonioActualDesc.Caption = lblPatrimonioActualDesc.Caption & "(" & MesAnio(ffFecActual) & "):"
    Me.lblPatrimonioActual.Caption = Format(fnPatrimonioEfec, "###," & String(15, "#") & "#0.00") & " "
    Me.lblMontoMax.Caption = Format(n7PorcPatriEfec * fnPatrimonioEfec, "###," & String(15, "#") & "#0.00") & " "
    Me.lblReservaCred.Caption = Format(nReservaCred, "###," & String(15, "#") & "#0.00") & " "
    Me.lblSaldoActual.Caption = Format(nSaldosVinculados + nSaldosAceptados, "###," & String(15, "#") & "#0.00") & " "
    'Me.lblSaldoDisponible.Caption = Format((n7PorcPatriEfec * nPatrimonioEfec) - nSaldoActual, "###," & String(15, "#") & "#0.00") & " "
    Me.lblSaldoDisponible.Caption = Format((n7PorcPatriEfec * fnPatrimonioEfec) - nSaldosVinculados - nSaldosAceptados - nReservaCred, "###," & String(15, "#") & "#0.00") & " "
    Me.lblMontoMaxGE.Caption = Format(nMontoMaxGE, "###," & String(15, "#") & "#0.00") & " "
    
    'Set oConsSist = Nothing
    Set oPar = Nothing
    'END ORCR20140314**************************
End Sub


'ORCR20140314**************************
'Private Function ObtenerSaldos(ByVal psAnio As String, ByVal psMes As String) As Double
'Dim oNContabilidad As COMNContabilidad.NCOMContFunciones
'Dim nSaldo As Double
'Set oNContabilidad = New COMNContabilidad.NCOMContFunciones
'
'nSaldo = oNContabilidad.PatrimonioEfecAjustInfl(psAnio, psMes)
'ObtenerSaldos = nSaldo
'Set oNContabilidad = Nothing
'End Function
'END ORCR20140314**************************

Private Function MesAnio(ByVal dFecha As Date) As String
    Dim sFechaDesc As String
    sFechaDesc = ""
    
    Select Case Month(dFecha)
        Case 1: sFechaDesc = "Enero"
        Case 2: sFechaDesc = "Febrero"
        Case 3: sFechaDesc = "Marzo"
        Case 4: sFechaDesc = "Abril"
        Case 5: sFechaDesc = "Mayo"
        Case 6: sFechaDesc = "Junio"
        Case 7: sFechaDesc = "Julio"
        Case 8: sFechaDesc = "Agosto"
        Case 9: sFechaDesc = "Septiembre"
        Case 10: sFechaDesc = "Octubre"
        Case 11: sFechaDesc = "Noviembre"
        Case 12: sFechaDesc = "Diciembre"
    End Select
    
    sFechaDesc = sFechaDesc & " " & CStr(Year(dFecha))
    MesAnio = sFechaDesc
End Function
