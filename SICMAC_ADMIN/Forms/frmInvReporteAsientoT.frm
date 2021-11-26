VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInvReporteAsientoT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPORTE DE ASIENTO DE LAS TRANSFERENCIAS"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvReporteAsientoT.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "F. Transferencia"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   59572225
         CurrentDate     =   39896
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   59572225
         CurrentDate     =   39896
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmInvReporteAsientoT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim s As String
    s = ReporteExcelAsientoTransferencia
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
End Sub

Public Function ReporteExcelAsientoTransferencia() As String
    Dim oDInventario As DInvTransferencia
    Set oDInventario = New DInvTransferencia
    
    Dim lsSQL As String
    Dim lrDataRep As New ADODB.Recordset
    Dim lscadimp As String
    Dim lsCadBuffer As String
    Dim lnIndice As Long
    Dim lnLineas As Integer
    Dim lnPage As Integer
    Dim lsOperaciones As String
    Dim P As String
    Dim whereAgencias As String
    Dim nTotalD As Double, nTotalS As Double
    Dim psMensaje As String
    Dim nTotalCas As Integer, i As Integer, j As Integer
 
    Set lrDataRep = oDInventario.CargarReporteAsientoTransferencia
            
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psMensaje = " No Existen Datos para el reporte en la Agencia "
        Exit Function
    Else
    
        Dim ApExcel As Variant
        Set ApExcel = CreateObject("Excel.application")
        
        ApExcel.Workbooks.Add
        
        ApExcel.Cells(2, 3).Formula = "REPORTE DE ASIENTO DE TRANSFERENCIA" '& Format(pdFecCastDe, "dd/MM/YYYY") & " AL " & Format(pdFecCastHasta, "dd/mm/yyyy")
        ApExcel.Range("C2", "H2").MergeCells = True
        ApExcel.Cells(2, 3).Font.Bold = True
        ApExcel.Cells(2, 3).HorizontalAlignment = 3
        
        ApExcel.Cells(8, 3).Formula = "COD. INVENTARIO"
        ApExcel.Cells(8, 4).Formula = "DEBE"
        ApExcel.Cells(8, 5).Formula = "HABER"
        
        ApExcel.Cells(10, 3).Formula = "Valor Activo Fijo"
        
        ApExcel.Cells(14 + lrDataRep.RecordCount, 3).Formula = "Depreciaciacion Activo Fijo"

        ApExcel.Range("C7", "H8").Interior.Color = RGB(10, 190, 160)
        ApExcel.Range("B7", "H8").Font.Bold = True
        ApExcel.Range("B7", "H8").HorizontalAlignment = 3
    
    i = 11
    
        Do While Not lrDataRep.EOF
    
            i = i + 1
                    ApExcel.Cells(i, 3).Formula = lrDataRep!vCodInventario
                    If lrDataRep!vTipo = "O" Then
                        ApExcel.Cells(i, 5).Formula = lrDataRep!nBSValor
                        ApExcel.Range("D" & 8 & ":" & "D" & lrDataRep.RecordCount + 9).NumberFormat = "#,##0.00"
                    Else
                        ApExcel.Cells(i, 4).Formula = lrDataRep!nBSValor
                        ApExcel.Range("E" & 8 & ":" & "E" & lrDataRep.RecordCount + 9).NumberFormat = "#,##0.00"
                    End If
                    
                    ApExcel.Cells(i + lrDataRep.RecordCount + 4, 3).Formula = lrDataRep!vCodInventario
                    If lrDataRep!vTipo = "O" Then
                        ApExcel.Cells(i + lrDataRep.RecordCount + 4, 4).Formula = lrDataRep!nValorDepre
                    Else
                        
                        ApExcel.Cells(i + lrDataRep.RecordCount + 4, 5).Formula = lrDataRep!nValorDepre
                    End If
                    
                    
                    'ApExcel.Cells(i + 1, 4).Formula = lrDataRep!vCodInventario
'                    ApExcel.Cells(i, 5).Formula = lrDataRep!cPersCod
'                    ApExcel.Cells(i, 6).Formula = lrDataRep!Cliente
'                    ApExcel.Cells(i, 7).Formula = lrDataRep!direccion
'                    ApExcel.Cells(i, 8).Formula = "'" & lrDataRep!DocID
'
'                    ApExcel.Cells(i, 9).Formula = "'" & Trim(lrDataRep!PersRel)
'                    ApExcel.Cells(i, 10).Formula = "'" & Trim(lrDataRep!PersRelNombre)
'                    ApExcel.Cells(i, 11).Formula = "'" & Trim(lrDataRep!PersRelDire)
'
'                    ApExcel.Cells(i, 12).Formula = Format(lrDataRep!Fecha, "mm/dd/yyyy")
'                    ApExcel.Cells(i, 13).Formula = lrDataRep!nMontoCol
'                    ApExcel.Cells(i, 14).Formula = lrDataRep!nMontoCol - lrDataRep!nSaldo
'                    ApExcel.Cells(i, 15).Formula = lrDataRep!nSaldo
'                    ApExcel.Cells(i, 16).Formula = lrDataRep!Interes
'                    ApExcel.Cells(i, 17).Formula = lrDataRep!Mora
'                    ApExcel.Cells(i, 18).Formula = lrDataRep!Gastos
'                    ApExcel.Cells(i, 19).Formula = lrDataRep!TotalDCast
'                    'ApExcel.Cells(i, 17).Formula = lrDataRep!TotalACastMN
'                    ApExcel.Cells(i, 20).Formula = lrDataRep!TotalDCastMN
'                    ApExcel.Cells(i, 21).Formula = lrDataRep!CapiCas
'                    ApExcel.Cells(i, 22).Formula = lrDataRep!InteCas
'                    ApExcel.Cells(i, 23).Formula = lrDataRep!MoraCas
'                    ApExcel.Cells(i, 24).Formula = lrDataRep!GastCas
'                    ApExcel.Cells(i, 25).Formula = lrDataRep!TotalACast
'                    ApExcel.Cells(i, 26).Formula = lrDataRep!TotalACastMN
'                    ApExcel.Cells(i, 27).Formula = UCase(lrDataRep!Analista)
'                    ApExcel.Cells(i, 28).Formula = lrDataRep!nDiasAtraso
'                    ApExcel.Cells(i, 29).Formula = lrDataRep!NombAgencia
'                    ApExcel.Cells(i, 30).Formula = IIf(IsNull(lrDataRep!dPersFallec), "", IIf(lrDataRep!dPersFallec = "", "", Format(lrDataRep!dPersFallec, "mm/dd/yyyy")))
                                    
                                    
                    'ApExcel.Range("M" & Trim(Str(i)) & ":" & "Z" & Trim(Str(i))).NumberFormat = "#,##0.00"
                    
                    lrDataRep.MoveNext
'                    If lrDataRep.EOF Then
'                        Exit Do
'                    End If
                    
            'Loop
        Loop
        
        lrDataRep.Close
        Set lrDataRep = Nothing
       
        ApExcel.Cells.Select
        ApExcel.Cells.EntireColumn.AutoFit
        ApExcel.Columns("B:B").ColumnWidth = 6#
        ApExcel.Range("B2").Select
    
        ApExcel.Visible = True
        Set ApExcel = Nothing
           
        End If
End Function

