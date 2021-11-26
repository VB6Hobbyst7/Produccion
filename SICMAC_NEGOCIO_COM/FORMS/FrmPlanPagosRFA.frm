VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmPlanPagosRFA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de Pagos RFA"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   435
      Left            =   9930
      TabIndex        =   1
      Top             =   5970
      Width           =   1245
   End
   Begin TabDlg.SSTab Stab 
      Height          =   6555
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   11562
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Resumen del Plan de Pagos"
      TabPicture(0)   =   "FrmPlanPagosRFA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Flex"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle del Plan de Pagos"
      TabPicture(1)   =   "FrmPlanPagosRFA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FlexEdit1"
      Tab(1).ControlCount=   1
      Begin SICMACT.FlexEdit FlexEdit1 
         Height          =   5205
         Left            =   -74760
         TabIndex        =   3
         Top             =   600
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   9181
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
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
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   -1
         RowHeight0      =   240
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit Flex 
         Height          =   4875
         Left            =   150
         TabIndex        =   2
         Top             =   750
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   8599
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Cuota-Estado-Fec.Venc-Fec.Pago-Mont.Pend.Dif-Mont.Pag.Dif-Mont.Pend.Rfc-Mont.Pend.Rfc-Mont.Pen.Rfa-Mont.Pag.Rfa-Saldo Total"
         EncabezadosAnchos=   "1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "Cuota"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   1200
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "FrmPlanPagosRFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nCodCli As String

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Flex.Clear
    Flex.Rows = 2
    Flex.FormaCabecera
'    ConfigurarDGResumen
    Stab.Tab = 0
    Stab.TabVisible(1) = False
    Cargar_CuotasResumen
End Sub

Private Sub Cargar_CuotasResumen()
    Dim objPlanRFA As COMDCredito.DCOMRFA
    Dim rs As ADODB.Recordset
    Dim nCuota As Integer
    Dim nMontoPos As Double
    Dim nMontoNeg As Double
    
    On Error GoTo ErrHandler
        
        nMontoPos = 0
        nMontoNeg = 0
        
        Set objPlanRFA = New COMDCredito.DCOMRFA
        Set rs = objPlanRFA.ListaCuotas(nCodCli)
        Set objPlanRFA = Nothing
               Do Until rs.EOF
               Flex.AdicionaFila
               Flex.TextMatrix(Flex.Rows - 1, 0) = rs!nCuotaDIF
               Flex.TextMatrix(Flex.Rows - 1, 1) = DeterminarEstadoPlanResumen(rs!EstadoDIF, rs!EstadoRFC, rs!EstadoRFA)
               Flex.TextMatrix(Flex.Rows - 1, 2) = Format(rs!dVencDIF, "dd/MM/yyyy")
               Flex.TextMatrix(Flex.Rows - 1, 3) = ObtenerFechaPago(rs!dPagoDIF, rs!dPagoRFC, rs!dPagoRFA, rs!EstadoDIF, rs!EstadoRFC, rs!EstadoRFA)
               Flex.TextMatrix(Flex.Rows - 1, 4) = Format(rs!nMontoPagoDIF, "#0.00")
               Flex.TextMatrix(Flex.Rows - 1, 5) = Format(rs!nMontoPagadoDif, "#0.00")
               Flex.TextMatrix(Flex.Rows - 1, 6) = Format(rs!nMontoPagoRFC, "#0.00")
               Flex.TextMatrix(Flex.Rows - 1, 7) = Format(rs!nMontoPagadoRFC, "#0.00")
               Flex.TextMatrix(Flex.Rows - 1, 8) = Format(rs!nMontoPagoRFA, "#0.00")
               Flex.TextMatrix(Flex.Rows - 1, 9) = Format(rs!nMontoPagadoRFA, "#0.00")
           
                If rs!EstadoDIF = "CANC." Then
                    Flex.Row = Flex.Rows - 1
                    
                    Flex.Col = 4
                    Flex.CellForeColor = vbBlue
                    Flex.TextMatrix(Flex.Rows - 1, 4) = Format(rs!nMontoPagoDIF, "#0.00")
                    Flex.Col = 5
                    Flex.CellForeColor = vbBlue
                    Flex.TextMatrix(Flex.Rows - 1, 5) = Format(rs!nMontoPagadoDif, "#0.00")
                    
                Else
                    Flex.Row = Flex.Rows - 1
                    
                    Flex.Col = 4
                    Flex.CellForeColor = vbRed
                    Flex.TextMatrix(Flex.Rows - 1, 4) = Format(rs!nMontoPagoDIF, "#0.00")
                    Flex.Col = 5
                    Flex.CellForeColor = vbRed
                    Flex.TextMatrix(Flex.Rows - 1, 5) = Format(rs!nMontoPagadoDif, "#0.00")
                    
                End If
                
                If rs!EstadoRFC = "CANC." Then
                    Flex.Row = Flex.Rows - 1
                    
                    Flex.Col = 6
                    Flex.CellForeColor = vbBlue
                    Flex.TextMatrix(Flex.Rows - 1, 6) = Format(rs!nMontoPagoRFC, "#0.00")
                    Flex.Col = 7
                    Flex.CellForeColor = vbBlue
                    Flex.TextMatrix(Flex.Rows - 1, 7) = Format(rs!nMontoPagadoRFC, "#0.00")
                Else
                    Flex.Row = Flex.Rows - 1
                    
                    Flex.Col = 6
                    Flex.CellForeColor = vbRed
                    Flex.TextMatrix(Flex.Rows - 1, 6) = Format(rs!nMontoPagoRFC, "#0.00")
                    Flex.Col = 7
                    Flex.CellForeColor = vbRed
                    Flex.TextMatrix(Flex.Rows - 1, 7) = Format(rs!nMontoPagadoRFC, "#0.00")
                End If
                If rs!EstadoRFA = "CANC." Then
                    Flex.Row = Flex.Rows - 1
                    
                    Flex.Col = 8
                    Flex.CellForeColor = vbBlue
                    Flex.TextMatrix(Flex.Rows - 1, 8) = Format(rs!nMontoPagoRFA, "#0.00")
                    Flex.Col = 9
                    Flex.CellForeColor = vbBlue
                    Flex.TextMatrix(Flex.Rows - 1, 9) = Format(rs!nMontoPagadoRFA, "#0.00")
                Else
                    Flex.Row = Flex.Rows - 1
                    
                    Flex.Col = 8
                    Flex.CellForeColor = vbRed
                    Flex.TextMatrix(Flex.Rows - 1, 8) = Format(rs!nMontoPagoRFA, "#0.00")
                    Flex.Col = 9
                    Flex.CellForeColor = vbRed
                    Flex.TextMatrix(Flex.Rows - 1, 9) = Format(rs!nMontoPagadoRFA, "#0.00")
                End If
                
                nMontoPos = CDbl(rs!nMontoPagadoDif) + CDbl(rs!nMontoPagadoRFC) + CDbl(rs!nMontoPagadoRFA)
                nMontoNeg = CDbl(rs!nMontoPagoDIF) + CDbl(rs!nMontoPagoRFC) + CDbl(rs!nMontoPagoRFA)
                Flex.TextMatrix(Flex.Rows - 1, 10) = Format(nMontoNeg - nMontoPos, "#0.00")
                
                If (nMontoNeg - nMontoPos) > 0 Then
                    Flex.Col = 10
                    Flex.Row = Flex.Rows - 1
                    Flex.CellForeColor = vbRed
                    Flex.TextMatrix(Flex.Rows - 1, 10) = Format((nMontoNeg - nMontoPos), "#0.00")
                Else
                    Flex.Col = 10
                    Flex.Row = Flex.Rows - 1
                    Flex.CellForeColor = vbBlue
                    Flex.TextMatrix(Flex.Rows - 1, 10) = Format(Abs((nMontoNeg - nMontoPos)), "#0.00")
                End If
                rs.MoveNext
            Loop
        
    Exit Sub
ErrHandler:
    If Not objPlanRFA Is Nothing Then Set objPlanRFA = Nothing
    MsgBox "Se ha producido un error al cargar cuotas", vbInformation, "AVISO"
End Sub

Function DeterminarEstadoPlanResumen(ByVal cEstadoDif As String, ByVal cEstadoRFC As String, _
                                     ByVal cEstadoRFA As String) As String

    Dim nEstadoDIF As Integer
    Dim nEstadoRFC As Integer
    Dim nEstadoRFA As Integer
            
            If cEstadoDif = "CANC." Then
                nEstadoDIF = 0
            Else
                nEstadoDIF = 1
            End If
            
            If cEstadoRFC = "CANC." Then
                nEstadoRFC = 0
            Else
                nEstadoRFC = 1
            End If
            
            If cEstadoRFA = "CANC." Then
                nEstadoRFA = 0
            Else
                nEstadoRFA = 1
            End If
            
            If nEstadoDIF = 1 Or nEstadoRFC = 1 Or nEstadoRFA = 1 Then
                DeterminarEstadoPlanResumen = "PEND."
            Else
                DeterminarEstadoPlanResumen = "CANC."
            End If
End Function


Function ObtenerFechaPago(ByVal pdFecDif As Date, ByVal pdFecRFC As Date, ByVal pdFecRFA As Date, _
                          ByVal cEstadoDif As String, ByVal cEstadoRFC As String, ByVal cEstadoRFA As String) As String
    Dim nDIf As Integer
    Dim nRFC As Integer
    Dim nRFA As Integer
    Dim dTemp As Date
            
                nDIf = DetEstado(cEstadoDif)
                nRFC = DetEstado(cEstadoRFC)
                nRFA = DetEstado(cEstadoRFA)
                
                If nDIf = 1 And nRFC = 0 And nRFA = 0 Then
                    ObtenerFechaPago = IIf(pdFecDif = "01/01/1900", "", pdFecDif)
                ElseIf nDIf = 0 And nRFC = 1 And nRFA = 0 Then
                    ObtenerFechaPago = IIf(pdFecRFC = "01/01/1900", "", pdFecDif)
                ElseIf nDIf = 0 And nRFC = 0 And nRFA = 1 Then
                    ObtenerFechaPago = IIf(pdFecRFA = "01/01/1900", "", pdFecDif)
                ElseIf nDIf = 1 And nRFC = 1 And nRFA = 0 Then
                    If DateDiff("d", pdFecDif, pdFecRFC) > 0 Then
                        ObtenerFechaPago = IIf(pdFecDif = "01/01/1900", "", pdFecDif)
                    Else
                        ObtenerFechaPago = IIf(pdFecRFC = "01/01/1900", "", pdFecDif)
                    End If
                    
                ElseIf nDIf = 0 And nRFC = 1 And nRFA = 1 Then
                    If DateDiff("d", pdFecRFC, pdFecRFA) > 0 Then
                        ObtenerFechaPago = IIf(pdFecRFC = "01/01/1900", "", pdFecRFC)
                    Else
                        ObtenerFechaPago = IIf(pdFecRFA = "01/01/1900", "", pdFecRFA)
                    End If
                    
                ElseIf nDIf = 1 And nRFC = 0 And nRFA = 1 Then
                    If DateDiff("d", pdFecDif, pdFecRFA) > 0 Then
                        ObtenerFechaPago = IIf(pdFecDif = "01/01/1900", "", pdFecDif)
                    Else
                        ObtenerFechaPago = IIf(pdFecRFA = "01/01/1900", "", pdFecRFA)
                    End If
                ElseIf nDIf = 1 And nRFC = 1 And nRFA = 1 Then
                        dTemp = pdFecDif
                    If DateDiff("d", pdFecDif, pdFecRFA) < 0 Then
                        dTemp = pdFecRFA
                    End If
                    If DateDiff("d", dTemp, pdFecRFC) < 0 Then
                        dTemp = pdFecRFC
                    End If
                      ObtenerFechaPago = IIf(pdFecRFC = "01/01/1900", "", pdFecRFC)
                                         
                 ElseIf nDIf = 0 And nRFC = 0 And nRFA = 0 Then
                    
                End If
End Function


Function DetEstado(ByVal cEstado As String) As Integer
    If cEstado = "CANC." Then
        DetEstado = 0
    Else
        DetEstado = 1
    End If
End Function
