VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCredTransARecup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferencia A Recuperaciones"
   ClientHeight    =   6195
   ClientLeft      =   2055
   ClientTop       =   1800
   ClientWidth     =   8670
   Icon            =   "frmCredTransFerARecup.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   825
      Left            =   135
      TabIndex        =   9
      Top             =   5190
      Width           =   8490
      Begin VB.CommandButton cmdImprimirActa 
         Caption         =   "&Imprimir Acta"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         TabIndex        =   20
         Top             =   255
         Width           =   1500
      End
      Begin MSComctlLib.ProgressBar PBBarra 
         Height          =   285
         Left            =   3600
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   435
         Left            =   6645
         TabIndex        =   13
         Top             =   255
         Width           =   1500
      End
      Begin VB.CommandButton CmdTransferir 
         Caption         =   "&Transferir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   11
         Top             =   255
         Width           =   1500
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3990
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   8505
      Begin VB.Frame Frame3 
         Height          =   510
         Left            =   105
         TabIndex        =   15
         Top             =   3405
         Width           =   2805
         Begin VB.OptionButton OptSelecc 
            Caption         =   "&Ninguno"
            Height          =   240
            Index           =   1
            Left            =   1530
            TabIndex        =   17
            Top             =   195
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton OptSelecc 
            Caption         =   "&Todos"
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   16
            Top             =   195
            Width           =   960
         End
      End
      Begin VB.CommandButton CmdPasarAJud 
         Caption         =   "Pasar A Judicial"
         Height          =   345
         Left            =   5100
         TabIndex        =   14
         Top             =   3480
         Width           =   1650
      End
      Begin VB.CommandButton CmdDemanda 
         Caption         =   "Con Demanda"
         Height          =   345
         Left            =   6780
         TabIndex        =   12
         Top             =   3480
         Width           =   1365
      End
      Begin SICMACT.FlexEdit FECreditos 
         Height          =   3165
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   5583
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Tras-Credito-Demanda-Cliente-Prestamo-Saldo Cap.-Refinanciado-Analista-Atraso-NroCalen-Estado"
         EncabezadosAnchos=   "400-400-2000-900-4000-1200-1200-1200-3000-1200-200-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         BackColor       =   16777215
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-L-R-R-C-C-R-C-C"
         FormatosEdit    =   "0-0-0-0-0-2-2-0-0-3-3-0"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483635
         CellBackColor   =   16777215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8520
      Begin VB.Frame Frame2 
         Caption         =   "Busqueda"
         Height          =   720
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   4500
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "&Archivo"
            Height          =   195
            Index           =   3
            Left            =   3480
            TabIndex        =   22
            Top             =   285
            Width           =   945
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "&General"
            Height          =   195
            Index           =   2
            Left            =   2520
            TabIndex        =   18
            Top             =   285
            Width           =   1185
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Por Nombre"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   3
            Top             =   285
            Width           =   1185
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Por Cuenta"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   285
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame FraCta 
         Height          =   735
         Left            =   4680
         TabIndex        =   4
         Top             =   180
         Width           =   3855
         Begin SICMACT.ActXCodCta ActxCta 
            Height          =   390
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   3720
            _ExtentX        =   6562
            _ExtentY        =   688
            Texto           =   "Credito :"
            EnabledCMAC     =   -1  'True
            EnabledCta      =   -1  'True
            EnabledProd     =   -1  'True
            EnabledAge      =   -1  'True
         End
      End
      Begin VB.Frame FraBusqNom 
         Height          =   675
         Left            =   4680
         TabIndex        =   6
         Top             =   180
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "&Buscar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2355
            TabIndex        =   7
            Top             =   195
            Width           =   1380
         End
      End
      Begin VB.Frame FraCargar 
         Height          =   675
         Left            =   4680
         TabIndex        =   23
         Top             =   180
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton CmdCargar 
            Caption         =   "Cargar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2280
            TabIndex        =   24
            Top             =   190
            Width           =   1380
         End
      End
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   330
      Left            =   3600
      TabIndex        =   21
      Top             =   5760
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   582
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmCredTransFerARecup.frx":030A
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Archivo de Excel (*.xls)|*.xls"
      FilterIndex     =   1
   End
End
Attribute VB_Name = "frmCredTransARecup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnPorcGastoTransf As Double
Dim fsMetLiquid As String
Dim sFilename As String

Private Sub CargaDatos(ByVal psBusqueda As String, ByVal pnTipoBusq As Integer)
Dim oCredito As COMDCredito.DCOMCredito
Dim oNCredito As COMNCredito.NCOMCredito
Dim R As ADODB.Recordset
Dim MatCalend As Variant

    On Error GoTo ErrorCargaDatos
    Set oCredito = New COMDCredito.DCOMCredito
    Select Case pnTipoBusq
        Case 1 'Total
            Set R = oCredito.RecuperaCreditosParaJudicalTotal
        Case 2 'Por Nombre
            Set R = oCredito.RecuperaCreditosParaJudicalTotal(psBusqueda)
        Case 3 'Por Cuenta
            Set R = oCredito.RecuperaCreditosParaJudicalTotal(, psBusqueda)
    End Select
    LimpiaFlex FECreditos
    Set oCredito = Nothing
    If R.BOF And R.EOF Then
        MsgBox "No se Encontraron Registros", vbInformation, "Aviso"
        R.Close
        Set R = Nothing
        Exit Sub
    End If
    Do While Not R.EOF
        FECreditos.AdicionaFila
        FECreditos.TextMatrix(R.Bookmark, 1) = "."
        FECreditos.TextMatrix(R.Bookmark, 2) = R!cCtaCod
        FECreditos.TextMatrix(R.Bookmark, 3) = "NO"
        FECreditos.TextMatrix(R.Bookmark, 4) = PstaNombre(R!cTitular)
        FECreditos.TextMatrix(R.Bookmark, 5) = Format(R!nMontoCol, "#0.00")
        FECreditos.TextMatrix(R.Bookmark, 6) = Format(R!nSaldo, "#0.00")
        If R!nPrdEstado = gColocEstRefMor Or R!nPrdEstado = gColocEstRefNorm Or R!nPrdEstado = gColocEstRefVenc Then
            FECreditos.TextMatrix(R.Bookmark, 7) = "SI"
        Else
            FECreditos.TextMatrix(R.Bookmark, 7) = "NO"
        End If
        FECreditos.TextMatrix(R.Bookmark, 8) = PstaNombre(IIf(IsNull(R!cAnalista), "", R!cAnalista))
        FECreditos.TextMatrix(R.Bookmark, 9) = Trim(str(R!nDiasAtraso))
        FECreditos.TextMatrix(R.Bookmark, 10) = Trim(str(R!nNroCalen))
        FECreditos.TextMatrix(R.Bookmark, 11) = Trim(str(R!nPrdEstado))
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Exit Sub

ErrorCargaDatos:
    MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CargaDatos(ActxCta.NroCuenta, 3)
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim oPers As COMDPersona.UCOMPersona 'UPersona
    If OptBusqueda(1).value Then
        Set oPers = frmBuscaPersona.Inicio
        If Not oPers Is Nothing Then
            Call CargaDatos(oPers.sPersCod, 2)
        End If
    ElseIf OptBusqueda(3).value Then
        CargaArchivoTransferir sFilename
    Else
        Call CargaDatos("", 1)
    End If
    cmdImprimirActa.Enabled = False
End Sub
'20110502
Private Sub cmdCargar_Click()
    CargaArchivoTransferir sFilename
End Sub

Sub CargaArchivoTransferir(psNomArchivo As String)
    Dim xlApp As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja As Excel.Worksheet
    Dim varMatriz As Variant
    Dim cNombreHoja As String
    Dim i As Long, n As Long
    'madm 20110702
    Dim oNCredito As COMNCredito.NCOMGasto
    Set oNCredito = New COMNCredito.NCOMGasto
    Dim nNoCastigar As Integer
    'end madm
    i = 0
    nNoCastigar = 0
    Set xlApp = New Excel.Application
    
    If Trim(psNomArchivo) = "" Then
        MsgBox "Debe indicar la ruta del Archivo Excel", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    If Trim(psNomArchivo) <> "" Then
    
        Set xlLibro = xlApp.Workbooks.Open(psNomArchivo, True, True, , "")
        cNombreHoja = "Hoja1"
        'validar nombre de hoja
        Set xlHoja = xlApp.Worksheets(cNombreHoja)
        varMatriz = xlHoja.Range("A2:AA2000").value
        xlLibro.Close SaveChanges:=False
        xlApp.Quit
        Set xlHoja = Nothing
        Set xlLibro = Nothing
        Set xlApp = Nothing
        n = UBound(varMatriz)
        LimpiaFlex FECreditos
        For i = 8 To n
            If varMatriz(i, 1) = "" Then
                If i = 8 Then
                    MsgBox "Archivo No tiene Estructura Correcta, la informacion debe estar en la Celda (A9)", vbCritical, "Mensaje"
                Else
                    If nNoCastigar >= 1 Then
                        MsgBox "Archivo Cargado Correctamente pero hay: " & nNoCastigar & " Créditos con Estados Diferentes, señalados con Rojo", vbCritical, "Mensaje"
                    Else
                        MsgBox "Archivo Cargado Correctamente ", vbInformation, "Mensaje"
                    End If
                    Set oNCredito = Nothing
                End If
                Me.OptBusqueda(3).value = False
                psNomArchivo = ""
                Exit For
            Else
                    If varMatriz(i, 26) = "" Then
                       MsgBox "Archivo No tiene Estructura Correcta, la informacion debe a partir de la Celda (A9 - Z9)", vbCritical, "Mensaje"
                       Exit For
                    Else
                         If Trim(varMatriz(i, 22)) = oNCredito.DevolverEstadoProducto(Trim(varMatriz(i, 2))) Then
                            FECreditos.AdicionaFila
                            If FECreditos.CellBackColor <> 16777215 Then
                               FECreditos.CellBackColor = 16777215
                            End If
                            FECreditos.TextMatrix(FECreditos.Row, 1) = "S"
                            FECreditos.TextMatrix(FECreditos.Row, 2) = varMatriz(i, 2) 'cuenta
                            FECreditos.TextMatrix(FECreditos.Row, 3) = "NO"
                            FECreditos.TextMatrix(FECreditos.Row, 4) = varMatriz(i, 4) 'titular
                            FECreditos.TextMatrix(FECreditos.Row, 5) = Format(varMatriz(i, 6), "#0.00") 'R!nMontoCol
                            FECreditos.TextMatrix(FECreditos.Row, 6) = Format(varMatriz(i, 23), "#0.00") 'R!nSaldo
                            If varMatriz(i, 22) = gColocEstRefMor Or varMatriz(i, 22) = gColocEstRefNorm Or varMatriz(i, 22) = gColocEstRefVenc Then 'R!nPrdEstado
                                FECreditos.TextMatrix(FECreditos.Row, 7) = "SI"
                            Else
                                FECreditos.TextMatrix(FECreditos.Row, 7) = "NO"
                            End If
                            FECreditos.TextMatrix(FECreditos.Row, 8) = varMatriz(i, 7) ' analista
                            FECreditos.TextMatrix(FECreditos.Row, 9) = Trim(str(varMatriz(i, 11))) 'atraso
                            FECreditos.TextMatrix(FECreditos.Row, 10) = Trim(str(varMatriz(i, 25))) 'R!nNroCalen
                            FECreditos.TextMatrix(FECreditos.Row, 11) = Trim(str(varMatriz(i, 22)))
                        Else
                            nNoCastigar = nNoCastigar + 1
                            FECreditos.AdicionaFila
                              If FECreditos.CellBackColor <> vbGreen Then
                                    FECreditos.CellBackColor = vbRed
                              End If
                            FECreditos.TextMatrix(FECreditos.Row, 1) = "N"
                            FECreditos.TextMatrix(FECreditos.Row, 2) = varMatriz(i, 2) 'cuenta
                            FECreditos.TextMatrix(FECreditos.Row, 3) = "NO"
                            FECreditos.TextMatrix(FECreditos.Row, 4) = varMatriz(i, 4) 'titular
                            FECreditos.TextMatrix(FECreditos.Row, 5) = Format(varMatriz(i, 6), "#0.00") 'R!nMontoCol
                            FECreditos.TextMatrix(FECreditos.Row, 6) = Format(varMatriz(i, 23), "#0.00") 'R!nSaldo
                            If varMatriz(i, 22) = gColocEstRefMor Or varMatriz(i, 22) = gColocEstRefNorm Or varMatriz(i, 22) = gColocEstRefVenc Then 'R!nPrdEstado
                                FECreditos.TextMatrix(FECreditos.Row, 7) = "SI"
                            Else
                                FECreditos.TextMatrix(FECreditos.Row, 7) = "NO"
                            End If
                            FECreditos.TextMatrix(FECreditos.Row, 8) = varMatriz(i, 7) ' analista
                            FECreditos.TextMatrix(FECreditos.Row, 9) = Trim(str(varMatriz(i, 11))) 'atraso
                            FECreditos.TextMatrix(FECreditos.Row, 10) = Trim(str(varMatriz(i, 25))) 'R!nNroCalen
                            FECreditos.TextMatrix(FECreditos.Row, 11) = Trim(str(varMatriz(i, 22)))
                        End If
                    End If
            End If
        Next i
    End If

End Sub

Private Sub CmdDemanda_Click()
Dim nCol As Integer
    nCol = FECreditos.Col
    If Trim(FECreditos.TextMatrix(1, 1)) <> "" Then
        FECreditos.Col = 3
        If FECreditos.CellBackColor = &HD2FFFF Then
            FECreditos.CellBackColor = vbWhite
            CmdDemanda.Caption = "Con &Demanda"
            FECreditos.TextMatrix(FECreditos.Row, 3) = "NO"
        Else
            FECreditos.CellBackColor = &HD2FFFF
            CmdDemanda.Caption = "Si&n Demanda"
            FECreditos.TextMatrix(FECreditos.Row, 3) = "SI"
        End If
    End If
    FECreditos.Col = nCol
End Sub

Private Sub cmdImprimirActa_Click()
On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim loPrevio As previo.clsprevio

Dim lnAge As Integer

    lsCadImprimir = ""
    rtfCartas.Filename = App.path & "\FormatoCarta\CartaActaTransferenciaCredito.txt"
    
    Set loImprime = New COMNColoCPig.NCOMColPRecGar
        lsCadImprimir = lsCadImprimir & nImprimeActaTransferencia(rtfCartas.Text, Format(gdFecSis, "mm/dd/yyyy"))
    Set loImprime = Nothing
    
    If Len(Trim(lsCadImprimir)) = 0 Then
        MsgBox "No se hay datos para mostrar en el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set loPrevio = New previo.clsprevio
        loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", False
    Set loPrevio = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    If Err.Number = 32755 Then
        MsgBox " Grabación Cancelada ", vbInformation, " Aviso "
    Else
        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
            " Avise al Area de Sistemas ", vbInformation, " Aviso "
    End If

End Sub

Private Sub CmdPasarAJud_Click()
Dim nCol As Integer
    nCol = FECreditos.Col
    If Trim(FECreditos.TextMatrix(FECreditos.Row, 1)) <> "" Then
        FECreditos.Col = 1
        If FECreditos.CellBackColor = vbGreen Then
            FECreditos.CellBackColor = vbWhite
            FECreditos.TextMatrix(FECreditos.Row, 1) = "."
            CmdPasarAJud.Caption = "&Pasar a Judicial"
        Else
            FECreditos.CellBackColor = vbGreen
            FECreditos.TextMatrix(FECreditos.Row, 1) = "S"
            CmdPasarAJud.Caption = "&No Pasar a Judicial"
        End If
    End If
    FECreditos.Col = nCol

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub CmdTransferir_Click()
Dim i As Integer
Dim nCol As Integer
Dim oNCred As COMNCredito.NCOMCredito
Dim nMaxBarra As Integer
Dim nContCred As Integer
Dim lsmensaje As String

Dim rs As ADODB.Recordset
'OptBusqueda

    If MsgBox("Creditos seran Transferidos a Recuperaciones, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    If Trim(FECreditos.TextMatrix(1, 1)) = "" Then
        MsgBox "No Se Encontraron Registros para Transferir", vbInformation, "Aviso"
        Exit Sub
    End If
    Screen.MousePointer = 11
    PBBarra.Visible = True
    nCol = FECreditos.Col
    FECreditos.Col = 1
    nContCred = 0
    For i = 1 To FECreditos.Rows - 1
        FECreditos.Row = i
        If FECreditos.CellBackColor = vbGreen Then
            nContCred = nContCred + 1
        End If
    Next i
    nMaxBarra = nContCred
    nContCred = 0
    
    'Mandar los Datos a un Recordset para Grabar en Lote
    Set rs = New ADODB.Recordset

    With rs
        'Crear RecordSet
        .Fields.Append "cCtaCod", adVarChar, 20
        .Fields.Append "nSaldoCap", adDouble
        .Fields.Append "nDemanda", adInteger
        .Fields.Append "nDiasAtraso", adInteger
        .Fields.Append "nNroCalen", adInteger
        .Fields.Append "nPrdEstado", adInteger
        .Open
        'Llenar Recordset
    
    For i = 1 To FECreditos.Rows - 1
        FECreditos.Row = i
        'If FECreditos.CellBackColor = vbGreen Then
        If FECreditos.CellBackColor = vbGreen Then
            nContCred = nContCred + 1
            
            .AddNew
            .Fields("cCtaCod") = FECreditos.TextMatrix(i, 2)
            .Fields("nSaldoCap") = CDbl(FECreditos.TextMatrix(i, 6))
            .Fields("nDemanda") = IIf(FECreditos.TextMatrix(i, 3) = "SI", gColRecDemandaSi, gColRecDemandaNo)
            .Fields("nDiasAtraso") = CInt(FECreditos.TextMatrix(i, 9))
            .Fields("nNroCalen") = CInt(FECreditos.TextMatrix(i, 10))
            .Fields("nPrdEstado") = CInt(FECreditos.TextMatrix(i, 11))
            
            'Call oNCred.TransferirARecuperaciones(FECreditos.TextMatrix(i, 2), CDbl(FECreditos.TextMatrix(i, 6)), _
                 IIf(FECreditos.TextMatrix(i, 3) = "SI", gColRecDemandaSi, gColRecDemandaNo), CInt(FECreditos.TextMatrix(i, 9)), _
                gdFecSis, gsCodAge, gsCodUser, CInt(FECreditos.TextMatrix(i, 10)), CInt(FECreditos.TextMatrix(i, 11)), , fsMetLiquid)
            FECreditos.CellBackColor = vbRed
            PBBarra.value = (nContCred / nMaxBarra) * 100
        End If
    Next i
    End With
    
    Set oNCred = New COMNCredito.NCOMCredito
    If Not (rs.EOF And rs.BOF) Then
        Call oNCred.TransferirARecuperacionesLote(rs, gdFecSis, gsCodAge, gsCodUser, fsMetLiquid, gsProyectoActual, lsmensaje)
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    Else
        MsgBox "Seleccione un Crédito", vbInformation, "Aviso"
        Screen.MousePointer = 0
        Exit Sub
    End If
                
    Set oNCred = Nothing
    FECreditos.Col = nCol
    PBBarra.Visible = False
    CmdPasarAJud.Enabled = False
    CmdDemanda.Enabled = False
    CmdTransferir.Enabled = False
    cmdImprimirActa.Enabled = True
            
    Screen.MousePointer = 0
    
    Call Impresion
    
End Sub

Private Sub Impresion()
Dim oPrev As previo.clsprevio
Dim sCad As String
Dim i As Integer

Dim loImpre As COMNColocRec.NCOMColRecImpre
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

   With rs
        'Crear RecordSet
        .Fields.Append "cTran", adVarChar, 2
        .Fields.Append "cCtaCod", adVarChar, 20
        .Fields.Append "cCond", adVarChar, 2
        .Fields.Append "cCliente", adVarChar, 100
        .Fields.Append "nMonto", adCurrency
        .Fields.Append "nSaldo", adCurrency
        .Fields.Append "cEstado", adVarChar, 2
        .Open
        'Llenar Recordset
        For i = 1 To FECreditos.Rows - 1
            FECreditos.Row = i
            FECreditos.Col = 1
            If FECreditos.CellBackColor = vbRed Then
                .AddNew
                .Fields("cTran") = IIf(IsNull(FECreditos.TextMatrix(i, 1)), "", FECreditos.TextMatrix(i, 1))
                .Fields("cCtaCod") = IIf(IsNull(FECreditos.TextMatrix(i, 2)), "", FECreditos.TextMatrix(i, 2))
                .Fields("cCond") = IIf(IsNull(FECreditos.TextMatrix(i, 3)), "", FECreditos.TextMatrix(i, 3))
                .Fields("cCliente") = IIf(IsNull(FECreditos.TextMatrix(i, 4)), "", FECreditos.TextMatrix(i, 4))
                .Fields("nMonto") = IIf(IsNull(FECreditos.TextMatrix(i, 5)), 0, FECreditos.TextMatrix(i, 5))
                .Fields("nSaldo") = IIf(IsNull(FECreditos.TextMatrix(i, 6)), 0, FECreditos.TextMatrix(i, 6))
                .Fields("cEstado") = IIf(IsNull(FECreditos.TextMatrix(i, 7)), 0, FECreditos.TextMatrix(i, 7))
            End If
        Next i
    End With
    
    Set loImpre = New COMNColocRec.NCOMColRecImpre
        sCad = loImpre.ImpresionTransferencia(rs, gsNomCmac, gdFecSis, gsNomAge, gsCodUser)
    Set loImpre = Nothing
    
    rs.Close
    
'    sCad = Chr$(10)
'    sCad = sCad & Space(2) & gsNomCmac & Space(85 - Len(gsNomCmac)) & gdFecSis & Chr$(10)
'    sCad = sCad & Space(2) & gsNomAge & Space(85 - Len(gsNomAge)) & gsCodUser & Chr$(10) & Chr$(10)
'    sCad = sCad & Space(40) & " TRANSFERENCIA A JUDICIAL " & Chr$(10)
'    sCad = sCad & Space(40) & String(30, "-") & Chr$(10)
'    sCad = sCad & Chr$(10) & Chr$(10)
'    sCad = sCad & Space(6) & ImpreFormat("CREDITO", 20) & ImpreFormat("DEMANDA", 7) & ImpreFormat("CLIENTE", 34)
'    sCad = sCad & ImpreFormat("PRESTAMO", 12) & ImpreFormat("SALDO", 5) & ImpreFormat("REFINAN", 10) & Chr$(10)
'    sCad = sCad & Space(5) & String(110, "-") & Chr$(10)
'
'    For i = 1 To FECreditos.Rows - 1
'        FECreditos.Row = i
'        FECreditos.Col = 1
'        'If FECreditos.CellBackColor = vbRed Then
'        If FECreditos.TextMatrix(i, 1) = "S" Or FECreditos.CellBackColor = vbRed Then
'            sCad = sCad & Space(2) & ImpreFormat(i, 4, 0, False) & ImpreFormat(Me.FECreditos.TextMatrix(i, 2), 20)
'            sCad = sCad & ImpreFormat(Me.FECreditos.TextMatrix(i, 3), 7)
'            sCad = sCad & ImpreFormat(Me.FECreditos.TextMatrix(i, 4), 30)
'            sCad = sCad & ImpreFormat(CDbl(Me.FECreditos.TextMatrix(i, 5)), 10, , True)
'            sCad = sCad & ImpreFormat(CDbl(Me.FECreditos.TextMatrix(i, 6)), 10, , True)
'            sCad = sCad & ImpreFormat(Me.FECreditos.TextMatrix(i, 7), 10) & Chr$(10)
'        End If
'    Next i
'    sCad = sCad & Space(5) & String(110, "-") & Chr$(10)
    
    Set oPrev = New previo.clsprevio
    oPrev.Show sCad, "Transferencia A Recuperaciones"
    Set oPrev = Nothing
    
    
End Sub

Private Sub FECreditos_RowColChange()
Dim nCol As Integer
    If Trim(FECreditos.TextMatrix(1, 1)) = "" Then
        Exit Sub
    End If
    nCol = FECreditos.Col
    FECreditos.Col = 1
    If FECreditos.CellBackColor = vbRed Then
        CmdPasarAJud.Enabled = False
        CmdDemanda.Enabled = False
    Else
        CmdPasarAJud.Enabled = True
        CmdDemanda.Enabled = True
    End If
    
    If FECreditos.CellBackColor = vbGreen Then
        CmdPasarAJud.Caption = "&No Pasar a Judicial"
    Else
        CmdPasarAJud.Caption = "&Pasar a Judicial"
    End If
    
    FECreditos.Col = 3
    If FECreditos.CellBackColor = &HD2FFFF Then
        CmdDemanda.Caption = "Si&n Demanda"
    Else
        CmdDemanda.Caption = "Con &Demanda"
    End If
    
    FECreditos.Col = nCol
End Sub

Private Sub Form_Load()
    CentraSdi Me
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Call fCargaParametro
End Sub

Private Sub OptBusqueda_Click(Index As Integer)
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    If Index = 0 Then
        FraCta.Visible = True
        FraBusqNom.Visible = False
        FraCargar.Visible = False
        CmdCargar.Visible = False
    ElseIf Index = 3 Then
        FraBusqNom.Visible = False
        FraCta.Visible = False
        FraCargar.Visible = True
        CmdCargar.Visible = True
        CommonDialog1.ShowOpen
        sFilename = CommonDialog1.Filename
    Else
        FraCta.Visible = False
        FraBusqNom.Visible = True
        FraCargar.Visible = False
        CmdCargar.Visible = False
    End If
    cmdImprimirActa.Enabled = False
End Sub

Private Sub OptSelecc_Click(Index As Integer)
Dim i As Integer
Dim nCol As Integer
    nCol = FECreditos.Col
    If FECreditos.TextMatrix(1, 1) <> "" Then
        If Index = 0 Then
            For i = 1 To FECreditos.Rows - 1
                FECreditos.Col = 1
                FECreditos.Row = i
                If FECreditos.CellBackColor <> vbRed Then
                    FECreditos.CellBackColor = vbGreen
                    CmdPasarAJud.Caption = "&No Pasar A Judicial"
                End If
            Next i
        Else
            For i = 1 To FECreditos.Rows - 1
                FECreditos.Col = 1
                FECreditos.Row = i
                If FECreditos.CellBackColor <> vbRed Then
                    FECreditos.CellBackColor = vbWhite
                    CmdPasarAJud.Caption = "&Pasar A Judicial"
                End If
            Next i
            
        End If
    End If
    FECreditos.Col = nCol
End Sub

Private Sub fCargaParametro()
'Dim lsSql As String
'Dim lrParam As ADODB.Recordset
'Dim loConec As COMConecta.DCOMConecta
'
'lsSql = " Select nValor  from ProductoConcepto  where nprdconceptocod like '3211' "
'
'Set loConec = New COMConecta.DCOMConecta
'    loConec.AbreConexion
'    'Set lrParam = New ADODB.Recordset
'    Set lrParam = loConec.CargaRecordSet(lsSql)
'Set loConec = Nothing
'If lrParam.BOF And lrParam.EOF Then
'    fnPorcGastoTransf = 0
'Else
'    fnPorcGastoTransf = lrParam!nValor
'End If
'Set lrParam = Nothing
'
'Dim loParam As COMDConstSistema.NCOMConstSistema
'Set loParam = New COMDConstSistema.NCOMConstSistema
'fsMetLiquid = loParam.LeeConstSistema(153)
'Set loParam = Nothing

Dim oCred As COMNCredito.NCOMCredito
Set oCred = New COMNCredito.NCOMCredito
Call oCred.CargarParametrosTransferencia(fnPorcGastoTransf, fsMetLiquid)
Set oCred = Nothing
End Sub

'Imprime Acta de Transferencia
Public Function nImprimeActaTransferencia(ByVal psTextoCarta As String, ByVal psFecha As String) As String
'Dim lsCadImp As String
'Dim lsCartaModelo As String
'Dim lsFechaHoraGrab As String
'Dim liItem As Integer
'Dim lsListaTransfer As String
'Dim lsTotalTransfer As String
'Dim lnTotalTransfer As Double
Dim i As Integer

Dim loImpre As COMNColocRec.NCOMColRecImpre
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset


    With rs
        'Crear RecordSet
        .Fields.Append "cTran", adVarChar, 2
        .Fields.Append "cCtaCod", adVarChar, 20
        .Fields.Append "cCliente", adVarChar, 100
        .Fields.Append "nSaldo", adCurrency
        .Fields.Append "nDiasA", adInteger
        .Open
        'Llenar Recordset
        For i = 1 To FECreditos.Rows - 1
            FECreditos.Row = i
            FECreditos.Col = 1
            If FECreditos.CellBackColor = vbRed Then
                .AddNew
                .Fields("cTran") = IIf(IsNull(FECreditos.TextMatrix(i, 1)), "", FECreditos.TextMatrix(i, 1))
                .Fields("cCtaCod") = IIf(IsNull(FECreditos.TextMatrix(i, 2)), "", FECreditos.TextMatrix(i, 2))
                .Fields("cCliente") = IIf(IsNull(FECreditos.TextMatrix(i, 4)), "", FECreditos.TextMatrix(i, 4))
                .Fields("nSaldo") = IIf(IsNull(FECreditos.TextMatrix(i, 6)), 0, FECreditos.TextMatrix(i, 6))
                .Fields("nDiasA") = IIf(IsNull(FECreditos.TextMatrix(i, 9)), 0, FECreditos.TextMatrix(i, 9))
            End If
        Next i
    End With

    Set loImpre = New COMNColocRec.NCOMColRecImpre
        nImprimeActaTransferencia = loImpre.nImprimeActaTransferencia(psTextoCarta, psFecha, rs)
    Set loImpre = Nothing

    rs.Close

'    lsListaTransfer = lsListaTransfer & Space(8) & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34)
'    lsListaTransfer = lsListaTransfer & ImpreFormat("SALDO", 5) & ImpreFormat("ATRASO", 15) & Chr$(10)
'    lsListaTransfer = lsListaTransfer & Space(2) & String(80, "-") & Chr$(10)
'
'    For i = 1 To FECreditos.Rows - 1
'        FECreditos.Row = i
'        'If FECreditos.CellBackColor = vbRed Then
'        If FECreditos.TextMatrix(i, 1) = "S" Then
'            lsListaTransfer = lsListaTransfer & Space(2) & ImpreFormat(i, 4, 0, False) & ImpreFormat(Me.FECreditos.TextMatrix(i, 2), 20)
'            lsListaTransfer = lsListaTransfer & ImpreFormat(Me.FECreditos.TextMatrix(i, 4), 30)
'            lsListaTransfer = lsListaTransfer & ImpreFormat(CDbl(Me.FECreditos.TextMatrix(i, 6)), 10, , True)
'            lsListaTransfer = lsListaTransfer & ImpreFormat(Me.FECreditos.TextMatrix(i, 9), 8) & Chr$(10)
'            'lsListaTransfer = lsListaTransfer & ImpreFormat(Me.FECreditos.TextMatrix(i, 7), 10)
'            'Acumulo el total
'            lnTotalTransfer = lnTotalTransfer + ImpreFormat(CDbl(Me.FECreditos.TextMatrix(i, 6)), 10, , True)
'        End If
'    Next
'    lsListaTransfer = lsListaTransfer & Space(2) & String(80, "-") & Chr$(10)
'    lsTotalTransfer = "Total Creditos : " & ImpreFormat(i - 1, 5, 0, False) & Space(20) & "Saldo Capital : " & ImpreFormat(lnTotalTransfer, 12, 2, True)
'
'    'Llena cartas
'    lsCartaModelo = psTextoCarta
'    lsCartaModelo = Replace(lsCartaModelo, "<<FECHAC>>", Format(psFecha, "dd/mm/yyyy"), , 1, vbTextCompare)
'    lsCartaModelo = Replace(lsCartaModelo, "<<FECHAL>>", Format(psFecha, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
'    lsCartaModelo = Replace(lsCartaModelo, "<<LISTA>>", lsListaTransfer, , 1, vbTextCompare)
'    lsCartaModelo = Replace(lsCartaModelo, "<<TOTAL>>", lsTotalTransfer, , 1, vbTextCompare)
'
'    lsCadImp = lsCadImp & lsCartaModelo & Chr(12)
    

End Function

