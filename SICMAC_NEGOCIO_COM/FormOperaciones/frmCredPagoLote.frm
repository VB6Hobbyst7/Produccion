VERSION 5.00
Object = "{2E37A9B8-9906-4590-9F3A-E67DB58122F5}#7.0#0"; "OcxLabelX.ocx"
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form frmCredPagoLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pago en Lote"
   ClientHeight    =   7095
   ClientLeft      =   705
   ClientTop       =   1725
   ClientWidth     =   11430
   Icon            =   "frmCredPagoLote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   750
      Left            =   135
      TabIndex        =   16
      Top             =   6270
      Width           =   11220
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   390
         Left            =   9825
         TabIndex        =   21
         Top             =   210
         Width           =   1305
      End
      Begin VB.CommandButton CmdCargaArch 
         Caption         =   "&Cargar Archivo"
         Height          =   390
         Left            =   4140
         TabIndex        =   20
         Top             =   210
         Width           =   1785
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   390
         Left            =   2805
         TabIndex        =   19
         Top             =   210
         Width           =   1305
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   1470
         TabIndex        =   18
         Top             =   210
         Width           =   1305
      End
      Begin VB.CommandButton CmdPagar 
         Caption         =   "&Pagar"
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
         Height          =   390
         Left            =   135
         TabIndex        =   17
         Top             =   210
         Width           =   1305
      End
      Begin VB.Label lblMens 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Left            =   6060
         TabIndex        =   29
         Top             =   165
         Visible         =   0   'False
         Width           =   3600
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5070
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   11220
      Begin VB.OptionButton optmoneda 
         Caption         =   "Dolares"
         Height          =   285
         Index           =   1
         Left            =   7665
         TabIndex        =   31
         Top             =   180
         Width           =   1035
      End
      Begin VB.OptionButton optmoneda 
         Caption         =   "Soles"
         Height          =   285
         Index           =   0
         Left            =   6720
         TabIndex        =   30
         Top             =   180
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.CheckBox ChkImpBol 
         Caption         =   "Imprimir Boletas"
         Height          =   270
         Left            =   9615
         TabIndex        =   28
         Top             =   4680
         Width           =   1470
      End
      Begin SICMACT.FlexEdit FEPagoLote 
         Height          =   4125
         Left            =   120
         TabIndex        =   11
         Top             =   525
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   7276
         Cols0           =   23
         HighLight       =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmCredPagoLote.frx":030A
         EncabezadosAnchos=   "400-400-2200-1200-900-1200-4500-1200-1200-800-1200-0-0-0-0-1200-1200-1200-1200-1200-1200-0-0"
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
         ColumnasAEditar =   "X-1-X-X-X-5-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   65535
         BackColorControl=   65535
         BackColorControl=   65535
         EncabezadosAlineacion=   "C-C-C-C-C-R-L-R-R-C-R-L-C-C-C-R-R-R-R-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-2-0-2-2-0-2-0-0-0-0-2-2-2-2-2-5-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "Nº"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483635
      End
      Begin VB.CheckBox ChkCobrarMora 
         Caption         =   "&Cobrar Mora"
         Height          =   195
         Left            =   9750
         TabIndex        =   10
         Top             =   225
         Width           =   1245
      End
      Begin VB.OptionButton OptSelec 
         Caption         =   "&Ninguno"
         Height          =   255
         Index           =   1
         Left            =   1050
         TabIndex        =   9
         Top             =   180
         Width           =   1065
      End
      Begin VB.OptionButton OptSelec 
         Caption         =   "&Todos"
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   8
         Top             =   180
         Width           =   885
      End
      Begin VB.Label LblTotCredPag 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   300
         Left            =   1755
         TabIndex        =   15
         Top             =   4710
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total de Creditos :"
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
         Height          =   195
         Left            =   105
         TabIndex        =   14
         Top             =   4755
         Width           =   1590
      End
      Begin VB.Label LblTotPag 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   300
         Left            =   7995
         TabIndex        =   13
         Top             =   4695
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "A Pagar :"
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
         Height          =   195
         Left            =   7035
         TabIndex        =   12
         Top             =   4755
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   11220
      Begin OcxLabelX.LabelX LblxBanco 
         Height          =   435
         Left            =   3000
         TabIndex        =   25
         Top             =   645
         Visible         =   0   'False
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   767
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX LblXMonto 
         Height          =   465
         Left            =   8940
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   820
         FondoBlanco     =   0   'False
         Resalte         =   16711680
         Caption         =   "0.00"
         Bold            =   -1  'True
         Alignment       =   1
      End
      Begin VB.CommandButton CmdAplicar 
         Caption         =   "&Aplicar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   630
         Width           =   1590
      End
      Begin VB.ComboBox CmbForPag 
         Height          =   315
         Left            =   8190
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   210
         Width           =   1965
      End
      Begin VB.ComboBox CmbInstitucion 
         Height          =   315
         Left            =   1755
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   4785
      End
      Begin OcxLabelX.LabelX LblXCheque 
         Height          =   435
         Left            =   6105
         TabIndex        =   26
         Top             =   645
         Visible         =   0   'False
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   767
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
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
         Height          =   195
         Left            =   8265
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label LblBancos 
         AutoSize        =   -1  'True
         Caption         =   "Bancos"
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
         Height          =   195
         Left            =   2310
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cheque :"
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
         Height          =   195
         Left            =   5250
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago :"
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
         Height          =   195
         Left            =   6705
         TabIndex        =   3
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Institucion : "
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
         Height          =   195
         Left            =   690
         TabIndex        =   1
         Top             =   270
         Width           =   1080
      End
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgFile 
      Height          =   480
      Left            =   11475
      TabIndex        =   27
      Top             =   5175
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   847
      Filtro          =   "Archivos Excel|*.XLS|Todos los Archivos|*.*"
      Altura          =   280
   End
End
Attribute VB_Name = "frmCredPagoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TPagoLote
    CodMod As String
    nomcli As String
    PagoCli As String
    CodInst As String
    Enc As Boolean
End Type
Private MatCarPagos() As TPagoLote
Private ContMatCarPagos As Integer

Dim sPersCodBco As String
Dim sCadImpre As String
Dim nTotCred As Integer
Dim nMontoAPagar As Double
Dim nRedondeoITF As Double
'Dim WithEvents oNegCred As COMNCredito.NCOMCredito
Dim oDocRec As UDocRec 'EJVG20140226

Private Sub CargaMatrizdeArchivo()
Dim i As Integer
Dim oDCreditos As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim MadCadena() As String
Dim j As Integer
Dim MatDatosNOCargados() As Integer
Dim nNumRegNOCar As Integer

    LimpiaFlex FEPagoLote
    ReDim MadCadena(ContMatCarPagos)
    For i = 0 To ContMatCarPagos - 1
        MadCadena(i) = MatCarPagos(i).CodMod
    Next i
    Set oDCreditos = New COMDCredito.DCOMCreditos
    Set R = oDCreditos.RecuperaCreditosArchivoPagoLote(MadCadena, MatCarPagos(0).CodInst)
    nNumRegNOCar = 0
    ReDim MatDatosNOCargados(0)
    If Not R.EOF And Not R.BOF Then
        'Set FEPagoLote.Recordset = R
        j = 0
        For i = 0 To UBound(MatCarPagos) - 1
            R.MoveFirst
            R.Find " cCodModular = '" & Trim(MatCarPagos(i).CodMod) & "' "
            If Not R.EOF Then
                j = j + 1
                FEPagoLote.AdicionaFila , , True
                FEPagoLote.TextMatrix(j, 2) = Trim(R!cCtaCod)
                FEPagoLote.TextMatrix(j, 3) = Trim(R!cCodModular)
                FEPagoLote.TextMatrix(j, 4) = Trim(Str(R!nCuota))
                FEPagoLote.TextMatrix(j, 5) = "0.00"
                FEPagoLote.TextMatrix(j, 6) = PstaNombre(R!cPersNombre)
                FEPagoLote.TextMatrix(j, 7) = Format(R!nSaldoCuota, "#0.00")
                FEPagoLote.TextMatrix(j, 8) = Format(R!nMora, "#0.00")
                FEPagoLote.TextMatrix(j, 9) = Trim(Str(R!nDiasAtraso))
                FEPagoLote.TextMatrix(j, 10) = Format(R!nDeudaTotal, "#0.00")
                FEPagoLote.TextMatrix(j, 11) = Trim(R!cMetLiquidacion)
                FEPagoLote.TextMatrix(j, 12) = Trim(Str(R!nTransacc))
                FEPagoLote.TextMatrix(j, 13) = Trim(Str(R!nPrdEstado))
                FEPagoLote.TextMatrix(j, 15) = "0.00"
                FEPagoLote.TextMatrix(j, 20) = Format(R!Vencim, "dd/mm/yyyy")
            Else
                nNumRegNOCar = nNumRegNOCar + 1
                ReDim Preserve MatDatosNOCargados(nNumRegNOCar)
                MatDatosNOCargados(nNumRegNOCar - 1) = i
            End If
        Next i
        CmbInstitucion.ListIndex = IndiceListaCombo(CmbInstitucion, MatCarPagos(0).CodInst)
    End If
    FEPagoLote.lbEditarFlex = True
    If FEPagoLote.TextMatrix(1, 2) <> "" Then
        Call HabilitaDatos(True)
    End If
    For i = 1 To FEPagoLote.Rows - 1
        For j = 0 To ContMatCarPagos - 1
            If MatCarPagos(j).CodMod = Trim(FEPagoLote.TextMatrix(i, 3)) Then
                FEPagoLote.TextMatrix(i, 5) = Format(fgITFCalculaImpuestoIncluido(CDbl(MatCarPagos(j).PagoCli)), "0.00")
                FEPagoLote.TextMatrix(i, 15) = Format(CDbl(MatCarPagos(j).PagoCli) - fgITFCalculaImpuestoIncluido(CDbl(MatCarPagos(j).PagoCli)), "0.00")
                'FEPagoLote.TextMatrix(I, 5) = Format(CDbl(MatCarPagos(j).PagoCli), "0.00")
                'FEPagoLote.TextMatrix(I, 15) = Format(CDbl(MatCarPagos(j).PagoCli) - fgITFCalculaImpuestoIncluido(CDbl(MatCarPagos(j).PagoCli)), "0.00")
            End If
        Next j
    Next i
    Set oDCreditos = Nothing
    
    Dim nSumTot As Double
    Dim sCadImpre As String
    'Imprime Creditos que no se cargaron
    nSumTot = 0
    sCadImpre = Chr(10)
    sCadImpre = sCadImpre & Space(30) & "LISTADO DE CREDITOS QUE NO SE CARGARON PARA EL PAGO EN LOTE" & Chr(10)
    sCadImpre = sCadImpre & Space(30) & String(50, "-") & Chr(10) & Chr(10) & Chr(10)
    sCadImpre = sCadImpre & Space(5) & String(80, "-") & Chr(10)
    sCadImpre = sCadImpre & Space(5) & ImpreFormat("Item", 4) & ImpreFormat("CODIGO MOD", 15) & ImpreFormat("CLIENTE", 40) & ImpreFormat("PAGO", 15) & Chr(10)
    sCadImpre = sCadImpre & Space(5) & String(80, "-") & Chr(10)
    For i = 0 To UBound(MatDatosNOCargados) - 1
        sCadImpre = sCadImpre & Space(5) & ImpreFormat(i + 1, 4, 0)
        sCadImpre = sCadImpre & ImpreFormat(MatCarPagos(MatDatosNOCargados(i)).CodMod, 15)
        sCadImpre = sCadImpre & ImpreFormat(MatCarPagos(MatDatosNOCargados(i)).nomcli, 40)
        sCadImpre = sCadImpre & ImpreFormat(CDbl(MatCarPagos(MatDatosNOCargados(i)).PagoCli), 8, 2, True) & Chr(10)
        nSumTot = nSumTot + MatCarPagos(MatDatosNOCargados(i)).PagoCli
    Next i
    sCadImpre = sCadImpre & Space(5) & String(80, "-") & Chr(10) & Chr(10)
    sCadImpre = sCadImpre & Space(5) & "TOTAL : " & ImpreFormat(nSumTot, 12) & Chr(10)
    Dim oPrev As New previo.clsprevio
    Call oPrev.Show(sCadImpre, "Creditos No Cargados")
    Set oPrev = Nothing
    
End Sub

Private Sub EliminaEspacio(CodMod As String)
Dim i As Integer
Dim cad As String
    cad = ""
    For i = 1 To Len(CodMod)
        If Mid(CodMod, i, 1) <> " " And Mid(CodMod, i, 1) <> "-" Then
            cad = cad + Mid(CodMod, i, 1)
        End If
    Next i
    CodMod = cad
End Sub

Private Sub CargaArchivo(ByVal psRuta As String)
Dim xlsApp As Excel.Application
Dim xlsBok As Excel.Workbook
Dim xlsSht As Excel.Worksheet
Dim Fin As Boolean
Dim ContEsp As Integer
Dim Ruta As String
Dim Colum As Integer
Dim fila As Integer
Dim ColCodMod() As Integer
Dim IndCodMod As Integer
Dim ColCli As Integer
Dim ColPago As Integer
Dim CodMod As String
Dim nomcli As String
Dim PagoCli As String
Dim TotalFilas As Integer
Dim Institucion As String
Dim i As Integer
    Ruta = psRuta
    If psRuta = "" Then
        MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Set xlsApp = New Excel.Application
    Set xlsBok = xlsApp.Workbooks.Open(Ruta, , True)
    Set xlsSht = xlsBok.Worksheets(1)
        
    'Reconocimiento de Estructura de Archivo Excel
    ColCli = 0
    ColPago = 0
    IndCodMod = 0
    ReDim ColCodMod(0)
    'Columnas que conforman el codigo Modular
    Colum = 1
    
    Fin = False
    Do While Not Fin
        If UCase(Mid(Trim(xlsSht.Cells(1, Colum)), 1, 2)) = "CM" Then
            IndCodMod = IndCodMod + 1
            ReDim Preserve ColCodMod(IndCodMod)
            ColCodMod(IndCodMod - 1) = Colum
            Colum = Colum + 1
        Else
            Fin = True
        End If
    Loop
    
    'Columna que conforma el nombre del cliente o pago
    If UCase(Trim(xlsSht.Cells(1, Colum))) = "CLIENTE" Then
        ColCli = Colum
    Else
        If UCase(Trim(xlsSht.Cells(1, Colum))) = "PAGO" Then
            ColPago = Colum
        End If
    End If
    Colum = Colum + 1
    
    If UCase(Trim(xlsSht.Cells(1, Colum))) = "PAGO" Then
        ColPago = Colum
    Else
        If UCase(Trim(xlsSht.Cells(1, Colum))) = "CLIENTE" Then
            ColCli = Colum
        End If
    End If
    Colum = Colum + 1
    
    Institucion = Trim(xlsSht.Cells(1, Colum))
    Colum = Colum + 1
    
    'Verificando  estructura
    If ColCli = 0 Then
        MsgBox "No se Encontro Columna de Cliente en el Documento de Excel", vbInformation, "Aviso"
            xlsBok.Close
            xlsApp.Quit
            Set xlsApp = Nothing
            Set xlsBok = Nothing
            Set xlsSht = Nothing
            Screen.MousePointer = 0
            Exit Sub
    End If
    If ColPago = 0 Then
        MsgBox "No se Encontro Columna de Pago en el Documento de Excel", vbInformation, "Aviso"
            xlsBok.Close
            xlsApp.Quit
            Set xlsApp = Nothing
            Set xlsBok = Nothing
            Set xlsSht = Nothing
            Screen.MousePointer = 0
            Exit Sub
    End If
    If IndCodMod = 0 Then
        MsgBox "No se Encontro Columna de Codigo Modular en el Documento de Excel", vbInformation, "Aviso"
            xlsBok.Close
            xlsApp.Quit
            Set xlsApp = Nothing
            Set xlsBok = Nothing
            Set xlsSht = Nothing
            Screen.MousePointer = 0
            Exit Sub
    End If
    
    If Len(Trim(Institucion)) = 0 Then
        MsgBox "No se Encontro Codigo de Institucion", vbInformation, "Aviso"
            xlsBok.Close
            xlsApp.Quit
            Set xlsApp = Nothing
            Set xlsBok = Nothing
            Set xlsSht = Nothing
            Screen.MousePointer = 0
            Exit Sub
    End If
    'Proceso de carga
    'para la barra de progreso Halla el total de filas
        Fin = False
        TotalFilas = 0
        fila = 2
        Do While Not Fin
            If Trim(xlsSht.Cells(fila, 1)) = "" Then
                ContEsp = ContEsp + 1
                If ContEsp >= 5 Then
                    Fin = True
                End If
            Else
                TotalFilas = TotalFilas + 1
            End If
             fila = fila + 1
        Loop
    
    'Carga
    Fin = False
    ContEsp = 0
    fila = 2
    ContMatCarPagos = 0
    Do While Not Fin
        If Trim(xlsSht.Cells(fila, 1)) = "" Then
            ContEsp = ContEsp + 1
            If ContEsp >= 5 Then
                Fin = True
            End If
        Else
            'Cargar una Matriz con los datos
            ContEsp = 0
            'Cargando columnas
            CodMod = ""
            For i = 0 To IndCodMod - 1
                CodMod = CodMod + Trim(xlsSht.Cells(fila, ColCodMod(i)))
            Next i
            Call EliminaEspacio(CodMod)
            'carga cliente
            nomcli = ""
            nomcli = Trim(xlsSht.Cells(fila, ColCli))
            'carga pago
            PagoCli = ""
            PagoCli = Trim(xlsSht.Cells(fila, ColPago))
            'Carga a matriz
            ContMatCarPagos = ContMatCarPagos + 1
            ReDim Preserve MatCarPagos(ContMatCarPagos)
            MatCarPagos(ContMatCarPagos - 1).CodMod = CodMod
            MatCarPagos(ContMatCarPagos - 1).nomcli = nomcli
            MatCarPagos(ContMatCarPagos - 1).PagoCli = PagoCli
            MatCarPagos(ContMatCarPagos - 1).CodInst = Institucion
            MatCarPagos(ContMatCarPagos - 1).Enc = False
        End If
        fila = fila + 1
    Loop
    xlsBok.Close
    xlsApp.Quit
    Set xlsApp = Nothing
    Set xlsBok = Nothing
    Set xlsSht = Nothing
    Screen.MousePointer = 0

End Sub

Public Sub HabilitaDatos(ByVal pbHabilita As Boolean)
    Frame1.Enabled = Not pbHabilita
    Frame2.Enabled = pbHabilita
    CmdPagar.Enabled = pbHabilita
    CmdImprimir.Enabled = Not pbHabilita
    CmdCargaArch.Enabled = Not pbHabilita
    FEPagoLote.Enabled = pbHabilita
    FEPagoLote.lbEditarFlex = pbHabilita
End Sub

Private Sub LimpiaDatos()
    Call InicializaCombos(Me)
    'Call LimpiaFlex(FEPagoLote)
    FormateaFlex FEPagoLote 'EJVG20140226
    LblXMonto.Caption = "0.00"
    LblTotCredPag.Caption = "0.00"
    LblTotPag.Caption = "0.00"
    ChkCobrarMora.value = 0
    ChkImpBol.value = 0
    LblBancos.Visible = False
    LblxBanco.Visible = False
    Label3.Visible = False
    LblXCheque.Visible = False
    Label6.Visible = False
    LblXMonto.Visible = False
End Sub

Private Sub CargaControles()
 Dim oCredito As COMDCredito.DCOMCredito
 Set oCredito = New COMDCredito.DCOMCredito
 Dim rsInstit As ADODB.Recordset
 Dim rsTipoPago As ADODB.Recordset
    'Carga Instituciones
'    Call CargaComboPersonasTipo(gPersTipoConvenio, cmbInstitucion)
    'Carga Formas de Pago
'    Call CargaComboConstante(gColocTipoPago, CmbForPag)
    Call oCredito.CargarControlesPagoLote(rsInstit, rsTipoPago)
    Call Llenar_Combo_con_Recordset(rsTipoPago, CmbForPag)
    Set oCredito = Nothing
    
    CmbInstitucion.Clear
    Do While Not rsInstit.EOF
        CmbInstitucion.AddItem PstaNombre(rsInstit!cPersNombre) & Space(250) & rsInstit!cPersCod
        rsInstit.MoveNext
    Loop

End Sub

'Private Sub CargaDatos(ByVal psCodInstitucion As String)
'Dim oDCred As DCreditos
'Dim R As ADODB.Recordset
'Dim i As Integer
'Dim nMontoGastoCargo As Double
'Dim nNumGastosCancel As Integer
'Dim oGastos As nGasto
'Dim MatGastosCancelacion As Variant
'Dim MatCalend As Variant
'Dim oNegCredito As NCredito
'
'    On Error GoTo ErrorCargaDatos
'    Set oDCred = New DCredito
'    Set R = oDCred.RecuperaCreditosPagoLote(psCodInstitucion)
'    Set oDCred = Nothing
'    Set FEPagoLote.Recordset = R
'
'    Set oGastos = New nGasto
'    Set oNegCredito = New NCredito
'    For i = 1 To FEPagoLote.Rows - 1
'        MatCalend = oNegCredito.RecuperaMatrizCalendarioPendiente(FEPagoLote.TextMatrix(i, 2))
'        MatGastosCancelacion = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosCancel, gdFecSis, psCtaCod, 1, "CA", , , CDbl(FEPagoLote.TextMatrix(i, 10)), oNegCredito.MatrizMontoCapitalAPagar(MatCalend, gdFecSis), MatCalend(0, 2), , , , , R!nDiasAtraso)
'        nMontoGastoCargo = MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("CA", "PA", ""))
'        FEPagoLote.TextMatrix(i, 5) = Format(CDbl(FEPagoLote.TextMatrix(i, 5)) + nMontoGastoCargo, "#0.00")
'        FEPagoLote.TextMatrix(i, 10) = Format(CDbl(FEPagoLote.TextMatrix(i, 10)) + nMontoGastoCargo, "#0.00")
'    Next i
'    Set oGastos = Nothing
'    Set oNegCredito = Nothing
'    Exit Sub
'
'ErrorCargaDatos:
'    MsgBox Err.Description, vbCritical, "Aviso"
'
'End Sub

'Private Sub ChkCobrarMora_Click()
''Esto he agregado -- CAFF
'Dim i As Integer
'
'    If ChkCobrarMora.value Then
'        nTotCred = 0
'        nMontoAPagar = 0
'        For i = 1 To FEPagoLote.Rows - 1
'            If FEPagoLote.TextMatrix(i, 1) = "." Then
'                nTotCred = nTotCred + 1
'                nMontoAPagar = nMontoAPagar + CDbl(FEPagoLote.TextMatrix(i, 7)) + CDbl(FEPagoLote.TextMatrix(i, 8))
'            End If
'        Next i
'        LblTotCredPag.Caption = nTotCred
'        LblTotPag.Caption = Format(nMontoAPagar, "#0.00")
'    Else
'        nTotCred = 0
'        nMontoAPagar = 0
'        For i = 1 To FEPagoLote.Rows - 1
'            If FEPagoLote.TextMatrix(i, 1) = "." Then
'                nTotCred = nTotCred + 1
'                nMontoAPagar = nMontoAPagar + CDbl(FEPagoLote.TextMatrix(i, 7))
'            End If
'        Next i
'        LblTotCredPag.Caption = nTotCred
'        LblTotPag.Caption = Format(nMontoAPagar, "#0.00")
'    End If
'End Sub

Private Sub CmbForPag_Click()
Dim MatDatos As Variant
    
    If CmbForPag.ListIndex = -1 Then
        Exit Sub
    End If
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCheque Then
        LblBancos.Visible = True
        LblxBanco.Visible = True
        Label3.Visible = True
        LblXCheque.Visible = True
        Label6.Visible = True
        LblXMonto.Visible = True
        'MatDatos(0) 'Descripcion Banco
        'MatDatos(1) 'Codigo de Persona
        'MatDatos(2) =  'Moneda
        'MatDatos(3) = 'Monto
        'MatDatos(4) = 'Nro de Cheque
        'EJVG20140226 ***
        Dim oForm As New frmChequeBusqueda
        Set oDocRec = oForm.Iniciar(IIf(optmoneda(0).value = True, gMonedaNacional, gMonedaExtranjera), CRED_PagoLote)
        Set oForm = Nothing
        LblxBanco.Caption = oDocRec.fsPersNombre
        sPersCodBco = oDocRec.fsPersCod
        If optmoneda(0).value = True Then
            LblXMonto.Resalte = eAzul
        Else
            LblXMonto.Resalte = eVerde
        End If
        LblXCheque.Caption = oDocRec.fsNroDoc
        LblXMonto.Caption = Format(oDocRec.fnMonto, gsFormatoNumeroView)
        'MatDatos = frmBuscaCheque.BuscaCheque(gChqEstEnValorizacion, , 0)
        'If MatDatos(0) <> "" Then
        '    LblxBanco.Caption = MatDatos(0)
        '    sPersCodBco = MatDatos(1)
        '    If MatDatos(2) = "1" Then
        '        LblXMonto.Resalte = eAzul
        '    Else
        '        LblXMonto.Resalte = eVerde
        '    End If
        '    LblXCheque.Caption = MatDatos(4)
        '    LblXMonto.Caption = MatDatos(3)
        'Else
        '    LblxBanco.Caption = ""
        '    sPersCodBco = ""
        '    LblXCheque.Caption = ""
        '    LblXMonto.Resalte = eNegro
        '    LblXMonto.Caption = "0.00"
        'End If
        'END EJVG *******
    Else
        LblBancos.Visible = False
        LblxBanco.Visible = False
        Label3.Visible = False
        LblXCheque.Visible = False
        Label6.Visible = False
        LblXMonto.Visible = False
    End If
End Sub

Private Sub cmdAplicar_Click()

'Dim oCred As COMdCredito.DCOMCreditos
Dim R As ADODB.Recordset
'Dim oGastos As COMNCredito.NCOMGasto
'Dim oNegCredito As COMNCredito.NCOMCredito
'Dim MatCalend As Variant
'Dim MatGastosCancelacion As Variant
'Dim nMontoGastoCargo As Double
Dim i As Integer
'Dim nNumGastosCancel As Integer
Dim pnMoneda As Moneda
Dim oCred As COMNCredito.NCOMCredito
Dim MatPago As Variant
Dim MatDeuda As Variant

    If CmbForPag.ListIndex = -1 Then
       MsgBox "Seleccione la forma de Pago", vbInformation, "Aviso"
       CmbForPag.SetFocus
       Exit Sub
    End If

    If optmoneda(0).value = True Then
        pnMoneda = gMonedaNacional
    Else
        pnMoneda = gMonedaExtranjera
    End If
    'Set oCred = New COMdCredito.DCOMCreditos
    'Set R = oCred.RecuperaCreditosPagoLote(Trim(Right(CmbInstitucion.Text, 20)), pnMoneda)
    'Set oCred = Nothing
    Set oCred = New COMNCredito.NCOMCredito
    Call oCred.AplicarPagoLote(Trim(Right(CmbInstitucion.Text, 20)), pnMoneda, gdFecSis, R, MatPago, MatDeuda, IIf(ChkCobrarMora.value = 1, True, False))
    Set oCred = Nothing
    
    FEPagoLote.Clear
    FEPagoLote.FormaCabecera
    FEPagoLote.Rows = 2
    If Not R.EOF And Not R.BOF Then
        Set FEPagoLote.Recordset = R
    End If
    FEPagoLote.lbEditarFlex = True
    If FEPagoLote.TextMatrix(1, 3) <> "" Or R.RecordCount > 0 Then
        Call HabilitaDatos(True)
    Else
        MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
        Exit Sub
    End If
    nTotCred = 0
    nMontoAPagar = 0
    
    '**************************************************
    'Carga Gastos
    '**************************************************
'    Set oGastos = New COMNCredito.NCOMGasto
'    Set oNegCredito = New COMNCredito.NCOMCredito
    Me.Enabled = False
    lblMens.Visible = True
    Me.MousePointer = 11
    For i = 1 To FEPagoLote.Rows - 1
        lblMens.Caption = "Procesando Gastos Credito N°:" & FEPagoLote.TextMatrix(i, 2) & " [" & i & "/" & FEPagoLote.Rows - 1 & "]"
'        lblMens.Refresh
'        MatCalend = oNegCredito.RecuperaMatrizCalendarioPendiente(FEPagoLote.TextMatrix(i, 2))
'        MatGastosCancelacion = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosCancel, gdFecSis, FEPagoLote.TextMatrix(i, 2), 1, "CA", , , CDbl(FEPagoLote.TextMatrix(i, 10)), oNegCredito.MatrizMontoCapitalAPagar(MatCalend, gdFecSis), MatCalend(0, 2), , , , , CInt(FEPagoLote.TextMatrix(i, 9)))
'        nMontoGastoCargo = MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("CA", "PA", ""))
'        FEPagoLote.TextMatrix(i, 5) = Format(CDbl(FEPagoLote.TextMatrix(i, 5)) + nMontoGastoCargo, "#0.00")
'        FEPagoLote.TextMatrix(i, 10) = Format(CDbl(FEPagoLote.TextMatrix(i, 10)) + nMontoGastoCargo, "#0.00")

'ARCV 24-04-2007 (Gastos de Cancelacion Manuales)
'        FEPagoLote.TextMatrix(i, 5) = MatPago(i)
'        FEPagoLote.TextMatrix(i, 10) = MatDeuda(i)
'---
    Next i
    Me.MousePointer = 0
    Me.Enabled = True
    lblMens.Visible = False
    
'    Set oGastos = Nothing
'    Set oNegCredito = Nothing
    '**************************************************
    
End Sub


Private Sub cmdCancelar_Click()
    LimpiaDatos
    Call HabilitaDatos(False)
End Sub

Private Sub CmdCargaArch_Click()
Dim sRuta As String
    CdlgFile.nHwd = Me.hwnd
    CdlgFile.Show
    sRuta = CdlgFile.Ruta
    Screen.MousePointer = 11
    Call CargaArchivo(sRuta)
    If Trim(sRuta) <> "" Then
        Call CargaMatrizdeArchivo
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdImprimir_Click()
Dim previo As previo.clsprevio
    Set previo = New previo.clsprevio
    previo.Show sCadImpre, "Pago en Lote", True
    Set previo = Nothing
End Sub

Private Sub CmdPagar_Click()
Dim MatPagos() As String
Dim nContPagos As Integer
Dim i As Integer
Dim nSumTotPago As Double
Dim previo As previo.clsprevio
Dim sImpreBol As String
Dim oCred As COMNCredito.NCOMCredito

Dim oCredPagLote As COMDCredito.DCOMCredito 'JIPR
Dim sMovNro As String 'JIPR
    

    If CmbForPag.ListIndex = -1 Then
        MsgBox "Seleccione la Forma de Pago ", vbInformation, "Aviso"
        Exit Sub
    End If
    'EJVG20140226 ***
    If FEPagoLote.TextMatrix(1, 0) = "" Then
        MsgBox "No se puede continuar ya que no existen registros a procesar", vbInformation, "Aviso"
        Exit Sub
    End If
    'END EJVG *******
    For i = 1 To FEPagoLote.Rows - 1
        Call FEPagoLote.BackColorRow(vbWhite, False)
    Next i
    
    For i = 1 To FEPagoLote.Rows - 1
        If FEPagoLote.TextMatrix(i, 1) = "." And CDbl(FEPagoLote.TextMatrix(i, 5)) > 0 Then
            'If (CDbl(FEPagoLote.TextMatrix(I, 5)) - CDbl(FEPagoLote.TextMatrix(I, 15))) > CDbl(FEPagoLote.TextMatrix(I, 10)) And FEPagoLote.TextMatrix(I, 0) = "1" Then
            'If CDbl(FEPagoLote.TextMatrix(I, 5)) > CDbl(FEPagoLote.TextMatrix(I, 10)) And FEPagoLote.TextMatrix(I, 0) = "1" Then
            If CDbl(FEPagoLote.TextMatrix(i, 5)) > CDbl(FEPagoLote.TextMatrix(i, 10)) Then
                MsgBox "Advertencia en la Fila " & i & ", No se puede Amortizar mas del Monto de la Deuda", vbInformation, "Aviso"
                FEPagoLote.row = i
                FEPagoLote.Col = 5
                FEPagoLote.SetFocus
                Call FEPagoLote.BackColorRow(vbYellow, True)
                Exit Sub
            End If
        End If
    Next i
        
    nContPagos = 0
    For i = 1 To FEPagoLote.Rows - 1
        If FEPagoLote.TextMatrix(i, 1) = "." And CDbl(FEPagoLote.TextMatrix(i, 5)) > 0 Then
            nContPagos = nContPagos + 1
        End If
    Next i
    'ARCV 22-06-2007
    'ReDim MatPagos(nContPagos, 10)
    'ReDim MatPagos(nContPagos, 11)
    '----
    ReDim MatPagos(nContPagos, 12) 'BRGO 20110914
    nContPagos = 0
    For i = 1 To FEPagoLote.Rows - 1
        If FEPagoLote.TextMatrix(i, 1) = "." And CDbl(FEPagoLote.TextMatrix(i, 5)) > 0 Then
            MatPagos(nContPagos, 0) = Trim(FEPagoLote.TextMatrix(i, 2)) 'Codigo de Cuenta
            'MatPagos(nContPagos, 1) = Format(CDbl(FEPagoLote.TextMatrix(i, 5)) - CDbl(FEPagoLote.TextMatrix(i, 15)), "0.00") 'Pago
            MatPagos(nContPagos, 1) = Format(CDbl(FEPagoLote.TextMatrix(i, 5)), "0.00") 'Pago
            MatPagos(nContPagos, 2) = Trim(FEPagoLote.TextMatrix(i, 12)) 'Numero de Transaccion
            MatPagos(nContPagos, 3) = Trim(FEPagoLote.TextMatrix(i, 6)) 'Nombre del Cliente
            MatPagos(nContPagos, 4) = Trim(FEPagoLote.TextMatrix(i, 11)) 'Metodo de Liquidacion
            MatPagos(nContPagos, 5) = Trim(FEPagoLote.TextMatrix(i, 13)) 'Estado del Credito
            MatPagos(nContPagos, 6) = Trim(FEPagoLote.TextMatrix(i, 10)) 'Deuda Total
            MatPagos(nContPagos, 7) = Trim(FEPagoLote.TextMatrix(i, 9)) 'Dias de Atraso
            MatPagos(nContPagos, 8) = Trim(FEPagoLote.TextMatrix(i, 14)) 'Numero de Calendario
            MatPagos(nContPagos, 9) = Trim(FEPagoLote.TextMatrix(i, 15)) 'ITF
            MatPagos(nContPagos, 10) = Trim(FEPagoLote.TextMatrix(i, 21)) 'ARCV 22-06-2007: nPlazo
            MatPagos(nContPagos, 11) = Trim(FEPagoLote.TextMatrix(i, 22)) 'BRGO 14-09-2011: ITF sin Redondeo
            nContPagos = nContPagos + 1
        End If
    Next i
    'EJVG20140226 ***
    If UBound(MatPagos) <= 0 Then
        MsgBox "Ud. debe seleccionar los registros a procesar", vbInformation, "Aviso"
        Exit Sub
    End If
    'END EJVG *******
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCheque Then
        'EJVG20140226 ***
        If Len(Trim(LblXCheque.Caption)) = 0 Then
            MsgBox "Ud. debe seleccionar el cheque", vbInformation, "Aviso"
            If Me.CmbForPag.Visible And CmbForPag.Enabled Then CmbForPag.SetFocus
            Exit Sub
        End If
        If Not ValidaSeleccionCheque Then
            MsgBox "Ud. debe seleccionar el cheque", vbInformation, "Aviso"
            If Me.CmbForPag.Visible And CmbForPag.Enabled Then CmbForPag.SetFocus
            Exit Sub
        End If
        'END EJVG *******
        For i = 0 To nContPagos - 1
            nSumTotPago = nSumTotPago + CDbl(MatPagos(i, 1))
            nSumTotPago = nSumTotPago + CDbl(MatPagos(i, 9)) 'CTI2 20190429 Suma el ITF al total para la validación del cheque
        Next i
        nSumTotPago = CDbl(Format(nSumTotPago, "#0.00"))
        'If nSumTotPago > CDbl(LblXMonto.Caption) Then
        If nSumTotPago > CDbl(oDocRec.fnMonto) Then 'EJVG20140226
            MsgBox "El Monto a Pagar No Debe Exceder el Monto del Cheque", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    If MsgBox("Se va a Realizar el Pago en Lote, Desea Continuar ? ", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    If oDocRec Is Nothing Then Set oDocRec = New UDocRec 'EJVG20140226
    Set oCred = New COMNCredito.NCOMCredito
    Dim psCodCtaError As String
    'sCadImpre = oCred.AmortizarPagoLote(MatPagos, IIf(CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoEfectivo, gColocTipoPagoEfectivo, gColocTipoPagoCheque), gdFecSis, gsNomAge, gsCodAge, gsCodUser, sLpt, Trim(Mid(CmbInstitucion.Text, 1, 30)), Trim(LblXCheque.Caption), ChkImpBol.value, IIf(ChkCobrarMora.value = 1, True, False), , sImpreBol, gsProyectoActual, psCodCtaError)
    sCadImpre = oCred.AmortizarPagoLote(MatPagos, IIf(CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoEfectivo, gColocTipoPagoEfectivo, gColocTipoPagoCheque), gdFecSis, gsNomAge, gsCodAge, gsCodUser, sLpt, Trim(Mid(CmbInstitucion.Text, 1, 30)), oDocRec.fsNroDoc, ChkImpBol.value, IIf(ChkCobrarMora.value = 1, True, False), , sImpreBol, gsProyectoActual, psCodCtaError, oDocRec.fsPersCod, oDocRec.fsIFTpo, oDocRec.fsIFCta, sMovNro) 'EJVG20140226
    If Left(sCadImpre, 5) = "ERROR" Then
        'MsgBox "El Pago no es correcto" & vbCrLf & "Monto a Pagar " & CDbl(FEPagoLote.TextMatrix(CInt(Mid(sCadImpre, 6, InStr(1, sCadImpre, "+") - 6)), 5)) & " Deuda a la Fecha " & Mid(sCadImpre, InStr(6, sCadImpre, "+") + 1, Len(sCadImpre) - InStr(6, sCadImpre, "+")), vbInformation, "Mensaje"
        MsgBox "El Pago no es correcto" & vbCrLf & " en Monto a Pagar " & ", Deuda a la Fecha es " & Format(fgITFCalculaImpuestoNOIncluido(CDbl(Mid(sCadImpre, InStr(6, sCadImpre, "+") + 1, Len(sCadImpre) - InStr(6, sCadImpre, "+")))), "#0.00"), vbInformation, "Mensaje"
        
        Set oCred = Nothing
        
'        For I = 1 To FEPagoLote.Rows - 1
'            If Trim(FEPagoLote.TextMatrix(I, 2)) = psCodCtaError Then
'                FEPagoLote.SelectionMode = flexSelectionFree
'                FEPagoLote.Row = I
'                FEPagoLote.Col = 5
'                Call FEPagoLote.BackColorRow(vbWhite, False)
'            End If
'        Next I
'
'        For I = 1 To FEPagoLote.Rows - 1
'            If Trim(FEPagoLote.TextMatrix(I, 2)) = psCodCtaError Then
'                FEPagoLote.SelectionMode = flexSelectionFree
'                FEPagoLote.Row = I
'                FEPagoLote.Col = 5
'                Call FEPagoLote.BackColorRow(vbYellow, True)
'            End If
'        Next I
        
        Exit Sub
    End If
    
    'JIPR 22/06/2018
    Set oCredPagLote = New COMDCredito.DCOMCredito
    Call oCredPagLote.InsertaPagoLoteDetalle(sMovNro, CStr(Trim(Right(CmbInstitucion.Text, 13))), IIf(CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoEfectivo, gColocTipoPagoEfectivo, gColocTipoPagoCheque))
    'JIPR 22/06/2018
    
    Set oCred = Nothing
    lblMens.Visible = False
    
    MsgBox "Pago en Lote Finalizado", vbInformation, "Aviso"
    Call HabilitaDatos(False)
    Set previo = New previo.clsprevio
    
    If sImpreBol <> "" Then
    '    Previo.Show sImpreBol, "Boleta", True
        previo.PrintSpool sLpt, sImpreBol, True
    End If
    
    'Previo.Show sCadImpre, "Pago en Lote", True
    previo.PrintSpool sLpt, sCadImpre, True
    Set previo = Nothing
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub FEPagoLote_DblClick()
If Trim(FEPagoLote.TextMatrix(1, 0)) <> "" Then
    Call frmCredHistCalendario.PagoCuotas(Trim(FEPagoLote.TextMatrix(FEPagoLote.row, 2)))
End If
End Sub

Private Sub FEPagoLote_OnCellChange(pnRow As Long, pnCol As Long)
Dim nITF As Currency
Dim lnMontoPago As Currency
Dim lnMontoIng As Currency

    lnMontoIng = FEPagoLote.TextMatrix(pnRow, 5)
    'lnMontoPago = fgITFCalculaImpuestoNOIncluido(CDbl(lnMontoIng))
    lnMontoPago = fgITFCalculaImpuestoIncluido(CDbl(lnMontoIng), False)
    If Abs(lnMontoPago - CDbl(FEPagoLote.TextMatrix(pnRow, 10))) <= 0.02 Then
        lnMontoPago = CDbl(FEPagoLote.TextMatrix(pnRow, 10))
    End If
    'lnMontoPago = CDbl(FEPagoLote.TextMatrix(pnRow, 10))
    'nitf = Format(lnMontoPago - CDbl(lnMontoIng), "0.00")
    nITF = Format(CDbl(lnMontoIng) - lnMontoPago, "0.00")
    
    '*** BRGO 20110908 ************************************************
    nRedondeoITF = fgDiferenciaRedondeoITF(nITF)
    If nRedondeoITF > 0 Then
        nITF = nITF - nRedondeoITF
    End If
    '*** END BRGO
    
    FEPagoLote.TextMatrix(pnRow, 15) = Format(nITF, "0.00")
    FEPagoLote.TextMatrix(pnRow, 5) = Format(lnMontoPago, "#0.00")
    FEPagoLote.TextMatrix(pnRow, 22) = Format(nITF + nRedondeoITF, "0.00") 'BRGO 20110908 ITF sin Redondeo
    Call FEPagoLote_OnCellCheck(1, 1)
    nRedondeoITF = 0
End Sub

Private Sub FEPagoLote_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Dim i As Integer
    nTotCred = 0
    nMontoAPagar = 0
    For i = 1 To FEPagoLote.Rows - 1
        If FEPagoLote.TextMatrix(i, 1) = "." Then
            nTotCred = nTotCred + 1
            nMontoAPagar = nMontoAPagar + CDbl(FEPagoLote.TextMatrix(i, 5)) + CDbl(FEPagoLote.TextMatrix(i, 15))
        End If
    Next i
    LblTotCredPag.Caption = nTotCred
    LblTotPag.Caption = Format(nMontoAPagar, "#0.00")
End Sub

Private Sub FEPagoLote_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
'If pnCol = 5 Then
'    If Val(FEPagoLote.TextMatrix(pnRow, pnCol)) > Val(FEPagoLote.TextMatrix(pnRow, 10)) Then
'        MsgBox "Monto de pago no debe ser mayor al total de deuda de credito", vbInformation, "aviso"
'        FEPagoLote.TextMatrix(pnRow, pnCol) = FEPagoLote.TextMatrix(pnRow, 10)
'        Cancel = False
'    End If
'End If

'CTI2 ADD 20190429
 Dim sColumnas() As String
sColumnas = Split(FEPagoLote.ColumnasAEditar, "-")
If sColumnas(pnCol) = "X" Then
    Cancel = False
    MsgBox "Esta celda no es editable", vbInformation, "Aviso"
    SendKeys "{Tab}", True
    Exit Sub
End If
    
If Not IsNumeric(FEPagoLote.TextMatrix(pnRow, 5)) Then
    Cancel = False
    MsgBox "Debes ingresar un valor numérico", vbInformation, "Aviso"
    SendKeys "{Tab}", True
    Exit Sub
End If
'END CTI2

Dim nITF As Currency
Dim lnMontoPago As Currency
Dim nMontoIng As Currency
'Modificado por LMMD
If pnCol = 5 Then
    nMontoIng = FEPagoLote.TextMatrix(pnRow, 5)
    lnMontoPago = fgITFCalculaImpuestoIncluido(CDbl(nMontoIng), False)
    nITF = Format(CDbl(nMontoIng) - lnMontoPago, "0.00")
    
    'If Val(FEPagoLote.TextMatrix(pnRow, pnCol)) > Val(FEPagoLote.TextMatrix(pnRow, 10)) Then
    'If Val(lnMontoPago) > Val(FEPagoLote.TextMatrix(pnRow, 10)) Then
    If Val(lnMontoPago) > CDbl(FEPagoLote.TextMatrix(pnRow, 10)) Then
        MsgBox "Monto de pago no debe ser mayor al total de deuda de credito" & vbCrLf & _
        "El Monto deberia ser " & Format(fgITFCalculaImpuestoNOIncluido(Val(FEPagoLote.TextMatrix(pnRow, 10))), "#0.00"), vbInformation, "aviso"
        FEPagoLote.TextMatrix(pnRow, pnCol) = FEPagoLote.TextMatrix(pnRow, 10)
        Cancel = False
    End If
End If

End Sub

'CTI2 20190429
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End Sub
'END CTI2

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CentraSdi Me
    Call CargaControles
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oDocRec = Nothing
End Sub

'EJVG20140225 ***
Private Sub OptMoneda_Click(Index As Integer)
    Set oDocRec = Nothing
    FormateaFlex FEPagoLote
    LblxBanco.Caption = ""
    sPersCodBco = ""
    LblXCheque.Caption = ""
    LblXMonto.Resalte = eNegro
    LblXMonto.Caption = "0.00"
End Sub
'END EJVG *******
'Private Sub oNegCred_MostrarMensaje()
'lblMens.Visible = True
'lblMens.Caption = oNegCred.pMensaje
'lblMens.Refresh
'End Sub

Private Sub OptSelec_Click(Index As Integer)
Dim i As Integer
    If Index = 0 Then
        For i = 1 To FEPagoLote.Rows - 1
            FEPagoLote.TextMatrix(i, 1) = "1"
        Next i
    Else
        For i = 1 To FEPagoLote.Rows - 1
            FEPagoLote.TextMatrix(i, 1) = " "
        Next i
    End If
    Call FEPagoLote_OnCellCheck(1, 1)
End Sub
'EJVG20140212 ***
Private Function ValidaSeleccionCheque() As Boolean
    ValidaSeleccionCheque = True
    If oDocRec Is Nothing Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
    If Len(Trim(oDocRec.fsNroDoc)) = 0 Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
End Function
'END EJVG *******
