VERSION 5.00
Begin VB.Form FrmCredSimuReprog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simulador de Reprogramacion de Creditos"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1590
      Left            =   30
      TabIndex        =   20
      Top             =   4650
      Width           =   10770
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   195
         TabIndex        =   38
         Top             =   675
         Width           =   1170
      End
      Begin VB.Frame fraIntRep 
         Caption         =   "Interes Reprogramado"
         Height          =   1395
         Left            =   4155
         TabIndex        =   33
         Top             =   105
         Width           =   2400
         Begin VB.OptionButton OptTipoRep 
            Caption         =   "Proratear"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   37
            Top             =   270
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton OptTipoRep 
            Caption         =   "Adicionar a la Ultima Cuota"
            Height          =   360
            Index           =   1
            Left            =   105
            TabIndex        =   36
            Top             =   450
            Width           =   2235
         End
         Begin VB.OptionButton OptTipoRep 
            Caption         =   "Reprogramacion Especial"
            Height          =   360
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   2235
         End
         Begin VB.OptionButton OptTipoRep 
            Caption         =   "Segun CMAC ICA"
            Height          =   360
            Index           =   3
            Left            =   105
            TabIndex        =   34
            Top             =   975
            Width           =   2235
         End
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   435
         Left            =   9270
         TabIndex        =   32
         Top             =   345
         Width           =   1290
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   435
         Left            =   9270
         TabIndex        =   31
         Top             =   825
         Width           =   1305
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Enabled         =   0   'False
         Height          =   360
         Left            =   1410
         TabIndex        =   30
         Top             =   675
         Width           =   1170
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   360
         Left            =   2625
         TabIndex        =   29
         Top             =   675
         Width           =   1170
      End
      Begin VB.CommandButton CmdReprogramar 
         Caption         =   "Renovar Credito"
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
         Left            =   705
         TabIndex        =   28
         Top             =   1080
         Width           =   2520
      End
      Begin VB.Frame Frame1 
         Caption         =   "Renovar A "
         Height          =   960
         Left            =   6600
         TabIndex        =   25
         Top             =   150
         Width           =   2460
         Begin VB.OptionButton OptRenov 
            Caption         =   "Misma Cuota Menor Plazo"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   27
            Top             =   315
            Value           =   -1  'True
            Width           =   2160
         End
         Begin VB.OptionButton OptRenov 
            Caption         =   "Menor Cuota Mismo Plazo"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   26
            Top             =   660
            Width           =   2160
         End
      End
      Begin VB.Frame FraOptRepro 
         Height          =   465
         Left            =   165
         TabIndex        =   22
         Top             =   120
         Width           =   3690
         Begin VB.OptionButton OptRepro 
            Caption         =   "Reprogramar"
            Height          =   195
            Index           =   0
            Left            =   435
            TabIndex        =   24
            Top             =   195
            Value           =   -1  'True
            Width           =   1290
         End
         Begin VB.OptionButton OptRepro 
            Caption         =   "Renovar"
            Height          =   195
            Index           =   1
            Left            =   2010
            TabIndex        =   23
            Top             =   195
            Width           =   1290
         End
      End
      Begin VB.CheckBox chkcalendOrig 
         Caption         =   "De acuerdo a condiciones Originales del Crédito."
         Enabled         =   0   'False
         Height          =   375
         Left            =   6630
         TabIndex        =   21
         Top             =   1140
         Value           =   1  'Checked
         Width           =   2505
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   0
      TabIndex        =   18
      Top             =   1710
      Width           =   10785
      Begin SICMACT.FlexEdit FECalend 
         Height          =   2640
         Left            =   60
         TabIndex        =   19
         Top             =   180
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   4657
         Cols0           =   13
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Fecha-Nro-Monto-Capital-Int. Comp-Int. Mor-Int. Reprog-Int Gracia-Gasto-Saldo-Estado-nCapPag"
         EncabezadosAnchos=   "400-1000-400-1000-1000-1000-1000-1000-1000-1000-1200-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-2-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   65535
         BackColorControl=   65535
         BackColorControl=   65535
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483635
      End
   End
   Begin VB.Frame FraDatos 
      Height          =   1680
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10770
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
         Left            =   6840
         TabIndex        =   2
         Top             =   315
         Width           =   1230
      End
      Begin VB.ListBox LstCtas 
         Height          =   840
         Left            =   8325
         TabIndex        =   1
         Top             =   210
         Width           =   1995
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   480
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   847
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
         CMAC            =   "108"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Interes : "
         Height          =   195
         Left            =   3735
         TabIndex        =   17
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Label LblTasa 
         Alignment       =   1  'Right Justify
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
         Left            =   4815
         TabIndex        =   16
         Top             =   1275
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular :"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   705
         Width           =   525
      End
      Begin VB.Label lblTitular 
         Height          =   195
         Left            =   780
         TabIndex        =   14
         Top             =   705
         Width           =   4260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Analista :"
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   960
         Width           =   645
      End
      Begin VB.Label LblAnalista 
         Height          =   195
         Left            =   930
         TabIndex        =   12
         Top             =   960
         Width           =   3645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prestamo :"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label LblPrestamo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1020
         TabIndex        =   10
         Top             =   1305
         Width           =   900
      End
      Begin VB.Label Saldo 
         AutoSize        =   -1  'True
         Caption         =   "Saldo :"
         Height          =   195
         Left            =   2145
         TabIndex        =   9
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label LblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2730
         TabIndex        =   8
         Top             =   1305
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Capital Reprog.:"
         Height          =   195
         Left            =   7665
         TabIndex        =   7
         Top             =   1335
         Width           =   1590
      End
      Begin VB.Label lblSaldoRep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   9255
         TabIndex        =   6
         Top             =   1305
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ult.Cuota:"
         Height          =   195
         Left            =   4830
         TabIndex        =   5
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label lblfecUltCuota 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6075
         TabIndex        =   4
         Top             =   975
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FrmCredSimuReprog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private nFilaEditar As Integer
Private dFecTemp As Date
Private nMontoApr As Double
Private nTasaInteres As Double
Private MatCalend As Variant
Private nTipoReprogCred As Integer
Dim ldVigencia As Date

Private Sub HabilitaControlesReprog(ByVal pbHabilita As Boolean)
    FECalend.lbEditarFlex = pbHabilita
    cmdEditar.Enabled = Not pbHabilita
    CmdReprogramar.Enabled = Not pbHabilita
    cmdNuevo.Enabled = Not pbHabilita
    cmdSalir.Enabled = Not pbHabilita
    fraIntRep.Enabled = Not pbHabilita
    cmdAceptar.Enabled = pbHabilita
    cmdCancelar.Enabled = pbHabilita
    FraOptRepro.Enabled = Not pbHabilita
    If Me.OptTipoRep(2).value Then
        FECalend.lbEditarFlex = False
    End If
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
Dim oCalend As Dcalendario
Dim oCredito As DCredito
Dim R As ADODB.Recordset
Dim R2 As ADODB.Recordset
Dim nPrdEstado As Integer
Dim lnSaldoNew As Double

    On Error GoTo ErrorCargaDatos
    LimpiaFlex FECalend
    MatCalend = ""
    
    Set oCalend = New Dcalendario
    Set R = oCalend.RecuperaCalendarioPagos(psCtaCod)
    Set oCalend = Nothing
    
    Set oCredito = New DCredito
    nMontoApr = oCredito.SaldoPactadoCredito(psCtaCod)
    
    Set R2 = oCredito.RecuperaDatosComunes(psCtaCod)
    nPrdEstado = R2!nPrdEstado
    nTasaInteres = CDbl(Format(R2!nTasaInteres, "#0.00"))
    lblTitular.Caption = PstaNombre(R2!cTitular)
    lblAnalista.Caption = PstaNombre(R2!cAnalista)
    lblSaldo.Caption = Format(R2!nSaldo, "#0.00")
    LblPrestamo.Caption = Format(R2!nMontoCol, "#0.00")
    lblTasa.Caption = Format(nTasaInteres, "#0.00")
    ldVigencia = Format(R2!dVigencia, "dd/mm/yyyy")
    R2.Close
    Set R2 = Nothing
    Set oCredito = Nothing
    
    If R.BOF Or R.EOF Then
        CargaDatos = False
        R.Close
        Exit Function
    Else
        CargaDatos = True
    End If
    
    If nPrdEstado <> gColocEstVigNorm And nPrdEstado <> gColocEstVigVenc And nPrdEstado <> gColocEstVigMor _
        And nPrdEstado <> gColocEstRefNorm And nPrdEstado <> gColocEstRefMor And nPrdEstado <> gColocEstRefVenc Then
        CargaDatos = False
        R.Close
        Exit Function
    End If
    lnSaldoNew = 0
    Do While Not R.EOF
        FECalend.AdicionaFila
        FECalend.TextMatrix(R.Bookmark, 1) = Format(R!dVenc, "dd/mm/yyyy")
        FECalend.TextMatrix(R.Bookmark, 2) = Trim(Str(R!nCuota))
        FECalend.TextMatrix(R.Bookmark, 3) = Format(IIf(IsNull(R!nCapital), 0, R!nCapital) + _
                                        IIf(IsNull(R!nIntComp), 0, R!nIntComp) + _
                                        IIf(IsNull(R!nIntGracia), 0, R!nIntGracia) + _
                                        IIf(IsNull(R!nIntMor), 0, R!nIntMor) + _
                                        IIf(IsNull(R!nIntReprog), 0, R!nIntReprog) + _
                                        IIf(IsNull(R!nGasto), 0, R!nGasto), "#0.00")
' IIf(IsNull(R!nIntSuspenso), 0, R!nIntSuspenso) +
        FECalend.TextMatrix(R.Bookmark, 4) = Format(IIf(IsNull(R!nCapital), 0, R!nCapital), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 5) = Format(IIf(IsNull(R!nIntComp), 0, R!nIntComp), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 6) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 7) = Format(IIf(IsNull(R!nIntReprog), 0, R!nIntReprog), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 8) = Format(IIf(IsNull(R!nIntGracia), 0, R!nIntGracia), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 9) = Format(IIf(IsNull(R!nGasto), 0, R!nGasto), "#0.00")
        nMontoApr = nMontoApr - IIf(IsNull(R!nCapital), 0, R!nCapital)
        nMontoApr = CDbl(Format(nMontoApr, "#0.00"))
        FECalend.TextMatrix(R.Bookmark, 10) = Format(nMontoApr, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 11) = Trim(Str(R!nColocCalendEstado))
        FECalend.TextMatrix(R.Bookmark, 12) = Format(IIf(IsNull(R!nCapitalPag), 0, R!nCapitalPag), "#0.00")
        
        lnSaldoNew = lnSaldoNew + IIf(IsNull(R!nCapital), 0, R!nCapital) - IIf(IsNull(R!nCapitalPag), 0, R!nCapitalPag)
        
        If R!nColocCalendEstado = gColocCalendEstadoPagado Then
            FECalend.Row = R.Bookmark
            Call FECalend.ForeColorRow(vbRed)
        End If
        If R.RecordCount = R.Bookmark Then
            lblfecUltCuota = Format(R!dVenc, "dd/mm/yyyy")
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    lblSaldoRep = Format(lnSaldoNew, "#0.00")
    Exit Function
ErrorCargaDatos:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Private Sub Carga_Mat_A_Flex(ByVal MatCalend As Variant)
Dim i As Integer
Dim lnSaldoNew As Double
    lnSaldoNew = 0
    For i = 0 To UBound(MatCalend) - 1
        FECalend.AdicionaFila
        FECalend.TextMatrix(i + 1, 1) = MatCalend(i, 0)
        FECalend.TextMatrix(i + 1, 2) = MatCalend(i, 1)
        FECalend.TextMatrix(i + 1, 3) = Format(CDbl(MatCalend(i, 2)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 10)) + CDbl(MatCalend(i, 14)), "#0.00")
        FECalend.TextMatrix(i + 1, 4) = Format(CDbl(MatCalend(i, 2)), "#0.00")
        FECalend.TextMatrix(i + 1, 5) = Format(CDbl(MatCalend(i, 4)), "#0.00")
        FECalend.TextMatrix(i + 1, 6) = Format(CDbl(MatCalend(i, 6)), "#0.00")
        FECalend.TextMatrix(i + 1, 7) = Format(CDbl(MatCalend(i, 8)), "#0.00")
        FECalend.TextMatrix(i + 1, 8) = Format(CDbl(MatCalend(i, 10)), "#0.00")
        FECalend.TextMatrix(i + 1, 9) = Format(CDbl(MatCalend(i, 14)), "#0.00")
        FECalend.TextMatrix(i + 1, 10) = MatCalend(i, 16)
        FECalend.TextMatrix(i + 1, 11) = MatCalend(i, 17)
        FECalend.TextMatrix(i + 1, 12) = Format(CDbl(MatCalend(i, 3)), "#0.00")
        FECalend.Row = i + 1
        
        If CInt(FECalend.TextMatrix(i + 1, 11)) = gColocCalendEstadoPagado Then
            FECalend.Row = i + 1
            Call FECalend.ForeColorRow(vbRed)
        Else
            FECalend.ForeColorRow (vbBlack)
        End If
        lnSaldoNew = lnSaldoNew + CDbl(MatCalend(i, 2)) - CDbl(MatCalend(i, 3))
    Next i
    lblSaldoRep = Format(lnSaldoNew, "#0.00")
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    Dim odCredito As New DCredito
    Dim bCredito As Boolean
    If KeyAscii = 13 Then
        Set odCredito = New DCredito
        bCredito = odCredito.VerificarVencimientoCredito(Me.ActxCta.NroCuenta)
        Set odCredito = Nothing
        If bCredito = True Then
            If Not CargaDatos(ActxCta.NroCuenta) Then
                MsgBox "No se Pudo Cargar el calendario de Pagos", vbInformation, "Aviso"
                fraDatos.Enabled = True
                cmdEditar.Enabled = False
                CmdReprogramar.Enabled = False
                
            Else
                fraDatos.Enabled = False
                
                cmdEditar.Enabled = True
                CmdReprogramar.Enabled = True
            End If
        Else
         MsgBox "El credito tiene dias de atraso no se puede reprogramar", vbInformation, "AVISO"
        End If
    End If
End Sub

Private Sub Impresion()
Dim sCad As String
Dim i As Integer
Dim oPrev As previo.clsprevio

    sCad = Chr$(10)
    sCad = sCad & Space(40) & "Reprogramacion de Credito" & Chr$(10)
    sCad = sCad & Space(38) & String(30, "-") & Chr$(10)
    sCad = sCad & Chr$(10) & Chr$(10)
    sCad = sCad & Space(2) & "Credito : " & Me.ActxCta.NroCuenta
    sCad = sCad & Space(2) & "Titular : " & lblTitular.Caption & Chr$(10)
    sCad = sCad & Space(2) & "Analista : " & LblPrestamo.Caption & Chr$(10)
    sCad = sCad & Space(2) & "Saldo Capital: " & lblSaldo.Caption
    sCad = sCad & Space(2) & "Tasa : " & lblTasa.Caption & Chr$(10) & Chr$(10) & Chr$(10)
    
    sCad = sCad & Space(2) & "Justificacion : " & Chr$(10) & Chr$(10)
    'sCad = sCad & Space(15) & Trim(Me.TxtGlosa.Text) & Chr$(10) & Chr$(10) & Chr$(10)
    sCad = sCad & Space(2) & "Calendario Nuevo : " & Chr$(10)
    
    sCad = sCad & Space(2) & ImpreFormat("Fecha", 10) & ImpreFormat("Nro", 3) & ImpreFormat("Monto", 7)
    sCad = sCad & Space(2) & ImpreFormat("Capital", 10) & ImpreFormat("Interes", 10) & ImpreFormat("Mora", 8)
    sCad = sCad & Space(2) & ImpreFormat("Int.Rep", 10) & ImpreFormat("Int.Gra", 10) & ImpreFormat("Gastos", 10) & ImpreFormat("Saldo", 10) & Chr$(10)
    sCad = sCad & Space(2) & String(110, "-") & Chr$(10)
    
    For i = 1 To FECalend.Rows - 1
        sCad = sCad & Space(2) & ImpreFormat(FECalend.TextMatrix(i, 1), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 2), 3)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 3), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 4), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 5), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 6), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 7), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 8), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 9), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 10), 10) & Chr$(10)
                
    Next i
    
    sCad = sCad & Space(2) & String(110, "-") & Chr$(10)
    
    Set oPrev = New previo.clsprevio
    'oPrev.Show sCad, "Reprogramacion Credito", True
    oPrev.Show sCad, "Reprogramacion Credito", True, , gImpresora
    Set oPrev = Nothing
    
    
End Sub

Private Sub CmdAceptar_Click()
Dim oNCredito As NCredito
Dim i As Integer

    On Error GoTo ErrorCmdAceptar_Click
    If dFecTemp = CDate(Me.FECalend.TextMatrix(FECalend.Row, 1)) And Not OptTipoRep(2).value Then
        MsgBox "La Fecha de Reprogramacion No Puede ser la misma", vbInformation, "Aviso"
        Exit Sub
    End If
    If CCur(lblSaldoRep) <> CCur(lblSaldo) Then
        MsgBox "Saldo de Capital Original no coincide con Saldo Reprogramado. Por favor Verificar", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Se va a Reprogramar el Credito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    'Set oNCredito = New NCredito
    
    'Call oNCredito.ReprogramarCredito(ActxCta.NroCuenta, MatCalend, nTipoReprogCred, , , gdFecSis, , gsCodUser, gsCodAge)
    Call Impresion
   ' Set oNCredito = Nothing
    Call cmdNuevo_Click
    
    Exit Sub

ErrorCmdAceptar_Click:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub cmdBuscar_Click()
Dim oCred As DCredito
Dim R As ADODB.Recordset
Dim sPerscod As String
Dim oPers As UPersona
    On Error GoTo ErrorCmdBuscar_Click
    Set oCred = New DCredito
    Set oPers = frmBuscaPersona.Inicio()
    If oPers Is Nothing Then
        MsgBox "No se Selecciono Ninguna Persona", vbInformation, "Aviso"
        Exit Sub
    Else
        sPerscod = oPers.sPerscod
    End If
    Set oPers = Nothing
    
    Set R = oCred.RecuperaCreditosVigentes(sPerscod, , Array(gColocEstVigMor, gColocEstVigNorm, gColocEstVigVenc, gColocEstRefMor, gColocEstRefNorm, gColocEstRefVenc))
    Set oCred = Nothing
    LstCtas.Clear
    If R.BOF And R.EOF Then
        MsgBox "No Existen Creditos Vigentes", vbInformation, "Aviso"
    End If
    Do While Not R.EOF
        LstCtas.AddItem R!cCtaCod
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Exit Sub

ErrorCmdBuscar_Click:
        MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub cmdCancelar_Click()
    If nTipoReprogCred = 1 Then
        FECalend.Row = nFilaEditar
        Call FECalend.BackColorRow(vbWhite)
        nFilaEditar = -1
    End If
    HabilitaControlesReprog False
    nTipoReprogCred = -1
    MatCalend = ""
    Call ActxCta_KeyPress(13)
End Sub

Private Sub cmdEditar_Click()
    If CInt(FECalend.TextMatrix(FECalend.Row, 11)) = gColocCalendEstadoPagado Then
        MsgBox "No se Puede Reprogramar Cuotas Canceladas", vbInformation, "Aviso"
        Exit Sub
    End If
    nFilaEditar = FECalend.Row
    dFecTemp = CDate(FECalend.TextMatrix(FECalend.Row, 1))
    'Call FECalend.BackColorRow(&HC0FFFF)
    HabilitaControlesReprog True
    nTipoReprogCred = 1
    
    If Me.OptTipoRep(2).value Then
        Call FECalend_OnValidate(1, 1, False)
    End If
End Sub

Private Sub cmdNuevo_Click()
    nTipoReprogCred = -1
    fraDatos.Enabled = True
    LimpiaFlex FECalend
    cmdEditar.Enabled = False
    nFilaEditar = -1
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    cmdSalir.Enabled = True
    HabilitaControlesReprog False
    cmdEditar.Enabled = False
    CmdReprogramar.Enabled = False
    lblAnalista.Caption = ""
    LblPrestamo.Caption = "0.00"
    lblSaldo.Caption = "0.00"
    lblTitular.Caption = ""
    lblTasa.Caption = "0.00"
    'Me.TxtGlosa.Text = ""
    MatCalend = ""
    lblSaldoRep = "0.00"
End Sub


Private Sub CmdReprogramar_Click()
Dim oCredito As NCredito
Dim i As Integer

    nTipoReprogCred = 2
    Set oCredito = New NCredito
    MatCalend = oCredito.ReprogramarCreditoenMemoriaTotal(ActxCta.NroCuenta, gdFecSis, IIf(OptRenov(0).value, True, False))
    Set oCredito = Nothing
    
    HabilitaControlesReprog True
    
    LimpiaFlex FECalend
    For i = 0 To UBound(MatCalend) - 1
        FECalend.AdicionaFila
        FECalend.TextMatrix(i + 1, 1) = MatCalend(i, 0)
        FECalend.TextMatrix(i + 1, 2) = MatCalend(i, 1)
        FECalend.TextMatrix(i + 1, 3) = MatCalend(i, 2)
        FECalend.TextMatrix(i + 1, 4) = MatCalend(i, 3)
        FECalend.TextMatrix(i + 1, 5) = MatCalend(i, 4)
        FECalend.TextMatrix(i + 1, 6) = "0.00"
        FECalend.TextMatrix(i + 1, 7) = "0.00"
        FECalend.TextMatrix(i + 1, 8) = MatCalend(i, 5)
        FECalend.TextMatrix(i + 1, 9) = MatCalend(i, 6)
        FECalend.TextMatrix(i + 1, 10) = MatCalend(i, 7)
        FECalend.Row = i + 1
        Call FECalend.ForeColorRow(vbBlack)
    Next i
        
    FECalend.lbEditarFlex = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub FECalend_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim oNCredito As NCredito
    
    'On Error GoTo ErrorFECalend_OnValidate
    
    If CDbl(FECalend.TextMatrix(nFilaEditar, 3)) < CDbl(FECalend.TextMatrix(nFilaEditar, 5)) + CDbl(FECalend.TextMatrix(nFilaEditar, 6)) + CDbl(FECalend.TextMatrix(nFilaEditar, 7)) + CDbl(FECalend.TextMatrix(nFilaEditar, 8)) + CDbl(FECalend.TextMatrix(nFilaEditar, 9)) Then
        MsgBox "Monto de Cuota debe ser mayor a la suma de los intereses ", vbInformation, "Aviso"
        Exit Sub
    End If
    If OptTipoRep(3).value = False Then
        If CDbl(FECalend.TextMatrix(nFilaEditar, 3)) > CDbl(FECalend.TextMatrix(nFilaEditar, 4)) + CDbl(FECalend.TextMatrix(nFilaEditar, 5)) + CDbl(FECalend.TextMatrix(nFilaEditar, 6)) + CDbl(FECalend.TextMatrix(nFilaEditar, 7)) + CDbl(FECalend.TextMatrix(nFilaEditar, 8)) + CDbl(FECalend.TextMatrix(nFilaEditar, 9)) Then
            MsgBox "Monto de Cuota debe ser menor a la cuota anterior", vbInformation, "Aviso"
            FECalend.TextMatrix(nFilaEditar, 3) = Format(CDbl(FECalend.TextMatrix(nFilaEditar, 4)) + CDbl(FECalend.TextMatrix(nFilaEditar, 5)) + CDbl(FECalend.TextMatrix(nFilaEditar, 6)) + CDbl(FECalend.TextMatrix(nFilaEditar, 7)) + CDbl(FECalend.TextMatrix(nFilaEditar, 8)) + CDbl(FECalend.TextMatrix(nFilaEditar, 9)), "#0.00")
            Exit Sub
        End If
    End If
    
    If OptTipoRep(3).value = False Then
        If CDate(FECalend.TextMatrix(nFilaEditar, 1)) < dFecTemp And Not Me.OptTipoRep(2).value Then
            MsgBox "La Fecha de Reprogramacion debe ser Mayor a la Anterior", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        End If
    End If
    
    If CDate(FECalend.TextMatrix(nFilaEditar, 1)) > CDate(lblfecUltCuota) Then
        MsgBox "La Fecha de Reprogramacion no puede ser Mayor al vencimiento de la ultima cuota", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
    If CDate(FECalend.TextMatrix(nFilaEditar, 1)) < gdFecSis Then
        MsgBox "Fecha no puede ser menor a la fecha actual del sistema", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
    
    Set oNCredito = New NCredito
    If OptTipoRep(0).value Then
        MatCalend = oNCredito.ReprogramarCreditoenMemoria(ActxCta.NroCuenta, nTasaInteres, dFecTemp, CDate(FECalend.TextMatrix(nFilaEditar, 1)), nFilaEditar - 1, 1, True, MatCalend)
    Else
        If OptTipoRep(1).value Then
            MatCalend = oNCredito.ReprogramarCreditoenMemoria(ActxCta.NroCuenta, nTasaInteres, dFecTemp, CDate(FECalend.TextMatrix(nFilaEditar, 1)), nFilaEditar - 1, 2, False, MatCalend)
        Else
            If OptTipoRep(2).value Then
                MatCalend = oNCredito.ReprogramarCreditoenMemoria(ActxCta.NroCuenta, nTasaInteres, dFecTemp, CDate(FECalend.TextMatrix(nFilaEditar, 1)), nFilaEditar - 1, 3, False, MatCalend)
            Else
                If Val(FECalend.TextMatrix(nFilaEditar, 12)) > 0 Then
                    MsgBox "Cuota posee pagado capital no podrá continuar", vbInformation, "Aviso"
                    Cancel = False
                    Exit Sub
                End If
                
                Dim lnMonto As Double
                If (CDbl(FECalend.TextMatrix(nFilaEditar, 3)) <> CDbl(FECalend.TextMatrix(nFilaEditar, 4)) + CDbl(FECalend.TextMatrix(nFilaEditar, 5)) + CDbl(FECalend.TextMatrix(nFilaEditar, 6)) + CDbl(FECalend.TextMatrix(nFilaEditar, 7)) + CDbl(FECalend.TextMatrix(nFilaEditar, 8)) + CDbl(FECalend.TextMatrix(nFilaEditar, 9))) And (nFilaEditar <> FECalend.Rows - 1) Then
                    'MatCalend = oNCredito.ReprogramarCreditoMonto(MatCalend, nFilaEditar - 1, CDbl(FECalend.TextMatrix(FECalend.Row, 3)))
                    lnMonto = CDbl(FECalend.TextMatrix(FECalend.Row, 3))
                Else
                    lnMonto = 0
                End If
                If IsDate(FECalend.TextMatrix(nFilaEditar - 1, 1)) = True Then
                    dFecTemp = FECalend.TextMatrix(nFilaEditar - 1, 1)
                End If
                MatCalend = oNCredito.ReprogramarCreditoenMemoria(ActxCta.NroCuenta, nTasaInteres, dFecTemp, CDate(FECalend.TextMatrix(nFilaEditar, 1)), nFilaEditar - 1, 4, False, MatCalend, lnMonto, ldVigencia, chkcalendOrig.value, CDate(lblfecUltCuota))
                
                'If (CDbl(FECalend.TextMatrix(nFilaEditar, 3)) - (CDbl(MatCalend(nFilaEditar, 4)) + CDbl(MatCalend(nFilaEditar - 1, 6)) + CDbl(MatCalend(nFilaEditar - 1, 8)) + CDbl(MatCalend(nFilaEditar - 1, 10)) + CDbl(MatCalend(nFilaEditar - 1, 12)) + CDbl(MatCalend(nFilaEditar - 1, 14)))) < CDbl(MatCalend(nFilaEditar - 1, 3)) Then
                '        MsgBox "El monto de la cuota debe ser " & Format((CDbl(MatCalend(nFilaEditar - 1, 4)) + CDbl(MatCalend(nFilaEditar - 1, 6)) + CDbl(MatCalend(nFilaEditar - 1, 8)) + CDbl(MatCalend(nFilaEditar - 1, 10)) + CDbl(MatCalend(nFilaEditar - 1, 12)) + CDbl(MatCalend(nFilaEditar - 1, 14)) + CDbl(MatCalend(nFilaEditar - 1, 3))), "#0.00") & " por tener  capital pagado ", vbInformation, "Aviso"
                '        Exit Sub
                'End If
            End If
        End If
    End If
    
    If (CDbl(FECalend.TextMatrix(nFilaEditar, 3)) - (CDbl(MatCalend(nFilaEditar - 1, 4)) + CDbl(MatCalend(nFilaEditar - 1, 6)) + CDbl(MatCalend(nFilaEditar - 1, 8)) + CDbl(MatCalend(nFilaEditar - 1, 10)) + CDbl(MatCalend(nFilaEditar - 1, 12)) + CDbl(MatCalend(nFilaEditar - 1, 14)))) < CDbl(MatCalend(nFilaEditar - 1, 3)) Then
        MsgBox "El monto de la cuota debe ser " & Format((CDbl(MatCalend(nFilaEditar - 1, 4)) + CDbl(MatCalend(nFilaEditar - 1, 6)) + CDbl(MatCalend(nFilaEditar - 1, 8)) + CDbl(MatCalend(nFilaEditar - 1, 10)) + CDbl(MatCalend(nFilaEditar - 1, 12)) + CDbl(MatCalend(nFilaEditar - 1, 14)) + CDbl(MatCalend(nFilaEditar - 1, 3))), "#0.00") & " por tener  capital pagado ", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If (nFilaEditar = FECalend.Rows - 1) And FECalend.Col = 3 Then
        MsgBox "No se puede Actualizar el monto de la Ultima Cuota", vbInformation, "Aviso"
        FECalend.TextMatrix(nFilaEditar, 3) = Format(CDbl(FECalend.TextMatrix(nFilaEditar, 4)) + CDbl(FECalend.TextMatrix(nFilaEditar, 5)) + CDbl(FECalend.TextMatrix(nFilaEditar, 6)) + CDbl(FECalend.TextMatrix(nFilaEditar, 7)) + CDbl(FECalend.TextMatrix(nFilaEditar, 8)) + CDbl(FECalend.TextMatrix(nFilaEditar, 9)), "#0.00")
        Exit Sub
    End If
    
    Set oNCredito = Nothing
    LimpiaFlex FECalend
    Call Carga_Mat_A_Flex(MatCalend)
    
    If nFilaEditar > 0 Then
        FECalend.Row = nFilaEditar
        FECalend.BackColorRow vbYellow
        FECalend.SetFocus
    End If
    
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
    cmdNuevo.Enabled = True
    
    
    Exit Sub

ErrorFECalend_OnValidate:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub FECalend_RowColChange()
    'If nFilaEditar <> -1 Then
    '    FECalend.Row = nFilaEditar
    '    FECalend.Col = 1
    'End If
    If nFilaEditar <> -1 Then
        nFilaEditar = FECalend.Row
        'FECalend.Col = 1
    End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then 'F12
'And ActxCta.Enabled = True
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColComercEmp, bRetSinTarjeta)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CentraForm Me
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    nFilaEditar = -1
End Sub


Private Sub LstCtas_Click()
    If LstCtas.ListCount > 0 And LstCtas.ListIndex <> -1 Then
        ActxCta.NroCuenta = LstCtas.Text
    End If
End Sub

Private Sub OptRepro_Click(Index As Integer)
    If Index = 0 Then
        CmdReprogramar.Enabled = False
        OptRenov(0).Enabled = False
        OptRenov(1).Enabled = False
        cmdEditar.Enabled = True
        OptTipoRep(0).Enabled = True
        OptTipoRep(1).Enabled = True
        OptTipoRep(2).Enabled = True
    Else
        CmdReprogramar.Enabled = True
        OptRenov(0).Enabled = True
        OptRenov(1).Enabled = True
        cmdEditar.Enabled = False
        OptTipoRep(0).Enabled = False
        OptTipoRep(1).Enabled = False
        OptTipoRep(2).Enabled = False
    End If
End Sub

Private Sub OptTipoRep_Click(Index As Integer)
chkcalendOrig.Enabled = False
If OptTipoRep(3).value = True Then
    chkcalendOrig.Enabled = True
End If
End Sub


