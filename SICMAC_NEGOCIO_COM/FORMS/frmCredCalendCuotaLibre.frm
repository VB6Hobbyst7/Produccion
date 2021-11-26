VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCredCalendCuotaLibre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario de Cuota libre"
   ClientHeight    =   6465
   ClientLeft      =   2025
   ClientTop       =   1515
   ClientWidth     =   7635
   Icon            =   "frmCredCalendCuotaLibre.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7635
   Begin SICMACT.FlexEdit FECalend 
      Height          =   2865
      Left            =   135
      TabIndex        =   8
      Top             =   2895
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   5054
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "-Fecha de Venc.-Cuota-Capital-Interes-Saldo Capital"
      EncabezadosAnchos=   "400-1500-1200-1200-1200-1500"
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
      ColumnasAEditar =   "X-1-2-X-X-X"
      ListaControles  =   "0-2-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-R-R-R-R"
      FormatosEdit    =   "0-0-2-2-2-2"
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame Frame1 
      Caption         =   "Condiciones"
      Height          =   2715
      Left            =   525
      TabIndex        =   11
      Top             =   60
      Width           =   6480
      Begin VB.Frame fraCondicion 
         Height          =   1350
         Left            =   120
         TabIndex        =   16
         Top             =   705
         Width           =   3000
         Begin VB.TextBox txtMonto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1635
            MaxLength       =   9
            TabIndex        =   2
            Top             =   225
            Width           =   1245
         End
         Begin VB.TextBox txtinteres 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1635
            MaxLength       =   7
            TabIndex        =   3
            Top             =   525
            Width           =   1245
         End
         Begin MSComCtl2.DTPicker DTPFecdes 
            Height          =   300
            Left            =   1620
            TabIndex        =   4
            Top             =   960
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
            _Version        =   393216
            Format          =   71237633
            CurrentDate     =   36586
         End
         Begin VB.Label Label8 
            Caption         =   "Fecha Desembolso"
            Height          =   435
            Left            =   165
            TabIndex        =   19
            Top             =   885
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Interes (Mensual)"
            Height          =   195
            Left            =   150
            TabIndex        =   18
            Top             =   540
            Width           =   1215
         End
         Begin VB.Label lblmonto 
            AutoSize        =   -1  'True
            Caption         =   "&Monto Total"
            Height          =   195
            Left            =   165
            TabIndex        =   17
            Top             =   240
            Width           =   855
         End
         Begin VB.Line Line1 
            X1              =   60
            X2              =   2955
            Y1              =   855
            Y2              =   855
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            BorderStyle     =   6  'Inside Solid
            X1              =   60
            X2              =   2895
            Y1              =   870
            Y2              =   870
         End
      End
      Begin VB.CommandButton cmdRemCuota 
         Caption         =   "&Remover Ultima Cuota"
         Height          =   390
         Left            =   3330
         TabIndex        =   7
         Top             =   1695
         Width           =   2940
      End
      Begin VB.Frame fraCuota 
         Enabled         =   0   'False
         Height          =   510
         Left            =   120
         TabIndex        =   13
         Top             =   2115
         Width           =   6150
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Deuda Total :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3375
            TabIndex        =   15
            Top             =   225
            Width           =   1185
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
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
            Height          =   195
            Left            =   4845
            TabIndex        =   14
            Top             =   225
            Width           =   1125
         End
      End
      Begin VB.Frame Frame2 
         Height          =   540
         Left            =   135
         TabIndex        =   12
         Top             =   165
         Width           =   6135
         Begin VB.OptionButton OptDesemb 
            Caption         =   "Desembolso Total"
            Height          =   225
            Index           =   0
            Left            =   495
            TabIndex        =   0
            Top             =   210
            Value           =   -1  'True
            Width           =   2250
         End
         Begin VB.OptionButton OptDesemb 
            Caption         =   "Desembolso Parcial"
            Height          =   225
            Index           =   1
            Left            =   3360
            TabIndex        =   1
            Top             =   210
            Width           =   2250
         End
      End
      Begin VB.CommandButton CmdAddDesemb 
         Caption         =   "Agregar &Desembolso"
         Enabled         =   0   'False
         Height          =   390
         Left            =   3330
         TabIndex        =   5
         Top             =   810
         Width           =   2940
      End
      Begin VB.CommandButton cmdAddCuota 
         Caption         =   "Agregar &Cuota"
         Height          =   390
         Left            =   3330
         TabIndex        =   6
         Top             =   1245
         Width           =   2940
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   4830
         TabIndex        =   21
         Top             =   1245
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton cmdaceptar 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   3345
         TabIndex        =   20
         Top             =   1245
         Visible         =   0   'False
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   5850
      TabIndex        =   10
      Top             =   5880
      Width           =   1665
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   405
      Left            =   4095
      TabIndex        =   9
      Top             =   5880
      Width           =   1665
   End
End
Attribute VB_Name = "frmCredCalendCuotaLibre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredCalendCuotaLibre
'***     Descripcion:      Opcion que Permite Generar un Calendario de Cuota Libre
'                          A Libre Demanda del Cliente
'***     Creado por:        NSSE
'***     Maquina:           07SISTDES03
'***     Fecha-Tiempo:         09/07/2001 05:10:48 PM
'***     Ultima Modificacion: Creacion de la Opcion
'*****************************************************************************************
Option Explicit
Private MatDesPar() As String
Private MatCalendario() As String
Private nUltimaFila As Integer
Private vsCadImpresion As String
Private lsNegritaOn As String
Private lsNegritaOff  As String
Private lsSaltoLin As String
Private lsTab As String
Private EspHorxPag As Integer
Private EspVerxPag As Integer
Private nPuntPag As Integer
Private sCadCab As String
Dim csNomCMAC As String
Dim csNomAgencia As String
Dim csCodUser As String
Dim csFechaSis As String


Public Function CalendarioLibre(ByVal pnSimulacion As Boolean, ByVal pdFecDesemb As Date, ByVal MatCalend As Variant, Optional ByVal pnPrestamo As Double = 0#, _
        Optional ByVal pnTipoDes As Integer = 0, Optional pnTasa As Double = 0#) As Variant
Dim i As Integer

        Frame2.Enabled = pnSimulacion
        txtMonto.Enabled = pnSimulacion
        txtinteres.Enabled = pnSimulacion
        DTPFecdes.Enabled = pnSimulacion

        OptDesemb(pnTipoDes).value = True
        If pnSimulacion Then
            Call Optdesemb_Click(pnTipoDes)
        Else
            HabilitaDesembolsoTotal False
            CmdAddDesemb.Enabled = False
        End If
        txtMonto.Text = Format(pnPrestamo, "#0.00")
        txtinteres.Text = Format(pnTasa, "#0.00")
        DTPFecdes.value = Format(pdFecDesemb, "dd/mm/yyyy")

        Call LimpiaFlex(FECalend)
        If UBound(MatCalend) > 0 Then
            If UBound(MatCalend) > 1 Then
                FECalend.Rows = UBound(MatCalend) + 1
            End If
            For i = 0 To UBound(MatCalend) - 1
                FECalend.TextMatrix(i + 1, 0) = MatCalend(i, 0)
                FECalend.TextMatrix(i + 1, 1) = MatCalend(i, 1)
                FECalend.TextMatrix(i + 1, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)), "#0.00")
                FECalend.TextMatrix(i + 1, 3) = MatCalend(i, 3)
                FECalend.TextMatrix(i + 1, 4) = MatCalend(i, 4)
                FECalend.TextMatrix(i + 1, 5) = MatCalend(i, 5)
            Next i
        End If
        Me.Show 1

        CalendarioLibre = MatCalendario
End Function

Private Sub HabilitaDesembolsoTotal(ByVal pbHabilita As Boolean)
    txtMonto.Enabled = pbHabilita
    CmdAddDesemb.Enabled = Not pbHabilita
    DTPFecdes.Enabled = pbHabilita
End Sub
Private Function MontoPrestamoParcial() As Double
Dim i As Integer
    MontoPrestamoParcial = 0
    If UBound(MatDesPar) > 0 Then
        If Trim(MatDesPar(0, 1)) = "" Then
            Exit Function
        End If
    End If
    For i = 0 To UBound(MatDesPar) - 1
        MontoPrestamoParcial = MontoPrestamoParcial + CDbl(MatDesPar(i, 1))
    Next i

End Function
Private Sub HabilitaNuevaCuota(ByVal pbHabilita As Boolean)

    cmdAddCuota.Visible = Not pbHabilita
    cmdaceptar.Visible = pbHabilita
    cmdcancelar.Visible = pbHabilita
    fraCondicion.Enabled = False
    CmdAddDesemb.Enabled = False
    Frame2.Enabled = False
    cmdRemCuota.Enabled = Not pbHabilita
    cmdImprimir.Enabled = Not pbHabilita
    cmdsalir.Enabled = Not pbHabilita
    If pbHabilita Then
        CmdAddDesemb.Enabled = False
    Else
        If OptDesemb(1).value Then
            CmdAddDesemb.Enabled = True
        Else
            CmdAddDesemb.Enabled = False
        End If
    End If
End Sub

Private Sub CmdAceptar_Click()

    If ValidaFecha(FECalend.TextMatrix(nUltimaFila, 1)) <> "" Then
        MsgBox ValidaFecha(FECalend.TextMatrix(nUltimaFila, 1)), vbInformation, "Aviso"
        Exit Sub
    End If

    If Trim(FECalend.TextMatrix(nUltimaFila, 2)) = "" Then
        MsgBox "Ingrese la Cuota", vbInformation, "Aviso"
        Exit Sub
    End If


    If CDbl(FECalend.TextMatrix(nUltimaFila, 2)) < CDbl(FECalend.TextMatrix(nUltimaFila, 4)) Then
        MsgBox "La Cuota No Puede ser Menor que el Interes ", vbInformation, "Aviso"
        Exit Sub
    End If

    Call FECalend.BackColorRow(vbWhite)
    FECalend.lbEditarFlex = False
    HabilitaNuevaCuota False
    lblTotal.Caption = Format(CDbl(FECalend.TextMatrix(nUltimaFila, 5)), "#0.00")
    nUltimaFila = -1
End Sub

Private Sub cmdAddCuota_Click()
    If CDbl(txtMonto.Text) = 0 Then
        MsgBox "Ingrese Monto del Prestamo", vbInformation, "Aviso"
        If txtMonto.Enabled Then
            txtMonto.SetFocus
        End If
        Exit Sub
    End If
    If CDbl(txtinteres.Text) = 0 Then
        MsgBox "Ingrese la Tasa de Interes", vbInformation, "Aviso"
        If txtinteres.Enabled = True Then
            txtinteres.SetFocus
        End If
        Exit Sub
    End If
    If FECalend.TextMatrix(1, 1) <> "" Then
       If CDbl(FECalend.TextMatrix(FECalend.Rows - 1, 5)) = 0# Then
            MsgBox "El Saldo de Capital es Cero, No puede Ingresar Mas Cuotas ", vbInformation, "Aviso"
            Exit Sub
       End If
    End If
    FECalend.lbEditarFlex = True
    FECalend.AdicionaFila
    nUltimaFila = FECalend.Rows - 1
    FECalend.row = nUltimaFila
    Call FECalend.BackColorRow(vbYellow)
    HabilitaNuevaCuota True
End Sub

Private Sub CmdAddDesemb_Click()
    MatDesPar = frmCredDesembParcial.Inicio(gdFecSis) ', MatDesPar)
    txtMonto.Text = Format(MontoPrestamoParcial, "#0.00")
    If FECalend.TextMatrix(1, 1) <> "" Then
        If UBound(MatDesPar) > 0 Then
            DTPFecdes.value = CDate(MatDesPar(UBound(MatDesPar) - 1, 0))
        End If
    End If
End Sub

Private Sub CmdCancelar_Click()
    Call FECalend.BackColorRow(vbWhite)
    Call FECalend.EliminaFila(nUltimaFila)
    FECalend.lbEditarFlex = False
    HabilitaNuevaCuota False
    nUltimaFila = -1
    If FECalend.TextMatrix(1, 1) = "" Then
        Frame2.Enabled = True
        fraCondicion.Enabled = True
        If OptDesemb(1).value Then
            CmdAddDesemb.Enabled = True
        Else
            CmdAddDesemb.Enabled = False
        End If
    End If
    If Trim(FECalend.TextMatrix(1, 1)) = "" Then
        lblTotal.Caption = Format(CDbl(txtMonto.Text), "#0.00")
    Else
        lblTotal.Caption = Format(CDbl(FECalend.TextMatrix(FECalend.Rows - 1, 5)), "#0.00")
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim sCadImp As String
    Dim oImpresion As clsprevio
    Dim i As Integer
    Dim nCapitalTotal As Double
    Dim nInteresTotal As Double
    Dim nCuotaTotal As Double

    Inicio
    Call ImprimeCabeceraDocumento(sCadImp, gsNomAge, Format(gdFecSis, "dd/mm/yyyy hh:mm:ss"), gsCodUser, "PLAN DE PAGOS", "CMAC - ICA", , False)

    sCadImp = sCadImp & String(150, "---") & Chr(10)
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & String(150, "---") & Chr(10)

    sCadImp = sCadImp & ImpreFormat("NRO DE CUOTA", 15) & ImpreFormat("FECHA", 15) & ImpreFormat("CAPITAL", 8) & ImpreFormat("INTERES", 8) & ImpreFormat("CUOTA", 15) & ImpreFormat("SALDO CAPITAL", 15) & Chr(10)
    sCadImp = sCadImp & String(150, "---") & Chr(10)
   ' sCadImp = sCadImp & Chr(10)

    For i = 1 To FECalend.Rows - 1
        sCadImp = sCadImp & ImpreFormat(i, 17, 0) & ImpreFormat(FECalend.TextMatrix(i, 1), 15) & ImpreFormat(Format(FECalend.TextMatrix(i, 3), "#0.00"), 8, 2) & ImpreFormat(Format(FECalend.TextMatrix(i, 4), "#0.00"), 8, 2) & ImpreFormat(Format(FECalend.TextMatrix(i, 2), "#0.00"), 16, 2) & ImpreFormat(Format(FECalend.TextMatrix(i, 5), "#0.00"), 8, 2) & Chr(10)
        nCapitalTotal = nCapitalTotal + CDbl(FECalend.TextMatrix(i, 3))
        nInteresTotal = nInteresTotal + CDbl(FECalend.TextMatrix(i, 4))
        nCuotaTotal = nCuotaTotal + CDbl(FECalend.TextMatrix(i, 2))
    Next i

    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & "Resumen de Calendario de Cuota Libre" & Chr(10)
    sCadImp = sCadImp & String(150, "---") & Chr(10)

    sCadImp = sCadImp & ImpreFormat("Capital: ", 8) & ImpreFormat(Format(nCapitalTotal, "#0.00"), 10, 2) & Chr(10)
    sCadImp = sCadImp & ImpreFormat("Interes: ", 8) & ImpreFormat(Format(nInteresTotal, "#0.00"), 10, 2) & Chr(10)
    sCadImp = sCadImp & ImpreFormat("Cuota Total:", 8) & ImpreFormat(Format(nCuotaTotal, "#0.00"), 10, 2) & Chr(10)

    Set oImpresion = New clsprevio
    oImpresion.Show sCadImp, "PLAN DE PAGOS DE CUOTA LIBRE", , , gImpresora

End Sub
Private Function lnSaltoLinDoc() As String
    nPuntPag = nPuntPag + 1
    If nPuntPag > EspVerxPag Then
        nPuntPag = 0
        lnSaltoLinDoc = Chr$(12) & sCadCab
    Else
        lnSaltoLinDoc = Chr$(10)
    End If

End Function
Private Sub ImprimeCabeceraDocumento(ByRef psCadImp As String, ByVal psNomAge As String, _
    ByVal psFechaHora As String, ByVal psCodUsu, ByVal psTitulo As String, ByVal psNomCmac As String, _
    Optional psTab As String = "", Optional pbCondensado As Boolean = True, Optional psCodRepo As String = "")

    nPuntPag = 0
    psCadImp = psCadImp & lnSaltoLinDoc
    If Len(psCodRepo) > 0 Then
        psTitulo = psCodRepo & " - " & psTitulo
    End If
    If pbCondensado Then
        psCadImp = psCadImp & psTab & psNomCmac & Space(70) & "Fecha :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 45, 0) & Space(51) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
    Else
        psCadImp = psCadImp & psTab & psNomCmac & Space(45) & "Fecha :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 40, 0) & Space(36) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space((120 - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
    End If
End Sub

Private Sub cmdRemCuota_Click()
    If MsgBox("Se va Ha Eliminar la Cuota " & Trim(FECalend.TextMatrix(FECalend.row, 0)) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        FECalend.EliminaFila FECalend.row
    End If
    If FECalend.TextMatrix(1, 1) = "" Then
        Frame2.Enabled = True
        fraCondicion.Enabled = True
        If OptDesemb(1).value Then
            CmdAddDesemb.Enabled = True
        Else
            CmdAddDesemb.Enabled = False
        End If
    End If
    If Trim(FECalend.TextMatrix(1, 1)) = "" Then
        lblTotal.Caption = Format(CDbl(txtMonto.Text), "#0.00")
    Else
        lblTotal.Caption = Format(CDbl(FECalend.TextMatrix(FECalend.Rows - 1, 5)), "#0.00")
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub FECalend_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim oNCredito As COMNCredito.NCOMCredito
    On Error GoTo ErrorFECalend_OnValidate
    If nUltimaFila <> -1 Then
        Set oNCredito = New COMNCredito.NCOMCredito
        If pnCol = 1 Then
            If Trim(FECalend.TextMatrix(pnRow, 1)) = "" Then
                FECalend.TextMatrix(pnRow, 4) = "0.00"
                FECalend.TextMatrix(pnRow, 2) = "0.00"
                FECalend.TextMatrix(pnRow, 3) = "0.00"
                Exit Sub
            End If
            If pnRow = 1 Then
                If CDate(FECalend.TextMatrix(pnRow, 1)) <= CDate(DTPFecdes.value) Then
                    MsgBox "La Fecha de Pago No Puede ser Menor que la Fecha del Desembolso", vbInformation, "Mensaje"
                    Set oNCredito = Nothing
                    Cancel = False
                Else
                    FECalend.TextMatrix(pnRow, 4) = Format(oNCredito.TasaIntPerDias(CDbl(txtinteres.Text), DateDiff("d", CDate(DTPFecdes.value), CDate(FECalend.TextMatrix(pnRow, 1)))) * CDbl(txtMonto.Text), "#0.00")
                    lblTotal.Caption = Format(CDbl(FECalend.TextMatrix(pnRow, 4)) + CDbl(txtMonto.Text), "#0.00")
                End If
            Else
                If CDate(FECalend.TextMatrix(pnRow, 1)) <= CDate(FECalend.TextMatrix(pnRow - 1, 1)) Then
                    MsgBox "La Fecha de Pago No Puede ser Menor que la Anterior", vbInformation, "Mensaje"
                    Set oNCredito = Nothing
                    Cancel = False
                Else
                    FECalend.TextMatrix(pnRow, 4) = Format(oNCredito.TasaIntPerDias(CDbl(txtinteres.Text), DateDiff("d", CDate(FECalend.TextMatrix(pnRow - 1, 1)), CDate(FECalend.TextMatrix(pnRow, 1)))) * CDbl(FECalend.TextMatrix(pnRow - 1, 5)), "#0.00")
                    lblTotal.Caption = Format(CDbl(FECalend.TextMatrix(pnRow, 4)) + CDbl(FECalend.TextMatrix(pnRow - 1, 5)), "#0.00")
                End If
            End If
            FECalend.TextMatrix(pnRow, 2) = "0.00"
        End If

        'Validando Cuota
        If pnCol = 2 Then
            If Trim(FECalend.TextMatrix(pnRow, 2)) = "" Then
                Exit Sub
            End If
            If CDbl(FECalend.TextMatrix(pnRow, 2)) < CDbl(FECalend.TextMatrix(pnRow, 4)) Then
                MsgBox "La Cuota No Puede ser Menor que el Interes ", vbInformation, "Aviso"
                Cancel = False
            End If
            If pnRow = 1 Then
                If CDbl(Format(CDbl(FECalend.TextMatrix(pnRow, 2)) - CDbl(FECalend.TextMatrix(pnRow, 4)), "#0.00")) > CDbl(txtMonto.Text) Then
                    MsgBox "El Monto de la Cuota debe ser como Maximo " & Format(CDbl(txtMonto.Text) + CDbl(FECalend.TextMatrix(pnRow, 4)), "#0.00"), vbInformation, "Aviso"
                    FECalend.TextMatrix(pnRow, 2) = Format(CDbl(txtMonto.Text) + CDbl(FECalend.TextMatrix(pnRow, 4)), "#0.00")
                    Cancel = False
                Else
                    FECalend.TextMatrix(pnRow, 3) = Format(CDbl(FECalend.TextMatrix(pnRow, 2)) - CDbl(FECalend.TextMatrix(pnRow, 4)), "#0.00")
                    If pnRow = 1 Then
                        FECalend.TextMatrix(pnRow, 5) = Format(CDbl(txtMonto.Text) - CDbl(FECalend.TextMatrix(pnRow, 3)), "#0.00")
                    Else
                        FECalend.TextMatrix(pnRow, 5) = Format(CDbl(FECalend.TextMatrix(pnRow - 1, 5)) - CDbl(FECalend.TextMatrix(pnRow, 3)), "#0.00")
                    End If
                End If
            Else
                If CDbl(Format(CDbl(FECalend.TextMatrix(pnRow, 2)) - CDbl(FECalend.TextMatrix(pnRow, 4)), "#0.00")) > CDbl(FECalend.TextMatrix(pnRow - 1, 5)) Then
                    MsgBox "El Monto de la Cuota debe ser como Maximo " & Format(CDbl(FECalend.TextMatrix(pnRow - 1, 5)) + CDbl(FECalend.TextMatrix(pnRow, 4)), "#0.00"), vbInformation, "Aviso"
                    FECalend.TextMatrix(pnRow, 2) = Format(CDbl(FECalend.TextMatrix(pnRow - 1, 5)) + CDbl(FECalend.TextMatrix(pnRow, 4)), "#0.00")
                    Cancel = False
                Else
                    FECalend.TextMatrix(pnRow, 3) = Format(CDbl(FECalend.TextMatrix(pnRow, 2)) - CDbl(FECalend.TextMatrix(pnRow, 4)), "#0.00")
                    If pnRow = 1 Then
                        FECalend.TextMatrix(pnRow, 5) = Format(CDbl(txtMonto.Text) - CDbl(FECalend.TextMatrix(pnRow, 3)), "#0.00")
                    Else
                        FECalend.TextMatrix(pnRow, 5) = Format(CDbl(FECalend.TextMatrix(pnRow - 1, 5)) - CDbl(FECalend.TextMatrix(pnRow, 3)), "#0.00")
                    End If
                End If
            End If
        End If
        Set oNCredito = Nothing
    End If

    Exit Sub

ErrorFECalend_OnValidate:
        MsgBox Err.Description, vbCritical, "Aviso"


End Sub

Private Sub FECalend_RowColChange()
    If nUltimaFila <> -1 Then
        FECalend.row = nUltimaFila
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CentraForm Me
    nUltimaFila = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
    If FECalend.Rows = 0 Then
        Exit Sub
    End If
    If FECalend.TextMatrix(1, 1) <> "" Then
        If CDbl(FECalend.TextMatrix(FECalend.Rows - 1, 5)) <> 0# Then
            MsgBox "No puede Salir de la Opcion Hasta que halla Distribuido todo el Capital", vbInformation, "Aviso"
            Cancel = 1
            Exit Sub
        End If
        If cmdaceptar.Visible Then
            MsgBox "No puede Salir de la Opcion Hasta que halla Confirmado o Cancelado el Ingreso de la Cuota", vbInformation, "Aviso"
            Cancel = 1
            Exit Sub
        End If
    End If

    If FECalend.TextMatrix(1, 1) = "" Then
            ReDim MatCalendario(0)
    Else
            ReDim MatCalendario(FECalend.Rows - 1, FECalend.Cols)
            For i = 1 To FECalend.Rows - 1
                MatCalendario(i - 1, 0) = FECalend.TextMatrix(i, 0)
                MatCalendario(i - 1, 1) = FECalend.TextMatrix(i, 1)
                MatCalendario(i - 1, 2) = FECalend.TextMatrix(i, 2)
                MatCalendario(i - 1, 3) = FECalend.TextMatrix(i, 3)
                MatCalendario(i - 1, 4) = FECalend.TextMatrix(i, 4)
                MatCalendario(i - 1, 5) = FECalend.TextMatrix(i, 5)
            Next i
    End If
End Sub

Private Sub Optdesemb_Click(Index As Integer)
    If Index = 0 Then
        HabilitaDesembolsoTotal True
    Else
        HabilitaDesembolsoTotal False
    End If
End Sub


Private Sub txtinteres_GotFocus()
    fEnfoque txtinteres
End Sub

Private Sub txtinteres_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtinteres, KeyAscii)
    If KeyAscii = 13 And DTPFecdes.Enabled Then
        DTPFecdes.SetFocus
    End If
End Sub

Private Sub txtinteres_LostFocus()
    If Trim(txtinteres.Text) = "" Then
        txtinteres.Text = "0.00"
    Else
        txtinteres.Text = Format(txtinteres.Text, "#0.00")
    End If
End Sub

Private Sub txtMonto_GotFocus()
    fEnfoque txtMonto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonto, KeyAscii)
    If KeyAscii = 13 Then
        txtinteres.SetFocus
    End If
End Sub

Private Sub txtMonto_LostFocus()
    If Trim(txtMonto.Text) = "" Then
        txtMonto.Text = "0.00"
    Else
        txtMonto.Text = Format(txtMonto.Text, "#0.00")
    End If
End Sub

Sub Inicio()
    lsNegritaOn = Chr$(27) + Chr$(71)
    lsNegritaOff = Chr$(27) + Chr$(72)
    lsSaltoLin = Chr$(10)
    EspHorxPag = 170
    EspVerxPag = 56
    nPuntPag = 0
    lsTab = Space(1)
End Sub
