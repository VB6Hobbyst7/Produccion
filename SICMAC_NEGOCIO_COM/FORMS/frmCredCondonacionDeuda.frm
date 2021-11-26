VERSION 5.00
Begin VB.Form frmCredCondonacionDeuda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Condonación de Deuda"
   ClientHeight    =   6345
   ClientLeft      =   2430
   ClientTop       =   3180
   ClientWidth     =   9915
   Icon            =   "frmCredCondonacionDeuda.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkGastos 
      Caption         =   "Gastos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8430
      TabIndex        =   24
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CheckBox chkIntVenc 
      Caption         =   "Int. Venc"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7230
      TabIndex        =   23
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CheckBox chkMora 
      Caption         =   "Mora"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6030
      TabIndex        =   22
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CheckBox chkInteres 
      Caption         =   "Intereses"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4830
      TabIndex        =   21
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CheckBox chkCapital 
      Caption         =   "Capital"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3630
      TabIndex        =   20
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CheckBox chkTotalDeuda 
      Caption         =   "Total Deuda"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1710
      TabIndex        =   19
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txttotalGasto 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8430
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4920
      Width           =   1200
   End
   Begin VB.TextBox txttotalCapital 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3630
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4920
      Width           =   1200
   End
   Begin VB.TextBox txtTotalCuota 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4920
      Width           =   1200
   End
   Begin VB.TextBox txttotalIntVenc 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   7230
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4920
      Width           =   1200
   End
   Begin VB.TextBox txttotalMora 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6030
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4920
      Width           =   1200
   End
   Begin VB.TextBox txttotalInteres 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4830
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9870
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   661
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblMetLiq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1800
         TabIndex        =   29
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label6 
         Caption         =   "Met. Liquidación :"
         Height          =   240
         Left            =   480
         TabIndex        =   28
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label lblCuota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   240
         Left            =   6120
         TabIndex        =   17
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Titular :"
         Height          =   240
         Left            =   3855
         TabIndex        =   10
         Top             =   270
         Width           =   570
      End
      Begin VB.Label LblTitu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   4560
         TabIndex        =   9
         Top             =   270
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "Prestamo :"
         Height          =   240
         Left            =   2865
         TabIndex        =   8
         Top             =   840
         Width           =   750
      End
      Begin VB.Label LblMonto 
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
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   3675
         TabIndex        =   7
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label5 
         Caption         =   "Nro Calendario :"
         Height          =   240
         Left            =   4830
         TabIndex        =   6
         Top             =   870
         Width           =   1170
      End
      Begin VB.Label LblDiasAtraso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   240
         Left            =   7800
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Dias Atraso: "
         Height          =   240
         Left            =   6840
         TabIndex        =   4
         Top             =   870
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   5580
      Width           =   9930
      Begin VB.CommandButton CmdCondonar 
         Caption         =   "Condonar Deuda"
         Height          =   405
         Left            =   1920
         TabIndex        =   25
         Top             =   210
         Width           =   1530
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   405
         Left            =   8205
         TabIndex        =   2
         Top             =   210
         Width           =   1350
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "Nueva Busqueda"
         Height          =   405
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Width           =   1530
      End
   End
   Begin SICMACT.FlexEdit FECalendario 
      Height          =   3465
      Left            =   0
      TabIndex        =   27
      Top             =   1320
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   6112
      Cols0           =   9
      FixedCols       =   0
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "No-Fecha-Cuota-Estado-Capital-Interes-Mora-Int Venc-Gastos"
      EncabezadosAnchos=   "400-1200-1200-700-1200-1200-1200-1200-1200"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   1
      ListaControles  =   "0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
      TextArray0      =   "No"
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColor       =   -2147483630
      ForeColorFixed  =   -2147483635
      CellForeColor   =   -2147483630
   End
   Begin VB.Label Label2 
      Caption         =   "Condonar:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   5280
      Width           =   1095
   End
End
Attribute VB_Name = "frmCredCondonacionDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bPagoCuotas As Boolean
Private sCtaCod As String
Dim oCredito As COMNCredito.NCOMCredito
'Dim MatCalenPend As Variant
Dim MatCalend As Variant
Dim MatCalendDistrib As Variant
Dim sMonto, nSumCuota, nSumCapital, nSumInteres, nSumMora, nSumIntVenc, nSumGastos As Double

Public Sub PagoCuotas(ByVal psCtaCod As String)
    bPagoCuotas = True
    ActxCta.NroCuenta = psCtaCod
    
    Call ActxCta_KeyPress(13)
    
    CboCuota.Enabled = False
    CmdNuevo.Enabled = False
    Me.Show 1
End Sub

Public Sub Inicio()
    bPagoCuotas = False
    Me.Show 1
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)

Dim oCalend As COMDCredito.DCOMCalendario
Dim oCredDat As COMDCredito.DCOMCredito
Dim rsCredVig As ADODB.Recordset

Dim dFecPago As String
Dim i As Integer
Dim pos As Integer

'Dim rsCredVig As ADODB.Recordset
Dim rsCalend As ADODB.Recordset

'Dim MatCalend As Variant
'Dim MatCalendDistrib As Variant
'Dim oCredN As COMNCredito.NCOMCredito
    
    If KeyAscii = 13 Then
        bPagoCuotas = True
        Set oCredDat = New COMDCredito.DCOMCredito
        'Set rsCredVig = oCredDat.RecuperaDatosCreditoVigente(ActxCta.NroCuenta)
        
        Set oCalend = New COMDCredito.DCOMCalendario
        Set rsCalend = oCalend.RecuperaCalendarioPagosPendiente(ActxCta.NroCuenta, False, False)
'        Set oCredN = New COMNCredito.NCOMCredito
'        MatCalend = oCredN.RecuperaMatrizCalendarioPendiente(ActxCta.NroCuenta)
'        MatCalendDistrib = oCredN.CrearMatrizparaAmortizacion(MatCalend)
        
        Set rsCredVig = oCredDat.RecuperaDatosCreditoJudicial(ActxCta.NroCuenta)
        
        'Call oCredDat.CargarHistoriaCreditoJUDICIAL(ActxCta.NroCuenta, bPagoCuotas, rsCredVig, rsCalend)
        'Set oCredDat = Nothing
        
        If rsCredVig.RecordCount > 0 Then
            
            LblTitu.Caption = Space(3) & PstaNombre(rsCredVig!cPersNombre)
            'ChkMiViv.value = rsCredVig!bMiVivienda
            LblMonto.Caption = Format(rsCredVig!nMontoCol, "#0.00")
            
            'CboCuota.Clear
            'For i = 1 To rsCredVig!nNroCalen
            '    CboCuota.AddItem i
            'Next i
            lblCuota.Caption = rsCredVig.Fields("nNroCalen")
            lblMetLiq.Caption = rsCredVig.Fields("cMetLiquidacion")
            If Not bPagoCuotas Then
               'CboCuota.ListIndex = 0
               LblDiasAtraso.Visible = False
               Label4.Visible = False
               lblCuota.Visible = False
            Else
               pos = i - 1
               'CboCuota.ListIndex = IndiceListaCombo(CboCuota, pos)
               LblDiasAtraso.Visible = True
               Label4.Visible = True
               LblDiasAtraso.Visible = Format(rsCredVig!nDiasAtraso, "0")
               lblCuota.Visible = True
            End If
            'Set oCred = New COMDCredito.DCOMCalendario
            'If bPagoCuotas = False Then
            '    Set rsCredVig = oCred.RecuperaCalendarioPagos(ActxCta.NroCuenta, CInt(CboCuota.Text))
            'Else
            '    Set rsCredVig = oCred.RecuperaCalendarioPagos(ActxCta.NroCuenta, CInt(CboCuota.Text), , , , True)
            'End If
            'Set oCred = Nothing
            
            nSumCuota = 0#
            nSumCapital = 0#
            nSumInteres = 0#
            nSumMora = 0#
            nSumIntVenc = 0#
            nSumGastos = 0#
            
            If rsCalend.RecordCount > 0 Then ReDim MatCalenPend(rsCalend.RecordCount, 9) As Variant
            
            LimpiaFlex FECalendario
            Do While Not rsCalend.EOF
                If rsCalend.Bookmark <> 1 Then
                    FECalendario.AdicionaFila
                End If
                If rsCalend!nColocCalendEstado = 1 Then
                    'FECalend.ForeColorRow vbRed
                    FECalendario.BackColorRow vbYellow
                End If
                
                'Agregado por LMMD
                If rsCalend!nColocCalendEstado = 0 And DateDiff("d", rsCalend!dvenc, gdFecSis) >= 0 Then
                    FECalendario.BackColorRow vbRed
                    FECalendario.ForeColorRow vbWhite
                End If
                
'                If IsNull(rsCalend!dPago) Then
'                    dFecPago = ""
'                Else
'                    If Year(rsCalend!dPago) = "1900" Then
'                        dFecPago = ""
'                    Else
'                        dFecPago = Format(rsCalend!dPago, "dd/mm/yyyy")
'                    End If
'                End If
                
                If Not rsCalend!nColocCalendEstado = 0 Then  ''si la cuota esta cancelada
'                    FECalendario.TextMatrix(rsCalend.Bookmark, 0) = Trim(Str(rsCalend!nCuota))
'                    FECalendario.TextMatrix(rsCalend.Bookmark, 1) = Format(rsCalend!dvenc, "dd/mm/yyyy")
'                    FECalendario.TextMatrix(rsCalend.Bookmark, 2) = Format(rsCalend!nCapital + rsCalend!nIntComp + rsCalend!nIntGracia + rsCalend!nIntMor + rsCalend!nIntReprog + rsCalend!nIntSuspenso + rsCalend!nGasto + rsCalend!nIntCompVenc, "#0.00")
'                    FECalendario.TextMatrix(rsCalend.Bookmark, 3) = IIf(rsCalend!nColocCalendEstado = 0, "Pend.", "Cancel.")
'                    FECalendario.TextMatrix(rsCalend.Bookmark, 4) = Format(rsCalend!nCapital, "#0.00")
'                    FECalendario.TextMatrix(rsCalend.Bookmark, 5) = Format(rsCalend!nIntComp + rsCalend!nIntGracia + rsCalend!nIntReprog + rsCalend!nIntSuspenso, "#0.00")
'                    FECalendario.TextMatrix(rsCalend.Bookmark, 6) = Format(rsCalend!nIntMor, "#0.00")
'                    FECalendario.TextMatrix(rsCalend.Bookmark, 7) = Format(rsCalend!nIntCompVenc, "#0.00")
'                    FECalendario.TextMatrix(rsCalend.Bookmark, 8) = Format(rsCalend!nGasto, "#0.00")
                Else
                    FECalendario.TextMatrix(rsCalend.Bookmark, 0) = Trim(Str(rsCalend!nCuota))
                    FECalendario.TextMatrix(rsCalend.Bookmark, 1) = Format(rsCalend!dvenc, "dd/mm/yyyy")
                    FECalendario.TextMatrix(rsCalend.Bookmark, 2) = Format(IIf(IsNull(rsCalend!nCapital), 0#, rsCalend!nCapital) + IIf(IsNull(rsCalend!nIntComp), 0#, rsCalend!nIntComp) + IIf(IsNull(rsCalend!nIntGracia), 0#, rsCalend!nIntGracia) + IIf(IsNull(rsCalend!nIntMor), 0#, rsCalend!nIntMor) + IIf(IsNull(rsCalend!nIntReprog), 0#, rsCalend!nIntReprog) + IIf(IsNull(rsCalend!nIntSuspenso), 0#, rsCalend!nIntSuspenso) + IIf(IsNull(rsCalend!nGasto), 0#, rsCalend!nGasto) + IIf(IsNull(rsCalend!nIntCompVenc), 0#, rsCalend!nIntCompVenc), "#0.00")
                    FECalendario.TextMatrix(rsCalend.Bookmark, 3) = IIf(rsCalend!nColocCalendEstado = 0, "Pend.", "Cancel.")
                    FECalendario.TextMatrix(rsCalend.Bookmark, 4) = Format(IIf(IsNull(rsCalend!nCapital), 0#, rsCalend!nCapital), "#0.00")
                    FECalendario.TextMatrix(rsCalend.Bookmark, 5) = Format(IIf(IsNull(rsCalend!nIntComp), 0#, rsCalend!nIntComp) + IIf(IsNull(rsCalend!nIntGracia), 0#, rsCalend!nIntGracia) + IIf(IsNull(rsCalend!nIntReprog), 0#, rsCalend!nIntReprog) + IIf(IsNull(rsCalend!nIntSuspenso), 0#, rsCalend!nIntSuspenso), "#0.00")
                    FECalendario.TextMatrix(rsCalend.Bookmark, 6) = Format(IIf(IsNull(rsCalend!nIntMor), 0#, rsCalend!nIntMor), "#0.00")
                    FECalendario.TextMatrix(rsCalend.Bookmark, 7) = Format(IIf(IsNull(rsCalend!nIntCompVenc), 0#, rsCalend!nIntCompVenc), "#0.00")
                    FECalendario.TextMatrix(rsCalend.Bookmark, 8) = Format(IIf(IsNull(rsCalend!nGasto), 0#, rsCalend!nGasto), "#0.00")
                    
                    nSumCuota = nSumCuota + (IIf(IsNull(rsCalend!nCapital), 0#, rsCalend!nCapital) + IIf(IsNull(rsCalend!nIntComp), 0#, rsCalend!nIntComp) + IIf(IsNull(rsCalend!nIntGracia), 0#, rsCalend!nIntGracia) + IIf(IsNull(rsCalend!nIntMor), 0#, rsCalend!nIntMor) + IIf(IsNull(rsCalend!nIntReprog), 0#, rsCalend!nIntReprog) + IIf(IsNull(rsCalend!nIntSuspenso), 0#, rsCalend!nIntSuspenso) + IIf(IsNull(rsCalend!nGasto), 0#, rsCalend!nGasto) + IIf(IsNull(rsCalend!nIntCompVenc), 0#, rsCalend!nIntCompVenc))
                    nSumCapital = nSumCapital + (IIf(IsNull(rsCalend!nCapital), 0#, rsCalend!nCapital))
                    If gdFecSis >= CDate(rsCalend!dvenc) Then
                        nSumInteres = nSumInteres + (IIf(IsNull(rsCalend!nIntComp), 0#, rsCalend!nIntComp) + IIf(IsNull(rsCalend!nIntGracia), 0#, rsCalend!nIntGracia) + IIf(IsNull(rsCalend!nIntReprog), 0#, rsCalend!nIntReprog) + IIf(IsNull(rsCalend!nIntSuspenso), 0#, rsCalend!nIntSuspenso))
                        nSumMora = nSumMora + (IIf(IsNull(rsCalend!nIntMor), 0#, rsCalend!nIntMor))
                        nSumIntVenc = nSumIntVenc + (IIf(IsNull(rsCalend!nIntCompVenc), 0#, rsCalend!nIntCompVenc))
                        nSumGastos = nSumGastos + (IIf(IsNull(rsCalend!nGasto), 0#, rsCalend!nGasto))
                    End If
                End If
                
'                If bPagoCuotas = True Then
'                    FECalendario.TextMatrix(rsCalend.Bookmark, 15) = IIf(IsNull(rsCalend!Cuser), "", rsCalend!Cuser)
'                End If
                
                rsCalend.MoveNext
            Loop
            ActxCta.Enabled = False
            FECalendario.Enabled = True
            
            txtTotalCuota = Format(nSumCapital + nSumInteres + nSumMora + nSumIntVenc + nSumGastos, "###.00")
            txttotalCapital = Format(nSumCapital, "###.00")
            txttotalInteres = Format(nSumInteres, "###.00")
            txttotalMora = Format(nSumMora, "###.00")
            txttotalIntVenc = Format(nSumIntVenc, "###.00")
            txttotalGasto = Format(nSumGastos, "###.00")
        Else
            MsgBox "No Se pudo encontrar el Credito o no esta Vigente"
            'rsCalend.Close
            Exit Sub
        End If
    End If
End Sub

Private Sub chkCapital_Click()
    If chkCapital.value = 1 Then chkTotalDeuda.value = 0
End Sub

Private Sub chkGastos_Click()
    If chkGastos.value = 1 Then
        If CDbl(txttotalGasto) = 0 Or txttotalGasto = "" Then
            MsgBox "El monto a condonar debe ser mayor a cero ", vbInformation, "Condonar Deuda"
            chkGastos.value = 0
            Exit Sub
        End If
    Else
        chkTotalDeuda.value = 0
    End If
End Sub

Private Sub chkInteres_Click()
    If chkInteres.value = 1 Then
        If CDbl(txttotalInteres) = 0 Or txttotalInteres = "" Then
            MsgBox "El monto a condonar debe ser mayor a cero ", vbInformation, "Condonar Deuda"
            chkInteres.value = 0
            Exit Sub
        End If
    Else
        chkTotalDeuda.value = 0
    End If
End Sub

Private Sub chkIntVenc_Click()
    If chkIntVenc.value = 1 Then
        If CDbl(txttotalIntVenc) = 0 Or txttotalIntVenc = "" Then
            MsgBox "El monto a condonar debe ser mayor a cero ", vbInformation, "Condonar Deuda"
            chkIntVenc.value = 0
            Exit Sub
        End If
    Else
        chkTotalDeuda.value = 0
    End If
End Sub

Private Sub chkMora_Click()
    If chkMora.value = 1 Then
        If CDbl(txttotalMora) = 0 Or txttotalMora = "" Then
            MsgBox "El monto a condonar debe ser mayor a cero ", vbInformation, "Condonar Deuda"
            chkMora.value = 0
            Exit Sub
        End If
    Else
        chkTotalDeuda.value = 0
    End If
End Sub

Private Sub chkTotalDeuda_Click()
    If chkTotalDeuda.value = 1 Then
        If CDbl(txtTotalCuota) = 0 Or txtTotalCuota = "" Then
            MsgBox "El monto a condonar debe ser mayor a cero ", vbInformation, "Condonar Deuda"
            chkTotalDeuda.value = 0
            Exit Sub
        End If
    Else
        chkCapital.value = 0
        chkInteres.value = 0
        chkMora.value = 0
        chkIntVenc.value = 0
        chkGastos.value = 0
    End If
End Sub

Private Sub CmdCondonar_Click()
Dim itm As ListItem
Dim lcFlagCondonacion As String
Dim MatCptos As Variant
Dim nSaldoK As Double
Dim nMonto

    nMonto = 0#
    lcFlagCondonacion = "000000"
    If chkTotalDeuda.value = 1 Then
        lcFlagCondonacion = "011111"
        nMonto = CDbl(txttotalCapital) + CDbl(txttotalInteres) + CDbl(txttotalMora) + CDbl(txttotalIntVenc) + CDbl(txttotalGasto)
        nSaldoK = 0
    Else
        If chkCapital.value = 1 Then
            lcFlagCondonacion = Left$(lcFlagCondonacion, 1) & "1" & Right$(lcFlagCondonacion, 4)
            nMonto = nMonto + CDbl(txttotalCapital)
            nSaldoK = 0
        Else
            lcFlagCondonacion = Left$(lcFlagCondonacion, 1) & "0" & Right$(lcFlagCondonacion, 4)
            nSaldoK = CDbl(txttotalCapital)
        End If
        If chkInteres.value = 1 Then
            lcFlagCondonacion = Left$(lcFlagCondonacion, 2) & "1" & Right$(lcFlagCondonacion, 3)
            nMonto = nMonto + CDbl(txttotalInteres)
        Else
            lcFlagCondonacion = Left$(lcFlagCondonacion, 2) & "0" & Right$(lcFlagCondonacion, 3)
        End If
        If chkMora.value = 1 Then
            lcFlagCondonacion = Left$(lcFlagCondonacion, 3) & "1" & Right$(lcFlagCondonacion, 2)
            nMonto = nMonto + CDbl(txttotalMora)
        Else
            lcFlagCondonacion = Left$(lcFlagCondonacion, 3) & "0" & Right$(lcFlagCondonacion, 2)
        End If
        If chkIntVenc.value = 1 Then
            lcFlagCondonacion = Left$(lcFlagCondonacion, 4) & "1" & Right$(lcFlagCondonacion, 1)
            nMonto = nMonto + CDbl(txttotalIntVenc)
        Else
            lcFlagCondonacion = Left$(lcFlagCondonacion, 4) & "0" & Right$(lcFlagCondonacion, 1)
        End If
        If chkGastos.value = 1 Then
            lcFlagCondonacion = Left$(lcFlagCondonacion, 5) & "1" & Right$(lcFlagCondonacion, 0)
            nMonto = nMonto + CDbl(txttotalGasto)
        Else
            lcFlagCondonacion = Left$(lcFlagCondonacion, 5) & "1" & Right$(lcFlagCondonacion, 0)
        End If
    End If
    If CondonarDeuda(ActxCta.NroCuenta, Trim(lblMetLiq.Caption), lcFlagCondonacion, nSaldoK, nMonto) Then
        MsgBox "El proceso de condonación se ha realizado satisfactoriamente ", vbInformation, "Condonación de Deuda"
        CmdCondonar.Enabled = False
    Else
        MsgBox "Se ha producido un error en el proceso de Condonacón de Deuda", vbExclamation, "Condonación de Deuda"
    End If
End Sub

Private Sub cmdNuevo_Click()
    ActxCta.Enabled = True
    CmdCondonar.Enabled = True
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    LimpiaFlex FECalendario
    LblMonto.Caption = "0.00"
    LblTitu.Caption = ""
    lblCuota.Caption = ""
    
    txtTotalCuota = 0#
    txttotalCapital = 0#
    txttotalInteres = 0#
    txttotalMora = 0#
    txttotalIntVenc = 0#
    txttotalGasto = 0#
    
    chkTotalDeuda.value = 0
    chkCapital.value = 0
    chkInteres.value = 0
    chkMora.value = 0
    chkIntVenc.value = 0
    chkGastos.value = 0
        
    'ChkMiViv.value = 0
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If Not bPagoCuotas Then
       ActxCta.NroCuenta = ""
       ActxCta.CMAC = gsCodCMAC
       ActxCta.Age = gsCodAge
    End If
'    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CentraForm Me
End Sub

Private Function CondonarDeuda(ByVal pcCtaCod As String, ByVal psMetLiq As String, ByVal psFlag As String, ByVal psSaldoKFecha As String, ByVal pnMonto As Double) As Boolean
Dim bConforme As Boolean
    Set oCredito = New COMNCredito.NCOMCredito
    MatCalend = oCredito.RecuperaMatrizCalendarioPendiente(pcCtaCod)
    MatCalendDistrib = oCredito.MatrizDistribuirMontoAcondonar(MatCalend, pnMonto, psMetLiq, psFlag, gdFecSis)
    Call oCredito.CondonarDeuda(pcCtaCod, MatCalend, MatCalendDistrib, pnMonto, gdFecSis, psMetLiq, psFlag, 5, gsCodAge, gsCodUser, bConforme)
    CondonarDeuda = bConforme
End Function


