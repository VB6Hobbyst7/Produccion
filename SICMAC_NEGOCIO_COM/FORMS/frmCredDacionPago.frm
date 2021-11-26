VERSION 5.00
Object = "{2E37A9B8-9906-4590-9F3A-E67DB58122F5}#7.0#0"; "OcxLabelX.ocx"
Begin VB.Form frmCredDacionPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dacion de pago"
   ClientHeight    =   5265
   ClientLeft      =   3180
   ClientTop       =   1980
   ClientWidth     =   6060
   Icon            =   "frmCredDacionPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   75
      TabIndex        =   30
      Top             =   4515
      Width           =   5940
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   405
         Left            =   4725
         TabIndex        =   33
         Top             =   195
         Width           =   1110
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   405
         Left            =   3570
         TabIndex        =   32
         Top             =   195
         Width           =   1110
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
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
         Height          =   405
         Left            =   90
         TabIndex        =   31
         Top             =   195
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3675
      Left            =   75
      TabIndex        =   4
      Top             =   840
      Width           =   5940
      Begin OcxLabelX.LabelX LblXMoneda 
         Height          =   420
         Left            =   1155
         TabIndex        =   8
         Top             =   630
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   16711680
         Bold            =   0   'False
         Alignment       =   2
      End
      Begin OcxLabelX.LabelX LblXNomCli 
         Height          =   420
         Left            =   1155
         TabIndex        =   6
         Top             =   195
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   16711680
         Bold            =   -1  'True
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX LblXMontoCred 
         Height          =   420
         Left            =   4020
         TabIndex        =   10
         Top             =   645
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   1
      End
      Begin OcxLabelX.LabelX LblXSaldoCap 
         Height          =   420
         Left            =   1155
         TabIndex        =   12
         Top             =   1050
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   1
      End
      Begin OcxLabelX.LabelX LblXDeuda 
         Height          =   420
         Left            =   4020
         TabIndex        =   14
         Top             =   1050
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   1
      End
      Begin OcxLabelX.LabelX LblXForPag 
         Height          =   420
         Left            =   1155
         TabIndex        =   17
         Top             =   1440
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   2
      End
      Begin OcxLabelX.LabelX LblXCuota 
         Height          =   420
         Left            =   4020
         TabIndex        =   19
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   1
      End
      Begin OcxLabelX.LabelX LblXCuotaPend 
         Height          =   420
         Left            =   1155
         TabIndex        =   22
         Top             =   1845
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   2
      End
      Begin OcxLabelX.LabelX LblXDiasAtr 
         Height          =   420
         Left            =   4020
         TabIndex        =   23
         Top             =   1830
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   2
      End
      Begin OcxLabelX.LabelX LblXMetLiq 
         Height          =   420
         Left            =   1155
         TabIndex        =   25
         Top             =   2220
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   2
      End
      Begin OcxLabelX.LabelX LblXCredito 
         Height          =   420
         Left            =   3330
         TabIndex        =   29
         Top             =   3210
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   16711680
         Bold            =   -1  'True
         Alignment       =   2
      End
      Begin OcxLabelX.LabelX LblXMora 
         Height          =   420
         Left            =   4020
         TabIndex        =   34
         Top             =   2205
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   1
      End
      Begin OcxLabelX.LabelX LblXCalDin 
         Height          =   420
         Left            =   1140
         TabIndex        =   36
         Top             =   2595
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   1
      End
      Begin OcxLabelX.LabelX LblXMontoPago 
         Height          =   450
         Left            =   1275
         TabIndex        =   37
         Top             =   3180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   794
         FondoBlanco     =   0   'False
         Resalte         =   16711680
         Bold            =   0   'False
         Alignment       =   1
      End
      Begin OcxLabelX.LabelX LblITF 
         Height          =   420
         Left            =   4020
         TabIndex        =   38
         Top             =   2580
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         FondoBlanco     =   0   'False
         Resalte         =   16711680
         Bold            =   0   'False
         Alignment       =   1
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "I.T.F.                   :"
         Height          =   195
         Left            =   2580
         TabIndex        =   39
         Top             =   2640
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Calend. Din:"
         Height          =   195
         Left            =   135
         TabIndex        =   35
         Top             =   2670
         Width           =   870
      End
      Begin VB.Label Label7 
         Caption         =   "Credito :"
         Height          =   285
         Left            =   2700
         TabIndex        =   28
         Top             =   3255
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Monto a Pagar :"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   3255
         Width           =   1140
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         X1              =   45
         X2              =   5865
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   45
         X2              =   5880
         Y1              =   3045
         Y2              =   3045
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Mora                    :"
         Height          =   195
         Left            =   2580
         TabIndex        =   26
         Top             =   2280
         Width           =   1305
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Met Liquid. :"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   2295
         Width           =   870
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Dias Atrasados    :"
         Height          =   195
         Left            =   2580
         TabIndex        =   21
         Top             =   1920
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cuota Pend. :"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Monto Cuota        :"
         Height          =   195
         Left            =   2580
         TabIndex        =   18
         Top             =   1530
         Width           =   1320
      End
      Begin VB.Label LblTotCuo 
         AutoSize        =   -1  'True
         Caption         =   "Cuotas"
         Height          =   195
         Left            =   1785
         TabIndex        =   16
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Lbl2 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago  :"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1500
         Width           =   990
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Deuda a la Fecha : "
         Height          =   195
         Left            =   2580
         TabIndex        =   13
         Top             =   1125
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Capital :"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1125
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Monto del Credito  :"
         Height          =   195
         Left            =   2580
         TabIndex        =   9
         Top             =   705
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Moneda        :"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   690
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente           :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   255
         Width           =   1020
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Dacion"
      Height          =   840
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   5940
      Begin OcxLabelX.LabelX LblXNrodacion 
         Height          =   465
         Left            =   1080
         TabIndex        =   2
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   820
         FondoBlanco     =   0   'False
         Resalte         =   16711680
         Bold            =   -1  'True
         Alignment       =   2
      End
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
         Height          =   390
         Left            =   2685
         TabIndex        =   1
         Top             =   285
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dacion Nro :"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   345
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmCredDacionPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private MatCalend As Variant
Private MatCalendDistribuido As Variant
Private nNroTransac As Long
Private bCalenDinamic As Boolean
Private nMontoPago As Double
Private nITF As Double

Private Sub LimpiaDatos()
    LblxNomCli.Caption = ""
    LblXNrodacion.Caption = ""
    LblXMoneda.Caption = ""
    LblXMontoCred.Caption = ""
    LblXSaldoCap.Caption = ""
    LblXDeuda.Caption = ""
    LblXForPag.Caption = ""
    LblXCuota.Caption = ""
    LblXCuotaPend.Caption = ""
    LblXDiasAtr.Caption = ""
    LblXMetLiq.Caption = ""
    LblXMora.Caption = ""
    LblXCalDin.Caption = ""
    LblXMontoPago.Caption = ""
    LblXCredito.Caption = ""
    LblItf.Caption = "0.00"
End Sub

Private Sub HabilitaDatos(ByVal pbHabilita As Boolean)
    cmdGrabar.Enabled = pbHabilita
    CmdBuscar.Enabled = Not pbHabilita
End Sub

Private Sub CargaDatos(ByVal pnDacion As Long)

'Dim oDCred As COMDCredito.DCOMCredito
Dim oNCred As COMNCredito.NCOMCredito
Dim R As ADODB.Recordset
Dim sCuotaPend As String
Dim sDeuda As String
Dim sCuota As String
Dim sMora As String
Dim sCalDin As String
Dim nNroCuota As Integer
Dim sMetLiq As String
Dim nMontoAPagar As Double
Dim sCtaCod As String

'    Set oDCred = New COMDCredito.DCOMCredito
'    Set R = oDCred.RecuperaDatosDacionPagoCredito(pnDacion)
    Set oNCred = New COMNCredito.NCOMCredito
    Call oNCred.CargarDatosDacionPago(pnDacion, gdFecSis, nNroCuota, sMetLiq, nMontoAPagar, _
                                      sCtaCod, R, bCalenDinamic, MatCalend, MatCalendDistribuido, nMontoPago, nITF, _
                                      sCuotaPend, sDeuda, sCuota, sMora, sCalDin)
    Set oNCred = Nothing
    
    If Not R.BOF And Not R.EOF Then
        LblXNrodacion.Caption = pnDacion
        LblxNomCli.Caption = PstaNombre(R!cPersNombre)
        LblXMoneda.Caption = R!cMoneda
        LblXMontoCred.Caption = Format(R!nMontoCol, "#0.00")
        LblXSaldoCap.Caption = Format(R!nSaldo, "#0.00")
        'LblXForPag.Caption = R!nCuotas
        LblXDiasAtr.Caption = R!nDiasAtraso
        'LblXMetLiq.Caption = R!cMetLiquidacion
        'LblXCredito.Caption = R!cCtaCod
        'LblXMontoPago.Caption = Format(R!nValorTotal, "#0.00")
        nNroTransac = R!nTransacc
'        If IsNull(R!nCalendDinamico) Then
'            bCalenDinamic = False
'        Else
'            If R!nCalendDinamico = 1 Then
'                bCalenDinamic = True
'            Else
'                bCalenDinamic = False
'            End If
'        End If
    End If
    R.Close
    Set R = Nothing
    'Set oDCred = Nothing
    
    'Set oNCred = New COMNCredito.NCOMCredito
    'MatCalend = oNCred.RecuperaMatrizCalendarioPendiente(LblXCredito.Caption)
    'LblXCuotaPend.Caption = oNCred.MatrizCuotaPendiente(MatCalend, MatCalend)
    'LblXDeuda.Caption = Format(oNCred.MatrizDeudaAlaFecha(LblXCredito.Caption, MatCalend, gdFecSis), "#0.00")
    'LblXCuota.Caption = Format(oNCred.MatrizMontoCuota(MatCalend, CInt(LblXForPag.Caption)), "#0.00")
    'LblXMora.Caption = Format(oNCred.MatrizInteresMorFecha(LblXCredito.Caption, MatCalend), "#0.00")
    'LblXCalDin.Caption = Format(oNCred.MatrizMontoCalendDinamico(LblXCredito.Caption, MatCalend, gdFecSis), "#0.00")
    
    'MatCalendDistribuido = oNCred.CrearMatrizparaAmortizacion(MatCalend)
    
    LblXForPag.Caption = nNroCuota
    LblXMetLiq.Caption = sMetLiq
    LblXMontoPago.Caption = Format(nMontoAPagar, "#0.00")
    
    LblXCuotaPend.Caption = sCuotaPend
    LblXDeuda.Caption = sDeuda
    LblXCuota.Caption = sCuota
    LblXMora.Caption = sMora
    LblXCalDin.Caption = sCalDin
    LblXCredito.Caption = sCtaCod
    
    'nMontoPago = fgITFCalculaImpuestoIncluido(CDbl(LblXMontoPago.Caption))
    'nITF = Format(CDbl(LblXMontoPago.Caption) - nMontoPago, "0.00")
    
    'El ITF no se va considerar para la Dacion en Pago
    'LblITF.Caption = Format(nITF, "0.00")
    nITF = 0
    '*************************************
    If nMontoPago > CDbl(LblXDeuda.Caption) Then
        MsgBox "Monto De Dacion es Mayor a la Deuda", vbInformation, "Aviso"
        Call cmdCancelar_Click
        Exit Sub
    End If
    
'    If bCalenDinamic And (nMontoPago < CDbl(LblXDeuda.Caption)) Then
'        If nMontoPago > CDbl(LblXCalDin.Caption) Then
'            MatCalendDistribuido = oNCred.MatrizDistribuirCalendDinamico(LblXCredito.Caption, MatCalend, nMontoPago, Trim(LblXMetLiq.Caption), gdFecSis)
'        Else
'            MatCalendDistribuido = oNCred.MatrizDistribuirMonto(MatCalend, nMontoPago, Trim(LblXMetLiq.Caption))
'        End If
'    Else
'        If nMontoPago <> CDbl(LblXDeuda.Caption) Then
'            MatCalendDistribuido = oNCred.MatrizDistribuirMonto(MatCalend, nMontoPago, Trim(LblXMetLiq.Caption))
'        Else
'            MatCalendDistribuido = oNCred.MatrizDistribuirCancelacion(LblXCredito.Caption, MatCalend, nMontoPago, Trim(LblXMetLiq.Caption), gdFecSis)
'        End If
'    End If
'   Set oNCred = Nothing
End Sub

Private Sub cmdBuscar_Click()
Dim nDacion As Long
    nDacion = frmCredPersEstado.BuscaDacionesPago("Daciones en Pago")
    If nDacion <> -1 Then
        Call CargaDatos(nDacion)
        If Trim(LblXNrodacion.Caption) <> "" Then
            HabilitaDatos True
        End If
        
    Else
        HabilitaDatos False
    End If
End Sub

Private Sub cmdCancelar_Click()
    LimpiaDatos
    HabilitaDatos False
End Sub

Private Sub cmdGrabar_Click()

'Dim oNegCred As COMNCredito.NCOMCredito
'Dim oDoc As COMNCredito.NCOMCredDoc
'Dim oConstante As COMDConstantes.DCOMConstantes
Dim oCred As COMNCredito.NCOMCredito
Dim sError As String
Dim sReporte As String
'Dim sTipoCred As String
'Dim MatCalDinam As Variant
Dim oPrevio As Previo.clsPrevio
        
    If nMontoPago > CDbl(Me.LblXDeuda.Caption) Then
        MsgBox "Monto de Pago de la Dacion No Puede Ser Mayor a la Deuda del Credito", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Se va a Efectuar el Pago del Credito, Desea Continuar ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    '    Set oNegCred = New COMNCredito.NCOMCredito
    '    sError = oNegCred.AmortizarCredito(LblXCredito.Caption, MatCalend, MatCalendDistribuido, nMontoPago, gdFecSis, LblXMetLiq.Caption, gColocTipoPagoDacionPago, gsCodAge, gsCodUser, , , , CLng(LblXNrodacion.Caption), , , , , , , , , , , , , nITF)
        Set oCred = New COMNCredito.NCOMCredito
        Call oCred.GrabarDacionPago(LblXCredito.Caption, MatCalend, MatCalendDistribuido, nMontoPago, gdFecSis, LblXMetLiq.Caption, gsCodAge, gsCodUser, CLng(LblXNrodacion.Caption), _
                                    nITF, bCalenDinamic, CDbl(LblXDeuda.Caption), CDbl(LblXCalDin.Caption), LblxNomCli.Caption, gsNomAge, LblXMoneda.Caption, nNroTransac, sLpt, gsCodCMAC, _
                                    sError, sReporte)
        Set oCred = Nothing
        
        If sError <> "" Then
            MsgBox sError, vbInformation, "Aviso"
            Exit Sub
        'Else
            'Verifica si fue un pago para Calendario Dinamico
        '    If bCalenDinamic And (nMontoPago < CDbl(LblXDeuda.Caption)) Then
        '        If nMontoPago > CDbl(LblXCalDin.Caption) Then
        '            MatCalDinam = oNegCred.ReprogramarCreditoenMemoriaTotal(LblXCredito.Caption, gdFecSis)
        '            oNegCred.ReprogramarCredito LblXCredito.Caption, MatCalDinam, 2
        '        End If
        End If
        Set oPrevio = New Previo.clsPrevio
        oPrevio.Show sReporte, "Reporte de Dacion de Pago"
        '    Set oConstante = New COMDConstantes.DCOMConstantes
        '    sTipoCred = oConstante.DameDescripcionConstante(gProducto, CInt(Mid(LblXCredito.Caption, 6, 3)))
        '    Set oConstante = Nothing
        '    Set oDoc = New COMNCredito.NCOMCredDoc
        '    Call oDoc.ImprimeBoleta(LblXCredito.Caption, LblXNomCli.Caption, gsNomAge, LblXMoneda.Caption, _
                oNegCred.MatrizCuotasPagadas(MatCalendDistribuido), gdFecSis, Format(Date, "hh:mm:ss"), nNroTransac + 1, "", _
                oNegCred.MatrizCapitalPagado(MatCalendDistribuido), oNegCred.MatrizIntCompPagado(MatCalendDistribuido), _
                oNegCred.MatrizIntCompVencPagado(MatCalendDistribuido), _
                oNegCred.MatrizIntMorPagado(MatCalendDistribuido), oNegCred.MatrizGastoPag(MatCalendDistribuido), _
                oNegCred.MatrizIntGraciaPagado(MatCalendDistribuido), _
                oNegCred.MatrizIntSuspensoPag(MatCalendDistribuido) + oNegCred.MatrizIntReprogPag(MatCalendDistribuido), _
                oNegCred.MatrizSaldoCapital(MatCalend, MatCalendDistribuido), oNegCred.MatrizFechaCuotaPendiente(MatCalend, MatCalendDistribuido), _
                gsCodUser, sLpt, gsNomAge, , , gsCodCMAC, nITF)
            
            Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes
            '    Call oDoc.ImprimeBoleta(LblXCredito.Caption, LblXNomCli.Caption, gsNomAge, LblXMoneda.Caption, _
                oNegCred.MatrizCuotasPagadas(MatCalendDistribuido), gdFecSis, Format(Date, "hh:mm:ss"), nNroTransac + 1, "", _
                oNegCred.MatrizCapitalPagado(MatCalendDistribuido), oNegCred.MatrizIntCompPagado(MatCalendDistribuido), _
                oNegCred.MatrizIntCompVencPagado(MatCalendDistribuido), _
                oNegCred.MatrizIntMorPagado(MatCalendDistribuido), oNegCred.MatrizGastoPag(MatCalendDistribuido), _
                oNegCred.MatrizIntGraciaPagado(MatCalendDistribuido), _
                oNegCred.MatrizIntSuspensoPag(MatCalendDistribuido) + oNegCred.MatrizIntReprogPag(MatCalendDistribuido), _
                oNegCred.MatrizSaldoCapital(MatCalend, MatCalendDistribuido), oNegCred.MatrizFechaCuotaPendiente(MatCalend, MatCalendDistribuido), _
                gsCodUser, sLpt, gsNomAge, , , gsCodCMAC, nITF)
                
                oPrevio.Show sReporte, "Reporte de Dacion de Pago"
            Loop
            'Set oDoc = Nothing
            
            Call cmdCancelar_Click
'        End If
        'Set oNegCred = Nothing
    'End If
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
End Sub

