VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredConsultaGarantDPF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Créditos con Garantía DPF"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   Icon            =   "frmCredConsultaGarantDPF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Left            =   9240
      TabIndex        =   15
      Top             =   5280
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Créditos Garantizados"
      TabPicture(0)   =   "frmCredConsultaGarantDPF.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSaldoDPF"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblSaldoCred"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblOtrosBloq"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblLineaDisp"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "feCredGarant"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin SICMACT.FlexEdit feCredGarant 
         Height          =   1815
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   10140
         _ExtentX        =   17886
         _ExtentY        =   3201
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nº-Nro Credito-Agencia-Analista-F. Desemb-Cuotas Pend.-Moneda-Cuota-Saldo Cred."
         EncabezadosAnchos=   "300-1800-1800-800-1200-1050-800-1000-1000"
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
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "R-L-L-C-C-C-C-R-R"
         FormatosEdit    =   "3-0-0-0-0-0-0-2-2"
         TextArray0      =   "Nº"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblLineaDisp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8880
         TabIndex        =   21
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Línea Disponible"
         Height          =   195
         Left            =   7560
         TabIndex        =   20
         Top             =   2445
         Width           =   1200
      End
      Begin VB.Label lblOtrosBloq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6000
         TabIndex        =   19
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Otros Bloq. :"
         Height          =   195
         Left            =   5040
         TabIndex        =   18
         Top             =   2445
         Width           =   870
      End
      Begin VB.Label lblSaldoCred 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3600
         TabIndex        =   17
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Cred. :"
         Height          =   195
         Left            =   2640
         TabIndex        =   16
         Top             =   2445
         Width           =   915
      End
      Begin VB.Label lblSaldoDPF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Saldo DPF MN :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2445
         Width           =   1155
      End
   End
   Begin VB.Frame FraDatos 
      Caption         =   " Datos de la cuenta "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Texto           =   "N° Cuenta:"
      End
      Begin SICMACT.FlexEdit feCliente 
         Height          =   1605
         Left            =   3840
         TabIndex        =   10
         Top             =   240
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   2831
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   1
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "Nº-Codigo-Nombre-Relacion"
         EncabezadosAnchos=   "300-1500-2700-1500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         ColumnasAEditar =   "X-1-X-3"
         TextStyleFixed  =   4
         ListaControles  =   "0-1-0-3"
         EncabezadosAlineacion=   "C-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "Nº"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
      End
      Begin VB.Label lblTipoCuenta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1605
         Width           =   960
      End
      Begin VB.Label lblCobert 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3000
         TabIndex        =   7
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cobertura :"
         Height          =   195
         Left            =   2160
         TabIndex        =   6
         Top             =   1245
         Width           =   780
      End
      Begin VB.Label lblTEA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "T.E.A. :"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1250
         Width           =   540
      End
      Begin VB.Label lblSubProd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "SubProducto :"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   880
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmCredConsultaGarantDPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCredConsultaGarantDPF
'** Descripción : Formulario para mostrar datos de la Garantia DPF según TI-ERS138-2013
'** Creación : JUEZ, 20140103 03:30:00 PM
'*****************************************************************************************************

Option Explicit

Dim rs As ADODB.Recordset
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

Public Sub Inicio(ByVal psCtaCod As String, ByVal pMatSaldos As Variant, ByVal psCobert As String, ByVal cTpoPrograma As String)
    txtCuenta.NroCuenta = psCtaCod
    lblSubProd.Caption = cTpoPrograma
    
    lblCobert.Caption = psCobert
    
    lblSaldoDPF.Caption = Format(pMatSaldos(0, 1), "#,##0.00")
    lblOtrosBloq.Caption = Format(pMatSaldos(3, 1), "#,##0.00")
    lblLineaDisp.Caption = Format(pMatSaldos(1, 1), "#,##0.00")
    
    If CargaDatos(psCtaCod) Then
        Me.Show 1
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
    Dim oNCapGen As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim oDCred As COMDCredito.DCOMCredito
    Dim rsPers As ADODB.Recordset, rsCred As ADODB.Recordset
    Dim lnFila As Integer
    Dim nSaldoCred As Double, nSumSaldoCred As Double
    
    CargaDatos = False
    
    Set oNCapGen = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rs = oNCapGen.GetDatosCuenta(psCtaCod)
    If rs.RecordCount > 0 Then
        lblTEA.Caption = Format(ConvierteTNAaTEA(rs("nTasaInteres")), "#,##0.00") & " % "
        lblTipoCuenta.Caption = rs("cTipoCuenta")
        
        Set rsPers = oNCapGen.GetProductoPersona(psCtaCod, gCapRelPersPromotor)
        If rsPers.RecordCount > 0 Then
            lnFila = 0
            Do While Not rsPers.EOF
                feCliente.AdicionaFila
                lnFila = feCliente.row
                feCliente.TextMatrix(lnFila, 1) = rsPers!cPersCod
                feCliente.TextMatrix(lnFila, 2) = rsPers!Nombre
                feCliente.TextMatrix(lnFila, 3) = rsPers!cRelacion
                rsPers.MoveNext
            Loop
        End If
        
        Set oDCred = New COMDCredito.DCOMCredito
        Set rsCred = oDCred.RecuperaCreditosGarantiaDPF(psCtaCod)
        If rsCred.RecordCount > 0 Then
            lnFila = 0
            Do While Not rsCred.EOF
                feCredGarant.AdicionaFila
                lnFila = feCredGarant.row
                feCredGarant.TextMatrix(lnFila, 1) = rsCred!cCtaCod
                feCredGarant.TextMatrix(lnFila, 2) = rsCred!cAgeDescripcion
                feCredGarant.TextMatrix(lnFila, 3) = rsCred!cAnalista
                feCredGarant.TextMatrix(lnFila, 4) = Format(rsCred!dVigencia, "dd/MM/yyyy")
                feCredGarant.TextMatrix(lnFila, 5) = rsCred!nCuotasPend
                feCredGarant.TextMatrix(lnFila, 6) = rsCred!cMoneda
                feCredGarant.TextMatrix(lnFila, 7) = Format(rsCred!nCuota, "#,##0.00")
                
                nSaldoCred = CalculaSaldoCred(rsCred!cCtaCod, rsCred!nPorcMontoCobert)
                feCredGarant.TextMatrix(lnFila, 8) = Format(nSaldoCred, "#,##0.00")
                
                nSumSaldoCred = nSumSaldoCred + nSaldoCred
                rsCred.MoveNext
            Loop
            lblSaldoCred.Caption = Format(nSumSaldoCred, "#,##0.00")
        End If
        CargaDatos = True
    Else
        MsgBox "No se puede obtener datos de la cuenta", vbInformation, "Aviso"
    End If
End Function

Private Function CalculaSaldoCred(ByVal psCtaCod As String, ByVal pnPorcMontoCobert As Double) As Double
    Dim oNCred As COMNCredito.NCOMCredito
    Dim oDCred As COMDCredito.DCOMCredito
    Dim oDCredAct As COMDCredito.DCOMCredActBD
    Dim oNCF As COMNCartaFianza.NCOMCartaFianzaValida
    Dim rsCred As ADODB.Recordset, RGas As ADODB.Recordset
    Dim MatCalend As Variant
    Dim nSaldoKFecha As Double, nIntCompFecha As Double, nGastoFecha As Double, nIntMorFecha As Double, nDeudaFecha As Double
    Dim nSaldoCredCob As Double
    
    CalculaSaldoCred = 0
    
    Set oNCred = New COMNCredito.NCOMCredito
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000093", Mid(psCtaCod, 6, 3)) Then
    'If Mid(psCtaCod, 6, 3) = "121" Or Mid(psCtaCod, 6, 3) = "221" Or Mid(psCtaCod, 6, 3) = "514" Then
    '**ARLO20180712 ERS042 - 2018
        Set oNCF = New COMNCartaFianza.NCOMCartaFianzaValida
        Set RGas = oNCF.RecuperaDatosT(psCtaCod)
        Set oNCF = Nothing
        nDeudaFecha = RGas!nSaldo
    Else
        MatCalend = oNCred.RecuperaMatrizCalendarioPendienteHistorial(psCtaCod)
        nSaldoKFecha = Format(oNCred.MatrizCapitalAFecha(psCtaCod, MatCalend), "#0.00")
        nIntCompFecha = Format(oNCred.MatrizInteresTotalesAFechaSinMora(psCtaCod, MatCalend, gdFecSis) + oNCred.MatrizInteresGraAFecha(psCtaCod, MatCalend, gdFecSis), "#0.00")
        Set oDCredAct = New COMDCredito.DCOMCredActBD
        Set RGas = oDCredAct.CargaRecordSet(" SELECT nGasto=dbo.ColocCred_ObtieneGastoFechaCredito('" & psCtaCod & "','" & Format(gdFecSis, "mm/dd/yyyy") & "')")
        nGastoFecha = RGas!nGasto
        Set oDCredAct = Nothing
        nIntMorFecha = Format(oNCred.ObtenerMoraVencida(gdFecSis, MatCalend), "#0.00")
        
        nDeudaFecha = (nSaldoKFecha + nIntCompFecha + nGastoFecha + nIntMorFecha)
    End If
    nSaldoCredCob = pnPorcMontoCobert * nDeudaFecha
    'CalculaSaldoCred = CalculaSaldoCred + nDeudaFecha
    CalculaSaldoCred = CalculaSaldoCred + nSaldoCredCob
    
    Set oNCred = Nothing
End Function
