VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCarteraFondemi 
   Caption         =   "Cartera FONDEMI"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
   Icon            =   "frmCarteraFondemi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin Sicmact.FlexEdit FECab 
      Height          =   2535
      Left            =   240
      TabIndex        =   30
      Top             =   3840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3625
      Cols0           =   12
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "-Fecha-Paquete-Tip.Infor.-Fec.Ini-Fec.Fin-Núm.Clientes-MontAprob-SaldoTotal-DesemFondemi-Moneda-Estado"
      EncabezadosAnchos=   "400-1200-800-800-1200-1200-1200-0-0-0-0-1200"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-C-C-R-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Cerrar Mes"
      Height          =   495
      Left            =   6360
      TabIndex        =   24
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "Reporte Fondemi"
      Height          =   495
      Left            =   3120
      TabIndex        =   23
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   495
      Left            =   4800
      TabIndex        =   22
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   7920
      TabIndex        =   21
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.ComboBox cboTipInf 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtSaldoTotal 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   2520
         TabIndex        =   29
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtPaquete 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   320
         Left            =   2520
         TabIndex        =   27
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Fecha desembolso de crédito"
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
         Height          =   855
         Left            =   4680
         TabIndex        =   14
         Top             =   2040
         Width           =   4095
         Begin MSMask.MaskEdBox txtFechaIni 
            Height          =   300
            Left            =   960
            TabIndex        =   15
            Top             =   360
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFechaFin 
            Height          =   300
            Left            =   2640
            TabIndex        =   16
            Top             =   360
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
            Caption         =   "Al"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   18
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label7 
            Caption         =   "Del"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cartera Al"
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
         Height          =   855
         Left            =   4680
         TabIndex        =   12
         Top             =   360
         Width           =   4095
         Begin MSMask.MaskEdBox txtFechaDel 
            Height          =   300
            Left            =   2640
            TabIndex        =   13
            Top             =   360
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            Caption         =   "Fecha de cierre de mes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tasas"
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
         Height          =   735
         Left            =   4680
         TabIndex        =   7
         Top             =   1200
         Width           =   4095
         Begin VB.TextBox txtTasaFin 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2880
            TabIndex        =   9
            Text            =   "101.22"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtTasaIni 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   960
            TabIndex        =   8
            Text            =   "12.68"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Máxima"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Mínima"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox txtNroDesembolso 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   320
         Left            =   2520
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox cboDesembolso 
         Height          =   315
         ItemData        =   "frmCarteraFondemi.frx":030A
         Left            =   2520
         List            =   "frmCarteraFondemi.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtMontoAprobado 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Información"
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
         Left            =   240
         TabIndex        =   32
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "Saldo Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Nro Paquete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Nro Desembolso"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Desembolso FONDEMI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Monto Aprobado"
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCarteraFondemi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dFechaIni As Date
Private Sub cboDesembolso_Click()
    Dim dFechaCierre As String
    dFechaCierre = Mid(cboDesembolso.Text, 511, 8)
    txtPaquete.Text = Mid(cboDesembolso.Text, 529, 3)
    txtNroDesembolso.Text = Mid(cboDesembolso.Text, 542, 13)
'    If Right(Trim(cboDesembolso.Text), 8) = "20091109" Then
'            txtPaquete.Text = "001"
'            txtNroDesembolso.Text = "85628"
'    End If
'    If Right(Trim(cboDesembolso.Text), 8) = "20091201" Then
'            txtPaquete.Text = "002"
'            txtNroDesembolso.Text = "86112"
'    End If
End Sub

Private Sub cmdActualizar_Click()
Dim oDCajas As DCaja_Adeudados
Set oDCajas = New DCaja_Adeudados
Call oDCajas.ActualizarFONDEMICAB_Estado(txtFechaDel.Text, txtPaquete.Text)
Set oDCajas = Nothing
'ALPA20130723
MsgBox "Información del cierre " & Format(txtFechaDel.Text, "DD/MM/YYYY") & " fue cerrado satisfactoriamente ", vbInformation, "Aviso"
Call Reporte_Fonddemi 'ALPA20130820
End Sub

Private Sub Reporte_Fonddemi()
Dim sCondBala As String, sCondBala2 As String
Dim sCondBala6 As String
Dim sCond1 As String, sCond2 As String
Dim sCta   As String
Dim sCta2   As String
Dim N      As Integer
Dim nPos   As Variant
Dim FecIni As Date
Dim FecFin As Date
'Recorsed
Dim RDet As New ADODB.Recordset
Dim RCab As New ADODB.Recordset
Dim sCodigo As String
Dim sSalIni As String
Dim sDebe As String
Dim sHaber As String
Dim sSalFin As String
Dim CadRep As String
Dim nTotal As Integer
Dim cPar2(6) As Integer
Dim i As Integer
Dim ContBarra As Long
Dim Total As Integer
Dim CadTemp As String

Dim oDCajas As DCaja_Adeudados
Set oDCajas = New DCaja_Adeudados
Dim lnBalanceCate As Integer


'**
Dim sCodigoInst As String
Dim sFechaProceso As String
Dim sNroPaquete As String
Dim sIndicaTI As String
Dim sMoneda As String
Dim sCodiOpeCofide As String
Dim sFechaPrestamo As String
Dim sMontoPrestamo As String
Dim sSaldoACofide As String
Dim sPenJustificar As String
Dim sTotalImporteSubPrestamos As String
Dim sTotalActivosFijos As String
Dim sTotalCapitalTraba As String
Dim sTotalImpoSaldSubPres As String
Dim sNroSubPrestamos, sCliente, sDocumentoDNI, sDocumentoRUC, sTipoPersona, sSaldoCapital, cPlazo, cPlazoTotal, cCuotas, cSector As String
Dim nSaldoCapital, nPlazo As Currency
'**
Dim fs As New Scripting.FileSystemObject
Dim psArchivoAGrabarC As String
Dim psArchivoAGrabarD As String
Dim sCad As String
Dim ArcSal As Integer
Dim lnSigno As String
'On Error GoTo SucaveERR
psArchivoAGrabarC = App.path & "\SPOOLER\SC" & "0" & Trim(Right(cboMoneda.Text, 4)) & Trim(txtPaquete.Text) & ".051"
psArchivoAGrabarD = App.path & "\SPOOLER\SD" & "0" & Trim(Right(cboMoneda.Text, 4)) & Trim(txtPaquete.Text) & ".051"
lnBalanceCate = 5
   cPar2(0) = 0
   cPar2(1) = 1
   cPar2(2) = 2
   cPar2(3) = 3
   cPar2(4) = 4
   cPar2(5) = 6
   FecIni = txtFechaIni.Text
   FecFin = txtFechaFin.Text
       
   DoEvents
   MousePointer = 11
   'Set oBarra = New clsProgressBar
   'oBarra.ShowForm Me
   'oBarra.Max = 7
   'oBarra.Progress 0, "RABE", "", "Eliminando BCient"
   DoEvents
   CadRep = ""
   CadTemp = ""
   
   
   Set RCab = oDCajas.ObtenerFONDEMICAB(txtFechaDel.Text, txtPaquete.Text)
   If Not RCab.BOF And Not RCab.EOF Then
      'oBarra.Max = RCab.RecordCount
      Do While Not RCab.EOF
        sCodigoInst = RCab!cCodIFI '"051"
        sFechaProceso = Format(RCab!dFechPr, "YYYYMMDD") '"20110630"
        sNroPaquete = RCab!cNroPaq '"001"
        sIndicaTI = RCab!cIndInfo '"S"
        sMoneda = RCab!cMoneda '"01"
        sCodiOpeCofide = FillNum(RCab!cCodCofi, 13, "0") '"0000000085628"
        sFechaPrestamo = Format(RCab!dFecDesCofide, "YYYYMMDD") '"20091109"
        sMontoPrestamo = FillNum(sDecimal(RCab!nImporteDesem), 14, "0") '"00002000000.00"
        sSaldoACofide = FillNum(sDecimal(RCab!nSaldoPorPagar), 14, "0") '"00001000000.00"
        If RCab!nSaldoPorJustificar < 0 Then
            lnSigno = "-"
        End If
        sPenJustificar = lnSigno & FillNum(sDecimal(Abs(RCab!nSaldoPorJustificar)), 13, "0") '"000000-2000.00"
        sTotalImporteSubPrestamos = FillNum(sDecimal(RCab!nTotalSubPres), 14, "0") '"00003000000.00"
        sTotalActivosFijos = FillNum(sDecimal(RCab!nTotalActiFijo), 14, "0") '"00001083250.00"
        sTotalCapitalTraba = FillNum(sDecimal(RCab!nTotalCapTra), 14, "0") '"00000918750.00"
        sTotalImpoSaldSubPres = FillNum(sDecimal(RCab!nTotalSSubPr), 14, "0") '"00002002000.00"
        sNroSubPrestamos = FillNum(RCab!nCantiReg, 10, "0") '"0000002219"
        CadTemp = sCodigoInst + sFechaProceso + sNroPaquete + sIndicaTI + sMoneda + sCodiOpeCofide + sFechaPrestamo + sMontoPrestamo + sSaldoACofide + sPenJustificar + sTotalImporteSubPrestamos + sTotalActivosFijos + sTotalCapitalTraba + sTotalImpoSaldSubPres + sNroSubPrestamos + oImpresora.gPrnSaltoLinea
        RCab.MoveNext
      Loop
   End If
   RSClose RCab
   sCad = ""
   ArcSal = FreeFile
   Open psArchivoAGrabarC For Output As ArcSal
   
    If CadTemp <> "" Then
       Print #1, sCad; CadTemp
    End If

    Close ArcSal
   CadTemp = ""
   Set RDet = oDCajas.ObtenerFONDEMIDET(txtFechaDel.Text, txtPaquete.Text)
   
    If Not RDet.BOF And Not RDet.EOF Then
'      oBarra.Max = RDet.RecordCount
      Do While Not RDet.EOF
        DoEvents
         
         sCliente = RDet!cNombre + Space(40)
         sDocumentoDNI = FillNum(RDet!cDNI, 8, "0")
         sDocumentoRUC = FillNum(RDet!cRuc, 11, "0")
         sTipoPersona = RDet!cTipPer
         CadTemp = CadTemp + sCodigoInst + sFechaProceso + sNroPaquete + sIndicaTI + sMoneda
         CadTemp = CadTemp + FillNum(RDet!cPersCod, 15, "0") + Mid(sCliente, 1, 40)
         CadTemp = CadTemp + sTipoPersona
         CadTemp = CadTemp + Mid(sDocumentoDNI, 1, 8)
         CadTemp = CadTemp + Mid(sDocumentoRUC, 1, 11) & RDet!cGenero
         CadTemp = CadTemp + FillNum(RDet!cCtaCod, 20, "0")
         CadTemp = CadTemp + Format(RDet!dFecDese, "YYYYMMDD")
         CadTemp = CadTemp + FillNum(sDecimal(CStr(RDet!nImpSPre)), 12, "0")
         CadTemp = CadTemp + IIf(RDet!nActivoFijo > 0, FillNum(sDecimal(CStr(RDet!nActivoFijo)), 12, "0"), FillNum("0.00", 12, "0"))
         CadTemp = CadTemp + IIf(RDet!nCapitalTra > 0, FillNum(sDecimal(CStr(RDet!nCapitalTra)), 12, "0"), FillNum("0.00", 12, "0"))


         CadTemp = CadTemp + FillNum(sDecimal(CStr(RDet!nSaldoCapit)), 12, "0")
         nPlazo = RDet!cFrecPago
         If RDet!cFrecPago = 0 Then
            nPlazo = 30
         End If
         
         If nPlazo = 7 Then
            cPlazo = "07"
         ElseIf nPlazo = 30 Then
            cPlazo = "01"
         ElseIf nPlazo = 90 Then
            cPlazo = "03"
         End If
        CadTemp = CadTemp + cPlazo
        
        cPlazoTotal = Round(((RDet!cPlazoTotal) / 30), 0)
        CadTemp = CadTemp + FillNum(CStr(cPlazoTotal), 3, "0")
        CadTemp = CadTemp + FillNum(CStr(Round((RDet!cPlazoGraci / 30), 0)), 3, "0")
        CadTemp = CadTemp + FillNum(sDecimal(CStr(RDet!nTEA)), 6, "0")
        CadTemp = CadTemp + FillNum(CStr(RDet!nNroCuota), 3, "0")
        CadTemp = CadTemp + IIf(RDet!nDiasMora <= 0, "000", IIf(Len(RDet!nDiasMora) = 1, "00" + CStr(RDet!nDiasMora), IIf(Len(RDet!nDiasMora) = 2, "0" + CStr(RDet!nDiasMora), CStr(RDet!nDiasMora)))) & CStr(CInt(Mid(RDet!nCalifiSBS, 1, 1)) + 1)
        cSector = RDet!cCodigoSec
        CadTemp = CadTemp + cSector + Right(RDet!cCodigoAct, 4) + RDet!cCodUbiGeo + "U"
        
        CadTemp = CadTemp + FillNum(RDet!cCodigoSBS, 10, "0") + oImpresora.gPrnSaltoLinea
'        If Len(CadTemp) >= 1000 Then
'           CadRep = CadRep & CadTemp
'           CadTemp = ""
'        End If
        
        ContBarra = ContBarra + 1
        'oBarra.Progress RDet.Bookmark, "BCIENT", "Generando Formato TXT", "Procesando...", vbBlue
        RDet.MoveNext
      Loop
   End If
   RSClose RDet
   CadRep = CadRep & CadTemp
   'oBarra.CloseForm Me
   Set oDCajas = Nothing
   MousePointer = 0
   
   sCad = ""
   ArcSal = FreeFile
   Open psArchivoAGrabarD For Output As ArcSal
   
   If CadRep <> "" Then
       Print #1, sCad; CadRep
   End If
'   If CadRep <> "" Then
'       Print #1, sCad; CadRep
'   End If
   Close ArcSal
   CadTemp = ""
    If Len(Trim(CadTemp)) > 0 Then
     CadRep = CadRep & CadTemp
     CadTemp = ""
   End If
   EnviaPrevio CadRep, "SUCAVE: BCIENT", gnLinPage, False
   CadRep = ""
End Sub
Private Function sDecimal(ByRef cNumero As String) As String
    If Left(Right(cNumero, 3), 1) = "." Then
        cNumero = cNumero
    ElseIf Mid(Right(cNumero, 3), 2, 1) = "." Then
        cNumero = cNumero & "0"
    Else
        cNumero = cNumero & ".00"
    End If
sDecimal = cNumero
End Function

Private Sub cmdReporte_Click()
    Call Mostrar_Cabecera
    Call Reporte_Fonddemi
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdGenerar_Click()
    Dim oDCajas As DCaja_Adeudados
    Set oDCajas = New DCaja_Adeudados

    Call oDCajas.RegistrarFONDEMIDET(txtFechaDel.Text, txtFechaIni.Text, txtFechaFin.Text, txtPaquete.Text, Trim(Right(cboMoneda.Text, 4)), Right(cboTipInf.Text, 1))
    Call oDCajas.RegistrarFONDEMICAB(txtFechaDel.Text, txtFechaIni.Text, txtFechaFin.Text, txtPaquete.Text, Trim(Right(cboMoneda.Text, 4)), CDbl(txtMontoAprobado.Text), CDbl(txtSaldoTotal.Text), Right(cboTipInf.Text, 1))
    Set oDCajas = Nothing
    Call Mostrar_Cabecera
    Call Reporte_Fonddemi
End Sub

Private Sub Mostrar_Cabecera()
    Dim oDCajas As DCaja_Adeudados
    Set oDCajas = New DCaja_Adeudados
    Dim oRs As ADODB.Recordset
    Dim i As Integer
    Set oDCajas = New DCaja_Adeudados
    Set oRs = New ADODB.Recordset
    Set oRs = oDCajas.ObtenerFONDEMICAB_SP
    
    LimpiaFlex FECab
    i = 1
    Do While Not oRs.EOF
        FECab.AdicionaFila
        FECab.TextMatrix(oRs.Bookmark, 1) = Format(oRs!dFecha, "YYYY/MM/DD")
        FECab.TextMatrix(oRs.Bookmark, 2) = oRs!cNroPaq
        FECab.TextMatrix(oRs.Bookmark, 3) = oRs!cIndInfo
        FECab.TextMatrix(oRs.Bookmark, 4) = Format(IIf(IsNull(oRs!dFechaIni), oRs!dFecha, oRs!dFechaIni), "YYYY/MM/DD")
        FECab.TextMatrix(oRs.Bookmark, 5) = Format(IIf(IsNull(oRs!dFechaFin), oRs!dFecha, oRs!dFechaFin), "YYYY/MM/DD")
        FECab.TextMatrix(oRs.Bookmark, 6) = IIf(IsNull(oRs!nCantiReg), 0, oRs!nCantiReg)
        FECab.TextMatrix(oRs.Bookmark, 7) = IIf(IsNull(oRs!nImporteDesem), 0, oRs!nImporteDesem)
        FECab.TextMatrix(oRs.Bookmark, 8) = IIf(IsNull(oRs!nSaldoPorPagar), 0, oRs!nSaldoPorPagar)
        FECab.TextMatrix(oRs.Bookmark, 9) = Format(IIf(IsNull(oRs!dFecDesCofide), oRs!dFecha, oRs!dFecDesCofide), "YYYYMMDD")
        FECab.TextMatrix(oRs.Bookmark, 10) = oRs!cMoneda
        FECab.TextMatrix(oRs.Bookmark, 11) = IIf(oRs!nestado = 1, "Cerrado", "Pendiente")
        dFechaIni = DateAdd("D", 1, Format(IIf(IsNull(oRs!dFechaFin), oRs!dFecha, oRs!dFechaFin), "DD/MM/YYYY"))
        oRs.MoveNext
    Loop
    Set oDCajas = Nothing
End Sub

Private Sub FECab_Click()
    Dim nPosi As Integer

    nPosi = FECab.Row
    txtFechaDel.Text = Format(CDate(FECab.TextMatrix(nPosi, 1)), "DD/MM/YYYY")
    txtMontoAprobado.Text = FECab.TextMatrix(nPosi, 7)
    txtSaldoTotal.Text = FECab.TextMatrix(nPosi, 8)
   
    cboDesembolso.ListIndex = IndiceListaCombo(cboDesembolso, FECab.TextMatrix(nPosi, 9))
    Call cboDesembolso_Click
    cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, Right(Trim(FECab.TextMatrix(nPosi, 10)), 1))
    cboTipInf.ListIndex = IndiceListaCombo(cboTipInf, Right(Trim(FECab.TextMatrix(nPosi, 3)), 1))
    txtFechaIni.Text = Format(FECab.TextMatrix(nPosi, 4), "DD/MM/YYYY")
    txtFechaFin.Text = Format(FECab.TextMatrix(nPosi, 5), "DD/MM/YYYY")
End Sub

Private Sub Form_Load()
    Call Mostrar_Cabecera
    Call FechaDesembolso
    Call Moneda
    Call TipInf
    txtFechaIni.Text = dFechaIni
    txtFechaFin.Text = dFechaIni
End Sub

Private Sub FechaDesembolso()
    Dim oDCajas As DCaja_Adeudados
    Set oDCajas = New DCaja_Adeudados
    Dim oRs As ADODB.Recordset
    Set oRs = oDCajas.ObtenerPaqueteFONDEMI
    
    If (oRs.BOF Or oRs.EOF) Then
        Exit Sub
    End If
    Do While Not (oRs.EOF)
        cboDesembolso.AddItem oRs!dFechaPaquete & Space(500) & Format(oRs!dFechaPaquete, "YYYYMMDD") & Space(10) & Trim(oRs!cNumerPaquete) & Space(10) & Trim(oRs!cNumerDesemb)
        oRs.MoveNext
    Loop
End Sub

Private Sub Moneda()
    cboMoneda.AddItem "Soles" & Space(500) & "1"
End Sub
Private Sub TipInf()
    cboTipInf.AddItem "Justificacion" & Space(500) & "J"
    cboTipInf.AddItem "Situacion" & Space(500) & "S"
End Sub


