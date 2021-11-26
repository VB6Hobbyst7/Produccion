VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdeuAlertaVencimientoPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alerta de Vencimiento de Pago"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15870
   Icon            =   "frmAdeuAlertaVencimientoPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   15870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   13800
      TabIndex        =   3
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "Reporte"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1695
   End
   Begin TabDlg.SSTab TabVencimiento 
      Height          =   4635
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   8176
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabMaxWidth     =   3528
      TabCaption(0)   =   "Vencimientos"
      TabPicture(0)   =   "frmAdeuAlertaVencimientoPago.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Cuotas Vencida"
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
         Height          =   4095
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   15255
         Begin Sicmact.FlexEdit FELista 
            Height          =   3495
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   14895
            _ExtentX        =   26273
            _ExtentY        =   6165
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Fondo-Linea-Tasa-Capital-Cuota Vencida-Fecha Venc.-Moneda"
            EncabezadosAnchos=   "1200-2000-4000-1200-2000-1200-1200-2000"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-R-R-R-C-L"
            FormatosEdit    =   "0-0-0-0-0-0-5-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   1200
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
   End
End
Attribute VB_Name = "frmAdeuAlertaVencimientoPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fbSalir As Boolean

Private Sub cmdReporte_Click()
Dim oTipCambio As nTipoCambio
Dim lnTipCambio As Currency
Set oTipCambio = New nTipoCambio
    lnTipCambio = Format(oTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "#0.000")
Set oTipCambio = Nothing

Call MostrarReporteAlertaVencimiento(gdFecSis, lnTipCambio)
End Sub
Private Sub MostrarReporteAlertaVencimiento(ByVal pdFechaProc As Date, ByVal pnTipCamb As Double)
   Dim objLinea As DLineaCreditoV2
    
Dim R As ADODB.Recordset
Dim lMatCabecera As Variant
Dim lsMensaje As String
Dim lsNombreArchivo As String

    lsNombreArchivo = "ReporteAlertaVencimiento"
    
    ReDim lMatCabecera(6, 0)
    Dim objAde As DAdeudCal
    Set objAde = New DAdeudCal
    Set R = objAde.ObtenerAlertaVencimientoPago(gdFecSis)
    Set objLinea = Nothing
    If Not R Is Nothing Then
        Call GeneraReporteEnArchivoExcelInicio(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Priododidad", "", lsNombreArchivo, lMatCabecera, R, 2, , , True, True)
    Else
        MsgBox lsMensaje, vbInformation, "AVISO"
    End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub ListarAlertaVencimiento()
Dim objAde As DAdeudCal
Dim lbCargo As Boolean
Dim objRs As ADODB.Recordset
Set objRs = New ADODB.Recordset
Set objAde = New DAdeudCal
lbCargo = objAde.obtenerSitieneElCargo(gsCodPersUser)
Set objAde = Nothing
fbSalir = False
If lbCargo = True Then
Set objAde = New DAdeudCal
Set objRs = objAde.ObtenerAlertaVencimientoPago(gdFecSis)



    Set objAde = Nothing
     Dim nNumero As Integer
        nNumero = 0
        FormateaFlex FELista
        If Not (objRs.BOF Or objRs.EOF) Then
        Do While Not objRs.EOF
            FELista.AdicionaFila
            nNumero = nNumero + 1
            FELista.TextMatrix(nNumero, 0) = nNumero
            FELista.TextMatrix(nNumero, 1) = objRs!cPersNombre
            FELista.TextMatrix(nNumero, 2) = objRs!cCtaIFDesc
            FELista.TextMatrix(nNumero, 3) = objRs!nCtaIFIntValor
            FELista.TextMatrix(nNumero, 4) = objRs!nCapital
            FELista.TextMatrix(nNumero, 5) = objRs!nCuotaVencida
            FELista.TextMatrix(nNumero, 6) = Format(objRs!dVencimiento, "YYYY/MM/DD")
            FELista.TextMatrix(nNumero, 7) = IIf(objRs!cMonedaPago = 1, "SOLES", "DOLARES")
            
            objRs.MoveNext
        Loop
        Set objRs = Nothing
        Else
            fbSalir = True
        End If
    Else
        fbSalir = True
    End If
End Sub

Private Sub Form_Activate()
    If fbSalir Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Screen.MousePointer = 0
Call ListarAlertaVencimiento
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = 11
End Sub
