VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCapServGenReporte 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "frmCapServGenReporte.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5220
      Width           =   1035
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5220
      Width           =   1035
   End
   Begin VB.Frame fraCuentas 
      Caption         =   "Parámetros :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1320
      Left            =   90
      TabIndex        =   7
      Top             =   90
      Width           =   4695
      Begin SICMACT.FlexEdit grdCuentas 
         Height          =   1005
         Left            =   135
         TabIndex        =   8
         Top             =   225
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   1773
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Descripcion-Valor"
         EncabezadosAnchos=   "350-1600-1900"
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Fecha de Proceso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1320
      Left            =   4860
      TabIndex        =   3
      Top             =   90
      Width           =   4155
      Begin VB.CheckBox chkReciboYaAbono 
         Alignment       =   1  'Right Justify
         Caption         =   "Recibos Ya Abonados"
         Height          =   195
         Left            =   1620
         TabIndex        =   11
         Top             =   990
         Width           =   2220
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   2970
         Picture         =   "frmCapServGenReporte.frx":030A
         TabIndex        =   5
         Top             =   405
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker txtInicio 
         Height          =   330
         Left            =   945
         TabIndex        =   4
         Top             =   240
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   66977793
         CurrentDate     =   37636
      End
      Begin MSComCtl2.DTPicker txtFin 
         Height          =   330
         Left            =   945
         TabIndex        =   12
         Top             =   630
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   66977793
         CurrentDate     =   37636
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al :"
         Height          =   195
         Left            =   225
         TabIndex        =   13
         Top             =   705
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del :"
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   315
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdDistribuir 
      Caption         =   "&Distribuir Fondos"
      Height          =   375
      Left            =   6210
      TabIndex        =   2
      Top             =   5220
      Width           =   1710
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Recibos Cobrados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3660
      Left            =   90
      TabIndex        =   1
      Top             =   1485
      Width           =   8925
      Begin SICMACT.FlexEdit grdRecibos 
         Height          =   2850
         Left            =   90
         TabIndex        =   10
         Top             =   270
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   5027
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Fecha-Recibo-Codigo-Monto-Agencia-Usuario-nMovNro"
         EncabezadosAnchos=   "350-1900-1200-1200-1200-1700-800-0"
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-2-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comisón(S/.) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   5850
         TabIndex        =   19
         Top             =   3195
         Width           =   1590
      End
      Begin VB.Label lblComision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   7425
         TabIndex        =   18
         Top             =   3195
         Width           =   1200
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cobrado (S/.) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3015
         TabIndex        =   17
         Top             =   3195
         Width           =   1590
      End
      Begin VB.Label lblCobrado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   4590
         TabIndex        =   16
         Top             =   3195
         Width           =   1200
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "# Recibos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   135
         TabIndex        =   15
         Top             =   3195
         Width           =   1590
      End
      Begin VB.Label lblNumRecibos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1710
         TabIndex        =   14
         Top             =   3195
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5220
      Width           =   1035
   End
End
Attribute VB_Name = "frmCapServGenReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nServicio As CaptacInstServicios
Dim nTipoComision As CapServTipoComision
Dim nComision As Double
Dim sCuenta As String

Private Sub CalculaTotales()
lblNumRecibos.Caption = Format$(grdRecibos.Rows - 1, "#,##0")
lblCobrado.Caption = Format$(grdRecibos.SumaRow(4), "#,##0.00")
If nTipoComision = gCapServTpoComMontoRecibo Then
    lblComision.Caption = Format$(nComision * CDbl(lblNumRecibos.Caption), "#,##0.00")
ElseIf nTipoComision = gCapServTpoComPorcentaje Then
    lblComision.Caption = Format$(nComision * CDbl(lblCobrado.Caption), "#,##0.00")
End If
End Sub

Public Sub Inicia(ByVal nServ As CaptacInstServicios)
Dim oServ As NCapServicios
Dim rsServ As Recordset
Dim sTipoComision As String
Dim nFila As Long

nServicio = nServ
Set oServ = New NCapServicios
Select Case nServicio
    Case gCapServSedalib
        Set rsServ = oServ.GetServicioParametros(gCapServSedalib)
        Me.Caption = "Captaciones - Servicios - Reportes - SEDALIB"
    Case gCapServHidrandina
        Set rsServ = oServ.GetServicioParametros(gCapServHidrandina)
        Me.Caption = "Captaciones - Servicios - Reportes - HIDRANDINA"
    Case gCapServEdelnor
        Set rsServ = oServ.GetServicioParametros(gCapServEdelnor)
        Me.Caption = "Captaciones - Servicios - Reportes - EDELNOR"
End Select
nTipoComision = rsServ("nTipoComision")
If nTipoComision = gCapServTpoComMontoRecibo Then
    sTipoComision = "(x Recibo)"
ElseIf nTipoComision = gCapServTpoComPorcentaje Then
    sTipoComision = "(%)"
End If
grdCuentas.AdicionaFila
sCuenta = rsServ("cCtaCodAbono")
grdCuentas.TextMatrix(1, 1) = "Cuenta :"
grdCuentas.TextMatrix(1, 2) = sCuenta
grdCuentas.AdicionaFila
grdCuentas.TextMatrix(2, 1) = "Comision " & sTipoComision & " :"
nComision = rsServ("nComision")
grdCuentas.TextMatrix(2, 2) = nComision
rsServ.Close
Set rsServ = Nothing
cmdDistribuir.Enabled = False
cmdImprimir.Enabled = False
cmdCancelar.Enabled = False
Me.Show 1
End Sub

Private Sub chkReciboYaAbono_Click()
If chkReciboYaAbono.value = 1 Then
    cmdDistribuir.Enabled = False
Else
    cmdDistribuir.Enabled = True
End If
End Sub

Private Sub cmdBuscar_Click()
Dim clsServ As NCapServicios
Dim dInicio As Date, dFin As Date
Dim bRecYaAbo As Boolean
Dim rsServ As Recordset

dInicio = CDate(txtInicio.value)
dFin = CDate(txtFin.value)

bRecYaAbo = IIf(chkReciboYaAbono.value = 1, True, False)

Set clsServ = New NCapServicios
Select Case nServicio
    Case gCapServSedalib
        Set rsServ = clsServ.GetMovServicios(gServCobSedalib, dInicio, dFin, bRecYaAbo)
    Case gCapServHidrandina
        Set rsServ = clsServ.GetMovServicios(gServCobHidrandina, dInicio, dFin, bRecYaAbo)
    Case gCapServEdelnor
        Set rsServ = clsServ.GetMovServicios(gServCobEdelnor, dInicio, dFin, bRecYaAbo)
End Select
Set clsServ = Nothing
If Not (rsServ.EOF And rsServ.BOF) Then
    Set grdRecibos.Recordset = rsServ
    grdRecibos.FormateaColumnas
    CalculaTotales
    fraFecha.Enabled = False
    If Not bRecYaAbo Then cmdDistribuir.Enabled = True
    cmdImprimir.Enabled = True
    cmdCancelar.Enabled = True
Else
    MsgBox "No se encontraron recibos en la fecha indicada", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdCancelar_Click()
txtInicio.value = gdFecSis
txtFin.value = gdFecSis
cmdDistribuir.Enabled = False
cmdCancelar.Enabled = False
cmdImprimir.Enabled = False
fraFecha.Enabled = True
chkReciboYaAbono.value = 0
txtInicio.SetFocus
grdRecibos.Clear
grdRecibos.Rows = 2
grdRecibos.FormaCabecera
lblNumRecibos.Caption = "0"
lblCobrado.Caption = "0.00"
lblComision.Caption = "0.00"
End Sub

Private Sub cmdDistribuir_Click()

If MsgBox("¿Desea Grabar la Operación de distribución?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim nMontoAbono As Double, nMontoCargo As Double
    Dim oServ As NCapServicios
    Dim rsServ As Recordset
    Dim sGlosa As String
    nMontoAbono = CDbl(lblCobrado)
    nMontoCargo = CDbl(lblComision)
    Set rsServ = grdRecibos.GetRsNew()
    Set oServ = New NCapServicios
    sGlosa = "Por cobranza del " & Format$(txtInicio.value, "dd/mm/yyyy") & " Al " & Format$(txtFin.value, "dd/mm/yyyy")
    If nServicio = gCapServSedalib Then
        oServ.GrabaAbonoServicios rsServ, nMontoAbono, nMontoCargo, sCuenta, gAhoDepPagServSedalib, _
                gAhoRetComServSEDALIB, sGlosa, gdFecSis, gsCodUser, gsCodAge, gsNomCmac, gsNomAge, sLpt
    ElseIf nServicio = gCapServHidrandina Then
        oServ.GrabaAbonoServicios rsServ, nMontoAbono, nMontoCargo, sCuenta, gAhoDepPagServHidrandina, _
                gAhoRetComServHidrandina, sGlosa, gdFecSis, gsCodUser, gsCodAge, gsNomCmac, gsNomAge, sLpt
    ElseIf nServicio = gCapServEdelnor Then
        oServ.GrabaAbonoServicios rsServ, nMontoAbono, nMontoCargo, sCuenta, gAhoDepPagServEdelnor, _
                gAhoRetComServEDELNOR, sGlosa, gdFecSis, gsCodUser, gsCodAge, gsNomCmac, gsNomAge, sLpt
    End If
    Set oServ = Nothing
    cmdCancelar_Click
End If
End Sub

Private Sub cmdImprimir_Click()
Dim clsServ As NCapServicios
Dim clsPrev As previo.clsprevio
Dim sCad As String
Dim dFechaCobro As Date
Dim rsServ As Recordset
Dim dInicio As Date, dFin As Date
dInicio = CDate(txtInicio.value)
dFin = CDate(txtFin.value)

Set rsServ = grdRecibos.GetRsNew()
Set clsServ = New NCapServicios
sCad = clsServ.GeneraReporteServicios(rsServ, nServicio, CLng(lblNumRecibos), CDbl(lblCobrado), _
            nComision, CDbl(lblComision), nTipoComision, dInicio, dFin, gdFecSis, gsNomCmac)
Set clsServ = Nothing
If sCad <> "" Then
    Set clsPrev = New previo.clsprevio
    'ALPA 20200202****************************
    'clsPrev.Show sCad, "Captaciones - Servicios - Reporte Cobranza", True
    clsPrev.Show sCad, "Captaciones - Servicios - Reporte Cobranza", True, , gImpresora
    Set clsPrev = Nothing
Else
    MsgBox "No se encontraron datos para la fecha indicada", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtInicio.value = gdFecSis
Me.Icon = LoadPicture(App.path & gsRutaIcono)
txtFin.value = gdFecSis
End Sub
