VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmServDisfondos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "frmServDisFondos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCuentas 
      Caption         =   "Monto Recaudado Distribuido :"
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
      Height          =   2550
      Left            =   150
      TabIndex        =   16
      Top             =   810
      Width           =   7290
      Begin SICMACT.FlexEdit FeMontoRec 
         Height          =   1875
         Left            =   75
         TabIndex        =   17
         Top             =   195
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   3307
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Tipo-Concepto-Importe"
         EncabezadosAnchos=   "350-2800-2500-1200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         ColumnasAEditar =   "X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R"
         FormatosEdit    =   "0-0-0-4"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comis."
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
         Left            =   1620
         TabIndex        =   27
         Top             =   2160
         Width           =   630
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
         Left            =   2235
         TabIndex        =   26
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto"
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
         Index           =   0
         Left            =   60
         TabIndex        =   25
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label lblImporte 
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
         Left            =   615
         TabIndex        =   24
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G.Adm."
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
         Index           =   1
         Left            =   3030
         TabIndex        =   23
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label lblGastosA 
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
         Left            =   3675
         TabIndex        =   22
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Costas"
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
         Index           =   2
         Left            =   4470
         TabIndex        =   21
         Top             =   2160
         Width           =   630
      End
      Begin VB.Label lblcostas 
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
         Left            =   5085
         TabIndex        =   20
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D.Emi."
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
         Index           =   3
         Left            =   5895
         TabIndex        =   19
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label lblDemis 
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
         Left            =   6480
         TabIndex        =   18
         Top             =   2160
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Monto  Distribuido por %"
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
      Height          =   2550
      Left            =   165
      TabIndex        =   14
      Top             =   3780
      Width           =   7290
      Begin SICMACT.FlexEdit FeDistrib 
         Height          =   2220
         Left            =   90
         TabIndex        =   15
         Top             =   225
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   3916
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Tipo-Cuenta-Concepto-Importe"
         EncabezadosAnchos=   "400-1000-1700-2000-1200"
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
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-R"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Distribucion Fondos"
      Height          =   375
      Left            =   5820
      TabIndex        =   13
      Top             =   3420
      Width           =   1515
   End
   Begin VB.CommandButton cmdDistribuir 
      Caption         =   "&Distribuir Fondos"
      Height          =   375
      Left            =   4755
      TabIndex        =   12
      Top             =   6435
      Width           =   1620
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6435
      Width           =   1035
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3705
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6435
      Width           =   1035
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
      Height          =   720
      Left            =   150
      TabIndex        =   2
      Top             =   90
      Width           =   6150
      Begin VB.CheckBox chkReciboYaAbono 
         Alignment       =   1  'Right Justify
         Caption         =   "Recibos Ya Abonados"
         Height          =   195
         Left            =   3675
         TabIndex        =   8
         Top             =   735
         Width           =   2220
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   4950
         Picture         =   "frmServDisFondos.frx":030A
         TabIndex        =   4
         Top             =   225
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker txtInicio 
         Height          =   330
         Left            =   615
         TabIndex        =   3
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
         Format          =   49414145
         CurrentDate     =   37636
      End
      Begin MSComCtl2.DTPicker txtFin 
         Height          =   330
         Left            =   2865
         TabIndex        =   9
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
         Format          =   49414145
         CurrentDate     =   37636
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al :"
         Height          =   195
         Left            =   2625
         TabIndex        =   10
         Top             =   315
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del :"
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   315
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6435
      Width           =   1035
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Movimientos Efectuados"
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
      Height          =   1410
      Left            =   780
      TabIndex        =   1
      Top             =   1200
      Width           =   5670
      Begin SICMACT.FlexEdit FeRecibos 
         Height          =   1860
         Left            =   90
         TabIndex        =   7
         Top             =   270
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   3281
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-NroMov-Usua-nMonto-cMovDesc-Fecha"
         EncabezadosAnchos=   "350-1000-800-1000-2200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-R-R-L-R"
         FormatosEdit    =   "0-0-0-4-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmServDisfondos"
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
Dim oServ As NCapServicios
Set oServ = New NCapServicios

lblImporte.Caption = oServ.GetSatSuma("1") 'Suma de Importes

End Sub

Public Sub Inicia(ByVal nServ As CaptacInstServicios)
'Dim oServ As NCapServicios
'Dim rsServ As Recordset
'Dim sTipoComision As String
'Dim nFila As Long
'
'nServicio = nServ
'Set oServ = New NCapServicios
'Set rsServ = oServ.GetSatMontoFecha(txtInicio.value, txtFin.value)
'Set rsServ = oServ.GetServicioParametros(gCapServSATTInfraccion)
'Me.Caption = "Captaciones - Servicios - Reportes - SAT - INFRACCIONES"
'
'Set oServ = New NCapServicios
'Select Case nServicio
'       Case gCapServSedalib
'        Set rsServ = oServ.GetServicioParametros(gCapServSedalib)
'        Me.Caption = "Captaciones - Servicios - Reportes - SEDALIB"
'    Case gCapServHidrandina
'        Set rsServ = oServ.GetServicioParametros(gCapServHidrandina)
'        Me.Caption = "Captaciones - Servicios - Reportes - HIDRANDINA"
'    Case gCapServEdelnor
'        Set rsServ = oServ.GetServicioParametros(gCapServEdelnor)
'        Me.Caption = "Captaciones - Servicios - Reportes - EDELNOR"
'    Case gCapServSATTInfraccion
'        Set rsServ = oServ.GetServicioParametros(gCapServSATTInfraccion)
'        Me.Caption = "Captaciones - Servicios - Reportes - SAT - INFRACCIONES"
'
'End Select
'nTipoComision = rsServ("nTipoComision")
'If nTipoComision = gCapServTpoComMontoRecibo Then
'    sTipoComision = "(x Recibo)"
'ElseIf nTipoComision = gCapServTpoComPorcentaje Then
'    sTipoComision = "(%)"
'End If
'grdCuentas.AdicionaFila
'sCuenta = rsServ("cCtaCodAbono")
'grdCuentas.TextMatrix(1, 1) = "Cuenta :"
'grdCuentas.TextMatrix(1, 2) = sCuenta
'grdCuentas.AdicionaFila
'grdCuentas.TextMatrix(2, 1) = "Comision " & sTipoComision & " :"
'nComision = rsServ("nComision")
'grdCuentas.TextMatrix(2, 2) = nComision
'rsServ.Close
'Set rsServ = Nothing
'cmdDistribuir.Enabled = False
'cmdImprimir.Enabled = False
'cmdCancelar.Enabled = False
'Me.Show 1
End Sub

Private Sub chkReciboYaAbono_Click()
If chkReciboYaAbono.value = 1 Then
    cmdDistribuir.Enabled = False
Else
    cmdDistribuir.Enabled = True
End If
End Sub
'MODIFICACION EFECTUADA POR CMCPL - CRSF 10/06
Private Sub cmdBuscar_Click()
Dim clsServ As NCapServicios
Dim dInicio As Date, dFin As Date
Dim bRecYaAbo As Boolean
Dim rsserv As Recordset
Dim i As Integer
Dim sConcepto As String, sopera As String
'DECLARACION DE VARIABLES PARA SIGNACION

dInicio = CDate(txtInicio.value)
dFin = CDate(txtFin.value)
bRecYaAbo = IIf(chkReciboYaAbono.value = 1, True, False)
Set clsServ = New NCapServicios
'LISTA LOS MOVIMIENTOS -TODOS
'Set rsServ = clsServ.GetSatMontoFecha(txtInicio.value, txtFin.value)
'Set clsServ = Nothing
'If rsServ.EOF And rsServ.BOF Then
'    MsgBox "Datos No encontrados, Pagos de SAT no efectuados este dia", vbInformation, "Aviso"
'Else
'    FeRecibos.Clear
'    FeRecibos.Rows = 2
'    FeRecibos.FormaCabecera
'    Do While Not rsServ.EOF
'        FeRecibos.AdicionaFila , , True
'        FeRecibos.TextMatrix(FeRecibos.Rows - 1, 1) = rsServ!nMovNro
'        FeRecibos.TextMatrix(FeRecibos.Rows - 1, 2) = rsServ!Usua
'        FeRecibos.TextMatrix(FeRecibos.Rows - 1, 3) = Format$(rsServ!nMonto, "#,##0") 'rsServ!nMonto
'        FeRecibos.TextMatrix(FeRecibos.Rows - 1, 4) = rsServ!cMovDesc
'        FeRecibos.TextMatrix(FeRecibos.Rows - 1, 5) = rsServ!Fecha
'        rsServ.MoveNext
'    Loop
' ********************************************************************************
' SUMA POR TIPO LOS DETALLES
Set clsServ = New NCapServicios
Set rsserv = clsServ.GetSatMontoFechaDis(txtInicio.value, txtFin.value)
Set clsServ = Nothing
    FeMontoRec.Clear
    FeMontoRec.Rows = 2
    FeMontoRec.FormaCabecera
    Do While Not rsserv.EOF
        FeMontoRec.AdicionaFila , , True
        If rsserv!cOpecod = "300105" Then
            sopera = "1 - Pago de Infraciones"
        Else
            sopera = "2 - Pago de Tributos/Derechos"
        End If
        FeMontoRec.TextMatrix(FeMontoRec.Rows - 1, 1) = sopera
        Select Case rsserv!nPrdConcepto
            Case 1
                        sConcepto = "1 - Monto de Importe"
            Case 2
                        sConcepto = "2 - Monto de Comision"
            Case 3
                        sConcepto = "3 - Gastos Administrativos."
            Case 4
                        sConcepto = "4 - Gastos de Costas"
            Case 5
                        sConcepto = "5 - Derecho Emision"
         End Select
        FeMontoRec.TextMatrix(FeMontoRec.Rows - 1, 2) = sConcepto
        FeMontoRec.TextMatrix(FeMontoRec.Rows - 1, 3) = rsserv!nMonto
        rsserv.MoveNext
    Loop
 CalculaTotales
'End If
' *****************************************************************************
'DISTRIBUCION DE FONDOS
'distribucion_fondos
'Select Case nServicio
'    Case gCapServSedalib
'        Set rsServ = clsServ.GetMovServicios(gServCobSedalib, dInicio, dFin, bRecYaAbo)
'    Case gCapServHidrandina
'        Set rsServ = clsServ.GetMovServicios(gServCobHidrandina, dInicio, dFin, bRecYaAbo)
'    Case gCapServEdelnor
'        Set rsServ = clsServ.GetMovServicios(gServCobEdelnor, dInicio, dFin, bRecYaAbo)
'End Select
'Set clsServ = Nothing
'If Not (rsServ.EOF And rsServ.BOF) Then
'    Set grdRecibos.Recordset = rsServ
'    grdRecibos.FormateaColumnas
'    CalculaTotales
'    fraFecha.Enabled = False
'    If Not bRecYaAbo Then cmdDistribuir.Enabled = True
'    cmdImprimir.Enabled = True
'    cmdCancelar.Enabled = True
'Else
'    MsgBox "No se encontraron recibos en la fecha indicada", vbInformation, "Aviso"
'End If

 End Sub
'NUEVO CMCPL  - CRSF 11/08
Public Sub distribucion_fondos()
Dim Reg As NCapServicios
Dim rsserv As Recordset
    FeDistrib.Clear
    FeDistrib.Rows = 2
    FeDistrib.FormaCabecera
    Set Reg = New NCapServicios
    'ABONO EN TOTAL A LA CUENTA DE MUNI (TOTAL - DERE.EMIS)
    FeDistrib.AdicionaFila
    FeDistrib.TextMatrix(FeDistrib.Rows - 1, 1) = "ABONAR"
    Set rsserv = Reg.GetSatDistribucion("1")
    Set Reg = Nothing
    FeDistrib.TextMatrix(FeDistrib.Rows - 1, 2) = rsserv("cctacodabono")
    FeDistrib.TextMatrix(FeDistrib.Rows - 1, 3) = "Suma Total de  Ingresos"
    
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
lblNumRecibos.Caption = "0"
lblCobrado.Caption = "0.00"
lblComision.Caption = "0.00"
End Sub
 ' MODIFICAR PARA LA DISTRIBUCION DE FONDOS DE LA SAT - CRSF 10/08
Private Sub cmdDistribuir_Click()

If MsgBox("¿Desea Grabar la Operación de distribución?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim nMontoAbono As Double, nMontoCargo As Double
    Dim oServ As NCapServicios
    Dim rsserv As Recordset
    Dim sGlosa As String
    nMontoAbono = CDbl(lblCobrado)
    nMontoCargo = CDbl(lblComision)
    Set rsserv = grdRecibos.GetRsNew()
    Set oServ = New NCapServicios
    sGlosa = "Por cobranza del " & Format$(txtInicio.value, "dd/mm/yyyy") & " Al " & Format$(txtFin.value, "dd/mm/yyyy")
    If nServicio = gCapServSedalib Then
        oServ.GrabaAbonoServicios rsserv, nMontoAbono, nMontoCargo, sCuenta, gAhoDepPagServSedalib, _
                gAhoRetComServSEDALIB, sGlosa, gdFecSis, gsCodUser, gsCodAge, gsNomCmac, gsNomAge, sLpt
    ElseIf nServicio = gCapServHidrandina Then
        oServ.GrabaAbonoServicios rsserv, nMontoAbono, nMontoCargo, sCuenta, gAhoDepPagServHidrandina, _
                gAhoRetComServHidrandina, sGlosa, gdFecSis, gsCodUser, gsCodAge, gsNomCmac, gsNomAge, sLpt
    ElseIf nServicio = gCapServEdelnor Then
        oServ.GrabaAbonoServicios rsserv, nMontoAbono, nMontoCargo, sCuenta, gAhoDepPagServEdelnor, _
                gAhoRetComServEDELNOR, sGlosa, gdFecSis, gsCodUser, gsCodAge, gsNomCmac, gsNomAge, sLpt
    End If
    Set oServ = Nothing
    cmdCancelar_Click
End If
End Sub

Private Sub cmdImprimir_Click()
Dim clsServ As NCapServicios
Dim clsPrev As Previo.clsPrevio
Dim sCad As String
Dim dFechaCobro As Date
Dim rsserv As Recordset
Dim dInicio As Date, dFin As Date
dInicio = CDate(txtInicio.value)
dFin = CDate(txtFin.value)

Set rsserv = grdRecibos.GetRsNew()
Set clsServ = New NCapServicios
sCad = clsServ.GeneraReporteServicios(rsserv, nServicio, CLng(lblNumRecibos), CDbl(lblCobrado), _
            nComision, CDbl(lblComision), nTipoComision, dInicio, dFin, gdFecSis, gsNomCmac)
Set clsServ = Nothing
If sCad <> "" Then
    Set clsPrev = New Previo.clsPrevio
    clsPrev.Show sCad, "Captaciones - Servicios - Reporte Cobranza", True
    Set clsPrev = Nothing
Else
    MsgBox "No se encontraron datos para la fecha indicada", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtInicio.value = gdFecSis
Me.Icon = LoadPicture(App.path & gsRutaIcono)
txtFin.value = gdFecSis
Me.Caption = "Distribución de Fondos - SAT"
End Sub

