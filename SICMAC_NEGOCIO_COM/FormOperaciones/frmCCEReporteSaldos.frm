VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCCEReporteSaldos 
   BackColor       =   &H80000016&
   Caption         =   "Reporte Saldos"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16905
   Icon            =   "frmCCEReporteSaldos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   16905
   StartUpPosition =   3  'Windows Default
   Begin SICMACT.FlexEdit flxReportSaldo 
      Height          =   2655
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   4683
      Cols0           =   11
      EncabezadosNombres=   $"frmCCEReporteSaldos.frx":030A
      EncabezadosAnchos=   "1200-4200-4200-3200-3200-4200-1200-4200-4200-3200-3200"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "L-C-C-C-C-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "Movimiento"
      ColWidth0       =   1200
      RowHeight0      =   300
   End
   Begin VB.Frame FraInstrumento 
      BackColor       =   &H80000016&
      Caption         =   " Instrumentos "
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
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton cmdReporte 
         BackColor       =   &H80000016&
         Caption         =   "Generar Reporte"
         Height          =   495
         Left            =   5640
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboInstrumento 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txtFechaSaldo 
         Height          =   315
         Left            =   3960
         TabIndex        =   5
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H80000016&
         Caption         =   "Fecha"
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
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblTipoSesion 
         BackColor       =   &H80000016&
         Caption         =   "Sesion"
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
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCargaFecha 
         BackColor       =   &H80000016&
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblCargaHora 
         BackColor       =   &H80000016&
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCCEReporteSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'** Nombre : frmCCEReporteSaldos
'** Descripción : Para Reporte de Saldos
'** Creación : VAPA20170630
'**********************************************************************
Option Explicit
Dim oCCE As COMNCajaGeneral.NCOMCCE
Dim lsAplicacion As String
Dim bLogicoDev As Boolean
Dim bLogicoCon As Boolean
Private Sub cboInstrumento_Click()
lsAplicacion = cboInstrumento.Text
txtFechaSaldo.SetFocus
End Sub
Private Sub optTipoBus_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtFechaSaldo.Visible Then
        txtFechaSaldo.SetFocus
        cmdReporte.Enabled = True
    End If
End If
End Sub
Private Sub CargaDatosflxReportSaldo(ByVal psSesion As String, ByVal pdFecha As Date)
    Dim rsReporte As ADODB.Recordset
    Set rsReporte = oCCE.CCE_ReporteSaldos(psSesion, pdFecha)
    'VAPA20170720
    LimpiaFlex flxReportSaldo
    If Not rsReporte.EOF Then
    Set Me.flxReportSaldo.Recordset = rsReporte
    Else
    MsgBox "No se encuentra reporte de Saldos en esta sesión.", vbInformation, "¡Aviso!"
    Set Me.flxReportSaldo.Recordset = rsReporte
    End If
End Sub
Private Sub cmdReporte_Click()

 Call CargaDatosflxReportSaldo(lsAplicacion, CDate(txtFechaSaldo))
End Sub



Private Sub txtFechaSaldo_GotFocus()
With txtFechaSaldo
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtFechaSaldo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdReporte.Enabled = True
End If
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
    cmdReporte.Enabled = False
    Set oCCE = New COMNCajaGeneral.NCOMCCE
    Set rs = oCCE.CCE_ObtieneAplicacionReporte
    Do While Not rs.EOF
        cboInstrumento.AddItem rs!cCodAplicacion
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub
