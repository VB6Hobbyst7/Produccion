VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDistrExposicion 
   Caption         =   "Distribucion por tipos de exposiciones"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   Icon            =   "frmDistrExposicion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdRegistrar 
      Caption         =   "Registrar"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   8520
      TabIndex        =   5
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.CommandButton CmdMostrar 
         Caption         =   "Mostrar"
         Height          =   495
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin SICMACT.FlexEdit FEDE 
         Height          =   3015
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5318
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Item-Fecha-Descripción-Monto"
         EncabezadosAnchos=   "400-0-1400-4200-2200"
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
         ColumnasAEditar =   "X-X-X-X-4"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-R"
         FormatosEdit    =   "0-0-5-5-4"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin MSMask.MaskEdBox mskPeriodo 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   390
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmDistrExposicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdMostrar_Click()
    Call MostrarExposicion
End Sub
Private Sub MostrarExposicion()
    Dim oColEval As COMNCredito.NCOMColocEval
    Set oColEval = New COMNCredito.NCOMColocEval
    Dim oRs As ADODB.Recordset
    Dim i As Integer
    Set oRs = New ADODB.Recordset
    Set oRs = oColEval.ObtenerExposiciones2A1(mskPeriodo.Text)
    
    LimpiaFlex FEDE
    i = 1
    Do While Not oRs.EOF
        FEDE.AdicionaFila
        FEDE.TextMatrix(oRs.Bookmark, 1) = oRs!nOrdenTCredito
        FEDE.TextMatrix(oRs.Bookmark, 2) = Format(oRs!dFecha, "YYYY/MM/DD")
        FEDE.TextMatrix(oRs.Bookmark, 3) = oRs!cConsDescripcion
        FEDE.TextMatrix(oRs.Bookmark, 4) = Format(IIf(IsNull(oRs!nValor), 0, oRs!nValor), "###,###,###,##0.00")
        oRs.MoveNext
    Loop
    Set oColEval = Nothing
End Sub

Private Sub cmdRegistrar_Click()
    Dim i As Integer
    Dim oColEval As COMNCredito.NCOMColocEval
    Set oColEval = New COMNCredito.NCOMColocEval
    For i = 1 To FEDE.Rows - 1
        Call oColEval.InsertaExposiciones2A1(mskPeriodo.Text, FEDE.TextMatrix(i, 1), FEDE.TextMatrix(i, 4))
    Next i
    Call cmdMostrar_Click
    MsgBox "Los datos se guardaron correctamente", vbApplicationModal
End Sub

