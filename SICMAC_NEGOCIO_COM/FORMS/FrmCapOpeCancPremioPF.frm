VERSION 5.00
Begin VB.Form FrmCapOpeCancPremioPF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobro de Premio por Cancelacion de Plazo Fijo"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin SICMACT.FlexEdit Flex 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   2143
      Cols0           =   3
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Cantidad-Descripcion-Monto"
      EncabezadosAnchos=   "1000-4000-1000"
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
      ColumnasAEditar =   "X-X-X"
      ListaControles  =   "0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "R-L-R"
      FormatosEdit    =   "3-0-2"
      TextArray0      =   "Cantidad"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   1005
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4815
      TabIndex        =   2
      Top             =   1890
      Width           =   1455
   End
   Begin VB.Label LBLMONTO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label LblET 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MONTO TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PLAZO FIJO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label LblPF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "FrmCapOpeCancPremioPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub
Sub Inicion(ByVal psCtaCod As String, ByVal psRs As ADODB.Recordset, ByRef pnMonto As Currency)
    Call CargaPremios(psCtaCod, psRs, pnMonto)
    Me.Show 1
End Sub

Sub CargaPremios(ByVal psCtaCod As String, ByVal psRs As ADODB.Recordset, ByRef pnMonto As Currency)
    Dim nSuma As Currency
    Flex.Clear
    Flex.FormaCabecera
    psRs.MoveFirst
    nSuma = 0
    While Not psRs.EOF
        Flex.AdicionaFila
        Flex.TextMatrix(Me.Flex.Rows - 1, 0) = psRs!nCantidad
        Flex.TextMatrix(Me.Flex.Rows - 1, 1) = psRs!cDescripcion
        Flex.TextMatrix(Me.Flex.Rows - 1, 2) = Format(psRs!nMontoPremio, "#.00")
        nSuma = nSuma + psRs!nMontoPremio
        psRs.MoveNext
    Wend
    Me.LBLMONTO = Format(nSuma, "#.00")
    Me.LblPF.Caption = psCtaCod
    If Mid(psCtaCod, 9, 1) = "1" Then
        pnMonto = nSuma
        Me.LblET.Caption = "MONTO TOTAL S/."
        Me.LblET.BackColor = &HC0FFFF
    Else
        Me.LblET.BackColor = &HC0FFC0
        Me.LblET.Caption = "MONTO TOTAL $ "
        Dim TC As COMDConstSistema.NCOMTipoCambio
        Set TC = New COMDConstSistema.NCOMTipoCambio
        pnMonto = (nSuma / TC.EmiteTipoCambio(gdFecSis, TipoCambio.TCCompra))
        Set TC = Nothing
    End If
End Sub

