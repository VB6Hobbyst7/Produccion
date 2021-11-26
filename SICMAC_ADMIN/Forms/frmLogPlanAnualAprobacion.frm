VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogPlanAnualAprobacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5460
   ClientLeft      =   345
   ClientTop       =   2385
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   11415
   Begin VB.CommandButton cmdAreas 
      Caption         =   "Areas"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton cmdAgencias 
      Caption         =   "Agencias"
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      Top             =   5040
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   7995
      Begin VB.TextBox txtAnio 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   6600
         TabIndex        =   3
         Text            =   "2005"
         Top             =   105
         Width           =   615
      End
      Begin VB.Label lblPlan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plan Anual de Adquisiciones y Contrataciones del Año  "
         BeginProperty Font 
            Name            =   "Helvetica"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   120
         Width           =   6375
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   525
         Left            =   0
         Top             =   0
         Width           =   7395
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1305
      Left            =   120
      TabIndex        =   5
      Top             =   525
      Width           =   11175
      Begin VB.TextBox txtEntidad 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   280
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   300
         Width           =   7395
      End
      Begin VB.TextBox txtRUC 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   280
         Left            =   9180
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   300
         Width           =   1755
      End
      Begin VB.TextBox txtSiglas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   280
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox txtPliego 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   280
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   900
         Width           =   1635
      End
      Begin VB.TextBox txtEjecutora 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   280
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   5655
      End
      Begin VB.TextBox txtAprueba 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   280
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   900
         Width           =   5655
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Entidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8520
         TabIndex        =   16
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Siglas"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   660
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pliego"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Unidad Ejecutora"
         Height          =   195
         Left            =   3900
         TabIndex        =   13
         Top             =   660
         Width           =   1230
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Instrumento - Aprueba o Modifica"
         Height          =   195
         Left            =   2820
         TabIndex        =   12
         Top             =   960
         Width           =   2340
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
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
      Left            =   10080
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAprobar 
      Caption         =   "Aprobar Plan Anual"
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
      Left            =   7560
      TabIndex        =   0
      Top             =   5040
      Width           =   2475
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   3075
      Left            =   120
      TabIndex        =   18
      Top             =   1860
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5424
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      BackColorSel    =   14942183
      ForeColorSel    =   8388608
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   16
   End
End
Attribute VB_Name = "frmLogPlanAnualAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nPlanAnualNro As Integer

Private Sub cmdAgencias_Click()
frmLogPlanAgeArea.Inicio 2, nPlanAnualNro, CInt(txtAnio.Text)
End Sub

Private Sub cmdAprobar_Click()
Dim oConn As New DConecta, sSql As String, nAnio As Integer
Dim cLogMov As String

sSql = ""
nAnio = CInt(txtAnio.Text)

If MsgBox("¿ Está seguro de aprobar el Plan Anual " & txtAnio.Text & " ?" + Space(10), vbQuestion + vbYesNo, "Confirme operación") = vbYes Then

   cLogMov = GetLogMovNro
   
   sSql = "UPDATE LogPlanAnual SET nPlanAnualEstado = 2 WHERE nPlanAnualNro = " & nPlanAnualNro & " "
   If oConn.AbreConexion Then
      oConn.Ejecutar sSql
   End If
   oConn.CierraConexion
   
   'DatosPlanAnual nAnio
   MsgBox "Se ha aprobó el Plan Anual en este nivel!" + Space(10), vbInformation
End If
End Sub

Private Sub cmdAreas_Click()
frmLogPlanAgeArea.Inicio 1, nPlanAnualNro, CInt(txtAnio.Text)
End Sub

Private Sub Form_Load()
CentraForm Me
Me.Caption = "Aprobación del Plan Anual de Adquisiciones y Contrataciones"
FormaFlex
'sstReq.Tab = 0
nPlanAnualNro = 0
txtAnio.Text = Year(gdFecSis) + 1
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Sub DatosPlanAnual(ByVal pnAnio As Integer)
Dim rs As New ADODB.Recordset, oConn As New DConecta
Dim sSql As String

sSql = "SELECT p.nPlanAnualNro,p.nPlanAnualAnio,p.cPlanAnualEntidad,p.cPlanAnualRUC,p.cPlanAnualSiglas, " & _
       "       p.cPlanAnualPliego , p.cPlanAnualEjecutor, p.cPlanAnualAprueba, p.nPlanAnualEstado, e.cEstado, " & _
       "       p.cUsuApro1,p.cUsuApro2,p.cUsuApro3 " & _
       "  from LogPlanAnual p inner join (select nConsValor as nEstado, cEstado=cConsDescripcion from Constante where nConsCod = 9049 and nConsCod<>nConsValor) e on p.nPlanAnualEstado = e.nEstado " & _
       " where nPlanAnualAnio = " & pnAnio & " and nPlanAnualEstado = 2"
       
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSql)
   oConn.CierraConexion
   If Not rs.EOF Then
      nPlanAnualNro = rs!nPlanAnualNro
      txtEntidad.Text = rs!cPlanAnualEntidad
      txtSiglas.Text = rs!cPlanAnualSiglas
      txtRUC.Text = rs!cPlanAnualRUC
      txtEjecutora.Text = rs!cPlanAnualEjecutor
      txtAprueba.Text = rs!cPlanAnualAprueba
      txtPliego.Text = rs!cPlanAnualPliego
      ListaPlanAnual pnAnio
   Else
      MsgBox "No hay un Plan Anual de Adquisiciones o Contrataciones por aprobar..." + Space(10), vbInformation
      cmdAgencias.Visible = False
      cmdAreas.Visible = False
      cmdAprobar.Visible = False
   End If
End If
End Sub

Sub ListaPlanAnual(ByVal pnAnio As Integer)
Dim oConn As New DConecta, rs As New ADODB.Recordset, i As Integer, nSuma As Currency
Dim sSql As String
sSql = ""
nSuma = 0
FormaFlex

If oConn.AbreConexion Then
  
  sSql = "SELECT d.nPlanAnualNro,d.nPlanAnualItem,d.nProSelTpoCod,d.nProSelSubTpo,d.nObjetoCod, o.cObjeto,d.nPlanAnualMes, d.cUbigeoCod, " & _
  "       d.cCIIU , d.cSintesis, d.nMoneda, d.nValorEstimado, r.cAbreviatura, f.cFuenteFinanciamiento " & _
  "  from LogPlanAnualDetalle d " & _
  "  left outer join (select nProSelTpoCod,nProSelSubTpo,cAbreviatura from LogProSelTpoRangos) r on r.nProSelTpoCod = d.nProSelTpoCod and r.nProSelSubTpo = d.nProSelSubTpo " & _
  "  left outer join (select nConsValor as nFuenteFinCod, cFuenteFinanciamiento=cConsDescripcion from Constante where nConsCod = 9046 and nConsCod<>nConsValor) f on d.nFuenteFinCod = f.nFuenteFinCod " & _
  "  left outer join (select nConsValor as nObjetoCod, cObjeto=cConsDescripcion from Constante where nConsCod = 9048 and nConsCod<>nConsValor) o on d.nObjetoCod = o.nObjetoCod  " & _
  " Where d.nPlanAnualAnio = " & pnAnio & " And d.nPlanAnualEstado = 1"

   If Len(sSql) = 0 Then Exit Sub
   
   Set rs = oConn.CargaRecordSet(sSql)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.RowHeight(i) = 500
         MSFlex.TextMatrix(i, 0) = rs!nPlanAnualNro
         MSFlex.TextMatrix(i, 1) = rs!nPlanAnualItem
         MSFlex.TextMatrix(i, 2) = rs!nProSelTpoCod
         MSFlex.TextMatrix(i, 3) = rs!nProSelSubTpo
         MSFlex.TextMatrix(i, 4) = rs!nObjetoCod
         MSFlex.TextMatrix(i, 5) = rs!cAbreviatura
         MSFlex.TextMatrix(i, 6) = rs!cObjeto
         MSFlex.TextMatrix(i, 7) = rs!cCIIU
         MSFlex.TextMatrix(i, 8) = rs!cSintesis
         If rs!nPlanAnualMes > 0 Then
            MSFlex.TextMatrix(i, 9) = UCase(mMes(rs!nPlanAnualMes))
         Else
            MSFlex.TextMatrix(i, 9) = ""
         End If
         MSFlex.TextMatrix(i, 10) = IIf(rs!nMoneda = 1, "SOLES", "DOLARES")
         MSFlex.TextMatrix(i, 11) = FNumero(rs!nValorEstimado)
         MSFlex.TextMatrix(i, 12) = rs!cFuenteFinanciamiento
         MSFlex.TextMatrix(i, 13) = GetUbigeoConsucode(rs!cUbiGeoCod)
         rs.MoveNext
      Loop
   End If
End If
End Sub

Sub FormaFlex()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(0) = 420
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 400:    MSFlex.TextMatrix(0, 1) = "Nro":     MSFlex.ColAlignment(1) = 4
MSFlex.ColWidth(2) = 0
MSFlex.ColWidth(3) = 0
MSFlex.ColWidth(4) = 0:      MSFlex.TextMatrix(0, 4) = "Nro":     MSFlex.ColAlignment(4) = 4
MSFlex.ColWidth(5) = 600:    MSFlex.TextMatrix(0, 5) = "Tipo de Proceso"
MSFlex.ColWidth(6) = 1000:   MSFlex.TextMatrix(0, 6) = "Objeto"
MSFlex.ColWidth(7) = 600:    MSFlex.TextMatrix(0, 7) = " CIIU":    MSFlex.ColAlignment(7) = 4
MSFlex.ColWidth(8) = 3800:   MSFlex.TextMatrix(0, 8) = "Síntesis de Especificaciones Técnicas"
MSFlex.ColWidth(9) = 1200:   MSFlex.TextMatrix(0, 9) = "Fecha Probable de Convocatoria":    MSFlex.ColAlignment(9) = 4
MSFlex.ColWidth(10) = 900:   MSFlex.TextMatrix(0, 10) = "Moneda":    MSFlex.ColAlignment(10) = 4
MSFlex.ColWidth(11) = 1000:  MSFlex.TextMatrix(0, 11) = "Valor Estimado"
MSFlex.ColWidth(12) = 2500:  MSFlex.TextMatrix(0, 12) = "Fuente de Financiamiento" ' MSFlex.ColAlignment(12) = 4
MSFlex.ColWidth(13) = 3500:  MSFlex.TextMatrix(0, 13) = "Ubicación Geográfica" ' MSFlex.ColAlignment(12) = 4
MSFlex.WordWrap = True
End Sub

Private Sub txtAnio_Change()
If Len(Trim(txtAnio)) = 4 Then
   DatosPlanAnual CInt(txtAnio.Text)
Else
   FormaFlex
End If
End Sub
