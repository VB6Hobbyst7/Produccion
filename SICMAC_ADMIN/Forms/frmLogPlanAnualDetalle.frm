VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogPlanAnualDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan Anual de Adquisiciones y Contrataciones"
   ClientHeight    =   6555
   ClientLeft      =   1620
   ClientTop       =   1755
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   9255
      Begin VB.TextBox txtItem 
         BackColor       =   &H00EAFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         MaxLength       =   4
         TabIndex        =   52
         Top             =   300
         Width           =   615
      End
      Begin VB.ComboBox cboTipoCon 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   3840
         Width           =   2355
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   555
         Left            =   1860
         Locked          =   -1  'True
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Top             =   5280
         Width           =   7215
      End
      Begin VB.TextBox txtEntidadAdquisicion 
         Height          =   315
         Left            =   1860
         MaxLength       =   20
         TabIndex        =   46
         Top             =   4560
         Width           =   7155
      End
      Begin VB.TextBox txtEntidadConvocante 
         Height          =   315
         Left            =   1860
         MaxLength       =   20
         TabIndex        =   44
         Top             =   4200
         Width           =   7155
      End
      Begin VB.CommandButton cmdUbigeo 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8640
         TabIndex        =   42
         Top             =   4920
         Width           =   400
      End
      Begin VB.TextBox txtUbigeo 
         Height          =   315
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   4920
         Width           =   6735
      End
      Begin VB.ComboBox cboFteFin 
         Height          =   315
         Left            =   6060
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   3120
         Width           =   3015
      End
      Begin VB.ComboBox cboMeses 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   3480
         Width           =   2355
      End
      Begin VB.ComboBox cboTipoCompra 
         Height          =   315
         Left            =   6060
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   3480
         Width           =   3015
      End
      Begin VB.CommandButton cmdCatalogo 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3780
         TabIndex        =   33
         Top             =   2445
         Width           =   405
      End
      Begin VB.TextBox txtCatalogo 
         Height          =   315
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   32
         Top             =   2415
         Width           =   2355
      End
      Begin VB.TextBox txtCIIU 
         Height          =   315
         Left            =   7620
         MaxLength       =   5
         TabIndex        =   29
         Top             =   1455
         Width           =   1440
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1860
         MaxLength       =   15
         TabIndex        =   27
         Top             =   2775
         Width           =   2355
      End
      Begin VB.ComboBox cboUnidad 
         Height          =   315
         Left            =   6060
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2775
         Width           =   3015
      End
      Begin VB.ComboBox cboObj 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1455
         Width           =   4275
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmLogPlanAnualDetalle.frx":0000
         Left            =   1860
         List            =   "frmLogPlanAnualDetalle.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3135
         Width           =   855
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2700
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   3135
         Width           =   1515
      End
      Begin VB.TextBox txtSintesis 
         Height          =   555
         Left            =   1860
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1815
         Width           =   7215
      End
      Begin VB.CommandButton cmdProSel 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8640
         TabIndex        =   12
         Top             =   1095
         Width           =   400
      End
      Begin VB.TextBox txtProSelDescripcion 
         Height          =   315
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1095
         Width           =   6735
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   6060
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2415
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CheckBox chkItemUnico 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Unico"
         Height          =   195
         Left            =   2880
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtAntecedentes 
         Height          =   315
         Left            =   1860
         MaxLength       =   100
         TabIndex        =   8
         Top             =   735
         Width           =   7185
      End
      Begin VB.ComboBox cboPrecedente 
         Height          =   315
         ItemData        =   "frmLogPlanAnualDetalle.frx":0018
         Left            =   5220
         List            =   "frmLogPlanAnualDetalle.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   315
         Width           =   915
      End
      Begin VB.TextBox txtAnio 
         BackColor       =   &H00EAFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1860
         MaxLength       =   4
         TabIndex        =   6
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Item Nro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   7560
         TabIndex        =   53
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Convocatoria"
         Height          =   195
         Left            =   180
         TabIndex        =   51
         Top             =   3900
         Width           =   1530
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   180
         TabIndex        =   49
         Top             =   5280
         Width           =   1065
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Org. Adq./Contratación"
         Height          =   195
         Left            =   180
         TabIndex        =   47
         Top             =   4620
         Width           =   1650
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Entidad Convocante"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   4260
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lugar de Ejecución"
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Top             =   4960
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fuente Financiamiento"
         Height          =   195
         Left            =   4380
         TabIndex        =   40
         Top             =   3180
         Width           =   1605
      End
      Begin VB.Label lblEspecificaciones 
         AutoSize        =   -1  'True
         Caption         =   "Mes de Convocatoria"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Top             =   3540
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Modalidad Adquisión"
         Height          =   195
         Left            =   4380
         TabIndex        =   38
         Top             =   3540
         Width           =   1470
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Catálogo Bienes/Serv."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   2475
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Código CIIU"
         Height          =   195
         Left            =   6600
         TabIndex        =   30
         Top             =   1515
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Unidad Medida"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4380
         TabIndex        =   28
         Top             =   2835
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   2835
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Objeto Selección"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   1515
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor Estimado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   3195
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sintesis Especificaciones  >> Técnicas"
         Height          =   585
         Left            =   180
         TabIndex        =   22
         Top             =   1785
         Width           =   1560
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTipoProceso 
         AutoSize        =   -1  'True
         Caption         =   "Proceso Selección"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   1155
         Width           =   1335
      End
      Begin VB.Label lblGrupo 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Bienes"
         Height          =   195
         Left            =   4380
         TabIndex        =   20
         Top             =   2475
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Antecedentes"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   795
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Precedente"
         Height          =   195
         Left            =   4260
         TabIndex        =   18
         Top             =   375
         Width           =   825
      End
      Begin VB.Label lblPlan 
         AutoSize        =   -1  'True
         Caption         =   "Plan Anual Año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   360
         Width           =   1320
      End
   End
   Begin VB.TextBox txtUbigeoCod 
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   5700
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txtProSelSub 
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Top             =   5700
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txtProSelTpo 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   5700
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8220
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   735
      Left            =   240
      TabIndex        =   43
      Top             =   4800
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1296
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   5
   End
End
Attribute VB_Name = "frmLogPlanAnualDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vpHaGrabado As Boolean
Public vpProSelCod As String
Public vpSintesis As String
Public vpCIIU As String
Public vpMes As String
Public vpFteFinCod As String

Dim sSQL As String, NuevoRegistro As Boolean
Dim cObjeto As String, cObjCod As String, nMonto As Currency, nMoneda As Integer
Dim cDescripcion As String, cAbreviatura As String, cSintesis As String
Dim nPlanNro As Integer, nPlanItem As Integer, nTpoCod As Integer, nSubTpo As Integer

Public Sub Inicio(ByVal pnPlanNroReg As Integer, ByVal pnPlanItem As Integer, Optional EsNuevo As Boolean = False)
nPlanNro = pnPlanNroReg
nPlanItem = pnPlanItem
NuevoRegistro = EsNuevo
Me.Show 1
End Sub

Private Sub cmdCatalogo_Click()
Dim k As Integer

sSQL = "select * from LogProSelCatalogo order by cCatalogoDescripcion"
frmLogProSelSelector.Consulta sSQL, "Seleccione Bien del Catálogo"
If frmLogProSelSelector.vpHaySeleccion Then
   txtCatalogo.Text = frmLogProSelSelector.vpCodigo
   'txtProSelDescripcion.Text = frmLogProSelSelector.vpDescripcion
   'txtProSelSub.Text = "1"
End If

End Sub




Private Sub Form_Load()
CentraForm Me
FormaFlex
txtAnio = Year(Date)
Me.vpHaGrabado = False
cboMoneda.ListIndex = 0
LimpiaCampos
If NuevoRegistro Then
   lblGrupo.Visible = True
   cboGrupo.Visible = True
Else
   RecuperaDatosRequerimiento
End If
End Sub

Sub LimpiaCampos()
Dim Rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim nMes As Integer

If oConn.AbreConexion Then
   If NuevoRegistro Then
      Set Rs = oConn.CargaRecordSet("Select nMaxItem=Max(nPlanAnualItem) from LogPlanAnualDetalle where nPlanAnualNro = " & nPlanNro & " ")
      If Not Rs.EOF Then
         nPlanItem = Rs!nMaxItem + 1
      Else
         nPlanItem = 1
      End If
   End If
   
   cboMeses.Clear
   sSQL = "select cValor,cNomTab from DBComunes..TablaCod where cCodTab like 'EZ%' and len(cCodTab)=4"
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         cboMeses.AddItem Rs!cNomTab
         Rs.MoveNext
      Loop
      If nMes > 0 Then
         cboMeses.ListIndex = nMes - 1
      Else
         cboMeses.ListIndex = 0
      End If
   End If
   
   '-------------------------------------------------------------------------------
   cboFteFin.Clear
   sSQL = "select nConsValor, cConsDescripcion from Constante where nConsCod = '9046' and nConsCod<>nConsValor"
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         cboFteFin.AddItem Rs!cConsDescripcion
         cboFteFin.ItemData(cboFteFin.ListCount - 1) = Rs!nConsValor
         Rs.MoveNext
      Loop
      cboFteFin.ListIndex = 0
   End If
   
   '-------------------------------------------------------------------------------
   sSQL = "select nConsValor,cConsDescripcion from Constante where nConsCod =9089  and nConsCod <>nConsValor"
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         cboUnidad.AddItem Rs!cConsDescripcion
         cboUnidad.ItemData(cboUnidad.ListCount - 1) = Rs!nConsValor
         Rs.MoveNext
      Loop
      cboUnidad.ListIndex = 0
   End If
   '-------------------------------------------------------------------------------
   
   sSQL = "select nConsValor,cConsDescripcion from Constante where nConsCod =9096  and nConsCod <>nConsValor"
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         cboTipoCon.AddItem Rs!cConsDescripcion
         cboTipoCon.ItemData(cboTipoCon.ListCount - 1) = Rs!nConsValor
         Rs.MoveNext
      Loop
      cboTipoCon.ListIndex = 0
   End If
   
   '-------------------------------------------------------------------------------
   cboTipoCompra.Clear
   sSQL = "select nConsValor, cConsDescripcion from Constante where nConsCod = '9081' and nConsCod<>nConsValor"
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         cboTipoCompra.AddItem Rs!cConsDescripcion
         cboTipoCompra.ItemData(cboTipoCompra.ListCount - 1) = Rs!nConsValor
         Rs.MoveNext
      Loop
      cboTipoCompra.ListIndex = 0
   End If
   
   '-------------------------------------------------------------------------------
   cboGrupo.Clear
   sSQL = "select * from BSGrupos where len(cBSGrupoCod)=4 order by cBSGrupoCod"
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         cboGrupo.AddItem Rs!cBSGrupoDescripcion
         cboGrupo.ItemData(cboGrupo.ListCount - 1) = Rs!cBSGrupoCod
         Rs.MoveNext
      Loop
      cboGrupo.ListIndex = 0
   End If
   
   '-------------------------------------------------------------------------------
   CboObj.Clear
   sSQL = "select nConsValor as nObjetoCod, cObjeto=cConsDescripcion from Constante where nConsCod = 9048 and nConsCod<>nConsValor"
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         CboObj.AddItem Rs!cObjeto
         CboObj.ItemData(CboObj.ListCount - 1) = Rs!nObjetoCod
         Rs.MoveNext
      Loop
      cboGrupo.ListIndex = 0
   End If
   
   
End If


End Sub

Sub RecuperaDatosRequerimiento()
Dim Rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim nObj As Integer, nMes As Integer
Dim i As Integer, nAnio As Integer
Dim k As Integer
Dim nIndex As Integer

If oConn.AbreConexion Then

   sSQL = "SELECT d.nProSelTpoCod,d.nProSelSubTpo,d.nObjetoCod,o.cObjeto,nConsucodeCod,d.cUbigeoCod,d.nItem, " & _
   "       d.cCIIU, d.cSintesis, d.nMoneda, d.nValorEstimado, cProceso=coalesce(r.cProceso,''),r.cSubTipo ,f.cFuenteFinanciamiento, " & _
   "       d.nPlanAnualMes, d.nPlanAnualAnio, d.nItemUnico,d.nItem,d.nPrecedente,d.nUnidad, d.nCantidad,d.cEntidadAdquisicion, " & _
   "       d.cCatalogo " & _
   "  from LogPlanAnualDetalle d " & _
   "  left outer join (select r.nProSelTpoCod,r.nConsucodeCod, r.nProSelSubTpo,cProceso=t.cProSelTpoDescripcion,cSubTipo=r.cProSelSubTpo from LogProSelTpoRangos r inner join LogProSelTpo t on r.nProSelTpoCod = t.nProSelTpoCod) r on r.nProSelTpoCod = d.nProSelTpoCod and r.nProSelSubTpo = d.nProSelSubTpo " & _
   "  left outer join (select nConsValor as nFuenteFinCod, cFuenteFinanciamiento=cConsDescripcion from Constante where nConsCod = 9046 and nConsCod<>nConsValor) f on d.nFuenteFinCod = f.nFuenteFinCod " & _
   "  left outer join (select nConsValor as nObjetoCod, cObjeto=cConsDescripcion from Constante where nConsCod = 9048 and nConsCod<>nConsValor) o on d.nObjetoCod = o.nObjetoCod " & _
   " where d.nPlanAnualEstado=1 and " & _
   "       d.nPlanAnualNro = " & nPlanNro & " and d.nPlanAnualItem = " & nPlanItem & " "
   
   ' d.nPlanAnualAnio = 2006 and
   
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      'txtCIIU.Text = rs!cCIIU
      nAnio = Rs!nPlanAnualAnio
      nMes = Rs!nPlanAnualMes
      txtAnio = Rs!nPlanAnualAnio
      txtMonto.Text = FNumero(Rs!nValorEstimado)
      txtSintesis.Text = Rs!cSintesis
      txtProSelDescripcion = Rs!cProceso
      txtProSelTpo = Rs!nProSelTpoCod
      txtProSelSub = Rs!nProSelSubTpo
      txtUbigeoCod.Text = Rs!cUbigeoCod
      txtCIIU.Text = Rs!cCIIU
      txtCatalogo.Text = Rs!cCatalogo
      If Len(Rs!cUbigeoCod) = 6 Then
         txtUbigeo.Text = GetUbigeoConsucode(Rs!cUbigeoCod)
      End If
      chkItemUnico.value = Rs!nItemUnico
      txtItem.Text = Rs!nItem
      cboPrecedente.ListIndex = Rs!nPrecedente
      txtCantidad.Text = Rs!nCantidad
      txtItem.Text = Rs!nItem
      txtEntidadAdquisicion.Text = Rs!cEntidadAdquisicion
      nIndex = -1
      For k = 0 To CboObj.ListCount - 1
          If CboObj.ItemData(k) = Rs!nObjetoCod Then
             nIndex = k
             Exit For
          End If
      Next
      CboObj.ListIndex = nIndex


      nIndex = -1
      For k = 0 To cboUnidad.ListCount - 1
          If cboUnidad.ItemData(k) = Rs!nUnidad Then
             nIndex = k
             Exit For
          End If
      Next
      cboUnidad.ListIndex = nIndex


      Select Case Rs!nMoneda
          Case 1
               cboMoneda.ListIndex = 0
          Case 2
               cboMoneda.ListIndex = 1
      End Select
      
   End If

   
  ' i = 0
  ' sSQL = "select d.cBSCod, b.cBSDescripcion, d.nCantidad, t.cUnidad " & _
  ' "  from LogPlanAnualDetalleBS d inner join LogProSelBienesServicios b on d.cBSCod = b.cProSelBSCod " & _
  ' "       inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) t on b.nBSUnidad = t.nBSUnidad " & _
  ' " Where d.nPlanAnualNro = " & nPlanNro & " And d.nPlanAnualItem = " & nPlanItem & " "

   'Set rs = oConn.CargaRecordSet(sSQL)
   'If Not rs.EOF Then
   '   Do While Not rs.EOF
   '      i = i + 1
   '      InsRow MSFlex, i
   '      MSFlex.TextMatrix(i, 1) = rs!cBSCod
   '      MSFlex.TextMatrix(i, 2) = rs!cBSDescripcion
   '      MSFlex.TextMatrix(i, 3) = rs!nCantidad
   '      MSFlex.TextMatrix(i, 4) = rs!cUnidad
   '      rs.MoveNext
   '   Loop
   'End If
Else
   MsgBox "No se puede establecer conexión..." + Space(10), vbInformation, "Aviso"
   Exit Sub
End If
End Sub

Sub FormaFlex()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(1) = 8
MSFlex.RowHeight(0) = 300
MSFlex.ColWidth(0) = 0
MSFlex.TextMatrix(0, 1) = "Código":       MSFlex.ColWidth(1) = 840
MSFlex.TextMatrix(0, 2) = "Descripción":  MSFlex.ColWidth(2) = 3200
MSFlex.TextMatrix(0, 3) = "Cantidad":     MSFlex.ColWidth(3) = 700: MSFlex.ColAlignment(3) = 4
MSFlex.TextMatrix(0, 4) = "Unidad":       MSFlex.ColWidth(4) = 1200
End Sub


Private Sub CmdAceptar_Click()
Dim nMes As Integer
Dim oConn As New DConecta
Dim cFteFin As String, cTpoCom As String, nFteFin As Integer
Dim nTpoCom As Integer, nAnio As Integer, nMoneda As Integer
Dim cBSGrupoCod As String, nTpoCon As Integer, nModAdq As Integer
Dim nObjCod As Integer, nUnidad As Integer, cCatalogo As String
Dim nPrecedente As Integer, nItemUnico As Integer

nAnio = CInt(txtAnio.Text)

If MsgBox("¿ Está seguro de asignar este proceso ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   nMes = cboMeses.ListIndex + 1
   nItemUnico = chkItemUnico.value
   nFteFin = cboFteFin.ItemData(cboFteFin.ListIndex)
   'nTpoCom = cboTipoCompra.ItemData(cboTipoCompra.ListIndex)
   nMoneda = cboMoneda.ListIndex + 1
   nUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)
   nTpoCon = cboTipoCon.ItemData(cboTipoCon.ListIndex)
   nModAdq = cboTipoCompra.ItemData(cboTipoCompra.ListIndex)
   
   
   If NuevoRegistro Then
      nObjCod = CboObj.ItemData(CboObj.ListIndex)
      cBSGrupoCod = Format(cboGrupo.ItemData(cboGrupo.ListIndex), "0000")
   
      sSQL = "INSERT INTO LogPlanAnualDetalle (nPlanAnualNro, nPlanAnualItem, nPlanAnualAnio, nPlanAnualMes, nProSelTpoCod, nProSelSubTpo, cBSGrupoCod, " & _
             "            nObjetoCod , cCIIU, cSintesis, nMoneda, nValorEstimado, nFuenteFinCod, cUbigeoCod, nPlanAnualEstado ) " & _
             "  VALUES (" & nPlanNro & "," & nPlanItem & "," & nAnio & "," & nMes & "," & CInt(VNumero(txtProSelTpo.Text)) & "," & CInt(VNumero(txtProSelSub.Text)) & " ,'" & cBSGrupoCod & "', " & _
             "  " & nObjCod & ", '" & txtCIIU.Text & "','" & txtSintesis.Text & "'," & nMoneda & "," & VNumero(txtMonto.Text) & "," & nFteFin & ",'" & txtUbigeoCod & "',1) "
             
   Else
      sSQL = "update LogPlanAnualDetalle SET nProSelTpoCod = " & CInt(VNumero(txtProSelTpo.Text)) & ", " & _
             "       nProSelSubTpo = " & CInt(VNumero(txtProSelSub.Text)) & ", nItemUnico = " & nItemUnico & " , " & _
             "          cUbigeoCod = '" & txtUbigeoCod & "',          nPlanAnualMes = " & nMes & ",  nCantidad = " & Val(txtCantidad.Text) & ", " & _
             "             nUnidad = " & nUnidad & ",           cEntidadAdquisicion = '" & txtEntidadAdquisicion.Text & "', " & _
             "             nItem   = " & CInt(txtItem.Text) & " , nTipoConvocatoria = '" & nTpoCon & "', cCatalogo = '" & txtCatalogo.Text & "', " & _
             "           cSintesis = '" & txtSintesis.Text & "',              cCIIU = '" & txtCIIU.Text & "', " & _
             "      nValorEstimado = " & VNumero(txtMonto.Text) & " , nFuenteFinCod = " & nFteFin & "  " & _
             " WHERE nPlanAnualNro = " & nPlanNro & " and nPlanAnualItem = " & nPlanItem & " "
             
             'cCatalogo = '" & txtCatalogo.Text & "',
   End If
   
   If oConn.AbreConexion Then
      oConn.Ejecutar sSQL
   End If
   Me.vpHaGrabado = True
   Unload Me
End If
End Sub

Private Sub cmdCancelar_Click()
Me.vpHaGrabado = False
Unload Me
End Sub

Private Sub cmdProSel_Click()
Dim sSQL As String
sSQL = "select nProSelTpoCod,cProSelTpoDescripcion from LogProSelTpo"
frmLogProSelSelector.Consulta sSQL, "Seleccione el Proceso"
If frmLogProSelSelector.vpHaySeleccion Then
   txtProSelTpo.Text = frmLogProSelSelector.vpCodigo
   txtProSelDescripcion.Text = frmLogProSelSelector.vpDescripcion
   txtProSelSub.Text = "1"
End If
End Sub

Private Sub cmdUbiGeo_Click()
frmLogProSelSeleUbiGeo.FuenteConsucode True
'If Len(Trim(frmLogProSelSeleUbiGeo.vpCodUbigeo)) > 0 Then
   'txtUbigeoCod.Text = frmLogProSelSeleUbiGeo.vpCodUbigeo
   'txtUbigeo.Text = frmLogProSelSeleUbiGeo.vpUbigeoDesc
If Len(Trim(frmLogProSelSeleUbiGeo.gvCodigo)) > 0 Then
   txtUbigeoCod.Text = frmLogProSelSeleUbiGeo.gvCodigo
   txtUbigeo.Text = frmLogProSelSeleUbiGeo.gvNoddo
   
End If
End Sub

'Private Sub cmdareaagencia_Click()
'frmSelectorArbol.FormaArbolConsulta "Select cAgeCod,cAgeDescripcion from Agencias where nEstado=1", "Seleccion de Agencias"
'If Len(Trim(frmSelectorArbol.vpSeleccion)) > 0 Then
'   txtAgeCod.Text = frmSelectorArbol.vpCodigo
'   txtAgencia.Text = frmSelectorArbol.vpDescripcion
'End If
'End Sub


Private Sub txtCIIU_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtSintesis.SetFocus
End If
End Sub

'Private Sub txtMonto_Change()
'If VNumero(txtMonto) > 0 Then
'   cDescripcion = ""
'   cAbreviatura = ""
'   txtProSelCod.Text = DeterminaProcesoSeleccion(cObjCod, nMonto, cDescripcion, cAbreviatura)
'   txtProSel.Text = cDescripcion
'End If
'End Sub

Private Sub txtSintesis_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Function DeterminaProcesoSeleccion(ByVal vObjetoCod As String, ByVal vMonto As Currency, ByRef vDescripcion As String, ByRef vAbreviatura As String) As String
Dim oConn As New DConecta
Dim Rs As New ADODB.Recordset

DeterminaProcesoSeleccion = ""
sSQL = ""
Select Case vObjetoCod
    Case "11"
         'sSQL = "select nProSelTpoCod,cProSelDescripcion,cAbreviatura " & _
         '       "  from LogProSelTpo " & _
         '       " where " & vMonto & " > nBienesMin  and  " & vMonto & " < nBienesMax and " & _
         '       "       nBienesMin>0 and nBienesMax>0"
                
         sSQL = "select r.nProSelTpoCod,t.cProSelTpoDescripcion,r.cProSelSubTpo,r.cAbreviatura " & _
                " from LogProSelTpoRangos r inner join LogProSelTpo t on r.nProSelTpoCod = t.nProSelTpoCod " & _
                "where " & vMonto & " > nBienesMin  and  " & vMonto & " < nBienesMax and " & _
                "       nBienesMin>0 and nBienesMax>0"

    Case "12"

         sSQL = "select r.nProSelTpoCod,t.cProSelTpoDescripcion,r.cProSelSubTpo,r.cAbreviatura " & _
                " from LogProSelTpoRangos r inner join LogProSelTpo t on r.nProSelTpoCod = t.nProSelTpoCod " & _
                " where " & vMonto & " > nServiMin  and  " & vMonto & " < nServiMax and " & _
                "      nServiMin>0 and nServiMax>0 "
    Case "13"
    
         sSQL = ""
End Select

If Len(sSQL) = 0 Then Exit Function
   
If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      DeterminaProcesoSeleccion = Rs!nProSelTpoCod
      vDescripcion = Rs!cProSelTpoDescripcion
      vAbreviatura = Rs!cAbreviatura
   End If
End If

End Function


