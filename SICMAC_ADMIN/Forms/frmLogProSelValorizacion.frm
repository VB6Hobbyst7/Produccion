VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogProSelValorizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Valor Referencial "
   ClientHeight    =   4740
   ClientLeft      =   735
   ClientTop       =   2490
   ClientWidth     =   10515
   Icon            =   "frmLogProSelValorizacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   10515
   Begin VB.TextBox txtEdit 
      BackColor       =   &H00FCFFE1&
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   5280
      TabIndex        =   4
      Top             =   1620
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox cboMoneda 
      BackColor       =   &H00FCFFE1&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      ItemData        =   "frmLogProSelValorizacion.frx":08CA
      Left            =   4200
      List            =   "frmLogProSelValorizacion.frx":08D4
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   -30
      Width           =   10275
      Begin VB.CommandButton cmdConsolidar 
         Caption         =   "Consolidar &Requerimientos"
         Height          =   375
         Left            =   5880
         TabIndex        =   19
         Top             =   210
         Width           =   2115
      End
      Begin VB.CommandButton cmdEjeVal 
         Caption         =   "Asignar &Valor Referencial"
         Height          =   375
         Left            =   8040
         TabIndex        =   18
         Top             =   210
         Width           =   2115
      End
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
         Height          =   330
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "2005"
         Top             =   225
         Width           =   675
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proceso de Seleccion y Aquisiciones"
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
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   3150
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   2865
      Left            =   120
      TabIndex        =   5
      Top             =   630
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   5054
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   10
      FixedCols       =   0
      ForeColorFixed  =   -2147483646
      BackColorSel    =   16580577
      ForeColorSel    =   4194304
      BackColorBkg    =   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      GridColorUnpopulated=   -2147483633
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAFFFF&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   120
      TabIndex        =   6
      Top             =   3420
      Width           =   10275
      Begin VB.TextBox txtTotalDol 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox txtTotalSol 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL   S/.                                US$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4980
         TabIndex        =   9
         Top             =   240
         Width           =   3345
      End
   End
   Begin VB.Frame fraBot1 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   10275
      Begin VB.CommandButton cmdAgencias 
         Caption         =   "Agencias"
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdAreas 
         Caption         =   "Areas"
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   9060
         TabIndex        =   11
         Top             =   150
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DADEDC&
         Caption         =   "Tipos de Moneda"
         ForeColor       =   &H80000008&
         Height          =   580
         Left            =   5085
         TabIndex        =   12
         Top             =   0
         Width           =   3810
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "Imprimir"
            Height          =   375
            Left            =   2415
            TabIndex        =   22
            Top             =   150
            Width           =   1275
         End
         Begin VB.OptionButton opSol 
            BackColor       =   &H00DADEDC&
            Caption         =   "Soles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   260
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opDol 
            BackColor       =   &H00DADEDC&
            Caption         =   "Dólares"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00185B11&
            Height          =   195
            Left            =   1320
            TabIndex        =   13
            Top             =   260
            Width           =   1095
         End
      End
   End
   Begin VB.Frame fraBot2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7500
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   420
         TabIndex        =   16
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Menu cModificar 
      Caption         =   "cModificar"
      Visible         =   0   'False
      Begin VB.Menu cModificar1 
         Caption         =   "Quitar Requerimiento"
      End
      Begin VB.Menu cModificar2 
         Caption         =   "Agrupar por Usuario"
      End
      Begin VB.Menu cModificar3 
         Caption         =   "Agrupar por Articulo"
      End
   End
End
Attribute VB_Name = "frmLogProSelValorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim pbEstaValorizando As Boolean
Dim nFila As Integer
Dim bRecuperado As Boolean

Private Sub cmdAgencias_Click()
frmLogPlanAgeArea.Inicio 2, 0, CInt(txtAnio.Text)
End Sub

Private Sub cmdAreas_Click()
frmLogPlanAgeArea.Inicio 1, 0, CInt(txtAnio.Text)
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer, n As Integer

n = MSFlex.Rows - 1
For i = 1 To n
    MSFlex.TextMatrix(i, 5) = MSFlex.TextMatrix(i, 8)
    MSFlex.TextMatrix(i, 6) = MSFlex.TextMatrix(i, 9)
Next
SumaTotal
pbEstaValorizando = False
fraBot1.Visible = True
fraBot2.Visible = False
End Sub

Private Sub cModificar1_Click()
On Error GoTo cModificar1_ClickErr
    Dim i As Integer
    Dim K As Integer
    
    i = MSFlex.row
    If Len(Trim(MSFlex.TextMatrix(i, 1))) = 0 Then
       Exit Sub
    End If
    
    If MsgBox("¿ está seguro de quitar el elemento ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
       If MSFlex.Rows - 1 > 1 Then
          MSFlex.RemoveItem i
       Else
          'MSFlex.Clear          Quita las cabeceras
          For K = 0 To MSFlex.Cols - 1
              MSFlex.TextMatrix(i, K) = ""
          Next
          MSFlex.RowHeight(i) = 8
       End If
    End If
    Exit Sub
cModificar1_ClickErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cModificar2_Click()
    ConsolidacionRequerimientos "Usuario"
End Sub

Private Sub cModificar3_Click()
    ConsolidacionRequerimientos "Articulo"
End Sub

Private Sub Form_Load()
CentraForm Me
txtAnio.Text = Year(gdFecSis)
FormaFlexValor
pbEstaValorizando = False
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdEjeVal_Click()
pbEstaValorizando = True
fraBot1.Visible = False
fraBot2.Visible = True
MSFlex.row = 1
MSFlex.Col = 6
MSFlex.SetFocus
End Sub

Private Sub cmdGrabar_Click()
Dim oConn As New DConecta, i As Integer, n As Integer
Dim nPrecio As Currency, nMoneda As Integer, cProSelBSCod As String

n = MSFlex.Rows - 1

If MsgBox("¿ Está seguro de grabar los valores indicados ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then

   If oConn.AbreConexion Then
      For i = 1 To n
          If Len(MSFlex.TextMatrix(i, 5)) > 0 And VNumero(MSFlex.TextMatrix(i, 6)) > 0 Then
             nMoneda = 0
             Select Case MSFlex.TextMatrix(i, 5)
                 Case "S/."
                      nMoneda = 1
                 Case "US$"
                      nMoneda = 2
             End Select
             nPrecio = VNumero(MSFlex.TextMatrix(i, 6))
             cProSelBSCod = MSFlex.TextMatrix(i, 1)
             
             sSQL = "UPDATE LogProSelReqDetalle SET LogProSelReqDetalle.nMoneda = " & nMoneda & ", " & _
                    "                                  LogProSelReqDetalle.nPrecioUnitario = " & nPrecio & " " & _
                    " WHERE LogProSelReqDetalle.cBSCod = '" & cProSelBSCod & "' and LogProSelReqDetalle.nEstado = 1 "
             oConn.Ejecutar sSQL
             
          End If
      Next
      If CDbl(txtTotalSol.Text) = 0 Then
        nMoneda = 2
        nPrecio = CDbl(txtTotalDol.Text)
      Else
        nMoneda = 1
        nPrecio = CDbl(txtTotalSol.Text)
      End If
   End If
   pbEstaValorizando = False
   fraBot1.Visible = True
   fraBot2.Visible = False
End If
End Sub

Private Sub cmdConsolidar_Click()
Dim rs As New ADODB.Recordset, oConn As New DConecta
Dim YaValorizo As Boolean, nAnio As Integer, nResp As Integer

nAnio = CInt(VNumero(txtAnio.Text))

If RequerimientosNoAprobadosPS(nAnio) > 0 Then
   MsgBox "Existen requerimientos sin aprobar completamente..." + Space(10) + vbCrLf + _
          "    Sólo se consolidarán todos los aprobados ", vbInformation, "Verifique requerimientos"
End If

sSQL = "select top 1 nPrecioUnitario from LogProSelReqDetalle where nProSelNro=0 and nAnio=" & nAnio & " and nPrecioUnitario>0"
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
End If

If Not rs.EOF Then
   YaValorizo = True
Else
   YaValorizo = False
End If

Set rs = Nothing

If YaValorizo Then
   nResp = MsgBox(Space(30) + "A T E N C I O N" + vbCrLf + "El Consolidado ya tiene valores referenciales asignados" + Space(10) + vbCrLf + vbCrLf + Space(20) + "¿ Consolidar requerimientos ?" + vbCrLf + vbCrLf + _
                  "Presione SI para consolidar requerimientos nuevamente" + Space(10) + vbCrLf + _
                  "Presione NO para recuperar el consolidado anterior   " + Space(10), vbQuestion + vbYesNoCancel + vbDefaultButton3, "Aviso")

   If nResp = vbYes Then
      ConsolidacionRequerimientos
      bRecuperado = False
   ElseIf nResp = vbNo Then
      RecuperaConsolidado
      bRecuperado = True
   ElseIf nResp = vbCancel Then
   
   End If
Else
   ConsolidacionRequerimientos
End If
End Sub

Sub ConsolidacionRequerimientos(Optional ByVal psGrupo As String = "Articulo")
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta, i As Integer
Dim nAnio As Integer
nAnio = CInt(VNumero(txtAnio.Text))

FormaFlexValor
txtTotalSol = ""
txtTotalDol = ""
Select Case psGrupo
    Case "SG"
        sSQL = "select d.nProSelReqNro, d.cBSCod, g.cBSDescripcion, u.cUnidad , nCantidad " & _
            "  from LogProSelReqDetalle d " & _
            "       inner join LogProSelBienesServicios g on d.cBSCod = g.cProSelBSCod " & _
            "       inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 1019) u on g.nBSUnidad = u.nBSUnidad " & _
            " WHERE nEstado = 1 and (select count(*) from LogProSelAprobacion where nProSelReqNro = d.nProSelReqNro and nEstadoAprobacion=0) = 0 and nProSelNro=0 and d.nAnio = " & nAnio & " " & _
            "  group by d.nProSelReqNro, d.cBSCod, g.cBSDescripcion, u.cUnidad, nCantidad "
    Case "Articulo"
        sSQL = "select d.cBSCod, g.cBSDescripcion, u.cUnidad , nCantidad=sum(nCantidad) " & _
            "  from LogProSelReqDetalle d " & _
            "       inner join LogProSelBienesServicios g on d.cBSCod = g.cProSelBSCod " & _
            "       inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 1019) u on g.nBSUnidad = u.nBSUnidad " & _
            " WHERE nEstado >= 1 and (select count(*) from LogProSelAprobacion where nProSelReqNro = d.nProSelReqNro and nEstadoAprobacion=0) = 0 and nProSelNro=0 and d.nAnio = " & nAnio & " " & _
            "  group by d.cBSCod, g.cBSDescripcion, u.cUnidad "
    Case "Usuario"
        sSQL = "select d.cBSCod, g.cBSDescripcion, u.cUnidad , nCantidad=sum(nCantidad) " & _
            "  from LogProSelReqDetalle d " & _
            "       inner join LogProSelBienesServicios g on d.cBSCod = g.cProSelBSCod " & _
            "       inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 1019) u on g.nBSUnidad = u.nBSUnidad " & _
            " WHERE nEstado = 1 and (select count(*) from LogProSelAprobacion where nProSelReqNro = d.nProSelReqNro and nEstadoAprobacion=0) = 0 and nProSelNro=0 and d.nAnio = " & nAnio & " " & _
            "  group by d.cPersCod,d.cBSCod, g.cBSDescripcion, u.cUnidad "
End Select
    
    'sSQL = "select d.cBSCod, g.cBSDescripcion, u.cUnidad , nCantidad=sum(nCantidad) " & _
        "  from LogProSelReqDetalle d " & _
        "       inner join LogProSelBienesServicios g on d.cBSCod = g.cProSelBSCod " & _
        "       inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 1019) u on g.nBSUnidad = u.nBSUnidad " & _
        " WHERE nProSelNro=0 and d.nAnio = " & nAnio & " " & _
        "  group by d.cBSCod, g.cBSDescripcion, u.cUnidad "
       
   If oConn.AbreConexion Then
      Set rs = oConn.CargaRecordSet(sSQL)
      oConn.CierraConexion
   End If
   If Not rs.EOF Then
      i = 0
      Do While Not rs.EOF
'         If RequerimientoAprobado(Rs!nProSelReqNro) Then
            i = i + 1
            InsRow MSFlex, i
            MSFlex.RowHeight(i) = 330
            MSFlex.TextMatrix(i, 0) = i
            MSFlex.TextMatrix(i, 1) = rs!cBSCod
            MSFlex.TextMatrix(i, 2) = rs!cBSDescripcion
            MSFlex.TextMatrix(i, 3) = rs!nCantidad
            MSFlex.TextMatrix(i, 4) = rs!cUnidad
            MSFlex.TextMatrix(i, 5) = "S/."
            MSFlex.TextMatrix(i, 6) = FNumero(CargarValorRef(rs!cBSCod)) 'GetPrecioUnitario(nAnio - 1, 2, Rs!cBSCod)
            MSFlex.TextMatrix(i, 7) = VNumero(MSFlex.TextMatrix(i, 3)) * VNumero(MSFlex.TextMatrix(i, 6))
            DoEvents
'         End If
         rs.MoveNext
      Loop
      SumaTotal
    End If
End Sub

Sub RecuperaConsolidado()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta, i As Integer
Dim nAnio As Integer
nAnio = CInt(VNumero(txtAnio.Text))

FormaFlexValor
txtTotalSol = ""
txtTotalDol = ""

sSQL = "select d.nProSelReqNro, d.cBSCod, g.cBSDescripcion, d.nMoneda, d.nPrecioUnitario," & _
       "            u.cUnidad , nCantidad " & _
       "  from LogProSelReqDetalle d " & _
       "       inner join LogProSelBienesServicios g on d.cBSCod = g.cProSelBSCod " & _
       "       inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 1019) u on g.nBSUnidad = u.nBSUnidad " & _
       " WHERE d.nProSelNro = 0 and d.nAnio = " & nAnio & " and d.nEstado = 1 " & _
       "  group by d.nProSelReqNro, d.cBSCod, g.cBSDescripcion, u.cUnidad, d.nMoneda, d.nPrecioUnitario, nCantidad  "
      
    If oConn.AbreConexion Then
       Set rs = oConn.CargaRecordSet(sSQL)
       oConn.CierraConexion
    End If
    
    If Not rs.EOF Then
       i = 0
       Do While Not rs.EOF
            If RequerimientoAprobado(rs!nProSelReqNro) Then
                i = i + 1
                InsRow MSFlex, i
                MSFlex.RowHeight(i) = 330
                MSFlex.TextMatrix(i, 0) = Format(i, "00")
                MSFlex.TextMatrix(i, 1) = rs!cBSCod
                MSFlex.TextMatrix(i, 2) = rs!cBSDescripcion
                MSFlex.TextMatrix(i, 3) = rs!nCantidad
                MSFlex.TextMatrix(i, 4) = rs!cUnidad
                MSFlex.TextMatrix(i, 5) = IIf(rs!nMoneda = 1, "S/.", "US$")
                MSFlex.TextMatrix(i, 6) = IIf(rs!nPrecioUnitario = 0, FNumero(CargarValorRef(rs!cBSCod)), FNumero(rs!nPrecioUnitario))
                MSFlex.TextMatrix(i, 7) = FNumero(VNumero(MSFlex.TextMatrix(i, 3)) * VNumero(MSFlex.TextMatrix(i, 6)))
                'Para restaurar en caso cancele operación
                MSFlex.TextMatrix(i, 8) = IIf(rs!nMoneda = 1, "S/.", "US$")
                MSFlex.TextMatrix(i, 9) = FNumero(rs!nPrecioUnitario)
                DoEvents
            End If
          rs.MoveNext
       Loop
       SumaTotal
    End If
End Sub

Sub FormaFlexValor()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(-1) = 280
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 350:       MSFlex.TextMatrix(0, 0) = "Nº":     MSFlex.ColAlignment(0) = 4
MSFlex.ColWidth(1) = 850:       MSFlex.TextMatrix(0, 1) = "Codigo"
MSFlex.ColWidth(2) = 4000:      MSFlex.TextMatrix(0, 2) = "Descripción"
MSFlex.ColWidth(3) = 700:       MSFlex.TextMatrix(0, 3) = "Cantidad":     MSFlex.ColAlignment(3) = 4
MSFlex.ColWidth(4) = 1500:      MSFlex.TextMatrix(0, 4) = " U. Medida":   MSFlex.ColAlignment(4) = 1
MSFlex.ColWidth(5) = 600:       MSFlex.TextMatrix(0, 5) = "Moneda":       MSFlex.ColAlignment(5) = 4
MSFlex.ColWidth(6) = 950:       MSFlex.TextMatrix(0, 6) = " Precio Unit"
MSFlex.ColWidth(7) = 1000:       MSFlex.TextMatrix(0, 7) = " Sub-Total"
MSFlex.ColWidth(8) = 0
MSFlex.ColWidth(9) = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmLogProSelValorizacion = Nothing
End Sub

Private Sub MSFlex_DblClick()
If MSFlex.Col = 5 And pbEstaValorizando Then
   cboMoneda.Visible = True
   cboMoneda.Text = MSFlex.TextMatrix(MSFlex.row, 5)
   cboMoneda.Move MSFlex.Left + MSFlex.CellLeft - 30, MSFlex.Top + MSFlex.CellTop - 30, MSFlex.CellWidth + 30
End If
End Sub

Private Sub MSFlex_GotFocus()
If cboMoneda.Visible Then
   MSFlex.TextMatrix(MSFlex.row, 5) = cboMoneda.Text
   cboMoneda.Visible = False
End If
If txtEdit.Visible = False Then Exit Sub
MSFlex = txtEdit
txtEdit.Visible = False
SumaTotal
End Sub

Private Sub MSFlex_LeaveCell()
If cboMoneda.Visible Then
   MSFlex.TextMatrix(MSFlex.row, 5) = cboMoneda.Text
   cboMoneda.Visible = False
   SumaTotal
End If
If txtEdit.Visible = False Then Exit Sub
MSFlex = FNumero(txtEdit)
txtEdit.Visible = False
SumaTotal
End Sub


'*********************************************************************
'PROCEDIMIENTOS DEL FLEX
'*********************************************************************

Private Sub MSFlex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And MSFlex.Col = 6 And pbEstaValorizando Then
   MSFlex.TextMatrix(MSFlex.row, 6) = ""
   SumaTotal
End If
End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
If MSFlex.Col = 6 And pbEstaValorizando Then
   If Len(MSFlex.TextMatrix(MSFlex.row, 5)) = 0 Then
      MsgBox "Debe seleccionar un tipo de moneda..." + Space(10), vbInformation
   Else
      EditaFlex MSFlex, txtEdit, KeyAscii
   End If
End If

If MSFlex.Col = 5 And pbEstaValorizando Then
   nFila = MSFlex.row
   cboMoneda.Move MSFlex.Left + MSFlex.CellLeft - 30, MSFlex.Top + MSFlex.CellTop - 30, MSFlex.CellWidth + 30
   cboMoneda.Visible = True
   cboMoneda.SetFocus
End If
End Sub

Sub EditaFlex(MSFlex As Control, Edt As Control, KeyAscii As Integer)
Select Case KeyAscii
    Case 0 To 32
         Edt = MSFlex
         Edt.SelStart = 1000
    Case Else
         Edt = Chr(KeyAscii)
         Edt.SelStart = 1
End Select
Edt.Move MSFlex.Left + MSFlex.CellLeft - 15, MSFlex.Top + MSFlex.CellTop - 15, _
         MSFlex.CellWidth, MSFlex.CellHeight
Edt.Visible = True
Edt.SetFocus
End Sub

Private Sub MSFlex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 Then PopupMenu cModificar
End Sub

Private Sub txtAnio_GotFocus()
SelTexto txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   cmdConsolidar.SetFocus
End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(vbCr) Then
   KeyAscii = 0
   'txtEdit = FNumero(txtEdit)
End If
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode MSFlex, txtEdit, KeyCode, Shift
End Sub

Sub EditKeyCode(MSFlex As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27
         Edt.Visible = False
         MSFlex.SetFocus
    Case 13
         MSFlex.SetFocus
    Case 37                     'Izquierda
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col > 1 Then
            MSFlex.Col = MSFlex.Col - 1
         End If
    Case 39                     'Derecha
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col < MSFlex.Cols - 1 Then
            MSFlex.Col = MSFlex.Col + 1
         End If
    Case 38
         MSFlex.SetFocus
         DoEvents
         If MSFlex.row > MSFlex.FixedRows + 1 Then
            MSFlex.row = MSFlex.row - 1
         End If
    Case 40
         MSFlex.SetFocus
         DoEvents
         If MSFlex.row < MSFlex.Rows - 1 Then
            MSFlex.row = MSFlex.row + 1
         End If
End Select
End Sub

Sub SumaTotal()
Dim i As Integer, n As Integer, nSumaSol As Currency, nSumaDol As Currency

n = MSFlex.Rows - 1
nSumaSol = 0
nSumaDol = 0
For i = 1 To n
    MSFlex.TextMatrix(i, 7) = FNumero(VNumero(MSFlex.TextMatrix(i, 6)) * VNumero(MSFlex.TextMatrix(i, 3)))
    
    If VNumero(MSFlex.TextMatrix(i, 6)) > 0 Then
       Select Case MSFlex.TextMatrix(i, 5)
           Case "S/."
                nSumaSol = nSumaSol + VNumero(MSFlex.TextMatrix(i, 7))
           Case "US$"
                nSumaDol = nSumaDol + VNumero(MSFlex.TextMatrix(i, 7))
       End Select
    Else
       MSFlex.TextMatrix(i, 7) = ""
    End If

Next
txtTotalSol = FNumero(nSumaSol)
txtTotalDol = FNumero(nSumaDol)
End Sub

Private Sub cmdImprimir_Click()
Dim i As Integer, n As Integer, f As Integer, v As Variant
Dim nTotal As Currency, cMoneda As String

If opSol.value Then
   cMoneda = "S/."
End If
If opDol.value Then
   cMoneda = "US$"
End If

n = MSFlex.Rows - 1
f = FreeFile
Open App.path + "\Val01.txt" For Output As #f
nTotal = 0
Print #f, ""
Print #f, Space(23) + "REQUERIMIENTO NO PROGRAMABLES"
Print #f, Space(18) + "  VALORIZACION DE REQUERIMIENTOS EN " + IIf(opSol.value, "SOLES", "DOLARES")
Print #f, ""
Print #f, String(100, "=")
Print #f, "No Codigo     Descripcion" + Space(35) + "Cantidad         Moneda     Sub - Total"
Print #f, String(100, "-")
For i = 1 To n
    If MSFlex.TextMatrix(i, 5) = cMoneda Then
       Print #f, Format(i, "00") + " " + MSFlex.TextMatrix(i, 1) + " " + JIZQ(MSFlex.TextMatrix(i, 2), 40) + " " + JDER(MSFlex.TextMatrix(i, 3), 8) + "  " + JIZQ(MSFlex.TextMatrix(i, 4), 15) + " " + MSFlex.TextMatrix(i, 5) + "   " + JDER(FNumero(VNumero(MSFlex.TextMatrix(i, 3)) * VNumero(MSFlex.TextMatrix(i, 6))), 12)
       Print #f, CargarDatosR(MSFlex.TextMatrix(i, 1))
       nTotal = nTotal + VNumero(MSFlex.TextMatrix(i, 3)) * VNumero(MSFlex.TextMatrix(i, 6))
    End If
Next
Print #f, String(100, "-")
Print #f, Space(65) + "TOTAL " + Space(10) + cMoneda + JDER(FNumero(nTotal), 15)
Print #f, String(100, "=")
Close #f

v = Shell("notepad.exe " + App.path + "\val01.txt", vbNormalFocus)
End Sub

Private Function CargarDatosR(ByVal psBSCod As String) As String
On Error GoTo CargarDatosRErr
    Dim sSQL As String, rs As ADODB.Recordset, oCon As DConecta, i As Integer, nAnio As Integer
    Set oCon = New DConecta
    nAnio = CInt(VNumero(txtAnio.Text))
    sSQL = " select cPersNombre=replace(p.cPersNombre,'/',' '), nCantidad=sum(nCantidad), a.cAgeDescripcion " & _
           " from LogProSelReqDetalle d " & _
           " inner join Persona p on d.cPersCod = p.cPersCod " & _
           " inner join rrhh r on d.cPersCod = r.cPersCod " & _
           " inner join agencias a on r.cAgenciaAsig = a.cAgeCod " & _
           " where d.cBSCod = '" & psBSCod & "' and d.nEstado = 1 and " & _
           " (select count(*) from LogProSelAprobacion where nEstadoAprobacion=0 and nProSelReqNro = d.nProSelReqNro) = 0 and nProSelNro=0 and d.nAnio = " & nAnio & _
           " group by p.cPersNombre, a.cAgeDescripcion "
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sSQL)
        oCon.CierraConexion
    End If
    Do While Not rs.EOF
        If CargarDatosR = "" Then
            CargarDatosR = Space(5) & JIZQ(rs!cAgeDescripcion, 40) & JIZQ(rs!cPersNombre, 40) & Space(5) & JDER(rs!nCantidad, 8)
        Else
            CargarDatosR = CargarDatosR & vbCrLf & Space(5) & JIZQ(rs!cAgeDescripcion, 40) & JIZQ(rs!cPersNombre, 40) & Space(10) & JDER(rs!nCantidad, 8)
        End If
        rs.MoveNext
    Loop
    Exit Function
CargarDatosRErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function
