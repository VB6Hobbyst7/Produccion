VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogPlanAnualValRef 
   ClientHeight    =   6075
   ClientLeft      =   1275
   ClientTop       =   2235
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   9120
   Begin TabDlg.SSTab sstLog 
      Height          =   5895
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   741
      TabCaption(0)   =   "Consolidación de Requerimientos - Plan Anual                   "
      TabPicture(0)   =   "frmLogPlanAnualValRef.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdGenPlan"
      Tab(0).Control(1)=   "cmdSalir"
      Tab(0).Control(2)=   "cmdActualizar"
      Tab(0).Control(3)=   "cmdConsolidar"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(5)=   "cboMoneda"
      Tab(0).Control(6)=   "txtAnio"
      Tab(0).Control(7)=   "txtEdit"
      Tab(0).Control(8)=   "MSValor"
      Tab(0).Control(9)=   "Frame2"
      Tab(0).Control(10)=   "cmdEjeVal"
      Tab(0).Control(11)=   "cmdGrabar"
      Tab(0).Control(12)=   "cmdCancelar"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Requerimientos Plan Anual - Consolidado por Areas                   "
      TabPicture(1)   =   "frmLogPlanAnualValRef.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "MSFlex"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cboAreas"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdVerArea"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdGenPlan 
         Caption         =   "Generar Plan Anual"
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
         Left            =   -69420
         TabIndex        =   20
         Top             =   5460
         Width           =   1935
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   -67440
         TabIndex        =   19
         Top             =   5460
         Width           =   1215
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   -73440
         TabIndex        =   18
         Top             =   5460
         Width           =   1395
      End
      Begin VB.CommandButton cmdConsolidar 
         Caption         =   "Consolidar"
         Height          =   375
         Left            =   -74880
         TabIndex        =   17
         Top             =   5460
         Width           =   1395
      End
      Begin VB.Frame Frame1 
         Height          =   675
         Left            =   -74880
         TabIndex        =   12
         Top             =   480
         Width           =   5115
         Begin VB.Label lblTitulo 
            AutoSize        =   -1  'True
            Caption         =   "Requerimientos consolidados al"
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
            Left            =   180
            TabIndex        =   16
            Top             =   300
            Width           =   2700
         End
      End
      Begin VB.ComboBox cboMoneda 
         BackColor       =   &H00EAFFFF&
         Height          =   315
         ItemData        =   "frmLogPlanAnualValRef.frx":0038
         Left            =   -69840
         List            =   "frmLogPlanAnualValRef.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2160
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtAnio 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   -71460
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "2000"
         Top             =   130
         Width           =   495
      End
      Begin VB.TextBox txtEdit 
         BackColor       =   &H00DDFFFE&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -67920
         TabIndex        =   6
         Top             =   2040
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdVerArea 
         Caption         =   "Ver consolidado"
         Height          =   375
         Left            =   7080
         TabIndex        =   5
         Top             =   675
         Width           =   1695
      End
      Begin VB.ComboBox cboAreas 
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   6195
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSValor 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   1
         Top             =   1200
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   7435
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   10
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   10
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   4395
         Left            =   120
         TabIndex        =   2
         Top             =   1260
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   7752
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   10
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Height          =   675
         Left            =   -69720
         TabIndex        =   8
         Top             =   480
         Width           =   3495
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00EAFFFF&
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Acumulado"
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
            Left            =   240
            TabIndex        =   10
            Top             =   300
            Width           =   1515
         End
      End
      Begin VB.CommandButton cmdEjeVal 
         Caption         =   "Valor Referencial"
         Height          =   375
         Left            =   -72000
         TabIndex        =   15
         Top             =   5460
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
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
         Left            =   -68760
         TabIndex        =   14
         Top             =   5460
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
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
         Left            =   -67440
         TabIndex        =   13
         Top             =   5460
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Area"
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
         TabIndex        =   4
         Top             =   780
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmLogPlanAnualValRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String

Dim EstaValorizando As Boolean
Dim EstaAgrupado As Boolean
Dim nFILA As Integer

Private Sub cmdGenPlan_Click()
Dim oConn As New DConecta, rs As New ADODB.Recordset
Dim nPlanAnualNro As Integer, nItem As Integer

If oConn.AbreConexion Then

   sSql = "UPDATE LogPlanAnual SET nPlanAnualEstado=0 WHERE nPlanAnualAnio = 2006"
   oConn.Ejecutar sSql
      
   sSql = "UPDATE LogPlanAnualDetalle SET nPlanAnualEstado=0 WHERE nPlanAnualAnio = 2006"
   oConn.Ejecutar sSql
      
   sSql = "insert into LogPlanAnual (nPlanAnualAnio) Values (2006)"
   oConn.Ejecutar sSql
      
   Set rs = oConn.CargaRecordSet("Select max(nPlanAnualNro) as nMaxNro from LogPlanAnual")
   If Not rs.EOF Then
      nPlanAnualNro = rs!nMaxNro
   End If
      
   If nPlanAnualNro > 0 Then

      sSql = " select " & nPlanAnualNro & ",2006 as nPlanAnualAnio,cBSGrupoCod=coalesce(g.cBSGrupoCod,'')," & _
             " cObjetoCod=left(p.cBSCod,2), cSintesis=g.cBSGrupoDescripcion, p.nMoneda, sum(p.nPrecioUnitario*p.nCantidad) as nValorEstimado " & _
             "  from LogPlanAnualValor p  left join BienesServicios b on p.cBSCod = b.cBSCod " & _
             "  left join BSGrupos g on b.cBSGrupoCod = g.cBSGrupoCod " & _
             " Where P.nAnio = 2006 And P.nEstado = 1 " & _
             " group by g.cBSGrupoCod,left(p.cBSCod,2),g.cBSGrupoDescripcion,p.nMoneda "
             
      Set rs = oConn.CargaRecordSet(sSql)
      If Not rs.EOF Then
         nItem = 0
         Do While Not rs.EOF
            nItem = nItem + 1
            sSql = "insert into LogPlanAnualDetalle (nPlanAnualNro,nPlanAnualItem,nPlanAnualAnio,cBSGrupoCod,cObjetoCod,cSintesis,nMoneda,nValorEstimado) " & _
                   " VALUES (" & nPlanAnualNro & "," & nItem & "," & rs!nPlanAnualAnio & ",'" & rs!cBSGrupoCod & "','" & rs!cObjetoCod & "','" & rs!cSintesis & "'," & rs!nMoneda & "," & rs!nValorEstimado & ") "
            oConn.Ejecutar sSql
            rs.MoveNext
         Loop
      End If
   End If
   
   MsgBox "Se ha generado el PLAN ANUAL !" + Space(10), vbInformation
   Unload Me
End If

End Sub

Private Sub Form_Load()
CentraForm Me
txtAnio = "2006"
sstLog.Tab = 0
EstaAgrupado = False
EstaValorizando = False
cboMoneda.Font = "Tahoma"
cboMoneda.Fontsize = 7
FormaFlexValor

If HayConsolidado Then
   If MsgBox("Ya existe un consolidado de requerimientos" + Space(10) + vbCrLf + _
             "       ¿ Recuperar el Consolidado ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
      VisualizaConsolidado
   End If
End If
'lblTitulo.Caption = "Requerimientos consolidados al " + CStr(gdFecSis)
End Sub

Private Sub cmdCancelar_Click()
lblTitulo.Caption = "Requerimientos consolidados al " + CStr(gdFecSis)
lblTitulo.ForeColor = "&H00800000"
cmdConsolidar.Visible = True
cmdActualizar.Visible = True
cmdGrabar.Visible = False
cmdCancelar.Visible = False
cmdSalir.Enabled = True
End Sub

Sub GeneraConsolidado(vMantenerValor As Boolean)
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim cMsg As String
Dim nPlanAnualNro As Integer

'GENERA UN CONSOLIDADO A PARTIR DE LOS REQUERIMIENTOS APROBADOS
If oConn.AbreConexion Then

   cMsg = "La consolidación creará una estructura del Plan Anual" + vbCrLf + _
          "  ¿ Está seguro de consolidar los requerimientos ?" + Space(10)
          
   If MsgBox(cMsg + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
      
      If vMantenerValor Then
         sSql = "Select cBSCod,nPrecioUnitario from LogPlanAnualValor where nAnio=2006"
         Set rs = oConn.CargaRecordSet(sSql)
      End If

      sSql = "DELETE from LogPlanAnualValor where nAnio=2006"
      oConn.Ejecutar sSql

      sSql = "insert into LogPlanAnualValor (nAnio,dFecha,cBSCod,nCantidad) " & _
      " select 2006,getdate(),cBSCod,sum(nMes01+nMes02+nMes03+nMes04+nMes05+nMes06+nMes07+nMes08+nMes09+nMes10+nMes11+nMes12) as nCantidad  " & _
      " From LogPlanAnualReqDetalle " & _
      " where nAnio=2006 and nPlanNro in (select nPlanNro from LogPlanAnualAprobacionPro where nNivelProceso=3 and nEstadoProceso = 1) " & _
      " group by cBSCod "
           
      oConn.Ejecutar sSql

      If vMantenerValor Then
         If Not rs.EOF Then
            Do While Not rs.EOF
               sSql = "UPDATE LogPlanAnualValor SET nPrecioUnitario = " & rs!nPrecioUnitario & " WHERE cBSCod = '" & rs!cBSCod & "'"
               oConn.Ejecutar sSql
               rs.MoveNext
            Loop
         End If
      End If
      
      oConn.CierraConexion
   End If
Else
      MsgBox "No se puede establecer la conexión..." + Space(10), vbInformation
      Exit Sub
End If
End Sub


'************************************************************************
Private Sub cmdConsolidar_Click()
GeneraConsolidado False
VisualizaConsolidado
End Sub

Sub VisualizaConsolidado()
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset, i As Integer

FormaFlexValor
sSql = "select c.*, u.cUnidad, b.cBSDescripcion from LogPlanAnualValor c inner join BienesServicios b on c.cBSCod=b.cBSCod " & _
       " inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad " & _
       "              from Constante where nConsCod = 9097) u on b.nBSUnidad = u.nBSUnidad " & _
       " where c.nAnio=2006 "

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSql)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSValor, i
         MSValor.TextMatrix(i, 1) = rs!cBSCod
         MSValor.TextMatrix(i, 2) = rs!cBSDescripcion
         MSValor.TextMatrix(i, 3) = rs!nCantidad
         MSValor.TextMatrix(i, 4) = rs!cUnidad
         MSValor.TextMatrix(i, 5) = IIf(rs!nMoneda = 2, "US$", "S/.")
         MSValor.TextMatrix(i, 6) = FNumero(rs!nPrecioUnitario)
         MSValor.TextMatrix(i, 7) = FNumero(rs!nCantidad * rs!nPrecioUnitario)
         MSValor.TextMatrix(i, 8) = rs!nMoneda
         rs.MoveNext
      Loop
      cmdEjeVal.Visible = True
      cmdSalir.Enabled = True
   End If
End If
End Sub

Sub FormaFlexValor()
MSValor.Clear
MSValor.Rows = 2
MSValor.RowHeight(-1) = 280
MSValor.RowHeight(0) = 320
MSValor.RowHeight(1) = 8
MSValor.ColWidth(0) = 0
MSValor.ColWidth(1) = 850:       MSValor.TextMatrix(0, 1) = "Codigo"
MSValor.ColWidth(2) = 3000:      MSValor.TextMatrix(0, 2) = "Descripción"
MSValor.ColWidth(3) = 800:       MSValor.TextMatrix(0, 3) = "Cantidad":     MSValor.ColAlignment(3) = 4
MSValor.ColWidth(4) = 1000:      MSValor.TextMatrix(0, 4) = " U. Medida":   MSValor.ColAlignment(4) = 4
MSValor.ColWidth(5) = 600:       MSValor.TextMatrix(0, 5) = "Moneda":       MSValor.ColAlignment(5) = 4
MSValor.ColWidth(6) = 1000:      MSValor.TextMatrix(0, 6) = " Precio Unit"
MSValor.ColWidth(7) = 1000:      MSValor.TextMatrix(0, 7) = "  Sub-Total"
End Sub

'Function HayConsolidado() As Boolean
'Dim oConn As New DConecta, rs as New ADODB.Recordset
'
'HayConsolidado = False
'sSQL = "Select top 1 * from LogPlanAnualValor where nEstado=1"
'If oConn.AbreConexion Then
'   Set rs = oConn.CargaRecordSet(sSQL)
'   If Not rs.EOF Then
'      HayConsolidado = True
'   End If
'End If
'End Function

Sub VerificaSiHay()
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset, i As Integer

'   sSQL = " Select top 1 x.nCantidad, coalesce(y.nCantidad,0) from " & _
'   " (select cBSCod,sum(nMes01+nMes02+nMes03+nMes04+nMes05+nMes06+nMes07+nMes08+nMes09+nMes10+nMes11+nMes12) as nCantidad " & _
'   " From LogPlanAnualReqDetalle " & _
'   " where nPlanNro in (select a.nPlanNro from LogPlanAnualAprobacion a inner join " & _
'   "   (select nPlanNro,max(nNivelAprobacion) as nMaxNivel from LogPlanAnualAprobacion group by nPlanNro) n on a.nPlanNro = n.nPlanNro and a.nNivelAprobacion = n.nMaxNivel " & _
'   "     where a.nEstadoAprobacion = 1) and nAnio=2006 group by cBSCod) x " & _
'   " left join LogPlanAnualValor y on x.cBSCod = y.cBSCod " & _
'   " Where x.nCantidad <> coalesce(y.nCantidad, 0) "
   sSql = ""
   If oConn.AbreConexion Then
      Set rs = oConn.CargaRecordSet(sSql)
      If Not rs.EOF Then
         MsgBox "Existen requerimientos aprobados por consolidar..." + Space(10), vbInformation
      End If
   End If
End Sub


Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   MSValor.SetFocus
End If
End Sub

Private Sub cboMoneda_LostFocus()
If cboMoneda.Visible Then
   MSValor = cboMoneda.Text
   cboMoneda.Visible = False
End If
End Sub

Function HayConsolidado() As Boolean
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset

HayConsolidado = False
sSql = "select top 1 cBSCod from LogPlanAnualValor where nAnio=2006 and nEstado=1 "
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSql)
   If Not rs.EOF Then
      HayConsolidado = True
   End If
End If
End Function


Private Sub cmdEjeVal_Click()
Dim i As Integer, n As Integer

lblTitulo.Caption = "Proceso de Valorización"
lblTitulo.ForeColor = "&H00000080"
EstaValorizando = True

n = MSValor.Rows - 1
For i = 1 To n
    MSValor.row = i
    MSValor.Col = 5:  MSValor.CellBackColor = "&H00EAFFFF"
    MSValor.Col = 6:  MSValor.CellBackColor = "&H00EAFFFF"
Next
sstLog.TabEnabled(1) = False
cmdEjeVal.Visible = False
cmdConsolidar.Visible = False
cmdActualizar.Visible = False
cmdGrabar.Visible = True
cmdGenPlan.Visible = False
cmdCancelar.Visible = True
cmdSalir.Caption = "Cancelar"
MSValor.Col = 5
MSValor.SetFocus
End Sub

Private Sub CmdSalir_Click()
If EstaValorizando Then
   txtAnio.Locked = False
   'lblValor.Visible = False
   sstLog.TabEnabled(1) = True
   cmdEjeVal.Visible = True
   cmdConsolidar.Visible = True
   cmdActualizar.Visible = True
   cmdGrabar.Visible = False
   cmdSalir.Caption = "Salir"
   EstaValorizando = False
Else
   Unload Me
End If
End Sub

Private Sub cmdActualizar_Click()
If MsgBox("¿ Mantener valorización del Consolidado Actual?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   GeneraConsolidado True
Else
   GeneraConsolidado False
End If
End Sub

Private Sub cmdGrabar_Click()
Dim oConn As New DConecta
Dim i As Integer, n As Integer, cPersCod As String
Dim nTipMon As Integer
Dim nSuma As Currency

n = MSValor.Rows - 1
If MsgBox("¿ Esta seguro de grabar la valorización de Requerimientos ? " + Space(10), vbQuestion + vbYesNo, "Confirme aprobación") = vbYes Then

   If oConn.AbreConexion Then
      nSuma = 0
      For i = 1 To n
          nTipMon = IIf(MSValor.TextMatrix(i, 5) = "S/.", 1, 2)
          sSql = "UPDATE LogPlanAnualValor SET nMoneda=" & nTipMon & ", nPrecioUnitario = " & VNumero(MSValor.TextMatrix(i, 6)) & ", dFechaVal = '" & Format(gdFecSis, "YYYYMMDD") & "' where nAnio = 2006 and cBSCod = '" & MSValor.TextMatrix(i, 1) & "' "
          nSuma = nSuma + VNumero(MSValor.TextMatrix(i, 7))
          oConn.Ejecutar sSql
      Next
      
      sSql = "update LogPlanAnualReq set nValorizado=1 " & _
             " where nPlanNro in (select a.nPlanNro from LogPlanAnualAprobacion a inner join  " & _
             " (select nPlanNro,max(nNivelAprobacion) as nMaxNivel from LogPlanAnualAprobacion group by nPlanNro) n on a.nPlanNro = n.nPlanNro and a.nNivelAprobacion = n.nMaxNivel  " & _
             " where a.nEstadoAprobacion = 1) and nAnio=2006 "
      oConn.Ejecutar sSql
      
   End If
   MsgBox "Se ha grabado la valorización del consolidado!" + Space(10), vbInformation
   Unload Me
End If
End Sub

'************************************************************************
Private Sub cboAreas_Click()
Dim cAreaCod As String
If cboAreas.ListIndex > 0 Then
   cAreaCod = Format(cboAreas.ItemData(cboAreas.ListIndex), "000")
   ConsolidaArea cAreaCod, 0
Else
   FormaFlex
End If
End Sub

Private Sub cmdVerArea_Click()
Dim cAreaCod As String
End Sub

Private Sub MSValor_DblClick()
If MSValor.Col = 5 And EstaValorizando Then
   cboMoneda.Visible = True
   cboMoneda.ListIndex = 0
   cboMoneda.Move MSValor.Left + MSValor.CellLeft - 30, MSValor.Top + MSValor.CellTop - 30, 640
End If
End Sub

Private Sub sstLog_Click(PreviousTab As Integer)
Select Case sstLog.Tab
    Case 0
    
    Case 1
         CargaAreas
End Select
End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

End If
End Sub

Sub CargaAreas()
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset

cboAreas.Clear
cboAreas.AddItem "Seleccione el Area ---------"

sSql = "select distinct cAreaCod=a.cRHAreaCodAprobacion, c.cAreaDescripcion " & _
       "  From LogPlanAnualAprobacion a inner join Areas c on a.cRHAreaCodAprobacion = c.cAreaCod " & _
       " Where a.nNivelAprobacion = 1 "
             
If oConn.AbreConexion Then
Set rs = oConn.CargaRecordSet(sSql)
If Not rs.EOF Then
   Do While Not rs.EOF
      cboAreas.AddItem rs!cAreaDescripcion
      cboAreas.ItemData(cboAreas.ListCount - 1) = rs!cAreaCod
      rs.MoveNext
   Loop
End If
End If
cboAreas.ListIndex = 0
End Sub

Private Sub ConsolidaTodas()
Dim oConn As New DConecta
Dim ra As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim i As Integer, cAreaCod As String
Dim k As Integer

i = 0
FormaFlex
If oConn.AbreConexion Then
   sSql = "select cAreaCod,cAreaDescripcion from DBCmactAux..Areas where cAreaCod in " & _
          " (select distinct cRHAreaCod from LogPlanAnualRequerimientos where nAnio=2006 and nTipo=2 and nEstado=1) "
             
   Set ra = oConn.CargaRecordSet(sSql)
   If Not ra.EOF Then
      Do While Not ra.EOF
         i = MSFlex.Rows - 1
         If i = 1 Then
            MSFlex.RowHeight(i) = 270
         Else
            i = i + 1
            InsRow MSFlex, i
         End If
         cAreaCod = ra!cAreaCod
         MSFlex.TextMatrix(i, 0) = ra!cAreaCod
         MSFlex.TextMatrix(i, 1) = ra!cAreaCod
         MSFlex.TextMatrix(i, 2) = ra!cAreaDescripcion
         MSFlex.row = i
         For k = 0 To 5
             MSFlex.Col = k
             MSFlex.CellBackColor = "&H00EAFFFF"
             MSFlex.CellForeColor = "&H00C00000"
             MSFlex.CellFontBold = True
         Next
         ConsolidaArea cAreaCod, i
         ra.MoveNext
      Loop
   End If
End If
End Sub

Private Sub ConsolidaArea(vAreaCod As String, nIndex As Integer)
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim i As Integer

If cboAreas.ListIndex > 0 Then
   FormaFlex
End If
i = MSFlex.Rows - 1

If oConn.AbreConexion Then
          
   sSql = "select p.cBSCod,b.cBSDescripcion,u.cUnidad, nCantidad=sum(nMes01+nMes02+nMes03+nMes04+nMes05+nMes06+nMes07+nMes08+nMes09+nMes10+nMes11+nMes12)  " & _
   " from LogPlanAnualReqDetalle p inner join BienesServicios b on p.cBSCod = b.cBSCod " & _
   " inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) u on b.nBSUnidad = u.nBSUnidad " & _
   " where p.nEstado=1 and p.nAnio=2006 and p.nPlanNro in " & _
   " (select nPlanNro From LogPlanAnualAprobacion where cRHAreaCodAprobacion='" & vAreaCod & "' and nNivelAprobacion=1) " & _
   " group by p.cBSCod,b.cBSDescripcion,u.cUnidad"


   Set rs = oConn.CargaRecordSet(sSql)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 0) = vAreaCod
         MSFlex.TextMatrix(i, 1) = rs!cBSCod
         MSFlex.TextMatrix(i, 2) = rs!cBSDescripcion
         MSFlex.TextMatrix(i, 3) = rs!nCantidad
         MSFlex.TextMatrix(i, 4) = rs!cUnidad
         rs.MoveNext
      Loop
    End If
End If
End Sub

Sub FormaFlex()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(-1) = 280
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 850:       MSFlex.TextMatrix(0, 1) = "Codigo"
MSFlex.ColWidth(2) = 4500:      MSFlex.TextMatrix(0, 2) = "Descripción"
MSFlex.ColWidth(3) = 800:       MSFlex.TextMatrix(0, 3) = " Cantidad":   MSFlex.ColAlignment(3) = 4
MSFlex.ColWidth(4) = 1200:      MSFlex.TextMatrix(0, 4) = "   U. Medida":   MSFlex.ColAlignment(4) = 4
MSFlex.ColWidth(5) = 0
End Sub

'*********************************************************************
'PROCEDIMIENTOS DEL FLEX
'*********************************************************************

Private Sub MSValor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And MSValor.Col = 6 And EstaValorizando Then
   MSValor.TextMatrix(MSValor.row, 6) = ""
   SumaTotal
End If

'If KeyCode = vbKeyInsert Then
'   If EstaAgrupado Then
'      frmACAsignaProceso.Inicio MSValor.TextMatrix(MSValor.row, 3), 1, MSValor.TextMatrix(MSValor.row, 5)
'   End If
'End If
End Sub

Private Sub MSValor_KeyPress(KeyAscii As Integer)
If MSValor.Col = 6 And EstaValorizando Then
   EditaFlex MSValor, txtEdit, KeyAscii
End If
If MSValor.Col = 5 And EstaValorizando Then
   nFILA = MSValor.row
   cboMoneda.Move MSValor.Left + MSValor.CellLeft - 30, MSValor.Top + MSValor.CellTop - 30, 640
   cboMoneda.Visible = True
   cboMoneda.SetFocus
End If
End Sub

Sub EditaFlex(MSValor As Control, Edt As Control, KeyAscii As Integer)
Select Case KeyAscii
    Case 0 To 32
         Edt = MSValor
         Edt.SelStart = 1000
    Case Else
         Edt = Chr(KeyAscii)
         Edt.SelStart = 1
End Select
Edt.Move MSValor.Left + MSValor.CellLeft - 15, MSValor.Top + MSValor.CellTop - 15, _
         MSValor.CellWidth, MSValor.CellHeight
Edt.Visible = True
Edt.SetFocus
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
'nKeyAscii = KeyAscii
'KeyAscii = DigNumDec(txtEdit, KeyAscii)
If KeyAscii = Asc(vbCr) Then
   KeyAscii = 0
   'txtEdit = FNumero(txtEdit)
End If
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode MSValor, txtEdit, KeyCode, Shift
End Sub

Sub EditKeyCode(MSValor As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27
         Edt.Visible = False
         MSValor.SetFocus
    Case 13
         MSValor.SetFocus
    Case 37                     'Izquierda
         MSValor.SetFocus
         DoEvents
         If MSValor.Col > 1 Then
            MSValor.Col = MSValor.Col - 1
         End If
    Case 39                     'Derecha
         MSValor.SetFocus
         DoEvents
         If MSValor.Col < MSValor.Cols - 1 Then
            MSValor.Col = MSValor.Col + 1
         End If
    Case 38
         MSValor.SetFocus
         DoEvents
         If MSValor.row > MSValor.FixedRows + 1 Then
            MSValor.row = MSValor.row - 1
         End If
    Case 40
         MSValor.SetFocus
         DoEvents
         If MSValor.row < MSValor.Rows - 1 Then
            MSValor.row = MSValor.row + 1
         End If
End Select
End Sub

Private Sub MSValor_GotFocus()
If cboMoneda.Visible Then
   MSValor.TextMatrix(MSValor.row, 5) = cboMoneda.Text
   cboMoneda.Visible = False
End If
If txtEdit.Visible = False Then Exit Sub
MSValor = txtEdit
txtEdit.Visible = False
SumaTotal
'If MSValor.Row < MSValor.Rows - 1 Then
'   MSValor.Row = MSValor.Row + 1
'End If
End Sub

Private Sub MSValor_LeaveCell()
If cboMoneda.Visible Then
   MSValor.TextMatrix(MSValor.row, 5) = cboMoneda.Text
   cboMoneda.Visible = False
End If
If txtEdit.Visible = False Then Exit Sub
MSValor = txtEdit
txtEdit.Visible = False
SumaTotal
'If MSValor.Row < MSValor.Rows - 1 Then
'   MSValor.Row = MSValor.Row + 1
'End If
End Sub

Sub SumaTotal()
Dim i As Integer, n As Integer, nSuma As Currency

n = MSValor.Rows - 1
nSuma = 0
For i = 1 To n
    MSValor.TextMatrix(i, 7) = FNumero(VNumero(MSValor.TextMatrix(i, 6)) * VNumero(MSValor.TextMatrix(i, 3)))
    If VNumero(MSValor.TextMatrix(i, 6)) > 0 Then
       nSuma = nSuma + VNumero(MSValor.TextMatrix(i, 7))
    Else
       MSValor.TextMatrix(i, 7) = ""
    End If
Next
txtTotal = FNumero(nSuma)
End Sub


