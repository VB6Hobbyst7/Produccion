VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogPlanAnualReqConsolida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consolidación de Requerimientos"
   ClientHeight    =   6030
   ClientLeft      =   135
   ClientTop       =   1995
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   11715
   Begin TabDlg.SSTab sstReq 
      Height          =   3675
      Left            =   60
      TabIndex        =   18
      Top             =   1860
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   6482
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   5
      TabHeight       =   564
      TabCaption(0)   =   "Consolidado General                 "
      TabPicture(0)   =   "frmLogPlanAnualReqConsolida.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSFlex"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   " Detalle por Bien / Servicio      "
      TabPicture(1)   =   "frmLogPlanAnualReqConsolida.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSDet"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "cmdExcluyeReq(1)"
      Tab(1).Control(3)=   "cmdObservaReq(1)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Detalle por Usuario                 "
      TabPicture(2)   =   "frmLogPlanAnualReqConsolida.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSUsu"
      Tab(2).Control(1)=   "MSReq"
      Tab(2).Control(2)=   "cmdExcluyeReq(2)"
      Tab(2).Control(3)=   "cmdObservaReq(2)"
      Tab(2).ControlCount=   4
      Begin VB.CommandButton cmdObservaReq 
         Height          =   375
         Index           =   2
         Left            =   -65340
         TabIndex        =   32
         Top             =   900
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdExcluyeReq 
         Caption         =   "Excluir Requerimiento"
         Height          =   375
         Index           =   2
         Left            =   -65340
         TabIndex        =   31
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdObservaReq 
         Height          =   375
         Index           =   1
         Left            =   -72780
         TabIndex        =   30
         Top             =   3210
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdExcluyeReq 
         Caption         =   "Excluir Requerimiento"
         Height          =   375
         Index           =   1
         Left            =   -74880
         TabIndex        =   29
         Top             =   3210
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74880
         TabIndex        =   23
         Top             =   420
         Width           =   11355
         Begin VB.TextBox txtUnidad 
            BackColor       =   &H00ECFFEF&
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   9360
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   300
            Width           =   1695
         End
         Begin VB.TextBox txtBSDescripcion 
            BackColor       =   &H00ECFFEF&
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   300
            Width           =   6195
         End
         Begin VB.TextBox txtBSCod 
            BackColor       =   &H00ECFFEF&
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   180
            TabIndex        =   28
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Unidad"
            Height          =   195
            Left            =   8760
            TabIndex        =   27
            Top             =   360
            Width           =   510
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   3045
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   5371
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   18
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   14151167
         ForeColorSel    =   128
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
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
         _Band(0).Cols   =   18
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSReq 
         Height          =   1185
         Left            =   -74880
         TabIndex        =   20
         Top             =   480
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   2090
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   7
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   16775393
         ForeColorSel    =   8388608
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
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
         _Band(0).Cols   =   7
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSUsu 
         Height          =   1785
         Left            =   -74880
         TabIndex        =   21
         Top             =   1740
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   3149
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   18
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   16775393
         ForeColorSel    =   8388608
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
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
         _Band(0).Cols   =   18
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDet 
         Height          =   2025
         Left            =   -74880
         TabIndex        =   22
         Top             =   1140
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   3572
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   17
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   15532015
         ForeColorSel    =   1596177
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
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
         _Band(0).Cols   =   17
      End
   End
   Begin VB.CommandButton cmdConsolida 
      Caption         =   "Consolidar Requerimientos"
      Height          =   375
      Left            =   60
      TabIndex        =   16
      Top             =   5580
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   6735
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
         Left            =   5700
         TabIndex        =   0
         Text            =   "2005"
         Top             =   90
         Width           =   675
      End
      Begin VB.Label lblPlan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plan Anual de Adquisiciones y Contrataciones"
         BeginProperty Font 
            Name            =   "Helvetica"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   5310
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00EAFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   6735
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   11595
      Begin VB.CommandButton cmdPersona 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2325
         TabIndex        =   14
         Top             =   330
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtPersona 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   300
         Width           =   8655
      End
      Begin VB.TextBox txtCargo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   10455
      End
      Begin VB.TextBox txtArea 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   900
         Width           =   6375
      End
      Begin VB.TextBox txtAgencia 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8100
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   900
         Width           =   3255
      End
      Begin VB.TextBox txtPersCod 
         Height          =   300
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agencia"
         Height          =   195
         Left            =   7440
         TabIndex        =   17
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   660
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Area"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   960
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdAprobar 
      Caption         =   "Aprobar Consolidado"
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
      Left            =   7920
      TabIndex        =   2
      Top             =   5580
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10440
      TabIndex        =   1
      Top             =   5580
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Plan del año"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   720
      Width           =   1080
   End
   Begin VB.Menu mnuConsol 
      Caption         =   "MenuConsol"
      Visible         =   0   'False
      Begin VB.Menu mnuConsolidado 
         Caption         =   "Consolidado total del Area"
      End
      Begin VB.Menu mnuUsuario 
         Caption         =   "Requerimientos por Usuario"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBSCod 
         Caption         =   "Detalle de requerimientos por Usuario"
      End
   End
End
Attribute VB_Name = "frmLogPlanAnualReqConsolida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMes(1 To 12) As String, sSQL As String
Dim cRHAgeCod As String, cRHAreaCod As String, cRHCargoCod As String
Dim cRangosIN As String, nNivelAprobacion As Integer, bHaExcluido As Boolean

Private Sub cmdExcluyeReq_Click(Index As Integer)
Dim nPlanReqNro As Long
Dim cTipoDesc As String
Dim oConn As New DConecta
Dim cUsuNom As String

bHaExcluido = False
cTipoDesc = ""
If Index = 1 Then
   If Len(Trim(MSDet.TextMatrix(MSDet.row, 1))) > 0 Then
      nPlanReqNro = MSDet.TextMatrix(MSDet.row, 1)
      cUsuNom = MSDet.TextMatrix(MSDet.row, 3)
      cTipoDesc = "Se excluirá el requerimiento de :" + vbCrLf + Space(5) + UCase(txtBSDescripcion) + Space(10) + vbCrLf + "del usuario :" + vbCrLf + Space(5) + cUsuNom + Space(10) + vbCrLf + vbCrLf + Space(10) + "¿ Desea continuar ?"
   Else
      MsgBox "No es un Requerimiento válido..." + Space(10), vbInformation
      Exit Sub
   End If
End If

If Index = 2 Then
   If Len(Trim(MSReq.TextMatrix(MSReq.row, 1))) > 0 Then
      nPlanReqNro = MSReq.TextMatrix(MSReq.row, 1)
      cTipoDesc = "Se excluirá TODOS los requerimientos de: " + Space(10) + vbCrLf + MSReq.TextMatrix(MSReq.row, 3) + Space(10) + vbCrLf + vbCrLf + Space(10) + "¿ Desea continuar ?"
   Else
      MsgBox "No es un Requerimiento válido..." + Space(10), vbInformation
      Exit Sub
   End If
End If

If MsgBox(cTipoDesc, vbQuestion + vbYesNo + vbDefaultButton2, "Confirme") = vbYes Then
   bHaExcluido = True
   If oConn.AbreConexion Then
      If Index = 1 Then
         sSQL = "UPDATE LogPlanAnualReqDetalle SET nEstado = 0 WHERE nPlanReqNro = " & nPlanReqNro & " and cBSCod = '" & txtBSCod.Text & "'"
         oConn.Ejecutar sSQL
         cmdConsolida_Click
         DetalleBienServicio MSFlex.TextMatrix(MSFlex.row, 1)
         MSDet.SetFocus
      End If
      If Index = 2 Then
         sSQL = "UPDATE LogPlanAnualReq SET nEstado = 3 WHERE nPlanReqNro = " & nPlanReqNro & " "
         oConn.Ejecutar sSQL
         sSQL = "UPDATE LogPlanAnualReqDetalle SET nEstado = 0 WHERE nPlanReqNro = " & nPlanReqNro & " "
         oConn.Ejecutar sSQL
         cmdConsolida_Click
      End If
      oConn.CierraConexion
   End If
End If
End Sub

Private Sub Form_Load()
CentraForm Me
cRHAgeCod = ""
cRHAreaCod = ""
cRHCargoCod = ""
txtAgencia.Text = ""
txtCargo.Text = ""
txtArea.Text = ""
bHaExcluido = False
txtAnio = Year(gdFecSis) + 1
mMes(1) = "ENE"
mMes(2) = "FEB"
mMes(3) = "MAR"
mMes(4) = "ABR"
mMes(5) = "MAY"
mMes(6) = "JUN"
mMes(7) = "JUL"
mMes(8) = "AGO"
mMes(9) = "SEP"
mMes(10) = "OCT"
mMes(11) = "NOV"
mMes(12) = "DIC"
FormaFlex
FormaFlexReq
If gsCodArea = "036" Or gsCodUser = "SIST" Then
   cmdPersona.Visible = True
End If
sstReq.Tab = 0
txtPersCod.Text = gsCodPersUser
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub FormaFlexReq()
MSReq.Clear
MSReq.Rows = 2
MSReq.RowHeight(0) = 300
MSReq.RowHeight(1) = 8
MSReq.ColWidth(0) = 0
MSReq.ColWidth(1) = 0
MSReq.ColWidth(2) = 980:  MSReq.TextMatrix(0, 2) = "Código"
MSReq.ColWidth(3) = 3000: MSReq.TextMatrix(0, 3) = "Usuario"
MSReq.ColWidth(4) = 2000: MSReq.TextMatrix(0, 4) = "Cargo"
MSReq.ColWidth(5) = 2000: MSReq.TextMatrix(0, 5) = "Agencia"
MSReq.ColWidth(6) = 1200: MSReq.TextMatrix(0, 6) = "Estado"
End Sub

Private Sub cmdConsolida_Click()
Dim Rs As New ADODB.Recordset
Dim oConn As New DConecta, i As Integer
Dim YaAprobado As Boolean
Dim HayPendientes As Boolean
Dim nConsolidaAgencia As Boolean
Dim nNivelVerificacion As Integer

FormaFlex
FormaFlexReq
cRangosIN = ""
cmdConsolida.Visible = False
YaAprobado = False
sstReq.Tab = 0
nNivelVerificacion = 0
HayPendientes = False
nConsolidaAgencia = False

'--- NIVEL DE APROBACION DEL CARGO ACTUAL ---------------------------

sSQL = "select distinct nNivelAprobacion,nAgencia " & _
       "  from LogNivelAprobacion where cRHCargoCodAprobacion = '" & cRHCargoCod & "' " & _
       "    "
       
If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
   If Not Rs.EOF Then
      nConsolidaAgencia = Rs!nAgencia
      nNivelAprobacion = Rs!nNivelAprobacion
   Else
      nNivelAprobacion = 1
   End If
End If

If nNivelAprobacion > 1 Then
   nNivelVerificacion = nNivelAprobacion - 1
End If

'--- REQUERIMIENTOS QUE SERAN APROBADOS POR EL CARGO ACTUAL ---------
'gcEstadosRPA
'---------------------------------------------------------------------
'CONSOLIDACION DE REQUERIMIENTOS ACTIVOS - CABECERA REQUERIMIENTOS
'---------------------------------------------------------------------
If nNivelAprobacion <= 1 Then
   If nConsolidaAgencia Then
      sSQL = "select distinct a.nPlanReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' ')," & _
        "         t.cEstado, g.cAgeDescripcion, a.nEstadoAprobacion, c.cRHCargoDescripcion " & _
        "  from LogPlanAnualAprobacion a " & _
        "       inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro " & _
        "       inner join Persona p on r.cPersCod = p.cPersCod " & _
        "       inner join Agencias g on r.cRHAgeCod = g.cAgeCod " & _
        "       inner join RHCargosTabla c on r.cRHCargoCod = c.cRHCargoCod " & _
        "       inner join (select nConsValor as nEstadoAprobacion,cConsDescripcion as cEstado from Constante where nConsCod = " & gcEstadosRPA & " and nConsCod<>nConsValor ) t on a.nEstadoAprobacion = t.nEstadoAprobacion " & _
        " where a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and r.cRHAgeCod = '" & cRHAgeCod & "' and " & _
        "       r.nEstado=" & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & " "
   Else
      sSQL = "select a.nPlanReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' '), " & _
        "         t.cEstado, g.cAgeDescripcion, a.nEstadoAprobacion, c.cRHCargoDescripcion " & _
        "  from LogPlanAnualAprobacion a  " & _
        "       inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro " & _
        "       inner join Persona p on r.cPersCod = p.cPersCod " & _
        "       inner join Agencias g on r.cRHAgeCod = g.cAgeCod " & _
        "       inner join RHCargosTabla c on r.cRHCargoCod = c.cRHCargoCod " & _
        "       inner join (select nConsValor as nEstadoAprobacion,cConsDescripcion as cEstado from Constante where nConsCod = " & gcEstadosRPA & " and nConsCod<>nConsValor ) t on a.nEstadoAprobacion = t.nEstadoAprobacion " & _
        " where a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and " & _
        "       r.nEstado=" & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & " order by  r.cRHAgeCod,p.cPersNombre"
   End If
Else
   If nConsolidaAgencia Then
      sSQL = "select distinct a.nPlanReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' ')," & _
        "         t.cEstado, g.cAgeDescripcion, a.nEstadoAprobacion, e.nEstadoAprobacion as nEstadoNivelAnt, c.cRHCargoDescripcion " & _
        "  from LogPlanAnualAprobacion a " & _
        "       inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro " & _
        "       inner join Persona p on r.cPersCod = p.cPersCod " & _
        "       inner join Agencias g on r.cRHAgeCod = g.cAgeCod " & _
        "       inner join RHCargosTabla c on r.cRHCargoCod = c.cRHCargoCod " & _
        "       inner join (select nPlanReqNro, nEstadoAprobacion from LogPlanAnualAprobacion where nPlanReqNro in (select a.nPlanReqNro from LogPlanAnualAprobacion a " & _
        "                    inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro  where a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and r.nEstado = " & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & ") and nNivelAprobacion = " & nNivelVerificacion & ") e on a.nPlanReqNro = e.nPlanReqNro " & _
        "       inner join (select nConsValor as nEstadoAprobacion,cConsDescripcion as cEstado from Constante where nConsCod = " & gcEstadosRPA & " and nConsCod<>nConsValor ) t on e.nEstadoAprobacion = t.nEstadoAprobacion " & _
        " where a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and r.cRHAgeCod = '" & cRHAgeCod & "' and " & _
        "       r.nEstado=" & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & " "
   Else
      sSQL = "select distinct a.nPlanReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' '), " & _
        "            t.cEstado, g.cAgeDescripcion, a.nEstadoAprobacion, e.nEstadoAprobacion as nEstadoNivelAnt, c.cRHCargoDescripcion " & _
        "  from LogPlanAnualAprobacion a  " & _
        "       inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro " & _
        "       inner join Persona p on r.cPersCod = p.cPersCod " & _
        "       inner join Agencias g on r.cRHAgeCod = g.cAgeCod " & _
        "       inner join RHCargosTabla c on r.cRHCargoCod = c.cRHCargoCod " & _
        "       inner join (select nPlanReqNro, nEstadoAprobacion from LogPlanAnualAprobacion where nPlanReqNro in (select a.nPlanReqNro from LogPlanAnualAprobacion a " & _
        "                    inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro  where a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and r.nEstado = " & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & ") and nNivelAprobacion = " & nNivelVerificacion & ") e on a.nPlanReqNro = e.nPlanReqNro " & _
        "       inner join (select nConsValor as nEstadoAprobacion,cConsDescripcion as cEstado from Constante where nConsCod = " & gcEstadosRPA & " and nConsCod<>nConsValor ) t on e.nEstadoAprobacion = t.nEstadoAprobacion " & _
        " where a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and " & _
        "       r.nEstado = " & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & " order by  r.cRHAgeCod,p.cPersNombre"
   End If
End If

'and a.nEstadoAprobacion=0
If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
End If

If Rs.State = 0 Then
   MsgBox ""
Else
   If Not Rs.EOF Then
      i = 0
      Do While Not Rs.EOF
         i = i + 1
         InsRow MSReq, i
         MSReq.TextMatrix(i, 1) = Rs!nPlanReqNro
         cRangosIN = cRangosIN + MSReq.TextMatrix(i, 1) + ","
         MSReq.TextMatrix(i, 2) = Rs!cPersCod
         MSReq.TextMatrix(i, 3) = Rs!cPersona
         MSReq.TextMatrix(i, 4) = Rs!cRHCargoDescripcion
         MSReq.TextMatrix(i, 5) = Rs!cAgeDescripcion
         MSReq.TextMatrix(i, 6) = Rs!cEstado
         
         If Rs!nEstadoAprobacion = 1 Then
            YaAprobado = True
         End If
         
         If nNivelAprobacion > 1 Then
            If Rs!nEstadoNivelAnt = 0 Then
               HayPendientes = True
            End If
         End If
         Rs.MoveNext
      Loop
      cRangosIN = Left(cRangosIN, Len(cRangosIN) - 1)
   End If
End If

If Len(cRangosIN) = 0 Then
   MsgBox "No hay requerimientos por aprobar en este Nivel..." + Space(10), vbInformation
   Exit Sub
End If

'---------------------------------------------------------------------
'CONSOLIDACION DE DETALLE DE REQUERIMIENTOS - SOLO DETALLE ACTIVO
'---------------------------------------------------------------------
sSQL = "select d.cBSCod, b.cBSDescripcion,u.cUnidad, nMes01=sum(d.nMes01), nMes02=sum(d.nMes02),nMes03=sum(d.nMes03), " & _
       "       nMes04=sum(d.nMes04),nMes05=sum(d.nMes05),nMes06=sum(d.nMes06),nMes07=sum(d.nMes07), " & _
       "       nMes08=sum(d.nMes08),nMes09=sum(d.nMes09),nMes10=sum(d.nMes10),nMes11=sum(d.nMes11),nMes12=sum(d.nMes12) " & _
       "         " & _
       "  from LogPlanAnualReqDetalle d inner join BienesServicios b on d.cBSCod = b.cBSCod " & _
       "       inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) u on b.nBSUnidad = u.nBSUnidad " & _
       " where d.nEstado = " & gcActivo & " and d.nPlanReqNro in (" + cRangosIN + ") group by d.cBSCod, b.cBSDescripcion,u.cUnidad "

If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
End If
If Not Rs.EOF Then
   i = 0
      Do While Not Rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 1) = Rs!cBSCod
         MSFlex.TextMatrix(i, 2) = Rs!cBSDescripcion
         MSFlex.TextMatrix(i, 3) = Rs!cUnidad
         MSFlex.TextMatrix(i, 4) = IIf(Rs!nMes01 > 0, Rs!nMes01, "")
         MSFlex.TextMatrix(i, 5) = IIf(Rs!nMes02 > 0, Rs!nMes02, "")
         MSFlex.TextMatrix(i, 6) = IIf(Rs!nMes03 > 0, Rs!nMes03, "")
         MSFlex.TextMatrix(i, 7) = IIf(Rs!nMes04 > 0, Rs!nMes04, "")
         MSFlex.TextMatrix(i, 8) = IIf(Rs!nMes05 > 0, Rs!nMes05, "")
         MSFlex.TextMatrix(i, 9) = IIf(Rs!nMes06 > 0, Rs!nMes06, "")
         MSFlex.TextMatrix(i, 10) = IIf(Rs!nMes07 > 0, Rs!nMes07, "")
         MSFlex.TextMatrix(i, 11) = IIf(Rs!nMes08 > 0, Rs!nMes08, "")
         MSFlex.TextMatrix(i, 12) = IIf(Rs!nMes09 > 0, Rs!nMes09, "")
         MSFlex.TextMatrix(i, 13) = IIf(Rs!nMes10 > 0, Rs!nMes10, "")
         MSFlex.TextMatrix(i, 14) = IIf(Rs!nMes11 > 0, Rs!nMes11, "")
         MSFlex.TextMatrix(i, 15) = IIf(Rs!nMes12 > 0, Rs!nMes12, "")
         MSFlex.TextMatrix(i, 16) = Rs!nMes01 + Rs!nMes02 + Rs!nMes03 + Rs!nMes04 + Rs!nMes05 + Rs!nMes06 + Rs!nMes07 + Rs!nMes08 + Rs!nMes09 + Rs!nMes10 + Rs!nMes11 + Rs!nMes12
         Rs.MoveNext
      Loop
End If

'-------------------------------------------------------------
'Para un estado NORMAL ---------------------------------------
'-------------------------------------------------------------
cmdAprobar.Visible = True
cmdSalir.FontBold = True
cmdConsolida.Visible = True

If YaAprobado Then
   If bHaExcluido Then
      bHaExcluido = False
      cmdAprobar.Visible = True
      cmdSalir.FontBold = True
      cmdConsolida.Visible = True
      MSFlex.SetFocus
      Exit Sub
   End If
   
   If MsgBox("El consolidado ya fue aprobado..." + Space(10) + vbCrLf + " ¿ Desea consolidar nuevamente ? ", vbQuestion + vbYesNo, "") = vbNo Then
      cmdAprobar.Visible = False
      cmdSalir.FontBold = False
      cmdConsolida.Visible = False
   End If
Else
   If HayPendientes Then
      MsgBox "Existen requerimientos pendientes de aprobación en el Nivel anterior..." + Space(10), vbInformation
      cmdAprobar.Visible = False
      cmdSalir.FontBold = True
      cmdConsolida.Visible = True
   End If
End If
MSFlex.SetFocus
End Sub

Sub FormaFlex()
Dim i As Integer
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(-1) = 280
MSFlex.RowHeight(0) = 300
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 0:         MSFlex.TextMatrix(0, 1) = ""
MSFlex.ColWidth(2) = 2500:      MSFlex.TextMatrix(0, 2) = "Descripción"
MSFlex.ColWidth(3) = 1200:      MSFlex.TextMatrix(0, 3) = "   U. Medida":   MSFlex.ColAlignment(3) = 4
For i = 1 To 12
    MSFlex.TextMatrix(0, i + 3) = Space(2) + mMes(i)
    MSFlex.ColWidth(i + 3) = 540:   MSFlex.ColAlignment(i + 3) = 4
Next
MSFlex.ColWidth(17) = 0
MSFlex.ColWidth(18) = 0
End Sub

Private Sub cmdAprobar_Click()
Dim oConn As New DConecta

sstReq.Tab = 0
If MsgBox("¿ Esta seguro de aprobar el Consolidado de Requerimientos ? " + Space(10), vbQuestion + vbYesNo + vbDefaultButton2, "Confirme aprobación") = vbYes Then
   
   If oConn.AbreConexion Then
      sSQL = "UPDATE LogPlanAnualAprobacion Set nEstadoAprobacion = 1 " & _
             " WHERE nPlanReqNro IN (" & cRangosIN & ")  and nNivelAprobacion = " & nNivelAprobacion & " "
      oConn.Ejecutar sSQL
   End If
   
   MsgBox "El requerimiento consolidado del área ha sido APROBADO con éxito!" + Space(10), vbInformation
   Unload Me
End If
End Sub


Private Sub sstReq_Click(PreviousTab As Integer)
Select Case sstReq.Tab
    Case 0
         txtBSCod.Text = ""
         txtBSDescripcion.Text = ""
    Case 1
         txtBSCod.Text = MSFlex.TextMatrix(MSFlex.row, 1)
         txtBSDescripcion.Text = MSFlex.TextMatrix(MSFlex.row, 2)
         txtUnidad.Text = MSFlex.TextMatrix(MSFlex.row, 3)
         DetalleBienServicio MSFlex.TextMatrix(MSFlex.row, 1)
    Case 2
         MSReq.SetFocus
End Select
End Sub

Private Sub cmdPersona_Click()
Dim X As UPersona
Set X = frmBuscaPersona.Inicio

If X Is Nothing Then
    Exit Sub
End If

If Len(Trim(X.sPersNombre)) > 0 Then
   txtPersona.Text = X.sPersNombre
   txtPersCod = X.sPersCod
End If
'frmBuscaPersona.Show 1
'If frmBuscaPersona.vpOK Then
'   txtPersona.Text = frmBuscaPersona.vpPersNom
'   txtPersCod = frmBuscaPersona.vpPersCod
'End If
End Sub


Private Sub txtAnio_GotFocus()
SelTexto txtAnio
End Sub

Private Sub txtanio_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   txtPersCod.SetFocus
End If
End Sub

Private Sub txtPersCod_Change()
If Len(txtPersCod) > 0 Then
   DatosPersonal txtPersCod
End If
End Sub

Sub DatosPersonal(vPersCod As String)
Dim Rs As New ADODB.Recordset
Dim rn As New ADODB.Recordset
Dim oConn As DConecta
Dim i As Integer

FormaFlex
FormaFlexReq
cRHAgeCod = ""
cRHAreaCod = ""
cRHCargoCod = ""
txtAgencia.Text = ""
txtCargo.Text = ""
txtArea.Text = ""

sSQL = "select x.*,p.cPersNombre,cCargo=coalesce(c.cRHCargoDescripcion,''),cArea=coalesce(a.cAreaDescripcion,''),cAgencia=coalesce(b.cAgeDescripcion,'') " & _
" from (select top 1 cPersCod, cRHCargoCodOficial as cRHCargoCod, cRHAreaCodOficial as cAreaCod, cRHAgenciaCodOficial as cAgeCod " & _
"  from RHCargos where cPersCod='" & vPersCod & "' order by dRHCargoFecha desc) x " & _
"  left outer join Persona p on x.cPersCod = p.cPersCod " & _
"  left outer join Areas a on x.cAreaCod = a.cAreaCod " & _
"  left outer join Agencias b on x.cAgeCod = b.cAgeCod " & _
"  left outer join RHCargosTabla c on x.cRHCargoCod = c.cRHCargoCod "
 
Set oConn = New DConecta
If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
   If Not Rs.EOF Then
      cRHAgeCod = Rs!cAgeCod
      cRHAreaCod = Rs!cAreaCod
      cRHCargoCod = Rs!cRHCargoCod
      txtPersona.Text = Rs!cPersNombre
      txtAgencia.Text = Rs!cAgencia
      txtCargo.Text = Rs!cCargo
      txtArea.Text = Rs!cArea
      sstReq.Tab = 0
      cmdConsolida.Visible = True
   End If
End If
End Sub


'-----------------------------------------------------------------------
Private Sub MSReq_GotFocus()
If Len(MSReq.TextMatrix(MSReq.row, 1)) > 0 Then
   DetalleUsuario MSReq.TextMatrix(MSReq.row, 1)
End If
End Sub

Private Sub MSReq_RowColChange()
If Len(MSReq.TextMatrix(MSReq.row, 1)) > 0 Then
   DetalleUsuario MSReq.TextMatrix(MSReq.row, 1)
End If
End Sub

'------------------------------------------------------------------------------
' DETALLE POR USUARIO
'------------------------------------------------------------------------------
Private Sub DetalleUsuario(ByVal pnPlanNro As Integer)
Dim Rs As New ADODB.Recordset
Dim oPlan As New DLogPlanAnual
Dim i As Integer

MSUsu.Clear
MSUsu.Rows = 2
MSUsu.RowHeight(0) = 290
MSUsu.RowHeight(1) = 8
MSUsu.ColWidth(0) = 0
MSUsu.ColWidth(1) = 0
MSUsu.ColWidth(1) = 0:            MSUsu.TextMatrix(0, 1) = ""
MSUsu.ColWidth(2) = 2500:         MSUsu.TextMatrix(0, 2) = "Descripción"
MSUsu.ColWidth(3) = 1200:         MSUsu.TextMatrix(0, 3) = "   U. Medida":   MSUsu.ColAlignment(3) = 4
For i = 1 To 12
    MSUsu.TextMatrix(0, i + 3) = Space(2) + mMes(i)
    MSUsu.ColWidth(i + 3) = 540:  MSUsu.ColAlignment(i + 3) = 4
Next
MSUsu.ColWidth(17) = 0
MSUsu.ColWidth(18) = 0

Set Rs = oPlan.RequerimientoPlanAnual(pnPlanNro, True)
If Not Rs.EOF Then
   i = 0
   Do While Not Rs.EOF
      i = i + 1
      InsRow MSUsu, i
      MSUsu.TextMatrix(i, 1) = Rs!cBSCod
      MSUsu.TextMatrix(i, 2) = Rs!cBSDescripcion
      MSUsu.TextMatrix(i, 3) = Rs!cUnidad
      MSUsu.TextMatrix(i, 4) = IIf(Rs!nMes01 > 0, Rs!nMes01, "")
      MSUsu.TextMatrix(i, 5) = IIf(Rs!nMes02 > 0, Rs!nMes02, "")
      MSUsu.TextMatrix(i, 6) = IIf(Rs!nMes03 > 0, Rs!nMes03, "")
      MSUsu.TextMatrix(i, 7) = IIf(Rs!nMes04 > 0, Rs!nMes04, "")
      MSUsu.TextMatrix(i, 8) = IIf(Rs!nMes05 > 0, Rs!nMes05, "")
      MSUsu.TextMatrix(i, 9) = IIf(Rs!nMes06 > 0, Rs!nMes06, "")
      MSUsu.TextMatrix(i, 10) = IIf(Rs!nMes07 > 0, Rs!nMes07, "")
      MSUsu.TextMatrix(i, 11) = IIf(Rs!nMes08 > 0, Rs!nMes08, "")
      MSUsu.TextMatrix(i, 12) = IIf(Rs!nMes09 > 0, Rs!nMes09, "")
      MSUsu.TextMatrix(i, 13) = IIf(Rs!nMes10 > 0, Rs!nMes10, "")
      MSUsu.TextMatrix(i, 14) = IIf(Rs!nMes11 > 0, Rs!nMes11, "")
      MSUsu.TextMatrix(i, 15) = IIf(Rs!nMes12 > 0, Rs!nMes12, "")
      MSUsu.TextMatrix(i, 16) = Rs!nTotal
      Rs.MoveNext
   Loop
End If
End Sub

'------------------------------------------------------------------------------
' DETALLE POR BIEN/SERVICIO
'------------------------------------------------------------------------------
Private Sub DetalleBienServicio(psBSCod As String)
Dim Rs As New ADODB.Recordset
Dim oConn As New DConecta, i As Integer

MSDet.Clear
MSDet.Rows = 2
MSDet.RowHeight(-1) = 280
MSDet.RowHeight(0) = 300
MSDet.RowHeight(1) = 8
MSDet.ColWidth(0) = 0
MSDet.ColWidth(1) = 0
MSDet.ColWidth(2) = 0
MSDet.ColWidth(3) = 3700:    MSDet.TextMatrix(0, 2) = "Descripción"
For i = 1 To 12
    MSDet.TextMatrix(0, i + 3) = Space(2) + mMes(i)
    MSDet.ColWidth(i + 3) = 540:   MSDet.ColAlignment(i + 3) = 4
Next
MSDet.ColWidth(17) = 0
MSDet.ColWidth(18) = 0
'--------------------------------------------------------------------------

If Len(Trim(cRangosIN)) = 0 Then Exit Sub

sSQL = "select r.nPlanReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' '), d.* " & _
"  from LogPlanAnualReqDetalle d inner join LogPlanAnualReq r on d.nPlanReqNro = r.nPlanReqNro " & _
"       inner join Persona p on r.cPersCod = p.cPersCod  " & _
" where r.nPlanReqNro in (" & cRangosIN & ") and  d.cBSCod = '" & psBSCod & "' and d.nEstado=1 "

If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      i = 0
      Do While Not Rs.EOF
         i = i + 1
         InsRow MSDet, i
         MSDet.TextMatrix(i, 1) = Rs!nPlanReqNro
         MSDet.TextMatrix(i, 2) = Rs!cPersCod
         MSDet.TextMatrix(i, 3) = Rs!cPersona
         MSDet.TextMatrix(i, 4) = IIf(Rs!nMes01 > 0, Rs!nMes01, "")
         MSDet.TextMatrix(i, 5) = IIf(Rs!nMes02 > 0, Rs!nMes02, "")
         MSDet.TextMatrix(i, 6) = IIf(Rs!nMes03 > 0, Rs!nMes03, "")
         MSDet.TextMatrix(i, 7) = IIf(Rs!nMes04 > 0, Rs!nMes04, "")
         MSDet.TextMatrix(i, 8) = IIf(Rs!nMes05 > 0, Rs!nMes05, "")
         MSDet.TextMatrix(i, 9) = IIf(Rs!nMes06 > 0, Rs!nMes06, "")
         MSDet.TextMatrix(i, 10) = IIf(Rs!nMes07 > 0, Rs!nMes07, "")
         MSDet.TextMatrix(i, 11) = IIf(Rs!nMes08 > 0, Rs!nMes08, "")
         MSDet.TextMatrix(i, 12) = IIf(Rs!nMes09 > 0, Rs!nMes09, "")
         MSDet.TextMatrix(i, 13) = IIf(Rs!nMes10 > 0, Rs!nMes10, "")
         MSDet.TextMatrix(i, 14) = IIf(Rs!nMes11 > 0, Rs!nMes11, "")
         MSDet.TextMatrix(i, 15) = IIf(Rs!nMes12 > 0, Rs!nMes12, "")
         MSDet.TextMatrix(i, 16) = Rs!nMes01 + Rs!nMes02 + Rs!nMes03 + Rs!nMes04 + Rs!nMes05 + Rs!nMes06 + Rs!nMes07 + Rs!nMes08 + Rs!nMes09 + Rs!nMes10 + Rs!nMes11 + Rs!nMes12
         Rs.MoveNext
      Loop
   End If
   'Totaliza
   'lblTitulo.Caption = "Detalle del requerimiento: " + UCase(cBSDesc)
   'lblTitulo.Visible = True
   'cmdAprobar.Visible = False
End If
End Sub

Private Sub mnuConsolidado_Click()
cmdConsolida_Click
End Sub


'*******************************************************************************
Private Sub MSFlex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu mnuConsol
End If
End Sub

