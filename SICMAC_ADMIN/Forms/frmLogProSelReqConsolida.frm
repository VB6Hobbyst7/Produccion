VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogProSelReqConsolida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consolidación de Requerimientos"
   ClientHeight    =   5490
   ClientLeft      =   825
   ClientTop       =   2100
   ClientWidth     =   7665
   Icon            =   "frmLogProSelReqConsolida.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7665
   Begin VB.TextBox txtEdit 
      BackColor       =   &H00D7EDFF&
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   4800
      MaxLength       =   7
      TabIndex        =   16
      Top             =   2850
      Visible         =   0   'False
      Width           =   560
   End
   Begin VB.CommandButton cmdConsolida 
      Caption         =   "Consolidar Requerimientos"
      Height          =   375
      Left            =   60
      TabIndex        =   13
      Top             =   4980
      Width           =   2295
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
      Height          =   1680
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   7515
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
         Height          =   240
         Left            =   2550
         TabIndex        =   11
         Top             =   330
         Width           =   315
      End
      Begin VB.TextBox txtPersona 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   300
         Width           =   4215
      End
      Begin VB.TextBox txtCargo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   6015
      End
      Begin VB.TextBox txtArea 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   900
         Width           =   6015
      End
      Begin VB.TextBox txtAgencia 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   6015
      End
      Begin VB.TextBox txtPersCod 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agencia"
         Height          =   180
         Left            =   420
         TabIndex        =   14
         Top             =   1260
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   420
         TabIndex        =   9
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
         Height          =   195
         Left            =   420
         TabIndex        =   8
         Top             =   660
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Area"
         Height          =   195
         Left            =   420
         TabIndex        =   7
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
      Left            =   3480
      TabIndex        =   1
      Top             =   4980
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   4980
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   3045
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   5371
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   11
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
      _Band(0).Cols   =   11
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
      TabIndex        =   10
      Top             =   720
      Width           =   1080
   End
   Begin VB.Menu mnuConsol 
      Caption         =   "MenuConsol"
      Visible         =   0   'False
      Begin VB.Menu mnuConsolidado 
         Caption         =   "Consolidado total del Area"
      End
      Begin VB.Menu mnuQuitarReq 
         Caption         =   "Quitar Requerimiento"
      End
   End
End
Attribute VB_Name = "frmLogProSelReqConsolida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMes(1 To 12) As String, sSQL As String
Dim cRHAgeCod As String, cRHAreaCod As String, cRHCargoCod As String
Dim cRangosIN As String, nNivelAprobacion As Integer, bHaExcluido As Boolean

'Private Sub cmdExcluyeReq_Click(Index As Integer)
'Dim nPlanReqNro As Long
'Dim cTipoDesc As String
'Dim oConn As New DConecta
'Dim cUsuNom As String
'
'bHaExcluido = False
'cTipoDesc = ""
'If Index = 1 Then
'   If Len(Trim(MSDet.TextMatrix(MSDet.Row, 1))) > 0 Then
'      nPlanReqNro = MSDet.TextMatrix(MSDet.Row, 1)
'      cUsuNom = MSDet.TextMatrix(MSDet.Row, 3)
'      cTipoDesc = "Se excluirá el requerimiento de :" + vbCrLf + Space(5) + UCase(txtBSDescripcion) + Space(10) + vbCrLf + "del usuario :" + vbCrLf + Space(5) + cUsuNom + Space(10) + vbCrLf + vbCrLf + Space(10) + "¿ Desea continuar ?"
'   Else
'      MsgBox "No es un Requerimiento válido..." + Space(10), vbInformation
'      Exit Sub
'   End If
'End If
'
'If Index = 2 Then
'   If Len(Trim(MSReq.TextMatrix(MSReq.Row, 1))) > 0 Then
'      nPlanReqNro = MSReq.TextMatrix(MSReq.Row, 1)
'      cTipoDesc = "Se excluirá TODOS los requerimientos de: " + Space(10) + vbCrLf + MSReq.TextMatrix(MSReq.Row, 3) + Space(10) + vbCrLf + vbCrLf + Space(10) + "¿ Desea continuar ?"
'   Else
'      MsgBox "No es un Requerimiento válido..." + Space(10), vbInformation
'      Exit Sub
'   End If
'End If
'
'If MsgBox(cTipoDesc, vbQuestion + vbYesNo + vbDefaultButton2, "Confirme") = vbYes Then
'   bHaExcluido = True
'   If oConn.AbreConexion Then
'      If Index = 1 Then
'         sSQL = "UPDATE LogPlanAnualReqDetalle SET nEstado = 0 WHERE nPlanReqNro = " & nPlanReqNro & " and cProSelBSCod = '" & txtBSCod.Text & "'"
'         oConn.Ejecutar sSQL
'         cmdConsolida_Click
'         DetalleBienServicio MSFlex.TextMatrix(MSFlex.Row, 1)
'         MSDet.SetFocus
'      End If
'      If Index = 2 Then
'         sSQL = "UPDATE LogPlanAnualReq SET nEstado = 3 WHERE nPlanReqNro = " & nPlanReqNro & " "
'         oConn.Ejecutar sSQL
'         sSQL = "UPDATE LogPlanAnualReqDetalle SET nEstado = 0 WHERE nPlanReqNro = " & nPlanReqNro & " "
'         oConn.Ejecutar sSQL
'         cmdConsolida_Click
'      End If
'      oConn.CierraConexion
'   End If
'End If
'End Sub

Private Sub Form_Load()
Dim oAcceso As UAcceso
Set oAcceso = New UAcceso
CentraForm Me
cRHAgeCod = ""
cRHAreaCod = ""
cRHCargoCod = ""
txtAgencia.Text = ""
txtCargo.Text = ""
txtArea.Text = ""
bHaExcluido = False
'txtPersCod.Text = UCase(oAcceso.ObtenerUsuario)
txtPersCod.Text = gsCodPersUser
'txtAnio = Year(gdFecSis) + 1
'mMes(1) = "ENE"
'mMes(2) = "FEB"
'mMes(3) = "MAR"
'mMes(4) = "ABR"
'mMes(5) = "MAY"
'mMes(6) = "JUN"
'mMes(7) = "JUL"
'mMes(8) = "AGO"
'mMes(9) = "SEP"
'mMes(10) = "OCT"
'mMes(11) = "NOV"
'mMes(12) = "DIC"
FormaFlex
'FormaFlexReq
'sstReq.Tab = 0
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

'Sub FormaFlexReq()
'MSReq.Clear
'MSReq.Rows = 2
'MSReq.RowHeight(0) = 300
'MSReq.RowHeight(1) = 8
'MSReq.ColWidth(0) = 0
'MSReq.ColWidth(1) = 0
'MSReq.ColWidth(2) = 980:  MSReq.TextMatrix(0, 2) = "Código"
'MSReq.ColWidth(3) = 2000: MSReq.TextMatrix(0, 3) = "Usuario"
'MSReq.ColWidth(4) = 1500: MSReq.TextMatrix(0, 4) = "Cargo"
'MSReq.ColWidth(5) = 1500: MSReq.TextMatrix(0, 5) = "Agencia"
'MSReq.ColWidth(6) = 1000: MSReq.TextMatrix(0, 6) = "Estado"
'End Sub

Private Sub cmdConsolida_Click()
Dim Rs As New ADODB.Recordset
Dim oConn As New DConecta, i As Integer
Dim YaAprobado As Boolean
Dim HayPendientes As Boolean
Dim nConsolidaAgencia As Boolean
Dim nNivelVerificacion As Integer
Dim sBSCod As String, nPapa As Integer

FormaFlex
'FormaFlexReq
cRangosIN = ""
cmdConsolida.Visible = False
YaAprobado = False
'sstReq.Tab = 0
nNivelVerificacion = 0
HayPendientes = False
nConsolidaAgencia = False

'--- NIVEL DE APROBACION DEL CARGO ACTUAL ---------------------------

'sSQL = "select distinct nNivelAprobacion,nAgencia " & _
       "  from LogNivelAprobacion where cRHCargoCodAprobacion = '" & cRHCargoCod & "' " & _
       "    "
sSQL = "select distinct nNivelaprobacion  from LogProSelAprobacion where cRHCargoCodAprobacion = '" & cRHCargoCod & "' "
       
If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
   If Not Rs.EOF Then
'      nConsolidaAgencia = Rs!nAgencia
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
'      sSQL = "select a.nPlanReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' ')," & _
        "         t.cEstado, g.cAgeDescripcion, a.nEstadoAprobacion, c.cRHCargoDescripcion " & _
        "  from LogPlanAnualAprobacion a " & _
        "       inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro " & _
        "       inner join Persona p on r.cPersCod = p.cPersCod " & _
        "       inner join Agencias g on r.cRHAgeCod = g.cAgeCod " & _
        "       inner join RHCargosTabla c on r.cRHCargoCod = c.cRHCargoCod " & _
        "       inner join (select nConsValor as nEstadoAprobacion,cConsDescripcion as cEstado from Constante where nConsCod = " & gcEstadosRPA & " and nConsCod<>nConsValor ) t on a.nEstadoAprobacion = t.nEstadoAprobacion " & _
        " where a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and r.cRHAgeCod = '" & cRHAgeCod & "' and " & _
        "       r.nEstado=" & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & " "
       sSQL = "select a.nProSelReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' ')," & _
        "         t.cEstado, g.cAgeDescripcion, a.nEstadoAprobacion, c.cRHCargoDescripcion " & _
        "  from LogProSelAprobacion a " & _
        "       inner join LogProSelReq r on a.nProSelReqNro = r.nProSelReqNro " & _
        "       inner join Persona p on r.cPersCod = p.cPersCod " & _
        "       inner join Agencias g on r.cRHAgeCod = g.cAgeCod " & _
        "       inner join RHCargosTabla c on r.cRHCargoCod = c.cRHCargoCod " & _
        "       inner join (select nConsValor as nEstadoAprobacion,cConsDescripcion as cEstado from Constante where nConsCod = " & gcEstadosRPA & " and nConsCod<>nConsValor ) t on a.nEstadoAprobacion = t.nEstadoAprobacion " & _
        " where a.nEstadoAprobacion=0 and a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and r.cRHAgeCod = '" & cRHAgeCod & "' and " & _
        "       r.nEstado=" & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & " "
   Else
'      sSQL = "select a.nPlanReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' '), " & _
        "         t.cEstado, g.cAgeDescripcion, a.nEstadoAprobacion, c.cRHCargoDescripcion " & _
        "  from LogPlanAnualAprobacion a  " & _
        "       inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro " & _
        "       inner join Persona p on r.cPersCod = p.cPersCod " & _
        "       inner join Agencias g on r.cRHAgeCod = g.cAgeCod " & _
        "       inner join RHCargosTabla c on r.cRHCargoCod = c.cRHCargoCod " & _
        "       inner join (select nConsValor as nEstadoAprobacion,cConsDescripcion as cEstado from Constante where nConsCod = " & gcEstadosRPA & " and nConsCod<>nConsValor ) t on a.nEstadoAprobacion = t.nEstadoAprobacion " & _
        " where a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and " & _
        "       r.nEstado=" & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & " order by  r.cRHAgeCod,p.cPersNombre"
        sSQL = "select a.nProSelReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' '),          " & _
               " t.cEstado , g.cAgeDescripcion, a.nEstadoAprobacion, c.cRHCargoDescripcion " & _
               " from LogProSelAprobacion a " & _
               " inner join LogProSelReq r on a.nProSelReqNro = r.nProSelReqNro " & _
               " inner join Persona p on r.cPersCod = p.cPersCod " & _
               " inner join Agencias g on r.cRHAgeCod = g.cAgeCod " & _
               " inner join RHCargosTabla c on r.cRHCargoCod = c.cRHCargoCod " & _
               " inner join (select nConsValor as nEstadoAprobacion,cConsDescripcion as cEstado " & _
               " from Constante where nConsCod = " & gcEstadosRPA & " and nConsCod<>nConsValor ) t " & _
               " on a.nEstadoAprobacion = t.nEstadoAprobacion " & _
               " where a.nEstadoAprobacion=0 and a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and r.nEstado = " & gcActivo & " And a.nNivelAprobacion = " & nNivelAprobacion & _
               " order by  r.cRHAgeCod,p.cPersNombre"
   End If
Else
   If nConsolidaAgencia Then
      'sSQL = "select a.nPlanReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' ')," & _
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
    sSQL = "select a.nProSelReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' '), " & _
               " t.cEstado, g.cAgeDescripcion, a.nEstadoAprobacion, e.nEstadoAprobacion as nEstadoNivelAnt, c.cRHCargoDescripcion " & _
               " from LogProSelAprobacion a " & _
               " inner join LogProSelReq r on a.nProSelReqNro = r.nProSelReqNro " & _
               " inner join Persona p on r.cPersCod = p.cPersCod " & _
               " inner join Agencias g on r.cRHAgeCod = g.cAgeCod " & _
               " inner join RHCargosTabla c on r.cRHCargoCod = c.cRHCargoCod " & _
               " inner join (select nProSelReqNro, nEstadoAprobacion from LogProSelAprobacion where nProSelReqNro in (select a.nProSelReqNro from LogProSelAprobacion a " & _
               " inner join LogProSelReq r on a.nProSelReqNro = r.nProSelReqNro  where a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and r.nEstado = " & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & ") and nNivelAprobacion = " & nNivelVerificacion & ") e on a.nProSelReqNro = e.nProSelReqNro " & _
               " inner join (select nConsValor as nEstadoAprobacion,cConsDescripcion as cEstado from Constante where nConsCod = " & gcEstadosRPA & " and nConsCod<>nConsValor ) t on e.nEstadoAprobacion = t.nEstadoAprobacion " & _
               " where a.nEstadoAprobacion=0 and a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and r.cRHAgeCod = '" & cRHAgeCod & "' and " & _
               " r.nEstado=" & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & " "
   Else
      'sSQL = "select a.nPlanReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' '), " & _
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
        
        
        'sSQL = "select a.nProSelReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' '), " & _
               " t.cEstado, g.cAgeDescripcion, a.nEstadoAprobacion, e.nEstadoAprobacion as nEstadoNivelAnt, c.cRHCargoDescripcion " & _
               " from LogProSelAprobacion a " & _
               " inner join LogProSelReq r on a.nProSelReqNro = r.nProSelReqNro " & _
               " inner join Persona p on r.cPersCod = p.cPersCod " & _
               " inner join Agencias g on r.cRHAgeCod = g.cAgeCod " & _
               " inner join RHCargosTabla c on r.cRHCargoCod = c.cRHCargoCod " & _
               " inner join (select nProSelReqNro, nEstadoAprobacion from LogProSelAprobacion where nProSelReqNro in ( " & _
               " select a.nProSelReqNro from LogProSelAprobacion a " & _
               " inner join LogProSelReq r on a.nProSelReqNro = r.nProSelReqNro " & _
               " where a.nEstadoAprobacion=0 and a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and r.nEstado = " & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & ") and nNivelAprobacion = " & nNivelVerificacion & ") e on a.nProSelReqNro = e.nProSelReqNro " & _
               " inner join (select nConsValor as nEstadoAprobacion,cConsDescripcion as cEstado from Constante where nConsCod = " & gcEstadosRPA & " and nConsCod<>nConsValor ) t on " & _
               " e.nEstadoAprobacion = t.nEstadoAprobacion  where a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and " & _
               " r.nEstado = " & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & " order by  r.cRHAgeCod,p.cPersNombre"
        
        sSQL = "select a.nProSelReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' '), " & _
               " g.cAgeDescripcion, a.nEstadoAprobacion, c.cRHCargoDescripcion " & _
               " from LogProSelAprobacion a " & _
               " inner join LogProSelReq r on a.nProSelReqNro = r.nProSelReqNro " & _
               " inner join Persona p on r.cPersCod = p.cPersCod " & _
               " inner join Agencias g on r.cRHAgeCod = g.cAgeCod " & _
               " inner join RHCargosTabla c on r.cRHCargoCod = c.cRHCargoCod " & _
               " where a.cRHCargoCodAprobacion = '" & cRHCargoCod & "' and " & _
               " r.nEstado > " & gcActivo & " AND a.nNivelAprobacion = " & nNivelAprobacion & " and a.nEstadoAprobacion = 0 order by  r.cRHAgeCod,p.cPersNombre "
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
'         i = i + 1
'         InsRow MSReq, i
'         MSReq.TextMatrix(i, 1) = Rs!nProSelReqNro
         cRangosIN = cRangosIN + CStr(Rs!nProSelReqNro) + ","
'         MSReq.TextMatrix(i, 2) = Rs!cPersCod
'         MSReq.TextMatrix(i, 3) = Rs!cPersona
'         MSReq.TextMatrix(i, 4) = Rs!cRHCargoDescripcion
'         MSReq.TextMatrix(i, 5) = Rs!cAgeDescripcion
'         MSReq.TextMatrix(i, 6) = Rs!cEstado

         If Rs!nEstadoAprobacion = 1 Then
            YaAprobado = True
         End If

'         If nNivelAprobacion > 1 Then
'            If Rs!nEstadoNivelAnt = 0 Then
'               HayPendientes = True
'            End If
'         End If
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
'sSQL = "select d.cProSelBSCod, b.cBSDescripcion,u.cUnidad, nMes01=sum(d.nMes01), nMes02=sum(d.nMes02),nMes03=sum(d.nMes03), " & _
       "       nMes04=sum(d.nMes04),nMes05=sum(d.nMes05),nMes06=sum(d.nMes06),nMes07=sum(d.nMes07), " & _
       "       nMes08=sum(d.nMes08),nMes09=sum(d.nMes09),nMes10=sum(d.nMes10),nMes11=sum(d.nMes11),nMes12=sum(d.nMes12) " & _
       "         " & _
       "  from LogPlanAnualReqDetalle d inner join LogProSelBienesServicios b on d.cProSelBSCod = b.cProSelBSCod " & _
       "       inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) u on b.nBSUnidad = u.nBSUnidad " & _
       " where d.nEstado = " & gcActivo & " and d.nPlanReqNro in (" + cRangosIN + ") group by d.cProSelBSCod, b.cBSDescripcion,u.cUnidad "
'sSQL = "select d.cBSCod, b.cBSDescripcion,u.cUnidad, nCantidad=sum(nCantidad), x.nMesEje, x.nAnio from LogProSelReqDetalle d " & _
       " inner join LogProSelReq x on d.nProSelReqNro = x.nProSelReqNro " & _
       " inner join LogProSelBienesServicios b on d.cBSCod = b.cProSelBSCod " & _
       " inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) u on b.nBSUnidad = u.nBSUnidad " & _
       " where d.nEstado = 1 and d.nProSelReqNro in (" + cRangosIN + ") " & _
       " group by d.cBSCod, b.cBSDescripcion,u.cUnidad, x.nMesEje, x.nAnio"
sSQL = "select d.cBSCod, b.cBSDescripcion,cUnidad=coalesce(u.cUnidad,''), nCantidad, x.nMesEje, x.nAnio, p.cPersNombre, p.cPersCod, d.nProSelReqNro  from LogProSelReqDetalle d " & _
       " inner join LogProSelReq x on d.nProSelReqNro = x.nProSelReqNro " & _
       " inner join LogProSelBienesServicios b on d.cBSCod = b.cProSelBSCod " & _
       " left join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) u on b.nBSUnidad = u.nBSUnidad " & _
       " inner join Persona p on d.cPersCod = p.cPersCod " & _
       " where d.nEstado = 1 and d.nProSelReqNro in (" + cRangosIN + ") " & _
       " order by d.cBSCod, b.cBSDescripcion, p.cPersNombre "

If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
End If
If Not Rs.EOF Then
   i = 0
      Do While Not Rs.EOF
         If sBSCod <> Rs!cBSCod Then
            i = i + 1
            InsRow MSFlex, i
            MSFlex.Col = 2
            MSFlex.row = i
            MSFlex.RowHeight(i) = 300
            MSFlex.CellFontSize = 10
            MSFlex.CellFontBold = True
            MSFlex.TextMatrix(i, 10) = Rs!cBSCod
            MSFlex.TextMatrix(i, 2) = "+"
            MSFlex.TextMatrix(i, 3) = Rs!cBSDescripcion
            MSFlex.TextMatrix(i, 4) = Rs!cUnidad
            MSFlex.TextMatrix(i, 6) = "0.00"
            MSFlex.TextMatrix(i, 7) = Format("01/" & Rs!nMesEje & "/" & Rs!nAnio, "mmm")
            sBSCod = Rs!cBSCod
            nPapa = i
         End If
        i = i + 1
        InsRow MSFlex, i
        MSFlex.RowHeight(i) = 0
        MSFlex.TextMatrix(i, 0) = 1
        MSFlex.TextMatrix(i, 1) = nPapa
        MSFlex.TextMatrix(i, 3) = Rs!cPersNombre
        MSFlex.TextMatrix(i, 4) = Rs!cUnidad
        MSFlex.TextMatrix(i, 5) = Rs!nCantidad
        MSFlex.TextMatrix(nPapa, 5) = Val(MSFlex.TextMatrix(nPapa, 5)) + Rs!nCantidad
        MSFlex.TextMatrix(i, 6) = Format(Rs!nCantidad * CargarValorRef(Rs!cBSCod), "#,##0.00")
        MSFlex.TextMatrix(nPapa, 6) = Format(CDbl(MSFlex.TextMatrix(nPapa, 6)) + (Rs!nCantidad * CargarValorRef(Rs!cBSCod)), "#,##0.00")
        MSFlex.TextMatrix(i, 8) = Rs!nProSelReqNro
        MSFlex.TextMatrix(i, 9) = Rs!cPersCod
        MSFlex.TextMatrix(i, 10) = Rs!cBSCod
        Rs.MoveNext
      Loop
End If

'-------------------------------------------------------------
'Para un estado NORMAL ---------------------------------------
'-------------------------------------------------------------
cmdAprobar.Visible = True
cmdSalir.FontBold = True
cmdConsolida.Visible = True

'If YaAprobado Then
'   If bHaExcluido Then
'      bHaExcluido = False
'      cmdAprobar.Visible = True
'      cmdSalir.FontBold = True
'      cmdConsolida.Visible = True
'      MSFlex.SetFocus
'      Exit Sub
'   End If
'
'   If MsgBox("El consolidado ya fue aprobado..." + Space(10) + vbCrLf + " ¿ Desea consolidar nuevamente ? ", vbQuestion + vbYesNo, "") = vbNo Then
'      cmdAprobar.Visible = False
'      cmdSalir.FontBold = False
'      cmdConsolida.Visible = False
'   End If
'Else
'   If HayPendientes Then
'      MsgBox "Existen requerimientos pendientes de aprobación en el Nivel anterior..." + Space(10), vbInformation
'      cmdAprobar.Visible = False
'      cmdSalir.FontBold = True
'      cmdConsolida.Visible = True
'   End If
'End If
MSFlex.SetFocus
End Sub

Sub FormaFlex()
Dim i As Integer
MSFlex.Clear
MSFlex.Rows = 2 ': MSFlex.Cols = 8
MSFlex.RowHeight(-1) = 280
MSFlex.RowHeight(0) = 300
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 0:         MSFlex.TextMatrix(0, 1) = ""
MSFlex.ColWidth(2) = 300:         MSFlex.TextMatrix(0, 2) = "":             MSFlex.ColAlignment(2) = 4
MSFlex.ColWidth(3) = 3500:      MSFlex.TextMatrix(0, 3) = "Descripción"
MSFlex.ColWidth(4) = 1200:      MSFlex.TextMatrix(0, 4) = "   U. Medida":   MSFlex.ColAlignment(4) = 4
MSFlex.ColWidth(5) = 800:      MSFlex.TextMatrix(0, 5) = "Cantidad"
MSFlex.ColWidth(6) = 1000:      MSFlex.TextMatrix(0, 6) = "P. Ref"
MSFlex.ColWidth(7) = 1000:      MSFlex.TextMatrix(0, 7) = "Mes"
MSFlex.ColWidth(8) = 0:      MSFlex.TextMatrix(0, 8) = "NroReq"
MSFlex.ColWidth(9) = 0:      MSFlex.TextMatrix(0, 9) = "cPersCod"
MSFlex.ColWidth(10) = 0:      MSFlex.TextMatrix(0, 10) = ""
End Sub

Private Sub cmdAprobar_Click()
Dim oConn As New DConecta

'sstReq.Tab = 0
If MsgBox("¿ Esta seguro de aprobar el Consolidado de Requerimientos ? " + Space(10), vbQuestion + vbYesNo + vbDefaultButton2, "Confirme aprobación") = vbYes Then
   Modificarequerimeintos
   If oConn.AbreConexion Then
      sSQL = "UPDATE LogProSelReq Set nEstado = nEstado + 1 " & _
             " WHERE nProSelReqNro IN (" & cRangosIN & ") "
      oConn.Ejecutar sSQL
      sSQL = "UPDATE LogProSelAprobacion Set nEstadoAprobacion = 1 " & _
             " WHERE nProSelReqNro IN (" & cRangosIN & ")  and nNivelAprobacion = " & nNivelAprobacion & " "
      oConn.Ejecutar sSQL
   End If
   
   MsgBox "El requerimiento consolidado del área ha sido APROBADO con éxito!" + Space(10), vbInformation
   Unload Me
End If
End Sub

Private Sub Modificarequerimeintos()
    Dim oCon As DConecta, sSQL As String, i As Integer
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        With MSFlex
            Do While i < .Rows
                If Val(.TextMatrix(i, 0)) = -1 Then
                    sSQL = "update LogProSelReqDetalle set nEstado = 0 " & _
                          " where nProSelReqNro = " & .TextMatrix(i, 8) & " and cPersCod ='" & .TextMatrix(i, 9) & "' and cBSCod = '" & .TextMatrix(i, 10) & "'"
                    oCon.Ejecutar sSQL
                    sSQL = "INSERT INTO LogProSelReqDetalle (nProSelReqNro,cPersCod,nAnio, nItem, cBSCod,nCantidad) " & _
                          " VALUES (" & .TextMatrix(i, 8) & ",'" & .TextMatrix(i, 9) & "'," & Year(gdFecSis) & "," & i & ",'" & MSFlex.TextMatrix(i, 10) & "'," & MSFlex.TextMatrix(i, 5) & ")"
                    oCon.Ejecutar sSQL
                End If
                i = i + 1
            Loop
        End With
        oCon.CierraConexion
    End If
End Sub

Private Sub QuitaRequerimeintos()
On Error GoTo QuitaRequerimeintosErr
    Dim oCon As DConecta, sSQL As String, i As Integer, j As Integer
    Set oCon = New DConecta
    If MSFlex.TextMatrix(MSFlex.row, 2) <> "" Then Exit Sub
    If MsgBox("¿Seguro que Desea Eliminar el Requerimiento?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    If oCon.AbreConexion Then
        With MSFlex
            i = .row
            sSQL = "update LogProSelReqDetalle set nEstado = 0 " & _
                  " where nProSelReqNro = " & .TextMatrix(i, 8) & " and cPersCod ='" & .TextMatrix(i, 9) & "' and cBSCod = '" & .TextMatrix(i, 10) & "'"
            oCon.Ejecutar sSQL
            
            cmdConsolida_Click
'            Do While i < .Rows
'                j = 0
'                Do While j < .Cols
'                    .Col = 2
'                    .row = i
'                    .CellFontSize = 10
'                    .CellFontBold = True
'                    If (i + 1) > .Rows - 1 Then
'                        .TextMatrix(i, j) = ""
'                    Else
'                        .TextMatrix(i, j) = .TextMatrix(i + 1, j)
'                    End If
'                    j = j + 1
'                Loop
'                If .TextMatrix(i, 2) = "" Then
'                    .RowHeight(i) = 0
'                Else
'                    .RowHeight(i) = 300
'                End If
'                i = i + 1
'            Loop
'            .Rows = .Rows - 1
        End With
        oCon.CierraConexion
    End If
    MsgBox "Requerimiento Eliminado", vbInformation, "Aviso"
    Exit Sub
QuitaRequerimeintosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

'Private Sub sstReq_Click(PreviousTab As Integer)
'Select Case sstReq.Tab
'    Case 0
'         txtBSCod.Text = ""
'         txtBSDescripcion.Text = ""
'    Case 1
'         txtBSCod.Text = MSFlex.TextMatrix(MSFlex.Row, 1)
'         txtBSDescripcion.Text = MSFlex.TextMatrix(MSFlex.Row, 2)
'         txtUnidad.Text = MSFlex.TextMatrix(MSFlex.Row, 3)
'         DetalleBienServicio MSFlex.TextMatrix(MSFlex.Row, 1)
'    Case 2
'         MSReq.SetFocus
'End Select
'End Sub

Private Sub cmdPersona_Click()
Dim X As UPersona
Set X = frmBuscaPersona.Inicio(True)

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


'Private Sub txtAnio_GotFocus()
'SelTexto txtanio
'End Sub

Private Sub txtanio_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   txtPersCod.SetFocus
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLogProSelReqConsolida = Nothing
End Sub

Private Sub mnuQuitarReq_Click()
    QuitaRequerimeintos
End Sub

Private Sub MSFlex_DblClick()
On Error GoTo MSFlexErr
    Dim i As Integer, bTipo As Boolean
    With MSFlex
        If Trim(.TextMatrix(.row, 2)) = "-" Then
           .TextMatrix(.row, 2) = "+"
           i = .row + 1
           bTipo = True
        ElseIf Trim(.TextMatrix(.row, 2)) = "+" Then
           .TextMatrix(.row, 2) = "-"
           i = .row + 1
           bTipo = False
        End If
        Do While i < .Rows
            If Trim(.TextMatrix(i, 2)) = "+" Or Trim(.TextMatrix(i, 2)) = "-" Then
                Exit Sub
            End If
            
            If bTipo Then
                .RowHeight(i) = 0
            Else
                .RowHeight(i) = 300
            End If
            i = i + 1
        Loop
    End With
Exit Sub
MSFlexErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub MSFlex_GotFocus()
    Dim nPrecio As Currency
    If txtEdit.Visible Then
        nPrecio = MSFlex.TextMatrix(MSFlex.row, 6) / MSFlex.TextMatrix(MSFlex.row, 5)
        MSFlex = txtEdit
        MSFlex.TextMatrix(MSFlex.row, 6) = Format(MSFlex.TextMatrix(MSFlex.row, 5) * nPrecio, "#,##0.00")
        'MSFlex.TextMatrix(MSItem.row, 12) = FNumero(CDbl(MSItem.TextMatrix(MSItem.row, 4)) * CDbl(MSItem.TextMatrix(MSItem.row, 11)))
        txtEdit.Visible = False
        CalcularCantidad MSFlex.TextMatrix(MSFlex.row, 1)
        MSFlex.TextMatrix(MSFlex.row, 0) = -1
    End If
End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
    Select Case MSFlex.Col
        Case 5
            If IsNumeric(Chr(KeyAscii)) Then _
                EditaFlex MSFlex, txtEdit, KeyAscii
    End Select
    If KeyAscii = 13 Then
        MSFlex_DblClick
    End If
End Sub

Private Sub MSFlex_LeaveCell()
    Dim nPrecio As Currency
    If txtEdit.Visible Then
        nPrecio = MSFlex.TextMatrix(MSFlex.row, 6) / MSFlex.TextMatrix(MSFlex.row, 5)
        MSFlex = txtEdit
        MSFlex.TextMatrix(MSFlex.row, 6) = Format(MSFlex.TextMatrix(MSFlex.row, 5) * nPrecio, "#,##0.00")
        'MSFlex.TextMatrix(MSFlex.row, 6) = FNumero(CDbl(MSFlex.TextMatrix(MSFlex.row, 5)) * CDbl(MSFlex.TextMatrix(MSFlex.row, 6)))
        txtEdit.Visible = False
        CalcularCantidad MSFlex.TextMatrix(MSFlex.row, 1)
        MSFlex.TextMatrix(MSFlex.row, 0) = -1
    End If
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode MSFlex, txtEdit, KeyCode, Shift
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then
       KeyAscii = 0
    End If
    Select Case MSFlex.Col
        Case 5
            KeyAscii = DigNumDec(txtEdit, KeyAscii)
    End Select
End Sub

Private Sub txtPersCod_Change()
    Dim sCargos As String
'If Len(txtPersCod) > 0 Then
'   DatosPersonal txtPersCod
If Len(txtPersCod) > 0 Then
    sCargos = VerificaCargoEncargado(IIf(Len(txtPersCod.Text) = 13, txtPersCod.Text, gsCodPersUser))
    If sCargos <> "" Then
        If MsgBox(sCargos & vbCrLf & "Desea Usar el Cargo Encargado", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            DatosPersonal IIf(Len(txtPersCod.Text) = 13, txtPersCod.Text, gsCodPersUser), True
        Else
            DatosPersonal IIf(Len(txtPersCod.Text) = 13, txtPersCod.Text, gsCodPersUser), False
        End If
    Else
        DatosPersonal IIf(Len(txtPersCod.Text) = 13, txtPersCod.Text, gsCodPersUser), False
    End If
End If
End Sub

Sub DatosPersonal(vPersCod As String, ByVal pbEncargado As Boolean)
Dim Rs As New ADODB.Recordset
Dim rn As New ADODB.Recordset
Dim oConn As DConecta
Dim i As Integer, sEncargado As String

FormaFlex
'FormaFlexReq
cRHAgeCod = ""
cRHAreaCod = ""
cRHCargoCod = ""
txtAgencia.Text = ""
txtCargo.Text = ""
txtArea.Text = ""

sEncargado = ""

If pbEncargado Then
    sEncargado = "select top 1 cPersCod, cRHCargoCodOficial as cRHCargoCod, cRHAreaCodOficial as cAreaCod, cRHAgenciaCodOficial as cAgeCod from RHCargos where cPersCod='" & vPersCod & "' order by dRHCargoFecha desc"
Else
    sEncargado = "select top 1 cPersCod, cRHCargoCodOficial as cRHCargoCod, cRHAreaCodOficial as cAreaCod, cRHAgenciaCodOficial as cAgeCod from RHCargos r inner join RHCargosTabla t on r.cRHCargoCod = t.cRHCargoCod where cPersCod='" & vPersCod & "' and not cRHCargoDescripcion like '%(E)' order by dRHCargoFecha desc"
End If

sSQL = "select x.*,p.cPersNombre,cCargo=coalesce(c.cRHCargoDescripcion,''),cArea=coalesce(a.cAreaDescripcion,''),cAgencia=coalesce(b.cAgeDescripcion,'') " & _
" from (" & sEncargado & ") x " & _
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
      'sstReq.Tab = 0
      cmdConsolida.Visible = True
   End If
End If
End Sub


'-----------------------------------------------------------------------
'Private Sub MSReq_GotFocus()
'If Len(MSReq.TextMatrix(MSReq.Row, 1)) > 0 Then
'   DetalleUsuario MSReq.TextMatrix(MSReq.Row, 1)
'End If
'End Sub

'Private Sub MSReq_RowColChange()
'If Len(MSReq.TextMatrix(MSReq.Row, 1)) > 0 Then
'   DetalleUsuario MSReq.TextMatrix(MSReq.Row, 1)
'End If
'End Sub

'------------------------------------------------------------------------------
' DETALLE POR USUARIO
'------------------------------------------------------------------------------
'Private Sub DetalleUsuario(ByVal pnProSelReqNro As Integer)
'Dim rs As New ADODB.Recordset
'Dim oPlan As New DLogPlanAnual
'Dim oCon As DConecta, sSQL As String
'Dim i As Integer
'
'Set oCon = New DConecta
'
'MSUsu.Clear
'MSUsu.Rows = 2
'MSUsu.RowHeight(0) = 290
'MSUsu.RowHeight(1) = 8
'MSUsu.ColWidth(0) = 0
'MSUsu.ColWidth(1) = 0
'MSUsu.ColWidth(1) = 0:            MSUsu.TextMatrix(0, 1) = ""
'MSUsu.ColWidth(2) = 2500:         MSUsu.TextMatrix(0, 2) = "Descripción"
'MSUsu.ColWidth(3) = 1200:         MSUsu.TextMatrix(0, 3) = "   U. Medida":   MSUsu.ColAlignment(3) = 4
'MSUsu.ColWidth(4) = 1200:         MSUsu.TextMatrix(0, 4) = "Cantidad"
'MSUsu.ColWidth(5) = 0
'
'If oCon.AbreConexion Then
'    sSQL = "select d.cProSelBSCod,cBSDescripcion=substring(b.cBSDescripcion,1,60),cUnidad, nCantidad " & _
'           " from LogProSelReqDetalle d inner join LogProSelBienesServicios b on d.cProSelBSCod = b.cProSelBSCod " & _
'           " inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) t on b.nBSUnidad = t.nBSUnidad " & _
'           " where nProSelReqNro = " & pnProSelReqNro
'    Set rs = oCon.CargaRecordSet(sSQL)
''    Set rs = oPlan.RequerimientoPlanAnual(pnProSelReqNro, True)
'    If Not rs.EOF Then
'       i = 0
'       Do While Not rs.EOF
'          i = i + 1
'          InsRow MSUsu, i
'          MSUsu.TextMatrix(i, 1) = rs!cProSelBSCod
'          MSUsu.TextMatrix(i, 2) = rs!cBSDescripcion
'          MSUsu.TextMatrix(i, 3) = rs!cUnidad
'          MSUsu.TextMatrix(i, 4) = rs!nCantidad
'          rs.MoveNext
'       Loop
'    End If
'    oCon.CierraConexion
'End If
'End Sub

'------------------------------------------------------------------------------
' DETALLE POR BIEN/SERVICIO
'------------------------------------------------------------------------------
'Private Sub DetalleBienServicio(psBSCod As String)
'Dim rs As New ADODB.Recordset
'Dim oConn As New DConecta, i As Integer
'
'MSDet.Clear
'MSDet.Rows = 2
'MSDet.RowHeight(-1) = 280
'MSDet.RowHeight(0) = 300
'MSDet.RowHeight(1) = 8
'MSDet.ColWidth(0) = 0
'MSDet.ColWidth(1) = 0
'MSDet.ColWidth(2) = 0
'MSDet.ColWidth(3) = 7000:    MSDet.TextMatrix(0, 2) = "Descripción"
'MSDet.ColWidth(4) = 0
'MSDet.ColWidth(5) = 0
''--------------------------------------------------------------------------
'
'If Len(Trim(cRangosIN)) = 0 Then Exit Sub
'
'sSQL = "select r.nProSelReqNro, r.cPersCod, cPersona=replace(p.cPersNombre,'/',' '), d.* " & _
'"  from LogProSelReqDetalle d inner join LogProSelReq r on d.nProSelReqNro = r.nProSelReqNro " & _
'"       inner join Persona p on r.cPersCod = p.cPersCod  " & _
'" where r.nProSelReqNro in (" & cRangosIN & ") and  d.cProSelBSCod = '" & psBSCod & "' and d.nEstado=1 "
'
'If oConn.AbreConexion Then
'   Set rs = oConn.CargaRecordSet(sSQL)
'   If Not rs.EOF Then
'      i = 0
'      Do While Not rs.EOF
'         i = i + 1
'         InsRow MSDet, i
'         MSDet.TextMatrix(i, 1) = rs!nProSelReqNro
'         MSDet.TextMatrix(i, 2) = rs!cPersCod
'         MSDet.TextMatrix(i, 3) = rs!cPersona
'         MSDet.TextMatrix(i, 4) = rs!nCantidad
'         rs.MoveNext
'      Loop
'   End If
'End If
'End Sub

Private Sub mnuConsolidado_Click()
cmdConsolida_Click
End Sub


'*******************************************************************************
Private Sub MSFlex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu mnuConsol
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
'Edt.Text = Chr(KeyAscii) ' & MSFlex
Edt.Visible = True
Edt.SetFocus
End Sub

Private Sub CalcularCantidad(ByVal pnPapa As Integer)
    Dim nCant As Integer, i As Integer, nVal As Currency
    With MSFlex
        Do While i < .Rows
            If Val(.TextMatrix(i, 1)) = pnPapa Then
                nCant = nCant + CDbl(.TextMatrix(i, 5))
                nVal = nVal + CDbl(.TextMatrix(i, 6))
            End If
            i = i + 1
        Loop
        .TextMatrix(pnPapa, 5) = nCant
        .TextMatrix(pnPapa, 6) = FNumero(nVal)  ' * nCant, "#,##0.00")
    End With
End Sub

Private Function VerificaCargoEncargado(ByVal pcPersCod As String) As String
On Error GoTo VerificaCargoEncargadoErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, sSQLX As String
    Set oCon = New DConecta
    sSQL = " select top 1 cPersCod, cRHCargoDescripcion, dRHCargoFecha, " & _
           " cRHCargoCodOficial as cRHCargoCod, " & _
           " cRHAreaCodOficial as cAreaCod, " & _
           " cRHAgenciaCodOficial As cAgeCod " & _
           " from RHCargos r " & _
           " inner join RHCargosTabla t on r.cRHCargoCod = t.cRHCargoCod " & _
           " where cPersCod='" & pcPersCod & "' order by dRHCargoFecha desc "
           
    sSQLX = " select top 2 cPersCod, cRHCargoDescripcion, dRHCargoFecha, " & _
           " cRHCargoCodOficial as cRHCargoCod, " & _
           " cRHAreaCodOficial as cAreaCod, " & _
           " cRHAgenciaCodOficial As cAgeCod " & _
           " from RHCargos r " & _
           " inner join RHCargosTabla t on r.cRHCargoCod = t.cRHCargoCod " & _
           " where cPersCod='" & pcPersCod & "' order by dRHCargoFecha desc "
    
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            If InStr(1, Rs!cRHCargoDescripcion, "(E)") > 0 Then
                Rs.Close
                Set Rs = oCon.CargaRecordSet(sSQLX)
                Do While Not Rs.EOF
                    If VerificaCargoEncargado = "" Then
                        VerificaCargoEncargado = "Sus cargos actualez son: " & vbCrLf & Rs!cRHCargoDescripcion
                    Else
                        VerificaCargoEncargado = VerificaCargoEncargado & vbCrLf & Rs!cRHCargoDescripcion
                    End If
                    Rs.MoveNext
                Loop
            End If
        End If
        oCon.CierraConexion
    End If
    Exit Function
VerificaCargoEncargadoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

