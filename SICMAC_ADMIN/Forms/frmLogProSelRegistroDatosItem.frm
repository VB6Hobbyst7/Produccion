VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogProSelRegistroDatosItem 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3180
   ClientLeft      =   1275
   ClientTop       =   3270
   ClientWidth     =   8130
   Icon            =   "frmLogProSelRegistroDatosItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   8130
   Begin VB.Frame fraEta 
      Height          =   3135
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   7995
      Begin MSMask.MaskEdBox txtFechaIni 
         Height          =   315
         Left            =   1620
         TabIndex        =   18
         Top             =   960
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtEtapa 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   540
         Width           =   6135
      End
      Begin VB.CommandButton cmdGrabarEta 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   5220
         TabIndex        =   8
         Top             =   2520
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelaEta 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6600
         TabIndex        =   7
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtObs 
         Height          =   615
         Left            =   1620
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1800
         Width           =   6195
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   1620
         TabIndex        =   19
         Top             =   1380
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Etapa"
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
         Left            =   420
         TabIndex        =   13
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Observación"
         Height          =   195
         Left            =   420
         TabIndex        =   11
         Top             =   1860
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         Height          =   195
         Left            =   420
         TabIndex        =   10
         Top             =   1020
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Término"
         Height          =   195
         Left            =   420
         TabIndex        =   9
         Top             =   1440
         Width           =   1065
      End
   End
   Begin VB.Frame fraVis 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   7995
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   2580
         Width           =   1155
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   2580
         Width           =   1155
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   6660
         TabIndex        =   1
         Top             =   2580
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   2535
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   4471
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         Cols            =   6
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   -2147483647
         ForeColorSel    =   -2147483624
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         WordWrap        =   -1  'True
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
   End
   Begin VB.Frame fraReg 
      Caption         =   "Descripcion"
      Height          =   2955
      Left            =   60
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   7995
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   5040
         TabIndex        =   17
         Top             =   2520
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6420
         TabIndex        =   16
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtConcepto 
         Height          =   2175
         Left            =   120
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   300
         Width           =   7755
      End
   End
End
Attribute VB_Name = "frmLogProSelRegistroDatosItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vpGrabado As Boolean

Dim nProselNro As Integer, nProSelItem As Integer, dFechaI As Date, nTipo As Integer, _
    cDescripcion As String, dFEchaF As Date, cProSelBSCod As String
Public gdFechaI As String, gdFechaF As String

Public Sub Inicio(ByVal vProSelNro As Integer, ByVal vProSelItem As Integer, Optional ByVal vFechaI As Date = "01/01/1900", Optional ByVal vFechaF As Date = "01/01/1900", Optional ByVal vTipo As Integer = 1, Optional ByVal vDescripcion As String = "", Optional ByVal pcBSCod As String = "")
nProselNro = vProSelNro
nProSelItem = vProSelItem
nTipo = vTipo
cDescripcion = vDescripcion
dFechaI = vFechaI
dFEchaF = vFechaF
cProSelBSCod = pcBSCod
Me.Show 1
End Sub

Private Sub Form_Load()
CentraForm Me
Me.vpGrabado = False
Select Case nTipo
    Case 1
         Caption = "Registro de Caracteristicas de Items"
         CargaCaracteristicas
    Case 2
         Caption = "Registro de Propuestas de Postores"
         CargaPropuestas
    Case 4
         Caption = "Registro de Fechas en Etapas - Proceso de Selección"
         fraVis.Visible = False
         fraEta.Visible = True
         txtFechaIni = dFechaI
         txtFechaFin = dFEchaF
         txtEtapa.Text = cDescripcion
         txtFechaIni.TabIndex = 0
         txtFechaFin.TabIndex = 1
         txtObs.TabIndex = 2
End Select
End Sub


Private Sub cmdCancelaEta_Click()
Me.vpGrabado = False
Unload Me
End Sub

Private Sub cmdGrabarEta_Click()
Dim oConn As New DConecta
Dim sSQL As String

If Not oConn.AbreConexion Then
   MsgBox "No se puede abrir la conexión...." + Space(10), vbInformation
   Exit Sub
End If

If txtFechaIni.Text = "__/__/____" Or Not IsDate(txtFechaIni.Text) Then
    MsgBox "Debe Ingrear una Fecha Correcta", vbInformation, "Aviso"
    txtFechaIni.SetFocus
    Exit Sub
End If

If txtFechaFin.Text = "__/__/____" Or Not IsDate(txtFechaFin.Text) Then
    MsgBox "Debe Ingrear una Fecha Correcta", vbInformation, "Aviso"
    txtFechaFin.SetFocus
    Exit Sub
End If

If CDate(txtFechaIni.Text) - CDate(txtFechaFin.Text) > 0 Then
    MsgBox "Rango de Fecha Incorrecto", vbInformation, "Aviso"
    txtFechaFin.SetFocus
    Exit Sub
End If

If CDate(txtFechaIni.Text) < gdFecSis Then
    MsgBox "Fecha de Inicio Incorrecta", vbInformation, "Aviso"
    Exit Sub
End If

If MsgBox("¿ Está seguro de agregar las Fechas a la etapa indicada ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then

   gdFechaI = IIf(txtFechaIni.Text = "__/__/____", "", txtFechaIni.Text)
   gdFechaF = IIf(txtFechaFin.Text = "__/__/____", "", txtFechaFin.Text)
   
   sSQL = "UPDATE LogProSelEtapa SET dFechaInicio = '" & Format(gdFechaI, "YYYYMMDD") & "', dFechaTermino = '" & Format(gdFechaF, "YYYYMMDD") & "', cObservacion = '" & txtObs.Text & "' " & _
          " WHERE nProSelNro = " & nProselNro & " and nEtapaCod = " & nProSelItem & " "
   oConn.Ejecutar sSQL
   Me.vpGrabado = True
   Unload Me
End If
End Sub

Private Sub cmdQuitar_Click()
On Error GoTo cmdQuitarErr
    Dim oCon As DConecta, sSQL As String
    
    If Len(MSFlex.TextMatrix(MSFlex.row, 1)) = 0 Then Exit Sub
    
    If MsgBox("Seguro que Desea Eliminar...?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "delete LogProSelItemCaracteristicas where nProSelNro=" & nProselNro & " and nProSelItem=" & nProSelItem & " and nItem=" & MSFlex.TextMatrix(MSFlex.row, 1)
        oCon.Ejecutar sSQL
        oCon.CierraConexion
        MsgBox "Caracteristica Eliminada...", vbInformation
        CargaCaracteristicas
    End If
    Exit Sub
cmdQuitarErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

'Private Sub cmdQuitarEsp_Click()
'Dim i As Integer
'Dim k As Integer
'
'
'i = MSFlexEsp.Row
'If Len(Trim(MSFlexEsp.TextMatrix(i, 0))) = 0 Then
'   Exit Sub
'End If
'
'If MsgBox("¿ está seguro de quitar el elemento ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
'
'   If MSFlexEsp.Rows - 1 > 1 Then
'      MSFlexEsp.RemoveItem i
'   Else
'      'MSFlex.Clear          Quita las cabeceras
'      For k = 0 To MSFlexEsp.Cols - 1
'          MSFlexEsp.TextMatrix(i, k) = ""
'      Next
'      MSFlexEsp.RowHeight(i) = 8
'   End If
'End If
'End Sub

Private Sub cmdAgregar_Click()
fraVis.Visible = False
Select Case nTipo
    Case 1
         fraReg.Visible = True
         txtConcepto.Text = ""
    Case 2
         'fraPos.Visible = True
         'txtPersCod.Text = ""
         'txtPersona.Text = ""
         'txtPropEcon.Text = ""
         'txtPuntaje.Text = ""
         'chkGanador.Value = 0
    Case 3
         'fraObs.Visible = True
         'txtPersCodObs.Text = ""
         'txtPersObs.Text = ""
         'txtConsulta.Text = ""
         'txtrespuesta.Text = ""
End Select
End Sub



'Private Sub cmdCancelaObs_Click()
'fraObs.Visible = False
'fraVis.Visible = True
'End Sub

'Private Sub cmdCancelaPos_Click()
'fraPos.Visible = False
'fraVis.Visible = True
'End Sub

Private Sub cmdCancelar_Click()
fraReg.Visible = False
fraVis.Visible = True
End Sub

'Private Sub cmdPersona_Click()
'Dim X As UPersona
'Set X = frmBuscaPersona.Inicio
'
'If X Is Nothing Then
'    Exit Sub
'End If
'
'If Len(Trim(X.sPersNombre)) > 0 Then
'   txtPersona.Text = X.sPersNombre
'   txtPersCod = X.sPersCod
'End If

'frmBuscaPersona.Show 1
'If frmBuscaPersona.vpOK Then
'   txtPersona.Text = frmBuscaPersona.vpPersNom
'   txtPersCod = frmBuscaPersona.vpPersCod
'   'valida si el postor ya esta para el item
'
'End If
'End Sub

'Private Sub cmdPersObs_Click()
'Dim X As UPersona
'Set X = frmBuscaPersona.Inicio
'
'If X Is Nothing Then
'    Exit Sub
'End If
'
'If Len(Trim(X.sPersNombre)) > 0 Then
'   txtPersona.Text = X.sPersNombre
'   txtPersCod = X.sPersCod
'End If

'frmBuscaPersona.Show 1
'If frmBuscaPersona.vpOK Then
'   txtPersObs.Text = frmBuscaPersona.vpPersNom
'   txtPersCodObs = frmBuscaPersona.vpPersCod
'End If
'End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'Sub CargaObservaciones()
'Dim oConn As New DConecta, sSQL As String, rs As ADODB.Recordset, i As Integer
'On Error GoTo CargaProp
'
'FormaFlex
'If oConn.AbreConexion Then
'
'   sSQL = "select x.nProSelNro, x.nTipo, x.cPersCod, p.cPersNombre,x.cConsulta, x.cRespuesta  " & _
'          " from LogProSelConsultas x inner join Persona p on x.cPersCod = p.cPersCod " & _
'          " Where x.nProSelNro = " & nProselNro & ""
'
'   Set rs = oConn.CargaRecordSet(sSQL)
'   i = 0
'   Do While Not rs.EOF
'        If bConObs And rs!nTipo = 1 Then
'            i = i + 1
'            InsRow MSFlex, i
'            MSFlex.RowHeight(i) = 280
'            MSFlex.TextMatrix(i, 2) = "Consulta"
'            MSFlex.TextMatrix(i, 3) = rs!cPersNombre
'            MSFlex.TextMatrix(i, 4) = rs!cConsulta
'            'MSFlex.TextMatrix(i, 5) = rs!cRespuesta
'            MSFlex.ScrollBars = flexScrollBarBoth
'        ElseIf bConObs = False And rs!nTipo = 2 Then
'            i = i + 1
'            InsRow MSFlex, i
'            MSFlex.RowHeight(i) = 280
'            MSFlex.TextMatrix(i, 2) = "Observacion"
'            MSFlex.TextMatrix(i, 3) = rs!cPersNombre
'            MSFlex.TextMatrix(i, 4) = rs!cConsulta
'            'MSFlex.TextMatrix(i, 5) = rs!cRespuesta
'            MSFlex.ScrollBars = flexScrollBarBoth
'        End If
'        rs.MoveNext
'    Loop
'End If
'Exit Sub
'CargaProp:
'    MsgBox Err.Number & vbCrLf & Err.Description
'End Sub

Sub CargaPropuestas()
Dim oConn As New DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
On Error GoTo CargaProp
   
FormaFlex
If oConn.AbreConexion Then
          
   sSQL = "select pp.*, pe.cPersNombre " & _
          " from LogProSelPostorPropuesta pp inner join Persona pe on pp.cPersCod = pe.cPersCod " & _
          " Where pp.nProSelNro = " & nProselNro & " And pp.nProSelItem = " & nProSelItem & " "
          
   Set Rs = oConn.CargaRecordSet(sSQL)
   i = 0
   Do While Not Rs.EOF
      i = i + 1
        InsRow MSFlex, i
        MSFlex.RowHeight(i) = 280
        MSFlex.TextMatrix(i, 2) = Rs!cPersNombre
        MSFlex.TextMatrix(i, 3) = Rs!nPropEconomica
        MSFlex.TextMatrix(i, 4) = IIf(IsNull(Rs!npuntaje), 0, Rs!npuntaje)
        Rs.MoveNext
    Loop
End If
Exit Sub

CargaProp:
    MsgBox Err.Number & vbCrLf & Err.Description

End Sub

Private Sub CargaCaracteristicas()
Dim oConn As New DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
On Error GoTo CargarCaracteristicasErr
   
FormaFlex
If oConn.AbreConexion Then
   sSQL = "select * from LogProSelItemCaracteristicas where nProSelNro=" & nProselNro & " and nProSelItem=" & nProSelItem & " and cBSCod='" & cProSelBSCod & "'"
   Set Rs = oConn.CargaRecordSet(sSQL)
   i = 0
   Do While Not Rs.EOF
      i = i + 1
        InsRow MSFlex, i
        MSFlex.RowHeight(i) = 1000
        MSFlex.TextMatrix(i, 1) = Rs!nItem
        MSFlex.TextMatrix(i, 2) = Rs!cDescripcion
'        MSFlex.TextMatrix(i, 3) = Rs!nValor
'        MSFlex.TextMatrix(i, 4) = Rs!nPonderacion
        Rs.MoveNext
    Loop
End If
Exit Sub

CargarCaracteristicasErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Sub FormaFlex()
Dim i As Integer
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.Cols = 6
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 10
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 0
Select Case nTipo
    Case 1
         MSFlex.ColWidth(2) = 7500:    MSFlex.TextMatrix(0, 2) = Space(15) & "Descripcion"
         MSFlex.ColWidth(3) = 0:    MSFlex.TextMatrix(0, 3) = "Valor"
         MSFlex.ColWidth(4) = 0:    MSFlex.TextMatrix(0, 4) = "Ponderacion"
         MSFlex.ColWidth(5) = 0
    Case 2
         MSFlex.ColWidth(2) = 5100:    MSFlex.TextMatrix(0, 2) = "Postor"
         MSFlex.ColWidth(3) = 1100:    MSFlex.TextMatrix(0, 3) = "Prop.Econ."
         MSFlex.ColWidth(4) = 800:     MSFlex.TextMatrix(0, 4) = "Puntaje"
         MSFlex.ColWidth(5) = 0
'         MSFlexEsp.RowHeight(0) = 320
'         MSFlexEsp.ColWidth(0) = 500:    MSFlexEsp.TextMatrix(0, 0) = "Item":  MSFlexEsp.ColAlignment(0) = 4
'         MSFlexEsp.ColWidth(1) = 5000:   MSFlexEsp.TextMatrix(0, 1) = "Concepto"
'         MSFlexEsp.ColWidth(2) = 1200:   MSFlexEsp.TextMatrix(0, 2) = "Valor"

    Case 3
         MSFlex.ColWidth(2) = 1000:    MSFlex.TextMatrix(0, 2) = "Tipo"
         MSFlex.ColWidth(3) = 3000:    MSFlex.TextMatrix(0, 3) = "Persona"
         MSFlex.ColWidth(4) = 3000:    MSFlex.TextMatrix(0, 4) = "Descripción"
         MSFlex.ColWidth(5) = 5000:    MSFlex.TextMatrix(0, 5) = "Respuesta"
End Select
End Sub

Private Sub cmdGrabar_Click()
Dim oConn As New DConecta, sSQL As String
On Error GoTo GrabarCaracteristicasErr

sSQL = " insert into LogProSelItemCaracteristicas (nProSelNro,nProSelItem,cBSCod,cDescripcion) " & _
       " values (" & nProselNro & "," & nProSelItem & ",'" & cProSelBSCod & "','" & txtConcepto & "')" ' & VNumero(TxtValor) & "," & VNumero(txtPonde) & ")"
           
If oConn.AbreConexion Then
   If MsgBox("¿ Está seguro de agregar la caracteríticas ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
      oConn.Ejecutar sSQL
      oConn.CierraConexion
      fraReg.Visible = False
      fraVis.Visible = True
      CargaCaracteristicas
   End If
End If
Exit Sub
GrabarCaracteristicasErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'Private Sub cmdGrabarPos_Click()
'Dim oConn As New DConecta, sSQL As String, i As Integer
'On Error GoTo GrabarPos
'
''sSQL = " insert into LogProSelPostorPropuesta (nProSelNro,nProSelItem,cPersCod,nPropEconomica,nPuntaje,bGanador) " & _
''       " values (" & nProSelNro & "," & nProSelItem & ",'" & txtPersCod & "'," & VNumero(txtPropEcon) & "," & VNumero(txtPuntaje) & "," & chkGanador.Value & ")"
'
'If Len(Trim(txtPersCod.Text)) > 0 Then
'   MsgBox "Debe ingresar un postor válido..." + Space(10), vbInformation
'   Exit Sub
'End If
'
'If Len(Trim(txtPropEcon.Text)) > 0 And VNumero(txtPropEcon.Text) > 0 Then
'   MsgBox "Debe ingresar una propuesta económica válida..." + Space(10), vbInformation
'   Exit Sub
'End If
'
'
'If Len(Trim(MSFlexEsp.TextMatrix(1, 0))) > 0 Then
'
'End If
'
'
'   sSQL = " insert into LogProSelPostorPropuesta (nProSelNro,nProSelItem,cPersCod,nPropEconomica) " & _
'          " values (" & nProSelNro & "," & nProSelItem & ",'" & txtPersCod & "'," & VNumero(txtPropEcon) & " )"
'
'   If oConn.AbreConexion Then
'      If MsgBox("¿ Está seguro de agregar la propuesta del Postor ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
'         oConn.Ejecutar sSQL
'
'
'         i = 1
'         Do While i < MSFlexEsp.Rows
'
'            'GrabarEspecificaciones nProSelNro, nProSelItem, txtPersCod, MSFlexEsp.TextMatrix(i, 0), MSFlexEsp.TextMatrix(i, 1), MSFlexEsp.TextMatrix(i, 2)
'
'            i = i + 1
'         Loop
'
'
'         oConn.CierraConexion
'         fraPos.Visible = False
'         fraVis.Visible = True
'         CargaPropuestas
'      End If
'   End If
'
'Exit Sub
'
'GrabarPos:
'    MsgBox Err.Number & vbCrLf & Err.Description
'End Sub

'Private Sub cmdGrabarPos_Click()
'Dim oConn As New DConecta, sSQL As String, i As Integer
'On Error GoTo GrabarPos
'
''sSQL = " insert into LogProSelPostorPropuesta (nProSelNro,nProSelItem,cPersCod,nPropEconomica,nPuntaje,bGanador) " & _
''       " values (" & nProSelNro & "," & nProSelItem & ",'" & txtPersCod & "'," & VNumero(txtPropEcon) & "," & VNumero(txtPuntaje) & "," & chkGanador.Value & ")"
'
'If Len(Trim(txtPersCod.Text)) = 0 Then
'   MsgBox "Debe ingresar un postor válido..." + Space(10), vbInformation
'   txtPersCod.SetFocus
'   Exit Sub
'End If

'If Len(Trim(txtPropEcon.Text)) = 0 Or VNumero(txtPropEcon.Text) < 0 Then
'   MsgBox "Debe ingresar una propuesta económica válida..." + Space(10), vbInformation
'   txtPropEcon.SetFocus
'   Exit Sub
'End If
'
'i = 1
'Do While i < MSFlexEsp.Rows
'    If Len(Trim(MSFlexEsp.TextMatrix(i, 0))) = 0 Or Len(Trim(MSFlexEsp.TextMatrix(i, 1))) = 0 Or _
'        Len(Trim(MSFlexEsp.TextMatrix(i, 2))) = 0 _
'        Or VNumero(MSFlexEsp.TextMatrix(i, 2)) < 0 Then
'        MsgBox "Debe ingresar Caracteristicas Validas para la Propuesta..." + Space(10), vbInformation
'        MSFlexEsp.SetFocus
'        Exit Sub
'    End If
'    i = i + 1
'Loop
'
'   sSQL = " insert into LogProSelPostorPropuesta (nProSelNro,nProSelItem,cPersCod,nPropEconomica) " & _
'          " values (" & nProselNro & "," & nProSelItem & ",'" & txtPersCod & "'," & VNumero(txtPropEcon) & " )"
'
'   If oConn.AbreConexion Then
'      If MsgBox("¿ Está seguro de agregar la propuesta del Postor ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
'         oConn.Ejecutar sSQL
'
'         i = 1
'         Do While i < MSFlexEsp.Rows
'            sSQL = "insert into LogProSelPostorPropEspecificacion(nProSelNro,nProSelItem,cPersCod,nItemEspecificaciones,cConcepto,nValor) " & _
'                   " values(" & nProselNro & "," & nProSelItem & "," & txtPersCod & "," & MSFlexEsp.TextMatrix(i, 0) & ",'" & MSFlexEsp.TextMatrix(i, 1) & "'," & MSFlexEsp.TextMatrix(i, 2) & ") "
'            oConn.Ejecutar sSQL
'            i = i + 1
'         Loop
'
'         oConn.CierraConexion
'         fraPos.Visible = False
'         fraVis.Visible = True
'         CargaPropuestas
'      End If
'   End If
'
'Exit Sub
'
'GrabarPos:
'    MsgBox Err.Number & vbCrLf & Err.Description
'End Sub

'Private Sub cmdGrabarObs_Click()
'Dim oConn As New DConecta, sSQL As String, xConsulta As String
'Dim nTipo As Integer
'
'If bConObs Then
'    nTipo = 1
'Else
'    nTipo = 2
'End If
'
'sSQL = " insert into LogProSelConsultas (nProSelNro,nTipo,cPersCod,cConsulta,cRespuesta) " & _
'       " values (" & nProselNro & "," & nTipo & ",'" & txtPersCodObs & "','" & txtConsulta & "','" & txtrespuesta & "' )"
'
'If oConn.AbreConexion Then
'   If MsgBox("¿ Está seguro de agregar la " & lblDesc.Caption & " del Postor ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
'      oConn.Ejecutar sSQL
'      oConn.CierraConexion
'      fraPos.Visible = False
'      fraVis.Visible = True
'      CargaObservaciones
'      cmdCancelaObs_Click
'   End If
'End If
'Exit Sub
'
'End Sub


'Private Sub opCon01_Click()
'If opCon01.Value Then
'   lblDesc.Caption = "Consulta"
'Else
'   lblDesc.Caption = "Observación"
'End If
'End Sub

'Private Sub opCon02_Click()
'If opCon02.Value Then
'   lblDesc.Caption = "Observación"
'Else
'   lblDesc.Caption = "Consulta"
'End If
'End Sub

'Private Sub cmdAgregarEsp_Click()
'On Error GoTo cmdGrabarEspErr
'
'Dim sSQL As String, oConn As New DConecta, rs As ADODB.Recordset, nItem As Integer, i As Integer
'Dim nRow As Integer
'
'nRow = MSFlexEsp.Rows - 1
'If Len(Trim(MSFlexEsp.TextMatrix(nRow, 1))) > 0 Then
'   nRow = nRow + 1
'End If
'
'    nItem = IIf(MSFlexEsp.TextMatrix(MSFlexEsp.Rows - 1, 0) <> "", MSFlexEsp.TextMatrix(MSFlexEsp.Rows - 1, 0), 0) + 1
'    With MSFlexEsp
'
'        InsRow MSFlexEsp, nRow
'        MSFlexEsp.TextMatrix(MSFlexEsp.Rows - 1, 0) = nItem
'        MSFlexEsp.Row = nRow
'        MSFlexEsp.Col = 1
'        MSFlexEsp.SetFocus
'    End With
'    Exit Sub
'cmdGrabarEspErr:
'    MsgBox Err.Number & vbCrLf & Err.Description
'End Sub

'*********************************************************************
'VALIDACION DE CAMPOS
'*********************************************************************
Private Sub txtPonde_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtPropEcon_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TxtValor_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub



'*********************************************************************
'CAMBIOS DE FOCOS
'*********************************************************************





'*********************************************************************
'PROCEDIMIENTOS DEL FLEX
'*********************************************************************
'Private Sub MSFlexEsp_GotFocus()
'If txtEdit.Visible = False Then Exit Sub
'MSFlexEsp = txtEdit
'txtEdit.Visible = False
'End Sub
'
'Private Sub MSFlexEsp_LeaveCell()
'If txtEdit.Visible = False Then Exit Sub
'MSFlexEsp = txtEdit
'txtEdit.Visible = False
'End Sub

'Private Sub MSFlexEsp_KeyPress(KeyAscii As Integer)
''If MSFlexEsp.Col >= 1 And MSFlexEsp.Col < 3 Then
''   EditaFlex MSFlex, txtEdit, KeyAscii
''End If
'Select Case MSFlexEsp.Col
'    Case 1
'        If Not IsNumeric(Chr(KeyAscii)) Then _
'            EditaFlex MSFlex, txtEdit, KeyAscii
'    Case 2
'        If IsNumeric(Chr(KeyAscii)) Then _
'            EditaFlex MSFlex, txtEdit, KeyAscii
'End Select
'End Sub

'Sub EditaFlex(MSFlex As Control, Edt As Control, KeyAscii As Integer)
'Select Case KeyAscii
'    Case 0 To 32
'         Edt = MSFlex
'         Edt.SelStart = 1000
'    Case Else
'         Edt = Chr(KeyAscii)
'         Edt.SelStart = 1
'End Select
'Edt.Move MSFlexEsp.Left + MSFlexEsp.CellLeft - 15, MSFlexEsp.Top + MSFlexEsp.CellTop - 15, _
'         MSFlexEsp.CellWidth, MSFlexEsp.CellHeight
'Edt.Visible = True
'Edt.SetFocus
'End Sub

'Private Sub txtEdit_KeyPress(KeyAscii As Integer)
'If KeyAscii = Asc(vbCr) Then
'   KeyAscii = 0
'End If
'End Sub

'Private Sub txtEdit_KeyPress(KeyAscii As Integer)
'    If KeyAscii = Asc(vbCr) Then
'       KeyAscii = 0
'    End If
'Select Case MSFlexEsp.Col
'    Case 1
'        If IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
'    Case 2
'        If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
'End Select
'End Sub

'Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
'EditKeyCode MSFlexEsp, txtEdit, KeyCode, Shift
'End Sub
'
'Sub EditKeyCode(MSFlex As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
'Select Case KeyCode
'    Case 27
'         Edt.Visible = False
'         MSFlexEsp.SetFocus
'    Case 13
'         MSFlexEsp.SetFocus
'    Case 37                     'Izquierda
'         MSFlexEsp.SetFocus
'         DoEvents
'         If MSFlexEsp.Col > 1 Then
'            MSFlexEsp.Col = MSFlexEsp.Col - 1
'         End If
'    Case 39                     'Derecha
'         MSFlex.SetFocus
'         DoEvents
'         If MSFlexEsp.Col < MSFlexEsp.Cols - 1 Then
'            MSFlexEsp.Col = MSFlexEsp.Col + 1
'         End If
'    Case 38
'         MSFlex.SetFocus
'         DoEvents
'         If MSFlexEsp.Row > MSFlexEsp.FixedRows + 1 Then
'            MSFlexEsp.Row = MSFlexEsp.Row - 1
'         End If
'    Case 40
'         MSFlexEsp.SetFocus
'         DoEvents
'         'If MSFlex.Row < MSFlex.FixedRows - 1 Then
'         If MSFlexEsp.Row < MSFlexEsp.Rows - 1 Then
'            MSFlexEsp.Row = MSFlexEsp.Row + 1
'         End If
'End Select
'End Sub

Private Sub txtFechaIni_GotFocus()
SelTexto txtFechaIni
End Sub

Private Sub txtFechaFin_GotFocus()
SelTexto txtFechaFin
End Sub

Private Sub txtFechaIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFechaFin.SetFocus
End If
End Sub

Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtObs.SetFocus
End If
End Sub

