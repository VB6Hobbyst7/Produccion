VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogPlanAnualReq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan Anual de Adquisiciones  y Contrataciones - Registro de Requerimientos de Usuario"
   ClientHeight    =   5595
   ClientLeft      =   105
   ClientTop       =   2175
   ClientWidth     =   11715
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   11715
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   11595
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
         Left            =   5580
         TabIndex        =   0
         Text            =   "2005"
         Top             =   90
         Width           =   675
      End
      Begin VB.Label lblReqNro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requerimientos de usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7800
         TabIndex        =   23
         Top             =   120
         Width           =   3090
      End
      Begin VB.Label lblPlan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plan Anual de Adquisiciones y Contrataciones"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   11
         Top             =   120
         Width           =   5280
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00EAFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   7095
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00EAFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   495
         Left            =   7140
         Top             =   0
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   60
      TabIndex        =   12
      Top             =   480
      Width           =   11595
      Begin VB.TextBox txtAgencia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8100
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   900
         Width           =   3255
      End
      Begin VB.TextBox txtArea 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   900
         Width           =   6195
      End
      Begin VB.TextBox txtCargo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   10455
      End
      Begin VB.TextBox txtPersona 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   300
         Width           =   8655
      End
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
         TabIndex        =   13
         Top             =   330
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtPersCod 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   960
         Width           =   330
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   660
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7440
         TabIndex        =   19
         Top             =   960
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   8
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   5160
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H00DDFFFE&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9300
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   1800
      Width           =   11595
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad mensual requerida"
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
         Left            =   6120
         TabIndex        =   9
         Top             =   180
         Width           =   2370
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bienes"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   570
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E3FFE7&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H004A6B36&
         Height          =   420
         Left            =   0
         Top             =   60
         Width           =   3735
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00EAFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H006B6BB8&
         Height          =   420
         Left            =   3720
         Top             =   60
         Width           =   7875
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9180
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   2835
      Left            =   60
      TabIndex        =   6
      Top             =   2280
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   5001
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   18
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
   Begin VB.Menu mnuMenu 
      Caption         =   "MenuReq"
      Visible         =   0   'False
      Begin VB.Menu mnuInfo 
         Caption         =   "Info del trámite "
      End
   End
End
Attribute VB_Name = "frmLogPlanAnualReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMes(1 To 12) As String, sSQL As String
Dim cRHAgeCod As String, cRHAreaCod As String, cRHCargoCod As String
Dim nEditable As Boolean, nPlanReqActual As Long


Private Sub Form_Load()
CentraForm Me
cRHAgeCod = ""
cRHAreaCod = ""
cRHCargoCod = ""
txtAgencia.Text = ""
txtCargo.Text = ""
txtArea.Text = ""
txtAnio.Text = Year(gdFecSis) + 1
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
If gsCodArea = "036" Or gsCodUser = "SIST" Then
   cmdPersona.Visible = True
End If
If Len(Trim(gsCodPersUser)) = 0 And gsCodUser <> "SIST" Then
   MsgBox "El Usuario actual no se puede identificar como personal de la CMAC-T" + Space(10), vbInformation, "Verificación de usuario como personal"
   cmdGrabar.Visible = False
   cmdAgregar.Visible = False
   cmdQuitar.Visible = False
Else
   txtPersCod = gsCodPersUser
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub CmdGrabar_Click()
Dim oConn As DConecta, sSQL As String
Dim cBSCod As String, i As Integer, n As Integer
Dim nPlanNro As Long, Rs As New ADODB.Recordset
Dim nItem As Integer
Dim cLogNro As String
Dim nAnio As Integer

nItem = 0
n = MSFlex.Rows - 1
For i = 1 To n
    If Len(MSFlex.TextMatrix(i, 1)) > 0 And Len(MSFlex.TextMatrix(i, 2)) > 0 Then
       nItem = nItem + 1
    End If
Next
If nItem = 0 Then
   MsgBox "Debe seleccionar al menos un Bien / Servicio..." + Space(10), vbInformation
   Exit Sub
End If
nItem = 0
nAnio = CInt(VNumero(txtAnio.Text))
Set oConn = New DConecta
If oConn.AbreConexion Then
   If MsgBox("¿ Está seguro de grabar ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   
      cLogNro = GetLogMovNro
      
      '1º Deshabilitamos los requerimientos anteriores del usuario
      sSQL = "UPDATE LogPlanAnualReq SET nEstado = 0 WHERE cPersCod = '" & txtPersCod.Text & "'"
      oConn.Ejecutar sSQL
      
      sSQL = "UPDATE LogPlanAnualReqDetalle SET nEstado = 0 WHERE cPersCod = '" & txtPersCod.Text & "'"
      oConn.Ejecutar sSQL
      
      '2º Insertamos cabecera del requerimiento actual del usuario
      sSQL = "INSERT INTO LogPlanAnualReq( nAnio, cPersCod, cRHCargoCod, cRHAreaCod, cRHAgeCod, cMovNro) " & _
             "    VALUES (" & nAnio & ",'" & txtPersCod.Text & "','" & cRHCargoCod & "','" & cRHAreaCod & "','" & cRHAgeCod & "','" & cLogNro & "') "
      oConn.Ejecutar sSQL
      
      '3º Hallamos ultima secuencia de los requerimientos
      nPlanNro = UltimaSecuenciaIdentidad("LogPlanAnualReq")
      
      '---------------------------------------------------------------------------------
      nItem = 0
      For i = 1 To n
          If Len(MSFlex.TextMatrix(i, 1)) > 0 Then
             nItem = nItem + 1
             sSQL = "INSERT INTO LogPlanAnualReqDetalle (nPlanReqNro,cPersCod,nAnio, nItem, cBSCod, " & _
                  "            nMes01, nMes02, nMes03, nMes04, nMes05, nMes06, nMes07, nMes08, nMes09, nMes10, nMes11, nMes12) " & _
                  " VALUES (" & nPlanNro & ",'" & txtPersCod.Text & "'," & nAnio & "," & nItem & ",'" & MSFlex.TextMatrix(i, 1) & "'," & _
                  "         " & VNumero(MSFlex.TextMatrix(i, 4)) & "," & VNumero(MSFlex.TextMatrix(i, 5)) & "," & VNumero(MSFlex.TextMatrix(i, 6)) & "," & VNumero(MSFlex.TextMatrix(i, 7)) & "," & VNumero(MSFlex.TextMatrix(i, 8)) & "," & VNumero(MSFlex.TextMatrix(i, 9)) & ", " & _
                  "         " & VNumero(MSFlex.TextMatrix(i, 10)) & "," & VNumero(MSFlex.TextMatrix(i, 11)) & "," & VNumero(MSFlex.TextMatrix(i, 12)) & "," & VNumero(MSFlex.TextMatrix(i, 13)) & "," & VNumero(MSFlex.TextMatrix(i, 14)) & "," & VNumero(MSFlex.TextMatrix(i, 15)) & " )"
             oConn.Ejecutar sSQL
          End If
      Next
      
      sSQL = " insert into LogPlanAnualAprobacion (nPlanReqNro,cRHCargoCodAprobacion,nNivelAprobacion) " & _
             " select " & nPlanNro & ",cRHCargoCodAprobacion,nNivelAprobacion " & _
             " from LogNivelAprobacion where cRHCargoCod = '" & cRHCargoCod & "' order by nNivelAprobacion "
      oConn.Ejecutar sSQL
      
      oConn.CierraConexion
      
      MsgBox "El requerimiento se ha grabado con éxito!" + Space(10), vbInformation
      txtPersCod = ""
      FormaFlex
      Unload Me
   End If
   oConn.CierraConexion
End If
End Sub

Private Sub cmdQuitar_Click()
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

Sub DatosPersonal(vPersCod As String)
Dim oPlan As New DLogPlanAnual
Dim Rs As New ADODB.Recordset
Dim rn As New ADODB.Recordset

Dim i As Integer, K As Integer, nAprobado As Integer
Dim nNivel As Integer, nSector As Integer, nAnio As Integer
Dim cAnioMesActual As String

nAnio = CInt(VNumero(txtAnio.Text))

nPlanReqActual = 0
cRHAgeCod = ""
cRHAreaCod = ""
cRHCargoCod = ""

nEditable = False
cmdAgregar.Visible = False
cmdQuitar.Visible = False
cmdGrabar.Visible = False

FormaFlex

'Set oConn = New DConecta
'sSQL = "select x.*,cCargo=coalesce(c.cRHCargoDescripcion,''),cArea=coalesce(a.cAreaDescripcion,''),cAgencia=coalesce(b.cAgeDescripcion,'') " & _
'       " from (select top 1 cRHCargoCodOficial as cRHCargoCod, cRHAreaCodOficial as cAreaCod, cRHAgenciaCodOficial as cAgeCod " & _
'       "  from RHCargos where cPersCod='" & vPersCod & "' order by dRHCargoFecha desc) x " & _
'       "  left outer join Areas a on x.cAreaCod = a.cAreaCod " & _
'       "  left outer join Agencias b on x.cAgeCod = b.cAgeCod " & _
'       "  left outer join RHCargosTabla c on x.cRHCargoCod = c.cRHCargoCod "

cAnioMesActual = CStr(Year(gdFecSis)) + Format(Month(gdFecSis), "00")
'If oConn.AbreConexion Then
   Set Rs = oPlan.AreaCargoAgencia(vPersCod, cAnioMesActual)
   'oConn.CierraConexion
   If Not Rs.EOF Then
      'cRHAgeCod = rs!cAgeCod
      'cRHAreaCod = rs!cAreaCod
      'cRHCargoCod = rs!cRHCargoCod
      'txtAgencia.Text = rs!cAgencia
      'txtCargo.Text = rs!cCargo
      'txtArea.Text = rs!cArea
      'xRHCargoCod = cRHCargoCod
      'xRHCargoDesc = ""
 
      cRHAgeCod = Rs!cRHAgeCod
      cRHAreaCod = Rs!cRHAreaCod
      cRHCargoCod = Rs!cRHCargoCod
      txtPersona = Rs!cPersona
      txtAgencia.Text = Rs!cRHAgencia
      txtCargo.Text = Rs!cRHCargo
      txtArea.Text = Rs!cRHArea

      'Verifica NIVELES DE APROBACION -----------------------------------------------
      Set rn = GetNivelesAprobacion(cRHCargoCod)
      If rn.State = 0 Then
         MsgBox "No se puede determinar niveles de aprobacion para: " + Space(10) + vbCrLf + txtCargo.Text + Space(10) + vbCrLf + txtArea.Text + Space(10), vbInformation
         cmdQuitar.Visible = False
         cmdAgregar.Visible = False
         cmdGrabar.Visible = False
         Exit Sub
      Else
         If rn.EOF And rn.BOF Then
            MsgBox "No se puede determinar niveles de aprobacion para: " + Space(10) + vbCrLf + txtCargo.Text + Space(10) + vbCrLf + txtArea.Text + Space(10), vbInformation
            cmdQuitar.Visible = False
            cmdAgregar.Visible = False
            cmdGrabar.Visible = False
            Exit Sub
         End If
      End If
      
      '--------------------------------------------------------------
      'sSQL = "select r.nPlanReqNro,a.nEstadoAprobacion " & _
      '     "  from LogPlanAnualReq r inner join LogPlanAnualAprobacion a on r.nPlanReqNro = a.nPlanReqNro " & _
      '     " where r.cPersCod = '" & vPersCod & "' and " & _
      '     "       r.nAnio=" & nAnio & "  and  " & _
      '     "       r.nEstado=1 and  a.nNivelAprobacion=1 "

      '--------------------------------------------------------------
      'Verifica si hay requerimientos grabados ----------------------
      '--------------------------------------------------------------
      nPlanReqActual = 0
      nAprobado = 0
      Set Rs = oPlan.EstadoAprobacionRequerimiento(nAnio, vPersCod, 1, 1)
      If Not Rs.EOF Then
         nPlanReqActual = Rs!nPlanReqNro
         nAprobado = Rs!nEstadoAprobacion
         'Si ya está aprobado |||||||||||||||||||||||||||||||||||||||
         If nAprobado = 1 Then
            nEditable = False
            cmdAgregar.Visible = False
            cmdQuitar.Visible = False
            cmdGrabar.Visible = False
            lblReqNro.Caption = "Requerimiento Nº " + CStr(nPlanReqActual)
            MsgBox "El requerimiento ya fue aprobado !" + Space(10), vbInformation
         Else
            nEditable = True
            cmdAgregar.Visible = True
            cmdQuitar.Visible = True
            cmdGrabar.Visible = True
         End If
      Else
         nEditable = True
         cmdAgregar.Visible = True
         cmdQuitar.Visible = True
         cmdGrabar.Visible = True
      End If
      Set Rs = Nothing
      '--------------------------------------------------------------
      'Obtenemos detalle de requerimientos grabados -----------------
      '--------------------------------------------------------------
      If nPlanReqActual > 0 Then
         Set Rs = oPlan.RequerimientoPlanAnual(nPlanReqActual)
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
               MSFlex.TextMatrix(i, 16) = IIf(Rs!nTotal > 0, Rs!nTotal, "")
               If Rs!nEstado = 0 Then
                  MSFlex.row = i
                  For K = 1 To 15
                      MSFlex.Col = K
                      MSFlex.CellForeColor = "&H8000000F"
                  Next
               End If
               Rs.MoveNext
            Loop
         End If
      End If
   End If
'End If
End Sub

Private Sub mnuInfo_Click()
frmLogPlanAnualInfo.PlanAnual nPlanReqActual
End Sub


Private Sub MSFlex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And MSFlex.Col >= 4 And MSFlex.Col <= 15 And nEditable Then
   MSFlex.TextMatrix(MSFlex.row, MSFlex.Col) = ""
   TotalFila MSFlex.row
End If
End Sub

Private Sub MSFlex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu mnuMenu
End If
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
If Len(txtPersCod) = 13 Then
   DatosPersonal txtPersCod
Else
   If Len(txtPersCod) = 0 Then
      Exit Sub
   End If
   MsgBox "Persona no se puede identificar como personal CMAC-T ..." + Space(10), vbInformation
   FormaFlex
   cmdGrabar.Enabled = False
   cmdAgregar.Visible = False
   cmdQuitar.Visible = False
   txtArea.Text = ""
   txtAgencia.Text = ""
   txtCargo.Text = ""
   txtPersona.Text = ""
End If
End Sub


Sub Totaliza()
Dim i As Integer, j As Integer, n As Integer
Dim nSuma As Currency
n = MSFlex.Rows - 1
For i = 1 To n
    nSuma = 0
    For j = 1 To 12
        nSuma = nSuma + VNumero(MSFlex.TextMatrix(i, j + 3))
    Next
    MSFlex.TextMatrix(i, 16) = nSuma
Next
End Sub

Private Sub cmdAgregar_Click()
Dim i As Integer, Rs As New ADODB.Recordset

i = MSFlex.Rows - 1
If Len(Trim(MSFlex.TextMatrix(i, 1))) = 0 Then
   i = i - 1
End If

frmLogBSSelector.SeleccionBienesCargo cRHCargoCod
Set Rs = frmLogBSSelector.gvrs
If Rs.State <> 0 Then
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         If Not YaEsta(Rs!cProSelBSCod) Then
            i = i + 1
            InsRow MSFlex, i
            MSFlex.TextMatrix(i, 1) = Rs!cProSelBSCod
            MSFlex.TextMatrix(i, 2) = Rs!cBSDescripcion
            MSFlex.TextMatrix(i, 3) = GetBSUnidadLog(Rs!cProSelBSCod)
            If Left(Rs!cProSelBSCod, 2) = "12" Then
               MSFlex.row = i
               MSFlex.Col = 4
               frmLogProSelEspecificaciones.Inicio MSFlex.Left + MSFlex.CellLeft + 120, MSFlex.Top + MSFlex.CellTop + 1720, ""
               MSFlex.TextMatrix(i, 17) = frmLogProSelEspecificaciones.vpTexto
            End If
         End If
         Rs.MoveNext
      Loop
   End If
End If



'frmLogProSelBSSelector.TodosConCheck False
'Set Rs = frmLogProSelBSSelector.gvrs
'If Rs.State <> 0 Then
'   If Not Rs.EOF Then
'      Do While Not Rs.EOF
'         If Not YaEsta(Rs!cProSelBSCod) Then
'            i = i + 1
'            InsRow MSFlex, i
'            MSFlex.TextMatrix(i, 1) = Rs!cProSelBSCod
'            MSFlex.TextMatrix(i, 2) = Rs!cBSDescripcion
'            MSFlex.TextMatrix(i, 3) = GetBSUnidadLog(Rs!cProSelBSCod)
'            If Left(Rs!cProSelBSCod, 2) = "12" Then
'               MSFlex.row = i
'               MSFlex.Col = 4
'               frmLogProSelEspecificaciones.Inicio MSFlex.Left + MSFlex.CellLeft + 120, MSFlex.Top + MSFlex.CellTop + 1720, ""
'               MSFlex.TextMatrix(i, 17) = frmLogProSelEspecificaciones.vpTexto
'            End If
'         End If
'         Rs.MoveNext
'      Loop
'   End If
'End If

End Sub

Private Sub MSflex_DblClick()
If Left(MSFlex.TextMatrix(MSFlex.row, 1), 2) = "12" Then
   MSFlex.row = MSFlex.row
   MSFlex.Col = 4
   'frmLogEspecificaciones.Inicio MSFlex.Left + MSFlex.CellLeft + 120, MSFlex.Top + MSFlex.CellTop + 1720, MSFlex.TextMatrix(MSFlex.row, 17)
   'MSFlex.TextMatrix(MSFlex.row, 17) = frmLogEspecificaciones.vpTexto
End If
End Sub


Function YaEsta(vBSCod As String) As Boolean
Dim i As Integer, n As Integer
YaEsta = False
n = MSFlex.Rows - 1

For i = 1 To n
    If MSFlex.TextMatrix(i, 1) = vBSCod Then
       YaEsta = True
       Exit Function
    End If
Next
End Function

Sub FormaFlex()
Dim i As Integer
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(-1) = 260
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 0:         MSFlex.TextMatrix(0, 1) = ""
MSFlex.ColWidth(2) = 2500:      MSFlex.TextMatrix(0, 2) = "Descripción"
MSFlex.ColWidth(3) = 1200:      MSFlex.TextMatrix(0, 3) = "Unidad":   MSFlex.ColAlignment(3) = 1
For i = 1 To 12
    MSFlex.TextMatrix(0, i + 3) = Space(2) + mMes(i)
    MSFlex.ColWidth(i + 3) = 550:   MSFlex.ColAlignment(i + 3) = 4
Next
MSFlex.ColWidth(16) = 900: MSFlex.TextMatrix(0, 16) = "   TOTAL"
MSFlex.ColWidth(17) = 0
End Sub

'*********************************************************************
'PROCEDIMIENTOS DEL FLEX
'*********************************************************************


'Private Sub MSFlex_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyDelete And MSFlex.Col = 3 Then
'   MSFlex.TextMatrix(MSFlex.Row, 3) = ""
'End If
'End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
If MSFlex.Col >= 3 And MSFlex.Col < 16 And nEditable Then
   EditaFlex MSFlex, txtEdit, KeyAscii
End If
End Sub

Sub EditaFlex(MSFlex As Control, Edt As Control, KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) Then
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
End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If KeyAscii = Asc(vbCr) Then
   KeyAscii = 0
   txtEdit = FNumero(txtEdit)
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
         'If MSFlex.Row < MSFlex.FixedRows - 1 Then
         If MSFlex.row < MSFlex.Rows - 1 Then
            MSFlex.row = MSFlex.row + 1
         End If
End Select
End Sub

Private Sub MSFlex_GotFocus()
If txtEdit.Visible = False Then Exit Sub
MSFlex = txtEdit
txtEdit.Visible = False
TotalFila MSFlex.row
'If MSFlex.Row < MSFlex.Rows - 1 Then
'   MSFlex.Row = MSFlex.Row + 1
'End If
End Sub

Private Sub MSFlex_LeaveCell()
If txtEdit.Visible = False Then Exit Sub
MSFlex = txtEdit
txtEdit.Visible = False
TotalFila MSFlex.row
'If MSFlex.Row < MSFlex.Rows - 1 Then
'   MSFlex.Row = MSFlex.Row + 1
'End If
End Sub

Sub TotalFila(i As Integer)
Dim j As Integer, n As Integer
Dim nSuma As Currency
nSuma = 0
For j = 1 To 12
    nSuma = nSuma + VNumero(MSFlex.TextMatrix(i, j + 3))
Next
MSFlex.TextMatrix(i, 16) = nSuma
End Sub

