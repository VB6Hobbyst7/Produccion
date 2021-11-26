VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogVehiculoSol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Vehicular - Solicitud de Asignación Vehicular"
   ClientHeight    =   5490
   ClientLeft      =   525
   ClientTop       =   2460
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   11190
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   11055
      Begin VB.TextBox txtAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   7620
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   420
         Width           =   2955
      End
      Begin VB.TextBox txtPersona 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   5115
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agencia"
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
         Left            =   6780
         TabIndex        =   7
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Administrador"
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
         TabIndex        =   6
         Top             =   420
         Width           =   1155
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   315
         Left            =   1380
         Top             =   360
         Width           =   5235
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00EAFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   315
         Left            =   7560
         Top             =   360
         Width           =   3075
      End
   End
   Begin VB.CommandButton cmdAsignaSi 
      Caption         =   "Asignar Vehículo"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   5040
      Width           =   1635
   End
   Begin VB.CommandButton cmdAsignaNo 
      Caption         =   "Quitar Asignación"
      Height          =   375
      Left            =   1740
      TabIndex        =   1
      Top             =   5040
      Width           =   1635
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      Top             =   5040
      Width           =   1275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   4035
      Left            =   60
      TabIndex        =   8
      Top             =   900
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7117
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
      HighLight       =   2
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
      _Band(0).Cols   =   11
   End
End
Attribute VB_Name = "frmLogVehiculoSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cAgeCod As String
Dim nModo As Integer

Public Sub Modalidad(pnModo As Integer)
nModo = pnModo
Me.Show 1
End Sub

Private Sub Form_Load()
Dim oConn As New DConecta, Rs As New ADODB.Recordset
CentraForm Me
IdentificaUsuario gsCodPersUser
'---------------------------------------------------------------
'If oConn.AbreConexion Then
'   Set rs = oConn.CargaRecordSet("select nConsValor,cConsDescripcion from Constante where nConsCod =9021 and nconscod<>nconsvalor order by nConsValor")
'   oConn.CierraConexion
'   If Not rs.EOF Then
'      Do While Not rs.EOF
'         cboMovil.AddItem rs!cConsDescripcion
'         cboMovil.ItemData(cboMovil.ListCount - 1) = rs!nConsValor
'         rs.MoveNext
'      Loop
'      cboMovil.ListIndex = 0
'   Else
'      MsgBox "Faltan definir tipos de vehiculos en la tabla Constante..." + Space(10), vbInformation
'   End If
'End If
'---------------------------------------------------------------
If Len(Trim(cAgeCod)) = 0 Then
   MsgBox "Ud. no puede realizar solicitudes de Asignación Vehicular..." + Space(10), vbInformation
   cmdAsignaSi.Visible = False
   cmdAsignaNo.Visible = False
   FlexLista
Else
   ListaConductores cAgeCod, nModo
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub IdentificaUsuario(ByVal psPersCod As String)
Dim Rs As New ADODB.Recordset

txtPersona.Text = ""
txtAgencia.Text = ""
cAgeCod = ""

Set Rs = GetAreaCargoAgencia(psPersCod)
If Not Rs.EOF Then
   txtPersona.Text = Rs!cPersona
   txtAgencia.Text = Rs!cRHAgencia
   cAgeCod = Rs!cRHAgeCod
End If
End Sub

Sub ListaConductores(ByVal psAgeCod As String, ByVal pnModo As Integer)
Dim sSQL As String, Rs As New ADODB.Recordset, oConn As New DConecta
Dim i As Integer

FlexLista
If Len(psAgeCod) = 0 Then Exit Sub


If pnModo = 1 Then
   sSQL = "select a.*, p.cPersNombre,r.cPlaca, t.cDescripcion, c.cAgeCod, e.cEstado " & _
   "  from LogVehiculoAsignacion a inner join Persona p on a.cPersCod = p.cPersCod " & _
   "     inner join LogVehiculo r on a.nVehiculoCod = r.nVehiculoCod " & _
   "     inner join LogVehiculoConductor c on a.cPersCod = c.cPersCod " & _
   "     inner join (select nConsValor as nTipoVehiculo, cConsDescripcion as cDescripcion " & _
   "                   from Constante where nConsCod = " & gcTipoVehiculo & " and nConsCod<>nConsValor) t on r.nTipoVehiculo = t.nTipoVehiculo " & _
   "     inner join (select nConsValor as nEstado, cConsDescripcion as cEstado " & _
   "                   from Constante where nConsCod = " & gcAsignacionEstado & " and nConsCod<>nConsValor) e on a.nEstado = e.nEstado " & _
   " WHERE a.nEstado = " & gcSolicitud & " and c.cAgeCod = '" & psAgeCod & "' "

Else
   sSQL = "select a.*, p.cPersNombre,r.cPlaca, t.cDescripcion, c.cAgeCod, e.cEstado " & _
   "  from LogVehiculoAsignacion a inner join Persona p on a.cPersCod = p.cPersCod " & _
   "     inner join LogVehiculo r on a.nVehiculoCod = r.nVehiculoCod " & _
   "     inner join LogVehiculoConductor c on a.cPersCod = c.cPersCod " & _
   "     inner join (select nConsValor as nTipoVehiculo, cConsDescripcion as cDescripcion " & _
   "                   from Constante where nConsCod = " & gcTipoVehiculo & " and nConsCod<>nConsValor) t on r.nTipoVehiculo = t.nTipoVehiculo " & _
   "     inner join (select nConsValor as nEstado, cConsDescripcion as cEstado " & _
   "                   from Constante where nConsCod = " & gcAsignacionEstado & " and nConsCod<>nConsValor) e on a.nEstado = e.nEstado " & _
   " WHERE a.nEstado < " & gcVistoBueno & " and c.cAgeCod = '" & psAgeCod & "' "
End If

If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      i = 0
      Do While Not Rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 0) = Format(i, "00")
         MSFlex.TextMatrix(i, 1) = Rs!nAsignacionNro
         MSFlex.TextMatrix(i, 2) = Rs!cPersCod
         MSFlex.TextMatrix(i, 3) = Rs!nVehiculoCod
         MSFlex.TextMatrix(i, 4) = Replace(Rs!cPersNombre, "/", " ")
         MSFlex.TextMatrix(i, 5) = Rs!cDescripcion
         MSFlex.TextMatrix(i, 6) = Rs!cPlaca
         MSFlex.TextMatrix(i, 7) = Rs!cEstado
         
         'MSFlex.TextMatrix(i, 7) = IIf(rs!nAsignacionTpo = 1, "INDEFINIDA", "TEMPORAL")
         
         MSFlex.TextMatrix(i, 8) = Rs!dFechaIni
         If Rs!nAsignacionTpo = 1 Then
            MSFlex.TextMatrix(i, 9) = "INDEFINIDA"
         Else
            MSFlex.TextMatrix(i, 9) = IIf(IsNull(Rs!dFechaFin), "", CStr(Rs!dFechaFin)) + " - TEMPORAL"
         End If
         
         'MSFlex.TextMatrix(i, 9) = IIf(IsNull(rs!dFechaFin), "", rs!dFechaFin)
         'MSFlex.TextMatrix(i, 10) = rs!cEstado
         Rs.MoveNext
      Loop
   End If
End If
End Sub

Sub FlexLista()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 240:      MSFlex.ColAlignment(0) = 4
MSFlex.ColWidth(1) = 0
MSFlex.ColWidth(2) = 0
MSFlex.ColWidth(3) = 0
MSFlex.ColWidth(4) = 3200:  MSFlex.TextMatrix(0, 4) = "Persona"
MSFlex.ColWidth(5) = 2100:  MSFlex.TextMatrix(0, 5) = "Vehículo"
MSFlex.ColWidth(6) = 1000:  MSFlex.TextMatrix(0, 6) = "Placa":      MSFlex.ColAlignment(6) = 4
MSFlex.ColWidth(7) = 2200:  MSFlex.TextMatrix(0, 7) = " Estado":
MSFlex.ColWidth(8) = 840:   MSFlex.TextMatrix(0, 8) = " Desde":   MSFlex.ColAlignment(8) = 4
MSFlex.ColWidth(9) = 1200:  MSFlex.TextMatrix(0, 9) = " Hasta"    'MSFlex.ColAlignment(9) = 4
MSFlex.ColWidth(10) = 0
End Sub

Private Sub cmdAsignaNo_Click()
Dim oConn As New DConecta
Dim nAsignaNro As Integer
Dim cPersCod  As String
Dim nVehiculoCod As Integer

Dim sSQL1 As String
Dim sSQL2 As String
Dim sSQL3 As String

nAsignaNro = CInt(VNumero(MSFlex.TextMatrix(MSFlex.row, 1)))
cPersCod = MSFlex.TextMatrix(MSFlex.row, 2)
nVehiculoCod = MSFlex.TextMatrix(MSFlex.row, 3)

If MsgBox("¿ Está seguro de quitar la asignación ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   
   sSQL1 = "UPDATE LogVehiculoAsignacion SET nEstado = 0 WHERE nAsignacionNro = " & nAsignaNro & " "
   sSQL2 = "UPDATE LogVehiculoConductor  SET nEstado = 1 WHERE cPersCod = '" & cPersCod & "' "
   sSQL3 = "UPDATE LogVehiculo           SET nEstado = 1 WHERE nVehiculoCod = '" & nVehiculoCod & "' "
   
   If oConn.AbreConexion Then
      oConn.Ejecutar sSQL1
      oConn.Ejecutar sSQL2
      oConn.Ejecutar sSQL3
      oConn.CierraConexion
   End If
   ListaConductores cAgeCod, nModo
End If
End Sub

Private Sub cmdAsignaSi_Click()
frmLogVehiculoSolDet.Agencia cAgeCod
If frmLogVehiculoSolDet.vpGrabado Then
   ListaConductores cAgeCod, nModo
End If
End Sub



