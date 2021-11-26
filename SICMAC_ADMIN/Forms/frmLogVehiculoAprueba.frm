VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogVehiculoAprueba 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4875
   ClientLeft      =   420
   ClientTop       =   2730
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   10575
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Detalle"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9180
      TabIndex        =   6
      Top             =   4440
      Width           =   1275
   End
   Begin VB.CommandButton cmdRechazar 
      Height          =   375
      Left            =   1620
      TabIndex        =   5
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdAprobar 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   -60
      Width           =   10335
      Begin VB.ComboBox cboAgencia 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label5 
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
         Left            =   240
         TabIndex        =   3
         Top             =   420
         Width           =   705
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   3555
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6271
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
      HighLight       =   2
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
   Begin VB.Image imgOK 
      Height          =   240
      Left            =   5760
      Picture         =   "frmLogVehiculoAprueba.frx":0000
      Top             =   4500
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNN 
      Height          =   240
      Left            =   5460
      Picture         =   "frmLogVehiculoAprueba.frx":0342
      Top             =   4500
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmLogVehiculoAprueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nEstadoAprueba As Integer, nEstadoBuscar As Integer

Public Sub CambiaEstado(pnEstadoBuscar As Integer, pnEstadoAprueba As Integer)
nEstadoAprueba = pnEstadoAprueba
nEstadoBuscar = pnEstadoBuscar
Me.Show 1
End Sub

Private Sub cmdDetalle_Click()
Dim cMovNro As String, cPersCod As String

cMovNro = MSFlex.TextMatrix(MSFlex.row, 8)
cPersCod = MSFlex.TextMatrix(MSFlex.row, 9)

frmLogVehiculoRegistro.Estado cPersCod, 4, True

End Sub

Private Sub Form_Load()
CentraForm Me
CargaAgencias
'gcSolicitud --> gcAprobado
If nEstadoBuscar = gcSolicitud Then
   Me.Caption = "Vehicular - Aprobación de Solicitudes de Asignación Vehicular"
   cmdAprobar.Caption = "Aprobar"
   cmdRechazar.Caption = "Rechazar"
End If
If nEstadoBuscar = gcAceptado Then
   Me.Caption = "Vehicular - Visto Bueno a Documentos"
   cmdAprobar.Caption = "Visto Bueno"
   cmdRechazar.Caption = "Observación"
   cmdDetalle.Visible = True
End If

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

'Sub CargaMeses()
'Dim i As Integer
'cboMes.Clear
'For i = 1 To 12
'    cboMes.AddItem UCase(mMes(i))
'Next
'cboMes.ListIndex = 0
'End Sub

Sub CargaAgencias()
Dim rs As New ADODB.Recordset, oConn As New DConecta
Dim sSQL As String
If oConn.AbreConexion Then

   sSQL = "Select cAgeCod, cAgeDescripcion from Agencias where nEstado=1"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
   
      cboAgencia.AddItem "TODAS LAS AGENCIAS"
      cboAgencia.ItemData(cboAgencia.ListCount - 1) = 0

      Do While Not rs.EOF
         cboAgencia.AddItem rs!cAgeDescripcion
         cboAgencia.ItemData(cboAgencia.ListCount - 1) = rs!cAgeCod
         rs.MoveNext
      Loop
      cboAgencia.ListIndex = 0
   End If
End If
End Sub

Private Sub cboMes_Click()
GeneraListaAprobacion Format(cboAgencia.ItemData(cboAgencia.ListIndex), "00")
End Sub

Private Sub cboAgencia_Click()
GeneraListaAprobacion Format(cboAgencia.ItemData(cboAgencia.ListIndex), "00")
End Sub

Sub GeneraListaAprobacion(Optional psAgeCod As String = "00")
Dim rs As New ADODB.Recordset, oConn As New DConecta, i As Integer
Dim sSQL As String, cConsulta As String

If Val(psAgeCod) > 0 Then
   cConsulta = " AND c.cAgeCod = '" & psAgeCod & "' "
Else
   cConsulta = ""
End If

FlexLista
sSQL = "select a.*, p.cPersNombre,r.cPlaca, t.cDescripcion,c.cAgeCod,cAgencia=g.cAgeDescripcion " & _
"  from LogVehiculoAsignacion a inner join Persona p on a.cPersCod = p.cPersCod " & _
"     inner join LogVehiculoConductor c on a.cPersCod = c.cPersCod " & _
"     inner join LogVehiculo r on a.nVehiculoCod = r.nVehiculoCod " & _
"     inner join Agencias g on c.cAgeCod = g.cAgeCod " & _
"     inner join (select nConsValor as nTipoVehiculo, cConsDescripcion as cDescripcion " & _
"     from Constante where nConsCod = 9026 and nConsCod<>nConsValor) t on r.nTipoVehiculo = t.nTipoVehiculo " & _
"   WHERE a.nEstado = " & nEstadoBuscar & " " + cConsulta & _
" order by c.cAgeCod,p.cPersNombre "

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 1) = ""  'Reservado para marcar check
         MSFlex.TextMatrix(i, 2) = Replace(rs!cPersNombre, "/", " ")
         MSFlex.TextMatrix(i, 3) = rs!cDescripcion
         MSFlex.TextMatrix(i, 4) = rs!cPlaca
         MSFlex.TextMatrix(i, 5) = rs!dFechaIni
         MSFlex.TextMatrix(i, 6) = IIf(IsNull(rs!dFechaFin), "", rs!dFechaFin)
         MSFlex.TextMatrix(i, 7) = rs!cAgencia
         MSFlex.TextMatrix(i, 8) = rs!nAsignacionNro
         MSFlex.TextMatrix(i, 9) = rs!cPersCod
         MSFlex.TextMatrix(i, 10) = rs!nVehiculoCod
         MSFlex.row = i
         MSFlex.Col = 0
         Set MSFlex.CellPicture = imgNN
         rs.MoveNext
      Loop
   End If
End If
End Sub

Sub FlexLista()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 260:  MSFlex.ColAlignment(0) = 4
MSFlex.ColWidth(1) = 0      'Reservado para marcar check
MSFlex.ColWidth(2) = 3000
MSFlex.ColWidth(3) = 2000
MSFlex.ColWidth(4) = 1000:  MSFlex.ColAlignment(4) = 4
MSFlex.ColWidth(5) = 840:   MSFlex.ColAlignment(5) = 4
MSFlex.ColWidth(6) = 840:   MSFlex.ColAlignment(6) = 4
MSFlex.ColWidth(7) = 2100
MSFlex.ColWidth(8) = 0
MSFlex.ColWidth(9) = 0
MSFlex.ColWidth(10) = 0
End Sub

Private Sub MSflex_DblClick()
Dim i As Integer
i = MSFlex.row
If Len(Trim(MSFlex.TextMatrix(i, 1))) = 0 Then
   MSFlex.TextMatrix(i, 1) = "."
   MSFlex.row = i
   MSFlex.Col = 0
   Set MSFlex.CellPicture = imgOK
Else
   MSFlex.row = i
   MSFlex.Col = 0
   Set MSFlex.CellPicture = imgNN
   MSFlex.TextMatrix(MSFlex.row, 1) = ""
End If
MSFlex.Col = 2
MSFlex.SetFocus
End Sub

Private Sub cmdAprobar_Click()
Dim i As Integer, n As Integer, sSQL As String
Dim oConn As New DConecta, rs As New ADODB.Recordset
Dim sMovNro As String, cPersCod As String, nVehiculoCod As Integer
Dim nAsignaNro As Integer

n = MSFlex.Rows - 1
If MsgBox("¿ Seguro de aprobar las personas indicadas ?" + Space(10), vbQuestion + vbYesNo, "Confirme operación") = vbYes Then
   If oConn.AbreConexion Then
      For i = 1 To n
          nAsignaNro = 0
          If Len(Trim(MSFlex.TextMatrix(i, 1))) > 0 Then
          
             sMovNro = GetLogMovNro
             nAsignaNro = CInt(MSFlex.TextMatrix(i, 8))
             cPersCod = MSFlex.TextMatrix(i, 9)
             nVehiculoCod = MSFlex.TextMatrix(i, 10)
             
             sSQL = "UPDATE LogVehiculoAsignacion SET nEstado = " & nEstadoAprueba & " WHERE nAsignacionNro = " & nAsignaNro & " "
             oConn.Ejecutar sSQL
             
             sSQL = "INSERT INTO LogVehiculoAsignacionMov (nAsignacionNro,cOpeCod,cMovNro,cComentario) " & _
                    "       VALUES (" & nAsignaNro & ", '" & gsOpeCod & "','" & sMovNro & "','') "
             oConn.Ejecutar sSQL
             
             If nEstadoAprueba = gcVistoBueno Then
             
                sSQL = "UPDATE LogVehiculo SET nEstado = " & gcDisponible & " WHERE nVehiculoCod = " & nVehiculoCod & " "
                oConn.Ejecutar sSQL
                
                sSQL = "UPDATE LogVehiculoConductor SET nEstado = " & gcDisponible & " WHERE cPersCod = '" & cPersCod & "' "
                oConn.Ejecutar sSQL
                
             End If
             
          End If
      Next
      GeneraListaAprobacion Format(cboAgencia.ItemData(cboAgencia.ListIndex), "00")
   End If
End If
End Sub

