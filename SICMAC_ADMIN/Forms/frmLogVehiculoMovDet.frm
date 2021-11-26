VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogVehiculoMovDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control vehicular - Detalle de Movimientos vehiculares"
   ClientHeight    =   5715
   ClientLeft      =   1530
   ClientTop       =   2145
   ClientWidth     =   8730
   Icon            =   "frmLogVehiculoMovDet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8730
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   8475
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "Procesar"
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
         Left            =   6600
         TabIndex        =   7
         Top             =   680
         Width           =   1695
      End
      Begin VB.TextBox txtFechaFin 
         Height          =   315
         Left            =   3900
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtFechaIni 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox txtVehiculo 
         BackColor       =   &H8000000F&
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   8115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Movimientos desde                         hasta"
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
         Top             =   780
         Width           =   3600
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSMov 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDet 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   3420
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
      HighLight       =   2
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
End
Attribute VB_Name = "frmLogVehiculoMovDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nVehiculoCod As Integer, cVehiculo As String

Public Sub Vehiculo(ByVal pnVehiculoCod As Integer, ByVal psVehiculo As String)
nVehiculoCod = pnVehiculoCod
cVehiculo = psVehiculo
Me.Show 1
End Sub

Private Sub cmdProcesar_Click()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim sSQL As String, i As Integer

FormaFlexMov

sSQL = "select m.nAsignacionNro,m.dFechaIni,m.dFechaFin,m.cPersCod,cPersona=coalesce(p.cPersNombre,''), e.cEstado " & _
       "  from LogVehiculoAsignacion m left outer join Persona p on m.cPersCod=p.cPersCod " & _
       "       left outer join (select nConsValor as nEstado,cConsDescripcion as cEstado from Constante where nConsCod =9020 and nconscod<>nconsvalor) e on m.nEstado=e.nEstado " & _
       " where m.nVehiculoCod = " & nVehiculoCod & " and m.dFechaIni>='" & Format(txtFechaIni, "YYYYMMDD") & "' and m.dFechaFin<='" & Format(txtFechaFin, "YYYYMMDD") & "'"
i = 0
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSMov, i
         MSMov.TextMatrix(i, 1) = rs!nAsignacionNro
         MSMov.TextMatrix(i, 2) = rs!dFechaIni
         MSMov.TextMatrix(i, 3) = rs!dFechaFin
         MSMov.TextMatrix(i, 4) = rs!cPersCod
         MSMov.TextMatrix(i, 5) = IIf(Len(rs!cPersona) > 0, rs!cPersona, rs!cEstado)
         rs.MoveNext
      Loop
      MSMov.SetFocus
   End If
End If

End Sub

Private Sub Form_Load()
CentraForm Me
txtVehiculo.Text = cVehiculo
txtFechaIni.Text = Date
txtFechaFin.Text = Date
FormaFlexMov
FormaFlexDet
End Sub

'------------------------------------------------------------------------------

Sub FormaFlexMov()
MSMov.Clear
MSMov.Rows = 2
MSMov.RowHeight(0) = 320
MSMov.RowHeight(1) = 8
MSMov.ColWidth(0) = 0
MSMov.ColWidth(1) = 380
MSMov.ColWidth(2) = 900: MSMov.TextMatrix(0, 2) = "Inicio":  MSMov.ColAlignment(2) = 4
MSMov.ColWidth(3) = 900: MSMov.TextMatrix(0, 3) = "Término":  MSMov.ColAlignment(3) = 4
MSMov.ColWidth(4) = 0
MSMov.ColWidth(5) = 6000: MSMov.TextMatrix(0, 5) = "Conductor"
MSMov.ColWidth(6) = 0
MSMov.ColWidth(7) = 0
End Sub

Private Sub MSMov_GotFocus()
If Len(Trim(MSMov.TextMatrix(MSMov.row, 1))) > 0 Then
   GeneraDetalleMov MSMov.TextMatrix(MSMov.row, 1)
End If
End Sub

Private Sub MSMov_RowColChange()
If Len(Trim(MSMov.TextMatrix(MSMov.row, 1))) > 0 Then
   GeneraDetalleMov MSMov.TextMatrix(MSMov.row, 1)
End If
End Sub

Sub GeneraDetalleMov(ByVal pnAsignacionNro As Integer)
Dim DLog As New DLogVehiculos
Dim rs As New ADODB.Recordset
Dim i As Integer

FormaFlexDet
Set rs = DLog.GetVehiculoMovDet(pnAsignacionNro)
If Not rs.EOF Then
   i = 0
   Do While Not rs.EOF
      i = i + 1
      InsRow MSDet, i
      MSDet.TextMatrix(i, 0) = rs!dFecha
      MSDet.TextMatrix(i, 1) = rs!cRegistro
      MSDet.TextMatrix(i, 2) = rs!cDescripcion
      If rs!nTipoReg = 4 Then
         MSDet.TextMatrix(i, 3) = rs!cValor1
         MSDet.TextMatrix(i, 4) = rs!cValor2
      Else
         MSDet.TextMatrix(i, 3) = rs!cDesc1
         MSDet.TextMatrix(i, 4) = rs!cDesc2
      End If
      MSDet.TextMatrix(i, 5) = FNumero(rs!nMonto)
      rs.MoveNext
   Loop
End If
End Sub

Sub FormaFlexDet()
MSDet.Clear
MSDet.Rows = 2
MSDet.RowHeight(0) = 320
MSDet.RowHeight(1) = 8
MSDet.ColWidth(0) = 800: MSDet.TextMatrix(0, 0) = "Fecha"
MSDet.ColWidth(1) = 2100: MSDet.TextMatrix(0, 1) = "Operación"
MSDet.ColWidth(2) = 2100: MSDet.TextMatrix(0, 2) = "Descripción"
MSDet.ColWidth(3) = 1200: MSDet.TextMatrix(0, 3) = "Lugar/Origen": MSDet.ColAlignment(3) = 1
MSDet.ColWidth(4) = 1200: MSDet.TextMatrix(0, 4) = "Destino": MSDet.ColAlignment(4) = 1
MSDet.ColWidth(5) = 800:  MSDet.TextMatrix(0, 5) = "Monto"
MSDet.ColWidth(6) = 0
MSDet.ColWidth(7) = 0
MSDet.ColWidth(8) = 0
MSDet.ColWidth(9) = 0
End Sub


'------------------------------------------------------------------------------

Private Sub txtFechaFin_GotFocus()
txtFechaFin.SelStart = 0
txtFechaFin.SelLength = Len(Trim(txtFechaFin))
End Sub

Private Sub txtFechaFin_LostFocus()
If CDate(txtFechaFin) < CDate(txtFechaIni) Then
   MsgBox "La fecha final es incorrecta..." + Space(10), vbInformation, "Error en dato"
   txtFechaFin.SetFocus
End If
End Sub

Private Sub txtFechaIni_GotFocus()
txtFechaIni.SelStart = 0
txtFechaIni.SelLength = Len(Trim(txtFechaIni))
End Sub

Private Sub txtFechaIni_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigFecha(txtFechaIni, KeyAscii)
If nKeyAscii = 13 Then
   txtFechaFin.SetFocus
End If
End Sub

Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigFecha(txtFechaFin, KeyAscii)
If nKeyAscii = 13 Then
   'txtVehiculoCod.SetFocus
End If
End Sub

