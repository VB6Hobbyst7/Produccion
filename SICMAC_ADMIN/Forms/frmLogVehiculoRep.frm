VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmLogVehiculoRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   5025
   ClientLeft      =   2040
   ClientTop       =   2385
   ClientWidth     =   7710
   Icon            =   "frmLogVehiculoRep.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7710
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   7575
      Begin VB.TextBox txtVehiculo 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   360
         Width           =   5475
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   6285
      End
      Begin VB.CommandButton cmdBuscaVehiculo 
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
         Left            =   1500
         TabIndex        =   1
         Top             =   390
         Width           =   340
      End
      Begin VB.TextBox txtVehiculoCod 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   780
      End
      Begin MSMask.MaskEdBox txtFecIni 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1080
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecFin 
         Height          =   315
         Left            =   3180
         TabIndex        =   4
         Top             =   1080
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Top             =   1140
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Reorte de"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   780
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vehículo"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   420
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3435
      Left            =   60
      TabIndex        =   11
      Top             =   1500
      Width           =   7575
      Begin VB.CheckBox chkPie 
         Caption         =   "Sectores"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkGrafico 
         Caption         =   "Gráfico"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   6120
         TabIndex        =   7
         Top             =   3000
         Width           =   1275
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   3000
         Width           =   1275
      End
      Begin MSComctlLib.ListView lsvRep 
         Height          =   2715
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   4789
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripción"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Frame fraGraf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         ForeColor       =   &H80000008&
         Height          =   2715
         Left            =   1080
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   6315
         Begin MSChart20Lib.MSChart Graf 
            Height          =   2595
            Left            =   60
            OleObjectBlob   =   "frmLogVehiculoRep.frx":08CA
            TabIndex        =   16
            Top             =   60
            Width           =   6195
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Selección"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmLogVehiculoRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtFecIni = Date
txtFecFin = Date
CargaCombo
End Sub

Private Sub cmdBuscaVehiculo_Click()
Dim sSQL As String, v As New DLogVehiculos

sSQL = "select r.nVehiculoCod,t.cTipoVehiculo +' '+ m.cMarca +' ['+ r.cPlaca +']  - '+ r.cModelo from LogVehiculo r " & _
       "  inner join (select nConsValor AS nTipoVehiculo,cConsDescripcion as cTipoVehiculo from Constante where nConsCod=9026 and nconscod<>nconsvalor) t on r.nTipoVehiculo = t.nTipoVehiculo " & _
       " inner join (select nConsValor AS nMarca,cConsDescripcion as cMarca from Constante where nConsCod=9022 and nconscod<>nconsvalor) m on r.nMarca = m.nMarca "
frmLogSelector.Consulta sSQL, "Seleccione Vehiculo"
If frmLogSelector.vpHaySeleccion Then
   txtVehiculoCod = frmLogSelector.vpCodigo
   txtVehiculo = frmLogSelector.vpDescripcion
End If
End Sub

Sub CargaCombo()
Dim oConn As DConecta
Dim rs As New ADODB.Recordset
Dim sSQL As String

Set oConn = New DConecta
If oConn.AbreConexion Then
   sSQL = "select nConsValor,cConsDescripcion from Constante where nConsCod =9024 and nconscod<>nconsvalor "
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         cboTipo.AddItem Format(rs!nConsValor, "00") + ". " + rs!cConsDescripcion
         rs.MoveNext
      Loop
   End If
   cboTipo.ListIndex = 0
   Set rs = Nothing
End If
End Sub

Private Sub cboTipo_Click()
Dim nTipo As Integer
nTipo = CInt(Mid(cboTipo.Text, 1, 2))
Select Case nTipo
    Case 0
         lsvRep.ListItems.Clear
    Case 1
         lsvRep.ListItems.Clear
    Case 2
         GeneraLista
    Case 3
         GeneraListaIncidencias
End Select
End Sub

Sub GeneraLista()
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim sSQL As String

lsvRep.ListItems.Clear

sSQL = "SELECT dFecha,cDescripcion,Origen=coalesce(o.cUbigeoDescripcion,''),Destino=coalesce(d.cUbigeoDescripcion,'') " & _
       "  FROM LogVehiculoAsignacionDet m " & _
       "  left outer join UbicacionGeografica o on m.cValor1=o.cUbigeoCod " & _
       "  left outer join UbicacionGeografica d on m.cValor2=d.cUbigeoCod " & _
       " Where nTipoReg = 2 " & _
       " "
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         lsvRep.ListItems.Add
         lsvRep.ListItems(lsvRep.ListItems.Count).SubItems(1) = rs!dFecha
         lsvRep.ListItems(lsvRep.ListItems.Count).SubItems(2) = rs!cDescripcion
         lsvRep.ListItems(lsvRep.ListItems.Count).Checked = True
         rs.MoveNext
      Loop
   End If
End If
End Sub

Sub GeneraListaIncidencias()
Dim oConn As DConecta
Dim rs As New ADODB.Recordset
Dim sSQL As String

Set oConn = New DConecta

lsvRep.ListItems.Clear
If oConn.AbreConexion Then
   sSQL = "select nConsValor,cConsDescripcion from Constante where nConsCod =9025 and nconscod<>nconsvalor"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         lsvRep.ListItems.Add
         lsvRep.ListItems(lsvRep.ListItems.Count).SubItems(1) = rs!nConsValor
         lsvRep.ListItems(lsvRep.ListItems.Count).SubItems(2) = rs!cConsDescripcion
         lsvRep.ListItems(lsvRep.ListItems.Count).Checked = True
         rs.MoveNext
      Loop
   End If
   Set rs = Nothing
   oConn.CierraConexion
End If
End Sub

Private Sub cmdImprimir_Click()
Dim rs As New ADODB.Recordset, oConn As DConecta, sSQL As String
Dim f As Integer, i As Integer, v As Variant, cArchivo As String
Dim n As Integer, nTipoReg As Integer, nVehiculoCod As Integer
Dim cValores As String, nSuma As Currency

Set oConn = New DConecta

If Not oConn.AbreConexion Then
   MsgBox "No se puede establecer la conexión..." + Space(10), vbInformation
   Exit Sub
End If

f = FreeFile
Close
cArchivo = App.path + "\RepVehic.txt"
Open cArchivo For Output As #f

Print #f, "REPORTE DE " + cboTipo.Text
Print #f, ""
Print #f, String(70, "=")
Print #f, "Fecha      Descripcion" + Space(35) + " Monto"
Print #f, String(70, "-")

n = lsvRep.ListItems.Count
cValores = ""
For i = 1 To n
    If lsvRep.ListItems(i).Checked Then
       cValores = cValores + lsvRep.ListItems(i).SubItems(1) + ","
    End If
Next
nVehiculoCod = VNumero(txtVehiculoCod)
If Len(cValores) = 0 Then Exit Sub
cValores = Mid(cValores, 1, Len(cValores) - 1)
nSuma = 0

'sSQL = "select d.dFecha,d.cDescripcion,d.nMonto, p.cPersNombre as cConductor" & _
'       " from LogVehiculoMovDet d  inner join LogVehiculoMov m on d.cMovNro = m.cMovNro " & _
'       " left join Persona p on m.cPersCod = p.cPersCod   " & _
'       " where nTipoReg = 3 and m.nVehiculoCod=" & nVehiculoCod & " and " & _
'       "       d.dFecha >= '" & Format(txtFecIni, "YYYYMMDD") & "' and " & _
'       "       d.dFecha <= '" & Format(txtFecFin, "YYYYMMDD") & "' and  " & _
'       "       convert(int,cValor0) in (" & cValores & ") " & _
'       " order by d.dFecha "
       
sSQL = "select d.nAsignacionNro, d.dFecha, d.cDescripcion, d.nMonto, a.cPersCod, p.cPersNombre as cConductor " & _
"  from LogVehiculoAsignacionDet d inner join LogVehiculoAsignacion a on d.nAsignacionNro = a.nAsignacionNro " & _
"       inner join Persona p on a.cPersCod = p.cPersCod " & _
" where d.nTipoReg = 3 and a.nVehiculoCod = " & nVehiculoCod & " and " & _
"      d.dFecha >= '" & Format(txtFecIni, "YYYYMMDD") & "' and " & _
"      d.dFecha <= '" & Format(txtFecFin, "YYYYMMDD") & "' and " & _
"     convert(int,d.cValor0) in (" & cValores & ") " & _
" order by d.dFecha"
       
Set rs = oConn.CargaRecordSet(sSQL)
If Not rs.EOF Then
   Do While Not rs.EOF
      Print #f, CStr(rs!dFecha) + " " + JIZQ(rs!cDescripcion + " / " + rs!cConductor, 40) + "  " + JDER(Format(rs!nMonto, "##,##0.00"), 12)
      'Print #f, Space(11) + rs!cConductor
      nSuma = nSuma + rs!nMonto
      rs.MoveNext
   Loop
End If
Print #f, String(70, "-")
Print #f, "TOTAL ........." + Space(38) + JDER(Format(nSuma, "###,##0.00"), 12)
Print #f, String(70, "=")
Close #f
v = Shell("notepad " & cArchivo & "", vbNormalFocus)
End Sub

Private Sub chkGrafico_Click()
If chkGrafico.value = 1 Then
   lsvRep.Visible = False
   chkPie.value = 0
   chkPie.Visible = True
   fraGraf.Visible = True
   FormaGrafico
Else
   lsvRep.Visible = True
   chkPie.Visible = False
   fraGraf.Visible = False
End If
End Sub

Sub FormaGrafico()
Dim i As Integer, n As Integer, nMax As Integer
Dim nTipoIncCod As Integer
Dim nMaximo As Currency
Dim nMonto As Currency

   Graf.RowCount = 0
   Graf.ColumnCount = 0
   Graf.ChartType = VtChChartType2dBar
   Graf.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
   Graf.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 50
   Graf.ShowLegend = True
   
   n = 0
   nMaximo = 0
   nMax = lsvRep.ListItems.Count
   For i = 1 To nMax
       If lsvRep.ListItems(i).Checked Then
          n = n + 1
       End If
   Next
   Graf.ColumnCount = n
   Graf.RowCount = 1
   n = 0
   For i = 1 To nMax
       If lsvRep.ListItems(i).Checked Then
          n = n + 1
          nTipoIncCod = lsvRep.ListItems(i).SubItems(1)
          Graf.Column = n
          Graf.row = 1
          nMonto = GetMontoItem(txtVehiculoCod, nTipoIncCod, txtFecIni, txtFecFin)
          If nMonto > nMaximo Then
             nMaximo = nMonto
          End If
          Graf.data = nMonto
          Graf.ColumnLabel = Mid(lsvRep.ListItems(i).SubItems(2), 1, 20)
       End If
   Next
   Graf.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = nMaximo + 20
End Sub

Private Sub chkPie_Click()
If chkPie.value = 1 Then
   Graf.ChartType = VtChChartType2dPie
Else
   Graf.ChartType = VtChChartType2dBar
End If
End Sub

Function GetMontoItem(vVehiculoCod As Integer, vTipoIncCod As Integer, vFecIni As Date, vFecFin As Date) As Currency
Dim oConn As DConecta, sSQL As String
Dim rs As New ADODB.Recordset

GetMontoItem = 0
Set oConn = New DConecta
If oConn.AbreConexion Then

   sSQL = "select nTotal=coalesce(sum(d.nMonto),0)  from LogVehiculoAsignacionDet d " & _
       "  inner join LogVehiculoAsignacion a on d.nAsignacionNro = a.nAsignacionNro " & _
       " where d.nTipoReg = 3 and a.nVehiculoCod=" & vVehiculoCod & " and " & _
       "       d.dFecha >= '" & Format(vFecIni, "YYYYMMDD") & "' and " & _
       "       d.dFecha <= '" & Format(vFecFin, "YYYYMMDD") & "' and  " & _
       "       convert(int,d.cValor0) = " & vTipoIncCod & " "
       
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      GetMontoItem = rs!nTotal
   Else
      GetMontoItem = 0
   End If
End If
End Function
