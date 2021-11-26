VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogProSelContratoDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ejecucion del Contrato"
   ClientHeight    =   6405
   ClientLeft      =   1695
   ClientTop       =   1875
   ClientWidth     =   9870
   Icon            =   "frmLogProSelContratoDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9870
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   8400
      TabIndex        =   18
      Top             =   5940
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Condiciones de Entrega "
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
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   9615
      Begin VB.TextBox txtCantidadFaltante 
         Height          =   315
         Left            =   6540
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   780
         Width           =   915
      End
      Begin VB.CommandButton cmdEntrega 
         Caption         =   "Registrar Entrega"
         Height          =   375
         Left            =   7920
         TabIndex        =   5
         Top             =   750
         Width           =   1575
      End
      Begin VB.TextBox txtFechaEnt 
         Height          =   315
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   4
         Top             =   780
         Width           =   1155
      End
      Begin VB.TextBox txtCantidad 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   780
         Width           =   675
      End
      Begin VB.ComboBox cboAgencia 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   8535
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   8160
         TabIndex        =   6
         Top             =   750
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Falta"
         Height          =   195
         Index           =   1
         Left            =   6120
         TabIndex        =   20
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha entrega"
         Height          =   195
         Index           =   0
         Left            =   2700
         TabIndex        =   14
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agencia"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   420
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   9615
      Begin VB.TextBox txtFecha 
         Height          =   315
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   0
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox txtNro 
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
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox txtPersona 
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
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   660
         Width           =   8175
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Contrato"
         Height          =   195
         Left            =   6900
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Contrato Nº"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Postor"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   720
         Width           =   450
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSItem 
      Height          =   1680
      Left            =   120
      TabIndex        =   1
      Top             =   1140
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   2963
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   5
      FixedCols       =   0
      ForeColorFixed  =   -2147483646
      BackColorSel    =   14151167
      ForeColorSel    =   -2147483630
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
      _Band(0).Cols   =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDet 
      Height          =   1800
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3175
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   7
      FixedCols       =   0
      ForeColorFixed  =   -2147483646
      BackColorSel    =   14343900
      ForeColorSel    =   -2147483630
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
End
Attribute VB_Name = "frmLogProSelContratoDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nProselNro As Integer, nProSelItem As Integer, cPersCod As String, cPersona As String, nMonto As Currency
Dim sSQL As String

Public Sub Contrato(vProSelNro As Integer, vPersCod As String, vPersona As String, vMonto As Currency, vProSelItem As Integer)
nProselNro = vProSelNro
cPersCod = vPersCod
cPersona = vPersona
nProSelItem = vProSelItem
nMonto = vMonto
Me.Show 1
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim oConn As New DConecta, Rs As New ADODB.Recordset

CentraForm Me
CargaAgencias
txtFecha = gdFecSis
txtFechaEnt = gdFecSis
txtPersona = cPersona
GeneraItemsPostor nProselNro, cPersCod, nProSelItem
txtNro = GetContratoNro(cPersCod, nProselNro, nProSelItem)
If VNumero(txtNro) = 0 Then
   If MsgBox("El postor no tiene un Contrato registrado " + Space(10) + vbCrLf + "¿ Generar un Contrato para el Postor indicado ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   
      If oConn.AbreConexion Then
         sSQL = "INSERT INTO LogProSelContrato (nProSelNro, nProSelItem, cPersCod, dContratoFecha ) " & _
              " VALUES (" & nProselNro & "," & nProSelItem & ",'" & cPersCod & "','" & Format(txtFecha.Text, "YYYYMMDD") & "')"
         oConn.Ejecutar sSQL
         Set Rs = oConn.CargaRecordSet("Select nUltNro = @@identity from LogProSelContrato")
         If Not Rs.EOF Then
            txtNro = Rs!nUltNro
         End If
         oConn.CierraConexion
      End If
   Else
      cboAgencia.BackColor = "&H8000000F":   cboAgencia.Locked = True
      txtFechaEnt.BackColor = "&H8000000F":  txtFechaEnt.Locked = True
      txtCantidad.BackColor = "&H8000000F":  txtCantidad.Locked = True
      cmdEntrega.Enabled = False
      cmdCancelar.Caption = "Salir"
   End If
Else
  MsgBox "Ya existe un contrato para el postor..." + Space(10), vbInformation
  txtFecha.Enabled = False
  MSItem.TabIndex = 0
End If
FormaFlexDetalle
End Sub

Private Sub cmdCancelar_Click()
If cboAgencia.ListCount > 0 Then
   cboAgencia.ListIndex = 0
End If
txtCantidad.Text = ""
txtFechaEnt.Text = gdFecSis
End Sub

Private Sub cmdEntrega_Click()
Dim nContratoNro As Integer, nItem As Integer, cProSelBSCod As String
Dim oConn As New DConecta, cAgeCod As String
Dim nFilas As Integer
Dim nSuma As Integer
Dim nMonto As Integer
Dim nTotal As Integer ', nfila As Integer

nSuma = 0

nFilas = MSDet.Rows - 1
For nItem = 1 To nFilas
    If MSDet.TextMatrix(nItem, 3) = MSItem.TextMatrix(MSItem.row, 1) Then
        nSuma = nSuma + VNumero(MSDet.TextMatrix(nItem, 5))
    End If
Next

sSQL = ""
nContratoNro = CInt(txtNro)
nItem = MSItem.TextMatrix(MSItem.row, 0)
cProSelBSCod = MSItem.TextMatrix(MSItem.row, 1)
nTotal = MSItem.TextMatrix(MSItem.row, 3)

If VNumero(txtCantidad) <= 0 Then
   MsgBox "Debe ingresar una cantidad válida..." + Space(10), vbInformation
   Exit Sub
End If

If nSuma + VNumero(txtCantidad) > nTotal Then
   MsgBox "No se puede agregar una cantidad mayor al Total del Item..." + Space(10), vbInformation
   Exit Sub
End If

cAgeCod = Format(cboAgencia.ItemData(cboAgencia.ListIndex), "00")

If Len(cAgeCod) = 0 Then
   MsgBox "No se indica la agencia..." + Space(10), vbInformation
   Exit Sub
End If

If MsgBox("¿ Seguro de registrar la entrega ?" + Space(10), vbQuestion + vbYesNo, "Confirme operación") = vbYes Then

   sSQL = "INSERT INTO LogProSelContratoEntrega (nProSelConNro, nProSelNro, nProSelItem, cPersCod, cBSCod, nCantidad, dFechaEntrega, cAgeCod ) " & _
          " VALUES (" & nContratoNro & "," & nProselNro & "," & nItem & ",'" & cPersCod & "','" & cProSelBSCod & "'," & VNumero(txtCantidad) & ",'" & Format(txtFechaEnt.Text, "YYYYMMDD") & "','" & cAgeCod & "')"
          
   If oConn.AbreConexion Then
      oConn.Ejecutar sSQL
   End If
   cmdCancelar_Click
   GeneraDetalleItem nProselNro, nItem
End If
End Sub

Sub CargaAgencias()
Dim oConn As New DConecta, Rs As New ADODB.Recordset

sSQL = "Select cAgeCod, cAgeDescripcion from Agencias where nEstado=1 "
If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
Else
   MsgBox "No se puede establecer conexión..." + Space(10), vbInformation
   Exit Sub
End If

If Not Rs.EOF Then
   Do While Not Rs.EOF
      cboAgencia.AddItem Rs!cAgeDescripcion
      cboAgencia.ItemData(cboAgencia.ListCount - 1) = Rs!cAgeCod
      Rs.MoveNext
   Loop
   cboAgencia.ListIndex = 0
End If
End Sub

Sub GeneraItemsPostor(vProSelNro As Integer, vPersCod As String, vProSelItem As Integer)
Dim oConn As New DConecta, Rs As New ADODB.Recordset, i As Integer

FormaFlexItem

sSQL = "select i.nProSelItem, i.cBSCod, b.cBSDescripcion, i.nCantidad, t.cUnidad,p.cPersCod " & _
"  from LogProSelItemBS i inner join LogProSelBienesServicios b on i.cBSCod = b.cProSelBSCod " & _
" inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) t on b.nBSUnidad = t.nBSUnidad " & _
" inner join (select distinct nProSelNro,nProSelItem,cPersCod from LogProSelPostorPropuesta where bGanador=1) p on p.nProSelNro = i.nProSelNro and p.nProSelItem = i.nProSelItem " & _
" where i.nProSelNro = " & vProSelNro & " and p.cPersCod = '" & vPersCod & "' and i.nProSelItem=" & vProSelItem

If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
Else
   MsgBox "No se puede establecer conexión..." + Space(10), vbInformation
   Exit Sub
End If

If Not Rs.EOF Then
   i = 0
   Do While Not Rs.EOF
      i = i + 1
      InsRow MSItem, i
      MSItem.TextMatrix(i, 0) = Rs!nProSelItem
      MSItem.TextMatrix(i, 1) = Rs!cBSCod
      MSItem.TextMatrix(i, 2) = Rs!cBSDescripcion
      MSItem.TextMatrix(i, 3) = Rs!nCantidad
      MSItem.TextMatrix(i, 4) = Rs!cUnidad
      Rs.MoveNext
   Loop
End If
End Sub

Sub FormaFlexItem()
MSItem.Clear
MSItem.Rows = 2
MSItem.RowHeight(0) = 360
MSItem.RowHeight(1) = 8
MSItem.ColWidth(0) = 0
MSItem.ColWidth(1) = 850:   MSItem.TextMatrix(0, 1) = "Código":       MSItem.ColAlignment(1) = 4
MSItem.ColWidth(2) = 4100:  MSItem.TextMatrix(0, 2) = "Descripción"
MSItem.ColWidth(3) = 950:   MSItem.TextMatrix(0, 3) = "  Cantidad":   MSItem.ColAlignment(3) = 4
MSItem.ColWidth(4) = 1600:  MSItem.TextMatrix(0, 4) = "   Unidad":    MSItem.ColAlignment(4) = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLogProSelContratoDet = Nothing
End Sub

Private Sub MSItem_GotFocus()
GeneraDetalleItem nProselNro, MSItem.TextMatrix(MSItem.row, 0)
End Sub

Private Sub MSItem_RowColChange()
GeneraDetalleItem nProselNro, MSItem.TextMatrix(MSItem.row, 0)
End Sub

Sub GeneraDetalleItem(vProSelNro As Integer, vProSelItem As Integer)
Dim oConn As New DConecta, Rs As New ADODB.Recordset, i As Integer, nSuma As Integer

FormaFlexDetalle

sSQL = "select e.nProSelItem, e.cAgeCod, a.cAgeDescripcion, e.dFechaEntrega, e.nCantidad, e.cBSCod, b.cBSDescripcion " & _
       "  from LogProSelContratoEntrega e " & _
       "    inner join Agencias a on e.cAgeCod = a.cAgeCod " & _
       "    inner join LogProSelBienesServicios b on e.cBSCod = b.cProSelBSCod " & _
       " Where e.bEstado = 1 And e.nProSelNro = " & vProSelNro & " And e.nProSelItem = " & vProSelItem & " " & _
       " "

If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
Else
   MsgBox "No se puede establecer conexión..." + Space(10), vbInformation
   Exit Sub
End If

If Not Rs.EOF Then
   i = 0
   Do While Not Rs.EOF
      i = i + 1
      InsRow MSDet, i
      MSDet.TextMatrix(i, 1) = Rs!cAgeCod
      MSDet.TextMatrix(i, 2) = Rs!cAgeDescripcion
      MSDet.TextMatrix(i, 3) = Rs!cBSCod
      MSDet.TextMatrix(i, 4) = Rs!cBSDescripcion
      MSDet.TextMatrix(i, 5) = Rs!nCantidad
      MSDet.TextMatrix(i, 6) = Rs!dFechaEntrega
      If Rs!cBSCod = MSItem.TextMatrix(MSItem.row, 1) Then nSuma = nSuma + Rs!nCantidad
      Rs.MoveNext
   Loop
   txtCantidadFaltante.Text = MSItem.TextMatrix(MSItem.row, 3) - nSuma
End If
End Sub

Sub FormaFlexDetalle()
MSDet.Clear
MSDet.Rows = 2
MSDet.Cols = 7
MSDet.RowHeight(0) = 360
MSDet.RowHeight(1) = 8
MSDet.ColWidth(0) = 0
MSDet.ColWidth(1) = 400:   MSDet.TextMatrix(0, 1) = "C.A.":      MSDet.ColAlignment(1) = 4
MSDet.ColWidth(2) = 3000:  MSDet.TextMatrix(0, 2) = "Agencia de entrega"
MSDet.ColWidth(3) = 1000:   MSDet.TextMatrix(0, 3) = "  Código":  MSDet.ColAlignment(3) = 4
MSDet.ColWidth(4) = 3100:  MSDet.TextMatrix(0, 4) = "Descripcion":   MSDet.ColAlignment(4) = 4
MSDet.ColWidth(5) = 800:   MSDet.TextMatrix(0, 5) = "  Cantidad":  MSDet.ColAlignment(5) = 4
MSDet.ColWidth(6) = 1000:  MSDet.TextMatrix(0, 6) = "Fecha Entrega":   MSDet.ColAlignment(6) = 4
End Sub

Function GetContratoNro(ByVal vPersCod As String, ByVal vProSelNro As Integer, ByVal pnProSelItem As Integer) As Integer
Dim cSQL As String, oConn As New DConecta, Rs As New ADODB.Recordset
GetContratoNro = 0
cSQL = "select isnull(max(nProSelConNro),0) from LogProSelContrato where cPersCod = '" & vPersCod & "' and nProSelNro = " & vProSelNro & " and nProSelItem= " & pnProSelItem
If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(cSQL)
   If Not Rs.EOF Then
      GetContratoNro = Rs(0)
   End If
End If
End Function

Private Sub cboAgencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtCantidad.SetFocus
End If
End Sub

Private Sub txtCantidad_GotFocus()
SelTexto txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   txtFechaEnt.SetFocus
End If
End Sub

Private Sub txtFecha_GotFocus()
SelTexto txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigFecha(txtFecha, KeyAscii)
If nKeyAscii = 13 Then
End If
End Sub

Private Sub txtFechaEnt_GotFocus()
SelTexto txtFechaEnt
End Sub

Private Sub txtFechaEnt_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigFecha(txtFechaEnt, KeyAscii)
If nKeyAscii = 13 Then
   cmdEntrega.SetFocus
End If
End Sub


