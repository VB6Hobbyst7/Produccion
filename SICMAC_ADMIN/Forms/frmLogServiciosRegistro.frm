VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogServiciosRegistro 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4035
   ClientLeft      =   1845
   ClientTop       =   2625
   ClientWidth     =   6690
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6690
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6435
      Begin VB.ComboBox cboServicio 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   5235
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
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
         TabIndex        =   2
         Top             =   420
         Width           =   705
      End
   End
   Begin VB.Frame fraReg 
      BorderStyle     =   0  'None
      Height          =   3075
      Left            =   120
      TabIndex        =   7
      Top             =   900
      Visible         =   0   'False
      Width           =   6435
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         Top             =   2640
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5160
         TabIndex        =   11
         Top             =   2640
         Width           =   1275
      End
      Begin VB.Frame fraAgencia 
         Height          =   855
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   6435
         Begin VB.ComboBox cboAgencia 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   5235
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Agencia"
            Height          =   195
            Left            =   180
            TabIndex        =   10
            Top             =   420
            Width           =   585
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Width           =   6435
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            ItemData        =   "frmLogServiciosRegistro.frx":0000
            Left            =   4080
            List            =   "frmLogServiciosRegistro.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   780
            Width           =   855
         End
         Begin VB.TextBox txtNroRecibo 
            Height          =   315
            Left            =   960
            MaxLength       =   20
            TabIndex        =   19
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtFechaPago 
            Height          =   315
            Left            =   4920
            MaxLength       =   10
            TabIndex        =   18
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtAnio 
            Height          =   315
            Left            =   960
            MaxLength       =   4
            TabIndex        =   17
            Top             =   780
            Width           =   555
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   780
            Width           =   1875
         End
         Begin VB.TextBox txtMonto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   15
            Top             =   780
            Width           =   1275
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Recibo"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   420
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Ultima fecha pago"
            Height          =   195
            Left            =   3540
            TabIndex        =   24
            Top             =   420
            Width           =   1290
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mes"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   840
            Width           =   300
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            Height          =   195
            Left            =   3540
            TabIndex        =   22
            Top             =   840
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   1260
            Width           =   495
         End
      End
   End
   Begin VB.Frame fraLis 
      BorderStyle     =   0  'None
      Height          =   3075
      Left            =   120
      TabIndex        =   3
      Top             =   900
      Width           =   6435
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   2700
         Width           =   1275
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   5160
         TabIndex        =   4
         Top             =   2700
         Width           =   1275
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   2655
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   4683
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   7
         FixedCols       =   0
         BackColorSel    =   16313054
         ForeColorSel    =   128
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
   End
End
Attribute VB_Name = "frmLogServiciosRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cOpeCod As String, nDocTpo As Integer
Dim sSQL As String, nKeyAscii As Integer

Public Sub Inicio(ByVal psOpeCod As String)
cOpeCod = psOpeCod
Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
fraReg.Visible = False
fraLis.Visible = True
End Sub

Private Sub cmdnuevo_Click()
fraLis.Visible = False
fraReg.Visible = True
txtFechaPago.Text = Date
txtNroRecibo.Text = ""
txtMonto.Text = ""
txtAnio.Text = Year(Date)
cboMoneda.ListIndex = 0
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
nDocTpo = 14
cOpeCod = ""
CargaListas
ListaServicios
End Sub

Private Sub cboServicio_Click()
If cboServicio.ListIndex > 0 Then
   cOpeCod = CStr(cboServicio.ItemData(cboServicio.ListIndex))
   ListaServicios
Else
   FormaFlex
End If
End Sub

Sub ListaServicios()
Dim oConn As New DConecta
Dim Rs As New ADODB.Recordset
Dim i As Integer

'inner join OpeTpo o on s.cOpeCod = o.cOpeCod
FormaFlex
sSQL = "select s.nMovServ, cDocNumero=substring(d.cDocDesc,1,3)+'-'+cDocNro, s.nMonto, " & _
       "       s.dFechaPago, cAgencia=coalesce(a.cAgeDescripcion,''),s.nEstado " & _
       "  from LogServiciosPublicos s inner join Documento d on s.nDocTpo = d.nDocTpo " & _
       "       left join Agencias a on s.cAgeCod = a.cAgeCod   " & _
       "  " & _
       " WHERE s.cOpeCod = '" & cOpeCod & "'"

If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 1) = Rs!nMovServ
         MSFlex.TextMatrix(i, 2) = Rs!cDocNumero
         MSFlex.TextMatrix(i, 3) = Rs!dFechaPago
         MSFlex.TextMatrix(i, 4) = FNumero(Rs!nMonto)
         MSFlex.TextMatrix(i, 5) = Rs!cAgencia
         MSFlex.TextMatrix(i, 6) = Rs!nEstado
         Rs.MoveNext
      Loop
   End If
End If
End Sub

Sub FormaFlex()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 0
MSFlex.ColWidth(2) = 2000
MSFlex.ColWidth(3) = 1000
MSFlex.ColWidth(4) = 1000
MSFlex.ColWidth(5) = 2200
MSFlex.ColWidth(6) = 0
End Sub

Sub CargaListas()
Dim Rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim k As Integer

If oConn.AbreConexion Then

   'sSQL = "Select cOpeDesc from OpeTpo where cOpeCod = '" & cOpeCod & "'"
   'Set rs = oConn.CargaRecordSet(sSQL)
   'If Not rs.EOF Then
   '   Me.Caption = rs!cOpeDesc
   'Else
   '   Me.Caption = "Operación no indicada"
   '   cmdGrabar.Visible = False
   'End If
   
   cboAgencia.Clear
   sSQL = "Select cAgeCod,cAgeDescripcion from Agencias where nEstado = 1"
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         cboAgencia.AddItem Rs!cAgeDescripcion
         cboAgencia.ItemData(cboAgencia.ListCount - 1) = Rs!cAgeCod
         Rs.MoveNext
      Loop
      cboAgencia.ListIndex = 0
   End If
   
   cboServicio.Clear
   cboServicio.AddItem "--- Selecione un Tipo de Servicio"
   'sSQL = "select nConsValor,cConsDescripcion from Constante where nConsCod<>nConsValor and nConsCod=9134"
   sSQL = "select cOpeCod,cOpeDesc from OpeTpo where cOpeCod like '54102%' and nOpeNiv=3"
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         'cboServicio.AddItem rs!cConsDescripcion
         'cboServicio.ItemData(cboServicio.ListCount - 1) = rs!nConsValor
         cboServicio.AddItem UCase(Rs!cOpeDesc)
         cboServicio.ItemData(cboServicio.ListCount - 1) = Rs!cOpeCod
         Rs.MoveNext
      Loop
      cboServicio.ListIndex = 0
   End If
   
   cboMes.AddItem "--- Seleccione Mes"
   cboMes.AddItem "ENERO":       cboMes.AddItem "FEBRERO"
   cboMes.AddItem "MARZO":       cboMes.AddItem "ABRIL"
   cboMes.AddItem "MAYO":        cboMes.AddItem "JUNIO"
   cboMes.AddItem "JULIO":       cboMes.AddItem "AGOSTO"
   cboMes.AddItem "SEPTIEMBRE":  cboMes.AddItem "OCTUBRE"
   cboMes.AddItem "NOVIEMBRE":   cboMes.AddItem "DICIEMBRE"
   cboMes.ListIndex = 0
End If
End Sub

'--------------------------------------------------------------------------
'------   EVENTOS DE TECLAS
'--------------------------------------------------------------------------

Private Sub cmdGrabar_Click()
Dim i As Integer
Dim nMes As Integer, nAnio As Integer
Dim nMonto As Currency, nMoneda As Integer
Dim cAgeCod As String
Dim oConn As New DConecta
cAgeCod = Format(cboAgencia.ItemData(cboAgencia.ListIndex), "00")

If cboServicio.ListIndex <= 0 Then Exit Sub
If cboMes.ListIndex <= 0 Then
   MsgBox "Debe seleccionar un Mes válido..." + Space(10), vbInformation, "Aviso"
   Exit Sub
Else
   nMes = cboMes.ListIndex
End If

If Len(Trim(txtNroRecibo.Text)) = 0 Then
   MsgBox "Debe ingresar un número de recibo válido..." + Space(10), vbInformation, "Aviso"
   Exit Sub
End If

nMoneda = cboMoneda.ItemData(cboMoneda.ListIndex)
nMonto = VNumero(txtMonto)

If MsgBox("¿Está seguro de grabar los datos ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then

   sSQL = "INSERT INTO LogServiciosPublicos (cOpeCod,nDocTpo,cDocNro,dFechaPago,nMes,nAnio,cAgeCod,nMonto,cCodUsu,dFecha) " & _
          "  VALUES ('" & cOpeCod & "'," & nDocTpo & ",'" & txtNroRecibo.Text & "','" & Format(txtFechaPago.Text, "YYYYMMDD") & "'," & nMes & "," & CInt(txtAnio.Text) & ",'" & cAgeCod & "'," & nMonto & ",'" & gsCodUser & "','" & Format(Date, "YYYYMMDD") & "') "
          
   If oConn.AbreConexion Then
      oConn.Ejecutar sSQL
   End If
   ListaServicios
   fraReg.Visible = False
   fraLis.Visible = True

End If
End Sub


'--------------------------------------------------------------------------
'------   EVENTOS DE TECLAS
'--------------------------------------------------------------------------

Private Sub cboServicio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtNroRecibo.SetFocus
End If
End Sub


Private Sub txtNroRecibo_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii, "/-:")
If nKeyAscii = 13 Then
   txtFechaPago.SetFocus
End If
End Sub

Private Sub txtFechaPago_GotFocus()
SelTexto txtFechaPago
End Sub

Private Sub txtFechaPago_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigFecha(txtFechaPago, KeyAscii)
If nKeyAscii = 13 Then
   txtAnio.SetFocus
End If
End Sub

Private Sub cboAgencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtNroRecibo.SetFocus
End If
End Sub

Private Sub txtAnio_GotFocus()
SelTexto txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   cboMes.SetFocus
End If
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboMoneda.SetFocus
End If
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtMonto.SetFocus
End If
End Sub

Private Sub txtMonto_GotFocus()
SelTexto txtMonto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumDec(txtMonto, KeyAscii)
If nKeyAscii = 13 Then
   txtMonto = FNumero(txtMonto.Text)
   cmdGrabar.SetFocus
End If
End Sub




