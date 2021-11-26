VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogActivoFijo 
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   2565
   ClientWidth     =   11790
   LinkTopic       =   "Form2"
   ScaleHeight     =   5475
   ScaleWidth      =   11790
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10320
      TabIndex        =   7
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton cmdCodBar 
      Caption         =   "Código de Barras"
      Height          =   375
      Left            =   180
      TabIndex        =   6
      Top             =   5040
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   4035
      Left            =   180
      TabIndex        =   2
      Top             =   900
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   7117
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorSel    =   -2147483647
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   180
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin VB.ComboBox cboAge 
         Height          =   315
         Left            =   7560
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtAnio 
         Height          =   315
         Left            =   780
         MaxLength       =   4
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   2580
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3735
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
         Left            =   6720
         TabIndex        =   9
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         TabIndex        =   4
         Top             =   420
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " Grupo"
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
         Left            =   1860
         TabIndex        =   3
         Top             =   420
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmLogActivoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String

Private Sub cboAge_Click()
cboTipo_Click
End Sub

Private Sub cmdCodBar_Click()
Dim k As Integer
k = MSFlex.row
frmLogAlmCodBarra.Codigo MSFlex.TextMatrix(k, 1), MSFlex.TextMatrix(k, 2), MSFlex.TextMatrix(k, 3)
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
txtAnio.Text = Year(Date)
LimpiaFlex
CargaDatos
End Sub

Sub CargaDatos()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim oGen As DGeneral

Set oGen = New DGeneral
    
Set rs = oGen.GetConstante(5062, False)
Me.cboTipo.Clear
While Not rs.EOF
      cboTipo.AddItem rs.Fields(0)
      cboTipo.ItemData(cboTipo.ListCount - 1) = rs.Fields(1)
      rs.MoveNext
Wend


sSQL = "Select cAgeCod,cAgeDescripcion from Agencias where nEstado=1"
Me.cboAge.Clear
cboAge.AddItem "-- Todas las Agencias"
cboAge.ItemData(cboAge.ListCount - 1) = 0
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   While Not rs.EOF
      cboAge.AddItem rs.Fields(1)
      cboAge.ItemData(cboAge.ListCount - 1) = rs.Fields(0)
      rs.MoveNext
   Wend
   cboAge.ListIndex = 0
End If

End Sub

Sub LimpiaFlex()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(1) = 280
MSFlex.ColWidth(0) = 0:      MSFlex.TextMatrix(0, 0) = ""
MSFlex.ColWidth(1) = 1000:   MSFlex.TextMatrix(0, 1) = ""
MSFlex.ColWidth(2) = 3000:   MSFlex.TextMatrix(0, 2) = ""
MSFlex.ColWidth(3) = 3000:   MSFlex.TextMatrix(0, 3) = "": MSFlex.ColAlignment(3) = 1
MSFlex.ColWidth(4) = 2000:   MSFlex.TextMatrix(0, 4) = ""
MSFlex.ColWidth(5) = 2000:   MSFlex.TextMatrix(0, 5) = ""
MSFlex.ColWidth(6) = 0:      MSFlex.TextMatrix(0, 6) = ""
End Sub

Private Sub cboTipo_Click()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta, i As Integer
Dim nAnio As Integer
Dim nTipo As Integer
Dim nAgeCod As Integer
Dim cConsulta As String

i = 0
LimpiaFlex
nAnio = CInt(txtAnio)
If cboTipo.ListIndex < 0 Then Exit Sub
nTipo = cboTipo.ItemData(cboTipo.ListIndex)
nAgeCod = cboAge.ItemData(cboAge.ListIndex)
If nAgeCod > 0 Then
   cConsulta = " and b.cAgeCod = '" & Format(nAgeCod, "00") & "'"
Else
   cConsulta = ""
End If

sSQL = "select b.cBSCod,cBSDescripcion=coalesce(x.cBSDescripcion,''), b.cSerie,cArea=coalesce(a.cAreaDescripcion,''), cAgencia=coalesce(g.cAgeDescripcion,'') " & _
       "  from BSActivoFijo b left join BienesServicios x on b.cBSCod = x.cBSCod " & _
       "       left join Areas a on b.cAreCod = a.cAreaCod " & _
       "       left join Agencias g on b.cAgeCod = g.cAgeCod " & _
       " Where b.nAnio = " & nAnio & " And b.ban = " & nTipo & " And b.bBaja = 0 " + cConsulta
       
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 1) = rs!cBSCod
         MSFlex.TextMatrix(i, 2) = rs!cBSDescripcion
         MSFlex.TextMatrix(i, 3) = rs!cSerie
         MSFlex.TextMatrix(i, 4) = rs!cArea
         MSFlex.TextMatrix(i, 5) = rs!cAgencia
         MSFlex.TextMatrix(i, 6) = ""
         rs.MoveNext
      Loop
   End If
End If
End Sub


