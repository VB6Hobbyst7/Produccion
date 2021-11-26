VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogVehiculoMov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Vehicular - Movimientos vehiculares"
   ClientHeight    =   4650
   ClientLeft      =   1065
   ClientTop       =   2355
   ClientWidth     =   9690
   Icon            =   "frmLogVehiculoMov.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   9690
   Begin VB.CommandButton cmdMovs 
      Caption         =   "Ver Movimientos"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   1515
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5741
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
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.ComboBox cboMovil 
         Height          =   315
         ItemData        =   "frmLogVehiculoMov.frx":08CA
         Left            =   1800
         List            =   "frmLogVehiculoMov.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   2475
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Vehículo"
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
         TabIndex        =   4
         Top             =   360
         Width           =   1485
      End
   End
   Begin VB.Menu mnuVehiculo 
      Caption         =   "Vehiculos"
      Visible         =   0   'False
      Begin VB.Menu mnuVerMovs 
         Caption         =   "Movimientos"
      End
   End
End
Attribute VB_Name = "frmLogVehiculoMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMovs_Click()
Dim i As Integer
i = MSFlex.row
If Len(Trim(MSFlex.TextMatrix(i, 1))) > 0 Then
   frmLogVehiculoMovDet.Vehiculo MSFlex.TextMatrix(i, 1), MSFlex.TextMatrix(i, 2) + " " + MSFlex.TextMatrix(i, 3) + " " + MSFlex.TextMatrix(i, 4)
End If
End Sub

Private Sub Form_Load()
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset

CentraForm Me
LimpiaFlex

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet("select nConsValor,cConsDescripcion from Constante where nConsCod =9021 and nconscod<>nconsvalor order by nConsValor")
   oConn.CierraConexion
   If Not rs.EOF Then
      Do While Not rs.EOF
         cboMovil.AddItem rs!cConsDescripcion
         cboMovil.ItemData(cboMovil.ListCount - 1) = rs!nConsValor
         rs.MoveNext
      Loop
      cboMovil.ListIndex = 0
   Else
      MsgBox "Faltan definir tipos de vehiculos en la tabla Constante..." + Space(10), vbInformation
   End If
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cboMovil_Click()
Dim nTipoV As Integer
nTipoV = cboMovil.ItemData(cboMovil.ListIndex)
GeneraVehiculos nTipoV
End Sub

Sub GeneraVehiculos(pnTipo As Integer)
Dim v As DLogVehiculos, i As Integer
Dim rs As New ADODB.Recordset
i = 0
LimpiaFlex
Set v = New DLogVehiculos

Set rs = v.ListaVehiculos(pnTipo)
If Not rs.EOF Then
   Do While Not rs.EOF
      i = i + 1
      InsRow MSFlex, i
      MSFlex.TextMatrix(i, 0) = "" 'rs!cMovNro
      MSFlex.TextMatrix(i, 1) = rs!nVehiculoCod
      MSFlex.TextMatrix(i, 2) = rs!cTipoVehiculo
      MSFlex.TextMatrix(i, 3) = rs!cMarca
      MSFlex.TextMatrix(i, 4) = rs!cPlaca
      MSFlex.TextMatrix(i, 5) = rs!cEstado
      MSFlex.TextMatrix(i, 6) = rs!nEstado
      If rs!nEstado = 1 Then
         MSFlex.Col = 5
         MSFlex.row = i
         MSFlex.CellForeColor = "&H00C00000"
      End If
      If rs!nEstado = 2 Then
         MSFlex.Col = 5
         MSFlex.row = i
         MSFlex.CellForeColor = "&H00000080"
      End If
      rs.MoveNext
   Loop
   MSFlex.Col = 1
   MSFlex.row = 1
End If
End Sub

Sub LimpiaFlex()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 300
MSFlex.ColWidth(2) = 4500:  MSFlex.TextMatrix(0, 2) = "Vehículo"
MSFlex.ColWidth(3) = 1200:  MSFlex.TextMatrix(0, 3) = "Marca"
MSFlex.ColWidth(4) = 1000:  MSFlex.TextMatrix(0, 4) = "Placa": MSFlex.ColAlignment(4) = 4
MSFlex.ColWidth(5) = 2200:  MSFlex.TextMatrix(0, 5) = "Estado"
MSFlex.ColWidth(6) = 0
End Sub


Private Sub MSFlex_KeyPress(KeyAscii As Integer)
'Dim nEst As Integer, nCod As Integer, cMov As String
'If KeyAscii = 13 Then
'   cMov = MSFlex.TextMatrix(MSFlex.Row, 0)
'   nCod = CInt(VNumero(MSFlex.TextMatrix(MSFlex.Row, 1)))
'   nEst = CInt(VNumero(MSFlex.TextMatrix(MSFlex.Row, 6)))
'   If nCod > 0 And nEst >= 0 Then
'      frmLogVehiculoAsigna.Operacion nCod, nEst, cMov
'      If frmLogVehiculoAsigna.vpGrabado Then
'         GeneraVehiculos cboMovil.ListIndex
'      End If
'   End If
'End If
End Sub


