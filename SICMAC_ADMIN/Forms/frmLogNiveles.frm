VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogNiveles 
   ClientHeight    =   5550
   ClientLeft      =   1830
   ClientTop       =   2250
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8145
   Begin VB.Frame fraVis1 
      Caption         =   "Cargos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2595
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   7995
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex1 
         Height          =   2175
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   3836
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   6
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
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
   Begin VB.Frame fraVis2 
      Caption         =   "Ruta de Aprobación "
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
      Height          =   2655
      Left            =   60
      TabIndex        =   18
      Top             =   2760
      Width           =   7995
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   6720
         TabIndex        =   22
         Top             =   2220
         Width           =   1155
      End
      Begin VB.CommandButton cmdAgregar2 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2220
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuitar2 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   1260
         TabIndex        =   19
         Top             =   2220
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex2 
         Height          =   1935
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   3413
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   7
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
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
         _Band(0).Cols   =   7
      End
   End
   Begin VB.Frame fraReg2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAFFFF&
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   7875
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   900
         TabIndex        =   9
         Top             =   660
         Width           =   3975
         Begin VB.CheckBox chkProceso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00EAFFFF&
            Caption         =   "Proceso"
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
            Left            =   2640
            TabIndex        =   23
            Top             =   300
            Width           =   1035
         End
         Begin VB.TextBox txtNivel 
            Height          =   315
            Left            =   1980
            TabIndex        =   10
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel de Aprobación"
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
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Width           =   1740
         End
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   5220
         TabIndex        =   8
         Top             =   1980
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancela 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6480
         TabIndex        =   7
         Top             =   1980
         Width           =   1155
      End
      Begin VB.TextBox txtCargoDesc 
         Height          =   315
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   5715
      End
      Begin VB.CommandButton cmdCargo 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1650
         TabIndex        =   5
         Top             =   390
         Width           =   315
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         Caption         =   "Sector "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   900
         TabIndex        =   1
         Top             =   1380
         Width           =   2475
         Begin VB.OptionButton opSector13 
            BackColor       =   &H00EAFFFF&
            Caption         =   "Ahorros/Finanzas"
            Height          =   195
            Left            =   300
            TabIndex        =   4
            Top             =   900
            Width           =   1695
         End
         Begin VB.OptionButton opSector12 
            BackColor       =   &H00EAFFFF&
            Caption         =   "Créditos"
            Height          =   195
            Left            =   300
            TabIndex        =   3
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton opSector11 
            BackColor       =   &H00EAFFFF&
            Caption         =   "Administración"
            Height          =   195
            Left            =   300
            TabIndex        =   2
            Top             =   300
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox txtCargo 
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Frame fraPro 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   5040
         TabIndex        =   12
         Top             =   660
         Width           =   2655
         Begin VB.CheckBox chkAge 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00EAFFFF&
            Caption         =   "Consolida toda su Agencia"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   300
            Width           =   2295
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   420
         Width           =   510
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "MenuUtiles"
      Visible         =   0   'False
      Begin VB.Menu mnuCopia 
         Caption         =   "Copiar Ruta"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPegar 
         Caption         =   "Pegar"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRenu 
         Caption         =   "Re-numerar"
      End
   End
End
Attribute VB_Name = "frmLogNiveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim nFuncion As Integer
Dim cRHCargoCopia As String

Private Sub cmdAgregar2_Click()
nFuncion = 1
fraVis2.Visible = False
fraReg2.Visible = True
txtCargo.Text = ""
txtCargoDesc.Text = ""
txtNivel.Text = ""
End Sub

Private Sub cmdCancela_Click()
nFuncion = 0
fraVis2.Visible = True
fraReg2.Visible = False
End Sub

Private Sub cmdCargo_Click()
sSql = "select cRHCargoCod as Codigo,cRHCargoDescripcion as Descripcion from RHCargosTabla where bRHCargoEstado=1 order by cRHCargoDescripcion "
'frmSelectorArbol.FormaArbolConsulta sSql, "Selección de Cargos"
'If frmSelectorArbol.vpSeleccion Then
'   txtCargo.Text = frmSelectorArbol.vpCodigo
'   txtCargoDesc.Text = frmSelectorArbol.vpDescripcion
'End If

frmLogSelector.Consulta sSql, "Selección de Cargos"
If frmLogSelector.vpHaySeleccion Then
   txtCargo.Text = frmLogSelector.vpCodigo
   txtCargoDesc.Text = frmLogSelector.vpDescripcion
End If

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Me.Caption = "Plan Anual de Adquisiciones y Contrataciones - Niveles de Aprobación"
cRHCargoCopia = ""
CargaCargos
End Sub

Sub CargaCargos()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim i As Integer

MSFlex1.Clear
MSFlex1.Rows = 2
MSFlex1.RowHeight(1) = 8
MSFlex1.ColWidth(0) = 0
MSFlex1.ColWidth(1) = 600:     MSFlex1.ColAlignment(1) = 1
MSFlex1.ColWidth(2) = 6000
MSFlex1.ColWidth(3) = 0:      MSFlex1.ColAlignment(3) = 4: MSFlex1.TextMatrix(0, 3) = "Nivel Proc"
MSFlex1.ColWidth(4) = 0:      MSFlex1.ColAlignment(4) = 4: MSFlex1.TextMatrix(0, 4) = "x Agencia"

sSql = "select cRHCargoCod, cRHCargoDescripcion  " & _
       " from RHCargosTabla where bRHCargoEstado = 1 and len(cRHCargoCod)=6 order by cRHCargoDescripcion "
       
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSql)
   If Not rs.EOF Then
      i = 0
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex1, i
         MSFlex1.TextMatrix(i, 1) = rs!cRHCargoCod
         MSFlex1.TextMatrix(i, 2) = rs!cRHCargoDescripcion
         MSFlex1.TextMatrix(i, 3) = ""
         MSFlex1.TextMatrix(i, 4) = ""
         MSFlex1.TextMatrix(i, 5) = ""
         rs.MoveNext
      Loop
   End If
End If
End Sub

Sub ListaRutaAprobacion(vRHCargoApro As String)
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim i As Integer

MSFlex2.Clear
MSFlex2.Rows = 2
MSFlex2.RowHeight(1) = 8
MSFlex2.ColWidth(0) = 0
MSFlex2.ColWidth(1) = 550:    MSFlex2.ColAlignment(1) = 4
MSFlex2.ColWidth(2) = 300:    MSFlex2.ColAlignment(2) = 4
MSFlex2.ColWidth(3) = 0:      MSFlex2.ColAlignment(3) = 4
MSFlex2.ColWidth(4) = 3600
MSFlex2.ColWidth(5) = 2000
MSFlex2.ColWidth(6) = 1000:   MSFlex2.TextMatrix(0, 6) = "Consol.Age": MSFlex2.ColAlignment(3) = 4
       
'sSQL = "select n.cRHCargoCodAprobacion,t.cRHCargoDescripcion,n.nNivelAprobacion,n.nNivelProceso, cSector " & _
'       "  from LogNivelAprobacion n " & _
'       "       inner join RHCargosTabla t on n.cRHCargoCodAprobacion = t.cRHCargoCod " & _
'       "       inner join (SELECT nConsValor as nSector, cConsDescripcion as cSector FROM Constante WHERE (nConsCod = 9045) AND (nConsCod <> nConsValor)) s on n.nSector = s.nSector " & _
'       " where n.cRHCargoCod = '" & vRHCargoApro & "' order by n.nNivelAprobacion,n.cRHCargoCodAprobacion "
 
 
sSql = "select n.cRHCargoCodAprobacion, t.cRHCargoDescripcion, n.nNivelAprobacion, s.cSector, n.nAgencia" & _
       "  from LogNivelAprobacion n " & _
       "       inner join RHCargosTabla t on n.cRHCargoCodAprobacion = t.cRHCargoCod " & _
       "       inner join (SELECT nConsValor as nSector, cConsDescripcion as cSector FROM Constante WHERE (nConsCod = 9045) AND (nConsCod <> nConsValor)) s on n.nSector = s.nSector " & _
       " where n.cRHCargoCod = '" & vRHCargoApro & "' order by n.nNivelAprobacion,n.cRHCargoCodAprobacion "
 
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSql)
   If Not rs.EOF Then
      i = 0
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex2, i
         MSFlex2.TextMatrix(i, 1) = rs!cRHCargoCodAprobacion
         MSFlex2.TextMatrix(i, 2) = rs!nNivelAprobacion
         MSFlex2.TextMatrix(i, 3) = ""
         MSFlex2.TextMatrix(i, 4) = rs!cRHCargoDescripcion
         MSFlex2.TextMatrix(i, 5) = rs!cSector
         MSFlex2.TextMatrix(i, 6) = IIf(rs!nAgencia = 1, "SI", "")
         rs.MoveNext
      Loop
   End If
End If
End Sub

Private Sub cmdQuitar2_Click()
Dim oConn As New DConecta
Dim cRHCargoCod As String
Dim cRHCargoCodApro As String
Dim nNivel As Integer

cRHCargoCod = MSFlex1.TextMatrix(MSFlex1.row, 1)

If Len(MSFlex2.TextMatrix(MSFlex2.row, 1)) = 0 Then
   MsgBox "Debe seleccionar un cargo válido..." + Space(10), vbInformation
   Exit Sub
End If

If MsgBox("¿ Está seguro de agregar el Cargo ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then

   cRHCargoCodApro = MSFlex2.TextMatrix(MSFlex2.row, 1)
   nNivel = MSFlex2.TextMatrix(MSFlex2.row, 2)
   
   sSql = "DELETE FROM LogNivelAprobacion WHERE cRHCargoCod = '" & cRHCargoCod & "' and cRHCargoCodAprobacion = '" & cRHCargoCodApro & "' and nNivelAprobacion = " & nNivel & ""
   If oConn.AbreConexion Then
      oConn.Ejecutar sSql
      ListaRutaAprobacion cRHCargoCod
   End If
End If
End Sub

Private Sub cmdGrabar_Click()
Dim oConn As New DConecta, rs As New ADODB.Recordset
Dim nNivel As Integer, nNivelPro As Integer, nConsol As Integer, nSector As Integer
Dim cRHCargoCod As String, cRHCargoCodApro As String
Dim nMaxNivel As Integer
Dim nNro As Integer

If opSector11.value = False And opSector12.value = False And opSector13.value = False Then
   MsgBox "Debe seleccionar un sector..." + Space(10), vbInformation
   Exit Sub
End If
   
If opSector11.value Then nSector = 1
If opSector12.value Then nSector = 2
If opSector13.value Then nSector = 3

cRHCargoCod = MSFlex1.TextMatrix(MSFlex1.row, 1)

If MsgBox("¿ Está seguro de agregar el Cargo ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   If oConn.AbreConexion Then
      
      cRHCargoCodApro = txtCargo.Text
      nNivel = CInt(VNumero(txtNivel.Text))
      nNro = 0
      nMaxNivel = 0
      nNivelPro = 0
      If chkProceso.value = 1 Then
         
         sSql = "select nNro=count(*) from LogNivelProceso where cRHCargoCod = '" & cRHCargoCodApro & "'"
         Set rs = oConn.Ejecutar(sSql)
         If Not rs.EOF Then
            nNro = rs!nNro
         End If
         If nNro > 1 Then
            sSql = "select nMaxNivel = Max(nNivelProceso) from LogNivelAprobacion where cRHCargoCod = '" & cRHCargoCod & "'"
            Set rs = oConn.Ejecutar(sSql)
            If Not rs.EOF Then
               nMaxNivel = rs!nMaxNivel
            End If
         End If
         
         sSql = "select nNivelProceso from LogNivelProceso where cRHCargoCod = '" & cRHCargoCodApro & "' and nNivelProceso > " & nMaxNivel & ""
         Set rs = oConn.Ejecutar(sSql)
         If Not rs.EOF Then
            nNivelPro = rs!nNivelProceso
         End If
         
      End If
   
      If nFuncion = 1 Then

         'nNivelPro = CInt(VNumero(txtNivelPro.Text))
         nConsol = chkAge.value
         
         'sSQL = "INSERT INTO LogNivelAprobacion (cRHCargoCod,cRHCargoCodAprobacion,nNivelAprobacion,nNivelProceso, nAgencia, nSector) " & _
         '       " VALUES ('" & cRHCargoCod & "','" & cRHCargoCodApro & "'," & nNivel & "," & nNivelPro & "," & nConsol & "," & nSector & ") "
         
         sSql = "INSERT INTO LogNivelAprobacion (cRHCargoCod, cRHCargoCodAprobacion, nNivelAprobacion,  nAgencia, nSector) " & _
                " VALUES ('" & cRHCargoCod & "','" & cRHCargoCodApro & "'," & nNivel & "," & nConsol & "," & nSector & ") "
         oConn.Ejecutar sSql
         
      End If
      
      'If nFuncion = 2 Then
      '   cRHCargoCod = txtCargo.Text
      '   nNivel = CInt(VNumero(txtNivel.Text))
      '   nConsol = chkAge.Value
      '
      '   sSQL = "UPDATE LogPlanAnualCargosApro SET nNivelProceso = " & nNivel & ", nConsolidaAgencia = " & nConsol & ", nSector = " & nSector & " " & _
      '          " WHERE cRHCargoCodAprobacion = '" & cRHCargoCod & "' "
      '   oConn.Ejecutar sSQL
      'End If
      
   End If
   fraVis2.Visible = True
   fraReg2.Visible = False
   ListaRutaAprobacion MSFlex1.TextMatrix(MSFlex1.row, 1)
   MSFlex1.SetFocus
End If
End Sub


Private Sub MSFlex1_GotFocus()
ListaRutaAprobacion MSFlex1.TextMatrix(MSFlex1.row, 1)
End Sub

Private Sub MSFlex1_RowColChange()
ListaRutaAprobacion MSFlex1.TextMatrix(MSFlex1.row, 1)
End Sub

'---------------------------------------------------------------------------------

Private Sub MSFlex1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
   If Len(cRHCargoCopia) = 0 Then
      mnuSep.Enabled = False
      mnuPegar.Enabled = False
   Else
      mnuSep.Enabled = True
      mnuPegar.Enabled = True
   End If
   PopupMenu mnuUtil
End If
End Sub

Private Sub MSFlex2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
   PopupMenu mnuUtil
End If
End Sub

Private Sub mnuCopia_Click()
cRHCargoCopia = MSFlex1.TextMatrix(MSFlex1.row, 1)
End Sub

Private Sub mnuPegar_Click()
Dim cRHCargoCod As String, oConn As New DConecta, sSQLX As String

cRHCargoCod = MSFlex1.TextMatrix(MSFlex1.row, 1)

sSQLX = "DELETE FROM LogNivelAprobacion WHERE cRHCargoCod = '" & cRHCargoCod & "'"

sSql = "insert into LogNivelAprobacion (cRHCargoCod,cRHCargoCodAprobacion,nAgencia,nNivelAprobacion,nSector) " & _
       " select '" & cRHCargoCod & "',cRHCargoCodAprobacion,nAgencia,nNivelAprobacion,nSector " & _
       "  from LogNivelAprobacion where cRHCargoCod = '" & cRHCargoCopia & "'"

If oConn.AbreConexion Then
   oConn.Ejecutar sSQLX
   oConn.Ejecutar sSql
   oConn.CierraConexion
   ListaRutaAprobacion cRHCargoCod
End If
End Sub


Private Sub mnuRenu_Click()
Dim oConn As New DConecta, rs As New ADODB.Recordset
Dim cRHCargoCod As String, cRHCargoCodApro As String, nNivel As Integer, i As Integer

cRHCargoCod = MSFlex1.TextMatrix(MSFlex1.row, 1)

sSql = "Select * from LogNivelAprobacion WHERE cRHCargoCod = '" & cRHCargoCod & "' order by nNivelAprobacion"
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSql)
   
   If Not rs.EOF Then
      i = 0
      Do While Not rs.EOF
         i = i + 1
         nNivel = rs!nNivelAprobacion
         cRHCargoCodApro = rs!cRHCargoCodAprobacion
         
         sSql = "UPDATE LogNivelAprobacion SET nNivelAprobacion = " & i & "  " & _
                " WHERE cRHCargoCod = '" & cRHCargoCod & "' and cRHCargoCodAprobacion = '" & cRHCargoCodApro & "' and nNivelAprobacion = " & nNivel & " "
         oConn.Ejecutar sSql
         
         rs.MoveNext
      Loop
   End If
   oConn.CierraConexion
End If
ListaRutaAprobacion cRHCargoCod
End Sub

