VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogVehiculo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Vehiculo"
   ClientHeight    =   4620
   ClientLeft      =   1080
   ClientTop       =   2355
   ClientWidth     =   10230
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraVis 
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   120
      TabIndex        =   36
      Top             =   120
      Width           =   10035
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   7500
         TabIndex        =   39
         Top             =   4020
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Quitar"
         Height          =   375
         Left            =   1260
         TabIndex        =   2
         Top             =   4020
         Width           =   1215
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   8760
         TabIndex        =   20
         Top             =   4020
         Width           =   1215
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   4020
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
         Height          =   3915
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6906
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
   End
   Begin VB.Frame fraReg 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   4395
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   10035
      Begin VB.Frame Frame4 
         Caption         =   "Descripción del Bien Activo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   9975
         Begin VB.TextBox lblDescripcion 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   420
            Width           =   7155
         End
         Begin VB.CommandButton cmdVehic 
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
            Left            =   2160
            TabIndex        =   4
            Top             =   450
            Width           =   375
         End
         Begin VB.TextBox txtBSCod 
            Height          =   315
            Left            =   900
            MaxLength       =   50
            TabIndex        =   3
            Top             =   420
            Width           =   1650
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Vehiculo"
            Height          =   195
            Left            =   180
            TabIndex        =   40
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Características Externas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   0
         TabIndex        =   26
         Top             =   1140
         Width           =   4905
         Begin VB.OptionButton opTipo2 
            Caption         =   "Motocicleta"
            Height          =   255
            Left            =   2940
            TabIndex        =   6
            Top             =   420
            Width           =   1575
         End
         Begin VB.OptionButton opTipo1 
            Caption         =   "Auto / Camioneta"
            Height          =   255
            Left            =   900
            TabIndex        =   5
            Top             =   420
            Width           =   1575
         End
         Begin VB.ComboBox CboVehiculo 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   900
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   900
            Width           =   3750
         End
         Begin VB.ComboBox CboColor 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2160
            Width           =   3750
         End
         Begin VB.ComboBox CboMarca 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   900
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1320
            Width           =   3750
         End
         Begin VB.TextBox TxtPlaca 
            Height          =   315
            Left            =   900
            MaxLength       =   4
            TabIndex        =   11
            Top             =   2640
            Width           =   645
         End
         Begin VB.TextBox txtPlacaNro 
            Height          =   315
            Left            =   1860
            MaxLength       =   5
            TabIndex        =   12
            Top             =   2640
            Width           =   915
         End
         Begin VB.ComboBox CboAño 
            Height          =   315
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2640
            Width           =   1065
         End
         Begin VB.TextBox TxtModelo 
            Height          =   315
            Left            =   900
            MaxLength       =   50
            TabIndex        =   9
            Top             =   1740
            Width           =   3750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Placa"
            Height          =   195
            Left            =   180
            TabIndex        =   33
            Top             =   2700
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo "
            Height          =   195
            Left            =   180
            TabIndex        =   32
            Top             =   960
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Modelo"
            Height          =   195
            Left            =   180
            TabIndex        =   31
            Top             =   1800
            Width           =   525
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Año"
            Height          =   195
            Left            =   3180
            TabIndex        =   30
            Top             =   2700
            Width           =   285
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Color"
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   2220
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Marca"
            Height          =   195
            Left            =   180
            TabIndex        =   28
            Top             =   1380
            Width           =   450
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1620
            TabIndex        =   27
            Top             =   2580
            Width           =   120
         End
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   8820
         TabIndex        =   19
         Top             =   3900
         Width           =   1155
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   7500
         TabIndex        =   18
         Top             =   3900
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Características Internas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   5040
         TabIndex        =   23
         Top             =   1140
         Width           =   4935
         Begin VB.TextBox txtSerie 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   21
            Top             =   360
            Width           =   3390
         End
         Begin VB.ComboBox cboCombustible 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   780
            Width           =   3390
         End
         Begin VB.TextBox TxtNroMotor 
            Height          =   315
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   15
            Top             =   1200
            Width           =   3390
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Nro Serie"
            Height          =   195
            Left            =   180
            TabIndex        =   38
            Top             =   420
            Width           =   660
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nro Motor"
            Height          =   195
            Left            =   180
            TabIndex        =   25
            Top             =   1260
            Width           =   705
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Combustible"
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   795
         Left            =   5040
         TabIndex        =   34
         Top             =   2880
         Width           =   4935
         Begin MSMask.MaskEdBox txtFechaIni 
            Height          =   315
            Left            =   1260
            TabIndex        =   16
            Top             =   300
            Width           =   1200
            _ExtentX        =   2117
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
         Begin MSMask.MaskEdBox txtFechaFin 
            Height          =   315
            Left            =   3420
            TabIndex        =   17
            Top             =   300
            Width           =   1200
            _ExtentX        =   2117
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
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   2880
            TabIndex        =   41
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "SOAT Desde"
            Height          =   195
            Left            =   180
            TabIndex        =   35
            Top             =   360
            Width           =   945
         End
      End
   End
End
Attribute VB_Name = "frmLogVehiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub opTipo1_Click()
CargaTiposMarcas 1
End Sub

Private Sub opTipo2_Click()
CargaTiposMarcas 2
End Sub

Sub CargaTiposMarcas(pnTipo As Integer)
Dim i As Integer, j As Integer
Dim DV As DLogVehiculos
Dim rs As New ADODB.Recordset

If opTipo1.value Or opTipo2.value Then
   CboVehiculo.BackColor = "&H80000005"
   CboMarca.BackColor = "&H80000005"
   CboVehiculo.Locked = False
   CboMarca.Locked = False
End If

Set DV = New DLogVehiculos

CboMarca.Clear
Set rs = DV.GetMarcasVehiculo(pnTipo)
If Not (rs.EOF And rs.BOF) Then
    While Not rs.EOF
        CboMarca.AddItem Trim(rs(1)) & Space(100) & rs(0)
        rs.MoveNext
    Wend
    CboMarca.ListIndex = 0
End If

CboVehiculo.Clear
Set rs = DV.GetTiposVehiculo(pnTipo)
If Not (rs.EOF And rs.BOF) Then
    While Not rs.EOF
        CboVehiculo.AddItem Trim(rs(1)) & Space(100) & rs(0)
        rs.MoveNext
    Wend
    CboVehiculo.ListIndex = 0
End If

End Sub

Private Sub cmdVehic_Click()
Dim oConn As New DConecta, rs As New ADODB.Recordset

frmLogVehiculoLista.Show 1

If frmLogVehiculoLista.vpSeleccion Then
   If oConn.AbreConexion Then
      Set rs = oConn.CargaRecordSet("select nVehiculoCod from LogVehiculoReg where cBSCod = '" & Trim(frmLogVehiculoLista.vpCodigo) & "' and cBSSerie='" & Trim(frmLogVehiculoLista.vpSerie) & "'")
      If Not rs.EOF Then
         MsgBox "Ya se ha registrado el vehículo indicado..." + Space(10), vbInformation, "Aviso"
         Exit Sub
      Else
         txtBSCod.Text = frmLogVehiculoLista.vpCodigo
         txtSerie.Text = frmLogVehiculoLista.vpSerie
         lblDescripcion = frmLogVehiculoLista.vpDescripcion
         opTipo1.SetFocus
      End If
      oConn.CierraConexion
   End If
End If

End Sub

Private Sub Form_Load()
CentraForm Me
txtFechaIni = Date
CargaLista
End Sub

Sub CargaLista()
Dim i As Integer, rs As New ADODB.Recordset
Dim sSQL As String, oConn As DConecta

Set oConn = New DConecta
FormaFlex 0
If oConn.AbreConexion Then
  
sSQL = "select r.nTipo, r.cBSCod,r.cBSSerie, r.nVehiculoCod,t.cTipoVehiculo,m.cMarca,c.cColor, r.cPlaca, r.cModelo, r.nAnioFab, r.nCombustible " & _
       " from LogVehiculoReg r " & _
       " inner join (select nConsValor AS nTipoVehiculo,cConsDescripcion as cTipoVehiculo from Constante where nConsCod=9026 and nconscod<>nconsvalor) t on r.nTipoVehiculo = t.nTipoVehiculo  " & _
       " inner join (select nConsValor AS nMarca,cConsDescripcion as cMarca from Constante where nConsCod=9022 and nconscod<>nconsvalor) m on r.nMarca = m.nMarca " & _
       " inner join (select nConsValor AS nColor,cConsDescripcion as cColor from Constante where nConsCod=9023 and nconscod<>nconsvalor) c on r.nColor = c.nColor " & _
       "   "
       
'Where R.nTipo = 1

   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      i = 0
      Do While Not rs.EOF
         i = i + 1
         InsRow MSH, i
         MSH.TextMatrix(i, 0) = Format(i, "00")
         MSH.TextMatrix(i, 1) = rs!cBSCod    'rs!cPlaca
         MSH.TextMatrix(i, 2) = rs!cBSSerie  'rs!cModelo
         MSH.TextMatrix(i, 3) = rs!cTipoVehiculo + " " + rs!cModelo + " - " + rs!cMarca + " " + " COLOR " + rs!cColor
         MSH.TextMatrix(i, 4) = rs!cPlaca    'rs!cColor
         'MSH.TextMatrix(i, 5) = IIf(rs!nEstado = 3, "EN REPARACION", "")
         rs.MoveNext
      Loop
   End If
   oConn.CierraConexion
End If
End Sub

Sub FormaFlex(nLineas As Integer)
MSH.Clear
MSH.Font = "Tahoma"
MSH.Font.Size = 7
MSH.RowHeight(0) = 300
Select Case nLineas
    Case 0
         MSH.Rows = 2
         MSH.RowHeight(1) = 8
    Case 1
         MSH.Rows = 2
         MSH.RowHeight(1) = 260
    Case Is > 1
         MSH.Rows = nLineas + 1
         MSH.RowHeight(1) = 260
End Select
MSH.ColWidth(0) = 300:   MSH.ColAlignment(0) = 4
MSH.ColWidth(1) = 900:   MSH.TextMatrix(0, 1) = " BSCod": MSH.ColAlignment(1) = 4
MSH.ColWidth(2) = 2000:  MSH.TextMatrix(0, 2) = " BSSerie": MSH.ColAlignment(2) = 1
MSH.ColWidth(3) = 5400:  MSH.TextMatrix(0, 3) = " Descripcion"
MSH.ColWidth(4) = 1100:  MSH.TextMatrix(0, 4) = " Placa": MSH.ColAlignment(4) = 4
MSH.ColWidth(5) = 0:     MSH.TextMatrix(0, 5) = " "
MSH.ColWidth(6) = 0:     MSH.TextMatrix(0, 6) = " "
MSH.ColWidth(7) = 0:     MSH.TextMatrix(0, 7) = " "
End Sub

Sub CargaCombos()
Dim i As Integer, j As Integer
Dim DV As DLogVehiculos
Dim rs As New ADODB.Recordset

Set DV = New DLogVehiculos
j = Year(Date)

CboAño.Clear
For i = 0 To 20
    CboAño.AddItem j - i
Next i

'CboVehiculo.Clear

CboColor.Clear
Set rs = DV.GetColorV
If Not (rs.EOF And rs.BOF) Then
    While Not rs.EOF
        CboColor.AddItem Trim(rs(1)) & Space(100) & rs(0)
        rs.MoveNext
    Wend
    CboColor.ListIndex = 0
End If

cboCombustible.Clear
Set rs = DV.GetCombustible
If Not (rs.EOF And rs.BOF) Then
    While Not rs.EOF
        cboCombustible.AddItem Trim(rs(1)) & Space(100) & rs(0)
        rs.MoveNext
    Wend
    cboCombustible.ListIndex = 0
End If
Set rs = Nothing
Set DV = Nothing
End Sub

Private Sub cmdAceptar_Click()
Dim opt As Integer
Dim DV As DLogVehiculos
Dim oConn As New DConecta, rs As New ADODB.Recordset
Dim nCombustible As Integer
Dim nTipVehiculo As Integer
Dim nMarcaV As Integer
Dim nColoV As Integer
Dim nTipo As Integer

Set DV = New DLogVehiculos

If Len(Trim(txtBSCod.Text)) = 0 Then
   MsgBox "Debe indicar el Vehículo como Bien de la institución..." + Space(10), vbInformation, "Aviso"
   Exit Sub
End If

If Len(Trim(txtSerie.Text)) = 0 Then
   MsgBox "Debe indicar el Vehículo como Bien de la institución..." + Space(10), vbInformation, "Aviso"
   Exit Sub
End If

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet("select nVehiculoCod from LogVehiculoReg where cBSCod = '" & Trim(txtBSCod.Text) & "' and cBSSerie='" & Trim(txtSerie.Text) & "'")
   If Not rs.EOF Then
      MsgBox "Ya se ha registrado el vehículo indicado..." + Space(10), vbInformation, "Aviso"
      Exit Sub
   End If
   oConn.CierraConexion
End If

If Me.CboVehiculo.ListIndex = -1 Then
    MsgBox "Elija un Tipo de Vehículo", vbInformation, "AVISO"
    Exit Sub
End If

If Me.CboMarca.ListIndex = -1 Then
    MsgBox "Elija un Tipo de Marca del Vehículo", vbInformation, "AVISO"
    Exit Sub
End If

If Me.CboColor.ListIndex = -1 Then
    MsgBox "Elija un Tipo de Color", vbInformation, "AVISO"
    Exit Sub
End If

If Me.CboAño.ListIndex = -1 Then
    MsgBox "Elija el año de fabricación Vehículo", vbInformation, "AVISO"
    Exit Sub
End If

If Trim(Me.TxtModelo) = "" Then
    MsgBox "Ingrese el modelo del vehiculo", vbInformation, "AVISO"
    Exit Sub
End If

If Trim(Me.TxtNroMotor) = "" Then
    MsgBox "Ingrese el nro del motor", vbInformation, "AVISO"
    Exit Sub
End If

If Trim(Me.TxtPlaca) = "" Or Trim(Me.txtPlacaNro) = "" Then
    MsgBox "Ingrese el de Placa Correctamente", vbInformation, "AVISO"
    Exit Sub
End If

nTipo = IIf(opTipo1.value, 1, IIf(opTipo2.value, 2, 0))

opt = MsgBox("¿ Esta seguro de grabar el vehículo indicado ?" + Space(10), vbQuestion + vbYesNo, "AVISO")
If vbNo = opt Then Exit Sub

nTipVehiculo = CInt(Right(Me.CboVehiculo, 3))
nMarcaV = CInt(Right(Me.CboMarca, 3))
nColoV = CInt(Right(Me.CboColor, 3))

If cboCombustible.ListIndex >= 0 Then
   nCombustible = CInt(Right(Me.cboCombustible, 3))
Else
   MsgBox "Debe indicar el combustible..." + Space(10), vbInformation
   Exit Sub
End If

Call DV.InsertaRegVehiculo(txtBSCod.Text, txtSerie, nTipo, Me.TxtModelo, Me.CboAño.Text, 0, nCombustible, Me.TxtNroMotor, Me.TxtPlaca + "-" + Me.txtPlacaNro, nTipVehiculo, nMarcaV, nColoV, txtFechaIni, txtFechaFin, gdFecSis, Right(gsCodAge, 2), gsCodUser)
CargaLista
CmdCancelar_Click
End Sub

Private Sub CmdCancelar_Click()
fraVis.Visible = True
fraReg.Visible = False
End Sub

Private Sub cmdNuevo_Click()
opTipo1.value = False
opTipo2.value = False

CboMarca.Clear
CboVehiculo.Clear
CboVehiculo.BackColor = "&H8000000F"
CboMarca.BackColor = "&H8000000F"
CboVehiculo.Locked = True
CboMarca.Locked = True

CargaCombos
fraVis.Visible = False
fraReg.Visible = True
txtBSCod.Text = ""
txtSerie.Text = ""
TxtModelo.Text = ""
TxtPlaca.Text = ""
txtPlacaNro.Text = ""
TxtNroMotor.Text = ""
lblDescripcion.Text = ""
txtFechaIni = Date
txtFechaFin = Date
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub Marco()
Dim rs As ADODB.Recordset
Dim LV As DLogVehiculos
Dim i As Integer
Set LV = New DLogVehiculos

With MSH
    .Rows = 2
    .Clear
    .ColWidth(0) = 100
    .ColWidth(1) = 1200
    .ColWidth(2) = 1800
    .ColWidth(3) = 1300
    .ColWidth(4) = 800
    .ColWidth(5) = 1600
    .ColWidth(6) = 1500
    .ColWidth(7) = 1800
    .ColWidth(8) = 1500
    .ColWidth(9) = 1700
    .ColWidth(10) = 1800
    .ColWidth(11) = 1000
    
    .TextMatrix(0, 1) = "BSCod"
    .TextMatrix(0, 2) = "BSSerie"
    .TextMatrix(0, 3) = "Modelo"
    .TextMatrix(0, 4) = "Año Fab"
    .TextMatrix(0, 5) = "Estado"
    .TextMatrix(0, 6) = "Combustible"
    .TextMatrix(0, 7) = "Nro Motor"
    .TextMatrix(0, 8) = "Placa"
    .TextMatrix(0, 9) = "Tipo V"
    .TextMatrix(0, 10) = "Marca"
    .TextMatrix(0, 11) = "Coloc"
End With
Set rs = LV.GetVehiculos(0)
i = 1
While Not rs.EOF
    MSH.TextMatrix(i, 1) = rs!cBSCod
    MSH.TextMatrix(i, 2) = rs!cBSSerie
    MSH.TextMatrix(i, 3) = rs!cModelo
    MSH.TextMatrix(i, 4) = rs!nAñoFab
    MSH.TextMatrix(i, 5) = rs!nEstado
    MSH.TextMatrix(i, 6) = rs!Combustible
    MSH.TextMatrix(i, 7) = rs!cNroMotor
    MSH.TextMatrix(i, 8) = rs!cPlaca
    MSH.TextMatrix(i, 9) = rs!TipoV
    MSH.TextMatrix(i, 10) = rs!Marcar
    MSH.TextMatrix(i, 11) = rs!Color
    MSH.Rows = MSH.Rows + 1
    rs.MoveNext
    i = 1 + i
Wend
If Not (rs.EOF And rs.BOF) Then MSH.Rows = MSH.Rows - 1
End Sub


Public Function intfMayusculas(intTecla As Integer) As Integer
 If Chr(intTecla) >= "a" And Chr(intTecla) <= "z" Then
    intTecla = intTecla - 32
 End If
 If intTecla = 39 Then
    intTecla = 0
 End If
 If intTecla = 209 Or intTecla = 241 Or intTecla = 8 Or intTecla = 32 Then
    intfMayusculas = Asc(UCase(Chr(intTecla)))
     Exit Function
 End If
 intfMayusculas = intTecla
End Function

Private Sub MSH_DblClick()
frmLogisticaVehiculoDetalle.Inicio MSH.TextMatrix(MSH.row, 5), MSH.TextMatrix(MSH.row, 6)
If frmLogisticaVehiculoDetalle.vpHayCambios Then
   CargaLista
End If
End Sub

Private Sub MSH_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyInsert Then
'   Me.Hide
'   frmLogisticaVehiculoAsigna.Vehiculo MSH.TextMatrix(MSH.row, 5), MSH.TextMatrix(MSH.row, 6)
'   Me.Show 1
'End If
End Sub


Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   CmdAceptar.SetFocus
End If
End Sub

Private Sub txtFechaIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFechaFin.SetFocus
End If
End Sub

Private Sub txtPlacaNro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   CboAño.SetFocus
End If
End Sub

Private Sub TxtModelo_KeyPress(KeyAscii As Integer)
KeyAscii = intfMayusculas(KeyAscii)
If KeyAscii = 13 Then
   CboColor.SetFocus
End If
End Sub

Private Sub TxtNroMotor_KeyPress(KeyAscii As Integer)
KeyAscii = intfMayusculas(KeyAscii)
If KeyAscii = 13 Then
   txtFechaIni.SetFocus
End If
End Sub

Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
KeyAscii = intfMayusculas(KeyAscii)
If KeyAscii = 13 Then
   txtPlacaNro.SetFocus
End If
End Sub

Private Sub CboVehiculo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   CboMarca.SetFocus
End If
End Sub

Private Sub CboMarca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TxtModelo.SetFocus
End If
End Sub

Private Sub CboColor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TxtPlaca.SetFocus
End If
End Sub

Private Sub CboAño_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboCombustible.SetFocus
End If
End Sub

Private Sub cboCombustible_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TxtNroMotor.SetFocus
End If
End Sub

