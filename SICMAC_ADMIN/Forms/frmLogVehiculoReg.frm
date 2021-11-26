VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3A34C7E1-4D73-49FB-9EA9-C6D17991498B}#10.0#0"; "MiPanel.ocx"
Begin VB.Form frmLogVehiculoReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Vehiculo"
   ClientHeight    =   4200
   ClientLeft      =   525
   ClientTop       =   2775
   ClientWidth     =   10815
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraVis 
      BorderStyle     =   0  'None
      Height          =   4035
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   10575
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   8100
         TabIndex        =   36
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Quitar"
         Height          =   375
         Left            =   1260
         TabIndex        =   2
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   9360
         TabIndex        =   3
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   3660
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
         Height          =   3555
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   6271
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
      Height          =   3975
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   10575
      Begin VB.Frame Frame4 
         Height          =   795
         Left            =   0
         TabIndex        =   33
         Top             =   -60
         Width           =   10515
         Begin LabelBoxOCX.LabelBox lblDescripcion 
            Height          =   315
            Left            =   2580
            TabIndex        =   35
            Top             =   300
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   556
            FteColor        =   -2147483630
            BeginProperty Fuente {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            TabIndex        =   5
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox txtBSCod 
            Height          =   315
            Left            =   900
            MaxLength       =   50
            TabIndex        =   4
            Top             =   300
            Width           =   1650
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Activo"
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
            TabIndex        =   37
            Top             =   360
            Width           =   555
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
         Height          =   2550
         Left            =   0
         TabIndex        =   22
         Top             =   900
         Width           =   5205
         Begin VB.ComboBox CboVehiculo 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   360
            Width           =   4050
         End
         Begin VB.ComboBox CboColor 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1620
            Width           =   4050
         End
         Begin VB.ComboBox CboMarca 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   780
            Width           =   4050
         End
         Begin VB.TextBox TxtPlaca 
            Height          =   315
            Left            =   900
            MaxLength       =   4
            TabIndex        =   10
            Top             =   2040
            Width           =   645
         End
         Begin VB.TextBox txtPlacaNro 
            Height          =   315
            Left            =   1860
            MaxLength       =   5
            TabIndex        =   11
            Top             =   2040
            Width           =   915
         End
         Begin VB.ComboBox CboAño 
            Height          =   315
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2040
            Width           =   1065
         End
         Begin VB.TextBox TxtModelo 
            Height          =   315
            Left            =   900
            MaxLength       =   50
            TabIndex        =   8
            Top             =   1200
            Width           =   4050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Placa"
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   2100
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo "
            Height          =   195
            Left            =   180
            TabIndex        =   28
            Top             =   420
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Modelo"
            Height          =   195
            Left            =   180
            TabIndex        =   27
            Top             =   1260
            Width           =   525
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Año"
            Height          =   195
            Left            =   3480
            TabIndex        =   26
            Top             =   2100
            Width           =   285
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Color"
            Height          =   195
            Left            =   180
            TabIndex        =   25
            Top             =   1680
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Marca"
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   840
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
            TabIndex        =   23
            Top             =   1980
            Width           =   120
         End
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   9300
         TabIndex        =   17
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   8040
         TabIndex        =   16
         Top             =   3600
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
         Left            =   5400
         TabIndex        =   19
         Top             =   900
         Width           =   5115
         Begin VB.TextBox txtSerie 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   13
            Top             =   360
            Width           =   3630
         End
         Begin VB.ComboBox cboCombustible 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   780
            Width           =   3630
         End
         Begin VB.TextBox TxtNroMotor 
            Height          =   315
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   15
            Top             =   1200
            Width           =   3630
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Nro Serie"
            Height          =   195
            Left            =   180
            TabIndex        =   34
            Top             =   420
            Width           =   660
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nro Motor"
            Height          =   195
            Left            =   180
            TabIndex        =   21
            Top             =   1260
            Width           =   705
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Combustible"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   795
         Left            =   5400
         TabIndex        =   30
         Top             =   2640
         Width           =   5115
         Begin MSMask.MaskEdBox txtFechaIni 
            Height          =   315
            Left            =   1260
            TabIndex        =   38
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
            Left            =   3720
            TabIndex        =   39
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
            Caption         =   "Vigente hasta"
            Height          =   195
            Left            =   2640
            TabIndex        =   40
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "SOAT Desde"
            Height          =   195
            Left            =   180
            TabIndex        =   31
            Top             =   360
            Width           =   945
         End
      End
   End
End
Attribute VB_Name = "FrmLogVehiculoReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdVehic_Click()
frmLogVehiculoLista.Show 1
End Sub

Private Sub Form_Load()
txtFechaIni = Date
CargaLista
End Sub

Sub CargaLista()
Dim i As Integer, rs As New ADODB.Recordset
Dim sSQL As String, oConn As DConecta

Set oConn = New DConecta
FormaFlex 0
If oConn.AbreConexion Then
  
sSQL = "select LV.cBSCod, LV.cBSSerie, LV.cModelo, LV.nAnioFab, " & _
       "       LV.cPlaca, LT.cDescripcion, MV.cMarca, LC.cColor, LV.nEstado " & _
       " from LogVehiculoReg LV inner join LogisticaTipoVehiculo LT on LT.nTipoVehiculo = LV.nTipoVehiculo" & _
       "   inner join logisticaMarcav MV on MV.nMarca = LV.nMarca  " & _
       "  inner join logisticacolor LC on LC.nColor = LV.nColor " & _
       "   order by LT.cDescripcion "
       
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      i = 0
      Do While Not rs.EOF
         i = i + 1
         InsRow MSH, i
         MSH.TextMatrix(i, 0) = Format(i, "00")
         MSH.TextMatrix(i, 1) = rs!cBSCod    'rs!cPlaca
         MSH.TextMatrix(i, 2) = rs!cBSSerie  'rs!cModelo
         MSH.TextMatrix(i, 3) = rs!cDescripcion + " " + rs!cModelo + " - " + rs!cMarca + " " + " COLOR " + rs!cColor
         MSH.TextMatrix(i, 4) = rs!cPlaca    'rs!cColor
         MSH.TextMatrix(i, 5) = IIf(rs!nEstado = 3, "EN REPARACION", "")
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
MSH.ColWidth(3) = 6000:  MSH.TextMatrix(0, 3) = " Descripcion"
MSH.ColWidth(4) = 1100:  MSH.TextMatrix(0, 4) = " Placa": MSH.ColAlignment(4) = 4
MSH.ColWidth(5) = 0:     MSH.TextMatrix(0, 5) = " "
MSH.ColWidth(6) = 0:     MSH.TextMatrix(0, 6) = " "
MSH.ColWidth(7) = 0:     MSH.TextMatrix(0, 7) = " "
End Sub

Sub CargaCombos()
Dim i As Integer, j As Integer
Dim DV As DLogVehiculo
Dim rs As New ADODB.Recordset

Set DV = New DLogVehiculo
j = Year(Date)

CboAño.Clear
For i = 0 To 20
    CboAño.AddItem j - i
Next i

CboVehiculo.Clear
Set rs = DV.GetTipoVehiculo
If Not (rs.EOF And rs.BOF) Then
    While Not rs.EOF
        CboVehiculo.AddItem Trim(rs(0)) & Space(100) & rs(1)
        rs.MoveNext
    Wend
End If

CboMarca.Clear
Set rs = DV.GetMarcaV
If Not (rs.EOF And rs.BOF) Then
    While Not rs.EOF
        CboMarca.AddItem Trim(rs(0)) & Space(100) & rs(1)
        rs.MoveNext
    Wend
End If

CboColor.Clear
Set rs = DV.GetColorV
If Not (rs.EOF And rs.BOF) Then
    While Not rs.EOF
        CboColor.AddItem Trim(rs(0)) & Space(100) & rs(1)
        rs.MoveNext
    Wend
End If

cboCombustible.Clear
cboCombustible.AddItem "GASOLINA"
cboCombustible.AddItem "PETROLEO"

Set rs = Nothing
Set DV = Nothing
End Sub

Private Sub CmdAceptar_Click()
Dim opt As Integer
Dim DV As DLogVehiculos
Dim nCombustible As Integer
Dim nTipVehiculo As Integer
Dim nMarcaV As Integer
Dim nColoV As Integer

Set DV = New DLogVehiculos

If Len(Trim(txtBSCod.Text)) = 0 Then
   MsgBox "Debe indicar el Vehículo como Bien de la institución..." + Space(10), vbInformation, "Aviso"
   Exit Sub
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

opt = MsgBox("Esta seguro de Guardar", vbQuestion + vbYesNo, "AVISO")
If vbNo = opt Then Exit Sub

nTipVehiculo = CInt(Right(Me.CboVehiculo, 3))
nMarcaV = CInt(Right(Me.CboMarca, 3))
nColoV = CInt(Right(Me.CboColor, 3))

If cboCombustible.ListIndex >= 0 Then
   nCombustible = cboCombustible.ListIndex
Else
   MsgBox "Debe indicar el combustible..." + Space(10), vbInformation
   Exit Sub
End If

Call DV.InsertaRegVehiculo(txtBSCod.Text, txtSerie, Me.TxtModelo, Me.CboAño.Text, 1, nCombustible, Me.TxtNroMotor, Me.TxtPlaca + "-" + Me.txtPlacaNro, nTipVehiculo, nMarcaV, nColoV, txtFechaIni, txtFechaFin, gdFecSis, Right(gsCodAge, 2), gsCodUser)
CargaLista
CmdCancelar_Click
End Sub

Private Sub CmdCancelar_Click()
fraVis.Visible = True
fraReg.Visible = False
End Sub

Private Sub cmdNuevo_Click()
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
Dim LV As DLogVehiculo
Dim i As Integer
Set LV = New DLogVehiculo

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
   CmdAceptar.SetFocus
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

