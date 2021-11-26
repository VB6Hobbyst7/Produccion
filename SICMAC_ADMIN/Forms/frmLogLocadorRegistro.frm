VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogLocadorRegistro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Locadores"
   ClientHeight    =   5040
   ClientLeft      =   570
   ClientTop       =   2055
   ClientWidth     =   11130
   Icon            =   "frmLogLocadorRegistro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   11130
   Begin VB.Frame fraVis 
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
      Height          =   4960
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10995
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "Actualizar Datos"
         Height          =   375
         Left            =   2640
         TabIndex        =   45
         Top             =   4500
         Width           =   1755
      End
      Begin VB.CommandButton cmdReportes 
         Caption         =   "Reportes"
         Height          =   375
         Left            =   8340
         TabIndex        =   40
         Top             =   4500
         Width           =   1275
      End
      Begin VB.TextBox txtFechaFin 
         Height          =   315
         Left            =   4320
         TabIndex        =   30
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtFechaIni 
         Height          =   315
         Left            =   2340
         TabIndex        =   29
         Top             =   300
         Width           =   1035
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSLista 
         Height          =   3675
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6482
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   10
         FixedCols       =   0
         BackColorSel    =   13887482
         ForeColorSel    =   128
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   10
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   9660
         TabIndex        =   7
         Top             =   4500
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   1380
         TabIndex        =   6
         Top             =   4500
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   4500
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Contrato vigente desde                        hasta"
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
         TabIndex        =   28
         Top             =   360
         Width           =   3900
      End
   End
   Begin VB.Frame fraReg 
      Caption         =   "Registro de Personal por Locación de Servicios "
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
      Height          =   4860
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   10995
      Begin VB.Frame Frame5 
         Height          =   795
         Left            =   5640
         TabIndex        =   32
         Top             =   720
         Width           =   5175
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            ItemData        =   "frmLogLocadorRegistro.frx":08CA
            Left            =   720
            List            =   "frmLogLocadorRegistro.frx":08D4
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   300
            Width           =   915
         End
         Begin VB.TextBox txtMonto 
            Height          =   315
            Left            =   1620
            TabIndex        =   34
            Top             =   300
            Width           =   1455
         End
         Begin VB.TextBox txtCuotas 
            Height          =   315
            Left            =   4380
            TabIndex        =   33
            Top             =   300
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Cuotas"
            Height          =   195
            Left            =   3300
            TabIndex        =   35
            Top             =   360
            Width           =   945
         End
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   8400
         TabIndex        =   39
         Top             =   4410
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   9660
         TabIndex        =   38
         Top             =   4410
         Width           =   1155
      End
      Begin VB.Frame Frame2 
         Caption         =   "Vigencia Contrato "
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
         Height          =   1515
         Left            =   3540
         TabIndex        =   20
         Top             =   720
         Width           =   2055
         Begin VB.TextBox txtFecIni 
            Height          =   315
            Left            =   720
            TabIndex        =   22
            Top             =   300
            Width           =   1140
         End
         Begin VB.TextBox txtFecFin 
            Height          =   315
            Left            =   720
            TabIndex        =   21
            Top             =   660
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   180
            TabIndex        =   23
            Top             =   720
            Width           =   420
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Locación del Servicio "
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
         Height          =   1995
         Left            =   180
         TabIndex        =   11
         Top             =   2340
         Width           =   5415
         Begin VB.CommandButton cmdFunciones 
            Caption         =   "Funciones de Locador"
            Height          =   375
            Left            =   960
            TabIndex        =   44
            Top             =   1530
            Width           =   1875
         End
         Begin VB.ComboBox cboFuncion 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1140
            Width           =   4275
         End
         Begin VB.ComboBox cboArea 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   720
            Width           =   4275
         End
         Begin VB.ComboBox cboAge 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   300
            Width           =   4275
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Función"
            Height          =   195
            Left            =   180
            TabIndex        =   19
            Top             =   1200
            Width           =   570
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Area"
            Height          =   195
            Left            =   180
            TabIndex        =   17
            Top             =   780
            Width           =   330
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Agencia"
            Height          =   195
            Left            =   180
            TabIndex        =   15
            Top             =   360
            Width           =   585
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Contrato "
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
         Height          =   1515
         Left            =   180
         TabIndex        =   8
         Top             =   720
         Width           =   3315
         Begin VB.TextBox Text1 
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
            Left            =   660
            TabIndex        =   46
            Text            =   "CLS"
            Top             =   300
            Width           =   495
         End
         Begin VB.TextBox txtNroContrato 
            Height          =   315
            Left            =   1140
            TabIndex        =   9
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox Text2 
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
            Left            =   2340
            TabIndex        =   47
            Text            =   "CMAC-T"
            Top             =   300
            Width           =   795
         End
         Begin VB.CheckBox chkCAP 
            Alignment       =   1  'Right Justify
            Caption         =   "Esta considerado en el C.A.P."
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   660
            TabIndex        =   41
            Top             =   1140
            Width           =   2475
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   660
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   660
            Width           =   2475
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
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
            TabIndex        =   13
            Top             =   720
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   180
            TabIndex        =   10
            Top             =   300
            Width           =   300
         End
      End
      Begin VB.TextBox txtPersona 
         BackColor       =   &H00EAFFFF&
         Height          =   315
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   330
         Width           =   8565
      End
      Begin VB.CommandButton cmdPersona 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1575
         TabIndex        =   3
         Top             =   360
         Width           =   390
      End
      Begin VB.TextBox txtPersCod 
         BackColor       =   &H00EAFFFF&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   330
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Height          =   2895
         Left            =   5640
         TabIndex        =   25
         Top             =   1440
         Width           =   5175
         Begin VB.CommandButton cmdAgrega 
            Caption         =   "Agregar Concepto Pago"
            Height          =   375
            Left            =   180
            TabIndex        =   43
            Top             =   2430
            Width           =   2055
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00DDFFFE&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2580
            TabIndex        =   31
            Top             =   1080
            Visible         =   0   'False
            Width           =   990
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
            Height          =   2070
            Left            =   180
            TabIndex        =   27
            Top             =   300
            Width           =   4800
            _ExtentX        =   8467
            _ExtentY        =   3651
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   5
            FixedCols       =   0
            BackColorSel    =   16773352
            ForeColorSel    =   16711680
            BackColorBkg    =   -2147483643
            GridColor       =   -2147483633
            FocusRect       =   0
            HighLight       =   2
            ScrollBars      =   2
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
         Begin VB.TextBox txtTotal 
            BackColor       =   &H00FFF8E1&
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   1620
            Visible         =   0   'False
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmLogLocadorRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HaCargado As Boolean
Dim cAgeCod As String, cAreaCod As String
Dim nEditable As Boolean, sSQL As String
Dim HaGrabado As Boolean
Dim nFuncion As Integer

Private Sub MSLista_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   VerDatos
End If
End Sub

Private Sub VerDatos()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta, i As Integer
Dim cPersCod As String, nRegLoc As Integer
Dim nCuota As Integer

nFuncion = 0

If Not HaCargado Then
   CargaCombosFuncion
End If
nEditable = False
cPersCod = MSLista.TextMatrix(MSLista.row, 1)
nRegLoc = MSLista.TextMatrix(MSLista.row, 0)

cmdFunciones.Visible = False
cmdAgrega.Visible = False
cmdGrabar.Visible = False
cmdPersona.Visible = False
DoEvents
If oConn.AbreConexion Then

   sSQL = "select x.nRegLocador, x.cPersCod, cPersona=replace(p.cPersNombre,'/',' '),cNroContrato,dFechaInicio,dFechaTermino, " & _
   "      aa.cAreaDescripcion as cArea, a.cAgeDescripcion as cAgencia,x.nMontoContrato,x.nMoneda, " & _
   "      cTipoContrato = coalesce(t.cConsDescripcion,'') ,cFuncion = coalesce(f.cConsDescripcion,'')  " & _
   " from LogLocadores x inner join Persona p on x.cPersCod=p.cPersCod " & _
   "                     inner join Agencias a on x.cAgeCod = a.cAgeCod " & _
   "                     inner join Areas aa on x.cAreaCod = aa.cAreaCod " & _
   "                     left join (select nConsValor,cConsDescripcion from Constante where nConsCod =9131 and nConsCod<>nConsValor) t on x.nTipoContrato = t.nConsValor " & _
   "                     left join (select nConsValor,cConsDescripcion from Constante where nConsCod =9132 and nConsCod<>nConsValor) f on x.nFuncionCod = f.nConsValor " & _
   " where x.cPersCod = '" & cPersCod & "'"

   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      nRegLoc = rs!nRegLocador
      txtPersCod.Text = rs!cPersCod
      txtPersona.Text = rs!cPersona:         txtPersona.BackColor = "&H00EFEFEF"
      txtNroContrato.Text = rs!cNroContrato: txtNroContrato.BackColor = "&H00EFEFEF":  txtNroContrato.Locked = True
      cboTipo.BackColor = "&H00EFEFEF":
      txtFecIni = rs!dFechaInicio:           txtFecIni.BackColor = "&H00EFEFEF":       txtFecIni.Locked = True
      txtFecFin = rs!dFechaTermino:          txtFecFin.BackColor = "&H00EFEFEF":       txtFecFin.Locked = True
      cboFuncion.BackColor = "&H00EFEFEF"
      txtMonto.Text = FNumero(rs!nMontoContrato): txtMonto.BackColor = "&H00EFEFEF":   txtMonto.Locked = True
      cboAge.Text = rs!cAgencia:             cboAge.BackColor = "&H00EFEFEF":          cboAge.Locked = True
      cboArea.Text = rs!cArea:               cboArea.BackColor = "&H00EFEFEF":         cboArea.Locked = True
      
      cboMoneda.ListIndex = rs!nMoneda - 1
      cboMoneda.BackColor = "&H00EFEFEF": cboMoneda.Locked = True
      
      If Len(rs!cTipoContrato) > 0 Then
         cboTipo.Text = rs!cTipoContrato
      Else
         cboTipo.ListIndex = -1
      End If
      cboFuncion.Locked = True
      
      If Len(Trim(rs!cFuncion)) > 0 Then
         cboFuncion.Text = rs!cFuncion
      Else
         cboFuncion.ListIndex = -1
      End If
      cboTipo.Locked = True
      MSFlex.BackColor = "&H00EFEFEF"
End If
   
   i = 0
   FlexCuotas
   
   'sSQL = "Select * from LogLocadoresCronograma where nRegLocador = " & nRegLoc & " "
   
   sSQL = "Select x.*, cConcepto = coalesce(y.cConsDescripcion,'') " & _
          "  from LogLocadoresCronograma x left join (select nConsValor,cConsDescripcion from Constante where nConsCod = 9133 and nConsCod <> nConsValor) y on x.nTipoCuota = y.nConsValor " & _
          " Where X.nRegLocador = " & nRegLoc & " "
          
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 1) = Format(rs!nNroCuota, "00")
         MSFlex.TextMatrix(i, 2) = rs!dFechaPago
         MSFlex.TextMatrix(i, 3) = rs!cConcepto
         MSFlex.TextMatrix(i, 4) = FNumero(rs!nMontoAPagar)
         rs.MoveNext
      Loop
   End If
   txtCuotas.Text = i:   txtCuotas.BackColor = "&H00EFEFEF":        txtCuotas.Locked = True
   fraVis.Visible = False
   fraReg.Visible = True
End If
End Sub

Private Sub cmdActualiza_Click()
Dim k As Integer
VerDatos
nFuncion = 2
cmdFunciones.Visible = True
cmdGrabar.Visible = True
txtNroContrato.BackColor = "&H80000005":  txtNroContrato.Locked = False
cboTipo.BackColor = "&H80000005":         cboTipo.Locked = False
txtFecIni.BackColor = "&H80000005":       txtFecIni.Locked = False
txtFecFin.BackColor = "&H80000005":       txtFecFin.Locked = False
txtFecIni.BackColor = "&H80000005":       txtFecIni.Locked = False
txtFecFin.BackColor = "&H80000005":       txtFecFin.Locked = False
cboFuncion.BackColor = "&H80000005":      cboFuncion.Locked = False
'txtMonto.BackColor = "&H80000005":        txtMonto.Locked = False
cboAge.BackColor = "&H80000005":          cboAge.Locked = False
cboArea.BackColor = "&H80000005":         cboArea.Locked = False
cboMoneda.BackColor = "&H80000005":       cboMoneda.Locked = False
End Sub

Private Sub cmdAgrega_Click()
Dim i As Integer, nCod As Integer, cDesc As String

nFuncion = 1

sSQL = "select nConsValor,cConsDescripcion from Constante where nConsCod = 9133 and nConsCod<>nConsValor"
frmLogSelector.Consulta sSQL, "Seleccione concepto"
i = MSFlex.Rows - 1
If Len(frmLogSelector.vpCodigo) > 0 Then
   i = i + 1
   InsRow MSFlex, i
   MSFlex.TextMatrix(i, 0) = frmLogSelector.vpCodigo   'CODIGO DE COncepto
   MSFlex.TextMatrix(i, 1) = "00"
   MSFlex.TextMatrix(i, 2) = txtFecIni.Text
   MSFlex.TextMatrix(i, 3) = frmLogSelector.vpDescripcion
   MSFlex.TextMatrix(i, 4) = "0.00"
End If
End Sub


Private Sub cmdFunciones_Click()
Dim k As Integer
frmMntConstantes.Inicio 9132
If frmMntConstantes.vpHaGrabado Then
   Call CargaComboConstante(9132, cboFuncion)
End If
End Sub

Private Sub cmdGrabar_Click()
Dim oConn As New DConecta, nReg As Integer
Dim cPersCod As String, i As Integer, n As Integer
Dim cNroContrato As String, nTipoContrato As Integer, dFecIni As String, dFecFin As String
Dim nFuncionCod As Integer, nMontoContrato As Currency, bPagoComision As Integer, nEstado As Integer
Dim nMontoCuota As Currency, nCuota As Integer, dFechaVen As String, nCAP As Integer
Dim rs As New ADODB.Recordset, nTipMon As Integer
Dim nRegLoc As Integer

If Len(Trim(txtPersCod.Text)) < 13 Then
   MsgBox "Debe indicar una persona válida..." + Space(10), vbInformation
   Exit Sub
End If

If Len(Trim(txtNroContrato.Text)) = 0 Then
   MsgBox "Debe indicar un número de contrato válido..." + Space(10), vbInformation
   Exit Sub
End If

nRegLoc = 0
If nFuncion = 1 Then
   If oConn.AbreConexion Then
      Set rs = oConn.CargaRecordSet("Select dFechaInicio,dFechaTermino,nEstado from LogLocadores where cPersCod = '" & txtPersCod.Text & "' ")
      If Not rs.EOF Then
         MsgBox "La persona tiene un contrato vigente desde " + CStr(rs!dFechaInicio) + " hasta el " + CStr(rs!dFechaTermino) + Space(10), vbInformation, "Aviso"
         Exit Sub
      End If
   End If
Else
   nRegLoc = MSLista.TextMatrix(MSLista.row, 0)
End If

cPersCod = txtPersCod.Text
cNroContrato = txtNroContrato.Text

If cboTipo.ListIndex < 0 Then
   MsgBox "Debe indicar el Tipo de Contrato..." + Space(10), vbInformation, "Verifique"
   cboTipo.SetFocus
   Exit Sub
End If

nTipoContrato = cboTipo.ItemData(cboTipo.ListIndex)
dFecIni = Format(txtFecIni, "YYYYMMDD")
dFecFin = Format(txtFecFin, "YYYYMMDD")
nFuncionCod = cboFuncion.ItemData(cboFuncion.ListIndex)
nMontoContrato = CCur(VNumero(txtMonto.Text))
cAgeCod = Format(cboAge.ItemData(cboAge.ListIndex), "00")
cAreaCod = Format(cboArea.ItemData(cboArea.ListIndex), "000")
nCAP = IIf(chkCAP.value = 0, 0, 1)
bPagoComision = 0
nEstado = 1
Select Case cboMoneda.ListIndex
    Case -1
         MsgBox "Debe indicar la moneda del Contrato..." + Space(10), vbInformation, "Revisar"
         Exit Sub
    Case 0
         nTipMon = 1
    Case 1
         nTipMon = 2
End Select


n = MSFlex.Rows - 1

If MsgBox("¿ Esta seguro de grabar los datos ?" + Space(10), vbQuestion + vbYesNo, "Confirme operacion") = vbYes Then
   
   If oConn.AbreConexion Then
      
      If nFuncion = 1 Then
      
         sSQL = "INSERT INTO LogLocadores (cPersCod , cNroContrato, nTipoContrato, dFechaInicio, dFechaTermino, cAreaCod, cAgeCod, nFuncionCod, nMoneda ,nMontoContrato, bPagoComision, nCap, nEstado) " & _
                " VALUES ('" & cPersCod & "','" & cNroContrato & "'," & nTipoContrato & ",'" & dFecIni & "','" & dFecFin & "','" & cAreaCod & "','" & cAgeCod & "'," & nFuncionCod & "," & nTipMon & "," & nMontoContrato & "," & bPagoComision & "," & nCAP & "," & nEstado & ") "
         oConn.ConexionActiva.Execute sSQL
   
         nReg = UltimaSecuenciaIdentidad("LogLocadores")
      
         For i = 1 To n
          nCuota = MSFlex.TextMatrix(i, 1)
          'MSFlex.TextMatrix(i, 2) = "30/09/2005"
          dFechaVen = Format(MSFlex.TextMatrix(i, 2), "YYYYMMDD")
          nMontoCuota = MSFlex.TextMatrix(i, 3)
          
          sSQL = "INSERT INTO LogLocadoresCronograma (nRegLocador, nNroCuota , dFechaPago, nMoneda ,nMontoAPagar, nEstadoCuota) " & _
                 " VALUES (" & nReg & ",'" & nCuota & "','" & dFechaVen & "'," & nTipMon & "," & nMontoCuota & ",0) "
          oConn.ConexionActiva.Execute sSQL
         Next
      End If
      
      If nFuncion = 2 Then
            
         sSQL = "UPDATE LogLocadores SET cNroContrato = '" & cNroContrato & "', nTipoContrato = " & nTipoContrato & ", dFechaInicio = '" & dFecIni & "', " & _
                " dFechaTermino = '" & dFecFin & "', cAreaCod = '" & cAreaCod & "', cAgeCod = '" & cAgeCod & "', nFuncionCod = " & nFuncionCod & ", " & _
                " nMoneda = " & nTipMon & ",nMontoContrato=" & nMontoContrato & ", bPagoComision=" & bPagoComision & ", nCap = " & nCAP & ", nEstado=" & nEstado & " " & _
                " WHERE nRegLocador = '" & nRegLoc & "' "
         oConn.ConexionActiva.Execute sSQL
      
      End If
      
   End If
   HaGrabado = True
   MsgBox "Se han grabado los datos correctamente" + Space(10), vbInformation
   fraVis.Visible = True
   fraReg.Visible = False
End If
End Sub

Private Sub cmdQuitar_Click()
Dim nRegLoc As Integer
Dim oConn As New DConecta

nRegLoc = MSLista.TextMatrix(MSLista.row, 0)

If Len(MSLista.TextMatrix(MSLista.row, 1)) = 0 Then Exit Sub

If MsgBox("¿ Está seguro de quitar la Persona indicada ?" + Space(10), vbQuestion + vbYesNo, "Confirme eliminación") = vbYes Then
   If oConn.AbreConexion Then
      sSQL = "UPDATE LogLocadores SET nEstado=0 where nRegLocador = " & nRegLoc & " "
      oConn.ConexionActiva.Execute sSQL
      oConn.CierraConexion
      ListaLocadores
   End If
End If
End Sub

Private Sub cmdReportes_Click()
frmLogLocadorReportes.Show 1
End Sub

Private Sub Form_Load()
'Me.Caption = "Control de Locadores"
CentraForm Me
CargaCombos
nEditable = 1
HaCargado = False
txtFechaIni = "01/01/" + CStr(Year(Date))
txtFechaFin = Date
'ListaLocadores
End Sub

Sub ListaLocadores()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim i As Integer

FormaFlex
i = 0
If oConn.AbreConexion Then

   sSQL = "select x.nRegLocador, x.cPersCod, p.cPersNombre,cNroContrato,dFechaInicio,dFechaTermino, " & _
          " aa.cAreaDescripcion as cArea, a.cAgeDescripcion as cAgencia " & _
          " from LogLocadores x inner join Persona p on x.cPersCod=p.cPersCod " & _
          "            inner join Agencias a on x.cAgeCod = a.cAgeCod " & _
          "            inner join Areas aa on x.cAreaCod = aa.cAreaCod  " & _
          " Where x.nEstado = 1 and " & _
          "      (x.dFechaInicio between '" & Format(txtFechaIni, "YYYYMMDD") & "' and '" & Format(txtFechaFin, "YYYYMMDD") & "' OR  " & _
          "       x.dFechaTermino between '" & Format(txtFechaIni, "YYYYMMDD") & "' and '" & Format(txtFechaFin, "YYYYMMDD") & "' )  "
          
' (x.dFechaTermino >= '" & Format(txtFechaIni, "YYYYMMDD") & "' and x.dFechaTermino <= '" & Format(txtFechaFin, "YYYYMMDD") & "')
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSLista, i
         MSLista.TextMatrix(i, 0) = rs!nRegLocador
         MSLista.TextMatrix(i, 1) = rs!cPersCod
         MSLista.TextMatrix(i, 2) = rs!cPersNombre
         MSLista.TextMatrix(i, 3) = rs!cNroContrato
         MSLista.TextMatrix(i, 4) = rs!dFechaInicio
         MSLista.TextMatrix(i, 5) = rs!dFechaTermino
         MSLista.TextMatrix(i, 6) = rs!cArea
         MSLista.TextMatrix(i, 7) = rs!cAgencia
         rs.MoveNext
      Loop
   End If
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdCancelar_Click()
fraVis.Visible = True
fraReg.Visible = False
End Sub

Private Sub cmdPersona_Click()
Dim X As UPersona
Set X = frmBuscaPersona.Inicio

If X Is Nothing Then
    Exit Sub
End If

If Len(Trim(X.sPersNombre)) > 0 Then
   txtPersona.Text = X.sPersNombre
   txtPersCod = X.sPersCod
   txtNroContrato.SetFocus
End If

'frmBuscaPersona.Show 1
'If frmBuscaPersona. Then
   'txtPersCod.Text = frmBuscaPersona.vpPersCod
   'txtPersona.Text = frmBuscaPersona.vpPersNom
'End If
End Sub

Private Sub cmdAgregar_Click()
If Not HaCargado Then
   CargaCombosFuncion
End If

cmdFunciones.Enabled = True
cmdAgrega.Enabled = True
cmdGrabar.Enabled = True
cmdPersona.Visible = True
cmdFunciones.Visible = True
cmdAgrega.Visible = True
cmdGrabar.Visible = True

cboMoneda.ListIndex = 0
txtPersCod.Text = ""
txtNroContrato.Text = "": txtNroContrato.BackColor = "&H80000005":  txtNroContrato.Locked = False
                          cboTipo.BackColor = "&H80000005":         cboTipo.Locked = False
txtFecIni.Text = Date:    txtFecIni.BackColor = "&H80000005":       txtFecIni.Locked = False
txtFecFin.Text = Date:    txtFecFin.BackColor = "&H80000005":       txtFecFin.Locked = False
txtMonto.Text = "":       txtMonto.BackColor = "&H80000005":        txtMonto.Locked = False
txtCuotas.Text = "":      txtCuotas.BackColor = "&H80000005":       txtCuotas.Locked = False

HaGrabado = False
cboFuncion.BackColor = "&H80000005":      cboFuncion.Locked = False
cboAge.BackColor = "&H80000005":          cboAge.Locked = False
cboArea.BackColor = "&H80000005":         cboArea.Locked = False
cboMoneda.BackColor = "&H80000005":       cboMoneda.Locked = False
MSFlex.BackColor = "&H80000005"

FlexCuotas

fraVis.Visible = False
fraReg.Visible = True
End Sub

Sub CargaCombos()
Call CargaComboBoxConstante(cboTipo, 9131)
Call CargaComboBoxConstante(cboFuncion, 9132)
End Sub

Sub CargaCombosFuncion()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta

If oConn.AbreConexion Then
   cboAge.Clear
   sSQL = "Select cAgeCod,cAgeDescripcion from Agencias where nEstado = 1"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         cboAge.AddItem rs!cAgeDescripcion
         cboAge.ItemData(cboAge.ListCount - 1) = rs!cAgeCod
         rs.MoveNext
      Loop
      cboAge.ListIndex = 0
   End If
   
   cboArea.Clear
   sSQL = "select cAreaCod, cAreaDescripcion from Areas order by cAreaDescripcion"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         cboArea.AddItem rs!cAreaDescripcion
         cboArea.ItemData(cboArea.ListCount - 1) = rs!cAreaCod
         rs.MoveNext
      Loop
      cboArea.ListIndex = 0
   End If
End If
End Sub


Sub cboAge_Click()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta

'cAgeCod = Format(cboAge.ItemData(cboAge.ListIndex), "00")
'cboArea.Clear
'If oConn.AbreConexion Then
'   sSQL = "select a.cAreaCod, x.cAreaDescripcion  from AreaAgencia a inner join Areas x on a.cAreaCod = x.cAreaCod where a.cAgeCod = '" & cAgeCod & "' order by x.cAreaDescripcion"
'   Set Rs = oConn.CargaRecordSet(sSQL)
'   If Not Rs.EOF Then
'      Do While Not Rs.EOF
'         cboArea.AddItem Rs!cAreaDescripcion
'         cboArea.ItemData(cboArea.ListCount - 1) = Rs!cAreaCod
'         Rs.MoveNext
'      Loop
'      cboArea.ListIndex = 0
'   End If
'End If
End Sub

Sub FormaFlex()
MSLista.Clear
MSLista.Rows = 2
MSLista.RowHeight(0) = 320
MSLista.RowHeight(1) = 8
MSLista.ColWidth(0) = 0
MSLista.ColWidth(1) = 1100: MSLista.TextMatrix(0, 1) = "Código"
MSLista.ColWidth(2) = 3500: MSLista.TextMatrix(0, 2) = "Locador"
MSLista.ColWidth(3) = 2200: MSLista.TextMatrix(0, 3) = " Nº Contrato"
MSLista.ColWidth(4) = 1000: MSLista.TextMatrix(0, 4) = "Inicio"
MSLista.ColWidth(5) = 1000: MSLista.TextMatrix(0, 5) = "Término"
MSLista.ColWidth(6) = 4300: MSLista.TextMatrix(0, 6) = "Area"
MSLista.ColWidth(7) = 3000: MSLista.TextMatrix(0, 7) = "Agencia"
MSLista.ColWidth(8) = 0
MSLista.ColWidth(9) = 0
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 400:  MSFlex.TextMatrix(0, 1) = " Nº":  MSFlex.ColAlignment(1) = 4
MSFlex.ColWidth(2) = 1100: MSFlex.TextMatrix(0, 2) = "Fecha Pago"
MSFlex.ColWidth(3) = 1200: MSFlex.TextMatrix(0, 3) = "  Monto Cuota"
End Sub


Private Sub txtFechaFin_Change()
If IsDate(txtFechaIni) And IsDate(txtFechaFin) And Len(Trim(txtFechaIni)) = 10 And Len(Trim(txtFechaFin)) = 10 Then
   ListaLocadores
End If
End Sub

Private Sub txtFechaIni_Change()
If IsDate(txtFechaIni) And IsDate(txtFechaFin) And Len(Trim(txtFechaIni)) = 10 And Len(Trim(txtFechaFin)) = 10 Then
   ListaLocadores
End If
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
   'MSFlex.SetFocus
End If
End Sub

Private Sub txtFecIni_GotFocus()
SelTexto txtFecIni
End Sub

Private Sub txtFecFin_GotFocus()
SelTexto txtFecFin
End Sub

Private Sub txtNroContrato_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   cboTipo.SetFocus
End If
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFecIni.SetFocus
End If
End Sub

Private Sub TxtFecIni_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigFecha(txtFecIni, KeyAscii)
If nKeyAscii = 13 Then
   txtFecFin.SetFocus
End If
End Sub

Private Sub TxtFecFin_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigFecha(txtFecFin, KeyAscii)
If nKeyAscii = 13 Then
   txtMonto.SetFocus
End If
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If VNumero(txtMonto) > 0 And VNumero(txtCuotas) > 0 Then
      GeneraListaCrono
   End If
   txtMonto = FNumero(txtMonto)
   txtCuotas.SetFocus
End If
End Sub

Private Sub txtCuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If VNumero(txtMonto) > 0 And VNumero(txtCuotas) > 0 Then
      GeneraListaCrono
   End If
   txtMonto.SetFocus
End If
End Sub

Private Sub cboAge_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboArea.SetFocus
End If
End Sub

Private Sub cboArea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboFuncion.SetFocus
End If
End Sub

Private Sub cboFuncion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdGrabar.SetFocus
End If
End Sub

Sub GeneraListaCrono()
Dim nMonto As Currency, nCuotas As Integer
Dim nMontoCuota As Currency, i As Integer, nDia As Integer
Dim cFecha As String, nMes As Integer, nAnio As Integer
Dim nSuma As Currency
Dim dFecha As Date

FlexCuotas
nMonto = CCur(VNumero(txtMonto.Text))

If Len(Trim(txtFecIni.Text)) = 10 Then
   dFecha = CDate(txtFecIni.Text)
Else
   Exit Sub
End If

nCuotas = CInt(Val(txtCuotas))
nMes = Month(dFecha)
nAnio = Year(dFecha)
nSuma = 0
nMontoCuota = Round(nMonto / nCuotas, 2)
For i = 1 To nCuotas
    InsRow MSFlex, i
    nDia = Day(DateSerial(nAnio, nMes + 1, 0))
    MSFlex.TextMatrix(i, 0) = 1                     'CONCEPTO DE PAGO
    MSFlex.TextMatrix(i, 1) = Format(i, "00")
    MSFlex.TextMatrix(i, 2) = Format(nDia, "00") + "/" + Format(nMes, "00") + "/" + CStr(nAnio)
    MSFlex.TextMatrix(i, 3) = CStr(i) + "º ARMADA"
    MSFlex.TextMatrix(i, 4) = FNumero(nMontoCuota)
    If nMes >= 12 Then
       nMes = 1
       nAnio = nAnio + 1
    Else
       nMes = nMes + 1
    End If
    nSuma = nSuma + nMontoCuota
Next
i = IIf(nSuma > nMonto, 1, MSFlex.Rows - 1)
MSFlex.TextMatrix(i, 4) = FNumero(VNumero(MSFlex.TextMatrix(i, 4)) + (nMonto - nSuma))
nSuma = 0
For i = 1 To MSFlex.Rows - 1
    nSuma = nSuma + VNumero(MSFlex.TextMatrix(i, 4))
Next
'i = MSFlex.Rows - 1
'txtFecFin = MSFlex.TextMatrix(i, 2)
txtTotal.Text = FNumero(nSuma)
nEditable = True
End Sub

Sub FlexCuotas()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 400:  MSFlex.TextMatrix(0, 1) = " Nº":  MSFlex.ColAlignment(1) = 4
MSFlex.ColWidth(2) = 1000: MSFlex.TextMatrix(0, 2) = "Fecha Pago":  MSFlex.ColAlignment(2) = 4
MSFlex.ColWidth(3) = 1900: MSFlex.TextMatrix(0, 3) = "  Concepto":  MSFlex.ColAlignment(3) = 1
MSFlex.ColWidth(4) = 1200: MSFlex.TextMatrix(0, 4) = "  Monto Cuota"
End Sub

'*********************************************************************
'PROCEDIMIENTOS DEL FLEX
'*********************************************************************

'Private Sub MSFlex_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyDelete And MSFlex.Col = 3 Then
'   MSFlex.TextMatrix(MSFlex.Row, 3) = ""
'End If
'End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
If MSFlex.Col >= 2 And nEditable Then
   txtEdit = MSFlex
   EditaFlex MSFlex, txtEdit, KeyAscii
End If
End Sub

Sub EditaFlex(MSFlex As Control, Edt As Control, KeyAscii As Integer)
If InStr("0123456789/", Chr(KeyAscii)) Then
Select Case KeyAscii
    Case 0 To 32
         Edt = MSFlex
         Edt.SelStart = 1000
    Case Else
         Edt = Chr(KeyAscii)
         Edt.SelStart = 1
End Select
Edt.Move MSFlex.Left + MSFlex.CellLeft - 15, MSFlex.Top + MSFlex.CellTop - 15, _
         MSFlex.CellWidth, MSFlex.CellHeight
Edt.Visible = True
Edt.SetFocus
End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
If MSFlex.Col = 3 Then
   KeyAscii = DigNumDec(txtEdit, KeyAscii)
End If
If MSFlex.Col = 2 Then
   KeyAscii = DigFecha(txtEdit, KeyAscii)
End If
If KeyAscii = Asc(vbCr) Then
   KeyAscii = 0
   txtEdit = FNumero(txtEdit)
End If
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode MSFlex, txtEdit, KeyCode, Shift
End Sub

Sub EditKeyCode(MSFlex As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27
         Edt.Visible = False
         MSFlex.SetFocus
    Case 13
         MSFlex.SetFocus
    Case 37                     'Izquierda
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col > 1 Then
            MSFlex.Col = MSFlex.Col - 1
         End If
    Case 39                     'Derecha
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col < MSFlex.Cols - 1 Then
            MSFlex.Col = MSFlex.Col + 1
         End If
    Case 38
         MSFlex.SetFocus
         DoEvents
         If MSFlex.row > MSFlex.FixedRows + 1 Then
            MSFlex.row = MSFlex.row - 1
         End If
    Case 40
         MSFlex.SetFocus
         DoEvents
         'If MSFlex.Row < MSFlex.FixedRows - 1 Then
         If MSFlex.row < MSFlex.Rows - 1 Then
            MSFlex.row = MSFlex.row + 1
         End If
End Select
End Sub

Private Sub MSFlex_GotFocus()
If txtEdit.Visible = False Then Exit Sub
If MSFlex.Col = 2 Then
   MSFlex = txtEdit
Else
   MSFlex = FNumero(txtEdit)
End If
txtEdit.Visible = False
VerificaSuma
End Sub

Private Sub MSFlex_LeaveCell()
If txtEdit.Visible = False Then Exit Sub
If MSFlex.Col = 2 Then
   MSFlex = txtEdit
Else
   MSFlex = FNumero(txtEdit)
End If
txtEdit.Visible = False
VerificaSuma
End Sub

Sub VerificaSuma()
Dim i As Integer, n As Integer
Dim nSuma As Currency
n = MSFlex.Rows - 1
nSuma = 0
For i = 1 To n
    If MSFlex.TextMatrix(i, 0) = 1 Then
       nSuma = nSuma + VNumero(MSFlex.TextMatrix(i, 4))
    End If
Next
If nSuma <> VNumero(txtMonto) Then
   MsgBox "Monto de Contrato =  " + txtMonto.Text + Space(10) + vbCrLf + "Suma de Cuotas ...=  " + FNumero(nSuma) + Space(10), vbInformation, "Verificación de montos"
End If
End Sub


