VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMetasAgencias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Metas de Agencias"
   ClientHeight    =   6630
   ClientLeft      =   4875
   ClientTop       =   3285
   ClientWidth     =   5895
   Icon            =   "frmMetasAgencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   5895
   Begin VB.Frame Frame3 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   5760
      Width           =   5655
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Metas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   5655
      Begin SICMACT.FlexEdit fgMetas 
         Height          =   2895
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5106
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-nTipoCred-Tipo Credito-Monto"
         EncabezadosAnchos=   "400-0-3300-1500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         ColumnasAEditar =   "X-X-X-3"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-R"
         FormatosEdit    =   "0-0-0-4"
         AvanceCeldas    =   1
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         Enabled         =   0   'False
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
         CellBackColor   =   -2147483624
      End
      Begin VB.Label lblFechaMeta 
         Caption         =   "__/__/____"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label lblSaldoTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Saldo Total Esperado al"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label lblSaldoMeta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo de Meta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H8000000D&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox cmbMes 
         Height          =   315
         ItemData        =   "frmMetasAgencias.frx":030A
         Left            =   1080
         List            =   "frmMetasAgencias.frx":0338
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   4200
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbAnio 
         Height          =   315
         ItemData        =   "frmMetasAgencias.frx":03A9
         Left            =   2640
         List            =   "frmMetasAgencias.frx":03AB
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin MSDataListLib.DataCombo dcAgencia 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblSaldoAnt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblFechaAnt 
         Caption         =   "__/__/____"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Saldo respecto al :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   630
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMetasAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbModificar As Boolean 'Modificar=1 ; No Modificar=0

Private Sub cmbAnio_Click()
    limpiar
End Sub
Private Sub cmbMes_Click()
    limpiar
End Sub

Private Sub cmdBuscar_Click()
     If dcAgencia.BoundText <> "0" And Me.cmbAnio.ListIndex > 0 Then
        fgMetas.Enabled = False
        cargarSaldoCarteraAgencia
        CargarMetas
        mostrarFechas
    ElseIf Me.cmbAnio.ListIndex <= 0 Then
        MsgBox "Seleccione un Año Valido", vbInformation, "Aviso!"
        Me.cmbAnio.SetFocus
    ElseIf dcAgencia.BoundText = "0" Then
        MsgBox "Seleccione una Agencia", vbInformation, "Aviso!"
        Me.dcAgencia.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
    Me.fgMetas.Enabled = False
    MostrarBoton True, True, False, False
    If CDbl(lblSaldoMeta) > 0 Then
        HabilitarBoton False, True, False, False
    Else
        HabilitarBoton True, False, False, False
    End If
End Sub

Private Sub cmdGrabar_Click()
    If Not valida Then
        Exit Sub
    End If
    
    If MsgBox("Se van a Guardar los Datos,Desea Continuar?", vbYesNo, "Guardar") = vbYes Then
        Dim sMovNro As String
        Dim oNCred As New COMNCredito.NCOMCredito
        Dim clsMov As New COMNContabilidad.NCOMContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
        If oNCred.guardarMetasAgencias(Me.fgMetas.GetRsNew, dcAgencia.BoundText, Me.cmbAnio.Text, Format(Me.cmbMes.ListIndex, "00"), Me.lblSaldoAnt, Me.lblSaldoTotal, sMovNro, lbModificar) Then
            lbModificar = False
            MsgBox "Se han Guardado los Datos con Exito!", vbInformation, "Datos Guardado"
            MostrarBoton True, True, False, False
            HabilitarBoton False, False, False, False
            limpiar
            Me.cmbAnio.ListIndex = 0
            Me.dcAgencia.BoundText = "0"
            Me.lblFechaAnt = ""
        Else
            MsgBox "No se Pudieron Guardar los Datos!", vbInformation, "Error al Guardar"
        End If
        
        
    End If
End Sub
Private Sub limpiar()

    Call LimpiaFlex(Me.fgMetas)
    Me.lblSaldoAnt = "0.00"
    Me.lblSaldoMeta = "0.00"
    Me.lblSaldoTotal = "0.00"
    
    Me.lblFechaAnt = "__/__/____"
    Me.lblFechaMeta = "__/__/____"
    
End Sub
Private Function valida() As Boolean
    valida = True
    If CDbl(Me.lblSaldoMeta) <= 0 Then
        MsgBox "Debe Ingresar las Metas", vbInformation, "Aviso!"
        Me.fgMetas.SetFocus
        valida = False
        Exit Function
    End If
'    If CDbl(Me.lblSaldoAnt) <= 0 Then
'        MsgBox "El saldo del Año Anterior es 0,Verifique", vbInformation, "Aviso!"
'        Me.cmbAnio.SetFocus
'        valida = False
'        Exit Function
'    End If

End Function
Private Sub cmdModificar_Click()
    fgMetas.Enabled = True
    lbModificar = True
    MostrarBoton False, False, True, True
    HabilitarBoton False, False, True, True
    Me.fgMetas.Row = 1
    Me.fgMetas.Col = 3
    Me.fgMetas.SetFocus
End Sub

Private Sub cmdNuevo_Click()
    fgMetas.Enabled = True
    Me.fgMetas.Row = 1
    Me.fgMetas.Col = 3
    Me.fgMetas.SetFocus
    MostrarBoton False, False, True, True
    HabilitarBoton False, False, True, True
End Sub

Private Sub cmdsalir_Click()
 Unload Me
End Sub

Private Sub dcAgencia_Change()
   limpiar
End Sub
Private Sub fgMetas_RowColChange()
    Me.lblSaldoMeta = Format(fgMetas.SumaRow(3), "##,##0.00")
    Me.lblSaldoTotal = Format(CDbl(Me.lblSaldoMeta) + CDbl(Me.lblSaldoAnt), "##,##0.00")
End Sub

Private Sub Form_Load()
    lbModificar = False
    cargarAnio
    CargarAgencias
'    CargarMetas
'    cargarSaldoCarteraAgencia
End Sub
Private Sub CargarAgencias()
    Dim rsAgencia As New adodb.Recordset
    Dim objCOMNCredito As COMNCredito.NCOMBPPR
    Set objCOMNCredito = New COMNCredito.NCOMBPPR
    Set rsAgencia.DataSource = objCOMNCredito.getCargarAgencias
    dcAgencia.BoundColumn = "cAgeCod"
    dcAgencia.DataField = "cAgeCod"
    Set dcAgencia.RowSource = rsAgencia
    dcAgencia.ListField = "cAgeDescripcion"
    dcAgencia.BoundText = "01"
End Sub
Private Sub cargarAnio()
    Dim i As Integer
    Dim indice As Integer
    indice = 1
    cmbAnio.AddItem "--Año--", 0
    For i = 2011 To val(Year(gdFecSis))
        cmbAnio.AddItem i
        cmbAnio.ItemData(indice) = i
        indice = indice + 1
    Next
    cmbAnio.ListIndex = cmbAnio.ListCount - 1
    cmbMes.ListIndex = Month(gdFecSis)
    mostrarFechas
    
End Sub
Private Sub mostrarFechas()

    Dim dFechaTmp As Date
    dFechaTmp = CDate("01/" + Format(Me.cmbMes.ListIndex, "00") + "/" + Me.cmbAnio.Text)
    Me.lblFechaAnt = DateAdd("d", -1, dFechaTmp)
    Me.lblFechaMeta = DateAdd("d", -1, DateAdd("M", 1, dFechaTmp))
End Sub
Private Sub CargarMetas()
    Dim oMetas As New COMNCredito.NCOMCredito
    Dim rs As New Recordset
    
    Set rs = oMetas.obtenerMetasAgenciasDet(dcAgencia.BoundText, cmbAnio.Text, Format(cmbMes.ListIndex, "00"))
    If Not (rs.EOF And rs.BOF) Then
        Me.fgMetas.rsFlex = rs.DataSource
    End If
    Me.lblSaldoMeta = Format(fgMetas.SumaRow(3), "##,##0.00")
    Me.lblSaldoTotal = Format(CDbl(Me.lblSaldoMeta) + CDbl(Me.lblSaldoAnt), "##,##0.00")
    MostrarBoton True, True, True, True
    If CDbl(lblSaldoMeta) > 0 Then
        HabilitarBoton False, True, False, False
    Else
        HabilitarBoton True, False, False, False
    End If
End Sub
Private Sub cargarSaldoCarteraAgencia()
    Dim oMetas As New COMNCredito.NCOMCredito
    Dim rs As New Recordset
    
    Dim clsTC As COMDConstSistema.NCOMTipoCambio
    Dim nTC As Double
    
    Me.MousePointer = vbArrowHourglass
    Set clsTC = New COMDConstSistema.NCOMTipoCambio
    nTC = clsTC.EmiteTipoCambio("01/" + Format(cmbMes.ListIndex, "00") + "/" + Me.cmbAnio.Text, TCFijoDia)

    mostrarFechas
    If nTC <> 0 Then
        Me.lblSaldoAnt = Format(oMetas.obtenerSaldoCarteraAgencia(dcAgencia.BoundText, Me.lblFechaAnt, nTC), "##,##0.00")
    Else
        Me.lblSaldoAnt = "0.00"
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub HabilitarBoton(pbNuevo As Boolean, pbModificar As Boolean, pbGrabar As Boolean, pbCancelar As Boolean)
    Me.cmdNuevo.Enabled = pbNuevo
    Me.cmdModificar.Enabled = pbModificar
    Me.cmdGrabar.Enabled = pbGrabar
    Me.cmdCancelar.Enabled = pbCancelar
    
End Sub

Private Sub MostrarBoton(pbNuevo As Boolean, pbModificar As Boolean, pbGrabar As Boolean, pbCancelar As Boolean)
    Me.cmdNuevo.Visible = pbNuevo
    Me.cmdModificar.Visible = pbModificar
    Me.cmdGrabar.Visible = pbGrabar
    Me.cmdCancelar.Visible = pbCancelar
End Sub

