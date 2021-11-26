VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPolizaSeguPatriListado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Polizas de Seguros Patrimoniales"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
   Icon            =   "frmPolizaSeguPatriListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Búsqueda"
      Height          =   2415
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   8190
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   315
         Left            =   480
         TabIndex        =   13
         Top             =   1800
         Width           =   2340
      End
      Begin VB.CommandButton cmdbuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4800
         TabIndex        =   12
         ToolTipText     =   "Buscar Credito"
         Top             =   1920
         Width           =   1200
      End
      Begin VB.CommandButton CmdBuscaCont 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7650
         TabIndex        =   4
         Top             =   600
         Width           =   390
      End
      Begin VB.OptionButton optFechas 
         Caption         =   "Por Fecha de Registro"
         Height          =   315
         Left            =   450
         TabIndex        =   2
         Top             =   1050
         Width           =   2340
      End
      Begin VB.OptionButton optContrat 
         Caption         =   "Por Aseguradora"
         Height          =   240
         Left            =   450
         TabIndex        =   1
         Top             =   300
         Width           =   2040
      End
      Begin MSMask.MaskEdBox mskInicio 
         Height          =   345
         Left            =   2550
         TabIndex        =   8
         Top             =   1350
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFin 
         Height          =   345
         Left            =   4650
         TabIndex        =   9
         Top             =   1350
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Del:"
         Height          =   240
         Left            =   2100
         TabIndex        =   11
         Top             =   1425
         Width           =   315
      End
      Begin VB.Label Label6 
         Caption         =   "Al:"
         Height          =   240
         Left            =   4275
         TabIndex        =   10
         Top             =   1425
         Width           =   240
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Aseguradora:"
         Height          =   195
         Left            =   1500
         TabIndex        =   7
         Top             =   630
         Width           =   945
      End
      Begin VB.Label LblContPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   3855
         TabIndex        =   6
         Top             =   615
         Width           =   3750
      End
      Begin VB.Label LblContPersCod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   2490
         TabIndex        =   5
         Top             =   615
         Width           =   1350
      End
   End
   Begin Sicmact.FlexEdit FePolizas 
      Height          =   2790
      Left            =   105
      TabIndex        =   3
      Top             =   2505
      Width           =   8190
      _extentx        =   14446
      _extenty        =   4921
      cols0           =   7
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#-N° Poliza-Aseguradora-Tipo Poliza-F. Registro-cPersCod-nTipoPoliza"
      encabezadosanchos=   "400-1000-3400-2050-1200-0-0"
      font            =   "frmPolizaSeguPatriListado.frx":030A
      fontfixed       =   "frmPolizaSeguPatriListado.frx":0336
      columnasaeditar =   "X-X-X-X-X-X-X"
      listacontroles  =   "0-0-0-0-0-0-0"
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      encabezadosalineacion=   "C-C-L-L-L-C-L"
      formatosedit    =   "0-0-0-0-0-0-0"
      textarray0      =   "#"
      lbultimainstancia=   -1  'True
      lbpuntero       =   -1  'True
      colwidth0       =   405
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
End
Attribute VB_Name = "frmPolizaSeguPatriListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum TipoOper
    nBusqueda = 0
    nConsulta = 1
End Enum

Dim nTipoOperacion As TipoOper

Public sNumPoliza As String
Public nTipoPoliza As Integer
Public sPersCodContr As String
Dim nEstadoPoliza As Integer

Public Sub Inicio(ByVal pnTipoBusqueda As TipoOper, _
                    Optional ByVal pnEstadoPoliza As Integer = -1)
    nTipoOperacion = pnTipoBusqueda
    nEstadoPoliza = pnEstadoPoliza
    sNumPoliza = ""
    Me.Show 1
    
End Sub

Private Sub cmdBuscaCont_Click()

Dim oPers As UPersona
    Set oPers = frmBuscaPersona.Inicio
    If Not oPers Is Nothing Then
        LblContPersCod.Caption = oPers.sPersCod
        LblContPersNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
    cmdbuscar.SetFocus
End Sub

Private Sub cmdBuscar_Click()

Dim oPol As DOperacion

'Dim oPol As COMDCredito.DCOMPoliza
Dim rs As ADODB.Recordset
Set oPol = New DOperacion
If optContrat.value Then
    If Len(Trim(LblContPersCod.Caption)) = 0 Then
        MsgBox "Ingrese una aseguradora para la busqueda.", vbOKOnly, "Atención"
        Exit Sub
    Else
        Set rs = oPol.CargaPoliSeguPatri(0, LblContPersCod.Caption)
    End If
ElseIf Me.optFechas.value Then
    If Len(ValidaFecha(mskInicio.Text)) = 0 And Len(ValidaFecha(mskFin.Text)) = 0 Then
        Set rs = oPol.CargaPoliSeguPatri(1, , Format(CDate(mskInicio.Text), "yyyymmdd"), Format(CDate(mskFin.Text), "yyyymmdd"))
    Else
        MsgBox "Ingrese fecha correctas.", vbOKOnly, "Atención"
        Exit Sub
    End If
ElseIf Me.optTodos.value Then
    Set rs = oPol.CargaPoliSeguPatri(2)
End If

If rs.EOF Then MsgBox "No se encontraron registros", vbInformation, "Mensaje"

FePolizas.Clear
FePolizas.FormaCabecera
FePolizas.Rows = 2
FePolizas.rsFlex = rs
FePolizas.SetFocus
Set oPol = Nothing
End Sub


Private Sub FePolizas_DblClick()
'If nTipoOperacion = nBusqueda Then
    'sNumPoliza = FePolizas.TextMatrix(FePolizas.Row, 1)
    'sPersCodContr = FePolizas.TextMatrix(FePolizas.Row, 6)
    'Unload Me
'Else
    'sNumPoliza = FePolizas.TextMatrix(FePolizas.Row, 1)

'MAVM 20090924
If nTipoOperacion = 2 Or nTipoOperacion = 0 Then
    sNumPoliza = FePolizas.TextMatrix(FePolizas.Row, 1)
    nTipoPoliza = FePolizas.TextMatrix(FePolizas.Row, 6)
    sPersCodContr = FePolizas.TextMatrix(FePolizas.Row, 5)
    'sNumPoliza = FePolizas.TextMatrix(FePolizas.Row, 1)
End If

'If nTipoOperacion = 1 Then 'Mantenimiento
'    frmCredPolizaGarantia.CargarDatos Trim(FePolizas.TextMatrix(FePolizas.Row, 1)), Trim(FePolizas.TextMatrix(FePolizas.Row, 6))
'    frmCredPolizaGarantia.LlenaGrillas Trim(FePolizas.TextMatrix(FePolizas.Row, 6)), Trim(FePolizas.TextMatrix(FePolizas.Row, 1))
'    frmCredPolizaGarantia.lblNumPoliza.Caption = Trim(FePolizas.TextMatrix(FePolizas.Row, 1))
'End If


'If nTipoOperacion = 3 Then
'    Call frmCredPolizaConsulta.Inicio(FePolizas.TextMatrix(FePolizas.Row, 1))
'End If

Unload Me

End Sub

Private Sub FePolizas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call FePolizas_DblClick
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
End Sub

Private Sub mskFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdbuscar.SetFocus
End Sub

Private Sub mskInicio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then mskFin.SetFocus
End Sub

Private Sub optContrat_Click()
If optContrat.value Then
    CmdBuscaCont.Enabled = True
    mskInicio.Enabled = False
    mskInicio.Text = "__/__/____"
    mskFin.Enabled = False
    mskFin.Text = "__/__/____"
End If
End Sub

Private Sub optFechas_Click()
If optFechas.value Then
    CmdBuscaCont.Enabled = False
    mskInicio.Enabled = True
    mskInicio.Text = "__/__/____"
    mskFin.Enabled = True
    mskFin.Text = "__/__/____"
    
    Me.LblContPersCod.Caption = ""
    Me.LblContPersNombre.Caption = ""
    
End If
End Sub

