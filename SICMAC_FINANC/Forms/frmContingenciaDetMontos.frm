VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmContingenciaDetMontos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contingencia: Detalle de Monto"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "frmContingenciaDetMontos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin MSDataGridLib.DataGrid DGDetalleMontos 
         Height          =   2055
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   17
         RowDividerStyle =   6
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "cMonto"
            Caption         =   "Concepto Monto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cMoneda"
            Caption         =   "Moneda"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "nMonto"
            Caption         =   "Monto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "dFecha"
            Caption         =   "Fecha Reg."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            ScrollBars      =   2
            BeginProperty Column00 
               ColumnWidth     =   2594.835
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1425.26
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1904.882
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3240
         Width           =   1035
      End
      Begin VB.ComboBox cboTipoMonto 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3240
         Width           =   3435
      End
      Begin VB.PictureBox Picture1 
         Height          =   585
         Left            =   120
         ScaleHeight     =   525
         ScaleWidth      =   7515
         TabIndex        =   3
         Top             =   3960
         Width           =   7575
         Begin VB.CommandButton cmdSalir 
            Cancel          =   -1  'True
            Caption         =   "&Salir"
            CausesValidation=   0   'False
            Height          =   400
            Left            =   6360
            TabIndex        =   9
            Top             =   60
            Width           =   1100
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   400
            Left            =   5160
            TabIndex        =   8
            Top             =   60
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.CommandButton cmdModificar 
            Caption         =   "&Modificar"
            Height          =   400
            Left            =   1320
            TabIndex        =   7
            Top             =   75
            Width           =   1100
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   400
            Left            =   2520
            TabIndex        =   6
            Top             =   75
            Width           =   1100
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   400
            Left            =   120
            TabIndex        =   5
            Top             =   75
            Width           =   1100
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            CausesValidation=   0   'False
            Height          =   400
            Left            =   6360
            TabIndex        =   4
            Top             =   60
            Visible         =   0   'False
            Width           =   1100
         End
      End
      Begin Sicmact.EditMoney txtMonto 
         Height          =   315
         Left            =   5040
         TabIndex        =   16
         Top             =   3240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin Sicmact.EditMoney txtTipoCambio 
         Height          =   315
         Left            =   6480
         TabIndex        =   18
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox txtFecRegContable 
         Height          =   285
         Left            =   6360
         TabIndex        =   19
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   15794175
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha Reg. Contable:"
         Height          =   255
         Left            =   4680
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Cambio"
         Height          =   210
         Left            =   6360
         TabIndex        =   17
         Top             =   3000
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label2 
         Caption         =   "Monto"
         Height          =   210
         Left            =   5400
         TabIndex        =   14
         Top             =   3000
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda"
         Height          =   210
         Left            =   4200
         TabIndex        =   13
         Top             =   3000
         Width           =   660
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         Left            =   120
         Top             =   2760
         Width           =   7575
      End
      Begin VB.Label LblTipoUnstFinan 
         Caption         =   "Concepto Monto"
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Label lblNroRegistroP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1515
         TabIndex        =   2
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "Nº de Registro:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmContingenciaDetMontos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CROB20170731
Option Explicit
Dim oConting As DContingencia 'TORE20032018
Dim oContTM As DContingenciaTipoMonto 'CROB20170724
Dim oConM As DContingenciaMontos
Dim rs As ADODB.Recordset
Dim ControlEjecutar As Integer 'TORE20032018
Dim nTipoInicio As Integer
Dim nTipoMontoAnterior As Integer 'TORE20032018


Private Sub cmdSalir_Click()
     Unload Me
End Sub

Public Sub Consultar(ByVal psNumRegistro As String)
    nTipoInicio = 2 'Consultar Datos TORE20032018
    DGDetalleMontos.Height = 3255
    Call CargarTiposDeMontos
    Call CargarTipoDeMoneda
    Call AcitvaControles(False)
    CargarDatos (psNumRegistro)
    Me.Show 1
End Sub
'TORE20032018
Public Function ValidaDatos() As Boolean
        If cboTipoMonto.ListIndex = -1 Then
            MsgBox "Falta seleccionar el tipo de monto", vbInformation, "Aviso"
            cboTipoMonto.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If cboMoneda.ListIndex = -1 Then
            MsgBox "El Monto de Perdida es obligatorio, seleccione el tipo de moneda del monto de perdida.", vbInformation, "Aviso"
            cboMoneda.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If txtMonto.value = 0 Then
            MsgBox "Ingrese el monto.", vbInformation, "Aviso"
            txtMonto.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    ValidaDatos = True
End Function

Private Sub cboMoneda_Click()
    If Right(cboMoneda.Text, 1) = 2 Then
        Label3.Visible = True
        txtTipoCambio.Visible = True
        txtTipoCambio.Enabled = False
        txtTipoCambio.value = ObtenerTipoCambioFecha(txtFecRegContable.Text)
    Else
    Label3.Visible = False
    txtTipoCambio.Visible = False
    txtTipoCambio.value = 0
    End If
End Sub

Private Sub CargarDatos(ByVal psNumRegistro As String)
    Set oConting = New DContingencia
    Set rs = oConting.CargaContingenciasMontos(psNumRegistro)
    Set DGDetalleMontos.DataSource = rs
    Set oConting = Nothing
    DGDetalleMontos.Refresh
    lblNroRegistroP.Caption = psNumRegistro
    txtFecRegContable.Text = gdFecSis
End Sub

Private Sub CargarTiposDeMontos()
    Set oContTM = New DContingenciaTipoMonto
    Set rs = oContTM.ListarTipoMontoPasivoContingente
    Call CargaCombo(rs, cboTipoMonto, 0, 1)
    'cboTipoMonto.ListIndex = 0
    Set oContTM = Nothing
    Set rs = Nothing
End Sub

Private Sub CargarTipoDeMoneda()
    Set oContTM = New DContingenciaTipoMonto
    Set rs = oContTM.ListarTipoMoneda
    Call CargaCombo(rs, cboMoneda, 0, 1)
    'cboMoneda.ListIndex = 0
    Set oContTM = Nothing
    Set rs = Nothing
End Sub

Private Sub DGDetalleMontos_DblClick()
    If cmdModificar.Enabled Then
        Call cmdModificar_Click
    End If
End Sub

Private Sub DGDetalleMontos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdModificar.Enabled Then
            Call cmdModificar_Click
        End If
    End If
End Sub

Private Sub cmdEliminar_Click()
On Error GoTo ERRORcmdEliminar
    If MsgBox("Esta seguro de eliminar el monto " & rs!cMonto & ", Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set oConM = New DContingenciaMontos
        Set rs = oConM.EliminarMontoDetalle(lblNroRegistroP.Caption, Right(rs!cTipoMonto, 2))
        If rs!nResultado Then
            MsgBox "Monto eliminado", vbInformation, "Aviso"
        End If
        Set oConM = Nothing
        Call CargarDatos(lblNroRegistroP.Caption)
    End If
    DGDetalleMontos.SetFocus
    Exit Sub
ERRORcmdEliminar:
    MsgBox Err.Description, vbExclamation, "Aviso"
End Sub

Private Sub cmdModificar_Click()
    Call LimpiaControles
    Call AcitvaControles(True)
    nTipoMontoAnterior = 0
    
     If Not rs.EOF And Not rs.BOF Then
        cboTipoMonto.ListIndex = IndiceListaCombo(cboTipoMonto, Trim(Str(CInt(Trim(Right(rs!cTipoMonto, 15))))))
        cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, Trim(Str(CInt(Trim(Right(rs!nnMoneda, 15))))))
        txtMonto.Text = rs!nMonto
        nTipoMontoAnterior = rs!cTipoMontoPCID_A
        'txtTipoCambio.Text = rs!nTipoCamb
        DGDetalleMontos.Height = 2055
        cboTipoMonto.SetFocus
        ControlEjecutar = 2
    Else
    MsgBox "No existe concepto asignado a la contingencia", vbInformation, "Aviso"
    Call cmdNuevo_Click
    End If
    
   
    
End Sub

Private Sub cmdNuevo_Click()
    DGDetalleMontos.Height = 2055
    Call LimpiaControles
    Call AcitvaControles(True)
    txtMonto.SetFocus
    ControlEjecutar = 1
End Sub

Private Sub LimpiaControles()
    txtMonto.Text = ""
End Sub


Private Sub AcitvaControles(ByVal pbHabilita As Boolean)
    DGDetalleMontos.Enabled = Not pbHabilita
    cmdNuevo.Visible = IIf(pbHabilita, False, True)
    cmdModificar.Visible = IIf(pbHabilita, False, True)
    cmdEliminar.Visible = IIf(pbHabilita, False, True)
    cmdSalir.Visible = IIf(pbHabilita, False, True)
    cmdAceptar.Visible = pbHabilita
    cmdCancelar.Visible = pbHabilita
End Sub



Public Sub CargaCombo(ByVal prsCombo As ADODB.Recordset, ByVal CtrlCombo As ComboBox, ByVal pnFiel1 As Integer, ByVal pnFiel2 As Integer)
    CtrlCombo.Clear
    While Not rs.EOF
        CtrlCombo.AddItem prsCombo.Fields(pnFiel1) & space(100) & prsCombo.Fields(pnFiel2) 'CROB20170721
        rs.MoveNext
    Wend
End Sub


Private Sub cmdAceptar_Click()
    If ValidaDatos Then
        Set oConM = New DContingenciaMontos
        Select Case ControlEjecutar
            Case 1 'Nuevo
                If MsgBox("¿Está seguro de registrar el concepto?", vbInformation + vbYesNo, "Confirmar") = vbNo Then Exit Sub
                Set rs = oConM.NuevoMontoDetalle(lblNroRegistroP.Caption, Trim(Right(cboTipoMonto.List(cboTipoMonto.ListIndex), 25)), txtMonto.Text, Trim(Right(cboMoneda.List(cboMoneda.ListIndex), 25)), IIf(Trim(Right(cboMoneda.List(cboMoneda.ListIndex), 25)) = 2, ObtenerTipoCambioFecha(txtFecRegContable.Text), 0)) 'txtTipoCambio.Text)
                If rs!nResultados = 1 Then
                    MsgBox "El concepto fue registrado", vbInformation, "Aviso"
                    Set oConM = Nothing
                    Set rs = Nothing
                    Call CargarDatos(lblNroRegistroP.Caption)
                    Exit Sub
                ElseIf rs!nResultados = 2 Then
                    MsgBox "El concepto del monto ya se encuentra registrado", vbInformation, "Aviso"
                    Set oConM = Nothing
                    Set rs = Nothing
                    Exit Sub
                End If
            Case 2 'Modificar
                If MsgBox("¿Está seguro de guardar los cambios?", vbInformation + vbYesNo, "Confirmar") = vbNo Then Exit Sub
                If nTipoMontoAnterior = 0 Then
                MsgBox "No se reconocio el concepto anterior", vbInformation, "Aviso"
                Else
                Set rs = oConM.ActualizarMontoDetalle(lblNroRegistroP.Caption, Trim(Right(cboTipoMonto.List(cboTipoMonto.ListIndex), 25)), txtMonto.Text, Trim(Right(cboMoneda.List(cboMoneda.ListIndex), 25)), IIf(Trim(Right(cboMoneda.List(cboMoneda.ListIndex), 25)) = 2, ObtenerTipoCambioFecha(txtFecRegContable.Text), 0), nTipoMontoAnterior)
                If rs!nResultado = 1 Then
                    MsgBox "Concepto actualizado", vbInformation, "Aviso"
                    Set oConM = Nothing
                    Set rs = Nothing
                    'No mostrar
                    Label3.Visible = True
                    txtTipoCambio.Visible = True
                    txtTipoCambio.value = 0
                    Call CargarDatos(lblNroRegistroP.Caption)
                    
                End If
                End If
        End Select
        DGDetalleMontos.Height = 3255
        Call LimpiaControles
        Call AcitvaControles(False)
        DGDetalleMontos.SetFocus
        ControlEjecutar = -1
        End If
End Sub

Private Sub cmdCancelar_Click()
    DGDetalleMontos.Height = 3255
    Call LimpiaControles
    Call AcitvaControles(False)
    DGDetalleMontos.SetFocus
    ControlEjecutar = -1
    Exit Sub
End Sub

Private Function ObtenerTipoCambioFecha(ByVal psFecha As String) As Double
    Dim rss As Recordset
    Dim ooContTM As DContingenciaTipoMonto
    Set ooContTM = New DContingenciaTipoMonto
    Set rss = New Recordset
    Set rss = ooContTM.ObtenerTipoCambioFecha(psFecha)
    ObtenerTipoCambioFecha = rss!nValFijo
    Set ooContTM = Nothing
    Set rss = Nothing
End Function

Private Sub txtFecRegContable_LostFocus()
    If Not IsDate(txtFecRegContable) Then
        MsgBox "Verifique Dia, Mes, Año , Fecha Incorrecta", vbInformation, "Aviso"
    End If
End Sub
'END TORE



