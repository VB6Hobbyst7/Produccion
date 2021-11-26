VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRHPlanillaMovExtra 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   Icon            =   "frmRHPlanillaMovExtra.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2400
      TabIndex        =   25
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame fraDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Datos Planilla"
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
      Height          =   4095
      Left            =   60
      TabIndex        =   21
      Top             =   1815
      Width           =   10800
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   8475
         TabIndex        =   24
         Top             =   3645
         Width           =   1095
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   9615
         TabIndex        =   23
         Top             =   3645
         Width           =   1095
      End
      Begin Sicmact.FlexEdit Flex 
         Height          =   3360
         Left            =   75
         TabIndex        =   22
         Top             =   240
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   5927
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Codigo-Nombre-Monto-Comentario-Cta-Tipo-a-Destino-Cod Ref-Referencia-Cta Referencia"
         EncabezadosAnchos=   "300-1500-2800-900-2800-1-1200-0-2100-1200-2800-1200"
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
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3-4-X-6-X-8-9-X-11"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-0-0-0-3-0-3-1-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-R-L-C-L-C-L-L-L-L"
         FormatosEdit    =   "0-0-0-2-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   7
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblTotAG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   2415
         TabIndex        =   29
         Top             =   3660
         Width           =   1590
      End
      Begin VB.Label lblTotalG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   5820
         TabIndex        =   27
         Top             =   3660
         Width           =   1590
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Cargos"
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
         Height          =   285
         Left            =   4515
         TabIndex        =   26
         Top             =   3660
         Width           =   2295
      End
      Begin VB.Label lblTotA 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Abonos"
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
         Height          =   285
         Left            =   1110
         TabIndex        =   30
         Top             =   3660
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9690
      TabIndex        =   11
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   45
      TabIndex        =   8
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame fraPlanilla 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Planilla"
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
      Height          =   1755
      Left            =   60
      TabIndex        =   12
      Top             =   30
      Width           =   10785
      Begin VB.Frame fraFijo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Valor Fijo"
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
         Height          =   1125
         Left            =   3660
         TabIndex        =   18
         Top             =   510
         Width           =   7035
         Begin VB.Frame fraTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            Height          =   450
            Left            =   3525
            TabIndex        =   28
            Top             =   135
            Width           =   3390
            Begin VB.OptionButton optAbono 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Caption         =   "Abono"
               ForeColor       =   &H80000008&
               Height          =   165
               Left            =   135
               TabIndex        =   5
               Top             =   195
               Width           =   1125
            End
            Begin VB.OptionButton opcCargo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Caption         =   "Cargo"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1665
               TabIndex        =   6
               Top             =   180
               Value           =   -1  'True
               Width           =   960
            End
         End
         Begin VB.TextBox txtComentarioFijo 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   1065
            MaxLength       =   150
            TabIndex        =   7
            Top             =   645
            Width           =   5850
         End
         Begin Sicmact.EditMoney txtMontoFijo 
            Height          =   375
            Left            =   1065
            TabIndex        =   4
            Top             =   225
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
         End
         Begin VB.Label lblComent 
            Caption         =   "Comentario :"
            Height          =   210
            Left            =   135
            TabIndex        =   20
            Top             =   615
            Width           =   900
         End
         Begin VB.Label lblMontoFijo 
            Caption         =   "Monto :"
            Height          =   195
            Left            =   135
            TabIndex        =   19
            Top             =   285
            Width           =   630
         End
      End
      Begin VB.Frame fraTipoOpe 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Operación"
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
         Height          =   540
         Left            =   810
         TabIndex        =   17
         Top             =   510
         Width           =   2790
         Begin VB.OptionButton opcMontoPersona 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Por Persona"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1320
            TabIndex        =   3
            Top             =   232
            Width           =   1245
         End
         Begin VB.OptionButton opcMontoFijo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Fijo"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   300
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   1035
         End
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   270
         Left            =   8580
         TabIndex        =   1
         Top             =   225
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin Sicmact.TxtBuscar txtPlanillas 
         Height          =   300
         Left            =   795
         TabIndex        =   0
         Top             =   195
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Appearance      =   0
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin VB.Label lblMes 
         Caption         =   "(yyyymm)"
         Height          =   225
         Left            =   9870
         TabIndex        =   16
         Top             =   255
         Width           =   810
      End
      Begin VB.Label lblPeriodo 
         Caption         =   "Periodo :"
         Height          =   195
         Left            =   7725
         TabIndex        =   15
         Top             =   255
         Width           =   690
      End
      Begin VB.Label lblPlanilla 
         Caption         =   "Planilla :"
         Height          =   180
         Left            =   135
         TabIndex        =   14
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblPlanillaRes 
         Height          =   255
         Left            =   2145
         TabIndex        =   13
         Top             =   210
         Width           =   4560
      End
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1215
      TabIndex        =   9
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1215
      TabIndex        =   10
      Top             =   6000
      Width           =   1095
   End
End
Attribute VB_Name = "frmRHPlanillaMovExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnTipo As TipoOpe

Private Sub cmdCancelar_Click()
    Flex.Clear
    Flex.Rows = 2
    Flex.FormaCabecera
    Limpia
    Activa False
End Sub

Private Sub cmdEditar_Click()
    If Me.txtPlanillas.Text = "" Or Me.mskFecha.Text = "______" Then Exit Sub
    Activa True
End Sub

Private Sub cmdEliminar_Click()
    Flex.EliminaFila Me.Flex.Row
End Sub

Private Sub cmdGrabar_Click()
    If Not Valida Then Exit Sub
    Dim oExtra As NActualizaMovimientoExtraPla
    Set oExtra = New NActualizaMovimientoExtraPla
    Dim oExtraV As DActualizaMovExtraPlanilla
    Set oExtraV = New DActualizaMovExtraPlanilla
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    
    If MsgBox("Desea guardar los cambios ? ", vbQuestion + vbYesNo, "Avios") = vbNo Then Exit Sub
    
    If lnTipo = gTipoOpeRegistro Then
        Set rsE = oExtraV.GetExtPlanilla(Me.mskFecha.Text, Me.txtPlanillas.Text)
        If rsE.EOF And rsE.BOF Then
            oExtra.ModificaExtPlanilla Me.mskFecha.Text, Me.txtPlanillas.Text, Flex.GetRsNew, GetMovNro(gsCodUser, gsCodAge), IIf(Me.opcMontoFijo, Me.txtMontoFijo.value, ""), IIf(Me.opcMontoFijo, Me.txtComentarioFijo.Text, ""), IIf(Me.opcMontoFijo, IIf(Me.opcCargo, RHExtraPlanillaOpeTpo.RHExtraPlanillaOpeTpoCargo, RHExtraPlanillaOpeTpo.RHExtraPlanillaOpeTpoAbono), "")
        Else
            MsgBox "Ya existen Movimientos extra para esta planilla de remuneraciones. " & Chr(13) & "Si desea modificarlos ingrese a la opcion de mantenimiento.", vbInformation, "Aviso"
        End If
        rsE.Close
        
    Else
        oExtra.ModificaExtPlanilla Me.mskFecha.Text, Me.txtPlanillas.Text, Flex.GetRsNew, GetMovNro(gsCodUser, gsCodAge), IIf(Me.opcMontoFijo, Me.txtMontoFijo.value, ""), IIf(Me.opcMontoFijo, Me.txtComentarioFijo.Text, ""), IIf(Me.opcMontoFijo, IIf(Me.opcCargo, RHExtraPlanillaOpeTpo.RHExtraPlanillaOpeTpoCargo, RHExtraPlanillaOpeTpo.RHExtraPlanillaOpeTpoAbono), "")
    End If
    
    Set rsE = Nothing
    Set oExtra = Nothing
    Set oExtraV = Nothing
    cmdCancelar_Click
End Sub

Private Sub cmdImprimir_Click()
    Dim oExt As NActualizaMovimientoExtraPla
    Set oExt = New NActualizaMovimientoExtraPla
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    Dim lsCadena As String
    
    lsCadena = oExt.GetReporte(Me.mskFecha.Text, Me.txtPlanillas.Text, Me.lblPlanillaRes, gsNomAge, gsEmpresa, gdFecSis)
    
    oPrevio.Show lsCadena, Me.Caption, True, 66
End Sub

Private Sub cmdNuevo_Click()
    Me.Flex.AdicionaFila
    Flex.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Flex_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    Dim lsCodCta As String
    Dim oExtra As DActualizaMovExtraPlanilla
    Set oExtra = New DActualizaMovExtraPlanilla

    Me.Flex.rsTextBuscar = oExtra.GetPersonasExtraPlanillaAdd
    
    If pnCol = 1 Then
        lsCodCta = oRH.GetCuentaRRHH(Me.Flex.TextMatrix(pnRow, 1), Me.txtPlanillas.Text)
        
        If lsCodCta = "" Then
            MsgBox "La persona seleccionada no tiene cuenta de empleado.", vbInformation, "Aviso"
            Me.Flex.TextMatrix(pnRow, 1) = ""
            Me.Flex.TextMatrix(pnRow, 2) = ""
            Me.Flex.SetFocus
            Exit Sub
        End If
        
        Me.Flex.TextMatrix(pnRow, 5) = lsCodCta
    ElseIf pnCol = 9 Then
        
        If Right(Flex.TextMatrix(pnRow, 8), 1) = 1 Then
            lsCodCta = oRH.GetCuentaRRHH(Me.Flex.TextMatrix(pnRow, 9), "E01")
            
            If lsCodCta = "" And Flex.TipoBusqueda = buscaempleado Then
                MsgBox "La persona seleccionada no tiene cuenta de empleado.", vbInformation, "Aviso"
                Me.Flex.TextMatrix(pnRow, 9) = ""
                Me.Flex.TextMatrix(pnRow, 10) = ""
                Me.Flex.SetFocus
                Exit Sub
            End If
        Else
            lsCodCta = oExtra.GetCtaPersonasExtraPlanillaAdd(Me.Flex.TextMatrix(pnRow, 9))
            
            If lsCodCta = "" And Flex.TipoBusqueda = buscaempleado Then
                MsgBox "La persona seleccionada no tiene cuenta de empleado.", vbInformation, "Aviso"
                Me.Flex.TextMatrix(pnRow, 9) = ""
                Me.Flex.TextMatrix(pnRow, 10) = ""
                Me.Flex.SetFocus
                Exit Sub
            End If
        
        End If
        
        Me.Flex.TextMatrix(pnRow, 11) = lsCodCta
    End If

End Sub

Private Sub Flex_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim lnI As Integer
    Dim lnAcumC As Currency
    Dim lnAcumA As Currency
    
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    lnAcumC = 0
    lnAcumA = 0
    
    For lnI = 1 To Me.Flex.Rows - 1
        If IsNumeric(Flex.TextMatrix(lnI, 3)) Then
            If Right(Flex.TextMatrix(lnI, 6), 1) = RHExtraPlanillaOpeTpo.RHExtraPlanillaOpeTpoCargo Then
                lnAcumC = lnAcumC + CCur(Flex.TextMatrix(lnI, 3))
            Else
                lnAcumA = lnAcumA + CCur(Flex.TextMatrix(lnI, 3))
            End If
        End If
    Next lnI
    
    Me.lblTotalG.Caption = Format(lnAcumC, "#,##0.00")
    Me.lblTotAG.Caption = Format(lnAcumA, "#,##0.00")
  
    If pnCol = 8 Then
        If Right(Flex.TextMatrix(Flex.Row, 8), 1) = "0" Then
            Me.Flex.TextMatrix(Flex.Row, 9) = ""
            Me.Flex.TextMatrix(Flex.Row, 10) = ""
            Me.Flex.TextMatrix(Flex.Row, 11) = ""
        End If
    ElseIf pnCol = 11 Then
        If Not oRH.ValidaCta(Flex.TextMatrix(pnRow, pnCol), gbBitCentral) Then
            MsgBox "La cuenta no tiene un estado valido.", vbInformation, "Aviso"
            Cancel = False
        End If
    End If
End Sub

Private Sub Flex_RowColChange()
    Dim oCons As DConstantes
    Set oCons = New DConstantes
    
    If Flex.TextMatrix(Flex.Row, 8) = "" Then
        If Me.opcMontoFijo.value Then
            Me.Flex.ColumnasAEditar = "X-1-X-X-X-X-X-X-8-9-X-11"
        Else
            Me.Flex.ColumnasAEditar = "X-1-X-3-4-X-6-X-8-9-X-11"
        End If
    Else
        If Right(Flex.TextMatrix(Flex.Row, 8), 1) = "0" Then
            If Me.opcMontoFijo.value Then
                Me.Flex.ColumnasAEditar = "X-1-X-X-X-X-X-X-8-X-X-X"
            Else
                Me.Flex.ColumnasAEditar = "X-1-X-3-4-X-6-X-8-X-X-X"
            End If
        Else
            If Me.opcMontoFijo.value Then
                Me.Flex.ColumnasAEditar = "X-1-X-X-X-X-X-X-8-9-X-11"
            Else
                Me.Flex.ColumnasAEditar = "X-1-X-3-4-X-6-X-8-9-X-11"
            End If
        End If
    End If
    
    If Me.Flex.Col = 6 Then
        Me.Flex.CargaCombo oCons.GetConstante(6038)
    ElseIf Me.Flex.Col = 8 Then
        Me.Flex.CargaCombo oCons.GetConstante(6053)
    ElseIf Me.Flex.Col = 9 Then
        If Right(Flex.TextMatrix(Flex.Row, 8), 1) = "1" Then
            Flex.TipoBusqueda = buscaempleado
        Else
            Flex.TipoBusqueda = BuscaArbol
            Dim oExtra As DActualizaMovExtraPlanilla
            Set oExtra = New DActualizaMovExtraPlanilla
        
            Me.Flex.rsTextBuscar = oExtra.GetPersonasExtraPlanillaAdd
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    

    Me.txtPlanillas.rs = oPla.GetPlanillas(, , False)
    Activa False
    
    opcMontoFijo_Click
    

End Sub

Private Sub mskFecha_GotFocus()
    mskFecha.SelStart = 0
    mskFecha.SelLength = 8
End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)
    Dim oExtra As DActualizaMovExtraPlanilla
    Set oExtra = New DActualizaMovExtraPlanilla
    
    If KeyAscii = 13 Then
        If opcMontoFijo.Enabled Then Me.opcMontoFijo.SetFocus
        
        If lnTipo = gTipoOpeReporte Then
            If Me.mskFecha.Text = "" Or Me.txtPlanillas.Text = "" Then Exit Sub
            Me.Flex.rsFlex = oExtra.GetExtPlanilla(Me.mskFecha.Text, Me.txtPlanillas.Text, True)
            Flex_OnValidate 1, 1, True
        ElseIf lnTipo <> gTipoOpeRegistro Then
            If Me.mskFecha.Text = "" Or Me.txtPlanillas.Text = "" Then Exit Sub
            Me.Flex.rsFlex = oExtra.GetExtPlanilla(Me.mskFecha.Text, Me.txtPlanillas.Text)
            Flex_OnValidate 1, 1, True
        End If
    End If
    Set oExtra = Nothing
End Sub

Private Sub mskFecha_LostFocus()
    If Not IsNumeric(mskFecha.Text) Then
        mskFecha.SetFocus
    End If
End Sub

Private Sub opcCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtComentarioFijo.SetFocus
    End If
End Sub

Private Sub opcMontoFijo_Click()
    If opcMontoPersona.value Then
        Me.fraFijo.Enabled = False
        Me.txtComentarioFijo.Enabled = False
        Me.txtMontoFijo.Enabled = False
    Else
        Me.fraFijo.Enabled = True
        Me.txtComentarioFijo.Enabled = True
        Me.txtMontoFijo.Enabled = True
    End If
End Sub

Private Sub opcMontoFijo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If opcMontoPersona.value Then
            If cmdEditar.Enabled Then
                Me.cmdEditar.SetFocus
            Else
                Me.cmdNuevo.SetFocus
            End If
        Else
            txtMontoFijo.Enabled = True
            Me.txtComentarioFijo.Enabled = True
            
            Me.txtMontoFijo.SetFocus
        End If
    End If
End Sub

Private Sub opcMontoPersona_Click()
    If opcMontoPersona.value Then
       Me.fraFijo.Enabled = False
       Me.txtComentarioFijo.Text = ""
       Me.txtMontoFijo.value = 0
    Else
        Me.fraFijo.Enabled = True
    End If
End Sub

Private Sub opcMontoPersona_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If opcMontoPersona.value Then
            If cmdEditar.Enabled Then
                Me.cmdEditar.SetFocus
            Else
                Me.cmdNuevo.SetFocus
            End If
        Else
            Me.txtMontoFijo.SetFocus
        End If
    End If
End Sub

Private Sub optAbono_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtComentarioFijo.SetFocus
    End If
End Sub

Private Sub txtComentarioFijo_GotFocus()
    txtComentarioFijo.SelStart = 0
    txtComentarioFijo.SelLength = 200
End Sub

Private Sub txtComentarioFijo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdEditar.Enabled And cmdEditar.Visible Then
            Me.cmdEditar.SetFocus
        Else
            Me.cmdNuevo.SetFocus
        End If
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtMontoFijo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.optAbono.SetFocus
    End If
End Sub

Private Sub txtPlanillas_EmiteDatos()
    If txtPlanillas.Text = "" Then Exit Sub
    
    Dim oRH As DActualizaMovExtraPlanilla
    Set oRH = New DActualizaMovExtraPlanilla
    
    Me.lblPlanillaRes.Caption = Me.txtPlanillas.psDescripcion
    Me.Flex.rsTextBuscar = oRH.GetListaRRHH(Me.txtPlanillas.Text)
    
    Set oRH = Nothing
End Sub

Private Sub txtPlanillas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFecha.SetFocus
    End If
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.cmdSalir.Enabled = Not pbValor
    Me.txtPlanillas.Enabled = Not pbValor
    
    If lnTipo = gTipoOpeRegistro Then
        Me.fraDatos.Enabled = pbValor
        Me.cmdEditar.Visible = Not pbValor
        Me.cmdGrabar.Enabled = pbValor
        Me.cmdCancelar.Visible = pbValor
        If cmdNuevo.Enabled Then Me.cmdNuevo.SetFocus
    ElseIf lnTipo = gTipoOpeMantenimiento Then
        Me.fraDatos.Enabled = pbValor
        Me.cmdEditar.Visible = Not pbValor
        Me.cmdGrabar.Enabled = pbValor
        Me.cmdCancelar.Visible = pbValor
        Me.txtPlanillas.Enabled = Not pbValor
        Me.opcCargo.value = False
        Me.optAbono.value = False
        Me.fraTipoOpe.Enabled = False
        Me.opcMontoFijo.value = False
        opcMontoPersona.value = True
        Me.fraFijo.Enabled = False
        Me.fraTipo.Enabled = False
        Me.cmdNuevo.Visible = True
    ElseIf lnTipo = gTipoOpeConsulta Then
        Me.cmdEditar.Visible = False
        Me.cmdGrabar.Visible = False
        Me.cmdCancelar.Visible = False
        Me.cmdNuevo.Visible = False
        Me.cmdEliminar.Visible = False
        Me.Flex.lbEditarFlex = False
        fraDatos.Enabled = True
        Me.cmdImprimir.Visible = False
        Me.Flex.lbEditarFlex = False
        Me.fraFijo.Enabled = False
        Me.fraTipoOpe.Enabled = False
    ElseIf lnTipo = gTipoOpeReporte Then
        Me.cmdEliminar.Enabled = pbValor
        Me.cmdNuevo.Enabled = pbValor
        Me.cmdEditar.Enabled = pbValor
        Me.fraDatos.Enabled = True
        Me.cmdGrabar.Visible = False
        Me.Flex.lbEditarFlex = False
        Me.cmdNuevo.Visible = False
        Me.cmdEditar.Visible = False
        Me.cmdEliminar.Visible = False
    End If
End Sub

Public Sub Ini(pnTipo As TipoOpe, psCaption As String)
    lnTipo = pnTipo
    Caption = psCaption
    Me.Show 1
End Sub

Private Sub Limpia()
    Me.txtMontoFijo.value = 0
    Me.txtComentarioFijo.Text = ""
    Me.txtPlanillas.Text = ""
    Me.lblPlanillaRes.Caption = ""
    Me.mskFecha.Text = "______"
End Sub

Private Function Valida() As Boolean
    Dim i As Integer
    For i = 1 To Me.Flex.Rows - 1
        Flex.Row = i
        If Me.Flex.TextMatrix(i, 1) = "" Then
            MsgBox "Debe Ingresar una persona, para el registro " & Me.Flex.TextMatrix(i, 0) & ".", vbInformation, "Aviso"
            Flex.Col = 1
            Flex.SetFocus
            Valida = False
            Flex.Row = i
        ElseIf Me.Flex.TextMatrix(i, 2) = "" And Me.opcMontoPersona Then
            MsgBox "Debe Ingresar un monto valido, para el registro " & Me.Flex.TextMatrix(i, 0) & ".", vbInformation, "Aviso"
            Flex.Col = 2
            Flex.SetFocus
            Valida = False
        ElseIf Me.Flex.TextMatrix(i, 3) = "" And Me.opcMontoPersona Then
            MsgBox "Debe Ingresar un comentario valido, para el registro " & Me.Flex.TextMatrix(i, 0) & ".", vbInformation, "Aviso"
            Flex.Col = 3
            Flex.SetFocus
            Valida = False
        ElseIf Me.Flex.TextMatrix(i, 6) = "" And Me.opcMontoPersona Then
            MsgBox "Debe Ingresar un tipo de operación, para el registro " & Me.Flex.TextMatrix(i, 0) & ".", vbInformation, "Aviso"
            Flex.Col = 6
            Flex.SetFocus
            Valida = False
        ElseIf Right(Me.Flex.TextMatrix(i, 8), 1) <> "0" And Me.Flex.TextMatrix(i, 9) = "" Then
            MsgBox "Debe Ingresar una persona  valida, para el registro " & Me.Flex.TextMatrix(i, 0) & ".", vbInformation, "Aviso"
            Flex.Col = 9
            Flex.SetFocus
            Valida = False
        ElseIf Right(Me.Flex.TextMatrix(i, 8), 1) <> "0" And Me.Flex.TextMatrix(i, 11) = "" Then
            MsgBox "Debe Ingresar un tipo de operación, para el registro " & Me.Flex.TextMatrix(i, 0) & ".", vbInformation, "Aviso"
            Flex.Col = 11
            Flex.SetFocus
            Valida = False
        Else
            Valida = True
        End If
    Next i
End Function
