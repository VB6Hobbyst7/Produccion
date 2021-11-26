VERSION 5.00
Begin VB.Form frmRHContratoSeleccion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   Icon            =   "frmRHContratoSeleccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInicializar 
      Caption         =   "&Inicializar"
      Height          =   375
      Left            =   4260
      TabIndex        =   16
      Top             =   4965
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   15
      Top             =   4965
      Width           =   975
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Elminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1110
      TabIndex        =   14
      Top             =   4965
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8385
      TabIndex        =   13
      Top             =   4965
      Width           =   975
   End
   Begin Sicmact.TxtBuscar txtAreaCargo 
      Height          =   285
      Left            =   1245
      TabIndex        =   9
      Top             =   750
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   503
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Enabled         =   0   'False
      Appearance      =   0
      sTitulo         =   ""
      EnabledText     =   0   'False
   End
   Begin Sicmact.TxtBuscar txtTpoCon 
      Height          =   285
      Left            =   1245
      TabIndex        =   7
      Top             =   405
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   503
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Enabled         =   0   'False
      Appearance      =   0
      sTitulo         =   ""
      EnabledText     =   0   'False
   End
   Begin Sicmact.TxtBuscar txtEval 
      Height          =   285
      Left            =   1245
      TabIndex        =   5
      Top             =   60
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   503
      Appearance      =   0
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
   Begin VB.CommandButton cmdDesierto 
      Caption         =   "&Desierto"
      Height          =   375
      Left            =   3195
      TabIndex        =   4
      Top             =   4965
      Width           =   975
   End
   Begin VB.CommandButton cmdContratar 
      Caption         =   "&Contratar"
      Height          =   375
      Left            =   2145
      TabIndex        =   3
      Top             =   4965
      Width           =   975
   End
   Begin VB.Frame fraExamenSeleccion 
      Caption         =   "Personas Seleccion"
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
      Height          =   3810
      Index           =   4
      Left            =   60
      TabIndex        =   0
      Top             =   1110
      Width           =   9315
      Begin Sicmact.FlexEdit FlexExamen 
         Height          =   3495
         Left            =   105
         TabIndex        =   1
         Top             =   225
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   6165
         Cols0           =   22
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   $"frmRHContratoSeleccion.frx":030A
         EncabezadosAnchos=   "500-1500-3000-1500-1500-1500-1500-1500-600-600-1200-1200-800-4000-2000-800-2500-800-1200-1200-1800-4000"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-9-10-11-12-X-14-15-X-17-X-19-X-21"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-3-0-4-2-2-1-0-0-1-0-1-0-1-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R-R-R-R-R-C-L-R-R-L-L-R-L-L-L-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-2-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbRsLoad        =   -1  'True
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   495
         RowHeight0      =   300
         CellForeColor   =   -2147483627
      End
   End
   Begin VB.Label lblAreaCargo 
      Caption         =   "Area Cargo :"
      Height          =   255
      Left            =   165
      TabIndex        =   12
      Top             =   780
      Width           =   1050
   End
   Begin VB.Label lblTpoCon 
      Caption         =   "Tpo.Contrato :"
      Height          =   255
      Left            =   165
      TabIndex        =   11
      Top             =   435
      Width           =   1065
   End
   Begin VB.Label lblAreaCargoRes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3150
      TabIndex        =   10
      Top             =   780
      Width           =   5760
   End
   Begin VB.Label lblTpoConRes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3150
      TabIndex        =   8
      Top             =   435
      Width           =   5760
   End
   Begin VB.Label lblEvaluacionRes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3150
      TabIndex        =   6
      Top             =   90
      Width           =   5760
   End
   Begin VB.Label lblEval 
      Caption         =   "Evaluación :"
      Height          =   255
      Left            =   165
      TabIndex        =   2
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "frmRHContratoSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnTipoContratacion As ContratoForma
'ALPA 20090122******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub cmdAgregar_Click()
    If Me.txtTpoCon.Text = "" Then
        MsgBox "Debe Ingresar un tipo contrato valido.", vbInformation, "Aviso"
        txtTpoCon.SetFocus
        Exit Sub
    ElseIf Me.txtAreaCargo.Text = "" Then
        MsgBox "Debe Ingresar una Area - Cargo valido.", vbInformation, "Aviso"
        txtAreaCargo.SetFocus
        Exit Sub
    End If
    
    Me.FlexExamen.AdicionaFila
    Me.txtTpoCon.Enabled = False
    Me.txtAreaCargo.Enabled = False
    FlexExamen.SetFocus
End Sub

Private Sub cmdContratar_Click()
    Dim oEva As NActualizaProcesoSeleccion
    Dim oRh As DActualizaDatosRRHH
    Dim i As Integer
    Dim rsEva As ADODB.Recordset
    Set rsEva = New ADODB.Recordset
    Set oEva = New NActualizaProcesoSeleccion
    Set oRh = New DActualizaDatosRRHH
    
    If Me.FlexExamen.TextMatrix(FlexExamen.Rows - 1, 1) = "" Then
        MsgBox "Debe Ingresar por lo menos a una persona para que sea contratada.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If lnTipoContratacion = ContratoFormaAutomatica Then
        If Me.txtEval.Text = "" Then Exit Sub
        If Valida(False) Then
            If MsgBox("Desea Contratatar a personal del proceso Seleccion ? " & oImpresora.gPrnSaltoLinea & " El Proceso de Seleccion no podra se modificado." & oImpresora.gPrnSaltoLinea & "(Solo se contrataran las personas que estan confirmadas)", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
            
            For i = 1 To Me.FlexExamen.Rows - 1
                If Me.FlexExamen.TextMatrix(i, 9) = "." Then
                    If Not ValidaItem(i) Then
                        Exit Sub
                    End If
                End If
            Next i
            'ALPA 20090122*************************************
            glsMovNro = GetMovNro(gsCodUser, gsCodAge)
            '**************************************************
            Set rsEva = Me.FlexExamen.GetRsNew
            oRh.AgregaRRHHLote rsEva, Left(Me.txtAreaCargo.Text, 3), Right(Me.txtAreaCargo.Text, 6), Me.txtTpoCon.Text, glsMovNro, Format(gdFecSis, gsFormatoFecha), gsFormatoFecha
            
            oEva.ModificaEstadoProSelec txtEval.Text, "2", glsMovNro
            'ALPA 20090122*************************************
             gsOpeCod = LogPistaRegistraContratoProcesoSelección
             objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", , txtEval.Text, gNumeroProcesoSeleccion
            '**************************************************
        Else
            MsgBox "Para poder contratar personal de un proceso de Seleccion debe confirmar por lo menos a un postulante.", vbInformation, "Aviso"
            Exit Sub
        End If
        Set oEva = Nothing
        Form_Load
        cmdContratar.Enabled = False
        cmdInicializar.Enabled = False
        cmdDesierto.Enabled = False
    Else
        If MsgBox("Desea Contratatar a las personas ingresadas ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
        
        For i = 1 To Me.FlexExamen.Rows - 1
            If Not ValidaItem(i) Then
                Me.FlexExamen.SetFocus
                Exit Sub
            End If
        Next i
        'ALPA 20090122*************************************
        glsMovNro = GetMovNro(gsCodUser, gsCodAge)
        '**************************************************
        Set rsEva = Me.FlexExamen.GetRsNew
        oRh.AgregaRRHHLote rsEva, Left(Me.txtAreaCargo.Text, 3), Right(Me.txtAreaCargo.Text, 6), Me.txtTpoCon.Text, glsMovNro, Format(gdFecSis, gsFormatoFecha), gsFormatoFecha, False
        oEva.ModificaEstadoProSelec txtEval.Text, "2", glsMovNro
        'ALPA 20090122*************************************
        gsOpeCod = LogPistaRegistraContratoManual
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", , txtEval.Text, gNumeroProcesoSeleccion
        '**************************************************
        Set oEva = Nothing
        cmdInicializar_Click
        Form_Load
    End If
    
End Sub

Private Sub cmdDesierto_Click()
    Dim oEva As NActualizaProcesoSeleccion
    Set oEva = New NActualizaProcesoSeleccion
    
    If Me.txtEval.Text = "" Then Exit Sub
    If Valida(True) Then
        If MsgBox("Desea Cerrar proceso Seleccion ?,  - El Procesa no podra se modificado.", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
        'ALPA 20090122*************************************
         glsMovNro = GetMovNro(gsCodUser, gsCodAge)
        '**************************************************
        oEva.ModificaEstadoProSelec txtEval.Text, "3", glsMovNro
        'ALPA 20090122*************************************
        gsOpeCod = LogPistaRegistraContratoProcesoSelección
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2", , txtEval.Text, gNumeroProcesoSeleccion
        '**************************************************
    Else
        MsgBox "Para declarar desierto un proceso de Seleccion no debe confirmar a ningun postulante.", vbInformation, "Aviso"
    End If
    
    Set oEva = Nothing
    Form_Load
End Sub

Private Sub CmdEliminar_Click()
    Me.FlexExamen.EliminaFila FlexExamen.Row
End Sub

Private Sub cmdInicializar_Click()
    If lnTipoContratacion = ContratoFormaManual Then
        Me.txtTpoCon.Text = ""
        Me.txtAreaCargo.Text = ""
        Me.txtTpoCon.Enabled = True
        Me.txtAreaCargo.Enabled = True
        FlexExamen.Clear
        FlexExamen.Rows = 2
        FlexExamen.FormaCabecera
        Me.lblAreaCargoRes.Caption = ""
        Me.lblTpoConRes.Caption = ""
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub Ini(lsCaption As String, pnTipoContratacion As ContratoForma)
    lnTipoContratacion = pnTipoContratacion
    Caption = lsCaption
    Me.Show 1
End Sub

Private Sub CargaDatos()
    Dim oEva As DActualizaProcesoSeleccion
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim rsEDet As ADODB.Recordset
    Set rsEDet = New ADODB.Recordset
    
    Set oEva = New DActualizaProcesoSeleccion
    Set rsE = oEva.GetProcesoSeleccion(Me.txtEval.Text)
    
    If Not (rsE.EOF And rsE.BOF) Then
        Me.txtAreaCargo.Text = rsE!area & rsE!Cargo
        Me.lblAreaCargoRes.Caption = Me.txtAreaCargo.psDescripcion
        Me.txtTpoCon.Text = rsE!TipoContrato
        Me.lblTpoConRes.Caption = Me.txtTpoCon.psDescripcion
        Set rsEDet = oEva.GetProcesosSeleccionDetExamen(Me.txtEval.Text, RHTipoOpeEvaluacion.RHTipoOpeEvaConsolidado, True)
        
        If rsEDet.EOF And rsEDet.BOF Then Exit Sub
        
        If Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Then
            Me.FlexExamen.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-9-10-X-12-X-14-15-X-17-X-19-X-21"
        ElseIf Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo Then
            Me.FlexExamen.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-9-10-11-12-X-14-15-X-17-X-19-X-21"
        Else
            Me.FlexExamen.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-9-10-11-X-X-14-15-X-X-X-X-X-21"
        End If
        
        Set Me.FlexExamen.Recordset = rsEDet
    Else
        FlexExamen.Clear
        FlexExamen.Rows = 2
        FlexExamen.FormaCabecera
    End If
End Sub

Private Sub FlexExamen_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim lnI As Integer
    
    If lnTipoContratacion = ContratoFormaAutomatica Then
        If pnCol = 17 And psDataCod <> "" Then
            If psDataCod <> RHEmpleadoFonfoTipo.RHEmpleadoFonfoTipoAFP Then
                Me.FlexExamen.TextMatrix(pnRow, 19) = ""
                Me.FlexExamen.TextMatrix(pnRow, 20) = ""
            End If
        ElseIf pnCol = 17 And psDataCod <> "" Then
            Me.FlexExamen.TextMatrix(pnRow, 19) = ""
            Me.FlexExamen.TextMatrix(pnRow, 20) = ""
        End If
    Else
        If pnCol = 10 And psDataCod <> "" And IsNumeric(psDataCod) Then
            If psDataCod <> RHEmpleadoFonfoTipo.RHEmpleadoFonfoTipoAFP Then
                Me.FlexExamen.TextMatrix(pnRow, 12) = ""
                Me.FlexExamen.TextMatrix(pnRow, 13) = ""
            End If
        ElseIf pnCol = 10 And psDataCod <> "" And IsNumeric(psDataCod) Then
            Me.FlexExamen.TextMatrix(pnRow, 12) = ""
            Me.FlexExamen.TextMatrix(pnRow, 13) = ""
        End If
        If pnCol = 1 Then
            For lnI = 1 To Me.FlexExamen.Rows - 1
                If psDataCod = Me.FlexExamen.TextMatrix(lnI, 1) And lnI <> FlexExamen.Row Then
                    Me.FlexExamen.TextMatrix(FlexExamen.Row, 1) = ""
                    Me.FlexExamen.TextMatrix(FlexExamen.Row, 2) = ""
                    Exit Sub
                End If
            Next lnI
        End If
    End If
End Sub

Private Sub FlexExamen_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If Me.txtTpoCon.Text <> 0 And pnCol = 4 Then
        If IsDate(FlexExamen.TextMatrix(pnRow, 3)) And IsDate(FlexExamen.TextMatrix(pnRow, 4)) Then
            If CDate(FlexExamen.TextMatrix(pnRow, 3)) > CDate(FlexExamen.TextMatrix(pnRow, 4)) Then
                Cancel = False
            End If
        End If
    End If
    If pnCol = 7 Then
        If FlexExamen.TextMatrix(pnRow, 7) < 0 Then
            Cancel = False
        End If
    End If
    
    
End Sub

Private Sub FlexExamen_RowColChange()
    Dim oAsisMedica As DActualizaAsistMedicaPrivada
    Dim oCon As DConstantes
    Dim oAFP As DActualizaDatosAFP
            
    If Me.txtTpoCon.Text = "" Then
        Exit Sub
    ElseIf Me.txtAreaCargo.Text = "" Then
        Exit Sub
    End If
            
            
    If lnTipoContratacion = ContratoFormaAutomatica Then
        If Me.FlexExamen.Col = 12 And (Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Or Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo) Then
            Set oAsisMedica = New DActualizaAsistMedicaPrivada
            Me.FlexExamen.rsTextBuscar = oAsisMedica.GetAsisMedPriv(True)
            Set oAsisMedica = Nothing
        ElseIf Me.FlexExamen.Col = 15 Then
            Set oCon = New DConstantes
            Me.FlexExamen.rsTextBuscar = oCon.GetAgencias(, , True, Left(txtAreaCargo, 3))
            Set oCon = Nothing
        ElseIf Me.FlexExamen.Col = 17 And (Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Or Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo) Then
            Set oCon = New DConstantes
            Me.FlexExamen.rsTextBuscar = oCon.GetConstante(gRHEmpleadoFonfoTipo, , , True)
            Set oCon = Nothing
        ElseIf Me.FlexExamen.Col = 19 And CInt(IIf(Me.FlexExamen.TextMatrix(Me.FlexExamen.Row, 17) = "", -1, Me.FlexExamen.TextMatrix(Me.FlexExamen.Row, 17))) = RHEmpleadoFonfoTipo.RHEmpleadoFonfoTipoAFP And (Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Or Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo) Then
            Set oAFP = New DActualizaDatosAFP
            Me.FlexExamen.rsTextBuscar = oAFP.GetAFP(True)
            Set oAFP = Nothing
            If Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Then
                Me.FlexExamen.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-9-10-X-12-X-14-15-X-17-X-19-X-21"
            ElseIf Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo Then
                Me.FlexExamen.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-9-10-11-12-X-14-15-X-17-X-19-X-21"
            End If
        ElseIf Me.FlexExamen.Col = 19 And (Me.FlexExamen.TextMatrix(Me.FlexExamen.Row, 17) = "" Or Me.FlexExamen.TextMatrix(Me.FlexExamen.Row, 17) = "1") Then
            If Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Then
                Me.FlexExamen.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-9-10-X-12-X-14-15-X-17-X-X-X-21"
            ElseIf Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo Then
                Me.FlexExamen.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-9-10-11-12-X-14-15-X-17-X-X-X-21"
            End If
        End If
    Else
        If Me.FlexExamen.Col = 1 Or Me.FlexExamen.Col = 2 Then
            Me.FlexExamen.TipoBusqueda = BuscaPersona
        Else
            If Me.FlexExamen.TipoBusqueda <> BuscaArbol Then Me.FlexExamen.TipoBusqueda = BuscaArbol
        End If
        
        If Me.FlexExamen.Col = 11 And Not IsNumeric(Me.FlexExamen.TextMatrix(Me.FlexExamen.Row, 10)) Then
            FlexExamen.TextMatrix(Me.FlexExamen.Row, 10) = ""
        End If
            
        If Me.FlexExamen.Col = 5 And (IIf(Me.txtTpoCon.Text = "", "-1", Me.txtTpoCon.Text) = RHContratoTipo.RHContratoTipoIndeterminado Or IIf(Me.txtTpoCon.Text = "", "-1", Me.txtTpoCon.Text) = RHContratoTipo.RHContratoTipoFijo Or IIf(Me.txtTpoCon.Text = "", "-1", Me.txtTpoCon.Text) = RHContratoTipo.RHContratoTipoDirector) Then
            Set oAsisMedica = New DActualizaAsistMedicaPrivada
            Me.FlexExamen.rsTextBuscar = oAsisMedica.GetAsisMedPriv(True)
            Set oAsisMedica = Nothing
        ElseIf Me.FlexExamen.Col = 8 Then
            Set oCon = New DConstantes
            Me.FlexExamen.rsTextBuscar = oCon.GetAgencias(, , True, Left(txtAreaCargo, 3))
            Set oCon = Nothing
        ElseIf Me.FlexExamen.Col = 10 And (IIf(Me.txtTpoCon.Text = "", "-1", Me.txtTpoCon.Text) = RHContratoTipo.RHContratoTipoIndeterminado Or IIf(Me.txtTpoCon.Text = "", "-1", Me.txtTpoCon.Text) = RHContratoTipo.RHContratoTipoFijo) Then
            Set oCon = New DConstantes
            Me.FlexExamen.rsTextBuscar = oCon.GetConstante(gRHEmpleadoFonfoTipo, , , True)
            Set oCon = Nothing
        ElseIf Me.FlexExamen.Col = 12 And CInt(IIf(Me.FlexExamen.TextMatrix(Me.FlexExamen.Row, 10) = "", -1, IIf(IsNumeric(Me.FlexExamen.TextMatrix(Me.FlexExamen.Row, 10)), Me.FlexExamen.TextMatrix(Me.FlexExamen.Row, 10), -1))) = RHEmpleadoFonfoTipo.RHEmpleadoFonfoTipoAFP And (Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Or Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo) Then
            Set oAFP = New DActualizaDatosAFP
            Me.FlexExamen.rsTextBuscar = oAFP.GetAFP(True)
            Set oAFP = Nothing
            If Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Then
                Me.FlexExamen.ColumnasAEditar = "X-1-X-3-X-5-X-7-8-X-10-X-12-X-14"
            ElseIf Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo Then
                Me.FlexExamen.ColumnasAEditar = "X-1-X-3-4-5-X-7-8-X-10-X-12-X-14"
            End If
        ElseIf Me.FlexExamen.Col = 12 And (Me.FlexExamen.TextMatrix(Me.FlexExamen.Row, 10) = "" Or Me.FlexExamen.TextMatrix(Me.FlexExamen.Row, 10) = "1") Then
            If Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Then
                Me.FlexExamen.ColumnasAEditar = "X-1-X-3-X-5-X-7-8-X-10-X-X-X-14"
            ElseIf Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo Then
                Me.FlexExamen.ColumnasAEditar = "X-1-X-3-4-5-X-7-8-X-10-X-X-X-14"
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim oEva As DActualizaProcesoSeleccion
    Dim oDatoAreas As DActualizaDatosArea
    Set oDatoAreas = New DActualizaDatosArea
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    
    Set oEva = New DActualizaProcesoSeleccion
    
    Me.txtTpoCon.rs = oCon.GetConstante(gRHContratoTipo, , , True)
    Me.txtAreaCargo.rs = oDatoAreas.GetCargosAreas
    
    If lnTipoContratacion = ContratoFormaAutomatica Then
        Set rsE = oEva.GetProcesosSeleccion(RHProcesoSeleccionEstado.gRHProcSelEstFinalizado)
        Me.txtEval.rs = rsE
    Else
        Me.FlexExamen.Cols = 15
        Me.FlexExamen.ColumnasAEditar = "X-1-X-3-4-5-X-7-8-X-10-X-12-X-14"
        Me.FlexExamen.EncabezadosAlineacion = "C-L-L-R-R-L-L-R-L-L-L-L-L-L-L"
        Me.FlexExamen.EncabezadosAnchos = "300-1500-5000-1200-1200-800-4000-2000-800-2500-800-1200-1200-1800-5000"
        Me.FlexExamen.EncabezadosNombres = "#-Codigo-Nombre-Inicio-Fin-Cod AMP-Asis Med Fam-Sueldo-Cod Age-Agencia-Cod Fon-Fondo-Cod AFP-AFP-Comentario"
        Me.FlexExamen.FormatosEdit = "0-0-0-0-0-0-0-2-0-0-3-0-0-0-0"
        Me.FlexExamen.ListaControles = "0-1-0-2-2-1-0-0-1-0-1-0-1-0-0"
        
        Me.cmdDesierto.Visible = False
        Me.cmdInicializar.Visible = False
        
        Me.txtEval.Visible = False
        Me.lblEval.Visible = False
        Me.lblEvaluacionRes.Visible = False
        Me.txtAreaCargo.Enabled = True
        Me.txtTpoCon.Enabled = True
        Me.cmdAgregar.Enabled = True
        Me.cmdEliminar.Enabled = True
    End If
    'ALPA 20090122 ***************************************************************************
    Set objPista = New COMManejador.Pista
    '*****************************************************************************************
    Set oEva = Nothing
    Set oDatoAreas = Nothing
    Set oCon = Nothing
End Sub


Private Sub txtAreaCargo_EmiteDatos()
    Me.lblAreaCargoRes = txtAreaCargo.psDescripcion
End Sub

Private Sub txtEval_EmiteDatos()
    Me.lblEvaluacionRes.Caption = txtEval.psDescripcion
    CargaDatos
End Sub

Private Sub Activa(pbvalor As Boolean)
    
End Sub

Private Sub Limpia()
    Me.txtEval.Text = ""
End Sub

Private Function Valida(pbDesiertos As Boolean) As Boolean
    Dim i As Integer
    Dim lbBan As Boolean
    
    If pbDesiertos Then
        lbBan = True
        For i = 1 To Me.FlexExamen.Rows - 1
            If Me.FlexExamen.TextMatrix(i, 9) = "." Then
                lbBan = False
                i = Me.FlexExamen.Rows - 1
            End If
        Next i
    Else
        lbBan = False
        For i = 1 To Me.FlexExamen.Rows - 1
            If Me.FlexExamen.TextMatrix(i, 9) = "." Then
                lbBan = True
                i = Me.FlexExamen.Rows - 1
            End If
        Next i
    End If
    
    Valida = lbBan
End Function

Private Function ValidaItem(pnI As Integer) As Boolean
    Me.FlexExamen.Row = pnI
    
    If lnTipoContratacion = ContratoFormaAutomatica Then
        If Not IsDate(Me.FlexExamen.TextMatrix(pnI, 10)) Then
            MsgBox "Debe Ingresar un Fecha Valida para el inicio de contrato del Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 10
            ValidaItem = False
            Exit Function
        ElseIf Not IsDate(Me.FlexExamen.TextMatrix(pnI, 11)) And Me.txtTpoCon.Text <> RHContratoTipo.RHContratoTipoIndeterminado And Me.txtTpoCon.Text <> RHContratoTipo.RHContratoTipoDirector Then
            MsgBox "Debe Ingresar un Fecha Valida para el fin de contrato del Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 11
            ValidaItem = False
            Exit Function
        ElseIf Me.FlexExamen.TextMatrix(pnI, 12) = "" And (Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Or Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo) Then
            MsgBox "Debe Ingresar una categoria de la asistencia medica privada Valida del Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 12
            ValidaItem = False
            Exit Function
        ElseIf Not IsNumeric(Me.FlexExamen.TextMatrix(pnI, 14)) Then
            MsgBox "Debe Ingresar un sueldo valido para el Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 14
            ValidaItem = False
            Exit Function
        ElseIf Me.FlexExamen.TextMatrix(pnI, 15) = "" Then
            MsgBox "Debe Ingresar una agencia de ubicacion valida para el Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 15
            ValidaItem = False
            Exit Function
        ElseIf Me.FlexExamen.TextMatrix(pnI, 17) = "" And (Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Or Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo) Then
            MsgBox "Debe Ingresar un fondo de pensiones valido para el Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 17
            ValidaItem = False
            Exit Function
        ElseIf Me.FlexExamen.TextMatrix(pnI, 19) = "" And (Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Or Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo) And IIf(IsNumeric(Me.FlexExamen.TextMatrix(pnI, 17)), Me.FlexExamen.TextMatrix(pnI, 17), -1) = RHEmpleadoFonfoTipo.RHEmpleadoFonfoTipoAFP Then
            MsgBox "Debe Ingresar un fondo de pensiones valido para el Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 19
            ValidaItem = False
            Exit Function
        Else
            ValidaItem = True
        End If
    Else
        If Not IsDate(Me.FlexExamen.TextMatrix(pnI, 3)) Then
            MsgBox "Debe Ingresar un Fecha Valida para el inicio de contrato del Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 3
            ValidaItem = False
            Exit Function
        ElseIf Not IsDate(Me.FlexExamen.TextMatrix(pnI, 4)) And Me.txtTpoCon.Text <> RHContratoTipo.RHContratoTipoIndeterminado And Me.txtTpoCon.Text <> RHContratoTipo.RHContratoTipoDirector Then
            MsgBox "Debe Ingresar un Fecha Valida para el fin de contrato del Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 4
            ValidaItem = False
            Exit Function
        ElseIf Me.FlexExamen.TextMatrix(pnI, 5) = "" And (Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Or Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo) Then
            MsgBox "Debe Ingresar una categoria de la asistencia medica privada Valida del Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 5
            ValidaItem = False
            Exit Function
        ElseIf Not IsNumeric(Me.FlexExamen.TextMatrix(pnI, 7)) Then
            MsgBox "Debe Ingresar un sueldo valido para el Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 7
            ValidaItem = False
            Exit Function
        ElseIf Me.FlexExamen.TextMatrix(pnI, 8) = "" Then
            MsgBox "Debe Ingresar una agencia de ubicacion valida para el Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 8
            ValidaItem = False
            Exit Function
        ElseIf Me.FlexExamen.TextMatrix(pnI, 10) = "" And (Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Or Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo) Then
            MsgBox "Debe Ingresar un fondo de pensiones valido para el Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 10
            ValidaItem = False
            Exit Function
        ElseIf Me.FlexExamen.TextMatrix(pnI, 12) = "" And (Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Or Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo) And IIf(Me.FlexExamen.TextMatrix(pnI, 10) = "", -1, Me.FlexExamen.TextMatrix(pnI, 10)) = RHEmpleadoFonfoTipo.RHEmpleadoFonfoTipoAFP Then
            MsgBox "Debe Ingresar un fondo de pensiones valido para el Postulante : " & oImpresora.gPrnSaltoLinea & " - " & Me.FlexExamen.TextMatrix(pnI, 2), vbInformation, "Aviso"
            FlexExamen.Col = 12
            ValidaItem = False
            Exit Function
        Else
            ValidaItem = True
        End If
    End If
End Function

Private Sub txtTpoCon_EmiteDatos()
    Me.lblTpoConRes.Caption = Me.txtTpoCon.psDescripcion
    
    If Me.txtTpoCon.Text = "" Then Exit Sub
    If lnTipoContratacion = ContratoFormaManual Then
        Me.FlexExamen.Cols = 15
        If Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoIndeterminado Then
            Me.FlexExamen.ColumnasAEditar = "X-1-X-3-X-5-X-7-8-X-10-X-12-X-14"
        ElseIf Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoFijo Then
            Me.FlexExamen.ColumnasAEditar = "X-1-X-3-4-5-X-7-8-X-10-X-12-X-14"
        ElseIf Me.txtTpoCon.Text = RHContratoTipo.RHContratoTipoDirector Then
            Me.FlexExamen.ColumnasAEditar = "X-1-X-3-X-5-X-7-8-X-X-X-X-X-14"
            
        Else
            Me.FlexExamen.ColumnasAEditar = "X-1-X-3-4-X-X-7-8-X-X-X-X-X-14"
        End If
    End If
End Sub
