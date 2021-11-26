VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRHCurriculum 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   Icon            =   "frmRHCurriculum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   4245
      Left            =   75
      TabIndex        =   6
      Top             =   1290
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   7488
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Curriculum Vitae"
      TabPicture(0)   =   "frmRHCurriculum.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCur"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Actividades Curriculares"
      TabPicture(1)   =   "frmRHCurriculum.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraExtra"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraExtra 
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
         ForeColor       =   &H00800000&
         Height          =   3765
         Left            =   -74865
         TabIndex        =   11
         Top             =   360
         Width           =   10020
         Begin VB.CommandButton cmdNuevoExtra 
            Caption         =   "&Nuevo"
            Height          =   345
            Left            =   7755
            TabIndex        =   13
            Top             =   3285
            Width           =   1050
         End
         Begin VB.CommandButton cmdEliminarExtra 
            Caption         =   "&Eliminar"
            Height          =   345
            Left            =   8865
            TabIndex        =   12
            Top             =   3285
            Width           =   1050
         End
         Begin Sicmact.FlexEdit FlexExtra 
            Height          =   2955
            Left            =   75
            TabIndex        =   14
            Top             =   240
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   5212
            Cols0           =   14
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Cod.Tpo-Tipo-Cod Activ-Actividad-Años.Pract-Costo-Cod.Niv-Nivel-Otorgado CMACT-Comentario-UltimaActualizacion-BitCod-BitItem"
            EncabezadosAnchos=   "300-750-3500-1000-3500-1000-1000-800-2500-1500-5000-2500-0-0"
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
            ColumnasAEditar =   "X-1-X-3-X-5-6-7-X-9-10-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-1-0-1-0-0-0-1-0-4-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-L-R-R-L-L-L-L-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-3-2-0-0-0-1-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraCur 
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
         ForeColor       =   &H00800000&
         Height          =   3765
         Left            =   135
         TabIndex        =   7
         Top             =   360
         Width           =   10020
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   345
            Left            =   8865
            TabIndex        =   9
            Top             =   3285
            Width           =   1050
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   345
            Left            =   7755
            TabIndex        =   8
            Top             =   3285
            Width           =   1050
         End
         Begin Sicmact.FlexEdit Flex 
            Height          =   2955
            Left            =   15
            TabIndex        =   10
            Top             =   195
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   5212
            Cols0           =   21
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   $"frmRHCurriculum.frx":0342
            EncabezadosAnchos=   "300-750-3500-3500-700-2000-1000-1000-1000-2000-1000-800-2500-1000-800-2500-1500-5000-2500-0-0"
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
            ColumnasAEditar =   "X-1-X-3-4-X-6-7-8-X-10-11-X-13-14-X-16-17-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-1-0-0-1-0-2-2-1-0-0-1-0-0-1-0-4-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-L-L-R-R-L-L-R-L-L-R-L-L-L-L-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-3-0-1-2-0-0-0-1-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
   End
   Begin Sicmact.ctrRRHHGen ctrRRHH 
      Height          =   1200
      Left            =   60
      TabIndex        =   5
      Top             =   15
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   2117
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9195
      TabIndex        =   2
      Top             =   5625
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2445
      TabIndex        =   1
      Top             =   5625
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1245
      TabIndex        =   0
      Top             =   5625
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   5625
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1245
      TabIndex        =   4
      Top             =   5625
      Width           =   1095
   End
End
Attribute VB_Name = "frmRHCurriculum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnTipo As TipoOpe
'ALPA 20090122***********************
Dim objPista As COMManejador.Pista
'************************************

Private Sub cmdCancelar_Click()
    Limpia
    CargaDatos
    Activa False
End Sub

Private Sub cmdEditar_Click()
    If Me.ctrRRHH.psCodigoPersona = "" Then
        MsgBox "Debe elegir a una persona.", vbInformation, "Aviso"
        Me.ctrRRHH.SetFocus
        Exit Sub
    End If
    Activa True
End Sub

Private Sub CmdEliminar_Click()
    If MsgBox("Desea eliminar el registro :" & Me.Flex.Row, vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Flex.EliminaFila Flex.Row
End Sub

Private Sub cmdEliminarExtra_Click()
    If MsgBox("Desea eliminar el registro :" & Me.FlexExtra.Row, vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    FlexExtra.EliminaFila FlexExtra.Row
End Sub

Private Sub cmdGrabar_Click()
    If Not Valida Then Exit Sub
    
    Dim oCur As NActualizaDatosCurriculum
    Dim lsFecIni As String
    Dim lsFecFin As String
    Dim lnTpoPeriodo As Integer
    Dim lnNumPeriodo As Integer
    Set oCur = New NActualizaDatosCurriculum
    'ALPA 20090122*************************************
    glsMovNro = GetMovNro(gsCodUser, gsCodAge)
    '**************************************************
    oCur.ModificaCurriculum Me.ctrRRHH.psCodigoPersona, Flex.GetRsNew, FlexExtra.GetRsNew, glsMovNro
    'ALPA 20090122*************************************
    If gTipoOpeRegistro = lnTipo Then
        gsOpeCod = LogPistaRegistrarCurriculum
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", , Me.ctrRRHH.psCodigoPersona, gCodigoCuenta
    ElseIf gTipoOpeMantenimiento = lnTipo Then
        gsOpeCod = LogPistaModificaCurriculum
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2", , Me.ctrRRHH.psCodigoPersona, gCodigoCuenta
    End If
    '**************************************************
    cmdCancelar_Click
End Sub

Private Sub cmdImprimir_Click()
    Dim oCur As NActualizaDatosCurriculum
    Dim oPrevio As Previo.clsPrevio
    Dim lsCadena As String
    Set oCur = New NActualizaDatosCurriculum
    Set oPrevio = New Previo.clsPrevio
    
    If Me.ctrRRHH.psCodigoPersona = "" Then
        MsgBox "Debe elegir a una persona.", vbInformation, "Aviso"
        Me.ctrRRHH.SetFocus
        Exit Sub
    End If
    
    lsCadena = oCur.GetReporteCurriculum(Me.ctrRRHH.psCodigoPersona, gsNomAge, gsEmpresa, gdFecSis)
    
    oPrevio.Show lsCadena, "Curriculum ", True, 66
    Set oPrevio = Nothing
End Sub

Private Sub CmdNuevo_Click()
    Flex.Col = 2
    Me.Flex.AdicionaFila
    Flex.Col = 1
    flex_RowColChange
    Me.Flex.SetFocus
End Sub

Private Sub cmdNuevoExtra_Click()
    Me.FlexExtra.Col = 2
    Me.FlexExtra.AdicionaFila
    FlexExtra.Col = 1
    FlexExtra_RowColChange
    Me.FlexExtra.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Activa(pbvalor As Boolean)
    Me.ctrRRHH.Enabled = Not pbvalor
    Me.Flex.lbEditarFlex = pbvalor
    Me.FlexExtra.lbEditarFlex = pbvalor
    Me.cmdEditar.Visible = Not pbvalor
    Me.cmdGrabar.Enabled = pbvalor
    Me.cmdCancelar.Visible = pbvalor
    Me.cmdSalir.Enabled = Not pbvalor
    
    If lnTipo = gTipoOpeRegistro Then
        Me.cmdEliminar.Enabled = False
        Me.cmdImprimir.Visible = False
    ElseIf lnTipo = gTipoOpeMantenimiento Then
        Me.cmdImprimir.Visible = False
    ElseIf lnTipo = gTipoOpeConsulta Then
        Me.cmdEliminar.Visible = False
        Me.cmdNuevo.Visible = False
        Me.cmdEditar.Visible = False
        Me.cmdImprimir.Visible = False
        Me.cmdGrabar.Visible = False
        Me.Flex.lbEditarFlex = False
        Me.fraCur.Enabled = True
    ElseIf lnTipo = gTipoOpeReporte Then
        Me.cmdEliminar.Enabled = pbvalor
        Me.cmdNuevo.Visible = False
        Me.cmdEliminar.Visible = False
        Me.cmdEditar.Enabled = False
    End If
End Sub

Private Sub ctrRRHH_EmiteDatos()
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    Set oRRHH = New DActualizaDatosRRHH
    Set oPersona = New UPersona
    Set oPersona = frmBuscaPersona.Inicio(True)
    If Not oPersona Is Nothing Then
        Limpia
        Me.ctrRRHH.psCodigoPersona = oPersona.sPersCod
        Me.ctrRRHH.psNombreEmpledo = oPersona.sPersNombre
        Me.ctrRRHH.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(Me.ctrRRHH.psCodigoPersona)
        CargaDatos
    End If
End Sub

Private Sub ctrRRHH_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oRRHH As DActualizaDatosRRHH
        Dim rsR As ADODB.Recordset
        Set oRRHH = New DActualizaDatosRRHH
        ctrRRHH.psCodigoEmpleado = Left(ctrRRHH.psCodigoEmpleado, 1) & Format(Trim(Mid(ctrRRHH.psCodigoEmpleado, 2)), "00000")
        Dim oCon As DActualizaDatosContrato
        Set oCon = New DActualizaDatosContrato
        
        Set rsR = oRRHH.GetRRHH(ctrRRHH.psCodigoEmpleado, gPersIdDNI)
           
        If Not (rsR.EOF And rsR.BOF) Then
            ctrRRHH.SpinnerValor = CInt(Right(ctrRRHH.psCodigoEmpleado, 5))
            ctrRRHH.psCodigoPersona = rsR.Fields("Codigo")
            ctrRRHH.psNombreEmpledo = rsR.Fields("Nombre")
            rsR.Close
            Set rsR = oRRHH.GetRRHHGeneralidades(ctrRRHH.psCodigoEmpleado)
            CargaDatos
            If cmdEditar.Enabled And cmdEditar.Visible Then
                Me.cmdEditar.SetFocus
            End If
        Else
            MsgBox "Codigo no Reconocido.", vbInformation, "Aviso"
            ctrRRHH.SetFocus
        End If
        
        rsR.Close
        Set rsR = Nothing
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub flex_RowColChange()
    Dim oCurT As DActualizaDatosCurriculum
    Dim rsC As ADODB.Recordset
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Set rsC = New ADODB.Recordset
    Set oCurT = New DActualizaDatosCurriculum
    
    If Flex.Col = 1 Then
        Set rsC = oCurT.GetCurriculumTabla(True)
        Flex.rsTextBuscar = rsC
    ElseIf Flex.Col = 4 Then
        Set rsC = oCon.GetConstante(6036, , , True)
        Flex.rsTextBuscar = rsC
    ElseIf Flex.Col = 8 Then
        Set rsC = oCon.GetConstante(gGenTipoPeriodos, , , True)
        Flex.rsTextBuscar = rsC
    ElseIf Flex.Col = 11 Then
        Set rsC = oCon.GetConstante(gRHProfesionesCurr, False, , True)
        Flex.rsTextBuscar = rsC
    ElseIf Flex.Col = 14 Then
        Set rsC = oCon.GetConstante(gRHNivelCurr, , , True)
        Flex.rsTextBuscar = rsC
    Else
        If Not IsNumeric(Me.Flex.TextMatrix(Flex.Row, 4)) Then
            Me.Flex.TextMatrix(Flex.Row, 4) = ""
        End If
        If Me.Flex.TextMatrix(Flex.Row, 4) = "" Then
            Flex.ColumnasAEditar = "X-1-X-3-4-X-X-X-X-X-X-11-X-13-14-X-16-17-X-X-X"
        ElseIf Me.Flex.TextMatrix(Flex.Row, 4) = RHPeriodosTpoTiempo Then
            Flex.ColumnasAEditar = "X-1-X-3-4-X-6-7-X-X-X-11-X-13-14-X-16-17-X-X-X"
        ElseIf Me.Flex.TextMatrix(Flex.Row, 4) = RHPeriodosTpoPeridos Then
            Flex.ColumnasAEditar = "X-1-X-3-4-X-X-X-8-X-10-11-X-13-14-X-16-17-X-X-X"
        ElseIf Me.Flex.TextMatrix(Flex.Row, 4) = RHPeriodosTpoTiempoPeridos Then
            Flex.ColumnasAEditar = "X-1-X-3-4-X-6-7-8-X-10-11-X-13-14-X-16-17-X-X-X"
        Else
            Flex.ColumnasAEditar = "X-1-X-3-4-X-X-X-X-X-X-11-X-13-14-X-16-17-X-X-X"
        End If
    End If
End Sub

Private Sub FlexExtra_RowColChange()
    Dim oCurT As DActualizaDatosCurriculum
    Dim rsC As ADODB.Recordset
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Set rsC = New ADODB.Recordset
    Set oCurT = New DActualizaDatosCurriculum
    
    If FlexExtra.Col = 1 Then
        Set rsC = oCon.GetConstante(6046, , , True)
        FlexExtra.rsTextBuscar = rsC
    ElseIf FlexExtra.Col = 3 Then
        If FlexExtra.TextMatrix(FlexExtra.Row, 1) = "" Then Exit Sub
        Set rsC = oCon.GetConstante(6047 + CInt(FlexExtra.TextMatrix(FlexExtra.Row, 1)), , , True)
        FlexExtra.rsTextBuscar = rsC
    ElseIf FlexExtra.Col = 7 Then
        Set rsC = oCon.GetConstante(6051, , , True)
        FlexExtra.rsTextBuscar = rsC
    End If
End Sub

Private Sub Form_Load()
    Me.SSTab.Tab = 0
    CargaDatos
    Activa False
    'ALPA 20090122 ***************************************************************************
    Set objPista = New COMManejador.Pista
    '*****************************************************************************************
End Sub

Private Sub CargaDatos()
    Dim oCurT As DActualizaDatosCurriculum
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Set oCurT = New DActualizaDatosCurriculum
    
    Set rsC = oCurT.GetCurriculums(Me.ctrRRHH.psCodigoPersona)
    If Not (rsC.EOF And rsC.BOF) Then
        'Flex.rsFlex = rsC
       Set Flex.Recordset = rsC
    Else
        Flex.Clear
        Flex.Rows = 2
        Flex.FormaCabecera
    End If
    
    Set rsC = oCurT.GetCurriculumsExtra(Me.ctrRRHH.psCodigoPersona)
    If Not (rsC.EOF And rsC.BOF) Then
        'FlexExtra.rsFlex = rsC
       Set FlexExtra.Recordset = rsC
    Else
        FlexExtra.Clear
        FlexExtra.Rows = 2
        FlexExtra.FormaCabecera
    End If
   
    Set oCurT = Nothing
    Set rsC = Nothing
End Sub

Private Sub Limpia()
    Me.Flex.Clear
    Me.Flex.Rows = 2
    Me.Flex.FormaCabecera
End Sub

Private Function Valida() As Boolean
    Dim i As Integer
    
    For i = 1 To Me.Flex.Rows - 1
        Flex.Row = i
        If Flex.TextMatrix(i, 1) = "" Then
            MsgBox "Debe Ingresar el tipo de curriculum.", vbInformation, "Aviso"
            Flex.Col = 1
            Flex.SetFocus
            Valida = False
            SSTab.Tab = 0
            Exit Function
        ElseIf Flex.TextMatrix(i, 3) = "" Then
            MsgBox "Debe Ingresar el lugar.", vbInformation, "Aviso"
            Flex.Col = 3
            Flex.SetFocus
            Valida = False
            SSTab.Tab = 0
            Exit Function
        ElseIf Flex.TextMatrix(i, 11) = "" Then
            MsgBox "Debe Ingresar la descripcion.", vbInformation, "Aviso"
            Flex.Col = 11
            Flex.SetFocus
            Valida = False
            SSTab.Tab = 0
            Exit Function
        ElseIf Flex.TextMatrix(i, 4) = "" Then
            MsgBox "Debe Ingresar un Tipo de Periodo valido.", vbInformation, "Aviso"
            Flex.Col = 4
            Flex.SetFocus
            Valida = False
            SSTab.Tab = 0
            Exit Function
        ElseIf Flex.TextMatrix(i, 4) = RHPeriodosTpoTiempo And (Not IsDate(Flex.TextMatrix(i, 6)) Or Not IsDate(Flex.TextMatrix(i, 7))) Then
            MsgBox "Debe Ingresar una fecha valida.", vbInformation, "Aviso"
            If Not IsDate(Flex.TextMatrix(i, 6)) Then
                Flex.Col = 6
            ElseIf Not IsDate(Flex.TextMatrix(i, 7)) Then
                Flex.Col = 7
            End If
            Flex.SetFocus
            Valida = False
            SSTab.Tab = 0
            Exit Function
        ElseIf Flex.TextMatrix(i, 4) = RHPeriodosTpoPeridos And (Flex.TextMatrix(i, 8) = "" Or Flex.TextMatrix(i, 9) = "") Then
            MsgBox "Debe Ingresar un periodo valido.", vbInformation, "Aviso"
            If Flex.TextMatrix(i, 8) = "" Then
                Flex.Col = 8
            ElseIf Flex.TextMatrix(i, 9) = "" Then
                Flex.Col = 9
            End If
            Flex.SetFocus
            Valida = False
            SSTab.Tab = 0
            Exit Function
        ElseIf Flex.TextMatrix(i, 13) = "" Then
            MsgBox "Debe Ingresar Costo.", vbInformation, "Aviso"
            Flex.Col = 13
            Flex.SetFocus
            Valida = False
            SSTab.Tab = 0
            Exit Function
        ElseIf Flex.TextMatrix(i, 15) = "" Then
            MsgBox "Debe Ingresar Nivel.", vbInformation, "Aviso"
            Flex.Col = 15
            Flex.SetFocus
            Valida = False
            SSTab.Tab = 0
            Exit Function
        Else
            Valida = True
        End If
    Next i
    
    ''Extra curricular
    'For i = 1 To Me.FlexExtra.Rows - 1
    '   FlexExtra.Row = i
    '    If FlexExtra.TextMatrix(i, 1) = "" Then
     '       MsgBox "Debe Ingresar el tipo de actividad extra curricular.", vbInformation, "Aviso"
     '       FlexExtra.Col = 1
     '       FlexExtra.SetFocus
     '       Valida = False
     '       SSTab.Tab = 1
     '       Exit Function
     '   ElseIf FlexExtra.TextMatrix(i, 3) = "" Then
     '       MsgBox "Debe Ingresar el codigo de la actividad extra curricular que realiza.", vbInformation, "Aviso"
     '       FlexExtra.Col = 3
     '       FlexExtra.SetFocus
     '       Valida = False
     '       SSTab.Tab = 1
     '       Exit Function
     '   ElseIf Not IsNumeric(FlexExtra.TextMatrix(i, 5)) Then
     '       MsgBox "Debe Ingresar numero valido.", vbInformation, "Aviso"
     '       FlexExtra.Col = 5
     '      FlexExtra.SetFocus
     '       Valida = False
     '       SSTab.Tab = 1
     '       Exit Function
     '   ElseIf Not IsNumeric(FlexExtra.TextMatrix(i, 6)) Then
     '       MsgBox "Debe Ingresar un Monto valido.", vbInformation, "Aviso"
     '       FlexExtra.Col = 6
     '       FlexExtra.SetFocus
     '       Valida = False
     '       SSTab.Tab = 1
     '       Exit Function
     '  ElseIf FlexExtra.TextMatrix(i, 7) = "" Then
     '       MsgBox "Debe Ingresar un nivel valido.", vbInformation, "Aviso"
     '       FlexExtra.Col = 7
     '       FlexExtra.SetFocus
     '       Valida = False
     '       SSTab.Tab = 1
     '       Exit Function
     '   ElseIf FlexExtra.TextMatrix(i, 10) = "" Then
     '       MsgBox "Debe Ingresar un comentario valido.", vbInformation, "Aviso"
     '       FlexExtra.Col = 10
     '       FlexExtra.SetFocus
     '       Valida = False
     '       SSTab.Tab = 1
     '       Exit Function
     '   Else
     '       Valida = True
     '   End If
    'Next i
    
End Function

Private Function GetUltCorr() As Long
    Dim lnMax As Integer
    Dim i As Integer
    
    If Me.Flex.Rows = 2 And Me.Flex.TextMatrix(Flex.Rows - 1, 2) = "" Then
        GetUltCorr = 1
    Else
        lnMax = CInt(Me.Flex.TextMatrix(Flex.Rows - 1, 2))
        For i = 2 To Me.Flex.Rows - 1
            If lnMax < CInt(Me.Flex.TextMatrix(i, 2)) Then
                lnMax = CInt(Me.Flex.TextMatrix(i, 2))
            End If
        Next i
        GetUltCorr = lnMax + 1
    End If
End Function

Public Sub Ini(pnTipo As TipoOpe, psCaption As String)
    lnTipo = pnTipo
    Caption = psCaption
    Me.Show 1
End Sub
