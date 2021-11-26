VERSION 5.00
Begin VB.Form frmRHAsistenciaAsig 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11430
   Icon            =   "frmRHAsistenciaAsig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEditar 
      Caption         =   "E&ditar"
      Height          =   375
      Left            =   1215
      TabIndex        =   10
      Top             =   5265
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   45
      TabIndex        =   9
      Top             =   5265
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10305
      TabIndex        =   6
      Top             =   5265
      Width           =   1095
   End
   Begin Sicmact.ctrRRHHGen ctrRRHHGen 
      Height          =   1200
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   11340
      _ExtentX        =   20003
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
   Begin VB.Frame fraHorarioLab 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Horario - Laboral"
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
      Height          =   3990
      Left            =   30
      TabIndex        =   1
      Top             =   1215
      Width           =   11355
      Begin VB.ComboBox cmbTipoHorario 
         Height          =   315
         ItemData        =   "frmRHAsistenciaAsig.frx":030A
         Left            =   2430
         List            =   "frmRHAsistenciaAsig.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3570
         Width           =   3615
      End
      Begin VB.CommandButton cmdIniciaHorario 
         Caption         =   "&Inicia Horario"
         Height          =   375
         Left            =   6120
         TabIndex        =   12
         Top             =   3540
         Width           =   1095
      End
      Begin VB.CommandButton cmdEliminarHorario 
         Caption         =   "E&liminar Hor."
         Height          =   375
         Left            =   1230
         TabIndex        =   8
         Top             =   3540
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregarHor 
         Caption         =   "Ag&regar Hor."
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Top             =   3540
         Width           =   1095
      End
      Begin VB.CommandButton cmdEliminarDia 
         Caption         =   "&Eliminar Dia"
         Height          =   375
         Left            =   8370
         TabIndex        =   5
         Top             =   3540
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregarDia 
         Caption         =   "&Agregar Dia"
         Height          =   375
         Left            =   7245
         TabIndex        =   4
         Top             =   3540
         Width           =   1095
      End
      Begin Sicmact.FlexEdit FlexDia 
         Height          =   3255
         Left            =   6120
         TabIndex        =   3
         Top             =   225
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   5741
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Cod Dia-Dia-Turno-Hora Ini-Hora Fin-Existe"
         EncabezadosAnchos=   "300-800-800-1200-800-800-0"
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
         ColumnasAEditar =   "X-1-X-3-4-5-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-3-2-2-0"
         EncabezadosAlineacion=   "C-L-L-L-R-R-C"
         FormatosEdit    =   "0-0-0-6-6-6-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit FlexHor 
         Height          =   3255
         Left            =   105
         TabIndex        =   2
         Top             =   225
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   5741
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Fecha-Rango-Comentario-Bit"
         EncabezadosAnchos=   "300-1200-600-3400-0"
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
         ColumnasAEditar =   "X-1-2-3-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-2-0-0-0"
         EncabezadosAlineacion=   "C-R-R-L-C"
         FormatosEdit    =   "0-0-3-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblHora 
         Caption         =   "Horas S."
         Height          =   255
         Left            =   9555
         TabIndex        =   14
         Top             =   3608
         Width           =   705
      End
      Begin VB.Label lblHoraSemana 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   10290
         TabIndex        =   13
         Top             =   3600
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1215
      TabIndex        =   11
      Top             =   5265
      Width           =   1095
   End
End
Attribute VB_Name = "frmRHAsistenciaAsig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnRowAnt As Integer
Dim lnTipo As TipoOpe
Dim LsCaption As String

Private Sub cmdAgregarDia_Click()
    If Not IsDate(Me.FlexHor.TextMatrix(FlexHor.Row, 1)) Then Exit Sub
    Me.FlexDia.AdicionaFila
    Me.FlexDia.TextMatrix(FlexDia.Row, 6) = Me.FlexHor.TextMatrix(FlexHor.Row, 1)
    FlexDia.SetFocus
End Sub

Private Sub cmdAgregarHor_Click()
    Me.FlexHor.AdicionaFila
    FlexHor_RowColChange
    FlexHor.SetFocus
    Me.FlexDia.Clear
    Me.FlexDia.Rows = 2
    Me.FlexDia.FormaCabecera
    FlexHor.ColumnasAEditar = "X-1-2-3-X"
End Sub

Private Sub Limpia(Optional pbTodos As Boolean = True)
    If pbTodos Then Me.ctrRRHHGen.ClearScreen
    FlexDia.Clear
    Me.FlexDia.Rows = 2
    FlexDia.FormaCabecera
    FlexHor.Clear
    Me.FlexHor.Rows = 2
    FlexHor.FormaCabecera
    lnRowAnt = -1
    Me.lblHoraSemana.Caption = "0.00"
End Sub

Private Sub cmdCancelar_Click()
    Limpia
    Activa False
End Sub

Private Sub cmdEditar_Click()
    If Me.ctrRRHHGen.psCodigoPersona = "" Then Exit Sub
    Activa True
    FlexHor_RowColChange
End Sub

Private Sub cmdEliminarDia_Click()
    If FlexDia.TextMatrix(FlexDia.Row, 6) = "" Then
        MsgBox "No puede borrar un registro grabado.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    FlexDia.EliminaFila FlexDia.Row
    
    Me.lblHoraSemana.Caption = Format(GetSuma(), "0#.00")
End Sub

Private Sub cmdEliminarHorario_Click()
    'If FlexHor.TextMatrix(FlexHor.Row, 3) = "1" Then
    '    MsgBox "No puede borrar un registro grabado.", vbInformation, "Aviso"
    '    Exit Sub
   ' End If
   If MsgBox("Desea Eliminar el horario seleccionado ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
   
    FlexHor.EliminaFila FlexHor.Row
    lnRowAnt = -1
    FlexHor_RowColChange
End Sub

Private Sub cmdGrabar_Click()
    If Not Valida Then Exit Sub
    
    If MsgBox("Desea Grabar ?? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    Dim oHor As DActualizaDatosHorarios
    Set oHor = New DActualizaDatosHorarios
    
    Dim rsH As ADODB.Recordset
    Dim rsT As ADODB.Recordset
    Set rsH = New ADODB.Recordset
    Set rsT = New ADODB.Recordset
    
    Set rsH = Me.FlexHor.GetRsNew
    Set rsT = Me.FlexDia.GetRsNew
    
    oHor.AgregaHorarios Me.ctrRRHHGen.psCodigoPersona, rsH, rsT, GetMovNro(gsCodUser, gsCodAge), gsFormatoFechaHora
    
    If Not rsT Is Nothing Then rsT.Close
    rsH.Close
    Set rsT = Nothing
    Set rsH = Nothing
    Limpia False
    CargaData Me.ctrRRHHGen.psCodigoPersona
    Activa False
End Sub

Private Sub cmdIniciaHorario_Click()
    Dim oHor As DActualizaDatosHorarios
    Set oHor = New DActualizaDatosHorarios
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If Not IsDate(Me.FlexHor.TextMatrix(FlexHor.Row, 1)) Then Exit Sub
    
    Set rs = oHor.GetHorarioTabla(Me.FlexHor.TextMatrix(FlexHor.Row, 1), Right(Me.cmbTipoHorario.Text, 3))
    
    Me.FlexDia.rsFlex = rs
    Me.lblHoraSemana.Caption = Format(GetSuma(), "0#.00")

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub ctrRRHHGen_EmiteDatos()
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    Set oRRHH = New DActualizaDatosRRHH
    Set oPersona = New UPersona
    Set oPersona = frmBuscaPersona.Inicio(True)
    If Not oPersona Is Nothing Then
        Me.ctrRRHHGen.psCodigoPersona = oPersona.sPersCod
        Me.ctrRRHHGen.psNombreEmpledo = oPersona.sPersNombre
        Me.ctrRRHHGen.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(Me.ctrRRHHGen.psCodigoPersona)
        CargaData Me.ctrRRHHGen.psCodigoPersona
    End If
End Sub

Private Sub ctrRRHHGen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oRRHH As DActualizaDatosRRHH
        Dim rsR As ADODB.Recordset
        Set oRRHH = New DActualizaDatosRRHH
        ctrRRHHGen.psCodigoEmpleado = Left(ctrRRHHGen.psCodigoEmpleado, 1) & Format(Trim(Mid(ctrRRHHGen.psCodigoEmpleado, 2)), "00000")
        Dim oCon As DActualizaDatosContrato
        Set oCon = New DActualizaDatosContrato
        
        Set rsR = oRRHH.GetRRHH(ctrRRHHGen.psCodigoEmpleado, gPersIdDNI)
           
        If Not (rsR.EOF And rsR.BOF) Then
            ctrRRHHGen.SpinnerValor = CInt(Right(ctrRRHHGen.psCodigoEmpleado, 5))
            ctrRRHHGen.psCodigoPersona = rsR.Fields("Codigo")
            ctrRRHHGen.psNombreEmpledo = rsR.Fields("Nombre")
            rsR.Close
            Set rsR = oRRHH.GetRRHHGeneralidades(ctrRRHHGen.psCodigoEmpleado)
            CargaData Me.ctrRRHHGen.psCodigoPersona
        Else
            MsgBox "Codigo no Reconocido.", vbInformation, "Aviso"
            ctrRRHHGen.SetFocus
        End If
        
        rsR.Close
        Set rsR = Nothing
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub FlexDia_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Me.lblHoraSemana.Caption = Format(GetSuma(), "0#.00")
End Sub

Private Sub FlexDia_RowColChange()
    If Not IsNumeric(Me.FlexDia.TextMatrix(FlexDia.Row, 1)) Then
        Me.FlexDia.TextMatrix(FlexDia.Row, 1) = ""
    End If
    
    Me.lblHoraSemana.Caption = Format(GetSuma(), "0#.00")
    
End Sub

Private Sub FlexHor_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    'Cancel = False
End Sub

Private Sub FlexHor_RowColChange()
    Dim rsD As ADODB.Recordset
    Dim oHor As DActualizaDatosHorarios
    
    If lnRowAnt <> FlexHor.Row Or lnRowAnt = -1 Then
        If Not IsDate(Me.FlexHor.TextMatrix(FlexHor.Row, 1)) Then
            lnRowAnt = FlexHor.Row
            Me.FlexDia.Clear
            Me.FlexDia.Rows = 2
            Me.FlexDia.FormaCabecera
            Exit Sub
        End If
        
        Set oHor = New DActualizaDatosHorarios
        Set rsD = New ADODB.Recordset
        
        Set rsD = oHor.GetHorariosDetalle(Me.ctrRRHHGen.psCodigoPersona, Format(CDate(Me.FlexHor.TextMatrix(FlexHor.Row, 1)), gsFormatoFecha))
        
        If Not (rsD.EOF And rsD.BOF) Then
            Set Me.FlexDia.Recordset = rsD
        Else
            FlexDia.Clear
            FlexDia.Rows = 2
            FlexDia.FormaCabecera
        End If
        lnRowAnt = FlexHor.Row
    End If
    
    If Me.FlexHor.TextMatrix(FlexHor.Row, 4) = "1" Then
        FlexHor.ColumnasAEditar = "X-X-X-X-X"
        Me.lblHoraSemana.Caption = Format(GetSuma(), "0#.00")
    Else
        FlexHor.ColumnasAEditar = "X-1-2-3-X"
    End If
End Sub

Private Sub Form_Load()
    Dim oCon As DConstantes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConstantes
    
    Set rs = oCon.GetConstante(1022, False, False, True)
    Me.FlexDia.rsTextBuscar = rs
    Me.FlexDia.CargaCombo oCon.GetConstante(gRHEmpleadoTurno)
    
    CargaCombo oCon.GetConstante(6041), Me.cmbTipoHorario
    cmbTipoHorario.ListIndex = 0
    Activa False
    Limpia
    Caption = LsCaption
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.cmdEditar.Visible = Not pbValor
    Me.cmdCancelar.Visible = pbValor
    Me.cmdGrabar.Enabled = pbValor
    Me.cmdSalir.Visible = Not pbValor
    Me.fraHorarioLab.Enabled = pbValor
    Me.ctrRRHHGen.Enabled = Not pbValor
    
    If lnTipo = gTipoOpeConsulta Then
        Me.cmdAgregarDia.Visible = False
        Me.cmdAgregarHor.Visible = False
        Me.cmdCancelar.Visible = False
        Me.cmdEditar.Visible = False
        Me.cmdEliminarDia.Visible = False
        Me.cmdEliminarHorario.Visible = False
        Me.cmdGrabar.Visible = False
        Me.fraHorarioLab.Enabled = True
        Me.FlexDia.lbEditarFlex = False
        Me.cmdIniciaHorario.Visible = False
    ElseIf lnTipo = gTipoOpeMantenimiento Then
            
    End If
End Sub

Private Sub CargaData(psPersCod As String)
    Dim rsD As ADODB.Recordset
    Dim oHor As DActualizaDatosHorarios
    Set oHor = New DActualizaDatosHorarios
    Set rsD = New ADODB.Recordset
    
    Set rsD = oHor.GetHorarios(psPersCod)
    If Not (rsD.EOF And rsD.BOF) Then
        Set Me.FlexHor.Recordset = rsD
        FlexHor.Col = 0
    Else
        FlexHor.Clear
        FlexHor.Rows = 2
        FlexHor.FormaCabecera
    End If
    FlexHor_RowColChange
End Sub

Private Function Valida() As Boolean
    Dim i As Integer
    
    For i = 1 To FlexHor.Rows - 1
        If Not IsDate(FlexHor.TextMatrix(i, 1)) Then
            MsgBox " Debe Ingresar una Fecha Valida para el Item = " & Me.FlexHor.TextMatrix(i, 0) & " .", vbInformation, "Aviso"
            FlexHor.Col = 1
            FlexHor.Row = i
            FlexHor.SetFocus
            Valida = False
            Exit Function
        ElseIf Not IsNumeric(FlexHor.TextMatrix(i, 2)) Then
            MsgBox " Debe Ingresar una rango valido para el Item = " & Me.FlexHor.TextMatrix(i, 0) & " .", vbInformation, "Aviso"
            FlexHor.Col = 2
            FlexHor.Row = i
            FlexHor.SetFocus
            Valida = False
            Exit Function
        ElseIf FlexHor.TextMatrix(i, 3) = "" Then
            MsgBox " Debe Ingresar un comentario valido para el Item = " & Me.FlexHor.TextMatrix(i, 0) & " .", vbInformation, "Aviso"
            FlexHor.Col = 3
            FlexHor.Row = i
            FlexHor.SetFocus
            Valida = False
            Exit Function
        End If
    Next i
    
    For i = 1 To FlexDia.Rows - 1
        If FlexDia.TextMatrix(i, 0) = "" Then
            
        ElseIf FlexDia.TextMatrix(i, 1) = "" Then
            MsgBox " Debe Ingresar una dia laboral para el Item = " & Me.FlexDia.TextMatrix(i, 0) & " .", vbInformation, "Aviso"
            FlexDia.SetFocus
            Valida = False
            Exit Function
        ElseIf FlexDia.TextMatrix(i, 3) = "" Then
            MsgBox " Debe Ingresar un Turno Valido para el Item = " & Me.FlexDia.TextMatrix(i, 0) & " .", vbInformation, "Aviso"
            FlexDia.SetFocus
            Valida = False
            Exit Function
        ElseIf Not IsDate(FlexDia.TextMatrix(i, 4)) Then
            MsgBox " Debe Ingresar una hora de Inicio Valida para el Item = " & Me.FlexDia.TextMatrix(i, 0) & " .", vbInformation, "Aviso"
            FlexDia.SetFocus
            Valida = False
            Exit Function
        ElseIf Not IsDate(FlexDia.TextMatrix(i, 5)) Then
            MsgBox " Debe Ingresar una hora de Salida Valida para el Item = " & Me.FlexDia.TextMatrix(i, 0) & " .", vbInformation, "Aviso"
            FlexDia.SetFocus
            Valida = False
            Exit Function
        End If
    Next i
    Valida = True
End Function

Public Function Ini(pnTipo As TipoOpe, psCaption As String)
    lnTipo = pnTipo
    LsCaption = psCaption
    Me.Show 1
End Function

Private Function GetSuma() As Double
    Dim lnI As Integer
    Dim lnSuma As Double
    
    lnSuma = 0
    For lnI = 1 To Me.FlexDia.Rows - 1
        If IsDate(Me.FlexDia.TextMatrix(lnI, 4)) And IsDate(Me.FlexDia.TextMatrix(lnI, 5)) Then
            lnSuma = lnSuma + DateDiff("n", CDate(Me.FlexDia.TextMatrix(lnI, 4)), CDate(Me.FlexDia.TextMatrix(lnI, 5)))
        End If
    Next lnI
    
    GetSuma = lnSuma / 60
End Function


