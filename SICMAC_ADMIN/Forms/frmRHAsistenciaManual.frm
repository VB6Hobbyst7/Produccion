VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRHAsistenciaManual 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12060
   Icon            =   "frmRHAsistenciaManual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdtardanza 
      Caption         =   "Tardanzas"
      Height          =   375
      Left            =   2505
      TabIndex        =   20
      Top             =   6705
      Width           =   1335
   End
   Begin VB.CommandButton cmdReiniciar 
      Caption         =   "&Reiniciar"
      Height          =   375
      Left            =   8115
      TabIndex        =   16
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   9285
      TabIndex        =   15
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdAgregaAsist 
      Caption         =   "&Add Asist"
      Height          =   375
      Left            =   6945
      TabIndex        =   14
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   10740
      TabIndex        =   13
      Top             =   615
      Width           =   1095
   End
   Begin VB.Frame fraAgencias 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Agencias"
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
      Height          =   615
      Left            =   45
      TabIndex        =   8
      Top             =   435
      Width           =   11925
      Begin Sicmact.TxtBuscar TxtAgencia 
         Height          =   285
         Left            =   1065
         TabIndex        =   10
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   930
      End
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2460
         TabIndex        =   11
         Top             =   195
         Width           =   8085
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10890
      TabIndex        =   7
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1230
      TabIndex        =   4
      Top             =   6720
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dpkFecha 
      Height          =   315
      Left            =   675
      TabIndex        =   3
      Top             =   75
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      _Version        =   393216
      Format          =   67239937
      CurrentDate     =   37056
   End
   Begin VB.Frame fraAsistencia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Asistencia"
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
      Height          =   5505
      Left            =   45
      TabIndex        =   0
      Top             =   1155
      Width           =   11970
      Begin Sicmact.FlexEdit FlexHorario 
         Height          =   5145
         Left            =   6960
         TabIndex        =   2
         Top             =   240
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   9075
         Cols0           =   9
         HighLight       =   1
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Turno-Ingreso-Salida-CodPersona-CodLleno-Codigo-Ingreso1-Salida1"
         EncabezadosAnchos=   "500-800-1400-1400-0-0-900-0-0"
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
         ColumnasAEditar =   "X-1-2-3-4-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-3-2-2-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-C-C-L-C-C"
         FormatosEdit    =   "0-0-5-5-0-0-0-5-5"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit flexTurno 
         Height          =   5145
         Left            =   75
         TabIndex        =   1
         Top             =   240
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   9075
         Cols0           =   9
         HighLight       =   1
         EncabezadosNombres=   "#-Cod...-Nombre-T.1 Ini-T.1 Fin-T.2 Ini-T.2 Fin-Indice-PersCod"
         EncabezadosAnchos=   "400-600-2400-750-750-750-750-0-0"
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
         ColumnasAEditar =   "X-X-2-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R-R-R-R-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1230
      TabIndex        =   6
      Top             =   6720
      Width           =   1095
   End
   Begin Sicmact.TxtBuscar TxtTipoPlanilla 
      Height          =   315
      Left            =   3735
      TabIndex        =   17
      Top             =   75
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
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
   Begin VB.Label lblTipCon 
      Caption         =   "Tipo Contrato"
      Height          =   255
      Left            =   2580
      TabIndex        =   19
      Top             =   105
      Width           =   1125
   End
   Begin VB.Label lblTipoPlanillaRes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5190
      TabIndex        =   18
      Top             =   90
      Width           =   6765
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha :"
      Height          =   210
      Left            =   60
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmRHAsistenciaManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnTipo As TipoOpe
Dim lbPermiso As Boolean

Private Sub Activa(pbValor As Boolean)
    Me.cmdSalir.Enabled = Not pbValor
    Me.cmdGrabar.Enabled = pbValor
    Me.cmdCancelar.Visible = pbValor
    Me.cmdAplicar.Enabled = Not pbValor
    Me.cmdEditar.Visible = Not pbValor
    Me.dpkFecha.Enabled = Not pbValor
    Me.fraAgencias.Enabled = Not pbValor
    
    
    If lnTipo = gTipoOpeConsulta Then
        Me.cmdCancelar.Enabled = True
        Me.cmdEditar.Visible = False
        Me.cmdGrabar.Visible = False
        Me.cmdSalir.Enabled = True
        Me.FlexHorario.lbEditarFlex = False
        Me.cmdAplicar.Enabled = True
        
    End If
End Sub

Private Sub Limpia()
    Me.FlexHorario.Clear
    Me.FlexHorario.Rows = 2
    Me.FlexHorario.FormaCabecera
    Me.flexTurno.Clear
    Me.flexTurno.Rows = 2
    Me.flexTurno.FormaCabecera
End Sub

Private Sub chkTodos_Click()

If gsCodArea = "022" Or gsCodArea = "044" Then
    Else
    MsgBox "Solo puede ver datos de la agencia asignada", vbInformation, "No tiene los permisos"
    Me.chkTodos.value = 0
    Exit Sub
End If

If Me.chkTodos.value = 1 Then
    Me.TxtAgencia.Text = ""
    Me.lblAgencia.Caption = ""
End If

End Sub

Private Sub cmdAgregaAsist_Click()
    Dim lnI As Long
    Dim lnJ As Long
    Dim lnBan1 As Boolean
    Dim lnBan2 As Boolean
    
    For lnI = 1 To Me.flexTurno.Rows - 1
        'Agrega turno 1
        If Me.flexTurno.TextMatrix(lnI, 3) <> "" Then
            lnBan1 = True
            For lnJ = CInt(Me.flexTurno.TextMatrix(lnI, 8)) To Me.FlexHorario.Rows - 1
                 If FlexHorario.TextMatrix(lnJ, 6) = Me.flexTurno.TextMatrix(lnI, 1) And lnBan1 Then
                    If FlexHorario.TextMatrix(lnJ, 1) = "TURNO 1                                                  1" Then
                        lnBan1 = False
                    End If
                Else
                    lnJ = Me.FlexHorario.Rows - 1
                End If
            Next lnJ
            
            If lnBan1 Then
                For lnJ = CInt(Me.flexTurno.TextMatrix(lnI, 8)) To Me.FlexHorario.Rows - 1
                     If FlexHorario.TextMatrix(lnJ, 6) = Me.flexTurno.TextMatrix(lnI, 1) And lnBan1 Then
                        If FlexHorario.TextMatrix(lnJ, 1) = "" Then
                            lnBan1 = False
                            FlexHorario.TextMatrix(lnJ, 1) = "TURNO 1                                                  1"
                            FlexHorario.TextMatrix(lnJ, 2) = Me.dpkFecha & " " & Me.flexTurno.TextMatrix(lnI, 3)
                            FlexHorario.TextMatrix(lnJ, 3) = Me.dpkFecha & " " & Me.flexTurno.TextMatrix(lnI, 4)
                        End If
                    Else
                        lnJ = Me.FlexHorario.Rows - 1
                    End If
                Next lnJ
            End If
        End If
    
        'Agrega turno 3
        If Me.flexTurno.TextMatrix(lnI, 5) <> "" Then
            lnBan1 = True
            For lnJ = CInt(Me.flexTurno.TextMatrix(lnI, 8)) To Me.FlexHorario.Rows - 1
                 If FlexHorario.TextMatrix(lnJ, 6) = Me.flexTurno.TextMatrix(lnI, 1) And lnBan1 Then
                    If FlexHorario.TextMatrix(lnJ, 1) = "TURNO 2                                                  2" Then
                        lnBan1 = False
                    End If
                Else
                    lnJ = Me.FlexHorario.Rows - 1
                End If
            Next lnJ
            
            If lnBan1 Then
                For lnJ = CInt(Me.flexTurno.TextMatrix(lnI, 8)) To Me.FlexHorario.Rows - 1
                     If FlexHorario.TextMatrix(lnJ, 6) = Me.flexTurno.TextMatrix(lnI, 1) And lnBan1 Then
                        If FlexHorario.TextMatrix(lnJ, 1) = "" Then
                            lnBan1 = False
                            FlexHorario.TextMatrix(lnJ, 1) = "TURNO 2                                                  2"
                            FlexHorario.TextMatrix(lnJ, 2) = Me.dpkFecha & " " & Me.flexTurno.TextMatrix(lnI, 5)
                            FlexHorario.TextMatrix(lnJ, 3) = Me.dpkFecha & " " & Me.flexTurno.TextMatrix(lnI, 6)
                        End If
                    Else
                        lnJ = Me.FlexHorario.Rows - 1
                    End If
                Next lnJ
            End If
        End If
    Next lnI
End Sub

Private Sub cmdAplicar_Click()
   
    GetData
    If Not (Me.chkTodos.value = 0 And TxtAgencia.Text = "") Then Activa True
End Sub

Private Sub cmdCancelar_Click()
    Limpia
    Activa False
    If Not lbPermiso Then
        Unload Me
    End If
End Sub

Private Sub cmdEditar_Click()
    Activa True
End Sub

Private Sub cmdGrabar_Click()
    Dim oAsis As DActualizaDatosHorarios
    Dim rsA As ADODB.Recordset
    Set rsA = New ADODB.Recordset
    Set oAsis = New DActualizaDatosHorarios
    
    
    Set rsA = Me.FlexHorario.GetRsNew
        
    If rsA Is Nothing Then
        Exit Sub
    End If
        
    oAsis.AgregaAsistencia FechaHora(CDate(Me.dpkFecha.value)), rsA, GetMovNro(gsCodUser, gsCodAge), gsFormatoFechaHora
    
    Set rsA = Nothing
    Set oAsis = Nothing
    Limpia
    Activa False
    If Not lbPermiso Then
        Unload Me
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    Dim oHor As DActualizaDatosHorarios
    Set oHor = New DActualizaDatosHorarios
    Dim lsCadena As String
    
    If Me.TxtTipoPlanilla.Text = "" Then
        MsgBox "Debe ingresar un tipo contrato a especificar en el control de Asistencia.", vbInformation, "Aviso"
        Me.TxtTipoPlanilla.SetFocus
        Exit Sub
    ElseIf Me.chkTodos.value = 0 And TxtAgencia.Text = "" Then
        MsgBox "Si no desea ver todas las agencias debe elegir una.", vbInformation, "Aviso"
        Me.TxtAgencia.SetFocus
        Exit Sub
    End If
    
    lsCadena = oHor.GetRepoAsistenciaDiaAge(Me.dpkFecha, Me.TxtAgencia.Text, Me.lblAgencia.Caption, gcEmpresa, Me.dpkFecha, Me.TxtTipoPlanilla.Text)
    
    oPrevio.Show lsCadena, Caption, True
End Sub

Private Sub cmdReiniciar_Click()
    Dim lnI As Long
    Dim lnJ As Long
    Dim lnBan1 As Boolean
    Dim lnBan2 As Boolean
    
    For lnI = 1 To Me.FlexHorario.Rows - 1
        FlexHorario.TextMatrix(lnI, 1) = ""
    Next lnI
        
    cmdAgregaAsist_Click

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdtardanza_Click()
frmRHReporteAsistencia.Show 1
End Sub

Private Sub FlexHorario_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If Me.flexTurno.TextMatrix(CInt(Me.FlexHorario.TextMatrix(pnRow, 5)), 5) = "" And Right(Me.FlexHorario.TextMatrix(pnRow, 0), 1) = RHEmpleadoTurnoDos Then
        FlexHorario.TextMatrix(pnRow, pnCol) = ""
        Cancel = False
    End If
End Sub

Private Sub flexTurno_RowColChange()
    Dim lnI As Integer
        
    For lnI = 1 To Me.FlexHorario.Rows - 1
         If FlexHorario.TextMatrix(lnI, 6) = Me.flexTurno.TextMatrix(Me.flexTurno.Row, 1) Then
            FlexHorario.RowHeight(lnI) = 285
        Else
            FlexHorario.RowHeight(lnI) = 0
        End If
    Next lnI
    
    If flexTurno.TextMatrix(flexTurno.Row, 8) <> "" Then
        Me.FlexHorario.Row = flexTurno.TextMatrix(flexTurno.Row, 8)
    End If
    
    FlexHorario.Col = 0
End Sub

Private Sub Form_Load()
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Me.FlexHorario.CargaCombo oCon.GetConstante(gRHEmpleadoTurno)
    Me.TxtAgencia.rs = oCon.GetAgencias(, , True)

    Activa False
   
    Set rs = oCon.GetRHTipoContrato
    TxtTipoPlanilla.rs = rs
    Set oCon = Nothing
    lbPermiso = True
    If Not TieneAcceso(gsPermisoHorario) Then
        lbPermiso = False
        Me.dpkFecha = gdFecSis
        Me.dpkFecha.Enabled = False
        Me.TxtAgencia.Text = gsCodAge
        Me.TxtAgencia.Enabled = False
        txtAgencia_EmiteDatos
        Me.chkTodos.value = 0
        Me.cmdAplicar.Enabled = False
        cmdAplicar_Click
    End If
End Sub

Private Sub GetData()
    Dim rsH As ADODB.Recordset
    Dim i As Integer
    Set rsH = New ADODB.Recordset
    Dim rsD As ADODB.Recordset
    Set rsD = New ADODB.Recordset
    Dim oHor As DActualizaDatosHorarios
    Set oHor = New DActualizaDatosHorarios
    Dim lbBan As Boolean
    Dim lbBandera As Boolean

    If Me.TxtTipoPlanilla.Text = "" Then
        MsgBox "Debe ingresar un tipo contrato a especificar en el control de Asistencia.", vbInformation, "Aviso"
        Me.TxtTipoPlanilla.SetFocus
        Exit Sub
    ElseIf Me.chkTodos.value = 0 And TxtAgencia.Text = "" Then
        MsgBox "Si no desea ver todas las agencias debe elegir una.", vbInformation, "Aviso"
        Me.TxtAgencia.SetFocus
        Exit Sub
    End If

    Set rsH = oHor.GetHorarioDia(CDate(Me.dpkFecha.value), gsFormatoFecha, IIf(Me.chkTodos.value = 1, "", Me.TxtAgencia.Text), Me.TxtTipoPlanilla.Text)
    If Not (rsH.EOF And rsH.BOF) Then
        FlexHorario.EnumeraItems

        Set Me.flexTurno.Recordset = rsH
        FlexHorario.Clear
        FlexHorario.Rows = 2
        FlexHorario.FormaCabecera
        lbBandera = True
        For i = 1 To flexTurno.Rows - 1
            Set rsD = oHor.GetHorarioDiaDet(Format(CDate(Me.dpkFecha.value), gsFormatoFecha), Me.flexTurno.TextMatrix(i, 7))

            If Not (rsD.EOF And rsD.EOF) Then
                If Me.flexTurno.TextMatrix(i, 7) = rsD.Fields(0) Then
                        Me.flexTurno.TextMatrix(i, 8) = FlexHorario.Rows - 1
                        lbBan = True
                        FlexHorario.AdicionaFila
                        While Not rsD.EOF And lbBan
                            If rsD.Fields(0) = Me.flexTurno.TextMatrix(i, 7) Then
                                If Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 1) <> "" Then
                                    FlexHorario.AdicionaFila , , True
                                    'FlexHorario.Rows = FlexHorario.Rows + 1
                                    'FlexHorario.ColWidth(FlexHorario.Rows - 1) = FlexHorario.ColWidth(1)
                                End If
                                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 1) = rsD.Fields(1)
                                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 2) = IIf(IsNull(rsD.Fields(2)), "__/__/____", rsD.Fields(2))
                                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 3) = IIf(IsNull(rsD.Fields(3)), "__/__/____", rsD.Fields(3))
                                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 4) = Me.flexTurno.TextMatrix(i, 7)
                                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 5) = i
                                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 6) = Me.flexTurno.TextMatrix(i, 1)
                                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 7) = IIf(IsNull(rsD.Fields(2)), "__/__/____", rsD.Fields(2))
                                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 8) = IIf(IsNull(rsD.Fields(3)), "__/__/____", rsD.Fields(3))

                                rsD.MoveNext
                            Else
                                lbBan = False
                            End If
                        Wend
                   lbBandera = True
                Else
                    Me.flexTurno.TextMatrix(i, 8) = Me.FlexHorario.Rows - 1
                End If
            Else
                FlexHorario.AdicionaFila
                Me.flexTurno.TextMatrix(i, 8) = Me.FlexHorario.Rows - 1
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 4) = Me.flexTurno.TextMatrix(i, 7)
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 5) = i
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 6) = Me.flexTurno.TextMatrix(i, 1)
            End If

            If lnTipo = gTipoOpeMantenimiento And lbBandera Then
                FlexHorario.AdicionaFila , , True
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 4) = Me.flexTurno.TextMatrix(i, 7)
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 5) = i
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 6) = Me.flexTurno.TextMatrix(i, 1)
                FlexHorario.AdicionaFila , , True
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 4) = Me.flexTurno.TextMatrix(i, 7)
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 5) = i
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 6) = Me.flexTurno.TextMatrix(i, 1)
                FlexHorario.AdicionaFila , , True
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 4) = Me.flexTurno.TextMatrix(i, 7)
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 5) = i
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 6) = Me.flexTurno.TextMatrix(i, 1)
                FlexHorario.AdicionaFila , , True
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 4) = Me.flexTurno.TextMatrix(i, 7)
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 5) = i
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 6) = Me.flexTurno.TextMatrix(i, 1)
                FlexHorario.AdicionaFila , , True
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 4) = Me.flexTurno.TextMatrix(i, 7)
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 5) = i
                Me.FlexHorario.TextMatrix(FlexHorario.Rows - 1, 6) = Me.flexTurno.TextMatrix(i, 1)
            End If
        Next i

    Else
        flexTurno.Clear
        flexTurno.Rows = 2
        flexTurno.FormaCabecera
        FlexHorario.Clear
        FlexHorario.Rows = 2
        FlexHorario.FormaCabecera
    End If

    Set rsH = Nothing
    Set oHor = Nothing
    Set rsD = Nothing
End Sub



Private Sub txtAgencia_EmiteDatos()
    Me.lblAgencia.Caption = TxtAgencia.psDescripcion
    
    If TxtAgencia.Text <> gsCodAge Then
        If gsCodArea = "022" Or gsCodArea = "044" Then
           Else
                MsgBox "Solo esta permitido ver asistencia de la agencia asignada", vbInformation, "no esta permitido"
                Me.lblAgencia.Caption = ""
                TxtAgencia.Text = ""
                Exit Sub
        End If
    End If
    
    
End Sub

Public Sub Ini(pnTipo As TipoOpe, psCaption As String)
    lnTipo = pnTipo
    Caption = psCaption
    Me.dpkFecha = Format(gdFecSis, "dd/mm/yyyy")
    Me.Show 1
End Sub

