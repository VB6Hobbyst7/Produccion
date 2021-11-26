VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmRHPeriodosNoLaborales 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "frmRHPeriodosNoLaborales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActEstado 
      Caption         =   "&Act Estado"
      Height          =   375
      Left            =   6510
      TabIndex        =   36
      Top             =   5655
      Width           =   1095
   End
   Begin TabDlg.SSTab Tab 
      Height          =   4335
      Left            =   45
      TabIndex        =   1
      Top             =   1230
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Vacaciones"
      TabPicture(0)   =   "frmRHPeriodosNoLaborales.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdCancelar(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdEditar(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdImprimir(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraPerNoLab(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdGrabar(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Permisos"
      TabPicture(1)   =   "frmRHPeriodosNoLaborales.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCancelar(1)"
      Tab(1).Control(1)=   "cmdGrabar(1)"
      Tab(1).Control(2)=   "cmdEditar(1)"
      Tab(1).Control(3)=   "fraPerNoLab(1)"
      Tab(1).Control(4)=   "cmdImprimir(1)"
      Tab(1).Control(5)=   "chkSoloPendientes(1)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Descansos/Subsidios"
      TabPicture(2)   =   "frmRHPeriodosNoLaborales.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "cmdCancelar(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdGrabar(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraPerNoLab(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdImprimir(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdEditar(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Sanciones"
      TabPicture(3)   =   "frmRHPeriodosNoLaborales.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdCancelar(3)"
      Tab(3).Control(1)=   "cmdGrabar(3)"
      Tab(3).Control(2)=   "fraPerNoLab(3)"
      Tab(3).Control(3)=   "cmdImprimir(3)"
      Tab(3).Control(4)=   "cmdEditar(3)"
      Tab(3).ControlCount=   5
      Begin VB.CheckBox chkSoloPendientes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Pendientes"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   -69540
         TabIndex        =   24
         Top             =   3915
         Width           =   2880
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   3
         Left            =   -73710
         TabIndex        =   20
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   3
         Left            =   -72510
         TabIndex        =   19
         Top             =   3825
         Width           =   1095
      End
      Begin VB.Frame fraPerNoLab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Sanciones"
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
         Height          =   3405
         Index           =   3
         Left            =   -74925
         TabIndex        =   18
         Top             =   360
         Width           =   8565
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   375
            Index           =   3
            Left            =   7410
            TabIndex        =   32
            Top             =   2955
            Width           =   1095
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   375
            Index           =   3
            Left            =   6240
            TabIndex        =   31
            Top             =   2955
            Width           =   1095
         End
         Begin Sicmact.FlexEdit FlexPerNoLab 
            Height          =   2625
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Top             =   255
            Width           =   8400
            _extentx        =   14817
            _extenty        =   4630
            cols0           =   13
            highlight       =   1
            encabezadosnombres=   "#-CodTipo-Tipo-Sol Ini-Sol Fin-Ejec Ini-Ejec Fin-Comentario-Cod.Estad-Estado-Observacion-bit1-bit2"
            encabezadosanchos=   "300-800-1500-1500-1500-0-0-3000-0-0-0-0-0"
            font            =   "frmRHPeriodosNoLaborales.frx":037A
            font            =   "frmRHPeriodosNoLaborales.frx":03A2
            font            =   "frmRHPeriodosNoLaborales.frx":03CA
            font            =   "frmRHPeriodosNoLaborales.frx":03F2
            font            =   "frmRHPeriodosNoLaborales.frx":041A
            fontfixed       =   "frmRHPeriodosNoLaborales.frx":0442
            columnasaeditar =   "X-1-X-3-4-5-6-7-8-X-10-X-X"
            textstylefixed  =   3
            listacontroles  =   "0-1-0-2-2-2-2-0-1-0-0-0-0"
            encabezadosalineacion=   "C-L-L-R-R-R-R-L-L-L-L-C-C"
            formatosedit    =   "0-0-0-5-5-5-5-0-0-0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1  'True
            appearance      =   0
            colwidth0       =   300
            rowheight0      =   300
            cellbackcolor   =   -2147483624
         End
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   2
         Left            =   1290
         TabIndex        =   15
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   2
         Left            =   2490
         TabIndex        =   14
         Top             =   3825
         Width           =   1095
      End
      Begin VB.Frame fraPerNoLab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Descansos"
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
         Height          =   3405
         Index           =   2
         Left            =   75
         TabIndex        =   13
         Top             =   360
         Width           =   8565
         Begin Sicmact.FlexEdit FlexPerNoLab 
            Height          =   2625
            Index           =   2
            Left            =   120
            TabIndex        =   34
            Top             =   255
            Width           =   8400
            _extentx        =   14817
            _extenty        =   4630
            cols0           =   13
            highlight       =   1
            encabezadosnombres=   "#-CodTipo-Tipo-Sol Ini-Sol Fin-Ejec Ini-Ejec Fin-Comentario-Cod Estado-Estado-Observacion-bit1-bit2"
            encabezadosanchos=   "300-800-1500-1500-1500-0-0-3000-0-0-0-0-0"
            font            =   "frmRHPeriodosNoLaborales.frx":0468
            font            =   "frmRHPeriodosNoLaborales.frx":0490
            font            =   "frmRHPeriodosNoLaborales.frx":04B8
            font            =   "frmRHPeriodosNoLaborales.frx":04E0
            font            =   "frmRHPeriodosNoLaborales.frx":0508
            fontfixed       =   "frmRHPeriodosNoLaborales.frx":0530
            columnasaeditar =   "X-1-X-3-4-5-6-7-8-X-10-X-X"
            textstylefixed  =   3
            listacontroles  =   "0-1-0-2-2-2-2-0-1-0-0-0-0"
            encabezadosalineacion=   "C-L-L-R-R-R-R-L-L-L-L-C-C"
            formatosedit    =   "0-0-0-5-5-5-5-0-0-0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1  'True
            appearance      =   0
            colwidth0       =   300
            rowheight0      =   300
            cellbackcolor   =   -2147483624
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   375
            Index           =   2
            Left            =   7410
            TabIndex        =   30
            Top             =   2955
            Width           =   1095
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   375
            Index           =   2
            Left            =   6240
            TabIndex        =   29
            Top             =   2955
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   0
         Left            =   -74910
         TabIndex        =   6
         Top             =   3825
         Width           =   1095
      End
      Begin VB.Frame fraPerNoLab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vacaciones"
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
         Height          =   3405
         Index           =   0
         Left            =   -74925
         TabIndex        =   2
         Top             =   360
         Width           =   8565
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   375
            Index           =   0
            Left            =   7410
            TabIndex        =   26
            Top             =   2955
            Width           =   1095
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   375
            Index           =   0
            Left            =   6240
            TabIndex        =   25
            Top             =   2955
            Width           =   1095
         End
         Begin Sicmact.FlexEdit FlexPerNoLab 
            Height          =   2625
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   255
            Width           =   8400
            _extentx        =   14817
            _extenty        =   4630
            cols0           =   13
            highlight       =   1
            encabezadosnombres=   "#-CodTipo-Tipo-Sol Ini-Sol Fin-Ejec Ini-Ejec Fin-Comentario-Cod Estado-Estado-Observacion-bit1-bit2"
            encabezadosanchos=   "300-800-2500-1500-1500-1500-1500-3000-0-0-0-0-0"
            font            =   "frmRHPeriodosNoLaborales.frx":0556
            font            =   "frmRHPeriodosNoLaborales.frx":057E
            font            =   "frmRHPeriodosNoLaborales.frx":05A6
            font            =   "frmRHPeriodosNoLaborales.frx":05CE
            font            =   "frmRHPeriodosNoLaborales.frx":05F6
            fontfixed       =   "frmRHPeriodosNoLaborales.frx":061E
            columnasaeditar =   "X-1-X-3-4-5-6-7-8-X-10-X-X"
            textstylefixed  =   3
            listacontroles  =   "0-1-0-2-2-2-2-0-1-0-0-0-0"
            encabezadosalineacion=   "C-L-L-R-R-R-R-L-L-L-L-C-C"
            formatosedit    =   "0-0-0-5-5-5-5-0-0-0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            appearance      =   0
            colwidth0       =   300
            rowheight0      =   300
            cellbackcolor   =   -2147483624
         End
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   0
         Left            =   -72510
         TabIndex        =   4
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   1
         Left            =   -72510
         TabIndex        =   9
         Top             =   3825
         Width           =   1095
      End
      Begin VB.Frame fraPerNoLab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Permisos"
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
         Height          =   3405
         Index           =   1
         Left            =   -74925
         TabIndex        =   8
         Top             =   360
         Width           =   8565
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   375
            Index           =   1
            Left            =   6240
            TabIndex        =   28
            Top             =   2955
            Width           =   1095
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   375
            Index           =   1
            Left            =   7410
            TabIndex        =   27
            Top             =   2955
            Width           =   1095
         End
         Begin Sicmact.FlexEdit FlexPerNoLab 
            Height          =   2625
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   255
            Width           =   8400
            _extentx        =   14817
            _extenty        =   4630
            cols0           =   13
            highlight       =   1
            encabezadosnombres=   "#-CodTipo-Tipo-Sol Ini-Sol Fin-Ejec Ini-Ejec Fin-Comentario-Cod Estado-Estado-Observacion-bit1-bit2"
            encabezadosanchos=   "300-800-1500-1500-1500-1500-1500-3000-800-2500-3000-0-0"
            font            =   "frmRHPeriodosNoLaborales.frx":0644
            font            =   "frmRHPeriodosNoLaborales.frx":066C
            font            =   "frmRHPeriodosNoLaborales.frx":0694
            font            =   "frmRHPeriodosNoLaborales.frx":06BC
            font            =   "frmRHPeriodosNoLaborales.frx":06E4
            fontfixed       =   "frmRHPeriodosNoLaborales.frx":070C
            columnasaeditar =   "X-1-X-3-4-5-6-7-8-X-10-X-X"
            textstylefixed  =   3
            listacontroles  =   "0-1-0-2-2-2-2-0-1-0-0-0-0"
            backcolor       =   -2147483634
            encabezadosalineacion=   "C-L-L-R-R-R-R-L-L-L-L-C-C"
            formatosedit    =   "0-0-0-5-5-5-5-0-0-0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1  'True
            appearance      =   0
            colwidth0       =   300
            rowheight0      =   300
            cellbackcolor   =   -2147483624
         End
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   1
         Left            =   -73710
         TabIndex        =   10
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   2
         Left            =   90
         TabIndex        =   16
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   1290
         TabIndex        =   17
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   1
         Left            =   -74910
         TabIndex        =   11
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   -73710
         TabIndex        =   12
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   3
         Left            =   -74910
         TabIndex        =   21
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   -73710
         TabIndex        =   22
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   0
         Left            =   -73710
         TabIndex        =   5
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   -73710
         TabIndex        =   7
         Top             =   3825
         Width           =   1095
      End
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   60
      Top             =   5610
      _extentx        =   820
      _extenty        =   820
   End
   Begin Sicmact.ctrRRHHGen ctrRRHHGen 
      Height          =   1200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8760
      _extentx        =   15452
      _extenty        =   2117
      font            =   "frmRHPeriodosNoLaborales.frx":0732
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7665
      TabIndex        =   23
      Top             =   5655
      Width           =   1095
   End
End
Attribute VB_Name = "frmRHPeriodosNoLaborales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnEstadosTipo As RHEstadosTpo
Dim lnOpeTpo As TipoOpe
Dim lnIndiceActual As Integer
Dim lbSalir As Boolean
Dim lnColAnt As Long
Dim lnRowAnt As Long
Dim lnColAct As Long
Dim lnRowAct As Long
Dim lnGradoApr As Integer

Private Sub chkSoloPendientes_Click(Index As Integer)
    If Me.chkSoloPendientes(Index).value = 1 Then
        CargaData Me.ctrRRHHGen.psCodigoPersona, True
    Else
        CargaData Me.ctrRRHHGen.psCodigoPersona, False
    End If
End Sub

Private Sub cmdActEstado_Click()
    Dim oPer As DPeriodoNoLaborado
    Set oPer = New DPeriodoNoLaborado
    
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    'Me.lblRHEstado.Caption = oRH.GetRRHHEstado(Me.ctrRRHHGen.psCodigoPersona)
    
    If MsgBox("Desea Actualizar el estado del empleado ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    If Not oPer.ActualizaEstado(Me.ctrRRHHGen.psCodigoPersona, gdFecSis, GetMovNro(gsCodUser, gsCodAge)) Then
        MsgBox "No se ha modificado el estado, es probable que no exista un periodo no laborado, que lo modifique."
    Else
        MsgBox "Se ha modificado el estado con exito a " & oRH.GetRRHHEstado(Me.ctrRRHHGen.psCodigoPersona), vbInformation
    End If
    
    Set oPer = Nothing
End Sub

Private Sub CmdCancelar_Click(Index As Integer)
    CargaData Me.ctrRRHHGen.psCodigoPersona
    Activa False
End Sub

Private Sub cmdEditar_Click(Index As Integer)
    If Me.ctrRRHHGen.psCodigoPersona = "" Then Exit Sub
    Activa True
End Sub

Private Sub CmdEliminar_Click(Index As Integer)
    If FlexPerNoLab(Index).TextMatrix(FlexPerNoLab(Index).Row, FlexPerNoLab(Index).Cols - 1) = "1" Then
            MsgBox "No se puede Eliminar", vbInformation, "Aviso"
    End If
    
    Me.FlexPerNoLab(Index).EliminaFila Me.FlexPerNoLab(Index).Row
End Sub

Private Sub cmdGrabar_Click(Index As Integer)
    If MsgBox("Desea Grabar ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    Dim oPer As NPeriodoNoLaborado
    Set oPer = New NPeriodoNoLaborado
    
    If lnEstadosTipo = RHEstadosTpoVacaciones Then
        lnIndiceActual = 0
    ElseIf lnEstadosTipo = RHEstadosTpoPermisosLicencias Then
        lnIndiceActual = 1
    ElseIf lnEstadosTipo = RHEstadosTpoSubsidiado Then
        lnIndiceActual = 2
    ElseIf lnEstadosTipo = RHEstadosTpoSuspendido Then
        lnIndiceActual = 3
    End If
    
    If Not Valida Then
        Me.FlexPerNoLab(lnIndiceActual).SetFocus
        Exit Sub
    End If
    
    oPer.ModificaPerNoLab Me.ctrRRHHGen.psCodigoPersona, Trim(Str(lnEstadosTipo)), Me.FlexPerNoLab(Index).GetRsNew, GetMovNro(gsCodUser, gsCodAge), gsFormatoFechaHora
    
    Set oPer = Nothing
    CmdCancelar_Click Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    Dim oPer As NPeriodoNoLaborado
    Set oPer = New NPeriodoNoLaborado
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    Dim lsCadena As String
    
    lsCadena = oPer.GetReporte(Me.ctrRRHHGen.psCodigoPersona, Me.ctrRRHHGen.psNombreEmpledo, CInt(lnEstadosTipo), "Hola", gsNomAge, gsEmpresa, gdFecSis)
    
    oPrevio.Show lsCadena, Caption, True
End Sub

Private Sub CmdNuevo_Click(Index As Integer)
    Me.FlexPerNoLab(Index).AdicionaFila
    
    If lnEstadosTipo = RHEstadosTpoVacaciones Then
        'Me.FlexPerNoLab(Index).TextMatrix(Me.FlexPerNoLab(Index).Row, 1) = "301"
        Me.FlexPerNoLab(Index).TextMatrix(Me.FlexPerNoLab(Index).Row, 8) = RHPerNoLab.RHPerNoLabAprovado
        Me.FlexPerNoLab(Index).Col = 1
    ElseIf lnEstadosTipo = RHEstadosTpoPermisosLicencias Then
        If lnOpeTpo = gTipoOpeRegistro Then
            Me.FlexPerNoLab(Index).TextMatrix(Me.FlexPerNoLab(Index).Row, 8) = "0"
            Me.FlexPerNoLab(Index).Col = 1
        End If
    ElseIf lnEstadosTipo = RHEstadosTpoSuspendido Or lnEstadosTipo = RHEstadosTpoSubsidiado Then
        Me.FlexPerNoLab(Index).TextMatrix(Me.FlexPerNoLab(Index).Row, 8) = RHPerNoLab.RHPerNoLabAprovado
    End If
    'FlexPerNoLab(Index).BackColorSel = &H8000000D
    FlexPerNoLab(Index).SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Public Sub Ini(pnOpeTpo As TipoOpe, pnEstadosTipo As RHEstadosTpo, psCaption As String)
    lnOpeTpo = pnOpeTpo
    lnEstadosTipo = pnEstadosTipo
    If lnEstadosTipo = RHEstadosTpoVacaciones Then
        lnIndiceActual = 0
    ElseIf lnEstadosTipo = RHEstadosTpoPermisosLicencias Then
        lnIndiceActual = 1
    ElseIf lnEstadosTipo = RHEstadosTpoSubsidiado Then
        lnIndiceActual = 2
    ElseIf lnEstadosTipo = RHEstadosTpoSuspendido Then
        lnIndiceActual = 3
    End If
    
    IniTab
    Caption = psCaption
    Me.Show 1
End Sub

Private Sub ctrRRHHGen_EmiteDatos()
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    Set oRRHH = New DActualizaDatosRRHH
    Set oPersona = New UPersona
    Set oPersona = frmBuscaPersona.Inicio(True)
    Dim lnEmpGA As Integer
    Dim lsEmpAge As String
    Dim lsEmpArea As String
    
    'oPersona.ObtieneClientexCodigo ()
    ClearScreen
    If Not oPersona Is Nothing Then
        If lnEstadosTipo = RHEstadosTpoVacaciones Then
            If Not oRRHH.GetRRHHControl(oPersona.sPersCod, TipoControl.TipoControlVacaciones) Then
                MsgBox "La Persona Pertenece a un grupo de empleados al que no se le controla las vacaciones.", vbInformation, "Aviso"
                Set oRRHH = Nothing
                Set oPersona = Nothing
                Exit Sub
            End If
        ElseIf lnEstadosTipo = RHEstadosTpoPermisosLicencias Then
            If Not oRRHH.GetRRHHControl(oPersona.sPersCod, TipoControl.TipoControlPermisos) Then
                MsgBox "La Persona Pertenece a un grupo de empleados al que no se le controla los permisos.", vbInformation, "Aviso"
                Set oRRHH = Nothing
                Set oPersona = Nothing
                Exit Sub
            End If
            
            lnEmpGA = oRRHH.GetNiveldeAprovacion(oPersona.sPersCod)
            oRRHH.GetAreaAgenciaRRHH oPersona.sPersCod, lsEmpArea, lsEmpAge
            
            'If lnEmpGA >= lnGradoApr And gnGradoMaxAut <> lnEmpGA Then
            '    Me.cmdEditar(lnIndiceActual).Enabled = False
            'ElseIf lnEmpGA + 1 = lnGradoApr Then
            '   If lsEmpArea <> Usuario.AreaCod Then
            '        Me.cmdEditar(lnIndiceActual).Enabled = False
            '    Else
            '        Me.cmdEditar(lnIndiceActual).Enabled = True
            '    End If
            'ElseIf lnEmpGA + 2 = lnGradoApr Then
            '    If lsEmpAge <> Usuario.CodAgeAsig Then
            '        Me.cmdEditar(lnIndiceActual).Enabled = False
            '    Else
            '        Me.cmdEditar(lnIndiceActual).Enabled = True
            '    End If
            'Else
            '    Me.cmdEditar(lnIndiceActual).Enabled = True
            'End If
        End If
        
        ClearScreen
        Me.ctrRRHHGen.psCodigoPersona = oPersona.sPersCod
        Me.ctrRRHHGen.psNombreEmpledo = oPersona.sPersNombre
        Me.ctrRRHHGen.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(Me.ctrRRHHGen.psCodigoPersona)
        CargaData Me.ctrRRHHGen.psCodigoPersona
    
    End If
End Sub

Private Sub CargaData(psPersCod As String, Optional pbSoloPendientes As Boolean = False)
    Dim oPNL As DPeriodoNoLaborado
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Set oPNL = New DPeriodoNoLaborado
    
    Set rsP = oPNL.GetRHPeriodoNoLabPersona(psPersCod, CInt(lnEstadosTipo), pbSoloPendientes)
    
    If Not (rsP.EOF And rsP.BOF) Then
        Set Me.FlexPerNoLab(lnIndiceActual).Recordset = rsP
    Else
        FlexPerNoLab(lnIndiceActual).Clear
        FlexPerNoLab(lnIndiceActual).Rows = 2
        FlexPerNoLab(lnIndiceActual).FormaCabecera
    End If
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.cmdSalir.Enabled = Not pbValor
    
    If lnEstadosTipo = RHEstadosTpoVacaciones Then
        If lnOpeTpo = gTipoOpeMantenimiento Then
            Me.fraPerNoLab(lnIndiceActual).Enabled = pbValor
            Me.cmdEditar(lnIndiceActual).Visible = Not pbValor
            Me.cmdGrabar(lnIndiceActual).Enabled = pbValor
            Me.cmdCancelar(lnIndiceActual).Visible = pbValor
            Me.ctrRRHHGen.Enabled = Not pbValor
        ElseIf lnOpeTpo = gTipoOpeConsulta Then
            Me.cmdEditar(lnIndiceActual).Visible = False
            Me.cmdGrabar(lnIndiceActual).Visible = False
            Me.cmdCancelar(lnIndiceActual).Visible = False
            Me.cmdNuevo(lnIndiceActual).Visible = False
            Me.cmdEliminar(lnIndiceActual).Visible = False
        End If
    ElseIf lnEstadosTipo = RHEstadosTpoPermisosLicencias Then
        If lnOpeTpo = gTipoOpeRegistro Then
            Me.fraPerNoLab(lnIndiceActual).Enabled = pbValor
            Me.cmdEditar(lnIndiceActual).Visible = Not pbValor
            Me.cmdGrabar(lnIndiceActual).Enabled = pbValor
            Me.cmdCancelar(lnIndiceActual).Visible = pbValor
            Me.ctrRRHHGen.Enabled = Not pbValor
        ElseIf lnOpeTpo = gTipoOpeMantenimiento Then
            Me.fraPerNoLab(lnIndiceActual).Enabled = pbValor
            Me.cmdEditar(lnIndiceActual).Visible = Not pbValor
            Me.cmdGrabar(lnIndiceActual).Enabled = pbValor
            Me.cmdCancelar(lnIndiceActual).Visible = pbValor
            Me.ctrRRHHGen.Enabled = Not pbValor
            Me.cmdNuevo(lnIndiceActual).Visible = False
            Me.cmdEliminar(lnIndiceActual).Visible = False
        ElseIf lnOpeTpo = gTipoOpeConsulta Then
            Me.cmdEditar(lnIndiceActual).Visible = False
            Me.cmdGrabar(lnIndiceActual).Visible = False
            Me.cmdCancelar(lnIndiceActual).Visible = False
            Me.cmdNuevo(lnIndiceActual).Visible = False
            Me.cmdEliminar(lnIndiceActual).Visible = False
        End If
    ElseIf lnEstadosTipo = RHEstadosTpoSuspendido Then
        If lnOpeTpo = gTipoOpeRegistro Then
            Me.fraPerNoLab(lnIndiceActual).Enabled = pbValor
            Me.cmdEditar(lnIndiceActual).Visible = Not pbValor
            Me.cmdGrabar(lnIndiceActual).Enabled = pbValor
            Me.cmdCancelar(lnIndiceActual).Visible = pbValor
            
        ElseIf lnOpeTpo = gTipoOpeMantenimiento Then
            Me.fraPerNoLab(lnIndiceActual).Enabled = pbValor
            Me.cmdEditar(lnIndiceActual).Visible = Not pbValor
            Me.cmdGrabar(lnIndiceActual).Enabled = pbValor
            Me.cmdCancelar(lnIndiceActual).Visible = pbValor
            Me.ctrRRHHGen.Enabled = Not pbValor
        ElseIf lnOpeTpo = gTipoOpeConsulta Then
            Me.cmdEditar(lnIndiceActual).Visible = False
            Me.cmdGrabar(lnIndiceActual).Visible = False
            Me.cmdCancelar(lnIndiceActual).Visible = False
            Me.cmdNuevo(lnIndiceActual).Visible = False
            Me.cmdEliminar(lnIndiceActual).Visible = False
        End If
    ElseIf lnEstadosTipo = RHEstadosTpoSubsidiado Then
        If lnOpeTpo = gTipoOpeRegistro Then
            Me.fraPerNoLab(lnIndiceActual).Enabled = pbValor
            Me.cmdEditar(lnIndiceActual).Visible = Not pbValor
            Me.cmdGrabar(lnIndiceActual).Enabled = pbValor
            Me.cmdCancelar(lnIndiceActual).Visible = pbValor
            
        ElseIf lnOpeTpo = gTipoOpeMantenimiento Then
            Me.fraPerNoLab(lnIndiceActual).Enabled = pbValor
            Me.cmdEditar(lnIndiceActual).Visible = Not pbValor
            Me.cmdGrabar(lnIndiceActual).Enabled = pbValor
            Me.cmdCancelar(lnIndiceActual).Visible = pbValor
            Me.ctrRRHHGen.Enabled = Not pbValor
        ElseIf lnOpeTpo = gTipoOpeConsulta Then
            Me.cmdEditar(lnIndiceActual).Visible = False
            Me.cmdGrabar(lnIndiceActual).Visible = False
            Me.cmdCancelar(lnIndiceActual).Visible = False
            Me.cmdNuevo(lnIndiceActual).Visible = False
            Me.cmdEliminar(lnIndiceActual).Visible = False
        End If
    End If
End Sub

Private Sub IniTab()
    Dim i As Integer
    
    For i = 0 To Me.Tab.Tabs - 1
        If lnIndiceActual <> i Then Me.Tab.TabVisible(i) = False
    Next i
End Sub

Private Sub ClearScreen()
    Me.ctrRRHHGen.ClearScreen
    FlexPerNoLab(lnIndiceActual).Clear
    FlexPerNoLab(lnIndiceActual).Rows = 2
    FlexPerNoLab(lnIndiceActual).FormaCabecera
End Sub

Private Sub ctrRRHHGen_KeyPress(KeyAscii As Integer)
    Dim lnEmpGA As Integer
    Dim lsEmpAge As String
    Dim lsEmpArea As String
    
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
            ClearScreen
            ctrRRHHGen.SetFocus
            Exit Sub
        End If
        
        rsR.Close
        Set rsR = Nothing
        
        If lnEstadosTipo = RHEstadosTpoPermisosLicencias Then
            lnEmpGA = oRRHH.GetNiveldeAprovacion(ctrRRHHGen.psCodigoPersona)
            oRRHH.GetAreaAgenciaRRHH ctrRRHHGen.psCodigoPersona, lsEmpArea, lsEmpAge
            
            If lnEmpGA >= lnGradoApr And gnGradoMaxAut <> lnEmpGA Then
                Me.cmdEditar(lnIndiceActual).Enabled = False
            ElseIf lnEmpGA + 1 = lnGradoApr Then
                If lsEmpArea <> Usuario.AreaCod Then
                    Me.cmdEditar(lnIndiceActual).Enabled = False
                Else
                    Me.cmdEditar(lnIndiceActual).Enabled = True
                End If
            ElseIf lnEmpGA + 2 = lnGradoApr Then
                If lsEmpAge <> Usuario.CodAgeAsig Then
                    Me.cmdEditar(lnIndiceActual).Enabled = False
                Else
                    Me.cmdEditar(lnIndiceActual).Enabled = True
                End If
            Else
                Me.cmdEditar(lnIndiceActual).Enabled = True
            End If
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub FlexPerNoLab_OnEnterTextBuscar(Index As Integer, psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    If lnEstadosTipo = RHEstadosTpoPermisosLicencias Then
        If Left(psDataCod, 1) = 1 And pnCol = 1 Then
            Me.FlexPerNoLab(lnIndiceActual).FormatosEdit = "0-0-0-5-5-5-5-0-0-0-0"
        Else
            Me.FlexPerNoLab(lnIndiceActual).FormatosEdit = "0-0-0-0-0-0-0-0-0-0-0"
        End If
    End If
End Sub

Private Sub FlexPerNoLab_OnValidate(Index As Integer, ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If pnCol = 3 Then
       FlexPerNoLab(Index).TextMatrix(pnRow, 5) = FlexPerNoLab(Index).TextMatrix(pnRow, 3)
    ElseIf pnCol = 4 Then
       FlexPerNoLab(Index).TextMatrix(pnRow, 6) = FlexPerNoLab(Index).TextMatrix(pnRow, 4)
    End If

End Sub

Private Sub FlexPerNoLab_RowColChange(Index As Integer)
    Dim oCon As DConstantes
    Dim rsC As ADODB.Recordset
    Set oCon = New DConstantes
    Set rsC = New ADODB.Recordset
    
'    lnColAct = FlexPerNoLab(Index).Col
'    lnRowAct = FlexPerNoLab(Index).Row
'    FlexPerNoLab(Index).CellBackColor = &H80000018
'
'    FlexPerNoLab(Index).Col = lnColAnt
'    FlexPerNoLab(Index).Row = lnRowAnt
'    FlexPerNoLab(Index).CellBackColor = &H80000005
'
'    FlexPerNoLab(Index).Col = lnColAnt
'    FlexPerNoLab(Index).Row = lnRowAnt
'
'    lnColAnt = lnColAct
'    lnRowAnt = lnRowAct
    
    If Not IsNumeric(FlexPerNoLab(Index).TextMatrix(FlexPerNoLab(Index).Row, 1)) Then
        FlexPerNoLab(Index).TextMatrix(FlexPerNoLab(Index).Row, 1) = ""
    End If
    
    If lnEstadosTipo = RHEstadosTpoPermisosLicencias Then
        If FlexPerNoLab(Index).Col = 1 Then
            Set rsC = oCon.GetPeriodosNoLabTpo(RHEstadosTpoPermisosLicencias)
            Me.FlexPerNoLab(lnIndiceActual).rsTextBuscar = rsC
        ElseIf FlexPerNoLab(Index).Col = 8 Then
            If lnOpeTpo = gTipoOpeRegistro Then
                Set rsC = oCon.GetConstante(gRHPeriodoNoLab, , , True)
            Else
                Set rsC = oCon.GetConstante(gRHPeriodoNoLab, , , True, , "0")
            End If
            
            Me.FlexPerNoLab(lnIndiceActual).rsTextBuscar = rsC
            Me.FlexPerNoLab(Index).ListaControles = "0-1-0-2-2-2-2-0-1-0-0-0-0"
        Else
            If Left(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(Me.FlexPerNoLab(lnIndiceActual).Row, 2), 1) = 1 Then
                Me.FlexPerNoLab(lnIndiceActual).FormatosEdit = "0-0-0-5-5-5-5-0-0"
            Else
                Me.FlexPerNoLab(lnIndiceActual).FormatosEdit = "0-0-0-0-0-0-0-0-0"
            End If
        End If
        
        If lnOpeTpo = gTipoOpeRegistro Then
            If Me.FlexPerNoLab(lnIndiceActual).TextMatrix(FlexPerNoLab(lnIndiceActual).Row, 12) = "" Then
                Me.FlexPerNoLab(lnIndiceActual).lbEditarFlex = True
            Else
                Me.FlexPerNoLab(lnIndiceActual).lbEditarFlex = False
            End If
        Else
            Me.FlexPerNoLab(lnIndiceActual).lbEditarFlex = True
        End If
        
    ElseIf lnEstadosTipo = RHEstadosTpoSubsidiado Then
        If FlexPerNoLab(Index).Col = 1 Then
            Set rsC = oCon.GetPeriodosNoLabTpo(RHEstadosTpoSubsidiado)
            Me.FlexPerNoLab(lnIndiceActual).rsTextBuscar = rsC
        End If
        Me.FlexPerNoLab(lnIndiceActual).FormatosEdit = "0-0-0-0-0-0-0-0-0"
    ElseIf lnEstadosTipo = RHEstadosTpoVacaciones Then
        If FlexPerNoLab(Index).Col = 1 Then
            Set rsC = oCon.GetPeriodosNoLabTpo(RHEstadosTpoVacaciones)
            Me.FlexPerNoLab(lnIndiceActual).rsTextBuscar = rsC
        ElseIf FlexPerNoLab(Index).Col = 8 Then
            Set rsC = oCon.GetConstante(gRHPeriodoNoLab, , , True, "1")
            Me.FlexPerNoLab(lnIndiceActual).rsTextBuscar = rsC
        End If
        Me.FlexPerNoLab(lnIndiceActual).FormatosEdit = "0-0-0-0-0-0-0-0-0"
    End If
    
    
End Sub

Private Sub Form_Activate()
    If lbSalir Then Unload Me
End Sub

Private Sub Form_Load()
    Activa False
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim oRRHH As DActualizaDatosRRHH
    Set oRRHH = New DActualizaDatosRRHH
    
    lbSalir = False
    
    If lnEstadosTipo = RHEstadosTpoPermisosLicencias Then
        Set rsC = oCon.GetPeriodosNoLabTpo(RHEstadosTpoPermisosLicencias)
        Me.FlexPerNoLab(lnIndiceActual).rsTextBuscar = rsC
    
        Me.Usuario.Inicio gsCodUser
        lnGradoApr = oRRHH.GetNiveldeAprovacion(Usuario.PersCod)
        
        If Not oRRHH.GetRRHHControl(Usuario.PersCod, TipoControl.TipoControlPermisos) Then
            MsgBox "Sr(ta). " & Me.Usuario.UserNom & Chr(13) & " Ud. Pertenece a un grupo de Recursos Humanos al que no se le controla los permisos.", vbInformation, "Aviso"
            Set oRRHH = Nothing
            lbSalir = True
        Else
            If lnOpeTpo = gTipoOpeRegistro Then
                'Me.ctrRRHHGen.Enabled = False
                Me.ctrRRHHGen.psCodigoPersona = Usuario.PersCod
                Me.ctrRRHHGen.psNombreEmpledo = Usuario.UserNom
                Me.ctrRRHHGen.psCodigoEmpleado = GetCodigoEmpleado(Usuario.PersCod)
                CargaData Me.ctrRRHHGen.psCodigoPersona
                Me.FlexPerNoLab(lnIndiceActual).ColumnasAEditar = "X-1-X-3-4-5-6-7-X-X-X-X-X"
                Me.cmdEliminar(lnIndiceActual).Visible = False
            End If
        End If
            
        If lnOpeTpo = gTipoOpeMantenimiento And lnGradoApr <> 9 Then
            Me.FlexPerNoLab(lnIndiceActual).ColumnasAEditar = "X-X-X-X-X-X-X-X-8-X-10-X-X"
        End If
    ElseIf lnEstadosTipo = RHEstadosTpoSuspendido Then
        Set rsC = oCon.GetPeriodosNoLabTpo(RHEstadosTpoSuspendido)
        Me.FlexPerNoLab(lnIndiceActual).rsTextBuscar = rsC
        Me.FlexPerNoLab(lnIndiceActual).FormatosEdit = "0-0-0-0-0-0-0-0-0"
    ElseIf lnEstadosTipo = RHEstadosTpoSubsidiado Then
        Set rsC = oCon.GetPeriodosNoLabTpo(RHEstadosTpoSubsidiado)
        Me.FlexPerNoLab(lnIndiceActual).rsTextBuscar = rsC
        Me.FlexPerNoLab(lnIndiceActual).FormatosEdit = "0-0-0-0-0-0-0-0-0"
    End If
End Sub

Private Function Valida() As Boolean
    Dim i As Integer
    
    For i = 1 To Me.FlexPerNoLab(lnIndiceActual).Rows - 1
        FlexPerNoLab(lnIndiceActual).Row = i
        If lnEstadosTipo = RHEstadosTpoVacaciones Then
            If Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 1) = "" Then
                MsgBox "Debe Ingresar un tipo valido para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 1
                Valida = False
                Exit Function
            ElseIf Not IsDate(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 3)) And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar una Fecha de Inicio de Programacion Valida para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 3
                Valida = False
                Exit Function
            ElseIf Not IsDate(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 4)) And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar una Fecha de Fin de Programacion Valida para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 4
                Valida = False
                Exit Function
            ElseIf Not IsDate(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 5)) And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar una Fecha de Inicio de Ejecucion Valida para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 5
                Valida = False
                Exit Function
            ElseIf Not IsDate(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 6)) And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar una Fecha de Fin de Ejecucion Valida para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 6
                Valida = False
                Exit Function
            ElseIf Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 7) = "" And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar un comentario para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 7
                Valida = False
                Exit Function
            Else
                Valida = True
            End If
        ElseIf lnEstadosTipo = RHEstadosTpoPermisosLicencias Or lnEstadosTipo = RHEstadosTpoSubsidiado Then
            If Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 1) = "" And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar un tipo de descanso Valido para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 1
                Valida = False
                Exit Function
            ElseIf Not IsDate(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 3)) And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar una Fecha de Inicio de Solcicitud Valida para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 3
                Valida = False
                Exit Function
            ElseIf Not IsDate(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 4)) And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar una Fecha de Fin de Solcicitud Valida para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 4
                Valida = False
                Exit Function
            ElseIf Not IsDate(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 5)) And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar una Fecha de Inicio de Ejecucion Valida para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 5
                Valida = False
                Exit Function
            ElseIf Not IsDate(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 6)) And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar una Fecha de Fin de Ejecucion Valida para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 6
                Valida = False
                Exit Function
            ElseIf Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 7) = "" And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar un comentario para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 7
                Valida = False
                Exit Function
            Else
                Valida = True
            End If
        ElseIf lnEstadosTipo = RHEstadosTpoSuspendido Then
            If Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 1) = "" Then
                MsgBox "Debe Ingresar un tipo de suspnesión Valido para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 1
                Valida = False
                Exit Function
            ElseIf Not IsDate(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 3)) And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar una Fecha de Inicio de Programacion Valida para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 3
                Valida = False
                Exit Function
            ElseIf Not IsDate(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 4)) And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar una Fecha de Fin de Programacion Valida para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 4
                Valida = False
                Exit Function
            ElseIf Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 7) = "" And IsNumeric(Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0)) Then
                MsgBox "Debe Ingresar un comentario para el registro : " & Me.FlexPerNoLab(lnIndiceActual).TextMatrix(i, 0), vbInformation, "Aviso"
                FlexPerNoLab(lnIndiceActual).Col = 7
                Valida = False
                Exit Function
            Else
                Valida = True
            End If
        End If
    Next i
End Function


