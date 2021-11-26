VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRHProgVacaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XXXXXXXX"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmRHProgVacaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   45
      TabIndex        =   2
      Top             =   5700
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1245
      TabIndex        =   3
      Top             =   5700
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2460
      TabIndex        =   6
      Top             =   5700
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3645
      TabIndex        =   7
      Top             =   5700
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7020
      TabIndex        =   17
      Top             =   5700
      Width           =   1095
   End
   Begin SicmactAdmin.ctrRRHH RRHH 
      Height          =   1905
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3360
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
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   45
      TabIndex        =   4
      Top             =   5700
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1245
      TabIndex        =   5
      Top             =   5700
      Width           =   1095
   End
   Begin VB.Frame fraAutorizacionDet 
      Caption         =   "Detalle de Autorizacion :"
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
      Height          =   3480
      Left            =   45
      TabIndex        =   21
      Top             =   2190
      Width           =   8040
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         Height          =   510
         Left            =   1335
         TabIndex        =   9
         Top             =   615
         Width           =   6570
      End
      Begin VB.ComboBox cmbEstado 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3060
         Width           =   6495
      End
      Begin VB.Frame fraEjecutado 
         Caption         =   "Ejecutado"
         Height          =   615
         Left            =   150
         TabIndex        =   29
         Top             =   2385
         Width           =   7740
         Begin MSMask.MaskEdBox mskEjecutadoIni 
            Height          =   330
            Left            =   1230
            TabIndex        =   14
            Top             =   195
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   19
            Mask            =   "##/##/#### ##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskEjecutadoFin 
            Height          =   330
            Left            =   5280
            TabIndex        =   15
            Top             =   195
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   19
            Mask            =   "##/##/#### ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblEjecutadoIni 
            Caption         =   "Inicio :"
            Height          =   225
            Left            =   225
            TabIndex        =   31
            Top             =   240
            Width           =   960
         End
         Begin VB.Label lblEjecutadoFin 
            Caption         =   "Fin :"
            Height          =   195
            Left            =   4125
            TabIndex        =   30
            Top             =   285
            Width           =   1050
         End
      End
      Begin VB.Frame FraProgramado 
         Caption         =   "Programado"
         Height          =   615
         Left            =   150
         TabIndex        =   23
         Top             =   1740
         Width           =   7740
         Begin MSMask.MaskEdBox mskProgramadoIni 
            Height          =   330
            Left            =   1230
            TabIndex        =   12
            Top             =   180
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   19
            Mask            =   "##/##/#### ##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskProgramadoFin 
            Height          =   330
            Left            =   5280
            TabIndex        =   13
            Top             =   195
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   19
            Mask            =   "##/##/#### ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblProgramadoFin 
            Caption         =   "Fin :"
            Height          =   195
            Left            =   4125
            TabIndex        =   25
            Top             =   285
            Width           =   1050
         End
         Begin VB.Label lblProgramadoIni 
            Caption         =   "Inicio :"
            Height          =   225
            Left            =   240
            TabIndex        =   24
            Top             =   255
            Width           =   960
         End
      End
      Begin VB.Frame fraSolicitud 
         Caption         =   "Solicitud"
         Height          =   615
         Left            =   150
         TabIndex        =   26
         Top             =   1095
         Width           =   7740
         Begin MSMask.MaskEdBox mskSolicitudIni 
            Height          =   330
            Left            =   1230
            TabIndex        =   10
            Top             =   195
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   19
            Mask            =   "##/##/#### ##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskSolicitudFin 
            Height          =   330
            Left            =   5280
            TabIndex        =   11
            Top             =   195
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   19
            Mask            =   "##/##/#### ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblSolicitudIni 
            Caption         =   "Inicio :"
            Height          =   225
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   960
         End
         Begin VB.Label lblSolicitadoFin 
            Caption         =   "Fin :"
            Height          =   195
            Left            =   4125
            TabIndex        =   27
            Top             =   285
            Width           =   1050
         End
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   210
         Width           =   6570
      End
      Begin VB.Label lblTipo 
         Caption         =   "Tipo :"
         Height          =   195
         Left            =   195
         TabIndex        =   34
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label lblEstado 
         Caption         =   "Estado :"
         Height          =   195
         Left            =   210
         TabIndex        =   33
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblComentario 
         Caption         =   "Comentario :"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   615
         Width           =   945
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Autorización "
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
      Height          =   3480
      Left            =   45
      TabIndex        =   20
      Top             =   2190
      Width           =   8040
      Begin SicmactAdmin.FlexEdit FlexEdit1 
         Height          =   3135
         Left            =   90
         TabIndex        =   1
         Top             =   240
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   5530
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Nombre-Tipo-FIni_Solicitud-FFin_Solicitud-FIni_Programado-FFin_Programado-FIni_Ejecutado-FFin_Ejecutado-Comentario-Estado"
         EncabezadosAnchos=   "400-4000-4000-2000-2000-2000-2000-2000-2000-4000-2000"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-L-L-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   285
      End
   End
   Begin VB.Label lblDiasSubsidioL 
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
      Height          =   195
      Left            =   4755
      TabIndex        =   35
      Top             =   1950
      Width           =   1545
   End
   Begin VB.Label lblDiasSubsidio 
      Caption         =   "Dias Subsidio"
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
      Height          =   195
      Left            =   3345
      TabIndex        =   22
      Top             =   1950
      Width           =   1245
   End
   Begin VB.Label lblFechaIngresoL 
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
      Height          =   195
      Left            =   1560
      TabIndex        =   19
      Top             =   1950
      Width           =   1425
   End
   Begin VB.Label lblFechaIng 
      Caption         =   "Fecha Ingreso :"
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
      Height          =   195
      Left            =   60
      TabIndex        =   18
      Top             =   1950
      Width           =   1500
   End
End
Attribute VB_Name = "frmRHProgVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbEditado As Boolean
Dim lnTipo As RHAutoFisicaGrupo
Dim lnTipoOpe As TipoOpe
Dim lnTipoEstAut As RHAutoFisicaTipoEstado

Private Sub cmbTipo_Click()
    If cmbTipo.Text <> "" Then
        If Left(Right(cmbTipo.Text, 3), 1) = "0" Then
            Me.mskEjecutadoFin.Mask = "##/##/####"
            Me.mskEjecutadoIni.Mask = "##/##/####"
            Me.mskProgramadoFin.Mask = "##/##/####"
            Me.mskProgramadoIni.Mask = "##/##/####"
            Me.mskSolicitudFin.Mask = "##/##/####"
            Me.mskSolicitudIni.Mask = "##/##/####"
        Else
            Me.mskEjecutadoFin.Mask = "##/##/#### ##:##:##"
            Me.mskEjecutadoIni.Mask = "##/##/#### ##:##:##"
            Me.mskProgramadoFin.Mask = "##/##/#### ##:##:##"
            Me.mskProgramadoIni.Mask = "##/##/#### ##:##:##"
            Me.mskSolicitudFin.Mask = "##/##/#### ##:##:##"
            Me.mskSolicitudIni.Mask = "##/##/#### ##:##:##"
        End If
    End If
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtComentario.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
    lbEditado = False
    Activa False
End Sub

Private Sub cmdEditar_Click()
    If Me.RRHH.psCodigoPersona = "" Or Me.FlexEdit1.TextMatrix(FlexEdit1.Row, 2) = "" Then Exit Sub
    
    lbEditado = True
    Activa True
    
    If Me.fraSolicitud.Enabled Then
        Me.mskSolicitudIni.SetFocus
    ElseIf Me.FraProgramado.Enabled Then
        Me.mskProgramadoIni.SetFocus
    ElseIf Me.fraEjecutado.Enabled Then
        Me.mskEjecutadoIni.SetFocus
    Else
        Me.cmbEstado.SetFocus
    End If
End Sub

Private Sub cmdEliminar_Click()
    Dim oAut As NAutorizacionFisica
    Set oAut = New NAutorizacionFisica
    
    If Me.RRHH.psCodigoPersona = "" Or Me.FlexEdit1.TextMatrix(FlexEdit1.Row, 2) = "" Then Exit Sub
    
    If MsgBox("Desea Eliminar el Registro ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oAut.Elimina Me.RRHH.psCodigoPersona, Right(Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 2), 2), Format(Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 3), gsFormatoFechaHora)
    
    CargaDatos
End Sub

Private Sub CmdGrabar_Click()
    Dim oAut As NAutorizacionFisica
    Set oAut = New NAutorizacionFisica
    
    If Not Valida Then Exit Sub
    
    If lnTipo = RHGrupoAutoFisicaSanciones Then
        If Not lbEditado Then
            If Not oAut.AgredaDatos(Me.RRHH.psCodigoPersona, Right(Me.cmbTipo.Text, 2), IIf(Not IsDate(Me.mskEjecutadoIni.Text), "", Format(Me.mskEjecutadoIni.Text, gsFormatoFechaHora)), IIf(Not IsDate(Me.mskEjecutadoFin.Text), "", Format(Me.mskEjecutadoFin.Text, gsFormatoFechaHora)), IIf(Not IsDate(Me.mskEjecutadoIni.Text), "", Format(Me.mskEjecutadoIni.Text, gsFormatoFechaHora)), IIf(Not IsDate(Me.mskEjecutadoFin.Text), "", Format(Me.mskEjecutadoFin.Text, gsFormatoFechaHora)), Format(Me.mskEjecutadoIni.Text, gsFormatoFechaHora), Format(Me.mskEjecutadoFin.Text, gsFormatoFechaHora), Me.txtComentario.Text, Right(Me.cmbEstado.Text, 1), GetMovNro(gsCodUser, gsCodAge)) Then
                MsgBox "Ud. no puede tener dos Autorizaciones del mismo tipo en el mimo momento.", vbInformation, "Aviso"
                If mskSolicitudIni.Enabled Then
                    Me.mskSolicitudIni.SetFocus
                End If
                Exit Sub
                
            End If
        Else
            If Not oAut.ModificaDatos(Me.RRHH.psCodigoPersona, Right(Me.cmbTipo.Text, 2), IIf(Not IsDate(Me.mskEjecutadoIni.Text), "", Format(Me.mskEjecutadoIni.Text, gsFormatoFechaHora)), IIf(Not IsDate(Me.mskEjecutadoFin.Text), "", Format(Me.mskEjecutadoFin.Text, gsFormatoFechaHora)), IIf(Not IsDate(Me.mskEjecutadoIni.Text), "", Format(Me.mskEjecutadoIni.Text, gsFormatoFechaHora)), IIf(Not IsDate(Me.mskEjecutadoFin.Text), "", Format(Me.mskEjecutadoFin.Text, gsFormatoFechaHora)), Format(Me.mskEjecutadoIni.Text, gsFormatoFechaHora), Format(Me.mskEjecutadoFin.Text, gsFormatoFechaHora), Me.txtComentario.Text, Right(Me.cmbEstado.Text, 1), GetMovNro(gsCodUser, gsCodAge), Right(Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 2), 2), Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 3)) Then
                MsgBox "Ud. no puede tener dos Autorizaciones del mismo tipo en el mimo momento.", vbInformation, "Aviso"
                Me.mskSolicitudIni.SetFocus
                Exit Sub
            End If
        End If
    Else
    
    If Not lbEditado Then
        If Not oAut.AgredaDatos(Me.RRHH.psCodigoPersona, Right(Me.cmbTipo.Text, 2), IIf(Not IsDate(Me.mskSolicitudIni.Text), "", Format(Me.mskSolicitudIni.Text, gsFormatoFechaHora)), IIf(Not IsDate(Me.mskSolicitudFin.Text), "", Format(Me.mskSolicitudFin.Text, gsFormatoFechaHora)), IIf(Not IsDate(Me.mskProgramadoIni.Text), "", Format(Me.mskProgramadoIni.Text, gsFormatoFechaHora)), IIf(Not IsDate(Me.mskProgramadoFin.Text), "", Format(Me.mskProgramadoFin.Text, gsFormatoFechaHora)), Format(Me.mskEjecutadoIni.Text, gsFormatoFechaHora), Format(Me.mskEjecutadoFin.Text, gsFormatoFechaHora), Me.txtComentario.Text, Right(Me.cmbEstado.Text, 1), GetMovNro(gsCodUser, gsCodAge)) Then
            MsgBox "Ud. no puede tener dos Autorizaciones del mismo tipo en el mimo momento.", vbInformation, "Aviso"
            If mskSolicitudIni.Enabled Then
                Me.mskSolicitudIni.SetFocus
                Exit Sub
            End If
            Exit Sub
        End If
    Else
        If Not oAut.ModificaDatos(Me.RRHH.psCodigoPersona, Right(Me.cmbTipo.Text, 2), IIf(Not IsDate(Me.mskSolicitudIni.Text), "", Format(Me.mskSolicitudIni.Text, gsFormatoFechaHora)), IIf(Not IsDate(Me.mskSolicitudFin.Text), "", Format(Me.mskSolicitudFin.Text, gsFormatoFechaHora)), IIf(Not IsDate(Me.mskProgramadoIni.Text), "", Format(Me.mskProgramadoIni.Text, gsFormatoFechaHora)), IIf(Not IsDate(Me.mskProgramadoFin.Text), "", Format(Me.mskProgramadoFin.Text, gsFormatoFechaHora)), Format(Me.mskEjecutadoIni.Text, gsFormatoFechaHora), Format(Me.mskEjecutadoFin.Text, gsFormatoFechaHora), Me.txtComentario.Text, Right(Me.cmbEstado.Text, 1), GetMovNro(gsCodUser, gsCodAge), Right(Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 2), 2), Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 3)) Then
            MsgBox "Ud. no puede tener dos Autorizaciones del mismo tipo en el mimo momento.", vbInformation, "Aviso"
            Me.mskSolicitudIni.SetFocus
            Exit Sub
        End If
    End If
    
    End If
    lbEditado = False
    Activa False
    CargaDatos
End Sub

Private Sub cmdImprimir_Click()
    Dim oAut As NAutorizacionFisica
    Dim oPrevio As Previo.clsPrevio
    Dim lsCadena As String
    Set oPrevio = New Previo.clsPrevio
    Set oAut = New NAutorizacionFisica
    
    If Me.RRHH.psCodigoPersona = "" Then Exit Sub
    
    lsCadena = oAut.GetReporte(Me.RRHH.psCodigoPersona, CInt(lnTipo), Me.Caption, gsNomAge, gsEmpresa, gdFecSis)
    
    If lsCadena <> "" Then
        oPrevio.Show lsCadena, Caption, True, 66
    End If
    
    Set oPrevio = Nothing
    Set oAut = Nothing
End Sub

Private Sub cmdNuevo_Click()
    If Me.RRHH.psCodigoPersona = "" Then Exit Sub
    Limpia
    If lnTipo = RHGrupoAutoFisicaSanciones Then
        UbicaCombo Me.cmbEstado, "1"
    Else
        UbicaCombo Me.cmbEstado, "0"
    End If
    lbEditado = False
    Activa True
    Me.cmbTipo.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.cmdNuevo.Visible = Not pbValor
    Me.cmdEditar.Visible = Not pbValor
    Me.cmdGrabar.Visible = pbValor
    Me.cmdCancelar.Visible = pbValor
    Me.cmdSalir.Enabled = Not pbValor
    Me.fraDatos.Visible = Not pbValor
    Me.RRHH.Enabled = Not pbValor
    Me.fraAutorizacionDet.Visible = pbValor
    
    If lnTipoOpe = gTipoOpeRegistro Then
        Me.cmdEditar.Enabled = pbValor
        Me.cmdEliminar.Enabled = False
        Me.cmdImprimir.Enabled = False
    ElseIf lnTipoOpe = gTipoOpeMantenimiento Then
        Me.cmdEliminar.Enabled = Not pbValor
        Me.cmdImprimir.Enabled = False
        Me.cmdNuevo.Enabled = False
    ElseIf lnTipoOpe = gTipoOpeConsulta Then
        Me.cmdEliminar.Enabled = False
        Me.cmdNuevo.Enabled = False
        Me.cmdEditar.Enabled = False
        Me.cmdImprimir.Enabled = False
    ElseIf lnTipoOpe = gTipoOpeReporte Then
        Me.cmdEliminar.Enabled = pbValor
        Me.cmdNuevo.Enabled = pbValor
        Me.cmdEditar.Enabled = pbValor
    End If
    
End Sub

Private Sub FlexEdit1_DblClick()
    'cmdEditar_Click
End Sub

Private Sub FlexEdit1_OnRowChange(pnRow As Long, pnCol As Long)
    If Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 3) = "" Then Exit Sub
    UbicaCombo Me.cmbTipo, Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 2), , 2
    Me.txtComentario.Text = Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 9)
    If Left(Right(cmbTipo.Text, 3), 1) = "0" Then
        Me.mskSolicitudIni.Text = Left(Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 3), 10)
        Me.mskSolicitudFin.Text = Left(Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 4), 10)
        Me.mskProgramadoIni.Text = Left(Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 5), 10)
        Me.mskProgramadoFin.Text = Left(Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 6), 10)
        Me.mskEjecutadoIni.Text = Left(Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 7), 10)
        Me.mskEjecutadoFin.Text = Left(Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 8), 10)
    Else
        Me.mskSolicitudIni.Text = Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 3)
        Me.mskSolicitudFin.Text = Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 4)
        Me.mskProgramadoIni.Text = Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 5)
        Me.mskProgramadoFin.Text = Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 6)
        Me.mskEjecutadoIni.Text = Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 7)
        Me.mskEjecutadoFin.Text = Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 8)
    End If
    UbicaCombo Me.cmbEstado, Me.FlexEdit1.TextMatrix(Me.FlexEdit1.Row, 10), , 1
End Sub

Private Sub Form_Load()
    Dim oCon As DConstantes
    Dim oAut As DAutorizacionFisica
    Dim rsC As ADODB.Recordset
    Set oCon = New DConstantes
    Set oAut = New DAutorizacionFisica
    Set rsC = New ADODB.Recordset
    
    Set rsC = oCon.GetConstante(6025)
    CargaCombo rsC, Me.cmbEstado, 150
    
    rsC.Close
    Set rsC = oAut.GetAutorizacionFisicaTipo(CInt(lnTipo))
    CargaCombo rsC, Me.cmbTipo, 150
    
    If lnTipo = RHGrupoAutoFisicaVacaciones Then
        Me.lblFechaIngresoL.Visible = False
        Me.lblFechaIng.Visible = False
        Me.lblDiasSubsidio.Visible = False
        Me.lblDiasSubsidioL.Visible = False
    Else
        Me.lblFechaIngresoL.Visible = True
        Me.lblFechaIng.Visible = True
        Me.lblDiasSubsidio.Visible = True
        Me.lblDiasSubsidioL.Visible = True
    End If
    
    If lnTipoEstAut = RHGrupoAutoFisicaSolicitada Then
        Me.fraSolicitud.Enabled = True
        Me.FraProgramado.Enabled = False
        Me.fraEjecutado.Enabled = False
        Me.cmbEstado.Enabled = False
        
    ElseIf lnTipoEstAut = RHGrupoAutoFisicaProgramada Then
        Me.fraSolicitud.Enabled = False
        Me.FraProgramado.Enabled = True
        Me.fraEjecutado.Enabled = False
        Me.cmbEstado.Enabled = False
    ElseIf lnTipoEstAut = RHGrupoAutoFisicaEjecutada Then
        Me.fraSolicitud.Enabled = False
        Me.FraProgramado.Enabled = False
        Me.fraEjecutado.Enabled = True
        Me.cmbEstado.Enabled = False
    Else
        Me.fraSolicitud.Enabled = False
        Me.FraProgramado.Enabled = False
        Me.fraEjecutado.Enabled = False
        Me.cmbEstado.Enabled = True
    End If
    Activa False
End Sub

Private Function Valida() As Boolean
    If Me.cmbTipo.Text = "" Then
        MsgBox "Debe Ingresar un Tipo.", vbInformation, "Aviso"
        Valida = False
        Me.txtComentario.SetFocus
    ElseIf lnTipoEstAut = RHGrupoAutoFisicaSolicitada Then
        If Not IsDate(Me.mskSolicitudIni.Text) Then
            MsgBox "Debe Ingresar una Fecha valida.", vbInformation, "Aviso"
            Valida = False
            mskSolicitudIni.SetFocus
        ElseIf Not IsDate(Me.mskSolicitudFin.Text) Then
            MsgBox "Debe Ingresar una Fecha valida.", vbInformation, "Aviso"
            Valida = False
            mskSolicitudFin.SetFocus
        Else
            Valida = True
        End If
    ElseIf lnTipoEstAut = RHGrupoAutoFisicaProgramada Then
        If Not IsDate(Me.mskProgramadoIni.Text) Then
            MsgBox "Debe Ingresar una Fecha valida.", vbInformation, "Aviso"
            Valida = False
            mskProgramadoIni.SetFocus
        ElseIf Not IsDate(Me.mskProgramadoFin.Text) Then
            MsgBox "Debe Ingresar una Fecha valida.", vbInformation, "Aviso"
            Valida = False
            mskProgramadoFin.SetFocus
        Else
            Valida = True
        End If
    ElseIf lnTipoEstAut = RHGrupoAutoFisicaEjecutada Then
        If Not IsDate(Me.mskEjecutadoIni.Text) Then
            MsgBox "Debe Ingresar una Fecha valida.", vbInformation, "Aviso"
            Valida = False
            mskEjecutadoIni.SetFocus
        ElseIf Not IsDate(Me.mskEjecutadoFin.Text) Then
            MsgBox "Debe Ingresar una Fecha valida.", vbInformation, "Aviso"
            Valida = False
            mskEjecutadoFin.SetFocus
        Else
            Valida = True
        End If
    Else
        Valida = True
    End If
    
End Function

Private Sub Limpia()
    Me.cmbTipo.ListIndex = -1
    Me.txtComentario.Text = ""
    Me.lblDiasSubsidioL.Caption = "0"
    If mskSolicitudIni.Mask = "##/##/####" Then
        Me.mskSolicitudIni.Text = "__/__/____"
        Me.mskSolicitudFin.Text = "__/__/____"
        Me.mskProgramadoIni.Text = "__/__/____"
        Me.mskProgramadoFin.Text = "__/__/____"
        Me.mskEjecutadoIni.Text = "__/__/____"
        Me.mskEjecutadoFin.Text = "__/__/____"
    Else
        Me.mskSolicitudIni.Text = "__/__/____ __:__:__"
        Me.mskSolicitudFin.Text = "__/__/____ __:__:__"
        Me.mskProgramadoIni.Text = "__/__/____ __:__:__"
        Me.mskProgramadoFin.Text = "__/__/____ __:__:__"
        Me.mskEjecutadoIni.Text = "__/__/____ __:__:__"
        Me.mskEjecutadoFin.Text = "__/__/____ __:__:__"
    End If
End Sub

Private Sub CargaDatos()
    Dim oAut As DAutorizacionFisica
    Set oAut = New DAutorizacionFisica
    Dim rsA As ADODB.Recordset
    Set rsA = New ADODB.Recordset
    
    Set rsA = oAut.GetAutorizacionFisica(Me.RRHH.psCodigoPersona, CInt(lnTipo))
    If rsA Is Nothing Then
        Me.FlexEdit1.FormaCabecera
        Exit Sub
    End If
    
    If rsA.EOF And rsA.BOF Then
        Me.FlexEdit1.Clear
        FlexEdit1.Rows = 2
        Me.FlexEdit1.FormaCabecera
        Exit Sub
    End If
    
    Set Me.FlexEdit1.Recordset = rsA
    Set rsA = Nothing
    Me.lblDiasSubsidioL = oAut.GetNroDias(Me.RRHH.psCodigoPersona, "01/01/" & Format(gdFecSis, "yyyy"), Format(gdFecSis, gcFormatoFecha), "99", , , False)
    
    FlexEdit1_OnRowChange FlexEdit1.Row, 2
End Sub

Private Sub mskEjecutadoFin_GotFocus()
    mskEjecutadoFin.SelStart = 0
    mskEjecutadoFin.SelLength = 50
End Sub

Private Sub mskEjecutadoFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.cmbEstado.Enabled Then
            Me.cmbEstado.SetFocus
        Else
            Me.cmdGrabar.SetFocus
        End If
    End If
End Sub

Private Sub mskEjecutadoIni_GotFocus()
    Me.mskEjecutadoIni.SelStart = 0
    Me.mskEjecutadoIni.SelLength = 50
End Sub

Private Sub mskEjecutadoIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskEjecutadoFin.SetFocus
    End If
End Sub

Private Sub mskProgramadoFin_GotFocus()
    Me.mskProgramadoFin.SelStart = 0
    Me.mskProgramadoFin.SelLength = 50
End Sub

Private Sub mskProgramadoFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.fraEjecutado.Enabled Then
            Me.mskEjecutadoIni.SetFocus
        ElseIf Me.cmbEstado.Enabled Then
            cmbEstado.SetFocus
        Else
            cmdGrabar.SetFocus
        End If
    End If
End Sub

Private Sub mskProgramadoFin_LostFocus()
    Me.mskEjecutadoFin.Text = Me.mskProgramadoFin.Text
End Sub

Private Sub mskProgramadoIni_GotFocus()
    Me.mskProgramadoIni.SelStart = 0
    Me.mskProgramadoIni.SelLength = 50
End Sub

Private Sub mskProgramadoIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskProgramadoFin.SetFocus
    End If
End Sub

Private Sub mskProgramadoIni_LostFocus()
    Me.mskEjecutadoIni.Text = Me.mskProgramadoIni.Text
End Sub

Private Sub mskSolicitudFin_GotFocus()
    mskSolicitudFin.SelStart = 0
    mskSolicitudFin.SelLength = 50
End Sub

Private Sub mskSolicitudFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.FraProgramado.Enabled Then
            Me.mskProgramadoIni.SetFocus
        Else
            Me.cmdGrabar.SetFocus
        End If
    End If
End Sub

Private Sub mskSolicitudFin_LostFocus()
    Me.mskProgramadoFin.Text = mskSolicitudFin.Text
    Me.mskEjecutadoFin.Text = mskSolicitudFin.Text
End Sub

Private Sub mskSolicitudIni_GotFocus()
    mskSolicitudIni.SelStart = 0
    mskSolicitudIni.SelLength = 50
End Sub

Private Sub mskSolicitudIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskSolicitudFin.SetFocus
    End If
End Sub

Private Sub mskSolicitudIni_LostFocus()
    Me.mskProgramadoIni.Text = mskSolicitudIni.Text
    Me.mskEjecutadoIni.Text = mskSolicitudIni.Text
End Sub

Private Sub RRHH_Click()
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    Set oRRHH = New DActualizaDatosRRHH
    Set oPersona = New UPersona
    Set oPersona = frmBuscaPersona.Inicio(True)
    If Not oPersona Is Nothing Then
        Limpia
        Me.RRHH.psCodigoPersona = oPersona.sPersCod
        Me.RRHH.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(oPersona.sPersCod)
        RRHH_KeyPress 13
    End If
End Sub

Private Sub RRHH_cmdRecodatorioClick()
    'MsgBox "HOla"
    
End Sub

Private Sub RRHH_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oRRHH As DActualizaDatosRRHH
        Dim rsR As ADODB.Recordset
        Set oRRHH = New DActualizaDatosRRHH
        RRHH.psCodigoEmpleado = Left(RRHH.psCodigoEmpleado, 1) & Format(Trim(Mid(RRHH.psCodigoEmpleado, 2)), "00000")
        Dim oCon As DActualizaDatosContrato
        Set oCon = New DActualizaDatosContrato
        
        Set rsR = oRRHH.GetRRHH(RRHH.psCodigoEmpleado, gPersIdDNI, 3)
        
        If Not (rsR.EOF And rsR.BOF) Then
            RRHH.SpinnerValor = CInt(Right(RRHH.psCodigoEmpleado, 5))
            If rsR.Fields("Control") Then
                RRHH.psCodigoPersona = rsR.Fields("Codigo")
                RRHH.psNombreEmpledo = rsR.Fields("Nombre")
                RRHH.psDireccionPersona = rsR.Fields("Direccion")
                RRHH.psDNIPersona = IIf(IsNull(rsR.Fields("ID")), "", rsR.Fields("ID"))
                RRHH.psSueldoContrato = Format(rsR.Fields("Sueldo"), "#,##0.00")
                RRHH.psFechaNacimiento = Format(rsR.Fields("Fecha"), gsFormatoFechaView)
                Me.lblFechaIngresoL.Caption = Format(IIf(IsNull(rsR.Fields("Ingreso")), "__/__/_____", Format(rsR.Fields("Ingreso"), gsFormatoFechaView)))
                If cmdNuevo.Enabled Then Me.cmdNuevo.SetFocus
                CargaDatos
            Else
                MsgBox "No se puede controlar al RRHH " & rsR.Fields("Nombre") & ", pues, pertenece a un tipo no configurado para ser controlado por esta opcion.", vbInformation, "Aviso"
                Limpia
                RRHH.ClearScreen
                Set oRRHH = Nothing
                RRHH.SetFocus
            End If
        Else
            MsgBox "Codigo no Reconocido.", vbInformation, "Aviso"
            Limpia
            RRHH.ClearScreen
            Set oRRHH = Nothing
            RRHH.SetFocus
        End If
        Set oRRHH = Nothing
    End If
End Sub

Private Sub txtComentario_GotFocus()
    txtComentario.SelStart = 0
    txtComentario.SelLength = 50
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.fraSolicitud.Enabled Then
            Me.mskSolicitudIni.SetFocus
        ElseIf Me.FraProgramado.Enabled Then
            Me.mskProgramadoIni.SetFocus
        ElseIf Me.fraEjecutado.Enabled Then
            Me.mskEjecutadoIni.SetFocus
        Else
            Me.cmbEstado.SetFocus
        End If
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Public Sub Ini(pTipoAut As RHAutoFisicaGrupo, pTipoEstAut As RHAutoFisicaTipoEstado, pTipoOpe As TipoOpe, psMensaje As String)
    lnTipo = pTipoAut
    lnTipoEstAut = pTipoEstAut
    lnTipoOpe = pTipoOpe
    Caption = psMensaje
    Me.Show 1
End Sub

