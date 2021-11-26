VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRHInformeSocial 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   Icon            =   "frmRHInformeSocial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Sicmact.ctrRRHHGen ctrRRHH 
      Height          =   1200
      Left            =   15
      TabIndex        =   7
      Top             =   15
      Width           =   8205
      _ExtentX        =   14473
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
   Begin TabDlg.SSTab Tab 
      Height          =   4155
      Left            =   30
      TabIndex        =   8
      Top             =   1275
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   7329
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   10485760
      TabCaption(0)   =   "Informe Social"
      TabPicture(0)   =   "frmRHInformeSocial.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraInfSoc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Texto"
      TabPicture(1)   =   "frmRHInformeSocial.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTexto"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraTexto 
         Caption         =   "Texto"
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
         Height          =   3675
         Left            =   -74880
         TabIndex        =   18
         Top             =   360
         Width           =   8070
         Begin VB.CommandButton cmdCargarArchivo 
            Caption         =   "&Carga Archivo"
            Height          =   375
            Left            =   3255
            TabIndex        =   19
            Top             =   3225
            Width           =   1545
         End
         Begin RichTextLib.RichTextBox RichInforme 
            Height          =   2910
            Left            =   75
            TabIndex        =   20
            Top             =   240
            Width           =   7905
            _ExtentX        =   13944
            _ExtentY        =   5133
            _Version        =   393217
            ReadOnly        =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmRHInformeSocial.frx":0342
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame fraInfSoc 
         Caption         =   "Informe Social"
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
         Height          =   3735
         Left            =   120
         TabIndex        =   9
         Top             =   330
         Width           =   8040
         Begin Sicmact.TxtBuscar txtReferencia 
            Height          =   315
            Left            =   990
            TabIndex        =   14
            Top             =   465
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            EditFlex        =   -1  'True
            sTitulo         =   ""
         End
         Begin VB.TextBox txtResumen 
            Appearance      =   0  'Flat
            Height          =   2055
            Left            =   1005
            TabIndex        =   10
            Top             =   1560
            Width           =   6900
         End
         Begin Sicmact.TxtBuscar txtTipo 
            Height          =   315
            Left            =   990
            TabIndex        =   16
            Top             =   1035
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            sTitulo         =   ""
         End
         Begin VB.Label lblTipoRes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3090
            TabIndex        =   17
            Top             =   1035
            Width           =   4785
         End
         Begin VB.Label lblRefRes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3090
            TabIndex        =   15
            Top             =   465
            Width           =   4785
         End
         Begin VB.Label lblTipo 
            Caption         =   "Tipo :"
            Height          =   240
            Left            =   150
            TabIndex        =   13
            Top             =   1065
            Width           =   810
         End
         Begin VB.Label lblResumen 
            Caption         =   "Resumen :"
            Height          =   225
            Left            =   150
            TabIndex        =   12
            Top             =   1530
            Width           =   930
         End
         Begin VB.Label lblDetalle 
            Caption         =   "Referenc :"
            Height          =   225
            Left            =   150
            TabIndex        =   11
            Top             =   495
            Width           =   915
         End
      End
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   5565
      Top             =   5430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7170
      TabIndex        =   5
      Top             =   5490
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3660
      TabIndex        =   3
      Top             =   5490
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2460
      TabIndex        =   2
      Top             =   5490
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1260
      TabIndex        =   1
      Top             =   5490
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   5490
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   5490
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1260
      TabIndex        =   6
      Top             =   5490
      Width           =   1095
   End
End
Attribute VB_Name = "frmRHInformeSocial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbEditado As Boolean
Dim lnTipo As TipoOpe

Public Sub Ini(pnTipo As TipoOpe, psCaption As String)
    lnTipo = pnTipo
    Caption = psCaption
    Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
    lbEditado = False
    Limpia
    Activa False
End Sub

Private Sub cmdCargarArchivo_Click()
    Dim lsArch As String
    CDialog.CancelError = False
    On Error GoTo ErrHandler
    CDialog.Flags = cdlOFNHideReadOnly
    CDialog.Filter = "Archivos de texto(*.txt)|*.txt"
    CDialog.FilterIndex = 2
    CDialog.ShowOpen
    lsArch = CDialog.FileName
    Me.RichInforme.LoadFile lsArch, 1
    Me.RichInforme.Text = FiltroCadena(Me.RichInforme.Text)
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub cmdEditar_Click()
    If Me.ctrRRHH.psCodigoPersona = "" Or Me.txtReferencia.Text = "" Then Exit Sub
    lbEditado = True
    Activa True
End Sub

Private Sub cmdEliminar_Click()
    Dim oInf As NActualizaDatosInformeSocial
    Set oInf = New NActualizaDatosInformeSocial
    
    If Me.ctrRRHH.psCodigoPersona = "" Or Me.txtReferencia.Text = "" Then Exit Sub
    If MsgBox("Desea eliminar el Registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    oInf.EliminaInformeSocial Me.ctrRRHH.psCodigoPersona, Format(Me.txtReferencia.Text, gsFormatoFechaHora)
    Limpia
    Set oInf = Nothing
    CargaData Me.ctrRRHH.psCodigoPersona
End Sub

Private Sub CmdGrabar_Click()
    If Not Valida Then Exit Sub
    
    Dim oInf As NActualizaDatosInformeSocial
    Set oInf = New NActualizaDatosInformeSocial
    
    If Not lbEditado Then
        oInf.AgregaInformeSocial Me.ctrRRHH.psCodigoPersona, FechaHora(gdFecSis), Me.txtTipo.Text, Me.RichInforme.Text, Me.txtResumen.Text, GetMovNro(gsCodUser, gsCodAge)
    Else
        oInf.ModificaInformeSocial Me.ctrRRHH.psCodigoPersona, Format(Me.txtReferencia.Text, gsFormatoFechaHora), txtTipo.Text, Me.RichInforme.Text, Me.txtResumen.Text, GetMovNro(gsCodUser, gsCodAge)
    End If
        
    Set oInf = Nothing
        
    Activa False
    lbEditado = False
    Limpia
    CargaData Me.ctrRRHH.psCodigoPersona
End Sub

Private Sub cmdImprimir_Click()
    Dim oPrevio As Previo.clsPrevio
    Dim lsCadena As String
    Dim oInf As NActualizaDatosInformeSocial
    Set oInf = New NActualizaDatosInformeSocial
    Set oPrevio = New Previo.clsPrevio
    
    If Me.ctrRRHH.psCodigoPersona = "" Then Exit Sub

    lsCadena = oInf.ReporteInformeSocial(Me.ctrRRHH.psCodigoPersona, 1, Caption & " - " & Me.ctrRRHH.psNombreEmpledo, gsNomAge, gsEmpresa, gdFecSis)
    
    oPrevio.Show lsCadena, Caption, True, 66
    Set oPrevio = Nothing
End Sub

Private Sub cmdNuevo_Click()
    If Me.ctrRRHH.psCodigoPersona = "" Then
        Me.ctrRRHH.SetFocus
        Exit Sub
    End If
    lbEditado = False
    Limpia False
    Activa True
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.cmdNuevo.Visible = Not pbValor
    Me.cmdSalir.Enabled = Not pbValor
    Me.cmdEditar.Visible = Not pbValor
    Me.cmdGrabar.Visible = pbValor
    Me.cmdCancelar.Visible = pbValor
    Me.cmdCargarArchivo.Enabled = pbValor
    Me.ctrRRHH.Enabled = Not pbValor
    
    
    If lnTipo = gTipoOpeRegistro Then
        Me.cmdEditar.Enabled = pbValor
        Me.cmdEliminar.Enabled = False
        Me.cmdImprimir.Enabled = False
        Me.fraInfSoc.Enabled = pbValor
        Me.txtReferencia.Enabled = Not pbValor
    ElseIf lnTipo = gTipoOpeMantenimiento Then
        Me.cmdEliminar.Enabled = Not pbValor
        Me.cmdImprimir.Enabled = False
        Me.cmdNuevo.Enabled = False
        Me.txtReferencia.Enabled = Not pbValor
        Me.txtTipo.Enabled = pbValor
        Me.txtResumen.Enabled = pbValor
        
    ElseIf lnTipo = gTipoOpeConsulta Then
        Me.cmdEliminar.Enabled = False
        Me.cmdNuevo.Enabled = False
        Me.cmdEditar.Enabled = False
        Me.cmdImprimir.Enabled = False
    ElseIf lnTipo = gTipoOpeReporte Then
        Me.cmdEliminar.Enabled = pbValor
        Me.cmdNuevo.Enabled = pbValor
        Me.cmdEditar.Enabled = pbValor
        Me.fraInfSoc.Enabled = False
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
        CargaData Me.ctrRRHH.psCodigoPersona
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
            CargaData Me.ctrRRHH.psCodigoPersona
            If cmdEditar.Enabled And cmdEditar.Visible Then
                Me.cmdEditar.SetFocus
            End If
        Else
            MsgBox "Codigo no Reconocido.", vbInformation, "Aviso"
            'ClearScreen
            ctrRRHH.SetFocus
        End If
        
        rsR.Close
        Set rsR = Nothing
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub Form_Load()
    CargaDatos
    Activa False
End Sub

Private Function Valida() As Boolean
    'valida
    If Me.txtTipo.Text = "" Then
        MsgBox "Debe Ingresar un tipo de Informe Social.", vbInformation, "Aviso"
        Me.txtTipo.SetFocus
        Valida = False
    ElseIf Me.txtResumen.Text = "" Then
        MsgBox "Debe Ingresar un Comentario.", vbInformation, "Aviso"
        Me.txtResumen.SetFocus
        Valida = False
    Else
        Valida = True
    End If
End Function

Private Sub Limpia(Optional pbTodo As Boolean = True)
    Me.txtTipo.Text = ""
    Me.txtResumen.Text = ""
    Me.txtReferencia.Text = ""
    Me.RichInforme.Text = ""
    Me.lblRefRes.Caption = ""
    Me.lblTipoRes.Caption = ""
End Sub

Private Sub CargaDatos()
    Dim oCon As DConstantes
    Dim rsC As ADODB.Recordset
    Set oCon = New DConstantes
    Set rsC = New ADODB.Recordset
    
    Set rsC = oCon.GetConstante(6024, , , True)
    Me.txtTipo.rs = rsC
    
    Set oCon = Nothing
End Sub

Private Sub RichInforme_DblClick()
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    oPrevio.Show Me.RichInforme.Text, "Contrato", True, 66
    Set oPrevio = Nothing
End Sub

'Private Sub RRHH_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Dim oRRHH As DActualizaDatosRRHH
'        Dim rsR As ADODB.Recordset
'        Set oRRHH = New DActualizaDatosRRHH
'        RRHH.psCodigoEmpleado = Left(RRHH.psCodigoEmpleado, 1) & Format(Trim(Mid(RRHH.psCodigoEmpleado, 2)), "00000")
'        Dim oCon As DActualizaDatosContrato
'        Set oCon = New DActualizaDatosContrato
'
'        Set rsR = oRRHH.GetRRHH(RRHH.psCodigoEmpleado, gPersIdDNI)
'
'        If Not (rsR.EOF And rsR.BOF) Then
'            RRHH.SpinnerValor = CInt(Right(RRHH.psCodigoEmpleado, 5))
'            RRHH.psCodigoPersona = rsR.Fields("Codigo")
'            RRHH.psNombreEmpledo = rsR.Fields("Nombre")
'            RRHH.psDireccionPersona = rsR.Fields("Direccion")
'            RRHH.psDNIPersona = IIf(IsNull(rsR.Fields("ID")), "", rsR.Fields("ID"))
'            RRHH.psSueldoContrato = Format(rsR.Fields("Sueldo"), "#,##0.00")
'            RRHH.psFechaNacimiento = Format(rsR.Fields("Fecha"), gsFormatoFechaView)
'            If cmdNuevo.Enabled Then Me.cmdNuevo.SetFocus
'        Else
'            MsgBox "Codigo no Reconocido.", vbInformation, "Aviso"
'            Limpia
'            RRHH.SetFocus
'        End If
'
'    rsI.Close
'    Set rsI = Nothing
'        rsR.Close
'        Set rsR = Nothing
'    Else
'        KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    End If
'End Sub

Private Sub txtReferencia_EmiteDatos()
    Me.lblRefRes.Caption = Me.txtReferencia.psDescripcion
    
    Dim oInf As DActualizaDatosInformeSocial
    Set oInf = New DActualizaDatosInformeSocial
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = oInf.GetValorInformeSocial(ctrRRHH.psCodigoPersona, Format(Me.txtReferencia.Text, gsFormatoFechaHora))
    If Not (rs.EOF And rs.BOF) Then
        Me.txtTipo.Text = rs!Tpo
        Me.lblTipoRes.Caption = txtTipo.psDescripcion
        Me.txtResumen.Text = rs!Coment
        If Not IsNull(rs!texto) Then
            Me.RichInforme.Text = rs!texto
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub txtResumen_GotFocus()
    txtResumen.SelStart = 0
    txtResumen.SelLength = 50
End Sub

Private Sub txtResumen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.cmdCargarArchivo.Enabled Then Me.cmdCargarArchivo.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtTipo_EmiteDatos()
    Me.lblTipoRes.Caption = Me.txtTipo.psDescripcion
    Me.txtResumen.SetFocus
End Sub

Private Sub CargaData(psCodPers As String)
    Dim oInf As DActualizaDatosInformeSocial
    Dim rsI As ADODB.Recordset
    Set rsI = New ADODB.Recordset
    Set oInf = New DActualizaDatosInformeSocial
    Set rsI = oInf.GetInformesSociales(psCodPers, 1)
    txtReferencia.rs = rsI
End Sub
