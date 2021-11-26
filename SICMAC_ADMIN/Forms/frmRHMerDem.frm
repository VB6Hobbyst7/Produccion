VERSION 5.00
Begin VB.Form frmRHMerDem 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "frmRHMerDem.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8130
      TabIndex        =   8
      Top             =   4695
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2415
      TabIndex        =   3
      Top             =   4695
      Width           =   1095
   End
   Begin Sicmact.ctrRRHHGen ctrRRHHGen 
      Height          =   1200
      Left            =   15
      TabIndex        =   2
      Top             =   15
      Width           =   9180
      _ExtentX        =   16193
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
   Begin VB.Frame fraPerNoLab 
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
      ForeColor       =   &H00800000&
      Height          =   3405
      Left            =   15
      TabIndex        =   4
      Top             =   1215
      Width           =   9195
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   8010
         TabIndex        =   0
         Top             =   2955
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   6840
         TabIndex        =   1
         Top             =   2955
         Width           =   1095
      End
      Begin Sicmact.FlexEdit FlexPer 
         Height          =   2685
         Left            =   120
         TabIndex        =   5
         Top             =   225
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   4736
         Cols0           =   8
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cod.Tpo-Tipo-Fecha-Observaciones-bit-bit1-Movimiento"
         EncabezadosAnchos=   "300-800-1800-900-3000-0-0-1800"
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
         ColumnasAEditar =   "X-1-X-3-4-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-2-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-C-C-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1230
      TabIndex        =   6
      Top             =   4695
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1230
      TabIndex        =   9
      Top             =   4695
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4680
      Width           =   1095
   End
End
Attribute VB_Name = "frmRHMerDem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ALPA 20090122******************
Dim objPista As COMManejador.Pista
'*******************************
Dim lsCodigo As String
Dim lbEditado  As Boolean
Dim lnTipo As TipoOpe

Private Sub cmdCancelar_Click()
    CargaData Me.ctrRRHHGen.psCodigoPersona
    Activa False
End Sub

Private Sub cmdEditar_Click()
    If Me.ctrRRHHGen.psCodigoPersona = "" Then Exit Sub
    Activa True
    Me.cmdNuevo.SetFocus
End Sub

Private Sub CmdEliminar_Click()
    FlexPer.EliminaFila FlexPer.Row
End Sub

Private Sub cmdGrabar_Click()
    If Me.ctrRRHHGen.psCodigoPersona = "" Then Exit Sub
    If MsgBox("Desea Grabar ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    Dim oMer As NMeritosDemeritos
    Set oMer = New NMeritosDemeritos
    
    If Not Valida Then
        Me.FlexPer.SetFocus
        Exit Sub
    End If
    glsMovNro = GetMovNro(gsCodUser, gsCodAge)
    oMer.ModificaMerDem Me.ctrRRHHGen.psCodigoPersona, Me.FlexPer.GetRsNew, glsMovNro, gsFormatoFecha
    'ALPA 20090122 **********************************************************
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2", , Me.ctrRRHHGen.psCodigoPersona, gCodigoPersona
    '************************************************************************
    Set oMer = Nothing
    cmdCancelar_Click
End Sub

Private Sub cmdImprimir_Click()
    If Me.ctrRRHHGen.psCodigoPersona = "" Then
        MsgBox "Debe elegir a una persona.", vbInformation, "Aviso"
        Me.ctrRRHHGen.SetFocus
        Exit Sub
    End If
    
    Dim oMer As NMeritosDemeritos
    Dim oPrevio As Previo.clsPrevio
    Dim lsCadena As String
    Set oMer = New NMeritosDemeritos
    Set oPrevio = New Previo.clsPrevio
    
    lsCadena = oMer.GetReporteMerDem(Me.ctrRRHHGen.psCodigoPersona, Me.ctrRRHHGen.psNombreEmpledo, gsNomAge, gsEmpresa, gdFecSis)
    
    oPrevio.Show lsCadena, "Meritos y Demeritos", True, 66
    Set oPrevio = Nothing
End Sub

Private Sub CmdNuevo_Click()
    FlexPer.AdicionaFila
    FlexPer.SetFocus
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
    'oPersona.ObtieneClientexCodigo ()
    ClearScreen
    If Not oPersona Is Nothing Then
        ClearScreen
        Me.ctrRRHHGen.psCodigoPersona = oPersona.sPersCod
        Me.ctrRRHHGen.psNombreEmpledo = oPersona.sPersNombre
        Me.ctrRRHHGen.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(Me.ctrRRHHGen.psCodigoPersona)
        CargaData Me.ctrRRHHGen.psCodigoPersona
        If cmdEditar.Enabled And cmdEditar.Visible Then
            Me.cmdEditar.SetFocus
        Else
            Me.cmdImprimir.SetFocus
            Me.FlexPer.lbEditarFlex = False
        End If
    End If
End Sub

Private Sub CargaData(psPersCod As String)
    Dim oPNL As DMeritosDemeritos
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Set oPNL = New DMeritosDemeritos
    
    Set rsP = oPNL.GetMerDems(psPersCod)
    
    If Not (rsP.EOF And rsP.BOF) Then
        Set Me.FlexPer.Recordset = rsP
    Else
        FlexPer.Clear
        FlexPer.Rows = 2
        FlexPer.FormaCabecera
    End If
End Sub

Public Sub Ini(pnTipo As TipoOpe, psCaption As String)
    lnTipo = pnTipo
    Caption = psCaption
    Me.Show 1
End Sub

Private Sub ClearScreen()
    ctrRRHHGen.ClearScreen
    FlexPer.Clear
    FlexPer.Rows = 2
    FlexPer.FormaCabecera
End Sub

Private Sub Activa(pbvalor As Boolean)
    Me.cmdSalir.Enabled = Not pbvalor
    If lnTipo = gTipoOpeMantenimiento Then
        Me.fraPerNoLab.Enabled = pbvalor
        Me.cmdEditar.Visible = Not pbvalor
        Me.cmdGrabar.Enabled = pbvalor
        Me.cmdCancelar.Visible = pbvalor
        Me.ctrRRHHGen.Enabled = Not pbvalor
    ElseIf lnTipo = gTipoOpeConsulta Then
        Me.cmdEditar.Visible = False
        Me.cmdGrabar.Visible = False
        Me.cmdCancelar.Visible = False
        Me.cmdNuevo.Visible = False
        Me.cmdEliminar.Visible = False
    End If
End Sub

Private Function Valida() As Boolean
    Dim i As Integer
    For i = 1 To Me.FlexPer.Rows - 1
        Me.FlexPer.Row = i
        If FlexPer.TextMatrix(i, 1) = "" Then
            MsgBox "Debe Ingresar una tipo de Merito o Demerito Valido para el registro : " & Me.FlexPer.TextMatrix(i, 0), vbInformation, "Aviso"
            FlexPer.Col = 1
            Valida = False
            Exit Function
        ElseIf Not IsDate(Me.FlexPer.TextMatrix(i, 3)) Then
            MsgBox "Debe Ingresar una Fecha Valida para el registro : " & Me.FlexPer.TextMatrix(i, 0), vbInformation, "Aviso"
            FlexPer.Col = 3
            Valida = False
            Exit Function
        ElseIf Me.FlexPer.TextMatrix(i, 4) = "" Then
            MsgBox "Debe Ingresar un comentario u observacion para el registro : " & Me.FlexPer.TextMatrix(i, 0), vbInformation, "Aviso"
            FlexPer.Col = 4
            Valida = False
            Exit Function
        Else
            Valida = True
        End If
    Next i
End Function

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
            If cmdEditar.Enabled And cmdEditar.Visible Then
                Me.cmdEditar.SetFocus
            Else
                Me.cmdImprimir.SetFocus
                Me.FlexPer.lbEditarFlex = False
            End If
        Else
            MsgBox "Codigo no Reconocido.", vbInformation, "Aviso"
            ClearScreen
            ctrRRHHGen.SetFocus
        End If
        
        rsR.Close
        Set rsR = Nothing
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub Form_Load()
    Dim rsRH As ADODB.Recordset
    Set rsRH = New ADODB.Recordset
    Dim oDem As DMeritosDemeritos
    Set oDem = New DMeritosDemeritos
    'ALPA 20090122 ***************************************************************************
    gsOpeCod = LogPistaMeritoDemerito
    Set objPista = New COMManejador.Pista
    '*****************************************************************************************
    Me.FlexPer.rsTextBuscar = oDem.GetMerDemTabla(True)
    Activa False
End Sub

