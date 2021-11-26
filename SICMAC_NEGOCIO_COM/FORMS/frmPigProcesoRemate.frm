VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmPigProcesoRemate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proceso de Remate"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   Icon            =   "frmPigProcesoRemate.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox RtfCartas 
      Height          =   150
      Left            =   5655
      TabIndex        =   21
      Top             =   5235
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   265
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmPigProcesoRemate.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraContenedor 
      Height          =   5190
      Index           =   0
      Left            =   75
      TabIndex        =   16
      Top             =   0
      Width           =   6795
      Begin VB.CommandButton cmdRepMartillero 
         Caption         =   "Reporte de Martillero"
         Enabled         =   0   'False
         Height          =   360
         Left            =   705
         TabIndex        =   20
         Top             =   4395
         Width           =   4890
      End
      Begin VB.TextBox txtNumRemate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   930
         TabIndex        =   2
         Top             =   390
         Width           =   750
      End
      Begin VB.CommandButton cmdSeleccion 
         Caption         =   "Selección de Contratos"
         Enabled         =   0   'False
         Height          =   360
         Left            =   720
         TabIndex        =   7
         Top             =   1185
         Width           =   4890
      End
      Begin VB.CommandButton cmdPrimeraSubasta 
         Caption         =   "Fin de Retasación - Calculo de Deuda 1era Subasta"
         Enabled         =   0   'False
         Height          =   360
         Left            =   720
         TabIndex        =   8
         Top             =   2130
         Width           =   4890
      End
      Begin VB.CommandButton cmdSegundaSubasta 
         Caption         =   "Generación 2da Subasta"
         Enabled         =   0   'False
         Height          =   360
         Left            =   720
         TabIndex        =   9
         Top             =   2595
         Width           =   4890
      End
      Begin VB.CommandButton cmdTerceraSubasta 
         Caption         =   "Generación 3era Subasta"
         Enabled         =   0   'False
         Height          =   360
         Left            =   705
         TabIndex        =   10
         Top             =   3030
         Width           =   4890
      End
      Begin VB.CommandButton cmdAdjudicacion 
         Caption         =   "Proceso de Adjudicación (Cliente, Tasador, Caja)"
         Enabled         =   0   'False
         Height          =   360
         Left            =   690
         TabIndex        =   11
         Top             =   3495
         Width           =   4890
      End
      Begin VB.CommandButton cmdRetasacionMan 
         Caption         =   "Retasación Manual"
         Enabled         =   0   'False
         Height          =   360
         Left            =   720
         TabIndex        =   13
         Top             =   1665
         Width           =   4890
      End
      Begin VB.CommandButton cmdRegistroSobrante 
         Caption         =   "Registro de Sobrantes"
         Enabled         =   0   'False
         Height          =   360
         Left            =   705
         TabIndex        =   12
         Top             =   3945
         Width           =   4890
      End
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   315
         Left            =   3015
         TabIndex        =   4
         Top             =   360
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskFecFin 
         Height          =   315
         Left            =   5415
         TabIndex        =   6
         Top             =   360
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecRef 
         Caption         =   "Label1"
         Height          =   30
         Left            =   3195
         TabIndex        =   19
         Top             =   5175
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Line Line1 
         X1              =   135
         X2              =   6660
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fecha Fin :"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   4530
         TabIndex        =   5
         Top             =   435
         Width           =   855
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Número :"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   7
         Left            =   195
         TabIndex        =   1
         Top             =   420
         Width           =   660
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fecha Inicio :"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   6
         Left            =   1995
         TabIndex        =   3
         Top             =   420
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5805
      TabIndex        =   17
      Top             =   5370
      Width           =   975
   End
   Begin VB.Frame fraImpresion 
      Caption         =   "Impresión"
      Height          =   540
      Left            =   60
      TabIndex        =   0
      Top             =   5265
      Visible         =   0   'False
      Width           =   2460
      Begin VB.OptionButton optImpresion 
         Caption         =   "Impresora"
         Height          =   225
         Index           =   1
         Left            =   1230
         TabIndex        =   15
         Top             =   210
         Width           =   990
      End
      Begin VB.OptionButton optImpresion 
         Caption         =   "Pantalla"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   210
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin MSComctlLib.ProgressBar prgList 
      Height          =   330
      Left            =   2805
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmPigProcesoRemate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim lnTipoProceso As Integer
'
'Private Sub cmdAdjudicacion_Click()
'Dim cmd As ADODB.Command
'Dim prm As ADODB.Parameter
'Dim oConn As DConecta
'Dim oValida As nPigValida
'
'    Set oValida = New nPigValida
'    If oValida.ValidaAdjudicacion(gdFecSis) Then
'
'        If Not oValida.ValidaSiRemesaAdjudica(txtNumRemate.Text) Then
'            MsgBox "Existen Piezas no vendidas en el Remate que no han sido Remesadas a la Boveda de Valores. Realizar la Remesa antes de proceder a la Adjudicacion", vbInformation, "Aviso"
'            Exit Sub
'        End If
'
'        'On Error GoTo Error
'
'        Set cmd = New ADODB.Command
'        Set prm = New ADODB.Parameter
'
'        Set oConn = New DConecta
'        oConn.AbreConexion
'        oConn.ConexionActiva.BeginTrans
'
'        cmd.CommandText = "ColocPigCompensacionDeuda"
'        cmd.CommandType = adCmdStoredProc
'        cmd.Name = "ColocPigCompensacionDeuda"
'        Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
'        cmd.Parameters.Append prm
'        Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
'        cmd.Parameters.Append prm
'        Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
'        cmd.Parameters.Append prm
'        Set cmd.ActiveConnection = oConn.ConexionActiva
'        cmd.CommandTimeout = 720
'        cmd.Parameters.Refresh
'
'        oConn.ConexionActiva.ColocPigCompensacionDeuda Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), gsCodUser, Right(gsCodAge, 2)
'
'        Set cmd = Nothing
'        Set prm = Nothing
'
'        Set cmd = New ADODB.Command
'        Set prm = New ADODB.Parameter
'
'        cmd.CommandText = "ColocPigAdjudica"
'        cmd.CommandType = adCmdStoredProc
'        cmd.Name = "ColocPigAdjudica"
'        Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
'        cmd.Parameters.Append prm
'        Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
'        cmd.Parameters.Append prm
'        Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
'        cmd.Parameters.Append prm
'        Set cmd.ActiveConnection = oConn.ConexionActiva
'        cmd.CommandTimeout = 720
'        cmd.Parameters.Refresh
'
'        oConn.ConexionActiva.ColocPigAdjudica Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), gsCodUser, Right(gsCodAge, 2)
'
'        oConn.ConexionActiva.CommitTrans
'
'        Set cmd = Nothing
'        Set prm = Nothing
'
'        MsgBox "El Proceso de Generación de Adjudicación finalizó satisfactoriamente"
'        oConn.CierraConexion
'        HabilitaControles False, False, False, False, False, False, True, True
'
'    Else
'        MsgBox "El Proceso de Adjudicación ya se efectuó", vbInformation, "Aviso"
'        Exit Sub
'    End If
'
'    Set oValida = Nothing
'
'End Sub
'
'Private Sub cmdPrimeraSubasta_Click()
'Dim cmd As ADODB.Command
'Dim prm As ADODB.Parameter
'Dim oConn As DConecta
'Dim oValida As nPigValida
'
'    Set oValida = New nPigValida
'
'    If oValida.ValidaPrimSubasta(gdFecSis) Then
'
'        On Error GoTo Error
'
'        Set cmd = New ADODB.Command
'        Set prm = New ADODB.Parameter
'
'        Set oConn = New DConecta
'        oConn.AbreConexion
'        oConn.ConexionActiva.BeginTrans
'
'        cmd.CommandText = "ColocPigFinRetasacion"
'        cmd.CommandType = adCmdStoredProc
'        cmd.Name = "ColocPigFinRetasacion"
'        Set prm = cmd.CreateParameter("NumRemate", adInteger, adParamInput)
'        cmd.Parameters.Append prm
'        Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
'        cmd.Parameters.Append prm
'        Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
'        cmd.Parameters.Append prm
'        Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
'        cmd.Parameters.Append prm
'        Set cmd.ActiveConnection = oConn.ConexionActiva
'        cmd.CommandTimeout = 720
'        cmd.Parameters.Refresh
'
'        oConn.ConexionActiva.ColocPigFinRetasacion txtNumRemate, Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), gsCodUser, Right(gsCodAge, 2)
'
'        oConn.ConexionActiva.CommitTrans
'
'        Set cmd = Nothing
'        Set prm = Nothing
'
'        MsgBox "El Proceso de Fin de Retasación finalizó satisfactoriamente"
'        lnTipoProceso = 1
'        HabilitaControles False, False, False, True, False, False, False, True
'        oConn.CierraConexion
'
'    Else
'        MsgBox "El Proceso de Fin de Retasación ya se efectuó"
'        Exit Sub
'    End If
'
'    Set oValida = Nothing
'
'    Exit Sub
'
'Error:
'    MsgBox Err.raiser + Err.Number, Err.Description, "Aviso"
'
'End Sub
'
'Private Sub cmdRegistroSobrante_Click()
'Dim oPigDatos As DPigContrato
'Dim oContFunc As NContFunciones
'Dim oPigRemate As NPigRemate
'Dim oPigImpre As NPigImpre
'Dim oPrevio As previo.clsPrevio
'Dim lsMovNro As String
'Dim rs As Recordset
'Dim sCadImpre As String
'
'    Set oPigDatos = New DPigContrato
'    Set rs = oPigDatos.dSeleccionaSobrantes(txtNumRemate)
'    Set oPigDatos = Nothing
'
'    Set oPigRemate = New NPigRemate
'        Call oPigRemate.nPigRegistraSobrante(rs, CInt(txtNumRemate), gdFecSis, gsCodAge, gsCodUser)
'
'    Set oPigRemate = Nothing
'    Set rs = Nothing
'
'    '======= Impresion de Avisos de Sobrante de Remate
'    RtfCartas.FileName = App.path & "\FormatoCarta\CartaAvisoSobranteRemate.txt"
'    Set oPigImpre = New NPigImpre
'    sCadImpre = oPigImpre.ImpreAvisoSobrante(CInt(txtNumRemate), gdFecSis, RtfCartas.Text)
'    Set oPrevio = New previo.clsPrevio
'    oPrevio.Show sCadImpre, "Avisos de Sobrante de Remate", True, 66
'    Set oPrevio = Nothing
'
'    Set oPigImpre = Nothing
'
'    MsgBox "El Proceso de Registro de Sobrante de Remate finalizó satisfactoriamente"
'    HabilitaControles False, False, False, False, False, False, False, True
'
'End Sub
'
'Private Sub cmdRepMartillero_Click()
'Dim oImpre As NPigImpre
'Dim lsCadImpre As String
'Dim oPrevio As previo.clsPrevio
'
'
'    If MsgBox("Imprimir Reporte de Martillero ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'        Screen.MousePointer = 11
'        Set oImpre = New NPigImpre
'        If lnTipoProceso = 0 Then lnTipoProceso = 1
'        lsCadImpre = oImpre.ImpreRepMartillero(gsNomCmac, gsNomAge, txtNumRemate, lnTipoProceso, "", gdFecSis)
'        Set oImpre = Nothing
'
'        Set oPrevio = New previo.clsPrevio
'        oPrevio.Show lsCadImpre, "Reporte del Martillero", True, 66
'
'        Do While True
'            If MsgBox("Reimprimir Reporte de Martillero? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'                oPrevio.Show lsCadImpre, "Reporte del Martillero", True, 66
'            Else
'                Set oPrevio = Nothing
'                Exit Do
'            End If
'        Loop
'    End If
'
'    Screen.MousePointer = 0
'
'End Sub
'
'Private Sub cmdRetasacionMan_Click()
'
'    FrmPigRetasacionMan.txtnremate = txtNumRemate
'    FrmPigRetasacionMan.lblFecIniRem = mskFecIni
'    FrmPigRetasacionMan.Show 1
'
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdSegundaSubasta_Click()
'Dim cmd As ADODB.Command
'Dim prm As ADODB.Parameter
'Dim oConn As DConecta
'Dim oValida As nPigValida
'
'    Set oValida = New nPigValida
'    If oValida.ValidaSegTerSubasta(gdFecSis, 2) Then
'        Set cmd = New ADODB.Command
'        Set prm = New ADODB.Parameter
'
'        Set oConn = New DConecta
'        oConn.AbreConexion
'        oConn.ConexionActiva.BeginTrans
'
'        cmd.CommandText = "ColocPigSegundaSubasta"
'        cmd.CommandType = adCmdStoredProc
'        cmd.Name = "ColocPigSegundaSubasta"
'        Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
'        cmd.Parameters.Append prm
'        Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
'        cmd.Parameters.Append prm
'        Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
'        cmd.Parameters.Append prm
'        Set cmd.ActiveConnection = oConn.ConexionActiva
'        cmd.CommandTimeout = 720
'        cmd.Parameters.Refresh
'
'        oConn.ConexionActiva.ColocPigSegundaSubasta Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), gsCodUser, Right(gsCodAge, 2)
'
'        oConn.ConexionActiva.CommitTrans
'
'        Set cmd = Nothing
'        Set prm = Nothing
'
'        MsgBox "El Proceso de Generación de Segunda Subasta finalizó satisfactoriamente"
'        lnTipoProceso = 2
'        HabilitaControles False, False, False, False, True, True, False, True
'        oConn.CierraConexion
'
'    Else
'        MsgBox "No se puede realizar el Proceso de Generación de Segunda Subasta"
'        Exit Sub
'    End If
'
'    Set oValida = Nothing
'
'End Sub
'
'Private Sub cmdSeleccion_Click()
'
'    frmPigSeleccionRemate.txtremate = txtNumRemate
'    frmPigSeleccionRemate.MskIniRemate = mskFecIni
'    frmPigSeleccionRemate.Show 1
'
'End Sub
'
'
'Private Sub cmdTerceraSubasta_Click()
'Dim cmd As ADODB.Command
'Dim prm As ADODB.Parameter
'Dim oConn As DConecta
'Dim oValida As nPigValida
'
'    Set oValida = New nPigValida
'
'    If oValida.ValidaSegTerSubasta(gdFecSis, 3) Then
'        Set cmd = New ADODB.Command
'        Set prm = New ADODB.Parameter
'
'        Set oConn = New DConecta
'        oConn.AbreConexion
'        oConn.ConexionActiva.BeginTrans
'
'        cmd.CommandText = "ColocPigTerceraSubasta"
'        cmd.CommandType = adCmdStoredProc
'        cmd.Name = "ColocPigTerceraSubasta"
'        Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
'        cmd.Parameters.Append prm
'        Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
'        cmd.Parameters.Append prm
'        Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
'        cmd.Parameters.Append prm
'        Set cmd.ActiveConnection = oConn.ConexionActiva
'        cmd.CommandTimeout = 720
'        cmd.Parameters.Refresh
'
'        oConn.ConexionActiva.ColocPigTerceraSubasta Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), gsCodUser, Right(gsCodAge, 2)
'
'        oConn.ConexionActiva.CommitTrans
'
'        Set cmd = Nothing
'        Set prm = Nothing
'
'        MsgBox "El Proceso de Generación de Tercera Subasta finalizó satisfactoriamente"
'        lnTipoProceso = 3
'        HabilitaControles False, False, False, False, False, True, False, True
'        oConn.CierraConexion
'
'    Else
'        MsgBox "No se puede realizar el Proceso de Generación de Tercera Subasta"
'        Exit Sub
'    End If
'
'    Set oValida = Nothing
'
'End Sub
'
'Private Sub Form_Load()
'Dim oRemate As DPigRemate
'Dim rs As Recordset
'
'    Set oRemate = New DPigRemate
'    Set rs = oRemate.GetNumRemate
'
'    If Not (rs.EOF And rs.BOF) Then
'        txtNumRemate = rs!NumRemate
'        lblFecRef = Format(rs!dReferencia, "dd/mm/yyyy")
'        mskFecIni = Format(rs!dInicio, "dd/mm/yyyy")
'        MskFecFin = Format(rs!dFin, "dd/mm/yyyy")
'        lnTipoProceso = rs!nTipoProceso
'    End If
'
'    Set rs = Nothing
'    Set oRemate = Nothing
'
'End Sub
'
'Private Sub txtNumRemate_KeyPress(KeyAscii As Integer)
'Dim oRemate As DPigRemate
'Dim rs As Recordset
'
'    If KeyAscii = 13 Then
'        If txtNumRemate <> "" Then
'            Set oRemate = New DPigRemate
'            Set rs = oRemate.GetNumRemate(txtNumRemate)
'
'            If Not (rs.EOF And rs.BOF) Then
''                 If CInt(rs!cUbicacion) <> CInt(Right(gsCodAge, 2)) Then
''                     MsgBox "Usuario no se encuentra asignado a la Agencia donde se efectuara el Remate", vbInformation, "Aviso"
''                     Exit Sub
''                End If
'                txtNumRemate = rs!NumRemate
'                lblFecRef = Format(rs!dReferencia, "dd/mm/yyyy")
'                mskFecIni = Format(rs!dInicio, "dd/mm/yyyy")
'                MskFecFin = Format(rs!dFin, "dd/mm/yyyy")
'                lnTipoProceso = rs!nTipoProceso
'
'                Select Case rs!nTipoProceso
'                Case 0
'                    Call HabilitaControles(True, True, True, False, False, False, False, True)
'                Case 1  'Segunda Subasta
'                   Call HabilitaControles(False, False, False, True, False, False, False, True)
'                Case 2  'Tercera Subasta
'                   Call HabilitaControles(False, False, False, False, True, True, False, True)
'                Case 3
'                    If rs!nAdjudicado = 0 Then
'                        Call HabilitaControles(False, False, False, False, False, True, False, True)
'                    ElseIf rs!nSobrante = 0 Then
'                        Call HabilitaControles(False, False, False, False, False, False, True, True)
'                    Else
'                        Call HabilitaControles(False, False, False, False, False, False, False, True)
'                    End If
'                End Select
'
'            Else
'                MsgBox "Número de Remate no valido", vbInformation, "Aviso"
'            End If
'
'            Set rs = Nothing
'            Set oRemate = Nothing
'
'        End If
'    End If
'
'End Sub
'
'Private Sub HabilitaControles(ByVal pbSR As Boolean, ByVal pbRM As Boolean, ByVal pbPS As Boolean, ByVal pbSS As Boolean, '        ByVal pbTS As Boolean, ByVal pbAd As Boolean, ByVal pbRS As Boolean, ByVal pbIM As Boolean)
'
'    cmdSeleccion.Enabled = pbSR
'    cmdRetasacionMan.Enabled = pbRM
'    cmdPrimeraSubasta.Enabled = pbPS
'    cmdSegundaSubasta.Enabled = pbSS
'    cmdTerceraSubasta.Enabled = pbTS
'    cmdAdjudicacion.Enabled = pbAd
'    cmdRegistroSobrante.Enabled = pbRS
'    cmdRepMartillero.Enabled = pbIM
'
'End Sub
