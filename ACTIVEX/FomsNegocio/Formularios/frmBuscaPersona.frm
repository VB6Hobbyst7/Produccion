VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuscaPersona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Persona"
   ClientHeight    =   2775
   ClientLeft      =   2370
   ClientTop       =   2535
   ClientWidth     =   7695
   Icon            =   "frmBuscaPersona.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNomPer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      TabIndex        =   0
      Tag             =   "1"
      Top             =   450
      Width           =   3990
   End
   Begin VB.TextBox txtDocPer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      MaxLength       =   15
      TabIndex        =   8
      Tag             =   "3"
      Top             =   450
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   6
      Top             =   2265
      Width           =   1230
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   5
      Top             =   1890
      Width           =   1230
   End
   Begin VB.CommandButton cmdNewCli 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   4
      Top             =   1515
      Width           =   1230
   End
   Begin VB.Frame frabusca 
      Caption         =   "Buscar por ...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1170
      Left            =   120
      TabIndex        =   9
      Top             =   105
      Width           =   1800
      Begin VB.OptionButton optOpcion 
         Caption         =   "A&pellido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Có&digo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   540
         Width           =   1245
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Nº Docu&mento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   825
         Width           =   1635
      End
   End
   Begin VB.TextBox txtCodPer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      MaxLength       =   13
      TabIndex        =   7
      Tag             =   "2"
      Top             =   450
      Visible         =   0   'False
      Width           =   1800
   End
   Begin MSDataGridLib.DataGrid dbgrdPersona 
      Height          =   1815
      Left            =   2085
      TabIndex        =   12
      Top             =   840
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   1
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "cPersNombre"
         Caption         =   "Nombre  o Razon Social"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cPersDireccDomicilio"
         Caption         =   "Dirección"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cPersCod"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "cPersIDnroDNI"
         Caption         =   "Doc. Natural"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "cPersIDnroRUC"
         Caption         =   "Doc. Juridico"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "cPersTelefono"
         Caption         =   "Telefono"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "dPersNacCreac"
         Caption         =   "Fecha Nac."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "cCodZon"
         Caption         =   "Zona"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "cPersPersoneria"
         Caption         =   "Tipo Persona"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         Size            =   182
         BeginProperty Column00 
            ColumnWidth     =   4515.024
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1725.165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1289.764
         EndProperty
      EndProperty
   End
   Begin VB.Label LblDoc 
      Height          =   195
      Left            =   4095
      TabIndex        =   11
      Top             =   525
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese Dato a Buscar :"
      Height          =   195
      Left            =   2070
      TabIndex        =   10
      Top             =   150
      Width           =   1680
   End
End
Attribute VB_Name = "frmBuscaPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Persona As COMDPersona.UCOMPersona
Dim R As ADODB.Recordset
Dim bBuscarEmpleado As Boolean
Public vcodper As String 'MADM 20090928
Dim cOpeCod As String 'MADM 20101012

'AGREGADO POR VAPI PARA LA PRESOLITUD DE CREDITO SEGÙN ERS TI-ERS001-2017
Public Function InicioAutomatico(ByVal pcPersCod As String) As COMDPersona.UCOMPersona

    'Me.Show 1
    Dim ClsPersona As COMDPersona.DCOMPersonas
    If Len(Trim(pcPersCod)) = 0 Then
        MsgBox "Falta Ingresar el Codigo de la Persona", vbInformation, "Aviso"
        Exit Function
    End If
    Screen.MousePointer = 11
    Set ClsPersona = New COMDPersona.DCOMPersonas
    Set R = ClsPersona.BuscaCliente(pcPersCod, BusquedaCodigo)
    Set dbgrdPersona.DataSource = R
    dbgrdPersona.Refresh
    Screen.MousePointer = 0
    If R.RecordCount = 0 Then
        MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
        txtCodPer.SetFocus
        cmdAceptar.Default = False
    Else
        'dbgrdPersona.SetFocus
        cmdAceptar.Default = True
    End If
    
    'Call CmdAceptar_Click
   
   Set InicioAutomatico = Persona
   Set Persona = Nothing
End Function


'FIN AGREAGDO POR VAPI

Public Function Inicio(Optional ByVal pbBuscarEmpleado As Boolean = False, Optional pcOpeCod As String = "") As COMDPersona.UCOMPersona
   bBuscarEmpleado = pbBuscarEmpleado
   
    'MADM 20101012
   cOpeCod = pcOpeCod
   
   'RIRO20140611, Se agrego gOtrOpeEgresoDevSobranteOtrasOpeVoucher
   'If cOpecod = "300503" Then
   If cOpeCod = "300503" Or cOpeCod = CStr(gOtrOpeEgresoDevSobranteOtrasOpeChq) Or cOpeCod = CStr(gOtrOpeEgresoDevSobranteOtrasOpeVoucher) Then 'EJVG20140213
        optOpcion(0).Visible = False
        optOpcion(1).Visible = False
        cmdNewCli.Visible = False
        optOpcion(2).Value = True
   Else
        optOpcion(0).Visible = True
        optOpcion(1).Visible = True
        cmdNewCli.Visible = True
   End If
   'PTI1 20170626
   If cOpeCod = "3010245" Then
        cmdNewCli.Enabled = False
   End If
   'END PTI1 20170626
   'END MADM 20101012
   
   'If Me.ActiveControl Is Nothing Then
    '    Me.Show vbModal
   'Else
     '   Me.Show
  ' End If
  
   Me.Show 1
   Set Inicio = Persona
   Set Persona = Nothing
End Function

'Private Sub CmdAceptar_Click()
'Dim lnCondicion As Integer, lcCondi As String
'Dim lbResultadoVisto As Boolean
'
'Dim loVistoElectronico As frmVistoElectronico
'Set loVistoElectronico = New frmVistoElectronico
'Dim lafirma As frmPersonaFirma
''MADM 20100825 - VISTO
'Dim lsTempUsuGrupo As String
'Dim lsUsuario As String
'Dim oAccesso As COMDPersona.UCOMAcceso
'Set oAccesso = New COMDPersona.UCOMAcceso
'
''END
''***RECO 20130701 ******
'Dim oPersona As UPersona_Cli
'Set oPersona = New UPersona_Cli
''***END RECO*******
'
'   Set Persona = New COMDPersona.UCOMPersona
'   If R Is Nothing Then
'        MsgBox "Seleccione un Cliente", vbInformation, "Aviso"
'        Exit Sub
'   Else
'        If R.RecordCount = 0 Then
'            MsgBox "Seleccione un Cliente", vbInformation, "Aviso"
'            Exit Sub
'        End If
'   End If
'   'RECO 20130701 ******
'   Call oPersona.RecuperaPersona(Trim(R!cPersCod))
'   '*****END RECO*******
'
'   '*** PEAC 20090731
'    'If Persona.ValidaEnListaNegativa(IIf(IsNull(R!cPersIDnroDNI), "", R!cPersIDnroDNI), IIf(IsNull(R!cPersIDnroRUC), "", R!cPersIDnroRUC), lnCondicion) Then
'    'ALPA 20091221**********************
'    If frmGrupoEcoEmpresa.nLogPerNegativa = 0 Then
'    '20090928 MADM **********************
'    If Persona.ValidaEnListaNegativaCondicion(IIf(IsNull(R!cPersIDnroDNI), "", R!cPersIDnroDNI), IIf(IsNull(R!cPersIDnroRUC), "", R!cPersIDnroRUC), lnCondicion, Me.txtNomPer.Text) Or oPersona.Nacionalidad <> "04028" Or oPersona.Residencia = 0 Then 'MODIFICADO:RECO-20130701
'    'If Persona.ValidaEnListaNegativaCondicion(IIf(IsNull(R!cPersIDnroDNI), "", R!cPersIDnroDNI), IIf(IsNull(R!cPersIDnroRUC), "", R!cPersIDnroRUC), lnCondicion, Me.txtNomPer.Text) Then
'        Select Case lnCondicion
'            Case 1
'                lcCondi = "NEGATIVO"
'            Case 2
'                lcCondi = "FRAUDULENTO"
'            Case 3
'                lcCondi = "PEPS"
'            Case 5
'                lcCondi = "LISTA OFAC"
'            Case 6
'                lcCondi = "LISTA ONU"
'            Case 7
'                lcCondi = "PEPS - NEGATIVO" 'MARG 13-05-2016
'        End Select
'        'If oPersona.Nacionalidad <> "" Then '*****RECO 20130701************'WIOR 20130909 COMENTÓ
'
'        '***MARG ERS046-2016 agrego 20161110***
'        Dim sMensaje As String
'        'END MARG***********************
'
'        If lnCondicion = 1 Or lnCondicion = 3 Or lnCondicion = 5 Or lnCondicion = 6 Or (oPersona.Nacionalidad <> "4028" And oPersona.Personeria = 1) Or (oPersona.Residencia = 0 And oPersona.Personeria = 1) Then  'madm 20100510---MODIFICADO:RECO-20130701'WIOR 20130909 SE INCLUYO oPersona.Personeria = 1 PARA LOS NO RESIDENTES Y EXTRANJEROS
'        'If lnCondicion = 1 Or lnCondicion = 3 Or lnCondicion = 5 Or lnCondicion = 6 Then 'madm 20100510
'
'            'WIOR 20120309----- Modificar el mensaje solo para Personas PEPS
'            If lnCondicion = 3 Then
'                '''MsgBox "Esta Persona es un " & lcCondi & ", necesitará un Visto electrónico", vbInformation, "Aviso" 'MARG ERS046-2016
'                sMensaje = "Esta Persona es un " & lcCondi 'MARG ERS046-2016
'            'WIOR 20130909***********************************************
'            ElseIf lnCondicion <> 3 And lnCondicion <> 0 Then
'                '''MsgBox "Este Cliente se encuentra en la lista Preventiva como " & lcCondi & ", necesitará un Visto electrónico", vbInformation, "Aviso" 'MARG ERS046-2016
'                sMensaje = "Este Cliente se encuentra en la lista Preventiva como " & lcCondi 'MARG ERS046-2016
'            'WIOR FIN ***************************************************
'            ElseIf (oPersona.Nacionalidad <> "04028" And oPersona.Personeria = 1) Then 'RECO****** SI ES EXTRANJERO'WIOR 20130909 AGREGO oPersona.Personeria = 1
'                '''MsgBox "Este Cliente es Extranjero, necesitará un Visto electrónico", vbInformation, "Aviso" 'MARG ERS046-2016
'                sMensaje = "Este Cliente es Extranjero" 'MARG ERS046-2016
'            ElseIf (oPersona.Residencia = 0 And oPersona.Personeria = 1) Then  'RECO****SI NO ES RESIDENTE'WIOR 20130909 AGREGO oPersona.Personeria = 1
'                '''MsgBox "Este Cliente es No Residente en el País, necesitará un Visto electrónico", vbInformation, "Aviso" 'MARG ERS046-2016
'                sMensaje = "Este Cliente es No Residente en el País" 'MARG ERS046-2016
'            'Else
'            '    MsgBox "Este Cliente esta en la Lista de Negativos como " & lcCondi & ", necesitará un Visto electrónico", vbInformation, "Aviso"
'            End If
'
'            'MADM 20100825 - MODIFICACION VISTO X OPERACION ------------------------------------------------------
'
'             Call oAccesso.CargaGruposUsuario(gsCodUser, gsDominio)
'
'             lsTempUsuGrupo = oAccesso.DameGrupoUsuario
'             lsUsuario = gsCodUser & "," & lsTempUsuGrupo
'
'             While lsTempUsuGrupo <> ""
'                lsTempUsuGrupo = oAccesso.DameGrupoUsuario
'                If lsTempUsuGrupo <> "" Then
'                    lsUsuario = lsUsuario & "," & lsTempUsuGrupo
'                End If
'              Wend
'
'             '***MARG ERS046-2016****************COMENTADO 20161108********************************************
''''             If Not oAccesso.ValidarVistoBuenoxGrupoUser(3, lsUsuario) Then
''''             'MADM 20101231
''''                lbResultadoVisto = loVistoElectronico.Inicio(3, "910000", R!cPersCod)
''''              Else
''''                lbResultadoVisto = loVistoElectronico.Inicio(4, "910000", R!cPersCod, gsCodUser)
''''             End If
''''
''''             If Not lbResultadoVisto Then
''''                 Exit Sub
''''             End If
'             '***END MARG********************************************************
'
'             'END MADM -------------------------------------------------------------------------------------------
'
'             '***MARG ERS046-2016***
'             If gsOpeCod = "190260" Or gsOpeCod = "190010" Or gsOpeCod = "190860" Or gsOpeCod = "190280" Then
'                Dim oPerVistoContinuidad As New COMDPersona.DCOMPersonas
'                Dim RsVisto As ADODB.Recordset
'                Dim nRespuesta As Integer
'                Dim nExito As Integer
'
'                nRespuesta = oPerVistoContinuidad.SolicitarVistoContinuidad(R!cPersCod, lnCondicion, gsOpeCod, gsCodUser, gsCodAge)
'                Set RsVisto = oPerVistoContinuidad.ObtenerPersRPLAFTVistoContinuidad(R!cPersCod, lnCondicion, gsOpeCod, gsCodUser, gsCodAge)
'
'                If nRespuesta = 1 Then '1:Tiene Solicitud(es) de Visto de Continuidad de proceso de crédito(s) pendiente(s) por admitir
'                    sMensaje = sMensaje & Chr(13) & " y Tiene Solicitud(es) de Visto de Continuidad del Proceso de Crédito Pendiente(s) por admitir"
'                    MsgBox sMensaje, vbInformation, "Aviso"
'                    Exit Sub
'                End If
'                If nRespuesta = 2 Then '2:Tiene vistos admitidos
'                    If gnCountVCAdmitido = 0 Then
'                        sMensaje = sMensaje & Chr(13) & vbNewLine & "Se Admite la Continuidad del Proceso de Crédito:" & Chr(13) & RsVisto!cComentario
'                        MsgBox sMensaje, vbInformation, "Aviso"
'                    End If
'                    gnCountVCAdmitido = gnCountVCAdmitido + 1
'                End If
'                If nRespuesta = 3 Then '3:Tiene vistos no admitidos
'                    sMensaje = sMensaje & Chr(13) & vbNewLine & "NO se admite la Continuidad del Proceso de Crédito:" & Chr(13) & RsVisto!cComentario
'                    MsgBox sMensaje, vbInformation, "Aviso"
'                    Exit Sub
'                End If
'                If nRespuesta = 4 Then '4:Se permite registrar la solicitud
'                    sMensaje = sMensaje & Chr(13) & "Se procederá a Solicitar Visto de Continuidad del Proceso del Crédito"
'                    MsgBox sMensaje, vbInformation, "Aviso"
'                    frmSolicitudContinuidadProcesoCredRPLAFT.Inicio R!cPersCod, lnCondicion, lcCondi, gsOpeCod, gsCodUser, gsCodAge
'                    Exit Sub
'                End If
'             Else
'                MsgBox sMensaje, vbInformation, "Aviso"
'             End If
'             '***END MARG***
'
'        ElseIf lnCondicion = 2 Then
'            MsgBox "Este Cliente esta en la Lista de Negativos como " & lcCondi & ", no se podrá continuar.", vbInformation, "Aviso"
'            Exit Sub
'        End If
'       'End If '*****RECO END************    End If'WIOR 20130909 COMENTÓ
'    End If
' End If
   '************************************
   'ALPA 20100930
   'Call Persona.CargaDatos(R!cPersCod, R!cPersNombre, Format(IIf(IsNull(R!dPersNacCreac), gdFecSis, R!dPersNacCreac), "dd/mm/yyyy"), IIf(IsNull(R!cPersDireccDomicilio), "", R!cPersDireccDomicilio), IIf(IsNull(R!cPersTelefono), "", R!cPersTelefono), R!nPersPersoneria, IIf(IsNull(R!cPersIDnroDNI), "", R!cPersIDnroDNI), IIf(IsNull(R!cPersIDnroRUC), "", R!cPersIDnroRUC), IIf(IsNull(R!cPersIDnro), "", R!cPersIDnro), IIf(IsNull(R!cPersNatSexo), "", R!cPersNatSexo), IIf(IsNull(R!cActiGiro1), "", R!cActiGiro1))
   'Call Persona.CargaDatos(R!cPersCod, R!cPersNombre, Format(IIf(IsNull(R!dPersNacCreac), gdFecSis, R!dPersNacCreac), "dd/mm/yyyy"), IIf(IsNull(R!cPersDireccDomicilio), "", R!cPersDireccDomicilio), IIf(IsNull(R!cPersTelefono), "", R!cPersTelefono), R!nPersPersoneria, IIf(IsNull(R!cPersIDnroDNI), "", R!cPersIDnroDNI), IIf(IsNull(R!cPersIDnroRUC), "", R!cPersIDnroRUC), IIf(IsNull(R!cPersIDnro), "", R!cPersIDnro), IIf(IsNull(R!cPersnatSexo), "", R!cPersnatSexo), IIf(IsNull(R!cActiGiro1), "", R!cActiGiro1), IIf(IsNull(R!nTipoId), "1", R!nTipoId), R!nEdad) 'APRI 20170622 ADD R!nEdad
   '***********************
   '************* firma madm 20090928
'    If R!nPersPersoneria = 1 Then
'    vcodper = R!cPersCod
'    Set lafirma = New frmPersonaFirma
'    Call frmPersonaFirma.Inicio(Trim(vcodper), Mid(vcodper, 4, 2), False)
'    End If
   '*********************
   
'   Set R = Nothing
'   Screen.MousePointer = 0
'   Unload Me
'End Sub

Private Sub cmdClose_Click()
    Set R = Nothing
    Set Persona = Nothing
    Unload Me
End Sub

'Private Sub cmdNewCli_Click()
'Dim sCodigo As String
'Dim sNombre As String
'Dim RNew As ADODB.Recordset
'Dim oConecta As DConecta
'Dim ssql As String
'
'    sCodigo = frmPersona.PersonaNueva
'    If sCodigo <> "" Then
'        sNombre = Mid(sCodigo, 14, Len(sCodigo) - 13)
'        sCodigo = Mid(sCodigo, 1, 13)
'        optOpcion(1).Value = True
'        txtCodPer.Text = sCodigo
'        Call txtCodPer_KeyPress(13)
'        CmdAceptar_Click
'        Unload Me
'    End If
'End Sub

Private Sub dbgrdPersona_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If dbgrdPersona.DataSource Is Nothing Then Exit Sub
    If txtNomPer.Visible Then
        txtNomPer.Text = dbgrdPersona.Columns(0)
    Else
        If txtCodPer.Visible Then
            txtCodPer.Text = dbgrdPersona.Columns(2)
        Else
            If txtDocPer.Visible Then
                LblDoc.Caption = Trim(IIf(IsNull(R!cTipo), "", R!cTipo))
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    dbgrdPersona.MarqueeStyle = dbgHighlightRow
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Private Sub optOpcion_Click(index As Integer)
    Select Case index
        Case 0 'Busqueda por Nombre
            txtNomPer.Text = ""
            txtNomPer.Visible = True
            txtNomPer.SetFocus
            txtCodPer.Text = ""
            txtCodPer.Visible = False
            txtDocPer.Text = ""
            txtDocPer.Visible = False
            LblDoc.Visible = False
        Case 1 'Busqueda por Codigo
            txtCodPer.Text = ""
            txtCodPer.Visible = True
            txtCodPer.SetFocus
            txtNomPer.Text = ""
            txtNomPer.Visible = False
            txtDocPer.Text = ""
            txtDocPer.Visible = False
            LblDoc.Visible = False
        Case 2 'Busqueda por Documento
            txtDocPer.Text = ""
            txtDocPer.Visible = True
            LblDoc.Visible = True
            'madm 20101012
            If txtDocPer.Visible Then
                txtDocPer.SetFocus
            End If
            'end madm
            txtCodPer.Text = ""
            txtCodPer.Visible = False
            txtNomPer.Text = ""
            txtNomPer.Visible = False
    End Select
End Sub

'Private Sub txtCodPer_GotFocus()
'    fEnfoque txtCodPer
'End Sub

Private Sub txtCodPer_KeyPress(KeyAscii As Integer)
Dim ClsPersona As COMDPersona.DCOMPersonas
    If KeyAscii = 13 Then
        If Len(Trim(txtCodPer.Text)) = 0 Then
            MsgBox "Falta Ingresar el Codigo de la Persona", vbInformation, "Aviso"
            Exit Sub
        End If
        Screen.MousePointer = 11
        Set ClsPersona = New COMDPersona.DCOMPersonas
        If bBuscarEmpleado Then
            Set R = ClsPersona.BuscaCliente(txtCodPer.Text, BusquedaEmpleadoCodigo)
        Else
            Set R = ClsPersona.BuscaCliente(txtCodPer.Text, BusquedaCodigo)
        End If
        Set dbgrdPersona.DataSource = R
        dbgrdPersona.Refresh
        Screen.MousePointer = 0
        If R.RecordCount = 0 Then
            MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
            txtCodPer.SetFocus
            cmdAceptar.Default = False
        Else
            dbgrdPersona.SetFocus
            cmdAceptar.Default = True
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
End Sub

'Private Sub txtDocPer_GotFocus()
'    fEnfoque txtDocPer
'End Sub

Private Sub txtDocPer_KeyPress(KeyAscii As Integer)
Dim ClsPersona As COMDPersona.DCOMPersonas
    If KeyAscii = 13 Then
        If Len(Trim(txtDocPer.Text)) = 0 Then
            MsgBox "Falta Ingresar el Documento de la Persona", vbInformation, "Aviso"
            Exit Sub
        End If
        Screen.MousePointer = 11
        Set ClsPersona = New COMDPersona.DCOMPersonas
        If bBuscarEmpleado Then
            Set R = ClsPersona.BuscaCliente(txtDocPer.Text, BusquedaEmpleadoDocumento)
        Else
            Set R = ClsPersona.BuscaCliente(txtDocPer.Text, BusquedaDocumento)
        End If
        Set dbgrdPersona.DataSource = R
        dbgrdPersona.Refresh
        Screen.MousePointer = 0
        If R.RecordCount = 0 Then
            MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
            txtDocPer.SetFocus
            cmdAceptar.Default = False
        Else
            dbgrdPersona.SetFocus
            cmdAceptar.Default = True
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
End Sub

'Private Sub txtNomPer_GotFocus()
'    fEnfoque txtNomPer
'End Sub
'
'Private Sub txtNomPer_KeyPress(KeyAscii As Integer)
'Dim ClsPersona As COMDPersona.DCOMPersonas
'   If KeyAscii = 13 Then
'      If Len(Trim(txtNomPer.Text)) = 0 Then
'        MsgBox "Falta Ingresar el Nombre de la Persona", vbInformation, "Aviso"
'        Exit Sub
'      End If
'      Screen.MousePointer = 11
'      Set ClsPersona = New COMDPersona.DCOMPersonas
'      If bBuscarEmpleado Then
'        Set R = ClsPersona.BuscaCliente(txtNomPer.Text, BusquedaEmpleadoNombre)
'      Else
'        Set R = ClsPersona.BuscaCliente(txtNomPer.Text)
'      End If
'      Set dbgrdPersona.DataSource = R
'      dbgrdPersona.Refresh
'      Screen.MousePointer = 0
'      If R.RecordCount = 0 Then
'        MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
'        txtNomPer.SetFocus
'        cmdAceptar.Default = False
'      Else
'        cmdAceptar.Default = True
'        txtNomPer.Text = Trim(R!cPersNombre)
'        dbgrdPersona.SetFocus
'      End If
'
'   Else
'        KeyAscii = Letras(KeyAscii)
'        cmdAceptar.Default = False
'   End If
'End Sub


'***RECO 20130701 ******
'Public Function VerificarEstadoPersona(cPersCod As String)
'    Dim ClsPersona As COMDPersona.DCOMPersonas
'    Set ClsPersona = New COMDPersona.DCOMPersonas
'    Set R = ClsPersona.BuscaCliente(cPersCod, BusquedaCodigo)
'
'    Call CmdAceptar_Click
'
'End Function
'***RECO END ******

