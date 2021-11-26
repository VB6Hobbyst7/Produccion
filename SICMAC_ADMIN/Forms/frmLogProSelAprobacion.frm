VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogProSelAprobacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobacion"
   ClientHeight    =   4995
   ClientLeft      =   15
   ClientTop       =   2310
   ClientWidth     =   11025
   Icon            =   "frmLogProSelAprobacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Proceso de Seleccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   10815
      Begin VB.TextBox TxtDescripcion 
         Height          =   555
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   600
         Width           =   6675
      End
      Begin VB.TextBox TxtMonto 
         Height          =   315
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox TxtTipo 
         Height          =   315
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3435
      End
      Begin VB.CommandButton CmdConsultarProceso 
         Caption         =   "Proceso de Seleccion"
         Height          =   375
         Left            =   8160
         TabIndex        =   6
         Top             =   780
         Width           =   2415
      End
      Begin VB.TextBox TxtProSelNro 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox txtanio 
         Height          =   315
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8280
         TabIndex        =   16
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   15
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nro"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   840
         TabIndex        =   14
         Top             =   360
         Width           =   255
      End
      Begin VB.Label LblMoneda 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8880
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2520
         TabIndex        =   13
         Top             =   360
         Width           =   285
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   6420
      TabIndex        =   8
      Top             =   4560
      Width           =   1275
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton CmndAprobar 
      Caption         =   "Aprobar"
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame FrameConObs 
      BorderStyle     =   0  'None
      Height          =   3075
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   10875
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFObsCon 
         Height          =   2895
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   5106
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         Cols            =   6
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   -2147483647
         ForeColorSel    =   -2147483624
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         WordWrap        =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
   End
End
Attribute VB_Name = "frmLogProSelAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gnProSelNro  As Integer, gcBSGrupoCod As String
Dim nTipo As Integer

Public Sub TipoFuncion(pnTipo As Integer)
nTipo = pnTipo
Me.Show 1
End Sub

Private Sub CmdConsultarProceso_Click()
On Error GoTo msflex_clckErr
    frmLogProSelCnsProcesoSeleccion.Inicio 2
    With frmLogProSelCnsProcesoSeleccion
        If Not .gbBandera Then Exit Sub
        gnProSelNro = .gvnProSelNro
        gcBSGrupoCod = .gvcBSGrupoCod
        TxtProSelNro.Text = .gvnNro
        TxtTipo.Text = .gvcTipo
        TxtMonto.Text = FNumero(.gvnMonto)
        LblMoneda.Caption = .gvcMoneda
        TxtDescripcion.Text = .gvcDescripcion
    End With
    Select Case nTipo
        Case 1
            If VerificarEtapa(gnProSelNro, cnAbsolucionConsultas) Then
                CargaConsultas
                CmndAprobar.Visible = True
                cmdImprimir.Visible = True
            Else
                MsgBox "No se Especificado esta Etapa para este proceso", vbInformation, "Aviso"
                CmndAprobar.Visible = False
                cmdImprimir.Visible = False
                Exit Sub
            End If
        Case 2
            If VerificarEtapa(gnProSelNro, cnObservaciones) Then
                CargaObservaciones
                CmndAprobar.Visible = True
                cmdImprimir.Visible = True
            Else
                MsgBox "No se Especificado esta Etapa para este proceso", vbInformation, "Aviso"
                CmndAprobar.Visible = False
                cmdImprimir.Visible = False
                Exit Sub
            End If
    End Select
Exit Sub
msflex_clckErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub cmdImprimir_Click()
    If gnProSelNro = 0 Then Exit Sub
    Select Case nTipo
        Case 1
            ImpConsultasWord gnProSelNro
        Case 2
            ImpObservacionesWord gnProSelNro
    End Select
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Function ValidarRespuesta() As Boolean
    On Error GoTo ValidarRespuestaErr
    Dim i As Integer
    i = 1
    ValidarRespuesta = False
    Do While i < MSFObsCon.Rows
        If MSFObsCon.TextMatrix(i, 5) = "" Then
            ValidarRespuesta = True
            Exit Function
        End If
        i = i + 1
    Loop
    Exit Function
ValidarRespuestaErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Function


Private Sub CmndAprobar_Click()
On Error GoTo CmndAprobarErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, sUser As String, cMovNro As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
    If ValidarRespuesta Then
        MsgBox "Existen Consultas sin Responder...", vbInformation, "Aviso"
        Exit Sub
    End If
'    gsCodUser = "AAAR" ' BORRAR AQUI
    If Not ValidaComite(gsCodUser) Then
        MsgBox "El Usuario no Pertenece al Comite...", vbInformation
        Exit Sub
    End If
    Select Case nTipo
        Case 1
            Set Rs = oCon.CargaRecordSet("select cResolucion1,cResolucion2,cResolucion3 from LogProSelConsultas where nProSelNro=" & gnProSelNro)
            If Not Rs.EOF Then
                cMovNro = GetLogMovNro
                If Rs!cResolucion1 = "" Then
                    sSQL = "update LogProSelConsultas set cResolucion1='" & cMovNro & "' where nProSelNro=" & gnProSelNro
                ElseIf Rs!cResolucion2 = "" Then
                    sUser = Right(Rs!cResolucion1, 4)
                    If sUser <> Right(cMovNro, 4) Then
                        sSQL = "update LogProSelConsultas set cResolucion2='" & cMovNro & "' where nProSelNro=" & gnProSelNro
                    Else
                        MsgBox "No Puede Aprobar dos Veces", vbInformation
'                        Exit Sub
                    End If
                ElseIf Rs!cResolucion3 = "" Then
                    sUser = Right(Rs!cResolucion2, 4)
                    If sUser <> Right(cMovNro, 4) Then
                        sSQL = "update LogProSelConsultas set cResolucion3='" & cMovNro & "' where nProSelNro=" & gnProSelNro
                    Else
                        MsgBox "No Puede Aprobar dos Veces", vbInformation
'                        Exit Sub
                    End If
                End If
                CierraEtapa gnProSelNro, cnPresentacionConsultas
                CierraEtapa gnProSelNro, cnAbsolucionConsultas
                CierraEtapa gnProSelNro, cnAtencionEvaluacionConsultas
            End If
            If sSQL = "" Then
                MsgBox "Consultas ya Fueron Aprobadas", vbInformation
            Else
                oCon.Ejecutar sSQL
                MsgBox "Aprobacion Registrada Correctamente...", vbInformation
            End If
        Case 2
            Set Rs = oCon.CargaRecordSet("select cResolucion1,cResolucion2,cResolucion3 from LogProSelObsBases where nProSelNro=" & gnProSelNro)
            If Not Rs.EOF Then
                cMovNro = GetLogMovNro
                If Rs!cResolucion1 = "" Then
                    sSQL = "update LogProSelObsBases set cResolucion1='" & cMovNro & "' where cRespuesta<>'' and nProSelNro=" & gnProSelNro
                ElseIf Rs!cResolucion2 = "" Then
                    sUser = Right(Rs!cResolucion1, 4)
                    If sUser <> Right(cMovNro, 4) Then
                        sSQL = "update LogProSelObsBases set cResolucion2='" & cMovNro & "' where cRespuesta<>'' and nProSelNro=" & gnProSelNro
                    Else
                        MsgBox "No Puede Aprobar dos Veces", vbInformation
'                        Exit Sub
                    End If
                ElseIf Rs!cResolucion3 = "" Then
                    sUser = Right(Rs!cResolucion2, 4)
                    If sUser <> Right(cMovNro, 4) Then
                        sSQL = "update LogProSelObsBases set cResolucion3='" & cMovNro & "' where cRespuesta<>'' and nProSelNro=" & gnProSelNro
                    Else
                        MsgBox "No Puede Aprobar dos Veces", vbInformation
'                        Exit Sub
                    End If
                End If
                CierraEtapa gnProSelNro, cnObservaciones
                CierraEtapa gnProSelNro, cnEvaluacionPropuestasObservaciones
                CierraEtapa gnProSelNro, cnResolucionObservaciones
            End If
            If sSQL = "" Then
                MsgBox "Observaciones ya Fueron Aprobadas", vbInformation
            Else
                oCon.Ejecutar sSQL
                MsgBox "Aprobacion Registrada Correctamente...", vbInformation
            End If
    End Select
        oCon.CierraConexion
    End If
    CmndAprobar.Enabled = False
    Exit Sub
CmndAprobarErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub Form_Load()
    CentraForm Me
    gnProSelNro = 0
    gcBSGrupoCod = ""
    txtanio.Text = Year(gdFecSis)
    Select Case nTipo
        Case 1
            CargaConsultas
        Case 2
            CargaObservaciones
    End Select
End Sub

Sub CargaConsultas()
Dim oConn As New DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
On Error GoTo CargaConsultasErr
   
FormaMSFObsCon
If oConn.AbreConexion Then
          
   sSQL = "select x.nProSelNro, x.nTipo, x.cPersCod, p.cPersNombre,x.cConsulta, x.cRespuesta  " & _
          " from LogProSelConsultas x inner join Persona p on x.cPersCod = p.cPersCod " & _
          " Where x.nProSelNro = " & gnProSelNro & ""
          
   Set Rs = oConn.CargaRecordSet(sSQL)
   i = 0
   Do While Not Rs.EOF
        i = i + 1
        InsRow MSFObsCon, i
        MSFObsCon.RowHeight(i) = 560
        MSFObsCon.TextMatrix(i, 2) = "Consulta"
        MSFObsCon.TextMatrix(i, 3) = Rs!cPersNombre
        MSFObsCon.TextMatrix(i, 4) = Rs!cConsulta
        MSFObsCon.TextMatrix(i, 5) = IIf(IsNull(Rs!cRespuesta), "", Rs!cRespuesta)
        MSFObsCon.ScrollBars = flexScrollBarBoth
        Rs.MoveNext
    Loop
End If
Exit Sub
CargaConsultasErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Sub CargaObservaciones()
Dim oConn As New DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
On Error GoTo CargaProp
   
FormaMSFObsCon
If oConn.AbreConexion Then
          
   sSQL = "select x.nProSelNro, x.cPersCod, p.cPersNombre,x.cObservacion, x.cRespuesta  " & _
          " from LogProSelObsBases x inner join Persona p on x.cPersCod = p.cPersCod " & _
          " Where x.nProSelNro = " & gnProSelNro & ""
          
   Set Rs = oConn.CargaRecordSet(sSQL)
   i = 0
   Do While Not Rs.EOF
        i = i + 1
        InsRow MSFObsCon, i
        MSFObsCon.RowHeight(i) = 560
        MSFObsCon.TextMatrix(i, 2) = "Observacion"
        MSFObsCon.TextMatrix(i, 3) = Rs!cPersNombre
        MSFObsCon.TextMatrix(i, 4) = Rs!cObservacion
        MSFObsCon.TextMatrix(i, 5) = IIf(IsNull(Rs!cRespuesta), "", Rs!cRespuesta)
        MSFObsCon.ScrollBars = flexScrollBarBoth
        Rs.MoveNext
    Loop
End If
Exit Sub
CargaProp:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Sub FormaMSFObsCon()
MSFObsCon.Clear
MSFObsCon.Rows = 2
MSFObsCon.RowHeight(0) = 320
MSFObsCon.RowHeight(1) = 10
MSFObsCon.ColWidth(0) = 0
MSFObsCon.ColWidth(1) = 0
MSFObsCon.ColWidth(2) = 1000:    MSFObsCon.TextMatrix(0, 2) = "Tipo"
MSFObsCon.ColWidth(3) = 3000:    MSFObsCon.TextMatrix(0, 3) = "Persona"
MSFObsCon.ColWidth(4) = 5000:    MSFObsCon.TextMatrix(0, 4) = "Descripción"
MSFObsCon.ColWidth(5) = 5000:    MSFObsCon.TextMatrix(0, 5) = "Respuesta"
End Sub

Private Function ValidaComite(ByVal pcCod As String) As Boolean
On Error GoTo ValidaComiteErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    sSQL = "select u.cUser,p.cPersCod from LogProSelComite p " & _
            "inner join rrhh u on p.cPersCod=u.cPersCod " & _
            "where u.cUser='" & pcCod & "' and p.nProSelNro=" & gnProSelNro
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            ValidaComite = True
        Else
            ValidaComite = False
        End If
        oCon.CierraConexion
    End If
    Exit Function
ValidaComiteErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Function

Private Function VerificarEtapa(ByVal pnProSelNro As Integer, ByVal nEtapa As Integer) As Boolean
On Error GoTo VerificarEtapaErr
    Dim oCon As DConecta, sSQL As String, Rs As New ADODB.Recordset
    Set oCon = New DConecta
    sSQL = "select nNro=count(*)  From LogProSelEtapa e " & _
           "      inner join LogEtapa t on t.nEstado = 1 and " & _
           "    e.nEtapaCod = t.nEtapaCod  Where e.nProSelNro = " & pnProSelNro & " and e.nEtapaCod = " & nEtapa
    If oCon.AbreConexion Then
       Set Rs = oCon.CargaRecordSet(sSQL)
       If Not Rs.EOF Then
          If Rs!nNro > 0 Then
            VerificarEtapa = True
          Else
            VerificarEtapa = False
          End If
       End If
       oCon.CierraConexion
    End If
    Exit Function
VerificarEtapaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmLogProSelAprobacion = Nothing
End Sub

Private Sub txtanio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Val(TxtProSelNro.Text) >= 0 Then
            ConsultarProcesoNro Val(TxtProSelNro.Text), Val(txtanio.Text)
            Exit Sub
        Else
            TxtProSelNro.SetFocus
        End If
    End If
    KeyAscii = DigNumEnt(KeyAscii)
End Sub

Private Sub TxtProSelNro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Val(txtanio.Text) > 0 Then
            ConsultarProcesoNro Val(TxtProSelNro.Text), Val(txtanio.Text)
            Exit Sub
        Else
            txtanio.SetFocus
        End If
    End If
    KeyAscii = DigNumEnt(KeyAscii)
End Sub

Private Sub ConsultarProcesoNro(ByVal pnNro As Integer, ByVal pnAnio As Integer)
    On Error GoTo ConsultarProcesoNroErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    sSQL = "select t.cProSelTpoDescripcion, s.nProSelNro, s.nPlanAnualNro, s.nPlanAnualAnio, " & _
            "s.nPlanAnualMes, s.nProSelTpoCod, s.nProSelSubTpo, nNroProceso, c.cConsDescripcion, " & _
            "s.nObjetoCod , s.nMoneda, s.nProSelMonto, s.nProSelEstado, cSintesis, nModalidadCompra " & _
            "from LogProcesoSeleccion s " & _
            "inner join LogProSelTpo t on s.nProSelTpoCod = t.nProSelTpoCod " & _
            "left outer join constante c on s.nObjetoCod=c.nConsValor and c.nConsCod = 9048 " & _
            "where s.nProSelEstado > -1 and s.nNroProceso=" & pnNro & " and nPlanAnualAnio = " & pnAnio
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            gnProSelNro = Rs!nProselNro
'            gnBSGrupoCod = rs!cBSGrupoCod
            TxtTipo.Text = Rs!cProSelTpoDescripcion
            TxtMonto.Text = FNumero(Rs!nProSelMonto)
            LblMoneda.Caption = IIf(Rs!nMoneda = 1, "S/.", "$")
            TxtDescripcion.Text = Rs!cSintesis
        Else
            gnProSelNro = 0
            TxtTipo.Text = ""
            TxtMonto.Text = ""
            LblMoneda.Caption = ""
            TxtDescripcion.Text = ""
            MsgBox "Proceso no Existe...", vbInformation, "Aviso"
            Exit Sub
        End If
        oCon.CierraConexion
    End If
    Select Case nTipo
        Case 1
            If VerificarEtapa(gnProSelNro, cnAbsolucionConsultas) Then
                CargaConsultas
                CmndAprobar.Visible = True
                cmdImprimir.Visible = True
            Else
                MsgBox "No se Especificado esta Etapa para este proceso", vbInformation, "Aviso"
                CmndAprobar.Visible = False
                cmdImprimir.Visible = False
                Exit Sub
            End If
        Case 2
            If VerificarEtapa(gnProSelNro, cnObservaciones) Then
                CargaObservaciones
                CmndAprobar.Visible = True
                cmdImprimir.Visible = True
            Else
                MsgBox "No se Especificado esta Etapa para este proceso", vbInformation, "Aviso"
                CmndAprobar.Visible = False
                cmdImprimir.Visible = False
                Exit Sub
            End If
    End Select
    Exit Sub
ConsultarProcesoNroErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

