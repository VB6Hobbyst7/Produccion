VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogSelAsignacionCriterios 
   Caption         =   "Asignacion de Criterios Tecnicos"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   Icon            =   "frmLogSelAsignacionCriterios.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   10005
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   390
      Left            =   8640
      TabIndex        =   14
      Top             =   6240
      Width           =   1305
   End
   Begin VB.Frame s 
      Caption         =   "Proceso de  Seleccion"
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   9855
      Begin VB.TextBox txtdescripcion 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   8655
      End
      Begin VB.TextBox txttipo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   600
         Width           =   6015
      End
      Begin Sicmact.TxtBuscar txtSeleccionA 
         Height          =   315
         Left            =   1080
         TabIndex        =   15
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   2
         sTitulo         =   ""
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C00000&
         BorderStyle     =   4  'Dash-Dot
         FillColor       =   &H8000000D&
         Height          =   405
         Left            =   7200
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblestado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
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
         Height          =   195
         Left            =   7320
         TabIndex        =   18
         Top             =   600
         Width           =   660
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Proceso"
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
         Height          =   195
         Index           =   7
         Left            =   7200
         TabIndex        =   17
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label Label5 
         Caption         =   "Numero"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   315
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Seleccion de Criterios "
      TabPicture(0)   =   "frmLogSelAsignacionCriterios.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FlexCriterios"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdMant(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdMant(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdMant 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   390
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Top             =   3120
         Width           =   1305
      End
      Begin VB.CommandButton cmdMant 
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   3120
         Width           =   1305
      End
      Begin Sicmact.FlexEdit FlexCriterios 
         Height          =   2775
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4895
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Codigo-Descripcion-Puntaje Maximo-Ultima Actualizacion"
         EncabezadosAnchos=   "550-1200-3500-1200-2800"
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
         ColumnasAEditar =   "X-1-X-3-X"
         ListaControles  =   "0-1-0-0-0"
         BackColorControl=   16761024
         BackColorControl=   16761024
         BackColorControl=   16761024
         EncabezadosAlineacion=   "C-L-L-R-C"
         FormatosEdit    =   "0-0-0-2-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         Enabled         =   0   'False
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   555
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   3
      Left            =   3000
      TabIndex        =   4
      Top             =   6240
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Top             =   6240
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Editar"
      Height          =   390
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   6240
      Width           =   1305
   End
   Begin VB.ComboBox cmbperiodo 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Periodo"
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
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   660
   End
End
Attribute VB_Name = "frmLogSelAsignacionCriterios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim clsDGnral As DLogGeneral
Dim clsDGAdqui As DLogAdquisi
Dim ClsNAdqui As NActualizaProcesoSelecLog
Dim oCons As DConstantes
Dim saccion As String

Private Sub cmbperiodo_Click()
txtSeleccionA.Text = ""
txttipo.Text = ""
txtdescripcion.Text = ""
FlexCriterios.Clear
FlexCriterios.FormaCabecera
FlexCriterios.Rows = 2
Me.txtSeleccionA.rs = clsDGAdqui.LogSeleccionLista(cmbperiodo.Text)

End Sub

Private Sub cmdMant_Click(Index As Integer)
Select Case Index
Case 0
        FlexCriterios.AdicionaFila
        FlexCriterios.SetFocus
Case 1
        nBSRow = FlexCriterios.Row
        If MsgBox("¿ Estás seguro de eliminar " & FlexCriterios.TextMatrix(nBSRow, 2) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            FlexCriterios.EliminaFila nBSRow
            'clsDGAdqui.EliminaSeleccionCriteriosProceso txtSeleccionA.Text
        End If

End Select

End Sub

Private Sub cmdReq_Click(Index As Integer)
Dim sActualiza As String
sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
Select Case Index
Case 1 'Editar
        If txtSeleccionA.Text = "" Then
           MsgBox "Debe Seleccionar un Numero de Proceso Primero", vbInformation, "Seleccione un Numero de Proceso"
           txtSeleccionA.SetFocus
           Exit Sub
        End If
        
        
        If Right(txtdescripcion.Text, 1) = 2 Then
            MsgBox "Este es un Proceso de Tipo Directo o Abreviado y no se considera una Evaluacion Tecnica ", vbInformation, "No se ingresa criterios tecnicos para este tipo de proceso "
            Exit Sub
        End If
        
        
        
        
        
            
        nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccionA.Text)
        If nestadoProc = SelEstProcesoCerrado Then
            MsgBox "No se puede Modificar,el Proceso de Seleccion " + txtSeleccionA.Text + "  se encuentra Cerrado", vbInformation, "Estado del proceso " + txtSeleccionA.Text + " esta Cerrado"
            Exit Sub
        End If
        
        
        If nestadoProc <> SelEstProcesoIniciado Then
            MsgBox "No se puede Modificar,el Proceso de Seleccion " + txtSeleccionA.Text + " Tiene un estado diferente al de  INICIADO", vbInformation, "Estado del proceso " + txtSeleccionA.Text + " es diferente a INICIADO "
            Exit Sub
        End If
        saccion = "E"
        cmdReq(1).Enabled = False  'Editar
        cmdReq(2).Enabled = True  'Cancelar
        cmdReq(3).Enabled = True  'Grabar
        cmdMant(1).Enabled = True  'Eliminar
        cmdMant(0).Enabled = True  'Nuevo
        FlexCriterios.Enabled = True
        txtSeleccionA.Enabled = False
Case 2 'Cancelar
        saccion = "C"
        cmdReq(1).Enabled = True  'Editar
        cmdReq(2).Enabled = False  'Cancelar
        cmdReq(3).Enabled = False 'Grabar
        cmdMant(1).Enabled = False  'Eliminar
        cmdMant(0).Enabled = False 'Nuevo
        mostrar_criterios_procesos txtSeleccionA.Text
        FlexCriterios.Enabled = False
        txtSeleccionA.Enabled = True
Case 3 'Grabar
        If txtSeleccionA.Text = "" Then
           MsgBox "Debe Seleccionar un Numero de Proceso Primero", vbInformation, "Seleccione un Numero de Proceso"
           txtSeleccionA.SetFocus
           Exit Sub
        End If
        If FlexCriterios.Rows <= 2 And FlexCriterios.TextMatrix(1, 1) = "" Then
            MsgBox "Debe Ingresar Los Criterios de Evaluacion Tecnica del Proceso  Nº " & txtSeleccionA.Text, vbInformation, "Ingrese un Numero de Criterio de Proceso"
            FlexCriterios.SetFocus
            Exit Sub
        End If
        For i = 0 To FlexCriterios.Rows - 1
            If FlexCriterios.TextMatrix(i, 1) = "" Then
                MsgBox "Falta Ingresar Un Numero de Criterio en el Item  Nº " & i, vbInformation, "Ingrese un Numero de Criterio de Proceso"
                FlexCriterios.SetFocus
                Exit Sub
            End If
        Next
        
        Select Case saccion
            Case "E"
                    If FlexCriterios.Rows = 2 And FlexCriterios.TextMatrix(1, 1) = "" Then
                                'Elimina
                                clsDGAdqui.EliminaSeleccionCriteriosProceso txtSeleccionA.Text
                                Exit Sub
                    End If
                            ClsNAdqui.AgregaSeleccionCriteriosTecnicos txtSeleccionA.Text, FlexCriterios.GetRsNew, sActualiza
         End Select
        cmdReq(1).Enabled = True  'Editar
        cmdReq(2).Enabled = False  'Cancelar
        cmdReq(3).Enabled = False 'Grabar
        cmdMant(1).Enabled = False  'Eliminar
        cmdMant(0).Enabled = False 'Nuevo
        FlexCriterios.Enabled = False
        saccion = "G"
        mostrar_criterios_procesos txtSeleccionA.Text
        txtSeleccionA.Enabled = True
End Select
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub





Private Sub FlexCriterios_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim rsCT As ADODB.Recordset
    'Agregar unidad al Flex
    If Not pbEsDuplicado Then
        Set rsCT = New ADODB.Recordset
        Set rsCT = clsDGAdqui.CargalogselcriteriosTecnicos(FlexCriterios.TextMatrix(pnRow, pnCol))
        If rsCT.RecordCount > 0 Then
        FlexCriterios.TextMatrix(pnRow, 3) = rsCT!nPuntajeDefault
        End If
        Set rsBS = Nothing
    End If
End Sub

Private Sub Form_Load()
Me.Height = 7185
Me.Width = 10100
Dim sAno As String
Set rs = New ADODB.Recordset
Set clsDGnral = New DLogGeneral
Set clsDGAdqui = New DLogAdquisi
Set ClsNAdqui = New NActualizaProcesoSelecLog
Set rs = clsDGnral.CargaPeriodo
Call CargaCombo(rs, cmbperiodo)
sAno = Year(gdFecSis)
ubicar_ano sAno, cmbperiodo
Set rs = clsDGAdqui.CargaSelCriteriosTecnicos(2)
Me.FlexCriterios.rsTextBuscar = rs
FlexCriterios.BackColorBkg = -2147483643
End Sub

Sub mostrar_criterios_procesos(nLogSelProceso As Long)
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs = clsDGAdqui.CargaLogSelCriteriosProceso(nLogSelProceso, 1)
If rs.EOF = True Then
    FlexCriterios.Rows = 2
    FlexCriterios.Clear
    FlexCriterios.FormaCabecera
    Else
    Set FlexCriterios.Recordset = rs
End If
End Sub

Private Sub txtSeleccionA_EmiteDatos()
If txtSeleccionA.Text = "" Then Exit Sub
mostrar_criterios_procesos txtSeleccionA.Text
mostrar_descripcion txtSeleccionA.Text
End Sub
Sub mostrar_descripcion(nLogSelProceso As Long)
Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = clsDGAdqui.CargaLogSelDescripcionProceso(nLogSelProceso)
    If rs.EOF = True Then
        txttipo.Text = ""
        txtdescripcion.Text = ""
        lblestado.Caption = ""
        Else
            txttipo.Text = UCase(rs!cTipo)
            txtdescripcion.Text = "COTIZACION Nº: " + rs!nLogSelNumeroCot + " - " + rs!cDescripcionProceso + " - TIPO PROCESO: " + rs!nLogSelDescProceso + Space(200) + Str(rs!nLogSelTipoProceso)
         If rs!nLogSelEstado = SelEstProcesoIniciado Then
                lblestado.Caption = "INICIADO"
         ElseIf rs!nLogSelEstado = SelEstProcesoEvaluacionTec Then
                lblestado.Caption = "EVALUACION TECNICA"
         ElseIf rs!nLogSelEstado = SelEstProcesoEvaluacionEco Then
                lblestado.Caption = "EVALUACION ECONOMICA"
         ElseIf rs!nLogSelEstado = SelEstProcesoFinEvaluacion Then
                lblestado.Caption = "FIN DE EVALUACION"
         ElseIf rs!nLogSelEstado = SelEstProcesoCerrado Then
                lblestado.Caption = "CERRADO"
         ElseIf rs!nLogSelEstado = SelEstProcesoCancelado Then
                lblestado.Caption = "CANCELADO"
         End If
    End If
End Sub
