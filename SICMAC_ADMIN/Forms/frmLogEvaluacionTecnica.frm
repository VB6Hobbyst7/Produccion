VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogEvaluacionTecnica 
   Caption         =   "Procesos de Seleccion : Evaluacion Tecnica"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   Icon            =   "frmLogEvaluacionTecnica.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   10695
   Begin VB.TextBox txtestado 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8280
      TabIndex        =   22
      Top             =   2100
      Width           =   1455
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Editar"
      Height          =   390
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   6240
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   1560
      TabIndex        =   16
      Top             =   6240
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   3
      Left            =   2880
      TabIndex        =   15
      Top             =   6240
      Width           =   1305
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   390
      Left            =   9360
      TabIndex        =   12
      Top             =   6240
      Width           =   1305
   End
   Begin VB.TextBox txtDescripcionProveedor 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3000
      TabIndex        =   9
      Top             =   2100
      Width           =   4095
   End
   Begin VB.ComboBox cboperiodo 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame s 
      Caption         =   "Proceso de  Seleccion"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   10455
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
         Height          =   525
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   9015
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
         TabIndex        =   1
         Top             =   600
         Width           =   5895
      End
      Begin Sicmact.TxtBuscar txtSeleccion 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Proceso "
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
         Left            =   7080
         TabIndex        =   25
         Top             =   240
         Width           =   1410
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C00000&
         BorderStyle     =   4  'Dash-Dot
         FillColor       =   &H8000000D&
         Height          =   400
         Left            =   7080
         Top             =   480
         Width           =   3015
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
         Left            =   7200
         TabIndex        =   24
         Top             =   600
         Width           =   660
      End
      Begin VB.Label Label7 
         Caption         =   "Numero"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   315
      End
   End
   Begin Sicmact.TxtBuscar txtProveedor 
      Height          =   315
      Left            =   1200
      TabIndex        =   10
      Top             =   2100
      Width           =   1695
      _ExtentX        =   2990
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
      sTitulo         =   ""
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Evaluacion Tecnica"
      TabPicture(0)   =   "frmLogEvaluacionTecnica.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FlexEvaluacionTecnica"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtPuntMinimo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtPuntMaximo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtPuntMaximo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         TabIndex        =   21
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox txtPuntMinimo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         TabIndex        =   20
         Top             =   0
         Width           =   735
      End
      Begin Sicmact.FlexEdit FlexEvaluacionTecnica 
         Height          =   3135
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5530
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Codigo-Descripcion-Puntaje Max-Calificacion-Observacion-Ultima Actualizacion"
         EncabezadosAnchos=   "550-800-2500-1000-900-1950-2500"
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
         ColumnasAEditar =   "X-X-X-X-4-5-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   12648447
         BackColorControl=   12648447
         BackColorControl=   12648447
         EncabezadosAlineacion=   "C-L-L-R-R-C-C"
         FormatosEdit    =   "0-0-0-2-2-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   555
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Punt.Maximo"
         Height          =   195
         Left            =   5760
         TabIndex        =   19
         Top             =   0
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Punt. Minimo"
         Height          =   195
         Left            =   3840
         TabIndex        =   18
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Estado Prov"
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
      Left            =   7200
      TabIndex        =   23
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label lblproveedor 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
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
      TabIndex        =   11
      Top             =   2160
      Width           =   885
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
      TabIndex        =   8
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmLogEvaluacionTecnica"
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

Private Sub cboPeriodo_Click()
txtSeleccion.Text = ""
txttipo.Text = ""
txtdescripcion.Text = ""
flexEvaluacionTecnica.Clear
flexEvaluacionTecnica.FormaCabecera
flexEvaluacionTecnica.Rows = 2
txtProveedor.Text = ""
txtdescripcion.Text = ""
Me.txtSeleccion.rs = clsDGAdqui.LogSeleccionLista(cboperiodo.Text)
End Sub

Private Sub cmdReq_Click(Index As Integer)
Dim sActualiza As String
Dim nSuma As Double
Dim nestado As Integer
Dim nestadoProc As Integer
sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
Select Case Index
Case 1 'Editar
        If txtSeleccion.Text = "" Then
           MsgBox "Debe Seleccionar un Numero de Proceso Primero", vbInformation, "Seleccione un Numero de Proceso"
           txtSeleccion.SetFocus
           Exit Sub
        End If
        nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccion.Text)
        

        If nestadoProc = SelEstProcesoCancelado Then
            MsgBox "No se puede Modificar,El Procesos de Seleccion " + txtSeleccion.Text + " Tiene un estado de Cancelado", vbInformation, "Estado del proceso" + txtSeleccion.Text + " esta Cancelado "
            Exit Sub
        End If
        
        If nestadoProc = SelEstProcesoCerrado Then
            MsgBox "No se puede Modificar,El Procesos de Seleccion " + txtSeleccion.Text + " Esta Cerrado", vbInformation, "Estado del proceso" + txtSeleccion.Text + " esta Cerrado "
            Exit Sub
        End If
        
        If nestadoProc <> SelEstProcesoIniciado Then
            If nestadoProc = SelEstProcesoEvaluacionTec Then
            Else
                MsgBox "No se puede Modificar,El Procesos de Seleccion " + txtSeleccion.Text + " el estado es Diferente a Iniciado o al de Evaluacion Tecnica", vbInformation, "Estado del proceso" + txtSeleccion.Text + " es diferente a Iniciado "
                Exit Sub
            End If
            
        End If
        
        
        
        
        If txtProveedor.Text = "" Then
           MsgBox "Debe Seleccionar un Proveedor", vbInformation, "Seleccione un Proveedor"
           txtProveedor.SetFocus
           Exit Sub
        End If
        
        If Trim(txtProveedor.Text) = "0" Then
           MsgBox "El Proceso No tiene Configurado  Proveedores", vbInformation, "El Proceso No tiene Asignados  Proveedores"
           txtProveedor.SetFocus
           Exit Sub
        End If
        
        cmdReq(1).Enabled = False  'Editar
        cmdReq(2).Enabled = True  'Cancelar
        cmdReq(3).Enabled = True  'Grabar
        flexEvaluacionTecnica.Enabled = True
        flexEvaluacionTecnica.EliminaFila (flexEvaluacionTecnica.Rows - 1)
        txtSeleccion.Enabled = False
        
        saccion = "E"
Case 2 'Cancelar
        saccion = "C"
        cmdReq(1).Enabled = True  'Editar
        cmdReq(2).Enabled = False  'Cancelar
        cmdReq(3).Enabled = False 'Grabar
        mostrar_evaluacion txtSeleccion.Text
        flexEvaluacionTecnica.Enabled = False
        txtSeleccion.Enabled = True
Case 3 'Grabar
        If txtSeleccion.Text = "" Then
           MsgBox "Debe Seleccionar un Numero de Proceso Primero", vbInformation, "Seleccione un Numero de Proceso"
           txtSeleccion.SetFocus
           Exit Sub
        End If
        If txtProveedor.Text = "" Or Trim(txtProveedor.Text) = "0" Then
           MsgBox "Debe Seleccionar un Proveedor", vbInformation, "Seleccione un Proveedor"
           txtProveedor.SetFocus
           Exit Sub
        End If
        If flexEvaluacionTecnica.Rows <= 2 And flexEvaluacionTecnica.TextMatrix(1, 1) = "" Then
            MsgBox "Debe Ingresar Los Criterios de Evaluacion Tecnica del Proceso  Nº " & txtSeleccion.Text, vbInformation, "Ingrese un Numero de Criterio de Proceso"
            flexEvaluacionTecnica.SetFocus
            Exit Sub
        End If
        For i = 0 To flexEvaluacionTecnica.Rows - 1
            If flexEvaluacionTecnica.TextMatrix(i, 1) = "" Then
                MsgBox "Falta Ingresar Un Numero de Criterio en el Item  Nº " & i, vbInformation, "Ingrese un Numero de Criterio de Proceso"
                flexEvaluacionTecnica.SetFocus
                Exit Sub
            End If
        Next
        For i = 1 To flexEvaluacionTecnica.Rows - 1
            
            nSuma = nSuma + flexEvaluacionTecnica.TextMatrix(i, 4)
        Next
           If nSuma > Val(txtPuntMaximo.Text) Then
                MsgBox "La Suma de Los Puntajes de Los Criterios No debe ser Mayor que el Puntaje Maximo Establecido ", vbInformation, "Sumatoria de Puntajes excede al Puntaje Maximo Establecido"
                flexEvaluacionTecnica.SetFocus
                Exit Sub
           End If
           If nSuma < Val(txtPuntMinimo.Text) Then
               If MsgBox("La Suma de Los Puntajes es Menor que el Puntaje Minimo Establecido, se Descalificara al Proveedor  ", vbYesNo + vbInformation, "Sumatoria de Puntajes es Menor que el Puntaje Minimo Establecido ") = vbYes Then
               Else
                  Exit Sub
               End If
                
           End If
           
           
           Select Case saccion
            Case "E"
                    If flexEvaluacionTecnica.Rows = 2 And flexEvaluacionTecnica.TextMatrix(1, 1) = "" Then
                                'Elimina
                                clsDGAdqui.EliminaSeleccionCriteriosProceso txtSeleccionA.Text
                                Exit Sub
                    End If
                            ClsNAdqui.AgregaSeleccionEvaluacionTecnica txtSeleccion.Text, flexEvaluacionTecnica.GetRsNew, txtProveedor.Text, sActualiza
                            'Actualiza estado Proveedor
                            If flexEvaluacionTecnica.SumaRow(3) < Val(txtPuntMinimo.Text) Then
                               nestado = TpoLogSelEstProveedor.SelProDesCalificado
                               Else
                               nestado = TpoLogSelEstProveedor.SelProCalificado
                            End If
                            clsDGAdqui.ActualizaEstadoProveedor txtSeleccion.Text, txtProveedor.Text, nestado
                            'Actualiza Estado Proceso De Seleccion
                            clsDGAdqui.ActualizaEstadoProcesoSeleccion txtSeleccion.Text, TpoLogSelEstProceso.SelEstProcesoEvaluacionTec
                            nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccion.Text)
                            If nestadoProc = SelEstProcesoIniciado Then
                                    lblestado.Caption = "INICIADO"
                            ElseIf nestadoProc = SelEstProcesoEvaluacionTec Then
                                    lblestado.Caption = "EVALUACION TECNICA"
                            ElseIf nestadoProc = SelEstProcesoEvaluacionEco Then
                                    lblestado.Caption = "EVALUACION ECONOMICA"
                            ElseIf rs!nLogSelEstado = SelEstProcesoFinEvaluacion Then
                                    lblestado.Caption = "FIN DE EVALUACION"
                            ElseIf nestadoProc = SelEstProcesoCerrado Then
                                    lblestado.Caption = "CERRADO"
                            ElseIf nestadoProc = SelEstProcesoCancelado Then
                                    lblestado.Caption = "CANCELADO"
                            End If
         End Select
        cmdReq(1).Enabled = True  'Editar
        cmdReq(2).Enabled = False  'Cancelar
        cmdReq(3).Enabled = False 'Grabar
        flexEvaluacionTecnica.Enabled = False
        mostrar_evaluacion txtSeleccion.Text
        txtSeleccion.Enabled = True
        saccion = "G"
End Select
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Height = 7155
Me.Width = 10815
Dim sAno As String
Set rs = New ADODB.Recordset
Set clsDGnral = New DLogGeneral
Set clsDGAdqui = New DLogAdquisi
Set ClsNAdqui = New NActualizaProcesoSelecLog
Set rs = clsDGnral.CargaPeriodo
Call CargaCombo(rs, cboperiodo)
sAno = Year(gdFecSis)
ubicar_ano sAno, cboperiodo
flexEvaluacionTecnica.BackColorBkg = -2147483643
flexEvaluacionTecnica.Enabled = False

End Sub


Private Sub txtProveedor_EmiteDatos()

If txtProveedor.Text = "" Then
    MsgBox "Seleccione un Codigo Proveedor antes", vbInformation, "Seleccione Codigo de Proveedor"
    Exit Sub
End If

txtDescripcionProveedor.Text = txtProveedor.psDescripcion
txtProveedor.Enabled = True
'if tine criterios configurados then
'Recuperar sus criterios Tecnicos y Puntajes Respectivos
'else
'recuperar la plantilla de criterios tecnicos configurados
mostrar_evaluacion txtSeleccion.Text
End Sub

Sub mostrar_evaluacion(nLogSelProceso As Long)
  If txtProveedor.Text = "" Or txtProveedor.Text = "0" Then
     MsgBox "Seleccione un Proveedor,Asegurese de que el proceso tenga Proveedores Configurados", vbInformation, "Seleccione un Proveedor"
     Exit Sub
  End If
  txtestado.Text = clsDGAdqui.CargaEstadoProveedor(txtSeleccion.Text, txtProveedor.Text)
  Set rs = clsDGAdqui.CargaLogSelCriteriosProceso(txtSeleccion.Text, 3, txtProveedor.Text)
  If rs.EOF = True Then
    mostrar_criterios_procesos txtSeleccion.Text
    Else
    Set flexEvaluacionTecnica.Recordset = rs
  End If
  flexEvaluacionTecnica.AdicionaFila
  flexEvaluacionTecnica.TextMatrix(flexEvaluacionTecnica.Rows - 1, 1) = "-------------"
  flexEvaluacionTecnica.TextMatrix(flexEvaluacionTecnica.Rows - 1, 2) = "---------------TOTAL----------------"
  flexEvaluacionTecnica.TextMatrix(flexEvaluacionTecnica.Rows - 1, 3) = "-------------"
  flexEvaluacionTecnica.TextMatrix(flexEvaluacionTecnica.Rows - 1, 4) = Format(flexEvaluacionTecnica.SumaRow(4), "#####.00")
  flexEvaluacionTecnica.TextMatrix(flexEvaluacionTecnica.Rows - 1, 5) = "--------------------------------------------------"
  flexEvaluacionTecnica.TextMatrix(flexEvaluacionTecnica.Rows - 1, 6) = "--------------------------------------------------"
  

End Sub
Private Sub txtProveedor_GotFocus()

If txtSeleccion.Text = "" Then
    txtProveedor.Text = ""
    txtDescripcionProveedor.Text = ""
    Exit Sub
End If
Me.txtProveedor.rs = clsDGAdqui.LogSeleccionListaProveedores(txtSeleccion.Text)
If txtProveedor.rs.RecordCount = 0 Then
    MsgBox "El proceso de Seleccion no tiene Proveedores ", vbInformation, "No Tiene Criterios de Provedore"
    txtProveedor.Enabled = True
        txtProveedor.Text = ""
        txtDescripcionProveedor.Text = ""
    Exit Sub
End If


End Sub

Private Sub txtSeleccion_EmiteDatos()
Dim nNumCriterios As Integer
If txtSeleccion.Text = "" Then Exit Sub
'mostrar_criterios_procesos txtSeleccionA.Text
'Mostrar  Los criterios tecnicos del proceso con la configuracion de puntajes
txtPuntMinimo.Text = ""
txtPuntMaximo.Text = ""
flexEvaluacionTecnica.Clear
flexEvaluacionTecnica.Rows = 2
flexEvaluacionTecnica.FormaCabecera
txtProveedor.Text = ""
txtDescripcionProveedor.Text = ""

mostrar_descripcion txtSeleccion.Text

If Right(txtdescripcion.Text, 1) = 2 Then
    MsgBox "Este es un tipo de proceso Directo, No Necesita Evaluacion Tecnica", vbInformation, "Este esun Tipo de Proceso Directo "
    Unload Me
    Else
    nNumCriterios = clsDGAdqui.ValidaSelNumCriteriosProceso(txtSeleccion.Text)
    If nNumCriterios = 0 Then
        MsgBox "Falta Asignar los Criterios Tecnicos a este Proceso  ", vbInformation, "Falta asignar los Criterios Tecnicos"
        Unload Me
        Exit Sub
    End If
End If



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
            txtdescripcion.Text = "COTIZACION Nº: " + rs!nLogSelNumeroCot + " - " + rs!cDescripcionProceso + " - TIPO PROCESO: " + rs!nLogSelDescProceso + Space(300) + Str(rs!nLogSelTipoProceso)
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
    Set rs = clsDGAdqui.CargaSeleccionPuntajes(nLogSelProceso)
    If rs.EOF = True Then
        txtPuntMinimo.Text = ""
        txtPuntMaximo.Text = ""
        
        Else
            txtPuntMinimo.Text = Format(rs!nPuntajeMinimo, "00.00")
            txtPuntMaximo.Text = Format(rs!nPuntajeMaximo, "00.00")
           
    End If
    
End Sub
Sub mostrar_criterios_procesos(nLogSelProceso As Long)
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs = clsDGAdqui.CargaLogSelCriteriosProceso(nLogSelProceso, 2)
If rs.EOF = True Then
    flexEvaluacionTecnica.Rows = 2
    flexEvaluacionTecnica.Clear
    flexEvaluacionTecnica.FormaCabecera
    Else
    Set flexEvaluacionTecnica.Recordset = rs
        
End If
End Sub

