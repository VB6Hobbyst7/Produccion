VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmLogSelSeleccionBienes 
   Caption         =   "Configuracion de Bienes"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   Icon            =   "frmLogSelSeleccionBienes.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11820
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10320
      TabIndex        =   30
      Top             =   7320
      Width           =   1455
   End
   Begin VB.ComboBox cmbPlantillaCotizacion 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   7320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtPlantilla 
      Height          =   315
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   7320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtestado 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8640
      TabIndex        =   36
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdObtConfigBienes 
      Caption         =   "Obtener Listado"
      Height          =   375
      Left            =   3000
      TabIndex        =   35
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtDescripcionProveedor 
      Height          =   315
      Left            =   3240
      TabIndex        =   33
      Top             =   2280
      Width           =   3975
   End
   Begin Sicmact.TxtBuscar txtProveedor 
      Height          =   315
      Left            =   1440
      TabIndex        =   32
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.CommandButton cmdSolicitudCot 
      Caption         =   "Generar  Solicitud"
      Height          =   375
      Left            =   8760
      TabIndex        =   31
      Top             =   7320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame s 
      Caption         =   "Proceso de  Seleccion"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   11655
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
         Left            =   1320
         TabIndex        =   10
         Top             =   600
         Width           =   5775
      End
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
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   8775
      End
      Begin Sicmact.TxtBuscar txtSeleccion 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
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
         Width           =   2895
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
         TabIndex        =   39
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
         Left            =   7320
         TabIndex        =   38
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label7 
         Caption         =   "Numero"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdMant 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   5
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdMant 
      Caption         =   "Grabar"
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdMant 
      Caption         =   "Editar"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdObtBienes 
      Caption         =   "Obtener Bienes "
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox cboperiodo 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Configuracion  de Bienes"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fgeBienesConfig"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdEdit(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdEdit(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Referencia Consolidado"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   1695
         TabIndex        =   28
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Frame Frame1 
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
         Height          =   3855
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   9375
         Begin VB.TextBox txtmesfin 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox txtmesini 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtRequerimiento 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtperiodo 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtconsolidado 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   840
            Width           =   4935
         End
         Begin VB.TextBox txtCategoria 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Consolidado Nº"
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
            Left            =   120
            TabIndex        =   27
            Top             =   900
            Width           =   1320
         End
         Begin VB.Label lblperiodo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   120
            TabIndex        =   26
            Top             =   420
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Requerimiento"
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
            Left            =   3000
            TabIndex        =   25
            Top             =   420
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mes Final"
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
            Left            =   3480
            TabIndex        =   24
            Top             =   1860
            Width           =   825
         End
         Begin VB.Label lblmes1 
            AutoSize        =   -1  'True
            Caption         =   "Mes Inicial"
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
            Left            =   120
            TabIndex        =   23
            Top             =   1860
            Width           =   930
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Categoria de Bien "
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
            Left            =   120
            TabIndex        =   22
            Top             =   1380
            Width           =   1590
         End
      End
      Begin Sicmact.FlexEdit fgeBienesConfig 
         Height          =   3495
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   6165
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "Item-Código-Descripción Bien-Unidad-Valor Unidad-Descripcion Adicional-Cantidad-Precio Ref-Sub Total-Participa"
         EncabezadosAnchos=   "450-1200-3500-700-0-2500-1000-1000-1000-1200"
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
         ColumnasAEditar =   "X-1-X-X-X-5-6-7-X-9"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-0-0-0-0-0-0-4"
         EncabezadosAlineacion=   "R-L-L-L-L-L-R-C-R-C"
         FormatosEdit    =   "0-0-0-0-0-0-3-2-2-0"
         CantEntero      =   10
         CantDecimales   =   4
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         Enabled         =   0   'False
         lbFlexDuplicados=   0   'False
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
      End
   End
   Begin VB.Label lblcotiza 
      Height          =   255
      Left            =   5640
      TabIndex        =   42
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblEstProveedor 
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
      Left            =   7440
      TabIndex        =   37
      Top             =   2340
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
      TabIndex        =   34
      Top             =   2340
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
      TabIndex        =   1
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmLogSelSeleccionBienes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
          As String, ByVal lpDirectory As String, ByVal nShowCmd _
          As Long) As Long

Dim rs As New ADODB.Recordset
Dim clsDGnral As DLogGeneral
Dim clsDGAdqui As DLogAdquisi
Dim ClsNAdqui As NActualizaProcesoSelecLog
Dim oCons As DConstantes
Public sAccionSelBienes As String
Dim bpuntaje As Boolean
Dim clsDBS As DLogBieSer
Dim saccion As String
Dim psTposel As String

Private Sub cboPeriodo_Click()
txtSeleccion.Text = ""
txtTipo.Text = ""
txtDescripcion.Text = ""
Me.txtSeleccion.rs = clsDGAdqui.LogSeleccionLista(cboperiodo.Text)

End Sub



Private Sub cmbPlantillaCotizacion_Click()
txtPlantilla.Text = clsDGAdqui.CargalogSelDescPlantilla(2, Right(cmbPlantillaCotizacion.Text, 1))
End Sub

Private Sub cmdedit_Click(Index As Integer)
Dim nBSRow As Integer
Select Case Index
Case 0
        fgeBienesConfig.AdicionaFila
        fgeBienesConfig.SetFocus
Case 1
        nBSRow = fgeBienesConfig.Row
        If MsgBox("¿ Estás seguro de eliminar " & fgeBienesConfig.TextMatrix(nBSRow, 2) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            fgeBienesConfig.EliminaFila nBSRow
            'fgeBienesConfig.EliminaFila nBSRow
        End If
End Select
End Sub

Private Sub CmdLimpiar_Click()
fgeBienesConfig.Clear
fgeBienesConfig.FormaCabecera
End Sub

Private Sub cmdMant_Click(Index As Integer)
    Dim nBSRow As Integer
    Dim sactualiza As String
    sactualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
    'Botones de comandos del detalle de bienes/servicios
    Select Case Index
    Case 1 'Editar
            If txtSeleccion.Text = "" Then
               MsgBox "Antes debe Seleccionar un Numero de Porceso de Seleccion", vbInformation, "Seleccione un Proceso de Seleccion"
               Exit Sub
            End If
            
            If psTposel = "3" Then
                   nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccion.Text)
                    
                    If nestadoProc = SelEstProcesoCancelado Then
                       MsgBox "No se puede Modificar,El Procesos de Seleccion " + txtSeleccion.Text + " esta Anulado", vbInformation, "Estado del proceso " + txtSeleccion.Text + " esta Anulado"
                       Exit Sub
                    End If
                    
                    If nestadoProc = SelEstProcesoCerrado Then
                       MsgBox "No se puede Modificar,El Procesos de Seleccion " + txtSeleccion.Text + " Tiene un estado de Cerrado", vbInformation, "Estado del proceso " + txtSeleccion.Text + " esta Cerrado"
                       Exit Sub
                    End If
                    
                    If nestadoProc = SelEstProcesoCerrado Then
                        MsgBox "No se puede Modificar,El Procesos de Seleccion " + txtSeleccionA.Text + " Tiene un estado de Cerrado", vbInformation, "Estado del proceso" + txtSeleccionA.Text + " esta Cerrado"
                        Exit Sub
                    End If
                    If txtProveedor.Text = "" Then
                          MsgBox "Seleccione el Proveedor", vbInformation, "Seleccione el Proveedor"
                          Exit Sub
                    End If
                    If Left(txtestado.Text, 1) = "D" Then
                          MsgBox "El Proveedor no puede Participar Esta Descalificado", vbInformation, "El Proveedor est Descalificado"
                          fgeBienesConfig.EliminaFila fgeBienesConfig.Rows - 1
                          ClsNAdqui.AgregaSeleccionCotizacionProveedores txtSeleccion.Text, txtProveedor.Text, fgeBienesConfig.GetRsNew, sactualiza
                          Mostrar_config_Bienes_Cotiza txtSeleccion.Text, txtProveedor.Text
                          Exit Sub
                   End If
                   If fgeBienesConfig.Rows = 2 And fgeBienesConfig.TextMatrix(1, 1) = "" Then
                          MsgBox "No Puede Editar , pues este Proveedor no tiene su Cotizacion Ingresada ", vbInformation, "No se puede Editar"
                          Exit Sub
                   End If
                ElseIf psTposel = "1" Then
                   nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccion.Text)
                     If nestadoProc = SelEstProcesoCerrado Then
                       MsgBox "No se puede Modificar,El Procesos de Seleccion " + txtSeleccion.Text + " Tiene un estado de Cerrado", vbInformation, "Estado del proceso" + txtSeleccion.Text + "  esta Cerrado"
                       Exit Sub
                     End If

                    If nestadoProc <> SelEstProcesoIniciado Then
                       MsgBox "No se puede Modificar,El Procesos de Seleccion " + txtSeleccion.Text + " Tiene un estado diferente al de  Iniciado", vbInformation, "Estado del proceso" + txtSeleccion.Text + " es diferente a INICIADO "
                       Exit Sub
                    End If
            End If
            cmdMant(1).Enabled = False  'editar
            cmdMant(2).Enabled = True  'Grabar
            cmdMant(3).Enabled = True  'Cancelar
            fgeBienesConfig.Enabled = True
            cmdObtConfigBienes.Enabled = False
            txtSeleccion.Enabled = False
            If psTposel = "1" Then
               cmdEdit(0).Enabled = True
               cmdEdit(1).Enabled = True
                fgeBienesConfig.EliminaFila fgeBienesConfig.Rows - 1
            ElseIf psTposel = "3" Then
                   txtProveedor.Enabled = False
                   fgeBienesConfig.EliminaFila fgeBienesConfig.Rows - 1
            
            End If
            saccion = "E"
    Case 2 'Grabar
                If MsgBox("Esta seguro Que desea Grabar ? ", vbInformation + vbYesNo, "Si esta seguro pulse Si") = vbYes Then
                Else
                    Exit Sub
                End If
                If txtSeleccion.Text = "" Then Exit Sub
                'clsDGAdqui.CargaLogSelTipoBs
                If psTposel = "1" Then
                           If ValidaGrilla = False Then Exit Sub
                              fgeBienesConfig.Enabled = False
                           Select Case sAccionSelBienes
                              Case "N"
                                   If fgeBienesConfig.Rows = 2 And fgeBienesConfig.TextMatrix(1, 1) = "" Then 'la grilla no tiene nada
                                      'Elimina
                                      clsDGAdqui.EliminaSeleccionConfigBienes txtSeleccion.Text
                                      Else
                                      ClsNAdqui.AgregaSeleccionConfigBienes txtSeleccion.Text, fgeBienesConfig.GetRsNew, sactualiza
                                   End If
                                    ClsNAdqui.AgregaSeleccionReferencia txtSeleccion.Text, txtperiodo.Text, Right(txtRequerimiento.Text, 1), Left(txtconsolidado.Text, 2), Mid(txtconsolidado.Text, 5, 40), Left(txtmesini.Text, 10) + Space(10) + Right(txtmesini.Text, 2), Left(txtmesfin.Text, 10) + Space(10) + Right(txtmesfin.Text, 2), Trim(txtCategoria.Text), sactualiza
                                    sAccionSelBienes = "E"
                                    
                              Case "E"
                                    If fgeBienesConfig.Rows = 2 And fgeBienesConfig.TextMatrix(1, 1) = "" Then
                                    'Elimina
                                    clsDGAdqui.EliminaSeleccionConfigBienes txtSeleccion.Text
                                      Exit Sub
                                    End If
                                    If saccion = "E" Then
                                    
                                      ClsNAdqui.AgregaSeleccionConfigBienes txtSeleccion.Text, fgeBienesConfig.GetRsNew, sactualiza
                                    End If
                           End Select
                                     Mostrar_Config_Bienes txtSeleccion.Text
                 
                ElseIf psTposel = "3" Then
                            If saccion = "E" Then
                                   ClsNAdqui.AgregaSeleccionCotizacionProveedores txtSeleccion.Text, txtProveedor.Text, fgeBienesConfig.GetRsNew, sactualiza
                                   Mostrar_config_Bienes_Cotiza txtSeleccion.Text, txtProveedor.Text
                                   'Actualiza Estado Proceso De Seleccion
                                    clsDGAdqui.ActualizaEstadoProcesoSeleccion txtSeleccion.Text, TpoLogSelEstProceso.SelEstProcesoEvaluacionEco
                                    nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccion.Text)
                                    If nestadoProc = SelEstProcesoIniciado Then
                                        lblEstado.Caption = "INICIADO"
                                    ElseIf nestadoProc = SelEstProcesoEvaluacionTec Then
                                        lblEstado.Caption = "EVALUACION TECNICA"
                                    ElseIf nestadoProc = SelEstProcesoEvaluacionEco Then
                                        lblEstado.Caption = "EVALUACION ECONOMICA"
                                    ElseIf rs!nLogSelEstado = SelEstProcesoFinEvaluacion Then
                                        lblEstado.Caption = "FIN DE EVALUACION"
                                    ElseIf nestadoProc = SelEstProcesoCerrado Then
                                        lblEstado.Caption = "CERRADO"
                                    ElseIf nestadoProc = SelEstProcesoCancelado Then
                                        lblEstado.Caption = "CANCELADO"
                                    End If
                            End If
                End If
                fgeBienesConfig.Enabled = False
                cmdObtConfigBienes.Enabled = True
                cmdMant(1).Enabled = True  'editar
                cmdMant(2).Enabled = False  'Grabar
                cmdMant(3).Enabled = False 'Cancelar
                cmdEdit(0).Enabled = False
                cmdEdit(1).Enabled = False
                txtSeleccion.Enabled = True
                txtProveedor.Enabled = True
    Case 3 'Cancelar
         If psTposel = "1" Then
            Mostrar_Config_Bienes Val(txtSeleccion.Text)
         ElseIf psTposel = "3" Then
            Mostrar_config_Bienes_Cotiza txtSeleccion.Text, txtProveedor.Text
         End If
         saccion = "C"
         cmdMant(1).Enabled = True  'editar
         cmdMant(2).Enabled = False  'Grabar
         cmdMant(3).Enabled = False 'Cancelar
         cmdEdit(0).Enabled = False
         cmdEdit(1).Enabled = False
         fgeBienesConfig.Enabled = False
         cmdObtConfigBienes.Enabled = True
         txtSeleccion.Enabled = True
         txtProveedor.Enabled = True
         'recupera el original
 End Select
End Sub

Private Sub cmdObtBienes_Click()
Dim nCantidad As Integer

If txtSeleccion.Text = "" Then Exit Sub
    nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccion.Text)
    If nestadoProc <> SelEstProcesoIniciado Then
       MsgBox "No se puede Modificar,El Procesos de Seleccion " + txtSeleccion.Text + " Tiene un estado diferente al de  INICIADO", vbInformation, "Estado del proceso" + txtSeleccion.Text + " es diferente a INICIADO "
       Exit Sub
    End If
'Validar
nCantidad = clsDGAdqui.ValidaLogSelDetalle(txtSeleccion.Text)
If nCantidad = 0 Then
   If MsgBox("El Proceso de Seleccion Nº " & txtSeleccion.Text & " No Tiene Items Configurados ", vbInformation + vbYesNo, "Desea Continuar?") = vbYes Then
    frmLogSelConfigBienes.Show
    Exit Sub
   End If
ElseIf nCantidad > 0 Then
   If MsgBox("El Proceso de Seleccion Nº " & txtSeleccion.Text & " Ya  tiene Items Configuradas ", vbInformation + vbYesNo, "Desea Continuar?") = vbYes Then
        frmLogSelConfigBienes.Show
    Exit Sub
   End If
End If
End Sub

Private Sub cmdObtConfigBienes_Click()
Dim nNumProvedores As Integer
Dim nNumProvedoresEvalTec As Integer

If txtSeleccion.Text = "" Then Exit Sub

If psTposel = "3" Then
                   nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccion.Text)
                    If nestadoProc = SelEstProcesoCerrado Or nestadoProc = SelEstProcesoCancelado Then
                       MsgBox "No se puede Modificar,El Procesos de Seleccion " + txtSeleccion.Text + " Tiene un estado diferente al de  Iniciado,Evaluacion Tecnica , Evaluacion Economica o Fin de Evaluacion", vbInformation, "Estado del proceso" + txtSeleccion.Text + " es diferente a INICIADO "
                       Exit Sub
                      End If
End If


nNumProvedores = clsDGAdqui.CuentaLogSelProveedor(txtSeleccion.Text)
nNumProvedoresEvalTec = clsDGAdqui.CuentaLogSelProveedorEvalTec(txtSeleccion.Text)
   'Todos Los proveedores tengan Evalaucion tecnica

If Right(txtDescripcion.Text, 1) = 2 Then
   Else
        If nNumProvedores - nNumProvedoresEvalTec = 0 Then
            'Ok
        Else
            MsgBox "Existen " + Str(nNumProvedores - nNumProvedoresEvalTec) + " Proveedores del proceso " + txtSeleccion.Text + "Que No Tienen Evaluacion Tecnica ", vbInformation, "Pendiente Evaluacion Tecnica de Proveedor"
            Exit Sub
        End If

End If




If txtProveedor.Text = "" Then
    MsgBox "Debe Seleccionar un Proveedor participante del Proceso de Seleccion Nº " + txtSeleccion.Text, vbInformation, "Seleccione Proveedor"
    txtProveedor.SetFocus
    Exit Sub
End If
    
    nCantidad = clsDGAdqui.ValidaLogSelDetalleCotizaProveedor(txtSeleccion.Text, txtProveedor.Text)
    If nCantidad = 0 Then
       If MsgBox("El Proceso de Seleccion Nº " & txtSeleccion.Text & " No tiene Cotizacion ingresada", vbInformation + vbYesNo, "Desea Continuar para Obtener el Listado de la Solicitud de Cotizacion ?") = vbYes Then
          frmLogSelListadoBienes.Show vbModal
          Exit Sub
       End If
        ElseIf nCantidad > 0 Then
           If MsgBox("El Proceso de Seleccion Nº " & txtSeleccion.Text & " Ya Tiene Su Cotizacion Ingresada ?  ", vbInformation + vbYesNo, "Desea Continuar? ") = vbYes Then
              frmLogSelListadoBienes.Show vbModal
            Exit Sub
           End If
        End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSolicitudCot_Click()
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range

If txtPlantilla.Text = "" Then
    MsgBox "Debe Seleccionar un Tipo de Plantilla", vbInformation, ""
    Exit Sub
End If

If txtProveedor.Text = "" Then
    MsgBox "Debe Seleccionar Un Proveedor", vbInformation, "Seleccione un Proveedor"
    Exit Sub
End If


If Left(txtestado.Text, 1) = "D" Then
   MsgBox "Este Proveedor esta Descalificado no se Puede Generar la Solicitud de Cotizacion", vbInformation, "Proveedor Descalificado No se Puede Generar Solicitud"
   Exit Sub
End If

Dim Row As Integer
Dim Col As Integer
Dim i As Integer
Dim N As Integer
Dim sTemp As String
Dim arr() As String
ReDim arr(fgeBienesConfig.Rows - 1, fgeBienesConfig.Cols - 1 - 2 - 1)
'Create an instance of Word
Set oWord = CreateObject("Word.Application")
'Show Word to the user
oWord.Visible = True
'Add a new, blank document
Set oDoc = oWord.Documents.Open(App.path & "\Solicitud.doc")
'Get the current document's range object
'Store FlexGrid items to a two dimensional array
For Row = 0 To fgeBienesConfig.Rows - 1
    N = 0
    For Col = 0 To fgeBienesConfig.Cols - 1 - 1
        If Col = 1 Or Col = 4 Then
        Else
        
        arr(i, N) = Trim(fgeBienesConfig.TextMatrix(Row, Col))
        N = N + 1
        End If
    Next
    i = i + 1
Next
'Store array items to a string
For i = LBound(arr, 1) To UBound(arr, 1)
    For N = LBound(arr, 2) To UBound(arr, 2)
        sTemp = sTemp & arr(i, N)
        If N = UBound(arr, 2) Then
           sTemp = sTemp & vbCrLf
        Else
           sTemp = sTemp & vbTab
        End If
    Next
Next
With oWord.Selection.Find
            .Text = "CampFecha"
            .Replacement.Text = Trim(ImpreFormat(Format(gdFecSis, "dddd, d mmmm yyyy"), 25))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
oWord.Selection.Find.Execute Replace:=wdReplaceAll

With oWord.Selection.Find
            .Text = "CampNumProceso"
            .Replacement.Text = txtSeleccion.Text
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
 oWord.Selection.Find.Execute Replace:=wdReplaceAll
 
 With oWord.Selection.Find
            .Text = "CampDescripcion" 'Mid(string, start[, length])
            .Replacement.Text = Mid(txtTipo.Text, InStr(1, txtTipo.Text, "-"))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
 oWord.Selection.Find.Execute Replace:=wdReplaceAll
 
 With oWord.Selection.Find
            .Text = "CampTipProceso" 'Mid(string, start[, length])
            .Replacement.Text = Left(txtTipo.Text, InStr(1, txtTipo.Text, "-") - 1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
 oWord.Selection.Find.Execute Replace:=wdReplaceAll
  
 With oWord.Selection.Find
            .Text = "CampProveedor"
            .Replacement.Text = txtDescripcionProveedor.Text
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
 oWord.Selection.Find.Execute Replace:=wdReplaceAll
 
 Set rs = clsDGAdqui.CargaDirRUCPersona(txtProveedor.Text)
    If rs.EOF = True Then
    Else
                With oWord.Selection.Find
                           .Text = "CampDireccion"
                           .Replacement.Text = IIf(IsNull(rs!cPersDireccDomicilio), "", rs!cPersDireccDomicilio)
                           .Forward = True
                           .Wrap = wdFindContinue
                           .Format = False
                         End With
                oWord.Selection.Find.Execute Replace:=wdReplaceAll
                
                With oWord.Selection.Find
                           .Text = "CampNroRUC"
                           .Replacement.Text = "Nº RUC :" + IIf(IsNull(rs!cPersIDnro), "", rs!cPersIDnro)
                           .Forward = True
                           .Wrap = wdFindContinue
                           .Format = False
                         End With
                oWord.Selection.Find.Execute Replace:=wdReplaceAll
    End If
 
                With oWord.Selection.Find
                    .Text = "CampNumCotiza"
                    .Replacement.Text = lblcotiza.Caption
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                End With
                oWord.Selection.Find.Execute Replace:=wdReplaceAll
                
                With oWord.Selection.Find
                    .Text = "FechaCotizacion"
                    .Replacement.Text = gdFecSis
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                End With
                oWord.Selection.Find.Execute Replace:=wdReplaceAll
                

                
                
 'get the current document's range object and move to end of document
    Set oRange = oDoc.Bookmarks("\EndOfDoc").Range
    oRange.Text = sTemp
  
'convert the text to a table and format the table
oRange.ConvertToTable vbTab, Format:=5   'wdTableFormatColorful2

Set oRange = Nothing
 
 
'wAppSource.ActiveDocument.Close
'wApp.Visible = True


End Sub

Private Sub fgeBienesConfig_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol = 6 Or pnCol = 7 Then
fgeBienesConfig.TextMatrix(pnRow, 8) = Format(IIf(fgeBienesConfig.TextMatrix(pnRow, 6) = "", 0, fgeBienesConfig.TextMatrix(pnRow, 6)) * IIf(fgeBienesConfig.TextMatrix(pnRow, 7) = "", 0, fgeBienesConfig.TextMatrix(pnRow, 7)), "######0.00")
End If
End Sub
Private Sub fgeBienesConfig_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim rsBS As ADODB.Recordset
    'Agregar unidad al Flex
    If Not pbEsDuplicado Then
        Set rsBS = New ADODB.Recordset
        Set rsBS = clsDBS.CargaBS(BsUnRegistro, psDataCod)
        If rsBS.RecordCount > 0 Then
        fgeBienesConfig.TextMatrix(pnRow, 3) = rsBS!cConsUnidad     'cBSUnidad
        fgeBienesConfig.TextMatrix(pnRow, 4) = Trim(Right(Trim(rsBS!cConsUnidad), 2))
        End If
        Set rsBS = Nothing
    End If
End Sub
Private Sub Form_Load()
Me.Width = 11880
Me.Height = 8250
Set clsDBS = New DLogBieSer
fgeBienesConfig.BackColorBkg = 16777215
Set rs = New ADODB.Recordset
Set clsDGnral = New DLogGeneral
Set clsDGAdqui = New DLogAdquisi
Set clsDBS = New DLogBieSer
Set rs = clsDGnral.CargaPeriodo
Set ClsNAdqui = New NActualizaProcesoSelecLog
Call CargaCombo(rs, cboperiodo)
ubicar_ano Year(gdFecSis), cboperiodo
fgeBienesConfig.rsTextBuscar = clsDBS.CargaBS(BsTodosArbol)
    cmdMant(1).Enabled = True  'editar
    cmdMant(2).Enabled = False  'Grabar
    cmdMant(3).Enabled = False 'Cancelar
    cmdEdit(0).Enabled = False
    cmdEdit(1).Enabled = False
If psTposel = "2" Then 'solicitud de cotizacion
        Set rs = clsDGAdqui.CargalogSelPlantilla(2)
        Call CargaCombo(rs, cmbPlantillaCotizacion)
        Me.Caption = "Solicitud de Cotizacion"
        cmdMant(1).Visible = False   'editar
        cmdMant(2).Visible = False   'Grabar
        cmdMant(3).Visible = False  'Cancelar
        cmdEdit(0).Visible = False
        cmdEdit(1).Visible = False
        cmdObtBienes.Visible = False
        cmdEdit(0).Visible = False
        cmdEdit(1).Visible = False
        lblproveedor.Visible = True
        txtProveedor.Visible = True
        cmbPlantillaCotizacion.Visible = True
        cmdSolicitudCot.Visible = True
        txtPlantilla.Visible = True
        
        txtDescripcionProveedor.Visible = True
        fgeBienesConfig.Enabled = True
        fgeBienesConfig.ListaControles = "0-1-0-0-0-0-0-0-0-0"
        fgeBienesConfig.EncabezadosAnchos = "450 - 1100 - 3400 - 700 - 0 - 2400 - 850 - 850 - 900 - 0 "
        
        cmdObtConfigBienes.Visible = False
        fgeBienesConfig.Enabled = False
    ElseIf psTposel = "1" Then 'Seleccion de bienes
        fgeBienesConfig.Enabled = True
        fgeBienesConfig.ListaControles = "0-1-0-0-0-0-0-0-0"
        fgeBienesConfig.EncabezadosAnchos = "450 - 1100 - 3400 - 700 - 0 - 2400 - 850 - 850 - 900 - 0"
        cmdSolicitudCot.Visible = False
        lblproveedor.Visible = False
        txtProveedor.Visible = False
        txtDescripcionProveedor.Visible = False
        cmdObtConfigBienes.Visible = False
        lblEstProveedor.Visible = False
        txtestado.Visible = False
    ElseIf psTposel = "3" Then
        Me.Caption = "Cotizacion de Proveedores"
        cmdSolicitudCot.Visible = False
        fgeBienesConfig.EncabezadosAnchos = "450 - 1100 - 3400 - 700 - 0 - 2400 - 850 - 850 - 900 - 700"
        lblproveedor.Visible = True
        txtProveedor.Visible = True
        txtDescripcionProveedor.Visible = True
        cmdObtBienes.Visible = False
        cmdEdit(0).Visible = False
        cmdEdit(1).Visible = False
        fgeBienesConfig.ColumnasAEditar = "X-X-X-X-X-X-X-7-X-9"
        fgeBienesConfig.EncabezadosNombres = "Item-Codigo-Descripcion Bien-Unidad-Valor Unidad-Descripcion Adicional-Cantidad-Precio Cot.-Sub Total-Participa"
        cmdObtConfigBienes.Visible = True
End If



End Sub
Sub ubicar_ano(codigo As String, combo As ComboBox)
Dim i As Integer
For i = 0 To combo.ListCount
If combo.List(i) = codigo Then
    combo.ListIndex = i
    Exit For
    End If
Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
sAccionSelBienes = ""
End Sub
Private Sub txtProveedor_EmiteDatos()
txtDescripcionProveedor.Text = txtProveedor.psDescripcion
txtestado.Text = ""
If txtProveedor.Text = "" Or txtProveedor.Text = "0" Then
          MsgBox "Seleccione el Proveedor", vbInformation, "Seleccione el Proveedor"
          Exit Sub
End If

'obtener el Numero de RUC del Proveedor

If psTposel = "3" Then
    'Validar si Para ese Provedor se Ingreso Su Evaluacion Tecnica
    If Right(txtDescripcion.Text, 1) = 2 Then
    Else
       If clsDGAdqui.ValidaEvalTecProvedor(txtSeleccion.Text, txtProveedor.Text) = 0 Then
       MsgBox "Este Proveedor No tiene Calificacion Tecnica Ingresada", vbInformation, "Ingrese Puntajes de Criterios Tecnicos para Este Proveedor"
       Exit Sub
        End If
    End If
    
    Mostrar_config_Bienes_Cotiza txtSeleccion.Text, txtProveedor.Text
ElseIf psTposel = "2" Then
    txtestado.Text = clsDGAdqui.CargaEstadoProveedor(txtSeleccion.Text, txtProveedor.Text)
End If
End Sub

Private Sub txtProveedor_GotFocus()
If txtSeleccion.Text = "" Then
    txtProveedor.Text = ""
    txtDescripcionProveedor.Text = ""
    Exit Sub
End If
Me.txtProveedor.rs = clsDGAdqui.LogSeleccionListaProveedores(txtSeleccion.Text)
End Sub

Private Sub txtSeleccion_EmiteDatos()
cmdObtBienes.Enabled = True
If txtSeleccion.Text = "" Then Exit Sub

    
   If psTposel <> "3" Then
       'Configuracion de Bienes para 1 y 2
       Mostrar_Config_Bienes txtSeleccion.Text
   ElseIf psTposel = "3" Then
       fgeBienesConfig.Clear
       fgeBienesConfig.FormaCabecera
       fgeBienesConfig.Rows = 2
       
       txtProveedor.Text = ""
       txtDescripcionProveedor.Text = ""
   End If
   mostrar_descripcion txtSeleccion.Text
   sAccionSelBienes = "E"
End Sub

Sub Mostrar_Config_Bienes(pnNumseleccion As Long)
'Mostrar la Referencia
    Dim nSuma As Currency
    If txtSeleccion.Text = "" Then Exit Sub
    Set rs = clsDGAdqui.CargaSelReferencia(pnNumseleccion)
    If Not rs.EOF = True Then
        txtperiodo.Text = rs!nLogSelPeriodo
        If rs!nLogSelTpoReq = ReqTipoRegular Then
            txtRequerimiento.Text = "Regular" & Space(10) & rs!nLogSelTpoReq
           ElseIf rs!nLogSelTpoReq = ReqTipoExtemporaneo Then
           txtRequerimiento.Text = "Extemporaneo" & Space(10) & rs!nLogSelTpoReq
        End If
        txtconsolidado.Text = Str(rs!nConsolidado) + " - " + rs!SDescripcionConsol
        txtmesini.Text = rs!nMesIni
        txtmesfin.Text = rs!nMesFin
        txtCategoria.Text = rs!sCategoriaBien
        'Mostrar el Detalle
    End If
    Set rs = clsDGAdqui.CargaSelDetalle(pnNumseleccion, 1)
    If Not rs.EOF = True Then
        Set fgeBienesConfig.Recordset = rs
            fgeBienesConfig.AdicionaFila
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 1) = " --------------------------- "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 2) = " --------------------- TOTAL ------------- "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 3) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 4) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 5) = " --------------- " 'fgeBienesConfig.SumaRow(5)
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 6) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 7) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 8) = Format(fgeBienesConfig.SumaRow(8), "########.00")
        Else
        fgeBienesConfig.Clear
        fgeBienesConfig.FormaCabecera
        fgeBienesConfig.Rows = 2
    End If
End Sub

Sub Mostrar_config_Bienes_Cotiza(pnNumseleccion As Long, psCodproveedor As String)
    Dim nSuma As Currency
    If txtSeleccion.Text = "" Then Exit Sub
    If txtProveedor.Text = "" Or txtProveedor.Text = "0" Then
         MsgBox "Seleccione un Proveedor,Asegurese de que el proceso tenga Proveedores Configurados", vbInformation, "Seleccione un Proveedor"
         Exit Sub
    End If
    txtestado.Text = clsDGAdqui.CargaEstadoProveedor(txtSeleccion.Text, txtProveedor.Text)
    Set rs = clsDGAdqui.CargaSelReferencia(pnNumseleccion)
    If Not rs.EOF = True Then
        txtperiodo.Text = rs!nLogSelPeriodo
        If rs!nLogSelTpoReq = ReqTipoRegular Then
            txtRequerimiento.Text = "Regular" & Space(10) & rs!nLogSelTpoReq
           ElseIf rs!nLogSelTpoReq = ReqTipoExtemporaneo Then
           txtRequerimiento.Text = "Extemporaneo" & Space(10) & rs!nLogSelTpoReq
        End If
        txtconsolidado.Text = Str(rs!nConsolidado) + " - " + rs!SDescripcionConsol
        txtmesini.Text = rs!nMesIni
        txtmesfin.Text = rs!nMesFin
        txtCategoria.Text = rs!sCategoriaBien
        'Mostrar el Detalle
    End If
    Set rs = clsDGAdqui.CargaSelDetalle(pnNumseleccion, 3, psCodproveedor)
    If Not rs.EOF = True Then
        Set fgeBienesConfig.Recordset = rs
            fgeBienesConfig.AdicionaFila
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 1) = " --------------------------- "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 2) = " --------------------- TOTAL ------------- "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 3) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 5) = " --------------- " 'fgeBienesConfig.SumaRow(5)
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 6) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 7) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 8) = Format(fgeBienesConfig.SumaRow(8), "########.00")
        Else
        fgeBienesConfig.Clear
        fgeBienesConfig.FormaCabecera
        fgeBienesConfig.Rows = 2
    End If
End Sub


Private Function ValidaGrilla() As Boolean
    Dim nBs As Integer, nBSMes As Integer, nCant As Integer
    'Validación de BienesServicios
    ValidaGrilla = True
    For nBs = 1 To fgeBienesConfig.Rows - 1
        If fgeBienesConfig.TextMatrix(nBs, 1) = "" Then
            MsgBox "Falta determinar el Bien/Servicio en el Item " & nBs, vbInformation, "Seleccione un Codigo de Bien"
            ValidaGrilla = False
            Exit Function
        End If
       
        If fgeBienesConfig.TextMatrix(nBs, 6) = "" Then
            MsgBox "Falta determinar La Cantidad en el Item " & nBs, vbInformation, "Ingrese la Cantidad en el Item"
            ValidaGrilla = False
            Exit Function
        End If
        If fgeBienesConfig.TextMatrix(nBs, 7) = "" Then
            MsgBox "Falta determinar el Precio Referencial en el Item " & nBs, vbInformation, "Seleccione un Codigo de Bien"
            ValidaGrilla = False
            Exit Function
        End If
    Next
End Function

Sub mostrar_descripcion(nLogSelProceso As Long)
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs = clsDGAdqui.CargaLogSelDescripcionProceso(nLogSelProceso)
If rs.EOF = True Then
    txtTipo.Text = ""
    txtDescripcion.Text = ""
    lblEstado.Caption = ""
    Else
    txtTipo.Text = UCase(rs!cTipo)
    txtDescripcion.Text = "COTIZACION Nº: " + rs!nLogSelNumeroCot + " - " + rs!cDescripcionProceso + " - TIPO PROCESO: " + rs!nLogSelDescProceso + Space(300) + Str(rs!nLogSelTipoProceso)
    lblcotiza.Caption = rs!nLogSelNumeroCot
    If rs!nLogSelEstado = SelEstProcesoIniciado Then
        lblEstado.Caption = "INICIADO"
    ElseIf rs!nLogSelEstado = SelEstProcesoEvaluacionTec Then
        lblEstado.Caption = "EVALUACION TECNICA"
    ElseIf rs!nLogSelEstado = SelEstProcesoEvaluacionEco Then
        lblEstado.Caption = "EVALUACION ECONOMICA"
    ElseIf rs!nLogSelEstado = SelEstProcesoFinEvaluacion Then
        lblEstado.Caption = "FIN DE EVALUACION"
    ElseIf rs!nLogSelEstado = SelEstProcesoCerrado Then
        lblEstado.Caption = "CERRADO"
    ElseIf rs!nLogSelEstado = SelEstProcesoCancelado Then
        lblEstado.Caption = "CANCELADO"
    End If
End If
End Sub

Public Sub Inicio(ByVal psTipoSel As String, ByVal psFormTpo As String, Optional ByVal psSeleccionNro As String = "")
psTposel = psTipoSel
psFrmTpo = psFormTpo
psReqNro = psSeleccionNro
Me.Show
End Sub

