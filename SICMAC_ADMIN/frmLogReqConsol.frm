VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogReqConsol 
   Caption         =   "Consolidacion de Requerimientos"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12915
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   12915
   Begin VB.TextBox txtconsolidado 
      Height          =   285
      Left            =   8280
      TabIndex        =   32
      Top             =   120
      Width           =   4455
   End
   Begin TabDlg.SSTab SSTabgrillas 
      Height          =   6015
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   10610
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Detalle"
      TabPicture(0)   =   "frmLogReqConsol.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmbconsolidado"
      Tab(0).Control(1)=   "cmdconsolidar(1)"
      Tab(0).Control(2)=   "cmbvista"
      Tab(0).Control(3)=   "cmdconsolidar(0)"
      Tab(0).Control(4)=   "cmdactualizar"
      Tab(0).Control(5)=   "MshListReq"
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(7)=   "lblconsol"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Consolidado"
      TabPicture(1)   =   "frmLogReqConsol.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MshListConsol"
      Tab(1).Control(1)=   "cmdexport"
      Tab(1).Control(2)=   "cmbvistaconsol"
      Tab(1).Control(3)=   "cmdvistaconsol"
      Tab(1).Control(4)=   "cmbmesfin"
      Tab(1).Control(5)=   "cmbmesini"
      Tab(1).Control(6)=   "cboMoneda"
      Tab(1).Control(7)=   "cmbtiparea"
      Tab(1).Control(8)=   "Txtarea"
      Tab(1).Control(9)=   "Label4"
      Tab(1).Control(10)=   "Label3"
      Tab(1).Control(11)=   "lblmes1"
      Tab(1).Control(12)=   "lblEtiqueta(8)"
      Tab(1).Control(13)=   "lblAreaDes"
      Tab(1).Control(14)=   "lblEtiqueta(0)"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Precios Standar Bienes Servicios"
      TabPicture(2)   =   "frmLogReqConsol.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fgeBSPrecios"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdreqprecio(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdreqprecio(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdreqprecio(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.ComboBox cmbconsolidado 
         Height          =   315
         Left            =   -68880
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdconsolidar 
         Caption         =   "Eliminar"
         Height          =   375
         Index           =   1
         Left            =   -63600
         TabIndex        =   27
         Top             =   480
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshListConsol 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   12
         Top             =   1680
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   7435
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   16777215
         FocusRect       =   2
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.CommandButton cmdexport 
         Caption         =   "Exportar"
         Height          =   375
         Left            =   -63600
         TabIndex        =   25
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox cmbvistaconsol 
         Height          =   315
         Left            =   -73920
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   360
         Width           =   4935
      End
      Begin VB.CommandButton cmdvistaconsol 
         Caption         =   "Ver"
         Height          =   375
         Left            =   -64920
         TabIndex        =   22
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox cmbmesfin 
         Height          =   315
         Left            =   -71400
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cmbmesini 
         Height          =   315
         Left            =   -73920
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   -65160
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox cmbtiparea 
         Height          =   315
         Left            =   -73920
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdreqprecio 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   7200
         TabIndex        =   11
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdreqprecio 
         Caption         =   "Grabar"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   10
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdreqprecio 
         Caption         =   "Editar"
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   9
         Top             =   5520
         Width           =   1335
      End
      Begin VB.ComboBox cmbvista 
         Height          =   315
         Left            =   -74280
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   4935
      End
      Begin VB.CommandButton cmdconsolidar 
         Caption         =   "Consolidar"
         Height          =   375
         Index           =   0
         Left            =   -64920
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdactualizar 
         Caption         =   "Ver"
         Height          =   375
         Left            =   -66240
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshListReq 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   5
         Top             =   960
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   8705
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   16777215
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin Sicmact.FlexEdit fgeBSPrecios 
         Height          =   4935
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   12540
         _ExtentX        =   22119
         _ExtentY        =   8705
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "Item-Código-Descripción-Unidad-Referencia-Precio Ref-Ultima Actualizacion"
         EncabezadosAnchos=   "450-1500-3500-700-1200-1000-2500"
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
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-1-0-0"
         EncabezadosAlineacion=   "R-L-L-L-L-R-L"
         FormatosEdit    =   "0-0-0-0-0-2-0"
         CantEntero      =   6
         CantDecimales   =   1
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         Enabled         =   0   'False
         TipoBusqueda    =   2
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
      End
      Begin Sicmact.TxtBuscar Txtarea 
         Height          =   300
         Left            =   -71400
         TabIndex        =   29
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
         EnabledText     =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Vista"
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
         Left            =   -74880
         TabIndex        =   34
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lblconsol 
         AutoSize        =   -1  'True
         Caption         =   "xx"
         Height          =   195
         Left            =   -62520
         TabIndex        =   30
         Top             =   600
         Width           =   150
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Vista"
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
         Left            =   -74640
         TabIndex        =   24
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes Fin:"
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
         Left            =   -72240
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblmes1 
         AutoSize        =   -1  'True
         Caption         =   "Mes Ini :"
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
         Left            =   -74760
         TabIndex        =   18
         Top             =   840
         Width           =   750
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
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
         Index           =   8
         Left            =   -66000
         TabIndex        =   17
         Top             =   1800
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblAreaDes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -70080
         TabIndex        =   15
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area :"
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
         Index           =   0
         Left            =   -74520
         TabIndex        =   14
         Top             =   1320
         Width           =   525
      End
   End
   Begin VB.ComboBox cmbtipconsol 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox cboPeriodo 
      Height          =   315
      ItemData        =   "frmLogReqConsol.frx":0054
      Left            =   840
      List            =   "frmLogReqConsol.frx":0056
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin Sicmact.TxtBuscar txtconsol 
      Height          =   300
      Left            =   7320
      TabIndex        =   31
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
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
      EnabledText     =   0   'False
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
      Left            =   5880
      TabIndex        =   33
      Top             =   120
      Width           =   1320
   End
   Begin VB.OLE OLE1 
      Height          =   255
      Left            =   12960
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
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
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   1230
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
      TabIndex        =   2
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmLogReqConsol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim psTpoReq As String
Dim clsDReq As DLogRequeri
Dim oArea As DActualizaDatosArea
Dim Progress As clsProgressBar
Dim rs As ADODB.Recordset
Dim codigoant As String
Dim clsDMov As DLogMov

'Pa exportar
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet



Private Sub cmdSalir_Click()
    Set clsDReq = Nothing
    Unload Me
End Sub




Private Sub cboPeriodo_Click()
txtconsol.Text = ""
txtconsolidado.Text = ""
MshListReq.Clear
MshListConsol.Clear
End Sub

Private Sub cmbconsolidado_Click()
 If Right(cmbconsolidado.Text, 1) = "1" Then
    cmdconsolidar(0).Enabled = False
    cmdconsolidar(1).Enabled = True
 ElseIf Right(cmbconsolidado.Text, 1) = "2" Then
    txtconsol.Text = ""
    txtconsolidado.Text = ""
    cmdconsolidar(0).Enabled = True
    cmdconsolidar(1).Enabled = False
 End If
 MshListReq.Clear
End Sub

Private Sub cmbtiparea_Click()
If cmbtiparea.ListIndex = 0 Then      'Todos
    lblAreaDes.Visible = False
    Txtarea.Visible = False
    Txtarea.Text = ""
ElseIf cmbtiparea.ListIndex = 1 Then  'Por Area
    lblAreaDes.Visible = True
    Txtarea.Visible = True
End If

End Sub
Private Sub cmbtipconsol_Click()
txtconsol.Text = ""
txtconsolidado.Text = ""
MshListReq.Clear
MshListConsol.Clear

End Sub


Private Sub cmdactualizar_Click()
Dim svisConsol As String
lblconsol.Caption = "xx"
svisConsol = Right(cmbconsolidado.Text, 1)
If svisConsol = "1" And txtconsol.Text = "" Then
    MsgBox "Debe Seleccionar el Numero de Consolidado ", vbInformation, "Seleccionar el Numero de Consolidado"
    MshListReq.Clear
    Exit Sub
End If
Set rs = clsDReq.CargaReqListaDetalle(Right(Trim(cmbtipconsol.Text), 1), True, "", cboperiodo.Text, ReqEstadoaprobado, False, svisConsol, IIf(txtconsol.Text = "", 0, txtconsol.Text))
'CargaReqListaDetalle
If rs.RecordCount > 0 Then
    Set MshListReq.Recordset = rs
    MshListReq.SetFocus
    Else
    MsgBox "No existen Registros Para los Parametros Ingresados ", vbInformation, "No existen Registros"
    MshListReq.Clear
End If
Set rs = Nothing
'carga precios referenciales
Set rs = clsDReq.CargaReqPrecios(cboperiodo.Text)
If rs.RecordCount > 0 Then
   Set fgeBSPrecios.Recordset = rs
       fgeBSPrecios.ColWidth(7) = 0
   Else
    MsgBox "No existen Registros Para los Parametros Ingresados ", vbInformation, "No existen Registros"
    fgeBSPrecios.Clear
End If
Set rs = Nothing
lblconsol.Caption = clsDReq.CargaReqControlConsolCodigo(cboperiodo.Text, Right(Trim(cmbtipconsol.Text), 1))

End Sub


Private Sub cmdconsolidar_Click(Index As Integer)
Dim result As Integer
Dim i As Long
Dim sActualiza As String
Dim nvalor As Integer
Dim nestado As Integer
Dim ncodigo As Integer
Dim bflag As Boolean

Dim nCant As Integer
Dim nCantNull As Integer

If cboperiodo.Text = "" Then Exit Sub
If cmbtipconsol.Text = "" Then Exit Sub
i = 1
sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
'los maximos
nestado = clsDReq.CargaReqControlConsolEstado(cboperiodo.Text, Right(Trim(cmbtipconsol.Text), 1))
ncodigo = clsDReq.CargaReqControlConsolCodigo(cboperiodo.Text, Right(Trim(cmbtipconsol.Text), 1))
Select Case Index
Case 0
                 
                  'Verifica si existen Requerimientos Nuevos
                  nCant = clsDReq.CargaReqSinConsolidar(cboperiodo.Text, Right(Trim(cmbtipconsol.Text), 1))
                  If nCant = 0 Then
                     MsgBox "No existen Requerimientos Con estado sin Consolidar para Su Consolidacion " & cboperiodo.Text & " y Requerimiento " & Left(cmbtipconsol.Text, 15), vbInformation, "No Existen Requerimientos Para Consolidar "
                     Exit Sub
                  End If
                 
                 
                 If Right(Trim(cmbtipconsol.Text), 1) = 1 Then
                    If nestado = 3 Then
                         MsgBox "No se puede  Consolidar ,el Periodo " & cboperiodo.Text & " y Requerimiento " & Left(cmbtipconsol.Text, 12) & " Esta Aprobado", vbInformation, "Consolidado se encuentra Aprobado"
                         Exit Sub
                    End If
                    If nestado = 1 Then
                         MsgBox "Existe un Consolidado Pendiente de Aprobacion, Apruebelo o Eliminelo " & cboperiodo.Text & " y Requerimiento " & Left(cmbtipconsol.Text, 15), vbInformation, "Existe Un Consolidado Pendiente de Aprobacion"
                         Exit Sub
                    End If
                    If nestado = 0 Then 'No tiene consolidado
                    End If
                End If
               If Right(Trim(cmbtipconsol.Text), 1) = 2 Then
                  'Verificar si existe un consolidado pendiente de Aprobacion
                  If nestado = 1 Then
                     MsgBox "Existe un Consolidado Pendiente de Aprobacion ,Periodo  " & cboperiodo.Text & " y Requerimiento " & Left(cmbtipconsol.Text, 15), vbInformation, "Existe Un Consolidado Pendiente de Aprobacion"
                     Exit Sub
                  End If
                  If nestado = 0 Then 'No tiene consolidado
                  End If
                  If nestado = 3 Then 'El Ultimo esta Aprobado
                  End If
                 
               End If
                  nCantNull = clsDReq.CargaValidaReqPrecios(cboperiodo.Text)
                  If nCantNull > 0 Then
                     MsgBox "Existen " & nCantNull & "  Requerimietos de Bienes que No tienen Precio Referenciale ", vbInformation, "Existen Precios Referenciales Con Valor Null"
                     Exit Sub
                  End If
                  If MsgBox("¿ Estás seguro de Consolidar los Requerimientos del Periodo " & cboperiodo.Text & "  de Tipo de Requerimiento  " & Left(cmbtipconsol.Text, 12) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                    'Validar si los precios referenciales estas completamente asignados
                    oPlaEvento_ShowProgress
                    result = clsDMov.ProcesaReqConsol(cboperiodo.Text, Right(Trim(cmbtipconsol.Text), 1), sActualiza)
                     For i = 1 To 2000 'MshListReq.Rows - 1
                         oPlaEvento_Progress i, 2000 'MshListReq.Rows - 1
                     Next
                     oPlaEvento_CloseProgress
                     SSTabgrillas.Tab = 1
                  End If
Case 1
               
               If txtconsol.Text = "" Then
                        MsgBox "Debe seleccionar Un numero de Consolidado ", vbInformation, "Seleccione Un Numero de Consolidado"
                        Exit Sub
               End If
               
               nestado = clsDReq.CargaReqControlConsolEstadopoCod(cboperiodo.Text, Right(Trim(cmbtipconsol.Text), 1), txtconsol.Text)
               
               If nestado = 0 Then 'No tiene consolidado
                        MsgBox "No Existe un Consolidado a Eliminar del Periodo " & cboperiodo.Text & " y Requerimiento " & Left(cmbtipconsol.Text, 12) & " ", vbInformation, "No Existe Consolidado a Eliminar"
                        Exit Sub
               End If
               If nestado = 2 Then
                        MsgBox "No Existe un Consolidado a Eliminar del Periodo " & cboperiodo.Text & " y Requerimiento " & Left(cmbtipconsol.Text, 12) & " Con estado Pendiente de Aprobacion ", vbInformation, "No Existe Consolidado a Eliminar"
                        Exit Sub
               End If
               If Right(Trim(cmbtipconsol.Text), 1) = 1 Then
                    If nestado = 3 Then
                         MsgBox "No se puede  Eliminar el Consolidado Nº " & txtconsol.Text & "  del Periodo " & cboperiodo.Text & " y Requerimiento " & Left(cmbtipconsol.Text, 12) & " Esta Aprobado", vbInformation, "Consolidado se encuentra Aprobado"
                         Exit Sub
                    End If
                    If nestado = 1 Then 'para aprobacion si se puede Eliminar
                    End If
                    
               End If
               
               If Right(Trim(cmbtipconsol.Text), 1) = 2 Then
                  'Verificar si existe un consolidado pendiente de Aprobacion
                  If nestado = 3 Then
                         MsgBox "No se puede  Eliminar el Consolidado Nº " & txtconsol.Text & "  del Periodo " & cboperiodo.Text & " y Requerimiento " & Left(cmbtipconsol.Text, 12) & " Esta Aprobado", vbInformation, "Consolidado se encuentra Aprobado"
                         Exit Sub
                  End If
                  If nestado = 1 Then 'Si puede Eliminar
                  End If
               End If
               'Preguntar
               If MsgBox("¿ Estás seguro de Eliminar el Consolidado Nº " & txtconsol.Text & " del Periodo " & cboperiodo.Text & "  de Tipo de Requerimiento  " & Left(cmbtipconsol.Text, 12) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                   oPlaEvento_ShowProgress
                   result = clsDMov.EliminaReqConsol(cboperiodo.Text, Right(Trim(cmbtipconsol.Text), 1), sActualiza, ncodigo)
                   For i = 1 To 2000
                        oPlaEvento_Progress i, 2000
                   Next
                   oPlaEvento_CloseProgress
                   SSTabgrillas.Tab = 1
                   
               End If
               MshListConsol.Clear
               txtconsol.Text = ""
               txtconsolidado.Text = ""
End Select

End Sub

Private Sub cmdexport_Click()
Dim svisConsol As String
svisConsol = Right(cmbvistaconsol.Text, 1)
exportar_rep cboperiodo.Text, cmbtipconsol.Text, txtconsol.Text, txtconsolidado.Text, cmbvistaconsol.Text, cmbmesini.Text, cmbmesfin.Text, lblAreaDes.Caption, svisConsol
mostrar_consolidado
End Sub
Private Sub cmdreqprecio_Click(Index As Integer)
Dim nBs As Integer
Dim sBSCod As String
Dim nRefPrecio As Currency
Dim nLogReqCod As String
Dim nperiodo As Integer
Dim sActualiza As String
nperiodo = cboperiodo.Text
sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
Select Case Index
Case 0 'Editar
        fgeBSPrecios.Enabled = True
        cmdreqprecio(0).Enabled = False 'editar
        cmdreqprecio(1).Enabled = True  'grabar
        cmdreqprecio(2).Enabled = True  'cancelar
Case 1 'Grabar
    For nBs = 1 To fgeBSPrecios.Rows - 1
           
        If fgeBSPrecios.TextMatrix(nBs, 7) = 1 Then
            sBSCod = fgeBSPrecios.TextMatrix(nBs, 1)
            nRefPrecio = CCur(IIf(fgeBSPrecios.TextMatrix(nBs, 5) = "", 0, fgeBSPrecios.TextMatrix(nBs, 5)))
            nLogReqCod = Trim(IIf(fgeBSPrecios.TextMatrix(nBs, 4) = "", "", fgeBSPrecios.TextMatrix(nBs, 4)))
            clsDMov.ActualizaReqListaPrecios sBSCod, nperiodo, nLogReqCod, nRefPrecio, sActualiza
            Else
            sBSCod = ""
            nRefPrecio = 0
            nLogReqCod = ""
        End If
        fgeBSPrecios.Enabled = False
        cmdreqprecio(0).Enabled = True  'editar
        cmdreqprecio(1).Enabled = False 'grabar
        cmdreqprecio(2).Enabled = False  'cancelar
        Next
        'carga precios referenciales
        Set rs = clsDReq.CargaReqPrecios(cboperiodo.Text)
        If rs.RecordCount > 0 Then
            Set fgeBSPrecios.Recordset = rs
        End If
        Set rs = Nothing
        
Case 2 'Cancelar
       fgeBSPrecios.Enabled = False
       cmdreqprecio(0).Enabled = True 'editar
       cmdreqprecio(1).Enabled = False 'grabar
       cmdreqprecio(2).Enabled = False  'cancelar
       'carga precios referenciales
       Set rs = clsDReq.CargaReqPrecios(cboperiodo.Text)
       If rs.RecordCount > 0 Then
           Set fgeBSPrecios.Recordset = rs
       End If
       Set rs = Nothing
       
       
End Select
End Sub


Private Sub cmdvistaconsol_Click()
mostrar_consolidado
End Sub

Private Sub fgeBSPrecios_Click()

If fgeBSPrecios.Row <= 0 Then Exit Sub
 codigoant = fgeBSPrecios.TextMatrix(fgeBSPrecios.Row, 4)
End Sub
Sub formato(Vista As String)
Select Case Vista
Case "r", "f"
        MshListConsol.ColWidth(0) = 1500
        MshListConsol.ColWidth(1) = 2500
        MshListConsol.ColWidth(2) = 2000
        MshListConsol.ColWidth(3) = 3200
        MshListConsol.ColWidth(4) = 1000
        MshListConsol.ColWidth(5) = 1000
        MshListConsol.ColWidth(6) = 1000
        
        MshListConsol.MergeCol(0) = True
        MshListConsol.MergeCol(1) = True
        MshListConsol.MergeCol(2) = False
        MshListConsol.MergeCol(3) = False
Case "g", "h"
        MshListConsol.ColWidth(0) = 2000
        MshListConsol.ColWidth(1) = 3200
        MshListConsol.ColWidth(2) = 1000
        MshListConsol.ColWidth(3) = 1000
        
        MshListConsol.MergeCol(0) = True
        MshListConsol.MergeCol(1) = False
        MshListConsol.MergeCol(2) = False
        MshListConsol.MergeCol(3) = False
Case "i", "m"
        MshListConsol.ColWidth(0) = 1500
        MshListConsol.ColWidth(1) = 2000
        MshListConsol.ColWidth(2) = 3200
        MshListConsol.ColWidth(3) = 1100
        MshListConsol.ColWidth(4) = 1100
        MshListConsol.ColWidth(5) = 1100
        MshListConsol.ColWidth(6) = 1100
        MshListConsol.ColWidth(7) = 1100
        MshListConsol.ColWidth(8) = 1100
        MshListConsol.ColWidth(9) = 1100
        MshListConsol.ColWidth(10) = 1100
        MshListConsol.ColWidth(11) = 1100
        MshListConsol.ColWidth(12) = 1100
        MshListConsol.ColWidth(13) = 1100
        MshListConsol.ColWidth(14) = 1100
        MshListConsol.ColWidth(15) = 1100
        MshListConsol.ColWidth(16) = 1100
        MshListConsol.ColWidth(18) = 1100
        MshListConsol.ColWidth(19) = 1100
        MshListConsol.ColWidth(20) = 1100
        MshListConsol.ColWidth(21) = 1100
        MshListConsol.ColWidth(22) = 1100
        MshListConsol.ColWidth(23) = 1100
        MshListConsol.ColWidth(24) = 1100
        MshListConsol.ColWidth(25) = 1100
        MshListConsol.ColWidth(26) = 1100
        MshListConsol.ColWidth(27) = 1100
        MshListConsol.ColWidth(28) = 1100
        MshListConsol.MergeCol(0) = True
        MshListConsol.MergeCol(1) = True
        MshListConsol.MergeCol(2) = False
        MshListConsol.MergeCol(3) = False
Case "k", "n"
        MshListConsol.ColWidth(0) = 1500
        MshListConsol.ColWidth(1) = 3200
        MshListConsol.ColWidth(2) = 1100
        MshListConsol.ColWidth(3) = 1100
        MshListConsol.ColWidth(4) = 1100
        MshListConsol.ColWidth(5) = 1100
        MshListConsol.ColWidth(6) = 1100
        MshListConsol.ColWidth(7) = 1100
        MshListConsol.ColWidth(8) = 1100
        MshListConsol.ColWidth(9) = 1100
        MshListConsol.ColWidth(10) = 1100
        MshListConsol.ColWidth(11) = 1100
        MshListConsol.ColWidth(12) = 1100
        MshListConsol.ColWidth(13) = 1100
        MshListConsol.ColWidth(14) = 1100
        MshListConsol.ColWidth(15) = 1100
        MshListConsol.ColWidth(16) = 1100
        MshListConsol.ColWidth(18) = 1100
        MshListConsol.ColWidth(19) = 1100
        MshListConsol.ColWidth(20) = 1100
        MshListConsol.ColWidth(21) = 1100
        MshListConsol.ColWidth(22) = 1100
        MshListConsol.ColWidth(23) = 1100
        MshListConsol.ColWidth(24) = 1100
        MshListConsol.ColWidth(25) = 1100
        MshListConsol.ColWidth(26) = 1100
        MshListConsol.ColWidth(27) = 1100
        MshListConsol.MergeCol(0) = True
        MshListConsol.MergeCol(1) = False
        MshListConsol.MergeCol(2) = False
        MshListConsol.MergeCol(3) = False
Case "l", "o"
        MshListConsol.ColWidth(0) = 3200
        MshListConsol.ColWidth(1) = 1100
        MshListConsol.ColWidth(2) = 1100
        MshListConsol.ColWidth(3) = 1100
        MshListConsol.ColWidth(4) = 1100
        MshListConsol.ColWidth(5) = 1100
        MshListConsol.ColWidth(6) = 1100
        MshListConsol.ColWidth(7) = 1100
        MshListConsol.ColWidth(8) = 1100
        MshListConsol.ColWidth(9) = 1100
        MshListConsol.ColWidth(10) = 1100
        MshListConsol.ColWidth(11) = 1100
        MshListConsol.ColWidth(12) = 1100
        MshListConsol.ColWidth(13) = 1100
        MshListConsol.ColWidth(14) = 1100
        MshListConsol.ColWidth(15) = 1100
        MshListConsol.ColWidth(16) = 1100
        MshListConsol.ColWidth(18) = 1100
        MshListConsol.ColWidth(19) = 1100
        MshListConsol.ColWidth(20) = 1100
        MshListConsol.ColWidth(21) = 1100
        MshListConsol.ColWidth(22) = 1100
        MshListConsol.ColWidth(23) = 1100
        MshListConsol.ColWidth(24) = 1100
        MshListConsol.ColWidth(25) = 1100
        MshListConsol.ColWidth(26) = 1100
        MshListConsol.MergeCol(0) = False
        MshListConsol.MergeCol(1) = False
        MshListConsol.MergeCol(2) = False
        MshListConsol.MergeCol(3) = False
  Case "p"
        MshListConsol.ColWidth(0) = 3200
        MshListConsol.ColWidth(1) = 1100
        MshListConsol.ColWidth(2) = 1100
        MshListConsol.ColWidth(3) = 1100
        MshListConsol.ColWidth(4) = 1100
        MshListConsol.ColWidth(5) = 1100
        MshListConsol.ColWidth(6) = 1100
        MshListConsol.ColWidth(7) = 1100
        MshListConsol.ColWidth(8) = 1100
        MshListConsol.ColWidth(9) = 1100
        MshListConsol.ColWidth(10) = 1100
        MshListConsol.ColWidth(11) = 1100
        MshListConsol.ColWidth(12) = 1100
        MshListConsol.ColWidth(13) = 1100
        MshListConsol.ColWidth(14) = 1100
        MshListConsol.ColWidth(15) = 1100
        MshListConsol.ColWidth(16) = 1100
        MshListConsol.ColWidth(18) = 1100
        MshListConsol.ColWidth(19) = 1100
        MshListConsol.ColWidth(20) = 1100
        MshListConsol.ColWidth(21) = 1100
        MshListConsol.ColWidth(22) = 1100
        MshListConsol.ColWidth(23) = 1100
        MshListConsol.ColWidth(24) = 1100
        MshListConsol.ColWidth(25) = 1100
        MshListConsol.ColWidth(26) = 1100
        MshListConsol.MergeCol(0) = True
        MshListConsol.MergeCol(1) = True
        MshListConsol.MergeCol(2) = False
        MshListConsol.MergeCol(3) = False
   Case "q"
        MshListConsol.ColWidth(0) = 3200
        MshListConsol.ColWidth(1) = 1100
        MshListConsol.ColWidth(2) = 1100
        MshListConsol.ColWidth(3) = 1100
        MshListConsol.ColWidth(4) = 1100
        MshListConsol.ColWidth(5) = 1100
        MshListConsol.ColWidth(6) = 1100
        MshListConsol.ColWidth(7) = 1100
        MshListConsol.ColWidth(8) = 1100
        MshListConsol.ColWidth(9) = 1100
        MshListConsol.ColWidth(10) = 1100
        MshListConsol.ColWidth(11) = 1100
        MshListConsol.ColWidth(12) = 1100
        MshListConsol.ColWidth(13) = 1100
        MshListConsol.ColWidth(14) = 1100
        MshListConsol.ColWidth(15) = 1100
        MshListConsol.ColWidth(16) = 1100
        MshListConsol.ColWidth(18) = 1100
        MshListConsol.ColWidth(19) = 1100
        MshListConsol.ColWidth(20) = 1100
        MshListConsol.ColWidth(21) = 1100
        MshListConsol.ColWidth(22) = 1100
        MshListConsol.ColWidth(23) = 1100
        MshListConsol.ColWidth(24) = 1100
        MshListConsol.ColWidth(25) = 1100
        MshListConsol.ColWidth(26) = 1100
        MshListConsol.MergeCol(0) = True
        MshListConsol.MergeCol(1) = True
        MshListConsol.MergeCol(2) = True
        MshListConsol.MergeCol(3) = True
        
End Select

End Sub

Private Sub fgeBSPrecios_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol = 5 Then
    fgeBSPrecios.TextMatrix(pnRow, 7) = 1
    fgeBSPrecios.TextMatrix(pnRow, 4) = ""
End If
'If pnCol = 4 Then
'    fgeBSPrecios.TextMatrix(pnRow, 7) = 1
'End If



End Sub

Private Sub fgeBSPrecios_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim rsBS As ADODB.Recordset
Set rsBS = New ADODB.Recordset


Set rsBS = clsDReq.CargaPrecioReq(fgeBSPrecios.TextMatrix(pnRow, 1), psDataCod)
If rsBS.RecordCount > 0 Then
fgeBSPrecios.TextMatrix(pnRow, 5) = Format(rsBS!nLogReqDetPrecio, "######.#0")
If Trim(fgeBSPrecios.TextMatrix(pnRow, 4)) = codigoant Then
Else
fgeBSPrecios.TextMatrix(pnRow, 7) = 1
End If


Set rsBS = Nothing
End If

End Sub

Private Sub fgeBSPrecios_RowColChange()

Dim codbien As String
If fgeBSPrecios.Col = 4 Then
codbien = fgeBSPrecios.TextMatrix(fgeBSPrecios.Row, 1)
If fgeBSPrecios.TextMatrix(fgeBSPrecios.Row, 4) = "" Then

End If
fgeBSPrecios.rsTextBuscar = clsDReq.CargaReqPreciosProm(cboperiodo.Text, codbien)
codigoant = Trim(fgeBSPrecios.TextMatrix(fgeBSPrecios.Row, 4))

End If
End Sub


Private Sub Form_Load()
    
    Set clsDReq = New DLogRequeri
    Set oArea = New DActualizaDatosArea
    Dim clsDGnral As DLogGeneral
    Set clsDMov = New DLogMov
    Set oArea = New DActualizaDatosArea
    Set Progress = New clsProgressBar
    Set clsDGnral = New DLogGeneral
    Set rs = New ADODB.Recordset
    Set rs = clsDGnral.CargaPeriodo
    Call CargaCombo(rs, cboperiodo)
    cboperiodo.ListIndex = 0
    cmbtipconsol.AddItem "Regular                                           1"
    cmbtipconsol.AddItem "Extemporaneo                                      2"
    cmbtipconsol.ListIndex = 0
    'cmbvista.AddItem "Detalle Nivel Categoria Bien   " & Space(200) & "r"
    cmbvista.AddItem "Detalle Nivel Codigo Bien      " & Space(200) & "d"
    cmbvista.ListIndex = 0
    
    'cmbvistaconsol.AddItem "Detalle Consolidado Rango de Meses           " & Space(200) & "d"
    cmbvistaconsol.AddItem "Resumen (RM) Por Agencia,Area y Codigo de Bien    " & Space(200) & "r"
    cmbvistaconsol.AddItem "Resumen (RM) Por Agencia,Area y Categoria de Bien " & Space(200) & "f"
    cmbvistaconsol.AddItem "Resumen (RM) Por Categoria de Bien                " & Space(200) & "g"
    cmbvistaconsol.AddItem "Resumen (RM) Por Codigo de Bien                   " & Space(200) & "h"
    cmbvistaconsol.AddItem "-------------------------------------------------------------------"
    cmbvistaconsol.AddItem "Resumen Mensual Por Agencia,Area y Codigo de Bien " & Space(200) & "i"
    cmbvistaconsol.AddItem "Resumen Mensual Por Agencia y Codigo de Bien" & Space(200) & "k"
    cmbvistaconsol.AddItem "Resumen Mensual Por Codigo de Bien " & Space(200) & "l"
    cmbvistaconsol.AddItem "Resumen Mensual Por Agencia,Area y Categoria de Bien " & Space(200) & "m"
    cmbvistaconsol.AddItem "Resumen Mensual Por Agencia y Categoria de Bien" & Space(200) & "n"
    cmbvistaconsol.AddItem "Resumen Mensual Por Categoria de Bien " & Space(200) & "o"
    cmbvistaconsol.AddItem "Resumen Mensual Por Agencia,Area,Categoria de Bien y Codigo de Bien " & Space(200) & "q"
    cmbvistaconsol.AddItem "Resumen Mensual Por Categoria de Bien y Codigo de Bien " & Space(200) & "p"
    
    cmbvistaconsol.ListIndex = 0
    
    cmbconsolidado.AddItem "Consolidado" & Space(200) & "1"
    cmbconsolidado.AddItem "Sin Consolidar" & Space(200) & "2"
    cmbconsolidado.ListIndex = 0
    
    
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
   'lblAreaDes.Caption = Usuario.AreaNom
   'vmTramite_MenuItemClick 1, 1
   fgeBSPrecios.BackColor = &HFFFFFF
   MshListReq.Cols = 10
   MshListReq.ColWidth(0) = 0
   MshListReq.ColWidth(1) = 1600
   MshListReq.ColWidth(2) = 2500
   MshListReq.ColWidth(3) = 700
   MshListReq.ColWidth(4) = 3450
   MshListReq.ColWidth(5) = 700
   MshListReq.ColWidth(6) = 700
   MshListReq.ColWidth(7) = 800
   MshListReq.ColWidth(8) = 800
   MshListReq.ColWidth(9) = 1000
   'Cabecera de la Grilla
   'Del Detalle
   MshListReq.TextMatrix(0, 1) = "Agencia"
   MshListReq.TextMatrix(0, 2) = "Area"
   MshListReq.TextMatrix(0, 3) = "Req.Nº"
   MshListReq.TextMatrix(0, 4) = "Descripcion"
   MshListReq.TextMatrix(0, 5) = "Precio"
   MshListReq.TextMatrix(0, 6) = "Cantidad"
   MshListReq.TextMatrix(0, 7) = "Subtotal"
   MshListReq.TextMatrix(0, 8) = "Estado"
   MshListReq.TextMatrix(0, 9) = "Consolidado"
   MshListReq.MergeCells = flexMergeRestrictColumns
   MshListReq.MergeCol(1) = True
   MshListReq.MergeCol(2) = True
   MshListReq.MergeCol(3) = True
   'Del consolidado
   
   MshListConsol.Cols = 6
   MshListConsol.MergeCells = flexMergeRestrictColumns
   MshListConsol.ColWidth(0) = 1800
   MshListConsol.ColWidth(1) = 2300
   MshListConsol.ColWidth(2) = 2000
   MshListConsol.ColWidth(3) = 3500
   MshListConsol.ColWidth(4) = 900
   MshListConsol.ColWidth(5) = 900
   MshListConsol.ColWidth(6) = 900
   
   
   'Meses
   Set rs = clsDGnral.CargaConstante(gMeses)
   Call CargaCombo(rs, cmbmesini, , 1, 0)
   cmbmesini.ListIndex = 0
   Call CargaCombo(rs, cmbmesfin, , 1, 0)
   cmbmesfin.ListIndex = 11
   
   'Tipos de area
   cmbtiparea.AddItem "Todas las Areas"
   cmbtiparea.AddItem "Seleccione Area"
   cmbtiparea.ListIndex = 0
   
   'Tipos de Moneda
   Set rs = clsDGnral.CargaConstante(gMoneda, False)
   Call CargaCombo(rs, cboMoneda)
   Me.Txtarea.rs = oArea.GetAgenciasAreas
   'Me.txtconsol.rs = clsDReq.CargaReqControlConsol(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1))
   'Set rs = clsDReq.CargaReqControlConsol(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1))
   Set rs = Nothing
   SSTabgrillas.Tab = 0
End Sub

Private Sub MshListReq_DblClick()
Dim nRow As Integer
Dim ncol As Integer
nRow = MshListReq.Row
ncol = MshListReq.Col
Dim psTpoReq As String
psTpoReq = Right(cmbtipconsol.Text, 1)
If nRow < 0 Then Exit Sub

If ncol <> 3 Then Exit Sub
    If MshListReq.TextMatrix(nRow, 1) <> "" Then
                Call frmLogReqInicio.Inicio(psTpoReq, "3", MshListReq.TextMatrix(nRow, ncol))
     End If

End Sub

Private Sub oPlaEvento_ShowProgress()
    Progress.ShowForm Me
End Sub

Private Sub oPlaEvento_Progress(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Consolidando Requerimientos"
End Sub
Private Sub oPlaEvento_Progress2(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Eliminando Consolidado"
End Sub

Private Sub oPlaEvento_CloseProgress()
    Progress.CloseForm Me
End Sub

Private Sub TxtArea_EmiteDatos()
Me.lblAreaDes.Caption = Txtarea.psDescripcion

End Sub

Private Sub txtconsol_EmiteDatos()
Me.txtconsolidado.Text = txtconsol.psDescripcion
cmbconsolidado.ListIndex = 0
MshListReq.Clear
MshListConsol.Clear
End Sub

Private Sub txtconsol_GotFocus()
 Me.txtconsol.rs = clsDReq.CargaReqControlConsol(cboperiodo.Text, Right(Trim(cmbtipconsol.Text), 1))
 txtconsol.Enabled = True
End Sub
Sub exportar_rep(periodo As Integer, requerimiento As String, ConsolNum As String, desconsol As String, reporte As String, mesini As String, mesfin As String, area As String, svistaConsol As String)
Dim i As Long
Dim n As Long
Dim lsArchivoN As String
Dim lbLibroOpen As Boolean
Dim lsCadAnt As String
Dim lnIni As Integer
Dim J As Integer
On Error Resume Next
lsArchivoN = App.path & "\prueba.xls"
OLE1.Class = "ExcelWorkSheet"
lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
If Not lbLibroOpen Then
   Err.Clear
   'Set objExcel = CreateObject("Excel.Application")
   If Err.Number Then
      MsgBox "Can't open Excel."
   End If
   Exit Sub
End If
Set xlHoja1 = xlLibro.Worksheets(1)
ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
Dim band  As Boolean
Dim letra As String
lnIni = 0

xlHoja1.Cells(2, 1).value = "Reporte"
xlHoja1.Cells(2, 2).value = Left(cmbvista.Text, 50)

xlHoja1.Cells(3, 1).value = "Periodo"
xlHoja1.Cells(3, 2).value = periodo
xlHoja1.Cells(4, 1).value = "Tipo Requerimiento"
xlHoja1.Cells(4, 2).value = Left(requerimiento, 12)
xlHoja1.Cells(5, 1).value = "Consolidado "
xlHoja1.Cells(5, 2).value = "Nº:" & ConsolNum & "  - " & desconsol
xlHoja1.Cells(6, 1).value = "Mes Inicial"
xlHoja1.Cells(6, 2).value = Left(mesini, 12)
xlHoja1.Cells(6, 3).value = "Mes Final"
xlHoja1.Cells(6, 4).value = Left(mesfin, 12)
xlHoja1.Cells(7, 1).value = "Area"
xlHoja1.Cells(7, 2).value = IIf(area = "", "Todos", area)
'formatea Cabecera
xlHoja1.Range("A2:F7").Select
Selection.AutoFormat Format:=xlRangeAutoFormatClassic2, Number:=True, Font:=True, Alignment:=True, Border:=True, Pattern:=True, Width:=True
Range("B4").Select

'xlHoja1.Cells(10 + MshListConsol.Rows - 1, 1).value = "******"
'xlHoja1.Cells(10 + MshListConsol.Rows - 1, 2).value = "******"
MshListConsol.AddItem "*******************************************"
For n = 0 To MshListConsol.Cols - 1
    MshListConsol.Col = n
    lnIni = 0
    For i = 0 To MshListConsol.Rows - 1
            MshListConsol.Row = i
            If n = 0 Then
               If lsCadAnt = MshListConsol.Text Then
                    Else
                        xlHoja1.Cells(i + 9, n + 1).value = MshListConsol.Text
                        lsCadAnt = MshListConsol.Text
                        If i = 0 Then
                        Else
                            xlHoja1.Range("A" & lnIni + 9 & ":A" & i + 8 & "").Merge
                        End If
                        lnIni = i
               End If
            End If
              If n = 1 Then
               If lsCadAnt = MshListConsol.Text Then
                    Else
                        xlHoja1.Cells(i + 9, n + 1).value = MshListConsol.Text
                        lsCadAnt = MshListConsol.Text
                        
                        If i = 0 Then
                        Else
                            xlHoja1.Range("B" & lnIni + 9 & ":B" & i + 8 & "").Merge
                        End If
                        lnIni = i
                   End If
              End If
              If n = 2 And svistaConsol = "q" Then
               If lsCadAnt = MshListConsol.Text Then
                    Else
                        xlHoja1.Cells(i + 9, n + 1).value = MshListConsol.Text
                        lsCadAnt = MshListConsol.Text
                        If i = 0 Then
                        Else
                            xlHoja1.Range("C" & lnIni + 9 & ":C" & i + 8 & "").Merge
                        End If
                        lnIni = i
                   End If
              End If
            If svistaConsol = "q" Then
                If n <> 0 And n <> 1 And n <> 2 Then
                    xlHoja1.Cells(i + 9, n + 1).value = MshListConsol.Text
                End If
            Else
            If n <> 0 And n <> 1 Then
                xlHoja1.Cells(i + 9, n + 1).value = MshListConsol.Text
            End If
            End If
    Next
Next
xlHoja1.Range("A:A").NumberFormat = "@"
xlHoja1.Range("A:A").VerticalAlignment = xlCenter
xlHoja1.Range("B:B").NumberFormat = "@"
xlHoja1.Range("B:B").VerticalAlignment = xlCenter
xlHoja1.Range("C:C").NumberFormat = "@"
xlHoja1.Range("C:C").VerticalAlignment = xlCenter

OLE1.Class = "ExcelWorkSheet"
ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
OLE1.SourceDoc = lsArchivoN
OLE1.Verb = 1
OLE1.Action = 1
OLE1.DoVerb -1
End Sub

Sub formatoMeses()
Dim i As Integer
i = 4
xlHoja1.Cells(5, i).value = "Mes Enero"
i = i + 2
xlHoja1.Cells(5, i).value = "Mes Febrero"
i = i + 2
xlHoja1.Cells(5, i).value = "Mes Marzo"
i = i + 2
xlHoja1.Cells(5, i).value = "Mes Abril"
i = i + 2
xlHoja1.Cells(5, i).value = "Mes Mayo"
i = i + 2
xlHoja1.Cells(5, i).value = "Mes Junio"
i = i + 2
xlHoja1.Cells(5, i).value = "Mes Julio"
i = i + 2
xlHoja1.Cells(5, i).value = "Mes Agosto"
i = i + 2
xlHoja1.Cells(5, i).value = "Mes Setiembre"
i = i + 2
xlHoja1.Cells(5, i).value = "Mes Octubre"
i = i + 2
xlHoja1.Cells(5, i).value = "Mes Noviembre"
i = i + 2
xlHoja1.Cells(5, i).value = "Mes Diciembre"
i = i + 2
xlHoja1.Cells(5, i).value = "Total Anual"
Range("C5:D5").Select
 With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
Selection.Merge
End Sub

Sub mostrar_consolidado()
Dim barea As Boolean
Dim scodagencia As String
Dim scodarea As String
Dim psCategoria As String
If txtconsol.Text = "" Then
MsgBox "Debe Seleccionar Un Numero de Consolidado ", vbInformation, "Seleccione el Numero de Consolidado"
txtconsol.SetFocus
Exit Sub
End If
If Val(Trim(Right(Trim(cmbmesfin.Text), 2))) < Val(Trim(Right(Trim(cmbmesini.Text), 2))) Then
      MsgBox "El Mes Final no Debe ser Menor Que el mes Inicial", vbInformation, "Seleccione el Mes Final"
      cmbmesfin.SetFocus
      MshListConsol.Clear
      Exit Sub
End If
If Txtarea.Visible = False Then 'todos
            barea = True
            scodagencia = "01"
            scodarea = Trim(Txtarea.Text)
     Else
     If Txtarea.Text = "" Then
        MsgBox "Debe Seleccionar un Area antes  ", vbInformation, "Seleccione Un Area"
        Txtarea.SetFocus
        Exit Sub
     End If
     
     If Len(Trim(Txtarea.Text)) = 3 Then
            scodagencia = "01"
            scodarea = Trim(Txtarea.Text)
        ElseIf Len(Trim(Txtarea.Text)) > 3 Then
            scodagencia = Right(Trim(Txtarea.Text), 2)
            scodarea = Left(Trim(Txtarea.Text), 3)
        End If
     barea = False 'por area
 End If
If Right(cmbvistaconsol.Text, 1) = "d" Then
    
    If Trim(Right(Trim(cmbmesfin.Text), 2)) = Trim(Right(Trim(cmbmesini.Text), 2)) Then
        MshListConsol.MergeCol(3) = False
    End If
    Set rs = clsDReq.CargaReqConsol(cboperiodo.Text, Right(Trim(cmbtipconsol.Text), 1), barea, scodagencia, scodarea, Trim(Right(Trim(cmbmesini.Text), 2)), Trim(Right(Trim(cmbmesfin.Text), 2)), "d", txtconsol.Text)
        If rs.RecordCount > 0 Then
            Set MshListConsol.Recordset = rs
    Else
            MshListConsol.Clear
            MsgBox "No existen Registros para los Parametros Ingresados ", vbInformation, "No existen Registros"
    End If


ElseIf Right(cmbvistaconsol.Text, 1) = "r" Or Right(cmbvistaconsol.Text, 1) = "f" Or Right(cmbvistaconsol.Text, 1) = "g" Or Right(cmbvistaconsol.Text, 1) = "h" Then
    psCategoria = Right(cmbvistaconsol.Text, 1)
    formato psCategoria
    Set rs = clsDReq.CargaReqConsol(cboperiodo.Text, Right(Trim(cmbtipconsol.Text), 1), barea, scodagencia, scodarea, Trim(Right(Trim(cmbmesini.Text), 2)), Trim(Right(Trim(cmbmesfin.Text), 2)), psCategoria, txtconsol.Text)
        If rs.RecordCount > 0 Then
            Set MshListConsol.Recordset = rs
    Else
            MshListConsol.Clear
            MsgBox "No existen Registros para los Parametros Ingresados ", vbInformation, "No existen Registros"
    End If
ElseIf Right(cmbvistaconsol.Text, 1) = "i" Or Right(cmbvistaconsol.Text, 1) = "k" Or Right(cmbvistaconsol.Text, 1) = "l" Or Right(cmbvistaconsol.Text, 1) = "m" Or Right(cmbvistaconsol.Text, 1) = "n" Or Right(cmbvistaconsol.Text, 1) = "o" Or Right(cmbvistaconsol.Text, 1) = "p" Or Right(cmbvistaconsol.Text, 1) = "q" Then
    psCategoria = Right(cmbvistaconsol.Text, 1)
    formato psCategoria
    Set rs = clsDReq.CargaReqConsolMensual(cboperiodo.Text, Right(Trim(cmbtipconsol.Text), 1), barea, scodagencia, scodarea, Trim(Right(Trim(cmbmesini.Text), 2)), Trim(Right(Trim(cmbmesfin.Text), 2)), psCategoria, 1, txtconsol.Text)
        If rs.RecordCount > 0 Then
            Set MshListConsol.Recordset = rs
            Else
            MshListConsol.Clear
            MsgBox "No existen Registros para los Parametros Ingresados ", vbInformation, "No existen Registros"
        End If
End If

End Sub

