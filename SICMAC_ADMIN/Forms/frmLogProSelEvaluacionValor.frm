VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogProSelEvaluacionValor 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5835
   ClientLeft      =   570
   ClientTop       =   2040
   ClientWidth     =   11070
   Icon            =   "frmLogProSelEvaluacionValor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   11070
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
      ForeColor       =   &H00000080&
      Height          =   1590
      Left            =   120
      TabIndex        =   24
      Top             =   60
      Width           =   10815
      Begin VB.TextBox TxtDescripcion 
         Appearance      =   0  'Flat
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
         Height          =   495
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   950
         Width           =   8835
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   9240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   630
         Width           =   1260
      End
      Begin VB.TextBox TxtTipo 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   630
         Width           =   6255
      End
      Begin VB.TextBox txtanio 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   300
         Width           =   3255
      End
      Begin VB.CommandButton CmdConsultarProceso 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2670
         TabIndex        =   26
         Top             =   310
         Width           =   350
      End
      Begin VB.TextBox txtObjeto 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   300
         Width           =   1860
      End
      Begin VB.TextBox TxtProSelNro 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8040
         TabIndex        =   37
         Top             =   690
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Proceso Selección"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   700
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nº Proceso"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   345
         Width           =   810
      End
      Begin VB.Label LblMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   8640
         TabIndex        =   34
         Top             =   630
         Width           =   580
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Ejecución"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3120
         TabIndex        =   33
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Objeto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8040
         TabIndex        =   32
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9480
      TabIndex        =   0
      Top             =   5220
      Width           =   1215
   End
   Begin VB.Frame frameEval 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4035
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   10815
      Begin VB.CommandButton cmdCancelaPos 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   8100
         TabIndex        =   23
         Top             =   3540
         Width           =   1215
      End
      Begin VB.ComboBox CboPostores 
         Height          =   315
         ItemData        =   "frmLogProSelEvaluacionValor.frx":08CA
         Left            =   1440
         List            =   "frmLogProSelEvaluacionValor.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   60
         Width           =   9135
      End
      Begin TabDlg.SSTab tabEvaluacion 
         Height          =   3555
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   6271
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   635
         TabCaption(0)   =   "   Registro de Propuestas Técnicas      "
         TabPicture(0)   =   "frmLogProSelEvaluacionValor.frx":08CE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frmEvaluacion"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "   Registro de Propuestas Economicas       "
         TabPicture(1)   =   "frmLogProSelEvaluacionValor.frx":08EA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FramePropuestaEconomica"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame frmEvaluacion 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   3015
            Left            =   120
            TabIndex        =   12
            Top             =   420
            Width           =   10515
            Begin VB.TextBox txtEdit 
               BackColor       =   &H80000001&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000018&
               Height          =   285
               Left            =   3840
               MaxLength       =   10
               TabIndex        =   15
               Top             =   1560
               Visible         =   0   'False
               Width           =   2835
            End
            Begin VB.CommandButton cmdGrabar 
               Caption         =   "Grabar"
               Height          =   375
               Left            =   6660
               TabIndex        =   14
               Top             =   2640
               Width           =   1275
            End
            Begin VB.ComboBox cboGrupoBS 
               Height          =   315
               ItemData        =   "frmLogProSelEvaluacionValor.frx":0906
               Left            =   1620
               List            =   "frmLogProSelEvaluacionValor.frx":0908
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   180
               Width           =   8835
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
               Height          =   1995
               Left            =   60
               TabIndex        =   16
               Top             =   540
               Width           =   10395
               _ExtentX        =   18336
               _ExtentY        =   3519
               _Version        =   393216
               BackColor       =   16777215
               ForeColor       =   -2147483630
               Cols            =   4
               FixedCols       =   0
               ForeColorFixed  =   -2147483646
               BackColorSel    =   -2147483647
               ForeColorSel    =   -2147483624
               BackColorBkg    =   16777215
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483633
               GridColorUnpopulated=   -2147483633
               FocusRect       =   0
               ScrollBars      =   2
               AllowUserResizing=   1
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
               _Band(0).Cols   =   4
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Grupos de Items"
               Height          =   195
               Left            =   180
               TabIndex        =   17
               Top             =   240
               Width           =   1155
            End
         End
         Begin VB.Frame FramePropuestaEconomica 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   3015
            Left            =   -74880
            TabIndex        =   18
            Top             =   480
            Width           =   10575
            Begin VB.TextBox txtEditItem 
               BackColor       =   &H00FFF2E1&
               BorderStyle     =   0  'None
               ForeColor       =   &H00400000&
               Height          =   285
               Left            =   4440
               MaxLength       =   10
               TabIndex        =   20
               Top             =   1080
               Visible         =   0   'False
               Width           =   2835
            End
            Begin VB.CommandButton cmdGrabarPos 
               Caption         =   "Grabar"
               Height          =   375
               Left            =   6660
               TabIndex        =   19
               Top             =   2580
               Width           =   1275
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSItem 
               Height          =   2355
               Left            =   60
               TabIndex        =   21
               Top             =   120
               Width           =   10395
               _ExtentX        =   18336
               _ExtentY        =   4154
               _Version        =   393216
               BackColor       =   16777215
               Cols            =   11
               FixedCols       =   0
               ForeColorFixed  =   -2147483646
               BackColorSel    =   16773857
               ForeColorSel    =   -2147483635
               BackColorBkg    =   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483633
               GridColorUnpopulated=   -2147483633
               Enabled         =   0   'False
               FocusRect       =   0
               ScrollBars      =   2
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _NumberOfBands  =   1
               _Band(0).Cols   =   11
            End
         End
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1080
      End
   End
   Begin VB.Frame FrameConcentimiento 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   1740
      Visible         =   0   'False
      Width           =   10815
      Begin VB.CommandButton cmdConcentimiento 
         Caption         =   "Consentimiento de Buena Pro"
         Height          =   375
         Left            =   5760
         TabIndex        =   2
         Top             =   3480
         Width           =   3495
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSItemBP 
         Height          =   2970
         Left            =   180
         TabIndex        =   3
         Top             =   420
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5239
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   7
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   16773857
         ForeColorSel    =   -2147483635
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
   End
   Begin VB.Frame frmBuenaPro 
      Caption         =   "Evaluación de Postores "
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
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   1740
      Visible         =   0   'False
      Width           =   10815
      Begin MSComctlLib.ProgressBar PBEvaluacion 
         Height          =   315
         Left            =   2820
         TabIndex        =   9
         Top             =   3495
         Visible         =   0   'False
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.CommandButton cmdEval 
         Caption         =   "Evaluar Propuestas"
         Height          =   375
         Left            =   7080
         TabIndex        =   10
         Top             =   3480
         Width           =   2235
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSPostor 
         Height          =   2910
         Left            =   240
         TabIndex        =   22
         Top             =   420
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5133
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   7
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   14942183
         ForeColorSel    =   -2147483635
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin VB.Label lblEvalua 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Evaluando Propuestas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   3500
         Visible         =   0   'False
         Width           =   2430
      End
   End
   Begin VB.Image imgNN 
      Height          =   240
      Left            =   600
      Picture         =   "frmLogProSelEvaluacionValor.frx":090A
      Top             =   5340
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgOK 
      Height          =   240
      Left            =   240
      Picture         =   "frmLogProSelEvaluacionValor.frx":0C4C
      Top             =   5340
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuPostor 
      Caption         =   "MenuPos"
      Visible         =   0   'False
      Begin VB.Menu mnuAsignaGana 
         Caption         =   "Asignar Ganador"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesestimar 
         Caption         =   "Desestimar postor"
      End
   End
End
Attribute VB_Name = "frmLogProSelEvaluacionValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gnProSelNro As Integer, gcBSGrupoCod As String, gnProSelTpoCod As Integer
Dim gnProSelSubTpo As Integer, ban As Boolean, gnNroProceso As Integer, gnAnio As Integer
Dim gcObjeto As String, gnObjeto As Integer, nTipo As Integer, cTitulo As String
    
Public Sub Inicio(ByVal pnTipo As Integer, ByVal psTitulo As String)
    nTipo = pnTipo
    cTitulo = psTitulo
    Me.Show 1
End Sub

Private Sub Form_Load()
CentraForm Me
Me.Caption = cTitulo
FlexPostores
txtanio.Text = Year(gdFecSis)
Select Case nTipo
    Case 1
        FormaFlexFactores
        FormaFlexItem
        ban = False
        'Caption = "Evaluacion de Postores"
'        frmEvaluacion.Visible = True
        frameEval.Visible = True
        tabEvaluacion.Tab = 0
    Case 2
        frmBuenaPro.Visible = True
        'Caption = "Calificacion de Postores"
        PBEvaluacion.value = 0
    Case 3
        FrameConcentimiento.Visible = True
        'Caption = "Concentimiento de Buena Pro"
        FormaFlexItemBP
'    Case 4
'        FormaFlexItem
'        FramePropuestaEconomica.Visible = True
'        Caption = "Evaluacion de Postores"
'        tabEvaluacion.Visible = True
End Select
End Sub

Private Sub cboGrupoBS_Click()
    CargarFactores gnProSelTpoCod, gnProSelSubTpo, Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
    CargarValores Right(CboPostores.Text, 13), Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
End Sub

Private Sub CboPostores_Click()
    If nTipo = 1 Then
        CargarFactores gnProSelTpoCod, gnProSelSubTpo, Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
        CargarValores Right(CboPostores.Text, 13), Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
        GeneraDetalleItemPostor gnProSelNro, Right(CboPostores.Text, 13)
        MSFlex.row = 1
        MSFlex.Col = 2
    End If
End Sub

Private Sub cmdCancelaPos_Click()
    FormaFlexFactores
    FormaFlexItem
    'Caption = "Evaluacion de Postores"
    frameEval.Visible = True
    tabEvaluacion.Tab = 0
    tabEvaluacion.TabEnabled(1) = False
    tabEvaluacion.TabEnabled(0) = True
    txtanio.Text = Year(gdFecSis)
    TxtProSelNro.Text = ""
    TxtTipo.Text = ""
    TxtMonto.Text = ""
    LblMoneda.Caption = ""
    TxtDescripcion.Text = ""
    cboGrupoBS.Clear
    CboPostores.Clear
End Sub

Private Sub cmdConcentimiento_Click()
On Error GoTo cmdConcentimientoErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    If gnProSelNro = 0 Then Exit Sub
    
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select RUC=isnull(i.cPersIDnro,''), cPersNombre = replace(p.cPersNombre,'/',' '), p.cPersDireccDomicilio, p.cPersTelefono, " & _
               " pp.bGanador , nPropEconomica, b.nMonto, cProSelTpoDescripcion, b.cSintesis, nNroProceso, nPlanAnualAnio " & _
               " from LogProSelPostorPropuesta pp " & _
               " inner join LogProcesoSeleccion x on pp.nProSelNro = x.nProSelNro " & _
               " inner join LogProSelTpo t on x.nProSelTpoCod = t.nProSelTpoCod " & _
               " inner join LogProSelItem b on pp.nProSelNro = b.nProSelNro and pp.nProSelItem = b.nProSelItem " & _
               " inner join Persona p  on pp.cPersCod = p.cPersCod " & _
               " left outer join PersID i on p.cPersCod = i.cPersCod and cPersIDTpo=2 " & _
               " where pp.nProSelNro=" & gnProSelNro & " and bganador=1 and pp.nProSelItem=" & Val(MSItemBP.TextMatrix(MSItemBP.row, 6))
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            ImpConcentimientoBuenaProWORD Format(gdFecSis, "dd - mmmm - yyyy"), Rs!cPersDireccDomicilio & " " & Rs!cPersTelefono, Rs!cProSelTpoDescripcion & " N° " & Rs!nNroProceso & "-" & Rs!nPlanAnualAnio & "-CMAC-T S.A.", Rs!cSintesis, Rs!nMonto, _
                Rs!cPersNombre, Rs!nPropEconomica / Rs!nMonto, Rs!nPropEconomica, Format(gdFecSis + 15, "dd - mmmm - yyyy")
        End If
        oCon.CierraConexion
    End If
    Exit Sub
cmdConcentimientoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CmdConsultarProceso_Click()
On Error GoTo msflex_clckErr
    frmLogProSelCnsProcesoSeleccion.Inicio 2
    With frmLogProSelCnsProcesoSeleccion
        If Not .gbBandera Then Exit Sub
        gnProSelNro = .gvnProSelNro
        gcBSGrupoCod = .gvcBSGrupoCod
        gnProSelTpoCod = .gvnProSelTpoCod
        gnProSelSubTpo = .gvnProSelSubTpo
        gnNroProceso = .gvnNroProceso
        TxtProSelNro.Text = gnProSelNro
        txtanio.Text = .gvnAnio
        gnAnio = .gvnAnio
        gcObjeto = .gvcObjeto
        gnObjeto = .gvnObjeto
        TxtTipo.Text = .gvcTipo
        TxtMonto.Text = FNumero(.gvnMonto)
        LblMoneda.Caption = .gvcMoneda
        TxtDescripcion.Text = .gvcDescripcion
    End With
    PBEvaluacion.value = 0
    Select Case nTipo
        Case 1
            CierraEtapa gnProSelNro, cnPresentacionPropuestas
            If Not VerificaEtapaCerrada(gnProSelNro, cnEvaluacionPropuestas) Then
                CargarGrupos gnProSelNro
                GeneraDetalleItem gnProSelNro
                CargarPostores gnProSelNro
                cmdGrabar.Visible = True
                cmdGrabarPos.Visible = True
                'tabEvaluacion.TabEnabled(1) = False
            Else
                MsgBox "Etapa Cerreda", vbInformation, "Aviso"
                cmdGrabar.Visible = False
                cmdGrabarPos.Visible = False
                Exit Sub
            End If
        Case 2
            CierraEtapa gnProSelNro, cnEvaluacionPropuestas
            If Not VerificaEtapaCerrada(gnProSelNro, cnOtorgamientoBP) Then
                CargarGrupos gnProSelNro
                ListaPostores gnProSelNro
                cmdEval.Visible = True
            Else
                cmdEval.Visible = False
                MsgBox "Etapa Cerrada", vbInformation, "Aviso"
                Exit Sub
            End If
        Case 3
            If Not VerificaEtapaCerrada(gnProSelNro, cnConcentimientoBP) Then
                CierraEtapa gnProSelNro, cnOtorgamientoBP
                CargarGrupos gnProSelNro
                GeneraDetalleItemBP gnProSelNro
            Else
                FormaFlexItemBP
                cmdConcentimiento.Visible = False
                MsgBox "Etapa Cerrada", vbInformation, "Aviso"
                Exit Sub
            End If
    End Select
'    CargarPostores gnProSelNro, Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
'    If nTipo = 1 Then
'        CargarFactores gnProSelTpoCod, gnProSelSubTpo, Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
'        CargarValores Right(CboPostores.Text, 13), Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
'    End If
Exit Sub
msflex_clckErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Sub ListaPostores(ByVal psProSelNro As Integer)
Dim oConn As New DConecta
Dim Rs As New ADODB.Recordset
Dim i As Integer, sSQL As String

i = 0
FlexPostores
sSQL = " select ps.nProSelItem, ps.cPersCod, p.cPersNombre, ps.nMoneda, ps.nPropEconomica " & _
       "   from LogProSelPostorPropuesta ps inner join Persona p on ps.cPersCod = p.cPersCod " & _
       " where ps.nProSelNro = " & psProSelNro & " and ps.bGanador = 0 and ps.bDesestimado = 0 " & _
       " "

If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         i = i + 1
         InsRow MSPostor, i
         MSPostor.TextMatrix(i, 0) = psProSelNro
         MSPostor.TextMatrix(i, 1) = Rs!nProSelItem
         MSPostor.TextMatrix(i, 2) = Rs!cPersCod
         MSPostor.TextMatrix(i, 3) = Rs!cPersNombre
         MSPostor.TextMatrix(i, 4) = Rs!nMoneda
         MSPostor.TextMatrix(i, 5) = FNumero(Rs!nPropEconomica)
         Rs.MoveNext
      Loop
   End If
End If
End Sub

Sub FlexPostores()
MSPostor.Clear
MSPostor.Rows = 2
MSPostor.RowHeight(0) = 320
MSPostor.RowHeight(1) = 8
MSPostor.ColWidth(0) = 0
MSPostor.ColWidth(1) = 500: MSPostor.ColAlignment(1) = 4
MSPostor.ColWidth(2) = 1200
MSPostor.ColWidth(3) = 4000
MSPostor.ColWidth(4) = 500: MSPostor.ColAlignment(4) = 4
MSPostor.ColWidth(5) = 1200
End Sub

Private Sub cmdEval_Click()
On Error GoTo cmdEvalErr
Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, cPersCod As String, _
    Puntaje As Currency, i As Integer, gcBSGrupoDescripcion As String, nFactor As Double
    
    If gnProSelNro = 0 Then Exit Sub
    If Not VerificaEtapaCerrada(gnProSelNro, cnOtorgamientoBP) Then
        Set oCon = New DConecta
        If MsgBox("Desea Volver a Evaluar a los Postores...?", vbQuestion + vbYesNo) = vbYes Then
           lblEvalua.Visible = True
           PBEvaluacion.Visible = True
           DoEvents
            nFactor = 100 / cboGrupoBS.ListCount
            If oCon.AbreConexion Then
                sSQL = "update LogProSelPostorPropuesta set bGanador=0, bDesestimado=0, cDesesDescripcion='' where not cDesesDescripcion in ('Propuesta Economica Fuera de Intervalo','no Cubrir todos los Items') and nProSelNro = " & gnProSelNro '& " and nProSelItem = " & nItem  & "' and cPersCod='" & sGanador & "'"
                oCon.Ejecutar sSQL
                sSQL = "update LogProSelEvalResultado set nPuntaje=0 where nProSelNro=" & gnProSelNro
                oCon.Ejecutar sSQL
                i = 0
                Do While i < cboGrupoBS.ListCount
                                        
                    gcBSGrupoCod = Right(cboGrupoBS.List(i), 4)
                    gcBSGrupoDescripcion = Left(cboGrupoBS.List(i), Len(cboGrupoBS.List(i)) - 154)
                    sSQL = "select distinct e.nFactorNro, e.nPuntaje, e.nFormula, f.cFactorDescripcion, v.nProSelItem, v.nValor, f.nTipo, cPersCod  " & _
                        "from LogProSelEvalFactor e " & _
                        "inner join LogProSelFactor f on e.nFactorNro=f.nFactorNro " & _
                        "inner join LogProSelEvalResultado v on e.nFactorNro=v.nFactorNro and e.cBSGrupoCod = v.cBSGrupoCod and v.nProSelNro = e.nProSelNro and e.nProSelItem = v.nProSelItem " & _
                        "where e.nProSelNro=" & gnProSelNro & " and e.cBSGrupoCod='" & gcBSGrupoCod & _
                        "' and nProSelTpoCod= " & gnProSelTpoCod & " and nProSelSubTpo= " & gnProSelSubTpo & _
                        " order by nTipo,nFormula"
                        
                    Set Rs = oCon.CargaRecordSet(sSQL)
                    Do While Not Rs.EOF
                        If Rs!nTipo Then
                            ValidaPuntos gnProSelNro, Rs!nProSelItem, Rs!npuntaje, Rs!nFactorNro, gcBSGrupoCod, Rs!nFormula, Rs!cPersCod
                        Else
                            CalculaPuntaje Rs!nFormula, gnProSelNro, Rs!npuntaje, Rs!nFactorNro, gcBSGrupoCod
                        End If
                        Rs.MoveNext
                    Loop
                    i = i + 1
                    PBEvaluacion.value = i * nFactor
                Loop
                
                '*************************************************************************************
                'economica*****************************************************************************
                '*************************************************************************************
                
                i = 0
                Do While i < cboGrupoBS.ListCount
                                        
                    gcBSGrupoCod = Right(cboGrupoBS.List(i), 4)
                    gcBSGrupoDescripcion = Left(cboGrupoBS.List(i), Len(cboGrupoBS.List(i)) - 154)
                    sSQL = "select distinct e.nFactorNro, e.nPuntaje, e.nFormula, f.cFactorDescripcion, v.nProSelItem, v.nValor, f.nTipo, cPersCod  " & _
                        "from LogProSelEvalFactor e " & _
                        "inner join LogProSelFactor f on e.nFactorNro=f.nFactorNro " & _
                        "inner join LogProSelEvalResultado v on e.nFactorNro=v.nFactorNro and e.cBSGrupoCod = v.cBSGrupoCod and e.nProSelNro = v.nProSelNro and e.nProSelItem = v.nProSelItem " & _
                        "where e.nProSelNro=" & gnProSelNro & " and e.cBSGrupoCod='" & gcBSGrupoCod & _
                        "' and e.nProSelTpoCod= " & gnProSelTpoCod & " and e.nProSelSubTpo= " & gnProSelSubTpo & _
                        " order by nTipo,nFormula"
                        
                    Set Rs = oCon.CargaRecordSet(sSQL)
                    Do While Not Rs.EOF
                        If Rs!nTipo Then
                            CalculaPuntaje Rs!nFormula, gnProSelNro, Rs!npuntaje, Rs!nFactorNro, gcBSGrupoCod
                        End If
                        Rs.MoveNext
                    Loop
                    i = i + 1
                    PBEvaluacion.value = i * nFactor
                
                    '*****************************************************************************************
                    'buena pro
                    '*****************************************************************************************
                
                    OtorgamientoBuenaPro gcBSGrupoCod, gnProSelNro
                Loop
                CierraEtapa gnProSelNro, cnOtorgamientoBP
                oCon.CierraConexion
            End If
        End If
    Else
        MsgBox "Etapa Cerrada", vbInformation, "Aviso"
    End If

    Generar_ActaBuenaPro gnProSelNro, TxtTipo & " N° " & TxtProSelNro & "-" & gnAnio & "-CMAC-T", gcObjeto & ": " & TxtDescripcion.Text, LblMoneda.Caption, TxtMonto.Text
    
    lblEvalua.Visible = False
    PBEvaluacion.Visible = False
    
    Exit Sub
cmdEvalErr:
    Screen.MousePointer = 0
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Function CalculaPuntaje(ByVal nFormula As Integer, ByVal pnProSelNro As Integer, _
                ByVal pnPuntaje As Currency, pnFactor As Integer, pgcBSGrupoCod As String) As Currency
On Error GoTo CalculaPuntajeErr
    Select Case nFormula
        Case 0
            'DIRECTAMENTE
            eval_Directamente pnProSelNro, pnPuntaje, pnFactor, nFormula, gcBSGrupoCod ', pnTipo
            CalculaPuntaje = 0
        Case 1
            'INVERSAMENTE
            eval_Inversamente pnProSelNro, pnPuntaje, pnFactor, nFormula, gcBSGrupoCod ', pnTipo
            CalculaPuntaje = 0
        Case 2
            'RANGO
            eval_Rangos pnProSelNro, pnFactor, gcBSGrupoCod ', pnTipo
            CalculaPuntaje = 0
        Case 3
            'SI/NO
            eval_SINO pnProSelNro, pnPuntaje, pnFactor, gcBSGrupoCod ', pnTipo
            CalculaPuntaje = 0
    End Select
    Exit Function
CalculaPuntajeErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Sub eval_Directamente(ByVal pnProSelNro As Integer, ByVal pnPuntaje As Currency, _
                ByVal pnFactor As Integer, ByVal pnFormula As Integer, pcBSGrupoCod As String)
On Error GoTo DirectamenteErr
    Dim sSQL As String, oCon As DConecta, Rs As ADODB.Recordset, nMenor As Currency, nPuntos As Currency
    Set oCon = New DConecta
    'sSQL = "select distinct e.nFactorNro,e.nFormula,e.nPuntaje, v.cPersCod, v.nValor, v.nProSelNro " & _
            "from LogProSelEvalFactor e " & _
            "inner join LogProSelEvalFactorValor v on e.nFactorNro=v.nFactorNro " & _
            "where e.nFormula=0 and nProSelNro=" & pnProSelNro & " and e.nFactorNro=" & pnfactor & _
            " order by v.nValor "
    sSQL = "select distinct e.nFactorNro,e.nFormula,e.nPuntaje, v.cPersCod, v.nValor, v.nProSelNro " & _
            "from LogProSelEvalFactor e " & _
            "inner join LogProSelEvalResultado v on e.nFactorNro=v.nFactorNro and e.cBSGrupoCod = v.cBSGrupoCod and e.nProSelNro = v.nProSelNro  and e.nProSelItem = v.nProSelItem " & _
            "where e.nFormula=" & pnFormula & " and e.nProSelNro=" & pnProSelNro & " and e.nFactorNro= " & pnFactor & _
            " and e.cBSGrupoCod='" & pcBSGrupoCod & "'" & _
            " order by v.nValor"
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            Rs.MoveLast
            nMenor = Rs!nValor
            nPuntos = pnPuntaje
            sSQL = "update LogProSelEvalResultado set nPuntaje=" & nPuntos & " where nProSelNro=" & pnProSelNro & " and cPersCod='" & Rs!cPersCod & "' and nFactorNro=" & pnFactor & " and cBSGrupoCod='" & pcBSGrupoCod & "'"
            oCon.Ejecutar sSQL
            Rs.MovePrevious
            Do While Not Rs.BOF
                nPuntos = (Rs!nValor * pnPuntaje) / nMenor
                sSQL = "update LogProSelEvalResultado set nPuntaje=" & nPuntos & " where nProSelNro=" & pnProSelNro & " and cPersCod='" & Rs!cPersCod & "' and nFactorNro=" & pnFactor & " and cBSGrupoCod='" & pcBSGrupoCod & "'"
                oCon.Ejecutar sSQL
                Rs.MovePrevious
            Loop
        End If
        oCon.CierraConexion
    End If
    Exit Sub
DirectamenteErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub eval_Inversamente(ByVal pnProSelNro As Integer, ByVal pnPuntaje As Currency, _
                pnFactor As Integer, pnFormula As Integer, pcBSGrupoCod As String)
On Error GoTo DirectamenteErr
    Dim sSQL As String, oCon As DConecta, Rs As ADODB.Recordset, nMayor As Currency, nPuntos As Currency
    Set oCon = New DConecta
    'sSQL = "select distinct e.nFactorNro,e.nFormula,e.nPuntaje, v.cPersCod, v.nValor, v.nProSelNro " & _
            "from LogProSelEvalFactor e " & _
            "inner join LogProSelEvalFactorValor v on e.nFactorNro=v.nFactorNro " & _
            "where e.nFormula=0 and nProSelNro=" & pnProSelNro & " and e.nFactorNro=" & pnFactor & _
            " order by v.nValor "
    sSQL = "select distinct e.nFactorNro,e.nFormula,e.nPuntaje, v.cPersCod, v.nValor, v.nProSelNro " & _
            "from LogProSelEvalFactor e " & _
            "inner join LogProSelEvalResultado v on e.nFactorNro=v.nFactorNro and e.cBSGrupoCod= v.cBSGrupoCod and e.nProSelNro = v.nProSelNro and e.nProSelItem = v.nProSelItem " & _
            "inner join LogProSelPostorPropuesta p on p.nProSelNro = v.nProSelNro and p.nProSelItem = v.nProSelItem and p.cPersCod = v.cPersCod " & _
            "where p.bDesestimado=0 and e.nFormula=" & pnFormula & " and e.nProSelNro=" & pnProSelNro & " and e.nFactorNro=" & pnFactor & " and e.cBSGrupoCod='" & pcBSGrupoCod & "'" & _
            " order by v.nValor"
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            nMayor = Rs!nValor
            nPuntos = pnPuntaje
            sSQL = "update LogProSelEvalResultado set nPuntaje=" & nPuntos & " where nProSelNro=" & pnProSelNro & " and cPersCod='" & Rs!cPersCod & "' and nFactorNro=" & pnFactor & " and cBSGrupoCod='" & pcBSGrupoCod & "'"
            oCon.Ejecutar sSQL
            Rs.MoveNext
            Do While Not Rs.EOF
                nPuntos = (nMayor * pnPuntaje) / Rs!nValor
                sSQL = "update LogProSelEvalResultado set nPuntaje=" & nPuntos & " where nProSelNro=" & pnProSelNro & " and cPersCod='" & Rs!cPersCod & "' and nFactorNro=" & pnFactor & " and cBSGrupoCod='" & pcBSGrupoCod & "'"
                oCon.Ejecutar sSQL
                Rs.MoveNext
            Loop
        End If
        oCon.CierraConexion
    End If
    Exit Sub
DirectamenteErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub eval_Rangos(ByVal pnProSelNro As Integer, pnFactor As Integer, pcBSGrupoCod As String)
On Error GoTo DirectamenteErr
    Dim sSQL As String, oCon As DConecta, Rs As ADODB.Recordset, _
        sSQL1 As String, nPuntos As Currency, rs1 As ADODB.Recordset
    Set oCon = New DConecta
    'sSQL = "select distinct e.nFactorNro,e.nFormula,e.nPuntaje, v.cPersCod, v.nValor, v.nProSelNro, e.cBSGrupoCod " & _
            "from LogProSelEvalFactor e " & _
            "inner join LogProSelEvalFactorValor v on e.nFactorNro=v.nFactorNro " & _
            "where e.nFormula=0 and nProSelNro=" & pnProSelNro & " and e.nFactorNro=" & pnFactor & _
            " order by v.nValor "
    'sSQL = "select distinct v.cPersCod,v.nValor, r.nPuntaje " & _
            "from LogProSelEvalFactorValor v " & _
            "inner join LogProSelEvalFactorRangos r on r.nFactorNro = v.nFactorNro " & _
            "where nProSelNro=" & pnProSelNro & " and v.nFactorNro=" & pnFactor & _
            " and v.nValor between r.nRangoMin and r.nRangoMax and cPersCod='" & pcPersCod & "'"
    sSQL = "select distinct cPersCod,nValor from LogProSelEvalResultado where nProSelNro=" & pnProSelNro & " and nFactorNro=" & pnFactor & " and cBSGrupoCod='" & pcBSGrupoCod & "'"
    
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
            Do While Not Rs.EOF
                sSQL1 = "select distinct  r.nPuntaje from LogProSelEvalFactor f " & _
                        "inner join LogProSelEvalFactorRangos r on  f.nFactorNro=r.nFactorNro and f.nProSelNro = r.nProSelNro and f.nProSelTpoCod = r.nProSelTpoCod and f.nProSelSubTpo = r.nProSelSubTpo and f.nObjeto = r.nObjeto and f.nProSelNro = r.nProSelNro and f.nProSelItem = r.nProSelItem " & _
                        "where " & Rs!nValor & " between r.nRangoMin and r.nRangoMax and f.nFactorNro=" & pnFactor & _
                        " and r.cBSGrupoCod='" & pcBSGrupoCod & "' and f.nProSelNro = " & pnProSelNro
                Set rs1 = oCon.CargaRecordSet(sSQL1)
                If Not rs1.EOF Then
                    nPuntos = rs1!npuntaje
                End If
                Set rs1 = Nothing
                If nPuntos = 0 Then
                    sSQL1 = "select distinct  r.nPuntaje,r.nRangoMax from LogProSelEvalFactor f " & _
                        "inner join LogProSelEvalFactorRangos r on f.nFactorNro=r.nFactorNro and f.nProSelNro = r.nProSelNro and f.nProSelTpoCod = r.nProSelTpoCod and f.nProSelSubTpo = r.nProSelSubTpo and f.nObjeto = r.nObjeto and f.nProSelNro = r.nProSelNro and f.nProSelItem = r.nProSelItem " & _
                        "where f.nFactorNro=" & pnFactor & " and f.nProSelNro = " & pnProSelNro & _
                        " and r.cBSGrupoCod='" & pcBSGrupoCod & "' order by r.nRangoMax desc"
                    Set rs1 = oCon.CargaRecordSet(sSQL1)
                    If Not rs1.EOF Then
                        If Rs!nValor > rs1!nRangoMax Then _
                            nPuntos = rs1!npuntaje
                    End If
                    Set rs1 = Nothing
                End If
                    sSQL = "update LogProSelEvalResultado set nPuntaje=" & nPuntos & " where nProSelNro=" & pnProSelNro & " and cPersCod='" & Rs!cPersCod & "' and nFactorNro=" & pnFactor & " and cBSGrupoCod='" & pcBSGrupoCod & "'"
                    oCon.Ejecutar sSQL
                Set rs1 = Nothing
                Rs.MoveNext
            Loop
        oCon.CierraConexion
    End If
    Exit Sub
DirectamenteErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub eval_SINO(ByVal pnProSelNro As Integer, ByVal pnPuntaje As Currency, pnFactor As Integer, ByVal pcBSGrupoCod As String)
On Error GoTo DirectamenteErr
    Dim sSQL As String, oCon As DConecta, Rs As ADODB.Recordset, nPuntos As Currency
    Set oCon = New DConecta
    'sSQL = "select distinct e.nFactorNro,e.nFormula,e.nPuntaje, v.cPersCod, v.nValor, v.nProSelNro " & _
            "from LogProSelEvalFactor e " & _
            "inner join LogProSelEvalFactorValor v on e.nFactorNro=v.nFactorNro " & _
            "where e.nFormula=0 and nProSelNro=" & pnProSelNro & " and e.nFactorNro=" & pnFactor & _
            " order by v.nValor "
    sSQL = "select distinct cPersCod,nValor from LogProSelEvalResultado where nProSelNro=" & pnProSelNro & " and nFactorNro=" & pnFactor
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            Do While Not Rs.EOF
                nPuntos = Rs!nValor * pnPuntaje
                sSQL = "update LogProSelEvalResultado set nPuntaje=" & nPuntos & " where nProSelNro=" & pnProSelNro & " and cPersCod='" & Rs!cPersCod & "' and nFactorNro=" & pnFactor & " and cBSGrupoCod='" & pcBSGrupoCod & "'"
                oCon.Ejecutar sSQL
                Rs.MoveNext
            Loop
        End If
        oCon.CierraConexion
    End If
    Exit Sub
DirectamenteErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub


Private Sub cmdGrabar_Click()
On Error GoTo cmdGrabarErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, cPersCod As String, _
        i As Integer, nPos As Integer, nValor As Currency
    Set oCon = New DConecta
    
    If gnProSelNro = 0 Then Exit Sub
        
    If Not ValidaValores Then Exit Sub
    
    If Not MSFlex.Enabled Then
        MsgBox "Debe Ingresar Valores Validos...", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If CboPostores.ListCount = 0 Then Exit Sub
    
    If oCon.AbreConexion Then
    
        '********************************************************************************************
        
        sSQL = "select distinct r.nValor, r.nFactorNro from LogProSelEvalResultado r " & _
            " where nProSelNro=" & gnProSelNro & _
            " and cPersCod= '" & Right(CboPostores.Text, 13) & "' and r.cBSGrupoCod='" & Right(cboGrupoBS.Text, 4) & "' order by nFactornro"
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            If MsgBox("Esta Seguro que Desea Cambiar los Datos del Postor...", vbQuestion + vbYesNo) = vbNo Then
'                sSQL = "delete LogProSelEvalResultado " & _
                        " where nProSelNro=" & gnProSelNro & _
                        " and cPersCod= '" & Right(CboPostores.Text, 13) & "' and cBSGrupoCod='" & Right(cboGrupoBS.Text, 4) & "'"
'                oCon.Ejecutar sSQL
            'Else
                Exit Sub
            End If
        End If
        sSQL = ""
        
        '********************************************************************************************
    
        cPersCod = Right(CboPostores.Text, 13)
        Set Rs = oCon.CargaRecordSet("select nProSelItem from LogProSelItem where nProSelNro=" & gnProSelNro & " and cBSGrupoCod='" & Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4) & "'")
        i = 1
        Do While i < MSFlex.Rows
            
            MSFlex.Col = 2
            MSFlex.row = i
            If MSFlex.CellPicture = imgNN Then
                nValor = 0
            ElseIf MSFlex.CellPicture = imgOK Then
                nValor = 1
            Else
                nValor = Val(MSFlex.TextMatrix(i, 2))
            End If
            
            'Do While Not Rs.EOF
            If Not Rs.EOF Then
                sSQL = "declare @tmp int "
                sSQL = sSQL & " set @tmp=(select count(*) from LogProSelEvalResultado where nProSelNro = " & gnProSelNro & " and nProSelItem = " & Rs!nProSelItem & " and cPersCod = '" & cPersCod & "' and nFactorNro = " & MSFlex.TextMatrix(i, 0) & " and cBSGrupoCod = '" & Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4) & "') "
                sSQL = sSQL & " if @tmp=0 "
                sSQL = sSQL & " insert into LogProSelEvalResultado(nProSelNro,cPersCod,nProSelItem,nFactorNro,nValor,cMovNro,cBSGrupoCod) " & _
                              " values(" & gnProSelNro & ",'" & cPersCod & "'," & Rs!nProSelItem & "," & MSFlex.TextMatrix(i, 0) & "," & nValor & ",'" & GetLogMovNro & "','" & Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4) & "') "
                sSQL = sSQL & " else "
                sSQL = sSQL & " update LogProSelEvalResultado set nValor =" & nValor & _
                              " ,cMovNro = '" & GetLogMovNro & "', " & _
                              " nPuntaje = 0 " & _
                              " where nProSelNro = " & gnProSelNro & " and nProSelItem = " & Rs!nProSelItem & " and cPersCod = '" & cPersCod & "' and nFactorNro = " & MSFlex.TextMatrix(i, 0) & " and cBSGrupoCod = '" & Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4) & "' "
                oCon.Ejecutar sSQL
            End If
            '    ban = True
            '    Rs.MoveNext
            'Loop
            'If ban Then Rs.MoveFirst
            i = i + 1
        Loop
        oCon.CierraConexion
        
        tabEvaluacion.TabEnabled(1) = ValidaPropuestaTecnica(gnProSelNro, Right(CboPostores, 13))
'        tabEvaluacion.TabEnabled(0) = False
'        tabEvaluacion.Tab = 1
        MsgBox "Valores Registrados", vbInformation
'        MSFlex.Enabled = False
    End If
    Exit Sub
cmdGrabarErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Function ValidaPropuestaTecnica(ByVal pnProSelNro As Integer, ByVal pcPersCod As String) As Boolean
On Error GoTo ValidaPropuestaTecnicaErr
Dim oCon As DConecta, Rs As ADODB.Recordset, sSQL As String, Nro As Integer, i As Integer
Set oCon = New DConecta
If oCon.AbreConexion Then
    Do While i < cboGrupoBS.ListCount
        sSQL = "select * from LogProSelEvalResultado where nProSelNro = " & pnProSelNro & " and cBSGrupoCod = '" & Right(cboGrupoBS.List(i), 4) & "' and cPersCod = '" & pcPersCod & "'"
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            Nro = Rs(0)
            ValidaPropuestaTecnica = True
        Else
            ValidaPropuestaTecnica = False
            Exit Function
        End If
        Set Rs = Nothing
        i = i + 1
    Loop
End If
oCon.CierraConexion
Exit Function
ValidaPropuestaTecnicaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

'Private Sub cmdPersona_Click()
'With frmLogCnsPostores
'    .Inicio gnProSelNro, ""
'    txtPersCod.Text = .gcPersCod
'    txtPersona.Text = .gcPersNombre
'    If txtPersCod.Text <> "" Then
'        MSItem.Enabled = True
'        'MSItem.SetFocus
'        GeneraDetalleItemPostor gnProSelNro, txtPersCod
'    End If
'End With
'End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Sub FormaFlexItemBP()
MSItemBP.Clear
MSItemBP.Rows = 2
MSItemBP.RowHeight(0) = 320
MSItemBP.RowHeight(1) = 8
MSItemBP.ColWidth(0) = 250
MSItemBP.ColWidth(1) = 850:   MSItemBP.ColAlignment(1) = 4:  MSItemBP.TextMatrix(0, 1) = " Item"
MSItemBP.ColWidth(2) = 0:   MSItemBP.ColAlignment(2) = 4:  MSItemBP.TextMatrix(0, 2) = " Código"
MSItemBP.ColWidth(3) = 7600:  MSItemBP.TextMatrix(0, 3) = " Descripción"
MSItemBP.ColWidth(4) = 1200:  MSItemBP.TextMatrix(0, 4) = " Cantidad"
MSItemBP.ColWidth(5) = 0:  MSItemBP.TextMatrix(0, 5) = " nProSelNro"
MSItemBP.ColWidth(6) = 0:  MSItemBP.TextMatrix(0, 6) = " nProSelItem"
End Sub

Sub GeneraDetalleItemBP(vProSelNro As Integer)
Dim oConn As New DConecta, Rs As New ADODB.Recordset, i As Integer, nSuma As Currency
Dim sSQL As String, sGrupo As String

sSQL = ""
nSuma = 0
FormaFlexItem

If oConn.AbreConexion Then

    sSQL = "select v.nProSelNro, v.nProSelItem, b.cBSGrupoDescripcion, b.cBSGrupoCod,x.cBSCod, y.cBSDescripcion, x.nCantidad, x.nMonto " & _
            "from LogProSelItem v " & _
            "inner join BSGrupos b on v.cBSGrupoCod = b.cBSGrupoCod " & _
            "inner join LogProSelItemBS x on v.nProSelNro = x.nProSelNro and v.nProSelItem = x.nProSelItem " & _
            "inner join LogProSelBienesServicios y on x.cBSCod = y.cProSelBSCod " & _
            "where v.nProSelNro = " & vProSelNro & " order by v.nProSelItem, b.cBSGrupoDescripcion "

   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
        If sGrupo <> Rs!cBSGrupoDescripcion Then
         sGrupo = Rs!cBSGrupoDescripcion
         i = i + 1
         InsRow MSItemBP, i
         MSItemBP.Col = 0
         MSItemBP.row = i
         MSItemBP.CellFontSize = 10
         MSItemBP.CellFontBold = True
         MSItemBP.TextMatrix(i, 0) = "+"
         MSItemBP.TextMatrix(i, 1) = Rs!nProSelItem
         MSItemBP.TextMatrix(i, 3) = Rs!cBSGrupoDescripcion
         MSItemBP.TextMatrix(i, 4) = ""
         MSItemBP.TextMatrix(i, 5) = Rs!nProselNro
         MSItemBP.TextMatrix(i, 6) = Rs!nProSelItem
        End If
        i = i + 1
        InsRow MSItemBP, i
        MSItemBP.RowHeight(i) = 0
        MSItemBP.TextMatrix(i, 1) = Rs!cBSCod
        MSItemBP.TextMatrix(i, 3) = Rs!cBSDescripcion
        MSItemBP.TextMatrix(i, 4) = Rs!nCantidad
        MSItemBP.TextMatrix(i, 5) = Rs!nProselNro
        MSItemBP.TextMatrix(i, 6) = Rs!nProSelItem
        Rs.MoveNext
      Loop
   End If
End If
End Sub


Private Sub CargarGrupos(ByVal pnProSelNro As Integer)
    On Error GoTo CargarGruposErr
    Dim sSQL As String, oCon As DConecta, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select distinct g.cBSGrupoCod, g.cBSGrupoDescripcion from LogProSelItem v " & _
                "inner join BSGrupos g on g.cBSGrupoCod = v.cBSGrupoCod where v.nProSelNro = " & pnProSelNro
        Set Rs = oCon.CargaRecordSet(sSQL)
        cboGrupoBS.Clear
        Do While Not Rs.EOF
            cboGrupoBS.AddItem Rs!cBSGrupoDescripcion & Space(200) & Rs!cBSGrupoCod, cboGrupoBS.ListCount
            'cboGrupoBS.ItemData(cboGrupoBS.ListCount - 1) = Rs!cBSGrupoCod
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    If cboGrupoBS.ListCount > 0 Then cboGrupoBS.ListIndex = 0
    Exit Sub
CargarGruposErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CargarPostores(ByVal pnProSelNro As Integer)
On Error GoTo CargarPostoresErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        'sSQL = "select distinct p.cPersCod, x.cPersNombre from LogProSelPostorPropuesta p " & _
                "inner join Persona x on x.cPersCod=p.cPersCod " & _
                "where p.bDesestimado=0 and p.nProSelNro=" & pnProSelNro
        sSQL = "select distinct p.cPersCod, x.cPersNombre from LogProSelPostor p " & _
                "inner join Persona x on x.cPersCod=p.cPersCod " & _
                "where p.nPresentoProp=1 and p.nProSelNro=" & pnProSelNro
        Set Rs = oCon.CargaRecordSet(sSQL)
        CboPostores.Clear
        Do While Not Rs.EOF
            CboPostores.AddItem Rs!cPersNombre & Space(150) & Rs!cPersCod, CboPostores.ListCount
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    If CboPostores.ListCount > 0 Then CboPostores.ListIndex = 0
    Exit Sub
CargarPostoresErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Sub FormaFlexFactores()
With MSFlex
    .Clear
    .Rows = 2
    .Cols = 6
    .RowHeight(0) = 320
    .RowHeight(1) = 10
    .ColWidth(0) = 1000:     .TextMatrix(0, 0) = "Codigo":        '.ColAlignment(1) = 4
    .ColWidth(1) = 4800:    .TextMatrix(0, 1) = "Factor"
    .ColWidth(2) = 1000:     .TextMatrix(0, 2) = "Valor":        .ColAlignment(2) = 4
    .ColWidth(3) = 2000:     .TextMatrix(0, 3) = "Unidades":
    .ColWidth(4) = 0:     .TextMatrix(0, 4) = "Tipo":
    .ColWidth(5) = 1200:     .TextMatrix(0, 5) = "Propuesta":
End With
End Sub

Private Sub CargarFactores(ByVal pnProSelTpoCod As Integer, pnProSelSubTpo As Integer, pcBSgrupo As String)
    On Error GoTo CargarFactoresErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, sSQL As String, i As Integer
    sSQL = "select x.nFactorNro, x.cFactorDescripcion, f.nFormula, x.cUnidades, x.nTipo from LogProSelEvalFactor f " & _
        "inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
        "where nTipo=0 and nVigente=1 and nProSelTpoCod=" & pnProSelTpoCod & " and nProSelSubTpo=" & pnProSelSubTpo & " and cBSGrupoCod='" & pcBSgrupo & "' and nProSelNro = " & gnProSelNro
    'sSQL = "select distinct x.nFactorNro, x.cFactorDescripcion, v.nValor from LogProSelEvalFactor f " & _
        "inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
        "left outer join LogProSelEvalFactorValor v on f.nFactorNro = v.nFactorNro " & _
        "where nProSelTpoCod=" & pnProSelTpoCod & " and nProSelSubTpo=" & pnProSelSubTpo & " and cBSGrupoCod='0" & pcBSgrupo & "'"
    Set oCon = New DConecta
    FormaFlexFactores
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            With MSFlex
                i = i + 1
                InsRow MSFlex, i
                .TextMatrix(i, 0) = Rs!nFactorNro
                .TextMatrix(i, 1) = Rs!cFactorDescripcion
                If Rs!nFormula = 3 Then
                    .row = i
                    .Col = 2
                    .CellPictureAlignment = 4
                    Set .CellPicture = imgNN
                    .row = 1
                End If
                .TextMatrix(i, 3) = Rs!cUnidades
                .TextMatrix(i, 4) = 0
                .TextMatrix(i, 5) = IIf(Rs!nTipo = 0, "Tecnica", "Economica")
            End With
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    Exit Sub
CargarFactoresErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLogProSelEvaluacionValor = Nothing
End Sub



'*********************************************************************
'PROCEDIMIENTOS DEL FLEX
'*********************************************************************

Private Sub MSFlex_Click()
    If MSFlex.CellPicture = imgNN Then
        Set MSFlex.CellPicture = imgOK
    ElseIf MSFlex.CellPicture = imgOK Then
        Set MSFlex.CellPicture = imgNN
    End If
End Sub


Private Sub MSFlex_GotFocus()
If txtEdit.Visible = False Then Exit Sub
MSFlex = txtEdit
txtEdit.Visible = False
End Sub

Private Sub MSFlex_LeaveCell()
If txtEdit.Visible = False Then Exit Sub
MSFlex = txtEdit
txtEdit.Visible = False
End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
'If MSFlex.Col >= 1 And MSFlex.Col < 3 Then
'   EditaFlex MSFlex, txtEdit, KeyAscii
'End If
On Error GoTo msgError
If MSFlex.TextMatrix(MSFlex.row, 5) = "Economica" Then Exit Sub

Select Case MSFlex.Col
    Case 2
    'Not IsNumeric(MSFlex.TextMatrix(MSFlex.Row, 4)) And
        If IsNumeric(Chr(KeyAscii)) And MSFlex.CellPicture <> imgNN And MSFlex.CellPicture <> imgOK Then
            EditaFlex MSFlex, txtEdit, KeyAscii
        ElseIf KeyAscii = 13 Then
            MSFlex_Click
        End If
End Select
Exit Sub
msgError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Sub EditaFlex(MSFlex As Control, Edt As Control, KeyAscii As Integer)
Select Case KeyAscii
    Case 0 To 32
         Edt = MSFlex
         Edt.SelStart = 1000
    Case Else
         Edt = Chr(KeyAscii)
         Edt.SelStart = 1
End Select
Edt.Move MSFlex.Left + MSFlex.CellLeft - 15, MSFlex.Top + MSFlex.CellTop - 15, _
         MSFlex.CellWidth, MSFlex.CellHeight
Edt.Visible = True
Edt.SetFocus
End Sub

Private Sub MSItemBP_DblClick()
On Error GoTo MSItemErr
    Dim i As Integer, bTipo As Boolean
    With MSItemBP
        If Trim(.TextMatrix(.row, 0)) = "-" Then
           .TextMatrix(.row, 0) = "+"
           i = .row + 1
           bTipo = True
        ElseIf Trim(.TextMatrix(.row, 0)) = "+" Then
           .TextMatrix(.row, 0) = "-"
           i = .row + 1
           bTipo = False
        Else
            Exit Sub
        End If
        
        Do While i < .Rows
            If Trim(.TextMatrix(i, 0)) = "+" Or Trim(.TextMatrix(i, 0)) = "-" Then
                Exit Sub
            End If
            
            If bTipo Then
                .RowHeight(i) = 0
            Else
                .RowHeight(i) = 260
            End If
            i = i + 1
        Loop
    End With
Exit Sub
MSItemErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub mnuDesestimar_Click()
Dim k As Integer, cPersCod As String
Dim sSQL As String
Dim oConn As New DConecta, nProselNro As Integer

k = MSPostor.row
nProselNro = MSPostor.TextMatrix(k, 0)
cPersCod = MSPostor.TextMatrix(k, 2)

If MsgBox("¿ Está seguro de asignar Ganador al postor indicado ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then

   sSQL = "UPDATE LogProSelPostorPropuesta SET bGanador=0, bDesestimado = 1 WHERE cPersCod = '" & cPersCod & "' and nProSelNro = " & nProselNro & " "
   If oConn.AbreConexion Then
      oConn.Ejecutar sSQL
      oConn.CierraConexion
   End If
End If

End Sub

Private Sub mnuAsignaGana_Click()
Dim k As Integer, cPersCod As String
Dim sSQL As String
Dim oConn As New DConecta, nProselNro As Integer
Dim nMoneda As Integer

k = MSPostor.row
nProselNro = MSPostor.TextMatrix(k, 0)
cPersCod = MSPostor.TextMatrix(k, 2)

If MsgBox("¿ Está seguro de asignar Ganador al postor indicado ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   nMoneda = IIf(LblMoneda.Caption = "S/.", 1, 2)
   sSQL = "UPDATE LogProSelPostorPropuesta SET nMoneda=" & nMoneda & ",nPropEconomica = " & VNumero(TxtMonto.Text) & ",bGanador = 1, bDesestimado = 0 WHERE cPersCod = '" & cPersCod & "' and nProSelNro = " & nProselNro & " "
   If oConn.AbreConexion Then
      oConn.Ejecutar sSQL
      oConn.CierraConexion
      ListaPostores nProselNro
   End If
End If
End Sub

Private Sub MSPostor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu mnuPostor
End If
End Sub

'Private Sub MSItem_DblClick()
'On Error GoTo MSItemErr
'    Dim i As Integer, bTipo As Boolean
'    With MSItem
'        If Trim(.TextMatrix(.Row, 0)) = "-" Then
'           .TextMatrix(.Row, 0) = "+"
'           i = .Row + 1
'           bTipo = True
'        ElseIf Trim(.TextMatrix(.Row, 0)) = "+" Then
'           .TextMatrix(.Row, 0) = "-"
'           i = .Row + 1
'           bTipo = False
'        Else
'            Exit Sub
'        End If
'        Do While i < .Rows
'            If Trim(.TextMatrix(i, 0)) = "+" Or Trim(.TextMatrix(i, 0)) = "-" Then
'                Exit Sub
'            End If
'
'            If bTipo Then
'                .RowHeight(i) = 0
'            Else
'                .RowHeight(i) = 260
'            End If
'            i = i + 1
'        Loop
'    End With
'Exit Sub
'MSItemErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'
'End Sub

'Private Sub MSItem_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then MSItem_DblClick
'End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
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

'Private Sub txtEdit_KeyPress(KeyAscii As Integer)
'If KeyAscii = Asc(vbCr) Then
'   KeyAscii = 0
'End If
'End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then
       KeyAscii = 0
    End If
Select Case MSFlex.Col
    Case 2
        KeyAscii = DigNumDec(txtEdit, KeyAscii)
'        If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Select
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode MSFlex, txtEdit, KeyCode, Shift
End Sub

Sub EditKeyCode(MSFlex As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27
         Edt.Visible = False
         MSFlex.SetFocus
    Case 13
         MSFlex.SetFocus
    Case 37                     'Izquierda
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col > 1 Then
            MSFlex.Col = MSFlex.Col - 1
         End If
    Case 39                     'Derecha
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col < MSFlex.Cols - 1 Then
            MSFlex.Col = MSFlex.Col + 1
         End If
    Case 38
         MSFlex.SetFocus
         DoEvents
         If MSFlex.row > MSFlex.FixedRows + 1 Then
            MSFlex.row = MSFlex.row - 1
         End If
    Case 40
         MSFlex.SetFocus
         DoEvents
         'If MSFlex.Row < MSFlex.FixedRows - 1 Then
         If MSFlex.row < MSFlex.Rows - 1 Then
            MSFlex.row = MSFlex.row + 1
         End If
End Select
End Sub

Private Function ValidaValores() As Boolean
    Dim i As Integer
    ValidaValores = True
    With MSFlex
        i = 1
        Do While i < .Rows
            .row = i: .Col = 2
            If Val(.TextMatrix(i, 2)) <= 0 And .CellPicture <> imgNN And .CellPicture <> imgOK Then
                ValidaValores = False
                Exit Function
            End If
            i = i + 1
        Loop
    End With
End Function

'Private Sub CargarPEconomica(pcPersCod As String)
'    Dim i As Integer
'    i = 1
'    With MSFlex
'        Do While i < .Rows
'            If .TextMatrix(i, 4) Then
'                .TextMatrix(i, 2) = fPropuestaEconomica(pcPersCod, gnProSelNro)
'            End If
'            i = i + 1
'        Loop
'    End With
'End Sub

Private Sub CargarValores(ByVal pcPersCod As String, pcBSGrupoCod As String)
On Error GoTo CargarValoresErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, sSQL As String, i As Integer
    'sSQL = "select distinct r.nValor, r.nFactorNro, p.nPropEconomica  from LogProSelEvalResultado r " & _
            " inner join LogProSelPostorPropuesta p on r.nProSelNro = p.nProSelNro " & _
            " and r.nProSelItem = p.nProSelItem and p.cPersCod = r.cPersCod" & _
            " where p.nProSelNro=" & gnProSelNro & _
            " and p.cPersCod= '" & pcPersCod & "' and r.cBSGrupoCod='" & pcBSGrupoCod & "' order by nFactornro"
    sSQL = "select distinct r.nValor, r.nFactorNro from LogProSelEvalResultado r " & _
            " where nProSelNro=" & gnProSelNro & _
            " and cPersCod= '" & pcPersCod & "' and r.cBSGrupoCod='" & pcBSGrupoCod & "' order by nFactornro"
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        i = 1
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Rs.EOF Then
            tabEvaluacion.TabEnabled(1) = False
            tabEvaluacion.Tab = 0
        Else
            tabEvaluacion.TabEnabled(1) = True
            tabEvaluacion.TabEnabled(0) = True
            tabEvaluacion.Tab = 0
        End If
        Do While Not Rs.EOF
            With MSFlex
                If .TextMatrix(i, 0) = Rs!nFactorNro Then
                    .row = i: .Col = 2
                    If .CellPicture <> imgNN And .CellPicture <> imgOK Then
                        .TextMatrix(i, 2) = Rs!nValor
                    ElseIf Rs!nValor = 1 Then
                        Set .CellPicture = imgOK
                    Else
                        Set .CellPicture = imgNN
                    End If
                End If
                i = i + 1
                If i = .Rows Then Exit Do
            End With
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    
    i = 1
    Do While i < MSFlex.Rows
        If MSFlex.TextMatrix(i, 5) = "Economica" Then
            MSFlex.TextMatrix(i, 2) = FNumero(CargarPropuestaEconomica(gnProSelNro, Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4), pcPersCod))
        End If
        i = i + 1
    Loop
    
    Set Rs = Nothing
    If MSFlex.Enabled And MSFlex.Visible Then MSFlex.SetFocus
    Exit Sub
CargarValoresErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

'Private Function fPropuestaEconomica(pcPersCod As String, pnProSelNro As Integer) As Currency
'On Error GoTo fPropuestaEconomicaErr
'    Dim sSQL As String, oCon As DConecta, Rs As ADODB.Recordset
'    sSQL = "SELECT nPE=sum(e.nPropEconomica * i.nCantidad), nNroItem=count(i.nProSelNro), e.cPersCod " & _
'        "   FROM LogProSelPostorPropuesta e " & _
'        "   inner join LogProSelItem i on i.nProSelItem = e.nProSelItem and i.nProSelNro = e.nProSelNro " & _
'        "where e.nProSelNro=" & pnProSelNro & " and e.cPersCod='" & pcPersCod & "' " & _
'        "group by e.cPersCod"
'    Set oCon = New DConecta
'    If oCon.AbreConexion Then
'        Set Rs = oCon.CargaRecordSet(sSQL)
'        If Not Rs.EOF Then
'            fPropuestaEconomica = Rs!nPE
'        End If
'        oCon.CierraConexion
'    End If
'    Exit Function
'fPropuestaEconomicaErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'End Function

'Private Sub EvaluarPropuestaEconomica()
'    On Error GoTo EvaluarPropuestaEconomicaErr
'    sSQL = "select distinct e.nFactorNro, e.nPuntaje, e.nFormula, f.cFactorDescripcion " & _
'            "from LogProSelEvalFactor e " & _
'            "inner join LogProSelFactor f on e.nFactorNro=f.nFactorNro " & _
'            "inner join LogProSelEvalResultado v on e.nFactorNro=v.nFactorNro " & _
'            "where nProSelNro=" & gnProSelNro & " and e.cBSGrupoCod='" & gcBSGrupoCod & "' order by nFormula"
'    Exit Sub
'EvaluarPropuestaEconomicaErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'End Sub

Private Function ValidaMonto(ByVal pMontoRef As Currency, ByVal pMonto As Currency, ByVal pnProSelNro As Integer) As Boolean
On Error GoTo ValidaMontoErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, sSQL As String, max As Currency, min As Currency
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select nRangoMayor, nRangoMenor from LogProcesoSeleccion where nProSelNro=" & pnProSelNro
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            max = Rs!nRangoMayor / 100#
            min = Rs!nRangoMenor / 100#
        End If
        Set Rs = Nothing
        oCon.CierraConexion
    End If
    ValidaMonto = False
    If pMonto >= (pMontoRef * min) And pMonto <= (pMontoRef * max) Then _
        ValidaMonto = True
    Exit Function
ValidaMontoErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Function

Private Function CargarPropuestaEconomica(ByVal pnProSelNro As Integer, ByVal pcBSGrupoCod As String, ByVal pcPersCod As String) As Currency
    On Error GoTo CargarPropuestaEconomicaErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, sSQL As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select p.* from LogProSelPostorPropuesta p " & _
                "inner join LogProSelItem i on i.nProSelNro = p.nProSelNro and i.nProSelItem = p.nProSelItem " & _
                "where p.nProSelNro=" & pnProSelNro & " and i.cBSGrupoCod='" & pcBSGrupoCod & "' and p.cPersCod= '" & pcPersCod & "' "
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            CargarPropuestaEconomica = Rs!nPropEconomica
        End If
        oCon.CierraConexion
    End If
    Exit Function
CargarPropuestaEconomicaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Sub OtorgamientoBuenaPro(ByVal pcBSGrupoCod As String, ByVal pnProSelNro As Integer)
    On Error GoTo OtorgamientoBuenaProErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, sGanador As String, _
        nMayor As Currency, nItem As Integer
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        
        'sSQL = "select nPTotal=sum(nPuntaje),cPersCod,nProSelItem,nValor from LogProSelEvalResultado where nProSelNro=" & pnProSelNro & " and cBSGrupoCod='" & pcBSGrupoCod & "' group by cPersCod, nProSelItem, nValor"
        sSQL = "select nPTotal=sum(r.nPuntaje),r.cPersCod,r.nProSelItem " & _
                "from LogProSelEvalResultado r inner join LogProSelPostorPropuesta p on " & _
                "r.nProSelNro = P.nProSelNro And r.nProSelItem = P.nProSelItem And r.cPersCod = P.cPersCod " & _
                "where r.nProSelNro=" & pnProSelNro & " and r.cBSGrupoCod='" & pcBSGrupoCod & "' and p.bDesestimado=0 " & _
                "group by r.cPersCod, r.nProSelItem"
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            If nMayor < Rs!nPTotal Then
               nMayor = Rs!nPTotal
               sGanador = Rs!cPersCod
               nItem = Rs!nProSelItem
            End If
            Rs.MoveNext
        Loop
        Set Rs = Nothing
        
        sSQL = "update LogProSelPostorPropuesta set bGanador=1 where nProSelNro = " & pnProSelNro & " and nProSelItem = " & nItem & " and cPersCod='" & sGanador & "'"
        oCon.Ejecutar sSQL
                
        oCon.CierraConexion
    End If
    Exit Sub
OtorgamientoBuenaProErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Function PuntajeMinimo(ByVal pnProSelNro As Integer) As Integer
    On Error GoTo PuntajeMinimoErr
        Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
        Set oCon = New DConecta
        If oCon.AbreConexion Then
            sSQL = "select nPuntajeMinimoTecnico from LogProcesoSeleccion where nProSelNro=" & pnProSelNro
            Set Rs = oCon.CargaRecordSet(sSQL)
            If Not Rs.EOF Then
                PuntajeMinimo = Rs(0)
            End If
            Set Rs = Nothing
            oCon.CierraConexion
        End If
    Exit Function
PuntajeMinimoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Function ValidaPuntos(ByVal pnProSelNro As Integer, ByVal pnProSelItem As Integer, ByVal pnPuntaje As Integer, ByVal pnFactorNro As Integer, ByVal gcBSGrupoCod As String, ByVal pnFormula As Integer, ByVal pcPersCod As String) As Boolean
On Error GoTo ValidaPuntosErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, nPuntajeMinimo As Integer
    Set oCon = New DConecta
    nPuntajeMinimo = PuntajeMinimo(gnProSelNro)
    If oCon.AbreConexion Then
        sSQL = "select nTotal=sum(r.nPuntaje),r.nProSelItem " & _
               " from LogProSelEvalResultado r " & _
               " inner join LogProSelPostorPropuesta p on r.nProSelNro = P.nProSelNro And r.nProSelItem = P.nProSelItem And r.cPersCod = P.cPersCod " & _
               " inner join LogProSelFactor f on r.nFactorNro = f.nFactorNro " & _
               " where f.nTipo=0 and r.nProSelNro=" & pnProSelNro & " and r.nProSelItem=" & pnProSelItem & " and r.cPersCod= '" & pcPersCod & "'" & _
               " group by r.cPersCod, r.nProSelItem"
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            If Rs!nTotal < nPuntajeMinimo Then
                sSQL = "update LogProSelPostorPropuesta set bDesestimado=1,cDesesDescripcion='No alcanzar el puntaje minimo en la propuesta tecnica' " & _
                       " where nProSelNro=" & pnProSelNro & " and nProSelItem=" & pnProSelItem & " and cPersCod='" & pcPersCod & "' "
                oCon.Ejecutar sSQL
                ValidaPuntos = False
'            Else
'                CalculaPuntaje pnFormula, gnProSelNro, pnPuntaje, pnFactorNro, gcBSGrupoCod
            End If
        End If
        oCon.CierraConexion
    End If
    Exit Function
ValidaPuntosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

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
'            gcBSGrupoCod = rs!cBSGrupoCod
            gnProSelTpoCod = Rs!nProSelTpoCod
            gnProSelSubTpo = Rs!nProSelSubTpo
            gnNroProceso = Rs!nNroProceso
            gnAnio = txtanio
            TxtTipo.Text = Rs!cProSelTpoDescripcion
            TxtMonto.Text = FNumero(Rs!nProSelMonto)
            LblMoneda.Caption = IIf(Rs!nMoneda = 1, "S/.", "$")
            TxtDescripcion.Text = Rs!cSintesis
        Else
            gnProSelNro = 0
            gcBSGrupoCod = ""
            gnProSelTpoCod = 0
            gnProSelSubTpo = 0
            gnNroProceso = 0
            gnAnio = 0
            gcObjeto = ""
            gnObjeto = 0
            TxtTipo.Text = ""
            TxtMonto.Text = ""
            LblMoneda.Caption = ""
            TxtDescripcion.Text = ""
            CargarPostores gnProSelNro ', Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
            CboPostores_Click
            MsgBox "Proceso no Existe", vbInformation
        End If
    End If
    PBEvaluacion.value = 0
    Select Case nTipo
        Case 1
            CierraEtapa gnProSelNro, cnPresentacionPropuestas
            If Not VerificaEtapaCerrada(gnProSelNro, cnEvaluacionPropuestas) Then
                CargarGrupos gnProSelNro
                CargarGrupos gnProSelNro
                GeneraDetalleItem gnProSelNro
                CargarPostores gnProSelNro
                cmdGrabar.Visible = True
                cmdGrabarPos.Visible = True
'                tabEvaluacion.TabEnabled(1) = False
            Else
                MsgBox "Etapa Cerreda", vbInformation, "Aviso"
                cmdGrabar.Visible = False
                cmdGrabarPos.Visible = False
                Exit Sub
            End If
        Case 2
            CierraEtapa gnProSelNro, cnEvaluacionPropuestas
            If Not VerificaEtapaCerrada(gnProSelNro, cnOtorgamientoBP) Then
                CargarGrupos gnProSelNro
                cmdEval.Visible = True
            Else
                cmdEval.Visible = False
                MsgBox "Etapa Cerrada", vbInformation, "Aviso"
                Exit Sub
            End If
        Case 3
'            CierraEtapa gnProSelNro, cnOtorgamientoBP
'            CargarGrupos gnProSelNro
'            CargarGrupos gnProSelNro
'            GeneraDetalleItemBP gnProSelNro
            If Not VerificaEtapaCerrada(gnProSelNro, cnConcentimientoBP) Then
                CierraEtapa gnProSelNro, cnOtorgamientoBP
                CargarGrupos gnProSelNro
                GeneraDetalleItemBP gnProSelNro
                cmdConcentimiento.Visible = True
            Else
                FormaFlexItemBP
                cmdConcentimiento.Visible = False
                MsgBox "Etapa Cerrada", vbInformation, "Aviso"
                Exit Sub
            End If
    End Select
    Exit Sub
ConsultarProcesoNroErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
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

'****************************************************************************************************
'*****************************************************************************************************


Private Sub MSItem_KeyPress(KeyAscii As Integer)
On Error GoTo MSItemErr
    Dim i As Integer
    If KeyAscii = 13 Then
        MSItem_DblClick
    End If
    
    Select Case MSItem.Col
        Case 7
            If Val(MSItem.TextMatrix(MSItem.row, 4)) = 0 Then Exit Sub
            If IsNumeric(Chr(KeyAscii)) Then _
                EditaFlex MSItem, txtEditItem, KeyAscii
        Case 9
            If MSItem.TextMatrix(MSItem.row, 0) = "-" Or MSItem.TextMatrix(MSItem.row, 0) = "+" Then
                If IsNumeric(Chr(KeyAscii)) Then _
                    EditaFlex MSItem, txtEditItem, KeyAscii
            End If
    End Select
Exit Sub
MSItemErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Sub GeneraDetalleItem(vProSelNro As Integer)
Dim oConn As New DConecta, Rs As New ADODB.Recordset, i As Integer, nSuma As Currency
Dim sSQL As String, sGrupo As String

sSQL = ""
nSuma = 0
FormaFlexItem

If oConn.AbreConexion Then
          
    sSQL = "select v.nProSelNro, v.nProSelItem, b.cBSGrupoDescripcion, b.cBSGrupoCod,x.cBSCod, y.cBSDescripcion, x.nCantidad, v.nMonto " & _
            "from LogProSelItem v " & _
            "inner join BSGrupos b on v.cBSGrupoCod = b.cBSGrupoCod " & _
            "inner join LogProSelItemBS x on v.nProSelNro = x.nProSelNro and v.nProSelItem = x.nProSelItem " & _
            "inner join LogProSelBienesServicios y on x.cBSCod = y.cProSelBSCod " & _
            "where v.nProSelNro = " & vProSelNro & " order by v.nProSelItem, b.cBSGrupoDescripcion "

   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      MSItem.Enabled = True
      Do While Not Rs.EOF
        If sGrupo <> Rs!nProSelItem Then
         sGrupo = Rs!nProSelItem
         i = i + 1
         InsRow MSItem, i
         MSItem.row = i
         MSItem.Col = 1
         Set MSItem.CellPicture = imgNN
         MSItem.Col = 0
         MSItem.CellFontSize = 10
         MSItem.CellFontBold = True
         MSItem.TextMatrix(i, 0) = "+" '"-"
         MSItem.TextMatrix(i, 1) = Rs!nProSelItem
         MSItem.TextMatrix(i, 3) = Rs!cBSGrupoDescripcion
         MSItem.TextMatrix(i, 4) = ""
         MSItem.TextMatrix(i, 5) = Rs!nProselNro
         'MSItem.TextMatrix(i, 7) = 0
         MSItem.TextMatrix(i, 6) = Rs!nProSelItem
         MSItem.TextMatrix(i, 10) = Rs!cBSGrupoCod
         MSItem.TextMatrix(i, 8) = FNumero(Rs!nMonto)
         MSItem.TextMatrix(i, 9) = "0.00"
        End If
        i = i + 1
        InsRow MSItem, i
'        MSItem.Col = 1
'        MSItem.Row = i
'        Set MSItem.CellPicture = imgNN
        MSItem.RowHeight(i) = 0 '260
        MSItem.TextMatrix(i, 1) = Rs!cBSCod
        MSItem.TextMatrix(i, 3) = Rs!cBSDescripcion
        MSItem.TextMatrix(i, 4) = FNumero(Rs!nCantidad)
        MSItem.TextMatrix(i, 5) = Rs!nProselNro
        MSItem.TextMatrix(i, 6) = Rs!nProSelItem
        MSItem.TextMatrix(i, 7) = "0.00"
        MSItem.TextMatrix(i, 10) = Rs!cBSGrupoCod
        MSItem.TextMatrix(i, 8) = 0
        MSItem.TextMatrix(i, 9) = "0.00"
        Rs.MoveNext
      Loop
   End If
End If
End Sub

Sub GeneraDetalleItemPostor(ByVal vProSelNro As Integer, ByVal pcPersCod As String)
Dim oConn As New DConecta, Rs As New ADODB.Recordset, i As Integer, nSuma As Currency
Dim sSQL As String, sGrupo As String, nPapa As Integer

sSQL = ""
nSuma = 0

If oConn.AbreConexion Then
          
    sSQL = "select v.nProSelNro, v.nProSelItem, b.cBSGrupoDescripcion, b.cBSGrupoCod,x.cBSCod, " & _
           " Y.cBSDescripcion , X.nCantidad, v.nMonto, Z.nPropEconomica, Z.nMoneda " & _
           " from LogProSelItem v " & _
           " inner join BSGrupos b on v.cBSGrupoCod = b.cBSGrupoCod " & _
           " inner join LogProSelItemBS x on v.nProSelNro = x.nProSelNro and v.nProSelItem = x.nProSelItem " & _
           " inner join LogProSelBienesServicios y on x.cBSCod = y.cProSelBSCod " & _
           " left outer join LogProSelPostorPropuestaBS z on x.nProSelNro = z.nProSelNro and x.nProSelItem = z.nProSelItem and x.cBSCod = z.cBSCod" & _
           " where v.nProSelNro = " & vProSelNro & " and z.cPersCod = '" & pcPersCod & "' " & _
           " order by v.nProSelItem, b.cBSGrupoDescripcion "

   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      MSItem.Clear
      FormaFlexItem
      MSItem.Enabled = True
      Do While Not Rs.EOF
        If sGrupo <> Rs!nProSelItem Then
         sGrupo = Rs!nProSelItem
         i = i + 1
         InsRow MSItem, i
         MSItem.row = i
         MSItem.Col = 1
         Set MSItem.CellPicture = imgOK
         MSItem.Col = 0
         If nPapa > 0 Then MSItem.TextMatrix(nPapa, 9) = FNumero(nSuma)
         MSItem.CellFontSize = 10
         MSItem.CellFontBold = True
         MSItem.TextMatrix(i, 0) = "+" '"-"
         MSItem.TextMatrix(i, 1) = Rs!nProSelItem
         MSItem.TextMatrix(i, 3) = Rs!cBSGrupoDescripcion
         MSItem.TextMatrix(i, 4) = ""
         MSItem.TextMatrix(i, 5) = Rs!nProselNro
         MSItem.TextMatrix(i, 6) = Rs!nProSelItem
         MSItem.TextMatrix(i, 10) = Rs!cBSGrupoCod
         MSItem.TextMatrix(i, 8) = FNumero(Rs!nMonto)
         MSItem.TextMatrix(i, 9) = 0
         nPapa = i
         nSuma = 0
        End If
        i = i + 1
        InsRow MSItem, i

        MSItem.RowHeight(i) = 0 '260
        MSItem.TextMatrix(i, 1) = Rs!cBSCod
        MSItem.TextMatrix(i, 3) = Rs!cBSDescripcion
        MSItem.TextMatrix(i, 4) = FNumero(Rs!nCantidad)
        MSItem.TextMatrix(i, 5) = Rs!nProselNro
        MSItem.TextMatrix(i, 6) = Rs!nProSelItem
        MSItem.TextMatrix(i, 7) = FNumero(Rs!nPropEconomica)
        MSItem.TextMatrix(i, 8) = 0
        MSItem.TextMatrix(i, 9) = FNumero(Rs!nPropEconomica * Rs!nCantidad)
        nSuma = nSuma + CDbl(MSItem.TextMatrix(i, 9))
        Rs.MoveNext
      Loop
      If nPapa > 0 Then MSItem.TextMatrix(nPapa, 9) = FNumero(nSuma)
   Else
       GeneraDetalleItem vProSelNro
   End If
End If
End Sub


Sub FormaFlexItem()
MSItem.Clear
MSItem.Rows = 2
MSItem.RowHeight(0) = 320
MSItem.RowHeight(1) = 8
MSItem.ColWidth(0) = 250: MSItem.ColAlignment(0) = 4
MSItem.ColWidth(1) = 1000:   MSItem.ColAlignment(1) = 4:  MSItem.TextMatrix(0, 1) = " Item"
MSItem.ColWidth(2) = 0:   MSItem.ColAlignment(2) = 4:  MSItem.TextMatrix(0, 2) = " Código"
MSItem.ColWidth(3) = 4700:  MSItem.TextMatrix(0, 3) = " Descripción"
MSItem.ColWidth(4) = 1000:  MSItem.TextMatrix(0, 4) = "Cant."
MSItem.ColWidth(5) = 0:  MSItem.TextMatrix(0, 5) = " nProSelNro"
MSItem.ColWidth(6) = 0:  MSItem.TextMatrix(0, 6) = " nProSelItem"
MSItem.ColWidth(7) = 1300:  MSItem.TextMatrix(0, 7) = "P. Uni."
MSItem.ColWidth(8) = 0:  MSItem.TextMatrix(0, 8) = " Monto"
MSItem.ColWidth(9) = 1500:  MSItem.TextMatrix(0, 9) = " Precio"
MSItem.ColWidth(10) = 0:  MSItem.TextMatrix(0, 10) = " Cod Grupo"
End Sub

Private Sub MSItem_DblClick()
On Error GoTo MSItemErr
    Dim i As Integer, bTipo As Boolean
    With MSItem
        If Trim(.TextMatrix(.row, 0)) = "-" Then
           .TextMatrix(.row, 0) = "+"
           i = .row + 1
           bTipo = True
        ElseIf Trim(.TextMatrix(.row, 0)) = "+" Then
           .TextMatrix(.row, 0) = "-"
           i = .row + 1
           bTipo = False
        Else
            Exit Sub
        End If
        
        Do While i < .Rows
            If Trim(.TextMatrix(i, 0)) = "+" Or Trim(.TextMatrix(i, 0)) = "-" Then
                Exit Sub
            End If
            
            If bTipo Then
                .RowHeight(i) = 0
            Else
                .RowHeight(i) = 260
            End If
            i = i + 1
        Loop
    End With
Exit Sub
MSItemErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub MSItem_GotFocus()
If txtEditItem.Visible = False Then Exit Sub
MSItem = FNumero(txtEditItem)
txtEditItem.Visible = False
Select Case MSItem.Col
    Case 7
        GeneraPropuestaEconomica CDbl(MSItem.TextMatrix(MSItem.row, 6))
        MSItem.Col = 7
    Case 9
        GeneraPropuestaEconomicaInversa CDbl(MSItem.TextMatrix(MSItem.row, 6))
        MSItem.Col = 9
End Select
End Sub

Private Sub MSItem_LeaveCell()
If txtEditItem.Visible = False Then Exit Sub
MSItem = FNumero(txtEditItem)
txtEditItem.Visible = False
Select Case MSItem.Col
    Case 7
        GeneraPropuestaEconomica CDbl(MSItem.TextMatrix(MSItem.row, 6))
        MSItem.Col = 7
    Case 9
        GeneraPropuestaEconomicaInversa CDbl(MSItem.TextMatrix(MSItem.row, 6))
        MSItem.Col = 9
End Select
End Sub

Private Sub txtEditItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then
       KeyAscii = 0
    End If
Select Case MSItem.Col
    Case 7, 9
        KeyAscii = DigNumDec(txtEditItem, KeyAscii)
End Select
End Sub

Private Sub txtEdititem_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode MSItem, txtEditItem, KeyCode, Shift
End Sub

'Private Sub GeneraPropuestaEconomica(ByVal pnProSelItem As Integer)
'On Error GoTo GeneraPropuestaEconomicaErr
'    Dim i As Integer, nPropuesta As Currency, npapa As Integer, nRow As Integer
'    With MSItem
'        i = 1
'        Do While i < .Rows
'            If .TextMatrix(i, 0) <> "+" And .TextMatrix(i, 0) <> "-" Then
''                If .TextMatrix(i, 6) = pnProSelItem And Val(.TextMatrix(i, 7)) = 0 Then
''                    nRow = .Row
''                    .Col = 1
''                    .Row = npapa
''                    Set .CellPicture = imgNN
'''                    .TextMatrix(npapa, 7) = ""
''                    .Col = 7
''                    .Row = nRow
'''                    Exit Sub
''                Else
'            ElseIf .TextMatrix(i, 6) = pnProSelItem Then
'                    nPropuesta = nPropuesta + (Val(.TextMatrix(i, 7)) * Val(.TextMatrix(i, 4)))
'                    .TextMatrix(i, 9) = FNumero((Val(.TextMatrix(i, 7)) * Val(.TextMatrix(i, 4))))
'            End If
'            If .TextMatrix(i, 6) = pnProSelItem Then
'                npapa = i
'            End If
'            i = i + 1
'        Loop
'        nRow = .Row
'        .Col = 1
'        .Row = npapa
'        If nPropuesta > 0 Then
'            Set .CellPicture = imgOK
'        Else
'            Set .CellPicture = imgNN
'        End If
'        .TextMatrix(npapa, 9) = FNumero(nPropuesta)
'        .Col = 7
'        .Row = nRow
'    End With
'Exit Sub
'GeneraPropuestaEconomicaErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'End Sub

Private Sub GeneraPropuestaEconomica(ByVal pnProSelItem As Integer)
On Error GoTo GeneraPropuestaEconomicaErr
    Dim i As Integer, nPropuesta As Currency, nPapa As Integer, nRow As Integer
    With MSItem
        i = 1
        Do While i < .Rows
            If .TextMatrix(i, 0) <> "+" And .TextMatrix(i, 0) <> "-" Then
'                If .TextMatrix(i, 6) = pnProSelItem And Val(.TextMatrix(i, 7)) = 0 Then
'                    nRow = .Row
'                    .Col = 1
'                    .Row = npapa
'                    Set .CellPicture = imgNN
''                    .TextMatrix(npapa, 7) = ""
'                    .Col = 7
'                    .Row = nRow
''                    Exit Sub
'                Else
            If .TextMatrix(i, 6) = pnProSelItem Then
                    nPropuesta = nPropuesta + (CDbl(.TextMatrix(i, 7)) * CDbl(.TextMatrix(i, 4)))
                    .TextMatrix(i, 9) = FNumero((CDbl(.TextMatrix(i, 7)) * CDbl(.TextMatrix(i, 4))))
                End If
            ElseIf .TextMatrix(i, 6) = pnProSelItem Then
                nPapa = i
            End If
            i = i + 1
        Loop
        nRow = .row
        .Col = 1
        .row = nPapa
        If nPropuesta > 0 Then
            Set .CellPicture = imgOK
        Else
            Set .CellPicture = imgNN
        End If
        .TextMatrix(nPapa, 9) = FNumero(nPropuesta)
        .Col = 7
        .row = nRow
    End With
Exit Sub
GeneraPropuestaEconomicaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub


Private Sub GeneraPropuestaEconomicaInversa(ByVal pnProSelItem As Integer)
On Error GoTo GeneraPropuestaEconomicaErr
    Dim i As Integer, nHijos As Integer
    With MSItem
        i = .row + 1
                
        .Col = 1
        Set .CellPicture = imgOK
        
        Do While i < .Rows
            If .TextMatrix(i, 0) = "+" Or .TextMatrix(i, 0) = "-" Then
                Exit Do
            Else
                nHijos = nHijos + 1
            End If
            i = i + 1
        Loop
        
        i = .row + 1
        Do While i < .Rows
            If .TextMatrix(i, 0) = "+" Or .TextMatrix(i, 0) = "-" Then
                Exit Do
            Else
                .TextMatrix(i, 9) = FNumero(CDbl(.TextMatrix(.row, 9)) / nHijos)
                .TextMatrix(i, 7) = FNumero(CDbl(.TextMatrix(i, 9)) / .TextMatrix(i, 4))
            End If
            i = i + 1
        Loop
    End With
Exit Sub
GeneraPropuestaEconomicaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CargarPropuesta(ByVal pcPersCod As String)
    On Error GoTo CargarPropuestaErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    sSQL = "select p.nProSelNro, p.nProSelItem, p.cPersCod, p.nMoneda, p.nPropEconomica, x.cProSelBSCod, nPropEconomicaBS=x.nPropEconomica " & _
            " from LogProSelPostorPropuesta p " & _
            " inner join LogProSelPostorPropuestaBS x on " & _
            " P.nProSelNro = X.nProSelNro And P.nProSelItem = X.nProSelItem "
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            
            Rs.MoveNext
        Loop
    End If
    Exit Sub
CargarPropuestaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Function ValidaBS(pnFila As Integer) As Boolean
    On Error GoTo ValidaBSErr
    Dim i As Integer
    i = pnFila + 1
    With MSItem
        Do While i < .Rows
            If .TextMatrix(i, 0) = "+" Or .TextMatrix(i, 0) = "-" Then
                ValidaBS = True
                Exit Function
            End If
            If Val(.TextMatrix(i, 7)) = 0 Then
                ValidaBS = False
                Exit Function
            End If
            i = i + 1
        Loop
        ValidaBS = True
    End With
    Exit Function
ValidaBSErr:
    MsgBox Err.Number & Err.Description, vbInformation
End Function

Private Function ValidaPropuestaEconomica(ByVal pnMonto As Currency, ByVal pnPropuesta As Currency, _
            ByVal pcPersCod As String, ByVal pnProSelNro As Integer, ByVal pnProSelItem As Integer) As Boolean
On Error GoTo ValidaPropuestaEconomicaErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, max As Currency, min As Currency
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        ValidaPropuestaEconomica = False
        
        sSQL = "select nRangoMayor, nRangoMenor from LogProcesoSeleccion where nProSelNro=" & pnProSelNro
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            max = Rs!nRangoMayor / 100#
            min = Rs!nRangoMenor / 100#
        End If
        Set Rs = Nothing
        
        If (pnMonto * min) > pnPropuesta Or (pnMonto * max) < pnPropuesta Then
            sSQL = "update LogProSelPostorPropuesta set bDesestimado=1, cDesesDescripcion='Propuesta Economica Fuera del Intervalo Permitido' " & _
                    " where nProSelItem=" & pnProSelItem & " and nProSelNro=" & pnProSelNro & " and cPersCod='" & pcPersCod & "'"
            oCon.Ejecutar sSQL
            ValidaPropuestaEconomica = True
        End If
        oCon.CierraConexion
    End If
    Exit Function
ValidaPropuestaEconomicaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Function ValidarPropuestaPostor(ByVal pnProSelNro As Integer, ByVal pcPersCod As String) As Boolean
    On Error GoTo ValidaPropuestaPostorErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    sSQL = "select * from LogProSelPostorPropuesta where nProSelNro=" & pnProSelNro & " and cPersCod='" & pcPersCod & "'"
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            ValidarPropuestaPostor = True
        End If
        oCon.CierraConexion
    End If
    Set Rs = Nothing
    Exit Function
ValidaPropuestaPostorErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Function VerificaModalidad(ByVal pnModalidad As Boolean) As Boolean
On Error GoTo VerificaModalidadErr
    Dim i As Integer
    With MSItem
        If pnModalidad Then
            Do While i < .Rows
                .row = i: .Col = 1
                If (.TextMatrix(0, i) = "+" Or .TextMatrix(0, i) = "-") And .CellPicture = imgNN Then
                    VerificaModalidad = False
                    Exit Function
                End If
                i = i + 1
            Loop
            VerificaModalidad = True
        End If
    End With
    Exit Function
VerificaModalidadErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Sub cmdGrabarPos_Click()
Dim oConn As New DConecta, sSQL As String, i As Integer, nItem As Integer, _
    nMoneda As Integer, nDesestimado As Integer, cDesestimado As String
On Error GoTo GrabarPos

If gnProSelNro = 0 Then Exit Sub

nMoneda = IIf(LblMoneda.Caption = "S/.", 1, 2)

'sSQL = " insert into LogProSelPostorPropuesta (nProSelNro,nProSelItem,cPersCod,nPropEconomica,nPuntaje,bGanador) " & _
'       " values (" & nProSelNro & "," & nProSelItem & ",'" & txtPersCod & "'," & VNumero(txtPropEcon) & "," & VNumero(txtPuntaje) & "," & chkGanador.Value & ")"

i = 1
With MSItem
    Do While i < .Rows
        If Val(Trim(.TextMatrix(i, 7))) > 0 Then
            Exit Do
        End If
        i = i + 1
    Loop
    If i = .Rows Then
        MsgBox "Debe Ingrear una Propuesta Valida....", vbInformation
        Exit Sub
    End If
End With

If ValidarPropuestaPostor(gnProSelNro, Right(CboPostores.Text, 13)) Then
    If MsgBox("Propuesta ya Fue Registrada..." & vbCrLf & "Seguro que Desea Cambiarla", vbInformation + vbYesNo) = vbNo Then
        Exit Sub
    Else
        If oConn.AbreConexion Then
            'oConn.BeginTrans
            sSQL = "update LogProSelEvalResultado set nPuntaje=0 where nProSelNro = " & gnProSelNro & " and cPersCod= '" & Right(CboPostores.Text, 13) & "'"
            oConn.Ejecutar sSQL
'            sSQL = "delete LogProSelPostorPropuestaBS where nProSelNro = " & gnProSelNro & " and cPersCod= '" & Right(CboPostores.Text, 13) & "'"
'            oConn.Ejecutar sSQL
'            sSQL = "delete LogProSelPostorPropuesta where nProSelNro = " & gnProSelNro & " and cPersCod= '" & Right(CboPostores.Text, 13) & "'"
'            oConn.Ejecutar sSQL
'            oConn.CommitTrans
            oConn.CierraConexion
            sSQL = ""
        End If
    End If
End If

'If MsgBox("¿ Está seguro de agregar la propuesta del Postor ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
    nItem = 1
    If oConn.AbreConexion Then
     Do While nItem < MSItem.Rows
        MSItem.row = nItem
        MSItem.Col = 1
        If MSItem.CellPicture = imgOK And Val(MSItem.TextMatrix(nItem, 9)) > 0 Then
'            If Val(MSItem.TextMatrix(nItem, 10)) > 0 Then
'                cboGrupoBS.AddItem MSItem.TextMatrix(nItem, 3) & Space(150) & MSItem.TextMatrix(nItem, 10), cboGrupoBS.ListCount
'                cboGrupoBS.ListIndex = 0
'            End If
            If Not ValidaBS(nItem) Then
                nDesestimado = 1
                cDesestimado = "no Cubrir todos los Items"
            ElseIf ValidaPropuestaEconomica(CDbl(MSItem.TextMatrix(nItem, 8)), CDbl(MSItem.TextMatrix(nItem, 9)), Right(CboPostores.Text, 13), gnProSelNro, Val(MSItem.TextMatrix(nItem, 6))) Then
                nDesestimado = 1
                cDesestimado = "Propuesta Economica Fuera de Intervalo"
'            ElseIf Not VerificaModalidad(gnModalidad) Then
'                nDesestimado = 1
'                cDesestimado = "no Cubrir Todo el Proceso"
            Else
                nDesestimado = 0
                cDesestimado = ""
            End If
            
            sSQL = "declare @tmp int "
            sSQL = sSQL & " set @tmp = (select count(*) from LogProSelPostorPropuesta where nProSelNro = " & Val(MSItem.TextMatrix(nItem, 5)) & " and nProSelItem = " & Val(MSItem.TextMatrix(nItem, 6)) & " and cPersCod = '" & Right(CboPostores.Text, 13) & "') "
            sSQL = sSQL & " if @tmp=0 "
            sSQL = sSQL & " insert into LogProSelPostorPropuesta (nProSelNro,nProSelItem,cPersCod,nPropEconomica,nMoneda,bDesestimado,cDesesDescripcion) " & _
                    " values (" & Val(MSItem.TextMatrix(nItem, 5)) & "," & Val(MSItem.TextMatrix(nItem, 6)) & ",'" & Right(CboPostores.Text, 13) & "'," & VNumero(MSItem.TextMatrix(nItem, 9)) & "," & nMoneda & "," & nDesestimado & ",'" & cDesestimado & "') "
            sSQL = sSQL & " else "
            sSQL = sSQL & " update LogProSelPostorPropuesta set nPropEconomica = " & VNumero(MSItem.TextMatrix(nItem, 9)) & ", " & _
                          " nMoneda=" & nMoneda & ", " & _
                          " bDesestimado = " & nDesestimado & ", " & _
                          " cDesesDescripcion = " & "'" & cDesestimado & "' " & _
                          " where nProSelNro = " & Val(MSItem.TextMatrix(nItem, 5)) & " and nProSelItem = " & Val(MSItem.TextMatrix(nItem, 6)) & " and cPersCod = '" & Right(CboPostores.Text, 13) & "' "
            oConn.Ejecutar sSQL
            
        ElseIf Len(MSItem.TextMatrix(nItem, 1)) > 1 And Val(MSItem.TextMatrix(nItem, 7)) > 0 Then
            sSQL = "declare @tmp int "
            sSQL = sSQL & " set @tmp = (select count(*) from LogProSelPostorPropuestaBS where nProSelNro = " & Val(MSItem.TextMatrix(nItem, 5)) & " and nProSelItem = " & Val(MSItem.TextMatrix(nItem, 6)) & " and cPersCod = '" & Right(CboPostores.Text, 13) & "' and cBSCod='" & MSItem.TextMatrix(nItem, 1) & "') "
            sSQL = sSQL & " if @tmp=0 "
            sSQL = sSQL & " insert into LogProSelPostorPropuestaBS (nProSelNro,nProSelItem,cBSCod,cPersCod,nPropEconomica,nMoneda) " & _
                    " values (" & Val(MSItem.TextMatrix(nItem, 5)) & "," & Val(MSItem.TextMatrix(nItem, 6)) & ",'" & MSItem.TextMatrix(nItem, 1) & "','" & Right(CboPostores.Text, 13) & "'," & VNumero(MSItem.TextMatrix(nItem, 7)) & "," & nMoneda & " ) "
            sSQL = sSQL & " else "
            sSQL = sSQL & " update LogProSelPostorPropuestaBS set nMoneda = " & VNumero(MSItem.TextMatrix(nItem, 9)) & ", " & _
                          " nPropEconomica = " & VNumero(MSItem.TextMatrix(nItem, 7)) & " " & _
                          " where nProSelNro = " & Val(MSItem.TextMatrix(nItem, 5)) & " and nProSelItem = " & Val(MSItem.TextMatrix(nItem, 6)) & " and cPersCod = '" & Right(CboPostores.Text, 13) & "' and cBSCod='" & MSItem.TextMatrix(nItem, 1) & "' "
            oConn.Ejecutar sSQL
        End If
        
        RegistrarFactorEconomico gnProSelTpoCod, gnProSelSubTpo
        
         nItem = nItem + 1
     Loop
     
     oConn.CierraConexion
    End If
        
    
    MSItem.SetFocus
    
    MsgBox "Propuesta Registrada Correctamente...", vbInformation
    GeneraDetalleItem gnProSelNro
    CargarValores Right(CboPostores.Text, 13), Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
    
'    frmEvaluacion.Visible = True
'    FramePropuestaEconomica.Visible = False
'    FormaFlexFactores
'    CargarFactores gnProSelTpoCod, gnProSelSubTpo, Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
'    CargarValores txtPersCod.Text, Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
'    cmdCancelaPos_Click
'End If
Exit Sub
GrabarPos:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub RegistrarFactorEconomico(ByVal pnProSelTpoCod As Integer, pnProSelSubTpo As Integer)
On Error GoTo RegistrarFactorEconomicoErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, nFactor As Integer, sGrupo As String, nItem As Integer, cPersCod  As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        nItem = 1
        cPersCod = Right(CboPostores.Text, 13)
        Do While nItem < MSItem.Rows
            MSItem.row = nItem
            sGrupo = MSItem.TextMatrix(nItem, 10)
            sSQL = "select x.nFactorNro, x.cFactorDescripcion, f.nFormula, x.cUnidades, x.nTipo from LogProSelEvalFactor f " & _
                "inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
                "where nTipo=1 and nVigente=1 and nProSelTpoCod=" & pnProSelTpoCod & " and nProSelSubTpo=" & pnProSelSubTpo & " and cBSGrupoCod='" & sGrupo & "' and nProSelNro = " & gnProSelNro
            Set Rs = oCon.CargaRecordSet(sSQL)
            If Not Rs.EOF Then
                nFactor = Rs!nFactorNro
                Rs.Close
            Else
                Rs.Close
                Exit Sub
            End If
            
            If MSItem.CellPicture = imgOK And Val(MSItem.TextMatrix(nItem, 9)) > 0 Then
                'sSQL = " insert into LogProSelPostorPropuesta (nProSelNro,nProSelItem,cPersCod,nPropEconomica,nMoneda,bDesestimado,cDesesDescripcion) " & _
                '        " values (" & Val(MSItem.TextMatrix(nItem, 5)) & "," & Val(MSItem.TextMatrix(nItem, 6)) & ",'" & Right(CboPostores.Text, 13) & "'," & VNumero(MSItem.TextMatrix(nItem, 9)) & "," & nMoneda & "," & nDesestimado & ",'" & cDesestimado & "')"
                'oConn.Ejecutar sSQL
                sSQL = "declare @tmp int "
                sSQL = sSQL & " set @tmp=(select count(*) from LogProSelEvalResultado where nProSelNro = " & Val(MSItem.TextMatrix(nItem, 5)) & " and cPersCod = '" & cPersCod & "' and nProSelItem = " & Val(MSItem.TextMatrix(nItem, 6)) & " and nFactorNro =" & nFactor & ") "
                sSQL = sSQL & " if @tmp=0 "
                sSQL = sSQL & " insert into LogProSelEvalResultado(nProSelNro,cPersCod,nProSelItem,nFactorNro,nValor,cMovNro,cBSGrupoCod) " & _
                        " values(" & Val(MSItem.TextMatrix(nItem, 5)) & ",'" & cPersCod & "'," & Val(MSItem.TextMatrix(nItem, 6)) & "," & nFactor & "," & VNumero(MSItem.TextMatrix(nItem, 9)) & ",'" & GetLogMovNro & "','" & sGrupo & "') "
                sSQL = sSQL & " else "
                sSQL = sSQL & " update LogProSelEvalResultado set " & _
                              " nValor = " & VNumero(MSItem.TextMatrix(nItem, 9)) & " ,cMovNro = '" & GetLogMovNro & "' ,cBSGrupoCod = '" & sGrupo & "' " & _
                              " where nProSelNro = " & Val(MSItem.TextMatrix(nItem, 5)) & " and cPersCod = '" & cPersCod & "' and nProSelItem = " & Val(MSItem.TextMatrix(nItem, 6)) & " and nFactorNro =" & nFactor
                oCon.Ejecutar sSQL
            End If
             nItem = nItem + 1
         Loop
        
        oCon.CierraConexion
    End If
    Exit Sub
RegistrarFactorEconomicoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub
        

