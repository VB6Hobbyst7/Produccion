VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmLogSelComienzo 
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   Icon            =   "frmLogSelComienzo.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   11250
   Begin VB.ComboBox cmbclaseProceso 
      Height          =   315
      ItemData        =   "frmLogSelComienzo.frx":030A
      Left            =   1080
      List            =   "frmLogSelComienzo.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   390
      Left            =   7980
      TabIndex        =   39
      Top             =   4815
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   3
      Left            =   6675
      TabIndex        =   11
      Top             =   4800
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   5370
      TabIndex        =   12
      Top             =   4800
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Editar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   4065
      TabIndex        =   14
      Top             =   4800
      Width           =   1305
   End
   Begin VB.ComboBox cboperiodo 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      Height          =   390
      Index           =   0
      Left            =   2760
      TabIndex        =   13
      Top             =   4800
      Width           =   1305
   End
   Begin TabDlg.SSTab SSTSeleccion 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6165
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Proceso de Seleccion"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label15"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label16"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DTPFechaConvo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbtipoproceso"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmbmoneda"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmbTipoSeleccion"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtrangoini"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtrangofin"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtdescripcion"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtnumerocotizacion"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Config. Puntajes Evaluacion"
      TabPicture(1)   =   "frmLogSelComienzo.frx":03D2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label14"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "uSTecPunMinimo"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "CkEvaluacionTecnica"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "CkEvaluacionEcon"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "MEEcoPeso"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "METecPeso"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdPunt(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmdPunt(2)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdPunt(0)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "uSEconPunMinimo"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "uSTecPunMaximo"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "uSEcoPunMaximo"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Comite de Evaluacion"
      TabPicture(2)   =   "frmLogSelComienzo.frx":03EE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FlexComite"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdNuevoComite"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmbtipocomite"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdcomitepreter"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdEliminarComite"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.TextBox txtnumerocotizacion 
         Height          =   285
         Left            =   6600
         MaxLength       =   20
         TabIndex        =   49
         Top             =   720
         Width           =   3375
      End
      Begin VB.CommandButton cmdEliminarComite 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   -65400
         TabIndex        =   32
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdcomitepreter 
         Caption         =   "Agregar Comite"
         Height          =   375
         Left            =   -72240
         TabIndex        =   47
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox cmbtipocomite 
         Height          =   315
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   2880
         Width           =   2535
      End
      Begin Spinner.uSpinner uSEcoPunMaximo 
         Height          =   255
         Left            =   -68280
         TabIndex        =   45
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin Spinner.uSpinner uSTecPunMaximo 
         Height          =   255
         Left            =   -68280
         TabIndex        =   44
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin Spinner.uSpinner uSEconPunMinimo 
         Height          =   255
         Left            =   -70800
         TabIndex        =   43
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.TextBox txtdescripcion 
         Height          =   495
         Left            =   2040
         MaxLength       =   130
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   2640
         Width           =   8175
      End
      Begin VB.CommandButton cmdPunt 
         Caption         =   "Editar"
         Height          =   375
         Index           =   0
         Left            =   -68040
         TabIndex        =   36
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmdPunt 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   2
         Left            =   -65400
         TabIndex        =   35
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdPunt 
         Caption         =   "Grabar"
         Height          =   375
         Index           =   1
         Left            =   -66720
         TabIndex        =   34
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmdNuevoComite 
         Caption         =   "N&uevo"
         Height          =   375
         Left            =   -66600
         TabIndex        =   33
         Top             =   2880
         Width           =   1215
      End
      Begin MSMask.MaskEdBox METecPeso 
         Height          =   285
         Left            =   -65040
         TabIndex        =   30
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   4
         Format          =   "#.#0"
         Mask            =   "#.#0"
         PromptChar      =   "0"
      End
      Begin MSMask.MaskEdBox MEEcoPeso 
         Height          =   285
         Left            =   -65040
         TabIndex        =   29
         Top             =   1560
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   4
         Format          =   "#.#0"
         Mask            =   "#.#0"
         PromptChar      =   "#"
      End
      Begin VB.CheckBox CkEvaluacionEcon 
         Caption         =   "Evaluacion Economica"
         Height          =   495
         Left            =   -74640
         TabIndex        =   22
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox CkEvaluacionTecnica 
         Caption         =   "Evaluacion Tecnica"
         Height          =   495
         Left            =   -74640
         TabIndex        =   21
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtrangofin 
         Height          =   285
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1215
         Width           =   1095
      End
      Begin VB.TextBox txtrangoini 
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1215
         Width           =   1095
      End
      Begin VB.ComboBox cmbTipoSeleccion 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   3135
      End
      Begin VB.ComboBox cmbmoneda 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
      End
      Begin VB.ComboBox cmbtipoproceso 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1200
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker DTPFechaConvo 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60030977
         CurrentDate     =   37973
      End
      Begin Sicmact.FlexEdit FlexComite 
         Height          =   2400
         Left            =   -74880
         TabIndex        =   31
         Top             =   360
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   4233
         Cols0           =   4
         HighLight       =   1
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre-Cargo"
         EncabezadosAnchos=   "350-1800-4000-3000"
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
         ColumnasAEditar =   "X-1-X-3"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-3"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   240
         CellBackColor   =   -2147483624
      End
      Begin Spinner.uSpinner uSTecPunMinimo 
         Height          =   255
         Left            =   -70800
         TabIndex        =   42
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Min             =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Nº Cotizacion"
         Height          =   195
         Left            =   5400
         TabIndex        =   48
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label Label14 
         Caption         =   "Peso Ponderado de 0 a 1"
         Height          =   255
         Left            =   -67200
         TabIndex        =   28
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "Peso Ponderado de 0 a 1"
         Height          =   255
         Left            =   -67200
         TabIndex        =   27
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Puntaje Maximo"
         Height          =   255
         Left            =   -69600
         TabIndex        =   26
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Puntaje Minimo"
         Height          =   255
         Left            =   -72120
         TabIndex        =   25
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Puntaje Maximo"
         Height          =   255
         Left            =   -69600
         TabIndex        =   24
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Puntaje Minimo"
         Height          =   255
         Left            =   -72120
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Rango Final"
         Height          =   195
         Left            =   7800
         TabIndex        =   16
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Rango Inicial"
         Height          =   195
         Left            =   5400
         TabIndex        =   15
         Top             =   1260
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Bienes"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Proceso Seleccion"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Moneda "
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Convocatoria"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   1665
      End
   End
   Begin Sicmact.TxtBuscar txtSeleccion 
      Height          =   315
      Left            =   4320
      TabIndex        =   9
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
   Begin VB.Label cmdtipoproceso 
      AutoSize        =   -1  'True
      Caption         =   "Proceso"
      Height          =   195
      Left            =   240
      TabIndex        =   51
      Top             =   720
      Width           =   585
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
      Left            =   6360
      TabIndex        =   41
      Top             =   240
      Width           =   1350
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
      Left            =   8280
      TabIndex        =   40
      Top             =   270
      Width           =   660
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BorderStyle     =   4  'Dash-Dot
      FillColor       =   &H8000000D&
      Height          =   495
      Left            =   8160
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Periodo"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   300
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Proceso"
      Height          =   195
      Left            =   3360
      TabIndex        =   10
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "frmLogSelComienzo"
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
Dim bpuntaje As Boolean
Public Sub Inicio(ByVal psTipoReq As String, ByVal psFormTpo As String, Optional ByVal psRequeriNro As String = "")
psTpoReq = psTipoReq
psFrmTpo = psFormTpo
psReqNro = psRequeriNro
Me.Show 1
End Sub
Private Sub cboPeriodo_Click()
txtSeleccion.Text = ""
FlexComite.Clear
FlexComite.FormaCabecera
FlexComite.Rows = 2
cmbTipoSeleccion.Clear
cmbtipoproceso.Clear
txtrangoini.Text = ""
txtrangofin.Text = ""
DTPFechaConvo.value = Now
txtDescripcion.Text = ""
Set rs = clsDGAdqui.CargaLogSelTipoBs
Call CargaCombo(rs, cmbTipoSeleccion)
puntaje
'If saccion = "N" Then Exit Sub
'If saccion = "E" Then Exit Sub
 Me.txtSeleccion.rs = clsDGAdqui.LogSeleccionLista(cboperiodo.Text)
 If Me.txtSeleccion.rs.EOF = True Then
    Me.txtSeleccion.Text = ""
 End If
 txtSeleccion.Enabled = True
 
End Sub
Private Sub CkEvaluacionEcon_Click()

If CkEvaluacionEcon.value = 0 Then
    uSEconPunMinimo.Enabled = False
    uSEcoPunMaximo.Enabled = False
    MEEcoPeso.Enabled = False
    uSEconPunMinimo.BackColor = 14737632
    uSEcoPunMaximo.BackColor = 14737632
    MEEcoPeso.BackColor = 14737632
ElseIf CkEvaluacionEcon.value = 1 Then
    uSEconPunMinimo.Enabled = True
    uSEcoPunMaximo.Enabled = True
    MEEcoPeso.Enabled = True
    uSEconPunMinimo.BackColor = -2147483643
    uSEcoPunMaximo.BackColor = -2147483643
    MEEcoPeso.BackColor = -2147483643
End If
End Sub

Private Sub CkEvaluacionTecnica_Click()




If CkEvaluacionTecnica.value = 1 Then
    uSTecPunMinimo.Enabled = True
    uSTecPunMaximo.Enabled = True
    METecPeso.Enabled = True
    uSTecPunMinimo.BackColor = -2147483643
    uSTecPunMaximo.BackColor = -2147483643
    METecPeso.BackColor = -2147483643
ElseIf CkEvaluacionTecnica.value = 0 Then
    uSTecPunMinimo.Enabled = False
    uSTecPunMaximo.Enabled = False
    METecPeso.Enabled = False
    uSTecPunMinimo.BackColor = 14737632
    uSTecPunMaximo.BackColor = 14737632
    METecPeso.BackColor = 14737632
End If
End Sub









Private Sub cmbtipoproceso_Click()
Set rs = clsDGAdqui.LogSelConfProceso(Right(cmbTipoSeleccion.Text, 1), Right(cmbtipoproceso.Text, 1))
If rs.RecordCount > 0 Then
    txtrangoini.Text = Format(rs!nRangoIni, "########0.#0")
    txtrangofin.Text = Format(rs!nRangoFin, "########0.#0")
End If

End Sub

Private Sub cmbTipoSeleccion_Click()
Dim sTipoSelect As String
cmbtipoproceso.Clear
sTipoSelect = Right(cmbTipoSeleccion.Text, 1)

Set rs = clsDGAdqui.LogSelTipoSeleccion(sTipoSelect)
Call CargaCombo(rs, cmbtipoproceso)
cmbtipoproceso.ListIndex = 0
Set rs = clsDGAdqui.LogSelConfProceso(Right(cmbTipoSeleccion.Text, 1), Right(cmbtipoproceso.Text, 1))
If rs.RecordCount > 0 Then
    txtrangoini.Text = Format(rs!nRangoIni, "########0.#0")
    txtrangofin.Text = Format(rs!nRangoFin, "########0.#0")
End If

'Establecer  los puntajes ponderados
Select Case sTipoSelect
Case "1" 'Bienes Suministros
         METecPeso.Mask = "#.00"
         MEEcoPeso.Mask = "#.00"
         METecPeso.PromptInclude = True
         MEEcoPeso.PromptInclude = True
         METecPeso.Text = "1.00"
         MEEcoPeso.Text = "1.00"
         uSTecPunMaximo.Valor = 50
         uSEcoPunMaximo.Valor = 50
Case "2", "3"
         METecPeso.PromptChar = 0
         MEEcoPeso.PromptChar = 0
         METecPeso.PromptInclude = True
         MEEcoPeso.PromptInclude = True
         METecPeso.Mask = "0.##"
         MEEcoPeso.Mask = "0.##"
         METecPeso.Text = "0.50"
         MEEcoPeso.Text = "0.50"
         uSTecPunMaximo.Valor = 100
         uSEcoPunMaximo.Valor = 100

End Select
End Sub

Private Sub cmdcomitepreter_Click()
If cmbtipoComite.Text = "" Then
    MsgBox "Seleccione Un Tipo de Comite ", vbInformation, "Seleccione Comite"
    Exit Sub
End If

If MsgBox("¿Se Agregarn Todos Los Miembros  que conforman el " & Left(cmbtipoComite.Text, 30) & " ? ", vbQuestion + vbYesNo, " Desea Agregar  ") = vbYes Then
    mostrar_Comite_Tipo Right(cmbtipoComite.Text, 1)
            
End If




End Sub

Private Sub cmdEliminarComite_Click()
    Me.FlexComite.EliminaFila Me.FlexComite.Row
End Sub


Private Sub cmdNuevoComite_Click()
    Dim oPersona As UPersona
    If Me.FlexComite.TextMatrix(Me.FlexComite.Rows - 1, 0) = "#" Then
        FlexComite.Rows = 2
    End If
    
    If Me.FlexComite.TextMatrix(Me.FlexComite.Rows - 1, 0) = "" Then
        FlexComite.AdicionaFila 1
    Else
        FlexComite.AdicionaFila CLng(Me.FlexComite.TextMatrix(FlexComite.Rows - 1, 0)) + 1
    End If
    FlexComite.SetFocus

End Sub


Private Sub cmdPunt_Click(Index As Integer)
Dim btec As Boolean
Dim beco As Boolean
Dim ncodigo As Long
Dim rs_sel As ADODB.Recordset

Dim sactualiza As String
Set rs_sel = New ADODB.Recordset
If txtSeleccion.Text = "" Then
    MsgBox "Debe Seleccionar Un codigo de Procesos de Seleccion valido", vbInformation, "Seleccione Codigo Valido"
    Exit Sub
End If
ncodigo = txtSeleccion.Text

sactualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)

Select Case Index
    Case 0 'editar
                cmdPunt(0).Enabled = False
                cmdPunt(1).Enabled = True
                cmdPunt(2).Enabled = True
                CkEvaluacionTecnica.Enabled = True
            If CkEvaluacionTecnica.value = 1 Then
                uSTecPunMinimo.BackColor = -2147483643
                uSTecPunMaximo.BackColor = -2147483643
                METecPeso.BackColor = -2147483643
               Else
                uSEconPunMinimo.BackColor = 14737632
                uSEcoPunMaximo.BackColor = 14737632
                MEEcoPeso.BackColor = 14737632
            End If
            If CkEvaluacionEcon.value = 1 Then
                uSEconPunMinimo.BackColor = -2147483643
                uSEcoPunMaximo.BackColor = -2147483643
                MEEcoPeso.BackColor = -2147483643
            Else
                uSTecPunMinimo.BackColor = 14737632
                uSTecPunMaximo.BackColor = 14737632
                METecPeso.BackColor = 14737632
            End If
                uSTecPunMinimo.Enabled = True
                uSTecPunMaximo.Enabled = True
                METecPeso.Enabled = True
                CkEvaluacionEcon.Enabled = True
                uSEconPunMinimo.Enabled = True
                uSEcoPunMaximo.Enabled = True
                MEEcoPeso.Enabled = True
                bpuntaje = True
    Case 1
                
                If valida_ponderado = -1 Then
                   Exit Sub
                End If
                cmdPunt(0).Enabled = True
                cmdPunt(1).Enabled = False
                cmdPunt(2).Enabled = False
                ClsNAdqui.EliminaSeleccionConfigPuntaje ncodigo
                'Inserto Nueva Configuracion
                If CkEvaluacionTecnica.value = 1 Then
                   ClsNAdqui.AgregaSeleccionConfigPuntaje ncodigo, SelEvalTecnica, uSTecPunMinimo.Valor, uSTecPunMaximo.Valor, METecPeso.Text, sactualiza
                   End If
                   If CkEvaluacionEcon.value = 1 Then
                   ClsNAdqui.AgregaSeleccionConfigPuntaje ncodigo, SelEvalEconomica, uSEconPunMinimo.Valor, uSEcoPunMaximo.Valor, MEEcoPeso.Text, sactualiza
                   End If
                bpuntaje = False
    Case 2
                            cmdPunt(0).Enabled = True
                            cmdPunt(1).Enabled = False
                            cmdPunt(2).Enabled = False
                            'Cancela Recupera Configuracion Original
                            Set rs_sel = clsDGAdqui.CargaSeleccionPuntajes(ncodigo)
                            If rs_sel.EOF = True Then
                                MsgBox "No se Pudo cargar Informacion sobre los Puntajes ,Consulte con Sistemas", vbInformation, "No se pudo Cargar"
                            End If
                            Do While rs_sel.EOF = False
                            If rs_sel!cCodTipoEvaluacion = SelEvalTecnica Then
                                    CkEvaluacionTecnica.value = 1
                                    uSTecPunMinimo.Valor = rs_sel!nPuntajeMinimo
                                    uSTecPunMaximo.Valor = rs_sel!nPuntajeMaximo
                                    METecPeso.Text = Format(rs_sel!nPesoPonderado, "00.00")
                                    btec = True
                            End If
                            If rs_sel!cCodTipoEvaluacion = SelEvalEconomica Then
                                    CkEvaluacionEcon.value = 1
                                    uSEconPunMinimo.Valor = rs_sel!nPuntajeMinimo
                                    uSEcoPunMaximo.Valor = rs_sel!nPuntajeMaximo
                                    MEEcoPeso.Text = Format(rs_sel!nPesoPonderado, "00.00")
                                    beco = True
                            End If
                                    rs_sel.MoveNext
                                    Loop
                                    bpuntaje = False
End Select

End Sub

Private Sub cmdReq_Click(Index As Integer)
Dim ncodigo As Long
Dim sactualiza As String
Dim nestadoProc As Integer
Dim oConec As DConecta
Set oConec = New DConecta
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Dim i As Integer
Select Case Index
Case 0 'Nuevo
    saccion = "N"
    cmdReq(0).Enabled = False
    cmdReq(1).Enabled = False
    cmdReq(2).Enabled = True
    cmdReq(3).Enabled = True
    txtSeleccion.Text = ""
    txtSeleccion.Enabled = False
    txtDescripcion.Text = ""
    cmbTipoSeleccion.Clear
    cmbtipoproceso.Clear
    Set rs = clsDGAdqui.CargaLogSelTipoBs
    Call CargaCombo(rs, cmbTipoSeleccion)
    txtrangoini.Text = ""
    txtrangofin.Text = ""
    DTPFechaConvo.value = Now
    FlexComite.Clear
    FlexComite.FormaCabecera
    FlexComite.Rows = 2
    CkEvaluacionTecnica.value = 0
    CkEvaluacionEcon.value = 0
    Bloquear_Controles "N"
Case 1 'Editar
    If txtSeleccion.Text = "" Then
       MsgBox "Seleccione un Codigo de Seleccion Valido", vbInformation, "Seleccione Un Codigo de Seleccion"
       Exit Sub
    End If
    
    nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccion.Text)
    If nestadoProc <> SelEstProcesoIniciado Then
       MsgBox "No se puede Modificar,El Procesos de Seleccion " + txtSeleccion.Text + " Tiene un estado diferente al de  INICIADO", vbInformation, "Estado del proceso" + txtSeleccion.Text + " es diferente a INICIADO "
       Exit Sub
    End If
    Bloquear_Controles "E"
    'bpuntaje = True
    saccion = "E"
Case 2 'Cancelar
    Select Case saccion
    Case "N"
        Limpiar_Controles
        Bloquear_Controles "M"
    Case "E"
        mostrar_seleccion Trim(txtSeleccion.Text)
        Bloquear_Controles "M"
    End Select
    txtSeleccion.Enabled = True
    cmdReq(0).Enabled = True
    cmdReq(1).Enabled = True
    cmdReq(2).Enabled = False
    cmdReq(3).Enabled = False
    cmdPunt(0).Visible = False
    cmdPunt(1).Visible = False
    cmdPunt(2).Visible = False
    cmdPunt(0).Enabled = False
    cmdPunt(1).Enabled = False
    cmdPunt(2).Enabled = False
    saccion = "C"
Case 3 'Grabar
    sTipoBS = Right(Trim(cmbTipoSeleccion.Text), 1)

    If cmbclaseProceso.Text = "" Then
        MsgBox "Debe ingresar el Tipo de Proceso  ", vbInformation, "Seleccione un tipo de Proceso"
        cboperiodo.SetFocus
        Exit Sub
    End If
    
    If cboperiodo.Text = "" Then
       MsgBox "Debe Seleccionar un Periodo Valido ", vbInformation, "Seleccione un Periodo valido"
       cboperiodo.SetFocus
       Exit Sub
    End If
    If cmbTipoSeleccion.Text = "" Then
           MsgBox "Debe Seleccionar un Tipo de Seleccion Valido ", vbInformation, "Seleccione un Tipo de Seleccion Valido"
        SSTSeleccion.Tab = 0
        cmbTipoSeleccion.SetFocus
           Exit Sub
    End If
    If cmbtipoproceso.Text = "" Then
           MsgBox "Debe Seleccionar un Tipo de Proceso Valido ", vbInformation, "Seleccione un Tipo de Proceso Valido"
        SSTSeleccion.Tab = 0
        cmbtipoproceso.SetFocus
        Exit Sub
    End If
          
    If txtnumerocotizacion.Text = "" Then
        MsgBox "Debe Ingresar un Numero de Cotizacion  ", vbInformation, "Ingrese Un Numero de Cotizacion para el Proceso"
        SSTSeleccion.Tab = 0
        txtnumerocotizacion.SetFocus
        Exit Sub
    End If
      
    If CkEvaluacionEcon.value = 0 Or CkEvaluacionTecnica.value = 0 Then
        MsgBox "Olvido Configurar los Puntajes ", vbInformation, "Configure los puntajes por favor "
        SSTSeleccion.Tab = 1
        Exit Sub
    End If
    
    If bpuntaje = True Then
        MsgBox "Antes Debe Grabar los Puntajes ", vbInformation, "Falta Grabar o Cancelar Los Puntajes"
        SSTSeleccion.Tab = 1
        Exit Sub
    End If
    
     If valida_ponderado = -1 Then
        Exit Sub
     End If
    
    If FlexComite.Rows <= 2 And FlexComite.TextMatrix(1, 1) = "" Then
        MsgBox "Debe Ingresar  Los Miembros del Comite ", vbInformation, "Ingrese los Integrantes del Comite"
        SSTSeleccion.Tab = 2
        FlexComite.SetFocus
        Exit Sub
    End If
    For i = 0 To FlexComite.Rows - 1
        If FlexComite.TextMatrix(i, 3) = "" Then
            SSTSeleccion.Tab = 2
            MsgBox "Falta Asignar Cargo  a Los Miembros del Comite en el Item  Nº" & i, vbInformation, "Ingrese los Integrantes del Comite"
            FlexComite.SetFocus
            Exit Sub
        End If
    Next
    
    
    txtSeleccion.Enabled = True
    cmdReq(0).Enabled = True
    cmdReq(1).Enabled = True
    cmdReq(2).Enabled = False
    cmdReq(3).Enabled = False
    cmdPunt(0).Visible = False
    cmdPunt(1).Visible = False
    cmdPunt(2).Visible = False
    cmdPunt(0).Enabled = False
    cmdPunt(1).Enabled = False
    cmdPunt(2).Enabled = False
    sactualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
        Select Case saccion
            Case "N"
                    ncodigo = clsDGAdqui.LogSeleccionCod(cboperiodo.Text)
                    txtSeleccion.Text = ncodigo
                    ncodigo = ClsNAdqui.AgregaProSeleccionLog(cboperiodo.Text, Right(cmbTipoSeleccion.Text, 1), Right(cmbtipoproceso.Text, 1), Right(cmbMoneda.Text, 1), txtrangoini.Text, txtrangofin.Text, Format(DTPFechaConvo.value, "ddmmyyyy"), sactualiza, UCase(Trim(txtDescripcion.Text)), UCase(txtnumerocotizacion.Text), Right(Trim(cmbclaseProceso.Text), 1))
                    If CkEvaluacionTecnica.value = 1 Then
                        ClsNAdqui.AgregaSeleccionConfigPuntaje ncodigo, 1, uSTecPunMinimo.Valor, uSTecPunMaximo.Valor, METecPeso.Text, sactualiza
                    End If
                    If CkEvaluacionEcon.value = 1 Then
                        ClsNAdqui.AgregaSeleccionConfigPuntaje ncodigo, 2, uSEconPunMinimo.Valor, uSEcoPunMaximo.Valor, MEEcoPeso.Text, sactualiza
                    End If
                    'Comite
                    'oEval.AgregaComiteProSelec Left(Me.cmbEval, 6), Me.FlexComite.GetRsNew, lsUltActualizacion
                    ClsNAdqui.AgregaSeleccionComite ncodigo, Me.FlexComite.GetRsNew, sactualiza
            Case "E"
                    ClsNAdqui.ActualizaProSeleccionLog txtSeleccion, cboperiodo.Text, Right(cmbTipoSeleccion.Text, 1), Right(cmbtipoproceso.Text, 1), Right(cmbMoneda.Text, 1), txtrangoini.Text, txtrangofin.Text, Format(DTPFechaConvo.value, "ddmmyyyy"), sactualiza, UCase(Trim(txtDescripcion.Text)), txtnumerocotizacion.Text, Right(Trim(cmbclaseProceso.Text), 1)
                    ClsNAdqui.AgregaSeleccionComite txtSeleccion, Me.FlexComite.GetRsNew, sactualiza
            Case "G"
            Case "C"
        End Select
                    nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccion.Text)
                    lblestado.Caption = clsDGAdqui.CargaLogSelEstadoDesc(nestadoProc)
                    Me.txtSeleccion.rs = clsDGAdqui.LogSeleccionLista(cboperiodo.Text)
            Bloquear_Controles "G"
End Select

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Height = 5790
Me.Width = 11320
Dim sAno As String
Set rs = New ADODB.Recordset
Set clsDGnral = New DLogGeneral
Set clsDGAdqui = New DLogAdquisi
Set ClsNAdqui = New NActualizaProcesoSelecLog
Set rs = clsDGnral.CargaPeriodo
Call CargaCombo(rs, cboperiodo)
Set rs = clsDGAdqui.CargaLogSelTipoBs
Call CargaCombo(rs, cmbTipoSeleccion)
Set rs = clsDGAdqui.CargaSelcargos
Me.FlexComite.CargaCombo rs
sAno = Year(gdFecSis)
ubicar_ano sAno, cboperiodo
'Carga el cbo con Monedas
Set rs = clsDGnral.CargaConstante(gMoneda, False)
Call CargaCombo(rs, cmbMoneda)
cmbMoneda.ListIndex = 0
'cmbmoneda.Locked = True
cmdReq(0).Enabled = True
cmdReq(1).Enabled = True
cmdReq(2).Enabled = False
cmdReq(3).Enabled = False
'Me.txtSeleccion.rs = clsDGAdqui.LogSeleccionLista(cboperiodo.Text)
puntaje
Bloquear_Controles "M"
SSTSeleccion.Tab = 0
Set rs = clsDGAdqui.CargaSelTipoComite()
cmbtipoComite.AddItem "Ninguno                                      0"
Call CargaCombo(rs, cmbtipoComite)
End Sub








Private Sub MEEcoPeso_GotFocus()
If cmbTipoSeleccion.Text = "" Then
   MsgBox "Antes debe seleccionar el tipo de proceso", vbInformation, "Seleccionar el Tipo de Proceso"
   SSTSeleccion.Tab = 0
   Exit Sub
End If

If Right(cmbTipoSeleccion.Text, 1) = "1" Then
        If Val(MEEcoPeso.Text) = 0 Then
           MsgBox "el peso ponderado no puede ser igual a 0", vbInformation, "peso ponderado no puede ser igual a 0"
           MEEcoPeso.SetFocus
           Exit Sub
        End If
        If Val(MEEcoPeso.Text) <> 1 Then
           MsgBox "el peso ponderado debe ser igual a 1", vbInformation, "peso ponderado debe ser igual a  1"
           MEEcoPeso.SetFocus
           Exit Sub
        End If
    Else
        If Val(MEEcoPeso.Text) = 0 Then
           MsgBox "el peso ponderado no puede ser igual a 0", vbInformation, "peso ponderado no puede ser igual a 0"
           MEEcoPeso.SetFocus
           Exit Sub
        End If
        If Val(MEEcoPeso.Text) >= 1 Then
           MsgBox "el peso ponderado debe ser menor a 1", vbInformation, "peso ponderado debe ser igual a  1"
           MEEcoPeso.SetFocus
           Exit Sub
        End If
End If

End Sub

Private Sub METecPeso_GotFocus()

If cmbTipoSeleccion.Text = "" Then
   MsgBox "Antes debe seleccionar el tipo de proceso", vbInformation, "Seleccionar el Tipo de Proceso"
   SSTSeleccion.Tab = 0
   Exit Sub
End If


If Right(cmbTipoSeleccion.Text, 1) = "1" Then
        If Val(METecPeso.Text) = 0 Then
           MsgBox "el peso ponderado no puede ser igual a 0", vbInformation, "peso ponderado no puede ser igual a 0"
           METecPeso.SetFocus
           Exit Sub
        End If
        If Val(METecPeso.Text) <> 1 Then
           MsgBox "el peso ponderado debe ser igual a 1", vbInformation, "peso ponderado debe ser igual a  1"
           METecPeso.SetFocus
           Exit Sub
        End If
    Else
        If Val(METecPeso.Text) = 0 Then
           MsgBox "el peso ponderado no puede ser igual a 0", vbInformation, "peso ponderado no puede ser igual a 0"
           METecPeso.SetFocus
           Exit Sub
        End If
        If Val(METecPeso.Text) >= 1 Then
           MsgBox "el peso ponderado debe ser menor a 1", vbInformation, "peso ponderado debe ser igual a  1"
           METecPeso.SetFocus
           Exit Sub
        End If
End If
End Sub

Private Sub METecPeso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If CkEvaluacionTecnica.value = 1 Then
   MEEcoPeso.SetFocus
End If


End If

End Sub

Private Sub txtSeleccion_EmiteDatos()
If txtSeleccion.Text = "" Then Exit Sub
    
mostrar_seleccion Trim(txtSeleccion.Text)
Bloquear_Controles "M"

End Sub

Sub mostrar_seleccion(LogSelCodigo As Long)
'Creacion de Proceso
Dim btec As Boolean
Dim beco As Boolean
Dim sTipoBS As String
Dim nestadoProc  As Integer
Dim rs_sel As ADODB.Recordset
Set rs_sel = New ADODB.Recordset
Set rs_sel = clsDGAdqui.CargaSeleccionProceso(LogSelCodigo)
If rs_sel.EOF = True Then
    MsgBox "No se Pudo cargar Informacion ,Consulte con Sistemas", vbInformation, "No se pudo Cargar"
    Exit Sub
End If
    sTipoBS = rs_sel!cCodTipoBS
    ubicar rs_sel!cCodTipoBS, cmbTipoSeleccion
    ubicar rs_sel!cCodTipoSelec, cmbtipoproceso
    ubicar rs_sel!cCodTipoSelec, cmbMoneda
    
    ubicar rs_sel!nLogSelTipoProceso, cmbclaseProceso
    txtDescripcion.Text = IIf(IsNull(rs_sel!cDescripcionProceso), "", rs_sel!cDescripcionProceso)
    txtrangoini.Text = Format(rs_sel!nRangoIni, "00.00")
    txtrangofin.Text = Format(rs_sel!nRangoFin, "00.00")
    DTPFechaConvo.value = rs_sel!FechaConvo
    txtnumerocotizacion.Text = IIf(IsNull(rs_sel!nLogSelNumeroCot), "", rs_sel!nLogSelNumeroCot)
'Configuracion de Puntaje
Set rs_sel = clsDGAdqui.CargaSeleccionPuntajes(LogSelCodigo)
If rs_sel.EOF = True Then
    MsgBox "Proceso No Tiene Ingresado Los Puntajes ", vbInformation, "No Tiene puntajes Ingresados"
    uSTecPunMinimo.Valor = 0
    uSTecPunMaximo.Valor = 50
    uSEconPunMinimo.Valor = 0
    uSEcoPunMaximo.Valor = 50
    METecPeso.Text = "00.50"
    MEEcoPeso.Text = "00.50"
End If
    Do While Not rs_sel.EOF = True
        If rs_sel!cCodTipoEvaluacion = SelEvalTecnica Then
            CkEvaluacionTecnica.value = 1
            uSTecPunMaximo.Valor = rs_sel!nPuntajeMaximo
            uSTecPunMinimo.Valor = rs_sel!nPuntajeMinimo
            Select Case sTipoBS
            Case "1" 'Bienes
                    METecPeso.Mask = "#.##"
                    METecPeso.Text = Format(rs_sel!nPesoPonderado, "0.00")
            Case "2", "3" 'demas
                    METecPeso.Mask = "#.##"
                    METecPeso.Text = Format(rs_sel!nPesoPonderado, "0.00")
            End Select
            btec = True
        End If
        If rs_sel!cCodTipoEvaluacion = SelEvalEconomica Then
            CkEvaluacionEcon.value = 1
            uSEcoPunMaximo.Valor = rs_sel!nPuntajeMaximo
            uSEconPunMinimo.Valor = rs_sel!nPuntajeMinimo
            Select Case sTipoBS
            Case "1" 'Bienes
                    MEEcoPeso.Mask = "#.##"
                    MEEcoPeso.Text = Format(rs_sel!nPesoPonderado, "0.00")
            Case "2", "3" 'demas
                    MEEcoPeso.Mask = "#.##"
                    MEEcoPeso.Text = Format(rs_sel!nPesoPonderado, "0.00")
            End Select
            
            
            beco = True
        End If
        rs_sel.MoveNext
    Loop
    If btec = False Then
    CkEvaluacionTecnica.value = 0
    End If
    If beco = False Then
    CkEvaluacionEcon.value = 0
    End If
'Comite
Set rs_sel = clsDGAdqui.CargaSeleccionComite(LogSelCodigo)
If rs_sel.EOF = True Then
   FlexComite.Rows = 2
   Else
   
   Set FlexComite.Recordset = rs_sel
End If
'Carga Estado Proceso
    nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccion.Text)
    lblestado.Caption = clsDGAdqui.CargaLogSelEstadoDesc(nestadoProc)

End Sub
Private Sub uSEconPunMinimo_Change()
If Val(uSEconPunMinimo.Valor) > Val(uSEcoPunMaximo.Valor) Then
    MsgBox "El Puntaje Minimo no Puede ser Mayor que el Maximo ", vbInformation, "Verifique Valores"
    uSEconPunMinimo.Valor = uSEcoPunMaximo.Valor
    Exit Sub
End If
End Sub

Private Sub uSEcoPunMaximo_Change()
If Val(uSEconPunMinimo.Valor) > Val(uSEcoPunMaximo.Valor) Then
    MsgBox "El Puntaje Maximo no Puede ser Menor que el Minimo ", vbInformation, "Verifique Valores"
    uSEcoPunMaximo.Valor = uSEconPunMinimo.Valor
    Exit Sub
End If
End Sub

Private Sub uSTecPunMaximo_Change()
If Val(uSTecPunMinimo.Valor) > Val(uSTecPunMaximo.Valor) Then
    MsgBox "El Puntaje Maximo no Puede ser Menor que el Minimo ", vbInformation, "Verifique Valores"
    uSTecPunMaximo.Valor = uSTecPunMinimo.Valor
    Exit Sub
End If
End Sub

Private Sub uSTecPunMinimo_Change()
If Val(uSTecPunMinimo.Valor) > Val(uSTecPunMaximo.Valor) Then
    MsgBox "El Puntaje Minimo no Puede ser Mayor que el Maximo ", vbInformation, "Verifique Valores"
    uSTecPunMinimo.Valor = uSTecPunMaximo.Valor
    Exit Sub
End If
End Sub

Sub ubicar(codigo As String, combo As ComboBox)
Dim i As Integer
For i = 0 To combo.ListCount
If Right(combo.List(i), 1) = codigo Then
    combo.ListIndex = i
    Exit For
    End If
Next
End Sub

Sub Bloquear_Controles(saccion As String)
Select Case saccion
Case "M", "G" 'mostrar
            cboperiodo.Enabled = True
            txtnumerocotizacion.Enabled = False
            cmbclaseProceso.Enabled = False
            cmbTipoSeleccion.Enabled = False
            cmbtipoproceso.Enabled = False
            cmbMoneda.Enabled = False
            DTPFechaConvo.Enabled = False
            CkEvaluacionTecnica.Enabled = False
            uSTecPunMinimo.Enabled = False
            uSTecPunMaximo.Enabled = False
            METecPeso.Enabled = False
            txtDescripcion.Enabled = False
            If CkEvaluacionTecnica.value = 1 Then
                uSTecPunMinimo.BackColor = -2147483643
                uSTecPunMaximo.BackColor = -2147483643
                METecPeso.BackColor = -2147483643
               Else
                uSEconPunMinimo.BackColor = 14737632
                uSEcoPunMaximo.BackColor = 14737632
                MEEcoPeso.BackColor = 14737632
            End If
            If CkEvaluacionEcon.value = 1 Then
                uSEconPunMinimo.BackColor = -2147483643
                uSEcoPunMaximo.BackColor = -2147483643
                MEEcoPeso.BackColor = -2147483643
            Else
                uSTecPunMinimo.BackColor = 14737632
                uSTecPunMaximo.BackColor = 14737632
                METecPeso.BackColor = 14737632
            End If
            CkEvaluacionEcon.Enabled = False
            uSEconPunMinimo.Enabled = False
            uSEcoPunMaximo.Enabled = False
            MEEcoPeso.Enabled = False
            FlexComite.Enabled = False
            cmdNuevoComite.Enabled = False
            cmdEliminarComite.Enabled = False
            cmdcomitepreter.Enabled = False
            cmbtipoComite.Enabled = False
            
Case "N" 'nuevo
            txtSeleccion.Text = ""
            cmbclaseProceso.Enabled = True
            txtnumerocotizacion.Enabled = True
            txtDescripcion.Enabled = True
            cboperiodo.Enabled = True
            cmbTipoSeleccion.Enabled = True
            cmbtipoproceso.Enabled = True
            cmbMoneda.Enabled = True
            DTPFechaConvo.Enabled = True
            CkEvaluacionTecnica.Enabled = True
            uSTecPunMinimo.Enabled = True
            uSTecPunMaximo.Enabled = True
            METecPeso.Enabled = True
            If CkEvaluacionTecnica.value = 1 Then
                uSTecPunMinimo.BackColor = -2147483643
                uSTecPunMaximo.BackColor = -2147483643
                METecPeso.BackColor = -2147483643
               Else
                uSEconPunMinimo.BackColor = 14737632
                uSEcoPunMaximo.BackColor = 14737632
                MEEcoPeso.BackColor = 14737632
            End If
            If CkEvaluacionEcon.value = 1 Then
                uSEconPunMinimo.BackColor = -2147483643
                uSEcoPunMaximo.BackColor = -2147483643
                MEEcoPeso.BackColor = -2147483643
            Else
                uSTecPunMinimo.BackColor = 14737632
                uSTecPunMaximo.BackColor = 14737632
                METecPeso.BackColor = 14737632
            End If
            CkEvaluacionEcon.Enabled = True
            uSEconPunMinimo.Enabled = True
            uSEcoPunMaximo.Enabled = True
            MEEcoPeso.Enabled = True
            FlexComite.Enabled = True
            cmdNuevoComite.Enabled = True
            cmdEliminarComite.Enabled = True
            cmdcomitepreter.Enabled = True
            cmbtipoComite.Enabled = True
Case "E" 'Editar
            txtnumerocotizacion.Enabled = True
            cmbclaseProceso.Enabled = True
            txtDescripcion.Enabled = True
            cboperiodo.Enabled = False
            txtSeleccion.Enabled = False
            cmdReq(0).Enabled = False
            cmdReq(1).Enabled = False
            cmdReq(2).Enabled = True
            cmdReq(3).Enabled = True
            cmdPunt(0).Visible = True
            cmdPunt(1).Visible = True
            cmdPunt(2).Visible = True
            cmdPunt(0).Enabled = True
            cmdPunt(1).Enabled = False
            cmdPunt(2).Enabled = False
            cmbTipoSeleccion.Enabled = True
            cmbtipoproceso.Enabled = True
            cmbMoneda.Enabled = True
            DTPFechaConvo.Enabled = True
            'Comite
            FlexComite.Enabled = True
            cmdNuevoComite.Enabled = True
            cmdEliminarComite.Enabled = True
            cmdcomitepreter.Enabled = True
            cmbtipoComite.Enabled = True
Case "G" 'Grabar
End Select
End Sub


Sub Limpiar_Controles()
      ubicar_ano Year(gdFecSis), cboperiodo
      txtSeleccion.Text = ""
      cmbTipoSeleccion.Enabled = True
      DTPFechaConvo.Enabled = True
      CkEvaluacionTecnica.value = 0
      uSTecPunMinimo.Valor = 0
      uSTecPunMaximo.Valor = 0
      METecPeso.Mask = "1.##"
      METecPeso.Text = "1.00"
      If CkEvaluacionTecnica.value = 1 Then
          uSTecPunMinimo.BackColor = -2147483643
          uSTecPunMaximo.BackColor = -2147483643
           METecPeso.BackColor = -2147483643
      Else
           uSEconPunMinimo.BackColor = 14737632
            uSEcoPunMaximo.BackColor = 14737632
            MEEcoPeso.BackColor = 14737632
      End If
          CkEvaluacionEcon.value = 0
          uSEconPunMinimo.Valor = 0
          uSEcoPunMaximo.Valor = 0
          MEEcoPeso.Mask = "1.##"
          MEEcoPeso.Text = "1.00"
           If CkEvaluacionEcon.value = 1 Then
                uSEconPunMinimo.BackColor = -2147483643
                uSEcoPunMaximo.BackColor = -2147483643
                MEEcoPeso.BackColor = -2147483643
            Else
                uSTecPunMinimo.BackColor = 14737632
                uSTecPunMaximo.BackColor = 14737632
                METecPeso.BackColor = 14737632
            End If
            FlexComite.Clear
            FlexComite.FormaCabecera
            FlexComite.Rows = 2
End Sub

Sub puntaje()

uSTecPunMaximo.Valor = 50
uSEcoPunMaximo.Valor = 50
uSTecPunMinimo.Valor = 30
uSEconPunMinimo.Valor = 30
METecPeso.Text = "1.00"
MEEcoPeso.Text = "1.00"
uSTecPunMinimo.Enabled = False
uSTecPunMaximo.Enabled = False
METecPeso.Enabled = False
uSTecPunMinimo.BackColor = 14737632
uSTecPunMaximo.BackColor = 14737632
METecPeso.BackColor = 14737632
uSEconPunMinimo.Enabled = False
uSEcoPunMaximo.Enabled = False
MEEcoPeso.Enabled = False
uSEconPunMinimo.BackColor = 14737632
uSEcoPunMaximo.BackColor = 14737632
MEEcoPeso.BackColor = 14737632
cmdPunt(0).Visible = False
cmdPunt(1).Visible = False
cmdPunt(2).Visible = False
End Sub


Function valida_ponderado() As Integer
Dim sTipoBS  As String
valida_ponderado = 0
sTipoBS = Right(Trim(cmbTipoSeleccion.Text), 1)

                If Val(uSTecPunMaximo.Valor) = 0 Then
                      MsgBox "el puntaje maximo no puede ser igual a 0", vbInformation, "el puntaje maximo no puede ser igual a 0"
                      uSTecPunMaximo.SetFocus
                      valida_ponderado = -1
                      Exit Function
                End If
                If Val(uSEcoPunMaximo.Valor) = 0 Then
                      MsgBox "el puntaje maximo no puede ser igual a 0", vbInformation, "el puntaje maximo no puede ser igual a 0"
                      uSEcoPunMaximo.SetFocus
                      valida_ponderado = -1
                      Exit Function
                End If
                If Val(MEEcoPeso.Text) <= 0 Then
                      MsgBox "El peso ponderado no Puede ser Igual a 0 ", vbInformation, "el peso ponderado no puede ser igual a 0"
                      MEEcoPeso.SetFocus
                      valida_ponderado = -1
                      Exit Function
                End If
                If Val(METecPeso.Text) <= 0 Then
                      MsgBox "El peso ponderado no Puede ser Igual a 0 ", vbInformation, "el peso ponderado no puede ser igual a 0"
                      METecPeso.SetFocus
                      valida_ponderado = -1
                      Exit Function
                End If
                If sTipoBS = "1" Then
                      If Val(uSTecPunMaximo.Valor) > 100 Then
                          MsgBox "el puntaje maximo no puede ser mayor a 100", vbInformation, "El Puntaje Maximo no puede ser mayor a 100"
                          uSTecPunMaximo.SetFocus
                          valida_ponderado = -1
                          Exit Function
                      End If
                      If Val(uSEcoPunMaximo.Valor) > 100 Then
                          MsgBox "el puntaje maximo no puede ser mayor a 100", vbInformation, "El Puntaje Maximo no puede ser mayor a 100"
                          uSEcoPunMaximo.SetFocus
                          valida_ponderado = -1
                          Exit Function
                      End If
                      If Val(MEEcoPeso.Text) <> 1 Then
                          MsgBox "El Peso Ponderado debe ser Igual a 1 ", vbInformation, "el peso Ponderado debe ser igual a 1"
                          METecPeso.SetFocus
                          valida_ponderado = -1
                          Exit Function
                      End If
                      If Val(METecPeso.Text) <> 1 Then
                          MsgBox "El Peso Ponderado debe ser Igual a 1 ", vbInformation, "el peso Ponderado debe ser igual a 1"
                          METecPeso.SetFocus
                          valida_ponderado = -1
                          Exit Function
                       End If
                   Else
                       If Val(uSTecPunMaximo.Valor) + Val(uSEcoPunMaximo.Valor) < 100 Then
                          MsgBox "La Suma de los Puntajes Maximos no pueden ser menor a 100", vbInformation, "La Suma de Los Puntajes Maximo debe ser 100"
                          valida_ponderado = -1
                          Exit Function
                       End If
                       If Val(uSTecPunMaximo.Valor) + Val(uSEcoPunMaximo.Valor) > 200 Then
                          MsgBox "La Suma de los Puntajes Maximos no pueden ser mayor a 100", vbInformation, "La Suma de Los Puntajes Maximo debe ser 100"
                          valida_ponderado = -1
                          uSTecPunMaximo.SetFocus
                          Exit Function
                       End If
                       If Val(METecPeso.Text) + Val(MEEcoPeso.Text) < 1 Then
                          MsgBox "La Suma de los Pesos Ponderados debe ser igual a 1 ", vbInformation, "La Suma de Peso Ponderados debe ser igual a 1"
                          METecPeso.SetFocus
                          valida_ponderado = -1
                          Exit Function
                       End If
                       If Val(METecPeso.Text) + Val(MEEcoPeso.Text) > 1 Then
                          MsgBox "La Suma de los Pesos Ponderados debe ser igual a 1 ", vbInformation, "La Suma de Peso Ponderados debe ser igual a 1"
                          METecPeso.SetFocus
                          valida_ponderado = -1
                          Exit Function
                       End If
                End If
End Function

Sub mostrar_Comite_Tipo(nLogTipoComite As Long)
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = clsDGAdqui.CargaLogSelComitePre(nLogTipoComite)
    If rs.EOF = True Then
        FlexComite.Rows = 2
        FlexComite.Clear
        
        FlexComite.FormaCabecera
        Else
        Set FlexComite.Recordset = rs
        FlexComite.ColWidth(4) = 0
    End If
End Sub
