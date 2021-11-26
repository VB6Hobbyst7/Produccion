VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCredLineaCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administracion de  Lineas de Credito"
   ClientHeight    =   7875
   ClientLeft      =   1530
   ClientTop       =   585
   ClientWidth     =   9360
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdTarifario 
      Caption         =   "&Consultar Tarifario"
      Height          =   330
      Left            =   7200
      TabIndex        =   54
      Top             =   7440
      Width           =   2010
   End
   Begin VB.Frame Frame1 
      Height          =   7440
      Left            =   15
      TabIndex        =   13
      Top             =   -45
      Width           =   9255
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   360
         Left            =   7620
         TabIndex        =   31
         Top             =   195
         Width           =   1275
      End
      Begin VB.Frame Frame3 
         Height          =   4530
         Left            =   90
         TabIndex        =   25
         Top             =   555
         Width           =   9045
         Begin VB.CheckBox ChkPreferencial 
            Caption         =   "Preferencial"
            Enabled         =   0   'False
            Height          =   345
            Left            =   4230
            TabIndex        =   48
            Top             =   4140
            Width           =   1575
         End
         Begin VB.Frame Frame5 
            Height          =   3885
            Left            =   4155
            TabIndex        =   38
            Top             =   120
            Width           =   4770
            Begin TabDlg.SSTab SSTablinea 
               Height          =   3555
               Left            =   60
               TabIndex        =   40
               Top             =   255
               Width           =   4650
               _ExtentX        =   8202
               _ExtentY        =   6271
               _Version        =   393216
               TabHeight       =   520
               TabCaption(0)   =   "Busqueda"
               TabPicture(0)   =   "frmCredLineaCredito.frx":0000
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "LstLineas"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "TxtLineaBusq"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).ControlCount=   2
               TabCaption(1)   =   "Datos de Linea"
               TabPicture(1)   =   "frmCredLineaCredito.frx":001C
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Label6"
               Tab(1).Control(1)=   "Label7"
               Tab(1).Control(2)=   "Label8"
               Tab(1).Control(3)=   "Label9"
               Tab(1).Control(4)=   "Label10"
               Tab(1).Control(5)=   "LblLineaDesc"
               Tab(1).Control(6)=   "Label11"
               Tab(1).Control(7)=   "CmbEstado"
               Tab(1).Control(8)=   "TxtPlazoMin"
               Tab(1).Control(9)=   "TxtPlazoMax"
               Tab(1).Control(10)=   "TxtMontoMin"
               Tab(1).Control(11)=   "TxtMontoMax"
               Tab(1).Control(12)=   "TxtLinDesc"
               Tab(1).ControlCount=   13
               TabCaption(2)   =   "Agencias"
               TabPicture(2)   =   "frmCredLineaCredito.frx":0038
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "fraContenedor"
               Tab(2).Control(1)=   "chkTodos"
               Tab(2).ControlCount=   2
               Begin VB.CheckBox chkTodos 
                  Caption         =   "Todos"
                  Height          =   255
                  Left            =   -74580
                  TabIndex        =   51
                  Top             =   420
                  Width           =   1335
               End
               Begin VB.Frame fraContenedor 
                  Caption         =   "Haga un Click en la Agencia a escoger "
                  Height          =   2475
                  Left            =   -74580
                  TabIndex        =   49
                  Top             =   660
                  Width           =   3645
                  Begin VB.ListBox lstAgencias 
                     Height          =   2085
                     ItemData        =   "frmCredLineaCredito.frx":0054
                     Left            =   105
                     List            =   "frmCredLineaCredito.frx":005B
                     Style           =   1  'Checkbox
                     TabIndex        =   50
                     Top             =   255
                     Width           =   3405
                  End
               End
               Begin VB.TextBox TxtLinDesc 
                  Height          =   285
                  Left            =   -73485
                  MaxLength       =   10
                  TabIndex        =   10
                  Top             =   2505
                  Width           =   1185
               End
               Begin VB.TextBox TxtLineaBusq 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2760
                  TabIndex        =   11
                  Top             =   480
                  Width           =   1755
               End
               Begin MSComctlLib.ListView LstLineas 
                  Height          =   2610
                  Left            =   135
                  TabIndex        =   12
                  Top             =   840
                  Width           =   4380
                  _ExtentX        =   7726
                  _ExtentY        =   4604
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   0   'False
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   7
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "Codigo"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "Descripcion"
                     Object.Width           =   6174
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Text            =   "Estado"
                     Object.Width           =   0
                  EndProperty
                  BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   3
                     Text            =   "Plazo Min"
                     Object.Width           =   0
                  EndProperty
                  BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   4
                     Text            =   "Plazo Max"
                     Object.Width           =   0
                  EndProperty
                  BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   5
                     Text            =   "Monto Min"
                     Object.Width           =   0
                  EndProperty
                  BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   6
                     Text            =   "Monto Max"
                     Object.Width           =   0
                  EndProperty
               End
               Begin VB.TextBox TxtMontoMax 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   -73485
                  TabIndex        =   9
                  Text            =   "0"
                  Top             =   2130
                  Width           =   1170
               End
               Begin VB.TextBox TxtMontoMin 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   -73485
                  TabIndex        =   8
                  Text            =   "0"
                  Top             =   1755
                  Width           =   1155
               End
               Begin VB.TextBox TxtPlazoMax 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   -71565
                  TabIndex        =   7
                  Text            =   "0"
                  Top             =   1395
                  Width           =   615
               End
               Begin VB.TextBox TxtPlazoMin 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   -73485
                  TabIndex        =   6
                  Text            =   "0"
                  Top             =   1365
                  Width           =   615
               End
               Begin VB.ComboBox CmbEstado 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   -73485
                  Style           =   2  'Dropdown List
                  TabIndex        =   5
                  Top             =   975
                  Width           =   1515
               End
               Begin VB.Label Label11 
                  Caption         =   "Descripcion"
                  Height          =   195
                  Left            =   -74640
                  TabIndex        =   47
                  Top             =   2535
                  Width           =   1125
               End
               Begin VB.Label LblLineaDesc 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   -73845
                  TabIndex        =   46
                  Top             =   510
                  Width           =   2295
               End
               Begin VB.Label Label10 
                  Caption         =   "Monto Maximo :"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   -74655
                  TabIndex        =   45
                  Top             =   2160
                  Width           =   1125
               End
               Begin VB.Label Label9 
                  Caption         =   "Monto Minimo :"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   -74655
                  TabIndex        =   44
                  Top             =   1800
                  Width           =   1080
               End
               Begin VB.Label Label8 
                  Caption         =   "Plazo Maximo :"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   -72720
                  TabIndex        =   43
                  Top             =   1410
                  Width           =   1080
               End
               Begin VB.Label Label7 
                  Caption         =   "Plazo Minimo :"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   -74655
                  TabIndex        =   42
                  Top             =   1410
                  Width           =   1080
               End
               Begin VB.Label Label6 
                  Caption         =   "Estado :"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   -74655
                  TabIndex        =   41
                  Top             =   1050
                  Width           =   630
               End
            End
         End
         Begin VB.Frame Frame4 
            Height          =   3885
            Left            =   105
            TabIndex        =   32
            Top             =   120
            Width           =   4005
            Begin VB.ComboBox cmbPaquete 
               Enabled         =   0   'False
               Height          =   315
               Left            =   180
               Style           =   2  'Dropdown List
               TabIndex        =   52
               Top             =   3450
               Width           =   3615
            End
            Begin VB.ComboBox CmbProd 
               Enabled         =   0   'False
               Height          =   315
               Left            =   150
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   2835
               Width           =   3615
            End
            Begin VB.ComboBox CmbPlazo 
               Enabled         =   0   'False
               Height          =   315
               Left            =   135
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   2205
               Width           =   2190
            End
            Begin VB.ComboBox CmbMoneda 
               Enabled         =   0   'False
               Height          =   315
               Left            =   135
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   1605
               Width           =   2175
            End
            Begin VB.ComboBox CmbSubFondo 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmCredLineaCredito.frx":006C
               Left            =   105
               List            =   "frmCredLineaCredito.frx":006E
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   990
               Width           =   3675
            End
            Begin VB.ComboBox CmbFondo 
               Enabled         =   0   'False
               Height          =   315
               Left            =   105
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   375
               Width           =   3690
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Paquete:"
               Enabled         =   0   'False
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
               TabIndex        =   53
               Top             =   3240
               Width           =   780
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sub Tipo de Credito :"
               Enabled         =   0   'False
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
               Left            =   90
               TabIndex        =   37
               Top             =   2625
               Width           =   1830
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Plazo :"
               Enabled         =   0   'False
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
               Left            =   90
               TabIndex        =   36
               Top             =   2000
               Width           =   600
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda :"
               Enabled         =   0   'False
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
               Left            =   90
               TabIndex        =   35
               Top             =   1380
               Width           =   810
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sub Fondo :"
               Enabled         =   0   'False
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
               Left            =   90
               TabIndex        =   34
               Top             =   780
               Width           =   1050
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fondo :"
               Enabled         =   0   'False
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
               Left            =   90
               TabIndex        =   33
               Top             =   135
               Width           =   660
            End
         End
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   360
            Left            =   7620
            TabIndex        =   30
            Top             =   4080
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CommandButton CmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   360
            Left            =   6285
            TabIndex        =   29
            Top             =   4080
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CommandButton CmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   360
            Left            =   2820
            TabIndex        =   28
            Top             =   4095
            Width           =   1275
         End
         Begin VB.CommandButton CmdEditar 
            Caption         =   "&Modificar"
            Height          =   360
            Left            =   1485
            TabIndex        =   27
            Top             =   4095
            Width           =   1275
         End
         Begin VB.CommandButton CmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   360
            Left            =   150
            TabIndex        =   26
            Top             =   4095
            Width           =   1275
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2250
         Left            =   90
         TabIndex        =   14
         Top             =   5070
         Width           =   9075
         Begin VB.CommandButton CmdTasasCancelar 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   330
            Left            =   7515
            TabIndex        =   24
            Top             =   1845
            Width           =   1410
         End
         Begin VB.CommandButton CmdTasasAceptar 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   330
            Left            =   7515
            TabIndex        =   23
            Top             =   1485
            Width           =   1410
         End
         Begin VB.CommandButton CmdTasaNuevo 
            Caption         =   "&Nuevo"
            Enabled         =   0   'False
            Height          =   330
            Left            =   7515
            TabIndex        =   17
            Top             =   375
            Width           =   1410
         End
         Begin VB.CommandButton CmdTasaEditar 
            Caption         =   "Editar"
            Enabled         =   0   'False
            Height          =   330
            Left            =   7515
            TabIndex        =   16
            Top             =   735
            Width           =   1410
         End
         Begin VB.CommandButton CmdTasaEliminar 
            Caption         =   "&Eliminar"
            Enabled         =   0   'False
            Height          =   330
            Left            =   7515
            TabIndex        =   15
            Top             =   1110
            Width           =   1410
         End
         Begin SICMACT.FlexEdit FETasas 
            Height          =   1695
            Left            =   105
            TabIndex        =   18
            Top             =   510
            Width           =   7185
            _ExtentX        =   12674
            _ExtentY        =   2990
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Tipo-Tasa Inicial-Tasa Final"
            EncabezadosAnchos=   "350-3200-1000-900"
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
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-3"
            ListaControles  =   "0-3-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-R"
            FormatosEdit    =   "0-0-2-2"
            CantDecimales   =   4
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
         End
         Begin VB.Label LblTasas 
            BackStyle       =   0  'Transparent
            Caption         =   "Tasas de Linea de Credito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   2850
         End
         Begin VB.Label LblTasas2 
            BackStyle       =   0  'Transparent
            Caption         =   "Tasas de Linea de Credito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   2850
         End
      End
      Begin VB.Label LblLinea 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   2640
         TabIndex        =   39
         Top             =   240
         Width           =   3990
      End
      Begin VB.Label LblLineaCred 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lineas de Credito :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   180
         TabIndex        =   20
         Top             =   255
         Width           =   1965
      End
      Begin VB.Label LblLineacred2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lineas de Credito :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   165
         TabIndex        =   21
         Top             =   240
         Width           =   1965
      End
   End
   Begin VB.Menu mnuLineaCred 
      Caption         =   "LineaCred"
      Visible         =   0   'False
      Begin VB.Menu mnuLCNueva 
         Caption         =   "Nueva Linea"
      End
      Begin VB.Menu mnuLCModificar 
         Caption         =   "Modificar Linea"
      End
      Begin VB.Menu mnuLCEliminar 
         Caption         =   "Eliminar Linea"
      End
   End
End
Attribute VB_Name = "frmCredLineaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Tipo Opcion
'1 : Registrar
'2 : Mantenimiento
'3 : Consulta
Private nTipoOpcion As Integer
Private nFilaLinea As Integer
Private nTipoProceso As Integer
Private nTipoProcesoTasa As Integer
Private sColBloq As String
Private sPersCodTemp As String
Dim bAplDes As Boolean

Dim MatAgencias() As String
Dim objPista As COMManejador.Pista
'************************
'Modificado 03-05-2006
'Para el manejo de los Paquetes de Adeudados

Private Sub CargaSubFondos(ByVal psFondo As String)
Dim oLinea As COMDCredito.DCOMLineaCredito
Dim R As ADODB.Recordset

    'Carga Sub Fondos de Linea de Credito
    CmbSubFondo.Clear
    Set oLinea = New COMDCredito.DCOMLineaCredito
    Set R = oLinea.RecuperaSubFondos(psFondo)
    Do While Not R.EOF
        CmbSubFondo.AddItem Trim(R!cDescripcion) & Space(100 - Len(Trim(R!cDescripcion))) & R!cSubFondo & Space(100) & Trim(R!cAbrev)
        R.MoveNext
    Loop
    R.Close
    CmbSubFondo.AddItem "<Nuevo SubFondo>" & Space(100) & "new"
    Set oLinea = Nothing
End Sub

Private Sub HabilitaCombos(ByVal pbHabilita As Boolean)
    Label1.Enabled = pbHabilita
    CmbFondo.Enabled = pbHabilita
    Label2.Enabled = pbHabilita
    CmbSubFondo.Enabled = pbHabilita
    Label3.Enabled = pbHabilita
    CmbMoneda.Enabled = pbHabilita
    Label4.Enabled = pbHabilita
    CmbPlazo.Enabled = pbHabilita
    Label5.Enabled = pbHabilita
    CmbProd.Enabled = pbHabilita
    ChkPreferencial.Enabled = pbHabilita
    cmbPaquete.Enabled = pbHabilita
    Label12.Enabled = pbHabilita
End Sub

Private Sub HabilitaBusqueda(ByVal pbHabilita As Boolean)
    TxtLineaBusq.Enabled = pbHabilita
    LstLineas.Enabled = pbHabilita
End Sub

Private Sub HabilitaDatosLinea(ByVal pbHabilita As Boolean)
    LblLineaDesc.Enabled = pbHabilita
    Label6.Enabled = pbHabilita
    CmbEstado.Enabled = pbHabilita
    Label7.Enabled = pbHabilita
    TxtPlazoMin.Enabled = pbHabilita
    Label8.Enabled = pbHabilita
    TxtPlazoMax.Enabled = pbHabilita
    Label9.Enabled = pbHabilita
    TxtMontoMin.Enabled = pbHabilita
    Label10.Enabled = pbHabilita
    TxtMontoMax.Enabled = pbHabilita
    Label11.Enabled = pbHabilita
    TxtLinDesc.Enabled = pbHabilita
    'CUSCO
    fraContenedor.Enabled = pbHabilita
    chkTodos.Enabled = pbHabilita
End Sub

Private Sub HabilitaBotones(ByVal pbHabilita As Boolean)
    CmdNuevo.Enabled = Not pbHabilita
    CmdEditar.Enabled = Not pbHabilita
    CmdEliminar.Enabled = Not pbHabilita
    CmdAceptar.Enabled = pbHabilita
    CmdCancelar.Enabled = pbHabilita
    CmdAceptar.Visible = pbHabilita
    CmdCancelar.Visible = pbHabilita
    Frame2.Enabled = Not pbHabilita
End Sub

Private Sub CargaLineasCredito(ByVal psCriterio As String)
Dim oLinea As COMDCredito.DCOMLineaCredito
Dim R As ADODB.Recordset
Dim L As ListItem

    Set oLinea = New COMDCredito.DCOMLineaCredito
    Set R = oLinea.RecuperaLineasCredito(5, psCriterio)
    LstLineas.ListItems.Clear
    Do While Not R.EOF
        'CL.cLineaCred, CL.cDescripcion, CL.nPlazoMax, CL.nPlazoMin, CL.nMontoMax, CL.nMontoMin, P.cPersNombre + space(50) + P.cPersCod as PersCod, convert(int,CL.bEstado)
        Set L = LstLineas.ListItems.Add(, , R!cLineaCred)
        L.SubItems(1) = R!cDescripcion
        L.SubItems(2) = R!nEstado
        L.SubItems(3) = R!nPlazoMin
        L.SubItems(4) = R!nplazomax
        L.SubItems(5) = R!nMontoMin
        L.SubItems(6) = R!nMontoMax
        R.MoveNext
    Loop
    Set oLinea = Nothing
End Sub

Private Function ObtieneAbrevProd(ByVal psProd As String) As String
Dim i, J As Integer
Dim Abrev1, Abrev2 As String
    Abrev1 = Left(psProd, 3)
    J = -1
    For i = 1 To Len(psProd)
        If Mid(psProd, i, 1) = " " Then
            J = i + 1
            Exit For
        End If
    Next i
    Abrev2 = ""
    If J <> -1 Then
        Abrev2 = Mid(psProd, J, 3)
        ObtieneAbrevProd = Abrev1 & "-" & Abrev2
    Else
        ObtieneAbrevProd = Abrev1
    End If
End Function


Private Sub CargaControles()
Dim oLinea As COMDCredito.DCOMLineaCredito
'Dim oCred As COMDCredito.DCOMCredito
'Dim oCons As DConstante
'Dim R As adodb.Recordset

Dim lrsFondos As ADODB.Recordset
Dim lrsProductos As ADODB.Recordset
'CUSCO
Dim lrsAgencias As ADODB.Recordset

Set oLinea = New COMDCredito.DCOMLineaCredito
Call oLinea.Cargar_Datos_Objetos_LineaCredito(lrsFondos, lrsProductos, lrsAgencias)

'Carga Fondos de Linea de Credito
    
    CmbFondo.Clear
'    Set oLinea = New COMDCredito.DCOMLineaCredito
'    Set R = oLinea.RecuperaFondos
    Do While Not lrsFondos.EOF
        CmbFondo.AddItem Trim(lrsFondos!cPersNombre) & Space(100 - Len(Trim(lrsFondos!cPersNombre))) & lrsFondos!cLineaCred & Space(100) & Trim(lrsFondos!cAbrev)
        lrsFondos.MoveNext
    Loop
    lrsFondos.Close
    CmbFondo.AddItem "<Nuevo Fondo>" & Space(100) & "new"
    
    'Carga Moneda
    CmbMoneda.Clear
    CmbMoneda.AddItem "SOLES" & Space(95) & "1" & Space(100) & "MN"
    CmbMoneda.AddItem "DOLARES" & Space(93) & "2" & Space(100) & "ME"
    
    'Carga Plazo
    CmbPlazo.Clear
    CmbPlazo.AddItem "CORTO PLAZO" & Space(89) & "1" & Space(100) & "CP"
    CmbPlazo.AddItem "LARGO PLAZO" & Space(89) & "2" & Space(100) & "LP"
    
    'Carga Tipos de Producto
    CmbProd.Clear
    'Set oCred = New COMDCredito.DCOMCredito
    'Set R = oCred.RecuperaProductosDeCredito
    Do While Not lrsProductos.EOF
        CmbProd.AddItem Trim(lrsProductos!cConsDescripcion) & Space(100 - Len(Trim(lrsProductos!cConsDescripcion))) & lrsProductos!nConsValor & Space(100) & ObtieneAbrevProd(Trim(lrsProductos!cConsDescripcion))
        lrsProductos.MoveNext
    Loop
    'Set oCred = Nothing
    
    'Carga Estados de la Linea
    CmbEstado.Clear
    CmbEstado.AddItem "ACTIVA" & Space(100) & "1"
    CmbEstado.AddItem "INACTIVA" & Space(100) & "0"
    
    'CUSCO
    Me.lstAgencias.Clear
    With lrsAgencias
        Do While Not .EOF
            lstAgencias.AddItem !cAgeCod & " " & Trim(!cAgeDescripcion)
            'If !cAgeCod = gsCodAge Then
            '    lstAgencias.Selected(lstAgencias.ListCount - 1) = True
            'End If
            .MoveNext
        Loop
    End With
    
    cmbPaquete.Clear
    SSTablinea.Tab = 0
End Sub

Private Sub HabilitaIngresoLineaTasa(ByVal pbHabilita As Boolean)
Dim oConstante As COMDConstantes.DCOMConstantes
    FETasas.lbEditarFlex = pbHabilita
    Select Case nTipoOpcion
        Case 1
            CmdTasaNuevo.Enabled = Not pbHabilita
            CmdTasaEditar.Enabled = Not pbHabilita
            CmdTasaEliminar.Enabled = Not pbHabilita
        Case 2
            CmdTasaNuevo.Enabled = False
            CmdTasaEditar.Enabled = Not pbHabilita
            CmdTasaEliminar.Enabled = Not pbHabilita
    End Select
    
    CmdTasasAceptar.Enabled = pbHabilita
    CmdTasasCancelar.Enabled = pbHabilita
    Set oConstante = New COMDConstantes.DCOMConstantes
        FETasas.CargaCombo oConstante.RecuperaConstantes(gColocLineaCredTasas)
    Set oConstante = Nothing
End Sub


Public Sub Registrar()
    nTipoOpcion = 1
    Me.Show 1
End Sub

Public Sub Actualizar()
    nTipoOpcion = 2
    CmdNuevo.Enabled = False
    CmdTasaNuevo.Enabled = False
    Me.Show 1
End Sub

Public Sub Consultar()
    nTipoOpcion = 3
    CmdNuevo.Enabled = False
    CmdEditar.Enabled = False
    CmdEliminar.Enabled = False
    CmdTasaNuevo.Enabled = False
    CmdTasaEditar.Enabled = False
    CmdTasaEliminar.Enabled = False
    Me.Show 1
End Sub

Private Sub CargaDetalleLinea(ByVal psLineaCred As String)
Dim RTasas As ADODB.Recordset
Dim oLinea As COMDCredito.DCOMLineaCredito
Dim RAgencias As ADODB.Recordset
Dim i As Integer
'Carga Tasas de Linea
    Set oLinea = New COMDCredito.DCOMLineaCredito
    LimpiaFlex FETasas
    'CUSCO
    'Set RTasas = oLinea.RecuperaLineasTasas(psLineaCred)
    Call oLinea.CargarDetalleLineaCredito(psLineaCred, RTasas, RAgencias)
    Do While Not RTasas.EOF
        FETasas.AdicionaFila
        FETasas.TextMatrix(RTasas.Bookmark, 1) = RTasas!cConsDescripcion & Space(50) & Trim(str(CInt(Trim(RTasas!nColocLinCredTasaTpo))))
        FETasas.TextMatrix(RTasas.Bookmark, 2) = Format(RTasas!nTasaIni, "#0.0000")
        FETasas.TextMatrix(RTasas.Bookmark, 3) = Format(RTasas!nTasafin, "#0.0000")
        RTasas.MoveNext
    Loop
    RTasas.Close
    Set RTasas = Nothing
    If Len(Trim(FETasas.TextMatrix(1, 1))) > 0 Then
        FETasas.Row = 1
    End If
    'CUSCO
    chkTodos.value = 0
    Call chkTodos_Click
    With RAgencias
        Do While Not .EOF
            For i = 0 To lstAgencias.ListCount - 1
                If !cAgeCod = Left(lstAgencias.List(i), 2) Then
                    lstAgencias.Selected(i) = True
                End If
            Next i
            .MoveNext
        Loop
    End With

End Sub

Private Function NuevoCodigo(ByVal pbFondo As Boolean) As String
Dim i As Integer
Dim nMay As Integer
    nMay = 0
    If pbFondo Then
        For i = 0 To CmbFondo.ListCount - 2
            If CInt(Mid(CmbFondo.List(i), 101, 2)) > nMay Then
                nMay = CInt(Mid(CmbFondo.List(i), 101, 2))
            End If
        Next i
    Else
        For i = 0 To CmbSubFondo.ListCount - 2
            If CInt(Mid(CmbSubFondo.List(i), 103, 2)) > nMay Then
                nMay = CInt(Mid(CmbSubFondo.List(i), 103, 2))
            End If
        Next i
    End If
    NuevoCodigo = Format$(Trim(str(nMay + 1)), "00")
End Function

Private Sub chkTodos_Click()
Dim i As Integer
Dim bTodos As Boolean

bTodos = IIf(chkTodos.value = 1, True, False)

For i = 0 To lstAgencias.ListCount - 1
    lstAgencias.Selected(i) = bTodos
Next

End Sub

Private Sub CmbEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtPlazoMin.SetFocus
    End If
End Sub

Private Sub CmbFondo_Click()
Dim sCad As String
Dim nPos As Integer
    If Right(CmbFondo.Text, 3) = "new" Then
        nPos = CmbFondo.ListIndex
        sCad = frmCredLineaConcepto.Fondo(NuevoCodigo(True), sPersCodTemp)
        If Trim(sCad) <> "" Then
            CmbFondo.AddItem sCad
            'CmbFondo.ListIndex = CmbFondo.ListCount - 1
            CmbFondo.Tag = CmbFondo.ListIndex
            CmbFondo.RemoveItem (nPos)
            CmbSubFondo.Clear
        End If
    End If
    LblLineaDesc.Caption = Mid(CmbFondo.Text, 101, 2) & IIf(Mid(CmbSubFondo.Text, 103, 2) = "", Mid(CmbSubFondo.Text, 101, 2), Mid(CmbSubFondo.Text, 103, 2)) & Trim(Mid(CmbMoneda.Text, 101, 1)) & Trim(Mid(CmbPlazo.Text, 101, 1)) & Trim(Mid(CmbProd.Text, 101, 3))
    LblLinea.Caption = Trim(Right(CmbFondo.Text, 5)) & "-" & Trim(Right(CmbSubFondo.Text, 5)) & "-" & Trim(Right(CmbMoneda.Text, 2)) & "-" & Trim(Right(CmbPlazo.Text, 2)) & "-" & Trim(Right(CmbProd.Text, 7))
    Call CargaSubFondos(Trim(Mid(CmbFondo.Text, 101, 4)))
    
End Sub

Private Sub CmbFondo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbSubFondo.SetFocus
    End If
End Sub

Private Sub cmbMoneda_Click()
    LblLineaDesc.Caption = Mid(CmbFondo.Text, 101, 2) & IIf(Trim(Mid(CmbSubFondo.Text, 103, 2)) = "", Mid(CmbSubFondo.Text, 101, 2), Mid(CmbSubFondo.Text, 103, 2)) & Trim(Mid(CmbMoneda.Text, 101, 1)) & Trim(Mid(CmbPlazo.Text, 101, 1)) & Trim(Mid(CmbProd.Text, 101, 3))
    LblLinea.Caption = Trim(Right(CmbFondo.Text, 5)) & "-" & Trim(Right(CmbSubFondo.Text, 5)) & "-" & Trim(Right(CmbMoneda.Text, 2)) & "-" & Trim(Right(CmbPlazo.Text, 2)) & "-" & Trim(Right(CmbProd.Text, 7))
End Sub

Private Sub CmbMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbPlazo.SetFocus
    End If
End Sub

Private Sub cmbPaquete_Click()
Dim sCad As String
Dim nPos As Integer
Dim oLinea As COMDCredito.DCOMLineaCredito
    If Right(cmbPaquete.Text, 3) = "new" Then
        nPos = cmbPaquete.ListIndex
        Set oLinea = New COMDCredito.DCOMLineaCredito
        'sCad = oLinea.CorrelativoPaquete(LblLineaDesc.Caption)
        sCad = oLinea.CorrelativoPaquete(Mid(LblLineaDesc.Caption, 1, 4))
        Set oLinea = Nothing
        cmbPaquete.AddItem sCad
        cmbPaquete.ListIndex = cmbPaquete.ListCount - 1
        cmbPaquete.RemoveItem (nPos)
    End If
    LblLineaDesc.Caption = Mid(CmbFondo.Text, 101, 2) & IIf(Trim(Mid(CmbSubFondo.Text, 103, 2)) = "", Mid(CmbSubFondo.Text, 101, 2), Mid(CmbSubFondo.Text, 103, 2)) & Trim(Mid(CmbMoneda.Text, 101, 1)) & Trim(Mid(CmbPlazo.Text, 101, 1)) & Trim(Mid(CmbProd.Text, 101, 3)) & Trim(Left(cmbPaquete.Text, 2))
End Sub

Private Sub cmbPaquete_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   CmbEstado.SetFocus
End If
End Sub

Private Sub CmbPlazo_Click()
    LblLineaDesc.Caption = Mid(CmbFondo.Text, 101, 2) & IIf(Trim(Mid(CmbSubFondo.Text, 103, 2)) = "", Mid(CmbSubFondo.Text, 101, 2), Mid(CmbSubFondo.Text, 103, 2)) & Trim(Mid(CmbMoneda.Text, 101, 1)) & Trim(Mid(CmbPlazo.Text, 101, 1)) & Trim(Mid(CmbProd.Text, 101, 3))
    LblLinea.Caption = Trim(Right(CmbFondo.Text, 5)) & "-" & Trim(Right(CmbSubFondo.Text, 5)) & "-" & Trim(Right(CmbMoneda.Text, 2)) & "-" & Trim(Right(CmbPlazo.Text, 2)) & "-" & Trim(Right(CmbProd.Text, 7))
End Sub

Private Sub CmbPlazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbProd.SetFocus
    End If
End Sub

Private Sub CmbProd_Click()
    LblLineaDesc.Caption = Mid(CmbFondo.Text, 101, 2) & IIf(Trim(Mid(CmbSubFondo.Text, 103, 2)) = "", Mid(CmbSubFondo.Text, 101, 2), Mid(CmbSubFondo.Text, 103, 2)) & Trim(Mid(CmbMoneda.Text, 101, 1)) & Trim(Mid(CmbPlazo.Text, 101, 1)) & Trim(Mid(CmbProd.Text, 101, 3))
    LblLinea.Caption = Trim(Right(CmbFondo.Text, 5)) & "-" & Trim(Right(CmbSubFondo.Text, 5)) & "-" & Trim(Right(CmbMoneda.Text, 2)) & "-" & Trim(Right(CmbPlazo.Text, 2)) & "-" & Trim(Right(CmbProd.Text, 7))
    'Nueva Opcion de Paquetes
    'Call CargarPaquetes(LblLineaDesc.Caption)
End Sub

Private Sub CargarPaquetes(ByVal psFiltro As String)
Dim oLinea As COMDCredito.DCOMLineaCredito
Dim R As ADODB.Recordset
Dim sFiltro As String

    'Carga Paquetes de Linea de Credito
    cmbPaquete.Clear
    sFiltro = ""
    Set oLinea = New COMDCredito.DCOMLineaCredito
    Set R = oLinea.RecuperaPaquetes(psFiltro)
    Do While Not R.EOF
        cmbPaquete.AddItem Trim(R!Paquete)
        R.MoveNext
    Loop
    R.Close
    cmbPaquete.AddItem "<Nuevo Paquete>" & Space(100) & "new"
    Set oLinea = Nothing
End Sub

Private Sub CmbProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbPaquete.SetFocus
    End If
End Sub

Private Sub CmbSubFondo_Click()
Dim sCad As String
Dim nPos As Integer

    If Right(CmbSubFondo.Text, 3) = "new" Then
        nPos = CmbSubFondo.ListIndex
        sCad = frmCredLineaConcepto.SubFondo(NuevoCodigo(False))
        If Trim(sCad) <> "" Then
            CmbSubFondo.AddItem sCad
            CmbSubFondo.ListIndex = CmbSubFondo.ListCount - 1
            Call CmbSubFondo.RemoveItem(nPos)
        End If
    End If
    
    LblLineaDesc.Caption = Mid(CmbFondo.Text, 101, 2) & IIf(Trim(Mid(CmbSubFondo.Text, 103, 2)) = "", Mid(CmbSubFondo.Text, 101, 2), Mid(CmbSubFondo.Text, 103, 2)) & Trim(Mid(CmbMoneda.Text, 101, 1)) & Trim(Mid(CmbPlazo.Text, 101, 1)) & Trim(Mid(CmbProd.Text, 101, 3))
    LblLinea.Caption = Trim(Right(CmbFondo.Text, 5)) & "-" & Trim(Right(CmbSubFondo.Text, 5)) & "-" & Trim(Right(CmbMoneda.Text, 2)) & "-" & Trim(Right(CmbPlazo.Text, 2)) & "-" & Trim(Right(CmbProd.Text, 7))
    
    '03-05
    Call CargarPaquetes(LblLineaDesc.Caption)
End Sub

Private Sub CmbSubFondo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbMoneda.SetFocus
    End If
End Sub

Private Sub CmdAceptar_Click()
Dim oLineaCredito As COMNCredito.NCOMLineaCredito
Dim dLineaCredito As COMDCredito.DCOMLineaCredito
Dim sError As String
    
    On Error GoTo ErrorCmdAceptar_Click
    
    'Manejo de Lineas de Credito por Agencias
    If RecuperaAgencias = False Then Exit Sub
    '********************

    Set dLineaCredito = New COMDCredito.DCOMLineaCredito
    Set oLineaCredito = New COMNCredito.NCOMLineaCredito
    
    If nTipoProceso = 1 Then
        If CInt(CmbFondo.Tag) = CmbFondo.ListIndex Then
            sError = oLineaCredito.NuevaLinea(5, LblLineaDesc.Caption, LblLinea.Caption, _
                    IIf(CmbEstado.ListIndex = 0, 1, 0), CDbl(TxtPlazoMax.Text), CDbl(TxtPlazoMin.Text), CDbl(TxtMontoMax.Text), _
                    CDbl(TxtMontoMin.Text), sPersCodTemp, Trim(Left(CmbFondo.Text, 100)), Trim(Left(CmbSubFondo.Text, 100)), Trim(Left(CmbProd.Text, 100)), Trim(Right(CmbFondo.Text, 5)), Trim(Right(CmbSubFondo.Text, 5)), IIf(ChkPreferencial.value = vbChecked, True, False), MatAgencias)
        Else
            sError = oLineaCredito.NuevaLinea(5, LblLineaDesc.Caption, LblLinea.Caption, _
                    IIf(CmbEstado.ListIndex = 0, 1, 0), CDbl(TxtPlazoMax.Text), CDbl(TxtPlazoMin.Text), CDbl(TxtMontoMax.Text), _
                    CDbl(TxtMontoMin.Text), dLineaCredito.RecuperaInstitucion(LblLineaDesc.Caption), Trim(Left(CmbFondo.Text, 100)), Trim(Left(CmbSubFondo.Text, 100)), Trim(Left(CmbProd.Text, 100)), , Trim(Right(CmbSubFondo.Text, 20)), IIf(ChkPreferencial.value = vbChecked, True, False), MatAgencias)
        End If
        objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Linea Credito", LblLineaDesc.Caption, gCodLineaCred
    Else
        If nTipoProceso = 2 Then
            sError = oLineaCredito.ModificarLinea(5, LblLineaDesc.Caption, LblLinea.Caption, _
                IIf(CmbEstado.ListIndex = 0, 1, 0), CDbl(TxtPlazoMax.Text), CDbl(TxtPlazoMin.Text), CDbl(TxtMontoMax.Text), _
                CDbl(TxtMontoMin.Text), dLineaCredito.RecuperaInstitucion(LblLineaDesc.Caption), IIf(ChkPreferencial.value = vbChecked, True, False), MatAgencias)
            objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Linea Credito", LblLineaDesc.Caption, gCodLineaCred
        End If
    End If
    Set oLineaCredito = Nothing
    Set dLineaCredito = Nothing
    If sError <> "" Then
        MsgBox sError, vbInformation, "Aviso"
        Exit Sub
    End If
    
    Call CargaControles
    Call CargaLineasCredito(TxtLineaBusq.Text)
    Call cmdcancelar_Click
    
    Exit Sub

ErrorCmdAceptar_Click:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub cmdcancelar_Click()
    nTipoProceso = 0
    HabilitaCombos False
    HabilitaBusqueda True
    HabilitaDatosLinea False
    HabilitaBotones False
    LimpiaFlex FETasas
    Call CargaControles
    LblLinea.Caption = ""
    LblLineaDesc.Caption = ""
    TxtLinDesc.Text = ""
    TxtPlazoMin.Text = "0"
    TxtPlazoMax.Text = "0"
    TxtMontoMin.Text = "0.00"
    TxtMontoMax.Text = "0.00"
    CmbEstado.ListIndex = -1
    CmbSubFondo.Clear
End Sub

Private Sub cmdEditar_Click()
    nTipoProceso = 2
    bAplDes = False
    HabilitaCombos False
    HabilitaBusqueda False
    HabilitaDatosLinea True
    HabilitaBotones True
    SSTablinea.Tab = 1
End Sub

Private Sub cmdeliminar_Click()
Dim oLinea As COMNCredito.NCOMLineaCredito
Dim sError As String
    If LstLineas.ListItems.Count <= 0 Then
        MsgBox "No Existen Lineas que Eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Se va a Eliminar la Linea de Credito : " & LstLineas.SelectedItem.Text, vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oLinea = New COMNCredito.NCOMLineaCredito
        sError = oLinea.EliminaLinea(LstLineas.SelectedItem.Text)
        objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gEliminar, "Linea Credito", (LstLineas.SelectedItem.Text), gCodLineaCred
        If sError <> "" Then
            MsgBox sError, vbInformation, "Aviso"
            Exit Sub
        End If
        Set oLinea = Nothing
        Call CargaLineasCredito(TxtLineaBusq.Text)
    End If
End Sub

Private Sub cmdNuevo_Click()
Dim i As Integer

    nTipoProceso = 1
    HabilitaCombos True
    HabilitaBusqueda False
    HabilitaDatosLinea True
    HabilitaBotones True
    
    CmbFondo.ListIndex = -1
    CmbSubFondo.ListIndex = -1
    CmbMoneda.ListIndex = -1
    CmbPlazo.ListIndex = -1
    CmbProd.ListIndex = -1
    CmbEstado.ListIndex = -1
    TxtMontoMax.Text = "0.00"
    TxtMontoMin.Text = "0.00"
    TxtPlazoMax.Text = "0"
    TxtPlazoMin.Text = "0"
    SSTablinea.Tab = 1
    CmbFondo.Tag = "-2"
    
    LimpiaFlex FETasas
    
    chkTodos.value = 0
    Call chkTodos_Click
    For i = 0 To lstAgencias.ListCount - 1
        If gsCodAge = Left(lstAgencias.List(i), 2) Then
            lstAgencias.Selected(i) = True
        End If
    Next i
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdTarifario_Click()
    FrmCredConsultaLineaCred.Show 1
End Sub

Private Sub CmdTasaEditar_Click()
Dim ColBloq(1) As Integer
    If Trim(FETasas.TextMatrix(1, 1)) = "" Then
        MsgBox "No Existen datos a Editar", vbInformation, "Aviso"
        Exit Sub
    End If
    Call HabilitaIngresoLineaTasa(True)
    ColBloq(0) = 1
    sColBloq = FETasas.ColumnasAEditar
    Call HabilitaFilaFlex(nFilaLinea, FETasas, ColBloq, &HC0FFFF, True)
    nTipoProcesoTasa = 2
End Sub

Private Sub CmdTasaEliminar_Click()
Dim oLinea As COMNCredito.NCOMLineaCredito
Dim sError As String
    If Trim(FETasas.TextMatrix(1, 1)) = "" Then
        MsgBox "No Existen datos a Eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Se va a Eliminar la Tasa : " & Trim(Left(FETasas.TextMatrix(FETasas.Row, 1), 50)) & " de la  Linea de Credito : " & LstLineas.SelectedItem.Text, vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oLinea = New COMNCredito.NCOMLineaCredito
        sError = oLinea.EliminarTasa(LstLineas.SelectedItem.Text, Trim(Right(FETasas.TextMatrix(FETasas.Row, 1), 20)))
        If sError <> "" Then
            MsgBox sError, vbInformation, "Aviso"
            Exit Sub
        End If
        Set oLinea = Nothing
        Call FETasas.EliminaFila(FETasas.Row)
    End If
End Sub


Private Sub CmdTasaNuevo_Click()
Dim ColBloq(0) As Integer
    
    Call HabilitaIngresoLineaTasa(True)
    sColBloq = FETasas.ColumnasAEditar
    Call HabilitaFilaFlex(nFilaLinea, FETasas, ColBloq, &HC0FFFF)
    FETasas.TextMatrix(nFilaLinea, 2) = "0.00"
    FETasas.TextMatrix(nFilaLinea, 3) = "0.00"
    FETasas.TopRow = nFilaLinea
    nTipoProcesoTasa = 1
End Sub

Private Sub CmdTasasAceptar_Click()
Dim oNLinea As COMNCredito.NCOMLineaCredito
Dim sError As String
    On Error GoTo ErrorCmdTasasAceptar_Click
    Set oNLinea = New COMNCredito.NCOMLineaCredito
    If nTipoProcesoTasa = 1 Then
        sError = oNLinea.NuevaTasa(LstLineas.SelectedItem.Text, Trim(Right(FETasas.TextMatrix(nFilaLinea, 1), 15)), CDbl(FETasas.TextMatrix(nFilaLinea, 2)), CDbl(FETasas.TextMatrix(nFilaLinea, 3)))
    Else
        If nTipoProcesoTasa = 2 Then
            sError = oNLinea.ModificarTasa(LstLineas.SelectedItem.Text, Trim(Right(FETasas.TextMatrix(nFilaLinea, 1), 15)), CDbl(FETasas.TextMatrix(nFilaLinea, 2)), CDbl(FETasas.TextMatrix(nFilaLinea, 3)))
        End If
    End If
    
    If sError <> "" Then
        MsgBox sError, vbInformation, "Aviso"
        Exit Sub
    End If
    Call CmdTasasCancelar_Click
    If Len(Trim(LstLineas.SelectedItem.Text)) <> "" Then
        Call CargaDetalleLinea(Trim(LstLineas.SelectedItem.Text))
    End If
    Exit Sub

ErrorCmdTasasAceptar_Click:
    MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub CmdTasasCancelar_Click()
    Call HabilitaFlexNormal(FETasas, nFilaLinea, sColBloq)
    If nTipoProcesoTasa = 1 Then
        Call FETasas.EliminaFila(nFilaLinea)
    End If
    Call HabilitaIngresoLineaTasa(False)
    nFilaLinea = -1
End Sub

Private Sub FETasas_RowColChange()
    If nFilaLinea <> -1 Then
        FETasas.Row = nFilaLinea
    End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    FETasas.lbEditarFlex = False
    nFilaLinea = -1
    nTipoProcesoTasa = -1
    nTipoProceso = -1
    Call HabilitaDatosLinea(False)
    Call CargaControles
    Call CargaLineasCredito("0101")
    TxtLineaBusq.Text = "0101"
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredLineaCred
End Sub

Private Sub LstLineas_Click()
Dim sLineDescTmp As String
Dim i As Integer
Dim nCont As Integer
Dim k As Integer
    If LstLineas.ListItems.Count <= 0 Then
        Exit Sub
    End If
    
    LblLineaDesc.Caption = Trim(LstLineas.SelectedItem.Text)
    sLineDescTmp = LblLineaDesc.Caption
    
    CmbFondo.ListIndex = IndiceListaCombo(CmbFondo, Mid(sLineDescTmp, 1, 2), 2)
    CmbSubFondo.ListIndex = IndiceListaCombo(CmbSubFondo, Mid(sLineDescTmp, 1, 4), 2)
    CmbMoneda.ListIndex = IndiceListaCombo(CmbMoneda, Mid(sLineDescTmp, 5, 1), 2)
    CmbPlazo.ListIndex = IndiceListaCombo(CmbPlazo, Mid(sLineDescTmp, 6, 1), 2)
    CmbProd.ListIndex = IndiceListaCombo(CmbProd, Mid(sLineDescTmp, 7, 3), 2)
    'Paquete
    Call CmbProd_Click
    cmbPaquete.ListIndex = IndiceListaCombo(cmbPaquete, Mid(sLineDescTmp, 10, 2), 1)
    If Trim(LstLineas.SelectedItem.SubItems(2)) = "1" Then
        CmbEstado.ListIndex = 0
    Else
        CmbEstado.ListIndex = 1
    End If
    
    LblLinea.Caption = LstLineas.SelectedItem.SubItems(1)
    
    TxtPlazoMin.Text = LstLineas.SelectedItem.SubItems(3)
    TxtPlazoMax.Text = LstLineas.SelectedItem.SubItems(4)
    TxtMontoMin.Text = LstLineas.SelectedItem.SubItems(5)
    TxtMontoMax.Text = LstLineas.SelectedItem.SubItems(6)
    
    ObtenerPreferencial LstLineas.SelectedItem.Text
    
    Call CargaDetalleLinea(LstLineas.SelectedItem.Text)
    CmdTasaNuevo.Enabled = True
    CmdTasaEliminar.Enabled = True
    CmdTasaEditar.Enabled = True
    
    LblLineaDesc.Caption = sLineDescTmp
    nCont = 0
    For i = 1 To Len(LblLinea.Caption)
        If Mid(LblLinea.Caption, i, 1) = "-" Then
           nCont = nCont + 1
        End If
    Next i
    If nCont = 6 Then
        k = 0
        For i = Len(LblLinea.Caption) To 1 Step -1
            If Mid(LblLinea.Caption, i, 1) = "-" Then
                k = i
                Exit For
            End If
        Next i
    End If
    
    If k <> 0 Then
       TxtLinDesc.Text = Mid(LblLinea.Caption, k + 1, Len(LblLinea.Caption))
    End If
    
End Sub


Private Sub LstLineas_KeyUp(KeyCode As Integer, Shift As Integer)
    Call LstLineas_Click
End Sub


Function RecuperaAgencias() As Boolean

Dim nContAge As Integer
Dim i As Integer
    
ReDim MatAgencias(0)
nContAge = 0
RecuperaAgencias = True

For i = 0 To lstAgencias.ListCount - 1
    If lstAgencias.Selected(i) = True Then
        nContAge = nContAge + 1
        ReDim Preserve MatAgencias(nContAge)
        MatAgencias(nContAge - 1) = Mid(lstAgencias.List(i), 1, 2)
    End If
Next i
    
If nContAge = 0 Then
    MsgBox "Debe seleccionar por lo menos una Agencia", vbInformation, "Mensaje"
    RecuperaAgencias = False
End If
    
End Function

Private Sub optElejir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SSTablinea.Tab = 1
        CmbEstado.SetFocus
    End If
End Sub

Private Sub TxtLinDesc_GotFocus()
    fEnfoque TxtLinDesc
End Sub

Private Sub TxtLinDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        CmdAceptar.SetFocus
    End If
End Sub

Private Sub TxtLinDesc_LostFocus()
Dim i As Integer
Dim k As Integer
Dim nCont As Integer

    nCont = 0
    For i = 1 To Len(LblLinea.Caption)
        If Mid(LblLinea.Caption, i, 1) = "-" Then
           nCont = nCont + 1
        End If
    Next i
    bAplDes = False
    If nCont = 6 Then
        bAplDes = True
    End If

    If bAplDes Then
        k = 0
        For i = Len(LblLinea.Caption) To 1 Step -1
            If Mid(LblLinea.Caption, i, 1) = "-" Then
                k = i
                Exit For
            End If
        Next i
        If k <> 0 Then
            LblLinea.Caption = Mid(LblLinea.Caption, 1, k - 1) & "-" & Trim(TxtLinDesc.Text)
        End If
    Else
        bAplDes = True
        LblLinea.Caption = LblLinea.Caption & "-" & Trim(TxtLinDesc.Text)
    End If
End Sub

Private Sub TxtLineaBusq_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        Call cmdcancelar_Click
        Call CargaLineasCredito(TxtLineaBusq.Text)
    End If
End Sub

Private Sub txtMontoMax_GotFocus()
    fEnfoque TxtMontoMax
End Sub

Private Sub txtMontoMax_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMontoMax, KeyAscii)
    If KeyAscii = 13 Then
        Label11.Enabled = True
        TxtLinDesc.Enabled = True
        TxtLinDesc.SetFocus
    End If
End Sub

Private Sub txtMontoMax_LostFocus()
    If Trim(TxtMontoMax.Text) = "" Then
        TxtMontoMax.Text = "0.00"
    Else
        TxtMontoMax.Text = Format(TxtMontoMax.Text, "0.00")
    End If
End Sub

Private Sub txtMontoMin_GotFocus()
    fEnfoque TxtMontoMin
End Sub

Private Sub txtMontoMin_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMontoMin, KeyAscii)
    If KeyAscii = 13 Then
        TxtMontoMax.SetFocus
    End If
End Sub

Private Sub txtMontoMin_LostFocus()
    If Trim(TxtMontoMin.Text) = "" Then
        TxtMontoMin.Text = "0.00"
    Else
        TxtMontoMin.Text = Format(TxtMontoMin.Text, "0.00")
    End If
End Sub

Private Sub TxtPlazoMax_GotFocus()
    fEnfoque TxtPlazoMax
End Sub

Private Sub TxtPlazoMax_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        TxtMontoMin.SetFocus
    End If
End Sub

Private Sub TxtPlazoMin_GotFocus()
    fEnfoque TxtPlazoMin
End Sub

Private Sub TxtPlazoMin_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        TxtPlazoMax.SetFocus
    End If
End Sub

Sub ObtenerPreferencial(ByVal pcLineaCredito As String)
    Dim objDLineaCredito As COMDCredito.DCOMLineaCredito
    Dim bValor As Boolean
    
    On Error GoTo ErrHandler
        Set objDLineaCredito = New COMDCredito.DCOMLineaCredito
        bValor = objDLineaCredito.ObtenerPreferencialLinea(pcLineaCredito)
        Set objDLineaCredito = Nothing
        If bValor = True Then
            ChkPreferencial.value = vbChecked
        Else
            ChkPreferencial.value = vbUnchecked
        End If
    Exit Sub
ErrHandler:
    If objDLineaCredito Is Nothing Then Set objDLineaCredito = Nothing
    MsgBox "Error al obtener linea preferencial", vbInformation, "AVISO"
End Sub

