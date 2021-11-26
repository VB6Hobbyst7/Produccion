VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmColocEvalCalCli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocaciones - Evaluación y Calificación de Clientes"
   ClientHeight    =   7485
   ClientLeft      =   1110
   ClientTop       =   1185
   ClientWidth     =   9915
   Icon            =   "frmColocEvalCalCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imgList 
      Left            =   11160
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   42
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColocEvalCalCli.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColocEvalCalCli.frx":0ADC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   600
      Left            =   240
      TabIndex        =   43
      Top             =   6480
      Width           =   9240
      Begin VB.CommandButton cmdEliminaGen 
         Caption         =   "&Eliminar"
         Height          =   360
         Left            =   2925
         TabIndex        =   9
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir Planilla"
         Height          =   360
         Left            =   4200
         TabIndex        =   10
         Top             =   180
         Width           =   1815
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7635
         TabIndex        =   44
         Top             =   180
         Width           =   1470
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   360
         Left            =   1455
         TabIndex        =   8
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   360
         Left            =   720
         TabIndex        =   7
         Top             =   120
         Width           =   1275
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   360
         Left            =   180
         TabIndex        =   11
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   360
         Left            =   1455
         TabIndex        =   12
         Top             =   180
         Width           =   1275
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   690
      Left            =   10320
      TabIndex        =   29
      Top             =   7200
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   1217
      _Version        =   393217
      TextRTF         =   $"frmColocEvalCalCli.frx":12AE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   165
      Left            =   6060
      TabIndex        =   41
      Top             =   7290
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin TabDlg.SSTab TabGen 
      Height          =   7185
      Left            =   60
      TabIndex        =   30
      Top             =   0
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   12674
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Lista de Créditos"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblCodPers1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblNomPers1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Shape1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Shape2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "TabPosicion"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdBuscar"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.Frame Frame4 
         Height          =   2055
         Left            =   -67545
         TabIndex        =   42
         Top             =   1845
         Width           =   1815
         Begin MSMask.MaskEdBox mskFechaMes 
            Height          =   315
            Left            =   585
            TabIndex        =   50
            Top             =   1185
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmdImprimeCreditos 
            Caption         =   "Im&primir"
            Height          =   345
            Left            =   180
            TabIndex        =   20
            Top             =   1590
            Width           =   1455
         End
         Begin VB.CheckBox chkImpFte 
            Caption         =   "&Fuentes"
            Height          =   300
            Left            =   300
            TabIndex        =   19
            Top             =   720
            Width           =   990
         End
         Begin VB.CheckBox chkImpGar 
            Caption         =   "&Garantías"
            Height          =   345
            Left            =   300
            TabIndex        =   18
            Top             =   420
            Width           =   1125
         End
         Begin VB.CheckBox chkImpCred 
            Caption         =   "&Créditos"
            Height          =   330
            Left            =   300
            TabIndex        =   17
            Top             =   165
            Width           =   900
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha :"
            Height          =   195
            Left            =   75
            TabIndex        =   49
            Top             =   1215
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar Datos Clientes"
         Height          =   345
         Left            =   -70680
         TabIndex        =   13
         Top             =   540
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos del Cliente"
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
         Left            =   195
         TabIndex        =   27
         Top             =   525
         Width           =   9450
         Begin VB.ListBox lsActividades 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            ItemData        =   "frmColocEvalCalCli.frx":132E
            Left            =   4920
            List            =   "frmColocEvalCalCli.frx":1335
            Sorted          =   -1  'True
            TabIndex        =   48
            Top             =   615
            Width           =   4245
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cliente "
            Height          =   195
            Left            =   165
            TabIndex        =   36
            Top             =   285
            Width           =   525
         End
         Begin VB.Label lblDocJur 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   3600
            TabIndex        =   3
            Top             =   615
            Width           =   1275
         End
         Begin VB.Label lblDocNat 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1200
            TabIndex        =   2
            Top             =   615
            Width           =   1155
         End
         Begin VB.Label lblNomPers 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   2400
            TabIndex        =   1
            Top             =   225
            Width           =   6810
         End
         Begin VB.Label lblCodPers 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   1200
            TabIndex        =   0
            Top             =   225
            Width           =   1170
         End
         Begin VB.Label lblDocJuridico 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Tributario"
            Height          =   195
            Left            =   2400
            TabIndex        =   35
            Top             =   690
            Width           =   1050
         End
         Begin VB.Label lblDocNatural 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Ident."
            Height          =   195
            Left            =   150
            TabIndex        =   34
            Top             =   675
            Width           =   795
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos de Auditoria"
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
         Height          =   4725
         Left            =   165
         TabIndex        =   33
         Top             =   1710
         Width           =   9240
         Begin MSAdodcLib.Adodc adoAudGen 
            Height          =   375
            Left            =   4500
            Top             =   1845
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSComDlg.CommonDialog cdlOpen 
            Left            =   7740
            Top             =   4140
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton cmdPlantilla 
            Caption         =   "&Plantilla"
            Height          =   360
            Left            =   7575
            TabIndex        =   46
            Top             =   3600
            Width           =   1320
         End
         Begin VB.CommandButton cmdPasar 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   8745
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   4335
            Width           =   420
         End
         Begin VB.TextBox txtObs 
            Height          =   1755
            Left            =   345
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   2820
            Width           =   6960
         End
         Begin MSDataGridLib.DataGrid dtgAudGen 
            Height          =   2070
            Left            =   240
            TabIndex        =   4
            Top             =   285
            Width           =   8715
            _ExtentX        =   15372
            _ExtentY        =   3651
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   2
            RowHeight       =   20
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "persona"
               Caption         =   "Persona"
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
               DataField       =   "Cal"
               Caption         =   "Calific."
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
               DataField       =   "Obs"
               Caption         =   "Observaciones"
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
               DataField       =   "CODPERS"
               Caption         =   "ccodpers"
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
               DataField       =   "DocNat"
               Caption         =   "docIdent"
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
               DataField       =   "DocTri"
               Caption         =   "docTrib"
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
               BeginProperty Column00 
                  ColumnWidth     =   6944.882
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  ColumnWidth     =   0
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   0
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   0
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  ColumnWidth     =   0
               EndProperty
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones Generales :"
            Height          =   195
            Left            =   150
            TabIndex        =   45
            Top             =   2565
            Width           =   1920
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Calificación"
            Height          =   195
            Left            =   7770
            TabIndex        =   40
            Top             =   2700
            Width           =   810
         End
         Begin VB.Label txtCalGen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   7860
            TabIndex        =   39
            ToolTipText     =   "Doble Clic para actualizar"
            Top             =   2985
            Width           =   615
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            X1              =   10365
            X2              =   0
            Y1              =   2475
            Y2              =   2475
         End
         Begin VB.Line Line1 
            X1              =   10320
            X2              =   15
            Y1              =   2490
            Y2              =   2490
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Calificación por Crédito"
         Height          =   2400
         Left            =   -74835
         TabIndex        =   31
         Top             =   4080
         Width           =   9240
         Begin VB.CommandButton cmdImprimeCalCred 
            Caption         =   "I&mprimir Créditos"
            Height          =   345
            Left            =   4665
            TabIndex        =   25
            Top             =   1980
            Width           =   1575
         End
         Begin VB.CommandButton cmdEditaCred 
            Caption         =   "&Modificar"
            Height          =   345
            Left            =   1320
            TabIndex        =   23
            Top             =   1980
            Width           =   1110
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            Height          =   345
            Left            =   180
            TabIndex        =   22
            Top             =   1980
            Width           =   1095
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   345
            Left            =   3480
            TabIndex        =   24
            Top             =   1980
            Width           =   1140
         End
         Begin VB.TextBox txtObsDet 
            Height          =   2010
            Left            =   6450
            Locked          =   -1  'True
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Top             =   270
            Width           =   2610
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgCredito 
            Height          =   1695
            Left            =   180
            TabIndex        =   21
            ToolTipText     =   "Seleccione el credito para mostrar Observaciones al lado derecho"
            Top             =   210
            Width           =   6150
            _ExtentX        =   10848
            _ExtentY        =   2990
            _Version        =   393216
            Cols            =   5
            FocusRect       =   0
            HighLight       =   2
            RowSizingMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin TabDlg.SSTab TabPosicion 
         Height          =   2685
         Left            =   -74835
         TabIndex        =   32
         Top             =   1395
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4736
         _Version        =   393216
         Style           =   1
         TabHeight       =   617
         WordWrap        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&Lista de Créditos   "
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fgCreditos"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Garantías"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fgGarantias"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Fuentes de Ingreso       "
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkFuentes"
         Tab(2).Control(1)=   "fgFuentes"
         Tab(2).ControlCount=   2
         Begin VB.CheckBox chkFuentes 
            Height          =   195
            Left            =   -70980
            TabIndex        =   47
            Top             =   105
            Width           =   210
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgCreditos 
            Height          =   1965
            Left            =   195
            TabIndex        =   14
            ToolTipText     =   "Doble click para visualizar información detallada del crédito"
            Top             =   525
            Width           =   7050
            _ExtentX        =   12435
            _ExtentY        =   3466
            _Version        =   393216
            FocusRect       =   0
            HighLight       =   2
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgFuentes 
            Height          =   1965
            Left            =   -74805
            TabIndex        =   16
            Top             =   525
            Width           =   7050
            _ExtentX        =   12435
            _ExtentY        =   3466
            _Version        =   393216
            FocusRect       =   0
            HighLight       =   2
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgGarantias 
            Height          =   1965
            Left            =   -74805
            TabIndex        =   15
            Top             =   525
            Width           =   7050
            _ExtentX        =   12435
            _ExtentY        =   3466
            _Version        =   393216
            FocusRect       =   0
            HighLight       =   2
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000005&
         Height          =   930
         Left            =   -74865
         Shape           =   4  'Rounded Rectangle
         Top             =   450
         Width           =   6420
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000006&
         Height          =   930
         Left            =   -74865
         Shape           =   4  'Rounded Rectangle
         Top             =   435
         Width           =   6390
      End
      Begin VB.Label lblNomPers1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   -74700
         TabIndex        =   38
         Top             =   945
         Width           =   6090
      End
      Begin VB.Label lblCodPers1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -74700
         TabIndex        =   37
         Top             =   600
         Width           =   1050
      End
   End
   Begin MSComctlLib.StatusBar barraestado 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   28
      Top             =   7185
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8467
            MinWidth        =   8467
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmColocEvalCalCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'** Registro de Evaluacion de Personas
Option Explicit
Dim lnNumLineas As Integer
Dim lnNumPag As Integer

Dim fnTotalCred As Integer
Dim fnTotalGar As Integer
Dim fnTotalFte As Integer
Dim fbNuevo As Boolean

Dim lnPos As Integer
Dim lbAud As Boolean
Dim fdFechaFinMes As Date

Dim fnTipoCambio  As Currency
Dim fnMontoRango As Currency

'Private Sub adoAudGen_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'If Not pRecordset.EOF And Not pRecordset.BOF Then
'    lblCodPers = pRecordset!CodPers
'    Me.lblNomPers = PstaNombre(pRecordset!Persona, False)
'    lblCodPers1 = pRecordset!CodPers
'    Me.lblNomPers1 = PstaNombre(pRecordset!Persona, False)
'    Me.lblDocJur = pRecordset!DocTri
'    Me.lblDocNat = pRecordset!DocNat
'    Me.txtCalGen = pRecordset!Cal
'    Me.txtObs = Trim(pRecordset!Obs)
'    DatosDetalleBase Trim(pRecordset!CodPers)
'    Me.lsActividades.Clear
'    CabGridCreditos
'    CabGridGarantias
'    CabGridFuentes
'End If
'End Sub



Private Sub CmdAgregar_Click()
Dim lsCta As String
Dim ldFechaCred As Date
Dim lsFecha As String
If Me.lblCodPers = "" Then
    MsgBox "Cliente no válido ", vbInformation, "Aviso"
    Me.cmdBuscar.SetFocus
    Exit Sub
End If
If Me.fgCreditos.TextMatrix(1, 0) = "" Then
    MsgBox "Clientes no Posee Créditos", vbInformation, "Aviso"
    Me.cmdBuscar.SetFocus
    Exit Sub
End If
If Trim(fgCreditos.TextMatrix(fgCreditos.Row, 21)) <> "F" Then
        If Mid(Trim(fgCreditos.TextMatrix(fgCreditos.Row, 2)), 3, 3) = "305" Then
            If MsgBox("Crédito es un Pignoraticio. Desea continuar con el Proceso?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                Me.fgCreditos.SetFocus
                Exit Sub
            End If
        Else
            If MsgBox("Estado de Crédito no es Vigente. Desea continuar con el Proceso?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                Me.fgCreditos.SetFocus
                Exit Sub
            End If
        End If

End If

If Me.cmdnuevo.Visible Then
    'frmAuditDatos.lbNuevaPers = False
Else
    'frmAuditDatos.lbNuevaPers = True
End If
'frmAuditDatos.fbNuevo = True
'frmAuditDatos.lsCodCta = fgCreditos.TextMatrix(fgCreditos.Row, 2)
'frmAuditDatos.lblEstado = fgCreditos.TextMatrix(fgCreditos.Row, 5)
'If frmAuditDatos.lsCodCta = "" Then
'    MsgBox "Seleccione la fila del crédito que desea agregar ", vbInformation, "aviso"
'    Exit Sub
'End If
'frmAuditDatos.lsCodPers = Trim(lblCodPers)
'frmAuditDatos.TxtNota = IIf(fgCreditos.TextMatrix(fgCreditos.Row, 9) = "", "0", fgCreditos.TextMatrix(fgCreditos.Row, 9))
'frmAuditDatos.txtSaldoCap = fgCreditos.TextMatrix(fgCreditos.Row, 11)
'frmAuditDatos.txtDiasAtraso = Val(fgCreditos.TextMatrix(fgCreditos.Row, 13))

'frmAuditDatos.txtFechaFM = IIf(fgCreditos.TextMatrix(fgCreditos.Row, 22) = "", gdFecSis, fgCreditos.TextMatrix(fgCreditos.Row, 22))
'frmAuditDatos.txtNotaFM = IIf(fgCreditos.TextMatrix(fgCreditos.Row, 10) = "", fgCreditos.TextMatrix(fgCreditos.Row, 9), fgCreditos.TextMatrix(fgCreditos.Row, 10))
'frmAuditDatos.txtSaldoCapFM = IIf(fgCreditos.TextMatrix(fgCreditos.Row, 12) = "0.00", fgCreditos.TextMatrix(fgCreditos.Row, 11), fgCreditos.TextMatrix(fgCreditos.Row, 12))
'frmAuditDatos.txtDiasAtrasoFM = Val(IIf(fgCreditos.TextMatrix(fgCreditos.Row, 14) = "", fgCreditos.TextMatrix(fgCreditos.Row, 13), fgCreditos.TextMatrix(fgCreditos.Row, 14)))
'frmAuditDatos.chkVigente.Value = 1

'frmAuditDatos.lblCalxDiasAtraso.Caption = fgAsinaCalificacionCredito(fgCredito.TextMatrix(fgCredito.Row, 1), fgCredito.TextMatrix(fgCredito.Row, 9), fgCredito.TextMatrix(fgCredito.Row, 7), fgCredito.TextMatrix(fgCredito.Row, 3), fnMontoRango, fnTipoCambio, False)

'frmAuditDatos.Show 1

'If frmAuditDatos.lbOk Then
'    If frmAuditDatos.lbOk And frmAuditDatos.lbNuevaPers = False Then
'        DatosDetalleBase Trim(lblCodPers)
'    Else
'        DatosDetalleGrid Trim(lblCodPers), frmAuditDatos.lsCodCta, frmAuditDatos.ldFechaCal, _
'                    frmAuditDatos.lnSaldoCap, frmAuditDatos.lsNota, frmAuditDatos.lsCalificacion, frmAuditDatos.lsObs, frmAuditDatos.lnDiasAtraso, "0", Trim(frmAuditDatos.lsEstado)
'
'       txtObsDet = frmAuditDatos.lsObs
'        lsCta = ""
'        lsFecha = ""
'        For I = 1 To Me.fgCredito.Rows - 1
'            If (lsCta = Me.fgCredito.TextMatrix(I, 1) And lsFecha = Me.fgCredito.TextMatrix(I, 2)) Then
'                MsgBox "Credito ya ha sido ingresado a esa fecha", vbInformation, "Aviso"
'                EliminaRow Me.fgCredito, I
'            Else
'                lsCta = fgCredito.TextMatrix(I, 1)
'                lsFecha = fgCredito.TextMatrix(I, 2)
'            End If
'        Next I
'    End If
'End If
End Sub

Private Sub cmdBuscar_Click()
Dim lbVacio As Boolean
Dim J As Integer
lbVacio = False
If Me.lblCodPers = "" Then
    MsgBox "Cliente no Válido", vbInformation, "Aviso"
    Me.TabGen.Tab = 0
    Me.cmdnuevo.SetFocus
    Exit Sub
End If
'If ValFecha(Me.mskFechaMes.Text) = False Then
'    Exit Sub
'End If
'For J = 0 To listaAgencias.ListCount - 1
'    If listaAgencias.Selected(J) Then
'        lbVacio = True
'        Exit For
'    End If
'Next J
'If lbVacio Then
'    Conexiones Trim(lblCodPers)
'Else
'    MsgBox "Seleccione por lo menos una Agencia para poder empezar el Proceso", vbInformation, "Aviso"
'    listaAgencias.SetFocus
'End If

End Sub

Private Sub cmdCancelar_Click()
Habilitacion False

If fbNuevo Then
    LimpiarControles
    CabGridCreditos
    CabGridGarantias
    CabGridFuentes
End If
fbNuevo = False
Me.dtgAudGen.SetFocus
End Sub

Private Sub cmdEditaCred_Click()
If Me.fgCredito.TextMatrix(1, 0) = "" Then
    MsgBox "Lista no posee Créditos para modificar", vbInformation, "aviso"
    Exit Sub
End If
If Me.fgCredito.TextMatrix(fgCredito.Row, 8) = "." Then
    MsgBox "Credito ya ha sido Aprobado. No se puede modificar", vbInformation, "aviso"
    Exit Sub
End If


''frmAuditDatos.fbNuevo = False
''frmAuditDatos.lsCodPers = Trim(lblCodPers)
''frmAuditDatos.lsCodCta = fgCredito.TextMatrix(fgCredito.Row, 1)
''frmAuditDatos.txtFecha = fgCredito.TextMatrix(fgCredito.Row, 2)
''frmAuditDatos.txtSaldoCap = fgCredito.TextMatrix(fgCredito.Row, 3)
''frmAuditDatos.txtCalificacion = fgCredito.TextMatrix(fgCredito.Row, 4)
''frmAuditDatos.txtNota = fgCredito.TextMatrix(fgCredito.Row, 5)
''frmAuditDatos.txtObs = fgCredito.TextMatrix(fgCredito.Row, 6)
''frmAuditDatos.txtDiasAtraso = Val(fgCredito.TextMatrix(fgCredito.Row, 7))
''frmAuditDatos.chkVigente.Value = IIf(Trim(fgCredito.TextMatrix(fgCredito.Row, 10)) = "1", 1, 0)
''
''frmAuditDatos.fraDatosFM.Visible = False
''frmAuditDatos.optSeleccion(0).Visible = False
''frmAuditDatos.optSeleccion(1).Visible = False
''frmAuditDatos.FraDatosActual.Enabled = True
'''frmAuditDatos.lblCalxDiasAtraso.Caption = fgAsinaCalificacionCredito(fgCredito.TextMatrix(fgCredito.Row, 1),  fgCredito.TextMatrix(fgCredito.Row, 1), fgCredito.TextMatrix(fgCredito.Row, 7), fgCredito.TextMatrix(fgCredito.Row, 1), fnMontoRango, fnTipoCambio, False)
''
''frmAuditDatos.Show 1
''If frmAuditDatos.lbOk Then
''    If frmAuditDatos.lbNuevaPers Then
''        fgCredito.TextMatrix(fgCredito.Row, 1) = frmAuditDatos.lsCodCta
''        fgCredito.TextMatrix(fgCredito.Row, 2) = frmAuditDatos.ldFechaCal
''        fgCredito.TextMatrix(fgCredito.Row, 3) = frmAuditDatos.lnSaldoCap
''        fgCredito.TextMatrix(fgCredito.Row, 4) = frmAuditDatos.lsCalificacion
''        fgCredito.TextMatrix(fgCredito.Row, 5) = frmAuditDatos.lsNota
''        fgCredito.TextMatrix(fgCredito.Row, 6) = frmAuditDatos.lsObs
''    Else
''        If frmAuditDatos.fbNuevo = False Then
''            DatosDetalleBase Trim(lblCodPers)
''            txtObsDet = fgCredito.TextMatrix(fgCredito.Rows - 1, 6)
''
''        Else
''            DatosDetalleGrid Trim(lblCodPers), frmAuditDatos.lsCodCta, frmAuditDatos.txtFecha, _
''                frmAuditDatos.txtSaldoCap, frmAuditDatos.txtNota, frmAuditDatos.txtCalificacion, frmAuditDatos.txtObs, frmAuditDatos.lnDiasAtraso, "0", Trim(frmAuditDatos.lsEstado)
''        End If
''    End If
''End If

End Sub

Private Sub cmdEliminaGen_Click()
'Dim SQL As String
'Dim rs As New ADODB.Recordset
'If Me.lblCodPers = "" Then
'    MsgBox "Código de Cliente no válido para realizar operación", vbInformation, "Aviso"
'    Me.TabGen.Tab = 0
'    Me.dtgAudGen.SetFocus
'    Exit Sub
'End If
'
'If adoAudGen.Recordset.EOF Then
'    MsgBox "No existen registros para eliminar", vbInformation, "Aviso"
'    Me.cmdNuevo.SetFocus
'    Exit Sub
'End If
'
'SQL = "Select cCodCta from " & gcServerAudit & "AudPersDet where cCodPers ='" & lblCodPers & "'"
'rs.Open SQL, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
'If Not RSVacio(rs) Then
'    MsgBox "Cliente posee creditos Ingresados." & gPrnSaltoLinea & "1° Elimine los créditos del Cliente para poder realizar el proceso", vbInformation, "Aviso"
'    Me.TabGen.Tab = 1
'    If Me.cmdEliminar.Enabled Then
'        Me.cmdEliminar.SetFocus
'    End If
'    Exit Sub
'End If
'rs.Close
'Set rs = Nothing
'
'If MsgBox("Desea elminar el Registro Seleccionado", vbQuestion + vbYesNo, "Aviso") = vbYes Then
'    SQL = "DELETE " & gcServerAudit & "AUDPERS WHERE cCodPers='" & lblCodPers & "'"
'
'    dbCmact.Execute SQL
'    DatosGenerales
'    If adoAudGen.Recordset.EOF Then
'        LimpiarControles
'    End If
'End If
End Sub

Private Sub cmdeliminar_Click()
'Dim SQL As String
'Dim lsCodCta As String
'
'If fgCredito.TextMatrix(1, 0) = "" Then
'    MsgBox "No existen registros detallados para eliminar", vbInformation, "Aviso"
'    Me.cmdAgregar.SetFocus
'    Exit Sub
'End If
'If Me.fgCredito.TextMatrix(fgCredito.Row, 8) = "." Then
'    MsgBox "Credito ya ha sido Aprobado. No se puede Eliminar", vbInformation, "aviso"
'    Exit Sub
'End If
'If MsgBox("Desea elminar el Crédito :" & Me.fgCredito.TextMatrix(fgCredito.Row, 1) & " de la Fecha : " & fgCredito.TextMatrix(fgCredito.Row, 2), vbQuestion + vbYesNo, "Aviso") = vbYes Then
'    If Me.cmdNuevo.Visible Then
'        lsCodCta = Trim(fgCredito.TextMatrix(fgCredito.Row, 1))
'        SQL = "DELETE " & gcServerAudit & "AUDPERSDET WHERE cCodPers='" & lblCodPers & "' and cCodCta='" & lsCodCta & "' and dFecCal='" & Format(fgCredito.TextMatrix(fgCredito.Row, 2), "mm/dd/yyyy") & "'"
'
'        dbCmact.Execute SQL
'        DatosDetalleBase Trim(lblCodPers)
'    Else
'        EliminaRow fgCredito, fgCredito.Row
'    End If
'    If fgCredito.TextMatrix(1, 0) = "" Then
'        Me.txtObsDet = ""
'    End If
'End If
End Sub

Private Sub cmdGrabar_Click()
'Dim SQL As String
'Dim I As Integer
'Dim lsCalGen As String
'Dim lnCalAux As Integer
'Dim lsAut As String
'On Error GoTo ErrorGrabar
'
'If Valida = False Then Exit Sub
'If MsgBox("Desea Grabar la Información", vbYesNo + vbQuestion, "Aviso") = vbYes Then
'    If fbNuevo Then
'        For I = 1 To fgCredito.Rows - 1
'            lsAut = IIf(fgCredito.TextMatrix(I, 8) = ".", "1", "0")
'
'            SQL = "INSERT INTO " & gcServerAudit & "AUDPERSDET(cCodPers, cCodCta, dFecCal, nSaldoCap, cCalAud, cNotaAud, cObsAud,nDiasAtraso, cAutorizado,cEstado )  " _
'                & "VALUES('" & Trim(lblCodPers) & "','" & Trim(fgCredito.TextMatrix(I, 1)) & "','" & Format(fgCredito.TextMatrix(I, 2), "mm/dd/yyyy") & "'," _
'                & fgCredito.TextMatrix(I, 3) & ",'" & Trim(fgCredito.TextMatrix(I, 4)) & "','" & Trim(fgCredito.TextMatrix(I, 5)) & "'," _
'                & IIf(fgCredito.TextMatrix(I, 6) = "", "Null", "'" & Mid(Trim(fgCredito.TextMatrix(I, 6)), 1, 200) & "'") & "," & Trim(fgCredito.TextMatrix(I, 7)) & ",'" & lsAut & "','" & Trim(fgCredito.TextMatrix(I, 10)) & "')"
'
'            dbCmact.Execute SQL
'        Next
'
'        lsCalGen = CalificaGen(Trim(lblCodPers), dbCmact)
'
'        SQL = "INSERT INTO " & gcServerAudit & "AUDPERS(cCodPers, cCalGen, cObsGen, dFecMod, cCodUsu ) " _
'            & "VALUES('" & Trim(lblCodPers) & "','" & Trim(lsCalGen) & "','" _
'            & Trim(Replace(txtObs, "'", "''")) & "','" & FechaHora(gdFecSis) & "','" & gsCodUser & "')"
'
'        dbCmact.Execute SQL
'    Else
'        For I = 1 To fgCredito.Rows - 1
'            lsAut = IIf(fgCredito.TextMatrix(I, 8) = ".", "1", "0")
'            SQL = "   UPDATE " & gcServerAudit & "AUDPERSDET SET NSALDOCAP =" & fgCredito.TextMatrix(I, 3) & ", CCALAUD='" & fgCredito.TextMatrix(I, 4) & "',CNOTAAUD='" & fgCredito.TextMatrix(I, 5) & "',COBSAUD='" & Trim(fgCredito.TextMatrix(I, 6)) & "', " _
'                    & "                     nDiasAtraso=" & fgCredito.TextMatrix(1, 7) & "," _
'                    & "                     cAutorizado ='" & lsAut & "',cEstado='" & Trim(fgCredito.TextMatrix(I, 10)) & "' " _
'                    & " WHERE CCODPERS ='" & Trim(lblCodPers) & "' AND DFECCAL='" & Format(fgCredito.TextMatrix(I, 2), "mm/dd/yyyy") & "' AND CCODCTA='" & Trim(fgCredito.TextMatrix(I, 1)) & "'"
'
'            dbCmact.Execute SQL
'        Next
'
'        lsCalGen = CalificaGen(Trim(lblCodPers), dbCmact)
'
'        SQL = " UPDATE " & gcServerAudit & "AUDPERS SET  cCalGen='" & Trim(lsCalGen) & "'," _
'            & "                     cObsGen='" & Trim(Replace(txtObs, "'", "''")) & "'," _
'            & "                     dFecMod='" & FechaHora(gdFecSis) & "',cCodUsu='" & gsCodUser & "' " _
'            & " WHERE cCodPers='" & Trim(lblCodPers) & "'"
'
'        dbCmact.Execute SQL
'    End If
'    Habilitacion False
'    DatosGenerales
'    If fbNuevo = False Then
'         Me.adoAudGen.Recordset.Move lnPos - 1
'         lnPos = 0
'    End If
'    dtgAudGen.SetFocus
'    fbNuevo = False
'End If
'
'Exit Sub
'ErrorGrabar:
'    MsgBox "Error N°[" & Err.Number & " ] " & Err.Description, vbInformation, "aviso"

End Sub
Private Sub cmdImprimeCalCred_Click()
'Dim lnLineasJust As Integer
'Dim lsCodCred As String
'If lblCodPers = "" Then
'    MsgBox "Datos de clientes no válidos", vbInformation, "aviso"
'    Me.TabGen.Tab = 0
'    dtgAudGen.SetFocus
'    Exit Sub
'End If
'If Me.fgCredito.TextMatrix(1, 0) = "" Then
'    MsgBox "No existen información para realizar la operacion", vbInformation, "Aviso"
'    Exit Sub
'End If
'rtf.Text = ""
'lnNumLineas = 0: lnNumPag = 1
'lsCodCred = ""
'EncabezadoGeneral False
'For I = 1 To Me.fgCredito.Rows - 1
'
'    rtf.Text = rtf.Text + Space(10) + ImpreFormat(IIf(lsCodCred <> Trim(fgCredito.TextMatrix(I, 1)), fgCredito.TextMatrix(I, 1), ""), 12) & ImpreFormat(Format(fgCredito.TextMatrix(I, 2), "dd/mm/yyyy"), 12) & _
'                ImpreFormat(CCur(fgCredito.TextMatrix(I, 3)), 10, 2, True) & ImpreFormat(Val(fgCredito.TextMatrix(I, 4)), 7, 0, False) & _
'                ImpreFormat(Val(fgCredito.TextMatrix(I, 5)), 6, 0, False) + Space(3) + JustificaTexto(Trim(fgCredito.TextMatrix(I, 6)), 40, lnLineasJust, 67, False) + gPrnSaltoLinea
'
'    lnNumLineas = lnNumLineas + lnLineasJust
'    If lnNumLineas > 60 Then
'        lnNumLineas = 0
'        lnNumPag = lnNumPag + 1
'        EncabezadoGeneral False
'    End If
'    lsCodCred = Trim(fgCredito.TextMatrix(I, 1))
'Next
'frmPrevio.Previo rtf, "Reporte De Auditoria Interna", False, 66
End Sub

Private Sub cmdImprimeCreditos_Click()
'On Error GoTo ErroAudit1
'If Me.chkImpCred.Value = 0 And Me.chkImpFte.Value = 0 And Me.chkImpGar.Value = 0 Then
'    MsgBox "Seleccione alguna opcion de Impresion por favor ", vbInformation, "Aviso"
'    Me.chkImpCred.SetFocus
'    Exit Sub
'End If
'If Len(lblCodPers) = 0 Then
'    MsgBox "Nombre de Cliente no Válido. Por favor Verifique", vbInformation, "Aviso"
'    Me.cmdBuscar.SetFocus
'    Exit Sub
'End If
'If chkImpCred.Value = 1 Then
'    If Me.fgCreditos.TextMatrix(1, 0) = "" Then
'        MsgBox "El Sistema no detecta Créditos de Clientes", vbInformation, "Aviso"
'        Me.cmdBuscar.SetFocus
'        Exit Sub
'    End If
'End If
'If Me.chkImpGar.Value = 1 Then
'    If Me.fgGarantias.TextMatrix(1, 0) = "" Then
'        MsgBox "El Sistema no detecta Garantias de Clientes", vbInformation, "Aviso"
'        Me.cmdBuscar.SetFocus
'        Exit Sub
'    End If
'End If
'If Me.chkImpFte.Value = 1 Then
'    If Me.fgFuentes.TextMatrix(1, 0) = "" Then
'        MsgBox "El Sistema no detecta fuentes de Ingresos de Clientes", vbInformation, "Aviso"
'        Me.cmdBuscar.SetFocus
'        Exit Sub
'    End If
'End If
'
'rtf.Text = ""
'lnNumLineas = 0
'If chkImpCred.Value = 1 Then
'    ImpresionCreditos
'End If
'If chkImpGar.Value = 1 Then
'    ImprimeGarantias1
'End If
'If chkImpFte.Value = 1 Then
'    ImprimeFuentes
'End If
'frmPrevio.Previo rtf, "REPORTE DE POSICION DE CLIENTE - AUDITORIA", False, 66
'Exit Sub
'ErroAudit1:
'       MsgBox "Error: [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
'       BarraEstado.Panels(2).Text = ""
'       BarraEstado.Panels(1).Text = ""
End Sub
Private Sub cmdImprimir_Click()
'Dim lnLineasJust As Integer
'Dim lsCodCred As String
'If lblCodPers = "" Then
'    MsgBox "Datos de clientes no válidos", vbInformation, "aviso"
'    dtgAudGen.SetFocus
'    Exit Sub
'End If
'rtf.Text = ""
'lnNumLineas = 0: lnNumPag = 1
'lsCodCred = ""
'EncabezadoGeneral True
'If Me.fgCredito.TextMatrix(1, 0) <> "" Then
'    For I = 1 To Me.fgCredito.Rows - 1
'        rtf.Text = rtf.Text + Space(10) + ImpreFormat(IIf(lsCodCred <> Trim(fgCredito.TextMatrix(I, 1)), fgCredito.TextMatrix(I, 1), ""), 12) & ImpreFormat(Format(fgCredito.TextMatrix(I, 2), "dd/mm/yyyy"), 12) & _
'                    ImpreFormat(CCur(fgCredito.TextMatrix(I, 3)), 10, 2, True) & ImpreFormat(Val(fgCredito.TextMatrix(I, 4)), 7, 0, False) & _
'                    ImpreFormat(Val(fgCredito.TextMatrix(I, 5)), 6, 0, False) + Space(3) + JustificaTexto(Trim(fgCredito.TextMatrix(I, 6)), 40, lnLineasJust, 67, False) + gPrnSaltoLinea
'
'        lnNumLineas = lnNumLineas + lnLineasJust
'        If lnNumLineas > 60 Then
'            lnNumLineas = 0
'            lnNumPag = lnNumPag + 1
'            EncabezadoGeneral False
'        End If
'        lsCodCred = Trim(fgCredito.TextMatrix(I, 1))
'    Next
'End If
'frmPrevio.Previo rtf, "Reporte De Auditoria Interna", False, 66
End Sub

Private Sub CmdModificar_Click()

'If Me.adoAudGen.Recordset.EOF Then
'    MsgBox "No existen datos para modificar", vbInformation, "Aviso"
'    cmdNuevo.SetFocus
'    Exit Sub
'End If
'If Me.lblCodPers = "" Then
'    MsgBox "Datos de Cliente no Válido", vbInformation, "Aviso"
'    cmdNuevo.SetFocus
'    Exit Sub
'End If
'
'frmColocEvalCalCliDatos.lbNuevaPers = False
'
'Habilitacion True
'If Not adoAudGen.Recordset.EOF Then
'    lnPos = Me.adoAudGen.Recordset.Bookmark
'End If
'fbNuevo = False
'Me.TabGen.Tab = 0
'Me.txtObs.SetFocus

End Sub

Private Sub cmdNuevo_Click()

Dim loPers As UPersona
Dim lsPersCod As String, lsPersNombre As String, lsPersDocId As String, lsPersDocTrib As String
Dim loExiste As nColocEvalCal
Dim lbExiste As Boolean
Dim lrCreditos  As ADODB.Recordset

'On Error GoTo ControlError

Set loPers = New UPersona
    Set loPers = frmBuscaPersona.Inicio
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
    lsPersDocId = loPers.sPersIdnroDNI
    lsPersDocTrib = loPers.sPersIdnroRUC
Set loPers = Nothing

Habilitacion True
DoEvents

If Len(Trim(lsPersCod)) > 0 Then
    LimpiarControles
    fbNuevo = True
    ' Muestra datos cliente
    lblCodPers.Caption = Trim(lsPersCod)
    'lblCodPers2.Caption = Trim(lsPersCod)
    lblNomPers.Caption = PstaNombre(lsPersNombre, False)
    lblNomPers1.Caption = PstaNombre(lsPersNombre, False)
    lblDocNat.Caption = lsPersDocId
    lblDocJur.Caption = lsPersDocTrib
    
    Set loExiste = New nColocEvalCal
        'lbExiste = loExiste.nVerifExisteEvaluacion(lsPersCod)
    Set loExiste = Nothing
    If lbExiste = True Then
        MsgBox "Cliente ya se encuentra registrado dentro de la Información de Auditoria", vbInformation, "aviso"
        Habilitacion False
        fbNuevo = False
        dtgAudGen.SetFocus
        Exit Sub
    End If
    CabeceraGrid
    'Obtiene datos de Cliente
    ObtieneDatosCliente (lsPersCod)
Else
    Habilitacion False
    fbNuevo = False
    Me.dtgAudGen.SetFocus
End If
End Sub

Private Sub cmdPasar_Click()
    TabGen.Tab = 1
End Sub

Private Sub cmdPlantilla_Click()
On Error GoTo ErrorPlantilla

'Me.cdlOpen.CancelError = True
'Me.cdlOpen.DialogTitle = "Plantillas de Observaciones de Auditoria"
'Me.cdlOpen.Filter = "Archivos de Texto|*.txt|"
'Me.cdlOpen.InitDir = App.Path
'Me.cdlOpen.ShowOpen
'rtf.LoadFile Me.cdlOpen.FileName, 1
'Me.txtObs = rtf.Text
Exit Sub
ErrorPlantilla:
    MsgBox "Error N°[" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dtgAudGen_HeadClick(ByVal ColIndex As Integer)
'If Not adoAudGen.Recordset Is Nothing Then
'    If Not adoAudGen.Recordset.EOF Then
'        adoAudGen.Recordset.Sort = dtgAudGen.Columns(ColIndex).DataField
'    End If
'End If
End Sub

Private Sub dtgAudGen_KeyPress(KeyAscii As Integer)
'Dim lsCriterio As String
'KeyAscii = intfMayusculas(KeyAscii)
'lsCriterio = " Persona LIKE '" & Chr(KeyAscii) & "*'"
'If Not adoAudGen.Recordset.EOF And Not adoAudGen.Recordset.BOF Then
'    If Asc(Mid(adoAudGen.Recordset!Persona, 1, 1)) <= KeyAscii Then
'        BuscaDato lsCriterio, adoAudGen.Recordset, 1, False, 1
'    Else
'        BuscaDato lsCriterio, adoAudGen.Recordset, 1, False, 0
'    End If
'End If

End Sub

Private Sub fgCredito_Click()
'Dim SQL As String
'If fgCredito.TextMatrix(1, 0) <> "" Then
'    If fgCredito.Col = 8 Then
'        If lbAutorizado Then
'            If fgCredito.TextMatrix(fgCredito.Row, 8) = "." Then
'                fgCredito.TextMatrix(fgCredito.Row, 8) = ""
'                Set fgCredito.CellPicture = imgList.ListImages(1).Picture
'                SQL = "UPDATE " & gcServerAudit & "AudPersDet SET cAutorizado='0', dFecMod='" & FechaHora(gdFecSis) & "', cUsuAut='" & gsCodUser & "'   WHERE cCodCta='" & fgCredito.TextMatrix(fgCredito.Row, 1) & "' and cCodPers='" & Trim(lblCodPers) & "' and dFecCal ='" & Format(fgCredito.TextMatrix(fgCredito.Row, 2), "mm/dd/yyyy") & "' "
'                dbCmact.Execute SQL
'            Else
'                fgCredito.TextMatrix(fgCredito.Row, 8) = "."
'                Set fgCredito.CellPicture = imgList.ListImages(2).Picture
'                SQL = "UPDATE " & gcServerAudit & "AudPersDet SET cAutorizado='1', dFecMod='" & FechaHora(gdFecSis) & "', cUsuAut='" & gsCodUser & "'  WHERE cCodCta='" & fgCredito.TextMatrix(fgCredito.Row, 1) & "' and cCodPers='" & Trim(lblCodPers) & "' and dFecCal ='" & Format(fgCredito.TextMatrix(fgCredito.Row, 2), "mm/dd/yyyy") & "' "
'                dbCmact.Execute SQL
'
'            End If
'        Else
'            MsgBox "Permiso denegado para Aprobar calificaciones", vbInformation, "Aviso"
'        End If
'    End If
'End If
End Sub
Private Sub fgCredito_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim Col As Integer
'Dim SQL As String
'If KeyCode = 32 Then
'    If lbAutorizado Then
'        Col = Me.fgCredito.Col
'        If fgCredito.TextMatrix(fgCredito.Row, 8) = "." Then
'            fgCredito.TextMatrix(fgCredito.Row, 8) = ""
'            fgCredito.Col = 8
'            Set fgCredito.CellPicture = imgList.ListImages(1).Picture
'            SQL = "UPDATE " & gcServerAudit & "AudPersDet SET cAutorizado='0', dFecMod='" & FechaHora(gdFecSis) & "', cUsuAut='" & gsCodUser & "'   WHERE cCodCta='" & fgCredito.TextMatrix(fgCredito.Row, 1) & "' and cCodPers='" & Trim(lblCodPers) & "' and dFecCal ='" & Format(fgCredito.TextMatrix(fgCredito.Row, 2), "mm/dd/yyyy") & "' "
'            dbCmact.Execute SQL
'        Else
'            fgCredito.TextMatrix(fgCredito.Row, 8) = "."
'            fgCredito.Col = 8
'            Set fgCredito.CellPicture = imgList.ListImages(2).Picture
'            SQL = "UPDATE " & gcServerAudit & "AudPersDet SET cAutorizado='1', dFecMod='" & FechaHora(gdFecSis) & "', cUsuAut='" & gsCodUser & "'  WHERE cCodCta='" & fgCredito.TextMatrix(fgCredito.Row, 1) & "' and cCodPers='" & Trim(lblCodPers) & "' and dFecCal ='" & Format(fgCredito.TextMatrix(fgCredito.Row, 2), "mm/dd/yyyy") & "' "
'            dbCmact.Execute SQL
'
'        End If
'        Me.fgCredito.Col = Col
'    Else
'        MsgBox "Permiso denegado para Aprobar calificaciones", vbInformation, "Aviso"
'    End If
'End If
End Sub

Private Sub fgCredito_RowColChange()
txtObsDet = fgCredito.TextMatrix(fgCredito.Row, 6)
End Sub


Private Sub fgCreditos_DblClick()
'If Me.fgCreditos.TextMatrix(1, 0) <> "" Then
'    Me.BarraEstado.Panels(1).Text = "Cargando Datos de Credito. por Favor Espere"
'    fgCreditos.Enabled = False
'    If Mid(Trim(fgCreditos.TextMatrix(fgCreditos.Row, 2)), 3, 3) = "305" Then ' Prendario
'        frmPstaPrend.frmIni Trim(fgCreditos.TextMatrix(fgCreditos.Row, 2))
'    Else
'        frmPstaCred.frmIni Trim(fgCreditos.TextMatrix(fgCreditos.Row, 2)), Trim(fgCreditos.TextMatrix(fgCreditos.Row, 5))
'    End If
'    fgCreditos.Enabled = True
'    BarraEstado.Panels(1).Text = "Evaluación y Clasificacion de Colocaciones"
'    'hacemos la conexion a la base puesto que la cierra el formulario que se ha llamado anteriormente
'    AbreConexion
'    adoAudGen.ConnectionString = gsConnection
'End If
End Sub

Private Sub Form_Load()
Dim J As Integer
Dim loConstSis As NConstSistemas
Set loConstSis = New NConstSistemas
    loConstSis.LeeConstSistema (42) ' Agencia Auditoria
    fdFechaFinMes = loConstSis.LeeConstSistema(gConstSistCierreMesNegocio)
Set loConstSis = Nothing

'If CargaVarAuditoria = False Then
'    lbAud = False
'End If

CabeceraGrid
TabGen.Tab = 0
TabPosicion.Tab = 0
'fdFechaFinMes = CDate(ReadVarSis("ADM", "dFecCierreMes"))
mskFechaMes.Text = fdFechaFinMes

DatosGenerales
fbNuevo = False
CabGridCreditos
CabGridGarantias
CabGridFuentes
Habilitacion False

'fnTipoCambio = EmiteCambioFijo(Format(gdFecSis, "dd/mm/yyyy"))
'fnMontoRango = Format(CDbl(ReadVarSis("CRE", "cRangSepProv")), "#,#0.00")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If Me.cmdNuevo.Visible = False Then
'    If MsgBox("Aun no se ha grabado la informacion" & gPrnSaltoLinea & "Desea Salir sin grabar la información ??", vbQuestion + vbYesNo) = vbYes Then
'        Cancel = 0
'    Else
'        Cancel = 1
'    End If
'End If
End Sub

Private Sub ObtieneDatosCliente(ByVal psCodPers As String)
'Dim Agencia As String
'Dim SQL As String
'Dim lsFecha As String
'Dim J As Integer
'Dim lnItem As Integer
'Dim lnItemPinta As Integer
'
''On Error GoTo ErrorConexion
'CabGridCreditos
'CabGridGarantias
'CabGridFuentes
'lsActividades.Clear
'fnTotalCred = 0
'fnTotalGar = 0
'fnTotalFte = 0
'
'    BarraEstado.Panels(1).Text = " Recuperando informacion del Cliente. Por Favor Espere..."
'    BarraEstado.Panels(2).Text = "Por Favor Espere ..."
'
'    Call MuestraDatosClienteCreditos(psCodPers)
'    'Actividades pscodpers
'    If Me.chkFuentes.Value = 1 Then
'    '    DatosFuentes psCodPers
'    End If
'    BarraEstado.Panels(2).Text = ""
'    DoEvents
'
'
''Colorea las Evaluaciones de creditos Vigentes
'For lnItem = 1 To fgCreditos.Rows - 1
'    If fgCreditos.TextMatrix(lnItem, 21) = "F" Then
'        For lnItemPinta = 1 To fgCredito.Rows - 1
'            If Trim(fgCreditos.TextMatrix(lnItem, 2)) = Trim(fgCredito.TextMatrix(lnItemPinta, 1)) Then
'                fgCredito.Row = lnItemPinta
'                fgCredito.Col = 0
'                fgCredito.CellBackColor = &HE1E1C0
'                'BackColorFg fgCredito, "&H00E0E0A0", True
'                fgCredito.CellFontBold = True
'            End If
'        Next
'    End If
'    If Trim(fgCreditos.TextMatrix(lnItem, 21) = "1") Or Trim(fgCreditos.TextMatrix(lnItem, 21) = "4") Or Trim(fgCreditos.TextMatrix(lnItem, 21) = "6") Or Trim(fgCreditos.TextMatrix(lnItem, 21) = "7") Then
'        For lnItemPinta = 1 To fgCredito.Rows - 1
'            If Trim(fgCreditos.TextMatrix(lnItem, 2)) = Trim(fgCredito.TextMatrix(lnItemPinta, 1)) Then
'                fgCredito.Row = lnItemPinta
'                fgCredito.Col = 0
'                fgCredito.CellBackColor = &HCABC2
'                'BackColorFg fgCredito, "&HCABD0", True
'                fgCredito.CellFontBold = True
'            End If
'        Next
'    End If
'    If fgCreditos.TextMatrix(lnItem, 21) = "V" Or fgCreditos.TextMatrix(lnItem, 21) = "W" Then
'        For lnItemPinta = 1 To fgCredito.Rows - 1
'            If Trim(fgCreditos.TextMatrix(lnItem, 2)) = Trim(fgCredito.TextMatrix(lnItemPinta, 1)) Then
'                fgCredito.Row = lnItemPinta
'                fgCredito.Col = 0
'                fgCredito.CellBackColor = &HCAB10
'                'BackColorFg fgCredito, "&HCABA0", True
'                fgCredito.CellFontBold = True
'            End If
'        Next
'    End If
'
'Next
'
'
'BarraEstado.Panels(1).Text = "Culminado el Proceso de Busqueda de Información."
'BarraEstado.Panels(2).Text = ""
'Me.barra.Value = 0
''Me.TabPosicion.Tab = 0
'Screen.MousePointer = 0
'If Me.fgCreditos.TextMatrix(1, 0) = "" And Me.fgFuentes.TextMatrix(1, 0) = "" And Me.fgGarantias.TextMatrix(1, 0) = "" Then
'    MsgBox "Cliente no posee créditos en agencias seleccionadas", vbInformation, "Aviso"
'End If
'Exit Sub
'ErrorConexion:
'    Screen.MousePointer = 0
'    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description & Chr(13) & "Por Favor Consulte al Dpto. de Sistemas", vbInformation, "Aviso"
'    BarraEstado.Panels(1).Text = "Error en busqueda de información de cliente"
'    BarraEstado.Panels(2).Text = "Por Favor consulte al Area de Sistemas"
End Sub
Private Sub MuestraDatosClienteCreditos(ByVal psCodPers As String)
Dim loDatCred As nColocEvalCal
Dim lrData As ADODB.Recordset

Dim lnCorreCred As Integer
Dim lsMonto As String
Dim lsFechaCred  As String
Dim lsAgencia As String

barraestado.Panels(1).Text = "Cargando información de Creditos. Por Favor Espere..."

Set loDatCred = New nColocEvalCal
    Set lrData = loDatCred.nObtieneDatosClienteCreditos(psCodPers, mskFechaMes.Text)
Set loDatCred = Nothing

'Total = lrData.RecordCount

If Not (lrData.BOF And lrData.EOF) Then
    Do While Not lrData.EOF
        'I = I + 1
        
        'lsAgenciaCod = Mid(lrdata!cCodCta, 1, 2)
        'lsAgencia = Tablacod("47", "112" & lsAgenciaCod)
        lnCorreCred = lnCorreCred + 1
        
        AdicionaRow fgCreditos
        lnCorreCred = fgCreditos.Row

        Select Case Trim(lrData!nPrdEstado)
            Case gColocEstSolic
            '    lsMonto = IIf(IsNull(lrData!nMontoSol), "0.00", Format(Trim(lrData!nMontoSol), "####0.00"))
            '    lsFechaCred = IIf(IsNull(lrData!dAsignacion), "", Format(Trim(lrData!dAsignacion), "dd/mm/yyyy"))
            'Case "B"
            '    lsMonto = IIf(IsNull(rs!nMontoSug), "0.00", Format(Trim(rs!nMontoSug), "####0.00"))
            '    lsFechaCred = IIf(IsNull(rs!dAsignacion), "", Format(Trim(rs!dAsignacion), "dd/mm/yyyy"))
            'Case "L"
            '    lsRechazo = Tablacod("27", IIf(IsNull(rs!cCauRech), "", Trim(rs!cCauRech)))
            '    If IsNull(rs!nMontoApr) Then
            '        lsFechaCred = IIf(IsNull(rs!dAsignacion), "NI        ", Format(Trim(rs!dAsignacion), "dd/mm/yyyy"))
            '        If IsNull(rs!nMontoSug) Then
            '            lsMonto = IIf(IsNull(rs!nMontoSol), "0.00", Format(Trim(rs!nMontoSol), "####0.00"))
            '        Else
            '            lsMonto = IIf(IsNull(rs!nMontoSug), "0.00", Format(Trim(rs!nMontoSug), "####0.00"))
            '        End If
            '    Else
            '        lsFechaCred = IIf(IsNull(rs!dFecApr), "NI        ", Format(Trim(rs!dFecApr), "dd/mm/yyyy"))
            '        lsMonto = IIf(IsNull(rs!nMontoApr), "0.00", Format(Trim(rs!nMontoApr), "####0.00"))
            '    End If
            'Case "C"
            '    lsFechaCred = IIf(IsNull(rs!dFecApr), "", Format(Trim(rs!dFecApr), "dd/mm/yyyy"))
            '    lsMonto = IIf(IsNull(rs!nMontoApr), "0.00", Format(Trim(rs!nMontoApr), "####0.00"))
            'Case "F", "R"
            '    lsFechaCred = IIf(IsNull(rs!dFecVig), "", Format(Trim(rs!dFecVig), "dd/mm/yyyy"))
            '    lsMonto = IIf(IsNull(rs!nMontoApr), "0.00", Format(Trim(rs!nMontoApr), "####0.00"))
            'Case "G"
            '    lsFechaCred = IIf(IsNull(rs!dFecVig), "", Format(Trim(rs!dFecVig), "dd/mm/yyyy"))
            '    lsMonto = IIf(IsNull(rs!nMontoApr), "0.00", Format(Trim(rs!nMontoApr), "####0.00"))
            'Case "H"
            '    lsFechaCred = IIf(IsNull(rs!dFecVig), "", Format(Trim(rs!dFecVig), "dd/mm/yyyy"))
            '    lsMonto = IIf(IsNull(rs!nMontoApr), "0.00", Format(Trim(rs!nMontoApr), "####0.00"))
            '***** Predarios  *****
            'Case "1", "4", "6", "7"
            '    lsFechaCred = IIf(IsNull(rs!dFecVig), "", Format(Trim(rs!dFecVig), "dd/mm/yyyy"))
            '    lsMonto = IIf(IsNull(rs!nMontoApr), "0.00", Format(Trim(rs!nMontoApr), "####0.00"))
            '    '***** Predarios  *****
            'Case "V", "W"
            '    lsFechaCred = IIf(IsNull(rs!dFecVig), "", Format(Trim(rs!dFecVig), "dd/mm/yyyy"))
            '    lsMonto = IIf(IsNull(rs!nMontoApr), "0.00", Format(Trim(rs!nMontoApr), "####0.00"))
            'Case Else
            '    lsFechaCred = IIf(IsNull(rs!dAsignacion), "NI       ", Format(Trim(rs!dAsignacion), "dd/mm/yyyy"))
        End Select
        
        fgCreditos.TextMatrix(lnCorreCred, 1) = lsFechaCred
        fgCreditos.TextMatrix(lnCorreCred, 2) = Trim(lrData!cCtaCod)
        fgCreditos.TextMatrix(lnCorreCred, 3) = Trim(lrData!cNomAgencia)
        'fgCreditos.TextMatrix(lnCorreCred, 4) = fgProductoCreditoTipo(Mid(lrData!cCtaCod, 6, 3))

        'If rs!cEstado = "V" Then
        '    If rs!cRefinan = "J" Then
        '        fgCreditos.TextMatrix(lnCorreCred, 5) = "JUDICIAL"
        '    Else
        '        fgCreditos.TextMatrix(lnCorreCred, 5) = "CASTIGADO"
        '    End If
        'Else
        '    fgCreditos.TextMatrix(lnCorreCred, 5) = Tablacod(IIf(Mid(rs!cCodCta, 3, 3) = "305", "46", "26"), Trim(rs!cEstado))
        'End If
        'fgCreditos.TextMatrix(lnCorreCred, 6) = Tablacod("25", Trim(rs!cRelaCta))
        'fgCreditos.TextMatrix(lnCorreCred, 7) = Trim(lrData!cCodAnalista)
        fgCreditos.TextMatrix(lnCorreCred, 8) = lsMonto
        fgCreditos.TextMatrix(lnCorreCred, 9) = Trim(IIf(IsNull(lrData!cNota1), 0, lrData!cNota1))
        fgCreditos.TextMatrix(lnCorreCred, 10) = Trim(IIf(IsNull(lrData!Nota1FM), 0, lrData!Nota1FM))
        fgCreditos.TextMatrix(lnCorreCred, 11) = Format(IIf(IsNull(lrData!nSaldo), 0, lrData!nSaldo), "#0.00")
        fgCreditos.TextMatrix(lnCorreCred, 12) = Format(IIf(IsNull(lrData!SaldoFM), 0, lrData!SaldoFM), "#0.00") '
        'fgCreditos.TextMatrix(lnCorreCred, 13) = lrData!nDiasAtraso '
        fgCreditos.TextMatrix(lnCorreCred, 14) = IIf(IsNull(lrData!AtrasoFM), "", lrData!AtrasoFM) '
        'fgCreditos.TextMatrix(lnCorreCred, 15) = CodigoAntiguo("CREDITOS", lrdata!cCodCta, dbConexion) '
        fgCreditos.TextMatrix(lnCorreCred, 16) = IIf(Mid(lrData!cCtaCod, 6, 1) = "1", "MN", "ME")
        'fgCreditos.TextMatrix(lnCorreCred, 17) = Trim(rs!cRefinan) '
        'fgCreditos.TextMatrix(lnCorreCred, 18) = lsRechazo  '
        'fgCreditos.TextMatrix(lnCorreCred, 19) = AbrevProd(Mid(lrData!cCtaCod, 3, 3)) '
        'fgCreditos.TextMatrix(lnCorreCred, 20) = IIf(IsNull(rs!dCancelado), "NI", Format(rs!dCancelado, "dd/mm/yyyy"))
        fgCreditos.TextMatrix(lnCorreCred, 21) = Trim(lrData!nPrdEstado)
        fgCreditos.TextMatrix(lnCorreCred, 22) = IIf(IsNull(lrData!dFecha), "", Format(lrData!dFecha, "dd/mm/yyyy"))
        
        'If Trim(lrData!nPrdEstado) = gColocEstVigNorm Then
        '    DatosGarantias rs!cCodCta, dbConexion, Trim(Right(listaAgencias.List(I), 5))
        '    fgCreditos.Row = fgCreditos.Row
        '    fgCreditos.Col = 0
         '   fgCreditos.CellBackColor = &HE1E1C0
         '   fgCreditos.CellFontBold = True
         '   If rs!cRelaCta = "TI" Then
         '       BackColorFg fgCreditos, "&H00E0E0A0", True
         '   End If
        'End If
        'If Trim(rs!cEstado) = "1" Or Trim(rs!cEstado) = "4" Or Trim(rs!cEstado) = "6" Or Trim(rs!cEstado) = "7" Then
        '    fgCreditos.Row = fgCreditos.Row
        '    fgCreditos.Col = 0
        '    fgCreditos.CellBackColor = &HCABC2
        '    fgCreditos.CellFontBold = True
        '   If rs!cRelaCta = "TI" Then
       '         BackColorFg fgCreditos, "&HCABD0", True
       '     End If
       ' End If
       ' If Trim(rs!cEstado) = "V" Or Trim(rs!cEstado) = "W" Then
       '     fgCreditos.Row = fgCreditos.Row
       '     fgCreditos.Col = 0
       '     fgCreditos.CellBackColor = &HCAB10
       '    fgCreditos.CellFontBold = True
       ''     If rs!cRelaCta = "TI" Then
       '         BackColorFg fgCreditos, "&HCABA0", True
       '     End If
       ' End If
        
        'Me.barra.Value = (I / Total) * 100
        DoEvents
        lrData.MoveNext
    Loop
End If
lrData.Close
Set lrData = Nothing
End Sub
Private Sub Actividades(lsCodPers As String, dbConexion As ADODB.Connection)
'Dim SQL As String
'Dim rs As New ADODB.Recordset
'SQL = "   SELECT  cCodpers, cActecon, cSector, TA.CNOMTAB as Actividad , TS.CNOMTAB as Sector" _
'    & " FROM    Fuenteingreso F " _
'    & "         JOIN " & gcCentralCom & "TablaCod TA    ON TA.CVALOR = F.CACTECON " _
'    & "         JOIN " & gcCentralCom & "TablaCod TS    ON TS.CVALOR = F.CSECTOR " _
'    & " WHERE   cCodPers ='" & lsCodPers & "' " _
'    & "         AND TA.CCODTAB LIKE '35__' AND  TS.CCODTAB LIKE '20__' " _
'    & " GROUP BY cCodpers, cActecon, cSector, TA.CNOMTAB , TS.CNOMTAB "
'
'rs.CursorLocation = adUseClient
'rs.Open SQL, dbConexion, adOpenStatic, adLockReadOnly, adCmdText
'Set rs.ActiveConnection = Nothing
'
'Do While Not rs.EOF
'    lsActividades.AddItem Trim(rs!Actividad) & Space(5) & Trim(rs!Sector)
'    rs.MoveNext
'Loop
'rs.Close
'Set rs = Nothing
End Sub
Private Sub DatosGarantias(lsCodCta As String, dbConexion As ADODB.Connection, psCodAge As String)
'Dim SQL As String
'Dim rs As New ADODB.Recordset
'Dim lsCredGarant As String
'Dim Total As Integer, I As Integer
'Dim N As Integer
'
''TabPosicion.Tab = 1
'barraestado.Panels(1).Text = "Cargando información de Garantias en Agencia : " & psCodAge & ".  Por Favor Espere..."
'
'
'SQL = "SELECT   GC.cCodCta, G.cNumGarant, PE1.cCodPers , PG.cNomPers, PG.cNudoci, PE1.cRelaCta, TG.cNomTab as TipoGarantia, G.cTipoGarant, " _
'    & "         Convert(VarChar(30),ISNULL(G.cDesGarant,'')) as  Descripcion, G.cDocGarant,G.cMoneda as Moneda, ISNULL(G.nMontoTasac,0) AS TASACION, " _
'    & "         ISNULL(G.nMontoRealiz,0) AS REALIZACION , ISNULL(G.nMontoxGrav,0) AS PORGRAVAR, G.cEstado, ISNULL(GC.nMontoGrava,0) as Gravado, " _
'    & "         GC.cMoneda as MonedaGC " _
'    & " FROM    GARANTCRED GC   " _
'    & "         JOIN GARANTIAS G        ON G.CNUMGARANT =GC.CNUMGARANT " _
'    & "         JOIN CREDITO C          ON C.CCODCTA =GC.CCODCTA " _
'    & "         JOIN " & gcCentralPers & "Persona PG         ON PG.CCODPERS =G.CCODPERS " _
'    & "         JOIN PERSCREDITO PE1    ON (PE1.CCODPERS=PG.CCODPERS AND PE1.CCODCTA =GC.CCODCTA)  " _
'    & "         JOIN " & gcCentralCom & "TablaCod TG        ON TG.CVALOR=G.cTipoGarant " _
'    & " WHERE   TG.CCODTAB LIKE '24__'  " _
'    & "         AND GC.cCodCta='" & lsCodCta & "' ORDER BY G.cNumGarant "
'
'
'rs.Open SQL, dbConexion, adOpenStatic, adLockReadOnly, adCmdText
'Total = rs.RecordCount
'lsCredGarant = ""
'Do While Not rs.EOF
'    I = I + 1
'    AdicionaRow fgGarantias
'    N = fgGarantias.Row
'    If lsCredGarant <> Trim(rs!cCodCta) Then
'        BackColorFg fgGarantias, "&H00E0E0E0"
'        fgGarantias.TextMatrix(N, 1) = Trim(rs!cCodCta)
'    End If
'    fgGarantias.TextMatrix(N, 2) = Trim(PstaNombre(rs!cNomPers, False))
'    fgGarantias.TextMatrix(N, 3) = Trim(rs!cRelaCta)
'    fgGarantias.TextMatrix(N, 4) = Trim(rs!TipoGarantia)
'    fgGarantias.TextMatrix(N, 5) = Trim(rs!Descripcion)
'    fgGarantias.TextMatrix(N, 6) = Tablacod("49", rs!cDocGarant)  'Trim(rs!TipoDocGar)
'    fgGarantias.TextMatrix(N, 7) = Trim(IIf(Trim(rs!Moneda) = "1", "MN", "ME"))
'    fgGarantias.TextMatrix(N, 8) = Format(rs!Tasacion, "#,#0.00")
'    fgGarantias.TextMatrix(N, 9) = Format(rs!REALIZACION, "#,#0.00")
'    fgGarantias.TextMatrix(N, 10) = Format(rs!PorGravar, "#,#0.00")
'    fgGarantias.TextMatrix(N, 11) = Trim(rs!cEstado)
'    fgGarantias.TextMatrix(N, 12) = Format(rs!Gravado, "#,#0.00")
'    fgGarantias.TextMatrix(N, 13) = Trim(IIf(rs!MonedaGC = "1", "MN", "ME"))
'    fgGarantias.TextMatrix(N, 14) = Trim(rs!cNumGarant)
'    lsCredGarant = Trim(rs!cCodCta)
'    Me.barra.Value = (I / Total) * 100
'    rs.MoveNext
'Loop
'rs.Close
'Set rs = Nothing
'Me.barraestado.Panels(1).Text = "Término de Búsqueda de garantías"
End Sub
Private Sub DatosFuentes(lsCodPers As String, dbConexion As ADODB.Connection, psCodAge As String)
'Dim SQL As String
'Dim rs As New ADODB.Recordset
'Dim Total As Integer, I As Integer
'Dim N As Integer
'Dim lsNumfuente As String
'
'SQL = "   SELECT  'TIPOFTE'=   CASE " _
'        & "                     WHEN F.cTipoFuente='D' THEN 'DEPENDIENTE' " _
'        & "                     Else 'INDEPENDIENTE' " _
'        & "                 END, " _
'        & "         ISNULL(F.cRazonSocial,'') AS cRazonSocial, ISNULL(F.cdireccion,'') AS cdireccion, ISNULL(Z.cDesZon,'NI') AS ZONA, ISNULL(F.ccodzon,'') AS cCodZon, " _
'        & "         ISNULL(S.CNOMTAB,'NI') as Sector, ISNULL(F.cSector,'') AS cSector, ISNULL(A.CNOMTAB,'NI') AS ACTIVIDAD,  ISNULL(F.cActeCon,'') AS cActEcon, ISNULL(F.ccargo,'') as Cargo, " _
'        & "         FD.dFecEval AS FDFECHA, ISNULL(FD.nIngClt,0) AS INGRESO , ISNULL(FD.nGasFam,0) AS GASTOS, " _
'        & "         B.dFecBalanc AS BALFECHA, ISNULL(B.ningfam,0) AS BALINGFAM ,ISNULL(B.ngasfam,0) AS BALGASFAM,F.CNUMFUENTE " _
'        & " FROM    FUENTEINGRESO F  JOIN " & gcCentralPers & "Persona P ON F.CCODPERS=P.CCODPERS " _
'        & "         LEFT JOIN FDEPENDIENTE FD   ON FD.CNUMFUENTE=F.CNUMFUENTE " _
'        & "         LEFT JOIN BALANCE B     ON B.CNUMFUENTE=F.CNUMFUENTE " _
'        & "         LEFT JOIN " & gcCentralCom & "Zonas Z       ON Z.CCODZON=F.CCODZON " _
'        & "         LEFT JOIN " & gcCentralCom & "TablaCod S         ON S.CVALOR = F.CSECTOR " _
'        & "         LEFT JOIN " & gcCentralCom & "TablaCod A         ON A.CVALOR = F.CACTECON " _
'        & " WHERE   P.CCODPERS='" & lsCodPers & "' AND S.CCODTAB LIKE '20__' " _
'        & "         AND A.CCODTAB LIKE '35__' ORDER BY  F.CNUMFUENTE "
'
''TabPosicion.Tab = 2
'barraestado.Panels(1).Text = "Cargando información de Fuentes Ingreso en " & psCodAge & ".  Por Favor Espere..."
'lsNumfuente = ""
'
'rs.CursorLocation = adUseClient
'rs.Open SQL, dbConexion, adOpenStatic, adLockReadOnly, adCmdText
'Set rs.ActiveConnection = Nothing
'
'Total = rs.RecordCount
'Do While Not rs.EOF
'    I = I + 1
'    AdicionaRow fgFuentes
'    N = fgFuentes.Row
'    If lsNumfuente <> Trim(rs!cNumfuente) Then
'        BackColorFg fgFuentes, "&H00E0E0E0"
'        fgFuentes.TextMatrix(N, 1) = Trim(rs!TIPOFTE)
'        fgFuentes.TextMatrix(N, 2) = Trim(rs!cRazonSocial)
'        fgFuentes.TextMatrix(N, 3) = Trim(rs!cdireccion)
'        fgFuentes.TextMatrix(N, 4) = Trim(rs!Zona)
'        fgFuentes.TextMatrix(N, 5) = Trim(rs!Sector)
'        fgFuentes.TextMatrix(N, 6) = Trim(rs!Actividad)
'        fgFuentes.TextMatrix(N, 7) = Trim(rs!Cargo)
'    End If
'    fgFuentes.TextMatrix(N, 8) = IIf(IsNull(rs!FDFECHA), "", Format(rs!FDFECHA, "dd/mm/yyyy"))
'    fgFuentes.TextMatrix(N, 9) = Format(rs!Ingreso, "#,#0.00")
'    fgFuentes.TextMatrix(N, 10) = Format(rs!Gastos, "#,#0.00")
'    fgFuentes.TextMatrix(N, 11) = IIf(IsNull(rs!BalFecha), "", Format(rs!BalFecha, "dd/mm/yyyy"))
'    fgFuentes.TextMatrix(N, 12) = Format(rs!BalIngFam, "#,#0.00")
'    fgFuentes.TextMatrix(N, 13) = Format(rs!BalGasFam, "#,#0.00")
'    fgFuentes.TextMatrix(N, 14) = Trim(rs!cNumfuente)
'    lsNumfuente = Trim(rs!cNumfuente)
'    Me.barra.Value = (I / Total) * 100
'    rs.MoveNext
'Loop
'rs.Close
'Set rs = Nothing
End Sub
Private Sub ImpresionCreditos()
Dim lsAgencia As String
rtf.Text = ""
lnNumPag = 1
Encabezado
EncabezadoCredito
DetalleCreditos

'rtf.Text = rtf.Text + PrnSet("C-") & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea
'rtf.Text = rtf.Text & Space(50) & String(30, "_") & gPrnSaltoLinea
'rtf.Text = rtf.Text & Space(62) & "Vo Bo"
End Sub
Private Sub Encabezado()
'Dim lsFecha As String
'Dim I As Integer
'Dim lsAgencia As String
'Dim lnNumAge As Integer
'
'lsFecha = Format(gdFecSis & Time, "dd/mm/yyyy hh:mm:ss AMPM")
'rtf.Text = rtf.Text + PrnSet("B+") + CabeRepo("", "", 100, "", "REPORTE DE CALIFICACION DE CLIENTES", "AUDITORIA INTERNA", "CMACT", Format(lnNumPag, "0000")) & PrnSet("B-") + gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 7
'rtf.Text = rtf.Text & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'rtf.Text = rtf.Text & ImpreFormat("CLIENTE        : " & PstaNombre(Trim(lblNomPers), False), 80, 0) & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'rtf.Text = rtf.Text & ImpreFormat("DOC. IDENTIDAD : " & LblDocNat.Caption, 30, 0) & ImpreFormat("DOC. JURIDICO :" & LblDocJur, 30) & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'If Me.lsActividades.ListCount > 0 Then
'    rtf.Text = rtf.Text + ImpreFormat("Actividade(s) / Sector(es) :", 30, 0) & gPrnSaltoLinea
'    For I = 0 To Me.lsActividades.ListCount - 1
'        rtf.Text = rtf.Text + Space(17) & ImpreFormat(Me.lsActividades.List(I), 80, 0) & gPrnSaltoLinea
'        lnNumLineas = lnNumLineas + 1
'    Next
'Else
'    rtf.Text = rtf.Text & gPrnSaltoLinea
'    lnNumLineas = lnNumLineas + 1
'End If
'If listaAgencias.ListCount > 0 Then
'    rtf.Text = rtf.Text + PrnSet("C+") + PrnSet("I+") + ImpreFormat("", 90) + PrnSet("B+") + "Agencias Consultadas" + PrnSet("B-") + PrnSet("I-") + gPrnSaltoLinea
'    rtf.Text = rtf.Text + PrnSet("C+") + ImpreFormat("", 60) + String(50, "=") & gPrnSaltoLinea
'    lnNumAge = 0
'    For I = 0 To Me.listaAgencias.ListCount - 1
'        If listaAgencias.Selected(I) = True Then
'            lnNumAge = lnNumAge + 1
'            If lnNumAge = 1 Then
'                lsAgencia = ImpreFormat("", 60) + lsAgencia
'            End If
'            lsAgencia = lsAgencia & ImpreFormat(listaAgencias.List(I), 15) & "-"
'            If lnNumAge Mod 3 = 0 Then
'                lsAgencia = "-" & lsAgencia + gPrnSaltoLinea
'                lsAgencia = lsAgencia + ImpreFormat("", 60)
'                lnNumLineas = lnNumLineas + 1
'            End If
'        End If
'    Next I
'    If lsAgencia <> "" Then
'        rtf.Text = rtf.Text & lsAgencia & IIf(lnNumAge Mod 3 = 0, "", gPrnSaltoLinea)
'    End If
'    rtf.Text = rtf.Text + IIf(lnNumAge Mod 3 = 0, "", ImpreFormat("", 60)) & String(50, "=") + PrnSet("C-") & gPrnSaltoLinea & gPrnSaltoLinea
'End If
End Sub
Private Sub EncabezadoCredito()
'rtf.Text = rtf.Text + PrnSet("B+") + PrnSet("I+") + "LISTADO DE CREDITOS" + PrnSet("C+") + PrnSet("I-") & gPrnSaltoLinea
'If Me.fgCreditos.TextMatrix(1, 0) <> "" Then
'    rtf.Text = rtf.Text + PrnSet("C+") + PrnSet("B+") + String(118, "-") & gPrnSaltoLinea
'    rtf.Text = rtf.Text & ImpreFormat("No.", 3, 0) & ImpreFormat("FECHA", 10) & ImpreFormat("TIPO", 6) & ImpreFormat("AGENCIA", 8) _
'                & ImpreFormat("CREDITO", 12) & ImpreFormat("COD ANT", 8) & ImpreFormat("EST", 3) & ImpreFormat("REL", 3) _
'                & ImpreFormat("ANAL.", 4) & ImpreFormat("NOT1", 4) & ImpreFormat("NOT2", 7) & ImpreFormat("MONTO", 9) & ImpreFormat("SALDO", 6) & ImpreFormat("REF", 3) & ImpreFormat("DIAS", 8) & gPrnSaltoLinea
'    rtf.Text = rtf.Text & String(118, "-") + PrnSet("B-") & gPrnSaltoLinea
'    lnNumLineas = lnNumLineas + 7
'Else
'    rtf.Text = rtf.Text & Space(5) & "Cliente no Posee Ningún Crédito " & gPrnSaltoLinea
'    rtf.Text = rtf.Text & gPrnSaltoLinea
'    lnNumLineas = lnNumLineas + 2
'End If
'
End Sub
Private Sub DetalleCreditos()
'Dim I As Integer
'Dim lsTipo As String
'Dim lsnro As String
'Dim Item As ListItem
'
'With Me.fgCreditos
'    For I = 1 To .Rows - 1
'        If lnNumLineas > 60 Then
'            lnNumLineas = 0
'            rtf.Text = rtf.Text + gPrnSaltoPagina
'            lnNumPag = lnNumPag + 1
'            EncabezadoCredito
'        End If
'        lsTipo = AbrevProd(Mid(.TextMatrix(I, 2), 3, 3))
'        lsnro = .TextMatrix(I, 0)
'        rtf.Text = rtf.Text & CadDerecha(lsnro, 2) _
'                & ImpreFormat(.TextMatrix(I, 1), 10) _
'                & ImpreFormat(Trim(lsTipo), 6) _
'                & ImpreFormat(.TextMatrix(I, 3), 8) _
'                & ImpreFormat(.TextMatrix(I, 2), 12) _
'                & ImpreFormat(IIf(Len(Trim(.TextMatrix(I, 15))) = 0, "", .TextMatrix(I, 15)), 8) _
'                & ImpreFormat(Mid(.TextMatrix(I, 5), 1, 4), 3) _
'                & ImpreFormat(Mid(.TextMatrix(I, 6), 1, 3), 3) _
'                & ImpreFormat(IIf(Len(Trim(.TextMatrix(I, 7))) = 0, "", .TextMatrix(I, 7)), 6) _
'                & ImpreFormat(.TextMatrix(I, 9), 4) _
'                & ImpreFormat(.TextMatrix(I, 10), 4) _
'                & ImpreFormat(Val(.TextMatrix(I, 8)), 8, 2, True) _
'                & ImpreFormat(Val(.TextMatrix(I, 11)), 8, 2, True) _
'                & ImpreFormat(.TextMatrix(I, 17), 3) _
'                & ImpreFormat(Val(.TextMatrix(I, 13)), 3, 0, True) _
'                & gPrnSaltoLinea
'
'            lnNumLineas = lnNumLineas + 1
'    Next I
'End With
'rtf.Text = rtf.Text & String(118, "-") & gPrnSaltoLinea
'rtf.Text = rtf.Text & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 2
End Sub
Private Function VerAuditPrinc(lsCodPers As String) As Boolean
'Dim SQL As String
'Dim rs As New ADODB.Recordset
'
'SQL = " Select cCodPers,cCalGen,cObsGen,dFecMod,cCodUsu " _
'    & " FROM " & gcServerAudit & "AUDPERS WHERE cCodPers='" & lsCodPers & "'"
'
'rs.Open SQL, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
'VerAuditPrinc = Not rs.EOF
'rs.Close
'Set rs = Nothing
End Function
Private Function VerAuditDet(lsCodPers As String, lsCodCta As String) As Boolean
'Dim SQL As String
'Dim rs As New ADODB.Recordset
'
'SQL = " Select cCodPers,cCodCta " _
'    & " FROM " & gcServerAudit & "AUDPERSDET WHERE cCodPers='" & lsCodPers & "' and cCodCta='" & lsCodCta & "'"
'
'rs.Open SQL, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
'VerAuditDet = Not rs.EOF
'rs.Close
'Set rs = Nothing
End Function
Private Function Valida() As Boolean
'Dim SQL As String
'Valida = True
'If Me.lblCodPers = "" Then
'    MsgBox "Codigo de Cliente no Válido", vbInformation, "aviso"
'    If Me.cmdNuevo.Enabled And Me.cmdNuevo.Visible Then
'        Me.cmdNuevo.SetFocus
'    End If
'    Valida = False
'    Exit Function
'End If
'If Me.txtObs = "" Then
'    MsgBox "No ha ingresado observaciones a Cliente", vbInformation, "Aviso"
'    Me.TabGen.Tab = 0
'    Me.txtObs.SetFocus
'    Valida = False
'    Exit Function
'End If
'
'If Me.fgCredito.TextMatrix(1, 0) = "" Then
'    MsgBox "Información Créditicia del cliente no ingresada", vbInformation, "Aviso"
'    Valida = False
'    Me.TabGen.Tab = 1
'    Me.cmdAgregar.SetFocus
'    Exit Function
'End If
End Function



Private Sub txtCalGen_DblClick()
'Actualiza los datos
    DatosGenerales
End Sub

Private Sub txtObs_GotFocus()
fEnfoque txtObs
End Sub
Private Sub txtObs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If Me.cmdGrabar.Enabled And cmdGrabar.Visible Then
        Me.cmdGrabar.SetFocus
    End If
End If
End Sub
Private Sub Habilitacion(lbHab As Boolean)
ControlesAutorizados
    Me.cmdGrabar.Visible = lbHab
    Me.cmdPlantilla.Enabled = lbHab
    Me.cmdnuevo.Visible = Not lbHab
    Me.cmdCancelar.Visible = lbHab
    Me.cmdModificar.Visible = Not lbHab
    
    Me.txtObs.Locked = Not lbHab
    Me.dtgAudGen.Enabled = Not lbHab
    
    Me.cmdEliminaGen.Enabled = Not lbHab
    Me.cmdImprimir.Enabled = Not lbHab
End Sub
Private Sub DatosGenerales()
Dim loDatos As nColocEvalCal
Dim lrDatCalif As ADODB.Recordset

Set loDatos = New nColocEvalCal
    Set lrDatCalif = loDatos.nObtieneCreditosCalificados
Set loDatos = Nothing

Set dtgAudGen.DataSource = lrDatCalif

End Sub

Private Sub LimpiarControles()
Me.lblCodPers = ""
Me.lblCodPers1 = ""
Me.lblNomPers = ""
Me.lblNomPers1 = ""
Me.lblDocNat = ""
Me.lblDocJur = ""
Me.txtCalGen = ""
Me.txtObsDet = ""
Me.txtObs = ""
End Sub
Private Sub ImprimeGarantias()
'Dim I As Integer
'Dim lsNumGarant As String
'If chkImpGar.Value = 1 And chkImpCred.Value = 0 Then
'    lnNumLineas = 0
'    lnNumPag = 1
'    rtf.Text = ""
'    Encabezado
'End If
'rtf.Text = rtf.Text + PrnSet("I+") + PrnSet("B+") + "LISTA DE GARANTIAS" + PrnSet("I-") & gPrnSaltoLinea
'rtf.Text = rtf.Text & "------------------" + PrnSet("B-") + gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 2
'For I = 1 To Me.fgGarantias.Rows - 1            '   Me.lstGarantias.ListItems.Count
'        If lnNumLineas > 60 Then
'            lnNumLineas = 0
'            rtf.Text = rtf.Text & gPrnSaltoPagina
'            lnNumPag = lnNumPag + 1
'            'Encabezado
'        End If
'    If lsNumGarant <> Trim(fgGarantias.TextMatrix(I, 13)) Then
'        rtf.Text = rtf.Text & gPrnSaltoLinea
'        lnNumLineas = lnNumLineas + 1
'        rtf.Text = rtf.Text & ImpreFormat("Tipo Garantia     :" & Me.fgGarantias.TextMatrix(I, 3), 50, 0) & ImpreFormat("Documento : " & Me.fgGarantias.TextMatrix(I, 4), 40) & ImpreFormat(fgGarantias.TextMatrix(I, 5), 10) & gPrnSaltoLinea
'        lnNumLineas = lnNumLineas + 1
'        rtf.Text = rtf.Text & ImpreFormat("Persona Garante   :" & Me.fgGarantias.TextMatrix(I, 1), 50, 0) & ImpreFormat("Relación : " & Me.fgGarantias.TextMatrix(I, 2), 40) & gPrnSaltoLinea
'        lnNumLineas = lnNumLineas + 1
'        rtf.Text = rtf.Text & ImpreFormat("Monto Tasación    :", 20, 0) & ImpreFormat(CCur(fgGarantias.TextMatrix(I, 6)), 10, 2, True) & ImpreFormat("Realización :", 15, 5) & ImpreFormat(CCur(fgGarantias.TextMatrix(I, 7)), 10, 2, True) & ImpreFormat("Por Gravar:", 12) & ImpreFormat(CCur(fgGarantias.TextMatrix(I, 8)), 10, 2, True) + gPrnSaltoLinea
'        lnNumLineas = lnNumLineas + 1
'        rtf.Text = rtf.Text & PrnSet("B+") & ImpreFormat("", 30) & String("50", "-") & gPrnSaltoLinea
'        lnNumLineas = lnNumLineas + 1
'        rtf.Text = rtf.Text & ImpreFormat("", 30) & ImpreFormat("Credito", 17) & ImpreFormat("Gravado", 12) & ImpreFormat("Moneda", 10) & gPrnSaltoLinea
'        lnNumLineas = lnNumLineas + 1
'        rtf.Text = rtf.Text & ImpreFormat("", 30) & String("50", "-") & PrnSet("B-") & gPrnSaltoLinea
'        lnNumLineas = lnNumLineas + 1
'    End If
'    rtf.Text = rtf.Text & ImpreFormat("", 30) & ImpreFormat(Me.fgGarantias.TextMatrix(I, 10), 12) & ImpreFormat(CCur(Me.fgGarantias.TextMatrix(I, 11)), 10, 2, True) & ImpreFormat(Me.fgGarantias.TextMatrix(I, 12), 10, 8) & gPrnSaltoLinea
'    lnNumLineas = lnNumLineas + 1
'    lsNumGarant = Trim(fgGarantias.TextMatrix(I, 13))
'Next
End Sub
Private Sub CabGarant()
'rtf.Text = rtf.Text + PrnSet("I+") + PrnSet("B+") + "LISTA DE GARANTIAS" + PrnSet("I-") & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'rtf.Text = rtf.Text + PrnSet("Esp", 8) + PrnSet("B+") + PrnSet("C+") + String(115, "-") & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'rtf.Text = rtf.Text + PrnSet("I+") + ImpreFormat("CREDITO", 15) + PrnSet("I-") + gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'rtf.Text = rtf.Text & ImpreFormat("Garantía", 8) & ImpreFormat("Garante", 30) & ImpreFormat("Rel", 3) & ImpreFormat("Tipo Garantia", 15) & ImpreFormat("Descripcion", 40) & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'rtf.Text = rtf.Text & ImpreFormat("Documento", 30, 12) & ImpreFormat("Mon", 5) + PrnSet("I+") + ImpreFormat("Tasacion", 13) & ImpreFormat("Realizacion", 15) & ImpreFormat("Por Gravar", 15) & ImpreFormat("Gravado", 8) & ImpreFormat("Mon", 3) + PrnSet("I-") + gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'rtf.Text = rtf.Text & String(115, "-") + PrnSet("B-") + PrnSet("C-") + PrnSet("EspN") + gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
End Sub
Private Sub ImprimeGarantias1()
'Dim I As Integer
'Dim lsCredGarant As String
'If chkImpGar.Value = 1 And chkImpCred.Value = 0 Then
'    lnNumLineas = 0
'    lnNumPag = 1
'    rtf.Text = ""
'    Encabezado
'    CabGarant
'Else
'    If lnNumLineas > 60 Then
'        lnNumLineas = 0
'        rtf.Text = rtf.Text & gPrnSaltoPagina
'        lnNumPag = lnNumPag + 1
'        CabGarant
'    Else
'        CabGarant
'    End If
'End If
'lsCredGarant = ""
'For I = 1 To Me.fgGarantias.Rows - 1            '   Me.lstGarantias.ListItems.Count
'        If lnNumLineas > 60 Then
'            lnNumLineas = 0
'            rtf.Text = rtf.Text & gPrnSaltoPagina
'            lnNumPag = lnNumPag + 1
'            CabGarant
'        End If
'        If lsCredGarant <> ImpreFormat(fgGarantias.TextMatrix(I, 1), 15) Then
'            If lsCredGarant <> "" Then
'                rtf.Text = rtf.Text & gPrnSaltoLinea
'            End If
'            lnNumLineas = lnNumLineas + 1
'            rtf.Text = rtf.Text + PrnSet("B+") + PrnSet("I+") + ImpreFormat(fgGarantias.TextMatrix(I, 1), 15) + PrnSet("I-") + PrnSet("B-") + gPrnSaltoLinea
'            lnNumLineas = lnNumLineas + 1
'        End If
'        rtf.Text = rtf.Text + PrnSet("C+") + ImpreFormat(fgGarantias.TextMatrix(I, 14), 8) & _
'                                            ImpreFormat(fgGarantias.TextMatrix(I, 2), 30) & _
'                                            ImpreFormat(fgGarantias.TextMatrix(I, 3), 3) & _
'                                            ImpreFormat(fgGarantias.TextMatrix(I, 4), 15) & _
'                                            ImpreFormat(fgGarantias.TextMatrix(I, 5), 40) & gPrnSaltoLinea
'        lnNumLineas = lnNumLineas + 1
'         rtf.Text = rtf.Text + ImpreFormat(fgGarantias.TextMatrix(I, 6), 30, 12) & _
'                               ImpreFormat(fgGarantias.TextMatrix(I, 7), 3) & _
'                               PrnSet("I+") + ImpreFormat(CCur(fgGarantias.TextMatrix(I, 8)), 8, , True) & _
'                               ImpreFormat(CCur(fgGarantias.TextMatrix(I, 9)), 13, , True) & _
'                               ImpreFormat(CCur(fgGarantias.TextMatrix(I, 10)), 15, , True) & _
'                               ImpreFormat(CCur(fgGarantias.TextMatrix(I, 12)), 12, , True) + PrnSet("I-") & _
'                               ImpreFormat(fgGarantias.TextMatrix(I, 13), 3) + PrnSet("C+") + gPrnSaltoLinea
'        lnNumLineas = lnNumLineas + 1
'        lsCredGarant = ImpreFormat(fgGarantias.TextMatrix(I, 1), 8)
'Next
End Sub
Private Sub EncabezadoGeneral(lbObs As Boolean, Optional lbDatosGen As Boolean = True)
'Dim I As Integer
'Dim lsAgencia As String
'Dim lnNumAge As Integer
'Dim lnNumJust As Integer
'
'rtf.Text = rtf.Text + Space(10) + PrnSet("B+") + CabeRepo2(100, "", "REPORTE DE CALIFICACION DE CLIENTES", "AUDITORIA INTERNA", "CMACT", Format(lnNumPag, "0000"), 10) & PrnSet("B-") + gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 7
'rtf.Text = rtf.Text & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'rtf.Text = rtf.Text + Space(10) & PrnSet("B+") & ImpreFormat("CLIENTE        : " & PstaNombre(Trim(lblNomPers), False), 80, 0) & PrnSet("B-") & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'rtf.Text = rtf.Text + Space(10) & ImpreFormat("DOC. IDENTIDAD : " & LblDocNat.Caption, 30, 0) & ImpreFormat("DOC. JURIDICO :" & LblDocJur, 30) & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'If Me.lsActividades.ListCount > 0 Then
'    rtf.Text = rtf.Text + Space(10) & ImpreFormat("Actividade(s) / Sector(es) :", 30, 0) & gPrnSaltoLinea
'    For I = 0 To Me.lsActividades.ListCount - 1
'        rtf.Text = rtf.Text + Space(20) & ImpreFormat(Me.lsActividades.List(I), 80, 0) & gPrnSaltoLinea
'        lnNumLineas = lnNumLineas + 1
'    Next
'Else
'    rtf.Text = rtf.Text & gPrnSaltoLinea
'    lnNumLineas = lnNumLineas + 1
'End If
'If lbObs Then
'    rtf.Text = rtf.Text + Space(10) & String(90, "=") & gPrnSaltoLinea
'    lnNumLineas = lnNumLineas + 1
'    rtf.Text = rtf.Text + Space(10) + PrnSet("12CPI") + PrnSet("B+") + ImpreFormat("Calificación : " & txtCalGen, 30, 50) + PrnSet("15CPI") + PrnSet("B-") + gPrnSaltoLinea
'    lnNumLineas = lnNumLineas + 1
'    rtf.Text = rtf.Text + Space(10) & PrnSet("B+") & PrnSet("I+") & ImpreFormat("Observaciones Generales :", 30, 0) & PrnSet("B-") & PrnSet("I-") & gPrnSaltoLinea
'    rtf.Text = rtf.Text + Space(15) & JustificaTexto(Trim(txtObs), 80, lnNumJust, 15, False) & gPrnSaltoLinea
'    rtf.Text = rtf.Text & gPrnSaltoLinea
'    lnNumLineas = lnNumLineas + lnNumJust
'End If
'rtf.Text = rtf.Text + Space(10) & PrnSet("B+") & "LISTA DE CREDITOS" & PrnSet("B-") & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'rtf.Text = rtf.Text + Space(10) & String(95, "-") & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'rtf.Text = rtf.Text + Space(10) & ImpreFormat("CREDITO", 14) & ImpreFormat("FECHA", 14) & ImpreFormat("SALDO", 10) & ImpreFormat("CAL", 4) & ImpreFormat("NOTA", 4) & ImpreFormat("OBSERVACIONES", 25) & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
'rtf.Text = rtf.Text + Space(10) & String(95, "-") & gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 1
End Sub
Private Sub ImprimeFuentes()
'Dim I As Integer
'Dim lsNumfuente As String
'If Me.chkImpFte.Value = 1 And (Me.chkImpGar.Value = 0 And Me.chkImpCred.Value = 0) Then
'    lnNumLineas = 0
'    lnNumPag = 1
'    rtf.Text = ""
'    Encabezado
'End If
'rtf.Text = rtf.Text & gPrnSaltoLinea
'rtf.Text = rtf.Text + PrnSet("I+") + PrnSet("B+") + "FUENTES DE INGRESOS" + PrnSet("I-") & gPrnSaltoLinea
'rtf.Text = rtf.Text & "------------------" + PrnSet("B-") + gPrnSaltoLinea
'lnNumLineas = lnNumLineas + 2
'For I = 1 To Me.fgFuentes.Rows - 1        'Me.lstfuentes.ListItems.Count
'    If lnNumLineas > 60 Then
'        lnNumLineas = 0
'        rtf.Text = rtf.Text & gPrnSaltoPagina
'        lnNumPag = lnNumPag + 1
'        'Encabezado
'    End If
'    If lsNumfuente <> Trim(Me.fgFuentes.TextMatrix(I, 14)) Then
'        rtf.Text = rtf.Text & gPrnSaltoLinea
'        rtf.Text = rtf.Text & ImpreFormat("Razon Social :" & Me.fgFuentes.TextMatrix(I, 2), 50, 0) & ImpreFormat("Tipo de Fuente : " & Me.fgFuentes.TextMatrix(I, 1), 40) & gPrnSaltoLinea
'        rtf.Text = rtf.Text & ImpreFormat("Dirección    :" & Me.fgFuentes.TextMatrix(I, 3) & "-" & Me.fgFuentes.TextMatrix(I, 4), 90, 0) & gPrnSaltoLinea
'        rtf.Text = rtf.Text & ImpreFormat("Actividad    :" & Me.fgFuentes.TextMatrix(I, 6), 50, 0) & ImpreFormat("Sector : " & Me.fgFuentes.TextMatrix(I, 5), 30) & gPrnSaltoLinea
'        rtf.Text = rtf.Text & ImpreFormat("Cargo        :" & Me.fgFuentes.TextMatrix(I, 7), 50, 0) & gPrnSaltoLinea
'        If Me.fgFuentes.TextMatrix(I, 8) <> "" Then
'            rtf.Text = rtf.Text + PrnSet("I+") & PrnSet("B+") & "[Fuentes Dependientes]" + PrnSet("I-") & gPrnSaltoLinea
'            rtf.Text = rtf.Text & String("50", "-") & gPrnSaltoLinea
'            rtf.Text = rtf.Text & ImpreFormat("Fecha", 16) & ImpreFormat("Ingresos", 12) & ImpreFormat("Gastos", 10) & gPrnSaltoLinea
'            rtf.Text = rtf.Text & String("50", "-") & PrnSet("B-") & gPrnSaltoLinea
'        Else
'            rtf.Text = rtf.Text + PrnSet("I+") & PrnSet("B+") & "[Balance Financiero]" + PrnSet("I-") & gPrnSaltoLinea
'            rtf.Text = rtf.Text & String("50", "-") & gPrnSaltoLinea
'            rtf.Text = rtf.Text & ImpreFormat("Fecha", 15) & ImpreFormat("Ing. Fam", 12) & ImpreFormat("Gas. Fam", 10) & gPrnSaltoLinea
'            rtf.Text = rtf.Text & String("50", "-") & PrnSet("B-") & gPrnSaltoLinea
'        End If
'        lnNumLineas = lnNumLineas + 8
'    End If
'    If Me.fgFuentes.TextMatrix(I, 8) <> "" Then
'        rtf.Text = rtf.Text & ImpreFormat(Me.fgFuentes.TextMatrix(I, 8), 12) & ImpreFormat(CCur(Me.fgFuentes.TextMatrix(I, 9)), 10, 2, True) & ImpreFormat(CCur(Me.fgFuentes.TextMatrix(I, 10)), 10, 2) & gPrnSaltoLinea
'    Else
'        rtf.Text = rtf.Text & ImpreFormat(Me.fgFuentes.TextMatrix(I, 11), 12) & ImpreFormat(CCur(Me.fgFuentes.TextMatrix(I, 12)), 10, 2, True) & ImpreFormat(CCur(Me.fgFuentes.TextMatrix(I, 13)), 10, 2) & gPrnSaltoLinea
'    End If
'    lnNumLineas = lnNumLineas + 1
'    lsNumfuente = Trim(Me.fgFuentes.TextMatrix(I, 14))
'Next
End Sub
Private Sub CabeceraGrid()
'Dim I As Integer
'
'fgCredito.Cols = 11
'fgCredito.Rows = 2
'fgCredito.Clear
'For I = 0 To fgCredito.Cols - 1
'    fgCredito.Row = 0
'    fgCredito.Col = I
'    fgCredito.CellTextStyle = 0
'    fgCredito.CellAlignment = 4
'    fgCredito.CellFontBold = True
'    Select Case I
'        Case 1
'            fgCredito.Text = "Credito"
'        Case 2
'            fgCredito.Text = "Fecha"
'        Case 3
'            fgCredito.Text = "Capital"
'        Case 4
'            fgCredito.Text = "Cal"
'        Case 5
'            fgCredito.Text = "Nota"
'        Case 6
'            fgCredito.Text = "Obs"
'        Case 7
'            fgCredito.Text = "Dias"
'        Case 8
'            fgCredito.Text = "Aut."
'        Case 10
'            fgCredito.Text = "Est."
'    End Select
'Next I
'fgCredito.ColAlignment(0) = 0
'fgCredito.ColAlignment(1) = 1
'fgCredito.ColAlignment(2) = 1
'fgCredito.ColAlignment(3) = 7
'fgCredito.ColAlignment(4) = 4
'fgCredito.ColAlignment(5) = 4
'fgCredito.ColAlignment(6) = 4
'fgCredito.ColAlignment(7) = 4
'fgCredito.ColAlignment(8) = 4
'fgCredito.ColAlignment(10) = 4
'
'fgCredito.ColWidth(0) = 350
'fgCredito.ColWidth(1) = 1300
'fgCredito.ColWidth(2) = 1100
'fgCredito.ColWidth(3) = 1000
'fgCredito.ColWidth(4) = 600
'fgCredito.ColWidth(5) = 500
'fgCredito.ColWidth(6) = 0
'fgCredito.ColWidth(7) = 500
'fgCredito.ColWidth(8) = 450
'fgCredito.ColWidth(9) = 0
'fgCredito.ColWidth(10) = 450

End Sub
Private Sub DatosDetalleGrid(lsCodPers As String, lsCodCta As String, lsFecha As String, _
                            lnCapital As Currency, lsNota As String, lsCal As String, _
                            lsObs As String, lnDiasAtraso As Integer, lsAutorizado As String, lsEstado As String)
'Dim N As Integer
'AdicionaRow fgCredito
'If lsEstado = "0" Then
'    BackColorFg fgCredito, &HC0C0&, False
'End If
'N = fgCredito.Row
'fgCredito.TextMatrix(N, 1) = lsCodCta
'fgCredito.TextMatrix(N, 2) = lsFecha
'fgCredito.TextMatrix(N, 3) = Format(lnCapital, "#0.00")
'fgCredito.TextMatrix(N, 4) = lsCal
'fgCredito.TextMatrix(N, 5) = lsNota
'fgCredito.TextMatrix(N, 6) = lsObs
'fgCredito.TextMatrix(N, 7) = Val(lnDiasAtraso)
'fgCredito.TextMatrix(N, 10) = Trim(lsEstado)
'
'If lsAutorizado = "1" Then
'    fgCredito.TextMatrix(N, 8) = "."
'    fgCredito.Col = 8
'    Set fgCredito.CellPicture = imgList.ListImages(2).Picture
'Else
'    fgCredito.TextMatrix(N, 8) = ""
'    fgCredito.Col = 8
'    Set fgCredito.CellPicture = imgList.ListImages(1).Picture
'End If

End Sub
Private Sub DatosDetalleBase(lsCodPers As String)
Dim SQL As String
Dim rs As New ADODB.Recordset

'CabeceraGrid
'SQL = " SELECT  cCodCta as Credito, dFecCal as Fecha, " _
'    & "         IsNull(cCalAud,'') as Cal, Isnull(cNotaAud,'') as Nota, " _
'    & "         ISNULL(cObsAud,'') as Obs, nSaldoCap as Capital, ISNULL(nDiasAtraso,0) AS nDiasAtraso, ISNULL(cAutorizado,'') AS cAutorizado, cEstado  " _
'    & " FROM    " & gcServerAudit & "AUDPERSDET where cCodPers ='" & lsCodPers & "' " _
'    & " ORDER BY dFecCal "
'
'rs.CursorLocation = adUseClient
'rs.Open SQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'Set rs.ActiveConnection = Nothing
'Do While Not rs.EOF
'    DatosDetalleGrid lsCodPers, Trim(rs!Credito), Format(rs!Fecha, "dd/mm/yyyy"), rs!Capital, _
'                        rs!Nota, rs!Cal, rs!Obs, Val(rs!nDiasAtraso), Trim(rs!cAutorizado), rs!cEstado
'
'    rs.MoveNext
'Loop
'rs.Close
End Sub
Private Sub CabGridCreditos()
'Dim I As Integer
'
'fgCreditos.Cols = 24
'fgCreditos.Rows = 2
'fgCreditos.Clear
'For I = 0 To fgCreditos.Cols - 1
'    fgCreditos.Row = 0
'    fgCreditos.Col = I
'    fgCreditos.CellTextStyle = 0
'    fgCreditos.CellAlignment = 4
'    fgCreditos.CellFontBold = True
'    Select Case I
'        Case 1
'            fgCreditos.Text = "Fecha"
'        Case 2
'            fgCreditos.Text = "Crédito"
'        Case 3
'            fgCreditos.Text = "Agencia"
'        Case 4
'            fgCreditos.Text = "Producto"
'        Case 5
'            fgCreditos.Text = "Estado"
'        Case 6
'            fgCreditos.Text = "Rel"
'        Case 7
'            fgCreditos.Text = "Analista"
'        Case 8
'            fgCreditos.Text = "Monto"
'        Case 9
'            fgCreditos.Text = "Nota1"
'        Case 10
'            fgCreditos.Text = "Nota1 F/M"
'        Case 11
'            fgCreditos.Text = "S.K Actual"
'        Case 12
'            fgCreditos.Text = "S.K F/M"
'        Case 13
'            fgCreditos.Text = "D/M Act"
'        Case 14
'            fgCreditos.Text = "D/M-F/M"
'        Case 15
'            fgCreditos.Text = "Cod.Ant"
'        Case 16
'            fgCreditos.Text = "Moneda "
'        Case 17
'            fgCreditos.Text = "Ref "
'        Case 18
'            fgCreditos.Text = "Motivo Rechazo "
'        Case 19
'            fgCreditos.Text = "Abrev Cred "
'        Case 20
'            fgCreditos.Text = "Cancelac."
'        Case 21
'            fgCreditos.Text = "Estado"
'        Case 22
'            fgCreditos.Text = "Fecha F/M"
'        Case 23
'            fgCreditos.Text = "Monto Aprob"
'    End Select
'Next I
'fgCreditos.ColAlignment(0) = 0
'fgCreditos.ColAlignment(1) = 0
'fgCreditos.ColAlignment(2) = 0
'fgCreditos.ColAlignment(3) = 0
'fgCreditos.ColAlignment(4) = 0
'fgCreditos.ColAlignment(5) = 0
'fgCreditos.ColAlignment(6) = 0
'
'fgCreditos.ColWidth(0) = 350
'fgCreditos.ColWidth(1) = 1100
'fgCreditos.ColWidth(2) = 1400
'fgCreditos.ColWidth(3) = 2500
'fgCreditos.ColWidth(4) = 2800
'fgCreditos.ColWidth(5) = 1500
'fgCreditos.ColWidth(6) = 1000
'fgCreditos.ColWidth(7) = 800
'fgCreditos.ColWidth(8) = 1000
'fgCreditos.ColWidth(9) = 900
'fgCreditos.ColWidth(10) = 1000
'fgCreditos.ColWidth(11) = 1200
'fgCreditos.ColWidth(12) = 1200
'fgCreditos.ColWidth(13) = 900
'fgCreditos.ColWidth(14) = 900
'fgCreditos.ColWidth(15) = 1000
'fgCreditos.ColWidth(16) = 0
'fgCreditos.ColWidth(17) = 800
'fgCreditos.ColWidth(18) = 0
'fgCreditos.ColWidth(19) = 0
'fgCreditos.ColWidth(20) = 1200
'fgCreditos.ColWidth(21) = 900
'fgCreditos.ColWidth(22) = 1200
'fgCreditos.ColWidth(23) = 1200
End Sub
Private Sub CabGridGarantias()
'Dim I As Integer
'Dim lsCabecera As String
'Dim lsAnchos As String
'lsCabecera = " - Crédito -  Persona - Relacion - Tipo Garantia - Descripción - Documento - " _
'            & " Moneda - Tasacion - Realizacion - Por Gravar - Estado - " _
'            & "Gravado - Moneda - Nro. Garantia "
'
'MSHFlex fgGarantias, 15, lsCabecera, " 300 - 1200 - 3500 - 1200 - 2500 - 3000 - 2800 - 900 - 1200 - 1200 - 1200 - 900 - 1200 - 1200 - 900 - 1200", _
'        "L-L-L-L-L-L-L-R-R-R-L-R-R-L-C"
'
'For I = 0 To fgGarantias.Cols - 1
'    fgGarantias.Row = 0
'    fgGarantias.Col = I
'    fgGarantias.CellTextStyle = 0
'    fgGarantias.CellAlignment = 4
'    fgGarantias.CellFontBold = True
'Next
End Sub
Private Sub CabGridFuentes()
'Dim I As Integer
'Dim lsCabecera As String
'Dim lsAnchos As String
'lsCabecera = " - Tipo Fuente - Razon Social - Direccion - Zona - " _
'            & " Sector - Actividad - Cargo - Fecha FD - Ingreso FD - Gastos FD - " _
'            & "Fecha Bal - Ingresos Bal - Gasto Bal - Numfuente"
'
'MSHFlex fgFuentes, 15, lsCabecera, " 300 - 1500 - 2500 - 2500 - 1500 - 1500 - 3500 - 2800 - 1200 - 1200 - 1200 - 1200 - 1200 - 1200 - 1200 ", _
'        "L-L-L-L-L-L-L-L-L-R-R-L-R-R-C"
'
'For I = 0 To fgFuentes.Cols - 1
'    fgFuentes.Row = 0
'    fgFuentes.Col = I
'    fgFuentes.CellTextStyle = 0
'    fgFuentes.CellAlignment = 4
'    fgFuentes.CellFontBold = True
'Next
End Sub
Private Sub ControlesAutorizados()
'If Me.lbAutorizado Then
'    Me.cmdAgregar.Enabled = False
'    Me.cmdEliminar.Enabled = False
'    Me.cmdEliminaGen.Enabled = False
'    Me.cmdnuevo.Enabled = False
'    Me.cmdEditaCred.Enabled = False
'Else
    Me.cmdAgregar.Enabled = True
    Me.cmdEliminar.Enabled = True
    Me.cmdEliminaGen.Enabled = True
    Me.cmdnuevo.Enabled = True
    Me.cmdEditaCred.Enabled = True
'End If
End Sub


