VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColocCalEvalCli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocaciones - Evaluación y Calificación de Clientes"
   ClientHeight    =   7920
   ClientLeft      =   825
   ClientTop       =   975
   ClientWidth     =   9945
   Icon            =   "frmColocCalEvalCli.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc AdoAudGen 
      Height          =   375
      Left            =   7155
      Top             =   105
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
            Picture         =   "frmColocCalEvalCli.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColocCalEvalCli.frx":0ADC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   600
      Left            =   225
      TabIndex        =   42
      Top             =   6840
      Width           =   9240
      Begin VB.CommandButton cmdEliminaGen 
         Caption         =   "&Eliminar"
         Height          =   360
         Left            =   2925
         TabIndex        =   8
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir Planilla"
         Height          =   360
         Left            =   4200
         TabIndex        =   9
         Top             =   180
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7635
         TabIndex        =   43
         Top             =   180
         Width           =   1470
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   360
         Left            =   1440
         TabIndex        =   7
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   360
         Left            =   180
         TabIndex        =   6
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   360
         Left            =   180
         TabIndex        =   10
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   360
         Left            =   1455
         TabIndex        =   11
         Top             =   180
         Width           =   1275
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   690
      Left            =   10050
      TabIndex        =   28
      Top             =   6150
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   1217
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmColocCalEvalCli.frx":12AE
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
      Left            =   6120
      TabIndex        =   40
      Top             =   7710
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin TabDlg.SSTab TabGen 
      Height          =   7185
      Left            =   75
      TabIndex        =   29
      Top             =   360
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   12674
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "frmColocCalEvalCli.frx":132E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Lista de Créditos"
      TabPicture(1)   =   "frmColocCalEvalCli.frx":134A
      Tab(1).ControlEnabled=   -1  'True
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
         Enabled         =   0   'False
         Height          =   2055
         Left            =   7455
         TabIndex        =   41
         Top             =   1845
         Width           =   1815
         Begin MSMask.MaskEdBox mskFechaMes 
            Height          =   315
            Left            =   585
            TabIndex        =   48
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
            TabIndex        =   19
            Top             =   1590
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkImpFte 
            Caption         =   "&Fuentes"
            Height          =   300
            Left            =   300
            TabIndex        =   18
            Top             =   720
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.CheckBox chkImpGar 
            Caption         =   "&Garantías"
            Height          =   345
            Left            =   300
            TabIndex        =   17
            Top             =   420
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CheckBox chkImpCred 
            Caption         =   "&Créditos"
            Height          =   330
            Left            =   300
            TabIndex        =   16
            Top             =   165
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha :"
            Height          =   195
            Left            =   75
            TabIndex        =   47
            Top             =   1215
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar Datos Clientes"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4320
         TabIndex        =   12
         Top             =   600
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
         Left            =   -74805
         TabIndex        =   26
         Top             =   525
         Width           =   9360
         Begin VB.ListBox lsActividades 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            ItemData        =   "frmColocCalEvalCli.frx":1366
            Left            =   4920
            List            =   "frmColocCalEvalCli.frx":1368
            Sorted          =   -1  'True
            TabIndex        =   46
            Top             =   615
            Width           =   4245
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cliente "
            Height          =   195
            Left            =   165
            TabIndex        =   35
            Top             =   285
            Width           =   525
         End
         Begin VB.Label lblDocJur 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   3600
            TabIndex        =   3
            Top             =   615
            Width           =   1275
         End
         Begin VB.Label lblDocNat 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1080
            TabIndex        =   2
            Top             =   615
            Width           =   1395
         End
         Begin VB.Label lblNomPers 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2520
            TabIndex        =   1
            Top             =   225
            Width           =   6630
         End
         Begin VB.Label lblCodPers 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   0
            Top             =   225
            Width           =   1410
         End
         Begin VB.Label lblDocJuridico 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Tributario"
            Height          =   195
            Left            =   2520
            TabIndex        =   34
            Top             =   660
            Width           =   1050
         End
         Begin VB.Label lblDocNatural 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Ident."
            Height          =   195
            Left            =   150
            TabIndex        =   33
            Top             =   660
            Width           =   795
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos de Evaluacion"
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
         Left            =   -74835
         TabIndex        =   32
         Top             =   1680
         Width           =   9240
         Begin MSComDlg.CommonDialog cdlOpen 
            Left            =   7740
            Top             =   4140
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txtObs 
            Enabled         =   0   'False
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
            Top             =   240
            Width           =   8715
            _ExtentX        =   15372
            _ExtentY        =   3651
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            WrapCellPointer =   -1  'True
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
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   3
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
            TabIndex        =   44
            Top             =   2565
            Width           =   1920
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Calificación"
            Height          =   195
            Left            =   7770
            TabIndex        =   39
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
            TabIndex        =   38
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
         Enabled         =   0   'False
         Height          =   2400
         Left            =   165
         TabIndex        =   30
         Top             =   4080
         Width           =   9240
         Begin VB.CommandButton cmdImprimeCalCred 
            Caption         =   "I&mprimir Créditos"
            Height          =   345
            Left            =   4680
            TabIndex        =   24
            Top             =   1980
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdEditaCred 
            Caption         =   "&Modificar"
            Height          =   345
            Left            =   1320
            TabIndex        =   22
            Top             =   1980
            Width           =   1110
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            Height          =   345
            Left            =   180
            TabIndex        =   21
            Top             =   1980
            Width           =   1095
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   345
            Left            =   3480
            TabIndex        =   23
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
            TabIndex        =   25
            Top             =   270
            Width           =   2610
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgCredCalif 
            Height          =   1695
            Left            =   180
            TabIndex        =   20
            ToolTipText     =   "Seleccione el credito para mostrar Observaciones al lado derecho"
            Top             =   180
            Width           =   6150
            _ExtentX        =   10848
            _ExtentY        =   2990
            _Version        =   393216
            Cols            =   5
            FocusRect       =   0
            HighLight       =   2
            AllowUserResizing=   1
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
         Left            =   180
         TabIndex        =   31
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
         TabPicture(0)   =   "frmColocCalEvalCli.frx":136A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fgCreditos"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Garantías"
         TabPicture(1)   =   "frmColocCalEvalCli.frx":1386
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fgGarantias"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Fuentes de Ingreso       "
         TabPicture(2)   =   "frmColocCalEvalCli.frx":13A2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fgFuentes"
         Tab(2).Control(1)=   "chkFuentes"
         Tab(2).ControlCount=   2
         Begin VB.CheckBox chkFuentes 
            Height          =   195
            Left            =   -70980
            TabIndex        =   45
            Top             =   105
            Width           =   210
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgCreditos 
            Height          =   1965
            Left            =   195
            TabIndex        =   13
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
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgFuentes 
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgGarantias 
            Height          =   1965
            Left            =   -74805
            TabIndex        =   14
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
         Left            =   135
         Shape           =   4  'Rounded Rectangle
         Top             =   450
         Width           =   6420
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000006&
         Height          =   930
         Left            =   135
         Shape           =   4  'Rounded Rectangle
         Top             =   435
         Width           =   6390
      End
      Begin VB.Label lblNomPers1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   300
         TabIndex        =   37
         Top             =   990
         Width           =   6090
      End
      Begin VB.Label lblCodPers1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   300
         TabIndex        =   36
         Top             =   660
         Width           =   1410
      End
   End
   Begin MSComctlLib.StatusBar barraestado 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   27
      Top             =   7620
      Width           =   9945
      _ExtentX        =   17542
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
Attribute VB_Name = "frmColocCalEvalCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* COLOCACIONES - CALIFICACION - EVALUACION DE CLIENTES
'Archivo:  frmColocCalEvalCli.frm
'LAYG   :  01/10/2002.
'Resumen:  Registra la evaluacion de los Clientes para la Calificacion

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

Public fnTipoEval As Integer  ' 0=Evaluacion / 1=Revision


Public Sub Inicio(ByVal pbEvaluacion As Boolean)

Dim loConstSis As COMDConstSistema.NCOMConstSistema
Set loConstSis = New COMDConstSistema.NCOMConstSistema
    loConstSis.LeeConstSistema (42) ' Agencia Auditoria
    fdFechaFinMes = loConstSis.LeeConstSistema(gConstSistCierreMesNegocio)
   
Set loConstSis = Nothing

CabeceraGrid
TabGen.Tab = 0
TabPosicion.Tab = 0
'fdFechaFinMes = CDate(ReadVarSis("ADM", "dFecCierreMes"))
mskFechaMes.Text = fdFechaFinMes
If pbEvaluacion = True Then
    fnTipoEval = 0
    Me.Caption = "Colocaciones : Calificacion - Evaluacion de Calificacion"
Else
    fnTipoEval = 1
    Me.Caption = "Colocaciones : Calificacion - Revision de Calificacion"
End If

DatosGenerales
fbNuevo = False
CabGridCreditos
CabGridGarantias
CabGridFuentes
Habilitacion False

Me.Show 1

End Sub

'TODOCOMPLETA   Se Comento para el Caso de Anita que no se puede referenciar al ADODB.Event... por el caso de tener instalado el .Net

'Private Sub adoAudGen_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'If Not pRecordset.EOF And Not pRecordset.BOF Then
'    lblCodPers = pRecordset!CodPers
'    Me.lblNomPers = PstaNombre(pRecordset!Persona, False)
'    lblCodPers1 = pRecordset!CodPers
'    Me.lblNomPers1 = PstaNombre(pRecordset!Persona, False)
'    Me.lblDocJur = IIf(IsNull(pRecordset!DocTri), "", pRecordset!DocTri)
'    Me.lblDocNat = IIf(IsNull(pRecordset!DocNat), "", pRecordset!DocNat)
'    Me.txtCalGen = pRecordset!Cal
'    Me.txtObs = Trim(pRecordset!Obs)
'    Call DatosCalifCredito(Trim(pRecordset!CodPers), fnTipoEval)
'    Me.lsActividades.Clear
'    CabGridCreditos
'    CabGridGarantias
'    CabGridFuentes
'End If
'End Sub

Private Sub cmdAgregar_Click()
Dim lsCta As String
Dim ldFechaCred As Date
Dim lsFecha As String
Dim i As Integer

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
If fgEstadoVigenteCredito(Trim(fgCreditos.TextMatrix(fgCreditos.row, 21))) = False Then
        If Mid(Trim(fgCreditos.TextMatrix(fgCreditos.row, 2)), 6, 3) = "305" Then
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

'If Me.cmdNuevo.Visible Then
'    frmColocCalEvalCliDatos.lbNuevaPers = False
'Else
'    frmColocCalEvalCliDatos.lbNuevaPers = True
'End If
frmColocCalEvalCliDatos.fbNuevo = True
frmColocCalEvalCliDatos.fsCodCta = fgCreditos.TextMatrix(fgCreditos.row, 2)
frmColocCalEvalCliDatos.lblEstado = fgCreditos.TextMatrix(fgCreditos.row, 5)
frmColocCalEvalCli.fnTipoEval = fnTipoEval
If frmColocCalEvalCliDatos.fsCodCta = "" Then
    MsgBox "Seleccione la fila del crédito que desea agregar ", vbInformation, "aviso"
    Exit Sub
End If
frmColocCalEvalCliDatos.fsCodPers = Trim(lblCodPers)
frmColocCalEvalCliDatos.txtNota = IIf(fgCreditos.TextMatrix(fgCreditos.row, 9) = "", "0", fgCreditos.TextMatrix(fgCreditos.row, 9))
frmColocCalEvalCliDatos.txtSaldoCap = fgCreditos.TextMatrix(fgCreditos.row, 11)
frmColocCalEvalCliDatos.txtDiasAtraso = val(fgCreditos.TextMatrix(fgCreditos.row, 13))
'***
frmColocCalEvalCliDatos.fnSaldoCap = fgCreditos.TextMatrix(fgCreditos.row, 11)
frmColocCalEvalCliDatos.fnDiasAtraso = val(fgCreditos.TextMatrix(fgCreditos.row, 13))

frmColocCalEvalCliDatos.txtFechaFM = IIf(fgCreditos.TextMatrix(fgCreditos.row, 22) = "", gdFecDataFM, fgCreditos.TextMatrix(fgCreditos.row, 22))
frmColocCalEvalCliDatos.txtNotaFM = IIf(fgCreditos.TextMatrix(fgCreditos.row, 10) = "", fgCreditos.TextMatrix(fgCreditos.row, 9), fgCreditos.TextMatrix(fgCreditos.row, 10))
frmColocCalEvalCliDatos.txtSaldoCapFM = IIf(fgCreditos.TextMatrix(fgCreditos.row, 12) = "0.00", fgCreditos.TextMatrix(fgCreditos.row, 11), fgCreditos.TextMatrix(fgCreditos.row, 12))
frmColocCalEvalCliDatos.txtDiasAtrasoFM = val(IIf(fgCreditos.TextMatrix(fgCreditos.row, 14) = "", fgCreditos.TextMatrix(fgCreditos.row, 13), fgCreditos.TextMatrix(fgCreditos.row, 14)))
frmColocCalEvalCliDatos.chkVigente.value = 1
frmColocCalEvalCliDatos.fdFechaEval = fdFechaFinMes

'frmAuditDatos.lblCalxDiasAtraso.Caption = fgAsinaCalificacionCredito(fgCredCalif.TextMatrix(fgCredCalif.Row, 1), fgCredCalif.TextMatrix(fgCredCalif.Row, 9), fgCredCalif.TextMatrix(fgCredCalif.Row, 7), fgCredCalif.TextMatrix(fgCredCalif.Row, 3), fnMontoRango, fnTipoCambio, False)
frmColocCalEvalCliDatos.fnTipoEval = fnTipoEval

frmColocCalEvalCliDatos.Show 1

If frmColocCalEvalCliDatos.fbOk Then
    If frmColocCalEvalCliDatos.fbOk Then
        Call DatosCalifCredito(Trim(lblCodPers), fnTipoEval)
    Else
        Call DatosDetalleGrid(Trim(lblCodPers), frmColocCalEvalCliDatos.fsCodCta, frmColocCalEvalCliDatos.fdFechaEval, _
                    frmColocCalEvalCliDatos.txtSaldoCap, frmColocCalEvalCliDatos.txtNota, frmColocCalEvalCliDatos.TxtCalificacion, frmColocCalEvalCliDatos.txtObs, frmColocCalEvalCliDatos.txtDiasAtraso)
       txtObsDet = frmColocCalEvalCliDatos.txtObs
        lsCta = ""
        lsFecha = ""
        For i = 1 To Me.fgCredCalif.rows - 1
            If (lsCta = Me.fgCredCalif.TextMatrix(i, 1) And lsFecha = Me.fgCredCalif.TextMatrix(i, 2)) Then
                MsgBox "Credito ya ha sido ingresado a esa fecha", vbInformation, "Aviso"
                EliminaRow Me.fgCredCalif, i
            Else
                lsCta = fgCredCalif.TextMatrix(i, 1)
                lsFecha = fgCredCalif.TextMatrix(i, 2)
            End If
        Next i
    End If
End If
End Sub

Private Sub cmdBuscar_Click()
Dim lbVacio As Boolean
Dim J As Integer
Dim loDatos As COMNCredito.NCOMColocEval
Dim lrDatos As ADODB.Recordset
lbVacio = False
If Me.lblCodPers = "" Then
    MsgBox "Cliente no Válido", vbInformation, "Aviso"
    Me.TabGen.Tab = 0
    Me.cmdNuevo.SetFocus
    Exit Sub
End If

Call ObtieneDatosCliente(lblCodPers)

Set loDatos = New COMNCredito.NCOMColocEval
    Set lrDatos = loDatos.nObtieneDatosClienteCreditos(lblCodPers, mskFechaMes.Text)
Set loDatos = Nothing
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

frmColocCalEvalCliDatos.fbNuevo = False
frmColocCalEvalCliDatos.fsCodPers = Trim(lblCodPers)
frmColocCalEvalCliDatos.fsCodCta = fgCredCalif.TextMatrix(fgCredCalif.row, 1)
frmColocCalEvalCliDatos.txtFecha = fgCredCalif.TextMatrix(fgCredCalif.row, 2)
frmColocCalEvalCliDatos.txtSaldoCap = fgCredCalif.TextMatrix(fgCredCalif.row, 3)
frmColocCalEvalCliDatos.TxtCalificacion = fgCredCalif.TextMatrix(fgCredCalif.row, 4)
'frmColocCalEvalCliDatos.txtNota = fgCredCalif.TextMatrix(fgCredCalif.Row, 5)
frmColocCalEvalCliDatos.txtObs = fgCredCalif.TextMatrix(fgCredCalif.row, 6)
frmColocCalEvalCliDatos.txtDiasAtraso = val(fgCredCalif.TextMatrix(fgCredCalif.row, 7))
'frmColocCalEvalCliDatos.chkVigente.Value = IIf(Trim(fgCredCalif.TextMatrix(fgCredCalif.Row, 10)) = "1", 1, 0)
frmColocCalEvalCliDatos.fdFechaEval = fdFechaFinMes

frmColocCalEvalCliDatos.fraDatosFM.Visible = False
frmColocCalEvalCliDatos.optSeleccion(0).Visible = False
frmColocCalEvalCliDatos.optSeleccion(1).Visible = False
frmColocCalEvalCliDatos.FraDatosActual.Enabled = True


frmColocCalEvalCliDatos.Show 1
If frmColocCalEvalCliDatos.fbOk Then
    If frmColocCalEvalCliDatos.lbNuevaPers Then
        fgCredCalif.TextMatrix(fgCredCalif.row, 1) = frmColocCalEvalCliDatos.fsCodCta
        fgCredCalif.TextMatrix(fgCredCalif.row, 2) = frmColocCalEvalCliDatos.fdFechaEval
        'fgCredCalif.TextMatrix(fgCredCalif.Row, 3) = frmColocCalEvalCliDatos.lnSaldoCap
        fgCredCalif.TextMatrix(fgCredCalif.row, 4) = frmColocCalEvalCliDatos.TxtCalificacion
        'fgCredCalif.TextMatrix(fgCredCalif.Row, 5) = frmColocCalEvalCliDatos.lsNota
        fgCredCalif.TextMatrix(fgCredCalif.row, 6) = frmColocCalEvalCliDatos.txtObs
    Else
        If frmColocCalEvalCliDatos.fbNuevo = False Then
            Call DatosCalifCredito(Trim(lblCodPers), fnTipoEval)
            txtObsDet = fgCredCalif.TextMatrix(fgCredCalif.rows - 1, 6)

        Else
            Call DatosDetalleGrid(Trim(lblCodPers), frmColocCalEvalCliDatos.fsCodCta, frmColocCalEvalCliDatos.txtFecha, _
                frmColocCalEvalCliDatos.TxtCalificacion, frmColocCalEvalCliDatos.txtObs)
        End If
    End If
End If

End Sub

Private Sub cmdEliminaGen_Click()
'Dim SQL As String
'Dim rs As New ADODB.Recordset
Dim lsUltimaAct As String

If Me.lblCodPers = "" Then
    MsgBox "Código de Cliente no válido para realizar operación", vbInformation, "Aviso"
    Me.TabGen.Tab = 0
    Me.dtgAudGen.SetFocus
    Exit Sub
End If

If AdoAudGen.Recordset.EOF Then
    MsgBox "No existen registros para eliminar", vbInformation, "Aviso"
    Me.cmdNuevo.SetFocus
    Exit Sub
End If

Dim loVal As COMNCredito.NCOMColocEval
Dim lrVal As ADODB.Recordset

Set loVal = New COMNCredito.NCOMColocEval
     Set lrVal = loVal.nObtieneCreditosEvaluadosPersDetalles(Me.lblCodPers, fnTipoEval)
Set loVal = Nothing
If Not lrVal Is Nothing Then
    MsgBox "Cliente posee calificaciones de Creditos ingresados" & Chr(10) & "Primero Elimine la Calificacion de créditos del Cliente", vbInformation, "Aviso"
    Me.TabGen.Tab = 1
    If Me.cmdEliminar.Enabled Then
        Me.cmdEliminar.SetFocus
    End If
    Exit Sub
End If
'lrVal.Close
Set lrVal = Nothing



If MsgBox("Desea elminar el Registro Seleccionado", vbQuestion + vbYesNo, "Aviso") = vbYes Then
Dim loReg As COMNCredito.NCOMColocEval

    Set loReg = New COMNCredito.NCOMColocEval
        Call loReg.nCalifPersonaElimina(Me.lblCodPers, fnTipoEval, lsUltimaAct, False)
    Set loReg = Nothing

    DatosGenerales
    If AdoAudGen.Recordset.EOF Then
        LimpiarControles
    End If
End If
End Sub

Private Sub cmdEliminar_Click()
Dim lsCuenta As String
Dim loNCal As COMNCredito.NCOMColocEval
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim lsMovNro As String

If fgCredCalif.TextMatrix(1, 0) = "" Then
    MsgBox "No existen registros detallados para eliminar", vbInformation, "Aviso"
    Me.cmdAgregar.SetFocus
    Exit Sub
End If
lsCuenta = Trim(fgCredCalif.TextMatrix(fgCredCalif.row, 1))

If Format(fgCredCalif.TextMatrix(fgCredCalif.row, 2), "dd/mm/yyyy") <> Format(fdFechaFinMes, "dd/mm/yyyy") Then
    MsgBox "No se puede eliminar registros de meses anteriores ", vbInformation, "Aviso"
    Me.cmdAgregar.SetFocus
    Exit Sub
End If

If MsgBox("Desea elminar el Crédito :" & Me.fgCredCalif.TextMatrix(fgCredCalif.row, 1) & " de la Fecha : " & fgCredCalif.TextMatrix(fgCredCalif.row, 2), vbQuestion + vbYesNo, "Aviso") = vbYes Then

    'Genera el Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    Set loNCal = New COMNCredito.NCOMColocEval
        Call loNCal.nCalifDetalleElimina(Me.lblCodPers, fnTipoEval, lsCuenta, fgCredCalif.TextMatrix(fgCredCalif.row, 2), "X", lsMovNro, False)
    Set loNCal = Nothing
        Call DatosCalifCredito(Trim(lblCodPers), fnTipoEval)
    If fgCredCalif.TextMatrix(1, 0) = "" Then
        Me.txtObsDet = ""
    End If
End If
End Sub

Private Sub cmdGrabar_Click()
Dim loNCal As COMNCredito.NCOMColocEval
Dim objPista As COMManejador.Pista 'MAVM: 15/08/2008 COMNAuditoria
Set objPista = New COMManejador.Pista '
Dim loContFunct As COMNContabilidad.NCOMContFunciones  'NContFunciones
Dim lsMovNro As String

'Dim I As Integer
'Dim lsCalGen As String
'Dim lnCalAux As Integer
'Dim lsAut As String
'On Error GoTo ErrorGrabar
Dim lsEvalCalif As String
Dim lsUltimaAct As String
Dim lsEvalObs As String

lsEvalObs = Me.txtObs

If Valida = False Then Exit Sub


If MsgBox("Desea Grabar la Información", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    If fbNuevo Then
    
        Set loContFunct = New COMNContabilidad.NCOMContFunciones   'NContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        Set loNCal = New COMNCredito.NCOMColocEval
            lsEvalCalif = loNCal.nCalificaCabecera(lblCodPers.Caption)
            Call loNCal.nCalifPersonaNuevo(lblCodPers.Caption, fnTipoEval, lsEvalCalif, lsUltimaAct, lsEvalObs, "", False)
        'MAVM 15/08/2008
        objPista.InsertarPista "191150", GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, "1", txtObs.Text
            
        Set loNCal = Nothing
        Call cmdCancelar_Click
    Else
    
        Set loContFunct = New COMNContabilidad.NCOMContFunciones   'NContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        Set loNCal = New COMNCredito.NCOMColocEval
            lsEvalCalif = loNCal.nCalificaCabecera(lblCodPers.Caption)
            Call loNCal.nCalifPersonaModifica(lblCodPers.Caption, fnTipoEval, lsEvalCalif, lsUltimaAct, lsEvalObs, False)
        Set loNCal = Nothing
        
        'MAVM 15/08/2008
        objPista.InsertarPista "191150", GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, "2", txtObs.Text
        
    End If
    Habilitacion False
    DatosGenerales
    If fbNuevo = False Then
         Me.AdoAudGen.Recordset.Move lnPos - 1
         lnPos = 0
    End If
    dtgAudGen.SetFocus
    fbNuevo = False
End If

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
'If Me.fgCredCalif.TextMatrix(1, 0) = "" Then
'    MsgBox "No existen información para realizar la operacion", vbInformation, "Aviso"
'    Exit Sub
'End If
'rtf.Text = ""
'lnNumLineas = 0: lnNumPag = 1
'lsCodCred = ""
'EncabezadoGeneral False
'For I = 1 To Me.fgCredCalif.Rows - 1
'
'    rtf.Text = rtf.Text + Space(10) + ImpreFormat(IIf(lsCodCred <> Trim(fgCredCalif.TextMatrix(I, 1)), fgCredCalif.TextMatrix(I, 1), ""), 12) & ImpreFormat(Format(fgCredCalif.TextMatrix(I, 2), "dd/mm/yyyy"), 12) & _
'                ImpreFormat(CCur(fgCredCalif.TextMatrix(I, 3)), 10, 2, True) & ImpreFormat(Val(fgCredCalif.TextMatrix(I, 4)), 7, 0, False) & _
'                ImpreFormat(Val(fgCredCalif.TextMatrix(I, 5)), 6, 0, False) + Space(3) + JustificaTexto(Trim(fgCredCalif.TextMatrix(I, 6)), 40, lnLineasJust, 67, False) + gPrnSaltoLinea
'
'    lnNumLineas = lnNumLineas + lnLineasJust
'    If lnNumLineas > 60 Then
'        lnNumLineas = 0
'        lnNumPag = lnNumPag + 1
'        EncabezadoGeneral False
'    End If
'    lsCodCred = Trim(fgCredCalif.TextMatrix(I, 1))
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
'If Me.fgCredCalif.TextMatrix(1, 0) <> "" Then
'    For I = 1 To Me.fgCredCalif.Rows - 1
'        rtf.Text = rtf.Text + Space(10) + ImpreFormat(IIf(lsCodCred <> Trim(fgCredCalif.TextMatrix(I, 1)), fgCredCalif.TextMatrix(I, 1), ""), 12) & ImpreFormat(Format(fgCredCalif.TextMatrix(I, 2), "dd/mm/yyyy"), 12) & _
'                    ImpreFormat(CCur(fgCredCalif.TextMatrix(I, 3)), 10, 2, True) & ImpreFormat(Val(fgCredCalif.TextMatrix(I, 4)), 7, 0, False) & _
'                    ImpreFormat(Val(fgCredCalif.TextMatrix(I, 5)), 6, 0, False) + Space(3) + JustificaTexto(Trim(fgCredCalif.TextMatrix(I, 6)), 40, lnLineasJust, 67, False) + gPrnSaltoLinea
'
'        lnNumLineas = lnNumLineas + lnLineasJust
'        If lnNumLineas > 60 Then
'            lnNumLineas = 0
'            lnNumPag = lnNumPag + 1
'            EncabezadoGeneral False
'        End If
'        lsCodCred = Trim(fgCredCalif.TextMatrix(I, 1))
'    Next
'End If
'frmPrevio.Previo rtf, "Reporte De Auditoria Interna", False, 66
End Sub

Private Sub CmdModificar_Click()

If Me.AdoAudGen.Recordset.EOF Then
    MsgBox "No existen datos para modificar", vbInformation, "Aviso"
    cmdNuevo.SetFocus
    Exit Sub
End If
If Me.lblCodPers = "" Then
    MsgBox "Datos de Cliente no Válido", vbInformation, "Aviso"
    cmdNuevo.SetFocus
    Exit Sub
End If

frmColocCalEvalCliDatos.lbNuevaPers = False

Habilitacion True
If Not AdoAudGen.Recordset.EOF Then
    lnPos = Me.AdoAudGen.Recordset.Bookmark
End If
fbNuevo = False
Me.TabGen.Tab = 0
Me.txtObs.SetFocus

End Sub

Private Sub cmdNuevo_Click()

Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String, lsPersDocId As String, lsPersDocTrib As String
Dim loExiste As COMNCredito.NCOMColocEval
Dim lbExiste As Boolean
Dim lrCreditos  As ADODB.Recordset

'On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
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
    lblCodPers1.Caption = Trim(lsPersCod)
    lblNomPers.Caption = PstaNombre(lsPersNombre, False)
    lblNomPers1.Caption = PstaNombre(lsPersNombre, False)
    lblDocnat.Caption = lsPersDocId
    lblDocJur.Caption = lsPersDocTrib
    
    Set loExiste = New COMNCredito.NCOMColocEval
        lbExiste = loExiste.nVerifExisteEvaluacion(lsPersCod, fnTipoEval)
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

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub dtgAudGen_HeadClick(ByVal ColIndex As Integer)
If Not AdoAudGen.Recordset Is Nothing Then
    If Not AdoAudGen.Recordset.EOF Then
        AdoAudGen.Recordset.Sort = dtgAudGen.Columns(ColIndex).DataField
    End If
End If
End Sub

Private Sub dtgAudGen_KeyPress(KeyAscii As Integer)
Dim lsCriterio As String

KeyAscii = Letras(KeyAscii)
lsCriterio = " Persona LIKE '" & Chr(KeyAscii) & "*'"
If Not (AdoAudGen.Recordset.BOF And AdoAudGen.Recordset.EOF) Then
    If Asc(Mid(AdoAudGen.Recordset!Persona, 1, 1)) <= KeyAscii Then
        BuscaDato lsCriterio, AdoAudGen.Recordset, 1, False
    Else
        BuscaDato lsCriterio, AdoAudGen.Recordset, 1, False
    End If
End If

End Sub

Private Sub dtgAudGen_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

'If Not pRecordset.EOF And Not pRecordset.BOF Then
'    lblCodPers = pRecordset!CodPers
'    Me.lblNomPers = PstaNombre(pRecordset!Persona, False)
'    lblCodPers1 = pRecordset!CodPers
'    Me.lblNomPers1 = PstaNombre(pRecordset!Persona, False)
'    Me.lblDocJur = IIf(IsNull(pRecordset!DocTri), "", pRecordset!DocTri)
'    Me.lblDocNat = IIf(IsNull(pRecordset!DocNat), "", pRecordset!DocNat)
'    Me.txtCalGen = pRecordset!Cal
'    Me.txtObs = Trim(pRecordset!Obs)
'    Call DatosCalifCredito(Trim(pRecordset!CodPers), fnTipoEval)
'    Me.lsActividades.Clear
'    CabGridCreditos
'    CabGridGarantias
'    CabGridFuentes
'End If

End Sub

Private Sub fgCredCalif_Click()
'Dim SQL As String
'If fgCredCalif.TextMatrix(1, 0) <> "" Then
'    If fgCredCalif.Col = 8 Then
'        If lbAutorizado Then
'            If fgCredCalif.TextMatrix(fgCredCalif.Row, 8) = "." Then
'                fgCredCalif.TextMatrix(fgCredCalif.Row, 8) = ""
'                Set fgCredCalif.CellPicture = imgList.ListImages(1).Picture
'                SQL = "UPDATE " & gcServerAudit & "AudPersDet SET cAutorizado='0', dFecMod='" & FechaHora(gdFecSis) & "', cUsuAut='" & gsCodUser & "'   WHERE cCodCta='" & fgCredCalif.TextMatrix(fgCredCalif.Row, 1) & "' and cCodPers='" & Trim(lblCodPers) & "' and dFecCal ='" & Format(fgCredCalif.TextMatrix(fgCredCalif.Row, 2), "mm/dd/yyyy") & "' "
'                dbCmact.Execute SQL
'            Else
'                fgCredCalif.TextMatrix(fgCredCalif.Row, 8) = "."
'                Set fgCredCalif.CellPicture = imgList.ListImages(2).Picture
'                SQL = "UPDATE " & gcServerAudit & "AudPersDet SET cAutorizado='1', dFecMod='" & FechaHora(gdFecSis) & "', cUsuAut='" & gsCodUser & "'  WHERE cCodCta='" & fgCredCalif.TextMatrix(fgCredCalif.Row, 1) & "' and cCodPers='" & Trim(lblCodPers) & "' and dFecCal ='" & Format(fgCredCalif.TextMatrix(fgCredCalif.Row, 2), "mm/dd/yyyy") & "' "
'                dbCmact.Execute SQL
'
'            End If
'        Else
'            MsgBox "Permiso denegado para Aprobar calificaciones", vbInformation, "Aviso"
'        End If
'    End If
'End If
End Sub
Private Sub fgCredCalif_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim Col As Integer
'Dim SQL As String
'If KeyCode = 32 Then
'    If lbAutorizado Then
'        Col = Me.fgCredCalif.Col
'        If fgCredCalif.TextMatrix(fgCredCalif.Row, 8) = "." Then
'            fgCredCalif.TextMatrix(fgCredCalif.Row, 8) = ""
'            fgCredCalif.Col = 8
'            Set fgCredCalif.CellPicture = imgList.ListImages(1).Picture
'            SQL = "UPDATE " & gcServerAudit & "AudPersDet SET cAutorizado='0', dFecMod='" & FechaHora(gdFecSis) & "', cUsuAut='" & gsCodUser & "'   WHERE cCodCta='" & fgCredCalif.TextMatrix(fgCredCalif.Row, 1) & "' and cCodPers='" & Trim(lblCodPers) & "' and dFecCal ='" & Format(fgCredCalif.TextMatrix(fgCredCalif.Row, 2), "mm/dd/yyyy") & "' "
'            dbCmact.Execute SQL
'        Else
'            fgCredCalif.TextMatrix(fgCredCalif.Row, 8) = "."
'            fgCredCalif.Col = 8
'            Set fgCredCalif.CellPicture = imgList.ListImages(2).Picture
'            SQL = "UPDATE " & gcServerAudit & "AudPersDet SET cAutorizado='1', dFecMod='" & FechaHora(gdFecSis) & "', cUsuAut='" & gsCodUser & "'  WHERE cCodCta='" & fgCredCalif.TextMatrix(fgCredCalif.Row, 1) & "' and cCodPers='" & Trim(lblCodPers) & "' and dFecCal ='" & Format(fgCredCalif.TextMatrix(fgCredCalif.Row, 2), "mm/dd/yyyy") & "' "
'            dbCmact.Execute SQL
'
'        End If
'        Me.fgCredCalif.Col = Col
'    Else
'        MsgBox "Permiso denegado para Aprobar calificaciones", vbInformation, "Aviso"
'    End If
'End If
End Sub

Private Sub fgCredCalif_RowColChange()
    txtObsDet = fgCredCalif.TextMatrix(fgCredCalif.row, 6)
End Sub

Private Sub fgCreditos_DblClick()
If Me.fgCreditos.TextMatrix(1, 0) <> "" Then
    Me.barraestado.Panels(1).Text = "Cargando Datos de Credito. por Favor Espere"
    fgCreditos.Enabled = False
    If Mid(Trim(fgCreditos.TextMatrix(fgCreditos.row, 2)), 6, 3) = "305" Then ' Prendario
        Call frmPigConsulta.MuestraContratoPosicion(fgCreditos.TextMatrix(fgCreditos.row, 2), lblCodPers)
    Else
        Call frmCredConsulta.ConsultaCliente(fgCreditos.TextMatrix(fgCreditos.row, 2))
    End If
    fgCreditos.Enabled = True
    'barraestado.Panels(1).Text = "Evaluación y Clasificacion de Colocaciones"
    'hacemos la conexion a la base puesto que la cierra el formulario que se ha llamado anteriormente
    'AbreConexion
    AdoAudGen.ConnectionString = gsConnection
End If

End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Me.cmdNuevo.Visible = False Then
    If MsgBox("Aun no se ha grabado la informacion" & gPrnSaltoLinea & "Desea Salir sin grabar la información ??", vbQuestion + vbYesNo) = vbYes Then
        Cancel = 0
    Else
        Cancel = 1
    End If
End If
End Sub

Private Sub ObtieneDatosCliente(ByVal psCodPers As String)
'Dim Agencia As String
'Dim SQL As String
'Dim lsFecha As String
'Dim J As Integer
Dim lnItem As Integer
Dim lnItemPinta As Integer
'
'On Error GoTo ErrorConexion

CabGridCreditos
CabGridGarantias
CabGridFuentes
lsActividades.Clear
fnTotalCred = 0
fnTotalGar = 0
fnTotalFte = 0

    Call MuestraDatosClienteCreditos(psCodPers)
    If Me.chkFuentes.value = 1 Then
        DatosFuentes psCodPers
    End If
    barraestado.Panels(2).Text = ""
    DoEvents
'
'
''Colorea las Evaluaciones de creditos Vigentes
If fgCreditos.rows = 0 Then Exit Sub

For lnItem = 1 To fgCreditos.rows - 1
    If fgEstadoVigenteCredito(Trim(fgCreditos.TextMatrix(fgCreditos.row, 21))) = True Then
 
        For lnItemPinta = 1 To fgCredCalif.rows - 1
            If Trim(fgCreditos.TextMatrix(lnItem, 2)) = Trim(fgCredCalif.TextMatrix(lnItemPinta, 1)) Then
                fgCredCalif.row = lnItemPinta
                fgCredCalif.col = 0
                fgCredCalif.CellBackColor = &HE1E1C0
                BackColorFg fgCredCalif, "&H00E0E0A0", True
                fgCredCalif.CellFontBold = True
            End If
        Next
    End If
    If Trim(fgCreditos.TextMatrix(lnItem, 21) = "1") Or Trim(fgCreditos.TextMatrix(lnItem, 21) = "4") Or Trim(fgCreditos.TextMatrix(lnItem, 21) = "6") Or Trim(fgCreditos.TextMatrix(lnItem, 21) = "7") Then
        For lnItemPinta = 1 To fgCredCalif.rows - 1
            If Trim(fgCreditos.TextMatrix(lnItem, 2)) = Trim(fgCredCalif.TextMatrix(lnItemPinta, 1)) Then
                fgCredCalif.row = lnItemPinta
                fgCredCalif.col = 0
                fgCredCalif.CellBackColor = &HCABC2
                BackColorFg fgCredCalif, "&HCABD0", True
                fgCredCalif.CellFontBold = True
            End If
        Next
    End If
    If fgCreditos.TextMatrix(lnItem, 21) = "V" Or fgCreditos.TextMatrix(lnItem, 21) = "W" Then
        For lnItemPinta = 1 To fgCredCalif.rows - 1
            If Trim(fgCreditos.TextMatrix(lnItem, 2)) = Trim(fgCredCalif.TextMatrix(lnItemPinta, 1)) Then
                fgCredCalif.row = lnItemPinta
                fgCredCalif.col = 0
                fgCredCalif.CellBackColor = &HCAB10
                'BackColorFg fgCredCalif, "&HCABA0", True
                fgCredCalif.CellFontBold = True
            End If
        Next
    End If

Next


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
Dim loDatCred As COMNCredito.NCOMColocEval
Dim lrData As ADODB.Recordset

Dim lnCorreCred As Integer
Dim lsMonto As String
Dim lsFechaCred  As String
Dim lsAgencia As String

barraestado.Panels(1).Text = "Cargando información de Creditos. Por Favor Espere..."

Set loDatCred = New COMNCredito.NCOMColocEval
    Set lrData = loDatCred.nObtieneDatosClienteCreditos(psCodPers, mskFechaMes.Text)
Set loDatCred = Nothing

'Total = lrData.RecordCount
If lrData Is Nothing Then Exit Sub

If Not (lrData.BOF And lrData.EOF) Then
    Do While Not lrData.EOF
        'I = I + 1
        
        'lsAgenciaCod = Mid(lrdata!cCodCta, 1, 2)
        'lsAgencia = Tablacod("47", "112" & lsAgenciaCod)
        lnCorreCred = lnCorreCred + 1
        
        AdicionaRow fgCreditos
        lnCorreCred = fgCreditos.row
        lsFechaCred = IIf(IsNull(lrData!dVigencia), "01/01/1950", Format(Trim(lrData!dVigencia), "dd/mm/yyyy"))
        Select Case Trim(lrData!nPrdEstado)
            Case gColocEstSolic, gColocEstSug
                lsMonto = IIf(IsNull(lrData!nMontoCol), "0.00", Format(Trim(lrData!nMontoCol), "####0.00"))
                lsFechaCred = IIf(IsNull(lrData!dVigencia), "", Format(Trim(lrData!dVigencia), "dd/mm/yyyy"))
            'Case gColocEstRech
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
            Case gColocEstAprob
                lsMonto = IIf(IsNull(lrData!nMontoCol), "0.00", Format(Trim(lrData!nMontoCol), "####0.00"))
                lsFechaCred = IIf(IsNull(lrData!dVigencia), "", Format(Trim(lrData!dVigencia), "dd/mm/yyyy"))
            Case gColocEstVigNorm, gColocEstVigMor, gColocEstVigVenc, gColocEstRefNorm, gColocEstRefMor, gColocEstRefVenc
                lsMonto = IIf(IsNull(lrData!nMontoCol), "0.00", Format(Trim(lrData!nMontoCol), "####0.00"))
                lsFechaCred = IIf(IsNull(lrData!dVigencia), "", Format(Trim(lrData!dVigencia), "dd/mm/yyyy"))
            Case gColocEstCancelado
                lsMonto = IIf(IsNull(lrData!nMontoCol), "0.00", Format(Trim(lrData!nMontoCol), "####0.00"))
                lsFechaCred = IIf(IsNull(lrData!dVigencia), "", Format(Trim(lrData!dVigencia), "dd/mm/yyyy"))            'Case "H"
            '***** Predarios  *****
            Case gColPEstDesem, gColPEstVenci, gColPEstPRema, gColPEstRenov
                lsMonto = IIf(IsNull(lrData!nMontoCol), "0.00", Format(Trim(lrData!nMontoCol), "####0.00"))
                lsFechaCred = IIf(IsNull(lrData!dVigencia), "", Format(Trim(lrData!dVigencia), "dd/mm/yyyy"))
            '    '***** Judiciales  *****
            Case gColocEstRecVigJud, gColocEstRecVigCast
                lsMonto = IIf(IsNull(lrData!nMontoCol), "0.00", Format(Trim(lrData!nMontoCol), "####0.00"))
                lsFechaCred = IIf(IsNull(lrData!dVigencia), "", Format(Trim(lrData!dVigencia), "dd/mm/yyyy"))
            'Case Else
            '    lsFechaCred = IIf(IsNull(rs!dAsignacion), "NI       ", Format(Trim(rs!dAsignacion), "dd/mm/yyyy"))
        End Select
        
        fgCreditos.TextMatrix(lnCorreCred, 1) = lsFechaCred
        fgCreditos.TextMatrix(lnCorreCred, 2) = Trim(lrData!cCtaCod)
        fgCreditos.TextMatrix(lnCorreCred, 3) = Trim(IIf(IsNull(lrData!cNomAgencia), "", lrData!cNomAgencia))
        fgCreditos.TextMatrix(lnCorreCred, 4) = fgProductoCreditoTipo(lrData!cCtaCod)

        If fgEstadoVigenteCredito(lrData!nPrdEstado) = False Then
                fgCreditos.TextMatrix(lnCorreCred, 5) = "CANCELADO"
        Else
                fgCreditos.TextMatrix(lnCorreCred, 5) = "VIGENTE"
        End If
        If lrData!nPrdEstado = gColocEstRecVigJud Then
                fgCreditos.TextMatrix(lnCorreCred, 5) = "JUDICIAL"
        ElseIf lrData!nPrdEstado = gColocEstRecVigCast Then
                fgCreditos.TextMatrix(lnCorreCred, 5) = "CASTIGADO"
        'Else
        '    fgCreditos.TextMatrix(lnCorreCred, 5) = Tablacod(IIf(Mid(rs!cCodCta, 3, 3) = "305", "46", "26"), Trim(rs!cEstado))
        End If
        'fgCreditos.TextMatrix(lnCorreCred, 6) = Tablacod("25", Trim(rs!cRelaCta))
        'fgCreditos.TextMatrix(lnCorreCred, 7) = Trim(lrData!cCodAnalista)
        fgCreditos.TextMatrix(lnCorreCred, 8) = lsMonto
        fgCreditos.TextMatrix(lnCorreCred, 9) = Trim(IIf(IsNull(lrData!cNota1), 0, lrData!cNota1))
        fgCreditos.TextMatrix(lnCorreCred, 10) = Trim(IIf(IsNull(lrData!cNota1FM), 0, lrData!cNota1FM))
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
        fgCreditos.TextMatrix(lnCorreCred, 22) = IIf(IsNull(lrData!dfecha), "", Format(lrData!dfecha, "dd/mm/yyyy"))
        
        If fgEstadoVigenteCredito(lrData!nPrdEstado) = True Then
            DatosGarantias lrData!cCtaCod
            fgCreditos.row = fgCreditos.row
            fgCreditos.col = 0
            fgCreditos.CellBackColor = &HE1E1C0
            fgCreditos.CellFontBold = True
            If lrData!nPrdPersRelac = gColRelPersTitular Then
                BackColorFg fgCreditos, "&H00E0E0A0", True
            End If
        End If
        If lrData!nPrdEstado = gColPEstDesem Or Trim(lrData!nPrdEstado) = gColPEstVenci Or Trim(lrData!nPrdEstado) = gColPEstRenov Or Trim(lrData!nPrdEstado) = gColPEstPRema Then
            fgCreditos.row = fgCreditos.row
            fgCreditos.col = 0
            fgCreditos.CellBackColor = &HCABC2
            fgCreditos.CellFontBold = True
           If lrData!nPrdPersRelac = gColRelPersTitular Then
               BackColorFg fgCreditos, "&HCABD0", True
            End If
        End If
        If lrData!nPrdEstado = gColocEstRecVigJud Or lrData!nPrdEstado = gColocEstRecVigCast Then
            fgCreditos.row = fgCreditos.row
            fgCreditos.col = 0
            fgCreditos.CellBackColor = &HCAB10
            fgCreditos.CellFontBold = True
            If lrData!nPrdPersRelac = gColRelPersTitular Then
                BackColorFg fgCreditos, "&HCABA0", True
            End If
        End If
        
        'Me.barra.Value = (I / Total) * 100
        DoEvents
        lrData.MoveNext
    Loop
End If
lrData.Close
Set lrData = Nothing
barraestado.Panels(1).Text = " "
End Sub

Private Sub DatosGarantias(ByVal psCodCta As String)

Dim loDatGar As COMNCredito.NCOMColocEval
Dim rs As ADODB.Recordset
Dim lsCredGarant As String
Dim Total As Integer, i As Integer
Dim n As Integer

''TabPosicion.Tab = 1

Set loDatGar = New COMNCredito.NCOMColocEval
    Set rs = loDatGar.nObtieneDatosClienteGarantias(psCodCta)
Set loDatGar = Nothing
If rs Is Nothing Then
    Exit Sub
End If
Total = rs.RecordCount
lsCredGarant = ""
Do While Not rs.EOF
    i = i + 1
    AdicionaRow fgGarantias
    n = fgGarantias.row
    If lsCredGarant <> Trim(rs!cCtaCod) Then
        'BackColorFg fgGarantias, "&H00E0E0E0"
        fgGarantias.TextMatrix(n, 1) = Trim(rs!cCtaCod)
    End If
    fgGarantias.TextMatrix(n, 2) = Trim(PstaNombre(rs!cPersNombre, False))
    fgGarantias.TextMatrix(n, 3) = Trim(rs!nPrdPersRelac)
    fgGarantias.TextMatrix(n, 4) = Trim(rs!TipoGarantia)
    fgGarantias.TextMatrix(n, 5) = Trim(rs!Descripcion)
    'fgGarantias.TextMatrix(N, 6) = Tablacod("49", rs!cDocGarant)  'Trim(rs!TipoDocGar)
    fgGarantias.TextMatrix(n, 7) = Trim(IIf(Trim(rs!Moneda) = "1", "MN", "ME"))
    fgGarantias.TextMatrix(n, 8) = Format(rs!Tasacion, "#,#0.00")
    fgGarantias.TextMatrix(n, 9) = Format(rs!Realizacion, "#,#0.00")
    fgGarantias.TextMatrix(n, 10) = Format(rs!PorGravar, "#,#0.00")
    fgGarantias.TextMatrix(n, 11) = Trim(rs!nEstado)
    fgGarantias.TextMatrix(n, 12) = Format(rs!Gravado, "#,#0.00")
    fgGarantias.TextMatrix(n, 13) = Trim(IIf(rs!MonedaGC = "1", "MN", "ME"))
    fgGarantias.TextMatrix(n, 14) = Trim(rs!cNumGarant)
    lsCredGarant = Trim(rs!cCtaCod)
    Me.barra.value = (i / Total) * 100
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub
Private Sub DatosFuentes(ByVal psCodPers As String)

Dim loDatFte As COMNCredito.NCOMColocEval
Dim rs As New ADODB.Recordset
Dim Total As Integer, i As Integer
Dim n As Integer
Dim lsNumfuente As String

Set loDatFte = New COMNCredito.NCOMColocEval
    Set rs = loDatFte.nObtieneDatosFuentesIngreso(psCodPers)
Set loDatFte = Nothing

Total = rs.RecordCount

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
TabPosicion.Tab = 2
barraestado.Panels(1).Text = "Cargando información de Fuentes Ingreso.  Por Favor Espere..."
lsNumfuente = ""

Do While Not rs.EOF
    i = i + 1
    AdicionaRow fgFuentes
    n = fgFuentes.row
    If lsNumfuente <> Trim(rs!cNumFuente) Then
        BackColorFg fgFuentes, "&H00E0E0E0"
        fgFuentes.TextMatrix(n, 1) = Trim(rs!TipoFte)
        fgFuentes.TextMatrix(n, 2) = Trim(rs!cRazonSocial)
        fgFuentes.TextMatrix(n, 3) = Trim(rs!cDireccion)
        fgFuentes.TextMatrix(n, 4) = Trim(rs!Zona)
        'fgFuentes.TextMatrix(N, 5) = Trim(rs!Sector)
        'fgFuentes.TextMatrix(N, 6) = Trim(rs!Actividad)
        fgFuentes.TextMatrix(n, 7) = Trim(rs!Cargo)
    End If
    'fgFuentes.TextMatrix(N, 8) = IIf(IsNull(rs!FDFECHA), "", Format(rs!FDFECHA, "dd/mm/yyyy"))
    'fgFuentes.TextMatrix(N, 9) = Format(rs!Ingreso, "#,#0.00")
    'fgFuentes.TextMatrix(N, 10) = Format(rs!Gastos, "#,#0.00")
    'fgFuentes.TextMatrix(N, 11) = IIf(IsNull(rs!BalFecha), "", Format(rs!BalFecha, "dd/mm/yyyy"))
    'fgFuentes.TextMatrix(N, 12) = Format(rs!BalIngFam, "#,#0.00")
    'fgFuentes.TextMatrix(N, 13) = Format(rs!BalGasFam, "#,#0.00")
    fgFuentes.TextMatrix(n, 14) = Trim(rs!cNumFuente)
    lsNumfuente = Trim(rs!cNumFuente)
    Me.barra.value = (i / Total) * 100
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
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
Valida = True
If Me.lblCodPers = "" Then
    MsgBox "Codigo de Cliente no Válido", vbInformation, "aviso"
    If Me.cmdNuevo.Enabled And Me.cmdNuevo.Visible Then
        Me.cmdNuevo.SetFocus
    End If
    Valida = False
    Exit Function
End If
If Me.txtObs = "" Then
    MsgBox "No ha ingresado observaciones a Cliente", vbInformation, "Aviso"
    Me.TabGen.Tab = 0
    Me.txtObs.SetFocus
    Valida = False
    Exit Function
End If

'If Me.fgCredCalif.TextMatrix(1, 0) = "" Then
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
    
    Me.cmdNuevo.Visible = Not lbHab
    Me.cmdCancelar.Visible = lbHab
    Me.cmdModificar.Visible = Not lbHab
    
    Me.txtObs.Locked = Not lbHab
    Me.dtgAudGen.Enabled = Not lbHab
    
    Me.cmdEliminaGen.Enabled = Not lbHab
    Me.cmdImprimir.Enabled = Not lbHab
    
    txtObs.Enabled = lbHab
    cmdBuscar.Enabled = lbHab
    Frame3.Enabled = lbHab
    
    
End Sub

Private Sub DatosGenerales()
Dim lsSQL As String
Dim loConex As COMConecta.DCOMConecta
Dim lsCadenaConex As String

Set loConex = New COMConecta.DCOMConecta
    loConex.AbreConexion
    lsCadenaConex = loConex.CadenaConexion
Set loConex = Nothing

'Set dtgAudGen.DataSource = lrDatCalif

   lsSQL = "SELECT   Per.CPersNombre Persona, CEval.cEvalCalif Cal, " _
        & " DocNat = (Select ISNULL(PerID.cPersIDnro,'') From PersId PerID " _
        & "          Where PerId.cPersCod = Per.cPersCod AND PerId.cPersIDTpo ='" & gPersIdDNI & "') ,  " _
        & " DocTri =(Select ISNULL(PerID.cPersIDnro,'') From PersId PerID " _
        & "          Where PerId.cPersCod = Per.cPersCod AND PerId.cPersIDTpo ='" & gPersIdRUC & "') ,  " _
        & " Per.cPersCod CodPers, ISNULL(CEval.cEvalObs,'') AS Obs " _
        & " FROM ColocEvalCalif CEval JOIN Persona Per ON Per.cPersCod = CEval.cPersCod " _
        & " WHERE nEvalTipo = " & fnTipoEval & " ORDER BY Per.cPersNombre "

AdoAudGen.CommandType = adCmdText
AdoAudGen.CursorType = adOpenStatic
AdoAudGen.RecordSource = lsSQL
AdoAudGen.ConnectionString = lsCadenaConex 'gsConnection
AdoAudGen.Refresh
Set dtgAudGen.DataSource = AdoAudGen

End Sub

Private Sub LimpiarControles()
Me.lblCodPers = ""
Me.lblCodPers1 = ""
Me.lblNomPers = ""
Me.lblNomPers1 = ""
Me.lblDocnat = ""
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
Dim i As Integer
'
fgCredCalif.cols = 11
fgCredCalif.rows = 2
fgCredCalif.Clear

For i = 0 To fgCredCalif.cols - 1
    fgCredCalif.row = 0
    fgCredCalif.col = i
    fgCredCalif.CellTextStyle = 0
    fgCredCalif.CellAlignment = 4
    fgCredCalif.CellFontBold = True
    Select Case i
        Case 1
            fgCredCalif.Text = "Credito"
        Case 2
            fgCredCalif.Text = "Fecha"
        Case 3
            fgCredCalif.Text = "Capital"
        Case 4
            fgCredCalif.Text = "Cal"
        Case 5
            fgCredCalif.Text = "Nota"
        Case 6
            fgCredCalif.Text = "Obs"
        Case 7
            fgCredCalif.Text = "Dias"
        Case 8
            fgCredCalif.Text = "Aut."
        Case 10
            fgCredCalif.Text = "Est."
    End Select
Next i
fgCredCalif.ColAlignment(0) = 0
fgCredCalif.ColAlignment(1) = 1
fgCredCalif.ColAlignment(2) = 1
fgCredCalif.ColAlignment(3) = 7
fgCredCalif.ColAlignment(4) = 4
fgCredCalif.ColAlignment(5) = 4
fgCredCalif.ColAlignment(6) = 4
fgCredCalif.ColAlignment(7) = 4
fgCredCalif.ColAlignment(8) = 4
fgCredCalif.ColAlignment(10) = 4
'
fgCredCalif.ColWidth(0) = 350
fgCredCalif.ColWidth(1) = 2000
fgCredCalif.ColWidth(2) = 1300
fgCredCalif.ColWidth(3) = 1000
fgCredCalif.ColWidth(4) = 600
fgCredCalif.ColWidth(5) = 0
fgCredCalif.ColWidth(6) = 0
fgCredCalif.ColWidth(7) = 500
fgCredCalif.ColWidth(8) = 0
fgCredCalif.ColWidth(9) = 0
fgCredCalif.ColWidth(10) = 0
End Sub

Private Sub DatosDetalleGrid(ByVal psCodPers As String, ByVal psCodCta As String, ByVal lsFecha As String, _
                            ByVal lsCal As String, ByVal lsObs As String, _
                            Optional ByVal lnCapital As Currency, Optional ByVal lsNota As String, _
                            Optional ByVal lnDiasAtraso As Integer)
Dim n As Integer
AdicionaRow fgCredCalif
'If lsEstado = "0" Then
'    BackColorFg fgCredCalif, &HC0C0&, False
'End If
n = fgCredCalif.row
fgCredCalif.TextMatrix(n, 1) = psCodCta
fgCredCalif.TextMatrix(n, 2) = lsFecha
fgCredCalif.TextMatrix(n, 3) = Format(lnCapital, "#0.00")
fgCredCalif.TextMatrix(n, 4) = lsCal
'fgCredCalif.TextMatrix(N, 5) = lsNota
fgCredCalif.TextMatrix(n, 6) = lsObs
fgCredCalif.TextMatrix(n, 7) = val(lnDiasAtraso)
'fgCredCalif.TextMatrix(N, 10) = Trim(lsEstado)
End Sub

Private Sub DatosCalifCredito(ByVal psCodPers As String, ByVal pnTipoEval As Integer)
Dim loDatos As COMNCredito.NCOMColocEval
Dim rs As ADODB.Recordset

Set loDatos = New COMNCredito.NCOMColocEval
    Set rs = loDatos.nObtieneCreditosEvaluadosPersDetalles(psCodPers, pnTipoEval)
Set loDatos = Nothing
CabeceraGrid
If rs Is Nothing Then Exit Sub
Do While Not rs.EOF
    Call DatosDetalleGrid(psCodPers, Trim(rs!Credito), Format(rs!Fecha, "dd/mm/yyyy"), rs!Cal, _
         rs!Obs, rs!nSaldoCap, , rs!nDiasAtraso)
        
    rs.MoveNext
Loop
rs.Close
End Sub
Private Sub CabGridCreditos()
Dim i As Integer

fgCreditos.cols = 24
fgCreditos.rows = 2
fgCreditos.Clear
For i = 0 To fgCreditos.cols - 1
    fgCreditos.row = 0
    fgCreditos.col = i
    fgCreditos.CellTextStyle = 0
    fgCreditos.CellAlignment = 4
    fgCreditos.CellFontBold = True
    Select Case i
        Case 0
            fgCreditos.ColWidth(i) = 400
        Case 1
            fgCreditos.Text = "Fecha"
            fgCreditos.ColWidth(i) = 1200
        Case 2
            fgCreditos.Text = "Crédito"
            fgCreditos.ColWidth(i) = 2000
        Case 3
            fgCreditos.Text = "Agencia"
            fgCreditos.ColWidth(i) = 1500
        Case 4
            fgCreditos.Text = "Producto"
            fgCreditos.ColWidth(i) = 2000
        Case 5
            fgCreditos.Text = "Estado"
            fgCreditos.ColWidth(i) = 2500
        Case 6
            fgCreditos.Text = "Rel"
            fgCreditos.ColWidth(i) = 1500
        Case 7
            fgCreditos.Text = "Analista"
            fgCreditos.ColWidth(i) = 800
        Case 8
            fgCreditos.Text = "Monto"
            fgCreditos.ColWidth(i) = 1000
        Case 9
            fgCreditos.Text = "Nota1"
            fgCreditos.ColWidth(i) = 900
        Case 10
            fgCreditos.Text = "Nota1 F/M"
            fgCreditos.ColWidth(i) = 1000
        Case 11
            fgCreditos.Text = "S.K Actual"
            fgCreditos.ColWidth(i) = 1200
        Case 12
            fgCreditos.Text = "S.K F/M"
            fgCreditos.ColWidth(i) = 1200
        Case 13
            fgCreditos.Text = "D/M Act"
            fgCreditos.ColWidth(i) = 900
        Case 14
            fgCreditos.Text = "D/M-F/M"
            fgCreditos.ColWidth(i) = 900
        Case 15
            fgCreditos.Text = "Cod.Ant"
            fgCreditos.ColWidth(i) = 1000
        Case 16
            fgCreditos.Text = "Moneda "
            fgCreditos.ColWidth(i) = 0
        Case 17
            fgCreditos.Text = "Ref "
            fgCreditos.ColWidth(i) = 800
        Case 18
            fgCreditos.Text = "Motivo Rechazo "
            fgCreditos.ColWidth(i) = 0
        Case 19
            fgCreditos.Text = "Abrev Cred "
            fgCreditos.ColWidth(i) = 0
        Case 20
            fgCreditos.Text = "Cancelac."
            fgCreditos.ColWidth(i) = 1200
        Case 21
            fgCreditos.Text = "Estado"
            fgCreditos.ColWidth(i) = 900
        Case 22
            fgCreditos.Text = "Fecha F/M"
            fgCreditos.ColWidth(i) = 1200
        Case 23
            fgCreditos.Text = "Monto Aprob"
            fgCreditos.ColWidth(i) = 1200
    End Select
Next i
fgCreditos.ColAlignment(0) = 0

End Sub
Private Sub CabGridGarantias()
Dim i As Integer
Dim lsCabecera As String
Dim lsAnchos As String
lsCabecera = " - Crédito -  Persona - Relacion - Tipo Garantia - Descripción - Documento - " _
            & " Moneda - Tasacion - Realizacion - Por Gravar - Estado - " _
            & "Gravado - Moneda - Nro. Garantia "

MSHFlex fgGarantias, 15, lsCabecera, " 500 - 1800 - 3500 - 1200 - 2500 - 3000 - 2800 - 900 - 1200 - 1200 - 1200 - 900 - 1200 - 1200 - 900 - 1200", _
        "L-L-L-L-L-L-L-R-R-R-L-R-R-L-C"

For i = 0 To fgGarantias.cols - 1
    fgGarantias.row = 0
    fgGarantias.col = i
    fgGarantias.CellTextStyle = 0
    fgGarantias.CellAlignment = 4
    fgGarantias.CellFontBold = True
Next
End Sub

Private Sub CabGridFuentes()
Dim i As Integer
Dim lsCabecera As String
Dim lsAnchos As String
lsCabecera = " - Tipo Fuente - Razon Social - Direccion - Zona - " _
            & " Sector - Actividad - Cargo - Fecha FD - Ingreso FD - Gastos FD - " _
            & "Fecha Bal - Ingresos Bal - Gasto Bal - Numfuente"

MSHFlex fgFuentes, 15, lsCabecera, " 300 - 1500 - 2500 - 2500 - 1500 - 1500 - 3500 - 2800 - 1200 - 1200 - 1200 - 1200 - 1200 - 1200 - 1200 ", _
        "L-L-L-L-L-L-L-L-L-R-R-L-R-R-C"

For i = 0 To fgFuentes.cols - 1
    fgFuentes.row = 0
    fgFuentes.col = i
    fgFuentes.CellTextStyle = 0
    fgFuentes.CellAlignment = 4
    fgFuentes.CellFontBold = True
Next
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
    Me.cmdNuevo.Enabled = True
    Me.cmdEditaCred.Enabled = True
'End If
End Sub


