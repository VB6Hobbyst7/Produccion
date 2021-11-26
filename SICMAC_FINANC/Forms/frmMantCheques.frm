VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMantCheques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Cheques"
   ClientHeight    =   6330
   ClientLeft      =   1515
   ClientTop       =   1815
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMantCheques.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   7230
      TabIndex        =   20
      Top             =   5925
      Width           =   1185
   End
   Begin TabDlg.SSTab stabCheques 
      Height          =   5805
      Left            =   15
      TabIndex        =   6
      Top             =   60
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   10239
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lista de Cheques"
      TabPicture(0)   =   "frmMantCheques.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos Generales"
      TabPicture(1)   =   "frmMantCheques.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAnular"
      Tab(1).Control(1)=   "cmdRechazar"
      Tab(1).Control(2)=   "cmdValorizar"
      Tab(1).Control(3)=   "cmdCambiar"
      Tab(1).Control(4)=   "FraIngCheque"
      Tab(1).Control(5)=   "txtGlosa"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Extornos"
      TabPicture(2)   =   "frmMantCheques.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtglosaExt"
      Tab(2).Control(1)=   "cmdExtornar"
      Tab(2).Control(2)=   "Frame3"
      Tab(2).ControlCount=   3
      Begin VB.TextBox txtglosaExt 
         Height          =   555
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Top             =   5070
         Width           =   6375
      End
      Begin VB.TextBox txtGlosa 
         Height          =   555
         Left            =   -74835
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   5040
         Width           =   5535
      End
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "&Extornar"
         Height          =   360
         Left            =   -68190
         TabIndex        =   27
         Top             =   5175
         Width           =   1275
      End
      Begin VB.Frame Frame3 
         Caption         =   "Lista de Extornos"
         Height          =   4485
         Left            =   -74850
         TabIndex        =   41
         Top             =   480
         Width           =   8100
         Begin VB.CommandButton cmdProcesarChq 
            Caption         =   "&Procesar"
            Height          =   345
            Left            =   6735
            TabIndex        =   24
            Top             =   210
            Width           =   1215
         End
         Begin VB.TextBox txtNumChequeExt 
            Height          =   330
            Left            =   4860
            TabIndex        =   23
            Top             =   255
            Width           =   1500
         End
         Begin Sicmact.FlexEdit fgExtornos 
            Height          =   3705
            Left            =   165
            TabIndex        =   25
            Top             =   690
            Width           =   7770
            _ExtentX        =   13705
            _ExtentY        =   6535
            Cols0           =   11
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "N°-Fecha-User-Cheque No-Banco-Estado-Moneda-Monto-Cuenta-cOpeCod-nMovnro"
            EncabezadosAnchos=   "350-900-600-1200-2300-800-700-900-1200-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-C-L-L-L-L-R-L-L-L"
            FormatosEdit    =   "0-0-0-0-0-0-0-2-0-0-0"
            TextArray0      =   "N°"
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin MSMask.MaskEdBox txtdesdeExt 
            Height          =   315
            Left            =   750
            TabIndex        =   21
            Top             =   240
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            _Version        =   393216
            ForeColor       =   -2147483635
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtHastaExt 
            Height          =   300
            Left            =   2415
            TabIndex        =   22
            Top             =   255
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            _Version        =   393216
            ForeColor       =   -2147483635
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   120
            TabIndex        =   54
            Top             =   270
            Width           =   570
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hasta :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   1830
            TabIndex        =   53
            Top             =   285
            Width           =   525
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Cheque N°:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   3915
            TabIndex        =   52
            Top             =   300
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5175
         Left            =   150
         TabIndex        =   39
         Top             =   450
         Width           =   8100
         Begin VB.TextBox txtNumCheque 
            Height          =   330
            Left            =   4845
            TabIndex        =   2
            Top             =   225
            Width           =   1500
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "&Buscar"
            Height          =   345
            Left            =   6795
            TabIndex        =   5
            Top             =   4725
            Width           =   1215
         End
         Begin VB.CommandButton cmdProcesar 
            Caption         =   "&Procesar"
            Height          =   345
            Left            =   6720
            TabIndex        =   3
            Top             =   180
            Width           =   1215
         End
         Begin MSDataGridLib.DataGrid fgCheques 
            Height          =   4005
            Left            =   105
            TabIndex        =   4
            Top             =   630
            Width           =   7905
            _ExtentX        =   13944
            _ExtentY        =   7064
            _Version        =   393216
            AllowUpdate     =   0   'False
            ColumnHeaders   =   -1  'True
            HeadLines       =   2
            RowHeight       =   15
            RowDividerStyle =   1
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "cNroDoc"
               Caption         =   "N° Cheque"
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
               DataField       =   "cPersNombre"
               Caption         =   "Institucion Financiera"
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
               DataField       =   "cCtaCod"
               Caption         =   "Cuenta"
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
               DataField       =   "nMonto"
               Caption         =   "Monto"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "cMoneda"
               Caption         =   "Moneda"
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
               DataField       =   "EstActual"
               Caption         =   "Estado"
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
            BeginProperty Column06 
               DataField       =   "dValorizaRef"
               Caption         =   "Valorizacion Referencia"
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
            BeginProperty Column07 
               DataField       =   "dValorizacion"
               Caption         =   "Valorizacion"
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
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2220.094
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1830.047
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   915.024
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   810.142
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column06 
                  Alignment       =   2
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column07 
                  Alignment       =   2
                  ColumnWidth     =   1154.835
               EndProperty
            EndProperty
         End
         Begin MSMask.MaskEdBox txtdesde 
            Height          =   315
            Left            =   735
            TabIndex        =   0
            Top             =   210
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            _Version        =   393216
            ForeColor       =   -2147483635
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txthasta 
            Height          =   300
            Left            =   2400
            TabIndex        =   1
            Top             =   225
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            _Version        =   393216
            ForeColor       =   -2147483635
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cheque N°:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   3900
            TabIndex        =   51
            Top             =   270
            Width           =   900
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   1815
            TabIndex        =   43
            Top             =   255
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   105
            TabIndex        =   42
            Top             =   240
            Width           =   570
         End
      End
      Begin VB.Frame FraIngCheque 
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
         Height          =   4485
         Left            =   -74895
         TabIndex        =   26
         Top             =   480
         Width           =   8040
         Begin VB.Frame fraCuentaAho 
            Caption         =   "Cuenta Ahorros"
            Enabled         =   0   'False
            Height          =   2490
            Left            =   3765
            TabIndex        =   44
            Top             =   1890
            Width           =   4035
            Begin Sicmact.FlexEdit fgCtaPers 
               Height          =   1635
               Left            =   105
               TabIndex        =   14
               Top             =   720
               Width           =   3870
               _ExtentX        =   6826
               _ExtentY        =   2884
               Cols0           =   3
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "N°-Persona-Rel."
               EncabezadosAnchos=   "350-2500-600"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X"
               ListaControles  =   "0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-C"
               FormatosEdit    =   "0-0-0"
               TextArray0      =   "N°"
               lbUltimaInstancia=   -1  'True
               ColWidth0       =   345
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin Sicmact.ActXCodCta txtCodCta 
               Height          =   390
               Left            =   195
               TabIndex        =   13
               Top             =   225
               Width           =   3630
               _ExtentX        =   6403
               _ExtentY        =   688
               Texto           =   "Cuenta N° :"
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Historia Estados :"
            Height          =   2115
            Left            =   75
            TabIndex        =   28
            Top             =   1890
            Width           =   3540
            Begin Sicmact.FlexEdit fgHistEst 
               Height          =   1800
               Left            =   90
               TabIndex        =   12
               Top             =   225
               Width           =   3360
               _ExtentX        =   5927
               _ExtentY        =   3175
               Cols0           =   3
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "N°-Fecha-Estado"
               EncabezadosAnchos=   "350-1000-1600"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X"
               TextStyleFixed  =   3
               ListaControles  =   "0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-L"
               FormatosEdit    =   "0-0-0"
               TextArray0      =   "N°"
               lbUltimaInstancia=   -1  'True
               lbFormatoCol    =   -1  'True
               lbPuntero       =   -1  'True
               ColWidth0       =   345
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
         End
         Begin VB.TextBox txtIngChqNumCheque 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   330
            Left            =   5970
            MaxLength       =   15
            TabIndex        =   7
            Top             =   195
            Width           =   1770
         End
         Begin MSMask.MaskEdBox txtIngChqFechaReg 
            Height          =   315
            Left            =   4410
            TabIndex        =   10
            Top             =   1448
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtIngChqFechaVal 
            Height          =   315
            Left            =   6690
            TabIndex        =   11
            Top             =   1455
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Frame fracajaGen 
            Enabled         =   0   'False
            Height          =   525
            Left            =   60
            TabIndex        =   40
            Top             =   1335
            Width           =   3000
            Begin VB.CheckBox chkConfirmar 
               Caption         =   "Por Confimar"
               Height          =   255
               Left            =   105
               TabIndex        =   8
               Top             =   210
               Width           =   1245
            End
            Begin VB.CheckBox chkDepositado 
               Caption         =   "Depositado"
               Height          =   285
               Left            =   1665
               TabIndex        =   9
               Top             =   195
               Width           =   1155
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "En Caja General"
               Height          =   210
               Left            =   105
               TabIndex        =   50
               Top             =   0
               Width           =   1155
            End
         End
         Begin VB.Label lblMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   1365
            TabIndex        =   49
            Top             =   4065
            Width           =   1500
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   300
            TabIndex        =   48
            Top             =   4080
            Width           =   750
         End
         Begin VB.Label lblMoneda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5745
            TabIndex        =   47
            Top             =   1035
            Width           =   1320
         End
         Begin VB.Label lblNroCtaIF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   915
            TabIndex        =   46
            Top             =   990
            Width           =   2070
         End
         Begin VB.Label lblPlaza 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3540
            TabIndex        =   45
            Top             =   1005
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Estado Actual:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   195
            Left            =   4980
            TabIndex        =   38
            Top             =   698
            Width           =   1260
         End
         Begin VB.Label lblEstado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Estado!!!!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   300
            Left            =   6315
            TabIndex        =   37
            Top             =   645
            Width           =   1470
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta N° :"
            Height          =   210
            Left            =   60
            TabIndex        =   36
            Top             =   1065
            Width           =   810
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Banco :"
            Height          =   210
            Left            =   90
            TabIndex        =   35
            Top             =   690
            Width           =   705
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "N° Cheque :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   4950
            TabIndex        =   34
            Top             =   248
            Width           =   975
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Ref. Valorizacion"
            Height          =   210
            Left            =   3120
            TabIndex        =   33
            Top             =   1500
            Width           =   1245
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Plaza"
            Height          =   210
            Left            =   3105
            TabIndex        =   32
            Top             =   1050
            Width           =   390
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Valorización:"
            Height          =   210
            Left            =   5625
            TabIndex        =   31
            Top             =   1500
            Width           =   945
         End
         Begin VB.Label lblIngChqDescIF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   930
            TabIndex        =   30
            Top             =   630
            Width           =   3915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   4995
            TabIndex        =   29
            Top             =   1080
            Width           =   630
         End
         Begin VB.Shape ShapeS 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   345
            Left            =   120
            Top             =   4035
            Width           =   2760
         End
      End
      Begin VB.CommandButton cmdCambiar 
         Caption         =   "&Cambiar"
         Height          =   345
         Left            =   -68085
         TabIndex        =   15
         Top             =   5145
         Width           =   1185
      End
      Begin VB.CommandButton cmdValorizar 
         Caption         =   "&Valorizar"
         Height          =   345
         Left            =   -68085
         TabIndex        =   17
         Top             =   5130
         Width           =   1185
      End
      Begin VB.CommandButton cmdRechazar 
         Caption         =   "&Rechazar"
         Height          =   345
         Left            =   -68085
         TabIndex        =   18
         Top             =   5130
         Width           =   1185
      End
      Begin VB.CommandButton cmdAnular 
         Caption         =   "&Anular"
         Height          =   345
         Left            =   -68085
         TabIndex        =   19
         Top             =   5130
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmMantCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Dim oDocRec As NDocRec

Sub CargaCheques()
Me.stabCheques.TabEnabled(1) = False
Me.cmdBuscar.Enabled = False
Set rs = New ADODB.Recordset
Set rs = oDocRec.GetCheques(txtdesde, txthasta, gsCodAge, txtNumCheque)
Set fgCheques.DataSource = rs
If Not rs.EOF And Not rs.BOF Then
    stabCheques.TabEnabled(1) = True
    Me.cmdBuscar.Enabled = True
Else
    MsgBox "Datos no Encontrados", vbInformation, "Aviso"
    txtdesde.SetFocus
End If

End Sub

Private Sub cmdAnular_Click()
Dim oCont As NContFunciones
Dim lsMovNro As String
Dim lsNroCheque As String
Set oCont = New NContFunciones
If Valida = False Then Exit Sub

If MsgBox("Desea realizar la anulacion del cheque Seleccionado??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsNroCheque = ""
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    oDocRec.GrabaCambioEstadoCheque gsFormatoFecha, lsMovNro, gsOpeCod, gsOpeDesc, Trim(txtGlosa), Trim(txtIngChqNumCheque), _
                        rs!cPersCod, rs!cIfTpo, CCur(lblMonto), gChqEstAnulado, rs!nConfCaja, rs!dValorizacion, Trim(txtCodCta.NroCuenta), gdFecSis
                        
    Dim oImp As NContImprimir
    Dim lsTexto As String
    Dim lbReimp As Boolean
    Set oImp = New NContImprimir
    
    lbReimp = True
    Do While lbReimp
        oImp.ImprimeBoletaGeneral gsOpeDescPadre, gsOpeDescHijo, gsOpeCod, CCur(lblMonto), _
                                 gsNomAge, lsMovNro, sLpt, , txtNumCheque, Trim(txtCodCta.NroCuenta), _
                                 txtGlosa
        If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            lbReimp = False
        End If
    Loop
    Set oImp = Nothing
    lsNroCheque = txtIngChqNumCheque
    CargaCheques
    rs.Find "cNroDoc='" & lsNroCheque & "'"
    txtGlosa = ""
End If

End Sub
Function Valida() As Boolean
Valida = True
If Len(Trim(txtIngChqNumCheque)) = "" Then
    MsgBox "Nro de Cheque no válido", vbInformation, "Aviso"
    Valida = False
    Exit Function
End If
Select Case gsOpeCod
    Case gOpeChequesAnulacion, gOpeChequesRechazo, gOpeChequesValorización
        If rs!nEstado <> gChqEstEnValorizacion Then
            MsgBox "Cheque se encuentra en Estado no permitido para realizar Operacion", vbInformation, "Aviso"
            Valida = False
            Exit Function
        End If
    Case gOpeChequesModFecVal
        If rs!nEstado <> gChqEstEnValorizacion Then
            MsgBox "Cheque se encuentra en Estado no permitido para realizar Operacion", vbInformation, "Aviso"
            Valida = False
            Exit Function
        End If
        If ValFecha(txtIngChqFechaVal) = False Then
            Valida = False
            Exit Function
        End If
        If CDate(txtIngChqFechaVal) < CDate(txtIngChqFechaReg) Then
            MsgBox "Fecha de Valorizacion no Puede ser menor que la Fecha Referencial", vbInformation, "Aviso"
            txtIngChqFechaVal.SetFocus
            Valida = False
            Exit Function
        End If
    Case gOpeChequesExtAnulación
    Case gOpeChequesExtRechazo
    Case gOpeChequesExtValorización
End Select
If Len(Trim(txtGlosa)) = 0 Then
    MsgBox "Glosa de Operación no Ingresada", vbInformation, "Aviso"
    txtGlosa.SetFocus
    Valida = False
    Exit Function
End If



End Function
Private Sub cmdBuscar_Click()
Dim oDesc As ClassDescObjeto
Set oDesc = New ClassDescObjeto
If Not rs Is Nothing Then
    If Not rs.EOF And Not rs.BOF Then
        oDesc.BuscarDato rs, 1, "Cheque N°"
    End If
End If
Set oDesc = Nothing
End Sub

Private Sub cmdCambiar_Click()
Dim oCont As NContFunciones
Dim lsMovNro As String
Dim lsNroCheque As String
Set oCont = New NContFunciones
If Valida = False Then Exit Sub

If MsgBox("Desea realizar el cambio de fecha de Valorización al cheque Seleccionado??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsNroCheque = ""
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    oDocRec.GrabaCambioEstadoCheque gsFormatoFecha, lsMovNro, gsOpeCod, gsOpeDesc, Trim(txtGlosa), Trim(txtIngChqNumCheque), _
                        rs!cPersCod, rs!cIfTpo, CCur(lblMonto), IIf(IsNull(rs!nEstado), 1, rs!nEstado), rs!nConfCaja, CDate(txtIngChqFechaVal), Trim(txtCodCta.NroCuenta), gdFecSis
                        
    Dim oImp As NContImprimir
    Dim lsTexto As String
    Dim lbReimp As Boolean
    Set oImp = New NContImprimir
    
    lbReimp = True
    Do While lbReimp
        oImp.ImprimeBoletaGeneral gsOpeDescPadre, gsOpeDescHijo, gsOpeCod, CCur(lblMonto), _
                                 gsNomAge, lsMovNro, sLpt, , txtNumCheque, Trim(txtCodCta.NroCuenta), _
                                 txtGlosa
        If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            lbReimp = False
        End If
    Loop
    Set oImp = Nothing
    lsNroCheque = txtIngChqNumCheque
    CargaCheques
    rs.Find "cNroDoc='" & lsNroCheque & "'"
    txtGlosa = ""
End If


End Sub

Private Sub cmdExtornar_Click()
Dim oCont As NContFunciones
Dim lnMovNroAnt As Long
Dim lsMovNro As String
Dim lsNumCheque As String
Dim lsPersCod As String
Dim lsPersNombre As String
Dim lsTpoIf As String
Dim lnMonto As Currency
Dim lsCtaAho As String

Set oCont = New NContFunciones
If fgExtornos.TextMatrix(1, 0) = "" Then
    MsgBox "No se Encuentran Operaciones para extornar", vbInformation, "Aviso"
    Exit Sub
End If
If Len(Trim(txtglosaExt)) = 0 Then
    MsgBox "Glosa de Operacion no ingresada", vbInformation, "Aviso"
    txtGlosa.SetFocus
    Exit Sub
End If
If MsgBox("Desea Extornar la Operación del cheque Seleccionado??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lnMovNroAnt = fgExtornos.TextMatrix(fgExtornos.Row, 10)
    lsNumCheque = fgExtornos.TextMatrix(fgExtornos.Row, 3)
    lsPersCod = fgExtornos.TextMatrix(fgExtornos.Row, 9)
    lsPersNombre = Trim(fgExtornos.TextMatrix(fgExtornos.Row, 4))
    lsTpoIf = fgExtornos.TextMatrix(fgExtornos.Row, 11)
    lnMonto = fgExtornos.TextMatrix(fgExtornos.Row, 7)
    lsCtaAho = fgExtornos.TextMatrix(fgExtornos.Row, 8)
    
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    oDocRec.GrabaExtornoCheque gsFormatoFecha, lnMovNroAnt, lsMovNro, gsOpeCod, gsOpeDesc, Trim(txtglosaExt), lsNumCheque, _
                        lsPersCod, lsTpoIf, lnMonto, gChqEstEnValorizacion, lsCtaAho, gdFecSis
                        
    Dim oImp As NContImprimir
    Dim lsTexto As String
    Dim lbReimp As Boolean
    Set oImp = New NContImprimir
    
    lbReimp = True
    Do While lbReimp
        oImp.ImprimeBoletaGeneral gsOpeDescPadre, gsOpeDescHijo, gsOpeCod, lnMonto, _
                                 gsNomAge, lsMovNro, sLpt, lsPersNombre, lsNumCheque, lsCtaAho, _
                                 txtglosaExt
        If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            lbReimp = False
        End If
    Loop
    Set oImp = Nothing
    cmdProcesarChq.Value = True
    txtglosaExt = ""
End If
End Sub

Private Sub cmdProcesar_Click()
If ValFecha(txtdesde) = False Then
    txtdesde.SetFocus
    Exit Sub
End If
If ValFecha(txthasta) = False Then
    txthasta.SetFocus
    Exit Sub
End If
If CDate(txtdesde) > CDate(txthasta) Then
    MsgBox "Fecha Inicial no pude ser mayor que fecha final", vbInformation, "Aviso"
    txtdesde.SetFocus
    Exit Sub
End If
CargaCheques
End Sub

Private Sub cmdProcesarChq_Click()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
If ValFecha(txtdesdeExt) = False Then
    txtdesde.SetFocus
    Exit Sub
End If
If ValFecha(txtHastaExt) = False Then
    txthasta.SetFocus
    Exit Sub
End If
If CDate(txtdesdeExt) > CDate(txtHastaExt) Then
    MsgBox "Fecha Inicial no pude ser mayor que fecha final", vbInformation, "Aviso"
    txtdesdeExt.SetFocus
    Exit Sub
End If
Dim lsOpeCod As String

Select Case gsOpeCod
Case gOpeChequesExtAnulación
    lsOpeCod = gOpeChequesAnulacion
Case gOpeChequesExtRechazo
    lsOpeCod = gOpeChequesRechazo
Case gOpeChequesExtValorización
    lsOpeCod = gOpeChequesValorización
End Select
fgExtornos.Clear
fgExtornos.FormaCabecera
fgExtornos.Rows = 2
Set rs = oDocRec.GetOpeCheques(lsOpeCod, txtdesdeExt, txtHastaExt, gsCodAge, Trim(txtNumChequeExt))
If Not rs.EOF And Not rs.BOF Then
    Set fgExtornos.Recordset = rs
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub cmdRechazar_Click()
Dim oCont As NContFunciones
Dim lsMovNro As String
Dim lsNroCheque As String
Set oCont = New NContFunciones
If Valida = False Then Exit Sub

If MsgBox("Desea realizar el Rechazo del cheque Seleccionado??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsNroCheque = ""
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    oDocRec.GrabaCambioEstadoCheque gsFormatoFecha, lsMovNro, gsOpeCod, gsOpeDesc, Trim(txtGlosa), Trim(txtIngChqNumCheque), _
                        rs!cPersCod, rs!cIfTpo, CCur(lblMonto), gChqEstRechazado, rs!nConfCaja, rs!dValorizacion, Trim(txtCodCta.NroCuenta), gdFecSis
                        
    Dim oImp As NContImprimir
    Dim lsTexto As String
    Dim lbReimp As Boolean
    Set oImp = New NContImprimir
    
    lbReimp = True
    Do While lbReimp
        oImp.ImprimeBoletaGeneral gsOpeDescPadre, gsOpeDescHijo, gsOpeCod, CCur(lblMonto), _
                                 gsNomAge, lsMovNro, sLpt, , txtNumCheque, Trim(txtCodCta.NroCuenta), _
                                 txtGlosa
        If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            lbReimp = False
        End If
    Loop
    Set oImp = Nothing
    lsNroCheque = txtIngChqNumCheque
    CargaCheques
    rs.Find "cNroDoc='" & lsNroCheque & "'"
    txtGlosa = ""
End If

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdValorizar_Click()
Dim oCont As NContFunciones
Dim lsMovNro As String
Dim lsNroCheque As String
Set oCont = New NContFunciones
If Valida = False Then Exit Sub

If MsgBox("Desea realizar la Valorización del cheque Seleccionado??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsNroCheque = ""
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    oDocRec.GrabaCambioEstadoCheque gsFormatoFecha, lsMovNro, gsOpeCod, gsOpeDesc, Trim(txtGlosa), Trim(txtIngChqNumCheque), _
                        rs!cPersCod, rs!cIfTpo, CCur(lblMonto), gChqEstValorizado, rs!nConfCaja, rs!dValorizacion, Trim(txtCodCta.NroCuenta), gdFecSis
                        
    Dim oImp As NContImprimir
    Dim lsTexto As String
    Dim lbReimp As Boolean
    Set oImp = New NContImprimir
    
    lbReimp = True
    Do While lbReimp
        oImp.ImprimeBoletaGeneral gsOpeDescPadre, gsOpeDescHijo, gsOpeCod, CCur(lblMonto), _
                                 gsNomAge, lsMovNro, sLpt, , txtNumCheque, Trim(txtCodCta.NroCuenta), _
                                 txtGlosa
        If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            lbReimp = False
        End If
    Loop
    Set oImp = Nothing
    lsNroCheque = txtIngChqNumCheque
    CargaCheques
    rs.Find "cNroDoc='" & lsNroCheque & "'"
    txtGlosa = ""
End If

End Sub

Private Sub fgCheques_GotFocus()
fgCheques.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub fgCheques_LostFocus()
'fgCheques.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub Form_Load()
CentraForm Me
Set oDocRec = New NDocRec
stabCheques.Tab = 0
stabCheques.TabEnabled(1) = False
stabCheques.TabVisible(2) = False
Me.cmdBuscar.Enabled = False
cmdCambiar.Visible = False
cmdAnular.Visible = False
cmdExtornar.Visible = False
cmdRechazar.Visible = False
cmdValorizar.Visible = False
txtGlosa.Visible = True
txtdesde = gdFecSis
txthasta = gdFecSis
txtdesdeExt = gdFecSis
txtHastaExt = gdFecSis
Select Case gsOpeCod
    Case gOpeChequesAnulacion
        cmdAnular.Visible = True
    Case gOpeChequesRechazo
        cmdRechazar.Visible = True
    Case gOpeChequesValorización
        cmdValorizar.Visible = True
    Case gOpeChequesModFecVal
        cmdCambiar.Visible = True
        txtIngChqFechaVal.Enabled = True
    Case gOpeChequesConsEstados
        txtGlosa.Visible = False
    Case gOpeChequesExtAnulación, gOpeChequesExtRechazo, gOpeChequesExtValorización
        stabCheques.TabVisible(0) = False
        stabCheques.TabVisible(1) = False
        stabCheques.TabVisible(2) = True
        stabCheques.Tab = 2
        cmdExtornar.Visible = True
End Select

End Sub

Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.ERROR, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
If Not pRecordset Is Nothing Then
    If Not pRecordset.EOF And Not pRecordset.BOF Then
        txtIngChqNumCheque = pRecordset!cNroDoc
        lblIngChqDescIF = pRecordset!cPersNombre
        lblMonto = Format(pRecordset!nMonto, "#,#0.00")
        txtIngChqFechaReg = pRecordset!dValorizaRef
        txtIngChqFechaVal = pRecordset!dValorizacion
        lblPlaza = pRecordset!cPlaza
        lblNroCtaIF = pRecordset!cIFCta
        chkDepositado = Val(pRecordset!cDepIF)
        chkConfirmar = pRecordset!nConfCaja
        lblEstado = pRecordset!EstActual
        lblMoneda = pRecordset!cMoneda
        fgHistEst.Clear
        fgHistEst.FormaCabecera
        fgHistEst.Rows = 2
        Set rs1 = oDocRec.GetEstadosCheques(pRecordset!cNroDoc, pRecordset!cPersCod, pRecordset!cIfTpo)
        If Not rs1.EOF And Not rs1.BOF Then
            Set fgHistEst.Recordset = rs1
        End If
        fraCuentaAho.Visible = True
        If pRecordset!cCtaCod <> "" Then
            txtCodCta.NroCuenta = pRecordset!cCtaCod
            Set rs1 = oDocRec.GetPersCuentaAho(pRecordset!cCtaCod)
        Else
            fraCuentaAho.Visible = False
        End If
    End If
End If

End Sub

Private Sub txtdesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txthasta.Enabled Then
        txthasta.SetFocus
    ElseIf cmdProcesar.Enabled Then
            cmdProcesar.SetFocus
        End If
End If
End Sub

Private Sub txtdesdeExt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtHastaExt.Enabled Then
        txthasta.SetFocus
    ElseIf cmdProcesar.Enabled Then
            cmdProcesarChq.SetFocus
        End If
End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdAnular.Visible Then cmdAnular.SetFocus
    If cmdCambiar.Visible Then cmdCambiar.SetFocus
    If cmdRechazar.Visible Then cmdRechazar.SetFocus
    If cmdValorizar.Visible Then cmdValorizar.SetFocus
End If
End Sub
Private Sub txtglosaExt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdExtornar.SetFocus
End If
End Sub

Private Sub txthasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNumCheque.SetFocus
End If

End Sub

Private Sub txtHastaExt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNumChequeExt.SetFocus
End If
End Sub

Private Sub txtIngChqFechaVal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGlosa.SetFocus
End If
End Sub

Private Sub txtNumCheque_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    If cmdProcesar.Enabled Then
        cmdProcesar.SetFocus
    ElseIf fgCheques.Enabled Then
        fgCheques.SetFocus
        End If
End If
End Sub
Private Sub txtNumChequeExt_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    If cmdProcesarChq.Enabled Then
        cmdProcesarChq.SetFocus
    ElseIf fgExtornos.Enabled Then
        fgExtornos.SetFocus
        End If
End If

End Sub
