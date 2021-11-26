VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Begin VB.Form frmEntidadesOpeRecipro 
   Caption         =   "Entidades con Operaciones Reciprocas"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   Icon            =   "frmEntidadesOpeRecipro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAbrevia 
      Height          =   375
      Left            =   7440
      TabIndex        =   52
      Top             =   3480
      Width           =   2775
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2805
      Left            =   120
      TabIndex        =   19
      Top             =   4080
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4948
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Disponibles - 1"
      TabPicture(0)   =   "frmEntidadesOpeRecipro.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdDisBorrar"
      Tab(0).Control(1)=   "cmdDisAgregar"
      Tab(0).Control(2)=   "cmdDisAh"
      Tab(0).Control(3)=   "Frame8"
      Tab(0).Control(4)=   "grdDisponibles"
      Tab(0).Control(5)=   "lblDisAho"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Depósitos - 2"
      TabPicture(1)   =   "frmEntidadesOpeRecipro.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdDepBorrar"
      Tab(1).Control(1)=   "cmdDepAgregar"
      Tab(1).Control(2)=   "cmdDepAh"
      Tab(1).Control(3)=   "Frame9"
      Tab(1).Control(4)=   "grdDepositos"
      Tab(1).Control(5)=   "lblDepAh"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Ingresos - 3"
      TabPicture(2)   =   "frmEntidadesOpeRecipro.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdIngBorrar"
      Tab(2).Control(1)=   "cmdIngAgregar"
      Tab(2).Control(2)=   "cmdIngAh"
      Tab(2).Control(3)=   "Frame10"
      Tab(2).Control(4)=   "grdIngresos"
      Tab(2).Control(5)=   "lblIngAh"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Gastos - 4"
      TabPicture(3)   =   "frmEntidadesOpeRecipro.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblGasAh"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "grdGastos"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame11"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdGasAh"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdGasAgregar"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdGasBorrar"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      Begin VB.CommandButton cmdGasBorrar 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   6240
         TabIndex        =   48
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdGasAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   6240
         TabIndex        =   47
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdIngBorrar 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   -68760
         TabIndex        =   46
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdIngAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   -68760
         TabIndex        =   45
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdDisBorrar 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   -68760
         TabIndex        =   44
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdDisAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   -68760
         TabIndex        =   43
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdDepBorrar 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   -68760
         TabIndex        =   42
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdDepAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   -68760
         TabIndex        =   41
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdGasAh 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5280
         TabIndex        =   40
         Top             =   960
         Width           =   375
      End
      Begin VB.Frame Frame11 
         Caption         =   "Tipo Cuenta"
         Height          =   615
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   3255
         Begin VB.OptionButton optGasPF 
            Caption         =   "Cta. Plazo Fijo-2"
            Height          =   195
            Left            =   1560
            TabIndex        =   38
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optGasAh 
            Caption         =   "Cta. Ahorro-1"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdIngAh 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69720
         TabIndex        =   35
         Top             =   960
         Width           =   375
      End
      Begin VB.Frame Frame10 
         Caption         =   "Tipo Cuenta"
         Height          =   615
         Left            =   -74880
         TabIndex        =   31
         Top             =   720
         Width           =   3255
         Begin VB.OptionButton optIngPF 
            Caption         =   "Cta. Plazo Fijo-2"
            Height          =   195
            Left            =   1560
            TabIndex        =   33
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optIngAh 
            Caption         =   "Cta. Ahorro-1"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdDepAh 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69720
         TabIndex        =   30
         Top             =   960
         Width           =   375
      End
      Begin VB.Frame Frame9 
         Caption         =   "Tipo Cuenta"
         Height          =   615
         Left            =   -74880
         TabIndex        =   26
         Top             =   720
         Width           =   3255
         Begin VB.OptionButton optDepPF 
            Caption         =   "Cta. Plazo Fijo-2"
            Height          =   195
            Left            =   1560
            TabIndex        =   28
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optDepAh 
            Caption         =   "Cta. Ahorro-1"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdDisAh 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69720
         TabIndex        =   24
         Top             =   960
         Width           =   375
      End
      Begin VB.Frame Frame8 
         Caption         =   "Tipo Cuenta"
         Height          =   615
         Left            =   -74880
         TabIndex        =   20
         Top             =   720
         Width           =   3255
         Begin VB.OptionButton optDisAh 
            Caption         =   "Cta. Ahorro-1"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optDisPF 
            Caption         =   "Cta. Plazo Fijo-2"
            Height          =   195
            Left            =   1560
            TabIndex        =   21
            Top             =   240
            Width           =   1575
         End
      End
      Begin MSDataGridLib.DataGrid grdDisponibles 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   25
         Top             =   1440
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2355
         _Version        =   393216
         BackColor       =   -2147483629
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cTipAho"
            Caption         =   "Tipo Ahorro"
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
            DataField       =   "cCtaContCod"
            Caption         =   "Cta.Cont."
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
            DataField       =   "cTipCta"
            Caption         =   "cTipCta"
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
            DataField       =   "cPersCod"
            Caption         =   "cPersCod"
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
            DataField       =   "cMovNro"
            Caption         =   "cMovNro"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grdDepositos 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   49
         Top             =   1440
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2355
         _Version        =   393216
         BackColor       =   -2147483629
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cTipAho"
            Caption         =   "Tipo Ahorro"
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
            DataField       =   "cCtaContCod"
            Caption         =   "Cta.Cont."
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
            DataField       =   "cTipCta"
            Caption         =   "cTipCta"
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
            DataField       =   "cPersCod"
            Caption         =   "cPersCod"
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
            DataField       =   "cMovNro"
            Caption         =   "cMovNro"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grdIngresos 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   50
         Top             =   1440
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2355
         _Version        =   393216
         BackColor       =   -2147483629
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cTipAho"
            Caption         =   "Tipo Ahorro"
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
            DataField       =   "cCtaContCod"
            Caption         =   "Cta.Cont."
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
            DataField       =   "cTipCta"
            Caption         =   "cTipCta"
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
            DataField       =   "cPersCod"
            Caption         =   "cPersCod"
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
            DataField       =   "cMovNro"
            Caption         =   "cMovNro"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grdGastos 
         Height          =   1335
         Left            =   120
         TabIndex        =   51
         Top             =   1440
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2355
         _Version        =   393216
         BackColor       =   -2147483629
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cTipAho"
            Caption         =   "Tipo Ahorro"
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
            DataField       =   "cCtaContCod"
            Caption         =   "Cta.Cont."
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
            DataField       =   "cTipCta"
            Caption         =   "cTipCta"
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
            DataField       =   "cPersCod"
            Caption         =   "cPersCod"
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
            DataField       =   "cMovNro"
            Caption         =   "cMovNro"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label lblGasAh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3600
         TabIndex        =   39
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblIngAh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -71400
         TabIndex        =   34
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblDepAh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -71400
         TabIndex        =   29
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblDisAho 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -71400
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2775
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   10095
      Begin MSDataGridLib.DataGrid grdEntiOpeRecipro 
         Height          =   2535
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4471
         _Version        =   393216
         BackColor       =   -2147483629
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cPersIDNro"
            Caption         =   "Cod. Pers."
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
            Caption         =   "Nombre"
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
            DataField       =   "cCodEnti"
            Caption         =   "Cod. Ent."
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
            DataField       =   "cConsDescripcion"
            Caption         =   "Sector"
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
            DataField       =   "cEstado"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4995.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1005.165
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Activo"
      Height          =   195
      Left            =   8520
      TabIndex        =   16
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   6840
      Width           =   8775
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6240
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   7800
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4920
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   2760
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sectores"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   4575
      Begin VB.OptionButton Option3 
         Caption         =   "Público Ahorros"
         Height          =   255
         Left            =   3000
         TabIndex        =   54
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Público Gastos"
         Height          =   195
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Financiero"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdBuscaAseg 
      Caption         =   "..."
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
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   3000
      Width           =   390
   End
   Begin VB.TextBox txtCodEntidad 
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Abrevia:"
      Height          =   195
      Left            =   6840
      TabIndex        =   53
      Top             =   3600
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nombre Entidad"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1140
   End
   Begin VB.Label LblAsegPersNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   2925
      TabIndex        =   4
      Top             =   3000
      Width           =   4845
   End
   Begin VB.Label LblAsegPersCod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   1560
      TabIndex        =   3
      Top             =   3000
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cód. Entidad:"
      Height          =   195
      Left            =   4680
      TabIndex        =   1
      Top             =   3600
      Width           =   960
   End
End
Attribute VB_Name = "frmEntidadesOpeRecipro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lConsulta As Boolean
Dim sSql As String
Dim rsConst As ADODB.Recordset
Dim rsEnti As ADODB.Recordset

Dim rsCtaContDisEnti As ADODB.Recordset
Dim rsCtaContDepEnti As ADODB.Recordset
Dim rsCtaContIngEnti As ADODB.Recordset
Dim rsCtaContEgrEnti As ADODB.Recordset

Dim i As Integer
Dim lNuevo As Boolean
Dim oConst As NConstSistemas

Private Sub cmdAceptar_Click()
Dim sqlCod As String
Dim Cod As Integer
On Error GoTo AceptarErr

If Not ValidaDatos Then
   Exit Sub
End If

If MsgBox(" ¿ Seguro de grabar datos ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   Select Case lNuevo
      Case True
            If oConst.BuscaEntiOpeRecipro(LblAsegPersCod, IIf(Me.Option1.value = True, 1, IIf(Me.Option2.value = True, 2, 3))) Then
                MsgBox "Esta persona ya se encuentra registrada con el mismo sector, verifique.", vbInformation, "Atención"
                Exit Sub
            End If
            oConst.InsertaEntiOpeRecipro Me.LblAsegPersCod, Me.txtCodEntidad, IIf(Me.Option1.value = True, 1, IIf(Me.Option2.value = True, 2, 3)), IIf(Me.Check1.value = 1, 1, 0), Trim(txtAbrevia.Text), gsMovNro  '', Me.lblDisAho.Caption, Me.lblDisPF.Caption, Me.lblDepAh.Caption, Me.lblDepPF.Caption, Me.lblIngAh.Caption, Me.lblIngPF.Caption, Me.lblGasAh.Caption, Me.lblGasPF.Caption, gsMovNro
      Case False
            oConst.ActualizaEntiOpeRecipro Me.LblAsegPersCod, Me.txtCodEntidad, IIf(Me.Option1.value = True, 1, IIf(Me.Option2.value = True, 2, 3)), IIf(Me.Check1.value = 1, 1, 0), Trim(txtAbrevia.Text), gsMovNro '', Me.lblDisAho.Caption, Me.lblDisPF.Caption, Me.lblDepAh.Caption, Me.lblDepPF.Caption, Me.lblIngAh.Caption, Me.lblIngPF.Caption, Me.lblGasAh.Caption, Me.lblGasPF.Caption, gsMovNro
   End Select
    CargaEntiOpeReciprocas
End If
LblAsegPersCod.Caption = ""
LblAsegPersNombre.Caption = ""
txtCodEntidad.Text = ""
txtAbrevia.Text = ""
Me.lblDisAho.Caption = ""
Me.lblDepAh.Caption = ""
Me.lblIngAh.Caption = ""
Me.lblGasAh.Caption = ""
CargaCtasContEntiFinan ("X")
ActivaBotones True
grdEntiOpeRecipro.SetFocus
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdBuscaAseg_Click()

    Dim oPers As UPersona
    Set oPers = frmBuscaPersona.Inicio
    If Not oPers Is Nothing Then
        LblAsegPersCod.Caption = oPers.sPersCod
        LblAsegPersNombre.Caption = oPers.sPersNombre
     
    End If
    Set oPers = Nothing
    cmdBuscaAseg.SetFocus

End Sub

Private Sub cmdCancelar_Click()

If lNuevo And Me.Option1.value Then
   If Not rsCtaContDisEnti.EOF Or Not rsCtaContDepEnti.EOF Or Not rsCtaContIngEnti.EOF Or Not rsCtaContEgrEnti.EOF Then
        MsgBox "Si va cancelar la institución primero elimine las cuentas creadas.", vbInformation, "Atención"
        Exit Sub
   End If
End If

LblAsegPersCod.Caption = ""
LblAsegPersNombre.Caption = ""
txtCodEntidad.Text = ""
txtAbrevia.Text = ""
Me.lblDisAho.Caption = ""
Me.lblDepAh.Caption = ""
Me.lblIngAh.Caption = ""
Me.lblGasAh.Caption = ""

CargaCtasContEntiFinan ("X")

ActivaBotones True
Me.grdEntiOpeRecipro.SetFocus
End Sub

Private Sub cmdDepAgregar_Click()
Dim sqlCod As String
Dim Cod As Integer
On Error GoTo AceptarErr

If (Me.optDepAh.value = False And Me.optDepPF.value = False) Then
    MsgBox "Seleccione un tipo de cuenta de ahorro.", vbInformation, "Atención"
    Exit Sub
End If
If Me.lblDepAh.Caption = "" Then
    MsgBox "Seleccione una cuenta contable.", vbInformation, "Atención"
    Exit Sub
End If

If oConst.BuscaCtasContEntiFinanOpeRecipro(IIf(Me.optDepAh.value = True, "1", "2"), Me.lblDepAh.Caption, "2", LblAsegPersCod) Then
    MsgBox "Esta cuenta contable ya fue ingresada.", vbInformation, "Atención"
    Exit Sub
End If

oConst.InsertaCtasContEntiFinanOpeRecipro IIf(Me.optDepAh.value = True, "1", "2"), Me.lblDepAh.Caption, "2", LblAsegPersCod
CargaCtasContEntiFinan

Me.lblDepAh.Caption = ""
'Me.optDepAh.value = False
'Me.optDepPF.value = False

Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"

End Sub

Private Sub cmdDepBorrar_Click()
Dim sELIM As String
Dim lrs   As ADODB.Recordset
    If rsCtaContDepEnti.EOF Then
       MsgBox "No existen datos para Borrar", vbInformation, "¡Aviso!"
       Exit Sub
    End If

    If MsgBox(" ¿ Está seguro de Eliminar este dato ? ", vbQuestion + vbYesNo, "Mensaje de confirmación") = vbYes Then
       oConst.EliminaCtasContEntiFinanOpeRecipro rsCtaContDepEnti!cTipAho, rsCtaContDepEnti!cCtaContCod, rsCtaContDepEnti!cTipCta, rsCtaContDepEnti!cPersCod
       CargaCtasContEntiFinan
'       Set oConst = Nothing
       RSClose lrs
    End If

End Sub

Private Sub cmdDisAgregar_Click()
Dim sqlCod As String
Dim Cod As Integer
On Error GoTo AceptarErr

If (Me.optDisAh.value = False And Me.optDisPF.value = False) Then
    MsgBox "Seleccione un tipo de cuenta de ahorro.", vbInformation, "Atención"
    Exit Sub
End If
If Me.lblDisAho.Caption = "" Then
    MsgBox "Seleccione una cuenta contable.", vbInformation, "Atención"
    Exit Sub
End If

If oConst.BuscaCtasContEntiFinanOpeRecipro(IIf(Me.optDisAh.value = True, "1", "2"), Me.lblDisAho.Caption, "1", LblAsegPersCod) Then
    MsgBox "Esta cuenta contable ya fue ingresada.", vbInformation, "Atención"
    Exit Sub
End If

oConst.InsertaCtasContEntiFinanOpeRecipro IIf(Me.optDisAh.value = True, "1", "2"), Me.lblDisAho.Caption, "1", LblAsegPersCod
CargaCtasContEntiFinan

Me.lblDisAho.Caption = ""
'Me.optDisAh.value = False
'Me.optDisPF.value = False

Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdDisBorrar_Click()
Dim sELIM As String
Dim lrs   As ADODB.Recordset
    If rsCtaContDisEnti.EOF Then
       MsgBox "No existen datos para Borrar", vbInformation, "¡Aviso!"
       Exit Sub
    End If

    If MsgBox(" ¿ Está seguro de Eliminar este dato ? ", vbQuestion + vbYesNo, "Mensaje de confirmación") = vbYes Then
       oConst.EliminaCtasContEntiFinanOpeRecipro rsCtaContDisEnti!cTipAho, rsCtaContDisEnti!cCtaContCod, rsCtaContDisEnti!cTipCta, rsCtaContDisEnti!cPersCod
       CargaCtasContEntiFinan
       'Set oConst = Nothing
       RSClose lrs
    End If
End Sub

Private Sub cmdEliminar_Click()
Dim sELIM As String
Dim lrs   As ADODB.Recordset
If rsEnti.EOF Then
   MsgBox "No existen datos para Eliminar", vbInformation, "¡Aviso!"
   Exit Sub
End If
If MsgBox(" ¿ Está seguro de Eliminar este dato ? ", vbQuestion + vbYesNo, "Mensaje de confirmación") = vbYes Then
   oConst.EliminaEntiOpeRecipro rsEnti!cpersidnro, rsEnti!nSector
   CargaEntiOpeReciprocas
   Set oDoc = Nothing
   RSClose lrs
End If

CargaCtasContEntiFinan ("X")

grdEntiOpeRecipro.SetFocus
End Sub

Private Sub cmdGasAgregar_Click()
Dim sqlCod As String
Dim Cod As Integer
On Error GoTo AceptarErr

If (Me.optGasAh.value = False And Me.optGasPF.value = False) Then
    MsgBox "Seleccione un tipo de cuenta de ahorro.", vbInformation, "Atención"
    Exit Sub
End If
If Me.lblGasAh.Caption = "" Then
    MsgBox "Seleccione una cuenta contable.", vbInformation, "Atención"
    Exit Sub
End If

If oConst.BuscaCtasContEntiFinanOpeRecipro(IIf(Me.optGasAh.value = True, "1", "2"), Me.lblGasAh.Caption, "4", LblAsegPersCod) Then
    MsgBox "Esta cuenta contable ya fue ingresada.", vbInformation, "Atención"
    Exit Sub
End If

oConst.InsertaCtasContEntiFinanOpeRecipro IIf(Me.optGasAh.value = True, "1", "2"), Me.lblGasAh.Caption, "4", LblAsegPersCod
CargaCtasContEntiFinan

Me.lblGasAh.Caption = ""
'Me.optGasAh.value = False
'Me.optGasPF.value = False

Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"

End Sub

Private Sub cmdGasBorrar_Click()
Dim sELIM As String
Dim lrs   As ADODB.Recordset
    If rsCtaContEgrEnti.EOF Then
       MsgBox "No existen datos para Borrar", vbInformation, "¡Aviso!"
       Exit Sub
    End If

    If MsgBox(" ¿ Está seguro de Eliminar este dato ? ", vbQuestion + vbYesNo, "Mensaje de confirmación") = vbYes Then
       oConst.EliminaCtasContEntiFinanOpeRecipro rsCtaContEgrEnti!cTipAho, rsCtaContEgrEnti!cCtaContCod, rsCtaContEgrEnti!cTipCta, rsCtaContEgrEnti!cPersCod
       CargaCtasContEntiFinan
'       Set oConst = Nothing
       RSClose lrs
    End If

End Sub

Private Sub cmdIngAgregar_Click()
Dim sqlCod As String
Dim Cod As Integer
On Error GoTo AceptarErr

If (Me.optIngAh.value = False And Me.optIngPF.value = False) Then
    MsgBox "Seleccione un tipo de cuenta de ahorro.", vbInformation, "Atención"
    Exit Sub
End If
If Me.lblIngAh.Caption = "" Then
    MsgBox "Seleccione una cuenta contable.", vbInformation, "Atención"
    Exit Sub
End If

If oConst.BuscaCtasContEntiFinanOpeRecipro(IIf(Me.optIngAh.value = True, "1", "2"), Me.lblIngAh.Caption, "3", LblAsegPersCod) Then
    MsgBox "Esta cuenta contable ya fue ingresada.", vbInformation, "Atención"
    Exit Sub
End If

oConst.InsertaCtasContEntiFinanOpeRecipro IIf(Me.optIngAh.value = True, "1", "2"), Me.lblIngAh.Caption, "3", LblAsegPersCod
CargaCtasContEntiFinan

Me.lblIngAh.Caption = ""
'Me.optIngAh.value = False
'Me.optIngPF.value = False

Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"

End Sub

Private Sub cmdIngAh_Click()
    Dim rsObj As ADODB.Recordset
    Dim oCta As DCtaCont
    Dim oDescObj As New ClassDescObjeto
    Dim m_lbUltimaInstancia As Boolean
    Set oCta = New DCtaCont
    
    'Set rsObj = oCta.CargaCtaCont("cCtaContCod LIKE '510%' and nCtaEstado=1", , adLockReadOnly, True)
    Set rsObj = oCta.CargaCtaContOpeRecipro(adLockReadOnly, True, "3", "")
    Set oCta = Nothing

    'cCtaContCod LIKE '510%' and nCtaEstado=1
    
    oDescObj.lbUltNivel = m_lbUltimaInstancia
    oDescObj.ShowGrid rsObj, "Cuentas"
    rsObj.Close

    If oDescObj.lbOk Then
        Me.lblIngAh.Caption = Trim(oDescObj.gsSelecCod)
    Else
        Me.lblIngAh.Caption = ""
        Exit Sub
    End If

End Sub

Private Sub cmdIngBorrar_Click()
Dim sELIM As String
Dim lrs   As ADODB.Recordset
    If rsCtaContIngEnti.EOF Then
       MsgBox "No existen datos para Borrar", vbInformation, "¡Aviso!"
       Exit Sub
    End If

    If MsgBox(" ¿ Está seguro de Eliminar este dato ? ", vbQuestion + vbYesNo, "Mensaje de confirmación") = vbYes Then
       oConst.EliminaCtasContEntiFinanOpeRecipro rsCtaContIngEnti!cTipAho, rsCtaContIngEnti!cCtaContCod, rsCtaContIngEnti!cTipCta, rsCtaContIngEnti!cPersCod
       CargaCtasContEntiFinan
       'Set oConst = Nothing
       RSClose lrs
    End If
End Sub

Private Sub CmdModificar_Click()
If rsEnti.EOF Then
   MsgBox "No existen datos para Modificar", vbInformation, "Aviso"
   Exit Sub
End If

If rsEnti!nSector = 1 Then
    Me.Option1.value = True
    Me.Option2.value = False
    Me.Option3.value = False
ElseIf rsEnti!nSector = 2 Then
    Me.Option1.value = False
    Me.Option2.value = True
    Me.Option3.value = False
Else
    Me.Option1.value = False
    Me.Option2.value = False
    Me.Option3.value = True
End If
txtCodEntidad.Text = rsEnti!cCodEnti
txtAbrevia.Text = IIf(IsNull(rsEnti!cAbrevia), "", rsEnti!cAbrevia)
If rsEnti!nestado = True Then
    Check1.value = 1
Else
    Check1.value = 0
End If
Me.LblAsegPersCod = rsEnti!cpersidnro
Me.LblAsegPersNombre = rsEnti!cPersNombre

CargaCtasContEntiFinan

ActivaBotones False

lNuevo = False
End Sub

Private Sub cmdNuevo_Click()

LblAsegPersCod.Caption = ""
LblAsegPersNombre.Caption = ""
txtCodEntidad.Text = ""
txtAbrevia.Text = ""
ActivaBotones False
lNuevo = True

CargaCtasContEntiFinan ("X")

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdDisAh_Click()
    Dim rsObj As ADODB.Recordset
    Dim oCta As DCtaCont
    Dim oDescObj As New ClassDescObjeto
    Dim m_lbUltimaInstancia As Boolean
    Set oCta = New DCtaCont
        
    Set rsObj = oCta.CargaCtaContOpeRecipro(adLockReadOnly, True, "1", "")
    Set oCta = Nothing

    oDescObj.lbUltNivel = m_lbUltimaInstancia
    oDescObj.ShowGrid rsObj, "Cuentas"
    rsObj.Close

    If oDescObj.lbOk Then
        Me.lblDisAho.Caption = Trim(oDescObj.gsSelecCod)
    Else
        Me.lblDisAho.Caption = ""
        Exit Sub
    End If
End Sub

Private Sub cmdDepAh_Click()
    Dim rsObj As ADODB.Recordset
    Dim oCta As DCtaCont
    Dim oDescObj As New ClassDescObjeto
    Dim m_lbUltimaInstancia As Boolean
    Set oCta = New DCtaCont
    
'    If Me.optDepAh.value Then
'        Set rsObj = oCta.CargaCtaCont("cCtaContCod LIKE '2_[12]2%' and nCtaEstado=1", , adLockReadOnly, True)
'        Set oCta = Nothing
'    Else
'        Set rsObj = oCta.CargaCtaCont("cCtaContCod LIKE '2_[12]3%' and nCtaEstado=1", , adLockReadOnly, True)
'        Set oCta = Nothing
'    End If

    Set rsObj = oCta.CargaCtaContOpeRecipro(adLockReadOnly, True, "2", IIf(Me.optDepAh.value, "1", "2"))
    Set oCta = Nothing

    oDescObj.lbUltNivel = m_lbUltimaInstancia
    oDescObj.ShowGrid rsObj, "Cuentas"
    rsObj.Close

    If oDescObj.lbOk Then
        Me.lblDepAh.Caption = Trim(oDescObj.gsSelecCod)
    Else
        Me.lblDepAh.Caption = ""
        Exit Sub
    End If
End Sub


Private Sub cmdGasAh_Click()
    Dim rsObj As ADODB.Recordset
    Dim oCta As DCtaCont
    Dim oDescObj As New ClassDescObjeto
    Dim m_lbUltimaInstancia As Boolean
    Set oCta = New DCtaCont
    
    'Set rsObj = oCta.CargaCtaCont("cCtaContCod LIKE '410%' and nCtaEstado=1", , adLockReadOnly, True)
    Set rsObj = oCta.CargaCtaContOpeRecipro(adLockReadOnly, True, "4", "")
    Set oCta = Nothing
    'cCtaContCod LIKE '410%' and nCtaEstado=1
    
    oDescObj.lbUltNivel = m_lbUltimaInstancia
    oDescObj.ShowGrid rsObj, "Cuentas"
    rsObj.Close

    If oDescObj.lbOk Then
        Me.lblGasAh.Caption = Trim(oDescObj.gsSelecCod)
    Else
        Me.lblGasAh.Caption = ""
        Exit Sub
    End If

End Sub

Private Sub Form_Load()
    Set oConst = New NConstSistemas
       
    CargaEntiOpeReciprocas
    
    If lConsulta Then
        cmdNuevo.Visible = False
        cmdModificar.Visible = False
        cmdEliminar.Visible = False
    End If
    
    Me.optDepAh.value = True
    Me.optDisAh.value = True
    Me.optGasAh.value = True
    Me.optIngAh.value = True
    
    ActivaBotones True
    
    CentraForm Me
End Sub
Public Sub Inicio(plConsulta As Boolean)
lConsulta = plConsulta
Me.Show 1
End Sub
Sub ActivaBotones(lActiva As Boolean)

cmdNuevo.Visible = lActiva
cmdModificar.Visible = lActiva
cmdEliminar.Visible = lActiva
cmdSalir.Visible = lActiva
cmdAceptar.Visible = Not lActiva
cmdCancelar.Visible = Not lActiva

cmdBuscaAseg.Enabled = Not lActiva
Frame1.Enabled = Not lActiva
txtCodEntidad.Enabled = Not lActiva
txtAbrevia.Enabled = Not lActiva
Check1.Enabled = Not lActiva

Me.cmdDisAh.Enabled = Not lActiva
Me.cmdDepAh.Enabled = Not lActiva
Me.cmdIngAh.Enabled = Not lActiva
Me.cmdGasAh.Enabled = Not lActiva

SSTab1.Enabled = Not lActiva

End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsEnti
Set rsEnti = Nothing
End Sub

Private Sub grdDepositos_GotFocus()
grdDepositos.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdDepositos_LostFocus()
grdDepositos.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub grdDisponibles_GotFocus()
grdEntiOpeRecipro.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdDisponibles_LostFocus()
grdDisponibles.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub grdEntiOpeRecipro_GotFocus()
grdDisponibles.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdEntiOpeRecipro_LostFocus()
grdEntiOpeRecipro.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub grdEntiOpeRecipro_HeadClick(ByVal ColIndex As Integer)
If Not rsEnti Is Nothing Then
   If Not rsEnti.EOF Then
      rsEnti.Sort = grdEntiOpeRecipro.Columns(ColIndex).DataField
   End If
End If
End Sub

Private Sub CargaEntiOpeReciprocas()
    Set rsEnti = oConst.LeeEntidadesOpeReciprocas()
    Set Me.grdEntiOpeRecipro.DataSource = rsEnti
End Sub

Private Sub CargaCtasContEntiFinan(Optional ByVal psSinDatos As String = "")

If psSinDatos = "" Then
    Set rsCtaContDisEnti = oConst.LeeCuentasContEntiFinan("1", LblAsegPersCod)
    Set Me.grdDisponibles.DataSource = rsCtaContDisEnti
    
    Set rsCtaContDepEnti = oConst.LeeCuentasContEntiFinan("2", LblAsegPersCod)
    Set Me.grdDepositos.DataSource = rsCtaContDepEnti
    
    Set rsCtaContIngEnti = oConst.LeeCuentasContEntiFinan("3", LblAsegPersCod)
    Set Me.grdIngresos.DataSource = rsCtaContIngEnti
    
    Set rsCtaContEgrEnti = oConst.LeeCuentasContEntiFinan("4", LblAsegPersCod)
    Set Me.grdGastos.DataSource = rsCtaContEgrEnti
Else
    Set rsCtaContDisEnti = oConst.LeeCuentasContEntiFinan("1", "123456789")
    Set Me.grdDisponibles.DataSource = rsCtaContDisEnti
    
    Set rsCtaContDepEnti = oConst.LeeCuentasContEntiFinan("2", "123456789")
    Set Me.grdDepositos.DataSource = rsCtaContDepEnti
    
    Set rsCtaContIngEnti = oConst.LeeCuentasContEntiFinan("3", "123456789")
    Set Me.grdIngresos.DataSource = rsCtaContIngEnti
    
    Set rsCtaContEgrEnti = oConst.LeeCuentasContEntiFinan("4", "123456789")
    Set Me.grdGastos.DataSource = rsCtaContEgrEnti
End If
    
End Sub


Private Function ValidaDatos() As Boolean
ValidaDatos = False

If Len(Trim(LblAsegPersCod.Caption)) = 0 Or Len(Trim(LblAsegPersNombre.Caption)) = 0 Then
   MsgBox "Seleccione una entidad.", vbCritical, "Atención"
   Exit Function
End If

If Len(Trim(txtCodEntidad.Text)) = 0 Then
   MsgBox "Ingrese el código de Entidad.", vbCritical, "Atención"
   Exit Function
End If

If Option1.value = False And Option2.value = False And Option3.value = False Then
    MsgBox "Seleccione un sector.", vbInformation, "Atención"
    Exit Function
End If

If Option1.value = True And txtAbrevia.Text = "" Then
    MsgBox "Para las entidades Financieras ingrese una Abreviatura.", vbInformation, "Atención"
    Exit Function
End If


ValidaDatos = True
End Function

Private Sub lblDisPF_Click()

End Sub

Private Sub grdGastos_GotFocus()
grdGastos.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdGastos_LostFocus()
grdGastos.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub grdIngresos_GotFocus()
grdIngresos.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdIngresos_LostFocus()
grdIngresos.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub Option1_Click()

If Option1.value = True Then
    Me.cmdDisAh.Enabled = True
    Me.cmdDepAh.Enabled = True
    Me.cmdIngAh.Enabled = True
    Me.cmdGasAh.Enabled = True
    Me.SSTab1.Enabled = True
End If

End Sub

Private Sub Option2_Click()
If Option1.value = False Then
    Me.cmdDisAh.Enabled = False
    Me.cmdDepAh.Enabled = False
    Me.cmdIngAh.Enabled = False
    Me.cmdGasAh.Enabled = False
    Me.SSTab1.Enabled = False
End If
End Sub

Private Sub Option3_Click()
If Option1.value = False Then
    Me.cmdDisAh.Enabled = False
    Me.cmdDepAh.Enabled = False
    Me.cmdIngAh.Enabled = False
    Me.cmdGasAh.Enabled = False
    Me.SSTab1.Enabled = False
End If

End Sub

