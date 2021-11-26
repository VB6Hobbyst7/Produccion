VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5F774E03-DB36-4DFC-AAC4-D35DC9379F2F}#1.1#0"; "VertMenu.ocx"
Begin VB.Form frmAsientoModificaCont 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asientos Contables: Mantenimiento"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11460
   Icon            =   "frmAsientoModificaCont.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkExcel 
      Caption         =   "Excel"
      Height          =   255
      Left            =   7560
      TabIndex        =   54
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecibo 
      Caption         =   "&Documento"
      Height          =   375
      Left            =   3885
      TabIndex        =   52
      ToolTipText     =   "Movimientos a los que Referencia este Asiento"
      Top             =   6060
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "&Anterior"
      Height          =   375
      Left            =   5100
      TabIndex        =   51
      ToolTipText     =   "Movimientos a los que Referencia este Asiento"
      Top             =   6060
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   "&Siguiente"
      Height          =   375
      Left            =   6315
      TabIndex        =   50
      ToolTipText     =   "Movimientos que Referencian a este Asiento"
      Top             =   6060
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10260
      Picture         =   "frmAsientoModificaCont.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "&Detalle ..."
      Height          =   375
      Left            =   8760
      TabIndex        =   19
      Top             =   6060
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   7530
      TabIndex        =   18
      Top             =   6060
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   9990
      TabIndex        =   20
      Top             =   6060
      Width           =   1200
   End
   Begin VB.Frame fraFechaBusca 
      Height          =   705
      Left            =   7140
      TabIndex        =   42
      Top             =   0
      Width           =   3045
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   390
         TabIndex        =   10
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFecha2 
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         Caption         =   "Al"
         Height          =   225
         Left            =   1590
         TabIndex        =   44
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Del"
         Height          =   225
         Left            =   90
         TabIndex        =   43
         Top             =   300
         Width           =   405
      End
   End
   Begin VB.Frame fraAsiento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4995
      Left            =   1500
      TabIndex        =   26
      Top             =   720
      Width           =   9855
      Begin VB.TextBox txtMovDesc 
         Height          =   345
         Left            =   750
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4500
         Width           =   8595
      End
      Begin MSDataGridLib.DataGrid dgMov 
         Height          =   3615
         Left            =   30
         TabIndex        =   13
         Top             =   180
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   6
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "cMovNro"
            Caption         =   "Nro.Movimiento"
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
            DataField       =   "nMovItem"
            Caption         =   "Item"
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
            DataField       =   "cCtaContCod"
            Caption         =   "Cuenta Contable"
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
            DataField       =   "cCtaContDesc"
            Caption         =   "Descripción"
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
            DataField       =   "nDebe"
            Caption         =   "Debe MN"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "nHaber"
            Caption         =   "Haber MN"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "nDebeME"
            Caption         =   "Debe ME"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "nHaberME"
            Caption         =   "Haber ME"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   2
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               DividerStyle    =   6
               ColumnWidth     =   2399.811
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   6
               ColumnWidth     =   1635.024
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   2099.906
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               DividerStyle    =   6
               Locked          =   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               DividerStyle    =   6
               Locked          =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Glosa"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   150
         TabIndex        =   36
         Top             =   4560
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total M.E. Haber"
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
         Left            =   6195
         TabIndex        =   34
         Top             =   4140
         Width           =   1470
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total M.E. Debe"
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
         Left            =   6195
         TabIndex        =   33
         Top             =   3870
         Width           =   1410
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total M.N. Haber"
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
         Left            =   3000
         TabIndex        =   32
         Top             =   4140
         Width           =   1485
      End
      Begin VB.Label LblTotHS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0.00"
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
         Height          =   255
         Left            =   4545
         TabIndex        =   31
         Top             =   4110
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total M.N. Debe"
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
         Left            =   2985
         TabIndex        =   30
         Top             =   3870
         Width           =   1425
      End
      Begin VB.Label LblTotDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0.00"
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
         Height          =   255
         Left            =   4560
         TabIndex        =   29
         Top             =   3870
         Width           =   1440
      End
      Begin VB.Label LblTotDD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0.00"
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
         Height          =   255
         Left            =   7695
         TabIndex        =   28
         Top             =   3870
         Width           =   1410
      End
      Begin VB.Label LblTotHD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0.00"
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
         Height          =   255
         Left            =   7710
         TabIndex        =   27
         Top             =   4110
         Width           =   1395
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000C&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   585
         Left            =   2880
         Top             =   3810
         Width           =   6465
      End
   End
   Begin VB.CommandButton cmdModifiDat 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   1560
      TabIndex        =   53
      ToolTipText     =   "Movimientos a los que Referencia este Asiento"
      Top             =   6060
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Extornar"
      Height          =   375
      Left            =   1560
      TabIndex        =   22
      ToolTipText     =   "Eliminar Asiento Contable"
      Top             =   6060
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Nueva Fecha"
      ForeColor       =   &H8000000D&
      Height          =   675
      Left            =   1500
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   3555
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2370
         TabIndex        =   17
         ToolTipText     =   "Cancelar Cambio de Fecha de Asiento"
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         Height          =   315
         Left            =   1290
         TabIndex        =   16
         ToolTipText     =   "Aplicar cambio de Fecha de Asiento"
         Top             =   240
         Width           =   1035
      End
      Begin MSMask.MaskEdBox txtMovFecha 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Mo&dificar"
      Height          =   375
      Left            =   2790
      TabIndex        =   24
      ToolTipText     =   "Modificar Asiento Contable"
      Top             =   6060
      Width           =   1185
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "Ca&mbiar"
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      ToolTipText     =   "Cambiar fecha de Registro de Asiento"
      Top             =   6060
      Width           =   1185
   End
   Begin VB.Frame fraAge 
      Height          =   705
      Left            =   1500
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   5595
      Begin Sicmact.TxtBuscar txtAgeCod 
         Height          =   330
         Left            =   870
         TabIndex        =   8
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
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
      Begin VB.TextBox txtAgeDesc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   3285
      End
      Begin VB.Label Label3 
         Caption         =   "Agencia"
         Height          =   225
         Left            =   150
         TabIndex        =   38
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.Frame fraOpe 
      Height          =   705
      Left            =   1500
      TabIndex        =   21
      Top             =   0
      Width           =   5595
      Begin Sicmact.TxtBuscar txtOpecod 
         Height          =   330
         Left            =   930
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
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
      Begin VB.TextBox txtOpeDes 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   3285
      End
      Begin VB.Label lblOpe 
         Caption         =   "Operación"
         Height          =   225
         Left            =   120
         TabIndex        =   35
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame fraDoc 
      Height          =   705
      Left            =   1500
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   5595
      Begin VB.TextBox txtDocSerie 
         Height          =   315
         Left            =   1860
         TabIndex        =   2
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox txtDocNro 
         Height          =   315
         Left            =   3540
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Documento :      Serie"
         Height          =   225
         Left            =   180
         TabIndex        =   41
         Top             =   300
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Número"
         Height          =   225
         Left            =   2850
         TabIndex        =   40
         Top             =   300
         Width           =   705
      End
   End
   Begin VB.Frame fraCta 
      Height          =   705
      Left            =   1500
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   5595
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3780
         MaxLength       =   16
         TabIndex        =   5
         Top             =   240
         Width           =   1005
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         ItemData        =   "frmAsientoModificaCont.frx":040C
         Left            =   4800
         List            =   "frmAsientoModificaCont.frx":0422
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   660
      End
      Begin VB.TextBox txtCtaCod 
         Height          =   315
         Left            =   1410
         TabIndex        =   4
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         Height          =   195
         Left            =   3180
         TabIndex        =   49
         Top             =   315
         Width           =   525
      End
      Begin VB.Label Label12 
         Caption         =   "Cuenta Contable"
         Height          =   225
         Left            =   120
         TabIndex        =   48
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.Frame fraMov 
      Height          =   705
      Left            =   1500
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   5595
      Begin VB.TextBox txtMovNro 
         Height          =   315
         Left            =   1860
         TabIndex        =   7
         Top             =   240
         Width           =   3315
      End
      Begin VB.Label Label11 
         Caption         =   "Nro. Movimiento"
         Height          =   225
         Left            =   420
         TabIndex        =   46
         Top             =   300
         Width           =   1275
      End
   End
   Begin VertMenu.VerticalMenu MnuOpe 
      Height          =   6465
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   11404
      MenuCaption1    =   "Buscar por ..."
      MenuItemsMax1   =   5
      MenuItemIcon11  =   "frmAsientoModificaCont.frx":043B
      MenuItemCaption11=   "Operación"
      MenuItemIcon12  =   "frmAsientoModificaCont.frx":0755
      MenuItemCaption12=   "Agencia"
      MenuItemIcon13  =   "frmAsientoModificaCont.frx":0A6F
      MenuItemCaption13=   "Documento"
      MenuItemIcon14  =   "frmAsientoModificaCont.frx":0D89
      MenuItemCaption14=   "Cuenta Contable"
      MenuItemIcon15  =   "frmAsientoModificaCont.frx":10A3
      MenuItemCaption15=   "Nro. Movimiento"
   End
End
Attribute VB_Name = "frmAsientoModificaCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim lbExtorno     As Boolean
Dim lbConsulta    As Boolean
Dim lbEliminaMov  As Boolean
Dim lbDatosNoCont As Boolean
Dim lnTpoBusca    As Integer
Dim lsUltMovNro   As String
Dim rsMov         As ADODB.Recordset
Dim lsOpeCod      As String
Dim MatOpeNoExtornable As Variant 'EJVG20140415

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(pbConsulta As Boolean, pbExtorno As Boolean, Optional pbEliminaMov As Boolean, Optional pnTpoBusca As Integer = 1, Optional pbDatosNoCont As Boolean = False)
lbConsulta = pbConsulta
lbExtorno = pbExtorno
lbEliminaMov = pbEliminaMov
lnTpoBusca = pnTpoBusca
lbDatosNoCont = pbDatosNoCont
   'lnTpoBusca = 2 Muestra Eliminados
Me.Show 1
End Sub

Private Sub cboFiltro_KeyPress(keyAscii As Integer)
If keyAscii = 13 Then
   txtFecha.SetFocus
End If
End Sub

Private Sub cmdAnterior_Click()
Dim clsAsiento As New ncontasientos
Dim sBusCond As String
If rsMov Is Nothing Then
    Exit Sub
End If
If rsMov.EOF Then
   Exit Sub
End If
sBusCond = " M.nMovNro IN (SELECT nMovNroRef FROM MovRef WHERE nMovNro = " & rsMov!nMovNro & " ) "
Set rsMov = clsAsiento.GetAsientoConsulta(sBusCond, "", "", "", "", "", "", gsMesCerrado)
Set dgMov.DataSource = rsMov
Set clsAsiento = Nothing
MousePointer = 0
dgMov.SetFocus
End Sub

Private Sub cmdAplicar_Click()
Dim nPos       As Variant
Dim sMovNro    As String
Dim sMovCambio As String
Dim oFun As New NContFunciones
Dim oMovCamb As New DMov 'NAGL 202008

   If MsgBox(" ¿ Seguro de Modificar Fecha de Movimiento ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
      If oMovCamb.PermiteModAsiContxCodOperación(rsMov!cOpeCod) = False Then
        MsgBox "Para el movimiento seleccionado, no está permitido la modificación de Fecha..!!", vbInformation, "Aviso"
        Exit Sub
      End If 'NAGL 202008 Según Acta N°063-2020
      
      If Month(txtMovFecha) <> nVal(Mid(rsMov!cMovNro, 5, 2)) And (Not IsNull(rsMov!nDebeME) Or Not IsNull(rsMov!nHaberME)) Then
         MsgBox "No se puede cambiar Asientos de Moneda Extranjera a un mes diferente", vbInformation, "¡Aviso!"
         Exit Sub
      End If
      sMovNro = oFun.GeneraMovNro(CDate(txtMovFecha), , , Format(txtMovFecha.Text, gsFormatoMovFecha) & Mid(rsMov!cMovNro, 9, 25))
      sMovCambio = oFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
      
      If Not oFun.PermiteModificarAsiento(sMovNro, False) Then
         MsgBox "No de puede cambiar fecha de Asiento a un mes ya Cerrado", vbInformation, "Aviso"
         Exit Sub
      End If
      Dim oMov As New DMov
      oMov.BeginTrans
      oMov.ActualizaMovimiento sMovNro, rsMov!nMovNro
      oMov.InsertaMovModifica sMovCambio, rsMov!cMovNro, sMovNro
      oMov.CommitTrans
      Set oMov = Nothing
      
      Dim oConst As New NConstSistemas
      oConst.ActualizaConstSistemas gConstSistUltActSaldos, GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge), txtMovFecha
      Set oConst = Nothing

      nPos = rsMov.Bookmark
      cmdBuscar_Click
      If nPos < rsMov.RecordCount Then
         rsMov.Bookmark = nPos
      Else
         If rsMov.RecordCount > 0 Then
            rsMov.MoveLast
         End If
      End If
      
      dgMov.SetFocus
   End If
   Set oFun = Nothing
   OpcionesCambiar False
End Sub

Private Sub cmdBuscar_Click()
Dim sOpeCond As String
Dim sAgeCond As String
Dim sMovCond As String
Dim sDocCond As String
Dim sCtaCond As String
Dim sFecCond As String
Dim sBusCond As String
sAgeCond = ""
sOpeCond = ""
sMovCond = ""
sDocCond = ""
sCtaCond = ""
sFecCond = ""
MousePointer = 11
   If fraOpe.Visible And txtOpeCod.Text <> "" Then
      sOpeCond = " and op.cOpeGruCod = '" & txtOpeCod.Text & "' "
   End If
   If fraAge.Visible And txtAgeCod.Text <> "" Then
      sAgeCond = " and cMovNro LIKE '______________" & gsCodCMAC & txtAgeCod.Text & "%' "
   End If
   If fraDoc.Visible And txtDocSerie <> "" Or txtDocNro <> "" Then
      sDocCond = " and md.cDocNro like '" & txtDocSerie & "%" & txtDocNro & "%' "
   End If
   If fraCta.Visible And txtCtaCod <> "" Then
      sCtaCond = " and a.cCtaContCod LIKE '" & txtCtaCod & "%' " & IIf(nVal(txtImporte) <> 0, " and ABS(nMovImporte) " & cboFiltro & nVal(txtImporte) & " ", "")
   End If
   If fraMov.Visible And txtMovNro <> "" Then
      sMovCond = " and cMovNro LIKE '" & txtMovNro & "%'"
   End If
   If fraFechaBusca.Visible Then
      If Trim(txtFecha) <> "/  /" And Trim(txtFecha2) <> "/  /" Then
         sFecCond = " LEFT(M.cMovNro,8) BETWEEN '" & Format(txtFecha, gsFormatoMovFecha) & "' and " _
                  & " '" & Format(txtFecha2, gsFormatoMovFecha) & "' "
      ElseIf Trim(txtFecha) <> "/  /" Then
         sFecCond = " M.cMovNro LIKE '" & Format(txtFecha, gsFormatoMovFecha) & "%' "
      ElseIf Trim(txtFecha2) <> "/  /" Then
         sFecCond = " M.cMovNro LIKE '" & Format(txtFecha2, gsFormatoMovFecha) & "%' "
      Else
         sFecCond = " M.cMovNro LIKE '_%' "
      End If
   End If
   Select Case lnTpoBusca
      Case 1: sBusCond = IIf(sFecCond = "", "", " and ") & "M.nMovEstado = " & gMovEstContabMovContable & " and M.nMovFlag <> " & gMovFlagEliminado & " and a.cCtaContCod <> '' "
      Case 2: sBusCond = IIf(sFecCond = "", "", " and ") & "M.nMovFlag = " & gMovFlagEliminado
      Case Else: sBusCond = ""
   End Select
Dim clsAsiento As New ncontasientos
Set rsMov = clsAsiento.GetAsientoConsulta(sBusCond, sOpeCond, sAgeCond, sDocCond, sCtaCond, sMovCond, sFecCond, gsMesCerrado)
Set dgMov.DataSource = rsMov
Set clsAsiento = Nothing
MousePointer = 0
dgMov.SetFocus
End Sub

Private Sub cmdCambiar_Click()
Dim oCont As New NContFunciones
On Error GoTo CambiarErr
If Left(lsOpeCod, 1) = "5" Then
    MsgBox "Las operaciones de Logistica no se pueden Modificar por esta opción. " & Chr(10) & "Realizarlo a través del respectivo Sistema ", vbInformation, "¡Aviso!"
    Exit Sub
End If
If Left(lsOpeCod, 1) = "4" Then
    MsgBox "Las operaciones de Caja General no se pueden Modificar por esta opción. " & Chr(10) & "Realizarlo a través de sus respectivas Opciones", vbInformation, "¡Aviso!"
    Exit Sub
End If

If rsMov Is Nothing Then
    Exit Sub
End If
If rsMov.EOF Then
    Exit Sub
End If
If Not oCont.PermiteModificarAsiento(rsMov!cMovNro) Then
   Exit Sub
End If
OpcionesCambiar True
txtMovFecha.SetFocus
Exit Sub
CambiarErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdCancelar_Click()
   OpcionesCambiar False

End Sub

Private Sub cmdDetalle_Click()
Dim N As Variant
On Error GoTo ErrModifica

If rsMov Is Nothing Then Exit Sub
If rsMov.EOF And rsMov.BOF Then Exit Sub

glAceptar = False
frmAsientoModDatosNoContable.Inicio rsMov!cMovNro, rsMov!nMovNro, False
If glAceptar Then
   N = rsMov.Bookmark
   cmdBuscar_Click
   Dim oCont As New NContFunciones
   oCont.UbicarEnRegistro rsMov, N
   Set oCont = Nothing
   dgMov.SetFocus
End If
Set frmAsientoModDatosNoContable = Nothing
Exit Sub
ErrModifica:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdEliminar_Click()
Dim nPos As Variant
Dim sMov As String
Dim pscCtaContCod As String, psOpeCod As String 'NAGL 202008
On Error GoTo ErrDel

'If Left(lsOpeCod, 1) = "5" Then
'    MsgBox "Las operaciones de Logistica no se pueden Modificar por esta opción. " & Chr(10) & "Realizarlo a través del respectivo Sistema ", vbInformation, "¡Aviso!"
'    Exit Sub
'End If
'If Left(lsOpeCod, 1) = "4" Then
'    MsgBox "Las operaciones de Caja General no se pueden Modificar por esta opción. " & Chr(10) & "Realizarlo a través de sus respectivas Opciones", vbInformation, "¡Aviso!"
'    Exit Sub
'End If

If rsMov!nMovFlag = gMovFlagExtornado Then
   MsgBox "Este Movimiento ya fue Extornado", vbInformation, "¡Aviso!"
   Exit Sub
Else
    If Not PermiteModificarAsiento(rsMov!cMovNro, False) Then
        If lbEliminaMov Or lnTpoBusca = 2 Then
            MsgBox " Operación corresponde a Mes ya Cerrado. Puede realizar Extorno con Generación de Asiento", vbInformation, "¡Aviso!"
            Exit Sub
        Else
            If MsgBox(" Operación corresponde a Mes ya Cerrado. ¿ Desea Continuar ? ", vbQuestion + vbYesNo, "¡Aviso!") = vbNo Then
                Exit Sub
            End If
        End If
    End If
End If
nPos = -1
If lbEliminaMov Or lnTpoBusca = 2 Then
    'EJVG20140415 ***
    If Not EsOperacionExtornable(rsMov!cOpeCod) Then
        MsgBox "No se puede continuar, ya que la operación " & rsMov!cOpeCod & " es una operación no extornable", vbInformation, "Aviso"
        Exit Sub
    End If
    'END EJVG *******
   If rsMov!nMovFlag = gMovFlagEliminado Then
      sMov = " Movimiento ya fue Eliminado. ¿ Desea Recuperarlo ? "
   Else
      sMov = " ¿ Seguro de Extornar Movimiento ? "
   End If
      If MsgBox(sMov, vbQuestion + vbYesNo, "Confirmación") = vbNo Then
         Exit Sub
      End If
   
      nPos = rsMov.Bookmark
      Dim oMov As New DMov
      oMov.BeginTrans
      sMov = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
      If rsMov!nMovFlag = gMovFlagEliminado Then
         oMov.ActualizaMov rsMov!nMovNro, , rsMov!nMovEstado, gMovFlagVigente
         If rsMov!nMovEstado = gMovEstContabMovContable Then
            oMov.ActualizaSaldoMovimiento rsMov!cMovNro, "+"
         End If
      Else
         oMov.ActualizaMovPendientesRendCambio rsMov!nMovNro
         oMov.EliminaCuentasProAsiento rsMov!nMovNro, rsMov!cMovNro, rsMov!cOpeCod 'NAGL 202008 Según Acta N°063-2020
         oMov.EliminaMov rsMov!cMovNro

         '*** PEAC 20120503
         oMov.ModificaActivosFijos rsMov!cMovNro, rsMov!cOpeCod

      End If
      oMov.InsertaMovModifica sMov, rsMov!cMovNro
      pscCtaContCod = rsMov!cCtaContCod 'NAGL 202008
      psOpeCod = rsMov!cOpeCod 'NAGL 202008
      oMov.CommitTrans
      Set oMov = Nothing
      lsUltMovNro = ""
      cmdBuscar_Click
Else
   If lbExtorno Then
        'EJVG20140415 ***
        If Not EsOperacionExtornable(rsMov!cOpeCod) Then
            MsgBox "No se puede continuar, ya que la operación " & rsMov!cOpeCod & " es una operación no extornable", vbInformation, "Aviso"
            Exit Sub
        End If
        'END EJVG *******
      glAceptar = False
      nPos = rsMov.Bookmark
      frmAsientoRegistro.Inicio rsMov!cMovNro, rsMov!nMovNro, True
      If glAceptar Then
         lsUltMovNro = ""
         cmdBuscar_Click
      End If
   End If
End If
If Not nPos = -1 Then
   Dim oCont As New NContFunciones
   oCont.UbicarEnRegistro rsMov, nPos
   Set oCont = Nothing
End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = rsMov!cCtaContCod 'Comentado by NAGL 202008
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Se Extorno Con Eliminación del asiento " & pscCtaContCod & "del Tipo de Operación : " & psOpeCod & " por Motivo " & txtMovDesc.Text
            'NAGL 202008 Cambió rsMov!cCtaContCod a pscCtaContCod y rsMov!cOpeCod a psOpeCod
            Set objPista = Nothing
            '*******
            txtMovDesc.Text = "" 'NAGL 202008
dgMov.SetFocus
Sumas
Exit Sub
ErrDel:
    MsgBox TextErr(Err.Description), vbInformation, "Error"
End Sub

Private Sub cmdImprimir_Click()
Dim oMov As New NContImprimir
Dim sTexto As String
Me.Enabled = False

If chkExcel.value = 1 Then
   ImprimeAsientoExcel
Else
   sTexto = oMov.ImprimeAsientoContable(lsUltMovNro, gnLinPage, gnColPage)
   EnviaPrevio sTexto, "Asiento Contable", False
End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaRegistoAsientCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio los Asientos Contables"
            Set objPista = Nothing
            '*******
Me.Enabled = True
End Sub
Private Sub ImprimeAsientoExcel()
Dim oCtaIf As NCajaCtaIF
Dim lsMoneda As String
Dim rs As ADODB.Recordset

Dim fs              As Scripting.FileSystemObject
Dim xlAplicacion    As Excel.Application
Dim xlLibro         As Excel.Workbook
Dim xlHoja1         As Excel.Worksheet
Dim lbExisteHoja    As Boolean
Dim lilineas        As Integer
Dim i               As Integer
Dim glsArchivo      As String
Dim lsNomHoja       As String

On Error GoTo ReporteAdeudadosVinculadosErr
    
    glsArchivo = "Asiento_Contable" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLSX"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    xlHoja1.PageSetup.Orientation = xlLandscape

            lbExisteHoja = False
            lsNomHoja = "Asiento Contable"
            For Each xlHoja1 In xlLibro.Worksheets
                If xlHoja1.Name = lsNomHoja Then
                    xlHoja1.Activate
                    lbExisteHoja = True
                    Exit For
                End If
            Next
            If lbExisteHoja = False Then
                Set xlHoja1 = xlLibro.Worksheets.Add
                xlHoja1.Name = lsNomHoja
            End If

            xlAplicacion.Range("A1:A1").ColumnWidth = 20
            xlAplicacion.Range("B1:B1").ColumnWidth = 37
            xlAplicacion.Range("c1:c1").ColumnWidth = 20
            xlAplicacion.Range("D1:D1").ColumnWidth = 20
            xlAplicacion.Range("E1:E1").ColumnWidth = 20
            xlAplicacion.Range("F1:F1").ColumnWidth = 20
           
            xlAplicacion.Range("A1:Z100").Font.Size = 9
            xlAplicacion.Range("A1:Z100").Font.Name = "Century Gothic"
       
            xlHoja1.Cells(1, 1) = gsNomCmac
            xlHoja1.Cells(2, 1) = "Asiento Contable" & space(5) & "DEL " & space(5) & Format(txtFecha, "dd/mm/yyyy")
            
            xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 3)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 3)).Merge True
            xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 2)).Merge True
                 
                      
            lilineas = 4
            
            xlHoja1.Cells(lilineas, 1) = "Cta"
            xlHoja1.Cells(lilineas, 2) = "Descripcion"
            xlHoja1.Cells(lilineas, 3) = "DEBE"
            xlHoja1.Cells(lilineas, 4) = "HABER"
            xlHoja1.Cells(lilineas, 5) = "DEBE ME"
            xlHoja1.Cells(lilineas, 6) = "HABER ME"
   
            
            xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas + 3, 1)).Font.Bold = True
            
            xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 12)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 12)).VerticalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas + 3, 1)).Merge True
            xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 12)).EntireRow.AutoFit
            xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 12)).WrapText = True
            
            xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 12)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 12)).Borders.LineStyle = 1
            xlHoja1.Range(xlHoja1.Cells(lilineas, 1), xlHoja1.Cells(lilineas, 12)).Interior.ColorIndex = 36 '.Color = RGB(159, 206, 238)
            
         
            
            lilineas = lilineas + 1
            rsMov.MoveFirst
         Do Until rsMov.EOF
            xlHoja1.Cells(lilineas, 1) = rsMov(3)
            xlHoja1.Cells(lilineas, 2) = rsMov(4)
            xlHoja1.Cells(lilineas, 3) = rsMov(5)
            xlHoja1.Cells(lilineas, 4) = rsMov(6)
            xlHoja1.Cells(lilineas, 5) = rsMov(7)
            xlHoja1.Cells(lilineas, 6) = rsMov(8)
            
            
            xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 5)).Style = "Comma"
            xlHoja1.Range(xlHoja1.Cells(lilineas, 8), xlHoja1.Cells(lilineas, 8)).Style = "Comma"
            
            xlHoja1.Range(xlHoja1.Cells(lilineas, 3), xlHoja1.Cells(lilineas, 5)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lilineas, 7), xlHoja1.Cells(lilineas, 7)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lilineas, 10), xlHoja1.Cells(lilineas, 11)).HorizontalAlignment = xlCenter
            
            xlHoja1.Range(xlHoja1.Cells(lilineas, 6), xlHoja1.Cells(lilineas, 6)).HorizontalAlignment = xlRight
            xlHoja1.Range(xlHoja1.Cells(lilineas, 10), xlHoja1.Cells(lilineas, 10)).HorizontalAlignment = xlRight
            
            lilineas = lilineas + 1
            rsMov.MoveNext
        Loop

        ExcelCuadro xlHoja1, 1, 4, 12, lilineas - 1
        
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        ExcelEnd App.path & "\Spooler\" & glsArchivo, xlAplicacion, xlLibro, xlHoja1
    
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        Call CargaArchivo(glsArchivo, App.path & "\SPOOLER\")
                    
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaRegistoAsientCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio en Excel los Asientos Contables"
            Set objPista = Nothing
            '*******
    
Set oCtaIf = Nothing
    Exit Sub
ReporteAdeudadosVinculadosErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub



End Sub


Private Sub CmdModificar_Click()
Dim N As Variant
Dim oMov As New DMov 'NAGL 202008
On Error GoTo ErrModifica
If rsMov Is Nothing Then
    Exit Sub
End If
If rsMov.EOF Then
    Exit Sub
End If
DoEvents
'If Left(lsOpeCod, 1) = "5" Then
'    MsgBox "Las operaciones de Logistica no se pueden Modificar por esta opción. " & Chr(10) & "Realizarlo a través del respectivo Sistema ", vbInformation, "¡Aviso!"
'    Exit Sub
'End If
'If Left(lsOpeCod, 1) = "4" Then
'    If Not (Left(lsOpeCod, 4) = Left(gCGArendirCtaSolMN, 4) Or Left(lsOpeCod, 4) = Left(gCGArendirCtaSolME, 4) Or Left(lsOpeCod, 4) = Left(gCGArendirViatSolMN, 4) Or Left(lsOpeCod, 4) = Left(gCGArendirViatSolME, 4) Or Left(lsOpeCod, 4) = Left(gCHAutorizaDesembMN, 4) Or Left(lsOpeCod, 4) = Left(gCHAutorizaDesembME, 4)) Then
'        MsgBox "Las operaciones de Caja General no se pueden Modificar por esta opción. " & Chr(10) & "Realizarlo a través de sus respectivas Opciones", vbInformation, "¡Aviso!"
'        Exit Sub
'    End If
'End If

If Not PermiteModificarAsiento(rsMov!cMovNro) Then
   Exit Sub
End If
If oMov.PermiteModAsiContxCodOperación(rsMov!cOpeCod) = False Then
    MsgBox "Para el movimiento seleccionado, no está permitido la modificación de Asientos..!!", vbInformation, "Aviso"
    Exit Sub
End If 'NAGL 202008 Según Acta N°063-2020

glAceptar = False
frmAsientoRegistro.Inicio rsMov!cMovNro, rsMov!nMovNro
If glAceptar Then
   N = rsMov.Bookmark
   cmdBuscar_Click
   Dim oCont As New NContFunciones
   oCont.UbicarEnRegistro rsMov, N
   Set oCont = Nothing
   dgMov.SetFocus
End If
Exit Sub
ErrModifica:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdModifiDat_Click()
Dim N As Variant
On Error GoTo ErrModifica
glAceptar = False
frmAsientoModDatosNoContable.Inicio rsMov!cMovNro, rsMov!nMovNro, True
If glAceptar Then
   N = rsMov.Bookmark
   cmdBuscar_Click
   Dim oCont As New NContFunciones
   oCont.UbicarEnRegistro rsMov, N
   Set oCont = Nothing
   dgMov.SetFocus
End If
Set frmAsientoModDatosNoContable = Nothing
Exit Sub
ErrModifica:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"

End Sub

Private Sub cmdRecibo_Click()
Dim oConImp As NContImprimir
Dim lsTexto As String
Dim sSql As String
Dim prs  As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Set oConImp = New NContImprimir
Set prs = New ADODB.Recordset
If rsMov Is Nothing Then
    Exit Sub
End If
If rsMov.EOF Then
    Exit Sub
End If
oCon.AbreConexion

    sSql = "SELECT mg.nMovNro, mg.cPersCod, p.cPersNombre, md.cDocNro, md.dDocFecha, m.cMovDesc, m.cOpecod , ABS(SUM(mc.nMovImporte)) nMovImporte " _
         & "FROM MovGasto mg JOIN Mov m ON m.nMovNro = mg.nMovNro JOIN Persona p ON p.cPersCod = mg.cPersCod " _
         & "     JOIN MovDoc md ON md.nMovNro = mg.nMovNro JOIN MovCta mc ON mc.nMovNro = mg.nMovNro " _
         & "WHERE mg.nMovNro = " & rsMov!nMovNro & " and LEFT(mc.cCtaContCod,6) IN ('251601','251602','252601','252602','191106','192106','281807','292807','111701','112701') and md.nDocTpo IN (67," & TpoDocRecEgreso & ")" _
         & "GROUP BY mg.nMovNro, mg.cPersCod, p.cPersNombre, md.cDocNro, md.dDocFecha, m.cMovDesc, m.cOpecod  "
    Set prs = oCon.CargaRecordSet(sSql)
    If Not prs.EOF Then
        lsTexto = oConImp.ImprimeReciboEgresos(gnColPage, rsMov!cMovNro, prs!cMovDesc, GetFechaMov(rsMov!cMovNro, True), gsNomCmac, prs!cOpeCod, False, "", ArendirAtencion, "", prs!cDocNro, _
                  prs!dDocFecha, prs!cPersCod, prs!cPersNombre, "", prs!nMovImporte)
    Else
        MsgBox "Sistema no emite formato de Tipo de Documento de esta Operación", vbInformation, "¡Aviso"
    End If
    RSClose prs
    If lsTexto <> "" Then
        EnviaPrevio lsTexto, Me.Caption, gnLinPage
    End If
oCon.CierraConexion
Set oCon = Nothing

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSiguiente_Click()
Dim clsAsiento As New ncontasientos
Dim sBusCond As String
If rsMov Is Nothing Then
    Exit Sub
End If
If rsMov.EOF Then
   Exit Sub
End If
sBusCond = " M.nMovNro IN (SELECT nMovNro FROM MovRef WHERE nMovNroRef = " & rsMov!nMovNro & " ) "
Set rsMov = clsAsiento.GetAsientoConsulta(sBusCond, "", "", "", "", "", "", gsMesCerrado)
Set dgMov.DataSource = rsMov
Set clsAsiento = Nothing
MousePointer = 0
dgMov.SetFocus
End Sub

Private Sub dgMov_HeadClick(ByVal ColIndex As Integer)
If Not rsMov Is Nothing Then
   If Not rsMov.EOF Then
      rsMov.Sort = dgMov.Columns(ColIndex).DataField
   End If
End If
End Sub

Private Sub dgMov_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
RefrescaDatos
End Sub

Private Sub Form_Load()
Dim clsOpe As New DOperacion
CentraForm Me
txtOpeCod.rs = clsOpe.CargaOpeGru
Set clsOpe = Nothing
cmdAnterior.Visible = True
cmdSiguiente.Visible = True
If lbConsulta Then
   Me.Caption = "Asientos Contables: Consulta"
   If lnTpoBusca = 2 Then
      Me.Caption = Me.Caption & " de Movimientos Eliminados"
   End If
End If
If lbExtorno Then
   Me.Caption = "Asientos Contables: Extorno con Generación de Asiento"
End If
If lbEliminaMov Then
   Me.Caption = "Asientos Contables: Extorno con Eliminación de Asiento"
End If
Dim clsRHArea As New DActualizaDatosArea
txtAgeCod.rs = clsRHArea.GetAgencias
Set clsRHArea = Nothing

If lbConsulta Or lbExtorno Then
   cmdRecibo.Visible = True
   cmdModificar.Visible = False
   cmdCambiar.Visible = False
End If
If lbExtorno Or lnTpoBusca = 2 Then
   cmdEliminar.Visible = True
End If
If lbDatosNoCont Then
    Me.Caption = "Operaciones: Modificación de Datos no Contables "
    cmdModifiDat.Visible = True
    cmdCambiar.Visible = False
    cmdRecibo.Visible = False
    cmdModificar.Visible = False
End If
cboFiltro.ListIndex = 0
lsUltMovNro = ""
MatOpeNoExtornable = DameOperacionesNoExtornables 'EJVG20140415
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsMov
End Sub

Private Sub MnuOpe_MenuItemClick(MenuNumber As Long, MenuItem As Long)
MnuOpe.MenuItemCur = MenuItem
Select Case MenuItem
   Case 1:  ActivaFrame True, False, False, False, False
            fraFechaBusca.Visible = True
            txtOpeCod.SetFocus
   Case 2:  ActivaFrame False, True, False, False, False
            fraFechaBusca.Visible = True
            txtAgeCod.SetFocus
   Case 3:  ActivaFrame False, False, True, False, False
            fraFechaBusca.Visible = True
            txtDocSerie.SetFocus
   Case 4:  ActivaFrame False, False, False, True, False
            fraFechaBusca.Visible = True
            txtCtaCod.SetFocus
   Case 5:  ActivaFrame False, False, False, False, True
            fraFechaBusca.Visible = False
            txtMovNro.SetFocus
End Select
End Sub

Private Sub txtAgeCod_EmiteDatos()
txtAgeDesc = txtAgeCod.psDescripcion
If txtAgeDesc <> "" Then
   txtFecha.SetFocus
End If
End Sub

Private Sub txtCtaCod_KeyPress(keyAscii As Integer)
keyAscii = NumerosEnteros(keyAscii)
If keyAscii = 13 Then
   txtImporte.SetFocus
End If
End Sub

Private Sub txtDocNro_KeyPress(keyAscii As Integer)
If keyAscii = 13 Then
   cmdBuscar.SetFocus
End If
End Sub

Private Sub txtDocSerie_KeyPress(keyAscii As Integer)
If keyAscii = 13 Then
   txtDocNro.SetFocus
End If
End Sub

Private Sub txtFecha_GotFocus()
txtFecha.SelStart = 0
txtFecha.SelLength = Len(txtFecha)
End Sub

Private Sub txtFecha_KeyPress(keyAscii As Integer)
If keyAscii = 13 Then
   If Trim(txtFecha) = "/  /" Then
   Else
      If ValidaFecha(txtFecha) <> "" Then
         MsgBox "Fecha no Valida...!"
         Exit Sub
      End If
   End If
   txtFecha2.SetFocus
End If
End Sub

Private Sub txtFecha2_GotFocus()
txtFecha2.SelStart = 0
txtFecha2.SelLength = Len(txtFecha2)
End Sub

Private Sub txtFecha2_KeyPress(keyAscii As Integer)
If keyAscii = 13 Then
   If Trim(txtFecha2) = "/  /" Then
   Else
      If ValidaFecha(txtFecha2) <> "" Then
         MsgBox "Fecha no Valida...!"
         Exit Sub
      End If
   End If
   cmdBuscar.SetFocus
End If
End Sub

Private Sub txtImporte_KeyPress(keyAscii As Integer)
keyAscii = NumerosDecimales(txtImporte, keyAscii, 16, 2)
If keyAscii = 13 Then
   txtImporte = Format(txtImporte, gsFormatoNumeroView)
   cboFiltro.SetFocus
End If
End Sub

Private Sub txtMovFecha_KeyPress(keyAscii As Integer)
Dim sMov As String
If keyAscii = 13 Then
   If ValidaFecha(txtMovFecha.Text) <> "" Then
      MsgBox "Fecha no Válida....!", vbInformation, "Error"
      Exit Sub
   End If
   Dim oMov As New DMov
   sMov = oMov.GeneraMovNro(CDate(txtMovFecha), Format(txtMovFecha, gsFormatoFecha) & Mid(lsUltMovNro, 9, 25))
   Set oMov = Nothing
   If Not PermiteModificarAsiento(sMov) Then
      Exit Sub
   End If
   cmdAplicar.SetFocus
End If
End Sub

Private Sub txtMovFecha_Validate(Cancel As Boolean)
Dim sMov As String
If ValidaFecha(txtMovFecha.Text) <> "" Then
   MsgBox "Fecha no Válida....!", vbInformation, "Error"
   Cancel = True
Else
   Dim oMov As New DMov
   sMov = oMov.GeneraMovNro(CDate(txtMovFecha), gsCodAge, gsCodUser)
   Set oMov = Nothing
   If Not PermiteModificarAsiento(sMov) Then
      Cancel = True
   End If
End If
End Sub

Private Sub txtMovNro_KeyPress(keyAscii As Integer)
keyAscii = Asc(UCase(Chr(keyAscii)))
If keyAscii = 13 Then
   cmdBuscar.SetFocus
End If
End Sub

Private Sub txtOpeCod_EmiteDatos()
txtOpeDes = txtOpeCod.psDescripcion
If txtOpeCod <> "" Then
   txtFecha.SetFocus
End If
End Sub

'**************** FUNCIONES PRIVADA *************
Private Sub ActivaFrame(pbOpe As Boolean, pbAge As Boolean, pbDoc As Boolean, pbCta As Boolean, pbMov As Boolean)
fraOpe.Visible = pbOpe
fraAge.Visible = pbAge
fraDoc.Visible = pbDoc
fraCta.Visible = pbCta
fraMov.Visible = pbMov
End Sub

Private Sub RefrescaDatos()
Dim lsEstado As String
If rsMov Is Nothing Then
   Exit Sub
End If
If Not rsMov.EOF And Not rsMov.BOF Then
   If rsMov!cMovNro <> lsUltMovNro Then
      MuestraDatosMov rsMov!cMovNro
      Sumas
      fraAsiento.Caption = ""
      lsEstado = "MOVIMIENTO "
      Select Case rsMov!nMovEstado
        Case gMovEstContabNoContable: lsEstado = lsEstado & "NO-CONTABLE "
        Case gMovEstContabPendiente: lsEstado = lsEstado & "PENDIENTE "
        Case gMovEstContabRechazado: lsEstado = lsEstado & "RECHAZADO "
        Case gMovEstLogIngBienAceptado: lsEstado = lsEstado & "ING.BIEN ACEPTADO "
        Case gMovEstLogIngBienRechazado: lsEstado = lsEstado & "ING.BIEN RECHAZADO "
        Case gMovEstLogSaleBienAlmacen: lsEstado = lsEstado & "SALIDA DE ALMACEN "
        Case 14: lsEstado = lsEstado & "NO-CONTABLE "
      End Select
      Select Case rsMov!nMovFlag
         Case gMovFlagEliminado: lsEstado = lsEstado & "ELIMINADO"
         Case gMovFlagDeExtorno: lsEstado = lsEstado & "DE EXTORNO"
         Case gMovFlagExtornado: lsEstado = lsEstado & "EXTORNADO"
         Case gMovFlagModificado: lsEstado = lsEstado & "MODIFICADO"
      End Select
      If Not lsEstado = "MOVIMIENTO " Then
          fraAsiento.Caption = lsEstado
      End If
   End If
End If
End Sub

Private Sub MuestraDatosMov(vMovNro As String)
Dim prs As New ADODB.Recordset
Dim oMov As New DMov
txtMovDesc = ""
Set prs = oMov.CargaMovOpeAsiento(0, vMovNro)
If Not prs.EOF Then
   txtMovDesc = prs!cMovDesc
   lsOpeCod = prs!cOpeCod
   lsUltMovNro = prs!cMovNro
End If
RSClose prs
Set oMov = Nothing
End Sub

Private Sub Sumas()
Dim prs As New ADODB.Recordset
Dim oMov As New DMov
LblTotDS.Caption = ""
LblTotDD.Caption = ""
LblTotHS.Caption = ""
LblTotHD.Caption = ""
Set prs = oMov.CargaSumaMovAsiento(lsUltMovNro)
If Not prs.EOF Then
   LblTotDS.Caption = Format(prs!nDebe, gsFormatoNumeroView)
   LblTotDD.Caption = Format(prs!nDebeME, gsFormatoNumeroView)
   LblTotHS.Caption = Format(prs!nHaber, gsFormatoNumeroView)
   LblTotHD.Caption = Format(prs!nHaberME, gsFormatoNumeroView)
End If
RSClose prs
Set oMov = Nothing
End Sub

Private Sub OpcionesCambiar(lOp As Boolean)
fraFecha.Visible = lOp
cmdCambiar.Visible = Not lOp
cmdModificar.Visible = Not lOp
If lbExtorno Then
   cmdEliminar.Visible = Not lOp
End If
txtMovFecha.Visible = lOp
End Sub

'EJVG20140415 ***
Public Function DameOperacionesNoExtornables() As Variant
    Dim oCS As New DConstSistemas
    Dim MatOpe As Variant
    Dim lsListaOpe As String
    On Error GoTo ErrDameOperacionesNoExtornables
    lsListaOpe = oCS.LeeConstSistema(469)
    MatOpe = Split(lsListaOpe, ",")
    DameOperacionesNoExtornables = MatOpe
    Set oCS = Nothing
    Exit Function
ErrDameOperacionesNoExtornables:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
Public Function EsOperacionExtornable(ByVal psOpeCod As String) As Boolean
    Dim i As Integer
    EsOperacionExtornable = True
    If IsArray(MatOpeNoExtornable) Then
        For i = 0 To UBound(MatOpeNoExtornable)
            If psOpeCod = MatOpeNoExtornable(i) Then
                EsOperacionExtornable = False
                Exit For
            End If
        Next
    End If
End Function
'END EJVG *******
