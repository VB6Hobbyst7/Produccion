VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCajaGenAperCtas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8400
   ClientLeft      =   765
   ClientTop       =   1455
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajaGenAperCtas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Fecha"
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
      Height          =   705
      Left            =   7410
      TabIndex        =   38
      Top             =   90
      Width           =   1605
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   345
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
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
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cuentas a Aperturar"
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
      Height          =   2175
      Left            =   150
      TabIndex        =   37
      Top             =   1200
      Width           =   8835
      Begin VB.TextBox txtTotApertura 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Height          =   285
         Left            =   6930
         TabIndex        =   7
         Tag             =   "0"
         Text            =   "0.00"
         Top             =   1710
         Width           =   1680
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   1290
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   150
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin Sicmact.FlexEdit fgCta 
         Height          =   1395
         Left            =   150
         TabIndex        =   4
         Top             =   240
         Width           =   8535
         _extentx        =   15055
         _extenty        =   2461
         cols0           =   9
         highlight       =   1
         rowsizingmode   =   1
         encabezadosnombres=   "#-Codigo-Cuenta_Nro-Importe-Plazo-PerTasa-T.Int.Efect-Check-MontoEuros"
         encabezadosanchos=   "300-1200-2700-1400-600-850-1200-1200-1400"
         font            =   "frmCajaGenAperCtas.frx":030A
         font            =   "frmCajaGenAperCtas.frx":0336
         font            =   "frmCajaGenAperCtas.frx":0362
         font            =   "frmCajaGenAperCtas.frx":038E
         font            =   "frmCajaGenAperCtas.frx":03BA
         fontfixed       =   "frmCajaGenAperCtas.frx":03E6
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-2-3-4-5-6-7-8"
         textstylefixed  =   3
         listacontroles  =   "0-0-0-0-0-0-0-4-0"
         encabezadosalineacion=   "C-L-L-R-R-R-R-L-R"
         formatosedit    =   "0-0-0-2-3-3-2-0-2"
         cantentero      =   12
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbbuscaduplicadotext=   -1
         colwidth0       =   300
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total Aperturas"
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
         Left            =   5490
         TabIndex        =   40
         Top             =   1720
         Width           =   1320
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   5340
         Top             =   1700
         Width           =   3285
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7650
      TabIndex        =   27
      Top             =   7920
      Width           =   1275
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6330
      TabIndex        =   26
      Top             =   7920
      Width           =   1275
   End
   Begin VB.Frame fradatosGen 
      Caption         =   "Datos Generales"
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
      Height          =   1110
      Left            =   150
      TabIndex        =   28
      Top             =   90
      Width           =   7185
      Begin VB.ComboBox cmbTipoPlazoFijo 
         Height          =   330
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   660
         Visible         =   0   'False
         Width           =   1950
      End
      Begin Sicmact.TxtBuscar txtBuscaIF 
         Height          =   360
         Left            =   1005
         TabIndex        =   1
         Top             =   240
         Width           =   1995
         _extentx        =   3519
         _extenty        =   635
         appearance      =   1
         appearance      =   1
         font            =   "frmCajaGenAperCtas.frx":0414
         appearance      =   1
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo :"
         Height          =   210
         Left            =   570
         TabIndex        =   46
         Top             =   690
         Visible         =   0   'False
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDescTipoCta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3030
         TabIndex        =   3
         Top             =   660
         Width           =   4005
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Institución Financiera :"
         Height          =   420
         Left            =   135
         TabIndex        =   29
         Top             =   195
         Width           =   900
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDescIF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3030
         TabIndex        =   2
         Top             =   255
         Width           =   4005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1380
      Left            =   150
      TabIndex        =   30
      Top             =   3360
      Width           =   8835
      Begin VB.CommandButton cmdCartAper 
         Caption         =   "&Carta Apertura"
         Height          =   375
         Left            =   7290
         TabIndex        =   9
         Top             =   210
         Width           =   1380
      End
      Begin VB.TextBox txtMovDesc 
         Height          =   600
         Left            =   135
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   255
         Width           =   7050
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Height          =   285
         Left            =   5235
         TabIndex        =   10
         Tag             =   "2"
         Text            =   "0.00"
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total Retiros :"
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
         Left            =   3780
         TabIndex        =   31
         Top             =   1005
         Width           =   1230
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         FillColor       =   &H00C0C0C0&
         Height          =   345
         Left            =   3600
         Top             =   930
         Width           =   3345
      End
   End
   Begin TabDlg.SSTab TabDoc 
      Height          =   3045
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   5371
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Transferencia"
      TabPicture(0)   =   "frmCajaGenAperCtas.frx":0438
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDocTrans"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTransferencia"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkDocOrigen"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Efectivo"
      TabPicture(1)   =   "frmCajaGenAperCtas.frx":0454
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdEfectivo"
      Tab(1).Control(1)=   "txtBilleteImporte"
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(3)=   "Shape2"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Cheque Recibido"
      TabPicture(2)   =   "frmCajaGenAperCtas.frx":0470
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdChqRecibido"
      Tab(2).Control(1)=   "txtChqRecImporte"
      Tab(2).Control(2)=   "fgChqRecibido"
      Tab(2).Control(3)=   "Label2"
      Tab(2).Control(4)=   "Shape5"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Otros "
      TabPicture(3)   =   "frmCajaGenAperCtas.frx":048C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtTotalOtrasCtas"
      Tab(3).Control(1)=   "fgOtros"
      Tab(3).Control(2)=   "cmdAgregarCta"
      Tab(3).Control(3)=   "cmdEliminarCta"
      Tab(3).Control(4)=   "fgObj"
      Tab(3).Control(5)=   "Label18"
      Tab(3).Control(6)=   "Shape3"
      Tab(3).ControlCount=   7
      Begin VB.CommandButton cmdChqRecibido 
         Caption         =   "&Cheques "
         Height          =   315
         Left            =   -74790
         TabIndex        =   44
         Top             =   2580
         Width           =   1395
      End
      Begin VB.CheckBox chkDocOrigen 
         Caption         =   " Documento"
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
         Left            =   420
         TabIndex        =   43
         Top             =   2130
         Width           =   1365
      End
      Begin VB.TextBox txtChqRecImporte 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   -68070
         TabIndex        =   20
         Tag             =   "0"
         Text            =   "0.00"
         Top             =   2565
         Width           =   1680
      End
      Begin Sicmact.FlexEdit fgChqRecibido 
         Height          =   1965
         Left            =   -74790
         TabIndex        =   19
         Top             =   540
         Width           =   8415
         _extentx        =   14843
         _extenty        =   3466
         cols0           =   12
         encabezadosnombres=   "#-Opc-Banco-NroCheque-Fecha-Importe-Cuenta-cAreaCod-cAgeCod-nMovNro-cPersCod-cIFTpo"
         encabezadosanchos=   "0-420-3200-1800-1200-1500-0-0-0-0-0-0"
         font            =   "frmCajaGenAperCtas.frx":04A8
         font            =   "frmCajaGenAperCtas.frx":04CC
         font            =   "frmCajaGenAperCtas.frx":04F0
         font            =   "frmCajaGenAperCtas.frx":0514
         font            =   "frmCajaGenAperCtas.frx":0538
         fontfixed       =   "frmCajaGenAperCtas.frx":055C
         columnasaeditar =   "X-1-X-X-X-X-X-X-X-X-X-X"
         textstylefixed  =   3
         listacontroles  =   "0-4-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-L-L-C-R-L-C-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-2-0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbbuscaduplicadotext=   -1
         rowheight0      =   300
      End
      Begin VB.TextBox txtTotalOtrasCtas 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   -69210
         TabIndex        =   24
         Tag             =   "3"
         Top             =   1545
         Width           =   1440
      End
      Begin Sicmact.FlexEdit fgOtros 
         Height          =   1005
         Left            =   -74820
         TabIndex        =   21
         Top             =   510
         Width           =   7275
         _extentx        =   12779
         _extenty        =   1667
         cols0           =   4
         highlight       =   1
         encabezadosnombres=   "#-Cuenta-Descripcion-Importe"
         encabezadosanchos=   "300-1800-3500-1400"
         font            =   "frmCajaGenAperCtas.frx":058A
         font            =   "frmCajaGenAperCtas.frx":05AE
         font            =   "frmCajaGenAperCtas.frx":05D2
         font            =   "frmCajaGenAperCtas.frx":05F6
         font            =   "frmCajaGenAperCtas.frx":061A
         fontfixed       =   "frmCajaGenAperCtas.frx":063E
         columnasaeditar =   "X-1-X-3"
         textstylefixed  =   3
         listacontroles  =   "0-1-0-0"
         encabezadosalineacion=   "C-L-L-R"
         formatosedit    =   "0-0-0-2"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbbuscaduplicadotext=   -1
         colwidth0       =   300
         rowheight0      =   300
      End
      Begin VB.CommandButton cmdAgregarCta 
         Caption         =   "A&gregar"
         Height          =   360
         Left            =   -67470
         TabIndex        =   22
         Top             =   510
         Width           =   1200
      End
      Begin VB.CommandButton cmdEliminarCta 
         Caption         =   "&Eliminar"
         Height          =   360
         Left            =   -67470
         TabIndex        =   23
         Top             =   885
         Width           =   1200
      End
      Begin VB.CommandButton cmdEfectivo 
         Caption         =   "Descomposición de Efectivo"
         Height          =   405
         Left            =   -69390
         TabIndex        =   17
         Top             =   2010
         Width           =   3045
      End
      Begin VB.TextBox txtBilleteImporte 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   -68040
         TabIndex        =   18
         Tag             =   "0"
         Top             =   2505
         Width           =   1680
      End
      Begin VB.Frame fraTransferencia 
         Caption         =   "Entidad Financiera"
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
         Height          =   1665
         Left            =   270
         TabIndex        =   33
         Top             =   420
         Width           =   8295
         Begin VB.TextBox txtBancoImporte 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   285
            Left            =   6330
            TabIndex        =   41
            Tag             =   "0"
            Top             =   1185
            Width           =   1680
         End
         Begin Sicmact.TxtBuscar txtBuscaEntidad 
            Height          =   360
            Left            =   1080
            TabIndex        =   12
            Top             =   300
            Width           =   2580
            _extentx        =   4551
            _extenty        =   635
            appearance      =   1
            appearance      =   1
            font            =   "frmCajaGenAperCtas.frx":066C
            appearance      =   1
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Importe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5190
            TabIndex        =   42
            Top             =   1230
            Width           =   615
         End
         Begin VB.Label lblDesCtaIfTransf 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1095
            TabIndex        =   14
            Top             =   720
            Width           =   6930
         End
         Begin VB.Label lblDescIfTransf 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3735
            TabIndex        =   13
            Top             =   300
            Width           =   4290
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta N° :"
            Height          =   210
            Left            =   180
            TabIndex        =   34
            Top             =   360
            Width           =   810
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   315
            Left            =   4980
            Top             =   1170
            Width           =   3045
         End
      End
      Begin VB.Frame fraDocTrans 
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
         Height          =   765
         Left            =   270
         TabIndex        =   32
         Top             =   2130
         Width           =   4560
         Begin VB.OptionButton optDoc 
            Caption         =   "Carta"
            Height          =   345
            Index           =   1
            Left            =   2310
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   270
            Width           =   2055
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Cheque"
            Height          =   345
            Index           =   0
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   270
            Width           =   2055
         End
      End
      Begin Sicmact.FlexEdit fgObj 
         Height          =   945
         Left            =   -74820
         TabIndex        =   25
         Top             =   1950
         Width           =   6195
         _extentx        =   10927
         _extenty        =   1667
         cols0           =   8
         highlight       =   2
         allowuserresizing=   1
         encabezadosnombres=   "#-Ord-Código-Descripción-CtaCont-SubCta-ObjPadre-ItemCtaCont"
         encabezadosanchos=   "350-400-1200-3000-0-900-0-0"
         font            =   "frmCajaGenAperCtas.frx":0690
         font            =   "frmCajaGenAperCtas.frx":06BC
         font            =   "frmCajaGenAperCtas.frx":06E8
         font            =   "frmCajaGenAperCtas.frx":0714
         font            =   "frmCajaGenAperCtas.frx":0740
         fontfixed       =   "frmCajaGenAperCtas.frx":076C
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X-X-X-X-X"
         textstylefixed  =   3
         listacontroles  =   "0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-L-L-C-C-C-C"
         formatosedit    =   "0-0-3-0-0-0-0-0"
         textarray0      =   "#"
         lbbuscaduplicadotext=   -1
         colwidth0       =   345
         rowheight0      =   300
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69210
         TabIndex        =   39
         Top             =   2610
         Width           =   615
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70620
         TabIndex        =   36
         Top             =   1575
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69180
         TabIndex        =   35
         Top             =   2550
         Width           =   615
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   -69390
         Top             =   2490
         Width           =   3045
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   -70830
         Top             =   1530
         Width           =   3105
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   -69420
         Top             =   2550
         Width           =   3045
      End
   End
End
Attribute VB_Name = "frmCajaGenAperCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnOpcion As Long
Dim oOpe As DOperacion
Dim oCtaIf As NCajaCtaIF
Dim lnTipoCtaIf As CGTipoCtaIF
Dim lsSubCtaIF As String
Dim lsNroCartaApert As String
Dim lsDocCartaAper As String
Dim lnTpoDocAper As TpoDoc
Dim lbCreaSubCta As Boolean
Dim lbCalendario As Boolean
Dim lnNroCuotas  As Integer
'Efectivo
Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset

'Documento de Transferencia
Dim lsDocumento As String
Dim lnTpoDoc As TpoDoc
Dim lsNroDoc As String
Dim lsNroVoucher As String
Dim ldFechaDoc  As Date
Dim lsSubCuentaIF As String  'Sub Cuenta Contable de la Entidad
Dim objPista As COMManejador.Pista 'ARLO20170217

Function Valida() As Boolean
Dim K As Integer
Valida = False
If txtImporte <> txtTotApertura Then
    MsgBox "Importe de Apertura diferente a Total de Importe retirado", vbInformation, "¡Aviso!"
    Exit Function
End If
If Len(Trim(txtBuscaIF)) = 0 Then
    MsgBox "Entidad Financiera no seleccionada", vbInformation, "Aviso"
    txtBuscaIF.SetFocus
    Exit Function
End If

If nVal(txtTotApertura) = 0 Then
    MsgBox "No se indicaron Cuentas a Aperturar. Por favor Verifique", vbInformation, "Aviso"
    fgCta.SetFocus
    Exit Function
End If

If nVal(txtImporte) = 0 Then
    MsgBox "No se indico Importe Origen para Apertura", vbInformation, "¡Aviso!"
    Exit Function
End If

If lnTipoCtaIf = gTpoCtaIFCtaPF Then
    For K = 1 To fgCta.Rows - 1
        If nVal(fgCta.TextMatrix(K, 4)) = 0 Then
            MsgBox "Plazo de Cuenta no válido ", vbInformation, "Aviso"
            fgCta.SetFocus
            fgCta.col = 4
            fgCta.row = K
            Exit Function
        End If
        If nVal(fgCta.TextMatrix(K, 5)) = 0 Then
            MsgBox "Periodo de Interés no válido ", vbInformation, "Aviso"
            fgCta.SetFocus
            fgCta.col = 5
            fgCta.row = K
            Exit Function
        End If
        If nVal(fgCta.TextMatrix(K, 5)) = 0 Then
            MsgBox "Porcentaje de Interés no válido ", vbInformation, "Aviso"
            fgCta.SetFocus
            fgCta.col = 6
            fgCta.row = K
            Exit Function
        End If
    Next
End If
If nVal(Me.txtBancoImporte) <> 0 Then
    If Len(Trim(txtBuscaEntidad)) = 0 Then
        MsgBox "Cuenta de Entidad Financiera no válida", vbInformation, "Aviso"
        txtBuscaEntidad.SetFocus
        Exit Function
    End If
End If
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Descripción de Operación no válida", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Exit Function
End If
Valida = True
End Function

Private Sub chkDocOrigen_Click()
    If chkDocOrigen.value = Checked Then
        fraDocTrans.Enabled = True
    Else
        fraDocTrans.Enabled = False
    End If
End Sub

'EJVG20120801 *** Adecuado x OverNight
Private Sub cmbTipoPlazoFijo_Click()
    Dim lnTpoPF As Integer
    lnTpoPF = CInt(Right(Me.cmbTipoPlazoFijo.Text, 1))
    If lnTpoPF = 1 Then 'PLAZO FIJO
        lblDescTipoCta = gsOpeDescHijo
        lnTipoCtaIf = gTpoCtaIFCtaPF
    ElseIf lnTpoPF = 2 Then 'OVERNIGHT
        lblDescTipoCta = "CUENTA OVERNIGHT EN " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N", "M.E")
        lnTipoCtaIf = gTpoCtaIFCtaPFOverNight
    End If
    'CtasIF
    lsSubCtaIF = ""
    lbCreaSubCta = False
    If txtBuscaIF <> "" Then
        lbCreaSubCta = Not oCtaIf.GetVerificaSubCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1))
        lsSubCtaIF = oCtaIf.GetSubCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1))
        fgCta.Rows = 2
        fgCta.Clear
        fgCta.FormaCabecera
    End If
End Sub
Private Sub cmbTipoPlazoFijo_Change()
    Dim lnTpoPF As Integer
    lnTpoPF = CInt(Right(Me.cmbTipoPlazoFijo.Text, 1))
    If lnTpoPF = 1 Then 'PLAZO FIJO
        lblDescTipoCta = gsOpeDescHijo
    ElseIf lnTpoPF = 2 Then 'OVERNIGHT
        lblDescTipoCta = "CUENTA OVERNIGHT EN " & IIf(gsOpeCod = "401514", "M.N", "M.E")
        lnTipoCtaIf = gTpoCtaIFCtaPFOverNight
    End If
End Sub
'END EJVG *******

Private Sub cmdAceptar_Click()
Dim lsCuentaAho As String

Dim lsMovNro As String
Dim oCon     As NContFunciones
Dim oCaja As nCajaGeneral
Dim rsAdeud  As ADODB.Recordset
Dim lTasaInteres() As Double 'RIRO20140530 ERS017
Dim I As Integer

On Error GoTo ErrApertura
If Valida = False Then
   Exit Sub
End If
If Len(Trim(lsNroCartaApert)) = 0 Then
    If MsgBox("Carta de Apertura no ha sido ingresada. ¿Desea Continuar con la Operación? ", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        cmdCartAper.SetFocus
        Exit Sub
    End If
End If

Set oCon = New NContFunciones
Set oCaja = New nCajaGeneral

'RIRO20140530 ERS017 *****
For I = 1 To fgCta.Rows - 1
ReDim Preserve lTasaInteres(I)
lTasaInteres(I) = fgCta.TextMatrix(I, 6)
Next
'END RIRO ****************

'guardando datos
If MsgBox(" ¿ Desea Grabar Operación de Apertura ? ", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oCon.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
    oCaja.GrabaAperturaCtaIF lsMovNro, gsOpeCod, txtFecha, txtMovDesc, nVal(txtImporte), _
            fgCta.GetRsNew, lbCreaSubCta, lsSubCtaIF, lblDescTipoCta, _
            Mid(txtBuscaIF, 4, 13), Mid(txtBuscaIF, 1, 2), rsBill, rsMon, _
            lsNroCartaApert, gdFecSis, txtBuscaEntidad, nVal(txtBancoImporte), _
            lnTpoDoc, lsNroDoc, ldFechaDoc, lsNroVoucher, _
            fgChqRecibido.GetRsNew, fgOtros.GetRsNew, fgObj.GetRsNew, IIf(Trim(Right(cmbTipoPlazoFijo.Text, 1)) = 2, 1, 0), lTasaInteres

    If lsDocCartaAper <> "" Then
        EnviaPrevio lsDocCartaAper + oImpresora.gPrnSaltoPagina + lsDocCartaAper, "Cartas de Apertura", gnLinPage, False
    End If
    ImprimeAsientoContable lsMovNro, lsNroVoucher, lnTpoDoc, lsDocumento, True, False
    
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
        Set objPista = Nothing
        '****
        
    If MsgBox("Desea Realizar otra operación de Apertura de Cuentas??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        txtBuscaEntidad = ""
        txtBuscaIF = ""
        txtImporte = "0.00"
        txtTotApertura = "0.00"
        txtBancoImporte = "0.00"
        txtBilleteImporte = "0.00"
        txtTotalOtrasCtas = "0.00"
        
        txtMovDesc = ""
        Set rsBill = Nothing
        Set rsMon = Nothing
        lblDescIF = ""
        lblDescIfTransf = ""
        lblDesCtaIfTransf = ""
        lblDescTipoCta = gsOpeDescHijo
        
        lsNroCartaApert = ""
        lsDocCartaAper = ""
        
        txtBancoImporte = "0.00"
        
        lsDocumento = ""
        lnTpoDoc = -1
        lsNroDoc = ""
        lsNroVoucher = ""
        
        fgCta.Clear
        fgCta.Rows = 2
        fgCta.FormaCabecera
        
        fgOtros.Clear
        fgOtros.Rows = 2
        fgOtros.FormaCabecera
        
        fgObj.Clear
        fgObj.Rows = 2
        fgObj.FormaCabecera
        
        If Not frmAdeudCal Is Nothing Then
            Unload frmAdeudCal
            Set frmAdeudCal = Nothing
        End If

Else
        Unload Me
    End If
End If
Exit Sub
ErrApertura:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdAgregar_Click()
Dim lsCuentaCod As String
If txtBuscaIF.Text <> "" And lsSubCtaIF <> "" Then
    If fgCta.Rows > 2 Or (fgCta.row = 1 And fgCta.TextMatrix(fgCta.row, 0) <> "") Then
        If Not ValidaDatosCuenta(CInt(fgCta.row)) Then
            fgCta.SetFocus
            Exit Sub
        End If
    End If
    If fgCta.TextMatrix(fgCta.Rows - 1, 1) <> "" Then
       lsCuentaCod = Left(fgCta.TextMatrix(fgCta.Rows - 1, 1), Len(fgCta.TextMatrix(fgCta.Rows - 1, 1)) - 2) + Format(nVal(Right(fgCta.TextMatrix(fgCta.Rows - 1, 1), 2)) + 1, "00")
    Else
        lsCuentaCod = oCtaIf.GetNewCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1), lsSubCtaIF)
    End If
    fgCta.AdicionaFila
    fgCta.TextMatrix(fgCta.Rows - 1, 1) = lsCuentaCod
    'fgCta.TextMatrix(fgCta.Rows - 1, 2) = "Pendiente"
    fgCta.TextMatrix(fgCta.Rows - 1, 2) = "Pendiente" & IIf(lnTipoCtaIf = gTpoCtaIFCtaPFOverNight, " OverNight", "") 'EJVG20120801 Para que se den cuenta que realizan OverNight
    fgCta.TextMatrix(fgCta.Rows - 1, 5) = "360"
    fgCta.col = 2
    fgCta.SetFocus
Else
    MsgBox "Aun no selecciono Institución donde Aperturar", vbInformation, "¡Aviso!"
    txtBuscaIF.SetFocus
End If
End Sub

Private Sub cmdAgregarCta_Click()
Dim oOpe As New DOperacion
fgOtros.AdicionaFila
fgOtros.rsTextBuscar = oOpe.EmiteOpeCtasNivel(gsOpeCod, , "4")
fgOtros.SetFocus
Set oOpe = Nothing
End Sub

Private Sub cmdCartAper_Click()
Dim oDoc As clsDocPago
Set oDoc = New clsDocPago
Dim lsCtaEntDest As String
Dim lsEntDest As String

If Valida = False Then Exit Sub
lsCtaEntDest = ""
lsEntDest = ""
If lblDesCtaIfTransf <> "" Then
    lsCtaEntDest = lblDesCtaIfTransf
    lsEntDest = lblDescIfTransf
End If
lsNroCartaApert = ""
lsDocCartaAper = ""
oDoc.InicioCarta "", "", gsOpeCod, gsOpeDesc, txtMovDesc, "", txtImporte, gdFecSis, _
            lblDescIF, "", lsEntDest, lsCtaEntDest, ""
If oDoc.vbOk Then
    lnTpoDocAper = TpoDocCarta
    lsNroCartaApert = oDoc.vsNroDoc
    lsDocCartaAper = oDoc.vsFormaDoc
    cmdAceptar.SetFocus
End If


End Sub

Private Sub cmdChqRecibido_Click()
Dim oDocRec As New NDocRec
Dim rs As ADODB.Recordset
Dim nRow As Integer
Set rs = New ADODB.Recordset
Set rs = oDocRec.GetChequesNoDepositados(Mid(gsOpeCod, 3, 1))
fgChqRecibido.Rows = 2
fgChqRecibido.Clear
fgChqRecibido.FormaCabecera
If Not rs.EOF And Not rs.BOF Then
    Do While Not rs.EOF
        fgChqRecibido.AdicionaFila
        nRow = fgChqRecibido.row
        fgChqRecibido.TextMatrix(nRow, 2) = rs!banco
        fgChqRecibido.TextMatrix(nRow, 3) = rs!cNroDoc
        fgChqRecibido.TextMatrix(nRow, 4) = rs!Fecha
        fgChqRecibido.TextMatrix(nRow, 5) = rs!nMonto
        fgChqRecibido.TextMatrix(nRow, 6) = rs!Objeto
        fgChqRecibido.TextMatrix(nRow, 7) = rs!cAreaCod
        fgChqRecibido.TextMatrix(nRow, 8) = rs!cAgeCod
        fgChqRecibido.TextMatrix(nRow, 9) = rs!nMovNro
        fgChqRecibido.TextMatrix(nRow, 10) = rs!cPersCod
        fgChqRecibido.TextMatrix(nRow, 11) = rs!cIFTpo
        rs.MoveNext
    Loop
End If
RSClose rs
Set oDocRec = Nothing
End Sub

Private Sub cmdefectivo_Click()

    frmCajaGenEfectivo.Inicio gsOpeCod, gsOpeDesc, 0, Mid(gsOpeCod, 3, 1), False
    If frmCajaGenEfectivo.lbOk Then
         Set rsBill = frmCajaGenEfectivo.rsBilletes
         Set rsMon = frmCajaGenEfectivo.rsMonedas
        txtBilleteImporte.Text = Format(frmCajaGenEfectivo.Total, gsFormatoNumeroView)
    Else
        Unload frmCajaGenEfectivo
        Set frmCajaGenEfectivo = Nothing
        txtBilleteImporte.Text = "0.00"
        RSClose rsBill
        RSClose rsMon
        Exit Sub
    End If
    CalculaTotalRetiros
    If rsBill Is Nothing And rsMon Is Nothing Then
        MsgBox "No se Ingreso de Billetaje", vbInformation, "Aviso"
        Exit Sub
    End If
End Sub

Private Sub cmdEliminar_Click()
fgCta.EliminaFila fgCta.row
txtTotApertura.Text = Format(fgCta.SumaRow(3), gsFormatoNumeroView)
End Sub

Private Sub cmdEliminarCta_Click()
If fgOtros.TextMatrix(fgOtros.row, 0) <> "" Then
   EliminaCuenta fgOtros.TextMatrix(fgOtros.row, 1), fgOtros.TextMatrix(fgOtros.row, 0)
   txtTotalOtrasCtas.Text = Format(fgOtros.SumaRow(3), gsFormatoNumeroView)
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Function ValidaDatosCuenta(pnRow As Integer) As Boolean
ValidaDatosCuenta = False
If fgCta.TextMatrix(pnRow, 1) = "" Then
    MsgBox "Falta indicar Codigo de Cuenta a aperturar", vbInformation, "¡AViso!"
    fgCta.col = 1
    Exit Function
End If
If fgCta.TextMatrix(pnRow, 2) = "" Then
    MsgBox "Falta indicar Descripción de Cuenta a Aperturar", vbInformation, "¡Aviso!"
    fgCta.col = 2
    Exit Function
End If
If nVal(fgCta.TextMatrix(pnRow, 3)) = 0 Then
    MsgBox "Falta indicar Importe de Apertura", vbInformation, "¡Aviso!"
    fgCta.col = 3
    Exit Function
End If
If lnTipoCtaIf = gTpoCtaIFCtaPF Then
    If nVal(fgCta.TextMatrix(pnRow, 4)) = 0 Then
        MsgBox "Falta indicar Plazo de Cuenta", vbInformation, "¡Aviso!"
        fgCta.col = 4
        Exit Function
    End If
    If nVal(fgCta.TextMatrix(pnRow, 5)) = 0 Then
        MsgBox "Falta indicar Periodo de Interes de Cuenta", vbInformation, "¡Aviso!"
        fgCta.col = 5
        Exit Function
    End If
    If nVal(fgCta.TextMatrix(pnRow, 6)) = 0 Then
        MsgBox "Falta indicar Interes de Cuenta", vbInformation, "¡Aviso!"
        fgCta.col = 6
        Exit Function
    End If
End If
ValidaDatosCuenta = True
End Function

Private Sub fgChqRecibido_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
If fgChqRecibido.TextMatrix(pnRow, 1) = "." Then
    txtChqRecImporte = nVal(txtChqRecImporte) + nVal(fgChqRecibido.TextMatrix(pnRow, 5))
Else
    txtChqRecImporte = nVal(txtChqRecImporte) - nVal(fgChqRecibido.TextMatrix(pnRow, 5))
End If
CalculaTotalRetiros
End Sub

Private Sub fgCta_OnRowChange(pnRow As Long, pnCol As Long)
txtTotApertura.Text = Format(fgCta.SumaRow(3), gsFormatoNumeroView)
End Sub

Private Sub fgCta_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
txtTotApertura.Text = Format(fgCta.SumaRow(3), gsFormatoNumeroView)
End Sub


Private Sub fgOtros_OnCellChange(pnRow As Long, pnCol As Long)
txtTotalOtrasCtas = Format(fgOtros.SumaRow(3), gsFormatoNumeroView)
CalculaTotalRetiros
End Sub

Private Sub fgOtros_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim oCta As New DGeneral
If psDataCod <> "" Then
    AsignaCtaObj psDataCod
    fgOtros.col = 2
End If
Set oCta = Nothing
End Sub

Private Sub fgOtros_RowColChange()
RefrescaFgObj Val(fgOtros.TextMatrix(fgOtros.row, 0))
End Sub

Private Sub Form_Load()
Set oOpe = New DOperacion
Set oCtaIf = New NCajaCtaIF
lbCreaSubCta = False
Me.Caption = gsOpeDesc
txtBuscaIF.psRaiz = "Instituciones Financieras"
txtBuscaIF.rs = oOpe.GetOpeObj(gsOpeCod, "1")
'txtBuscaEntidad.psRaiz = "Cuentas de Entidades Financieras"
txtBuscaEntidad.rs = oOpe.GetOpeObj(gsOpeCod, "2")
'txtCtaNro = "Pendiente"

lblDescTipoCta = gsOpeDescHijo
CentraForm Me
txtFecha = gdFecSis
lnTipoCtaIf = -1
Select Case gsOpeCod
    Case gOpeCGOpeAperCorrienteMN, gOpeCGOpeAperCorrienteME
        lnTipoCtaIf = gTpoCtaIFCtaCte
        
    Case gOpeCGOpeAperAhorroMN, gOpeCGOpeAperAhorroME, _
        gOpeCGOpeCMACAperAhorrosMN, gOpeCGOpeCMACAperAhorrosME
        lnTipoCtaIf = gTpoCtaIFCtaAho
    Case gOpeCGOpeAperPlazoMN, gOpeCGOpeAperPlazoME, _
         gOpeCGOpeCMACAperPFMN, gOpeCGOpeCMACAperPFME
        lnTipoCtaIf = gTpoCtaIFCtaPF
        '-----john ---------
        cmbTipoPlazoFijo.Visible = True
        lblTipo.Visible = True
        
        Dim rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        Dim oCons As DConstante
        Set oCons = New DConstante
        
        Set rs1 = oCons.CargaConstante(9064)
        If Not (rs1.EOF And rs1.BOF) Then
            cmbTipoPlazoFijo.Clear
            Do While Not rs1.EOF
                cmbTipoPlazoFijo.AddItem Trim(rs1(2)) & Space(100) & Trim(rs1(1))
                rs1.MoveNext
            Loop
            cmbTipoPlazoFijo.ListIndex = 0
        End If
       '-----------------------
End Select

            fgCta.ColWidth(7) = 1200
            fgCta.ColWidth(8) = 1400
    Select Case gsOpeCod
       Case gOpeCGOpeAperCorrienteMN, gOpeCGOpeAperAhorroMN, _
            gOpeCGOpeAperPlazoMN, gOpeCGOpeCMACAperAhorrosMN
            
            fgCta.ColWidth(7) = 0
            fgCta.ColWidth(8) = 0
    End Select

ldFechaDoc = gdFecSis

If gsCodCMAC = "102" Then
   fgCta.CantDecimales = 4
End If
fgCta.CantDecimales = 5
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not frmCajaGenEfectivo Is Nothing Then
    Unload frmCajaGenEfectivo
End If
Set frmCajaGenEfectivo = Nothing

End Sub

Private Sub OptDoc_Click(Index As Integer)
Dim oDocPago As clsDocPago
    Set oDocPago = New clsDocPago
    If optDoc(0).value Then
        oDocPago.InicioCheque "", True, Mid(txtBuscaEntidad, 4, 13), gsOpeCod, gsNomCmac, gsOpeDesc, txtMovDesc, _
                     CCur(txtImporte), gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lblDescIfTransf, _
                     lblDesCtaIfTransf, "", True, , Mid(Me.txtBuscaEntidad, 18, 10), , Mid(txtBuscaEntidad, 1, 2), Mid(txtBuscaEntidad, 4, 13), Mid(txtBuscaEntidad, 18, 10)
                     'lblDesCtaIfTransf , "", True, , Mid(Me.txtBuscaEntidad, 18, 10)
        If oDocPago.vbOk Then
            lsDocumento = oDocPago.vsFormaDoc
            lnTpoDoc = Val(oDocPago.vsTpoDoc)
            lsNroDoc = oDocPago.vsNroDoc
            lsNroVoucher = oDocPago.vsNroVoucher
            ldFechaDoc = oDocPago.vdFechaDoc
            optDoc(0).value = True
        Else
            optDoc(0).value = False
            Exit Sub
        End If
    Else
        oDocPago.InicioCarta "", "", gsOpeCod, gsOpeDesc, txtMovDesc, "", CCur(txtImporte), _
                     gdFecSis, lblDescIfTransf, lblDesCtaIfTransf, gsNomCmac, "", ""
        If oDocPago.vbOk Then
            lsDocumento = oDocPago.vsFormaDoc
            lnTpoDoc = Val(oDocPago.vsTpoDoc)
            lsNroDoc = oDocPago.vsNroDoc
            lsNroVoucher = oDocPago.vsNroVoucher
            ldFechaDoc = oDocPago.vdFechaDoc
            optDoc(1).value = True
        Else
            optDoc(1).value = False
            Exit Sub
        End If
    End If
    Set oDocPago = Nothing
End Sub


Private Sub txtBancoImporte_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtBancoImporte, KeyAscii, 16, 2)
If KeyAscii = 13 Then
    txtBancoImporte = Format(txtBancoImporte, gsFormatoNumeroView)
    CalculaTotalRetiros
    chkDocOrigen.SetFocus
End If
End Sub

Private Sub txtBancoImporte_LostFocus()
    txtBancoImporte = Format(txtBancoImporte, gsFormatoNumeroView)
    CalculaTotalRetiros
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
lblDescIfTransf = oCtaIf.NombreIF(Mid(txtBuscaEntidad.Text, 4, 13))
lblDesCtaIfTransf = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))) & " " & txtBuscaEntidad.psDescripcion
lsSubCuentaIF = oCtaIf.SubCuentaIF(Mid(txtBuscaEntidad.Text, 4, 13))

txtBancoImporte.SetFocus
End Sub
Private Sub txtBuscaIF_EmiteDatos()
lblDescIF = txtBuscaIF.psDescripcion
cmdAgregar.SetFocus
lsSubCtaIF = ""
lbCreaSubCta = False
If txtBuscaIF <> "" Then
    lbCreaSubCta = Not oCtaIf.GetVerificaSubCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1))
    lsSubCtaIF = oCtaIf.GetSubCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1))
    fgCta.Rows = 2
    fgCta.Clear
    fgCta.FormaCabecera
End If
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBuscaIF.SetFocus
End If
End Sub
Private Sub txtFecha_LostFocus()
If ValFecha(txtFecha) = False Then Exit Sub
'txtfechaiNT = DateAdd("d", spnPlazo.Valor, CDate(txtFecha))
End Sub
Private Sub txtFecha_Validate(Cancel As Boolean)
If ValFecha(txtFecha) = False Then Cancel = True
'txtfechaiNT = DateAdd("d", spnPlazo.Valor, CDate(txtFecha))
End Sub
Private Sub txtImporte_GotFocus()
fEnfoque txtImporte
End Sub
Private Sub txtImporte_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtImporte, KeyAscii, 18, 2)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub

Private Sub txtImporte_LostFocus()
If Val(txtImporte) = 0 Then txtImporte = 0
txtImporte = Format(txtImporte, "#,#0.00")
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    TabDoc.SetFocus
End If
End Sub

Private Sub CalculaTotalRetiros()
txtImporte = Format(nVal(txtBancoImporte) + nVal(txtBilleteImporte) + nVal(txtChqRecImporte) + nVal(txtTotalOtrasCtas), gsFormatoNumeroView)
End Sub

Private Sub AsignaCtaObj(ByVal psCtaContCod As String)
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim lsRaiz As String
Dim oDescObj As ClassDescObjeto
Dim UP As UPersona
Dim lsFiltro As String
Dim oRHAreas As DActualizaDatosArea
Dim oCtaCont As DCtaCont
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo
Dim oContFunct As NContFunciones

Set oEfect = New Defectivo
Set oCtaIf = New NCajaCtaIF
Set oRHAreas = New DActualizaDatosArea
Set oDescObj = New ClassDescObjeto
Set oCtaCont = New DCtaCont
Set oContFunct = New NContFunciones

Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
EliminaFgObj Val(fgOtros.TextMatrix(fgOtros.row, 0))
Set rs1 = oCtaCont.CargaCtaObj(psCtaContCod, , True)
If Not rs1.EOF And Not rs1.BOF Then
    Do While Not rs1.EOF
        lsRaiz = ""
        lsFiltro = ""
        Select Case Val(rs1!cObjetoCod)
            Case ObjCMACAgencias
                Set rs = oRHAreas.GetAgencias(rs1!cCtaObjFiltro)
            Case ObjCMACAgenciaArea
                lsRaiz = "Unidades Organizacionales"
                Set rs = oRHAreas.GetAgenciasAreas(rs1!cCtaObjFiltro)
            Case ObjCMACArea
                Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
            Case ObjEntidadesFinancieras
                lsRaiz = "Cuentas de Entidades Financieras"
                Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, psCtaContCod)
            Case ObjDescomEfectivo
                Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
            Case ObjPersona
                Set rs = Nothing
            Case Else
                lsRaiz = "Varios"
                Set rs = GetObjetos(Val(rs1!cObjetoCod))
        End Select
        If Not rs Is Nothing Then
            If rs.State = adStateOpen Then
                If Not rs.EOF And Not rs.BOF Then
                    If rs.RecordCount > 1 Then
                        oDescObj.Show rs, "", lsRaiz
                        If oDescObj.lbOk Then
                            lsFiltro = oContFunct.GetFiltroObjetos(Val(rs1!cObjetoCod), psCtaContCod, oDescObj.gsSelecCod, False)
                            AdicionaObj psCtaContCod, fgOtros.TextMatrix(fgOtros.row, 0), rs1!nCtaObjOrden, oDescObj.gsSelecCod, _
                                        oDescObj.gsSelecDesc, lsFiltro, rs1!cObjetoCod
                        Else
                            fgOtros.EliminaFila fgOtros.row, False
                            Exit Do
                        End If
                    Else
                        AdicionaObj psCtaContCod, fgOtros.TextMatrix(fgOtros.row, 0), rs1!nCtaObjOrden, rs1!cObjetoCod, _
                                        rs1!cObjetoDesc, lsFiltro, rs1!cObjetoCod
                    End If
                End If
            End If
        Else
            If Val(rs1!cObjetoCod) = ObjPersona Then
                Set UP = frmBuscaPersona.Inicio
                If Not UP Is Nothing Then
                    AdicionaObj psCtaContCod, fgOtros.TextMatrix(fgOtros.row, 0), rs1!nCtaObjOrden, _
                                    UP.sPersCod, UP.sPersNombre, _
                                    lsFiltro, rs1!cObjetoCod
                End If
                Set frmBuscaPersona = Nothing
            End If
        End If
        rs1.MoveNext
    Loop
End If
rs1.Close
Set rs1 = Nothing
Set oDescObj = Nothing
Set UP = Nothing
Set oCtaCont = Nothing
Set oCtaIf = Nothing
Set oEfect = Nothing
Set oContFunct = Nothing

End Sub
Private Sub AdicionaObj(sCodCta As String, nFila As Integer, _
                        psOrden As String, psObjetoCod As String, psObjDescripcion As String, _
                        psSubCta As String, psObjPadre As String)
Dim nItem As Integer
    fgObj.AdicionaFila
    nItem = fgObj.row
    fgObj.TextMatrix(nItem, 0) = nFila
    fgObj.TextMatrix(nItem, 1) = psOrden
    fgObj.TextMatrix(nItem, 2) = psObjetoCod
    fgObj.TextMatrix(nItem, 3) = psObjDescripcion
    fgObj.TextMatrix(nItem, 4) = sCodCta
    fgObj.TextMatrix(nItem, 5) = psSubCta
    fgObj.TextMatrix(nItem, 6) = psObjPadre
    fgObj.TextMatrix(nItem, 7) = nFila
    'fgOtros.TextMatrix(fgOtros.Row, 6) = psObjetoCod
    
End Sub

Private Sub EliminaCuenta(sCod As String, nItem As Integer)
If fgOtros.TextMatrix(1, 0) <> "" Then
    EliminaFgObj Val(fgOtros.TextMatrix(fgOtros.row, 0))
    fgOtros.EliminaFila fgOtros.row, False
End If
If Len(fgOtros.TextMatrix(1, 1)) > 0 Then
   RefrescaFgObj Val(fgOtros.TextMatrix(fgOtros.row, 0))
End If
End Sub

Private Sub EliminaFgObj(nItem As Integer)
Dim K  As Integer, m As Integer
K = 1
Do While K < fgObj.Rows
   If Len(fgObj.TextMatrix(K, 1)) > 0 Then
      If Val(fgObj.TextMatrix(K, 0)) = nItem Then
         fgObj.EliminaFila K, False
      Else
         K = K + 1
      End If
   Else
      K = K + 1
   End If
Loop
End Sub

Private Sub RefrescaFgObj(nItem As Integer)
Dim K  As Integer
For K = 1 To fgObj.Rows - 1
    If Len(fgObj.TextMatrix(K, 1)) Then
       If fgObj.TextMatrix(K, 0) = nItem Then
          fgObj.RowHeight(K) = 285
       Else
          fgObj.RowHeight(K) = 0
       End If
    End If
Next
End Sub

