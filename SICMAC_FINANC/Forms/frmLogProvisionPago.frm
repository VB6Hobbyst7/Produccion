VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogProvisionPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Provisión pago a Proveedores"
   ClientHeight    =   7650
   ClientLeft      =   450
   ClientTop       =   1740
   ClientWidth     =   10635
   Icon            =   "frmLogProvisionPago.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraRetencSistPens 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7800
      TabIndex        =   87
      Top             =   5900
      Visible         =   0   'False
      Width           =   2835
      Begin VB.CommandButton cmdRetSistPensActualizar 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   2520
         Picture         =   "frmLogProvisionPago.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Actualizar retención SNP/ONP"
         Top             =   135
         Width           =   280
      End
      Begin VB.TextBox txtRetSistPens 
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
         Left            =   1290
         TabIndex        =   89
         Text            =   "0.00"
         Top             =   150
         Width           =   1185
      End
      Begin VB.CommandButton cmdRetSistPensDetalle 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   990
         Picture         =   "frmLogProvisionPago.frx":06E7
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Ver retención SNP/ONP"
         Top             =   135
         Width           =   280
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "SNP/ONP :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   90
         Top             =   195
         Width           =   885
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   0
         Top             =   120
         Width           =   2505
      End
   End
   Begin VB.Frame fraComprobante 
      Caption         =   "Búscar"
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
      Height          =   735
      Left            =   45
      TabIndex        =   82
      Top             =   0
      Visible         =   0   'False
      Width           =   4350
      Begin VB.CommandButton cmdComprobanteLimpiar 
         Caption         =   "&Limpiar"
         Height          =   330
         Left            =   3400
         TabIndex        =   84
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdComprobanteCargar 
         Caption         =   "&Cargar"
         Height          =   330
         Left            =   2550
         TabIndex        =   83
         Top             =   240
         Width           =   855
      End
      Begin Sicmact.TxtBuscar txtComprobanteCod 
         Height          =   285
         Left            =   620
         TabIndex        =   85
         Top             =   260
         Width           =   1900
         _extentx        =   3360
         _extenty        =   503
         appearance      =   0
         appearance      =   0
         font            =   "frmLogProvisionPago.frx":0AA0
         appearance      =   0
         tipobusqueda    =   6
         stitulo         =   ""
         enabledtext     =   0
      End
      Begin VB.Label Label23 
         Caption         =   "Doc. Origen:"
         Height          =   495
         Left            =   80
         TabIndex        =   86
         Top             =   165
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Distribución"
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
      Height          =   735
      Left            =   9000
      TabIndex        =   80
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton cmdDistribución 
         Caption         =   "Distribución"
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
         Left            =   120
         TabIndex        =   81
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRetenciones 
      Caption         =   "Aplicar &Penalidad"
      Height          =   330
      Left            =   4560
      TabIndex        =   79
      Top             =   7280
      Width           =   2220
   End
   Begin VB.CommandButton cmdFielCumplimiento 
      Caption         =   "Aplicar &Fiel Cump."
      Height          =   330
      Left            =   4560
      TabIndex        =   72
      ToolTipText     =   "Adiciona Importe de Ajuste de Documento"
      Top             =   6855
      Width           =   2220
   End
   Begin VB.CommandButton cmdAjuste 
      Caption         =   "Asignar A&juste   >>>"
      Height          =   330
      Left            =   2310
      TabIndex        =   59
      ToolTipText     =   "Adiciona Importe de Ajuste de Documento"
      Top             =   6435
      Width           =   2220
   End
   Begin VB.CommandButton cmdDetraccion 
      Caption         =   "Aplicar &Detracción"
      Height          =   330
      Left            =   4560
      TabIndex        =   71
      ToolTipText     =   "Adiciona Importe de Ajuste de Documento"
      Top             =   6435
      Width           =   2220
   End
   Begin VB.Frame fraAjuste 
      Height          =   615
      Left            =   2640
      TabIndex        =   63
      Top             =   6165
      Visible         =   0   'False
      Width           =   4035
      Begin VB.TextBox txtAjuste 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   66
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         TabIndex        =   65
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2760
         TabIndex        =   64
         Top             =   180
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdValVenta 
      Caption         =   "&Deduce Impuesto"
      Height          =   330
      Left            =   60
      TabIndex        =   62
      ToolTipText     =   "Calcula el Valor Venta de Subtotales"
      Top             =   6435
      Width           =   2220
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "A&gregar"
      Height          =   330
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   6075
      Width           =   1110
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   330
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   6075
      Width           =   1080
   End
   Begin VB.CommandButton cmdExaminar 
      Caption         =   "..."
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   3030
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtObj 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   295
      Left            =   1050
      TabIndex        =   56
      Top             =   3030
      Visible         =   0   'False
      Width           =   1510
   End
   Begin VB.TextBox txtCant 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   295
      Left            =   3600
      TabIndex        =   55
      Top             =   3600
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Frame fraProvis 
      Caption         =   "Provisión de ..."
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
      Height          =   660
      Left            =   60
      TabIndex        =   51
      Top             =   6825
      Width           =   3345
      Begin VB.ComboBox cboProvis 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   225
         Width           =   3075
      End
   End
   Begin VB.Frame FrameTipCambio 
      Caption         =   "Tipo de Cambio"
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
      Height          =   735
      Left            =   6805
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   2145
      Begin VB.TextBox txtTipVariable 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1040
         TabIndex        =   48
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txtTipFijo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   110
         TabIndex        =   47
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Mercado"
         Height          =   255
         Left            =   1140
         TabIndex        =   50
         Top             =   175
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Fijo"
         Height          =   255
         Left            =   420
         TabIndex        =   49
         Top             =   175
         Width           =   345
      End
   End
   Begin VB.TextBox txtTotal 
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
      Left            =   9090
      TabIndex        =   43
      Top             =   6810
      Width           =   1185
   End
   Begin VB.TextBox txtSTotal 
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
      Left            =   9090
      TabIndex        =   42
      Top             =   6480
      Width           =   1185
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   9105
      TabIndex        =   13
      Top             =   7220
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      Caption         =   "Movimiento"
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
      Height          =   735
      Left            =   4440
      TabIndex        =   8
      Top             =   0
      Width           =   2325
      Begin VB.TextBox txtOpeCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   150
         TabIndex        =   9
         Top             =   360
         Width           =   840
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   285
         Left            =   1065
         TabIndex        =   10
         Top             =   360
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   1440
         TabIndex        =   12
         Top             =   180
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "Número"
         Height          =   255
         Left            =   270
         TabIndex        =   11
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Frame frameDestino 
      Caption         =   "Proveedor "
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
      Height          =   975
      Left            =   45
      TabIndex        =   6
      Top             =   750
      Width           =   5565
      Begin VB.CheckBox chkBuneCOnt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Buen Contribuyente"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3630
         TabIndex        =   68
         Top             =   660
         Width           =   1815
      End
      Begin VB.CheckBox chkRetencion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Agente de Retención"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   67
         Top             =   660
         Width           =   1935
      End
      Begin VB.TextBox lblProvNombre 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   1770
         TabIndex        =   7
         Tag             =   "txtnombre"
         Top             =   270
         Width           =   3675
      End
      Begin Sicmact.TxtBuscar txtBuscarProv 
         Height          =   360
         Left            =   120
         TabIndex        =   58
         Top             =   270
         Width           =   1605
         _extentx        =   2831
         _extenty        =   635
         appearance      =   1
         appearance      =   1
         font            =   "frmLogProvisionPago.frx":0ACC
         appearance      =   1
         tipobusqueda    =   3
         tipobuspers     =   2
      End
   End
   Begin VB.Frame frameMovDesc 
      Caption         =   "Observaciones"
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
      Height          =   885
      Left            =   45
      TabIndex        =   4
      Top             =   1740
      Width           =   5565
      Begin VB.TextBox txtMovDesc 
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   150
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   210
         Width           =   5295
      End
   End
   Begin VB.PictureBox PicOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6900
      Picture         =   "frmLogProvisionPago.frx":0AF8
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   6555
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   7815
      TabIndex        =   2
      Top             =   7220
      Width           =   1230
   End
   Begin VB.PictureBox picCuadroNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6630
      Picture         =   "frmLogProvisionPago.frx":0E3A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   6525
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picCuadroSi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6330
      Picture         =   "frmLogProvisionPago.frx":117C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   6525
      Visible         =   0   'False
      Width           =   255
   End
   Begin TabDlg.SSTab TabDocRef 
      Height          =   1845
      Left            =   5760
      TabIndex        =   14
      Top             =   830
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   3254
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Ord. de Com&pra"
      TabPicture(0)   =   "frmLogProvisionPago.frx":14BE
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtOCPlazo"
      Tab(0).Control(1)=   "txtOCFecha"
      Tab(0).Control(2)=   "txtOCNro"
      Tab(0).Control(3)=   "txtOCEntrega"
      Tab(0).Control(4)=   "Shape4"
      Tab(0).Control(5)=   "Shape1"
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(9)=   "Label19"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Guía de &Rem."
      TabPicture(1)   =   "frmLogProvisionPago.frx":14DA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtGRNro"
      Tab(1).Control(1)=   "txtGRSerie"
      Tab(1).Control(2)=   "txtGRFecha"
      Tab(1).Control(3)=   "Shape3"
      Tab(1).Control(4)=   "Shape2"
      Tab(1).Control(5)=   "Label11"
      Tab(1).Control(6)=   "Label2"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "&Comprobante     "
      TabPicture(2)   =   "frmLogProvisionPago.frx":14F6
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Shape5"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Shape6"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label18"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtFacFecha"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cboDoc"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtFacSerie"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtFacNro"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cboDocDestino"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "&Ref."
      TabPicture(3)   =   "frmLogProvisionPago.frx":1512
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtNroRef"
      Tab(3).Control(1)=   "lblProvNombreRef"
      Tab(3).Control(2)=   "txtBuscarProvRef"
      Tab(3).Control(3)=   "txtFechaRef"
      Tab(3).Control(4)=   "Label20"
      Tab(3).Control(5)=   "Label21"
      Tab(3).Control(6)=   "Shape8"
      Tab(3).ControlCount=   7
      Begin VB.TextBox txtNroRef 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73080
         MaxLength       =   12
         TabIndex        =   78
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox lblProvNombreRef 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   -73080
         TabIndex        =   73
         Tag             =   "txtnombre"
         Top             =   840
         Width           =   2715
      End
      Begin VB.ComboBox cboDocDestino 
         Height          =   315
         ItemData        =   "frmLogProvisionPago.frx":152E
         Left            =   1095
         List            =   "frmLogProvisionPago.frx":153E
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtGRNro 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74130
         MaxLength       =   12
         TabIndex        =   23
         Top             =   840
         Width           =   1350
      End
      Begin VB.TextBox txtGRSerie 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74505
         MaxLength       =   4
         TabIndex        =   22
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtOCPlazo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74235
         TabIndex        =   21
         Top             =   1080
         Width           =   1260
      End
      Begin VB.TextBox txtFacNro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         MaxLength       =   12
         TabIndex        =   20
         Top             =   930
         Width           =   1590
      End
      Begin VB.TextBox txtFacSerie 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   585
         MaxLength       =   4
         TabIndex        =   19
         Top             =   930
         Width           =   495
      End
      Begin VB.TextBox txtOCFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -71685
         TabIndex        =   18
         Top             =   585
         Width           =   1200
      End
      Begin VB.TextBox txtOCNro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74460
         MaxLength       =   13
         TabIndex        =   17
         Top             =   585
         Width           =   1485
      End
      Begin VB.ComboBox cboDoc 
         Height          =   315
         Left            =   615
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   540
         Width           =   3900
      End
      Begin VB.TextBox txtOCEntrega 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72270
         TabIndex        =   15
         Top             =   1065
         Width           =   1785
      End
      Begin MSMask.MaskEdBox txtFacFecha 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   3375
         TabIndex        =   24
         Top             =   930
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtGRFecha 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   -71565
         TabIndex        =   25
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin Sicmact.TxtBuscar txtBuscarProvRef 
         Height          =   360
         Left            =   -74760
         TabIndex        =   74
         Top             =   840
         Width           =   1605
         _extentx        =   2831
         _extenty        =   635
         appearance      =   1
         appearance      =   1
         font            =   "frmLogProvisionPago.frx":15B3
         appearance      =   1
         tipobusqueda    =   3
         tipobuspers     =   2
      End
      Begin MSMask.MaskEdBox txtFechaRef 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   -71480
         TabIndex        =   75
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Documento Ref :"
         Height          =   195
         Left            =   -74640
         TabIndex        =   77
         Top             =   1395
         Width           =   1215
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Persona Referencia"
         Height          =   195
         Left            =   -74760
         TabIndex        =   76
         Top             =   600
         Width           =   1410
      End
      Begin VB.Shape Shape8 
         Height          =   1215
         Left            =   -74880
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label18 
         Caption         =   "Destino"
         Height          =   255
         Left            =   375
         TabIndex        =   70
         Top             =   1365
         Width           =   645
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H8000000E&
         Height          =   1290
         Left            =   135
         Top             =   465
         Width           =   4605
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H8000000C&
         Height          =   1275
         Left            =   150
         Top             =   450
         Width           =   4605
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H8000000E&
         Height          =   1245
         Left            =   -74865
         Top             =   465
         Width           =   4620
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000C&
         Height          =   1230
         Left            =   -74880
         Top             =   450
         Width           =   4650
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H8000000E&
         Height          =   1290
         Left            =   -74850
         Top             =   450
         Width           =   4635
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000C&
         Height          =   1290
         Left            =   -74865
         Top             =   435
         Width           =   4590
      End
      Begin VB.Label Label11 
         Caption         =   "Nº"
         Height          =   165
         Left            =   -74790
         TabIndex        =   34
         Top             =   915
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Emisión"
         Height          =   165
         Left            =   -72660
         TabIndex        =   33
         Top             =   915
         Width           =   1035
      End
      Begin VB.Label Label14 
         Caption         =   "Plazo"
         Height          =   240
         Left            =   -74790
         TabIndex        =   32
         Top             =   1095
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Emisión"
         Height          =   165
         Left            =   -72870
         TabIndex        =   31
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label7 
         Caption         =   "Nº"
         Height          =   165
         Left            =   -74745
         TabIndex        =   30
         Top             =   660
         Width           =   315
      End
      Begin VB.Label Label6 
         Caption         =   "Emisión"
         Height          =   165
         Left            =   2775
         TabIndex        =   29
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "Nº"
         Height          =   165
         Left            =   225
         TabIndex        =   28
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo"
         Height          =   240
         Left            =   195
         TabIndex        =   27
         Top             =   570
         Width           =   360
      End
      Begin VB.Label Label19 
         Caption         =   "Entrega"
         Height          =   240
         Left            =   -72870
         TabIndex        =   26
         Top             =   1095
         Width           =   1245
      End
   End
   Begin VB.TextBox txtProvCod 
      Height          =   315
      Left            =   195
      MaxLength       =   20
      TabIndex        =   35
      Tag             =   "txtcodigo"
      Top             =   1020
      Visible         =   0   'False
      Width           =   555
   End
   Begin Sicmact.FlexEdit fgImp 
      Height          =   1710
      Left            =   7080
      TabIndex        =   36
      Top             =   4275
      Width           =   3615
      _extentx        =   6165
      _extenty        =   2170
      cols0           =   12
      highlight       =   2
      allowuserresizing=   3
      encabezadosnombres=   "-#1-Ok-Impuesto-Tasa-Monto-CtaCont-CtaContDesc-cDocImpDH-cImpDestino-cDocImpOpc-calculo"
      encabezadosanchos=   "0-0-350-1000-600-1200-0-0-0-0-0-0"
      font            =   "frmLogProvisionPago.frx":15DF
      font            =   "frmLogProvisionPago.frx":160B
      font            =   "frmLogProvisionPago.frx":1637
      font            =   "frmLogProvisionPago.frx":1663
      font            =   "frmLogProvisionPago.frx":168F
      fontfixed       =   "frmLogProvisionPago.frx":16BB
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-2-X-X-X-X-X-X-X-X-X"
      textstylefixed  =   4
      listacontroles  =   "0-0-4-0-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-C-L-C-R-R-C-C-C-C-L-L"
      formatosedit    =   "0-0-0-2-2-2-2-2-2-2-0-0"
      lbeditarflex    =   -1
      lbformatocol    =   -1
      lbbuscaduplicadotext=   -1
      rowheight0      =   300
   End
   Begin Sicmact.FlexEdit fgObj 
      Height          =   1710
      Left            =   30
      TabIndex        =   41
      Top             =   4275
      Width           =   5955
      _extentx        =   10504
      _extenty        =   2196
      cols0           =   8
      highlight       =   2
      allowuserresizing=   1
      encabezadosnombres=   "#-Ord-Código-Descripción-CtaCont-SubCta-ObjPadre-ItemCtaCont"
      encabezadosanchos=   "350-400-1200-3000-0-900-0-0"
      font            =   "frmLogProvisionPago.frx":16E9
      font            =   "frmLogProvisionPago.frx":1715
      font            =   "frmLogProvisionPago.frx":1741
      font            =   "frmLogProvisionPago.frx":176D
      font            =   "frmLogProvisionPago.frx":1799
      fontfixed       =   "frmLogProvisionPago.frx":17C5
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
   Begin MSMask.MaskEdBox txtFecPlazo 
      Height          =   315
      Left            =   7560
      TabIndex        =   53
      Top             =   3090
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BackColor       =   15794175
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDetalle 
      Height          =   1485
      Left            =   30
      TabIndex        =   57
      Top             =   2730
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   2619
      _Version        =   393216
      Cols            =   13
      ForeColorSel    =   -2147483643
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483637
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      FocusRect       =   0
      HighLight       =   2
      RowSizingMode   =   1
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   13
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8100
      TabIndex        =   45
      Top             =   6840
      Width           =   675
   End
   Begin VB.Label lblSTot 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "SubTotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7980
      TabIndex        =   44
      Top             =   6525
      Width           =   885
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Impuestos"
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
      Height          =   225
      Left            =   6090
      TabIndex        =   39
      Top             =   4740
      Width           =   885
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "y/o"
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
      Left            =   6375
      TabIndex        =   38
      Top             =   5025
      Width           =   315
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Retenc."
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
      Left            =   6195
      TabIndex        =   37
      Top             =   5325
      Width           =   705
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Height          =   1710
      Left            =   6000
      TabIndex        =   40
      Top             =   4275
      Width           =   1035
   End
   Begin VB.Shape ShapeIGV 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   345
      Left            =   7800
      Top             =   6450
      Width           =   2505
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   345
      Left            =   7800
      Top             =   6780
      Width           =   2505
   End
End
Attribute VB_Name = "frmLogProvisionPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String, sSqlObj As String
Dim lTransActiva As Boolean      ' Controla si la transaccion esta activa o no
Dim rs As New ADODB.Recordset    'Rs temporal para lectura de datos
Dim lSalir As Boolean
Dim lLlenaObj As Boolean, OK As Boolean, lbBienes As Boolean
Dim sObjCod      As String, sObjDesc As String, sObjUnid
Dim sCtaCod      As String, sCtaDesc As String
Dim sProvCod     As String
Dim nTasaIGV     As Currency, nVariaIGV As Currency
Dim lbNewProv    As Boolean
Dim sDocDesc     As String
Dim sCtaProvis   As String
Dim lnColorBien  As Double
Dim lnColorServ  As Double
Dim lnMovNro     As Long
Dim lbTieneIGV   As Boolean
Dim lbMismaOpe   As Boolean
'ALPA 20090303**************************
Dim lbRetenciones   As Boolean
Dim lnPenalidad As Integer
'***************************************
'*
Dim lPendiente   As Boolean
Dim lPendPasivo  As Boolean
Dim lbRegulaPend As Boolean
Dim oMov As DMov

Dim oContFunct As NContFunciones
Dim lnImporteRegula  As Currency
Dim lsCtaDetraccion  As String
Dim lnTasaDetraccion As Currency
Dim rsDetracc        As ADODB.Recordset
Dim ctaEmbargo As String
Dim ctaFielCump As String
Dim ctaRetenciones As String

Dim lnPosCosto As Integer
Dim rsCosto() As New ADODB.Recordset
Dim lbHayPersonasCosto As Boolean
Dim cArchPersonas As String
'ALPA 20090302*************************************
Dim nTipoRetenciones As Integer
Dim nTFielCumplimiento As Integer
Dim nTPenalidad As Integer
'**************************************************
'ALPA 20090302*************************************
Dim sMatrizDatos() As Variant
Dim nTipoMatriz As Integer
Dim nCantAgeSel As Integer
Dim nMontogasto As Currency
Dim nMontoInafecto As Currency
Dim lnCalBaseImpIGV As Currency '***PEAC 20101104
'**************************************************
Dim rsDocValidaRuc As ADODB.Recordset
Dim cSubCta As String
Dim fnOK As Boolean
'EJVG20131113 ***
Dim fRsComprobante As ADODB.Recordset
Dim fsDocOrigenNro As String
Dim fnComprobanteMovNro As Long
'END EJVG *******
'EJVG20140724 ***
Dim fnRetProvSPMontoBase As Currency
Dim fnRetProvSPAporte As Currency
Dim fnRetProvSPComisionAFP As Currency
Dim fnRetProvSPSeguroAFP As Currency
Dim fsCtaContAporteAFP As String, fsCtaContSeguroAFP As String, fsCtaContComisionAFP As String
Dim fsCtaContAporteONP As String
'END EJVG *******

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(LlenaObj As Boolean, OrdenCompra As Boolean, Optional ProvCod As String, Optional plPendiente As Boolean = False, Optional plPendPasivo As Boolean = True, Optional pbRegulaPend As Boolean = False, Optional pbMismaOpe As Boolean = False, Optional pbRetenciones As Boolean = False)
lLlenaObj = LlenaObj
lbBienes = OrdenCompra
sProvCod = ProvCod

'Pendientes
lPendiente = plPendiente
lPendPasivo = plPendPasivo
lbRegulaPend = pbRegulaPend
lbMismaOpe = pbMismaOpe
lbRetenciones = pbRetenciones
nCantAgeSel = 0
nMontogasto = 0
Me.Show 1
End Sub

Private Sub ControlesTabDocRef(lGR As Boolean, lFac As Boolean)
txtGRSerie.Enabled = lGR
txtGRNro.Enabled = lGR
txtGRFecha.Enabled = lGR
txtFacSerie.Enabled = lFac
txtFacNro.Enabled = lFac
txtFacFecha.Enabled = lFac
cboDoc.Enabled = lFac
End Sub

Private Sub FormatoOrden()
fgDetalle.TextMatrix(0, 0) = "#"
fgDetalle.TextMatrix(0, 1) = "Código"
fgDetalle.TextMatrix(0, 2) = "Descripción"
fgDetalle.TextMatrix(0, 3) = "Unidad"
fgDetalle.TextMatrix(0, 4) = "Solicitado"
fgDetalle.TextMatrix(0, 5) = "P.Unitario"
fgDetalle.TextMatrix(0, 6) = "Saldo"
fgDetalle.TextMatrix(0, 7) = "Sub Total"
fgDetalle.TextMatrix(0, 8) = "Cod.Bien"
fgDetalle.TextMatrix(0, 9) = "Gra."
fgDetalle.TextMatrix(0, 10) = "BienServicio"
fgDetalle.TextMatrix(0, 11) = "Cant.Orden"
fgDetalle.TextMatrix(0, 12) = "Monto Orden"

fgDetalle.ColWidth(0) = 350
fgDetalle.ColWidth(1) = 1500
fgDetalle.ColWidth(2) = 4800
fgDetalle.ColWidth(3) = 0
fgDetalle.ColWidth(4) = 0
fgDetalle.ColWidth(5) = 0
fgDetalle.ColWidth(6) = 0
fgDetalle.ColWidth(7) = 1300
fgDetalle.ColWidth(8) = 0
fgDetalle.ColWidth(9) = 0
fgDetalle.ColWidth(10) = 0
fgDetalle.ColWidth(11) = 0
fgDetalle.ColWidth(12) = 0

fgDetalle.ColAlignment(1) = 1
fgDetalle.ColAlignment(8) = 1
fgDetalle.ColAlignmentFixed(0) = 4
fgDetalle.ColAlignmentFixed(4) = 7
fgDetalle.ColAlignmentFixed(5) = 7
fgDetalle.ColAlignmentFixed(6) = 7
fgDetalle.ColAlignmentFixed(7) = 7
fgDetalle.RowHeight(-1) = 285
End Sub

Private Sub cboDoc_Click()
Dim rs As ADODB.Recordset
Dim oDoc As DDocumento
Dim oConst As NConstSistemas
Set oConst = New NConstSistemas
Set oDoc = New DDocumento

Dim nRow As Integer
   fgDetalle.Cols = 13
   fgImp.Clear
   fgImp.FormaCabecera
   fgImp.Rows = 2
   Set rs = New ADODB.Recordset
   Set rs = oDoc.CargaDocImpuesto(Trim(Val(Mid(cboDoc.Text, 1, 3))))
   
   lbTieneIGV = False
    
   Do While Not rs.EOF
      'Primero adicionamos Columna de Impuesto
        fgDetalle.Cols = fgDetalle.Cols + 1
        fgDetalle.ColWidth(fgDetalle.Cols - 1) = 1100
        fgDetalle.TextMatrix(0, fgDetalle.Cols - 1) = rs!cImpAbrev
                
        '*** PEAC 20101103 - SI ES FACTURA AGREGAMIOS UNA COLUMNA MAS PARA VER SI ESTA AFECTO
'        If nCantAgeSel > 0 And Trim(Val(Mid(cboDoc.Text, 1, 3))) = "1" Then
'            fgDetalle.Cols = fgDetalle.Cols + 1
'            fgDetalle.ColWidth(fgDetalle.Cols - 1) = 0
'            fgDetalle.TextMatrix(0, fgDetalle.Cols - 1) = "AfectoInafectoAIGV"
'        End If
        '*** FIN PEAC
        
       'Adicionamos los impuestos en el grid de impuestos
        fgImp.AdicionaFila
        fgImp.Col = 0
        nRow = fgImp.row
        fgImp.TextMatrix(nRow, 1) = fgImp.row
        If rs!cDocImpOpc = "1" Then
           fgImp.TextMatrix(nRow, 2) = "1"
        End If
        fgImp.TextMatrix(nRow, 3) = rs!cImpAbrev
        fgImp.TextMatrix(nRow, 4) = Format(rs!nImpTasa, gsFormatoNumeroView)
        fgImp.TextMatrix(nRow, 5) = Format(0, gsFormatoNumeroView)
        fgImp.TextMatrix(nRow, 6) = rs!cCtaContCod
        fgImp.TextMatrix(nRow, 7) = rs!cCtaContDesc
        fgImp.TextMatrix(nRow, 8) = rs!cDocImpDH
        fgImp.TextMatrix(nRow, 9) = rs!cImpDestino
        fgImp.TextMatrix(nRow, 10) = rs!cDocImpOpc
        fgImp.TextMatrix(nRow, 11) = rs!nCalculo
        If rs!cCtaContCod = gcCtaIGV Then
            lbTieneIGV = True
            nTasaIGV = rs!nImpTasa
        End If
        
        rs.MoveNext
   Loop
   
   fgImp.Col = 1
   
   If lbTieneIGV = False Then
      cboDocDestino.ListIndex = -1
      cboDocDestino.Enabled = False
   Else
      cboDocDestino.Enabled = True
    
      Me.cboDocDestino.ListIndex = oConst.LeeConstSistema(gConstSistDestinoIGVDefecto)
      
      '-----------John-----------------------------
'      Dim nDocNro As Integer
'      nDocNro = Trim(Val(Mid(cboDoc.Text, 1, 3)))
'      If nDocNro >= 32 And nDocNro <= 170 Then
'        Me.cboDocDestino.Enabled = True
'      End If
      '--------------------------------------------
      
   End If
   
   VerReciboEgreso
    'EJVG20140727 ***
    If gbBitRetencSistPensProv Then
        If Val(Trim(Left(cboDoc.Text, 3))) = TpoDoc.TpoDocRecHonorarios Then
            fraRetencSistPens.Visible = True
        Else
            fraRetencSistPens.Visible = False
        End If
    End If
    'END EJVG *******
   CalculaTotal
Set oDoc = Nothing
End Sub

Private Sub VerReciboEgreso()
Dim lsReciboEgreso As String
      txtFacSerie.MaxLength = 4 '***Agregado por ELRO el 20130611, según SATI INC1304290006
      txtFacSerie = ""
      txtFacSerie.Enabled = True
      txtFacNro.MaxLength = 11 '***Agregado por ELRO el 20130611, según SATI INC1304290006
      txtFacNro = ""
      txtFacNro.Enabled = True
   


'Dim lsReciboEgreso As String
'   If Mid(cboDoc.Text, 1, 2) = TpoDocRecEgreso Then
'      lsReciboEgreso = oContFunct.GeneraDocNro(TpoDocRecEgreso, gsCodUser, Mid(gsOpeCod, 3, 1))
'      txtFacSerie = Mid(lsReciboEgreso, 1, 4)
'      txtFacSerie.Enabled = False
'      txtFacNro = Mid(lsReciboEgreso, 6, 20)
'      txtFacNro.Enabled = False
'   Else
'      txtFacSerie = ""
'      txtFacSerie.Enabled = True
'      txtFacNro = ""
'      txtFacNro.Enabled = True
'   End If
End Sub

'Private Sub cboDoc_LostFocus()
    'If Trim(Left(cboDoc, 2)) = "05" Then
        'txtFacSerie = Trim(Str("3"))
    'End If
'End Sub '***NAGL ERS012-2017 Comentado by NAGL 20170927

Private Sub cboDoc_Validate(Cancel As Boolean)
If cboDoc = "" Then
    Cancel = True
End If
End Sub
Private Sub CalculaTotal(Optional lCalcImpuestos As Boolean = True, Optional lAsignaImporte As Boolean = True)
    Dim n As Integer, m As Integer
    Dim nSTot As Currency
    Dim nITot As Currency, nImp As Currency
    Dim nSTotI As Currency
    Dim nTot  As Currency
    Dim nTotImp As Currency
    Dim nTasaImp As Currency
    Dim nVV      As Currency
    Dim lnI As Integer
    Dim lnTotFila As Currency
    Dim lnAjuste As Currency
    Dim lnMontoRetencionProvSP As Currency 'EJVG20140726
    
    nSTot = 0: nTot = 0
     nSTotI = 0
    nTotImp = 0: nTasaImp = 0
    If fgImp.TextMatrix(1, 0) = "" Then
       lCalcImpuestos = False
    End If
    For m = 1 To fgImp.Rows - 1
       If fgImp.TextMatrix(m, 2) = "." Then
          If fgImp.TextMatrix(m, 11) = "0" Then
             nTasaImp = nTasaImp + nVal(fgImp.TextMatrix(m, 4))
          End If
       End If
    Next
    For m = 1 To fgDetalle.Rows - 1
        If Trim(sProvCod) <> "" And lCalcImpuestos And lAsignaImporte Then
    '      If fgImp.TextMatrix(m, 8) = "D" Then
    '        fgDetalle.TextMatrix(m, 7) = Round(Val(Format(fgDetalle.TextMatrix(N, 12), gsFormatoNumeroDato)) / (1 + (nTasaImp / 100)), 2)
    '      End If
            fgDetalle.TextMatrix(m, 7) = fgDetalle.TextMatrix(m, 12)
        End If
       nSTot = nSTot + nVal(fgDetalle.TextMatrix(m, 7))
    Next
    
    For m = 1 To fgImp.Rows - 1
       nITot = 0
       For n = 1 To fgDetalle.Rows - 1
          If fgImp.TextMatrix(m, 2) = "." And fgDetalle.TextMatrix(n, 1) <> "" Then
             If lCalcImpuestos Then
                If Trim(sProvCod) <> "" And lAsignaImporte Then
                    If fgImp.TextMatrix(m, 8) = "D" Then
                       nVV = Round(Val(Format(fgDetalle.TextMatrix(n, 12), gsFormatoNumeroDato)) / (1 + (nTasaImp / 100)), 2)
                       fgDetalle.TextMatrix(n, 7) = Format(nVV, gsFormatoNumeroView)
                       nImp = Round(nVV * ((Val(Format(fgImp.TextMatrix(m, 4), gsFormatoNumeroDato)) / 100)), 2)
                       
                       'John Cambio  ---------------
                       Dim T As Currency
                       T = Format(fgDetalle.TextMatrix(n, 12), gsFormatoNumeroDato)
                       
                        If nVV + nImp <> T Then
                        
                        lnAjuste = 0
                           'nImp = nVV * ((Val(Format(fgImp.TextMatrix(m, 4), gsFormatoNumeroDato)) / 100))
                           'nImp = fgTruncar(CDbl(nImp), 2)
                           lnAjuste = T - nVV - nImp
                           nImp = Round(CDbl(nImp) + CDbl(lnAjuste), 2)
                        End If
                        '-- Fin ---------------------
                       
                    Else
                       nImp = Round(Val(Format(fgDetalle.TextMatrix(n, 12), gsFormatoNumeroDato)) * (Val(Format(fgImp.TextMatrix(m, 4), gsFormatoNumeroDato)) / 100), 2)
                    End If
                Else
                    If fgImp.TextMatrix(m, 11) = "0" Then
                        nImp = Round(Val(Format(fgDetalle.TextMatrix(n, 7), gsFormatoNumeroDato)) * Val(Format(fgImp.TextMatrix(m, 4), gsFormatoNumeroDato)) / 100, 2)
                    Else
                        
                        lnTotFila = fgDetalle.TextMatrix(n, 7)
                        
                        For lnI = 13 To fgDetalle.Cols - 1
                            If Me.fgImp.TextMatrix(lnI - 13 + 1, 11) = "0" And fgDetalle.TextMatrix(n, lnI) <> "" Then
                                lnTotFila = lnTotFila + fgDetalle.TextMatrix(n, lnI)
                            End If
                        Next lnI
                        
                        nImp = Round(Val(Format(lnTotFila, gsFormatoNumeroDato)) * Val(Format(fgImp.TextMatrix(m, 4), gsFormatoNumeroDato)) / 100, 2)
                    End If
                End If
                
                'ALPA 20090303*******************************************
                        If fgImp.TextMatrix(m, 2) <> "." Or fgDetalle.TextMatrix(n, m + 12) = "" Then
                            '*** PEAC 20101104
                            If fgDetalle.Cols - 1 = 16 Then
                                If fgDetalle.TextMatrix(n, 16) = "A" Or fgDetalle.TextMatrix(n, 16) = "" Then
                                    fgDetalle.TextMatrix(n, m + 12) = Format(nImp, gsFormatoNumeroView)
                                End If
                            Else
                                fgDetalle.TextMatrix(n, m + 12) = Format(nImp, gsFormatoNumeroView)
                            End If
                            'fgDetalle.TextMatrix(N, m + 12) = Format(nImp, gsFormatoNumeroView)
                            '*** FIN PEAC
                        End If
                '********************************************************
                        nImp = nVal(fgDetalle.TextMatrix(n, m + 12))
             Else
                        nImp = nVal(fgDetalle.TextMatrix(n, m + 12))
             End If
             nITot = nITot + nImp
          Else
             If lCalcImpuestos Then fgDetalle.TextMatrix(n, m + 12) = ""
          End If
       Next
    '   If fgImp.TextMatrix(m, 9) = "0" And sProvCod <> "" Then
       If sProvCod <> "" Then
          nTotImp = nTotImp + nITot * IIf(fgImp.TextMatrix(m, 8) = "D", 1, -1)
       Else
          nTotImp = nTotImp + nITot * IIf(fgImp.TextMatrix(m, 8) = "D", 1, -1)
       End If
       'ALPA 20090303*******************************************
        If fgImp.TextMatrix(m, 2) = "." Then
            fgImp.TextMatrix(m, 5) = Format(nITot, gsFormatoNumeroView)
        Else
            fgImp.TextMatrix(m, 5) = Format(0, gsFormatoNumeroView)
        End If
        '*******************************************************
       nTot = nTot + nITot * IIf(fgImp.TextMatrix(m, 8) = "D", 1, -1)
    Next
    If sProvCod <> "" Then
        nSTot = 0
        For m = 1 To fgDetalle.Rows - 1
            nSTot = nSTot + nVal(fgDetalle.TextMatrix(m, 7))
        Next
    End If
    txtSTotal = Format(Abs(nSTot), gsFormatoNumeroView)
    If nTot < 0 Then
       txtSTotal.ForeColor = vbRed
    Else
       txtSTotal.ForeColor = vbBlack
    End If
    'EJVG20140726 ***
    fnRetProvSPMontoBase = 0#: fnRetProvSPAporte = 0#: fnRetProvSPSeguroAFP = 0#: fnRetProvSPComisionAFP = 0#
    If gbBitRetencSistPensProv Then
        If Val(Trim(Left(cboDoc.Text, 3))) = TpoDoc.TpoDocRecHonorarios Then
            If IsDate(txtFacFecha.Text) And txtBuscarProv.psCodigoPersona <> "" Then
                Dim oPSP As New NProveedorSistPens
                fnRetProvSPMontoBase = MontoBaseOperacion
                If oPSP.ExisteDatosSistemaPension(txtBuscarProv.psCodigoPersona) Then
                    If oPSP.AplicaRetencionSistemaPension(txtBuscarProv.psCodigoPersona, CDate(txtFacFecha.Text), fnRetProvSPMontoBase) Then
                        oPSP.SetDatosRetencionSistPens txtBuscarProv.psCodigoPersona, CDate(txtFacFecha.Text), fnRetProvSPMontoBase, Mid(gsOpeCod, 3, 1), fnRetProvSPAporte, fnRetProvSPSeguroAFP, fnRetProvSPComisionAFP
                    End If
                End If
                Set oPSP = Nothing
            End If
        End If
    End If
    lnMontoRetencionProvSP = fnRetProvSPAporte + fnRetProvSPSeguroAFP + fnRetProvSPComisionAFP
    txtRetSistPens.Text = Format(lnMontoRetencionProvSP, gsFormatoNumeroView)
    'END EJVG *******
    'If sProvCod = "" Then
    '    txtTotal = Format(nSTot + nTotImp, gsFormatoNumeroView)
    'Else
    '    txtTotal = Format(nSTot, gsFormatoNumeroView)
    'End If
    txtTotal = Format(nSTot + nTotImp - lnMontoRetencionProvSP, gsFormatoNumeroView)
End Sub


Private Sub cboDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtFacSerie.Enabled Then
        txtFacSerie.SetFocus
    Else
        If txtFacFecha.Enabled Then txtFacFecha.SetFocus
    End If
End If
End Sub



Private Sub cmdAgregar_Click()
'EJVG20130823 ***
If cboDoc.ListIndex = -1 Then
    MsgBox "Ud. debe de elegir tipo de documento", vbInformation, "Aviso"
    Exit Sub
End If
'END EJVG *******
'EJVG20140723 ***
If Len(Trim(txtBuscarProv.Text)) = 0 Then
    MsgBox "Ud. debe de seleccionar al Proveedor", vbInformation, "Aviso"
    Exit Sub
End If
'END EJVG *******
If fgDetalle.TextMatrix(1, 0) = "" Then
   AdicionaRow fgDetalle, 1
   EnfocaTexto txtObj, 0, fgDetalle
Else
   If Val(Format(fgDetalle.TextMatrix(fgDetalle.Rows - 1, 7), gsFormatoNumeroDato)) <> 0 And _
      Len(Format(fgDetalle.TextMatrix(fgDetalle.Rows - 1, 1), gsFormatoNumeroDato)) > 0 Then
      AdicionaRow fgDetalle, fgDetalle.Rows
      EnfocaTexto txtObj, 0, fgDetalle
   Else
      If fgDetalle.Enabled Then
         fgDetalle.SetFocus
      End If
   End If
End If

fgDetalle.row = fgDetalle.Rows - 1
fgDetalle.Col = 9
fgDetalle.TextMatrix(fgDetalle.row, 9) = ""
Set fgDetalle.CellPicture = picCuadroNo
fgDetalle.Col = 1
End Sub

Private Sub cmdAjuste_Click()
If nVal(txtTotal) = 0 Then
   MsgBox "Primero ingresar Concepto de Gastos de Documento", vbInformation, "Aviso"
   Exit Sub
End If
If cmdAjuste.Caption = "Asignar A&juste   >>>" Then
   cmdAjuste.Visible = False
   cmdDetraccion.Visible = False
   fraAjuste.Visible = True
   
   fgDetalle.Enabled = False
   fgImp.Enabled = False
   fgObj.Enabled = False
   
   TabDocRef.Enabled = False
   txtAjuste.SetFocus
Else
   fgImp.EliminaFila fgImp.Rows - 1
   cmdAjuste.Caption = "Asignar A&juste   >>>"
   fgDetalle.Cols = fgDetalle.Cols - 1
   fgDetalle.SetFocus
End If
End Sub
Private Sub DesactivaAjuste()
cmdAjuste.Visible = True
cmdDetraccion.Visible = True
fraAjuste.Visible = False
fgDetalle.Enabled = True
fgImp.Enabled = True
fgObj.Enabled = True
fgDetalle.Enabled = True
TabDocRef.Enabled = True
fgDetalle.SetFocus
End Sub

Private Sub cmdAplicar_Click()
Dim nRow As Integer
Dim nTot As Currency
Dim nImp As Currency
If txtAjuste.Text = "" Then
MsgBox "Debe ingresar Ajuste", vbInformation, "Aviso"
Exit Sub
Else
fgImp.AdicionaFila
nRow = fgImp.row
fgImp.Col = 0
fgImp.Col = 1
'Falta adicionar Check
fgImp.TextMatrix(nRow, 0) = nRow
fgImp.TextMatrix(nRow, 1) = "AJUSTE"
fgImp.TextMatrix(nRow, 2) = "1"
fgImp.TextMatrix(nRow, 3) = "AJUSTE"
fgImp.TextMatrix(nRow, 4) = Format(nVal(txtAjuste) * 100 / nVal(txtSTotal), gsFormatoNumeroView)
fgImp.TextMatrix(nRow, 6) = "AJUSTE"
fgImp.TextMatrix(nRow, 7) = "0"   'Impuesto Calculado
fgImp.TextMatrix(nRow, 8) = "D"
fgImp.TextMatrix(nRow, 9) = "1"  'Para q se grabe en MovOtrosItem Variable Ajuste
fgImp.TextMatrix(nRow, 10) = "2"
fgImp.TextMatrix(nRow, 11) = "0"

'Distribución del Ajuste entre Cuentas de Gasto
fgDetalle.Cols = fgDetalle.Cols + 1
fgDetalle.TextMatrix(0, fgDetalle.Cols - 1) = "AJUSTE"
nTot = 0
For nRow = 1 To fgDetalle.Rows - 1
   nImp = Round(nVal(txtAjuste) * nVal(fgDetalle.TextMatrix(nRow, 7)) / nVal(txtSTotal), 2)
   nTot = nTot + nImp
   fgDetalle.TextMatrix(nRow, fgDetalle.Cols - 1) = Format(nImp, gsFormatoNumeroView)
Next
If nTot <> nVal(txtAjuste) Then
   fgDetalle.TextMatrix(1, fgDetalle.Cols - 1) = nVal(fgDetalle.TextMatrix(1, fgDetalle.Cols - 1)) + (nVal(txtAjuste) - nTot)
End If
cmdAjuste.Caption = "Desasignar A&juste   >>>"
CalculaTotal False
DesactivaAjuste
End If
End Sub

Private Sub cmdCancelar_Click()
cmdAjuste.Caption = "Asignar A&juste   >>>"
DesactivaAjuste
End Sub

Private Sub cmdDetraccion_Click()
Dim nRow As Integer
Dim nTot As Currency
Dim nImp As Currency
Dim lnAjuste As Currency
Dim ctaDetra As String
Dim resp As Integer
Dim rs As ADODB.Recordset


If nVal(txtTotal) = 0 Then
   MsgBox "Primero ingresar Concepto de Gastos de Documento", vbInformation, "Aviso"
   Exit Sub
End If
If cmdDetraccion.Caption = "Aplicar &Detracción" Then
                
       Set oContFunct = New NContFunciones
       
       ctaDetra = Mid(fgDetalle.TextMatrix(fgDetalle.row, 1), 1, 2) & "0" & Mid(fgDetalle.TextMatrix(fgDetalle.row, 1), 4, Len(fgDetalle.TextMatrix(fgDetalle.row, 1)))
       resp = oContFunct.VerificaCtaDetraccion(ctaDetra)
       If resp = 0 Then
          MsgBox "Cuenta no permitada para Detraccion", vbInformation, "Aviso"
          Exit Sub
       End If
       '--------------------------------------------
       '   '--Detracción por IGV 14%
       Set rs = New ADODB.Recordset
       
       Set rs = oContFunct.GetImpuestoCuentaDetrac(ctaDetra)
       If Not rs.EOF Then
          lnTasaDetraccion = rs!nDetraPorc * 100
       Else
          MsgBox "Cuenta de Detracción " & ctaDetra & " no definida como Impuesto", vbInformation, "¡Aviso!"
          lSalir = True
          Exit Sub
       End If
       
      '------------------------------------------------
   lnAjuste = Round(nVal(txtTotal) * lnTasaDetraccion / 100, 2)
   '***********************************************************'
   '** Add by GITU 30-09-2008 segun el acta 195-2008/TI-D    **'
   '***********************************************************'
   lnAjuste = CInt(lnAjuste)
   '***********************************************************'
   frmOpeNegVentanilla.Inicio "", Mid(gsOpeCod, 3, 1), lnAjuste, txtBuscarProv, lblProvNombre, "Detracción", gsCodAge, lsCtaDetraccion, "D", False
   If frmOpeNegVentanilla.lbOk Then
      Set rsDetracc = frmOpeNegVentanilla.rsPago
      cmdDetraccion.Caption = "Sin &Detracción"
      fgImp.AdicionaFila
      nRow = fgImp.row
      fgImp.Col = 2
      fgImp.TextMatrix(nRow, 1) = "DETRACC."
      fgImp.TextMatrix(nRow, 2) = "1"
      fgImp.TextMatrix(nRow, 0) = nRow
      fgImp.TextMatrix(nRow, 3) = "DETRACC."
      fgImp.TextMatrix(nRow, 4) = Format(lnTasaDetraccion, gsFormatoNumeroView)
      fgImp.TextMatrix(nRow, 5) = Format(lnAjuste, "#,#0.00")
      fgImp.TextMatrix(nRow, 6) = lsCtaDetraccion
      fgImp.TextMatrix(nRow, 7) = "DETRACCION"
      fgImp.TextMatrix(nRow, 8) = "H"
      fgImp.TextMatrix(nRow, 9) = "0"
      fgImp.TextMatrix(nRow, 10) = "2"
      fgImp.TextMatrix(nRow, 11) = ""
      
      'Distribución del Ajuste entre Cuentas de Gasto
      fgDetalle.Cols = fgDetalle.Cols + 1
      fgDetalle.TextMatrix(0, fgDetalle.Cols - 1) = "DETRACC."
      nTot = 0
      For nRow = 1 To fgDetalle.Rows - 1
         nImp = Round(lnAjuste * nVal(fgDetalle.TextMatrix(nRow, 7)) / nVal(txtSTotal), 2)
         nTot = nTot + nImp
         fgDetalle.TextMatrix(nRow, fgDetalle.Cols - 1) = Format(nImp, gsFormatoNumeroView)
      Next
      If nTot <> lnAjuste Then
         fgDetalle.TextMatrix(1, fgDetalle.Cols - 1) = Format(Val(fgDetalle.TextMatrix(1, fgDetalle.Cols - 1)) + (lnAjuste - nTot), gsFormatoNumeroView)
      End If
   End If
Else
   fgImp.EliminaFila fgImp.Rows - 1
   cmdDetraccion.Caption = "Aplicar &Detracción"
   fgDetalle.Cols = fgDetalle.Cols - 1
   fgDetalle.SetFocus
End If
CalculaTotal False
End Sub
'ALPA 20091110*********************************
Private Sub cmdDistribución_Click()
If cboDoc.ListIndex = -1 Then
    MsgBox "Ud. debe de elegir tipo de documento", vbInformation, "Aviso"
    Exit Sub
End If
Call frmAgenciaPorcentajeGastosProvision.Inicio(sMatrizDatos, nTipoMatriz, nCantAgeSel, nMontogasto, , , nMontoInafecto)
nMontoInafecto = 0 'EJVG20130823
'*** PEAC 20101103 - SE AGREGO "nMontoInafecto"

End Sub
'**********************************************
Private Sub cmdEliminar_Click()
'If fgDetalle.TextMatrix(fgDetalle.Row, 0) <> "" Then
'   EliminaCuenta fgDetalle.TextMatrix(fgDetalle.Row, 1), fgDetalle.TextMatrix(fgDetalle.Row, 0)
'   If fgImp.TextMatrix(2, 0) = 1 Or fgImp.TextMatrix(2, 1) = "DETRACC." Then
'         fgImp.EliminaFila fgImp.Rows - 1
'         fgDetalle.Cols = fgDetalle.Cols - 1
'          cmdDetraccion.Caption = "Aplicar &Detracción"
'         fgDetalle.SetFocus
'   End If
'   CalculaTotal
'   If fgDetalle.Enabled Then
'      fgDetalle.SetFocus
'   End If
'End If
'edpyme - enviado por john
If fgDetalle.TextMatrix(fgDetalle.row, 0) <> "" Then
   EliminaCuenta fgDetalle.TextMatrix(fgDetalle.row, 1), fgDetalle.TextMatrix(fgDetalle.row, 0)
  
     If fgImp.Rows - 1 = 2 Then
      If Val(fgImp.TextMatrix(2, 0)) = 1 Or fgImp.TextMatrix(2, 1) = "DETRACC." Then
         fgImp.EliminaFila fgImp.Rows - 1
         fgDetalle.Cols = fgDetalle.Cols - 1
          cmdDetraccion.Caption = "Aplicar &Detracción"
         fgDetalle.SetFocus
      End If
     End If
   CalculaTotal
   If fgDetalle.Enabled Then
      fgDetalle.SetFocus
   End If
End If


End Sub

Private Sub EliminaCuenta(sCod As String, nItem As Integer)
EliminaRow fgDetalle, fgDetalle.row
EliminaFgObj nItem
If Len(fgDetalle.TextMatrix(1, 1)) > 0 Then
   RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
End If
End Sub

'Private Sub cmdEmbargo_Click()
'Dim nRow As Integer
'Dim nTot As Currency
'Dim nImp As Currency
'Dim lnAjuste As Currency
'
'Dim resp As Integer
'Dim rs As ADODB.Recordset
'
'
'If nVal(txtTotal) = 0 Then
'   MsgBox "Primero ingresar Concepto de Gastos de Documento", vbInformation, "Aviso"
'   Exit Sub
'End If
'If cmdEmbargo.Caption = "Aplicar &Fiel Cump." Then
'
'       Set oContFunct = New NContFunciones
'       Set rs = New ADODB.Recordset
'       Set rs = oContFunct.GetCuentaEmbargo(gsOpeCod, 6)
'       If Not rs.EOF Then
'          ctaEmbargo = rs!cCtaContCod
'       Else
'          MsgBox "Cuenta de Embargo " & ctaEmbargo & " No definida", vbInformation, "¡Aviso!"
'          Exit Sub
'       End If
'
'
'
'      cmdEmbargo.Caption = "Sin &Fiel Cumpli."
'      fgImp.AdicionaFila
'      nRow = fgImp.Row
'      fgImp.Col = 2
'      fgImp.TextMatrix(nRow, 1) = "FIEL."
'      fgImp.TextMatrix(nRow, 2) = "0"
'      fgImp.TextMatrix(nRow, 0) = nRow
'      fgImp.TextMatrix(nRow, 3) = "FIEL."
'      fgImp.TextMatrix(nRow, 4) = Format(0, "#,#0.00")
'      fgImp.TextMatrix(nRow, 5) = Format(0, "#,#0.00")
'      fgImp.TextMatrix(nRow, 6) = ctaEmbargo
'      fgImp.TextMatrix(nRow, 7) = "FIEL"
'      fgImp.TextMatrix(nRow, 8) = "H"
'      fgImp.TextMatrix(nRow, 9) = "0"
'      fgImp.TextMatrix(nRow, 10) = "2"
'      fgImp.TextMatrix(nRow, 11) = ""
'
'      'Distribución del Ajuste entre Cuentas de Gasto
'      fgDetalle.Cols = fgDetalle.Cols + 1
'      fgDetalle.TextMatrix(0, fgDetalle.Cols - 1) = "FIEL."
'      nTot = 0
'Else
'   fgImp.EliminaFila fgImp.Rows - 1
'   cmdEmbargo.Caption = "Aplicar &Fiel Cump."
'   fgDetalle.Cols = fgDetalle.Cols - 1
'   fgDetalle.SetFocus
'End If
'CalculaTotal False
'End Sub
Private Sub ObtenerValoresPorCtas()
   Dim nNiv  As Integer
   Dim sCtaCod  As String
   Dim sCtaDes  As String
 If nCantAgeSel > 0 Then
            Dim nContx As Integer
            Dim cAgeTemp As String
            Dim cAgeCodTemp As String
            Dim cDescAgeTemp As String
            Dim cCtaContTemp As String
            Dim nSumaItmes As Currency
            Dim sFacSerie As String
            Dim sFacNro As String
            nSumaItmes = 0
            sCtaCod = txtObj.Text
            cCtaContTemp = fgDetalle.TextMatrix(fgDetalle.row, 2)
            If nCantAgeSel > 0 Then
             'cCtaContTemp = fgObj.TextMatrix(fgObj.Row, 2)
             sFacNro = txtFacNro.Text
             sFacSerie = txtFacSerie.Text
             EliminaCuenta fgDetalle.TextMatrix(fgDetalle.row, 1), fgDetalle.TextMatrix(fgDetalle.row, 0)
             fgImp.EliminaFila fgImp.Rows - 1
             cAgeTemp = fgObj.TextMatrix(fgObj.row, 5)
             sCtaDes = cCtaContTemp
             Call cboDoc_Click
             txtFacSerie.Text = sFacSerie
             txtFacNro.Text = sFacNro
            For nContx = 1 To nCantAgeSel
             AdicionaRow fgDetalle, fgDetalle.Rows
             EnfocaTexto txtObj, 0, fgDetalle
             sCtaCod = IIf(Len(Trim(cSubCta)) > Len(Trim(sCtaCod)), cSubCta, sCtaCod)
             fgDetalle = sCtaCod
             cDescAgeTemp = sMatrizDatos(2, nContx)
             cAgeCodTemp = sMatrizDatos(1, nContx)
             fgDetalle.TextMatrix(fgDetalle.row, 2) = sCtaDes & "-" & cDescAgeTemp
             fgDetalle.TextMatrix(fgDetalle.row, 7) = sCtaCod
             'nSumaItmes = nSumaItmes + nMontogasto * (sMatrizDatos(3, nContx) / 100)
             If nCantAgeSel > nContx Then
                fgDetalle.TextMatrix(fgDetalle.row, 7) = Round(nMontogasto * (sMatrizDatos(3, nContx) / 100), 2)
                nSumaItmes = nSumaItmes + Round(nMontogasto * (sMatrizDatos(3, nContx) / 100), 2)
             Else
                fgDetalle.TextMatrix(fgDetalle.row, 7) = nMontogasto - Round(nSumaItmes, 2)
                nSumaItmes = nSumaItmes + (nMontogasto - nSumaItmes)
             End If
             CalculaTotal IIf(fgDetalle.Col > 12, False, True), False
    '                AdicionaObj sCtaCod, fgDetalle.TextMatrix(fgDetalle.Row, 0), fgObj.TextMatrix(fgObj.Row, 1), fgObj.TextMatrix(fgObj.Row, 2), _
    '                            fgObj.TextMatrix(fgObj.Row, 3), fgObj.TextMatrix(fgObj.Row, 5), fgObj.TextMatrix(fgObj.Row, 6)
                     AdicionaObj cSubCta, nContx, nContx, Mid(cCtaContTemp, 1, 3) & sMatrizDatos(1, nContx), _
                                cDescAgeTemp, cAgeCodTemp, fgObj.TextMatrix(fgObj.row, 6)

            Next nContx
            End If
            Else
                'sCtaCod = IIf(Len(Trim(cSubCta)) > Len(Trim(sCtaCod)), cSubCta, sCtaCod)
                'fgDetalle = sCtaCod
                'fgDetalle.TextMatrix(fgDetalle.Row, 7) = sCtaCod
            End If
       '**************************************************************************
       fgDetalle.Enabled = True
       txtObj.Visible = False
       cmdExaminar.Visible = False
       fgDetalle.Col = 7
       fgDetalle.SetFocus

    
    If Not rs Is Nothing Then
       If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
    End If
    If nCantAgeSel = 0 Then
        If fgDetalle.TextMatrix(fgDetalle.row, 0) <> "" Then
           RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
        End If
    End If
End Sub
Private Sub cmdexaminar_Click()
    Dim sSqlO As String
    Dim rsObj As ADODB.Recordset
    Set rsObj = New ADODB.Recordset
    Dim oCon As DConecta

    Set oCon = New DConecta
    
    Dim nNiv  As Integer
    Dim sCtaCod  As String
    Dim sCtaDes  As String
    Dim lnMontoAfecto As Currency
    oCon.AbreConexion
    If lbMismaOpe Then
        sSqlO = "SELECT DISTINCT a.cCtaContCod as cObjetoCod, b.cCtaContDesc, 2 as nObjetoNiv " _
              & "FROM  OpeCta a JOIN CtaCont b ON b.cCtaContCod = a.cCtaContCod " _
              & "WHERE a.cOpeCod = '" & gsOpeCod & "' AND a.cOpeCtaDH='D'"
    Else
        sSqlO = "SELECT DISTINCT a.cCtaContCod as cObjetoCod, b.cCtaContDesc, 2 as nObjetoNiv " _
              & "FROM  OpeCta a JOIN CtaCont b ON b.cCtaContCod = a.cCtaContCod " _
              & "WHERE a.cOpeCod like '" & gsOpeCod & "' AND a.cOpeCtaDH='D'"
    End If
    Set rs = oCon.CargaRecordSet(sSqlO)
    If rs.EOF Then
        MsgBox "No se asignaron Conceptos a Operación", vbCritical, "Error"
        txtObj.SetFocus
        Exit Sub
    End If
    Dim oDesc As New ClassDescObjeto
    oDesc.lbUltNivel = True
    oDesc.Show rs, txtObj.Text, Me.cboProvis
    
    
    If oDesc.lbOk Then
       sCtaCod = oDesc.gsSelecCod: sCtaDes = oDesc.gsSelecDesc
       AsignaObjetosSer sCtaCod, sCtaDes
'       ALPA 20091113*************************************************************
'       If gsOpeCod = "701401" Or gsOpeCod = "701402" Or gsOpeCod = "702401" Or gsOpeCod = "702402" Then
            
'       Else
'           EliminaFgObj fgDetalle.TextMatrix(fgDetalle.Row, 0)
'       End If
'       sMatrizDatos, nTipoMatriz, nCantAgeSel

        If nCantAgeSel > 0 Then
            Dim nContx As Integer
            Dim cAgeTemp As String
            Dim cAgeCodTemp As String
            Dim cDescAgeTemp As String
            Dim cCtaContTemp As String
            Dim nSumaItmes As Currency
            Dim sFacSerie As String
            Dim sFacNro As String
            
            Dim lnTotMonto As Currency '*** PEAC 20101104
            Dim lnCorrelativo As Integer
            
            lnCorrelativo = 0
            nSumaItmes = 0
                        
            If nCantAgeSel > 0 Then
             sFacNro = txtFacNro.Text
             sFacSerie = txtFacSerie.Text
             cCtaContTemp = fgObj.TextMatrix(fgObj.row, 2)
             EliminaCuenta fgDetalle.TextMatrix(fgDetalle.row, 1), fgDetalle.TextMatrix(fgDetalle.row, 0)
             fgImp.EliminaFila fgImp.Rows - 1
             Call cboDoc_Click
             txtFacSerie.Text = sFacSerie
             txtFacNro.Text = sFacNro
             cAgeTemp = fgObj.TextMatrix(fgObj.row, 5)

            '*** PEAC 20101103
            If nMontoInafecto > 0 Then
                lnMontoAfecto = nMontogasto - nMontoInafecto
                nMontogasto = Round(lnMontoAfecto / ((lnCalBaseImpIGV + 100) / 100), 2)
                lnTotMonto = 0

                For nContx = 1 To nCantAgeSel
                 AdicionaRow fgDetalle, fgDetalle.Rows
                 EnfocaTexto txtObj, 0, fgDetalle
                 fgDetalle = IIf(Len(Trim(cSubCta)) > Len(Trim(sCtaCod)), cSubCta, sCtaCod)
                 cDescAgeTemp = sMatrizDatos(2, nContx)
                 cAgeCodTemp = sMatrizDatos(1, nContx)
                 fgDetalle.TextMatrix(fgDetalle.row, 2) = sCtaDes & "-" & cDescAgeTemp
                 fgDetalle.TextMatrix(fgDetalle.row, 8) = sCtaCod
                 fgDetalle.TextMatrix(fgDetalle.row, 16) = "A"

                 If nCantAgeSel > nContx Then
                    fgDetalle.TextMatrix(fgDetalle.row, 7) = Round(nMontogasto * (sMatrizDatos(3, nContx) / 100), 2)
                    nSumaItmes = nSumaItmes + Round(nMontogasto * (sMatrizDatos(3, nContx) / 100), 2)
                 Else
                    fgDetalle.TextMatrix(fgDetalle.row, 7) = nMontogasto - Round(nSumaItmes, 2)
                    nSumaItmes = nSumaItmes + (nMontogasto - nSumaItmes)
                 End If
                 CalculaTotal IIf(fgDetalle.Col > 12, False, True), False
                         AdicionaObj sCtaCod, nContx, nContx, Mid(cCtaContTemp, 1, 3) & sMatrizDatos(1, nContx), _
                                    cDescAgeTemp, cAgeCodTemp, fgObj.TextMatrix(fgObj.row, 6)

                lnTotMonto = lnTotMonto + nMontogasto

                Next nContx

                If lnTotMonto Then

                End If

                nMontogasto = nMontoInafecto
                nSumaItmes = 0
                lnCorrelativo = nContx - 1
            End If
            '*** FIN PEAC
            
            For nContx = 1 To nCantAgeSel
            
            lnCorrelativo = lnCorrelativo + 1
            
             AdicionaRow fgDetalle, fgDetalle.Rows
             EnfocaTexto txtObj, 0, fgDetalle
             fgDetalle = IIf(Len(Trim(cSubCta)) > Len(Trim(sCtaCod)), cSubCta, sCtaCod)
             cDescAgeTemp = sMatrizDatos(2, nContx)
             cAgeCodTemp = sMatrizDatos(1, nContx)
             fgDetalle.TextMatrix(fgDetalle.row, 2) = sCtaDes & "-" & cDescAgeTemp
             fgDetalle.TextMatrix(fgDetalle.row, 8) = sCtaCod
             
             If Trim(Val(Mid(cboDoc.Text, 1, 3))) = "1" And nMontoInafecto > 0 Then
                fgDetalle.TextMatrix(fgDetalle.row, 16) = "I"
             End If
             
             'nSumaItmes = nSumaItmes + nMontogasto * (sMatrizDatos(3, nContx) / 100)
                          
             If nCantAgeSel > nContx Then
                fgDetalle.TextMatrix(fgDetalle.row, 7) = Round(nMontogasto * (sMatrizDatos(3, nContx) / 100), 2)
                nSumaItmes = nSumaItmes + Round(nMontogasto * (sMatrizDatos(3, nContx) / 100), 2)
             Else
                fgDetalle.TextMatrix(fgDetalle.row, 7) = nMontogasto - Round(nSumaItmes, 2)
                nSumaItmes = nSumaItmes + (nMontogasto - nSumaItmes)
             End If
             
             
             CalculaTotal IIf(fgDetalle.Col > 12, False, True), False
    '                AdicionaObj sCtaCod, fgDetalle.TextMatrix(fgDetalle.Row, 0), fgObj.TextMatrix(fgObj.Row, 1), fgObj.TextMatrix(fgObj.Row, 2), _
    '                            fgObj.TextMatrix(fgObj.Row, 3), fgObj.TextMatrix(fgObj.Row, 5), fgObj.TextMatrix(fgObj.Row, 6)
                     
'                     AdicionaObj sCtaCod, nContx, nContx, Mid(cCtaContTemp, 1, 3) & sMatrizDatos(1, nContx), _
'                                cDescAgeTemp, cAgeCodTemp, fgObj.TextMatrix(fgObj.Row, 6)

                     AdicionaObj sCtaCod, lnCorrelativo, lnCorrelativo, Mid(cCtaContTemp, 1, 3) & sMatrizDatos(1, nContx), _
                                cDescAgeTemp, cAgeCodTemp, fgObj.TextMatrix(fgObj.row, 6)

    
            Next nContx
            
            
            End If
        End If
       '**************************************************************************
       fgDetalle.Enabled = True
       txtObj.Visible = False
       cmdExaminar.Visible = False
       fgDetalle.Col = 7
       fgDetalle.SetFocus
    Else
       txtObj.SetFocus
    End If
    If Not rs Is Nothing Then
       If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
    End If
    If nCantAgeSel = 0 Then
        If fgDetalle.TextMatrix(fgDetalle.row, 0) <> "" Then
           RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
        End If
    End If
End Sub

Private Sub cmdFielCumplimiento_Click()
Dim nRow As Integer
Dim nTot As Currency
Dim nImp As Currency
Dim lnAjuste As Currency

Dim resp As Integer
Dim rs As ADODB.Recordset


If nVal(txtTotal) = 0 Then
   MsgBox "Primero ingresar Concepto de Gastos de Documento", vbInformation, "Aviso"
   Exit Sub
End If
If cmdFielCumplimiento.Caption = "Aplicar &Fiel Cump." Then
                
       'Set oContFunct = New NContFunciones
       'Set rs = New ADODB.Recordset
       'Set rs = oContFunct.GetCuentaEmbargo(gsOpeCod, 6)
       'If Not rs.EOF Then
       '   ctaEmbargo = rs!cCtaContCod
       'Else
       '   MsgBox "Cuenta de Embargo " & ctaEmbargo & " No definida", vbInformation, "¡Aviso!"
       '   Exit Sub
       'End If
       
       
       Dim oCon As NConstSistemas
       Set oCon = New NConstSistemas
       
      ctaFielCump = oCon.LeeConstSistema(172)
      If ctaFielCump = "" Then
         MsgBox "Cuenta de Fiel Cumplimiento No definida", vbInformation, "¡Aviso!"
         Exit Sub
       End If
       
       If Mid(gsOpeCod, 3, 1) = 1 Then
           ctaFielCump = Mid(ctaFielCump, 1, 2) & "1" & Mid(ctaFielCump, 4, Len(Trim(ctaFielCump)))
        Else
           ctaFielCump = Mid(ctaFielCump, 1, 2) & "2" & Mid(ctaFielCump, 4, Len(Trim(ctaFielCump)))
        End If
      
             
      cmdFielCumplimiento.Caption = "Sin &Fiel Cumpli."
      fgImp.AdicionaFila
      nRow = fgImp.row
      fgImp.Col = 2
      fgImp.TextMatrix(nRow, 1) = "FIEL."
      fgImp.TextMatrix(nRow, 2) = "0"
      fgImp.TextMatrix(nRow, 0) = nRow
      fgImp.TextMatrix(nRow, 3) = "FIEL."
      fgImp.TextMatrix(nRow, 4) = Format(0, "#,#0.00")
      fgImp.TextMatrix(nRow, 5) = Format(0, "#,#0.00")
      fgImp.TextMatrix(nRow, 6) = ctaFielCump
      fgImp.TextMatrix(nRow, 7) = "FIEL"
      fgImp.TextMatrix(nRow, 8) = "H"
      fgImp.TextMatrix(nRow, 9) = "0"
      fgImp.TextMatrix(nRow, 10) = "2"
      fgImp.TextMatrix(nRow, 11) = ""
      
      'Distribución del Ajuste entre Cuentas de Gasto
      fgDetalle.Cols = fgDetalle.Cols + 1
      fgDetalle.TextMatrix(0, fgDetalle.Cols - 1) = "FIEL."
      nTot = 0
'      For nRow = 1 To fgDetalle.Rows - 1
'         nImp = Round(lnAjuste * nVal(fgDetalle.TextMatrix(nRow, 7)) / nVal(txtSTotal), 2)
'         nTot = nTot + nImp
'         fgDetalle.TextMatrix(nRow, fgDetalle.Cols - 1) = Format(nImp, gsFormatoNumeroView)
'      Next
'      If nTot <> lnAjuste Then
'         fgDetalle.TextMatrix(1, fgDetalle.Cols - 1) = Format(Val(fgDetalle.TextMatrix(1, fgDetalle.Cols - 1)) + (lnAjuste - nTot), gsFormatoNumeroView)
'      End If
   'End If
   'ALPA 200903***************************
   nTFielCumplimiento = 1
   '**************************************
Else
   fgImp.EliminaFila fgImp.Rows - 1
   cmdFielCumplimiento.Caption = "Aplicar &Fiel Cump."
   fgDetalle.Cols = fgDetalle.Cols - 1
   fgDetalle.SetFocus
   'ALPA 200903***************************
   nTFielCumplimiento = 0
   '**************************************
End If
CalculaTotal False
End Sub
'ALPA 20090303*****************************************************************************************
Private Sub cmdRetenciones_Click()
Dim nRow As Integer
Dim nTot As Currency
Dim nImp As Currency
Dim lnAjuste As Currency

Dim resp As Integer
Dim rs As ADODB.Recordset


If nVal(txtTotal) = 0 Then
   MsgBox "Primero ingresar Concepto de Gastos de Documento", vbInformation, "Aviso"
   Exit Sub
End If
If cmdRetenciones.Caption = "Aplicar &Penalidad" Then

       
       Dim oCon As NConstSistemas
       Set oCon = New NConstSistemas
       
        If lbBienes Then
            ctaRetenciones = oCon.LeeConstSistema(402)
        Else
            ctaRetenciones = oCon.LeeConstSistema(401)
        End If
      If ctaRetenciones = "" Then
         MsgBox "Cuenta de Retenciones No definida", vbInformation, "¡Aviso!"
         Exit Sub
       End If
       
       If Mid(gsOpeCod, 3, 1) = 1 Then
           ctaRetenciones = Mid(ctaRetenciones, 1, 2) & "1" & Mid(ctaRetenciones, 4, Len(Trim(ctaRetenciones)))
        Else
           ctaRetenciones = Mid(ctaRetenciones, 1, 2) & "2" & Mid(ctaRetenciones, 4, Len(Trim(ctaRetenciones)))
        End If
      
             
      cmdRetenciones.Caption = "Sin &Penalidad"
      fgImp.AdicionaFila
      nRow = fgImp.row
      fgImp.Col = 2
      fgImp.TextMatrix(nRow, 1) = "RETE."
      fgImp.TextMatrix(nRow, 2) = "0"
      fgImp.TextMatrix(nRow, 0) = nRow
      fgImp.TextMatrix(nRow, 3) = "RETE."
      fgImp.TextMatrix(nRow, 4) = Format(0, "#,#0.00")
      fgImp.TextMatrix(nRow, 5) = Format(0, "#,#0.00")
      fgImp.TextMatrix(nRow, 6) = ctaRetenciones
      fgImp.TextMatrix(nRow, 7) = "RETE"
      fgImp.TextMatrix(nRow, 8) = "H"
      fgImp.TextMatrix(nRow, 9) = "0"
      fgImp.TextMatrix(nRow, 10) = "2"
      fgImp.TextMatrix(nRow, 11) = ""
      
      'Distribución del Ajuste entre Cuentas de Gasto
      fgDetalle.Cols = fgDetalle.Cols + 1
      fgDetalle.TextMatrix(0, fgDetalle.Cols - 1) = "RETE."
      nTot = 0
      '20090302*****************
      nTPenalidad = 1
      '*************************
Else
   fgImp.EliminaFila fgImp.Rows - 1
   cmdRetenciones.Caption = "Aplicar &Penalidad"
   fgDetalle.Cols = fgDetalle.Cols - 1
   fgDetalle.SetFocus
   '20090302*****************
    nTPenalidad = 0
    '*************************
End If
CalculaTotal False
End Sub

Private Sub cmdRetSistPensActualizar_Click()
    On Error GoTo ErrcmdRetSistPensActualizar
    Dim oPSP As NProveedorSistPens
    Dim frm As frmProveedorRegSistemaPension
    
    If Not IsDate(txtFacFecha.Text) Then
        MsgBox "Ud. debe ingresar la fecha del comprobante", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtBuscarProv.psCodigoPersona)) = 0 Then
        MsgBox "Ud. debe de seleccionar al Proveedor", vbInformation, "Aviso"
        Exit Sub
    Else
        Set oPSP = New NProveedorSistPens
        cmdRetSistPensActualizar.Enabled = False
        If oPSP.AplicaRetencionSistemaPension(txtBuscarProv.psCodigoPersona, CDate(Me.txtFacFecha.Text), MontoBaseOperacion) Then
            Do While Not oPSP.ExisteDatosSistemaPension(txtBuscarProv.psCodigoPersona)
                If MsgBox("Para continuar Ud. debe registrar los datos de Sistema Pensión del Proveedor", vbInformation + vbYesNo, "Aviso") = vbYes Then
                    Set frm = New frmProveedorRegSistemaPension
                    frm.Registrar (txtBuscarProv.psCodigoPersona)
                Else
                    cmdRetSistPensActualizar.Enabled = True
                    Set oPSP = Nothing
                    Exit Sub
                End If
            Loop
        End If
    End If
    Call CalculaTotal
    MsgBox "Retención de Sistema de Pensión recalculado", vbInformation, "Aviso"
    cmdRetSistPensActualizar.Enabled = True
    
    Set oPSP = Nothing
    Set frm = Nothing
    Exit Sub
ErrcmdRetSistPensActualizar:
    cmdRetSistPensActualizar.Enabled = True
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdRetSistPensDetalle_Click()
    Dim frm As frmProveedorMuestraRetencion
    On Error GoTo ErrcmdRetSistPensDetalle
    If Len(Trim(txtBuscarProv.psCodigoPersona)) = 0 Then
        MsgBox "Ud. debe de seleccionar al Proveedor", vbInformation, "Aviso"
        Exit Sub
    End If
    If Not IsDate(txtFacFecha) Then
        MsgBox "Ud. debe ingresar la fecha del comprobante", vbInformation, "Aviso"
        Exit Sub
    End If
    Set frm = New frmProveedorMuestraRetencion
    frm.Iniciar fnRetProvSPAporte, fnRetProvSPSeguroAFP, fnRetProvSPComisionAFP
    Set frm = Nothing
    Exit Sub
ErrcmdRetSistPensDetalle:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdValVenta_Click()
Dim n As Integer
If fgImp.TextMatrix(fgImp.row, 0) = "" Or fgImp.TextMatrix(fgImp.row, 2) = "" Then
    Exit Sub
End If

For n = 1 To fgDetalle.Rows - 1
    fgDetalle.TextMatrix(n, 7) = Format(Round(Val(Format(fgDetalle.TextMatrix(n, 7), gsFormatoNumeroDato)) / (1 + (nVal(fgImp.TextMatrix(fgImp.row, 4)) / 100)), 2), gsFormatoNumeroView)
Next
CalculaTotal
nVariaIGV = 0
End Sub

Private Sub fgDetalle_RowColChange()
If fgDetalle.TextMatrix(fgDetalle.row, 0) <> "" Then
    RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
    'ALPA 20091222**************
    Call ModidicarTotalPorItem
    '***************************
End If
End Sub

Private Sub fgImp_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
'John cambio
If fgImp.TextMatrix(pnRow, 3) = "IGV" And fgImp.TextMatrix(pnRow, 5) = "0.00" Then

   Dim oConst As NConstSistemas
   Set oConst = New NConstSistemas

   cboDocDestino.Enabled = True
   Me.cboDocDestino.ListIndex = oConst.LeeConstSistema(gConstSistDestinoIGVDefecto)
Else
   ' If fgImp.TextMatrix(pnRow, 3) = "FIEL." And fgImp.TextMatrix(pnRow, 5) = "0.00" Then
    If fgImp.TextMatrix(pnRow, 3) = "FIEL." Then
    
    ElseIf fgImp.TextMatrix(pnRow, 3) = "RETE." Then
  
    Else
       If fgImp.TextMatrix(pnRow, 3) = "IGV" And fgImp.TextMatrix(pnRow, 1) = 1 Then
          cboDocDestino.ListIndex = -1
          cboDocDestino.Enabled = False
       End If
    End If
End If
If fgDetalle.TextMatrix(1, 0) <> "" Then
    CalculaTotal
End If
End Sub

Private Sub fgImp_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim columnas() As String
    columnas = Split(fgImp.ColumnasAEditar, "-")
    If UCase(columnas(pnCol)) = "X" Then 'No editar
        MsgBox "La columna es no editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not OK Then
       If MsgBox(" ¿ Seguro que desea Salir sin grabar Operación ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
          Cancel = 1
          Exit Sub
       End If
    End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub fgDetalle_DblClick()
    If fgDetalle.TextMatrix(fgDetalle.row, 0) = "" Then
       Exit Sub
    End If
    
    If fgDetalle.Col = 9 Then
        If fgDetalle.TextMatrix(fgDetalle.row, 9) = "" Then
            fgDetalle.TextMatrix(fgDetalle.row, 9) = "."
            Set fgDetalle.CellPicture = picCuadroSi
        Else
            fgDetalle.TextMatrix(fgDetalle.row, 9) = ""
            Set fgDetalle.CellPicture = picCuadroNo
        End If
        Exit Sub
    End If
'JEOM
'Se deshabilito para poder modificar los Bienes de Almacen a su cuenta original

'    If fgDetalle.Col = 1 And fgDetalle.TextMatrix(fgDetalle.Row, 10) = "B" Then
'        MsgBox "No se permite modificación de Cuenta por estar relacionada con Bienes de Logística", vbInformation, "¡Aviso!"
'        Exit Sub
'    End If

    If fgDetalle.Col = 1 Then 'And fgDetalle.TextMatrix(fgDetalle.Row, 10) <> "B" Then
       EnfocaTexto txtObj, 0, fgDetalle
    End If
    If fgDetalle.Col = 7 And lbRegulaPend And fgDetalle.row = 1 And lPendPasivo Then
        MsgBox "Importe de regularización debe definirse en la opción previa", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    Call ModidicarTotalPorItem
    If fgDetalle.Col = 7 Or fgDetalle.Col > 12 Then
       If Val(Format(fgDetalle.Text, gsFormatoNumeroDato)) > 0 Then
          EnfocaTexto txtCant, 0, fgDetalle
       End If
    End If
End Sub
'ALPA 20091222*****************************
Private Sub ModidicarTotalPorItem()
    Dim Ip As Integer
    If nCantAgeSel > 0 Then
      If Val(Format(fgDetalle.Text, gsFormatoNumeroDato)) > 0 Then
        For Ip = 1 To nCantAgeSel
            If fgObj.TextMatrix(fgDetalle.row, 5) = sMatrizDatos(1, Ip) And sMatrizDatos(4, Ip) = 0 Then
                fgDetalle.TextMatrix(fgDetalle.row, 7) = fgDetalle.TextMatrix(fgDetalle.row, 7) '* (sMatrizDatos(3, Ip) / 100)
                sMatrizDatos(4, Ip) = 1
            End If
        Next Ip
      End If
    End If
End Sub
 '*******************************************
Private Sub fgDetalle_GotFocus()
VertxtObj
End Sub

Private Sub VertxtObj()
If txtObj.Visible Then
   txtObj.Visible = False
   cmdExaminar.Visible = False
End If
End Sub

Private Sub fgDetalle_KeyPress(KeyAscii As Integer)
If fgDetalle.TextMatrix(fgDetalle.row, 0) = "" Then
   Exit Sub
End If
'If fgDetalle.Col = 1 And fgDetalle.TextMatrix(fgDetalle.Row, 10) = "B" Then
'    MsgBox "No se permite modificación de Cuenta por estar relacionada con Bienes de Logística", vbInformation, "¡Aviso!"
'    Exit Sub
'End If

If fgDetalle.Col = 1 Then
   If KeyAscii = 13 Then EnfocaTexto txtObj, IIf(KeyAscii = 13, 0, KeyAscii), fgDetalle
End If

If fgDetalle.Col = 7 And lbRegulaPend And fgDetalle.row = 1 And lPendPasivo Then
    MsgBox "Importe de regularización debe definirse en la opción previa", vbInformation, "¡Aviso!"
    Exit Sub
End If
If fgDetalle.Col = 7 Or fgDetalle.Col > 12 Then
   If fgDetalle.Col > 12 And fgDetalle.Text = "" Then
      Exit Sub
   End If
   If InStr("-0123456789.", Chr(KeyAscii)) > 0 Then
      EnfocaTexto txtCant, KeyAscii, fgDetalle
   Else
      If KeyAscii = 13 Then EnfocaTexto txtCant, 0, fgDetalle
   End If
End If
Call ModidicarTotalPorItem
End Sub

Private Sub fgDetalle_KeyUp(KeyCode As Integer, Shift As Integer)
If fgDetalle.Col > 12 Then
   If fgImp.TextMatrix(fgDetalle.Col - 12, 0) = "." Then
      Flex_PresionaKey fgDetalle, KeyCode, Shift
      CalculaTotal False
   End If
End If
End Sub

Private Sub fgDetalle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   Select Case fgDetalle.Col
          Case 1:
          Case 4: 'mnuNoAtender.Enabled = False
                  'mnuAtender.Enabled = True
                  'PopupMenu mnuLog
          Case 5: 'mnuAtender.Enabled = False
                  'mnuNoAtender.Enabled = True
                  'PopupMenu mnuLog
   End Select
End If
End Sub

Private Sub Form_Activate()
If lSalir Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
    Dim n As Integer, nSaldo As Currency, nCant As Currency
    Dim nItem As Integer
    Dim sCtaCnt As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oDoc As DOperacion
    Set oDoc = New DOperacion
    Set oContFunct = New NContFunciones
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
        
    'ALPA 20090302****************************
    nTipoRetenciones = 0
    lnPenalidad = 0
    '*****************************************
    Me.cboDocDestino.ListIndex = oConst.LeeConstSistema(gConstSistDestinoIGVDefecto)
    CentraForm Me
    
    lnCalBaseImpIGV = oContFunct.GetImptoIGV
    
    'EJVG20140727 ***
    fsCtaContAporteAFP = oConst.LeeConstSistema(483)
    fsCtaContSeguroAFP = oConst.LeeConstSistema(484)
    fsCtaContComisionAFP = oConst.LeeConstSistema(485)
    fsCtaContAporteONP = oConst.LeeConstSistema(486)
    'END EJVG *******
    
    lSalir = False
    oCon.AbreConexion
    If Mid(gsOpeCod, 3, 1) = "2" Then  'Identificación de Tipo de Moneda
       gsSimbolo = gcME
       If gnTipCambio = 0 Then
          If Not GetTipCambio(gdFecSis) Then
             lSalir = True
             Exit Sub
          End If
       End If
       FrameTipCambio.Visible = True
       txtTipFijo = Format(gnTipCambio, gsFormatoNumeroView3Dec)
       'txtTipVariable = Format(gnTipCambioPonderadoVenta, gsFormatoNumeroView3Dec)
       If gbBitTCPonderado Then
            Label5.Caption = "Ponder."
            txtTipVariable = Format(gnTipCambioPonderadoVenta, gsFormatoNumeroView3Dec)
       Else
            txtTipVariable = Format(gnTipCambioV, gsFormatoNumeroView3Dec)
       End If
    Else
       gsSimbolo = gcMN
    End If
    
    Set rs = oDoc.CargaOpeCta(gsOpeCod, "H", "0")
    If rs.EOF Then
       MsgBox "Falta definir Cuenta de Provisión en Operación", vbInformation, "¡Aviso!"
       lSalir = True
       Exit Sub
    End If
    Do While Not rs.EOF
       cboProvis.AddItem rs!cCtaContDesc & space(100) & rs!cCtaContCod
       rs.MoveNext
    Loop
    RSClose rs
    'cboProvis.ListIndex = cboProvis.ListCount - 1
    If gsOpeCod = "701406" Or gsOpeCod = "702406" Or gsOpeCod = "701403" Or gsOpeCod = "702403" Then
        cboProvis.ListIndex = IndiceListaCombo(Me.cboProvis, IIf(Mid(gsOpeCod, 3, 1) = "1", "25160202", "25260202")) 'PASI20160524 agrego IIF
    Else
        cboProvis.ListIndex = cboProvis.ListCount - 1
    End If
    
'   '--Detracción por IGV 14%
    Set rs = New ADODB.Recordset
    Set rs = oDoc.CargaOpeCta(gsOpeCod, "H", "5")
    If Not rs.EOF Then
       lsCtaDetraccion = rs!cCtaContCod
       cmdDetraccion.Enabled = True
'       Dim clsImp As New DImpuesto
'       Set rs = clsImp.CargaImpuesto(lsCtaDetraccion)
'       If Not rs.EOF Then
'          lnTasaDetraccion = rs!nImpTasa
'       Else
'          MsgBox "Cuenta de Detracción " & lsCtaDetraccion & " No definida como Impuesto", vbInformation, "¡Aviso!"
'          lSalir = True
'          Exit Sub
'       End If
    Else
       MsgBox "No se definio Cuenta de Detraccion", vbInformation, "¡Aviso!"
       cmdDetraccion.Enabled = False
    End If
    
    
    lnColorBien = "&H00F0FFFF"
    lnColorServ = "&H00FFFFC0"
    
    ' Defino el Nro de Movimiento
    txtOpeCod = gsOpeCod
    txtFecha = Format(gdFecSis, "dd/mm/yyyy")
    If lbBienes Then
       fgDetalle.BackColor = lnColorBien
       fgDetalle.BackColorBkg = lnColorBien
    Else
       fgDetalle.BackColor = lnColorServ
       fgDetalle.BackColorBkg = lnColorServ
    End If
    If Trim(sProvCod) <> "" And Not lPendiente Then
       lnMovNro = gnMovNro
       If Trim(sProvCod) <> "" Then
            sSql = " Select cPersNombre cNomPers, ISNULL(cPersIDNro,'00000000') cProvRuc from Persona PE  Left Join PersID PID On PE.cPersCod = PID.cPersCod And Pid.cPersIDTpo = '2' Where PE.cPersCod = '" & sProvCod & "' "
            Set rs = oCon.CargaRecordSet(sSql)
            If RSVacio(rs) Then
               MsgBox "Proveedor no registrado. Por favor verificar", vbCritical, "Error"
               lSalir = True
               Exit Sub
            End If
            txtBuscarProv.Text = Trim(rs!cProvRuc)
            txtBuscarProv_EmiteDatos
            txtBuscarProv.psCodigoPersona = sProvCod
            txtBuscarProv.psDescripcion = rs!cNomPers
            lblProvNombre = rs!cNomPers
            txtBuscarProv.Enabled = False
       End If
       txtOCNro.Tag = gnDocTpo
       txtOCNro = gsDocNro
       txtOCFecha = gdFecha
       txtMovDesc = gsGlosa
       
       Dim rsObj    As ADODB.Recordset
       
       sSql = " SELECT  mc.nMovNro, mc.nMovItem, mo.nMovObjOrden cMovObjOrden, mc.cCtaContCod, " _
            & " mo.cObjetoCod, mcd.cDescrip, moc.nMovCant," & IIf(gsSimbolo = gcMN, "mc.nMovImporte ", "me.nMovMeImporte ") & " nMovImporte," _
            & " ISNULL(sust.nMontoAtendido,0) as nMontoAtendido," _
            & " ISNULL(Sust.nCantAtendido, 0) As nCantAtendido" _
            & " FROM MovCta mc " & IIf(gsSimbolo = gcMN, "", " JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem") _
            & " LEFT JOIN ( SELECT nMovNro,nMovItem, nMovObjOrden, cObjetoCod FROM MovObj " _
            & "             WHERE nMovNro = '" & gnMovNro & "') mo on mo.nMovNro = mc.nMovNro and mo.nMovItem = mc.nMovItem" _
            & " LEFT JOIN MovCant moc ON moc.nMovNro = mo.nMovNro and moc.nMovItem = mo.nMovItem" _
            & " LEFT JOIN MovCotizacDet mcd ON mcd.nMovNro = mc.nMovNro and mcd.nMovItem = mc.nMovItem" _
            & " LEFT JOIN (SELECT mr.nMovNroRef, mc.nMovItem, mo.nMovObjOrden, SUM(nMov" & IIf(gsSimbolo = gcME, "ME", "") & "Importe + ISNULL(MOI.NMOVOTROIMPORTE,0)) as nMontoAtendido, SUM(moc.nMovCant) as nCantAtendido " _
            & "                      FROM MovRef mr JOIN MovCta mc ON mc.nMovNro = mr.nMovNro " & IIf(gsSimbolo = gcME, " JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
            & "                                     JOIN MovObj mo ON mo.nMovNro = mc.nMovNro and mo.nMovItem = mc.nMovItem " _
            & "                  LEFT JOIN (SELECT NMOVNRO, NMOVITEM, SUM(NMOVOTROIMPORTE) NMOVOTROIMPORTE " _
            & "                             FROM MovOTROSITEM MOI where CMOVOTROTEXTO IN ('1','2') GROUP BY NMOVNRO, NMOVITEM " _
            & "                            ) MOI ON MOI.NMOVNRO = MC.NMOVNRO AND MOI.NMOVITEM = MC.NMOVITEM " _
            & "                                LEFT JOIN MovCant moc ON moc.nMovNro = mo.nMovNro and moc.nMovItem = mo.nMovItem " _
            & "                                     JOIN Mov m ON m.nMovNro = mr.nMovNro " _
            & "                      WHERE m.nMovEstado = '" & gMovEstContabMovContable & "' and m.nMovFlag <> '" & gMovFlagEliminado & "' " _
            & "                      GROUP BY mr.nMovNroRef, mc.nMovItem, mo.nMovObjOrden " _
            & "                    ) Sust ON Sust.nMovNroRef = '" & gnMovNro & "' and mo.nMovNro = " & gnMovNro & " and convert(int,Sust.nMovItem) = mo.nMovItem and convert(int,Sust.nMovObjOrden) = mo.nMovObjOrden " _
            & " WHERE mc.nMovNro = '" & gnMovNro & "' and mc.nMovImporte <> 0 and not mc.cctacontcod LIKE '25%' ORDER BY mc.nMovItem "
       
       
       Set rs = oCon.CargaRecordSet(sSql)
       n = 0
       nItem = 0
       Do While Not rs.EOF
          If nItem <> rs!nMovItem Then
             n = n + 1
             If n > fgDetalle.Rows - 1 Then
                AdicionaRow fgDetalle
             End If
             fgDetalle.TextMatrix(n, 0) = n
             fgDetalle.TextMatrix(n, 1) = rs!cCtaContCod
             fgDetalle.TextMatrix(n, 7) = Format(rs!nMovImporte - rs!nMontoAtendido, gsFormatoNumeroView)
             fgDetalle.TextMatrix(n, 8) = rs!cCtaContCod
             fgDetalle.TextMatrix(n, 12) = fgDetalle.TextMatrix(n, 7)

             If Not IsNull(rs!cDescrip) Then
                fgDetalle.TextMatrix(n, 2) = rs!cDescrip
             End If
          End If
          If Not IsNull(rs!cObjetoCod) Then
             Select Case rs!cObjetoCod
                Case ObjCMACAgenciaArea
                   sSql = "SELECT nMovObjOrden, " & ObjCMACAgenciaArea & " cObjPadre, aa.cAreaCod+aa.cAgeCod cObjetoCod, ISNULL(a.cAreaDescripcion,'') + ISNULL(ag.cAgeDescripcion,'') cObjetoDesc, NULL nMovCant " _
                        & "FROM MovObjAreaAgencia mo JOIN AreaAgencia aa ON (mo.cAreaCod = aa.cAreaCod and mo.cAgeCod = aa.cAgeCod) " _
                        & "    LEFT JOIN Areas a ON a.cAreaCod = aa.cAreaCod LEFT JOIN Agencias ag ON ag.cAgeCod = aa.cAgeCod " _
                        & "WHERE  mo.nMovNro = " & gnMovNro & " and mo.nMovItem = " & rs!nMovItem
    
                Case ObjBienesServicios
                   sSql = "SELECT bs.nMovBsOrden, " & ObjBienesServicios & " cObjPadre, bs.cBSCod cObjetoCod, b.cBSDescripcion cObjetoDesc, mc.nMovCant " _
                        & "FROM MovBS bs JOIN MovCant mc ON mc.nMovNro = bs.nMovNro and mc.nMovItem = bs.nMovItem " _
                        & "     JOIN BienesServicios b ON b.cBSCod = bs.cBsCod " _
                        & "WHERE  bs.nMovNro = " & gnMovNro & " and bs.nMovItem = " & rs!nMovItem
                Case Else
                   sSql = "SELECT nMovObjOrden, mo.cObjetoCod cObjPadre, mo.cObjetoCod, o.cObjetoDesc, NULL nMovCant " _
                        & "FROM MovObj mo JOIN  Objeto o ON o.cObjetoCod = mo.cObjetoCod " _
                        & "WHERE  mo.nMovNro = " & gnMovNro & " and mo.nMovItem = " & rs!nMovItem & " and not mo.cObjetoCod IN ('13','11','12','18') "
             End Select
             Set rsObj = oCon.CargaRecordSet(sSql)
             If Not rsObj.EOF Then
                If fgDetalle.TextMatrix(n, 2) = "" Then
                   fgDetalle.TextMatrix(n, 2) = rsObj!cObjetoDesc
                End If
                nItem = rs!nMovItem
                fgDetalle.TextMatrix(n, 1) = rs!cCtaContCod
                
                If Not IsNull(rs!nMovCant) Then
                    fgDetalle.TextMatrix(n, 10) = "B"
                    fgDetalle.TextMatrix(n, 8) = rs!cObjetoCod
                Else
                    fgDetalle.TextMatrix(n, 10) = "S"
                End If
                'FlexBackColor fgDetalle, N, lnColorServ
                Do While Not rsObj.EOF
                   If Not IsNull(rsObj!cObjetoCod) Then
                      fgObj.AdicionaFila
                      fgObj.TextMatrix(fgObj.row, 0) = n
                      fgObj.TextMatrix(fgObj.row, 1) = rs!cMovObjOrden
                      fgObj.TextMatrix(fgObj.row, 2) = rsObj!cObjetoCod
                      fgObj.TextMatrix(fgObj.row, 3) = rsObj!cObjetoDesc
                      fgObj.TextMatrix(fgObj.row, 6) = rsObj!cObjPadre
                   End If
                   rsObj.MoveNext
                Loop
                rs.MoveNext
             Else
                rs.MoveNext
             End If
          Else
             rs.MoveNext
          End If
       Loop
       sSql = "SELECT dMovPlazo, cMovLugarEntrega FROM MovCotizac WHERE nMovNro = '" & gnMovNro & "'"
       Set rs = oCon.CargaRecordSet(sSql)
       If Not rs.EOF Then
          txtOCPlazo = Format(rs!dMovPlazo, "dd/mm/yyyy")
          txtOCEntrega = rs!cMovLugarEntrega
       End If
       RSClose rs
       cmdAjuste.Enabled = False
       cmdAgregar.Enabled = False
       cmdEliminar.Enabled = False
    End If
    'EJVG20131113 ***
    If gsOpeCod = gContProvLogComprobanteMN Or gsOpeCod = gContProvLogComprobanteME Then
        ActualizaComprobantes
        fraComprobante.Visible = True
        cmdAgregar.Enabled = False
        cmdEliminar.Enabled = False
        txtBuscarProv.Enabled = False
    End If
    'END EJVG *******
    Me.Caption = gsOpeDesc
    FormatoOrden
       
    'GUIA DE REMISION
    txtGRNro.Tag = TpoDocGuiaRemision
    
    If lbBienes Then
    Else
       TabDocRef.TabCaption(0) = "Orden de Servicio"
       'lblDoc.Visible = False
    End If
    fgObj.BackColor = lnColorServ
    
    'Tipos de Comprobantes de Pago
    Set rs = oDoc.CargaOpeDoc(gsOpeCod, , OpeDocMetDigitado)
    Do While Not rs.EOF
       cboDoc.AddItem Format(rs!nDocTpo, "00") & " " & Mid(rs!cDocDesc & space(100), 1, 100) & Mid(rs!cDocAbrev & "   ", 1, 3)
       rs.MoveNext
    Loop
    If cboDoc.ListCount = 1 Then
       cboDoc.ListIndex = 0
    End If
    RSClose rs
    Call CalculaTotal
    
lnImporteRegula = 0
If lPendiente Then
    AdicionaRow fgDetalle
    If lbRegulaPend And lPendPasivo Then
        fgDetalle.TextMatrix(1, 1) = frmAnalisisRegulaPend.txtCtaPend
        fgDetalle.TextMatrix(1, 2) = frmAnalisisRegulaPend.txtCtaPendDes
        fgDetalle.TextMatrix(1, 7) = Format(Abs(gnImporte), gsFormatoNumeroView)
        fgDetalle.TextMatrix(1, 12) = fgDetalle.TextMatrix(1, 7)
        txtSTotal = fgDetalle.TextMatrix(1, 7)
        txtTotal = txtSTotal
        fgDetalle.Enabled = True
        cmdAjuste.Enabled = True
        cmdAgregar.Enabled = True
        cmdEliminar.Enabled = True
    Else
        lnImporteRegula = Abs(gnImporte)
    End If
End If
'ALPA 20090303******************
If lbRetenciones Then
    cmdRetenciones.Enabled = True
Else
    cmdRetenciones.Enabled = False
End If
'*******************************
End Sub

Private Sub Form_Unload(Cancel As Integer)
lnMovNro = 0
End Sub


Private Sub TabDocRef_Click(PreviousTab As Integer)
Select Case TabDocRef.Tab
       Case 0
            If txtOCNro.Enabled Then
               txtOCNro.SetFocus
            End If
            ControlesTabDocRef False, False
       Case 1
            ControlesTabDocRef True, False
            txtGRSerie.SetFocus
       Case 2
            ControlesTabDocRef False, True
            cboDoc.SetFocus
End Select
End Sub

Private Sub txtAjuste_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtAjuste, KeyAscii, 12, 2)
If KeyAscii = 13 Then
    cmdAplicar.SetFocus
End If
End Sub

Private Sub txtBuscarProv_EmiteDatos()
    Dim oProv As New DLogProveedor 'NAGL Según INC1712260008 Agregó New
    Dim lsMotivoNoHabido As String
    Set oProv = New DLogProveedor
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oPer As DPersonas
    Set oPer = New DPersonas
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Set rs3 = New ADODB.Recordset
    
    Dim lnRuc As String
    Dim psIdProv As String 'NAGL Según INC1712260008
    Dim rsDir As New ADODB.Recordset 'NAGL Según INC1712260008
   
    psIdProv = "" 'NAGL Según INC1712260008
    lbNewProv = False
    lblProvNombre = txtBuscarProv.psDescripcion
  

If gsOpeCod = 701401 Or gsOpeCod = 702401 Or gsOpeCod = 701402 Or gsOpeCod = 702402 Then

Else
' Busqueda de RUC o DNI  --------------------
    Set rs1 = oProv.GetProveedorRUC(txtBuscarProv.psCodigoPersona)
    If Not (rs1.EOF And rs1.BOF) Then
        lnRuc = rs1!cPersIdNro
        If oProv.GetProveedorNoHabido(lnRuc, lsMotivoNoHabido) Then
           MsgBox "Proveedor fue identificado como no Habido por la Sunat" & Chr(10) _
                  & "Con la siguiente Observación : " & lsMotivoNoHabido, vbInformation, "Aviso"
           txtBuscarProv = ""
           lblProvNombre = ""
           Exit Sub
        End If
    Else
        '***Modificado por ELRO el 20120428, según OYP-RFC038-2012
        '*** PEAC 20110222
        'If gsOpeCod = "701403" Or gsOpeCod = "702403" Then
        '    MsgBox "Este Proveedor no tiene RUC, por favor comuníquese con Contabilidad para actualizar sus datos.", vbOKOnly + vbInformation, "Aviso"
        '   txtBuscarProv = ""
        '   lblProvNombre = ""
        '    Exit Sub
        'End If
        '*** FIN PEAC
        If gsOpeCod = "701403" Or gsOpeCod = "702403" Then
            If Trim(Left(cboDoc, 3)) <> "95" And Trim(Left(cboDoc, 3)) <> "96" Then
                MsgBox "Este Proveedor no tiene RUC, por favor comuníquese con Contabilidad para actualizar sus datos.", vbOKOnly + vbInformation, "Aviso"
                txtBuscarProv = ""
                lblProvNombre = ""
                Exit Sub
            End If
        End If
        '***Fin Modificado por ELRO*******************************

       Set rs2 = oProv.GetProveedorDNI(txtBuscarProv.psCodigoPersona)
       If Not (rs2.EOF And rs2.BOF) Then
          txtBuscarProv.Text = rs2!cPersIdNro
       Else
          'Busqueda Carnet de extranjeria
          Set rs3 = oProv.GetProveedorCarnetExt(txtBuscarProv.psCodigoPersona)
          If Not (rs3.EOF And rs3.BOF) Then
             txtBuscarProv.Text = rs3!cPersIdNro
          Else
             MsgBox "Proveedor no tiene DNI ni RUC", vbInformation, "AVISO"
          End If
       End If
    End If
'--------------------------------------------
End If
    
    If lblProvNombre <> "" Then
        lbNewProv = Not oProv.IsExisProveedor(txtBuscarProv.psCodigoPersona)
        Set rs = oProv.GetProveedorAgeRetBuenCont(txtBuscarProv.psCodigoPersona)
        If rs.EOF And rs.BOF Then
            MsgBox "La persona ingresada no esta registrada como proveedor o tiene el estado de Desactivado, debe regsitrarlo o activarlo.", vbInformation, "Aviso"
            Me.chkBuneCOnt.value = 0
            Me.chkRetencion.value = 0
        Else
            Me.chkBuneCOnt.value = IIf(rs.Fields(1), 1, 0)
            Me.chkRetencion.value = IIf(rs.Fields(0), 1, 0)
        End If
        
        txtMovDesc.SetFocus
    
    End If
    Set oProv = Nothing
    'EJVG20140723 ***
    If txtBuscarProv.psCodigoPersona <> "" Then
        If gbBitRetencSistPensProv Then
            If Val(Trim(Left(cboDoc.Text, 3))) = TpoDoc.TpoDocRecHonorarios Then
                Call CalculaTotal
            End If
        End If
    End If
     'END EJVG *******
     '******************NAGL Según INC1712260008***************
             Set rsDir = oProv.GetProvDirector(gdFecSis)
             If (rsDir.RecordCount <> 0) Then
                If Not rsDir.BOF And Not rsDir.EOF Then
                   Do While Not rsDir.EOF
                      If txtBuscarProv.Text = rsDir!cPersIdNro Then
                          psIdProv = rsDir!cPersIdNro
                      End If
                      rsDir.MoveNext
                   Loop
                End If
            End If
    '**********END NAGL 2171230*******************************
    
    'ALPA 20160428***
    If gsOpeCod = "701406" Or gsOpeCod = "702406" Or gsOpeCod = "701403" Or gsOpeCod = "702403" Then
        If (txtBuscarProv.Text = "1090100942715" Or txtBuscarProv.Text = "20148116564") Then  'PASI20160524
            cboProvis.ListIndex = IndiceListaCombo(Me.cboProvis, IIf(Mid(gsOpeCod, 3, 1) = "1", "251702", "252702")) 'PASI20160524 Agrego IIF
        ElseIf psIdProv <> "" Then
            cboProvis.ListIndex = IndiceListaCombo(Me.cboProvis, IIf(Mid(gsOpeCod, 3, 1) = "1", "251506", "252506")) ' NAGL INC1712260008
        Else
            cboProvis.ListIndex = IndiceListaCombo(Me.cboProvis, IIf(Mid(gsOpeCod, 3, 1) = "1", "25160202", "25260202")) 'PASI20160524 Agrego IIF
        End If
    End If
    '****************
End Sub

Private Sub txtBuscarProv_Validate(Cancel As Boolean)
If (txtBuscarProv = "" And txtBuscarProv.psDescripcion = "") Then
   Cancel = True
End If
End Sub

Private Sub txtBuscarProvRef_EmiteDatos()
    Dim oProv As DLogProveedor
    Dim lsMotivoNoHabido As String
    Set oProv = New DLogProveedor
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oPer As DPersonas
    Set oPer = New DPersonas
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    
    Dim lnRuc As String
    
    lbNewProv = False
    
    lblProvNombreRef = txtBuscarProvRef.psDescripcion
    
' Busqueda de RUC o DNI  --------------------
    Set rs1 = oProv.GetProveedorRUC(txtBuscarProvRef.psCodigoPersona)
    If Not (rs1.EOF And rs1.BOF) Then
        lnRuc = rs1!cPersIdNro
'        If oProv.GetProveedorNoHabido(lnRuc, lsMotivoNoHabido) Then
'           MsgBox "Proveedor fue identificado como no Habido por la Sunat" & Chr(10) _
'                  & "Con la siguiente Observación : " & lsMotivoNoHabido, vbInformation, "Aviso"
'           txtBuscarProv = ""
'           lblProvNombre = ""
'           Exit Sub
'        End If
    Else
       Set rs2 = oProv.GetProveedorDNI(txtBuscarProvRef.psCodigoPersona)
       If Not (rs2.EOF And rs2.BOF) Then
          txtBuscarProvRef.Text = rs2!cPersIdNro
       Else
          MsgBox "Proveedor no tiene DNI ni RUC", vbInformation, "AVISO"
          txtBuscarProvRef.Text = ""
       End If
    End If
'--------------------------------------------
    
    If lblProvNombreRef <> "" Then
        lbNewProv = Not oProv.IsExisProveedor(txtBuscarProvRef.psCodigoPersona)
        Set rs = oProv.GetProveedorAgeRetBuenCont(txtBuscarProvRef.psCodigoPersona)
        If rs.EOF And rs.BOF Then
            'MsgBox "La persona ingresada no esta registrada como proveedor o tiene el estado de Desactivado, debe regsitrarlo o activarlo.", vbInformation, "Aviso"
            Me.chkBuneCOnt.value = 0
            Me.chkRetencion.value = 0
        Else
            Me.chkBuneCOnt.value = IIf(rs.Fields(1), 1, 0)
            Me.chkRetencion.value = IIf(rs.Fields(0), 1, 0)
        End If
        
        txtMovDesc.SetFocus
    
    End If
    Set oProv = Nothing




End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
Dim lnImporte As Currency
Dim n As Integer
KeyAscii = NumerosDecimales(txtCant, KeyAscii, 12, 2)
If KeyAscii = 13 Then
   If fgDetalle.Col = 4 Then
      If nVal(txtCant) > nVal(fgDetalle.TextMatrix(txtCant.Tag, 11)) Then
         MsgBox "Cantidad Atendida no puede mayor que Cantidad Solicitada", vbInformation, "¡Aviso!"
         Exit Sub
      End If
      fgDetalle.TextMatrix(txtCant.Tag, 7) = Format(Round(Val(txtCant) * Val(Format(fgDetalle.TextMatrix(fgDetalle.row, 5), "#0.00")), 2), gsFormatoNumeroView)
   End If
   If fgDetalle.Col = 7 And fgDetalle.TextMatrix(fgDetalle.row, 12) <> "" Then
      lnImporte = nVal(txtCant)
      If fgDetalle.Cols > 12 Then
         For n = 13 To fgDetalle.Cols - 1
            lnImporte = lnImporte + (nVal(fgDetalle.TextMatrix(fgDetalle.row, n)) * IIf(fgImp.TextMatrix(n - 12, 8) = "D", 1, -1))
         Next
      End If
      If lnImporte > nVal(fgDetalle.TextMatrix(txtCant.Tag, 12)) Then
         MsgBox "Monto no puede ser mayor a Monto pactado", vbInformation, "¡Aviso!"
         Exit Sub
      End If
   End If
   fgDetalle.Text = Format(txtCant.Text, gsFormatoNumeroView)
   txtCant.Visible = False
   fgDetalle.SetFocus
   Call ModidicarTotalPorItem
   CalculaTotal IIf(fgDetalle.Col > 12, False, True), False
End If
End Sub

Private Sub txtCant_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
   txtCant_KeyPress 13
   SendKeys IIf(KeyCode = 38, "{Up}", "{Down}"), True
End If
End Sub

Private Sub txtCant_LostFocus()
txtCant.Text = ""
txtCant.Visible = False
End Sub
'EJVG20131113 ***
Private Sub txtComprobanteCod_EmiteDatos()
    Dim lsDocNro As String
    Dim lnMovNro As Long
    Dim frmSel As New frmLogProvisionComprobanteSel
    cmdComprobanteLimpiar_click
    frmSel.Inicio fRsComprobante, lsDocNro, lnMovNro
    If lsDocNro <> "" And lnMovNro > 0 Then
        fsDocOrigenNro = lsDocNro
        fnComprobanteMovNro = lnMovNro
        txtComprobanteCod.Text = fsDocOrigenNro
        cmdComprobanteCargar.SetFocus
    End If
    Set frmSel = Nothing
End Sub
Private Sub cmdComprobanteCargar_click()
    On Error GoTo cmdComprobanteCargar
    Dim oLog As New DLogGeneral
    Dim rs As New ADODB.Recordset
    Dim rsObj As New ADODB.Recordset
    Dim i As Integer
    Dim sSql As String
    Dim oCon As DConecta
    
    If txtComprobanteCod.Text = "" Or fsDocOrigenNro = "" Or fnComprobanteMovNro <= 0 Then
        MsgBox "Ud. debe seleccionar primero el Comprobante", vbInformation, "Aviso"
        txtComprobanteCod.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = 11
    lnMovNro = fnComprobanteMovNro
    Set oLog = New DLogGeneral
    Set rs = New ADODB.Recordset
    Set rs = oLog.ComprobanteDet(fnComprobanteMovNro, Mid(gsOpeCod, 3, 1))
    If Not rs.EOF Then
        txtBuscarProv.Text = rs!cRuc
        lblProvNombre.Text = rs!cPersNombre
        txtBuscarProv.psCodigoPersona = rs!cPersCod
        sProvCod = rs!cPersCod 'EJVG20131212
        txtBuscarProv.psDescripcion = rs!cPersNombre
        txtBuscarProv_EmiteDatos
        txtMovDesc.Text = rs!cDescripcion
        cboDoc.ListIndex = ListaIndiceComprobante(rs!nDocTpo)
        txtFacSerie.Text = rs!cDocSerie
        txtFacNro.Text = rs!cDocNro
        txtFacFecha.Text = Format(rs!dDocFecha, gsFormatoFechaView)
        cboProvis.ListIndex = IndiceListaCombo(Me.cboProvis, rs!cCtaContCod2) 'EJVG20131213
        'PASI 20160428***
        If gsOpeCod = "701406" Or gsOpeCod = "702406" Or gsOpeCod = "701403" Or gsOpeCod = "702403" Then
            If (txtBuscarProv.Text = "1090100942715" Or txtBuscarProv.Text = "20148116564") Then
                cboProvis.ListIndex = IndiceListaCombo(Me.cboProvis, IIf(Mid(gsOpeCod, 3, 1) = "1", "251702", "252702"))
            Else
                cboProvis.ListIndex = IndiceListaCombo(Me.cboProvis, IIf(Mid(gsOpeCod, 3, 1) = "1", "25160202", "25260202"))
            End If
        End If
        '**********
        i = 0
        Do While Not rs.EOF
            i = i + 1
            If i > fgDetalle.Rows - 1 Then
               AdicionaRow fgDetalle
            End If
            fgDetalle.TextMatrix(i, 0) = i
            fgDetalle.TextMatrix(i, 1) = rs!cCtaContCod
            fgDetalle.TextMatrix(i, 7) = Format(rs!nMovImporte, gsFormatoNumeroView)
            fgDetalle.TextMatrix(i, 8) = rs!cCtaContCod
            fgDetalle.TextMatrix(i, 2) = rs!cCtaContDesc
            fgDetalle.TextMatrix(i, 12) = fgDetalle.TextMatrix(i, 7)

            If rs!cObjetoCod <> "" Then
                Select Case rs!cObjetoCod
                    Case ObjCMACAgenciaArea
                        sSql = "SELECT nMovObjOrden, " & ObjCMACAgenciaArea & " cObjPadre, aa.cAreaCod+aa.cAgeCod cObjetoCod, ISNULL(a.cAreaDescripcion,'') + ISNULL(ag.cAgeDescripcion,'') cObjetoDesc, NULL nMovCant " _
                            & "FROM MovObjAreaAgencia mo JOIN AreaAgencia aa ON (mo.cAreaCod = aa.cAreaCod and mo.cAgeCod = aa.cAgeCod) " _
                            & "    LEFT JOIN Areas a ON a.cAreaCod = aa.cAreaCod LEFT JOIN Agencias ag ON ag.cAgeCod = aa.cAgeCod " _
                            & "WHERE  mo.nMovNro = " & rs!nMovNro & " and mo.nMovItem = " & rs!nMovItem
                    Case ObjBienesServicios
                        sSql = "SELECT bs.nMovBsOrden, " & ObjBienesServicios & " cObjPadre, bs.cBSCod cObjetoCod, b.cBSDescripcion cObjetoDesc, mc.nMovCant " _
                            & "FROM MovBS bs JOIN MovCant mc ON mc.nMovNro = bs.nMovNro and mc.nMovItem = bs.nMovItem " _
                            & "     JOIN BienesServicios b ON b.cBSCod = bs.cBsCod " _
                            & "WHERE  bs.nMovNro = " & rs!nMovNro & " and bs.nMovItem = " & rs!nMovItem
                    Case Else
                        sSql = "SELECT nMovObjOrden, mo.cObjetoCod cObjPadre, mo.cObjetoCod, o.cObjetoDesc, NULL nMovCant " _
                            & "FROM MovObj mo JOIN  Objeto o ON o.cObjetoCod = mo.cObjetoCod " _
                            & "WHERE  mo.nMovNro = " & rs!nMovNro & " and mo.nMovItem = " & rs!nMovItem & " and not mo.cObjetoCod IN ('13','11','12','18') "
                End Select
                
                Set oCon = New DConecta
                Set rsObj = New ADODB.Recordset
                oCon.AbreConexion
                Set rsObj = oCon.CargaRecordSet(sSql)
                oCon.CierraConexion
                If Not rsObj.EOF Then
                    Do While Not rsObj.EOF
                        If Not IsNull(rsObj!cObjetoCod) Then
                            fgObj.AdicionaFila
                            fgObj.TextMatrix(fgObj.row, 0) = i
                            fgObj.TextMatrix(fgObj.row, 1) = rs!nMovObjOrden
                            fgObj.TextMatrix(fgObj.row, 2) = rsObj!cObjetoCod
                            fgObj.TextMatrix(fgObj.row, 3) = rsObj!cObjetoDesc
                            fgObj.TextMatrix(fgObj.row, 6) = rsObj!cObjPadre
                        End If
                        rsObj.MoveNext
                    Loop
                End If
            End If
            rs.MoveNext
        Loop
        Call CalculaTotal
        'EJVG20140724 ***
        If gbBitRetencSistPensProv Then
            If Val(Trim(Left(cboDoc.Text, 3))) = TpoDoc.TpoDocRecHonorarios Then
                Dim oPSP As New NProveedorSistPens
                Dim frm As frmProveedorRegSistemaPension
                If oPSP.AplicaRetencionSistemaPension(txtBuscarProv.psCodigoPersona, CDate(Me.txtFacFecha.Text), MontoBaseOperacion) Then
                    Do While Not oPSP.ExisteDatosSistemaPension(txtBuscarProv.psCodigoPersona)
                        Screen.MousePointer = 0
                        If MsgBox("Para continuar Ud. debe registrar los datos de Sistema Pensión del Proveedor", vbInformation + vbYesNo, "Aviso") = vbYes Then
                            Set frm = New frmProveedorRegSistemaPension
                            frm.Registrar (txtBuscarProv.psCodigoPersona)
                        Else
                            cmdComprobanteLimpiar_click
                            Set oPSP = Nothing
                            Exit Sub
                        End If
                        Set frm = Nothing
                    Loop
                End If
                Set oPSP = Nothing
            End If
        End If
        Call CalculaTotal
        'END EJVG *******
    Else
        MsgBox "No se encontraron datos, si el problema persiste comunicarse con el Dpto. de TI", vbInformation, "Aviso"
    End If
    
    Screen.MousePointer = 0
    Set rsObj = Nothing
    Set rs = Nothing
    Set oLog = Nothing
    Set oCon = Nothing
    Exit Sub
cmdComprobanteCargar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdComprobanteLimpiar_click()
    txtComprobanteCod.Text = ""
    txtProvCod.Text = ""
    lblProvNombre.Text = ""
    txtMovDesc.Text = ""
    cboDoc.ListIndex = -1
    txtFacSerie.Text = ""
    txtFacNro.Text = ""
    txtFacFecha.Text = Format(gdFecSis, gsFormatoFechaView)
    fgDetalle.Rows = 2
    EliminaRow fgDetalle, 1
    fgObj.Rows = 2
    fgObj.EliminaFila 1
    txtSTotal.Text = "0.00"
    txtTotal.Text = "0.00"
    txtBuscarProv.Text = ""
End Sub
'END EJVG ******
Private Sub txtFacFecha_GotFocus()
'***Agregado por ELRO el 20130611, según SATI INC1304290006****
If fnOK = False Then
    MsgBox "Documento de Proveedor ya registrado", vbInformation, "¡Aviso!"
    txtFacNro.SetFocus
    Exit Sub
End If
'***Fin Agregado por ELRO el 20130611, según SATI INC1304290006
fEnfoque txtFacFecha
End Sub

Private Sub txtFacFecha_KeyPress(KeyAscii As Integer)
Dim i As Integer
    Dim oTC As nTipoCambio
    Set oTC = New nTipoCambio
    
    If KeyAscii = 13 Then
       If Not ValidaFechaContab(txtFacFecha, gdFecSis, False) Then
          Exit Sub
       Else
          If gsSimbolo = gcME Then
             GetTipCambio CDate(txtFacFecha)
             txtTipVariable = Format(gnTipCambioV, gsFormatoNumeroView3Dec)
             If gbBitTCPonderado Then
                  Label5.Caption = "Ponder."
                  txtTipVariable = Format(gnTipCambioPonderadoVenta, gsFormatoNumeroView3Dec)
             Else
                  txtTipVariable = Format(gnTipCambioV, gsFormatoNumeroView3Dec)
             End If
             gnTipCambio = Val(Format(txtTipFijo, gsFormatoNumeroView3Dec))
             GetTipCambio gdFecSis
          End If
          If fgImp.TextMatrix(1, 0) <> "" Then
            Dim oImp As DImpuesto
            Set oImp = New DImpuesto
            For i = 1 To fgImp.Rows - 1
               fgImp.TextMatrix(i, 4) = Format(oImp.CargaImpuestoFechaValor(fgImp.TextMatrix(i, 6), txtFacFecha), gsFormatoNumeroView)
            Next
            Set oImp = Nothing
          End If
          fgDetalle.SetFocus
       End If
    End If
End Sub

Private Sub txtFacFecha_LostFocus()
'***Modificado por ELRO el 20130611, según SATI INC1304290006****
If oContFunct.ValidaDocumento(Val(Left(cboDoc.Text, 3)), txtFacSerie & "-" & txtFacNro, txtBuscarProv.psCodigoPersona) Then
    If Not ValidaFechaContab(txtFacFecha, gdFecSis, False) Then
       Exit Sub
    Else
        If gsSimbolo = gcME Then
           GetTipCambio CDate(txtFacFecha)
           txtTipVariable = Format(gnTipCambioV, gsFormatoNumeroView3Dec)
           If gbBitTCPonderado Then
                Label5.Caption = "Ponder."
                txtTipVariable = Format(gnTipCambioPonderadoVenta, gsFormatoNumeroView3Dec)
           Else
                txtTipVariable = Format(gnTipCambioV, gsFormatoNumeroView3Dec)
           End If
           gnTipCambio = Val(Format(txtTipFijo, gsFormatoNumeroView3Dec))
           GetTipCambio gdFecSis
        End If
        'EJVG20140727 ***
        If gbBitRetencSistPensProv Then
            If Val(Trim(Left(cboDoc.Text, 3))) = TpoDoc.TpoDocRecHonorarios Then
                Call CalculaTotal
            End If
        End If
        'END EJVG *******
    End If
End If
'***Fin Modificado por ELRO el 20130611, según SATI INC1304290006
End Sub

Private Sub txtFacFecha_Validate(Cancel As Boolean)
'    If Not ValidaFechaContab(txtFacFecha, gdFecSis, False) Then
'        Cancel = True
'    Else
'        GetTipCambio CDate(txtFacFecha)
'
'        If gbBitTCPonderado Then
'            Label5.Caption = "Ponder."
'            txtTipVariable = Format(gnTipCambioPonderado, gsFormatoNumeroView)
'        Else
'            txtTipVariable = Format(gnTipCambioV, gsFormatoNumeroView)
'        End If
'
'        gnTipCambio = Val(Format(txtTipFijo, gsFormatoNumeroDato))
'
'        GetTipCambio gdFecSis
'    End If
End Sub
'***Comentado por ELRO el 20130611, según SATI INC1304290006****
'Private Sub txtFacNro_LostFocus()
'txtFacNro = Format(txtFacNro, String(8, "0"))
'End Sub
'***Fin Comentado por ELRO el 20130611, según SATI INC1304290006

Private Sub txtFacNro_Validate(Cancel As Boolean)
    If Not oContFunct.ValidaDocumento(Val(Left(cboDoc.Text, 3)), txtFacSerie & "-" & txtFacNro, txtBuscarProv.psCodigoPersona) Then
        fnOK = False
        Exit Sub
    Else
        fnOK = True
    End If
End Sub

Private Sub txtFacSerie_LostFocus()
    '***Modificado por ELRO el 20130611, según SATI INC1304290006****
    'txtFacSerie = Format(txtFacSerie, String(3, "0"))
    'txtFacSerie = Trim(txtFacSerie)
    '***Fin Modificado por ELRO el 20130611, según SATI INC1304290006
    'If Trim(Left(cboDoc, 2)) = "05" Then
        'txtFacSerie = Trim(Str("3"))
    'Else 'Comentado by NAGL 20170926
        txtFacSerie = Right(String(4, "0") & txtFacSerie, 4)
    'End If '***NAGL ERS012-2017
End Sub
Private Sub txtFecha_GotFocus()
    fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Not ValidaFechaContab(txtFecha, gdFecSis) Then
          Exit Sub
       Else
          If Not PermiteModificarAsiento(Format(txtFecha, gsFormatoMovFecha), True) Then
            Exit Sub
          End If
          GetTipCambio CDate(txtFecha)
          txtTipVariable = Format(gnTipCambioPonderadoVenta, gsFormatoNumeroView3Dec)
          txtTipFijo = Format(gnTipCambio, gsFormatoNumeroView3Dec)
          GetTipCambio gdFecSis
          If txtBuscarProv.Enabled Then
             txtBuscarProv.SetFocus
          Else
             txtMovDesc.SetFocus
          End If
       End If
    End If
End Sub

Private Sub txtFecha_Validate(Cancel As Boolean)
    If Not ValidaFechaContab(txtFecha, gdFecSis) Then
       Cancel = True
    End If
End Sub

Private Sub txtGRFecha_GotFocus()
    fEnfoque txtGRFecha
End Sub

Private Sub txtGRFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Not ValidaFechaContab(txtGRFecha, gdFecSis) Then
          txtGRFecha.SetFocus
       Else
            TabDocRef.Tab = 2
       End If
    End If
End Sub

Private Sub txtGRNro_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
       txtGRNro = Format(txtGRNro, String(8, "0"))
       txtGRFecha.SetFocus
    End If
End Sub

Private Sub txtFacNro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        '***Modificado por ELRO el 20130611, según SATI INC1304290006****
        'txtFacNro = Format(txtFacNro, String(8, "0"))
        If Not oContFunct.ValidaDocumento(Val(Left(cboDoc.Text, 3)), txtFacSerie & "-" & txtFacNro, txtBuscarProv.psCodigoPersona) Then
            MsgBox "Documento de Proveedor ya registrado", vbInformation, "¡Aviso!"
            fnOK = False
            txtFacNro.SetFocus
            Exit Sub
        Else
            txtFacNro = Trim(txtFacNro)
            fnOK = True
            txtFacFecha.SetFocus
        End If
        '***Fin Modificado por ELRO el 20130611, según SATI INC1304290006
    End If
End Sub

Private Sub txtGRSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
       txtGRSerie = Format(txtGRSerie, String(3, "0"))
       txtGRNro.SetFocus
    End If
End Sub

Private Sub txtFacSerie_KeyPress(KeyAscii As Integer)
 KeyAscii = LetrasNumeros(KeyAscii)
    If KeyAscii = 13 Then
        '***Modificado por ELRO el 20130611, según SATI INC1304290006****
        'txtFacSerie = Format(txtFacSerie, String(3, "0"))
        'txtFacSerie = Trim(txtFacSerie)
        '***Fin Modificado por ELRO el 20130611, según SATI INC1304290006
        'If Trim(Left(cboDoc, 2)) <> "05" Then
            txtFacSerie = Right(String(4, "0") & txtFacSerie, 4)
        'Else
            'txtFacSerie = Trim(txtFacSerie)
        'End If 'NAGL ERS012-2017 20170710 Comentado by NAGL 20170927
        txtFacNro.SetFocus
    End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       KeyAscii = 0
       Select Case TabDocRef.Tab
            Case 0: TabDocRef.Tab = 2: cboDoc.SetFocus
            Case 1: txtGRSerie.SetFocus
            Case 2: cboDoc.SetFocus
       End Select
    End If
End Sub

Private Function ValidaDatos() As Boolean
    ValidaDatos = False
    If cboDoc.ListIndex < 0 Then
       MsgBox "No se especifico Tipo de Comprobante de Provisión"
       cboDoc.SetFocus
       Exit Function
    End If
    If txtFacSerie = "" Then
       MsgBox "Falta especificar Serie de Comprobante.", vbInformation, "Aviso"
       txtFacSerie.SetFocus
       Exit Function
    End If
    If txtFacNro = "" Then
       MsgBox "Falta especificar Número de Comprobante", vbInformation, "Aviso"
       txtFacNro.SetFocus
       Exit Function
    End If
    If Not ValidaFechaContab(txtFacFecha, gdFecSis, False) Then
       txtFacFecha.SetFocus
       Exit Function
    End If
    
    '*** PEAC 20110222 - SE AGREGO "Trim(lblProvNombre)"
    If Trim(txtBuscarProv) = "" Or Trim(lblProvNombre) = "" Then
       MsgBox "Falta especificar datos de Proveedor", vbInformation, "Aviso"
       txtBuscarProv.SetFocus
       Exit Function
    End If
    If Val(Format(txtTotal, gsFormatoNumeroDato)) = 0 Then
       MsgBox "Falta indicar Importe de Documento", vbInformation, "Aviso"
       fgDetalle.SetFocus
       Exit Function
    End If
    If txtMovDesc = "" Then
       MsgBox "Falta especificar Motivo de Provisión", vbInformation, "Aviso"
       txtMovDesc.SetFocus
       Exit Function
    End If
    
    If Not oContFunct.ValidaDocumento(Val(Left(cboDoc.Text, 3)), txtFacSerie & "-" & txtFacNro, txtBuscarProv.psCodigoPersona) Then
        MsgBox "Documento de Proveedor ya registrado", vbInformation, "¡Aviso!"
        txtFacSerie.SetFocus
        Exit Function
    End If
    ValidaDatos = True

End Function

Private Sub cmdAceptar_Click()
    Dim n As Integer 'Contador
    Dim nItem As Integer, nCol  As Integer
    Dim sTexto As String, lOk As Boolean
    Dim sMovNro As String
    Dim lnOrdenObj As Integer, m As Integer
    Dim lsMsgError As String
    
    Dim lsAgeCod  As String
    Dim lsAreaCod As String
    Dim lsCodPers As String
    Dim lnImporte As Currency
    Dim lnTotalReten As Currency
    Dim lnI As Integer
    Dim lnA As Integer
    Dim lnMontoNoGrabado As Currency
    Dim rsPend As ADODB.Recordset
    Dim lsTpoCta As String
    Dim lnDocTpo As Integer
    
    '*** PEAC 20110314
    Dim rsCtasOrden As ADODB.Recordset
    Dim lnMontoCtaOrd As Double
    Dim lcMovNroCtaOrd As String, lnMovNroCtaOrd As Long
    Dim nOtrosItem As Integer '*** PEAC 20130212
    'EJVG20131115 ***
    Dim ofrmVistoElect As frmVistoElectronico
    Dim lbVistoOK As Boolean
    'END EJVG *******
    'EJVG20140727 ***
    Dim frm As frmProveedorRegSistemaPension
    Dim bPSPSinDatos As Boolean
    Dim oPSP As NProveedorSistPens
    Dim lnRetProvSPMontoBase As Currency
    Dim lnRetProvSPAporte As Currency
    Dim lnRetProvSPComisionAFP As Currency
    Dim lnRetProvSPSeguroAFP As Currency
    Dim lnTpoSistPensProv As TipoSistemaPensionProveeedor
    'END EJVG *******
    
    On Error GoTo ErrAceptar

    If lnImporteRegula <> 0 And CCur(txtTotal) <> lnImporteRegula Then
        MsgBox "Regularización debe ser por " & Format(lnImporteRegula, gsFormatoNumeroView) & ". Verifique datos ingresados"
        Exit Sub
    End If
    If Not ValidaDatos() Then
        Exit Sub
    End If
    If Not PermiteModificarAsiento(Format(txtFecha, gsFormatoMovFecha), True) Then
        Exit Sub
    End If
    
    If Val(Left(cboDoc, 3)) = 14 Then 'Recibo de Servicios Publicos
        If txtFechaRef = "" Or txtFechaRef = "  /  /    " Then
           MsgBox "Debe Ingresar Fecha/Nro.Ref. de los Servicios Publicos", vbInformation, "Aviso"
           Exit Sub
        End If
    End If
    
    'EJVG20140724 ***
    If gbBitRetencSistPensProv Then
        If Val(Trim(Left(cboDoc.Text, 3))) = TpoDoc.TpoDocRecHonorarios Then
            Set oPSP = New NProveedorSistPens
            lnRetProvSPMontoBase = MontoBaseOperacion
            If oPSP.AplicaRetencionSistemaPension(txtBuscarProv.psCodigoPersona, CDate(txtFacFecha.Text), lnRetProvSPMontoBase) Then
                Do While Not oPSP.ExisteDatosSistemaPension(txtBuscarProv.psCodigoPersona)
                    bPSPSinDatos = True
                    If MsgBox("Para continuar Ud. debe registrar los datos de Sistema Pensión del Proveedor", vbInformation + vbYesNo, "Aviso") = vbYes Then
                        Set frm = New frmProveedorRegSistemaPension
                        frm.Registrar (txtBuscarProv.psCodigoPersona)
                    Else
                        Set oPSP = Nothing
                        Exit Sub
                    End If
                    Set frm = Nothing
                Loop
                If bPSPSinDatos Then 'Recalcula el monto de Retención
                    Call cmdRetSistPensActualizar_Click
                    Set oPSP = Nothing
                    Exit Sub
                End If
                'Verifica los montos calculados de retención
                oPSP.SetDatosRetencionSistPens txtBuscarProv.psCodigoPersona, CDate(txtFacFecha.Text), lnRetProvSPMontoBase, Mid(gsOpeCod, 3, 1), lnRetProvSPAporte, lnRetProvSPSeguroAFP, lnRetProvSPComisionAFP, lnTpoSistPensProv
                Set oPSP = Nothing
                If lnRetProvSPAporte <> fnRetProvSPAporte Or lnRetProvSPSeguroAFP <> fnRetProvSPSeguroAFP Or lnRetProvSPComisionAFP <> fnRetProvSPComisionAFP Then
                    MsgBox "El monto de retención de SNP/ONP debe ser recalculado, de click en actualizar SNP/ONP", vbInformation, "Aviso"
                    Exit Sub
                End If
                If fnRetProvSPAporte < 0# Or fnRetProvSPSeguroAFP < 0# Or fnRetProvSPComisionAFP < 0# Then
                    MsgBox "Uno de los conceptos de la Retención es negativo, verifique." & Chr(10) & "Si el problema persiste comuniquese con el Dpto. de TI", vbExclamation, "Aviso"
                    Exit Sub
                End If
            End If
            Set oPSP = Nothing
        End If
    End If
    'END EJVG *******
    
    If MsgBox(" ¿ Seguro de grabar Operación ? ", vbOKCancel, "Aviso de Confirmación") = vbCancel Then
        Exit Sub
    End If
    lnTotalReten = 0
    Set oMov = New DMov
    lTransActiva = True
    
    oMov.BeginTrans
    
    gsMovNro = oMov.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
    oMov.InsertaMov gsMovNro, gsOpeCod, txtMovDesc, gMovEstContabMovContable, gMovFlagVigente
    gnMovNro = oMov.GetnMovNro(gsMovNro)
    oMov.InsertaMovCont gnMovNro, nVal(txtTotal), 0, ""
    'Insertamos el Proveedor
    oMov.InsertaMovGasto gnMovNro, txtBuscarProv.psCodigoPersona, "", Me.cboDocDestino.ListIndex + 1
    
    'JEOM 21-02-2007
    If txtBuscarProvRef.Text <> "" Then
       oMov.InsertaMovGastoRef gnMovNro, txtBuscarProvRef.psCodigoPersona
    End If
    
    If Val(Left(cboDoc, 3)) = 14 Then 'Recibo de Servicios Publicos
        If txtFechaRef <> "" Then
           oMov.InsertaMovDocRef gnMovNro, IIf(txtFacSerie <> "", txtFacSerie & "-", "") & txtFacNro, Format(CDate(txtFechaRef), gsFormatoFecha), 1
        End If
    End If
    'Fin JEOM
    

    'Ahora la Guía de Remisión
    If txtGRNro <> "" Then
       If Not ValidaFechaContab(txtGRFecha, gdFecSis) Then
           oMov.InsertaMovDoc gnMovNro, TpoDocGuiaRemision, IIf(txtGRSerie <> "", txtGRSerie & "-", "") & txtGRNro, Format(txtGRFecha, gsFormatoFecha)
       End If
    End If

    'Por ultimo, el Comprobante del Proveedor
    If txtFacNro <> "" Then
       oMov.InsertaMovDoc gnMovNro, Val(Left(cboDoc, 3)), IIf(txtFacSerie <> "", txtFacSerie & "-", "") & txtFacNro, Format(CDate(txtFacFecha), gsFormatoFecha)
       lnDocTpo = Val(Left(cboDoc, 3))
    Else
       lnDocTpo = 0
    End If

    ' Grabamos en MovObj y MovCant
    nItem = 0
    gnImporte = 0
    For n = 1 To fgDetalle.Rows - 1
      If Len(fgDetalle.TextMatrix(n, 1)) > 0 Then
         nItem = nItem + 1
         If nVal(fgDetalle.TextMatrix(n, 7)) <> 0 Then
           sTexto = ""
           'Recorrido para formar cuenta
           For m = 1 To fgObj.Rows - 1
              If fgObj.TextMatrix(m, 0) = fgDetalle.TextMatrix(n, 0) Then
                  sTexto = sTexto & fgObj.TextMatrix(m, 5)
              End If
           Next m
           
           If Me.cboDocDestino.ListIndex = 3 Or Me.cboDocDestino.ListIndex = 2 Or (Me.cboDocDestino.ListIndex = 1 And fgDetalle.TextMatrix(n, 9) = "") Then
                lnImporte = nVal(fgDetalle.TextMatrix(n, 7))
                nOtrosItem = nItem '*** PEAC 20130212
                For m = 13 To fgDetalle.Cols - 1
                    If fgDetalle.TextMatrix(0, m) <> "AfectoInafectoAIGV" Then
                        If (nVal(fgDetalle.TextMatrix(n, m)) <> 0 And fgImp.TextMatrix(IIf((m - 12) = 1, m - 12, m - 13), 11) = "0") Or (fgImp.TextMatrix(IIf((m - 13) >= 0, m - 13, m - 12), 3) = "AJUSTE") Then
                            '*** PEAC 20100517 - estos cambios solo deben afectar al proceso de provision directa en MN y ME
                            'lnImporte = lnImporte + nVal(fgDetalle.TextMatrix(N, m))
                            'ALPA 20130606
                            'If fgDetalle.TextMatrix(0, m) = "DETRACC." And (gsOpeCod = "701403" Or gsOpeCod = "702403" Or gsOpeCod = "701401" Or gsOpeCod = "702401" Or gsOpeCod = "701402" Or gsOpeCod = "702402") Then
                            If fgDetalle.TextMatrix(0, m) = "DETRACC." And (gsOpeCod = "701403" Or gsOpeCod = "702403" Or gsOpeCod = "701401" Or gsOpeCod = "702401" Or gsOpeCod = "701402" Or gsOpeCod = "702402" Or gsOpeCod = gContProvLogComprobanteMN Or gsOpeCod = gContProvLogComprobanteME) Then  'EJVG20131210
                                lnImporte = lnImporte
                            'ALPA 20130525******************
                            'ElseIf fgDetalle.TextMatrix(0, m) = "RETE." And (gsOpeCod = "701403" Or gsOpeCod = "702403") Then
                            'ElseIf fgDetalle.TextMatrix(0, m) = "RETE." And (gsOpeCod = "701403" Or gsOpeCod = "702403" Or gsOpeCod = "701401" Or gsOpeCod = "702401" Or gsOpeCod = "701402" Or gsOpeCod = "702402") Then
                            ElseIf fgDetalle.TextMatrix(0, m) = "RETE." And (gsOpeCod = "701403" Or gsOpeCod = "702403" Or gsOpeCod = "701401" Or gsOpeCod = "702401" Or gsOpeCod = "701402" Or gsOpeCod = "702402" Or gsOpeCod = gContProvLogComprobanteMN Or gsOpeCod = gContProvLogComprobanteME) Then
                            '*******************************
                                lnImporte = lnImporte
                            Else
                                lnImporte = lnImporte + nVal(fgDetalle.TextMatrix(n, m))
                            End If
                            '*** FIN PEAC
                        End If
                        If nVal(fgDetalle.TextMatrix(n, m)) <> 0 And fgImp.TextMatrix(IIf((m - 12) = 1, m - 12, m - 13), 11) = "1" Then
                           lnTotalReten = lnTotalReten + nVal(fgDetalle.TextMatrix(n, m))
                        End If
                        If nVal(fgDetalle.TextMatrix(n, m)) <> 0 Then
                            '*** PEAC 20130212
                            'oMov.InsertaMovOtrosItem gnMovNro, nItem, fgImp.TextMatrix(IIf((m - 12) = 1, m - 12, m - 13), 6), nVal(fgDetalle.TextMatrix(N, m)), Me.cboDocDestino.ListIndex + 1
                            '***Modificado por ELRO el 20130607, según TI-ERS0064-2013****
                            'oMov.InsertaMovOtrosItem gnMovNro, nOtrosItem, fgImp.TextMatrix(IIf((m - 12) = 1, m - 12, m - 13), 6), nVal(fgDetalle.TextMatrix(N, m)), Me.cboDocDestino.ListIndex + 1
                            oMov.InsertaMovOtrosItem gnMovNro, nOtrosItem, fgImp.TextMatrix(m - 12, 6), nVal(fgDetalle.TextMatrix(n, m)), Me.cboDocDestino.ListIndex + 1
                            '***Fin Modificado por ELRO el 20130607, según TI-ERS0064-2013
                            nOtrosItem = nOtrosItem + 1
                            '*** FIN PEAC
                        End If
                    End If
                Next
           Else
                lnImporte = nVal(fgDetalle.TextMatrix(n, 7))
                nOtrosItem = nItem '*** PEAC 20130212
                For m = 13 To fgDetalle.Cols - 1
                    If nVal(fgDetalle.TextMatrix(n, m)) <> 0 Then
                        If fgDetalle.TextMatrix(0, m) = "AJUSTE" Then
                            '*** PEAC 20130212
                            'oMov.InsertaMovOtrosItem gnMovNro, nItem, fgImp.TextMatrix(m - 12, 6), nVal(fgDetalle.TextMatrix(N, m)), Me.cboDocDestino.ListIndex + 1
                            oMov.InsertaMovOtrosItem gnMovNro, nOtrosItem, fgImp.TextMatrix(m - 12, 6), nVal(fgDetalle.TextMatrix(n, m)), Me.cboDocDestino.ListIndex + 1
                            nOtrosItem = nOtrosItem + 1
                            '*** FIN PEAC
                            lnImporte = lnImporte + nVal(fgDetalle.TextMatrix(n, m))
                        Else
                            '*** PEAC 20130212
                            'oMov.InsertaMovOtrosItem gnMovNro, nItem, fgImp.TextMatrix(m - 12, 6), nVal(fgDetalle.TextMatrix(N, m)), Me.cboDocDestino.ListIndex + 1
                            oMov.InsertaMovOtrosItem gnMovNro, nOtrosItem, fgImp.TextMatrix(m - 12, 6), nVal(fgDetalle.TextMatrix(n, m)), Me.cboDocDestino.ListIndex + 1
                            nOtrosItem = nOtrosItem + 1
                            '*** FIN PEAC
                        End If
                        If nVal(fgDetalle.TextMatrix(n, m)) > 0 And fgImp.TextMatrix(m - 12, 11) = "1" Then
                           lnTotalReten = lnTotalReten + nVal(fgDetalle.TextMatrix(n, m))
                        End If
                    End If
                Next
           End If
           
           oMov.InsertaMovCta gnMovNro, nItem, fgDetalle.TextMatrix(n, 1) & sTexto, lnImporte * IIf(lnDocTpo = 7, -1, 1)
           
           lsTpoCta = fgDetalle.TextMatrix(n, 1)
           
           If oMov.CuentaEsPendiente(fgDetalle.TextMatrix(n, 1) & sTexto, , "D") Then
               oMov.InsertaMovPendientesRend gnMovNro, fgDetalle.TextMatrix(n, 1) & sTexto, lnImporte
           End If
                   
           'Graba Objetos
           lnOrdenObj = 0
           For m = 1 To fgObj.Rows - 1
              If fgObj.TextMatrix(m, 0) = fgDetalle.TextMatrix(n, 0) Then
                  Select Case fgObj.TextMatrix(m, 6)
                      Case ObjCMACAgencias
                          lnOrdenObj = lnOrdenObj + 1
                          lsAgeCod = fgObj.TextMatrix(m, 2)
                          oMov.InsertaMovObj gnMovNro, nItem, lnOrdenObj, ObjCMACAgencias
                          oMov.InsertaMovObjAgenciaArea gnMovNro, nItem, lnOrdenObj, lsAgeCod, ""
                      Case ObjCMACArea
                          lnOrdenObj = lnOrdenObj + 1
                          lsAreaCod = Mid(fgObj.TextMatrix(m, 2), 1, 3)
                          oMov.InsertaMovObj gnMovNro, nItem, lnOrdenObj, ObjCMACArea
                          oMov.InsertaMovObjAgenciaArea gnMovNro, nItem, lnOrdenObj, "", lsAreaCod
                      Case ObjCMACAgenciaArea
                          lsAreaCod = Mid(fgObj.TextMatrix(m, 2), 1, 3)
                          lsAgeCod = Mid(fgObj.TextMatrix(m, 2), 4, 2)
                          lnOrdenObj = lnOrdenObj + 1
                          oMov.InsertaMovObj gnMovNro, nItem, lnOrdenObj, ObjCMACAgenciaArea
                          oMov.InsertaMovObjAgenciaArea gnMovNro, nItem, lnOrdenObj, lsAgeCod, lsAreaCod
                      Case ObjPersona
                          lnOrdenObj = lnOrdenObj + 1
                          oMov.InsertaMovObj gnMovNro, nItem, lnOrdenObj, ObjPersona
                          oMov.InsertaMovGasto gnMovNro, lsCodPers, ""
                      Case ObjBienesServicios
                          lnOrdenObj = lnOrdenObj + 1
                          oMov.InsertaMovObj gnMovNro, nItem, lnOrdenObj, ObjBienesServicios
                          oMov.InsertaMovBS gnMovNro, nItem, lnOrdenObj, fgObj.TextMatrix(m, 2)
                      Case Else
                          lnOrdenObj = lnOrdenObj + 1
                          oMov.InsertaMovObj gnMovNro, nItem, lnOrdenObj, fgObj.TextMatrix(m, 2)
                  End Select
              End If
           Next
           gnImporte = gnImporte + Val(Format(fgDetalle.TextMatrix(n, 7), gsFormatoNumeroDato))
        End If
      End If
    Next
    
    For lnI = 1 To Me.fgImp.Rows - 1
    'ALPA 20090303*****************************
        'If fgImp.TextMatrix(lnI, 2) = "." And (cboDocDestino.ListIndex = -1 Or cboDocDestino.ListIndex = 0 Or fgImp.TextMatrix(lnI, 6) = lsCtaDetraccion Or fgImp.TextMatrix(lnI, 6) = ctaFielCump) Then
        If fgImp.TextMatrix(lnI, 2) = "." And (cboDocDestino.ListIndex = -1 Or cboDocDestino.ListIndex = 0 Or fgImp.TextMatrix(lnI, 6) = lsCtaDetraccion Or fgImp.TextMatrix(lnI, 6) = ctaFielCump Or fgImp.TextMatrix(lnI, 6) = ctaRetenciones) Then
    '*****************************************
            lnImporte = nVal(fgImp.TextMatrix(lnI, 5)) * IIf(fgImp.TextMatrix(lnI, 8) = "D", 1, -1)
            
            If lnImporte <> 0 And fgImp.TextMatrix(lnI, 3) <> "AJUSTE" Then
                nItem = nItem + 1
                If fgImp.TextMatrix(lnI, 3) = "RETE." Then
                    lnPenalidad = 1
                Else
                    lnPenalidad = 0
                End If
                'ALPA 20090317*************************************************************
                'oMov.InsertaMovCta gnMovNro, nItem, fgImp.TextMatrix(lnI, 6), lnImporte * IIf(lnDocTpo = 7, -1, 1)
                oMov.InsertaMovCta gnMovNro, nItem, fgImp.TextMatrix(lnI, 6), lnImporte * IIf(lnDocTpo = 7, -1, 1), lnPenalidad
                '**************************************************************************
            End If
        ElseIf fgImp.TextMatrix(lnI, 2) = "." And (cboDocDestino.ListIndex = 1) Then
            lnImporte = nVal(fgImp.TextMatrix(lnI, 5)) * IIf(fgImp.TextMatrix(lnI, 8) = "D", 1, -1)
            
            lnMontoNoGrabado = 0
            
            For lnA = 1 To Me.fgDetalle.Rows - 1
                If fgDetalle.TextMatrix(lnA, 9) = "" Then
                    lnMontoNoGrabado = lnMontoNoGrabado + fgDetalle.TextMatrix(lnA, 12 + lnI)
                End If
            Next lnA
            
            If fgImp.TextMatrix(lnI, 11) = "0" Then
                lnImporte = lnImporte - lnMontoNoGrabado
            End If
            
            If lnImporte <> 0 And fgImp.TextMatrix(lnI, 3) <> "AJUSTE" Then
                nItem = nItem + 1
                oMov.InsertaMovCta gnMovNro, nItem, fgImp.TextMatrix(lnI, 6), lnImporte * IIf(lnDocTpo = 7, -1, 1)
            End If
            
        'Para Retenciones
        ElseIf fgImp.TextMatrix(lnI, 2) = "." And fgImp.TextMatrix(lnI, 11) = "1" Then
            lnImporte = nVal(fgImp.TextMatrix(lnI, 5)) * IIf(fgImp.TextMatrix(lnI, 8) = "D", 1, -1)
            
            If lnImporte <> 0 And fgImp.TextMatrix(lnI, 3) <> "AJUSTE" Then
                nItem = nItem + 1
                oMov.InsertaMovCta gnMovNro, nItem, fgImp.TextMatrix(lnI, 6), lnImporte * IIf(lnDocTpo = 7, -1, 1)
            End If
            
        End If
    Next lnI
    
    'Grabamos la Provisión
    If gnImporte = 0 Then
       Err.Raise "50001", "Grabando", "No se puede grabar Documento sin Importe"
    End If
    
'EJVG20140727 ***
If gbBitRetencSistPensProv Then
    If lnRetProvSPMontoBase > 0 Then 'Entra al calculo acumulativo
        oMov.InsertaMovProvRetencSistPens gnMovNro, lnTpoSistPensProv, lnRetProvSPMontoBase, fnRetProvSPAporte + fnRetProvSPSeguroAFP + fnRetProvSPComisionAFP
    End If
    If fnRetProvSPAporte > 0 Then
        nItem = nItem + 1
        oMov.InsertaMovCta gnMovNro, nItem, ReemplazaCtaCont(IIf(lnTpoSistPensProv = AFP, fsCtaContAporteAFP, fsCtaContAporteONP), Mid(gsOpeCod, 3, 1)), fnRetProvSPAporte * -1
        oMov.InsertaMovProvRetencSistPensDet gnMovNro, nItem, ConceptoRetencionSistemaPensionProveedor.Aporte, fnRetProvSPAporte, fnRetProvSPAporte
    End If
    If fnRetProvSPSeguroAFP > 0 Then
        nItem = nItem + 1
        oMov.InsertaMovCta gnMovNro, nItem, ReemplazaCtaCont(fsCtaContSeguroAFP, Mid(gsOpeCod, 3, 1)), fnRetProvSPSeguroAFP * -1
        oMov.InsertaMovProvRetencSistPensDet gnMovNro, nItem, ConceptoRetencionSistemaPensionProveedor.SeguroAFP, fnRetProvSPSeguroAFP, fnRetProvSPSeguroAFP
    End If
    If fnRetProvSPComisionAFP > 0 Then
        nItem = nItem + 1
        oMov.InsertaMovCta gnMovNro, nItem, ReemplazaCtaCont(fsCtaContComisionAFP, Mid(gsOpeCod, 3, 1)), fnRetProvSPComisionAFP * -1
        oMov.InsertaMovProvRetencSistPensDet gnMovNro, nItem, ConceptoRetencionSistemaPensionProveedor.ComsionAFP, fnRetProvSPComisionAFP, fnRetProvSPComisionAFP
    End If
End If
'END EJVG *******
'Grabamos la Provisión
If nVal(txtTotal) <> 0 Then
   nItem = nItem + 1
   gnImporte = nVal(txtTotal)
   oMov.InsertaMovCta gnMovNro, nItem, Trim(Right(cboProvis, 22)), gnImporte * -1 * IIf(lnDocTpo = 7, -1, 1)
End If
If oMov.CuentaEsPendiente(Trim(Right(cboProvis.Text, 22)), , IIf(lPendPasivo, "A", "D")) Then
   oMov.InsertaMovPendientesRend gnMovNro, Trim(Right(cboProvis.Text, 22)), nVal(txtTotal)
End If
If lbRegulaPend Then
    Set rsPend = frmAnalisisRegulaPend.lvPend.GetRsNew()
    Do While Not rsPend.EOF
       If rsPend!OK = 1 Then
            oMov.InsertaMovRef gnMovNro, rsPend!nMovNro, frmAnalisisRegulaPend.txtAgecod
            oMov.ActualizaMovPendientesRend rsPend!nMovNro, frmAnalisisRegulaPend.txtCtaPend, rsPend!Rendicion
        End If
        rsPend.MoveNext
    Loop
End If
RSClose rs
If Not rsDetracc Is Nothing Then
   If rsDetracc.State = adStateOpen Then
      Do While Not rsDetracc.EOF
         If rsDetracc!OK = 1 Then
            oMov.InsertaMovRef gnMovNro, rsDetracc![Nro.Doc.], ""
            oMov.ActualizaMovPendientesRend rsDetracc![Nro.Doc.], rsDetracc!cCtaPendiente, rsDetracc![Regulariza]
         End If
         rsDetracc.MoveNext
      Loop
   End If
End If
'Grabamos MovRef Referencias a la Orden de Compra ...si existe
If lnMovNro > 0 Then
    oMov.InsertaMovRef gnMovNro, lnMovNro
End If

'Grabamos el Tipo de Cambio de Compra
'If RTrim(LTrim(Str(Mid(lsTpoCta, 1, 1)))) <> "1" And gsOpeCod = "702401" Then
'    If RTrim(LTrim(Str(Mid(lsTpoCta, 1, 1)))) <> "2" And gsOpeCod = "702401" Then
        If gsSimbolo = gcME Then
            'Mdificado PASI20150417 x INC1504160009
            'oMov.GeneraMovME gnMovNro, gsMovNro, Format(Me.txtTipVariable, gsFormatoNumeroDato), Format(gnTipCambioPonderado, gsFormatoNumeroView3Dec), gsOpeCod
            'TORE - ADD - RFC1712260009
            oMov.GeneraMovMExProvision gnMovNro, gsMovNro, Format(Me.txtTipVariable, gsFormatoNumeroDato), Format(gnTipCambioPonderado, gsFormatoNumeroView3Dec), gsOpeCod, , , , , , , , , , , , , , , txtFacFecha.Text
            'END PASI
        End If
'    End If
'End If


'*** PEAC 20110314 - Grabamos el asiento de cuentas de orden
'*** de los Bienes No Depreciables

lnMontoCtaOrd = 0
If gsOpeCod = "701401" Or gsOpeCod = "702401" Then 'provisión de una orden de compra
    For n = 1 To Me.fgDetalle.Rows - 1
        If Left(fgDetalle.TextMatrix(n, 1), 10) = "45" & Mid(gsOpeCod, 3, 1) & "3011101" Then
            lnMontoCtaOrd = lnMontoCtaOrd + fgDetalle.TextMatrix(n, 7)
        End If
    Next n
    
    If lnMontoCtaOrd > 0 Then
        Set rsCtasOrden = oMov.GetCuentaOrdenBienNoDepre(Mid(gsOpeCod, 3, 1), Right(fgDetalle.TextMatrix(1, 1), 2))
        If Not (rsCtasOrden.EOF And rsCtasOrden.BOF) Then
            nItem = nItem + 1
            oMov.InsertaMovCta gnMovNro, nItem, rsCtasOrden!cCtaDebe, lnMontoCtaOrd
            nItem = nItem + 1
            oMov.InsertaMovCta gnMovNro, nItem, rsCtasOrden!cCtaHaber, lnMontoCtaOrd * -1
        End If
    End If
ElseIf gsOpeCod Like "70140[36]" Or gsOpeCod Like "70240[36]" Then 'provisión de una orden de compra 'NAGL 20190327 Agregó LIKE [36]
    For n = 1 To Me.fgDetalle.Rows - 1
        sTexto = ""
        For m = 1 To fgObj.Rows - 1
           If fgObj.TextMatrix(m, 0) = fgDetalle.TextMatrix(n, 0) Then
               sTexto = sTexto & fgObj.TextMatrix(m, 5)
           End If
        Next m
            
        If Len(sTexto) = 2 Then
            If fgDetalle.TextMatrix(n, 1) = "45" & Mid(gsOpeCod, 3, 1) & "3011101" Then
                lnMontoCtaOrd = lnMontoCtaOrd + fgDetalle.TextMatrix(n, 7)
            End If
        Else
            If fgDetalle.TextMatrix(n, 1) = "45" & Mid(gsOpeCod, 3, 1) & "3011101" + Right(Trim(fgDetalle.TextMatrix(n, 1)), 2) Then  'NAGL 20190327 Agregó Right(Trim(fgDetalle.TextMatrix(n, 1)), 2) 'NAGL quito " + Left(sTexto, 2)"
                lnMontoCtaOrd = lnMontoCtaOrd + fgDetalle.TextMatrix(n, 7)
                If n < Me.fgDetalle.Rows - 1 Then
                    If (fgDetalle.TextMatrix(n, 1) <> fgDetalle.TextMatrix(n + 1, 1)) Then
                            If lnMontoCtaOrd > 0 Then
                                Set rsCtasOrden = oMov.GetCuentaOrdenBienNoDepre(Mid(gsOpeCod, 3, 1), Right(Trim(fgDetalle.TextMatrix(n, 1)), 2))
                                If Not (rsCtasOrden.EOF And rsCtasOrden.BOF) Then
                                    nItem = nItem + 1
                                    oMov.InsertaMovCta gnMovNro, nItem, rsCtasOrden!cCtaDebe, lnMontoCtaOrd
                                    nItem = nItem + 1
                                    oMov.InsertaMovCta gnMovNro, nItem, rsCtasOrden!cCtaHaber, lnMontoCtaOrd * -1
                                End If
                            End If
                            lnMontoCtaOrd = 0
                    End If
                Else
                        If lnMontoCtaOrd > 0 Then
                            Set rsCtasOrden = oMov.GetCuentaOrdenBienNoDepre(Mid(gsOpeCod, 3, 1), Right(Trim(fgDetalle.TextMatrix(n, 1)), 2))
                            If Not (rsCtasOrden.EOF And rsCtasOrden.BOF) Then
                                nItem = nItem + 1
                                oMov.InsertaMovCta gnMovNro, nItem, rsCtasOrden!cCtaDebe, lnMontoCtaOrd
                                nItem = nItem + 1
                                oMov.InsertaMovCta gnMovNro, nItem, rsCtasOrden!cCtaHaber, lnMontoCtaOrd * -1
                            End If
                        End If
                        lnMontoCtaOrd = 0
                End If
            End If 'Sección Agregado by NAGL 20190327 Según RFC1906060005
        End If
    Next n
    
'    If lnMontoCtaOrd > 0 Then
'        Set rsCtasOrden = oMov.GetCuentaOrdenBienNoDepre(Mid(gsOpeCod, 3, 1), Right(fgObj.TextMatrix(1, 2), 2))
'        If Not (rsCtasOrden.EOF And rsCtasOrden.BOF) Then
'            nItem = nItem + 1
'            oMov.InsertaMovCta gnMovNro, nItem, rsCtasOrden!cCtaDebe, lnMontoCtaOrd
'            nItem = nItem + 1
'            oMov.InsertaMovCta gnMovNro, nItem, rsCtasOrden!cCtaHaber, lnMontoCtaOrd * -1
'        End If
'    End If 'Comentado by NAGL 20190327
End If

'*** FIN PEAC
'EJVG20131116 ***
If gsOpeCod = gContProvLogComprobanteMN Or gsOpeCod = gContProvLogComprobanteME Then
    oMov.ActualizaComprobanteLog fnComprobanteMovNro, Val(Left(cboDoc, 3)), IIf(txtFacSerie.Text <> "", txtFacSerie.Text & "-", "") & txtFacNro.Text, CDate(txtFacFecha.Text)
    oMov.ActualizaComprobanteLogNew fnComprobanteMovNro, Val(Left(cboDoc, 3)), IIf(txtFacSerie.Text <> "", txtFacSerie.Text & "-", "") & txtFacNro.Text, CDate(txtFacFecha.Text) 'PASIERS0772014
    'oMov.RevierteAsientoActaConformidad gnMovNro, fnComprobanteMovNro
    oMov.ActualizaEstadoCuotasCronogramaxProvision fnComprobanteMovNro
    oMov.ActualizaEstadoCuotasCronogramaxProvisionNew fnComprobanteMovNro 'PASIERS0772014
    oMov.ActualizaEstadoBienesxProvision fnComprobanteMovNro 'PASIERS0772014
End If
'END EJVG *******
oMov.CommitTrans

lTransActiva = False
gnSaldo = gnSaldo - nVal(txtTotal)

   OK = True
   Dim lsImpre As String
   Dim oImpre As New NContImprimir
   lsImpre = ""
   oImpre.Inicio gsNomCmac, gsNomAge, Format(gdFecSis, gsFormatoFechaView)
   If Val(Left(cboDoc, 2)) = TpoDocRecEgreso Then
      lsImpre = oImpre.ImprimeRecibo(gsMovNro)
   End If
   
   'JEOM
   If gsOpeCod = 701401 Or gsOpeCod = 702401 Or gsOpeCod = 701402 Or gsOpeCod = 702402 Then
      lnMovNro = 0
   End If
   'FIN
   
   'If lnMovNro = 0 Then
   If lnMovNro = 0 Or (lnMovNro > 0 And (gsOpeCod = gContProvLogComprobanteMN Or gsOpeCod = gContProvLogComprobanteME)) Then 'EJVG20131116
      If Not lsImpre = "" Then
      lsImpre = lsImpre & oImpresora.gPrnSaltoPagina
      End If
      'lsImpre = lsImpre & oImpre.ImprimeAsientoContable(gsMovNro, gnLinPage, gnColPage, gsOpeDesc)
      lsImpre = lsImpre & oImpre.ImprimeAsientoContableProv(gsMovNro, gnLinPage, gnColPage, gsOpeDesc)
   End If
   Set oImpre = Nothing
   If Not lsImpre = "" Then
        EnviaPrevio lsImpre, gsOpeDesc, gnLinPage, False
   End If
   
    'EJVG20131115 ***
    If lbVistoOK Then
        ofrmVistoElect.RegistraVistoElectronico gnMovNro
        Set ofrmVistoElect = Nothing
    End If
    'END EJVG *******
    
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaMantDocumento
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se ha Registrado Garantia por Alquiler del Proveedor :" & lblProvNombre
            Set objPista = Nothing
            '*******
        
 'JEOM
  If gsOpeCod = 701401 Or gsOpeCod = 702401 Or gsOpeCod = 701402 Or gsOpeCod = 702402 Then
     Unload Me
  Else
        'If lnMovNro <= 0 Then
        If lnMovNro <= 0 Or (lnMovNro > 0 And (gsOpeCod = gContProvLogComprobanteMN Or gsOpeCod = gContProvLogComprobanteME)) Then 'EJVG20131116
             lnMovNro = 0
             If MsgBox("¿ Desea Registrar otro Documento ?", vbQuestion + vbYesNo, "¡Aviso!") = vbYes Then
                'ALPA 20100309*************
                 nCantAgeSel = 0
                 nMontogasto = 0
                 '*************************
                 txtProvCod = ""
                 txtBuscarProv = ""
                 lblProvNombre = ""
                 cboDoc.ListIndex = -1
                 txtFacSerie = ""
                 txtFacNro = ""
                 txtFacFecha = gdFecSis
                 fgDetalle.Rows = 2
                 EliminaRow fgDetalle, 1
                 fgObj.Rows = 2
                 fgObj.EliminaFila 1
                 Me.txtSTotal = "0.00"
                 Me.txtTotal = "0.00"
                 If gsOpeCod = gContProvLogComprobanteMN Or gsOpeCod = gContProvLogComprobanteME Then 'EJVG20131116
                    cmdComprobanteLimpiar_click
                    ActualizaComprobantes
                 End If
             Else
                Unload Me
             End If
        Else
            Unload Me
        End If
  End If
  'FIN
  
Exit Sub
ErrAceptar:
    lsMsgError = TextErr(Err.Description)
    If lTransActiva Then
        oMov.RollbackTrans
    End If
    lTransActiva = False
    MsgBox lsMsgError, vbInformation, "¡Aviso!"
End Sub

Public Property Get lOk() As Boolean
    lOk = OK
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
    OK = vNewValue
End Property

Private Sub txtObj_GotFocus()
    txtObj.Width = txtObj.Width - cmdExaminar.Width + 10
    cmdExaminar.Top = txtObj.Top + 15
    cmdExaminar.Left = txtObj.Left + txtObj.Width - 10
    cmdExaminar.Visible = True
    cmdExaminar.TabIndex = txtObj.TabIndex + 1
End Sub

Private Sub txtObj_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Or KeyCode = 38 Then
       txtObj_KeyPress 13
       MoveFlex fgDetalle, KeyCode
    End If
End Sub

Private Sub txtObj_KeyPress(KeyAscii As Integer)
    Dim nItem As Integer
    If KeyAscii = 13 Then
       If ValidaCta Then
          fgDetalle.Enabled = True
          txtObj.Visible = False
          cmdExaminar.Visible = False
          fgDetalle.Col = 7
          fgDetalle.SetFocus
       End If
    End If
End Sub


Private Function ValidaCta() As Boolean
    Dim oOpe As New DOperacion
    ValidaCta = False
    If Len(txtObj) = 0 Then
       txtObj.Visible = False
       cmdExaminar.Visible = False
       cmdEliminar_Click
       Exit Function
    End If
    Set rs = oOpe.CargaOpeCta(Left(gsOpeCod, 5) & "%", "D")
    If Not rs.EOF Then
       rs.Find "cCtaContCod = '" & txtObj & "'"
       If Not rs.EOF Then
          glAceptar = True
          If gsOpeCod = "701401" Or gsOpeCod = "701402" Or gsOpeCod = "702401" Or gsOpeCod = "702402" Then
             AsignaObjetosSer rs!cCtaContCod, fgDetalle.TextMatrix(fgDetalle.row, 2)
          Else
              AsignaObjetosSer rs!cCtaContCod, rs!cCtaContDesc
          End If
          If glAceptar Then
             ValidaCta = True
          End If
       End If
    End If
    RSClose rs
    Call ObtenerValoresPorCtas
    If Not ValidaCta Then
       MsgBox "Concepto no Asignado a Operación...!", vbInformation, "¡Aviso!"
    End If
End Function

Private Sub txtObj_LostFocus()
    txtObj.Width = txtObj.Width + cmdExaminar.Width - 10
End Sub

Private Sub AsignaObjetosSer(sCtaCod As String, sCtaDes As String)
    Dim rsObj As ADODB.Recordset
    Set rsObj = New ADODB.Recordset
    Dim nNiv  As Integer
    Dim nObj  As Integer
    Dim nObjs As Integer
    Dim n     As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oCtaCont As DCtaCont
    Set oCtaCont = New DCtaCont
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    Dim oRHAreas As DActualizaDatosArea
    Dim oCtaIf As NCajaCtaIF
    Dim oEfect As Defectivo
    Set oRHAreas = New DActualizaDatosArea
    Set oCtaIf = New NCajaCtaIF
    Set oEfect = New Defectivo
    Dim oDescObj As ClassDescObjeto
    Set oDescObj = New ClassDescObjeto
    Dim oContFunct As NContFunciones
    Set oContFunct = New NContFunciones
    Dim lsRaiz  As String
    Dim lsFiltro  As String
    
    
    
    oCon.AbreConexion
     If gsOpeCod = "701401" Or gsOpeCod = "701402" Or gsOpeCod = "702401" Or gsOpeCod = "702402" Then
     Else
      EliminaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
     End If
      fgDetalle = sCtaCod
      fgDetalle.TextMatrix(fgDetalle.row, 2) = sCtaDes
      fgDetalle.TextMatrix(fgDetalle.row, 8) = sCtaCod
      
      sSql = "SELECT MAX(nCtaObjOrden) as nNiveles FROM CtaObj WHERE cCtaContCod = '" & sCtaCod & "' and cObjetoCod <> '00' "
      Set rs = oCon.CargaRecordSet(sSql)
      nObjs = IIf(IsNull(rs!nNiveles), 0, rs!nNiveles)
     cSubCta = ""
    Set rs1 = oCtaCont.CargaCtaObj(sCtaCod, , True)
    If Not rs1.EOF And Not rs1.BOF Then
        Do While Not rs1.EOF
            lsRaiz = ""
            lsFiltro = ""
            Set rs = New ADODB.Recordset
            Select Case Val(rs1!cObjetoCod)
                Case ObjCMACAgencias
                    Set rs = oRHAreas.GetAgencias()
                Case ObjCMACAgenciaArea
                    lsRaiz = "Unidades Organizacionales"
                    Set rs = oRHAreas.GetAgenciasAreas(rs1!cCtaObjFiltro)
                Case ObjCMACArea
                    Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
                Case ObjEntidadesFinancieras
                    lsRaiz = "Cuentas de Entidades Financieras"
                    Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, sCtaCod)
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
                               glAceptar = True
                             
                                lsFiltro = oContFunct.GetFiltroObjetos(Val(rs1!cObjetoCod), sCtaCod, oDescObj.gsSelecCod, False)
                                If Len(Trim(cSubCta)) = 0 Then
                                cSubCta = oDescObj.gsSelecCod
                                End If
                                AdicionaObj sCtaCod, fgDetalle.TextMatrix(fgDetalle.row, 0), rs1!nCtaObjOrden, oDescObj.gsSelecCod, _
                                            oDescObj.gsSelecDesc, lsFiltro, rs1!cObjetoCod
                            Else
                                EliminaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
                                'fgDetalle.EliminaFila fgDetalle.Row, False
                                Exit Do
                            End If
                        Else
                            AdicionaObj sCtaCod, fgDetalle.TextMatrix(fgDetalle.row, 0), rs1!nCtaObjOrden, rs1!cObjetoCod, _
                                            rs1!cObjetoDesc, lsFiltro, rs1!cObjetoCod
                        End If
                    End If
                End If
            End If
            rs1.MoveNext
        Loop
    End If
    rs1.Close
    Set rs1 = Nothing
    Set oDescObj = Nothing
    
    Set oCtaCont = Nothing
    Set oCtaIf = Nothing
    Set oEfect = Nothing
      
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
    
    For m = 1 To Me.fgObj.Rows - 1
       If Me.fgObj.TextMatrix(m, 0) <> "" Then
         If Me.fgObj.TextMatrix(m, 0) > nItem Then
            Me.fgObj.TextMatrix(m, 0) = Me.fgObj.TextMatrix(m, 0) - 1
         End If
       End If
    Next m
    
    
End Sub

Private Sub AdicionaObj(sCodCta As String, nFila As Integer, pnCtaObjOrden As Integer, psSelecCod As String, psSelecDesc As String, psFiltro As String, psObjetoCod As String)
   Dim nItem As Integer
   fgObj.AdicionaFila
   nItem = fgObj.row
   fgObj.TextMatrix(nItem, 0) = nFila
   fgObj.TextMatrix(nItem, 1) = pnCtaObjOrden
   fgObj.TextMatrix(fgObj.row, 2) = psSelecCod
   fgObj.TextMatrix(fgObj.row, 3) = psSelecDesc
   fgObj.TextMatrix(fgObj.row, 4) = sCodCta
   fgObj.TextMatrix(fgObj.row, 5) = psFiltro
   fgObj.TextMatrix(fgObj.row, 6) = psObjetoCod
   fgObj.TextMatrix(fgObj.row, 7) = nFila
End Sub

Private Sub EliminaObjeto(nItem As Integer)
    EliminaFgObj nItem
    If Len(fgDetalle.TextMatrix(1, 1)) > 0 Then
       RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
    End If
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

Private Sub cboDocDestino_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cboDocDestino_Click
       If fgImp.Enabled Then
          fgImp.SetFocus
       Else
          fgDetalle.SetFocus
       End If
    End If
End Sub

Private Sub cboDocDestino_Click()
    Dim i As Integer
    For i = 1 To fgImp.Rows - 1
         'si el destino del impuesto es puede ser gravado y si es obligatorio
        If fgImp.TextMatrix(i, 9) = "1" And fgImp.TextMatrix(i, 10) = "1" Then
            If cboDocDestino.ListIndex = 3 Then
               fgImp.TextMatrix(i, 2) = ""
            Else
               fgImp.TextMatrix(i, 2) = "1"
            End If
            CalculaTotal
        Else
            'fgImp.TextMatrix(i, 2) = ""
        End If
    Next
    
    If Me.cboDocDestino.ListIndex = 3 Then
        fgImp.lbEditarFlex = False
    Else
        fgImp.lbEditarFlex = True
    End If
    
    If cboDocDestino.ListIndex <> 1 Then
        fgDetalle.ColWidth(9) = 0
        For i = 1 To fgDetalle.Rows - 1
            'limpia las posibles gravaciones manuales realizadas
            If fgDetalle.TextMatrix(i, 9) = "." Then
               fgDetalle.TextMatrix(i, 9) = ""
           End If
        Next
    Else
        fgDetalle.ColWidth(9) = 500
    End If
End Sub
'*******************************************************************************************

'EJVG20131113 ***
Private Sub ActualizaComprobantes()
    Dim oLog As New DLogGeneral
    Set fRsComprobante = New ADODB.Recordset
    Set fRsComprobante = oLog.ListaComprobantesxProvision(Mid(gsOpeCod, 3, 1))
    Set oLog = Nothing
End Sub
Private Function ListaIndiceComprobante(ByVal pnValor As Integer) As Integer
    ListaIndiceComprobante = -1
    Dim i As Integer
    For i = 0 To cboDoc.ListCount - 1
        If CInt(Trim(Left(cboDoc.List(i), 3))) = pnValor Then
            ListaIndiceComprobante = i
            Exit Function
        End If
    Next
End Function
'END EJVG *******
'EJVG20140724 ***
Private Function MontoBaseOperacion() As Currency
    Dim i As Integer
    Dim lnMonto As Currency
    If fgDetalle.TextMatrix(1, 0) <> "" Then
        For i = 1 To fgDetalle.Rows - 1
            lnMonto = lnMonto + CCur(IIf(fgDetalle.TextMatrix(i, 12) = "", fgDetalle.TextMatrix(i, 7), fgDetalle.TextMatrix(i, 12)))
        Next
    End If
    MontoBaseOperacion = lnMonto
End Function
Private Function ReemplazaCtaCont(ByVal psCtaContCod As String, Optional ByVal psMoneda As String = "") As String
    Dim Temp As String
    Temp = psCtaContCod
    If Len(psMoneda) <> 0 Then
        Temp = Replace(Temp, "M", psMoneda)
    End If
    
    ReemplazaCtaCont = Temp
End Function
'END EJVG *******
