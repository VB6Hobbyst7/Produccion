VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOpeDocChica 
   Caption         =   "Caja Chica: Registro de Documentos"
   ClientHeight    =   8940
   ClientLeft      =   660
   ClientTop       =   1065
   ClientWidth     =   10650
   Icon            =   "frmOpeDocChica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   10650
   Begin VB.Frame FRLugarMovilidad 
      Caption         =   "Lugar de Desplazamiento"
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
      Height          =   670
      Left            =   90
      TabIndex        =   71
      Top             =   8205
      Width           =   7320
      Begin VB.TextBox txtLugar 
         Height          =   350
         Left            =   120
         TabIndex        =   72
         Top             =   240
         Width           =   7095
      End
   End
   Begin VB.CheckBox ChkVisitaAge 
      Caption         =   "Viajes a Agencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   67
      Top             =   720
      Visible         =   0   'False
      Width           =   4305
   End
   Begin VB.Frame Frame4 
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
      Height          =   660
      Left            =   6210
      TabIndex        =   17
      Top             =   0
      Width           =   4035
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   2685
         TabIndex        =   1
         Top             =   210
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label txtMovNro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   690
         TabIndex        =   38
         Top             =   780
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Label txtOpeCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1005
         TabIndex        =   37
         Top             =   210
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   2040
         TabIndex        =   19
         Top             =   247
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Operacion :"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.CheckBox chkDocDolares 
      Alignment       =   1  'Right Justify
      Caption         =   "Doc Dolares."
      Height          =   255
      Left            =   9120
      TabIndex        =   62
      Top             =   1230
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Frame fraDoc 
      Caption         =   "Documento "
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
      Height          =   1620
      Left            =   6210
      TabIndex        =   21
      Top             =   1035
      Width           =   4395
      Begin VB.ComboBox cboDocDestino 
         Height          =   315
         ItemData        =   "frmOpeDocChica.frx":030A
         Left            =   840
         List            =   "frmOpeDocChica.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1230
         Width           =   3375
      End
      Begin VB.TextBox txtDocSerie 
         Height          =   315
         Left            =   510
         MaxLength       =   4
         TabIndex        =   5
         Top             =   855
         Width           =   600
      End
      Begin VB.ComboBox cboDoc 
         Height          =   315
         Left            =   510
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   495
         Width           =   3720
      End
      Begin VB.TextBox txtDocNro 
         Height          =   315
         Left            =   1110
         MaxLength       =   11
         TabIndex        =   6
         Top             =   855
         Width           =   1200
      End
      Begin MSMask.MaskEdBox txtDocFecha 
         Height          =   315
         Left            =   3090
         TabIndex        =   7
         Top             =   870
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
         Caption         =   "Destino"
         Height          =   255
         Left            =   135
         TabIndex        =   25
         Top             =   1275
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   540
         Width           =   360
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha"
         Height          =   240
         Left            =   2490
         TabIndex        =   23
         Top             =   915
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro."
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   915
         Width           =   300
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
      Height          =   660
      Left            =   2955
      TabIndex        =   43
      Top             =   8205
      Visible         =   0   'False
      Width           =   4350
      Begin VB.TextBox txtTipPondVenta 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3495
         TabIndex        =   63
         Top             =   210
         Width           =   735
      End
      Begin VB.TextBox txtTipFijo 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   360
         TabIndex        =   45
         Top             =   210
         Width           =   735
      End
      Begin VB.TextBox txtTipVariable 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1725
         TabIndex        =   44
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Venta Pond"
         Height          =   255
         Left            =   2550
         TabIndex        =   64
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label10 
         Caption         =   "Fijo"
         Height          =   255
         Left            =   75
         TabIndex        =   47
         Top             =   255
         Width           =   345
      End
      Begin VB.Label Label9 
         Caption         =   "Banco"
         Height          =   255
         Left            =   1170
         TabIndex        =   46
         Top             =   240
         Width           =   495
      End
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
      Left            =   45
      TabIndex        =   31
      Top             =   8205
      Width           =   2865
      Begin VB.ComboBox cboProvis 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   225
         Width           =   2595
      End
   End
   Begin VB.Frame fraServicio 
      Height          =   5535
      Left            =   60
      TabIndex        =   26
      Top             =   2640
      Width           =   10545
      Begin VB.Frame fraRetencSistPens 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   7680
         TabIndex        =   73
         Top             =   4200
         Visible         =   0   'False
         Width           =   2835
         Begin VB.CommandButton cmdRetSistPensDetalle 
            Appearance      =   0  'Flat
            Height          =   320
            Left            =   990
            Picture         =   "frmOpeDocChica.frx":038F
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Ver retención SNP/ONP"
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
            TabIndex        =   75
            Text            =   "0.00"
            Top             =   150
            Width           =   1185
         End
         Begin VB.CommandButton cmdRetSistPensActualizar 
            Appearance      =   0  'Flat
            Height          =   320
            Left            =   2520
            Picture         =   "frmOpeDocChica.frx":0748
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Actualizar retención SNP/ONP"
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
            TabIndex        =   77
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
      Begin VB.ComboBox cbogastos 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1965
         Width           =   3720
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "A&gregar"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4380
         Width           =   1110
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   330
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4380
         Width           =   1110
      End
      Begin VB.CommandButton cmdValVenta 
         Caption         =   "Ajuste &Valor Venta"
         Height          =   330
         Left            =   2400
         TabIndex        =   40
         ToolTipText     =   "Calcula el Valor Venta de Subtotales"
         Top             =   4380
         Width           =   2220
      End
      Begin VB.CommandButton cmdAjuste 
         Caption         =   "A&juste Manual   >>>"
         Height          =   330
         Left            =   4650
         TabIndex        =   39
         ToolTipText     =   "Adiciona Importe de Ajuste de Documento"
         Top             =   4380
         Width           =   2220
      End
      Begin Sicmact.FlexEdit fgImp 
         Height          =   1800
         Left            =   7140
         TabIndex        =   11
         Top             =   2445
         Width           =   3315
         _extentx        =   5847
         _extenty        =   3175
         cols0           =   12
         highlight       =   2
         allowuserresizing=   3
         encabezadosnombres=   "-#1-Ok-Impuesto-Tasa-Monto-CtaCont-CtaContDesc-cDocImpDH-cImpDestino-cDocImpOpc-nCalculo"
         encabezadosanchos=   "0-0-350-1000-600-1200-0-0-0-0-0-1"
         font            =   "frmOpeDocChica.frx":0B25
         font            =   "frmOpeDocChica.frx":0B51
         font            =   "frmOpeDocChica.frx":0B7D
         font            =   "frmOpeDocChica.frx":0BA9
         font            =   "frmOpeDocChica.frx":0BD5
         fontfixed       =   "frmOpeDocChica.frx":0C01
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-2-X-X-X-X-X-X-X-X-X"
         textstylefixed  =   4
         listacontroles  =   "0-0-4-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-L-C-R-R-C-C-C-C-L-C"
         formatosedit    =   "0-0-0-2-2-2-2-2-2-2-0-0"
         lbeditarflex    =   -1
         lbformatocol    =   -1
         lbbuscaduplicadotext=   -1
         rowheight0      =   300
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
         Left            =   9240
         TabIndex        =   33
         Top             =   4815
         Width           =   1185
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
         Left            =   9240
         TabIndex        =   32
         Top             =   5145
         Width           =   1185
      End
      Begin Sicmact.FlexEdit fgObj 
         Height          =   1800
         Left            =   105
         TabIndex        =   10
         Top             =   2445
         Width           =   5955
         _extentx        =   10504
         _extenty        =   3175
         cols0           =   8
         highlight       =   2
         allowuserresizing=   1
         encabezadosnombres=   "#-Ord-Código-Descripción-CtaCont-SubCta-ObjPadre-ItemCtaCont"
         encabezadosanchos=   "350-400-1200-3000-0-900-0-0"
         font            =   "frmOpeDocChica.frx":0C2F
         font            =   "frmOpeDocChica.frx":0C5B
         font            =   "frmOpeDocChica.frx":0C87
         font            =   "frmOpeDocChica.frx":0CB3
         font            =   "frmOpeDocChica.frx":0CDF
         fontfixed       =   "frmOpeDocChica.frx":0D0B
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
      Begin Sicmact.FlexEdit fgDetalle 
         Height          =   1470
         Left            =   120
         TabIndex        =   9
         Top             =   210
         Width           =   10350
         _extentx        =   17806
         _extenty        =   2593
         cols0           =   7
         highlight       =   2
         allowuserresizing=   1
         encabezadosnombres=   "#-Código-Descripción-Monto-DH-Grav.-ItemCtaCont"
         encabezadosanchos=   "350-1200-3500-1200-0-0-0"
         font            =   "frmOpeDocChica.frx":0D39
         font            =   "frmOpeDocChica.frx":0D65
         font            =   "frmOpeDocChica.frx":0D91
         font            =   "frmOpeDocChica.frx":0DBD
         font            =   "frmOpeDocChica.frx":0DE9
         fontfixed       =   "frmOpeDocChica.frx":0E15
         lbultimainstancia=   -1
         columnasaeditar =   "X-1-X-3-X-5-X"
         textstylefixed  =   3
         listacontroles  =   "0-1-0-0-0-4-0"
         encabezadosalineacion=   "C-L-L-R-C-C-C"
         formatosedit    =   "0-0-0-2-2-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbformatocol    =   -1
         lbpuntero       =   -1
         lbbuscaduplicadotext=   -1
         colwidth0       =   345
         rowheight0      =   300
      End
      Begin VB.Frame fraAjuste 
         Height          =   660
         Left            =   120
         TabIndex        =   58
         Top             =   4260
         Visible         =   0   'False
         Width           =   4035
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
            Height          =   360
            Left            =   1590
            TabIndex        =   61
            Top             =   195
            Width           =   1170
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
            Height          =   360
            Left            =   2745
            TabIndex        =   60
            Top             =   195
            Width           =   1170
         End
         Begin VB.TextBox txtAjuste 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   135
            TabIndex        =   59
            Top             =   210
            Width           =   1410
         End
      End
      Begin Sicmact.ActXCodCta ActxCta 
         Height          =   390
         Left            =   120
         TabIndex        =   69
         Top             =   1920
         Width           =   3720
         _extentx        =   6562
         _extenty        =   688
         texto           =   "Credito :"
         enabledcmac     =   -1
         enabledcta      =   -1
         enabledprod     =   -1
         enabledage      =   -1
      End
      Begin VB.Label lblGastos 
         Caption         =   "Gastos :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   70
         Top             =   2010
         Width           =   855
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
         Left            =   8190
         TabIndex        =   36
         Top             =   4860
         Width           =   885
      End
      Begin VB.Label Label12 
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
         Left            =   8190
         TabIndex        =   35
         Top             =   5175
         Width           =   915
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
         Left            =   6240
         TabIndex        =   30
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label Label14 
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
         Left            =   6360
         TabIndex        =   29
         Top             =   2880
         Width           =   315
      End
      Begin VB.Label Label11 
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
         Left            =   6120
         TabIndex        =   28
         Top             =   2520
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   "Servicios"
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
         Left            =   150
         TabIndex        =   27
         Top             =   0
         Width           =   795
      End
      Begin VB.Shape ShapeIGV 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   8070
         Top             =   4785
         Width           =   2385
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   8070
         Top             =   5115
         Width           =   2385
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Left            =   6090
         TabIndex        =   34
         Top             =   2445
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   8040
      TabIndex        =   13
      Top             =   8370
      Width           =   1230
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   9270
      TabIndex        =   14
      Top             =   8370
      Width           =   1230
   End
   Begin VB.Frame Frame3 
      Caption         =   "Glosa"
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
      Height          =   915
      Left            =   75
      TabIndex        =   16
      Top             =   1725
      Width           =   5985
      Begin VB.TextBox txtMovDesc 
         Height          =   630
         Left            =   90
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   195
         Width           =   5730
      End
   End
   Begin VB.Frame frameDestino 
      Caption         =   "Proveedor"
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
      Height          =   690
      Left            =   75
      TabIndex        =   20
      Top             =   1035
      Width           =   5985
      Begin Sicmact.TxtBuscar txtBuscarProv 
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   1725
         _extentx        =   3043
         _extenty        =   635
         appearance      =   1
         appearance      =   1
         font            =   "frmOpeDocChica.frx":0E43
         appearance      =   1
         tipobusqueda    =   3
         tipobuspers     =   2
      End
      Begin VB.Label lblProvNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1875
         TabIndex        =   15
         Top             =   225
         Width           =   3915
      End
   End
   Begin VB.Frame frameCaja 
      Caption         =   "CAJA CHICA"
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
      Height          =   1020
      Left            =   75
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   5835
      Begin VB.ComboBox cboAreaSoli 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   600
         Width           =   4440
      End
      Begin VB.TextBox lblCajaChicaDesc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   217
         Width           =   3885
      End
      Begin Sicmact.TxtBuscar txtBuscarAreaCH 
         Height          =   345
         Left            =   135
         TabIndex        =   0
         Top             =   210
         Width           =   1095
         _extentx        =   1535
         _extenty        =   609
         appearance      =   1
         appearance      =   1
         font            =   "frmOpeDocChica.frx":0E6F
         appearance      =   1
      End
      Begin VB.Label lblAreaSoli 
         AutoSize        =   -1  'True
         Caption         =   "Area Solicitante:"
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   650
         Width           =   1155
      End
      Begin VB.Label lblNroProc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   5175
         TabIndex        =   56
         Top             =   210
         Width           =   570
      End
   End
   Begin VB.Frame fraArendir 
      Caption         =   "Recibo de A rendir"
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
      Left            =   75
      TabIndex        =   49
      Top             =   0
      Width           =   5835
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Left            =   3870
         TabIndex        =   55
         Top             =   285
         Width           =   450
      End
      Begin VB.Label lblTotalArendir 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4470
         TabIndex        =   54
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nro"
         Height          =   195
         Left            =   75
         TabIndex        =   53
         Top             =   285
         Width           =   255
      End
      Begin VB.Label lblArendirNro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   405
         TabIndex        =   52
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label lblFechaArendir 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2685
         TabIndex        =   51
         Top             =   225
         Width           =   1065
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   2055
         TabIndex        =   50
         Top             =   285
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmOpeDocChica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lMN As Boolean, sMoney As String
Dim cCtaCodPend As String
Dim lTransActiva As Boolean
Dim nTasaIGV As Currency, nTasaImp As Currency
Dim nVariaIGV As Currency

Dim aCtaCambio(1, 2) As String
Dim sObjRendir As String
Dim sConcepCod As String
Dim nConcepNiv As Integer
Dim sCtaViaticos As String
Dim sCtaViaticosDesc As String

'''********************************************************************
Dim vsDocAbrev As String
Dim vsDocNro As String
Dim vsProveedor As String
Dim vdFechaDoc As Date
Dim vnImporteDoc As Currency
Dim vsMovNroSust As String
Dim vsMovDesc As String

Dim oContFunct As NContFunciones
Dim oNArendir As NARendir
Dim oOpe As DOperacion
Dim oCtaCont As DCtaCont
Dim lnArendirFase As ARendirFases
Dim lnTipoArendir As ArendirTipo
Dim lbCajaChica As Boolean
Dim lnTotalArendir As Currency
Dim lsCtaCodPend As String
Dim lbNewProv As Boolean
Dim lbTieneIGV As Boolean
Dim lsCtaCodProvis As String
Dim lsMovNroAten As String
Dim lnTempSaldo As Currency

Dim OK As Boolean
Dim lSalir As Boolean
Dim lsAgeCod As String
Dim lsAgeDes As String
Dim lsAreaCod As String
Dim lsAreaDes As String
Dim lsPersCod As String
Dim lsPersNom As String
Dim lsFechaARendir As String
Dim lsMovNroSol As String
Dim lsNroArendir As String
Dim lsAreaCh As String
Dim lsAgeCh As String
Dim lnNroProc As Integer

Dim lsCtaITFD As String
Dim lsCtaITFH As String

Dim bAlert As Integer

Private m_Pos As Integer '*** PEAC 20121119
Dim bSale As Boolean
Dim Frase As String
Dim PalabraFrase As String
Dim lsCodigoPersona As String '*** PEAC 20121122

Dim sMatrizDatos() As Variant
Dim nTipoMatriz As Integer
Dim nCantAgeSel As Integer
Dim i As Integer
Dim lcCodObjArea As String
'EJVG20140727 ***
Dim fnRetProvSPMontoBase As Currency
Dim fnRetProvSPAporte As Currency
Dim fnRetProvSPComisionAFP As Currency
Dim fnRetProvSPSeguroAFP As Currency
Dim fsCtaContAporteAFP As String, fsCtaContSeguroAFP As String, fsCtaContComisionAFP As String
Dim fsCtaContAporteONP As String
'END EJVG *******
'PASI20140813
Dim rsProv As ADODB.Recordset
'end PASI
'ARLO20170208****
Dim objPista As COMManejador.Pista
'************


Public Sub Inicio(ByVal pnArendirFase As ARendirFases, ByVal pnTipoArendir As ArendirTipo, _
                    ByVal psAreaCod As String, ByVal psAreaDes As String, _
                    ByVal psAgeCod As String, ByVal psAgeDes As String, _
                    ByVal psPersCod As String, ByVal psPersNom As String, _
                    ByVal psFechaARendir As String, ByVal psMovNroAten As String, _
                    ByVal psMovNroSol As String, ByVal psNroArendir As String, _
                    Optional ByVal pnTempSaldo As Currency, Optional pbCajaChica As Boolean = False, _
                    Optional pnTotalARendir As Currency = 0, _
                    Optional psAreaCh As String = "", Optional psAgeCh As String = "", _
                    Optional pnNroProc As Integer = 0, _
                    Optional psGlosa As String = "")
                    '***Parametro psGlosa agregado por ELRO el 20130221, según SATI INC1301300007
                    

lsAreaCh = psAreaCh
lsAgeCh = psAgeCh
lnNroProc = pnNroProc
                   
lsNroArendir = psNroArendir
lsAgeCod = psAgeCod
lsAgeDes = psAgeDes
lsAreaCod = psAreaCod
lsAreaDes = psAreaDes
lsMovNroAten = psMovNroAten
lsPersCod = psPersCod
lsPersNom = psPersNom
lnArendirFase = pnArendirFase
lnTipoArendir = pnTipoArendir
lbCajaChica = pbCajaChica
lsFechaARendir = psFechaARendir
lsMovNroSol = psMovNroSol
'*********************************************
'Agregado por PASI20131111 segun TI-ERS107-2013
ChkVisitaAge.Visible = False
'*********************************************

'edpyme
If CStr(pnTempSaldo) <> "" Then
    lnTempSaldo = pnTempSaldo
    bAlert = 1
Else
    bAlert = 0
End If

If pnArendirFase = ArendirRendicion Or pnArendirFase = ArendirSustentacion Then
   lnTotalArendir = pnTotalARendir
End If
txtMovDesc = psGlosa '***Agregado por ELRO el 20130221, según SATI INC1301300007
Me.Show 1
End Sub

Public Sub InicioEgresoDirecto()
lbCajaChica = True
Me.Show 1
End Sub
Private Sub FormatoImpuesto()
fgImp.ColWidth(0) = 250
fgImp.ColWidth(1) = 750
fgImp.ColWidth(2) = 550
fgImp.ColWidth(3) = 0    'CtaContCod
fgImp.ColWidth(4) = 0    'CtaContDes
fgImp.ColWidth(5) = 0    'D/H
fgImp.ColWidth(6) = 1200
fgImp.ColWidth(7) = 0 'Destino 0/1
fgImp.ColWidth(8) = 0 'Obligatorio, Opcional 1/2
fgImp.ColWidth(9) = 0 'Total Impuesto no Gravado
fgImp.TextMatrix(0, 1) = ""
fgImp.TextMatrix(0, 2) = "Tasa"
fgImp.TextMatrix(0, 6) = "Monto"
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

Private Sub CalculaTotal(Optional lCalcImpuestos As Boolean = True)
    Dim n As Integer, m As Integer
    Dim nSTot As Currency
    Dim nITot As Currency, nImp As Currency
    Dim nTot  As Currency
    Dim lnTotalFIla As Currency
    Dim lnI As Integer
    Dim lnMontoRetencionProvSP As Currency 'EJVG20140728

    nSTot = 0: nTot = 0
    If fgImp.TextMatrix(1, 1) = "" Then
       lCalcImpuestos = False
    End If
    For m = 1 To fgImp.Rows - 1
       nITot = 0
       For n = 1 To fgDetalle.Rows - 1
          If fgImp.TextMatrix(m, 2) = "." Then
             If lCalcImpuestos Then
                If fgImp.TextMatrix(m, 11) = "0" Then
                    
                    If m = 1 Then
                       nImp = Round(Val(Format(fgDetalle.TextMatrix(n, 3), gsFormatoNumeroDato)) * Val(Format(fgImp.TextMatrix(m, 4), gsFormatoNumeroDato)) / 100, 2)
                       fgDetalle.TextMatrix(n, m + 6) = Format(nImp, gsFormatoNumeroView)
                    End If
                    If m = 2 And fgImp.TextMatrix(m, 11) = "0" Then
                       nImp = Round(Val(Format(fgDetalle.TextMatrix(n, 8), gsFormatoNumeroDato)), 2)
                    End If
                    
                Else
                    lnTotalFIla = Val(Format(fgDetalle.TextMatrix(n, 3), gsFormatoNumeroDato))

                    For lnI = 7 To Me.fgDetalle.Cols - 1
                        If fgImp.TextMatrix(lnI - 7 + 1, 11) = "0" And fgDetalle.TextMatrix(n, lnI) <> "" Then
                        lnTotalFIla = lnTotalFIla + Val(Format(fgDetalle.TextMatrix(n, lnI), gsFormatoNumeroDato))
                        End If
                    Next lnI

                    nImp = Round(Val(lnTotalFIla) * Val(Format(fgImp.TextMatrix(m, 4), gsFormatoNumeroDato)) / 100, 2)
                    fgDetalle.TextMatrix(n, m + 6) = Format(nImp, gsFormatoNumeroView)
                End If
             Else
                nImp = IIf(fgDetalle.TextMatrix(n, m + 6) = "", 0, fgDetalle.TextMatrix(n, m + 6))
             End If
             nITot = nITot + nImp
          Else
             If lCalcImpuestos Then fgDetalle.TextMatrix(n, m + 6) = "0.00"
          End If
       Next
       
       fgImp.TextMatrix(m, 5) = Format(nITot, gsFormatoNumeroView)
       nTot = nTot + nITot * IIf(fgImp.TextMatrix(m, 8) = "D", 1, -1)
    Next
    For n = 1 To fgDetalle.Rows - 1
       nSTot = nSTot + Val(Format(fgDetalle.TextMatrix(n, 3), gsFormatoNumeroDato))
    Next
    txtSTotal = Format(nSTot, gsFormatoNumeroView)
    'EJVG20140728 ***
    fnRetProvSPMontoBase = 0#: fnRetProvSPAporte = 0#: fnRetProvSPSeguroAFP = 0#: fnRetProvSPComisionAFP = 0#
    If gbBitRetencSistPensProv Then
        If Val(Trim(Left(cboDoc.Text, 3))) = TpoDoc.TpoDocRecHonorarios Then
            If IsDate(txtDocFecha.Text) And txtBuscarProv.psCodigoPersona <> "" Then
                Dim oPSP As New NProveedorSistPens
                fnRetProvSPMontoBase = MontoBaseOperacion
                If oPSP.ExisteDatosSistemaPension(txtBuscarProv.psCodigoPersona) Then
                    If oPSP.AplicaRetencionSistemaPension(txtBuscarProv.psCodigoPersona, CDate(txtDocFecha.Text), fnRetProvSPMontoBase) Then
                        oPSP.SetDatosRetencionSistPens txtBuscarProv.psCodigoPersona, CDate(txtDocFecha.Text), fnRetProvSPMontoBase, Mid(gsOpeCod, 3, 1), fnRetProvSPAporte, fnRetProvSPSeguroAFP, fnRetProvSPComisionAFP
                    End If
                End If
                Set oPSP = Nothing
            End If
        End If
    End If
    lnMontoRetencionProvSP = fnRetProvSPAporte + fnRetProvSPSeguroAFP + fnRetProvSPComisionAFP
    txtRetSistPens.Text = Format(lnMontoRetencionProvSP, gsFormatoNumeroView)
    'txtTotal = Format(nSTot + nTot, gsFormatoNumeroView)
    txtTotal = Format(nSTot + nTot - lnMontoRetencionProvSP, gsFormatoNumeroView)
    'END EJVG *******
End Sub



Private Function ValidaCajaChica() As Boolean
Dim oCajaChica As nCajaChica
Dim lnSaldo As Currency
Set oCajaChica = New nCajaChica
Dim lnTope As Currency
Dim lnMAsig As Currency

ValidaCajaChica = True
lnSaldo = oCajaChica.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2))
If lnSaldo = 0 Then
    MsgBox "Caja Chica Sin Saldo. Es necesario Solicitar Autorización o Desembolso", vbInformation, "Aviso"
    ValidaCajaChica = False
    Exit Function
End If
lnTope = oCajaChica.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), MontoTope)
If CCur(txtTotal) > lnTope Then
    MsgBox "El Importe para solicitar a Caja Chica no puede ser mayor a " & Format(lnTope, gsFormatoNumeroView) & ". " & oImpresora.gPrnSaltoLinea & "En Caso Contrario Solicite A rendir Cuenta con Caja General", vbInformation, "Aviso"
    ValidaCajaChica = False
    Exit Function
End If
'--------John-----
'lnMAsig = oCajaChica.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), MontoAsig)
'If lnMAsig = 0 Or IsNull(lnMAsig) Then
'   MsgBox "Caja chica no tiene Monto Asignado", vbInformation, "Aviso"
'   ValidaCajaChica = False
'   Exit Function
'End If
'If CCur(txtTotal) > 10 * lnMAsig / 100 Then
'    MsgBox "El Importe a solicitar No puede ser mayor al 10 % :  " & 10 * lnMAsig / 100 & ". " & oImpresora.gPrnSaltoLinea, vbInformation, "Aviso"
'    ValidaCajaChica = False
'    Exit Function
'End If

'---------------------

If oCajaChica.VerificaTopeCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2)) = True Then
    MsgBox "No puede realizar Egreso porque el saldo de esta Caja Chica es menor que el permitido." & oImpresora.gPrnSaltoLinea _
            & "Por favor es necesario que realice Rendición", vbInformation, "Aviso"
    ValidaCajaChica = False
    If MsgBox(" ¿ Desea continuar con solicitud de Egreso de Caja Chica ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
         ValidaCajaChica = True
    End If
    Exit Function
End If
If lnSaldo < CCur(txtTotal) Then
    MsgBox "Egreso no puede ser mayor que " & Format(lnSaldo, gsFormatoNumeroView), vbInformation, "Aviso"
    ValidaCajaChica = False
    Exit Function
End If
End Function

'Private Sub FormatoDetalle()
'fgDetalle.TextMatrix(0, 0) = "#"
'fgDetalle.TextMatrix(0, 1) = "Código"
'fgDetalle.TextMatrix(0, 2) = "Descripción"
'fgDetalle.TextMatrix(0, 3) = "Monto"
'fgDetalle.TextMatrix(0, 4) = "D/H"
'fgDetalle.TextMatrix(0, 5) = "Gravado"
'fgDetalle.TextMatrix(0, 6) = ""
'fgDetalle.ColWidth(0) = 380
'fgDetalle.ColWidth(1) = 1300
'fgDetalle.ColWidth(2) = 3470
'fgDetalle.ColWidth(3) = 1200
'fgDetalle.ColWidth(4) = 0
'fgDetalle.ColWidth(5) = 0
'fgDetalle.ColWidth(6) = 0
'
'fgDetalle.ColAlignment(0) = 4
'fgDetalle.ColAlignment(1) = 1
'fgDetalle.ColAlignment(3) = 7
'fgDetalle.ColAlignmentFixed(0) = 4
'fgDetalle.ColAlignmentFixed(3) = 7
'fgDetalle.Row = 1
'fgDetalle.Col = 1
'fgDetalle.RowHeight(-1) = 285
'End Sub
Private Sub cboDoc_KeyPress(KeyAscii As Integer)
Dim n As Integer
Dim lvItem As ListItem
If KeyAscii = 13 Then
   cboDoc_Click
   If txtDocSerie.Enabled Then
      txtDocSerie.SetFocus
   Else
      txtDocFecha.SetFocus
   End If
End If
End Sub
Private Sub cboDoc_Click()
Dim rs As ADODB.Recordset
Dim oDoc As DDocumento
Dim oConst As NConstSistemas
Set oConst = New NConstSistemas

If cboDoc.ListIndex = -1 Then
    Exit Sub
End If
Set oDoc = New DDocumento

Dim nRow As Integer
   lbTieneIGV = False
   fgDetalle.Cols = 7
   fgImp.Clear
   fgImp.FormaCabecera
   fgImp.Rows = 2
   fgDetalle.ColumnasAEditar = Left(fgDetalle.ColumnasAEditar, 13)
   fgDetalle.EncabezadosAlineacion = Left(fgDetalle.EncabezadosAlineacion, 13)
   fgDetalle.FormatosEdit = Left(fgDetalle.FormatosEdit, 13)
   fgDetalle.ListaControles = Left(fgDetalle.ListaControles, 13)
   Set rs = New ADODB.Recordset
   Set rs = oDoc.CargaDocImpuesto(Trim(Mid(cboDoc.Text, 1, 3)))
   Do While Not rs.EOF
      'Primero adicionamos Columna de Impuesto
        fgDetalle.Cols = fgDetalle.Cols + 1
        fgDetalle.ColWidth(fgDetalle.Cols - 1) = 1200
        fgDetalle.ColumnasAEditar = fgDetalle.ColumnasAEditar & "-" & fgDetalle.Cols - 1
        fgDetalle.EncabezadosAlineacion = fgDetalle.EncabezadosAlineacion & "-R"
        fgDetalle.FormatosEdit = fgDetalle.FormatosEdit & "-2"
        fgDetalle.ListaControles = fgDetalle.ListaControles & "-0"
        fgDetalle.TextMatrix(0, fgDetalle.Cols - 1) = rs!cImpAbrev

       'Adicionamos los impuestos en el grid de impuestos
        fgImp.AdicionaFila
        fgImp.col = 0
        nRow = fgImp.row
        fgImp.TextMatrix(nRow, 1) = "."
        If rs!cDocImpOpc = "1" Then
            'activamos el check de impuesto enviandole el valor "1"
            If cboDocDestino.ListIndex <> 3 Then fgImp.TextMatrix(nRow, 2) = "1"
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
   fgImp.col = 1
   If lbTieneIGV = False Then
      cboDocDestino.ListIndex = -1
      cboDocDestino.Enabled = False
   Else
      cboDocDestino.Enabled = True
      cboDocDestino.ListIndex = oConst.LeeConstSistema(gConstSistDestinoIGVDefecto)
   End If
   VerReciboEgreso
   'EJVG20140728 ***
    If gbBitRetencSistPensProv Then
        If Val(Trim(Left(cboDoc.Text, 3))) = TpoDoc.TpoDocRecHonorarios Then
            fraRetencSistPens.Visible = True
        Else
            fraRetencSistPens.Visible = False
        End If
    End If
    'END EJVG *******
   CalculaTotal
    '***Agregado por ELRO el 20130111, según OYP-RFC126-2012
   If gsOpeCod = gCHArendirCtaSustMN Or gsOpeCod = gCHArendirCtaSustME Then
    If Trim(cboDoc.Text) <> "" Then
      If Left(Trim(cboDoc.Text), 2) = "97" Then
         Frame3.Caption = "Motivo de Desplazamiento"
         FRLugarMovilidad.Visible = True
         txtLugar.Text = ""
      Else
         Frame3.Caption = "Glosa"
         FRLugarMovilidad.Visible = False
         txtLugar.Text = ""
      End If
    End If
   End If
   '***Fin Agregado por ELRO el 20130111******************
Set oDoc = Nothing
End Sub
Private Sub VerReciboEgreso()
'*************************************************************************
' Comentado, Cusco no quiere que se genere numero autonumerico para Recibos de Egresos
'*************************************************************************
'Dim lsReciboEgreso As String
'   If Mid(cboDoc.Text, 1, 2) = TpoDocRecEgreso Then
'      lsReciboEgreso = oContFunct.GeneraDocNro(TpoDocRecEgreso, gsCodUser, Mid(gsOpeCod, 3, 1))
'      txtDocSerie.MaxLength = 4
'      txtDocSerie = Mid(lsReciboEgreso, 1, 4)
'      txtDocSerie.Enabled = False
'      txtDocNro = Mid(lsReciboEgreso, 6, 20)
'      txtDocNro.Enabled = False
'   Else
'      txtDocSerie.MaxLength = 3
'      txtDocSerie = ""
'      txtDocSerie.Enabled = True
'      txtDocNro = ""
'      txtDocNro.Enabled = True
'   End If
    
      txtDocSerie.MaxLength = 4
      txtDocSerie = ""
      txtDocSerie.Enabled = True
      txtDocNro.MaxLength = 11 '***Agregado por ELRO el 20130603, según SATI INC1304290006
      txtDocNro = ""
      txtDocNro.Enabled = True

End Sub

Private Sub cboDoc_LostFocus()
    '*** PEAC 20121121
    'If (gsOpeCod = "401240" Or gsOpeCod = "402240" Or gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401344") And Trim(Left(cboDoc, 3)) = "93" Then
    If Trim(Left(cboDoc, 3)) = "93" Then
        txtDocSerie = Trim(Str(Year(Date)))
        txtDocNro = BuscaNumCorrelaDocDDJJ(txtDocSerie, Trim(Left(cboDoc, 3)))
    End If
    '*** FIN PEAC
    'If Trim(Left(cboDoc, 3)) = "5" Then
        'txtDocSerie = Trim(Str("3"))
    'End If '***NAGL ERS012-2017 20170710 Comentado by NAGL 20170927
End Sub

Private Sub cboDoc_Validate(Cancel As Boolean)
If cboDoc = "" Then
    Cancel = True
End If
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
If cboDocDestino.ListIndex <> 1 Then
    fgDetalle.ColWidth(5) = 0
    For i = 1 To fgDetalle.Rows - 1
        'limpia las posibles gravaciones manuales realizadas
        If fgDetalle.TextMatrix(i, 5) = "." Then
           fgDetalle.TextMatrix(i, 5) = ""
       End If
    Next
Else
    fgDetalle.ColWidth(5) = 500
End If
End Sub




Private Sub chkDocDolares_Click()
    If Me.chkDocDolares.value = 1 Then
        FrameTipCambio.Visible = True
        txtTipFijo = gnTipCambio
        If gbBitTCPonderado Then
           Label9.Caption = "Ponder."
           txtTipVariable = Format(gnTipCambioPonderadoVenta, gsFormatoNumeroView3Dec)
           txtTipPondVenta = Format(gnTipCambioPonderadoVenta, gsFormatoNumeroView3Dec)
        Else
           txtTipVariable = Format(gnTipCambioV, gsFormatoNumeroView3Dec)
           txtTipPondVenta = Format(gnTipCambioV, gsFormatoNumeroView3Dec)
        End If
    Else
        FrameTipCambio.Visible = False
    End If
    Me.fgDetalle.EliminaFila 1
End Sub

Private Sub ChkVisitaAge_Click()
    Dim lnContaItem As Integer
    
    If fgDetalle.TextMatrix(1, 2) = "" Then
        ChkVisitaAge.value = 0
        Exit Sub
    End If
    
    If ChkVisitaAge.value = 1 Then
        '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
        'If (gsOpeCod = "401140" Or gsOpeCod = "402140") And fgDetalle.TextMatrix(fgDetalle.Row, 0) = "" Then
        If (gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401141" Or gsOpeCod = "402141") And fgDetalle.TextMatrix(fgDetalle.row, 0) = "" Then
        '***Fin Modificado por ELRO el 20120921*******************
            MsgBox "Por favor ingrese primero un gasto para realizar la distribución manual.", vbOKOnly + vbExclamation, "Atención"
            ChkVisitaAge.value = 0
            Exit Sub
        '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
        'ElseIf (gsOpeCod = "401140" Or gsOpeCod = "402140") And fgDetalle.TextMatrix(fgDetalle.Row, 0) <> "" And Me.txtSTotal.Text = 0 Then
        ElseIf (gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401141" Or gsOpeCod = "402141") And fgDetalle.TextMatrix(fgDetalle.row, 0) <> "" And Me.txtSTotal.Text = 0 Then
        '***Fin Modificado por ELRO el 20120921*******************
            MsgBox "Por favor ingrese el monto del gasto registrado.", vbOKOnly + vbExclamation, "Atención"
            ChkVisitaAge.value = 0
            Exit Sub
        End If
        
        '********************************************
        For i = 1 To fgObj.Rows - 1
            lcCodObjArea = Me.fgObj.TextMatrix(i, 2)
        Next
        '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
        'If Left(lcCodObjArea, 3) = "067" Or Left(lcCodObjArea, 3) = "041" Or Left(lcCodObjArea, 3) = "026" And (gsOpeCod = "401371" Or gsOpeCod = "402371") Then
        If Left(lcCodObjArea, 3) = "067" Or Left(lcCodObjArea, 3) = "041" Or Left(lcCodObjArea, 3) = "026" And (gsOpeCod = "401371" Or gsOpeCod = "402371" Or gsOpeCod = "401344") Then
        '***Fin Modificado por ELRO el 20120921*******************
            MsgBox "Este item no se puede distribuir porque no pertenece a un area Administrativa.", vbOKOnly + vbExclamation, "Atención"
            ChkVisitaAge.value = 0
            Exit Sub
        End If
        
        '********************************************
        lnContaItem = 0
        For i = 1 To fgDetalle.Rows - 1
            lnContaItem = lnContaItem + 1
        Next
        '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
        'If lnContaItem > 1 And (gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401371" Or gsOpeCod = "402371") Then
        If lnContaItem > 1 And (gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401141" Or gsOpeCod = "402141" Or gsOpeCod = "401371" Or gsOpeCod = "402371" Or gsOpeCod = "401344") Then
        '***Fin Modificado por ELRO el 20120921*******************
            MsgBox "Para realizar la distribución manual solo debe registrarse un item.", vbOKOnly + vbExclamation, "Atención"
            ChkVisitaAge.value = 0
            Exit Sub
        End If

        '********************************************
        Call frmViaticoVisitaAgencias.Inicio(sMatrizDatos, nTipoMatriz, nCantAgeSel, Me.txtSTotal.Text)
        If nCantAgeSel = 0 Then
            ChkVisitaAge.value = 0
        End If
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim n As Integer 'Contadorl
Dim m As Integer
Dim nItem As Integer, nCol  As Integer
Dim sTexto As String
Dim nImporte As Currency
Dim nImpPend As Currency
Dim lsMovUltAct As String
Dim lsMovNro As String
Dim ldFecha As Date
Dim lsDocNro As String
Dim lsReciboEgreso As String
Dim lsDocAbrev As String
Dim lnSal As Currency
Dim lnTop As Currency
Dim oNCaja As nCajaChica
'ALPA 20100412***********
Dim sAgeTipo As String
'************************
Dim lsCtasInexis As String '', N As Integer, M As Integer
Dim lnMontoTre As Currency, lnMontoSie  As Currency
'MADM 20110505
Dim lrDatCredito As New ADODB.Recordset
Dim lsFechaHoraGrab As String
Dim lnNroGastoCta As Integer
Dim lnConceptoGasto As Long
Dim lsFechaAsigna As String
Dim lnMontoAntes As Currency
Dim lnMonto As Currency
Dim lsMotivo As String
Dim fnUltimoGasto As Currency
Dim fnSaldoGasto As Currency
Dim lnNewSaldoGasto As Currency
Dim pnMovNro As Long
'EJVG20140727 ***
Dim frm As frmProveedorRegSistemaPension
Dim bPSPSinDatos As Boolean
Dim oPSP As NProveedorSistPens
Dim lnRetProvSPMontoBase As Currency
Dim lnRetProvSPAporte As Currency
Dim lnRetProvSPComisionAFP As Currency
Dim lnRetProvSPSeguroAFP As Currency
Dim lnRetProvSPTotal As Currency
Dim lnTpoSistPensProv As TipoSistemaPensionProveeedor
Dim bMeterMatrizPSP As Boolean
Dim loMatDatosRetencSistPensProv As Variant
'END EJVG *******

'MADM 20110505

On Error GoTo ErrAceptar

If Not CamposOk Then Exit Sub

'*** PEAC 20121122
'If (gsOpeCod = "401240" Or gsOpeCod = "402240" Or gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401344") And Trim(Left(cboDoc, 3)) = "93" Then 'si es doc. Declarac. jurada
If Trim(Left(cboDoc, 3)) = "93" Then 'si es doc. Declarac. jurada
    lsCodigoPersona = ""
Else
    lsCodigoPersona = txtBuscarProv.psCodigoPersona
End If
'*** FIN PEAC

Dim oDoc As DDocumento
Set oDoc = New DDocumento
If oDoc.GetValidaDocProv(lsCodigoPersona, Left(cboDoc, 2), txtDocSerie & "-" & txtDocNro) = True Then
    '*** PEAC 20121122
    'If (gsOpeCod = "401240" Or gsOpeCod = "402240" Or gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401344") And Trim(Left(cboDoc, 3)) = "93" Then
    If Trim(Left(cboDoc, 3)) = "93" Then
        '***Modificado por ELRO el 20130603, según SATI INC1304290006****
        'txtDocNro = Right("00000000000" + Trim(Str(Val(txtDocNro) + 1)), 11)
        txtDocNro = Trim(Str(Val(txtDocNro) + 1))
        '***Fin Modificado por ELRO el 20130603, según SATI INC1304290006
        MsgBox "El número de DDJJ acaba de ser registrado por otro usuario, por lo que se incrementó el número automaticamente, por favor vuelva a grabar.", vbInformation, "Aviso"
    Else
        MsgBox "Documento ya se encuentra registrado", vbInformation, "Aviso"
    End If
    'MsgBox "Documento ya se encuentra registrado", vbInformation, "Aviso"
    '*** FIN PEAC
    Exit Sub
End If

Set oDoc = Nothing
If lbCajaChica Then
   If ValidaCajaChica = False Then
      Exit Sub
   End If
End If

 If lnSal + lnTop > lnTop Then
    MsgBox "Monto Superior a Tope de Caja Chica", vbInformation, "Aviso"
    Exit Sub
 End If

'*** PEAC 20121216
    'If (gsOpeCod = "401240" Or gsOpeCod = "402240" Or gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401344") And Trim(Left(cboDoc, 3)) = "93" Then 'si es doc. Declarac. jurada
    If Trim(Left(cboDoc, 3)) = "93" Then 'si es doc. Declarac. jurada
        Call buscar_cadena(1)
        If bSale Then
        
            With txtMovDesc
                .SelStart = m_Pos - 1
                .SelLength = Len(PalabraFrase)
                .SetFocus
            End With
            Exit Sub
        End If
    End If
    
    'If (gsOpeCod = "401240" Or gsOpeCod = "402240" Or gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401344") And Trim(Left(cboDoc, 3)) = "93" And txtBuscarProv.psCodigoPersona <> lsPersCod Then
    If gsOpeCod <> "401344" Then 'PASI20140813
        If Trim(Left(cboDoc, 3)) = "93" And txtBuscarProv.psCodigoPersona <> lsPersCod Then
            MsgBox "El emisor debe ser la persona que fue beneficiaria de la entrega a rendir.", vbInformation, "Aviso"
        Exit Sub
        End If
    End If
'*** FIN PEAC

'*** PEAC 20100920
'***Modificado por ELRO el 20120921, según OYP-RFC111-2012
If gsOpeCod = "401371" Or gsOpeCod = "401344" Then
'***Fin Modificado por ELRO el 20120921*******************
    If Trim(Mid(cboDoc.Text, 1, 3)) = "2" Or Trim(Mid(cboDoc.Text, 1, 3)) = "80" Or Trim(Mid(cboDoc.Text, 1, 3)) = "95" Then
        lnMontoTre = 0: lnMontoSie = 0
        For n = 1 To fgDetalle.Rows - 1
            lnMontoTre = lnMontoTre + CCur(fgDetalle.TextMatrix(n, 3)) 'OTROS DSCTOS
            lnMontoSie = lnMontoSie + CCur(fgDetalle.TextMatrix(n, 7)) 'OTROS DSCTOS
        Next
        'If lnMontoTre - lnMontoSie <> CCur(Me.txtTotal) Then
        If lnMontoTre - lnMontoSie - (fnRetProvSPAporte + fnRetProvSPComisionAFP + fnRetProvSPSeguroAFP) <> CCur(Me.txtTotal) Then
            MsgBox "Se detectó incoherencias entre los montos ingresados y los totales, por favor revise.", vbOKOnly, "Atención"
            Exit Sub
        End If
    Else
        If fgDetalle.Cols = 9 Then
            lnMontoTre = 0: lnMontoSie = 0
            For n = 1 To fgDetalle.Rows - 1
                lnMontoTre = lnMontoTre + CCur(fgDetalle.TextMatrix(n, 3))
                lnMontoSie = lnMontoSie + CCur(fgDetalle.TextMatrix(n, 7))
            Next
            If lnMontoTre + lnMontoSie <> CCur(Me.txtTotal) Then
                MsgBox "Se detectó incoherencias entre los montos ingresados y los totales, por favor revise.", vbOKOnly, "Atención"
                Exit Sub
            End If
        End If
    End If
End If
'*** FIN PEAC

For i = 1 To fgObj.Rows - 1
    lcCodObjArea = Me.fgObj.TextMatrix(i, 2)
Next
'***Modificado por ELRO el 20120921, según OYP-RFC111-2012
'If Me.ChkVisitaAge.Visible = True And (gsOpeCod = "401240" Or gsOpeCod = "402240") Then  '*** PEAC 20110303
'If Me.ChkVisitaAge.Visible = True And (gsOpeCod = "401240" Or gsOpeCod = "402240" Or gsOpeCod = "401241" Or gsOpeCod = "402241") Then
If Me.ChkVisitaAge.Visible = True Then
'***Fin Modificado por ELRO el 20120921*******************
    If Me.ChkVisitaAge.value = 0 Then
        If MsgBox("¿Realizó viajes a Agencias?", vbYesNo + vbQuestion, "Pregunta") = vbYes Then
            MsgBox "Active la Check de ''Viajes a Agencias'' y seleccione las agencias visitadas."
            Me.ChkVisitaAge.SetFocus
            Exit Sub
        End If
    End If
'***Modificado por ELRO el 20120921, según OYP-RFC111-2012
'ElseIf Me.ChkVisitaAge.Visible = True And (gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401371" Or gsOpeCod = "402371") Then '*** PEAC 20110303
ElseIf Me.ChkVisitaAge.Visible = True And (gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401141" Or gsOpeCod = "402141" Or gsOpeCod = "401371" Or gsOpeCod = "402371" Or gsOpeCod = "401344") Then
'***Fin Modificado por ELRO el 20120921*******************
    If Me.ChkVisitaAge.value = 0 Then
        If Left(lcCodObjArea, 3) = "067" Or Left(lcCodObjArea, 3) = "041" Or Left(lcCodObjArea, 3) = "026" Then
        
        Else
            If MsgBox("¿Realizó Gastos para otras Agencias?", vbYesNo + vbQuestion, "Pregunta") = vbYes Then
                MsgBox "Active la Check de ''Distribución manual de Gastos'' y seleccione las agencias con sus respectivos montos."
                Me.ChkVisitaAge.SetFocus
                Exit Sub
            End If
        End If
    End If
End If

'***Agregado por ELRO el 20120809, según OYP-RFC015-2012, OYP-RFC016-2012 y OYP-RFC047-2012
If gsOpeCod = gCGArendirCtatSust2MN Or _
   gsOpeCod = gCGArendirCtatSust2ME Or _
   gsOpeCod = gCGArendirViatSust2MN Or _
   gsOpeCod = gCGArendirViatSust2ME Or _
   gsOpeCod = gCHArendirCtaSustMN Or _
   gsOpeCod = gCHArendirCtaSustME Then
    If lnTipoArendir = gArendirTipoCajaGeneral Or lnTipoArendir = gArendirTipoViaticos Or lnTipoArendir = gArendirTipoCajaChica Then
        If CCur(txtTotal) > lnTempSaldo Then
            MsgBox "El Total es mayor a Saldo: " & CStr(lnTempSaldo), vbInformation, "Aviso"
            Exit Sub
        End If
    End If
End If
'***Agregado por ELRO**********************************************************************
'ChkVisitaAge

    'EJVG20140724 ***
    If gbBitRetencSistPensProv Then
        If Val(Trim(Left(cboDoc.Text, 3))) = TpoDoc.TpoDocRecHonorarios Then
            lnRetProvSPMontoBase = MontoBaseOperacion
            Set oPSP = New NProveedorSistPens
            If oPSP.AplicaRetencionSistemaPension(txtBuscarProv.psCodigoPersona, CDate(txtDocFecha.Text), lnRetProvSPMontoBase) Then
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
                oPSP.SetDatosRetencionSistPens txtBuscarProv.psCodigoPersona, CDate(txtDocFecha.Text), lnRetProvSPMontoBase, Mid(gsOpeCod, 3, 1), lnRetProvSPAporte, lnRetProvSPSeguroAFP, lnRetProvSPComisionAFP, lnTpoSistPensProv
                Set oPSP = Nothing
                If lnRetProvSPAporte <> fnRetProvSPAporte Or lnRetProvSPSeguroAFP <> fnRetProvSPSeguroAFP Or lnRetProvSPComisionAFP <> fnRetProvSPComisionAFP Then
                    MsgBox "El monto de retención de SNP/ONP debe ser recalculado, de click en actualizar SNP/ONP", vbInformation, "Aviso"
                    Exit Sub
                End If
                If fnRetProvSPAporte < 0# Or fnRetProvSPSeguroAFP < 0# Or fnRetProvSPComisionAFP < 0# Then
                    MsgBox "Uno de los conceptos de la Retención es negativo, verifique." & Chr(10) & "Si el problema persiste comuniquese con el Dpto. de TI", vbExclamation, "Aviso"
                    Exit Sub
                End If
                lnRetProvSPTotal = fnRetProvSPAporte + fnRetProvSPSeguroAFP + fnRetProvSPComisionAFP
                ReDim loMatDatosRetencSistPensProv(6)
                loMatDatosRetencSistPensProv(1) = fnRetProvSPAporte
                loMatDatosRetencSistPensProv(2) = ReemplazaCtaCont(IIf(lnTpoSistPensProv = AFP, fsCtaContAporteAFP, fsCtaContAporteONP), Mid(gsOpeCod, 3, 1))
                loMatDatosRetencSistPensProv(3) = fnRetProvSPSeguroAFP
                loMatDatosRetencSistPensProv(4) = ReemplazaCtaCont(fsCtaContSeguroAFP, Mid(gsOpeCod, 3, 1))
                loMatDatosRetencSistPensProv(5) = fnRetProvSPComisionAFP
                loMatDatosRetencSistPensProv(6) = ReemplazaCtaCont(fsCtaContComisionAFP, Mid(gsOpeCod, 3, 1))
            End If
            Set oPSP = Nothing
        End If
    End If
    'END EJVG *******

If Me.chkDocDolares.value = 1 Then
    If MsgBox(" ¿ Seguro de grabar Operación ?, el monto a descontar es " & Format(CCur(Me.txtTotal.Text) * CCur(Me.txtTipVariable.Text), gsFormatoNumeroView), vbYesNo + vbQuestion, "Aviso de Confirmación") = vbNo Then Exit Sub
Else
    If MsgBox(" ¿ Seguro de grabar Operación ? ", vbYesNo + vbQuestion, "Aviso de Confirmación") = vbNo Then Exit Sub
End If
Dim lsSubCta As String

lsMovNro = oContFunct.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
lsMovUltAct = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
nImpPend = 0

'*** PEAC 20100129 - aqui debe ser la cta 251419 y no la 1117
'Dim oOpe As New DOperacion

If gsOpeCod = "401322" Or gsOpeCod = "401371" Then
    lsCtaCodProvis = oOpe.EmiteOpeCta("401322", "D", "1") 'cta 251419 y no la cta 111701
Else
    lsCtaCodProvis = Trim(Right(cboProvis.Text, 20))  'cta 111701
End If

''ALPA 20100412************************
'sAgeTipo = ""
''lsCtaCodProvis = oOpe.EmiteOpeCta("401322", "D", "1") 'cta 251419
'lsCtaCodProvis = oOpe.EmiteOpeCta(gsOpeCod, "D", "1") 'cta 251419
'If Right(Trim(lsCtaCodProvis), 2) = "AG" Then
'    lsCtaCodProvis = Replace(lsCtaCodProvis, "AG", gsCodAge)
'    sAgeTipo = "AG"
'End If
''*************************************
'lsCtaCodProvis = Trim(Right(cboProvis.Text, 20)) 'cta 111701
'Set oOpe = Nothing
'*** FIN PEAC

If sAgeTipo = "" Then
lsSubCta = oContFunct.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaCodProvis, Trim(txtBuscarAreaCH), False)
' JEOM

If lsSubCta <> "" Then
       If Mid(txtBuscarAreaCH, 1, 3) = "067" And Mid(txtBuscarAreaCH, 4, 2) = "01" Then ' Espinar
              lsSubCta = "01"
              'lsCtaCodProvis = lsCtaCodProvis & "02" & lsSubCta
              lsCtaCodProvis = lsCtaCodProvis & lsSubCta '*** PEAC 20101210
       ElseIf Mid(txtBuscarAreaCH, 1, 3) = "042" Then  'Recuperaciones
            '**Modificado por DAOR 20080326 *******************************************
            'lsSubCta = "03"
            'lsCtaCodProvis = lsCtaCodProvis & "01" & lsSubCta
            lsSubCta = "01"
            'lsCtaCodProvis = lsCtaCodProvis & "02" & lsSubCta
            lsCtaCodProvis = lsCtaCodProvis & lsSubCta '*** PEAC 20101210
            '**************************************************************************
       ElseIf Mid(txtBuscarAreaCH, 1, 3) = "023" Or Mid(txtBuscarAreaCH, 1, 3) = "043" Then  'Logistica y Secretaria GITU
              lsSubCta = "01"
              'lsCtaCodProvis = lsCtaCodProvis & "02" & lsSubCta
              lsCtaCodProvis = lsCtaCodProvis & lsSubCta '*** PEAC 20101210
       Else
              'lsCtaCodProvis = lsCtaCodProvis & IIf(CCur(lsSubCta) > 90, "01", "02") & lsSubCta
              lsCtaCodProvis = lsCtaCodProvis & lsSubCta
       End If
End If
End If
'Fin JEOM
'***************************
lsDocNro = txtDocSerie + "-" + txtDocNro
lsDocAbrev = Trim(Right(cboDoc, 3))

Dim lsCVIng As String
Dim lsCVEgr As String

lsCtaITFD = oOpe.EmiteOpeCta(gsOpeCod, "D", 2)
lsCtaITFH = oOpe.EmiteOpeCta(gsOpeCod, "H", 2)
      
    '****Modificado por ELRO el 20120924, según OYP-RFC111-20212
    'If lbCajaChica = False Then
    If lbCajaChica = False Or (lbCajaChica = True And gsOpeCod = "401344") Then
     '****Fin Modificado por ELRO el 20120924*******************
        If lnArendirFase = ArendirSustentacion Then
            
            If Me.chkDocDolares.value = 1 Then
                lsCVEgr = oOpe.EmiteOpeCta(gsOpeCod, "D", 8)
                lsCVIng = oOpe.EmiteOpeCta(gsOpeCod, "H", 8)
                
                If lsCVEgr = "" Or lsCVIng = "" Then
                    MsgBox "Cuentas de Diferencia de cambio no ingresadas.", vbInformation, "Aviso"
                    Exit Sub
                End If
            End If
            
            lsCVEgr = lsCVEgr & oContFunct.GetFiltroObjetos(13, lsCVEgr, gsCodArea & gsCodAge, False)
            lsCVIng = lsCVIng & oContFunct.GetFiltroObjetos(13, lsCVIng, gsCodArea & gsCodAge, False)
            'MADM 20110512 pnMovNro
            oNArendir.GrabaSustentacionArendir lbNewProv, txtBuscarProv.psCodigoPersona, gLogProvEstadoActivado, lsMovUltAct, lsMovNro, _
                                gsOpeCod, Trim(txtMovDesc), CCur(txtTotal), cboDocDestino.ListIndex, _
                                lsCtaCodProvis, lsCtaCodPend, lnTipoArendir, lnTotalArendir, lsDocNro, _
                                Left(cboDoc, 3), CDate(txtDocFecha), CCur(txtTotal), lsMovNroAten, lsMovNroSol, lsAgeCod, lsAreaCod, fgDetalle.GetRsNew, fgObj.GetRsNew, _
                                fgImp.GetRsNew, lsAreaCh, lsAgeCh, lnNroProc, nVal(txtTipVariable), lsCtaITFD, lsCtaITFH, gnImpITF, nVal(txtTipPondVenta), IIf(Me.chkDocDolares.value = 1, True, False), lsCVIng, lsCVEgr, lsCtasInexis, sMatrizDatos, nCantAgeSel, pnMovNro, _
                                txtLugar, lnRetProvSPMontoBase, lnRetProvSPTotal, lnTpoSistPensProv, loMatDatosRetencSistPensProv
            '***Parametro txtLugar agregado por ELRO el 20130111, según OYP-RFC126-2012
            lbNewProv = True
            OK = True
            
            '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
            'MADM 20110505
            'If (gsOpeCod = "401140" Or (gsOpeCod = "401371" And Trim(Me.txtBuscarAreaCH.Text) = "042") And pnMovNro <> 0) Then
            If (gsOpeCod = "401140" Or gsOpeCod = "401141" Or ((gsOpeCod = "401371" Or gsOpeCod = "401344") And Trim(Me.txtBuscarAreaCH.Text) = "042") And pnMovNro <> 0) Then
            '***Fin Modificado por ELRO el 20120921*******************
                If Me.ActxCta.Visible = True Then
                        fnUltimoGasto = 0
                        Set lrDatCredito = oOpe.dObtieneDatosAsignaGastoCredRecup(ActxCta.NroCuenta)
                        If Not (lrDatCredito.BOF And lrDatCredito.EOF) Then
                            fnUltimoGasto = lrDatCredito!nUltGasto
                            fnSaldoGasto = lrDatCredito!nSaldoGasto
                            
                            lnMontoAntes = 0
                            lsFechaAsigna = Format$(txtFecha.Text, "mm/dd/yyyy")
                            lnConceptoGasto = Trim(Left(Trim(cbogastos), 4))
                            lnNroGastoCta = fnUltimoGasto + 1
                            lnMonto = CCur(Format(Trim(txtTotal.Text), "###0.00"))
                            lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
                            lsMotivo = Me.txtMovDesc.Text ' "Gastos de Recuperaciones desde Financiero"
                            lnNewSaldoGasto = fnSaldoGasto + lnMonto - lnMontoAntes
                            
                            Call oOpe.nGastoRecupAsignaNuevo(ActxCta.NroCuenta, lsFechaHoraGrab, _
                                lsMovNro, lnNroGastoCta, lsFechaAsigna, lnConceptoGasto, lnMonto, _
                                 0, gColRecGastoEstPendiente, lsMotivo, lnNewSaldoGasto, False, pnMovNro)

                        
                        Else
                                   MsgBox "No se Pudo Agregar el Gasto al Crédito, deberá ser Agregado desde el SICMACT Negocio", vbCritical, "Aviso"
                        End If
                End If
            End If
            'END MADM
            
        End If
    Else
        Dim oCH As nCajaChica
        Set oCH = New nCajaChica
        'MADM 20110512 pnMovNro
        oCH.GrabaSolEgresoDirecto lbNewProv, txtBuscarProv.psCodigoPersona, gLogProvEstadoActivado, lsMovUltAct, lsMovNro, _
                                    gsOpeCod, Trim(txtMovDesc), IIf(Trim(Mid(cboDoc.Text, 1, 3)) = "1" And cboDocDestino.ListIndex = -1, 2, cboDocDestino.ListIndex), _
                                    lsCtaCodProvis, lsDocNro, Left(cboDoc, 3), CDate(txtDocFecha), CCur(txtTotal), _
                                    Mid(Me.txtBuscarAreaCH, 4, 2), Mid(Me.txtBuscarAreaCH, 1, 3), Val(lblNroProc), _
                                    fgDetalle.GetRsNew, fgObj.GetRsNew, _
                                    fgImp.GetRsNew, nVal(txtTipVariable), lsCtaITFD, lsCtaITFH, gnImpITF, lsCtasInexis, sMatrizDatos, nCantAgeSel, pnMovNro, lnRetProvSPMontoBase, lnRetProvSPTotal, lnTpoSistPensProv, loMatDatosRetencSistPensProv
        
        lbNewProv = True
        OK = True
        
        '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
        'MADM 20110505
        'If (gsOpeCod = "401140" Or (gsOpeCod = "401371" And Trim(Me.txtBuscarAreaCH.Text) = "042") And pnMovNro <> 0) Then
        If (gsOpeCod = "401140" Or gsOpeCod = "401141" Or ((gsOpeCod = "401371" Or gsOpeCod = "401344") And Trim(Me.txtBuscarAreaCH.Text) = "042") And pnMovNro <> 0) Then
        '***Fin Modificado por ELRO el 20120921*******************
                If Me.ActxCta.Visible = True Then
                        fnUltimoGasto = 0
                        Set lrDatCredito = oOpe.dObtieneDatosAsignaGastoCredRecup(ActxCta.NroCuenta)
                        If Not (lrDatCredito.BOF And lrDatCredito.EOF) Then
                            fnUltimoGasto = lrDatCredito!nUltGasto
                            fnSaldoGasto = lrDatCredito!nSaldoGasto
                            lnMontoAntes = 0
                            lsFechaAsigna = Format$(txtFecha.Text, "mm/dd/yyyy")
                            lnConceptoGasto = Trim(Left(Trim(cbogastos), 4))
                            lnNroGastoCta = fnUltimoGasto + 1
                            lnMonto = CCur(Format(Trim(txtTotal.Text), "###0.00"))
                            lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
                            lsMotivo = Me.txtMovDesc.Text ' "Gastos de Recuperaciones desde Financiero"
                            lnNewSaldoGasto = fnSaldoGasto + lnMonto - lnMontoAntes
                            
                            Call oOpe.nGastoRecupAsignaNuevo(ActxCta.NroCuenta, lsFechaHoraGrab, _
                                lsMovNro, lnNroGastoCta, lsFechaAsigna, lnConceptoGasto, lnMonto, _
                                 0, gColRecGastoEstPendiente, lsMotivo, lnNewSaldoGasto, False, pnMovNro)
                        Else
                             MsgBox "No se Pudo Agregar el Gasto al Crédito, deberá ser Agregado desde el SICMACT Negocio", vbCritical, "Aviso"
                        End If
                End If
            End If
            'END MADM
        Set oCH = Nothing
    End If
    
    pnMovNro = 0 'MADM 20110511
    nCantAgeSel = 0 '*** PEAC 20110304
    
    '*** PEAC 20101019
    If lsCtasInexis <> "" Then
        MsgBox "Las siguientes Ctas contables no figuran en el Plan Contable : " & Chr(10) & lsCtasInexis & Chr(10) & "Capture esta pantalla y enviásela a contabilidad, este mensaje no afecta este proceso.", vbOKOnly, "Aviso para Contabilidad"
    End If
    '*** FIN PEAC
    
    '*** PEAC 20121121
        txtDocSerie.Enabled = False
        txtDocNro.Enabled = False
    '*** FIN PEAC
    
    
    ldFecha = CDate(Mid(lsMovNro, 7, 2) & "/" & Mid(lsMovNro, 5, 2) & "/" & Mid(lsMovNro, 1, 4))
    If Left(cboDoc, 2) = Format(TpoDocRecEgreso, "00") Then
        Dim oContImp As NContImprimir
        Set oContImp = New NContImprimir
        lsReciboEgreso = oContImp.ImprimeReciboEgresos(gnColPage, lsMovNro, txtMovDesc, ldFecha, gsNomCmac, gsOpeCod, _
                        lbCajaChica, txtBuscarAreaCH & "-" & lblNroProc & " " & lblCajaChicaDesc, lnArendirFase, lsNroArendir, lsDocNro, CDate(txtDocFecha), _
                        txtBuscarProv.psCodigoPersona, lblProvNombre, "", CCur(txtTotal))
        EnviaPrevio lsReciboEgreso, Me.Caption, Int(gnLinPage / 2), False
        Set oContImp = Nothing
    End If
    '***Modificado por ELRO el 20120919, según OYP-RFC111-2012
    'If lbCajaChica = False Or lnTipoArendir = gArendirTipoViaticos Then
    If lbCajaChica = False Or lnTipoArendir = gArendirTipoViaticos Or (lbCajaChica = True Or gsOpeCod = "401344") Then
    '***Fin Modificado por ELRO el 20120919*******************
        vsDocAbrev = lsDocAbrev
        vsDocNro = lsDocNro
        vsProveedor = lblProvNombre
        
        vdFechaDoc = CDate(txtDocFecha)
        If Me.chkDocDolares.value = 1 Then
            vnImporteDoc = Format(CCur(txtTotal) * CCur(Me.txtTipVariable.Text), "#.00")
        Else
            vnImporteDoc = CCur(txtTotal)
        End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Proveedor : " & txtBuscarProv.psCodigoPersona
            Set objPista = Nothing
            '*******
        vsMovNroSust = lsMovNro
        vsMovDesc = txtMovDesc
        OK = True
        lSalir = True
        Unload Me
    Else
       If MsgBox(" ¿ Desea registrar Nuevo Documento de Proveedor ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
          txtMovNro = oContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
          VerReciboEgreso
          fgDetalle.Clear
          fgDetalle.FormaCabecera
          fgDetalle.Rows = 2
          
          fgObj.Clear
          fgObj.FormaCabecera
          fgObj.Rows = 2
          
          If cmdAjuste.Caption = "Desasignar A&juste   >>>" Then
             cmdAjuste_Click
          End If
          txtMovDesc.Text = ""
          txtBuscarProv.Text = ""
          lblProvNombre.Caption = ""
          cboDoc.ListIndex = -1
          CalculaTotal
          txtBuscarProv.SetFocus
          lbNewProv = False
          Me.txtDocFecha.Mask = ""
          Me.txtDocFecha.Text = ""
          Me.txtDocFecha.Mask = "##/##/####"
          
          Me.ChkVisitaAge.value = 0
          
          OK = False
       Else
           Unload Me
       End If
    End If
Exit Sub
ErrAceptar:
   MsgBox TextErr(Err.Description), vbInformation, "Aviso en Actualización"
End Sub

Private Function CamposOk() As Boolean
CamposOk = False
If frameCaja.Visible And txtBuscarAreaCH.Text = "" Then
   MsgBox "No se definiò Caja Chica a solicitar", vbInformation, "Aviso"
   txtBuscarAreaCH.SetFocus
   Exit Function
End If

If Trim(txtBuscarProv) = "" Or Trim(lblProvNombre) = "" Then
   MsgBox "Falta especificar datos de Proveedor"
   txtBuscarProv.SetFocus
   Exit Function
End If
'PASI20140813
If gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401141" Or gsOpeCod = "402141" Or gsOpeCod = "401371" Or gsOpeCod = "401344" Or gsOpeCod = "401241" Or gsOpeCod = "402241" Then
    If rsProv!TpoDoc = 1 Then
        Select Case Val(Left(Trim(Me.cboDoc.Text), 3))
               Case 1, 2, 3, 5, 10, 12, 14, 16, 28
                    MsgBox "Este Proveedor no tiene RUC, por favor comuníquese con Contabilidad para actualizar sus datos.", vbOKOnly + vbInformation, "Aviso"
                    txtBuscarProv = ""
                    lblProvNombre = ""
                Exit Function
        End Select
    End If
End If
'end PASI
If txtMovDesc = "" Then
   MsgBox "Falta especificar Motivo de Provisión", vbInformation, "Aviso"
   txtMovDesc.SetFocus
   Exit Function
End If
If cboDoc.ListIndex < 0 Then
   MsgBox "No se especifico Tipo de Comprobante de Provisión", vbInformation, "Aviso"
   cboDoc.SetFocus
   Exit Function
End If
If txtDocSerie = "" Then
   MsgBox "Falta especificar Serie de Comprobante", vbInformation, "Aviso"
   txtDocSerie.SetFocus
   Exit Function
End If
If txtDocNro = "" Then
   MsgBox "Falta especificar Número de Comprobante", vbInformation, "Aviso"
   txtDocNro.SetFocus
   Exit Function
End If
If ValidaFechaContab(txtDocFecha, gdFecSis, False) = False Then
   txtDocFecha.SetFocus
   Exit Function
End If
If Val(Format(txtTotal, gsFormatoNumeroDato)) = 0 Or Val(Format(txtSTotal, gsFormatoNumeroDato)) = 0 Then
   MsgBox "Falta indicar Importe de Documento", vbInformation, "Aviso"
   fgDetalle.SetFocus
   Exit Function
End If
If Me.fgDetalle.TextMatrix(1, 1) = "" Then
    MsgBox "No se definio concepto de Gasto", vbInformation, "¡Aviso!"
    fgDetalle.SetFocus
    Exit Function
End If

If Val(Left(Me.cboDoc, 3)) = TpoDocFactura Then
   If Left(txtBuscarProv, 4) = String(4, "0") Then
     MsgBox "Para registrar Factura el Proveedor debe tener RUC", vbInformation, "¡Aviso!"
     
     Exit Function
   End If
End If
'***Agregado por ELRO el 20120927, según OYP-RFC111-2012
If cbogastos.Visible = True And cbogastos.Enabled = True Then
    If cbogastos.ListIndex = -1 Then
        MsgBox "Debe elegir el motivo del gasto.", vbInformation, "¡Aviso!"
        cbogastos.SetFocus
        Exit Function
    End If
End If
'***Fin Agregado por ELRO el 20120927*******************
'***Agregado por ELRO el 20130111, según OYP-RFC126-2012
If gsOpeCod = gCHArendirCtaSustMN Or gsOpeCod = gCHArendirCtaSustME Then
   If Left(Trim(cboDoc.Text), 2) = "97" And txtLugar.Text = "" Then
     MsgBox "Debe ingresar el lugar de desplazamiento", vbInformation, "¡Aviso!"
     txtLugar.SetFocus
     Exit Function
   End If
End If
'***Fin Agregado por ELRO el 20130111*******************
CamposOk = True
End Function

Private Sub cmdAgregar_Click()
Dim lnFila As Integer
If Me.fgDetalle.TextMatrix(1, 0) = "" Then
    fgDetalle.AdicionaFila , Val(fgDetalle.TextMatrix(fgDetalle.row, 0)) + 1
    lnFila = fgDetalle.row
    fgDetalle.TextMatrix(lnFila, 6) = fgDetalle.TextMatrix(lnFila, 0)
    fgDetalle.SetFocus
    SendKeys "{ENTER}"
Else
    If Val(Format(fgDetalle.TextMatrix(fgDetalle.Rows - 1, 3), gsFormatoNumeroDato)) <> 0 And _
        Len(Format(fgDetalle.TextMatrix(fgDetalle.Rows - 1, 1), gsFormatoNumeroDato)) > 0 Then
        fgDetalle.AdicionaFila , fgDetalle.TextMatrix(fgDetalle.Rows - 1, 0) + 1
        lnFila = fgDetalle.row
        fgDetalle.TextMatrix(lnFila, 6) = fgDetalle.TextMatrix(lnFila, 0)
    Else
        If fgDetalle.Enabled Then
           fgDetalle.SetFocus
        End If
    End If
    fgDetalle.SetFocus
    SendKeys "{ENTER}"
End If
End Sub

Private Sub cmdAjuste_Click()
If Val(txtTotal) = 0 Then
   MsgBox "Primero ingresar Concepto de Gastos de Documento", vbInformation, "Aviso"
   Exit Sub
End If
If cmdAjuste.Caption = "A&juste Manual   >>>" Then
   ActivaAjusteBoton False
   txtAjuste.SetFocus
Else
    fgImp.EliminaFila fgImp.Rows - 1
    cmdAjuste.Caption = "A&juste Manual   >>>"
    fgDetalle.Cols = fgDetalle.Cols - 1
    fgDetalle.SetFocus
End If
End Sub

Private Sub ActivaAjusteBoton(plActiva As Boolean)
   fraAjuste.Visible = Not plActiva
   fgDetalle.Enabled = plActiva
   fgObj.Enabled = plActiva
   fgImp.Enabled = plActiva
   cmdAgregar.Visible = plActiva
   cmdEliminar.Visible = plActiva
   cmdAjuste.Visible = plActiva
   cmdValVenta.Visible = plActiva
   fraDoc.Enabled = plActiva
End Sub

Private Sub cmdAplicar_Click()
Dim nRow As Integer
Dim nTot As Currency
Dim nImp As Currency
fgImp.AdicionaFila
nRow = fgImp.row
fgImp.col = 2
fgImp.TextMatrix(nRow, 2) = "1"
'Set fgImp.CellPicture = picCuadroSi.Picture
'fgImp.Col = 1
'Falta adicionar Check
fgImp.TextMatrix(nRow, 0) = nRow
fgImp.TextMatrix(nRow, 3) = "AJUSTE"
fgImp.TextMatrix(nRow, 4) = Format(Round(nVal(txtAjuste) * 100 / nVal(txtSTotal), 2), gsFormatoNumeroView)
fgImp.TextMatrix(nRow, 5) = Format(txtAjuste, "#,#0.00")
fgImp.TextMatrix(nRow, 6) = "AJUSTE"
fgImp.TextMatrix(nRow, 7) = "AJUSTE"
fgImp.TextMatrix(nRow, 8) = "D"
fgImp.TextMatrix(nRow, 9) = "0"
fgImp.TextMatrix(nRow, 10) = "2"
'Distribución del Ajuste entre Cuentas de Gasto
fgDetalle.Cols = fgDetalle.Cols + 1
fgDetalle.TextMatrix(0, fgDetalle.Cols - 1) = "AJUSTE"
nTot = 0
For nRow = 1 To fgDetalle.Rows - 1
   nImp = Round(Val(txtAjuste) * Val(fgDetalle.TextMatrix(nRow, 3)) / Val(txtSTotal), 2)
   nTot = nTot + nImp
   fgDetalle.TextMatrix(nRow, fgDetalle.Cols - 1) = Format(nImp, gsFormatoNumeroView)
Next
If nTot <> Val(txtAjuste) Then
   fgDetalle.TextMatrix(1, fgDetalle.Cols - 1) = Val(fgDetalle.TextMatrix(1, fgDetalle.Cols - 1)) + (Val(txtAjuste) - nTot)
End If
cmdAjuste.Caption = "Desasignar A&juste   >>>"
CalculaTotal False
ActivaAjusteBoton True

End Sub

Private Sub cmdCancelar_Click()
cmdAjuste.Caption = "A&juste Manual   >>>"
ActivaAjusteBoton True
End Sub

Private Sub cmdEliminar_Click()
If fgDetalle.TextMatrix(fgDetalle.row, 0) <> "" Then
   EliminaCuenta fgDetalle.TextMatrix(fgDetalle.row, 1), fgDetalle.TextMatrix(fgDetalle.row, 0)
   CalculaTotal
   If fgDetalle.Enabled Then
      fgDetalle.SetFocus
   End If
End If
End Sub
Private Sub EliminaCuenta(sCod As String, nItem As Integer)
If fgDetalle.TextMatrix(1, 0) <> "" Then
    EliminaFgObj Val(fgDetalle.TextMatrix(fgDetalle.row, 0))
    fgDetalle.EliminaFila fgDetalle.row, False
End If
If Len(fgDetalle.TextMatrix(1, 1)) > 0 Then
   RefrescaFgObj Val(fgDetalle.TextMatrix(fgDetalle.row, 0))
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
Dim lbUltimaInstancia As Boolean
Dim oCon As New DConecta
oCon.AbreConexion

Set oEfect = New Defectivo
Set oCtaIf = New NCajaCtaIF
Set oRHAreas = New DActualizaDatosArea
Set oDescObj = New ClassDescObjeto
Set oCtaCont = New DCtaCont
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
lbUltimaInstancia = False
EliminaFgObj Val(fgDetalle.TextMatrix(fgDetalle.row, 0))
Set rs1 = oCtaCont.CargaCtaObj(psCtaContCod, , True)
If Not rs1.EOF And Not rs1.BOF Then
    Do While Not rs1.EOF
        lsRaiz = ""
        lsFiltro = ""
        Select Case Val(rs1!cObjetoCod)
            Case ObjCMACAgencias
                If lsAgeCod = "" Then
                    Set rs = oRHAreas.GetAgencias(rs1!cCtaObjFiltro)
                Else
                    Set rs = Nothing
                End If
                lbUltimaInstancia = True
            Case ObjCMACAgenciaArea
                If lsAreaCod = "" And lsAgeCod = "" Then
                    lsRaiz = "Unidades Organizacionales"
                    Set rs = oRHAreas.GetAgenciasAreas(rs1!cCtaObjFiltro)
                    
                '********Modificado por PASI 20131225 TI-ERS107-2013
'                '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
'                'ElseIf gsOpeCod = "401140" Or gsOpeCod = "402140" Then
'                ElseIf gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401141" Or gsOpeCod = "402141" Then
'                '***Fin Modificado por ELRO el 20120921*******************
                ElseIf gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401141" Or gsOpeCod = "402141" Or _
                     gsOpeCod = "401240" Or gsOpeCod = "402240" Or gsOpeCod = "401241" Or gsOpeCod = "402241" Or gsOpeCod = "401344" Then  'VAPI-ERS1972014 AGREGO Or gsOpeCod = "401344"
                '***Fin Modificado por PASI 20131225*******************
                
                    lsRaiz = "Unidades Organizacionales"
                    Set rs = oRHAreas.GetAgenciasAreas(rs1!cCtaObjFiltro)
                Else
                    Set rs = Nothing
                End If
                lbUltimaInstancia = True
            Case ObjCMACArea
                If lsAreaCod = "" Then
                    Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
                Else
                    Set rs = Nothing
                End If
            Case ObjEntidadesFinancieras
                lsRaiz = "Cuentas de Entidades Financieras"
                Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, psCtaContCod)
                lbUltimaInstancia = True
            Case ObjDescomEfectivo
                Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
                lbUltimaInstancia = True
            Case ObjPersona
                Set rs = Nothing
            Case Else
                Set rs = oCon.CargaRecordSet("spGetTreeObj '" & rs1!cObjetoCod & "'," & rs1!nCtaObjNiv + rs1!nObjetoNiv & ",'" & rs1!cCtaObjFiltro & "'")
                lbUltimaInstancia = True
        End Select
        If Not rs Is Nothing Then
            If rs.State = adStateOpen Then
                If Not rs.EOF And Not rs.BOF Then
                    If rs.RecordCount > 1 Then
                        oDescObj.lbUltNivel = lbUltimaInstancia
                        oDescObj.Show rs, "", lsRaiz
                    
                        If oDescObj.lbOk Then
                            lsFiltro = oContFunct.GetFiltroObjetos(Val(rs1!cObjetoCod), psCtaContCod, oDescObj.gsSelecCod, False)
                            AdicionaObj psCtaContCod, fgDetalle.TextMatrix(fgDetalle.row, 0), rs1!nCtaObjOrden, oDescObj.gsSelecCod, _
                                        oDescObj.gsSelecDesc, lsFiltro, rs1!cObjetoCod
                        Else
                            fgDetalle.EliminaFila fgDetalle.row, False
                            Exit Do
                        End If
                    Else
                        AdicionaObj psCtaContCod, fgDetalle.TextMatrix(fgDetalle.row, 0), rs1!nCtaObjOrden, rs1!cObjetoCod, _
                                        rs1!cObjetoDesc, lsFiltro, rs1!cObjetoCod
                    End If
                End If
            End If
            RSClose rs
        Else
            If Val(rs1!cObjetoCod) = ObjPersona Then
                Set UP = frmBuscaPersona.Inicio
                If Not UP Is Nothing Then
                    AdicionaObj psCtaContCod, fgDetalle.TextMatrix(fgDetalle.row, 0), rs1!nCtaObjOrden, _
                                    UP.sPersCod, UP.sPersNombre, _
                                    lsFiltro, rs1!cObjetoCod
                End If
            End If
            If Val(rs1!cObjetoCod) = ObjCMACAgenciaArea Or Val(rs1!cObjetoCod) = ObjCMACAgencias Or Val(rs1!cObjetoCod) = ObjCMACArea Then
                lsFiltro = oContFunct.GetFiltroObjetos(Val(rs1!cObjetoCod), psCtaContCod, lsAreaCod & lsAgeCod, False)
                AdicionaObj psCtaContCod, fgDetalle.TextMatrix(fgDetalle.row, 0), rs1!nCtaObjOrden, lsAreaCod & lsAgeCod, lsAreaDes, lsFiltro, rs1!cObjetoCod
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

oCon.CierraConexion
Set oCon = Nothing
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
    'fgDetalle.TextMatrix(fgDetalle.Row, 6) = psObjetoCod
    
End Sub


Private Sub cmdSalir_Click()
OK = False
Unload Me
End Sub
Private Sub cmdValVenta_Click()
Dim n As Integer
   
If MsgBox(" ¿ Esta Seguro de aplicar Ajuste ? ", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
Else
    For n = 1 To fgDetalle.Rows - 1
        fgDetalle.TextMatrix(n, 3) = Format(Round(Val(Format(fgDetalle.TextMatrix(n, 3), gsFormatoNumeroDato)) / (1 + (nVal(fgImp.TextMatrix(fgImp.row, 4)) / 100)), 2), gsFormatoNumeroView)
    Next
    CalculaTotal
    nVariaIGV = 0
End If
End Sub
'Private Sub fgDetalle_KeyPress(KeyAscii As Integer)
'If fgDetalle.TextMatrix(fgDetalle.Row, 1) = "" Then
'   Exit Sub
'End If
'If fgDetalle.Col = 1 Then
'   If KeyAscii = 13 Then EnfocaTexto txtCta, IIf(KeyAscii = 13, 0, KeyAscii), fgDetalle
'   If KeyAscii = 32 Then
'      If cboDocDestino.ListIndex = 1 Then
'         If fgDetalle.TextMatrix(fgDetalle.Row, 5) = "X" Then
'            mnuGravado.Checked = True
'         Else
 '           mnuGravado.Checked = False
 '        End If
'         mnuGravado_Click
'      End If
'   End If
'End If
'If fgDetalle.Col = 3 Or fgDetalle.Col > 6 Then
'   If fgDetalle.Col > 6 And fgDetalle.Text = "" Then
'      Exit Sub
'   End If
'   If InStr("-0123456789.", Chr(KeyAscii)) > 0 Then
'      EnfocaTexto txtCant, KeyAscii, fgDetalle
'   Else
'      If KeyAscii = 13 Then EnfocaTexto txtCant, 0, fgDetalle
'   End If
'End If
'End Sub
'Private Sub fgDetalle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 2 Then
'   If cboDocDestino.ListIndex = 1 And fgDetalle.Col = 1 Then
'      mnu_.Visible = True
'      mnuGravado.Visible = True
'      If fgDetalle.TextMatrix(fgDetalle.Row, 5) = "X" Then
'         mnuGravado.Checked = True
'      Else
'         mnuGravado.Checked = False
'      End If
'   Else
'      mnu_.Visible = False
'      mnuGravado.Visible = False
'   End If
'   PopupMenu mnuObj
'End If

Private Sub fgDetalle_OnCellChange(pnRow As Long, pnCol As Long)
If fgDetalle.TextMatrix(1, 0) <> "" Then
    CalculaTotal pnCol < 7
End If
End Sub

Private Sub fgDetalle_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
'If psDataCod <> "" Then
'    fgDetalle.TextMatrix(pnRow, 4) = oOpe.GetOpeCtaDebeHaber(gsOpeCod, psDataCod)
'    AsignaCtaObj psDataCod
'Else
'    fgDetalle.EliminaFila pnRow
'    cmdAgregar_Click
'End If
If psDataCod <> "" Then
    fgDetalle.TextMatrix(pnRow, 4) = oOpe.GetOpeCtaDebeHaber(gsOpeCod, psDataCod)
    AsignaCtaObj psDataCod
    '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
    'MADM 20110505
    'If (gsOpeCod = "401140" Or (gsOpeCod = "401371" And Trim(Me.txtBuscarAreaCH.Text) = "042")) Then
    If (gsOpeCod = "401140" Or gsOpeCod = "401141" Or ((gsOpeCod = "401371" Or gsOpeCod = "401344") And Trim(Me.txtBuscarAreaCH.Text) = "042")) Then
    '***Fin Modificado por ELRO el 20120921*******************
            If Mid(fgDetalle.TextMatrix(1, 1), 1, 2) = "15" And Mid(fgDetalle.TextMatrix(1, 1), 4, 5) = "70902" Then
                If Me.cbogastos.Visible = False Then
                    Me.cbogastos.Visible = True
                    Me.ActxCta.Visible = True
                    Me.lblGastos.Visible = True
                End If
            Else
                If Me.cbogastos.Visible Then
                    Me.cbogastos.Visible = False
                    Me.ActxCta.Visible = False
                    Me.lblGastos.Visible = False
                End If
            End If
       End If
    'END MADM
Else
    fgDetalle.EliminaFila pnRow
    cmdAgregar_Click
End If
End Sub

Private Sub fgDetalle_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim nTempMonto As Currency

If pnCol = 3 And fgDetalle.TextMatrix(pnRow, 3) <> "" Then
    CalculaTotal
    If bAlert = 1 Or lbCajaChica = True Then
        nTempMonto = txtTotal  'CDbl(fgDetalle.TextMatrix(pnRow, 3))
        If Me.chkDocDolares.value = 1 Then
            If Round(nTempMonto * CCur(Me.txtTipVariable.Text), 2) > lnTempSaldo Then   'Or (nTempMonto + Val(txtTotal.Text)) > lnTempSaldo Then
                MsgBox "Total mayor a Saldo: " & lnTempSaldo & " el monto que Ud. desea asignar es " & Format(Round(nTempMonto * CCur(Me.txtTipVariable.Text), 2), gsFormatoNumeroView), vbInformation, "Aviso"
                fgDetalle.TextMatrix(pnRow, 3) = ""
                fgDetalle.row = pnRow
                fgDetalle.col = pnCol
                'Exit Sub
            End If
        Else
            If nTempMonto > lnTempSaldo Then  'Or (nTempMonto + Val(txtTotal.Text)) > lnTempSaldo Then
                MsgBox "Total mayor a Saldo: " & lnTempSaldo, vbInformation, "Aviso"
                fgDetalle.TextMatrix(pnRow, 3) = ""
                fgDetalle.row = pnRow
                fgDetalle.col = pnCol
                'Exit Sub
            End If
        End If
    End If
End If



If pnCol = 8 Then
   If fgDetalle.TextMatrix(pnRow, 8) <> "" Then
    CalculaTotal
    If fgImp.TextMatrix(2, 3) = "OTROS" Then
        nTempMonto = Val(txtTotal)
    Else
        nTempMonto = Val(txtTotal) + Val(fgDetalle.TextMatrix(pnRow, 8))
    End If
    
        If nTempMonto > lnTempSaldo Then  'Or (nTempMonto + Val(txtTotal.Text)) > lnTempSaldo Then
           MsgBox "Total mayor a Saldo: " & lnTempSaldo, vbInformation, "Aviso"
           fgDetalle.TextMatrix(pnRow, 8) = ""
           fgDetalle.row = pnRow
           fgDetalle.col = pnCol
        'Exit Sub
        End If
   End If
End If
End Sub

Private Sub fgDetalle_RowColChange()
    If Len(fgDetalle.TextMatrix(1, 1)) > 0 Then
       If fgDetalle.TextMatrix(fgDetalle.row, 0) <> "" Then
            RefrescaFgObj Val(fgDetalle.TextMatrix(fgDetalle.row, 0))
       End If
    End If
'End If
End Sub

Private Sub fgDetalle_Validate(Cancel As Boolean)
If fgDetalle.TextMatrix(1, 0) <> "" Then
'    CalculaTotal
End If

End Sub

Private Sub fgImp_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
If fgDetalle.TextMatrix(1, 0) <> "" Then
      'John ----------------
      If fgImp.TextMatrix(pnRow, 3) = "IGV" And fgImp.TextMatrix(pnRow, 5) = "0.00" Then
      
         Dim oConst As NConstSistemas
         Set oConst = New NConstSistemas

         cboDocDestino.Enabled = True
         Me.cboDocDestino.ListIndex = oConst.LeeConstSistema(gConstSistDestinoIGVDefecto)
      Else
         If fgImp.TextMatrix(pnRow, 3) = "IGV" And fgImp.TextMatrix(pnRow, 1) = "." Then
            cboDocDestino.ListIndex = -1
            cboDocDestino.Enabled = False
         End If
      End If
      '---------------------

    CalculaTotal
End If
End Sub

Private Sub fgImp_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
If fgImp.TextMatrix(fgImp.row, 10) <> "2" Then
    Cancel = False
End If

End Sub

Private Sub Form_Activate()
If lSalir Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim n As Integer, nSaldo As Currency, nCant As Currency
Dim sCtaCod As String
Dim rs As ADODB.Recordset
Dim oConst As New NConstSistemas

Set rs = New ADODB.Recordset
Set oContFunct = New NContFunciones
Set oNArendir = New NARendir
Set oOpe = New DOperacion
Set oCtaCont = New DCtaCont
CentraForm Me
Me.Caption = gsOpeDesc
lSalir = False
OK = False

'EJVG20140727 ***
fsCtaContAporteAFP = oConst.LeeConstSistema(483)
fsCtaContSeguroAFP = oConst.LeeConstSistema(484)
fsCtaContComisionAFP = oConst.LeeConstSistema(485)
fsCtaContAporteONP = oConst.LeeConstSistema(486)
Set oConst = Nothing
'END EJVG *******

CambiaTamañoCombo cboDoc, 300
CambiaTamañoCombo cboDocDestino, 250

'Defino el Nro de Movimiento
txtOpeCod = gsOpeCod
txtMovNro = oContFunct.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
txtFecha = Format(gdFecSis, gsFormatoFechaView)
lblFechaArendir = lsFechaARendir
fgDetalle.TipoBusqueda = BuscaArbol
fgDetalle.psRaiz = "Cuentas Contables"
fgDetalle.rsTextBuscar = oOpe.EmiteOpeCtasNivel(gsOpeCod, "D")

If Mid(gsOpeCod, 3, 1) = gMonedaExtranjera Then    'Identificación de Tipo de Moneda
   lMN = False
   sMoney = gcME
   gsSimbolo = gcME
   If gnTipCambio = 0 Then
      If Not GetTipCambio(gdFecSis) Then
         lSalir = True
         Exit Sub
      End If
   End If
   FrameTipCambio.Visible = True
   txtTipFijo = gnTipCambio
   
    If gbBitTCPonderado Then
         Label9.Caption = "Ponder."
         txtTipVariable = Format(gnTipCambioPonderadoVenta, gsFormatoNumeroView3Dec)
         txtTipPondVenta = Format(gnTipCambioPonderadoVenta, gsFormatoNumeroView3Dec)
    Else
         txtTipVariable = Format(gnTipCambioV, gsFormatoNumeroView3Dec)
         txtTipPondVenta = Format(gnTipCambioV, gsFormatoNumeroView3Dec)
    End If
Else
   lMN = True
   sMoney = gcMN
   gsSimbolo = gcMN
End If

'*** PEAC 20100708

Set rs = oOpe.CargaAreaSoli()
Do While Not rs.EOF 'cAreaCod,cAreaDescripcion
   cboAreaSoli.AddItem rs!cAreaCod & "   " & Mid(rs!cAreaDescripcion & space(100), 1, 100)
   rs.MoveNext
Loop
'cboAreaSoli.ListIndex = cboAreaSoli.ListCount - 1

'*** FIN PEAC

'MADM 20110505
ActxCta.NroCuenta = ""
'***Modificado por ELRO el 20120921, según OYP-RFC111-2012
'If (gsOpeCod = "401140" Or gsOpeCod = "401371") Then
If (gsOpeCod = "401140" Or gsOpeCod = "401141" Or gsOpeCod = "401371" Or gsOpeCod = "401344") Then
'***Fin Modificado por ELRO el 20120921*******************
    ActxCta.CMAC = gsCodCMAC
    Set rs = oOpe.DevuelveGastosCreditos()
    Do While Not rs.EOF 'cAreaCod,cAreaDescripcion
       cbogastos.AddItem rs!nPrdConceptoCod & "   " & Mid(rs!cDescripcion & space(100), 1, 100)
       rs.MoveNext
    Loop
End If
ActxCta.Visible = False
cbogastos.Visible = False
lblGastos.Visible = False
'END MADM

Set rs = oOpe.CargaOpeDoc(gsOpeCod, , OpeDocMetDigitado)
Do While Not rs.EOF
   cboDoc.AddItem rs!nDocTpo & "   " & Mid(rs!cDocDesc & space(100), 1, 100) & Mid(rs!cDocAbrev & "   ", 1, 3)
   rs.MoveNext
Loop

Set rs = oOpe.CargaOpeCta(gsOpeCod, "H", "0")
If Not rs.EOF And Not rs.EOF Then
    Do While Not rs.EOF
        cboProvis.AddItem rs!cCtaContDesc & space(100) & rs!cCtaContCod
        rs.MoveNext
    Loop
    cboProvis.ListIndex = cboProvis.ListCount - 1
Else
    MsgBox "No se definieron Cuentas de Provisión para Operación. Por favor Consultar con Sistemas", vbInformation, "Aviso"
    lSalir = True
    Exit Sub
End If
If lnArendirFase = ArendirSustentacion Or lnArendirFase = ArendirRendicion Then 'Para a rendir de Caja General
   fraProvis.Visible = False
   If Mid(gsOpeCod, 3, 1) = Moneda.gMonedaNacional Then
      Me.chkDocDolares.Visible = True
      Me.chkDocDolares.value = 0
   End If
   lsCtaCodPend = oOpe.EmiteOpeCta(gsOpeCod, "H", "1")
   If lsCtaCodPend = "" Then
      MsgBox "No se definió Cuenta de Pendiente para Arendir", vbInformation, "Aviso"
      lSalir = True
      Exit Sub
   End If
End If
If lbCajaChica Then   'Cuando es Caja Chica
    txtFecha.Enabled = False
    fraArendir.Visible = False
    frameCaja.Visible = True
    fraProvis.Visible = False
    txtBuscarAreaCH.psRaiz = "Cajas Chicas"
    txtBuscarAreaCH.rs = oNArendir.EmiteCajasChicas

Else
    fraArendir.Visible = True
    lblArendirNro = lsNroArendir
    lblTotalArendir = Format(lnTotalArendir, "#,#0.00")
End If
'Modificado PASI20140613 TI-ERS061-2014
'txtFecha.Enabled = lbCajaChica
    txtFecha.Enabled = True
'end PASI
lbNewProv = False

'*** PEAC 20110303
'401371
Me.lblAreaSoli.Visible = False
Me.cboAreaSoli.Visible = False

'***Modificado por ELRO el 20120921, según OYP-RFC111-2012
'If (gsOpeCod = "401240" Or gsOpeCod = "402240" Or gsOpeCod = "401140" Or gsOpeCod = "402140") And (Right(Trim(lsAgeCod), 2) = "01" Or Right(Trim(lsAgeCod), 2) = "") Then
 'If (gsOpeCod = "401240" Or gsOpeCod = "402240" Or gsOpeCod = "401241" Or gsOpeCod = "402241" Or gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401141" Or gsOpeCod = "402141") And (Right(Trim(lsAgeCod), 2) = "01" Or Right(Trim(lsAgeCod), 2) = "") Then
 If (Right(Trim(lsAgeCod), 2) = "01" Or Right(Trim(lsAgeCod), 2) = "") Then
'***Fin Modificado por ELRO el 20120921*******************
    If lsAreaCod = "026" Or lsAreaCod = "041" Or lsAreaCod = "067" Then
        Me.ChkVisitaAge.Visible = False
    Else
        '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
        If gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401141" Or gsOpeCod = "402141" Then ChkVisitaAge.Caption = "Distribución Manual de Gastos"
        '***Fin Modificado por ELRO el 20120921*******************
        Me.ChkVisitaAge.Visible = True
    End If
Else
    Me.ChkVisitaAge.Visible = False
End If
'*** FIN PEAC
    '***Agrgegado por ELRO el 20120924, según OYP-RFC111-2012
If lbCajaChica = True And gsOpeCod = "401344" Then
    txtBuscarAreaCH.psDescripcion = lsAreaDes
    txtBuscarAreaCH = lsAreaCh & lsAgeCh
    txtBuscarAreaCH_EmiteDatos
    frameCaja.Enabled = False
End If
    '***Fin Agrgegado por ELRO el 20120924*******************
fraDoc.Enabled = True 'PASI20140613 TI-ERS061-2014
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not OK And lSalir = False Then
   If MsgBox(" ¿ Seguro de salir sin grabar Operación ? ", vbQuestion + vbYesNo) = vbNo Then
      Cancel = 1
      Exit Sub
   End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oContFunct = Nothing
Set oNArendir = Nothing
End Sub



Private Sub txtAjuste_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtAjuste, KeyAscii, 16, 2)
If KeyAscii = 13 Then
   txtAjuste = Format(txtAjuste, gsFormatoNumeroView)
   cmdAplicar.SetFocus
End If
End Sub

Private Sub txtAjuste_Validate(Cancel As Boolean)
txtAjuste = Format(txtAjuste, gsFormatoNumeroView)
End Sub
Private Sub txtBuscarAreaCH_EmiteDatos()
Dim oCajaCH As nCajaChica
Dim lnSaldo As Currency
Set oCajaCH = New nCajaChica
If txtBuscarAreaCH.Text = "" Then Exit Sub
lblCajaChicaDesc = txtBuscarAreaCH.psDescripcion
lblNroProc = oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), NroCajaChica)
lnSaldo = oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), SaldoActual)

'***Modificado por ELRO el 20120921, según OYP-RFC111-2012
'MADM 20110511
'If (gsOpeCod = "401371" And Trim(Me.txtBuscarAreaCH.Text) = "042") And (Mid(fgDetalle.TextMatrix(1, 1), 1, 2) = "15" And Mid(fgDetalle.TextMatrix(1, 1), 4, 5) = "70902") Then
If ((gsOpeCod = "401371" Or gsOpeCod = "401344") And Trim(Me.txtBuscarAreaCH.Text) = "042") And (Mid(fgDetalle.TextMatrix(1, 1), 1, 2) = "15" And Mid(fgDetalle.TextMatrix(1, 1), 4, 5) = "70902") Then
'***Fin Modificado por ELRO el 20120921*******************
     If Me.cbogastos.Visible = False Then
        Me.cbogastos.Visible = True
        Me.ActxCta.Visible = True
        Me.lblGastos.Visible = True
      End If
Else
        Me.cbogastos.Visible = False
        Me.ActxCta.Visible = False
        Me.lblGastos.Visible = False
End If

If lnSaldo = 0 Then
    MsgBox "Caja Chica no posee Saldo", vbInformation, "Aviso"
    lblCajaChicaDesc = ""
    txtBuscarAreaCH = ""
    lblNroProc = ""
    Exit Sub
End If
If oCajaCH.VerificaTopeCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2)) Then
    If MsgBox("Caja chica ha sobrepasado el limite establecido. debe realizar la Rendición Respectiva " & vbCrLf & "Desea Continuar ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        lblCajaChicaDesc = ""
        txtBuscarAreaCH = ""
        lblNroProc = ""
        Exit Sub
    End If
End If
'***Comentado por ELRO el 20120927, según OYP-RFC111-2012
'If lblCajaChicaDesc <> "" Then
'    If txtfecha.Enabled Then
'        txtfecha.SetFocus
'    End If
'End If
'***Fin Comentado por ELRO el 20120927*******************
'*** PEAC

'Me.cboAreaSoli.Enabled = False
Me.cboAreaSoli.Visible = False

If Left(Me.txtBuscarAreaCH, 3) = "067" Or Left(Me.txtBuscarAreaCH, 3) = "041" Or Left(Me.txtBuscarAreaCH, 3) = "026" Then
    Me.ChkVisitaAge.Visible = False
Else
    '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
    If gsOpeCod = "401371" Or gsOpeCod = "402371" Or gsOpeCod = "401344" Then ChkVisitaAge.Caption = "Distribución Manual de Gastos"
    '***Fin Modificado por ELRO el 20120921*******************
    Me.ChkVisitaAge.Visible = True
End If

'If Left(Me.txtBuscarAreaCH, 3) <> "067" Then
'    Me.cboAreaSoli.Enabled = True
'Else
'    If Right(Me.txtBuscarAreaCH, 2) = "01" Then
'        Me.cboAreaSoli.Enabled = True
'    Else
'        Me.cboAreaSoli.Enabled = False
'    End If
'End If
'*** FIN PEAC

Set oCajaCH = Nothing
'***Modificado por ELRO el 20120924, según OYP-RFC111-2012
'lnTempSaldo = lnSaldo

If gsOpeCod <> "401344" Then
'lnTempSaldo
'    lnTempSaldo = lnTotalArendir
'Else
    lnTempSaldo = lnSaldo
End If
'***Fin Modificado por ELRO el 20120924*******************
End Sub
Private Sub txtBuscarAreaCH_Validate(Cancel As Boolean)
If txtBuscarAreaCH = "" Then Cancel = True

End Sub

Private Sub txtBuscarProv_EmiteDatos()
Dim lsMotivoNoHabido As String
Dim oProv As DLogProveedor
Set oProv = New DLogProveedor
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
Dim codigopersona As String
Dim lnRuc As String
lbNewProv = False

'txtBuscarProv.Text = txtBuscarProv.psCodigoPersona

'*** PEAC 20081023 - PERMITIR QUE LA CAJA DE TEXTO MUESTRE CEROS CUANDO NO TIENE NI DNI NI RUC (EXTRANJEROS)
'   txtBuscarProv.Text = txtBuscarProv.sPersNroDoc
'*** END PEAC

lblProvNombre.Caption = txtBuscarProv.psDescripcion

codigopersona = txtBuscarProv.psCodigoPersona

    'Set rs1 = oProv.GetProveedorRUC(txtBuscarProv.psCodigoPersona)
    'Set rs1 = oProv.GetProveedorRUCoDNI(codigopersona) 'Comentado PASI20140813
    Set rsProv = New ADODB.Recordset 'PASI20140813
    Set rsProv = oProv.GetProveedorRUCoDNI(codigopersona)
    If Not (rsProv.EOF And rsProv.BOF) Then 'Se cambio rs1 por rsProv
        lnRuc = rsProv!nroDoc
        txtBuscarProv.Text = lnRuc
        
        'PASI TI-ERS061-2014
'        If Trim(Me.cboDoc.Text) = "" And Trim(lblProvNombre.Caption) <> "" Then 'Comentado PASI20140813
'                MsgBox "Por favor seleccione un tipo de Documento a registrar.", vbOKOnly + vbExclamation, "Atención"
'                txtBuscarProv = ""
'                lblProvNombre = ""
'                Me.cboDoc.Enabled = True '*** PEAC 20110426
'                Exit Sub
'        End If
'
'        If gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401141" Or gsOpeCod = "402141" Or gsOpeCod = "401371" Or gsOpeCod = "401344" Then 'Comentado PASI20140813
'            If rs1!TpoDoc = 1 Then 'PASI 20140718
'                Select Case Val(Left(Trim(Me.cboDoc.Text), 3))
'                       Case 1, 2, 3, 5, 10, 12, 14, 16, 28
'                            MsgBox "Este Proveedor no tiene RUC, por favor comuníquese con Contabilidad para actualizar sus datos.", vbOKOnly + vbInformation, "Aviso"
'                            txtBuscarProv = ""
'                            lblProvNombre = ""
'                        Exit Sub
'                End Select
'            End If 'End PASI
'        End If
'        'end PASI
        If rsProv!TpoDoc = 2 Then 'Sólo la Condicion agregado por PASI TI-ERS061-2014
            If oProv.GetProveedorNoHabido(lnRuc, lsMotivoNoHabido) Then
                MsgBox "Proveedor fue identificado como no Habido por la Sunat" & Chr(10) _
                    & "Con la siguiente Observación : " & lsMotivoNoHabido, vbInformation, "Aviso"
                txtBuscarProv = ""
                lblProvNombre = ""
                Exit Sub
            End If
        End If
'*** PEAC 20110222
    Else
        '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
        'If gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401371" Then
        If gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401141" Or gsOpeCod = "402141" Or gsOpeCod = "401371" Or gsOpeCod = "401344" Then
        '***Fin Modificado por ELRO el 20120921*******************
            
            If Trim(Me.cboDoc.Text) = "" And Trim(lblProvNombre.Caption) <> "" Then
                MsgBox "Por favor seleccione un tipo de Documento a registrar.", vbOKOnly + vbExclamation, "Atención"
                txtBuscarProv = ""
                lblProvNombre = ""
                Me.cboDoc.Enabled = True '*** PEAC 20110426
                Exit Sub
            Else
                Select Case Val(Left(Trim(Me.cboDoc.Text), 3))
                    Case 1, 2, 3, 5, 10, 12, 14, 16, 28
                        MsgBox "Este Proveedor no tiene RUC, por favor comuníquese con Contabilidad para actualizar sus datos.", vbOKOnly + vbInformation, "Aviso"
                        txtBuscarProv = ""
                        lblProvNombre = ""
                        Exit Sub
                End Select
            End If
        End If
        '*** FIN PEAC
    End If
If lblProvNombre.Caption <> "" Then
    lbNewProv = Not oProv.IsExisProveedor(txtBuscarProv.psCodigoPersona)
    txtMovDesc.SetFocus
End If
'EJVG20140728 ***
If txtBuscarProv.psCodigoPersona <> "" Then
    If gbBitRetencSistPensProv Then
        If Val(Trim(Left(cboDoc.Text, 3))) = TpoDoc.TpoDocRecHonorarios Then
            Call CalculaTotal
        End If
    End If
End If
'END EJVG *******
Set oProv = Nothing
End Sub

Private Sub txtBuscarProv_Validate(Cancel As Boolean)
If (txtBuscarProv = "" And txtBuscarProv.psDescripcion = "") Then
   Cancel = True
End If
End Sub
Private Sub txtDocFecha_KeyPress(KeyAscii As Integer)
Dim nTipFijo As Currency
Dim i As Integer
If KeyAscii = 13 Then
  ' If IsDate(txtDocFecha.Text) Then
   If Not ValidaFechaContab(txtDocFecha, gdFecSis, False) Then
      fEnfoque txtDocFecha
      Exit Sub
   Else
      If (lnArendirFase = ArendirRendicion Or lnArendirFase = ArendirSustentacion) Or lnTipoArendir = gArendirTipoViaticos Then
'         If CDate(txtDocFecha) < lsFechaARendir Then
'            MsgBox "Fecha no puede ser menor a fecha de A rendir", vbInformation, "Aviso"
'            Exit Sub
'         End If

          'JEOM
          If lnTipoArendir = gArendirTipoViaticos Then
             If CDate(txtDocFecha) < lsFechaARendir Then
                MsgBox "Fecha no puede ser menor a fecha de A Rendir", vbInformation, "Aviso"
                Exit Sub
             End If
          End If
          'FIN JEOM


      End If
      If CDate(txtDocFecha) - 3 > gdFecSis Then
         MsgBox "Fecha no puede ser mayor a fecha Actual", vbInformation, "Aviso"
         Exit Sub
      End If
      If Not lMN Or Me.chkDocDolares.value = 1 Then
         'nTipFijo = gnTipCambio
         'GetTipCambio CDate(txtDocFecha)
         '
        'If gbBitTCPonderado Then
        '     txtTipVariable = Format(gnTipCambioPonderadoVenta, gsFormatoNumeroView3Dec)
        '     txtTipPondVenta = Format(gnTipCambioPonderadoVenta, gsFormatoNumeroView3Dec)
        'Else
        '     txtTipVariable = Format(gnTipCambioV, gsFormatoNumeroView3Dec)
        '     txtTipPondVenta = Format(gnTipCambioV, gsFormatoNumeroView3Dec)
        'End If
        'gnTipCambio = nTipFijo
        '
        'GetTipCambio gdFecSis
      End If
      If fgImp.TextMatrix(1, 0) <> "" Then
        Dim oImp As DImpuesto
        Set oImp = New DImpuesto
        For i = 1 To fgImp.Rows - 1
           fgImp.TextMatrix(i, 4) = Format(oImp.CargaImpuestoFechaValor(fgImp.TextMatrix(i, 6), txtDocFecha), gsFormatoNumeroView)
        Next
        Set oImp = Nothing
      End If
      If cboDocDestino.Enabled Then
         cboDocDestino.SetFocus
      Else
         cmdAgregar.SetFocus
      End If
   End If
End If
End Sub

Private Sub txtDocFecha_LostFocus()
If lnTipoArendir = gArendirTipoViaticos Then
    '***Modificado por ELRO el 20120809, según OYP-RFC015-2012, OYP-RFC016-2012 y OYP-RFC047-2012
    'If CDate(txtDocFecha) < lsFechaARendir Then
    If Not ValidaFechaContab(txtDocFecha, gdFecSis, False) Then
       fEnfoque txtDocFecha
       Exit Sub
    Else
        If CDate(txtDocFecha) < lsFechaARendir Then
           MsgBox "Fecha no puede ser menor a fecha de A Rendir", vbInformation, "Aviso"
           txtDocFecha.SetFocus
           Exit Sub
        End If
    End If
    '***Modificado por ELRO**********************************************************************
End If
'EJVG20140727 ***
If IsDate(txtDocFecha) Then
    If gbBitRetencSistPensProv Then
        If Val(Trim(Left(cboDoc.Text, 3))) = TpoDoc.TpoDocRecHonorarios Then
            Call CalculaTotal
        End If
    End If
End If
'END EJVG *******
End Sub

Private Sub txtDocNro_GotFocus()
fEnfoque txtDocNro
End Sub

Private Sub txtDocNro_Validate(Cancel As Boolean)

'*** PEAC 20121122
'If (gsOpeCod = "401240" Or gsOpeCod = "402240" Or gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401344") And Trim(Left(cboDoc, 3)) = "93" Then 'si es doc. Declarac. jurada
If Trim(Left(cboDoc, 3)) = "93" Then  'si es doc. Declarac. jurada
    lsCodigoPersona = ""
Else
    lsCodigoPersona = txtBuscarProv.psCodigoPersona
End If
'*** FIN PEAC

Dim oDoc As DDocumento
Set oDoc = New DDocumento
If oDoc.GetValidaDocProv(lsCodigoPersona, Mid(cboDoc.Text, 1, 2), txtDocSerie & "-" & txtDocNro) Then
    '*** PEAC 20121122
    'If (gsOpeCod = "401240" Or gsOpeCod = "402240" Or gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401344") And Trim(Left(cboDoc, 3)) = "93" Then
    If Trim(Left(cboDoc, 3)) = "93" Then
        '***Modificado por ELRO el 20130603, según SATI INC1304290006****
        'txtDocNro = Right("00000000000" + Trim(Str(Val(txtDocNro) + 1)), 11)
        txtDocNro = Trim(Str(Val(txtDocNro) + 1))
        '***FIn Modificado por ELRO el 20130603, según SATI INC1304290006
        MsgBox "El número de DDJJ acaba de ser registrado por otro usuario, por lo que se incrementó el número automaticamente, por favor vuelva a grabar.", vbInformation, "Aviso"
    Else
        MsgBox "Documento ya se encuentra registrado", vbInformation, "Aviso"
    End If
    'MsgBox "Documento ya se encuentra registrado", vbInformation, "Aviso"
    '*** FIN PEAC
    
    Cancel = True
End If
Set oDoc = Nothing
End Sub


Private Sub txtDocSerie_LostFocus()
    '***Modificado por ELRO el 20130603, según SATI INC1304290006****
    'txtDocSerie = Trim(txtDocSerie)
    '***Fin Modificado por ELRO el 20130603, según SATI INC1304290006
      'If Trim(Left(cboDoc, 3)) = "5" Then
        'txtDocSerie = Trim(Str(3)) Comentado by NAGL 20170927
    If Trim(Left(cboDoc, 3)) = "93" Then
        txtDocSerie = Trim(Str(Year(Date)))
    Else
        txtDocSerie = Right(String(4, "0") & txtDocSerie, 4)
    End If '***NAGL ERS012-2017 20170710
End Sub

Private Sub txtFecha_GotFocus()
txtFecha.SelStart = 0
txtFecha.SelLength = Len(txtFecha.Text)
End Sub
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFecha.Text) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Error"
      Exit Sub
   End If
   txtMovNro = Format(txtFecha.Text & " " & Time, gsFormatoMovFechaHora) & Right(gsCodAge, 2) & "0000" & gsCodUser
   txtBuscarProv.SetFocus
End If
End Sub
Private Sub txtFecha_LostFocus()
   If ValidaFecha(txtFecha.Text) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Error"
      Exit Sub
   End If
   txtMovNro = Format(txtFecha.Text & " " & Time, gsFormatoMovFechaHora) & Right(gsCodAge, 2) & "0000" & gsCodUser
End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   cboDoc.SetFocus
End If
End Sub

Private Sub txtDocSerie_KeyPress(KeyAscii As Integer)
'KeyAscii = NumerosEnteros(KeyAscii)
KeyAscii = LetrasNumeros(KeyAscii)
If KeyAscii = 13 Then
    '***Modificado por ELRO el 20130530, según SATI INC1304290006****
    'txtDocSerie = Trim(txtDocSerie)
    '***Fin Modificado por ELRO el 20130530, según SATI INC1304290006
    'If Trim(Left(cboDoc, 3)) <> "5" Then
        txtDocSerie = Right(String(4, "0") & txtDocSerie, 4)
    'Else
        'txtDocSerie = Trim(txtDocSerie)
    'End If '***NAGL ERS 012-2017 20170710 Comentado by NAGL 20170927
    txtDocNro.SetFocus
End If
End Sub

Private Sub txtDocNro_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
Dim oDoc As DDocumento
Set oDoc = New DDocumento

If KeyAscii = 13 Then
    If txtDocNro = "" Or txtDocNro = "00000000000" Then
        MsgBox "Ingrese Numero de Documento", vbInformation, "Aviso"
        Exit Sub
    End If
    
   '***Modificación por ELRO el 20130603, según SATI INC1304290006****
   'txtDocNro = Right(String(11, "0") & txtDocNro, 11)
   txtDocNro = Trim(txtDocNro)
   '***Fin Modificación por ELRO el 20130603, según SATI INC1304290006
   
   If cboDoc.ListIndex > -1 Then
   
    '*** PEAC 20121122
    'If (gsOpeCod = "401240" Or gsOpeCod = "402240" Or gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401344") And Trim(Left(cboDoc, 3)) = "93" Then 'si es doc. Declarac. jurada
    If Trim(Left(cboDoc, 3)) = "93" Then  'si es doc. Declarac. jurada
        lsCodigoPersona = ""
    Else
        lsCodigoPersona = txtBuscarProv.psCodigoPersona
    End If
    '*** FIN PEAC
   
   If oDoc.GetValidaDocProv(lsCodigoPersona, Mid(cboDoc.Text, 1, 2), txtDocSerie & "-" & txtDocNro) Then
    '*** PEAC 20121122
    'If (gsOpeCod = "401240" Or gsOpeCod = "402240" Or gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401344") And Trim(Left(cboDoc, 3)) = "93" Then
    If Trim(Left(cboDoc, 3)) = "93" Then
        '***Modificado por ELRO el 20130603, según SATI INC1304290006****
        'txtDocNro = Right("00000000000" + Trim(Str(Val(txtDocNro) + 1)), 11)
        txtDocNro = Trim(Str(Val(txtDocNro) + 1))
        '***Fin Modificado por ELRO el 20130603, según SATI INC1304290006
        MsgBox "El número de DDJJ acaba de ser registrado por otro usuario, por lo que se incrementó el número automaticamente, por favor vuelva a grabar.", vbInformation, "Aviso"
    Else
        MsgBox "Documento ya se encuentra registrado", vbInformation, "Aviso"
    End If
    'MsgBox "Documento ya se encuentra registrado", vbInformation, "Aviso"
    '*** FIN PEAC

'        MsgBox "Documento ya ha sido Ingresado", vbInformation, "Aviso"
        txtDocNro.SetFocus
        Exit Sub
   End If
   
   End If
   txtDocFecha.SetFocus
End If
Set oDoc = Nothing
End Sub

Private Sub txtTipPondVenta_GotFocus()
    txtTipPondVenta.SelStart = 0
    txtTipPondVenta.SelLength = 50
End Sub

Private Sub txtTipPondVenta_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTipPondVenta, KeyAscii, 7, 3)
    If KeyAscii = 13 Then
        If cmdAceptar.Enabled Then
            cmdAceptar.SetFocus
        Else
             cmdAplicar.SetFocus
        End If
    End If
End Sub

Private Sub txtTipVariable_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTipVariable, KeyAscii, 7, 3)
    If KeyAscii = 13 Then
        If txtTipPondVenta.Enabled Then
            txtTipPondVenta.SetFocus
        Else
             txtTipPondVenta.SetFocus
        End If
    End If
End Sub
Public Property Get lOk() As Boolean
lOk = OK
End Property
Public Property Let lOk(ByVal vNewValue As Boolean)
OK = vNewValue
End Property
Public Property Get DocAbrev() As String
DocAbrev = vsDocAbrev
End Property
Public Property Let DocAbrev(ByVal vNewValue As String)
vsDocAbrev = vNewValue
End Property
Public Property Get DocNro() As String
DocNro = vsDocNro
End Property
Public Property Let DocNro(ByVal vNewValue As String)
vsDocNro = vNewValue
End Property
Public Property Get Proveedor() As String
 Proveedor = vsProveedor
End Property
Public Property Let Proveedor(ByVal vNewValue As String)
vsProveedor = vNewValue
End Property
Public Property Get FechaDoc() As Date
FechaDoc = vdFechaDoc
End Property
Public Property Let FechaDoc(ByVal vNewValue As Date)
vdFechaDoc = vNewValue
End Property
Public Property Get ImporteDoc() As Currency
ImporteDoc = vnImporteDoc
End Property
Public Property Let ImporteDoc(ByVal vNewValue As Currency)
vnImporteDoc = vNewValue
End Property
Public Property Get MovNroSust() As String
MovNroSust = vsMovNroSust
End Property
Public Property Let MovNroSust(ByVal vNewValue As String)
vsMovNroSust = vNewValue
End Property
Public Property Get MovDesc() As String
MovDesc = vsMovDesc
End Property
Public Property Let MovDesc(ByVal vNewValue As String)
vsMovDesc = vNewValue
End Property

Private Sub txtTipVariable_LostFocus()
    fgDetalle_OnValidate 1, 3, True
End Sub

'*** PEAC 20121119
Sub buscar_cadena(ByVal Posicion As Integer)

Dim P As Integer, Frase As String
Dim i As Integer
Dim cTexto As String
Dim cMsg As String

cMsg = "Concepto no válido, indique el lugar de Comisión y el motivo por el que se realiza la declaración jurada."
cTexto = UCase(txtMovDesc)

    bSale = False
    
    For i = 1 To 9
        Select Case i
            Case 1
                Frase = "DECLARACIÓN JURADA"
                P = InStr(Posicion, cTexto, Frase)
                If P > 0 Then
                    m_Pos = P
                    MsgBox cMsg, vbOKOnly + vbExclamation, "Atención"
                    bSale = True
                    PalabraFrase = Frase
                    Exit Sub
                End If
            Case 2
                Frase = "DECLARACION JURADA"
                P = InStr(Posicion, cTexto, Frase)
                If P > 0 Then
                    m_Pos = P
                    MsgBox cMsg, vbOKOnly + vbExclamation, "Atención"
                    bSale = True
                    PalabraFrase = Frase
                    Exit Sub
                End If
            Case 3
                Frase = "DDJJ"
                P = InStr(Posicion, cTexto, Frase)
                If P > 0 Then
                    m_Pos = P
                    MsgBox cMsg, vbOKOnly + vbExclamation, "Atención"
                    bSale = True
                    PalabraFrase = Frase
                    Exit Sub
                End If
            Case 4
                Frase = "DD.JJ."
                P = InStr(Posicion, cTexto, Frase)
                If P > 0 Then
                    m_Pos = P
                    MsgBox cMsg, vbOKOnly + vbExclamation, "Atención"
                    bSale = True
                    PalabraFrase = Frase
                    Exit Sub
                End If
            Case 5
                Frase = "D.J."
                P = InStr(Posicion, cTexto, Frase)
                If P > 0 Then
                    m_Pos = P
                    MsgBox cMsg, vbOKOnly + vbExclamation, "Atención"
                    bSale = True
                    PalabraFrase = Frase
                    Exit Sub
                End If
            Case 6
                Frase = "DJ"
                P = InStr(Posicion, cTexto, Frase)
                If P > 0 Then
                    m_Pos = P
                    MsgBox cMsg, vbOKOnly + vbExclamation, "Atención"
                    bSale = True
                    PalabraFrase = Frase
                    Exit Sub
                End If
            Case 7
                Frase = "DD"
                P = InStr(Posicion, cTexto, Frase)
                If P > 0 Then
                    m_Pos = P
                    MsgBox cMsg, vbOKOnly + vbExclamation, "Atención"
                    bSale = True
                    PalabraFrase = Frase
                    Exit Sub
                End If
            Case 8
                Frase = "JJ"
                P = InStr(Posicion, cTexto, Frase)
                If P > 0 Then
                    m_Pos = P
                    MsgBox cMsg, vbOKOnly + vbExclamation, "Atención"
                    bSale = True
                    PalabraFrase = Frase
                    Exit Sub
                End If
            Case 9
                Frase = "D J"
                P = InStr(Posicion, cTexto, Frase)
                If P > 0 Then
                    m_Pos = P
                    MsgBox cMsg, vbOKOnly + vbExclamation, "Atención"
                    bSale = True
                    PalabraFrase = Frase
                    Exit Sub
                End If
        End Select
    Next
End Sub
'EJVG20140724 ***
Private Function MontoBaseOperacion() As Currency
    Dim i As Integer
    Dim lnMonto As Currency
    If fgDetalle.TextMatrix(1, 0) <> "" Then
        For i = 1 To fgDetalle.Rows - 1
            lnMonto = lnMonto + CCur(IIf(fgDetalle.TextMatrix(i, 3) = "", 0, fgDetalle.TextMatrix(i, 3)))
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
Private Sub cmdRetSistPensActualizar_Click()
    On Error GoTo ErrcmdRetSistPensActualizar
    Dim oPSP As NProveedorSistPens
    Dim frm As frmProveedorRegSistemaPension
    
    If Not IsDate(txtDocFecha.Text) Then
        MsgBox "Ud. debe ingresar la fecha del comprobante", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtBuscarProv.psCodigoPersona)) = 0 Then
        MsgBox "Ud. debe de seleccionar al Proveedor", vbInformation, "Aviso"
        Exit Sub
    Else
        Set oPSP = New NProveedorSistPens
        cmdRetSistPensActualizar.Enabled = False
        If oPSP.AplicaRetencionSistemaPension(txtBuscarProv.psCodigoPersona, CDate(txtDocFecha.Text), MontoBaseOperacion) Then
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
    If Not IsDate(txtDocFecha.Text) Then
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
'END EJVG *******
