VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdeudOperaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adeudados: Operaciones"
   ClientHeight    =   8280
   ClientLeft      =   1305
   ClientTop       =   615
   ClientWidth     =   8850
   Icon            =   "frmAdeudOperaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
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
      Left            =   120
      TabIndex        =   45
      Tag             =   "2"
      Text            =   "0.00"
      Top             =   7860
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.TextBox lblTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7260
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3720
      Width           =   1380
   End
   Begin VB.TextBox txtDias 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7845
      TabIndex        =   5
      Top             =   2850
      Width           =   795
   End
   Begin VB.CheckBox chkCancelacion 
      Caption         =   "Cancelación"
      Height          =   270
      Left            =   7185
      TabIndex        =   2
      Top             =   90
      Width           =   1395
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "&Procesar"
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
      Left            =   7140
      TabIndex        =   1
      Top             =   375
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cuentas de CMACT"
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
      Height          =   1905
      Left            =   90
      TabIndex        =   34
      Top             =   720
      Width           =   8670
      Begin Sicmact.FlexEdit fgInteres 
         Height          =   1515
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   8415
         _extentx        =   14843
         _extenty        =   2672
         cols0           =   23
         highlight       =   1
         encabezadosnombres=   $"frmAdeudOperaciones.frx":030A
         encabezadosanchos=   "350-1600-1500-550-0-0-1200-0-1100-0-1100-0-0-0-0-1000-0-0-0-1100-700-0-0"
         font            =   "frmAdeudOperaciones.frx":03DA
         font            =   "frmAdeudOperaciones.frx":0406
         font            =   "frmAdeudOperaciones.frx":0432
         font            =   "frmAdeudOperaciones.frx":045E
         font            =   "frmAdeudOperaciones.frx":048A
         fontfixed       =   "frmAdeudOperaciones.frx":04B6
         columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         textstylefixed  =   3
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-C-C-C-R-R-C-R-R-R-R-C-C-R-C-C-L-C-R-L-C"
         formatosedit    =   "0-0-0-0-0-0-2-2-2-2-2-2-2-0-0-2-0-0-0-0-2-0-0"
         textarray0      =   "#"
         selectionmode   =   1
         appearance      =   0
         colwidth0       =   345
         rowheight0      =   300
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   6090
      TabIndex        =   26
      Top             =   7845
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7305
      TabIndex        =   27
      Top             =   7845
      Width           =   1125
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
      Height          =   600
      Left            =   105
      TabIndex        =   29
      Top             =   105
      Width           =   3960
      Begin VB.TextBox txtOpeCod 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   930
         TabIndex        =   30
         Top             =   195
         Width           =   900
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   2535
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Operación"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
         Height          =   195
         Left            =   1995
         TabIndex        =   31
         Top             =   225
         Width           =   495
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "&Glosa"
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
      Height          =   840
      Left            =   90
      TabIndex        =   32
      Top             =   4050
      Width           =   8670
      Begin VB.TextBox txtMovDesc 
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   210
         Width           =   8460
      End
   End
   Begin TabDlg.SSTab TabDoc 
      Height          =   2805
      Left            =   90
      TabIndex        =   9
      Top             =   4950
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   4948
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Transferencia"
      TabPicture(0)   =   "frmAdeudOperaciones.frx":04E4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTransferencia"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDocTrans"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkDocOrigen"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Efectivo"
      TabPicture(1)   =   "frmAdeudOperaciones.frx":0500
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtBilleteImporte"
      Tab(1).Control(1)=   "cmdEfectivo"
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(3)=   "Shape2"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Cheque Recibido"
      TabPicture(2)   =   "frmAdeudOperaciones.frx":051C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtChqRecImporte"
      Tab(2).Control(1)=   "cmdChqRecibido"
      Tab(2).Control(2)=   "fgChqRecibido"
      Tab(2).Control(3)=   "Label2"
      Tab(2).Control(4)=   "Shape5"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Otros "
      TabPicture(3)   =   "frmAdeudOperaciones.frx":0538
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdEliminarCta"
      Tab(3).Control(1)=   "cmdAgregarCta"
      Tab(3).Control(2)=   "txtTotalOtrasCtas"
      Tab(3).Control(3)=   "fgOtros"
      Tab(3).Control(4)=   "fgObj"
      Tab(3).Control(5)=   "Label18"
      Tab(3).Control(6)=   "Shape3"
      Tab(3).ControlCount=   7
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
         TabIndex        =   14
         Top             =   2040
         Width           =   1365
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
         Height          =   675
         Left            =   180
         TabIndex        =   40
         Top             =   2040
         Width           =   4560
         Begin VB.OptionButton optDoc 
            Caption         =   "Cheque"
            Height          =   345
            Index           =   0
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Carta"
            Height          =   345
            Index           =   1
            Left            =   2310
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   2055
         End
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
         Height          =   1575
         Left            =   180
         TabIndex        =   37
         Top             =   420
         Width           =   8295
         Begin Sicmact.EditMoney txtBancoImporte 
            Height          =   255
            Left            =   6270
            TabIndex        =   13
            Top             =   1170
            Width           =   1785
            _extentx        =   2937
            _extenty        =   450
            font            =   "frmAdeudOperaciones.frx":0554
            text            =   "0.00"
            enabled         =   -1
            borderstyle     =   0
         End
         Begin Sicmact.TxtBuscar txtBuscaEntidad 
            Height          =   360
            Left            =   1095
            TabIndex        =   10
            Top             =   300
            Width           =   2580
            _extentx        =   4551
            _extenty        =   635
            appearance      =   1
            appearance      =   1
            font            =   "frmAdeudOperaciones.frx":0580
            appearance      =   1
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta N° :"
            Height          =   210
            Left            =   180
            TabIndex        =   39
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblDescIfTransf 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3735
            TabIndex        =   11
            Top             =   300
            Width           =   4350
         End
         Begin VB.Label lblDesCtaIfTransf 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1095
            TabIndex        =   12
            Top             =   720
            Width           =   6990
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
            TabIndex        =   38
            Top             =   1200
            Width           =   615
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   315
            Left            =   4980
            Top             =   1140
            Width           =   3105
         End
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
         Left            =   -68190
         TabIndex        =   28
         Tag             =   "0"
         Top             =   2355
         Width           =   1680
      End
      Begin VB.CommandButton cmdEfectivo 
         Caption         =   "Descomposición de Efectivo"
         Height          =   405
         Left            =   -69540
         TabIndex        =   17
         Top             =   1860
         Width           =   3045
      End
      Begin VB.CommandButton cmdEliminarCta 
         Caption         =   "&Eliminar"
         Height          =   360
         Left            =   -67500
         TabIndex        =   24
         Top             =   840
         Width           =   1020
      End
      Begin VB.CommandButton cmdAgregarCta 
         Caption         =   "A&gregar"
         Height          =   360
         Left            =   -67500
         TabIndex        =   23
         Top             =   450
         Width           =   1020
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
         Left            =   -69270
         TabIndex        =   22
         Tag             =   "3"
         Top             =   1425
         Width           =   1440
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
         Top             =   2385
         Width           =   1680
      End
      Begin VB.CommandButton cmdChqRecibido 
         Caption         =   "&Cheques "
         Height          =   315
         Left            =   -74790
         TabIndex        =   18
         Top             =   2400
         Width           =   1395
      End
      Begin Sicmact.FlexEdit fgChqRecibido 
         Height          =   1725
         Left            =   -74880
         TabIndex        =   19
         Top             =   540
         Width           =   8415
         _extentx        =   14843
         _extenty        =   3149
         cols0           =   12
         encabezadosnombres=   "#-Opc-Banco-NroCheque-Fecha-Importe-Cuenta-cAreaCod-cAgeCod-nMovNro-cPersCod-cIFTpo"
         encabezadosanchos=   "0-420-3200-1800-1200-1500-0-0-0-0-0-0"
         font            =   "frmAdeudOperaciones.frx":05A4
         font            =   "frmAdeudOperaciones.frx":05C8
         font            =   "frmAdeudOperaciones.frx":05EC
         font            =   "frmAdeudOperaciones.frx":0610
         font            =   "frmAdeudOperaciones.frx":0634
         fontfixed       =   "frmAdeudOperaciones.frx":0658
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
      Begin Sicmact.FlexEdit fgOtros 
         Height          =   945
         Left            =   -74850
         TabIndex        =   21
         Top             =   450
         Width           =   7305
         _extentx        =   12779
         _extenty        =   1667
         cols0           =   4
         highlight       =   1
         encabezadosnombres=   "#-Cuenta-Descripcion-Importe"
         encabezadosanchos=   "300-1800-3500-1300"
         font            =   "frmAdeudOperaciones.frx":0686
         font            =   "frmAdeudOperaciones.frx":06AA
         font            =   "frmAdeudOperaciones.frx":06CE
         font            =   "frmAdeudOperaciones.frx":06F2
         font            =   "frmAdeudOperaciones.frx":0716
         fontfixed       =   "frmAdeudOperaciones.frx":073A
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
      Begin Sicmact.FlexEdit fgObj 
         Height          =   945
         Left            =   -74850
         TabIndex        =   25
         Top             =   1770
         Width           =   6195
         _extentx        =   10927
         _extenty        =   1667
         cols0           =   8
         highlight       =   2
         allowuserresizing=   1
         encabezadosnombres=   "#-Ord-Código-Descripción-CtaCont-SubCta-ObjPadre-ItemCtaCont"
         encabezadosanchos=   "350-400-1200-3000-0-900-0-0"
         font            =   "frmAdeudOperaciones.frx":0768
         font            =   "frmAdeudOperaciones.frx":0794
         font            =   "frmAdeudOperaciones.frx":07C0
         font            =   "frmAdeudOperaciones.frx":07EC
         font            =   "frmAdeudOperaciones.frx":0818
         fontfixed       =   "frmAdeudOperaciones.frx":0844
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
         Left            =   -69330
         TabIndex        =   43
         Top             =   2430
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
         Left            =   -70680
         TabIndex        =   42
         Top             =   1455
         Width           =   615
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
         TabIndex        =   41
         Top             =   2430
         Width           =   615
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   -69420
         Top             =   2370
         Width           =   3045
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   -69540
         Top             =   2340
         Width           =   3045
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   -70890
         Top             =   1410
         Width           =   3105
      End
   End
   Begin Sicmact.FlexEdit fgDetalle 
      Height          =   1365
      Left            =   90
      TabIndex        =   4
      Top             =   2670
      Width           =   6285
      _extentx        =   11086
      _extenty        =   2408
      cols0           =   7
      highlight       =   1
      encabezadosnombres=   "#-Cuenta-Descripcion-Monto-Pos-Objeto-Monto VAC"
      encabezadosanchos=   "350-1200-2000-1200-0-0-1200"
      font            =   "frmAdeudOperaciones.frx":0872
      font            =   "frmAdeudOperaciones.frx":089E
      font            =   "frmAdeudOperaciones.frx":08CA
      font            =   "frmAdeudOperaciones.frx":08F6
      font            =   "frmAdeudOperaciones.frx":0922
      fontfixed       =   "frmAdeudOperaciones.frx":094E
      columnasaeditar =   "X-X-X-3-X-X-X"
      textstylefixed  =   3
      listacontroles  =   "0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-L-L-R-L-L-R"
      formatosedit    =   "0-0-0-2-0-0-2"
      textarray0      =   "#"
      lbeditarflex    =   -1
      lbbuscaduplicadotext=   -1
      appearance      =   0
      colwidth0       =   345
      rowheight0      =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL :"
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
      Left            =   6510
      TabIndex        =   44
      Top             =   3750
      Width           =   735
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Dias :"
      Height          =   195
      Left            =   7380
      TabIndex        =   36
      Top             =   2940
      Width           =   405
   End
   Begin VB.Label lblTasaVAC 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   7440
      TabIndex        =   6
      Top             =   3285
      Width           =   1200
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Indice VAC :"
      Height          =   195
      Left            =   6525
      TabIndex        =   35
      Top             =   3345
      Width           =   885
   End
End
Attribute VB_Name = "frmAdeudOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lMN As Boolean
Dim lsCtaContDebe() As String
Dim lsCtaContHaber() As String
Dim aObj() As String
Dim lbCargar As Boolean
Dim lsGridDH As String
Dim lsPosCtaBusqueda As String
Dim lnTasaVac As Double
Dim lsCtaConcesional As String
Dim oAdeud As DCaja_Adeudados
Dim oOpe As DOperacion
Dim oCtaIf As NCajaCtaIF
Dim lsCtaOrdenD As String
Dim lsCtaOrdenH As String


'Efectivo
Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset

'Documento de Transferencia
Dim lsDocumento As String
Dim lnTpoDoc As TpoDoc
Dim lsNroDoc As String
Dim lsNroVoucher As String
Dim ldFechaDoc  As Date

'Variable para Refrescar Cuentas a Utilizar
Dim lsIFTpo As String

Private Sub CalculaTotalRetiros()
txtImporte = Format(nVal(txtBancoImporte) + nVal(txtBilleteImporte) + nVal(txtChqRecImporte) + nVal(txtTotalOtrasCtas), gsFormatoNumeroView)
End Sub

Private Sub chkCancelacion_Click()
    If chkCancelacion.value = 1 Then
        fgInteres.ColWidth(3) = 0
    Else
        fgInteres.ColWidth(3) = 500
    End If
    Me.cmdCalcular.SetFocus
End Sub

Private Sub chkDocOrigen_Click()
    If chkDocOrigen.value = Checked Then
        fraDocTrans.Enabled = True
    Else
        fraDocTrans.Enabled = False
    End If
End Sub

Private Function ValidaDatos() As Boolean
Dim lbMontoDet As Boolean
Dim i As Integer
ValidaDatos = False
    If lblTotal = 0 Then
        MsgBox "No se seleccionó Adeudado a pagar...", vbInformation, "¡AViso1"
        Exit Function
    End If
    If txtImporte <> lblTotal Then
        MsgBox "Monto a pagar no coincide con Cuota de deuda", vbInformation, "¡Aviso!"
        Exit Function
    End If
    If fgInteres.TextMatrix(1, 0) = "" Then
        MsgBox "No existen Cuentas de para realizar la Operación", vbInformation, "Aviso"
        Me.cmdSalir.SetFocus
        Exit Function
    End If
    If Val(lblTotal) <= 0 Then
        MsgBox "El Monto de operación no Válido", vbInformation, "Aviso"
        fgInteres.SetFocus
        Exit Function
    End If
    
    If nVal(txtBancoImporte) > 0 And txtBuscaEntidad = "" Then
        MsgBox "Cuenta de Banco no Válida", vbInformation, "Aviso"
        txtBuscaEntidad.SetFocus
        Exit Function
    End If
    If txtMovDesc = "" Then
        MsgBox "Ingrese Descripción de Operación !!!", vbInformation, "Aviso"
        txtMovDesc.SetFocus
        Exit Function
    End If
    lbMontoDet = False
ValidaDatos = True
End Function

Private Sub cmdAceptar_Click()
Dim oDocPago As clsDocPago
Dim lsCuentaAho As String

Dim lsMovNro As String
Dim oCon     As NContFunciones
Dim oCaja As nCajaGeneral
Dim rsAdeud  As ADODB.Recordset

On Error GoTo AceptarErr
If Not ValidaDatos() Then
    Exit Sub
End If

Set oDocPago = New clsDocPago
Set oCon = New NContFunciones
Set oCaja = New nCajaGeneral

If MsgBox(" ¿ Desea Grabar Operación ? ", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
    lsMovNro = oCon.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)

    oCaja.GrabaPagoCuotaAdeudados lsMovNro, gsOpeCod, txtFecha, txtMovDesc, nVal(txtImporte), _
            fgInteres.TextMatrix(fgInteres.Row, 18), fgInteres.TextMatrix(fgInteres.Row, 3), rsBill, rsMon, _
            gdFecSis, txtBuscaEntidad, txtBancoImporte, _
            lnTpoDoc, lsNroDoc, ldFechaDoc, lsNroVoucher, _
            fgChqRecibido.GetRsNew, fgOtros.GetRsNew, fgObj.GetRsNew, fgDetalle.GetRsNew, _
            chkCancelacion.value = vbChecked, Format(lblTasaVAC, "#,##0.00###"), lsCtaConcesional, fgInteres.TextMatrix(fgInteres.Row, 21), lsCtaOrdenD, lsCtaOrdenH, fgInteres.TextMatrix(fgInteres.Row, 22)
    ImprimeAsientoContable lsMovNro, lsNroVoucher, lnTpoDoc, lsDocumento, True, False
    If MsgBox("Desea Realizar otra operación de Pago ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        txtBuscaEntidad = ""
        txtImporte = "0.00"
        txtBancoImporte = "0.00"
        txtBilleteImporte = "0.00"
        txtTotalOtrasCtas = "0.00"
        
        txtMovDesc = ""
        Set rsBill = Nothing
        Set rsMon = Nothing
        lblDescIfTransf = ""
        lblDesCtaIfTransf = ""
        
        lsDocumento = ""
        lnTpoDoc = -1
        lsNroDoc = ""
        lsNroVoucher = ""
        
        LimpiaControles
        fgOtros.Clear
        fgOtros.Rows = 2
        fgOtros.FormaCabecera
        
        fgObj.Clear
        fgObj.Rows = 2
        fgObj.FormaCabecera
    Else
        Unload Me
    End If
End If
    
Exit Sub
AceptarErr:
    MsgBox "Error N° [" & Err.Number & "] " & TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub LimpiaControles(Optional lsTipo As Integer = 0)
    txtMovDesc = ""
    lblTotal = "0.00"
    txtBilleteImporte = "0.00"
    CargaCuentasGrid ""
    
End Sub

Private Sub cmdAgregarCta_Click()
Dim oOpe As New DOperacion
fgOtros.AdicionaFila
fgOtros.rsTextBuscar = oOpe.EmiteOpeCtasNivel(gsOpeCod, , "4")
fgOtros.SetFocus
Set oOpe = Nothing
End Sub

Private Sub cmdCalcular_Click()
lsIFTpo = ""
    If ValFecha(Me.txtFecha) = False Then Exit Sub
    CargaBancos CDate(txtFecha)
    fgInteres.SetFocus
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
        nRow = fgChqRecibido.Row
        fgChqRecibido.TextMatrix(nRow, 2) = rs!banco
        fgChqRecibido.TextMatrix(nRow, 3) = rs!cNroDoc
        fgChqRecibido.TextMatrix(nRow, 4) = rs!Fecha
        fgChqRecibido.TextMatrix(nRow, 5) = rs!nMonto
        fgChqRecibido.TextMatrix(nRow, 5) = rs!Objeto
        fgChqRecibido.TextMatrix(nRow, 6) = rs!cAreaCod
        fgChqRecibido.TextMatrix(nRow, 7) = rs!cAgecod
        fgChqRecibido.TextMatrix(nRow, 8) = rs!nMovNro
        fgChqRecibido.TextMatrix(nRow, 9) = rs!cPersCod
        fgChqRecibido.TextMatrix(nRow, 10) = rs!cIFTpo
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

Private Sub cmdEliminarCta_Click()
If fgOtros.TextMatrix(fgOtros.Row, 0) <> "" Then
   EliminaCuenta fgOtros.TextMatrix(fgOtros.Row, 1), fgOtros.TextMatrix(fgOtros.Row, 0)
   txtTotalOtrasCtas.Text = Format(fgOtros.SumaRow(3), gsFormatoNumeroView)
End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fgDetalle_OnCellChange(pnRow As Long, pnCol As Long)
If fgInteres.TextMatrix(fgInteres.Row, 17) = "2" And Mid(fgInteres.TextMatrix(fgInteres.Row, 18), 20, 1) = "1" Then
    If nVal(lblTasaVAC) > 0 Then
       fgDetalle.TextMatrix(pnRow, pnCol + 1) = Format(nVal(fgDetalle.TextMatrix(pnRow, pnCol)) / Format(lblTasaVAC, "#,#0.00###"), "#0.00")
    End If
Else
    fgDetalle.TextMatrix(fgDetalle.Row, 6) = Format(nVal(fgDetalle.TextMatrix(pnRow, pnCol)), "#0.00")
End If
Total True
End Sub

Private Sub fgInteres_GotFocus()
    Call fgInteres_RowColChange
End Sub

Private Sub fgInteres_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDias.SetFocus
    End If
End Sub

Private Sub fgInteres_RowColChange()
    If Me.fgInteres.TextMatrix(1, 0) <> "" Then
        Call Total(False)
        
    End If
End Sub

Private Sub fgOtros_OnCellChange(pnRow As Long, pnCol As Long)
txtTotalOtrasCtas = Format(fgOtros.SumaRow(3), gsFormatoNumeroView)
CalculaTotalRetiros
End Sub

Private Sub fgOtros_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim oCta As New DGeneral
If psDataCod <> "" Then
    AsignaCtaObj psDataCod
    fgOtros.Col = 2
End If
Set oCta = Nothing
End Sub

Private Sub fgOtros_RowColChange()
RefrescaFgObj Val(fgOtros.TextMatrix(fgOtros.Row, 0))
End Sub

Private Sub Form_Activate()
    If lbCargar = False Then
        Unload Me
    End If
End Sub

Public Sub Inicio(psGridDH As String, psPosCtaBusqueda As String)
    lsGridDH = psGridDH
    lsPosCtaBusqueda = psPosCtaBusqueda
    Me.Show 1
End Sub
Private Sub Form_Load()
    Dim N As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer, J As Integer

    CentraForm Me
    TabDoc.Tab = 0
    Me.Caption = gsOpeDesc
    gsSimbolo = gcMN
    If Mid(gsOpeCod, 3, 1) = "2" Then
        gsSimbolo = gcME
    End If

    lbCargar = True
    Set oOpe = New DOperacion
    Set oAdeud = New DCaja_Adeudados
    Set oCtaIf = New NCajaCtaIF


    lnTasaVac = oAdeud.CargaIndiceVAC(gdFecSis)
    lblTasaVAC = lnTasaVac

    LimpiaControles

    txtOpeCod = gsOpeCod
    txtFecha.Text = Format(gdFecSis, gsFormatoFechaView)
    txtBuscaEntidad.rs = oOpe.GetOpeObj(gsOpeCod, "2")
    Set rs = oOpe.CargaOpeCta(gsOpeCod, "H", "6")
    If Not rs.EOF Then
        lsCtaConcesional = rs!cCtaContCod
    End If
    RSClose rs
    
    lsCtaOrdenD = oOpe.EmiteOpeCta(gsOpeCod, "D", 7)
    lsCtaOrdenH = oOpe.EmiteOpeCta(gsOpeCod, "H", 7)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oAdeud = Nothing
End Sub

Private Sub OptDoc_Click(Index As Integer)
Dim oDocPago As clsDocPago
    Set oDocPago = New clsDocPago
    If optDoc(0).value Then
        oDocPago.InicioCheque "", True, Mid(txtBuscaEntidad, 4, 13), gsOpeCod, gsNomCmac, gsOpeDesc, txtMovDesc, _
                     CCur(txtImporte), gdFecSis, gsNomCmacRUC, txtBuscaEntidad, lblDescIfTransf, _
                     lblDesCtaIfTransf, "", True, , Mid(txtBuscaEntidad, 18, 10), , Mid(txtBuscaEntidad, 1, 2), Mid(txtBuscaEntidad, 4, 13), Mid(txtBuscaEntidad, 18, 10) 'EJVG20121130
                     'lblDesCtaIfTransf, "", True, , Mid(txtBuscaEntidad, 18, 10)
        If oDocPago.vbOk Then
            lsDocumento = oDocPago.vsFormaDoc
            lnTpoDoc = Val(oDocPago.vsTpoDoc)
            lsNroDoc = oDocPago.vsNroDoc
            lsNroVoucher = oDocPago.vsNroVoucher
            ldFechaDoc = oDocPago.vdFechaDoc
            txtMovDesc = oDocPago.vsGlosa
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
            txtMovDesc = oDocPago.vsGlosa
            optDoc(1).value = True
        Else
            optDoc(1).value = False
            Exit Sub
        End If
    End If
    Set oDocPago = Nothing
End Sub

Private Sub txtBancoImporte_GotFocus()
fEnfoque txtBancoImporte
End Sub

Private Sub txtBancoImporte_KeyPress(KeyAscii As Integer)
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
If txtBancoImporte.Visible Then
    txtBancoImporte.SetFocus
End If
End Sub

Private Sub txtDias_GotFocus()
    fEnfoque txtDias
End Sub

Private Sub txtDias_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        CalculoInteres Int(txtDias)
        fgInteres.SetFocus
    End If
End Sub

Private Sub txtFecha_GotFocus()
    fEnfoque txtFecha
End Sub
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ValFecha(txtFecha) = False Then
            txtFecha.SetFocus
        Else
            cmdCalcular.SetFocus
            lnTasaVac = oAdeud.CargaIndiceVAC(txtFecha)
            lblTasaVAC = lnTasaVac
        End If
    End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Select Case TabDoc.Tab
            Case 0:
                txtBuscaEntidad.SetFocus
            Case 1
                cmdEfectivo.SetFocus
            Case 2
                cmdChqRecibido.SetFocus
            Case 3
                cmdAgregarCta.SetFocus
        End Select
    End If
End Sub

Private Sub CargaBancos(ldFecha As Date)
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim N As Integer
    Dim lnMontoTotal As Currency
    Dim lnInteres As Currency
    Dim lnTotal As Integer, i As Integer
    Dim lnCapital As Currency

    lnTasaVac = oAdeud.CargaIndiceVAC(ldFecha)
    If lnTasaVac = 0 Then
        If MsgBox("Tasa VAC no ha sido definida para la fecha Ingresada" & Chr(13) & "Desea Proseguir con al Operación??", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
    End If
    lblTasaVAC = lnTasaVac
    
    'CARGAMOS LOS ADEUDADOS PENDIENTES
    Set rs = oAdeud.GetAdeudadosProvision(gsOpeCod, ldFecha, Mid(gsOpeCod, 3, 1))

    lnTotal = rs.RecordCount
    i = 0
    fgInteres.Rows = 2
    fgInteres.Clear
    fgInteres.FormaCabecera
    
    Do While Not rs.EOF
        i = i + 1
        fgInteres.AdicionaFila
        N = fgInteres.Row

        fgInteres.TextMatrix(N, 1) = Trim(rs!cPersNombre)   'entidad
        fgInteres.TextMatrix(N, 2) = Trim(rs!cCtaIFDesc)    'cuenta
        fgInteres.TextMatrix(N, 3) = Trim(rs!nNroCuota)    ' numero de cuota pendiente
        'se oculta *
        fgInteres.TextMatrix(N, 4) = Format(rs!nSaldoCap, "#,#0.00") 'Saldocapital
        lnCapital = IIf(Me.chkCancelacion.value = 1, rs!nSaldoCap, rs!nCapitalCuota)
        'se oculta *
        fgInteres.TextMatrix(N, 5) = Format(lnCapital, "#,#0.00")  ' Saldo de Capital Base
        'Se muestra
        If rs!cMonedaPago = "2" And Mid(rs!cCtaIFCod, 3, 1) = "1" Then
            fgInteres.TextMatrix(N, 6) = Format(lnCapital * lnTasaVac, "#,#0.00") ' Saldo * la tasa vac
        Else
            fgInteres.TextMatrix(N, 6) = Format(lnCapital, "#,#0.00")  ' Saldo de Capital Normal
        End If
        'se oculta *

        fgInteres.TextMatrix(N, 7) = Format(rs!nInteresPagado, "#,#0.00")  ' Interes acumulado pagado por cuota
        fgInteres.TextMatrix(N, 8) = Format(rs!nInteresPagado, "#,#0.00")
        
        If Val(rs!cIFTpo) = gTpoIFFuenteFinanciamiento Then
            lnMontoTotal = rs!nSaldoCap - rs!nSaldoConcesion
        Else
            lnMontoTotal = rs!nSaldoCap + rs!nInteresPagado
        End If
        lnInteres = oAdeud.CalculaInteres(rs!nDiasUltPAgo, rs!nPeriodo, rs!nInteres, lnMontoTotal)
        
        'se oculta
        fgInteres.TextMatrix(N, 9) = Format(lnInteres, "#,#0.00")
        'se muestra
        If rs!cMonedaPago = "2" And Mid(rs!cCtaIFCod, 3, 1) = "1" Then
            lnInteres = lnInteres * lnTasaVac
        End If
        fgInteres.TextMatrix(N, 10) = Format(lnInteres, "#,#0.00")
        'se oculta *

        fgInteres.TextMatrix(N, 11) = Format(rs!nInteresPagado + lnInteres, "#,#0.00")
        fgInteres.TextMatrix(N, 12) = Format((rs!nInteresPagado + lnInteres), "#0.00")

        fgInteres.TextMatrix(N, 13) = Format(rs!dCuotaUltPago, "dd/mm/yyyy")
        fgInteres.TextMatrix(N, 14) = Trim(rs!nPeriodo)
        fgInteres.TextMatrix(N, 20) = Trim(rs!nInteres)
        fgInteres.TextMatrix(N, 16) = Trim(rs!nDiasUltPAgo)
        fgInteres.TextMatrix(N, 17) = Trim(rs!cMonedaPago)
        fgInteres.TextMatrix(N, 18) = Trim(rs!cIFTpo & "." & rs!cPersCod & "." & rs!cCtaIFCod)
        fgInteres.TextMatrix(N, 19) = rs!dVencimiento
        fgInteres.TextMatrix(N, 15) = rs!nComision
        fgInteres.TextMatrix(N, 21) = rs!nSaldoCapLP
        fgInteres.TextMatrix(N, 22) = rs!cCodLinCred
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
Private Function Total(lbTotal As Boolean) As Currency
    Dim i As Integer
    Dim lnTotal As Currency
    If Not lsIFTpo = Left(fgInteres.TextMatrix(fgInteres.Row, 18), 2) Then
        lsIFTpo = Left(fgInteres.TextMatrix(fgInteres.Row, 18), 2)
        CargaCuentasGrid fgInteres.TextMatrix(fgInteres.Row, 18)
    End If
    If lbTotal = False Then
        For i = 1 To Me.fgDetalle.Rows - 1
            fgDetalle.TextMatrix(i, 5) = Trim(fgInteres.TextMatrix(fgInteres.Row, 18))
            Select Case Trim(fgDetalle.TextMatrix(i, 4))
                Case "0"
                    fgDetalle.TextMatrix(i, 3) = Format(fgInteres.TextMatrix(fgInteres.Row, 6), "#0.00")  'capital que se muestra
                    fgDetalle.TextMatrix(i, 6) = Format(fgInteres.TextMatrix(fgInteres.Row, 5), "#0.00")  'capital que se oculta
                Case "1"
                    fgDetalle.TextMatrix(i, 3) = Format(fgInteres.TextMatrix(fgInteres.Row, 10), "#0.00")
                    fgDetalle.TextMatrix(i, 6) = Format(fgInteres.TextMatrix(fgInteres.Row, 9), "#0.00")
                Case "2"
                    fgDetalle.TextMatrix(i, 3) = Format(fgInteres.TextMatrix(fgInteres.Row, 8), "#0.00")
                    fgDetalle.TextMatrix(i, 6) = Format(fgInteres.TextMatrix(fgInteres.Row, 7), "#0.00")
                Case "3"  'Comision
                    fgDetalle.TextMatrix(i, 3) = Format(fgInteres.TextMatrix(fgInteres.Row, 15), "#0.00")
                    fgDetalle.TextMatrix(i, 6) = Format(fgInteres.TextMatrix(fgInteres.Row, 15), "#0.00")
            End Select
        Next
        txtDias = fgInteres.TextMatrix(fgInteres.Row, 16)
    End If
    lnTotal = 0
    For i = 1 To fgDetalle.Rows - 1
        lnTotal = lnTotal + CCur(IIf(fgDetalle.TextMatrix(i, 3) = "", "0", fgDetalle.TextMatrix(i, 3)))
    Next
    lblTotal = Format(lnTotal, "#,#0.00")
    lblTasaVAC = 0
    If fgInteres.TextMatrix(fgInteres.Row, 17) = "2" And Mid(fgInteres.TextMatrix(fgInteres.Row, 18), 20, 1) = "1" Then
       lblTasaVAC = lnTasaVac
    End If
End Function

'Private Sub CargaFlex(psObjCod As String)
'    Dim nLong As Integer, nItem As Integer, Nivel As Integer
'    Dim lsObjDesc As String, SQLObjs As String
'    Dim rs As ADODB.Recordset
'
'    fgBillete.Clear
'    fgBillete.Rows = 2
'    fgBillete.TextMatrix(0, 2) = "Descripción"
'    fgBillete.TextMatrix(0, 3) = "Monto "
'    fgBillete.ColWidth(0) = 250
'    fgBillete.ColWidth(1) = 0
'    fgBillete.ColWidth(2) = 5800
'    fgBillete.ColWidth(3) = 1700
'    fgBillete.ColWidth(4) = 0
'    fgBillete.ColAlignment(3) = 7
'    fgBillete.ColAlignmentFixed(3) = 7
'    fgBillete.Row = 1
'    fgBillete.TopRow = 1
'
'    nLong = Len(Trim(psObjCod))  ' Del Objeto Tipo
'
'    SQLObjs = "SELECT cObjetoCod, cObjetoDesc from  Objeto " _
'       & "where SubString(cObjetoCod,1," & nLong & ") = '" & psObjCod _
'       & "' and nobjetoNiv = 4 Order by cObjetoCod Desc"
'    Set rs = CargaRecord(SQLObjs)
'    If RSVacio(rs) Then
'        MsgBox "Error al asignar Objetos a la Operación. Consultar con Sistemas" & psObjCod & " no tiene Instancias ", vbCritical, "Error de Objetos"
'        Unload Me
'    End If
'    Do While Not rs.EOF
'        AdicionaRow fgBillete
'        nItem = fgBillete.Row
'        fgBillete.TextMatrix(nItem, 0) = ""
'        fgBillete.TextMatrix(nItem, 1) = rs!cObjetoCod
'        fgBillete.TextMatrix(nItem, 2) = rs!cObjetoDesc
'        fgBillete.TextMatrix(nItem, 4) = "D"
'        fgBillete.Col = 2
'        fgBillete.CellBackColor = "&H00DBDBDB"
'        rs.MoveNext
'    Loop
'    rs.Close
'    Set rs = Nothing
'    fgBillete.Row = 1
'    fgBillete.TopRow = 1
'End Sub

Private Sub CargaCuentasGrid(psIFCod As String)
    Dim rs As ADODB.Recordset
    Dim oOpe As New DOperacion
If Not psIFCod = "" Then
    Set rs = oOpe.CargaOpeCtaIF(gsOpeCod, psIFCod, "D")
    If Not RSVacio(rs) Then
        fgDetalle.Rows = 2
        fgDetalle.Clear
        fgDetalle.FormaCabecera
        Do While Not rs.EOF
            fgDetalle.AdicionaFila
            fgDetalle.TextMatrix(fgDetalle.Row, 1) = rs!cCtaContCod
            fgDetalle.TextMatrix(fgDetalle.Row, 2) = Trim(rs!cCtaContDesc)
            fgDetalle.TextMatrix(fgDetalle.Row, 4) = Trim(rs!cOpeCtaOrden)
            rs.MoveNext
        Loop
    Else
        lbCargar = False
        MsgBox "No se han definido Cuentas Contables para Operación", vbInformation, "Aviso"
    End If
End If
    fgDetalle.TopRow = 1
    fgDetalle.Row = 1
    RSClose rs
    Set oOpe = Nothing
End Sub

Private Sub CalculoInteres(ByVal lnDias As Long)
    Dim lnPeriodo As Long
    Dim lnTasaInt As Currency
    Dim lnMontoTotal As Currency
    Dim lnInteres As Currency

    fgInteres.TextMatrix(fgInteres.Row, 16) = lnDias
    If Val(Mid(fgInteres.TextMatrix(fgInteres.Row, 18), 1, 2)) = gTpoIFFuenteFinanciamiento Then
        lnMontoTotal = CCur(fgInteres.TextMatrix(fgInteres.Row, 4))
    Else
        lnMontoTotal = CCur(fgInteres.TextMatrix(fgInteres.Row, 4)) + CCur(fgInteres.TextMatrix(fgInteres.Row, 7))
    End If
    lnTasaInt = CCur(fgInteres.TextMatrix(fgInteres.Row, 20))
    lnPeriodo = Val(fgInteres.TextMatrix(fgInteres.Row, 14))
    lnInteres = oAdeud.CalculaInteres(lnDias, lnPeriodo, lnTasaInt, lnMontoTotal)

    fgInteres.TextMatrix(fgInteres.Row, 9) = Format(lnInteres, "#0.00")
    If fgInteres.TextMatrix(fgInteres.Row, 17) = "2" And Mid(fgInteres.TextMatrix(fgInteres.Row, 18), 9, 1) = "1" Then
        fgInteres.TextMatrix(fgInteres.Row, 10) = Format(lnInteres * lnTasaVac, "#0.00")
    Else
        fgInteres.TextMatrix(fgInteres.Row, 10) = Format(lnInteres, "#0.00")
    End If
End Sub

Public Sub AsignaCtaObj(ByVal psCtaContCod As String)
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
EliminaFgObj Val(fgOtros.TextMatrix(fgOtros.Row, 0))
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
                            AdicionaObj psCtaContCod, fgOtros.TextMatrix(fgOtros.Row, 0), rs1!nCtaObjOrden, oDescObj.gsSelecCod, _
                                        oDescObj.gsSelecDesc, lsFiltro, rs1!cObjetoCod
                        Else
                            fgOtros.EliminaFila fgOtros.Row, False
                            Exit Do
                        End If
                    Else
                        AdicionaObj psCtaContCod, fgOtros.TextMatrix(fgOtros.Row, 0), rs1!nCtaObjOrden, rs1!cObjetoCod, _
                                        rs1!cObjetoDesc, lsFiltro, rs1!cObjetoCod
                    End If
                End If
            End If
        Else
            If Val(rs1!cObjetoCod) = ObjPersona Then
                Set UP = frmBuscaPersona.Inicio
                If Not UP Is Nothing Then
                    AdicionaObj psCtaContCod, fgOtros.TextMatrix(fgOtros.Row, 0), rs1!nCtaObjOrden, _
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
    nItem = fgObj.Row
    fgObj.TextMatrix(nItem, 0) = nFila
    fgObj.TextMatrix(nItem, 1) = psOrden
    fgObj.TextMatrix(nItem, 2) = psObjetoCod
    fgObj.TextMatrix(nItem, 3) = psObjDescripcion
    fgObj.TextMatrix(nItem, 4) = sCodCta
    fgObj.TextMatrix(nItem, 5) = psSubCta
    fgObj.TextMatrix(nItem, 6) = psObjPadre
    fgObj.TextMatrix(nItem, 7) = nFila
    
End Sub

Private Sub EliminaCuenta(sCod As String, nItem As Integer)
If fgOtros.TextMatrix(1, 0) <> "" Then
    EliminaFgObj Val(fgOtros.TextMatrix(fgOtros.Row, 0))
    fgOtros.EliminaFila fgOtros.Row
End If
If Len(fgOtros.TextMatrix(1, 1)) > 0 Then
   RefrescaFgObj Val(fgOtros.TextMatrix(fgOtros.Row, 0))
End If
End Sub

Private Sub RefrescaFgObj(nItem As Integer)
Dim k  As Integer
For k = 1 To fgObj.Rows - 1
    If Len(fgObj.TextMatrix(k, 1)) Then
       If fgObj.TextMatrix(k, 0) = nItem Then
          fgObj.RowHeight(k) = 285
       Else
          fgObj.RowHeight(k) = 0
       End If
    End If
Next
End Sub


Private Sub EliminaFgObj(nItem As Integer)
Dim k  As Integer, m As Integer
k = 1
Do While k < fgObj.Rows
   If Len(fgObj.TextMatrix(k, 0)) > 0 Then
      If Val(fgObj.TextMatrix(k, 0)) = nItem Then
         fgObj.EliminaFila k, False
      Else
        If CCur(fgObj.TextMatrix(k, 0)) > nItem Then
           fgObj.TextMatrix(k, 0) = CCur(fgObj.TextMatrix(k, 0)) - 1
        End If
         k = k + 1
      End If
   Else
      k = k + 1
   End If
Loop
End Sub


