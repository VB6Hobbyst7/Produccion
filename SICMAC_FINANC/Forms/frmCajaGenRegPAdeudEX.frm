VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCajaGenRegPAdeudEX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adeudados: Registro de Pagare"
   ClientHeight    =   8040
   ClientLeft      =   765
   ClientTop       =   1455
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajaGenRegPAdeudEX.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   9075
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
      TabIndex        =   43
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
      Caption         =   "Datos de Pagaré"
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
      Height          =   1755
      Left            =   150
      TabIndex        =   42
      Top             =   1230
      Width           =   8835
      Begin Sicmact.EditMoney txtPorcentAfect 
         Height          =   315
         Left            =   7800
         TabIndex        =   4
         Top             =   285
         Width           =   645
         _extentx        =   1138
         _extenty        =   556
         font            =   "frmCajaGenRegPAdeudEX.frx":000C
         forecolor       =   128
         text            =   "0"
         enabled         =   -1
      End
      Begin Sicmact.TxtBuscar txtLinCredCod 
         Height          =   315
         Left            =   4200
         TabIndex        =   6
         Top             =   720
         Width           =   1485
         _extentx        =   2619
         _extenty        =   556
         appearance      =   1
         appearance      =   1
         font            =   "frmCajaGenRegPAdeudEX.frx":0038
         appearance      =   1
         stitulo         =   ""
      End
      Begin VB.TextBox txtCuentaDesc 
         Height          =   315
         Left            =   2880
         TabIndex        =   3
         Top             =   285
         Width           =   2295
      End
      Begin VB.TextBox txtCuentaCod 
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   2
         Top             =   285
         Width           =   1005
      End
      Begin VB.TextBox txtComisionInicial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2055
         TabIndex        =   7
         Text            =   "0"
         Top             =   1140
         Width           =   750
      End
      Begin VB.TextBox txtComisionMonto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3450
         TabIndex        =   8
         Text            =   "0"
         Top             =   1140
         Width           =   1410
      End
      Begin VB.TextBox txtCapital 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   765
         Width           =   1410
      End
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
         Left            =   6960
         TabIndex        =   9
         Tag             =   "0"
         Text            =   "0.00"
         Top             =   1140
         Width           =   1680
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   8490
         TabIndex        =   59
         Top             =   375
         Width           =   135
      End
      Begin VB.Label lblPorcentAfect 
         Caption         =   "Porcent. de Afectacion : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   5745
         TabIndex        =   58
         Top             =   330
         Width           =   2220
      End
      Begin VB.Label lblLinCredDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5700
         TabIndex        =   57
         Top             =   720
         Width           =   2955
      End
      Begin VB.Label Label9 
         Caption         =   "Linea de Crédito"
         Height          =   225
         Left            =   2970
         TabIndex        =   56
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Número"
         Height          =   285
         Left            =   2190
         TabIndex        =   55
         Top             =   315
         Width           =   825
      End
      Begin VB.Label Label7 
         Caption         =   "Código"
         Height          =   285
         Left            =   240
         TabIndex        =   54
         Top             =   315
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "Tasa"
         Height          =   285
         Left            =   1590
         TabIndex        =   53
         Top             =   1185
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comision Inicial :"
         Height          =   210
         Left            =   240
         TabIndex        =   52
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label lblComision 
         Alignment       =   2  'Center
         Caption         =   "Monto"
         Height          =   285
         Left            =   2940
         TabIndex        =   51
         Top             =   1185
         Width           =   465
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Capital :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   50
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Neto a Desembolsar"
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
         Left            =   5130
         TabIndex        =   45
         Top             =   1170
         Width           =   1725
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   5010
         Top             =   1125
         Width           =   3645
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7650
      TabIndex        =   32
      Top             =   7560
      Width           =   1275
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6330
      TabIndex        =   31
      Top             =   7560
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
      TabIndex        =   33
      Top             =   90
      Width           =   7185
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
         font            =   "frmCajaGenRegPAdeudEX.frx":0064
         appearance      =   1
      End
      Begin VB.Label lblDescTipoCta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3030
         TabIndex        =   11
         Top             =   660
         Width           =   4005
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Institución Financiera :"
         Height          =   420
         Left            =   135
         TabIndex        =   34
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
         TabIndex        =   10
         Top             =   255
         Width           =   4005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1380
      Left            =   150
      TabIndex        =   35
      Top             =   2970
      Width           =   8835
      Begin VB.CommandButton cmdCalendario 
         Caption         =   "Calen&dario"
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
         Left            =   7290
         TabIndex        =   14
         Top             =   630
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton cmdCartAper 
         Caption         =   "&Carta Apertura"
         Height          =   375
         Left            =   7290
         TabIndex        =   13
         Top             =   210
         Width           =   1380
      End
      Begin VB.TextBox txtMovDesc 
         Height          =   600
         Left            =   135
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   270
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
         TabIndex        =   15
         Tag             =   "2"
         Text            =   "0.00"
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total Ingresos :"
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
         TabIndex        =   36
         Top             =   1005
         Width           =   1350
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
      Left            =   150
      TabIndex        =   16
      Top             =   4410
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   5371
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Transferencia"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chkDocOrigen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTransferencia"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDocTrans"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Efectivo"
      TabPicture(1)   =   "frmCajaGenRegPAdeudEX.frx":0088
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdEfectivo"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtBilleteImporte"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Shape2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Cheque Recibido"
      TabPicture(2)   =   "frmCajaGenRegPAdeudEX.frx":00A4
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Shape5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fgChqRecibido"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtChqRecImporte"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdChqRecibido"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Otros "
      TabPicture(3)   =   "frmCajaGenRegPAdeudEX.frx":00C0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtTotalOtrasCtas"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "fgOtros"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdAgregarCta"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdEliminarCta"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "fgObj"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label18"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Shape3"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      Begin VB.CommandButton cmdChqRecibido 
         Caption         =   "&Cheques "
         Height          =   315
         Left            =   210
         TabIndex        =   49
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
         Left            =   -74580
         TabIndex        =   48
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
         Left            =   6930
         TabIndex        =   25
         Tag             =   "0"
         Text            =   "0.00"
         Top             =   2565
         Width           =   1680
      End
      Begin Sicmact.FlexEdit fgChqRecibido 
         Height          =   1965
         Left            =   210
         TabIndex        =   24
         Top             =   540
         Width           =   8415
         _extentx        =   14843
         _extenty        =   3466
         cols0           =   12
         encabezadosnombres=   "#-Opc-Banco-NroCheque-Fecha-Importe-Cuenta-cAreaCod-cAgeCod-nMovNro-cPersCod-cIFTpo"
         encabezadosanchos=   "0-420-3200-1800-1200-1500-0-0-0-0-0-0"
         font            =   "frmCajaGenRegPAdeudEX.frx":00DC
         font            =   "frmCajaGenRegPAdeudEX.frx":0100
         font            =   "frmCajaGenRegPAdeudEX.frx":0124
         font            =   "frmCajaGenRegPAdeudEX.frx":0148
         font            =   "frmCajaGenRegPAdeudEX.frx":016C
         fontfixed       =   "frmCajaGenRegPAdeudEX.frx":0190
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
         TabIndex        =   29
         Tag             =   "3"
         Top             =   1545
         Width           =   1440
      End
      Begin Sicmact.FlexEdit fgOtros 
         Height          =   1005
         Left            =   -74790
         TabIndex        =   26
         Top             =   510
         Width           =   7245
         _extentx        =   12779
         _extenty        =   1667
         cols0           =   4
         highlight       =   1
         encabezadosnombres=   "#-Cuenta-Descripcion-Importe"
         encabezadosanchos=   "300-1800-3500-1300"
         font            =   "frmCajaGenRegPAdeudEX.frx":01B6
         font            =   "frmCajaGenRegPAdeudEX.frx":01DA
         font            =   "frmCajaGenRegPAdeudEX.frx":01FE
         font            =   "frmCajaGenRegPAdeudEX.frx":0222
         font            =   "frmCajaGenRegPAdeudEX.frx":0246
         fontfixed       =   "frmCajaGenRegPAdeudEX.frx":026A
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
         TabIndex        =   27
         Top             =   510
         Width           =   1200
      End
      Begin VB.CommandButton cmdEliminarCta 
         Caption         =   "&Eliminar"
         Height          =   360
         Left            =   -67470
         TabIndex        =   28
         Top             =   885
         Width           =   1200
      End
      Begin VB.CommandButton cmdEfectivo 
         Caption         =   "Descomposición de Efectivo"
         Height          =   405
         Left            =   -69390
         TabIndex        =   22
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
         TabIndex        =   23
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
         Left            =   -74730
         TabIndex        =   38
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
            TabIndex        =   46
            Tag             =   "0"
            Top             =   1185
            Width           =   1680
         End
         Begin Sicmact.TxtBuscar txtBuscaEntidad 
            Height          =   360
            Left            =   1095
            TabIndex        =   17
            Top             =   300
            Width           =   2580
            _extentx        =   4551
            _extenty        =   635
            appearance      =   1
            appearance      =   1
            font            =   "frmCajaGenRegPAdeudEX.frx":0290
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
            TabIndex        =   47
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
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   300
            Width           =   4290
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
         Left            =   -74730
         TabIndex        =   37
         Top             =   2130
         Width           =   4560
         Begin VB.OptionButton optDoc 
            Caption         =   "Carta"
            Height          =   345
            Index           =   1
            Left            =   2310
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   270
            Width           =   2055
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Cheque"
            Height          =   345
            Index           =   0
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   270
            Width           =   2055
         End
      End
      Begin Sicmact.FlexEdit fgObj 
         Height          =   945
         Left            =   -74790
         TabIndex        =   30
         Top             =   1950
         Width           =   6195
         _extentx        =   10927
         _extenty        =   1667
         cols0           =   8
         highlight       =   2
         allowuserresizing=   1
         encabezadosnombres=   "#-Ord-Código-Descripción-CtaCont-SubCta-ObjPadre-ItemCtaCont"
         encabezadosanchos=   "350-400-1200-3000-0-900-0-0"
         font            =   "frmCajaGenRegPAdeudEX.frx":02B4
         font            =   "frmCajaGenRegPAdeudEX.frx":02D8
         font            =   "frmCajaGenRegPAdeudEX.frx":02FC
         font            =   "frmCajaGenRegPAdeudEX.frx":0320
         font            =   "frmCajaGenRegPAdeudEX.frx":0344
         fontfixed       =   "frmCajaGenRegPAdeudEX.frx":0368
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
         Left            =   5790
         TabIndex        =   44
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         Left            =   5580
         Top             =   2550
         Width           =   3045
      End
   End
End
Attribute VB_Name = "frmCajaGenRegPAdeudEX"
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

'-Graba Cuentas de Orden por Importe Total de Pagare
Dim lsCtaOrdenD As String
Dim lsCtaOrdenH As String

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
    MsgBox "No se indicó Monto del Préstamo. Por favor Verifique", vbInformation, "Aviso"
    txtCapital.SetFocus
    Exit Function
End If

If nVal(txtImporte) = 0 Then
    MsgBox "No se indico Importe Origen para Apertura", vbInformation, "¡Aviso!"
    Exit Function
End If

If lnTipoCtaIf = gTpoCtaIFCtaAdeud Then
    If Not lbCalendario Then
        MsgBox "No se ha definido calendario de pagos para pagaré respectivo", vbInformation, "Aviso"
        Me.cmdCalendario.SetFocus
        Exit Function
    End If
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

Private Sub cmdAceptar_Click()
Dim oDocPago As clsDocPago
Dim lnTpoCuota As CGAdeudCalTpoCuota
Dim lsObjetoPadre As String
Dim lsObjeto     As String
Dim lsPersNombre As String
Dim lsPersDireccion As String
Dim lsUbigeo As String
Dim lsMovNro As String
Dim oCon     As NContFunciones
Dim oCaja As nCajaGeneral
Dim rsAdeud  As ADODB.Recordset

If Valida = False Then
   Exit Sub
End If
If Len(Trim(lsNroCartaApert)) = 0 And cmdCartAper.Visible Then
    If MsgBox("Carta de Apertura no ha sido ingresada. ¿Desea Continuar con la Operación? ", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        cmdCartAper.SetFocus
        Exit Sub
    End If
End If

Set oDocPago = New clsDocPago
Set oCon = New NContFunciones
Set oCaja = New nCajaGeneral
If lnTipoCtaIf = gTpoCtaIFCtaAdeud Then
   If txtCapital <> frmAdeudCal.nCapital Then
      MsgBox "Capital de Adeudado debe ser igual al definido en el Calendario. Por favor verificar", vbInformation, "¡Aviso!"
      Exit Sub
   End If
   If frmAdeudCal.fgCronograma.TextMatrix(1, 1) = "" Then
      MsgBox "No se genero Calendario de Pagare", vbInformation, "¡Aviso!"
      Exit Sub
   End If
   Set rsAdeud = frmAdeudCal.fgCronograma.GetRsNew(1)
   If frmAdeudCal.optTpoCuota(0).value Then
      lnTpoCuota = gAdeudTpoCuotaFija
   Else
      lnTpoCuota = gAdeudTpoCuotaVariable
   End If
Else
    Set rsAdeud = Nothing
End If

If MsgBox("Desea Grabar Operación de Apertura??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    cmdAceptar.Enabled = False
    lsMovNro = oCon.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
    oCaja.GrabaRegPagareAdeudado lsMovNro, gsOpeCod, txtFecha, txtMovDesc, nVal(txtCapital), frmAdeudCal.txtCapital, _
           Me.txtCuentaCod, txtCuentaDesc, lbCreaSubCta, lsSubCtaIF, lblDescTipoCta, _
            Mid(txtBuscaIF, 4, 13), Mid(txtBuscaIF, 1, 2), rsBill, rsMon, _
            lsNroCartaApert, gdFecSis, txtBuscaEntidad, txtBancoImporte, _
            lnTpoDoc, lsNroDoc, ldFechaDoc, lsNroVoucher, _
            fgChqRecibido.GetRsNew, fgOtros.GetRsNew, fgObj.GetRsNew, rsAdeud, _
            frmAdeudCal.SpnGracia.Valor, txtComisionInicial, nVal(txtComisionMonto), Val(frmAdeudCal.chkInterno.value), frmAdeudCal.txtCuotaPagoK, lnNroCuotas, frmAdeudCal.txtPlazoCuotas, IIf(frmAdeudCal.optPeriodo(0).value, 360, 30), frmAdeudCal.txtInteres, lnTpoCuota, CCur(frmAdeudCal.txtTramo), IIf(frmAdeudCal.chkVac = vbChecked And Mid(gsOpeCod, 3, 1) = "1", gMonedaExtranjera, Mid(gsOpeCod, 3, 1)), _
            nVal(frmAdeudCal.txtComision), txtLinCredCod, frmAdeudCal.txtFechaCuota _
            , Me.txtPorcentAfect.value, lsCtaOrdenD, lsCtaOrdenH

    If lsDocCartaAper <> "" Then
       EnviaPrevio lsDocCartaAper + oImpresora.gPrnSaltoPagina + lsDocCartaAper, "Cartas de Apertura", gnLinPage, False
    End If
    ImprimeAsientoContable lsMovNro, lsNroVoucher, lnTpoDoc, lsDocumento, False, False
    
    If MsgBox("¿ Desea Realizar otra operación de Apertura de Cuentas ?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        txtBuscaEntidad = ""
        txtBuscaIF = ""
        txtCuentaCod = ""
        txtCuentaDesc = ""
        
        txtComisionInicial = "0.00"
        txtComisionMonto = "0.00"
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
        
        fgOtros.Clear
        fgOtros.Rows = 2
        fgOtros.FormaCabecera
        
        fgObj.Clear
        fgObj.Rows = 2
        fgObj.FormaCabecera
        
        lbCalendario = False
        Set rsAdeud = Nothing
        If Not frmAdeudCal Is Nothing Then
            Unload frmAdeudCal
            Set frmAdeudCal = Nothing
        End If
        cmdAceptar.Enabled = True
    Else
        If Not frmAdeudCal Is Nothing Then
            Unload frmAdeudCal
            Set frmAdeudCal = Nothing
        End If
        Unload Me
    End If
    
End If
End Sub

Private Sub cmdAgregarCta_Click()
Dim oOpe As New DOperacion
fgOtros.AdicionaFila
fgOtros.rsTextBuscar = oOpe.EmiteOpeCtasNivel(gsOpeCod, , "4")
fgOtros.SetFocus
Set oOpe = Nothing
End Sub

Private Sub cmdCalendario_Click()
If txtBuscaIF <> "" And Val(txtTotApertura) > 0 Then
    frmAdeudCal.Inicio True, txtBuscaIF + "." + txtCuentaCod, lblDescIF, nVal(txtCapital), txtFecha
    If frmAdeudCal.OK Then
        lbCalendario = True
        lnNroCuotas = frmAdeudCal.SpnCuotas.Valor
    Else
        lbCalendario = False
        Set frmAdeudCal = Nothing
    End If
Else
    MsgBox "Debe indicar Institución e Importe del Adeudado", vbInformation, "¡Aviso!"
End If
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

Private Function ValidaDatosCuenta(pnRow As Integer) As Boolean
ValidaDatosCuenta = False
If txtCuentaCod = "" Then
    MsgBox "Falta indicar Codigo de Cuenta a aperturar", vbInformation, "¡AViso!"
    txtCuentaCod.SetFocus
    Exit Function
End If
If txtCuentaDesc = "" Then
    MsgBox "Falta indicar Descripción de Cuenta a Aperturar", vbInformation, "¡Aviso!"
    txtCuentaDesc.SetFocus
    Exit Function
End If
If nVal(txtCapital) = 0 Then
    MsgBox "Falta indicar Importe de Apertura", vbInformation, "¡Aviso!"
    txtCapital.SetFocus
    Exit Function
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

Private Sub Form_Load()
Set oOpe = New DOperacion
Set oCtaIf = New NCajaCtaIF
lbCreaSubCta = False
Me.Caption = gsOpeDesc
txtBuscaIF.psRaiz = "Instituciones Financieras"
txtBuscaIF.rs = oOpe.GetOpeObj(gsOpeCod, "1")

txtBuscaEntidad.psRaiz = "Cuentas de Entidades Financieras"
txtBuscaEntidad.rs = oOpe.GetOpeObj(gsOpeCod, "2")
'txtCtaNro = "Pendiente"

lblDescTipoCta = gsOpeDescHijo
CentraForm Me
txtFecha = gdFecSis
lnTipoCtaIf = -1
Select Case gsOpeCod
    Case gOpeCGAdeudaRegPagareMN, gOpeCGAdeudaRegPagareME
        lnTipoCtaIf = gTpoCtaIFCtaAdeud
        cmdCalendario.Visible = True
End Select
lbCalendario = False
Me.cmdCartAper.Visible = False

lsCtaOrdenD = oOpe.EmiteOpeCta(gsOpeCod, "D", 4)
lsCtaOrdenH = oOpe.EmiteOpeCta(gsOpeCod, "H", 4)

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
                     CCur(txtImporte), gdFecSis, gsNomCmacRUC, txtBuscaEntidad, lblDescIfTransf, _
                     lblDesCtaIfTransf, "", True, , Mid(txtBuscaEntidad, 18, 10)
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
txtBancoImporte.SetFocus
End Sub
Private Sub txtBuscaIF_EmiteDatos()
Dim oAdeud As New NCajaAdeudados

lblDescIF = txtBuscaIF.psDescripcion
lsSubCtaIF = ""
lbCreaSubCta = False
If txtBuscaIF <> "" Then
    txtLinCredCod = ""
    lblLinCredDesc = ""
    txtLinCredCod.rs = oAdeud.GetLineaCredito(Mid(gsOpeCod, 3, 1), Mid(txtBuscaIF, 4, 13))
    lbCreaSubCta = Not oCtaIf.GetVerificaSubCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1))
    lsSubCtaIF = oCtaIf.GetSubCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1))
    txtCuentaCod = oCtaIf.GetNewCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1), lsSubCtaIF)
    txtCuentaDesc = "Pendiente"
    txtCuentaDesc.SetFocus
End If
Set oAdeud = Nothing
End Sub

Private Sub txtcapital_GotFocus()
fEnfoque txtCapital
End Sub

Private Sub txtcapital_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCapital, KeyAscii, 15, 2)
If KeyAscii = 13 Then
    txtCapital = Format(txtCapital, gsFormatoNumeroView)
    txtTotApertura = Format(nVal(txtCapital) - nVal(txtComisionMonto), gsFormatoNumeroView)
    If txtLinCredCod.Enabled Then
        txtLinCredCod.SetFocus
    End If
End If
End Sub

Private Sub txtcapital_LostFocus()
If lbCalendario Then
   If txtCapital <> frmAdeudCal.txtCapital Then
      lbCalendario = False
      Unload frmAdeudCal
      Set frmAdeudCal = Nothing
   End If
End If

End Sub

Private Sub txtCapital_Validate(Cancel As Boolean)
txtTotApertura = Format(nVal(txtCapital) - nVal(txtComisionMonto), gsFormatoNumeroView)
End Sub

Private Sub txtComisionInicial_GotFocus()
fEnfoque txtComisionInicial
End Sub

Private Sub txtComisionInicial_KeyPress(KeyAscii As Integer)
If nVal(txtCapital) = 0 Then
    MsgBox "Primero ingresar Monto de Prestamo", vbInformation, "¡Aviso!"
    txtCapital.SetFocus
    Exit Sub
End If
KeyAscii = NumerosDecimales(txtComisionInicial, KeyAscii, 10, 4)
If KeyAscii = 13 Then
   txtComisionMonto = Format(Round(nVal(txtCapital) * txtComisionInicial / 100, 2), gsFormatoNumeroView)
   txtComisionInicial = Format(txtComisionInicial, gsFormatoNumeroView)
   txtTotApertura = Format(nVal(txtCapital) - nVal(txtComisionMonto), gsFormatoNumeroView)
   txtComisionMonto.SetFocus
End If
End Sub

Private Sub txtComisionMonto_GotFocus()
fEnfoque txtComisionMonto
End Sub

Private Sub txtComisionMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtComisionMonto, KeyAscii, 14, 4)
If KeyAscii = 13 Then
   txtTotApertura = Format(nVal(txtCapital) - nVal(txtComisionMonto), gsFormatoNumeroView)
   txtMovDesc.SetFocus
End If
End Sub

Private Sub txtComisionMonto_Validate(Cancel As Boolean)
txtTotApertura = Format(nVal(txtCapital) - nVal(txtComisionMonto), gsFormatoNumeroView)
End Sub

Private Sub txtCuentaCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCuentaDesc.SetFocus
End If
End Sub

Private Sub txtCuentaDesc_GotFocus()
fEnfoque txtCuentaDesc
End Sub

Private Sub txtCuentaDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPorcentAfect.SetFocus
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

Private Sub txtLinCredCod_EmiteDatos()
lblLinCredDesc = txtLinCredCod.psDescripcion
txtComisionInicial.SetFocus
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
    'fgOtros.TextMatrix(fgOtros.Row, 6) = psObjetoCod
    
End Sub

Private Sub EliminaCuenta(sCod As String, nItem As Integer)
If fgOtros.TextMatrix(1, 0) <> "" Then
    EliminaFgObj Val(fgOtros.TextMatrix(fgOtros.Row, 0))
    fgOtros.EliminaFila fgOtros.Row, False
End If
If Len(fgOtros.TextMatrix(1, 1)) > 0 Then
   RefrescaFgObj Val(fgOtros.TextMatrix(fgOtros.Row, 0))
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

Private Sub txtPorcentAfect_GotFocus()
    txtPorcentAfect.SelStart = 0
    txtPorcentAfect.SelLength = 50
End Sub

Private Sub txtPorcentAfect_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCapital.SetFocus
End If
End Sub

