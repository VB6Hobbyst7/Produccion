VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredFormEvalFormato5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Créditos - Evaluación - Formato 5"
   ClientHeight    =   10410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13485
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredFormEvalFormato5.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMNME 
      Caption         =   "MN - ME"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   11040
      TabIndex        =   139
      Top             =   400
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Hoja Evaluación"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11040
      TabIndex        =   37
      Top             =   740
      Width           =   2370
   End
   Begin VB.CommandButton cmdInformeVisita 
      Caption         =   "Infor&me de Visita"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   11040
      TabIndex        =   35
      Top             =   1100
      Width           =   2370
   End
   Begin VB.CommandButton cmdVerCar 
      Caption         =   "&Ver CAR"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   11040
      TabIndex        =   36
      Top             =   1440
      Width           =   2370
   End
   Begin VB.CommandButton cmdFlujoCaja5 
      Caption         =   "Generar &Flujo Caja"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   11040
      TabIndex        =   38
      Top             =   1790
      Width           =   2370
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   11040
      TabIndex        =   33
      Top             =   40
      Width           =   2370
   End
   Begin TabDlg.SSTab SSTabRatios 
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   9600
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   1561
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Ratios e Indicadores"
      TabPicture(0)   =   "frmCredFormEvalFormato5.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label32"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label33"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblCapaAceptable"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblEndeAceptable"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblRentabilidad"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblLiquidez"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblEndeudamiento"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtRentabilidadPat"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtLiquidezCte"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtExcedenteMensual"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtIngresoNeto"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtEndeudamiento"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCapacidadNeta"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Datos Flujo Caja Proyectada"
      TabPicture(1)   =   "frmCredFormEvalFormato5.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "EditMoneyForm5IVC"
      Tab(1).Control(1)=   "EditMoneyForm5ICM"
      Tab(1).Control(2)=   "EditMoneyForm5IPP"
      Tab(1).Control(3)=   "EditMoneyForm5IGV"
      Tab(1).Control(4)=   "EditMoneyForm5IC"
      Tab(1).Control(5)=   "Label13(12)"
      Tab(1).Control(6)=   "Label13(11)"
      Tab(1).Control(7)=   "Label13(10)"
      Tab(1).Control(8)=   "Label13(9)"
      Tab(1).Control(9)=   "Label13(8)"
      Tab(1).Control(10)=   "Label34"
      Tab(1).Control(11)=   "Label21"
      Tab(1).Control(12)=   "Label22"
      Tab(1).Control(13)=   "Label28"
      Tab(1).Control(14)=   "Label29"
      Tab(1).Control(15)=   "Label13(7)"
      Tab(1).Control(16)=   "Label13(6)"
      Tab(1).Control(17)=   "Label13(5)"
      Tab(1).Control(18)=   "Label13(4)"
      Tab(1).Control(19)=   "Label13(3)"
      Tab(1).ControlCount=   20
      Begin SICMACT.EditMoney txtCapacidadNeta 
         Height          =   300
         Left            =   1460
         TabIndex        =   39
         Top             =   340
         Width           =   850
         _extentx        =   1508
         _extenty        =   529
         font            =   "frmCredFormEvalFormato5.frx":0342
         forecolor       =   8421504
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtEndeudamiento 
         Height          =   300
         Left            =   3640
         TabIndex        =   40
         Top             =   345
         Width           =   850
         _extentx        =   1508
         _extenty        =   529
         font            =   "frmCredFormEvalFormato5.frx":036A
         forecolor       =   8421504
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtIngresoNeto 
         Height          =   300
         Left            =   10305
         TabIndex        =   41
         Top             =   390
         Width           =   1095
         _extentx        =   1931
         _extenty        =   529
         font            =   "frmCredFormEvalFormato5.frx":0392
         forecolor       =   8421504
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtExcedenteMensual 
         Height          =   300
         Left            =   12300
         TabIndex        =   42
         Top             =   390
         Width           =   1095
         _extentx        =   1931
         _extenty        =   529
         font            =   "frmCredFormEvalFormato5.frx":03BA
         forecolor       =   8421504
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtLiquidezCte 
         Height          =   300
         Left            =   8090
         TabIndex        =   43
         Top             =   390
         Width           =   1095
         _extentx        =   1931
         _extenty        =   529
         font            =   "frmCredFormEvalFormato5.frx":03E2
         forecolor       =   8421504
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtRentabilidadPat 
         Height          =   300
         Left            =   5860
         TabIndex        =   44
         Top             =   345
         Width           =   850
         _extentx        =   1508
         _extenty        =   529
         font            =   "frmCredFormEvalFormato5.frx":040A
         forecolor       =   8421504
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyForm5IVC 
         Height          =   300
         Left            =   -74880
         TabIndex        =   119
         ToolTipText     =   "Incremento de ventas al contado - Anual"
         Top             =   520
         Width           =   1305
         _extentx        =   1508
         _extenty        =   529
         font            =   "frmCredFormEvalFormato5.frx":0432
         forecolor       =   -2147483640
         text            =   "00.0"
         enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyForm5ICM 
         Height          =   300
         Left            =   -72600
         TabIndex        =   120
         ToolTipText     =   "Incremento de Compras de Mercaderias - Anual"
         Top             =   525
         Width           =   1305
         _extentx        =   1508
         _extenty        =   529
         font            =   "frmCredFormEvalFormato5.frx":045A
         forecolor       =   -2147483640
         text            =   "00.0"
         enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyForm5IPP 
         Height          =   300
         Left            =   -70440
         TabIndex        =   121
         ToolTipText     =   "Incremento de Consumo - Anual"
         Top             =   520
         Width           =   1305
         _extentx        =   1508
         _extenty        =   529
         font            =   "frmCredFormEvalFormato5.frx":0482
         forecolor       =   -2147483640
         text            =   "00.0"
         enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyForm5IGV 
         Height          =   300
         Left            =   -68160
         TabIndex        =   122
         ToolTipText     =   "Incremento de Pago Personal -Anual"
         Top             =   525
         Width           =   1305
         _extentx        =   1508
         _extenty        =   529
         font            =   "frmCredFormEvalFormato5.frx":04AA
         forecolor       =   -2147483640
         text            =   "00.0"
         enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyForm5IC 
         Height          =   300
         Left            =   -63840
         TabIndex        =   123
         ToolTipText     =   "Incremento de Gastos de Ventas - Anual"
         Top             =   525
         Width           =   1305
         _extentx        =   1508
         _extenty        =   529
         font            =   "frmCredFormEvalFormato5.frx":04D2
         forecolor       =   8421504
         text            =   "00.0"
      End
      Begin VB.Label lblEndeudamiento 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endeudamiento:"
         Height          =   195
         Left            =   2400
         TabIndex        =   142
         Top             =   405
         Width           =   1170
      End
      Begin VB.Label lblLiquidez 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidez Cte:"
         Height          =   195
         Left            =   6960
         TabIndex        =   141
         Top             =   405
         Width           =   930
      End
      Begin VB.Label lblRentabilidad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rentabilidad Pat.:"
         Height          =   195
         Left            =   4560
         TabIndex        =   140
         Top             =   405
         Width           =   1290
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   12
         Left            =   -62520
         TabIndex        =   138
         Top             =   555
         Width           =   165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   11
         Left            =   -66840
         TabIndex        =   137
         Top             =   550
         Width           =   165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   10
         Left            =   -69120
         TabIndex        =   136
         Top             =   550
         Width           =   165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   9
         Left            =   -71280
         TabIndex        =   135
         Top             =   550
         Width           =   165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   8
         Left            =   -73560
         TabIndex        =   134
         Top             =   550
         Width           =   165
      End
      Begin VB.Label Label34 
         Caption         =   "Anual"
         Height          =   255
         Left            =   -73320
         TabIndex        =   133
         Top             =   550
         Width           =   495
      End
      Begin VB.Label Label21 
         Caption         =   "Anual"
         Height          =   255
         Left            =   -71040
         TabIndex        =   132
         Top             =   550
         Width           =   495
      End
      Begin VB.Label Label22 
         Caption         =   "Anual"
         Height          =   255
         Left            =   -68880
         TabIndex        =   131
         Top             =   550
         Width           =   495
      End
      Begin VB.Label Label28 
         Caption         =   "Anual"
         Height          =   255
         Left            =   -66600
         TabIndex        =   130
         Top             =   550
         Width           =   495
      End
      Begin VB.Label Label29 
         Caption         =   "Anual"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -62280
         TabIndex        =   129
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. Gasto Ventas:"
         Height          =   195
         Index           =   7
         Left            =   -68160
         TabIndex        =   128
         Top             =   320
         Width           =   1410
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. Pago Personal:"
         Height          =   195
         Index           =   6
         Left            =   -70440
         TabIndex        =   127
         Top             =   320
         Width           =   1470
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. de Consumo:"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   5
         Left            =   -63840
         TabIndex        =   126
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. compra mercaderias:"
         Height          =   195
         Index           =   4
         Left            =   -72600
         TabIndex        =   125
         Top             =   320
         Width           =   1890
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incremento ventas contado:"
         Height          =   195
         Index           =   3
         Left            =   -74880
         TabIndex        =   124
         Top             =   320
         Width           =   2055
      End
      Begin VB.Label lblEndeAceptable 
         Caption         =   "Aceptable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3720
         TabIndex        =   118
         Top             =   660
         Width           =   855
      End
      Begin VB.Label lblCapaAceptable 
         Caption         =   "Aceptable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1480
         TabIndex        =   117
         Top             =   660
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   6720
         X2              =   6720
         Y1              =   360
         Y2              =   840
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excedente:"
         Height          =   195
         Left            =   11475
         TabIndex        =   47
         Top             =   435
         Width           =   825
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso Neto:"
         Height          =   195
         Left            =   9315
         TabIndex        =   46
         Top             =   435
         Width           =   1005
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacidad de Pago:"
         Height          =   195
         Index           =   2
         Left            =   50
         TabIndex        =   45
         Top             =   400
         Width           =   1440
      End
   End
   Begin TabDlg.SSTab SSTabInfoNego 
      Height          =   2115
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   10850
      _ExtentX        =   19129
      _ExtentY        =   3731
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Información del Negocio"
      TabPicture(0)   =   "frmCredFormEvalFormato5.frx":04FA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtFechaEvaluacion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ActXCodCta"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtGiroNeg"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.Frame Frame1 
         Height          =   1420
         Left            =   120
         TabIndex        =   49
         Top             =   650
         Width           =   10660
         Begin VB.TextBox txtNombreCliente 
            Height          =   300
            Left            =   1800
            TabIndex        =   4
            Top             =   120
            Width           =   4155
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Propia"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   11
            Top             =   1100
            Width           =   855
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Alquilada"
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   12
            Top             =   1100
            Width           =   1095
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Ambulante"
            Height          =   255
            Index           =   3
            Left            =   3720
            TabIndex        =   13
            Top             =   1100
            Width           =   1095
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Otros"
            Height          =   255
            Index           =   4
            Left            =   4920
            TabIndex        =   14
            Top             =   1100
            Width           =   855
         End
         Begin VB.TextBox txtCondLocalOtros 
            Height          =   300
            Left            =   5820
            MaxLength       =   250
            TabIndex        =   15
            Top             =   1080
            Visible         =   0   'False
            Width           =   4755
         End
         Begin MSMask.MaskEdBox txtFecUltEndeuda 
            Height          =   300
            Left            =   9240
            TabIndex        =   9
            Top             =   120
            Width           =   1210
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Spinner.uSpinner spnTiempoLocalAnio 
            Height          =   300
            Left            =   1800
            TabIndex        =   0
            Top             =   780
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            Max             =   99
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin Spinner.uSpinner spnTiempoLocalMes 
            Height          =   300
            Left            =   3360
            TabIndex        =   10
            Top             =   780
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            Max             =   12
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin SICMACT.EditMoney txtExposicionCredito 
            Height          =   300
            Left            =   9120
            TabIndex        =   5
            Top             =   760
            Width           =   1335
            _extentx        =   2355
            _extenty        =   529
            font            =   "frmCredFormEvalFormato5.frx":0516
            backcolor       =   -2147483643
            forecolor       =   8421504
            text            =   "0"
            enabled         =   -1  'True
         End
         Begin Spinner.uSpinner spnExpEmpAnio 
            Height          =   300
            Left            =   1800
            TabIndex        =   6
            Top             =   450
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            Max             =   99
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   8421504
         End
         Begin Spinner.uSpinner spnExpEmpMes 
            Height          =   300
            Left            =   3360
            TabIndex        =   7
            Top             =   450
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            Max             =   12
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   8421504
         End
         Begin SICMACT.EditMoney txtUltEndeuda 
            Height          =   300
            Left            =   9120
            TabIndex        =   8
            Top             =   450
            Width           =   1335
            _extentx        =   2355
            _extenty        =   529
            font            =   "frmCredFormEvalFormato5.frx":053E
            backcolor       =   -2147483643
            forecolor       =   8421504
            text            =   "0"
            enabled         =   -1  'True
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   1200
            TabIndex        =   60
            Top             =   150
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exp. como empresario:"
            Height          =   195
            Left            =   75
            TabIndex        =   59
            Top             =   460
            Width           =   1650
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tiempo en el local:"
            Height          =   195
            Left            =   420
            TabIndex        =   58
            Top             =   795
            Width           =   1320
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condición local:"
            Height          =   195
            Left            =   600
            TabIndex        =   57
            Top             =   1095
            Width           =   1110
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exposición con este crédito:"
            Height          =   195
            Left            =   7060
            TabIndex        =   56
            Top             =   825
            Width           =   2010
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "años"
            Height          =   255
            Left            =   2595
            TabIndex        =   55
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "años"
            Height          =   255
            Left            =   2595
            TabIndex        =   54
            Top             =   795
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "meses"
            Height          =   255
            Left            =   4155
            TabIndex        =   53
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "meses"
            Height          =   255
            Left            =   4155
            TabIndex        =   52
            Top             =   795
            Width           =   615
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Último endeudamiento RCC:"
            Height          =   195
            Left            =   7060
            TabIndex        =   51
            Top             =   480
            Width           =   2010
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Ult. endeudamiento RCC:"
            Height          =   195
            Left            =   6760
            TabIndex        =   50
            Top             =   150
            Width           =   2310
         End
      End
      Begin VB.TextBox txtGiroNeg 
         Height          =   300
         Left            =   5640
         TabIndex        =   3
         Top             =   360
         Width           =   5115
      End
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   320
         Width           =   3735
         _extentx        =   6588
         _extenty        =   661
         texto           =   "Crédito"
      End
      Begin MSMask.MaskEdBox txtFechaEvaluacion 
         Height          =   300
         Left            =   9360
         TabIndex        =   144
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de evaluación al:"
         Height          =   195
         Left            =   7520
         TabIndex        =   143
         Top             =   45
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Giro del Negocio :"
         Height          =   255
         Left            =   4320
         TabIndex        =   61
         Top             =   375
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab SSTabIngresos5 
      Height          =   7500
      Left            =   0
      TabIndex        =   62
      Top             =   2115
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   13229
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Activos y Pasivos"
      TabPicture(0)   =   "frmCredFormEvalFormato5.frx":0566
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Evaluación"
      TabPicture(1)   =   "frmCredFormEvalFormato5.frx":0582
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Line2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame13"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame14"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Propuesta del Crédito"
      TabPicture(2)   =   "frmCredFormEvalFormato5.frx":059E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "framePropuesta"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Comentarios y Referidos"
      TabPicture(3)   =   "frmCredFormEvalFormato5.frx":05BA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frameReferido"
      Tab(3).Control(1)=   "frameComentario"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame14 
         Caption         =   "Otros Ingresos :"
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
         Height          =   2175
         Left            =   6720
         TabIndex        =   116
         Top             =   3600
         Width           =   6135
         Begin SICMACT.FlexEdit feOtrosIngresos 
            Height          =   1815
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   5715
            _extentx        =   10081
            _extenty        =   3201
            cols0           =   5
            highlight       =   1
            encabezadosnombres=   "-N-Concepto-Monto-Aux"
            encabezadosanchos=   "0-300-3500-1800-0"
            font            =   "frmCredFormEvalFormato5.frx":05D6
            font            =   "frmCredFormEvalFormato5.frx":05FE
            font            =   "frmCredFormEvalFormato5.frx":0626
            font            =   "frmCredFormEvalFormato5.frx":064E
            font            =   "frmCredFormEvalFormato5.frx":0676
            fontfixed       =   "frmCredFormEvalFormato5.frx":069E
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-3-X"
            listacontroles  =   "0-0-0-0-0"
            encabezadosalineacion=   "C-C-L-R-C"
            formatosedit    =   "0-0-0-2-0"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            rowheight0      =   300
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Gastos Familiares : "
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
         Height          =   3135
         Left            =   6720
         TabIndex        =   115
         Top             =   360
         Width           =   6135
         Begin SICMACT.FlexEdit feGastosFamiliares 
            Height          =   2775
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   5760
            _extentx        =   10160
            _extenty        =   4895
            cols0           =   5
            highlight       =   1
            encabezadosnombres=   "-N-Concepto-Monto-Aux"
            encabezadosanchos=   "0-300-3500-1800-0"
            font            =   "frmCredFormEvalFormato5.frx":06C4
            font            =   "frmCredFormEvalFormato5.frx":06EC
            font            =   "frmCredFormEvalFormato5.frx":0714
            font            =   "frmCredFormEvalFormato5.frx":073C
            font            =   "frmCredFormEvalFormato5.frx":0764
            fontfixed       =   "frmCredFormEvalFormato5.frx":078C
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            tipobusqueda    =   6
            columnasaeditar =   "X-X-X-3-X"
            listacontroles  =   "0-0-0-0-0"
            encabezadosalineacion=   "C-C-L-R-C"
            formatosedit    =   "0-0-0-2-0"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            rowheight0      =   300
         End
      End
      Begin VB.Frame frameReferido 
         Caption         =   "Referidos :"
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
         Height          =   3255
         Left            =   -74640
         TabIndex        =   113
         Top             =   3120
         Width           =   12735
         Begin VB.CommandButton cmdAgregarRef 
            Caption         =   "&Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   9840
            TabIndex        =   31
            Top             =   2760
            Width           =   1170
         End
         Begin VB.CommandButton cmdQuitar5 
            Caption         =   "&Quitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   11160
            TabIndex        =   32
            Top             =   2760
            Width           =   1170
         End
         Begin SICMACT.FlexEdit feReferidos 
            Height          =   2415
            Left            =   240
            TabIndex        =   30
            Top             =   240
            Width           =   12075
            _extentx        =   21299
            _extenty        =   4260
            cols0           =   7
            highlight       =   1
            encabezadosnombres=   "N-Nombres-DNI-Teléfono-Comentario-NroDNI-Aux"
            encabezadosanchos=   "350-4250-960-1260-5100-0-0"
            font            =   "frmCredFormEvalFormato5.frx":07B2
            font            =   "frmCredFormEvalFormato5.frx":07DA
            font            =   "frmCredFormEvalFormato5.frx":0802
            font            =   "frmCredFormEvalFormato5.frx":082A
            font            =   "frmCredFormEvalFormato5.frx":0852
            fontfixed       =   "frmCredFormEvalFormato5.frx":087A
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            tipobusqueda    =   3
            columnasaeditar =   "X-1-2-3-4-X-X"
            listacontroles  =   "0-0-0-0-0-0-0"
            encabezadosalineacion=   "L-L-L-L-L-L-C"
            formatosedit    =   "0-0-0-0-0-0-0"
            textarray0      =   "N"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            colwidth0       =   345
            rowheight0      =   300
         End
      End
      Begin VB.Frame frameComentario 
         Caption         =   "Comentarios :"
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
         Height          =   2535
         Left            =   -74640
         TabIndex        =   112
         Top             =   480
         Width           =   12735
         Begin VB.TextBox txtComentario5 
            Height          =   2130
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   3000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   240
            Width           =   12135
         End
      End
      Begin VB.Frame framePropuesta 
         Caption         =   "Propuesta del Credito:"
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
         Height          =   6495
         Left            =   -74640
         TabIndex        =   105
         Top             =   360
         Width           =   12550
         Begin VB.TextBox txtEntornoFamiliar5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   960
            Width           =   12015
         End
         Begin VB.TextBox txtGiroUbicacion5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   1920
            Width           =   12015
         End
         Begin VB.TextBox txtExperiencia5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   2880
            Width           =   12015
         End
         Begin VB.TextBox txtFormalidadNegocio5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   3840
            Width           =   12015
         End
         Begin VB.TextBox txtColaterales5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   4800
            Width           =   12015
         End
         Begin VB.TextBox txtDestino5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   5760
            Width           =   12015
         End
         Begin MSMask.MaskEdBox txtFechaVisita 
            Height          =   300
            Left            =   1440
            TabIndex        =   22
            Top             =   240
            Width           =   1090
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Visita:"
            Height          =   195
            Left            =   240
            TabIndex        =   114
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el entorno familiar del cliente o representante:"
            Height          =   195
            Left            =   360
            TabIndex        =   111
            Top             =   720
            Width           =   3795
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el giro y la ubicación del negocio:"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   110
            Top             =   1680
            Width           =   2820
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sólo la experiencia crediticia:"
            Height          =   195
            Left            =   360
            TabIndex        =   109
            Top             =   2640
            Width           =   2070
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la consistencia de la información y la formalidad del negocio:"
            Height          =   195
            Left            =   360
            TabIndex        =   108
            Top             =   3600
            Width           =   4770
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   360
            TabIndex        =   107
            Top             =   4560
            Width           =   2400
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el destino y el impacto del mismo"
            Height          =   195
            Left            =   360
            TabIndex        =   106
            Top             =   5520
            Width           =   2790
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Declaración PDT:"
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
         Height          =   1420
         Left            =   5160
         TabIndex        =   104
         Top             =   5760
         Width           =   8235
         Begin SICMACT.FlexEdit feDeclaracionPDT 
            Height          =   1095
            Left            =   45
            TabIndex        =   21
            Top             =   240
            Width           =   8120
            _extentx        =   14314
            _extenty        =   1931
            rows            =   3
            cols0           =   9
            fixedcols       =   2
            highlight       =   1
            encabezadosnombres=   "-Mes/Detalle-nConsCod-nConsValor----Promedio-%Declarado"
            encabezadosanchos=   "0-1200-0-0-1350-1350-1350-1500-1230"
            font            =   "frmCredFormEvalFormato5.frx":08A0
            font            =   "frmCredFormEvalFormato5.frx":08C8
            font            =   "frmCredFormEvalFormato5.frx":08F0
            font            =   "frmCredFormEvalFormato5.frx":0918
            font            =   "frmCredFormEvalFormato5.frx":0940
            fontfixed       =   "frmCredFormEvalFormato5.frx":0968
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-X-4-5-6-X-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-R-L-R-R-R-R-C-L"
            formatosedit    =   "0-0-0-0-2-2-2-0-0"
            lbeditarflex    =   -1  'True
            lbformatocol    =   -1  'True
            lbpuntero       =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            rowheight0      =   300
            cellbackcolor   =   -2147483633
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Flujo de Caja Mensual :"
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
         Height          =   7140
         Left            =   50
         TabIndex        =   103
         Top             =   330
         Width           =   5010
         Begin SICMACT.FlexEdit feFlujoCajaMensual 
            Height          =   6930
            Left            =   0
            TabIndex        =   18
            Top             =   200
            Width           =   4950
            _extentx        =   8731
            _extenty        =   12224
            cols0           =   6
            highlight       =   1
            encabezadosnombres=   "N-nConsCod-nConsValor-Concepto-Monto-Aux"
            encabezadosanchos=   "0-0-0-3000-1600-0"
            font            =   "frmCredFormEvalFormato5.frx":098E
            font            =   "frmCredFormEvalFormato5.frx":09B6
            font            =   "frmCredFormEvalFormato5.frx":09DE
            font            =   "frmCredFormEvalFormato5.frx":0A06
            font            =   "frmCredFormEvalFormato5.frx":0A2E
            fontfixed       =   "frmCredFormEvalFormato5.frx":0A56
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            tipobusqueda    =   6
            columnasaeditar =   "X-X-X-X-4-X"
            listacontroles  =   "0-0-0-0-0-0"
            encabezadosalineacion=   "C-L-L-L-R-C"
            formatosedit    =   "0-0-0-0-2-0"
            textarray0      =   "N"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            rowheight0      =   300
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Pasivos :"
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
         Height          =   6620
         Left            =   -68240
         TabIndex        =   102
         Top             =   340
         Width           =   6615
         Begin SICMACT.FlexEdit fePasivos 
            Height          =   6375
            Left            =   60
            TabIndex        =   17
            Top             =   200
            Width           =   6495
            _extentx        =   11456
            _extenty        =   11245
            cols0           =   7
            highlight       =   1
            encabezadosnombres=   "-Concepto-P. P.-P. E.-Total-nConsCod-nConsValor"
            encabezadosanchos=   "0-2630-1150-1150-1450-0-0"
            font            =   "frmCredFormEvalFormato5.frx":0A7C
            font            =   "frmCredFormEvalFormato5.frx":0AA4
            font            =   "frmCredFormEvalFormato5.frx":0ACC
            font            =   "frmCredFormEvalFormato5.frx":0AF4
            font            =   "frmCredFormEvalFormato5.frx":0B1C
            fontfixed       =   "frmCredFormEvalFormato5.frx":0B44
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            tipobusqueda    =   6
            columnasaeditar =   "X-X-2-3-X-X-X"
            listacontroles  =   "0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-L-R-R-R-C-C"
            formatosedit    =   "0-0-2-2-2-2-2"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            rowheight0      =   300
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Activos :"
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
         Height          =   5895
         Left            =   -74880
         TabIndex        =   101
         Top             =   340
         Width           =   6630
         Begin SICMACT.FlexEdit feActivos 
            Height          =   5535
            Left            =   75
            TabIndex        =   16
            Top             =   200
            Width           =   6480
            _extentx        =   11430
            _extenty        =   9763
            cols0           =   7
            highlight       =   1
            encabezadosnombres=   "-Concepto-P. P.-P. E.-Total-nConsCod-nConsValor"
            encabezadosanchos=   "0-2630-1150-1150-1450-0-0"
            font            =   "frmCredFormEvalFormato5.frx":0B6A
            font            =   "frmCredFormEvalFormato5.frx":0B92
            font            =   "frmCredFormEvalFormato5.frx":0BBA
            font            =   "frmCredFormEvalFormato5.frx":0BE2
            font            =   "frmCredFormEvalFormato5.frx":0C0A
            fontfixed       =   "frmCredFormEvalFormato5.frx":0C32
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            tipobusqueda    =   6
            columnasaeditar =   "X-X-2-3-X-X-X"
            listacontroles  =   "0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-L-R-R-R-R-C"
            formatosedit    =   "0-0-2-2-2-3-3"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            rowheight0      =   300
         End
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73200
         TabIndex        =   100
         Top             =   6120
         Width           =   1170
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74640
         TabIndex        =   99
         Top             =   6120
         Width           =   1170
      End
      Begin VB.Frame Frame8 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74760
         TabIndex        =   97
         Top             =   3360
         Width           =   9975
         Begin SICMACT.FlexEdit FlexEdit1 
            Height          =   1935
            Left            =   120
            TabIndex        =   98
            Top             =   360
            Width           =   9720
            _extentx        =   17145
            _extenty        =   3413
            cols0           =   6
            highlight       =   1
            encabezadosnombres=   "N°-Nombre-DNI-Telef.-Referido-DNI"
            encabezadosanchos=   "1000-2800-1000-1500-2300-1000"
            font            =   "frmCredFormEvalFormato5.frx":0C58
            font            =   "frmCredFormEvalFormato5.frx":0C80
            font            =   "frmCredFormEvalFormato5.frx":0CA8
            font            =   "frmCredFormEvalFormato5.frx":0CD0
            font            =   "frmCredFormEvalFormato5.frx":0CF8
            fontfixed       =   "frmCredFormEvalFormato5.frx":0D20
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0-0"
            encabezadosalineacion=   "L-R-L-C-C-C"
            formatosedit    =   "0-2-0-0-0-0"
            textarray0      =   "N°"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            colwidth0       =   1005
            rowheight0      =   300
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74760
         TabIndex        =   95
         Top             =   360
         Width           =   9975
         Begin VB.TextBox Text1 
            Height          =   2010
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   96
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.Frame frmCredEvalFormato1 
         Caption         =   " Gastos del Negocio :"
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
         Height          =   6015
         Left            =   -74880
         TabIndex        =   82
         Top             =   360
         Width           =   9975
         Begin VB.TextBox txtDestino 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   88
            Top             =   5280
            Width           =   9735
         End
         Begin VB.TextBox txtColaterales 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   87
            Top             =   4320
            Width           =   9735
         End
         Begin VB.TextBox txtFormalidadNegocio 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   86
            Top             =   3360
            Width           =   9735
         End
         Begin VB.TextBox txtExperiencia 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   85
            Top             =   2400
            Width           =   9735
         End
         Begin VB.TextBox txtGiroUbicacion 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   84
            Top             =   1440
            Width           =   9735
         End
         Begin VB.TextBox txtEntornoFamiliar 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   83
            Top             =   480
            Width           =   9735
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Top             =   5040
            Width           =   2400
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   4080
            Width           =   2400
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la consistencia de la información y la formalidad del negocio:"
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   3120
            Width           =   4770
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sólo la experiencia crediticia:"
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   2160
            Width           =   2070
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el giro y la ubicación del negocio:"
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   1200
            Width           =   2820
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el entorno familiar del cliente o representante:"
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   3795
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Propuesta del Credito:"
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
         Height          =   6015
         Left            =   -74760
         TabIndex        =   69
         Top             =   480
         Width           =   9975
         Begin VB.TextBox txtDestino2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   75
            Top             =   5280
            Width           =   9735
         End
         Begin VB.TextBox txtColaterales2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   74
            Top             =   4320
            Width           =   9735
         End
         Begin VB.TextBox txtFormalidadNegocio2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   73
            Top             =   3360
            Width           =   9735
         End
         Begin VB.TextBox txtExperiencia2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   72
            Top             =   2400
            Width           =   9735
         End
         Begin VB.TextBox txtGiroUbicacion2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   71
            Top             =   1440
            Width           =   9735
         End
         Begin VB.TextBox txtEntornoFamiliar2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   70
            Top             =   480
            Width           =   9735
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   81
            Top             =   5040
            Width           =   2400
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   4080
            Width           =   2400
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la consistencia de la información y la formalidad del negocio:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   79
            Top             =   3120
            Width           =   4770
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sólo la experiencia crediticia:"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   2160
            Width           =   2070
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el giro y la ubicación del negocio:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   77
            Top             =   1200
            Width           =   2820
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el entorno familiar del cliente o representante:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   3795
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74880
         TabIndex        =   67
         Top             =   360
         Width           =   9975
         Begin VB.TextBox txtComentario2 
            Height          =   2010
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   68
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Referidos :"
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
         Height          =   2895
         Left            =   -74880
         TabIndex        =   65
         Top             =   3240
         Width           =   9975
         Begin SICMACT.FlexEdit feReferidos2 
            Height          =   1935
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   9705
            _extentx        =   17119
            _extenty        =   3413
            cols0           =   7
            highlight       =   1
            encabezadosnombres=   "N°-Nombre-DNI-Telef.-Referido-DNI-Aux"
            encabezadosanchos=   "500-2900-920-1250-2900-920-0"
            font            =   "frmCredFormEvalFormato5.frx":0D46
            font            =   "frmCredFormEvalFormato5.frx":0D6E
            font            =   "frmCredFormEvalFormato5.frx":0D96
            font            =   "frmCredFormEvalFormato5.frx":0DBE
            font            =   "frmCredFormEvalFormato5.frx":0DE6
            fontfixed       =   "frmCredFormEvalFormato5.frx":0E0E
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            tipobusqueda    =   3
            columnasaeditar =   "X-1-2-3-4-5-X"
            listacontroles  =   "0-0-0-0-0-0-0"
            encabezadosalineacion=   "L-L-L-L-C-L-L"
            formatosedit    =   "0-0-0-0-0-0-0"
            textarray0      =   "N°"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            colwidth0       =   495
            rowheight0      =   300
         End
      End
      Begin VB.CommandButton cmdQuitar2 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73200
         TabIndex        =   64
         Top             =   5640
         Width           =   1170
      End
      Begin VB.CommandButton cmdAgregar2 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74640
         TabIndex        =   63
         Top             =   5640
         Width           =   1170
      End
      Begin VB.Line Line2 
         X1              =   6000
         X2              =   6000
         Y1              =   480
         Y2              =   5280
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   11040
      TabIndex        =   34
      Top             =   40
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   13320
      Y1              =   10080
      Y2              =   10080
   End
End
Attribute VB_Name = "frmCredFormEvalFormato5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************************
'*** Nombre      : frmCredFormEvalFormato5                                                   ***
'*** Descripción : Formulario para evaluación de Creditos que tienen el tipo de evaluación 5 ***
'*** Referencia  : ERS004-2016                                                               ***
'*** Creación    : LUCV, 20160525 09:00:00 AM                                                ***
'***********************************************************************************************
Option Explicit
    Dim sCtaCod As String
    Dim sPersCod As String
    Dim gsOpeCod As String
    Dim fnTipoRegMant As Integer
    Dim fnTipoPermiso As Integer
    Dim fbPermiteGrabar As Boolean
    Dim fbBloqueaTodo As Boolean
    Dim fnTotalRefGastoNego As Currency
    Dim fnTotalRefGastoFami As Currency
    Dim fsCliente As String
    Dim fsGiroNego As String
    Dim fsAnioExp As Integer
    Dim fsMesExp As Integer
    Dim fnEstado As Integer
    Dim fnMontoDeudaSbs As Currency
    Dim fnFechaDeudaSbs As Currency
    
    Dim fnCondLocal As Integer
    Dim MatIfiGastoNego As Variant
    Dim MatIfiGastoFami As Variant
    Dim MatReferidos As Variant
    
    Dim MatIfiNoSupervisadaGastoNego As Variant 'CTI320200110 ERS003-2020. Agregó
    Dim MatIfiNoSupervisadaGastoFami As Variant 'CTI320200110 ERS003-2020. Agregó
    
    Dim rsFeGastoNeg As ADODB.Recordset
    Dim rsFeDatGastoFam As ADODB.Recordset
    Dim rsFeDatOtrosIng As ADODB.Recordset
    Dim rsFeDatBalanGen As ADODB.Recordset
    Dim rsFeDatActivos As ADODB.Recordset
    Dim rsFeDatPasivos As ADODB.Recordset
    Dim rsFeDatPasivosNo As ADODB.Recordset
    Dim rsFeDatPatrimonio As ADODB.Recordset
    Dim rsFeDatRef As ADODB.Recordset
    Dim rsFeFlujoCaja As ADODB.Recordset
    Dim rsFeDatPDT As ADODB.Recordset
    
    Dim rsCredEval As ADODB.Recordset
    Dim rsDCredito As ADODB.Recordset
    Dim rsAceptableCritico As ADODB.Recordset
    Dim rsCapacPagoNeta As ADODB.Recordset
    Dim rsCuotaIFIs As ADODB.Recordset
    Dim rsPropuesta As ADODB.Recordset
        
    Dim rsDatPasivosNo As ADODB.Recordset
    Dim rsDatActivoPasivo As ADODB.Recordset
    Dim rsDatGastoNeg As ADODB.Recordset
    Dim rsDatGastoFam As ADODB.Recordset
    Dim rsDatOtrosIng As ADODB.Recordset
    Dim rsDatRef As ADODB.Recordset
    Dim rsDatRatioInd As ADODB.Recordset
    Dim rsDatIfiGastoNego As ADODB.Recordset
    Dim rsDatIfiGastoFami As ADODB.Recordset
    Dim rsDatPDT As ADODB.Recordset
    Dim rsDatPDTDet As ADODB.Recordset
    Dim rsDatFlujoCaja As ADODB.Recordset
    Dim rsDatActivos As ADODB.Recordset
    Dim rsDatPasivos As ADODB.Recordset
    
    Dim nMontoAct As Double
    Dim nMontoPas As Double
    Dim nMontoPat As Double
    
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim objPista As COMManejador.Pista
    Dim fnFormato As Integer
    Dim fnMontoIni As Double
    Dim lnMin As Double
    Dim lnMax As Double
    Dim lnMinDol As Double
    Dim lnMaxDol As Double
    Dim nTC As Double
    Dim i As Integer, j As Integer, K As Integer
    
    Dim sMes1 As String, sMes2 As String, sMes3 As String
    Dim nMes1 As Integer, nMes2 As Integer, nMes3 As Integer
    Dim nAnio1 As Integer, nAnio2 As Integer, nAnio3 As Integer
    
    'LUCV20160705 **********-> Trabajando con Matrices TYPE
    Dim lvPrincipalActivos() As tFormEvalPrincipalActivosFormato5 'Matriz Principal-> Activos
    Dim lvPrincipalPasivos() As tFormEvalPrincipalPasivosFormato5 'Matriz Principal-> Pasivos
    'Detalle de Activos
    Dim lvDetalleActivosCtasCobrar() As tFormEvalDetalleActivosCtasCobrarFormato5 'Ctas x Cobrar
    Dim lvDetalleActivosInventario() As tFormEvalDetalleActivosInventarioFormato5 'Inventario
    Dim lvDetalleActivosActivoFijo() As tFormEvalDetalleActivosActivoFijoFormato5 'Activo Fijo
    'Detalle de Pasivos
    Dim lvDetallePasivosSobreGiros() As tFormEvalDetallePasivosSobreGirosFormato5 'Sobregiros y Prestamos
    'Dim lvDetallePasivosCtasCobrar() As tFormEvalDetallePasivosCtasCobrarFormato5 'Ctas x Cobrar
    
    Dim oFrmCtaCobrar As frmCredFormEvalCtasCobrar      'Formulario: Ctas x Cobrar
    Dim oFrmInve As frmCredFormEvalInventario           'Formulario: Inventario
    Dim oFrmActiFijo As frmCredFormEvalActivosFijos     'Formulario: Activo Fijo
    Dim oFrmSobre As frmCredFormEvalSobregirosPrestamos 'Formulario: Sobregiros y Prestamos
    'Fin LUCV20160705 <-**********
    Dim fbGrabar As Boolean
    Dim fnColocCondi As Integer
    Dim fbTieneReferido6Meses As Boolean 'LUCV20171115, Agregó segun correo: RUSI
    
    'RECO20160916******************************
    Dim bActivoDetPP(11) As Boolean
    Dim bActivoDetPE(11) As Boolean
    Dim bPasivoDetPP(5) As Boolean
    Dim bPasivoDetPE(5) As Boolean
    'RECO FIN *********************************

    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    Dim rsDatParamFlujoCajaForm5 As ADODB.Recordset
    Dim nMaximo As Integer
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    
    Dim lcMovNro As String 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
    Dim fnProducto As Integer 'CTI320200110 ERS003-2020
    Dim rsDatIfiNoSupervisadaGastoNego As ADODB.Recordset 'CTI320200110 ERS003-2020
    Dim rsDatIfiNoSupervisadaGastoFami As ADODB.Recordset 'CTI320200110 ERS003-2020
    Dim fbImprimirVB As Boolean 'CTI320200110 ERS003-2020
    Dim pnMontoOtrasIfisConsumo As Double
    Dim pnMontoOtrasIfisEmpresarial As Double
    
Private Sub cmdFlujoCaja5_Click()
On Error GoTo ErrorInicioExcel  'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
    
    Dim lsArchivo As String
    Dim lbLibroOpen As Boolean
    Dim bGeneraExcel As Boolean 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM

    'lsArchivo = App.Path & "\Spooler\FlujoCaja_Formato5" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls"  'comentado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
    lsArchivo = App.Path & "\Spooler\FlujoCaja_Formato5" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls" 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
    lbLibroOpen = ExcelInicio(lsArchivo, xlAplicacion, xlLibro)
    
    If lbLibroOpen Then
    bGeneraExcel = False 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
        bGeneraExcel = generaExcelForm5 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
        If bGeneraExcel Then 'modificado pti1 20180726
            ExcelFin lsArchivo, xlAplicacion, xlLibro, xlHoja1
            'AbrirArchivo "FlujoCaja_Formato5" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls", App.Path & "\Spooler"  'comentado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
            AbrirArchivo "FlujoCaja_Formato5" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls", App.Path & "\Spooler"
        End If
    End If
    
Exit Sub 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
ErrorInicioExcel: 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
MsgBox Err.Description + "Error1: Error al iniciar la creación del excel, Comuníquese con el Area de TI", vbInformation, "Error" 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
End Sub

Public Function generaExcelForm5() As Boolean
  On Error GoTo ErrorInicioExcel 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM

    Dim ssql As String
    Dim rs As New ADODB.Recordset
    Dim rsCabcera As New ADODB.Recordset
    Dim rsCuotas As New ADODB.Recordset
    Dim rsParFlujoCaja As New ADODB.Recordset
    Dim oCont As COMConecta.DCOMConecta
    Dim i As Integer
    Dim nCon As Integer
    Dim nFila As Integer
    Dim nCol As Integer
    Dim nColFin As Integer
    Dim nColInicio As Integer
    Dim a As Integer
    Dim Z As Integer
    Dim dFechaEval As Date
    Dim nTotalReg As Integer
    
    generaExcelForm5 = True
    
    'proteger Libro
    'xlAplicacion.ActiveWorkbook.Protect (123) 'pti comentado

    'Adiciona una hoja
    ExcelAddHoja "Hoja1", xlLibro, xlHoja1, True
               
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    
    xlHoja1.Cells(2, 2) = "FLUJO DE CAJA MENSUAL PRESUPUESTADO"
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 12)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 12)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 12)).Font.Bold = True
    
    xlHoja1.Cells(4, 1) = "CLIENTE: "
    xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 1)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(5, 1) = "ANALISTA: "
    xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 1)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(6, 1) = "DNI: "
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, 1)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(7, 1) = "RUC: "
    xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(7, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(7, 1)).HorizontalAlignment = xlLeft
    
    ssql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosCabecera  '" & ActXCodCta.NroCuenta & "'"

    Set oCont = New COMConecta.DCOMConecta
    oCont.AbreConexion
    Set rsCabcera = oCont.CargaRecordSet(ssql)
    oCont.CierraConexion
    Set oCont = Nothing
    
    ssql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosCuotas  '" & ActXCodCta.NroCuenta & "'"

    Set oCont = New COMConecta.DCOMConecta
    oCont.AbreConexion
    Set rsCuotas = oCont.CargaRecordSet(ssql)
    oCont.CierraConexion
    Set oCont = Nothing
    
    ssql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosConceptos  '" & ActXCodCta.NroCuenta & "'"

    Set oCont = New COMConecta.DCOMConecta
    oCont.AbreConexion
    Set rs = oCont.CargaRecordSet(ssql)
    oCont.CierraConexion
    Set oCont = Nothing
    
    ssql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosParametros  '" & ActXCodCta.NroCuenta & "'"

    Set oCont = New COMConecta.DCOMConecta
    oCont.AbreConexion
    Set rsParFlujoCaja = oCont.CargaRecordSet(ssql)
    oCont.CierraConexion
    Set oCont = Nothing

'Cabecera
If Not (rsCabcera.EOF And rsCabcera.BOF) Then
    dFechaEval = rsCabcera!fechaEval
    
    xlHoja1.Cells(4, 2) = rsCabcera!NombreClie
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 6)).Font.Bold = True

    xlHoja1.Cells(5, 2) = rsCabcera!NombreAnal
    xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(5, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(5, 6)).Font.Bold = True
    
    xlHoja1.Cells(6, 2) = rsCabcera!nDoc
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).Font.Bold = True
    
    xlHoja1.Cells(7, 2) = rsCabcera!nDocTrib
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).Font.Bold = True
    
Else
        MsgBox "Error, Comuníquese con el Área de TI", vbInformation, "!Error!"
        generaExcelForm5 = False
        Exit Function
End If
    
    
    xlHoja1.Cells(9, 2) = "Conceptos / Meses"
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(10, 2)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(9, 2)).Cells.Interior.Color = RGB(141, 180, 226)
        
    xlHoja1.Cells(9, 3) = "Flujo Mensual"
    xlHoja1.Cells(10, 3) = Format(dFechaEval, "mmm-yyyy")
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(10, 3)).Cells.Interior.Color = RGB(141, 180, 226)
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(10, 3)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(10, 3)).HorizontalAlignment = xlCenter
    
    CuadroExcel xlHoja1, 2, 9, 3, 10, True
    CuadroExcel xlHoja1, 2, 9, 3, 10, False
    
    nCon = 11
    
'Conceptos
    If Not (rs.EOF And rs.BOF) Then
        For i = 1 To rs.RecordCount
        
            CuadroExcel xlHoja1, 2, nCon, 3, nCon
        
            xlHoja1.Cells(nCon, 2) = rs!Descripcion
            xlHoja1.Cells(nCon, 3) = rs!Monto
            
            If rs!Descripcion = "INVERSION" Then
                nCon = nCon + 2
            Else
                nCon = nCon + 1
            End If
                        
            CuadroExcel xlHoja1, 2, nCon, 3, nCon - 1
            
            rs.MoveNext
        Next i
    Else
        MsgBox "Error, Comuníquese con el Área de TI", vbInformation, "!Error!"
        generaExcelForm5 = False
        Exit Function
    End If
  
'Pie
    If Not (rsParFlujoCaja.EOF And rsParFlujoCaja.BOF) Then

        xlHoja1.Cells(nCon + 1, 2) = "DATOS ADICIONALES"
        xlHoja1.Range(xlHoja1.Cells(nCon + 1, 2), xlHoja1.Cells(nCon + 1, 2)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(nCon + 1, 2), xlHoja1.Cells(nCon + 1, 2)).HorizontalAlignment = xlCenter
        CuadroExcel xlHoja1, 2, nCon + 1, 3, nCon + 1
        xlHoja1.Cells(nCon + 2, 2) = "Fecha de Pago"
        xlHoja1.Cells(nCon + 2, 3) = Format(rsParFlujoCaja!dFechaPago, "YYYY/mm/dd")
        CuadroExcel xlHoja1, 2, nCon + 2, 3, nCon + 2

        xlHoja1.Cells(nCon + 4, 3) = "Mes"
        xlHoja1.Cells(nCon + 4, 4) = "Anual"
        CuadroExcel xlHoja1, 3, nCon + 4, 4, nCon + 4
        xlHoja1.Range(xlHoja1.Cells(nCon + 4, 3), xlHoja1.Cells(nCon + 4, 4)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(nCon + 4, 3), xlHoja1.Cells(nCon + 4, 4)).HorizontalAlignment = xlCenter

        xlHoja1.Cells(nCon + 5, 2) = "Incremento de ventas al contado "
        xlHoja1.Cells(nCon + 6, 2) = "Incremento de Compra de Mercaderias"
        xlHoja1.Cells(nCon + 7, 2) = "Incremento de Consumo"
        xlHoja1.Cells(nCon + 8, 2) = "Incremento de Pago Personal"
        xlHoja1.Cells(nCon + 9, 2) = "Ingremento de Gastos de Ventas"

        xlHoja1.Cells(nCon + 5, 3) = Format(((1 + rsParFlujoCaja!nIncVentCont / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlHoja1.Cells(nCon + 6, 3) = Format(((1 + rsParFlujoCaja!nIncCompMerc / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlHoja1.Cells(nCon + 7, 3) = Format(((1 + rsParFlujoCaja!nIncConsu / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlHoja1.Cells(nCon + 8, 3) = Format(((1 + rsParFlujoCaja!nIncPagPers / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlHoja1.Cells(nCon + 9, 3) = Format(((1 + rsParFlujoCaja!nIncGastvent / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"

        xlHoja1.Cells(nCon + 5, 4) = Format(rsParFlujoCaja!nIncVentCont, "#0.0") & "%"
        xlHoja1.Cells(nCon + 6, 4) = Format(rsParFlujoCaja!nIncCompMerc, "#0.0") & "%"
        xlHoja1.Cells(nCon + 7, 4) = Format(rsParFlujoCaja!nIncConsu, "#0.0") & "%"
        xlHoja1.Cells(nCon + 8, 4) = Format(rsParFlujoCaja!nIncPagPers, "#0.0") & "%"
        xlHoja1.Cells(nCon + 9, 4) = Format(rsParFlujoCaja!nIncGastvent, "#0.0") & "%"

        CuadroExcel xlHoja1, 2, nCon + 5, 4, nCon + 9, True
        CuadroExcel xlHoja1, 2, nCon + 5, 4, nCon + 9, False
        xlHoja1.Range(xlHoja1.Cells(nCon + 5, 3), xlHoja1.Cells(nCon + 9, 4)).HorizontalAlignment = xlCenter

    Else
        MsgBox "Registre los Datos de Flujo de Caja Proyectado, y dar click en Guardar", vbInformation, "!Aviso!"
        generaExcelForm5 = False
        Exit Function
    End If
    
'Obtener las Letras del Abecedario A-Z
    Dim MatAZ5 As Variant
    Dim P As Integer
    P = 1
    Set MatAZ5 = Nothing
    ReDim MatAZ5(1, 140)
    For i = 65 To 90
        MatAZ5(1, P) = ChrW(i)
        P = P + 1
    Next i
           
    Dim MatLetrasRep5 As Variant
    Dim Y As Integer
    Set MatLetrasRep5 = Nothing
    Y = 1
    ReDim MatLetrasRep5(1, 131)
    For a = 1 To 130
        If a <= 26 Then
                MatLetrasRep5(1, Y) = ChrW(65) & MatAZ5(1, Y) 'AA,AB,AC......AZ
            Y = Y + 1
        ElseIf (a >= 27 And a <= 52) Then
            If a = 27 Then
                P = 1
            End If
                MatLetrasRep5(1, Y) = ChrW(66) & MatAZ5(1, P) 'BA,BB,BC......BZ
            Y = Y + 1
            P = P + 1
        ElseIf (a >= 53 And a <= 78) Then
            If a = 53 Then
                P = 1
            End If
                MatLetrasRep5(1, Y) = ChrW(67) & MatAZ5(1, P) 'CA,CB,CC......CZ
            Y = Y + 1
            P = P + 1
        ElseIf (a >= 79 And a <= 104) Then
            If a = 79 Then
                P = 1
            End If
                MatLetrasRep5(1, Y) = ChrW(68) & MatAZ5(1, P) 'DA,DB,DC......DZ
            Y = Y + 1
            P = P + 1
        End If
    Next a
    
    
'Cuotas
i = 0
Z = 0
Y = 0
nFila = 39
nCol = 4
nColInicio = 4
nColFin = 0

   If Not (rsCuotas.EOF And rsCuotas.BOF) Then
   
   nTotalReg = rsCuotas.RecordCount
   
        For i = 1 To rsCuotas.RecordCount

        If i >= 24 Then
            Y = Y + 1
        End If
                                    
            xlHoja1.Cells(9, nCol) = rsCuotas!nCuota
            xlHoja1.Range(xlHoja1.Cells(9, 4), xlHoja1.Cells(9, nCol)).Cells.Interior.Color = RGB(141, 180, 226)
            xlHoja1.Range(xlHoja1.Cells(9, nCol), xlHoja1.Cells(9, nCon)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(9, nCol), xlHoja1.Cells(9, nCon)).HorizontalAlignment = xlCenter
            
            xlHoja1.Cells(10, nCol) = Format(rsCuotas!dFechaCuotas, "mmm-yyyy")
            xlHoja1.Range(xlHoja1.Cells(10, 4), xlHoja1.Cells(10, nCol)).Cells.Interior.Color = RGB(141, 180, 226)
            xlHoja1.Range(xlHoja1.Cells(10, nCol), xlHoja1.Cells(10, nCon)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(10, nCol), xlHoja1.Cells(10, nCon)).HorizontalAlignment = xlCenter

            'calculo Ingresos Operativos
            xlHoja1.Range(xlHoja1.Cells(11, 2), xlHoja1.Cells(11, nCol)).Cells.Interior.Color = RGB(190, 190, 190)
            xlHoja1.Cells(11, nCol) = "=SUM(" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "12" & ":" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "15)"
            
            'Ventas al Contado
            xlHoja1.Cells(12, nCol) = Round((xlHoja1.Cells(12, nCol - 1) * ((1 + rsParFlujoCaja!nIncVentCont / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(12, nCol - 1)))
            
            'Ventas al Credito
            xlHoja1.Cells(13, nCol) = "=C13"
            
            'Ventas de Activos Fijos
            
            'Otros Ingresos
            xlHoja1.Cells(15, nCol) = "=C15"
            
            'calculo Engresos Operativos
            xlHoja1.Range(xlHoja1.Cells(16, 2), xlHoja1.Cells(16, nCol)).Cells.Interior.Color = RGB(190, 190, 190)
            xlHoja1.Cells(16, nCol) = "=SUM(" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "17" & ":" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "31)"
            
            'Egresos por Compras (Mercaderia)
            xlHoja1.Cells(17, nCol) = Round((xlHoja1.Cells(17, nCol - 1) * ((1 + rsParFlujoCaja!nIncCompMerc / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(17, nCol - 1)))
            
            'Personal
            xlHoja1.Cells(18, nCol) = Round((xlHoja1.Cells(18, nCol - 1) * ((1 + rsParFlujoCaja!nIncPagPers / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(18, nCol - 1)))
            
            'calculo Alquiler de Locales
            xlHoja1.Cells(19, nCol) = "=C19"

            'calculo Alquiler de Equipos
            xlHoja1.Cells(20, nCol) = "=C20"

            'calculo Servicios (luz....)
            xlHoja1.Cells(21, nCol) = "=C21"

            'calculo Utiles de oficinas
            xlHoja1.Cells(22, nCol) = "=C22"

            'calculo Rep y Mtto de Equipos
            xlHoja1.Cells(23, nCol) = "=C23"

            'calculo Rep y Mtto de Vehiculo
            xlHoja1.Cells(24, nCol) = "=C24"

            'calculo Seguro
            xlHoja1.Cells(25, nCol) = "=C25"

            'calculo Transporte/Combustible/ Gas
            xlHoja1.Cells(26, nCol) = "=C26"

            'calculo Contador
            xlHoja1.Cells(27, nCol) = "=C27"

            'calculo Sunat + Impuestos
            xlHoja1.Cells(28, nCol) = "=C28"

            'calculo Publicidad y otros gastos de ventas (**Nuevo)
            xlHoja1.Cells(29, nCol) = Round((xlHoja1.Cells(29, nCol - 1) * ((1 + rsParFlujoCaja!nIncPagPers / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(29, nCol - 1)))
            
            'calculo Otros
            xlHoja1.Cells(30, nCol) = "=C30"

            'calculo Consumo Per.Nat.
            xlHoja1.Cells(31, nCol) = Round((xlHoja1.Cells(31, nCol - 1) * ((1 + rsParFlujoCaja!nIncConsu / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(31, nCol - 1)))

            'calculo Flujo Operativo
            xlHoja1.Cells(32, nCol) = "=(" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "11" & "-" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "16)"
            xlHoja1.Range(xlHoja1.Cells(32, 2), xlHoja1.Cells(32, nCol)).Cells.Interior.Color = RGB(190, 190, 190)

            'Cobro de Prestamo y dividendos
            xlHoja1.Cells(33, nCol) = 0

            'Pago de cuota Prestamos vigentes
            xlHoja1.Cells(34, nCol) = "=C34"

            'Pago de cuotas de prestamos solicitado
            xlHoja1.Cells(35, nCol) = "=C35"

            'calculo Flujo Financiero
            xlHoja1.Cells(36, nCol) = "=(" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "32" & "+" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "33" & "-" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "34" & "-" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "35)"
            xlHoja1.Range(xlHoja1.Cells(36, 2), xlHoja1.Cells(36, nCol)).Cells.Interior.Color = RGB(190, 190, 190)
            
            'Inversiones
            xlHoja1.Cells(37, nCol) = 0
            
            'calculo Saldo
            xlHoja1.Cells(39, nCol) = "=(" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "36" & "-" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "37)"
            'Si los datos son numero negativos se pone rojo SALDO
            If xlHoja1.Cells(39, nCol) < 0 Then
                xlHoja1.Range(xlHoja1.Cells(39, nCol), xlHoja1.Cells(39, nCol)).Cells.Interior.Color = RGB(255, 0, 0)
            End If
            
            If i >= 25 Then
                Z = Z + 1
            End If

            'calculo Saldo Disponible
            xlHoja1.Cells(40, nCol) = "=(" & IIf(i >= 25, MatLetrasRep5(1, Z), MatAZ5(1, i + 2)) & "41)"
            'Si los datos son numero negativos se pone rojo Saldo Disponible
            If xlHoja1.Cells(40, nCol) < 0 Then
                xlHoja1.Range(xlHoja1.Cells(40, nCol), xlHoja1.Cells(40, nCol)).Cells.Interior.Color = RGB(255, 0, 0)
            End If

            'calculo Saldo Acumulado
            xlHoja1.Cells(41, nCol) = "=(" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "39" & "+" & IIf(i >= 24, MatLetrasRep5(1, Y), MatAZ5(1, i + 3)) & "40)"
            'Si los datos son numero negativos se pone rojo Saldo Disponible
            If xlHoja1.Cells(41, nCol) < 0 Then
                xlHoja1.Range(xlHoja1.Cells(41, nCol), xlHoja1.Cells(41, nCol)).Cells.Interior.Color = RGB(255, 0, 0)
            End If
            
            nCol = nCol + 1

            If (i Mod 12) = 0 Then
                nColFin = nCol - 1
                    xlHoja1.Cells(8, nColInicio) = "Año" & (i / 12)
                    xlHoja1.Range(xlHoja1.Cells(8, nColInicio), xlHoja1.Cells(8, nColFin)).HorizontalAlignment = xlCenter
                    xlHoja1.Range(xlHoja1.Cells(8, nColInicio), xlHoja1.Cells(8, nColFin)).MergeCells = True
                    xlHoja1.Range(xlHoja1.Cells(8, nColInicio), xlHoja1.Cells(8, nColFin)).Font.Bold = True
                nColInicio = nColFin + 1
            End If
            rsCuotas.MoveNext
        Next i
        
        If nColInicio <> nCol Then
        'Para la celda si no cumple un año
        xlHoja1.Range(xlHoja1.Cells(8, nColInicio), xlHoja1.Cells(8, nCol - 1)).MergeCells = True
        End If
        
        xlHoja1.Range(xlHoja1.Cells(8, 4), xlHoja1.Cells(8, nCol - 1)).Cells.Interior.Color = RGB(141, 180, 226)
        CuadroExcel xlHoja1, 4, 8, nCol - 1, 8

        For i = 0 To 33
            If i <= 28 Then
                CuadroExcel xlHoja1, 4, 9 + i, nCol - 1, 9 + i
            ElseIf i >= 31 Then
                CuadroExcel xlHoja1, 4, nFila, nCol - 1, nFila
                nFila = nFila + 1
            End If
        Next i
        
    Else
        MsgBox "Error al crear el Excel, Comuníquese con el Área de TI", vbInformation, "!Error!"
        Exit Function
    End If
    
xlHoja1.Cells.Select
xlHoja1.Cells.Font.Name = "Arial"
xlHoja1.Cells.Font.Size = 9
xlHoja1.Cells.EntireColumn.AutoFit

'xlAplicacion.Worksheets("Hoja1").Protect ("123")
      
MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "!Exito!"

rs.Close
rsCabcera.Close
rsParFlujoCaja.Close
rsCuotas.Close

Exit Function 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
ErrorInicioExcel: 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
MsgBox Err.Description + "Error 2: Error al iniciar la creación del excel, Comuníquese con el Área de TI", vbInformation, "Error" 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM

End Function

'JOEP20180725 ERS034-2018
Private Sub cmdMNME_Click()
    Call frmCredFormEvalCredCel.Inicio(ActXCodCta.NroCuenta, 11)
End Sub
'JOEP20180725 ERS034-2018

Private Sub EditMoneyForm5IVC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EditMoneyForm5ICM.SetFocus
        fEnfoque EditMoneyForm5ICM
    End If
End Sub

Private Sub EditMoneyForm5ICM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EditMoneyForm5IPP.SetFocus
        fEnfoque EditMoneyForm5IPP
    End If
End Sub

Private Sub EditMoneyForm5IPP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EditMoneyForm5IGV.SetFocus
        fEnfoque EditMoneyForm5IGV
    End If
End Sub
'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja

'_____________________________________________________________________________________________________________
'******************************************LUCV20160525: EVENTOS Varios***************************************
Private Sub Form_Load()
    fbGrabar = False
    CentraForm Me
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    EnfocaControl spnTiempoLocalAnio
    'RECO20160916 ***********************
    bActivoDetPP(0) = False
    bActivoDetPP(1) = False
    bActivoDetPP(2) = False
    bActivoDetPP(3) = False
    bActivoDetPP(4) = False
    bActivoDetPP(5) = False
    bActivoDetPP(6) = False
    bActivoDetPP(7) = False
    bActivoDetPP(8) = False
    bActivoDetPP(9) = False
    bActivoDetPP(10) = False
    
    bActivoDetPE(0) = False
    bActivoDetPE(1) = False
    bActivoDetPE(2) = False
    bActivoDetPE(3) = False
    bActivoDetPE(4) = False
    bActivoDetPE(5) = False
    bActivoDetPE(6) = False
    bActivoDetPE(7) = False
    bActivoDetPE(8) = False
    bActivoDetPE(9) = False
    bActivoDetPE(10) = False
    
    
    bPasivoDetPP(0) = False
    bPasivoDetPP(1) = False
    bPasivoDetPP(2) = False
    bPasivoDetPP(3) = False
    bPasivoDetPP(4) = False
    
    bPasivoDetPE(0) = False
    bPasivoDetPE(1) = False
    bPasivoDetPE(2) = False
    bPasivoDetPE(3) = False
    bPasivoDetPE(4) = False
    'RECO FIN 20160916*********

'JOEP20180725 ERS034-2018
    If fnTipoRegMant = 3 Then
        If Not ConsultaRiesgoCamCred(sCtaCod) Then
            cmdMNME.Visible = True
        End If
    End If
'JOEP20180725 ERS034-2018
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set MatIfiGastoNego = Nothing 'LUCV20161115
    Set MatIfiGastoFami = Nothing 'LUCV20161115
    
    Set MatIfiNoSupervisadaGastoNego = Nothing 'CTI320200110 ERS003-2020. Agregó
    Set MatIfiNoSupervisadaGastoFami = Nothing 'CTI320200110 ERS003-2020. Agregó
End Sub
Private Sub feActivos_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean) ' Activos
    Dim pnMonto As Double
    Dim Index As Integer
    Dim nTotal As Double
    Dim nTipoPat As Integer
    
    nTipoPat = IIf(feActivos.col = 2, 1, 2)
    If feActivos.TextMatrix(1, 0) = "" Then Exit Sub
    Index = CInt(feActivos.TextMatrix(feActivos.row, 0))
    
    Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0))
        Case 3, 4 '*************************-> Ctas x Cobrar
            Set oFrmCtaCobrar = New frmCredFormEvalCtasCobrar
            If feActivos.col = 2 Then 'Column P.P.
                If IsArray(lvPrincipalActivos(Index).vPPActivoCtaCobrar) Then
                    lvDetalleActivosCtasCobrar = lvPrincipalActivos(Index).vPPActivoCtaCobrar
                Else
                    ReDim lvDetalleActivosCtasCobrar(0)
                End If
            End If
            If feActivos.col = 3 Then 'Column P.E.
                If IsArray(lvPrincipalActivos(Index).vPEActivoCtaCobrar) Then
                    lvDetalleActivosCtasCobrar = lvPrincipalActivos(Index).vPEActivoCtaCobrar
                Else
                    ReDim lvDetalleActivosCtasCobrar(0)
                End If
            End If
            
            If oFrmCtaCobrar.Inicio(lvDetalleActivosCtasCobrar, nTotal, CInt(feActivos.TextMatrix(Me.feActivos.row, 5)), CInt(feActivos.TextMatrix(Me.feActivos.row, 6)), feActivos.TextMatrix(Me.feActivos.row, 1), ActXCodCta.NroCuenta, IIf(feActivos.col = 2, 1, 2)) Then
                If feActivos.col = 2 Then 'column P.P.
                    lvPrincipalActivos(Index).vPPActivoCtaCobrar = lvDetalleActivosCtasCobrar
                    'RECO20160916***************
                        Select Case Index
                        Case 3
                            bActivoDetPP(0) = True
                        Case 4
                            bActivoDetPP(1) = True
                        End Select
                    'RECO FIN*******************
                End If
                If feActivos.col = 3 Then ' columna P.E.
                    lvPrincipalActivos(Index).vPEActivoCtaCobrar = lvDetalleActivosCtasCobrar
                    'RECO20160916***************
                        Select Case Index
                        Case 3
                            bActivoDetPE(0) = True
                        Case 4
                            bActivoDetPE(1) = True
                        End Select
                    'RECO FIN*******************
                End If
            End If
            'Call CalculaCeldas
            If feActivos.col = 2 Then
                Me.feActivos.TextMatrix(Me.feActivos.row, Me.feActivos.col) = Format(nTotal, "#,#0.00")
            End If
            If feActivos.col = 3 Then
                Me.feActivos.TextMatrix(Me.feActivos.row, Me.feActivos.col) = Format(nTotal, "#,#0.00")
            End If
            Call CalculoTotal(1)
            'Fin - Ctas x Cobrar <-**********
            
        Case 7, 8, 9, 10 '********************-> Inventario
            Set oFrmInve = New frmCredFormEvalInventario
            If feActivos.col = 2 Then 'Column P.P.
                If IsArray(lvPrincipalActivos(Index).vPPActivoInventario) Then
                    lvDetalleActivosInventario = lvPrincipalActivos(Index).vPPActivoInventario
                Else
                    ReDim lvDetalleActivosInventario(0)
                End If
            End If
            If feActivos.col = 3 Then 'column P.E.
                If IsArray(lvPrincipalActivos(Index).vPEActivoInventario) Then
                    lvDetalleActivosInventario = lvPrincipalActivos(Index).vPEActivoInventario
                Else
                    ReDim lvDetalleActivosInventario(0)
                End If
            End If

            If oFrmInve.Inicio(lvDetalleActivosInventario, nTotal, CInt(feActivos.TextMatrix(Me.feActivos.row, 5)), CInt(feActivos.TextMatrix(Me.feActivos.row, 6)), feActivos.TextMatrix(Me.feActivos.row, 1), ActXCodCta.NroCuenta, nTipoPat) Then
                If feActivos.col = 2 Then 'column P.P.
                    lvPrincipalActivos(Index).vPPActivoInventario = lvDetalleActivosInventario
                    'RECO20160916***************
                        Select Case Index
                        Case 7
                            bActivoDetPP(2) = True
                        Case 8
                            bActivoDetPP(3) = True
                        Case 9
                            bActivoDetPP(4) = True
                        Case 10
                            bActivoDetPP(5) = True
                        End Select
                    'RECO FIN*******************
                End If
                If feActivos.col = 3 Then ' columna P.E.
                    lvPrincipalActivos(Index).vPEActivoInventario = lvDetalleActivosInventario
                    'RECO20160916***************
                        Select Case Index
                        Case 7
                            bActivoDetPE(2) = True
                        Case 8
                            bActivoDetPE(3) = True
                        Case 9
                            bActivoDetPE(4) = True
                        Case 10
                            bActivoDetPE(5) = True
                        End Select
                    'RECO FIN*******************
                End If
            End If
                 'Call CalculoTotal(1)
            If feActivos.col = 2 Then
                Me.feActivos.TextMatrix(Me.feActivos.row, Me.feActivos.col) = Format(nTotal, "#,#0.00")
            End If
            If feActivos.col = 3 Then
                Me.feActivos.TextMatrix(Me.feActivos.row, Me.feActivos.col) = Format(nTotal, "#,#0.00")
            End If
            Call CalculoTotal(1)
            'Fin - Inventario <-**********

        Case 12, 13, 14, 15, 16 '********************-> Activo / Fijo
            Set oFrmActiFijo = New frmCredFormEvalActivosFijos
            If feActivos.col = 2 Then 'column P.P.
                If IsArray(lvPrincipalActivos(Index).vPPActivoActivoFijo) Then
                    lvDetalleActivosActivoFijo = lvPrincipalActivos(Index).vPPActivoActivoFijo
                Else
                    ReDim lvDetalleActivosActivoFijo(0)
                End If
            End If
            If feActivos.col = 3 Then 'column P.E.
                If IsArray(lvPrincipalActivos(Index).vPPEctivoActivoFijo) Then
                    lvDetalleActivosActivoFijo = lvPrincipalActivos(Index).vPPEctivoActivoFijo
                Else
                    ReDim lvDetalleActivosActivoFijo(0)
                End If
            End If

            If oFrmActiFijo.Inicio(lvDetalleActivosActivoFijo, nTotal, CInt(feActivos.TextMatrix(Me.feActivos.row, 5)), CInt(feActivos.TextMatrix(Me.feActivos.row, 6)), feActivos.TextMatrix(Me.feActivos.row, 1), ActXCodCta.NroCuenta, nTipoPat) Then
                If feActivos.col = 2 Then 'column P.P.
                    lvPrincipalActivos(Index).vPPActivoActivoFijo = lvDetalleActivosActivoFijo
                    'RECO20160916***************
                        Select Case Index
                        Case 12
                            bActivoDetPP(6) = True
                        Case 13
                            bActivoDetPP(7) = True
                        Case 14
                            bActivoDetPP(8) = True
                        Case 15
                            bActivoDetPP(9) = True
                        Case 16
                            bActivoDetPP(10) = True
                        End Select
                    'RECO FIN*******************
                End If
                If feActivos.col = 3 Then ' columna P.E.
                    lvPrincipalActivos(Index).vPPEctivoActivoFijo = lvDetalleActivosActivoFijo
                    'RECO20160916***************
                        Select Case Index
                        Case 12
                            bActivoDetPE(6) = True
                        Case 13
                            bActivoDetPE(7) = True
                        Case 14
                            bActivoDetPE(8) = True
                        Case 15
                            bActivoDetPE(9) = True
                        Case 16
                            bActivoDetPE(10) = True
                        End Select
                    'RECO FIN*******************
                End If
            End If
            'Call CalculaCeldas
            If feActivos.col = 2 Then
                Me.feActivos.TextMatrix(Me.feActivos.row, Me.feActivos.col) = Format(nTotal, "#,#0.00")
            End If
            If feActivos.col = 3 Then
                Me.feActivos.TextMatrix(Me.feActivos.row, Me.feActivos.col) = Format(nTotal, "#,#0.00")
            End If
            Call CalculoTotal(1)
            'Fin - Activo/Fijo <-**********
        End Select
End Sub
Private Sub fePasivos_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean) 'Pasivos
Dim pnMonto As Double
Dim Index As Integer
Dim nTotal As Double
      
    If fePasivos.TextMatrix(1, 0) = "" Then Exit Sub
    Index = CInt(fePasivos.TextMatrix(fePasivos.row, 0))

    Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0))
        Case 5, 6, 7, 9, 10 '*************************-> Ctas x Cobrar
            Dim oFrm As Form
            If pnRow = 7 Or pnRow = 9 Then
                Set oFrm = New frmCredFormEvalCtasCobrarIfis
            Else
                Set oFrm = New frmCredFormEvalCtasCobrar
            End If
            If fePasivos.col = 2 Then 'column P.P.
                If IsArray(lvPrincipalPasivos(Index).vPPActivoCtaCobrar) Then
                    lvDetalleActivosCtasCobrar = lvPrincipalPasivos(Index).vPPActivoCtaCobrar
                Else
                    ReDim lvDetalleActivosCtasCobrar(0)
                End If
            End If
            If fePasivos.col = 3 Then 'column P.E.
                If IsArray(lvPrincipalPasivos(Index).vPEActivoCtaCobrar) Then
                    lvDetalleActivosCtasCobrar = lvPrincipalPasivos(Index).vPEActivoCtaCobrar
                Else
                    ReDim lvDetalleActivosCtasCobrar(0)
                End If
            End If
            
            If oFrm.Inicio(lvDetalleActivosCtasCobrar, nTotal, CInt(fePasivos.TextMatrix(Me.fePasivos.row, 5)), CInt(fePasivos.TextMatrix(Me.fePasivos.row, 6)), fePasivos.TextMatrix(Me.fePasivos.row, 1), ActXCodCta.NroCuenta, IIf(fePasivos.col = 2, 1, 2)) Then
                If fePasivos.col = 2 Then 'column P.P.
                    lvPrincipalPasivos(Index).vPPActivoCtaCobrar = lvDetalleActivosCtasCobrar
                    If IsArray(lvPrincipalPasivos(Index).vPPActivoCtaCobrar) Then
                        lvDetalleActivosCtasCobrar = lvPrincipalPasivos(Index).vPPActivoCtaCobrar
                    Else
                        ReDim lvDetalleActivosCtasCobrar(0)
                    End If
                End If
                If fePasivos.col = 3 Then ' columna P.E.
                    lvPrincipalPasivos(Index).vPEActivoCtaCobrar = lvDetalleActivosCtasCobrar
                    'RECO20160916***************
                        Select Case Index
                        Case 5
                            bPasivoDetPP(0) = True
                        Case 6
                            bPasivoDetPP(1) = True
                        Case 7
                            bPasivoDetPP(2) = True
                        Case 9
                            bPasivoDetPP(3) = True
                        Case 10
                            bPasivoDetPP(4) = True
                        End Select
                    'RECO FIN*******************
                End If
            End If
            If fePasivos.col = 2 Then
                Me.fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.col) = Format(nTotal, "#,#0.00")
            End If

            If fePasivos.col = 3 Then
                Me.fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.col) = Format(nTotal, "#,#0.00")
            End If
              Call CalculoTotal(2)
            'Fin - Ctas x Cobrar <-**********
        End Select
End Sub
Private Sub Cmdguardar_Click()
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim GrabarDatos As Boolean
    Dim rsGastoFam As ADODB.Recordset
    Dim rsOtrosIng As ADODB.Recordset
    Dim rsBalGen As ADODB.Recordset
    Dim rsFlujoCaja As ADODB.Recordset
    Dim rsPDT As ADODB.Recordset
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Dim rsRatiosActual As ADODB.Recordset
    Dim rsRatiosAceptableCritico As ADODB.Recordset
    
    'feDeclaracionPDT.TextMatrix(0, 3) = "MesDetalle"
    feDeclaracionPDT.TextMatrix(0, 4) = "Mes1"
    feDeclaracionPDT.TextMatrix(0, 5) = "Mes2"
    feDeclaracionPDT.TextMatrix(0, 6) = "Mes3"
    feDeclaracionPDT.TextMatrix(0, 8) = "VentasDeclaradas"
    Set rsPDT = IIf(feDeclaracionPDT.rows - 1 > 0, feDeclaracionPDT.GetRsNew(), Nothing)
    'feDeclaracionPDT.TextMatrix(0, 3) = "Mes/Detalle"
    feDeclaracionPDT.TextMatrix(0, 4) = sMes3
    feDeclaracionPDT.TextMatrix(0, 5) = sMes2
    feDeclaracionPDT.TextMatrix(0, 6) = sMes1
    feDeclaracionPDT.TextMatrix(0, 8) = "%Vent. Decl."

    Set rsFlujoCaja = IIf(feFlujoCajaMensual.rows - 1 > 0, feFlujoCajaMensual.GetRsNew(), Nothing)
    Set rsOtrosIng = IIf(feOtrosIngresos.rows - 1 > 0, feOtrosIngresos.GetRsNew(), Nothing)
    Set rsGastoFam = IIf(feGastosFamiliares.rows - 1 > 0, feGastosFamiliares.GetRsNew(), Nothing)
    'Flex a Matriz Referidos **********->
        ReDim MatReferidos(feReferidos.rows - 1, 6)
        For i = 1 To feReferidos.rows - 1
            MatReferidos(i, 1) = feReferidos.TextMatrix(i, 0)
            MatReferidos(i, 2) = feReferidos.TextMatrix(i, 1)
            MatReferidos(i, 3) = feReferidos.TextMatrix(i, 2)
            MatReferidos(i, 4) = feReferidos.TextMatrix(i, 3)
            MatReferidos(i, 5) = feReferidos.TextMatrix(i, 4)
            MatReferidos(i, 6) = feReferidos.TextMatrix(i, 5)
         Next i
    'Fin Referidos
    
    If ValidaDatos Then
        If Not fbImprimirVB Then
            If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        End If
        If txtUltEndeuda.Text = "__/__/____" Then
            txtUltEndeuda.Text = "01/01/1900"
        End If

        Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
        Set objPista = New COMManejador.Pista
        
        If fnTipoPermiso = 3 Then
    '********************************************************************************************
    '*** LUCV20160709, PARA EL LLENADO DEL FORMATO  5                                         ***
    '********************************************************************************************
    'RECO20160730********************************************************************************
            '->*****Activos
            If IsArray(lvPrincipalActivos(3).vPPActivoCtaCobrar) Then
                If UBound(lvPrincipalActivos(3).vPPActivoCtaCobrar) <= 0 And bActivoDetPP(0) = False Then '-> PP: Ctas x Cobrar
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7025, 102, 1)
                    lvPrincipalActivos(3).vPPActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalActivos(3).vPEActivoCtaCobrar) Then
                If UBound(lvPrincipalActivos(3).vPEActivoCtaCobrar) <= 0 And bActivoDetPE(0) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7025, 102, 2)
                    lvPrincipalActivos(3).vPEActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            
            If IsArray(lvPrincipalActivos(4).vPPActivoCtaCobrar) Then '-> PP:Otras Ctas x Cobrar
                If UBound(lvPrincipalActivos(4).vPPActivoCtaCobrar) <= 0 And bActivoDetPP(1) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7025, 106, 1)
                    lvPrincipalActivos(4).vPPActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalActivos(4).vPEActivoCtaCobrar) Then
                If UBound(lvPrincipalActivos(4).vPEActivoCtaCobrar) <= 0 And bActivoDetPE(1) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7025, 106, 2)
                    lvPrincipalActivos(4).vPEActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            '->***** Inventario
            If IsArray(lvPrincipalActivos(7).vPPActivoInventario) Then '->PP: 301-Mercaderia
                If UBound(lvPrincipalActivos(7).vPPActivoInventario) <= 0 And bActivoDetPP(2) = False Then
                    Call CargarMatrizDatosMantenimientoInvetario(lvDetalleActivosInventario, ActXCodCta.NroCuenta, 7025, 301, 1)
                    lvPrincipalActivos(7).vPPActivoInventario = lvDetalleActivosInventario
                End If
                ReDim lvDetalleActivosInventario(0)
            End If
            If IsArray(lvPrincipalActivos(7).vPEActivoInventario) Then '->PE
                If UBound(lvPrincipalActivos(7).vPEActivoInventario) <= 0 And bActivoDetPE(2) = False Then
                    Call CargarMatrizDatosMantenimientoInvetario(lvDetalleActivosInventario, ActXCodCta.NroCuenta, 7025, 301, 2)
                    lvPrincipalActivos(7).vPEActivoInventario = lvDetalleActivosInventario
                End If
                ReDim lvDetalleActivosInventario(0)
            End If
            If IsArray(lvPrincipalActivos(8).vPPActivoInventario) Then '->PP: 302-Productos Terminados
                If UBound(lvPrincipalActivos(8).vPPActivoInventario) <= 0 And bActivoDetPP(3) = False Then
                    Call CargarMatrizDatosMantenimientoInvetario(lvDetalleActivosInventario, ActXCodCta.NroCuenta, 7025, 302, 1)
                    lvPrincipalActivos(8).vPPActivoInventario = lvDetalleActivosInventario
                End If
                ReDim lvDetalleActivosInventario(0)
            End If
            If IsArray(lvPrincipalActivos(8).vPEActivoInventario) Then  '->PE
                If UBound(lvPrincipalActivos(8).vPEActivoInventario) <= 0 And bActivoDetPE(3) = False Then
                    Call CargarMatrizDatosMantenimientoInvetario(lvDetalleActivosInventario, ActXCodCta.NroCuenta, 7025, 302, 2)
                    lvPrincipalActivos(8).vPEActivoInventario = lvDetalleActivosInventario
                End If
                ReDim lvDetalleActivosInventario(0)
            End If
            If IsArray(lvPrincipalActivos(9).vPPActivoInventario) Then  '->PP: 303-Productos en Proceso
                If UBound(lvPrincipalActivos(9).vPPActivoInventario) <= 0 And bActivoDetPP(4) = False Then
                    Call CargarMatrizDatosMantenimientoInvetario(lvDetalleActivosInventario, ActXCodCta.NroCuenta, 7025, 303, 1)
                    lvPrincipalActivos(9).vPPActivoInventario = lvDetalleActivosInventario
                End If
                ReDim lvDetalleActivosInventario(0)
            End If
            If IsArray(lvPrincipalActivos(9).vPEActivoInventario) Then  '->PE
                If UBound(lvPrincipalActivos(9).vPEActivoInventario) <= 0 And bActivoDetPE(4) = False Then
                    Call CargarMatrizDatosMantenimientoInvetario(lvDetalleActivosInventario, ActXCodCta.NroCuenta, 7025, 303, 2)
                    lvPrincipalActivos(9).vPEActivoInventario = lvDetalleActivosInventario
                End If
                ReDim lvDetalleActivosInventario(0)
            End If
            If IsArray(lvPrincipalActivos(10).vPPActivoInventario) Then '-> PP: 304-Materia Prima (insumos)
                If UBound(lvPrincipalActivos(10).vPPActivoInventario) <= 0 And bActivoDetPP(5) = False Then
                    Call CargarMatrizDatosMantenimientoInvetario(lvDetalleActivosInventario, ActXCodCta.NroCuenta, 7025, 304, 1)
                    lvPrincipalActivos(10).vPPActivoInventario = lvDetalleActivosInventario
                End If
                ReDim lvDetalleActivosInventario(0)
            End If
            If IsArray(lvPrincipalActivos(10).vPEActivoInventario) Then '->PE
                If UBound(lvPrincipalActivos(10).vPEActivoInventario) <= 0 And bActivoDetPE(5) = False Then
                    Call CargarMatrizDatosMantenimientoInvetario(lvDetalleActivosInventario, ActXCodCta.NroCuenta, 7025, 304, 2)
                    lvPrincipalActivos(10).vPEActivoInventario = lvDetalleActivosInventario
                End If
                ReDim lvDetalleActivosInventario(0)
            End If
            '*****************************************************************************************************************
            '->***** Activos Fijos
            If IsArray(lvPrincipalActivos(12).vPPActivoActivoFijo) Then '->PP: 401-Muebles y equipos de Ofic.
                If UBound(lvPrincipalActivos(12).vPPActivoActivoFijo) <= 0 And bActivoDetPP(6) = False Then
                    Call CargarMatrizDatosMantenimientoActivoFijo(lvDetalleActivosActivoFijo, ActXCodCta.NroCuenta, 7025, 401, 1)
                    lvPrincipalActivos(12).vPPActivoActivoFijo = lvDetalleActivosActivoFijo
                End If
                ReDim lvDetalleActivosActivoFijo(0)
            End If
            If IsArray(lvPrincipalActivos(12).vPPEctivoActivoFijo) Then '->PE
                If UBound(lvPrincipalActivos(12).vPPEctivoActivoFijo) <= 0 And bActivoDetPE(6) = False Then
                    Call CargarMatrizDatosMantenimientoActivoFijo(lvDetalleActivosActivoFijo, ActXCodCta.NroCuenta, 7025, 401, 2)
                    lvPrincipalActivos(12).vPPEctivoActivoFijo = lvDetalleActivosActivoFijo
                End If
                ReDim lvDetalleActivosActivoFijo(0)
            End If
            If IsArray(lvPrincipalActivos(13).vPPActivoActivoFijo) Then '->PP:402-Vehiculos
                If UBound(lvPrincipalActivos(13).vPPActivoActivoFijo) <= 0 And bActivoDetPP(7) = False Then
                    Call CargarMatrizDatosMantenimientoActivoFijo(lvDetalleActivosActivoFijo, ActXCodCta.NroCuenta, 7025, 402, 1)
                    lvPrincipalActivos(13).vPPActivoActivoFijo = lvDetalleActivosActivoFijo
                End If
                ReDim lvDetalleActivosActivoFijo(0)
            End If
            If IsArray(lvPrincipalActivos(13).vPPEctivoActivoFijo) Then '->PE
                If UBound(lvPrincipalActivos(13).vPPEctivoActivoFijo) <= 0 And bActivoDetPE(7) = False Then
                    Call CargarMatrizDatosMantenimientoActivoFijo(lvDetalleActivosActivoFijo, ActXCodCta.NroCuenta, 7025, 402, 2)
                    lvPrincipalActivos(13).vPPEctivoActivoFijo = lvDetalleActivosActivoFijo
                End If
                ReDim lvDetalleActivosActivoFijo(0)
            End If
            If IsArray(lvPrincipalActivos(14).vPPActivoActivoFijo) Then '->PP: 403-Maquinarias y Equipos
                If UBound(lvPrincipalActivos(14).vPPActivoActivoFijo) <= 0 And bActivoDetPP(8) = False Then
                    Call CargarMatrizDatosMantenimientoActivoFijo(lvDetalleActivosActivoFijo, ActXCodCta.NroCuenta, 7025, 403, 1)
                    lvPrincipalActivos(14).vPPActivoActivoFijo = lvDetalleActivosActivoFijo
                End If
                ReDim lvDetalleActivosActivoFijo(0)
            End If
            If IsArray(lvPrincipalActivos(14).vPPEctivoActivoFijo) Then '->PE
                If UBound(lvPrincipalActivos(14).vPPEctivoActivoFijo) <= 0 And bActivoDetPE(8) = False Then
                    Call CargarMatrizDatosMantenimientoActivoFijo(lvDetalleActivosActivoFijo, ActXCodCta.NroCuenta, 7025, 403, 2)
                    lvPrincipalActivos(14).vPPEctivoActivoFijo = lvDetalleActivosActivoFijo
                End If
                ReDim lvDetalleActivosActivoFijo(0)
            End If
            If IsArray(lvPrincipalActivos(15).vPPActivoActivoFijo) Then '->PP: 404-Edificios y Terrenos
                If UBound(lvPrincipalActivos(15).vPPActivoActivoFijo) <= 0 And bActivoDetPP(9) = False Then
                    Call CargarMatrizDatosMantenimientoActivoFijo(lvDetalleActivosActivoFijo, ActXCodCta.NroCuenta, 7025, 404, 1)
                    lvPrincipalActivos(15).vPPActivoActivoFijo = lvDetalleActivosActivoFijo
                End If
                ReDim lvDetalleActivosActivoFijo(0)
            End If
            If IsArray(lvPrincipalActivos(15).vPPEctivoActivoFijo) Then '->PE
                If UBound(lvPrincipalActivos(15).vPPEctivoActivoFijo) <= 0 And bActivoDetPE(9) = False Then
                    Call CargarMatrizDatosMantenimientoActivoFijo(lvDetalleActivosActivoFijo, ActXCodCta.NroCuenta, 7025, 404, 2)
                    lvPrincipalActivos(15).vPPEctivoActivoFijo = lvDetalleActivosActivoFijo
                End If
                ReDim lvDetalleActivosActivoFijo(0)
            End If
            If IsArray(lvPrincipalActivos(16).vPPActivoActivoFijo) Then '->PP: 405-Herramientas y Otros
                If UBound(lvPrincipalActivos(16).vPPActivoActivoFijo) <= 0 And bActivoDetPP(10) = False Then
                    Call CargarMatrizDatosMantenimientoActivoFijo(lvDetalleActivosActivoFijo, ActXCodCta.NroCuenta, 7025, 405, 1)
                    lvPrincipalActivos(16).vPPActivoActivoFijo = lvDetalleActivosActivoFijo
                End If
                ReDim lvDetalleActivosActivoFijo(0)
            End If
            If IsArray(lvPrincipalActivos(16).vPPEctivoActivoFijo) Then '->PE
                If UBound(lvPrincipalActivos(16).vPPEctivoActivoFijo) <= 0 And bActivoDetPE(10) = False Then
                    Call CargarMatrizDatosMantenimientoActivoFijo(lvDetalleActivosActivoFijo, ActXCodCta.NroCuenta, 7025, 405, 2)
                    lvPrincipalActivos(16).vPPEctivoActivoFijo = lvDetalleActivosActivoFijo
                End If
                ReDim lvDetalleActivosActivoFijo(0)
            End If
    '********************************************************************************************
    'PASIVO*************************************
            If IsArray(lvPrincipalPasivos(5).vPPActivoCtaCobrar) Then
                If UBound(lvPrincipalPasivos(5).vPPActivoCtaCobrar) <= 0 And bPasivoDetPP(0) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 104, 1)
                    lvPrincipalPasivos(5).vPPActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalPasivos(5).vPEActivoCtaCobrar) Then
                If UBound(lvPrincipalPasivos(5).vPEActivoCtaCobrar) <= 0 And bPasivoDetPP(0) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 104, 2)
                    lvPrincipalPasivos(5).vPEActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalPasivos(6).vPPActivoCtaCobrar) Then
                If UBound(lvPrincipalPasivos(6).vPPActivoCtaCobrar) <= 0 And bPasivoDetPP(1) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 107, 1)
                    lvPrincipalPasivos(6).vPPActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalPasivos(6).vPEActivoCtaCobrar) Then
                If UBound(lvPrincipalPasivos(6).vPEActivoCtaCobrar) <= 0 And bPasivoDetPP(1) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 107, 2)
                    lvPrincipalPasivos(6).vPEActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalPasivos(7).vPPActivoCtaCobrar) Then
                If UBound(lvPrincipalPasivos(7).vPPActivoCtaCobrar) <= 0 And bPasivoDetPP(2) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 109, 1)
                    lvPrincipalPasivos(7).vPPActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalPasivos(7).vPEActivoCtaCobrar) Then
                If UBound(lvPrincipalPasivos(7).vPEActivoCtaCobrar) <= 0 And bPasivoDetPP(2) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 109, 2)
                    lvPrincipalPasivos(7).vPEActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalPasivos(9).vPPActivoCtaCobrar) Then
                If UBound(lvPrincipalPasivos(9).vPPActivoCtaCobrar) <= 0 And bPasivoDetPP(3) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 201, 1)
                    lvPrincipalPasivos(9).vPPActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalPasivos(9).vPEActivoCtaCobrar) Then
                If UBound(lvPrincipalPasivos(9).vPEActivoCtaCobrar) <= 0 And bPasivoDetPP(3) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 201, 2)
                    lvPrincipalPasivos(9).vPEActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalPasivos(10).vPPActivoCtaCobrar) Then
                If UBound(lvPrincipalPasivos(10).vPPActivoCtaCobrar) <= 0 And bPasivoDetPP(4) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 203, 1)
                    lvPrincipalPasivos(10).vPPActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalPasivos(10).vPEActivoCtaCobrar) Then
                If UBound(lvPrincipalPasivos(10).vPEActivoCtaCobrar) <= 0 And bPasivoDetPP(4) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 203, 2)
                    lvPrincipalPasivos(10).vPEActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
    '*******************************************
    
    
    
    'RECO FIN ***********************************************************************************
    '********************************************************************************************
    'Eliminamos Datos existentes
    If UBound(lvPrincipalActivos) > 0 Or UBound(lvPrincipalPasivos) > 0 Then
        Call oDCOMFormatosEval.EliminaCredFormEvalGrillaActiPasi(sCtaCod, fnFormato)
        Call oDCOMFormatosEval.EliminaCredFormEvalGrillaActiPasiDet(sCtaCod, fnFormato)
    End If
        '--------------------- ACTIVOS(CredFormEvalActivoPasivo / CredFormEvalActivoPasivoDet)
        If UBound(lvPrincipalActivos) > 0 Then
            For i = 1 To UBound(lvPrincipalActivos)
                'If CDbl(Me.feActivos.TextMatrix(i, 4)) > 0 Then
                    If i = 17 Then
                        Call oDCOMFormatosEval.InsertaCredFormEvalActiPasi(sCtaCod, fnFormato, Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                        CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), CCur(Me.feActivos.TextMatrix(i, 4)), _
                        CCur(Me.feActivos.TextMatrix(i, 2)), CCur(Me.feActivos.TextMatrix(i, 3)))
                    Else
                        Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiDet(sCtaCod, fnFormato, Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                        CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), CCur(Me.feActivos.TextMatrix(i, 4)), _
                        CCur(Me.feActivos.TextMatrix(i, 2)), CCur(Me.feActivos.TextMatrix(i, 3)))
                    End If
                'End If
            Next i
        End If
    '---------------------- ACTIVOS -> Detalle Celdas (PP / PE)
        If UBound(lvPrincipalActivos) > 0 Then
        For i = 1 To UBound(lvPrincipalActivos)
                
                'Detalle de Celdas -> Cuentas por Cobrar ********************->
                    If IsArray(lvPrincipalActivos(i).vPPActivoCtaCobrar) Then 'CtasxCobrar->PP
                        For j = 1 To UBound(lvPrincipalActivos(i).vPPActivoCtaCobrar)
                            Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiCtaCobrar(sCtaCod, fnFormato, _
                                                                                    gCodPatrimonioPersonal, j, _
                                                                                    CInt(Me.feActivos.TextMatrix(i, 5)), _
                                                                                    CInt(Me.feActivos.TextMatrix(i, 6)), _
                                                                                    Format(lvPrincipalActivos(i).vPPActivoCtaCobrar(j).dfecha, "yyyyMMdd"), _
                                                                                    lvPrincipalActivos(i).vPPActivoCtaCobrar(j).cCtaporCobrar, _
                                                                                    lvPrincipalActivos(i).vPPActivoCtaCobrar(j).nTotal)
                        Next j
                    End If
 
                    If IsArray(lvPrincipalActivos(i).vPEActivoCtaCobrar) Then 'CtasxCobrar->PE
                        For j = 1 To UBound(lvPrincipalActivos(i).vPEActivoCtaCobrar)
                            Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiCtaCobrar(sCtaCod, fnFormato, _
                                                                                    gCodPatrimonioEmpresarial, j, _
                                                                                    CInt(Me.feActivos.TextMatrix(i, 5)), _
                                                                                    CInt(Me.feActivos.TextMatrix(i, 6)), _
                                                                                    Format(lvPrincipalActivos(i).vPEActivoCtaCobrar(j).dfecha, "yyyyMMdd"), _
                                                                                    lvPrincipalActivos(i).vPEActivoCtaCobrar(j).cCtaporCobrar, _
                                                                                    lvPrincipalActivos(i).vPEActivoCtaCobrar(j).nTotal)
                        Next j
                    End If

               'Detalle de Celdas -> Inventario   **************************->
                 If IsArray(lvPrincipalActivos(i).vPPActivoInventario) Then 'Inventario ->PP
                     For j = 1 To UBound(lvPrincipalActivos(i).vPPActivoInventario)
                         Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiInventario(sCtaCod, _
                                                                                      fnFormato, _
                                                                                      gCodPatrimonioPersonal, _
                                                                                      j, _
                                                                                      CInt(Me.feActivos.TextMatrix(i, 5)), _
                                                                                      CInt(Me.feActivos.TextMatrix(i, 6)), _
                                                                                      Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                                                                                      lvPrincipalActivos(i).vPPActivoInventario(j).cMercaderia, _
                                                                                      lvPrincipalActivos(i).vPPActivoInventario(j).nCantidad, _
                                                                                      lvPrincipalActivos(i).vPPActivoInventario(j).cUnidMed, _
                                                                                      lvPrincipalActivos(i).vPPActivoInventario(j).nCostoUnit, _
                                                                                      lvPrincipalActivos(i).vPPActivoInventario(j).nTotal)
                     Next j
                 End If
                 If IsArray(lvPrincipalActivos(i).vPEActivoInventario) Then 'Invetario->PE
                     For j = 1 To UBound(lvPrincipalActivos(i).vPEActivoInventario)
                         Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiInventario(sCtaCod, _
                                                                                      fnFormato, _
                                                                                      gCodPatrimonioEmpresarial, _
                                                                                       j, _
                                                                                      CInt(Me.feActivos.TextMatrix(i, 5)), _
                                                                                      CInt(Me.feActivos.TextMatrix(i, 6)), _
                                                                                      Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                                                                                      lvPrincipalActivos(i).vPEActivoInventario(j).cMercaderia, _
                                                                                      lvPrincipalActivos(i).vPEActivoInventario(j).nCantidad, _
                                                                                      lvPrincipalActivos(i).vPEActivoInventario(j).cUnidMed, _
                                                                                      lvPrincipalActivos(i).vPEActivoInventario(j).nCostoUnit, _
                                                                                      lvPrincipalActivos(i).vPEActivoInventario(j).nTotal)
                     Next j
                End If

                'Detalle de Celdas -> Activo Fijo   **************************->
                 If IsArray(lvPrincipalActivos(i).vPPActivoActivoFijo) Then 'Activo Fijo ->PP
                     For j = 1 To UBound(lvPrincipalActivos(i).vPPActivoActivoFijo)
                         Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiActivoFijo(sCtaCod, _
                                                                                      fnFormato, _
                                                                                      j, _
                                                                                      gCodPatrimonioPersonal, _
                                                                                      CInt(Me.feActivos.TextMatrix(i, 5)), _
                                                                                      CInt(Me.feActivos.TextMatrix(i, 6)), _
                                                                                      lvPrincipalActivos(i).vPPActivoActivoFijo(j).CDescripcion, _
                                                                                      lvPrincipalActivos(i).vPPActivoActivoFijo(j).nCantidad, _
                                                                                      lvPrincipalActivos(i).vPPActivoActivoFijo(j).nPrecio, _
                                                                                      lvPrincipalActivos(i).vPPActivoActivoFijo(j).nTotal)
                     Next j
                 End If
                 If IsArray(lvPrincipalActivos(i).vPPEctivoActivoFijo) Then 'Activo Fijo->PE
                     For j = 1 To UBound(lvPrincipalActivos(i).vPPEctivoActivoFijo)
                         Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiActivoFijo(sCtaCod, _
                                                                                      fnFormato, _
                                                                                      j, _
                                                                                      gCodPatrimonioEmpresarial, _
                                                                                      CInt(Me.feActivos.TextMatrix(i, 5)), _
                                                                                      CInt(Me.feActivos.TextMatrix(i, 6)), _
                                                                                      lvPrincipalActivos(i).vPPEctivoActivoFijo(j).CDescripcion, _
                                                                                      lvPrincipalActivos(i).vPPEctivoActivoFijo(j).nCantidad, _
                                                                                      lvPrincipalActivos(i).vPPEctivoActivoFijo(j).nPrecio, _
                                                                                      lvPrincipalActivos(i).vPPEctivoActivoFijo(j).nTotal)
                     Next j
                End If
     Next i
End If
    '--------------------- PASIVOS (Formato)
        If UBound(lvPrincipalPasivos) > 0 Then
            For i = 1 To UBound(lvPrincipalPasivos)
                'If CDbl(Me.fePasivos.TextMatrix(i, 4)) > 0 Then
                   
                    If (i = 13) Then
                        Call oDCOMFormatosEval.InsertaCredFormEvalActiPasi(sCtaCod, fnFormato, Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                        CInt(Me.fePasivos.TextMatrix(i, 5)), gCodTotalPatrimonio, CDbl(Me.fePasivos.TextMatrix(i, 4)), _
                        CDbl(Me.fePasivos.TextMatrix(i, 2)), CDbl(Me.fePasivos.TextMatrix(i, 3)))
                    End If
                   
                    If (i = 20) Or (i = 21) Then
                        Call oDCOMFormatosEval.InsertaCredFormEvalActiPasi(sCtaCod, fnFormato, Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                        CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), CDbl(Me.fePasivos.TextMatrix(i, 4)), _
                        CDbl(Me.fePasivos.TextMatrix(i, 2)), CDbl(Me.fePasivos.TextMatrix(i, 3)))
                    
                    Else
                        Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiDet(sCtaCod, fnFormato, Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                        CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), CDbl(Me.fePasivos.TextMatrix(i, 4)), _
                        CDbl(Me.fePasivos.TextMatrix(i, 2)), CDbl(Me.fePasivos.TextMatrix(i, 3)))
                    End If
                    
               ' End If
            Next i
        End If
        '---------------------- PASIVOS -> Detalle Celdas (PP / PE)
        If UBound(lvPrincipalActivos) > 0 Then
            
        For i = 1 To UBound(lvPrincipalPasivos)
                'Detalle de Celdas -> Sobregiros   **************************->
                 If IsArray(lvPrincipalPasivos(i).vPPPasivoSobregiro) Then 'Sobregiros ->PP
                     For j = 1 To UBound(lvPrincipalPasivos(i).vPPPasivoSobregiro)
                         Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiSobregiros(sCtaCod, _
                                                                                      fnFormato, _
                                                                                      CInt(Me.fePasivos.TextMatrix(i, 5)), _
                                                                                      CInt(Me.fePasivos.TextMatrix(i, 6)), _
                                                                                      gCodPatrimonioPersonal, _
                                                                                      j, _
                                                                                      lvPrincipalPasivos(i).vPPPasivoSobregiro(j).cEntidad, _
                                                                                      lvPrincipalPasivos(i).vPPPasivoSobregiro(j).cCredito, _
                                                                                      lvPrincipalPasivos(i).vPPPasivoSobregiro(j).nTEA, _
                                                                                      lvPrincipalPasivos(i).vPPPasivoSobregiro(j).nCuotas, _
                                                                                      lvPrincipalPasivos(i).vPPPasivoSobregiro(j).nMontoCuota, _
                                                                                      lvPrincipalPasivos(i).vPPPasivoSobregiro(j).nCuotasPend)
                     Next j
                 End If
                 If IsArray(lvPrincipalPasivos(i).vPEPasivoSobregiro) Then 'Sobregiros->PE
                     For j = 1 To UBound(lvPrincipalPasivos(i).vPEPasivoSobregiro)
                         Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiSobregiros(sCtaCod, _
                                                                                      fnFormato, _
                                                                                      CInt(Me.fePasivos.TextMatrix(i, 5)), _
                                                                                      CInt(Me.fePasivos.TextMatrix(i, 6)), _
                                                                                      gCodPatrimonioEmpresarial, _
                                                                                      j, _
                                                                                      lvPrincipalPasivos(i).vPEPasivoSobregiro(j).cEntidad, _
                                                                                      lvPrincipalPasivos(i).vPEPasivoSobregiro(j).cCredito, _
                                                                                      lvPrincipalPasivos(i).vPEPasivoSobregiro(j).nTEA, _
                                                                                      lvPrincipalPasivos(i).vPEPasivoSobregiro(j).nCuotas, _
                                                                                      lvPrincipalPasivos(i).vPEPasivoSobregiro(j).nMontoCuota, _
                                                                                      lvPrincipalPasivos(i).vPEPasivoSobregiro(j).nCuotasPend)
                     Next j
                End If
                
                'Detalle de Celdas -> Cuentas por Cobrar ********************->
                    If IsArray(lvPrincipalPasivos(i).vPPActivoCtaCobrar) Then 'CtasxCobrar->PP
                        For j = 1 To UBound(lvPrincipalPasivos(i).vPPActivoCtaCobrar)
                            Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiCtaCobrar(sCtaCod, fnFormato, _
                                                                                    gCodPatrimonioPersonal, j, _
                                                                                    CInt(Me.fePasivos.TextMatrix(i, 5)), _
                                                                                    CInt(Me.fePasivos.TextMatrix(i, 6)), _
                                                                                    Format(lvPrincipalPasivos(i).vPPActivoCtaCobrar(j).dfecha, "yyyyMMdd"), _
                                                                                    lvPrincipalPasivos(i).vPPActivoCtaCobrar(j).cCtaporCobrar, _
                                                                                    lvPrincipalPasivos(i).vPPActivoCtaCobrar(j).nTotal)
                        Next j
                    End If
 
                    If IsArray(lvPrincipalPasivos(i).vPEActivoCtaCobrar) Then 'CtasxCobrar->PE
                        For j = 1 To UBound(lvPrincipalPasivos(i).vPEActivoCtaCobrar)
                            Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiCtaCobrar(sCtaCod, fnFormato, _
                                                                                    gCodPatrimonioEmpresarial, j, _
                                                                                    CInt(Me.fePasivos.TextMatrix(i, 5)), _
                                                                                    CInt(Me.fePasivos.TextMatrix(i, 6)), _
                                                                                    Format(lvPrincipalPasivos(i).vPEActivoCtaCobrar(j).dfecha, "yyyyMMdd"), _
                                                                                    lvPrincipalPasivos(i).vPEActivoCtaCobrar(j).cCtaporCobrar, _
                                                                                    lvPrincipalPasivos(i).vPEActivoCtaCobrar(j).nTotal)
                        Next j
                    End If
     Next i
End If

'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    Dim MatFlujoCaja As Variant
    Set MatFlujoCaja = Nothing
    ReDim MatFlujoCaja(1, 5)
        For i = 1 To 1
            MatFlujoCaja(i, 1) = EditMoneyForm5IVC
            MatFlujoCaja(i, 2) = EditMoneyForm5ICM
            MatFlujoCaja(i, 3) = EditMoneyForm5IPP
            MatFlujoCaja(i, 4) = EditMoneyForm5IGV
            MatFlujoCaja(i, 5) = EditMoneyForm5IC
        Next i
'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja

'Fin <- ********** LUCV20160709
        Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
        GrabarDatos = oNCOMFormatosEval.GrabarCredFormEvalFormato1_5(sCtaCod, fnFormato, fnTipoRegMant, _
                                                                    Trim(txtGiroNeg.Text), CInt(spnExpEmpAnio.valor), CInt(spnExpEmpMes.valor), CInt(spnTiempoLocalAnio.valor), _
                                                                    CInt(spnTiempoLocalMes.valor), CDbl(txtUltEndeuda.Text), Format(txtFecUltEndeuda.Text, "yyyymmdd"), _
                                                                    fnCondLocal, IIf(txtCondLocalOtros.Visible = False, "", txtCondLocalOtros.Text), CDbl(txtExposicionCredito.Text), _
                                                                    Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                                                                    Format(txtFechaVisita.Text, "yyyymmdd"), _
                                                                    txtEntornoFamiliar5.Text, txtGiroUbicacion5.Text, _
                                                                    txtExperiencia5.Text, txtFormalidadNegocio5.Text, _
                                                                    txtColaterales5, txtDestino5.Text, _
                                                                    txtComentario5.Text, MatReferidos, MatIfiGastoNego, MatIfiGastoFami, _
                                                                    rsGastoFam, rsOtrosIng, _
                                                                    , , , , , , _
                                                                    rsFlujoCaja, rsPDT, _
                                                                    gRatioCapacidadPago, _
                                                                    CDbl(Replace(txtCapacidadNeta.Text, "%", "")), _
                                                                    gRatioEndeudamiento, _
                                                                    CDbl(Replace(txtEndeudamiento.Text, "%", "")), _
                                                                    gRatioIngresoNetoNego, _
                                                                    CDbl(txtIngresoNeto.Text), _
                                                                    gRatioExcedenteMensual, _
                                                                    CDbl(txtExcedenteMensual.Text), _
                                                                    nMes1, nMes2, nMes3, nAnio1, nAnio2, nAnio3, fnColocCondi, MatFlujoCaja, , , _
                                                                    MatIfiNoSupervisadaGastoNego, MatIfiNoSupervisadaGastoFami)
                                                                    
                                                                    'MatIfiNoSupervisadaGastoNego, MatIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agregó
                                                                    'JOEP20171015 Flujo de Caja MatFlujoCaja
                                                                    
                    Call oDCOMFormatosEval.RecalculaIndicadoresyRatiosEvaluacion(sCtaCod)
                    Set rsRatiosActual = oDCOMFormatosEval.RecuperaDatosRatios(sCtaCod)
                    Set rsRatiosAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(sCtaCod)
                    
                    'JOEP20180725 ERS034-2018
                    Call EmiteFormRiesgoCamCred(sCtaCod)
                    'JOEP20180725 ERS034-2018
'        Else
        'GrabarDatos = oNCOMFormatosEval.GrabarCredEvaluacionVerif(sCtaCod, Trim(txtVerif.Text), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        End If
            If GrabarDatos Then
                'RECO20161020 ERS060-2016 **********************************************************
                Dim oNCOMColocEval As New NCOMColocEval
                'Dim lcMovNro As String 'LUCV20181220 Comentó, Anexo01 de Acta 199-2018
                lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                
                If Not ValidaExisteRegProceso(sCtaCod, gTpoRegCtrlEvaluacion) Then
                   'lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Comentó, Anexo01 de Acta 199-2018
                   'objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 5", sCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                   Call oNCOMColocEval.insEstadosExpediente(sCtaCod, "Evaluacion de Credito", lcMovNro, "", "", "", 1, 2001, gTpoRegCtrlEvaluacion)
                   Set oNCOMColocEval = Nothing
                End If
                'RECO FIN **************************************************************************
                fbGrabar = True
                If fnTipoRegMant = 1 Then
                    objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 5", sCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                    Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                    If Not fbImprimirVB Then
                        MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
                    End If
                Else
                    objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Evaluacion Credito Formato 5", sCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                    Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                    If Not fbImprimirVB Then
                        MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
                    End If
                    Dim objCredito As COMDCredito.DCOMCredito
                    Set objCredito = New COMDCredito.DCOMCredito
                    Call objCredito.ActualizarEstadoxVB(ActXCodCta.NroCuenta, 1)
                End If
                
                'CTI320200110 ERS003-2020. Comentó, método sin finalidad o existencia
                'FondoCrecerBitacora
                'Dim objFCBS_UP As COMDCredito.DCOMCredito
                'Set objFCBS_UP = New COMDCredito.DCOMCredito
                'objFCBS_UP.FondoCrecerBitacora IIf(fnTipoRegMant = 1, gCredRegistrarEvaluacionCred, gCredMantenimientoEvaluacionCred), lcMovNro, gsCodPersUser, sCtaCod, "Formato de Evaluación (Sicmac Negocio)"
                'Set objFCBS_UP = Nothing
                'FondoCrecerBitacora
                'Fin CTI320200110
                
                'Habilita / Deshabilita Botones - Text
                 If fnEstado = 2000 Then         '*****-> Si es Solicitado
                    If fnColocCondi <> 4 Then
                        Me.cmdInformeVisita.Enabled = True
                        Me.cmdVerCar.Enabled = False
                    Else
                        Me.cmdInformeVisita.Enabled = False
                        Me.cmdVerCar.Enabled = False
                    End If
                    Me.cmdImprimir.Enabled = False
                    Me.cmdGuardar.Enabled = False
                Else                             '*****-> Sugerido +
                    Me.cmdImprimir.Enabled = True
                    Me.cmdGuardar.Enabled = False
                    If fnColocCondi <> 4 Then
                        Me.cmdVerCar.Enabled = True
                        Me.cmdInformeVisita.Enabled = True
                    Else
                        Me.cmdVerCar.Enabled = False
                        Me.cmdInformeVisita.Enabled = False
                    End If
                    
                End If
                '*****->No Refinanciados (Propuesta Credito)
                    If fnColocCondi <> 4 Then
                        txtFechaVisita.Enabled = True
                        txtEntornoFamiliar5.Enabled = True
                        txtGiroUbicacion5.Enabled = True
                        txtExperiencia5.Enabled = True
                        txtFormalidadNegocio5.Enabled = True
                        txtColaterales5.Enabled = True
                        txtDestino5.Enabled = True
                     Else
                        framePropuesta.Enabled = False
                        txtFechaVisita.Enabled = False
                        txtEntornoFamiliar5.Enabled = False
                        txtGiroUbicacion5.Enabled = False
                        txtExperiencia5.Enabled = False
                        txtFormalidadNegocio5.Enabled = False
                        txtColaterales5.Enabled = False
                        txtDestino5.Enabled = False
                    End If
                '*****->Fin No Refinanciados
                
                'Actualizacion de los Ratios
                    txtCapacidadNeta.Text = CStr(rsRatiosActual!nCapPagNeta * 100) & "%"
                    txtEndeudamiento.Text = CStr(rsRatiosActual!nEndeuPat * 100) & "%"
                    txtLiquidezCte.Text = CStr(Format(rsRatiosActual!nLiquidezCte, "#0.00"))
                    txtRentabilidadPat.Text = CStr(rsRatiosActual!nRentaPatri * 100) & "%"
                    txtIngresoNeto.Text = Format(rsRatiosActual!nIngreNeto, "#,##0.00")
                    txtExcedenteMensual.Text = Format(rsRatiosActual!nExceMensual, "#,##0.00")
                    
                'Ratios: Aceptable / Critico ->*****
                    If Not (rsRatiosAceptableCritico.EOF Or rsRatiosAceptableCritico.BOF) Then
                    If rsRatiosAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
                        Me.lblCapaAceptable.Caption = "Aceptable"
                        Me.lblCapaAceptable.ForeColor = &H8000&
                    Else
                        Me.lblCapaAceptable.Caption = "Crítico"
                        Me.lblCapaAceptable.ForeColor = vbRed
                    End If
                    
                    If rsRatiosAceptableCritico!nEndeud = 1 Then 'Endeudamiento Pat.
                        Me.lblEndeAceptable.Caption = "Aceptable"
                        Me.lblEndeAceptable.ForeColor = &H8000&
                    Else
                        Me.lblEndeAceptable.Caption = "Crítico"
                        Me.lblEndeAceptable.ForeColor = vbRed
                    End If
                    Else
                        lblCapaAceptable.Visible = False
                        lblEndeAceptable.Visible = False
                    End If
                'Fin Ratios <-****
                    Set rsRatiosActual = Nothing
                    Set rsRatiosAceptableCritico = Nothing
            Else
                MsgBox "Hubo errores al grabar la información", vbError, "Error"
            End If
    'Else
    'MsgBox "Ha Ocurrido un Problema o Faltan Ingresar Datos", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdImprimir_Click()
    Call ImprimirFormatoEvaluacion
End Sub

Private Sub cmdVerCar_Click()
   Call GeneraVerCar
End Sub
Private Sub cmdInformeVisita_Click()
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsInfVisita As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsInfVisita = New ADODB.Recordset
    Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(sCtaCod)
    cmdInformeVisita.Enabled = False
    If (rsInfVisita.EOF And rsInfVisita.BOF) Then
        Set oDCOMFormatosEval = Nothing
        MsgBox "No existe datos para este reporte.", vbOKOnly, "Atención"
        Exit Sub
    End If
    Call CargaInformeVisitaPDF(rsInfVisita) 'gCredReportes
    Set rsInfVisita = Nothing
    cmdInformeVisita.Enabled = True
End Sub
Private Sub cmdCancelar_Click()
    Unload frmCredFormEvalCuotasIfis
    Unload Me
    Set MatIfiGastoNego = Nothing 'LUCV20161115
    Set MatIfiGastoFami = Nothing 'LUCV20161115
    
    Set MatIfiNoSupervisadaGastoNego = Nothing 'CTI320200110 ERS003-2020. Agregó
    Set MatIfiNoSupervisadaGastoFami = Nothing 'CTI320200110 ERS003-2020. Agregó
End Sub
Private Sub cmdAgregarRef_Click()
    If feReferidos.rows - 1 < 25 Then
        feReferidos.lbEditarFlex = True
        feReferidos.AdicionaFila
        feReferidos.SetFocus
        feReferidos.AvanceCeldas = Horizontal
        SendKeys "{Enter}"
    Else
        MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub
Private Sub cmdQuitar5_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feReferidos.EliminaFila (feReferidos.row)
    End If
End Sub


'LUCV20160620, KeyPress / GotFocus / LostFocus ->**********
    'TAB0 -> Ingresos/Egresos
Private Sub spnTiempoLocalAnio_KeyPress(KeyAscii As Integer) 'TiempoMismoLocal
    If KeyAscii = 13 Then
        spnTiempoLocalMes.SetFocus
    End If
End Sub
Private Sub spnTiempoLocalMes_KeyPress(KeyAscii As Integer) 'TiempoMismoLocal
    If KeyAscii = 13 Then
        OptCondLocal(1).SetFocus
    End If
End Sub
Private Sub OptCondLocal_KeyPress(Index As Integer, KeyAscii As Integer) 'CondicionLocal
     If fnProducto <> "800" Then 'CTI320200110 ERS003-2020
        If KeyAscii = 13 Then
            SSTabIngresos5.Tab = 0
            feActivos.row = 2
            feActivos.col = 2
            EnfocaControl feActivos
            SendKeys "{Enter}"
        End If
    Else
        If KeyAscii = 13 Then
            SSTabIngresos5.Tab = 1
            feFlujoCajaMensual.row = 1
            feFlujoCajaMensual.col = 4
            EnfocaControl feFlujoCajaMensual
            SendKeys "{Enter}"
        End If
    End If
End Sub

Private Sub txtCondLocalOtros_KeyPress(KeyAscii As Integer) 'OtroCondicionLocal
    If fnProducto <> "800" Then 'CTI320200110 ERS003-2020
        If KeyAscii = 13 Then
            SSTabIngresos5.Tab = 0
            feActivos.row = 2
            feActivos.col = 2
            EnfocaControl feActivos
            SendKeys "{Enter}"
        End If
    Else
       If KeyAscii = 13 Then
            SSTabIngresos5.Tab = 1
            feFlujoCajaMensual.row = 1
            feFlujoCajaMensual.col = 4
            EnfocaControl feFlujoCajaMensual
            SendKeys "{Enter}"
        End If
    End If
End Sub

   'TAB1 ->PropuestaCredito
Private Sub txtFechaVisita_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEntornoFamiliar5.SetFocus
        
        If Not IsDate(txtFechaVisita) Then
            MsgBox "Verifique Dia,Mes,Año , Fecha Incorrecta", vbInformation, "Aviso"
            txtFechaVisita.SetFocus
        End If
        
    End If
End Sub

Private Sub txtEntornoFamiliar5_KeyPress(KeyAscii As Integer) 'Entornofamiliar
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtGiroUbicacion5.SetFocus
    End If
End Sub
Private Sub txtGiroUbicacion5_KeyPress(KeyAscii As Integer) 'SobreGiro
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtExperiencia5.SetFocus
    End If
End Sub
Private Sub txtExperiencia5_KeyPress(KeyAscii As Integer) 'ExperienciaCrediticia
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtFormalidadNegocio5.SetFocus
    End If
End Sub
Private Sub txtFormalidadNegocio5_KeyPress(KeyAscii As Integer) 'ConsistenciaInformacion
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtColaterales5.SetFocus
    End If
End Sub
Private Sub txtColaterales5_KeyPress(KeyAscii As Integer) 'Colaterales_Garantias
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtDestino5.SetFocus
    End If
End Sub
Private Sub txtDestino5_KeyPress(KeyAscii As Integer) 'Destino del crédito
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        SSTabIngresos5.Tab = 3
        'If fnColocCondi = 1 Then 'LUCV20171115, Agregó segun correo: RUSI
        If Not fbTieneReferido6Meses Then
            txtComentario5.SetFocus
        Else
            cmdGuardar.SetFocus
        End If
    End If
End Sub
    'TAB1 ->ComentarioReferido
Private Sub txtComentario5_KeyPress(KeyAscii As Integer) 'Referidos/ ComentariosReferidos
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        If fnColocCondi = 1 Then
            cmdAgregarRef.SetFocus
        End If
    End If
End Sub
'LUCV20160620, KeyPress / GotFocus / LostFocus Fin <-**********

'Calcular Activos / Pasivos / FlujoCajaMensual
'Private Sub feActivos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
'    Call CalculoTotal(1)
'End Sub
Private Sub feActivos_KeyPress(KeyAscci As Integer)
    If (Me.feActivos.col = 4 And Me.feActivos.row = 17) Then
        Me.fePasivos.SetFocus
        fePasivos.row = 2
        fePasivos.col = 2
        SendKeys "{F2}"
    End If
End Sub


Private Sub feActivos_GotFocus()
    If (Me.feActivos.col = 3 And Me.feActivos.row = 17) Then
        Me.fePasivos.SetFocus
        fePasivos.row = 2
        fePasivos.col = 2
        SendKeys "{Enter}"
    End If
End Sub
Private Sub feActivos_RowColChange()
    If feActivos.col = 2 Then
        feActivos.AvanceCeldas = Horizontal
    ElseIf feActivos.col = 3 Then
        feActivos.AvanceCeldas = Vertical
    End If
    
    If feActivos.col = 2 Or feActivos.col = 3 Then
    Select Case CInt(feActivos.TextMatrix(feActivos.row, 6))
        Case 102, 106, 301, 302, 303, 304, 401, 402, 403, 404, 405
            feActivos.ListaControles = "0-0-1-1-0-0-0"
        Case Else
            feActivos.ListaControles = "0-0-0-0-0-0-0"
    End Select
    End If
    
    Select Case CInt(feActivos.TextMatrix(feActivos.row, 6)) 'celda que  o se puede editar
        Case 1000, 100, 200, 300, 400, 500
            Me.feActivos.BackColorRow (&H80000000)
            Me.feActivos.ForeColorRow vbBlack, True
            Me.feActivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 102, 106, 301, 302, 303, 304, 401, 402, 403, 404, 405 'Colorear color amarillo
            Me.feActivos.BackColorRow &HC0FFFF, True
            Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        Case Else
            Me.feActivos.BackColorRow (&HFFFFFF)
            Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
     End Select
     Call CalculoTotal(1)
End Sub
Private Sub feActivos_EnterCell()
    If feActivos.col = 2 Or feActivos.col = 3 Then
        Select Case CInt(feActivos.TextMatrix(feActivos.row, 6))
            Case 102, 106, 301, 302, 303, 304, 401, 402, 403, 404, 405
              feActivos.ListaControles = "0-0-1-1-0-0-0"
            Case Else
              feActivos.ListaControles = "0-0-0-0-0-0-0"
            End Select
        End If
    
        Select Case CInt(feActivos.TextMatrix(feActivos.row, 6)) 'celda que  o se puede editar
            Case 1000, 100, 200, 300, 400, 500
                Me.feActivos.BackColorRow (&H80000000)
                Me.feActivos.ForeColorRow vbBlack, True
                Me.feActivos.ColumnasAEditar = "X-X-X-X-X-X-X"
            Case 102, 106, 301, 302, 303, 304, 401, 402, 403, 404, 405 'Colorear color amarillo
                Me.feActivos.BackColorRow &HC0FFFF, True
                Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
            Case Else
                Me.feActivos.BackColorRow (&HFFFFFF)
                Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
         End Select
         Call CalculoTotal(1)
End Sub

Private Sub feActivos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean) 'Activos
    Dim Editar() As String
    Editar = Split(Me.feActivos.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}", True
        Exit Sub
    End If
End Sub
Private Sub feActivos_Click()
    If feActivos.col = 2 Or feActivos.col = 3 Then
        Select Case CInt(feActivos.TextMatrix(feActivos.row, 6))
            Case 102, 106, 301, 302, 303, 304, 401, 402, 403, 404, 405
              feActivos.ListaControles = "0-0-1-1-0-0-0"
            Case Else
              feActivos.ListaControles = "0-0-0-0-0-0-0"
        End Select
    End If
    
    Select Case CInt(feActivos.TextMatrix(feActivos.row, 6)) 'celda que  o se puede editar
        Case 1000, 100, 200, 300, 400, 500
            Me.feActivos.BackColorRow (&H80000000)
            Me.feActivos.ForeColorRow vbBlack, True
            Me.feActivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 102, 106, 301, 302, 303, 304, 401, 402, 403, 404, 405 'Colorear color amarillo
            Me.feActivos.BackColorRow &HC0FFFF, True
            Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        Case Else
            Me.feActivos.BackColorRow (&HFFFFFF)
            Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
    End Select
End Sub
Private Sub feActivos_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feActivos.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feActivos.TextMatrix(pnRow, pnCol) < 0 Then
            feActivos.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feActivos.TextMatrix(pnRow, pnCol) = 0
    End If
'    If (Me.feActivos.Col = 3 And Me.feActivos.row = 16) Then
'        Me.fePasivos.SetFocus
'        SSTabIngresos5.Tab = 0
'        fePasivos.row = 2
'        fePasivos.Col = 2
'        SendKeys "{TAB}"
'        SendKeys "{TAB}"
'    End If
    Call CalculoTotal(1)
End Sub

Private Sub fePasivos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    'Call CalculoTotal(2)
    Dim Editar() As String
    Editar = Split(Me.fePasivos.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}", True
        Exit Sub
    End If
End Sub
Private Sub fePasivos_KeyPress(KeyAscii As Integer)
    If (Me.fePasivos.col = 4 And Me.fePasivos.row = 21) Then
        SSTabIngresos5.Tab = 1
        Me.feFlujoCajaMensual.SetFocus
        feFlujoCajaMensual.row = 1
        feFlujoCajaMensual.col = 4
        SendKeys "{F2}"
    End If
End Sub
Private Sub fePasivos_RowColChange()
    If fePasivos.col = 2 Then
        fePasivos.AvanceCeldas = Horizontal
    ElseIf fePasivos.col = 3 Then
        fePasivos.AvanceCeldas = Vertical
    End If
    
    If fePasivos.col = 2 Or fePasivos.col = 3 Then
        Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6))
            Case 104, 107, 109, 201, 203
              fePasivos.ListaControles = "0-0-1-1-0-0-0"
            Case Else
              fePasivos.ListaControles = "0-0-0-0-0-0-0"
        End Select
    End If

   Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6)) 'celda que  o se puede editar
        Case 1000, 1001, 1002, 100, 200, 300, 400, 500
            Me.fePasivos.BackColorRow (&H80000000)
            Me.fePasivos.ForeColorRow vbBlack, True
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 206
            Me.fePasivos.BackColorRow vbWhite, True
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 301
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 104, 107, 109, 201, 203 'Color Amarillo
            Me.fePasivos.BackColorRow &HC0FFFF, True
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        Case Else
            Me.fePasivos.BackColorRow (&HFFFFFF)
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
    End Select
    Call CalculoTotal(2)
End Sub
Private Sub fePasivos_EnterCell()
    If fePasivos.col = 2 Or fePasivos.col = 3 Then
        Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6))
            Case 104, 107, 109, 201, 203
              fePasivos.ListaControles = "0-0-1-1-0-0-0"
            Case Else
              fePasivos.ListaControles = "0-0-0-0-0-0-0"
        End Select
    End If

    Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6)) 'celda que  o se puede editar
        Case 1000, 1001, 1002, 100, 200, 300, 400, 500
            Me.fePasivos.BackColorRow (&H80000000)
            Me.fePasivos.ForeColorRow vbBlack, True
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 206
            Me.fePasivos.BackColorRow vbWhite, True
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 301
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 104, 107, 109, 201, 203 'Color Amarillo
            Me.fePasivos.BackColorRow &HC0FFFF, True
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        Case Else
            Me.fePasivos.BackColorRow (&HFFFFFF)
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
    End Select
Call CalculoTotal(2)
End Sub


Private Sub fePasivos_Click()
    If fePasivos.col = 2 Or fePasivos.col = 3 Then
    Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6))
        Case 104, 107, 109, 201, 203
          fePasivos.ListaControles = "0-0-1-1-0-0-0"
        Case Else
          fePasivos.ListaControles = "0-0-0-0-0-0-0"
    End Select
    End If

    Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6)) 'celda que  o se puede editar
        Case 1000, 1001, 1002, 100, 200, 300, 400, 500
            Me.fePasivos.BackColorRow (&H80000000)
            Me.fePasivos.ForeColorRow vbBlack, True
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 206
            Me.fePasivos.BackColorRow vbWhite, True
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 301
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 104, 107, 109, 201, 203 'Color Amarillo
            Me.fePasivos.BackColorRow &HC0FFFF, True
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        Case Else
            Me.fePasivos.BackColorRow (&HFFFFFF)
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
    End Select
End Sub
Private Sub fePasivos_OnCellChange(pnRow As Long, pnCol As Long)
     If IsNumeric(fePasivos.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6))
            Case 305, 306 'Valores Negativos
                 fePasivos.TextMatrix(pnRow, pnCol) = Format((CCur(fePasivos.TextMatrix(pnRow, pnCol))), "#,#0.00")
            Case Else 'Valores Positivos
                If fePasivos.TextMatrix(pnRow, pnCol) < 0 Then
                  fePasivos.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(fePasivos.TextMatrix(pnRow, pnCol))), "#,#0.00")
                End If
         End Select
    Else
        fePasivos.TextMatrix(pnRow, pnCol) = "0.00"
    End If
    If (Me.fePasivos.col = 3 And Me.fePasivos.row = 19) Then
        SSTabIngresos5.Tab = 1
        Me.feFlujoCajaMensual.SetFocus
        feFlujoCajaMensual.row = 1
        feFlujoCajaMensual.col = 4
        SendKeys "{TAB}"
        SendKeys "{F2}"
    End If
    
    Call CalculoTotal(2)
End Sub
'Para Buscar Cuotas IFIs (Flujo Caja)**********->
Private Sub feFlujoCajaMensual_KeyPress(KeyAscii As Integer)
    'If (Me.feFlujoCajaMensual.Col = 4 And Me.feFlujoCajaMensual.row = 20) Then 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
    'If (Me.feFlujoCajaMensual.Col = 4 And Me.feFlujoCajaMensual.row = 22) Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Comentó
    If (Me.feFlujoCajaMensual.col = 4 And Me.feFlujoCajaMensual.row = 23) Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Agregó
        SSTabIngresos5.Tab = 1
        Me.feGastosFamiliares.SetFocus
        feGastosFamiliares.row = 1
        feGastosFamiliares.col = 3
        SendKeys "{Enter}"
    End If
End Sub
Private Sub feFlujoCajaMensual_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    'Call CalculoTotal(3)
    Dim Editar() As String
    Editar = Split(Me.feFlujoCajaMensual.ColumnasAEditar, "-")
    If Me.feFlujoCajaMensual.row <> 1 Then
        If Editar(pnCol) = "X" Then
            MsgBox "Esta celda no es editable", vbInformation, "Aviso"
            Cancel = False
            SendKeys "{TAB}", True
            Exit Sub
        End If
    End If
End Sub

Private Sub feFlujoCajaMensual_Click() 'GastosNegocio
    If feFlujoCajaMensual.col = 4 Then
        If CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 0)) = gCodCuotaIfiFlujoCaja _
            Or (CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 0)) = gCodCuotaIfiNoSupervisadaFlujoCaja) Then 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaFlujoCaja
            feFlujoCajaMensual.ListaControles = "0-0-0-0-1-0"
        Else
            feFlujoCajaMensual.ListaControles = "0-0-0-0-0-0"
        End If
    End If
        
        Select Case CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2)) 'celda que  o se puede editar
            'Case 4, 5, 20, 1000 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            'Case 5, 6, 22, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Comentó
            Case 5, 6, 23, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Agregó
                Me.feFlujoCajaMensual.BackColorRow (&H80000000)
                Me.feFlujoCajaMensual.ForeColorRow vbBlack, True
                Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
            'Case 17 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            Case 18 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                Me.feFlujoCajaMensual.ForeColorRow (&H80000007)
                Me.feFlujoCajaMensual.BackColorRow vbWhite, True
                Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
            'Case 18 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            Case 19, gCodCuotaIfiNoSupervisadaFlujoCaja 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiNoSupervisadaFlujoCaja
                Me.feFlujoCajaMensual.BackColorRow &HC0FFFF, True
                Me.feFlujoCajaMensual.ForeColorRow (&H80000007), True
                Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
            Case Else
                Me.feFlujoCajaMensual.BackColorRow (&HFFFFFF)
                Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
        End Select
End Sub

Private Sub feFlujoCajaMensual_EnterCell() 'LUCV20160525 - Me permite Buscar OtrasCuotasIFIs (GastosNegocio)
    If feFlujoCajaMensual.col = 4 Then
        If CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 0)) = gCodCuotaIfiFlujoCaja _
            Or (CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 0)) = gCodCuotaIfiNoSupervisadaFlujoCaja) Then 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaFlujoCaja
            feFlujoCajaMensual.ListaControles = "0-0-0-0-1-0"
        Else
            feFlujoCajaMensual.ListaControles = "0-0-0-0-0-0"
        End If
    End If
    
    Select Case CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2)) 'celda que  o se puede editar
        'Case 4, 5, 20, 1000 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
        'Case 5, 6, 22, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja'CTI320200110 ERS003-2020. Comentó
        Case 5, 6, 23, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Agregó
            Me.feFlujoCajaMensual.BackColorRow (&H80000000)
            Me.feFlujoCajaMensual.ForeColorRow vbBlack, True
             Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
        'Case 17 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case 18 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            Me.feFlujoCajaMensual.ForeColorRow (&H80000007)
            Me.feFlujoCajaMensual.BackColorRow vbWhite, True
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
        'Case 18 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case 19, gCodCuotaIfiNoSupervisadaFlujoCaja 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiNoSupervisadaFlujoCaja
            Me.feFlujoCajaMensual.BackColorRow &HC0FFFF, True
            Me.feFlujoCajaMensual.ForeColorRow (&H80000007), True
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
        Case Else
            Me.feFlujoCajaMensual.BackColorRow (&HFFFFFF)
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
    End Select
End Sub
Private Sub feFlujoCajaMensual_RowColChange() 'PresionarEnter:Monto
    If feFlujoCajaMensual.col = 4 Then
        feFlujoCajaMensual.AvanceCeldas = Vertical
    Else
        feFlujoCajaMensual.AvanceCeldas = Horizontal
    End If
    
    If feFlujoCajaMensual.col = 4 Then
        If CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 0)) = gCodCuotaIfiFlujoCaja _
            Or (CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 0)) = gCodCuotaIfiNoSupervisadaFlujoCaja) Then 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaFlujoCaja
            feFlujoCajaMensual.ListaControles = "0-0-0-0-1-0"
        Else
            feFlujoCajaMensual.ListaControles = "0-0-0-0-0-0"
        End If
    End If
    
    Select Case CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2)) 'celda que  o se puede editar
        'Case 4, 5, 20, 1000 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
        'Case 5, 6, 22, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Comentó
        Case 5, 6, 23, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Agregó
            Me.feFlujoCajaMensual.BackColorRow (&H80000000)
            Me.feFlujoCajaMensual.ForeColorRow vbBlack, True
             Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
        'Case 17 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case 18 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            Me.feFlujoCajaMensual.ForeColorRow (&H80000007)
            Me.feFlujoCajaMensual.BackColorRow vbWhite, True
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
        'Case 18 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
         Case 19, gCodCuotaIfiNoSupervisadaFlujoCaja 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiNoSupervisadaFlujoCaja
            Me.feFlujoCajaMensual.BackColorRow &HC0FFFF, True
            Me.feFlujoCajaMensual.ForeColorRow (&H80000007), True
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
        Case Else
            Me.feFlujoCajaMensual.BackColorRow (&HFFFFFF)
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
    End Select
End Sub
Private Sub feFlujoCajaMensual_OnClickTxtBuscar(psMontoIfiGastoNego As String, psDescripcion As String) 'Fujo Caja Mensual
    psMontoIfiGastoNego = 0
    psDescripcion = ""
    psDescripcion = feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 3) 'Cuotas Otras IFIs
    psMontoIfiGastoNego = feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 4) 'Monto
    
    If feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2) = gCodCuotaIfiFlujoCaja Then 'CTI320200110 ERS003-2020. Agregó
        If psMontoIfiGastoNego = 0 Then
            fnTotalRefGastoNego = 0
            Set MatIfiGastoNego = Nothing
            frmCredFormEvalCuotasIfis.Inicio (CLng(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 4))), fnTotalRefGastoNego, MatIfiGastoNego, feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 3), gFormatoGastosNegoFlujoCaja, gCodCuotaIfiFlujoCaja
            psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 4))), fnTotalRefGastoNego, MatIfiGastoNego, feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 3), gFormatoGastosNegoFlujoCaja, gCodCuotaIfiFlujoCaja
            psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
        End If
    Else
        If psMontoIfiGastoNego = 0 Then
            fnTotalRefGastoNego = 0
            Set MatIfiNoSupervisadaGastoNego = Nothing
            frmCredFormEvalCuotasIfis.Inicio (CLng(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 4))), fnTotalRefGastoNego, MatIfiNoSupervisadaGastoNego, feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 3), _
                                              gFormatoGastosNegoFlujoCaja, gCodCuotaIfiNoSupervisadaFlujoCaja 'CTI320200110 ERS003-2020. Agregó
            psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 4))), fnTotalRefGastoNego, MatIfiNoSupervisadaGastoNego, feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 3), _
                                              gFormatoGastosNegoFlujoCaja, gCodCuotaIfiNoSupervisadaFlujoCaja 'CTI320200110 ERS003-2020. Agregó
            psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
        End If
    End If
End Sub
Private Sub feFlujoCajaMensual_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feFlujoCajaMensual.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feFlujoCajaMensual.TextMatrix(pnRow, pnCol) < 0 Then
            feFlujoCajaMensual.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feFlujoCajaMensual.TextMatrix(pnRow, pnCol) = 0
    End If
    
    'If (Me.feFlujoCajaMensual.Col = 4 And Me.feFlujoCajaMensual.row = 19) Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    If (Me.feFlujoCajaMensual.col = 4 And Me.feFlujoCajaMensual.row = 21) Then 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
        SSTabIngresos5.Tab = 1
        Me.feGastosFamiliares.SetFocus
        feGastosFamiliares.row = 1
        feGastosFamiliares.col = 3
        SendKeys "{Enter}"
    End If
    
    Call CalculoTotal(3)
    Call CalculoTotal(4)
End Sub
Private Sub feGastosFamiliares_KeyPress(KeyAscii As Integer)
        If (feGastosFamiliares.col = 1 And feGastosFamiliares.row = 1) Or (feGastosFamiliares.col = 3 And feGastosFamiliares.row = 7) Then
        If KeyAscii = 13 Then
            feOtrosIngresos.row = 1
            feOtrosIngresos.col = 3
            EnfocaControl feOtrosIngresos
            SendKeys "{Enter}", True
        End If
    End If
End Sub
Private Sub feGastosFamiliares_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(Me.feGastosFamiliares.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}", True
        Exit Sub
    End If
End Sub
Private Sub feGastosFamiliares_Click() 'GastosFamiliares
    If feGastosFamiliares.col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami _
            Or CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiNoSupervisadaGastoFami Then 'CTI320200110 ERS003-2020, Agregó
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If

        Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
            Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
                Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
                Me.feGastosFamiliares.ForeColorRow (&H80000007), True
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
            Case gCodDeudaLCNUGastoFami
                Me.feGastosFamiliares.BackColorRow vbWhite, True
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
            Case Else
                Me.feGastosFamiliares.BackColorRow (&HFFFFFF)
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
        End Select
End Sub
Private Sub feGastosFamiliares_EnterCell() 'LUCV20160525 - Me permite Buscar CuotasIFIs(GastosFamiliares)
    If feGastosFamiliares.col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami _
            Or CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiNoSupervisadaGastoFami Then 'CTI320200110 ERS003-2020, Agregó
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If

    Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
        Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
            Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
            Me.feGastosFamiliares.ForeColorRow (&H80000007), True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
        Case gCodDeudaLCNUGastoFami
            Me.feGastosFamiliares.BackColorRow vbWhite, True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
        Case Else
            Me.feGastosFamiliares.BackColorRow (&HFFFFFF)
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End Select
End Sub
Private Sub feGastosFamiliares_RowColChange() 'PresionarEnter:Monto
    If feGastosFamiliares.col = 3 Then
        feGastosFamiliares.AvanceCeldas = Vertical
    Else
        feGastosFamiliares.AvanceCeldas = Horizontal
    End If
    
    If feGastosFamiliares.col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 0)) = gCodCuotaIfiGastoFami _
            Or (CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 0)) = gCodCuotaIfiNoSupervisadaGastoFami) Then 'CTI320200110 ERS003-2020, Agregó
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If
    
        Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
        Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
            Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
            Me.feGastosFamiliares.ForeColorRow (&H80000007), True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
        Case gCodDeudaLCNUGastoFami
            Me.feGastosFamiliares.BackColorRow vbWhite, True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
        Case Else
            Me.feGastosFamiliares.BackColorRow (&HFFFFFF)
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
        End Select
End Sub
Private Sub feGastosFamiliares_OnClickTxtBuscar(psMontoIfiGastoFami As String, psDescripcion As String) 'GastosFamiliares
    psMontoIfiGastoFami = 0
    psDescripcion = ""
    psDescripcion = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2) 'Cuotas Otras IFIs
    psMontoIfiGastoFami = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3) 'Monto
    
    If CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami Then 'CTI320200110 ERS003-2020. Agregó
        If psMontoIfiGastoFami = 0 Then
            fnTotalRefGastoFami = 0
            Set MatIfiGastoFami = Nothing
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), gFormatoGastosFami, gCodCuotaIfiGastoFami 'CTI320200110 ERS003-2020. Agregó
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), gFormatoGastosFami, gCodCuotaIfiGastoFami 'CTI320200110 ERS003-2020. Agregó
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        End If
    Else
        If psMontoIfiGastoFami = 0 Then
            fnTotalRefGastoFami = 0
            Set MatIfiNoSupervisadaGastoFami = Nothing
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiNoSupervisadaGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), _
                                             gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agregó
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiNoSupervisadaGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), _
                                            gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami  'CTI320200110 ERS003-2020. Agregó
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        End If
    End If
End Sub
Private Sub feGastosFamiliares_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feGastosFamiliares.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feGastosFamiliares.TextMatrix(pnRow, pnCol) < 0 Then
            feGastosFamiliares.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feGastosFamiliares.TextMatrix(pnRow, pnCol) = 0
    End If
End Sub
Private Sub OptCondLocal_Click(Index As Integer)
    Select Case Index
    Case 1, 2, 3
        Me.txtCondLocalOtros.Visible = False
        Me.txtCondLocalOtros.Text = ""
    Case 4
        Me.txtCondLocalOtros.Visible = True
        Me.txtCondLocalOtros.Text = ""
    End Select
    fnCondLocal = Index
End Sub

'***** LUCV20160528 - OnCellChange / RowColChange
Private Sub feReferidos_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 1 Or pnCol = 4 Then
        feReferidos.TextMatrix(pnRow, pnCol) = UCase(feReferidos.TextMatrix(pnRow, pnCol))
    End If
    
    Select Case pnCol
    Case 2
        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidos.TextMatrix(pnRow, pnCol)) = 8 Then
                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
                    Case Is > 0
                    Case Else
                        MsgBox "Por favor, verifique el DNI", vbInformation, "Alerta"
                        feReferidos.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "El DNI, tiene que ser 8 dígitos.", vbInformation, "Alerta"
                feReferidos.TextMatrix(pnRow, pnCol) = 0
            End If
            
        Else
            MsgBox "El DNI, tiene que ser numérico.", vbInformation, "Alerta"
            feReferidos.TextMatrix(pnRow, pnCol) = 0
        End If
    Case 3
        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidos.TextMatrix(pnRow, pnCol)) = 9 Then
                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
                Case Is > 0
                Case Else
                    MsgBox "Teléfono Mal Ingresado", vbInformation, "Alerta"
                    feReferidos.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "Faltan caracteres en el teléfono / celular.", vbInformation, "Alerta"
                feReferidos.TextMatrix(pnRow, pnCol) = 0
            End If
        Else
            MsgBox "El telefono, solo permite ingreso de datos tipo numérico." & Chr(10) & "Ejemplo: 065404040, 984047523 ", vbInformation, "Alerta"
            feReferidos.TextMatrix(pnRow, pnCol) = 0
        End If
    Case 5
        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidos.TextMatrix(pnRow, pnCol)) = 8 Then
                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
                Case Is > 0
                Case Else
                    MsgBox "El DNI del referido, tiene que contener 8 dígitos", vbInformation, "Alerta"
                    feReferidos.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "El DNI del referido, tiene que ser 8 dígitos", vbInformation, "Alerta"
                feReferidos.TextMatrix(pnRow, pnCol) = 0
            End If
        Else
            MsgBox "El DNI del referido, sólo permite ingreso de datos tipo numérico.", vbInformation, "Alerta"
            feReferidos.TextMatrix(pnRow, pnCol) = 0
        End If
    End Select
End Sub

Private Sub feReferidos_RowColChange()
    If feReferidos.col = 1 Then
        feReferidos.MaxLength = "200"
    ElseIf feReferidos.col = 2 Then
        feReferidos.MaxLength = "8"
    ElseIf feReferidos.col = 3 Then
        feReferidos.MaxLength = "9"
    ElseIf feReferidos.col = 4 Then
        feReferidos.MaxLength = "200"
    ElseIf feReferidos.col = 5 Then
        feReferidos.MaxLength = "8"
    End If
End Sub

Private Sub feOtrosIngresos_RowColChange() 'PresionarEnter:Monto
    If feOtrosIngresos.col = 3 Then
        feOtrosIngresos.AvanceCeldas = Vertical
    Else
        feOtrosIngresos.AvanceCeldas = Horizontal
    End If
End Sub
Private Sub feOtrosIngresos_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feOtrosIngresos.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feOtrosIngresos.TextMatrix(pnRow, pnCol) < 0 Then
            feOtrosIngresos.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feOtrosIngresos.TextMatrix(pnRow, pnCol) = 0
    End If
    If Me.feOtrosIngresos.col = 3 And Me.feOtrosIngresos.row = 5 Then
        Me.SSTabIngresos5.Tab = 1
        Me.feDeclaracionPDT.SetFocus
        Me.feDeclaracionPDT.row = 1
        Me.feDeclaracionPDT.col = 4
        SendKeys "{TAB}"
   End If
    
End Sub
Private Sub feDeclaracionPDT_KeyPress(KeyAscii As Integer)
 If (Me.feDeclaracionPDT.col = 7 And Me.feDeclaracionPDT.row = 2) Or (Me.feDeclaracionPDT.col = 8 And Me.feDeclaracionPDT.row = 2) Then
        Me.SSTabIngresos5.Tab = 2
        If txtFechaVisita.Enabled Then 'ARLO20190330
            Me.txtFechaVisita.SetFocus
        End If
        SendKeys "{TAB}"
   End If
End Sub
Private Sub feDeclaracionPDT_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
'CalculoTotal (4)
End Sub
Private Sub feDeclaracionPDT_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feDeclaracionPDT.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feDeclaracionPDT.TextMatrix(pnRow, pnCol) < 0 Then
            feDeclaracionPDT.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feDeclaracionPDT.TextMatrix(pnRow, pnCol) = 0
    End If
    If Me.feDeclaracionPDT.col = 6 And Me.feDeclaracionPDT.row = 2 Then
        Me.SSTabIngresos5.Tab = 2
        If txtFechaVisita.Enabled Then  'ARLO20190330
            Me.txtFechaVisita.SetFocus
        End If
        SendKeys "{TAB}"
   End If
    Call CalculoTotal(4)
End Sub

Private Sub feDeclaracionPDT_Click()
    CalculoTotal (4)
End Sub


'Fin <- LUCV20160528 - OnCellChange / RowColChange *****

'________________________________________________________________________________________________________________________
'*************************************************LUCV20160525: METODOS Varios **************************************************
Public Function Inicio(ByVal psTipoRegMant As Integer, ByVal psCtaCod As String, ByVal pfnFormato As Integer, ByVal pnProducto As Integer, _
                     ByVal pnSubProducto As Integer, ByVal pnMontoExpEsteCred As Double, ByVal pbImprimir As Boolean, ByVal pnEstado As Integer, _
                     Optional ByVal pbImprimirVB As Boolean = False) As Boolean
                  
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsDCredEval As ADODB.Recordset
    Dim rsDColCred As ADODB.Recordset
    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
    fnProducto = pnProducto 'CTI320200110 ERS003-2020
    fbImprimirVB = pbImprimirVB 'CTI3ERS0032020
    If psCtaCod <> -1 Then '*****-> CtaCod
        gsOpeCod = ""
        lcMovNro = "" 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        nTC = oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia)
        sCtaCod = psCtaCod
        fnTipoRegMant = psTipoRegMant
        ActXCodCta.NroCuenta = sCtaCod
        
        '(3: Analista, 2: Coordinador, 1: JefeAgencia)
        fnTipoPermiso = oNCOMFormatosEval.ObtieneTipoPermisoCredEval(gsCodCargo)  ' Obtener el tipo de Permiso, Segun Cargo
        Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval

        Set rsDCredito = oDCOMFormatosEval.RecuperaSolicitudDatoBasicosEval(sCtaCod) ' Datos Basicos del Credito Solicitado
        
        If (rsDCredito!cActiGiro) = "" Then
            MsgBox "Por favor, actualizar los datos del cliente. " & Chr(13) & " (Actividad o Giro del negocio)", vbInformation, "Alerta"
            Exit Function
        End If
        
        '*****-> Datos básicos de cabecera de Formato
        fsGiroNego = IIf((rsDCredito!cActiGiro) = "", "", (rsDCredito!cActiGiro))
        fsCliente = Trim(rsDCredito!cPersNombre)
        fsAnioExp = CInt(rsDCredito!nAnio)
        fsMesExp = CInt(rsDCredito!nMes)
        fnColocCondi = rsDCredito!nColocCondicion
        fbTieneReferido6Meses = rsDCredito!bTieneReferido6Meses   'Si tiene evaluacion registrada 6 meses (LUCV20171115, agregó según correo: RUSI)
        fnFechaDeudaSbs = IIf(rsDCredito!dFechaUltimaDeudaSBS = "", "__/__/____", rsDCredito!dFechaUltimaDeudaSBS)
        fnMontoDeudaSbs = CCur(rsDCredito!nMontoUltimaDeudaSBS)
        
        spnExpEmpAnio.valor = fsAnioExp
        spnExpEmpMes.valor = fsMesExp
        txtUltEndeuda.Text = Format(fnMontoDeudaSbs, "#,##0.00")
        txtFecUltEndeuda.Text = Format(fnFechaDeudaSbs, "dd/mm/yyyy")
        txtExposicionCredito.Text = Format(pnMontoExpEsteCred, "#,##0.00")
        txtFechaEvaluacion.Text = Format(gdFecSis, "dd/mm/yyyy")
        '<-***** Fin datos de cabecera
       
        Set rsDCredEval = oDCOMFormatosEval.RecuperaColocacCredEval(sCtaCod) 'Ojo: Recuperar Credito Si ha sido Registrado el Form. Eval.
        Set rsAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(sCtaCod) 'Obtenemos Datos, Aceptable / Critico de los Ratios
        If fnTipoPermiso = 2 Then
           If rsDCredEval.RecordCount = 0 Then ' Si no hay credito registrado
                MsgBox "El analista no ha registrado la Evaluacion respectiva", vbInformation, "Aviso"
                fbPermiteGrabar = False
            Else
                fbPermiteGrabar = True
             End If
        End If
        Set rsDCredito = Nothing
        Set rsDCredEval = Nothing
        
        fnFormato = pfnFormato
        fnEstado = pnEstado
        'CTI320200110 ERS003-2020
        If fnProducto <> "800" Then
            SSTabIngresos5.Tab = 0
        Else
            SSTabIngresos5.Tab = 1
        End If
        'fin CTI320200110
        
        fbPermiteGrabar = False
        fbBloqueaTodo = False
        'frameLinea.Visible = False 'CTI320200110 ERS003-2020. Comentó
        SSTabRatios.Visible = False
        
    Else
        MsgBox "No se ha registrado el número de cuenta del crédito a evaluar ", vbInformation, "Aviso"
    End If
    'Fin cCtaCod <-**********
    
    Set oDCOMFormatosEval = Nothing
    Set oTipoCam = Nothing
    Call CargaControlesInicio

    If fnTipoRegMant = 3 Then
        fbBloqueaTodo = True
        'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        gsOpeCod = gCredConsultarEvaluacionCred
        lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gConsultar, "Evaluacion Credito Formato 5", sCtaCod, gCodigoCuenta
        Set objPista = Nothing
        'Fin LUCV20181220
    End If

    'Carga de Datos Segun Evento: (Registrar / Mantenimiento) *****->
    If CargaDatos Then
        If CargaControlesTipoPermiso(fnTipoPermiso, fbPermiteGrabar, fbBloqueaTodo) Then
            If fnTipoRegMant = 1 Then   'Para el Evento: "Registrar"
                If Not rsCredEval.EOF Then
                    Call Mantenimiento
                    fnTipoRegMant = 2
                Else
                    Call Registro
                    fnTipoRegMant = 1
                End If
            ElseIf fnTipoRegMant = 2 Then 'Para el Evento. "Mantenimiento"
                If rsCredEval.EOF Then
                    Call Registro
                    fnTipoRegMant = 1
                Else
                    Call Mantenimiento
                    fnTipoRegMant = 2
                End If
            ElseIf fnTipoRegMant = 3 Then  ' Para el Evento. "Consulta"
                    Call Mantenimiento
                    fnTipoRegMant = 3
            End If
        Else
            Unload Me
            Exit Function
        End If
    Else
        If CargaControlesTipoPermiso(1, False) Then
        End If
    End If
    'Fin Carga <-*****
    
    
    'Habilita / Deshabilita Botones - Text
    If fnEstado = 2000 Then             '*****-> Si es Solicitado
            'Me.cmdGuardar.Enabled = True
            Me.cmdImprimir.Enabled = False
            Me.cmdInformeVisita.Enabled = False
            cmdFlujoCaja5.Enabled = False 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            If fnColocCondi <> 4 Then
                Me.cmdVerCar.Enabled = False
            Else
                Me.cmdVerCar.Enabled = False
            End If
        Else                                '*****-> Sugerido +
            'Me.cmdGuardar.Enabled = True
            Me.cmdImprimir.Enabled = True
            cmdFlujoCaja5.Enabled = True 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            If fnColocCondi <> 4 Then
                Me.cmdVerCar.Enabled = True 'No refinanciado
                Me.cmdInformeVisita.Enabled = True
            Else
                Me.cmdVerCar.Enabled = False
                Me.cmdInformeVisita.Enabled = False
            End If
        End If
    '*****->No Refinanciados (Propuesta Credito)
    If fnColocCondi <> 4 Then
        txtFechaVisita.Enabled = True
        txtEntornoFamiliar5.Enabled = True
        txtGiroUbicacion5.Enabled = True
        txtExperiencia5.Enabled = True
        txtFormalidadNegocio5.Enabled = True
        txtColaterales5.Enabled = True
        txtDestino5.Enabled = True
     Else
        framePropuesta.Enabled = False
        txtFechaVisita.Enabled = False
        txtEntornoFamiliar5.Enabled = False
        txtGiroUbicacion5.Enabled = False
        txtExperiencia5.Enabled = False
        txtFormalidadNegocio5.Enabled = False
        txtColaterales5.Enabled = False
        txtDestino5.Enabled = False
    End If
    '*****->Fin No Refinanciados
    
    Set rsAceptableCritico = Nothing
    fbGrabar = False
    Call CalculoTotal(2)
    If Not pbImprimir Then
        If fbImprimirVB Then
            Call Cmdguardar_Click
            fbImprimirVB = False
            cmdGuardar.Enabled = True
        End If
        Me.Show 1
    Else
        cmdImprimir_Click
    End If
    Inicio = fbGrabar
End Function

Private Function DevolverMes(ByVal pnMes As Integer, ByRef pnAnio As Integer, ByRef pnMesN As Integer) As String 'Cargar Ultimo 3 Meses -> Registrar
    Dim nIndMes As Integer
    nIndMes = CInt(Mid(gdFecSis, 4, 2)) - pnMes
    pnAnio = CInt(Mid(gdFecSis, 7, 4))
    If nIndMes < 1 Then
        nIndMes = nIndMes + 12
        pnAnio = pnAnio - 1
    End If
        pnMesN = nIndMes
        DevolverMes = Choose(nIndMes, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
End Function

Private Function DevolverMesDatos(ByVal pnMes As Integer) As String 'Cargar 3 Ultimos Meses -> Para el Mantenimiento
    DevolverMesDatos = Choose(pnMes, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
End Function

'***** LUCV20160529 / feReferidos
Public Function ValidaDatosReferencia() As Boolean
    Dim i As Integer, j As Integer
    ValidaDatosReferencia = False
        If feReferidos.rows - 1 < 2 Then
            MsgBox "Debe registrar por lo menos 2 referencias para continuar", vbInformation, "Aviso"
            cmdAgregarRef.SetFocus
            ValidaDatosReferencia = False
            Exit Function
        End If
        For i = 1 To feReferidos.rows - 1  'Verfica Tipo de Valores del DNI
            If Trim(feReferidos.TextMatrix(i, 1)) <> "" Then
                For j = 1 To Len(Trim(feReferidos.TextMatrix(i, 2)))
                    If (Mid(feReferidos.TextMatrix(i, 2), j, 1) < "0" Or Mid(feReferidos.TextMatrix(i, 2), j, 1) > "9") Then
                       MsgBox "Uno de los Digitos del primer DNI de la fila " & i & " no es un Numero", vbInformation, "Aviso"
                       feReferidos.SetFocus
                       ValidaDatosReferencia = False
                       Exit Function
                    End If
                Next j
            End If
        Next i
        For i = 1 To feReferidos.rows - 1  'Verfica Longitud del DNI
            If Trim(feReferidos.TextMatrix(i, 1)) <> "" Then
                If Len(Trim(feReferidos.TextMatrix(i, 2))) <> gnNroDigitosDNI Then
                    MsgBox "Primer DNI de la fila " & i & " no es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                    feReferidos.SetFocus
                    ValidaDatosReferencia = False
                    Exit Function
                End If
            End If
        Next i
        For i = 1 To feReferidos.rows - 1  'Verfica Tipo de Valores del Telefono
            If Trim(feReferidos.TextMatrix(i, 1)) <> "" Then
                For j = 1 To Len(Trim(feReferidos.TextMatrix(i, 3)))
                    If (Mid(feReferidos.TextMatrix(i, 3), j, 1) < "0" Or Mid(feReferidos.TextMatrix(i, 3), j, 1) > "9") Then
                       MsgBox "Uno de los Digitos del teléfono de la fila " & i & " no es un Numero", vbInformation, "Aviso"
                       feReferidos.SetFocus
                       ValidaDatosReferencia = False
                       Exit Function
                    End If
                Next j
            End If
        Next i

        For i = 1 To feReferidos.rows - 1
            For j = 1 To feReferidos.rows - 1
                If i <> j Then
                    If feReferidos.TextMatrix(i, 2) = feReferidos.TextMatrix(j, 2) Then
                        MsgBox "No se puede ingresar el mismo DNI mas de una vez en los referidos", vbInformation, "Alerta"
                        ValidaDatosReferencia = False
                        Exit Function
                    End If
                End If
            Next
        Next
    ValidaDatosReferencia = True
End Function

Public Function ValidaGrillas(ByVal Flex As FlexEdit) As Boolean
    Dim i As Integer
    ValidaGrillas = False
    For i = 1 To Flex.rows - 1
        If Flex.TextMatrix(i, 0) <> "" Then
            If Trim(Flex.TextMatrix(i, 1)) = "" Or Trim(Flex.TextMatrix(i, 3)) = "" Then
                ValidaGrillas = False
                Exit Function
            End If
        End If
    Next i
    ValidaGrillas = True
End Function

Public Function ValidaDatos() As Boolean
    Dim nIndice As Integer
    ValidaDatos = False
    Dim lsMensajeIfi As String 'LUCV20161115
    If fnTipoPermiso = 3 Then
    '********** Para TAB:0 -> Ingresos y Egresos
        If spnTiempoLocalAnio.valor = "" Then
            MsgBox "Ingrese Tiempo en el mismo local: Años", vbInformation, "Aviso"
            'CTI320200110 ERS003-2020
            If fnProducto <> "800" Then
                SSTabIngresos5.Tab = 0
            Else
                SSTabIngresos5.Tab = 1
            End If
            'fin CTI320200110
            
            spnTiempoLocalAnio.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If spnTiempoLocalMes.valor = "" Then
        MsgBox "Ingrese Tiempo en el mismo local: Meses", vbInformation, "Aviso"
            'CTI320200110 ERS003-2020
            If fnProducto <> "800" Then
                SSTabIngresos5.Tab = 0
            Else
                SSTabIngresos5.Tab = 1
            End If
            'fin CTI320200110
            spnTiempoLocalMes.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If OptCondLocal(1).value = 0 And OptCondLocal(2).value = 0 And OptCondLocal(3).value = 0 And OptCondLocal(4).value = 0 Then
            MsgBox "Falta elegir la Condicion del local", vbInformation, "Aviso"
            ValidaDatos = False
            'CTI320200110 ERS003-2020
            If fnProducto <> "800" Then
                SSTabIngresos5.Tab = 0
            Else
                SSTabIngresos5.Tab = 1
            End If
            'fin CTI320200110
            Exit Function
        End If
        If txtCondLocalOtros.Visible = True Then
            If txtCondLocalOtros.Text = "" Then
            MsgBox "Ingrese la Descripcion de la Opcion: Otro Local", vbInformation, "Aviso"
                ValidaDatos = False
                'CTI320200110 ERS003-2020
                If fnProducto <> "800" Then
                    SSTabIngresos5.Tab = 0
                Else
                    SSTabIngresos5.Tab = 1
                End If
                'fin CTI320200110
                Exit Function
            End If
        End If
        If Trim(txtGiroNeg.Text) = "" Then
            MsgBox "Falta ingresar el Giro del Negocio, Favor Actualizar los Datos del Cliente", vbInformation, "Aviso"
            'CTI320200110 ERS003-2020
            If fnProducto <> "800" Then
                SSTabIngresos5.Tab = 0
            Else
                SSTabIngresos5.Tab = 1
            End If
            'fin CTI320200110
            txtGiroNeg.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtFechaEvaluacion.Text) = "__/__/____" Then
            MsgBox "Falta Ingresar la Fecha de Evaluacion", vbInformation, "Aviso"
            'CTI320200110 ERS003-2020
            If fnProducto <> "800" Then
                SSTabIngresos5.Tab = 0
            Else
                SSTabIngresos5.Tab = 1
            End If
            'fin CTI320200110
            txtFechaEvaluacion.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    '********** Para TAB:1 -> Propuesta del Credito
      If fnColocCondi <> 4 Then 'Valida, si el credito no es refinanciado
        If Trim(txtFechaVisita.Text) = "__/__/____" Or Not IsDate(Trim(txtFechaVisita.Text)) Then
            MsgBox "Falta ingresar la fecha de visita o el formato de la fecha no es el correcto." & Chr(10) & " Formato: DD/MM/YYY", vbInformation, "Aviso"
            SSTabIngresos5.Tab = 2
            txtFechaVisita.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If txtEntornoFamiliar5.Text = "" Then
            MsgBox "Por favor Ingrese, El Entorno Familiar del Cliente o Representante", vbInformation, "Aviso"
            SSTabIngresos5.Tab = 2
            txtEntornoFamiliar5.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If txtGiroUbicacion5.Text = "" Then
            MsgBox "Por favor Ingrese, El Giro y la Ubicacion del Negocio", vbInformation, "Aviso"
            SSTabIngresos5.Tab = 2
            txtGiroUbicacion5.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If txtExperiencia5.Text = "" Then
            MsgBox "Por favor Ingrese, Sobre la Experiencia Crediticia", vbInformation, "Aviso"
            SSTabIngresos5.Tab = 2
            txtExperiencia5.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If txtFormalidadNegocio5.Text = "" Then
            MsgBox "Por favor Ingrese, La Formalidad del Negocio", vbInformation, "Aviso"
            SSTabIngresos5.Tab = 2
            txtFormalidadNegocio5.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If txtColaterales5.Text = "" Then
            MsgBox "Por favor Ingrese, Sobre las Garantias y Colaterales", vbInformation, "Aviso"
            SSTabIngresos5.Tab = 2
            txtColaterales5.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If txtDestino5.Text = "" Then
            MsgBox "Por favor Ingrese, El destino del Credito", vbInformation, "Aviso"
            SSTabIngresos5.Tab = 2
            txtDestino5.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
    '********** Para TAB:3 -> Comentarios y referidos
        'LUCV25072016->*****, Si el cliente es Nuevo -> Referente es Obligatorio
            'If fnColocCondi = 1 Then 'LUCV2017115, Según correo: RUSI
            If Not fbTieneReferido6Meses Then
                frameReferido.Enabled = True
                frameComentario.Enabled = True
                    For i = 0 To feReferidos.rows - 1
                        If feReferidos.TextMatrix(i, 0) <> "" Then
                            If Trim(feReferidos.TextMatrix(i, 0)) = "" Or Trim(feReferidos.TextMatrix(i, 1)) = "" _
                                Or Trim(feReferidos.TextMatrix(i, 2)) = "" Or Trim(feReferidos.TextMatrix(i, 3)) = "" Or Trim(feReferidos.TextMatrix(i, 4)) = "" Then
                                MsgBox "Faltan datos en la lista de Referencias", vbInformation, "Aviso"
                                SSTabIngresos5.Tab = 3
                                ValidaDatos = False
                                Exit Function
                            End If
                        End If
                    Next i
            
                    If ValidaDatosReferencia = False Then 'Contenido de feReferidos2: Referidos
                        SSTabIngresos5.Tab = 3
                        ValidaDatos = False
                        Exit Function
                    End If
                    
                    If txtComentario5.Text = "" Then
                        MsgBox "Por favor Ingrese, Comentarios", vbInformation, "Aviso"
                        SSTabIngresos5.Tab = 3
                        txtComentario5.SetFocus
                        ValidaDatos = False
                        Exit Function
                    End If
            Else
                'si el cliente es nuevo-> referido obligatorio
                    frameReferido.Enabled = False
                    feReferidos.Enabled = False
                    cmdAgregarRef.Enabled = False
                    cmdQuitar5.Enabled = False
                    txtComentario5.Enabled = False
                    frameComentario.Enabled = False
            End If
            'Fin LUCV25072016 <-*****
            
            '********** Para TAB:0 -> Validacion Grillas: GastosNegocio, OtrosIngresos, GastosFamiliares
            If ValidaGrillas(feOtrosIngresos) = False Then
                MsgBox "Faltan datos en la lista de Otros Ingresos", vbInformation, "Aviso"
                'CTI320200110 ERS003-2020
                If fnProducto <> "800" Then
                    SSTabIngresos5.Tab = 0
                Else
                    SSTabIngresos5.Tab = 1
                End If
                'fin CTI320200110
                ValidaDatos = False
                Exit Function
            End If
            If ValidaGrillas(feGastosFamiliares) = False Then
                MsgBox "Faltan datos en la lista de Gastos Familiares", vbInformation, "Aviso"
                'CTI320200110 ERS003-2020
                If fnProducto <> "800" Then
                    SSTabIngresos5.Tab = 0
                Else
                    SSTabIngresos5.Tab = 1
                End If
                'fin CTI320200110
                ValidaDatos = False
                Exit Function
            End If
           
           '********** Para TAB:0 -> Grilla Balance General
           For nIndice = 1 To feFlujoCajaMensual.rows - 1   'Flujo de Caja Mensual - Saldo disponible
                'If feFlujoCajaMensual.TextMatrix(nIndice, 2) = 20 Then 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                'If feFlujoCajaMensual.TextMatrix(nIndice, 2) = 22 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Comentó
                If feFlujoCajaMensual.TextMatrix(nIndice, 2) = 23 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Agregó
                    If val(Replace(feFlujoCajaMensual.TextMatrix(nIndice, 4), ",", "")) <= 0 Then
                        MsgBox "Flujo de Caja Mensual: Saldo disponible = (Margen Bruto) - (Otros Egresos) " & Chr(10) & " - El saldo disponible no puede ser un valor menor que cero.", vbInformation, "Alerta"
                        ValidaDatos = False
                        SSTabIngresos5.Tab = 1
                        Exit Function
                    End If
                End If
                
                'If feFlujoCajaMensual.TextMatrix(nIndice, 2) = 3 Then 'Flujo de Caja Mensual - Egresos Compras 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                If feFlujoCajaMensual.TextMatrix(nIndice, 2) = 4 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                    If val(Replace(feFlujoCajaMensual.TextMatrix(nIndice, 4), ",", "")) <= 0 Then
                        MsgBox "Flujo de Caja Mensual: Egresos por compras " & Chr(10) & " - El Valor ingresado tiene que ser un monto mayor a cero", vbInformation, "Alerta"
                        ValidaDatos = False
                        SSTabIngresos5.Tab = 1
                        Exit Function
                    End If
                End If
            Next
           
           If fnProducto <> "800" Then 'CTI320200110 ERS003-2020
                For nIndice = 1 To feActivos.rows - 1
                    If feActivos.TextMatrix(nIndice, 6) = 100 Then 'Activo Corriente
                        If val(Replace(feActivos.TextMatrix(nIndice, 4), ",", "")) <= 0 Then
                            MsgBox "Activo Corriente: " & Chr(10) & " El total, no tiene que ser un valor menor o igual que cero. ", vbInformation, "Alerta"
                            ValidaDatos = False
                            SSTabIngresos5.Tab = 0
                            Exit Function
                        End If
                    End If
                    
                    If feActivos.TextMatrix(nIndice, 6) = 1000 Then 'Total Activo
                        If val(Replace(feActivos.TextMatrix(nIndice, 4), ",", "")) <= 0 Then
                            MsgBox "Total Activo: " & Chr(10) & " El total, no tiene que ser un valor menor o igual que cero. ", vbInformation, "Alerta"
                            ValidaDatos = False
                            SSTabIngresos5.Tab = 0
                            Exit Function
                        End If
                    End If
                Next
                
                For nIndice = 1 To fePasivos.rows - 1
                    If fePasivos.TextMatrix(nIndice, 6) = 100 Then 'Pasivo  Corriente
                        If val(Replace(fePasivos.TextMatrix(nIndice, 4), ",", "")) <= 0 Then
                            MsgBox "Pasivo Corriente: " & Chr(10) & " - El total, no tiene que ser un valor menor o igual que cero ", vbInformation, "Alerta"
                            ValidaDatos = False
                             SSTabIngresos5.Tab = 0
                            Exit Function
                        End If
                    End If
                    
    '                If fePasivos.TextMatrix(nIndice, 6) = 200 Then 'Pasivo No Corriente
    '                    If val(Replace(fePasivos.TextMatrix(nIndice, 4), ",", "")) <= 0 Then
    '                        MsgBox "Pasivo No Corriente: " & Chr(10) & " - El total, no tiene que ser un valor menor o igual que cero ", vbInformation, "Alerta"
    '                        ValidaDatos = False
    '                         SSTabIngresos5.Tab = 0
    '                        Exit Function
    '                    End If
    '                End If
                    
                    If fePasivos.TextMatrix(nIndice, 6) = 300 Then 'Patrimonio
                        If val(Replace(fePasivos.TextMatrix(nIndice, 4), ",", "")) <= 0 Then
                            MsgBox "Patrimonio: " & Chr(10) & " - El total, no tiene que ser un valor menor o igual que cero ", vbInformation, "Alerta"
                            ValidaDatos = False
                             SSTabIngresos5.Tab = 0
                             Me.fePasivos.SetFocus
                            Exit Function
                        End If
                    End If
                Next
                'LUCV20161115, Agregó->Según ERS068-2016
                If Not ValidaIfiExisteCompraDeuda(sCtaCod, MatIfiGastoFami, MatIfiGastoNego, lsMensajeIfi, MatIfiNoSupervisadaGastoFami, MatIfiNoSupervisadaGastoNego) Or Len(Trim(lsMensajeIfi)) > 0 Then
                    MsgBox "Ifi y Cuota registrada en detalle de cambio de estructura de pasivos no coincide:  " & Chr(10) & Chr(10) & " " & lsMensajeIfi & " ", vbInformation, "Aviso"
                    SSTabIngresos5.Tab = 1
                    Exit Function
                End If
            End If 'Fin CTI320200110
   End If
    ValidaDatos = True
End Function

Private Function CargaControlesTipoPermiso(ByVal TipoPermiso As Integer, ByVal pPermiteGrabar As Boolean, Optional ByVal pBloqueaTodo As Boolean = False) As Boolean
    '1: JefeAgencia->
    If TipoPermiso = 1 Then
        Call HabilitaControles(False, False, False)
        CargaControlesTipoPermiso = True
     '2: Coordinador->
    ElseIf TipoPermiso = 2 Then
        Call HabilitaControles(False, False, pPermiteGrabar)
        CargaControlesTipoPermiso = True
     '3: Analista ->
    ElseIf TipoPermiso = 3 Then
        Call HabilitaControles(True, False, True)
        CargaControlesTipoPermiso = True
     'Usuario sin Permisos al formato
    Else
        MsgBox "No tiene Permisos para este módulo", vbInformation, "Aviso"
        CargaControlesTipoPermiso = False
    End If
    
    If pBloqueaTodo Then 'Para el Caso despues de dar Verificacion
        Call HabilitaControles(True, True, False)
        CargaControlesTipoPermiso = True
    End If
End Function

Private Function HabilitaControles(ByVal pbHabilitaA As Boolean, ByVal pbHabilitaRatios As Boolean, ByVal pbHabilitaGuardar As Boolean)
'HabilitacionControlesAnalistas:     pbHabilitaA = True
    'Tab0: Ingresos/Egresos
    spnTiempoLocalAnio.Enabled = pbHabilitaA
    spnTiempoLocalMes.Enabled = pbHabilitaA
    OptCondLocal(1).Enabled = pbHabilitaA
    OptCondLocal(2).Enabled = pbHabilitaA
    OptCondLocal(3).Enabled = pbHabilitaA
    OptCondLocal(4).Enabled = pbHabilitaA
    txtCondLocalOtros.Enabled = pbHabilitaA
    'txtFechaEvaluacion.Enabled = pbHabilitaA
    feActivos.Enabled = pbHabilitaA
    fePasivos.Enabled = pbHabilitaA
    
    'Tab1:  Flujo Caja Mensual
    feFlujoCajaMensual.Enabled = pbHabilitaA
    feGastosFamiliares.Enabled = pbHabilitaA
    feOtrosIngresos.Enabled = pbHabilitaA
    feDeclaracionPDT.Enabled = pbHabilitaA
    
    'Tab2: Propuesta/Credito
    txtFechaVisita.Enabled = pbHabilitaA
    txtEntornoFamiliar5.Enabled = pbHabilitaA
    txtGiroUbicacion5.Enabled = pbHabilitaA
    txtExperiencia5.Enabled = pbHabilitaA
    txtFormalidadNegocio5.Enabled = pbHabilitaA
    txtColaterales5.Enabled = pbHabilitaA
    txtDestino5.Enabled = pbHabilitaA

    'Tab3: Comentarios/Referidos
    txtComentario5.Enabled = pbHabilitaA
    feReferidos.Enabled = pbHabilitaA
    cmdAgregarRef.Enabled = pbHabilitaA
    cmdQuitar2.Enabled = pbHabilitaA
    frameReferido.Enabled = pbHabilitaA
    
      'txtVerif.Enabled = pbHabilitaB
    If fnEstado = 2000 Then
        SSTabRatios.Visible = False
    Else
        SSTabRatios.Visible = pbHabilitaRatios
    End If

    'cmdInformeVisita.Enabled = pbHabilitaRatios
    'cmdVerCar.Enabled = pbHabilitaRatios
    'cmdImprimir.Enabled = pbHabilitaRatios
    cmdGuardar.Enabled = pbHabilitaGuardar
End Function
Private Sub CargaControlesInicio()
    Call CargarFlexEdit
        
    'DesHabilita la CargaInicial de Controles
    ActXCodCta.Enabled = False
    txtNombreCliente.Enabled = False
    txtExposicionCredito.Enabled = False
    txtGiroNeg.Enabled = False
    txtUltEndeuda.Enabled = False
    txtFecUltEndeuda.Enabled = False
    spnExpEmpAnio.Enabled = False
    spnExpEmpMes.Enabled = False
    
    txtCapacidadNeta.Enabled = False
    txtEndeudamiento.Enabled = False
    txtRentabilidadPat.Enabled = False
    txtLiquidezCte.Enabled = False
    txtIngresoNeto.Enabled = False
    txtExcedenteMensual.Enabled = False
End Sub
Private Sub CargarFlexEdit() 'Registrar New Formato Evaluacion
    Dim lnFila As Integer
    Dim CargarFlexEdit As Boolean
    Dim nMontoIni As Double
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim i As Integer
    Dim nFila, NumRegRS As Integer
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    nMontoIni = Format(0, "00.00")
    
   CargarFlexEdit = oNCOMFormatosEval.CargaDatosFlexEdit(fnFormato, sCtaCod, _
                                                        rsFeGastoNeg, _
                                                        rsFeDatGastoFam, _
                                                        rsFeDatOtrosIng, _
                                                        rsFeDatBalanGen, _
                                                        rsFeDatActivos, _
                                                        rsFeDatPasivos, _
                                                        rsFeDatPasivosNo, _
                                                        rsFeDatPatrimonio, _
                                                        rsFeDatRef, _
                                                        rsFeFlujoCaja, _
                                                        , , , , rsFeDatPDT)
                                                                                                      
    'Flex Activos ->CargaInicial
    feActivos.Clear
    feActivos.FormaCabecera
    feActivos.rows = 2
    Call LimpiaFlex(feActivos)
        nFila = 0
        NumRegRS = 0
        NumRegRS = rsFeDatActivos.RecordCount
        ReDim lvPrincipalActivos(NumRegRS)
        Do While Not rsFeDatActivos.EOF
            feActivos.AdicionaFila
            lnFila = feActivos.row
            feActivos.TextMatrix(lnFila, 1) = rsFeDatActivos!cConsDescripcion
            feActivos.TextMatrix(lnFila, 2) = Format(rsFeDatActivos!nPP, "#,#0.00")
            feActivos.TextMatrix(lnFila, 3) = Format(rsFeDatActivos!nPE, "#,#0.00")
            feActivos.TextMatrix(lnFila, 4) = Format(rsFeDatActivos!nTotal, "#,#0.00")
            feActivos.TextMatrix(lnFila, 5) = rsFeDatActivos!nConsCod
            feActivos.TextMatrix(lnFila, 6) = rsFeDatActivos!nConsValor
            
            'Lena datos de Registro en Matrix "lvPrincipalActivosPasivos"
            lvPrincipalActivos(lnFila).cConcepto = rsFeDatActivos!cConsDescripcion
            lvPrincipalActivos(lnFila).nImportePP = rsFeDatActivos!nPP
            lvPrincipalActivos(lnFila).nImportePE = rsFeDatActivos!nPP
            lvPrincipalActivos(lnFila).nConsCod = rsFeDatActivos!nConsCod
            lvPrincipalActivos(lnFila).nConsValor = rsFeDatActivos!nConsValor
            
                Select Case CInt(feActivos.TextMatrix(feActivos.row, 6)) 'celda que  o se puede editar
                 Case 1000, 100, 200, 300, 400, 500
                     Me.feActivos.BackColorRow (&H80000000)
                     Me.feActivos.ForeColorRow vbBlack, True
                 Case 102, 106, 301, 302, 303, 304, 401, 402, 403, 404, 405 'Colorear color amarillo
                     Me.feActivos.BackColorRow &HC0FFFF
                     Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
                 Case Else
                     Me.feActivos.BackColorRow (&HFFFFFF)
                     Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
                 End Select
        rsFeDatActivos.MoveNext
        Loop
    rsFeDatActivos.Close
    Set rsFeDatActivos = Nothing
                                                                                                                                                                                              
    'Flex Pasivos->CargaInicial
    fePasivos.Clear
    fePasivos.FormaCabecera
    fePasivos.rows = 2
    Call LimpiaFlex(fePasivos)
    nFila = 0
    NumRegRS = 0
    NumRegRS = rsFeDatPasivos.RecordCount
    ReDim lvPrincipalPasivos(NumRegRS)
    
        Do While Not rsFeDatPasivos.EOF
            fePasivos.AdicionaFila
            lnFila = fePasivos.row
            fePasivos.TextMatrix(lnFila, 1) = rsFeDatPasivos!cConsDescripcion
            fePasivos.TextMatrix(lnFila, 2) = Format(rsFeDatPasivos!nPP, "#,#0.00")
            fePasivos.TextMatrix(lnFila, 3) = Format(rsFeDatPasivos!nPE, "#,#0.00")
            fePasivos.TextMatrix(lnFila, 4) = Format(rsFeDatPasivos!nTotal, "#,#0.00")
            fePasivos.TextMatrix(lnFila, 5) = rsFeDatPasivos!nConsCod
            fePasivos.TextMatrix(lnFila, 6) = rsFeDatPasivos!nConsValor
            
            lvPrincipalPasivos(lnFila).cConcepto = rsFeDatPasivos!cConsDescripcion
            lvPrincipalPasivos(lnFila).nImportePP = rsFeDatPasivos!nPP
            lvPrincipalPasivos(lnFila).nImportePE = rsFeDatPasivos!nPP
            lvPrincipalPasivos(lnFila).nConsCod = rsFeDatPasivos!nConsCod
            lvPrincipalPasivos(lnFila).nConsValor = rsFeDatPasivos!nConsValor
            
            Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6)) 'celda que  o se puede editar
                Case 1000, 1001, 1002, 100, 200, 300, 500
                    Me.fePasivos.BackColorRow (&H80000000)
                    Me.fePasivos.ForeColorRow vbBlack, True
                Case 206
                    'Me.fePasivos.BackColorRow vbWhite, True        'CTI320200110 ERS003-2020. Comentó
                    'Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X" 'CTI320200110 ERS003-2020. Comentó
                    'CTI320200110 ERS003-2020. Agregó
                    If CDbl((fePasivos.TextMatrix(lnFila, 2)) > 0) Or CDbl((fePasivos.TextMatrix(lnFila, 3)) > 0) Then
                        Me.fePasivos.ForeColorRow vbBlack, True
                        Me.fePasivos.BackColorRow vbWhite, True
                        Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
                    Else
                        Me.fePasivos.BackColorRow vbWhite, True
                        Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
                        Me.fePasivos.RowHeight(lnFila) = 1
                    End If
                    'Fin CTI320200110 ERS003-2020
                Case 301
                    Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
                Case 104, 107, 109, 201, 203 'Color Amarillo
                    Me.fePasivos.BackColorRow &HC0FFFF
                    Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
                Case Else
                    Me.fePasivos.BackColorRow (&HFFFFFF)
                    Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
            End Select
            rsFeDatPasivos.MoveNext
        Loop
    rsFeDatPasivos.Close
    Set rsFeDatPasivos = Nothing
                                                                                                                                                
    'Flex Flujo Caja Mensual
    feFlujoCajaMensual.Clear
    feFlujoCajaMensual.FormaCabecera
    feFlujoCajaMensual.rows = 2
    Call LimpiaFlex(feFlujoCajaMensual)
        Do While Not rsFeFlujoCaja.EOF
            feFlujoCajaMensual.AdicionaFila
            lnFila = feFlujoCajaMensual.row
            feFlujoCajaMensual.TextMatrix(lnFila, 1) = rsFeFlujoCaja!nConsCod
            feFlujoCajaMensual.TextMatrix(lnFila, 2) = rsFeFlujoCaja!nConsValor
            feFlujoCajaMensual.TextMatrix(lnFila, 3) = rsFeFlujoCaja!cConsDescripcion
            feFlujoCajaMensual.TextMatrix(lnFila, 4) = Format(rsFeFlujoCaja!nMonto, "#,##0.00")
            
            Select Case CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2)) 'celda que  o se puede editar
                'Case 4, 5, 20, 1000 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                'Case 5, 6, 22, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Comentó
                Case 5, 6, 23, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Agregó
                    Me.feFlujoCajaMensual.BackColorRow (&H80000000)
                    Me.feFlujoCajaMensual.ForeColorRow vbBlack, True
                    Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
                'Case 17 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                Case 18 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                    Me.feFlujoCajaMensual.ForeColorRow (&H80000007)
                    Me.feFlujoCajaMensual.BackColorRow vbWhite, True
                    Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
                'Case 18 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                Case 19, gCodCuotaIfiNoSupervisadaFlujoCaja 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiNoSupervisadaFlujoCaja
                    Me.feFlujoCajaMensual.BackColorRow &HC0FFFF, True
                    Me.feFlujoCajaMensual.ForeColorRow (&H80000007), True
                    Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
                Case Else
                    Me.feFlujoCajaMensual.BackColorRow (&HFFFFFF)
                    Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
            End Select
            rsFeFlujoCaja.MoveNext
        Loop
    rsFeFlujoCaja.Close
    Set rsFeFlujoCaja = Nothing
                                                                              
    'Flex otros Ingresos
    feOtrosIngresos.Clear
    feOtrosIngresos.FormaCabecera
    feOtrosIngresos.rows = 2
    Call LimpiaFlex(feOtrosIngresos)
        Do While Not rsFeDatOtrosIng.EOF
            feOtrosIngresos.AdicionaFila
            lnFila = feOtrosIngresos.row
            feOtrosIngresos.TextMatrix(lnFila, 1) = rsFeDatOtrosIng!nConsValor
            feOtrosIngresos.TextMatrix(lnFila, 2) = rsFeDatOtrosIng!cConsDescripcion
            feOtrosIngresos.TextMatrix(lnFila, 3) = Format(rsFeDatOtrosIng!nMonto, "#,##0.00")
            rsFeDatOtrosIng.MoveNext
        Loop
    rsFeDatOtrosIng.Close
    Set rsFeDatOtrosIng = Nothing

    'Gastos Familiares
    feGastosFamiliares.Clear
    feGastosFamiliares.FormaCabecera
    feGastosFamiliares.rows = 2
    Call LimpiaFlex(feGastosFamiliares)
        Do While Not rsFeDatGastoFam.EOF
            feGastosFamiliares.AdicionaFila
            lnFila = feGastosFamiliares.row
            feGastosFamiliares.TextMatrix(lnFila, 1) = rsFeDatGastoFam!nConsValor
            feGastosFamiliares.TextMatrix(lnFila, 2) = rsFeDatGastoFam!cConsDescripcion
            feGastosFamiliares.TextMatrix(lnFila, 3) = Format(rsFeDatGastoFam!nMonto, "#,##0.00")

            Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
                Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiNoSupervisadaGastoFami
                    Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
                    Me.feGastosFamiliares.ForeColorRow (&H80000007), True
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                Case gCodDeudaLCNUGastoFami
                    Me.feGastosFamiliares.BackColorRow vbWhite, True
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                Case Else
                    Me.feGastosFamiliares.BackColorRow (&HFFFFFF)
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
            End Select
            rsFeDatGastoFam.MoveNext
        Loop
    rsFeDatGastoFam.Close
    Set rsFeDatGastoFam = Nothing
    
    'Declaracion PDT
    sMes1 = DevolverMes(1, nAnio3, nMes3)
    sMes2 = DevolverMes(2, nAnio2, nMes2)
    sMes3 = DevolverMes(3, nAnio1, nMes1)
    
    feDeclaracionPDT.Clear
    feDeclaracionPDT.FormaCabecera
    feDeclaracionPDT.rows = 2
        
    feDeclaracionPDT.TextMatrix(0, 4) = sMes3
    feDeclaracionPDT.TextMatrix(0, 5) = sMes2
    feDeclaracionPDT.TextMatrix(0, 6) = sMes1
    
    feDeclaracionPDT.TextMatrix(0, 1) = "Mes/Detalle" '& Space(8)
    For i = 1 To 2
        feDeclaracionPDT.AdicionaFila
        'feDeclaracionPDT.TextMatrix(i, 1) = Choose(i, "Compras" & Space(8), "Ventas" & Space(8))
        feDeclaracionPDT.TextMatrix(i, 1) = rsFeDatPDT!cConsDescripcion
        feDeclaracionPDT.TextMatrix(i, 2) = rsFeDatPDT!nConsCod
        feDeclaracionPDT.TextMatrix(i, 3) = rsFeDatPDT!nConsValor
        feDeclaracionPDT.TextMatrix(i, 4) = Choose(i, "0.00", "0.00") 'Mes3
        feDeclaracionPDT.TextMatrix(i, 5) = Choose(i, "0.00", "0.00") 'Mes2
        feDeclaracionPDT.TextMatrix(i, 6) = Choose(i, "0.00", "0.00") 'Mes1
        feDeclaracionPDT.TextMatrix(i, 7) = Choose(i, "0.00", "0.00") 'Promedio
        feDeclaracionPDT.TextMatrix(i, 8) = Choose(i, "0.00", "0.00") '%Ventas
        rsFeDatPDT.MoveNext
    Next i
End Sub
Private Function CargaDatos() As Boolean 'Mantenimiento Formatos
On Error GoTo ErrorCargaDatos
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim i As Integer
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
 
    CargaDatos = oNCOMFormatosEval.CargaDatosCredEvaluacion2(sCtaCod, _
                                                            fnFormato, _
                                                            rsCredEval, _
                                                            rsDatGastoNeg, _
                                                            rsDatGastoFam, _
                                                            rsDatOtrosIng, _
                                                            rsDatRef, _
                                                            rsDatActivos, _
                                                            rsDatPasivos, _
                                                            rsCuotaIFIs, _
                                                            rsPropuesta, _
                                                            rsCapacPagoNeta, _
                                                            rsDatRatioInd, _
                                                            rsDatActivoPasivo, _
                                                            rsDatIfiGastoNego, _
                                                            rsDatIfiGastoFami, , _
                                                            rsDatFlujoCaja, _
                                                            rsDatPDT, _
                                                            rsDatPDTDet, _
                                                            gFormatoActivos, _
                                                            gFormatoPasivos, rsDatParamFlujoCajaForm5, , , _
                                                            rsDatIfiNoSupervisadaGastoNego, rsDatIfiNoSupervisadaGastoFami)
                                                            
                                                            'rsDatIfiNoSupervisadaGastoNego, rsDatIfiNoSupervisadaGastoFami CTI320200110 ERS003-2020. Agregó
                                                            'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja rsDatParamFlujoCajaForm5
    Exit Function
ErrorCargaDatos:
    CargaDatos = False
    MsgBox Err.Description + ": Error al carga datos", vbInformation, "Error"
End Function

Private Sub CalculoTotal(ByVal pnTipo As Integer)
    Dim nTotalActPP As Double 'Total Activos (PP:Patrimonio Empresarial | PE: Patrionio Personal)
    Dim nTotalActPE As Double
    Dim nTotalPasPP As Double 'Total Pasivos
    Dim nTotalPasPE As Double 'Total Pasivos
    
    Dim nActivoCtePP As Double 'Activo Cuenta Corriente PP
    Dim nActivoCtePE As Double 'Activo Cuenta Corriente PE
    Dim nInventarioPP As Double 'Inventario PP
    Dim nInventarioPE As Double 'Inventario PE
    Dim nActiFijoPP As Double 'Activo Fijo PP
    Dim nActiFijoPE As Double 'Activo Fijo PE
    
    Dim nPasiCtePP As Double 'Pasivo CtaCte
    Dim nPasiCtePE As Double 'Pasivo CtaCte
    Dim nPasiNoCtePP As Double 'Pasivo NO CtaCte
    Dim nPasiNoCtePE As Double 'Pasivo NO CtaCte
    Dim nPatriPP As Double 'Patrimonio
    Dim nPatriPE As Double 'Patrimonio
    Dim nExistePat As Double 'Saber si el monto del detalle es > 0
    Dim nPatriTotal As Double
    
    Dim nMargenCaja As Double 'FlujoCaja
    Dim nOtrosEgresos As Double 'FlujoCaja
    Dim nMontoDeclarado As Double
    Dim nTotalPP As Currency, nTotalPE As Currency
    
    Dim nTotalActiCte As Currency
    Dim nTotalActivo As Currency
    Dim nTotalPasivo As Currency
    Dim nCuotasIfisFlujo As Double
    
    Dim nCapitalPP As Currency 'Para Calculo Fila -> Capital(Patrimonio)
    Dim nCapitalPE As Currency
    Dim nCapitalTotal As Currency
    Dim nCapitalAdicPP As Currency 'Fila -> Capital Adicional
    Dim nCapitalAdicPE As Currency
    Dim nCapitalAdicTotal As Currency
    Dim nExcedenteRevalPP As Currency 'Fila ->Excedente de Revaluación
    Dim nExcedenteRevalPE As Currency
    Dim nExcedenteRevalTotal As Currency
    Dim nReservaLegalPP As Currency 'Fila ->Reserva Legal
    Dim nReservaLegalPE As Currency
    Dim nReservaLegalTotal As Currency
    Dim nResultadoEjercicioPP As Currency 'Fila -> Resultado del Ejercicio
    Dim nResultadoEjercicioPE As Currency
    Dim nResultadoEjercicioTotal As Currency
    Dim nResultadoAcumuladoPP As Currency 'Fila ->Resultado Acumulado
    Dim nResultadoAcumuladoPE As Currency
    Dim nResultadoAcumuladoTotal As Currency
    Dim nTotalActivoPP As Currency
    Dim nTotalActivoPE As Currency
    
    Dim nTotalPasivoPatrimonio As Currency
    Dim nTotalPatrimonio As Currency
           
    Dim nPorcentajeVentas As Double
    Dim nPorcentajeCompras As Double
    
    nTotalActPP = 0: nTotalActPE = 0: nTotalPasPP = 0: nTotalPasPE = 0
    nActivoCtePP = 0: nActivoCtePE = 0: nInventarioPP = 0: nInventarioPE = 0: nActiFijoPP = 0: nActiFijoPE = 0
    nPasiCtePP = 0: nPasiCtePE = 0: nPasiNoCtePP = 0: nPasiNoCtePE = 0: nPatriPP = 0: nPatriPE = 0
    nExistePat = 0: nPatriTotal = 0
    nMargenCaja = 0: nOtrosEgresos = 0
    nMontoDeclarado = 0
    nTotalActiCte = 0: nTotalActivo = 0: nTotalPasivo = 0
    nPorcentajeVentas = 0: nPorcentajeCompras = 0

    nCapitalPP = 0: nCapitalPE = 0: nCapitalTotal = 0:
    nCapitalAdicPP = 0: nCapitalAdicPE = 0: nCapitalAdicTotal = 0
    nExcedenteRevalPP = 0: nExcedenteRevalPE = 0: nExcedenteRevalTotal = 0
    nReservaLegalPP = 0: nReservaLegalPE = 0: nReservaLegalTotal = 0
    nResultadoEjercicioPP = 0: nResultadoEjercicioPE = 0: nResultadoEjercicioTotal = 0:
    nResultadoAcumuladoPP = 0: nResultadoAcumuladoPE = 0: nResultadoAcumuladoTotal = 0:
    nTotalPasivoPatrimonio = 0: nTotalPatrimonio = 0
        
On Error GoTo ErrorCalculo
Select Case pnTipo
    Case 1:
            'ACTIVOS:**********->
            'Sumatoria: Activo Corriente
            For i = 2 To 5
                nActivoCtePP = nActivoCtePP + CDbl(Me.feActivos.TextMatrix(i, 2))
                nActivoCtePE = nActivoCtePE + CDbl(Me.feActivos.TextMatrix(i, 3))
            Next i
            Me.feActivos.TextMatrix(1, 2) = Format(nActivoCtePP, "#,#0.00") 'Resultado: Activo Cte PP
            Me.feActivos.TextMatrix(1, 3) = Format(nActivoCtePE, "#,#0.00") 'Resultado: Activo Cte PE
            nTotalPP = nActivoCtePP
            nTotalPE = nActivoCtePE
            'Sumatoria: Inventario
            For i = 7 To 10
                nInventarioPP = nInventarioPP + CDbl(Me.feActivos.TextMatrix(i, 2))
                nInventarioPE = nInventarioPE + CDbl(Me.feActivos.TextMatrix(i, 3))
            Next i
            Me.feActivos.TextMatrix(6, 2) = Format(nInventarioPP, "#,#0.00")
            Me.feActivos.TextMatrix(6, 3) = Format(nInventarioPE, "#,#0.00")
            
            'Sumatoria: Activo Fijo
            For i = 12 To 16
                nActiFijoPP = nActiFijoPP + CDbl(Me.feActivos.TextMatrix(i, 2))
                nActiFijoPE = nActiFijoPE + CDbl(Me.feActivos.TextMatrix(i, 3))
            Next i
            Me.feActivos.TextMatrix(11, 2) = Format(nActiFijoPP, "#,#0.00")
            Me.feActivos.TextMatrix(11, 3) = Format(nActiFijoPE, "#,#0.00")
            nTotalPP = nTotalPP + nActiFijoPP
            nTotalPE = nTotalPE + nActiFijoPE
            
            Me.feActivos.TextMatrix(17, 2) = Format(nTotalPP + nInventarioPP, "#,#0.00")
            Me.feActivos.TextMatrix(17, 3) = Format(nTotalPE + nInventarioPE, "#,#0.00")
                
            'Columna Total
            For i = 1 To Me.feActivos.rows - 1
                nTotalActPP = CDbl(Me.feActivos.TextMatrix(i, 2))
                nTotalActPE = CDbl(Me.feActivos.TextMatrix(i, 3))
                Me.feActivos.TextMatrix(i, 4) = Format(nTotalActPP + nTotalActPE, "#,#0.00")
            Next i
            
            'Calculo del "TotalActivoCte":
            nTotalActiCte = Format(Me.feActivos.TextMatrix(6, 4) + nActivoCtePP + nActivoCtePE, "#,#0.00")
            Me.feActivos.TextMatrix(1, 4) = Format(CCur(nTotalActiCte), "#,#0.00")
            'Calculo del "TOTAL":
            nTotalActivo = Format(CCur(Me.feActivos.TextMatrix(1, 4)) + CCur(Me.feActivos.TextMatrix(11, 4)), "#,#0.00")
            Me.feActivos.TextMatrix(17, 4) = Format(nTotalActivo, "#,#0.00")

            Call CalculoTotal(2)
    Case 2:
            'PASIVOS:**********->
            'Sumatoria (PP/ PE): Pasivo Corriente
            For i = 2 To 7
                nPasiCtePP = nPasiCtePP + CDbl(Me.fePasivos.TextMatrix(i, 2))
                nPasiCtePE = nPasiCtePE + CDbl(Me.fePasivos.TextMatrix(i, 3))
            Next i
            Me.fePasivos.TextMatrix(1, 2) = Format(nPasiCtePP, "#,#0.00") 'Resultado: PAsivo Cte PP
            Me.fePasivos.TextMatrix(1, 3) = Format(nPasiCtePE, "#,#0.00") 'Resultado: Pasivo Cte PE
            nTotalPP = nPasiCtePP
            nTotalPE = nPasiCtePE
            'Sumatoria (PP/ PE): Pasivo No Corriente
            For i = 9 To 12
                nPasiNoCtePP = nPasiNoCtePP + CDbl(Me.fePasivos.TextMatrix(i, 2))
                nPasiNoCtePE = nPasiNoCtePE + CDbl(Me.fePasivos.TextMatrix(i, 3))
            Next i
            Me.fePasivos.TextMatrix(8, 2) = Format(nPasiNoCtePP, "#,#0.00")
            Me.fePasivos.TextMatrix(8, 3) = Format(nPasiNoCtePE, "#,#0.00")
            nTotalPP = nTotalPP + nPasiNoCtePP
            nTotalPE = nTotalPE + nPasiNoCtePE
            
           'Sumatoria Capital (Patrimonio)
            nTotalActivoPP = Me.feActivos.TextMatrix(17, 2)
            nTotalActivoPE = Me.feActivos.TextMatrix(17, 3)
            
            nCapitalAdicPP = Me.fePasivos.TextMatrix(15, 2)
            nCapitalAdicPE = Me.fePasivos.TextMatrix(15, 3)
            nCapitalAdicTotal = Me.fePasivos.TextMatrix(15, 4)
            nExcedenteRevalPP = Me.fePasivos.TextMatrix(16, 2)
            nExcedenteRevalPE = Me.fePasivos.TextMatrix(16, 3)
            nExcedenteRevalTotal = Me.fePasivos.TextMatrix(16, 4)
            nReservaLegalPP = Me.fePasivos.TextMatrix(17, 2)
            nReservaLegalPE = Me.fePasivos.TextMatrix(17, 3)
            nReservaLegalTotal = Me.fePasivos.TextMatrix(17, 4)
            nResultadoEjercicioPP = Me.fePasivos.TextMatrix(18, 2)
            nResultadoEjercicioPE = Me.fePasivos.TextMatrix(18, 3)
            nResultadoEjercicioTotal = Me.fePasivos.TextMatrix(18, 4)
            nResultadoAcumuladoPP = Me.fePasivos.TextMatrix(19, 2)
            nResultadoAcumuladoPE = Me.fePasivos.TextMatrix(19, 3)
            nResultadoAcumuladoTotal = Me.fePasivos.TextMatrix(19, 4)
            nTotalPasivo = Me.fePasivos.TextMatrix(20, 4)
                       
            nCapitalPP = nTotalActivoPP - nTotalPP - (nCapitalAdicPP + nExcedenteRevalPP + nReservaLegalPP + nResultadoEjercicioPP + nResultadoAcumuladoPP) 'Capital - Patrimonio PP
            nCapitalPE = nTotalActivoPE - nTotalPE - (nCapitalAdicPE + nExcedenteRevalPE + nReservaLegalPE + nResultadoEjercicioPE + nResultadoAcumuladoPE) 'Capital - Patrimonio PE
            nCapitalTotal = nTotalActivo - nTotalPasivo - (nCapitalAdicTotal + nExcedenteRevalTotal + nReservaLegalTotal + nResultadoEjercicioTotal + nResultadoAcumuladoTotal) 'Total Capital -Patrimonio
           
            Me.fePasivos.TextMatrix(14, 2) = Format(nCapitalPP, "#,#0.00")
            Me.fePasivos.TextMatrix(14, 3) = Format(nCapitalPE, "#,#0.00")
            Me.fePasivos.TextMatrix(14, 4) = Format(nCapitalTotal, "#,#0.00")
           
           'Verificar si Existe detalle Patrimonio
           For i = 14 To 19
            nExistePat = nExistePat + CDbl(Me.fePasivos.TextMatrix(i, 4))
           Next i
                
            'Sumatoria (PP/ PE): Patrimonio
            If nExistePat <> 0 Then
                For i = 14 To 19
                nPatriPP = nPatriPP + CDbl(Me.fePasivos.TextMatrix(i, 2))
                nPatriPE = nPatriPE + CDbl(Me.fePasivos.TextMatrix(i, 3))
                Next i
                Me.fePasivos.TextMatrix(13, 2) = Format(nPatriPP, "#,#0.00")
                Me.fePasivos.TextMatrix(13, 3) = Format(nPatriPE, "#,#0.00")
            Else
                nPatriTotal = Me.feActivos.TextMatrix(17, 4) - Me.fePasivos.TextMatrix(20, 4)
                Me.fePasivos.TextMatrix(13, 4) = Format(nPatriTotal, "#,#0.00")
            End If
            
            'Total Pasivo y Total Pasivo y Patrimonio (PP | PE)
            Me.fePasivos.TextMatrix(20, 2) = Format(nPasiCtePP + nPasiNoCtePP, "#,#0.00") 'Total Pasivo PP
            Me.fePasivos.TextMatrix(20, 3) = Format(nPasiCtePE + nPasiNoCtePE, "#,#0.00") 'Total Pasivo PE
   
            'Me.fePasivos.TextMatrix(21, 2) = Format(nPasiCtePP + nPasiNoCtePP + nPatriPP, "#,#0.00") 'Total Pasivo y Patrimonio
            'Me.fePasivos.TextMatrix(21, 3) = Format(nPasiCtePE + nPasiNoCtePE + nPatriPE, "#,#0.00") 'Total Pasivo y Patrimonio
             
            'Columna Total= PP + PE
            For i = 1 To Me.fePasivos.rows - 1
                nTotalPasPP = CDbl(Me.fePasivos.TextMatrix(i, 2))
                nTotalPasPE = CDbl(Me.fePasivos.TextMatrix(i, 3))
                Me.fePasivos.TextMatrix(i, 4) = Format(nTotalPasPP + nTotalPasPE, "#,#0.00")
            Next i
            
            If nPatriTotal <> 0 Then ' Para el Caso que no Exista detalle de Patrimonio
               Me.fePasivos.TextMatrix(13, 4) = Format(nPatriTotal, "#,#0.00")
               Me.fePasivos.TextMatrix(21, 4) = Format(CDbl(Me.fePasivos.TextMatrix(1, 4)) + CDbl(Me.fePasivos.TextMatrix(8, 4)) + CDbl(Me.fePasivos.TextMatrix(13, 4)), "#,#0.00")
            End If
            
            'Calculo de Total Pasivo y Patrimonio
             nTotalPatrimonio = Me.fePasivos.TextMatrix(13, 4)
             nTotalPasivoPatrimonio = nTotalPasivo + nTotalPatrimonio
             Me.fePasivos.TextMatrix(21, 4) = Format(nTotalPasivoPatrimonio, "#,#0.00")
            
             Me.fePasivos.TextMatrix(21, 2) = Format(nPasiCtePP + nPasiNoCtePP + nPatriPP, "#,#0.00") 'Total Pasivo y Patrimonio
             Me.fePasivos.TextMatrix(21, 3) = Format(nPasiCtePE + nPasiNoCtePE + nPatriPE, "#,#0.00") 'Total Pasivo y Patrimonio
             
            
        Case 3:
            'Margen Bruto Caja
            'For i = 1 To 2 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            For i = 1 To 3 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                nMargenCaja = nMargenCaja + CDbl(Me.feFlujoCajaMensual.TextMatrix(i, 4))
            Next i
            nMargenCaja = nMargenCaja - CDbl(Me.feFlujoCajaMensual.TextMatrix(i, 4))
            'Me.feFlujoCajaMensual.TextMatrix(4, 4) = Format(nMargenCaja, "#,#0.00") 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            Me.feFlujoCajaMensual.TextMatrix(5, 4) = Format(nMargenCaja, "#,#0.00") 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            
           'Otros Egresos
            'For i = 6 To 19 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            For i = 7 To 21 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                nOtrosEgresos = nOtrosEgresos + CDbl(Me.feFlujoCajaMensual.TextMatrix(i, 4))
            Next i
            'Me.feFlujoCajaMensual.TextMatrix(5, 4) = Format(nOtrosEgresos, "#,#0.00") 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            Me.feFlujoCajaMensual.TextMatrix(6, 4) = Format(nOtrosEgresos, "#,#0.00") 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja

           'Saldo Disponible
           'nCuotasIfisFlujo = CDbl(feFlujoCajaMensual.TextMatrix(17, 4)) + CDbl(feFlujoCajaMensual.TextMatrix(18, 4))'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
           nCuotasIfisFlujo = CDbl(feFlujoCajaMensual.TextMatrix(18, 4)) + CDbl(feFlujoCajaMensual.TextMatrix(19, 4)) 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
           'Me.feFlujoCajaMensual.TextMatrix(20, 4) = Format(nMargenCaja - (CDbl(nOtrosEgresos) - CDbl(nCuotasIfisFlujo)), "#,#0.00")'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
           'Me.feFlujoCajaMensual.TextMatrix(22, 4) = Format(nMargenCaja - (CDbl(nOtrosEgresos) - CDbl(nCuotasIfisFlujo)), "#,#0.00") 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Comentó
           Me.feFlujoCajaMensual.TextMatrix(23, 4) = Format(nMargenCaja - (CDbl(nOtrosEgresos) - CDbl(nCuotasIfisFlujo)), "#,#0.00") 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Agregó
           
        Case 4:
        'Promedio Declaracion PDT
            'If CCur(feFlujoCajaMensual.TextMatrix(1, 4)) = 0 Or CCur(feFlujoCajaMensual.TextMatrix(3, 4)) = 0 Then 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            If CCur(feFlujoCajaMensual.TextMatrix(1, 4)) = 0 Or CCur(feFlujoCajaMensual.TextMatrix(4, 4)) = 0 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                MsgBox "Consideraciones en el flujo de caja mensual " & Chr(10) & " - El monto de ventas al contado, no tiene que ser cero." & Chr(10) & " - El monto de egresos por compras, no tiene que ser cero.", vbInformation, "Alerta"
                Exit Sub
            End If
            For i = 1 To feDeclaracionPDT.rows - 1
                nMontoDeclarado = CDbl(Me.feDeclaracionPDT.TextMatrix(i, 4)) + CDbl(Me.feDeclaracionPDT.TextMatrix(i, 5)) + CDbl(Me.feDeclaracionPDT.TextMatrix(i, 6))
                nMontoDeclarado = nMontoDeclarado / 3
                Me.feDeclaracionPDT.TextMatrix(i, 7) = Format(nMontoDeclarado, "#,#0.00")
            Next
                
        'Para el %Declarado
        nPorcentajeVentas = Round(CDbl(feDeclaracionPDT.TextMatrix(1, 7)) / CDbl(feFlujoCajaMensual.TextMatrix(1, 4)), 4)
        'nPorcentajeCompras = Round(CDbl(feDeclaracionPDT.TextMatrix(2, 7)) / CDbl(feFlujoCajaMensual.TextMatrix(3, 4)), 4)'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
        nPorcentajeCompras = Round(CDbl(feDeclaracionPDT.TextMatrix(2, 7)) / CDbl(feFlujoCajaMensual.TextMatrix(4, 4)), 4) 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    
        Me.feDeclaracionPDT.TextMatrix(1, 8) = CStr(nPorcentajeVentas * 100) & "%"
        Me.feDeclaracionPDT.TextMatrix(2, 8) = CStr(nPorcentajeCompras * 100) & "%"
End Select

Exit Sub
ErrorCalculo:
MsgBox "Aviso: Ingrese los datos Correctamente." & Chr(13) & "Detalles: " & Err.Description, vbInformation, "Aviso"
'Select Case pnTipo
 '  Case 1:
'            txtIngresoNegocio.Text = "0.00"
'            txtEgresoNegocio.Text = "0.00"
'End Select
 'Call CalculoTotal(pnTipo)
End Sub

Private Function Registro()
    gsOpeCod = gCredRegistrarEvaluacionCred
    txtNombreCliente.Text = fsCliente
    txtGiroNeg.Text = fsGiroNego
    
    txtCapacidadNeta.Enabled = False
    txtEndeudamiento.Enabled = False
    txtRentabilidadPat.Enabled = False
    txtLiquidezCte.Enabled = False
    txtIngresoNeto.Enabled = False
    txtExcedenteMensual.Enabled = False
    
    'si el cliente es nuevo-> referido obligatorio
    'If fnColocCondi = 1 Then 'LUCV2017115, Según correo: RUSI
    If Not fbTieneReferido6Meses Then
        frameReferido.Enabled = True
        feReferidos.Enabled = True
        cmdAgregarRef.Enabled = True
        cmdQuitar5.Enabled = True
        txtComentario5.Enabled = True
        frameComentario.Enabled = True
    Else
        frameReferido.Enabled = False
        feReferidos.Enabled = False
        cmdAgregarRef.Enabled = False
        cmdQuitar5.Enabled = False
        txtComentario5.Enabled = False
        frameComentario.Enabled = False
    End If

    'Ratios: Aceptable / Critico ->*****
    If Not (rsAceptableCritico.BOF Or rsAceptableCritico.EOF) Then
        If rsAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
            Me.lblCapaAceptable.Caption = "Aceptable"
            Me.lblCapaAceptable.ForeColor = &H8000&
        Else
            Me.lblCapaAceptable.Caption = "Crítico"
            Me.lblCapaAceptable.ForeColor = vbRed
        End If
        
        If rsAceptableCritico!nEndeud = 1 Then 'Endeudamiento Pat.
            Me.lblEndeAceptable.Caption = "Aceptable"
            Me.lblEndeAceptable.ForeColor = &H8000&
        Else
            Me.lblEndeAceptable.Caption = "Crítico"
            Me.lblEndeAceptable.ForeColor = vbRed
        End If
    Else
        lblCapaAceptable.Visible = False
        lblCapaAceptable.Visible = False
    End If
    'Fin Ratios <-****
    
     '*****->No Refinanciados (Propuesta Credito)
    If fnColocCondi <> 4 Then
        txtFechaVisita.Enabled = True
        txtEntornoFamiliar5.Enabled = True
        txtGiroUbicacion5.Enabled = True
        txtExperiencia5.Enabled = True
        txtFormalidadNegocio5.Enabled = True
        txtColaterales5.Enabled = True
        txtDestino5.Enabled = True
     Else
        framePropuesta.Enabled = False
        txtFechaVisita.Enabled = False
        txtEntornoFamiliar5.Enabled = False
        txtGiroUbicacion5.Enabled = False
        txtExperiencia5.Enabled = False
        txtFormalidadNegocio5.Enabled = False
        txtColaterales5.Enabled = False
        txtDestino5.Enabled = False
    End If
        '*****->Fin No Refinanciados
    
        'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        If Not (rsDatParamFlujoCajaForm5.BOF And rsDatParamFlujoCajaForm5.EOF) Then
            EditMoneyForm5IVC.Text = Format(rsDatParamFlujoCajaForm5!nIncVentCont, "#0.00")
            EditMoneyForm5ICM.Text = Format(rsDatParamFlujoCajaForm5!nIncCompMerc, "#0.00")
            EditMoneyForm5IPP.Text = Format(rsDatParamFlujoCajaForm5!nIncPagPers, "#0.00")
            EditMoneyForm5IGV.Text = Format(rsDatParamFlujoCajaForm5!nIncGastvent, "#0.00")
            EditMoneyForm5IC.Text = Format(rsDatParamFlujoCajaForm5!nIncConsu, "#0.00")
        End If
        rsDatParamFlujoCajaForm5.Close
        Set rsDatParamFlujoCajaForm5 = Nothing
        'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
 'CTI3 ERS0032020
    ReDim MatIfiGastoNego(rsDatIfiGastoNego.RecordCount, 4)
    i = 0
    Do While Not rsDatIfiGastoNego.EOF
        MatIfiGastoNego(i, 0) = rsDatIfiGastoNego!nNroCuota
        MatIfiGastoNego(i, 1) = rsDatIfiGastoNego!CDescripcion
        MatIfiGastoNego(i, 2) = Format(IIf(IsNull(rsDatIfiGastoNego!nMonto), 0, rsDatIfiGastoNego!nMonto), "#0.00")
        rsDatIfiGastoNego.MoveNext
          i = i + 1
    Loop
    rsDatIfiGastoNego.Close
    Set rsDatIfiGastoNego = Nothing
    
     'Carga de rsDatIfiGastoFami -> Matrix
    ReDim MatIfiGastoFami(rsDatIfiGastoFami.RecordCount, 4)
    j = 0
    Do While Not rsDatIfiGastoFami.EOF
        MatIfiGastoFami(j, 0) = rsDatIfiGastoFami!nNroCuota
        MatIfiGastoFami(j, 1) = rsDatIfiGastoFami!CDescripcion
        MatIfiGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiGastoFami!nMonto), 0, rsDatIfiGastoFami!nMonto), "#0.00")
        rsDatIfiGastoFami.MoveNext
    j = j + 1
    Loop
    rsDatIfiGastoFami.Close
    Set rsDatIfiGastoFami = Nothing
   
    If fnProducto = "800" Then
        SSTabIngresos5.TabVisible(0) = False
    End If
    'Fin CTI320200110
    
End Function

Private Function Mantenimiento()
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Dim lnFila As Integer
    pnMontoOtrasIfisConsumo = 0
    pnMontoOtrasIfisEmpresarial = 0
    If fnTipoPermiso = 3 Then
        gsOpeCod = gCredMantenimientoEvaluacionCred
    Else
        'gsOpeCod = gCredVerificacionEvaluacionCred
    End If
    
      'Para Botones *****
    If Not fbBloqueaTodo Then
        cmdInformeVisita.Enabled = False
        cmdVerCar.Enabled = False
        cmdImprimir.Enabled = False
    End If
    
    'Ver Ratios
    If fnEstado > 2000 Then
        SSTabRatios.Visible = True
    Else
        SSTabRatios.Visible = False
        cmdInformeVisita.Enabled = False
        cmdVerCar.Enabled = False
        cmdImprimir.Enabled = False
    End If
    
    'Ratios/ Indicadores
    txtCapacidadNeta.Enabled = False
    txtEndeudamiento.Enabled = False
    txtRentabilidadPat.Enabled = False
    txtLiquidezCte.Enabled = False
    txtIngresoNeto.Enabled = False
    txtExcedenteMensual.Enabled = False
    
    'Si el cliente es nuevo-> referido obligatorio
    'If fnColocCondi = 1 Then 'LUCV2017115, Según correo: RUSI
    If Not fbTieneReferido6Meses Then
        frameReferido.Enabled = True
        feReferidos.Enabled = True
        cmdAgregarRef.Enabled = True
        cmdQuitar5.Enabled = True
        txtComentario5.Enabled = True
        frameComentario.Enabled = True
    Else
        frameReferido.Enabled = False
        feReferidos.Enabled = False
        cmdAgregarRef.Enabled = False
        cmdQuitar5.Enabled = False
        txtComentario5.Enabled = False
        frameComentario.Enabled = False
    End If
    
     'Ratios: Aceptable / Critico ->*****
     If Not (rsAceptableCritico.EOF Or rsAceptableCritico.BOF) Then
        If rsAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
            Me.lblCapaAceptable.Caption = "Aceptable"
            Me.lblCapaAceptable.ForeColor = &H8000&
        Else
            Me.lblCapaAceptable.Caption = "Crítico"
            Me.lblCapaAceptable.ForeColor = vbRed
        End If
        
        If rsAceptableCritico!nEndeud = 1 Then 'Endeudamiento Pat.
            Me.lblEndeAceptable.Caption = "Aceptable"
            Me.lblEndeAceptable.ForeColor = &H8000&
        Else
            Me.lblEndeAceptable.Caption = "Crítico"
            Me.lblEndeAceptable.ForeColor = vbRed
        End If
    Else
        Me.lblCapaAceptable.Visible = False
        Me.lblEndeAceptable.Visible = False
    End If
    'Fin Ratios <-****
    
    '*****->No Refinanciados (Propuesta Credito)
    If fnColocCondi <> 4 Then
        txtFechaVisita.Enabled = True
        txtEntornoFamiliar5.Enabled = True
        txtGiroUbicacion5.Enabled = True
        txtExperiencia5.Enabled = True
        txtFormalidadNegocio5.Enabled = True
        txtColaterales5.Enabled = True
        txtDestino5.Enabled = True
     Else
        framePropuesta.Enabled = False
        txtFechaVisita.Enabled = False
        txtEntornoFamiliar5.Enabled = False
        txtGiroUbicacion5.Enabled = False
        txtExperiencia5.Enabled = False
        txtFormalidadNegocio5.Enabled = False
        txtColaterales5.Enabled = False
        txtDestino5.Enabled = False
    End If
    '*****->Fin No Refinanciados
    
    'LUCV20160626, Para CARGAR CABECERA->**********
    Set rsDCredito = oDCOMFormatosEval.RecuperaSolicitudDatoBasicosEval(sCtaCod) ' Datos Basicos del Credito Solicitado
    ActXCodCta.NroCuenta = sCtaCod
    txtGiroNeg.Text = rsCredEval!cActividad
    txtNombreCliente.Text = fsCliente
    spnExpEmpAnio.valor = rsCredEval!nExpEmpAnio
    spnExpEmpMes.valor = rsCredEval!nExpEmpMes
    spnTiempoLocalAnio.valor = rsCredEval!nTmpoLocalAnio
    spnTiempoLocalMes.valor = rsCredEval!nTmpoLocalMes
    OptCondLocal(rsCredEval!nCondiLocal).value = 1
    txtCondLocalOtros.Text = rsCredEval!cCondiLocalOtro
    txtExposicionCredito.Text = Format(rsCredEval!nExposiCred, "#,##0.00")
    txtFechaEvaluacion.Text = Format(rsCredEval!dFecEval, "dd/mm/yyyy")
    txtUltEndeuda.Text = Format(rsCredEval!nUltEndeSBS, "#,##0.00")
    txtFecUltEndeuda.Text = Format(rsCredEval!dUltEndeuSBS, "dd/mm/yyyy")
     
    'LUCV20160626, Para CARGAR PROPUESTA->**********
     If fnColocCondi <> 4 Then
        txtFechaVisita.Text = Format(rsPropuesta!dFecVisita, "dd/mm/yyyy")
        txtEntornoFamiliar5.Text = Trim(rsPropuesta!cEntornoFami)
        txtGiroUbicacion5.Text = Trim(rsPropuesta!cGiroUbica)
        txtExperiencia5.Text = Trim(rsPropuesta!cExpeCrediticia)
        txtFormalidadNegocio5.Text = Trim(rsPropuesta!cFormalNegocio)
        txtColaterales5.Text = Trim(rsPropuesta!cColateGarantia)
        txtDestino5.Text = Trim(rsPropuesta!cDestino)
        txtComentario5.Text = Trim(rsCredEval!cComentario)
    End If
    'LUCV20160626, Para la CARGAR FLEX - Mantenimiento **********->
    
    If Not (rsDatIfiGastoFami.BOF Or rsDatIfiGastoFami.EOF) Then
        For i = 1 To rsDatIfiGastoFami.RecordCount
           pnMontoOtrasIfisConsumo = pnMontoOtrasIfisConsumo + rsDatIfiGastoFami!nMonto
           rsDatIfiGastoFami.MoveNext
        Next i
        rsDatIfiGastoFami.MoveFirst
    End If
    If Not (rsDatIfiGastoNego.BOF Or rsDatIfiGastoNego.EOF) Then
        For i = 1 To rsDatIfiGastoNego.RecordCount
           pnMontoOtrasIfisEmpresarial = pnMontoOtrasIfisEmpresarial + rsDatIfiGastoNego!nMonto
           rsDatIfiGastoNego.MoveNext
        Next i
        rsDatIfiGastoNego.MoveFirst
    End If
    'Call FormatearGrillas(feGastosFamiliares2)
    Call LimpiaFlex(feGastosFamiliares)
        Do While Not rsDatGastoFam.EOF
            feGastosFamiliares.AdicionaFila
            lnFila = feGastosFamiliares.row
            feGastosFamiliares.TextMatrix(lnFila, 1) = rsDatGastoFam!nConsValor
            feGastosFamiliares.TextMatrix(lnFila, 2) = rsDatGastoFam!cConsDescripcion
            feGastosFamiliares.TextMatrix(lnFila, 3) = Format(rsDatGastoFam!nMonto, "#,##0.00")
            
            If fbImprimirVB And rsDatGastoFam!nConsValor = 5 Then
                feGastosFamiliares.TextMatrix(lnFila, 3) = Format(pnMontoOtrasIfisConsumo, "#,##0.00")
            End If

            Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
                Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiNoSupervisadaGastoFami
                    Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
                    Me.feGastosFamiliares.ForeColorRow (&H80000007), True
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                Case gCodDeudaLCNUGastoFami
                    Me.feGastosFamiliares.BackColorRow vbWhite, True
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                Case Else
                    Me.feGastosFamiliares.BackColorRow (&HFFFFFF)
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
            End Select
            rsDatGastoFam.MoveNext
        Loop

    rsDatGastoFam.Close
    Set rsDatGastoFam = Nothing
    
    'Call FormatearGrillas(feOtrosIngresos2)
    Call LimpiaFlex(feOtrosIngresos)
        Do While Not rsDatOtrosIng.EOF
            feOtrosIngresos.AdicionaFila
            lnFila = feOtrosIngresos.row
            feOtrosIngresos.TextMatrix(lnFila, 1) = rsDatOtrosIng!nConsValor
            feOtrosIngresos.TextMatrix(lnFila, 2) = rsDatOtrosIng!cConsDescripcion
            feOtrosIngresos.TextMatrix(lnFila, 3) = Format(rsDatOtrosIng!nMonto, "#,##0.00")
            rsDatOtrosIng.MoveNext
        Loop
    rsDatOtrosIng.Close
    Set rsDatOtrosIng = Nothing
    
    'Call FormatearGrillas(feCuotaIfis)
    Call LimpiaFlex(frmCredFormEvalCuotasIfis.feCuotaIfis)
        Do While Not rsCuotaIFIs.EOF
            frmCredFormEvalCuotasIfis.feCuotaIfis.AdicionaFila
            lnFila = frmCredFormEvalCuotasIfis.feCuotaIfis.row
            frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 1) = rsCuotaIFIs!CDescripcion
            frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 2) = Format(rsCuotaIFIs!nMonto, "#,##0.00")
            rsCuotaIFIs.MoveNext
        Loop
    rsCuotaIFIs.Close
    Set rsCuotaIFIs = Nothing
    
    'Call FormatearGrillas(feReferidos2)
    Call LimpiaFlex(feReferidos)
        Do While Not rsDatRef.EOF
            feReferidos.AdicionaFila
            lnFila = feReferidos.row
            feReferidos.TextMatrix(lnFila, 0) = rsDatRef!nCodRef
            feReferidos.TextMatrix(lnFila, 1) = rsDatRef!cNombre
            feReferidos.TextMatrix(lnFila, 2) = rsDatRef!cDniNom
            feReferidos.TextMatrix(lnFila, 3) = rsDatRef!cTelf
            feReferidos.TextMatrix(lnFila, 4) = rsDatRef!cReferido
            feReferidos.TextMatrix(lnFila, 5) = rsDatRef!cDNIRef
            rsDatRef.MoveNext
        Loop
    rsDatRef.Close
    Set rsDatRef = Nothing
    
    'Call FormatearGrillas(feDeclaracionPDT)
    'Call LimpiaFlex(feDeclaracionPDT)
    lnFila = 1
    Do While Not rsDatPDTDet.EOF
        'feDeclaracionPDT.AdicionaFila
        feDeclaracionPDT.TextMatrix(lnFila, 2) = Format(rsDatPDTDet!nConsCod, "#,##0.00")
        feDeclaracionPDT.TextMatrix(lnFila, 3) = Format(rsDatPDTDet!nConsValor, "#,##0.00")
        feDeclaracionPDT.TextMatrix(lnFila, 4) = Format(rsDatPDTDet!nMontoMes1, "#,##0.00")
        feDeclaracionPDT.TextMatrix(lnFila, 5) = Format(rsDatPDTDet!nMontoMes2, "#,##0.00")
        feDeclaracionPDT.TextMatrix(lnFila, 6) = Format(rsDatPDTDet!nMontoMes3, "#,##0.00")
        feDeclaracionPDT.TextMatrix(lnFila, 7) = Format(rsDatPDTDet!nPromedio, "#,##0.00")
        feDeclaracionPDT.TextMatrix(lnFila, 8) = Format(rsDatPDTDet!nPorcentajeVent, "#,##0.00")
        rsDatPDTDet.MoveNext
        lnFila = lnFila + 1
    Loop
    rsDatPDTDet.Close
    Set rsDatPDTDet = Nothing

    feDeclaracionPDT.TextMatrix(0, 4) = DevolverMesDatos(CInt(rsDatPDT!nMes1))
    feDeclaracionPDT.TextMatrix(0, 5) = DevolverMesDatos(CInt(rsDatPDT!nMes2))
    feDeclaracionPDT.TextMatrix(0, 6) = DevolverMesDatos(CInt(rsDatPDT!nMes3))
    
    'Call FormatearGrillas(feFlujoCajaMensual)
    Call LimpiaFlex(feFlujoCajaMensual)
        Do While Not rsDatFlujoCaja.EOF
            feFlujoCajaMensual.AdicionaFila
            lnFila = feFlujoCajaMensual.row
            feFlujoCajaMensual.TextMatrix(lnFila, 1) = rsDatFlujoCaja!nConsCod
            feFlujoCajaMensual.TextMatrix(lnFila, 2) = rsDatFlujoCaja!nConsValor
            feFlujoCajaMensual.TextMatrix(lnFila, 3) = rsDatFlujoCaja!cConcepto
            feFlujoCajaMensual.TextMatrix(lnFila, 4) = Format(rsDatFlujoCaja!nMonto, "#,##0.00")
            
            If fbImprimirVB And rsDatFlujoCaja!nConsValor = 19 Then
                feFlujoCajaMensual.TextMatrix(lnFila, 4) = Format(pnMontoOtrasIfisEmpresarial, "#,##0.00")
            End If
            
            Select Case CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2)) 'celda que  o se puede editar
                'Case 4, 5, 20, 1000'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                'Case 5, 6, 22, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Comentó
                Case 5, 6, 23, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Agregó
                    Me.feFlujoCajaMensual.BackColorRow (&H80000000)
                    Me.feFlujoCajaMensual.ForeColorRow vbBlack, True
                'Case 17'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                Case 18 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                    Me.feFlujoCajaMensual.ForeColorRow (&H80000007)
                    Me.feFlujoCajaMensual.BackColorRow vbWhite, True
                    Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
                'Case 18'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                Case 19, gCodCuotaIfiNoSupervisadaFlujoCaja  'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiNoSupervisadaFlujoCaja
                    Me.feFlujoCajaMensual.BackColorRow &HC0FFFF, True
                    Me.feFlujoCajaMensual.ForeColorRow (&H80000007), True
                    Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
                Case Else
                    Me.feFlujoCajaMensual.BackColorRow (&HFFFFFF)
                    Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
            End Select
            rsDatFlujoCaja.MoveNext
        Loop
    rsDatFlujoCaja.Close
    Set rsDatFlujoCaja = Nothing


    'Call FormatearGrillas(feActivo)
    Call LimpiaFlex(feActivos)
        Do While Not rsDatActivos.EOF
            feActivos.AdicionaFila
            lnFila = feActivos.row
            feActivos.TextMatrix(lnFila, 1) = rsDatActivos!cConsDescripcion
            feActivos.TextMatrix(lnFila, 2) = Format(rsDatActivos!PP, "#,##0.00")
            feActivos.TextMatrix(lnFila, 3) = Format(rsDatActivos!PE, "#,##0.00")
            feActivos.TextMatrix(lnFila, 4) = Format(rsDatActivos!nTotal, "#,##0.00")
            feActivos.TextMatrix(lnFila, 5) = rsDatActivos!nConsCod
            feActivos.TextMatrix(lnFila, 6) = rsDatActivos!nConsValor
            
            Select Case CInt(feActivos.TextMatrix(feActivos.row, 6)) 'celda que  o se puede editar
                 Case 1000, 100, 200, 300, 400, 500
                     Me.feActivos.BackColorRow (&H80000000)
                     Me.feActivos.ForeColorRow vbBlack, True
                 Case 102, 106, 301, 302, 303, 304, 401, 402, 403, 404, 405 'Colorear color amarillo
                     Me.feActivos.BackColorRow &HC0FFFF
                     Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
                 Case Else
                     Me.feActivos.BackColorRow (&HFFFFFF)
                     Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
                 End Select
                 
            rsDatActivos.MoveNext
        Loop
    rsDatActivos.Close
    Set rsDatActivos = Nothing
    
    'Call FormatearGrillas(fePasivo)
    Call LimpiaFlex(fePasivos)
        Do While Not rsDatPasivos.EOF
            fePasivos.AdicionaFila
            lnFila = fePasivos.row
            fePasivos.TextMatrix(lnFila, 1) = rsDatPasivos!cConsDescripcion
            fePasivos.TextMatrix(lnFila, 2) = Format(rsDatPasivos!PP, "#,##0.00")
            fePasivos.TextMatrix(lnFila, 3) = Format(rsDatPasivos!PE, "#,##0.00")
            fePasivos.TextMatrix(lnFila, 4) = Format(rsDatPasivos!nTotal, "#,##0.00")
            fePasivos.TextMatrix(lnFila, 5) = rsDatPasivos!nConsCod
            fePasivos.TextMatrix(lnFila, 6) = rsDatPasivos!nConsValor
            
            Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6)) 'celda que  o se puede editar
                Case 1000, 1001, 1002, 100, 200, 300, 400, 500
                    Me.fePasivos.BackColorRow (&H80000000)
                    Me.fePasivos.ForeColorRow vbBlack, True
                    Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
                Case 206
                    'Me.fePasivos.BackColorRow vbWhite, True        'CTI320200110 ERS003-2020. Comentó
                    'Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X" 'CTI320200110 ERS003-2020. Comentó
                    'CTI320200110 ERS003-2020. Agregó
                    If CDbl((fePasivos.TextMatrix(lnFila, 2)) > 0) Or CDbl((fePasivos.TextMatrix(lnFila, 3)) > 0) Then
                        Me.fePasivos.ForeColorRow vbBlack, True
                        Me.fePasivos.BackColorRow vbWhite, True
                        Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
                    Else
                        Me.fePasivos.BackColorRow vbWhite, True
                        Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
                        Me.fePasivos.RowHeight(lnFila) = 1
                    End If
                    'Fin CTI320200110 ERS003-2020
                                        
                Case 104, 107, 109, 201, 203 'Color Amarillo
                    Me.fePasivos.BackColorRow &HC0FFFF
                    Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
                Case Else
                    Me.fePasivos.BackColorRow (&HFFFFFF)
                    Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
            End Select
            rsDatPasivos.MoveNext
        Loop
    rsDatPasivos.Close
    Set rsDatPasivos = Nothing
    'LUCV20160626, Fin Carga Flex <-**********
    
        'Carga de rsDatIfiGastoNego -> Matrix
        ReDim MatIfiGastoNego(rsDatIfiGastoNego.RecordCount, 4)
        i = 0
        Do While Not rsDatIfiGastoNego.EOF
            MatIfiGastoNego(i, 0) = rsDatIfiGastoNego!nNroCuota
            MatIfiGastoNego(i, 1) = rsDatIfiGastoNego!CDescripcion
            MatIfiGastoNego(i, 2) = Format(IIf(IsNull(rsDatIfiGastoNego!nMonto), 0, rsDatIfiGastoNego!nMonto), "#,##0.00")
            rsDatIfiGastoNego.MoveNext
              i = i + 1
        Loop
        rsDatIfiGastoNego.Close
        Set rsDatIfiGastoNego = Nothing

        'Carga de rsDatIfiGastoFami -> Matrix
        ReDim MatIfiGastoFami(rsDatIfiGastoFami.RecordCount, 4)
        j = 0
        Do While Not rsDatIfiGastoFami.EOF
            MatIfiGastoFami(j, 0) = rsDatIfiGastoFami!nNroCuota
            MatIfiGastoFami(j, 1) = rsDatIfiGastoFami!CDescripcion
            MatIfiGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiGastoFami!nMonto), 0, rsDatIfiGastoFami!nMonto), "#,##0.00")
            rsDatIfiGastoFami.MoveNext
            j = j + 1
        Loop
        rsDatIfiGastoFami.Close
        Set rsDatIfiGastoFami = Nothing
    
    'CTI320200110 ERS003-2020. Agregó:
        '(Carga de rsDatIfiNoSupervisadaGastoNego -> Matrix)
        ReDim MatIfiNoSupervisadaGastoNego(rsDatIfiNoSupervisadaGastoNego.RecordCount, 4)
        i = 0
        Do While Not rsDatIfiNoSupervisadaGastoNego.EOF
            MatIfiNoSupervisadaGastoNego(i, 0) = rsDatIfiNoSupervisadaGastoNego!nNroCuota
            MatIfiNoSupervisadaGastoNego(i, 1) = rsDatIfiNoSupervisadaGastoNego!CDescripcion
            MatIfiNoSupervisadaGastoNego(i, 2) = Format(IIf(IsNull(rsDatIfiNoSupervisadaGastoNego!nMonto), 0, rsDatIfiNoSupervisadaGastoNego!nMonto), "#0.00")
            rsDatIfiNoSupervisadaGastoNego.MoveNext
              i = i + 1
        Loop
        rsDatIfiNoSupervisadaGastoNego.Close
        Set rsDatIfiNoSupervisadaGastoNego = Nothing
        
        'Carga de rsDatIfiNoSupervisadaGastoFami -> Matrix
        ReDim MatIfiNoSupervisadaGastoFami(rsDatIfiNoSupervisadaGastoFami.RecordCount, 4)
        j = 0
        Do While Not rsDatIfiNoSupervisadaGastoFami.EOF
            MatIfiNoSupervisadaGastoFami(j, 0) = rsDatIfiNoSupervisadaGastoFami!nNroCuota
            MatIfiNoSupervisadaGastoFami(j, 1) = rsDatIfiNoSupervisadaGastoFami!CDescripcion
            MatIfiNoSupervisadaGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiNoSupervisadaGastoFami!nMonto), 0, rsDatIfiNoSupervisadaGastoFami!nMonto), "#0.00")
            rsDatIfiNoSupervisadaGastoFami.MoveNext
        j = j + 1
        Loop
        rsDatIfiNoSupervisadaGastoFami.Close
        Set rsDatIfiNoSupervisadaGastoFami = Nothing
        'Fin CTI320200110 ERS003-2020
    
    
    
    'LUCV20160628, Para CARGA RATIOS/INDICADORES
    txtCapacidadNeta.Text = CStr(rsDatRatioInd!nCapPagNeta * 100) & "%"
    txtEndeudamiento.Text = CStr(rsDatRatioInd!nEndeuPat * 100) & "%"
    txtLiquidezCte.Text = CStr(Format(rsDatRatioInd!nLiquidezCte, "#0.00"))
    txtRentabilidadPat.Text = CStr(rsDatRatioInd!nRentaPatri * 100) & "%"
    txtIngresoNeto.Text = Format(rsDatRatioInd!nIngreNeto, "#,##0.00")
    txtExcedenteMensual.Text = Format(rsDatRatioInd!nExceMensual, "#,##0.00")
    
    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    If Not (rsDatParamFlujoCajaForm5.BOF And rsDatParamFlujoCajaForm5.EOF) Then
        EditMoneyForm5IVC.Text = Format(rsDatParamFlujoCajaForm5!nIncVentCont, "#,##0.00")
        EditMoneyForm5ICM.Text = Format(rsDatParamFlujoCajaForm5!nIncCompMerc, "#,##0.00")
        EditMoneyForm5IPP.Text = Format(rsDatParamFlujoCajaForm5!nIncPagPers, "#,##0.00")
        EditMoneyForm5IGV.Text = Format(rsDatParamFlujoCajaForm5!nIncGastvent, "#,##0.00")
        EditMoneyForm5IC.Text = Format(rsDatParamFlujoCajaForm5!nIncConsu, "#,##0.00")
    End If
    Set rsDatParamFlujoCajaForm5 = Nothing
    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    
    'CTI320200110 ERS003-2020
        If (fnProducto = "800") And CDbl(rsDatRatioInd!nEndeuPat) <= 0 And CDbl(rsDatRatioInd!nLiquidezCte) <= 0 And CDbl(rsDatRatioInd!nRentaPatri) <= 0 Then
            'Balance
            SSTabIngresos5.TabVisible(0) = False
            
            'Ratios
            Me.lblEndeudamiento.Visible = False
            Me.txtEndeudamiento.Visible = False
            Me.lblEndeAceptable.Visible = False
            
            Me.lblRentabilidad.Visible = False
            Me.txtRentabilidadPat.Visible = False
            Me.lblLiquidez.Visible = False
            Me.txtLiquidezCte.Visible = False
            Me.Line1.Visible = False
        End If
    'Fin CTI320200110
Set rsDCredito = Nothing
End Function

Private Sub GeneraVerCar()
    Dim oCred As COMNCredito.NCOMFormatosEval
    Dim oDCredSbs As COMDCredito.DCOMFormatosEval
    Dim R As ADODB.Recordset
    Dim lcDNI, lcRUC As String
    Dim RSbs, RDatFin1, RCap As ADODB.Recordset
    
    Set oCred = New COMNCredito.NCOMFormatosEval
    Call oCred.RecuperaDatosInformeComercial(ActXCodCta.NroCuenta, R)
    Set oCred = Nothing
    
    If R.EOF And R.BOF Then
    MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
    Exit Sub
    End If
    
    lcDNI = Trim(R!dni_deudor)
    lcRUC = Trim(R!ruc_deudor)
    
    Set oDCredSbs = New COMDCredito.DCOMFormatosEval
    Set RSbs = oDCredSbs.RecuperaCaliSbs(lcDNI, lcRUC)
    Set RDatFin1 = oDCredSbs.RecuperaDatosFinan(ActXCodCta.NroCuenta, fnFormato)
    Set oDCredSbs = Nothing
    Call ImprimeInformeCriteriosAceptacionRiesgoFormatoEval(ActXCodCta.NroCuenta, gsNomAge, gsCodUser, R, RSbs, RDatFin1)
End Sub

Private Sub ImprimirFormatoEvaluacion()
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsInfVisita As ADODB.Recordset
    Dim oDoc  As cPDF
    Dim psCtaCod As String
    
    Dim rsMostrarCuotasIfis As ADODB.Recordset
    Dim rsMostrarCuotasIfisGF As ADODB.Recordset
    Dim rsRatiosIndicadores As ADODB.Recordset
    
    Set oDoc = New cPDF
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsInfVisita = New ADODB.Recordset
    'Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(sCtaCod)
    Set rsInfVisita = oDCOMFormatosEval.MostrarFormatoSinConvenioInfVisCabecera(sCtaCod, fnFormato)
    
    Set rsMostrarCuotasIfis = oDCOMFormatosEval.MostrarCuotasIfis(sCtaCod, fnFormato, 7022)
    Set rsMostrarCuotasIfisGF = oDCOMFormatosEval.MostrarCuotasIfis(sCtaCod, fnFormato, 7023)
    Set rsRatiosIndicadores = oDCOMFormatosEval.RecuperaDatosRatios(sCtaCod)
    
    Dim a As Currency
    Dim nFila As Integer
    
    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Informe de Visita Nº " & sCtaCod
    oDoc.Title = "Informe de Visita Nº " & sCtaCod
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\FormatoEvaluacion_" & sCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    'Contenido
    oDoc.Fonts.Add "F1", "Courier New", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Courier New", TrueType, Bold, WinAnsiEncoding
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
        
    If Not (rsInfVisita.BOF Or rsInfVisita.EOF) Then
        'Tamaño de hoja A4
        oDoc.NewPage A4_Vertical
    
            '---------- cabecera
        oDoc.WImage 45, 45, 45, 113, "Logo"
        oDoc.WTextBox 40, 60, 35, 390, UCase(rsInfVisita!cAgeDescripcion), "F2", 7.5, hLeft
    
        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
        oDoc.WTextBox 60, 450, 10, 410, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
        oDoc.WTextBox 70, 450, 10, 490, "ANALISTA: " & UCase(Trim(rsInfVisita!cUser)), "F1", 7.5, hLeft
          
        oDoc.WTextBox 80, 100, 10, 400, "HOJA DE EVALUACION", "F2", 10, hCenter
        oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & Trim(rsInfVisita!cCtaCod), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsInfVisita!cPersCod), "F1", 7.5, hLeft
        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsInfVisita!cPersNombre), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 450, 10, 200, "DNI: " & Trim(rsInfVisita!cPersDni) & "   ", "F1", 7.5, hLeft
        oDoc.WTextBox 110, 450, 10, 200, "RUC: " & Trim(IIf(rsInfVisita!cPersRuc = "-", Space(11), rsInfVisita!cPersRuc)), "F1", 7.5, hLeft
        '-------------- fin cabecera
    
        '----------------------------------------------------------------------------------------------------------------
        oDoc.WTextBox 120, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        oDoc.WTextBox 130, 55, 1, 160, "ACTIVOS", "F2", 7.5, hjustify
        oDoc.WTextBox 140, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = 140
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "P.P", "F2", 7.5, hRight
        oDoc.WTextBox nFila, 240, 1, 160, "P.E.", "F2", 7.5, hRight
        oDoc.WTextBox nFila, 340, 1, 160, "TOTAL", "F2", 7.5, hRight
        
            For i = 1 To feActivos.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feActivos.TextMatrix(i, 1), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(feActivos.TextMatrix(i, 2), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 250, 15, 150, Format(feActivos.TextMatrix(i, 3), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 350, 15, 150, Format(feActivos.TextMatrix(i, 4), "#,#0.00"), "F1", 7.5, hRight
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
    
        '----------------------------------------------------------------------------------------------------------------
        oDoc.WTextBox nFila, 55, 1, 160, "PASIVOS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "P.P", "F2", 7.5, hRight
        oDoc.WTextBox nFila, 240, 1, 160, "P.E.", "F2", 7.5, hRight
        oDoc.WTextBox nFila, 340, 1, 160, "TOTAL", "F2", 7.5, hRight
        
            For i = 1 To fePasivos.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, fePasivos.TextMatrix(i, 1), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(fePasivos.TextMatrix(i, 2), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 250, 15, 150, Format(fePasivos.TextMatrix(i, 3), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 350, 15, 150, Format(fePasivos.TextMatrix(i, 4), "#,#0.00"), "F1", 7.5, hRight
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        '----------------------------------------------------------------------------------------------------------------
        
        oDoc.NewPage A4_Vertical
            '---------- cabecera
        oDoc.WImage 45, 45, 45, 113, "Logo"
        oDoc.WTextBox 40, 60, 35, 390, UCase(rsInfVisita!cAgeDescripcion), "F2", 7.5, hLeft
    
        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
        oDoc.WTextBox 60, 450, 10, 410, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
        oDoc.WTextBox 70, 450, 10, 490, "ANALISTA: " & UCase(Trim(rsInfVisita!cUser)), "F1", 7.5, hLeft
          
        oDoc.WTextBox 80, 100, 10, 400, "HOJA DE EVALUACION", "F2", 10, hCenter
        oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & Trim(rsInfVisita!cCtaCod), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsInfVisita!cPersCod), "F1", 7.5, hLeft
        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsInfVisita!cPersNombre), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 450, 10, 200, "DNI: " & Trim(rsInfVisita!cPersDni) & "   ", "F1", 7.5, hLeft
        oDoc.WTextBox 110, 450, 10, 200, "RUC: " & Trim(IIf(rsInfVisita!cPersRuc = "-", Space(11), rsInfVisita!cPersRuc)), "F1", 7.5, hLeft
    
        nFila = 140
        '----------------------------------------------------------------------------------------------------------------
        
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "FLUJO DE CAJA MENSUAL", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
        a = 0
            For i = 1 To feFlujoCajaMensual.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feFlujoCajaMensual.TextMatrix(i, 3), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(feFlujoCajaMensual.TextMatrix(i, 4), "#,#0.00"), "F1", 7.5, hRight
                a = a + feFlujoCajaMensual.TextMatrix(i, 4)
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(a, "#,##0.00"), "F2", 7.5, hRight
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        
        oDoc.WTextBox nFila, 55, 1, 200, "FLUJO DE CAJA MENSUAL - CUOTAS IFIS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        
        a = 0
        If Not (rsMostrarCuotasIfis.BOF And rsMostrarCuotasIfis.EOF) Then
            For i = 1 To rsMostrarCuotasIfis.RecordCount
                'oDoc.WTextBox nFila, 55, 1, 160, rsMostrarCuotasIfis!nNroCuota, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 55, 1, 300, rsMostrarCuotasIfis!CDescripcion, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 140, 1, 160, Format(rsMostrarCuotasIfis!nMonto, "#,##0.00"), "F1", 7.5, hRight
                a = a + rsMostrarCuotasIfis!nMonto
                rsMostrarCuotasIfis.MoveNext
                nFila = nFila + 10
            Next i
            'nFila = nFila + 10
                oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(a, "#,##0.00"), "F2", 7.5, hRight
         End If
        
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        
        '----------------------------------------------------------------------------------------------------------------
        oDoc.WTextBox nFila, 55, 1, 160, "GASTOS FAMILIARES", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
        a = 0
            For i = 1 To feGastosFamiliares.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feGastosFamiliares.TextMatrix(i, 2), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(feGastosFamiliares.TextMatrix(i, 3), "#,#0.00"), "F1", 7.5, hRight
                a = a + feGastosFamiliares.TextMatrix(i, 3)
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(a, "#,##0.00"), "F2", 7.5, hRight
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
    
    
        oDoc.WTextBox nFila, 55, 1, 160, "GASTOS FAMILIARES  - CUOTAS IFIS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        a = 0
        If Not (rsMostrarCuotasIfisGF.BOF And rsMostrarCuotasIfisGF.EOF) Then
            For i = 1 To rsMostrarCuotasIfisGF.RecordCount
                'oDoc.WTextBox nFila, 55, 1, 160, rsMostrarCuotasIfisGF!nNroCuota, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 55, 1, 300, rsMostrarCuotasIfisGF!CDescripcion, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 140, 1, 160, Format(rsMostrarCuotasIfisGF!nMonto, "#,##0.00"), "F1", 7.5, hRight
                a = a + rsMostrarCuotasIfisGF!nMonto
                nFila = nFila + 10
                rsMostrarCuotasIfisGF.MoveNext
            Next i
            'nFila = nFila + 10
                oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(a, "#,##0.00"), "F2", 7.5, hRight
         End If
        
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
    
        '----------------------------------------------------------------------------------------------------------------
        oDoc.WTextBox nFila, 55, 1, 160, "OTROS INGRESOS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
        a = 0
            For i = 1 To feOtrosIngresos.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feOtrosIngresos.TextMatrix(i, 2), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(feOtrosIngresos.TextMatrix(i, 3), "#,#0.00"), "F1", 7.5, hRight
                a = a + feOtrosIngresos.TextMatrix(i, 3)
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(a, "#,##0.00"), "F2", 7.5, hRight
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
    
        '----------------------------------------------------------------------------------------------------------------
        
        If nFila >= 700 Then
            
            'Tamaño de hoja A4
        oDoc.NewPage A4_Vertical
        
        oDoc.WImage 45, 45, 45, 113, "Logo"
        oDoc.WTextBox 40, 60, 35, 390, UCase(rsInfVisita!cAgeDescripcion), "F2", 7.5, hLeft
    
        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
        oDoc.WTextBox 60, 450, 10, 410, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
        oDoc.WTextBox 70, 450, 10, 490, "ANALISTA: " & UCase(Trim(rsInfVisita!cUser)), "F1", 7.5, hLeft
          
        oDoc.WTextBox 80, 100, 10, 400, "HOJA DE EVALUACION", "F2", 10, hCenter
        oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & Trim(rsInfVisita!cCtaCod), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsInfVisita!cPersCod), "F1", 7.5, hLeft
        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsInfVisita!cPersNombre), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 450, 10, 200, "DNI: " & Trim(rsInfVisita!cPersDni) & "   ", "F1", 7.5, hLeft
        oDoc.WTextBox 110, 450, 10, 200, "RUC: " & Trim(IIf(rsInfVisita!cPersRuc = "-", Space(11), rsInfVisita!cPersRuc)), "F1", 7.5, hLeft
            
        
        nFila = 110
        
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "DECLARACION PDT", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140 - 80, 1, 160, feDeclaracionPDT.TextMatrix(0, 4), "F2", 7.5, hRight
        oDoc.WTextBox nFila, 240 - 80, 1, 160, feDeclaracionPDT.TextMatrix(0, 5), "F2", 7.5, hRight
        oDoc.WTextBox nFila, 340 - 80, 1, 160, feDeclaracionPDT.TextMatrix(0, 6), "F2", 7.5, hRight
        oDoc.WTextBox nFila, 440 - 80, 1, 160, "MONTO", "F2", 7.5, hRight
        
            For i = 1 To feDeclaracionPDT.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feDeclaracionPDT.TextMatrix(i, 1), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 4), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 250 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 5), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 350 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 6), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 450 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 7), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 550 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 8), "#,#0.00"), "F1", 7.5, hRight
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
    
        '----------------------------------------------------------------------------------------------------------------
        oDoc.WTextBox nFila, 55, 1, 160, "RATIOS E INDICADORES", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "Capacidad de pago", "F1", 7.5, hjustify
'        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 10, 55, 1, 160, "Endeudamiento Patrimonial", "F1", 7.5, hjustify
'            oDoc.WTextBox nFila + 20, 55, 1, 160, "Liquidez Cte.", "F1", 7.5, hjustify
'            oDoc.WTextBox nFila + 30, 55, 1, 160, "Rentabilidad Pat.", "F1", 7.5, hjustify
'        End If
'        oDoc.WTextBox nFila + 40, 55, 1, 160, "Ingreso Neto Empresarial", "F1", 7.5, hjustify
'        oDoc.WTextBox nFila + 50, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
'
'        oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
'        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 10, 150, 15, 150, CStr(rsRatiosIndicadores!nEndeuPat * 100) & "%", "F1", 7.5, hRight
'            oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nLiquidezCte, "#,#0.00"), "F1", 7.5, hRight
'            oDoc.WTextBox nFila + 30, 150, 15, 150, CStr(rsRatiosIndicadores!nRentaPatri * 100) & "%", "F1", 7.5, hRight
'        End If
'        oDoc.WTextBox nFila + 40, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
'        oDoc.WTextBox nFila + 50, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
'
'        oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
'        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 10, 320, 1, 250, "EN RELACION A SU PATRIMONIO TOTAL", "F1", 7.5, hLeft
'        End If
        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
            oDoc.WTextBox nFila + 10, 55, 1, 160, "Endeudamiento Patrimonial", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 20, 55, 1, 160, "Liquidez Cte", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 30, 55, 1, 160, "Rentabilidad Pat.", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 40, 55, 1, 160, "Ingreso Neto Empresarial", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 50, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
            oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 10, 150, 15, 150, CStr(rsRatiosIndicadores!nEndeuPat * 100) & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nLiquidezCte, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila + 30, 150, 15, 150, CStr(rsRatiosIndicadores!nRentaPatri * 100) & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 40, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila + 50, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
            oDoc.WTextBox nFila + 10, 320, 1, 250, "EN RELACION A SU PATRIMONIO TOTAL", "F1", 7.5, hLeft
        Else
             oDoc.WTextBox nFila + 10, 55, 1, 160, "Ingreso Neto Empresarial", "F1", 7.5, hjustify
             oDoc.WTextBox nFila + 20, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
             oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
             oDoc.WTextBox nFila + 10, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
             oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
             oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
        End If
        '----------------------------------------------------------------------------------------------------------------
        Else
        
        
            oDoc.WTextBox nFila, 55, 1, 160, "DECLARACION PDT", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140 - 80, 1, 160, feDeclaracionPDT.TextMatrix(0, 4), "F2", 7.5, hRight
        oDoc.WTextBox nFila, 240 - 80, 1, 160, feDeclaracionPDT.TextMatrix(0, 5), "F2", 7.5, hRight
        oDoc.WTextBox nFila, 340 - 80, 1, 160, feDeclaracionPDT.TextMatrix(0, 6), "F2", 7.5, hRight
        oDoc.WTextBox nFila, 440 - 80, 1, 160, "MONTO", "F2", 7.5, hRight
        
            For i = 1 To feDeclaracionPDT.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feDeclaracionPDT.TextMatrix(i, 1), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 4), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 250 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 5), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 350 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 6), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 450 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 7), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 550 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 8), "#,#0.00"), "F1", 7.5, hRight
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
    
        '----------------------------------------------------------------------------------------------------------------
        oDoc.WTextBox nFila, 55, 1, 160, "RATIOS E INDICADORES", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "Capacidad de pago", "F1", 7.5, hjustify
'        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 10, 55, 1, 160, "Endeudamiento Patrimonial", "F1", 7.5, hjustify
'            oDoc.WTextBox nFila + 20, 55, 1, 160, "Liquidez Cte.", "F1", 7.5, hjustify
'            oDoc.WTextBox nFila + 30, 55, 1, 160, "Rentabilidad Pat.", "F1", 7.5, hjustify
'        End If
'        oDoc.WTextBox nFila + 40, 55, 1, 160, "Ingreso Neto Empresarial", "F1", 7.5, hjustify
'        oDoc.WTextBox nFila + 50, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
'
'        oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
'        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 10, 150, 15, 150, CStr(rsRatiosIndicadores!nEndeuPat * 100) & "%", "F1", 7.5, hRight
'            oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nLiquidezCte, "#,#0.00"), "F1", 7.5, hRight
'            oDoc.WTextBox nFila + 30, 150, 15, 150, CStr(rsRatiosIndicadores!nRentaPatri * 100) & "%", "F1", 7.5, hRight
'        End If
'        oDoc.WTextBox nFila + 40, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
'        oDoc.WTextBox nFila + 50, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
'
'        oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
'        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 10, 320, 1, 250, "EN RELACION A SU PATRIMONIO TOTAL", "F1", 7.5, hLeft
'        End If
        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
            oDoc.WTextBox nFila + 10, 55, 1, 160, "Endeudamiento Patrimonial", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 20, 55, 1, 160, "Liquidez Cte", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 30, 55, 1, 160, "Rentabilidad Pat.", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 40, 55, 1, 160, "Ingreso Neto Empresarial", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 50, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
            oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 10, 150, 15, 150, CStr(rsRatiosIndicadores!nEndeuPat * 100) & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nLiquidezCte, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila + 30, 150, 15, 150, CStr(rsRatiosIndicadores!nRentaPatri * 100) & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 40, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila + 50, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
            oDoc.WTextBox nFila + 10, 320, 1, 250, "EN RELACION A SU PATRIMONIO TOTAL", "F1", 7.5, hLeft
        Else
             oDoc.WTextBox nFila + 10, 55, 1, 160, "Ingreso Neto Empresarial", "F1", 7.5, hjustify
             oDoc.WTextBox nFila + 20, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
             oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
             oDoc.WTextBox nFila + 10, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
             oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
             oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
        End If
        '----------------------------------------------------------------------------------------------------------------
        End If
    
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "Los Datos de la propuesta del Credito no han sido Registrados Correctamente", vbInformation, "Aviso"
    End If
    Set rsInfVisita = Nothing
End Sub


