VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredMntGastos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Gastos"
   ClientHeight    =   8160
   ClientLeft      =   2370
   ClientTop       =   945
   ClientWidth     =   7725
   Icon            =   "frmCredMntGastos.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   75
      TabIndex        =   11
      Top             =   -45
      Width           =   7545
      Begin VB.ComboBox CboProd 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   180
         Width           =   2025
      End
      Begin VB.ComboBox CboAplicado 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   165
         Width           =   1830
      End
      Begin VB.Label Label10 
         Caption         =   "Producto :"
         Height          =   225
         Left            =   90
         TabIndex        =   15
         Top             =   225
         Width           =   780
      End
      Begin VB.Label Label9 
         Caption         =   "Aplicado A :"
         Height          =   225
         Left            =   4275
         TabIndex        =   13
         Top             =   195
         Width           =   1050
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   465
      Left            =   6390
      TabIndex        =   10
      Top             =   7515
      Width           =   1215
   End
   Begin VB.Frame fraDetalle 
      Height          =   4695
      Left            =   90
      TabIndex        =   8
      Top             =   2610
      Width           =   7530
      Begin TabDlg.SSTab SSComision 
         Height          =   4380
         Left            =   120
         TabIndex        =   16
         Top             =   225
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   7726
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Datos de Comision"
         TabPicture(0)   =   "frmCredMntGastos.frx":030A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "FraDatos"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "CmdFiltraProducto"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Valor de la Comision"
         TabPicture(1)   =   "frmCredMntGastos.frx":0326
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "FRaAplMonto"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "FraTipoValor"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "FraAplProceso"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "FraRestric"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Frame4"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "fraFactor"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "fraPagoConvenio"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).ControlCount=   7
         TabCaption(2)   =   "Restricciones"
         TabPicture(2)   =   "frmCredMntGastos.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraFacTit"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Frame3"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Frame2"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "CmdFiltroGarant"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "FraEdad"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "FraApl"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "FraGarant"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "FraAplDias"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).ControlCount=   8
         Begin VB.Frame fraPagoConvenio 
            Caption         =   "Convenios"
            Height          =   735
            Left            =   1860
            TabIndex        =   113
            Top             =   3556
            Width           =   2040
            Begin VB.ComboBox cbConvenio 
               Height          =   315
               ItemData        =   "frmCredMntGastos.frx":035E
               Left            =   135
               List            =   "frmCredMntGastos.frx":036B
               Style           =   2  'Dropdown List
               TabIndex        =   114
               Top             =   270
               Width           =   1500
            End
         End
         Begin VB.Frame fraFactor 
            Caption         =   " Multiplicar por el Número De"
            Height          =   1275
            Left            =   3960
            TabIndex        =   107
            Top             =   3015
            Width           =   3255
            Begin VB.CheckBox chkAplNumConMic 
               Caption         =   "Consulta Score Microfinanzas"
               Height          =   255
               Left            =   120
               TabIndex        =   111
               Top             =   740
               Width           =   2415
            End
            Begin VB.CheckBox chkAplNumMeses 
               Caption         =   "Meses Para la Primera Cuota"
               Height          =   255
               Left            =   120
               TabIndex        =   109
               Top             =   480
               Width           =   2445
            End
            Begin VB.CheckBox chkAplNumConCer 
               Caption         =   "Consultas Certicom"
               Height          =   255
               Left            =   120
               TabIndex        =   108
               Top             =   240
               Width           =   2565
            End
         End
         Begin VB.Frame fraFacTit 
            Caption         =   "Para Mas de Un Titular"
            Height          =   1005
            Left            =   -70200
            TabIndex        =   103
            Top             =   1530
            Width           =   2175
            Begin VB.TextBox txtValorDosTit 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   660
               MaxLength       =   6
               TabIndex        =   105
               Top             =   240
               Width           =   1410
            End
            Begin VB.CheckBox chkAplValorDosTit 
               Caption         =   "Habilitar esta Restriccion"
               Height          =   315
               Left            =   60
               TabIndex        =   104
               Top             =   600
               Width           =   2070
            End
            Begin VB.Label lblValorDos 
               Caption         =   "Valor:"
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
               Height          =   240
               Left            =   120
               TabIndex        =   106
               Top             =   240
               Width           =   645
            End
         End
         Begin VB.Frame Frame4 
            Height          =   525
            Left            =   1860
            TabIndex        =   98
            Top             =   3000
            Width           =   2040
            Begin VB.CheckBox ChkAplProcMens 
               Caption         =   "Mensualmente"
               Height          =   255
               Left            =   75
               TabIndex        =   99
               Top             =   180
               Width           =   1365
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Aplicable todos los dias"
            Height          =   1020
            Left            =   -72450
            TabIndex        =   90
            Top             =   2535
            Width           =   2820
            Begin VB.CheckBox ChkHabDiasVen 
               Caption         =   "Habilitar esta Restriccion"
               Height          =   315
               Left            =   60
               TabIndex        =   93
               Top             =   645
               Width           =   2070
            End
            Begin VB.TextBox TxtNroDiasVen 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1755
               TabIndex        =   92
               Text            =   "0"
               Top             =   285
               Width           =   315
            End
            Begin VB.ComboBox CboTodosDiasApl 
               Enabled         =   0   'False
               Height          =   315
               Left            =   105
               Style           =   2  'Dropdown List
               TabIndex        =   91
               Top             =   255
               Width           =   1455
            End
            Begin VB.Label Label19 
               Caption         =   "dias de Vencido"
               Height          =   420
               Left            =   2130
               TabIndex        =   95
               Top             =   210
               Width           =   600
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "a"
               Height          =   195
               Left            =   1605
               TabIndex        =   94
               Top             =   300
               Width           =   90
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Tipo de Gasto"
            Height          =   1245
            Left            =   -74940
            TabIndex        =   86
            Top             =   2535
            Width           =   2400
            Begin VB.OptionButton OptGastoFijoVar 
               Caption         =   "No Considerar"
               Enabled         =   0   'False
               Height          =   225
               Index           =   2
               Left            =   105
               TabIndex        =   89
               Top             =   840
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton OptGastoFijoVar 
               Caption         =   "Gasto Variable"
               Enabled         =   0   'False
               Height          =   225
               Index           =   1
               Left            =   105
               TabIndex        =   88
               Top             =   540
               Width           =   1335
            End
            Begin VB.OptionButton OptGastoFijoVar 
               Caption         =   "Gasto Fijo"
               Enabled         =   0   'False
               Height          =   225
               Index           =   0
               Left            =   105
               TabIndex        =   87
               Top             =   300
               Width           =   1005
            End
         End
         Begin VB.CommandButton CmdFiltraProducto 
            Height          =   465
            Left            =   -68670
            Picture         =   "frmCredMntGastos.frx":03A0
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Filtrar los Gastos para Ciertos Productos"
            Top             =   2895
            Width           =   510
         End
         Begin VB.CommandButton CmdFiltroGarant 
            Height          =   465
            Left            =   -68460
            Picture         =   "frmCredMntGastos.frx":09CA
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Filtrar los Gastos para Ciertas Garantias"
            Top             =   3135
            Width           =   510
         End
         Begin VB.Frame FraEdad 
            Caption         =   "Por Edad"
            Height          =   1080
            Left            =   -70515
            TabIndex        =   74
            Top             =   420
            Width           =   2520
            Begin VB.ComboBox CboEdadOper 
               Enabled         =   0   'False
               Height          =   315
               Left            =   105
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   255
               Width           =   1455
            End
            Begin VB.TextBox TxtEdad 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1710
               TabIndex        =   76
               Text            =   "0"
               Top             =   285
               Width           =   315
            End
            Begin VB.CheckBox ChkEdad 
               Caption         =   "Habilitar esta Restriccion"
               Height          =   315
               Left            =   135
               TabIndex        =   75
               Top             =   660
               Width           =   2070
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "a"
               Height          =   195
               Left            =   1575
               TabIndex        =   85
               Top             =   300
               Width           =   90
            End
            Begin VB.Label Label16 
               Caption         =   "años."
               Height          =   240
               Left            =   2055
               TabIndex        =   77
               Top             =   315
               Width           =   390
            End
         End
         Begin VB.Frame FraRestric 
            Caption         =   "Aplicar Cuando el Monto de la Operacion "
            Height          =   2640
            Left            =   3960
            TabIndex        =   61
            Top             =   345
            Width           =   3255
            Begin VB.CheckBox ChkHabRest 
               Caption         =   "Habilitar esta Restriccion"
               Height          =   435
               Left            =   150
               TabIndex        =   72
               Top             =   2115
               Width           =   1305
            End
            Begin VB.TextBox TxtOperPorc 
               Enabled         =   0   'False
               Height          =   285
               Left            =   750
               TabIndex        =   70
               Text            =   "0.00"
               Top             =   705
               Width           =   540
            End
            Begin VB.ComboBox CboOperador 
               Enabled         =   0   'False
               Height          =   315
               Left            =   750
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   315
               Width           =   2415
            End
            Begin VB.Frame FraDel 
               Caption         =   "Del"
               Enabled         =   0   'False
               Height          =   1830
               Left            =   1560
               TabIndex        =   62
               Top             =   735
               Width           =   1620
               Begin VB.OptionButton OptOperMonto 
                  Caption         =   "Saldo PrePago Hip"
                  Height          =   435
                  Index           =   4
                  Left            =   120
                  TabIndex        =   73
                  Tag             =   "C"
                  Top             =   1365
                  Width           =   1395
               End
               Begin VB.OptionButton OptOperMonto 
                  Caption         =   "Saldo"
                  Height          =   195
                  Index           =   3
                  Left            =   105
                  TabIndex        =   66
                  Tag             =   "C"
                  Top             =   1089
                  Width           =   1395
               End
               Begin VB.OptionButton OptOperMonto 
                  Caption         =   "Prestamo"
                  Height          =   195
                  Index           =   0
                  Left            =   105
                  TabIndex        =   65
                  Tag             =   "D"
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   1395
               End
               Begin VB.OptionButton OptOperMonto 
                  Caption         =   "Cuota"
                  Height          =   195
                  Index           =   1
                  Left            =   105
                  TabIndex        =   64
                  Tag             =   "C"
                  Top             =   543
                  Width           =   1395
               End
               Begin VB.OptionButton OptOperMonto 
                  Caption         =   "Garantia"
                  Height          =   195
                  Index           =   2
                  Left            =   105
                  TabIndex        =   63
                  Tag             =   "C"
                  Top             =   816
                  Width           =   1395
               End
            End
            Begin VB.Label Label14 
               Caption         =   "%"
               Height          =   240
               Left            =   1350
               TabIndex        =   71
               Top             =   735
               Width           =   255
            End
            Begin VB.Label Label13 
               Caption         =   "Que el :"
               Height          =   240
               Left            =   120
               TabIndex        =   69
               Top             =   735
               Width           =   630
            End
            Begin VB.Label Label12 
               Caption         =   "Sea :"
               Height          =   240
               Left            =   105
               TabIndex        =   67
               Top             =   330
               Width           =   480
            End
         End
         Begin VB.Frame FraApl 
            Caption         =   "Aplicado"
            Height          =   2130
            Left            =   -74940
            TabIndex        =   52
            Top             =   405
            Width           =   2385
            Begin VB.OptionButton OptAplicado 
               Caption         =   "Todos los Desem y Cuotas"
               Height          =   210
               Index           =   0
               Left            =   60
               TabIndex        =   58
               Top             =   255
               Value           =   -1  'True
               Width           =   2235
            End
            Begin VB.OptionButton OptAplicado 
               Caption         =   "Todos los Desembolsos"
               Height          =   210
               Index           =   1
               Left            =   60
               TabIndex        =   57
               Top             =   570
               Width           =   2055
            End
            Begin VB.OptionButton OptAplicado 
               Caption         =   "Solo un Desembolso"
               Height          =   210
               Index           =   2
               Left            =   60
               TabIndex        =   56
               Top             =   885
               Width           =   1830
            End
            Begin VB.OptionButton OptAplicado 
               Caption         =   "Todas las Cuotas"
               Height          =   210
               Index           =   3
               Left            =   60
               TabIndex        =   55
               Top             =   1215
               Width           =   1605
            End
            Begin VB.OptionButton OptAplicado 
               Caption         =   "Solo una Cuota"
               Height          =   210
               Index           =   4
               Left            =   60
               TabIndex        =   54
               Top             =   1530
               Width           =   2265
            End
            Begin VB.OptionButton OptAplicado 
               Caption         =   "No Considerar"
               Height          =   210
               Index           =   5
               Left            =   60
               TabIndex        =   53
               Top             =   1830
               Width           =   2265
            End
         End
         Begin VB.Frame FraAplProceso 
            Caption         =   "Aplica en Proceso"
            Height          =   2640
            Left            =   1860
            TabIndex        =   48
            Top             =   360
            Width           =   2025
            Begin VB.OptionButton OptAplicable 
               Caption         =   "Asig. Manual"
               Height          =   225
               Index           =   6
               Left            =   60
               TabIndex        =   97
               Top             =   2130
               Width           =   1230
            End
            Begin VB.OptionButton OptAplicable 
               Caption         =   "Refinanciacion"
               Height          =   225
               Index           =   3
               Left            =   60
               TabIndex        =   96
               Top             =   1170
               Width           =   1380
            End
            Begin VB.OptionButton OptAplicable 
               Caption         =   "Cierre de Dia"
               Height          =   225
               Index           =   5
               Left            =   60
               TabIndex        =   60
               Top             =   1815
               Width           =   1230
            End
            Begin VB.OptionButton OptAplicable 
               Caption         =   "Pre - Pagos"
               Height          =   225
               Index           =   4
               Left            =   60
               TabIndex        =   59
               Top             =   1500
               Width           =   1230
            End
            Begin VB.OptionButton OptAplicable 
               Caption         =   "Aprobacion"
               Height          =   210
               Index           =   0
               Left            =   60
               TabIndex        =   51
               Top             =   270
               Value           =   -1  'True
               Width           =   1260
            End
            Begin VB.OptionButton OptAplicable 
               Caption         =   "Pagos"
               Height          =   210
               Index           =   1
               Left            =   60
               TabIndex        =   50
               Top             =   570
               Width           =   1260
            End
            Begin VB.OptionButton OptAplicable 
               Caption         =   "Cancel/Levant"
               Height          =   210
               Index           =   2
               Left            =   60
               TabIndex        =   49
               Top             =   870
               Width           =   1365
            End
         End
         Begin VB.Frame FraTipoValor 
            Caption         =   "Tipo de Valor"
            Height          =   1275
            Left            =   90
            TabIndex        =   41
            Top             =   3000
            Width           =   1710
            Begin VB.OptionButton OptVPersonalizado 
               Caption         =   "Personalizado"
               Height          =   225
               Left            =   105
               TabIndex        =   110
               Top             =   810
               Width           =   1300
            End
            Begin VB.OptionButton OptVValor 
               Caption         =   "Valor"
               Height          =   210
               Left            =   105
               TabIndex        =   43
               Top             =   330
               Value           =   -1  'True
               Width           =   705
            End
            Begin VB.OptionButton OptVPorcentaje 
               Caption         =   "Porcentaje"
               Height          =   225
               Left            =   105
               TabIndex        =   42
               Top             =   555
               Width           =   1110
            End
         End
         Begin VB.Frame FraGarant 
            Caption         =   "Relaciones Cred"
            Height          =   1080
            Left            =   -72450
            TabIndex        =   37
            Top             =   420
            Width           =   1860
            Begin VB.OptionButton OptGarTpo 
               Caption         =   "No Considerar"
               Height          =   210
               Index           =   2
               Left            =   60
               TabIndex        =   40
               Top             =   780
               Value           =   -1  'True
               Width           =   1380
            End
            Begin VB.OptionButton OptGarTpo 
               Caption         =   "Mancomunada"
               Height          =   210
               Index           =   1
               Left            =   60
               TabIndex        =   39
               Top             =   525
               Width           =   1380
            End
            Begin VB.OptionButton OptGarTpo 
               Caption         =   "Individual"
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   38
               Top             =   270
               Width           =   975
            End
         End
         Begin VB.Frame FraAplDias 
            Caption         =   "Dias Aplicable"
            Height          =   1005
            Left            =   -72450
            TabIndex        =   36
            Top             =   1530
            Width           =   2205
            Begin VB.TextBox TxtAplDias 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1065
               TabIndex        =   80
               Text            =   "0"
               Top             =   240
               Width           =   315
            End
            Begin VB.CheckBox ChkAplDias 
               Caption         =   "Habilitar esta Restriccion"
               Enabled         =   0   'False
               Height          =   315
               Left            =   60
               TabIndex        =   78
               Top             =   585
               Width           =   2070
            End
            Begin VB.Label Label17 
               Caption         =   "dias de Vencido"
               Height          =   360
               Left            =   1440
               TabIndex        =   81
               Top             =   165
               Width           =   690
            End
            Begin VB.Label Label11 
               Caption         =   "Aplicar en :"
               Height          =   240
               Left            =   165
               TabIndex        =   79
               Top             =   270
               Width           =   810
            End
         End
         Begin VB.Frame FraDatos 
            Height          =   3150
            Left            =   -74820
            TabIndex        =   21
            Top             =   390
            Width           =   6780
            Begin VB.CheckBox chkAplTipCamb 
               Caption         =   "Aplicar tipo cambio para cuentas en Dólares USD"
               Height          =   375
               Left            =   3240
               TabIndex        =   101
               Top             =   2280
               Width           =   2775
            End
            Begin VB.Frame Frame7 
               Caption         =   "Moneda"
               Height          =   495
               Left            =   90
               TabIndex        =   44
               Top             =   2235
               Width           =   2955
               Begin VB.OptionButton OptMDolares 
                  Caption         =   "Dolares USD"
                  Height          =   195
                  Left            =   1650
                  TabIndex        =   46
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.OptionButton OptMSoles 
                  Caption         =   "Nuevos Soles"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   45
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1290
               End
            End
            Begin VB.TextBox txtMontoMin 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1380
               MaxLength       =   10
               TabIndex        =   28
               Top             =   1410
               Width           =   1290
            End
            Begin VB.TextBox txtMontoMax 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   4695
               MaxLength       =   10
               TabIndex        =   27
               Top             =   1395
               Width           =   1290
            End
            Begin VB.TextBox TxtRangoI 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1380
               MaxLength       =   10
               TabIndex        =   26
               Top             =   1065
               Width           =   1290
            End
            Begin VB.TextBox TxtRangoF 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4695
               MaxLength       =   10
               TabIndex        =   25
               Top             =   1050
               Width           =   1290
            End
            Begin VB.TextBox TxtDesc 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1380
               MaxLength       =   50
               TabIndex        =   24
               Top             =   675
               Width           =   4605
            End
            Begin VB.TextBox TxtValor 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1380
               MaxLength       =   6
               TabIndex        =   23
               Top             =   1740
               Width           =   1290
            End
            Begin VB.TextBox txtCodigo 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C16A0B&
               Height          =   300
               Left            =   1380
               MaxLength       =   8
               TabIndex        =   22
               Top             =   315
               Width           =   1290
            End
            Begin VB.Label Label8 
               Caption         =   "Monto Maximo"
               Height          =   255
               Left            =   3480
               TabIndex        =   35
               Top             =   1425
               Width           =   1080
            End
            Begin VB.Label Label4 
               Caption         =   "Monto Minimo"
               Height          =   285
               Left            =   240
               TabIndex        =   34
               Top             =   1410
               Width           =   1050
            End
            Begin VB.Label Label7 
               Caption         =   "Rango Final :"
               Height          =   285
               Left            =   3480
               TabIndex        =   33
               Top             =   1080
               Width           =   960
            End
            Begin VB.Label Label5 
               Caption         =   "Rango Inicial:"
               Height          =   255
               Left            =   240
               TabIndex        =   32
               Top             =   1080
               Width           =   1080
            End
            Begin VB.Label Label3 
               Caption         =   "Valor:"
               Height          =   240
               Left            =   240
               TabIndex        =   31
               Top             =   1740
               Width           =   525
            End
            Begin VB.Label Label2 
               Caption         =   "Descripción:"
               Height          =   240
               Left            =   240
               TabIndex        =   30
               Top             =   720
               Width           =   1005
            End
            Begin VB.Label Label1 
               Caption         =   "Código:"
               Height          =   255
               Left            =   240
               TabIndex        =   29
               Top             =   390
               Width           =   630
            End
         End
         Begin VB.Frame FRaAplMonto 
            Caption         =   "Aplicar al Monto"
            Height          =   2640
            Left            =   90
            TabIndex        =   17
            Top             =   360
            Width           =   1725
            Begin VB.OptionButton OptAplMonto 
               Caption         =   "Saldo Cap. Ant"
               Height          =   405
               Index           =   6
               Left            =   150
               TabIndex        =   112
               Tag             =   "C"
               Top             =   2160
               Width           =   1425
            End
            Begin VB.OptionButton OptAplMonto 
               Caption         =   "Interes Comp. + Saldo Cap. Ant"
               Height          =   450
               Index           =   5
               Left            =   150
               TabIndex        =   102
               Tag             =   "C"
               Top             =   1695
               Width           =   1425
            End
            Begin VB.OptionButton OptAplMonto 
               Caption         =   "V. Constr."
               Height          =   210
               Index           =   4
               Left            =   150
               TabIndex        =   100
               Tag             =   "C"
               Top             =   1500
               Width           =   1065
            End
            Begin VB.OptionButton OptAplMonto 
               Caption         =   "Saldo Cap"
               Height          =   210
               Index           =   3
               Left            =   150
               TabIndex        =   47
               Tag             =   "C"
               Top             =   1185
               Width           =   1065
            End
            Begin VB.OptionButton OptAplMonto 
               Caption         =   "Prestamo"
               Height          =   225
               Index           =   0
               Left            =   150
               TabIndex        =   20
               Tag             =   "D"
               Top             =   255
               Value           =   -1  'True
               Width           =   1185
            End
            Begin VB.OptionButton OptAplMonto 
               Caption         =   "Cuota"
               Height          =   210
               Index           =   1
               Left            =   150
               TabIndex        =   19
               Tag             =   "C"
               Top             =   570
               Width           =   1065
            End
            Begin VB.OptionButton OptAplMonto 
               Caption         =   "Garantia"
               Height          =   210
               Index           =   2
               Left            =   150
               TabIndex        =   18
               Tag             =   "C"
               Top             =   885
               Width           =   1065
            End
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   15
         Left            =   390
         TabIndex        =   9
         Top             =   1575
         Width           =   135
      End
   End
   Begin MSDataGridLib.DataGrid DGGastos 
      Bindings        =   "frmCredMntGastos.frx":0FF4
      Height          =   2025
      Left            =   90
      TabIndex        =   0
      Top             =   570
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   3572
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      ForeColor       =   0
      HeadLines       =   2
      RowHeight       =   15
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "nPrdConceptoCod"
         Caption         =   "Codigo"
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
         DataField       =   "cDescripcion"
         Caption         =   "Descripcion"
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
         MarqueeStyle    =   3
         ScrollBars      =   2
         BeginProperty Column00 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5699.906
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraActualiza 
      Height          =   705
      Left            =   90
      TabIndex        =   4
      Top             =   7380
      Width           =   3915
      Begin VB.CommandButton CmdModificar 
         Caption         =   "&Modificar"
         Height          =   450
         Left            =   2610
         TabIndex        =   6
         ToolTipText     =   "Modificar un Dato"
         Top             =   165
         Width           =   1215
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   450
         Left            =   1350
         TabIndex        =   7
         ToolTipText     =   "Nuevo SubGrupo"
         Top             =   165
         Width           =   1215
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   435
         Left            =   45
         TabIndex        =   5
         ToolTipText     =   "Nuevo Grupo"
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.Frame fraConfirmar 
      Height          =   705
      Left            =   810
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   2640
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "G&rabar"
         Height          =   450
         Left            =   120
         TabIndex        =   3
         Top             =   165
         Width           =   1215
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   450
         Left            =   1335
         TabIndex        =   2
         Top             =   165
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCredMntGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RGastos As ADODB.Recordset
Private cmdEjecutar As Integer
Enum TInicioGastos
    InicioGastosActualizar = 1
    InicioGastosConsultar = 2
End Enum

Private Sub HabilitaMontoOPeracion(ByVal pbHabilita As Boolean)
    CboOperador.Enabled = pbHabilita
    TxtOperPorc.Enabled = pbHabilita
    FraDel.Enabled = pbHabilita
End Sub

Private Sub HabilitaActializacion(ByVal pbHabilita As Boolean)
    fraDetalle.Enabled = True
    FraDatos.Enabled = pbHabilita
    FRaAplMonto.Enabled = pbHabilita
    FraTipoValor.Enabled = pbHabilita
    FraAplProceso.Enabled = pbHabilita
    FraRestric.Enabled = pbHabilita
    FraApl.Enabled = pbHabilita
    FraGarant.Enabled = pbHabilita
    FraEdad.Enabled = pbHabilita
    FraAplDias.Enabled = pbHabilita
    fraFacTit.Enabled = pbHabilita
    fraPagoConvenio.Enabled = pbHabilita 'CTI2 20181226 ADD ERS075-2018
    CmdFiltroGarant.Enabled = pbHabilita
    
    DGGastos.Enabled = Not pbHabilita
    CmdModificar.Visible = Not pbHabilita
    fraConfirmar.Visible = pbHabilita
    fraActualiza.Visible = Not pbHabilita
    CmdSalir.Visible = Not pbHabilita
    txtCodigo.Enabled = Not pbHabilita
    CmdFiltraProducto.Enabled = Not pbHabilita
    CmdFiltroGarant.Enabled = Not pbHabilita
End Sub

Private Function ValidaDatos() As Boolean
    ValidaDatos = True
    'Si registro ya existe
    If cmdEjecutar = 1 Then
        If RGastos.RecordCount > 0 Then
            RGastos.Find " nPrdConceptoCod  = " & Trim(txtCodigo.Text), , , 1
            If Not RGastos.EOF Then
                ValidaDatos = False
                MsgBox "Gasto ya Existe", vbInformation, "Aviso"
                Exit Function
            End If
        End If
    End If
    
    If ChkHabRest.value = 1 Then
        If CboOperador.ListIndex = -1 Then
            MsgBox "Seleccione un Operador para la Restriccion", vbInformation, "Aviso"
            CboOperador.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If CDbl(TxtOperPorc.Text) <= 0 Then
            MsgBox "Porcentaje para la Operacion debe de ser Mayor que Cero", vbInformation, "Aviso"
            TxtOperPorc.SetFocus
            ValidaDatos = False
            Exit Function
        End If
                
    End If
    
    If ChkEdad.value = 1 Then
        If CInt(TxtEdad.Text) <= 0 Then
            MsgBox "Edad debe ser Mayor que Cero", vbInformation, "Aviso"
            TxtEdad.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
    
    If ChkAplDias.value = 1 Then
        If CInt(TxtAplDias.Text) <= 0 Then
            MsgBox "El numero de dias para Aplicar la comision debe ser Mayor que Cero", vbInformation, "Aviso"
            TxtAplDias.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
    
End Function

Public Sub Inicio(ByVal TipoInicio As TInicioGastos)
    gsOpeCod = 190210 'FRHU 20150326 Para la pista(Gastos - Mantenimiento)
    If TipoInicio = InicioGastosConsultar Then
        CmdNuevo.Enabled = False
        CmdModificar.Enabled = False
        CmdEliminar.Enabled = False
    End If
    
    Me.Caption = "Mantenimiento de Gastos y Comisiones" ' CTI2 20181229 ERS075-2018
    Me.Show 1
End Sub

Private Sub LimpiaFrameMontoOper()
    CboOperador.ListIndex = -1
    TxtOperPorc.Text = "0.00"
    OptAplMonto(0).value = True
End Sub
Private Sub LimpiaPantalla()
    txtCodigo.Text = ""
    TxtDesc.Text = ""
    TxtRangoI.Text = "0.00"
    TxtRangoF.Text = "0.00"
    txtMontoMax.Text = "0.00"
    txtMontoMin.Text = "0.00"
    TxtValor.Text = "0.00"
    OptAplicado(0).value = True
    OptVValor.value = True
    OptMSoles.value = True
    OptVPersonalizado.value = True  'peac 20071231
    
    OptAplicable(0).value = True
    OptGarTpo(2).value = True
    OptOperMonto(0).value = True
    OptAplMonto(0).value = True
    ChkHabRest.value = 0
    CboOperador.ListIndex = -1
    TxtOperPorc.Text = "0.00"
    ChkEdad.value = 0
    CboEdadOper.ListIndex = -1
    TxtEdad.Text = "0"
    ChkAplDias.value = 0
    TxtAplDias.Text = "0"
End Sub

Private Sub CargaControles()

Dim oGasto As COMDCredito.DCOMGasto
Dim rsApli As ADODB.Recordset
Dim rsProd As ADODB.Recordset
Dim rsOpe As ADODB.Recordset
Dim rsEdad As ADODB.Recordset
Dim rsDias As ADODB.Recordset

Set oGasto = New COMDCredito.DCOMGasto
Call oGasto.CargarControlesGastos(rsApli, rsProd, rsOpe, rsEdad, rsDias)
Set oGasto = Nothing

Call Llenar_Combo_con_Recordset(rsApli, CboAplicado)
Call Llenar_Combo_con_Recordset(rsProd, CboProd)
CboProd.ListIndex = 0
Call Llenar_Combo_con_Recordset(rsOpe, CboOperador)
Call Llenar_Combo_con_Recordset(rsEdad, CboEdadOper)
Call Llenar_Combo_con_Recordset(rsDias, CboTodosDiasApl)
'    Call CargaComboConstante(gColocCalendApl, CboAplicado, gCredFiltroCalendApl)
'    Call CargaComboConstante(gGastosProd, cboProd)
'    cboProd.ListIndex = 0
'    Call CargaComboConstante(6031, CboOperador, gCredFiltroGastosOperad)
'    Call CargaComboConstante(6031, CboEdadOper, gCredFiltroGastosOperad)
'    Call CargaComboConstante(6031, CboTodosDiasApl, gCredFiltroGastosOperad)
    
Set rsApli = Nothing
Set rsProd = Nothing
Set rsOpe = Nothing
Set rsEdad = Nothing
Set rsDias = Nothing
    
End Sub

Private Sub CargaGridGastos(ByVal pnProd As Integer)
Dim oGastos As COMDCredito.DCOMGasto

    Set oGastos = New COMDCredito.DCOMGasto
    Set RGastos = Nothing
    Set RGastos = oGastos.RecuperaGastosCabecera(pnProd)
    Set DGGastos.DataSource = RGastos
    DGGastos.Refresh
    Set oGastos = Nothing
    
End Sub

Private Sub cbConvenio_Click()
'    MsgBox (cbConvenio.ListIndex)
End Sub

Private Sub CboAplicado_Click()
    Select Case CInt(Trim(Right(CboAplicado.Text, 10)))
        Case gColocCalendAplDesembolso
            RGastos.Filter = " nAplicado = " & gColocConceptoAplTodosDC & " OR nAplicado = " & gColocConceptoAplDesembolso & " OR nAplicado = " & gColocConceptoAplTodosD
        Case gColocCalendAplCuota
            RGastos.Filter = " nAplicado = " & gColocConceptoAplTodosDC & " OR nAplicado = " & gColocConceptoAplCuota & " OR nAplicado = " & gColocConceptoAplTodosC
    End Select
    DGGastos.Refresh
End Sub

Private Sub CboOperador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      TxtOperPorc.SetFocus
    End If
End Sub

Private Sub cboProd_Click()
    Call CargaGridGastos(CInt(Trim(Right(CboProd.Text, 2))))
End Sub


Private Sub ChkAplDias_Click()
    TxtAplDias.Text = "0"
    If ChkAplDias.value = 1 Then
        TxtAplDias.Enabled = True
    Else
        TxtAplDias.Enabled = False
    End If
End Sub

Private Sub chkAplValorDosTit_Click()
    If chkAplValorDosTit.value = 1 Then
        txtValorDosTit.Enabled = True
    Else
        txtValorDosTit.Enabled = False
    End If
End Sub

Private Sub ChkEdad_Click()
    TxtEdad.Text = "0"
    If ChkEdad.value = 1 Then
        TxtEdad.Enabled = True
        CboEdadOper.Enabled = True
    Else
        TxtEdad.Enabled = False
        CboEdadOper.Enabled = False
    End If
End Sub

Private Sub ChkHabDiasVen_Click()
    If ChkHabDiasVen.value = 1 Then
        CboTodosDiasApl.Enabled = True
        TxtNroDiasVen.Enabled = True
    Else
        CboTodosDiasApl.Enabled = False
        TxtNroDiasVen.Enabled = False
    End If
End Sub

Private Sub ChkHabRest_Click()
    Call LimpiaFrameMontoOper
    If ChkHabRest.value = 1 Then
        Call HabilitaMontoOPeracion(True)
    Else
        Call HabilitaMontoOPeracion(False)
    End If
End Sub

Private Sub cmdCancelar_Click()
    HabilitaActializacion False
    Call DGGastos_RowColChange(0, 0)
End Sub

Private Sub CmdEliminar_Click()
Dim oGasto As COMDCredito.DCOMGasto
Dim objPista As COMManejador.Pista 'FRHU 20150326 INCIDENTE
    If MsgBox("Se Va ha Eliminar el Gasto?, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oGasto = New COMDCredito.DCOMGasto
        If oGasto.GastoAsignadoACredito(RGastos!nPrdConceptoCod) Then
            MsgBox "No se puede Eliminar Concepto porque esta siendo Usado", vbInformation, "Aviso"
            Exit Sub
        End If
        Call oGasto.EliminarGasto(RGastos!nPrdConceptoCod)
        Call CargaGridGastos(CInt(Trim(Right(CboProd.Text, 2))))
        'FRHU 20150326 INCIDENTE
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gEliminar, "Elimino Gasto", "Concepto: " & Trim(txtCodigo.Text), gCodigoPersona
        Set objPista = Nothing
        'FIN FRHU 20150326
    End If
End Sub

Private Sub CmdFiltraProducto_Click()
    frmCredGastosXProd.Inicio RGastos!nPrdConceptoCod, RGastos!cDescripcion, IIf(OptMSoles.value, 1, 2)
End Sub

Private Sub CmdFiltroGarant_Click()
    frmCredGastosXProd.Inicio RGastos!nPrdConceptoCod, RGastos!cDescripcion, IIf(OptMSoles.value, 1, 2), True
End Sub

Private Sub CmdGrabar_Click()
Dim oGasto As COMDCredito.DCOMGasto
Dim sAplTmp As String
Dim sCrit As String
Dim sFiltro As String
Dim sAplMonto As String
Dim sAplicProc As String
Dim sOperMonto As String
Dim nOper As Integer
Dim nOperEdad As Integer
Dim nEdad As Integer
Dim nPorc As Double
Dim sTpoGastoFV As String
Dim nProdConceptoCod As Long
Dim sMontoMens As String
Dim objPista As COMManejador.Pista 'FRHU 20150326 INCIDENTE

    If MsgBox("Desea Guardar los Cambios ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    nProdConceptoCod = CLng(txtCodigo.Text)
    If Not ValidaDatos Then
        Exit Sub
    End If
    Set oGasto = New COMDCredito.DCOMGasto
    If OptAplicado(0).value Then
        sAplTmp = Format(gColocConceptoAplTodosDC, "0#")
    Else
        If OptAplicado(1).value Then
            sAplTmp = Format(gColocConceptoAplTodosD, "0#")
        Else
            If OptAplicado(2).value Then
                sAplTmp = Format(gColocConceptoAplDesembolso, "0#")
            Else
                If OptAplicado(3).value Then
                    sAplTmp = Format(gColocConceptoAplTodosC, "0#")
                Else
                    If OptAplicado(4).value Then
                        sAplTmp = Format(gColocConceptoAplCuota, "0#")
                    Else
                        If OptAplicado(5).value Then
                            sAplTmp = Format(gColocConceptoAplNingun, "0#")
                        End If
                    End If
                End If
            End If
        End If
    End If
    If OptAplMonto(0).value Then
        sAplMonto = "D"
    Else
        If Me.OptAplMonto(1).value Then
            sAplMonto = "C"
        Else
            If OptAplMonto(2).value Then
                sAplMonto = "H"
            Else
                If OptAplMonto(3).value Then
                    sAplMonto = "S"
                Else
                    If OptAplMonto(4).value Then
                        sAplMonto = "T"
                    Else
                        'JUEZ 20140310 ********************
                        If OptAplMonto(5).value Then
                            sAplMonto = "I" ' DAOR 20061213, se agregó esta opción que representa Interes Compensatorio + Saldo Capital Anterior
                        Else
                            sAplMonto = "K"
                        End If
                        'END JUEZ *************************
                    End If
                End If
            End If
        End If
    End If
    If OptGarTpo(0).value Then
        sFiltro = "I"
    Else
        If OptGarTpo(1).value Then
            sFiltro = "M"
        Else
            sFiltro = "N"
        End If
    End If
    
    If OptAplicable(0).value Then 'Desembolsos
        sAplicProc = "DE"
    Else
        If OptAplicable(1).value Then 'Pagos
            sAplicProc = "PA"
        Else
            If OptAplicable(2).value Then 'Cancelaciones
                sAplicProc = "CA"
            Else
                If OptAplicable(3).value Then 'Refinanciaciones
                    sAplicProc = "RE"
                Else
                    If OptAplicable(4).value Then 'Prepagos
                        sAplicProc = "PP"
                    Else
                        If OptAplicable(5).value Then 'Cierre de Dia
                            sAplicProc = "CD"
                        Else
                            If OptAplicable(6).value Then 'Asignacion Manual
                                sAplicProc = "MA"
                            Else
                                sAplicProc = ""
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If OptOperMonto(0).value Then 'Prestamo
        sOperMonto = "P"
    Else
        If OptOperMonto(1).value Then 'Cuota
            sOperMonto = "C"
        Else
            If OptOperMonto(2).value Then 'Garantia
                sOperMonto = "G"
            Else
                If OptOperMonto(3).value Then 'Saldo
                    sOperMonto = "S"
                Else
                    If OptOperMonto(4).value Then 'Saldo Hip Prepago
                        sOperMonto = "H"
                    Else
                        sOperMonto = ""
                    End If
                End If
            End If
        End If
    End If
    
    If ChkHabRest.value = 1 Then
        nOper = CInt(Trim(Right(CboOperador.Text, 2)))
        nPorc = CDbl(TxtOperPorc.Text)
    Else
        nOper = -1
        nPorc = -1
    End If
        
    If ChkEdad.value = 1 Then
        nOperEdad = CInt(Trim(Right(CboEdadOper.Text, 2)))
        nEdad = CInt(TxtEdad.Text)
    Else
        nOperEdad = -1
        nEdad = -1
    End If
       
    If OptGastoFijoVar(0).value Then
        sTpoGastoFV = "F"
    Else
        If OptGastoFijoVar(1).value Then
            sTpoGastoFV = "V"
        Else
            sTpoGastoFV = "N"
        End If
    End If
    
    If ChkAplProcMens.value = 1 And ChkAplProcMens.Enabled Then
        sMontoMens = "1"
    Else
        sMontoMens = "0"
    End If
    
   
    'DAOR 20061213, se aumentaron nuevos campos
    If cmdEjecutar = 1 Then

'peac 20080107
'            Call oGasto.NuevoGasto(Trim(txtCodigo.Text), Trim(txtDesc.Text), sAplTmp, _
'            Format(TxtRangoI, "#0.00"), Format(TxtRangoF, "#0.00"), _
'             IIf(OptVPorcentaje.value, "2", "1"), Format(TxtValor.Text, "#0.0000"), _
'            IIf(OptMSoles.value, gMonedaNacional, gMonedaExtranjera), CDbl(Me.TxtMontoMin.Text), _
'            CDbl(TxtMontoMax.Text), sAplMonto, sFiltro, sAplicProc, nOper, nPorc, sOperMonto, _
'            nEdad, IIf(ChkAplDias.value = 0, -1, CInt(TxtAplDias.Text)), _
'            CInt(Trim(Right(CboProd.Text, 2))), nOperEdad, sTpoGastoFV, CInt(IIf(ChkHabDiasVen.value = 1, Trim(Right(CboTodosDiasApl.Text, 2)), "-1")), _
'            CInt(IIf(ChkHabDiasVen.value = 1, Trim(TxtNroDiasVen.Text), "-1")), 0, sMontoMens, chkAplTipCamb.value, chkAplValorDosTit.value, _
'            Format(IIf(txtValorDosTit.Text = "", -1, txtValorDosTit.Text), "#0.0000"), chkAplNumConCer.value, chkAplNumMeses.value)
            
            'peac 20080107
            Call oGasto.NuevoGasto(Trim(txtCodigo.Text), Trim(TxtDesc.Text), sAplTmp, _
            Format(TxtRangoI, "#0.00"), Format(TxtRangoF, "#0.00"), _
            IIf(OptVPorcentaje.value, "2", IIf(OptVValor.value, "1", "3")), Format(TxtValor.Text, "#0.0000"), _
            IIf(OptMSoles.value, gMonedaNacional, gMonedaExtranjera), CDbl(Me.txtMontoMin.Text), _
            CDbl(txtMontoMax.Text), sAplMonto, sFiltro, sAplicProc, nOper, nPorc, sOperMonto, _
            nEdad, IIf(ChkAplDias.value = 0, -1, CInt(TxtAplDias.Text)), _
            CInt(Trim(Right(CboProd.Text, 2))), nOperEdad, sTpoGastoFV, CInt(IIf(ChkHabDiasVen.value = 1, Trim(Right(CboTodosDiasApl.Text, 2)), "-1")), _
            CInt(IIf(ChkHabDiasVen.value = 1, Trim(TxtNroDiasVen.Text), "-1")), 0, sMontoMens, chkAplTipCamb.value, chkAplValorDosTit.value, _
            Format(IIf(txtValorDosTit.Text = "", -1, txtValorDosTit.Text), "#0.0000"), chkAplNumConCer.value, chkAplNumMeses.value, chkAplNumConMic.value, _
            cbConvenio.ListIndex)
            
            
            
'        Call oGasto.NuevoGasto(Trim(txtCodigo.Text), Trim(TxtDesc.Text), sAplTmp, _
'            Format(TxtRangoI, "#0.00"), Format(TxtRangoF, "#0.00"), _
'             IIf(OptVPorcentaje.value, "2", "1"), Format(TxtValor.Text, "#0.0000"), _
'            IIf(OptMSoles.value, gMonedaNacional, gMonedaExtranjera), CDbl(Me.txtMontoMin.Text), _
'            CDbl(txtMontoMax.Text), sAplMonto, sFiltro, sAplicProc, nOper, nPorc, sOperMonto, _
'            nEdad, IIf(ChkAplDias.value = 0, -1, CInt(TxtAplDias.Text)), _
'            CInt(Trim(Right(CboProd.Text, 2))), nOperEdad, sTpoGastoFV, CInt(IIf(ChkHabDiasVen.value = 1, Trim(Right(CboTodosDiasApl.Text, 2)), "-1")), CInt(IIf(ChkHabDiasVen.value = 1, Trim(TxtNroDiasVen.Text), "-1")), 0, sMontoMens)
            
            'FRHU 20150326 INCIDENTE
            Set objPista = New COMManejador.Pista
            objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Inserta Gasto", "Concepto: " & Trim(txtCodigo.Text), gCodigoPersona
            Set objPista = Nothing
            'FIN FRHU 20150326
    Else
    
'peac 20080107
Call oGasto.ActualizarGasto(Trim(txtCodigo.Text), Trim(TxtDesc.Text), sAplTmp, _
Format(TxtRangoI, "#0.00"), Format(TxtRangoF, "#0.00"), _
IIf(OptVPorcentaje.value, "2", IIf(OptVValor.value, "1", "3")), Format(TxtValor.Text, "#0.0000"), _
IIf(OptMSoles.value, gMonedaNacional, gMonedaExtranjera), _
CDbl(Me.txtMontoMin.Text), CDbl(Me.txtMontoMax.Text), sAplMonto, sFiltro, sAplicProc, nOper, nPorc, sOperMonto, _
nEdad, IIf(ChkAplDias.value = 0, -1, CInt(TxtAplDias.Text)), nOperEdad, sTpoGastoFV, CInt(IIf(ChkHabDiasVen.value = 1, Trim(Right(CboTodosDiasApl.Text, 2)), "-1")), _
CInt(IIf(ChkHabDiasVen.value = 1, Trim(TxtNroDiasVen.Text), "-1")), 0, sMontoMens, chkAplTipCamb.value, chkAplValorDosTit.value, _
Format(IIf(txtValorDosTit.Text = "", -1, txtValorDosTit.Text), "#0.0000"), chkAplNumConCer.value, chkAplNumMeses.value, chkAplNumConMic.value, _
cbConvenio.ListIndex)
    
    
    'peac 20080107
'            Call oGasto.ActualizarGasto(Trim(txtCodigo.Text), Trim(TxtDesc.Text), sAplTmp, _
'            Format(TxtRangoI, "#0.00"), Format(TxtRangoF, "#0.00"), _
'            IIf(OptVPorcentaje.value, "2", "1"), Format(TxtValor.Text, "#0.0000"), _
'            IIf(OptMSoles.value, gMonedaNacional, gMonedaExtranjera), _
'            CDbl(Me.txtMontoMin.Text), CDbl(Me.txtMontoMax.Text), sAplMonto, sFiltro, sAplicProc, nOper, nPorc, sOperMonto, _
'            nEdad, IIf(ChkAplDias.value = 0, -1, CInt(TxtAplDias.Text)), nOperEdad, sTpoGastoFV, CInt(IIf(ChkHabDiasVen.value = 1, Trim(Right(CboTodosDiasApl.Text, 2)), "-1")), _
'            CInt(IIf(ChkHabDiasVen.value = 1, Trim(TxtNroDiasVen.Text), "-1")), 0, sMontoMens, chkAplTipCamb.value, chkAplValorDosTit.value, _
'            Format(IIf(txtValorDosTit.Text = "", -1, txtValorDosTit.Text), "#0.0000"), chkAplNumConCer.value, chkAplNumMeses.value)
            
'        Call oGasto.ActualizarGasto(Trim(txtCodigo.Text), Trim(TxtDesc.Text), sAplTmp, _
'            Format(TxtRangoI, "#0.00"), Format(TxtRangoF, "#0.00"), _
'            IIf(OptVPorcentaje.value, "2", "1"), Format(TxtValor.Text, "#0.0000"), _
'            IIf(OptMSoles.value, gMonedaNacional, gMonedaExtranjera), _
'            CDbl(Me.txtMontoMin.Text), CDbl(Me.txtMontoMax.Text), sAplMonto, sFiltro, sAplicProc, nOper, nPorc, sOperMonto, _
'            nEdad, IIf(ChkAplDias.value = 0, -1, CInt(TxtAplDias.Text)), nOperEdad, sTpoGastoFV, CInt(IIf(ChkHabDiasVen.value = 1, Trim(Right(CboTodosDiasApl.Text, 2)), "-1")), CInt(IIf(ChkHabDiasVen.value = 1, Trim(TxtNroDiasVen.Text), "-1")), 0, sMontoMens)

            'FRHU 20150326 INCIDENTE
            Set objPista = New COMManejador.Pista
            objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Modifico Gasto", "Concepto: " & Trim(txtCodigo.Text), gCodigoPersona
            Set objPista = Nothing
            'FIN FRHU 20150326
    End If
        
    Call CargaGridGastos(CInt(Trim(Right(CboProd.Text, 2))))
    If Not RGastos.EOF Then
        sCrit = "nPrdConceptoCod = " & nProdConceptoCod
        RGastos.Find sCrit
    End If
    Call HabilitaActializacion(False)
    DGGastos.SetFocus
                
            
End Sub

Private Sub cmdModificar_Click()
    
    cmdEjecutar = 2
    Call HabilitaActializacion(True)
    
End Sub

Private Sub cmdNuevo_Click()
Dim oGasto As COMDCredito.DCOMGasto

    Call LimpiaPantalla
    Call HabilitaActializacion(True)
    
    Set oGasto = New COMDCredito.DCOMGasto
        txtCodigo.Tag = oGasto.UltimoCorrelXProd + 1
        txtCodigo.Text = txtCodigo.Tag
    Set oGasto = Nothing
    txtCodigo.Enabled = True
    cmdEjecutar = 1
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub DGGastos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim sAplMonto As String
Dim sFiltro As String

    If Not RGastos.EOF Then
            txtCodigo.Text = RGastos!nPrdConceptoCod
            TxtDesc.Text = RGastos!cDescripcion
            TxtRangoI.Text = Format(RGastos!nInicial, "#0.00")
            TxtRangoF.Text = Format(RGastos!nFinal, "#0.00")
            TxtValor.Text = Format(RGastos!nValor, "#0.0000")
            txtMontoMax.Text = Format(IIf(IsNull(RGastos!nMontoMax), 0#, RGastos!nMontoMax), "#0.00")
            txtMontoMin.Text = Format(IIf(IsNull(RGastos!nMontoMin), 0#, RGastos!nMontoMin), "#0.00")
            Select Case CInt(RGastos!nAplicado)
                Case gColocConceptoAplCuota
                    OptAplicado(4).value = True
                Case gColocConceptoAplDesembolso
                    OptAplicado(2).value = True
                Case gColocConceptoAplTodosC
                    OptAplicado(3).value = True
                Case gColocConceptoAplTodosD
                    OptAplicado(1).value = True
                Case gColocConceptoAplTodosDC
                    OptAplicado(0).value = True
                Case gColocConceptoAplNingun
                    OptAplicado(5).value = True
            End Select
            
            If RGastos!nTpoValor = 1 Then
                OptVValor.value = True
            ElseIf RGastos!nTpoValor = 2 Then  'peac 20071231
                    OptVPorcentaje.value = True
                Else
                    OptVPersonalizado.value = True
            End If
            
            If RGastos!nMoneda = gMonedaNacional Then
                OptMSoles.value = True
            Else
                OptMDolares.value = True
            End If
            
            sFiltro = IIf(IsNull(RGastos!cFiltro), "N", RGastos!cFiltro)
            If sFiltro = "I" Then
                OptGarTpo(0).value = True
            Else
                If sFiltro = "M" Then
                    OptGarTpo(1).value = True
                Else
                    OptGarTpo(2).value = True
                End If
            End If
            
            If IsNull(RGastos!cAplicaProceso) Then
                OptAplicable(0).value = True
            Else
                Select Case RGastos!cAplicaProceso
                    Case "DE"
                        OptAplicable(0).value = True
                    Case "PA"
                        OptAplicable(1).value = True
                    Case "CA"
                        OptAplicable(2).value = True
                    Case "RE"
                        OptAplicable(3).value = True
                    Case "PP"
                        OptAplicable(4).value = True
                    Case "CD"
                        OptAplicable(5).value = True
                    Case "MA"
                        OptAplicable(6).value = True
                End Select
            End If
            
            If IsNull(RGastos!nOperador) Then
                CboOperador.ListIndex = -1
                ChkHabRest.value = 0
                TxtOperPorc.Enabled = False
                FraDel.Enabled = False
                OptOperMonto(0).value = True
            Else
                ChkHabRest.value = 1
                CboOperador.ListIndex = IndiceListaCombo(CboOperador, Trim(str(RGastos!nOperador)))
                TxtOperPorc.Text = Format(RGastos!nOperPorc, "#0.00")
                Select Case RGastos!cOperMonto
                    Case "P"
                        OptOperMonto(0).value = True
                    Case "C"
                        OptOperMonto(1).value = True
                    Case "G"
                        OptOperMonto(2).value = True
                    Case "S"
                        OptOperMonto(3).value = True
                    Case "H"
                        OptOperMonto(4).value = True
                End Select
            End If
            
            If IsNull(RGastos!nEdadOper) Then
                ChkEdad.value = 0
                CboEdadOper.ListIndex = -1
                CboEdadOper.Enabled = False
                TxtEdad.Text = "0"
                TxtEdad.Enabled = False
                
            Else
                ChkEdad.value = 1
                CboEdadOper.Enabled = True
                CboEdadOper.ListIndex = IndiceListaCombo(CboEdadOper, RGastos!nEdadOper)
                TxtEdad.Text = Trim(str(RGastos!nEdad))
                TxtEdad.Enabled = False
            End If
            
            If IsNull(RGastos!nDiasApl) Then
                ChkAplDias.Enabled = True
                ChkAplDias.value = 0
                TxtAplDias.Text = "0"
                TxtAplDias.Enabled = False
            Else
                ChkAplDias.Enabled = True
                ChkAplDias.value = 1
                TxtAplDias.Text = Trim(str(RGastos!nDiasApl))
                TxtAplDias.Enabled = True
            End If
            
            sAplMonto = IIf(IsNull(RGastos!cAplicaMonto), "C", RGastos!cAplicaMonto)
            If sAplMonto = "D" Then
                OptAplMonto(0).value = True
            Else
                If sAplMonto = "C" Then
                    OptAplMonto(1).value = True
                Else
                    If sAplMonto = "H" Then
                        OptAplMonto(2).value = True
                    Else
                        If sAplMonto = "S" Then
                            OptAplMonto(3).value = True
                        Else
                            If sAplMonto = "T" Then
                                OptAplMonto(4).value = True
                            Else
                                'JUEZ 20140310 ***********************
                                If sAplMonto = "I" Then
                                    OptAplMonto(5).value = True 'DAOR 20061214
                                Else
                                    OptAplMonto(6).value = True
                                End If
                                'END JUEZ ****************************
                            End If
                        End If
                    End If
                End If
            End If
            
            If IsNull(RGastos!cGastoFijoVar) Then
                OptGastoFijoVar(0).value = True
            Else
                If RGastos!cGastoFijoVar = "F" Then
                    OptGastoFijoVar(0).value = True
                Else
                    If RGastos!cGastoFijoVar = "V" Then
                        OptGastoFijoVar(1).value = True
                    Else
                        OptGastoFijoVar(2).value = True
                    End If
                End If
            End If
            
            If IsNull(RGastos!nOperDiasVenc) Then
                ChkHabDiasVen.value = 0
                CboTodosDiasApl.ListIndex = -1
                CboTodosDiasApl.Enabled = False
                TxtNroDiasVen.Text = "0"
                TxtNroDiasVen.Enabled = False
                
            Else
                ChkHabDiasVen.value = 1
                CboTodosDiasApl.Enabled = True
                CboTodosDiasApl.ListIndex = IndiceListaCombo(CboEdadOper, RGastos!nOperDiasVenc)
                TxtNroDiasVen.Text = Trim(str(RGastos!nDiasVenc))
                TxtNroDiasVen.Enabled = False
            End If
            
            ChkAplProcMens.value = CInt(IIf(IsNull(RGastos!nMontoMensual), 0, RGastos!nMontoMensual))
            cbConvenio.ListIndex = CInt(IIf(IsNull(RGastos!nAplicaConvenio), 0, RGastos!nAplicaConvenio))
            
            'DAOR 20061213
            chkAplTipCamb.value = IIf(IsNull(RGastos!bAplTipCamb), 0, IIf(RGastos!bAplTipCamb = True, 1, 0))
            chkAplValorDosTit.value = CInt(IIf(IsNull(RGastos!bAplValorDosTit), False, IIf(RGastos!bAplValorDosTit = True, 1, 0)))
            txtValorDosTit.Text = IIf(IsNull(RGastos!nValorDosTit), "", Format(RGastos!nValorDosTit, "#0.0000"))
            chkAplNumConCer.value = IIf(IsNull(RGastos!bAplNumConCer), 0, IIf(RGastos!bAplNumConCer = True, 1, 0))
            chkAplNumMeses.value = IIf(IsNull(RGastos!bAplNumMeses), 0, IIf(RGastos!bAplNumMeses = True, 1, 0))
            'Add by Gitu 03-06-09 Score Microfinanzas
            chkAplNumConMic.value = IIf(IsNull(RGastos!bAplNumConMic), 0, IIf(RGastos!bAplNumConMic = True, 1, 0))
    End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    CargaControles
    HabilitaActializacion False
    cmdEjecutar = -1
End Sub

Private Sub OptAplicable_Click(Index As Integer)
    If Index = 0 Or Index = 6 Then  'Desembolso
        'FraApl.Enabled = True
        'OptAplicado(0).value = True
        OptAplicado(0).Enabled = True
        OptAplicado(1).Enabled = True
        OptAplicado(2).Enabled = True
        OptAplicado(3).Enabled = True
        OptAplicado(4).Enabled = True
        OptAplicado(5).Enabled = True
        
        'FraGarant.Enabled = True
        OptGarTpo(0).Enabled = True
        OptGarTpo(1).Enabled = True
        OptGarTpo(2).Enabled = True
        
        If OptAplMonto(0).value And (OptGarTpo(0).value Or OptGarTpo(1).value) Then
            OptGastoFijoVar(2).value = True
            OptGastoFijoVar(0).Enabled = True
            OptGastoFijoVar(1).Enabled = True
            OptGastoFijoVar(2).Enabled = True
        End If
        
        ChkHabRest.value = 0
        FraRestric.Enabled = False
    Else
        'FraApl.Enabled = False
        'OptAplicado(0).value = True
        OptAplicado(0).Enabled = False
        OptAplicado(1).Enabled = False
        OptAplicado(2).Enabled = False
        OptAplicado(3).Enabled = False
        OptAplicado(4).Enabled = False
        OptAplicado(5).Enabled = False
        
        'FraGarant.Enabled = False
        OptGarTpo(0).Enabled = False
        OptGarTpo(1).Enabled = False
        OptGarTpo(2).Enabled = False
                
        OptGastoFijoVar(2).value = True
        OptGastoFijoVar(0).Enabled = False
        OptGastoFijoVar(1).Enabled = False
        OptGastoFijoVar(2).Enabled = False
        
        ChkHabRest.value = 0
        FraRestric.Enabled = True
    End If
    
    If Index = 5 Then 'Cierre de Dia
        'FraApl.Enabled = True
        ChkHabDiasVen.Enabled = True
        TxtNroDiasVen.Enabled = True
        CboAplicado.Enabled = True
    Else
        If Not (Index = 0 Or Index = 6) Then
            'FraApl.Enabled = False
        End If
        ChkHabDiasVen.Enabled = False
        TxtNroDiasVen.Enabled = False
        CboAplicado.Enabled = False
    End If
    
    If Index = 0 Then
        ChkAplProcMens.Enabled = True
        fraFactor.Enabled = True 'DAOR 20061216
    Else
        ChkAplProcMens.Enabled = False
        fraFactor.Enabled = False 'DAOR 20061216
    End If
    
    ChkHabDiasVen.Enabled = True
End Sub

Private Sub OptAplMonto_Click(Index As Integer)
    If Index = 0 Or Index = 3 Then
        If OptAplicable(0).value And (OptGarTpo(0).value Or OptGarTpo(1).value) Then
            OptGastoFijoVar(2).value = True
            OptGastoFijoVar(0).Enabled = True
            OptGastoFijoVar(1).Enabled = True
            OptGastoFijoVar(2).Enabled = True
        End If
    Else
        OptGastoFijoVar(2).value = True
        OptGastoFijoVar(0).Enabled = False
        OptGastoFijoVar(1).Enabled = False
        OptGastoFijoVar(2).Enabled = False
    End If
End Sub

Private Sub OptGarTpo_Click(Index As Integer)
    If Index = 0 Or Index = 1 Then
        If (OptAplMonto(0).value Or OptAplMonto(3).value) And OptAplicable(0).value Then
            OptGastoFijoVar(2).value = True
            OptGastoFijoVar(0).Enabled = True
            OptGastoFijoVar(1).Enabled = True
            OptGastoFijoVar(2).Enabled = True
        End If
    Else
            OptGastoFijoVar(2).value = True
            OptGastoFijoVar(0).Enabled = False
            OptGastoFijoVar(1).Enabled = False
            OptGastoFijoVar(2).Enabled = False
    End If
End Sub

'DAOR 20061213
Private Sub OptMDolares_Click()
    If OptMDolares.value Then
        chkAplTipCamb.Enabled = False
        chkAplTipCamb.value = 0
    End If
End Sub
'DAOR 20061213
Private Sub OptMSoles_Click()
    If OptMSoles.value Then
        chkAplTipCamb.Enabled = True
    End If
End Sub

Private Sub txtcodigo_GotFocus()
    fEnfoque txtCodigo
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        TxtDesc.SetFocus
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    If Trim(txtCodigo.Text) = "" Then
        txtCodigo.Text = "1201"
    End If
End Sub

Private Sub TxtDesc_GotFocus()
    fEnfoque TxtDesc
End Sub

Private Sub TxtDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        TxtRangoI.SetFocus
    End If
End Sub

Private Sub txtMontoMax_GotFocus()
    fEnfoque txtMontoMax
End Sub

Private Sub txtMontoMax_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMontoMax, KeyAscii)
    If KeyAscii = 13 Then
        TxtValor.SetFocus
    End If
End Sub

Private Sub txtMontoMax_LostFocus()
    If Trim(txtMontoMax.Text) = "" Then
        txtMontoMax.Text = "0.00"
    End If
    txtMontoMax.Text = Format(txtMontoMax.Text, "#0.00")
End Sub


Private Sub txtMontoMin_GotFocus()
    fEnfoque txtMontoMin
End Sub

Private Sub txtMontoMin_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMontoMin, KeyAscii)
    If KeyAscii = 13 Then
        txtMontoMax.SetFocus
    End If
End Sub

Private Sub txtMontoMin_LostFocus()
    If Trim(txtMontoMin.Text) = "" Then
        txtMontoMin.Text = "0.00"
    End If
    txtMontoMin.Text = Format(txtMontoMin.Text, "#0.00")
    
End Sub


Private Sub TxtOperPorc_GotFocus()
    fEnfoque TxtOperPorc
End Sub

Private Sub TxtOperPorc_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtOperPorc, KeyAscii, 5)
    
End Sub

Private Sub TxtOperPorc_LostFocus()
    TxtOperPorc.Text = Format(IIf(Trim(TxtOperPorc.Text) = "", "0.00", TxtOperPorc.Text), "0.00")
End Sub

Private Sub TxtRangoF_GotFocus()
    fEnfoque TxtRangoF
End Sub

Private Sub TxtRangoF_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtRangoF, KeyAscii)
    If KeyAscii = 13 Then
        txtMontoMin.SetFocus
    End If
End Sub

Private Sub TxtRangoF_LostFocus()
    If Trim(TxtRangoF.Text) = "" Then
        TxtRangoF.Text = "0.00"
    Else
        TxtRangoF.Text = Format(TxtRangoF.Text, "#0.00")
    End If
End Sub

Private Sub TxtRangoI_GotFocus()
    fEnfoque TxtRangoI
End Sub

Private Sub TxtRangoI_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtRangoI, KeyAscii)
    If KeyAscii = 13 Then
        TxtRangoF.SetFocus
    End If
End Sub

Private Sub TxtRangoI_LostFocus()
    If Trim(TxtRangoI.Text) = "" Then
        TxtRangoI.Text = "0.00"
    Else
        TxtRangoI.Text = Format(TxtRangoI.Text, "#0.00")
    End If
End Sub

Private Sub TxtValor_GotFocus()
    fEnfoque TxtValor
End Sub

Private Sub TxtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtValor, KeyAscii, , 4)
    If KeyAscii = 13 Then
        CmdGrabar.SetFocus
    End If
End Sub

Private Sub TxtValor_LostFocus()
    If Trim(TxtValor.Text) = "" Then
        TxtValor.Text = "0.00"
    Else
        TxtValor.Text = Format(TxtValor.Text, "#0.0000")
    End If
End Sub

