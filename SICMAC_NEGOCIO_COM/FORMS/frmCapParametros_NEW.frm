VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapParametros_NEW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Productos Pasivos - "
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "frmCapParametros_NEW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8281
      _Version        =   393216
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Ahorros"
      TabPicture(0)   =   "frmCapParametros_NEW.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboAhoTpoPrograma"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdAhoSeleccionar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FraAho"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Plazo Fijo"
      TabPicture(1)   =   "frmCapParametros_NEW.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label19"
      Tab(1).Control(1)=   "cmdPFSeleccionar"
      Tab(1).Control(2)=   "cboPFTpoPrograma"
      Tab(1).Control(3)=   "FraPF"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "CTS"
      TabPicture(2)   =   "frmCapParametros_NEW.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label20"
      Tab(2).Control(1)=   "cmdCTSSeleccionar"
      Tab(2).Control(2)=   "cboCTSTpoPrograma"
      Tab(2).Control(3)=   "FraCTS"
      Tab(2).ControlCount=   4
      Begin VB.Frame FraCTS 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   72
         Top             =   840
         Width           =   6495
         Begin VB.TextBox txtCTSPorcTangRet 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4200
            TabIndex        =   95
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtCTSMesesSinDepTransNoAct 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4200
            TabIndex        =   92
            Top             =   2400
            Width           =   855
         End
         Begin VB.TextBox txtCTSDiasVerifUltRegSueldo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4200
            TabIndex        =   75
            Top             =   2880
            Width           =   855
         End
         Begin VB.TextBox txtCTSUltRemunBrutas 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4200
            TabIndex        =   74
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox txtCTSCantOpeDepAnio 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4200
            TabIndex        =   73
            Top             =   480
            Width           =   855
         End
         Begin SICMACT.EditMoney txtCTSSueldoMin 
            Height          =   300
            Left            =   4200
            TabIndex        =   76
            Top             =   1440
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label Label35 
            Caption         =   "Días"
            Height          =   255
            Left            =   5160
            TabIndex        =   94
            Top             =   2925
            Width           =   615
         End
         Begin VB.Label Label32 
            Caption         =   "Meses"
            Height          =   255
            Left            =   5160
            TabIndex        =   93
            Top             =   2445
            Width           =   615
         End
         Begin VB.Label Label54 
            Caption         =   "Nº de Operaciones de depósito x Año:"
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   510
            Width           =   2805
         End
         Begin VB.Label Label47 
            Caption         =   "%"
            Height          =   255
            Left            =   5160
            TabIndex        =   82
            Top             =   1005
            Width           =   255
         End
         Begin VB.Label Label42 
            Caption         =   "Días verificación último registro sueldo CTS:"
            Height          =   255
            Left            =   240
            TabIndex        =   81
            Top             =   2925
            Width           =   3375
         End
         Begin VB.Label Label41 
            Caption         =   "Meses sin depósitos para traslado a CTS No Activo:"
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   2445
            Width           =   3855
         End
         Begin VB.Label Label40 
            Caption         =   "Factor últimas Remuneraciones Brutas:"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   1965
            Width           =   3015
         End
         Begin VB.Label Label39 
            Caption         =   "Sueldo Mínimo S/.:"
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   1470
            Width           =   1815
         End
         Begin VB.Label Label38 
            Caption         =   "Porcentaje Tangible de Retiro:"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   1005
            Width           =   2775
         End
      End
      Begin VB.Frame FraPF 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   48
         Top             =   840
         Width           =   6495
         Begin VB.TextBox txtPFMaxAumCapMes 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5520
            TabIndex        =   89
            Top             =   3120
            Width           =   855
         End
         Begin VB.OptionButton optAumCap 
            Caption         =   "Si"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   88
            Top             =   2775
            Width           =   495
         End
         Begin VB.OptionButton optAumCap 
            Caption         =   "No"
            Height          =   255
            Index           =   0
            Left            =   2160
            TabIndex        =   87
            Top             =   2775
            Width           =   615
         End
         Begin VB.CheckBox chkPFMedRetIniPlazo 
            Caption         =   "Inicio de Plazo"
            Height          =   195
            Left            =   2160
            TabIndex        =   86
            Top             =   2320
            Width           =   2175
         End
         Begin VB.CheckBox chkPFMedRetMensual 
            Caption         =   "Mensualmente"
            Height          =   195
            Left            =   2160
            TabIndex        =   85
            Top             =   2050
            Width           =   2175
         End
         Begin VB.CheckBox chkPFMedRetFinPlazo 
            Caption         =   "Al vencimiento del plazo"
            Height          =   195
            Left            =   2160
            TabIndex        =   84
            Top             =   1800
            Width           =   2175
         End
         Begin VB.CheckBox chkPFPersNat 
            Caption         =   "Persona Natural"
            Height          =   195
            Left            =   2160
            TabIndex        =   54
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox chkPFPersJur 
            Caption         =   "Persona Jurídica"
            Height          =   195
            Left            =   4320
            TabIndex        =   53
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox chkPFSol 
            Caption         =   "Nuevos Soles"
            Height          =   195
            Left            =   2160
            TabIndex        =   52
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox chkPFDol 
            Caption         =   "Dólares Americanos"
            Height          =   195
            Left            =   4320
            TabIndex        =   51
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtPFPlazoMin 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2520
            TabIndex        =   50
            Top             =   1330
            Width           =   1095
         End
         Begin VB.TextBox txtPFPlazoMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4800
            TabIndex        =   49
            Top             =   1335
            Width           =   975
         End
         Begin SICMACT.EditMoney txtPFMontoMinApertSol 
            Height          =   300
            Left            =   2520
            TabIndex        =   55
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin SICMACT.EditMoney txtPFMontoMinApertDol 
            Height          =   300
            Left            =   4560
            TabIndex        =   56
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin SICMACT.EditMoney txtPFAumCapSol 
            Height          =   300
            Left            =   4320
            TabIndex        =   57
            Top             =   2760
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin SICMACT.EditMoney txtPFAumCapDol 
            Height          =   300
            Left            =   5520
            TabIndex        =   58
            Top             =   2760
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin VB.Label Label27 
            Caption         =   "Min:"
            Height          =   255
            Left            =   2160
            TabIndex        =   91
            Top             =   1380
            Width           =   375
         End
         Begin VB.Label Label26 
            Caption         =   "Mínimo:"
            Height          =   255
            Left            =   3360
            TabIndex        =   90
            Top             =   2790
            Width           =   615
         End
         Begin VB.Label Label37 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   585
            Width           =   735
         End
         Begin VB.Label Label36 
            Caption         =   "Monto Mínimo Apertura:"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   1005
            Width           =   1815
         End
         Begin VB.Label Label34 
            Caption         =   "Medios de Retiro válidos:"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   1755
            Width           =   2175
         End
         Begin VB.Label Label33 
            Caption         =   "Plazo (días):"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   1380
            Width           =   2415
         End
         Begin VB.Label Label31 
            Caption         =   "Aumento de Capital:"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   2805
            Width           =   1935
         End
         Begin VB.Label Label30 
            Caption         =   "Nº Max Aumento Cap x Mes:"
            Height          =   255
            Left            =   3120
            TabIndex        =   66
            Top             =   3195
            Width           =   2055
         End
         Begin VB.Label Label29 
            Caption         =   "S/."
            Height          =   255
            Left            =   2160
            TabIndex        =   65
            Top             =   1005
            Width           =   255
         End
         Begin VB.Label Label28 
            Caption         =   "$."
            Height          =   255
            Left            =   4320
            TabIndex        =   64
            Top             =   1005
            Width           =   255
         End
         Begin VB.Label Label25 
            Caption         =   "Min:"
            Height          =   255
            Left            =   2160
            TabIndex        =   63
            Top             =   1380
            Width           =   735
         End
         Begin VB.Label Label24 
            Caption         =   "Max:"
            Height          =   255
            Left            =   4320
            TabIndex        =   62
            Top             =   1380
            Width           =   735
         End
         Begin VB.Label Label23 
            Caption         =   "S/."
            Height          =   255
            Left            =   3960
            TabIndex        =   61
            Top             =   2805
            Width           =   255
         End
         Begin VB.Label Label22 
            Caption         =   "$."
            Height          =   255
            Left            =   5280
            TabIndex        =   60
            Top             =   2805
            Width           =   255
         End
         Begin VB.Label Label21 
            Caption         =   "Personeria:"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame FraAho 
         Height          =   3735
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   6495
         Begin VB.ComboBox cboAhoAplicaDesemb 
            Height          =   315
            ItemData        =   "frmCapParametros_NEW.frx":035E
            Left            =   2880
            List            =   "frmCapParametros_NEW.frx":0368
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   3270
            Width           =   1455
         End
         Begin VB.TextBox txtAhoCantOpeCajeros 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3840
            TabIndex        =   23
            Top             =   2490
            Width           =   495
         End
         Begin VB.TextBox txtAhoCantOpeVentRet 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5520
            TabIndex        =   22
            Top             =   2115
            Width           =   495
         End
         Begin VB.TextBox txtAhoCantOpeVentDep 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3840
            TabIndex        =   21
            Top             =   2115
            Width           =   495
         End
         Begin VB.ComboBox cboAhoOrdPag 
            Height          =   315
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1725
            Width           =   1455
         End
         Begin VB.CheckBox chkAhoDol 
            Caption         =   "Dólares Americanos"
            Height          =   195
            Left            =   4560
            TabIndex        =   19
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox chkAhoSol 
            Caption         =   "Nuevos Soles"
            Height          =   195
            Left            =   2880
            TabIndex        =   18
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox chkAhoPersJur 
            Caption         =   "Persona Jurídica"
            Height          =   195
            Left            =   4560
            TabIndex        =   17
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox chkAhoPersNat 
            Caption         =   "Persona Natural"
            Height          =   195
            Left            =   2880
            TabIndex        =   16
            Top             =   240
            Width           =   1575
         End
         Begin SICMACT.EditMoney txtAhoMontoMinApertSol 
            Height          =   300
            Left            =   3240
            TabIndex        =   25
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin SICMACT.EditMoney txtAhoMontoMinApertDol 
            Height          =   300
            Left            =   4800
            TabIndex        =   26
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin SICMACT.EditMoney txtAhoMontoMinDepSol 
            Height          =   300
            Left            =   3240
            TabIndex        =   27
            Top             =   1335
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin SICMACT.EditMoney txtAhoMontoMinDepDol 
            Height          =   300
            Left            =   4800
            TabIndex        =   28
            Top             =   1335
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin SICMACT.EditMoney txtAhoSaldoMinSol 
            Height          =   300
            Left            =   3240
            TabIndex        =   29
            Top             =   2880
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin SICMACT.EditMoney txtAhoSaldoMinDol 
            Height          =   300
            Left            =   4800
            TabIndex        =   30
            Top             =   2880
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin VB.Label Label2 
            Caption         =   "Personeria:"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label18 
            Caption         =   "$."
            Height          =   255
            Left            =   4560
            TabIndex        =   46
            Top             =   2910
            Width           =   255
         End
         Begin VB.Label Label17 
            Caption         =   "S/."
            Height          =   255
            Left            =   2880
            TabIndex        =   45
            Top             =   2910
            Width           =   255
         End
         Begin VB.Label Label16 
            Caption         =   "Retiros"
            Height          =   255
            Left            =   4560
            TabIndex        =   44
            Top             =   2145
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "Depósitos"
            Height          =   255
            Left            =   2880
            TabIndex        =   43
            Top             =   2145
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "$."
            Height          =   255
            Left            =   4560
            TabIndex        =   42
            Top             =   1365
            Width           =   255
         End
         Begin VB.Label Label13 
            Caption         =   "S/."
            Height          =   255
            Left            =   2880
            TabIndex        =   41
            Top             =   1365
            Width           =   255
         End
         Begin VB.Label Label12 
            Caption         =   "$."
            Height          =   255
            Left            =   4560
            TabIndex        =   40
            Top             =   1005
            Width           =   255
         End
         Begin VB.Label Label11 
            Caption         =   "S/."
            Height          =   255
            Left            =   2880
            TabIndex        =   39
            Top             =   1005
            Width           =   255
         End
         Begin VB.Label Label10 
            Caption         =   "Aplica destino desembolso de crédito:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   3315
            Width           =   2775
         End
         Begin VB.Label Label9 
            Caption         =   "Saldo mínimo en cuenta:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   2910
            Width           =   1935
         End
         Begin VB.Label Label8 
            Caption         =   "Operaciones en cajero libres:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   2535
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "Operaciones en Ventanillas libres:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   2145
            Width           =   2415
         End
         Begin VB.Label Label6 
            Caption         =   "Orden de Pago:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1755
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Monto Mínimo Depósito:"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1365
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Monto Mínimo Apertura:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1005
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   585
            Width           =   735
         End
      End
      Begin VB.ComboBox cboCTSTpoPrograma 
         Height          =   315
         Left            =   -73920
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmdCTSSeleccionar 
         Caption         =   "Seleccionar"
         Height          =   330
         Left            =   -71160
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox cboPFTpoPrograma 
         Height          =   315
         Left            =   -73920
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmdPFSeleccionar 
         Caption         =   "Seleccionar"
         Height          =   330
         Left            =   -71160
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdAhoSeleccionar 
         Caption         =   "Seleccionar"
         Height          =   330
         Left            =   3840
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox cboAhoTpoPrograma 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label20 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   11
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   530
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCapParametros_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCapParametros_NEW
'** Descripción : Formulario para administrar la configuración de parámetros de cuentas pasivas
'** Creación : JUEZ, 20140908 09:00:00 AM
'*****************************************************************************************************

Option Explicit

Dim clsCapDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rs As ADODB.Recordset
Dim nProd As Producto
Dim nTipoOpe As Integer
Dim lbAumCap As Integer
Dim lnUltRemunBrutasAnt As Integer
Dim lnPorcTangRet As Double

Public Sub Mantenimiento(ByVal pnProducto As Producto)
nProd = pnProducto
nTipoOpe = 1
Select Case nProd
    Case gCapAhorros
        Me.Caption = "Mantenimiento " & Me.Caption & "Ahorros"
        CargaCombo cboAhoTpoPrograma, gCaptacSubProdAhorros
        CargaCombo cboAhoOrdPag, 2040
        FraAho.Enabled = False
        SSTab.TabVisible(1) = False
        SSTab.TabVisible(2) = False
    Case gCapPlazoFijo
        Me.Caption = "Mantenimiento " & Me.Caption & "Plazo Fijo"
        CargaCombo cboPFTpoPrograma, gCaptacSubProdPlazoFijo
        FraPF.Enabled = False
        SSTab.TabVisible(0) = False
        SSTab.TabVisible(2) = False
    Case gCapCTS
        Me.Caption = "Mantenimiento " & Me.Caption & "CTS"
        CargaCombo cboCTSTpoPrograma, gCaptacSubProdCTS
        FraCTS.Enabled = False
        SSTab.TabVisible(0) = False
        SSTab.TabVisible(1) = False
End Select
cmdEditar.Enabled = False
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
Me.Show 1
End Sub

Public Sub Consulta(ByVal pnProducto As Producto)
nProd = pnProducto
nTipoOpe = 2
Select Case nProd
    Case gCapAhorros
        Me.Caption = "Consulta " & Me.Caption & "Ahorros"
        CargaCombo cboAhoTpoPrograma, gCaptacSubProdAhorros
        CargaCombo cboAhoOrdPag, 2040
        FraAho.Enabled = False
        SSTab.TabVisible(1) = False
        SSTab.TabVisible(2) = False
    Case gCapPlazoFijo
        Me.Caption = "Consulta " & Me.Caption & "Plazo Fijo"
        CargaCombo cboPFTpoPrograma, gCaptacSubProdPlazoFijo
        FraPF.Enabled = False
        SSTab.TabVisible(0) = False
        SSTab.TabVisible(2) = False
    Case gCapCTS
        Me.Caption = "Consulta " & Me.Caption & "CTS"
        CargaCombo cboCTSTpoPrograma, gCaptacSubProdCTS
        FraCTS.Enabled = False
        SSTab.TabVisible(0) = False
        SSTab.TabVisible(1) = False
End Select
cmdEditar.Visible = False
cmdGrabar.Visible = False
cmdCancelar.Visible = False
Me.Show 1
End Sub

Private Sub CargaCombo(ByVal cboCombo As ComboBox, ByVal pnConstante As Integer)
    Dim rsConst As New ADODB.Recordset
    Dim clsGen As New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(pnConstante, , , 1)
    Set clsGen = Nothing
    
    cboCombo.Clear
    While Not rsConst.EOF
        cboCombo.AddItem rsConst.Fields(0) & Space(100) & rsConst.Fields(1)
        rsConst.MoveNext
    Wend
    'If pnConstante = gCaptacSubProdCTS Then
    '    cboCombo.RemoveItem (IndiceListaCombo(cboCombo, 2))
    'End If
End Sub

Function SoloNumeros(ByVal KeyAscii As Integer, Optional ByVal pbDecimal As Boolean = False) As Integer
    'permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
    If pbDecimal Then
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then
            SoloNumeros = 0
        Else
            SoloNumeros = KeyAscii
        End If
    Else
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then
            SoloNumeros = 0
        Else
            SoloNumeros = KeyAscii
        End If
    End If
    ' teclas especiales permitidas
    If KeyAscii = 8 Then SoloNumeros = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function

Private Sub chkAhoDol_Click()
If chkAhoDol.value Then
    txtAhoMontoMinApertDol.Enabled = True
    txtAhoMontoMinDepDol.Enabled = True
    txtAhoSaldoMinDol.Enabled = True
Else
    txtAhoMontoMinApertDol.Text = "0"
    txtAhoMontoMinDepDol.Text = "0"
    txtAhoSaldoMinDol.Text = "0"
    txtAhoMontoMinApertDol.Enabled = False
    txtAhoMontoMinDepDol.Enabled = False
    txtAhoSaldoMinDol.Enabled = False
End If
End Sub

Private Sub chkAhoSol_Click()
If chkAhoSol.value Then
    txtAhoMontoMinApertSol.Enabled = True
    txtAhoMontoMinDepSol.Enabled = True
    txtAhoSaldoMinSol.Enabled = True
Else
    txtAhoMontoMinApertSol.Text = "0"
    txtAhoMontoMinDepSol.Text = "0"
    txtAhoSaldoMinSol.Text = "0"
    txtAhoMontoMinApertSol.Enabled = False
    txtAhoMontoMinDepSol.Enabled = False
    txtAhoSaldoMinSol.Enabled = False
End If
End Sub

Private Sub chkPFDol_Click()
If chkPFDol.value Then
    txtPFMontoMinApertDol.Enabled = True
    If lbAumCap = 1 Then
        txtPFAumCapDol.Enabled = True
    End If
Else
    txtPFMontoMinApertDol.Text = "0"
    txtPFAumCapDol.Text = "0"
    txtPFMontoMinApertDol.Enabled = False
    txtPFAumCapDol.Enabled = False
End If
End Sub

Private Sub chkPFSol_Click()
If chkPFSol.value Then
    txtPFMontoMinApertSol.Enabled = True
    If lbAumCap = 1 Then
        txtPFAumCapSol.Enabled = True
    End If
Else
    txtPFMontoMinApertSol.Text = "0"
    txtPFAumCapSol.Text = "0"
    txtPFMontoMinApertSol.Enabled = False
    txtPFAumCapSol.Enabled = False
End If
End Sub

Private Sub cmdAhoSeleccionar_Click()
If Trim(Right(cboAhoTpoPrograma.Text, 2)) = "" Then Exit Sub
CargarParametroSubProducto
If nTipoOpe = 1 Then
    cboAhoTpoPrograma.Enabled = False
    cmdAhoSeleccionar.Enabled = False
    cmdEditar.Enabled = True
    cmdCancelar.Enabled = True
End If
End Sub

Private Sub CargarParametroSubProducto()
Set clsCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
Select Case nProd
    Case gCapAhorros
        Set rs = clsCapDef.GetCapParametroNew(nProd, CInt(Trim(Right(cboAhoTpoPrograma.Text, 2))))
        If Not rs.EOF Then
            chkAhoPersNat.value = IIf(rs!bPersNat, 1, 0)
            chkAhoPersJur.value = IIf(rs!bPersJur, 1, 0)
            chkAhoSol.value = IIf(rs!bMonSol, 1, 0)
            chkAhoDol.value = IIf(rs!bMonDol, 1, 0)
            txtAhoMontoMinApertSol.Text = Format(rs!nMontoMinApertSol, "#,##0.00")
            txtAhoMontoMinApertDol.Text = Format(rs!nMontoMinApertDol, "#,##0.00")
            txtAhoMontoMinDepSol.Text = Format(rs!nMontoMinDepSol, "#,##0.00")
            txtAhoMontoMinDepDol.Text = Format(rs!nMontoMinDepDol, "#,##0.00")
            cboAhoOrdPag.ListIndex = IndiceListaCombo(cboAhoOrdPag, rs!nOrdPago)
            txtAhoCantOpeVentDep.Text = rs!nCantOpeVentDep
            txtAhoCantOpeVentRet.Text = rs!nCantOpeVentRet
            txtAhoCantOpeCajeros.Text = rs!nCantOpeCajero
            txtAhoSaldoMinSol.Text = Format(rs!nSaldoMinCtaSol, "#,##0.00")
            txtAhoSaldoMinDol.Text = Format(rs!nSaldoMinCtaDol, "#,##0.00")
            cboAhoAplicaDesemb.ListIndex = IndiceListaCombo(cboAhoAplicaDesemb, IIf(rs!bAplicaDesembCred, 1, 0))
        Else
            Limpiar
        End If
    Case gCapPlazoFijo
        Set rs = clsCapDef.GetCapParametroNew(nProd, CInt(Trim(Right(cboPFTpoPrograma.Text, 2))))
        If Not rs.EOF Then
            chkPFPersNat.value = IIf(rs!bPersNat, 1, 0)
            chkPFPersJur.value = IIf(rs!bPersJur, 1, 0)
            chkPFSol.value = IIf(rs!bMonSol, 1, 0)
            chkPFDol.value = IIf(rs!bMonDol, 1, 0)
            txtPFMontoMinApertSol.Text = Format(rs!nMontoMinApertSol, "#,##0.00")
            txtPFMontoMinApertDol.Text = Format(rs!nMontoMinApertDol, "#,##0.00")
            txtPFPlazoMin.Text = rs!nPlazoMin
            txtPFPlazoMax.Text = rs!nPlazoMax
            chkPFMedRetFinPlazo.value = IIf(rs!bFormaRetFinPlazo, 1, 0)
            chkPFMedRetIniPlazo.value = IIf(rs!bFormaRetInicioPlazo, 1, 0)
            chkPFMedRetMensual.value = IIf(rs!bFormaRetMensual, 1, 0)
            optAumCap(IIf(rs!bAumCap, 1, 0)).value = True
            txtPFAumCapSol.Text = Format(rs!nAumCapMinSol, "#,##0.00")
            txtPFAumCapDol.Text = Format(rs!nAumCapMinDol, "#,##0.00")
            txtPFMaxAumCapMes.Text = rs!nAumCapCantMaxMes
        Else
            Limpiar
        End If
    Case gCapCTS
        lnUltRemunBrutasAnt = 0
        lnPorcTangRet = 0
        Set rs = clsCapDef.GetCapParametroNew(nProd, CInt(Trim(Right(cboCTSTpoPrograma.Text, 2))))
        If Not rs.EOF Then
            txtCTSCantOpeDepAnio.Text = rs!nCantOpeDepAnio
            txtCTSPorcTangRet.Text = Format(rs!nPorcTangRet, "#,##0.00")
            lnPorcTangRet = rs!nPorcTangRet
            txtCTSSueldoMin.Text = Format(rs!nSueldoMin, "#,##0.00")
            txtCTSUltRemunBrutas.Text = rs!nUltRemunBrutas
            lnUltRemunBrutasAnt = rs!nUltRemunBrutas
            txtCTSMesesSinDepTransNoAct.Text = rs!nMesesSinDepTransNoActivo
            txtCTSDiasVerifUltRegSueldo.Text = rs!nDiasVerifUltRegSueldo
        Else
            Limpiar
        End If
End Select
Set clsCapDef = Nothing
End Sub

Private Sub cmdCancelar_Click()
Limpiar
cmdEditar.Enabled = False
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
lbAumCap = 0
Select Case nProd
    Case gCapAhorros
        cboAhoTpoPrograma.Enabled = True
        cmdAhoSeleccionar.Enabled = True
        FraAho.Enabled = False
    Case gCapPlazoFijo
        cboPFTpoPrograma.Enabled = True
        cmdPFSeleccionar.Enabled = True
        FraPF.Enabled = False
    Case gCapCTS
        cboCTSTpoPrograma.Enabled = True
        cmdCTSSeleccionar.Enabled = True
        FraCTS.Enabled = False
End Select
End Sub

Private Sub Limpiar()
Select Case nProd
    Case gCapAhorros
        chkAhoPersNat.value = 0
        chkAhoPersJur.value = 0
        chkAhoSol.value = 0
        chkAhoDol.value = 0
        txtAhoMontoMinApertSol.Text = "0"
        txtAhoMontoMinApertDol.Text = "0"
        txtAhoMontoMinDepSol.Text = "0"
        txtAhoMontoMinDepDol.Text = "0"
        cboAhoOrdPag.ListIndex = -1
        txtAhoCantOpeVentDep.Text = "0"
        txtAhoCantOpeVentRet.Text = "0"
        txtAhoCantOpeCajeros.Text = "0"
        txtAhoSaldoMinSol.Text = "0"
        txtAhoSaldoMinDol.Text = "0"
        cboAhoAplicaDesemb.ListIndex = -1
    Case gCapPlazoFijo
        chkPFPersNat.value = 0
        chkPFPersJur.value = 0
        chkPFSol.value = 0
        chkPFDol.value = 0
        txtPFMontoMinApertSol.Text = "0"
        txtPFMontoMinApertDol.Text = "0"
        txtPFPlazoMin.Text = "0"
        txtPFPlazoMax.Text = "0"
        chkPFMedRetFinPlazo.value = 0
        chkPFMedRetIniPlazo.value = 0
        chkPFMedRetMensual.value = 0
        optAumCap(0).value = False
        optAumCap(1).value = False
        txtPFAumCapSol.Text = 0
        txtPFAumCapDol.Text = 0
        txtPFMaxAumCapMes.Text = 0
    Case gCapCTS
        txtCTSCantOpeDepAnio.Text = "0"
        txtCTSPorcTangRet.Text = "0"
        txtCTSSueldoMin.Text = "0"
        txtCTSUltRemunBrutas.Text = "0"
        txtCTSMesesSinDepTransNoAct.Text = "0"
        txtCTSDiasVerifUltRegSueldo.Text = "0"
End Select
End Sub

Private Sub cmdCTSSeleccionar_Click()
If Trim(Right(cboCTSTpoPrograma.Text, 2)) = "" Then Exit Sub
CargarParametroSubProducto
If nTipoOpe = 1 Then
    cboCTSTpoPrograma.Enabled = False
    cmdCTSSeleccionar.Enabled = False
    cmdEditar.Enabled = True
    cmdCancelar.Enabled = True
End If
End Sub

Private Sub CmdEditar_Click()
If nProd = gCapAhorros Then
    FraAho.Enabled = True
ElseIf nProd = gCapPlazoFijo Then
    FraPF.Enabled = True
ElseIf nProd = gCapCTS Then
    FraCTS.Enabled = True
End If
        
cmdEditar.Enabled = False
cmdGrabar.Enabled = True
End Sub

Private Sub cmdExportar_Click()
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim lsArchivo As String
Dim lsHoja As String
Dim xlHoja1 As Excel.Worksheet
Dim xlHoja2 As Excel.Worksheet
Dim nLin As Long
Dim nItem As Long
Dim sColumna As String
Dim sProd As String
    
    If nProd = gCapAhorros Then
        sProd = "Ahorros"
    ElseIf nProd = gCapPlazoFijo Then
        sProd = "Plazo Fijo"
    ElseIf nProd = gCapCTS Then
        sProd = "CTS"
    End If
    
    Set clsCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set rs = clsCapDef.GetCapParametroNew(nProd, -1)
    Set clsCapDef = Nothing
    
    If Not rs.EOF Then
        lsArchivo = App.path & "\SPOOLER\Parametros" & sProd & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".xls"
        lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
        If Not lbLibroOpen Then
            Exit Sub
        End If
        nLin = 1
        lsHoja = sProd
        gFunGeneral.ExcelAddHoja lsHoja, xlLibro, xlHoja1
        
        xlHoja1.Range("A1:Y99999").EntireColumn.Font.FontStyle = "Arial"
        xlHoja1.Range("A1:Y99999").EntireColumn.Font.Size = 10
        xlHoja1.PageSetup.CenterHorizontally = True
        xlHoja1.PageSetup.Zoom = 75
        xlHoja1.PageSetup.TopMargin = 2
        
        Select Case nProd
            Case gCapAhorros
                xlHoja1.Range("A1:A1").RowHeight = 30
                xlHoja1.Range("A2:A2").RowHeight = 18
                xlHoja1.Range("A1:A1").ColumnWidth = 18
                xlHoja1.Range("B1:O1").ColumnWidth = 8
                xlHoja1.Range("L1:M1").ColumnWidth = 11
                xlHoja1.Range("P1:P1").ColumnWidth = 11
                
                xlHoja1.Cells(nLin, 1) = "Producto"
                xlHoja1.Cells(nLin, 2) = "Personeria"
                xlHoja1.Cells(nLin, 4) = "Moneda"
                xlHoja1.Cells(nLin, 6) = "Monto mínimo" & Chr(10) & "apertura"
                xlHoja1.Cells(nLin, 8) = "Monto mínimo" & Chr(10) & "de depósito"
                xlHoja1.Cells(nLin, 10) = "Orden de " & Chr(10) & "Pago"
                xlHoja1.Cells(nLin, 11) = "Operaciones en" & Chr(10) & "ventanilla libres"
                xlHoja1.Cells(nLin, 13) = "Operaciones" & Chr(10) & "en cajero" & Chr(10) & "libres"
                xlHoja1.Cells(nLin, 14) = "Saldo mínimo" & Chr(10) & "en cuentas"
                xlHoja1.Cells(nLin, 16) = "Aplica" & Chr(10) & "desembolso" & Chr(10) & "de crédito"
                
                xlHoja1.Range("A1", "A2").MergeCells = True
                xlHoja1.Range("B1", "C1").MergeCells = True
                xlHoja1.Range("D1", "E1").MergeCells = True
                xlHoja1.Range("F1", "G1").MergeCells = True
                xlHoja1.Range("H1", "I1").MergeCells = True
                xlHoja1.Range("J1", "J2").MergeCells = True
                xlHoja1.Range("K1", "L1").MergeCells = True
                xlHoja1.Range("M1", "M2").MergeCells = True
                xlHoja1.Range("N1", "O1").MergeCells = True
                xlHoja1.Range("P1", "P2").MergeCells = True
                
                nLin = nLin + 1
                xlHoja1.Cells(nLin, 2) = "Natural"
                xlHoja1.Cells(nLin, 3) = "Jurídica"
                xlHoja1.Cells(nLin, 4) = "Soles"
                xlHoja1.Cells(nLin, 5) = "Dólares"
                xlHoja1.Cells(nLin, 6) = "Soles"
                xlHoja1.Cells(nLin, 7) = "Dólares"
                xlHoja1.Cells(nLin, 8) = "Soles"
                xlHoja1.Cells(nLin, 9) = "Dólares"
                xlHoja1.Cells(nLin, 11) = "Depósitos"
                xlHoja1.Cells(nLin, 12) = "Retiros"
                xlHoja1.Cells(nLin, 14) = "Soles"
                xlHoja1.Cells(nLin, 15) = "Dólares"
                
                xlHoja1.Range("A1:P2").Font.Bold = True
                xlHoja1.Range("A1:P2").HorizontalAlignment = xlHAlignCenter
                xlHoja1.Range("A1:P2").VerticalAlignment = xlVAlignCenter
                xlHoja1.Range("A1:P2").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
                xlHoja1.Range("A1:P2").Borders.LineStyle = xlContinuous
                xlHoja1.Range("A1:P2").Borders.Color = vbBlack
                xlHoja1.Range("A1:P2").Interior.Color = RGB(255, 60, 60)
                xlHoja1.Range("A1:P2").Font.Color = RGB(255, 255, 255)
                
                With xlHoja1.PageSetup
                    .LeftHeader = ""
                    .CenterHeader = ""
                    .RightHeader = ""
                    .LeftFooter = ""
                    .CenterFooter = ""
                    .RightFooter = ""
                
                    .PrintHeadings = False
                    .PrintGridlines = False
                    .PrintComments = xlPrintNoComments
                    .CenterHorizontally = True
                    .CenterVertically = False
                    .Draft = False
                    .FirstPageNumber = xlAutomatic
                    .Order = xlDownThenOver
                    .BlackAndWhite = False
                    .Zoom = 55
                End With
                
                nItem = 1
                nLin = nLin + 1
                For nItem = 1 To rs.RecordCount
                    xlHoja1.Range("A" & nLin & ":O" & nLin).HorizontalAlignment = xlHAlignLeft
                    xlHoja1.Cells(nLin, 1) = rs!cTpoPrograma
                    xlHoja1.Cells(nLin, 2) = IIf(rs!bPersNat, "Si", "No")
                    xlHoja1.Cells(nLin, 3) = IIf(rs!bPersJur, "Si", "No")
                    xlHoja1.Cells(nLin, 4) = IIf(rs!bMonSol, "Si", "No")
                    xlHoja1.Cells(nLin, 5) = IIf(rs!bMonDol, "Si", "No")
                    xlHoja1.Cells(nLin, 6) = "'" & Format(rs!nMontoMinApertSol, "#,##0.00")
                    xlHoja1.Cells(nLin, 7) = "'" & Format(rs!nMontoMinApertDol, "#,##0.00")
                    xlHoja1.Cells(nLin, 8) = "'" & Format(rs!nMontoMinDepSol, "#,##0.00")
                    xlHoja1.Cells(nLin, 9) = "'" & Format(rs!nMontoMinDepDol, "#,##0.00")
                    xlHoja1.Cells(nLin, 10) = rs!cOrdPago
                    xlHoja1.Cells(nLin, 11) = rs!nCantOpeVentDep
                    xlHoja1.Cells(nLin, 12) = rs!nCantOpeVentRet
                    xlHoja1.Cells(nLin, 13) = rs!nCantOpeCajero
                    xlHoja1.Cells(nLin, 14) = "'" & Format(rs!nSaldoMinCtaSol, "#,##0.00")
                    xlHoja1.Cells(nLin, 15) = "'" & Format(rs!nSaldoMinCtaDol, "#,##0.00")
                    xlHoja1.Cells(nLin, 16) = IIf(rs!bAplicaDesembCred, "Si", "No")
                    nLin = nLin + 1
                    rs.MoveNext
                Next nItem
            Case gCapPlazoFijo
                xlHoja1.Range("A1:A1").RowHeight = 30
                xlHoja1.Range("A2:A2").RowHeight = 18
                xlHoja1.Range("A1:A1").ColumnWidth = 18
                xlHoja1.Range("B1:O1").ColumnWidth = 8
                xlHoja1.Range("L1:L1").ColumnWidth = 11
                xlHoja1.Range("P1:P1").ColumnWidth = 11
                
                xlHoja1.Cells(nLin, 1) = "Producto"
                xlHoja1.Cells(nLin, 2) = "Personeria"
                xlHoja1.Cells(nLin, 4) = "Moneda"
                xlHoja1.Cells(nLin, 6) = "Monto mínimo" & Chr(10) & "apertura"
                xlHoja1.Cells(nLin, 8) = "Plazo en días"
                xlHoja1.Cells(nLin, 10) = "Medios de retiro intereses"
                xlHoja1.Cells(nLin, 13) = "Aumento" & Chr(10) & "de capital"
                xlHoja1.Cells(nLin, 14) = "Monto" & Chr(10) & "Aumento"
                xlHoja1.Cells(nLin, 16) = "Nº Max" & Chr(10) & "aumento" & Chr(10) & "capital (mes)"
                
                xlHoja1.Range("A1", "A2").MergeCells = True
                xlHoja1.Range("B1", "C1").MergeCells = True
                xlHoja1.Range("D1", "E1").MergeCells = True
                xlHoja1.Range("F1", "G1").MergeCells = True
                xlHoja1.Range("H1", "I1").MergeCells = True
                xlHoja1.Range("J1", "L1").MergeCells = True
                xlHoja1.Range("M1", "M2").MergeCells = True
                xlHoja1.Range("N1", "O2").MergeCells = True
                xlHoja1.Range("P1", "P2").MergeCells = True
                
                nLin = nLin + 1
                xlHoja1.Cells(nLin, 2) = "Natural"
                xlHoja1.Cells(nLin, 3) = "Jurídica"
                xlHoja1.Cells(nLin, 4) = "Soles"
                xlHoja1.Cells(nLin, 5) = "Dólares"
                xlHoja1.Cells(nLin, 6) = "Soles"
                xlHoja1.Cells(nLin, 7) = "Dólares"
                xlHoja1.Cells(nLin, 8) = "Mínimo"
                xlHoja1.Cells(nLin, 9) = "Máximo"
                xlHoja1.Cells(nLin, 10) = "Al Vencimiento"
                xlHoja1.Cells(nLin, 11) = "Mensual"
                xlHoja1.Cells(nLin, 12) = "Inicio"
                xlHoja1.Cells(nLin, 14) = "Soles"
                xlHoja1.Cells(nLin, 15) = "Dólares"
                
                xlHoja1.Range("A1:P2").Font.Bold = True
                xlHoja1.Range("A1:P2").HorizontalAlignment = xlHAlignCenter
                xlHoja1.Range("A1:P2").VerticalAlignment = xlVAlignCenter
                xlHoja1.Range("A1:P2").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
                xlHoja1.Range("A1:P2").Borders.LineStyle = xlContinuous
                xlHoja1.Range("A1:P2").Borders.Color = vbBlack
                xlHoja1.Range("A1:P2").Interior.Color = RGB(255, 60, 60)
                xlHoja1.Range("A1:P2").Font.Color = RGB(255, 255, 255)
                
                With xlHoja1.PageSetup
                    .LeftHeader = ""
                    .CenterHeader = ""
                    .RightHeader = ""
                    .LeftFooter = ""
                    .CenterFooter = ""
                    .RightFooter = ""
                
                    .PrintHeadings = False
                    .PrintGridlines = False
                    .PrintComments = xlPrintNoComments
                    .CenterHorizontally = True
                    .CenterVertically = False
                    .Draft = False
                    .FirstPageNumber = xlAutomatic
                    .Order = xlDownThenOver
                    .BlackAndWhite = False
                    .Zoom = 55
                End With
                
                nItem = 1
                nLin = nLin + 1
                For nItem = 1 To rs.RecordCount
                    xlHoja1.Range("A" & nLin & ":O" & nLin).HorizontalAlignment = xlHAlignLeft
                    xlHoja1.Cells(nLin, 1) = rs!cTpoPrograma
                    xlHoja1.Cells(nLin, 2) = IIf(rs!bPersNat, "Si", "No")
                    xlHoja1.Cells(nLin, 3) = IIf(rs!bPersJur, "Si", "No")
                    xlHoja1.Cells(nLin, 4) = IIf(rs!bMonSol, "Si", "No")
                    xlHoja1.Cells(nLin, 5) = IIf(rs!bMonDol, "Si", "No")
                    xlHoja1.Cells(nLin, 6) = "'" & Format(rs!nMontoMinApertSol, "#,##0.00")
                    xlHoja1.Cells(nLin, 7) = "'" & Format(rs!nMontoMinApertDol, "#,##0.00")
                    xlHoja1.Cells(nLin, 8) = rs!nPlazoMin
                    xlHoja1.Cells(nLin, 9) = rs!nPlazoMax
                    xlHoja1.Cells(nLin, 10) = IIf(rs!bFormaRetFinPlazo, "Si", "No")
                    xlHoja1.Cells(nLin, 11) = IIf(rs!bFormaRetMensual, "Si", "No")
                    xlHoja1.Cells(nLin, 12) = IIf(rs!bFormaRetInicioPlazo, "Si", "No")
                    xlHoja1.Cells(nLin, 13) = IIf(rs!bAumCap, "Si", "No")
                    xlHoja1.Cells(nLin, 14) = "'" & Format(rs!nAumCapMinSol, "#,##0.00")
                    xlHoja1.Cells(nLin, 15) = "'" & Format(rs!nAumCapMinDol, "#,##0.00")
                    xlHoja1.Cells(nLin, 16) = rs!nAumCapCantMaxMes
                    nLin = nLin + 1
                    rs.MoveNext
                Next nItem
            Case gCapCTS
                xlHoja1.Range("A1:A1").RowHeight = 25
                xlHoja1.Range("A1:A1").ColumnWidth = 18
                xlHoja1.Range("B1:D1").ColumnWidth = 15
                xlHoja1.Range("E1:G1").ColumnWidth = 19
                
                xlHoja1.Cells(nLin, 1) = "Producto"
                xlHoja1.Cells(nLin, 2) = "Nº Operaciones de" & Chr(10) & "Depósito x Año"
                xlHoja1.Cells(nLin, 3) = "Pocentaje Tangible" & Chr(10) & "de Retiro"
                xlHoja1.Cells(nLin, 4) = "Sueldo" & Chr(10) & "Mínimo S/."
                xlHoja1.Cells(nLin, 5) = "Factor últimas" & Chr(10) & "remuneraciones brutas"
                xlHoja1.Cells(nLin, 6) = "Meses sin depósio para" & Chr(10) & "traslado a CTS No Activo"
                xlHoja1.Cells(nLin, 7) = "Días verificación último" & Chr(10) & "registro sueldos CTS"
                
                xlHoja1.Range("A1:G1").Font.Bold = True
                xlHoja1.Range("A1:G1").HorizontalAlignment = xlHAlignCenter
                xlHoja1.Range("A1:G1").VerticalAlignment = xlVAlignCenter
                xlHoja1.Range("A1:G1").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
                xlHoja1.Range("A1:G1").Borders.LineStyle = xlContinuous
                xlHoja1.Range("A1:G1").Borders.Color = vbBlack
                xlHoja1.Range("A1:G1").Interior.Color = RGB(255, 60, 60)
                xlHoja1.Range("A1:G1").Font.Color = RGB(255, 255, 255)
                
                With xlHoja1.PageSetup
                    .LeftHeader = ""
                    .CenterHeader = ""
                    .RightHeader = ""
                    .LeftFooter = ""
                    .CenterFooter = ""
                    .RightFooter = ""
                
                    .PrintHeadings = False
                    .PrintGridlines = False
                    .PrintComments = xlPrintNoComments
                    .CenterHorizontally = True
                    .CenterVertically = False
                    .Draft = False
                    .FirstPageNumber = xlAutomatic
                    .Order = xlDownThenOver
                    .BlackAndWhite = False
                    .Zoom = 55
                End With
                
                nItem = 1
                nLin = nLin + 1
                For nItem = 1 To rs.RecordCount
                    xlHoja1.Range("A" & nLin & ":G" & nLin).HorizontalAlignment = xlHAlignLeft
                    xlHoja1.Cells(nLin, 1) = rs!cTpoPrograma
                    xlHoja1.Cells(nLin, 2) = rs!nCantOpeDepAnio
                    xlHoja1.Cells(nLin, 3) = "'" & Format(rs!nPorcTangRet, "#,##0.00")
                    xlHoja1.Cells(nLin, 4) = "'" & Format(rs!nSueldoMin, "#,##0.00")
                    xlHoja1.Cells(nLin, 5) = rs!nUltRemunBrutas
                    xlHoja1.Cells(nLin, 6) = rs!nMesesSinDepTransNoActivo
                    xlHoja1.Cells(nLin, 7) = rs!nDiasVerifUltRegSueldo
                    nLin = nLin + 1
                    rs.MoveNext
                Next nItem
        End Select
        
        gFunGeneral.ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
        gFunGeneral.CargaArchivo lsArchivo, App.path & "\SPOOLER\"
    Else
        MsgBox "No se encontraron datos", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdGrabar_Click()
Dim MatParametros() As String
Dim nTpoProg As Integer
Dim lsMovNro As String
Dim bActualiza As Boolean
Dim bActRemunBrutas As Boolean

If Not ValidaDatos Then Exit Sub

Set clsCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
Select Case nProd
    Case gCapAhorros
        nTpoProg = CInt(Trim(Right(cboAhoTpoPrograma.Text, 2)))
        ReDim MatParametros(14)
        MatParametros(0) = chkAhoPersNat.value
        MatParametros(1) = chkAhoPersJur.value
        MatParametros(2) = chkAhoSol.value
        MatParametros(3) = chkAhoDol.value
        MatParametros(4) = CDbl(IIf(txtAhoMontoMinApertSol.Text = "", 0, txtAhoMontoMinApertSol.Text))
        MatParametros(5) = CDbl(IIf(txtAhoMontoMinApertDol.Text = "", 0, txtAhoMontoMinApertDol.Text))
        MatParametros(6) = CDbl(IIf(txtAhoMontoMinDepSol.Text = "", 0, txtAhoMontoMinDepSol.Text))
        MatParametros(7) = CDbl(IIf(txtAhoMontoMinDepDol.Text = "", 0, txtAhoMontoMinDepDol.Text))
        MatParametros(8) = IIf(Trim(Right(cboAhoOrdPag.Text, 2)) = "", 0, Trim(Right(cboAhoOrdPag.Text, 2)))
        MatParametros(9) = CLng(IIf(txtAhoCantOpeVentDep.Text = "", 0, txtAhoCantOpeVentDep.Text))
        MatParametros(10) = CLng(IIf(txtAhoCantOpeVentRet.Text = "", 0, txtAhoCantOpeVentRet.Text))
        MatParametros(11) = CLng(IIf(txtAhoCantOpeCajeros.Text = "", 0, txtAhoCantOpeCajeros.Text))
        MatParametros(12) = CDbl(IIf(txtAhoSaldoMinSol.Text = "", 0, txtAhoSaldoMinSol.Text))
        MatParametros(13) = CDbl(IIf(txtAhoSaldoMinDol.Text = "", 0, txtAhoSaldoMinDol.Text))
        MatParametros(14) = IIf(Trim(Right(cboAhoAplicaDesemb.Text, 2)) = "", 0, Trim(Right(cboAhoAplicaDesemb.Text, 2)))
    Case gCapPlazoFijo
        nTpoProg = CInt(Trim(Right(cboPFTpoPrograma.Text, 2)))
        ReDim MatParametros(14)
        MatParametros(0) = chkPFPersNat.value
        MatParametros(1) = chkPFPersJur.value
        MatParametros(2) = chkPFSol.value
        MatParametros(3) = chkPFDol.value
        MatParametros(4) = CDbl(IIf(txtPFMontoMinApertSol.Text = "", 0, txtPFMontoMinApertSol.Text))
        MatParametros(5) = CDbl(IIf(txtPFMontoMinApertDol.Text = "", 0, txtPFMontoMinApertDol.Text))
        MatParametros(6) = CLng(IIf(txtPFPlazoMin.Text = "", 0, txtPFPlazoMin.Text))
        MatParametros(7) = CLng(IIf(txtPFPlazoMax.Text = "", 0, txtPFPlazoMax.Text))
        MatParametros(8) = chkPFMedRetFinPlazo.value
        MatParametros(9) = chkPFMedRetIniPlazo.value
        MatParametros(10) = chkPFMedRetMensual.value
        MatParametros(11) = lbAumCap
        MatParametros(12) = CDbl(IIf(txtPFAumCapSol.Text = "", 0, txtPFAumCapSol.Text))
        MatParametros(13) = CDbl(IIf(txtPFAumCapDol.Text = "", 0, txtPFAumCapDol.Text))
        MatParametros(14) = CLng(IIf(txtPFMaxAumCapMes.Text = "", 0, txtPFMaxAumCapMes.Text))
    Case gCapCTS
        nTpoProg = CInt(Trim(Right(cboCTSTpoPrograma.Text, 2)))
        If lnPorcTangRet = 0 Then lnPorcTangRet = clsCapDef.GetCapParametro(2021)
        If lnUltRemunBrutasAnt = 0 Then lnUltRemunBrutasAnt = clsCapDef.GetCapParametro(2129)
        If IIf(txtCTSUltRemunBrutas.Text = "", 0, txtCTSUltRemunBrutas.Text) = "0" Then
            MsgBox "El factor ùltimas remuneraciones brutas debe ser mayor de 0", vbInformation, "Aviso"
            txtCTSUltRemunBrutas.SetFocus
            Exit Sub
        End If
        ReDim MatParametros(5)
        MatParametros(0) = CLng(IIf(txtCTSCantOpeDepAnio.Text = "", 0, txtCTSCantOpeDepAnio.Text))
        MatParametros(1) = CDbl(IIf(txtCTSPorcTangRet.Text = "", 0, txtCTSPorcTangRet.Text))
        MatParametros(2) = CDbl(IIf(txtCTSSueldoMin.Text = "", 0, txtCTSSueldoMin.Text))
        MatParametros(3) = CLng(IIf(txtCTSUltRemunBrutas.Text = "", 0, txtCTSUltRemunBrutas.Text))
        MatParametros(4) = CLng(IIf(txtCTSMesesSinDepTransNoAct.Text = "", 0, txtCTSMesesSinDepTransNoAct.Text))
        MatParametros(5) = CLng(IIf(txtCTSDiasVerifUltRegSueldo.Text = "", 0, txtCTSDiasVerifUltRegSueldo.Text))
End Select

If MsgBox("Se guardarán los parámetros, ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub

If nProd = gCapCTS Then
    If lnUltRemunBrutasAnt <> CInt(IIf(txtCTSUltRemunBrutas.Text = "", 0, txtCTSUltRemunBrutas.Text)) Or _
       lnPorcTangRet <> CDbl(IIf(txtCTSPorcTangRet.Text = "", 0, txtCTSPorcTangRet.Text)) Then
        If MsgBox("Se ha cambiado la configuración del Factor de última remuneraciones brutas o el Pocetaje Tangible de Retiro." & Chr(10) & _
                  "Se realizará la actualización del registro de los sueldos y se actualizará el saldo disponible de retiro, " & _
                  "¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            bActRemunBrutas = True
        Else
            Exit Sub
        End If
    End If
End If

Dim clsCont As COMNContabilidad.NCOMContFunciones
Set clsCont = New COMNContabilidad.NCOMContFunciones
lsMovNro = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
Set clsCont = Nothing

bActualiza = clsCapDef.ActualizaCapParametroNew(nProd, nTpoProg, lsMovNro, MatParametros())
Set clsCapDef = Nothing
If bActualiza Then
    If bActRemunBrutas Then
        Dim oContFun As COMNContabilidad.NCOMContFunciones
        Dim oCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento
        Dim sMovNro As String
        
        Set oContFun = New COMNContabilidad.NCOMContFunciones
            sMovNro = oContFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set oContFun = Nothing
        
        Set oCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
            Call oCapMov.ActualizaSueldosCTS(nTpoProg, lnUltRemunBrutasAnt, CInt(IIf(txtCTSUltRemunBrutas.Text = "", 0, txtCTSUltRemunBrutas.Text)), sMovNro)
        Set oCapMov = Nothing
    End If
    MsgBox "Los parámetros fueron actualizados", vbInformation, "Aviso"
    Call cmdCancelar_Click
End If
End Sub

Private Sub cmdPFSeleccionar_Click()
If Trim(Right(cboPFTpoPrograma.Text, 2)) = "" Then Exit Sub
CargarParametroSubProducto
If nTipoOpe = 1 Then
    cboPFTpoPrograma.Enabled = False
    cmdPFSeleccionar.Enabled = False
    cmdEditar.Enabled = True
    cmdCancelar.Enabled = True
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub optAumCap_Click(Index As Integer)
Select Case Index
    Case 0
        txtPFAumCapSol.Enabled = False
        txtPFAumCapDol.Enabled = False
        txtPFMaxAumCapMes.Enabled = False
    Case 1
        txtPFAumCapSol.Enabled = IIf(chkPFSol.value, True, False)
        txtPFAumCapDol.Enabled = IIf(chkPFDol.value, True, False)
        txtPFMaxAumCapMes.Enabled = True
End Select
lbAumCap = Index
End Sub

Private Sub txtAhoCantOpeCajeros_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        If Me.chkAhoSol.value = 1 Then
            txtAhoSaldoMinSol.SetFocus
        Else
            txtAhoSaldoMinDol.SetFocus
        End If
    End If
End Sub

Private Sub txtAhoCantOpeVentDep_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        txtAhoCantOpeVentRet.SetFocus
    End If
End Sub

Private Sub txtAhoCantOpeVentRet_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        txtAhoCantOpeCajeros.SetFocus
    End If
End Sub

Private Sub txtAhoMontoMinApertDol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkAhoSol.value Then
            txtAhoMontoMinDepSol.SetFocus
        Else
            txtAhoMontoMinDepDol.SetFocus
        End If
    End If
End Sub

Private Sub txtAhoMontoMinApertSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkAhoDol.value Then
            txtAhoMontoMinApertDol.SetFocus
        Else
            txtAhoMontoMinDepSol.SetFocus
        End If
    End If
End Sub

Private Sub txtAhoMontoMinDepDol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboAhoOrdPag.SetFocus
    End If
End Sub

Private Sub txtAhoMontoMinDepSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkAhoDol.value Then
            txtAhoMontoMinDepDol.SetFocus
        Else
            cboAhoOrdPag.SetFocus
        End If
    End If
End Sub

Private Sub cboAhoOrdPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAhoCantOpeVentDep.SetFocus
    End If
End Sub

Private Sub txtAhoSaldoMinDol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboAhoAplicaDesemb.SetFocus
    End If
End Sub

Private Sub txtAhoSaldoMinSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkAhoDol.value Then
            txtAhoSaldoMinDol.SetFocus
        Else
            cboAhoAplicaDesemb.SetFocus
        End If
    End If
End Sub

Private Sub cboAhoAplicaDesemb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub

Private Sub txtCTSCantOpeDepAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        txtCTSPorcTangRet.SetFocus
    End If
End Sub

Private Sub txtCTSDiasVerifUltRegSueldo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub

Private Sub txtCTSMesesSinDepTransNoAct_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCTSDiasVerifUltRegSueldo.SetFocus
    End If
End Sub

Private Sub txtCTSPorcTangRet_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii, True)
    If KeyAscii = 13 Then
        txtCTSSueldoMin.SetFocus
    End If
End Sub

Private Sub txtCTSPorcTangRet_KeyUp(KeyCode As Integer, Shift As Integer)
Dim nCantDecimales As Integer
    If txtCTSPorcTangRet.Text <> "" Then
        If CInt(txtCTSPorcTangRet.Text) > 100 Then
            MsgBox "El porcentaje tangible de retiro no debe ser superior a 100.00", vbInformation, "Aviso"
            txtCTSPorcTangRet.Text = "100.00"
        End If
        If InStr(1, txtCTSPorcTangRet.Text, ".") <> 0 Then
            nCantDecimales = Len(Mid(txtCTSPorcTangRet.Text, InStr(1, txtCTSPorcTangRet.Text, ".") + 1, Len(txtCTSPorcTangRet.Text)))
            If nCantDecimales > 4 Then
                MsgBox "El porcentaje solo permite 4 decimales", vbInformation, "Aviso"
                txtCTSPorcTangRet.Text = Mid(txtCTSPorcTangRet.Text, 1, Len(txtCTSPorcTangRet.Text) - (nCantDecimales - 1))
            End If
        End If
    Else
        txtCTSPorcTangRet.Text = "0"
    End If
End Sub

Private Sub txtCTSSueldoMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCTSUltRemunBrutas.SetFocus
    End If
End Sub

Private Sub txtCTSUltRemunBrutas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCTSMesesSinDepTransNoAct.SetFocus
    End If
End Sub

Private Sub txtPFAumCapDol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPFMaxAumCapMes.SetFocus
    End If
End Sub

Private Sub txtPFAumCapSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.chkPFDol.value = 1 Then
            txtPFAumCapDol.SetFocus
        Else
            txtPFMaxAumCapMes.SetFocus
        End If
    End If
End Sub

Private Sub txtPFMaxAumCapMes_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub

Private Sub txtPFMontoMinApertDol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPFPlazoMin.SetFocus
    End If
End Sub

Private Sub txtPFMontoMinApertSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.chkPFDol.value = 1 Then
            txtPFMontoMinApertDol.SetFocus
        Else
            txtPFPlazoMin.SetFocus
        End If
    End If
End Sub

Private Sub txtPFPlazoMax_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        chkPFMedRetFinPlazo.SetFocus
    End If
End Sub

Private Sub txtPFPlazoMin_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        txtPFPlazoMax.SetFocus
    End If
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False

Select Case nProd
    Case gCapAhorros
        If chkAhoPersNat.value = 0 And chkAhoPersJur.value = 0 Then
            MsgBox "Debe seleccionar al menos una de las personerias", vbInformation, "Aviso"
            Exit Function
        End If
        If chkAhoSol.value = 0 And chkAhoDol.value = 0 Then
            MsgBox "Debe seleccionar al menos una de las monedas", vbInformation, "Aviso"
            Exit Function
        End If
    Case gCapPlazoFijo
        If chkPFPersNat.value = 0 And chkPFPersJur.value = 0 Then
            MsgBox "Debe seleccionar al menos una de las personerias", vbInformation, "Aviso"
            Exit Function
        End If
        If chkPFSol.value = 0 And chkPFDol.value = 0 Then
            MsgBox "Debe seleccionar al menos una de las monedas", vbInformation, "Aviso"
            Exit Function
        End If
        If CLng(IIf(Trim(txtPFPlazoMax) = "", 0, txtPFPlazoMax)) < CLng(IIf(Trim(txtPFPlazoMin) = "", 0, txtPFPlazoMin)) Then
            MsgBox "El plazo maximo no vebe ser menor o igual al plazo mínimo", vbInformation, "Aviso"
            Exit Function
        End If
End Select

ValidaDatos = True
End Function
