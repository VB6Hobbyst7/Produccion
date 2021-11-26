VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCajeroIngEgre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresos y Egresos de Cajero"
   ClientHeight    =   8145
   ClientLeft      =   1185
   ClientTop       =   2085
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajeroIngEgre.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImpBill 
      Caption         =   "Imprimir Billetaje"
      Height          =   375
      Left            =   240
      TabIndex        =   73
      Top             =   7560
      Width           =   1500
   End
   Begin VB.CommandButton cmdRepoArqueo 
      BackColor       =   &H80000004&
      Caption         =   "Reporte para Arqueo"
      Height          =   360
      Left            =   9240
      TabIndex        =   72
      Top             =   7530
      Width           =   2625
   End
   Begin MSComctlLib.ProgressBar Pg 
      Height          =   255
      Left            =   4620
      TabIndex        =   67
      Top             =   7230
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer TDesatendido 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   4980
      Top             =   7215
   End
   Begin VB.CommandButton cmdPreCuadre 
      BackColor       =   &H80000000&
      Caption         =   "Pre-Cuadre de Agencias"
      Height          =   360
      Left            =   2445
      TabIndex        =   66
      Top             =   7170
      Width           =   2100
   End
   Begin SICMACT.Usuario Usuario 
      Left            =   3390
      Top             =   7185
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   10560
      TabIndex        =   65
      Top             =   7170
      Width           =   1305
   End
   Begin VB.CommandButton cmdRegSobFalt 
      Caption         =   "&Registrar Sobrante Faltante"
      Height          =   360
      Left            =   180
      TabIndex        =   64
      Top             =   7170
      Width           =   2280
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H80000004&
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   9270
      TabIndex        =   63
      Top             =   7170
      Width           =   1305
   End
   Begin TabDlg.SSTab tabIngEgre 
      Height          =   6435
      Left            =   75
      TabIndex        =   7
      Top             =   630
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   11351
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      TabCaption(0)   =   "Moneda &Nacional"
      TabPicture(0)   =   "frmCajeroIngEgre.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fgPlanillaMN"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Moneda &Extranjera"
      TabPicture(1)   =   "frmCajeroIngEgre.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgPlanillaME"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Resumen"
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
         Height          =   1920
         Left            =   -74910
         TabIndex        =   34
         Top             =   4380
         Width           =   11565
         Begin VB.Label lblNroMovME 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   7605
            TabIndex        =   71
            Top             =   1020
            Width           =   1395
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Total Mov."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   6405
            TabIndex        =   70
            Top             =   1110
            Width           =   1170
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ingresos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   240
            TabIndex        =   58
            Top             =   345
            Width           =   750
         End
         Begin VB.Label lblTotalIngME 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   1410
            TabIndex        =   57
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Habilitaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   240
            TabIndex        =   56
            Top             =   690
            Width           =   1140
         End
         Begin VB.Label lblTotalHabME 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   1410
            TabIndex        =   55
            Top             =   645
            Width           =   1305
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Egresos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   2880
            TabIndex        =   54
            Top             =   345
            Width           =   690
         End
         Begin VB.Label lblTotalEgreME 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   4020
            TabIndex        =   53
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Devoluciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   2880
            TabIndex        =   52
            Top             =   690
            Width           =   1110
         End
         Begin VB.Label lblTotalDevME 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   4020
            TabIndex        =   51
            Top             =   645
            Width           =   1305
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dev. Billetaje:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   2880
            TabIndex        =   50
            Top             =   1035
            Width           =   1080
         End
         Begin VB.Label lblTotalDevBillME 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   4020
            TabIndex        =   49
            Top             =   990
            Width           =   1305
         End
         Begin VB.Label lblSobFaltMEEtq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sob/Faltante :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   240
            TabIndex        =   48
            Top             =   1035
            Width           =   1095
         End
         Begin VB.Label lblTotalSobFalME 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   1410
            TabIndex        =   47
            Top             =   990
            Width           =   1305
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   2910
            TabIndex        =   46
            Top             =   1470
            Width           =   645
         End
         Begin VB.Label lblTotalME 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   3675
            TabIndex        =   45
            Top             =   1425
            Width           =   1620
         End
         Begin VB.Line Line2 
            X1              =   5865
            X2              =   5865
            Y1              =   240
            Y2              =   1800
         End
         Begin VB.Label lblSaldoAntME 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   7605
            TabIndex        =   44
            Top             =   315
            Width           =   1395
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Anterior"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   6405
            TabIndex        =   43
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label lblNetoMovME 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   7605
            TabIndex        =   42
            Top             =   675
            Width           =   1395
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Neto Movim."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   6405
            TabIndex        =   41
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Final"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   9075
            TabIndex        =   40
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Caja"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   9090
            TabIndex        =   39
            Top             =   720
            Width           =   795
         End
         Begin VB.Label lblSaldoFinalME 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   9975
            TabIndex        =   38
            Top             =   315
            Width           =   1320
         End
         Begin VB.Label lblTotalCajaME 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   9975
            TabIndex        =   37
            Top             =   675
            Width           =   1320
         End
         Begin VB.Label lblDifME 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   9930
            TabIndex        =   36
            Top             =   1365
            Width           =   1305
         End
         Begin VB.Label lblDifMEEtq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diferencia :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   8910
            TabIndex        =   35
            Top             =   1410
            Width           =   915
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            Height          =   1590
            Left            =   45
            Top             =   225
            Width           =   11460
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000B&
         Caption         =   "Resumen"
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
         Height          =   1920
         Left            =   90
         TabIndex        =   9
         Top             =   4380
         Width           =   11565
         Begin VB.Label lblNroMovMN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   7605
            TabIndex        =   69
            Top             =   1020
            Width           =   1395
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Total Mov."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   6405
            TabIndex        =   68
            Top             =   1110
            Width           =   1170
         End
         Begin VB.Label lblDifMNEtq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diferencia :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   8910
            TabIndex        =   33
            Top             =   1410
            Width           =   915
         End
         Begin VB.Label lblDifMN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   9930
            TabIndex        =   32
            Top             =   1365
            Width           =   1305
         End
         Begin VB.Label lblTotalCajaMN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   9975
            TabIndex        =   31
            Top             =   675
            Width           =   1320
         End
         Begin VB.Label lblSaldoFinalMN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   9975
            TabIndex        =   30
            Top             =   315
            Width           =   1320
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Caja"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   9090
            TabIndex        =   29
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Final"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   9075
            TabIndex        =   28
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Neto Movim."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   6405
            TabIndex        =   27
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label lblNetoMovMN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   7605
            TabIndex        =   26
            Top             =   675
            Width           =   1395
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Anterior"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   6405
            TabIndex        =   25
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label lblSaldoAntMN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   7605
            TabIndex        =   24
            Top             =   315
            Width           =   1395
         End
         Begin VB.Line Line1 
            X1              =   5865
            X2              =   5865
            Y1              =   240
            Y2              =   1800
         End
         Begin VB.Label lblTotalMN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   3675
            TabIndex        =   23
            Top             =   1425
            Width           =   1620
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   2910
            TabIndex        =   22
            Top             =   1470
            Width           =   645
         End
         Begin VB.Label lblTotalSobFalMN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   1410
            TabIndex        =   21
            Top             =   990
            Width           =   1305
         End
         Begin VB.Label lblSobFaltMNEtq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sob/Faltante :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   240
            TabIndex        =   20
            Top             =   1035
            Width           =   1095
         End
         Begin VB.Label lblTotalDevBillMN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   4020
            TabIndex        =   19
            Top             =   990
            Width           =   1305
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dev. Billetaje:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   2880
            TabIndex        =   18
            Top             =   1035
            Width           =   1080
         End
         Begin VB.Label lblTotalDevMN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   4020
            TabIndex        =   17
            Top             =   645
            Width           =   1305
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Devoluciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   2880
            TabIndex        =   16
            Top             =   690
            Width           =   1110
         End
         Begin VB.Label lblTotalEgreMN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   4020
            TabIndex        =   15
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Egresos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   2880
            TabIndex        =   14
            Top             =   345
            Width           =   690
         End
         Begin VB.Label lblTotalHabMN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   1410
            TabIndex        =   13
            Top             =   645
            Width           =   1305
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Habilitaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   240
            TabIndex        =   12
            Top             =   690
            Width           =   1140
         End
         Begin VB.Label lblTotalIngMN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   1410
            TabIndex        =   11
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ingresos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   240
            TabIndex        =   10
            Top             =   345
            Width           =   750
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            Height          =   1590
            Left            =   45
            Top             =   225
            Width           =   11460
         End
      End
      Begin SICMACT.FlexEdit fgPlanillaMN 
         Height          =   3945
         Left            =   90
         TabIndex        =   8
         Top             =   450
         Width           =   11565
         _ExtentX        =   20399
         _ExtentY        =   6959
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Descripción-N° Mov-Efectivo-Cheque-Orden Pago-No Efectivo-Total-cGruCod"
         EncabezadosAnchos=   "350-4000-800-1200-1200-1200-1200-1200-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColor       =   16777215
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-R-R-R-R-C"
         FormatosEdit    =   "0-0-3-2-2-2-2-2-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
         CellBackColor   =   16777215
      End
      Begin SICMACT.FlexEdit fgPlanillaME 
         Height          =   3945
         Left            =   -74910
         TabIndex        =   59
         Top             =   450
         Width           =   11565
         _ExtentX        =   20399
         _ExtentY        =   6959
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Descripción-N° Mov-Efectivo-Cheque-Orden Pago-No Efectivo-Total-cGrupoCod"
         EncabezadosAnchos=   "350-4000-800-1200-1200-1200-1200-1200-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-R-R-R-R-C"
         FormatosEdit    =   "0-0-3-2-2-2-2-2-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin MSComctlLib.StatusBar barra 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   7905
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame frageneral 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   97
      TabIndex        =   0
      Top             =   -15
      Width           =   11760
      Begin MSComCtl2.DTPicker txtFecha 
         Height          =   330
         Left            =   720
         TabIndex        =   6
         Top             =   165
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         Format          =   87752705
         CurrentDate     =   37090
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10440
         TabIndex        =   5
         Top             =   180
         Width           =   1230
      End
      Begin SICMACT.TxtBuscar TxtBuscarUser 
         Height          =   345
         Left            =   7935
         TabIndex        =   1
         Top             =   158
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
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
         ForeColor       =   12582912
      End
      Begin SICMACT.TxtBuscar TxtBuscarAge 
         Height          =   345
         Left            =   2790
         TabIndex        =   60
         Top             =   158
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
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
         ForeColor       =   12582912
      End
      Begin VB.Label lblDescAge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   3735
         TabIndex        =   62
         Top             =   165
         Width           =   3570
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1980
         TabIndex        =   61
         Top             =   225
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   7440
         TabIndex        =   4
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   225
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmCajeroIngEgre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oGen As COMDConstSistema.DCOMGeneral  'DGeneral

Dim bPlanillaBoveda As Boolean
Dim nIngresos As Double, nEgresos As Double
Dim nHabilitacion As Double
Dim nDevolucion As Double, nDevolucionBilletaje As Double
Dim nSobrante As Double, nFaltante As Double

'**DAOR 20080125 ******************************
Dim bExisteRegSobrFalt As Boolean, bProcesado As Boolean, bRegistrado As Boolean
Dim fsMovNroRegEfec As String
Dim fnMovNroRegEfec As Long, fnNumRegEfec As Integer
'**********************************************
'**MADM 20101006 *************************************
Dim loVistoElectronico As frmVistoElectronico
Dim lbVistoVal As Boolean
'*****************************************************
'**MADM 20110118 *************************************
Dim fsOpeCod As Long
'*****************************************************
'**EJVG 20110809 *************************************
Dim fnNumOpeMN As Integer
Dim fnNumOpeME As Integer
'*****************************************************
'*** RIRO 20130702 ***********************************
Dim rsEgresos As ADODB.Recordset
'*****************************************************
Private sGrupoOpe As String 'RIRO20150608 ERS162-2014

Private Sub ProcesaGrid(ByRef grid As FlexEdit, ByVal bBoveda As Boolean, ByVal bConsolidada As Boolean)
    Dim i As Long
    Dim nEfectivo As Double
    Dim nMontoEfect As Double, nMontoOrdPag As Double
    Dim sCodigoGrupo As String

    nIngresos = 0
    nEgresos = 0
    nHabilitacion = 0
    nDevolucion = 0
    nDevolucionBilletaje = 0
    nFaltante = 0
    nSobrante = 0
    For i = 1 To grid.Rows - 1
        'sCodigoGrupo = grid.TextMatrix(i, 7)
        sCodigoGrupo = grid.TextMatrix(i, 8) 'FRHU 20150311 OBSERVACION
        nMontoEfect = 0
        nMontoOrdPag = 0
        If grid.TextMatrix(i, 3) <> "" Then
            nMontoEfect = CDbl(grid.TextMatrix(i, 3))
        End If
        If grid.TextMatrix(i, 5) <> "" Then
            nMontoOrdPag = CDbl(grid.TextMatrix(i, 5))
        End If
        nEfectivo = nMontoEfect + nMontoOrdPag
        If Abs(nEfectivo) > 0 Then
            If bConsolidada Then
                If sCodigoGrupo = gsCajHabCajGen Then
                    nHabilitacion = nEfectivo
                ElseIf sCodigoGrupo = gsCajDevCajGen Then
                    nDevolucion = nEfectivo
                ElseIf sCodigoGrupo = gsCajDevBilletaje Then
                    nDevolucionBilletaje = nEfectivo
                ElseIf sCodigoGrupo = gsCajFaltCaja Then
                    nFaltante = nEfectivo
                ElseIf sCodigoGrupo = gsCajSobCaja Then
                    nSobrante = nEfectivo
                Else
                    If nEfectivo >= 0 Then
                        nIngresos = nIngresos + nEfectivo
                    Else
                        nEgresos = nEgresos + nEfectivo
                    End If
                End If
            Else
                If sCodigoGrupo = gsCajHabDeBove Then
                    nHabilitacion = nEfectivo
                ElseIf sCodigoGrupo = gsCajDevABove Then
                    nDevolucion = nEfectivo
                ElseIf sCodigoGrupo = gsCajDevBilletaje Then
                    nDevolucionBilletaje = nEfectivo
                ElseIf sCodigoGrupo = gsCajFaltCaja Then
                    nFaltante = nEfectivo
                ElseIf sCodigoGrupo = gsCajSobCaja Then
                    nSobrante = nEfectivo
                Else
                    If nEfectivo >= 0 Then
                        nIngresos = nIngresos + nEfectivo
                    Else
                        nEgresos = nEgresos + nEfectivo
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Sub CalculaDatosPlanilla(Optional bBoveda As Boolean = False, Optional bConsolidada As Boolean = False)
    Dim sCodigoGrupo As String
    Dim oEfectivo As COMDCajaGeneral.DCOMEfectivo  'Defectivo
    Dim nSaldoMovimientos As Double, nBilletaje As Double
    Dim nSaldoEfectivoAyer As Double, nSaldoEfectivoHoy As Double
    Dim nSobFalt As Double
    Dim sMensajeMN As String, sMensajeME As String

    'Moneda Nacional
    ProcesaGrid fgPlanillaMN, bBoveda, bConsolidada

    Dim COnSOb As Double

    lblTotalIngMN = Format$(nIngresos, "#,##0.00")
    lblTotalDevBillMN = Format$(nDevolucionBilletaje, "#,##0.00")
    lblTotalEgreMN = Format$(nEgresos, "#,##0.00")
    lblTotalHabMN = Format$(nHabilitacion, "#,##0.00")
    lblTotalDevMN = Format$(nDevolucion, "#,##0.00")
    lblTotalSobFalMN = Format$(nSobrante + nFaltante, "#,##0.00")
    Me.lblNroMovMN.Caption = fgPlanillaMN.SumaRow(2)
    nSaldoMovimientos = nIngresos + nEgresos + nHabilitacion + nDevolucion + nSobrante + nFaltante

    Set oEfectivo = New COMDCajaGeneral.DCOMEfectivo
    If bConsolidada Then
        nSaldoEfectivoAyer = oEfectivo.GetSaldoEfectivo(TxtBuscarAge, DateAdd("d", -1, CDate(txtFecha)), gMonedaNacional, gdFecSis, gsCodUser, , True)
        nSaldoEfectivoHoy = oEfectivo.GetSaldoEfectivo(TxtBuscarAge, CDate(txtFecha), gMonedaNacional, gdFecSis, gsCodUser)
        lblDifMN = Format$(nSaldoEfectivoAyer + nSaldoMovimientos - nSaldoEfectivoHoy, "#,##0.00")
        lblSaldoAntMN = Format$(nSaldoEfectivoAyer, "#,##0.00")
        lblTotalCajaMN = Format$(nSaldoEfectivoHoy, "#,##0.00")
        lblSaldoFinalMN = Format$(nSaldoEfectivoAyer + nSaldoMovimientos, "#,##0.00")
        lblNetoMovMN = Format$(nSaldoMovimientos, "#,##0.00")
        Set oEfectivo = Nothing
    Else
        If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
            nSaldoEfectivoHoy = oEfectivo.GetSaldoEfectivoPreCuadre(TxtBuscarAge, CDate(txtFecha), gMonedaNacional, gdFecSis, gsCodUser, txtBuscarUser.Text)
        Else
            nSaldoEfectivoHoy = oEfectivo.GetSaldoEfectivo(TxtBuscarAge, CDate(txtFecha), gMonedaNacional, gdFecSis, gsCodUser, txtBuscarUser.Text)
        End If
    
        If bBoveda Then
            nSaldoEfectivoAyer = oEfectivo.GetSaldoEfectivo(TxtBuscarAge, DateAdd("d", -1, CDate(txtFecha)), gMonedaNacional, gdFecSis, gsCodUser, , True)
            nSobFalt = Round(nSaldoEfectivoHoy - (nSaldoEfectivoAyer + nSaldoMovimientos), 2)
            lblSaldoFinalMN = Format$(nSaldoEfectivoAyer + nSaldoMovimientos, "#,##0.00")
        Else
            nSaldoEfectivoAyer = 0
            nSobFalt = Round(nSaldoEfectivoHoy - nSaldoMovimientos, 2)
            lblSaldoFinalMN = Format$(nSaldoMovimientos, "#,##0.00")
        End If
    
        lblSaldoAntMN = Format$(nSaldoEfectivoAyer, "#,##0.00")
        lblTotalCajaMN = Format$(nSaldoEfectivoHoy, "#,##0.00")
    
        lblNetoMovMN = Format$(nSaldoMovimientos, "#,##0.00")
    
        If nSobFalt > 0 Then
            lblDifMN = Format$(nSobFalt, "#,##0.00")
            lblDifMNEtq.Caption = "SOBRANTE !!"
            lblDifMN.ForeColor = &H800000
            lblDifMNEtq.ForeColor = &H800000
            COnSOb = nSobFalt
        ElseIf nSobFalt < 0 Then
            lblDifMN = Format$(nSobFalt, "#,##0.00")
            lblDifMNEtq.Caption = "FALTANTE !!"
            lblDifMN.ForeColor = &HC0&
            lblDifMNEtq.ForeColor = &HC0&
            COnSOb = (nSobFalt) * -1
        Else
            lblDifMN = "SIN SOB/FALT"
            lblDifMNEtq.Caption = "Sob/Faltante"
            lblDifMN.ForeColor = &H800000
            lblDifMNEtq.ForeColor = &H800000
        End If
    
    End If
    lblTotalMN = Format$(nSaldoMovimientos + nDevolucionBilletaje, "#,##0.00")

    'Moneda Extranjera
    ProcesaGrid fgPlanillaME, bBoveda, bConsolidada
    lblTotalIngME = Format$(nIngresos, "#,##0.00")
    lblTotalDevBillME = Format$(nDevolucionBilletaje, "#,##0.00")
    lblTotalEgreME = Format$(nEgresos, "#,##0.00")
    lblTotalHabME = Format$(nHabilitacion, "#,##0.00")
    lblTotalDevME = Format$(nDevolucion, "#,##0.00")
    lblTotalSobFalME = Format$(nSobrante + nFaltante, "#,##0.00")
    Me.lblNroMovME.Caption = fgPlanillaME.SumaRow(2)
    nSaldoMovimientos = nIngresos + nEgresos + nHabilitacion + nDevolucion + nSobrante + nFaltante

    If bConsolidada Then
        Set oEfectivo = New COMDCajaGeneral.DCOMEfectivo
        nSaldoEfectivoAyer = oEfectivo.GetSaldoEfectivo(TxtBuscarAge, DateAdd("d", -1, CDate(txtFecha)), gMonedaExtranjera, gdFecSis, gsCodUser, , True)
        nSaldoEfectivoHoy = oEfectivo.GetSaldoEfectivo(TxtBuscarAge, CDate(txtFecha), gMonedaExtranjera, gdFecSis, gsCodUser)
        lblDifME = Format$(nSaldoEfectivoAyer + nSaldoMovimientos - nSaldoEfectivoHoy, "#,##0.00")
        lblSaldoAntME = Format$(nSaldoEfectivoAyer, "#,##0.00")
        lblTotalCajaME = Format$(nSaldoEfectivoHoy, "#,##0.00")
        lblSaldoFinalME = Format$(nSaldoEfectivoAyer + nSaldoMovimientos, "#,##0.00")
        lblNetoMovME = Format$(nSaldoMovimientos, "#,##0.00")
        Set oEfectivo = Nothing
    Else
        If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
            nSaldoEfectivoHoy = oEfectivo.GetSaldoEfectivoPreCuadre(TxtBuscarAge, CDate(txtFecha), gMonedaExtranjera, gdFecSis, gsCodUser, txtBuscarUser.Text)
        Else
            nSaldoEfectivoHoy = oEfectivo.GetSaldoEfectivo(TxtBuscarAge, CDate(txtFecha), gMonedaExtranjera, gdFecSis, gsCodUser, txtBuscarUser.Text)
        End If
        
        If bBoveda Then
            nSaldoEfectivoAyer = oEfectivo.GetSaldoEfectivo(TxtBuscarAge, DateAdd("d", -1, CDate(txtFecha)), gMonedaExtranjera, gdFecSis, gsCodUser, , True)
            nSobFalt = Round(nSaldoEfectivoHoy - (nSaldoEfectivoAyer + nSaldoMovimientos), 2)
            lblSaldoFinalME = Format$(nSaldoEfectivoAyer + nSaldoMovimientos, "#,##0.00")
        Else
            'nSaldoEfectivoAyer = oEfectivo.GetSaldoEfectivo(TxtBuscarAge, DateAdd("d", -1, CDate(txtFecha)), gMonedaExtranjera, gdFecSis, gsCodUser, TxtBuscarUser.Text)
            nSaldoEfectivoAyer = 0
            nSobFalt = Round(nSaldoEfectivoHoy - nSaldoMovimientos, 2)
            lblSaldoFinalME = Format$(nSaldoMovimientos, "#,##0.00")
        End If
    
        lblSaldoAntME = Format$(nSaldoEfectivoAyer, "#,##0.00")
        lblTotalCajaME = Format$(nSaldoEfectivoHoy, "#,##0.00")
        lblNetoMovME = Format$(nSaldoMovimientos, "#,##0.00")
    
        If nSobFalt > 0 Then
            lblDifME = Format$(nSobFalt, "#,##0.00")
            lblDifMEEtq.Caption = "SOBRANTE !!"
            lblDifME.ForeColor = &H800000
            lblDifMEEtq.ForeColor = &H800000
            COnSOb = COnSOb + nSobFalt
        ElseIf nSobFalt < 0 Then
            lblDifME = Format$(nSobFalt, "#,##0.00")
            lblDifMEEtq.Caption = "FALTANTE !!"
            lblDifME.ForeColor = &HC0&
            lblDifMEEtq.ForeColor = &HC0&
            COnSOb = COnSOb + (nSobFalt) * -1
        Else
            lblDifME = "SIN SOB/FALT"
            lblDifMEEtq.Caption = "Sob/Faltante"
            lblDifME.ForeColor = &H800000
            lblDifMEEtq.ForeColor = &H800000
        End If
    End If

    '**Comentado por DAOR 20080125
    ''If COnSOb > 0 Then
    'Me.cmdRegSobFalt.Visible = True
    'Me.cmdRegSobFalt.Enabled = IIf(bConsolidada = True, False, True)
    ''End If

    '**DAOR 20080125, Verificar si operación de sobr. falt. ya existe******
    Dim loCajero As COMNCajaGeneral.NCOMCajero
    Set loCajero = New COMNCajaGeneral.NCOMCajero
    If Not bConsolidada Then
        If loCajero.YaRealizoDevBilletaje(txtBuscarUser.Text, txtFecha.value, gsCodAge) Then
            cmdRegSobFalt.Visible = False
        Else
            cmdRegSobFalt.Visible = True
        End If
    End If
    Set loCajero = Nothing
    '**********************************************************************
    LblTotalME = Format$(nSaldoMovimientos + nDevolucionBilletaje, "#,##0.00")
End Sub

Private Sub cmdImpBill_Click()
Call frmCajaGenEfectivo.ImprimirBilletaje 'MIOL 20130506, SEGUN RQ13207
End Sub

Private Sub cmdImprimir_Click()
Dim lscadimp As String
Dim nTotalEfectivo As Double, nTotalCheque As Double, nTotalOrdPag As Double, nTotalTotal As Double
Dim nCarLin As Integer
Dim sTitRp1 As String, sTitRp2 As String
Dim sMoneda As String, sNumPag As String, sFchtra As String
Dim nLinPag As Integer, nCntPag As Integer
Dim sNombre As String * 35, sNombre2 As String * 30
Dim sEfectivo As String * 12, sCheques As String * 12, sOrdPag As String * 12, sTotFila As String * 12
Dim sNumReg As String * 5
Dim sUsuario As String

Dim sTotalIngresosN As String * 12, sTotalEgresosN As String * 12
Dim sTotalHabilitacionN As String * 12, sTotalDevolucionN As String * 12
Dim sTotalDevBilletajeN As String * 12, sTotalSobFaltN As String * 12
Dim sTotalPlanillaN As String * 12, sSaldoAnteriorN As String * 12
Dim sNetoMovimientosN As String * 12, sNroMovimientosN As String * 10, sSaldoFinalN As String * 12
Dim sDiferenciaN As String * 12

Dim sTotalIngresosE As String * 12, sTotalEgresosE As String * 12
Dim sTotalHabilitacionE As String * 12, sTotalDevolucionE As String * 12
Dim sTotalDevBilletajeE As String * 12, sTotalSobFaltE As String * 12
Dim sTotalPlanillaE As String * 12, sSaldoAnteriorE As String * 12
Dim sNetoMovimientosE As String * 12, sNroMovimientosE As String * 10, sSaldoFinalE As String * 12
Dim sDiferenciaE As String * 12
Dim sNumOpeMN As String * 12, sNumOpeME As String * 12 'EJVG 20110815

 
Dim sBoldON As String, sBoldOFF As String
Dim i As Long

Dim rsMN As New ADODB.Recordset
Dim rsME As New ADODB.Recordset
 
Dim oCajImp As COMNCajaGeneral.NCOMCajero
Dim oPrevio As previo.clsprevio

'***RIRO 20130702 ***************
Dim dOtrosEgresosEfectivo As ADODB.Recordset
'Fin RIRO ***********************


If txtBuscarUser.Text = "" Then
    sUsuario = "<CONSOLIDADO>"
Else
    sUsuario = txtBuscarUser.Text
End If

With rsMN
    'Crear RecordSet
    .Fields.Append "sGrupoNom", adVarChar, 50
    .Fields.Append "nTotalMov", adVarChar, 50
    .Fields.Append "sEfectivo", adVarChar, 50
    .Fields.Append "sCheque", adVarChar, 50
    .Fields.Append "sOrdenPago", adVarChar, 50
    .Fields.Append "sNoEfectivo", adVarChar, 50 'FRHU 20150206 ERS048-2014
    .Fields.Append "nTotalMonto", adVarChar, 50
    .Open
    'Llenar Recordset
    For i = 0 To fgPlanillaMN.Rows - 1
        .AddNew
        .Fields("sGrupoNom") = Trim(fgPlanillaMN.TextMatrix(i, 1))
        .Fields("nTotalMov") = Trim(fgPlanillaMN.TextMatrix(i, 2))
        .Fields("sEfectivo") = Trim(fgPlanillaMN.TextMatrix(i, 3))
        .Fields("sCheque") = Trim(fgPlanillaMN.TextMatrix(i, 4))
        .Fields("sOrdenPago") = Trim(fgPlanillaMN.TextMatrix(i, 5))
        .Fields("sNoEfectivo") = Trim(fgPlanillaMN.TextMatrix(i, 6)) 'FRHU 20150206 ERS048-2014
        '.Fields("nTotalMonto") = Trim(fgPlanillaMN.TextMatrix(i, 6))
        .Fields("nTotalMonto") = Trim(fgPlanillaMN.TextMatrix(i, 7)) 'FRHU 20150206 ERS048-2014
    Next i
End With

With rsME
    'Crear RecordSet
    .Fields.Append "sGrupoNom", adVarChar, 50
    .Fields.Append "nTotalMov", adVarChar, 50
    .Fields.Append "sEfectivo", adVarChar, 50
    .Fields.Append "sCheque", adVarChar, 50
    .Fields.Append "sOrdenPago", adVarChar, 50
    .Fields.Append "sNoEfectivo", adVarChar, 50 'FRHU 20150206 ERS048-2014
    .Fields.Append "nTotalMonto", adVarChar, 50
    .Open
    'Llenar Recordset
    For i = 0 To fgPlanillaME.Rows - 1
        .AddNew
        .Fields("sGrupoNom") = Trim(fgPlanillaME.TextMatrix(i, 1))
        .Fields("nTotalMov") = Trim(fgPlanillaME.TextMatrix(i, 2))
        .Fields("sEfectivo") = Trim(fgPlanillaME.TextMatrix(i, 3))
        .Fields("sCheque") = Trim(fgPlanillaME.TextMatrix(i, 4))
        .Fields("sOrdenPago") = Trim(fgPlanillaME.TextMatrix(i, 5))
        .Fields("sNoEfectivo") = Trim(fgPlanillaME.TextMatrix(i, 6)) 'FRHU 20150206 ERS048-2014
        '.Fields("nTotalMonto") = Trim(fgPlanillaME.TextMatrix(i, 6))
        .Fields("nTotalMonto") = Trim(fgPlanillaME.TextMatrix(i, 7)) 'FRHU 20150206 ERS048-2014
    Next i
End With


RSet sTotalIngresosN = Trim(lblTotalIngMN)
RSet sTotalEgresosN = Trim(lblTotalEgreMN)
RSet sTotalHabilitacionN = Trim(lblTotalHabMN)
RSet sTotalDevolucionN = Trim(lblTotalDevMN)
RSet sTotalDevBilletajeN = Trim(lblTotalDevBillMN)
RSet sTotalSobFaltN = Trim(lblTotalSobFalMN)
RSet sTotalPlanillaN = Trim(lblTotalMN)

RSet sTotalIngresosE = Trim(lblTotalIngME)
RSet sTotalEgresosE = Trim(lblTotalEgreME)
RSet sTotalHabilitacionE = Trim(lblTotalHabME)
RSet sTotalDevolucionE = Trim(lblTotalDevME)
RSet sTotalDevBilletajeE = Trim(lblTotalDevBillME)
RSet sTotalSobFaltE = Trim(lblTotalSobFalME)
RSet sTotalPlanillaE = Trim(LblTotalME)

RSet sSaldoAnteriorN = Trim(lblSaldoAntMN)
RSet sNetoMovimientosN = Trim(lblNetoMovMN)
RSet sSaldoFinalN = Trim(lblSaldoFinalMN)
RSet sDiferenciaN = Trim(lblDifMN)

RSet sSaldoAnteriorE = Trim(lblSaldoAntME)
RSet sNetoMovimientosE = Trim(lblNetoMovME)
RSet sSaldoFinalE = Trim(lblSaldoFinalME)
RSet sDiferenciaE = Trim(lblDifME)
'EJVG 20110815**********************
RSet sNumOpeMN = CStr(fnNumOpeMN)
RSet sNumOpeME = CStr(fnNumOpeME)
'EJVG*******************************

Set oCajImp = New COMNCajaGeneral.NCOMCajero

Set rsEgresos = oCajImp.ResumenEgresosEfectivo(Format(Me.txtFecha, "yyyyMMdd"), txtBuscarUser.Text, Trim(TxtBuscarAge.Text))

    oCajImp.IniciaImpresora gImpresora
    lscadimp = oCajImp.ImpCajeroIngEgre(rsMN, rsME, Me.txtFecha, gsNomAge, lblNroMovMN.Caption, lblNroMovME.Caption, sTotalIngresosN, sTotalEgresosN, sTotalHabilitacionN, sTotalDevolucionN, _
                 sTotalDevBilletajeN, sTotalSobFaltN, sTotalPlanillaN, sSaldoAnteriorN, _
                 sNetoMovimientosN, sSaldoFinalN, sDiferenciaN, sTotalIngresosE, sTotalEgresosE, sTotalHabilitacionE, sTotalDevolucionE, _
                 sTotalDevBilletajeE, sTotalSobFaltE, sTotalPlanillaE, sSaldoAnteriorE, _
                 sNetoMovimientosE, sSaldoFinalE, sDiferenciaE, gsCodUser, gdFecSis, gsNomCmac, lblDescAge.Caption, , , sNumOpeMN, sNumOpeME, rsEgresos) 'Add by EJVG 20110809 sNumOpeMN y sNumOpeMN
Set oCajImp = Nothing
 
Set oPrevio = New previo.clsprevio
oPrevio.Show lscadimp, "RESUMEN DE INGRESOS Y EGRESOS", True, 66, gImpresora
Set oPrevio = Nothing
End Sub

Private Sub cmdPreCuadre_Click()
Dim oPrevio As previo.clsprevio
Dim oCajImp As COMNCajaGeneral.NCOMCajero
Dim lscadimp As String

Set oCajImp = New COMNCajaGeneral.NCOMCajero
    lscadimp = oCajImp.ImpPrecuadre(gsCodAge, gsNomAge, lblDifMN.Caption, lblDifME.Caption, gdFecSis)
Set oCajImp = Nothing

'Verificamos Usuarios que tengan movimientos en efectivo y no tengan registro de billetaje
If lscadimp <> "" Then
    Set oPrevio = New previo.clsprevio
    oPrevio.Show lscadimp, "PRECUADRE DE CAJEROS", True, , gImpresora
    cmdRegSobFalt.Enabled = False
    Exit Sub
Else
    MsgBox "Precuadre Finalizado Correctamente", vbInformation, "Aviso"
    cmdRegSobFalt.Enabled = True
    cmdRegSobFalt.SetFocus
End If
End Sub

Private Sub cmdProcesar_Click()
    bProcesado = True 'DAOR 20080125
    cmdProcesar.Enabled = False
    TDesatendido.Enabled = True
    DoEvents
End Sub

Private Sub cmdRepoArqueo_Click()

'*** PEAC 20090925
If gsCodUser = Trim(Me.txtBuscarUser) Then
    MsgBox "No se puede realizar el arqueo a si mismo.", vbInformation, "Aviso"
    Exit Sub
End If

Dim lscadimp As String
Dim nTotalEfectivo As Double, nTotalCheque As Double, nTotalOrdPag As Double, nTotalTotal As Double
Dim nCarLin As Integer
Dim sTitRp1 As String, sTitRp2 As String
Dim sMoneda As String, sNumPag As String, sFchtra As String
Dim nLinPag As Integer, nCntPag As Integer
Dim sNombre As String * 35, sNombre2 As String * 30
Dim sEfectivo As String * 12, sCheques As String * 12, sOrdPag As String * 12, sTotFila As String * 12
Dim sNumReg As String * 5
Dim sUsuario As String

Dim sTotalIngresosN As String * 12, sTotalEgresosN As String * 12
Dim sTotalHabilitacionN As String * 12, sTotalDevolucionN As String * 12
Dim sTotalDevBilletajeN As String * 12, sTotalSobFaltN As String * 12
Dim sTotalPlanillaN As String * 12, sSaldoAnteriorN As String * 12
Dim sNetoMovimientosN As String * 12, sNroMovimientosN As String * 10, sSaldoFinalN As String * 12
Dim sDiferenciaN As String * 12

Dim sTotalIngresosE As String * 12, sTotalEgresosE As String * 12
Dim sTotalHabilitacionE As String * 12, sTotalDevolucionE As String * 12
Dim sTotalDevBilletajeE As String * 12, sTotalSobFaltE As String * 12
Dim sTotalPlanillaE As String * 12, sSaldoAnteriorE As String * 12
Dim sNetoMovimientosE As String * 12, sNroMovimientosE As String * 10, sSaldoFinalE As String * 12
Dim sDiferenciaE As String * 12

Dim sBoldON As String, sBoldOFF As String
Dim i As Long

Dim rsMN As New ADODB.Recordset
Dim rsME As New ADODB.Recordset
 
Dim oCajImp As COMNCajaGeneral.NCOMCajero
Dim oPrevio As previo.clsprevio

If txtBuscarUser.Text = "" Then
    sUsuario = "<CONSOLIDADO>"
Else
    sUsuario = txtBuscarUser.Text
End If

With rsMN
    'Crear RecordSet
    .Fields.Append "sGrupoNom", adVarChar, 50
    .Fields.Append "nTotalMov", adVarChar, 50
    .Fields.Append "sEfectivo", adVarChar, 50
    .Fields.Append "sCheque", adVarChar, 50
    .Fields.Append "sOrdenPago", adVarChar, 50
    .Fields.Append "sNoEfectivo", adVarChar, 50 'EAAS20180705
    .Fields.Append "nTotalMonto", adVarChar, 50
    .Open
    'Llenar Recordset
    For i = 0 To fgPlanillaMN.Rows - 1
        .AddNew
        .Fields("sGrupoNom") = Trim(fgPlanillaMN.TextMatrix(i, 1))
        .Fields("nTotalMov") = Trim(fgPlanillaMN.TextMatrix(i, 2))
        .Fields("sEfectivo") = Trim(fgPlanillaMN.TextMatrix(i, 3))
        .Fields("sCheque") = Trim(fgPlanillaMN.TextMatrix(i, 4))
        .Fields("sOrdenPago") = Trim(fgPlanillaMN.TextMatrix(i, 5))
        .Fields("sNoEfectivo") = Trim(fgPlanillaMN.TextMatrix(i, 6)) 'EAAS20180705
        .Fields("nTotalMonto") = Trim(fgPlanillaMN.TextMatrix(i, 7))
    Next i
End With

With rsME
    'Crear RecordSet
    .Fields.Append "sGrupoNom", adVarChar, 50
    .Fields.Append "nTotalMov", adVarChar, 50
    .Fields.Append "sEfectivo", adVarChar, 50
    .Fields.Append "sCheque", adVarChar, 50
    .Fields.Append "sOrdenPago", adVarChar, 50
    .Fields.Append "sNoEfectivo", adVarChar, 50 'EAAS20180705
    .Fields.Append "nTotalMonto", adVarChar, 50
    .Open
    'Llenar Recordset
    For i = 0 To fgPlanillaME.Rows - 1
        .AddNew
        .Fields("sGrupoNom") = Trim(fgPlanillaME.TextMatrix(i, 1))
        .Fields("nTotalMov") = Trim(fgPlanillaME.TextMatrix(i, 2))
        .Fields("sEfectivo") = Trim(fgPlanillaME.TextMatrix(i, 3))
        .Fields("sCheque") = Trim(fgPlanillaME.TextMatrix(i, 4))
        .Fields("sOrdenPago") = Trim(fgPlanillaME.TextMatrix(i, 5))
        .Fields("sNoEfectivo") = Trim(fgPlanillaMN.TextMatrix(i, 6)) 'EAAS20180705
        .Fields("nTotalMonto") = Trim(fgPlanillaME.TextMatrix(i, 7))
    Next i
End With


RSet sTotalIngresosN = Trim(lblTotalIngMN)
RSet sTotalEgresosN = Trim(lblTotalEgreMN)
RSet sTotalHabilitacionN = Trim(lblTotalHabMN)
RSet sTotalDevolucionN = Trim(lblTotalDevMN)
RSet sTotalDevBilletajeN = Trim(lblTotalDevBillMN)
RSet sTotalSobFaltN = Trim(lblTotalSobFalMN)
RSet sTotalPlanillaN = Trim(lblTotalMN)

RSet sTotalIngresosE = Trim(lblTotalIngME)
RSet sTotalEgresosE = Trim(lblTotalEgreME)
RSet sTotalHabilitacionE = Trim(lblTotalHabME)
RSet sTotalDevolucionE = Trim(lblTotalDevME)
RSet sTotalDevBilletajeE = Trim(lblTotalDevBillME)
RSet sTotalSobFaltE = Trim(lblTotalSobFalME)
RSet sTotalPlanillaE = Trim(LblTotalME)

RSet sSaldoAnteriorN = Trim(lblSaldoAntMN)
RSet sNetoMovimientosN = Trim(lblNetoMovMN)
RSet sSaldoFinalN = Trim(lblSaldoFinalMN)
RSet sDiferenciaN = Trim(lblDifMN)

RSet sSaldoAnteriorE = Trim(lblSaldoAntME)
RSet sNetoMovimientosE = Trim(lblNetoMovME)
RSet sSaldoFinalE = Trim(lblSaldoFinalME)
RSet sDiferenciaE = Trim(lblDifME)

Set oCajImp = New COMNCajaGeneral.NCOMCajero
    oCajImp.IniciaImpresora gImpresora
    lscadimp = oCajImp.ImpCajeroIngEgre(rsMN, rsME, Me.txtFecha, gsNomAge, lblNroMovMN.Caption, lblNroMovME.Caption, sTotalIngresosN, sTotalEgresosN, sTotalHabilitacionN, sTotalDevolucionN, _
                 sTotalDevBilletajeN, sTotalSobFaltN, sTotalPlanillaN, sSaldoAnteriorN, _
                 sNetoMovimientosN, sSaldoFinalN, sDiferenciaN, sTotalIngresosE, sTotalEgresosE, sTotalHabilitacionE, sTotalDevolucionE, _
                 sTotalDevBilletajeE, sTotalSobFaltE, sTotalPlanillaE, sSaldoAnteriorE, _
                 sNetoMovimientosE, sSaldoFinalE, sDiferenciaE, gsCodUser, gdFecSis, gsNomCmac, lblDescAge.Caption, True, Me.txtBuscarUser)
Set oCajImp = Nothing
 
Set oPrevio = New previo.clsprevio
'oPrevio.Show lscadimp, "RESUMEN DE INGRESOS Y EGRESOS", True, 66
oPrevio.Show lscadimp, "REPORTE PARA ARQUEO", True, 66, gImpresora
Set oPrevio = Nothing

End Sub

Private Sub TDesatendido_Timer()
    TDesatendido.Enabled = False
    Pg.Visible = True
    Call CargaPlanilla
    DoEvents
    cmdProcesar.Enabled = True
    Pg.Visible = False
    DoEvents
End Sub

Private Sub cmdRegSobFalt_Click()
Dim oCajero As COMNCajaGeneral.NCOMCajero  'nCajero
Dim oCont As COMNContabilidad.NCOMContFunciones  'NContFunciones
Dim sMovNro As String, sOperacion As String
Dim nDiferenciaMN As Double, nDiferenciaME As Double
Dim nCajaMN As Double, nCajaME As Double
Dim sMovDesc As String
Dim lnSobrFaltMN As Currency, lnSobrFaltME As Currency 'DAOR 20080125
Dim mensajebox As String
'**DAOR 20080125 ********************************
If Not bProcesado Then
    MsgBox "Primero debe procesar la operación", vbInformation, "Aviso"
    cmdProcesar.SetFocus
    Exit Sub
End If
'************************************************
If cmdRegSobFalt.Caption = "Cierre Caja Agencia" Then

    'Add by GITU 18-10-2011
    If Not ValidaConfDevTarjetas Then
        cmdSalir.SetFocus
        Exit Sub
    End If
    'End GITU
    
    If MsgBox("¿Desea Cerrar la Caja de la " & gsNomAge & " ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set oCont = New COMNContabilidad.NCOMContFunciones
        sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set oCont = Nothing
        Set oCajero = New COMNCajaGeneral.NCOMCajero
        sMovDesc = "Cierre Caja Agencia " & gsCodAge & " - " & gsNomAge
        oCajero.GrabaCierreAgencia sMovNro, gOpeCajaCierreAgencia, sMovDesc
        Set oCajero = Nothing
        cmdRegSobFalt.Enabled = False
        cmdPreCuadre.Enabled = False
    End If
    Exit Sub
End If

If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
    mensajebox = "¿Desea Registrar Pre Cuadre de Operaciones?"
Else
    mensajebox = "¿Desea Registrar el Sobrante/Faltante y la Devolución de su Billetaje?"
End If

If MsgBox(mensajebox, vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim sUsuario As String
    Set oCont = New COMNContabilidad.NCOMContFunciones
    '**DAOR 20080125 *****************************************************
    'sMovNro = oCont.GeneraMovNro(gdFecSis, TxtBuscarAge, gsCodUser)
    If fsMovNroRegEfec = "" Then
        sMovNro = oCont.GeneraMovNro(gdFecSis, TxtBuscarAge, gsCodUser)
    Else
        sMovNro = fsMovNroRegEfec
    End If
    '*********************************************************************
    Set oCont = Nothing
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    nDiferenciaMN = 0
    nDiferenciaME = 0
    
    If IsNumeric(lblDifMN) Then nDiferenciaMN = CDbl(lblDifMN)
    If IsNumeric(lblDifME) Then nDiferenciaME = CDbl(lblDifME)
    nCajaMN = CDbl(lblTotalCajaMN)
    nCajaME = CDbl(lblTotalCajaME)
    
    ' *** control adicionado por ejrs  *********
    '**Modificado por DAOR 20080125 *******************************************************
    '**Si la operación no se realiza seguidamente después del registro de efectivo*********
    If fsMovNroRegEfec = "" Then
        If Abs(nDiferenciaMN) > 1 Then
            If MsgBox("Sobrante/faltante de [" & nDiferenciaMN & "] es mayor a 1 sol" & Chr(13) & "Desea continuar", vbYesNo + vbExclamation, "aviso") = vbNo Then
                Exit Sub
            End If
        End If
        
        If Abs(nDiferenciaME) > 1 Then
            If MsgBox("Sobrante/faltante de [" & nDiferenciaME & "] es mayor a 1 dolar" & Chr(13) & "Desea continuar", vbYesNo + vbExclamation, "aviso") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    '***************************************************************************************
    '********************************************
    sUsuario = gsCodUser
    If bPlanillaBoveda Then
        sUsuario = gsUsuarioBOVEDA
        If nDiferenciaMN <> 0 Then
            sMovDesc = lblDifMNEtq & " = " & lblDifMN
            sOperacion = IIf(nDiferenciaMN >= 0, gOpeBoveAgeRegSobrante, gOpeBoveAgeRegFaltante)
            oCajero.GrabaRegistroSobranteFaltante sMovNro, sOperacion, sMovDesc, nDiferenciaMN, gMonedaNacional
        End If
        If nDiferenciaME <> 0 Then
            sMovDesc = lblDifMEEtq & " = " & lblDifME
            sOperacion = IIf(nDiferenciaME >= 0, gOpeBoveAgeRegSobrante, gOpeBoveAgeRegFaltante)
            oCajero.GrabaRegistroSobranteFaltante sMovNro, sOperacion, sMovDesc, nDiferenciaME, gMonedaExtranjera
        End If
    Else
        If nDiferenciaMN <> 0 Then
            sMovDesc = lblDifMNEtq & " = " & lblDifMN
            If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
                sOperacion = IIf(nDiferenciaMN >= 0, COMDConstSistema.gOpeBoveAgeRegSobrantePreCuadre, COMDConstSistema.gOpeBoveAgeRegFaltantePreCuadre)
                oCajero.GrabaRegistroSobranteFaltante sMovNro, sOperacion, sMovDesc, nDiferenciaMN, gMonedaNacional
            Else
                sOperacion = IIf(nDiferenciaMN >= 0, gOpeHabCajRegSobrante, gOpeHabCajRegFaltante)
                oCajero.GrabaRegistroSobranteFaltante sMovNro, sOperacion, sMovDesc, nDiferenciaMN, gMonedaNacional
            End If
        End If
        If nDiferenciaME <> 0 Then
            sMovDesc = lblDifMEEtq & " = " & lblDifME
            If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
                sOperacion = IIf(nDiferenciaME >= 0, "901041", "901042")
                oCajero.GrabaRegistroSobranteFaltante sMovNro, sOperacion, sMovDesc, nDiferenciaME, gMonedaExtranjera
            Else
                sOperacion = IIf(nDiferenciaME >= 0, gOpeHabCajRegSobrante, gOpeHabCajRegFaltante)
                oCajero.GrabaRegistroSobranteFaltante sMovNro, sOperacion, sMovDesc, nDiferenciaME, gMonedaExtranjera
            End If
        End If
    End If
    
    
    If oCajero.GetBuscarMov(gOpeHabCajDevBilletaje, sUsuario, Left(sMovNro, 8), gMonedaNacional, gsCodAge) = False Then
          sMovDesc = "Devolución Billetaje MN = " & lblTotalCajaMN
          If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
            oCajero.GrabaHabilitaAgenciaPreCuadre sMovNro, gOpeHabCajDevBilletaje, sMovDesc, nCajaMN, Usuario.cAreaCodAct, TxtBuscarAge, gMonedaNacional, sUsuario, gsUsuarioBOVEDA
          Else
            oCajero.GrabaHabilitaAgencia sMovNro, gOpeHabCajDevBilletaje, sMovDesc, nCajaMN, Usuario.cAreaCodAct, TxtBuscarAge, gMonedaNacional, sUsuario, gsUsuarioBOVEDA
          End If
    Else
           MsgBox "Cajero ya registros Sobrante y Faltante", vbInformation, "aviso"
           cmdRegSobFalt.Enabled = False
           Exit Sub
    End If
    
    If oCajero.GetBuscarMov(gOpeHabCajDevBilletaje, sUsuario, Left(sMovNro, 8), gMonedaExtranjera, gsCodAge) = False Then
          sMovDesc = "Devolución Billetaje ME = " & lblTotalCajaME
          
        If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
            oCajero.GrabaHabilitaAgenciaPreCuadre sMovNro, gOpeHabCajDevBilletaje, sMovDesc, nCajaME, Usuario.cAreaCodAct, TxtBuscarAge, gMonedaExtranjera, sUsuario, gsUsuarioBOVEDA
        Else
            oCajero.GrabaHabilitaAgencia sMovNro, gOpeHabCajDevBilletaje, sMovDesc, nCajaME, Usuario.cAreaCodAct, TxtBuscarAge, gMonedaExtranjera, sUsuario, gsUsuarioBOVEDA
        End If
        
    End If
     
    '**DAOR 20080125 *************************************************************
    If fnMovNroRegEfec > 0 Then
        If IsNumeric(lblDifMN.Caption) Then lnSobrFaltMN = CCur(lblDifMN.Caption) Else lnSobrFaltMN = 0
        If IsNumeric(lblDifME.Caption) Then lnSobrFaltME = CCur(lblDifME.Caption) Else lnSobrFaltME = 0
        
        If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
            oCajero.GrabaRegistroEfectivoResumenPreCuadre fnMovNroRegEfec, sUsuario, fnNumRegEfec, CCur(lblTotalIngMN.Caption), CCur(lblTotalEgreMN.Caption), CCur(lblTotalCajaMN.Caption), lnSobrFaltMN, CCur(lblTotalIngME.Caption), CCur(lblTotalEgreME.Caption), CCur(lblTotalCajaME.Caption), lnSobrFaltME
        Else
            oCajero.GrabaRegistroEfectivoResumen fnMovNroRegEfec, sUsuario, fnNumRegEfec, CCur(lblTotalIngMN.Caption), CCur(lblTotalEgreMN.Caption), CCur(lblTotalCajaMN.Caption), lnSobrFaltMN, CCur(lblTotalIngME.Caption), CCur(lblTotalEgreME.Caption), CCur(lblTotalCajaME.Caption), lnSobrFaltME
        End If
        
    End If
    bRegistrado = True
    '*****************************************************************************
     
    Set oCajero = Nothing
    MsgBox "Registro finalizado con éxito", vbInformation, "Aviso"
    
    cmdRegSobFalt.Enabled = False
    cmdProcesar_Click
        
End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Public Sub Inicia(ByVal bBoveda As Boolean, ByVal bConsolidada As Boolean, Optional bTodasAgencias As Boolean = False, _
    Optional ByVal psMovNroRegEfec As String = "", Optional ByVal pnMovNroRegEfec As Long = 0, Optional ByVal pnNumRegEfec As Integer = 0, Optional ByVal psOpeCod As Long = 0)
    
    'RIRO20140930 ERS115 ***************************
    Dim bPermisoCargo As Boolean
    bPermisoCargo = False
    Set oGen = New COMDConstSistema.DCOMGeneral
    bPermisoCargo = oGen.VerificaExistePermisoCargo(gsCodCargo, 5, gsCodPersUser) 'RIRO20141027 ERS159
    bTodasAgencias = bPermisoCargo
    'END RIRO **************************************
    
    Dim oCajero As COMNCajaGeneral.NCOMCajero 'nCajero

    'Set oGen = New COMDConstSistema.DCOMGeneral  'DGeneral RIRO 20141113 ERS115

    Usuario.Inicio gsCodUser
    
    '**MADM 20110118 *************************************
    If psOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
        fsOpeCod = psOpeCod
        cmdPreCuadre.Visible = False
        txtBuscarUser.Enabled = False
        TxtBuscarAge.Enabled = False
        txtBuscarUser.Text = gsCodUser
        TxtBuscarAge.Text = gsCodAge
        lblDescAge = gsNomAge
    Else
    '******************************************************
            bPlanillaBoveda = bBoveda
            If bConsolidada And Not bTodasAgencias Then
                'Abre el boton de Cierre de Agencia pero no lo habilita hasta que no se realice el precuadre
                cmdRegSobFalt.Visible = True
                cmdRegSobFalt.Caption = "Cierre Caja Agencia"
                cmdRegSobFalt.Enabled = False
                cmdPreCuadre.Visible = True
                txtBuscarUser.Enabled = True
                TxtBuscarAge.Enabled = False
                TxtBuscarAge.Text = gsCodAge
                lblDescAge = gsNomAge
                txtBuscarUser.psRaiz = "USUARIOS"
                txtBuscarUser.rs = oGen.GetUserAreaAgencia(Usuario.cAreaCodAct, Usuario.CodAgeAct, "", False)
                txtBuscarUser.Enabled = True
                txtBuscarUser.Text = ""
                Set oCajero = New COMNCajaGeneral.NCOMCajero
                If oCajero.YaRealizoCierreAgencia(gsCodAge, gdFecSis) Then
                    cmdRegSobFalt.Visible = False
                    cmdPreCuadre.Visible = False
                Else
                    cmdRegSobFalt.Visible = True
                    cmdPreCuadre.Visible = True
                End If
                Set oCajero = Nothing
            ElseIf bBoveda Then
                cmdPreCuadre.Visible = False
                If psMovNroRegEfec = "" Then
                    cmdRegSobFalt.Visible = False
                Else
                    cmdRegSobFalt.Visible = True
                End If
                txtBuscarUser.Enabled = False
                TxtBuscarAge.Enabled = False
                txtBuscarUser.Text = gsUsuarioBOVEDA
                TxtBuscarAge.Text = gsCodAge
                lblDescAge = gsNomAge
                Set oCajero = New COMNCajaGeneral.NCOMCajero
                If oCajero.YaRealizoDevBilletaje(gsUsuarioBOVEDA, gdFecSis, gsCodAge) Then
                    cmdRegSobFalt.Visible = False
                Else
                    cmdRegSobFalt.Visible = True
                End If
                Set oCajero = Nothing
            ElseIf bConsolidada And bTodasAgencias Then
                cmdPreCuadre.Visible = False
                cmdRegSobFalt.Visible = False
                TxtBuscarAge.psRaiz = "Agencias"
                TxtBuscarAge.rs = oGen.GetNombreAgencias
                TxtBuscarAge.Enabled = True
                TxtBuscarAge.Text = gsCodAge
                lblDescAge = gsNomAge
                txtBuscarUser.psRaiz = "USUARIOS"
                txtBuscarUser.rs = oGen.GetUserAreaAgencia(Usuario.cAreaCodAct, Usuario.CodAgeAct, "", False)
                txtBuscarUser.Enabled = True
                txtBuscarUser.Text = ""
            Else
                '**DAOR 20080125 *********************************************************
                If psMovNroRegEfec = "" Then
                    Set oCajero = New COMNCajaGeneral.NCOMCajero
                    If Not oCajero.YaRealizoDevBilletaje(gsCodUser, gdFecSis, gsCodAge) Then
                        MsgBox "Aun no ha realizado la operación de registro de efectivo, ésta opción es de sólo lectura y está disponible despues del registro de efectivo", vbInformation, "Aviso"
                        Unload Me
                        Exit Sub
                    End If
                    Set oCajero = Nothing
                End If
                '*************************************************************************
                cmdPreCuadre.Visible = False
                txtBuscarUser.Enabled = False
                TxtBuscarAge.Enabled = False
                txtBuscarUser.Text = gsCodUser
                TxtBuscarAge.Text = gsCodAge
                lblDescAge = gsNomAge
                '**Comentado por DAOR 20080128 *******************************************
                '**La validación para deshabilitar el boton se encuentra en la opción de Procesar
                'Set oCajero = New COMNCajaGeneral.NCOMCajero
                'If oCajero.YaRealizoDevBilletaje(gsCodUser, gdFecSis, gsCodAge) Then
                '    cmdRegSobFalt.Visible = False
                'Else
                '    cmdRegSobFalt.Visible = True
                'End If
                'Set oCajero = Nothing
                '**************************************************************************
            End If
    End If
    '**DAOR 20080125, Datos necesarios del registro de fectivo*********************
    fsMovNroRegEfec = psMovNroRegEfec
    fnMovNroRegEfec = pnMovNroRegEfec
    fnNumRegEfec = pnNumRegEfec
    '******************************************************************************
''''''''     'MADM 20110201
    If gsOpeCod = "901014" Then
         Set loVistoElectronico = New frmVistoElectronico
         lbVistoVal = loVistoElectronico.Inicio(3, gsOpeCod)
         If lbVistoVal = False Then
            Unload Me
            Exit Sub
         End If
    End If
''''''''    'MADM 20101013
    Set oGen = Nothing
    Me.Show 1
End Sub

Private Sub Form_Load()
    '**DAOR 20080125 *******************************
    bProcesado = False
    bRegistrado = False
    '***********************************************
    tabIngEgre.Tab = 0
    txtFecha.value = gdFecSis
    Me.Caption = IIf(Trim(gsOpeDesc) = "", "Resumen de Ingresos y Egresos", gsOpeDesc)
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    txtBuscarUser = ""
    fsOpeCod = 0
    sGrupoOpe = "" 'RIRO20150608 ERS162-2014
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not bRegistrado And fsMovNroRegEfec <> "" Then
        MsgBox "Necesita registrar la operación de sobrante/faltante", vbInformation, "Aviso"
        Cancel = 1
        fsOpeCod = 0
        'EJVG 20110809
        fnNumOpeMN = 0
        fnNumOpeME = 0
        'END EJVG
        Exit Sub
    End If
    Set oGen = Nothing
End Sub

Private Sub TxtBuscarAge_EmiteDatos()
lblDescAge = TxtBuscarAge.psDescripcion
If TxtBuscarAge <> "" And lblDescAge <> "" Then
    txtBuscarUser = ""
    txtBuscarUser.psRaiz = "USUARIOS " & TxtBuscarAge.psDescripcion
    txtBuscarUser.Enabled = True
    'If bConsolidada And bTodasAgencias Then
        'TxtBuscarUser.rs = oGen.GetUserAreaAgenciaResumenIngEgre(Usuario.cAreaCodAct, TxtBuscarAge)
        Set oGen = New COMDConstSistema.DCOMGeneral
            txtBuscarUser.rs = oGen.GetUserAreaAgenciaResumenIngEgre("026", TxtBuscarAge)
        Set oGen = Nothing
    'Else
    '    TxtBuscarUser.rs = oGen.GetUserAreaAgencia(Usuario.cAreaCodAct, TxtBuscarAge)
    'End If
End If
End Sub

Private Sub CargaPlanilla()

    Dim rsC As ADODB.Recordset
    Dim rsM As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim bBoveda As Boolean
    Dim bConsolidada As Boolean
    Dim oCajero As COMNCajaGeneral.NCOMCajero
    Dim Pbarra As Integer

    Set oCajero = New COMNCajaGeneral.NCOMCajero
    
    'RIRO20150608 ERS162-2014 **********************
    Dim oConst As COMDConstSistema.DCOMGeneral
    Set oConst = New COMDConstSistema.DCOMGeneral
    sGrupoOpe = oConst.LeeConstSistema(506) & "," & oConst.LeeConstSistema(517)
    Set oConst = Nothing
    'END RIRO **************************************
    
    If txtBuscarUser.Text = "" Then
        bConsolidada = True
        bBoveda = False
    ElseIf txtBuscarUser.Text = gsUsuarioBOVEDA Then
        bConsolidada = False
        bBoveda = True
    Else
        bConsolidada = False
        bBoveda = False
    End If
    Set oGen = New COMDConstSistema.DCOMGeneral
    Set rs = New ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Set rsM = oGen.GetConstante(gMoneda)

    'Limpiamos y  formateamos todo
    fgPlanillaMN.Clear
    fgPlanillaMN.FormaCabecera
    fgPlanillaMN.Rows = 2
    fgPlanillaME.Clear
    fgPlanillaME.FormaCabecera
    fgPlanillaME.Rows = 2

    Pbarra = 0

    Do While Not rsM.EOF
    
        'Cargamos la estructura
        Set rsC = oGen.GetConstante(gGruposIngEgre)
        'recorremos para tod la estructura
        Pg.Min = 0
        If Pbarra = 0 Then
            Pbarra = 1
            Pg.Max = (rsC.RecordCount) * 2
            Pg.value = 0
        End If

        Do While Not rsC.EOF
        Pg.value = Pg.value + 1
            Select Case rsM!nConsValor
                'segun la moneda agrega fila en el tipo de grid
                Case gMonedaNacional
                    fgPlanillaMN.AdicionaFila
                    fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 1) = rsC("cDescripcion")
                    fgPlanillaMN.BackColorRow &HFDCF09, True
                Case gMonedaExtranjera
                    fgPlanillaME.AdicionaFila
                    fgPlanillaME.TextMatrix(fgPlanillaME.row, 1) = rsC("cDescripcion")
                    fgPlanillaME.BackColorRow &HFDCF09, True
            End Select
            
            Set oCajero = New COMNCajaGeneral.NCOMCajero
            
            'segun el grupo
            Select Case rsC!nConsValor
                Case gGruposIngEgreAgeLocal
                    If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
                        Set rs = oCajero.GetOpeIngEgreCaptacionesPreCuadre(gGruposIngEgreAgeLocal, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser, gsCodCMAC)
                        CargaGrid rsM!nConsValor, rs
                        Set rs = oCajero.GetOpeIngEgreColocacionesPreCuadre(gGruposIngEgreAgeLocal, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser, gsCodCMAC)
                        CargaGrid rsM!nConsValor, rs
                    Else
                        Set rs = oCajero.GetOpeIngEgreCaptaciones(gGruposIngEgreAgeLocal, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser, gsCodCMAC)
                        CargaGrid rsM!nConsValor, rs
                        Set rs = oCajero.GetOpeIngEgreColocaciones(gGruposIngEgreAgeLocal, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser, gsCodCMAC)
                        CargaGrid rsM!nConsValor, rs
                    End If
                Case gGruposIngEgreOtraAgencia
                    If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
                        Set rs = oCajero.GetOpeIngEgreCaptacionesPreCuadre(gGruposIngEgreOtraAgencia, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser, gsCodCMAC)
                        CargaGrid rsM!nConsValor, rs
                        Set rs = oCajero.GetOpeIngEgreColocacionesPreCuadre(gGruposIngEgreOtraAgencia, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser, gsCodCMAC)
                        CargaGrid rsM!nConsValor, rs
                    Else
                        Set rs = oCajero.GetOpeIngEgreCaptaciones(gGruposIngEgreOtraAgencia, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser, gsCodCMAC)
                        CargaGrid rsM!nConsValor, rs
                        Set rs = oCajero.GetOpeIngEgreColocaciones(gGruposIngEgreOtraAgencia, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser, gsCodCMAC)
                        CargaGrid rsM!nConsValor, rs
                    End If
                Case gGruposIngEgreOtraCMAC
                    If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
                        Set rs = oCajero.GetOpeIngEgreOpeCMACsPreCuadre(gGruposIngEgreOtraCMAC, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                        CargaGrid rsM!nConsValor, rs
                    Else
                        Set rs = oCajero.GetOpeIngEgreOpeCMACs(gGruposIngEgreOtraCMAC, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                        CargaGrid rsM!nConsValor, rs
                    End If
                Case gGruposIngEgreOtrasOpe     'otras operaciones
                    If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
                        Set rs = oCajero.GetOpeIngEgreOtrasOpePreCuadre(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                        CargaGrid rsM!nConsValor, rs
                    Else
                        Set rs = oCajero.GetOpeIngEgreOtrasOpe(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                        CargaGrid rsM!nConsValor, rs
                    End If
                Case gGruposIngEgreServicios    'servicios
                   If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
                        Set rs = oCajero.GetOpeIngEgreServiciosPreCuadre(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                        CargaGrid rsM!nConsValor, rs
                   Else
                        Set rs = oCajero.GetOpeIngEgreServicios(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                        CargaGrid rsM!nConsValor, rs
                    End If
                Case gGruposIngEgreCompraVenta  'Compra venta
                   If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
                        Set rs = oCajero.GetOpeIngEgreCompraVentaPreCuadre(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                        CargaGrid rsM!nConsValor, rs
                    Else
                        Set rs = oCajero.GetOpeIngEgreCompraVenta(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                        CargaGrid rsM!nConsValor, rs
                    End If
                Case gGruposIngEgreSobFalt      'sobrante faltante
                     If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
                        Set rs = oCajero.GetOpeIngEgreSobranteFaltantePreCuadre(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                     Else
                        Set rs = oCajero.GetOpeIngEgreSobranteFaltante(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                    End If
                    CargaGrid rsM!nConsValor, rs
                Case gGruposIngEgreHabDev
                    If bConsolidada Then
                        Set rs = oCajero.GetOpeIngEgreHabDev(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge.Text, rsM!nConsValor, gdFecSis, gsCodUser)
                    ElseIf bBoveda Then
                        Set rs = oCajero.GetOpeIngEgreHabDevCajero(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, gsUsuarioBOVEDA, rsM!nConsValor, gdFecSis, gsCodUser)
                        CargaGrid rsM!nConsValor, rs
                        Set rs = oCajero.GetOpeIngEgreHabDev(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge.Text, rsM!nConsValor, gdFecSis, gsCodUser)
                    Else
                        If fsOpeCod = COMDConstSistema.gOpeBoveAgeRegPreCuadre Then
                            Set rs = oCajero.GetOpeIngEgreHabDevCajero(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                            CargaGrid rsM!nConsValor, rs
                            Set rs = oCajero.GetOpeIngEgreHabDevBillCajero(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                        Else
                            Set rs = oCajero.GetOpeIngEgreHabDevCajero(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                            CargaGrid rsM!nConsValor, rs
                            Set rs = oCajero.GetOpeIngEgreHabDevBillCajero(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                        End If
                    End If
                    CargaGrid rsM!nConsValor, rs
                Case 8
                    Set rs = oCajero.GetOpePreCuadreOperaciones(rsC!nConsValor, CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis, gsCodUser)
                    CargaGrid rsM!nConsValor, rs
                
                'Agregado por RIRO el 20130410 **
                Case gGruposIngRecaudo
             
                    Set rs = oCajero.GetOpeIngEgreRecaudo(CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis)
                    CargaGrid rsM!nConsValor, rs
                Case gGruposIngEgreNoEfectivo 'FRHU 20141203 ERS048-2014
                    Set rs = oCajero.GetOpeIngEgreNoEfectivo(CDate(txtFecha), TxtBuscarAge, txtBuscarUser, rsM!nConsValor, gdFecSis)
                    CargaGrid rsM!nConsValor, rs
                ''ALPA 20161209
                'Case gGruposTransferencia_CCE
                '    Set rs = oCajero.GetTransferenciaCCE(CDate(txtFecha), TxtBuscarAge, TxtBuscarUser, rsM!nConsValor, gdFecSis)
                '    CargaGrid rsM!nConsValor, rs

                    
            End Select
            Set oCajero = Nothing
            rsC.MoveNext
        Loop
        rsC.Close
        Set rsC = Nothing
        rsM.MoveNext
    Loop
    rsM.Close
    Set rsM = Nothing

    CalculaDatosPlanilla bBoveda, bConsolidada
    Set oGen = Nothing
    
    'EJVG 20110809*****************
    fnNumOpeMN = obtenerNroOperacion(gMonedaNacional)
    fnNumOpeME = obtenerNroOperacion(gMonedaExtranjera)
    '******************************
End Sub

'**EJVG 20110809 **************************************************
Private Function obtenerNroOperacion(pnMoneda As Moneda) As Integer
    Dim oCajero As New COMNCajaGeneral.NCOMCajero
    Dim rsGruposOpeITF As New ADODB.Recordset
    Dim lbITF As Boolean
    Dim lnNumOpe As Integer
    Dim fg As FlexEdit
    Dim i As Integer
    
    lbITF = False
    Set rsGruposOpeITF = oCajero.obtenerGruposOpeITF()
    
    Select Case pnMoneda
        Case gMonedaNacional
            Set fg = fgPlanillaMN
        Case gMonedaExtranjera
            Set fg = fgPlanillaME
        Case Else
            obtenerNroOperacion = 0
            Exit Function
    End Select
    
    For i = 0 To fg.Rows - 1
        Do While Not rsGruposOpeITF.EOF
            If fg.TextMatrix(i, 7) = rsGruposOpeITF("cGrupoCod") Then
                lbITF = True
                Exit Do
            End If
            rsGruposOpeITF.MoveNext
        Loop
        If lbITF = False And IsNumeric(fg.TextMatrix(i, 2)) Then
            lnNumOpe = lnNumOpe + fg.TextMatrix(i, 2)
        End If
        lbITF = False
        If i < fg.Rows - 1 Then
            rsGruposOpeITF.MoveFirst
        End If
    Next
 
    Set rsGruposOpeITF = Nothing
    Set oCajero = Nothing
    obtenerNroOperacion = lnNumOpe
End Function
'End EJVG********************************************************************

Private Sub CargaGrid(ByVal pnMoneda As Moneda, ByVal rs As ADODB.Recordset)
Select Case pnMoneda
    Case gMonedaNacional
        If rs Is Nothing Then Exit Sub
        Do While Not rs.EOF
            If gsCodAge <> rs("cCodAge") And rs("cGrupoCod") = "" Then
                fgPlanillaMN.AdicionaFila
                fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 1) = rs!cGrupoNombre
                fgPlanillaMN.BackColorRow &HE0E0E0, True
            End If
            If rs!cGrupoCod <> "" Then
                fgPlanillaMN.AdicionaFila
                fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 1) = rs!cGrupoNombre
                fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 2) = rs!nTotalMov
                fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 3) = rs!Efectivo
                fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 4) = rs!cheque
                fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 5) = rs!OrdenPago
                'FRHU 20141203 ERS48-2014
                'fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 6) = rs!Efectivo + rs!cheque + rs!OrdenPago
                'fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 7) = rs("cGrupoCod")
                'If rs("cGrupoCod") = "10001001" Or rs("cGrupoCod") = "10001002" Then RIRO20150608 ERS162-2014
                If InStr(1, Trim(sGrupoOpe), Trim(rs("cGrupoCod"))) > 0 Then 'RIRO20150608 ERS162-2014
                    fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 6) = rs!NoEfectivo
                Else
                    fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 6) = 0#
                End If
                fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 7) = rs!Efectivo + rs!cheque + rs!OrdenPago
                fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 8) = rs("cGrupoCod")
                'FIN FRHU 20141203
            End If
            rs.MoveNext
        Loop
        fgPlanillaMN.FormateaColumnas
        rs.Close
        Set rs = Nothing
    Case gMonedaExtranjera
        If rs Is Nothing Then Exit Sub
        Do While Not rs.EOF
            If gsCodAge <> rs("cCodAge") And rs("cGrupoCod") = "" Then
                fgPlanillaME.AdicionaFila
                fgPlanillaME.TextMatrix(fgPlanillaME.row, 1) = rs!cGrupoNombre
                fgPlanillaME.BackColorRow &HE0E0E0, True
            End If
            If rs!cGrupoCod <> "" Then
                fgPlanillaME.AdicionaFila
                fgPlanillaME.TextMatrix(fgPlanillaME.row, 1) = rs!cGrupoNombre
                fgPlanillaME.TextMatrix(fgPlanillaME.row, 1) = rs!cGrupoNombre
                fgPlanillaME.TextMatrix(fgPlanillaME.row, 2) = rs!nTotalMov
                fgPlanillaME.TextMatrix(fgPlanillaME.row, 3) = rs!Efectivo
                fgPlanillaME.TextMatrix(fgPlanillaME.row, 4) = rs!cheque
                fgPlanillaME.TextMatrix(fgPlanillaME.row, 5) = rs!OrdenPago
                'FRHU 20141203 ERS48-2014
                'fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 6) = rs!Efectivo + rs!cheque + rs!OrdenPago
                'fgPlanillaMN.TextMatrix(fgPlanillaMN.row, 7) = rs("cGrupoCod")
                'If rs("cGrupoCod") = "10001001" Or rs("cGrupoCod") = "10001002" Then
                If InStr(1, Trim(sGrupoOpe), Trim(rs("cGrupoCod"))) > 0 Then 'RIRO20151211
                    fgPlanillaME.TextMatrix(fgPlanillaME.row, 6) = rs!NoEfectivo
                Else
                    fgPlanillaME.TextMatrix(fgPlanillaME.row, 6) = 0#
                End If
                fgPlanillaME.TextMatrix(fgPlanillaME.row, 7) = rs!Efectivo + rs!cheque + rs!OrdenPago
                fgPlanillaME.TextMatrix(fgPlanillaME.row, 8) = rs("cGrupoCod")
                'FIN FRHU 20141203
            End If
            rs.MoveNext
        Loop
        fgPlanillaME.FormateaColumnas
        rs.Close
        Set rs = Nothing
End Select
End Sub
 

