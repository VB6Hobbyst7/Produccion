VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCajeroIngEgre 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresos y Egresos de Cajero"
   ClientHeight    =   7815
   ClientLeft      =   510
   ClientTop       =   735
   ClientWidth     =   10905
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOtrasAge 
      BackColor       =   &H80000000&
      Caption         =   "&OtrasAgencias"
      Height          =   360
      Left            =   2400
      TabIndex        =   68
      Top             =   7155
      Width           =   1305
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   3390
      Top             =   7185
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   9480
      TabIndex        =   65
      Top             =   7125
      Width           =   1305
   End
   Begin VB.CommandButton cmdRegSobFalt 
      Caption         =   "&Registrar Sobrante Faltante"
      Height          =   360
      Left            =   135
      TabIndex        =   64
      Top             =   7155
      Width           =   2280
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H80000004&
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   8190
      TabIndex        =   63
      Top             =   7125
      Width           =   1305
   End
   Begin TabDlg.SSTab tabIngEgre 
      Height          =   6435
      Left            =   75
      TabIndex        =   7
      Top             =   630
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   11351
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   617
      TabCaption(0)   =   "Moneda &Nacional"
      TabPicture(0)   =   "frmCajeroIngEgre.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "fgPlanillaMN"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Moneda &Extranjera"
      TabPicture(1)   =   "frmCajeroIngEgre.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fgPlanillaME"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Otras Agencias MN"
      TabPicture(2)   =   "frmCajeroIngEgre.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fgOtrasAgMN"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Otras Agencias ME"
      TabPicture(3)   =   "frmCajeroIngEgre.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fgOtrasAgME"
      Tab(3).ControlCount=   1
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
         Left            =   90
         TabIndex        =   34
         Top             =   4380
         Width           =   10500
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
            Left            =   120
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
            Left            =   1290
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
            Left            =   120
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
            Left            =   1290
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
            Left            =   2760
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
            Left            =   3900
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
            Left            =   2760
            TabIndex        =   52
            Top             =   690
            Width           =   1110
         End
         Begin VB.Label lbltotalDevME 
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
            Left            =   3900
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
            Left            =   2760
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
            Left            =   3900
            TabIndex        =   49
            Top             =   990
            Width           =   1305
         End
         Begin VB.Label Label29 
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
            Left            =   120
            TabIndex        =   48
            Top             =   1035
            Width           =   1095
         End
         Begin VB.Label lbltotalSobFalME 
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
            Left            =   1290
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
            Left            =   2790
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
            Left            =   3555
            TabIndex        =   45
            Top             =   1425
            Width           =   1620
         End
         Begin VB.Line Line2 
            X1              =   5385
            X2              =   5385
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
            Left            =   6645
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
            Left            =   5445
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
            Left            =   6645
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
            Left            =   5445
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
            Left            =   8115
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
            Left            =   8130
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
            Left            =   9015
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
            Left            =   9015
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
            Left            =   8970
            TabIndex        =   36
            Top             =   1365
            Width           =   1305
         End
         Begin VB.Label Label9 
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
            Left            =   7950
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
            Width           =   10380
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
         Left            =   -74910
         TabIndex        =   9
         Top             =   4380
         Width           =   10500
         Begin VB.Label Label19 
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
            Left            =   7950
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
            Left            =   8970
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
            Left            =   9015
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
            Left            =   9015
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
            Left            =   8130
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
            Left            =   8115
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
            Left            =   5445
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
            Left            =   6645
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
            Left            =   5445
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
            Left            =   6645
            TabIndex        =   24
            Top             =   315
            Width           =   1395
         End
         Begin VB.Line Line1 
            X1              =   5385
            X2              =   5385
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
            Left            =   3555
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
            Left            =   2790
            TabIndex        =   22
            Top             =   1470
            Width           =   645
         End
         Begin VB.Label lbltotalSobFalMN 
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
            Left            =   1290
            TabIndex        =   21
            Top             =   990
            Width           =   1305
         End
         Begin VB.Label Label12 
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
            Left            =   120
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
            Left            =   3900
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
            Left            =   2760
            TabIndex        =   18
            Top             =   1035
            Width           =   1080
         End
         Begin VB.Label lbltotalDevMN 
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
            Left            =   3900
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
            Left            =   2760
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
            Left            =   3900
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
            Left            =   2760
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
            Left            =   1290
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
            Left            =   120
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
            Left            =   1290
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
            Left            =   120
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
            Width           =   10380
         End
      End
      Begin Sicmact.FlexEdit fgPlanillaMN 
         Height          =   3945
         Left            =   -74910
         TabIndex        =   8
         Top             =   450
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   6959
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Descripción-N° Mov-Efectivo-Cheque-Orden Pago-Total"
         EncabezadosAnchos=   "350-4000-800-1200-1200-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColor       =   16777215
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-R-R-R"
         FormatosEdit    =   "0-0-3-2-2-2-2"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit fgPlanillaME 
         Height          =   3945
         Left            =   90
         TabIndex        =   59
         Top             =   450
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   6959
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Descripción-N° Mov-Efectivo-Cheque-Orden Pago-Total"
         EncabezadosAnchos=   "350-4000-800-1200-1200-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-R-R-R"
         FormatosEdit    =   "0-0-3-2-2-2-2"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit fgOtrasAgMN 
         Height          =   5580
         Left            =   -74895
         TabIndex        =   66
         Top             =   585
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   9843
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Descripción-N° Mov-Efectivo-Cheque-Orden Pago-Total"
         EncabezadosAnchos=   "350-4000-800-1200-1200-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-R-R-R-R-R"
         FormatosEdit    =   "0-0-3-2-2-2-2"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit fgOtrasAgME 
         Height          =   5580
         Left            =   -74895
         TabIndex        =   67
         Top             =   585
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   9843
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Descripción-N° Mov-Efectivo-Cheque-Orden Pago-Total"
         EncabezadosAnchos=   "350-4000-800-1200-1200-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-R-R-R-R-R"
         FormatosEdit    =   "0-0-3-2-2-2-2"
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
      Top             =   7575
      Width           =   10905
      _ExtentX        =   19235
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
      Width           =   10680
      Begin MSComCtl2.DTPicker txtfecha 
         Height          =   330
         Left            =   720
         TabIndex        =   6
         Top             =   165
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         Format          =   59834369
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
         Left            =   9345
         TabIndex        =   5
         Top             =   180
         Width           =   1230
      End
      Begin Sicmact.TxtBuscar TxtBuscarUser 
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
      Begin Sicmact.TxtBuscar TxtBuscarAge 
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
Dim oAreaAge As DActualizaDatosArea
Dim oGen As DGeneral
Dim oCajero As nCajero

Private Sub cmdOtrasAge_Click()
tabIngEgre.TabVisible(2) = True
tabIngEgre.TabVisible(3) = True
End Sub

Private Sub cmdProcesar_Click()
CargaPlanilla
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Form_Load()
Set oAreaAge = New DActualizaDatosArea
Set oGen = New DGeneral
Set oCajero = New nCajero
Usuario.Inicio gsCodUser
tabIngEgre.Tab = 0
tabIngEgre.TabVisible(2) = False
tabIngEgre.TabVisible(3) = False
CentraForm Me
Me.Caption = gsOpeDesc
txtBuscarAge.rs = oAreaAge.GetAgencias
txtBuscarUser.psRaiz = "USUARIOS"
txtBuscarUser.rs = oGen.GetUserAreaAgencia(Usuario.cAreaCodAct, Usuario.CodAgeAct)
Select Case gsOpeCod
    Case gOpeHabCajRegSobFaltMN, gOpeHabCajRegSobFaltME
        txtBuscarUser.Enabled = False
        txtBuscarAge.Enabled = False
        'TxtBuscarUser.Text = gsCodUser
        txtBuscarAge.Text = gsCodAge
        lblDescAge = gsNomAge
End Select

End Sub
Private Sub Form_Unload(Cancel As Integer)
Set oAreaAge = Nothing
Set oGen = Nothing
Set oCajero = Nothing
End Sub

Private Sub TxtBuscarAge_EmiteDatos()
lblDescAge = txtBuscarAge.psDescripcion
If txtBuscarAge <> "" And lblDescAge <> "" Then
    txtBuscarUser = ""
    txtBuscarUser.psRaiz = "USUARIOS " & txtBuscarAge.psDescripcion
    txtBuscarUser.Enabled = True
    txtBuscarUser.rs = oGen.GetUserAreaAgencia(Usuario.cAreaCodAct, txtBuscarAge)
End If
End Sub
Sub CargaPlanilla()
Dim rsC As ADODB.Recordset
Dim rsM As ADODB.Recordset
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set rsC = New ADODB.Recordset
Set rsM = oGen.GetConstante(gMoneda)
fgPlanillaMN.Clear
fgPlanillaMN.FormaCabecera
fgPlanillaMN.Rows = 2
fgPlanillaME.Clear
fgPlanillaME.FormaCabecera
fgPlanillaME.Rows = 2
Do While Not rsM.EOF
    Set rsC = oGen.GetConstante(gGruposIngEgre)
    Do While Not rsC.EOF
        Select Case rsM!nConsValor
            Case gMonedaNacional
                fgPlanillaMN.AdicionaFila
                fgPlanillaMN.TextMatrix(fgPlanillaMN.Row, 1) = rsC!cDescripcion
                fgPlanillaMN.BackColorRow &HFDCF09, True
            Case gMonedaExtranjera
                fgPlanillaME.AdicionaFila
                fgPlanillaME.TextMatrix(fgPlanillaME.Row, 1) = rsC!cDescripcion
                fgPlanillaME.BackColorRow &HFDCF09, True
        End Select
        Select Case rsC!nConsValor
            Case gGruposIngEgreAgeLocal
                Set rs = oCajero.GetOpeIngEgreCaptaciones(gGruposIngEgreAgeLocal, CDate(txtFecha), txtBuscarAge, txtBuscarUser, rsM!nConsValor)
                CargaGrid rsM!nConsValor, rs
            Case gGruposIngEgreOtraAgencia
                Set rs = oCajero.GetOpeIngEgreCaptaciones(gGruposIngEgreOtraAgencia, CDate(txtFecha), txtBuscarAge, txtBuscarUser, rsM!nConsValor)
                CargaGrid rsM!nConsValor, rs
            Case gGruposIngEgreOtrasOpe
                Set rs = oCajero.GetOpeIngEgreOtrasOpe(rsC!nConsValor, CDate(txtFecha), txtBuscarAge, txtBuscarUser, rsM!nConsValor)
                CargaGrid rsM!nConsValor, rs
            Case gGruposIngEgreCompraVenta
                Set rs = oCajero.GetOpeIngEgreCompraVenta(rsC!nConsValor, CDate(txtFecha), txtBuscarAge, txtBuscarUser, rsM!nConsValor)
                CargaGrid rsM!nConsValor, rs
        End Select
        rsC.MoveNext
    Loop
    rsC.Close
    Set rsC = Nothing
    rsM.MoveNext
Loop
rsM.Close
Set rsM = Nothing
End Sub
Sub CargaGrid(ByVal pnMoneda As Moneda, rs As ADODB.Recordset)
Select Case pnMoneda
    Case gMonedaNacional
        Do While Not rs.EOF
            fgPlanillaMN.AdicionaFila
            fgPlanillaMN.TextMatrix(fgPlanillaMN.Row, 1) = rs!cGrupoNombre
            If rs!cGrupoCod <> "" Then
                fgPlanillaMN.TextMatrix(fgPlanillaMN.Row, 2) = rs!nTotalMov
                fgPlanillaMN.TextMatrix(fgPlanillaMN.Row, 3) = rs!Efectivo
                fgPlanillaMN.TextMatrix(fgPlanillaMN.Row, 4) = rs!cheque
                fgPlanillaMN.TextMatrix(fgPlanillaMN.Row, 5) = rs!OrdenPago
                fgPlanillaMN.TextMatrix(fgPlanillaMN.Row, 6) = rs!Efectivo + rs!cheque + rs!OrdenPago
            Else
                fgPlanillaMN.BackColorRow &HE0E0E0, True
            End If
            rs.MoveNext
        Loop
        fgPlanillaMN.FormateaColumnas
        rs.Close
        Set rs = Nothing
    Case gMonedaExtranjera
        Do While Not rs.EOF
            fgPlanillaME.AdicionaFila
            fgPlanillaME.TextMatrix(fgPlanillaME.Row, 1) = rs!cGrupoNombre
            If rs!cGrupoCod <> "" Then
                fgPlanillaME.TextMatrix(fgPlanillaME.Row, 2) = rs!nTotalMov
                fgPlanillaME.TextMatrix(fgPlanillaME.Row, 3) = rs!Efectivo
                fgPlanillaME.TextMatrix(fgPlanillaME.Row, 4) = rs!cheque
                fgPlanillaME.TextMatrix(fgPlanillaME.Row, 5) = rs!OrdenPago
                fgPlanillaME.TextMatrix(fgPlanillaME.Row, 6) = rs!Efectivo + rs!cheque + rs!OrdenPago
            Else
                fgPlanillaME.BackColorRow &HE0E0E0, True
            End If
            rs.MoveNext
        Loop
        fgPlanillaME.FormateaColumnas
        rs.Close
        Set rs = Nothing
End Select

End Sub
