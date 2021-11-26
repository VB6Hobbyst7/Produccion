VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmCapReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPORTES DE CAPTACIONES"
   ClientHeight    =   8445
   ClientLeft      =   2700
   ClientTop       =   1590
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCapReportes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraBusqEstadoGiro 
      Caption         =   "Estados Giro"
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
      Height          =   855
      Left            =   6840
      TabIndex        =   60
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
      Begin VB.OptionButton optTodosEstadoGiro 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optPendientesGiro 
         Caption         =   "Pendientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   61
         Top             =   525
         Width           =   1695
      End
   End
   Begin VB.CheckBox chkaTasaPactada 
      Caption         =   "Solo a Tasa Pactada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6960
      TabIndex        =   59
      Top             =   7800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame fraFormaRetiro 
      Caption         =   "Formas de Retiro"
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
      Height          =   1150
      Left            =   6840
      TabIndex        =   55
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
      Begin VB.ListBox lstFormaRet 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         ItemData        =   "frmCapReportes.frx":030A
         Left            =   110
         List            =   "frmCapReportes.frx":0317
         Style           =   1  'Checkbox
         TabIndex        =   56
         Top             =   250
         Width           =   1680
      End
   End
   Begin VB.Frame fraProducto 
      Caption         =   "Producto"
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
      Height          =   855
      Left            =   6840
      TabIndex        =   52
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CheckBox chkPlazoFijo 
         Caption         =   "Plazo Fijo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   500
         Width           =   1455
      End
      Begin VB.CheckBox chkAhorro 
         Caption         =   "Ahorro Cte."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.ListBox lstSubProd 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      ItemData        =   "frmCapReportes.frx":034B
      Left            =   6960
      List            =   "frmCapReportes.frx":0364
      Style           =   1  'Checkbox
      TabIndex        =   49
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame fraSubProducto 
      Caption         =   "SubProducto"
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
      Height          =   1695
      Left            =   6840
      TabIndex        =   48
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame FraCond 
      Caption         =   "Condición"
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
      Height          =   855
      Left            =   6840
      TabIndex        =   47
      Top             =   6000
      Visible         =   0   'False
      Width           =   1935
      Begin VB.OptionButton optSoloCan 
         Caption         =   "Solo Cancelados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   51
         Top             =   520
         Width           =   1695
      End
      Begin VB.OptionButton optSoloVig 
         Caption         =   "Solo Vigentes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fracmacs 
      Caption         =   "Incluir"
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
      Height          =   885
      Left            =   6825
      TabIndex        =   34
      Top             =   6360
      Visible         =   0   'False
      Width           =   1500
      Begin VB.CheckBox chkLlamadas 
         Caption         =   "Llamadas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   36
         Top             =   285
         Width           =   1170
      End
      Begin VB.CheckBox chkRecepcion 
         Caption         =   "Recepcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   35
         Top             =   540
         Width           =   1080
      End
   End
   Begin VB.Frame FraProductos 
      Caption         =   "Incluir"
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
      Height          =   885
      Left            =   6825
      TabIndex        =   44
      Top             =   3405
      Visible         =   0   'False
      Width           =   1890
      Begin VB.CheckBox chkCred 
         Caption         =   "Créditos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   46
         Top             =   540
         Width           =   1080
      End
      Begin VB.CheckBox chkCap 
         Caption         =   "Captaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   45
         Top             =   285
         Width           =   1560
      End
   End
   Begin VB.Frame Frapersoneria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Personería"
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
      Height          =   2295
      Left            =   6810
      TabIndex        =   37
      Top             =   1995
      Visible         =   0   'False
      Width           =   2115
      Begin VB.ListBox lstpersoneria 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         ItemData        =   "frmCapReportes.frx":0441
         Left            =   105
         List            =   "frmCapReportes.frx":0463
         TabIndex        =   38
         Top             =   240
         Width           =   1905
      End
   End
   Begin VB.CheckBox chkParaCierre 
      Caption         =   "Para cierre"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6810
      TabIndex        =   42
      Top             =   690
      Width           =   1635
   End
   Begin ComctlLib.ProgressBar prgBar 
      Height          =   270
      Left            =   75
      TabIndex        =   41
      Top             =   7815
      Visible         =   0   'False
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   476
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdExportarExcel 
      Caption         =   "&Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6870
      TabIndex        =   39
      Top             =   7260
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Frame fraOrden 
      Caption         =   "Ordenar"
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
      Height          =   885
      Left            =   6825
      TabIndex        =   30
      Top             =   6345
      Width           =   1890
      Begin VB.CheckBox chkTotal 
         Caption         =   "Total"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   33
         Top             =   495
         Width           =   765
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Transacción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   32
         Top             =   270
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   45
         TabIndex        =   31
         Top             =   510
         Width           =   1080
      End
   End
   Begin VB.Frame frafechacheques 
      Caption         =   "Fecha"
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
      Height          =   885
      Left            =   6825
      TabIndex        =   27
      Top             =   6315
      Width           =   1845
      Begin VB.OptionButton Option1 
         Caption         =   "de Valorizacion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   29
         Top             =   510
         Width           =   1515
      End
      Begin VB.OptionButton Option1 
         Caption         =   "de Registro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   270
         Value           =   -1  'True
         Width           =   1515
      End
   End
   Begin VB.Frame fraCheque 
      Caption         =   "Estado Actual Cheq"
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
      Height          =   1965
      Left            =   6825
      TabIndex        =   25
      Top             =   4350
      Visible         =   0   'False
      Width           =   1845
      Begin VB.ListBox lstcheques 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         ItemData        =   "frmCapReportes.frx":04EE
         Left            =   180
         List            =   "frmCapReportes.frx":0504
         Style           =   1  'Checkbox
         TabIndex        =   26
         Top             =   315
         Width           =   1560
      End
   End
   Begin VB.Frame fraUser 
      Height          =   675
      Left            =   6825
      TabIndex        =   22
      Top             =   3645
      Width           =   1770
      Begin VB.CheckBox Check1 
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   30
         TabIndex        =   24
         Top             =   255
         Width           =   690
      End
      Begin SICMACT.TxtBuscar TxtBuscarUser 
         Height          =   345
         Left            =   720
         TabIndex        =   23
         Top             =   195
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
         Enabled         =   0   'False
         Enabled         =   0   'False
         sTitulo         =   ""
         EnabledText     =   0   'False
         ForeColor       =   12582912
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   75
      TabIndex        =   20
      Top             =   675
      Width           =   6705
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   4215
         Top             =   4470
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCapReportes.frx":0569
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCapReportes.frx":08BB
               Key             =   "Bebe"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCapReportes.frx":0C0D
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCapReportes.frx":0F5F
               Key             =   "Hijito"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeRep 
         Height          =   6780
         Left            =   120
         TabIndex        =   21
         Top             =   165
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   11959
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imglstFiguras"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.OLE OleExcel 
         Class           =   "Excel.Sheet.8"
         Height          =   870
         Left            =   4680
         OleObjectBlob   =   "frmCapReportes.frx":12B1
         TabIndex        =   40
         Top             =   5820
         Visible         =   0   'False
         Width           =   1800
      End
   End
   Begin VB.Frame fraAgencias 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Agencias"
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
      Height          =   660
      Left            =   60
      TabIndex        =   15
      Top             =   -15
      Width           =   8715
      Begin SICMACT.TxtBuscar TxtAgencia 
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   165
         TabIndex        =   17
         Top             =   300
         Width           =   930
      End
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2475
         TabIndex        =   18
         Top             =   240
         Width           =   6045
      End
   End
   Begin VB.Frame fraTipoCambio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   6840
      TabIndex        =   13
      Top             =   2955
      Visible         =   0   'False
      Width           =   1725
      Begin VB.TextBox EditMoney3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   210
         MaxLength       =   10
         TabIndex        =   43
         Top             =   255
         Width           =   1410
      End
      Begin SICMACT.EditMoney EditMoney0 
         Height          =   285
         Left            =   255
         TabIndex        =   14
         Top             =   570
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame fraMonto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Montos"
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
      Height          =   975
      Left            =   6825
      TabIndex        =   8
      Top             =   1995
      Visible         =   0   'False
      Width           =   1725
      Begin SICMACT.EditMoney txtMonto 
         Height          =   285
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtMontoF 
         Height          =   285
         Left            =   480
         TabIndex        =   12
         Top             =   585
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Del:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   278
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Al:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   630
         Width           =   180
      End
   End
   Begin VB.Frame fraFecha 
      Appearance      =   0  'Flat
      Caption         =   "Fechas"
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
      Height          =   975
      Left            =   6825
      TabIndex        =   3
      Top             =   975
      Width           =   1725
      Begin MSMask.MaskEdBox txtFechaF 
         Height          =   300
         Left            =   480
         TabIndex        =   4
         Top             =   570
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   480
         TabIndex        =   5
         Top             =   225
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblAl 
         AutoSize        =   -1  'True
         Caption         =   "Al:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   630
         Width           =   180
      End
      Begin VB.Label lblDel 
         AutoSize        =   -1  'True
         Caption         =   "Del:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   6
         Top             =   285
         Width           =   285
      End
   End
   Begin MSComDlg.CommonDialog dlgGrabar 
      Left            =   2085
      Top             =   6405
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkCondensado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Condensado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   6480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7965
      TabIndex        =   1
      Top             =   8040
      Width           =   1035
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6870
      TabIndex        =   0
      Top             =   8040
      Width           =   1035
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   360
      Left            =   6900
      TabIndex        =   19
      Top             =   3810
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   635
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCapReportes.frx":2AC9
   End
   Begin SICMACT.Usuario Usuario 
      Left            =   90
      Top             =   7170
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin SICMACT.TxtBuscar txtInstitucion 
      Height          =   330
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4313
      _ExtentY        =   582
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
      TipoBusqueda    =   3
      TipoBusPers     =   1
   End
   Begin SICMACT.TxtBuscar TxtBuscaInst 
      Height          =   330
      Left            =   6840
      TabIndex        =   58
      Top             =   7320
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4313
      _ExtentY        =   582
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
      TipoBusqueda    =   3
      TipoBusPers     =   1
   End
End
Attribute VB_Name = "frmCapReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim Flag As Boolean
Dim flag1 As Boolean
Dim Char12 As Boolean
Dim lbPrtCom As Boolean
Dim NumA As String
Dim SalA As String
Dim NumC As String
Dim NumP As String
Dim SalP As String
Dim SalC As String
Dim TB As Currency
Dim TBD As Currency
Dim ca As Currency
Dim CAE As Currency
Dim RegCmacS As Integer
Dim vBuffer As String
Dim lsCadena As String
Dim lcCodInst As String * 13
Dim Progreso As clsProgressBar
'Dim WithEvents lsRep As nCaptaReportes
Dim lsRep As COMNCaptaGenerales.NCOMCaptaReportes
Attribute lsRep.VB_VarHelpID = -1

Dim oGen As COMDConstSistema.DCOMGeneral

Dim Index280100 As Integer ' *** MAVM:Auditoria
Dim Index280200 As Integer ' *** MAVM:Auditoria
Dim Index280201 As Integer ' *** MAVM:Auditoria
Dim Index280210 As Integer ' *** MAVM:Auditoria

Dim Index280208 As Integer ' *** MAVM:Auditoria 06012010
Dim fnEstadoGiro As Integer 'EJVG 20110720

Private Sub Check1_Click()
If Check1.value = 1 Then
    TxtBuscarUser.Enabled = True
Else
    TxtBuscarUser.Enabled = False
    TxtBuscarUser.Text = ""
End If
End Sub

Private Sub chkAhorro_Click()
    If chkAhorro.value = 1 Then
        fraSubProducto.Visible = True
        lstSubProd.Visible = True
    Else
        fraSubProducto.Visible = False
        lstSubProd.Visible = False
    End If
End Sub

Private Sub chkPlazoFijo_Click()
    If chkPlazoFijo.value = 1 Then
        fraFormaRetiro.Visible = True
    Else
        fraFormaRetiro.Visible = False
    End If
End Sub

Private Sub chkTodos_Click()
    If chkTodos.value = 1 Then
    
        TxtAgencia.Text = ""
        TxtAgencia_EmiteDatos
        
        fraUser.Enabled = False
        Check1.Enabled = False
        Check1.value = 0
        TxtBuscarUser.Text = ""
        
    Else
        fraUser.Enabled = True
        Check1.Enabled = True
        Check1.value = 0
        TxtBuscarUser.Text = ""
        
    End If
End Sub

Private Sub cmdExportarExcel_Click()
    Dim sResp As String
    Dim oConst As New COMDConstSistema.DCOMTipoCambio
    Dim nTipoCam As Double
    
    nTipoCam = gnTipCambio
    sResp = RepHavDevBoveda(txtFecha.Text, txtFechaF.Text, gsNomCmac, gsNomAge, Trim(TxtAgencia.Text))
    'ImpExportarExcelCapta Me.txtMonto.Text, nTipoCam, Me.TxtAgencia.Text, gsNomCmac, gsNomAge, gdFecSis, sResp, Me.txtFecha.Text, Me.txtFechaF.Text, Trim(Mid(TreeRep.SelectedItem.Key, 2, 7))

    
End Sub

Public Function RepHavDevBoveda(ByVal psFecIni As Date, ByVal psFecFin As Date, ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psCodAge As String) As String
     'comemtado por copilacion
     ' RepHavDevBoveda = ImpRepHavDevBoveda(psFecIni, psFecFin, psNomCMAC, psNomAge, pscodage)
End Function

Private Sub cmdImprimir_Click()
    Dim oPrevio As previo.clsprevio
    Set oPrevio = New previo.clsprevio
    Dim lsCadena As String
    Dim lsRep As String
    Dim lsEstadosCheques As String
    Dim lsSubProductos As String
    Dim lsCondi As String
    Dim lsOptionsCheques As String
    Dim lscmacllamada As String
    Dim lscmacrecepcion As String
    Dim lsOrden As String
    Dim lscheck As String
    Dim lnCondiAho As Integer
    Dim lsFormaRetiro As String
    Dim i As Integer
    Dim oRep As COMNCaptaGenerales.NCOMCaptaReportes
    Set oRep = New COMNCaptaGenerales.NCOMCaptaReportes

    Dim lsPersCodEmpCTS As String
    lsPersCodEmpCTS = ""
    
    If fraFecha.Visible = True And txtFecha.Visible = True Then
        If Not IsDate(txtFecha) Then
            MsgBox "Fecha no valida", vbInformation, "Aviso"
            Me.txtFecha.SetFocus
            Exit Sub
        End If
    End If
    
    If fraFecha.Visible = True And txtFechaF.Visible = True Then
        If Not IsDate(txtFechaF) Then
            MsgBox "Fecha no valida", vbInformation, "Aviso"
            Me.txtFechaF.SetFocus
            Exit Sub
        End If
    End If
    
    
    If fraFecha.Visible = True And txtFecha.Visible = True And txtFechaF.Visible = True Then
        If CDate(txtFechaF) < CDate(txtFecha) Then
            MsgBox "La fecha de finalizacion debe ser mayor a la fecha de inicio.", vbInformation, "Aviso"
            Me.txtFechaF.SetFocus
            Exit Sub
        End If
    End If
           
    If TxtAgencia.Visible = True Then
        If TxtAgencia.Text = "" And Me.chkTodos.value = 0 Then
            MsgBox "Agencia No valida", vbInformation, "Aviso"
            Me.TxtAgencia.SetFocus
            Exit Sub
        End If
    End If
    
    'By Capi 07082008
'        If Mid(TreeRep.SelectedItem.Key, 2, 7) = "280710" Then
'            Call PrintExtractos(Trim(TxtAgencia.Text), CDate(txtFecha), CDate(txtFechaF))
'            Exit Sub
'        End If
       lcCodInst = TxtBuscaInst
       
        If Mid(TreeRep.SelectedItem.Key, 2, 7) = gCapRepExtractoCtasOP Then
            Call PrintExtractos(Trim(TxtAgencia.Text), CDate(txtFecha), CDate(txtFechaF), gCapAhorros, lcCodInst, 1)
            Exit Sub
        End If
        
        If Mid(TreeRep.SelectedItem.Key, 2, 7) = gCapRepExtractoCtasCTS Then
            If Mid(lcCodInst, 1, 3) <> "109" Then
                Call PrintExtractos(Trim(TxtAgencia.Text), CDate(txtFecha), CDate(txtFechaF), gCapCTS, lcCodInst, 2)
                Exit Sub
            Else
                Call PrintExtractos(Trim(TxtAgencia.Text), CDate(txtFecha), CDate(txtFechaF), gCapCTS, lcCodInst, 3)
                Exit Sub
            End If
        End If
        'By Capi 01102008
        If Mid(TreeRep.SelectedItem.Key, 2, 7) = gCapRepDiaPersoneriaPF Then
            Call generarReportePFPorPersoneria
            Exit Sub
        End If
        '
        If Mid(TreeRep.SelectedItem.Key, 2, 7) = gCapRepPagosWesternUnion Then
            Call ReportePagosWesternUnion(CDate(txtFecha), CDate(txtFechaF))
            Exit Sub
        End If
    'End by
    
    'Add by Gitu 04-03-2010
    If Mid(TreeRep.SelectedItem.Key, 2, 7) = "280306" Then
        'frmGenerarCarta.Inicia "280306"
        Exit Sub
    End If
    'End by
    
    If fraUser.Visible = True Then
        If Check1.value = 1 Then
            If Len(Trim(TxtBuscarUser.Text)) > 0 Then
            Else
                MsgBox "Seleccione un usuario", vbInformation, "Aviso"
                Me.TxtBuscarUser.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    If fraTipoCambio.Visible = True Then
        If val(EditMoney3.Text) = 0 Then
            MsgBox "Ingrese Tipo de Cambio válido", vbExclamation, "Aviso"
            If fraFecha.Visible = True Then
                GetTipCambio (txtFecha.Text)
            Else
                GetTipCambio (gdFecSis)
            End If
            EditMoney3.Text = gnTipCambio
            EditMoney3.SetFocus
        End If
    End If
    
    If fracmacs.Visible = True Then
        If chkLlamadas.value = 0 And chkRecepcion.value = 0 Then
            MsgBox "Seleccione una opcion de llamada/recepcion", vbExclamation, "Aviso"
            chkLlamadas.SetFocus
            Exit Sub
        End If
    End If
    
    If fraCheque.Visible = True Then
        For i = 0 To lstcheques.ListCount - 1
            If lstcheques.Selected(i) = True Then
                If Len(Trim(lsEstadosCheques)) > 0 Then
                    lsEstadosCheques = lsEstadosCheques & ", " & Right(lstcheques.List(i), 1) & ""
                Else
                    lsEstadosCheques = "" & Right(lstcheques.List(i), 1) & ""
                End If
            End If
        Next
        
        If Option1(0).value = True Then
            lsOptionsCheques = "1"
        ElseIf Option1(1).value = True Then
            lsOptionsCheques = "2"
        End If
    Else
        lsEstadosCheques = ""
        lsOptionsCheques = ""
    End If
    
    '***** GITU 08/05/2008 *****'
    lnCondiAho = 0
    
    If fraProducto.Visible Then
        If chkAhorro.value = chkPlazoFijo.value Then
            lnCondiAho = 2
        Else
            If chkAhorro.value = 1 Then
                lnCondiAho = 0
            End If
            If chkPlazoFijo.value = 1 Then
                lnCondiAho = 1
            End If
        End If
    End If
    
    If FraCond.Visible Then
        If optSoloVig.value = True Then
            lsCondi = "1"
        End If
        If optSoloCan.value = True Then
            lsCondi = "2"
        End If
    End If
    
    If fraSubProducto.Visible = True And lstSubProd.Visible = True Then
        For i = 0 To lstSubProd.ListCount - 1
            If lstSubProd.Selected(i) = True Then
                If Len(Trim(lsSubProductos)) > 0 Then
                    lsSubProductos = lsSubProductos & ", " & Right(lstSubProd.List(i), 1) & ""
                Else
                    lsSubProductos = "" & Right(lstSubProd.List(i), 1) & ""
                End If
            End If
        Next
        If chkAhorro = 1 And lsSubProductos = "" Then
            MsgBox "Seleccione uno o todos los SubProductos", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
    Else
        lsSubProductos = ""
    End If
    
    If fraFormaRetiro.Visible = True Then
        For i = 0 To lstFormaRet.ListCount - 1
            If lstFormaRet.Selected(i) = True Then
                If Len(Trim(lsFormaRetiro)) > 0 Then
                    lsFormaRetiro = lsFormaRetiro & ", " & Right(lstFormaRet.List(i), 1) & ""
                Else
                    lsFormaRetiro = "" & Right(lstFormaRet.List(i), 1) & ""
                End If
            End If
        Next
        If chkPlazoFijo = 1 And lsFormaRetiro = "" Then
            MsgBox "Seleccione Alguna de las Formas de retiro", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
    Else
        lsFormaRetiro = ""
    End If
    '*** Fin GITU ***'
    
    If Mid(TreeRep.SelectedItem.Key, 2, 7) = "280615" Then
        lsPersCodEmpCTS = frmEmpresa.Ini
        Set frmEmpresa = Nothing
    End If
    
    If fraOrden.Visible = True Then
        If Option2(0).value = True Then
            lsOrden = "1"
            lscheck = "0"
        Else
            lsOrden = "2"
            If chkTotal.value = 1 Then
                lscheck = "1"
            Else
                lscheck = "0"
            End If
        End If
    End If
    
    
    If fracmacs.Visible = True Then
        If chkLlamadas.value = 1 Then
            lscmacllamada = "1"
        Else
            lscmacllamada = "0"
        End If
        If chkRecepcion.value = 1 Then
            lscmacrecepcion = "1"
        Else
            lscmacrecepcion = "0"
        End If
    End If
    
    rtfCartas.Filename = App.Path & cPlantillaCartaRenPF
    
    Dim pspersoneria As Integer
    If lstpersoneria.ListCount > 0 And lstpersoneria.Visible Then
        If lstpersoneria.ListIndex = -1 Then
            MsgBox "Seleccione el tipo de personería para este reporte", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        pspersoneria = lstpersoneria.ItemData(lstpersoneria.ListIndex)
    End If
    lsRep = Mid(TreeRep.SelectedItem.Key, 2, 7)
    
    Dim sProd As String
    sProd = ""
    If Me.FraProductos.Visible Then
        If chkCap.value = vbChecked And chkCred.value = vbUnchecked Then
            sProd = "CAPTACIONES"
        ElseIf chkCap.value = vbUnchecked And chkCred.value = vbChecked Then
            sProd = "CREDITOS"
        ElseIf (chkCap.value = vbUnchecked And chkCred.value = vbUnchecked) Or (chkCap.value = vbChecked And chkCred.value = vbChecked) Then
            sProd = ""
        End If
    End If
    'By Capi 13082008
        
        If Mid(TreeRep.SelectedItem.Key, 2, 7) = "280211" Then
            lsCadena = gPrnTamLetra10CPIDef & oRep.Reporte(lsRep, IIf(lsRep = "280719" Or lsRep = "280720" Or lsRep = "280721", gdFecSis, Me.txtFecha), Me.txtFechaF, Me.txtMonto.value, Me.txtMontoF.value, gsNomAge, gsNomCmac, gdFecSis, Me.TxtAgencia.Text, TxtBuscarUser.Text, rtfCartas.Text, val(EditMoney3.Text), lsEstadosCheques, lsOptionsCheques, lsOrden, lscheck, lscmacllamada, lscmacrecepcion, pspersoneria, sProd, lsPersCodEmpCTS, lsSubProductos, lsFormaRetiro, lnCondiAho, lsCondi, IIf(chkaTasaPactada = 1, True, False))
            oPrevio.Show lsCadena, Caption, True, 66, gImpresora
            Exit Sub
        End If
    '
    If chkParaCierre.value = vbChecked Then
        Dim clsAge As COMDConstantes.DCOMAgencias, rsTmp As ADODB.Recordset

        Set clsAge = New COMDConstantes.DCOMAgencias

        Set rsTmp = clsAge.RecuperaAgencias

        'While Not rsTmp.EOF

                lsCadena = gPrnTamLetra10CPIDef & oRep.Reporte(lsRep, IIf(lsRep = "280719" Or lsRep = "280720" Or lsRep = "280721", gdFecSis, Me.txtFecha), Me.txtFechaF, Me.txtMonto.value, Me.txtMontoF.value, gsNomAge, gsNomCmac, gdFecSis, rsTmp.Fields("cAgeCod"), TxtBuscarUser.Text, rtfCartas.Text, val(EditMoney3.Text), lsEstadosCheques, lsOptionsCheques, lsOrden, lscheck, lscmacllamada, lscmacrecepcion, pspersoneria)
                oPrevio.Show lsCadena, Caption, True, 66, gImpresora
'                 oPrevio.Show lscadena, Caption, True, 66
                rsTmp.MoveNext
        'Wend

        Exit Sub

     Else
        'MADM 20091119
        If Me.LblAgencia.Caption = "" Then
        lsCadena = gPrnTamLetra10CPIDef & oRep.Reporte(lsRep, IIf(lsRep = "280719" Or lsRep = "280720" Or lsRep = "280721", gdFecSis, Me.txtFecha), Me.txtFechaF, Me.txtMonto.value, Me.txtMontoF.value, gsNomAge, gsNomCmac, gdFecSis, Me.TxtAgencia.Text, TxtBuscarUser.Text, rtfCartas.Text, val(EditMoney3.Text), lsEstadosCheques, lsOptionsCheques, lsOrden, lscheck, lscmacllamada, lscmacrecepcion, pspersoneria, sProd, lsPersCodEmpCTS, lsSubProductos, lsFormaRetiro, lnCondiAho, lsCondi, , fnEstadoGiro)
        Else
        lsCadena = gPrnTamLetra10CPIDef & oRep.Reporte(lsRep, IIf(lsRep = "280719" Or lsRep = "280720" Or lsRep = "280721", gdFecSis, Me.txtFecha), Me.txtFechaF, Me.txtMonto.value, Me.txtMontoF.value, LblAgencia.Caption, gsNomCmac, gdFecSis, Me.TxtAgencia.Text, TxtBuscarUser.Text, rtfCartas.Text, val(EditMoney3.Text), lsEstadosCheques, lsOptionsCheques, lsOrden, lscheck, lscmacllamada, lscmacrecepcion, pspersoneria, sProd, lsPersCodEmpCTS, lsSubProductos, lsFormaRetiro, lnCondiAho, lsCondi, , fnEstadoGiro)
        End If
     End If
    
    
    If lsRep = "280231" Then
        oPrevio.Show lsCadena, Caption, True, 66, gImpresora
        Exit Sub
    End If
    
    If chkCondensado.value = 1 Then
        oPrevio.Show lsCadena, Caption, True, 66, gImpresora
    Else
        oPrevio.Show lsCadena, Caption, False, 66, gImpresora
    End If
    Set clsAge = Nothing
End Sub
'By Capi 07082008
'Private Sub PrintExtractos(ByVal SCodAge As String, ByVal fechaini As Date, ByVal fechafin As Date)
Private Sub PrintExtractos(ByVal sCodAge As String, ByVal fechaini As Date, ByVal FechaFin As Date, Optional ByVal pcProducto As String = "", Optional ByVal pcCodInst As String = "", Optional ByVal pnAccion As Integer)
    Dim oRep As COMNCaptaGenerales.NCOMCaptaReportes
    Set oRep = New COMNCaptaGenerales.NCOMCaptaReportes
    Dim oPrevio As New previo.clsprevio
    Dim lsCadImp As String
    
    
    lsCadImp = oRep.ImpExtractosBatch(sCodAge, gdFecSis, fechaini, FechaFin, sLpt, gsCodUser, gsNomCmac, gsNomAge, pcProducto, pcCodInst, pnAccion)
    oPrevio.Show lsCadImp, "Extractos", False, , gImpresora
    Set oPrevio = Nothing
    Set oRep = Nothing
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim lsCodCab As String
    Dim lsCodCab1 As String
    Dim oCons As COMDConstantes.DCOMConstantes
    Set oCons = New COMDConstantes.DCOMConstantes
    Set oGen = New COMDConstSistema.DCOMGeneral
    LlenaArbol
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    
    'Set Progreso = New clsProgressBar
    
    txtFecha = Format(gdFecSis, gsFormatoFechaView)
    txtFechaF = Format(gdFecSis, gsFormatoFechaView)
    
    lstpersoneria.ItemData(0) = gPersonaNat
    lstpersoneria.ItemData(1) = gPersonaJurSFL
    lstpersoneria.ItemData(2) = gPersonaJurCFL
    lstpersoneria.ItemData(3) = gPersonaJurCFLCMAC
    lstpersoneria.ItemData(4) = gPersonaJurCFLCRAC
    lstpersoneria.ItemData(5) = gPersonaJurCFLFONCODES
    lstpersoneria.ItemData(6) = gPersonaJurCFLCooperativa
    lstpersoneria.ItemData(7) = gPersonaJurCFLEdpyme
    '***Modificado por ELRO el 20111105, según Acta 307-2011/TI-D
    'lstpersoneria.ItemData(8) = 0
    lstpersoneria.ItemData(8) = gPersonaJurCFLFinancieras
    lstpersoneria.ItemData(9) = 0
    '***Fin Modificado por ELRO**********************************
    
    
    
    Me.TxtAgencia.rs = oCons.getAgencias(, , True)
    Usuario.Inicio gsCodUser
    
    Set oCons = Nothing
    Set oGen = Nothing
End Sub
Private Sub lsRep_CloseProgress()
    Progreso.CloseForm Me
End Sub

Private Sub lsRep_Progress(pnValor As Long, pnTotal As Long)
    Progreso.Max = pnTotal
    Progreso.Progress pnValor, "Generando Reporte", "Procesando ..."
End Sub

Private Sub lsRep_ShowProgress()
    Progreso.ShowForm Me
End Sub

Private Sub LlenaArbol()
    Dim sqlv As String
    Dim rsUsu As New ADODB.Recordset
    Dim sOperacion As String, sOpeCod As String
    Dim nodOpe As Node
    Dim sOpePadre As String
    Dim sOpeHijo As String
    Dim sOpeHijito As String
    Dim i As Integer ' ***MAVM:Auditoria
    Dim oCapR As COMDCaptaGenerales.COMDCaptaReportes
    'pti1 20-03-2019 ACTA N° 034-2019*********************
    Dim oPer As COMDPersona.UComAcceso
    Dim cGrupos As String 'pti1
    Dim sCadGruposdeUsu As String
    Dim pCadGrupo As String
    Dim sTemporal As String
    pCadGrupo = ""
    'pti1 *********************************

    Set oCapR = New COMDCaptaGenerales.COMDCaptaReportes
   
    'Set rsUsu = oCapR.Cap_ReportesCaptaciones 'comentado por pti1 ACTA N° 034-2019
    'add pti1 20-03-2019 ACTA N° 034-2019 ***********************************
    Set oPer = New COMDPersona.UComAcceso
    cGrupos = oPer.GruposUsuario(gsCodUser, gsDominio)
    sTemporal = cGrupos
    
    If cGrupos = "" Then
                sCadGruposdeUsu = "'" & gsCodUser & "'"
                sTemporal = "'" & gsCodUser & "'"
                pCadGrupo = sTemporal
    Else
                cGrupos = cGrupos & ",'" & gsCodUser & "' "
                sTemporal = sTemporal & ",'" & gsCodUser & "' "
                pCadGrupo = Replace(cGrupos, "'", "")
                
    End If
    Set rsUsu = oCapR.Cap_ReportesCaptaciones(gsCodUser, pCadGrupo)
    'FIN pti1 ***********************************************
    Set oCapR = Nothing
    
    Do While Not rsUsu.EOF
        sOpeCod = rsUsu("Codigo")
        sOperacion = sOpeCod & " " & UCase(rsUsu("Descripcion"))
        Select Case rsUsu("Nivel")
            Case "1"
                sOpePadre = "P" & sOpeCod
                Set nodOpe = TreeRep.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
                nodOpe.Tag = sOpeCod
            Case "2"
                sOpeHijo = "H" & sOpeCod
                Set nodOpe = TreeRep.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
                nodOpe.Tag = sOpeCod
            Case "3"
                sOpeHijito = "J" & sOpeCod
                Set nodOpe = TreeRep.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
                nodOpe.Tag = sOpeCod
            Case "4"
                Set nodOpe = TreeRep.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
                nodOpe.Tag = sOpeCod
        End Select
        i = i + 1 ' ***MAVM:Auditoria
        If sOpeCod = "280100" Then Index280100 = i ' ***MAVM:Auditoria
        If sOpeCod = "280200" Then Index280200 = i ' ***MAVM:Auditoria
        If sOpeCod = "280201" Then Index280201 = i ' ***MAVM:Auditoria
        If sOpeCod = "280210" Then Index280210 = i ' ***MAVM:Auditoria
        
        If sOpeCod = "280208" Then Index280208 = i ' ***MAVM:Auditoria 06012010
        rsUsu.MoveNext
    Loop
    rsUsu.Close
    Set rsUsu = Nothing
    
  Set oCapR = Nothing
End Sub

Private Sub Option2_Click(index As Integer)
    chkTotal.Enabled = IIf(index = 0, False, True)
    chkTotal.value = 0
End Sub

'*** MAVM: Modulo de Auditoria 21/08/2008
' Para Mostrar seleccionado El Reporte de Informacion que generan las operaciones de ahorros
' por defecto en el Modulo de Auditoria
'Public Sub Inicializar_OperacionAhorros()
'    TreeRep.Nodes(Index280100).Selected = True
'    treeRep_Click
'    TreeRep.Nodes(Index280100).Expanded = True
'
'    TreeRep.Nodes(Index280200).Selected = True
'    treeRep_Click
'    TreeRep.Nodes(Index280200).Expanded = True
'
'    TreeRep.Nodes(Index280201).Selected = True
'    treeRep_Click
'    TreeRep.Nodes(Index280201).Expanded = True
'
'    TreeRep.Enabled = False
'    TreeRep.HideSelection = False
'End Sub
'*** MAVM: Modulo de Auditoria 21/08/2008

'*** MAVM: Modulo de Auditoria 21/08/2008
' Para Mostrar seleccionado El Reporte de las cuentas aperturadas
' por defecto en el Modulo de Auditoria
'Public Sub Inicializar_CtaAperturadas()
'    TreeRep.Nodes(Index280100).Selected = True
'    treeRep_Click
'    TreeRep.Nodes(Index280100).Expanded = True
'
'    TreeRep.Nodes(Index280200).Selected = True
'    treeRep_Click
'    TreeRep.Nodes(Index280200).Expanded = True
'
'    TreeRep.Nodes(Index280210).Selected = True
'    treeRep_Click
'    TreeRep.Nodes(Index280210).Expanded = True
'
'    TreeRep.Enabled = False
'    TreeRep.HideSelection = False
'End Sub
'*** MAVM: Modulo de Auditoria 21/08/2008

'*** MAVM: Modulo de Auditoria 06/01/2010
' Para Mostrar Seleccionado El Reporte de las Cuentas Inactivas
' Por Defecto en el Modulo de Auditoria
Public Sub Inicializar_CtaInactivas()
    TreeRep.Nodes(Index280100).Selected = True
    treeRep_Click
    TreeRep.Nodes(Index280100).Expanded = True
    
    TreeRep.Nodes(Index280200).Selected = True
    treeRep_Click
    TreeRep.Nodes(Index280200).Expanded = True
    
    TreeRep.Nodes(Index280208).Selected = True
    treeRep_Click
    TreeRep.Nodes(Index280208).Expanded = True
    
    TreeRep.Enabled = False
    TreeRep.HideSelection = False
End Sub
'*** MAVM: Modulo de Auditoria 21/08/2008

Private Sub optPendientesGiro_Click()
    fnEstadoGiro = 2
End Sub
Private Sub optTodosEstadoGiro_Click()
    fnEstadoGiro = 1
End Sub

Private Sub treeRep_Click()
    Dim i As Integer
    Dim lsOpc As String
        
      Limpia
      cmdImprimir.Visible = True
      Me.Caption = "REPORTES DE CAPTACIONES - " & Mid(TreeRep.SelectedItem.Text, 8, Len(TreeRep.SelectedItem.Text) - 14)
      chkParaCierre.Visible = False
      
''''        lsOpc = Mid(TreeRep.SelectedItem, InStr(1, TreeRep.SelectedItem, "[") + 1, InStr(InStr(1, TreeRep.SelectedItem, "[") + 1, TreeRep.SelectedItem, "-") - InStr(1, TreeRep.SelectedItem, "[") - 1)
''''
''''        If lsOpc = "0" Then
''''            Me.fraFecha.Visible = False
''''        ElseIf lsOpc = "1" Then
''''            Me.fraFecha.Visible = True
''''            Me.txtFecha.Text = Format(gdFecSis, gsFormatoFechaView)
''''            Me.txtFechaF.Text = Format(gdFecSis, gsFormatoFechaView)
''''            Me.txtFecha.Visible = True
''''            Me.txtFechaF.Visible = False
''''            Me.txtFecha.Enabled = False
''''            Me.txtFechaF.Enabled = False
''''            Me.lblAl.Visible = False
''''        ElseIf lsOpc = "2" Then
''''            Me.fraFecha.Visible = True
''''            Me.txtFechaF.Text = Format(gdFecSis, gsFormatoFechaView)
''''            Me.txtFecha.Visible = True
''''            Me.txtFechaF.Visible = False
''''            Me.txtFecha.Enabled = True
''''            Me.txtFechaF.Enabled = False
''''            Me.lblAl.Visible = False
''''        ElseIf lsOpc = "3" Then
''''            Me.fraFecha.Visible = True
''''            Me.txtFecha.Visible = True
''''            Me.txtFechaF.Visible = True
''''            Me.txtFecha.Enabled = True
''''            Me.txtFechaF.Enabled = True
''''            Me.lblAl.Visible = True
''''        End If
''''
''''        lsOpc = Mid(TreeRep.SelectedItem, InStr(1, TreeRep.SelectedItem, "-") + 1, InStr(InStr(1, TreeRep.SelectedItem, "-") + 1, TreeRep.SelectedItem, "-") - InStr(1, TreeRep.SelectedItem, "-") - 1)
''''
''''        If lsOpc = "0" Then
''''            Me.fraMonto.Visible = False
''''        ElseIf lsOpc = "1" Then
''''            Me.fraMonto.Visible = True
''''            Me.txtMonto.Text = "0"
''''            Me.txtMontoF.Text = "0"
''''            Me.txtMonto.Visible = True
''''            Me.txtMontoF.Visible = False
''''            Me.txtMonto.Enabled = False
''''            Me.txtMontoF.Enabled = False
''''            Me.lblAl.Visible = False
''''        ElseIf lsOpc = "2" Then
''''            Me.fraMonto.Visible = True
''''            Me.txtMontoF.Text = "0"
''''            Me.txtMonto.Visible = True
''''            Me.txtMontoF.Visible = False
''''            Me.txtMonto.Enabled = True
''''            Me.txtMontoF.Enabled = False
''''            Me.lblAl.Visible = False
''''        ElseIf lsOpc = "3" Then
''''            Me.fraMonto.Visible = True
''''            Me.txtMonto.Visible = True
''''            Me.txtMontoF.Visible = True
''''            Me.txtMonto.Enabled = True
''''            Me.txtMontoF.Enabled = True
''''            Me.lblAl.Visible = True
''''        End If
        
        fraMonto.Caption = "Montos"
        Frapersoneria.Visible = False
        
        'Nuevos
        Select Case Mid(TreeRep.SelectedItem, 1, 6)
        
        '''''''''''''''''INICIO'''''''''''''''''
        
        Case gCapRepDiaCapEstadAho
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case gCapRepDiaCapEstadPF
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case gCapRepDiaCapEstadCTS
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case gCapRepDiaCapCtasMov
            HabilitaControles True, True, True, True, False, False, False, False, False
        Case gCapRepDiaCapSaldTpoCta
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepDiaCapEstratCta
            HabilitaControles True, False, True, True, False, False, False, False, False
        Case gCapRepDiaCapPFVenc
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case gCapRepDiaCapInact
            HabilitaControles False, False, False, False, False, False, False, False, False
        Case 280232
            cmdExportarExcel.Visible = True
            cmdImprimir.Visible = False
            HabilitaControles True, False, True, True, True, False, False, False, False
            
        Case gCapRepDiaCapConsInact
            HabilitaControles False, False, True, True, True, False, False, False, False
        Case gCapRepDiaCapApert
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case gCapRepDiaCapCanc
            HabilitaControles True, False, True, True, True, False, False, False, False, , , , , , , True
            
        Case gCapRepDiaCartaRenPF
            HabilitaControles True, False, True, True, False, False, False, False, False
        Case gCapRepDiaServGirosApert
            HabilitaControles True, False, True, True, True, False, False, False, False, , , , , , , , True
        Case gCapRepDiaServGirosCanc
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case "280303"
            fraMonto.Caption = "Con N y más días"
            HabilitaControles True, False, False, False, False, True, False, True, False
        Case "280304"
            HabilitaControles True, False, True, True, True, False, False, False, False
            
        
'        Case gCapRepDiaServConvCob
'            HabilitaControles False, False, True, True, True, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoPN
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoPJSFL
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoPJCFL
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoCMAC
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoCRAC
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoFoncodes
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoCooperativa
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoEdipyme
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case "280632"
                Frapersoneria.Visible = True
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case "280633"
                Frapersoneria.Visible = True
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case "280634"
                Frapersoneria.Visible = True
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case "280635"
                Frapersoneria.Visible = True
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case "280636"
                Frapersoneria.Visible = True
                HabilitaControles False, False, False, False, False, False, False, False, False
        Case "280637"
                Frapersoneria.Visible = True
                HabilitaControles False, False, False, False, False, False, False, False, False
        Case "280639"
                HabilitaControles True, False, False, False, False, False, False, False, False
        
        'By Capi 07082008
        Case gCapRepExtractoCtasOP
             HabilitaControles True, False, True, True, True, False, False, False, False
        Case gCapRepExtractoCtasCTS
             HabilitaControles True, False, True, True, True, False, False, False, False, , , , , , True
        Case gCapRepPagosWesternUnion
             HabilitaControles False, False, True, True, True, False, False, False, False
        '
                
            
        Case gCapRepMensCapSaldCtaPFPN
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFPJSFL
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFPJCFL
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFCMAC
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFCRAC
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFFoncodes
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFCooperativa
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFEdpyme
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFPNPJSFL
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFNascaPalpa
                 Frapersoneria.Visible = True
                HabilitaControles False, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFCaneteMala
                 Frapersoneria.Visible = True
                HabilitaControles False, False, False, False, False, False, False, False, False
        Case gCapRepMensBloqCapPF
                Frapersoneria.Visible = True
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case "280631"
                Frapersoneria.Visible = True
                HabilitaControles True, False, False, False, False, False, False, False, False
        '***Agregado por ELRO el 20111105, según Acta 307-2011/TI-D
        Case gCapRepMensCapSaldCtaAhoFinancieras
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFFinancieras
            HabilitaControles True, False, False, False, False, False, False, False, False
        '***Fin Agregado por ELRO**********************************
                
        Case gCapRepMensCapSaldCtaCTS
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaCTSConvenio
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaCTSExternos
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaCTSNascaPalpa
            HabilitaControles False, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaCTSCaneteMala
            HabilitaControles False, False, False, False, False, False, False, False, False
        Case gCapRepMensBloqCapCTS
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensBloqCapCTSConvenio
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensBloqCapCTSExternos
                HabilitaControles True, False, False, False, False, False, False, False, False
                
        
        Case gCapRepMensCapListGralCtas
            HabilitaControles True, False, True, True, False, False, False, False, False
        Case "V03"
            HabilitaControles False, False, True, True, True, False, False, False, False
        Case "280701"
            HabilitaControles True, False, True, True, True, True, True, True, True
        Case "280702"
            HabilitaControles True, False, True, True, True, True, True, True, True
        Case "280703"
            HabilitaControles True, False, True, True, True, True, True, True, True
        Case "280704"
            HabilitaControles True, False, True, True, True, True, True, True, True
        Case "280705"
            HabilitaControles True, True, True, True, True, True, True, True, True
        Case "280706"
            HabilitaControles True, False, True, True, True, True, True, True, True
        Case "280707"
            HabilitaControles True, False, True, True, True, True, True, True, True
        Case "280213"
            HabilitaControles True, True, True, True, False, False, False, False, False, , True 'ojo
        Case "280214"
            HabilitaControles True, True, True, True, False, False, False, False, False, , True 'ojo
        Case "280215"
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case "280216"
            chkParaCierre.Visible = True
            HabilitaControles True, False, True, True, False, False, False, False, False, False
            
        Case "280217"
            HabilitaControles True, False, True, True, False, False, False, False, False
        Case "280708"
            fraMonto.Caption = "Nro. Clientes"
            cmdExportarExcel.Visible = True
            HabilitaControles True, False, False, False, False, True, True, True, False
        Case "280709"
            fraMonto.Caption = "Nro. Clientes"
            cmdExportarExcel.Visible = True
            HabilitaControles True, False, False, False, False, True, True, True, False
        Case "280710"
            HabilitaControles True, False, True, True, True, False, False, False, False
                       
            
        Case "280224"
            HabilitaControles False, False, True, True, True, False, False, False, False
        Case "280225"
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case "280227"
            HabilitaControles True, False, True, True, False, False, False, False, False
        Case "280228"
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case "280219"
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case "280220"
            HabilitaControles True, False, True, True, True, False, False, False, False
        '*** PEAC 20100524
        Case "280221"
            HabilitaControles True, False, True, True, True, False, False, False, False
        '*** FIN PEAC
        Case "280230" 'cheques recibidos
            Me.chkCap.value = vbUnchecked
            Me.chkCred.value = vbUnchecked
            Me.FraProductos.Visible = True
            HabilitaControles True, False, True, True, True, False, False, False, False, True
        Case "280231"
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case "280232"
            HabilitaControles True, False, True, True, True, False, False, False, False
'        Case "280233"
'            HabilitaControles True, , False, True, True, True, False, False, False, False
        Case "280233"
            HabilitaControles True, False, True, True, True, False, False, False, False
            
        Case "280237"
            HabilitaControles True, False, True, True, False, False, False, False, False
            
        Case "280238" 'Agregado por GITU 07/05/2008
            HabilitaControles True, False, True, True, True, False, False, False, False, , , , True, True
            
        Case "280223"
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case "280222"
            HabilitaControles True, False, True, True, False, False, False, False, False, , , True
        
        Case "280719"
            HabilitaControles True, False, False, False, False, False, False, False, False, , , False
        Case "280720"
            HabilitaControles True, False, False, False, False, False, False, False, False, , , False
        Case "280721"
            HabilitaControles True, False, False, False, False, False, False, False, False, , , False
        Case "280722"
            HabilitaControles True, False, True, True, False, False, False, False, False, , , False
        Case gCapRepVariosCapPromotores
            HabilitaControles True, False, True, True, True, False, False, False, False, False, False, False
            
        
        
        '''''''''''''FIN'''''''''''''''''''''
        '*******REPORTES PARA CONTABILIDAD

        Case 280801
                Frapersoneria.Visible = True
                HabilitaControles True, False, True, True, False, False, False, False, False
                
        Case 280802
                Frapersoneria.Visible = True
                HabilitaControles True, False, True, True, False, False, False, False, False
        
        Case 280803
                Frapersoneria.Visible = True
                HabilitaControles True, False, True, True, False, False, False, False, False
        
        Case 280804
                Frapersoneria.Visible = True
                HabilitaControles True, False, True, True, False, False, False, False, False
                
        Case 280805
                Frapersoneria.Visible = True
                HabilitaControles True, False, True, True, False, False, False, False, False
                
        Case 280811
                Frapersoneria.Visible = True
                HabilitaControles True, False, True, True, False, False, False, False, False
        
        Case 280812
                Frapersoneria.Visible = True
                HabilitaControles True, False, True, True, False, False, False, False, False
        
        Case 280813
                Frapersoneria.Visible = True
                HabilitaControles True, False, True, True, False, False, False, False, False
        
        Case 280814
                Frapersoneria.Visible = True
                HabilitaControles True, False, True, True, False, False, False, False, False
        
        Case 280821
                Frapersoneria.Visible = False
                HabilitaControles True, False, True, True, False, False, False, False, False
        
        Case 280822
                Frapersoneria.Visible = False
                HabilitaControles True, False, True, True, False, False, False, False, False
        
        Case 280823
                Frapersoneria.Visible = False
                HabilitaControles True, False, True, True, False, False, False, False, False
                
        Case 280824
                Frapersoneria.Visible = False
                HabilitaControles True, False, True, True, False, False, False, False, False
        
        Case 280831
                
                HabilitaControles False, False, True, True, False, False, False, False, False
        
        Case 280832
                
                HabilitaControles False, False, True, True, False, False, False, False, False
        
        Case 280833
                
                HabilitaControles False, False, True, True, False, False, False, False, False
        
        Case 280834
                
                HabilitaControles False, False, True, True, False, False, False, False, False
        
        Case 280835
                
                HabilitaControles False, False, True, True, False, False, False, False, False
        
        Case 280836
                HabilitaControles False, False, True, True, False, False, False, False, False
        
        Case 280837
                HabilitaControles False, False, True, True, False, False, False, False, False
        
        Case 280838
                HabilitaControles False, False, True, True, False, False, False, False, False
        
        Case 280839
                HabilitaControles False, False, True, True, False, False, False, False, False

        Case "280851"
            cmdExportarExcel.Visible = True
            fraMonto.Caption = "Nro. Clientes"
            HabilitaControles True, False, True, True, False, True, True, True, False
        Case "280852"
            cmdExportarExcel.Visible = True
            fraMonto.Caption = "Nro. Clientes"
            HabilitaControles True, False, True, True, False, True, True, True, False
        Case "280853"
            cmdExportarExcel.Visible = True
            fraMonto.Caption = "Nro. Clientes"
            HabilitaControles True, False, True, True, False, True, True, True, False
        Case "280854"
            cmdExportarExcel.Visible = True
            fraMonto.Caption = "Nro. Clientes"
            HabilitaControles True, False, True, True, False, True, True, True, False
        Case "280855"
            cmdExportarExcel.Visible = True
            fraMonto.Caption = "Nro. Clientes"
            HabilitaControles True, False, True, True, False, True, True, True, False
        Case "280856"
            cmdExportarExcel.Visible = True
            fraMonto.Caption = "Nro. Clientes"
            HabilitaControles True, False, True, True, False, True, True, True, False
        Case "280857"
            cmdExportarExcel.Visible = True
            fraMonto.Caption = "Nro. Clientes"
            HabilitaControles True, False, True, True, False, True, True, True, False
        
        '********** FIN REPORTES PARA CONTABILIDAD

        '*** PEAC 20080616
        Case gCapRepSeguiCtasAhoPandero
            HabilitaControles True, False, True, True, True, False, False, False, False
        
        'By Capi 01102008
        Case gCapRepDiaPersoneriaPF
        HabilitaControles False, False, False, False, False, False, False, False, False
        '

        'Case 280213, 280214
        '    HabilitaControles True
        Case Else
            HabilitaControles False, False, False, False, False, False, False, False, False
        End Select
        End Sub
'By Capi 07082008 se adiciono parametro opcional pbInstitucion para CTS,pbaTasaPactada para plazo fijo,
Private Sub HabilitaControles(ByVal pfraAgencia As Boolean, ByVal pUser As Boolean, ByVal pfraFecha As Boolean, pFecha As Boolean, pfechaf As Boolean, _
                              ByVal pfraMonto As Boolean, ByVal ptipocambio As Boolean, ByVal pMonto As Boolean, _
                              ByVal pMontof As Boolean, Optional pFraCheque As Boolean = False, Optional pFraOrden As Boolean = False, _
                              Optional pFraCmacs As Boolean = False, Optional pFraCond As Boolean = False, Optional pFraProducto As Boolean = False, _
                              Optional pbInstitucion As Boolean = False, Optional pbaTasaPactada As Boolean = False, _
                              Optional pfraBusqEstadoGiro As Boolean = False)
    'EJVG 20110720 se adiciono pfraBusqEstadoGiro
    
    fraCheque.Visible = pFraCheque
    frafechacheques.Visible = pFraCheque
    
    fraOrden.Visible = pFraOrden
    
    fraAgencias.Visible = pfraAgencia
    
    fraUser.Visible = pUser
    fraFecha.Visible = pFecha
    
    fraMonto.Visible = pfraMonto
    txtMonto.Visible = pMonto
    txtMontoF.Visible = pMontof
    Label3.Visible = pMontof
    
    fraFecha.Visible = pfraFecha
    txtFecha.Visible = pFecha
    txtFechaF.Visible = pfechaf
    lblAl.Visible = pfechaf
    
    fraTipoCambio.Visible = ptipocambio
    
    fracmacs.Visible = pFraCmacs
    
    '*** GITU ***' 07/05/2008
    FraCond.Visible = pFraCond
    fraProducto.Visible = pFraProducto
    fraSubProducto.Visible = False
    lstSubProd.Visible = False
    fraFormaRetiro.Visible = False
    chkAhorro.value = 0
    chkPlazoFijo.value = 0
    optSoloVig.value = False
    optSoloCan.value = False
    '*** FIN ***'
    'By Capi 07082008
    If pbInstitucion = True Then
        TxtBuscaInst.Visible = True
    End If
    If pbaTasaPactada = True Then
        chkaTasaPactada.Visible = True
    End If
    FraBusqEstadoGiro.Visible = pfraBusqEstadoGiro 'EJVG 20110720
    If pfraBusqEstadoGiro = True Then
        optTodosEstadoGiro.value = True
        Call optTodosEstadoGiro_Click
    End If
End Sub

Private Sub Limpia()
    txtFecha.Text = "__/__/____"
    txtFechaF.Text = "__/__/____"
    txtMonto.Text = ""
    txtMontoF.Text = ""
    chkTodos.value = 0
    TxtAgencia.Text = ""
    LblAgencia.Caption = ""
    TxtBuscarUser.Text = ""
    chkLlamadas.value = 0
    chkRecepcion.value = 0
    cmdExportarExcel.Visible = False
    TxtBuscaInst.Visible = False
    TxtBuscaInst.Text = ""
End Sub

Private Sub treeRep_Collapse(ByVal Node As MSComctlLib.Node)
'    If Right(Node.Key, 2) = "00" Then
'        Node.ExpandedImage = 2
'    Else
'        Node.ExpandedImage = 1
'    End If
End Sub
Private Sub treeRep_Expand(ByVal Node As MSComctlLib.Node)
'    If Right(Node.Key, 2) = "00" Then
'        Node.ExpandedImage = 2
'    Else
'        Node.ExpandedImage = 1
'    End If
End Sub
Private Sub TreeRep_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    treeRep_Click
End If
End Sub
Private Sub treeRep_NodeCheck(ByVal Node As MSComctlLib.Node)
'    Dim i As Integer
'    TreeRep.SelectedItem = Node
'    Select Case Len(Node.Key)
'       Case Node.Key = "P"
'         If Node.Checked = True Then
'              For i = 1 To TreeRep.Nodes.Count
'                  TreeRep.Nodes(i).Checked = True
'              Next
'         Else
'              For i = 1 To TreeRep.Nodes.Count
'                  TreeRep.Nodes(i).Checked = False
'                 Next
'         End If
'       Case 7 And Right(Node.Key, 2) = "00"
'           If Node.Checked = True Then
'              For i = 1 To TreeRep.Nodes.Count
'                If Mid(TreeRep.Nodes(i).Key, 2, 4) = Mid(Node.Key, 2, 4) Then
'                      TreeRep.Nodes(i).Checked = True
'                      TreeRep.Nodes(i).Image = "Hoja"
'                      TreeRep.Nodes(i).ForeColor = vbBlue
'                End If
'              Next
'              TreeRep.SelectedItem.Image = "Close"
'           Else
'              For i = 1 To TreeRep.Nodes.Count
'                If Mid(TreeRep.Nodes(i).Key, 2, 4) = Mid(Node.Key, 2, 4) Then
'                      TreeRep.Nodes(i).Checked = False
'                      TreeRep.Nodes(i).Image = "Hoja1"
'                     TreeRep.Nodes(i).ForeColor = vbBlack
'                 End If
'              Next
'              TreeRep.SelectedItem.Image = "Open"
'           End If
'       Case Else
'
'           If Node.Checked = True Then
'               Node.Image = "Hoja"
'               Node.ForeColor = vbBlue
'           Else
'               Node.Image = "Hoja1"
'               Node.ForeColor = vbBlack
'           End If
'    End Select
End Sub

Private Sub TreeRep_NodeClick(ByVal Node As MSComctlLib.Node)
    treeRep_Click
End Sub

'Private Sub TxtAgencia_EmiteDatos()
'    Me.lblAgencia.Caption = TxtAgencia.psDescripcion
'End Sub

Private Sub TxtAgencia_EmiteDatos()
Set oGen = New COMDConstSistema.DCOMGeneral
Me.LblAgencia.Caption = TxtAgencia.psDescripcion
If chkTodos.value = 0 Then
    If TxtAgencia <> "" And LblAgencia <> "" Then
        TxtBuscarUser = ""
        TxtBuscarUser.psRaiz = "USUARIOS " & TxtAgencia.psDescripcion
        TxtBuscarUser.Enabled = True
        TxtBuscarUser.rs = oGen.GetUserAreaAgencia(Usuario.cAreaCodAct, TxtAgencia)
    End If
Else
    TxtBuscarUser.Text = ""
    TxtBuscarUser.Enabled = False
End If
End Sub



Private Sub TxtBuscaInst_EmiteDatos()
    lcCodInst = TxtBuscaInst.ColCodigo
End Sub

Private Sub txtFecha_GotFocus()
    fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtFechaF.Enabled And txtFechaF.Visible Then txtFechaF.SetFocus
    End If
End Sub

Private Sub txtFechaF_GotFocus()
fEnfoque txtFechaF
End Sub

Private Sub txtFechaF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TreeRep.SetFocus
End If
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtMontoF.Enabled And txtMontoF.Visible Then txtMontoF.SetFocus
    End If
End Sub
'By Capi 18082008
Private Sub ReportePagosWesternUnion(ByVal pdFecInicial As Date, ByVal pdFecFinal As Date)

Dim oNAho As COMNCaptaGenerales.NCOMCaptaReportes
Dim R As ADODB.Recordset
Dim lsNombreArchivo As String
Dim lMatCabecera(1, 1) As Variant
Dim lsmensaje As String

'Dim i As Integer
'Dim sCadAge As String


'    Screen.MousePointer = 11
'    sCadAge = ""
'    For i = 0 To UBound(pMatAgencias) - 1
'    sCadAge = sCadAge & pMatAgencias(i) & ","
'    Next i
'    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
'
    Set oNAho = New COMNCaptaGenerales.NCOMCaptaReportes
    Set R = oNAho.ReportePagosWesternUnion(pdFecInicial, pdFecFinal)
    Set oNAho = Nothing
    
    lsNombreArchivo = "Pagos Western Union"
            
    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Pagos Efectuados - Western Union ", " Desde " & CStr(pdFecInicial) & " Hasta " & CStr(pdFecFinal), lsNombreArchivo, lMatCabecera, R, 2, True, , True, True)

    Screen.MousePointer = 0

End Sub
'By Capi 01102008
Private Sub generarReportePFPorPersoneria()
    Dim loNCapRep As COMNCaptaGenerales.NCOMCaptaReportes
    Dim R As ADODB.Recordset
    Dim lsNombreArchivo As String
    Dim lMatCabecera(1, 1) As Variant
    Dim lsmensaje As String

    Set loNCapRep = New COMNCaptaGenerales.NCOMCaptaReportes
    Set R = loNCapRep.generarReportePFPersoneria
    Set loNCapRep = Nothing

    lsNombreArchivo = "Plazo Fijo Por Personeria"

    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Detalle Saldos Plazo Fijo por Personeria ", " Al: " & CStr(gdFecSis), lsNombreArchivo, lMatCabecera, R, 2, True, , True, True)

    Screen.MousePointer = 0

End Sub

