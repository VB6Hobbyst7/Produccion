VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmEstadDiariaMora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadistica Diaria sobre la Mora"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   3525
      Left            =   30
      TabIndex        =   12
      Top             =   1860
      Width           =   8655
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
         Height          =   3105
         Left            =   60
         TabIndex        =   13
         Top             =   180
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   5477
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   8685
      Begin VB.CheckBox ChkExcel 
         Caption         =   "Excel"
         Height          =   345
         Left            =   6510
         TabIndex        =   11
         Top             =   1380
         Width           =   1965
      End
      Begin MSComctlLib.ListView LstAgencias 
         Height          =   975
         Left            =   180
         TabIndex        =   9
         Top             =   690
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   1720
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo Agencia"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre de Agencia"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   6990
         TabIndex        =   8
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   5700
         TabIndex        =   7
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox txtTipoCambio 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4710
         TabIndex        =   6
         Top             =   1380
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker DTPInicial 
         Height          =   345
         Left            =   1380
         TabIndex        =   2
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   67502081
         CurrentDate     =   38545
      End
      Begin MSComCtl2.DTPicker DTPFinal 
         Height          =   345
         Left            =   4230
         TabIndex        =   4
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   67502081
         CurrentDate     =   38545
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seleccione las Agencias que desea hacer  la consulta "
         ForeColor       =   &H000000C0&
         Height          =   555
         Left            =   6390
         TabIndex        =   10
         Top             =   630
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4740
         TabIndex        =   5
         Top             =   1080
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3210
         TabIndex        =   3
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   210
         Width           =   1080
      End
   End
End
Attribute VB_Name = "FrmEstadDiariaMora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    DTPInicial.value = Date
    DTPFinal.value = Date
    MSH.Clear
    MSH.Rows = 2
End Sub

Sub CargarAgencias()
    Dim rs As ADODB.Recordset
    Dim objEstaDiaria As EstaDiaria
    
    Set objEstaDiaria = New EstaDiaria
    Set rs = objEstaDiaria.ListaAgencias
    Set objEstaDiaria = Nothing
End Sub
